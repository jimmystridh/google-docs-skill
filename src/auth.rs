use anyhow::{Context, Result, anyhow};
use chrono::Utc;
use reqwest::blocking::Client;
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use serde_yaml::{Mapping, Value as YamlValue};
use std::fs;
use std::path::{Path, PathBuf};
use url::Url;

pub const DOCS_SCOPE: &str = "https://www.googleapis.com/auth/documents";
pub const DRIVE_SCOPE: &str = "https://www.googleapis.com/auth/drive";
pub const SHEETS_SCOPE: &str = "https://www.googleapis.com/auth/spreadsheets";
pub const CALENDAR_SCOPE: &str = "https://www.googleapis.com/auth/calendar";
pub const CONTACTS_SCOPE: &str = "https://www.googleapis.com/auth/contacts";
pub const GMAIL_SCOPE: &str = "https://www.googleapis.com/auth/gmail.modify";

pub const SHARED_SCOPES: &[&str] = &[
    DRIVE_SCOPE,
    SHEETS_SCOPE,
    DOCS_SCOPE,
    CALENDAR_SCOPE,
    CONTACTS_SCOPE,
    GMAIL_SCOPE,
];

const DEFAULT_AUTH_URI: &str = "https://accounts.google.com/o/oauth2/auth";
const DEFAULT_TOKEN_URI: &str = "https://oauth2.googleapis.com/token";
const OOB_REDIRECT_URI: &str = "urn:ietf:wg:oauth:2.0:oob";

#[derive(Debug, Clone)]
pub struct AuthPaths {
    pub credentials_path: PathBuf,
    pub token_path: PathBuf,
}

impl AuthPaths {
    pub fn from_home(home: &Path) -> Self {
        Self {
            credentials_path: home.join(".claude/.google/client_secret.json"),
            token_path: home.join(".claude/.google/token.json"),
        }
    }
}

#[derive(Debug, Clone)]
pub struct OAuthClientConfig {
    pub client_id: String,
    pub client_secret: String,
    pub auth_uri: String,
    pub token_uri: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum ScopeField {
    Single(String),
    Multiple(Vec<String>),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StoredToken {
    pub client_id: String,
    pub access_token: String,
    #[serde(default)]
    pub refresh_token: Option<String>,
    #[serde(default)]
    pub scope: Option<ScopeField>,
    pub expiration_time_millis: i64,
}

#[derive(Debug, Deserialize)]
struct SecretFile {
    installed: Option<SecretSection>,
    web: Option<SecretSection>,
}

#[derive(Debug, Deserialize)]
struct SecretSection {
    client_id: String,
    client_secret: String,
    auth_uri: Option<String>,
    token_uri: Option<String>,
}

#[derive(Debug, Deserialize)]
struct TokenResponse {
    access_token: String,
    expires_in: Option<i64>,
    refresh_token: Option<String>,
    scope: Option<String>,
}

#[derive(Debug)]
pub enum TokenState {
    Authorized(StoredToken),
    AuthorizationRequired { auth_url: String },
}

pub fn load_oauth_client_config(path: &Path) -> Result<OAuthClientConfig> {
    let raw = fs::read_to_string(path)
        .with_context(|| format!("Failed to read credentials file: {}", path.display()))?;
    let parsed: SecretFile = serde_json::from_str(&raw)
        .with_context(|| format!("Failed to parse credentials JSON: {}", path.display()))?;

    let section = parsed
        .installed
        .or(parsed.web)
        .ok_or_else(|| anyhow!("Expected 'installed' or 'web' section in client secret JSON"))?;

    Ok(OAuthClientConfig {
        client_id: section.client_id,
        client_secret: section.client_secret,
        auth_uri: section
            .auth_uri
            .unwrap_or_else(|| DEFAULT_AUTH_URI.to_string()),
        token_uri: section
            .token_uri
            .unwrap_or_else(|| DEFAULT_TOKEN_URI.to_string()),
    })
}

pub fn build_auth_url(config: &OAuthClientConfig, scopes: &[&str]) -> Result<String> {
    let mut url = Url::parse(&config.auth_uri).context("Invalid auth URI")?;
    {
        let mut qp = url.query_pairs_mut();
        qp.append_pair("client_id", &config.client_id);
        qp.append_pair("redirect_uri", OOB_REDIRECT_URI);
        qp.append_pair("response_type", "code");
        qp.append_pair("scope", &scopes.join(" "));
        qp.append_pair("access_type", "offline");
        qp.append_pair("prompt", "consent");
    }
    Ok(url.to_string())
}

pub fn load_stored_token(path: &Path) -> Result<StoredToken> {
    let raw = fs::read_to_string(path)
        .with_context(|| format!("Failed to read token file: {}", path.display()))?;

    if let Ok(token) = serde_json::from_str::<StoredToken>(&raw) {
        return Ok(token);
    }

    let yaml: YamlValue = serde_yaml::from_str(&raw).with_context(|| {
        format!(
            "Failed to parse token file as JSON or YAML: {}",
            path.display()
        )
    })?;

    if let Some(token) = parse_token_from_yaml(&yaml)? {
        return Ok(token);
    }

    Err(anyhow!(
        "Token file format not recognized: {}",
        path.display()
    ))
}

fn parse_token_from_yaml(yaml: &YamlValue) -> Result<Option<StoredToken>> {
    let Some(mapping) = yaml.as_mapping() else {
        return Ok(None);
    };

    let key = YamlValue::String("default".to_string());
    let Some(value) = mapping.get(&key) else {
        return Ok(None);
    };

    match value {
        YamlValue::String(json_payload) => {
            let parsed = serde_json::from_str::<StoredToken>(json_payload)
                .context("Failed parsing YAML default token payload as JSON")?;
            Ok(Some(parsed))
        }
        other => {
            let parsed = serde_yaml::from_value::<StoredToken>(other.clone())
                .context("Failed parsing YAML default token object")?;
            Ok(Some(parsed))
        }
    }
}

pub fn save_stored_token(path: &Path, token: &StoredToken) -> Result<()> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).with_context(|| {
            format!(
                "Failed to create token parent directory: {}",
                parent.display()
            )
        })?;
    }

    let payload = serde_json::to_string(token).context("Failed serializing token JSON payload")?;

    let mut map = Mapping::new();
    map.insert(
        YamlValue::String("default".to_string()),
        YamlValue::String(payload),
    );

    let serialized = serde_yaml::to_string(&map).context("Failed serializing token YAML")?;
    fs::write(path, serialized)
        .with_context(|| format!("Failed writing token file: {}", path.display()))?;

    Ok(())
}

pub fn complete_authorization(
    config: &OAuthClientConfig,
    code: &str,
    existing_refresh_token: Option<String>,
) -> Result<StoredToken> {
    let client = Client::builder()
        .user_agent("google-docs-rust/1.0")
        .build()
        .context("Failed building HTTP client")?;

    let resp = client
        .post(&config.token_uri)
        .form(&[
            ("code", code),
            ("client_id", config.client_id.as_str()),
            ("client_secret", config.client_secret.as_str()),
            ("redirect_uri", OOB_REDIRECT_URI),
            ("grant_type", "authorization_code"),
        ])
        .send()
        .context("Token exchange request failed")?;

    let status = resp.status();
    let body = resp
        .text()
        .context("Failed reading token exchange response body")?;

    if !status.is_success() {
        let msg = extract_google_error_message(&body)
            .unwrap_or_else(|| format!("Token exchange failed with status {status}"));
        return Err(anyhow!("{msg}"));
    }

    let payload: TokenResponse = serde_json::from_str(&body)
        .with_context(|| format!("Failed parsing token exchange response JSON. Body: {body}"))?;

    let refresh_token = payload
        .refresh_token
        .or(existing_refresh_token)
        .ok_or_else(|| anyhow!("No refresh token received. Re-run auth with consent prompt."))?;

    Ok(StoredToken {
        client_id: config.client_id.clone(),
        access_token: payload.access_token,
        refresh_token: Some(refresh_token),
        scope: payload
            .scope
            .map(|s| ScopeField::Multiple(s.split_whitespace().map(ToString::to_string).collect())),
        expiration_time_millis: compute_expiration(payload.expires_in),
    })
}

pub fn ensure_token(paths: &AuthPaths, scopes: &[&str]) -> Result<TokenState> {
    let config = load_oauth_client_config(&paths.credentials_path)?;

    let mut token = match load_stored_token(&paths.token_path) {
        Ok(t) => t,
        Err(_) => {
            return Ok(TokenState::AuthorizationRequired {
                auth_url: build_auth_url(&config, scopes)?,
            });
        }
    };

    if token_is_expired(&token) {
        if token.refresh_token.is_none() {
            return Ok(TokenState::AuthorizationRequired {
                auth_url: build_auth_url(&config, scopes)?,
            });
        }

        refresh_token(&config, &mut token)?;
        save_stored_token(&paths.token_path, &token)?;
    }

    Ok(TokenState::Authorized(token))
}

pub fn refresh_token(config: &OAuthClientConfig, token: &mut StoredToken) -> Result<()> {
    let refresh_token = token
        .refresh_token
        .clone()
        .ok_or_else(|| anyhow!("Cannot refresh token without refresh_token"))?;

    let client = Client::builder()
        .user_agent("google-docs-rust/1.0")
        .build()
        .context("Failed building HTTP client")?;

    let resp = client
        .post(&config.token_uri)
        .form(&[
            ("client_id", config.client_id.as_str()),
            ("client_secret", config.client_secret.as_str()),
            ("refresh_token", refresh_token.as_str()),
            ("grant_type", "refresh_token"),
        ])
        .send()
        .context("Refresh token request failed")?;

    let status = resp.status();
    let body = resp
        .text()
        .context("Failed reading refresh token response body")?;

    if !status.is_success() {
        let msg = extract_google_error_message(&body)
            .unwrap_or_else(|| format!("Token refresh failed with status {status}"));
        return Err(anyhow!("{msg}"));
    }

    let payload: TokenResponse = serde_json::from_str(&body)
        .with_context(|| format!("Failed parsing refresh token response JSON. Body: {body}"))?;

    token.access_token = payload.access_token;
    token.expiration_time_millis = compute_expiration(payload.expires_in);
    if let Some(scope) = payload.scope {
        token.scope = Some(ScopeField::Multiple(
            scope.split_whitespace().map(ToString::to_string).collect(),
        ));
    }

    Ok(())
}

fn compute_expiration(expires_in: Option<i64>) -> i64 {
    let ttl_seconds = expires_in.unwrap_or(3600).max(1);
    Utc::now().timestamp_millis() + ttl_seconds * 1000
}

pub fn token_is_expired(token: &StoredToken) -> bool {
    let now = Utc::now().timestamp_millis();
    now >= (token.expiration_time_millis - 60_000)
}

pub fn extract_google_error_message(body: &str) -> Option<String> {
    let value = serde_json::from_str::<Value>(body).ok()?;
    value
        .get("error")
        .and_then(|e| {
            if e.is_string() {
                e.as_str().map(ToString::to_string)
            } else {
                e.get("message")
                    .and_then(|m| m.as_str())
                    .map(ToString::to_string)
            }
        })
        .or_else(|| {
            value
                .get("error_description")
                .and_then(|m| m.as_str())
                .map(ToString::to_string)
        })
}

pub fn auth_required_payload(auth_url: &str, message: &str, script_hint: &str) -> Value {
    json!({
      "status": "error",
      "error_code": "AUTH_REQUIRED",
      "message": message,
      "auth_url": auth_url,
      "instructions": [
        "1. Visit the authorization URL",
        "2. Grant access in your browser",
        "3. Copy the authorization code",
        format!("4. Run: {script_hint} auth <code>"),
        "5. Retry the original command"
      ]
    })
}
