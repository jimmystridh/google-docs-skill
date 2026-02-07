use anyhow::{Context, Result, anyhow};
use reqwest::Method;
use reqwest::blocking::{Client, Response, multipart};
use serde_json::Value;
use std::fs;
use std::path::Path;

#[derive(Debug, Clone)]
pub struct GoogleClient {
    http: Client,
    access_token: String,
}

#[derive(Debug, thiserror::Error)]
pub enum GoogleApiError {
    #[error("{message}")]
    Api {
        status: u16,
        message: String,
        body: Option<String>,
    },
    #[error("{0}")]
    Network(String),
    #[error("{0}")]
    Parse(String),
}

impl GoogleClient {
    pub fn new(access_token: impl Into<String>) -> Result<Self> {
        let http = Client::builder()
            .user_agent("google-docs-skill/1.0")
            .build()
            .context("Failed building HTTP client")?;
        Ok(Self {
            http,
            access_token: access_token.into(),
        })
    }

    pub fn get_json(
        &self,
        url: &str,
        query: &[(String, String)],
    ) -> std::result::Result<Value, GoogleApiError> {
        self.request_json(Method::GET, url, query, None)
    }

    pub fn post_json(
        &self,
        url: &str,
        query: &[(String, String)],
        body: &Value,
    ) -> std::result::Result<Value, GoogleApiError> {
        self.request_json(Method::POST, url, query, Some(body))
    }

    pub fn put_json(
        &self,
        url: &str,
        query: &[(String, String)],
        body: &Value,
    ) -> std::result::Result<Value, GoogleApiError> {
        self.request_json(Method::PUT, url, query, Some(body))
    }

    pub fn patch_json(
        &self,
        url: &str,
        query: &[(String, String)],
        body: &Value,
    ) -> std::result::Result<Value, GoogleApiError> {
        self.request_json(Method::PATCH, url, query, Some(body))
    }

    pub fn delete_no_content(
        &self,
        url: &str,
        query: &[(String, String)],
    ) -> std::result::Result<(), GoogleApiError> {
        let request = self
            .http
            .request(Method::DELETE, url)
            .bearer_auth(&self.access_token)
            .query(query);

        let response = request
            .send()
            .map_err(|e| GoogleApiError::Network(e.to_string()))?;

        if response.status().is_success() {
            return Ok(());
        }

        Err(error_from_response(response))
    }

    pub fn get_bytes_to_path(
        &self,
        url: &str,
        query: &[(String, String)],
        output_path: &Path,
    ) -> std::result::Result<(), GoogleApiError> {
        let request = self
            .http
            .request(Method::GET, url)
            .bearer_auth(&self.access_token)
            .query(query);

        let response = request
            .send()
            .map_err(|e| GoogleApiError::Network(e.to_string()))?;

        if !response.status().is_success() {
            return Err(error_from_response(response));
        }

        let bytes = response
            .bytes()
            .map_err(|e| GoogleApiError::Network(e.to_string()))?;

        if let Some(parent) = output_path.parent() {
            fs::create_dir_all(parent).map_err(|e| GoogleApiError::Network(e.to_string()))?;
        }

        fs::write(output_path, bytes).map_err(|e| GoogleApiError::Network(e.to_string()))?;
        Ok(())
    }

    pub fn post_multipart(
        &self,
        url: &str,
        query: &[(String, String)],
        metadata: &Value,
        file_path: &Path,
        mime_type: &str,
        file_name: &str,
    ) -> std::result::Result<Value, GoogleApiError> {
        let metadata_part = multipart::Part::text(metadata.to_string())
            .mime_str("application/json")
            .map_err(|e| GoogleApiError::Parse(e.to_string()))?;
        let file_bytes = fs::read(file_path).map_err(|e| GoogleApiError::Network(e.to_string()))?;
        let file_part = multipart::Part::bytes(file_bytes)
            .file_name(file_name.to_string())
            .mime_str(mime_type)
            .map_err(|e| GoogleApiError::Parse(e.to_string()))?;

        let form = multipart::Form::new()
            .part("metadata", metadata_part)
            .part("file", file_part);

        let response = self
            .http
            .post(url)
            .bearer_auth(&self.access_token)
            .query(query)
            .multipart(form)
            .send()
            .map_err(|e| GoogleApiError::Network(e.to_string()))?;

        parse_json_response(response)
    }

    pub fn patch_multipart(
        &self,
        url: &str,
        query: &[(String, String)],
        metadata: &Value,
        file_path: &Path,
        mime_type: &str,
        file_name: &str,
    ) -> std::result::Result<Value, GoogleApiError> {
        let metadata_part = multipart::Part::text(metadata.to_string())
            .mime_str("application/json")
            .map_err(|e| GoogleApiError::Parse(e.to_string()))?;
        let file_bytes = fs::read(file_path).map_err(|e| GoogleApiError::Network(e.to_string()))?;
        let file_part = multipart::Part::bytes(file_bytes)
            .file_name(file_name.to_string())
            .mime_str(mime_type)
            .map_err(|e| GoogleApiError::Parse(e.to_string()))?;

        let form = multipart::Form::new()
            .part("metadata", metadata_part)
            .part("file", file_part);

        let response = self
            .http
            .request(Method::PATCH, url)
            .bearer_auth(&self.access_token)
            .query(query)
            .multipart(form)
            .send()
            .map_err(|e| GoogleApiError::Network(e.to_string()))?;

        parse_json_response(response)
    }

    fn request_json(
        &self,
        method: Method,
        url: &str,
        query: &[(String, String)],
        body: Option<&Value>,
    ) -> std::result::Result<Value, GoogleApiError> {
        let mut request = self
            .http
            .request(method, url)
            .bearer_auth(&self.access_token)
            .query(query);

        if let Some(payload) = body {
            request = request.json(payload);
        }

        let response = request
            .send()
            .map_err(|e| GoogleApiError::Network(e.to_string()))?;

        parse_json_response(response)
    }
}

fn parse_json_response(response: Response) -> std::result::Result<Value, GoogleApiError> {
    if !response.status().is_success() {
        return Err(error_from_response(response));
    }

    let text = response
        .text()
        .map_err(|e| GoogleApiError::Network(e.to_string()))?;

    if text.trim().is_empty() {
        return Ok(Value::Null);
    }

    serde_json::from_str(&text).map_err(|e| GoogleApiError::Parse(e.to_string()))
}

fn error_from_response(response: Response) -> GoogleApiError {
    let status = response.status().as_u16();
    let body = response.text().ok();

    let message = body
        .as_deref()
        .and_then(extract_google_error_message)
        .unwrap_or_else(|| format!("Google API request failed with HTTP {status}"));

    GoogleApiError::Api {
        status,
        message,
        body,
    }
}

pub fn extract_google_error_message(body: &str) -> Option<String> {
    let value: Value = serde_json::from_str(body).ok()?;
    if let Some(msg) = value
        .get("error")
        .and_then(|e| e.get("message"))
        .and_then(|m| m.as_str())
    {
        return Some(msg.to_string());
    }
    if let Some(msg) = value.get("error_description").and_then(|m| m.as_str()) {
        return Some(msg.to_string());
    }
    None
}

pub fn map_api_error(operation: &str, err: &GoogleApiError) -> Value {
    match err {
        GoogleApiError::Api { message, body, .. } => {
            serde_json::json!({
                "status": "error",
                "error_code": "API_ERROR",
                "operation": operation,
                "message": format!("Google API error: {message}"),
                "details": body
            })
        }
        GoogleApiError::Network(message) | GoogleApiError::Parse(message) => {
            serde_json::json!({
                "status": "error",
                "error_code": "API_ERROR",
                "operation": operation,
                "message": format!("Google API error: {message}")
            })
        }
    }
}

pub fn ensure_file_exists(path: &Path) -> Result<()> {
    if path.exists() {
        Ok(())
    } else {
        Err(anyhow!("File not found: {}", path.display()))
    }
}

pub fn detect_drive_mime_type(path: &Path) -> &'static str {
    match path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase()
        .as_str()
    {
        "excalidraw" => "application/json",
        "json" => "application/json",
        "txt" => "text/plain",
        "md" => "text/markdown",
        "html" | "htm" => "text/html",
        "css" => "text/css",
        "js" => "application/javascript",
        "pdf" => "application/pdf",
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "svg" => "image/svg+xml",
        "zip" => "application/zip",
        "csv" => "text/csv",
        "xml" => "application/xml",
        "yaml" | "yml" => "application/x-yaml",
        _ => "application/octet-stream",
    }
}
