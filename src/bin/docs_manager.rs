use anyhow::{Context, Result};
use google_docs_rust::auth::{
    AuthPaths, SHARED_SCOPES, TokenState, auth_required_payload, build_auth_url,
    complete_authorization, ensure_token, load_oauth_client_config, load_stored_token,
    save_stored_token,
};
use google_docs_rust::google_api::{GoogleApiError, GoogleClient, map_api_error};
use google_docs_rust::io_helpers::{home_dir, print_json, read_stdin_json};
use serde_json::{Value, json};
use std::env;

const EXIT_SUCCESS: i32 = 0;
const EXIT_AUTH_ERROR: i32 = 2;
const EXIT_API_ERROR: i32 = 3;
const EXIT_INVALID_ARGS: i32 = 4;

#[derive(Debug, Clone)]
enum FormatType {
    Heading1,
    Heading2,
    Heading3,
    Bold,
    Italic,
    Code,
}

#[derive(Debug, Clone)]
struct FormatInfo {
    format_type: FormatType,
    start: i64,
    end: i64,
}

#[derive(Debug, Clone)]
struct TableInfo {
    rows: Vec<Vec<String>>,
    insert_index: i64,
    num_rows: i64,
    num_cols: i64,
}

#[derive(Debug, Clone)]
struct ParsedMarkdown {
    text: String,
    formats: Vec<FormatInfo>,
    tables: Vec<TableInfo>,
}

fn main() {
    let args: Vec<String> = env::args().collect();
    let program = args
        .first()
        .map(|p| {
            std::path::Path::new(p)
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("docs_manager")
                .to_string()
        })
        .unwrap_or_else(|| "docs_manager".to_string());

    if args.len() < 2 {
        usage(&program);
        std::process::exit(EXIT_INVALID_ARGS);
    }

    let command = args[1].as_str();

    if command == "--help" || command == "-h" {
        usage(&program);
        std::process::exit(EXIT_SUCCESS);
    }

    if command == "auth" {
        if args.len() < 3 {
            print_json(&json!({
                "status": "error",
                "error_code": "MISSING_CODE",
                "message": "Authorization code required",
                "usage": format!("{program} auth <code>")
            }));
            std::process::exit(EXIT_INVALID_ARGS);
        }

        if let Err(err) = complete_auth(&program, &args[2]) {
            print_json(&json!({
                "status": "error",
                "error_code": "AUTH_FAILED",
                "message": format!("Authorization failed: {err}")
            }));
            std::process::exit(EXIT_AUTH_ERROR);
        }

        std::process::exit(EXIT_SUCCESS);
    }

    let client = match initialize_client(
        &program,
        "Authorization required. Please visit the URL and enter the code.",
    ) {
        Ok(client) => client,
        Err(exit_code) => std::process::exit(exit_code),
    };

    let exit_code = match command {
        "read" => {
            if args.len() < 3 {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_DOCUMENT_ID",
                    "message": "Document ID required"
                }));
                EXIT_INVALID_ARGS
            } else {
                match read_document(&client, &args[2]) {
                    Ok(payload) => {
                        print_json(&payload);
                        EXIT_SUCCESS
                    }
                    Err(err) => handle_google_error("read", &err),
                }
            }
        }
        "structure" => {
            if args.len() < 3 {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_DOCUMENT_ID",
                    "message": "Document ID required"
                }));
                EXIT_INVALID_ARGS
            } else {
                match get_structure(&client, &args[2]) {
                    Ok(payload) => {
                        print_json(&payload);
                        EXIT_SUCCESS
                    }
                    Err(err) => handle_google_error("structure", &err),
                }
            }
        }
        "insert" => dispatch_json_command("insert", || {
            let input = read_stdin_json()?;
            let document_id = required_string(&input, "document_id")?;
            let text = required_string(&input, "text")?;
            let index = input.get("index").and_then(value_to_i64).unwrap_or(1);
            insert_text(&client, &document_id, &text, index)
        }),
        "append" => dispatch_json_command("append", || {
            let input = read_stdin_json()?;
            let document_id = required_string(&input, "document_id")?;
            let text = required_string(&input, "text")?;
            append_text(&client, &document_id, &text)
        }),
        "replace" => dispatch_json_command("replace", || {
            let input = read_stdin_json()?;
            let document_id = required_string(&input, "document_id")?;
            let find = required_string(&input, "find")?;
            let replace = required_string(&input, "replace")?;
            let match_case = input
                .get("match_case")
                .and_then(|v| v.as_bool())
                .unwrap_or(false);
            replace_text(&client, &document_id, &find, &replace, match_case)
        }),
        "format" => dispatch_json_command("format", || {
            let input = read_stdin_json()?;
            let document_id = required_string(&input, "document_id")?;
            let start_index = required_i64(&input, "start_index")?;
            let end_index = required_i64(&input, "end_index")?;
            let bold = input.get("bold").and_then(|v| v.as_bool());
            let italic = input.get("italic").and_then(|v| v.as_bool());
            let underline = input.get("underline").and_then(|v| v.as_bool());
            format_text(
                &client,
                &document_id,
                start_index,
                end_index,
                bold,
                italic,
                underline,
            )
        }),
        "page-break" => dispatch_json_command("page_break", || {
            let input = read_stdin_json()?;
            let document_id = required_string(&input, "document_id")?;
            let index = required_i64(&input, "index")?;
            insert_page_break(&client, &document_id, index)
        }),
        "create" => dispatch_json_command("create", || {
            let input = read_stdin_json()?;
            let title = required_string(&input, "title")?;
            let content = input
                .get("content")
                .and_then(|v| v.as_str())
                .map(ToString::to_string);
            create_document(&client, &title, content)
        }),
        "create-from-markdown" => dispatch_json_command("create_from_markdown", || {
            let input = read_stdin_json()?;
            let title = required_string(&input, "title")?;
            let markdown = required_string(&input, "markdown")?;
            create_from_markdown(&client, &title, &markdown)
        }),
        "insert-from-markdown" => dispatch_json_command("insert_from_markdown", || {
            let input = read_stdin_json()?;
            let document_id = required_string(&input, "document_id")?;
            let markdown = required_string(&input, "markdown")?;
            let index = input.get("index").and_then(value_to_i64);
            insert_from_markdown(&client, &document_id, &markdown, index)
        }),
        "delete" => dispatch_json_command("delete", || {
            let input = read_stdin_json()?;
            let document_id = required_string(&input, "document_id")?;
            let start_index = required_i64(&input, "start_index")?;
            let end_index = required_i64(&input, "end_index")?;
            delete_content(&client, &document_id, start_index, end_index)
        }),
        "insert-image" => dispatch_json_command("insert_image", || {
            let input = read_stdin_json()?;
            let document_id = required_string(&input, "document_id")?;
            let image_url = required_string(&input, "image_url")?;
            let index = input.get("index").and_then(value_to_i64);
            let width = input.get("width").and_then(value_to_f64);
            let height = input.get("height").and_then(value_to_f64);
            insert_image(&client, &document_id, &image_url, index, width, height)
        }),
        "insert-table" => dispatch_json_command("insert_table", || {
            let input = read_stdin_json()?;
            let document_id = required_string(&input, "document_id")?;
            let rows = required_i64(&input, "rows")?;
            let cols = required_i64(&input, "cols")?;
            let index = input.get("index").and_then(value_to_i64);
            let data = input
                .get("data")
                .and_then(|v| v.as_array())
                .cloned()
                .unwrap_or_default();
            insert_table(&client, &document_id, rows, cols, index, &data)
        }),
        _ => {
            print_json(&json!({
                "status": "error",
                "error_code": "INVALID_COMMAND",
                "message": format!("Unknown command: {command}"),
                "valid_commands": [
                    "auth",
                    "read",
                    "structure",
                    "insert",
                    "append",
                    "replace",
                    "format",
                    "page-break",
                    "create",
                    "create-from-markdown",
                    "insert-from-markdown",
                    "delete",
                    "insert-image",
                    "insert-table"
                ]
            }));
            usage(&program);
            EXIT_INVALID_ARGS
        }
    };

    std::process::exit(exit_code);
}

fn complete_auth(program: &str, code: &str) -> Result<()> {
    let home = home_dir()?;
    let paths = AuthPaths::from_home(&home);
    let config = load_oauth_client_config(&paths.credentials_path)?;
    let existing_refresh = load_stored_token(&paths.token_path)
        .ok()
        .and_then(|t| t.refresh_token.clone());
    let token = complete_authorization(&config, code, existing_refresh)?;
    save_stored_token(&paths.token_path, &token)?;

    print_json(&json!({
        "status": "success",
        "message": "Authorization complete. Token stored successfully.",
        "token_path": paths.token_path.display().to_string(),
        "scopes": SHARED_SCOPES
    }));

    let _ = program;
    Ok(())
}

fn initialize_client(program: &str, auth_message: &str) -> std::result::Result<GoogleClient, i32> {
    let home = match home_dir() {
        Ok(h) => h,
        Err(err) => {
            print_json(&json!({
                "status": "error",
                "error_code": "AUTH_FAILED",
                "message": format!("Authorization setup failed: {err}")
            }));
            return Err(EXIT_AUTH_ERROR);
        }
    };
    let paths = AuthPaths::from_home(&home);

    match ensure_token(&paths, SHARED_SCOPES) {
        Ok(TokenState::AuthorizationRequired { auth_url }) => {
            print_json(&auth_required_payload(&auth_url, auth_message, program));
            Err(EXIT_AUTH_ERROR)
        }
        Ok(TokenState::Authorized(token)) => match GoogleClient::new(token.access_token) {
            Ok(client) => Ok(client),
            Err(err) => {
                print_json(&json!({
                    "status": "error",
                    "error_code": "AUTH_FAILED",
                    "message": format!("Failed to initialize API client: {err}")
                }));
                Err(EXIT_AUTH_ERROR)
            }
        },
        Err(err) => {
            let auth_url = load_oauth_client_config(&paths.credentials_path)
                .ok()
                .and_then(|cfg| build_auth_url(&cfg, SHARED_SCOPES).ok());

            if let Some(url) = auth_url {
                print_json(&auth_required_payload(&url, auth_message, program));
            } else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "AUTH_FAILED",
                    "message": format!("Authorization failed: {err}")
                }));
            }
            Err(EXIT_AUTH_ERROR)
        }
    }
}

fn usage(program: &str) {
    println!(
        "Google Docs Manager - Document Operations CLI\n\nUsage:\n  {program} <command> [options]\n\nCommands:\n  auth <code>              Complete OAuth authorization with code\n  read <document_id>       Read document content\n  structure <document_id>  Get document structure (headings)\n  insert                   Insert text at specific index (JSON via stdin)\n  append                   Append text to end of document (JSON via stdin)\n  replace                  Find and replace text (JSON via stdin)\n  format                   Format text (JSON via stdin)\n  page-break               Insert page break (JSON via stdin)\n  create                   Create new document (JSON via stdin)\n  create-from-markdown     Create new document from markdown (JSON via stdin)\n  insert-from-markdown     Insert formatted markdown into existing doc (JSON via stdin)\n  delete                   Delete content range (JSON via stdin)\n  insert-image             Insert inline image from URL (JSON via stdin)\n  insert-table             Insert table (JSON via stdin)\n\nExit Codes:\n  0 - Success\n  1 - Operation failed\n  2 - Authentication error\n  3 - API error\n  4 - Invalid arguments"
    );
}

fn handle_google_error(operation: &str, err: &GoogleApiError) -> i32 {
    print_json(&map_api_error(operation, err));
    EXIT_API_ERROR
}

fn dispatch_json_command<F>(operation: &str, f: F) -> i32
where
    F: FnOnce() -> Result<Value>,
{
    match f() {
        Ok(payload) => {
            print_json(&payload);
            EXIT_SUCCESS
        }
        Err(err) => {
            if let Some(api_err) = err.downcast_ref::<GoogleApiError>() {
                return handle_google_error(operation, api_err);
            }

            print_json(&json!({
                "status": "error",
                "error_code": "MISSING_REQUIRED_FIELDS",
                "message": err.to_string()
            }));
            EXIT_INVALID_ARGS
        }
    }
}

fn required_string(input: &Value, key: &str) -> Result<String> {
    input
        .get(key)
        .and_then(|v| v.as_str())
        .map(ToString::to_string)
        .ok_or_else(|| anyhow::anyhow!(required_fields_message(&[key])))
}

fn required_i64(input: &Value, key: &str) -> Result<i64> {
    input
        .get(key)
        .and_then(value_to_i64)
        .ok_or_else(|| anyhow::anyhow!(required_fields_message(&[key])))
}

fn required_fields_message(fields: &[&str]) -> String {
    if fields.len() == 1 {
        format!("Required field: {}", fields[0])
    } else {
        format!("Required fields: {}", fields.join(", "))
    }
}

fn value_to_i64(value: &Value) -> Option<i64> {
    if let Some(v) = value.as_i64() {
        return Some(v);
    }
    if let Some(v) = value.as_u64() {
        return i64::try_from(v).ok();
    }
    if let Some(v) = value.as_f64() {
        return Some(v as i64);
    }
    None
}

fn value_to_f64(value: &Value) -> Option<f64> {
    if let Some(v) = value.as_f64() {
        Some(v)
    } else if let Some(v) = value.as_i64() {
        Some(v as f64)
    } else {
        value.as_u64().map(|v| v as f64)
    }
}

fn read_document(
    client: &GoogleClient,
    document_id: &str,
) -> std::result::Result<Value, GoogleApiError> {
    let document = get_document(client, document_id)?;
    let content = document
        .get("body")
        .and_then(|b| b.get("content"))
        .and_then(|c| c.as_array())
        .map(|items| extract_text_content(items))
        .unwrap_or_default();

    Ok(json!({
        "status": "success",
        "operation": "read",
        "document_id": document.get("documentId").and_then(|v| v.as_str()),
        "title": document.get("title").and_then(|v| v.as_str()),
        "content": content,
        "revision_id": document.get("revisionId").and_then(|v| v.as_str())
    }))
}

fn get_structure(
    client: &GoogleClient,
    document_id: &str,
) -> std::result::Result<Value, GoogleApiError> {
    let document = get_document(client, document_id)?;
    let mut structure = Vec::new();

    if let Some(elements) = document
        .get("body")
        .and_then(|b| b.get("content"))
        .and_then(|c| c.as_array())
    {
        for element in elements {
            let Some(paragraph) = element.get("paragraph") else {
                continue;
            };
            let Some(style) = paragraph
                .get("paragraphStyle")
                .and_then(|s| s.get("namedStyleType"))
                .and_then(|s| s.as_str())
            else {
                continue;
            };

            if !style.starts_with("HEADING_") {
                continue;
            }

            let level = style
                .rsplit('_')
                .next()
                .and_then(|n| n.parse::<i64>().ok())
                .unwrap_or(0);
            let text = extract_paragraph_text(paragraph);

            structure.push(json!({
                "level": level,
                "text": text,
                "start_index": element.get("startIndex").and_then(value_to_i64),
                "end_index": element.get("endIndex").and_then(value_to_i64)
            }));
        }
    }

    Ok(json!({
        "status": "success",
        "operation": "structure",
        "document_id": document.get("documentId").and_then(|v| v.as_str()),
        "title": document.get("title").and_then(|v| v.as_str()),
        "structure": structure
    }))
}

fn insert_text(client: &GoogleClient, document_id: &str, text: &str, index: i64) -> Result<Value> {
    let requests = vec![json!({
        "insertText": {
            "location": { "index": index },
            "text": text
        }
    })];
    let result = docs_batch_update(client, document_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "insert",
        "document_id": document_id,
        "inserted_at": index,
        "text_length": text.chars().count(),
        "revision_id": result.get("documentId").and_then(|v| v.as_str())
    }))
}

fn append_text(client: &GoogleClient, document_id: &str, text: &str) -> Result<Value> {
    let document = get_document(client, document_id)?;
    let end_index = last_body_end_index(&document).unwrap_or(1) - 1;
    let requests = vec![json!({
        "insertText": {
            "location": { "index": end_index },
            "text": text
        }
    })];

    let result = docs_batch_update(client, document_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "append",
        "document_id": document_id,
        "appended_at": end_index,
        "text_length": text.chars().count(),
        "revision_id": result.get("documentId").and_then(|v| v.as_str())
    }))
}

fn replace_text(
    client: &GoogleClient,
    document_id: &str,
    find: &str,
    replace: &str,
    match_case: bool,
) -> Result<Value> {
    let requests = vec![json!({
        "replaceAllText": {
            "containsText": {
                "text": find,
                "matchCase": match_case
            },
            "replaceText": replace
        }
    })];

    let result = docs_batch_update(client, document_id, requests)?;
    let occurrences = result
        .get("replies")
        .and_then(|r| r.as_array())
        .and_then(|r| r.first())
        .and_then(|r| r.get("replaceAllText"))
        .and_then(|r| r.get("occurrencesChanged"))
        .and_then(value_to_i64)
        .unwrap_or(0);

    Ok(json!({
        "status": "success",
        "operation": "replace",
        "document_id": document_id,
        "find": find,
        "replace": replace,
        "occurrences": occurrences
    }))
}

fn format_text(
    client: &GoogleClient,
    document_id: &str,
    start_index: i64,
    end_index: i64,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
) -> Result<Value> {
    let mut style = serde_json::Map::new();
    let mut fields = Vec::new();

    if let Some(v) = bold {
        style.insert("bold".to_string(), Value::Bool(v));
        fields.push("bold");
    }
    if let Some(v) = italic {
        style.insert("italic".to_string(), Value::Bool(v));
        fields.push("italic");
    }
    if let Some(v) = underline {
        style.insert("underline".to_string(), Value::Bool(v));
        fields.push("underline");
    }

    let requests = vec![json!({
        "updateTextStyle": {
            "range": {
                "startIndex": start_index,
                "endIndex": end_index
            },
            "textStyle": Value::Object(style.clone()),
            "fields": fields.join(",")
        }
    })];

    let _ = docs_batch_update(client, document_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "format",
        "document_id": document_id,
        "range": {"start": start_index, "end": end_index},
        "formatting": Value::Object(style)
    }))
}

fn insert_page_break(client: &GoogleClient, document_id: &str, index: i64) -> Result<Value> {
    let requests = vec![json!({
        "insertPageBreak": {
            "location": { "index": index }
        }
    })];

    let _ = docs_batch_update(client, document_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "page_break",
        "document_id": document_id,
        "inserted_at": index
    }))
}

fn insert_image(
    client: &GoogleClient,
    document_id: &str,
    image_url: &str,
    index: Option<i64>,
    width: Option<f64>,
    height: Option<f64>,
) -> Result<Value> {
    let insertion_index = match index {
        Some(i) => i,
        None => {
            let doc = get_document(client, document_id)?;
            last_body_end_index(&doc).unwrap_or(1) - 1
        }
    };

    let mut insert_inline_image = json!({
        "location": { "index": insertion_index },
        "uri": image_url
    });

    if width.is_some() || height.is_some() {
        let mut size = serde_json::Map::new();
        if let Some(w) = width {
            size.insert(
                "width".to_string(),
                json!({
                    "magnitude": w,
                    "unit": "PT"
                }),
            );
        }
        if let Some(h) = height {
            size.insert(
                "height".to_string(),
                json!({
                    "magnitude": h,
                    "unit": "PT"
                }),
            );
        }

        insert_inline_image
            .as_object_mut()
            .expect("object")
            .insert("objectSize".to_string(), Value::Object(size));
    }

    let requests = vec![json!({
        "insertInlineImage": insert_inline_image
    })];

    let result = docs_batch_update(client, document_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "insert_image",
        "document_id": document_id,
        "inserted_at": insertion_index,
        "image_url": image_url,
        "revision_id": result.get("documentId").and_then(|v| v.as_str())
    }))
}

fn create_document(client: &GoogleClient, title: &str, content: Option<String>) -> Result<Value> {
    let result = client
        .post_json(
            "https://docs.googleapis.com/v1/documents",
            &[],
            &json!({"title": title}),
        )
        .map_err(anyhow::Error::from)?;

    let document_id = result
        .get("documentId")
        .and_then(|v| v.as_str())
        .context("Failed to parse documentId from create response")?
        .to_string();

    if let Some(content) = content {
        let requests = vec![json!({
            "insertText": {
                "location": { "index": 1 },
                "text": content
            }
        })];
        let _ = docs_batch_update(client, &document_id, requests)?;
    }

    Ok(json!({
        "status": "success",
        "operation": "create",
        "document_id": document_id,
        "title": result.get("title").and_then(|v| v.as_str()),
        "revision_id": result.get("revisionId").and_then(|v| v.as_str())
    }))
}

fn delete_content(
    client: &GoogleClient,
    document_id: &str,
    start_index: i64,
    end_index: i64,
) -> Result<Value> {
    let requests = vec![json!({
        "deleteContentRange": {
            "range": {
                "startIndex": start_index,
                "endIndex": end_index
            }
        }
    })];

    let _ = docs_batch_update(client, document_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "delete",
        "document_id": document_id,
        "deleted_range": {"start": start_index, "end": end_index}
    }))
}

fn insert_table(
    client: &GoogleClient,
    document_id: &str,
    rows: i64,
    cols: i64,
    index: Option<i64>,
    data: &[Value],
) -> Result<Value> {
    let insertion_index = match index {
        Some(i) => i,
        None => {
            let document = get_document(client, document_id)?;
            last_body_end_index(&document).unwrap_or(1) - 1
        }
    };

    insert_table_internal(client, document_id, rows, cols, insertion_index, data)?;

    Ok(json!({
        "status": "success",
        "operation": "insert_table",
        "document_id": document_id,
        "rows": rows,
        "columns": cols,
        "inserted_at": insertion_index
    }))
}

fn insert_table_internal(
    client: &GoogleClient,
    document_id: &str,
    rows: i64,
    cols: i64,
    index: i64,
    data: &[Value],
) -> Result<()> {
    let insert_requests = vec![json!({
        "insertTable": {
            "rows": rows,
            "columns": cols,
            "location": { "index": index }
        }
    })];

    let _ = docs_batch_update(client, document_id, insert_requests)?;

    if data.is_empty() {
        return Ok(());
    }

    let document = get_document(client, document_id)?;
    let table_element = document
        .get("body")
        .and_then(|b| b.get("content"))
        .and_then(|c| c.as_array())
        .and_then(|items| {
            items.iter().find(|element| {
                element.get("table").is_some()
                    && element
                        .get("startIndex")
                        .and_then(value_to_i64)
                        .map(|v| v >= index)
                        .unwrap_or(false)
            })
        })
        .cloned();

    let Some(table_element) = table_element else {
        return Ok(());
    };

    let table_rows = table_element
        .get("table")
        .and_then(|t| t.get("tableRows"))
        .and_then(|r| r.as_array())
        .cloned()
        .unwrap_or_default();

    let mut cell_requests = Vec::new();

    for row_idx in (0..data.len()).rev() {
        if row_idx as i64 >= rows {
            continue;
        }
        let row_data = data
            .get(row_idx)
            .and_then(|v| v.as_array())
            .cloned()
            .unwrap_or_default();

        for col_idx in (0..row_data.len()).rev() {
            if col_idx as i64 >= cols {
                continue;
            }

            let Some(table_row) = table_rows.get(row_idx) else {
                continue;
            };

            let Some(table_cells) = table_row.get("tableCells").and_then(|v| v.as_array()) else {
                continue;
            };

            let Some(table_cell) = table_cells.get(col_idx) else {
                continue;
            };

            let cell_start = table_cell
                .get("content")
                .and_then(|v| v.as_array())
                .and_then(|items| items.first())
                .and_then(|first| first.get("startIndex"))
                .and_then(value_to_i64);

            let Some(cell_start) = cell_start else {
                continue;
            };

            let text = row_data
                .get(col_idx)
                .map(value_to_string)
                .unwrap_or_default();

            cell_requests.push(json!({
                "insertText": {
                    "location": {"index": cell_start},
                    "text": text
                }
            }));
        }
    }

    if !cell_requests.is_empty() {
        let _ = docs_batch_update(client, document_id, cell_requests)?;
    }

    Ok(())
}

fn create_from_markdown(client: &GoogleClient, title: &str, markdown: &str) -> Result<Value> {
    let create = client
        .post_json(
            "https://docs.googleapis.com/v1/documents",
            &[],
            &json!({"title": title}),
        )
        .map_err(anyhow::Error::from)?;

    let document_id = create
        .get("documentId")
        .and_then(|v| v.as_str())
        .context("Failed to parse documentId from create response")?
        .to_string();

    let parsed = parse_markdown(markdown);

    if !parsed.text.is_empty() {
        let requests = vec![json!({
            "insertText": {
                "location": { "index": 1 },
                "text": parsed.text.clone()
            }
        })];
        let _ = docs_batch_update(client, &document_id, requests)?;
    }

    let mut format_requests = Vec::new();
    for fmt in parsed.formats.iter().rev() {
        if let Some(req) = build_format_request(fmt) {
            format_requests.push(req);
        }
    }

    if !format_requests.is_empty() {
        let _ = docs_batch_update(client, &document_id, format_requests)?;
    }

    for table in parsed.tables.iter().rev() {
        let data: Vec<Value> = table
            .rows
            .iter()
            .map(|r| Value::Array(r.iter().map(|cell| Value::String(cell.clone())).collect()))
            .collect();
        insert_table_internal(
            client,
            &document_id,
            table.num_rows,
            table.num_cols,
            table.insert_index,
            &data,
        )?;
    }

    Ok(json!({
        "status": "success",
        "operation": "create_from_markdown",
        "document_id": document_id,
        "title": title,
        "revision_id": create.get("revisionId").and_then(|v| v.as_str()),
        "tables_inserted": parsed.tables.len()
    }))
}

fn insert_from_markdown(
    client: &GoogleClient,
    document_id: &str,
    markdown: &str,
    index: Option<i64>,
) -> Result<Value> {
    let insertion_index = match index {
        Some(v) => v,
        None => {
            let document = get_document(client, document_id)?;
            last_body_end_index(&document).unwrap_or(1) - 1
        }
    };

    let parsed = parse_markdown(markdown);

    if !parsed.text.is_empty() {
        let requests = vec![json!({
            "insertText": {
                "location": {"index": insertion_index},
                "text": parsed.text.clone()
            }
        })];
        let _ = docs_batch_update(client, document_id, requests)?;
    }

    let offset = insertion_index - 1;
    let adjusted_formats: Vec<FormatInfo> = parsed
        .formats
        .iter()
        .map(|fmt| FormatInfo {
            format_type: fmt.format_type.clone(),
            start: fmt.start + offset,
            end: fmt.end + offset,
        })
        .collect();

    let mut requests = Vec::new();
    for fmt in adjusted_formats.iter().rev() {
        if let Some(req) = build_format_request(fmt) {
            requests.push(req);
        }
    }

    if !requests.is_empty() {
        let _ = docs_batch_update(client, document_id, requests)?;
    }

    Ok(json!({
        "status": "success",
        "operation": "insert_from_markdown",
        "document_id": document_id,
        "inserted_at": insertion_index,
        "text_length": parsed.text.chars().count(),
        "formats_applied": parsed.formats.len()
    }))
}

fn parse_markdown(markdown: &str) -> ParsedMarkdown {
    let mut text = String::new();
    let mut formats = Vec::new();
    let mut tables = Vec::new();
    let mut current_index: i64 = 1;

    let lines: Vec<&str> = markdown.lines().collect();
    let mut i = 0usize;

    while i < lines.len() {
        let line = lines[i].trim_end();

        if let Some(rest) = line.strip_prefix("# ") {
            let heading = format!("{rest}\n");
            formats.push(FormatInfo {
                format_type: FormatType::Heading1,
                start: current_index,
                end: current_index + char_len(&heading) - 1,
            });
            text.push_str(&heading);
            current_index += char_len(&heading);
        } else if let Some(rest) = line.strip_prefix("## ") {
            let heading = format!("{rest}\n");
            formats.push(FormatInfo {
                format_type: FormatType::Heading2,
                start: current_index,
                end: current_index + char_len(&heading) - 1,
            });
            text.push_str(&heading);
            current_index += char_len(&heading);
        } else if let Some(rest) = line.strip_prefix("### ") {
            let heading = format!("{rest}\n");
            formats.push(FormatInfo {
                format_type: FormatType::Heading3,
                start: current_index,
                end: current_index + char_len(&heading) - 1,
            });
            text.push_str(&heading);
            current_index += char_len(&heading);
        } else if line.starts_with("- [ ] ") || line.starts_with("* [ ] ") {
            let item = &line[6..];
            let prefix = "☐ ";
            let processed =
                process_inline_formatting(item, current_index + char_len(prefix), &mut formats);
            let rendered = format!("{prefix}{processed}\n");
            text.push_str(&rendered);
            current_index += char_len(&rendered);
        } else if line.starts_with("- [x] ")
            || line.starts_with("* [x] ")
            || line.starts_with("- [X] ")
            || line.starts_with("* [X] ")
        {
            let item = &line[6..];
            let prefix = "☑ ";
            let processed =
                process_inline_formatting(item, current_index + char_len(prefix), &mut formats);
            let rendered = format!("{prefix}{processed}\n");
            text.push_str(&rendered);
            current_index += char_len(&rendered);
        } else if let Some(item) = line.strip_prefix("- ").or_else(|| line.strip_prefix("* ")) {
            let prefix = "• ";
            let processed =
                process_inline_formatting(item, current_index + char_len(prefix), &mut formats);
            let rendered = format!("{prefix}{processed}\n");
            text.push_str(&rendered);
            current_index += char_len(&rendered);
        } else if let Some((num, item)) = parse_numbered_list_item(line) {
            let prefix = format!("{num}. ");
            let processed =
                process_inline_formatting(&item, current_index + char_len(&prefix), &mut formats);
            let rendered = format!("{prefix}{processed}\n");
            text.push_str(&rendered);
            current_index += char_len(&rendered);
        } else if line == "---" {
            let hr = "———————————————————————————\n";
            text.push_str(hr);
            current_index += char_len(hr);
        } else if line.starts_with('|') && line.ends_with('|') {
            let mut table_rows: Vec<Vec<String>> = Vec::new();
            while i < lines.len() {
                let current = lines[i].trim_end();
                if !(current.starts_with('|') && current.ends_with('|')) {
                    break;
                }
                let cells = current[1..current.len() - 1]
                    .split('|')
                    .map(|c| c.trim().to_string())
                    .collect::<Vec<_>>();

                let separator = !cells.is_empty()
                    && cells
                        .iter()
                        .all(|c| !c.is_empty() && c.chars().all(|ch| ch == '-' || ch == ':'));
                if !separator {
                    table_rows.push(cells);
                }

                i += 1;
            }
            i = i.saturating_sub(1);

            if !table_rows.is_empty() {
                let num_rows = table_rows.len() as i64;
                let num_cols = table_rows.first().map(|r| r.len()).unwrap_or(0) as i64;
                tables.push(TableInfo {
                    rows: table_rows,
                    insert_index: current_index,
                    num_rows,
                    num_cols,
                });
                text.push('\n');
                current_index += 1;
            }
        } else if line.is_empty() {
            text.push('\n');
            current_index += 1;
        } else {
            let processed = process_inline_formatting(line, current_index, &mut formats);
            let rendered = format!("{processed}\n");
            text.push_str(&rendered);
            current_index += char_len(&rendered);
        }

        i += 1;
    }

    ParsedMarkdown {
        text,
        formats,
        tables,
    }
}

fn parse_numbered_list_item(line: &str) -> Option<(String, String)> {
    let dot = line.find('.')?;
    let (num, rest) = line.split_at(dot);
    if num.is_empty() || !num.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    let rest = rest.strip_prefix(". ")?;
    Some((num.to_string(), rest.to_string()))
}

fn process_inline_formatting(line: &str, base_index: i64, formats: &mut Vec<FormatInfo>) -> String {
    let mut result = String::new();
    let mut pos = 0usize;

    while pos < line.len() {
        if line[pos..].starts_with("**") {
            let search_start = pos + 2;
            if search_start <= line.len()
                && let Some(rel_end) = line[search_start..].find("**")
            {
                let end = search_start + rel_end;
                let bold_text = &line[search_start..end];
                let start_idx = base_index + char_len(&result);
                result.push_str(bold_text);
                formats.push(FormatInfo {
                    format_type: FormatType::Bold,
                    start: start_idx,
                    end: start_idx + char_len(bold_text),
                });
                pos = end + 2;
                continue;
            }
        }

        if line[pos..].starts_with('*') && !line[pos..].starts_with("**") {
            let search_start = pos + 1;
            if search_start <= line.len()
                && let Some(rel_end) = line[search_start..].find('*')
            {
                let end = search_start + rel_end;
                if !line[end..].starts_with("**") {
                    let italic_text = &line[search_start..end];
                    let start_idx = base_index + char_len(&result);
                    result.push_str(italic_text);
                    formats.push(FormatInfo {
                        format_type: FormatType::Italic,
                        start: start_idx,
                        end: start_idx + char_len(italic_text),
                    });
                    pos = end + 1;
                    continue;
                }
            }
        }

        if line[pos..].starts_with('`') {
            let search_start = pos + 1;
            if search_start <= line.len()
                && let Some(rel_end) = line[search_start..].find('`')
            {
                let end = search_start + rel_end;
                let code_text = &line[search_start..end];
                let start_idx = base_index + char_len(&result);
                result.push_str(code_text);
                formats.push(FormatInfo {
                    format_type: FormatType::Code,
                    start: start_idx,
                    end: start_idx + char_len(code_text),
                });
                pos = end + 1;
                continue;
            }
        }

        if let Some(ch) = line[pos..].chars().next() {
            result.push(ch);
            pos += ch.len_utf8();
        } else {
            break;
        }
    }

    result
}

fn build_format_request(fmt: &FormatInfo) -> Option<Value> {
    match fmt.format_type {
        FormatType::Heading1 => Some(json!({
            "updateParagraphStyle": {
                "range": {"startIndex": fmt.start, "endIndex": fmt.end},
                "paragraphStyle": {"namedStyleType": "HEADING_1"},
                "fields": "namedStyleType"
            }
        })),
        FormatType::Heading2 => Some(json!({
            "updateParagraphStyle": {
                "range": {"startIndex": fmt.start, "endIndex": fmt.end},
                "paragraphStyle": {"namedStyleType": "HEADING_2"},
                "fields": "namedStyleType"
            }
        })),
        FormatType::Heading3 => Some(json!({
            "updateParagraphStyle": {
                "range": {"startIndex": fmt.start, "endIndex": fmt.end},
                "paragraphStyle": {"namedStyleType": "HEADING_3"},
                "fields": "namedStyleType"
            }
        })),
        FormatType::Bold => Some(json!({
            "updateTextStyle": {
                "range": {"startIndex": fmt.start, "endIndex": fmt.end},
                "textStyle": {"bold": true},
                "fields": "bold"
            }
        })),
        FormatType::Italic => Some(json!({
            "updateTextStyle": {
                "range": {"startIndex": fmt.start, "endIndex": fmt.end},
                "textStyle": {"italic": true},
                "fields": "italic"
            }
        })),
        FormatType::Code => Some(json!({
            "updateTextStyle": {
                "range": {"startIndex": fmt.start, "endIndex": fmt.end},
                "textStyle": {
                    "fontFamily": "Courier New",
                    "backgroundColor": {
                        "color": {
                            "rgbColor": {"red": 0.95, "green": 0.95, "blue": 0.95}
                        }
                    }
                },
                "fields": "fontFamily,backgroundColor"
            }
        })),
    }
}

fn char_len(text: &str) -> i64 {
    text.chars().count() as i64
}

fn value_to_string(value: &Value) -> String {
    match value {
        Value::String(s) => s.clone(),
        Value::Null => String::new(),
        other => other.to_string(),
    }
}

fn get_document(
    client: &GoogleClient,
    document_id: &str,
) -> std::result::Result<Value, GoogleApiError> {
    let url = format!("https://docs.googleapis.com/v1/documents/{document_id}");
    client.get_json(&url, &[])
}

fn docs_batch_update(
    client: &GoogleClient,
    document_id: &str,
    requests: Vec<Value>,
) -> Result<Value> {
    let url = format!("https://docs.googleapis.com/v1/documents/{document_id}:batchUpdate");
    let payload = json!({ "requests": requests });
    client
        .post_json(&url, &[], &payload)
        .map_err(anyhow::Error::from)
}

fn last_body_end_index(document: &Value) -> Option<i64> {
    document
        .get("body")
        .and_then(|b| b.get("content"))
        .and_then(|c| c.as_array())
        .and_then(|content| content.last())
        .and_then(|element| element.get("endIndex"))
        .and_then(value_to_i64)
}

fn extract_text_content(content_elements: &[Value]) -> String {
    let mut text_blocks = Vec::new();
    for element in content_elements {
        if let Some(paragraph) = element.get("paragraph") {
            text_blocks.push(extract_paragraph_text(paragraph));
        } else if let Some(table) = element.get("table") {
            text_blocks.push(extract_table_text(table));
        }
    }
    text_blocks.join("\n")
}

fn extract_paragraph_text(paragraph: &Value) -> String {
    paragraph
        .get("elements")
        .and_then(|e| e.as_array())
        .map(|elements| {
            elements
                .iter()
                .filter_map(|el| {
                    el.get("textRun")
                        .and_then(|tr| tr.get("content"))
                        .and_then(|c| c.as_str())
                })
                .collect::<String>()
        })
        .unwrap_or_default()
}

fn extract_table_text(table: &Value) -> String {
    let mut rows = Vec::new();

    if let Some(table_rows) = table.get("tableRows").and_then(|r| r.as_array()) {
        for row in table_rows {
            let mut cells = Vec::new();
            if let Some(table_cells) = row.get("tableCells").and_then(|c| c.as_array()) {
                for cell in table_cells {
                    let text = cell
                        .get("content")
                        .and_then(|v| v.as_array())
                        .map(|elements| extract_text_content(elements))
                        .unwrap_or_default();
                    cells.push(text);
                }
            }
            rows.push(cells.join(" | "));
        }
    }

    rows.join("\n")
}
