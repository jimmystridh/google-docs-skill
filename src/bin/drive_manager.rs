use google_docs_rust::auth::{
    AuthPaths, SHARED_SCOPES, TokenState, auth_required_payload, build_auth_url, ensure_token,
    load_oauth_client_config,
};
use google_docs_rust::google_api::{
    GoogleApiError, GoogleClient, detect_drive_mime_type, ensure_file_exists, map_api_error,
};
use google_docs_rust::io_helpers::{home_dir, print_json};
use serde_json::{Value, json};
use std::collections::HashMap;
use std::env;
use std::path::Path;

const EXIT_SUCCESS: i32 = 0;
const EXIT_OPERATION_FAILED: i32 = 1;
const EXIT_AUTH_ERROR: i32 = 2;
const EXIT_API_ERROR: i32 = 3;
const EXIT_INVALID_ARGS: i32 = 4;

fn main() {
    let args: Vec<String> = env::args().collect();
    let program = args
        .first()
        .map(|p| {
            std::path::Path::new(p)
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("drive_manager")
                .to_string()
        })
        .unwrap_or_else(|| "drive_manager".to_string());

    if args.len() < 2 {
        usage(&program);
        std::process::exit(EXIT_INVALID_ARGS);
    }

    let command = args[1].as_str();
    if command == "--help" || command == "-h" {
        usage(&program);
        std::process::exit(EXIT_SUCCESS);
    }

    let client = match initialize_client(&program) {
        Ok(client) => client,
        Err(exit_code) => std::process::exit(exit_code),
    };

    let options = parse_args(&args[2..]);

    let exit = match command {
        "upload" => {
            let Some(file_path) = options.get("file") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_FILE",
                    "message": "File path required: --file <path>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };

            match upload(
                &client,
                Path::new(file_path),
                options.get("folder_id").map(String::as_str),
                options.get("name").map(String::as_str),
                options.get("mime_type").map(String::as_str),
            ) {
                Ok(payload) => {
                    print_json(&payload);
                    EXIT_SUCCESS
                }
                Err(CommandError::Api(err)) => {
                    print_json(&map_api_error("upload", &err));
                    EXIT_API_ERROR
                }
                Err(CommandError::Operation {
                    error_code,
                    message,
                }) => {
                    print_json(&json!({
                        "status": "error",
                        "error_code": error_code,
                        "operation": "upload",
                        "message": message
                    }));
                    EXIT_OPERATION_FAILED
                }
            }
        }
        "download" => {
            let Some(file_id) = options.get("file_id") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_ARGS",
                    "message": "File ID and output path required: --file-id <id> --output <path>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };
            let Some(output) = options.get("output") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_ARGS",
                    "message": "File ID and output path required: --file-id <id> --output <path>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };

            match download(&client, file_id, Path::new(output)) {
                Ok(payload) => {
                    print_json(&payload);
                    EXIT_SUCCESS
                }
                Err(CommandError::Api(err)) => {
                    print_json(&map_api_error("download", &err));
                    EXIT_API_ERROR
                }
                Err(CommandError::Operation {
                    error_code,
                    message,
                }) => {
                    print_json(&json!({
                        "status": "error",
                        "error_code": error_code,
                        "operation": "download",
                        "message": message
                    }));
                    EXIT_OPERATION_FAILED
                }
            }
        }
        "list" => match list_files(
            &client,
            options.get("folder_id").map(String::as_str),
            options
                .get("max_results")
                .and_then(|v| v.parse::<i64>().ok())
                .unwrap_or(100),
            None,
        ) {
            Ok(payload) => {
                print_json(&payload);
                EXIT_SUCCESS
            }
            Err(err) => {
                print_json(&map_api_error("list", &err));
                EXIT_API_ERROR
            }
        },
        "search" => {
            let Some(query) = options.get("query") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_QUERY",
                    "message": "Search query required: --query <query>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };

            match search_files(
                &client,
                query,
                options
                    .get("max_results")
                    .and_then(|v| v.parse::<i64>().ok())
                    .unwrap_or(100),
                None,
            ) {
                Ok(payload) => {
                    print_json(&payload);
                    EXIT_SUCCESS
                }
                Err(err) => {
                    print_json(&map_api_error("search", &err));
                    EXIT_API_ERROR
                }
            }
        }
        "get-metadata" => {
            let Some(file_id) = options.get("file_id") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_FILE_ID",
                    "message": "File ID required: --file-id <id>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };

            match get_metadata(&client, file_id) {
                Ok(payload) => {
                    print_json(&payload);
                    EXIT_SUCCESS
                }
                Err(err) => {
                    print_json(&map_api_error("get_metadata", &err));
                    EXIT_API_ERROR
                }
            }
        }
        "create-folder" => {
            let Some(name) = options.get("name") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_NAME",
                    "message": "Folder name required: --name <name>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };

            match create_folder(
                &client,
                name,
                options
                    .get("parent_id")
                    .or_else(|| options.get("folder_id"))
                    .map(String::as_str),
            ) {
                Ok(payload) => {
                    print_json(&payload);
                    EXIT_SUCCESS
                }
                Err(err) => {
                    print_json(&map_api_error("create_folder", &err));
                    EXIT_API_ERROR
                }
            }
        }
        "move" => {
            let Some(file_id) = options.get("file_id") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_ARGS",
                    "message": "File ID and folder ID required: --file-id <id> --folder-id <id>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };
            let Some(folder_id) = options.get("folder_id") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_ARGS",
                    "message": "File ID and folder ID required: --file-id <id> --folder-id <id>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };

            match move_file(&client, file_id, folder_id) {
                Ok(payload) => {
                    print_json(&payload);
                    EXIT_SUCCESS
                }
                Err(err) => {
                    print_json(&map_api_error("move", &err));
                    EXIT_API_ERROR
                }
            }
        }
        "share" => {
            let Some(file_id) = options.get("file_id") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_FILE_ID",
                    "message": "File ID required: --file-id <id>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };

            match share_file(
                &client,
                file_id,
                options.get("email").map(String::as_str),
                options.get("role").map(String::as_str).unwrap_or("reader"),
                options.get("type").map(String::as_str),
            ) {
                Ok(payload) => {
                    print_json(&payload);
                    EXIT_SUCCESS
                }
                Err(err) => {
                    print_json(&map_api_error("share", &err));
                    EXIT_API_ERROR
                }
            }
        }
        "delete" => {
            let Some(file_id) = options.get("file_id") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_FILE_ID",
                    "message": "File ID required: --file-id <id>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };

            let permanent = options
                .get("permanent")
                .map(|v| v == "true")
                .unwrap_or(false);
            match delete_file(&client, file_id, permanent) {
                Ok(payload) => {
                    print_json(&payload);
                    EXIT_SUCCESS
                }
                Err(err) => {
                    print_json(&map_api_error("delete", &err));
                    EXIT_API_ERROR
                }
            }
        }
        "copy" => {
            let Some(file_id) = options.get("file_id") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_FILE_ID",
                    "message": "File ID required: --file-id <id>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };

            match copy_file(
                &client,
                file_id,
                options.get("name").map(String::as_str),
                options.get("folder_id").map(String::as_str),
            ) {
                Ok(payload) => {
                    print_json(&payload);
                    EXIT_SUCCESS
                }
                Err(err) => {
                    print_json(&map_api_error("copy", &err));
                    EXIT_API_ERROR
                }
            }
        }
        "update" => {
            let Some(file_id) = options.get("file_id") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_ARGS",
                    "message": "File ID and file path required: --file-id <id> --file <path>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };
            let Some(file_path) = options.get("file") else {
                print_json(&json!({
                    "status": "error",
                    "error_code": "MISSING_ARGS",
                    "message": "File ID and file path required: --file-id <id> --file <path>"
                }));
                std::process::exit(EXIT_INVALID_ARGS);
            };

            match update_file(
                &client,
                file_id,
                Path::new(file_path),
                options.get("name").map(String::as_str),
            ) {
                Ok(payload) => {
                    print_json(&payload);
                    EXIT_SUCCESS
                }
                Err(CommandError::Api(err)) => {
                    print_json(&map_api_error("update", &err));
                    EXIT_API_ERROR
                }
                Err(CommandError::Operation {
                    error_code,
                    message,
                }) => {
                    print_json(&json!({
                        "status": "error",
                        "error_code": error_code,
                        "operation": "update",
                        "message": message
                    }));
                    EXIT_OPERATION_FAILED
                }
            }
        }
        _ => {
            print_json(&json!({
                "status": "error",
                "error_code": "UNKNOWN_COMMAND",
                "message": format!("Unknown command: {command}"),
                "hint": format!("Run '{program} --help' for usage")
            }));
            EXIT_INVALID_ARGS
        }
    };

    std::process::exit(exit);
}

enum CommandError {
    Api(GoogleApiError),
    Operation { error_code: String, message: String },
}

fn initialize_client(_program: &str) -> std::result::Result<GoogleClient, i32> {
    let home = match home_dir() {
        Ok(home) => home,
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
        Ok(TokenState::AuthorizationRequired { auth_url }) => {
            print_json(&auth_required_payload(
                &auth_url,
                "Authorization required. Please use docs_manager auth flow.",
                "docs_manager",
            ));
            Err(EXIT_AUTH_ERROR)
        }
        Err(err) => {
            let auth_url = load_oauth_client_config(&paths.credentials_path)
                .ok()
                .and_then(|cfg| build_auth_url(&cfg, SHARED_SCOPES).ok());

            if let Some(url) = auth_url {
                print_json(&auth_required_payload(
                    &url,
                    "Authorization required. Please use docs_manager auth flow.",
                    "docs_manager",
                ));
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
        "Google Drive Manager - File Operations CLI\n\nUsage:\n  {program} <command> [options]\n\nCommands:\n  upload          Upload a file to Drive\n  download        Download a file from Drive\n  list            List files in Drive or folder\n  search          Search files with query\n  get-metadata    Get file metadata\n  create-folder   Create a new folder\n  move            Move file to folder\n  share           Share file with user or make public\n  delete          Delete file (trash or permanent)\n  copy            Copy a file\n  update          Update file content\n\nOptions:\n  --file <path>       Local file path (for upload/update)\n  --file-id <id>      Drive file ID\n  --folder-id <id>    Drive folder ID\n  --output <path>     Output file path (for download)\n  --name <name>       File/folder name\n  --query <query>     Search query (Drive query syntax)\n  --email <email>     Email address (for sharing)\n  --role <role>       Permission role: reader, writer, commenter\n  --type <type>       Permission type: user, anyone, domain\n  --max-results <n>   Max results to return (default: 100)\n  --permanent         Permanently delete (not trash)\n  --mime-type <type>  Override MIME type for upload\n\nExit Codes:\n  0 - Success\n  1 - Operation failed\n  2 - Authentication error\n  3 - API error\n  4 - Invalid arguments"
    );
}

fn parse_args(args: &[String]) -> HashMap<String, String> {
    let mut options = HashMap::new();
    let mut i = 0usize;

    while i < args.len() {
        match args[i].as_str() {
            "--file" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("file".to_string(), value.clone());
                }
                i += 2;
            }
            "--file-id" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("file_id".to_string(), value.clone());
                }
                i += 2;
            }
            "--folder-id" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("folder_id".to_string(), value.clone());
                }
                i += 2;
            }
            "--output" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("output".to_string(), value.clone());
                }
                i += 2;
            }
            "--name" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("name".to_string(), value.clone());
                }
                i += 2;
            }
            "--query" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("query".to_string(), value.clone());
                }
                i += 2;
            }
            "--email" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("email".to_string(), value.clone());
                }
                i += 2;
            }
            "--role" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("role".to_string(), value.clone());
                }
                i += 2;
            }
            "--type" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("type".to_string(), value.clone());
                }
                i += 2;
            }
            "--max-results" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("max_results".to_string(), value.clone());
                }
                i += 2;
            }
            "--permanent" => {
                options.insert("permanent".to_string(), "true".to_string());
                i += 1;
            }
            "--mime-type" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("mime_type".to_string(), value.clone());
                }
                i += 2;
            }
            "--parent-id" => {
                if let Some(value) = args.get(i + 1) {
                    options.insert("parent_id".to_string(), value.clone());
                }
                i += 2;
            }
            _ => {
                i += 1;
            }
        }
    }

    options
}

fn upload(
    client: &GoogleClient,
    file_path: &Path,
    folder_id: Option<&str>,
    name: Option<&str>,
    mime_type: Option<&str>,
) -> std::result::Result<Value, CommandError> {
    ensure_file_exists(file_path).map_err(|_| CommandError::Operation {
        error_code: "FILE_NOT_FOUND".to_string(),
        message: format!("File not found: {}", file_path.display()),
    })?;

    let file_name = name
        .map(ToString::to_string)
        .or_else(|| {
            file_path
                .file_name()
                .and_then(|n| n.to_str())
                .map(ToString::to_string)
        })
        .unwrap_or_else(|| "upload.bin".to_string());

    let detected_mime = mime_type
        .map(ToString::to_string)
        .unwrap_or_else(|| detect_drive_mime_type(file_path).to_string());

    let mut metadata = json!({
        "name": file_name
    });
    if let Some(folder_id) = folder_id {
        metadata
            .as_object_mut()
            .expect("object")
            .insert("parents".to_string(), json!([folder_id]));
    }

    let query = vec![
        ("uploadType".to_string(), "multipart".to_string()),
        (
            "fields".to_string(),
            "id,name,mimeType,webViewLink,webContentLink,parents,createdTime,modifiedTime,size"
                .to_string(),
        ),
    ];

    let result = client
        .post_multipart(
            "https://www.googleapis.com/upload/drive/v3/files",
            &query,
            &metadata,
            file_path,
            &detected_mime,
            &file_name,
        )
        .map_err(CommandError::Api)?;

    Ok(json!({
        "status": "success",
        "operation": "upload",
        "file": {
            "id": result.get("id").and_then(|v| v.as_str()),
            "name": result.get("name").and_then(|v| v.as_str()),
            "mime_type": result.get("mimeType").and_then(|v| v.as_str()),
            "web_view_link": result.get("webViewLink").and_then(|v| v.as_str()),
            "web_content_link": result.get("webContentLink").and_then(|v| v.as_str()),
            "parents": result.get("parents"),
            "created_time": result.get("createdTime").and_then(|v| v.as_str()),
            "modified_time": result.get("modifiedTime").and_then(|v| v.as_str()),
            "size": result.get("size")
        }
    }))
}

fn download(
    client: &GoogleClient,
    file_id: &str,
    output_path: &Path,
) -> std::result::Result<Value, CommandError> {
    let metadata = client
        .get_json(
            &format!("https://www.googleapis.com/drive/v3/files/{file_id}"),
            &[("fields".to_string(), "id,name,mimeType".to_string())],
        )
        .map_err(CommandError::Api)?;

    let mime_type = metadata
        .get("mimeType")
        .and_then(|v| v.as_str())
        .unwrap_or_default();

    if mime_type.starts_with("application/vnd.google-apps.") {
        return export_google_doc(client, file_id, output_path, mime_type, None);
    }

    client
        .get_bytes_to_path(
            &format!("https://www.googleapis.com/drive/v3/files/{file_id}"),
            &[("alt".to_string(), "media".to_string())],
            output_path,
        )
        .map_err(CommandError::Api)?;

    Ok(json!({
        "status": "success",
        "operation": "download",
        "file_id": file_id,
        "output_path": output_path.display().to_string(),
        "name": metadata.get("name").and_then(|v| v.as_str()),
        "mime_type": mime_type
    }))
}

fn export_google_doc(
    client: &GoogleClient,
    file_id: &str,
    output_path: &Path,
    source_mime: &str,
    export_mime: Option<&str>,
) -> std::result::Result<Value, CommandError> {
    let selected_export =
        export_mime
            .map(ToString::to_string)
            .unwrap_or_else(|| match source_mime {
                "application/vnd.google-apps.document" => "application/pdf".to_string(),
                "application/vnd.google-apps.spreadsheet" => "text/csv".to_string(),
                "application/vnd.google-apps.presentation" => "application/pdf".to_string(),
                "application/vnd.google-apps.drawing" => "image/png".to_string(),
                _ => "application/pdf".to_string(),
            });

    client
        .get_bytes_to_path(
            &format!("https://www.googleapis.com/drive/v3/files/{file_id}/export"),
            &[("mimeType".to_string(), selected_export.clone())],
            output_path,
        )
        .map_err(CommandError::Api)?;

    Ok(json!({
        "status": "success",
        "operation": "export",
        "file_id": file_id,
        "output_path": output_path.display().to_string(),
        "export_mime_type": selected_export
    }))
}

fn list_files(
    client: &GoogleClient,
    folder_id: Option<&str>,
    max_results: i64,
    page_token: Option<&str>,
) -> std::result::Result<Value, GoogleApiError> {
    let mut query_parts = vec!["trashed = false".to_string()];
    if let Some(folder_id) = folder_id {
        query_parts.push(format!("'{folder_id}' in parents"));
    }

    let mut query = vec![
        ("q".to_string(), query_parts.join(" and ")),
        ("pageSize".to_string(), max_results.to_string()),
        (
            "fields".to_string(),
            "nextPageToken,files(id,name,mimeType,webViewLink,parents,createdTime,modifiedTime,size)"
                .to_string(),
        ),
    ];
    if let Some(token) = page_token {
        query.push(("pageToken".to_string(), token.to_string()));
    }

    let result = client.get_json("https://www.googleapis.com/drive/v3/files", &query)?;

    let files = result
        .get("files")
        .and_then(|v| v.as_array())
        .cloned()
        .unwrap_or_default()
        .into_iter()
        .map(|f| {
            json!({
                "id": f.get("id").and_then(|v| v.as_str()),
                "name": f.get("name").and_then(|v| v.as_str()),
                "mime_type": f.get("mimeType").and_then(|v| v.as_str()),
                "web_view_link": f.get("webViewLink").and_then(|v| v.as_str()),
                "parents": f.get("parents"),
                "created_time": f.get("createdTime").and_then(|v| v.as_str()),
                "modified_time": f.get("modifiedTime").and_then(|v| v.as_str()),
                "size": f.get("size")
            })
        })
        .collect::<Vec<_>>();

    Ok(json!({
        "status": "success",
        "operation": "list",
        "folder_id": folder_id,
        "files": files,
        "next_page_token": result.get("nextPageToken").and_then(|v| v.as_str()),
        "count": files.len()
    }))
}

fn search_files(
    client: &GoogleClient,
    query: &str,
    max_results: i64,
    page_token: Option<&str>,
) -> std::result::Result<Value, GoogleApiError> {
    let full_query = if query.contains("trashed") {
        query.to_string()
    } else {
        format!("{query} and trashed = false")
    };

    let mut params = vec![
        ("q".to_string(), full_query),
        ("pageSize".to_string(), max_results.to_string()),
        (
            "fields".to_string(),
            "nextPageToken,files(id,name,mimeType,webViewLink,parents,createdTime,modifiedTime,size)"
                .to_string(),
        ),
    ];
    if let Some(token) = page_token {
        params.push(("pageToken".to_string(), token.to_string()));
    }

    let result = client.get_json("https://www.googleapis.com/drive/v3/files", &params)?;

    let files = result
        .get("files")
        .and_then(|v| v.as_array())
        .cloned()
        .unwrap_or_default()
        .into_iter()
        .map(|f| {
            json!({
                "id": f.get("id").and_then(|v| v.as_str()),
                "name": f.get("name").and_then(|v| v.as_str()),
                "mime_type": f.get("mimeType").and_then(|v| v.as_str()),
                "web_view_link": f.get("webViewLink").and_then(|v| v.as_str()),
                "parents": f.get("parents"),
                "created_time": f.get("createdTime").and_then(|v| v.as_str()),
                "modified_time": f.get("modifiedTime").and_then(|v| v.as_str()),
                "size": f.get("size")
            })
        })
        .collect::<Vec<_>>();

    Ok(json!({
        "status": "success",
        "operation": "search",
        "query": query,
        "files": files,
        "next_page_token": result.get("nextPageToken").and_then(|v| v.as_str()),
        "count": files.len()
    }))
}

fn get_metadata(
    client: &GoogleClient,
    file_id: &str,
) -> std::result::Result<Value, GoogleApiError> {
    let file = client.get_json(
        &format!("https://www.googleapis.com/drive/v3/files/{file_id}"),
        &[ (
            "fields".to_string(),
            "id,name,mimeType,webViewLink,webContentLink,parents,createdTime,modifiedTime,size,description,starred,trashed,owners,permissions".to_string(),
        )],
    )?;

    let owners = file
        .get("owners")
        .and_then(|v| v.as_array())
        .cloned()
        .unwrap_or_default()
        .into_iter()
        .map(|owner| {
            json!({
                "email": owner.get("emailAddress").and_then(|v| v.as_str()),
                "name": owner.get("displayName").and_then(|v| v.as_str())
            })
        })
        .collect::<Vec<_>>();

    let permissions = file
        .get("permissions")
        .and_then(|v| v.as_array())
        .cloned()
        .unwrap_or_default()
        .into_iter()
        .map(|perm| {
            json!({
                "id": perm.get("id").and_then(|v| v.as_str()),
                "type": perm.get("type").and_then(|v| v.as_str()),
                "role": perm.get("role").and_then(|v| v.as_str()),
                "email": perm.get("emailAddress").and_then(|v| v.as_str())
            })
        })
        .collect::<Vec<_>>();

    Ok(json!({
        "status": "success",
        "operation": "get_metadata",
        "file": {
            "id": file.get("id").and_then(|v| v.as_str()),
            "name": file.get("name").and_then(|v| v.as_str()),
            "mime_type": file.get("mimeType").and_then(|v| v.as_str()),
            "web_view_link": file.get("webViewLink").and_then(|v| v.as_str()),
            "web_content_link": file.get("webContentLink").and_then(|v| v.as_str()),
            "parents": file.get("parents"),
            "created_time": file.get("createdTime").and_then(|v| v.as_str()),
            "modified_time": file.get("modifiedTime").and_then(|v| v.as_str()),
            "size": file.get("size"),
            "description": file.get("description").and_then(|v| v.as_str()),
            "starred": file.get("starred"),
            "trashed": file.get("trashed"),
            "owners": owners,
            "permissions": permissions
        }
    }))
}

fn create_folder(
    client: &GoogleClient,
    name: &str,
    parent_id: Option<&str>,
) -> std::result::Result<Value, GoogleApiError> {
    let mut metadata = json!({
        "name": name,
        "mimeType": "application/vnd.google-apps.folder"
    });
    if let Some(parent_id) = parent_id {
        metadata
            .as_object_mut()
            .expect("object")
            .insert("parents".to_string(), json!([parent_id]));
    }

    let result = client.post_json(
        "https://www.googleapis.com/drive/v3/files",
        &[(
            "fields".to_string(),
            "id,name,mimeType,webViewLink,parents,createdTime".to_string(),
        )],
        &metadata,
    )?;

    Ok(json!({
        "status": "success",
        "operation": "create_folder",
        "folder": {
            "id": result.get("id").and_then(|v| v.as_str()),
            "name": result.get("name").and_then(|v| v.as_str()),
            "web_view_link": result.get("webViewLink").and_then(|v| v.as_str()),
            "parents": result.get("parents"),
            "created_time": result.get("createdTime").and_then(|v| v.as_str())
        }
    }))
}

fn move_file(
    client: &GoogleClient,
    file_id: &str,
    folder_id: &str,
) -> std::result::Result<Value, GoogleApiError> {
    let file = client.get_json(
        &format!("https://www.googleapis.com/drive/v3/files/{file_id}"),
        &[("fields".to_string(), "parents".to_string())],
    )?;

    let previous_parents = file
        .get("parents")
        .and_then(|v| v.as_array())
        .map(|p| {
            p.iter()
                .filter_map(|v| v.as_str())
                .collect::<Vec<_>>()
                .join(",")
        })
        .unwrap_or_default();

    let query = vec![
        ("addParents".to_string(), folder_id.to_string()),
        ("removeParents".to_string(), previous_parents),
        (
            "fields".to_string(),
            "id,name,parents,webViewLink".to_string(),
        ),
    ];

    let result = client.patch_json(
        &format!("https://www.googleapis.com/drive/v3/files/{file_id}"),
        &query,
        &json!({}),
    )?;

    Ok(json!({
        "status": "success",
        "operation": "move",
        "file": {
            "id": result.get("id").and_then(|v| v.as_str()),
            "name": result.get("name").and_then(|v| v.as_str()),
            "parents": result.get("parents"),
            "web_view_link": result.get("webViewLink").and_then(|v| v.as_str())
        }
    }))
}

fn share_file(
    client: &GoogleClient,
    file_id: &str,
    email: Option<&str>,
    role: &str,
    permission_type: Option<&str>,
) -> std::result::Result<Value, GoogleApiError> {
    let perm_type = permission_type.unwrap_or(if email.is_some() { "user" } else { "anyone" });

    let mut permission = json!({
        "type": perm_type,
        "role": role
    });
    if let Some(email) = email.filter(|_| perm_type == "user") {
        permission
            .as_object_mut()
            .expect("object")
            .insert("emailAddress".to_string(), Value::String(email.to_string()));
    }

    let created = client.post_json(
        &format!("https://www.googleapis.com/drive/v3/files/{file_id}/permissions"),
        &[(
            "fields".to_string(),
            "id,type,role,emailAddress".to_string(),
        )],
        &permission,
    )?;

    let file = client.get_json(
        &format!("https://www.googleapis.com/drive/v3/files/{file_id}"),
        &[(
            "fields".to_string(),
            "webViewLink,webContentLink".to_string(),
        )],
    )?;

    Ok(json!({
        "status": "success",
        "operation": "share",
        "permission": {
            "id": created.get("id").and_then(|v| v.as_str()),
            "type": created.get("type").and_then(|v| v.as_str()),
            "role": created.get("role").and_then(|v| v.as_str()),
            "email": created.get("emailAddress").and_then(|v| v.as_str())
        },
        "web_view_link": file.get("webViewLink").and_then(|v| v.as_str()),
        "web_content_link": file.get("webContentLink").and_then(|v| v.as_str())
    }))
}

fn delete_file(
    client: &GoogleClient,
    file_id: &str,
    permanent: bool,
) -> std::result::Result<Value, GoogleApiError> {
    if permanent {
        client.delete_no_content(
            &format!("https://www.googleapis.com/drive/v3/files/{file_id}"),
            &[],
        )?;
    } else {
        let _ = client.patch_json(
            &format!("https://www.googleapis.com/drive/v3/files/{file_id}"),
            &[],
            &json!({"trashed": true}),
        )?;
    }

    Ok(json!({
        "status": "success",
        "operation": "delete",
        "file_id": file_id,
        "permanent": permanent
    }))
}

fn copy_file(
    client: &GoogleClient,
    file_id: &str,
    name: Option<&str>,
    folder_id: Option<&str>,
) -> std::result::Result<Value, GoogleApiError> {
    let mut metadata = json!({});
    if let Some(name) = name {
        metadata
            .as_object_mut()
            .expect("object")
            .insert("name".to_string(), Value::String(name.to_string()));
    }
    if let Some(folder_id) = folder_id {
        metadata
            .as_object_mut()
            .expect("object")
            .insert("parents".to_string(), json!([folder_id]));
    }

    let result = client.post_json(
        &format!("https://www.googleapis.com/drive/v3/files/{file_id}/copy"),
        &[(
            "fields".to_string(),
            "id,name,mimeType,webViewLink,parents,createdTime".to_string(),
        )],
        &metadata,
    )?;

    Ok(json!({
        "status": "success",
        "operation": "copy",
        "file": {
            "id": result.get("id").and_then(|v| v.as_str()),
            "name": result.get("name").and_then(|v| v.as_str()),
            "mime_type": result.get("mimeType").and_then(|v| v.as_str()),
            "web_view_link": result.get("webViewLink").and_then(|v| v.as_str()),
            "parents": result.get("parents"),
            "created_time": result.get("createdTime").and_then(|v| v.as_str())
        }
    }))
}

fn update_file(
    client: &GoogleClient,
    file_id: &str,
    file_path: &Path,
    name: Option<&str>,
) -> std::result::Result<Value, CommandError> {
    ensure_file_exists(file_path).map_err(|_| CommandError::Operation {
        error_code: "FILE_NOT_FOUND".to_string(),
        message: format!("File not found: {}", file_path.display()),
    })?;

    let mut metadata = json!({});
    if let Some(name) = name {
        metadata
            .as_object_mut()
            .expect("object")
            .insert("name".to_string(), Value::String(name.to_string()));
    }

    let mime_type = detect_drive_mime_type(file_path).to_string();
    let file_name = file_path
        .file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("file.bin");

    let query = vec![
        ("uploadType".to_string(), "multipart".to_string()),
        (
            "fields".to_string(),
            "id,name,mimeType,webViewLink,modifiedTime,size".to_string(),
        ),
    ];

    let result = client
        .patch_multipart(
            &format!("https://www.googleapis.com/upload/drive/v3/files/{file_id}"),
            &query,
            &metadata,
            file_path,
            &mime_type,
            file_name,
        )
        .map_err(CommandError::Api)?;

    Ok(json!({
        "status": "success",
        "operation": "update",
        "file": {
            "id": result.get("id").and_then(|v| v.as_str()),
            "name": result.get("name").and_then(|v| v.as_str()),
            "mime_type": result.get("mimeType").and_then(|v| v.as_str()),
            "web_view_link": result.get("webViewLink").and_then(|v| v.as_str()),
            "modified_time": result.get("modifiedTime").and_then(|v| v.as_str()),
            "size": result.get("size")
        }
    }))
}
