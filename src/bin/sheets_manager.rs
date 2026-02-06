use anyhow::{Context, Result};
use google_docs_rust::auth::{
    AuthPaths, SHARED_SCOPES, TokenState, auth_required_payload, build_auth_url,
    complete_authorization, ensure_token, load_oauth_client_config, load_stored_token,
    save_stored_token,
};
use google_docs_rust::google_api::{GoogleApiError, GoogleClient, map_api_error};
use google_docs_rust::io_helpers::{home_dir, print_json, read_stdin_json};
use serde_json::{Map, Value, json};
use std::env;

const EXIT_SUCCESS: i32 = 0;
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
                .unwrap_or("sheets_manager")
                .to_string()
        })
        .unwrap_or_else(|| "sheets_manager".to_string());

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

        if let Err(err) = complete_auth(&args[2]) {
            print_json(&json!({
                "status": "error",
                "error_code": "AUTH_FAILED",
                "message": format!("Authorization failed: {err}")
            }));
            std::process::exit(EXIT_AUTH_ERROR);
        }

        std::process::exit(EXIT_SUCCESS);
    }

    let client = match initialize_client(&program) {
        Ok(client) => client,
        Err(code) => std::process::exit(code),
    };

    let exit = match command {
        "create" => dispatch_json_command("create", || {
            let input = read_stdin_json()?;
            let title = required_string(&input, "title")?;
            let sheets = input.get("sheets").and_then(|v| v.as_array()).cloned();
            let data = input.get("data").and_then(|v| v.as_array()).cloned();
            create_spreadsheet(&client, &title, sheets, data)
        }),
        "read" => dispatch_json_command("read", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let range = required_string(&input, "range")?;
            read_range(&client, &spreadsheet_id, &range)
        }),
        "write" => dispatch_json_command("write", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let range = required_string(&input, "range")?;
            let values = required_array(&input, "values")?;
            write_range(&client, &spreadsheet_id, &range, values)
        }),
        "append" => dispatch_json_command("append", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let range = required_string(&input, "range")?;
            let values = required_array(&input, "values")?;
            append_rows(&client, &spreadsheet_id, &range, values)
        }),
        "clear" => dispatch_json_command("clear", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let range = required_string(&input, "range")?;
            clear_range(&client, &spreadsheet_id, &range)
        }),
        "batch-read" => dispatch_json_command("batch-read", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let ranges = required_array(&input, "ranges")?
                .iter()
                .filter_map(|v| v.as_str().map(ToString::to_string))
                .collect::<Vec<_>>();
            if ranges.is_empty() {
                anyhow::bail!("Required fields: spreadsheet_id, ranges");
            }
            batch_read(&client, &spreadsheet_id, &ranges)
        }),
        "batch-write" => dispatch_json_command("batch-write", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let data = required_array(&input, "data")?;
            batch_write(&client, &spreadsheet_id, data)
        }),
        "get-metadata" => dispatch_json_command("get-metadata", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            get_metadata(&client, &spreadsheet_id)
        }),
        "add-sheet" => dispatch_json_command("add-sheet", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let title = required_string(&input, "title")?;
            add_sheet(&client, &spreadsheet_id, &title)
        }),
        "delete-sheet" => dispatch_json_command("delete-sheet", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            delete_sheet(&client, &spreadsheet_id, sheet_id)
        }),
        "rename-sheet" => dispatch_json_command("rename-sheet", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let title = required_string(&input, "title")?;
            rename_sheet(&client, &spreadsheet_id, sheet_id, &title)
        }),
        "copy-sheet" => dispatch_json_command("copy-sheet", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let destination = input
                .get("destination_spreadsheet_id")
                .and_then(|v| v.as_str())
                .map(ToString::to_string);
            copy_sheet(&client, &spreadsheet_id, sheet_id, destination)
        }),
        "format" => dispatch_json_command("format", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let range = required_string(&input, "range")?;

            let mut options = Map::new();
            for key in [
                "bold",
                "italic",
                "underline",
                "font_size",
                "font_family",
                "foreground_color",
                "background_color",
                "horizontal_alignment",
                "vertical_alignment",
                "number_format",
                "wrap_strategy",
                "text_rotation",
                "borders",
            ] {
                if let Some(value) = input.get(key) {
                    options.insert(key.to_string(), value.clone());
                }
            }

            format_cells(&client, &spreadsheet_id, sheet_id, &range, &options)
        }),
        "merge-cells" => dispatch_json_command("merge-cells", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let range = required_string(&input, "range")?;
            let merge_type = input
                .get("merge_type")
                .and_then(|v| v.as_str())
                .unwrap_or("MERGE_ALL")
                .to_string();
            merge_cells(&client, &spreadsheet_id, sheet_id, &range, &merge_type)
        }),
        "unmerge-cells" => dispatch_json_command("unmerge-cells", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let range = required_string(&input, "range")?;
            unmerge_cells(&client, &spreadsheet_id, sheet_id, &range)
        }),
        "freeze" => dispatch_json_command("freeze", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let rows = input.get("rows").and_then(value_to_i64);
            let cols = input.get("cols").and_then(value_to_i64);
            freeze(&client, &spreadsheet_id, sheet_id, rows, cols)
        }),
        "auto-resize" => dispatch_json_command("auto-resize", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let start_col = required_i64(&input, "start_col")?;
            let end_col = required_i64(&input, "end_col")?;
            auto_resize(&client, &spreadsheet_id, sheet_id, start_col, end_col)
        }),
        "sort" => dispatch_json_command("sort", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let range = required_string(&input, "range")?;
            let sort_column = required_i64(&input, "sort_column")?;
            let ascending = input
                .get("ascending")
                .and_then(|v| v.as_bool())
                .unwrap_or(true);
            sort_range(
                &client,
                &spreadsheet_id,
                sheet_id,
                &range,
                sort_column,
                ascending,
            )
        }),
        "find-replace" => dispatch_json_command("find-replace", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let find = required_string(&input, "find")?;
            let replace = required_string(&input, "replace")?;
            let sheet_id = input.get("sheet_id").and_then(value_to_i64);
            let match_case = input
                .get("match_case")
                .and_then(|v| v.as_bool())
                .unwrap_or(false);
            let match_entire_cell = input
                .get("match_entire_cell")
                .and_then(|v| v.as_bool())
                .unwrap_or(false);

            find_replace(
                &client,
                &spreadsheet_id,
                &find,
                &replace,
                sheet_id,
                match_case,
                match_entire_cell,
            )
        }),
        "set-column-width" => dispatch_json_command("set-column-width", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let start_col = required_i64(&input, "start_col")?;
            let end_col = required_i64(&input, "end_col")?;
            let width = required_i64(&input, "width")?;
            set_column_width(
                &client,
                &spreadsheet_id,
                sheet_id,
                start_col,
                end_col,
                width,
            )
        }),
        "set-row-height" => dispatch_json_command("set-row-height", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let start_row = required_i64(&input, "start_row")?;
            let end_row = required_i64(&input, "end_row")?;
            let height = required_i64(&input, "height")?;
            set_row_height(
                &client,
                &spreadsheet_id,
                sheet_id,
                start_row,
                end_row,
                height,
            )
        }),
        "add-filter" => dispatch_json_command("add-filter", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let range = required_string(&input, "range")?;
            add_filter(&client, &spreadsheet_id, sheet_id, &range)
        }),
        "add-chart" => dispatch_json_command("add-chart", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let range = required_string(&input, "range")?;
            let chart_type = required_string(&input, "chart_type")?;
            let title = required_string(&input, "title")?;
            add_chart(
                &client,
                &spreadsheet_id,
                sheet_id,
                &range,
                &chart_type,
                &title,
            )
        }),
        "protect-range" => dispatch_json_command("protect-range", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let range = required_string(&input, "range")?;
            let description = input
                .get("description")
                .and_then(|v| v.as_str())
                .map(ToString::to_string);
            let editors = input.get("editors").and_then(|v| v.as_array()).map(|arr| {
                arr.iter()
                    .filter_map(|v| v.as_str().map(ToString::to_string))
                    .collect::<Vec<_>>()
            });
            protect_range(
                &client,
                &spreadsheet_id,
                sheet_id,
                &range,
                description,
                editors,
            )
        }),
        "add-conditional-format" => dispatch_json_command("add-conditional-format", || {
            let input = read_stdin_json()?;
            let spreadsheet_id = required_string(&input, "spreadsheet_id")?;
            let sheet_id = required_i64(&input, "sheet_id")?;
            let range = required_string(&input, "range")?;
            let rule_type = required_string(&input, "rule_type")?;

            let mut rule_params = Map::new();
            if let Some(obj) = input.as_object() {
                for (k, v) in obj {
                    if ["spreadsheet_id", "sheet_id", "range", "rule_type"].contains(&k.as_str()) {
                        continue;
                    }
                    rule_params.insert(k.clone(), v.clone());
                }
            }

            add_conditional_format(
                &client,
                &spreadsheet_id,
                sheet_id,
                &range,
                &rule_type,
                &rule_params,
            )
        }),
        _ => {
            print_json(&json!({
                "status": "error",
                "error_code": "INVALID_COMMAND",
                "message": format!("Unknown command: {command}"),
                "valid_commands": [
                    "auth",
                    "create",
                    "read",
                    "write",
                    "append",
                    "clear",
                    "batch-read",
                    "batch-write",
                    "get-metadata",
                    "add-sheet",
                    "delete-sheet",
                    "rename-sheet",
                    "copy-sheet",
                    "format",
                    "merge-cells",
                    "unmerge-cells",
                    "freeze",
                    "auto-resize",
                    "sort",
                    "find-replace",
                    "set-column-width",
                    "set-row-height",
                    "add-filter",
                    "add-chart",
                    "protect-range",
                    "add-conditional-format"
                ]
            }));
            usage(&program);
            EXIT_INVALID_ARGS
        }
    };

    std::process::exit(exit);
}

fn complete_auth(code: &str) -> Result<()> {
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

    Ok(())
}

fn initialize_client(program: &str) -> std::result::Result<GoogleClient, i32> {
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
                "Authorization required. Please visit the URL and enter the code.",
                program,
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
                    "Authorization required. Please visit the URL and enter the code.",
                    program,
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
        "Google Sheets Manager - Spreadsheet Operations CLI\n\nUsage:\n  {program} <command> [options]\n\nAll commands accept JSON via stdin (except auth).\n\nCommands:\n  auth <code>              Complete OAuth authorization with code\n  create                   Create new spreadsheet\n  read                     Read cell range\n  write                    Write values to range\n  append                   Append rows after existing data\n  clear                    Clear cell range\n  batch-read               Read multiple ranges\n  batch-write              Write to multiple ranges\n  get-metadata             Get spreadsheet info\n  add-sheet                Add new sheet/tab\n  delete-sheet             Delete sheet/tab\n  rename-sheet             Rename sheet/tab\n  copy-sheet               Copy sheet to same or other spreadsheet\n  format                   Format cells\n  merge-cells              Merge cell range\n  unmerge-cells            Unmerge cell range\n  freeze                   Freeze rows/columns\n  auto-resize              Auto-resize columns to fit content\n  sort                     Sort range by column\n  find-replace             Find and replace text\n  set-column-width         Set column width in pixels\n  set-row-height           Set row height in pixels\n  add-filter               Add basic filter to range\n  add-chart                Add chart from data range\n  protect-range            Protect cells from editing\n  add-conditional-format   Add conditional formatting rule\n\nExit Codes:\n  0 - Success\n  1 - Operation failed\n  2 - Authentication error\n  3 - API error\n  4 - Invalid arguments"
    );
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
                print_json(&map_api_error(operation, api_err));
                return EXIT_API_ERROR;
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
        .ok_or_else(|| anyhow::anyhow!(format!("Required fields: {key}")))
}

fn required_array<'a>(input: &'a Value, key: &str) -> Result<&'a Vec<Value>> {
    input
        .get(key)
        .and_then(|v| v.as_array())
        .ok_or_else(|| anyhow::anyhow!(format!("Required fields: {key}")))
}

fn required_i64(input: &Value, key: &str) -> Result<i64> {
    input
        .get(key)
        .and_then(value_to_i64)
        .ok_or_else(|| anyhow::anyhow!(format!("Required fields: {key}")))
}

fn value_to_i64(value: &Value) -> Option<i64> {
    if let Some(v) = value.as_i64() {
        Some(v)
    } else if let Some(v) = value.as_u64() {
        i64::try_from(v).ok()
    } else {
        value.as_f64().map(|v| v as i64)
    }
}

fn encode_range(range: &str) -> String {
    urlencoding::encode(range).to_string()
}

fn create_spreadsheet(
    client: &GoogleClient,
    title: &str,
    sheets: Option<Vec<Value>>,
    data: Option<Vec<Value>>,
) -> Result<Value> {
    let mut spreadsheet = json!({
        "properties": { "title": title }
    });

    if let Some(sheet_names) = sheets {
        let configured = sheet_names
            .iter()
            .enumerate()
            .filter_map(|(i, value)| {
                value.as_str().map(|name| {
                    json!({
                        "properties": {
                            "title": name,
                            "index": i
                        }
                    })
                })
            })
            .collect::<Vec<_>>();

        if !configured.is_empty() {
            spreadsheet
                .as_object_mut()
                .expect("object")
                .insert("sheets".to_string(), Value::Array(configured));
        }
    }

    let result = client
        .post_json(
            "https://sheets.googleapis.com/v4/spreadsheets",
            &[],
            &spreadsheet,
        )
        .map_err(anyhow::Error::from)?;

    let spreadsheet_id = result
        .get("spreadsheetId")
        .and_then(|v| v.as_str())
        .context("Missing spreadsheetId in create response")?
        .to_string();

    if let Some(values) = data
        && !values.is_empty()
    {
        let first_sheet = result
            .get("sheets")
            .and_then(|v| v.as_array())
            .and_then(|arr| arr.first())
            .and_then(|sheet| sheet.get("properties"))
            .and_then(|props| props.get("title"))
            .and_then(|title| title.as_str())
            .unwrap_or("Sheet1")
            .to_string();

        let range = format!("{first_sheet}!A1");
        let payload = json!({
            "range": range,
            "values": values
        });

        let _ = client
            .put_json(
                &format!(
                    "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}",
                    spreadsheet_id,
                    encode_range(&range)
                ),
                &[("valueInputOption".to_string(), "USER_ENTERED".to_string())],
                &payload,
            )
            .map_err(anyhow::Error::from)?;
    }

    let sheets_out = result
        .get("sheets")
        .and_then(|v| v.as_array())
        .cloned()
        .unwrap_or_default()
        .into_iter()
        .map(|s| {
            json!({
                "title": s.get("properties").and_then(|p| p.get("title")).and_then(|v| v.as_str()),
                "sheet_id": s.get("properties").and_then(|p| p.get("sheetId"))
            })
        })
        .collect::<Vec<_>>();

    Ok(json!({
        "status": "success",
        "operation": "create",
        "spreadsheet_id": spreadsheet_id,
        "title": result.get("properties").and_then(|p| p.get("title")).and_then(|v| v.as_str()),
        "spreadsheet_url": result.get("spreadsheetUrl").and_then(|v| v.as_str()),
        "sheets": sheets_out
    }))
}

fn read_range(client: &GoogleClient, spreadsheet_id: &str, range: &str) -> Result<Value> {
    let result = client
        .get_json(
            &format!(
                "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}",
                spreadsheet_id,
                encode_range(range)
            ),
            &[],
        )
        .map_err(anyhow::Error::from)?;

    let values = result
        .get("values")
        .and_then(|v| v.as_array())
        .cloned()
        .unwrap_or_default();

    Ok(json!({
        "status": "success",
        "operation": "read",
        "spreadsheet_id": spreadsheet_id,
        "range": result.get("range").and_then(|v| v.as_str()),
        "values": values,
        "rows": values.len(),
        "columns": values
            .first()
            .and_then(|row| row.as_array())
            .map(|r| r.len())
            .unwrap_or(0)
    }))
}

fn write_range(
    client: &GoogleClient,
    spreadsheet_id: &str,
    range: &str,
    values: &Vec<Value>,
) -> Result<Value> {
    let payload = json!({
        "range": range,
        "values": values
    });

    let result = client
        .put_json(
            &format!(
                "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}",
                spreadsheet_id,
                encode_range(range)
            ),
            &[("valueInputOption".to_string(), "USER_ENTERED".to_string())],
            &payload,
        )
        .map_err(anyhow::Error::from)?;

    Ok(json!({
        "status": "success",
        "operation": "write",
        "spreadsheet_id": spreadsheet_id,
        "updated_range": result.get("updatedRange").and_then(|v| v.as_str()),
        "updated_rows": result.get("updatedRows"),
        "updated_columns": result.get("updatedColumns"),
        "updated_cells": result.get("updatedCells")
    }))
}

fn append_rows(
    client: &GoogleClient,
    spreadsheet_id: &str,
    range: &str,
    values: &Vec<Value>,
) -> Result<Value> {
    let payload = json!({
        "range": range,
        "values": values
    });

    let result = client
        .post_json(
            &format!(
                "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}:append",
                spreadsheet_id,
                encode_range(range)
            ),
            &[
                ("valueInputOption".to_string(), "USER_ENTERED".to_string()),
                ("insertDataOption".to_string(), "INSERT_ROWS".to_string()),
            ],
            &payload,
        )
        .map_err(anyhow::Error::from)?;

    let updates = result.get("updates").cloned().unwrap_or(Value::Null);

    Ok(json!({
        "status": "success",
        "operation": "append",
        "spreadsheet_id": spreadsheet_id,
        "updated_range": updates.get("updatedRange").and_then(|v| v.as_str()),
        "updated_rows": updates.get("updatedRows"),
        "updated_columns": updates.get("updatedColumns"),
        "updated_cells": updates.get("updatedCells")
    }))
}

fn clear_range(client: &GoogleClient, spreadsheet_id: &str, range: &str) -> Result<Value> {
    let _ = client
        .post_json(
            &format!(
                "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}:clear",
                spreadsheet_id,
                encode_range(range)
            ),
            &[],
            &json!({}),
        )
        .map_err(anyhow::Error::from)?;

    Ok(json!({
        "status": "success",
        "operation": "clear",
        "spreadsheet_id": spreadsheet_id,
        "cleared_range": range
    }))
}

fn batch_read(client: &GoogleClient, spreadsheet_id: &str, ranges: &[String]) -> Result<Value> {
    let mut query = vec![];
    for range in ranges {
        query.push(("ranges".to_string(), range.clone()));
    }

    let result = client
        .get_json(
            &format!(
                "https://sheets.googleapis.com/v4/spreadsheets/{}/values:batchGet",
                spreadsheet_id
            ),
            &query,
        )
        .map_err(anyhow::Error::from)?;

    let range_data = result
        .get("valueRanges")
        .and_then(|v| v.as_array())
        .cloned()
        .unwrap_or_default()
        .into_iter()
        .map(|vr| {
            let values = vr
                .get("values")
                .and_then(|v| v.as_array())
                .cloned()
                .unwrap_or_default();
            json!({
                "range": vr.get("range").and_then(|v| v.as_str()),
                "values": values,
                "rows": values.len(),
                "columns": values.first().and_then(|row| row.as_array()).map(|r| r.len()).unwrap_or(0)
            })
        })
        .collect::<Vec<_>>();

    Ok(json!({
        "status": "success",
        "operation": "batch-read",
        "spreadsheet_id": spreadsheet_id,
        "ranges": range_data
    }))
}

fn batch_write(client: &GoogleClient, spreadsheet_id: &str, data: &[Value]) -> Result<Value> {
    let value_ranges = data
        .iter()
        .map(|entry| {
            json!({
                "range": entry.get("range"),
                "values": entry.get("values")
            })
        })
        .collect::<Vec<_>>();

    let payload = json!({
        "valueInputOption": "USER_ENTERED",
        "data": value_ranges
    });

    let result = client
        .post_json(
            &format!(
                "https://sheets.googleapis.com/v4/spreadsheets/{}/values:batchUpdate",
                spreadsheet_id
            ),
            &[],
            &payload,
        )
        .map_err(anyhow::Error::from)?;

    Ok(json!({
        "status": "success",
        "operation": "batch-write",
        "spreadsheet_id": spreadsheet_id,
        "total_updated_rows": result.get("totalUpdatedRows"),
        "total_updated_columns": result.get("totalUpdatedColumns"),
        "total_updated_cells": result.get("totalUpdatedCells"),
        "total_updated_sheets": result.get("totalUpdatedSheets")
    }))
}

fn get_metadata(client: &GoogleClient, spreadsheet_id: &str) -> Result<Value> {
    let result = client
        .get_json(
            &format!("https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}"),
            &[],
        )
        .map_err(anyhow::Error::from)?;

    let sheets_info = result
        .get("sheets")
        .and_then(|v| v.as_array())
        .cloned()
        .unwrap_or_default()
        .into_iter()
        .map(|sheet| {
            let props = sheet.get("properties").cloned().unwrap_or(Value::Null);
            json!({
                "title": props.get("title").and_then(|v| v.as_str()),
                "sheet_id": props.get("sheetId"),
                "index": props.get("index"),
                "sheet_type": props.get("sheetType").and_then(|v| v.as_str()),
                "row_count": props.get("gridProperties").and_then(|g| g.get("rowCount")),
                "column_count": props.get("gridProperties").and_then(|g| g.get("columnCount")),
                "frozen_row_count": props.get("gridProperties").and_then(|g| g.get("frozenRowCount")),
                "frozen_column_count": props.get("gridProperties").and_then(|g| g.get("frozenColumnCount"))
            })
        })
        .collect::<Vec<_>>();

    Ok(json!({
        "status": "success",
        "operation": "get-metadata",
        "spreadsheet_id": result.get("spreadsheetId").and_then(|v| v.as_str()),
        "title": result.get("properties").and_then(|p| p.get("title")).and_then(|v| v.as_str()),
        "locale": result.get("properties").and_then(|p| p.get("locale")).and_then(|v| v.as_str()),
        "time_zone": result.get("properties").and_then(|p| p.get("timeZone")).and_then(|v| v.as_str()),
        "spreadsheet_url": result.get("spreadsheetUrl").and_then(|v| v.as_str()),
        "sheets": sheets_info
    }))
}

fn add_sheet(client: &GoogleClient, spreadsheet_id: &str, title: &str) -> Result<Value> {
    let requests = vec![json!({
        "addSheet": {
            "properties": {"title": title}
        }
    })];

    let result = batch_update_spreadsheet(client, spreadsheet_id, requests)?;
    let new_sheet = result
        .get("replies")
        .and_then(|v| v.as_array())
        .and_then(|arr| arr.first())
        .and_then(|reply| reply.get("addSheet"))
        .and_then(|add| add.get("properties"))
        .cloned()
        .unwrap_or(Value::Null);

    Ok(json!({
        "status": "success",
        "operation": "add-sheet",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": new_sheet.get("sheetId"),
        "title": new_sheet.get("title").and_then(|v| v.as_str()),
        "index": new_sheet.get("index")
    }))
}

fn delete_sheet(client: &GoogleClient, spreadsheet_id: &str, sheet_id: i64) -> Result<Value> {
    let requests = vec![json!({
        "deleteSheet": {"sheetId": sheet_id}
    })];
    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "delete-sheet",
        "spreadsheet_id": spreadsheet_id,
        "deleted_sheet_id": sheet_id
    }))
}

fn rename_sheet(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    title: &str,
) -> Result<Value> {
    let requests = vec![json!({
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "title": title},
            "fields": "title"
        }
    })];

    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "rename-sheet",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "new_title": title
    }))
}

fn copy_sheet(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    destination_spreadsheet_id: Option<String>,
) -> Result<Value> {
    let destination = destination_spreadsheet_id
        .clone()
        .unwrap_or_else(|| spreadsheet_id.to_string());

    let payload = json!({
        "destinationSpreadsheetId": destination
    });

    let result = client
        .post_json(
            &format!(
                "https://sheets.googleapis.com/v4/spreadsheets/{}/sheets/{}:copyTo",
                spreadsheet_id, sheet_id
            ),
            &[],
            &payload,
        )
        .map_err(anyhow::Error::from)?;

    Ok(json!({
        "status": "success",
        "operation": "copy-sheet",
        "spreadsheet_id": spreadsheet_id,
        "source_sheet_id": sheet_id,
        "destination_spreadsheet_id": destination,
        "new_sheet_id": result.get("sheetId"),
        "new_title": result.get("title").and_then(|v| v.as_str())
    }))
}

fn format_cells(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    range: &str,
    format_options: &Map<String, Value>,
) -> Result<Value> {
    let grid_range = parse_a1_to_grid_range(range, sheet_id);
    let cell_format = build_cell_format(format_options);
    let fields = build_format_fields(format_options);

    let requests = vec![json!({
        "repeatCell": {
            "range": grid_range,
            "cell": {"userEnteredFormat": cell_format},
            "fields": format!("userEnteredFormat({fields})")
        }
    })];

    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "format",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "range": range,
        "format_applied": format_options
    }))
}

fn merge_cells(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    range: &str,
    merge_type: &str,
) -> Result<Value> {
    let grid_range = parse_a1_to_grid_range(range, sheet_id);
    let requests = vec![json!({
        "mergeCells": {
            "range": grid_range,
            "mergeType": merge_type
        }
    })];

    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "merge-cells",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "range": range,
        "merge_type": merge_type
    }))
}

fn unmerge_cells(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    range: &str,
) -> Result<Value> {
    let grid_range = parse_a1_to_grid_range(range, sheet_id);
    let requests = vec![json!({
        "unmergeCells": {
            "range": grid_range
        }
    })];

    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "unmerge-cells",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "range": range
    }))
}

fn freeze(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    rows: Option<i64>,
    cols: Option<i64>,
) -> Result<Value> {
    let mut grid_properties = Map::new();
    let mut fields = Vec::new();

    if let Some(rows) = rows {
        grid_properties.insert("frozenRowCount".to_string(), Value::Number(rows.into()));
        fields.push("gridProperties.frozenRowCount");
    }

    if let Some(cols) = cols {
        grid_properties.insert("frozenColumnCount".to_string(), Value::Number(cols.into()));
        fields.push("gridProperties.frozenColumnCount");
    }

    let requests = vec![json!({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": grid_properties
            },
            "fields": fields.join(",")
        }
    })];

    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "freeze",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "frozen_rows": rows,
        "frozen_cols": cols
    }))
}

fn auto_resize(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    start_col: i64,
    end_col: i64,
) -> Result<Value> {
    let requests = vec![json!({
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": start_col,
                "endIndex": end_col
            }
        }
    })];

    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "auto-resize",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "start_col": start_col,
        "end_col": end_col
    }))
}

fn sort_range(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    range: &str,
    sort_column: i64,
    ascending: bool,
) -> Result<Value> {
    let grid_range = parse_a1_to_grid_range(range, sheet_id);

    let requests = vec![json!({
        "sortRange": {
            "range": grid_range,
            "sortSpecs": [{
                "dimensionIndex": sort_column,
                "sortOrder": if ascending {"ASCENDING"} else {"DESCENDING"}
            }]
        }
    })];

    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "sort",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "range": range,
        "sort_column": sort_column,
        "ascending": ascending
    }))
}

fn find_replace(
    client: &GoogleClient,
    spreadsheet_id: &str,
    find: &str,
    replace: &str,
    sheet_id: Option<i64>,
    match_case: bool,
    match_entire_cell: bool,
) -> Result<Value> {
    let mut request = json!({
        "find": find,
        "replacement": replace,
        "matchCase": match_case,
        "matchEntireCell": match_entire_cell,
        "searchByRegex": false,
        "includeFormulas": false
    });

    if let Some(sheet_id) = sheet_id {
        request
            .as_object_mut()
            .expect("object")
            .insert("sheetId".to_string(), Value::Number(sheet_id.into()));
    }

    let requests = vec![json!({
        "findReplace": request
    })];

    let result = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    let fr = result
        .get("replies")
        .and_then(|v| v.as_array())
        .and_then(|arr| arr.first())
        .and_then(|reply| reply.get("findReplace"))
        .cloned()
        .unwrap_or(Value::Null);

    Ok(json!({
        "status": "success",
        "operation": "find-replace",
        "spreadsheet_id": spreadsheet_id,
        "find": find,
        "replace": replace,
        "occurrences_changed": fr.get("occurrencesChanged").and_then(value_to_i64).unwrap_or(0),
        "values_changed": fr.get("valuesChanged").and_then(value_to_i64).unwrap_or(0),
        "sheets_changed": fr.get("sheetsChanged").and_then(value_to_i64).unwrap_or(0),
        "formulas_changed": fr.get("formulasChanged").and_then(value_to_i64).unwrap_or(0)
    }))
}

fn set_column_width(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    start_col: i64,
    end_col: i64,
    width: i64,
) -> Result<Value> {
    let requests = vec![json!({
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": start_col,
                "endIndex": end_col
            },
            "properties": {"pixelSize": width},
            "fields": "pixelSize"
        }
    })];

    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "set-column-width",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "start_col": start_col,
        "end_col": end_col,
        "width": width
    }))
}

fn set_row_height(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    start_row: i64,
    end_row: i64,
    height: i64,
) -> Result<Value> {
    let requests = vec![json!({
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": start_row,
                "endIndex": end_row
            },
            "properties": {"pixelSize": height},
            "fields": "pixelSize"
        }
    })];

    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "set-row-height",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "start_row": start_row,
        "end_row": end_row,
        "height": height
    }))
}

fn add_filter(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    range: &str,
) -> Result<Value> {
    let grid_range = parse_a1_to_grid_range(range, sheet_id);
    let requests = vec![json!({
        "setBasicFilter": {
            "filter": {"range": grid_range}
        }
    })];

    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "add-filter",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "range": range
    }))
}

fn add_chart(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    range: &str,
    chart_type: &str,
    title: &str,
) -> Result<Value> {
    let grid_range = parse_a1_to_grid_range(range, sheet_id);
    let chart_spec = build_chart_spec(chart_type, title, &grid_range);

    let anchor_col = grid_range
        .get("endColumnIndex")
        .and_then(value_to_i64)
        .unwrap_or(0)
        + 1;

    let requests = vec![json!({
        "addChart": {
            "chart": {
                "spec": chart_spec,
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": sheet_id,
                            "rowIndex": 0,
                            "columnIndex": anchor_col
                        }
                    }
                }
            }
        }
    })];

    let result = batch_update_spreadsheet(client, spreadsheet_id, requests)?;
    let chart = result
        .get("replies")
        .and_then(|v| v.as_array())
        .and_then(|arr| arr.first())
        .and_then(|reply| reply.get("addChart"))
        .and_then(|item| item.get("chart"))
        .cloned()
        .unwrap_or(Value::Null);

    Ok(json!({
        "status": "success",
        "operation": "add-chart",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "chart_id": chart.get("chartId"),
        "title": title,
        "chart_type": chart_type
    }))
}

fn protect_range(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    range: &str,
    description: Option<String>,
    editors: Option<Vec<String>>,
) -> Result<Value> {
    let grid_range = parse_a1_to_grid_range(range, sheet_id);

    let mut protected_range = json!({
        "range": grid_range,
        "warningOnly": false
    });

    if let Some(description) = description {
        protected_range
            .as_object_mut()
            .expect("object")
            .insert("description".to_string(), Value::String(description));
    }

    if let Some(editors) = editors
        && !editors.is_empty()
    {
        protected_range
            .as_object_mut()
            .expect("object")
            .insert("editors".to_string(), json!({"users": editors}));
    }

    let requests = vec![json!({
        "addProtectedRange": {
            "protectedRange": protected_range
        }
    })];

    let result = batch_update_spreadsheet(client, spreadsheet_id, requests)?;
    let protected = result
        .get("replies")
        .and_then(|v| v.as_array())
        .and_then(|arr| arr.first())
        .and_then(|reply| reply.get("addProtectedRange"))
        .and_then(|entry| entry.get("protectedRange"))
        .cloned()
        .unwrap_or(Value::Null);

    Ok(json!({
        "status": "success",
        "operation": "protect-range",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "range": range,
        "protected_range_id": protected.get("protectedRangeId"),
        "description": protected.get("description").and_then(|v| v.as_str())
    }))
}

fn add_conditional_format(
    client: &GoogleClient,
    spreadsheet_id: &str,
    sheet_id: i64,
    range: &str,
    rule_type: &str,
    rule_params: &Map<String, Value>,
) -> Result<Value> {
    let grid_range = parse_a1_to_grid_range(range, sheet_id);
    let rule = build_conditional_format_rule(rule_type, &grid_range, rule_params);

    let requests = vec![json!({
        "addConditionalFormatRule": {
            "rule": rule,
            "index": 0
        }
    })];

    let _ = batch_update_spreadsheet(client, spreadsheet_id, requests)?;

    Ok(json!({
        "status": "success",
        "operation": "add-conditional-format",
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": sheet_id,
        "range": range,
        "rule_type": rule_type
    }))
}

fn batch_update_spreadsheet(
    client: &GoogleClient,
    spreadsheet_id: &str,
    requests: Vec<Value>,
) -> Result<Value> {
    let payload = json!({"requests": requests});
    client
        .post_json(
            &format!("https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}:batchUpdate"),
            &[],
            &payload,
        )
        .map_err(anyhow::Error::from)
}

fn parse_a1_to_grid_range(range: &str, sheet_id: i64) -> Value {
    let cell_range = if let Some((_, suffix)) = range.split_once('!') {
        suffix
    } else {
        range
    };

    let mut grid = Map::new();
    grid.insert("sheetId".to_string(), Value::Number(sheet_id.into()));

    if let Some((start_ref, end_ref)) = cell_range.split_once(':') {
        let (start_col, start_row) = parse_cell_ref(start_ref);
        let (end_col, end_row) = parse_cell_ref(end_ref);

        if let Some(col) = start_col {
            grid.insert("startColumnIndex".to_string(), Value::Number(col.into()));
        }
        if let Some(row) = start_row {
            grid.insert("startRowIndex".to_string(), Value::Number(row.into()));
        }
        if let Some(col) = end_col {
            grid.insert(
                "endColumnIndex".to_string(),
                Value::Number((col + 1).into()),
            );
        }
        if let Some(row) = end_row {
            grid.insert("endRowIndex".to_string(), Value::Number((row + 1).into()));
        }
    } else {
        let (col, row) = parse_cell_ref(cell_range);
        if let Some(col) = col {
            grid.insert("startColumnIndex".to_string(), Value::Number(col.into()));
            grid.insert(
                "endColumnIndex".to_string(),
                Value::Number((col + 1).into()),
            );
        }
        if let Some(row) = row {
            grid.insert("startRowIndex".to_string(), Value::Number(row.into()));
            grid.insert("endRowIndex".to_string(), Value::Number((row + 1).into()));
        }
    }

    Value::Object(grid)
}

fn parse_cell_ref(reference: &str) -> (Option<i64>, Option<i64>) {
    let mut letters = String::new();
    let mut numbers = String::new();

    for ch in reference.chars() {
        if ch.is_ascii_alphabetic() && numbers.is_empty() {
            letters.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            numbers.push(ch);
        }
    }

    let col = if letters.is_empty() {
        None
    } else {
        Some(col_letters_to_index(&letters))
    };

    let row = if numbers.is_empty() {
        None
    } else {
        numbers.parse::<i64>().ok().map(|r| r - 1)
    };

    (col, row)
}

fn col_letters_to_index(letters: &str) -> i64 {
    let mut result = 0i64;
    for c in letters.chars() {
        result = result * 26 + (c as i64 - 'A' as i64 + 1);
    }
    result - 1
}

fn build_cell_format(options: &Map<String, Value>) -> Value {
    let mut format = Map::new();
    let mut text_format = Map::new();

    for key in ["bold", "italic", "underline"] {
        if let Some(value) = options.get(key) {
            let target = match key {
                "bold" => "bold",
                "italic" => "italic",
                _ => "underline",
            };
            text_format.insert(target.to_string(), value.clone());
        }
    }

    if let Some(font_size) = options.get("font_size") {
        text_format.insert("fontSize".to_string(), font_size.clone());
    }
    if let Some(font_family) = options.get("font_family") {
        text_format.insert("fontFamily".to_string(), font_family.clone());
    }
    if let Some(foreground) = options.get("foreground_color") {
        text_format.insert(
            "foregroundColorStyle".to_string(),
            json!({"rgbColor": foreground}),
        );
    }

    if !text_format.is_empty() {
        format.insert("textFormat".to_string(), Value::Object(text_format));
    }

    if let Some(background) = options.get("background_color") {
        format.insert(
            "backgroundColorStyle".to_string(),
            json!({"rgbColor": background}),
        );
    }
    if let Some(horizontal) = options.get("horizontal_alignment") {
        format.insert("horizontalAlignment".to_string(), horizontal.clone());
    }
    if let Some(vertical) = options.get("vertical_alignment") {
        format.insert("verticalAlignment".to_string(), vertical.clone());
    }

    if let Some(number_format) = options.get("number_format").and_then(|v| v.as_object()) {
        format.insert(
            "numberFormat".to_string(),
            json!({
                "type": number_format.get("type"),
                "pattern": number_format.get("pattern")
            }),
        );
    }

    if let Some(wrap) = options.get("wrap_strategy") {
        format.insert("wrapStrategy".to_string(), wrap.clone());
    }

    if let Some(rotation) = options.get("text_rotation") {
        format.insert("textRotation".to_string(), json!({"angle": rotation}));
    }

    if let Some(borders) = options.get("borders").and_then(|v| v.as_object()) {
        format.insert("borders".to_string(), build_borders(borders));
    }

    Value::Object(format)
}

fn build_format_fields(options: &Map<String, Value>) -> String {
    let mut fields = Vec::new();
    if options.contains_key("bold") {
        fields.push("textFormat.bold");
    }
    if options.contains_key("italic") {
        fields.push("textFormat.italic");
    }
    if options.contains_key("underline") {
        fields.push("textFormat.underline");
    }
    if options.contains_key("font_size") {
        fields.push("textFormat.fontSize");
    }
    if options.contains_key("font_family") {
        fields.push("textFormat.fontFamily");
    }
    if options.contains_key("foreground_color") {
        fields.push("textFormat.foregroundColorStyle");
    }
    if options.contains_key("background_color") {
        fields.push("backgroundColorStyle");
    }
    if options.contains_key("horizontal_alignment") {
        fields.push("horizontalAlignment");
    }
    if options.contains_key("vertical_alignment") {
        fields.push("verticalAlignment");
    }
    if options.contains_key("number_format") {
        fields.push("numberFormat");
    }
    if options.contains_key("wrap_strategy") {
        fields.push("wrapStrategy");
    }
    if options.contains_key("text_rotation") {
        fields.push("textRotation");
    }
    if options.contains_key("borders") {
        fields.push("borders");
    }
    fields.join(",")
}

fn build_borders(border_config: &Map<String, Value>) -> Value {
    let mut borders = Map::new();

    for side in ["top", "bottom", "left", "right"] {
        let Some(side_cfg) = border_config.get(side).and_then(|v| v.as_object()) else {
            continue;
        };

        let mut border = Map::new();
        border.insert(
            "style".to_string(),
            side_cfg
                .get("style")
                .cloned()
                .unwrap_or(Value::String("SOLID".to_string())),
        );
        if let Some(color) = side_cfg.get("color") {
            border.insert("colorStyle".to_string(), json!({"rgbColor": color}));
        }

        borders.insert(side.to_string(), Value::Object(border));
    }

    Value::Object(borders)
}

fn build_chart_spec(chart_type: &str, title: &str, grid_range: &Value) -> Value {
    json!({
        "title": title,
        "basicChart": {
            "chartType": chart_type.to_uppercase(),
            "legendPosition": "BOTTOM_LEGEND",
            "domains": [{
                "domain": {
                    "sourceRange": {"sources": [grid_range]}
                }
            }],
            "series": [{
                "series": {
                    "sourceRange": {"sources": [grid_range]}
                },
                "targetAxis": "LEFT_AXIS"
            }],
            "headerCount": 1
        }
    })
}

fn build_conditional_format_rule(
    rule_type: &str,
    grid_range: &Value,
    params: &Map<String, Value>,
) -> Value {
    let mut rule = json!({
        "ranges": [grid_range]
    });

    if rule_type.eq_ignore_ascii_case("boolean") {
        let condition_type = params
            .get("condition_type")
            .and_then(|v| v.as_str())
            .unwrap_or("NUMBER_GREATER");

        let values = params
            .get("condition_values")
            .and_then(|v| v.as_array())
            .cloned()
            .unwrap_or_default()
            .into_iter()
            .map(|v| {
                let s = match v {
                    Value::String(s) => s,
                    other => other.to_string(),
                };
                json!({"userEnteredValue": s})
            })
            .collect::<Vec<_>>();

        let mut format = Map::new();
        if let Some(bg) = params.get("format_background_color") {
            format.insert("backgroundColorStyle".to_string(), json!({"rgbColor": bg}));
        }

        let mut text_format = Map::new();
        if let Some(bold) = params.get("format_bold").and_then(|v| v.as_bool())
            && bold
        {
            text_format.insert("bold".to_string(), Value::Bool(true));
        }
        if let Some(fg) = params.get("format_foreground_color") {
            text_format.insert("foregroundColorStyle".to_string(), json!({"rgbColor": fg}));
        }
        if !text_format.is_empty() {
            format.insert("textFormat".to_string(), Value::Object(text_format));
        }

        let mut boolean_rule = json!({
            "condition": {
                "type": condition_type,
            },
            "format": format
        });

        if !values.is_empty() {
            boolean_rule
                .as_object_mut()
                .expect("object")
                .get_mut("condition")
                .and_then(|c| c.as_object_mut())
                .expect("condition")
                .insert("values".to_string(), Value::Array(values));
        }

        rule.as_object_mut()
            .expect("object")
            .insert("booleanRule".to_string(), boolean_rule);
    } else if rule_type.eq_ignore_ascii_case("gradient") {
        let min_color = params.get("min_color").cloned().unwrap_or(json!({
            "red": 0.8,
            "green": 0.2,
            "blue": 0.2
        }));
        let max_color = params.get("max_color").cloned().unwrap_or(json!({
            "red": 0.2,
            "green": 0.8,
            "blue": 0.2
        }));

        let mut gradient_rule = json!({
            "minpoint": {
                "colorStyle": {"rgbColor": min_color},
                "type": params.get("min_type").cloned().unwrap_or(Value::String("MIN".to_string()))
            },
            "maxpoint": {
                "colorStyle": {"rgbColor": max_color},
                "type": params.get("max_type").cloned().unwrap_or(Value::String("MAX".to_string()))
            }
        });

        if let Some(mid_color) = params.get("mid_color") {
            gradient_rule
                .as_object_mut()
                .expect("object")
                .insert(
                    "midpoint".to_string(),
                    json!({
                        "colorStyle": {"rgbColor": mid_color},
                        "type": params.get("mid_type").cloned().unwrap_or(Value::String("PERCENTILE".to_string())),
                        "value": params.get("mid_value").cloned().unwrap_or(Value::String("50".to_string()))
                    }),
                );
        }

        rule.as_object_mut()
            .expect("object")
            .insert("gradientRule".to_string(), gradient_rule);
    }

    rule
}
