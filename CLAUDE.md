# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Rust CLI toolkit for managing Google Docs, Google Drive, and Google Sheets. Ships as three separate binaries (`docs_manager`, `drive_manager`, `sheets_manager`) that share OAuth credentials and output JSON for agent consumption. Distributed as an [agent skill](https://agentskills.io).

## Build & Validate

```bash
cargo build --release          # Build all three binaries
cargo fmt --all --check        # Check formatting
cargo clippy --all-targets --all-features -- -D warnings  # Lint
cargo test --all-features      # Run tests
```

## Running the CLIs (Development)

Wrapper scripts in `scripts/` invoke `cargo run` automatically:

```bash
scripts/docs_manager --help
scripts/drive_manager --help
scripts/sheets_manager --help
```

## Architecture

### Binary Layout

Three binaries in `src/bin/`, each a self-contained CLI with its own `main()`, argument parsing, and Google API call implementations:

- **`docs_manager.rs`** (~1400 lines) — Document CRUD, markdown-to-Docs conversion, formatting, tables, images. Commands that mutate accept JSON via stdin. Includes a markdown parser that converts to Google Docs API `batchUpdate` requests.
- **`drive_manager.rs`** (~1100 lines) — File upload/download/list/search/share/move/copy/delete. Uses `--flag value` CLI args (not stdin JSON). Handles Google Apps file export (Docs→PDF, Sheets→CSV, etc.).
- **`sheets_manager.rs`** (~1850 lines) — Full spreadsheet operations: read/write/append/format/charts/conditional formatting/protection. All commands accept JSON via stdin. Converts A1 notation to `GridRange` objects internally.

### Shared Library (`src/lib.rs`)

Three modules re-exported from `src/lib.rs`:

- **`auth.rs`** — OAuth2 token flow: loads `client_secret.json`, builds auth URLs, exchanges codes, refreshes tokens, persists as YAML-wrapped JSON at `~/.claude/.google/token.json`. Tokens are shared across all three CLIs. Scopes include Drive, Docs, Sheets, Calendar, Contacts, and Gmail.
- **`google_api.rs`** — `GoogleClient` wrapping `reqwest::blocking::Client` with Bearer auth. Provides typed HTTP methods (`get_json`, `post_json`, `patch_json`, `delete_no_content`, `post_multipart`, `patch_multipart`, `get_bytes_to_path`). Shared error mapping via `map_api_error`.
- **`io_helpers.rs`** — JSON stdin reading, JSON pretty-print output, cross-platform home directory resolution.

### Key Patterns

- **All output is JSON** — success and error responses follow `{"status": "success/error", ...}` convention.
- **Exit codes are semantic**: 0=success, 1=operation failed, 2=auth error, 3=API error, 4=invalid args.
- **`docs_manager` and `sheets_manager`** use `read_stdin_json()` for mutation commands (pipe JSON to stdin). **`drive_manager`** uses `--flag value` CLI args exclusively.
- **Token storage** uses YAML with a `default` key wrapping a JSON string payload — not plain JSON. Both formats are readable but writes always use the YAML format.
- **Rust edition 2024** — uses let-chains (`if let Some(x) = ... && condition`).

## Release

Tag pushes (`v*`) trigger `.github/workflows/release.yml` building cross-platform archives (Linux musl, macOS, Windows) that include binaries, scripts, docs, and examples.
