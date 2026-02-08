# Google Docs Skill (Rust)

An [agent skill](https://agentskills.io) for managing Google Docs, Google Sheets, and Google Drive. Zero-dependency native binaries — no Ruby, Python, or Node.js runtime needed.

## Features

### Google Docs
- Read document content and structure (headings)
- Create documents from Markdown with formatting, tables, and checkboxes
- Insert, append, find/replace, and delete text
- Text formatting (bold, italic, underline)
- Insert page breaks, inline images, and tables

### Google Sheets
- Create, read, write, and append spreadsheet data
- Batch read/write across multiple ranges
- Format cells (bold, colors, fonts, alignment, borders, number formats)
- Merge/unmerge cells, freeze rows/columns, sort, find/replace
- Add charts, filters, conditional formatting, protected ranges
- Manage sheets/tabs (add, delete, rename, copy)

### Google Drive
- Upload, download, and update files
- Search, list, and get file metadata
- Create folders, move, copy, and delete files
- Share files with users or publicly
- Export Google Docs/Sheets/Slides to PDF, CSV, etc.

## Install as a skill

```bash
npx skills add jimmystridh/google-docs-skill --skill google-docs-skill -g -a claude-code -y
```

Or by repository URL:

```bash
npx skills add https://github.com/jimmystridh/google-docs-skill --skill google-docs-skill -g -a claude-code -y
```

Release archives include prebuilt binaries for Linux, macOS, and Windows — no Rust toolchain required.

## Auth setup

1. Create a Google Cloud project and enable the **Drive**, **Docs**, and **Sheets** APIs.
2. Create OAuth 2.0 credentials (Desktop application type).
3. Save the downloaded client JSON:
   - macOS/Linux: `~/.claude/.google/client_secret.json`
   - Windows: `%USERPROFILE%\.claude\.google\client_secret.json`
4. Run any command to trigger the auth flow:

```bash
scripts/drive_manager list --max-results 1
```

If not yet authorized, you'll get a JSON response with an `auth_url`. Open it in your browser and complete the consent flow.

5. Store the token:

```bash
scripts/docs_manager auth <code>
```

Tokens are stored at `~/.claude/.google/token.json` and shared across all three tools. The auth URL also requests scopes for Calendar, Contacts, and Gmail for use with related Google skills.

## Usage

```bash
scripts/docs_manager --help
scripts/drive_manager --help
scripts/sheets_manager --help
```

On Windows release archives: `scripts\docs_manager.cmd`, etc.

## Building from source

```bash
cargo build --release
```

Binaries are output to `target/release/`. The `scripts/` wrappers invoke `cargo run` for development.

### Validation

```bash
cargo fmt --all --check
cargo clippy --all-targets --all-features -- -D warnings
cargo test --all-features
```

## Quickstart (Release Archive)

1. Download the archive matching your platform from GitHub Releases.
2. Extract and run:

```bash
tar -xzf google-docs-skill-vX.Y.Z-aarch64-apple-darwin.tar.gz
cd google-docs-skill-vX.Y.Z-aarch64-apple-darwin
scripts/drive_manager --help
```

### macOS Gatekeeper

If macOS blocks the binaries after a browser download:

```bash
xattr -dr com.apple.quarantine google-docs-skill-vX.Y.Z-aarch64-apple-darwin
```

Downloading via `gh release download` or `curl` avoids this.

## Release

Tag pushes (`v*`) trigger CI to build archives for all six targets (x86_64/aarch64 for Linux musl, macOS, Windows).

```bash
git tag v0.1.0
git push origin v0.1.0
```

## License

MIT
