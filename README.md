# Google Docs Skill

Manage Google Docs, Google Drive, and Google Sheets via three CLIs with JSON-friendly input/output.

## What is included

- `docs_manager` CLI: Google Docs operations (read, structure, edit, markdown, tables, images)
- `drive_manager` CLI: Google Drive file/folder operations
- `sheets_manager` CLI: Google Sheets data + formatting/chart/protection operations
- Shared OAuth token flow using:
  - `~/.claude/.google/client_secret.json`
  - `~/.claude/.google/token.json`

## Repository layout

- `src/bin/docs_manager.rs`
- `src/bin/drive_manager.rs`
- `src/bin/sheets_manager.rs`
- `src/auth.rs` shared OAuth logic
- `src/google_api.rs` shared Google HTTP client
- `scripts/*` runnable wrappers
- `references/*` operation guides and troubleshooting
- `examples/sample_operations.md` end-to-end examples

## Build

```bash
cargo build --release
```

## Usage

Use wrapper scripts (recommended):

```bash
scripts/docs_manager --help
scripts/drive_manager --help
scripts/sheets_manager --help
```

On Windows, the release archives also include `scripts\\docs_manager.cmd`, `scripts\\drive_manager.cmd`, and `scripts\\sheets_manager.cmd`.

Compatibility wrappers are also included for existing references:

```bash
scripts/docs_manager.rb --help
scripts/drive_manager.rb --help
scripts/sheets_manager.rb --help
```

## Quickstart (Release Archive)

Release archives include prebuilt binaries and runnable wrappers, so you do not need Rust/Cargo installed.

1. Download the archive matching your platform from GitHub Releases.
2. Extract it (it contains a single directory named `google-docs-skill-<version>-<target>/`).
3. Run the wrappers from inside that directory.

macOS/Linux example:

```bash
tar -xzf google-docs-skill-vX.Y.Z-aarch64-apple-darwin.tar.gz
cd google-docs-skill-vX.Y.Z-aarch64-apple-darwin
scripts/drive_manager --help
```

Windows example (PowerShell):

```powershell
Expand-Archive .\\google-docs-skill-vX.Y.Z-x86_64-pc-windows-msvc.zip -DestinationPath .
cd .\\google-docs-skill-vX.Y.Z-x86_64-pc-windows-msvc
.\\scripts\\drive_manager.cmd --help
```

## Install as a skill

Install directly from GitHub with the community `skills` CLI:

```bash
npx skills add jimmystridh/google-docs-skill --skill google-docs-skill -g -a claude-code -y
```

Repository URL also works:

```bash
npx skills add https://github.com/jimmystridh/google-docs-skill --skill google-docs-skill -g -a claude-code -y
```

## Auth setup

The tools look for OAuth credentials under your home directory:

- macOS/Linux: `~/.claude/.google/client_secret.json`
- Windows: `%USERPROFILE%\\.claude\\.google\\client_secret.json`

1. Create Google Cloud OAuth credentials (OAuth Client ID: **Desktop app**).
2. Enable the APIs in your Google Cloud project:
   - Google Drive API
   - Google Docs API
   - Google Sheets API
3. Save the downloaded OAuth client JSON to:
   - `~/.claude/.google/client_secret.json`
   - (Create the folder first if needed: `mkdir -p ~/.claude/.google`)
4. Trigger the auth flow once to get an authorization URL:

```bash
scripts/drive_manager list --max-results 1
```

If you are not authorized yet, you will get a JSON error containing `auth_url`. Open that URL in your browser and complete the consent flow to get an authorization code.

5. Complete auth by storing the token:

```bash
scripts/docs_manager auth <code>
# or
scripts/sheets_manager auth <code>
```

Tokens are stored at `~/.claude/.google/token.json`. Drive commands use the same shared token and do not have a separate `auth` command.

Note: For convenience across related Google skills, the authorization URL requests scopes for Drive/Docs/Sheets plus Calendar, Contacts (People API), and Gmail modify.

## Validation

Validated with:

```bash
cargo check --offline
cargo clippy --offline --all-targets --all-features
```

## Release Matrix

Tag pushes (`v*`) trigger `.github/workflows/release.yml`, which builds and publishes archives for:

- `x86_64-unknown-linux-musl`
- `aarch64-unknown-linux-musl`
- `x86_64-pc-windows-msvc`
- `aarch64-pc-windows-msvc`
- `x86_64-apple-darwin`
- `aarch64-apple-darwin`

Each archive contains:

- `bin/docs_manager`, `bin/drive_manager`, `bin/sheets_manager` (or `.exe` on Windows)
- runnable wrappers under `scripts/`
- `SKILL.md`, `README.md`, `LICENSE`, `examples/`, `references/`

Release process:

```bash
git tag v0.1.0
git push origin v0.1.0
```

The release workflow uploads archives that can be used directly by agents without a local Rust toolchain.

## License

MIT (same as original)
