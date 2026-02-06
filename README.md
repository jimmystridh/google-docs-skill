# Google Docs Skill (Rust Port)

Rust port of the original Google Docs/Drive/Sheets skill with the same command surface and JSON output patterns.

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

Compatibility wrappers are also included for existing references:

```bash
scripts/docs_manager.rb --help
scripts/drive_manager.rb --help
scripts/sheets_manager.rb --help
```

## Install as a skill

Install directly from GitHub with the community `skills` CLI:

```bash
npx skills add jimmystridh/google-docs-rust --skill google-docs-rust -g -a claude-code -y
```

Repository URL also works:

```bash
npx skills add https://github.com/jimmystridh/google-docs-rust --skill google-docs-rust -g -a claude-code -y
```

## Auth setup

1. Create Google Cloud OAuth Desktop credentials.
2. Save client secret JSON to:
   - `~/.claude/.google/client_secret.json`
3. Complete auth once:

```bash
scripts/docs_manager auth <code>
# or
scripts/sheets_manager auth <code>
```

Drive commands use the same shared token and do not have a separate `auth` command.

## Validation

This port was validated with:

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
