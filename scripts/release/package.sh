#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 3 ]]; then
  echo "Usage: $0 <version> <target> <output-dir>" >&2
  exit 1
fi

VERSION="$1"
TARGET="$2"
OUTPUT_DIR="$3"

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
BIN_DIR="${ROOT_DIR}/target/${TARGET}/release"

BIN_EXT=""
if [[ "${TARGET}" == *"windows"* ]]; then
  BIN_EXT=".exe"
fi

required_bins=(docs_manager drive_manager sheets_manager)
for bin in "${required_bins[@]}"; do
  if [[ ! -f "${BIN_DIR}/${bin}${BIN_EXT}" ]]; then
    echo "Missing binary: ${BIN_DIR}/${bin}${BIN_EXT}" >&2
    exit 1
  fi
done

PKG_NAME="google-docs-rust-${VERSION}-${TARGET}"
PKG_ROOT="${ROOT_DIR}/${OUTPUT_DIR}"
PKG_DIR="${PKG_ROOT}/${PKG_NAME}"

rm -rf "${PKG_DIR}"
mkdir -p "${PKG_DIR}/bin" "${PKG_DIR}/scripts"

for bin in "${required_bins[@]}"; do
  cp "${BIN_DIR}/${bin}${BIN_EXT}" "${PKG_DIR}/bin/${bin}${BIN_EXT}"
done

cp "${ROOT_DIR}/README.md" "${PKG_DIR}/README.md"
cp "${ROOT_DIR}/SKILL.md" "${PKG_DIR}/SKILL.md"
cp "${ROOT_DIR}/LICENSE" "${PKG_DIR}/LICENSE"
cp -R "${ROOT_DIR}/examples" "${PKG_DIR}/examples"
cp -R "${ROOT_DIR}/references" "${PKG_DIR}/references"

for bin in "${required_bins[@]}"; do
  cat > "${PKG_DIR}/scripts/${bin}" <<EOF
#!/usr/bin/env sh
set -eu
SCRIPT_DIR="\$(CDPATH= cd -- "\$(dirname -- "\$0")" && pwd)"
exec "\${SCRIPT_DIR}/../bin/${bin}${BIN_EXT}" "\$@"
EOF

  cat > "${PKG_DIR}/scripts/${bin}.rb" <<EOF
#!/usr/bin/env ruby
# frozen_string_literal: true

script_dir = File.expand_path(__dir__)
exec(File.join(script_dir, '..', 'bin', '${bin}${BIN_EXT}'), *ARGV)
EOF

done

chmod +x "${PKG_DIR}/scripts/docs_manager" "${PKG_DIR}/scripts/drive_manager" "${PKG_DIR}/scripts/sheets_manager"
chmod +x "${PKG_DIR}/scripts/docs_manager.rb" "${PKG_DIR}/scripts/drive_manager.rb" "${PKG_DIR}/scripts/sheets_manager.rb"

if [[ "${TARGET}" == *"windows"* ]]; then
  for bin in "${required_bins[@]}"; do
    cat > "${PKG_DIR}/scripts/${bin}.cmd" <<EOF
@echo off
set SCRIPT_DIR=%~dp0
"%SCRIPT_DIR%..\\bin\\${bin}${BIN_EXT}" %*
EOF
  done
fi

mkdir -p "${PKG_ROOT}"

if [[ "${TARGET}" == *"windows"* ]]; then
  (
    cd "${PKG_ROOT}"
    7z a -tzip "${PKG_NAME}.zip" "${PKG_NAME}" > /dev/null
  )
  echo "${OUTPUT_DIR}/${PKG_NAME}.zip"
else
  (
    cd "${PKG_ROOT}"
    tar -czf "${PKG_NAME}.tar.gz" "${PKG_NAME}"
  )
  echo "${OUTPUT_DIR}/${PKG_NAME}.tar.gz"
fi
