#!/usr/bin/env bash
# Auto-setup script for claude-pptx-mcp dependencies
# Called automatically on first run or when dependencies are missing

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
VENV_DIR="$PROJECT_ROOT/.venv"
MARKER="$VENV_DIR/.setup-done"

# Skip if already set up
if [ -f "$MARKER" ]; then
  exit 0
fi

echo "[claude-pptx-mcp] Setting up dependencies..." >&2

# --- Python venv + packages ---
if [ ! -d "$VENV_DIR" ]; then
  echo "[claude-pptx-mcp] Creating Python venv..." >&2
  python3 -m venv "$VENV_DIR"
fi

echo "[claude-pptx-mcp] Installing Python packages..." >&2
"$VENV_DIR/bin/pip" install --quiet --upgrade pip
"$VENV_DIR/bin/pip" install --quiet \
  python-pptx \
  Pillow \
  defusedxml \
  six \
  markitdown

# --- Node packages (pptxgenjs etc) ---
if [ ! -d "$PROJECT_ROOT/node_modules/pptxgenjs" ]; then
  echo "[claude-pptx-mcp] Installing Node packages..." >&2
  cd "$PROJECT_ROOT"
  npm install --save --quiet pptxgenjs sharp 2>&1 >&2 || true
fi

# --- LibreOffice (macOS only, via brew) ---
if ! command -v soffice &>/dev/null; then
  if command -v brew &>/dev/null; then
    echo "[claude-pptx-mcp] Installing LibreOffice via Homebrew (this may take a few minutes)..." >&2
    brew install --cask libreoffice 2>&1 >&2 || true
  else
    echo "[claude-pptx-mcp] WARNING: LibreOffice not found. Thumbnail generation will be limited." >&2
    echo "[claude-pptx-mcp] Install manually: https://www.libreoffice.org/download/" >&2
  fi
fi

# Mark as done
touch "$MARKER"
echo "[claude-pptx-mcp] Setup complete." >&2
