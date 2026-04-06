#!/usr/bin/env bash
# claude-pptx-mcp launcher — 毎回最新版を使う
# MCP起動時にgit pullしてから実行する
set -euo pipefail

REPO_DIR="${PPTX_MCP_DIR:-$HOME/.claude/mcp-servers/claude-pptx-mcp}"
REPO_URL="https://github.com/knishida-td/claude-pptx-mcp.git"

if [ -d "$REPO_DIR/.git" ]; then
  # 既存: pull
  cd "$REPO_DIR"
  git pull --ff-only -q 2>/dev/null || true
else
  # 初回: clone
  mkdir -p "$(dirname "$REPO_DIR")"
  git clone -q "$REPO_URL" "$REPO_DIR"
  cd "$REPO_DIR"
fi

# 依存が変わっていたら再インストール
if [ package.json -nt node_modules/.package-lock.json ] 2>/dev/null; then
  npm install --prefer-offline -q 2>/dev/null || npm install -q
fi

# ビルド済みでなければビルド
if [ ! -f dist/index.js ] || [ src/index.ts -nt dist/index.js ]; then
  npx tsc 2>/dev/null || true
fi

exec node dist/index.js
