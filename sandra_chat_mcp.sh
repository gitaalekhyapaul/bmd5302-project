#!/usr/bin/env bash

set -euo pipefail

# Defaults match the LLM-backed MCP App chat server.
PORT="${1:-${SANDRA_CHAT_PORT:-8001}}"
HOST="${2:-${SANDRA_CHAT_HOST:-0.0.0.0}}"
STREAMABLE_HTTP_PATH="${3:-/mcp}"
TRANSPORT="${4:-streamable-http}"
MOUNT_PATH="${5:-}"
EXTRA_ARGS=("${@:6}")

CMD=(
  uv run python sandra_chat_server.py
  --transport "${TRANSPORT}"
  --host "${HOST}"
  --port "${PORT}"
  --streamable-http-path "${STREAMABLE_HTTP_PATH}"
)

if [[ -n "${MOUNT_PATH}" ]]; then
  CMD+=(--mount-path "${MOUNT_PATH}")
fi

if [[ ${#EXTRA_ARGS[@]} -gt 0 ]]; then
  CMD+=("${EXTRA_ARGS[@]}")
fi

DISPLAY_HOST="127.0.0.1"

echo
echo "======================================================================"
echo "  Sandra LLM Chat Browser + MCP Server Configuration"
echo "======================================================================"
echo
echo "Server endpoint:"
echo "  http://${DISPLAY_HOST}:${PORT}${STREAMABLE_HTTP_PATH}"
echo
echo "Browser chat UI:"
echo "  http://${DISPLAY_HOST}:${PORT}/app"
echo
echo "Expected workbook MCP endpoint:"
echo "  ${SANDRA_WORKBOOK_MCP_URL:-http://127.0.0.1:8000/mcp}"
echo
echo "Copy/paste MCP config:"
echo
cat <<EOF
{
  "mcpServers": {
    "sandra-chat": {
      "url": "http://${DISPLAY_HOST}:${PORT}${STREAMABLE_HTTP_PATH}"
    }
  }
}
EOF
echo
echo "Starting Sandra chat MCP server..."
echo "======================================================================"
echo

exec "${CMD[@]}"
