#!/usr/bin/env bash

set -euo pipefail

# Defaults match the current HTTP MCP setup.
PORT="${1:-8000}"
HOST="${2:-0.0.0.0}"
STREAMABLE_HTTP_PATH="${3:-/mcp}"
TRANSPORT="${4:-streamable-http}"
MOUNT_PATH="${5:-}"
EXTRA_ARGS=("${@:6}")

CMD=(
  uv run python mcp_server.py
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
echo "  Excel MCP Server Configuration"
echo "======================================================================"
echo
echo "Server endpoint:"
echo "  http://${DISPLAY_HOST}:${PORT}${STREAMABLE_HTTP_PATH}"
echo
echo "Copy/paste MCP config:"
echo
cat <<EOF
{
  "mcpServers": {
    "excel-workbook": {
      "url": "http://${DISPLAY_HOST}:${PORT}${STREAMABLE_HTTP_PATH}"
    }
  }
}
EOF
echo
echo "Disclaimer: for non-local/public access, replace 127.0.0.1 with a reachable IP or domain."
echo
echo "Starting MCP server..."
echo "======================================================================"
echo

exec "${CMD[@]}"
