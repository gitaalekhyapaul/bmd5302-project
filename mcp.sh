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
  uv run python excel_mcp_server.py
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

exec "${CMD[@]}"
