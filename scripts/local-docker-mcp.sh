#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
IMAGE_NAME="spreadsheet-mcp-full:test"

docker build -q -f "$PROJECT_ROOT/Dockerfile.full" -t "$IMAGE_NAME" "$PROJECT_ROOT" >/dev/null

exec docker run --rm -i \
    -v "${WORKSPACE_ROOT:-$PROJECT_ROOT/stest}:/data" \
    "$IMAGE_NAME" \
    --workspace-root /data \
    --transport stdio \
    --recalc-enabled \
    --vba-enabled
