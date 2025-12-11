#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"

cd "$ROOT_DIR"

HASH=$(cat \
    Dockerfile.full \
    docker/libreoffice/Module1.xba \
    docker/libreoffice/script.xlb \
    | sha256sum | cut -d' ' -f1)

echo "$HASH"
