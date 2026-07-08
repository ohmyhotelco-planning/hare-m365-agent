#!/usr/bin/env bash
# OMH M365 Agent - Cowork runner for a mounted package folder.
set -euo pipefail

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
USER_DATA_DIR="${OMH_M365_USER_DATA_DIR:-$HOME/Library/Application Support/OMH/M365Agent}"

mkdir -p "$USER_DATA_DIR/.cache" "$USER_DATA_DIR/downloads" "$USER_DATA_DIR/logs"
chmod 700 "$USER_DATA_DIR" "$USER_DATA_DIR/.cache" "$USER_DATA_DIR/logs" 2>/dev/null || true

if [ -f "$APP_DIR/.env" ]; then
  set -a
  # shellcheck disable=SC1091
  . "$APP_DIR/.env"
  set +a
fi

export NODE_USE_ENV_PROXY=1
export NODE_NO_WARNINGS=1
export OMH_M365_POLICY_PATH="$APP_DIR/policy.json"
export OMH_M365_CACHE_DIR="$USER_DATA_DIR/.cache"
export OMH_M365_DOWNLOAD_DIR="$USER_DATA_DIR/downloads"

cd "$USER_DATA_DIR"
exec node "$APP_DIR/omh-m365.cjs" "$@"
