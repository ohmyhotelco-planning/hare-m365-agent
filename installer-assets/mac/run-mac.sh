#!/usr/bin/env bash
set -euo pipefail

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
USER_DATA_DIR="${OMH_M365_USER_DATA_DIR:-$HOME/Library/Application Support/OMH/M365Agent}"
RUNTIME_DIR="$USER_DATA_DIR/runtime"

NODE_VERSION="24.15.0"
ARCH="$(uname -m)"

case "$ARCH" in
  arm64)
    NODE_DIST="node-v${NODE_VERSION}-darwin-arm64"
    ;;
  x86_64)
    NODE_DIST="node-v${NODE_VERSION}-darwin-x64"
    ;;
  *)
    echo "Unsupported Mac architecture: $ARCH"
    echo "Supported: Apple Silicon arm64, Intel x86_64"
    exit 2
    ;;
esac

mkdir -p "$USER_DATA_DIR/.cache" "$USER_DATA_DIR/downloads" "$USER_DATA_DIR/logs" "$RUNTIME_DIR"
chmod 700 "$USER_DATA_DIR" "$USER_DATA_DIR/.cache" "$USER_DATA_DIR/logs" 2>/dev/null || true

if [ -f "$APP_DIR/.env" ]; then
  set -a
  # shellcheck disable=SC1091
  . "$APP_DIR/.env"
  set +a
fi

export OMH_M365_POLICY_PATH="$APP_DIR/policy.json"
export OMH_M365_CACHE_DIR="$USER_DATA_DIR/.cache"
export OMH_M365_DOWNLOAD_DIR="$USER_DATA_DIR/downloads"

NODE_HOME="$RUNTIME_DIR/$NODE_DIST"
NODE_BIN="$NODE_HOME/bin/node"
NODE_ARCHIVE="$APP_DIR/runtime/$NODE_DIST.tar.gz"

if [ ! -x "$NODE_BIN" ]; then
  if [ ! -f "$NODE_ARCHIVE" ]; then
    echo "Bundled Node runtime was not found:"
    echo "  $NODE_ARCHIVE"
    echo "Ask the distributor for a complete Mac package."
    exit 2
  fi

  tar -xzf "$NODE_ARCHIVE" -C "$RUNTIME_DIR"
fi

cd "$USER_DATA_DIR"
exec "$NODE_BIN" "$APP_DIR/omh-m365.cjs" "$@"
