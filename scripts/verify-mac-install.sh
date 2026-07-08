#!/usr/bin/env bash
set -euo pipefail

APP_DIR="/Library/Application Support/OMH/M365Agent"
RUNNER="$APP_DIR/run-mac.sh"

if [ "$(uname -s)" != "Darwin" ]; then
  echo "This verification script must be run on macOS."
  exit 2
fi

if [ ! -x "$RUNNER" ]; then
  echo "OMH M365 Agent is not installed or run-mac.sh is not executable:"
  echo "  $RUNNER"
  exit 2
fi

echo "Checking installed files..."
test -f "$APP_DIR/omh-m365.cjs"
test -f "$APP_DIR/policy.json"
test -f "$APP_DIR/.env"
test -f "$APP_DIR/start.html"
test -d "$APP_DIR/runtime"

echo
echo "Running doctor..."
"$RUNNER" doctor

echo
echo "Running auth status..."
AUTH_OUTPUT="$("$RUNNER" auth status || true)"
printf "%s\n" "$AUTH_OUTPUT"

if ! printf "%s\n" "$AUTH_OUTPUT" | grep -q '"loggedIn"[[:space:]]*:[[:space:]]*true'; then
  echo
  echo "Login is required."
  echo "Open /Applications/OMH M365 Agent Login.app, complete Microsoft device login, then run this script again."
  exit 1
fi

if [ "${1:-}" = "--graph-smoke" ]; then
  echo
  echo "Running Outlook smoke test..."
  "$RUNNER" outlook inbox --limit 1

  echo
  echo "Running Teams smoke test..."
  "$RUNNER" teams chats --limit 5

  echo
  echo "Running Files smoke test..."
  "$RUNNER" files search --query test --limit 1
fi

echo
echo "Mac installation verification completed."
