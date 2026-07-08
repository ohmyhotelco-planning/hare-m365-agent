#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

echo "OMH M365 Agent - LLM startup check"
echo
echo "[1/5] cwd"
pwd
echo
echo "[2/5] node"
node --version
echo
echo "[3/5] command"
if [ -f omh-m365.cjs ]; then
  CLI=(node omh-m365.cjs)
  echo "Using bundled CLI: node omh-m365.cjs"
else
  if [ "${OMH_M365_ALLOW_NPM_INSTALL:-0}" != "1" ]; then
    echo "Bundled CLI omh-m365.cjs was not found."
    echo "Do not run npm ci/npm install in Cowork-style LLM containers."
    echo "Use a packaged linux-llm release that contains omh-m365.cjs, or rebuild on a developer machine with:"
    echo "  npm run package:linux"
    exit 2
  fi

  echo "Bundled CLI not found. OMH_M365_ALLOW_NPM_INSTALL=1 is set, so source fallback is allowed."
  echo
  echo "[3a] dependencies"
  test -d node_modules || npm ci
  echo
  echo "[4/5] build"
  test -f dist/cli.js || npm run build
  CLI=(node dist/cli.js)
  echo "Using source CLI: node dist/cli.js"
fi
echo
echo "[5/5] status"
STATUS_JSON="$("${CLI[@]}" doctor)"
echo "$STATUS_JSON"

CONFIGURED="$(STATUS_JSON="$STATUS_JSON" node -e 'const s=JSON.parse(process.env.STATUS_JSON); process.stdout.write(String(Boolean(s.configured)));')"
LOGGED_IN="$(STATUS_JSON="$STATUS_JSON" node -e 'const s=JSON.parse(process.env.STATUS_JSON); process.stdout.write(String(Boolean(s.loggedIn)));')"

if [ "$CONFIGURED" != "true" ]; then
  echo
  echo "Configuration is missing."
  echo "The human user must create .env from .env.example and fill approved Azure Application values locally."
  echo "Do not ask the user to paste client ID or tenant ID into chat."
  exit 3
fi

if [ "$LOGGED_IN" != "true" ]; then
  echo
  echo "Authentication is missing or expired."
  echo "Before any Outlook, Teams, or Files command, complete login first:"
  echo "  ./START_LOGIN_FOR_USER.sh"
  echo "After the human says login is complete, rerun this startup check and then resume the original task."
  exit 4
fi

"${CLI[@]}" auth status
echo
echo "Ready for read-only Microsoft 365 commands."
