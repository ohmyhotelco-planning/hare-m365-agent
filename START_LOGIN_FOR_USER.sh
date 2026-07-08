#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

echo "OMH M365 Agent - Microsoft login"
echo
echo "The human user must complete Microsoft device login."
echo "LLMs must not copy, repeat, store, or summarize the device code."
echo

if command -v xdg-open >/dev/null 2>&1; then
  xdg-open "https://microsoft.com/devicelogin" >/dev/null 2>&1 || true
elif command -v open >/dev/null 2>&1; then
  open "https://microsoft.com/devicelogin" >/dev/null 2>&1 || true
fi

if [ -f omh-m365.cjs ]; then
  node omh-m365.cjs auth login
elif [ "${OMH_M365_ALLOW_SOURCE_CLI:-0}" = "1" ] && [ -f dist/cli.js ] && [ -d node_modules ]; then
  node dist/cli.js auth login
else
  echo "Bundled CLI omh-m365.cjs was not found."
  echo "Do not run npm ci/npm install in Cowork-style LLM containers."
  echo "Use a packaged linux-llm release that contains omh-m365.cjs."
  exit 2
fi
