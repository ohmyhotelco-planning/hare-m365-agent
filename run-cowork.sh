#!/usr/bin/env bash
# OMH M365 Agent - Cowork runner
# Routes MSAL + Microsoft Graph traffic through the Cowork host proxy
# (only login.microsoftonline.com + graph.microsoft.com must be allowlisted),
# and uses the packaged self-contained bundle. Read-only POC usage only.
set -euo pipefail
cd "$(dirname "$0")"
export NODE_USE_ENV_PROXY=1     # Node 22 built-in: make fetch + https honor HTTPS_PROXY
export NODE_NO_WARNINGS=1       # silence the experimental-proxy notice on stderr
exec node omh-m365.cjs "$@"
