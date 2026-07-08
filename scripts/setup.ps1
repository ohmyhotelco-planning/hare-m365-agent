Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

Write-Host "OMH M365 Agent setup"

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw "Node.js is not installed or not in PATH. Install Node.js 20 or later."
}

$nodeMajor = [int]((& node -p "process.versions.node.split('.')[0]") -as [string])
if ($nodeMajor -lt 20) {
  throw "Node.js 20 or later is required. Current major version: $nodeMajor"
}

if (-not (Test-Path -LiteralPath ".env")) {
  Copy-Item -LiteralPath ".env.example" -Destination ".env"
  Write-Host "Created .env from .env.example. Fill approved Azure Application values locally."
}

npm install
npm run typecheck
npm run build
npm run start -- doctor

Write-Host ""
Write-Host "Setup complete."
Write-Host "Next: fill .env, then run: npm run start -- auth login"
