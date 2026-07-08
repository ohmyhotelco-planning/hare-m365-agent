Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

$releaseDir = Join-Path $root "releases"
$stage = Join-Path $releaseDir "linux-llm"

if (Test-Path -LiteralPath ".env") {
  Write-Host "Note: .env exists locally and will be packaged as delegated public-client configuration."
}

npm run typecheck
npm run build

if (Test-Path -LiteralPath $stage) {
  Remove-Item -LiteralPath $stage -Recurse -Force
}
New-Item -ItemType Directory -Path $stage -Force | Out-Null

$files = @(
  "README.md",
  "START_HERE_FOR_LLM.md",
  "AGENTS.md",
  "CLAUDE.md",
  "COWORK_REQUIRED_README.md",
  "policy.json",
  ".env.example",
  "config.example.json",
  "LOGIN_FOR_COWORK.cmd",
  "run-cowork.sh",
  "RUN_FIRST_FOR_LLM.sh",
  "START_LOGIN_FOR_USER.sh"
)

foreach ($file in $files) {
  Copy-Item -LiteralPath (Join-Path $root $file) -Destination (Join-Path $stage $file)
}

if (Test-Path -LiteralPath ".env") {
  Copy-Item -LiteralPath (Join-Path $root ".env") -Destination (Join-Path $stage ".env")
}

npx esbuild src\cli.ts `
  --bundle `
  --platform=node `
  --target=node20 `
  --format=cjs `
  --outfile="$(Join-Path $stage 'omh-m365.cjs')"

$bundlePath = Join-Path $stage "omh-m365.cjs"
node --check $bundlePath
$bundleHash = Get-FileHash -LiteralPath $bundlePath -Algorithm SHA256
Set-Content -LiteralPath (Join-Path $stage "omh-m365.cjs.sha256") -Value "$($bundleHash.Hash.ToLower())  omh-m365.cjs" -Encoding ASCII

$blocked = @(".cache", "downloads", "logs", "node_modules", "build", "releases")
foreach ($name in $blocked) {
  if (Test-Path -LiteralPath (Join-Path $stage $name)) {
    throw "Blocked path was staged unexpectedly: $name"
  }
}

$secretPattern = '(client_secret|refresh_token|access_token|password|cookie)[[:space:]]*[:=]|Bearer[[:space:]][A-Za-z0-9._-]{20,}'
$secretHits = rg -l --glob '!omh-m365.cjs' -- "$secretPattern" $stage 2>$null
if ($LASTEXITCODE -eq 0) {
  throw "Potential secret pattern found in Linux release. Refusing to package secret-like values."
}

Write-Host "Created Linux LLM release folder:"
Write-Host $stage
Write-Host ""
Write-Host "Bundled CLI: omh-m365.cjs"
Write-Host "On Linux, run: chmod +x run-cowork.sh RUN_FIRST_FOR_LLM.sh START_LOGIN_FOR_USER.sh"
