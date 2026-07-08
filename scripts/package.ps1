Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

$packageJson = Get-Content -LiteralPath "package.json" -Raw | ConvertFrom-Json
$version = $packageJson.version
$releaseDir = Join-Path $root "releases"
$stageRoot = Join-Path $releaseDir "stage"
$stage = Join-Path $stageRoot "omh-m365-agent"
$zipPath = Join-Path $releaseDir "omh-m365-agent-$version.zip"

if (Test-Path -LiteralPath ".env") {
  Write-Host "Note: .env exists locally and will be packaged as delegated public-client configuration."
}

npm run typecheck
npm run build

New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
if (Test-Path -LiteralPath $stageRoot) {
  Remove-Item -LiteralPath $stageRoot -Recurse -Force
}
if (Test-Path -LiteralPath $zipPath) {
  Remove-Item -LiteralPath $zipPath -Force
}

New-Item -ItemType Directory -Path $stage -Force | Out-Null

$files = @(
  "package.json",
  "package-lock.json",
  "tsconfig.json",
  "README.md",
  "START_HERE_FOR_LLM.md",
  "AGENTS.md",
  "CLAUDE.md",
  "CLAUDE_COWORK_RUNBOOK.md",
  "policy.json",
  ".env.example",
  ".gitignore",
  "config.example.json",
  "RUN_FIRST_FOR_LLM.cmd",
  "START_LOGIN_FOR_USER.cmd",
  "RUN_FIRST_FOR_LLM.sh",
  "START_LOGIN_FOR_USER.sh"
)

foreach ($file in $files) {
  Copy-Item -LiteralPath (Join-Path $root $file) -Destination (Join-Path $stage $file)
}

if (Test-Path -LiteralPath ".env") {
  Copy-Item -LiteralPath (Join-Path $root ".env") -Destination (Join-Path $stage ".env")
}

foreach ($dir in @("src", "dist", "docs", "scripts")) {
  Copy-Item -LiteralPath (Join-Path $root $dir) -Destination (Join-Path $stage $dir) -Recurse
}

$blocked = @(".cache", "downloads", "logs", "node_modules", "releases")
foreach ($name in $blocked) {
  if (Test-Path -LiteralPath (Join-Path $stage $name)) {
    throw "Blocked path was staged unexpectedly: $name"
  }
}

$secretPattern = '(client_secret|refresh_token|access_token|password|cookie)[[:space:]]*[:=]|Bearer[[:space:]][A-Za-z0-9._-]{20,}'
$secretHits = rg -l --glob '!omh-m365.cjs' -- "$secretPattern" $stage 2>$null
if ($LASTEXITCODE -eq 0) {
  throw "Potential secret pattern found in staged package. Refusing to package secret-like values."
}

Compress-Archive -LiteralPath $stage -DestinationPath $zipPath -Force
Remove-Item -LiteralPath $stageRoot -Recurse -Force

Write-Host "Created package: $zipPath"
