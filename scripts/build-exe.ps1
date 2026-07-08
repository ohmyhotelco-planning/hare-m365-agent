Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

$packageJson = Get-Content -LiteralPath "package.json" -Raw | ConvertFrom-Json
$buildDir = Join-Path $root "build\sea"
$releaseDir = Join-Path $root "releases\win-x64"
$templateDir = Join-Path $root "release-templates\win-x64"
$bundlePath = Join-Path $buildDir "cli.cjs"
$blobPath = Join-Path $buildDir "omh-m365.blob"
$seaConfigPath = Join-Path $buildDir "sea-config.json"
$buildExePath = Join-Path $buildDir "omh-m365.exe"
$releaseExePath = Join-Path $releaseDir "omh-m365.exe"

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw "Node.js is not installed or not in PATH."
}

$nodeMajor = [int]((& node -p "process.versions.node.split('.')[0]") -as [string])
if ($nodeMajor -lt 20) {
  throw "Node.js 20 or later is required. Current major version: $nodeMajor"
}

npm run typecheck

New-Item -ItemType Directory -Path $buildDir -Force | Out-Null
New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null

npx esbuild src\cli.ts `
  --bundle `
  --platform=node `
  --target=node20 `
  --format=cjs `
  --outfile="$bundlePath"

$seaConfig = [ordered]@{
  main = "build/sea/cli.cjs"
  output = "build/sea/omh-m365.blob"
  disableExperimentalSEAWarning = $true
  useSnapshot = $false
  useCodeCache = $true
}

$seaConfig | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $seaConfigPath -Encoding UTF8
node --experimental-sea-config $seaConfigPath

$nodeExe = & node -p "process.execPath"
Copy-Item -LiteralPath $nodeExe -Destination $buildExePath -Force

npx postject $buildExePath NODE_SEA_BLOB $blobPath --sentinel-fuse NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2

Copy-Item -LiteralPath $buildExePath -Destination $releaseExePath -Force

foreach ($file in @("README.md", "START_HERE_FOR_LLM.md", "AGENTS.md", "CLAUDE.md")) {
  $templatePath = Join-Path $templateDir $file
  $sourcePath = if (Test-Path -LiteralPath $templatePath) { $templatePath } else { Join-Path $root $file }
  Copy-Item -LiteralPath $sourcePath -Destination (Join-Path $releaseDir $file) -Force
}

Copy-Item -LiteralPath "CLAUDE_COWORK_RUNBOOK.md" -Destination (Join-Path $releaseDir "CLAUDE_COWORK_RUNBOOK.md") -Force
Copy-Item -LiteralPath "policy.json" -Destination (Join-Path $releaseDir "policy.json") -Force
Copy-Item -LiteralPath ".env.example" -Destination (Join-Path $releaseDir ".env.example") -Force
if (Test-Path -LiteralPath ".env") {
  Copy-Item -LiteralPath ".env" -Destination (Join-Path $releaseDir ".env") -Force
}
Copy-Item -LiteralPath "RUN_FIRST_FOR_LLM.cmd" -Destination (Join-Path $releaseDir "RUN_FIRST_FOR_LLM.cmd") -Force
Copy-Item -LiteralPath "START_LOGIN_FOR_USER.cmd" -Destination (Join-Path $releaseDir "START_LOGIN_FOR_USER.cmd") -Force

if (Test-Path -LiteralPath (Join-Path $releaseDir "docs")) {
  Remove-Item -LiteralPath (Join-Path $releaseDir "docs") -Recurse -Force
}
$templateDocs = Join-Path $templateDir "docs"
if (Test-Path -LiteralPath $templateDocs) {
  Copy-Item -LiteralPath $templateDocs -Destination (Join-Path $releaseDir "docs") -Recurse
} else {
  Copy-Item -LiteralPath "docs" -Destination (Join-Path $releaseDir "docs") -Recurse
}

$runtimeLocal = @(".env", ".cache", "downloads", "logs")
foreach ($name in $runtimeLocal) {
  if (Test-Path -LiteralPath (Join-Path $releaseDir $name)) {
    Write-Host "Note: local runtime path exists in the installed release folder and was not touched: $name"
  }
}

$blocked = @("node_modules", "msal-cache.json")
foreach ($name in $blocked) {
  if (Test-Path -LiteralPath (Join-Path $releaseDir $name)) {
    throw "Blocked path was staged unexpectedly: $name"
  }
}

$exeHash = Get-FileHash -LiteralPath $releaseExePath -Algorithm SHA256
Set-Content -LiteralPath (Join-Path $releaseDir "omh-m365.exe.sha256") -Value "$($exeHash.Hash.ToLower())  omh-m365.exe" -Encoding ASCII

Write-Host "Created Windows executable release:"
Write-Host $releaseExePath
Write-Host ""
Write-Host "Note: this POC executable is not code-signed. Windows or endpoint security may show warnings."
