param(
  [string]$Tag
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

$packageJson = Get-Content -LiteralPath "package.json" -Raw | ConvertFrom-Json
$version = $packageJson.version
if (-not $Tag) {
  $Tag = "v$version"
}

$repo = if ($env:HARE_M365_GITHUB_REPO) { $env:HARE_M365_GITHUB_REPO } else { "ohmyhotelco-planning/hare-m365-agent" }
$packageFile = "ohmyhotel-hare-m365-agent-$version.tgz"
$uploadOnly = Join-Path $root "releases\github-release\$Tag-upload-only"

function Assert-True([bool]$Condition, [string]$Message) {
  if (-not $Condition) {
    throw $Message
  }
}

function Assert-Contains([string]$Text, [string]$Needle, [string]$Label) {
  Assert-True ($Text.Contains($Needle)) "Missing required text in ${Label}: $Needle"
}

Assert-True (Test-Path -LiteralPath $uploadOnly) "Upload-only release folder does not exist: $uploadOnly"

$expectedFiles = @(
  $packageFile,
  "SHA256SUMS.txt"
) | Sort-Object

$actualFiles = @(Get-ChildItem -LiteralPath $uploadOnly -File | Select-Object -ExpandProperty Name | Sort-Object)
$diff = Compare-Object -ReferenceObject $expectedFiles -DifferenceObject $actualFiles
Assert-True (-not $diff) "Upload-only release folder does not contain exactly the expected files."

$tgzPath = Join-Path $uploadOnly $packageFile
$shaPath = Join-Path $uploadOnly "SHA256SUMS.txt"
$shaText = (Get-Content -LiteralPath $shaPath -Raw).Trim()
$expectedHash = ($shaText -split "\s+")[0].ToLowerInvariant()
$actualHash = (Get-FileHash -LiteralPath $tgzPath -Algorithm SHA256).Hash.ToLowerInvariant()
Assert-True ($expectedHash -eq $actualHash) "SHA256 mismatch for $packageFile. Expected $expectedHash, actual $actualHash."

$tarEntries = tar -tf $tgzPath
foreach ($forbiddenEntry in @(
  "package/.cache/",
  "package/runtime/",
  "package/downloads/",
  "package/logs/",
  "package/node_modules/",
  "package/releases/",
  "package/scripts/"
)) {
  Assert-True (-not ($tarEntries | Where-Object { $_ -like "$forbiddenEntry*" })) "Forbidden package entry found: $forbiddenEntry"
}

$oldDataDir = $env:HARE_M365_DATA_DIR
$smokeRuntime = Join-Path $env:TEMP "hare-m365-release-validate-runtime"
try {
  $env:HARE_M365_DATA_DIR = $smokeRuntime
  $startupOutput = (& node "dist\cli.js" | Out-String)
  Assert-Contains $startupOutput '"nextCommand": "LOGIN_REQUIRED_HARD_GATE"' "cli startup"
  Assert-Contains $startupOutput "npm exec --yes --package" "cli startup"
  Assert-Contains $startupOutput "HARE_M365_DATA_DIR" "cli startup"
}
finally {
  if ($null -eq $oldDataDir) {
    Remove-Item Env:\HARE_M365_DATA_DIR -ErrorAction SilentlyContinue
  } else {
    $env:HARE_M365_DATA_DIR = $oldDataDir
  }

  $resolved = Resolve-Path -LiteralPath $smokeRuntime -ErrorAction SilentlyContinue
  if ($resolved -and $resolved.Path.StartsWith($env:TEMP, [System.StringComparison]::OrdinalIgnoreCase)) {
    Remove-Item -LiteralPath $resolved.Path -Recurse -Force
  }
}

Write-Host "Release validation passed:"
Write-Host $uploadOnly
