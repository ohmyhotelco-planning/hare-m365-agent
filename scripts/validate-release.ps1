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
$packageUrl = "https://github.com/$repo/releases/download/$Tag/$packageFile"
$uploadOnly = Join-Path $root "releases\github-release\$Tag-upload-only"

function Assert-True([bool]$Condition, [string]$Message) {
  if (-not $Condition) {
    throw $Message
  }
}

function Read-Utf8([string]$Path) {
  return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
}

function From-Base64Utf8([string]$Value) {
  return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($Value))
}

function Assert-Contains([string]$Text, [string]$Needle, [string]$Label) {
  Assert-True ($Text.Contains($Needle)) "Missing required text in ${Label}: $Needle"
}

function Assert-NotContains([string]$Text, [string]$Needle, [string]$Label) {
  Assert-True (-not $Text.Contains($Needle)) "Forbidden text found in ${Label}: $Needle"
}

function Assert-Order([string]$Text, [string[]]$Needles, [string]$Label) {
  $last = -1
  foreach ($needle in $Needles) {
    $current = $Text.IndexOf($needle, [System.StringComparison]::Ordinal)
    Assert-True ($current -ge 0) "Missing ordered text in ${Label}: $needle"
    Assert-True ($current -gt $last) "Wrong text order in ${Label}: $needle"
    $last = $current
  }
}

Assert-True (Test-Path -LiteralPath $uploadOnly) "Upload-only release folder does not exist: $uploadOnly"

$expectedFiles = @(
  $packageFile,
  "SHA256SUMS.txt",
  "Hare_M365_Start_Windows.zip",
  "Hare_M365_Start_Mac_Linux.sh",
  "START_HERE.html",
  "LLM_FIRST_PROMPT_KO.txt",
  "README.md",
  "github-release-npm-guide.md"
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

foreach ($asset in Get-ChildItem -LiteralPath $uploadOnly -File | Where-Object { $_.Extension -in @(".md", ".txt", ".html", ".sh") }) {
  $content = Read-Utf8 $asset.FullName
  foreach ($placeholder in @("__PACKAGE_URL__", "__PACKAGE_FILE__", "__VERSION__", "__TAG__", "__REPO__")) {
    Assert-NotContains $content $placeholder $asset.Name
  }
}

$startHere = Read-Utf8 (Join-Path $uploadOnly "START_HERE.html")
$llmPreparesFolder = From-Base64Utf8 "6rCA64ql7ZWcIOuPhOq1rOuhnA=="
$userApprovesOnly = From-Base64Utf8 "7IKs7Jqp7J6Q6rCAIO2VoCDsnbzsnYAg7Jew6rKwIO2XiOyaqeu/kA=="

Assert-Order $startHere @(
  '<span class="num">1</span>',
  '<span class="num">2</span>',
  '<span class="num">3</span>',
  '<span class="num">4</span>',
  '<span class="num">5</span>',
  '<span class="num">6</span>'
) "START_HERE.html"
Assert-Contains $startHere "Claude/Cowork" "START_HERE.html"
Assert-Contains $startHere "github.com" "START_HERE.html"
Assert-Contains $startHere "release-assets.githubusercontent.com" "START_HERE.html"
Assert-Contains $startHere "registry.npmjs.org" "START_HERE.html"
Assert-Contains $startHere "graph.microsoft.com" "START_HERE.html"
Assert-Contains $startHere "login.microsoftonline.com" "START_HERE.html"
Assert-Contains $startHere "Documents" "START_HERE.html"
Assert-Contains $startHere "Hare M365 Agent" "START_HERE.html"
Assert-Contains $startHere "llm-guide, doctor, auth status" "START_HERE.html"
Assert-Contains $startHere "npm exec, npx, curl" "START_HERE.html"
Assert-Contains $startHere $llmPreparesFolder "START_HERE.html"
Assert-Contains $startHere $userApprovesOnly "START_HERE.html"
Assert-Order $startHere @("npm exec, npx, curl", "llm-guide, doctor, auth status", "Documents") "START_HERE.html gated folder order"

$firstPrompt = Read-Utf8 (Join-Path $uploadOnly "LLM_FIRST_PROMPT_KO.txt")
Assert-Contains $firstPrompt $packageUrl "LLM_FIRST_PROMPT_KO.txt"
Assert-Contains $firstPrompt "npm exec, npx, curl" "LLM_FIRST_PROMPT_KO.txt"
Assert-Contains $firstPrompt "llm-guide, doctor, auth status" "LLM_FIRST_PROMPT_KO.txt"
Assert-Contains $firstPrompt "Documents" "LLM_FIRST_PROMPT_KO.txt"
Assert-Contains $firstPrompt $llmPreparesFolder "LLM_FIRST_PROMPT_KO.txt"
Assert-Contains $firstPrompt $userApprovesOnly "LLM_FIRST_PROMPT_KO.txt"
Assert-Order $firstPrompt @("npm exec, npx, curl", "llm-guide, doctor, auth status", "Documents") "LLM_FIRST_PROMPT_KO.txt gated folder order"

$releaseGuide = Read-Utf8 (Join-Path $uploadOnly "github-release-npm-guide.md")
Assert-Contains $releaseGuide $packageUrl "github-release-npm-guide.md"
Assert-Contains $releaseGuide "llm-guide, doctor, auth status" "github-release-npm-guide.md"
Assert-Contains $releaseGuide "Documents" "github-release-npm-guide.md"
Assert-Contains $releaseGuide $llmPreparesFolder "github-release-npm-guide.md"

foreach ($assetName in @("START_HERE.html", "LLM_FIRST_PROMPT_KO.txt", "README.md", "github-release-npm-guide.md")) {
  $content = Read-Utf8 (Join-Path $uploadOnly $assetName)
  foreach ($forbidden in @(
    "C:\Users\OMH\Documents\Hare M365 Agent",
    "folder does not exist",
    "not found",
    "create the folder first",
    "File Explorer",
    "double-click execution",
    "__PACKAGE_",
    (From-Base64Utf8 "7IOIIO2PtOuNlA=="),
    (From-Base64Utf8 "66eM65Ok6rOgIOydtOumhA==")
  )) {
    Assert-NotContains $content $forbidden $assetName
  }
}

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
  $llmPromptOutput = (& node "dist\cli.js" "llm-prompt" | Out-String)
  Assert-Contains $llmPromptOutput $packageUrl "cli llm-prompt"
  Assert-Contains $llmPromptOutput "npm exec" "cli llm-prompt"
  Assert-Contains $llmPromptOutput "Documents" "cli llm-prompt"
  Assert-Contains $llmPromptOutput $llmPreparesFolder "cli llm-prompt"
  Assert-Order $llmPromptOutput @("npm exec", "llm-guide, doctor, auth status", "Documents") "cli llm-prompt gated folder order"

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
