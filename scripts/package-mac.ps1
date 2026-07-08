Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

$packageJson = Get-Content -LiteralPath "package.json" -Raw | ConvertFrom-Json
$version = $packageJson.version
$nodeVersion = (& node -p "process.version.slice(1)").Trim()

$releaseDir = Join-Path $root "releases"
$stage = Join-Path $releaseDir "mac-llm"
$assetDir = Join-Path $root "installer-assets\mac"
$nodeCacheDir = Join-Path $root "build\node-cache\v$nodeVersion"
$archiveBaseName = "OMH-M365-Agent-mac"
$tarGzPath = Join-Path $releaseDir "$archiveBaseName-$version.tar.gz"
$zipPath = Join-Path $releaseDir "$archiveBaseName-$version.zip"

function Copy-RequiredFile {
  param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Destination
  )

  if (-not (Test-Path -LiteralPath $Source)) {
    throw "Required file is missing: $Source"
  }

  $parent = Split-Path -Parent $Destination
  New-Item -ItemType Directory -Path $parent -Force | Out-Null
  Copy-Item -LiteralPath $Source -Destination $Destination -Force
}

function ConvertFrom-Base64Utf8 {
  param([Parameter(Mandatory = $true)][string]$Value)
  return [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Value))
}

function ConvertTo-HtmlText {
  param([Parameter(Mandatory = $true)][string]$Value)
  return [System.Net.WebUtility]::HtmlEncode($Value)
}

function Write-MacStartFiles {
  $startHtmlName = ConvertFrom-Base64Utf8 "7Iuc7J6R7ZWY6riwLmh0bWw="
  $promptFileName = ConvertFrom-Base64Utf8 "Q293b3Jr7JeQIOu2meyXrOuEo+ydhCDtlITroaztlITtirgudHh0"
  $promptText = ConvertFrom-Base64Utf8 "7J20IE1hYyDrsLDtj6wg7Y+0642U66W8IENvd29ya+yXkCDsl7DqsrDtlZjqs6AgQ09XT1JLX1JFUVVJUkVEX1JFQURNRS5tZOulvCDrqLzsoIAg7J297J2AIOuSpCDqt7jrjIDroZwg65Sw6528LgrsnpHsl4U6IDzsm5DtlZjripQgTWljcm9zb2Z0IDM2NSDsnpHsl4U+"
  $loginHint = ConvertFrom-Base64Utf8 "7J20IO2PtOuNlOydmCDroZzqt7jsnbgg7Iuc7J6RLmNvbW1hbmQ="
  $promptHint = ConvertFrom-Base64Utf8 "7J20IO2PtOuNlOydmCBDb3dvcmVsl5Ag67aZ7Jes64Sj7J2EIO2UhOuhrO2UhO2KuC50eHQ="
  $folderHint = ConvertFrom-Base64Utf8 "7J20IEhUTUwg7YyM7J287J20IOyeiOuKlCBNYWMg67Cw7Y+sIO2PtOuNlA=="

  Set-Content -LiteralPath (Join-Path $stage $promptFileName) -Value $promptText -Encoding UTF8
  Set-Content -LiteralPath (Join-Path $stage "COWORK_FIRST_PROMPT.txt") -Value $promptText -Encoding UTF8

  $templatePath = Join-Path $assetDir "start-guide-template.html"
  if (-not (Test-Path -LiteralPath $templatePath)) {
    throw "Mac start guide template is missing: $templatePath"
  }

  $html = Get-Content -LiteralPath $templatePath -Raw -Encoding UTF8
  $html = $html.
    Replace("__ESCAPED_LOGIN_CMD__", (ConvertTo-HtmlText $loginHint)).
    Replace("__ESCAPED_PROMPT__", (ConvertTo-HtmlText $promptText)).
    Replace("__ESCAPED_PROMPT_PATH__", (ConvertTo-HtmlText $promptHint)).
    Replace("__ESCAPED_LINUX_PATH__", (ConvertTo-HtmlText $folderHint))

  Set-Content -LiteralPath (Join-Path $stage $startHtmlName) -Value $html -Encoding UTF8
}

function Get-NodeRuntime {
  param(
    [Parameter(Mandatory = $true)][string]$RuntimeName
  )

  New-Item -ItemType Directory -Path $nodeCacheDir -Force | Out-Null

  $shasumsPath = Join-Path $nodeCacheDir "SHASUMS256.txt"
  $shasumsUrl = "https://nodejs.org/dist/v$nodeVersion/SHASUMS256.txt"

  if (-not (Test-Path -LiteralPath $shasumsPath)) {
    Invoke-WebRequest -Uri $shasumsUrl -OutFile $shasumsPath -UseBasicParsing
  }

  $archiveName = "$RuntimeName.tar.gz"
  $archivePath = Join-Path $nodeCacheDir $archiveName
  $archiveUrl = "https://nodejs.org/dist/v$nodeVersion/$archiveName"

  if (-not (Test-Path -LiteralPath $archivePath)) {
    Invoke-WebRequest -Uri $archiveUrl -OutFile $archivePath -UseBasicParsing
  }

  $escapedArchiveName = [regex]::Escape($archiveName)
  $line = Get-Content -LiteralPath $shasumsPath | Where-Object { $_ -match "\s+$escapedArchiveName$" } | Select-Object -First 1
  if (-not $line) {
    throw "Cannot find checksum for $archiveName in Node SHASUMS256.txt"
  }

  $expected = ($line -split "\s+")[0].ToLowerInvariant()
  $actual = (Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash.ToLowerInvariant()
  if ($expected -ne $actual) {
    throw "Checksum mismatch for $archiveName"
  }

  Copy-RequiredFile -Source $archivePath -Destination (Join-Path $stage "runtime\$archiveName")
}

function New-Archives {
  $python = @'
import os
import stat
import tarfile
import zipfile
from pathlib import Path

stage = Path(os.environ["OMH_M365_STAGE"])
base = os.environ["OMH_M365_ARCHIVE_BASE"]
tar_gz = Path(os.environ["OMH_M365_TAR_GZ"])
zip_path = Path(os.environ["OMH_M365_ZIP"])

def mode_for(path: Path) -> int:
    if path.is_dir():
        return 0o755
    if path.suffix in {".sh", ".command"}:
        return 0o755
    return 0o644

def iter_paths():
    yield stage
    for path in sorted(stage.rglob("*"), key=lambda p: p.as_posix().lower()):
        yield path

def arcname(path: Path) -> str:
    if path == stage:
        return base
    return f"{base}/{path.relative_to(stage).as_posix()}"

with tarfile.open(tar_gz, "w:gz") as archive:
    for path in iter_paths():
        info = archive.gettarinfo(str(path), arcname(path))
        info.mode = mode_for(path)
        if path.is_file():
            with path.open("rb") as handle:
                archive.addfile(info, handle)
        else:
            archive.addfile(info)

with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
    for path in iter_paths():
        name = arcname(path)
        if path.is_dir():
            if not name.endswith("/"):
                name = f"{name}/"
            info = zipfile.ZipInfo(name)
            mode = mode_for(path) | stat.S_IFDIR
            info.external_attr = mode << 16
            archive.writestr(info, b"")
        else:
            info = zipfile.ZipInfo(name)
            mode = mode_for(path) | stat.S_IFREG
            info.external_attr = mode << 16
            with path.open("rb") as handle:
                archive.writestr(info, handle.read())
'@

  $env:OMH_M365_STAGE = $stage
  $env:OMH_M365_ARCHIVE_BASE = $archiveBaseName
  $env:OMH_M365_TAR_GZ = $tarGzPath
  $env:OMH_M365_ZIP = $zipPath
  try {
    $python | python -
  } finally {
    Remove-Item Env:\OMH_M365_STAGE -ErrorAction SilentlyContinue
    Remove-Item Env:\OMH_M365_ARCHIVE_BASE -ErrorAction SilentlyContinue
    Remove-Item Env:\OMH_M365_TAR_GZ -ErrorAction SilentlyContinue
    Remove-Item Env:\OMH_M365_ZIP -ErrorAction SilentlyContinue
  }
}

function Test-Stage {
  $startHtmlName = ConvertFrom-Base64Utf8 "7Iuc7J6R7ZWY6riwLmh0bWw="
  $promptFileName = ConvertFrom-Base64Utf8 "Q293b3Jr7JeQIOu2meyXrOuEo+ydhCDtlITroaztlITtirgudHh0"
  $loginCommandName = ConvertFrom-Base64Utf8 "66Gc6re47J24IOyLnOyekS5jb21tYW5k"

  foreach ($required in @(
    ".env",
    "policy.json",
    "README.md",
    "COWORK_REQUIRED_README.md",
    $startHtmlName,
    $promptFileName,
    "COWORK_FIRST_PROMPT.txt",
    "run-mac.sh",
    "run-cowork.sh",
    $loginCommandName,
    "omh-m365.cjs",
    "omh-m365.cjs.sha256",
    "runtime\node-v$nodeVersion-darwin-arm64.tar.gz",
    "runtime\node-v$nodeVersion-darwin-x64.tar.gz"
  )) {
    if (-not (Test-Path -LiteralPath (Join-Path $stage $required))) {
      throw "Required Mac package item is missing: $required"
    }
  }

  foreach ($blocked in @(".cache", "downloads", "logs", "node_modules", "src", "dist", "scripts", "build")) {
    $hit = Get-ChildItem -LiteralPath $stage -Force -Recurse | Where-Object { $_.Name -eq $blocked } | Select-Object -First 1
    if ($hit) {
      throw "Blocked item was staged unexpectedly: $blocked"
    }
  }
}

New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null

if (-not (Test-Path -LiteralPath $assetDir)) {
  throw "Mac asset directory is missing: $assetDir"
}

if (Test-Path -LiteralPath $stage) {
  Remove-Item -LiteralPath $stage -Recurse -Force
}
New-Item -ItemType Directory -Path (Join-Path $stage "runtime") -Force | Out-Null

npm run typecheck

npx esbuild src\cli.ts `
  --bundle `
  --platform=node `
  --target=node20 `
  --format=cjs `
  --outfile="$(Join-Path $stage 'omh-m365.cjs')"

$bundlePath = Join-Path $stage "omh-m365.cjs"
node --check $bundlePath
$bundleHash = Get-FileHash -LiteralPath $bundlePath -Algorithm SHA256
Set-Content -LiteralPath (Join-Path $stage "omh-m365.cjs.sha256") -Value "$($bundleHash.Hash.ToLowerInvariant())  omh-m365.cjs" -Encoding ASCII

Copy-RequiredFile -Source (Join-Path $root ".env") -Destination (Join-Path $stage ".env")
Copy-RequiredFile -Source (Join-Path $root "policy.json") -Destination (Join-Path $stage "policy.json")

foreach ($file in @(
  "README.md",
  "COWORK_REQUIRED_README.md",
  "run-mac.sh",
  "run-cowork.sh",
  (ConvertFrom-Base64Utf8 "66Gc6re47J24IOyLnOyekS5jb21tYW5k")
)) {
  Copy-RequiredFile -Source (Join-Path $assetDir $file) -Destination (Join-Path $stage $file)
}

$runMacPath = Join-Path $stage "run-mac.sh"
$runMac = Get-Content -LiteralPath $runMacPath -Raw -Encoding UTF8
$runMac = $runMac -replace 'NODE_VERSION="[0-9]+\.[0-9]+\.[0-9]+"', "NODE_VERSION=""$nodeVersion"""
Set-Content -LiteralPath $runMacPath -Value $runMac -Encoding UTF8

Write-MacStartFiles

Get-NodeRuntime -RuntimeName "node-v$nodeVersion-darwin-arm64"
Get-NodeRuntime -RuntimeName "node-v$nodeVersion-darwin-x64"

Test-Stage

foreach ($path in @($tarGzPath, $zipPath, "$tarGzPath.sha256", "$zipPath.sha256")) {
  if (Test-Path -LiteralPath $path) {
    Remove-Item -LiteralPath $path -Force
  }
}

New-Archives

foreach ($archive in @($tarGzPath, $zipPath)) {
  if (-not (Test-Path -LiteralPath $archive)) {
    throw "Archive was not created: $archive"
  }
  $hash = Get-FileHash -LiteralPath $archive -Algorithm SHA256
  Set-Content -LiteralPath "$archive.sha256" -Value "$($hash.Hash.ToLowerInvariant())  $(Split-Path -Leaf $archive)" -Encoding ASCII
}

Write-Host "Created Mac release folder:"
Write-Host $stage
Write-Host "Created Mac archives:"
Write-Host $tarGzPath
Write-Host $zipPath
