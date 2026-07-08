Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

$packageJson = Get-Content -LiteralPath "package.json" -Raw | ConvertFrom-Json
$version = $packageJson.version

$releaseDir = Join-Path $root "releases"
$stageRoot = Join-Path $root "build\mac-build-input"
$stage = Join-Path $stageRoot "omh-m365-agent-mac-build-input"
$zipPath = Join-Path $releaseDir "OMH-M365-Agent-mac-build-input-$version.zip"
$shaPath = "$zipPath.sha256"

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

function Copy-RequiredDirectory {
  param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Destination
  )

  if (-not (Test-Path -LiteralPath $Source)) {
    throw "Required directory is missing: $Source"
  }

  New-Item -ItemType Directory -Path (Split-Path -Parent $Destination) -Force | Out-Null
  Copy-Item -LiteralPath $Source -Destination $Destination -Recurse -Force
}

function Test-StagedInput {
  $macPkgBuildGuideName = ConvertFrom-Base64Utf8 "TUFDX1BLR1/ruYzrk5xf6rCA7J2065OcLm1k"
  $macDevicePromptName = ConvertFrom-Base64Utf8 "TUFDX+yLpOq4sOq4sF/qsoDspp1f7ZSE66Gs7ZSE7Yq4LnR4dA=="

  foreach ($required in @(
    "package.json",
    ".env",
    "policy.json",
    "releases\mac-llm\omh-m365.cjs",
    "releases\mac-llm\runtime\node-v24.15.0-darwin-arm64.tar.gz",
    "releases\mac-llm\runtime\node-v24.15.0-darwin-x64.tar.gz",
    "installer-assets\mac\start-guide-template.html",
    "scripts\build-mac-pkg.sh",
    "scripts\verify-mac-install.sh",
    "docs\mac-pkg-deployment-guide.md",
    (Join-Path "releases" $macPkgBuildGuideName),
    (Join-Path "releases" $macDevicePromptName)
  )) {
    if (-not (Test-Path -LiteralPath (Join-Path $stage $required))) {
      throw "Required build-input item is missing: $required"
    }
  }

  foreach ($blocked in @(".cache", "downloads", "logs", "node_modules", "dist")) {
    $hit = Get-ChildItem -LiteralPath $stage -Force -Recurse | Where-Object { $_.Name -eq $blocked } | Select-Object -First 1
    if ($hit) {
      throw "Blocked item was staged unexpectedly: $blocked"
    }
  }
}

New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null

if (Test-Path -LiteralPath $stageRoot) {
  Remove-Item -LiteralPath $stageRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $stage -Force | Out-Null

foreach ($file in @(
  "package.json",
  "package-lock.json",
  "tsconfig.json",
  ".env",
  ".env.example",
  "policy.json",
  "README.md"
)) {
  Copy-RequiredFile -Source (Join-Path $root $file) -Destination (Join-Path $stage $file)
}

Copy-RequiredDirectory -Source (Join-Path $root "src") -Destination (Join-Path $stage "src")
Copy-RequiredDirectory -Source (Join-Path $root "installer-assets\mac") -Destination (Join-Path $stage "installer-assets\mac")
Copy-RequiredDirectory -Source (Join-Path $root "releases\mac-llm") -Destination (Join-Path $stage "releases\mac-llm")

foreach ($file in @(
  "build-mac-pkg.sh",
  "verify-mac-install.sh"
)) {
  Copy-RequiredFile -Source (Join-Path $root "scripts\$file") -Destination (Join-Path $stage "scripts\$file")
}

New-Item -ItemType Directory -Path (Join-Path $stage "docs") -Force | Out-Null
Copy-RequiredFile -Source (Join-Path $root "docs\mac-pkg-deployment-guide.md") -Destination (Join-Path $stage "docs\mac-pkg-deployment-guide.md")

New-Item -ItemType Directory -Path (Join-Path $stage "releases") -Force | Out-Null
foreach ($file in @(
  (ConvertFrom-Base64Utf8 "TUFDX1BLR1/ruYzrk5xf6rCA7J2065OcLm1k"),
  (ConvertFrom-Base64Utf8 "TUFDX+yLpOq4sOq4sF/qsoDspp1f7ZSE66Gs7ZSE7Yq4LnR4dA=="),
  (ConvertFrom-Base64Utf8 "TUFDX1BLR1/sg53shLFf7ZWE7JqULnR4dA==")
)) {
  Copy-RequiredFile -Source (Join-Path $root "releases\$file") -Destination (Join-Path $stage "releases\$file")
}

Test-StagedInput

foreach ($path in @($zipPath, $shaPath)) {
  if (Test-Path -LiteralPath $path) {
    Remove-Item -LiteralPath $path -Force
  }
}

Compress-Archive -LiteralPath $stage -DestinationPath $zipPath -CompressionLevel Optimal

$hash = Get-FileHash -LiteralPath $zipPath -Algorithm SHA256
Set-Content -LiteralPath $shaPath -Value "$($hash.Hash.ToLowerInvariant())  $(Split-Path -Leaf $zipPath)" -Encoding ASCII

Write-Host "Created Mac PKG build input:"
Write-Host $zipPath
Write-Host "SHA256:"
Write-Host $hash.Hash.ToLowerInvariant()
