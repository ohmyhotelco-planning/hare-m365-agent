Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

$packageJson = Get-Content -LiteralPath "package.json" -Raw | ConvertFrom-Json
$version = $packageJson.version
$repo = if ($env:HARE_M365_GITHUB_REPO) { $env:HARE_M365_GITHUB_REPO } else { "ohmyhotelco-planning/hare-m365-agent" }
$tag = if ($env:HARE_M365_RELEASE_TAG) { $env:HARE_M365_RELEASE_TAG } else { "v$version" }

npm run package:npm

$npmReleaseDir = Join-Path $root "releases\npm"
$tgz = Get-ChildItem -LiteralPath $npmReleaseDir -Filter "*.tgz" | Select-Object -First 1
if (-not $tgz) {
  throw "npm package was not found under releases\npm."
}

$sha = "$($tgz.FullName).sha256"
if (-not (Test-Path -LiteralPath $sha)) {
  throw "SHA256 file was not found: $sha"
}

$githubReleaseRoot = Join-Path $root "releases\github-release"
$target = Join-Path $githubReleaseRoot $tag
$uploadOnly = Join-Path $githubReleaseRoot "$tag-upload-only"
$resolvedReleaseRoot = [System.IO.Path]::GetFullPath($githubReleaseRoot)
$resolvedTarget = [System.IO.Path]::GetFullPath($target)
$resolvedUploadOnly = [System.IO.Path]::GetFullPath($uploadOnly)
if (-not $resolvedTarget.StartsWith($resolvedReleaseRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
  throw "Refusing to write outside releases\github-release: $resolvedTarget"
}
if (-not $resolvedUploadOnly.StartsWith($resolvedReleaseRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
  throw "Refusing to write outside releases\github-release: $resolvedUploadOnly"
}

if (Test-Path -LiteralPath $target) {
  Remove-Item -LiteralPath $target -Recurse -Force
}
if (Test-Path -LiteralPath $uploadOnly) {
  Remove-Item -LiteralPath $uploadOnly -Recurse -Force
}
New-Item -ItemType Directory -Path $target -Force | Out-Null
New-Item -ItemType Directory -Path $uploadOnly -Force | Out-Null

Copy-Item -LiteralPath $tgz.FullName -Destination (Join-Path $target $tgz.Name) -Force
$shaText = Get-Content -LiteralPath $sha -Raw
Set-Content -LiteralPath (Join-Path $target "SHA256SUMS.txt") -Value $shaText -Encoding ASCII
$packageUrl = "https://github.com/$repo/releases/download/$tag/$($tgz.Name)"

$required = @(
  $tgz.Name,
  "SHA256SUMS.txt"
)

foreach ($name in $required) {
  if (-not (Test-Path -LiteralPath (Join-Path $target $name))) {
    throw "Required GitHub Release asset was not created: $name"
  }
  Copy-Item -LiteralPath (Join-Path $target $name) -Destination (Join-Path $uploadOnly $name) -Force
}

& (Join-Path $PSScriptRoot "validate-release.ps1") -Tag $tag

Write-Host "Prepared GitHub Release assets:"
Write-Host $target
Write-Host "Upload only these files:"
Write-Host $uploadOnly
Write-Host "Package URL:"
Write-Host $packageUrl
