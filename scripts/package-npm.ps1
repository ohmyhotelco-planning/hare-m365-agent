Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

$packageJson = Get-Content -LiteralPath "package.json" -Raw | ConvertFrom-Json
$version = $packageJson.version
$releaseDir = Join-Path $root "releases\npm"

if (Test-Path -LiteralPath $releaseDir) {
  Remove-Item -LiteralPath $releaseDir -Recurse -Force
}
New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null

npm run typecheck
npm run build

npm pack --pack-destination $releaseDir

$tgz = Get-ChildItem -LiteralPath $releaseDir -Filter "*.tgz" | Select-Object -First 1
if (-not $tgz) {
  throw "npm pack did not create a .tgz file."
}

$entries = tar -tf $tgz.FullName

foreach ($required in @(
  "package/package.json",
  "package/dist/cli.js",
  "package/dist/config.js",
  "package/docs/npm-cli-guide.md",
  "package/docs/github-release-npm-guide.md",
  "package/policy.json",
  "package/.env",
  "package/README.md"
)) {
  if ($entries -notcontains $required) {
    throw "Required npm package entry is missing: $required"
  }
}

foreach ($forbidden in @(
  "package/.cache/",
  "package/runtime/",
  "package/downloads/",
  "package/logs/",
  "package/node_modules/",
  "package/src/",
  "package/releases/",
  "package/scripts/",
  "package/build/"
)) {
  if ($entries | Where-Object { $_ -like "$forbidden*" }) {
    throw "Forbidden npm package entry found: $forbidden"
  }
}

$hash = Get-FileHash -LiteralPath $tgz.FullName -Algorithm SHA256
Set-Content -LiteralPath "$($tgz.FullName).sha256" -Value "$($hash.Hash.ToLowerInvariant())  $($tgz.Name)" -Encoding ASCII

Write-Host "Created npm package:"
Write-Host $tgz.FullName
Write-Host "SHA256:"
Write-Host $hash.Hash.ToLowerInvariant()
