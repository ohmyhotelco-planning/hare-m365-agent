Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

$packageJson = Get-Content -LiteralPath "package.json" -Raw | ConvertFrom-Json
$version = $packageJson.version
$releaseDir = Join-Path $root "releases"
$target = Join-Path $releaseDir "sharepoint-upload"

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

if (Test-Path -LiteralPath $target) {
  Remove-Item -LiteralPath $target -Recurse -Force
}

$windowsDir = Join-Path $target "Windows"
$npmDir = Join-Path $target "NPM_CLI"
$macBuildDir = Join-Path $target "Mac_PKG_Build_Input"
$docsDir = Join-Path $target "Docs"
New-Item -ItemType Directory -Path $windowsDir, $npmDir, $macBuildDir, $docsDir -Force | Out-Null

$installer = "OMH-M365-Agent-Setup-$version.exe"
Copy-RequiredFile -Source (Join-Path $releaseDir $installer) -Destination (Join-Path $windowsDir $installer)
Copy-RequiredFile -Source (Join-Path $releaseDir "$installer.sha256") -Destination (Join-Path $windowsDir "$installer.sha256")

$npmPackage = "ohmyhotel-hare-m365-agent-$version.tgz"
Copy-RequiredFile -Source (Join-Path $releaseDir "npm\$npmPackage") -Destination (Join-Path $npmDir $npmPackage)
Copy-RequiredFile -Source (Join-Path $releaseDir "npm\$npmPackage.sha256") -Destination (Join-Path $npmDir "$npmPackage.sha256")

$macBuildInput = "OMH-M365-Agent-mac-build-input-$version.zip"
Copy-RequiredFile -Source (Join-Path $releaseDir $macBuildInput) -Destination (Join-Path $macBuildDir $macBuildInput)
Copy-RequiredFile -Source (Join-Path $releaseDir "$macBuildInput.sha256") -Destination (Join-Path $macBuildDir "$macBuildInput.sha256")

foreach ($file in @(
  (ConvertFrom-Base64Utf8 "TUFDX1BLR1/ruYzrk5xf6rCA7J2065OcLm1k"),
  (ConvertFrom-Base64Utf8 "TUFDX+yLpOq4sOq4sF/qsoDspp1f7ZSE66Gs7ZSE7Yq4LnR4dA=="),
  (ConvertFrom-Base64Utf8 "TUFDX1BLR1/sg53shLFf7ZWE7JqULnR4dA=="),
  (ConvertFrom-Base64Utf8 "VEVBTVNf67Cw7Y+sX+uplOyLnOyngF/thZztlIzrpr8udHh0")
)) {
  Copy-RequiredFile -Source (Join-Path $releaseDir $file) -Destination (Join-Path $docsDir $file)
}

Copy-RequiredFile -Source (Join-Path $root "docs\distribution-channel-guide.md") -Destination (Join-Path $docsDir "distribution-channel-guide.md")
Copy-RequiredFile -Source (Join-Path $root "docs\mac-pkg-deployment-guide.md") -Destination (Join-Path $docsDir "mac-pkg-deployment-guide.md")
Copy-RequiredFile -Source (Join-Path $root "docs\npm-cli-guide.md") -Destination (Join-Path $docsDir "npm-cli-guide.md")

$readme = ConvertFrom-Base64Utf8 "IyBIYXJlIE0zNjUgQWdlbnQgU2hhcmVQb2ludCBVcGxvYWQKCuydtCDtj7TrjZQg7KCE7LK066W8IFNoYXJlUG9pbnQg67Cw7Y+sIOychOy5mOyXkCDsl4XroZzrk5ztlanri4jri6QuCgojIyDtj7TrjZQg6rWs7ISxCgotIE5QTV9DTEk6IFdpbmRvd3MvTWFjL0xpbnV4L0xMTSDtmZjqsr0g6rO17Ya1IOq2jOyepSDrsLDtj6wg7YyM7J287J6F64uI64ukLgotIFdpbmRvd3M6IFdpbmRvd3Mg7IKs7Jqp7J6Q7JqpIOyEpOy5mCDtjIzsnbzsnoXri4jri6QuCi0gTWFjX1BLR19CdWlsZF9JbnB1dDogTWFj7JeQ7IScIC5wa2frpbwg7IOd7ISx7ZWY6riwIOychO2VnCDsnoXroKXrrLzsnoXri4jri6QuIOy1nOyihSDsgqzsmqnsnpDsmqkgTWFjIHBrZ+qwgCDslYTri5nri4jri6QuCi0gRG9jczog67Cw7Y+sIOyViOuCtCwgVGVhbXMg66mU7Iuc7KeAIO2FnO2UjOumvywgbnBtIENMSS9NYWMg67mM65OcL+qygOymnSDslYjrgrTsnoXri4jri6QuCgojIyDqtozsnqUg67Cw7Y+sOiBucG0gQ0xJCgpMTE3snbQgc2hlbGzsnYQg7Iuk7ZaJ7ZWgIOyImCDsnojripQg7ZmY6rK97JeQ7ISc64qUIG5wbSBDTEnrpbwg7Jqw7ISgIOyCrOyaqe2VqeuLiOuLpC4KCk5QTV9DTEkvb2hteWhvdGVsLWhhcmUtbTM2NS1hZ2VudC0wLjEuMC50Z3oKCuyCrOyaqSDsmIg6CgpgYGBiYXNoCm5wbSBpbnN0YWxsIC1nIC4vb2hteWhvdGVsLWhhcmUtbTM2NS1hZ2VudC0wLjEuMC50Z3oKaGFyZS1tMzY1IGxsbS1ndWlkZQpoYXJlLW0zNjUgZG9jdG9yCmBgYAoK7IKs64K0IG5wbSByZWdpc3RyeeyXkCBwdWJsaXNo7ZWcIOuSpOyXkOuKlCDri6TsnYwg67Cp7Iud7Jy866GcIOyCrOyaqe2VqeuLiOuLpC4KCmBgYGJhc2gKbnB4IEBvaG15aG90ZWwvaGFyZS1tMzY1LWFnZW50IGxsbS1ndWlkZQpgYGAKCiMjIFdpbmRvd3Mg67O07KGwIOuwsO2PrAoKV2luZG93cyDsgqzsmqnsnpDripQg7JWE656YIO2MjOydvOydhCDri6TsmrTroZzrk5ztlbQg7Iuk7ZaJ7ZWp64uI64ukLgoKV2luZG93cy9PTUgtTTM2NS1BZ2VudC1TZXR1cC0wLjEuMC5leGUKCiMjIE1hYyBwa2cg7IOd7ISxCgrtmITsnqwg7J20IO2PtOuNlOyXkOuKlCDstZzsooUgTWFjIC5wa2fqsIAg7JeG7Iq164uI64ukLgoKTWFjIOyLpOq4sOq4sCDrmJDripQgbWFjT1MgQ0kgcnVubmVy7JeQ7IScIOyVhOuemCDsnoXroKXrrLzsnYQg67Cb7JWEIC5wa2frpbwg7IOd7ISx7ZWp64uI64ukLgoKTWFjX1BLR19CdWlsZF9JbnB1dC9PTUgtTTM2NS1BZ2VudC1tYWMtYnVpbGQtaW5wdXQtMC4xLjAuemlwCgpNYWPsl5DshJwg7IOd7ISxIO2bhCDstZzsooUg67Cw7Y+sIO2PtOuNlOyXkCDstpTqsIDtlaAg7YyM7J28OgoKLSBNYWMvT01ILU0zNjUtQWdlbnQtbWFjLTAuMS4wLnBrZwotIE1hYy9PTUgtTTM2NS1BZ2VudC1tYWMtMC4xLjAucGtnLnNoYTI1NgoKIyMg7KO87J2YCgotIFRlYW1z7JeQ64qUIOyLpO2WiSDtjIzsnbzsnYQg7KeB7KCRIOyyqOu2gO2VmOyngCDslYrqs6AgU2hhcmVQb2ludCDrp4Htgazrp4wg6rO17Jyg7ZWp64uI64ukLgotIC5jYWNoZSwgdG9rZW4sIGRldmljZSBjb2Rl64qUIO2PrO2VqO2VmOyngCDslYrsirXri4jri6QuCi0gLmVuduyXkOuKlCBkZWxlZ2F0ZWQgcHVibGljLWNsaWVudCDshKTsoJXsnbQg7Y+s7ZWo65CY7Ja0IOyeiOyKteuLiOuLpC4K"
$readme = $readme.Replace("0.1.0", $version)
Set-Content -LiteralPath (Join-Path $target "README.md") -Value $readme -Encoding UTF8

$manifest = ConvertFrom-Base64Utf8 "SGFyZSBNMzY1IEFnZW50IFNoYXJlUG9pbnQgVXBsb2FkIE1hbmlmZXN0ClZlcnNpb246IDAuMS4wCgpOUE0gQ0xJOgotIE5QTV9DTEkvb2hteWhvdGVsLWhhcmUtbTM2NS1hZ2VudC0wLjEuMC50Z3oKLSBOUE1fQ0xJL29obXlob3RlbC1oYXJlLW0zNjUtYWdlbnQtMC4xLjAudGd6LnNoYTI1NgoKV2luZG93czoKLSBXaW5kb3dzL09NSC1NMzY1LUFnZW50LVNldHVwLTAuMS4wLmV4ZQotIFdpbmRvd3MvT01ILU0zNjUtQWdlbnQtU2V0dXAtMC4xLjAuZXhlLnNoYTI1NgoKTWFjIGJ1aWxkIGlucHV0OgotIE1hY19QS0dfQnVpbGRfSW5wdXQvT01ILU0zNjUtQWdlbnQtbWFjLWJ1aWxkLWlucHV0LTAuMS4wLnppcAotIE1hY19QS0dfQnVpbGRfSW5wdXQvT01ILU0zNjUtQWdlbnQtbWFjLWJ1aWxkLWlucHV0LTAuMS4wLnppcC5zaGEyNTYKCkRvY3M6Ci0gRG9jcy9ucG0tY2xpLWd1aWRlLm1kCi0gRG9jcy9NQUNfUEtHX+u5jOuTnF/qsIDsnbTrk5wubWQKLSBEb2NzL01BQ1/si6TquLDquLBf6rKA7KadX+2UhOuhrO2UhO2KuC50eHQKLSBEb2NzL01BQ19QS0df7IOd7ISxX+2VhOyalC50eHQKLSBEb2NzL1RFQU1TX+uwsO2PrF/rqZTsi5zsp4Bf7YWc7ZSM66a/LnR4dAotIERvY3MvZGlzdHJpYnV0aW9uLWNoYW5uZWwtZ3VpZGUubWQKLSBEb2NzL21hYy1wa2ctZGVwbG95bWVudC1ndWlkZS5tZAo="
$manifest = $manifest.Replace("0.1.0", $version)
Set-Content -LiteralPath (Join-Path $target "MANIFEST.txt") -Value $manifest -Encoding UTF8

$blocked = @(".cache", "downloads", "logs", "node_modules", "dist", "src", "scripts")
foreach ($name in $blocked) {
  $hit = Get-ChildItem -LiteralPath $target -Force -Recurse | Where-Object { $_.Name -eq $name } | Select-Object -First 1
  if ($hit) {
    throw "Blocked item was staged unexpectedly: $name"
  }
}

Write-Host "Prepared SharePoint upload folder:"
Write-Host $target
