Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

function Invoke-CliJson {
  param([string[]]$CliArgs)

  $outFile = New-TemporaryFile
  $errFile = New-TemporaryFile
  try {
    $proc = Start-Process -FilePath "node" `
      -ArgumentList (@("dist/cli.js") + $CliArgs) `
      -WorkingDirectory $root `
      -NoNewWindow `
      -Wait `
      -PassThru `
      -RedirectStandardOutput $outFile.FullName `
      -RedirectStandardError $errFile.FullName

    $text = ((Get-Content -LiteralPath $outFile.FullName -Raw -Encoding UTF8) + (Get-Content -LiteralPath $errFile.FullName -Raw -Encoding UTF8)).Trim()
    if (-not $text) {
      return [pscustomobject]@{ ok = $false; error = "No CLI output."; exitCode = $proc.ExitCode }
    }

    $json = $text | ConvertFrom-Json
    $hasOk = $json.PSObject.Properties.Name -contains "ok"
    if ($hasOk -and $json.ok -eq $false) {
      return [pscustomobject]@{ ok = $false; error = $json.error; exitCode = $proc.ExitCode }
    }

    return [pscustomobject]@{ ok = ($proc.ExitCode -eq 0); json = $json; exitCode = $proc.ExitCode }
  }
  finally {
    Remove-Item -LiteralPath $outFile.FullName, $errFile.FullName -Force -ErrorAction SilentlyContinue
  }
}

function Get-CountFromJson {
  param($Json)

  $properties = $Json.PSObject.Properties.Name
  if ($properties -contains "messages") { return @($Json.messages).Count }
  if ($properties -contains "teams") { return @($Json.teams).Count }
  if ($properties -contains "chats") { return @($Json.chats).Count }
  if ($properties -contains "files") { return @($Json.files).Count }
  return $null
}

$checks = @(
  @{ name = "doctor"; args = @("doctor") },
  @{ name = "auth status"; args = @("auth", "status") },
  @{ name = "outlook inbox"; args = @("outlook", "inbox", "--limit", "3") },
  @{ name = "teams teams"; args = @("teams", "teams") },
  @{ name = "teams chats"; args = @("teams", "chats", "--limit", "3") },
  @{ name = "files search"; args = @("files", "search", "--query", "test", "--limit", "3") }
)

$results = foreach ($check in $checks) {
  $result = Invoke-CliJson -CliArgs $check.args
  $loggedIn = $null
  $count = $null
  $errorMessage = $null
  if ($result.json) {
    if ($result.json.PSObject.Properties.Name -contains "loggedIn") { $loggedIn = [bool]$result.json.loggedIn }
    $count = Get-CountFromJson -Json $result.json
  }
  if ($result.PSObject.Properties.Name -contains "error") {
    $errorMessage = $result.error
  }

  [pscustomobject]@{
    name = $check.name
    ok = [bool]$result.ok
    loggedIn = $loggedIn
    count = $count
    error = $errorMessage
  }
}

$results | ConvertTo-Json -Depth 4

if ($results | Where-Object { -not $_.ok }) {
  exit 1
}
