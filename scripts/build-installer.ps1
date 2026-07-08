Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $root

$packageJson = Get-Content -LiteralPath "package.json" -Raw | ConvertFrom-Json
$version = $packageJson.version
$releaseDir = Join-Path $root "releases"
$linuxDir = Join-Path $releaseDir "linux-llm"
$winDir = Join-Path $releaseDir "win-x64"
$buildDir = Join-Path $root "build\installer"
$payloadDir = Join-Path $buildDir "payload"
$payloadZip = Join-Path $buildDir "payload.zip"
$installerPath = Join-Path $releaseDir "OMH-M365-Agent-Setup-$version.exe"
$guideTemplatePath = Join-Path $root "installer-assets\start-guide-template.html"

function Copy-RequiredFile {
  param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Destination
  )

  if (-not (Test-Path -LiteralPath $Source)) {
    throw "Required installer source file is missing: $Source"
  }

  $parent = Split-Path -Parent $Destination
  New-Item -ItemType Directory -Path $parent -Force | Out-Null
  Copy-Item -LiteralPath $Source -Destination $Destination -Force
}

function Sync-LinuxRelease {
  New-Item -ItemType Directory -Path $linuxDir -Force | Out-Null

  npm run typecheck

  npx esbuild src\cli.ts `
    --bundle `
    --platform=node `
    --target=node20 `
    --format=cjs `
    --outfile="$(Join-Path $linuxDir 'omh-m365.cjs')"

  $bundlePath = Join-Path $linuxDir "omh-m365.cjs"
  node --check $bundlePath
  $bundleHash = Get-FileHash -LiteralPath $bundlePath -Algorithm SHA256
  Set-Content -LiteralPath (Join-Path $linuxDir "omh-m365.cjs.sha256") -Value "$($bundleHash.Hash.ToLower())  omh-m365.cjs" -Encoding ASCII

  foreach ($file in @(
    "README.md",
    "START_HERE_FOR_LLM.md",
    "AGENTS.md",
    "CLAUDE.md",
    "COWORK_REQUIRED_README.md",
    "policy.json",
    ".env.example",
    "config.example.json",
    "LOGIN_FOR_COWORK.cmd",
    "run-cowork.sh",
    "RUN_FIRST_FOR_LLM.sh",
    "START_LOGIN_FOR_USER.sh"
  )) {
    Copy-RequiredFile -Source (Join-Path $root $file) -Destination (Join-Path $linuxDir $file)
  }

  if (Test-Path -LiteralPath ".env") {
    Copy-RequiredFile -Source (Join-Path $root ".env") -Destination (Join-Path $linuxDir ".env")
  }
}

function Sync-WindowsRelease {
  npm run build:exe:win
}

function New-Payload {
  if (Test-Path -LiteralPath $buildDir) {
    Remove-Item -LiteralPath $buildDir -Recurse -Force
  }
  New-Item -ItemType Directory -Path $payloadDir -Force | Out-Null

  $linuxFiles = @(
    ".env",
    ".env.example",
    "AGENTS.md",
    "CLAUDE.md",
    "config.example.json",
    "COWORK_REQUIRED_README.md",
    "LOGIN_FOR_COWORK.cmd",
    "omh-m365.cjs",
    "omh-m365.cjs.sha256",
    "policy.json",
    "README.md",
    "run-cowork.sh",
    "RUN_FIRST_FOR_LLM.sh",
    "START_HERE_FOR_LLM.md",
    "START_LOGIN_FOR_USER.sh"
  )

  foreach ($file in $linuxFiles) {
    Copy-RequiredFile `
      -Source (Join-Path $linuxDir $file) `
      -Destination (Join-Path $payloadDir "releases\linux-llm\$file")
  }

  $winFiles = @(
    ".env",
    ".env.example",
    "omh-m365.exe",
    "omh-m365.exe.sha256",
    "policy.json"
  )

  foreach ($file in $winFiles) {
    Copy-RequiredFile `
      -Source (Join-Path $winDir $file) `
      -Destination (Join-Path $payloadDir "releases\win-x64\$file")
  }

  foreach ($blocked in @(".cache", "downloads", "logs", "node_modules", "src", "dist", "build", "scripts")) {
    if (Test-Path -LiteralPath (Join-Path $payloadDir $blocked)) {
      throw "Blocked path was staged unexpectedly: $blocked"
    }
  }

  if (Test-Path -LiteralPath $payloadZip) {
    Remove-Item -LiteralPath $payloadZip -Force
  }
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  [System.IO.Compression.ZipFile]::CreateFromDirectory(
    $payloadDir,
    $payloadZip,
    [System.IO.Compression.CompressionLevel]::Optimal,
    $false
  )
}

function New-NodeInstaller {
  $nodeInstallerDir = Join-Path $buildDir "node-installer"
  $installerJsPath = Join-Path $nodeInstallerDir "installer.cjs"
  $blobPath = Join-Path $nodeInstallerDir "installer.blob"
  $seaConfigPath = Join-Path $nodeInstallerDir "sea-config.json"

  New-Item -ItemType Directory -Path $nodeInstallerDir -Force | Out-Null

  $base64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($payloadZip))
  $chunks = New-Object System.Collections.Generic.List[string]
  for ($i = 0; $i -lt $base64.Length; $i += 1048576) {
    $length = [Math]::Min(1048576, $base64.Length - $i)
    $chunks.Add($base64.Substring($i, $length))
  }
  $partsLiteral = ($chunks | ForEach-Object { "  " + ($_ | ConvertTo-Json -Compress) }) -join ",`n"
  if (-not (Test-Path -LiteralPath $guideTemplatePath)) {
    throw "Required installer guide template is missing: $guideTemplatePath"
  }
  $guideTemplateBase64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($guideTemplatePath))

  $installerJsTemplate = @'
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const childProcess = require("node:child_process");

const payloadBase64Parts = [
__PAYLOAD_PARTS__
];
const guideTemplateBase64 = "__GUIDE_TEMPLATE_BASE64__";

function getArgValue(name) {
  const index = process.argv.indexOf(name);
  if (index >= 0 && process.argv[index + 1]) return process.argv[index + 1];
  return "";
}

function waitForEnter() {
  if (process.argv.includes("--quiet") || process.argv.includes("/quiet") || process.argv.includes("/Q")) return;
  if (!process.stdin.isTTY) return;
  process.stdout.write("\nPress Enter to close.");
  const buffer = Buffer.alloc(1);
  try {
    while (true) {
      const count = fs.readSync(0, buffer, 0, 1);
      if (count === 0 || buffer[0] === 10 || buffer[0] === 13) break;
    }
  } catch {
    // Ignore non-interactive stdin edge cases.
  }
}

function isQuiet() {
  return process.argv.includes("--quiet") || process.argv.includes("/quiet") || process.argv.includes("/Q");
}

function htmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function writeTextFile(filePath, value) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, value, "utf8");
}

function writeCmdFile(filePath, value) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, value, "utf8");
}

function openGuide(filePath) {
  if (isQuiet()) return;
  try {
    childProcess.spawn("cmd.exe", ["/c", "start", "", filePath], {
      detached: true,
      stdio: "ignore",
      windowsHide: true
    }).unref();
  } catch {
    // The guide file is still created even if Windows cannot launch it.
  }
}

function main() {
  const defaultBase = process.env.LOCALAPPDATA || process.env.USERPROFILE || process.cwd();
  const targetDir = path.resolve(getArgValue("--target") || path.join(defaultBase, "OMH", "M365Agent"));
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "omh-m365-agent-installer-"));
  const zipPath = path.join(tempDir, "payload.zip");

  try {
    fs.writeFileSync(zipPath, Buffer.from(payloadBase64Parts.join(""), "base64"));
    fs.mkdirSync(targetDir, { recursive: true });

    const psScript = [
      "$ErrorActionPreference = 'Stop'",
      "$zip = $env:OMH_M365_INSTALL_ZIP",
      "$target = $env:OMH_M365_INSTALL_TARGET",
      "New-Item -ItemType Directory -Path $target -Force | Out-Null",
      "Expand-Archive -LiteralPath $zip -DestinationPath $target -Force"
    ].join("; ");

    const result = childProcess.spawnSync(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", psScript],
      {
        stdio: "inherit",
        env: {
          ...process.env,
          OMH_M365_INSTALL_ZIP: zipPath,
          OMH_M365_INSTALL_TARGET: targetDir
        }
      }
    );

    if (result.error) throw result.error;
    if (result.status !== 0) throw new Error("Payload extraction failed.");

    for (const dir of [
      path.join(targetDir, "releases", "linux-llm", ".cache"),
      path.join(targetDir, "releases", "linux-llm", "downloads"),
      path.join(targetDir, "releases", "linux-llm", "logs")
    ]) {
      fs.mkdirSync(dir, { recursive: true });
    }

    const linuxPath = path.join(targetDir, "releases", "linux-llm");
    const promptPath = path.join(targetDir, "Cowork\uc5d0 \ubd99\uc5ec\ub123\uc744 \ud504\ub86c\ud504\ud2b8.txt");
    const legacyPromptPath = path.join(targetDir, "COWORK_FIRST_PROMPT.txt");
    const startHtmlPath = path.join(targetDir, "\uc2dc\uc791\ud558\uae30.html");
    const startCmdPath = path.join(targetDir, "\uc2dc\uc791\ud558\uae30.cmd");
    const loginCmdPath = path.join(targetDir, "\ub85c\uadf8\uc778 \uc2dc\uc791.cmd");
    const promptInstruction = "\uc774 \ud3f4\ub354\ub97c Cowork\uc5d0 \uc5f0\uacb0\ud558\uace0 COWORK_REQUIRED_README.md\ub97c \uba3c\uc800 \uc77d\uc740 \ub4a4 \uadf8\ub300\ub85c \ub530\ub77c.";
    const promptTask = "\uc791\uc5c5: <\uc6d0\ud558\ub294 Microsoft 365 \uc791\uc5c5>";
    const promptText = `${linuxPath}\r\n\r\n${promptInstruction}\r\n${promptTask}\r\n`;

    writeTextFile(promptPath, promptText);
    writeTextFile(legacyPromptPath, promptText);

    writeCmdFile(
      loginCmdPath,
      `@echo off\r\ncd /d "%~dp0releases\\linux-llm"\r\ncall "%~dp0releases\\linux-llm\\LOGIN_FOR_COWORK.cmd"\r\n`
    );

    writeCmdFile(
      startCmdPath,
      `@echo off\r\nstart "" "%~dp0\uc2dc\uc791\ud558\uae30.html"\r\n`
    );

    const escapedPrompt = htmlEscape(promptText);
    const escapedLinuxPath = htmlEscape(linuxPath);
    const escapedLoginCmd = htmlEscape(loginCmdPath);
    const escapedPromptPath = htmlEscape(promptPath);
    let html = `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>OMH M365 Agent \uc2dc\uc791\ud558\uae30</title>
  <style>
    :root { color-scheme: light; --accent: #FF6000; --text: #1f2933; --muted: #5c6670; --line: #d8dde3; --bg: #f6f8fa; --panel: #ffffff; }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Segoe UI", "Malgun Gothic", Arial, sans-serif; color: var(--text); background: var(--bg); line-height: 1.5; }
    main { max-width: 960px; margin: 0 auto; padding: 32px 24px 48px; }
    h1 { margin: 0 0 8px; font-size: 28px; }
    h2 { margin: 0 0 12px; font-size: 18px; }
    p { margin: 0 0 12px; }
    .accent { width: 64px; height: 4px; background: var(--accent); margin: 18px 0 28px; }
    .step { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 20px; margin: 16px 0; }
    .num { display: inline-flex; align-items: center; justify-content: center; width: 28px; height: 28px; margin-right: 8px; border-radius: 50%; background: var(--accent); color: #fff; font-weight: 700; }
    code, textarea { font-family: Consolas, "Courier New", monospace; }
    code { display: block; padding: 10px 12px; background: #f0f3f6; border: 1px solid var(--line); border-radius: 6px; overflow-wrap: anywhere; }
    textarea { width: 100%; min-height: 150px; resize: vertical; padding: 12px; border: 1px solid var(--line); border-radius: 6px; font-size: 14px; }
    button { border: 0; border-radius: 6px; background: var(--accent); color: white; font-weight: 700; padding: 10px 14px; cursor: pointer; }
    button.secondary { background: #344054; }
    .row { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 10px; }
    .muted { color: var(--muted); }
    .ok { font-weight: 700; }
    .guide-list { margin: 12px 0 16px; padding-left: 24px; }
    .guide-list li { margin: 8px 0; }
    .callout { background: #fff7ed; border: 1px solid #fed7aa; border-left: 4px solid var(--accent); border-radius: 6px; padding: 12px 14px; margin: 12px 0; }
    .mock-settings { display: grid; grid-template-columns: 180px 1fr; gap: 0; border: 1px solid var(--line); border-radius: 8px; overflow: hidden; background: #202124; color: #f5f5f5; margin-top: 14px; }
    .mock-sidebar { background: #18191b; padding: 14px; min-height: 300px; }
    .mock-search { background: #303134; border-radius: 6px; padding: 8px 10px; color: #b8bdc7; margin-bottom: 14px; }
    .mock-nav { padding: 8px 10px; border-radius: 6px; margin: 5px 0; color: #c9ced6; }
    .mock-nav.active { background: #3a3b3f; color: #fff; font-weight: 700; outline: 2px solid var(--accent); }
    .mock-content { background: #242528; padding: 18px; }
    .mock-section-title { font-size: 18px; font-weight: 700; margin-bottom: 12px; }
    .mock-card { border: 1px solid #55585f; border-radius: 8px; padding: 16px; background: #313236; }
    .mock-label { font-weight: 700; margin-bottom: 4px; }
    .mock-row { display: grid; grid-template-columns: 1fr 70px; gap: 8px; margin: 12px 0; }
    .mock-input { border: 2px solid var(--accent); background: #4a4b4f; border-radius: 6px; padding: 10px; color: #d8dde3; }
    .mock-add { background: var(--accent); color: #fff; border-radius: 6px; display: flex; align-items: center; justify-content: center; font-weight: 700; }
    .mock-domain { background: #505257; border-radius: 6px; padding: 10px 12px; margin-top: 8px; font-family: Consolas, "Courier New", monospace; }
    .mock-note { color: #c9ced6; font-size: 13px; margin-top: 10px; }
    @media (max-width: 720px) { .mock-settings { grid-template-columns: 1fr; } .mock-sidebar { min-height: auto; } }
  </style>
</head>
<body>
  <main>
    <h1>OMH M365 Agent \uc2dc\uc791\ud558\uae30</h1>
    <p class="muted">\uc774 \ud654\uba74\uc5d0\uc11c \ub3c4\uba54\uc778 \ud5c8\uc6a9, \ub85c\uadf8\uc778, Cowork \ud504\ub86c\ud504\ud2b8\uae4c\uc9c0 \uc21c\uc11c\ub300\ub85c \uc9c4\ud589\ud558\uba74 \ub429\ub2c8\ub2e4.</p>
    <div class="accent"></div>

    <section class="step">
      <h2><span class="num">1</span>Claude/Cowork \ub3c4\uba54\uc778 \ud5c8\uc6a9</h2>
      <p>Claude/Cowork\uac00 Microsoft Graph\uc5d0 \uc811\uc18d\ud560 \uc218 \uc788\ub3c4\ub85d \uc124\uc815\uc5d0\uc11c \ub3c4\uba54\uc778 2\uac1c\ub97c \ud5c8\uc6a9\ud569\ub2c8\ub2e4.</p>
      <ol class="guide-list">
        <li><strong>Claude \uc571 \uc67c\ucabd \uc544\ub798 \uacc4\uc815/\uc870\uc9c1\uba85</strong>(\uc608: CTO Office \u00b7 Max)\uc744 \ud074\ub9ad\ud558\uace0 <strong>\uc124\uc815</strong>\uc744 \uc5fd\ub2c8\ub2e4.</li>
        <li>\uc67c\ucabd \uc124\uc815 \uba54\ub274\uc5d0\uc11c <strong>\uae30\ub2a5</strong>\uc744 \ud074\ub9ad\ud569\ub2c8\ub2e4.</li>
        <li>\uc624\ub978\ucabd\uc5d0\uc11c <strong>\ub124\ud2b8\uc6cc\ud06c \uc811\uadfc</strong> \uc139\uc158\uc744 \ucc3e\uc2b5\ub2c8\ub2e4.</li>
        <li><strong>\ub3c4\uba54\uc778 \ud5c8\uc6a9 \ubaa9\ub85d</strong> \uc544\ub798 <strong>\ucd94\uac00 \ud5c8\uc6a9 \ub3c4\uba54\uc778</strong> \uc785\ub825\uce78\uc5d0 \uc544\ub798 \ub3c4\uba54\uc778\uc744 \ud558\ub098\uc529 \uc785\ub825\ud558\uace0 <strong>\ucd94\uac00</strong>\ub97c \ub204\ub985\ub2c8\ub2e4.</li>
      </ol>
      <code>graph.microsoft.com</code>
      <br>
      <code>login.microsoftonline.com</code>
      <div class="callout">\ub450 \ub3c4\uba54\uc778\uc774 \ubaa9\ub85d\uc5d0 \ubcf4\uc774\uba74 \uc644\ub8cc\uc785\ub2c8\ub2e4. \uc774 \uc124\uc815\uc774 \uc5c6\uc73c\uba74 Cowork\uc5d0\uc11c <code style="display:inline;padding:2px 5px;">network_error</code>, proxy, <code style="display:inline;padding:2px 5px;">403</code> \uc624\ub958\uac00 \ub0a0 \uc218 \uc788\uc2b5\ub2c8\ub2e4.</div>
      <div class="mock-settings" aria-label="Claude \ub3c4\uba54\uc778 \ud5c8\uc6a9 \uc124\uc815 \ud654\uba74 \uc608\uc2dc">
        <div class="mock-sidebar">
          <div class="mock-search">\uac80\uc0c9</div>
          <div class="mock-nav">\uc77c\ubc18</div>
          <div class="mock-nav">\uacc4\uc815</div>
          <div class="mock-nav">\uac1c\uc778\uc815\ubcf4\ubcf4\ud638</div>
          <div class="mock-nav active">\uae30\ub2a5</div>
          <div class="mock-nav">Claude Code</div>
          <div class="mock-nav">Chrome\uc6a9 Claude</div>
        </div>
        <div class="mock-content">
          <div class="mock-section-title">\ub124\ud2b8\uc6cc\ud06c \uc811\uadfc</div>
          <div class="mock-card">
            <div class="mock-label">\ub3c4\uba54\uc778 \ud5c8\uc6a9 \ubaa9\ub85d</div>
            <p class="mock-note">\uc0cc\ub4dc\ubc15\uc2a4\uac00 \uc561\uc138\uc2a4\ud560 \uc218 \uc788\ub294 \ub3c4\uba54\uc778\uc744 \uc120\ud0dd\ud569\ub2c8\ub2e4.</p>
            <div class="mock-label">\ucd94\uac00 \ud5c8\uc6a9 \ub3c4\uba54\uc778</div>
            <div class="mock-row">
              <div class="mock-input">example.com \ub610\ub294 *.example.com</div>
              <div class="mock-add">\ucd94\uac00</div>
            </div>
            <div class="mock-domain">graph.microsoft.com</div>
            <div class="mock-domain">login.microsoftonline.com</div>
          </div>
        </div>
      </div>
    </section>

    <section class="step">
      <h2><span class="num">2</span>\ucd5c\ucd08 1\ud68c \ub85c\uadf8\uc778</h2>
      <p><span class="ok">\ub85c\uadf8\uc778 \uc2dc\uc791.cmd</span>\ub97c \uc2e4\ud589\ud558\uace0 Microsoft device-code \ub85c\uadf8\uc778\uc744 \uc644\ub8cc\ud569\ub2c8\ub2e4.</p>
      <p class="muted">device code\ub294 \uc0ac\uc6a9\uc790\uac00 \uc9c1\uc811 \uc785\ub825\ud558\uace0, \ucc44\ud305\uc5d0 \ubd99\uc5ec\ub123\uc9c0 \uc54a\uc2b5\ub2c8\ub2e4.</p>
      <code>${escapedLoginCmd}</code>
    </section>

    <section class="step">
      <h2><span class="num">3</span>Cowork\uc5d0 \ud504\ub86c\ud504\ud2b8 \ubd99\uc5ec\ub123\uae30</h2>
      <p>\uc544\ub798 \ub0b4\uc6a9\uc744 \ubcf5\uc0ac\ud574 Claude Cowork\uc5d0 \ubd99\uc5ec\ub123\uc740 \ub4a4, <code style="display:inline;padding:2px 5px;">\uc791\uc5c5:</code> \ub4a4\uc5d0 \uc6d0\ud558\ub294 \uc694\uccad\uc744 \uc801\uc73c\uba74 \ub429\ub2c8\ub2e4.</p>
      <textarea id="prompt" readonly>${escapedPrompt}</textarea>
      <div class="row">
        <button id="copyPrompt">\ud504\ub86c\ud504\ud2b8 \ubcf5\uc0ac</button>
        <button class="secondary" id="selectPrompt">\uc120\ud0dd\ud558\uae30</button>
      </div>
      <p class="muted">\ud30c\uc77c\ub85c\ub3c4 \uc800\uc7a5\ub428: ${escapedPromptPath}</p>
    </section>

    <section class="step">
      <h2>\uc124\uce58 \uacbd\ub85c</h2>
      <code>${escapedLinuxPath}</code>
    </section>
  </main>
  <script>
    const promptBox = document.getElementById("prompt");
    document.getElementById("selectPrompt").addEventListener("click", () => {
      promptBox.focus();
      promptBox.select();
    });
    document.getElementById("copyPrompt").addEventListener("click", async () => {
      try {
        await navigator.clipboard.writeText(promptBox.value);
        alert("\ud504\ub86c\ud504\ud2b8\ub97c \ubcf5\uc0ac\ud588\uc2b5\ub2c8\ub2e4.");
      } catch {
        promptBox.focus();
        promptBox.select();
        alert("\ubcf5\uc0ac \ubc84\ud2bc\uc774 \ucc28\ub2e8\ub418\uba74 Ctrl+C\ub85c \ubcf5\uc0ac\ud558\uc138\uc694.");
      }
    });
  </script>
</body>
</html>`;

    const guideTemplate = Buffer.from(guideTemplateBase64, "base64").toString("utf8");
    html = guideTemplate
      .replaceAll("__ESCAPED_LOGIN_CMD__", escapedLoginCmd)
      .replaceAll("__ESCAPED_PROMPT__", escapedPrompt)
      .replaceAll("__ESCAPED_PROMPT_PATH__", escapedPromptPath)
      .replaceAll("__ESCAPED_LINUX_PATH__", escapedLinuxPath);

    writeTextFile(startHtmlPath, html);

    console.log("");
    console.log("OMH M365 Agent installed.");
    console.log(`Install path: ${targetDir}`);
    console.log("");
    console.log("Cowork folder path:");
    console.log(linuxPath);
    console.log("");
    console.log("First prompt was written to:");
    console.log(promptPath);
    console.log("");
    console.log("If login is needed, run:");
    console.log(loginCmdPath);
    console.log("");
    console.log("Start guide:");
    console.log(startHtmlPath);

    openGuide(startHtmlPath);
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

try {
  main();
  waitForEnter();
} catch (error) {
  console.error("");
  console.error("Installation failed.");
  console.error(error && error.stack ? error.stack : String(error));
  waitForEnter();
  process.exitCode = 1;
}
'@

  $installerJs = $installerJsTemplate.
    Replace("__PAYLOAD_PARTS__", $partsLiteral).
    Replace("__GUIDE_TEMPLATE_BASE64__", $guideTemplateBase64)

  Set-Content -LiteralPath $installerJsPath -Value $installerJs -Encoding UTF8
  node --check $installerJsPath

  $seaConfig = [ordered]@{
    main = "build/installer/node-installer/installer.cjs"
    output = "build/installer/node-installer/installer.blob"
    disableExperimentalSEAWarning = $true
    useSnapshot = $false
    useCodeCache = $true
  }

  $seaConfig | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $seaConfigPath -Encoding UTF8
  node --experimental-sea-config $seaConfigPath

  $nodeExe = & node -p "process.execPath"
  Copy-Item -LiteralPath $nodeExe -Destination $installerPath -Force

  npx postject $installerPath NODE_SEA_BLOB $blobPath --sentinel-fuse NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2
}

function Test-Payload {
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $zip = [System.IO.Compression.ZipFile]::OpenRead($payloadZip)
  try {
    $entries = $zip.Entries | ForEach-Object { $_.FullName.Replace("\", "/") }
    foreach ($required in @(
      "releases/linux-llm/.env",
      "releases/linux-llm/omh-m365.cjs",
      "releases/linux-llm/COWORK_REQUIRED_README.md",
      "releases/win-x64/.env",
      "releases/win-x64/omh-m365.exe"
    )) {
      if ($entries -notcontains $required) {
        throw "Required payload entry missing: $required"
      }
    }

    foreach ($forbidden in @(
      ".cache/",
      "downloads/",
      "logs/",
      "node_modules/",
      "src/",
      "dist/",
      "build/",
      "scripts/"
    )) {
      if ($entries | Where-Object { $_ -like "*$forbidden*" }) {
        throw "Forbidden payload entry found: $forbidden"
      }
    }
  } finally {
    $zip.Dispose()
  }
}

New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
Sync-LinuxRelease
Sync-WindowsRelease
New-Payload
Test-Payload

if (Test-Path -LiteralPath $installerPath) {
  Remove-Item -LiteralPath $installerPath -Force
}
New-NodeInstaller

if (-not (Test-Path -LiteralPath $installerPath)) {
  throw "Installer was not created: $installerPath"
}

$installerHash = Get-FileHash -LiteralPath $installerPath -Algorithm SHA256
Set-Content -LiteralPath "$installerPath.sha256" -Value "$($installerHash.Hash.ToLower())  $(Split-Path -Leaf $installerPath)" -Encoding ASCII

Write-Host "Created installer:"
Write-Host $installerPath
Write-Host "SHA256:"
Write-Host $installerHash.Hash.ToLower()
