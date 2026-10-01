export type LocalInstallOptions = {
  dataDir: string;
  repository: string;
  branch: string;
  environment?: "cowork" | "claude-code";
  shell?: "posix" | "powershell";
};

export function buildLocalSetupCommand(options: LocalInstallOptions): string {
  const environment = options.environment ?? "cowork";
  const shell = options.shell ?? "posix";
  if (environment !== "cowork" && environment !== "claude-code") {
    throw new Error("Unsupported Hare local environment.");
  }
  if (shell !== "posix" && shell !== "powershell") {
    throw new Error("Unsupported Hare command shell.");
  }
  for (const [name, value] of Object.entries({
    dataDir: options.dataDir, repository: options.repository, branch: options.branch
  })) {
    if (typeof value !== "string" || !value.trim() || /[\x00-\x1f\x7f]/.test(value)) {
      throw new Error(`Hare local ${name} must be nonempty and contain no control characters.`);
    }
    if (value.startsWith("-")) {
      throw new Error(`Hare local ${name} cannot start with a dash.`);
    }
    // Windows PowerShell's legacy native argument passing cannot preserve embedded double quotes.
    if (shell === "powershell" && value.includes('"')) {
      throw new Error(`Hare PowerShell ${name} cannot contain double quotes.`);
    }
  }
  if (options.branch === "@" || /[\s~^:?*\[\\]|\.\.|@\{/.test(options.branch)
    || options.branch.endsWith(".")
    || options.branch.split("/").some((part) => !part || part.startsWith(".") || part.endsWith(".lock"))) {
    throw new Error("Hare local branch must be a valid literal Git branch name.");
  }
  return shell === "powershell"
    ? buildPowerShellSetup(options, environment)
    : buildPosixSetup(options, environment);
}

function buildPosixSetup(options: LocalInstallOptions, environment: "cowork" | "claude-code"): string {
  const hostFlags = environment === "claude-code" ? " --host claude-code --command-shell posix" : "";
  // A separate shell keeps errexit effective even when the caller uses an AND/OR list.
  return `sh -eu <<'HARE_LOCAL_SETUP'
hare_fail() { printf '%s\\n' "$1" >&2; exit 1; }
HARE_DATA_DIR=${shellQuote(options.dataDir)}
HARE_RUNTIME_ROOT="\${XDG_CACHE_HOME:-$HOME/.cache}/hare-m365-agent-runtime"
case "$HARE_RUNTIME_ROOT" in [A-Za-z]:*) HARE_RUNTIME_ROOT=$(node -e 'process.stdout.write(process.argv[1].replaceAll(String.fromCharCode(92), String.fromCharCode(47)))' "$HARE_RUNTIME_ROOT") ;; esac
case "$HARE_RUNTIME_ROOT" in /*|[A-Za-z]:/*) ;; *) HARE_RUNTIME_ROOT="$PWD/$HARE_RUNTIME_ROOT" ;; esac
HARE_DATA_DIR=$(node -e ${shellQuote(runtimePathCheck)} "$HARE_DATA_DIR" "$HARE_RUNTIME_ROOT")
HARE_APP="$HARE_RUNTIME_ROOT/app"
HARE_BUILD_HEAD="$HARE_RUNTIME_ROOT/.hare-app-build-head"
HARE_REPO=${shellQuote(options.repository)}
HARE_BRANCH=${shellQuote(options.branch)}
HARE_REF="refs/heads/$HARE_BRANCH"
git check-ref-format --branch "$HARE_BRANCH" >/dev/null
if [ -e "$HARE_APP" ] || [ -L "$HARE_APP" ]; then
  [ -d "$HARE_APP/.git" ] || [ -f "$HARE_APP/.git" ] || hare_fail 'Existing Hare app folder is not a Git repository; it was left unchanged.'
fi
mkdir -p "$HARE_RUNTIME_ROOT" "$HARE_DATA_DIR"
REMOTE_LINE=$(git ls-remote --exit-code --refs "$HARE_REPO" "$HARE_REF")
REMOTE_HEAD=\${REMOTE_LINE%%[[:space:]]*}
case "$REMOTE_HEAD" in ''|*[!0-9a-f]*) hare_fail 'Invalid remote branch HEAD.' ;; esac
[ "$REMOTE_LINE" = "$(printf '%s\\t%s' "$REMOTE_HEAD" "$HARE_REF")" ] || hare_fail 'Remote branch did not resolve exactly.'
if [ ! -e "$HARE_APP/.git" ]; then
  git clone --branch "$HARE_BRANCH" --single-branch --no-tags -- "$HARE_REPO" "$HARE_APP"
else
  ORIGIN=$(git -C "$HARE_APP" remote get-url origin)
  [ "$ORIGIN" = "$HARE_REPO" ] || hare_fail 'Hare app remote does not match the requested repository.'
  CURRENT_BRANCH=$(git -C "$HARE_APP" branch --show-current)
  [ "$CURRENT_BRANCH" = "$HARE_BRANCH" ] || hare_fail 'Hare app branch does not match the requested branch.'
  WORKTREE_STATUS=$(git -C "$HARE_APP" status --porcelain --untracked-files=all)
  [ -z "$WORKTREE_STATUS" ] || hare_fail 'Hare app repository is dirty; changes were left unchanged.'
  git -C "$HARE_APP" fetch origin "$HARE_REF"
  git -C "$HARE_APP" pull --ff-only origin "$HARE_REF"
fi
LOCAL_HEAD=$(git -C "$HARE_APP" rev-parse HEAD)
[ "$LOCAL_HEAD" = "$REMOTE_HEAD" ] || hare_fail 'Hare app HEAD does not match the remote branch.'
BUILT_HEAD=''
if [ -f "$HARE_BUILD_HEAD" ]; then BUILT_HEAD=$(cat "$HARE_BUILD_HEAD"); fi
if [ "$BUILT_HEAD" != "$LOCAL_HEAD" ] || [ ! -d "$HARE_APP/node_modules" ] || [ ! -f "$HARE_APP/dist/cli.js" ] || [ ! -f "$HARE_APP/dist/proxy.js" ] || [ ! -f "$HARE_APP/dist/msal-network.js" ]; then
  cd "$HARE_APP"
  npm ci --prefer-offline --no-audit --no-fund
  npm run build
  [ -d "$HARE_APP/node_modules" ] && [ -f "$HARE_APP/dist/cli.js" ] && [ -f "$HARE_APP/dist/proxy.js" ] && [ -f "$HARE_APP/dist/msal-network.js" ] || hare_fail 'Hare build output is incomplete.'
  printf '%s\\n' "$LOCAL_HEAD" > "$HARE_BUILD_HEAD"
fi
node "$HARE_APP/dist/cli.js" --data-dir "$HARE_DATA_DIR"${hostFlags} network check --environment ${environment} &&
node "$HARE_APP/dist/cli.js" --data-dir "$HARE_DATA_DIR"${hostFlags}
HARE_LOCAL_SETUP`;
}

function buildPowerShellSetup(options: LocalInstallOptions, environment: "cowork" | "claude-code"): string {
  const hostFlags = environment === "claude-code" ? ", '--host', 'claude-code', '--command-shell', 'powershell'" : "";
  return `& {
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
function Invoke-HareNative {
  param([string]$File, [string[]]$Arguments)
  & $File @Arguments
  if ($LASTEXITCODE -ne 0) { throw "$File failed (exit code $LASTEXITCODE)." }
}
try {
  $HARE_DATA_DIR = [IO.Path]::GetFullPath(${powerShellQuote(options.dataDir)})
  $HARE_REPO = ${powerShellQuote(options.repository)}
  $HARE_BRANCH = ${powerShellQuote(options.branch)}
  $HARE_REF = "refs/heads/$HARE_BRANCH"
  if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    $HARE_RUNTIME_ROOT = Join-Path $HOME '.cache/HareM365Agent/runtime'
  } else {
    $HARE_RUNTIME_ROOT = Join-Path $env:LOCALAPPDATA 'HareM365Agent/runtime'
  }
  $HARE_RUNTIME_ROOT = [IO.Path]::GetFullPath($HARE_RUNTIME_ROOT)
  Invoke-HareNative 'node' @('-e', ${powerShellQuote(runtimePathCheck)}, $HARE_DATA_DIR, $HARE_RUNTIME_ROOT) | Out-Null
  $HARE_APP = Join-Path $HARE_RUNTIME_ROOT 'app'
  $HARE_BUILD_HEAD = Join-Path $HARE_RUNTIME_ROOT '.hare-app-build-head'
  Invoke-HareNative 'git' @('check-ref-format', '--branch', $HARE_BRANCH) | Out-Null
  if ((Test-Path -LiteralPath $HARE_APP) -and !(Test-Path -LiteralPath (Join-Path $HARE_APP '.git'))) {
    throw 'Existing Hare app folder is not a Git repository; it was left unchanged.'
  }
  [IO.Directory]::CreateDirectory($HARE_RUNTIME_ROOT) | Out-Null
  [IO.Directory]::CreateDirectory($HARE_DATA_DIR) | Out-Null
  $REMOTE_LINES = @(Invoke-HareNative 'git' @('ls-remote', '--exit-code', '--refs', $HARE_REPO, $HARE_REF))
  if ($REMOTE_LINES.Count -ne 1) { throw 'Remote branch did not resolve exactly.' }
  $REMOTE_PARTS = $REMOTE_LINES[0] -split "\`t"
  if ($REMOTE_PARTS.Count -ne 2 -or $REMOTE_PARTS[0] -notmatch '^(?:[0-9a-f]{40}|[0-9a-f]{64})$' -or $REMOTE_PARTS[1] -cne $HARE_REF) {
    throw 'Invalid remote branch HEAD.'
  }
  $REMOTE_HEAD = $REMOTE_PARTS[0]
  if (!(Test-Path -LiteralPath (Join-Path $HARE_APP '.git'))) {
    Invoke-HareNative 'git' @('clone', '--branch', $HARE_BRANCH, '--single-branch', '--no-tags', '--', $HARE_REPO, $HARE_APP)
  } else {
    $ORIGIN = Invoke-HareNative 'git' @('-C', $HARE_APP, 'remote', 'get-url', 'origin')
    if ($ORIGIN -cne $HARE_REPO) { throw 'Hare app remote does not match the requested repository.' }
    $CURRENT_BRANCH = Invoke-HareNative 'git' @('-C', $HARE_APP, 'branch', '--show-current')
    if ($CURRENT_BRANCH -cne $HARE_BRANCH) { throw 'Hare app branch does not match the requested branch.' }
    $WORKTREE_STATUS = @(Invoke-HareNative 'git' @('-C', $HARE_APP, 'status', '--porcelain', '--untracked-files=all'))
    if ($WORKTREE_STATUS.Count -ne 0) { throw 'Hare app repository is dirty; changes were left unchanged.' }
    Invoke-HareNative 'git' @('-C', $HARE_APP, 'fetch', 'origin', $HARE_REF)
    Invoke-HareNative 'git' @('-C', $HARE_APP, 'pull', '--ff-only', 'origin', $HARE_REF)
  }
  $LOCAL_HEAD = Invoke-HareNative 'git' @('-C', $HARE_APP, 'rev-parse', 'HEAD')
  if ($LOCAL_HEAD -cne $REMOTE_HEAD) { throw 'Hare app HEAD does not match the remote branch.' }
  $BUILT_HEAD = ''
  if (Test-Path -LiteralPath $HARE_BUILD_HEAD -PathType Leaf) { $BUILT_HEAD = [IO.File]::ReadAllText($HARE_BUILD_HEAD).Trim() }
  $BUILD_COMPLETE = (Test-Path -LiteralPath (Join-Path $HARE_APP 'node_modules') -PathType Container)
  foreach ($output in @('dist/cli.js', 'dist/proxy.js', 'dist/msal-network.js')) {
    $BUILD_COMPLETE = $BUILD_COMPLETE -and (Test-Path -LiteralPath (Join-Path $HARE_APP $output) -PathType Leaf)
  }
  if ($BUILT_HEAD -cne $LOCAL_HEAD -or !$BUILD_COMPLETE) {
    Push-Location -LiteralPath $HARE_APP
    try {
      Invoke-HareNative 'npm.cmd' @('ci', '--prefer-offline', '--no-audit', '--no-fund')
      Invoke-HareNative 'npm.cmd' @('run', 'build')
    } finally { Pop-Location }
    if (!(Test-Path -LiteralPath (Join-Path $HARE_APP 'node_modules') -PathType Container)) { throw 'Hare build output is incomplete.' }
    foreach ($output in @('dist/cli.js', 'dist/proxy.js', 'dist/msal-network.js')) {
      if (!(Test-Path -LiteralPath (Join-Path $HARE_APP $output) -PathType Leaf)) { throw 'Hare build output is incomplete.' }
    }
    [IO.File]::WriteAllText($HARE_BUILD_HEAD, "$LOCAL_HEAD\`n")
  }
  $HARE_ARGS = @((Join-Path $HARE_APP 'dist/cli.js'), '--data-dir', $HARE_DATA_DIR${hostFlags})
  Invoke-HareNative 'node' ($HARE_ARGS + @('network', 'check', '--environment', '${environment}'))
  Invoke-HareNative 'node' $HARE_ARGS
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
}`;
}

function shellQuote(value: string): string {
  return `'${value.replaceAll("'", `'"'"'`)}'`;
}

function powerShellQuote(value: string): string {
  return `'${value.replaceAll("'", "''")}'`;
}

// Node is already required by the installer and resolves junctions/symlinks on both platforms.
const runtimePathCheck = `const fs = require('node:fs');
const path = require('node:path');
function canonical(value) {
  try { return fs.realpathSync.native(value); }
  catch (error) {
    if (error.code !== 'ENOENT') throw error;
    const parent = path.dirname(value);
    if (parent === value) throw error;
    return path.join(canonical(parent), path.basename(value));
  }
}
function contains(parent, child) {
  const relative = path.relative(parent, child);
  return relative === '' || (relative !== '..' && !relative.startsWith('..' + path.sep) && !path.isAbsolute(relative));
}
const dataDir = path.resolve(process.argv[1]);
const runtime = path.resolve(process.argv[2]);
const data = canonical(dataDir);
for (const candidate of [runtime, path.join(runtime, 'app')]) {
  const target = canonical(candidate);
  if (contains(data, target) || contains(target, data)) {
    throw new Error('Hare data directory and runtime/app must not overlap.');
  }
}
process.stdout.write(dataDir.split(path.sep).join('/'));`;
