import path from "node:path";

export type LocalInstallOptions = {
  dataDir: string;
  repository: string;
  branch: string;
};

export type LocalInstallPaths = {
  appDir: string;
  buildHeadFile: string;
};

export function getLocalInstallPaths(dataDir: string): LocalInstallPaths {
  return {
    appDir: path.join(dataDir, "app"),
    buildHeadFile: path.join(dataDir, ".hare-app-build-head")
  };
}

export function buildLocalSetupCommand(options: LocalInstallOptions): string {
  const root = shellQuote(options.dataDir);
  const repository = shellQuote(options.repository);
  const branch = shellQuote(options.branch);

  return `HARE_ROOT=${root}
HARE_APP="$HARE_ROOT/app"
HARE_BUILD_HEAD="$HARE_ROOT/.hare-app-build-head"
HARE_REPO=${repository}
HARE_BRANCH=${branch}
mkdir -p "$HARE_ROOT"
REMOTE_HEAD=$(git ls-remote "$HARE_REPO" "refs/heads/${options.branch}" | awk '{print $1}')
test -n "$REMOTE_HEAD"
if [ ! -d "$HARE_APP/.git" ]; then
  if [ -e "$HARE_APP" ] && [ -n "$(ls -A "$HARE_APP" 2>/dev/null)" ]; then
    echo "Hare app directory exists but is not a Git checkout: $HARE_APP" >&2
    exit 1
  fi
  git clone --branch "$HARE_BRANCH" --single-branch --no-tags "$HARE_REPO" "$HARE_APP"
else
  test "$(git -C "$HARE_APP" remote get-url origin)" = "$HARE_REPO"
  test "$(git -C "$HARE_APP" branch --show-current)" = "$HARE_BRANCH"
  test -z "$(git -C "$HARE_APP" status --porcelain)"
  git -C "$HARE_APP" fetch origin "$HARE_BRANCH"
  git -C "$HARE_APP" pull --ff-only origin "$HARE_BRANCH"
fi
LOCAL_HEAD=$(git -C "$HARE_APP" rev-parse HEAD)
test "$LOCAL_HEAD" = "$REMOTE_HEAD"
BUILT_HEAD=$(cat "$HARE_BUILD_HEAD" 2>/dev/null || true)
if [ "$BUILT_HEAD" != "$LOCAL_HEAD" ] || [ ! -d "$HARE_APP/node_modules" ] || [ ! -f "$HARE_APP/dist/cli.js" ] || [ ! -f "$HARE_APP/dist/proxy.js" ] || [ ! -f "$HARE_APP/dist/msal-network.js" ]; then
  cd "$HARE_APP" && npm ci --prefer-offline --no-audit --no-fund && npm run build
  printf '%s\\n' "$LOCAL_HEAD" > "$HARE_BUILD_HEAD"
fi
node "$HARE_APP/dist/cli.js" --data-dir "$HARE_ROOT"`;
}

function shellQuote(value: string): string {
  if (value.includes("\0") || value.includes("\n") || value.includes("\r")) {
    throw new Error("Hare local paths cannot contain null bytes or line breaks.");
  }
  return `'${value.replaceAll("'", `'"'"'`)}'`;
}
