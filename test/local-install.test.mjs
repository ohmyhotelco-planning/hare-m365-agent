import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import test from "node:test";
import { buildLocalSetupCommand } from "../dist/local-install.js";

const bash = process.platform === "win32"
  ? "C:\\Program Files\\Git\\bin\\bash.exe"
  : "bash";

test("setup keeps app in session runtime and data in the selected project", {
  skip: !commandWorks(bash, ["--version"])
}, () => {
  const fixture = fs.mkdtempSync(path.join(os.tmpdir(), "hare-local-install-"));
  const source = path.join(fixture, "source");
  const dataDir = path.join(fixture, "selected-project");
  const runtimeCache = path.join(fixture, "runtime-cache");
  const runtimeRoot = path.join(runtimeCache, "hare-m365-agent-runtime");
  fs.mkdirSync(source, { recursive: true });

  writeFixtureRepository(source);
  git(source, "init", "-b", "master");
  git(source, "config", "user.email", "hare-test@example.invalid");
  git(source, "config", "user.name", "Hare Test");
  git(source, "add", ".");
  git(source, "commit", "-m", "initial");

  const command = buildLocalSetupCommand({
    dataDir: toBashPath(dataDir),
    repository: toBashPath(source),
    branch: "master"
  });
  const runtimeEnv = { XDG_CACHE_HOME: toBashPath(runtimeCache) };

  const first = runBash(command, runtimeEnv);
  assert.equal(first.status, 0, first.stderr);
  assert.match(first.stdout, /--data-dir/);
  assert.equal(readBuildCount(runtimeRoot), 1);
  assert.equal(fs.existsSync(path.join(runtimeRoot, "app", ".git")), true);
  assert.equal(fs.existsSync(path.join(dataDir, "app")), false);

  fs.mkdirSync(path.join(dataDir, ".cache"), { recursive: true });
  const cacheFile = path.join(dataDir, ".cache", "msal-cache.json");
  fs.writeFileSync(cacheFile, "fixture-cache-must-survive", "utf8");

  const second = runBash(command, runtimeEnv);
  assert.equal(second.status, 0, second.stderr);
  assert.equal(readBuildCount(runtimeRoot), 1, "unchanged HEAD must not rebuild");

  fs.appendFileSync(path.join(source, "build.mjs"), "\n// remote update\n", "utf8");
  git(source, "add", "build.mjs");
  git(source, "commit", "-m", "update");

  const third = runBash(command, runtimeEnv);
  assert.equal(third.status, 0, third.stderr);
  assert.equal(readBuildCount(runtimeRoot), 2, "changed HEAD must rebuild exactly once");
  assert.equal(fs.readFileSync(cacheFile, "utf8"), "fixture-cache-must-survive");
  assert.equal(
    git(source, "rev-parse", "HEAD"),
    git(path.join(runtimeRoot, "app"), "rev-parse", "HEAD")
  );
  assert.equal(fs.existsSync(path.join(dataDir, ".hare-app-snapshot.tar.gz")), false);
});

test("setup command separates runtime operations from the persistent data directory", () => {
  const command = buildLocalSetupCommand({
    dataDir: "/sessions/example/mnt/HareM365Agent",
    repository: "https://github.com/ohmyhotelco-planning/hare-m365-agent.git",
    branch: "master"
  });

  assert.match(command, /HARE_DATA_DIR='\/sessions\/example\/mnt\/HareM365Agent'/);
  assert.match(command, /HARE_RUNTIME_ROOT="\$\{XDG_CACHE_HOME:-\$HOME\/\.cache\}\/hare-m365-agent-runtime"/);
  assert.match(command, /HARE_APP="\$HARE_RUNTIME_ROOT\/app"/);
  assert.match(command, /git -C "\$HARE_APP" pull --ff-only/);
  assert.match(command, /npm ci --prefer-offline --no-audit --no-fund/);
  assert.match(command, /--data-dir "\$HARE_DATA_DIR"/);
  assert.doesNotMatch(command, /HARE_SNAPSHOT|tar -[ctx]zf|\.hare-app-snapshot/);
  assert.doesNotMatch(command, /cd "\$HARE_DATA_DIR"|\$HARE_DATA_DIR\/app/);
  assert.doesNotMatch(command, /\/tmp\/|\/dev\/shm|\/home\/claude|\/root\/\.local\/share/);
});

function writeFixtureRepository(source) {
  fs.writeFileSync(path.join(source, "package.json"), JSON.stringify({
    name: "hare-local-install-fixture",
    version: "1.0.0",
    type: "module",
    scripts: { build: "node build.mjs" }
  }, null, 2), "utf8");
  fs.writeFileSync(path.join(source, "package-lock.json"), JSON.stringify({
    name: "hare-local-install-fixture",
    version: "1.0.0",
    lockfileVersion: 3,
    requires: true,
    packages: {
      "": { name: "hare-local-install-fixture", version: "1.0.0" }
    }
  }, null, 2), "utf8");
  fs.writeFileSync(path.join(source, ".gitignore"), "dist/\nnode_modules/\n", "utf8");
  fs.writeFileSync(path.join(source, "build.mjs"), `import fs from "node:fs";
import path from "node:path";
const countFile = path.resolve("..", ".fixture-build-count");
const count = Number(fs.existsSync(countFile) ? fs.readFileSync(countFile, "utf8") : "0") + 1;
fs.writeFileSync(countFile, String(count), "utf8");
fs.mkdirSync("node_modules", { recursive: true });
fs.mkdirSync("dist", { recursive: true });
fs.writeFileSync("dist/cli.js", 'console.log(JSON.stringify({ args: process.argv.slice(2) }));\\n', "utf8");
fs.writeFileSync("dist/proxy.js", "export {};\\n", "utf8");
fs.writeFileSync("dist/msal-network.js", "export {};\\n", "utf8");
`, "utf8");
}

function commandWorks(command, args) {
  return spawnSync(command, args, { encoding: "utf8" }).status === 0;
}

function runBash(command, envOverrides = {}) {
  return spawnSync(bash, ["-lc", command], {
    encoding: "utf8",
    timeout: 120_000,
    env: { ...withoutNpxNodeShim(process.env), ...envOverrides }
  });
}

function toBashPath(value) {
  if (process.platform !== "win32") return value;
  const result = spawnSync(bash, ["-lc", `cygpath -u '${value.replaceAll("'", `'"'"'`)}'`], {
    encoding: "utf8"
  });
  assert.equal(result.status, 0, result.stderr);
  return result.stdout.trim();
}

function git(cwd, ...args) {
  const result = spawnSync("git", args, { cwd, encoding: "utf8" });
  assert.equal(result.status, 0, result.stderr);
  return result.stdout.trim();
}

function readBuildCount(runtimeRoot) {
  return Number(fs.readFileSync(path.join(runtimeRoot, ".fixture-build-count"), "utf8"));
}

function withoutNpxNodeShim(sourceEnv) {
  const env = { ...sourceEnv };
  for (const key of Object.keys(env)) {
    if (key.toLowerCase() !== "path" || !env[key]) continue;
    env[key] = env[key]
      .split(path.delimiter)
      .filter((entry) => !/[\\/]npm-cache[\\/]_npx[\\/]/i.test(entry))
      .join(path.delimiter);
  }
  return env;
}
