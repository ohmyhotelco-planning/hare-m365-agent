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

test("local setup clones once, reuses the app, and rebuilds only after HEAD changes", {
  skip: !commandWorks(bash, ["--version"])
}, () => {
  const fixture = fs.mkdtempSync(path.join(os.tmpdir(), "hare-local-install-"));
  const source = path.join(fixture, "source");
  const dataDir = path.join(fixture, "HareM365Agent");
  fs.mkdirSync(source, { recursive: true });

  writeFixtureRepository(source);
  git(source, "init", "-b", "master");
  git(source, "config", "user.email", "hare-test@example.invalid");
  git(source, "config", "user.name", "Hare Test");
  git(source, "add", ".");
  git(source, "commit", "-m", "initial");

  const sourceForBash = toBashPath(source);
  const dataDirForBash = toBashPath(dataDir);
  const command = buildLocalSetupCommand({
    dataDir: dataDirForBash,
    repository: sourceForBash,
    branch: "master"
  });

  const first = runBash(command);
  assert.equal(first.status, 0, first.stderr);
  assert.match(first.stdout, /--data-dir/);
  assert.equal(readBuildCount(dataDir), 1);
  assert.equal(fs.existsSync(path.join(dataDir, "app", ".git")), true);
  assert.equal(fs.existsSync(path.join(dataDir, ".hare-app-build-head")), true);
  fs.mkdirSync(path.join(dataDir, ".cache"), { recursive: true });
  const cacheFile = path.join(dataDir, ".cache", "msal-cache.json");
  fs.writeFileSync(cacheFile, "fixture-cache-must-survive", "utf8");

  const second = runBash(command);
  assert.equal(second.status, 0, second.stderr);
  assert.equal(readBuildCount(dataDir), 1, "unchanged HEAD must not reinstall or rebuild");

  fs.appendFileSync(path.join(source, "build.mjs"), "\n// remote update\n", "utf8");
  git(source, "add", "build.mjs");
  git(source, "commit", "-m", "update");

  const third = runBash(command);
  assert.equal(third.status, 0, third.stderr);
  assert.equal(readBuildCount(dataDir), 2, "changed HEAD must rebuild exactly once");
  assert.equal(fs.readFileSync(cacheFile, "utf8"), "fixture-cache-must-survive");
  assert.equal(git(source, "rev-parse", "HEAD"), git(path.join(dataDir, "app"), "rev-parse", "HEAD"));

  const snapshotFile = path.join(dataDir, ".hare-app-snapshot.tar.gz");
  assert.equal(fs.existsSync(snapshotFile), true, "rebuild must refresh the app snapshot");

  fs.rmSync(path.join(dataDir, "app"), { recursive: true, force: true });
  const fourth = runBash(command);
  assert.equal(fourth.status, 0, fourth.stderr);
  assert.equal(readBuildCount(dataDir), 2, "snapshot restore with matching HEAD must not clone or rebuild");
  assert.equal(fs.existsSync(path.join(dataDir, "app", ".git")), true, "snapshot must restore the git checkout");
  assert.equal(git(source, "rev-parse", "HEAD"), git(path.join(dataDir, "app"), "rev-parse", "HEAD"));
});

test("local setup command never uses a destructive temporary checkout", () => {
  const command = buildLocalSetupCommand({
    dataDir: "/mounted/HareM365Agent",
    repository: "https://github.com/ohmyhotelco-planning/hare-m365-agent.git",
    branch: "master"
  });

  assert.match(command, /HARE_APP="\$HARE_ROOT\/app"/);
  assert.match(command, /git -C "\$HARE_APP" pull --ff-only/);
  assert.match(command, /\.hare-app-build-head/);
  assert.match(command, /HARE_SNAPSHOT="\$HARE_ROOT\/\.hare-app-snapshot\.tar\.gz"/);
  assert.match(command, /HARE_SNAPSHOT_TMP="\$HARE_ROOT\/\.hare-app-snapshot\.tar\.gz\.tmp\.\$\$"/);
  assert.match(command, /tar -tzf "\$HARE_SNAPSHOT" >\/dev\/null/);
  assert.match(command, /tar -xzf "\$HARE_SNAPSHOT" -C "\$HARE_ROOT"/);
  assert.match(command, /tar -czf "\$HARE_SNAPSHOT_TMP" -C "\$HARE_ROOT" app/);
  assert.match(command, /mv -f "\$HARE_SNAPSHOT_TMP" "\$HARE_SNAPSHOT"/);
  assert.match(command, /--data-dir "\$HARE_ROOT"/);
  assert.doesNotMatch(command, /rm -rf|git reset|\/tmp\/|\/root\//);
});

test("local setup refuses a corrupt snapshot before extraction", {
  skip: !commandWorks(bash, ["--version"])
}, () => {
  const fixture = fs.mkdtempSync(path.join(os.tmpdir(), "hare-corrupt-snapshot-"));
  const dataDir = path.join(fixture, "HareM365Agent");
  fs.mkdirSync(dataDir, { recursive: true });
  fs.writeFileSync(path.join(dataDir, ".hare-app-snapshot.tar.gz"), "not-a-tarball", "utf8");

  const command = buildLocalSetupCommand({
    dataDir: toBashPath(dataDir),
    repository: toBashPath(path.join(fixture, "unused-source")),
    branch: "master"
  });
  const result = runBash(command);

  assert.notEqual(result.status, 0);
  assert.equal(fs.existsSync(path.join(dataDir, "app")), false);
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

function runBash(command) {
  return spawnSync(bash, ["-lc", command], {
    encoding: "utf8",
    timeout: 120_000,
    env: withoutNpxNodeShim(process.env)
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

function readBuildCount(dataDir) {
  return Number(fs.readFileSync(path.join(dataDir, ".fixture-build-count"), "utf8"));
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
