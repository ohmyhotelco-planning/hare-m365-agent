import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn, spawnSync } from "node:child_process";
import test from "node:test";

const cli = path.resolve("dist/cli.js");

function run(args, dataDir) {
  return spawnSync(process.execPath, [cli, ...args], {
    cwd: process.cwd(),
    env: { ...process.env, HARE_M365_DATA_DIR: dataDir },
    encoding: "utf8"
  });
}

function runAsync(args, dataDir) {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, [cli, ...args], {
      cwd: process.cwd(),
      env: { ...process.env, HARE_M365_DATA_DIR: dataDir }
    });
    let stdout = "";
    let stderr = "";
    child.stdout.setEncoding("utf8").on("data", (chunk) => { stdout += chunk; });
    child.stderr.setEncoding("utf8").on("data", (chunk) => { stderr += chunk; });
    child.on("close", (status) => resolve({ status, stdout, stderr }));
  });
}

test("startup never treats cache existence alone as a successful login", () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-status-"));
  fs.mkdirSync(path.join(dataDir, ".cache"), { recursive: true });
  fs.writeFileSync(path.join(dataDir, ".cache", "msal-cache.json"), "{}", "utf8");

  const result = run([], dataDir);
  assert.equal(result.status, 0, result.stderr);
  const output = JSON.parse(result.stdout);
  assert.equal(output.status.cacheFileExists, true);
  assert.equal(output.status.loggedIn, false);
  assert.equal(output.status.tokenUsable, false);
  assert.equal(output.nextCommand, "LOGIN_REQUIRED_HARD_GATE");
  assert.deepEqual(output.requiredDomains, [
    "github.com",
    "registry.npmjs.org",
    "graph.microsoft.com",
    "login.microsoftonline.com",
    "ohmylab-my.sharepoint.com",
    "ohmylab.sharepoint.com"
  ]);
  assert.match(output.setupCommand, /npm ci --prefer-offline --no-audit --no-fund/);
  assert.doesNotMatch(JSON.stringify(output), /required only if npm ci/);
  assert.match(output.instruction, /foreground/);
  assert.match(output.instruction, /Never use background/);
});

test("LLM guide requires foreground login and all generated runtime files", () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-guide-"));
  const result = run(["llm-guide"], dataDir);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /백그라운드, detached, setsid, nohup/);
  assert.match(result.stdout, /아래 6개 도메인/);
  assert.match(result.stdout, /registry\.npmjs\.org/);
  assert.match(result.stdout, /npm ci --prefer-offline --no-audit --no-fund/);
  assert.match(result.stdout, /npm install로 바꾸거나 여러 셸 호출에 나눠 반복하지 않는다/);
  assert.match(result.stdout, /dist\/msal-network\.js/);
});

test("parallel status checks serialize cache access without leaving a lock", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-lock-"));
  const cacheDir = path.join(dataDir, ".cache");
  fs.mkdirSync(cacheDir, { recursive: true });
  fs.writeFileSync(path.join(cacheDir, "msal-cache.json"), "{}", "utf8");

  const results = await Promise.all(Array.from({ length: 6 }, () => runAsync(["auth", "status"], dataDir)));
  for (const result of results) {
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(result.stdout).loggedIn, false);
  }
  assert.equal(fs.existsSync(path.join(cacheDir, "msal-cache.json.lock")), false);
});

test("list commands reject invalid limits before making Graph calls", () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-limit-"));
  const result = run(["files", "search", "--query", "test", "--limit", "0"], dataDir);
  assert.notEqual(result.status, 0);
  assert.match(`${result.stdout}\n${result.stderr}`, /limit must be a positive number/);
});
