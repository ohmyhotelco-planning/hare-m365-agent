import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn, spawnSync } from "node:child_process";
import test from "node:test";

const cli = path.resolve("dist/cli.js");
const htmlGuide = path.resolve(
  "release-templates/cowork-git-clone/Hare_M365_Claude_Cowork_연결가이드_fixed.html"
);

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
  assert.match(output.setupCommand, /refs\/heads\/master/);
  assert.match(output.setupCommand, /test "\$LOCAL_HEAD" = "\$REMOTE_HEAD"/);
  assert.doesNotMatch(JSON.stringify(output), /required only if npm ci/);
  assert.match(output.humanLoginCommand, /auth login-start/);
  assert.match(output.humanLoginCompleteCommand, /auth login-complete/);
  assert.match(output.instruction, /must finish within 25 seconds/);
  assert.match(output.instruction, /Never keep a poller alive/);
});

test("LLM guide requires split login and all generated runtime files", () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-guide-"));
  const result = run(["llm-guide"], dataDir);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /백그라운드, detached, setsid, nohup/);
  assert.match(result.stdout, /아래 6개 도메인/);
  assert.match(result.stdout, /registry\.npmjs\.org/);
  assert.match(result.stdout, /npm ci --prefer-offline --no-audit --no-fund/);
  assert.match(result.stdout, /npm install로 바꾸거나 여러 셸 호출에 나눠 반복하지 않는다/);
  assert.match(result.stdout, /dist\/msal-network\.js/);
  assert.match(result.stdout, /auth login-start/);
  assert.match(result.stdout, /auth login-complete/);
  assert.match(result.stdout, /dataDirPersistent가 false/);
});

test("startup blocks login when the default data directory is a hosted-session path", () => {
  const env = { ...process.env, LOCALAPPDATA: "C:\\sessions\\temporary-user" };
  delete env.HARE_M365_DATA_DIR;
  delete env.OMH_M365_DATA_DIR;
  const result = spawnSync(process.execPath, [cli], {
    cwd: process.cwd(),
    env,
    encoding: "utf8"
  });
  assert.equal(result.status, 0, result.stderr);
  const output = JSON.parse(result.stdout);
  assert.equal(output.status.dataDirPersistent, false);
  assert.equal(output.nextCommand, "PERSISTENT_DATA_DIR_REQUIRED");
  assert.equal(output.llmAction, "REQUEST_DOCUMENTS_FOLDER_ACCESS");
  assert.equal(output.humanLoginCommand, undefined);
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

test("human guide requires the current version, persistent folder, and split login", () => {
  const html = fs.readFileSync(htmlGuide, "utf8");
  assert.equal((html.match(/class="domain-row"/g) ?? []).length, 6);
  assert.match(html, /test "\$VERSION" = "0\.2\.0"/);
  assert.match(html, /test "\$LOCAL_HEAD" = "\$REMOTE_HEAD"/);
  assert.match(html, /auth login-start/);
  assert.match(html, /auth login-complete/);
  assert.match(html, /문서\(Documents\).*선택하고 허용하세요/);
  assert.doesNotMatch(html, /셸 호출을 닫거나 반환하지 말고 Microsoft 로그인 완료 후/);
});
