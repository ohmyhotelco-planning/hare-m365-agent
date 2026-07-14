import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn, spawnSync } from "node:child_process";
import test, { after } from "node:test";

// os.tmpdir() resolves to /tmp on Linux, which Hare intentionally treats as a
// hosted-session (non-persistent) path. Use a home-directory fixture root so
// persistence-dependent assertions behave the same on every platform.
const fixtureRoot = fs.mkdtempSync(path.join(os.homedir(), "hare-cli-contract-"));
after(() => fs.rmSync(fixtureRoot, { recursive: true, force: true }));

function makeDataDir(prefix) {
  return fs.mkdtempSync(path.join(fixtureRoot, prefix));
}

const cli = path.resolve("dist/cli.js");
const htmlGuideDirectory = path.resolve("release-templates/cowork-git-clone");
const htmlGuideFiles = fs
  .readdirSync(htmlGuideDirectory)
  .filter((name) => name.startsWith("Hare_M365_Claude_Cowork_") && name.endsWith(".html"));
assert.equal(htmlGuideFiles.length, 1, "Exactly one distributable Cowork HTML guide must exist");
const htmlGuide = path.join(htmlGuideDirectory, htmlGuideFiles[0]);

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

test("startup migrates a legacy cache and requests one sign-in for the new application", () => {
  const dataDir = makeDataDir("hare-status-");
  fs.mkdirSync(path.join(dataDir, ".cache"), { recursive: true });
  fs.writeFileSync(path.join(dataDir, ".cache", "msal-cache.json"), "{}", "utf8");

  const result = run([], dataDir);
  assert.equal(result.status, 0, result.stderr);
  const output = JSON.parse(result.stdout);
  assert.equal(output.status.cacheFileExists, false);
  assert.equal(output.status.loggedIn, false);
  assert.equal(output.status.tokenUsable, false);
  assert.equal(output.status.authMigrationRequired, true);
  assert.equal(output.status.authReason, "AUTH_APP_CHANGED");
  assert.equal(output.setup.state, "LOGIN_START_REQUIRED");
  assert.equal(output.setup.nextAction, "RUN_LOGIN_START");
  assert.match(output.setup.instruction, /updated to a new Microsoft application/);
  assert.deepEqual(output.requiredDomains, [
    "github.com",
    "registry.npmjs.org",
    "graph.microsoft.com",
    "login.microsoftonline.com",
    "outlook.office.com",
    "ohmylab-my.sharepoint.com",
    "ohmylab.sharepoint.com"
  ]);
  assert.equal(output.appDir, path.resolve("."));
  assert.match(output.setupCommand, /npm ci --prefer-offline --no-audit --no-fund/);
  assert.match(output.setupCommand, /refs\/heads\/master/);
  assert.match(output.setupCommand, /HARE_RUNTIME_ROOT=/);
  assert.match(output.setupCommand, /HARE_APP="\$HARE_RUNTIME_ROOT\/app"/);
  assert.match(output.setupCommand, /HARE_DATA_DIR=/);
  assert.doesNotMatch(output.setupCommand, /HARE_SNAPSHOT|\.hare-app-snapshot|tar -[ctx]zf/);
  assert.match(output.setupCommand, /test "\$LOCAL_HEAD" = "\$REMOTE_HEAD"/);
  assert.match(output.setupCommand, /pull --ff-only/);
  assert.match(output.setupCommand, /--data-dir/);
  assert.match(output.setupCommand, new RegExp(escapeRegExp(dataDir)));
  assert.doesNotMatch(output.setupCommand, /rm -rf "\$HARE_DATA_DIR"|\/tmp\/hare-m365-agent|\/dev\/shm|\/home\/claude/);
  assert.doesNotMatch(JSON.stringify(output), /required only if npm ci/);
  assert.match(output.setup.nextCommand, /auth login-start/);
  assert.match(output.setup.nextCommand, /--data-dir/);
  assert.match(output.setup.nextCommand, new RegExp(escapeRegExp(dataDir)));
  assert.match(output.setup.instruction, /Never start a background or detached poller/);
  assert.equal(output.nextCommand, undefined);
  assert.equal(output.llmAction, undefined);
});

test("explicit data-dir survives as part of every follow-up command", () => {
  const dataDir = makeDataDir("hare explicit data ");
  const result = spawnSync(process.execPath, [cli, "--data-dir", dataDir], {
    cwd: process.cwd(),
    env: { ...process.env },
    encoding: "utf8"
  });

  assert.equal(result.status, 0, result.stderr);
  const output = JSON.parse(result.stdout);
  assert.equal(output.status.dataDir, dataDir);
  assert.equal(output.status.dataDirSource, "command-line");
  assert.match(output.setup.nextCommand, /--data-dir/);
  assert.match(output.setup.nextCommand, new RegExp(escapeRegExp(dataDir)));

  const statusResult = spawnSync(
    process.execPath,
    [cli, "--data-dir", dataDir, "auth", "status"],
    { cwd: process.cwd(), env: { ...process.env }, encoding: "utf8" }
  );
  assert.equal(statusResult.status, 0, statusResult.stderr);
  const statusOutput = JSON.parse(statusResult.stdout);
  assert.equal(statusOutput.dataDir, dataDir);
  assert.equal(statusOutput.setup.state, "LOGIN_START_REQUIRED");
});

test("startup writes persistent Claude rules with the exact Hare paths", () => {
  const dataDir = makeDataDir("hare-rules-");
  const result = run([], dataDir);
  assert.equal(result.status, 0, result.stderr);

  const output = JSON.parse(result.stdout);
  const rulesFile = path.join(dataDir, "claude", "hare-m365-agent-rules.md");
  assert.equal(output.status.rulesFile, rulesFile);
  assert.equal(output.status.rulesFileExists, true);
  assert.equal(output.sessionRules.path, rulesFile);
  assert.equal(output.sessionRules.exists, true);
  assert.match(output.sessionRules.instruction, /Read this file before login or Microsoft 365 lookup/);
  assert.equal(fs.existsSync(rulesFile), true);

  const rules = fs.readFileSync(rulesFile, "utf8");
  assert.match(rules, new RegExp(escapeRegExp(dataDir)));
  assert.match(rules, new RegExp(escapeRegExp(path.join(dataDir, ".cache", "msal-cache.json"))));
  assert.match(rules, new RegExp(escapeRegExp(path.join(dataDir, "downloads"))));
  assert.match(rules, new RegExp(escapeRegExp(path.join(dataDir, "results"))));
  assert.match(rules, /--data-dir/);
  assert.match(rules, /Current session app directory/);
  assert.doesNotMatch(rules, /\/tmp\/hare-m365-agent/);
  assert.match(rules, /refs\/heads\/master/);
  assert.match(rules, /loggedIn.*tokenUsable/s);
  assert.match(rules, /Do not start a new login/);
  assert.match(rules, /default lookback.*90 days/i);
  assert.match(rules, /selected when this Cowork task was opened/);
  assert.match(rules, /Do not clone the repository, run npm ci, or build inside it/);
  assert.match(rules, /without requesting folder deletion permission/);
  assert.match(rules, /NETWORK_PERMISSION_REQUIRED/);
  assert.match(rules, /session sandbox shell/);
  assert.match(rules, /AWAITING_USER_APPROVAL/);
  assert.match(rules, /explicit user approval/);
  assert.match(rules, /Hare cannot send a draft/);
  assert.match(rules, /complete untruncated message/);
  assert.match(rules, /fullBodyUnavailableCount/);
  assert.doesNotMatch(rules, /hybrid|\.hare-app-snapshot|\/home\/claude/i);
  assert.doesNotMatch(rules, /single network allowlist/);
});

test("LLM guide follows the explicit setup state contract", () => {
  const dataDir = makeDataDir("hare-guide-");
  const result = run(["llm-guide"], dataDir);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /아래 7개 도메인/);
  assert.match(result.stdout, /outlook\.office\.com/);
  assert.match(result.stdout, /registry\.npmjs\.org/);
  assert.match(result.stdout, /npm ci --prefer-offline --no-audit --no-fund/);
  assert.match(result.stdout, /outlook recent --folder all/);
  assert.match(result.stdout, /outlook flagged --folder all/);
  assert.match(result.stdout, /outlook draft new/);
  assert.match(result.stdout, /approval-token/);
  assert.match(result.stdout, /메일 발송은 지원하지 않는다/);
  assert.match(result.stdout, /chat-messages의 body와 bodyHtml은 잘리지 않은 전체 본문/);
  assert.match(result.stdout, /fullBodyUnavailableCount/);
  assert.match(result.stdout, /searchSummary를 전체 본문으로 간주하지 않는다/);
  assert.match(result.stdout, /outlook inbox는 사용자가 받은편지함을 명시한 경우에만/);
  assert.match(result.stdout, /pull --ff-only/);
  assert.doesNotMatch(result.stdout, /rm -rf "\$HARE_DATA_DIR"/);
  assert.match(result.stdout, /dist\/msal-network\.js/);
  assert.match(result.stdout, /setup\.state/);
  assert.match(result.stdout, /FOLDER_REQUIRED/);
  assert.match(result.stdout, /LOGIN_START_REQUIRED/);
  assert.match(result.stdout, /LOGIN_COMPLETE_REQUIRED/);
  assert.match(result.stdout, /setup\.nextCommand를 수정하지 않고/);
  assert.match(result.stdout, /\/root\/\.local\/share/);
  assert.match(result.stdout, /현재 Cowork 작업에 선택된 프로젝트 마운트/);
  assert.doesNotMatch(result.stdout, /computer-use|folder-access tool|%USERPROFILE%|~\/HareM365Agent/);
  assert.match(result.stdout, /\/sessions\/<session>\/mnt\/<selected-project>/);
  assert.match(result.stdout, /NETWORK_PERMISSION_REQUIRED/);
  assert.match(result.stdout, /X-Proxy-Error: blocked-by-allowlist/);
  assert.match(result.stdout, /새 Cowork 채팅/);
  assert.match(result.stdout, /세션 런타임/);
  assert.match(result.stdout, /삭제 권한을 요청하지 않는다/);
  assert.doesNotMatch(result.stdout, /하이브리드|\.hare-app-snapshot/i);
  assert.doesNotMatch(result.stdout, /유일한 네트워크 허용 목록/);
  assert.doesNotMatch(result.stdout, /%USERPROFILE%\\Documents/);
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
  assert.equal(output.setup.state, "FOLDER_REQUIRED");
  assert.equal(output.setup.nextAction, "SELECT_PROJECT_FOLDER");
  assert.equal(output.sessionRules.exists, false);
  assert.equal(output.appDir, undefined);
  assert.equal(output.setupCommand, undefined);
  assert.match(output.setup.instruction, /existing Hare project or folder/);
  assert.match(output.setup.instruction, /stop/i);
  assert.doesNotMatch(output.setup.instruction, /computer-use|folder-access|%USERPROFILE%|~\/HareM365Agent/);
});

test("--data-dir cannot falsely mark a hosted-session path as persistent", () => {
  const hostedPath = path.join(path.parse(process.cwd()).root, "sessions", "temporary-hare-data");
  const result = spawnSync(process.execPath, [cli, "--data-dir", hostedPath], {
    cwd: process.cwd(),
    env: { ...process.env },
    encoding: "utf8"
  });

  assert.equal(result.status, 0, result.stderr);
  const output = JSON.parse(result.stdout);
  assert.equal(output.status.dataDirPersistent, false);
  assert.equal(output.setup.state, "FOLDER_REQUIRED");
  assert.equal(output.status.rulesFile, undefined);
  assert.equal(
    fs.existsSync(path.join(hostedPath, "claude", "hare-m365-agent-rules.md")),
    false
  );

  const guideResult = spawnSync(
    process.execPath,
    [cli, "--data-dir", hostedPath, "llm-guide"],
    { cwd: process.cwd(), env: { ...process.env }, encoding: "utf8" }
  );
  assert.equal(guideResult.status, 0, guideResult.stderr);
  assert.match(guideResult.stdout, /FOLDER_REQUIRED/);
  assert.doesNotMatch(guideResult.stdout, new RegExp(escapeRegExp(hostedPath)));
});

test("startup, doctor, and auth status expose the same setup state", () => {
  const dataDir = makeDataDir("hare-consistent-state-");
  const startup = run([], dataDir);
  const doctor = run(["doctor"], dataDir);
  const authStatus = run(["auth", "status"], dataDir);

  assert.equal(startup.status, 0, startup.stderr);
  assert.equal(doctor.status, 0, doctor.stderr);
  assert.equal(authStatus.status, 0, authStatus.stderr);

  const setupStates = [startup, doctor, authStatus].map(
    (result) => JSON.parse(result.stdout).setup.state
  );
  assert.deepEqual(setupStates, [
    "LOGIN_START_REQUIRED",
    "LOGIN_START_REQUIRED",
    "LOGIN_START_REQUIRED"
  ]);
});

test("parallel status checks serialize cache access without leaving a lock", async () => {
  const dataDir = makeDataDir("hare-lock-");
  const cacheDir = path.join(dataDir, ".cache");
  fs.mkdirSync(cacheDir, { recursive: true });
  fs.writeFileSync(path.join(cacheDir, "msal-cache.json"), "{}", "utf8");

  const results = await Promise.all(Array.from({ length: 6 }, () => runAsync(["auth", "status"], dataDir)));
  for (const result of results) {
    assert.equal(result.status, 0, result.stderr);
    const output = JSON.parse(result.stdout);
    assert.equal(output.loggedIn, false);
    assert.equal(output.authReason, "AUTH_APP_CHANGED");
    assert.equal(output.authMigrationRequired, true);
  }
  assert.equal(fs.existsSync(path.join(cacheDir, "msal-cache.json.lock")), false);
});

test("list commands reject invalid limits before making Graph calls", () => {
  const dataDir = makeDataDir("hare-limit-");
  const result = run(["files", "search", "--query", "test", "--limit", "0"], dataDir);
  assert.notEqual(result.status, 0);
  assert.match(`${result.stdout}\n${result.stderr}`, /limit must be a positive number/);
});

test("Outlook exposes whole-mailbox recent and flagged commands", () => {
  const dataDir = makeDataDir("hare-outlook-help-");
  const result = run(["outlook", "--help"], dataDir);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /recent/);
  assert.match(result.stdout, /flagged/);
  assert.match(result.stdout, /inbox/);
});

test("Outlook exposes draft preview commands but no send command", () => {
  const dataDir = makeDataDir("hare-outlook-draft-help-");
  const result = run(["outlook", "draft", "--help"], dataDir);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /new/);
  assert.match(result.stdout, /reply/);
  assert.match(result.stdout, /forward/);
  assert.doesNotMatch(result.stdout, /\bsend\b/i);
});

test("human guide verifies split-login features without a hardcoded version", () => {
  const html = fs.readFileSync(htmlGuide, "utf8");
  assert.match(html, /\.project-step > \.cowork-choice-figure \{ order: 2;/);
  assert.match(html, /\.project-step > \.action-steps-continuation \{ order: 3;/);
  assert.match(html, /class="cowork-choice-figure"/);
  assert.equal((html.match(/class="domain-row"/g) ?? []).length, 7);
  assert.match(html, /outlook\.office\.com/);
  assert.match(html, /신규·답장·전체답장·전달 메일 초안/);
  assert.match(html, /전체 미리보기를 보여주고 내 동의를 받은 뒤에만 생성/);
  assert.match(html, /auth login-start --help/);
  assert.match(html, /auth login-complete --help/);
  assert.match(html, /HARE_DATA_DIR/);
  assert.match(html, /HARE_RUNTIME_ROOT/);
  assert.match(html, /pull --ff-only/);
  assert.doesNotMatch(html, /rm -rf "\$HARE_DATA_DIR"/);
  assert.doesNotMatch(html, /test "\$VERSION" = "/);
  assert.match(html, /test "\$LOCAL_HEAD" = "\$REMOTE_HEAD"/);
  assert.match(html, /auth login-start/);
  assert.match(html, /auth login-complete/);
  assert.match(html, /Outlook 전체 메일함에서 최근 메일이나 플래그된 메일/);
  assert.match(html, /삭제된 항목을 제외한 Outlook 전체 메일함의 최신 메일 3건/);
  assert.match(html, /setup\.state만 확인하고/);
  assert.match(html, /FOLDER_REQUIRED/);
  assert.match(html, /LOGIN_START_REQUIRED/);
  assert.match(html, /LOGIN_COMPLETE_REQUIRED/);
  assert.match(html, /setup\.nextCommand를 수정하지 않고/);
  assert.match(html, /\/root\/\.local\/share/);
  assert.match(html, /프로젝트 또는 폴더/);
  assert.match(html, /HareM365Agent.*프로젝트/s);
  assert.match(html, /alt="Cowork에서 HareM365Agent 프로젝트를 선택하는 화면"/);
  assert.match(html, /alt="Windows 폴더 선택창에서 HareM365Agent를 선택하는 화면"/);
  assert.match(html, /class="image-grid project-selection-images"/);
  assert.match(html, /class="project-menu-crop"/);
  assert.match(html, /class="project-folder-shot"/);
  assert.match(html, /1-7/);
  assert.doesNotMatch(html, /__HARE_(?:PROJECT_DROPDOWN|FOLDER_PICKER)_IMAGE__/);
  assert.doesNotMatch(html, /src=["'](?:file:|[A-Za-z]:\\)/);
  assert.doesNotMatch(html, /computer-use|%USERPROFILE%|~\/HareM365Agent/);
  assert.match(html, /2-6/);
  assert.match(html, /새 Cowork 채팅/);
  assert.match(html, /<details class="troubleshoot">\s*<summary>Cowork가 열리지 않을 때만 펼치세요<\/summary>/);
  assert.match(html, /<details class="optional-details">\s*<summary>프롬프트 직접 보기<\/summary>/);
  assert.match(html, /id="prompt" rows="16"/);
  assert.doesNotMatch(html, /<details class="(?:troubleshoot|optional-details)" open/);
  assert.doesNotMatch(html, /id="prompt" rows="64"/);
  assert.match(html, /NETWORK_PERMISSION_REQUIRED/);
  assert.match(html, /X-Proxy-Error: blocked-by-allowlist/);
  assert.match(html, /현재 선택된 프로젝트 마운트 경로/);
  assert.match(html, /현재 선택된 프로젝트 루트 하나가 Hare의 영구 dataDir/);
  assert.match(html, /설정을 바꿨다면 새 Cowork 채팅을 여세요/);
  assert.match(html, /선택 프로젝트의 삭제 권한을 요청하지 마/);
  assert.doesNotMatch(html, /하이브리드|\.hare-app-snapshot|HARE_ROOT 판별/i);
  assert.doesNotMatch(html, /유일한 네트워크 허용 목록/);
  assert.doesNotMatch(html, /로컬 실행 환경은 프록시 허용 목록/);
  assert.doesNotMatch(html, /클라우드 쪽 환경은/);
  assert.doesNotMatch(html, /%USERPROFILE%\\Documents/);
  assert.doesNotMatch(html, /문서\(Documents\) 폴더 접근을 요청해/);
  assert.doesNotMatch(html, /마운트된 Documents/);
  assert.doesNotMatch(html, /셸 호출을 닫거나 반환하지 말고 Microsoft 로그인 완료 후/);
});

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
