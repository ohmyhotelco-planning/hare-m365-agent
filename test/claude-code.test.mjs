import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { spawnSync } from "node:child_process";
import test, { after } from "node:test";
import { quoteCommandArgument, commandPath } from "../dist/command-context.js";
import { buildSetupContract } from "../dist/setup-state.js";
import { checkMicrosoftConnectivity } from "../dist/network-check.js";
import { ProxyAwareNetworkClient } from "../dist/msal-network.js";

const root = fs.mkdtempSync(path.join(os.homedir(), "hare-code-contract-"));
after(() => fs.rmSync(root, { recursive: true, force: true }));
const cli = path.resolve("dist/cli.js");
// A new process never contacts Graph: these fixtures have no real credentials.
const preload = `import undici from ${JSON.stringify(pathToFileURL(path.resolve("node_modules/undici/index.js")).href)};
const mock = new undici.MockAgent(); mock.disableNetConnect(); undici.setGlobalDispatcher(mock);
mock.get('https://login.microsoftonline.com').intercept({path:'/common/v2.0/.well-known/openid-configuration',method:'GET'}).reply(200,'{}');`;
function run(dataDir, shell, args = []) {
  const env = { ...process.env, HTTPS_PROXY: "", HTTP_PROXY: "", https_proxy: "", http_proxy: "" };
  delete env.HARE_M365_COMMAND;
  return spawnSync(process.execPath, ["--import", `data:text/javascript,${encodeURIComponent(preload)}`,
    cli, "--host", "claude-code", "--command-shell", shell, "--data-dir", dataDir, ...args], {
    encoding: "utf8", timeout: 15000,
    env
  });
}

test("command quoting follows the requested shell, including Windows Git Bash", () => {
  assert.equal(quoteCommandArgument("O'Brien $HOME", "powershell"), "'O''Brien $HOME'");
  assert.equal(quoteCommandArgument("O'Brien $HOME", "posix"), `'O'"'"'Brien $HOME'`);
  assert.equal(commandPath("C:\\Users\\O'Brien\\Hare", "posix", "win32"), "C:/Users/O'Brien/Hare");
  assert.equal(commandPath("C:\\Hare", "powershell", "win32"), "C:\\Hare");
  for (const shell of ["powershell", "posix"]) {
    assert.throws(() => quoteCommandArgument("bad\nargument", shell));
  }
});

for (const shell of ["powershell", "posix"]) {
  test(`${shell}: startup and a new session retain folder, host and approval rules`, () => {
    const dataDir = fs.mkdtempSync(path.join(root, `O'Brien \uD55C\uAE00 \u65E5\u672C\u8A9E ${shell} `));
    fs.writeFileSync(path.join(dataDir, "CLAUDE.md"), "# Existing project instructions\nKeep this section.\n");
    const first = run(dataDir, shell);
    assert.equal(first.status, 0, first.stderr);
    const output = JSON.parse(first.stdout);
    assert.equal(output.host, "claude-code");
    assert.equal(output.commandShell, shell);
    assert.equal(output.status.dataDir, dataDir);
    assert.equal(output.setup.state, "LOGIN_START_REQUIRED");
    assert.match(output.setup.nextCommand, new RegExp(`--host claude-code --command-shell ${shell}`));
    assert.ok(output.setup.nextCommand.includes(quoteCommandArgument(commandPath(dataDir, shell), shell)));
    if (process.platform === "win32") {
      const command = output.setup.nextCommand.replace(/ auth login-start$/, " auth status");
      const executable = shell === "powershell" ? "powershell.exe" : "C:\\Program Files\\Git\\bin\\bash.exe";
      const args = shell === "powershell" ? ["-NoProfile", "-NonInteractive", "-Command", command] : ["-lc", command];
      const resumed = spawnSync(executable, args, { encoding: "utf8", timeout: 15000 });
      assert.equal(resumed.status, 0, resumed.stderr);
      assert.equal(JSON.parse(resumed.stdout).dataDir, dataDir);
      assert.equal(JSON.parse(resumed.stdout).setup.nextCommand, output.setup.nextCommand);
    }
    const rulesFile = path.join(dataDir, "claude", "hare-m365-agent-rules.md");
    const rules = fs.readFileSync(rulesFile, "utf8");
    assert.match(rules, /Claude Code Desktop Local/);
    assert.doesNotMatch(rules, /## Cowork network permission|new Cowork task|Cowork session runtime/);
    assert.match(rules, /only after explicit user approval/i);
    assert.match(rules, /complete AWAITING_USER_APPROVAL preview/);
    assert.match(rules, /Hare CLI is the exclusive tool/);
    assert.match(rules, /Never inspect raw authentication/);
    const next = run(dataDir, shell, ["auth", "status"]);
    assert.equal(next.status, 0, next.stderr);
    assert.equal(JSON.parse(next.stdout).dataDir, dataDir);
    assert.equal(fs.readFileSync(rulesFile, "utf8"), rules);
    const instructions = fs.readFileSync(path.join(dataDir, "CLAUDE.md"), "utf8");
    assert.match(instructions, /Keep this section/);
    assert.equal(instructions.match(/HARE_M365_AGENT_RULES_START/g).length, 1);
  });

  test(`${shell}: both LLM entry points use Local instructions, not Cowork setup`, () => {
    const dataDir = fs.mkdtempSync(path.join(root, `guide-${shell}-`));
    for (const cmd of ["llm-guide", "llm-prompt"]) {
      const result = run(dataDir, shell, [cmd]);
      assert.equal(result.status, 0, result.stderr);
      assert.match(result.stdout, /--host claude-code/);
      assert.match(result.stdout, /network check --environment claude-code/);
      assert.doesNotMatch(result.stdout, /new Cowork task|--environment cowork|Cowork session runtime/);
      assert.match(result.stdout, /Only after explicit user approval/);
      assert.match(result.stdout, /Cloud, remote, WSL/);
    }
  });
}

test("local mode never silently chooses OS-default or environment data folders", () => {
  const dataDir = path.join(root, "must-not-create");
  for (const args of [[], ["--data-dir", "relative-folder"]]) {
    const result = spawnSync(process.execPath, [cli, "--host", "claude-code", ...args], {
      encoding: "utf8", env: { ...process.env, HARE_M365_DATA_DIR: dataDir }
    });
    assert.notEqual(result.status, 0);
    assert.match(result.stderr, /FOLDER_REQUIRED/);
    assert.equal(fs.existsSync(dataDir), false);
  }
  const help = spawnSync(process.execPath, [cli, "--host", "claude-code", "--help"], { encoding: "utf8" });
  assert.equal(help.status, 0, help.stderr);
  assert.match(help.stdout, /--command-shell/);
});

test("switching host preserves unknown legacy authentication instead of replacing it", () => {
  const dataDir = fs.mkdtempSync(path.join(root, "legacy-"));
  fs.mkdirSync(path.join(dataDir, ".cache"));
  const file = path.join(dataDir, ".cache", "msal-cache.json");
  fs.writeFileSync(file, "{}");
  const result = run(dataDir, "powershell");
  assert.equal(result.status, 0, result.stderr);
  const output = JSON.parse(result.stdout);
  assert.equal(output.setup.state, "BLOCKED");
  assert.equal(output.setup.nextCommand, undefined);
  assert.equal(fs.readFileSync(file, "utf8"), "{}");
});

test("setup states retain auth decisions but choose host-appropriate recovery", () => {
  const snapshot = { configured: true, dataDirPersistent: true, loggedIn: true,
    tokenUsable: true, authMigrationRequired: false, pendingLoginStateExists: false };
  assert.equal(buildSetupContract(snapshot, "prefix", "claude-code").state, "READY");
  assert.equal(buildSetupContract(snapshot, "prefix", "claude-code").nextCommand, undefined);
  const pending = buildSetupContract({ ...snapshot, loggedIn: false, tokenUsable: false, pendingLoginStateExists: true }, "prefix", "claude-code");
  assert.equal(pending.state, "LOGIN_COMPLETE_REQUIRED");
  assert.match(pending.instruction, /After the user says login is complete/);
  const folder = buildSetupContract({ ...snapshot, dataDirPersistent: false }, "prefix", "claude-code");
  assert.match(folder.instruction, /Claude Code Desktop Local/);
  assert.doesNotMatch(folder.instruction, /new Cowork/);
  const blocked = buildSetupContract({ ...snapshot, loggedIn: null, tokenUsable: null, authReason: "AUTH_CHECK_BLOCKED: ETIMEDOUT" }, "prefix", "claude-code");
  assert.equal(blocked.state, "BLOCKED");
  assert.match(blocked.instruction, /network check --environment claude-code/);
  assert.equal(blocked.nextCommand, undefined);
});

test("Claude Code preflight is cache-free and does not masquerade as Codex elevation", async () => {
  for (const status of [200, 403]) {
    const result = await checkMicrosoftConnectivity("claude-code", new ProxyAwareNetworkClient(async () =>
      new Response("{}", { status, headers: status === 403 ? { "X-Proxy-Error": "blocked-by-allowlist" } : {} })));
    assert.equal(result.cacheAccessed, false);
    assert.equal(result.state, status === 200 ? "REACHABLE" : "NETWORK_PERMISSION_REQUIRED");
    assert.notEqual(result.nextAction, "REQUEST_EXECUTION_PERMISSION");
  }
  const denied = await checkMicrosoftConnectivity("claude-code", new ProxyAwareNetworkClient(async () => {
    throw Object.assign(new Error("denied"), { code: "EACCES" });
  }));
  assert.equal(denied.state, "NETWORK_CHECK_BLOCKED");
  assert.match(denied.instruction, /Claude Code's normal permission process/);
  const missing = path.join(root, "network-no-data");
  const check = run(missing, "powershell", ["network", "check"]);
  assert.equal(check.status, 0, check.stderr);
  assert.equal(JSON.parse(check.stdout).environment, "claude-code");
  assert.equal(fs.existsSync(missing), false);
});
