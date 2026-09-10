import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { pathToFileURL } from "node:url";
import test from "node:test";
import { checkMicrosoftConnectivity, networkExecutionGuidance } from "../dist/network-check.js";
import { ProxyAwareNetworkClient } from "../dist/msal-network.js";
import { buildBlockedSetupContract } from "../dist/setup-state.js";
import { buildLocalSetupCommand } from "../dist/local-install.js";

const marker = "PRIVATE_NETWORK_TEST_MARKER";
const endpoint = "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration";

test("public preflight uses one bounded unauthenticated request, not a login check", async () => {
  let calls = 0;
  const network = new ProxyAwareNetworkClient(async (url, options) => {
    calls++;
    assert.equal(url, endpoint);
    assert.equal(options.method, "GET");
    assert.equal(options.headers, undefined);
    assert.equal(options.body, undefined);
    assert.ok(options.signal);
    return new Response("{}", { status: 200 });
  });
  const result = await checkMicrosoftConnectivity("codex", network);
  assert.equal(calls, 1);
  assert.equal(result.state, "REACHABLE");
  assert.equal(result.timeoutMs, 3000);
  assert.equal(result.authenticationChecked, false);
  assert.equal(result.cacheAccessed, false);
});

for (const environment of ["codex", "cowork", "unknown"]) {
  for (const code of ["EACCES", "EPERM", "ENOTFOUND", "ETIMEDOUT"]) {
    test(`${environment} ${code} has bounded host-specific permission guidance`, async () => {
      let calls = 0;
      const network = new ProxyAwareNetworkClient(async () => {
        calls++;
        throw new Error(marker, { cause: Object.assign(new Error(marker), { code }) });
      });
      const result = await checkMicrosoftConnectivity(environment, network);
      const canRequest = environment === "codex" && ["EACCES", "EPERM"].includes(code);
      assert.equal(result.state, canRequest ? "EXECUTION_PERMISSION_REQUIRED" : "NETWORK_CHECK_BLOCKED");
      assert.equal(result.nextAction, canRequest ? "REQUEST_EXECUTION_PERMISSION" : "REPORT_BLOCKER");
      assert.equal(result.code, code);
      assert.equal(calls, 1);
      assert.equal(result.authenticationChecked, false);
      assert.doesNotMatch(JSON.stringify(result), new RegExp(marker));
      if (canRequest) {
        assert.match(result.instruction, /until permission is granted/);
        assert.match(result.instruction, /denied, unavailable, or still blocked, stop/);
        assert.match(result.instruction, /same Hare executable and exact dataDir/);
        assert.match(result.instruction, /Never.*automatically retry a write/);
      }
    });
  }
}

for (const environment of ["codex", "cowork", "unknown"]) {
  test(`${environment} explicit allowlist denial never requests execution elevation`, async () => {
    const result = await checkMicrosoftConnectivity(environment, new ProxyAwareNetworkClient(async () =>
      new Response(marker, { status: 403, headers: { "X-Proxy-Error": "blocked-by-allowlist" } })));
    assert.equal(result.state, "NETWORK_PERMISSION_REQUIRED");
    assert.equal(result.nextAction, "REPORT_BLOCKER");
    assert.doesNotMatch(JSON.stringify(result), new RegExp(marker));
    assert.match(result.instruction, /never move to a local shell to bypass/);
  });
}

test("ordinary HTTP denial and server failure are not allowlist claims", async () => {
  for (const status of [403, 429, 503]) {
    const result = await checkMicrosoftConnectivity("codex", new ProxyAwareNetworkClient(async () =>
      new Response(marker, { status })));
    assert.equal(result.state, "NETWORK_CHECK_BLOCKED");
    assert.equal(result.httpStatus, status);
    assert.equal(result.code, "HTTP_ERROR");
  }
});

for (const code of ["EACCES", "EPERM", "ETIMEDOUT"]) {
  test(`allowlist headers cannot be overridden by body ${code}`, async () => {
    const result = await checkMicrosoftConnectivity("codex", new ProxyAwareNetworkClient(async () => ({
      status: 403, headers: new Headers({ "X-Proxy-Error": "blocked-by-allowlist" }),
      body: { cancel: async () => { throw Object.assign(new Error(marker), { code }); } },
      text: async () => { assert.fail("preflight must never read a response body"); }
    })));
    assert.equal(result.state, "NETWORK_PERMISSION_REQUIRED");
    assert.equal(result.nextAction, "REPORT_BLOCKER");
  });
}

function runCli(dataDir, args, mode) {
  const preload = `
    import fs from 'node:fs';
    import path from 'node:path';
    import undici from ${JSON.stringify(pathToFileURL(path.resolve("node_modules/undici/index.js")).href)};
    const root = ${JSON.stringify(dataDir)};
    for (const name of ['readFileSync', 'writeFileSync', 'existsSync', 'mkdirSync', 'readdirSync', 'statSync', 'rmSync']) {
      const original = fs[name];
      fs[name] = function(p, ...rest) {
        if (typeof p === 'string' && (path.resolve(p) === root || path.resolve(p).startsWith(root + path.sep))) {
          throw new Error('PROBE_TOUCHED_DATA_DIRECTORY');
        }
        return original.call(this, p, ...rest);
      };
    }
    const mock = new undici.MockAgent();
    mock.disableNetConnect(); undici.setGlobalDispatcher(mock);
    const interceptor = mock.get('https://login.microsoftonline.com').intercept({
      path: '/common/v2.0/.well-known/openid-configuration', method: 'GET'
    });
    if (${JSON.stringify(mode)} === 'eacces') interceptor.replyWithError(Object.assign(new Error('${marker}'), {code:'EACCES'}));
    else if (${JSON.stringify(mode)} === 'redirect') interceptor.reply(302, '', { headers: { location: 'https://example.com/not-allowed' } });
    else if (${JSON.stringify(mode)} === 'timeout') interceptor.reply(200, '{}').delay(10000);
    else interceptor.reply(200, '{}');
  `;
  return spawnSync(process.execPath, [
    "--import", `data:text/javascript,${encodeURIComponent(preload)}`,
    path.resolve("dist/cli.js"), ...args
  ], {
    encoding: "utf8", timeout: 7000,
    env: { ...process.env, HARE_M365_DATA_DIR: dataDir, HTTPS_PROXY: "", https_proxy: "", HTTP_PROXY: "", http_proxy: "" }
  });
}

test("CLI preflight does not read or write the dataDir, even with existing invalid auth files", (t) => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "hare-network-cli-"));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const dataDir = path.join(root, "existing project with spaces");
  fs.mkdirSync(path.join(dataDir, ".cache"), { recursive: true });
  fs.writeFileSync(path.join(dataDir, ".cache", "msal-cache.json"), marker);
  for (const [mode, status, state] of [["ok", 0, "REACHABLE"], ["eacces", 1, "EXECUTION_PERMISSION_REQUIRED"], ["redirect", 1, "NETWORK_CHECK_BLOCKED"]]) {
    const result = runCli(dataDir, ["--data-dir", dataDir, "network", "check", "--environment", "codex"], mode);
    assert.equal(result.status, status, result.stderr);
    const output = JSON.parse(result.stdout);
    assert.equal(output.state, state);
    if (mode === "redirect") assert.equal(output.httpStatus, 302);
    assert.doesNotMatch(result.stdout + result.stderr, /PRIVATE_NETWORK_TEST_MARKER|PROBE_TOUCHED_DATA_DIRECTORY/);
    assert.deepEqual(fs.readdirSync(dataDir), [".cache"]);
    assert.equal(fs.readFileSync(path.join(dataDir, ".cache", "msal-cache.json"), "utf8"), marker);
  }
  const missing = path.join(root, "not-created");
  const result = runCli(missing, ["network", "check"], "ok");
  assert.equal(result.status, 0, result.stderr);
  assert.equal(JSON.parse(result.stdout).environment, "unknown");
  assert.equal(fs.existsSync(missing), false);
});

test("CLI preflight timeout is bounded and help/version are offline", (t) => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "hare-network-offline-"));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const dataDir = path.join(root, "not-created");
  const start = Date.now();
  const timedOut = runCli(dataDir, ["network", "check", "--environment", "codex"], "timeout");
  assert.equal(timedOut.status, 1, timedOut.stderr);
  assert.equal(JSON.parse(timedOut.stdout).code, "ETIMEDOUT");
  assert.ok(Date.now() - start < 6500);
  for (const args of [["network", "check", "--help"], ["--version"]]) {
    const result = runCli(dataDir, args, "eacces");
    assert.equal(result.status, 0, result.stderr);
    assert.doesNotMatch(result.stdout, /EXECUTION_PERMISSION_REQUIRED/);
  }
  assert.equal(fs.existsSync(dataDir), false);
});

test("machine contract and installation chain preserve preflight order and approval boundary", () => {
  const contract = buildBlockedSetupContract("AUTH_CHECK_BLOCKED: EACCES");
  assert.equal(contract.nextAction, "CHECK_EXECUTION_ENVIRONMENT");
  assert.equal(contract.nextCommand, undefined);
  assert.match(contract.instruction, /standard permission request, not an automatic retry/);
  const command = buildLocalSetupCommand({ dataDir: "/selected/project", repository: "https://example.com/repo.git", branch: "master" });
  assert.match(command, /--data-dir "\$HARE_DATA_DIR" network check --environment cowork &&\r?\nnode/);
  assert.match(networkExecutionGuidance, /Unknown hosts must not be treated as Codex/);
  assert.match(networkExecutionGuidance, /same executable and dataDir/);
});

test("generated guidance retains the exact Node, app and explicit selected project", (t) => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "hare-guidance-"));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const selected = path.join(root, "selected $literal 'project'");
  const other = path.join(root, "wrong project");
  const env = { ...process.env, HARE_M365_DATA_DIR: other };
  delete env.HARE_M365_COMMAND;
  const result = spawnSync(process.execPath, [path.resolve("dist/cli.js"), "--data-dir", selected, "llm-guide"], {
    env, encoding: "utf8", timeout: 5000
  });
  assert.equal(result.status, 0, result.stderr);
  assert.ok(result.stdout.includes(networkExecutionGuidance));
  // On Windows this selected fixture is persistent; Linux /tmp is intentionally not.
  if (process.platform === "win32") {
    const rules = fs.readFileSync(path.join(selected, "claude", "hare-m365-agent-rules.md"), "utf8");
    assert.ok(rules.includes(networkExecutionGuidance));
    const prefix = `& '${process.execPath.replaceAll("'", "''")}' '${path.resolve("dist/cli.js").replaceAll("'", "''")}' --data-dir '${selected.replaceAll("'", "''")}'`;
    assert.ok(rules.includes(prefix));
    const parsed = spawnSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", `${prefix} network check --help`], {
      env, encoding: "utf8", timeout: 5000
    });
    assert.equal(parsed.status, 0, parsed.stderr);
    assert.match(parsed.stdout, /Actual host/);
    const managed = fs.readFileSync(path.join(selected, "CLAUDE.md"), "utf8");
    assert.match(managed, /bounded network-check procedure/);
  }
  assert.equal(fs.existsSync(other), false);
});
