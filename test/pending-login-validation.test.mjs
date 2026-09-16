import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath, pathToFileURL } from "node:url";
import test, { after } from "node:test";
import { completeLogin, getScopeList } from "../dist/auth.js";
import {
  deviceLoginStatePath,
  hasPendingDeviceLoginState,
  readDeviceLoginState,
  startDeviceLogin
} from "../dist/device-login.js";
import { isPersistentDataDir } from "../dist/config.js";
import { buildSetupContract } from "../dist/setup-state.js";

const repoRoot = fileURLToPath(new URL("../", import.meta.url));
const fixturePrefix = path.join(repoRoot, ".pending-login-validation-");
const fixtureRoot = fs.mkdtempSync(fixturePrefix);
after(() => {
  const resolved = fs.realpathSync(fixtureRoot);
  assert.equal(path.dirname(resolved), fs.realpathSync(repoRoot));
  assert.ok(path.basename(resolved).startsWith(".pending-login-validation-"));
  fs.rmSync(resolved, { recursive: true, force: true });
});

const clientId = "11111111-1111-1111-1111-111111111111";
const tenantId = "22222222-2222-2222-2222-222222222222";
const scopes = getScopeList();
const syntheticSecrets = [
  "SYNTHETIC-USER-CODE-DO-NOT-PRINT",
  "SYNTHETIC-DEVICE-CODE-DO-NOT-PRINT",
  "SYNTHETIC-RESPONSE-MESSAGE-DO-NOT-PRINT"
];

function makeConfig() {
  const dataDir = fs.mkdtempSync(path.join(fixtureRoot, "case-"));
  assert.equal(isPersistentDataDir(dataDir, "command-line"), true,
    "CLI fixtures must use a persistent, non-/tmp directory");
  const config = {
    clientId,
    tenantId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    dataDir,
    dataDirSource: "command-line",
    dataDirPersistent: true,
    cacheDir: path.join(dataDir, ".cache")
  };
  fs.mkdirSync(config.cacheDir);
  return config;
}

function validState(config, now = Date.now()) {
  return {
    version: 1,
    clientId: config.clientId,
    authority: config.authority,
    scopes: [...scopes],
    createdAt: new Date(now - 60_000).toISOString(),
    expiresAt: new Date(now + 600_000).toISOString(),
    response: {
      user_code: syntheticSecrets[0],
      device_code: syntheticSecrets[1],
      verification_uri: "https://example.invalid/device-login",
      expires_in: 900,
      interval: 5,
      message: syntheticSecrets[2]
    }
  };
}

function stateCase(name, change, options = {}) {
  return { name, ...options, contents: (config) => JSON.stringify(change(validState(config))) };
}

const invalidCases = [
  stateCase("expired", (state) => ({ ...state,
    createdAt: new Date(Date.now() - 120_000).toISOString(),
    expiresAt: new Date(Date.now() - 60_000).toISOString()
  }), { expired: true, cli: true }),
  stateCase("wrong app", (state) => ({ ...state,
    clientId: "33333333-3333-3333-3333-333333333333"
  }), { cli: true }),
  stateCase("wrong authority", (state) => ({ ...state,
    authority: "https://login.microsoftonline.com/44444444-4444-4444-4444-444444444444"
  }), { cli: true }),
  stateCase("wrong scopes", (state) => ({ ...state, scopes: ["User.Read"] }),
    { cli: true, validWithoutExpectedScopes: true }),
  stateCase("extra scope", (state) => ({ ...state, scopes: [...scopes, "Synthetic.Extra"] }),
    { validWithoutExpectedScopes: true }),
  stateCase("same-size wrong scope set", (state) => ({ ...state,
    scopes: ["Synthetic.Replacement", ...scopes.slice(1)]
  }), { validWithoutExpectedScopes: true }),
  { name: "malformed JSON", contents: () => `{\"device_code\":\"${syntheticSecrets[1]}\",`, cli: true },
  { name: "null JSON", contents: () => "null", cli: true },
  { name: "array JSON", contents: () => "[]" },
  stateCase("wrong schema", (state) => ({ ...state, version: 2 })),
  stateCase("missing schema", (state) => ({ ...state, version: undefined })),
  stateCase("invalid created date", (state) => ({ ...state, createdAt: "not-a-date" }), { cli: true }),
  stateCase("invalid expiry date", (state) => ({ ...state, expiresAt: "not-a-date" })),
  stateCase("missing created date", (state) => ({ ...state, createdAt: undefined })),
  stateCase("missing expiry date", (state) => ({ ...state, expiresAt: undefined })),
  stateCase("numeric created date", (state) => ({ ...state, createdAt: 0 })),
  stateCase("null expiry date", (state) => ({ ...state, expiresAt: null })),
  stateCase("out-of-range expiry date", (state) => ({ ...state,
    expiresAt: "+999999-01-01T00:00:00.000Z"
  })),
  stateCase("expiry equals creation", (state) => ({ ...state, createdAt: state.expiresAt })),
  stateCase("expiry precedes creation", (state) => ({ ...state,
    createdAt: new Date(Date.parse(state.expiresAt) + 1).toISOString()
  })),
  stateCase("missing scopes", (state) => ({ ...state, scopes: undefined })),
  stateCase("empty scopes", (state) => ({ ...state, scopes: [] })),
  stateCase("non-array scopes", (state) => ({ ...state, scopes: scopes.join(" ") })),
  stateCase("non-string scope", (state) => ({ ...state, scopes: [...scopes, 42] })),
  stateCase("empty string scope", (state) => ({ ...state, scopes: [...scopes, ""] })),
  stateCase("whitespace scope", (state) => ({ ...state, scopes: [...scopes, " \t"] })),
  stateCase("missing response", (state) => ({ ...state, response: undefined })),
  stateCase("null response", (state) => ({ ...state, response: null }), { cli: true }),
  stateCase("malformed response field", (state) => ({ ...state,
    response: { ...state.response, device_code: 42 }
  })),
  stateCase("nonpositive response lifetime", (state) => ({ ...state,
    response: { ...state.response, expires_in: 0 }
  })),
  stateCase("non-numeric response interval", (state) => ({ ...state,
    response: { ...state.response, interval: "5" }
  })),
  { name: "empty file", contents: () => "", cli: true },
  { name: "whitespace file", contents: () => " \n\t" },
  { name: "missing file", contents: () => undefined, cli: true }
];

function writeCase(config, entry) {
  const contents = entry.contents(config);
  if (contents !== undefined) fs.writeFileSync(deviceLoginStatePath(config), contents, "utf8");
}

function snapshotCache(config) {
  return fs.readdirSync(config.cacheDir).sort().map((name) => {
    const file = path.join(config.cacheDir, name);
    const stat = fs.statSync(file);
    return { name, contents: fs.readFileSync(file).toString("base64"),
      mtimeMs: stat.mtimeMs, size: stat.size };
  });
}

function seedPreservationSentinels(config) {
  for (const name of ["auth-profile.json", "msal-cache.json"]) {
    fs.writeFileSync(path.join(config.cacheDir, name), JSON.stringify({ synthetic: name }), "utf8");
  }
  for (const name of fs.readdirSync(config.cacheDir)) {
    fs.utimesSync(path.join(config.cacheDir, name), new Date(0), new Date(0));
  }
}

function assertNoFixtureSecrets(output) {
  for (const value of syntheticSecrets) {
    assert.equal(output.includes(value), false, "synthetic login secrets must not appear in output");
  }
}

function meaningfulRejection(error) {
  assert.ok(error instanceof Error);
  assert.ok(error.message.trim().length > 0);
  assertNoFixtureSecrets(error.message);
  return true;
}

function assertExplicitExpiryCleanup(config) {
  const file = deviceLoginStatePath(config);
  assert.ok(!fs.existsSync(file) || fs.readFileSync(file, "utf8") === "",
    "explicit expired completion must remove or empty the pending state");
  assert.equal(hasPendingDeviceLoginState(config, scopes), false);
}

for (const entry of invalidCases) {
  test(`pending inspection is read-only: ${entry.name}`, () => {
    const config = makeConfig();
    writeCase(config, entry);
    seedPreservationSentinels(config);
    const before = snapshotCache(config);
    assert.equal(hasPendingDeviceLoginState(config, scopes), false);
    assert.equal(hasPendingDeviceLoginState(config), Boolean(entry.validWithoutExpectedScopes));
    assert.deepEqual(snapshotCache(config), before);
  });

  test(`read and full completion reject before network: ${entry.name}`, async () => {
    const config = makeConfig();
    writeCase(config, entry);
    const before = snapshotCache(config);
    assert.throws(() => readDeviceLoginState(config, scopes), meaningfulRejection);
    if (entry.expired) {
      assertExplicitExpiryCleanup(config);
      writeCase(config, entry);
    } else {
      assert.deepEqual(snapshotCache(config), before);
    }

    let networkCalls = 0;
    const rejectNetwork = async () => {
      networkCalls += 1;
      throw new Error("SYNTHETIC_UNEXPECTED_NETWORK_CALL");
    };
    const network = { sendGetRequestAsync: rejectNetwork, sendPostRequestAsync: rejectNetwork };
    await assert.rejects(() => completeLogin(config, network), meaningfulRejection);
    assert.equal(networkCalls, 0, "invalid pending state must reject before metadata discovery or token polling");
    if (entry.expired) assertExplicitExpiryCleanup(config);
    else assert.deepEqual(snapshotCache(config), before);
    assert.equal(fs.existsSync(path.join(config.cacheDir, "auth-profile.json")), false);
    assert.equal(fs.existsSync(path.join(config.cacheDir, "msal-cache.json")), false);
  });
}

test("valid state accepts reordered and duplicate-equivalent scope sets without mutation", () => {
  const config = makeConfig();
  for (const pendingScopes of [scopes, [...scopes].reverse(), [...scopes, scopes[0]]]) {
    const state = { ...validState(config), scopes: pendingScopes };
    fs.writeFileSync(deviceLoginStatePath(config), JSON.stringify(state), "utf8");
    seedPreservationSentinels(config);
    const before = snapshotCache(config);
    assert.equal(hasPendingDeviceLoginState(config), true);
    assert.equal(hasPendingDeviceLoginState(config, scopes), true);
    assert.equal(hasPendingDeviceLoginState(config, [...scopes].reverse()), true);
    assert.equal(hasPendingDeviceLoginState(config, [...scopes, scopes[0]]), true);
    assert.deepEqual(readDeviceLoginState(config, scopes), state);
    assert.deepEqual(readDeviceLoginState(config), state);
    assert.deepEqual(snapshotCache(config), before);
  }
  assert.equal(hasPendingDeviceLoginState(config, []), false);
  assert.throws(() => readDeviceLoginState(config, []), meaningfulRejection);
});

test("optional expected scopes do not impose current CLI scopes when omitted", () => {
  const config = makeConfig();
  const state = { ...validState(config), scopes: ["Synthetic.Custom"] };
  fs.writeFileSync(deviceLoginStatePath(config), JSON.stringify(state), "utf8");
  assert.equal(hasPendingDeviceLoginState(config), true);
  assert.deepEqual(readDeviceLoginState(config), state);
  assert.equal(hasPendingDeviceLoginState(config, state.scopes), true);
  assert.deepEqual(readDeviceLoginState(config, state.scopes), state);
  assert.equal(hasPendingDeviceLoginState(config, scopes), false);
  assert.throws(() => readDeviceLoginState(config, scopes), meaningfulRejection);
});

test("scope comparisons cannot confuse a joined string with separate permissions", () => {
  const config = makeConfig();
  const state = { ...validState(config), scopes: ["openid\nprofile"] };
  fs.writeFileSync(deviceLoginStatePath(config), JSON.stringify(state), "utf8");
  assert.equal(hasPendingDeviceLoginState(config, ["openid", "profile"]), false);
  assert.throws(() => readDeviceLoginState(config, ["openid", "profile"]), meaningfulRejection);
});

test("explicit login-start replaces stale state without changing profile or token cache", async () => {
  const config = makeConfig();
  writeCase(config, invalidCases[0]);
  fs.writeFileSync(path.join(config.cacheDir, "auth-profile.json"), JSON.stringify({
    version: 1, clientId, tenantId, scopes, migrationRequired: false
  }));
  fs.writeFileSync(path.join(config.cacheDir, "msal-cache.json"), "{}");
  const before = snapshotCache(config).filter((entry) => entry.name !== "device-login-state.json");
  let calls = 0;
  const network = {
    async sendGetRequestAsync() { throw new Error("Unexpected GET"); },
    async sendPostRequestAsync() {
      calls++;
      return { status: 200, headers: {}, body: validState(config).response };
    }
  };
  const result = await startDeviceLogin(config, scopes, network);
  assert.equal(calls, 1);
  assert.equal(result.stage, "WAITING_FOR_USER");
  assert.equal(hasPendingDeviceLoginState(config, scopes), true);
  assert.deepEqual(snapshotCache(config).filter((entry) => entry.name !== "device-login-state.json"), before);
});

test("exact expiry boundary is invalid without inspection cleanup", (t) => {
  const config = makeConfig();
  const now = Date.parse("2030-01-02T03:04:05.000Z");
  t.mock.method(Date, "now", () => now);
  const state = { ...validState(config, now), expiresAt: new Date(now + 1).toISOString() };
  fs.writeFileSync(deviceLoginStatePath(config), JSON.stringify(state), "utf8");
  assert.equal(hasPendingDeviceLoginState(config, scopes), true);
  assert.deepEqual(readDeviceLoginState(config, scopes), state);

  state.expiresAt = new Date(now).toISOString();
  fs.writeFileSync(deviceLoginStatePath(config), JSON.stringify(state), "utf8");
  const before = snapshotCache(config);
  assert.equal(hasPendingDeviceLoginState(config, scopes), false);
  assert.deepEqual(snapshotCache(config), before);
  assert.throws(() => readDeviceLoginState(config, scopes), meaningfulRejection);
  assertExplicitExpiryCleanup(config);
});

test("READY and BLOCKED retain priority over both valid and invalid pending states", () => {
  const config = makeConfig();
  for (const expiresAt of [new Date(Date.now() + 600_000).toISOString(), "invalid"]) {
    fs.writeFileSync(deviceLoginStatePath(config), JSON.stringify({ ...validState(config), expiresAt }), "utf8");
    const snapshot = {
      configured: true, dataDirPersistent: true, loggedIn: false, tokenUsable: false,
      authMigrationRequired: false, pendingLoginStateExists: hasPendingDeviceLoginState(config, scopes)
    };
    const ready = buildSetupContract({ ...snapshot, loggedIn: true, tokenUsable: true }, "hare-test");
    assert.equal(ready.state, "READY");
    assert.equal(ready.nextCommand, undefined);
    for (const authReason of ["AUTH_CHECK_BLOCKED: network_error", "AUTH_PROFILE_INVALID",
      "TOKEN_ACQUISITION_FAILED: unclassified_error"]) {
      const blocked = buildSetupContract({ ...snapshot, authReason }, "hare-test");
      assert.equal(blocked.state, "BLOCKED");
      assert.equal(blocked.nextCommand, undefined);
    }
    const migration = buildSetupContract({ ...snapshot, authMigrationRequired: true }, "hare-test");
    assert.equal(migration.state, "BLOCKED");
    assert.equal(migration.nextCommand, undefined);
  }
});

const cli = path.join(repoRoot, "dist", "cli.js");
// Run real CLI parsing/actions, but refuse every auth request and socket connection.
// The guard is inline so this suite owns no additional fixture source files.
const cliBootstrap = `
  import net from "node:net";
  import tls from "node:tls";
  import { syncBuiltinESMExports } from "node:module";
  let calls = 0;
  const deny = () => { calls += 1; throw new Error("SYNTHETIC_UNEXPECTED_NETWORK_CALL"); };
  net.Socket.prototype.connect = deny;
  tls.connect = deny;
  globalThis.fetch = deny;
  syncBuiltinESMExports();
  process.on("exit", () => {
    if (calls) { process.stderr.write("SYNTHETIC_UNEXPECTED_NETWORK_CALL"); process.exitCode = 91; }
  });
  const { ProxyAwareNetworkClient } = await import(${JSON.stringify(pathToFileURL(path.join(repoRoot, "dist", "msal-network.js")).href)});
  for (const method of ["sendGetRequestAsync", "sendPostRequestAsync", "sendGetHeadersAsync"]) {
    ProxyAwareNetworkClient.prototype[method] = deny;
  }
  process.argv = [process.execPath, ${JSON.stringify(cli)}, ...process.argv.slice(1)];
  await import(${JSON.stringify(pathToFileURL(cli).href)});
`;

function runCli(config, args) {
  const env = {
    OMH_M365_CLIENT_ID: clientId,
    OMH_M365_TENANT_ID: tenantId,
    HARE_M365_DATA_DIR: config.dataDir,
    HARE_M365_COMMAND: "hare-test",
    OMH_M365_POLICY_PATH: path.join(config.dataDir, "absent-synthetic-policy.json")
  };
  for (const key of ["SystemRoot", "WINDIR"]) {
    if (process.env[key]) env[key] = process.env[key];
  }
  return spawnSync(process.execPath,
    ["--input-type=module", "--eval", cliBootstrap, "--", "--data-dir", config.dataDir, ...args],
    { cwd: config.dataDir, env, encoding: "utf8", timeout: 10_000, maxBuffer: 1024 * 1024 });
}

const cliCases = [
  ...invalidCases.filter((entry) => entry.cli),
  stateCase("valid", (state) => state, { pending: true }),
  stateCase("reordered equivalent scopes", (state) => ({ ...state, scopes: [...scopes].reverse() }),
    { pending: true })
];

for (const [entrypoint, args] of [["startup", []], ["doctor", ["doctor"]], ["auth status", ["auth", "status"]]]) {
  for (const entry of cliCases) {
    test(`${entrypoint} routes ${entry.name} without auth-file mutation or secret output`, () => {
      const config = makeConfig();
      writeCase(config, entry);
      const before = snapshotCache(config);
      const result = runCli(config, args);
      assertNoFixtureSecrets(`${result.stdout}\n${result.stderr}`);
      assert.ifError(result.error);
      assert.equal(result.status, 0, result.stderr);
      assert.equal(result.stderr, "");
      const output = JSON.parse(result.stdout);
      const status = entrypoint === "startup" ? output.status : output;
      assert.equal(status.dataDir, config.dataDir);
      assert.equal(status.dataDirPersistent, true);
      assert.equal(status.cacheFileExists, false);
      assert.equal(status.loggedIn, false);
      assert.equal(status.tokenUsable, false);
      assert.equal(status.authReason, "NO_ACCOUNT_IN_CACHE");
      assert.equal(status.pendingLoginStateExists, Boolean(entry.pending));
      assert.equal(output.setup.state, entry.pending ? "LOGIN_COMPLETE_REQUIRED" : "LOGIN_START_REQUIRED");
      assert.equal(output.setup.nextAction,
        entry.pending ? "WAIT_FOR_USER_THEN_RUN_LOGIN_COMPLETE" : "RUN_LOGIN_START");
      const quotedDataDir = process.platform === "win32"
        ? `'${config.dataDir.replaceAll("'", "''")}'`
        : `'${config.dataDir.replaceAll("'", `'"'"'`)}'`;
      assert.equal(output.setup.nextCommand,
        `hare-test --data-dir ${quotedDataDir} auth ${entry.pending ? "login-complete" : "login-start"}`);
      assert.deepEqual(snapshotCache(config), before);
    });
  }
}
