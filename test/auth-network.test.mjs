import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { pathToFileURL } from "node:url";
import test from "node:test";
import { PublicClientApplication } from "@azure/msal-node";
import { completeLogin, getAccount, getAuthStatus, getAccessToken, getScopeList, logout } from "../dist/auth.js";
import { hasPendingDeviceLoginState, startDeviceLogin } from "../dist/device-login.js";
import { getAuthProfilePath, markAuthProfileReady } from "../dist/auth-profile.js";
import { ProxyAwareNetworkClient } from "../dist/msal-network.js";
import { buildSetupContract } from "../dist/setup-state.js";

const marker = "PRIVATE_TEST_MARKER";
const scopes = getScopeList();
const host = "login.microsoftonline.com";
const jsonResponse = (body, status = 200) => new Response(JSON.stringify(body), { status });

function tokenBody(config) {
  const now = Math.floor(Date.now() / 1000);
  const encode = (value) => Buffer.from(JSON.stringify(value)).toString("base64url");
  return {
    token_type: "Bearer", scope: scopes.join(" "), expires_in: 3600,
    access_token: "synthetic-access", refresh_token: "synthetic-refresh",
    id_token: `${encode({ alg: "none" })}.${encode({
      aud: config.clientId, iss: `${config.authority}/v2.0`, iat: now, nbf: now,
      exp: now + 3600, oid: "33333333-3333-3333-3333-333333333333",
      sub: "synthetic-subject", tid: config.tenantId, preferred_username: "test@example.com", ver: "2.0"
    })}.signature`,
    client_info: encode({ uid: "33333333-3333-3333-3333-333333333333", utid: config.tenantId })
  };
}

function successfulFetch(config) {
  return async (url, options) => {
    if (options.method === "GET") {
      if (url.includes("discovery/instance")) return jsonResponse({
        tenant_discovery_endpoint: `${config.authority}/v2.0/.well-known/openid-configuration`,
        metadata: [{ preferred_network: host, preferred_cache: host, aliases: [host] }]
      });
      return jsonResponse({
        authorization_endpoint: `${config.authority}/oauth2/v2.0/authorize`,
        token_endpoint: `${config.authority}/oauth2/v2.0/token`,
        issuer: `${config.authority}/v2.0`, jwks_uri: `${config.authority}/discovery/v2.0/keys`,
        device_authorization_endpoint: `${config.authority}/oauth2/v2.0/devicecode`
      });
    }
    if (new URL(url).pathname.endsWith("/devicecode")) return jsonResponse({
      user_code: "SYNTHETIC", device_code: "synthetic-device", verification_uri: "https://example.com",
      expires_in: 900, interval: 1, message: "Synthetic login only"
    });
    assert.ok(new URL(url).pathname.endsWith("/token"));
    return jsonResponse(tokenBody(config));
  };
}

async function fixture(t) {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-network-test-"));
  t.after(() => fs.rmSync(dataDir, { recursive: true, force: true }));
  const config = {
    clientId: "11111111-1111-1111-1111-111111111111",
    tenantId: "22222222-2222-2222-2222-222222222222",
    authority: `https://${host}/22222222-2222-2222-2222-222222222222`,
    dataDir, dataDirPersistent: true, dataDirSource: "environment", cacheDir: path.join(dataDir, ".cache")
  };
  const pca = new PublicClientApplication({
    auth: { clientId: config.clientId, authority: config.authority },
    system: { networkClient: new ProxyAwareNetworkClient(successfulFetch(config)) }
  });
  await pca.acquireTokenByDeviceCode({ scopes, deviceCodeCallback: () => {} });
  markAuthProfileReady(config, scopes);
  const cache = JSON.parse(pca.getTokenCache().serialize());
  for (const value of Object.values(cache.AccessToken)) value.expires_on = "1";
  const cacheFile = path.join(config.cacheDir, "msal-cache.json");
  fs.writeFileSync(cacheFile, JSON.stringify(cache));
  const pendingFile = path.join(config.cacheDir, "device-login-state.json");
  fs.writeFileSync(pendingFile, "synthetic-pending-state");
  return { config, cacheFile, pendingFile };
}

for (const code of ["EACCES", "ENOTFOUND", "EAI_AGAIN", "ETIMEDOUT", "UND_ERR_CONNECT_TIMEOUT"]) {
  test(`MSAL ${code}: verification unknown, cache preserved, recovery without login`, async (t) => {
    const { config, cacheFile, pendingFile } = await fixture(t);
    const before = fs.readFileSync(cacheFile, "utf8");
    const okFetch = successfulFetch(config);
    let restored = false;
    const network = new ProxyAwareNetworkClient(async (url, options) => {
      if (!restored && options.method === "POST") {
        throw new TypeError(`fetch failed ${marker}`, {
          cause: new AggregateError([Object.assign(new Error(marker), { code })])
        });
      }
      return okFetch(url, options);
    });
    const status = await getAuthStatus(config, network);
    assert.equal(status.account.username, "test@example.com");
    assert.equal(status.loggedIn, null, status.reason);
    assert.equal(status.tokenUsable, null);
    assert.match(status.reason, /^AUTH_CHECK_BLOCKED:/);
    assert.deepEqual(status.networkFailure, { method: "POST", hostname: host, stage: "request", code });
    const setup = buildSetupContract({
      configured: true, dataDirPersistent: true, loggedIn: status.loggedIn, tokenUsable: status.tokenUsable,
      authMigrationRequired: false, authReason: status.reason, pendingLoginStateExists: true
    }, "hare");
    assert.equal(setup.state, "BLOCKED");
    assert.equal(setup.nextCommand, undefined);
    assert.doesNotMatch(JSON.stringify({ status, setup }), new RegExp(marker));
    assert.equal(fs.readFileSync(cacheFile, "utf8"), before);
    await assert.rejects(getAccessToken(config, network), /AUTH_CHECK_BLOCKED/);
    assert.equal(fs.readFileSync(cacheFile, "utf8"), before);
    assert.equal(fs.readFileSync(pendingFile, "utf8"), "synthetic-pending-state");
    restored = true;
    const recovered = await getAuthStatus(config, network);
    assert.equal(recovered.loggedIn, true, recovered.reason);
    assert.equal(recovered.tokenUsable, true);
    assert.equal(recovered.networkFailure, undefined);
    assert.equal(await getAccessToken(config, network), "synthetic-access");
  });
}

test("a real OAuth rejection still requests login without exposing response details", async (t) => {
  const { config } = await fixture(t);
  const okFetch = successfulFetch(config);
  const network = new ProxyAwareNetworkClient(async (url, options) => options.method === "POST"
    ? jsonResponse({ error: "interaction_required", error_description: `${marker} network_error` }, 400)
    : okFetch(url, options));
  const status = await getAuthStatus(config, network);
  assert.equal(status.loggedIn, false);
  assert.equal(status.tokenUsable, false);
  assert.equal(status.reason, "TOKEN_ACQUISITION_FAILED: interaction_required");
  assert.equal(status.networkFailure, undefined);
  assert.doesNotMatch(JSON.stringify(status), new RegExp(marker));
  assert.equal(buildSetupContract({
    configured: true, dataDirPersistent: true, loggedIn: false, tokenUsable: false,
    authMigrationRequired: false, authReason: status.reason, pendingLoginStateExists: false
  }, "hare").state, "LOGIN_START_REQUIRED");
});

test("missing account remains login-required without any network request", async (t) => {
  const { config, cacheFile } = await fixture(t);
  fs.writeFileSync(cacheFile, "{}");
  const status = await getAuthStatus(config, new ProxyAwareNetworkClient(async () => {
    assert.fail("no account should not make a network request");
  }));
  assert.equal(status.reason, "NO_ACCOUNT_IN_CACHE");
  assert.equal(status.loggedIn, false);
});

test("post-login verification blockage preserves tokens but retires the redeemed code", async (t) => {
  const { config, cacheFile } = await fixture(t);
  const okFetch = successfulFetch(config);
  await startDeviceLogin(config, scopes, new ProxyAwareNetworkClient(okFetch));
  let issued = false;
  const network = new ProxyAwareNetworkClient(async (url, options) => {
    if (options.method !== "POST") return okFetch(url, options);
    if (!issued) {
      issued = true;
      return jsonResponse({ ...tokenBody(config), expires_in: 1, ext_expires_in: 1 });
    }
    throw Object.assign(new Error(marker), { code: "EACCES" });
  });
  await assert.rejects(completeLogin(config, network), (error) => {
    assert.match(error.message, /^AUTH_CHECK_BLOCKED:/);
    assert.doesNotMatch(error.message, /Run auth login-start|PRIVATE_TEST_MARKER/);
    return true;
  });
  assert.ok(Object.keys(JSON.parse(fs.readFileSync(cacheFile, "utf8")).RefreshToken).length > 0);
  assert.equal(hasPendingDeviceLoginState(config), false);
  const rejectedNetwork = new ProxyAwareNetworkClient(async (url, options) => options.method === "POST"
    ? jsonResponse({ error: "interaction_required", error_description: marker }, 400)
    : okFetch(url, options));
  const rejected = await getAuthStatus(config, rejectedNetwork);
  assert.equal(buildSetupContract({
    configured: true, dataDirPersistent: true, loggedIn: rejected.loggedIn, tokenUsable: rejected.tokenUsable,
    authMigrationRequired: false, authReason: rejected.reason,
    pendingLoginStateExists: hasPendingDeviceLoginState(config)
  }, "hare").state, "LOGIN_START_REQUIRED");
  assert.equal((await getAuthStatus(config, new ProxyAwareNetworkClient(okFetch))).tokenUsable, true);
});

test("transport timeout, body failure, sanitization and diagnostic reset", async () => {
  const url = `https://${host}/${marker}?secret=${marker}`;
  const timedOut = new ProxyAwareNetworkClient(async (_url, options) => new Promise((_resolve, reject) => {
    options.signal.addEventListener("abort", () => reject(new Error(marker)), { once: true });
  }));
  await assert.rejects(timedOut.sendGetRequestAsync(url, undefined, 5));
  assert.deepEqual(timedOut.takeNetworkFailure(), {
    method: "GET", hostname: host, stage: "request", code: "ETIMEDOUT"
  });
  assert.equal(timedOut.takeNetworkFailure(), undefined);

  const brokenBody = new ProxyAwareNetworkClient(async () => ({
    text: async () => { throw Object.assign(new Error(marker), { code: "ECONNRESET" }); }
  }));
  await assert.rejects(brokenBody.sendGetRequestAsync(url));
  assert.deepEqual(brokenBody.takeNetworkFailure(), {
    method: "GET", hostname: host, stage: "response", code: "ECONNRESET"
  });
  const cyclic = new Error(marker);
  cyclic.cause = cyclic;
  cyclic.code = marker;
  let failing = true;
  const reset = new ProxyAwareNetworkClient(async (_url, options) => {
    assert.equal(options.headers.Authorization, marker);
    assert.equal(options.body, marker);
    if (failing) throw cyclic;
    return jsonResponse({ ok: true });
  });
  await assert.rejects(reset.sendPostRequestAsync(url, { headers: { Authorization: marker }, body: marker }));
  assert.doesNotMatch(JSON.stringify(reset.takeNetworkFailure()), new RegExp(marker));
  failing = false;
  assert.equal((await reset.sendPostRequestAsync(url, { headers: { Authorization: marker }, body: marker })).body.ok, true);
  assert.equal(reset.takeNetworkFailure(), undefined);
});

test("startup, doctor and auth status expose blocked diagnostics without leaking raw errors", async (t) => {
  const { config, cacheFile, assertPreserved } = await scopeFixture(t);
  const before = fs.readFileSync(cacheFile, "utf8");
  const preload = `
    import undici from ${JSON.stringify(pathToFileURL(path.resolve("node_modules/undici/index.js")).href)};
    const mock = new undici.MockAgent();
    mock.disableNetConnect();
    undici.setGlobalDispatcher(mock);
    const pool = mock.get("https://${host}");
    pool.intercept({ path: /discovery\\/instance/, method: "GET" }).reply(200, {
      tenant_discovery_endpoint: "${config.authority}/v2.0/.well-known/openid-configuration",
      metadata: [{ preferred_network: "${host}", preferred_cache: "${host}", aliases: ["${host}"] }]
    }).persist();
    pool.intercept({ path: /openid-configuration/, method: "GET" }).reply(200, {
      authorization_endpoint: "${config.authority}/oauth2/v2.0/authorize",
      token_endpoint: "${config.authority}/oauth2/v2.0/token",
      issuer: "${config.authority}/v2.0", jwks_uri: "${config.authority}/discovery/v2.0/keys"
    }).persist();
    pool.intercept({ path: /token/, method: "POST" }).replyWithError(
      Object.assign(new Error("${marker}"), { code: "EACCES" })
    ).persist();
  `;
  for (const command of [[], ["doctor"], ["auth", "status"]]) {
    const result = spawnSync(process.execPath, [
      "--import", `data:text/javascript,${encodeURIComponent(preload)}`,
      path.resolve("dist/cli.js"), "--data-dir", config.dataDir, ...command
    ], {
      encoding: "utf8", timeout: 10_000,
      env: {
        ...process.env, OMH_M365_CLIENT_ID: config.clientId, OMH_M365_TENANT_ID: config.tenantId,
        HTTPS_PROXY: "", https_proxy: "", HTTP_PROXY: "", http_proxy: ""
      }
    });
    assert.equal(result.status, 0, result.stderr);
    const output = JSON.parse(result.stdout);
    const status = output.status ?? output;
    assert.equal(status.loggedIn, null);
    assert.equal(status.tokenUsable, null);
    assert.equal(status.authNetworkFailure.code, "EACCES");
    assert.equal(output.setup.state, "BLOCKED");
    assert.equal(output.setup.nextCommand, undefined);
    assert.doesNotMatch(result.stdout + result.stderr, new RegExp(marker));
    assert.equal(fs.readFileSync(cacheFile, "utf8"), before);
    assertPreserved();
  }
});

async function scopeFixture(t, storedScopes = scopes.filter((scope) => scope !== "Mail.Read.Shared")) {
  const fixtureData = await fixture(t);
  const { config } = fixtureData;
  const homeAccountId = `33333333-3333-3333-3333-333333333333.${config.tenantId}`;
  markAuthProfileReady(config, storedScopes, homeAccountId);
  const profileFile = getAuthProfilePath(config);
  const files = [profileFile, fixtureData.cacheFile, fixtureData.pendingFile];
  const before = files.map((file) => fs.readFileSync(file, "utf8"));
  return { ...fixtureData, homeAccountId, profileFile, assertPreserved() {
    assert.deepEqual(files.map((file) => fs.readFileSync(file, "utf8")), before);
  } };
}

for (const [label, storedScopes] of [
  ["added", scopes.filter((scope) => scope !== "Mail.Read.Shared")],
  ["removed", [...scopes, "Calendars.Read"]],
  ["replaced", scopes.map((scope) => scope === "Mail.ReadWrite" ? "Mail.Read" : scope)],
  ["reordered", [...scopes].reverse().concat(scopes[0])]
]) {
  test(`scope ${label}: all read APIs silently validate the bound account with current scopes`, async (t) => {
    const { config, homeAccountId, profileFile } = await scopeFixture(t, storedScopes);
    const profileBefore = fs.readFileSync(profileFile, "utf8");
    const ok = successfulFetch(config);
    let refreshes = 0;
    const network = new ProxyAwareNetworkClient(async (url, options) => {
      if (options.method === "POST") {
        assert.ok(new URL(url).pathname.endsWith("/token"), "must never start device login");
        const body = new URLSearchParams(options.body);
        assert.equal(body.get("grant_type"), "refresh_token");
        const requested = new Set(body.get("scope").split(" "));
        for (const scope of scopes) assert.ok(requested.has(scope), scope);
        refreshes++;
      }
      return ok(url, options);
    });
    const status = await getAuthStatus(config, network);
    assert.equal(status.tokenUsable, true, status.reason);
    assert.equal(status.migrationRequired, false);
    assert.equal(status.account.homeAccountId, homeAccountId);
    assert.equal((await getAccount(config, network)).homeAccountId, homeAccountId);
    assert.equal(await getAccessToken(config, network), "synthetic-access");
    assert.equal(refreshes, 1, "a successful refresh is persisted and reused");
    assert.equal(fs.readFileSync(profileFile, "utf8"), profileBefore);
  });
}

for (const [label, response, reason] of [
  ["consent", { error: "consent_required", error_description: marker }, "TOKEN_ACQUISITION_FAILED: consent_required"],
  ["revoked refresh", { error: "interaction_required", suberror: "bad_token", error_description: marker }, "TOKEN_ACQUISITION_FAILED: interaction_required"],
  ["unknown", { error: "server_error", error_description: `${marker} consent_required` }, "TOKEN_ACQUISITION_FAILED: unclassified_error"],
  ["network", null, "AUTH_CHECK_BLOCKED:"]
]) {
  test(`scope drift + ${label} preserves profile/cache/pending state and classifies safely`, async (t) => {
    const { config, assertPreserved } = await scopeFixture(t);
    const ok = successfulFetch(config);
    const network = new ProxyAwareNetworkClient(async (url, options) => {
      if (options.method !== "POST") return ok(url, options);
      assert.ok(new URL(url).pathname.endsWith("/token"));
      if (!response) throw Object.assign(new Error(marker), { code: "EACCES" });
      return jsonResponse(response, 400);
    });
    const status = await getAuthStatus(config, network);
    assert.ok(status.reason.startsWith(reason), status.reason);
    assert.equal(status.tokenUsable, label === "network" ? null : false);
    const setup = buildSetupContract({ configured: true, dataDirPersistent: true,
      loggedIn: status.loggedIn, tokenUsable: status.tokenUsable, authReason: status.reason,
      authMigrationRequired: status.migrationRequired, pendingLoginStateExists: false }, "hare");
    assert.equal(setup.state, ["consent", "revoked refresh"].includes(label) ? "LOGIN_START_REQUIRED" : "BLOCKED");
    assertPreserved();
    await assert.rejects(getAccessToken(config, network), (error) => {
      assert.ok(error.message.startsWith(reason), error.message);
      assert.doesNotMatch(error.message, new RegExp(marker));
      return true;
    });
    assert.equal(await getAccount(config, network), null);
    assertPreserved();
  });
}

for (const kind of ["insufficient-scopes", "different-account"]) {
  test(`silent response with ${kind} cannot become READY or persist new tokens`, async (t) => {
    const { config, assertPreserved } = await scopeFixture(t);
    const ok = successfulFetch(config);
    const network = new ProxyAwareNetworkClient(async (url, options) => {
      if (options.method !== "POST") return ok(url, options);
      const token = tokenBody(config);
      if (kind === "insufficient-scopes") token.scope = "User.Read";
      else token.client_info = Buffer.from(JSON.stringify({ uid: "other-user", utid: config.tenantId })).toString("base64url");
      return jsonResponse(token);
    });
    const expected = kind === "insufficient-scopes" ? "AUTH_SCOPES_INSUFFICIENT" : "AUTH_ACCOUNT_MISMATCH";
    const status = await getAuthStatus(config, network);
    assert.equal(status.reason, expected);
    assertPreserved();
    await assert.rejects(getAccessToken(config, network), new RegExp(expected));
    assertPreserved();
  });
}

test("hard profile blockers prevent reads and explicit login without touching credentials", async (t) => {
  const { config, profileFile, assertPreserved } = await scopeFixture(t);
  const blockedNetwork = new ProxyAwareNetworkClient(async () => assert.fail("no network on identity mismatch"));
  for (const [field, reason] of [["clientId", "AUTH_APP_CHANGED"], ["tenantId", "AUTH_TENANT_CHANGED"]]) {
    const changed = { ...config, [field]: "other" };
    assert.equal((await getAuthStatus(changed, blockedNetwork)).reason, reason);
    assert.equal(await getAccount(changed, blockedNetwork), null);
    await assert.rejects(getAccessToken(changed, blockedNetwork), new RegExp(reason));
    await assert.rejects(startDeviceLogin(changed, scopes, blockedNetwork), new RegExp(reason));
    await assert.rejects(completeLogin(changed, blockedNetwork), new RegExp(reason));
    assertPreserved();
  }
  const saved = fs.readFileSync(profileFile, "utf8");
  for (const value of ["bad-json", JSON.stringify({ ...JSON.parse(saved), migrationRequired: true })]) {
    fs.writeFileSync(profileFile, value);
    const reason = value === "bad-json" ? "AUTH_PROFILE_INVALID" : "AUTH_MIGRATION_REQUIRED";
    assert.equal((await getAuthStatus(config, blockedNetwork)).reason, reason);
    await assert.rejects(startDeviceLogin(config, scopes, blockedNetwork), new RegExp(reason));
    assert.equal(fs.readFileSync(profileFile, "utf8"), value);
  }
  fs.writeFileSync(profileFile, saved);
  assertPreserved();
});

test("bound account never falls back and ambiguous unbound accounts do not pick the first", async (t) => {
  const { config, cacheFile, profileFile } = await scopeFixture(t);
  const cache = JSON.parse(fs.readFileSync(cacheFile, "utf8"));
  const account = Object.values(cache.Account)[0];
  const otherHome = `44444444-4444-4444-4444-444444444444.${config.tenantId}`;
  cache.Account[`${otherHome}-${host}-${config.tenantId}`] = {
    ...account, home_account_id: otherHome, local_account_id: "44444444-4444-4444-4444-444444444444",
    username: "other@example.com"
  };
  fs.writeFileSync(cacheFile, JSON.stringify(cache));
  const blocked = new ProxyAwareNetworkClient(async () => assert.fail("ambiguous accounts must not request tokens"));
  markAuthProfileReady(config, scopes.slice(1), "missing-bound-account");
  assert.equal((await getAuthStatus(config, blocked)).reason, "AUTH_ACCOUNT_MISMATCH");
  await assert.rejects(getAccessToken(config, blocked), /AUTH_ACCOUNT_MISMATCH/);
  markAuthProfileReady(config, scopes.slice(1));
  const profileBefore = fs.readFileSync(profileFile, "utf8");
  assert.equal((await getAuthStatus(config, blocked)).reason, "AUTH_ACCOUNT_SELECTION_REQUIRED");
  assert.equal(await getAccount(config, blocked), null);
  assert.equal(fs.readFileSync(profileFile, "utf8"), profileBefore);
});

test("a concurrent cache change is not overwritten by a successful silent refresh", async (t) => {
  const { config, cacheFile } = await scopeFixture(t);
  const ok = successfulFetch(config);
  const network = new ProxyAwareNetworkClient(async (url, options) => {
    if (options.method === "POST") fs.writeFileSync(cacheFile, "{}");
    return ok(url, options);
  });
  const status = await getAuthStatus(config, network);
  assert.equal(status.reason, "AUTH_CACHE_CHANGED");
  assert.equal(status.tokenUsable, false);
  assert.equal(fs.readFileSync(cacheFile, "utf8"), "{}");
});

test("a missing token result remains blocked rather than proving another login is needed", async (t) => {
  const { config, assertPreserved } = await scopeFixture(t);
  t.mock.method(PublicClientApplication.prototype, "acquireTokenSilent", async () => null);
  const network = new ProxyAwareNetworkClient(async () => assert.fail("no live network"));
  const status = await getAuthStatus(config, network);
  assert.equal(status.reason, "NO_ACCESS_TOKEN");
  assert.equal(buildSetupContract({ configured: true, dataDirPersistent: true,
    loggedIn: status.loggedIn, tokenUsable: status.tokenUsable, authReason: status.reason,
    authMigrationRequired: false, pendingLoginStateExists: false }, "hare").state, "BLOCKED");
  await assert.rejects(getAccessToken(config, network), /NO_ACCESS_TOKEN/);
  assertPreserved();
});

test("explicit login cannot switch a bound account or persist its returned credentials", async (t) => {
  const { config, cacheFile, profileFile } = await scopeFixture(t);
  const ok = successfulFetch(config);
  await startDeviceLogin(config, scopes, new ProxyAwareNetworkClient(ok));
  const before = [cacheFile, profileFile].map((file) => fs.readFileSync(file, "utf8"));
  const network = new ProxyAwareNetworkClient(async (url, options) => options.method !== "POST"
    ? ok(url, options) : jsonResponse({ ...tokenBody(config),
      client_info: Buffer.from(JSON.stringify({ uid: "other-user", utid: config.tenantId })).toString("base64url") }));
  await assert.rejects(completeLogin(config, network), /AUTH_ACCOUNT_MISMATCH/);
  assert.deepEqual([cacheFile, profileFile].map((file) => fs.readFileSync(file, "utf8")), before);
  assert.equal(hasPendingDeviceLoginState(config), false, "a redeemed device code is retired even on validation failure");
});

for (const change of ["cache-and-profile", "profile-only"]) {
  test(`login completion preserves a concurrent ${change} update during the token request`, async (t) => {
    const { config, cacheFile, profileFile } = await scopeFixture(t);
    const ok = successfulFetch(config);
    await startDeviceLogin(config, scopes, new ProxyAwareNetworkClient(ok));
    let newerCache;
    let newerProfile;
    const network = new ProxyAwareNetworkClient(async (url, options) => {
      if (options.method !== "POST") return ok(url, options);
      if (change === "cache-and-profile") await logout(config);
      markAuthProfileReady(config, scopes, "newer-account");
      newerCache = fs.existsSync(cacheFile) ? fs.readFileSync(cacheFile, "utf8") : null;
      newerProfile = fs.readFileSync(profileFile, "utf8");
      return ok(url, options);
    });
    await assert.rejects(completeLogin(config, network), /AUTH_(CACHE|PROFILE)_CHANGED/);
    assert.equal(fs.existsSync(cacheFile) ? fs.readFileSync(cacheFile, "utf8") : null, newerCache);
    assert.equal(fs.readFileSync(profileFile, "utf8"), newerProfile);
  });
}

test("logout remains explicit and allows a fresh login without leaving a lock", async (t) => {
  const { config, cacheFile } = await scopeFixture(t);
  await logout(config);
  assert.equal(fs.existsSync(cacheFile), false);
  assert.equal(fs.existsSync(`${cacheFile}.lock`), false);
  const network = new ProxyAwareNetworkClient(successfulFetch(config));
  assert.equal((await getAuthStatus(config, network)).reason, "NO_ACCOUNT_IN_CACHE");
  await startDeviceLogin(config, scopes, network);
  await completeLogin(config, network);
  assert.equal((await getAuthStatus(config, network)).tokenUsable, true);
});
