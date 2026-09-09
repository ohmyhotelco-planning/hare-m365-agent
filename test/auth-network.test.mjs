import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { pathToFileURL } from "node:url";
import test from "node:test";
import { PublicClientApplication } from "@azure/msal-node";
import { completeLogin, getAuthStatus, getAccessToken, getScopeList } from "../dist/auth.js";
import { hasPendingDeviceLoginState, startDeviceLogin } from "../dist/device-login.js";
import { markAuthProfileReady } from "../dist/auth-profile.js";
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
  const { config, cacheFile } = await fixture(t);
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
  }
});
