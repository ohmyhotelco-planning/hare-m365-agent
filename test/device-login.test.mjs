import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import {
  readDeviceLoginState,
  ResumeDeviceCodeNetworkClient,
  startDeviceLogin
} from "../dist/device-login.js";
import { completeLogin, getAuthStatus, getScopeList } from "../dist/auth.js";

function configFor(dataDir) {
  return {
    clientId: "11111111-1111-1111-1111-111111111111",
    tenantId: "22222222-2222-2222-2222-222222222222",
    authority: "https://login.microsoftonline.com/22222222-2222-2222-2222-222222222222",
    dataDir,
    dataDirSource: "environment",
    dataDirPersistent: true,
    cacheDir: path.join(dataDir, ".cache")
  };
}

const serverResponse = {
  user_code: "ABCD-EFGH",
  device_code: "device-code-secret",
  verification_uri: "https://microsoft.com/devicelogin",
  expires_in: 900,
  interval: 5,
  message: "Open the Microsoft login page and enter the code."
};

test("login-start returns user instructions immediately and stores resumable state", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-device-start-"));
  const config = configFor(dataDir);
  const network = {
    async sendGetRequestAsync() { throw new Error("unexpected GET"); },
    async sendPostRequestAsync() {
      return { status: 200, headers: {}, body: serverResponse };
    }
  };

  const startedAt = Date.now();
  const result = await startDeviceLogin(config, ["User.Read"], network);
  assert.ok(Date.now() - startedAt < 1000);
  assert.equal(result.stage, "WAITING_FOR_USER");
  assert.equal(result.userCode, serverResponse.user_code);
  assert.equal("deviceCode" in result, false);
  assert.equal(readDeviceLoginState(config).response.device_code, serverResponse.device_code);
});

test("resume network client reuses the saved code once, then delegates token polling", async () => {
  const delegated = [];
  const delegate = {
    async sendGetRequestAsync(url) {
      delegated.push(["GET", url]);
      return { status: 200, headers: {}, body: {} };
    },
    async sendPostRequestAsync(url) {
      delegated.push(["POST", url]);
      return { status: 400, headers: {}, body: { error: "authorization_pending" } };
    }
  };
  const state = {
    version: 1,
    clientId: "client",
    authority: "authority",
    scopes: ["User.Read"],
    createdAt: new Date().toISOString(),
    expiresAt: new Date(Date.now() + 60_000).toISOString(),
    response: serverResponse
  };
  const network = new ResumeDeviceCodeNetworkClient(state, delegate);
  const endpoint = "https://login.microsoftonline.com/t/oauth2/v2.0/devicecode";

  const first = await network.sendPostRequestAsync(endpoint);
  assert.equal(first.status, 200);
  assert.equal(first.body.device_code, serverResponse.device_code);
  await network.sendPostRequestAsync(endpoint);
  await network.sendGetRequestAsync("https://login.microsoftonline.com/metadata");
  assert.deepEqual(delegated.map(([method]) => method), ["POST", "GET"]);
});

test("login-start refuses an unmounted hosted-session data directory", async () => {
  const config = {
    ...configFor(path.join(path.sep, "sessions", "temporary", "hare")),
    dataDirSource: "os-default",
    dataDirPersistent: false
  };
  await assert.rejects(
    () => startDeviceLogin(config, ["User.Read"]),
    /FOLDER_REQUIRED/
  );
});

test("login-complete resumes the saved device flow and writes an MSAL cache quickly", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-device-complete-"));
  const config = configFor(dataDir);
  const startNetwork = {
    async sendGetRequestAsync() { throw new Error("unexpected GET"); },
    async sendPostRequestAsync() { return { status: 200, headers: {}, body: serverResponse }; }
  };
  const scopes = getScopeList();
  await startDeviceLogin(config, scopes, startNetwork);

  const now = Math.floor(Date.now() / 1000);
  const idToken = unsignedJwt({
    aud: config.clientId,
    iss: `${config.authority}/v2.0`,
    iat: now,
    nbf: now,
    exp: now + 3600,
    oid: "33333333-3333-3333-3333-333333333333",
    sub: "subject",
    tid: config.tenantId,
    preferred_username: "test@example.com",
    name: "Test User",
    ver: "2.0"
  });
  const clientInfo = base64Url(JSON.stringify({
    uid: "33333333-3333-3333-3333-333333333333",
    utid: config.tenantId
  }));
  const completeNetwork = {
    async sendGetRequestAsync(url) {
      if (url.includes("discovery/instance")) {
        return {
          status: 200,
          headers: {},
          body: {
            tenant_discovery_endpoint: `${config.authority}/v2.0/.well-known/openid-configuration`,
            metadata: [{
              preferred_network: "login.microsoftonline.com",
              preferred_cache: "login.windows.net",
              aliases: ["login.microsoftonline.com", "login.windows.net"]
            }]
          }
        };
      }
      return {
        status: 200,
        headers: {},
        body: {
          authorization_endpoint: `${config.authority}/oauth2/v2.0/authorize`,
          token_endpoint: `${config.authority}/oauth2/v2.0/token`,
          issuer: `${config.authority}/v2.0`,
          jwks_uri: `${config.authority}/discovery/v2.0/keys`,
          device_authorization_endpoint: `${config.authority}/oauth2/v2.0/devicecode`,
          end_session_endpoint: `${config.authority}/oauth2/v2.0/logout`
        }
      };
    },
    async sendPostRequestAsync(url) {
      assert.match(url, /\/oauth2\/v2\.0\/token/);
      return {
        status: 200,
        headers: {},
        body: {
          token_type: "Bearer",
          scope: scopes.join(" "),
          expires_in: 3600,
          ext_expires_in: 3600,
          access_token: "mock-access-token",
          refresh_token: "mock-refresh-token",
          id_token: idToken,
          client_info: clientInfo
        }
      };
    }
  };

  const startedAt = Date.now();
  const result = await completeLogin(config, completeNetwork);
  assert.ok(Date.now() - startedAt < 1000);
  assert.equal(result.account?.username, "test@example.com");
  assert.equal(fs.existsSync(path.join(config.cacheDir, "msal-cache.json")), true);
  assert.equal(fs.existsSync(path.join(config.cacheDir, "device-login-state.json")), false);
  const status = await getAuthStatus(config);
  assert.equal(status.loggedIn, true, status.reason);
  assert.equal(status.tokenUsable, true, status.reason);
});

function unsignedJwt(payload) {
  return `${base64Url(JSON.stringify({ alg: "none", typ: "JWT" }))}.${base64Url(JSON.stringify(payload))}.signature`;
}

function base64Url(value) {
  return Buffer.from(value).toString("base64url");
}
