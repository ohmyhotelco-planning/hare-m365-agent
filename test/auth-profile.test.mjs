import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { getAuthProfilePath, markAuthProfileReady, prepareAuthProfile,
  resetAuthProfileAfterLogout } from "../dist/auth-profile.js";

const scopes = ["User.Read", "Mail.ReadWrite"];
function fixture(t) {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-auth-profile-"));
  t.after(() => fs.rmSync(dataDir, { recursive: true, force: true }));
  const config = { clientId: "client", tenantId: "tenant", dataDir,
    cacheDir: path.join(dataDir, ".cache") };
  fs.mkdirSync(config.cacheDir);
  const cache = path.join(config.cacheDir, "msal-cache.json");
  const pending = path.join(config.cacheDir, "device-login-state.json");
  return { config, cache, pending, profile: getAuthProfilePath(config) };
}

test("fresh inspection is read-only and pending first login does not require a profile", (t) => {
  const { config, pending, profile } = fixture(t);
  const expected = { migrationRequired: false, authStateCleared: false };
  assert.deepEqual(prepareAuthProfile(config, scopes), expected);
  assert.equal(fs.existsSync(profile), false);
  fs.writeFileSync(pending, "synthetic-pending");
  assert.deepEqual(prepareAuthProfile(config, scopes), expected);
  resetAuthProfileAfterLogout(config, scopes);
  assert.deepEqual(prepareAuthProfile(config, scopes), expected);
});

test("missing, invalid and unreadable profiles block without clearing existing state", (t) => {
  const { config, cache, pending, profile } = fixture(t);
  fs.writeFileSync(cache, "synthetic-cache");
  fs.writeFileSync(pending, "synthetic-pending");
  assert.equal(prepareAuthProfile(config, scopes).reason, "AUTH_PROFILE_MISSING");
  assert.equal(fs.existsSync(profile), false);
  for (const value of ["broken-json", "null", "{}", ""]) {
    fs.writeFileSync(profile, value);
    assert.equal(prepareAuthProfile(config, scopes).reason, "AUTH_PROFILE_INVALID");
    assert.equal(fs.readFileSync(profile, "utf8"), value);
  }
  fs.rmSync(profile);
  fs.mkdirSync(profile);
  assert.equal(prepareAuthProfile(config, scopes).reason, "AUTH_PROFILE_UNREADABLE");
  assert.equal(fs.readFileSync(cache, "utf8"), "synthetic-cache");
  assert.equal(fs.readFileSync(pending, "utf8"), "synthetic-pending");
});

test("client, tenant and explicit migration mismatches preserve all authentication files", (t) => {
  const { config, cache, pending, profile } = fixture(t);
  markAuthProfileReady(config, scopes, "account");
  fs.writeFileSync(cache, "synthetic-cache");
  fs.writeFileSync(pending, "synthetic-pending");
  const original = fs.readFileSync(profile, "utf8");
  for (const [field, reason] of [["clientId", "AUTH_APP_CHANGED"], ["tenantId", "AUTH_TENANT_CHANGED"]]) {
    assert.deepEqual(prepareAuthProfile({ ...config, [field]: "other" }, scopes), {
      migrationRequired: true, authStateCleared: false, reason
    });
    assert.equal(fs.readFileSync(profile, "utf8"), original);
  }
  const migration = JSON.stringify({ ...JSON.parse(original), migrationRequired: true });
  fs.writeFileSync(profile, migration);
  assert.equal(prepareAuthProfile(config, scopes).reason, "AUTH_MIGRATION_REQUIRED");
  assert.equal(fs.readFileSync(profile, "utf8"), migration);
  assert.equal(fs.readFileSync(cache, "utf8"), "synthetic-cache");
  assert.equal(fs.readFileSync(pending, "utf8"), "synthetic-pending");
});

test("scope changes retain identity for silent validation without rewriting metadata", (t) => {
  const { config, profile } = fixture(t);
  markAuthProfileReady(config, scopes, "account");
  const original = fs.readFileSync(profile, "utf8");
  for (const requested of [[...scopes, "Chat.Read"], ["User.Read"], ["User.Read", "Mail.Read"]]) {
    assert.deepEqual(prepareAuthProfile(config, requested), {
      migrationRequired: false, authStateCleared: false, homeAccountId: "account", scopesChanged: true
    });
    assert.equal(fs.readFileSync(profile, "utf8"), original);
  }
  assert.deepEqual(prepareAuthProfile(config, ["Mail.ReadWrite", "User.Read", "Mail.ReadWrite"]), {
    migrationRequired: false, authStateCleared: false, homeAccountId: "account"
  });
});
