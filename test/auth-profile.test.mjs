import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import {
  getAuthProfilePath,
  markAuthProfileReady,
  prepareAuthProfile,
  resetAuthProfileAfterLogout
} from "../dist/auth-profile.js";

function configFor(dataDir, clientId = "new-client") {
  return {
    clientId,
    tenantId: "tenant",
    authority: "https://login.microsoftonline.com/tenant",
    dataDir,
    dataDirSource: "environment",
    dataDirPersistent: true,
    cacheDir: path.join(dataDir, ".cache")
  };
}

const scopes = ["User.Read", "Mail.ReadWrite"];

test("legacy authentication state is cleared while user data is preserved", () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-auth-profile-legacy-"));
  const config = configFor(dataDir);
  const cacheFile = path.join(config.cacheDir, "msal-cache.json");
  const pendingFile = path.join(config.cacheDir, "device-login-state.json");
  const downloadFile = path.join(dataDir, "downloads", "keep.txt");
  const rulesFile = path.join(dataDir, "claude", "hare-m365-agent-rules.md");

  fs.mkdirSync(config.cacheDir, { recursive: true });
  fs.mkdirSync(path.dirname(downloadFile), { recursive: true });
  fs.mkdirSync(path.dirname(rulesFile), { recursive: true });
  fs.writeFileSync(cacheFile, "old-cache", "utf8");
  fs.writeFileSync(pendingFile, "old-pending-login", "utf8");
  fs.writeFileSync(downloadFile, "download", "utf8");
  fs.writeFileSync(rulesFile, "rules", "utf8");

  const preparation = prepareAuthProfile(config, scopes);

  assert.deepEqual(preparation, { migrationRequired: true, authStateCleared: true });
  assert.equal(fs.existsSync(cacheFile), false);
  assert.equal(fs.existsSync(pendingFile), false);
  assert.equal(fs.readFileSync(downloadFile, "utf8"), "download");
  assert.equal(fs.readFileSync(rulesFile, "utf8"), "rules");
  assert.equal(JSON.parse(fs.readFileSync(getAuthProfilePath(config), "utf8")).migrationRequired, true);
});

test("an application or scope change requires one migration login", () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-auth-profile-change-"));
  const oldConfig = configFor(dataDir, "old-client");
  const newConfig = configFor(dataDir, "new-client");

  prepareAuthProfile(oldConfig, ["User.Read", "Mail.Read"]);
  markAuthProfileReady(oldConfig, ["User.Read", "Mail.Read"]);
  fs.writeFileSync(path.join(oldConfig.cacheDir, "msal-cache.json"), "old-cache", "utf8");

  const preparation = prepareAuthProfile(newConfig, scopes);
  assert.equal(preparation.migrationRequired, true);
  assert.equal(preparation.authStateCleared, true);

  markAuthProfileReady(newConfig, scopes);
  assert.deepEqual(prepareAuthProfile(newConfig, scopes), {
    migrationRequired: false,
    authStateCleared: false
  });
});

test("fresh setup and logout do not report an application migration", () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-auth-profile-fresh-"));
  const config = configFor(dataDir);

  assert.deepEqual(prepareAuthProfile(config, scopes), {
    migrationRequired: false,
    authStateCleared: false
  });

  resetAuthProfileAfterLogout(config, scopes);
  assert.equal(JSON.parse(fs.readFileSync(getAuthProfilePath(config), "utf8")).migrationRequired, false);
});
