import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { spawnSync } from "node:child_process";
import { pathToFileURL } from "node:url";

const repoRoot = path.resolve(import.meta.dirname, "..");
const configModule = pathToFileURL(path.join(repoRoot, "dist", "config.js")).href;

function readConfig(cwd, env) {
  const script = `
    import { loadConfig } from ${JSON.stringify(configModule)};
    const config = loadConfig();
    process.stdout.write(JSON.stringify({ clientId: config.clientId, tenantId: config.tenantId }));
  `;
  const result = spawnSync(process.execPath, ["--input-type=module", "--eval", script], {
    cwd,
    encoding: "utf8",
    env
  });
  assert.equal(result.status, 0, result.stderr);
  return JSON.parse(result.stdout);
}

test("stale .env cannot override the bundled Azure application", () => {
  const cwd = fs.mkdtempSync(path.join(os.tmpdir(), "hare-stale-env-"));
  fs.writeFileSync(
    path.join(cwd, ".env"),
    "OMH_M365_CLIENT_ID=old-client-id\nOMH_M365_TENANT_ID=old-tenant-id\n",
    "utf8"
  );

  const env = { ...process.env };
  delete env.OMH_M365_CLIENT_ID;
  delete env.OMH_M365_TENANT_ID;
  const actual = readConfig(cwd, env);
  const expected = JSON.parse(fs.readFileSync(path.join(repoRoot, "hare.config.json"), "utf8"));
  assert.equal(actual.clientId, expected.clientId);
  assert.equal(actual.tenantId, expected.tenantId);
});

test("explicit process environment overrides bundled configuration", () => {
  const actual = readConfig(repoRoot, {
    ...process.env,
    OMH_M365_CLIENT_ID: "explicit-client",
    OMH_M365_TENANT_ID: "explicit-tenant"
  });

  assert.deepEqual(actual, {
    clientId: "explicit-client",
    tenantId: "explicit-tenant"
  });
});
