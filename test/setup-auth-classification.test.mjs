import assert from "node:assert/strict";
import test from "node:test";
import { buildSetupContract, determineSetupState } from "../dist/setup-state.js";

const baseSnapshot = {
  configured: true,
  dataDirPersistent: true,
  loggedIn: false,
  tokenUsable: false,
  authMigrationRequired: false,
  pendingLoginStateExists: false
};

test("cached account token network failures do not start a new login", () => {
  const snapshot = {
    ...baseSnapshot,
    authReason: "TOKEN_ACQUISITION_FAILED: network_error"
  };
  assert.equal(determineSetupState(snapshot), "BLOCKED");
  const contract = buildSetupContract(snapshot, "node dist/cli.js");
  assert.equal(contract.nextAction, "REPORT_BLOCKER");
  assert.equal(contract.nextCommand, undefined);
  assert.doesNotMatch(JSON.stringify(contract), /auth login-start/);
  assert.match(contract.instruction, /network_error/);
});

test("pending device login does not override a token-check network failure", () => {
  const snapshot = {
    ...baseSnapshot,
    pendingLoginStateExists: true,
    authReason: "TOKEN_ACQUISITION_FAILED: network_error"
  };
  assert.equal(determineSetupState(snapshot), "BLOCKED");
  assert.equal(
    buildSetupContract(snapshot, "node dist/cli.js").nextCommand,
    undefined
  );
});
test("missing account still starts login", () => {
  const snapshot = {
    ...baseSnapshot,
    authReason: "NO_ACCOUNT_IN_CACHE"
  };
  assert.equal(determineSetupState(snapshot), "LOGIN_START_REQUIRED");
  assert.equal(
    buildSetupContract(snapshot, "node dist/cli.js").nextCommand,
    "node dist/cli.js auth login-start"
  );
});

test("application migration blocks rather than automatically starting login", () => {
  const snapshot = {
    ...baseSnapshot,
    authMigrationRequired: true,
    authReason: "AUTH_APP_CHANGED"
  };
  assert.equal(determineSetupState(snapshot), "BLOCKED");
});

test("pending state and migration flags cannot override hard or unknown blockers", () => {
  for (const authReason of ["AUTH_APP_CHANGED", "AUTH_TENANT_CHANGED", "AUTH_MIGRATION_REQUIRED",
    "AUTH_PROFILE_INVALID", "AUTH_PROFILE_MISSING", "AUTH_PROFILE_UNREADABLE",
    "AUTH_ACCOUNT_SELECTION_REQUIRED", "AUTH_ACCOUNT_MISMATCH", "AUTH_CACHE_CHANGED",
    "AUTH_SCOPES_INSUFFICIENT", "NO_ACCESS_TOKEN", "TOKEN_ACQUISITION_FAILED: unclassified_error"]) {
    for (const authMigrationRequired of [true, false]) {
      const contract = buildSetupContract({ ...baseSnapshot, authReason,
        authMigrationRequired, pendingLoginStateExists: true }, "hare");
      assert.equal(contract.state, "BLOCKED", authReason);
      assert.equal(contract.nextCommand, undefined);
    }
  }
});
