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

test("pending device login remains completable during a token-check network failure", () => {
  const snapshot = {
    ...baseSnapshot,
    pendingLoginStateExists: true,
    authReason: "TOKEN_ACQUISITION_FAILED: network_error"
  };
  assert.equal(determineSetupState(snapshot), "LOGIN_COMPLETE_REQUIRED");
  assert.equal(
    buildSetupContract(snapshot, "node dist/cli.js").nextCommand,
    "node dist/cli.js auth login-complete"
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

test("application migration still starts one-time login", () => {
  const snapshot = {
    ...baseSnapshot,
    authMigrationRequired: true,
    authReason: "AUTH_APP_CHANGED"
  };
  assert.equal(determineSetupState(snapshot), "LOGIN_START_REQUIRED");
});
