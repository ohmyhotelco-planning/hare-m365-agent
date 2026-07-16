import assert from "node:assert/strict";
import test from "node:test";
import {
  buildBlockedSetupContract,
  buildSetupContract,
  determineSetupState
} from "../dist/setup-state.js";

const readySnapshot = {
  configured: true,
  dataDirPersistent: true,
  loggedIn: true,
  tokenUsable: true,
  authMigrationRequired: false,
  pendingLoginStateExists: false
};

test("setup state follows one deterministic priority order", () => {
  assert.equal(determineSetupState({ ...readySnapshot, configured: false }), "SETUP_REQUIRED");
  assert.equal(determineSetupState({ ...readySnapshot, dataDirPersistent: false }), "FOLDER_REQUIRED");
  assert.equal(determineSetupState(readySnapshot), "READY");
  assert.equal(
    determineSetupState({
      ...readySnapshot,
      loggedIn: false,
      tokenUsable: false,
      pendingLoginStateExists: true
    }),
    "LOGIN_COMPLETE_REQUIRED"
  );
  assert.equal(
    determineSetupState({
      ...readySnapshot,
      loggedIn: false,
      tokenUsable: false
    }),
    "LOGIN_START_REQUIRED"
  );
});

test("setup contract exposes exactly one next action for every state", () => {
  const cases = [
    [{ ...readySnapshot, configured: false }, "CHECK_CONFIGURATION"],
    [{ ...readySnapshot, dataDirPersistent: false }, "SELECT_PROJECT_FOLDER"],
    [
      { ...readySnapshot, loggedIn: false, tokenUsable: false },
      "RUN_LOGIN_START"
    ],
    [
      {
        ...readySnapshot,
        loggedIn: false,
        tokenUsable: false,
        pendingLoginStateExists: true
      },
      "WAIT_FOR_USER_THEN_RUN_LOGIN_COMPLETE"
    ],
    [readySnapshot, "WAIT_FOR_USER_REQUEST"]
  ];

  for (const [snapshot, expectedAction] of cases) {
    const contract = buildSetupContract(snapshot, "node dist/cli.js");
    assert.equal(contract.nextAction, expectedAction);
  }
});

test("missing project folder stops without folder automation", () => {
  const contract = buildSetupContract(
    { ...readySnapshot, dataDirPersistent: false },
    "node dist/cli.js"
  );
  assert.equal(contract.nextAction, "SELECT_PROJECT_FOLDER");
  assert.match(contract.instruction, /existing Hare project or folder/);
  assert.match(contract.instruction, /stop/i);
  assert.doesNotMatch(contract.instruction, /computer-use|folder-access|%USERPROFILE%|~\/HareM365Agent/);
});

test("pending login contract keeps the same device code flow", () => {
  const contract = buildSetupContract(
    {
      ...readySnapshot,
      loggedIn: false,
      tokenUsable: false,
      pendingLoginStateExists: true
    },
    "node dist/cli.js"
  );

  assert.equal(contract.nextCommand, "node dist/cli.js auth login-complete");
  assert.equal(contract.stopAfterAction, true);
  assert.doesNotMatch(JSON.stringify(contract), /login-start/);
});

test("application migration explains the one-time sign-in without changing the flow", () => {
  const contract = buildSetupContract(
    {
      ...readySnapshot,
      loggedIn: false,
      tokenUsable: false,
      authMigrationRequired: true
    },
    "node dist/cli.js"
  );

  assert.equal(contract.state, "LOGIN_START_REQUIRED");
  assert.equal(contract.nextCommand, "node dist/cli.js auth login-start");
  assert.match(contract.instruction, /updated to a new Microsoft application/);
  assert.match(contract.instruction, /one Microsoft sign-in/);
  assert.match(contract.instruction, /their own company Microsoft account/);
  assert.match(contract.instruction, /Never name, recommend, or preselect a specific email address/);
});

test("login start never recommends a specific account", () => {
  const contract = buildSetupContract(
    {
      ...readySnapshot,
      loggedIn: false,
      tokenUsable: false
    },
    "node dist/cli.js"
  );

  assert.equal(contract.state, "LOGIN_START_REQUIRED");
  assert.match(contract.instruction, /their own company Microsoft account/);
  assert.match(contract.instruction, /Never name, recommend, or preselect a specific email address/);
});

test("blocked setup contract reports one blocker and stops", () => {
  const contract = buildBlockedSetupContract("npm ci failed");
  assert.equal(contract.state, "BLOCKED");
  assert.equal(contract.nextAction, "REPORT_BLOCKER");
  assert.equal(contract.stopAfterAction, true);
  assert.match(contract.instruction, /npm ci failed/);
});
