import assert from "node:assert/strict";
import test from "node:test";
import {
  FIXED_HOST_DATA_DIRS,
  buildBlockedSetupContract,
  buildSetupContract,
  determineSetupState
} from "../dist/setup-state.js";

const readySnapshot = {
  configured: true,
  dataDirPersistent: true,
  loggedIn: true,
  tokenUsable: true,
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
    [{ ...readySnapshot, dataDirPersistent: false }, "CONNECT_FIXED_FOLDER"],
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

test("fixed host folders avoid Documents and OneDrive", () => {
  assert.deepEqual(FIXED_HOST_DATA_DIRS, {
    windows: "%USERPROFILE%\\HareM365Agent",
    mac: "~/HareM365Agent"
  });
  assert.doesNotMatch(JSON.stringify(FIXED_HOST_DATA_DIRS), /Documents|OneDrive/i);
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

test("blocked setup contract reports one blocker and stops", () => {
  const contract = buildBlockedSetupContract("npm ci failed");
  assert.equal(contract.state, "BLOCKED");
  assert.equal(contract.nextAction, "REPORT_BLOCKER");
  assert.equal(contract.stopAfterAction, true);
  assert.match(contract.instruction, /npm ci failed/);
});
