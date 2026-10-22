import assert from "node:assert/strict";
import test from "node:test";
import { authFailure, copyAuthProvenance, readAuthProvenance } from "../dist/auth-diagnostics.js";
import { groupReadError, retryableGroupReadFailure } from "../dist/groups-common.js";

test("auth diagnostics only copy fixed categories and leave messages unchanged", () => {
  const error = authFailure("TOKEN_ACQUISITION_FAILED: unclassified_error", "silent",
    { message: "PRIVATE_MESSAGE", stack: "PRIVATE_STACK", cause: "PRIVATE_CAUSE", errorCode: "server_error", code: "EPERM", path: "PRIVATE_PATH" });
  assert.equal(error.message, "TOKEN_ACQUISITION_FAILED: unclassified_error");
  assert.deepEqual(readAuthProvenance(error), { phase: "silent", reason: "TOKEN_ACQUISITION_FAILED", errorCode: "server_error", ioCode: "EPERM" });
  const wrapped = groupReadError("POST_SUBJECT", copyAuthProvenance(new Error(`Remote failure: ${error.message}`), error));
  assert.match(wrapped.message, /^GROUP_POST_SUBJECT_AUTH_BLOCKED:/);
  assert.deepEqual(readAuthProvenance(wrapped), readAuthProvenance(error));
  assert.doesNotMatch(JSON.stringify(wrapped), /PRIVATE_/);
});

test("untrusted or unknown diagnostic fields cannot escape the allowlist", () => {
  for (const authProvenance of [null, "PRIVATE", { phase: "PRIVATE", reason: "AUTH_CACHE_CHANGED" }, { phase: "commit", reason: "PRIVATE" }]) {
    assert.equal(readAuthProvenance({ authProvenance }), undefined);
  }
  const copied = copyAuthProvenance(new Error("safe"), { authProvenance: { phase: "commit", reason: "AUTH_CACHE_CHANGED",
    errorCode: "PRIVATE", ioCode: "PRIVATE", path: "PRIVATE", token: "PRIVATE" } });
  assert.deepEqual(readAuthProvenance(copied), { phase: "commit", reason: "AUTH_CACHE_CHANGED" });
  assert.doesNotMatch(JSON.stringify(copied), /PRIVATE/);
  assert.equal(readAuthProvenance(new Error("unknown")), undefined);
});

test("unverified token connectivity is not an authentication rejection and does not auto-retry", () => {
  const original = authFailure("AUTH_CHECK_BLOCKED: private network details", "silent", { code: "ETIMEDOUT" });
  const safe = groupReadError("POST_SUBJECT", copyAuthProvenance(new Error(`Remote failure: ${original.message}`), original));
  assert.match(safe.message, /^GROUP_POST_SUBJECT_NETWORK_BLOCKED:/);
  assert.equal(retryableGroupReadFailure(safe.message), false);
  assert.equal(readAuthProvenance(safe).ioCode, "ETIMEDOUT");
  assert.doesNotMatch(JSON.stringify(safe), /private network details/);
  assert.match(groupReadError("POST_SUBJECT", new Error("Graph GET failed (401 Unauthorized): private")).message, /AUTH_BLOCKED/);
  assert.doesNotMatch(groupReadError("POST_SUBJECT", new Error("Graph GET failed (409 Conflict): TOKEN_SYNTHETIC")).message, /AUTH_BLOCKED|TOKEN_SYNTHETIC/);
});
