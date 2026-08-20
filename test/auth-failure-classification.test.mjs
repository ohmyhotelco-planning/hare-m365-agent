import assert from "node:assert/strict";
import test from "node:test";
import { classifyAuthFailure } from "../dist/setup-state.js";

test("auth failures distinguish network blockers from real login requirements", () => {
  for (const reason of [
    "TOKEN_ACQUISITION_FAILED: network_error",
    "TOKEN_ACQUISITION_FAILED: fetch failed",
    "TOKEN_ACQUISITION_FAILED: getaddrinfo EAI_AGAIN login.microsoftonline.com",
    "TOKEN_ACQUISITION_FAILED: ETIMEDOUT"
  ]) {
    assert.equal(classifyAuthFailure(reason), "NETWORK_BLOCKED", reason);
  }

  for (const reason of [
    "NO_ACCOUNT_IN_CACHE",
    "NO_ACCESS_TOKEN",
    "AUTH_APP_CHANGED",
    "TOKEN_ACQUISITION_FAILED: interaction_required",
    "TOKEN_ACQUISITION_FAILED: invalid_grant"
  ]) {
    assert.equal(classifyAuthFailure(reason), "LOGIN_REQUIRED", reason);
  }

  assert.equal(
    classifyAuthFailure("TOKEN_ACQUISITION_FAILED: unexpected response"),
    "UNKNOWN_BLOCKED"
  );
});
