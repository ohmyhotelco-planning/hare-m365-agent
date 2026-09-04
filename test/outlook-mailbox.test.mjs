import assert from "node:assert/strict";
import test from "node:test";
import { getScopeList } from "../dist/auth.js";
import {
  mailboxKey,
  mailboxPath,
  resolveMailboxTarget,
  runMailboxRead,
  selfMailboxTarget
} from "../dist/outlook-mailbox.js";

test("shared mailbox scopes are requested without adding shared send permission", () => {
  const scopes = getScopeList();
  assert.ok(scopes.includes("Mail.Read.Shared"));
  assert.ok(scopes.includes("User.ReadBasic.All"));
  assert.equal(scopes.includes("Mail.Send.Shared"), false);
});

test("an exact mailbox address resolves to the directory user id", async () => {
  const calls = [];
  const target = await resolveMailboxTarget(config(), " CTO@example.com ", {
    async get(url, options) {
      calls.push({ url, options });
      return {
        value: [{
          id: "shared-user-id",
          displayName: "CTO",
          mail: "CTO@example.com",
          userPrincipalName: "cto-mailbox@tenant.onmicrosoft.com"
        }]
      };
    }
  });

  const url = new URL(`https://graph.microsoft.com/v1.0${calls[0].url}`);
  assert.equal(url.searchParams.get("$filter"), "mail eq 'CTO@example.com' or userPrincipalName eq 'CTO@example.com'");
  assert.equal(url.searchParams.get("$count"), "true");
  assert.equal(calls[0].options.headers.ConsistencyLevel, "eventual");
  assert.equal(target.scope, "shared");
  assert.equal(target.address, "CTO@example.com");
  assert.equal(target.userId, "shared-user-id");
  assert.equal(mailboxPath(target, "/messages"), "/users/shared-user-id/messages");
  assert.equal(mailboxKey(target), "shared:shared-user-id");
});

test("a unique exact mailbox display name resolves to its canonical id and address", async () => {
  const calls = [];
  const target = await resolveMailboxTarget(config(), "CTO", {
    async get(url, options) {
      calls.push({ url, options });
      return {
        value: [{
          id: "user-1",
          displayName: "CTO",
          mail: "cto@example.com",
          userPrincipalName: "cto-mailbox@tenant.onmicrosoft.com"
        }]
      };
    }
  });

  const url = new URL(`https://graph.microsoft.com/v1.0${calls[0].url}`);
  assert.equal(url.pathname, "/v1.0/users");
  assert.equal(url.searchParams.get("$search"), '"displayName:CTO"');
  assert.equal(url.searchParams.get("$count"), "true");
  assert.equal(calls[0].options.headers.ConsistencyLevel, "eventual");
  assert.equal(target.displayName, "CTO");
  assert.equal(target.address, "cto@example.com");
  assert.equal(mailboxPath(target, "messages"), "/users/user-1/messages");
});

test("mailbox name resolution fails closed for zero or multiple exact candidates", async () => {
  await assert.rejects(
    () => resolveMailboxTarget(config(), "Missing", { async get() { return { value: [] }; } }),
    /SHARED_MAILBOX_NOT_FOUND/
  );

  await assert.rejects(
    () => resolveMailboxTarget(config(), "CTO", {
      async get() {
        return {
          value: [
            { id: "1", displayName: "CTO", mail: "cto-office@example.com" },
            { id: "2", displayName: "CTO", mail: "cto-shared@example.com" }
          ]
        };
      }
    }),
    /SHARED_MAILBOX_AMBIGUOUS.*cto-office@example\.com.*cto-shared@example\.com/
  );
});

test("partial names and address local-parts are never selected automatically", async () => {
  await assert.rejects(
    () => resolveMailboxTarget(config(), "CT", {
      async get() {
        return { value: [{ id: "1", displayName: "CTO", mail: "cto@example.com" }] };
      }
    }),
    /SHARED_MAILBOX_NAME_NOT_EXACT.*cto@example\.com/
  );

  await assert.rejects(
    () => resolveMailboxTarget(config(), "cto", {
      async get() {
        return { value: [{ id: "1", displayName: "CTO Office", mail: "cto@example.com" }] };
      }
    }),
    /SHARED_MAILBOX_NAME_NOT_EXACT/
  );
});

test("directory and shared mailbox access failures are classified without fallback", async () => {
  await assert.rejects(
    () => resolveMailboxTarget(config(), "CTO", {
      async get() { throw new Error("Graph GET failed (403 Forbidden): denied"); }
    }),
    /MAILBOX_DIRECTORY_LOOKUP_DENIED/
  );

  const shared = await sharedMailbox("cto@example.com", "shared-user-id");
  await assert.rejects(
    () => runMailboxRead(shared, async () => {
      throw new Error("Graph GET failed (403 Forbidden): denied");
    }),
    /SHARED_MAILBOX_ACCESS_DENIED.*cto@example\.com/
  );
  await assert.rejects(
    () => runMailboxRead(shared, async () => {
      throw new Error("Graph GET failed (404 Not Found): missing");
    }),
    /SHARED_MAILBOX_NOT_FOUND.*cto@example\.com/
  );

  const original = new Error("Graph GET failed (403 Forbidden): denied");
  await assert.rejects(
    () => runMailboxRead(selfMailboxTarget(), async () => { throw original; }),
    (error) => error === original
  );
});

async function sharedMailbox(address, id) {
  return resolveMailboxTarget(config(), address, {
    async get() {
      return {
        value: [{
          id,
          displayName: "Shared mailbox",
          mail: address,
          userPrincipalName: `${id}@tenant.onmicrosoft.com`
        }]
      };
    }
  });
}

function config() {
  return {};
}
