import assert from "node:assert/strict";
import test from "node:test";
import {
  countMailboxMessages,
  listFlaggedMessages,
  listRecentMailbox,
  searchMailbox
} from "../dist/outlook.js";
import { resolveMailboxTarget } from "../dist/outlook-mailbox.js";

const scopes = ["all", "all-with-deleted", "deleted"];
const operations = ["recent", "flagged", "search", "count"];
const since = "2026-07-01";
const until = "2026-07-22";
const startDateTime = "2026-06-30T15:00:00.000Z";
const endDateTimeExclusive = "2026-07-22T15:00:00.000Z";
const active = mail("active", "inbox-id", "2026-07-20T00:00:00Z");
const deletedFirst = mail("deleted-first", "deleted-id", "2026-07-21T00:00:00Z");
const deletedLast = mail("deleted-last", "deleted-id", "2026-07-19T00:00:00Z");

for (const operation of operations) {
  test(`${operation}: scope routing, exclusion, paging and received dates for self/shared`, async () => {
    for (const target of [undefined, await sharedMailbox()]) {
      for (const scope of scopes) {
        const base = target ? "/users/shared-fixture-id" : "/me";
        const pages = scope === "deleted"
          ? [[deletedFirst], [deletedLast]]
          : [[deletedFirst], [active, deletedLast]];
        const client = pagedClient(base, scope, pages);
        let result = await invoke(operation, scope, target, client);
        const initial = metadata(operation, result);
        assert.equal(initial.folderScope, scope);
        assert.equal(initial.excludedDeletedItems, scope === "all");
        assert.equal(initial.mailbox.scope, target ? "shared" : "self");
        if (target) assert.equal(initial.mailbox.address, "shared@example.com");
        let messages = result.messages;

        if (operation === "search") {
          assert.equal(initial.continuationAvailable, true);
          assert.equal(initial.returnedCount, scope === "all" ? 0 : 1);
          assert.equal(client.messageCalls.length, 1);
          result = await invoke(operation, scope, target, client, { cursor: initial.nextCursor });
          messages = [...messages, ...result.messages];
          assert.equal(result.search.continuationAvailable, false);
          assert.equal(result.search.nextCursor, undefined);
          assert.equal(result.search.partialResult, false);
        }

        const expected = scope === "all" ? [active]
          : scope === "deleted" ? [deletedFirst, deletedLast]
          : [deletedFirst, active, deletedLast];
        if (operation === "count") {
          assert.equal(result.count.scannedCount, expected.length);
          assert.equal(result.count.matchedCount, expected.length);
          assert.equal(result.count.complete, true);
          assert.equal(result.count.continuationAvailable, false);
          assert.equal(result.count.dateProperty, "receivedDateTime");
          assert.equal(result.count.earliestDateTime, expected.at(-1).receivedDateTime);
          assert.equal(result.count.latestDateTime, expected[0].receivedDateTime);
          assert.deepEqual(result.breakdownBySender, [{ sender: "sender@example.com", count: expected.length }]);
        } else {
          assert.deepEqual(messages.map((item) => item.id), expected.map((item) => item.id));
          assert.ok(messages.every((item) => item.flagStatus === "flagged"));
          if (operation === "flagged") assert.equal(result.flagged.limitReached, false);
        }

        assert.equal(client.messageCalls.length, 2);
        assert.equal(client.messageCalls[1], client.nextUrl);
        assert.equal(client.metadataCalls.length, scope === "all" ? (operation === "search" ? 2 : 1) : 0);
        assertReceivedQuery(operation, graphUrl(client.messageCalls[0]));
      }
    }
  });
}

test("count: shared all-with-deleted timeout resumes cumulative totals without replay", async () => {
  const target = await sharedMailbox();
  let now = 0;
  const client = pagedClient("/users/shared-fixture-id", "all-with-deleted", [[deletedFirst, active], [deletedLast]]);
  const get = client.get.bind(client);
  let interrupt = true;
  client.get = async (path) => {
    if (path === client.nextUrl && interrupt) {
      interrupt = false;
      now = 6;
      const error = new Error("Synthetic request timed out");
      error.name = "GraphTimeoutError";
      throw error;
    }
    return get(path);
  };
  const partial = await invoke("count", "all-with-deleted", target, client, {
    timeBudgetMs: 5, now: () => now
  });
  assert.equal(interrupt, false);
  assert.equal(partial.count.partialResult, true);
  assert.equal(partial.count.partialReason, "time-budget-exceeded");
  assert.equal(partial.count.complete, false);
  assert.equal(partial.count.continuationAvailable, true);
  assert.equal(partial.count.scannedCount, 2);
  assert.equal(partial.count.matchedCount, 2);
  assert.equal(decodeCursor(partial.count.nextCursor).nextUrl, client.nextUrl);

  const result = await invoke("count", "all-with-deleted", target, client, { cursor: partial.count.nextCursor });
  assert.equal(result.count.continuationMode, "cumulative");
  assert.equal(result.count.complete, true);
  assert.equal(result.count.partialResult, false);
  assert.equal(result.count.scannedCount, 3);
  assert.equal(result.count.matchedCount, 3);
  assert.equal(result.count.earliestDateTime, deletedLast.receivedDateTime);
  assert.equal(result.count.latestDateTime, deletedFirst.receivedDateTime);
  assert.equal(result.count.nextCursor, undefined);
  assert.deepEqual(result.breakdownBySender, [{ sender: "sender@example.com", count: 3 }]);
  assert.equal(client.messageCalls.length, 2);
  assert.equal(client.metadataCalls.length, 0);
});

for (const operation of ["search", "count"]) {
  test(`${operation}: scope-bound cursors reject every cross-scope resume before client access`, async () => {
    for (const sourceScope of scopes) {
      let now = 0;
      const client = pagedClient("/me", sourceScope, [[deletedFirst], [deletedLast]]);
      const get = client.get.bind(client);
      client.get = async (path) => {
        const page = await get(path);
        if (page["@odata.nextLink"]) now = 6;
        return page;
      };
      const first = await invoke(operation, sourceScope, undefined, client, { timeBudgetMs: 5, now: () => now });
      const cursor = metadata(operation, first).nextCursor;
      assert.ok(cursor);
      assert.equal(JSON.parse(decodeCursor(cursor).criteriaKey).folderScope, sourceScope);
      for (const destination of scopes.filter((scope) => scope !== sourceScope)) {
        let calls = 0;
        await assert.rejects(
          () => invoke(operation, destination, undefined, {
            async get() { calls += 1; throw new Error("Unexpected client access"); }
          }, { cursor }),
          /cursor does not match/
        );
        assert.equal(calls, 0);
      }
    }
  });

  test(`${operation}: legacy all cursor shape and criteria remain usable`, async () => {
    let now = 0;
    const client = pagedClient("/me", "all", [[active, deletedFirst], [deletedLast, active]]);
    const get = client.get.bind(client);
    client.get = async (path) => {
      const page = await get(path);
      if (page["@odata.nextLink"]) now = 6;
      return page;
    };
    const first = await invoke(operation, "all", undefined, client, { timeBudgetMs: 5, now: () => now });
    // Construct the pre-change wire format independently of the production encoder.
    const criteriaKey = JSON.stringify(operation === "search" ? {
      query: "fixture", mailbox: "me", folderScope: "all", startDateTime, endDateTimeExclusive
    } : {
      subjectNeedle: "fixture", mailbox: "me", folderScope: "all", startDateTime,
      endDateTimeExclusive, dateProperty: "receivedDateTime"
    });
    const legacy = { version: 1, nextUrl: client.nextUrl, criteriaKey };
    if (operation === "count") Object.assign(legacy, {
      scannedCount: 1, matchedCount: 1,
      earliestDateTime: active.receivedDateTime, latestDateTime: active.receivedDateTime,
      senderCounts: [["sender@example.com", 1]]
    });
    assert.deepEqual(decodeCursor(metadata(operation, first).nextCursor), legacy);
    const cursor = Buffer.from(JSON.stringify(legacy), "utf8").toString("base64url");
    const resumed = await invoke(operation, "all", undefined, client, { cursor });
    assert.equal(metadata(operation, resumed).continuationAvailable, false);
    assert.equal(metadata(operation, resumed).excludedDeletedItems, true);
    if (operation === "search") assert.deepEqual(resumed.messages.map((item) => item.id), ["active"]);
    else {
      assert.equal(resumed.count.complete, true);
      assert.equal(resumed.count.scannedCount, 2);
      assert.equal(resumed.count.matchedCount, 2);
      assert.deepEqual(resumed.breakdownBySender, [{ sender: "sender@example.com", count: 2 }]);
    }
    assert.equal(client.messageCalls[1], client.nextUrl);
  });
}

for (const operation of operations) {
  test(`${operation}: shared errors propagate without self, recovery or broad-scope fallback`, async () => {
    const target = await sharedMailbox();
    for (const scope of scopes) {
      const expectedPath = scope === "all"
        ? "/v1.0/users/shared-fixture-id/mailFolders/deleteditems"
        : `/v1.0/users/shared-fixture-id${messagePath(scope)}`;
      const error = new Error(`Graph GET failed (${scope === "deleted" ? 404 : 403}): synthetic denial`);
      const calls = [];
      await assert.rejects(() => invoke(operation, scope, target, {
        async get(path) {
          calls.push(path);
          assert.equal(graphUrl(path).pathname, expectedPath);
          throw error;
        }
      }), (actual) => actual === error);
      assert.equal(calls.length, 1);
      if (scope === "all") assert.equal(graphUrl(calls[0]).search, "?$select=id");
    }
  });
}

function invoke(operation, scope, mailbox, client, options = {}) {
  if (operation === "recent") return listRecentMailbox(config(), scope, 10, mailbox, client);
  if (operation === "flagged") return listFlaggedMessages(config(), since, until, scope, 10, mailbox, client);
  if (operation === "search") return searchMailbox(config(), "fixture", since, until, scope, 10, { mailbox, ...options }, client);
  return countMailboxMessages(config(), "fixture", undefined, since, until, scope, { mailbox, ...options }, client);
}

function metadata(operation, result) {
  return result[operation === "recent" ? "list" : operation];
}

function pagedClient(base, scope, pages) {
  const path = `/v1.0${base}${messagePath(scope)}`;
  const metadataPath = `/v1.0${base}/mailFolders/deleteditems`;
  const nextUrl = `https://graph.microsoft.com${path}?$skiptoken=fixture-page-2`;
  const metadataCalls = [];
  const messageCalls = [];
  return {
    metadataCalls, messageCalls, nextUrl,
    async get(raw) {
      const url = graphUrl(raw);
      assert.equal(url.origin, "https://graph.microsoft.com");
      assert.doesNotMatch(url.pathname, /recoverable|purges|restore|permanent/i);
      if (url.pathname === metadataPath) {
        assert.equal(scope, "all", "Only default all should resolve Deleted Items metadata");
        assert.equal(url.search, "?$select=id");
        metadataCalls.push(raw);
        return { id: "deleted-id" };
      }
      assert.equal(url.pathname, path, "No other mailbox, folder or recovery endpoint is allowed");
      messageCalls.push(raw);
      assert.ok(messageCalls.length <= 2, "Unexpected page replay");
      if (messageCalls.length === 2) {
        assert.equal(raw, nextUrl);
        return { value: pages[1] };
      }
      assert.equal(url.searchParams.has("$skiptoken"), false);
      return { value: pages[0], "@odata.nextLink": nextUrl };
    }
  };
}

function assertReceivedQuery(operation, url) {
  if (operation === "search") {
    assert.equal(url.searchParams.get("$search"), `"fixture AND received>=${since} AND received<=${until}"`);
    return;
  }
  assert.equal(url.searchParams.get("$orderby"), "receivedDateTime desc");
  if (operation !== "recent") {
    const flag = operation === "flagged" ? " and flag/flagStatus eq 'flagged'" : "";
    assert.equal(url.searchParams.get("$filter"), `receivedDateTime ge ${startDateTime} and receivedDateTime lt ${endDateTimeExclusive}${flag}`);
  }
}

function messagePath(scope) {
  return scope === "deleted" ? "/mailFolders/deleteditems/messages" : "/messages";
}

function graphUrl(path) {
  return new URL(path.startsWith("/") ? `https://graph.microsoft.com/v1.0${path}` : path);
}

function decodeCursor(cursor) {
  assert.equal(typeof cursor, "string");
  return JSON.parse(Buffer.from(cursor, "base64url").toString("utf8"));
}

function mail(id, parentFolderId, receivedDateTime) {
  return {
    id, parentFolderId, receivedDateTime,
    sentDateTime: "2025-01-01T00:00:00Z",
    subject: "fixture message",
    from: { emailAddress: { address: "sender@example.com" } },
    flag: { flagStatus: "flagged" },
    body: { contentType: "text", content: "fixture body" }
  };
}

function config() {
  return {
    timeZone: "Asia/Seoul",
    policy: { defaultSearchLookbackDays: 90, maxSearchResults: 1000, maxMailFetchLimit: 20 }
  };
}

async function sharedMailbox() {
  return resolveMailboxTarget(config(), "shared@example.com", {
    async get(path) {
      assert.equal(graphUrl(path).pathname, "/v1.0/users");
      return { value: [{
        id: "shared-fixture-id", displayName: "Shared fixture",
        mail: "shared@example.com", userPrincipalName: "shared@example.com"
      }] };
    }
  });
}
