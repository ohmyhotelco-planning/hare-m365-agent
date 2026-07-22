import assert from "node:assert/strict";
import test from "node:test";
import {
  buildFlaggedMessagesPath,
  buildRecentMessagesPath,
  countMailboxMessages,
  searchMailbox,
  toMailSummary
} from "../dist/outlook.js";

test("recent mail defaults can target the whole mailbox and include flag state", () => {
  const url = graphUrl(buildRecentMessagesPath("all", 10));
  assert.equal(url.pathname, "/v1.0/me/messages");
  assert.equal(url.searchParams.get("$top"), "10");
  assert.equal(url.searchParams.get("$orderby"), "receivedDateTime desc");
  assert.match(url.searchParams.get("$select") ?? "", /flag/);
  assert.match(url.searchParams.get("$select") ?? "", /parentFolderId/);
});

test("flagged mail query filters the whole mailbox by flag and date", () => {
  const range = {
    since: "2026-07-01",
    until: "2026-07-14",
    startDateTime: "2026-06-30T15:00:00.000Z",
    endDateTimeExclusive: "2026-07-14T15:00:00.000Z",
    timeZone: "Asia/Seoul",
    days: 14,
    usedDefaultLookback: false,
    notice: "fixture"
  };
  const url = graphUrl(buildFlaggedMessagesPath("all", range, 1000));
  const filter = url.searchParams.get("$filter") ?? "";
  assert.equal(url.pathname, "/v1.0/me/messages");
  assert.equal(url.searchParams.get("$top"), "100");
  assert.match(filter, /^receivedDateTime ge /);
  assert.match(filter, /flag\/flagStatus eq 'flagged'/);
  assert.match(filter, /receivedDateTime ge 2026-06-30T15:00:00\.000Z/);
  assert.match(filter, /receivedDateTime lt 2026-07-14T15:00:00\.000Z/);
});

test("mail summaries expose the Outlook flag status", () => {
  const summary = toMailSummary({
    id: "message-id",
    subject: "Flagged follow-up",
    flag: { flagStatus: "flagged" }
  });
  assert.equal(summary.flagStatus, "flagged");
});

test("mail search returns complete bodies and an opaque continuation cursor", async () => {
  const urls = [];
  const client = {
    async get(url) {
      if (url.includes("deleteditems")) return { id: "deleted" };
      urls.push(url);
      if (urls.length === 1) {
        return {
          value: [
            {
              id: "m1",
              subject: "Reservation details",
              bodyPreview: "short preview",
              body: { contentType: "html", content: "<p>Complete reservation details</p>" }
            }
          ],
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=100"
        };
      }
      return { value: [{ id: "m2", body: { contentType: "html", content: "<p>Next page</p>" } }] };
    }
  };

  const first = await searchMailbox(
    config(),
    "reservation",
    "2026-07-01",
    "2026-07-22",
    "all",
    100,
    {},
    client
  );
  assert.equal(first.messages[0].body, "Complete reservation details");
  assert.equal(first.messages[0].bodyHtml, "<p>Complete reservation details</p>");
  assert.equal(first.messages[0].fullBodyAvailable, true);
  assert.equal(first.search.fullBodyUnavailableCount, 0);
  assert.equal(first.search.continuationAvailable, true);
  assert.ok(first.search.nextCursor);

  const second = await searchMailbox(
    config(),
    "reservation",
    "2026-07-01",
    "2026-07-22",
    "all",
    100,
    { cursor: first.search.nextCursor },
    client
  );
  assert.equal(urls[1], "https://graph.microsoft.com/v1.0/me/messages?$skip=100");
  assert.equal(second.messages[0].body, "Next page");
  assert.equal(second.search.continuationAvailable, false);

  await assert.rejects(
    () => searchMailbox(
      config(),
      "different query",
      "2026-07-01",
      "2026-07-22",
      "all",
      100,
      { cursor: first.search.nextCursor },
      client
    ),
    /cursor does not match/
  );
});

test("mail search returns a resumable partial result when its time budget expires", async () => {
  let now = 0;
  const client = {
    async get() {
      now = 6;
      const error = new Error("Graph request exceeded the 5ms time budget.");
      error.name = "GraphTimeoutError";
      throw error;
    }
  };

  const result = await searchMailbox(
    config(),
    "reservation",
    "2026-07-01",
    "2026-07-22",
    "all",
    100,
    { timeBudgetMs: 5, now: () => now },
    client
  );
  assert.equal(result.search.partialResult, true);
  assert.equal(result.search.partialReason, "time-budget-exceeded");
  assert.equal(result.search.continuationAvailable, true);
  assert.ok(result.search.nextCursor);
});

test("mail count marks an interrupted scan as partial and resumable", async () => {
  let now = 0;
  let calls = 0;
  const client = {
    async get(url) {
      calls += 1;
      if (url.includes("deleteditems")) return { id: "deleted" };
      if (calls === 2) {
        now = 4;
        return {
          value: [{ id: "m1", subject: "[RPA] first", receivedDateTime: "2026-07-20T00:00:00Z" }],
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=500"
        };
      }
      now = 6;
      const error = new Error("Graph request exceeded the 5ms time budget.");
      error.name = "GraphTimeoutError";
      throw error;
    }
  };

  const result = await countMailboxMessages(
    config(),
    "[RPA]",
    undefined,
    "2026-07-01",
    "2026-07-22",
    "all",
    { timeBudgetMs: 5, now: () => now },
    client
  );
  assert.equal(result.count.complete, false);
  assert.equal(result.count.partialResult, true);
  assert.equal(result.count.matchedCount, 1);
  assert.equal(result.count.continuationAvailable, true);
  assert.ok(result.count.nextCursor);
  assert.equal(result.count.continuationMode, "cumulative");

  const resumed = await countMailboxMessages(
    config(),
    "[RPA]",
    undefined,
    "2026-07-01",
    "2026-07-22",
    "all",
    { cursor: result.count.nextCursor },
    {
      async get(url) {
        if (url.includes("deleteditems")) return { id: "deleted" };
        return {
          value: [{ id: "m2", subject: "[RPA] second", receivedDateTime: "2026-07-19T00:00:00Z" }]
        };
      }
    }
  );
  assert.equal(resumed.count.complete, true);
  assert.equal(resumed.count.scannedCount, 2);
  assert.equal(resumed.count.matchedCount, 2);
  assert.equal(resumed.count.continuationAvailable, false);
});

function config() {
  return {
    timeZone: "Asia/Seoul",
    policy: {
      defaultSearchLookbackDays: 90,
      maxSearchResults: 1000,
      maxMailFetchLimit: 20
    }
  };
}

function graphUrl(relativePath) {
  return new URL(`https://graph.microsoft.com/v1.0${relativePath}`);
}
