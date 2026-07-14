import assert from "node:assert/strict";
import test from "node:test";
import {
  buildFlaggedMessagesPath,
  buildRecentMessagesPath,
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

function graphUrl(relativePath) {
  return new URL(`https://graph.microsoft.com/v1.0${relativePath}`);
}
