import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { listChatMessages, searchChatMessages } from "../dist/teams.js";

function configFor() {
  return {
    dataDir: path.join(os.tmpdir(), "hare-teams-full-body"),
    timeZone: "Asia/Seoul",
    policy: {
      defaultSearchLookbackDays: 90,
      maxSearchResults: 1000,
      maxTeamsFetchLimit: 50
    }
  };
}

test("chat-messages returns complete text and original HTML without a 500-character cut", async () => {
  const longText = "A".repeat(800);
  const bodyHtml = `<div>${longText}</div>`;
  const client = {
    async get(url) {
      assert.match(url, /\/chats\/chat-1\/messages/);
      return {
        value: [{
          id: "message-1",
          body: { contentType: "html", content: bodyHtml }
        }]
      };
    },
    async post() {
      throw new Error("Unexpected POST");
    }
  };

  const messages = await listChatMessages(configFor(), "chat-1", 1, client);
  assert.equal(messages[0].body, longText);
  assert.equal(messages[0].bodyHtml, bodyHtml);
  assert.equal(messages[0].bodyPreview, longText);
  assert.equal(messages[0].body.length, 800);
});

test("search-messages resolves complete bodies for chat and channel hits", async () => {
  const chatText = "Chat ".repeat(140);
  const channelText = "Channel ".repeat(100);
  const getUrls = [];
  const client = {
    async post(url) {
      assert.equal(url, "/search/query");
      return {
        value: [{
          hitsContainers: [{
            total: 3,
            moreResultsAvailable: false,
            hits: [
              { summary: "chat snippet", resource: { id: "m1", chatId: "c1" } },
              {
                summary: "channel snippet",
                resource: {
                  id: "m2",
                  channelIdentity: { teamId: "t1", channelId: "ch1" }
                }
              },
              { summary: "unlocated snippet", resource: { id: "m3" } }
            ]
          }]
        }]
      };
    },
    async get(url) {
      getUrls.push(url);
      if (url === "/chats/c1?$select=id,topic") return { id: "c1", topic: "Chat topic" };
      if (url === "/chats/c1/messages/m1") {
        return { id: "m1", body: { contentType: "html", content: `<p>${chatText}</p>` } };
      }
      if (url === "/teams/t1/channels/ch1/messages/m2") {
        return { id: "m2", body: { contentType: "html", content: `<div>${channelText}</div>` } };
      }
      throw new Error(`Unexpected GET ${url}`);
    }
  };

  const result = await searchChatMessages(
    configFor(),
    "keyword",
    "2026-07-01",
    "2026-07-14",
    10,
    {},
    client
  );

  assert.equal(result.search.fullBodyReturnedCount, 2);
  assert.equal(result.search.fullBodyUnavailableCount, 1);
  assert.equal(result.messages[0].body, chatText.trim());
  assert.equal(result.messages[0].bodyPreview, chatText.trim());
  assert.equal(result.messages[0].chatTopic, "Chat topic");
  assert.equal(result.messages[1].body, channelText.trim());
  assert.deepEqual(result.messages[1].channelIdentity, { teamId: "t1", channelId: "ch1" });
  assert.equal(result.messages[2].fullBodyAvailable, false);
  assert.equal(result.messages[2].bodyUnavailableReason, "missing-message-location");
  assert.equal(result.messages[2].bodyPreview, "unlocated snippet");
  assert.equal(getUrls.filter((url) => url === "/chats/c1?$select=id,topic").length, 1);
  assert.equal(getUrls.some((url) => url.startsWith("/me/chats")), false);
});

test("search-messages marks a failed detail request instead of presenting a snippet as full body", async () => {
  const client = {
    async post() {
      return {
        value: [{ hitsContainers: [{ total: 1, hits: [
          { summary: "partial only", resource: { id: "m1", chatId: "c1" } }
        ] }] }]
      };
    },
    async get(url) {
      if (url.startsWith("/me/chats")) return { value: [] };
      throw new Error("message deleted");
    }
  };

  const result = await searchChatMessages(
    configFor(),
    "keyword",
    "2026-07-01",
    "2026-07-14",
    10,
    {},
    client
  );

  assert.equal(result.search.fullBodyReturnedCount, 0);
  assert.equal(result.search.fullBodyUnavailableCount, 1);
  assert.equal(result.messages[0].fullBodyAvailable, false);
  assert.equal(result.messages[0].body, undefined);
  assert.equal(result.messages[0].searchSummary, "partial only");
  assert.equal(result.messages[0].bodyUnavailableReason, "detail-request-failed");
});

test("search-messages identifies a detail response with no message body", async () => {
  const client = {
    async post() {
      return {
        value: [{ hitsContainers: [{ total: 1, hits: [
          { summary: "system event", resource: { id: "m1", chatId: "c1" } }
        ] }] }]
      };
    },
    async get(url) {
      if (url.startsWith("/me/chats")) return { value: [] };
      return { id: "m1", body: {} };
    }
  };

  const result = await searchChatMessages(
    configFor(),
    "keyword",
    "2026-07-01",
    "2026-07-14",
    10,
    {},
    client
  );

  assert.equal(result.search.fullBodyUnavailableCount, 1);
  assert.equal(result.messages[0].fullBodyAvailable, false);
  assert.equal(result.messages[0].bodyUnavailableReason, "message-body-missing");
});

test("search-messages returns partial results when the processing budget is exhausted", async () => {
  let currentTime = 0;
  const client = {
    async post() {
      return {
        value: [{ hitsContainers: [{
          total: 4,
          moreResultsAvailable: false,
          hits: Array.from({ length: 4 }, (_, index) => ({
            summary: `snippet ${index + 1}`,
            resource: { id: `m${index + 1}`, chatId: `c${index + 1}` }
          }))
        }] }]
      };
    },
    async get(url) {
      currentTime += 20;
      const match = url.match(/messages\/(m\d+)$/);
      if (match) return { id: match[1], body: { contentType: "html", content: `<p>${match[1]}</p>` } };
      return { id: url, topic: "Topic" };
    }
  };

  const result = await searchChatMessages(
    configFor(),
    "keyword",
    "2026-07-01",
    "2026-07-14",
    4,
    { timeBudgetMs: 30, now: () => currentTime },
    client
  );

  assert.equal(result.search.returnedCount, 4);
  assert.equal(result.search.partialResult, true);
  assert.equal(result.search.partialReason, "time-budget-exceeded");
  assert.ok(result.search.fullBodyReturnedCount < 4);
  assert.ok(result.messages.some((message) => message.bodyUnavailableReason === "time-budget-exceeded"));
});

test("search-messages applies an offset and reports the next page", async () => {
  const client = {
    async post(_url, body) {
      assert.equal(body.requests[0].from, 100);
      assert.equal(body.requests[0].size, 1);
      return {
        value: [{ hitsContainers: [{
          total: 250,
          moreResultsAvailable: true,
          hits: [{ summary: "next", resource: { id: "m101", chatId: "c1" } }]
        }] }]
      };
    },
    async get(url) {
      if (url === "/chats/c1/messages/m101") {
        return { id: "m101", body: { contentType: "html", content: "full" } };
      }
      if (url === "/chats/c1?$select=id,topic") return { id: "c1", topic: "Topic" };
      throw new Error(`Unexpected GET ${url}`);
    }
  };

  const result = await searchChatMessages(
    configFor(),
    "keyword",
    "2026-07-01",
    "2026-07-14",
    1,
    { offset: 100 },
    client
  );

  assert.equal(result.search.offset, 100);
  assert.equal(result.search.nextOffset, 101);
  assert.equal(result.search.continuationAvailable, true);
  assert.equal(result.messages[0].chatTopic, "Topic");
});
