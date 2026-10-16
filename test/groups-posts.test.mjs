import assert from "node:assert/strict";
import test, { after } from "node:test";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { listGroupPosts } from "../dist/groups-posts.js";
import { groupPageUrl, groupReadError } from "../dist/groups-common.js";

const fixtureRoot = fs.mkdtempSync(path.join(os.tmpdir(), "hare-group-posts-"));
after(() => {
  assert.equal(path.dirname(path.resolve(fixtureRoot)), path.resolve(os.tmpdir()));
  assert.match(path.basename(fixtureRoot), /^hare-group-posts-/);
  fs.rmSync(fixtureRoot, { recursive: true, force: true });
});
const config = { cacheDir: path.join(fixtureRoot, "cache"), policy: { defaultSearchLookbackDays: 90, maxSearchResults: 1000 } };
const group = { id: "group-one", mail: "team@example.test", groupTypes: ["Unified"], classification: "microsoft365", evidence: ["groupTypes includes Unified"], mailEnabled: true, securityEnabled: false, exchangeRecipientType: "unverified" };
const base = "/v1.0/groups/group-one/conversations";
const options = { since: "2026-10-01", until: "2026-10-15", now: () => Date.parse("2026-10-15T01:00:00Z") };
const post = (id, extras = {}) => ({ id, createdDateTime: "2026-10-10T01:00:00Z", body: { contentType: "html", content: `<p>${id} ${"full body ".repeat(100)}</p>` }, hasAttachments: true, ...extras });
function clientFor(routes) {
  const calls = [];
  return {
    calls,
    async get(url, requestOptions) {
      const parsed = new URL(url.startsWith("/") ? `https://graph.microsoft.com/v1.0${url}` : url);
      const key = parsed.pathname + (parsed.searchParams.has("page") ? `?page=${parsed.searchParams.get("page")}` : "");
      calls.push({ key, url, requestOptions });
      assert.ok(requestOptions.totalTimeoutMs > 0);
      if (!(key in routes)) throw new Error(`Unexpected request ${key}`);
      const result = routes[key];
      if (result instanceof Error) throw result;
      return typeof result === "function" ? result() : result;
    }
  };
}
function oneThread(posts) {
  return { [base]: { value: [{ id: "c1", preview: "not a body", lastDeliveredDateTime: "2030-01-01T00:00:00Z" }] },
    [`${base}/c1/threads`]: { value: [{ id: "t1" }] },
    [`${base}/c1/threads/t1/posts`]: { value: posts } };
}

test("group posts traverse multiple conversations/threads/pages and retain complete body and chain", async () => {
  const routes = oneThread([post("p1")]);
  routes[base]["@odata.nextLink"] = `https://graph.microsoft.com${base}?page=2`;
  routes[`${base}?page=2`] = { value: [{ id: "c2" }] };
  routes[`${base}/c2/threads`] = { value: [{ id: "t2" }, { id: "t3" }] };
  routes[`${base}/c2/threads/t2/posts`] = { value: [post("p2", { conversationId: "c2", conversationThreadId: "t2" })] };
  routes[`${base}/c2/threads/t3/posts`] = { value: [post("p3", { from: { emailAddress: { address: "sender@example.test" } } })] };
  const client = clientFor(routes);
  const result = await listGroupPosts(config, group, 20, options, client);
  assert.equal(result.page.complete, true);
  assert.equal(result.posts.length, 3);
  assert.ok(result.posts[0].bodyContent.length > 900);
  assert.equal(result.posts[0].bodyHtml, routes[`${base}/c1/threads/t1/posts`].value[0].body.content);
  assert.equal(result.posts[1].threadId, "t2");
  assert.equal(result.posts[1].conversationId, "c2");
  assert.ok(client.calls.every(call => !/\$filter|\$orderby|\$top/.test(call.url)));
});

test("null conversationId retains the requested parent path and verifies the returned thread", async () => {
  const result = await listGroupPosts(config, group, 3, options,
    clientFor(oneThread([post("p1", { conversationId: null, conversationThreadId: "t1" })])));
  assert.equal(result.posts.length, 1);
  assert.equal(result.posts[0].conversationId, "c1");
  assert.equal(result.posts[0].threadId, "t1");
  assert.equal(result.posts[0].fullBodyAvailable, true);
  await assert.rejects(listGroupPosts(config, group, 3, options,
    clientFor(oneThread([post("p1", { conversationId: null, conversationThreadId: "other" })]))), /GROUP_ID_CHAIN_MISMATCH/);
  await assert.rejects(listGroupPosts(config, group, 3, options,
    clientFor(oneThread([post("p1", { conversationId: "other", conversationThreadId: "t1" })]))), /GROUP_ID_CHAIN_MISMATCH/);
});

test("limit uses a resumable position without discarding overfull pages or serializing bodies", async () => {
  const client = clientFor(oneThread([post("p1"), post("p2"), post("p3")]));
  const first = await listGroupPosts(config, group, 2, options, client);
  assert.equal(first.page.continuationAvailable, true);
  assert.deepEqual(first.posts.map(p => p.id), ["p1", "p2"]);
  assert.doesNotMatch(Buffer.from(first.page.nextCursor.split(".")[1], "base64url").toString(), /full body|sender@example/);
  const second = await listGroupPosts(config, group, 2, { ...options, cursor: first.page.nextCursor }, client);
  assert.deepEqual(second.posts.map(p => p.id), ["p3"]);
  assert.equal(second.page.complete, true);
});

test("post dates, not conversation delivery summary, determine inclusion; no assumed order", async () => {
  const client = clientFor(oneThread([post("old", { createdDateTime: "2026-09-01T00:00:00Z" }), post("new"), post("future", { createdDateTime: "2026-11-01T00:00:00Z" })]));
  const result = await listGroupPosts(config, group, 20, options, client);
  assert.deepEqual(result.posts.map(p => p.id), ["new"]);
  assert.equal(result.range.dateField, "createdDateTime");
  assert.equal(result.page.complete, true);
});

test("default date range is 90 days in Asia/Seoul", async () => {
  const result = await listGroupPosts(config, group, 20, { now: options.now }, clientFor({ [base]: { value: [] } }));
  assert.equal(result.range.days, 90);
  assert.equal(result.range.timeZone, "Asia/Seoul");
  assert.equal(result.range.usedDefaultLookback, true);
});

test("missing body hydrates post detail without treating preview as complete", async () => {
  const routes = oneThread([post("p1", { body: undefined, preview: "not full" })]);
  routes[`${base}/c1/threads/t1/posts/p1`] = post("p1", { body: { contentType: "text", content: "<literal text>" } });
  const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(result.posts[0].body, "<literal text>");
  assert.equal(result.posts[0].bodyHtml, undefined);
  assert.equal(result.posts[0].fullBodyAvailable, true);
});

test("body failure and unknown timestamp remain partial even after enumeration ends", async () => {
  const routes = oneThread([post("p1", { body: undefined }), post("p2", { createdDateTime: undefined })]);
  routes[`${base}/c1/threads/t1/posts/p1`] = new Error("Graph GET failed (403 Forbidden): PRIVATE_SECRET");
  routes[`${base}/c1/threads/t1/posts/p2`] = new Error("Graph GET failed (403 Forbidden): PRIVATE_SECRET");
  const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(result.page.complete, false);
  assert.equal(result.page.fullBodyUnavailableCount, 1);
  assert.equal(result.page.cumulativeTimestampUnavailableCount, 1);
  assert.equal(result.posts[0].fullBodyAvailable, false);
  assert.equal(result.page.continuationAvailable, false);
  assert.doesNotMatch(JSON.stringify(result), /PRIVATE_SECRET/);
});

test("deduplicates overlap and prevents repeating server continuations", async () => {
  const routes = oneThread([post("p1")]);
  const postPath = `${base}/c1/threads/t1/posts`;
  routes[postPath]["@odata.nextLink"] = `https://graph.microsoft.com${postPath}?page=2`;
  routes[`${postPath}?page=2`] = { value: [post("p1"), post("p2")] };
  let result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.deepEqual(result.posts.map(p => p.id), ["p1", "p2"]);
  assert.equal(result.page.duplicatesRemoved, 1);
  routes[`${postPath}?page=2`]["@odata.nextLink"] = `https://graph.microsoft.com${postPath}?page=2`;
  result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(result.page.noProgressDetected, true);
  assert.equal(result.page.continuationAvailable, false);
});

test("empty Graph pages with nextLink are followed", async () => {
  const routes = { [base]: { value: [], "@odata.nextLink": `https://graph.microsoft.com${base}?page=2` }, [`${base}?page=2`]: { value: [] } };
  const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(result.page.complete, true);
  assert.equal(result.page.requests, 2);
});

test("scan/request budgets return partial with continuation even before first post", async () => {
  const result = await listGroupPosts(config, group, 20, { ...options, maxRequests: 2 }, clientFor(oneThread([post("p1")])));
  assert.equal(result.page.complete, false);
  assert.equal(result.posts.length, 0);
  assert.equal(result.page.continuationAvailable, true);
  assert.ok(result.page.partialReasons.includes("SCAN_BUDGET"));
});

test("deadline returns resumable partial and bounds request timeout", async () => {
  let clock = options.now();
  const client = clientFor(oneThread([post("p1")]));
  const get = client.get;
  client.get = async (...args) => { const value = await get(...args); clock += 10; return value; };
  const result = await listGroupPosts(config, group, 20, { ...options, now: () => clock, timeBudgetMs: 5 }, client);
  assert.ok(result.page.partialReasons.includes("TIME_BUDGET"));
  assert.equal(result.page.continuationAvailable, true);
  assert.equal(client.calls[0].requestOptions.totalTimeoutMs, 5);
});

test("first access denial is distinct, later failure returns partial stage evidence", async () => {
  await assert.rejects(listGroupPosts(config, group, 20, options, clientFor({ [base]: new Error("Graph GET failed (403 Forbidden): secret") })), /GROUP_CONVERSATIONS_ACCESS_DENIED/);
  const routes = oneThread([post("p1")]);
  routes[`${base}/c1/threads/t1/posts`] = new Error("Graph GET failed (403 Forbidden): secret");
  const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(result.page.complete, false);
  assert.match(result.page.failure, /GROUP_POSTS_ACCESS_DENIED/);
  assert.doesNotMatch(result.page.failure, /secret/);
});

test("rejects cross-conversation response and cross-group/host continuation", async () => {
  await assert.rejects(listGroupPosts(config, group, 20, options, clientFor(oneThread([post("p1", { conversationThreadId: "other" })]))), /GROUP_ID_CHAIN_MISMATCH/);
  for (const next of ["https://evil.example/v1.0/groups/group-one/conversations", "https://graph.microsoft.com/v1.0/users", "https://graph.microsoft.com/v1.0/groups/other/conversations"]) {
    await assert.rejects(listGroupPosts(config, group, 20, options, clientFor({ [base]: { value: [], "@odata.nextLink": next } })), /GROUP_INVALID_CONTINUATION/);
  }
});

test("cursor is bound to group and date range and detects changed page", async () => {
  const routes = oneThread([post("p1"), post("p2")]);
  const first = await listGroupPosts(config, group, 1, options, clientFor(routes));
  await assert.rejects(listGroupPosts(config, group, 2, { ...options, since: "2026-10-02", cursor: first.page.nextCursor }, clientFor(routes)), /GROUP_INVALID_CURSOR/);
  await assert.rejects(listGroupPosts(config, { ...group, id: "other" }, 2, { ...options, cursor: first.page.nextCursor }, clientFor(routes)), /GROUP_INVALID_CURSOR/);
  routes[`${base}/c1/threads/t1/posts`].value.unshift(post("inserted"));
  const result = await listGroupPosts(config, group, 2, { ...options, cursor: first.page.nextCursor }, clientFor(routes));
  assert.match(result.page.failure, /GROUP_CURSOR_STALE/);
  assert.equal(result.page.continuationAvailable, false);
  assert.equal(result.page.complete, false);
});

test("unknown and non-Unified groups cannot call conversation APIs", async () => {
  let called = false;
  await assert.rejects(listGroupPosts(config, { ...group, groupTypes: [], classification: "distribution" }, 20, options, { get: async () => { called = true; } }), /GROUP_TYPE_UNSUPPORTED/);
  assert.equal(called, false);
});

test("strict options and safe stage classification", async () => {
  for (const limit of [0, 1.5, 101, NaN]) await assert.rejects(listGroupPosts(config, group, limit), /GROUP_INVALID_OPTION/);
  assert.match(groupReadError("POSTS", new Error("AUTH_CHECK_BLOCKED internal")).message, /GROUP_POSTS_AUTH_BLOCKED/);
  assert.match(groupReadError("POSTS", new Error("fetch failed ENOTFOUND secret")).message, /GROUP_POSTS_NETWORK_BLOCKED/);
  assert.throws(() => groupPageUrl("https://name:pass@graph.microsoft.com/v1.0/groups/g", "/groups/g"), /GROUP_INVALID_CONTINUATION/);
});

test("minimum request budget can hydrate after resume and rejects nonviable budget", async () => {
  const routes = oneThread([post("p1", { body: undefined })]);
  routes[`${base}/c1/threads/t1/posts/p1`] = post("p1");
  const client = clientFor(routes);
  await assert.rejects(listGroupPosts(config, group, 20, { ...options, maxRequests: 1 }, client), /GROUP_INVALID_OPTION/);
  let result = await listGroupPosts(config, group, 20, { ...options, maxRequests: 2 }, client);
  const found = [];
  for (let call = 0; call < 4 && result.page.continuationAvailable; call++) {
    result = await listGroupPosts(config, group, 20, { ...options, maxRequests: 2, cursor: result.page.nextCursor }, client);
    found.push(...result.posts);
  }
  assert.equal(result.page.continuationAvailable, false);
  assert.equal(result.page.complete, true);
  assert.deepEqual(found.map(p => p.id), ["p1"]);
});

test("duplicate-only intermediary page does not hide a later unique post", async () => {
  const routes = oneThread([post("p1")]);
  const postPath = `${base}/c1/threads/t1/posts`;
  routes[postPath]["@odata.nextLink"] = `https://graph.microsoft.com${postPath}?page=2`;
  routes[`${postPath}?page=2`] = { value: [post("p1")], "@odata.nextLink": `https://graph.microsoft.com${postPath}?page=3` };
  routes[`${postPath}?page=3`] = { value: [post("p2")] };
  const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(result.page.complete, true);
  assert.deepEqual(result.posts.map(p => p.id), ["p1", "p2"]);
});

test("modified query, offsets, job stack or counters are rejected before network access", async () => {
  const first = await listGroupPosts(config, group, 1, options, clientFor(oneThread([post("p1"), post("p2")])));
  for (const mutate of [
    state => { state.jobs[0].url += "&$filter=id eq 'other'"; },
    state => { state.jobs[0].offset = 1000; },
    state => { state.jobs = state.jobs.slice(0, 1); },
    state => { state.missingBodies = 0; state.missingTimestamps = 99; },
    state => { state.seen = []; }
  ]) {
    const parts = first.page.nextCursor.split(".");
    const envelope = JSON.parse(Buffer.from(parts[1], "base64url").toString());
    mutate(envelope.state);
    parts[1] = Buffer.from(JSON.stringify(envelope)).toString("base64url");
    const client = clientFor({});
    await assert.rejects(listGroupPosts(config, group, 20, { ...options, cursor: parts.join(".") }, client), /GROUP_INVALID_CURSOR/);
    assert.equal(client.calls.length, 0);
  }
});

test("failed resumed page stops with sanitized failure instead of a non-advancing cursor", async () => {
  const first = await listGroupPosts(config, group, 20, { ...options, maxRequests: 2 }, clientFor(oneThread([post("p1")])));
  const routes = oneThread([post("p1")]);
  routes[`${base}/c1/threads/t1/posts`] = new Error("Graph GET failed (403 Forbidden): PRIVATE_SECRET");
  const result = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, clientFor(routes));
  assert.equal(result.page.noProgressDetected, true);
  assert.equal(result.page.continuationAvailable, false);
  assert.match(result.page.failure, /GROUP_POSTS_ACCESS_DENIED/);
  assert.doesNotMatch(result.page.failure, /PRIVATE_SECRET/);
});

test("a stricter existing search policy caps group results", async () => {
  const result = await listGroupPosts({ ...config, policy: { ...config.policy, maxSearchResults: 1 } }, group, 20, options, clientFor(oneThread([post("p1"), post("p2")])));
  assert.equal(result.posts.length, 1);
  assert.equal(result.page.limit, 1);
  assert.equal(result.page.continuationAvailable, true);
});
