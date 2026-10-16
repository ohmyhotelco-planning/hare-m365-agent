import assert from "node:assert/strict";
import test, { after } from "node:test";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createHash } from "node:crypto";
import { listGroupPosts } from "../dist/groups-posts.js";
import { groupPageUrl, groupReadError } from "../dist/groups-common.js";
import { encodeGroupCursor } from "../dist/groups-cursor.js";
import { listGroupPostAttachments } from "../dist/groups-attachments.js";

const fixtureRoot = fs.mkdtempSync(path.join(os.tmpdir(), "hare-group-posts-"));
after(() => {
  assert.equal(path.dirname(path.resolve(fixtureRoot)), path.resolve(os.tmpdir()));
  assert.match(path.basename(fixtureRoot), /^hare-group-posts-/);
  fs.rmSync(fixtureRoot, { recursive: true, force: true });
});
const config = { cacheDir: path.join(fixtureRoot, "cache"), policy: { defaultSearchLookbackDays: 90, maxSearchResults: 1000 } };
const group = { id: "group-one", mail: "team@example.test", groupTypes: ["Unified"], classification: "microsoft365", evidence: ["groupTypes includes Unified"], mailEnabled: true, securityEnabled: false, exchangeRecipientType: "unverified" };
const base = "/v1.0/groups/group-one/conversations";
const resumedPostPath = "/v1.0/groups/group-one/threads/t1/posts";
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
    [`${base}/c1/threads/t1/posts`]: { value: posts },
    ...Object.fromEntries(posts.map(value => [`${resumedPostPath}/${encodeURIComponent(value.id)}`, value])) };
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
  for (const extra of [{ conversationId: null, conversationThreadId: "other" }, { conversationId: "other", conversationThreadId: "t1" }]) {
    const rejected = await listGroupPosts(config, group, 3, options, clientFor(oneThread([post("p1", extra)])));
    assert.equal(rejected.posts.length, 0);
    assert.equal(rejected.page.complete, false);
    assert.equal(rejected.page.continuationAvailable, false);
    assert.ok(rejected.page.failure);
  }
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
  const rejected = await listGroupPosts(config, group, 20, options, clientFor(oneThread([post("p1", { conversationThreadId: "other" })])));
  assert.equal(rejected.posts.length, 0);
  assert.equal(rejected.page.continuationAvailable, false);
  assert.ok(rejected.page.failure);
  for (const next of ["https://evil.example/v1.0/groups/group-one/conversations", "https://graph.microsoft.com/v1.0/users", "https://graph.microsoft.com/v1.0/groups/other/conversations"]) {
    await assert.rejects(listGroupPosts(config, group, 20, options, clientFor({ [base]: { value: [], "@odata.nextLink": next } })), /GROUP_INVALID_CONTINUATION/);
  }
});

test("cursor is bound to group and date range and resumes pinned posts despite changed page", async () => {
  const routes = oneThread([post("p1"), post("p2")]);
  const first = await listGroupPosts(config, group, 1, options, clientFor(routes));
  await assert.rejects(listGroupPosts(config, group, 2, { ...options, since: "2026-10-02", cursor: first.page.nextCursor }, clientFor(routes)), /GROUP_INVALID_CURSOR/);
  await assert.rejects(listGroupPosts(config, { ...group, id: "other" }, 2, { ...options, cursor: first.page.nextCursor }, clientFor(routes)), /GROUP_INVALID_CURSOR/);
  routes[`${base}/c1/threads/t1/posts`].value.unshift(post("inserted"));
  const client = clientFor(routes);
  const result = await listGroupPosts(config, group, 2, { ...options, cursor: first.page.nextCursor }, client);
  assert.deepEqual(result.posts.map(value => value.id), ["p2"]);
  assert.deepEqual(client.calls.map(call => call.key), [`${resumedPostPath}/p2`]);
  assert.equal(result.page.continuationAvailable, false);
  assert.equal(result.page.complete, true);
  assert.equal(result.page.snapshotConsistent, false);
  assert.equal(result.page.completionScope, "PINNED_PAGE_INVENTORIES_WITH_LIVE_CONTINUATIONS");
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
  routes[`${resumedPostPath}/p1`] = post("p1");
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

const verifiedPath = "/v1.0/groups/group-one/threads/returned-thread/posts";
function mismatchedRoutes(extra = {}) {
  return {
    ...oneThread([post("p1", { conversationId: null, conversationThreadId: "returned-thread", ...extra })]),
    [verifiedPath]: { value: [post("p1", { conversationId: null, conversationThreadId: "returned-thread" })] }
  };
}

test("revalidates exact post membership and passes the verified thread to unchanged attachment guards", async () => {
  const routes = mismatchedRoutes();
  routes[verifiedPath].value[0].body.content = "verified body";
  const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  const p = result.posts[0];
  assert.equal(result.page.complete, true);
  assert.equal(result.page.requests, 4);
  assert.equal(p.threadId, "returned-thread");
  assert.equal(p.traversalThreadId, "t1");
  assert.equal(p.threadVerification, "POST_COLLECTION_REVALIDATED");
  assert.equal(p.bodyContent, "verified body");
  const client = clientFor({
    [`${verifiedPath}/p1`]: { id: "p1", conversationThreadId: "returned-thread" },
    [`${verifiedPath}/p1/attachments`]: { value: [{ id: "a1", name: "test.pdf", size: 20, "@odata.type": "#microsoft.graph.fileAttachment" }] }
  });
  const attachments = await listGroupPostAttachments(config, { groupId: p.groupId, threadId: p.threadId, postId: p.id }, 20, {}, client);
  assert.equal(attachments.attachments[0].id, "a1");
  assert.equal(client.calls.length, 2);
});

test("30-result first page, 82-result scan budget, then revalidated third page retain unique posts", async () => {
  const routes = oneThread([...Array.from({ length: 112 }, (_, i) => post(`p${i}`)), post("p112", { conversationThreadId: "returned-thread" })]);
  routes[verifiedPath] = { value: [post("p112", { conversationThreadId: "returned-thread" })] };
  const client = clientFor(routes);
  const first = await listGroupPosts(config, group, 30, options, client);
  assert.equal(first.posts.length, 30);
  assert.deepEqual(first.page.partialReasons, ["RESULT_LIMIT"]);
  const second = await listGroupPosts(config, group, 100, { ...options, maxRecords: 82, cursor: first.page.nextCursor }, client);
  assert.equal(second.posts.length, 82);
  assert.deepEqual(second.page.partialReasons, ["SCAN_BUDGET"]);
  const third = await listGroupPosts(config, group, 100, { ...options, cursor: second.page.nextCursor }, client);
  assert.equal(third.page.complete, true);
  assert.equal(third.posts[0].threadVerification, "POST_COLLECTION_REVALIDATED");
  assert.equal(new Set([...first.posts, ...second.posts, ...third.posts].map(p => p.id)).size, 113);
});

test("deduplicates the same post across different traversal threads and cursor calls", async () => {
  const routes = mismatchedRoutes();
  routes[`${base}/c1/threads`].value.push({ id: "t2" });
  routes[`${base}/c1/threads/t2/posts`] = { value: [post("p1"), post("p2")] };
  const first = await listGroupPosts(config, group, 1, options, clientFor(routes));
  const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, clientFor(routes));
  assert.deepEqual(second.posts.map(p => p.id), ["p2"]);
  assert.equal(second.page.duplicatesRemoved, 1);
  assert.equal(second.page.complete, true);
});

test("unverified thread, post, conversation and malformed evidence fail closed while retaining prior posts", async () => {
  for (const value of [[], [post("different", { conversationThreadId: "returned-thread" })],
    [post("p1", { conversationThreadId: "wrong" })], [post("p1", { conversationThreadId: null })],
    [post("p1", { conversationThreadId: "returned-thread", conversationId: "other-conversation" })],
    [post("p1", { conversationThreadId: "returned-thread" }), post("p1", { conversationThreadId: "returned-thread" })]]) {
    const routes = mismatchedRoutes();
    routes[`${base}/c1/threads/t1/posts`].value.unshift(post("safe"));
    routes[verifiedPath] = { value };
    const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
    assert.deepEqual(result.posts.map(p => p.id), ["safe"]);
    assert.match(result.page.failure, /GROUP_ID_CHAIN_MISMATCH/);
    assert.equal(result.page.complete, false);
    assert.equal(result.page.continuationAvailable, false);
  }
});

test("verification page continuations cannot change group, thread or host and have a three-page bound", async () => {
  for (const next of ["https://evil.example/posts", "https://graph.microsoft.com/v1.0/groups/other/threads/returned-thread/posts", "https://graph.microsoft.com/v1.0/groups/group-one/threads/other/posts"]) {
    const routes = mismatchedRoutes();
    routes[verifiedPath] = { value: [], "@odata.nextLink": next };
    const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
    assert.match(result.page.failure, /GROUP_INVALID_CONTINUATION/);
    assert.equal(result.posts.length, 0);
  }
  const routes = mismatchedRoutes();
  for (let i = 1; i <= 4; i++) routes[verifiedPath + (i === 1 ? "" : `?page=${i}`)] = { value: [], "@odata.nextLink": `https://graph.microsoft.com${verifiedPath}?page=${i + 1}` };
  const client = clientFor(routes);
  const result = await listGroupPosts(config, group, 20, options, client);
  assert.match(result.page.failure, /GROUP_ID_CHAIN_MISMATCH/);
  assert.equal(client.calls.length, 6);
  assert.equal(result.page.continuationAvailable, false);
});

test("verification uses a bounded next page and accounts for requests and scanned records", async () => {
  const routes = mismatchedRoutes();
  routes[`${verifiedPath}?page=2`] = routes[verifiedPath];
  routes[verifiedPath] = { value: [post("other", { conversationThreadId: "returned-thread" })], "@odata.nextLink": `https://graph.microsoft.com${verifiedPath}?page=2` };
  const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(result.page.complete, true);
  assert.equal(result.page.requests, 5);
  assert.equal(result.page.scannedRecords, 5);
});

test("verification budget interruption resumes without skipping the unresolved post", async () => {
  const routes = mismatchedRoutes();
  const first = await listGroupPosts(config, group, 20, { ...options, maxRequests: 3 }, clientFor(routes));
  assert.deepEqual(first.posts, []);
  assert.deepEqual(first.page.partialReasons, ["SCAN_BUDGET"]);
  assert.equal(first.page.continuationAvailable, true);
  const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, clientFor(routes));
  assert.equal(second.posts[0].id, "p1");
  assert.equal(second.page.complete, true);
});

test("null thread is not accepted and mismatched conversation does not trigger thread verification", async () => {
  for (const extra of [{ conversationThreadId: null }, { conversationId: "other", conversationThreadId: "returned-thread" }]) {
    const client = clientFor(mismatchedRoutes(extra));
    const result = await listGroupPosts(config, group, 20, options, client);
    assert.match(result.page.failure, /GROUP_ID_CHAIN_MISMATCH/);
    assert.equal(client.calls.length, 3);
    assert.equal(result.posts.length, 0);
  }
});

test("hydration after revalidation stays in verified thread and rejects changed or missing chain evidence", async () => {
  for (const detail of [post("wrong", { conversationThreadId: "returned-thread" }), post("p1", { conversationThreadId: "other" }), post("p1")]) {
    const routes = mismatchedRoutes();
    routes[verifiedPath].value[0].body = undefined;
    routes[`${verifiedPath}/p1`] = detail;
    const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
    assert.match(result.page.failure, /GROUP_ID_CHAIN_MISMATCH/);
    assert.equal(result.posts.length, 0);
  }
  const routes = mismatchedRoutes();
  routes[verifiedPath].value[0].body = undefined;
  routes[`${verifiedPath}/p1`] = post("p1", { conversationThreadId: "returned-thread" });
  const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(result.posts[0].fullBodyAvailable, true);
});

test("old cursor versions require restart before network access", async () => {
  const first = await listGroupPosts(config, group, 1, options, clientFor(oneThread([post("p1"), post("p2")])));
  const envelope = JSON.parse(Buffer.from(first.page.nextCursor.split(".")[1], "base64url").toString());
  for (const version of [1, 2]) {
    envelope.state.version = version;
    const cursor = encodeGroupCursor(config, "posts", envelope.state);
    const client = clientFor({});
    await assert.rejects(listGroupPosts(config, group, 20, { ...options, cursor }, client), /GROUP_INVALID_CURSOR.*Restart/);
    assert.equal(client.calls.length, 0);
  }
});

test("a mismatch discovered only during hydration also requires collection membership", async () => {
  const routes = mismatchedRoutes();
  routes[`${base}/c1/threads/t1/posts`] = { value: [post("p1", { body: undefined })] };
  routes[`${base}/c1/threads/t1/posts/p1`] = post("p1", { conversationThreadId: "returned-thread" });
  routes[verifiedPath].value[0].body.content = "verified detail";
  const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(result.posts[0].threadId, "returned-thread");
  assert.equal(result.posts[0].bodyContent, "verified detail");
  assert.equal(result.page.complete, true);
  routes[verifiedPath] = new Error("Graph GET failed (403 Forbidden): PRIVATE_SECRET");
  const rejected = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(rejected.posts.length, 0);
  assert.equal(rejected.page.continuationAvailable, false);
  assert.match(rejected.page.failure, /GROUP_POSTS_ACCESS_DENIED/);
  assert.doesNotMatch(JSON.stringify(rejected), /PRIVATE_SECRET/);
});

test("resumed validation failure is not obscured by NO_PROGRESS and malformed IDs never leave the process", async () => {
  for (const thread of [null, " returned-thread ", ".", "..", ""]) {
    const routes = mismatchedRoutes({ conversationThreadId: thread });
    const first = await listGroupPosts(config, group, 20, { ...options, maxRequests: 2 }, clientFor(routes));
    const client = clientFor(routes);
    const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, client);
    assert.deepEqual(second.page.partialReasons, ["READ_FAILED"]);
    assert.match(second.page.failure, /GROUP_ID_CHAIN_MISMATCH/);
    assert.equal(second.page.continuationAvailable, false);
    assert.equal(client.calls.length, 1);
  }
});

test("verification failure preserves earlier hydrated posts and hides remote error details", async () => {
  for (const remote of [{ value: null }, new Error("Graph GET failed (404 Not Found): PRIVATE_SECRET")]) {
    const routes = mismatchedRoutes();
    routes[`${base}/c1/threads/t1/posts`].value.unshift(post("safe"));
    routes[verifiedPath] = remote;
    const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
    assert.deepEqual(result.posts.map(p => p.id), ["safe"]);
    assert.equal(result.page.continuationAvailable, false);
    assert.doesNotMatch(JSON.stringify(result), /PRIVATE_SECRET/);
  }
});

test("third verification page succeeds, cyclic pages fail and deadlines stop before another request", async () => {
  const routes = mismatchedRoutes();
  routes[`${verifiedPath}?page=3`] = routes[verifiedPath];
  routes[verifiedPath] = { value: [], "@odata.nextLink": `https://graph.microsoft.com${verifiedPath}?page=2` };
  routes[`${verifiedPath}?page=2`] = { value: [], "@odata.nextLink": `https://graph.microsoft.com${verifiedPath}?page=3` };
  const result = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(result.page.complete, true);
  assert.equal(result.page.requests, 6);
  routes[`${verifiedPath}?page=2`]["@odata.nextLink"] = `https://graph.microsoft.com${verifiedPath}?page=2`;
  const rejected = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.match(rejected.page.failure, /GROUP_ID_CHAIN_MISMATCH/);
  const client = clientFor(routes), get = client.get;
  let now = options.now();
  client.get = async (...args) => { const result = await get(...args); now += 10; return result; };
  const timed = await listGroupPosts(config, group, 20, { ...options, now: () => now, timeBudgetMs: 35 }, client);
  assert.deepEqual(timed.page.partialReasons, ["TIME_BUDGET"]);
  assert.equal(timed.page.requests, 4);
  assert.equal(timed.posts.length, 0);
});

test("verified timestamp, not the original mismatched row, controls the requested date range", async () => {
  const routes = mismatchedRoutes();
  routes[verifiedPath].value[0].createdDateTime = "2026-09-01T00:00:00Z";
  const outside = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(outside.posts.length, 0);
  assert.equal(outside.page.complete, true);
  routes[`${base}/c1/threads/t1/posts`].value[0].createdDateTime = "2026-09-01T00:00:00Z";
  routes[verifiedPath].value[0].createdDateTime = "2026-10-10T01:00:00Z";
  const inside = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(inside.posts.length, 1);
  assert.equal(inside.posts[0].createdDateTime, "2026-10-10T01:00:00Z");
});

test("verification that cannot fit a record budget stops without advancing or looping forever", async () => {
  const routes = mismatchedRoutes();
  routes[verifiedPath] = { value: Array.from({ length: 1000 }, (_, i) => post(`v${i}`, { conversationThreadId: "returned-thread" })) };
  const first = await listGroupPosts(config, group, 20, options, clientFor(routes));
  assert.equal(first.posts.length, 0);
  assert.deepEqual(first.page.partialReasons, ["SCAN_BUDGET"]);
  assert.equal(first.page.continuationAvailable, true);
  const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, clientFor(routes));
  assert.equal(second.posts.length, 0);
  assert.deepEqual(second.page.partialReasons, ["NO_PROGRESS"]);
  assert.equal(second.page.continuationAvailable, false);
  assert.equal(second.page.complete, false);
});

function cursorState(cursor) {
  return JSON.parse(Buffer.from(cursor.split(".")[1], "base64url").toString()).state;
}

test("oversized pinned inventory preserves collected posts and stops with explicit partial status", async () => {
  const routes = oneThread([post("p1")]);
  routes[base].value.push(...Array.from({ length: 999 }, (_, index) => ({ id: `${index}-` + "x".repeat(4080) })));
  const result = await listGroupPosts(config, group, 1, options, clientFor(routes));
  assert.deepEqual(result.posts.map(value => value.id), ["p1"]);
  assert.equal(result.posts[0].fullBodyAvailable, true);
  assert.equal(result.page.complete, false);
  assert.equal(result.page.partialResult, true);
  assert.deepEqual(result.page.partialReasons, ["CURSOR_SIZE_LIMIT"]);
  assert.equal(result.page.continuationAvailable, false);
  assert.equal(result.page.nextCursor, undefined);
});
function refreshFingerprint(snapshot) {
  snapshot.fingerprint = createHash("sha256").update(JSON.stringify([snapshot.ids, snapshot.nextLink ?? null])).digest("hex");
}

test("pinned parent inventories survive insert, reorder and removal without rereading parent pages", async () => {
  for (const changed of [["inserted", "c1", "c2"], ["c2", "c1"], ["c1"]]) {
    const routes = oneThread([post("p1")]);
    routes[base].value.push({ id: "c2" });
    routes[`${base}/c1/threads`].value.push({ id: "t2" });
    routes[`${base}/c1/threads/t2/posts`] = { value: [post("p2")] };
    routes[`${base}/c2/threads`] = { value: [{ id: "t3" }] };
    routes[`${base}/c2/threads/t3/posts`] = { value: [post("p3")] };
    const first = await listGroupPosts(config, group, 1, options, clientFor(routes));
    routes[base].value = changed.map(id => ({ id }));
    routes[`${base}/c1/threads`].value = [{ id: "inserted-thread" }, { id: "t1" }];
    const client = clientFor(routes);
    const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, client);
    assert.deepEqual(second.posts.map(value => value.id), ["p2", "p3"]);
    assert.equal(second.page.complete, true);
    assert.equal(second.page.snapshotConsistent, false);
    assert.ok(!client.calls.some(call => call.key === base || call.key === `${base}/c1/threads`));
  }
});

test("pinned post order and nextLink survive removal, reorder and replacement of the live page", async () => {
  const routes = oneThread([post("p1"), post("p2"), post("p3")]);
  const collection = `${base}/c1/threads/t1/posts`;
  routes[collection]["@odata.nextLink"] = `https://graph.microsoft.com${collection}?page=2`;
  routes[`${collection}?page=2`] = { value: [post("p3"), post("p4")] };
  const first = await listGroupPosts(config, group, 1, options, clientFor(routes));
  routes[collection] = { value: [post("p3"), post("inserted")], "@odata.nextLink": "https://evil.example/changed" };
  routes[base] = new Error("Parent collection must not be reread");
  routes[`${base}/c1/threads`] = new Error("Parent collection must not be reread");
  const client = clientFor(routes);
  const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, client);
  assert.deepEqual(second.posts.map(value => value.id), ["p2", "p3", "p4"]);
  assert.equal(second.page.duplicatesRemoved, 1);
  assert.equal(second.page.complete, true);
  assert.deepEqual(client.calls.map(call => call.key), [`${resumedPostPath}/p2`, `${resumedPostPath}/p3`, `${collection}?page=2`]);
});

test("missing pinned parents stop partially instead of skipping their descendants", async () => {
  for (const missing of [`${base}/c1/threads/t2/posts`, `${base}/c2/threads`]) {
    const routes = oneThread([post("p1")]);
    if (missing.includes("/t2/")) routes[`${base}/c1/threads`].value.push({ id: "t2" });
    else routes[base].value.push({ id: "c2" });
    const first = await listGroupPosts(config, group, 1, options, clientFor(routes));
    routes[missing] = new Error("Graph GET failed (404 Not Found): PRIVATE_SECRET");
    const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, clientFor(routes));
    assert.equal(second.page.complete, false);
    assert.match(second.page.failure, /NOT_FOUND_OR_UNAVAILABLE/);
    assert.deepEqual(second.posts, []);
    assert.doesNotMatch(JSON.stringify(second), /PRIVATE_SECRET/);
  }
});

test("exact resumed lookup rejects missing, wrong-ID and wrong-chain responses without skipping", async () => {
  for (const detail of [null, {}, post("different"), post("P2"), post("p2", { conversationId: "other" }),
    post("p2", { conversationThreadId: null }), new Error("Graph GET failed (404 Not Found): PRIVATE_SECRET")]) {
    const routes = oneThread([post("p1"), post("p2"), post("p3")]);
    const first = await listGroupPosts(config, group, 1, options, clientFor(routes));
    routes[`${resumedPostPath}/p2`] = detail;
    const client = clientFor(routes);
    const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, client);
    assert.deepEqual(second.posts, []);
    assert.equal(second.page.complete, false);
    assert.equal(second.page.continuationAvailable, false);
    assert.deepEqual(second.page.partialReasons, ["READ_FAILED"]);
    assert.match(second.page.failure, /GROUP_ID_CHAIN_MISMATCH|NOT_FOUND_OR_UNAVAILABLE/);
    assert.doesNotMatch(JSON.stringify(second), /PRIVATE_SECRET/);
    assert.deepEqual(client.calls.map(call => call.key), [`${resumedPostPath}/p2`]);
  }
});

test("signed v3 inventory validation rejects malformed IDs, offsets, chains, nextLinks and bodies before reads", async () => {
  const first = await listGroupPosts(config, group, 1, options, clientFor(oneThread([post("p1"), post("p2")])));
  const mutations = [
    state => { state.jobs[2].offset = 3; },
    state => { delete state.jobs[2].snapshot; },
    state => { state.jobs[0].offset = 0; },
    state => { state.jobs[2].snapshot.ids.reverse(); },
    state => { state.jobs[2].snapshot.ids = Array(1001).fill("p2"); refreshFingerprint(state.jobs[2].snapshot); },
    state => { state.jobs[1].conversationId = "other"; state.jobs[1].url = "/groups/group-one/conversations/other/threads?$select=id"; },
    state => { state.jobs[2].threadId = "other"; state.jobs[2].url = "/groups/group-one/conversations/c1/threads/other/posts?$select=id"; },
    state => { state.jobs[2].snapshot.body = "must not be retained"; },
    state => { state.body = "must not be retained"; },
    ...[" p2", "p2 ", ".", "..", "", "bad\u0000id", "\ud800", null, 123, {}].map(id => state => {
      state.jobs[2].snapshot.ids[1] = id; refreshFingerprint(state.jobs[2].snapshot);
    }),
    ...["https://evil.example/posts", "https://graph.microsoft.com/v1.0/groups/other/conversations/c1/threads/t1/posts",
      "https://graph.microsoft.com/v1.0/groups/group-one/conversations/c1/threads/other/posts", "", null].map(next => state => {
      state.jobs[2].snapshot.nextLink = next; refreshFingerprint(state.jobs[2].snapshot);
    })
  ];
  for (const mutate of mutations) {
    const state = cursorState(first.page.nextCursor);
    mutate(state);
    const client = clientFor({});
    const cursor = encodeGroupCursor(config, "posts", state);
    await assert.rejects(listGroupPosts(config, group, 20, { ...options, cursor }, client), /GROUP_INVALID_CURSOR/);
    assert.equal(client.calls.length, 0);
  }
});

test("opaque IDs retain their exact bytes and cursor inventories contain IDs only", async () => {
  const id = "Case+Sensitive/opaque=%2F?end";
  const routes = oneThread([post("p1"), post(id, { from: { emailAddress: { address: "PRIVATE_SENDER" } } })]);
  const first = await listGroupPosts(config, group, 1, options, clientFor(routes));
  const state = cursorState(first.page.nextCursor);
  assert.equal(state.version, 3);
  assert.deepEqual(state.jobs[2].snapshot.ids, ["p1", id]);
  assert.doesNotMatch(JSON.stringify(state), /PRIVATE_SENDER|full body|"createdDateTime":|"body":|"bodyContent":/);
  const client = clientFor(routes);
  const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, client);
  assert.deepEqual(second.posts.map(value => value.id), [id]);
  assert.equal(client.calls[0].key, `${resumedPostPath}/${encodeURIComponent(id)}`);
});

test("inventory acquisition is progress even when time expires before the first pinned item", async () => {
  const routes = oneThread([post("p1")]);
  const first = await listGroupPosts(config, group, 20, { ...options, maxRequests: 2 }, clientFor(routes));
  let clock = options.now();
  const client = clientFor(routes), get = client.get;
  client.get = async (...args) => { const value = await get(...args); clock += 10; return value; };
  const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor, now: () => clock, timeBudgetMs: 5 }, client);
  assert.deepEqual(second.posts, []);
  assert.deepEqual(second.page.partialReasons, ["TIME_BUDGET"]);
  assert.equal(second.page.noProgressDetected, false);
  assert.equal(second.page.continuationAvailable, true);
  assert.deepEqual(cursorState(second.page.nextCursor).jobs[2].snapshot.ids, ["p1"]);
  const third = await listGroupPosts(config, group, 20, { ...options, cursor: second.page.nextCursor }, clientFor(routes));
  assert.deepEqual(third.posts.map(value => value.id), ["p1"]);
  assert.equal(third.page.complete, true);
});
