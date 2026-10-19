import assert from "node:assert/strict";
import test, { after } from "node:test";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { listGroupPosts } from "../dist/groups-posts.js";
import { readGroupPostSubject } from "../dist/groups-post-subject.js";
import { readGroupCursorInput } from "../dist/groups-cursor-input.js";
import { listGroupPostAttachments } from "../dist/groups-attachments.js";

const root = fs.mkdtempSync(path.join(os.tmpdir(), "hare-subject-test-"));
after(() => {
  assert.equal(path.dirname(path.resolve(root)), path.resolve(os.tmpdir()));
  assert.match(path.basename(root), /^hare-subject-test-/);
  fs.rmSync(root, { recursive: true, force: true });
});
const config = { cacheDir: path.join(root, "cache"), policy: { defaultSearchLookbackDays: 90, maxSearchResults: 1000 } };
const group = { id: "g", groupTypes: ["Unified"], classification: "microsoft365" };
const options = { since: "2026-10-01", until: "2026-10-16", includeSubject: true, now: () => Date.parse("2026-10-17T01:00:00Z") };
const chain = { groupId: "g", conversationId: "c", threadId: "t", postId: "p", revalidated: false };
const subject = value => [{ id: "String 0x0037", value }];
const post = (id, extras = {}) => ({ id, conversationId: "c", conversationThreadId: "t", createdDateTime: "2026-10-10T01:00:00Z", body: { contentType: "html", content: `<p>${"full body ".repeat(100)}</p>` }, hasAttachments: true, ...extras });
function fixture(rows, subjects = {}) {
  const routes = {
    "/groups/g/conversations": { value: [{ id: "c", topic: "wrong parent title" }] },
    "/groups/g/conversations/c/threads": { value: [{ id: "t", topic: "wrong thread title" }] },
    "/groups/g/conversations/c/threads/t/posts": { value: rows },
    ...Object.fromEntries(rows.map(p => [`/groups/g/threads/t/posts/${p.id}`, p]))
  };
  const calls = [];
  return { routes, calls, subjects, async get(url, requestOptions) {
    assert.ok(requestOptions?.totalTimeoutMs > 0);
    const parsed = new URL(url.startsWith("/") ? `https://graph.microsoft.com/v1.0${url}` : url);
    const key = parsed.pathname.replace(/^\/v1\.0/, "");
    const expanded = parsed.searchParams.has("$expand");
    calls.push({ key, expanded });
    let result;
    if (expanded) {
      assert.equal(parsed.searchParams.get("$expand"), "singleValueExtendedProperties($filter=id eq 'String 0x0037')");
      assert.equal(parsed.searchParams.get("$select"), "id,conversationId,conversationThreadId");
      assert.ok(Object.hasOwn(subjects, key), `unexpected subject path ${key}`);
      result = subjects[key];
    } else {
      assert.ok(Object.hasOwn(routes, key), `unexpected path ${key}`);
      result = routes[key];
    }
    if (result instanceof Error) throw result;
    return typeof result === "function" ? result() : result;
  } };
}

test("subject values retain exact prefixes and whitespace, without topic or body fallback", async () => {
  const client = fixture([post("p")], { "/groups/g/threads/t/posts/p": { id: "p", conversationThreadId: "t", singleValueExtendedProperties: subject(" RE: rate input  ") } });
  const result = await listGroupPosts(config, group, 20, options, client);
  assert.deepEqual(result.posts[0].postSubject, { source: "PidTagSubject", propertyId: "String 0x0037", scope: "POST", status: "PRESENT", value: " RE: rate input  " });
  assert.equal(result.posts[0].bodyContent, post("p").body.content);
  assert.equal(result.posts[0].hasAttachments, true);
  assert.equal(result.page.complete, true);
  assert.equal(result.page.requests, 4);
  assert.equal(result.subjectRead.availableCount, 1);
});

for (const [properties, status] of [[subject(""), "EMPTY"], [[], "MISSING"], [undefined, "UNVERIFIED"], [null, "INVALID_RESPONSE"], [[{ id: "String 0x0037", value: null }], "INVALID_RESPONSE"], [[...subject("a"), ...subject("b")], "INVALID_RESPONSE"], [[{ id: "String 0x0038", value: "not subject" }], "MISSING"]]) {
  test(`subject property state ${status}: ${JSON.stringify(properties)}`, async () => {
    const result = await readGroupPostSubject(chain, async () => ({ id: "p", singleValueExtendedProperties: properties }));
    assert.equal(result.status, status);
    assert.equal(result.value, status === "EMPTY" ? "" : undefined);
  });
}

test("equivalent hexadecimal property ID is accepted, incomplete expansion is unverified", async () => {
  assert.equal((await readGroupPostSubject(chain, async () => ({ id: "p", singleValueExtendedProperties: [{ id: "String 0x37", value: "exact" }] }))).value, "exact");
  assert.equal((await readGroupPostSubject(chain, async () => ({ id: "p", singleValueExtendedProperties: [], "singleValueExtendedProperties@odata.nextLink": "unfollowed" }))).status, "UNVERIFIED");
});

test("opt-out preserves requests and output shape", async () => {
  const client = fixture([post("p")]);
  const result = await listGroupPosts(config, group, 20, { ...options, includeSubject: false }, client);
  assert.equal(result.page.requests, 3);
  assert.equal("postSubject" in result.posts[0], false);
  assert.equal("subjectRead" in result, false);
  assert.equal(client.calls.some(c => c.expanded), false);
});

test("different conversations with identical subject remain separate, overlap is deduplicated", async () => {
  const client = fixture([post("p"), post("p")], {
    "/groups/g/threads/t/posts/p": { id: "p", singleValueExtendedProperties: subject("same title") },
    "/groups/g/threads/t2/posts/p2": { id: "p2", conversationId: "c2", conversationThreadId: "t2", singleValueExtendedProperties: subject("same title") }
  });
  client.routes["/groups/g/conversations"].value.push({ id: "c2", topic: "different parent" });
  client.routes["/groups/g/conversations/c2/threads"] = { value: [{ id: "t2" }] };
  client.routes["/groups/g/conversations/c2/threads/t2/posts"] = { value: [post("p2", { conversationId: "c2", conversationThreadId: "t2" })] };
  const result = await listGroupPosts(config, group, 20, options, client);
  assert.deepEqual(result.posts.map(p => [p.id, p.conversationId, p.threadId, p.postSubject.value]), [["p", "c", "t", "same title"], ["p2", "c2", "t2", "same title"]]);
  assert.equal(result.page.duplicatesRemoved, 1);
  assert.equal(client.calls.filter(c => c.expanded).length, 2);
});

test("cursor-file resume preserves pinned IDs and does not serialize titles or reread parents", async () => {
  const client = fixture([post("p"), post("p2")], Object.fromEntries(["p", "p2"].map(id => [`/groups/g/threads/t/posts/${id}`, { id, singleValueExtendedProperties: subject(`private-title-${id}`) }])));
  const first = await listGroupPosts(config, group, 1, options, client);
  const stateText = Buffer.from(first.page.nextCursor.split(".")[1], "base64url").toString();
  assert.doesNotMatch(stateText, /private-title|PidTagSubject|includeSubject/);
  const file = path.join(root, "page.json");
  fs.writeFileSync(file, JSON.stringify(first));
  client.calls.length = 0;
  delete client.routes["/groups/g/conversations"];
  const second = await listGroupPosts(config, group, 20, { ...options, cursor: readGroupCursorInput({ cursorFile: file }) }, client);
  assert.deepEqual(second.posts.map(p => p.id), ["p2"]);
  assert.equal(second.posts[0].postSubject.value, "private-title-p2");
  assert.equal(second.page.complete, true);
  assert.ok(client.calls.every(c => c.key === "/groups/g/threads/t/posts/p2"));
});

for (const hydrate of [false, true]) {
  test(`subject uses final revalidated thread (body hydration=${hydrate}) and attachment chain stays valid`, async () => {
    const client = fixture([post("p", hydrate ? { body: undefined } : { conversationThreadId: "verified" })], {
      "/groups/g/threads/verified/posts/p": { id: "p", conversationThreadId: "verified", singleValueExtendedProperties: subject("verified title") }
    });
    client.routes["/groups/g/conversations/c/threads/t/posts/p"] = post("p", { conversationThreadId: "verified" });
    client.routes["/groups/g/threads/verified/posts"] = { value: [post("p", { conversationThreadId: "verified" })] };
    client.routes["/groups/g/threads/verified/posts/p"] = { id: "p", conversationThreadId: "verified" };
    client.routes["/groups/g/threads/verified/posts/p/attachments"] = { value: [{ id: "a", name: "test.pdf", size: 20, "@odata.type": "#microsoft.graph.fileAttachment" }] };
    const result = await listGroupPosts(config, group, 20, options, client);
    const p = result.posts[0];
    assert.equal(p.threadVerification, "POST_COLLECTION_REVALIDATED");
    assert.equal(p.postSubject.value, "verified title");
    const attachments = await listGroupPostAttachments(config, { groupId: p.groupId, threadId: p.threadId, postId: p.id }, 20, {}, client);
    assert.equal(attachments.attachments[0].id, "a");
  });
}

test("subject identity contradictions fail closed and preserve preceding verified post", async () => {
  for (const mismatch of [{ id: "wrong" }, { id: "p2", conversationThreadId: "other" }, { id: "p2", conversationId: "other" }]) {
    const client = fixture([post("p"), post("p2")], {
      "/groups/g/threads/t/posts/p": { id: "p", singleValueExtendedProperties: subject("good") },
      "/groups/g/threads/t/posts/p2": { ...mismatch, singleValueExtendedProperties: subject("unsafe") }
    });
    const result = await listGroupPosts(config, group, 20, options, client);
    assert.deepEqual(result.posts.map(p => p.id), ["p"]);
    assert.match(result.page.failure, /GROUP_ID_CHAIN_MISMATCH/);
    assert.equal(result.page.complete, false);
    assert.equal(result.page.continuationAvailable, false);
  }
  await assert.rejects(readGroupPostSubject({ ...chain, revalidated: true }, async () => ({ id: "p", singleValueExtendedProperties: subject("unverified") })), /GROUP_ID_CHAIN_MISMATCH/);
});

test("subject API failures are sanitized, preserve bodies and do not imply subject completion", async () => {
  for (const error of ["Graph GET failed (403 Forbidden): PRIVATE_SECRET", "Graph GET failed (400 Bad Request): PRIVATE_SECRET", "network timeout PRIVATE_SECRET", "GROUP_REMOTE_PRIVATE_SECRET"]) {
    const client = fixture([post("p")], { "/groups/g/threads/t/posts/p": new Error(error) });
    const result = await listGroupPosts(config, group, 20, options, client);
    assert.equal(result.posts[0].postSubject.status, "READ_FAILED");
    assert.equal(result.posts[0].fullBodyAvailable, true);
    assert.equal(result.subjectRead.unavailableCount, 1);
    assert.equal(result.page.complete, true);
    assert.doesNotMatch(JSON.stringify(result), /PRIVATE_SECRET/);
    assert.equal(client.calls.filter(c => c.expanded).length, 1);
  }
});

test("budget exhaustion returns explicit unavailable title without losing any post on resume", async () => {
  const client = fixture([post("p"), post("p2")], { "/groups/g/threads/t/posts/p2": { id: "p2", singleValueExtendedProperties: subject("second") } });
  const first = await listGroupPosts(config, group, 20, { ...options, maxRequests: 3 }, client);
  assert.deepEqual(first.posts.map(p => p.id), ["p"]);
  assert.equal(first.posts[0].postSubject.status, "NOT_ATTEMPTED_BUDGET");
  assert.equal(first.page.requests, 3);
  const second = await listGroupPosts(config, group, 20, { ...options, maxRequests: 2, cursor: first.page.nextCursor }, client);
  assert.deepEqual(second.posts.map(p => p.id), ["p2"]);
  assert.equal(second.posts[0].postSubject.value, "second");
  assert.equal(second.page.requests, 2);
  const third = await listGroupPosts(config, group, 20, { ...options, cursor: second.page.nextCursor }, client);
  assert.equal(third.posts.length, 0);
  assert.equal(third.page.complete, true);
});

test("subject reads are not attempted for posts outside the requested range", async () => {
  const client = fixture([post("old", { createdDateTime: "2026-09-01T01:00:00Z" })]);
  const result = await listGroupPosts(config, group, 20, options, client);
  assert.equal(result.posts.length, 0);
  assert.equal(client.calls.filter(c => c.expanded).length, 0);
});

test("deadline exhaustion during body hydration retains the post with budget status", async () => {
  let clock = options.now();
  const client = fixture([post("p", { body: undefined }), post("p2")]);
  client.routes["/groups/g/conversations/c/threads/t/posts/p"] = () => {
    clock += 100;
    return post("p");
  };
  const result = await listGroupPosts(config, group, 20, { ...options, now: () => clock, timeBudgetMs: 50 }, client);
  assert.deepEqual(result.posts.map(p => p.id), ["p"]);
  assert.equal(result.posts[0].postSubject.status, "NOT_ATTEMPTED_BUDGET");
  assert.equal(result.posts[0].fullBodyAvailable, true);
  assert.equal(result.page.continuationAvailable, true);
  assert.deepEqual(result.page.partialReasons, ["TIME_BUDGET"]);
  assert.equal(client.calls.filter(c => c.expanded).length, 0);
});
