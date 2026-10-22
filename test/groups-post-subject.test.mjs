import assert from "node:assert/strict";
import test, { after } from "node:test";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { listGroupPosts } from "../dist/groups-posts.js";
import { readGroupPostSubject } from "../dist/groups-post-subject.js";
import { readGroupCursorInput } from "../dist/groups-cursor-input.js";
import { listGroupPostAttachments } from "../dist/groups-attachments.js";
import { decodeGroupCursor, encodeGroupCursor } from "../dist/groups-cursor.js";
import { encodePostContinuation } from "../dist/groups-post-checkpoint.js";
import { createHash } from "node:crypto";
import { authFailure } from "../dist/auth-diagnostics.js";

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

test("terminal subject API failures are sanitized and never consume the pending post", async () => {
  for (const error of ["Graph GET failed (403 Forbidden): PRIVATE_SECRET", "Graph GET failed (400 Bad Request): PRIVATE_SECRET", "network timeout PRIVATE_SECRET", "GROUP_REMOTE_PRIVATE_SECRET"]) {
    const client = fixture([post("p")], { "/groups/g/threads/t/posts/p": new Error(error) });
    const result = await listGroupPosts(config, group, 20, options, client);
    assert.equal(result.posts.length, 0);
    assert.equal(result.page.scannedPosts, 0);
    assert.equal(result.page.failureDetails.stage, "POST_SUBJECT");
    assert.equal(result.page.continuationAvailable, false);
    assert.equal(result.page.complete, false);
    assert.doesNotMatch(JSON.stringify(result), /PRIVATE_SECRET/);
    assert.equal(client.calls.filter(c => c.expanded).length, 1);
  }
});

test("budget exhaustion defers the same post until its subject is read", async () => {
  const client = fixture([post("p"), post("p2")], {
    "/groups/g/threads/t/posts/p": { id: "p", singleValueExtendedProperties: subject("first") },
    "/groups/g/threads/t/posts/p2": { id: "p2", singleValueExtendedProperties: subject("second") }
  });
  const first = await listGroupPosts(config, group, 20, { ...options, maxRequests: 3 }, client);
  assert.deepEqual(first.posts, []);
  assert.deepEqual(first.page.partialReasons, ["SUBJECT_BUDGET"]);
  assert.equal(first.page.scannedPosts, 0);
  assert.equal(first.page.requests, 3);
  const second = await listGroupPosts(config, group, 20, { ...options, maxRequests: 2, cursor: first.page.nextCursor }, client);
  assert.deepEqual(second.posts.map(p => p.id), ["p"]);
  assert.equal(second.posts[0].postSubject.value, "first");
  assert.equal(second.page.requests, 2);
  const third = await listGroupPosts(config, group, 20, { ...options, cursor: second.page.nextCursor }, client);
  assert.deepEqual(third.posts.map(p => p.id), ["p2"]);
  assert.equal(third.posts[0].postSubject.value, "second");
  assert.equal(third.page.complete, true);
});

test("subject reads are not attempted for posts outside the requested range", async () => {
  const client = fixture([post("old", { createdDateTime: "2026-09-01T01:00:00Z" })]);
  const result = await listGroupPosts(config, group, 20, options, client);
  assert.equal(result.posts.length, 0);
  assert.equal(client.calls.filter(c => c.expanded).length, 0);
});

test("deadline exhaustion during hydration leaves the post pending for subject resume", async () => {
  let clock = options.now();
  const client = fixture([post("p", { body: undefined })], { "/groups/g/threads/t/posts/p": { id: "p", singleValueExtendedProperties: subject("recovered") } });
  client.routes["/groups/g/conversations/c/threads/t/posts/p"] = () => {
    clock += 100;
    return post("p");
  };
  const result = await listGroupPosts(config, group, 20, { ...options, now: () => clock, timeBudgetMs: 50 }, client);
  assert.deepEqual(result.posts, []);
  assert.equal(result.page.continuationAvailable, true);
  assert.deepEqual(result.page.partialReasons, ["SUBJECT_BUDGET"]);
  assert.equal(client.calls.filter(c => c.expanded).length, 0);
  client.routes["/groups/g/threads/t/posts/p"] = post("p");
  const resumed = await listGroupPosts(config, group, 20, { ...options, cursor: result.page.nextCursor }, client);
  assert.equal(resumed.posts[0].postSubject.value, "recovered");
  assert.equal(resumed.posts[0].fullBodyAvailable, true);
  assert.equal(resumed.page.complete, true);
});

test("subject option cannot be dropped on continuation", async () => {
  const client = fixture([post("p")]);
  const first = await listGroupPosts(config, group, 20, { ...options, maxRequests: 3 }, client);
  await assert.rejects(listGroupPosts(config, group, 20, { ...options, includeSubject: false, cursor: first.page.nextCursor }, client), /GROUP_SUBJECT_OPTION_REQUIRED/);
});

test("safe token failure provenance survives the full subject error chain without retry or consumption", async () => {
  const error = authFailure("AUTH_CACHE_CHANGED", "commit", { message: "PRIVATE", code: "EPERM" });
  const client = fixture([post("p")], { "/groups/g/threads/t/posts/p": error });
  const result = await listGroupPosts(config, group, 20, options, client);
  assert.equal(result.posts.length, 0);
  assert.equal(result.page.continuationAvailable, false);
  assert.equal(result.page.failureDetails.retryAllowed, false);
  assert.deepEqual(result.page.failureDetails.authProvenance, { phase: "commit", reason: "AUTH_CACHE_CHANGED", ioCode: "EPERM" });
  assert.doesNotMatch(JSON.stringify(result), /PRIVATE/);
});

test("same-position subject budget deferrals stop after three attempts without consuming", async () => {
  let clock = options.now();
  const client = fixture([post("p")]);
  const first = await listGroupPosts(config, group, 20, { ...options, maxRequests: 3 }, client);
  client.routes["/groups/g/threads/t/posts/p"] = () => { clock += 100; return post("p"); };
  const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor, now: () => clock, timeBudgetMs: 50 }, client);
  assert.equal(second.page.continuationAvailable, true);
  assert.equal(second.page.noProgressDetected, false);
  const third = await listGroupPosts(config, group, 20, { ...options, cursor: second.page.nextCursor, now: () => clock, timeBudgetMs: 50 }, client);
  assert.deepEqual(third.page.partialReasons, ["SUBJECT_BUDGET_RETRY_LIMIT"]);
  assert.equal(third.page.continuationAvailable, false);
  assert.equal(third.page.scannedPosts, 0);
});

for (const failure of ["Graph GET failed (429 Too Many Requests)", "Graph GET failed (503 Service Unavailable)", "Graph request exceeded the 1ms time budget", "fetch failed ECONNRESET"]) {
  test(`transient subject failure resumes the same post: ${failure}`, async () => {
    const client = fixture([post("p")], { "/groups/g/threads/t/posts/p": new Error(failure) });
    const first = await listGroupPosts(config, group, 20, options, client);
    assert.equal(first.page.failureDetails.retryAllowed, true);
    assert.equal(first.page.failureDetails.stage, "POST_SUBJECT");
    assert.equal(first.posts.length, 0);
    client.subjects["/groups/g/threads/t/posts/p"] = { id: "p", singleValueExtendedProperties: subject("recovered") };
    const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, client);
    assert.equal(second.posts[0].postSubject.value, "recovered");
    assert.equal(second.page.complete, true);
  });
}

test("repeated transient subject failures stop at the existing retry limit", async () => {
  const client = fixture([post("p")], { "/groups/g/threads/t/posts/p": new Error("Graph GET failed (503 Service Unavailable)") });
  let result;
  for (let attempt = 1; attempt <= 3; attempt++) {
    result = await listGroupPosts(config, group, 20, { ...options, cursor: result?.page.nextCursor }, client);
    assert.equal(result.page.failureDetails.consecutiveAttempts, attempt);
    assert.equal(result.page.scannedPosts, 0);
  }
  assert.deepEqual(result.page.partialReasons, ["TRANSIENT_RETRY_LIMIT"]);
  assert.equal(result.page.continuationAvailable, false);
});

test("alternating budget and transient failures share a bounded pending-position allowance", async () => {
  let clock = options.now();
  const client = fixture([post("p")], { "/groups/g/threads/t/posts/p": new Error("Graph GET failed (503 Service Unavailable)") });
  const first = await listGroupPosts(config, group, 20, { ...options, maxRequests: 3 }, client);
  const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, client);
  assert.equal(second.page.continuationAvailable, true);
  client.routes["/groups/g/threads/t/posts/p"] = () => { clock += 100; return post("p"); };
  const third = await listGroupPosts(config, group, 20, { ...options, cursor: second.page.nextCursor, now: () => clock, timeBudgetMs: 50 }, client);
  assert.deepEqual(third.page.partialReasons, ["SUBJECT_BUDGET_RETRY_LIMIT"]);
  assert.equal(third.page.continuationAvailable, false);
  assert.equal(third.posts.length, 0);
});

for (const local of [false, true]) {
  test(`signed subject retry state is validated before reads (local=${local})`, async () => {
    const client = fixture([post("p")]);
    const first = await listGroupPosts(config, group, 20, { ...options, maxRequests: 3 }, client);
    const original = decodeGroupCursor(config, "posts", first.page.nextCursor);
    for (const mutate of [
      s => { s.subjectDeferral.position = [s.subjectDeferral.position]; },
      s => { s.retry = { position: [s.subjectDeferral.position], attempts: 1 }; },
      s => { s.subjectDeferral.position = "f".repeat(64); },
      s => { s.retry = { position: "f".repeat(64), attempts: 1 }; },
      s => { s.retry = { ...s.subjectDeferral }; s.subjectDeferral.attempts = 2; },
      s => { s.subjectRequired = false; },
      s => { s.missingSubjects = -1; }
    ]) {
      const state = structuredClone(original);
      mutate(state);
      if (local) state.seen.push(...Array.from({ length: 8200 }, (_, i) => createHash("sha256").update(`filler${i}`).digest("hex")));
      const cursor = local ? encodePostContinuation(config, state) : encodeGroupCursor(config, "posts", state);
      client.calls.length = 0;
      await assert.rejects(listGroupPosts(config, group, 20, { ...options, cursor }, client), /GROUP_INVALID_CURSOR/);
      assert.equal(client.calls.length, 0);
    }
  });
}

for (const [properties, status] of [[[], "MISSING"], [undefined, "UNVERIFIED"], [null, "INVALID_RESPONSE"]]) {
  test(`${status} is explicit and prevents completion across pages`, async () => {
    const client = fixture([post("p"), post("p2")], {
      "/groups/g/threads/t/posts/p": { id: "p", singleValueExtendedProperties: properties },
      "/groups/g/threads/t/posts/p2": { id: "p2", singleValueExtendedProperties: subject("") }
    });
    const first = await listGroupPosts(config, group, 1, options, client);
    assert.equal(first.posts[0].postSubject.status, status);
    const second = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, client);
    assert.equal(second.posts[0].postSubject.status, "EMPTY");
    assert.equal(second.subjectRead.cumulativeUnavailableCount, 1);
    assert.equal(second.page.complete, false);
    assert.equal(second.page.continuationAvailable, false);
    assert.deepEqual(second.page.partialReasons, ["SUBJECT_UNAVAILABLE"]);
  });
}
