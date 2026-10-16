import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { downloadGroupPostAttachment, listGroupPostAttachments } from "../dist/groups-attachments.js";
import { authorizeDownload } from "../dist/download-approval.js";

const target = { groupId: "group/id", threadId: "thread+id", postId: "post=id" };
const postPath = "/groups/group%2Fid/threads/thread%2Bid/posts/post%3Did";
const collection = `${postPath}/attachments`;
const parent = { id: target.postId, conversationThreadId: target.threadId };
const file = (id = "attachment/id", extra = {}) => ({
  "@odata.type": "#microsoft.graph.fileAttachment", id, name: "fixture.txt", size: 5, ...extra
});

function config(t, policy = {}) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "hare-group-fixture-"));
  t.after(() => {
    assert.equal(path.dirname(path.resolve(root)), path.resolve(os.tmpdir()));
    assert.match(path.basename(root), /^hare-group-fixture-/);
    fs.rmSync(root, { recursive: true, force: true });
  });
  return {
    dataDir: root, cacheDir: path.join(root, "cache"), downloadDir: path.join(root, "downloads"),
    policy: { allowDownloads: true, maxSearchResults: 1000, maxDownloadBytes: 10, maxApprovedDownloadBytes: 100, ...policy }
  };
}

function clientFor({ post = parent, page = { value: [] }, metadata = file(), bytes = "hello", onDownload } = {}) {
  const calls = [];
  return {
    calls,
    async get(url, options) {
      calls.push({ kind: "get", url, options });
      assert.ok(options.totalTimeoutMs > 0 && options.totalTimeoutMs <= 60_000);
      assert.doesNotMatch(url, /\/(?:me|users)(?:\/|\?)/);
      if (!url.includes("/attachments")) return typeof post === "function" ? post(url) : post;
      if (new URL(url, "https://graph.microsoft.com").pathname.endsWith("/attachments")) {
        return typeof page === "function" ? page(url) : page;
      }
      return typeof metadata === "function" ? metadata(url) : metadata;
    },
    async download(url, options) {
      calls.push({ kind: "download", url, options });
      if (onDownload) return onDownload(url);
      return new Response(bytes);
    }
  };
}

test("group attachment listing verifies the exact parent and returns metadata only", async (t) => {
  const client = clientFor({ page: { value: [file(undefined, { contentBytes: "must-not-escape" }), file("ref", { "@odata.type": "#microsoft.graph.referenceAttachment" })] } });
  const result = await listGroupPostAttachments(config(t), target, undefined, {}, client);
  assert.equal(client.calls[0].url, `${postPath}?$select=id,conversationThreadId`);
  assert.match(client.calls[1].url, /\$top=20$/);
  assert.deepEqual(result.target, target);
  assert.equal(result.metadataVerified, true);
  assert.equal(result.list.returnedCount, 2);
  assert.equal(result.list.partial, false);
  assert.equal(result.list.continuationAvailable, false);
  assert.equal(result.list.stopReason, "COMPLETE");
  assert.equal(result.attachments[0].id, "attachment/id");
  assert.equal(result.attachments[0].attachmentType, "file");
  assert.equal("contentBytes" in result.attachments[0], false);
  assert.equal(result.attachments[1].downloadable, false);
});

for (const badParent of [{ ...parent, id: "wrong-post" }, { ...parent, conversationThreadId: "wrong-thread" }, { id: target.postId }, null]) {
  test(`parent chain mismatch blocks list and download: ${JSON.stringify(badParent)}`, async (t) => {
    const cfg = config(t);
    for (const operation of [
      (client) => listGroupPostAttachments(cfg, target, 20, {}, client),
      (client) => downloadGroupPostAttachment(cfg, target, "attachment/id", undefined, undefined, client)
    ]) {
      const client = clientFor({ post: badParent });
      await assert.rejects(() => operation(client), /GROUP_POSTS_CHAIN_MISMATCH/);
      assert.equal(client.calls.length, 1);
    }
  });
}

test("listing paginates and deduplicates attachment IDs", async (t) => {
  const client = clientFor({ page: (url) => url.includes("skiptoken")
    ? { value: [file("a"), file("b")] }
    : { value: [file("a")], "@odata.nextLink": `${collection}?$skiptoken=second` }
  });
  const result = await listGroupPostAttachments(config(t), target, 20, {}, client);
  assert.deepEqual(result.attachments.map((value) => value.id), ["a", "b"]);
  assert.equal(result.list.pagesRead, 2);
  assert.equal(result.list.partial, false);
});

test("overfull pages retain a replay position and deduplicate across continuations", async (t) => {
  const cfg = config(t, { maxSearchResults: 2 });
  const client = clientFor({ page: (url) => url.includes("skiptoken")
    ? { value: [file("c"), file("d"), file("e")] }
    : { value: [file("a"), file("a"), file("b"), file("c")], "@odata.nextLink": `${collection}?$skiptoken=second` }
  });
  const first = await listGroupPostAttachments(cfg, target, 999, {}, client);
  assert.deepEqual(first.attachments.map((value) => value.id), ["a", "b"]);
  assert.equal(first.list.maxResults, 2);
  assert.equal(first.list.limitReached, true);
  assert.equal(first.list.continuationAvailable, true);
  const second = await listGroupPostAttachments(cfg, target, 999, { cursor: first.list.nextCursor }, client);
  assert.deepEqual(second.attachments.map((value) => value.id), ["c", "d"]);
  const third = await listGroupPostAttachments(cfg, target, 999, { cursor: second.list.nextCursor }, client);
  assert.deepEqual(third.attachments.map((value) => value.id), ["e"]);
  assert.equal(third.list.partial, false);
  assert.equal(third.list.nextCursor, undefined);
  assert.equal(client.calls.filter(({ url }) => !url.includes("/attachments")).length, 3);
});

test("cursor binds all parent IDs and the effective result limit before any request", async (t) => {
  const cfg = config(t);
  const client = clientFor({ page: { value: [file("a"), file("b")] } });
  const first = await listGroupPostAttachments(cfg, target, 1, {}, client);
  for (const field of ["groupId", "threadId", "postId"]) {
    const fresh = clientFor();
    await assert.rejects(() => listGroupPostAttachments(cfg, { ...target, [field]: "other" }, 1, { cursor: first.list.nextCursor }, fresh), /INVALID_CURSOR/);
    assert.equal(fresh.calls.length, 0);
  }
  const fresh = clientFor();
  await assert.rejects(() => listGroupPostAttachments(cfg, target, 2, { cursor: first.list.nextCursor }, fresh), /INVALID_CURSOR/);
  assert.equal(fresh.calls.length, 0);
  const [version, payload, signature] = first.list.nextCursor.split(".");
  assert.equal(version, "g1");
  for (const mutate of [
    (envelope) => { envelope.state.url = "https://example.invalid/private"; },
    (envelope) => { envelope.state.offset += 1; },
    (envelope) => { envelope.state.seen.push(createHash("sha256").update("b").digest("hex")); },
    (envelope) => { envelope.expiresAt += 1; },
    (envelope) => { envelope.kind = "posts"; }
  ]) {
    const envelope = JSON.parse(Buffer.from(payload, "base64url").toString());
    mutate(envelope);
    const tampered = `${version}.${Buffer.from(JSON.stringify(envelope)).toString("base64url")}.${signature}`;
    await assert.rejects(() => listGroupPostAttachments(cfg, target, 1, { cursor: tampered }, fresh), /INVALID_CURSOR/);
    assert.equal(fresh.calls.length, 0);
  }
  await assert.rejects(() => listGroupPostAttachments(config(t), target, 1, { cursor: first.list.nextCursor }, fresh), /INVALID_CURSOR/);
  assert.equal(fresh.calls.length, 0);
});

test("replayed page changes fail explicitly rather than skipping new items", async (t) => {
  const cfg = config(t);
  const first = await listGroupPostAttachments(cfg, target, 1, {}, clientFor({ page: { value: [file("a"), file("b")] } }));
  await assert.rejects(() => listGroupPostAttachments(cfg, target, 1, { cursor: first.list.nextCursor }, clientFor({ page: { value: [file("new"), file("a"), file("b")] } })), /CURSOR_STALE/);
});

for (const link of [
  "https://example.invalid/attachments", "/me/messages/a/attachments", "/users/a/messages/b/attachments",
  "/groups/other/threads/thread%2Bid/posts/post%3Did/attachments",
  `${collection.replace("thread%2Bid", "other")}?$skiptoken=2`,
  `${collection.replace("post%3Did", "other")}?$skiptoken=2`,
  `https://user:password@graph.microsoft.com/v1.0${collection}`,
  `https://graph.microsoft.com/v1.0${collection}#secret`,
  `${collection}?$expand=item`, `${collection}?$filter=name%20eq%20'x'`,
  `${collection}?$select=contentBytes`, `${collection}?$top=1&$top=2`, "not-a-url"
]) {
  test(`bad continuation is rejected before following it: ${link}`, async (t) => {
    const client = clientFor({ page: { value: [file()], "@odata.nextLink": link } });
    await assert.rejects(() => listGroupPostAttachments(config(t), target, 20, {}, client), /GROUP_ATTACHMENT_LIST_INVALID_CONTINUATION/);
    assert.equal(client.calls.length, 2);
  });
}

test("repeated page URL stops partial without a looping cursor", async (t) => {
  const client = clientFor({ page: { value: [file("a")], "@odata.nextLink": `${collection}?$skiptoken=repeat` } });
  const result = await listGroupPostAttachments(config(t), target, 20, {}, client);
  assert.equal(result.list.pagesRead, 2);
  assert.equal(result.list.partial, true);
  assert.equal(result.list.stopReason, "NO_PROGRESS");
  assert.equal(result.list.nextCursor, undefined);
});

for (const limit of [20, 1]) {
  test(`duplicate-only intermediary page follows a new nextLink (limit=${limit})`, async (t) => {
    const cfg = config(t);
    const client = clientFor({ page: (url) => {
      const token = new URL(url).searchParams.get("$skiptoken");
      if (token === "third") return { value: [file("b")] };
      return { value: [file("a")], "@odata.nextLink": `${collection}?$skiptoken=${token === "second" ? "third" : "second"}` };
    } });
    const first = await listGroupPostAttachments(cfg, target, limit, {}, client);
    const results = [...first.attachments];
    const final = first.list.nextCursor
      ? await listGroupPostAttachments(cfg, target, limit, { cursor: first.list.nextCursor }, client)
      : first;
    if (final !== first) results.push(...final.attachments);
    assert.deepEqual(results.map(({ id }) => id), ["a", "b"]);
    assert.equal(final.list.partial, false);
    assert.equal(final.list.stopReason, "COMPLETE");
    assert.equal(final.list.continuationAvailable, false);
    assert.equal(client.calls.filter(({ url }) => url.includes("/attachments")).length, 3);
  });
}

test("duplicate-only pages with distinct URLs stop at the page budget", async (t) => {
  let pageNumber = 0;
  const client = clientFor({ page: () => ({ value: [file("a")], "@odata.nextLink": `${collection}?$skiptoken=${++pageNumber}` }) });
  const result = await listGroupPostAttachments(config(t), target, 20, {}, client);
  assert.deepEqual(result.attachments.map(({ id }) => id), ["a"]);
  assert.equal(result.list.pagesRead, 100);
  assert.equal(result.list.partial, true);
  assert.equal(result.list.stopReason, "PAGE_LIMIT");
  assert.equal(result.list.continuationAvailable, true);
});

test("empty nonterminal pages follow nextLink to later attachments", async (t) => {
  const client = clientFor({ page: (url) => url.includes("skiptoken")
    ? { value: [file("after-empty")] }
    : { value: [], "@odata.nextLink": `${collection}?$skiptoken=next` }
  });
  const result = await listGroupPostAttachments(config(t), target, 20, {}, client);
  assert.equal(result.list.partial, false);
  assert.equal(result.list.stopReason, "COMPLETE");
  assert.equal(result.list.pagesRead, 2);
  assert.deepEqual(result.attachments.map(({ id }) => id), ["after-empty"]);
});

test("empty-page cycles stop partial instead of repeating requests", async (t) => {
  const next = `${collection}?$skiptoken=repeat`;
  const result = await listGroupPostAttachments(config(t), target, 20, {}, clientFor({ page: { value: [], "@odata.nextLink": next } }));
  assert.equal(result.list.partial, true);
  assert.equal(result.list.stopReason, "NO_PROGRESS");
  assert.equal(result.list.pagesRead, 2);
  assert.equal(result.list.continuationAvailable, false);
});

test("empty nonterminal pages remain bounded by the page budget", async (t) => {
  let index = 0;
  const result = await listGroupPostAttachments(config(t), target, 20, {}, clientFor({ page: () => ({ value: [], "@odata.nextLink": `${collection}?$skiptoken=${++index}` }) }));
  assert.equal(result.list.stopReason, "PAGE_LIMIT");
  assert.equal(result.list.pagesRead, 100);
  assert.equal(result.list.continuationAvailable, true);
});

test("later page failure preserves results and a safe exact retry cursor", async (t) => {
  const cfg = config(t);
  const next = `${collection}?$skiptoken=retry`;
  const client = clientFor({ page: (url) => {
    if (url.includes("skiptoken")) throw new Error("Graph GET failed (403 Forbidden): synthetic-private-body");
    return { value: [file("a")], "@odata.nextLink": next };
  } });
  const partial = await listGroupPostAttachments(cfg, target, 20, {}, client);
  assert.deepEqual(partial.attachments.map(({ id }) => id), ["a"]);
  assert.equal(partial.list.partial, true);
  assert.equal(partial.list.stopReason, "READ_FAILED");
  assert.equal(partial.list.continuationAvailable, true);
  assert.equal(partial.list.failure.stage, "ATTACHMENT_LIST");
  assert.match(partial.list.failure.message, /^GROUP_ATTACHMENT_LIST_ACCESS_DENIED:/);
  assert.doesNotMatch(JSON.stringify(partial), /synthetic-private-body/);
  const retryClient = clientFor({ page: { value: [file("a"), file("b")] } });
  const retried = await listGroupPostAttachments(cfg, target, 20, { cursor: partial.list.nextCursor }, retryClient);
  assert.equal(retryClient.calls[1].url, `https://graph.microsoft.com/v1.0${next}`);
  assert.deepEqual(retried.attachments.map(({ id }) => id), ["b"]);
  assert.equal(retried.list.partial, false);
  assert.equal(retried.list.continuationAvailable, false);
  assert.equal(retried.list.failure, undefined);
});

test("failure after empty pages with no results still throws", async (t) => {
  const client = clientFor({ page: (url) => {
    if (url.includes("skiptoken")) throw new Error("Graph GET failed (403 Forbidden): synthetic-private-body");
    return { value: [], "@odata.nextLink": `${collection}?$skiptoken=next` };
  } });
  await assert.rejects(() => listGroupPostAttachments(config(t), target, 20, {}, client), /GROUP_ATTACHMENT_LIST_ACCESS_DENIED/);
});

test("page budget stops and offers an exact continuation", async (t) => {
  let index = 0;
  const client = clientFor({ page: () => ({ value: [file(`a-${++index}`)], "@odata.nextLink": `${collection}?$skiptoken=${index}` }) });
  const result = await listGroupPostAttachments(config(t), target, 1000, {}, client);
  assert.equal(result.list.pagesRead, 100);
  assert.equal(result.list.returnedCount, 100);
  assert.equal(result.list.stopReason, "PAGE_LIMIT");
  assert.ok(result.list.nextCursor);
});

test("elapsed time budget returns partial without another page request", async (t) => {
  const cfg = config(t);
  const originalNow = Date.now;
  let now = originalNow();
  t.after(() => { Date.now = originalNow; });
  Date.now = () => now;
  const client = clientFor({ page: () => {
    now += 60_001;
    return { value: [file()], "@odata.nextLink": `${collection}?$skiptoken=next` };
  } });
  const result = await listGroupPostAttachments(cfg, target, 20, {}, client);
  assert.equal(result.list.pagesRead, 1);
  assert.equal(result.list.stopReason, "TIME_LIMIT");
  assert.ok(result.list.nextCursor);
});

for (const stage of ["POSTS", "ATTACHMENT_LIST", "ATTACHMENT_METADATA", "ATTACHMENT_BYTES"]) {
  test(`${stage} 403 is distinct and does not expose the error body`, async (t) => {
    const fail = () => { throw new Error("Graph GET failed (403 Forbidden): synthetic-private-body"); };
    const client = clientFor({
      ...(stage === "POSTS" ? { post: fail } : {}),
      ...(stage === "ATTACHMENT_LIST" ? { page: fail } : {}),
      ...(stage === "ATTACHMENT_METADATA" ? { metadata: fail } : {}),
      ...(stage === "ATTACHMENT_BYTES" ? { onDownload: fail } : {})
    });
    const cfg = config(t);
    await assert.rejects(() => stage === "POSTS" || stage === "ATTACHMENT_LIST"
      ? listGroupPostAttachments(cfg, target, 20, {}, client)
      : downloadGroupPostAttachment(cfg, target, "attachment/id", undefined, undefined, client), (error) => {
      assert.match(error.message, new RegExp(`GROUP_${stage}_ACCESS_DENIED`));
      assert.doesNotMatch(error.message, /synthetic-private-body/);
      return true;
    });
  });
}

for (const type of ["file", "item"]) {
  test(`${type} downloads raw bytes and reports metadata size separately`, async (t) => {
    const cfg = config(t);
    const client = clientFor({ metadata: file(undefined, { "@odata.type": `#microsoft.graph.${type}Attachment`, size: 8 }) });
    const result = await downloadGroupPostAttachment(cfg, target, "attachment/id", undefined, undefined, client);
    assert.equal(result.stage, "DOWNLOADED");
    assert.equal(result.metadataVerified, true);
    assert.equal(result.structureValidation, "NOT_PERFORMED");
    assert.deepEqual(result.target, target);
    assert.equal(result.attachment.size, 8);
    assert.equal(result.downloadedBytes, 5);
    assert.equal(fs.statSync(result.outputPath).size, 5);
    assert.equal(fs.readFileSync(result.outputPath, "utf8"), "hello");
    assert.equal(client.calls.at(-1).url, `${collection}/attachment%2Fid/$value`);
    assert.equal(client.calls.some(({ url }) => url.includes("$expand")), false);
  });
}

test("attachment metadata ID mismatch blocks raw bytes", async (t) => {
  const client = clientFor({ metadata: file("another") });
  await assert.rejects(() => downloadGroupPostAttachment(config(t), target, "attachment/id", undefined, undefined, client), /ATTACHMENT_METADATA_CHAIN_MISMATCH/);
  assert.equal(client.calls.length, 2);
});

for (const type of ["reference", "unknown", "fakefile"]) {
  test(`${type} attachments do not trigger bytes or expansion fallback`, async (t) => {
    const client = clientFor({ metadata: file(undefined, { "@odata.type": `#microsoft.graph.${type}Attachment` }) });
    await assert.rejects(() => downloadGroupPostAttachment(config(t), target, "attachment/id", undefined, undefined, client), /TYPE_UNSUPPORTED/);
    assert.equal(client.calls.length, 2);
  });
}

test("size approval previews target and attachment before download", async (t) => {
  const cfg = config(t);
  const client = clientFor({ metadata: file(undefined, { size: 20 }), bytes: "approved raw bytes" });
  const pending = await downloadGroupPostAttachment(cfg, target, "attachment/id", "approved.txt", undefined, client);
  assert.equal(pending.stage, "AWAITING_USER_APPROVAL");
  assert.equal(pending.preview.source, "group");
  assert.deepEqual(pending.preview.target, target);
  assert.equal(pending.preview.attachmentId, "attachment/id");
  assert.equal(pending.preview.outputName, "approved.txt");
  assert.equal(pending.preview.size, 20);
  assert.equal(client.calls.some(({ kind }) => kind === "download"), false);
  const result = await downloadGroupPostAttachment(cfg, target, "attachment/id", "approved.txt", pending.approval.token, client);
  assert.equal(result.stage, "DOWNLOADED");
  assert.equal(result.downloadedBytes, Buffer.byteLength("approved raw bytes"));
  await assert.rejects(() => downloadGroupPostAttachment(cfg, target, "attachment/id", "approved.txt", pending.approval.token, client), /DOWNLOAD_APPROVAL_REQUIRED/);
});

for (const changed of ["groupId", "threadId", "postId", "attachmentId", "output", "size", "source"]) {
  test(`approval cannot be reused with changed ${changed}`, async (t) => {
    const cfg = config(t);
    const metadata = file(undefined, { size: 20 });
    const pending = await downloadGroupPostAttachment(cfg, target, metadata.id, "approved.txt", undefined, clientFor({ metadata }));
    if (changed === "source") {
      assert.throws(() => authorizeDownload(cfg, {
        source: "outlook", resourceKey: JSON.stringify([target.groupId, target.threadId, target.postId, metadata.id]),
        name: metadata.name, size: metadata.size, outputName: "approved.txt"
      }, pending.approval.token), /approved file does not match/);
      return;
    }
    const nextTarget = { ...target, ...(["groupId", "threadId", "postId"].includes(changed) ? { [changed]: "other" } : {}) };
    const nextId = changed === "attachmentId" ? "other" : metadata.id;
    const client = clientFor({
      post: { id: nextTarget.postId, conversationThreadId: nextTarget.threadId },
      metadata: { ...metadata, id: nextId, size: changed === "size" ? 21 : metadata.size }
    });
    await assert.rejects(() => downloadGroupPostAttachment(cfg, nextTarget, nextId, changed === "output" ? "other.txt" : "approved.txt", pending.approval.token, client), /approved file does not match/);
    assert.equal(client.calls.some(({ kind }) => kind === "download"), false);
  });
}

test("approval resource key is a JSON tuple of all opaque IDs", async (t) => {
  const cfg = config(t);
  const metadata = file(undefined, { size: 20 });
  const pending = await downloadGroupPostAttachment(cfg, target, metadata.id, undefined, undefined, clientFor({ metadata }));
  const authorization = authorizeDownload(cfg, {
    source: "group", resourceKey: JSON.stringify([target.groupId, target.threadId, target.postId, metadata.id]),
    name: metadata.name, size: metadata.size, outputName: metadata.name
  }, pending.approval.token);
  assert.equal(authorization.approved, true);
});

for (const size of [undefined, -1, NaN, Infinity, "5", 101]) {
  test(`invalid or excessive metadata size blocks raw content: ${String(size)}`, async (t) => {
    const client = clientFor({ metadata: file(undefined, { size }) });
    await assert.rejects(() => downloadGroupPostAttachment(config(t), target, "attachment/id", undefined, undefined, client), /Download (?:size metadata is required|exceeds approved policy limit)/);
    assert.equal(client.calls.some(({ kind }) => kind === "download"), false);
  });
}

for (const withLength of [false, true]) {
  test(`raw bytes enforce the default policy even when metadata is smaller (length=${withLength})`, async (t) => {
    const cfg = config(t);
    const client = clientFor({ onDownload: () => new Response("x".repeat(11), { headers: withLength ? { "content-length": "11" } : {} }) });
    await assert.rejects(() => downloadGroupPostAttachment(cfg, target, "attachment/id", undefined, undefined, client), /Download exceeds policy limit of 10 bytes/);
    if (fs.existsSync(cfg.downloadDir)) assert.deepEqual(fs.readdirSync(cfg.downloadDir), []);
  });
}

test("approved download still enforces the approved raw-byte limit", async (t) => {
  const cfg = config(t);
  const client = clientFor({ metadata: file(undefined, { size: 20 }), bytes: "x".repeat(101) });
  const pending = await downloadGroupPostAttachment(cfg, target, "attachment/id", undefined, undefined, client);
  await assert.rejects(() => downloadGroupPostAttachment(cfg, target, "attachment/id", undefined, pending.approval.token, client), /Download exceeds policy limit of 100 bytes/);
});

test("safe filename and unique output preserve existing files", async (t) => {
  const cfg = config(t);
  const client = clientFor();
  const first = await downloadGroupPostAttachment(cfg, target, "attachment/id", "../fixture.txt", undefined, client);
  fs.writeFileSync(first.outputPath, "preserve-existing-fixture");
  const second = await downloadGroupPostAttachment(cfg, target, "attachment/id", "../fixture.txt", undefined, client);
  assert.notEqual(first.outputPath, second.outputPath);
  assert.equal(path.dirname(first.outputPath), cfg.downloadDir);
  assert.equal(path.dirname(second.outputPath), cfg.downloadDir);
  assert.equal(fs.readFileSync(first.outputPath, "utf8"), "preserve-existing-fixture");
  assert.equal(fs.readFileSync(second.outputPath, "utf8"), "hello");
});

test("local saving errors are safe and distinct from Graph and policy errors", async (t) => {
  const cfg = config(t);
  fs.writeFileSync(cfg.downloadDir, "not-a-directory");
  await assert.rejects(() => downloadGroupPostAttachment(cfg, target, "attachment/id", undefined, undefined, clientFor()), (error) => {
    assert.match(error.message, /^GROUP_ATTACHMENT_SAVE_FAILED:/);
    assert.equal(error.message.includes(cfg.dataDir), false);
    return true;
  });
});

test("disabled downloads and invalid identifiers make no network requests", async (t) => {
  const client = clientFor();
  await assert.rejects(() => downloadGroupPostAttachment(config(t, { allowDownloads: false }), target, "attachment/id", undefined, undefined, client), /disabled by policy/);
  const cfg = config(t);
  for (const id of ["", "..", ".", " with-space", "control\u0000id"]) {
    await assert.rejects(() => downloadGroupPostAttachment(cfg, { ...target, postId: id }, "attachment/id", undefined, undefined, client), /INVALID_IDENTIFIER/);
  }
  await assert.rejects(() => listGroupPostAttachments(cfg, target, 0, {}, client), /positive number/);
  assert.equal(client.calls.length, 0);
});
