import assert from "node:assert/strict";
import test, { after } from "node:test";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createHash } from "node:crypto";
import { listGroupPosts } from "../dist/groups-posts.js";
import { encodeGroupCursor, decodeGroupCursor } from "../dist/groups-cursor.js";
import { encodePostContinuation, decodePostContinuation } from "../dist/groups-post-checkpoint.js";
import { readGroupCursorInput } from "../dist/groups-cursor-input.js";
import { groupReadError, retryableGroupReadFailure } from "../dist/groups-common.js";

const root = fs.mkdtempSync(path.join(os.tmpdir(), "hare-history-test-"));
after(() => {
  assert.equal(path.dirname(path.resolve(root)), path.resolve(os.tmpdir()));
  assert.match(path.basename(root), /^hare-history-test-/);
  fs.rmSync(root, { recursive: true, force: true });
});
const configFor = name => ({ cacheDir: path.join(root, name), policy: { defaultSearchLookbackDays: 90, maxSearchResults: 1000 } });
const group = { id: "g", groupTypes: ["Unified"], classification: "microsoft365" };
const options = { since: "2026-01-01", until: "2026-08-15" };
const hash = x => createHash("sha256").update(String(x)).digest("hex");
const post = id => ({ id, conversationId: "c", conversationThreadId: "t", createdDateTime: "2026-07-01T00:00:00Z", body: { contentType: "text", content: "synthetic full body" } });
const state = (count = 9000) => ({ version: 3, criteria: "query", jobs: [], seen: Array.from({ length: count }, (_, i) => hash(i)), pages: [], missingBodies: 0, missingTimestamps: 0 });
const envelope = cursor => JSON.parse(Buffer.from(cursor.split(".")[1], "base64url").toString()).state;

test("chunked checkpoints preserve exact histories and old versions with bounded cursors", () => {
  const config = configFor("immutable");
  const first = state();
  const cursor = encodePostContinuation(config, first);
  assert.equal(envelope(cursor).version, 4);
  assert.ok(cursor.length < 1000);
  const later = { ...first, seen: [...first.seen, ...Array.from({ length: 3000 }, (_, i) => hash(i + 9000))] };
  const next = encodePostContinuation(config, later);
  assert.deepEqual(decodePostContinuation(config, cursor, "query").state, first);
  assert.deepEqual(decodePostContinuation(config, next, "query").state, later);
  const names = fs.readdirSync(path.join(config.cacheDir, "group-post-checkpoints"));
  assert.equal(names.length, Math.floor(later.seen.length / 1024) + 2);
  assert.throws(() => decodePostContinuation(config, next, "another-query"), /GROUP_INVALID_CHECKPOINT/);
  assert.throws(() => decodePostContinuation(configFor("other-cache"), next, "query"), /GROUP_INVALID_CURSOR/);
});

test("missing, corrupted or reordered checkpoint chunks fail closed; signed path escape rejected", () => {
  for (const mode of ["missing", "corrupt", "reordered"]) {
    const config = configFor(mode);
    const cursor = encodePostContinuation(config, state());
    const dir = path.join(config.cacheDir, "group-post-checkpoints");
    const cp = JSON.parse(fs.readFileSync(path.join(dir, envelope(cursor).checkpoint + ".json"), "utf8"));
    const file = path.join(dir, cp.seen.chunks[0] + ".json");
    if (mode === "missing") fs.rmSync(file);
    if (mode === "corrupt") fs.writeFileSync(file, "broken");
    if (mode === "reordered") fs.writeFileSync(file, JSON.stringify(JSON.parse(fs.readFileSync(file, "utf8")).reverse()));
    assert.throws(() => decodePostContinuation(config, cursor, "query"), /GROUP_INVALID_CHECKPOINT/);
  }
  const config = configFor("escape");
  const cursor = encodeGroupCursor(config, "posts", { version: 4, criteria: "query", checkpoint: "../outside" });
  assert.throws(() => decodePostContinuation(config, cursor, "query"), /GROUP_INVALID_CHECKPOINT/);
});

test("storage quota and corrupt existing objects never replace previous committed state", () => {
  const config = configFor("quota");
  const cursor = encodePostContinuation(config, state());
  const dir = path.join(config.cacheDir, "group-post-checkpoints");
  const fd = fs.openSync(path.join(dir, "reserved.tmp"), "wx");
  fs.ftruncateSync(fd, 256 * 1024 * 1024);
  fs.closeSync(fd);
  assert.throws(() => encodePostContinuation(config, state(11000)), /GROUP_CHECKPOINT_STORAGE_BLOCKED/);
  assert.deepEqual(decodePostContinuation(config, cursor, "query").state, state());
});

test("actual traversal crosses 10000 seen identities with exact dedup and file resume", async () => {
  const config = configFor("large-seen");
  const total = 12030;
  const client = { async get(raw) {
    const u = new URL(raw.startsWith("/") ? `https://graph.microsoft.com/v1.0${raw}` : raw);
    const p = u.pathname;
    if (p.endsWith("/conversations")) return { value: [{ id: "c" }] };
    if (p.endsWith("/threads")) return { value: [{ id: "t" }] };
    if (/\/posts\/p\d+$/.test(p)) return post(p.split("/").at(-1));
    const start = Number(u.searchParams.get("start") ?? 0);
    const rows = Array.from({ length: Math.min(500, total - start) }, (_, i) => post(`p${start + i}`));
    if (start > 0) rows.unshift(post("p0"));
    return { value: rows, ...(start + 500 < total ? { "@odata.nextLink": `https://graph.microsoft.com/v1.0/groups/g/conversations/c/threads/t/posts?start=${start + 500}` } : {}) };
  } };
  const unique = new Set();
  let cursor;
  let localSeen = false;
  let result;
  for (let call = 0; call < 160; call++) {
    result = await listGroupPosts(config, group, 100, { ...options, cursor }, client);
    assert.equal(result.page.failure, undefined);
    assert.equal(result.page.partialReasons.includes("HISTORY_LIMIT"), false);
    for (const p of result.posts) { assert.equal(unique.has(p.id), false); unique.add(p.id); }
    if (result.page.history.seenEntries > 10000) localSeen = true;
    if (!result.page.continuationAvailable) break;
    const file = path.join(root, "progress.json");
    fs.writeFileSync(file, JSON.stringify({ page: { nextCursor: result.page.nextCursor } }));
    cursor = readGroupCursorInput({ cursorFile: file });
  }
  assert.equal(result.page.complete, true);
  assert.equal(localSeen, true);
  assert.equal(unique.size, total);
});

test("more than 10000 visited pages can complete without weakening cycle detection", async () => {
  const config = configFor("large-pages");
  const client = { async get(raw) {
    const u = new URL(raw.startsWith("/") ? `https://graph.microsoft.com/v1.0${raw}` : raw);
    const n = Number(u.searchParams.get("n") ?? 0);
    return { value: [], ...(n < 10005 ? { "@odata.nextLink": `https://graph.microsoft.com/v1.0/groups/g/conversations?n=${n + 1}` } : {}) };
  } };
  let cursor, result;
  for (let i = 0; i < 110; i++) {
    result = await listGroupPosts(config, group, 100, { ...options, cursor }, client);
    if (!result.page.continuationAvailable) break;
    cursor = result.page.nextCursor;
  }
  assert.equal(result.page.complete, true);
  assert.equal(result.page.history.visitedPages, 10005);
});

function faultClient(mode, failures) {
  return { async get(raw) {
    const u = new URL(raw.startsWith("/") ? `https://graph.microsoft.com/v1.0${raw}` : raw);
    const p = u.pathname;
    if (p.endsWith("/conversations")) return { value: [{ id: "c" }] };
    if (p.endsWith("/threads")) return { value: [{ id: "t" }] };
    if (p.endsWith("/posts")) {
      if (mode === "collection" && failures.left-- > 0) throw failures.error;
      return { value: [mode === "hydrate" ? { ...post("p1"), body: undefined } : post("p1"), post("p2")] };
    }
    if (failures.left-- > 0) throw failures.error;
    return post(p.split("/").at(-1));
  } };
}
for (const mode of ["collection", "detail", "hydrate"]) {
  test(`transient ${mode} failure resumes the unresolved ID without skip`, async () => {
    const config = configFor(mode + "retry");
    const failures = { left: mode === "detail" ? 0 : 1, error: new Error("Graph GET failed (503 Unavailable): PRIVATE_RESPONSE") };
    const client = faultClient(mode, failures);
    let first = await listGroupPosts(config, group, mode === "detail" ? 1 : 20, options, client);
    if (mode === "detail") {
      failures.left = 1;
      first = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, client);
    }
    assert.equal(first.page.failureDetails.retryAllowed, true);
    assert.equal(first.page.continuationAvailable, true);
    assert.equal(first.page.noProgressDetected, false);
    assert.doesNotMatch(JSON.stringify(first), /PRIVATE_RESPONSE/);
    const resumed = await listGroupPosts(config, group, 20, { ...options, cursor: first.page.nextCursor }, client);
    assert.deepEqual(resumed.posts.map(p => p.id), mode === "detail" ? ["p2"] : ["p1", "p2"]);
    assert.equal(resumed.page.complete, true);
  });
}

test("retry lineage stops after three failures at the same position", async () => {
  const config = configFor("retry-limit");
  const client = faultClient("collection", { left: 100, error: new Error("Graph request exceeded the 20ms time budget.") });
  let cursor;
  for (let i = 1; i <= 3; i++) {
    const result = await listGroupPosts(config, group, 20, { ...options, cursor }, client);
    assert.equal(result.page.failureDetails.consecutiveAttempts, i);
    assert.equal(result.page.failureDetails.retryAllowed, i < 3);
    assert.equal(result.page.continuationAvailable, i < 3);
    assert.equal(result.page.complete, false);
    if (i === 3) assert.ok(result.page.partialReasons.includes("TRANSIENT_RETRY_LIMIT"));
    cursor = result.page.nextCursor;
  }
});

test("HTTP status dominates body words; unknown/auth/access/identity failures are not retryable", () => {
  for (const message of ["Graph GET failed (403 Forbidden): fetch failed network timeout", "Graph GET failed (401 Unauthorized): socket error", "Graph GET failed (404 Not Found): ECONNRESET", "unknown PRIVATE_RESPONSE", "AUTH_REQUIRED: fetch failed", "fetch failed ENOTFOUND", "fetch failed EACCES", "fetch failed blocked-by-allowlist"]) {
    const e = groupReadError("POSTS", new Error(message));
    assert.equal(retryableGroupReadFailure(e.message), false);
    assert.doesNotMatch(e.message, /PRIVATE_RESPONSE/);
  }
  for (const message of ["Graph GET failed (429 Throttled): private", "Graph GET failed (503 Unavailable): private", "Graph request exceeded the 100ms time budget.", "fetch failed ECONNRESET"]) {
    assert.equal(retryableGroupReadFailure(groupReadError("POSTS", new Error(message)).message), true);
  }
});

test("interrupted checkpoint publication does not poison retry from previous cursor", t => {
  const config = configFor("partial-write");
  const previous = encodePostContinuation(config, state());
  const write = fs.writeFileSync;
  let injected = false;
  const mock = t.mock.method(fs, "writeFileSync", (file, data, ...args) => {
    if (typeof file === "number" && !injected) {
      injected = true;
      write(file, data.subarray(0, 5), ...args);
      throw Object.assign(new Error("synthetic disk failure"), { code: "EIO" });
    }
    return write(file, data, ...args);
  });
  assert.throws(() => encodePostContinuation(config, state(12000)), /GROUP_CHECKPOINT_STORAGE_BLOCKED/);
  mock.mock.restore();
  assert.deepEqual(decodePostContinuation(config, previous, "query").state, state());
  const next = encodePostContinuation(config, state(12000));
  assert.deepEqual(decodePostContinuation(config, next, "query").state, state(12000));
});

test("unknown hydration failure stops without consuming the unresolved post", async () => {
  const result = await listGroupPosts(configFor("unknown-hydration"), group, 20, options,
    faultClient("hydrate", { left: 1, error: new Error("PRIVATE_RESPONSE") }));
  assert.equal(result.posts.length, 0);
  assert.equal(result.page.scannedPosts, 0);
  assert.equal(result.page.failureDetails.code, "GROUP_POSTS_FAILED");
  assert.equal(result.page.continuationAvailable, false);
  assert.doesNotMatch(JSON.stringify(result), /PRIVATE_RESPONSE/);
});

test("slow local checkpoint load does not spend the Graph traversal budget", async t => {
  const config = configFor("slow-load");
  const client = faultClient("detail", { left: 0, error: new Error("unused") });
  const first = await listGroupPosts(config, group, 1, options, client);
  const s = decodeGroupCursor(config, "posts", first.page.nextCursor);
  s.seen.push(...state().seen);
  const cursor = encodePostContinuation(config, s);
  let clock = Date.now();
  const read = fs.readSync;
  t.mock.method(fs, "readSync", (...args) => { clock += 30000; return read(...args); });
  const result = await listGroupPosts(config, group, 20, { ...options, cursor, now: () => clock }, client);
  assert.equal(result.page.complete, true);
  assert.deepEqual(result.posts.map(p => p.id), ["p2"]);
});

test("signed v3 migration and local retry exhaustion retain exact history", async () => {
  const config = configFor("migration");
  const failures = { left: 0, error: new Error("Graph GET failed (503 Unavailable): private") };
  const client = faultClient("detail", failures);
  const first = await listGroupPosts(config, group, 1, options, client);
  const s = decodeGroupCursor(config, "posts", first.page.nextCursor);
  s.seen.push(...state(9900).seen);
  let cursor = encodeGroupCursor(config, "posts", s);
  failures.left = 10;
  for (let i = 1; i <= 3; i++) {
    const result = await listGroupPosts(config, group, 20, { ...options, cursor }, client);
    assert.equal(result.page.failureDetails.consecutiveAttempts, i);
    assert.equal(result.page.history.seenEntries, s.seen.length);
    if (i < 3) assert.equal(envelope(result.page.nextCursor).version, 4);
    else assert.equal(result.page.continuationAvailable, false);
    cursor = result.page.nextCursor;
  }
});

test("checkpoint per-set capacity accepts the boundary and rejects overflow", () => {
  const config = configFor("capacity");
  const s = state(500000);
  const cursor = encodePostContinuation(config, s);
  assert.equal(decodePostContinuation(config, cursor, "query").state.seen.length, 500000);
  s.seen.push(hash("overflow"));
  assert.throws(() => encodePostContinuation(config, s), /GROUP_CHECKPOINT_STORAGE_BLOCKED/);
});
