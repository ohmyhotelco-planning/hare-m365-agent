import { createHash } from "node:crypto";
import type { AppConfig } from "./config.js";
import type { GraphPage } from "./graph.js";
import {
  defaultGroupClient, groupIdentifier, groupPageUrl, groupPath, groupReadError,
  requireMicrosoft365Group, type GroupReadClient, type GroupReference, type GroupStage
} from "./groups-common.js";
import { resolveSearchRange } from "./search-range.js";
import { decodeGroupCursor, encodeGroupCursor } from "./groups-cursor.js";

type Entity = { id?: string };
type Recipient = { emailAddress?: { name?: string; address?: string } };
type Post = Entity & {
  conversationId?: string | null; conversationThreadId?: string | null;
  body?: { content?: string; contentType?: string };
  from?: Recipient; sender?: Recipient;
  createdDateTime?: string; receivedDateTime?: string; lastModifiedDateTime?: string;
  hasAttachments?: boolean;
};
type Job = {
  kind: "conversations" | "threads" | "posts";
  conversationId?: string; threadId?: string;
  url: string; offset: number;
  snapshot?: { ids: string[]; nextLink?: string; fingerprint: string };
};
type CursorState = {
  version: 3; criteria: string; jobs: Job[]; seen: string[]; pages: string[];
  missingBodies: number; missingTimestamps: number;
};

export type GroupPostSummary = {
  id: string; groupId: string; conversationId: string; threadId: string;
  traversalThreadId?: string;
  threadVerification: "REQUEST_PATH" | "POST_COLLECTION_REVALIDATED";
  from?: Recipient; sender?: Recipient;
  createdDateTime?: string; receivedDateTime?: string; lastModifiedDateTime?: string;
  hasAttachments: boolean | null;
  body?: string; bodyHtml?: string; bodyContent?: string; bodyContentType?: string;
  fullBodyAvailable: boolean; bodyFailure?: string;
};

export type GroupPostOptions = {
  since?: string; until?: string; cursor?: string;
  timeBudgetMs?: number; maxRequests?: number; maxRecords?: number;
  now?: () => number;
};
const maxHistory = 10_000;

// Pin bounded page inventories, not remote offsets or message bodies, across calls.
export async function listGroupPosts(
  config: AppConfig, group: GroupReference, limit: number,
  options: GroupPostOptions = {}, client: GroupReadClient = defaultGroupClient(config)
) {
  requireMicrosoft365Group(group);
  const requestedLimit = boundedInteger(limit, 100, "limit");
  if (!Number.isSafeInteger(config.policy.maxSearchResults) || config.policy.maxSearchResults < 1) {
    throw new Error("GROUP_INVALID_POLICY: maxSearchResults must be a positive integer.");
  }
  const maxResults = Math.min(requestedLimit, config.policy.maxSearchResults);
  const timeBudgetMs = boundedInteger(options.timeBudgetMs ?? 20_000, 30_000, "time-budget-ms");
  const maxRequests = boundedInteger(options.maxRequests ?? 100, 100, "max-requests");
  if (maxRequests < 2) throw new Error("GROUP_INVALID_OPTION: max-requests must allow at least a page read and post detail read (2).");
  const maxRecords = boundedInteger(options.maxRecords ?? 1000, 1000, "max-records");
  const now = options.now ?? Date.now;
  const deadline = now() + timeBudgetMs;
  const range = resolveSearchRange(options.since, options.until, config.policy.defaultSearchLookbackDays, new Date(now()));
  const criteria = hash(JSON.stringify([group.id, range.startDateTime, range.endDateTimeExclusive, "createdDateTime"]));
  const state = options.cursor ? decodeCursor(config, options.cursor, criteria, group.id) : {
    version: 3 as const, criteria, jobs: [newJob(group.id, "conversations")],
    seen: [], pages: [], missingBodies: 0, missingTimestamps: 0
  };
  const seen = new Set(state.seen);
  const pages = new Set(state.pages);
  const initialPosition = position(state);
  const posts: GroupPostSummary[] = [];
  let requests = 0;
  let scannedRecords = 0;
  let scannedPosts = 0;
  let duplicates = 0;
  let stopReason: string | undefined;
  let failure: string | undefined;
  let resumable = true;
  // Same-call objects avoid detail reads; only their exact IDs enter the signed cursor.
  const loaded = new Map<string, GraphPage<Entity>>();
  const readVerificationPage = async (verificationUrl: string): Promise<GraphPage<Post>> => {
    if (now() >= deadline) throw new Error("GROUP_VERIFICATION_TIME_BUDGET");
    if (requests >= maxRequests || scannedRecords >= maxRecords) throw new Error("GROUP_VERIFICATION_SCAN_BUDGET");
    requests++;
    let verificationPage: GraphPage<Post>;
    try {
      verificationPage = await client.get<GraphPage<Post>>(verificationUrl, { totalTimeoutMs: Math.max(1, deadline - now()) });
    } catch (error) {
      throw groupReadError("POSTS", new Error(`Remote failure: ${error instanceof Error ? error.message : "unknown"}`));
    }
    if (!Array.isArray(verificationPage.value) || verificationPage.value.length > 1000) {
      throw new Error("GROUP_ID_CHAIN_MISMATCH: Invalid verification collection.");
    }
    if (scannedRecords + verificationPage.value.length >= maxRecords) throw new Error("GROUP_VERIFICATION_SCAN_BUDGET");
    scannedRecords += verificationPage.value.length;
    return verificationPage;
  };
  while (state.jobs.length) {
    if (posts.length >= maxResults) { stopReason = "RESULT_LIMIT"; break; }
    if (now() >= deadline) { stopReason = "TIME_BUDGET"; break; }
    if (requests >= maxRequests || scannedRecords >= maxRecords) { stopReason = "SCAN_BUDGET"; break; }
    if (seen.size >= maxHistory || pages.size >= maxHistory) { stopReason = "HISTORY_LIMIT"; resumable = false; break; }
    const job = state.jobs[state.jobs.length - 1];
    const collection = collectionPath(group.id, job);
    const url = groupPageUrl(job.url, collection);
    let page = loaded.get(url);
    if (!job.snapshot) {
      try {
        requests++;
        page = await client.get<GraphPage<Entity>>(url, { totalTimeoutMs: Math.max(1, deadline - now()) });
        if (!page || !Array.isArray(page.value)) throw new Error("GROUP_INVALID_RESPONSE: Expected a collection.");
        if (page.value.length > 1000) throw new Error("GROUP_PAGE_TOO_LARGE: The server page exceeds the bounded reader capacity.");
        const ids = page.value.map(item => exactIdentifier(item?.id));
        const next = page["@odata.nextLink"];
        if (next !== undefined && (typeof next !== "string" || !next)) {
          throw new Error("GROUP_INVALID_CONTINUATION: Expected a nonempty continuation URL.");
        }
        const nextLink = next === undefined ? undefined : groupPageUrl(next, collection);
        job.snapshot = { ids, nextLink, fingerprint: inventoryFingerprint(ids, nextLink) };
        loaded.set(url, page);
      } catch (error) {
        const safe = groupReadError(stageFor(job), error);
        if (!posts.length && !options.cursor && scannedRecords === 0) throw safe;
        stopReason = "READ_FAILED";
        failure = safe.message;
        resumable = !safe.message.startsWith("GROUP_INVALID") && !safe.message.startsWith("GROUP_PAGE_TOO_LARGE");
        break;
      }
    }
    if (now() >= deadline) { stopReason = "TIME_BUDGET"; break; }
    if (job.offset >= job.snapshot.ids.length) {
      const next = job.snapshot.nextLink;
      if (next) {
        const nextUrl = groupPageUrl(next, collection);
        if (nextUrl === url || pages.has(hash(nextUrl))) {
          stopReason = "NO_PROGRESS"; resumable = false; break;
        }
        pages.add(hash(url));
        loaded.delete(url);
        job.url = nextUrl;
        job.offset = 0;
        job.snapshot = undefined;
      } else {
        state.jobs.pop();
        loaded.delete(url);
      }
      continue;
    }
    const id = job.snapshot.ids[job.offset];
    const identity = hash(JSON.stringify([job.kind, group.id, id]));
    if (seen.has(identity)) { job.offset++; scannedRecords++; duplicates++; continue; }
    if (job.kind !== "posts") {
      seen.add(identity); job.offset++; scannedRecords++;
      state.jobs.push(job.kind === "conversations"
        ? newJob(group.id, "threads", id)
        : newJob(group.id, "posts", job.conversationId, id));
      continue;
    }
    let post: Post;
    let threadId = job.threadId!;
    let revalidated = false;
    try {
      if (page) {
        post = page.value![job.offset] as Post;
      } else {
        if (requests >= maxRequests) { stopReason = "SCAN_BUDGET"; break; }
        requests++;
        // Never reread a mutable collection to recover a pinned post's old offset.
        try {
          post = await client.get<Post>(`${threadPostsPath(group.id, job.threadId!)}/${encodeURIComponent(id)}?${postSelect()}`, { totalTimeoutMs: Math.max(1, deadline - now()) });
        } catch (error) {
          throw groupReadError("POSTS", new Error(`Remote failure: ${error instanceof Error ? error.message : "unknown"}`));
        }
      }
      if (!post || post.id !== id) throw new Error("GROUP_ID_CHAIN_MISMATCH: The post response does not match the pinned ID.");
      const verified = await verifyPostThread(post, job, group.id, readVerificationPage);
      post = verified.post;
      threadId = verified.threadId;
      revalidated = threadId !== job.threadId;
    } catch (error) {
      const message = error instanceof Error ? error.message : "";
      if (message === "GROUP_VERIFICATION_TIME_BUDGET" || message === "GROUP_VERIFICATION_SCAN_BUDGET") {
        stopReason = message === "GROUP_VERIFICATION_TIME_BUDGET" ? "TIME_BUDGET" : "SCAN_BUDGET";
      } else {
        stopReason = "READ_FAILED";
        failure = groupReadError("POSTS", error).message;
        resumable = false;
      }
      break;
    }
    let bodyFailure: string | undefined;
    if ((page || revalidated) && (typeof post.body?.content !== "string" || !validDate(post.createdDateTime))) {
      if (requests >= maxRequests || now() >= deadline) { stopReason = "SCAN_BUDGET"; break; }
      let verifyingDetail = false;
      try {
        requests++;
        const detailCollection = revalidated ? threadPostsPath(group.id, threadId) : collection;
        const detail = await client.get<Post>(`${detailCollection}/${encodeURIComponent(id)}?${postSelect()}`, { totalTimeoutMs: Math.max(1, deadline - now()) });
        if (detail.id !== id) throw new Error("GROUP_ID_CHAIN_MISMATCH: The post response does not match the requested ID.");
        if (!revalidated && detail.conversationThreadId !== undefined && detail.conversationThreadId !== threadId) {
          verifyingDetail = true;
          const verified = await verifyPostThread(detail, job, group.id, readVerificationPage);
          post = verified.post;
          threadId = verified.threadId;
          revalidated = true;
        } else {
          assertPostChain(detail, { ...job, threadId });
          if (revalidated && detail.conversationThreadId !== threadId) throw new Error("GROUP_ID_CHAIN_MISMATCH: Verified thread evidence is missing from the post detail.");
          post = { ...post, ...detail };
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : "";
        if (message === "GROUP_VERIFICATION_TIME_BUDGET" || message === "GROUP_VERIFICATION_SCAN_BUDGET") {
          stopReason = message === "GROUP_VERIFICATION_TIME_BUDGET" ? "TIME_BUDGET" : "SCAN_BUDGET"; break;
        }
        if (verifyingDetail || message.startsWith("GROUP_ID_CHAIN_MISMATCH")) {
          stopReason = "READ_FAILED"; failure = groupReadError("POSTS", error).message; resumable = false; break;
        }
        bodyFailure = groupReadError("POSTS", error).message;
      }
    }
    seen.add(identity); job.offset++; scannedRecords++; scannedPosts++;
    if (!validDate(post.createdDateTime)) {
      state.missingTimestamps++;
      continue;
    }
    const created = Date.parse(post.createdDateTime!);
    if (created < Date.parse(range.startDateTime) || created >= Date.parse(range.endDateTimeExclusive)) continue;
    const fullBodyAvailable = typeof post.body?.content === "string";
    if (!fullBodyAvailable) state.missingBodies++;
    const html = post.body?.contentType?.toLowerCase() === "html";
    posts.push({
      id, groupId: group.id, conversationId: job.conversationId!, threadId,
      ...(revalidated ? { traversalThreadId: job.threadId } : {}),
      threadVerification: revalidated ? "POST_COLLECTION_REVALIDATED" : "REQUEST_PATH",
      from: post.from, sender: post.sender, createdDateTime: post.createdDateTime,
      receivedDateTime: post.receivedDateTime, lastModifiedDateTime: post.lastModifiedDateTime,
      hasAttachments: typeof post.hasAttachments === "boolean" ? post.hasAttachments : null,
      body: fullBodyAvailable ? (html ? textFromHtml(post.body!.content!) : post.body!.content) : undefined,
      bodyHtml: html ? post.body?.content : undefined,
      bodyContent: post.body?.content, bodyContentType: post.body?.contentType,
      fullBodyAvailable, ...(bodyFailure ? { bodyFailure } : {})
    });
  }
  state.seen = [...seen];
  state.pages = [...pages];
  if (options.cursor && resumable && state.jobs.length && position(state) === initialPosition) {
    stopReason = "NO_PROGRESS";
    resumable = false;
  }
  let nextCursor: string | undefined;
  if (state.jobs.length > 0 && resumable) {
    try { nextCursor = encodeGroupCursor(config, "posts", state); }
    catch (error) {
      if (!(error instanceof Error) || !error.message.startsWith("GROUP_CURSOR_SIZE_LIMIT:")) throw error;
      stopReason = "CURSOR_SIZE_LIMIT";
      resumable = false;
    }
  }
  const continuationAvailable = state.jobs.length > 0 && resumable;
  const reasons = [stopReason, state.missingBodies ? "BODY_UNAVAILABLE" : undefined, state.missingTimestamps ? "TIMESTAMP_UNAVAILABLE" : undefined].filter((v): v is string => Boolean(v));
  return {
    group, range: { ...range, dateField: "createdDateTime", filtering: "client-side" }, posts,
    page: {
      returnedCount: posts.length, limit: maxResults, scannedRecords, scannedPosts, requests, duplicatesRemoved: duplicates,
      complete: state.jobs.length === 0 && reasons.length === 0,
      snapshotConsistent: false,
      completionScope: "PINNED_PAGE_INVENTORIES_WITH_LIVE_CONTINUATIONS",
      partialResult: reasons.length > 0 || state.jobs.length > 0, partialReasons: reasons,
      ...(failure ? { failure } : {}),
      fullBodyUnavailableCount: posts.filter(post => !post.fullBodyAvailable).length,
      cumulativeFullBodyUnavailableCount: state.missingBodies,
      cumulativeTimestampUnavailableCount: state.missingTimestamps,
      continuationAvailable, nextCursor,
      noProgressDetected: stopReason === "NO_PROGRESS", timeBudgetMs,
      notice: "Traversal order is not newest-first. Conversation preview and lastDeliveredDateTime are not post bodies or creation times. Live pagination is not a point-in-time snapshot."
    }
  };
}

function newJob(groupId: string, kind: Job["kind"], conversationId?: string, threadId?: string): Job {
  const job: Job = { kind, conversationId, threadId, url: "", offset: 0 };
  const select = kind === "posts" ? postSelect() : "$select=id";
  job.url = `${collectionPath(groupId, job)}?${select}`;
  return job;
}
function collectionPath(groupId: string, job: Job): string {
  const base = `${groupPath(groupId)}/conversations`;
  if (job.kind === "conversations") return base;
  const threads = `${base}/${encodeURIComponent(groupIdentifier(job.conversationId ?? "", "conversation-id"))}/threads`;
  return job.kind === "threads" ? threads : `${threads}/${encodeURIComponent(groupIdentifier(job.threadId ?? "", "thread-id"))}/posts`;
}
function postSelect(): string {
  return "$select=id,conversationId,conversationThreadId,body,from,sender,createdDateTime,receivedDateTime,lastModifiedDateTime,hasAttachments";
}
function threadPostsPath(groupId: string, threadId: string): string {
  return `${groupPath(groupId)}/threads/${encodeURIComponent(threadId)}/posts`;
}

// Never infer ID equivalence. Require the same post in the returned thread's own collection.
async function verifyPostThread(
  post: Post, job: Job, groupId: string,
  read: (url: string) => Promise<GraphPage<Post>>
): Promise<{ post: Post; threadId: string }> {
  if (post.conversationId != null && post.conversationId !== job.conversationId) {
    throw new Error("GROUP_ID_CHAIN_MISMATCH: The returned conversation does not match the traversal.");
  }
  if (post.conversationThreadId === undefined || post.conversationThreadId === job.threadId) {
    return { post, threadId: job.threadId! };
  }
  const threadId = post.conversationThreadId;
  if (typeof threadId !== "string" || !threadId || threadId !== groupIdentifier(threadId, "thread-id") || threadId === "." || threadId === "..") {
    throw new Error("GROUP_ID_CHAIN_MISMATCH: A valid returned thread ID is required for verification.");
  }
  const collection = threadPostsPath(groupId, threadId);
  let url = groupPageUrl(`${collection}?${postSelect()}`, collection);
  const visited = new Set<string>();
  for (let pageNumber = 0; pageNumber < 3; pageNumber++) {
    if (visited.has(url)) break;
    visited.add(url);
    const page = await read(url);
    const matches = page.value!.filter(value => value?.id === post.id);
    if (matches.length) {
      const verified = matches[0];
      if (matches.length !== 1 || verified.conversationThreadId !== threadId) break;
      assertPostChain(verified, { ...job, threadId });
      return { post: verified, threadId };
    }
    if (!page["@odata.nextLink"]) break;
    url = groupPageUrl(page["@odata.nextLink"]!, collection);
  }
  throw new Error("GROUP_ID_CHAIN_MISMATCH: The same post could not be verified in the returned thread within the bounded check.");
}
function assertPostChain(post: Post, job: Job): void {
  if ((post.conversationThreadId !== undefined && post.conversationThreadId !== job.threadId) ||
      (post.conversationId != null && post.conversationId !== job.conversationId)) {
    throw new Error("GROUP_ID_CHAIN_MISMATCH: The returned post does not belong to the requested conversation/thread.");
  }
}
function stageFor(job: Job): GroupStage {
  return job.kind === "conversations" ? "CONVERSATIONS" : job.kind === "threads" ? "THREADS" : "POSTS";
}
function validDate(value: string | undefined): boolean { return typeof value === "string" && Number.isFinite(Date.parse(value)); }
function hash(value: string): string { return createHash("sha256").update(value).digest("hex"); }
function position(state: CursorState): string {
  return JSON.stringify(state.jobs.map(job => [job.kind, job.url, job.offset, job.snapshot?.fingerprint]));
}
function exactIdentifier(value: unknown): string {
  if (typeof value !== "string" || value !== groupIdentifier(value) || value === "." || value === "..") {
    throw new Error("GROUP_INVALID_IDENTIFIER: Expected an exact opaque identifier.");
  }
  try { encodeURIComponent(value); }
  catch { throw new Error("GROUP_INVALID_IDENTIFIER: Identifier cannot be encoded without alteration."); }
  return value;
}
function inventoryFingerprint(ids: string[], nextLink?: string): string {
  return hash(JSON.stringify([ids, nextLink ?? null]));
}
function boundedInteger(value: number, max: number, name: string): number {
  if (!Number.isInteger(value) || value < 1 || value > max) throw new Error(`GROUP_INVALID_OPTION: ${name} must be an integer from 1 to ${max}.`);
  return value;
}
function decodeCursor(config: AppConfig, cursor: string, criteria: string, groupId: string): CursorState {
  const state = decodeGroupCursor<CursorState>(config, "posts", cursor);
  if (!state || state.version !== 3 || state.criteria !== criteria ||
      Object.keys(state).some(key => !["version", "criteria", "jobs", "seen", "pages", "missingBodies", "missingTimestamps"].includes(key)) ||
      !Array.isArray(state.jobs) || state.jobs.length < 1 || state.jobs.length > 3 ||
      !Array.isArray(state.seen) || state.seen.length > maxHistory || !Array.isArray(state.pages) || state.pages.length > maxHistory ||
      ![...state.seen, ...state.pages].every(v => typeof v === "string" && /^[a-f0-9]{64}$/.test(v)) ||
      !Number.isSafeInteger(state.missingBodies) || state.missingBodies < 0 || !Number.isSafeInteger(state.missingTimestamps) || state.missingTimestamps < 0) {
    throw new Error("GROUP_INVALID_CURSOR: Cursor does not match this group/date range or reader version. Restart the query.");
  }
  const kinds = ["conversations", "threads", "posts"];
  state.jobs.forEach((job, index) => {
    if (!job || job.kind !== kinds[index] || !Number.isInteger(job.offset) || job.offset < 0 ||
        typeof job.url !== "string" || Object.keys(job).some(key => !["kind", "conversationId", "threadId", "url", "offset", "snapshot"].includes(key))) {
      throw new Error("GROUP_INVALID_CURSOR: Invalid traversal position.");
    }
    try {
      if (index === 0 && (job.conversationId !== undefined || job.threadId !== undefined) ||
          index === 1 && job.threadId !== undefined) throw new Error();
      if (index > 0) exactIdentifier(job.conversationId);
      if (index === 2) exactIdentifier(job.threadId);
      const collection = collectionPath(groupId, job);
      groupPageUrl(job.url, collection);
      const snapshot = job.snapshot;
      if (snapshot !== undefined) {
        if (!snapshot || !Array.isArray(snapshot.ids) || snapshot.ids.length > 1000 || job.offset > snapshot.ids.length ||
            Object.keys(snapshot).some(key => !["ids", "nextLink", "fingerprint"].includes(key))) throw new Error();
        snapshot.ids.forEach(exactIdentifier);
        if (snapshot.nextLink !== undefined && (typeof snapshot.nextLink !== "string" ||
            groupPageUrl(snapshot.nextLink, collection) !== snapshot.nextLink)) throw new Error();
        if (snapshot.fingerprint !== inventoryFingerprint(snapshot.ids, snapshot.nextLink)) throw new Error();
      } else if (job.offset !== 0) throw new Error();
      if (index > 0) {
        const parent = state.jobs[index - 1];
        const parentId = parent.snapshot?.ids[parent.offset - 1];
        if (!parentId || parentId !== (index === 1 ? job.conversationId : job.threadId) ||
            index === 2 && job.conversationId !== parent.conversationId) throw new Error();
      }
    } catch {
      throw new Error("GROUP_INVALID_CURSOR: Invalid pinned inventory or parent/collection binding. Restart the query.");
    }
  });
  return state;
}
function textFromHtml(value: string): string {
  return value.replace(/<br\s*\/?>/gi, "\n").replace(/<\/p>/gi, "\n").replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/&lt;/gi, "<").replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"').replace(/&#39;/gi, "'");
}
