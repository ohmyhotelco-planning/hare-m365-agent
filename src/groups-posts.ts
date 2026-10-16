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
  conversationId?: string | null; conversationThreadId?: string;
  body?: { content?: string; contentType?: string };
  from?: Recipient; sender?: Recipient;
  createdDateTime?: string; receivedDateTime?: string; lastModifiedDateTime?: string;
  hasAttachments?: boolean;
};
type Job = {
  kind: "conversations" | "threads" | "posts";
  conversationId?: string; threadId?: string;
  url: string; offset: number; fingerprint?: string;
};
type CursorState = {
  version: 1; criteria: string; jobs: Job[]; seen: string[]; pages: string[];
  missingBodies: number; missingTimestamps: number;
};

export type GroupPostSummary = {
  id: string; groupId: string; conversationId: string; threadId: string;
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

// Depth-first traversal retains only identifiers and page positions, never message bodies.
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
    version: 1 as const, criteria, jobs: [newJob(group.id, "conversations")],
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
  // Cache a page only during this call. On resume, detect changed page IDs before skipping.
  const loaded = new Map<string, GraphPage<Entity>>();
  while (state.jobs.length) {
    if (posts.length >= maxResults) { stopReason = "RESULT_LIMIT"; break; }
    if (now() >= deadline) { stopReason = "TIME_BUDGET"; break; }
    if (requests >= maxRequests || scannedRecords >= maxRecords) { stopReason = "SCAN_BUDGET"; break; }
    if (seen.size >= maxHistory || pages.size >= maxHistory) { stopReason = "HISTORY_LIMIT"; resumable = false; break; }
    const job = state.jobs[state.jobs.length - 1];
    const collection = collectionPath(group.id, job);
    const url = groupPageUrl(job.url, collection);
    let page = loaded.get(url);
    if (!page) {
      try {
        requests++;
        page = await client.get<GraphPage<Entity>>(url, { totalTimeoutMs: Math.max(1, deadline - now()) });
        if (!Array.isArray(page.value)) throw new Error("GROUP_INVALID_RESPONSE: Expected a collection.");
        if (page.value.length > 1000) throw new Error("GROUP_PAGE_TOO_LARGE: The server page exceeds the bounded reader capacity.");
        const fingerprint = hash(JSON.stringify(page.value.map(item => item.id)));
        if (job.fingerprint && job.fingerprint !== fingerprint) {
          throw new Error("GROUP_CURSOR_STALE: The current page changed. Restart the query; no snapshot completeness is claimed.");
        }
        job.fingerprint = fingerprint;
        loaded.set(url, page);
      } catch (error) {
        const safe = groupReadError(stageFor(job), error);
        if (!posts.length && !options.cursor && scannedRecords === 0) throw safe;
        stopReason = "READ_FAILED";
        failure = safe.message;
        resumable = !safe.message.startsWith("GROUP_CURSOR_STALE") && !safe.message.startsWith("GROUP_INVALID");
        break;
      }
    }
    const values = page.value!;
    if (job.offset >= values.length) {
      const next = page["@odata.nextLink"];
      if (next) {
        const nextUrl = groupPageUrl(next, collection);
        if (nextUrl === url || pages.has(hash(nextUrl))) {
          stopReason = "NO_PROGRESS"; resumable = false; break;
        }
        pages.add(hash(url));
        loaded.delete(url);
        job.url = nextUrl;
        job.offset = 0;
        job.fingerprint = undefined;
      } else {
        state.jobs.pop();
        loaded.delete(url);
      }
      continue;
    }
    const item = values[job.offset];
    const id = groupIdentifier(item.id ?? "", `${job.kind} id`);
    const identity = hash(JSON.stringify([job.kind, group.id, job.kind === "posts" ? job.threadId : null, id]));
    if (seen.has(identity)) { job.offset++; scannedRecords++; duplicates++; continue; }
    if (job.kind !== "posts") {
      seen.add(identity); job.offset++; scannedRecords++;
      state.jobs.push(job.kind === "conversations"
        ? newJob(group.id, "threads", id)
        : newJob(group.id, "posts", job.conversationId, id));
      continue;
    }
    let post = item as Post;
    assertPostChain(post, job);
    let bodyFailure: string | undefined;
    if (typeof post.body?.content !== "string" || !validDate(post.createdDateTime)) {
      if (requests >= maxRequests || now() >= deadline) { stopReason = "SCAN_BUDGET"; break; }
      try {
        requests++;
        const detail = await client.get<Post>(`${collection}/${encodeURIComponent(id)}?${postSelect()}`, { totalTimeoutMs: Math.max(1, deadline - now()) });
        if (detail.id !== id) throw new Error("GROUP_ID_CHAIN_MISMATCH: The post response does not match the requested ID.");
        assertPostChain(detail, job);
        post = { ...post, ...detail };
      } catch (error) {
        if (error instanceof Error && error.message.startsWith("GROUP_ID_CHAIN_MISMATCH")) throw error;
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
      id, groupId: group.id, conversationId: job.conversationId!, threadId: job.threadId!,
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
  if (options.cursor && state.jobs.length && position(state) === initialPosition) {
    stopReason = "NO_PROGRESS";
    resumable = false;
  }
  const continuationAvailable = state.jobs.length > 0 && resumable;
  const reasons = [stopReason, state.missingBodies ? "BODY_UNAVAILABLE" : undefined, state.missingTimestamps ? "TIMESTAMP_UNAVAILABLE" : undefined].filter((v): v is string => Boolean(v));
  return {
    group, range: { ...range, dateField: "createdDateTime", filtering: "client-side" }, posts,
    page: {
      returnedCount: posts.length, limit: maxResults, scannedRecords, scannedPosts, requests, duplicatesRemoved: duplicates,
      complete: state.jobs.length === 0 && reasons.length === 0,
      partialResult: reasons.length > 0 || state.jobs.length > 0, partialReasons: reasons,
      ...(failure ? { failure } : {}),
      fullBodyUnavailableCount: posts.filter(post => !post.fullBodyAvailable).length,
      cumulativeFullBodyUnavailableCount: state.missingBodies,
      cumulativeTimestampUnavailableCount: state.missingTimestamps,
      continuationAvailable, nextCursor: continuationAvailable ? encodeGroupCursor(config, "posts", state) : undefined,
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
  return JSON.stringify(state.jobs.map(job => [job.kind, job.url, job.offset]));
}
function boundedInteger(value: number, max: number, name: string): number {
  if (!Number.isInteger(value) || value < 1 || value > max) throw new Error(`GROUP_INVALID_OPTION: ${name} must be an integer from 1 to ${max}.`);
  return value;
}
function decodeCursor(config: AppConfig, cursor: string, criteria: string, groupId: string): CursorState {
  const state = decodeGroupCursor<CursorState>(config, "posts", cursor);
  if (!state || state.version !== 1 || state.criteria !== criteria || !Array.isArray(state.jobs) || state.jobs.length < 1 || state.jobs.length > 3 ||
      !Array.isArray(state.seen) || state.seen.length > maxHistory || !Array.isArray(state.pages) || state.pages.length > maxHistory ||
      ![...state.seen, ...state.pages].every(v => typeof v === "string" && /^[a-f0-9]{64}$/.test(v)) ||
      !Number.isSafeInteger(state.missingBodies) || state.missingBodies < 0 || !Number.isSafeInteger(state.missingTimestamps) || state.missingTimestamps < 0) {
    throw new Error("GROUP_INVALID_CURSOR: Cursor does not match this group/date range or reader version.");
  }
  const kinds = ["conversations", "threads", "posts"];
  state.jobs.forEach((job, index) => {
    if (!job || job.kind !== kinds[index] || !Number.isInteger(job.offset) || job.offset < 0 || job.offset > 1000 ||
        (job.fingerprint !== undefined && !/^[a-f0-9]{64}$/.test(job.fingerprint))) {
      throw new Error("GROUP_INVALID_CURSOR: Invalid traversal position.");
    }
    if (index === 2 && job.conversationId !== state.jobs[1].conversationId) throw new Error("GROUP_INVALID_CURSOR: Conversation chain mismatch.");
    groupPageUrl(job.url, collectionPath(groupId, job));
  });
  return state;
}
function textFromHtml(value: string): string {
  return value.replace(/<br\s*\/?>/gi, "\n").replace(/<\/p>/gi, "\n").replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/&lt;/gi, "<").replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"').replace(/&#39;/gi, "'");
}
