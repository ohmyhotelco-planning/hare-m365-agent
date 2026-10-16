import { createHash } from "node:crypto";
import fs from "node:fs";
import type { AppConfig } from "./config.js";
import { authorizeDownload, type PendingDownloadApproval } from "./download-approval.js";
import { sanitizeFilename, saveDownloadResponse, type DownloadResponse } from "./downloads.js";
import { graphDownloadResponse, graphGet, type GraphRequestOptions } from "./graph.js";
import { groupIdentifier, groupPageUrl, groupReadError, type GroupStage } from "./groups-common.js";
import { decodeGroupCursor, encodeGroupCursor } from "./groups-cursor.js";

export type GroupPostTarget = { groupId: string; threadId: string; postId: string };
export type GroupAttachmentClient = {
  get<T>(url: string, options?: GraphRequestOptions): Promise<T>;
  download(url: string, options?: GraphRequestOptions): Promise<DownloadResponse>;
};
export type GroupAttachmentSummary = {
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  lastModifiedDateTime?: string;
  attachmentType: "file" | "item" | "reference" | "unknown";
  downloadable: boolean;
};
type StopReason = "COMPLETE" | "LIMIT_REACHED" | "PAGE_LIMIT" | "TIME_LIMIT" | "NO_PROGRESS" | "CURSOR_LIMIT" | "READ_FAILED";
type AttachmentPage = { value?: unknown[]; "@odata.nextLink"?: unknown };
export type GroupAttachmentListResult = {
  stage: "ATTACHMENT_LIST";
  target: GroupPostTarget;
  metadataVerified: true;
  list: {
    returnedCount: number;
    maxResults: number;
    limitReached: boolean;
    partial: boolean;
    stopReason: StopReason;
    pagesRead: number;
    continuationAvailable: boolean;
    nextCursor?: string;
    failure?: { stage: "ATTACHMENT_LIST"; message: string };
  };
  attachments: GroupAttachmentSummary[];
};
type Cursor = {
  version: 1;
  target: GroupPostTarget;
  maxResults: number;
  url: string;
  offset: number;
  pageHash?: string;
  seen: string[];
};
type PendingGroupApproval = PendingDownloadApproval & {
  target: GroupPostTarget;
  preview: PendingDownloadApproval["preview"] & { target: GroupPostTarget; attachmentId: string };
};
export type GroupAttachmentDownloadResult = PendingGroupApproval | {
  stage: "DOWNLOADED";
  outputPath: string;
  target: GroupPostTarget;
  attachment: GroupAttachmentSummary;
  metadataVerified: true;
  downloadedBytes: number;
  structureValidation: "NOT_PERFORMED";
};

const select = "id,name,contentType,size,isInline,lastModifiedDateTime";
const timeBudgetMs = 60_000;
const maxPages = 100;
const maxSeen = 10_000;
const maxCursorLength = 1_000_000;

// Delegated Group-Conversation.Read.All or Group.Read.All; application is unsupported.
// https://learn.microsoft.com/en-us/graph/api/post-list-attachments?view=graph-rest-1.0
export async function listGroupPostAttachments(
  config: AppConfig,
  target: GroupPostTarget,
  limit = 20,
  options: { cursor?: string } = {},
  client: GroupAttachmentClient = defaultClient(config)
): Promise<GroupAttachmentListResult> {
  target = checkedTarget(target);
  if (!Number.isFinite(limit) || limit < 1) throw new Error("limit must be a positive number.");
  if (!Number.isSafeInteger(config.policy.maxSearchResults) || config.policy.maxSearchResults < 1) {
    throw new Error("Search policy maxSearchResults must be a positive integer.");
  }
  const maxResults = Math.min(Math.floor(limit), config.policy.maxSearchResults);
  const collection = `${postPath(target)}/attachments`;
  let state: Cursor = options.cursor === undefined ? {
    version: 1, target, maxResults,
    url: groupPageUrl(`${collection}?$select=${select}&$top=${Math.min(100, maxResults)}`, collection),
    offset: 0, seen: []
  } : decodeCursor(config, options.cursor, target, maxResults, collection);
  const deadline = Date.now() + timeBudgetMs;
  await verifyPost(client, target, deadline);
  const seen = new Set(state.seen);
  const visited = new Set<string>();
  const attachments: GroupAttachmentSummary[] = [];
  let pagesRead = 0;
  let stopReason: StopReason = "COMPLETE";
  let failure: GroupAttachmentListResult["list"]["failure"];
  let hasMore = true;
  while (hasMore) {
    if (Date.now() >= deadline) { stopReason = "TIME_LIMIT"; break; }
    if (pagesRead >= maxPages) { stopReason = "PAGE_LIMIT"; break; }
    if (visited.has(state.url)) { stopReason = "NO_PROGRESS"; break; }
    visited.add(state.url);
    let page: AttachmentPage;
    try {
      page = await remoteRead("ATTACHMENT_LIST", () => client.get<AttachmentPage>(
        state.url, { totalTimeoutMs: Math.max(1, deadline - Date.now()) }
      ));
    } catch (error) {
      if (attachments.length === 0) throw error;
      stopReason = "READ_FAILED";
      failure = { stage: "ATTACHMENT_LIST", message: (error as Error).message };
      break;
    }
    pagesRead += 1;
    if (!page || !Array.isArray(page.value) || page.value.length > maxSeen) {
      throw new Error("GROUP_ATTACHMENT_LIST_INVALID_RESPONSE: Attachment page is missing or exceeds the bounded page size.");
    }
    const values = page.value.map((value) => summary(value, "ATTACHMENT_LIST"));
    const pageHash = hash(JSON.stringify(values.map((value) => value.id)));
    if (state.offset > values.length || (state.pageHash && state.pageHash !== pageHash)) {
      throw new Error("GROUP_ATTACHMENT_LIST_CURSOR_STALE: The replayed page changed. Restart the listing.");
    }
    let next: string | undefined;
    if (page["@odata.nextLink"] !== undefined) {
      next = safePageUrl(page["@odata.nextLink"], collection);
    }
    let index = state.offset;
    for (; index < values.length; index += 1) {
      const value = values[index];
      const key = hash(value.id);
      if (seen.has(key)) continue;
      if (attachments.length >= maxResults) { stopReason = "LIMIT_REACHED"; break; }
      if (seen.size >= maxSeen) { stopReason = "CURSOR_LIMIT"; break; }
      seen.add(key);
      attachments.push(value);
    }
    if (index < values.length) {
      state = { ...state, offset: index, pageHash, seen: [...seen] };
      break;
    }
    hasMore = next !== undefined;
    if (!hasMore) break;
    state = { ...state, url: next!, offset: 0, pageHash: undefined, seen: [...seen] };
    if (attachments.length >= maxResults) { stopReason = "LIMIT_REACHED"; break; }
  }
  const nextCursor = hasMore && stopReason !== "NO_PROGRESS" && stopReason !== "CURSOR_LIMIT"
    ? encodeGroupCursor(config, "attachments", state) : undefined;
  return {
    stage: "ATTACHMENT_LIST", target, metadataVerified: true,
    list: {
      returnedCount: attachments.length, maxResults,
      limitReached: hasMore && attachments.length >= maxResults,
      partial: hasMore, stopReason, pagesRead,
      continuationAvailable: nextCursor !== undefined, nextCursor, failure
    },
    attachments
  };
}

// Get metadata/raw content documents delegated Group.Read.All only (no narrower scope).
// Application access is unsupported; Hare remains delegated-only without auth changes.
export async function downloadGroupPostAttachment(
  config: AppConfig,
  target: GroupPostTarget,
  attachmentId: string,
  filename?: string,
  approvalToken?: string,
  client: GroupAttachmentClient = defaultClient(config)
): Promise<GroupAttachmentDownloadResult> {
  if (!config.policy.allowDownloads) throw new Error("Downloads are disabled by policy.");
  target = checkedTarget(target);
  attachmentId = exactIdentifier(attachmentId, "attachment-id");
  const deadline = Date.now() + timeBudgetMs;
  await verifyPost(client, target, deadline);
  const attachmentPath = `${postPath(target)}/attachments/${encodeURIComponent(attachmentId)}`;
  const metadata = await remoteRead("ATTACHMENT_METADATA", () => client.get<unknown>(
    `${attachmentPath}?$select=${select}`, { totalTimeoutMs: Math.max(1, deadline - Date.now()) }
  ));
  const attachment = summary(metadata, "ATTACHMENT_METADATA");
  if (attachment.id !== attachmentId) {
    throw new Error("GROUP_ATTACHMENT_METADATA_CHAIN_MISMATCH: Attachment ID does not match the requested chain.");
  }
  if (!attachment.downloadable) {
    throw new Error("GROUP_ATTACHMENT_METADATA_TYPE_UNSUPPORTED: Only file and item attachments support raw downloads; reference attachments are unsupported.");
  }
  const outputName = sanitizeFilename(filename?.trim() || attachment.name || `${attachmentId}.bin`);
  let authorization;
  try {
    authorization = authorizeDownload(config, {
      source: "group",
      resourceKey: JSON.stringify([target.groupId, target.threadId, target.postId, attachmentId]),
      name: attachment.name || outputName, size: attachment.size ?? Number.NaN, outputName
    }, approvalToken);
  } catch (error) {
    throw safePolicyError(error, "GROUP_ATTACHMENT_APPROVAL_FAILED: Local download authorization failed.");
  }
  if (!authorization.approved) return {
    ...authorization.pending, target,
    preview: { ...authorization.pending.preview, target, attachmentId }
  };

  // Post attachments support delegated access only. Raw item bytes are MIME, not metadata size.
  // https://learn.microsoft.com/en-us/graph/api/attachment-get?view=graph-rest-1.0
  const response = await remoteRead("ATTACHMENT_BYTES", () => client.download(
    `${attachmentPath}/$value`, { totalTimeoutMs: timeBudgetMs }
  ));
  let outputPath: string;
  try {
    outputPath = await saveDownloadResponse(config, response, outputName, authorization.maxBytes);
  } catch (error) {
    throw safePolicyError(error, "GROUP_ATTACHMENT_SAVE_FAILED: Attachment bytes could not be saved locally.");
  }
  let downloadedBytes: number;
  try { downloadedBytes = fs.statSync(outputPath).size; }
  catch { throw new Error("GROUP_ATTACHMENT_SAVE_VERIFICATION_FAILED: Saved byte count could not be verified."); }
  return {
    stage: "DOWNLOADED", outputPath, target, attachment, metadataVerified: true,
    downloadedBytes, structureValidation: "NOT_PERFORMED"
  };
}

function exactIdentifier(value: string, label: string): string {
  const normalized = groupIdentifier(value, label);
  if (normalized !== value || value === "." || value === "..") {
    throw new Error("GROUP_INVALID_IDENTIFIER: Identifiers must be exact opaque IDs without surrounding whitespace or dot segments.");
  }
  return value;
}

function checkedTarget(target: GroupPostTarget): GroupPostTarget {
  return {
    groupId: exactIdentifier(target?.groupId, "group-id"),
    threadId: exactIdentifier(target?.threadId, "thread-id"),
    postId: exactIdentifier(target?.postId, "post-id")
  };
}

function postPath(target: GroupPostTarget): string {
  return `/groups/${encodeURIComponent(target.groupId)}/threads/${encodeURIComponent(target.threadId)}/posts/${encodeURIComponent(target.postId)}`;
}

async function verifyPost(client: GroupAttachmentClient, target: GroupPostTarget, deadline: number): Promise<void> {
  const post = await remoteRead("POSTS", () => client.get<{ id?: unknown; conversationThreadId?: unknown }>(
    `${postPath(target)}?$select=id,conversationThreadId`, { totalTimeoutMs: Math.max(1, deadline - Date.now()) }
  ));
  if (!post || post.id !== target.postId || post.conversationThreadId !== target.threadId) {
    throw new Error("GROUP_POSTS_CHAIN_MISMATCH: Parent post and conversationThreadId must match the requested target.");
  }
}

async function remoteRead<T>(stage: GroupStage, action: () => Promise<T>): Promise<T> {
  try { return await action(); }
  catch (error) {
    // Do not let a remote message masquerade as an already-sanitized GROUP_ error.
    throw groupReadError(stage, new Error(`Remote failure: ${error instanceof Error ? error.message : "unknown"}`));
  }
}

function summary(value: unknown, stage: "ATTACHMENT_LIST" | "ATTACHMENT_METADATA"): GroupAttachmentSummary {
  if (!value || typeof value !== "object" || !("id" in value) || typeof value.id !== "string") {
    throw new Error(`GROUP_${stage}_INVALID_RESPONSE: Attachment metadata requires an ID.`);
  }
  const data = value as Record<string, unknown>;
  let id: string;
  try { id = exactIdentifier(value.id, "attachment-id"); }
  catch { throw new Error(`GROUP_${stage}_INVALID_RESPONSE: Attachment metadata contains an invalid ID.`); }
  const attachmentType = data["@odata.type"] === "#microsoft.graph.fileAttachment" ? "file"
    : data["@odata.type"] === "#microsoft.graph.itemAttachment" ? "item"
    : data["@odata.type"] === "#microsoft.graph.referenceAttachment" ? "reference" : "unknown";
  return {
    id,
    name: typeof data.name === "string" ? data.name : undefined,
    contentType: typeof data.contentType === "string" ? data.contentType : undefined,
    size: typeof data.size === "number" && Number.isSafeInteger(data.size) && data.size >= 0 ? data.size : undefined,
    isInline: typeof data.isInline === "boolean" ? data.isInline : undefined,
    lastModifiedDateTime: typeof data.lastModifiedDateTime === "string" ? data.lastModifiedDateTime : undefined,
    attachmentType, downloadable: attachmentType === "file" || attachmentType === "item"
  };
}

function safePageUrl(value: unknown, collection: string): string {
  try {
    if (typeof value !== "string" || !value || value.length > 16_384) throw new Error();
    const url = groupPageUrl(value, collection);
    const params = new URL(url).searchParams;
    const allowed = new Set(["$select", "$top", "$skip", "$skiptoken"]);
    for (const key of params.keys()) {
      if (!allowed.has(key) || params.getAll(key).length !== 1) throw new Error();
    }
    if (params.has("$select") && params.get("$select") !== select) throw new Error();
    return url;
  } catch {
    throw new Error("GROUP_ATTACHMENT_LIST_INVALID_CONTINUATION: Continuation must preserve the exact attachment collection and metadata criteria.");
  }
}

function decodeCursor(config: AppConfig, cursor: string, target: GroupPostTarget, maxResults: number, collection: string): Cursor {
  try {
    if (!cursor || cursor.length > maxCursorLength) throw new Error();
    const value = decodeGroupCursor<Cursor>(config, "attachments", cursor);
    if (value.version !== 1 || JSON.stringify(value.target) !== JSON.stringify(target) || value.maxResults !== maxResults ||
        !Number.isSafeInteger(value.offset) || value.offset < 0 || value.offset > maxSeen ||
        !Array.isArray(value.seen) || value.seen.length > maxSeen ||
        !value.seen.every((id) => typeof id === "string" && /^[a-f0-9]{64}$/.test(id)) ||
        (value.offset > 0 && !value.pageHash) ||
        (value.pageHash !== undefined && !/^[a-f0-9]{64}$/.test(value.pageHash))) throw new Error();
    return { ...value, url: safePageUrl(value.url, collection) };
  } catch {
    throw new Error("GROUP_ATTACHMENT_LIST_INVALID_CURSOR: Cursor must match the exact target and result limit; restart the listing.");
  }
}

function hash(value: string): string { return createHash("sha256").update(value).digest("hex"); }

function safePolicyError(error: unknown, fallback: string): Error {
  const message = error instanceof Error ? error.message : "";
  if (/^(?:Download exceeds (?:approved )?policy limit of \d+ bytes\.|Download size metadata is required for policy approval\.|Download policy is invalid: maxApprovedDownloadBytes must be at least maxDownloadBytes\.|Download approval plan is incomplete\.|DOWNLOAD_APPROVAL_REQUIRED: The approval token is missing, expired, or already used\. Show a new preview and obtain approval again\.|DOWNLOAD_APPROVAL_REQUIRED: The approved file does not match this download\. Show a new preview and obtain approval again\.)$/.test(message)) {
    return new Error(message);
  }
  return new Error(fallback);
}

function defaultClient(config: AppConfig): GroupAttachmentClient {
  return {
    get: <T>(url: string, options?: GraphRequestOptions) => graphGet<T>(config, url, options),
    download: (url: string, options?: GraphRequestOptions) => graphDownloadResponse(config, url, options)
  };
}
