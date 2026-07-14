import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";
import { graphDelete, graphGet, graphPatch, graphPost } from "./graph.js";
import { fetchWithProxy } from "./proxy.js";

export type DraftKind = "new" | "reply" | "replyAll" | "forward";
export type DraftContentType = "text" | "html";

export type DraftInput = {
  kind: DraftKind;
  sourceMessageId?: string;
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body: string;
  contentType: DraftContentType;
  attachmentPaths?: string[];
};

export type DraftPreview = {
  kind: DraftKind;
  sourceMessageId?: string;
  subject: string;
  to: string[];
  cc: string[];
  bcc: string[];
  body: string;
  contentType: DraftContentType;
  originalThreadWillBeIncluded: boolean;
  attachments: Array<{
    name: string;
    size: number;
    uploadMode: "direct" | "uploadSession";
  }>;
  warning: string;
};

export type PreparedDraft = {
  preview: DraftPreview;
  approvalToken: string;
  plan: DraftPlan;
};

export type CreatedDraftResult = {
  ok: true;
  stage: "DRAFT_CREATED";
  draft: {
    id: string;
    kind: DraftKind;
    subject?: string;
    to: string[];
    cc: string[];
    bcc: string[];
    isDraft: boolean;
    hasAttachments: boolean;
    webLink?: string;
  };
  sendAvailable: false;
  instruction: string;
};

type EmailAddress = { name?: string; address?: string };
type Recipient = { emailAddress?: EmailAddress };

type GraphMessage = {
  id: string;
  subject?: string;
  body?: { contentType?: string; content?: string };
  from?: Recipient;
  replyTo?: Recipient[];
  toRecipients?: Recipient[];
  ccRecipients?: Recipient[];
  bccRecipients?: Recipient[];
  hasAttachments?: boolean;
  isDraft?: boolean;
  webLink?: string;
};

type GraphUser = {
  mail?: string;
  userPrincipalName?: string;
};

type PreparedAttachment = {
  filePath: string;
  name: string;
  size: number;
  sha256: string;
  uploadMode: "direct" | "uploadSession";
};

type DraftPlan = {
  kind: DraftKind;
  sourceMessageId?: string;
  subject: string;
  to: string[];
  cc: string[];
  bcc: string[];
  body: string;
  contentType: DraftContentType;
  attachments: PreparedAttachment[];
};

export type DraftGraphClient = {
  get<T>(pathOrUrl: string): Promise<T>;
  post<T>(pathOrUrl: string, body?: unknown): Promise<T>;
  patch<T>(pathOrUrl: string, body: unknown): Promise<T>;
  delete(pathOrUrl: string): Promise<void>;
  uploadChunk(uploadUrl: string, bytes: Uint8Array, start: number, total: number): Promise<void>;
};

const directAttachmentLimitBytes = 3 * 1024 * 1024;
const largeUploadChunkBytes = 12 * 320 * 1024;
const maximumRecipientCount = 500;

export async function prepareDraft(
  config: AppConfig,
  input: DraftInput,
  client: DraftGraphClient = defaultDraftClient(config)
): Promise<PreparedDraft> {
  requireDraftActionsEnabled(config);
  const body = input.body.trim();
  if (!body) throw new Error("Draft body must not be empty.");
  if (input.contentType !== "text" && input.contentType !== "html") {
    throw new Error("content type must be text or html.");
  }

  const suppliedTo = normalizeAddresses(input.to ?? []);
  const suppliedCc = normalizeAddresses(input.cc ?? []);
  const suppliedBcc = normalizeAddresses(input.bcc ?? []);
  let subject = input.subject?.trim() ?? "";
  let to = suppliedTo;
  let cc = suppliedCc;
  const bcc = suppliedBcc;

  if (input.kind === "new") {
    if (!subject) throw new Error("subject must not be empty for a new draft.");
    if (to.length === 0) throw new Error("at least one To recipient is required.");
  } else {
    if (!input.sourceMessageId?.trim()) {
      throw new Error("source message ID is required for reply and forward drafts.");
    }
    const source = await client.get<GraphMessage>(
      `/me/messages/${encodeURIComponent(input.sourceMessageId)}?$select=id,subject,from,replyTo,toRecipients,ccRecipients,hasAttachments`
    );

    if (input.kind === "forward") {
      if (to.length === 0) throw new Error("at least one To recipient is required for a forward draft.");
      subject = addSubjectPrefix(source.subject, "FW:");
    } else {
      const replyTargets = addressesFromRecipients(
        source.replyTo && source.replyTo.length > 0 ? source.replyTo : source.from ? [source.from] : []
      );
      if (replyTargets.length === 0) throw new Error("The source message has no reply recipient.");
      subject = addSubjectPrefix(source.subject, "RE:");

      if (input.kind === "replyAll") {
        const me = await client.get<GraphUser>("/me?$select=mail,userPrincipalName");
        const ownAddresses = normalizeAddresses([me.mail ?? "", me.userPrincipalName ?? ""]);
        const excluded = new Set(ownAddresses.map(normalizeAddressKey));
        to = uniqueAddresses([
          ...replyTargets,
          ...addressesFromRecipients(source.toRecipients),
          ...suppliedTo
        ]).filter((address) => !excluded.has(normalizeAddressKey(address)));
        const toKeys = new Set(to.map(normalizeAddressKey));
        cc = uniqueAddresses([
          ...addressesFromRecipients(source.ccRecipients),
          ...suppliedCc
        ]).filter(
          (address) => !excluded.has(normalizeAddressKey(address)) && !toKeys.has(normalizeAddressKey(address))
        );
      } else {
        to = uniqueAddresses([...replyTargets, ...suppliedTo]);
      }
    }
  }

  ensureRecipientLimit(to, cc, bcc);
  const attachments = await inspectAttachments(config, input.attachmentPaths ?? []);
  const plan: DraftPlan = {
    kind: input.kind,
    sourceMessageId: input.sourceMessageId?.trim() || undefined,
    subject,
    to,
    cc,
    bcc,
    body,
    contentType: input.contentType,
    attachments
  };

  const approvalToken = createHash("sha256")
    .update(JSON.stringify(plan))
    .digest("hex");

  return {
    approvalToken,
    plan,
    preview: {
      kind: plan.kind,
      sourceMessageId: plan.sourceMessageId,
      subject: plan.subject,
      to: plan.to,
      cc: plan.cc,
      bcc: plan.bcc,
      body: plan.body,
      contentType: plan.contentType,
      originalThreadWillBeIncluded: plan.kind !== "new",
      attachments: plan.attachments.map(({ name, size, uploadMode }) => ({ name, size, uploadMode })),
      warning: "This action creates an Outlook draft only. Hare cannot send mail."
    }
  };
}

export async function createApprovedDraft(
  config: AppConfig,
  input: DraftInput,
  approvalToken: string,
  client: DraftGraphClient = defaultDraftClient(config)
): Promise<CreatedDraftResult> {
  const prepared = await prepareDraft(config, input, client);
  if (!approvalToken || approvalToken !== prepared.approvalToken) {
    throw new Error(
      "DRAFT_APPROVAL_REQUIRED: The approved content does not match this draft. Show a new preview and obtain user approval again."
    );
  }

  let createdDraft: GraphMessage | undefined;
  try {
    createdDraft = await createBaseDraft(prepared.plan, client);
    if (!createdDraft.id) throw new Error("Microsoft Graph returned a draft without an ID.");

    for (const attachment of prepared.plan.attachments) {
      await addAttachment(client, createdDraft.id, attachment);
    }

    return {
      ok: true,
      stage: "DRAFT_CREATED",
      draft: {
        id: createdDraft.id,
        kind: prepared.plan.kind,
        subject: createdDraft.subject ?? prepared.plan.subject,
        to: prepared.plan.to,
        cc: prepared.plan.cc,
        bcc: prepared.plan.bcc,
        isDraft: createdDraft.isDraft ?? true,
        hasAttachments: prepared.plan.attachments.length > 0 || Boolean(createdDraft.hasAttachments),
        webLink: createdDraft.webLink
      },
      sendAvailable: false,
      instruction: "Draft created. The user must review it in Outlook. Hare cannot send mail."
    };
  } catch (error) {
    if (!createdDraft?.id) throw error;

    try {
      await client.delete(`/me/messages/${encodeURIComponent(createdDraft.id)}`);
    } catch (cleanupError) {
      throw new Error(
        `DRAFT_CREATION_FAILED_AND_CLEANUP_FAILED: ${errorMessage(error)} Draft ID: ${createdDraft.id}. Cleanup error: ${errorMessage(cleanupError)}`
      );
    }
    throw new Error(`DRAFT_CREATION_FAILED_CLEANED_UP: ${errorMessage(error)}`);
  }
}

export async function deleteDraft(
  config: AppConfig,
  draftId: string,
  client: DraftGraphClient = defaultDraftClient(config)
): Promise<void> {
  requireDraftActionsEnabled(config);
  await client.delete(`/me/messages/${encodeURIComponent(draftId)}`);
}

async function createBaseDraft(plan: DraftPlan, client: DraftGraphClient): Promise<GraphMessage> {
  if (plan.kind === "new") {
    return client.post<GraphMessage>("/me/messages", {
      subject: plan.subject,
      body: {
        contentType: graphContentType(plan.contentType),
        content: plan.body
      },
      toRecipients: toGraphRecipients(plan.to),
      ccRecipients: toGraphRecipients(plan.cc),
      bccRecipients: toGraphRecipients(plan.bcc)
    });
  }

  const action = plan.kind === "replyAll"
    ? "createReplyAll"
    : plan.kind === "reply"
      ? "createReply"
      : "createForward";
  const draft = await client.post<GraphMessage>(
    `/me/messages/${encodeURIComponent(plan.sourceMessageId ?? "")}/${action}`
  );
  if (!draft.id) return draft;

  return client.patch<GraphMessage>(`/me/messages/${encodeURIComponent(draft.id)}`, {
    body: combineThreadBody(plan.body, plan.contentType, draft.body),
    toRecipients: toGraphRecipients(plan.to),
    ccRecipients: toGraphRecipients(plan.cc),
    bccRecipients: toGraphRecipients(plan.bcc)
  });
}

async function addAttachment(
  client: DraftGraphClient,
  draftId: string,
  attachment: PreparedAttachment
): Promise<void> {
  const encodedDraftId = encodeURIComponent(draftId);
  if (attachment.uploadMode === "direct") {
    await client.post(`/me/messages/${encodedDraftId}/attachments`, {
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: attachment.name,
      contentType: contentTypeForFile(attachment.name),
      contentBytes: fs.readFileSync(attachment.filePath).toString("base64")
    });
    return;
  }

  const session = await client.post<{ uploadUrl?: string }>(
    `/me/messages/${encodedDraftId}/attachments/createUploadSession`,
    {
      AttachmentItem: {
        attachmentType: "file",
        name: attachment.name,
        size: attachment.size,
        isInline: false
      }
    }
  );
  if (!session.uploadUrl) throw new Error(`No upload URL was returned for ${attachment.name}.`);

  const handle = fs.openSync(attachment.filePath, "r");
  try {
    let offset = 0;
    while (offset < attachment.size) {
      const length = Math.min(largeUploadChunkBytes, attachment.size - offset);
      const buffer = Buffer.allocUnsafe(length);
      const bytesRead = fs.readSync(handle, buffer, 0, length, offset);
      if (bytesRead <= 0) throw new Error(`Could not read attachment ${attachment.name}.`);
      await client.uploadChunk(session.uploadUrl, buffer.subarray(0, bytesRead), offset, attachment.size);
      offset += bytesRead;
    }
  } finally {
    fs.closeSync(handle);
  }
}

async function inspectAttachments(
  config: AppConfig,
  attachmentPaths: string[]
): Promise<PreparedAttachment[]> {
  const files: PreparedAttachment[] = [];
  let totalBytes = 0;

  for (const suppliedPath of attachmentPaths) {
    const filePath = path.resolve(suppliedPath);
    let stat: fs.Stats;
    try {
      stat = fs.statSync(filePath);
    } catch {
      throw new Error(`Attachment file does not exist: ${suppliedPath}`);
    }
    if (!stat.isFile()) throw new Error(`Attachment path is not a file: ${suppliedPath}`);
    if (stat.size > config.policy.maxDraftAttachmentBytes) {
      throw new Error(`Attachment exceeds the ${config.policy.maxDraftAttachmentBytes}-byte limit: ${suppliedPath}`);
    }
    totalBytes += stat.size;
    if (totalBytes > config.policy.maxDraftTotalAttachmentBytes) {
      throw new Error(
        `Attachments exceed the ${config.policy.maxDraftTotalAttachmentBytes}-byte total limit.`
      );
    }

    files.push({
      filePath,
      name: path.basename(filePath),
      size: stat.size,
      sha256: await sha256File(filePath),
      uploadMode: stat.size < directAttachmentLimitBytes ? "direct" : "uploadSession"
    });
  }

  return files;
}

function defaultDraftClient(config: AppConfig): DraftGraphClient {
  return {
    get: <T>(pathOrUrl: string) => graphGet<T>(config, pathOrUrl),
    post: <T>(pathOrUrl: string, body?: unknown) => graphPost<T>(config, pathOrUrl, body),
    patch: <T>(pathOrUrl: string, body: unknown) => graphPatch<T>(config, pathOrUrl, body),
    delete: (pathOrUrl: string) => graphDelete(config, pathOrUrl),
    uploadChunk: uploadChunk
  };
}

async function uploadChunk(
  uploadUrl: string,
  bytes: Uint8Array,
  start: number,
  total: number
): Promise<void> {
  const end = start + bytes.byteLength - 1;
  let lastError: unknown;

  for (let attempt = 1; attempt <= 4; attempt += 1) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 60_000);
    try {
      const response = await fetchWithProxy(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": String(bytes.byteLength),
          "Content-Range": `bytes ${start}-${end}/${total}`
        },
        body: Buffer.from(bytes),
        signal: controller.signal
      });
      if (response.ok) {
        await response.body?.cancel();
        return;
      }
      const responseBody = await response.text();
      if (![429, 500, 502, 503, 504].includes(response.status)) {
        throw new Error(`ATTACHMENT_UPLOAD_REJECTED: ${response.status} ${response.statusText}: ${responseBody}`);
      }
      if (attempt === 4) throw new Error(`Attachment upload failed after retries: ${responseBody}`);
      await delay(500 * 2 ** (attempt - 1));
    } catch (error) {
      lastError = error;
      if (/^ATTACHMENT_UPLOAD_REJECTED:/.test(errorMessage(error))) throw error;
      if (attempt === 4) throw error;
      await delay(500 * 2 ** (attempt - 1));
    } finally {
      clearTimeout(timer);
    }
  }

  throw lastError instanceof Error ? lastError : new Error("Attachment upload failed.");
}

function combineThreadBody(
  body: string,
  contentType: DraftContentType,
  originalBody: GraphMessage["body"]
): { contentType: "HTML"; content: string } {
  const userContent = contentType === "html"
    ? body
    : escapeHtml(body).replace(/\r?\n/g, "<br>");
  const originalContent = originalBody?.contentType?.toLowerCase() === "html"
    ? originalBody.content ?? ""
    : `<pre>${escapeHtml(originalBody?.content ?? "")}</pre>`;
  return {
    contentType: "HTML",
    content: `<div>${userContent}</div><br>${originalContent}`
  };
}

function addSubjectPrefix(subject: string | undefined, prefix: "RE:" | "FW:"): string {
  const value = subject?.trim() || "(no subject)";
  if (prefix === "RE:" && /^re:/i.test(value)) return value;
  if (prefix === "FW:" && /^(?:fw|fwd):/i.test(value)) return value;
  return `${prefix} ${value}`;
}

function addressesFromRecipients(recipients: Recipient[] | undefined): string[] {
  return normalizeAddresses(
    (recipients ?? []).map((recipient) => recipient.emailAddress?.address ?? "")
  );
}

function normalizeAddresses(addresses: string[]): string[] {
  const split = addresses.flatMap((value) => value.split(/[;,]/));
  const normalized = split.map((value) => value.trim()).filter(Boolean);
  for (const address of normalized) {
    if (!/^[^\s@]+@[^\s@]+$/.test(address)) {
      throw new Error(`Invalid email address: ${address}`);
    }
  }
  return uniqueAddresses(normalized);
}

function uniqueAddresses(addresses: string[]): string[] {
  const seen = new Set<string>();
  return addresses.filter((address) => {
    const key = normalizeAddressKey(address);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function normalizeAddressKey(address: string): string {
  return address.toLocaleLowerCase();
}

function ensureRecipientLimit(to: string[], cc: string[], bcc: string[]): void {
  if (to.length + cc.length + bcc.length > maximumRecipientCount) {
    throw new Error(`A draft can contain at most ${maximumRecipientCount} recipients.`);
  }
}

function toGraphRecipients(addresses: string[]) {
  return addresses.map((address) => ({ emailAddress: { address } }));
}

function graphContentType(contentType: DraftContentType): "Text" | "HTML" {
  return contentType === "html" ? "HTML" : "Text";
}

function contentTypeForFile(fileName: string): string {
  const extension = path.extname(fileName).toLowerCase();
  return {
    ".csv": "text/csv",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".gif": "image/gif",
    ".html": "text/html",
    ".jpeg": "image/jpeg",
    ".jpg": "image/jpeg",
    ".json": "application/json",
    ".pdf": "application/pdf",
    ".png": "image/png",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".txt": "text/plain",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".zip": "application/zip"
  }[extension] ?? "application/octet-stream";
}

function requireDraftActionsEnabled(config: AppConfig): void {
  if (!config.policy.allowDraftActions) {
    throw new Error("Outlook draft actions are disabled by policy.");
  }
}

async function sha256File(filePath: string): Promise<string> {
  const hash = createHash("sha256");
  const stream = fs.createReadStream(filePath);
  for await (const chunk of stream) hash.update(chunk as Buffer);
  return hash.digest("hex");
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function delay(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}
