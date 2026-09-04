import type { AppConfig } from "./config.js";
import {
  authorizeDownload,
  type PendingDownloadApproval
} from "./download-approval.js";
import {
  sanitizeFilename,
  saveDownloadResponse,
  type DownloadResponse
} from "./downloads.js";
import { graphDownloadResponse, graphGet, type GraphPage } from "./graph.js";
import {
  mailboxKey,
  mailboxPath,
  mailboxReference,
  selfMailboxTarget,
  type MailboxReference,
  type MailboxTarget
} from "./outlook-mailbox.js";

export type OutlookAttachmentSummary = {
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  lastModifiedDateTime?: string;
  attachmentType: "file" | "item" | "reference" | "unknown";
  downloadable: boolean;
};

export type OutlookAttachmentListResult = {
  list: {
    mailbox: MailboxReference;
    messageId: string;
    returnedCount: number;
    maxResults: number;
    limitReached: boolean;
  };
  attachments: OutlookAttachmentSummary[];
};

type GraphAttachment = {
  "@odata.type"?: string;
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  lastModifiedDateTime?: string;
};

export type OutlookAttachmentGraphClient = {
  get<T>(pathOrUrl: string): Promise<T>;
  download(pathOrUrl: string): Promise<DownloadResponse>;
};

type PendingOutlookAttachmentApproval = PendingDownloadApproval & {
  preview: PendingDownloadApproval["preview"] & {
    mailbox: MailboxReference;
  };
};

const attachmentSelect = [
  "id",
  "name",
  "contentType",
  "size",
  "isInline",
  "lastModifiedDateTime"
].join(",");

export async function listMessageAttachments(
  config: AppConfig,
  messageId: string,
  limit: number,
  client: OutlookAttachmentGraphClient = defaultAttachmentClient(config),
  mailbox: MailboxTarget = selfMailboxTarget()
): Promise<OutlookAttachmentListResult> {
  const normalizedMessageId = requireIdentifier(messageId, "message-id");
  if (!Number.isFinite(limit) || limit < 1) throw new Error("limit must be a positive number.");

  const maxResults = Math.min(Math.floor(limit), config.policy.maxSearchResults);
  let nextUrl: string | undefined = buildAttachmentListPath(
    normalizedMessageId,
    maxResults,
    mailbox
  );
  const attachments: OutlookAttachmentSummary[] = [];
  let limitReached = false;

  while (nextUrl && attachments.length < maxResults) {
    const page: GraphPage<GraphAttachment> = await client.get<GraphPage<GraphAttachment>>(nextUrl);
    const values = page.value ?? [];
    const remaining = maxResults - attachments.length;
    attachments.push(...values.slice(0, remaining).map(toAttachmentSummary));
    nextUrl = page["@odata.nextLink"];
    limitReached = Boolean(nextUrl) || values.length > remaining;
  }

  return {
    list: {
      mailbox: mailboxReference(mailbox),
      messageId: normalizedMessageId,
      returnedCount: attachments.length,
      maxResults,
      limitReached
    },
    attachments
  };
}

export async function downloadMessageAttachment(
  config: AppConfig,
  messageId: string,
  attachmentId: string,
  filename?: string,
  approvalToken?: string,
  client: OutlookAttachmentGraphClient = defaultAttachmentClient(config),
  mailbox: MailboxTarget = selfMailboxTarget()
): Promise<
  | {
      stage: "DOWNLOADED";
      outputPath: string;
      mailbox: MailboxReference;
      attachment: OutlookAttachmentSummary;
    }
  | PendingOutlookAttachmentApproval
> {
  if (!config.policy.allowDownloads) throw new Error("Downloads are disabled by policy.");

  const normalizedMessageId = requireIdentifier(messageId, "message-id");
  const normalizedAttachmentId = requireIdentifier(attachmentId, "attachment-id");
  const metadataPath = buildAttachmentMetadataPath(
    normalizedMessageId,
    normalizedAttachmentId,
    mailbox
  );
  const attachment = toAttachmentSummary(await client.get<GraphAttachment>(metadataPath));

  if (!attachment.downloadable) {
    throw new Error(
      `Attachment type ${attachment.attachmentType} cannot be downloaded as a mailbox file.`
    );
  }
  const outputName = sanitizeFilename(
    filename?.trim() || attachment.name || `${normalizedAttachmentId}.bin`
  );
  const authorization = authorizeDownload(
    config,
    {
      source: "outlook",
      resourceKey: `${mailboxKey(mailbox)}\u0000${normalizedMessageId}\u0000${normalizedAttachmentId}`,
      name: attachment.name || outputName,
      size: attachment.size ?? Number.NaN,
      outputName
    },
    approvalToken
  );
  if (!authorization.approved) {
    return {
      ...authorization.pending,
      preview: {
        ...authorization.pending.preview,
        mailbox: mailboxReference(mailbox)
      }
    };
  }

  const response = await client.download(
    buildAttachmentContentPath(normalizedMessageId, normalizedAttachmentId, mailbox)
  );
  const outputPath = await saveDownloadResponse(
    config,
    response,
    outputName,
    authorization.maxBytes
  );
  return { stage: "DOWNLOADED", outputPath, mailbox: mailboxReference(mailbox), attachment };
}

export function buildAttachmentListPath(
  messageId: string,
  limit: number,
  mailbox: MailboxTarget = selfMailboxTarget()
): string {
  return `${mailboxPath(mailbox, `/messages/${encodeURIComponent(messageId)}/attachments`)}?$select=${attachmentSelect}&$top=${Math.min(100, limit)}`;
}

export function buildAttachmentMetadataPath(
  messageId: string,
  attachmentId: string,
  mailbox: MailboxTarget = selfMailboxTarget()
): string {
  return `${mailboxPath(mailbox, `/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`)}?$select=${attachmentSelect}`;
}

export function buildAttachmentContentPath(
  messageId: string,
  attachmentId: string,
  mailbox: MailboxTarget = selfMailboxTarget()
): string {
  return mailboxPath(
    mailbox,
    `/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}/$value`
  );
}

function toAttachmentSummary(attachment: GraphAttachment): OutlookAttachmentSummary {
  const attachmentType = normalizeAttachmentType(attachment["@odata.type"]);
  return {
    id: attachment.id,
    name: attachment.name,
    contentType: attachment.contentType,
    size: attachment.size,
    isInline: attachment.isInline,
    lastModifiedDateTime: attachment.lastModifiedDateTime,
    attachmentType,
    downloadable: attachmentType === "file" || attachmentType === "item"
  };
}

function normalizeAttachmentType(
  odataType: string | undefined
): OutlookAttachmentSummary["attachmentType"] {
  if (odataType?.endsWith("fileAttachment")) return "file";
  if (odataType?.endsWith("itemAttachment")) return "item";
  if (odataType?.endsWith("referenceAttachment")) return "reference";
  return "unknown";
}

function requireIdentifier(value: string, optionName: string): string {
  const normalized = value.trim();
  if (!normalized) throw new Error(`${optionName} must not be empty.`);
  return normalized;
}

function defaultAttachmentClient(config: AppConfig): OutlookAttachmentGraphClient {
  return {
    get: <T>(pathOrUrl: string) => graphGet<T>(config, pathOrUrl),
    download: (pathOrUrl: string) => graphDownloadResponse(config, pathOrUrl)
  };
}
