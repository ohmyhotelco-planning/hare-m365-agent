import type { AppConfig } from "./config.js";
import type { PendingDownloadApproval } from "./download-approval.js";
import { graphGet } from "./graph.js";
import { downloadDriveItem } from "./sharepoint.js";

export type TeamsAttachmentSummary = {
  id: string;
  name?: string;
  contentType?: string;
  attachmentType: "reference" | "unknown";
  downloadable: boolean;
};

export type TeamsAttachmentListResult = {
  list: {
    chatId: string;
    messageId: string;
    returnedCount: number;
  };
  attachments: TeamsAttachmentSummary[];
};

export type TeamsAttachmentDownloadResult =
  | {
      stage: "DOWNLOADED";
      outputPath: string;
      attachment: TeamsAttachmentSummary;
      file: { name?: string; size?: number; siteName?: string; libraryName?: string };
    }
  | PendingDownloadApproval;

type GraphTeamsAttachment = {
  id: string;
  name?: string;
  contentType?: string;
  contentUrl?: string;
};

type GraphTeamsMessage = {
  attachments?: GraphTeamsAttachment[];
};

type GraphSite = {
  id: string;
  displayName?: string;
};

type GraphDrive = {
  id: string;
  name?: string;
  webUrl?: string;
};

type GraphDriveItem = {
  id: string;
  name?: string;
  size?: number;
  file?: unknown;
  folder?: unknown;
  parentReference?: { driveId?: string };
};

export type TeamsAttachmentGraphClient = {
  get<T>(pathOrUrl: string): Promise<T>;
  downloadDriveItem(
    driveId: string,
    itemId: string,
    filename?: string,
    approvalToken?: string
  ): Promise<string | PendingDownloadApproval>;
};

const allowedSharePointHosts = new Set([
  "ohmylab.sharepoint.com",
  "ohmylab-my.sharepoint.com"
]);

export async function listTeamsMessageAttachments(
  config: AppConfig,
  chatId: string,
  messageId: string,
  client: TeamsAttachmentGraphClient = defaultTeamsAttachmentClient(config)
): Promise<TeamsAttachmentListResult> {
  const normalizedChatId = requireIdentifier(chatId, "chat-id");
  const normalizedMessageId = requireIdentifier(messageId, "message-id");
  const message = await client.get<GraphTeamsMessage>(
    buildTeamsMessagePath(normalizedChatId, normalizedMessageId)
  );
  const attachments = (message.attachments ?? []).map(toAttachmentSummary);

  return {
    list: {
      chatId: normalizedChatId,
      messageId: normalizedMessageId,
      returnedCount: attachments.length
    },
    attachments
  };
}

export async function downloadTeamsMessageAttachment(
  config: AppConfig,
  chatId: string,
  messageId: string,
  attachmentId: string,
  filename?: string,
  approvalToken?: string,
  client: TeamsAttachmentGraphClient = defaultTeamsAttachmentClient(config)
): Promise<TeamsAttachmentDownloadResult> {
  if (!config.policy.allowDownloads) throw new Error("Downloads are disabled by policy.");

  const normalizedChatId = requireIdentifier(chatId, "chat-id");
  const normalizedMessageId = requireIdentifier(messageId, "message-id");
  const normalizedAttachmentId = requireIdentifier(attachmentId, "attachment-id");
  const message = await client.get<GraphTeamsMessage>(
    buildTeamsMessagePath(normalizedChatId, normalizedMessageId)
  );
  const attachment = (message.attachments ?? []).find(
    (candidate) => candidate.id === normalizedAttachmentId
  );
  if (!attachment) throw new Error("The selected Teams attachment was not found.");

  const summary = toAttachmentSummary(attachment);
  if (!summary.downloadable || !attachment.contentUrl) {
    throw new Error(
      "The selected Teams attachment is not a downloadable SharePoint reference file."
    );
  }

  const resolved = await resolveReferenceAttachment(client, attachment.contentUrl);
  if (!resolved.item.file || resolved.item.folder) {
    throw new Error("The selected Teams attachment does not resolve to a file.");
  }
  const driveId = resolved.item.parentReference?.driveId ?? resolved.drive.id;
  const result = await client.downloadDriveItem(
    driveId,
    resolved.item.id,
    filename?.trim() || attachment.name || resolved.item.name,
    approvalToken
  );
  if (typeof result !== "string") return result;

  return {
    stage: "DOWNLOADED",
    outputPath: result,
    attachment: summary,
    file: {
      name: resolved.item.name,
      size: resolved.item.size,
      siteName: resolved.site.displayName,
      libraryName: resolved.drive.name
    }
  };
}

export function buildTeamsMessagePath(chatId: string, messageId: string): string {
  return `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`;
}

function toAttachmentSummary(attachment: GraphTeamsAttachment): TeamsAttachmentSummary {
  const attachmentType = attachment.contentType === "reference" ? "reference" : "unknown";
  return {
    id: attachment.id,
    name: attachment.name,
    contentType: attachment.contentType,
    attachmentType,
    downloadable:
      attachmentType === "reference" &&
      attachment.contentUrl !== undefined &&
      isAllowedReferenceUrl(attachment.contentUrl)
  };
}

async function resolveReferenceAttachment(
  client: TeamsAttachmentGraphClient,
  contentUrl: string
): Promise<{ site: GraphSite; drive: GraphDrive; item: GraphDriveItem }> {
  const url = parseAllowedReferenceUrl(contentUrl);
  const segments = url.pathname.split("/").filter(Boolean).map(decodePathSegment);
  if (segments.length < 4 || (segments[0] !== "sites" && segments[0] !== "personal")) {
    throw new Error("The Teams attachment URL is not a supported SharePoint file path.");
  }

  const sitePath = `/${segments.slice(0, 2).join("/")}`;
  const site = await client.get<GraphSite>(
    `/sites/${url.hostname}:${encodeServerRelativePath(sitePath)}?$select=id,displayName`
  );
  const drives = await client.get<{ value?: GraphDrive[] }>(
    `/sites/${encodeURIComponent(site.id)}/drives?$select=id,name,webUrl`
  );
  const drive = selectDriveForUrl(url, drives.value ?? []);
  const drivePath = new URL(drive.webUrl as string).pathname.replace(/\/$/, "");
  const relativeEncodedPath = url.pathname.slice(drivePath.length).replace(/^\//, "");
  if (!relativeEncodedPath) throw new Error("The Teams attachment URL does not identify a file.");
  const relativePath = relativeEncodedPath
    .split("/")
    .map(decodePathSegment)
    .map(encodeURIComponent)
    .join("/");
  const item = await client.get<GraphDriveItem>(
    `/drives/${encodeURIComponent(drive.id)}/root:/${relativePath}:?$select=id,name,size,file,folder,parentReference`
  );
  return { site, drive, item };
}

function parseAllowedReferenceUrl(value: string): URL {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new Error("The Teams attachment URL is invalid.");
  }
  if (
    url.protocol !== "https:" ||
    url.port ||
    url.username ||
    url.password ||
    !allowedSharePointHosts.has(url.hostname.toLowerCase())
  ) {
    throw new Error("The Teams attachment URL host is not allowed.");
  }
  return url;
}

function isAllowedReferenceUrl(value: string): boolean {
  try {
    const url = parseAllowedReferenceUrl(value);
    const segments = url.pathname.split("/").filter(Boolean).map(decodePathSegment);
    return segments.length >= 4 && (segments[0] === "sites" || segments[0] === "personal");
  } catch {
    return false;
  }
}

function selectDriveForUrl(url: URL, drives: GraphDrive[]): GraphDrive {
  let selected: GraphDrive | undefined;
  let selectedPath = "";
  for (const drive of drives) {
    if (!drive.webUrl) continue;
    try {
      const driveUrl = new URL(drive.webUrl);
      if (driveUrl.hostname.toLowerCase() !== url.hostname.toLowerCase()) continue;
      const drivePath = driveUrl.pathname.replace(/\/$/, "");
      if (
        url.pathname.toLowerCase().startsWith(`${drivePath}/`.toLowerCase()) &&
        drivePath.length > selectedPath.length
      ) {
        selected = drive;
        selectedPath = drivePath;
      }
    } catch {
      continue;
    }
  }
  if (!selected) {
    throw new Error("The SharePoint document library for this Teams attachment was not found.");
  }
  return selected;
}

function encodeServerRelativePath(value: string): string {
  return value.split("/").map((segment) => encodeURIComponent(segment)).join("/");
}

function decodePathSegment(value: string): string {
  try {
    const decoded = decodeURIComponent(value);
    if (decoded === "." || decoded === ".." || decoded.includes("\\")) {
      throw new Error("unsafe path segment");
    }
    return decoded;
  } catch {
    throw new Error("The Teams attachment URL contains an invalid encoded path.");
  }
}

function requireIdentifier(value: string, optionName: string): string {
  const normalized = value.trim();
  if (!normalized) throw new Error(`${optionName} must not be empty.`);
  return normalized;
}

function defaultTeamsAttachmentClient(config: AppConfig): TeamsAttachmentGraphClient {
  return {
    get: <T>(pathOrUrl: string) => graphGet<T>(config, pathOrUrl),
    downloadDriveItem: (driveId, itemId, filename, approvalToken) =>
      downloadDriveItem(config, driveId, itemId, filename, approvalToken)
  };
}
