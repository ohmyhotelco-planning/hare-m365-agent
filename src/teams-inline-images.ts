import path from "node:path";
import type { AppConfig } from "./config.js";
import { sanitizeFilename, saveDownloadResponse, type DownloadResponse } from "./downloads.js";
import { graphDownloadResponse, graphGet, type GraphPage } from "./graph.js";

export type TeamsInlineImageSummary = {
  id: string;
  contentType?: string;
  downloadable?: boolean;
  verificationRequired?: boolean;
};

export type TeamsInlineImageGraphClient = {
  get<T>(pathOrUrl: string): Promise<T>;
  download(pathOrUrl: string): Promise<DownloadResponse>;
};

type GraphHostedContent = {
  id: string;
};

export async function listTeamsInlineImages(
  config: AppConfig,
  chatId: string,
  messageId: string,
  limit: number,
  client: TeamsInlineImageGraphClient = defaultInlineImageClient(config)
): Promise<{
  list: { chatId: string; messageId: string; returnedCount: number; limitReached: boolean };
  images: TeamsInlineImageSummary[];
}> {
  const normalizedChatId = requireIdentifier(chatId, "chat-id");
  const normalizedMessageId = requireIdentifier(messageId, "message-id");
  if (!Number.isFinite(limit) || limit < 1) throw new Error("limit must be a positive number.");
  const maximum = Math.min(Math.floor(limit), 100);
  const images: TeamsInlineImageSummary[] = [];
  let nextUrl: string | undefined = buildInlineImageListPath(normalizedChatId, normalizedMessageId);

  while (nextUrl && images.length < maximum) {
    const page: GraphPage<GraphHostedContent> = await client.get(nextUrl);
    const remaining = maximum - images.length;
    images.push(
      ...(page.value ?? []).slice(0, remaining).map((content) => ({
        id: content.id,
        verificationRequired: true
      }))
    );
    nextUrl = page["@odata.nextLink"];
  }

  return {
    list: {
      chatId: normalizedChatId,
      messageId: normalizedMessageId,
      returnedCount: images.length,
      limitReached: Boolean(nextUrl)
    },
    images
  };
}

export async function downloadTeamsInlineImage(
  config: AppConfig,
  chatId: string,
  messageId: string,
  hostedContentId: string,
  filename?: string,
  client: TeamsInlineImageGraphClient = defaultInlineImageClient(config)
): Promise<{
  stage: "DOWNLOADED";
  outputPath: string;
  image: TeamsInlineImageSummary;
}> {
  if (!config.policy.allowDownloads) throw new Error("Downloads are disabled by policy.");
  const normalizedChatId = requireIdentifier(chatId, "chat-id");
  const normalizedMessageId = requireIdentifier(messageId, "message-id");
  const normalizedHostedContentId = requireIdentifier(hostedContentId, "hosted-content-id");
  const response = await client.download(
    buildInlineImageContentPath(
      normalizedChatId,
      normalizedMessageId,
      normalizedHostedContentId
    )
  );
  const contentType = response.headers.get("content-type")?.split(";", 1)[0]?.trim().toLowerCase();
  const extension = extensionFor(contentType);
  if (!extension) {
    await response.body?.cancel();
    throw new Error("The selected Teams hosted content is not an inline image.");
  }
  const image: TeamsInlineImageSummary = {
    id: normalizedHostedContentId,
    contentType,
    downloadable: true
  };
  const requestedName = sanitizeFilename(filename?.trim() || `teams-inline-${normalizedHostedContentId}`);
  const outputName = sanitizeFilename(`${path.parse(requestedName).name || "teams-inline"}${extension}`);
  const outputPath = await saveDownloadResponse(
    config,
    response,
    outputName,
    config.policy.maxDownloadBytes
  );
  return { stage: "DOWNLOADED", outputPath, image };
}

export function buildInlineImageListPath(chatId: string, messageId: string): string {
  return buildHostedContentsBasePath(chatId, messageId);
}

export function buildInlineImageContentPath(
  chatId: string,
  messageId: string,
  hostedContentId: string
): string {
  return `${buildHostedContentsBasePath(chatId, messageId)}/${encodeURIComponent(hostedContentId)}/$value`;
}

function buildHostedContentsBasePath(chatId: string, messageId: string): string {
  return `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}/hostedContents`;
}

function extensionFor(contentType: string | undefined): string | undefined {
  switch (contentType?.toLowerCase()) {
    case "image/jpeg": return ".jpg";
    case "image/gif": return ".gif";
    case "image/webp": return ".webp";
    case "image/bmp":
    case "image/x-ms-bmp": return ".bmp";
    case "image/tiff": return ".tiff";
    case "image/png": return ".png";
    default: return undefined;
  }
}

function requireIdentifier(value: string, optionName: string): string {
  const normalized = value.trim();
  if (!normalized) throw new Error(`${optionName} must not be empty.`);
  return normalized;
}

function defaultInlineImageClient(config: AppConfig): TeamsInlineImageGraphClient {
  return {
    get: <T>(pathOrUrl: string) => graphGet<T>(config, pathOrUrl),
    download: (pathOrUrl: string) => graphDownloadResponse(config, pathOrUrl)
  };
}
