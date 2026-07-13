import fs from "node:fs";
import path from "node:path";
import { Readable, Transform } from "node:stream";
import { pipeline } from "node:stream/promises";
import type { AppConfig } from "./config.js";
import { encodeQuery, graphDownloadResponse, graphGet, type GraphPage } from "./graph.js";

export type FileSummary = {
  id: string;
  name?: string;
  webUrl?: string;
  size?: number;
  lastModifiedDateTime?: string;
  parentReference?: {
    driveId?: string;
    siteId?: string;
    path?: string;
  };
};

export type SiteSummary = {
  id: string;
  displayName?: string;
  name?: string;
  description?: string;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  isPersonalSite?: boolean;
};

export type SiteSearchResult = {
  search: {
    query: string;
    scope: "sharepointSites";
    returnedCount: number;
    maxResults: number;
    limitReached: boolean;
  };
  sites: SiteSummary[];
};

type GraphDriveItem = FileSummary & {
  file?: unknown;
  folder?: unknown;
};

type GraphSite = SiteSummary;

export async function searchSites(
  config: AppConfig,
  query: string,
  limit: number
): Promise<SiteSearchResult> {
  const trimmedQuery = query.trim();
  if (!trimmedQuery) {
    throw new Error("query must not be empty.");
  }
  if (!Number.isFinite(limit) || limit < 1) {
    throw new Error("limit must be a positive number.");
  }

  const maxResults = Math.min(Math.floor(limit), config.policy.maxFileSearchLimit);
  let nextUrl: string | undefined = `/sites?search=${encodeURIComponent(trimmedQuery)}`;
  const sites: GraphSite[] = [];
  let limitReached = false;

  while (nextUrl && sites.length < maxResults) {
    const page: GraphPage<GraphSite> = await graphGet<GraphPage<GraphSite>>(config, nextUrl);
    const remaining = maxResults - sites.length;
    const values = page.value ?? [];
    sites.push(...values.slice(0, remaining));
    nextUrl = page["@odata.nextLink"];
    limitReached = Boolean(nextUrl) || values.length > remaining;
  }

  return {
    search: {
      query: trimmedQuery,
      scope: "sharepointSites",
      returnedCount: sites.length,
      maxResults,
      limitReached
    },
    sites: sites.map((site) => ({
      id: site.id,
      displayName: site.displayName,
      name: site.name,
      description: site.description,
      webUrl: site.webUrl,
      createdDateTime: site.createdDateTime,
      lastModifiedDateTime: site.lastModifiedDateTime,
      isPersonalSite: site.isPersonalSite
    }))
  };
}

export async function searchFiles(config: AppConfig, query: string, limit: number): Promise<FileSummary[]> {
  const trimmedQuery = query.trim();
  if (!trimmedQuery) throw new Error("query must not be empty.");
  if (!Number.isFinite(limit) || limit < 1) throw new Error("limit must be a positive number.");

  const top = Math.min(Math.floor(limit), config.policy.maxFileSearchLimit);
  let nextUrl: string | undefined =
    `/me/drive/root/search(q='${encodeQuery(trimmedQuery)}')?$top=${top}&$select=id,name,webUrl,size,lastModifiedDateTime,parentReference,file,folder`;
  const files: GraphDriveItem[] = [];

  while (nextUrl && files.length < top) {
    const page: GraphPage<GraphDriveItem> = await graphGet<GraphPage<GraphDriveItem>>(config, nextUrl);
    const eligible = (page.value ?? []).filter((item) => item.file);
    files.push(...eligible.slice(0, top - files.length));
    nextUrl = page["@odata.nextLink"];
  }

  return files
    .map((item) => ({
      id: item.id,
      name: item.name,
      webUrl: item.webUrl,
      size: item.size,
      lastModifiedDateTime: item.lastModifiedDateTime,
      parentReference: item.parentReference
    }));
}

export async function downloadDriveItem(
  config: AppConfig,
  driveId: string,
  itemId: string,
  filename?: string
): Promise<string> {
  if (!config.policy.allowDownloads) {
    throw new Error("Downloads are disabled by policy.");
  }

  if (!driveId.trim()) throw new Error("drive-id must not be empty.");
  if (!itemId.trim()) throw new Error("item-id must not be empty.");

  const response = await graphDownloadResponse(
    config,
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`
  );
  const contentLength = Number(response.headers.get("content-length"));
  if (Number.isFinite(contentLength) && contentLength > config.policy.maxDownloadBytes) {
    await response.body?.cancel();
    throw new Error(`Download exceeds policy limit of ${config.policy.maxDownloadBytes} bytes.`);
  }
  if (!response.body) throw new Error("Download response contained no file body.");

  fs.mkdirSync(config.downloadDir, { recursive: true });
  const safeName = sanitizeFilename(filename || `${itemId}.bin`);
  const outputPath = uniqueOutputPath(config.downloadDir, safeName);
  let receivedBytes = 0;
  const limiter = new Transform({
    transform(chunk: Buffer, _encoding, callback) {
      receivedBytes += chunk.length;
      if (receivedBytes > config.policy.maxDownloadBytes) {
        callback(new Error(`Download exceeds policy limit of ${config.policy.maxDownloadBytes} bytes.`));
        return;
      }
      callback(null, chunk);
    }
  });

  try {
    await pipeline(
      Readable.fromWeb(response.body as never),
      limiter,
      fs.createWriteStream(outputPath, { flags: "wx", mode: 0o600 })
    );
  } catch (error) {
    fs.rmSync(outputPath, { force: true });
    throw error;
  }
  return outputPath;
}

export function sanitizeFilename(value: string): string {
  const sanitized = value
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
    .replace(/[. ]+$/g, "")
    .trim()
    .slice(0, 180);
  const fallback = sanitized || "download.bin";
  const baseName = path.parse(fallback).name;
  return /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(baseName) ? `_${fallback}` : fallback;
}

function uniqueOutputPath(directory: string, filename: string): string {
  const parsed = path.parse(filename);
  for (let index = 0; index < 10_000; index += 1) {
    const suffix = index === 0 ? "" : ` (${index})`;
    const candidate = path.join(directory, `${parsed.name}${suffix}${parsed.ext}`);
    if (!fs.existsSync(candidate)) return candidate;
  }
  throw new Error("Could not allocate a unique download filename.");
}
