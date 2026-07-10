import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";
import { encodeQuery, graphDownload, graphGet, type GraphPage } from "./graph.js";

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
  const top = Math.min(limit, config.policy.maxFileSearchLimit);
  const page = await graphGet<GraphPage<GraphDriveItem>>(
    config,
    `/me/drive/root/search(q='${encodeQuery(query)}')?$top=${top}&$select=id,name,webUrl,size,lastModifiedDateTime,parentReference,file,folder`
  );

  return (page.value ?? [])
    .filter((item) => item.file)
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

  const bytes = await graphDownload(
    config,
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`
  );
  fs.mkdirSync(config.downloadDir, { recursive: true });
  const safeName = sanitizeFilename(filename || `${itemId}.bin`);
  const outputPath = path.join(config.downloadDir, safeName);
  fs.writeFileSync(outputPath, bytes);
  return outputPath;
}

function sanitizeFilename(value: string): string {
  return value.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").slice(0, 180);
}
