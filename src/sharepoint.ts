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
import {
  graphDownloadResponse,
  graphGet,
  graphPost,
  type GraphPage,
  type GraphRequestOptions
} from "./graph.js";

export type FileSummary = {
  id: string;
  name?: string;
  webUrl?: string;
  size?: number;
  lastModifiedDateTime?: string;
  mimeType?: string;
  parentReference?: {
    driveId?: string;
    siteId?: string;
    path?: string;
  };
};

export type FileSearchResult = {
  search: {
    query: string;
    scope: "sharePointAndOneDrive";
    source: "microsoftSearch";
    returnedCount: number;
    totalMatchesReported: number;
    maxResults: number;
    limitReached: boolean;
    offset: number;
    nextOffset?: number;
    continuationAvailable: boolean;
    partialResult: boolean;
    partialReason?: "time-budget-exceeded";
    timeBudgetMs: number;
  };
  files: FileSummary[];
};

export type FileSearchOptions = {
  offset?: number;
  timeBudgetMs?: number;
  now?: () => number;
};

export type SharePointGraphClient = {
  post<T>(pathOrUrl: string, body: unknown, options?: GraphRequestOptions): Promise<T>;
};

export type DriveItemDownloadGraphClient = {
  get<T>(pathOrUrl: string): Promise<T>;
  download(pathOrUrl: string): Promise<DownloadResponse>;
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
  file?: { mimeType?: string };
  folder?: unknown;
};

type GraphSite = SiteSummary;

type GraphDriveItemSearchHit = { resource?: GraphDriveItem };
type GraphDriveItemSearchResponse = {
  value?: Array<{
    hitsContainers?: Array<{
      total?: number;
      moreResultsAvailable?: boolean;
      hits?: GraphDriveItemSearchHit[];
    }>;
  }>;
};

const defaultSearchTimeBudgetMs = 35_000;

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

export async function searchFiles(
  config: AppConfig,
  query: string,
  limit: number,
  options: FileSearchOptions = {},
  client: SharePointGraphClient = defaultSharePointClient(config)
): Promise<FileSearchResult> {
  const trimmedQuery = query.trim();
  if (!trimmedQuery) throw new Error("query must not be empty.");
  if (!Number.isFinite(limit) || limit < 1) throw new Error("limit must be a positive number.");

  const maxResults = Math.min(Math.floor(limit), config.policy.maxFileSearchLimit);
  const offset = normalizeOffset(options.offset ?? 0);
  const timeBudgetMs = normalizeTimeBudget(options.timeBudgetMs ?? defaultSearchTimeBudgetMs);
  const now = options.now ?? Date.now;
  const deadline = now() + timeBudgetMs;
  let totalMatchesReported = 0;
  let moreResultsAvailable = false;
  let partialResult = false;
  let partialReason: FileSearchResult["search"]["partialReason"];
  let hits: GraphDriveItemSearchHit[] = [];

  try {
    const response = await client.post<GraphDriveItemSearchResponse>(
      "/search/query",
      {
        requests: [
          {
            entityTypes: ["driveItem"],
            query: {
              queryString: trimmedQuery,
              queryTemplate: "({searchTerms}) IsDocument:True"
            },
            fields: [
              "id",
              "name",
              "webUrl",
              "size",
              "lastModifiedDateTime",
              "parentReference",
              "file",
              "folder"
            ],
            from: offset,
            size: maxResults
          }
        ]
      },
      { totalTimeoutMs: Math.max(1, deadline - now()) }
    );
    const container = response.value?.[0]?.hitsContainers?.[0];
    totalMatchesReported = container?.total ?? 0;
    moreResultsAvailable = Boolean(container?.moreResultsAvailable);
    hits = container?.hits ?? [];
  } catch (error) {
    if (now() >= deadline || isGraphTimeBudgetError(error)) {
      partialResult = true;
      partialReason = "time-budget-exceeded";
    } else {
      throw error;
    }
  }

  const files = hits
    .map((hit) => hit.resource)
    .filter((item): item is GraphDriveItem => Boolean(item?.file))
    .map((item) => ({
      id: item.id,
      name: item.name,
      webUrl: item.webUrl,
      size: item.size,
      lastModifiedDateTime: item.lastModifiedDateTime,
      mimeType: item.file?.mimeType,
      parentReference: item.parentReference
    }));
  const continuationAvailable = partialResult || moreResultsAvailable;

  return {
    search: {
      query: trimmedQuery,
      scope: "sharePointAndOneDrive",
      source: "microsoftSearch",
      returnedCount: files.length,
      totalMatchesReported,
      maxResults,
      limitReached: continuationAvailable,
      offset,
      nextOffset: continuationAvailable ? offset + hits.length : undefined,
      continuationAvailable,
      partialResult,
      partialReason,
      timeBudgetMs
    },
    files
  };
}

export async function downloadDriveItem(
  config: AppConfig,
  driveId: string,
  itemId: string,
  filename?: string,
  approvalToken?: string,
  client: DriveItemDownloadGraphClient = defaultDriveItemDownloadClient(config)
): Promise<string | PendingDownloadApproval> {
  if (!config.policy.allowDownloads) {
    throw new Error("Downloads are disabled by policy.");
  }

  if (!driveId.trim()) throw new Error("drive-id must not be empty.");
  if (!itemId.trim()) throw new Error("item-id must not be empty.");

  const metadata = await client.get<GraphDriveItem>(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}?$select=id,name,size,file,folder`
  );
  if (!metadata.file || metadata.folder) throw new Error("The selected drive item is not a file.");
  const outputName = sanitizeFilename(filename?.trim() || metadata.name || `${itemId}.bin`);
  const authorization = authorizeDownload(
    config,
    {
      source: "sharepoint",
      resourceKey: `${driveId}\u0000${itemId}`,
      name: metadata.name || outputName,
      size: metadata.size ?? Number.NaN,
      outputName
    },
    approvalToken
  );
  if (!authorization.approved) return authorization.pending;

  const response = await client.download(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`
  );
  return saveDownloadResponse(config, response, outputName, authorization.maxBytes);
}

export { sanitizeFilename } from "./downloads.js";

function defaultSharePointClient(config: AppConfig): SharePointGraphClient {
  return {
    post: <T>(pathOrUrl: string, body: unknown, options?: GraphRequestOptions) =>
      graphPost<T>(config, pathOrUrl, body, options)
  };
}

function defaultDriveItemDownloadClient(config: AppConfig): DriveItemDownloadGraphClient {
  return {
    get: <T>(pathOrUrl: string) => graphGet<T>(config, pathOrUrl),
    download: (pathOrUrl: string) => graphDownloadResponse(config, pathOrUrl)
  };
}

function normalizeOffset(value: number): number {
  if (!Number.isFinite(value) || value < 0) {
    throw new Error("offset must be a non-negative number.");
  }
  return Math.floor(value);
}

function normalizeTimeBudget(value: number): number {
  if (!Number.isFinite(value) || value < 1) {
    throw new Error("time budget must be a positive number.");
  }
  return Math.floor(value);
}

function isGraphTimeBudgetError(error: unknown): boolean {
  if (!(error instanceof Error)) return false;
  return error.name === "GraphTimeoutError" || /time budget|timed out/i.test(error.message);
}
