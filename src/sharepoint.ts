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

type GraphDriveItem = FileSummary & {
  file?: unknown;
  folder?: unknown;
};

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
