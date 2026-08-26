import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";
import {
  finalizeTemporaryDownload,
  saveDownloadResponseToPath,
  sanitizeFilename,
  type DownloadResponse
} from "./downloads.js";
import {
  authorizeSharePointExport,
  completeSharePointExportApproval,
  hashExportPlan,
  type SharePointExportPlan
} from "./sharepoint-export-approval.js";
import {
  graphDownloadResponse,
  graphGet,
  type GraphPage
} from "./graph.js";
import {
  clearStoredFile,
  usesDeleteRestrictedStorage,
  writeStoredText
} from "./persistent-storage.js";

type GraphSite = {
  id: string;
  displayName?: string;
  webUrl?: string;
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
  eTag?: string;
  file?: unknown;
  folder?: { childCount?: number };
};

type ExportFile = {
  itemId: string;
  relativePath: string;
  size: number;
  eTag?: string;
};

type UnsupportedItem = {
  itemId: string;
  relativePath: string;
  reason: string;
};

export type SharePointExportGraphClient = {
  get<T>(pathOrUrl: string): Promise<T>;
  download(
    pathOrUrl: string,
    options?: { rangeStart?: number; eTag?: string }
  ): Promise<DownloadResponse & { status?: number }>;
};

export type SharePointExportOptions = {
  accountId?: string;
  approvalToken?: string;
  timeBudgetMs?: number;
  now?: () => number;
  availableDiskBytes?: (destinationRoot: string) => number;
  copyFinalizationRequired?: (outputPath: string) => boolean;
};

export type SharePointExportResult =
  | {
      stage: "ENUMERATION_IN_PROGRESS";
      scanId: string;
      siteUrl: string;
      destinationRoot: string;
      discoveredFileCount: number;
      discoveredFolderCount: number;
      instruction: string;
    }
  | {
      stage: "AWAITING_USER_APPROVAL";
      jobId: string;
      preview: ExportPreview;
      approval: {
        token: string;
        expiresAt: string;
        instruction: string;
      };
    }
  | {
      stage: "BLOCKED_INSUFFICIENT_DISK";
      jobId: string;
      preview: ExportPreview;
      error: string;
    }
  | {
      stage: "COMPLETE" | "COMPLETE_WITH_ERRORS" | "IN_PROGRESS";
      jobId: string;
      siteUrl: string;
      destinationRoot: string;
      downloadedCount: number;
      downloadedBytes: number;
      resumedFileCount: number;
      skippedExistingCount: number;
      skippedOversizedCount: number;
      unsupportedItemCount: number;
      failedCount: number;
      failures: Array<{ relativePath: string; error: string }>;
      approvalReused: boolean;
      instruction?: string;
    };

type ExportPreview = {
  siteUrl: string;
  siteName?: string;
  libraryName?: string;
  destinationRoot: string;
  fileCount: number;
  folderCount: number;
  totalBytes: number;
  pendingDownloadCount: number;
  pendingDownloadBytes: number;
  finalizationReserveBytes: number;
  requiredDiskBytes: number;
  existingFileCount: number;
  existingSizeConflictCount: number;
  oversizedFileCount: number;
  unsupportedItemCount: number;
  renamedPathCount: number;
  availableDiskBytes: number;
  maxFileBytes: number;
  oversizedFiles: Array<{ relativePath: string; size: number }>;
  warning: string;
};

type ExportManifest = {
  checkpointPath: string;
  plan: SharePointExportPlan;
  siteName?: string;
  libraryName?: string;
  folderCount: number;
  renamedPathCount: number;
  unsupported: UnsupportedItem[];
};

type ManifestCheckpoint = {
  schemaVersion: 1;
  tenantId: string;
  accountBindingHash: string;
  siteUrl: string;
  destinationRoot: string;
  maxFileBytes: number;
  siteId: string;
  siteName?: string;
  driveId: string;
  libraryName?: string;
  files: ExportFile[];
  unsupported: UnsupportedItem[];
  folders: Array<{ itemId?: string; segments: string[] }>;
  folderIndex: number;
  currentItems: GraphDriveItem[];
  currentNextUrl: string | null;
  seenNextLinks: string[];
  folderCount: number;
  renamedPathCount: number;
  complete: boolean;
};

type ManifestBuildResult =
  | { complete: true; manifest: ExportManifest }
  | {
      complete: false;
      scanId: string;
      siteUrl: string;
      discoveredFileCount: number;
      discoveredFolderCount: number;
    };


const allowedSharePointHost = "ohmylab.sharepoint.com";
const defaultTimeBudgetMs = 4 * 60 * 1000;

export async function exportSharePointFiles(
  config: AppConfig,
  siteUrl: string,
  destination: string,
  options: SharePointExportOptions = {},
  client: SharePointExportGraphClient = defaultClient(config)
): Promise<SharePointExportResult> {
  if (!config.policy.allowDownloads) {
    throw new Error("Downloads are disabled by policy.");
  }
  const maxFileBytes = config.policy.maxSharePointExportFileBytes;
  if (!Number.isFinite(maxFileBytes) || maxFileBytes < 1) {
    throw new Error("SharePoint export policy has an invalid maxSharePointExportFileBytes.");
  }

  const timeBudgetMs = normalizeTimeBudget(options.timeBudgetMs ?? defaultTimeBudgetMs);
  const now = options.now ?? Date.now;
  const deadline = now() + timeBudgetMs;
  const accountId = options.accountId?.trim();
  if (!accountId || !config.tenantId) {
    throw new Error("A logged-in account binding is required for SharePoint export.");
  }
  const accountBindingHash = createHash("sha256")
    .update(`${config.tenantId}\u0000${accountId}`)
    .digest("hex");
  const destinationRoot = normalizeDestination(destination);
  assertSafeDestinationRoot(destinationRoot);
  const manifestBuild = await createManifest(
    config,
    accountBindingHash,
    siteUrl,
    destinationRoot,
    maxFileBytes,
    deadline,
    now,
    client
  );
  if (!manifestBuild.complete) {
    return {
      stage: "ENUMERATION_IN_PROGRESS",
      scanId: manifestBuild.scanId,
      siteUrl: manifestBuild.siteUrl,
      destinationRoot,
      discoveredFileCount: manifestBuild.discoveredFileCount,
      discoveredFolderCount: manifestBuild.discoveredFolderCount,
      instruction:
        "Rerun the exact same command without an approval token to continue the saved SharePoint enumeration."
    };
  }
  const manifest = manifestBuild.manifest;
  const jobId = hashExportPlan(manifest.plan).slice(0, 16);
  const diskSpace = options.availableDiskBytes ?? availableDiskBytes;
  const forceCopyFinalization =
    options.copyFinalizationRequired ?? usesDeleteRestrictedStorage;
  const preview = buildPreview(
    manifest,
    diskSpace,
    jobId,
    () => true
  );

  if (preview.requiredDiskBytes > preview.availableDiskBytes) {
    return {
      stage: "BLOCKED_INSUFFICIENT_DISK",
      jobId,
      preview,
      error:
        "The destination does not have enough free space for the files that are not already present."
    };
  }

  if (preview.pendingDownloadCount === 0) {
    completeSharePointExportApproval(config, manifest.plan);
    clearStoredFile(manifest.checkpointPath);
    const failures: Array<{ relativePath: string; error: string }> = [];
    return {
      stage: "COMPLETE",
      jobId,
      siteUrl: manifest.plan.siteUrl,
      destinationRoot,
      downloadedCount: 0,
      downloadedBytes: 0,
      resumedFileCount: 0,
      skippedExistingCount: preview.existingFileCount,
      skippedOversizedCount: preview.oversizedFileCount,
      unsupportedItemCount: manifest.unsupported.length,
      failedCount: failures.length,
      failures,
      approvalReused: false
    };
  }

  const authorization = authorizeSharePointExport(
    config,
    manifest.plan,
    options.approvalToken
  );
  if (!authorization.approved) {
    return {
      stage: "AWAITING_USER_APPROVAL",
      jobId: authorization.jobId,
      preview,
      approval: {
        token: authorization.token,
        expiresAt: authorization.expiresAt,
        instruction:
          "Show this complete batch preview and stop. After explicit approval, rerun the exact same command once with --approval-token. The approved unchanged job can then resume for 24 hours without per-file approvals."
      }
    };
  }

  const failures: Array<{ relativePath: string; error: string }> = [];
  let downloadedCount = 0;
  let downloadedBytes = 0;
  let resumedFileCount = 0;
  let skippedExistingCount = 0;
  let skippedOversizedCount = 0;
  let interrupted = false;
  let sourceChanged = false;

  fs.mkdirSync(destinationRoot, { recursive: true });
  assertNoSymbolicLinks(destinationRoot, destinationRoot);

  for (const file of manifest.plan.files) {
    if (file.size > maxFileBytes) {
      skippedOversizedCount += 1;
      continue;
    }
    const outputPath = resolveOutputPath(destinationRoot, file.relativePath);
    if (fs.existsSync(outputPath)) {
      skippedExistingCount += 1;
      continue;
    }
    if (now() >= deadline) {
      interrupted = true;
      break;
    }

    try {
      assertNoSymbolicLinks(destinationRoot, outputPath);
      fs.mkdirSync(path.dirname(outputPath), { recursive: true });
      assertNoSymbolicLinks(destinationRoot, outputPath);

      const freeBytes = diskSpace(destinationRoot);
      const partialPath = partialOutputPath(outputPath, authorization.jobId);
      const partialSize = fs.existsSync(partialPath) ? fs.statSync(partialPath).size : 0;
      const safeOffset = partialSize > 0 && partialSize <= file.size ? partialSize : 0;
      const remainingBytes = file.size - safeOffset;
      const requiresCopyFinalization = forceCopyFinalization(outputPath);
      const requiredFreeBytes =
        remainingBytes + file.size;
      if (requiredFreeBytes > freeBytes) {
        throw new Error("Insufficient destination disk space for this file.");
      }
      if (safeOffset === file.size) {
        await verifySourceItem(client, manifest.plan.driveId, file);
        finalizeTemporaryDownload(partialPath, outputPath, {
          deleteRestricted: requiresCopyFinalization
        });
        downloadedCount += 1;
        downloadedBytes += file.size;
        resumedFileCount += 1;
        continue;
      }

      const response = await client.download(
        `/drives/${encodeURIComponent(manifest.plan.driveId)}/items/${encodeURIComponent(file.itemId)}/content`,
        {
          rangeStart: safeOffset > 0 ? safeOffset : undefined,
          eTag: file.eTag
        }
      );
      await saveDownloadResponseToPath(response, outputPath, {
        maxBytes: maxFileBytes,
        expectedBytes: file.size,
        startOffset: safeOffset,
        temporarySuffix: `.hare-part-${authorization.jobId}`,
        forceCopyFinalization: requiresCopyFinalization
      });
      downloadedCount += 1;
      downloadedBytes += file.size;
      if (safeOffset > 0) resumedFileCount += 1;
    } catch (error) {
      const message = safeErrorMessage(error);
      if (isSourceChangedError(message)) sourceChanged = true;
      failures.push({
        relativePath: file.relativePath,
        error: message
      });
    }
  }

  const stage = interrupted
    ? "IN_PROGRESS"
    : failures.length > 0
      ? "COMPLETE_WITH_ERRORS"
      : "COMPLETE";
  if (stage === "COMPLETE" || sourceChanged) {
    completeSharePointExportApproval(config, manifest.plan);
    clearStoredFile(manifest.checkpointPath);
  }

  return {
    stage,
    jobId: authorization.jobId,
    siteUrl: manifest.plan.siteUrl,
    destinationRoot,
    downloadedCount,
    downloadedBytes,
    resumedFileCount,
    skippedExistingCount,
    skippedOversizedCount,
    unsupportedItemCount: manifest.unsupported.length,
    failedCount: failures.length,
    failures,
    approvalReused: authorization.resumed,
    instruction: sourceChanged
      ? "The SharePoint source changed after approval. Rerun without an approval token to build and review a new manifest."
      : stage === "IN_PROGRESS" || stage === "COMPLETE_WITH_ERRORS"
        ? "Rerun the exact same command without an approval token to continue the approved unchanged job."
        : undefined
  };
}

async function createManifest(
  config: AppConfig,
  accountBindingHash: string,
  siteUrl: string,
  destinationRoot: string,
  maxFileBytes: number,
  deadline: number,
  now: () => number,
  client: SharePointExportGraphClient
): Promise<ManifestBuildResult> {
  const site = normalizeSiteUrl(siteUrl);
  const checkpointPath = manifestCheckpointPath(
    config,
    accountBindingHash,
    site.canonicalUrl,
    destinationRoot,
    maxFileBytes
  );
  const scanId = path.basename(checkpointPath, ".json").replace(
    "sharepoint-export-manifest-",
    ""
  );
  let state = readManifestCheckpoint(
    checkpointPath,
    config.tenantId,
    accountBindingHash,
    site.canonicalUrl,
    destinationRoot,
    maxFileBytes
  );

  if (!state) {
    if (now() >= deadline) {
      return {
        complete: false,
        scanId,
        siteUrl: site.canonicalUrl,
        discoveredFileCount: 0,
        discoveredFolderCount: 0
      };
    }
    const graphSite = await client.get<GraphSite>(site.graphPath);
    if (!graphSite.id) throw new Error("SharePoint site lookup returned no site ID.");
    if (now() >= deadline) {
      return {
        complete: false,
        scanId,
        siteUrl: site.canonicalUrl,
        discoveredFileCount: 0,
        discoveredFolderCount: 0
      };
    }
    const drive = await client.get<GraphDrive>(
      `/sites/${encodeURIComponent(graphSite.id)}/drive?$select=id,name,webUrl`
    );
    if (!drive.id) throw new Error("SharePoint site has no default document library.");
    state = {
      schemaVersion: 1,
      tenantId: config.tenantId,
      accountBindingHash,
      siteUrl: site.canonicalUrl,
      destinationRoot,
      maxFileBytes,
      siteId: graphSite.id,
      siteName: graphSite.displayName,
      driveId: drive.id,
      libraryName: drive.name,
      files: [],
      unsupported: [],
      folders: [{ segments: [] }],
      folderIndex: 0,
      currentItems: [],
      currentNextUrl: null,
      seenNextLinks: [],
      folderCount: 0,
      renamedPathCount: 0,
      complete: false
    };
    writeManifestCheckpoint(checkpointPath, state);
  }

  while (!state.complete && state.folderIndex < state.folders.length) {
    const folder = state.folders[state.folderIndex];
    if (!state.currentNextUrl) {
      state.currentNextUrl = childrenUrl(state.driveId, folder.itemId);
      state.seenNextLinks = [];
    }
    if (now() >= deadline) {
      writeManifestCheckpoint(checkpointPath, state);
      return {
        complete: false,
        scanId,
        siteUrl: state.siteUrl,
        discoveredFileCount: state.files.length,
        discoveredFolderCount: state.folderCount
      };
    }

    const requestUrl = state.currentNextUrl;
    if (state.seenNextLinks.includes(requestUrl)) {
      throw new Error("Graph pagination repeated a nextLink while enumerating SharePoint.");
    }
    state.seenNextLinks.push(requestUrl);
    const page = await client.get<GraphPage<GraphDriveItem>>(requestUrl);
    state.currentItems.push(...(page.value ?? []));
    const nextUrl = validateNextLink(page["@odata.nextLink"]);
    if (nextUrl) {
      state.currentNextUrl = nextUrl;
      writeManifestCheckpoint(checkpointPath, state);
      continue;
    }

    processFolderItems(state, folder);
    state.folderIndex += 1;
    state.currentItems = [];
    state.currentNextUrl = null;
    state.seenNextLinks = [];
    writeManifestCheckpoint(checkpointPath, state);
  }

  state.complete = state.folderIndex >= state.folders.length;
  writeManifestCheckpoint(checkpointPath, state);
  if (!state.complete) {
    return {
      complete: false,
      scanId,
      siteUrl: state.siteUrl,
      discoveredFileCount: state.files.length,
      discoveredFolderCount: state.folderCount
    };
  }
  return { complete: true, manifest: manifestFromCheckpoint(checkpointPath, state) };
}

function processFolderItems(
  state: ManifestCheckpoint,
  folder: { itemId?: string; segments: string[] }
): void {
  state.currentItems.sort((left, right) =>
    `${left.name ?? ""}\u0000${left.id}`.localeCompare(
      `${right.name ?? ""}\u0000${right.id}`
    )
  );
  const usedSegments = new Set<string>();

  for (const item of state.currentItems) {
    const originalName = item.name?.trim() || `unnamed-${shortId(item.id)}`;
    const segment = allocateSegment(originalName, item.id, usedSegments);
    if (segment !== originalName) state.renamedPathCount += 1;
    const segments = [...folder.segments, segment];
    const relativePath = segments.join("/");

    if (item.folder) {
      state.folderCount += 1;
      state.folders.push({ itemId: item.id, segments });
      continue;
    }
    if (!item.file) {
      state.unsupported.push({
        itemId: item.id,
        relativePath,
        reason: "Unsupported drive item type."
      });
      continue;
    }
    if (!Number.isFinite(item.size) || (item.size ?? -1) < 0) {
      state.unsupported.push({
        itemId: item.id,
        relativePath,
        reason: "File size metadata is unavailable."
      });
      continue;
    }
    if (!item.eTag?.trim()) {
      state.unsupported.push({
        itemId: item.id,
        relativePath,
        reason: "File eTag metadata is unavailable."
      });
      continue;
    }
    state.files.push({
      itemId: item.id,
      relativePath,
      size: item.size as number,
      eTag: item.eTag
    });
  }
}

function childrenUrl(driveId: string, itemId?: string): string {
  const select = "$select=id,name,size,eTag,file,folder";
  return itemId
    ? `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children?${select}&$top=200`
    : `/drives/${encodeURIComponent(driveId)}/root/children?${select}&$top=200`;
}

function manifestFromCheckpoint(
  checkpointPath: string,
  state: ManifestCheckpoint
): ExportManifest {
  const files = [...state.files].sort((left, right) =>
    left.relativePath.localeCompare(right.relativePath)
  );
  return {
    checkpointPath,
    plan: {
      schemaVersion: 1,
      tenantId: state.tenantId,
      accountBindingHash: state.accountBindingHash,
      siteUrl: state.siteUrl,
      siteId: state.siteId,
      driveId: state.driveId,
      destinationRoot: state.destinationRoot,
      maxFileBytes: state.maxFileBytes,
      files
    },
    siteName: state.siteName,
    libraryName: state.libraryName,
    folderCount: state.folderCount,
    renamedPathCount: state.renamedPathCount,
    unsupported: state.unsupported
  };
}

function manifestCheckpointPath(
  config: AppConfig,
  accountBindingHash: string,
  siteUrl: string,
  destinationRoot: string,
  maxFileBytes: number
): string {
  const key = createHash("sha256")
    .update(
      JSON.stringify({
        schemaVersion: 1,
        tenantId: config.tenantId,
        accountBindingHash,
        siteUrl,
        destinationRoot,
        maxFileBytes
      })
    )
    .digest("hex")
    .slice(0, 24);
  return path.join(config.cacheDir, `sharepoint-export-manifest-${key}.json`);
}

function readManifestCheckpoint(
  checkpointPath: string,
  tenantId: string,
  accountBindingHash: string,
  siteUrl: string,
  destinationRoot: string,
  maxFileBytes: number
): ManifestCheckpoint | undefined {
  if (!fs.existsSync(checkpointPath) || fs.statSync(checkpointPath).size === 0) {
    return undefined;
  }
  try {
    const state = JSON.parse(
      fs.readFileSync(checkpointPath, "utf8")
    ) as ManifestCheckpoint;
    if (
      state.schemaVersion !== 1 ||
      state.tenantId !== tenantId ||
      state.accountBindingHash !== accountBindingHash ||
      state.siteUrl !== siteUrl ||
      state.destinationRoot !== destinationRoot ||
      state.maxFileBytes !== maxFileBytes ||
      !state.siteId ||
      !state.driveId ||
      !Array.isArray(state.files) ||
      !Array.isArray(state.unsupported) ||
      !Array.isArray(state.folders) ||
      !Array.isArray(state.currentItems) ||
      !Array.isArray(state.seenNextLinks)
    ) {
      throw new Error("invalid checkpoint");
    }
    return state;
  } catch {
    clearStoredFile(checkpointPath);
    return undefined;
  }
}

function writeManifestCheckpoint(
  checkpointPath: string,
  state: ManifestCheckpoint
): void {
  writeStoredText(checkpointPath, JSON.stringify(state));
}


function buildPreview(
  manifest: ExportManifest,
  diskSpace: (destinationRoot: string) => number,
  jobId: string,
  copyFinalizationRequired: (outputPath: string) => boolean
): ExportPreview {
  let pendingDownloadCount = 0;
  let pendingDownloadBytes = 0;
  let finalizationReserveBytes = 0;
  let existingFileCount = 0;
  let existingSizeConflictCount = 0;
  const oversizedFiles: Array<{ relativePath: string; size: number }> = [];

  for (const file of manifest.plan.files) {
    if (file.size > manifest.plan.maxFileBytes) {
      oversizedFiles.push({ relativePath: file.relativePath, size: file.size });
      continue;
    }
    const outputPath = resolveOutputPath(
      manifest.plan.destinationRoot,
      file.relativePath
    );
    assertNoSymbolicLinks(manifest.plan.destinationRoot, outputPath);
    if (fs.existsSync(outputPath)) {
      existingFileCount += 1;
      const stat = fs.statSync(outputPath);
      if (!stat.isFile() || stat.size !== file.size) existingSizeConflictCount += 1;
      continue;
    }
    const partialPath = partialOutputPath(outputPath, jobId);
    assertNoSymbolicLinks(manifest.plan.destinationRoot, partialPath);
    const partialSize = fs.existsSync(partialPath) ? fs.statSync(partialPath).size : 0;
    const resumableBytes = partialSize > 0 && partialSize <= file.size ? partialSize : 0;
    pendingDownloadCount += 1;
    pendingDownloadBytes += file.size - resumableBytes;
    if (copyFinalizationRequired(outputPath)) {
      finalizationReserveBytes = Math.max(finalizationReserveBytes, file.size);
    }
  }

  return {
    siteUrl: manifest.plan.siteUrl,
    siteName: manifest.siteName,
    libraryName: manifest.libraryName,
    destinationRoot: manifest.plan.destinationRoot,
    fileCount: manifest.plan.files.length,
    folderCount: manifest.folderCount,
    totalBytes: manifest.plan.files.reduce((sum, file) => sum + file.size, 0),
    pendingDownloadCount,
    pendingDownloadBytes,
    finalizationReserveBytes,
    requiredDiskBytes: pendingDownloadBytes + finalizationReserveBytes,
    existingFileCount,
    existingSizeConflictCount,
    oversizedFileCount: oversizedFiles.length,
    unsupportedItemCount: manifest.unsupported.length,
    renamedPathCount: manifest.renamedPathCount,
    availableDiskBytes: diskSpace(manifest.plan.destinationRoot),
    maxFileBytes: manifest.plan.maxFileBytes,
    oversizedFiles: oversizedFiles.slice(0, 20),
    warning:
      "This exports the default SharePoint document library to the exact destination without overwriting existing files. One approval covers the complete unchanged manifest; files above the per-file cap are skipped."
  };
}

function normalizeSiteUrl(value: string): {
  canonicalUrl: string;
  graphPath: string;
} {
  let parsed: URL;
  try {
    parsed = new URL(value.trim());
  } catch {
    throw new Error("site-url must be a valid HTTPS SharePoint site URL.");
  }
  if (
    parsed.protocol !== "https:" ||
    parsed.username ||
    parsed.password ||
    parsed.port ||
    parsed.search ||
    parsed.hash ||
    parsed.hostname.toLowerCase() !== allowedSharePointHost
  ) {
    throw new Error(
      `site-url must be an HTTPS URL on ${allowedSharePointHost} without credentials, port, query, or fragment.`
    );
  }

  const segments = parsed.pathname
    .split("/")
    .filter(Boolean)
    .map((segment) => decodeURIComponent(segment));
  if (
    segments.length < 2 ||
    !["sites", "teams"].includes(segments[0].toLowerCase()) ||
    segments.some((segment) => segment === "." || segment === "..")
  ) {
    throw new Error("site-url must identify one /sites/... or /teams/... SharePoint site.");
  }
  const canonicalPath = `/${segments.map((segment) => encodeURIComponent(segment)).join("/")}`;
  return {
    canonicalUrl: `https://${allowedSharePointHost}${canonicalPath}`,
    graphPath:
      `/sites/${allowedSharePointHost}:${canonicalPath}` +
      "?$select=id,displayName,webUrl"
  };
}

function normalizeDestination(value: string): string {
  const trimmed = value.trim();
  if (!trimmed || !path.isAbsolute(trimmed)) {
    throw new Error("destination must be an absolute filesystem path.");
  }
  return path.resolve(trimmed);
}

function assertSafeDestinationRoot(destinationRoot: string): void {
  assertNoLinksInExistingChain(destinationRoot);
  const existing = nearestExistingPath(destinationRoot);
  const stat = fs.lstatSync(existing);
  if (!stat.isDirectory() || stat.isSymbolicLink()) {
    throw new Error("destination must resolve through a real directory, not a file or symbolic link.");
  }
  if (fs.existsSync(destinationRoot)) {
    const rootStat = fs.lstatSync(destinationRoot);
    if (!rootStat.isDirectory() || rootStat.isSymbolicLink()) {
      throw new Error("destination must be a real directory.");
    }
  }
}

function assertNoSymbolicLinks(destinationRoot: string, candidate: string): void {
  assertNoLinksInExistingChain(destinationRoot);
  assertNoLinksInExistingChain(candidate);
  const root = path.resolve(destinationRoot);
  const resolved = path.resolve(candidate);
  if (resolved !== root && !resolved.startsWith(`${root}${path.sep}`)) {
    throw new Error("Export path escaped the destination root.");
  }

  const rootParent = nearestExistingPath(root);
  if (fs.lstatSync(rootParent).isSymbolicLink()) {
    throw new Error("Destination path contains a symbolic link or junction.");
  }
  if (!fs.existsSync(root)) return;

  let current = root;
  if (fs.lstatSync(current).isSymbolicLink()) {
    throw new Error("Destination path contains a symbolic link or junction.");
  }
  for (const segment of path.relative(root, resolved).split(path.sep).filter(Boolean)) {
    current = path.join(current, segment);
    if (!fs.existsSync(current)) break;
    if (fs.lstatSync(current).isSymbolicLink()) {
      throw new Error("Destination path contains a symbolic link or junction.");
    }
  }
}

function assertNoLinksInExistingChain(value: string): void {
  const resolved = path.resolve(value);
  const parsed = path.parse(resolved);
  let current = parsed.root;
  for (const segment of resolved.slice(parsed.root.length).split(path.sep).filter(Boolean)) {
    current = path.join(current, segment);
    if (!fs.existsSync(current)) break;
    if (fs.lstatSync(current).isSymbolicLink()) {
      throw new Error("Destination path contains a symbolic link or junction.");
    }
  }
}


function resolveOutputPath(destinationRoot: string, relativePath: string): string {
  const segments = relativePath.split("/");
  if (
    segments.length === 0 ||
    segments.some((segment) => !segment || segment === "." || segment === "..")
  ) {
    throw new Error("Manifest contains an unsafe relative path.");
  }
  const outputPath = path.resolve(destinationRoot, ...segments);
  if (
    outputPath !== destinationRoot &&
    !outputPath.startsWith(`${destinationRoot}${path.sep}`)
  ) {
    throw new Error("Manifest path escaped the destination root.");
  }
  return outputPath;
}

function allocateSegment(
  originalName: string,
  itemId: string,
  usedSegments: Set<string>
): string {
  let candidate = sanitizeFilename(originalName);
  let key = candidate.toLocaleLowerCase("en-US");
  if (!usedSegments.has(key)) {
    usedSegments.add(key);
    return candidate;
  }

  const parsed = path.parse(candidate);
  const suffix = ` [${shortId(itemId)}]`;
  candidate = sanitizeFilename(
    `${parsed.name.slice(0, Math.max(1, 180 - parsed.ext.length - suffix.length))}${suffix}${parsed.ext}`
  );
  key = candidate.toLocaleLowerCase("en-US");
  let index = 2;
  while (usedSegments.has(key)) {
    candidate = sanitizeFilename(
      `${parsed.name.slice(0, 140)}${suffix}-${index}${parsed.ext}`
    );
    key = candidate.toLocaleLowerCase("en-US");
    index += 1;
  }
  usedSegments.add(key);
  return candidate;
}

function validateNextLink(value: string | undefined): string | undefined {
  if (!value) return undefined;
  if (value.startsWith("/")) return value;
  const parsed = new URL(value);
  if (
    parsed.protocol !== "https:" ||
    parsed.hostname !== "graph.microsoft.com" ||
    !parsed.pathname.startsWith("/v1.0/")
  ) {
    throw new Error("Graph pagination returned an unexpected nextLink.");
  }
  return value;
}

function availableDiskBytes(destinationRoot: string): number {
  const existing = nearestExistingPath(destinationRoot);
  const stat = fs.statfsSync(existing, { bigint: true });
  const available = stat.bavail * stat.bsize;
  return available > BigInt(Number.MAX_SAFE_INTEGER)
    ? Number.MAX_SAFE_INTEGER
    : Number(available);
}

function nearestExistingPath(value: string): string {
  let current = path.resolve(value);
  while (!fs.existsSync(current)) {
    const parent = path.dirname(current);
    if (parent === current) throw new Error("destination has no accessible parent directory.");
    current = parent;
  }
  return current;
}

function partialOutputPath(outputPath: string, jobId: string): string {
  return `${outputPath}.hare-part-${jobId}`;
}

function shortId(value: string): string {
  return value.replace(/[^A-Za-z0-9]/g, "").slice(-8) || "item";
}

function normalizeTimeBudget(value: number): number {
  if (!Number.isFinite(value) || value < 1) {
    throw new Error("time-budget-ms must be a positive number.");
  }
  return Math.floor(value);
}

function isSourceChangedError(message: string): boolean {
  return /\b(404|410|412)\b|not found|gone|precondition failed|etag/i.test(message);
}


async function verifySourceItem(
  client: SharePointExportGraphClient,
  driveId: string,
  file: ExportFile
): Promise<void> {
  const current = await client.get<GraphDriveItem>(
    `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(file.itemId)}?$select=id,size,eTag,file`
  );
  if (
    current.id !== file.itemId ||
    !current.file ||
    current.size !== file.size ||
    !current.eTag ||
    current.eTag !== file.eTag
  ) {
    throw new Error("SharePoint source eTag or size changed after approval.");
  }
}

function safeErrorMessage(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  return message.replace(/[\r\n]+/g, " ").slice(0, 500);
}

function defaultClient(config: AppConfig): SharePointExportGraphClient {
  return {
    get: <T>(pathOrUrl: string) => graphGet<T>(config, pathOrUrl),
    download: (pathOrUrl: string, options = {}) => {
      const headers: Record<string, string> = {};
      if (options.rangeStart !== undefined) {
        headers.Range = `bytes=${options.rangeStart}-`;
      }
      if (options.eTag) headers["If-Match"] = options.eTag;
      return graphDownloadResponse(config, pathOrUrl, { headers });
    }
  };
}
