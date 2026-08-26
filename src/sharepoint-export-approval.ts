import { createHash, randomBytes } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";
import { clearStoredFile, writeStoredText } from "./persistent-storage.js";

export type SharePointExportPlan = {
  schemaVersion: 1;
  tenantId: string;
  accountBindingHash: string;
  siteUrl: string;
  siteId: string;
  driveId: string;
  destinationRoot: string;
  maxFileBytes: number;
  files: Array<{
    itemId: string;
    relativePath: string;
    size: number;
    eTag?: string;
  }>;
};

export type ExportAuthorization =
  | { approved: true; resumed: boolean; jobId: string }
  | {
      approved: false;
      jobId: string;
      token: string;
      expiresAt: string;
    };

type StoredPendingApproval = {
  tokenHash: string;
  planHash: string;
  expiresAtMs: number;
};

type StoredActiveApproval = {
  planHash: string;
  expiresAtMs: number;
};

type StoredApprovalFile = {
  pending: StoredPendingApproval[];
  active: StoredActiveApproval[];
};

const pendingApprovalTtlMs = 10 * 60 * 1000;
const activeApprovalTtlMs = 24 * 60 * 60 * 1000;
const maxStoredApprovals = 20;

export function authorizeSharePointExport(
  config: AppConfig,
  plan: SharePointExportPlan,
  approvalToken?: string,
  now = Date.now()
): ExportAuthorization {
  validatePlan(plan);
  const planHash = hashExportPlan(plan);
  const jobId = planHash.slice(0, 16);
  const filePath = approvalFile(config);
  const state = readState(filePath);
  state.pending = state.pending.filter((entry) => entry.expiresAtMs > now);
  state.active = state.active.filter((entry) => entry.expiresAtMs > now);

  if (state.active.some((entry) => entry.planHash === planHash)) {
    writeState(filePath, state);
    return { approved: true, resumed: true, jobId };
  }

  if (!approvalToken) {
    const token = randomBytes(24).toString("hex");
    const expiresAtMs = now + pendingApprovalTtlMs;
    state.pending.push({
      tokenHash: hashToken(token),
      planHash,
      expiresAtMs
    });
    state.pending = state.pending.slice(-maxStoredApprovals);
    writeState(filePath, state);
    return {
      approved: false,
      jobId,
      token,
      expiresAt: new Date(expiresAtMs).toISOString()
    };
  }

  const tokenHash = hashToken(approvalToken);
  const matched = state.pending.find((entry) => entry.tokenHash === tokenHash);
  state.pending = state.pending.filter((entry) => entry.tokenHash !== tokenHash);
  if (!matched) {
    writeState(filePath, state);
    throw new Error(
      "EXPORT_APPROVAL_REQUIRED: The approval token is missing, expired, or already used. Show a new export preview and obtain approval again."
    );
  }
  if (matched.planHash !== planHash) {
    writeState(filePath, state);
    throw new Error(
      "EXPORT_APPROVAL_REQUIRED: The approved export plan no longer matches the current site, files, policy, or destination."
    );
  }

  state.active.push({
    planHash,
    expiresAtMs: now + activeApprovalTtlMs
  });
  state.active = state.active.slice(-maxStoredApprovals);
  writeState(filePath, state);
  return { approved: true, resumed: false, jobId };
}

export function completeSharePointExportApproval(
  config: AppConfig,
  plan: SharePointExportPlan
): void {
  const filePath = approvalFile(config);
  const state = readState(filePath);
  const planHash = hashExportPlan(plan);
  state.active = state.active.filter((entry) => entry.planHash !== planHash);
  writeState(filePath, state);
}

export function hashExportPlan(plan: SharePointExportPlan): string {
  return createHash("sha256").update(JSON.stringify(plan)).digest("hex");
}

function validatePlan(plan: SharePointExportPlan): void {
  if (
    plan.schemaVersion !== 1 ||
    !plan.tenantId ||
    !plan.accountBindingHash ||
    !plan.siteUrl ||
    !plan.siteId ||
    !plan.driveId ||
    !path.isAbsolute(plan.destinationRoot) ||
    !Number.isFinite(plan.maxFileBytes) ||
    plan.maxFileBytes < 1
  ) {
    throw new Error("SharePoint export approval plan is incomplete.");
  }
  for (const file of plan.files) {
    if (
      !file.itemId ||
      !file.relativePath ||
      !Number.isFinite(file.size) ||
      file.size < 0
    ) {
      throw new Error("SharePoint export approval plan contains an invalid file.");
    }
  }
}

function approvalFile(config: AppConfig): string {
  return path.join(config.cacheDir, "sharepoint-export-approvals.json");
}

function readState(filePath: string): StoredApprovalFile {
  if (!fs.existsSync(filePath) || fs.statSync(filePath).size === 0) {
    return { pending: [], active: [] };
  }
  try {
    const parsed = JSON.parse(fs.readFileSync(filePath, "utf8")) as Partial<StoredApprovalFile>;
    return {
      pending: Array.isArray(parsed.pending) ? parsed.pending : [],
      active: Array.isArray(parsed.active) ? parsed.active : []
    };
  } catch {
    return { pending: [], active: [] };
  }
}

function writeState(filePath: string, state: StoredApprovalFile): void {
  if (state.pending.length === 0 && state.active.length === 0) {
    clearStoredFile(filePath);
    return;
  }
  writeStoredText(filePath, JSON.stringify(state));
}

function hashToken(token: string): string {
  return createHash("sha256").update(token).digest("hex");
}
