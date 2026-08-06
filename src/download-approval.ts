import { createHash, randomBytes } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";
import { clearStoredFile, writeStoredText } from "./persistent-storage.js";

export type DownloadApprovalPlan = {
  source: "sharepoint" | "outlook";
  resourceKey: string;
  name: string;
  size: number;
  outputName: string;
};

export type PendingDownloadApproval = {
  stage: "AWAITING_USER_APPROVAL";
  preview: {
    source: DownloadApprovalPlan["source"];
    name: string;
    size: number;
    sizeMiB: number;
    outputName: string;
    downloadDir: string;
    defaultLimitBytes: number;
    approvedLimitBytes: number;
    warning: string;
  };
  approval: {
    token: string;
    expiresAt: string;
    instruction: string;
  };
};

export type DownloadAuthorization =
  | { approved: true; maxBytes: number; elevated: boolean }
  | { approved: false; pending: PendingDownloadApproval };

type StoredApproval = {
  tokenHash: string;
  planHash: string;
  expiresAtMs: number;
};

type StoredApprovalFile = {
  approvals: StoredApproval[];
};

const approvalTtlMs = 10 * 60 * 1000;
const maxPendingApprovals = 20;

export function authorizeDownload(
  config: AppConfig,
  plan: DownloadApprovalPlan,
  approvalToken?: string,
  now = Date.now()
): DownloadAuthorization {
  validatePlan(plan);
  const defaultLimit = config.policy.maxDownloadBytes;
  const approvedLimit = config.policy.maxApprovedDownloadBytes;
  if (
    !Number.isFinite(defaultLimit) ||
    defaultLimit < 1 ||
    !Number.isFinite(approvedLimit) ||
    approvedLimit < defaultLimit
  ) {
    throw new Error(
      "Download policy is invalid: maxApprovedDownloadBytes must be at least maxDownloadBytes."
    );
  }

  if (plan.size <= defaultLimit) {
    return { approved: true, maxBytes: defaultLimit, elevated: false };
  }
  if (plan.size > approvedLimit) {
    throw new Error(
      `Download exceeds approved policy limit of ${approvedLimit} bytes.`
    );
  }

  const approvalFile = downloadApprovalFile(config);
  const activeApprovals = readApprovals(approvalFile).filter(
    (approval) => approval.expiresAtMs > now
  );
  const planHash = hashPlan(plan);

  if (!approvalToken) {
    const token = randomBytes(24).toString("hex");
    const expiresAtMs = now + approvalTtlMs;
    activeApprovals.push({
      tokenHash: hashToken(token),
      planHash,
      expiresAtMs
    });
    writeApprovals(approvalFile, activeApprovals.slice(-maxPendingApprovals));
    return {
      approved: false,
      pending: {
        stage: "AWAITING_USER_APPROVAL",
        preview: {
          source: plan.source,
          name: plan.name,
          size: plan.size,
          sizeMiB: Number((plan.size / 1024 / 1024).toFixed(1)),
          outputName: plan.outputName,
          downloadDir: config.downloadDir,
          defaultLimitBytes: defaultLimit,
          approvedLimitBytes: approvedLimit,
          warning: "This file exceeds the default download limit. Download it only after the user explicitly approves this exact preview."
        },
        approval: {
          token,
          expiresAt: new Date(expiresAtMs).toISOString(),
          instruction:
            "Show the complete preview to the user and stop. After explicit approval, rerun the exact same command once with --approval-token set to this token."
        }
      }
    };
  }

  const tokenHash = hashToken(approvalToken);
  const matched = activeApprovals.find((approval) => approval.tokenHash === tokenHash);
  writeApprovals(
    approvalFile,
    activeApprovals.filter((approval) => approval.tokenHash !== tokenHash)
  );
  if (!matched) {
    throw new Error(
      "DOWNLOAD_APPROVAL_REQUIRED: The approval token is missing, expired, or already used. Show a new preview and obtain approval again."
    );
  }
  if (matched.planHash !== planHash) {
    throw new Error(
      "DOWNLOAD_APPROVAL_REQUIRED: The approved file does not match this download. Show a new preview and obtain approval again."
    );
  }

  return { approved: true, maxBytes: approvedLimit, elevated: true };
}

function validatePlan(plan: DownloadApprovalPlan): void {
  if (!Number.isFinite(plan.size) || plan.size < 0) {
    throw new Error("Download size metadata is required for policy approval.");
  }
  if (!plan.resourceKey || !plan.name || !plan.outputName) {
    throw new Error("Download approval plan is incomplete.");
  }
}

function downloadApprovalFile(config: AppConfig): string {
  return path.join(config.cacheDir, "download-approvals.json");
}

function readApprovals(filePath: string): StoredApproval[] {
  if (!fs.existsSync(filePath) || fs.statSync(filePath).size === 0) return [];
  try {
    const parsed = JSON.parse(fs.readFileSync(filePath, "utf8")) as StoredApprovalFile;
    return Array.isArray(parsed.approvals) ? parsed.approvals : [];
  } catch {
    return [];
  }
}

function writeApprovals(filePath: string, approvals: StoredApproval[]): void {
  if (approvals.length === 0) {
    clearStoredFile(filePath);
    return;
  }
  writeStoredText(filePath, JSON.stringify({ approvals } satisfies StoredApprovalFile));
}

function hashPlan(plan: DownloadApprovalPlan): string {
  return createHash("sha256").update(JSON.stringify(plan)).digest("hex");
}

function hashToken(token: string): string {
  return createHash("sha256").update(token).digest("hex");
}
