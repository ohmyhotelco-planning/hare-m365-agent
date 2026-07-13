import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import dotenv from "dotenv";

const moduleDir = path.dirname(fileURLToPath(import.meta.url));
const packageRoot = path.resolve(moduleDir, "..");

dotenv.config({ path: path.resolve(".env") });
dotenv.config({ path: path.join(packageRoot, ".env") });

type HareDefaultConfig = {
  clientId?: string;
  tenantId?: string;
  policyPath?: string;
};

export type Policy = {
  defaultMode: "readOnly" | "readWrite";
  allowWriteActions: boolean;
  allowDownloads: boolean;
  storeRawMessages: boolean;
  defaultSearchLookbackDays: number;
  maxSearchResults: number;
  maxMailFetchLimit: number;
  maxTeamsFetchLimit: number;
  maxFileSearchLimit: number;
  maxDownloadBytes: number;
  retentionDays: number;
  requireConfirmationFor: string[];
};

export type AppConfig = {
  clientId: string;
  tenantId: string;
  authority: string;
  dataDir: string;
  cacheDir: string;
  downloadDir: string;
  logsDir: string;
  resultsDir: string;
  timeZone: string;
  policyPath: string;
  policy: Policy;
};

const defaultPolicy: Policy = {
  defaultMode: "readOnly",
  allowWriteActions: false,
  allowDownloads: true,
  storeRawMessages: false,
  defaultSearchLookbackDays: 90,
  maxSearchResults: 1000,
  maxMailFetchLimit: 20,
  maxTeamsFetchLimit: 50,
  maxFileSearchLimit: 25,
  maxDownloadBytes: 104857600,
  retentionDays: 7,
  requireConfirmationFor: [
    "send_mail",
    "post_teams_message",
    "create_calendar_event",
    "upload_file",
    "delete_file",
    "change_permission"
  ]
};

function readJson<T>(filePath: string, fallback: T): T {
  if (!fs.existsSync(filePath)) return fallback;
  return JSON.parse(fs.readFileSync(filePath, "utf8")) as T;
}

function readDefaultConfig(): HareDefaultConfig {
  return readJson<HareDefaultConfig>(path.join(packageRoot, "hare.config.json"), {});
}

function defaultDataDir(): string {
  const explicitDataDir = process.env.HARE_M365_DATA_DIR ?? process.env.OMH_M365_DATA_DIR;
  if (explicitDataDir) {
    return explicitDataDir;
  }

  if (process.platform === "win32") {
    return path.join(process.env.LOCALAPPDATA ?? os.homedir(), "Ohmyhotel", "HareM365Agent");
  }

  if (process.platform === "darwin") {
    return path.join(os.homedir(), "Library", "Application Support", "Ohmyhotel", "HareM365Agent");
  }

  return path.join(process.env.XDG_DATA_HOME ?? path.join(os.homedir(), ".local", "share"), "ohmyhotel", "hare-m365-agent");
}

function resolvePackagePath(value: string | undefined, fallback: string): string {
  if (!value) return path.resolve(fallback);
  return path.resolve(path.isAbsolute(value) ? value : path.join(packageRoot, value));
}

export function loadConfig(): AppConfig {
  const defaults = readDefaultConfig();
  const clientId = process.env.OMH_M365_CLIENT_ID ?? defaults.clientId ?? "";
  const tenantId = process.env.OMH_M365_TENANT_ID ?? defaults.tenantId ?? "";
  const dataDir = path.resolve(defaultDataDir());
  const policyPath = resolvePackagePath(process.env.OMH_M365_POLICY_PATH ?? defaults.policyPath, path.join(packageRoot, "policy.json"));
  const cacheDir = path.join(dataDir, ".cache");
  const downloadDir = path.join(dataDir, "downloads");
  const logsDir = path.join(dataDir, "logs");
  const resultsDir = path.join(dataDir, "results");
  const timeZone = process.env.HARE_M365_TIME_ZONE ?? "Asia/Seoul";

  const policy = {
    ...defaultPolicy,
    ...readJson<Partial<Policy>>(policyPath, {})
  };

  return {
    clientId,
    tenantId,
    authority: tenantId ? `https://login.microsoftonline.com/${tenantId}` : "",
    dataDir,
    cacheDir,
    downloadDir,
    logsDir,
    resultsDir,
    timeZone,
    policyPath,
    policy
  };
}

export function ensureRuntimeDirs(config: AppConfig): void {
  for (const dir of [
    config.dataDir,
    config.cacheDir,
    config.downloadDir,
    config.logsDir,
    config.resultsDir
  ]) {
    fs.mkdirSync(path.resolve(dir), { recursive: true });
  }
}

export function requireConfigured(config: AppConfig): void {
  const missing: string[] = [];
  if (!config.clientId) missing.push("OMH_M365_CLIENT_ID");
  if (!config.tenantId) missing.push("OMH_M365_TENANT_ID");
  if (missing.length > 0) {
    throw new Error(
      `Missing required configuration values: ${missing.join(", ")}. Check hare.config.json or set local environment overrides.`
    );
  }
}
