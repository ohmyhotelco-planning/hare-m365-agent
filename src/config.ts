import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import dotenv from "dotenv";

const moduleDir = path.dirname(fileURLToPath(import.meta.url));
const packageRoot = path.resolve(moduleDir, "..");

dotenv.config({ path: path.resolve(".env") });
dotenv.config({ path: path.join(packageRoot, ".env") });

export type Policy = {
  defaultMode: "readOnly" | "readWrite";
  allowWriteActions: boolean;
  allowDownloads: boolean;
  storeRawMessages: boolean;
  maxMailFetchLimit: number;
  maxTeamsFetchLimit: number;
  maxFileSearchLimit: number;
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
  policyPath: string;
  policy: Policy;
};

const defaultPolicy: Policy = {
  defaultMode: "readOnly",
  allowWriteActions: false,
  allowDownloads: true,
  storeRawMessages: false,
  maxMailFetchLimit: 20,
  maxTeamsFetchLimit: 50,
  maxFileSearchLimit: 25,
  retentionDays: 0,
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

function defaultDataDir(): string {
  if (process.platform === "win32") {
    return path.join(process.env.LOCALAPPDATA ?? os.homedir(), "Ohmyhotel", "HareM365Agent");
  }

  if (process.platform === "darwin") {
    return path.join(os.homedir(), "Library", "Application Support", "Ohmyhotel", "HareM365Agent");
  }

  return path.join(process.env.XDG_DATA_HOME ?? path.join(os.homedir(), ".local", "share"), "ohmyhotel", "hare-m365-agent");
}

function resolveStoragePath(value: string | undefined, fallback: string, baseDir: string): string {
  if (!value) return path.resolve(fallback);
  return path.resolve(path.isAbsolute(value) ? value : path.join(baseDir, value));
}

function resolvePackagePath(value: string | undefined, fallback: string): string {
  if (!value) return path.resolve(fallback);
  return path.resolve(path.isAbsolute(value) ? value : path.join(packageRoot, value));
}

export function loadConfig(): AppConfig {
  const clientId = process.env.OMH_M365_CLIENT_ID ?? "";
  const tenantId = process.env.OMH_M365_TENANT_ID ?? "";
  const dataDir = path.resolve(process.env.HARE_M365_DATA_DIR ?? process.env.OMH_M365_DATA_DIR ?? defaultDataDir());
  const policyPath = resolvePackagePath(process.env.OMH_M365_POLICY_PATH, path.join(packageRoot, "policy.json"));
  const cacheDir = resolveStoragePath(process.env.OMH_M365_CACHE_DIR, path.join(dataDir, ".cache"), dataDir);
  const downloadDir = resolveStoragePath(process.env.OMH_M365_DOWNLOAD_DIR, path.join(dataDir, "downloads"), dataDir);
  const logsDir = resolveStoragePath(process.env.OMH_M365_LOGS_DIR, path.join(dataDir, "logs"), dataDir);

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
    policyPath,
    policy
  };
}

export function ensureRuntimeDirs(config: AppConfig): void {
  for (const dir of [config.dataDir, config.cacheDir, config.downloadDir, config.logsDir]) {
    fs.mkdirSync(path.resolve(dir), { recursive: true });
  }
}

export function requireConfigured(config: AppConfig): void {
  const missing: string[] = [];
  if (!config.clientId) missing.push("OMH_M365_CLIENT_ID");
  if (!config.tenantId) missing.push("OMH_M365_TENANT_ID");
  if (missing.length > 0) {
    throw new Error(
      `Missing required environment values: ${missing.join(", ")}. Use the approved Hare M365 Agent package configuration or set these values in a local .env file.`
    );
  }
}
