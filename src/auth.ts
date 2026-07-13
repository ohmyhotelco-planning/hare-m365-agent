import fs from "node:fs";
import path from "node:path";
import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration
} from "@azure/msal-node";
import type { AppConfig } from "./config.js";
import { ProxyAwareNetworkClient } from "./msal-network.js";

const scopes = [
  "User.Read",
  "Mail.Read",
  "Chat.Read",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "Files.Read.All",
  "Sites.Read.All",
  "offline_access"
];

const cacheLockRetryMs = 50;
const cacheLockTimeoutMs = 10_000;
const staleLockMs = 60_000;

function cachePath(config: AppConfig): string {
  return path.join(config.cacheDir, "msal-cache.json");
}

async function buildPca(config: AppConfig): Promise<PublicClientApplication> {
  let releaseCacheLock: (() => void) | undefined;
  const msalConfig: Configuration = {
    auth: {
      clientId: config.clientId,
      authority: config.authority
    },
    cache: {
      cachePlugin: {
        beforeCacheAccess: async (cacheContext) => {
          const file = cachePath(config);
          fs.mkdirSync(config.cacheDir, { recursive: true });
          releaseCacheLock = await acquireFileLock(`${file}.lock`);
          if (fs.existsSync(file)) {
            try {
              cacheContext.tokenCache.deserialize(fs.readFileSync(file, "utf8"));
            } catch (error) {
              releaseCacheLock?.();
              releaseCacheLock = undefined;
              throw new Error(`Hare login cache is unreadable: ${errorMessage(error)}`);
            }
          }
        },
        afterCacheAccess: async (cacheContext) => {
          try {
            if (cacheContext.cacheHasChanged) {
              writeFileAtomically(cachePath(config), cacheContext.tokenCache.serialize());
            }
          } finally {
            releaseCacheLock?.();
            releaseCacheLock = undefined;
          }
        }
      }
    },
    system: {
      networkClient: new ProxyAwareNetworkClient()
    }
  };

  return new PublicClientApplication(msalConfig);
}

async function acquireFileLock(lockPath: string): Promise<() => void> {
  const deadline = Date.now() + cacheLockTimeoutMs;

  while (true) {
    try {
      const handle = fs.openSync(lockPath, "wx", 0o600);
      fs.writeFileSync(handle, `${process.pid}\n`, "utf8");
      return () => {
        try {
          fs.closeSync(handle);
        } finally {
          fs.rmSync(lockPath, { force: true });
        }
      };
    } catch (error) {
      const code = (error as NodeJS.ErrnoException).code;
      if (code !== "EEXIST") throw error;

      try {
        const lockAge = Date.now() - fs.statSync(lockPath).mtimeMs;
        const lockOwner = Number.parseInt(fs.readFileSync(lockPath, "utf8").trim(), 10);
        if (lockAge > staleLockMs && !isProcessRunning(lockOwner)) {
          fs.rmSync(lockPath, { force: true });
          continue;
        }
      } catch (statError) {
        if ((statError as NodeJS.ErrnoException).code === "ENOENT") continue;
        throw statError;
      }

      if (Date.now() >= deadline) {
        throw new Error("Timed out waiting for the Hare login cache lock.");
      }
      await new Promise((resolve) => setTimeout(resolve, cacheLockRetryMs));
    }
  }
}

function isProcessRunning(pid: number): boolean {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === "EPERM";
  }
}

function writeFileAtomically(filePath: string, contents: string): void {
  const temporaryPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  try {
    fs.writeFileSync(temporaryPath, contents, { encoding: "utf8", mode: 0o600 });
    fs.renameSync(temporaryPath, filePath);
  } finally {
    fs.rmSync(temporaryPath, { force: true });
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export async function login(config: AppConfig): Promise<AuthenticationResult> {
  const pca = await buildPca(config);
  const result = await pca.acquireTokenByDeviceCode({
    scopes,
    deviceCodeCallback: (response) => {
      console.log(response.message);
    }
  });

  if (!result) throw new Error("Login failed: Microsoft returned no authentication result.");
  return result;
}

export async function getAccount(config: AppConfig): Promise<AccountInfo | null> {
  const pca = await buildPca(config);
  const accounts = await pca.getTokenCache().getAllAccounts();
  return accounts[0] ?? null;
}

export type AuthStatus = {
  account: AccountInfo | null;
  loggedIn: boolean;
  tokenUsable: boolean;
  reason?: string;
};

export async function getAuthStatus(config: AppConfig): Promise<AuthStatus> {
  const pca = await buildPca(config);
  const accounts = await pca.getTokenCache().getAllAccounts();
  const account = accounts[0] ?? null;
  if (!account) {
    return { account: null, loggedIn: false, tokenUsable: false, reason: "NO_ACCOUNT_IN_CACHE" };
  }

  try {
    const result = await pca.acquireTokenSilent({ account, scopes });
    const tokenUsable = Boolean(result?.accessToken);
    return {
      account,
      loggedIn: tokenUsable,
      tokenUsable,
      reason: tokenUsable ? undefined : "NO_ACCESS_TOKEN"
    };
  } catch (error) {
    return {
      account,
      loggedIn: false,
      tokenUsable: false,
      reason: `TOKEN_ACQUISITION_FAILED: ${errorMessage(error)}`
    };
  }
}

export async function getAccessToken(config: AppConfig): Promise<string> {
  const pca = await buildPca(config);
  const accounts = await pca.getTokenCache().getAllAccounts();
  const account = accounts[0] ?? null;
  if (!account) {
    throw new Error(
      "Not logged in for this Hare dataDir/cacheFile. During initial connection, run Hare auth login in the same shell, let the user complete Microsoft device-code login, then retry in the same dataDir/cacheFile."
    );
  }

  const result = await pca.acquireTokenSilent({
    account,
    scopes
  });

  if (!result?.accessToken) throw new Error("Could not acquire access token.");
  return result.accessToken;
}

export function logout(config: AppConfig): void {
  const file = cachePath(config);
  if (fs.existsSync(file)) fs.rmSync(file, { force: true });
}

export function getScopeList(): string[] {
  return [...scopes];
}
