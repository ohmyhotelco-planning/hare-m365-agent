import fs from "node:fs";
import path from "node:path";
import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration,
  type INetworkModule
} from "@azure/msal-node";
import type { AppConfig } from "./config.js";
import {
  clearDeviceLoginState,
  readDeviceLoginState,
  ResumeDeviceCodeNetworkClient,
  startDeviceLogin,
  type DeviceLoginStartResult
} from "./device-login.js";
import { ProxyAwareNetworkClient } from "./msal-network.js";
import {
  clearStoredFile,
  storedFileHasContent,
  usesDeleteRestrictedStorage,
  writeStoredText
} from "./persistent-storage.js";

const scopes = [
  "User.Read",
  "Mail.Read",
  "Chat.Read",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "Files.Read.All",
  "Sites.Read.All",
  "openid",
  "profile",
  "offline_access"
];

const cacheLockRetryMs = 50;
const cacheLockTimeoutMs = 10_000;
const staleLockMs = 60_000;

function cachePath(config: AppConfig): string {
  return path.join(config.cacheDir, "msal-cache.json");
}

async function buildPca(
  config: AppConfig,
  networkClient: INetworkModule = new ProxyAwareNetworkClient()
): Promise<PublicClientApplication> {
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
          releaseCacheLock = usesDeleteRestrictedStorage(file)
            ? undefined
            : await acquireFileLock(`${file}.lock`);
          if (storedFileHasContent(file)) {
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
              writeStoredText(cachePath(config), cacheContext.tokenCache.serialize());
            }
          } finally {
            releaseCacheLock?.();
            releaseCacheLock = undefined;
          }
        }
      }
    },
    system: {
      networkClient
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
      const windowsLockContention =
        (code === "EPERM" || code === "EACCES") && fs.existsSync(lockPath);
      if (code !== "EEXIST" && !windowsLockContention) throw error;

      try {
        const lockAge = Date.now() - fs.statSync(lockPath).mtimeMs;
        const lockOwner = Number.parseInt(fs.readFileSync(lockPath, "utf8").trim(), 10);
        if (lockAge > staleLockMs && !isProcessRunning(lockOwner)) {
          fs.rmSync(lockPath, { force: true });
          continue;
        }
      } catch (statError) {
        const statCode = (statError as NodeJS.ErrnoException).code;
        if (statCode === "ENOENT") continue;
        if (statCode === "EPERM" || statCode === "EACCES") {
          if (Date.now() >= deadline) {
            throw new Error("Timed out waiting for the Hare login cache lock.");
          }
          await new Promise((resolve) => setTimeout(resolve, cacheLockRetryMs));
          continue;
        }
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

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export async function startLogin(config: AppConfig): Promise<DeviceLoginStartResult> {
  return startDeviceLogin(config, scopes);
}

export async function completeLogin(
  config: AppConfig,
  networkClient: INetworkModule = new ProxyAwareNetworkClient()
): Promise<AuthenticationResult> {
  const state = readDeviceLoginState(config);
  if ([...state.scopes].sort().join(" ") !== [...scopes].sort().join(" ")) {
    clearDeviceLoginState(config);
    throw new Error("Pending login scopes changed. Run auth login-start again.");
  }
  const pca = await buildPca(config, new ResumeDeviceCodeNetworkClient(state, networkClient));
  let result: AuthenticationResult | null;
  try {
    result = await pca.acquireTokenByDeviceCode({
      scopes: state.scopes,
      deviceCodeCallback: () => undefined,
      timeout: 25
    });
  } catch (error) {
    const message = errorMessage(error);
    const diagnostic = `${(error as { errorCode?: string }).errorCode ?? ""} ${message}`;
    if (/expired/i.test(diagnostic)) {
      clearDeviceLoginState(config);
      throw new Error("LOGIN_CODE_EXPIRED: Run auth login-start again.");
    }
    if (/polling_cancelled|authorization_pending|timeout/i.test(diagnostic)) {
      throw new Error(
        "LOGIN_PENDING: Microsoft sign-in is not complete yet. Keep the existing code, finish the browser sign-in, then run auth login-complete again."
      );
    }
    throw error;
  }

  if (!result) throw new Error("Login failed: Microsoft returned no authentication result.");
  if (!result.account || !result.accessToken) {
    clearDeviceLoginState(config);
    throw new Error(
      "LOGIN_CACHE_VERIFICATION_FAILED: Microsoft sign-in returned without a usable account or access token. Run auth login-start again."
    );
  }

  const verifiedStatus = await getAuthStatus(config);
  if (!verifiedStatus.loggedIn || !verifiedStatus.tokenUsable) {
    clearDeviceLoginState(config);
    throw new Error(
      `LOGIN_CACHE_VERIFICATION_FAILED: ${verifiedStatus.reason ?? "Persisted token is unavailable"}. Run auth login-start again.`
    );
  }
  clearDeviceLoginState(config);
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
      "Not logged in for this Hare dataDir/cacheFile. Run auth login-start, let the user finish Microsoft sign-in, then run auth login-complete and retry in the same dataDir."
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
  clearStoredFile(file);
  clearDeviceLoginState(config);
}

export function getScopeList(): string[] {
  return [...scopes];
}
