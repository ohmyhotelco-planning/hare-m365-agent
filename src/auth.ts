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
  markAuthProfileReady,
  prepareAuthProfile,
  resetAuthProfileAfterLogout
} from "./auth-profile.js";
import {
  clearDeviceLoginState,
  readDeviceLoginState,
  ResumeDeviceCodeNetworkClient,
  startDeviceLogin,
  type DeviceLoginStartResult
} from "./device-login.js";
import { ProxyAwareNetworkClient, type AuthNetworkFailure } from "./msal-network.js";
import { classifyAuthFailure } from "./setup-state.js";
import {
  clearStoredFile,
  storedFileHasContent,
  usesDeleteRestrictedStorage,
  writeStoredText
} from "./persistent-storage.js";

const scopes = [
  "User.Read",
  "User.ReadBasic.All",
  "Mail.ReadWrite",
  "Mail.Read.Shared",
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
  prepareAuthProfile(config, scopes);
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

  markAuthProfileReady(config, scopes);
  // The code was redeemed successfully; token verification must never resume it again.
  clearDeviceLoginState(config);
  const verifiedStatus = await getAuthStatus(config, networkClient);
  if (verifiedStatus.reason?.startsWith("AUTH_CHECK_BLOCKED:")) {
    throw new Error(verifiedStatus.reason);
  }
  if (!verifiedStatus.loggedIn || !verifiedStatus.tokenUsable) {
    throw new Error(
      `LOGIN_CACHE_VERIFICATION_FAILED: ${verifiedStatus.reason ?? "Persisted token is unavailable"}. Run auth login-start again.`
    );
  }
  return result;
}

export async function getAccount(config: AppConfig): Promise<AccountInfo | null> {
  if (prepareAuthProfile(config, scopes).migrationRequired) return null;
  const pca = await buildPca(config);
  const accounts = await pca.getTokenCache().getAllAccounts();
  return accounts[0] ?? null;
}

export type AuthStatus = {
  account: AccountInfo | null;
  // null means connectivity prevented verification, not that sign-in was rejected.
  loggedIn: boolean | null;
  tokenUsable: boolean | null;
  migrationRequired: boolean;
  reason?: string;
  networkFailure?: AuthNetworkFailure;
};

function networkBlockedReason(error: unknown, failure?: AuthNetworkFailure): string | undefined {
  const code = (error as { errorCode?: unknown } | null)?.errorCode;
  if (typeof code === "string" && /^(interaction_required|invalid_grant|login_required|consent_required)$/i.test(code)) {
    return undefined;
  }
  if (!failure && classifyAuthFailure(`TOKEN_ACQUISITION_FAILED: ${typeof code === "string" ? code : ""} ${errorMessage(error)}`) !== "NETWORK_BLOCKED") {
    return undefined;
  }
  const detail = failure ? `${failure.code} ${failure.method} ${failure.hostname} ${failure.stage}` : "network_error";
  return `AUTH_CHECK_BLOCKED: ${detail}. Token validity is unknown because Microsoft connectivity could not be verified. Keep the existing cache; do not sign in again or reset it. Retry the same cache only after connectivity is restored in an approved execution environment.`;
}

function tokenFailureReason(error: unknown): string {
  const code = (error as { errorCode?: unknown } | null)?.errorCode;
  const diagnostic = `${typeof code === "string" ? code : ""} ${errorMessage(error)}`;
  const knownCode = diagnostic.match(/\b(interaction_required|invalid_grant|login_required|consent_required)\b/i)?.[1];
  return `TOKEN_ACQUISITION_FAILED: ${knownCode?.toLowerCase() ?? "unclassified_error"}`;
}

export async function getAuthStatus(
  config: AppConfig,
  networkClient: INetworkModule = new ProxyAwareNetworkClient()
): Promise<AuthStatus> {
  if (networkClient instanceof ProxyAwareNetworkClient) networkClient.takeNetworkFailure();
  const profile = prepareAuthProfile(config, scopes);
  if (profile.migrationRequired) {
    return {
      account: null,
      loggedIn: false,
      tokenUsable: false,
      migrationRequired: true,
      reason: "AUTH_APP_CHANGED"
    };
  }

  const pca = await buildPca(config, networkClient);
  const accounts = await pca.getTokenCache().getAllAccounts();
  const account = accounts[0] ?? null;
  if (!account) {
    return {
      account: null,
      loggedIn: false,
      tokenUsable: false,
      migrationRequired: false,
      reason: "NO_ACCOUNT_IN_CACHE"
    };
  }

  try {
    const result = await pca.acquireTokenSilent({ account, scopes });
    const tokenUsable = Boolean(result?.accessToken);
    return {
      account,
      loggedIn: tokenUsable,
      tokenUsable,
      migrationRequired: false,
      reason: tokenUsable ? undefined : "NO_ACCESS_TOKEN"
    };
  } catch (error) {
    const networkFailure = networkClient instanceof ProxyAwareNetworkClient
      ? networkClient.takeNetworkFailure() : undefined;
    const blockedReason = networkBlockedReason(error, networkFailure);
    return {
      account,
      loggedIn: blockedReason ? null : false,
      tokenUsable: blockedReason ? null : false,
      migrationRequired: false,
      reason: blockedReason ?? tokenFailureReason(error),
      networkFailure: blockedReason ? networkFailure : undefined
    };
  }
}

export async function getAccessToken(
  config: AppConfig,
  networkClient: INetworkModule = new ProxyAwareNetworkClient()
): Promise<string> {
  if (networkClient instanceof ProxyAwareNetworkClient) networkClient.takeNetworkFailure();
  if (prepareAuthProfile(config, scopes).migrationRequired) {
    throw new Error(
      "Hare M365 Agent authentication permissions or application changed. Complete Microsoft sign-in once, then retry."
    );
  }
  const pca = await buildPca(config, networkClient);
  const accounts = await pca.getTokenCache().getAllAccounts();
  const account = accounts[0] ?? null;
  if (!account) {
    throw new Error(
      "Not logged in for this Hare dataDir/cacheFile. Run auth login-start, let the user finish Microsoft sign-in, then run auth login-complete and retry in the same dataDir."
    );
  }

  let result: AuthenticationResult | null;
  try {
    result = await pca.acquireTokenSilent({ account, scopes });
  } catch (error) {
    const failure = networkClient instanceof ProxyAwareNetworkClient
      ? networkClient.takeNetworkFailure() : undefined;
    const reason = networkBlockedReason(error, failure);
    if (reason) throw new Error(reason);
    throw error;
  }

  if (!result?.accessToken) throw new Error("Could not acquire access token.");
  return result.accessToken;
}

export function logout(config: AppConfig): void {
  const file = cachePath(config);
  clearStoredFile(file);
  clearDeviceLoginState(config);
  resetAuthProfileAfterLogout(config, scopes);
}

export function getScopeList(): string[] {
  return [...scopes];
}
