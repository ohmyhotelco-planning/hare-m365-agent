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
  getAuthProfilePath,
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
): Promise<{ pca: PublicClientApplication; commit: (homeAccountId?: string) => Promise<void> }> {
  let releaseCacheLock: (() => void) | undefined;
  let pendingCache: string | undefined;
  const readCache = () => storedFileHasContent(cachePath(config))
    ? fs.readFileSync(cachePath(config), "utf8") : "";
  const readProfile = () => {
    try {
      return fs.readFileSync(getAuthProfilePath(config), "utf8");
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") return undefined;
      throw new Error("AUTH_PROFILE_UNREADABLE");
    }
  };
  // Device-code MSAL hooks run after the response; capture state before the request.
  let initialCache: string;
  try {
    initialCache = readCache();
  } catch {
    throw new Error("AUTH_CACHE_UNREADABLE");
  }
  const initialProfile = readProfile();
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
          try {
            const text = pendingCache ?? initialCache;
            if (text) cacheContext.tokenCache.deserialize(text);
          } catch {
            releaseCacheLock?.();
            releaseCacheLock = undefined;
            throw new Error("AUTH_CACHE_UNREADABLE");
          }
        },
        afterCacheAccess: async (cacheContext) => {
          try {
            if (cacheContext.cacheHasChanged) {
              pendingCache = cacheContext.tokenCache.serialize();
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

  return {
    pca: new PublicClientApplication(msalConfig),
    commit: async (homeAccountId?: string) => {
      const file = cachePath(config);
      const release = usesDeleteRestrictedStorage(file) ? undefined : await acquireFileLock(`${file}.lock`);
      try {
        // Do not overwrite a login/logout or refresh performed by another process.
        if (readCache() !== initialCache) throw new Error("AUTH_CACHE_CHANGED");
        if (readProfile() !== initialProfile) throw new Error("AUTH_PROFILE_CHANGED");
        if (pendingCache !== undefined) writeStoredText(file, pendingCache);
        if (homeAccountId) markAuthProfileReady(config, scopes, homeAccountId);
      } finally {
        release?.();
      }
    }
  };
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
  const profile = prepareAuthProfile(config, scopes);
  if (profile.reason) throw new Error(profile.reason);
  const state = readDeviceLoginState(config, scopes);
  const { pca, commit } = await buildPca(config, new ResumeDeviceCodeNetworkClient(state, networkClient));
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

  // The code was redeemed successfully; token verification must never resume it again.
  clearDeviceLoginState(config);
  if (result.account.tenantId !== config.tenantId ||
      (profile.homeAccountId && result.account.homeAccountId !== profile.homeAccountId)) {
    throw new Error("AUTH_ACCOUNT_MISMATCH");
  }
  verifySilentResult(result, result.account);
  await commit(result.account.homeAccountId);
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

export async function getAccount(
  config: AppConfig,
  networkClient: INetworkModule = new ProxyAwareNetworkClient()
): Promise<AccountInfo | null> {
  const status = await getAuthStatus(config, networkClient);
  return status.tokenUsable ? status.account : null;
}

function selectAccount(accounts: AccountInfo[], config: AppConfig, homeAccountId?: string): AccountInfo {
  if (!homeAccountId && accounts.length > 1) throw new Error("AUTH_ACCOUNT_SELECTION_REQUIRED");
  const account = homeAccountId
    ? accounts.find((candidate) => candidate.homeAccountId === homeAccountId)
    : accounts[0];
  if (!account) throw new Error(homeAccountId ? "AUTH_ACCOUNT_MISMATCH" : "NO_ACCOUNT_IN_CACHE");
  if (account.tenantId !== config.tenantId) throw new Error("AUTH_ACCOUNT_MISMATCH");
  return account;
}

function verifySilentResult(result: AuthenticationResult | null, account: AccountInfo): string {
  if (!result?.accessToken) throw new Error("NO_ACCESS_TOKEN");
  if (result.account?.homeAccountId !== account.homeAccountId || result.account.tenantId !== account.tenantId) {
    throw new Error("AUTH_ACCOUNT_MISMATCH");
  }
  const granted = new Set(result.scopes.map((scope) => scope.toLowerCase()));
  const missing = scopes.some((scope) => !["openid", "profile", "offline_access"].includes(scope)
    && !granted.has(scope.toLowerCase()));
  if (missing) throw new Error("AUTH_SCOPES_INSUFFICIENT");
  return result.accessToken;
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
  const message = errorMessage(error);
  if (/^(AUTH_ACCOUNT_MISMATCH|AUTH_ACCOUNT_SELECTION_REQUIRED|AUTH_CACHE_UNREADABLE|AUTH_CACHE_CHANGED|AUTH_PROFILE_UNREADABLE|AUTH_PROFILE_CHANGED|AUTH_SCOPES_INSUFFICIENT|NO_ACCOUNT_IN_CACHE|NO_ACCESS_TOKEN)$/.test(message)) {
    return message;
  }
  const code = (error as { errorCode?: unknown } | null)?.errorCode;
  const knownCode = typeof code === "string"
    ? code.match(/^(interaction_required|invalid_grant|login_required|consent_required)$/i)?.[1] : undefined;
  return `TOKEN_ACQUISITION_FAILED: ${knownCode?.toLowerCase() ?? "unclassified_error"}`;
}

export async function getAuthStatus(
  config: AppConfig,
  networkClient: INetworkModule = new ProxyAwareNetworkClient()
): Promise<AuthStatus> {
  if (networkClient instanceof ProxyAwareNetworkClient) networkClient.takeNetworkFailure();
  const profile = prepareAuthProfile(config, scopes);
  if (profile.migrationRequired || profile.reason) {
    return {
      account: null,
      loggedIn: false,
      tokenUsable: false,
      migrationRequired: profile.migrationRequired,
      reason: profile.reason ?? "AUTH_MIGRATION_REQUIRED"
    };
  }

  let account: AccountInfo | null = null;
  try {
    const { pca, commit } = await buildPca(config, networkClient);
    const accounts = await pca.getTokenCache().getAllAccounts();
    account = selectAccount(accounts, config, profile.homeAccountId);
    const result = await pca.acquireTokenSilent({ account, scopes });
    verifySilentResult(result, account);
    await commit();
    return {
      account,
      loggedIn: true,
      tokenUsable: true,
      migrationRequired: false
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
  const profile = prepareAuthProfile(config, scopes);
  if (profile.migrationRequired || profile.reason) {
    throw new Error(
      `${profile.reason ?? "AUTH_MIGRATION_REQUIRED"}: Authentication profile validation failed. Existing authentication files were preserved.`
    );
  }
  try {
    const { pca, commit } = await buildPca(config, networkClient);
    const accounts = await pca.getTokenCache().getAllAccounts();
    const account = selectAccount(accounts, config, profile.homeAccountId);
    const result = await pca.acquireTokenSilent({ account, scopes });
    const token = verifySilentResult(result, account);
    await commit();
    return token;
  } catch (error) {
    const failure = networkClient instanceof ProxyAwareNetworkClient
      ? networkClient.takeNetworkFailure() : undefined;
    const reason = networkBlockedReason(error, failure);
    if (reason) throw new Error(reason);
    throw new Error(tokenFailureReason(error));
  }
}

export async function logout(config: AppConfig): Promise<void> {
  const file = cachePath(config);
  fs.mkdirSync(config.cacheDir, { recursive: true });
  const release = usesDeleteRestrictedStorage(file) ? undefined : await acquireFileLock(`${file}.lock`);
  try {
    clearStoredFile(file);
    clearDeviceLoginState(config);
    resetAuthProfileAfterLogout(config, scopes);
  } finally {
    release?.();
  }
}

export function getScopeList(): string[] {
  return [...scopes];
}
