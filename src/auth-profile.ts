import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";
import { clearStoredFile, storedFileHasContent, writeStoredText } from "./persistent-storage.js";

type AuthProfileState = {
  version: 1;
  clientId: string;
  tenantId: string;
  scopes: string[];
  migrationRequired: boolean;
};

export type AuthProfilePreparation = {
  migrationRequired: boolean;
  authStateCleared: boolean;
};

function authProfilePath(config: AppConfig): string {
  return path.join(config.cacheDir, "auth-profile.json");
}

function msalCachePath(config: AppConfig): string {
  return path.join(config.cacheDir, "msal-cache.json");
}

function pendingLoginPath(config: AppConfig): string {
  return path.join(config.cacheDir, "device-login-state.json");
}

function normalizedScopes(scopes: string[]): string[] {
  return [...new Set(scopes)].sort();
}

function currentProfile(
  config: AppConfig,
  scopes: string[],
  migrationRequired: boolean
): AuthProfileState {
  return {
    version: 1,
    clientId: config.clientId,
    tenantId: config.tenantId,
    scopes: normalizedScopes(scopes),
    migrationRequired
  };
}

function readProfile(config: AppConfig): AuthProfileState | undefined {
  const file = authProfilePath(config);
  if (!storedFileHasContent(file)) return undefined;

  try {
    const value = JSON.parse(fs.readFileSync(file, "utf8")) as Partial<AuthProfileState>;
    if (
      value.version !== 1 ||
      typeof value.clientId !== "string" ||
      typeof value.tenantId !== "string" ||
      !Array.isArray(value.scopes) ||
      !value.scopes.every((scope) => typeof scope === "string") ||
      typeof value.migrationRequired !== "boolean"
    ) {
      return undefined;
    }
    return value as AuthProfileState;
  } catch {
    return undefined;
  }
}

function profileMatches(
  profile: AuthProfileState,
  config: AppConfig,
  scopes: string[]
): boolean {
  return (
    profile.clientId === config.clientId &&
    profile.tenantId === config.tenantId &&
    profile.scopes.join("\n") === normalizedScopes(scopes).join("\n")
  );
}

function writeProfile(config: AppConfig, scopes: string[], migrationRequired: boolean): void {
  writeStoredText(
    authProfilePath(config),
    `${JSON.stringify(currentProfile(config, scopes, migrationRequired), null, 2)}\n`
  );
}

function createProfileIfAbsent(
  config: AppConfig,
  scopes: string[],
  migrationRequired: boolean
): boolean {
  const file = authProfilePath(config);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  try {
    fs.writeFileSync(
      file,
      `${JSON.stringify(currentProfile(config, scopes, migrationRequired), null, 2)}\n`,
      { encoding: "utf8", mode: 0o600, flag: "wx" }
    );
    return true;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "EEXIST") return false;
    throw error;
  }
}

function clearAuthenticationState(config: AppConfig): void {
  clearStoredFile(msalCachePath(config));
  clearStoredFile(pendingLoginPath(config));
}

export function prepareAuthProfile(
  config: AppConfig,
  scopes: string[]
): AuthProfilePreparation {
  const profile = readProfile(config);
  const hasExistingAuthState =
    storedFileHasContent(msalCachePath(config)) || storedFileHasContent(pendingLoginPath(config));

  if (!profile) {
    if (hasExistingAuthState) {
      createProfileIfAbsent(config, scopes, true);
      clearAuthenticationState(config);
      writeProfile(config, scopes, true);
      return { migrationRequired: true, authStateCleared: true };
    }

    if (createProfileIfAbsent(config, scopes, false)) {
      return { migrationRequired: false, authStateCleared: false };
    }

    const concurrentlyCreatedProfile = readProfile(config);
    if (concurrentlyCreatedProfile && profileMatches(concurrentlyCreatedProfile, config, scopes)) {
      return {
        migrationRequired: concurrentlyCreatedProfile.migrationRequired,
        authStateCleared: false
      };
    }

    writeProfile(config, scopes, false);
    return { migrationRequired: false, authStateCleared: false };
  }

  if (!profileMatches(profile, config, scopes)) {
    clearAuthenticationState(config);
    writeProfile(config, scopes, true);
    return { migrationRequired: true, authStateCleared: hasExistingAuthState };
  }

  return {
    migrationRequired: profile.migrationRequired,
    authStateCleared: false
  };
}

export function markAuthProfileReady(config: AppConfig, scopes: string[]): void {
  writeProfile(config, scopes, false);
}

export function resetAuthProfileAfterLogout(config: AppConfig, scopes: string[]): void {
  writeProfile(config, scopes, false);
}

export function getAuthProfilePath(config: AppConfig): string {
  return authProfilePath(config);
}
