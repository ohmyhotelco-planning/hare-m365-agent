import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";
import { storedFileHasContent, writeStoredText } from "./persistent-storage.js";

type AuthProfileState = {
  version: 1;
  clientId: string;
  tenantId: string;
  scopes: string[];
  migrationRequired: boolean;
  homeAccountId?: string;
};

export type AuthProfilePreparation = {
  migrationRequired: boolean;
  authStateCleared: boolean;
  homeAccountId?: string;
  scopesChanged?: boolean;
  reason?: "AUTH_APP_CHANGED" | "AUTH_TENANT_CHANGED" | "AUTH_MIGRATION_REQUIRED"
    | "AUTH_PROFILE_MISSING" | "AUTH_PROFILE_INVALID" | "AUTH_PROFILE_UNREADABLE";
};

function authProfilePath(config: AppConfig): string {
  return path.join(config.cacheDir, "auth-profile.json");
}

function msalCachePath(config: AppConfig): string {
  return path.join(config.cacheDir, "msal-cache.json");
}

function normalizedScopes(scopes: string[]): string[] {
  return [...new Set(scopes)].sort();
}

function currentProfile(
  config: AppConfig,
  scopes: string[],
  migrationRequired: boolean,
  homeAccountId?: string
): AuthProfileState {
  return {
    version: 1,
    clientId: config.clientId,
    tenantId: config.tenantId,
    scopes: normalizedScopes(scopes),
    migrationRequired,
    ...(homeAccountId ? { homeAccountId } : {})
  };
}

function readProfile(config: AppConfig): { profile?: AuthProfileState; reason?: AuthProfilePreparation["reason"] } {
  let text: string;
  try {
    text = fs.readFileSync(authProfilePath(config), "utf8");
  } catch (error) {
    return { reason: (error as NodeJS.ErrnoException).code === "ENOENT"
      ? "AUTH_PROFILE_MISSING" : "AUTH_PROFILE_UNREADABLE" };
  }
  try {
    const value = JSON.parse(text) as Partial<AuthProfileState> | null;
    if (
      !value ||
      value.version !== 1 ||
      typeof value.clientId !== "string" ||
      typeof value.tenantId !== "string" ||
      !Array.isArray(value.scopes) ||
      !value.scopes.every((scope) => typeof scope === "string") ||
      typeof value.migrationRequired !== "boolean" ||
      (value.homeAccountId !== undefined && (typeof value.homeAccountId !== "string" || !value.homeAccountId))
    ) {
      return { reason: "AUTH_PROFILE_INVALID" };
    }
    return { profile: value as AuthProfileState };
  } catch {
    return { reason: "AUTH_PROFILE_INVALID" };
  }
}

function writeProfile(config: AppConfig, scopes: string[], migrationRequired: boolean, homeAccountId?: string): void {
  writeStoredText(
    authProfilePath(config),
    `${JSON.stringify(currentProfile(config, scopes, migrationRequired, homeAccountId), null, 2)}\n`
  );
}

// Inspection never rewrites profile metadata or resets authentication state.
export function prepareAuthProfile(
  config: AppConfig,
  scopes: string[]
): AuthProfilePreparation {
  const { profile, reason } = readProfile(config);
  if (!profile) {
    if (reason === "AUTH_PROFILE_MISSING") {
      try {
        if (!storedFileHasContent(msalCachePath(config))) {
          return { migrationRequired: false, authStateCleared: false };
        }
      } catch {
        return { migrationRequired: false, authStateCleared: false, reason: "AUTH_PROFILE_UNREADABLE" };
      }
    }
    return { migrationRequired: false, authStateCleared: false, reason };
  }
  const mismatch = profile.clientId !== config.clientId ? "AUTH_APP_CHANGED"
    : profile.tenantId !== config.tenantId ? "AUTH_TENANT_CHANGED"
    : profile.migrationRequired ? "AUTH_MIGRATION_REQUIRED" : undefined;
  if (mismatch) return { migrationRequired: true, authStateCleared: false, reason: mismatch };
  return {
    migrationRequired: false,
    authStateCleared: false,
    ...(profile.homeAccountId ? { homeAccountId: profile.homeAccountId } : {}),
    ...(normalizedScopes(profile.scopes).join("\n") !== normalizedScopes(scopes).join("\n")
      ? { scopesChanged: true } : {})
  };
}

export function markAuthProfileReady(config: AppConfig, scopes: string[], homeAccountId?: string): void {
  writeProfile(config, scopes, false, homeAccountId);
}

export function resetAuthProfileAfterLogout(config: AppConfig, scopes: string[]): void {
  writeProfile(config, scopes, false);
}

export function getAuthProfilePath(config: AppConfig): string {
  return authProfilePath(config);
}
