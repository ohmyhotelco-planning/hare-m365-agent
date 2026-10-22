const phases = ["load", "account", "silent", "validate", "commit"] as const;
export type AuthPhase = typeof phases[number];
const reasons = new Set([
  "AUTH_APP_CHANGED", "AUTH_TENANT_CHANGED", "AUTH_MIGRATION_REQUIRED", "AUTH_PROFILE_INVALID", "AUTH_PROFILE_MISSING",
  "AUTH_PROFILE_UNREADABLE", "AUTH_PROFILE_CHANGED", "AUTH_ACCOUNT_MISMATCH", "AUTH_ACCOUNT_SELECTION_REQUIRED",
  "AUTH_CACHE_UNREADABLE", "AUTH_CACHE_CHANGED", "AUTH_SCOPES_INSUFFICIENT", "NO_ACCOUNT_IN_CACHE", "NO_ACCESS_TOKEN",
  "AUTH_CHECK_BLOCKED", "TOKEN_ACQUISITION_FAILED"
]);
const oauthCodes = new Set(["interaction_required", "invalid_grant", "login_required", "consent_required", "server_error"]);
const ioCodes = new Set(["EACCES", "EPERM", "ENOENT", "EEXIST", "EIO", "ENOSPC", "ETIMEDOUT", "ECONNRESET", "ENOTFOUND", "EAI_AGAIN", "UND_ERR_CONNECT_TIMEOUT"]);
export type AuthProvenance = { phase: AuthPhase; reason: string; errorCode?: string; ioCode?: string };
type Annotated = { authProvenance?: unknown; errorCode?: unknown; code?: unknown };
const allowed = (value: unknown, values: Set<string>): string | undefined => typeof value === "string" && values.has(value) ? value : undefined;

// Only fixed categories cross this boundary; never copy messages, paths, causes or tokens.
export function readAuthProvenance(error: unknown): AuthProvenance | undefined {
  const value = (error as Annotated | null)?.authProvenance as Partial<AuthProvenance> | undefined;
  if (!value || !phases.includes(value.phase as AuthPhase) || !allowed(value.reason, reasons)) return undefined;
  return { phase: value.phase!, reason: value.reason!,
    ...(allowed(value.errorCode, oauthCodes) ? { errorCode: value.errorCode } : {}),
    ...(allowed(value.ioCode, ioCodes) ? { ioCode: value.ioCode } : {}) };
}

export function copyAuthProvenance<T extends Error>(target: T, source: unknown): T {
  const authProvenance = readAuthProvenance(source);
  return authProvenance ? Object.assign(target, { authProvenance }) : target;
}

export function authFailure(message: string, phase: AuthPhase, source: unknown, networkCode?: string): Error {
  const error = new Error(message);
  const reason = allowed(message.split(":", 1)[0], reasons);
  if (!reason) return error;
  const original = source as Annotated | null;
  const errorCode = allowed(original?.errorCode, oauthCodes);
  const ioCode = allowed(original?.code, ioCodes) ?? allowed(networkCode, ioCodes);
  return Object.assign(error, { authProvenance: { phase, reason, ...(errorCode ? { errorCode } : {}), ...(ioCode ? { ioCode } : {}) } });
}
