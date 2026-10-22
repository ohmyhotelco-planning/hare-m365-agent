import type { AppConfig } from "./config.js";
import { graphGet, type GraphRequestOptions } from "./graph.js";

export type GroupReference = {
  id: string;
  displayName?: string;
  mail: string | null;
  groupTypes: string[] | null;
  mailEnabled: boolean | null;
  securityEnabled: boolean | null;
  classification: "microsoft365" | "mail-enabled-security" | "distribution" | "security" | "unknown";
  evidence: string[];
  exchangeRecipientType: "unverified";
};

export type GroupReadClient = {
  get<T>(pathOrUrl: string, options?: GraphRequestOptions): Promise<T>;
};

export type GroupStage = "DIRECTORY" | "CONVERSATIONS" | "THREADS" | "POSTS" | "POST_SUBJECT" | "ATTACHMENT_LIST" | "ATTACHMENT_METADATA" | "ATTACHMENT_BYTES";

export function defaultGroupClient(config: AppConfig): GroupReadClient {
  return { get: <T>(url: string, options?: GraphRequestOptions) => graphGet<T>(config, url, options) };
}

export function groupIdentifier(value: string, label = "id"): string {
  if (typeof value !== "string" || !value.trim() || value.length > 4096 || /[\u0000-\u001f\u007f]/.test(value)) {
    throw new Error(`GROUP_INVALID_IDENTIFIER: ${label} is required and must not contain control characters.`);
  }
  return value.trim();
}

export function groupPath(groupId: string): string {
  return `/groups/${encodeURIComponent(groupIdentifier(groupId, "group-id"))}`;
}

export function requireMicrosoft365Group(group: GroupReference): void {
  if (group.classification !== "microsoft365" || !group.groupTypes?.includes("Unified")) {
    throw new Error("GROUP_TYPE_UNSUPPORTED: Direct conversation reads require directory evidence groupTypes includes Unified. Exchange recipient subtype remains unverified.");
  }
}

// Bind server continuations to the exact collection before passing a bearer token.
export function groupPageUrl(value: string, collectionPath: string): string {
  const url = new URL(value.startsWith("/") ? `https://graph.microsoft.com/v1.0${value}` : value);
  const expected = `/v1.0${collectionPath}`;
  if (url.origin !== "https://graph.microsoft.com" || url.username || url.password || url.hash || url.pathname !== expected) {
    throw new Error("GROUP_INVALID_CONTINUATION: Continuation must stay in the same Graph group collection.");
  }
  return url.toString();
}

export function groupReadError(stage: GroupStage, error: unknown): Error {
  if (error instanceof Error && error.message.startsWith("GROUP_")) return error;
  const message = error instanceof Error ? error.message : String(error);
  const status = /Graph \w+ failed \((\d{3})\b/i.exec(message)?.[1];
  let code: string;
  let explanation: string;
  if (status === "429" || ["500", "502", "503", "504"].includes(status ?? "")) {
    code = status === "429" ? "THROTTLED" : "SERVICE_UNAVAILABLE";
    explanation = "The bounded request retries did not complete. Resume only with the returned continuation and retry allowance.";
  } else if (!status && /Graph request exceeded (?:the )?\d+ms time budget/i.test(message)) {
    code = "TIME_BUDGET_EXCEEDED";
    explanation = "The request exhausted its time budget; this is not an authentication or identity failure.";
  } else if (!status && !/AUTH_|TOKEN_|LOGIN_|interaction_required|invalid_grant|ENOTFOUND|EACCES|EPERM|blocked-by-allowlist/i.test(message) && /fetch failed|socket|ECONNRESET|ETIMEDOUT|EAI_AGAIN|AbortError/i.test(message)) {
    code = "TRANSIENT_NETWORK";
    explanation = "A transient transport failure prevented this bounded read.";
  } else if (status === "403") {
    code = "ACCESS_DENIED";
    explanation = "The current delegated token or user access does not permit this stage. No automatic consent or sign-in was started.";
  } else if (status === "404") {
    code = "NOT_FOUND_OR_UNAVAILABLE";
    explanation = "The requested resource chain was not found or is unavailable to the signed-in user.";
  } else if (status === "400" || status === "405" || status === "501") {
    code = "API_UNSUPPORTED";
    explanation = "The API rejected this operation for the requested resource.";
  } else if (status === "401" || /AUTH_|TOKEN_|LOGIN_|interaction_required|invalid_grant/i.test(message)) {
    code = "AUTH_BLOCKED";
    explanation = "Authentication could not be used. Preserve the existing login and follow the authentication gate.";
  } else if (/timeout|timed out|network|fetch failed|ENOTFOUND|EACCES|ECONN|EAI_AGAIN|ETIMEDOUT/i.test(message)) {
    code = "NETWORK_BLOCKED";
    explanation = "Connectivity or the request time budget prevented this stage; this is not evidence of logout.";
  } else {
    code = "FAILED";
    explanation = "This read stage failed. No mailbox fallback was attempted.";
  }
  return new Error(`GROUP_${stage}_${code}: ${explanation}${status ? ` HTTP ${status}.` : ""}`);
}

export function retryableGroupReadFailure(message: string): boolean {
  return /^GROUP_[A-Z_]+_(THROTTLED|SERVICE_UNAVAILABLE|TIME_BUDGET_EXCEEDED|TRANSIENT_NETWORK):/.test(message);
}
