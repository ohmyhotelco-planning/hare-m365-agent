import type { AppConfig } from "./config.js";
import { graphGet, type GraphPage, type GraphRequestOptions } from "./graph.js";

export type MailboxTarget =
  | { scope: "self" }
  | {
      scope: "shared";
      input: string;
      displayName?: string;
      address: string;
      userId: string;
    };

export type MailboxReference = {
  scope: "self" | "shared";
  input?: string;
  displayName?: string;
  address?: string;
};

type DirectoryUser = {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
};

export type MailboxDirectoryClient = {
  get<T>(pathOrUrl: string, options?: GraphRequestOptions): Promise<T>;
};

const directoryResultLimit = 25;

export function selfMailboxTarget(): MailboxTarget {
  return { scope: "self" };
}

export async function resolveMailboxTarget(
  config: AppConfig,
  input: string | undefined,
  client: MailboxDirectoryClient = defaultDirectoryClient(config)
): Promise<MailboxTarget> {
  const normalized = input?.trim();
  if (!normalized) return selfMailboxTarget();
  if (/\p{Cc}/u.test(normalized)) throw new Error("mailbox must not contain control characters.");

  const candidates = normalized.includes("@")
    ? await findAddressCandidates(client, normalized)
    : await findNameCandidates(client, normalized);
  const exactCandidates = candidates.filter((candidate) =>
    normalized.includes("@")
      ? isExactAddress(candidate, normalized)
      : isExactDisplayName(candidate, normalized)
  );

  if (exactCandidates.length === 0 && candidates.length === 0) {
    throw new Error(
      `SHARED_MAILBOX_NOT_FOUND: No directory mailbox matched "${normalized}". Use the exact mailbox address.`
    );
  }
  if (exactCandidates.length === 0) {
    throw new Error(
      `SHARED_MAILBOX_NAME_NOT_EXACT: "${normalized}" did not exactly match a mailbox name. Candidates: ${formatCandidates(candidates)}. Use the exact mailbox address.`
    );
  }
  if (exactCandidates.length > 1) {
    throw new Error(
      `SHARED_MAILBOX_AMBIGUOUS: "${normalized}" matched multiple mailboxes: ${formatCandidates(exactCandidates)}. Use the exact mailbox address.`
    );
  }

  const selected = exactCandidates[0];
  return sharedMailboxTarget(
    normalized,
    mailboxAddress(selected),
    selected.id,
    selected.displayName
  );
}

export function mailboxReference(target: MailboxTarget): MailboxReference {
  return target.scope === "self"
    ? { scope: "self" }
    : {
        scope: "shared",
        input: target.input,
        displayName: target.displayName,
        address: target.address
      };
}

export function mailboxKey(target: MailboxTarget): string {
  return target.scope === "self" ? "me" : `shared:${target.userId.toLocaleLowerCase()}`;
}

export function mailboxPath(target: MailboxTarget, suffix: string): string {
  const normalizedSuffix = suffix.startsWith("/") ? suffix : `/${suffix}`;
  const basePath = target.scope === "self"
    ? "/me"
    : `/users/${encodeURIComponent(target.userId)}`;
  return `${basePath}${normalizedSuffix}`;
}

export async function runMailboxRead<T>(
  target: MailboxTarget,
  operation: () => Promise<T>
): Promise<T> {
  try {
    return await operation();
  } catch (error) {
    if (target.scope !== "shared") throw error;
    if (isGraphStatus(error, 403)) {
      throw new Error(
        `SHARED_MAILBOX_ACCESS_DENIED: The signed-in user cannot read ${target.address}.`
      );
    }
    if (isGraphStatus(error, 404)) {
      throw new Error(
        `SHARED_MAILBOX_NOT_FOUND: ${target.address} was not found or is unavailable to the signed-in user.`
      );
    }
    throw error;
  }
}

function sharedMailboxTarget(
  input: string,
  address: string,
  userId: string,
  displayName?: string
): MailboxTarget {
  return {
    scope: "shared",
    input,
    displayName,
    address: address.trim(),
    userId
  };
}

async function findAddressCandidates(
  client: MailboxDirectoryClient,
  address: string
): Promise<DirectoryUser[]> {
  const escapedAddress = escapeODataString(address);
  const params = new URLSearchParams({
    "$filter": `mail eq '${escapedAddress}' or userPrincipalName eq '${escapedAddress}'`,
    "$select": "id,displayName,mail,userPrincipalName",
    "$top": String(directoryResultLimit),
    "$count": "true"
  });
  return queryDirectory(client, params);
}

async function findNameCandidates(
  client: MailboxDirectoryClient,
  name: string
): Promise<DirectoryUser[]> {
  const params = new URLSearchParams({
    "$search": `\"displayName:${escapeDirectorySearch(name)}\"`,
    "$select": "id,displayName,mail,userPrincipalName",
    "$top": String(directoryResultLimit),
    "$count": "true"
  });
  return queryDirectory(client, params);
}

async function queryDirectory(
  client: MailboxDirectoryClient,
  params: URLSearchParams
): Promise<DirectoryUser[]> {
  try {
    const page = await client.get<GraphPage<DirectoryUser>>(`/users?${params.toString()}`, {
      headers: { ConsistencyLevel: "eventual" }
    });
    return uniqueCandidates(page.value ?? []);
  } catch (error) {
    if (isGraphStatus(error, 403)) {
      throw new Error(
        "MAILBOX_DIRECTORY_LOOKUP_DENIED: Hare cannot resolve a shared mailbox. Sign in again after User.ReadBasic.All is granted."
      );
    }
    throw error;
  }
}

function uniqueCandidates(users: DirectoryUser[]): DirectoryUser[] {
  const unique = new Map<string, DirectoryUser>();
  for (const user of users) {
    const key = user.id?.trim();
    if (!key || !mailboxAddress(user, false)) continue;
    unique.set(key.toLocaleLowerCase(), user);
  }
  return [...unique.values()];
}

function mailboxAddress(user: DirectoryUser, required = true): string {
  const value = user.mail?.trim() || user.userPrincipalName?.trim();
  if (!value && required) throw new Error("Directory mailbox candidate has no usable address.");
  return value ?? "";
}

function isExactAddress(user: DirectoryUser, input: string): boolean {
  const normalized = input.toLocaleLowerCase();
  return user.mail?.trim().toLocaleLowerCase() === normalized
    || user.userPrincipalName?.trim().toLocaleLowerCase() === normalized;
}

function isExactDisplayName(user: DirectoryUser, input: string): boolean {
  return user.displayName?.trim().toLocaleLowerCase() === input.toLocaleLowerCase();
}

function formatCandidates(candidates: DirectoryUser[]): string {
  return candidates
    .slice(0, 5)
    .map((candidate) => `${candidate.displayName ?? "(unnamed)"} <${mailboxAddress(candidate)}>`)
    .join(", ");
}

function escapeDirectorySearch(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/\"/g, '\\"');
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

function isGraphStatus(error: unknown, status: number): boolean {
  return error instanceof Error && new RegExp(`Graph \\w+ failed \\(${status}\\b`, "i").test(error.message);
}

function defaultDirectoryClient(config: AppConfig): MailboxDirectoryClient {
  return {
    get: <T>(pathOrUrl: string, options?: GraphRequestOptions) =>
      graphGet<T>(config, pathOrUrl, options)
  };
}
