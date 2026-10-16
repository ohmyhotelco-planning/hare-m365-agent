import type { AppConfig } from "./config.js";
import {
  defaultGroupClient,
  groupIdentifier,
  groupPageUrl,
  groupPath,
  groupReadError,
  type GroupReadClient,
  type GroupReference
} from "./groups-common.js";

const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const pageSize = 25;
const maxPages = 10;
const maxResults = pageSize * maxPages;
const lookupTimeoutMs = 15_000;
const select = "id,displayName,mail,mailNickname,proxyAddresses,groupTypes,mailEnabled,securityEnabled";

type DirectoryGroup = {
  id: string;
  displayName?: string | null;
  mail?: string | null;
  mailNickname?: string | null;
  proxyAddresses?: string[] | null;
  groupTypes?: string[] | null;
  mailEnabled?: boolean | null;
  securityEnabled?: boolean | null;
};

export async function resolveGroup(
  config: AppConfig,
  input: string,
  client: GroupReadClient = defaultGroupClient(config)
): Promise<GroupReference> {
  try {
    const identifier = groupIdentifier(input, "group");
    const deadline = Date.now() + lookupTimeoutMs;
    if (guid.test(identifier)) {
      const params = new URLSearchParams({ "$select": select });
      const group = readGroup(await boundedGet(client, `${groupPath(identifier)}?${params}`, deadline));
      if (group.id.toLowerCase() !== identifier.toLowerCase()) {
        throw incomplete("The returned group ID does not match the requested ID.");
      }
      return reference(group, ["id: exact GUID match"]);
    }

    const smtp = identifier.includes("@");
    if (smtp ? !/^[^\s@:<>]+@[^\s@:<>]+$/.test(identifier) :
      identifier.length > 64 || /[^\x21-\x7e]|[@()\\\[\]";:<>,]/.test(identifier)) {
      throw new Error("GROUP_INVALID_IDENTIFIER: Use a GUID, exact SMTP address, or bare mailNickname alias.");
    }
    const escaped = identifier.replace(/'/g, "''");
    // These equality filters use ordinary queries, not advanced directory queries.
    const filter = smtp
      ? `mail eq '${escaped}' or proxyAddresses/any(p:p eq 'smtp:${escaped}') or proxyAddresses/any(p:p eq 'SMTP:${escaped}')`
      : `mailNickname eq '${escaped}'`;
    const params = new URLSearchParams({ "$filter": filter, "$select": select, "$top": String(pageSize) });
    let next: string | undefined = `/groups?${params}`;
    const visited = new Set<string>();
    const ids = new Set<string>();
    const matches: GroupReference[] = [];
    let pages = 0;
    let reportedCount: number | undefined;

    while (next !== undefined) {
      if (pages >= maxPages) throw incomplete("The directory lookup reached its page limit.");
      const pageUrl = continuation(next);
      const canonical = new URL(pageUrl);
      canonical.searchParams.sort();
      if (visited.has(canonical.toString())) throw incomplete("The directory continuation loops.");
      visited.add(canonical.toString());
      const page = await boundedGet(client, next, deadline);
      pages += 1;
      if (!isRecord(page) || !Array.isArray(page.value) || page.value.length > pageSize || "@odata.deltaLink" in page) {
        throw incomplete("The directory response is not a complete group page.");
      }
      if ("@odata.count" in page) {
        const count = page["@odata.count"];
        if (typeof count !== "number" || !Number.isSafeInteger(count) || count < 0 ||
          count > maxResults || (reportedCount !== undefined && reportedCount !== count)) {
          throw incomplete("The directory result count is invalid, changed, or exceeds the lookup limit.");
        }
        reportedCount = count;
      }
      for (const value of page.value) {
        const group = readGroup(value);
        const id = group.id.toLowerCase();
        // Repeated IDs across pages can hide a changed or truncated result set.
        if (ids.has(id)) throw incomplete("The directory lookup repeated a group ID.");
        ids.add(id);
        if (ids.size > maxResults) throw incomplete("The directory lookup reached its result limit.");
        const evidence = matchEvidence(group, identifier, smtp);
        if (evidence.length) matches.push(reference(group, evidence));
      }
      if (reportedCount !== undefined && ids.size > reportedCount) {
        throw incomplete("The directory result count does not match the returned groups.");
      }
      next = undefined;
      if ("@odata.nextLink" in page) {
        if (typeof page["@odata.nextLink"] !== "string" || !page["@odata.nextLink"]) {
          throw incomplete("The directory continuation is missing or incomplete.");
        }
        next = continuation(page["@odata.nextLink"]);
      } else if (reportedCount !== undefined && ids.size !== reportedCount) {
        throw incomplete("The directory response was truncated before all reported groups were returned.");
      }
    }
    if (matches.length === 0) {
      throw new Error("GROUP_DIRECTORY_NO_MATCH: No exact group match was visible in this lookup. This is not proof that the group does not exist or that it is a shared mailbox.");
    }
    if (matches.length !== 1) {
      throw new Error("GROUP_DIRECTORY_AMBIGUOUS: Multiple group IDs matched exactly. Supply the intended group GUID; no candidate was selected.");
    }
    return matches[0];
  } catch (error) {
    throw groupReadError("DIRECTORY", error);
  }
}

async function boundedGet(client: GroupReadClient, url: string, deadline: number): Promise<unknown> {
  const remaining = deadline - Date.now();
  if (remaining <= 0) throw new Error("Group directory lookup timed out.");
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    // Also bound token acquisition and JSON decoding outside graphRequest's timer.
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new Error("Group directory lookup timed out.")), remaining);
    });
    const result = await Promise.race([client.get<unknown>(url, { totalTimeoutMs: remaining }), timeout]);
    if (Date.now() >= deadline) throw new Error("Group directory lookup timed out.");
    return result;
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}

function continuation(value: string): string {
  if (value.trim() !== value || /[\u0000-\u0020\u007f\\]/.test(value)) {
    throw incomplete("The directory continuation is malformed.");
  }
  try {
    return groupPageUrl(value, "/groups");
  } catch {
    throw incomplete("The directory continuation must stay in the same Graph group collection.");
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function readGroup(value: unknown): DirectoryGroup {
  if (!isRecord(value) || typeof value.id !== "string" || !value.id.trim() ||
    value.id !== value.id.trim() || value.id.length > 4096 || /[\u0000-\u001f\u007f]/.test(value.id) ||
    (value["@odata.type"] !== undefined && value["@odata.type"] !== "#microsoft.graph.group")) {
    throw incomplete("The directory response contains an invalid group identity.");
  }
  for (const key of ["displayName", "mail", "mailNickname"] as const) {
    if (value[key] != null && typeof value[key] !== "string") throw incomplete(`Invalid group ${key} field.`);
  }
  for (const key of ["mailEnabled", "securityEnabled"] as const) {
    if (value[key] != null && typeof value[key] !== "boolean") throw incomplete(`Invalid group ${key} field.`);
  }
  for (const key of ["groupTypes", "proxyAddresses"] as const) {
    const field = value[key];
    if (field != null && (!Array.isArray(field) || !field.every((item) => typeof item === "string"))) {
      throw incomplete(`Invalid group ${key} field.`);
    }
  }
  if ("proxyAddresses@odata.nextLink" in value || "groupTypes@odata.nextLink" in value) {
    throw incomplete("The directory group contains a truncated identity or classification collection.");
  }
  return value as DirectoryGroup;
}

function matchEvidence(group: DirectoryGroup, input: string, smtp: boolean): string[] {
  const normalized = input.toLowerCase();
  const evidence: string[] = [];
  if (smtp) {
    if (group.mail?.toLowerCase() === normalized) evidence.push("mail: exact SMTP match");
    if (group.proxyAddresses?.some((address) => address.toLowerCase() === `smtp:${normalized}`)) {
      evidence.push("proxyAddresses: exact SMTP match");
    }
    if (!evidence.length && (group.mail === undefined || group.proxyAddresses == null)) {
      throw incomplete("A directory candidate lacks the fields needed to exclude an exact SMTP match.");
    }
  } else {
    if (group.mailNickname?.toLowerCase() === normalized) evidence.push("mailNickname: exact bare alias match (not SMTP evidence)");
    if (group.mailNickname == null) throw incomplete("A directory candidate lacks mailNickname match evidence.");
  }
  return evidence;
}

function reference(group: DirectoryGroup, matchingEvidence: string[]): GroupReference {
  const groupTypes = group.groupTypes ?? null;
  const mailEnabled = group.mailEnabled ?? null;
  const securityEnabled = group.securityEnabled ?? null;
  let classification: GroupReference["classification"] = "unknown";
  if (groupTypes?.includes("Unified")) {
    classification = "microsoft365";
  } else if (groupTypes !== null && mailEnabled !== null && securityEnabled !== null) {
    if (mailEnabled && securityEnabled) classification = "mail-enabled-security";
    else if (mailEnabled && !securityEnabled) classification = "distribution";
    else if (!mailEnabled && securityEnabled) classification = "security";
  }
  const evidence = [...matchingEvidence];
  if (groupTypes !== null) evidence.push(`groupTypes: ${JSON.stringify(groupTypes)}`);
  if (mailEnabled !== null) evidence.push(`mailEnabled: ${mailEnabled}`);
  if (securityEnabled !== null) evidence.push(`securityEnabled: ${securityEnabled}`);
  if (classification === "unknown") evidence.push("classification: insufficient or unrecognized directory evidence");
  return {
    id: group.id,
    ...(group.displayName != null ? { displayName: group.displayName } : {}),
    mail: group.mail ?? null,
    groupTypes,
    mailEnabled,
    securityEnabled,
    classification,
    evidence,
    exchangeRecipientType: "unverified"
  };
}

function incomplete(reason: string): Error {
  return new Error(`GROUP_DIRECTORY_LOOKUP_INCOMPLETE: ${reason} No candidate was selected.`);
}
