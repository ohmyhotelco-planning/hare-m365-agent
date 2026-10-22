import { groupPath, groupReadError, type GroupReadClient } from "./groups-common.js";
import { copyAuthProvenance, readAuthProvenance, type AuthProvenance } from "./auth-diagnostics.js";

export type GroupPostSubject = {
  source: "PidTagSubject";
  propertyId: "String 0x0037";
  scope: "POST";
  status: "PRESENT" | "EMPTY" | "MISSING" | "UNVERIFIED" | "INVALID_RESPONSE" | "READ_FAILED" | "NOT_ATTEMPTED_BUDGET";
  value?: string;
  failure?: string;
  authProvenance?: AuthProvenance;
};

export function unavailablePostSubject(status: GroupPostSubject["status"]): GroupPostSubject {
  return { source: "PidTagSubject", propertyId: "String 0x0037", scope: "POST", status };
}

// Read only the final verified post; parent topic and body text are never subject evidence.
export async function readGroupPostSubject(
  chain: { groupId: string; conversationId: string; threadId: string; postId: string; revalidated: boolean },
  read: GroupReadClient["get"]
): Promise<GroupPostSubject> {
  const query = new URLSearchParams({
    "$select": "id,conversationId,conversationThreadId",
    "$expand": "singleValueExtendedProperties($filter=id eq 'String 0x0037')"
  });
  const url = `${groupPath(chain.groupId)}/threads/${encodeURIComponent(chain.threadId)}/posts/${encodeURIComponent(chain.postId)}?${query}`;
  let detail: {
    id?: unknown; conversationId?: unknown; conversationThreadId?: unknown;
    singleValueExtendedProperties?: unknown;
    "singleValueExtendedProperties@odata.nextLink"?: unknown;
  };
  try {
    detail = await read(url);
  } catch (error) {
    const failure = groupReadError("POST_SUBJECT", copyAuthProvenance(new Error(`Remote failure: ${error instanceof Error ? error.message : "unknown"}`), error));
    const authProvenance = readAuthProvenance(failure);
    return { ...unavailablePostSubject("READ_FAILED"), failure: failure.message, ...(authProvenance ? { authProvenance } : {}) };
  }
  if (!detail || detail.id !== chain.postId ||
      (detail.conversationId != null && detail.conversationId !== chain.conversationId) ||
      (detail.conversationThreadId !== undefined && detail.conversationThreadId !== chain.threadId) ||
      (chain.revalidated && detail.conversationThreadId !== chain.threadId)) {
    throw new Error("GROUP_ID_CHAIN_MISMATCH: Subject response does not match the verified post chain.");
  }
  const properties = detail.singleValueExtendedProperties;
  if (properties === undefined) return unavailablePostSubject("UNVERIFIED");
  if (!Array.isArray(properties) || properties.some(p => !p || typeof p.id !== "string" || typeof p.value !== "string")) {
    return unavailablePostSubject("INVALID_RESPONSE");
  }
  const matches = properties.filter(p => /^String 0x0*37$/i.test(p.id));
  if (matches.length > 1) return unavailablePostSubject("INVALID_RESPONSE");
  if (!matches.length) return unavailablePostSubject(detail["singleValueExtendedProperties@odata.nextLink"] === undefined ? "MISSING" : "UNVERIFIED");
  const value: string = matches[0].value;
  return { ...unavailablePostSubject(value.length ? "PRESENT" : "EMPTY"), value };
}
