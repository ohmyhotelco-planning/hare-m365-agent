const graphHost = "graph.microsoft.com";
const graphRoot = "https://graph.microsoft.com/v1.0";

export function encodeGraphContinuation(url: string): string {
  const normalized = normalizeGraphUrl(url);
  return Buffer.from(normalized, "utf8").toString("base64url");
}

export function decodeGraphContinuation(cursor: string): string {
  if (!cursor.trim()) throw new Error("cursor must not be empty.");
  let decoded: string;
  try {
    decoded = Buffer.from(cursor, "base64url").toString("utf8");
  } catch {
    throw new Error("cursor is invalid.");
  }
  validateGraphUrl(decoded);
  return decoded;
}

function validateGraphUrl(value: string): void {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new Error("cursor does not contain a valid URL.");
  }
  if (url.protocol !== "https:" || url.hostname !== graphHost || !url.pathname.startsWith("/v1.0/")) {
    throw new Error("cursor must reference a Microsoft Graph v1.0 continuation URL.");
  }
}

function normalizeGraphUrl(value: string): string {
  const normalized = value.startsWith("/") ? `${graphRoot}${value}` : value;
  validateGraphUrl(normalized);
  return normalized;
}
