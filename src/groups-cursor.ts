import { createHmac, randomBytes, timingSafeEqual } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";

const maximumLength = 3_000_000;
const ttlMs = 24 * 60 * 60 * 1000;

// This local key signs only traversal state. It is independent of M365 credentials.
function keyFor(config: AppConfig, create: boolean): Buffer {
  if (!config.cacheDir || !path.isAbsolute(config.cacheDir)) throw invalid();
  const file = path.join(config.cacheDir, "group-continuation.key");
  try {
    if (create && !fs.existsSync(file)) {
      fs.mkdirSync(config.cacheDir, { recursive: true, mode: 0o700 });
      try { fs.writeFileSync(file, randomBytes(32), { flag: "wx", mode: 0o600 }); }
      catch (error) { if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error; }
    }
    const key = fs.readFileSync(file);
    if (key.length !== 32) throw invalid();
    return key;
  } catch { throw invalid(); }
}

export function encodeGroupCursor(config: AppConfig, kind: "posts" | "attachments", state: unknown, now = Date.now()): string {
  const payload = Buffer.from(JSON.stringify({ kind, expiresAt: now + ttlMs, state })).toString("base64url");
  if (payload.length > maximumLength) throw new Error("GROUP_CURSOR_SIZE_LIMIT: Continuation exceeds the bounded cursor capacity.");
  const signature = createHmac("sha256", keyFor(config, true)).update(payload).digest("base64url");
  return `g1.${payload}.${signature}`;
}

export function decodeGroupCursor<T>(config: AppConfig, kind: "posts" | "attachments", cursor: string, now = Date.now()): T {
  if (typeof cursor !== "string" || cursor.length > maximumLength + 64) throw invalid();
  const match = /^g1\.([A-Za-z0-9_-]+)\.([A-Za-z0-9_-]{43})$/.exec(cursor);
  if (!match) throw invalid();
  const expected = createHmac("sha256", keyFor(config, false)).update(match[1]).digest();
  const actual = Buffer.from(match[2], "base64url");
  if (actual.length !== expected.length || !timingSafeEqual(actual, expected)) throw invalid();
  try {
    const envelope = JSON.parse(Buffer.from(match[1], "base64url").toString("utf8"));
    if (envelope.kind !== kind || !Number.isSafeInteger(envelope.expiresAt) || envelope.expiresAt <= now || envelope.expiresAt > now + ttlMs) throw invalid();
    return envelope.state as T;
  } catch { throw invalid(); }
}

function invalid(): Error {
  return new Error("GROUP_INVALID_CURSOR: Group continuation is invalid, modified, expired, or belongs to a different data directory. Restart the query; do not edit the cursor or authentication cache.");
}
