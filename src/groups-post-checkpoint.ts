import { createHash, randomUUID } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import type { AppConfig } from "./config.js";
import { decodeGroupCursor, encodeGroupCursor } from "./groups-cursor.js";

const chunkSize = 1024;
const digestPattern = /^[a-f0-9]{64}$/;
const maximumObjectBytes = 16 * 1024 * 1024;
const maximumStoreBytes = 256 * 1024 * 1024;
export const maximumHistoryEntries = 500_000;
type History = { chunks: string[]; tail: string[]; count: number };
type State = { criteria: string; seen: string[]; pages: string[] };

// Immutable, content-addressed chunks preserve old cursors without duplicating full histories.
// Only identifiers/hashes and traversal metadata are stored, never message bodies or subjects.
class Store {
  readonly root: string;
  private bytes?: number;
  constructor(config: AppConfig) {
    if (!path.isAbsolute(config.cacheDir)) throw unavailable();
    this.root = path.join(config.cacheDir, "group-post-checkpoints");
    if (fs.existsSync(this.root) && fs.lstatSync(this.root).isSymbolicLink()) throw unavailable();
  }
  read(id: string): unknown {
    if (!digestPattern.test(id)) throw invalid();
    let fd: number | undefined;
    try {
      const target = path.join(this.root, `${id}.json`);
      if (fs.lstatSync(target).isSymbolicLink()) throw invalid();
      fd = fs.openSync(target, "r");
      const stat = fs.fstatSync(fd);
      if (!stat.isFile() || stat.size < 1 || stat.size > maximumObjectBytes) throw invalid();
      const data = Buffer.alloc(stat.size + 1);
      let length = 0;
      while (length < data.length) {
        const n = fs.readSync(fd, data, length, data.length - length, null);
        if (!n) break;
        length += n;
      }
      if (length !== stat.size || digest(data.subarray(0, length)) !== id) throw invalid();
      return JSON.parse(data.subarray(0, length).toString("utf8"));
    } catch { throw invalid(); }
    finally { if (fd !== undefined) fs.closeSync(fd); }
  }
  write(value: unknown): string {
    const data = Buffer.from(JSON.stringify(value));
    if (data.length > maximumObjectBytes) throw unavailable();
    const id = digest(data);
    const target = path.join(this.root, `${id}.json`);
    if (fs.existsSync(target)) { this.read(id); return id; }
    try {
      fs.mkdirSync(this.root, { recursive: true, mode: 0o700 });
      if (this.bytes === undefined) {
        const names = fs.readdirSync(this.root);
        if (names.length > 20_000) throw unavailable();
        this.bytes = 0;
        for (const name of names) {
          const stat = fs.lstatSync(path.join(this.root, name));
          if (!stat.isFile() || stat.isSymbolicLink()) throw unavailable();
          this.bytes += stat.size;
        }
      }
      if (this.bytes + data.length > maximumStoreBytes) throw unavailable();
      // Publish only a complete object. Interrupted temporary files remain quota-counted.
      const temporary = path.join(this.root, `${randomUUID()}.tmp`);
      const fd = fs.openSync(temporary, "wx", 0o600);
      try { fs.writeFileSync(fd, data); fs.fsyncSync(fd); }
      finally { fs.closeSync(fd); }
      if (digest(fs.readFileSync(temporary)) !== id) throw unavailable();
      fs.renameSync(temporary, target);
      this.read(id);
      this.bytes += data.length;
      return id;
    } catch { throw unavailable(); }
  }
}

export function encodePostContinuation<T extends State>(config: AppConfig, state: T): string {
  if (state.seen.length + state.pages.length < 8192) {
    try { return encodeGroupCursor(config, "posts", state); }
    catch (error) {
      if (!(error instanceof Error) || !error.message.startsWith("GROUP_CURSOR_SIZE_LIMIT:")) throw error;
    }
  }
  const store = new Store(config);
  const pack = (values: string[]): History => {
    if (values.length > maximumHistoryEntries || !values.every(v => digestPattern.test(v))) throw unavailable();
    const chunks: string[] = [];
    let offset = 0;
    while (values.length - offset >= chunkSize) {
      chunks.push(store.write(values.slice(offset, offset + chunkSize)));
      offset += chunkSize;
    }
    return { chunks, tail: values.slice(offset), count: values.length };
  };
  const checkpoint = store.write({ ...state, seen: pack(state.seen), pages: pack(state.pages) });
  return encodeGroupCursor(config, "posts", { version: 4, criteria: state.criteria, checkpoint });
}

export function decodePostContinuation<T extends State>(config: AppConfig, cursor: string, criteria: string): { state: T; local: boolean } {
  const envelope = decodeGroupCursor<Record<string, unknown>>(config, "posts", cursor);
  if (envelope?.version !== 4) return { state: envelope as T, local: false };
  if (envelope.criteria !== criteria || typeof envelope.checkpoint !== "string" ||
      Object.keys(envelope).some(k => !["version", "criteria", "checkpoint"].includes(k))) throw invalid();
  const store = new Store(config);
  const stored = store.read(envelope.checkpoint) as Record<string, unknown>;
  if (!stored || stored.criteria !== criteria) throw invalid();
  const unpack = (raw: unknown): string[] => {
    const h = raw as History;
    if (!h || !Array.isArray(h.chunks) || !Array.isArray(h.tail) || h.tail.length >= chunkSize ||
        !Number.isSafeInteger(h.count) || h.count < 0 || h.count > maximumHistoryEntries ||
        h.count !== h.chunks.length * chunkSize + h.tail.length ||
        Object.keys(h).some(k => !["chunks", "tail", "count"].includes(k)) ||
        ![...h.chunks, ...h.tail].every(v => typeof v === "string" && digestPattern.test(v))) throw invalid();
    const values: string[] = [];
    for (const id of h.chunks) {
      const chunk = store.read(id);
      if (!Array.isArray(chunk) || chunk.length !== chunkSize || !chunk.every(v => typeof v === "string" && digestPattern.test(v))) throw invalid();
      values.push(...chunk);
    }
    return values.concat(h.tail);
  };
  return { state: { ...stored, seen: unpack(stored.seen), pages: unpack(stored.pages) } as T, local: true };
}
function digest(value: Buffer): string { return createHash("sha256").update(value).digest("hex"); }
function invalid(): Error { return new Error("GROUP_INVALID_CHECKPOINT: Local traversal state is missing, changed or invalid. Preserve the cursor and state; do not edit them or the authentication cache."); }
function unavailable(): Error { return new Error("GROUP_CHECKPOINT_STORAGE_BLOCKED: Local traversal storage is unavailable or exceeds its resource bounds (16 MiB/object, 256 MiB/store, 500000 identities/set). No progress result was committed; preserve the previous cursor. Do not delete authentication files."); }
