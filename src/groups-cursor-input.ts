import fs from "node:fs";
import { TextDecoder } from "node:util";

const maximumFileBytes = 32 * 1024 * 1024;
const maximumCursorLength = 3_000_064;

// Transport only: signature, expiry and query binding remain the reader's responsibility.
export function readGroupCursorInput(options: { cursor?: string; cursorFile?: string }): string | undefined {
  if (options.cursor !== undefined && options.cursorFile !== undefined) {
    throw new Error("GROUP_INVALID_CURSOR_INPUT: Use either --cursor or --cursor-file, not both.");
  }
  if (options.cursorFile === undefined) return options.cursor;
  let fd: number | undefined;
  try {
    fd = fs.openSync(options.cursorFile, "r");
    const stat = fs.fstatSync(fd);
    if (!stat.isFile() || stat.size < 1 || stat.size > maximumFileBytes) throw new Error();
    // Bound reads even if another process changes the file after fstat.
    const buffer = Buffer.alloc(stat.size + 1);
    let length = 0;
    while (length < buffer.length) {
      const count = fs.readSync(fd, buffer, length, buffer.length - length, null);
      if (!count) break;
      length += count;
    }
    if (length !== stat.size) throw new Error();
    const text = new TextDecoder("utf-8", { fatal: true }).decode(buffer.subarray(0, length)).trim();
    const cursor: unknown = text.startsWith("{") ? JSON.parse(text)?.page?.nextCursor : text;
    if (typeof cursor !== "string" || cursor.length > maximumCursorLength || !/^g1\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]{43}$/.test(cursor)) throw new Error();
    return cursor;
  } catch {
    throw new Error("GROUP_INVALID_CURSOR_INPUT: Expected a readable UTF-8 cursor or groups posts result JSON file (maximum 32 MiB) containing page.nextCursor. File contents and paths are not reported.");
  } finally {
    if (fd !== undefined) fs.closeSync(fd);
  }
}
