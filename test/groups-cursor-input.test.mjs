import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { readGroupCursorInput } from "../dist/groups-cursor-input.js";
import { encodeGroupCursor, decodeGroupCursor } from "../dist/groups-cursor.js";

function fixture(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "hare-cursor-input-"));
  t.after(() => {
    assert.equal(path.dirname(path.resolve(root)), path.resolve(os.tmpdir()));
    assert.match(path.basename(root), /^hare-cursor-input-/);
    fs.rmSync(root, { recursive: true, force: true });
  });
  return { root, file: path.join(root, "prior-page.json"), config: { cacheDir: path.join(root, "cache") } };
}

test("cursor input accepts absent/inline transport and rejects two sources", () => {
  assert.equal(readGroupCursorInput({}), undefined);
  assert.equal(readGroupCursorInput({ cursor: "unchanged" }), "unchanged");
  assert.throws(() => readGroupCursorInput({ cursor: "", cursorFile: "unused" }), /Use either/);
});

test("long signed cursor survives raw UTF-8 BOM/CRLF and previous result JSON unchanged", t => {
  const { file, config } = fixture(t);
  const state = { synthetic: "x".repeat(40_000) };
  const cursor = encodeGroupCursor(config, "posts", state);
  assert.ok(cursor.length > 32_767);
  for (const contents of ["\ufeff" + cursor + "\r\n", JSON.stringify({ posts: [], page: { nextCursor: cursor } })]) {
    fs.writeFileSync(file, contents);
    const read = readGroupCursorInput({ cursorFile: file });
    assert.equal(read, cursor);
    assert.deepEqual(decodeGroupCursor(config, "posts", read), state);
  }
});

test("file transport does not bypass signature or cursor-kind verification", t => {
  const { file, config } = fixture(t);
  const cursor = encodeGroupCursor(config, "posts", { synthetic: true });
  fs.writeFileSync(file, cursor);
  assert.throws(() => decodeGroupCursor(config, "attachments", readGroupCursorInput({ cursorFile: file })), /GROUP_INVALID_CURSOR/);
  const parts = cursor.split(".");
  parts[1] = Buffer.from(JSON.stringify({ altered: true })).toString("base64url");
  fs.writeFileSync(file, parts.join("."));
  assert.throws(() => decodeGroupCursor(config, "posts", readGroupCursorInput({ cursorFile: file })), /GROUP_INVALID_CURSOR/);
});

test("expired signed cursors remain rejected through both file formats", t => {
  const { file, config } = fixture(t);
  const cursor = encodeGroupCursor(config, "posts", {}, 1000);
  for (const contents of [cursor, JSON.stringify({ page: { nextCursor: cursor } })]) {
    fs.writeFileSync(file, contents);
    assert.throws(() => decodeGroupCursor(config, "posts", readGroupCursorInput({ cursorFile: file })), /GROUP_INVALID_CURSOR/);
  }
});

test("malformed, missing, non-file, oversized and non-UTF8 inputs fail without echoing content or paths", t => {
  const { root, file } = fixture(t);
  const secret = "synthetic-sensitive-marker";
  const check = input => assert.throws(() => readGroupCursorInput({ cursorFile: input }), error => {
    assert.match(error.message, /^GROUP_INVALID_CURSOR_INPUT:/);
    assert.equal(error.message.includes(root), false);
    assert.equal(error.message.includes(secret), false);
    return true;
  });
  check(path.join(root, "missing"));
  check(root);
  for (const value of ["", secret, "{", "{}", '{"page":{"nextCursor":42}}', '{"page":{"continuationAvailable":false}}', Buffer.from([0xff, 0xfe, 0x61])]) {
    fs.writeFileSync(file, value);
    check(file);
  }
  fs.writeFileSync(file, "");
  fs.truncateSync(file, 32 * 1024 * 1024 + 1);
  check(file);
  fs.writeFileSync(file, `g1.${"a".repeat(3_000_064)}.${"b".repeat(43)}`);
  check(file);
});
