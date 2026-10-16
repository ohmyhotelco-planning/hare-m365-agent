import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { encodeGroupCursor, decodeGroupCursor } from "../dist/groups-cursor.js";

function fixture(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "hare-group-cursor-"));
  t.after(() => {
    assert.equal(path.dirname(path.resolve(root)), path.resolve(os.tmpdir()));
    assert.match(path.basename(root), /^hare-group-cursor-/);
    fs.rmSync(root, { recursive: true, force: true });
  });
  return { cacheDir: path.join(root, "cache") };
}
test("group cursor authenticates the full payload and survives another call", t => {
  const config = fixture(t);
  const state = { group: "synthetic-group", offset: 10, missingBodies: 1 };
  const cursor = encodeGroupCursor(config, "posts", state, 1000);
  assert.deepEqual(decodeGroupCursor(config, "posts", cursor, 2000), state);
  const [prefix, data, signature] = cursor.split(".");
  const envelope = JSON.parse(Buffer.from(data, "base64url").toString());
  envelope.state.offset = 1000;
  const changed = `${prefix}.${Buffer.from(JSON.stringify(envelope)).toString("base64url")}.${signature}`;
  assert.throws(() => decodeGroupCursor(config, "posts", changed, 2000), /GROUP_INVALID_CURSOR/);
});
test("cursor kind, data directory and expiry cannot be substituted", t => {
  const config = fixture(t);
  const other = fixture(t);
  const cursor = encodeGroupCursor(config, "posts", {}, 1000);
  encodeGroupCursor(other, "posts", {}, 1000);
  assert.throws(() => decodeGroupCursor(other, "posts", cursor, 2000), /GROUP_INVALID_CURSOR/);
  assert.throws(() => decodeGroupCursor(config, "attachments", cursor, 2000), /GROUP_INVALID_CURSOR/);
  assert.throws(() => decodeGroupCursor(config, "posts", cursor, 1000 + 24 * 60 * 60 * 1000), /GROUP_INVALID_CURSOR/);
});
test("missing or corrupt traversal key fails closed and never resets authentication", t => {
  const config = fixture(t);
  const cursor = encodeGroupCursor(config, "posts", {});
  const keyPath = path.join(config.cacheDir, "group-continuation.key");
  fs.writeFileSync(keyPath, "invalid-test-key");
  assert.throws(() => decodeGroupCursor(config, "posts", cursor), /GROUP_INVALID_CURSOR/);
  assert.throws(() => encodeGroupCursor(config, "posts", {}), /GROUP_INVALID_CURSOR/);
  assert.equal(fs.readFileSync(keyPath, "utf8"), "invalid-test-key");
  assert.deepEqual(fs.readdirSync(config.cacheDir), ["group-continuation.key"]);
});
test("invalid envelopes are rejected without creating directories", t => {
  const config = fixture(t);
  for (const cursor of ["", "abc", "g1.abc.bad", "x".repeat(3_000_100)]) {
    assert.throws(() => decodeGroupCursor(config, "posts", cursor), /GROUP_INVALID_CURSOR/);
  }
  assert.equal(fs.existsSync(config.cacheDir), false);
});
