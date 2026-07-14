import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import {
  clearStoredFile,
  storedFileHasContent,
  usesDeleteRestrictedStorage,
  writeStoredText
} from "../dist/persistent-storage.js";

test("Cowork project mounts are classified as delete-restricted storage", () => {
  assert.equal(
    usesDeleteRestrictedStorage("/sessions/example/mnt/HareM365Agent/.cache/msal-cache.json"),
    true
  );
  assert.equal(
    usesDeleteRestrictedStorage("/sessions/example/hare/.cache/msal-cache.json"),
    false
  );
});

test("delete-restricted storage overwrites and clears files without removing them", () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "hare-delete-restricted-"));
  const file = path.join(directory, "msal-cache.json");

  writeStoredText(file, "first", { deleteRestricted: true });
  writeStoredText(file, "second", { deleteRestricted: true });
  assert.equal(fs.readFileSync(file, "utf8"), "second");
  assert.equal(storedFileHasContent(file), true);

  clearStoredFile(file, { deleteRestricted: true });
  assert.equal(fs.existsSync(file), true, "clear must not unlink a Cowork-mounted file");
  assert.equal(fs.readFileSync(file, "utf8"), "");
  assert.equal(storedFileHasContent(file), false);
});
