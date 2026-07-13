import assert from "node:assert/strict";
import test from "node:test";
import { isPersistentDataDir } from "../dist/config.js";

test("Linux OS-default data directories are not persistent in hosted execution", () => {
  assert.equal(
    isPersistentDataDir(
      "/root/.local/share/ohmyhotel/hare-m365-agent",
      "os-default",
      "linux"
    ),
    false
  );
  assert.equal(
    isPersistentDataDir(
      "/home/user/.local/share/ohmyhotel/hare-m365-agent",
      "os-default",
      "linux"
    ),
    false
  );
});

test("an explicit mounted Linux data directory can be persistent", () => {
  assert.equal(
    isPersistentDataDir("/mnt/hare/HareM365Agent", "command-line", "linux"),
    true
  );
  assert.equal(
    isPersistentDataDir("/mnt/hare/HareM365Agent", "environment", "linux"),
    true
  );
  assert.equal(
    isPersistentDataDir(
      "/sessions/example/mnt/HareM365Agent2",
      "command-line",
      "linux"
    ),
    true
  );
  assert.equal(
    isPersistentDataDir(
      "/sessions/example/mnt/company-selected-project",
      "command-line",
      "linux"
    ),
    true
  );
});

test("hosted temporary paths are rejected even when explicitly selected", () => {
  assert.equal(
    isPersistentDataDir("/sessions/example/HareM365Agent", "command-line", "linux"),
    false
  );
  assert.equal(
    isPersistentDataDir("/tmp/HareM365Agent", "environment", "linux"),
    false
  );
});

test("Windows and Mac OS-default data directories remain persistent", () => {
  assert.equal(
    isPersistentDataDir(
      "C:\\Users\\user\\AppData\\Local\\Ohmyhotel\\HareM365Agent",
      "os-default",
      "win32"
    ),
    true
  );
  assert.equal(
    isPersistentDataDir(
      "/Users/user/Library/Application Support/Ohmyhotel/HareM365Agent",
      "os-default",
      "darwin"
    ),
    true
  );
});
