import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";

test("group CLI help exposes read-only commands and scoped identifiers offline", () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "hare-groups-help-"));
  const dataDir = path.join(root, "unused");
  try {
    for (const [args, expected] of [
      [["groups"], /resolve[\s\S]*posts[\s\S]*attachments/],
      [["groups", "resolve"], /--group/],
      [["groups", "posts"], /--since[\s\S]*--until[\s\S]*--limit[\s\S]*--cursor/],
      [["groups", "attachments", "list"], /--group[\s\S]*--thread-id[\s\S]*--post-id/],
      [["groups", "attachments", "download"], /--attachment-id[\s\S]*--name[\s\S]*--approval-token/]
    ]) {
      const result = spawnSync(process.execPath, ["dist/cli.js", "--data-dir", dataDir, ...args, "--help"], { encoding: "utf8", timeout: 10000 });
      assert.equal(result.status, 0, result.stderr);
      assert.match(result.stdout, expected);
    }
    assert.equal(fs.existsSync(dataDir), false, "help must not create authentication/data files");
  } finally { fs.rmSync(root, { recursive: true, force: true }); }
});

test("group reads use the approved read scope without automatic permission escalation", () => {
  const auth = fs.readFileSync("src/auth.ts", "utf8");
  assert.match(auth, /"Group\.Read\.All"/);
  assert.doesNotMatch(auth, /"Group(?:-Conversation)?\.ReadWrite\.All"/);
  const rules = fs.readFileSync("src/session-rules.ts", "utf8");
  assert.match(rules, /groups resolve --group/);
  assert.match(rules, /Do not expand scopes, consent or restart sign-in automatically/);
});
