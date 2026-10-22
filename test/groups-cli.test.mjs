import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { pathToFileURL } from "node:url";

test("group CLI help exposes read-only commands and scoped identifiers offline", () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "hare-groups-help-"));
  const dataDir = path.join(root, "unused");
  try {
    for (const [args, expected] of [
      [["groups"], /resolve[\s\S]*posts[\s\S]*attachments/],
      [["groups", "resolve"], /--group/],
      [["groups", "posts"], /--since[\s\S]*--until[\s\S]*--limit[\s\S]*--cursor[\s\S]*--cursor-file[\s\S]*--include-subject/],
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
  assert.match(rules, /groups posts --include-subject/);
  assert.match(rules, /SUBJECT_BUDGET retains the pending post for continuation/);
  assert.match(rules, /Missing or invalid properties remain explicit and prevent page.complete/);
  const cli = fs.readFileSync("src/cli.ts", "utf8");
  assert.match(cli, /groups posts --include-subject/);
  assert.match(cli, /PRESENT의 value로만 제목 조건을 판단/);
});

test("native CLI file transport forwards long cursors and rejects conflicting sources before group lookup", t => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "hare-groups-cli-input-"));
  t.after(() => {
    assert.equal(path.dirname(path.resolve(root)), path.resolve(os.tmpdir()));
    assert.match(path.basename(root), /^hare-groups-cli-input-/);
    fs.rmSync(root, { recursive: true, force: true });
  });
  // Only the network-facing actions are replaced; the real CLI parser/input reader runs.
  const loader = path.join(root, "offline-loader.mjs");
  fs.writeFileSync(loader, `export async function load(url, context, next) {
    if (url.endsWith('/dist/groups-directory.js')) return {format:'module',shortCircuit:true,source:
      'export async function resolveGroup(){ if(process.env.TEST_DENY_DIRECTORY) throw new Error("UNEXPECTED_DIRECTORY_LOOKUP"); return {}; }'};
    if (url.endsWith('/dist/groups-posts.js')) return {format:'module',shortCircuit:true,source:
      'export async function listGroupPosts(c,g,l,o){return {receivedCursor:o.cursor,includeSubject:o.includeSubject};}'};
    return next(url, context);
  }`);
  const preload = `import {register} from 'node:module'; register(${JSON.stringify(pathToFileURL(loader).href)});`;
  const file = path.join(root, "previous.json");
  const cursor = `g1.${"a".repeat(40_000)}.${"b".repeat(43)}`;
  const run = (args = [], deny = false) => spawnSync(process.execPath, ["--import", `data:text/javascript,${encodeURIComponent(preload)}`,
    "dist/cli.js", "--data-dir", path.join(root, "data"), "groups", "posts", "--group", "synthetic@example.test", "--cursor-file", file, ...args], {
    encoding: "utf8", timeout: 15000,
    env: { ...process.env, OMH_M365_CLIENT_ID: "synthetic-client", OMH_M365_TENANT_ID: "synthetic-tenant", TEST_DENY_DIRECTORY: deny ? "1" : "" }
  });
  for (const value of [cursor, JSON.stringify({ page: { nextCursor: cursor } })]) {
    fs.writeFileSync(file, value);
    const result = run();
    assert.equal(result.status, 0, result.stderr);
    assert.equal(JSON.parse(result.stdout).receivedCursor, cursor);
  }
  const enriched = run(["--include-subject"]);
  assert.equal(enriched.status, 0, enriched.stderr);
  assert.equal(JSON.parse(enriched.stdout).includeSubject, true);
  const rejected = run(["--cursor", "inline"], true);
  assert.notEqual(rejected.status, 0);
  assert.match(rejected.stderr, /GROUP_INVALID_CURSOR_INPUT/);
  assert.doesNotMatch(rejected.stderr, /UNEXPECTED_DIRECTORY_LOOKUP/);
});
