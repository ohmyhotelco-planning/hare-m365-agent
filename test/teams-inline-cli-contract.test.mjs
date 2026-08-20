import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import test, { after } from "node:test";

const fixtureRoot = fs.mkdtempSync(path.join(os.homedir(), "hare-inline-cli-"));
after(() => fs.rmSync(fixtureRoot, { recursive: true, force: true }));
const cli = path.resolve("dist/cli.js");

function run(args) {
  const dataDir = fs.mkdtempSync(path.join(fixtureRoot, "data-"));
  return spawnSync(process.execPath, [cli, "--data-dir", dataDir, ...args], {
    cwd: path.dirname(cli),
    encoding: "utf8"
  });
}

test("Teams exposes inline image list and download commands", () => {
  const result = run(["teams", "inline-images", "--help"]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /list/);
  assert.match(result.stdout, /download/);

  const download = run(["teams", "inline-images", "download", "--help"]);
  assert.equal(download.status, 0, download.stderr);
  assert.match(download.stdout, /--hosted-content-id/);
  assert.match(download.stdout, /--chat-id/);
  assert.match(download.stdout, /--message-id/);
});
