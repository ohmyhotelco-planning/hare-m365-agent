import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { authorizeDownload } from "../dist/download-approval.js";
import { downloadDriveItem } from "../dist/sharepoint.js";

test("downloads at or below the default limit do not require approval", () => {
  const fixture = makeFixture();
  try {
    const result = authorizeDownload(fixture.config, plan(100));
    assert.deepEqual(result, { approved: true, maxBytes: 100, elevated: false });
    assert.equal(fs.existsSync(path.join(fixture.cacheDir, "download-approvals.json")), false);
  } finally {
    fixture.cleanup();
  }
});

test("an elevated approval token is exact, expiring, and single-use", () => {
  const fixture = makeFixture();
  try {
    const first = authorizeDownload(fixture.config, plan(608), undefined, 1_000);
    assert.equal(first.approved, false);
    assert.equal(first.pending.stage, "AWAITING_USER_APPROVAL");
    assert.equal(first.pending.preview.size, 608);
    const token = first.pending.approval.token;
    const stored = fs.readFileSync(path.join(fixture.cacheDir, "download-approvals.json"), "utf8");
    assert.equal(stored.includes(token), false);

    const approved = authorizeDownload(fixture.config, plan(608), token, 2_000);
    assert.deepEqual(approved, { approved: true, maxBytes: 1000, elevated: true });
    assert.throws(
      () => authorizeDownload(fixture.config, plan(608), token, 2_001),
      /missing, expired, or already used/
    );

    const expiring = authorizeDownload(fixture.config, plan(609), undefined, 10_000);
    assert.equal(expiring.approved, false);
    assert.throws(
      () => authorizeDownload(fixture.config, plan(609), expiring.pending.approval.token, 610_001),
      /missing, expired, or already used/
    );
  } finally {
    fixture.cleanup();
  }
});

test("approval is rejected when the file plan changes or exceeds the hard cap", () => {
  const fixture = makeFixture();
  try {
    const first = authorizeDownload(fixture.config, plan(608), undefined, 1_000);
    assert.equal(first.approved, false);
    assert.throws(
      () => authorizeDownload(
        fixture.config,
        { ...plan(608), outputName: "different.rar" },
        first.pending.approval.token,
        2_000
      ),
      /approved file does not match/
    );
    assert.throws(
      () => authorizeDownload(fixture.config, plan(1001), undefined, 2_000),
      /approved policy limit/
    );
  } finally {
    fixture.cleanup();
  }
});

test("an invalid approved limit fails closed", () => {
  const fixture = makeFixture();
  try {
    fixture.config.policy.maxApprovedDownloadBytes = 50;
    assert.throws(
      () => authorizeDownload(fixture.config, plan(60)),
      /maxApprovedDownloadBytes must be at least maxDownloadBytes/
    );
  } finally {
    fixture.cleanup();
  }
});

test("SharePoint large-file preview performs no content download before approval", async () => {
  const fixture = makeFixture();
  let downloadCalls = 0;
  const client = {
    async get() {
      return { id: "item-1", name: "Accounting package.rar", size: 608, file: {} };
    },
    async download() {
      downloadCalls += 1;
      return new Response(Buffer.from("payload"), { headers: { "content-length": "7" } });
    }
  };

  try {
    const preview = await downloadDriveItem(
      fixture.config,
      "drive-1",
      "item-1",
      undefined,
      undefined,
      client
    );
    assert.equal(preview.stage, "AWAITING_USER_APPROVAL");
    assert.equal(downloadCalls, 0);

    const outputPath = await downloadDriveItem(
      fixture.config,
      "drive-1",
      "item-1",
      undefined,
      preview.approval.token,
      client
    );
    assert.equal(typeof outputPath, "string");
    assert.equal(downloadCalls, 1);
    assert.equal(fs.readFileSync(outputPath, "utf8"), "payload");
  } finally {
    fixture.cleanup();
  }
});

function plan(size) {
  return {
    source: "sharepoint",
    resourceKey: "drive-1\0item-1",
    name: "Accounting package.rar",
    size,
    outputName: "Accounting package.rar"
  };
}

function makeFixture() {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-download-approval-"));
  const cacheDir = path.join(dataDir, ".cache");
  const downloadDir = path.join(dataDir, "downloads");
  fs.mkdirSync(cacheDir, { recursive: true });
  return {
    cacheDir,
    config: {
      dataDir,
      cacheDir,
      downloadDir,
      policy: {
        allowDownloads: true,
        maxDownloadBytes: 100,
        maxApprovedDownloadBytes: 1000
      }
    },
    cleanup: () => fs.rmSync(dataDir, { recursive: true, force: true })
  };
}
