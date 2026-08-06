import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import {
  buildAttachmentContentPath,
  buildAttachmentListPath,
  downloadMessageAttachment,
  listMessageAttachments
} from "../dist/outlook-attachments.js";

test("Outlook attachment paths encode message and attachment IDs", () => {
  assert.equal(
    buildAttachmentListPath("message/id", 20),
    "/me/messages/message%2Fid/attachments?$select=id,name,contentType,size,isInline,lastModifiedDateTime&$top=20"
  );
  assert.equal(
    buildAttachmentContentPath("message/id", "attachment+id"),
    "/me/messages/message%2Fid/attachments/attachment%2Bid/$value"
  );
});

test("attachment listing returns file metadata without content bytes", async () => {
  const calls = [];
  const client = {
    async get(url) {
      calls.push(url);
      return {
        value: [
          {
            "@odata.type": "#microsoft.graph.fileAttachment",
            id: "attachment-1",
            name: "SCM report.xlsx",
            contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            size: 1024,
            isInline: false
          },
          {
            "@odata.type": "#microsoft.graph.referenceAttachment",
            id: "attachment-2",
            name: "Cloud link"
          }
        ]
      };
    },
    async download() {
      throw new Error("not used");
    }
  };

  const result = await listMessageAttachments(config(), "message-1", 20, client);
  assert.equal(calls.length, 1);
  assert.equal(result.list.returnedCount, 2);
  assert.equal(result.attachments[0].attachmentType, "file");
  assert.equal(result.attachments[0].downloadable, true);
  assert.equal(result.attachments[1].attachmentType, "reference");
  assert.equal(result.attachments[1].downloadable, false);
  assert.equal("contentBytes" in result.attachments[0], false);
});

test("file attachments download into Hare downloadsDir", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-outlook-attachment-"));
  const downloadDir = path.join(dataDir, "downloads");
  const calls = [];
  const client = {
    async get(url) {
      calls.push(url);
      return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        id: "attachment-1",
        name: "SCM report.xlsx",
        size: 5,
        isInline: false
      };
    },
    async download(url) {
      calls.push(url);
      return new Response(Buffer.from("hello"), {
        headers: { "content-length": "5" }
      });
    }
  };

  try {
    const result = await downloadMessageAttachment(
      config({ dataDir, downloadDir }),
      "message-1",
      "attachment-1",
      undefined,
      undefined,
      client
    );
    assert.equal(path.dirname(result.outputPath), downloadDir);
    assert.equal(path.basename(result.outputPath), "SCM report.xlsx");
    assert.equal(fs.readFileSync(result.outputPath, "utf8"), "hello");
    assert.match(calls[1], /\/\$value$/);
  } finally {
    fs.rmSync(dataDir, { recursive: true, force: true });
  }
});

test("reference attachments are not downloaded as mailbox files", async () => {
  const client = {
    async get() {
      return {
        "@odata.type": "#microsoft.graph.referenceAttachment",
        id: "reference-1",
        name: "Cloud link"
      };
    },
    async download() {
      throw new Error("download must not run");
    }
  };

  await assert.rejects(
    () => downloadMessageAttachment(config(), "message-1", "reference-1", undefined, undefined, client),
    /reference cannot be downloaded/
  );
});

test("large Outlook attachments return a preview without downloading content", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-outlook-large-preview-"));
  let downloadCalls = 0;
  const client = {
    async get() {
      return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        id: "attachment-large",
        name: "Accounting package.rar",
        size: 2 * 1024 * 1024,
        isInline: false
      };
    },
    async download() {
      downloadCalls += 1;
      throw new Error("download must not run before approval");
    }
  };

  try {
    const result = await downloadMessageAttachment(
      config({ dataDir, downloadDir: path.join(dataDir, "downloads") }),
      "message-1",
      "attachment-large",
      undefined,
      undefined,
      client
    );
    assert.equal(result.stage, "AWAITING_USER_APPROVAL");
    assert.equal(result.preview.source, "outlook");
    assert.equal(downloadCalls, 0);
  } finally {
    fs.rmSync(dataDir, { recursive: true, force: true });
  }
});

function config(overrides = {}) {
  const dataDir = overrides.dataDir ?? os.tmpdir();
  return {
    dataDir,
    cacheDir: overrides.cacheDir ?? path.join(dataDir, ".cache"),
    downloadDir: overrides.downloadDir ?? os.tmpdir(),
    policy: {
      allowDownloads: true,
      maxDownloadBytes: 1024 * 1024,
      maxApprovedDownloadBytes: 10 * 1024 * 1024,
      maxSearchResults: 1000
    }
  };
}
