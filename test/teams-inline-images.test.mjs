import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import {
  buildInlineImageContentPath,
  buildInlineImageListPath,
  downloadTeamsInlineImage,
  listTeamsInlineImages
} from "../dist/teams-inline-images.js";

test("Teams inline image paths encode every identifier", () => {
  assert.equal(
    buildInlineImageListPath("chat/id", "message/id"),
    "/chats/chat%2Fid/messages/message%2Fid/hostedContents"
  );
  assert.equal(
    buildInlineImageContentPath("chat/id", "message/id", "hosted/id"),
    "/chats/chat%2Fid/messages/message%2Fid/hostedContents/hosted%2Fid/$value"
  );
});

test("inline image listing returns metadata without content bytes", async () => {
  const result = await listTeamsInlineImages(config(), "chat-1", "message-1", 20, {
    async get() {
      return {
        value: [
          { id: "image-1", contentBytes: null, contentType: null },
          { id: "content-2", contentBytes: null, contentType: null }
        ]
      };
    },
    async download() { throw new Error("not used"); }
  });
  assert.deepEqual(result.images, [
    { id: "image-1", verificationRequired: true },
    { id: "content-2", verificationRequired: true }
  ]);
  assert.equal("contentBytes" in result.images[0], false);
});

test("inline images download into Hare downloadsDir with the default size limit", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-teams-inline-"));
  const downloadDir = path.join(dataDir, "downloads");
  const calls = [];
  try {
    const result = await downloadTeamsInlineImage(
      config({ dataDir, downloadDir }),
      "chat-1",
      "message-1",
      "image-1",
      undefined,
      {
        async get() { throw new Error("download does not fetch unreadable metadata"); },
        async download(url) {
          calls.push(url);
          return new Response(Buffer.from("image"), {
            headers: { "content-length": "5", "content-type": "image/png" }
          });
        }
      }
    );
    assert.equal(result.stage, "DOWNLOADED");
    assert.equal(path.dirname(result.outputPath), downloadDir);
    assert.equal(fs.readFileSync(result.outputPath, "utf8"), "image");
    assert.match(calls.at(-1), /hostedContents\/image-1\/\$value$/);
  } finally {
    fs.rmSync(dataDir, { recursive: true, force: true });
  }
});

test("non-image hosted content is rejected from the download response", async () => {
  let downloads = 0;
  await assert.rejects(
    () => downloadTeamsInlineImage(config(), "chat-1", "message-1", "content-1", undefined, {
      async get() { throw new Error("not used"); },
      async download() {
        downloads += 1;
        return new Response(Buffer.from("not-an-image"), {
          headers: { "content-type": "text/plain" }
        });
      }
    }),
    /not an inline image/
  );
  assert.equal(downloads, 1);
});

test("inline image streams cannot exceed the default download limit", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-teams-inline-limit-"));
  try {
    await assert.rejects(
      () => downloadTeamsInlineImage(
        config({ dataDir, downloadDir: path.join(dataDir, "downloads"), maxDownloadBytes: 4 }),
        "chat-1",
        "message-1",
        "image-1",
        undefined,
        {
          async get() { throw new Error("not used"); },
          async download() {
            return new Response(Buffer.from("image"), {
              headers: { "content-length": "5", "content-type": "image/png" }
            });
          }
        }
      ),
      /exceeds policy limit/
    );
  } finally {
    fs.rmSync(dataDir, { recursive: true, force: true });
  }
});

test("active or unknown image formats are rejected before saving", async () => {
  for (const contentType of ["image/svg+xml", "image/unknown", undefined]) {
    let cancelled = false;
    await assert.rejects(
      () => downloadTeamsInlineImage(config(), "chat-1", "message-1", "content-1", undefined, {
        async get() { throw new Error("not used"); },
        async download() {
          return {
            headers: { get: (name) => name.toLowerCase() === "content-type" ? contentType ?? null : null },
            body: { async cancel() { cancelled = true; } }
          };
        }
      }),
      /not an inline image/
    );
    assert.equal(cancelled, true);
  }
});

test("a user-provided extension is replaced with the verified raster extension", async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-teams-inline-name-"));
  try {
    const result = await downloadTeamsInlineImage(
      config({ dataDir, downloadDir: path.join(dataDir, "downloads") }),
      "chat-1",
      "message-1",
      "image-1",
      "report.html",
      {
        async get() { throw new Error("not used"); },
        async download() {
          return new Response(Buffer.from("image"), {
            headers: { "content-length": "5", "content-type": "image/jpeg" }
          });
        }
      }
    );
    assert.equal(path.basename(result.outputPath), "report.jpg");
  } finally {
    fs.rmSync(dataDir, { recursive: true, force: true });
  }
});
test("download policy blocks inline images before any network request", async () => {
  let requests = 0;
  await assert.rejects(
    () => downloadTeamsInlineImage(
      config({ allowDownloads: false }),
      "chat-1",
      "message-1",
      "image-1",
      undefined,
      {
        async get() { requests += 1; throw new Error("not used"); },
        async download() { requests += 1; throw new Error("not used"); }
      }
    ),
    /disabled by policy/
  );
  assert.equal(requests, 0);
});

function config(overrides = {}) {
  const dataDir = overrides.dataDir ?? os.tmpdir();
  return {
    dataDir,
    cacheDir: path.join(dataDir, ".cache"),
    downloadDir: overrides.downloadDir ?? os.tmpdir(),
    policy: {
      allowDownloads: overrides.allowDownloads ?? true,
      maxDownloadBytes: overrides.maxDownloadBytes ?? 1024 * 1024
    }
  };
}
