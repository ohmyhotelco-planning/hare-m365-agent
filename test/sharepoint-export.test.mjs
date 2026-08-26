import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import test from "node:test";
import { finalizeTemporaryDownload } from "../dist/downloads.js";
import { exportSharePointFiles } from "../dist/sharepoint-export.js";

test("SharePoint export previews the full paged tree and downloads after one approval", async () => {
  const fixture = makeFixture(10);
  fs.writeFileSync(path.join(fixture.destination, "existing.txt"), "old", "utf8");
  const graph = makeGraph({
    rootPages: [
      {
        value: [
          { id: "folder-1", name: "Reports", folder: { childCount: 1 } },
          { id: "file-1", name: "summary.txt", size: 3, eTag: "etag-1", file: {} },
          { id: "large-1", name: "large.bin", size: 11, eTag: "etag-large", file: {} }
        ],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/mock/root-page-2"
      },
      {
        value: [
          { id: "existing-1", name: "existing.txt", size: 3, eTag: "etag-existing", file: {} }
        ]
      }
    ],
    folders: {
      "folder-1": [
        { id: "nested-1", name: "detail.txt", size: 4, eTag: "etag-2", file: {} }
      ]
    },
    payloads: {
      "file-1": "abc",
      "nested-1": "defg",
      "existing-1": "new"
    }
  });

  try {
    const preview = await runExport(fixture, graph.client);
    assert.equal(preview.stage, "AWAITING_USER_APPROVAL");
    assert.equal(preview.preview.fileCount, 4);
    assert.equal(preview.preview.folderCount, 1);
    assert.equal(preview.preview.pendingDownloadCount, 2);
    assert.equal(preview.preview.existingFileCount, 1);
    assert.equal(preview.preview.oversizedFileCount, 1);
    assert.equal(graph.downloadCalls.length, 0);
    assert.equal(graph.getCalls.some((url) => url.includes("root-page-2")), true);
    assert.equal(graph.getCalls.some((url) => url.includes("items/folder-1/children")), true);

    const storedApproval = fs.readFileSync(
      path.join(fixture.cacheDir, "sharepoint-export-approvals.json"),
      "utf8"
    );
    assert.equal(storedApproval.includes(preview.approval.token), false);

    const result = await runExport(fixture, graph.client, {
      approvalToken: preview.approval.token
    });
    assert.equal(result.stage, "COMPLETE");
    assert.equal(result.downloadedCount, 2);
    assert.equal(result.skippedExistingCount, 1);
    assert.equal(result.skippedOversizedCount, 1);
    assert.equal(fs.readFileSync(path.join(fixture.destination, "summary.txt"), "utf8"), "abc");
    assert.equal(
      fs.readFileSync(path.join(fixture.destination, "Reports", "detail.txt"), "utf8"),
      "defg"
    );
    assert.equal(fs.readFileSync(path.join(fixture.destination, "existing.txt"), "utf8"), "old");

    const rerun = await runExport(fixture, graph.client);
    assert.equal(rerun.stage, "COMPLETE");
    assert.equal(rerun.downloadedCount, 0);
    assert.equal(rerun.skippedExistingCount, 3);
  } finally {
    fixture.cleanup();
  }
});

test("SharePoint export resumes a partial file with Range under the approved manifest", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({
    rootPages: [
      {
        value: [
          { id: "file-1", name: "package.bin", size: 6, eTag: "etag-1", file: {} }
        ]
      }
    ],
    payloads: { "file-1": "abcdef" }
  });

  try {
    const preview = await runExport(fixture, graph.client);
    assert.equal(preview.stage, "AWAITING_USER_APPROVAL");
    fs.writeFileSync(
      path.join(fixture.destination, `package.bin.hare-part-${preview.jobId}`),
      "ab",
      "utf8"
    );

    const result = await runExport(fixture, graph.client, {
      approvalToken: preview.approval.token
    });
    assert.equal(result.stage, "COMPLETE");
    assert.equal(result.resumedFileCount, 1);
    assert.equal(graph.downloadCalls[0].options.rangeStart, 2);
    assert.equal(graph.downloadCalls[0].options.eTag, "etag-1");
    assert.equal(fs.readFileSync(path.join(fixture.destination, "package.bin"), "utf8"), "abcdef");
  } finally {
    fixture.cleanup();
  }
});

test("SharePoint export safely restarts when the server ignores a Range request", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({
    rootPages: [
      {
        value: [
          { id: "file-1", name: "package.bin", size: 6, eTag: "etag-1", file: {} }
        ]
      }
    ],
    payloads: { "file-1": "abcdef" }
  });
  let requestedOffset;

  try {
    const preview = await runExport(fixture, graph.client);
    fs.writeFileSync(
      path.join(fixture.destination, `package.bin.hare-part-${preview.jobId}`),
      "ab",
      "utf8"
    );
    graph.client.download = async (_url, options = {}) => {
      requestedOffset = options.rangeStart;
      return new Response(Buffer.from("abcdef"), {
        status: 200,
        headers: { "content-length": "6" }
      });
    };

    const result = await runExport(fixture, graph.client, {
      approvalToken: preview.approval.token
    });
    assert.equal(result.stage, "COMPLETE");
    assert.equal(requestedOffset, 2);
    assert.equal(fs.readFileSync(path.join(fixture.destination, "package.bin"), "utf8"), "abcdef");
  } finally {
    fixture.cleanup();
  }
});


test("SharePoint export rejects a mismatched Content-Range", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({
    rootPages: [
      {
        value: [
          { id: "file-1", name: "package.bin", size: 6, eTag: "etag-1", file: {} }
        ]
      }
    ],
    payloads: { "file-1": "abcdef" }
  });

  try {
    const preview = await runExport(fixture, graph.client);
    fs.writeFileSync(
      path.join(fixture.destination, `package.bin.hare-part-${preview.jobId}`),
      "ab",
      "utf8"
    );
    graph.client.download = async () =>
      new Response(Buffer.from("def"), {
        status: 206,
        headers: {
          "content-length": "3",
          "content-range": "bytes 3-5/6"
        }
      });

    const result = await runExport(fixture, graph.client, {
      approvalToken: preview.approval.token
    });
    assert.equal(result.stage, "COMPLETE_WITH_ERRORS");
    assert.match(result.failures[0].error, /invalid Content-Range/);
    assert.equal(fs.existsSync(path.join(fixture.destination, "package.bin")), false);
  } finally {
    fixture.cleanup();
  }
});

test("SharePoint enumeration checkpoints Graph pages before the time budget expires", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({
    rootPages: [
      {
        value: [
          { id: "file-1", name: "one.txt", size: 1, eTag: "etag-1", file: {} }
        ],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/mock/root-page-2"
      },
      {
        value: [
          { id: "file-2", name: "two.txt", size: 1, eTag: "etag-2", file: {} }
        ]
      }
    ],
    payloads: { "file-1": "1", "file-2": "2" }
  });

  try {
    const ticks = [0, 0, 0, 0, 10];
    const first = await runExport(fixture, graph.client, {
      timeBudgetMs: 5,
      now: () => ticks.shift() ?? 10
    });
    assert.equal(first.stage, "ENUMERATION_IN_PROGRESS");
    assert.equal(first.discoveredFileCount, 0);

    const second = await runExport(fixture, graph.client);
    assert.equal(second.stage, "AWAITING_USER_APPROVAL");
    assert.equal(second.preview.fileCount, 2);
    assert.equal(
      graph.getCalls.filter((url) => url.includes("/root/children")).length,
      1
    );
    assert.equal(
      graph.getCalls.filter((url) => url.includes("root-page-2")).length,
      1
    );
  } finally {
    fixture.cleanup();
  }
});


test("SharePoint export keeps batch approval active across a time-budget continuation", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({
    rootPages: [
      {
        value: [
          { id: "file-1", name: "one.txt", size: 1, eTag: "etag-1", file: {} },
          { id: "file-2", name: "two.txt", size: 1, eTag: "etag-2", file: {} }
        ]
      }
    ],
    payloads: { "file-1": "1", "file-2": "2" }
  });

  try {
    const preview = await runExport(fixture, graph.client);
    let clock = 0;
    const first = await runExport(fixture, graph.client, {
      approvalToken: preview.approval.token,
      timeBudgetMs: 3,
      now: () => {
        const value = clock;
        clock += 2;
        return value;
      }
    });
    assert.equal(first.stage, "IN_PROGRESS");
    assert.equal(first.downloadedCount, 1);

    const resumed = await runExport(fixture, graph.client);
    assert.equal(resumed.stage, "COMPLETE");
    assert.equal(resumed.approvalReused, true);
    assert.equal(resumed.downloadedCount, 1);
    assert.equal(resumed.skippedExistingCount, 1);
  } finally {
    fixture.cleanup();
  }
});

test("SharePoint export approval is bound to the logged-in account", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({
    rootPages: [
      { value: [{ id: "file-1", name: "one.txt", size: 3, eTag: "etag-1", file: {} }] }
    ],
    payloads: { "file-1": "one" }
  });

  try {
    const preview = await runExport(fixture, graph.client);
    await assert.rejects(
      () =>
        runExport(fixture, graph.client, {
          accountId: "account-2",
          approvalToken: preview.approval.token
        }),
      /approved export plan no longer matches/
    );
  } finally {
    fixture.cleanup();
  }
});

test("SharePoint source changes invalidate the approved checkpoint", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({
    rootPages: [
      { value: [{ id: "file-1", name: "one.txt", size: 3, eTag: "etag-1", file: {} }] }
    ],
    payloads: { "file-1": "one" }
  });

  try {
    const preview = await runExport(fixture, graph.client);
    graph.client.download = async () => {
      throw new Error("Graph GET failed (412 Precondition Failed)");
    };
    const result = await runExport(fixture, graph.client, {
      approvalToken: preview.approval.token
    });
    assert.equal(result.stage, "COMPLETE_WITH_ERRORS");
    assert.match(result.instruction, /source changed/i);

    const next = await runExport(fixture, graph.client);
    assert.equal(next.stage, "AWAITING_USER_APPROVAL");
  } finally {
    fixture.cleanup();
  }
});


test("SharePoint export fails closed on changed manifests and insufficient disk", async () => {
  const fixture = makeFixture(100);
  const firstGraph = makeGraph({
    rootPages: [
      { value: [{ id: "file-1", name: "one.txt", size: 3, eTag: "etag-1", file: {} }] }
    ],
    payloads: { "file-1": "one" }
  });
  const changedGraph = makeGraph({
    rootPages: [
      { value: [{ id: "file-1", name: "one.txt", size: 3, eTag: "etag-2", file: {} }] }
    ],
    payloads: { "file-1": "one" }
  });

  try {
    const preview = await runExport(fixture, firstGraph.client);
    for (const name of fs.readdirSync(fixture.cacheDir)) {
      if (name.startsWith("sharepoint-export-manifest-")) {
        fs.rmSync(path.join(fixture.cacheDir, name), { force: true });
      }
    }
    await assert.rejects(
      () =>
        runExport(fixture, changedGraph.client, {
          approvalToken: preview.approval.token
        }),
      /approved export plan no longer matches/
    );

    const blocked = await runExport(fixture, firstGraph.client, {
      availableDiskBytes: () => 2
    });
    assert.equal(blocked.stage, "BLOCKED_INSUFFICIENT_DISK");
    assert.equal(firstGraph.downloadCalls.length, 0);
  } finally {
    fixture.cleanup();
  }
});

test("SharePoint export allows the exact file cap and skips one byte above it", async () => {
  const fixture = makeFixture(3);
  const graph = makeGraph({
    rootPages: [
      {
        value: [
          { id: "exact", name: "exact.bin", size: 3, eTag: "etag-exact", file: {} },
          { id: "over", name: "over.bin", size: 4, eTag: "etag-over", file: {} }
        ]
      }
    ],
    payloads: { exact: "abc", over: "abcd" }
  });

  try {
    const preview = await runExport(fixture, graph.client);
    assert.equal(preview.stage, "AWAITING_USER_APPROVAL");
    assert.equal(preview.preview.pendingDownloadCount, 1);
    assert.equal(preview.preview.oversizedFileCount, 1);

    const result = await runExport(fixture, graph.client, {
      approvalToken: preview.approval.token
    });
    assert.equal(result.stage, "COMPLETE");
    assert.equal(result.downloadedCount, 1);
    assert.equal(result.skippedOversizedCount, 1);
    assert.equal(fs.readFileSync(path.join(fixture.destination, "exact.bin"), "utf8"), "abc");
    assert.equal(fs.existsSync(path.join(fixture.destination, "over.bin")), false);
  } finally {
    fixture.cleanup();
  }
});

test("SharePoint export rejects a destination path containing a link", async (t) => {
  const fixture = makeFixture(100);
  const graph = makeGraph({ rootPages: [{ value: [] }], payloads: {} });
  const realDestination = path.join(fixture.config.dataDir, "real-destination");
  const linkedDestination = path.join(fixture.config.dataDir, "linked-destination");
  fs.mkdirSync(realDestination);

  try {
    try {
      fs.symlinkSync(
        realDestination,
        linkedDestination,
        process.platform === "win32" ? "junction" : "dir"
      );
    } catch (error) {
      if (error?.code === "EPERM" || error?.code === "EACCES") {
        t.skip("Creating a test link is not permitted in this environment.");
        return;
      }
      throw error;
    }
    await assert.rejects(
      () =>
        exportSharePointFiles(
          fixture.config,
          "https://ohmylab.sharepoint.com/sites/Finance",
          path.join(linkedDestination, "export"),
          { accountId: "account-1", availableDiskBytes: () => 1000 },
          graph.client
        ),
      /symbolic link or junction/
    );
    assert.equal(graph.getCalls.length, 0);
  } finally {
    fixture.cleanup();
  }
});

test("SharePoint export excludes files without an eTag from the approved manifest", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({
    rootPages: [
      { value: [{ id: "file-1", name: "one.txt", size: 3, file: {} }] }
    ],
    payloads: { "file-1": "one" }
  });

  try {
    const result = await runExport(fixture, graph.client);
    assert.equal(result.stage, "COMPLETE");
    assert.equal(result.unsupportedItemCount, 1);
    assert.equal(graph.downloadCalls.length, 0);
    assert.equal(fs.existsSync(path.join(fixture.destination, "one.txt")), false);
  } finally {
    fixture.cleanup();
  }
});

test("SharePoint export verifies a complete partial file before finalizing it", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({
    rootPages: [
      { value: [{ id: "file-1", name: "one.txt", size: 3, eTag: "etag-1", file: {} }] }
    ],
    payloads: { "file-1": "one" }
  });

  try {
    const preview = await runExport(fixture, graph.client);
    const partialPath = path.join(
      fixture.destination,
      "one.txt.hare-part-" + preview.jobId
    );
    fs.writeFileSync(partialPath, "one", "utf8");
    const originalGet = graph.client.get.bind(graph.client);
    graph.client.get = async (url) => {
      if (url.includes("/items/file-1?$select=")) {
        return { id: "file-1", size: 3, eTag: "etag-2", file: {} };
      }
      return originalGet(url);
    };

    const result = await runExport(fixture, graph.client, {
      approvalToken: preview.approval.token
    });
    assert.equal(result.stage, "COMPLETE_WITH_ERRORS");
    assert.match(result.instruction, /source changed/i);
    assert.equal(fs.existsSync(path.join(fixture.destination, "one.txt")), false);
    assert.equal(fs.existsSync(partialPath), true);
  } finally {
    fixture.cleanup();
  }
});

test("SharePoint export never overwrites a file created during download finalization", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({
    rootPages: [
      { value: [{ id: "file-1", name: "race.txt", size: 3, eTag: "etag-1", file: {} }] }
    ],
    payloads: { "file-1": "new" }
  });

  try {
    const preview = await runExport(fixture, graph.client);
    const outputPath = path.join(fixture.destination, "race.txt");
    graph.client.download = async () => {
      fs.writeFileSync(outputPath, "user", "utf8");
      return new Response(Buffer.from("new"), {
        status: 200,
        headers: { "content-length": "3" }
      });
    };

    const result = await runExport(fixture, graph.client, {
      approvalToken: preview.approval.token
    });
    assert.equal(result.stage, "COMPLETE_WITH_ERRORS");
    assert.match(result.failures[0].error, /already exists/);
    assert.equal(fs.readFileSync(outputPath, "utf8"), "user");
  } finally {
    fixture.cleanup();
  }
});

test("delete-restricted finalization preserves the completed file", () => {
  const fixture = makeFixture(100);
  const temporaryPath = path.join(fixture.destination, "file.bin.part");
  const outputPath = path.join(fixture.destination, "file.bin");
  fs.writeFileSync(temporaryPath, "complete", "utf8");

  try {
    finalizeTemporaryDownload(temporaryPath, outputPath, {
      deleteRestricted: true,
      availableDiskBytes: () => 100
    });
    assert.equal(fs.readFileSync(outputPath, "utf8"), "complete");
    assert.equal(fs.existsSync(temporaryPath), true);
    assert.equal(fs.statSync(temporaryPath).size, 0);
  } finally {
    fixture.cleanup();
  }
});

test("copy finalization reserves one largest file of additional disk space", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({
    rootPages: [
      { value: [{ id: "file-1", name: "ten.bin", size: 10, eTag: "etag-1", file: {} }] }
    ],
    payloads: { "file-1": "0123456789" }
  });

  try {
    const result = await runExport(fixture, graph.client, {
      availableDiskBytes: () => 15,
      copyFinalizationRequired: () => true
    });
    assert.equal(result.stage, "BLOCKED_INSUFFICIENT_DISK");
    assert.equal(result.preview.pendingDownloadBytes, 10);
    assert.equal(result.preview.finalizationReserveBytes, 10);
    assert.equal(result.preview.requiredDiskBytes, 20);
    assert.equal(graph.downloadCalls.length, 0);
  } finally {
    fixture.cleanup();
  }
});

for (const sourceStatus of [404, 410]) {
  test("SharePoint source " + sourceStatus + " invalidates approval and checkpoint", async () => {
    const fixture = makeFixture(100);
    const graph = makeGraph({
      rootPages: [
        { value: [{ id: "file-1", name: "one.txt", size: 3, eTag: "etag-1", file: {} }] }
      ],
      payloads: { "file-1": "one" }
    });

    try {
      const preview = await runExport(fixture, graph.client);
      graph.client.download = async () => {
        throw new Error("Graph GET failed (" + sourceStatus + ")");
      };
      const result = await runExport(fixture, graph.client, {
        approvalToken: preview.approval.token
      });
      assert.equal(result.stage, "COMPLETE_WITH_ERRORS");
      assert.match(result.instruction, /source changed/i);

      const next = await runExport(fixture, graph.client);
      assert.equal(next.stage, "AWAITING_USER_APPROVAL");
    } finally {
      fixture.cleanup();
    }
  });
}

test("SharePoint export validates the company site URL and CLI contract", async () => {
  const fixture = makeFixture(100);
  const graph = makeGraph({ rootPages: [{ value: [] }], payloads: {} });
  try {
    await assert.rejects(
      () =>
        exportSharePointFiles(
          fixture.config,
          "https://example.com/sites/Finance",
          fixture.destination,
          { accountId: "account-1", availableDiskBytes: () => 1000 },
          graph.client
        ),
      /ohmylab\.sharepoint\.com/
    );
    await assert.rejects(
      () =>
        exportSharePointFiles(
          fixture.config,
          "https://ohmylab.sharepoint.com/sites/../Finance",
          fixture.destination,
          { accountId: "account-1", availableDiskBytes: () => 1000 },
          graph.client
        ),
      /\/sites\/\.\.\.|identify one/
    );

    const help = spawnSync(
      process.execPath,
      [path.resolve("dist/cli.js"), "sharepoint", "export-files", "--help"],
      {
        cwd: process.cwd(),
        encoding: "utf8",
        env: { ...process.env, HARE_M365_DATA_DIR: fixture.config.dataDir }
      }
    );
    assert.equal(help.status, 0, help.stderr);
    assert.match(help.stdout, /--site-url/);
    assert.match(help.stdout, /--destination/);
    assert.match(help.stdout, /--approval-token/);
    assert.match(help.stdout, /--time-budget-ms/);
  } finally {
    fixture.cleanup();
  }
});

function makeFixture(maxSharePointExportFileBytes) {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "hare-sharepoint-export-"));
  const cacheDir = path.join(dataDir, ".cache");
  const destination = path.join(dataDir, "external");
  fs.mkdirSync(cacheDir, { recursive: true });
  fs.mkdirSync(destination, { recursive: true });
  return {
    cacheDir,
    destination,
    config: {
      dataDir,
      cacheDir,
      tenantId: "tenant-1",
      policy: {
        allowDownloads: true,
        maxSharePointExportFileBytes
      }
    },
    cleanup: () => fs.rmSync(dataDir, { recursive: true, force: true })
  };
}

function runExport(fixture, client, options = {}) {
  return exportSharePointFiles(
    fixture.config,
    "https://ohmylab.sharepoint.com/sites/Finance",
    fixture.destination,
    {
      accountId: "account-1",
      availableDiskBytes: () => 1_000_000,
      ...options
    },
    client
  );
}

function makeGraph({ rootPages, folders = {}, payloads }) {
  const getCalls = [];
  const downloadCalls = [];
  let rootPageIndex = 0;
  const client = {
    async get(url) {
      getCalls.push(url);
      if (url.startsWith("/sites/ohmylab.sharepoint.com:")) {
        return {
          id: "site-1",
          displayName: "Finance",
          webUrl: "https://ohmylab.sharepoint.com/sites/Finance"
        };
      }
      if (url.startsWith("/sites/site-1/drive")) {
        return { id: "drive-1", name: "Documents" };
      }
      if (url.includes("root-page-2")) {
        return rootPages[1];
      }
      if (url.includes("/root/children")) {
        rootPageIndex += 1;
        return rootPages[0];
      }
      const folderMatch = url.match(/\/items\/([^/]+)\/children/);
      if (folderMatch) {
        return { value: folders[decodeURIComponent(folderMatch[1])] ?? [] };
      }
      throw new Error(`Unexpected Graph GET: ${url}`);
    },
    async download(url, options = {}) {
      const match = url.match(/\/items\/([^/]+)\/content/);
      const itemId = decodeURIComponent(match?.[1] ?? "");
      const payload = payloads[itemId];
      if (payload === undefined) throw new Error(`Missing payload for ${itemId}`);
      downloadCalls.push({ url, options });
      const offset = options.rangeStart ?? 0;
      const body = Buffer.from(payload).subarray(offset);
      return new Response(body, {
        status: offset > 0 ? 206 : 200,
        headers: {
          "content-length": String(body.length),
          ...(offset > 0
            ? { "content-range": `bytes ${offset}-${Buffer.byteLength(payload) - 1}/${Buffer.byteLength(payload)}` }
            : {})
        }
      });
    }
  };
  return { client, getCalls, downloadCalls };
}
