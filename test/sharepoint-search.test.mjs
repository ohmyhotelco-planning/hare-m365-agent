import assert from "node:assert/strict";
import test from "node:test";
import { searchFiles } from "../dist/sharepoint.js";

test("file search uses Microsoft Search across SharePoint and OneDrive", async () => {
  const calls = [];
  const client = {
    async post(url, body) {
      calls.push({ url, body });
      return {
        value: [
          {
            hitsContainers: [
              {
                total: 52,
                moreResultsAvailable: true,
                hits: [
                  {
                    resource: {
                      id: "file-1",
                      name: "Agent Automation.pdf",
                      webUrl: "https://contoso.sharepoint.com/file-1",
                      file: { mimeType: "application/pdf" },
                      parentReference: { driveId: "drive-1", siteId: "site-1" }
                    }
                  },
                  { resource: { id: "folder-1", name: "Folder", folder: {} } }
                ]
              }
            ]
          }
        ]
      };
    }
  };

  const result = await searchFiles(config(), "Agent Automation", 25, {}, client);
  assert.equal(calls[0].url, "/search/query");
  assert.deepEqual(calls[0].body.requests[0].entityTypes, ["driveItem"]);
  assert.equal(calls[0].body.requests[0].from, 0);
  assert.equal(calls[0].body.requests[0].size, 25);
  assert.equal(result.search.scope, "sharePointAndOneDrive");
  assert.equal(result.search.totalMatchesReported, 52);
  assert.equal(result.search.returnedCount, 1);
  assert.equal(result.search.nextOffset, 2);
  assert.equal(result.search.continuationAvailable, true);
  assert.equal(result.files[0].parentReference.siteId, "site-1");
});

test("file search supports offset continuation", async () => {
  const client = {
    async post(_url, body) {
      assert.equal(body.requests[0].from, 25);
      return { value: [{ hitsContainers: [{ total: 26, moreResultsAvailable: false, hits: [] }] }] };
    }
  };
  const result = await searchFiles(config(), "Agent", 25, { offset: 25 }, client);
  assert.equal(result.search.offset, 25);
  assert.equal(result.search.continuationAvailable, false);
});

test("file search returns a resumable partial result on timeout", async () => {
  let now = 0;
  const client = {
    async post() {
      now = 6;
      const error = new Error("Graph request exceeded the 5ms time budget.");
      error.name = "GraphTimeoutError";
      throw error;
    }
  };
  const result = await searchFiles(
    config(),
    "Agent",
    25,
    { timeBudgetMs: 5, now: () => now },
    client
  );
  assert.equal(result.search.partialResult, true);
  assert.equal(result.search.partialReason, "time-budget-exceeded");
  assert.equal(result.search.nextOffset, 0);
  assert.equal(result.search.continuationAvailable, true);
});

test("file search retains downloadable NDJSON outside the document classification", async () => {
  const query = 'path:"https://contoso.sharepoint.com/sites/usage" AND filename:"events.ndjson"';
  const client = {
    async post(url, body) {
      assert.equal(url, "/search/query");
      assert.equal(body.requests[0].query.queryString, query);
      const documentOnly = /IsDocument\s*[:=]\s*(?:True|1)/i.test(body.requests[0].query.queryTemplate ?? "");
      const hits = documentOnly ? [] : [
        { resource: { id: "ndjson", name: "events.ndjson", size: 1712,
          file: { mimeType: "application/octet-stream" } } },
        { resource: { id: "pdf", name: "report.pdf", file: { mimeType: "application/pdf" } } },
        { resource: { id: "folder", name: "events", folder: {} } }
      ];
      return { value: [{ hitsContainers: [{ total: hits.length, moreResultsAvailable: false, hits }] }] };
    }
  };
  const result = await searchFiles(config(), query, 25, {}, client);
  assert.deepEqual(result.files.map(file => file.id), ["ndjson", "pdf"]);
  assert.equal(result.files[0].size, 1712);
  assert.equal(result.search.returnedCount, 2);
  assert.equal(result.search.partialResult, false);
});

test("file search excludes folders mislabeled with a file facet without excluding empty files", async () => {
  const client = {
    async post(url, body) {
      assert.equal(url, "/search/query");
      const excludesFolders = body.requests[0].query.queryTemplate === "({searchTerms}) AND NOT ContentTypeId:0x0120*";
      // Search can attach a file facet to a real folder and omit its folder facet.
      const folder = { resource: { id: "folder", name: "folder", size: 0,
        file: { mimeType: "application/octet-stream" } } };
      const hits = [
        { resource: { id: "ndjson", name: "events.ndjson", size: 1712,
          file: { mimeType: "application/octet-stream" } } },
        { resource: { id: "empty", name: "empty.txt", size: 0, file: { mimeType: "text/plain" } } },
        ...(excludesFolders ? [] : [folder])
      ];
      return { value: [{ hitsContainers: [{ total: hits.length, moreResultsAvailable: false, hits }] }] };
    }
  };
  const result = await searchFiles(config(), "fixture", 25, {}, client);
  assert.deepEqual(result.files.map(file => file.id), ["ndjson", "empty"]);
  assert.equal(result.files[1].size, 0);
});

function config() {
  return {
    policy: { maxFileSearchLimit: 25 }
  };
}
