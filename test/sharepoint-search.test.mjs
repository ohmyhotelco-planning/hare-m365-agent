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

function config() {
  return {
    policy: { maxFileSearchLimit: 25 }
  };
}
