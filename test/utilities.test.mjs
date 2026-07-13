import assert from "node:assert/strict";
import test from "node:test";
import { retryDelayMs } from "../dist/graph.js";
import { sanitizeFilename } from "../dist/sharepoint.js";

test("retry delay honors Retry-After seconds and caps exponential delay", () => {
  assert.equal(retryDelayMs("3", 1), 3000);
  assert.equal(retryDelayMs(undefined, 1), 500);
  assert.equal(retryDelayMs(undefined, 10), 8000);
});

test("download filenames are cross-platform safe", () => {
  assert.equal(sanitizeFilename("report?.pdf"), "report_.pdf");
  assert.equal(sanitizeFilename("CON"), "_CON");
  assert.equal(sanitizeFilename("..."), "download.bin");
});
