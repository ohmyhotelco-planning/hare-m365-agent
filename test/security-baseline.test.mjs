import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";

const read = (file) => JSON.parse(fs.readFileSync(new URL(file, import.meta.url), "utf8"));

test("package and lock declare the October 2026 patched Node and undici floors", () => {
  const pkg = read("../package.json");
  const lock = read("../package-lock.json");
  assert.equal(pkg.engines.node, "^22.23.4 || ^24.22.1 || ^26.12.1");
  assert.deepEqual(lock.packages[""].engines, pkg.engines);
  assert.equal(pkg.dependencies.undici, "^7.30.1");
  assert.equal(lock.packages[""].dependencies.undici, pkg.dependencies.undici);
  assert.equal(lock.packages["node_modules/undici"].version, "7.30.1");
  assert.equal(read("../node_modules/undici/package.json").version, "7.30.1");
});
