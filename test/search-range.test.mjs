import assert from "node:assert/strict";
import test from "node:test";
import { resolveSearchRange } from "../dist/search-range.js";

test("default range uses the Asia/Seoul calendar date", () => {
  const range = resolveSearchRange(undefined, undefined, 90, new Date("2026-07-12T15:30:00Z"), "Asia/Seoul");
  assert.equal(range.until, "2026-07-13");
  assert.equal(range.since, "2026-04-15");
  assert.equal(range.startDateTime, "2026-04-14T15:00:00.000Z");
  assert.equal(range.endDateTimeExclusive, "2026-07-13T15:00:00.000Z");
  assert.match(range.notice, /Asia\/Seoul/);
});

test("explicit range is inclusive and validates calendar dates", () => {
  const range = resolveSearchRange("2026-07-01", "2026-07-13", 90, new Date(), "Asia/Seoul");
  assert.equal(range.days, 13);
  assert.equal(range.usedDefaultLookback, false);
  assert.throws(
    () => resolveSearchRange("2026-02-30", "2026-03-01", 90, new Date(), "Asia/Seoul"),
    /valid calendar date/
  );
});
