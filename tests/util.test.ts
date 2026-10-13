import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { formatUpdatedLabel } from "../src/dates.ts";
import { chunk } from "../src/util.ts";

describe("chunk", () => {
  it("splits into batches with a shorter final batch", () => {
    assert.deepEqual(chunk([1, 2, 3, 4, 5], 2), [[1, 2], [3, 4], [5]]);
  });

  it("returns no batches for an empty list", () => {
    assert.deepEqual(chunk([], 100), []);
  });
});

describe("formatUpdatedLabel", () => {
  const fetchedAt = new Date("2026-10-12T09:00:00Z");
  const after = (ms: number) => new Date(fetchedAt.getTime() + ms);

  it("says just now for under a minute", () => {
    assert.equal(formatUpdatedLabel(fetchedAt, after(59_000)), "Updated just now");
  });

  it("uses singular and plural minutes", () => {
    assert.equal(formatUpdatedLabel(fetchedAt, after(60_000)), "Updated 1 min ago");
    assert.equal(
      formatUpdatedLabel(fetchedAt, after(29 * 60_000)),
      "Updated 29 min ago",
    );
  });
});
