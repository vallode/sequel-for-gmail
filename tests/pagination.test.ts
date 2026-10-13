import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { paginateSearch } from "../src/follow-ups-core.ts";

/** Fake Gmail search over `total` results that records each call. */
function fakeSearch(total: number) {
  const calls: Array<[number, number]> = [];
  const search = (start: number, max: number): number[] => {
    calls.push([start, max]);
    const end = Math.min(start + max, total);
    return Array.from({ length: Math.max(0, end - start) }, (_, i) => start + i);
  };
  return { search, calls };
}

describe("pagination limits", () => {
  it("stops early when fewer results than the cap exist", () => {
    const { search, calls } = fakeSearch(150);
    const result = paginateSearch(search, 100, 300);
    assert.equal(result.items.length, 150);
    assert.equal(result.truncated, false);
    assert.deepEqual(calls, [
      [0, 100],
      [100, 100],
    ]);
  });

  it("does not flag truncation when results exactly fill the cap", () => {
    const { search, calls } = fakeSearch(300);
    const result = paginateSearch(search, 100, 300);
    assert.equal(result.items.length, 300);
    assert.equal(result.truncated, false);
    assert.deepEqual(calls.at(-1), [300, 1]);
  });

  it("flags truncation when results exceed the cap", () => {
    const { search } = fakeSearch(301);
    const result = paginateSearch(search, 100, 300);
    assert.equal(result.items.length, 300);
    assert.equal(result.truncated, true);
  });

  it("shrinks the last page to respect a cap that is not a page multiple", () => {
    const { search, calls } = fakeSearch(1000);
    const result = paginateSearch(search, 100, 250);
    assert.equal(result.items.length, 250);
    assert.deepEqual(calls, [
      [0, 100],
      [100, 100],
      [200, 50],
      [250, 1],
    ]);
  });

  it("handles an empty search", () => {
    const { search } = fakeSearch(0);
    assert.deepEqual(paginateSearch(search, 100, 300), { items: [], truncated: false });
  });
});
