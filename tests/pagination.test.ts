import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { GMAIL_SEARCH_PAGE_SIZE, MAX_THREADS_TO_SCAN } from "../src/constants.ts";
import { planThreadPagination } from "../src/follow-ups-core.ts";

describe("pagination limits", () => {
  it("plans pages until max threads cap", () => {
    const plan = planThreadPagination(null);
    assert.equal(plan.pages.length, MAX_THREADS_TO_SCAN / GMAIL_SEARCH_PAGE_SIZE);
    assert.equal(plan.pages[0], 0);
    assert.equal(plan.pages[1], GMAIL_SEARCH_PAGE_SIZE);
    assert.equal(plan.truncatedByCap, true);
  });

  it("stops early when gmail reports fewer threads", () => {
    const plan = planThreadPagination(150);
    assert.equal(plan.pages.length, 2);
    assert.equal(plan.truncatedByCap, false);
  });
});
