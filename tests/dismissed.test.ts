import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { DISMISSED_CAP } from "../src/constants.ts";
import {
  addDismissedEntry,
  capDismissedList,
  parseDismissedJson,
} from "../src/dismissed-store.ts";

describe("dismissed list", () => {
  it("returns empty list for malformed JSON", () => {
    assert.deepEqual(parseDismissedJson("{"), []);
    assert.deepEqual(parseDismissedJson("null"), []);
  });

  it("skips invalid entries", () => {
    const raw = JSON.stringify([
      { threadId: "ok", subject: "s", to: "a@b.com", dateMs: 1 },
      { threadId: "", subject: "bad" },
      "nope",
    ]);
    assert.equal(parseDismissedJson(raw).length, 1);
  });

  it("caps list and drops oldest", () => {
    const items = Array.from({ length: DISMISSED_CAP + 5 }, (_, i) => ({
      threadId: `t${i}`,
      subject: "s",
      to: "a@b.com",
      dateMs: i,
    }));
    const capped = capDismissedList(items);
    assert.equal(capped.length, DISMISSED_CAP);
    assert.equal(capped[0].threadId, "t5");
    assert.equal(capped[capped.length - 1].threadId, `t${DISMISSED_CAP + 4}`);
  });

  it("does not duplicate thread ids", () => {
    const base = [{ threadId: "t1", subject: "s", to: "a@b.com", dateMs: 1 }];
    const next = addDismissedEntry(base, {
      threadId: "t1",
      subject: "s2",
      to: "c@d.com",
      dateMs: 2,
    });
    assert.equal(next.length, 1);
  });
});
