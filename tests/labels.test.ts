import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { FOLLOWUP_LABELS } from "../src/constants.ts";
import { diffLabelMembership, missingLabelNames } from "../src/labels-core.ts";

describe("follow-up labels", () => {
  it("lists the parent before its children", () => {
    const [parent, ...children] = FOLLOWUP_LABELS;
    for (const child of children) assert.ok(child.startsWith(`${parent}/`));
  });

  it("creates only missing labels, parent first, ignoring case", () => {
    assert.deepEqual(missingLabelNames(FOLLOWUP_LABELS, []), [...FOLLOWUP_LABELS]);
    assert.deepEqual(
      missingLabelNames(FOLLOWUP_LABELS, ["follow-ups", "FOLLOW-UPS/WAITING", "Other"]),
      ["Follow-ups/Auto", "Follow-ups/Done"],
    );
    assert.deepEqual(missingLabelNames(FOLLOWUP_LABELS, [...FOLLOWUP_LABELS]), []);
  });

  it("adds new pending threads and removes ones that dropped out", () => {
    assert.deepEqual(
      diffLabelMembership(new Set(["a", "b"]), new Set(["b", "c"]), true),
      { add: ["a"], remove: ["c"] },
    );
  });

  it("never removes when the pending list may be incomplete", () => {
    assert.deepEqual(diffLabelMembership(new Set(["a"]), new Set(["b"]), false), {
      add: ["a"],
      remove: [],
    });
  });
});
