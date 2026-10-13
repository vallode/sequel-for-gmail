import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { describe, it } from "node:test";
import { CAT_WALK_GIF_DATA_URI } from "../src/cat-walk-gif.ts";

describe("cat easter egg", () => {
  it("embeds the committed GIF, not a stale copy", () => {
    const prefix = "data:image/gif;base64,";
    assert.ok(CAT_WALK_GIF_DATA_URI.startsWith(prefix));
    const embedded = Buffer.from(CAT_WALK_GIF_DATA_URI.slice(prefix.length), "base64");
    assert.equal(embedded.toString("ascii", 0, 6), "GIF89a");
    assert.deepEqual(embedded, readFileSync("assets/cat-walk.gif"));
  });
});
