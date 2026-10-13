import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  buildFollowUpCacheKey,
  parseCachedFollowUps,
  serializeCachedFollowUps,
} from "../src/cache-key.ts";
import { MAX_CACHE_KEY_LENGTH } from "../src/constants.ts";
import type { UserSettings } from "../src/types.ts";

const sampleSettings = (): UserSettings => ({
  days: 3,
  excludeReplied: true,
  autoLabel: "Follow up",
  excludedDomains: "a.example.com,b.example.com",
  staleDays: 90,
  excludeInternal: false,
});

describe("cache keys", () => {
  it("stays within Apps Script cache key limit", () => {
    const key = buildFollowUpCacheKey(sampleSettings());
    assert.ok(key.length <= MAX_CACHE_KEY_LENGTH);
  });

  it("changes when settings change", () => {
    const a = buildFollowUpCacheKey(sampleSettings());
    const b = buildFollowUpCacheKey({ ...sampleSettings(), days: 4 });
    assert.notEqual(a, b);
  });
});

describe("cached follow-up payloads", () => {
  const emails = [
    {
      threadId: "t1",
      subject: "Hello",
      to: "a@example.com",
      date: new Date("2026-10-01T12:00:00Z"),
    },
  ];

  it("round-trips emails and the truncation flag", () => {
    for (const searchTruncated of [true, false]) {
      const restored = parseCachedFollowUps(
        serializeCachedFollowUps(emails, searchTruncated),
      );
      assert.deepEqual(restored, { emails, searchTruncated });
    }
  });

  it("rejects malformed and legacy array payloads", () => {
    assert.equal(parseCachedFollowUps("not json"), null);
    assert.equal(parseCachedFollowUps("null"), null);
    assert.equal(parseCachedFollowUps("[]"), null);
  });
});
