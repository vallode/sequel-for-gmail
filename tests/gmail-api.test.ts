import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { paginateTokenSearch, type TokenPage } from "../src/follow-ups-core.ts";
import {
  buildThreadMetadataUrl,
  findHeader,
  isRetryableStatus,
  parseThreadMetadata,
} from "../src/gmail-api-core.ts";

/** Fake token-paged search over `total` results; `cap` limits page size. */
function fakeTokenSearch(total: number, cap = Number.POSITIVE_INFINITY) {
  const calls: Array<[string | undefined, number]> = [];
  const fetchPage = (pageToken: string | undefined, max: number): TokenPage<number> => {
    calls.push([pageToken, max]);
    const start = pageToken ? Number(pageToken) : 0;
    const end = Math.min(start + Math.min(max, cap), total);
    const items = Array.from({ length: Math.max(0, end - start) }, (_, i) => start + i);
    return { items, nextPageToken: end < total ? String(end) : undefined };
  };
  return { fetchPage, calls };
}

describe("paginateTokenSearch", () => {
  it("fetches one page when results fit under the cap", () => {
    const { fetchPage, calls } = fakeTokenSearch(120);
    const result = paginateTokenSearch(fetchPage, 500, 300);
    assert.equal(result.items.length, 120);
    assert.equal(result.truncated, false);
    assert.deepEqual(calls, [[undefined, 301]]);
  });

  it("does not flag truncation when results exactly fill the cap", () => {
    const { fetchPage } = fakeTokenSearch(300);
    assert.equal(paginateTokenSearch(fetchPage, 500, 300).truncated, false);
  });

  it("flags truncation and trims when more results exist", () => {
    const { fetchPage } = fakeTokenSearch(1000);
    const result = paginateTokenSearch(fetchPage, 500, 300);
    assert.equal(result.items.length, 300);
    assert.equal(result.truncated, true);
  });

  it("follows page tokens when the API returns short pages", () => {
    const { fetchPage, calls } = fakeTokenSearch(250, 100);
    const result = paginateTokenSearch(fetchPage, 500, 300);
    assert.equal(result.items.length, 250);
    assert.equal(result.truncated, false);
    assert.deepEqual(calls, [
      [undefined, 301],
      ["100", 201],
      ["200", 101],
    ]);
  });

  it("handles an empty search", () => {
    const { fetchPage } = fakeTokenSearch(0);
    assert.deepEqual(paginateTokenSearch(fetchPage, 500, 300), {
      items: [],
      truncated: false,
    });
  });
});

describe("thread metadata", () => {
  it("requests only metadata headers and needed fields", () => {
    const url = buildThreadMetadataUrl("abc123");
    assert.ok(
      url.startsWith("https://gmail.googleapis.com/gmail/v1/users/me/threads/abc123?"),
    );
    assert.ok(url.includes("format=metadata"));
    for (const h of ["From", "To", "Subject"])
      assert.ok(url.includes(`metadataHeaders=${h}`));
    assert.ok(url.includes("fields="));
  });

  it("matches header names case-insensitively", () => {
    const headers = [{ name: "subject", value: "Hi" }];
    assert.equal(findHeader(headers, "Subject"), "Hi");
    assert.equal(findHeader(headers, "To"), "");
    assert.equal(findHeader(undefined, "To"), "");
  });

  it("maps a thread response to filter input and message IDs", () => {
    const parsed = parseThreadMetadata({
      id: "t1",
      messages: [
        {
          id: "m1",
          internalDate: "1767225600000",
          payload: {
            headers: [
              { name: "From", value: "Me <me@corp.com>" },
              { name: "To", value: "a@example.com" },
              { name: "Subject", value: "Proposal" },
            ],
          },
        },
        { id: "m2", internalDate: "1767312000000", payload: { headers: [] } },
      ],
    });
    assert.deepEqual(parsed, {
      thread: {
        threadId: "t1",
        messages: [
          {
            from: "Me <me@corp.com>",
            to: "a@example.com",
            date: new Date(1767225600000),
            subject: "Proposal",
          },
          { from: "", to: "", date: new Date(1767312000000), subject: "" },
        ],
      },
      messageIds: ["m1", "m2"],
    });
  });

  it("rejects responses without a thread ID", () => {
    assert.equal(parseThreadMetadata({}), null);
  });

  it("retries only rate limits and server errors", () => {
    assert.equal(isRetryableStatus(429), true);
    assert.equal(isRetryableStatus(503), true);
    assert.equal(isRetryableStatus(404), false);
    assert.equal(isRetryableStatus(403), false);
  });
});
