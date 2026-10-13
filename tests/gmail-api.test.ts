import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { paginateTokenSearch, type TokenPage } from "../src/follow-ups-core.ts";
import {
  AUTO_FOLLOWUP_HEADERS,
  boundaryFromContentType,
  buildBatchRequestBody,
  buildThreadMetadataPath,
  describeHttpError,
  findHeader,
  isRetryableStatus,
  parseAutoFollowUpThread,
  parseBatchResponse,
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
    const url = buildThreadMetadataPath("abc123");
    assert.ok(url.startsWith("/gmail/v1/users/me/threads/abc123?"));
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

  it("describes Google API errors from the response body", () => {
    const body = JSON.stringify({
      error: { code: 403, message: "Gmail API has not been used" },
    });
    assert.equal(describeHttpError(403, body), "403: Gmail API has not been used");
    assert.equal(describeHttpError(502, "<html>Bad gateway</html>"), "HTTP 502");
    assert.ok(
      describeHttpError(400, JSON.stringify({ error: { message: "x".repeat(500) } }))
        .length <= 200,
    );
  });
});

describe("batch requests", () => {
  it("builds one tagged GET part per path", () => {
    const body = buildBatchRequestBody(["/a", "/b"], "XYZ");
    assert.equal(
      body,
      "--XYZ\r\nContent-Type: application/http\r\nContent-ID: <item-0>\r\n\r\nGET /a\r\n" +
        "\r\n--XYZ\r\nContent-Type: application/http\r\nContent-ID: <item-1>\r\n\r\nGET /b\r\n" +
        "\r\n--XYZ--",
    );
  });

  it("reads the boundary from a Content-Type header", () => {
    assert.equal(
      boundaryFromContentType("multipart/mixed; boundary=batch_abc"),
      "batch_abc",
    );
    assert.equal(boundaryFromContentType('multipart/mixed; boundary="q"'), "q");
    assert.equal(boundaryFromContentType("application/json"), null);
  });

  it("parses statuses and bodies, matching parts by Content-ID", () => {
    const response = [
      "--batch_abc",
      "Content-Type: application/http",
      "Content-ID: <response-item-1>",
      "",
      "HTTP/1.1 404 Not Found",
      "Content-Type: application/json; charset=UTF-8",
      "",
      '{"error":{"code":404,"message":"Not Found"}}',
      "--batch_abc",
      "Content-Type: application/http",
      "Content-ID: <response-item-0>",
      "",
      "HTTP/1.1 200 OK",
      "Content-Type: application/json; charset=UTF-8",
      "Vary: Origin",
      "",
      '{"id":"t1","messages":[]}',
      "--batch_abc--",
      "",
    ].join("\r\n");
    assert.deepEqual(parseBatchResponse(response, "batch_abc"), [
      { index: 1, status: 404, body: '{"error":{"code":404,"message":"Not Found"}}' },
      { index: 0, status: 200, body: '{"id":"t1","messages":[]}' },
    ]);
  });

  it("skips parts without a Content-ID or status line", () => {
    const response = "--b\r\nContent-Type: text/plain\r\n\r\nhello\r\n--b--";
    assert.deepEqual(parseBatchResponse(response, "b"), []);
  });
});

describe("auto follow-up thread metadata", () => {
  const header = (name: string, value: string) => ({ name, value });

  it("requests the reply headers along with the filter headers", () => {
    const path = buildThreadMetadataPath("t1", AUTO_FOLLOWUP_HEADERS);
    for (const h of ["From", "To", "Cc", "Subject", "Message-ID", "References"]) {
      assert.ok(path.includes(`metadataHeaders=${h}`), h);
    }
    assert.ok(decodeURIComponent(path).includes("labelIds"));
  });

  it("drops trashed messages, marks drafts and keeps the last message's reply headers", () => {
    const parsed = parseAutoFollowUpThread({
      id: "t1",
      messages: [
        {
          id: "m1",
          internalDate: "1767225600000",
          labelIds: ["SENT"],
          payload: {
            headers: [
              header("From", "me@corp.com"),
              header("To", "Ann <ann@example.com>"),
              header("Cc", "bob@example.com"),
              header("Subject", "Proposal"),
              header("Message-Id", "<m1@corp.com>"),
            ],
          },
        },
        {
          id: "m2",
          internalDate: "1767312000000",
          labelIds: ["TRASH"],
          payload: { headers: [header("From", "ann@example.com")] },
        },
        {
          id: "m3",
          internalDate: "1767398400000",
          labelIds: ["SENT"],
          payload: {
            headers: [
              header("From", "me@corp.com"),
              header("To", "ann@example.com"),
              header("Subject", "Re: Proposal"),
              header("message-id", "<m3@corp.com>"),
              header("References", "<m1@corp.com>"),
            ],
          },
        },
        {
          id: "m4",
          internalDate: "1767484800000",
          labelIds: ["DRAFT"],
          payload: { headers: [header("From", "me@corp.com")] },
        },
      ],
    });
    assert.ok(parsed);
    assert.deepEqual(
      parsed.thread.messages.map((m) => [m.from, m.isDraft]),
      [
        ["me@corp.com", false],
        ["me@corp.com", false],
        ["me@corp.com", true],
      ],
    );
    assert.equal(parsed.thread.messages[0].cc, "bob@example.com");
    assert.equal(parsed.thread.messages[0].subject, "Proposal");
    // The draft is last, so its (absent) headers are what replyTo carries
    assert.deepEqual(parsed.replyTo, { messageId: "", references: "" });
  });

  it("takes reply headers from the last non-trashed message", () => {
    const parsed = parseAutoFollowUpThread({
      id: "t2",
      messages: [
        {
          id: "m1",
          internalDate: "1",
          payload: {
            headers: [header("Message-ID", "<a@x>"), header("References", "<r@x>")],
          },
        },
        { id: "m2", internalDate: "2", labelIds: ["TRASH"], payload: { headers: [] } },
      ],
    });
    assert.deepEqual(parsed?.replyTo, { messageId: "<a@x>", references: "<r@x>" });
  });

  it("rejects responses without a thread ID", () => {
    assert.equal(parseAutoFollowUpThread({ messages: [] }), null);
  });
});
