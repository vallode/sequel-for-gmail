import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  applySortOrder,
  evaluateThreadForFollowUp,
  filterThreadsToPending,
  sortPendingEmailsAsc,
} from "../src/follow-ups-core.ts";
import type { FollowUpFilterOptions, ThreadInput } from "../src/types.ts";

const baseOptions = (
  overrides: Partial<FollowUpFilterOptions> = {},
): FollowUpFilterOptions => {
  const cutoff = new Date("2026-01-10T12:00:00Z");
  const staleCutoff = new Date("2025-10-01T12:00:00Z");
  return {
    cutoff,
    staleCutoff,
    excludeReplied: true,
    excludeInternal: false,
    excludedDomainList: [],
    myEmail: "me@corp.com",
    myDomain: "corp.com",
    ...overrides,
  };
};

describe("follow-up core logic", () => {
  it("treats last message not from user as no follow-up row", () => {
    const thread: ThreadInput = {
      threadId: "t1",
      messages: [
        {
          from: "me@corp.com",
          to: "client@other.com",
          date: new Date("2026-01-05T12:00:00Z"),
          subject: "Hello",
        },
        {
          from: "client@other.com",
          to: "me@corp.com",
          date: new Date("2026-01-06T12:00:00Z"),
          subject: "Re: Hello",
        },
      ],
    };
    assert.equal(evaluateThreadForFollowUp(thread, baseOptions()), null);
  });

  it("includes thread when last message is from user within window", () => {
    const thread: ThreadInput = {
      threadId: "t2",
      messages: [
        {
          from: "me@corp.com",
          to: "Client <client@other.com>",
          date: new Date("2026-01-05T12:00:00Z"),
          subject: "Ping",
        },
      ],
    };
    const row = evaluateThreadForFollowUp(thread, baseOptions());
    assert.ok(row);
    assert.equal(row?.to, "client@other.com");
  });

  it("excludes messages outside date window", () => {
    const thread: ThreadInput = {
      threadId: "t3",
      messages: [
        {
          from: "me@corp.com",
          to: "client@other.com",
          date: new Date("2026-01-11T12:00:00Z"),
          subject: "Too new",
        },
      ],
    };
    assert.equal(evaluateThreadForFollowUp(thread, baseOptions()), null);
  });

  it("sorts ascending and toggles display order", () => {
    const a = {
      threadId: "a",
      subject: "s",
      to: "x@y.com",
      date: new Date("2026-01-01T00:00:00Z"),
    };
    const b = {
      ...a,
      threadId: "b",
      date: new Date("2026-01-03T00:00:00Z"),
    };
    const sorted = sortPendingEmailsAsc([b, a]);
    assert.equal(sorted[0].threadId, "a");
    assert.equal(applySortOrder(sorted, false)[0].threadId, "b");
  });

  it("filters batch with internal and domain rules", () => {
    const threads: ThreadInput[] = [
      {
        threadId: "internal",
        messages: [
          {
            from: "me@corp.com",
            to: "peer@corp.com",
            date: new Date("2026-01-05T12:00:00Z"),
            subject: "Internal",
          },
        ],
      },
      {
        threadId: "external",
        messages: [
          {
            from: "me@corp.com",
            to: "client@other.com",
            date: new Date("2026-01-05T12:00:00Z"),
            subject: "External",
          },
        ],
      },
    ];
    const internalExcluded = filterThreadsToPending(
      threads,
      baseOptions({ excludeInternal: true }),
    );
    assert.equal(internalExcluded.length, 1);
    assert.equal(internalExcluded[0].threadId, "external");

    const domainExcluded = filterThreadsToPending(
      threads,
      baseOptions({ excludedDomainList: ["other.com"] }),
    );
    assert.equal(domainExcluded.length, 1);
    assert.equal(domainExcluded[0].threadId, "internal");
  });
});
