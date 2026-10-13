import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  buildReplyMime,
  countTrailingSentByUser,
  decideAutoFollowUp,
  firstNameFromRecipientHeader,
  parseAutoFollowUpSteps,
  renderTemplate,
  replySubject,
  serializeAutoFollowUpSteps,
} from "../src/auto-follow-up-core.ts";
import type {
  AutoFollowUpMessageInput,
  AutoFollowUpOptions,
  AutoFollowUpThreadInput,
} from "../src/types.ts";

const msg = (
  overrides: Partial<AutoFollowUpMessageInput> = {},
): AutoFollowUpMessageInput => ({
  from: "me@corp.com",
  to: "Jane Doe <jane@client.com>",
  cc: "",
  date: new Date("2026-01-01T12:00:00Z"),
  subject: "Proposal",
  isDraft: false,
  ...overrides,
});

const thread = (...messages: AutoFollowUpMessageInput[]): AutoFollowUpThreadInput => ({
  threadId: "t1",
  messages,
});

const baseOptions = (
  overrides: Partial<AutoFollowUpOptions> = {},
): AutoFollowUpOptions => ({
  now: new Date("2026-01-05T12:00:00Z"),
  steps: [
    { waitDays: 3, template: "Hi {{firstName}}, any thoughts on {{subject}}?" },
    {
      waitDays: 4,
      template: "Hi {{firstName}}, bumping this ({{daysSinceSent}} days).",
    },
  ],
  excludeInternal: false,
  excludedDomainList: [],
  myEmail: "me@corp.com",
  myDomain: "corp.com",
  ...overrides,
});

describe("auto follow-up decision", () => {
  it("drafts the first step once the wait has elapsed", () => {
    const decision = decideAutoFollowUp(
      thread(msg({ cc: "boss@corp.com" })),
      baseOptions(),
    );
    assert.deepEqual(decision, {
      action: "draft",
      threadId: "t1",
      stepIndex: 0,
      to: "Jane Doe <jane@client.com>",
      cc: "boss@corp.com",
      body: "Hi Jane, any thoughts on Proposal?",
    });
  });

  it("waits until the step is due", () => {
    const decision = decideAutoFollowUp(
      thread(msg()),
      baseOptions({ now: new Date("2026-01-03T12:00:00Z") }),
    );
    assert.deepEqual(decision, { action: "wait", threadId: "t1", reason: "not-due" });
  });

  it("waits while any draft is open on the thread", () => {
    const decision = decideAutoFollowUp(
      thread(msg(), msg({ isDraft: true })),
      baseOptions(),
    );
    assert.equal(decision.action, "wait");
    assert.equal(decision.action === "wait" && decision.reason, "draft-pending");
  });

  it("stops when the recipient replied", () => {
    const decision = decideAutoFollowUp(
      thread(msg(), msg({ from: "Jane <jane@client.com>", to: "me@corp.com" })),
      baseOptions(),
    );
    assert.equal(decision.action === "stop" && decision.reason, "replied");
  });

  it("advances to the next step after a sent nudge, timed from that nudge", () => {
    const nudge = msg({ date: new Date("2026-01-04T12:00:00Z") });
    const early = decideAutoFollowUp(
      thread(msg(), nudge),
      baseOptions({ now: new Date("2026-01-07T12:00:00Z") }),
    );
    assert.equal(early.action, "wait");

    const due = decideAutoFollowUp(
      thread(msg(), nudge),
      baseOptions({ now: new Date("2026-01-08T12:00:00Z") }),
    );
    assert.equal(due.action, "draft");
    if (due.action === "draft") {
      assert.equal(due.stepIndex, 1);
      assert.equal(due.body, "Hi Jane, bumping this (7 days).");
    }
  });

  it("stops once every step has been sent", () => {
    const decision = decideAutoFollowUp(thread(msg(), msg(), msg()), baseOptions());
    assert.equal(decision.action === "stop" && decision.reason, "sequence-complete");
  });

  it("restarts the sequence after a reply followed by a new message", () => {
    const decision = decideAutoFollowUp(
      thread(msg(), msg({ from: "jane@client.com" }), msg()),
      baseOptions(),
    );
    assert.equal(decision.action === "draft" && decision.stepIndex, 0);
  });

  it("stops for excluded domains and internal recipients", () => {
    const excluded = decideAutoFollowUp(
      thread(msg()),
      baseOptions({ excludedDomainList: ["client.com"] }),
    );
    assert.equal(excluded.action === "stop" && excluded.reason, "excluded");

    const internal = decideAutoFollowUp(
      thread(msg({ to: "colleague@corp.com" })),
      baseOptions({ excludeInternal: true }),
    );
    assert.equal(internal.action === "stop" && internal.reason, "excluded");
  });

  it("stops when there is no parseable recipient", () => {
    const decision = decideAutoFollowUp(thread(msg({ to: "" })), baseOptions());
    assert.equal(decision.action === "stop" && decision.reason, "no-recipient");
  });

  it("falls back to a generic greeting without a display name", () => {
    const decision = decideAutoFollowUp(
      thread(msg({ to: "jane@client.com" })),
      baseOptions(),
    );
    assert.equal(
      decision.action === "draft" && decision.body,
      "Hi there, any thoughts on Proposal?",
    );
  });
});

describe("auto follow-up helpers", () => {
  it("counts trailing messages sent by the user", () => {
    const messages = [msg(), msg({ from: "jane@client.com" }), msg(), msg()];
    assert.equal(countTrailingSentByUser(messages, "me@corp.com"), 2);
    assert.equal(countTrailingSentByUser([], "me@corp.com"), 0);
  });

  it("renders known placeholders and leaves unknown ones", () => {
    assert.equal(
      renderTemplate("Hi {{ firstName }}, re {{subject}} {{unknown}}", {
        firstName: "Jane",
        subject: "$& Plan",
      }),
      "Hi Jane, re $& Plan {{unknown}}",
    );
  });

  it("extracts first names from display names", () => {
    assert.equal(firstNameFromRecipientHeader("Jane Doe <jane@x.com>"), "Jane");
    assert.equal(firstNameFromRecipientHeader('"Doe, Jane" <jane@x.com>'), "Jane");
    assert.equal(firstNameFromRecipientHeader("jane@x.com"), null);
    assert.equal(firstNameFromRecipientHeader("a@x.com, Bob <b@y.com>"), null);
    assert.equal(firstNameFromRecipientHeader("<jane@x.com>"), null);
  });

  it("parses, clamps, and caps stored steps", () => {
    const steps = parseAutoFollowUpSteps(
      JSON.stringify([
        { waitDays: 0, template: " First " },
        { waitDays: 3, template: "" },
        { waitDays: "x", template: "Bad wait" },
        { waitDays: 9999, template: "Second" },
        { waitDays: 5, template: "Third" },
        { waitDays: 6, template: "Fourth" },
      ]),
    );
    assert.deepEqual(steps, [
      { waitDays: 1, template: "First" },
      { waitDays: 365, template: "Second" },
      { waitDays: 5, template: "Third" },
    ]);
    assert.deepEqual(parseAutoFollowUpSteps(serializeAutoFollowUpSteps(steps)), steps);
  });

  it("returns no steps for missing or malformed JSON", () => {
    assert.deepEqual(parseAutoFollowUpSteps(null), []);
    assert.deepEqual(parseAutoFollowUpSteps("{not json"), []);
    assert.deepEqual(parseAutoFollowUpSteps('{"a":1}'), []);
  });
});

describe("reply MIME", () => {
  const b64 = (text: string) => Buffer.from(text, "utf8").toString("base64");
  const decodeBody = (mime: string) =>
    Buffer.from(mime.split("\r\n\r\n")[1].replace(/\r\n/g, ""), "base64").toString(
      "utf8",
    );

  it("prefixes Re: once", () => {
    assert.equal(replySubject("Proposal"), "Re: Proposal");
    assert.equal(replySubject("RE: Proposal"), "RE: Proposal");
    assert.equal(replySubject("  "), "Re: (no subject)");
  });

  it("builds threaded headers and a UTF-8 body", () => {
    const mime = buildReplyMime(
      {
        to: "Jane <jane@client.com>",
        cc: "boss@corp.com",
        subject: "Proposal",
        inReplyTo: "<m2@mail>",
        references: "<m1@mail>",
        body: "Hi Jane,\nCzęść — thanks",
      },
      b64,
    );
    const [head] = mime.split("\r\n\r\n");
    assert.deepEqual(head.split("\r\n"), [
      "To: Jane <jane@client.com>",
      "Cc: boss@corp.com",
      "Subject: Re: Proposal",
      "In-Reply-To: <m2@mail>",
      "References: <m1@mail> <m2@mail>",
      "MIME-Version: 1.0",
      'Content-Type: text/plain; charset="UTF-8"',
      "Content-Transfer-Encoding: base64",
    ]);
    assert.equal(decodeBody(mime), "Hi Jane,\r\nCzęść — thanks");
  });

  it("omits empty Cc and threading headers, encodes non-ASCII subjects", () => {
    const mime = buildReplyMime(
      {
        to: "jane@client.com",
        cc: "",
        subject: "Café",
        inReplyTo: "",
        references: "",
        body: "x",
      },
      b64,
    );
    const head = mime.split("\r\n\r\n")[0];
    assert.ok(!head.includes("Cc:"));
    assert.ok(!head.includes("In-Reply-To:"));
    assert.ok(!head.includes("References:"));
    assert.ok(head.includes(`Subject: =?UTF-8?B?${b64("Re: Café")}?=`));
  });

  it("strips line breaks from header values", () => {
    const mime = buildReplyMime(
      {
        to: "jane@client.com\r\nBcc: evil@x.com",
        cc: "",
        subject: "Hi\nBcc: evil@x.com",
        inReplyTo: "",
        references: "",
        body: "x",
      },
      b64,
    );
    const head = mime.split("\r\n\r\n")[0];
    assert.ok(!head.split("\r\n").some((line) => line.startsWith("Bcc:")));
  });

  it("wraps long base64 bodies at 76 characters", () => {
    const mime = buildReplyMime(
      {
        to: "a@b.com",
        cc: "",
        subject: "s",
        inReplyTo: "",
        references: "",
        body: "y".repeat(500),
      },
      b64,
    );
    const bodyLines = mime.split("\r\n\r\n")[1].split("\r\n");
    assert.ok(bodyLines.length > 1);
    assert.ok(bodyLines.every((line) => line.length <= 76));
  });
});
