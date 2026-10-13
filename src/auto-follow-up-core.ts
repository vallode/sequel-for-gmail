import {
  AUTO_FOLLOWUP_NAME_FALLBACK,
  MAX_AUTO_FOLLOWUP_STEPS,
  MAX_AUTO_FOLLOWUP_TEMPLATE_LENGTH,
  MAX_FOLLOWUP_DAYS,
  MIN_FOLLOWUP_DAYS,
} from "./constants.ts";
import {
  getEmailDomain,
  isAnyDomainExcluded,
  isInternalRecipient,
  parseFirstRecipient,
  senderIsUser,
} from "./email-utils.ts";
import type {
  AutoFollowUpDecision,
  AutoFollowUpOptions,
  AutoFollowUpStep,
  AutoFollowUpThreadInput,
} from "./types.ts";

const DAY_MS = 86400000;
const TEMPLATE_VAR = /\{\{\s*(\w+)\s*\}\}/g;

/**
 * Drops entries without a template or numeric wait, clamps waits, trims and caps
 * templates, and keeps at most MAX_AUTO_FOLLOWUP_STEPS. Used for stored JSON and
 * settings-form input alike.
 */
export function normalizeAutoFollowUpSteps(items: unknown[]): AutoFollowUpStep[] {
  const steps: AutoFollowUpStep[] = [];
  for (const item of items) {
    if (item == null || typeof item !== "object") continue;
    const record = item as Record<string, unknown>;
    const template =
      typeof record.template === "string"
        ? record.template.trim().slice(0, MAX_AUTO_FOLLOWUP_TEMPLATE_LENGTH)
        : "";
    if (!template) continue;
    const waitRaw = typeof record.waitDays === "number" ? record.waitDays : Number.NaN;
    if (!Number.isFinite(waitRaw)) continue;
    const waitDays = Math.max(
      MIN_FOLLOWUP_DAYS,
      Math.min(MAX_FOLLOWUP_DAYS, Math.floor(waitRaw)),
    );
    steps.push({ waitDays, template });
    if (steps.length === MAX_AUTO_FOLLOWUP_STEPS) break;
  }
  return steps;
}

export function parseAutoFollowUpSteps(
  raw: string | null | undefined,
): AutoFollowUpStep[] {
  if (!raw || raw.trim() === "") return [];
  try {
    const parsed = JSON.parse(raw) as unknown;
    return Array.isArray(parsed) ? normalizeAutoFollowUpSteps(parsed) : [];
  } catch {
    return [];
  }
}

export function serializeAutoFollowUpSteps(steps: AutoFollowUpStep[]): string {
  return JSON.stringify(steps.slice(0, MAX_AUTO_FOLLOWUP_STEPS));
}

/** Replaces `{{name}}` placeholders; unknown placeholders are left as written. */
export function renderTemplate(template: string, vars: Record<string, string>): string {
  return template.replace(TEMPLATE_VAR, (match, name: string) =>
    typeof vars[name] === "string" ? vars[name] : match,
  );
}

/**
 * Best-effort first name from the display name of the first recipient in a raw
 * To header. Handles `Jane Doe <jane@x>` and `"Doe, Jane" <jane@x>`.
 */
export function firstNameFromRecipientHeader(rawTo: string): string | null {
  const angle = rawTo.indexOf("<");
  if (angle === -1) return null;
  let name = rawTo.slice(0, angle).trim();
  if (!name || name.includes("@")) return null;

  const quoted = name.startsWith('"') && name.endsWith('"') && name.length > 1;
  if (quoted) {
    name = name.slice(1, -1).trim();
    const comma = name.indexOf(",");
    if (comma !== -1) name = name.slice(comma + 1).trim();
  } else if (name.includes(",")) {
    // Unquoted comma means an earlier recipient had no display name.
    return null;
  }

  const first = name.split(/\s+/)[0];
  return first || null;
}

/** Messages the user sent at the end of the thread: the original plus any nudges. */
export function countTrailingSentByUser(
  messages: AutoFollowUpThreadInput["messages"],
  myEmail: string,
): number {
  let count = 0;
  for (let i = messages.length - 1; i >= 0; i--) {
    if (!senderIsUser(messages[i].from, myEmail)) break;
    count++;
  }
  return count;
}

export function decideAutoFollowUp(
  thread: AutoFollowUpThreadInput,
  options: AutoFollowUpOptions,
): AutoFollowUpDecision {
  const { threadId } = thread;

  // An open draft means a nudge (ours or the user's) is waiting for review.
  if (thread.messages.some((m) => m.isDraft)) {
    return { action: "wait", threadId, reason: "draft-pending" };
  }

  const sent = thread.messages; // no drafts past this point
  if (sent.length === 0) return { action: "stop", threadId, reason: "no-recipient" };

  const last = sent[sent.length - 1];
  if (!senderIsUser(last.from, options.myEmail)) {
    return { action: "stop", threadId, reason: "replied" };
  }

  const trailing = countTrailingSentByUser(sent, options.myEmail);
  const stepIndex = trailing - 1;
  const step = options.steps[stepIndex];
  if (!step) return { action: "stop", threadId, reason: "sequence-complete" };

  const recipient = parseFirstRecipient(last.to);
  if (!recipient) return { action: "stop", threadId, reason: "no-recipient" };

  if (
    (options.excludeInternal &&
      isInternalRecipient(recipient, options.myEmail, options.myDomain)) ||
    (options.excludedDomainList.length > 0 &&
      isAnyDomainExcluded(getEmailDomain(recipient), options.excludedDomainList))
  ) {
    return { action: "stop", threadId, reason: "excluded" };
  }

  const lastMs = last.date.getTime();
  if (options.now.getTime() < lastMs + step.waitDays * DAY_MS) {
    return { action: "wait", threadId, reason: "not-due" };
  }

  const original = sent[sent.length - trailing];
  const body = renderTemplate(step.template, {
    firstName: firstNameFromRecipientHeader(last.to) ?? AUTO_FOLLOWUP_NAME_FALLBACK,
    subject: sent[0]?.subject?.trim() || "(no subject)",
    daysSinceSent: String(
      Math.floor((options.now.getTime() - original.date.getTime()) / DAY_MS),
    ),
  });

  return { action: "draft", threadId, stepIndex, to: last.to, cc: last.cc, body };
}

export function replySubject(subject: string): string {
  const trimmed = subject.trim() || "(no subject)";
  return /^re:/i.test(trimmed) ? trimmed : `Re: ${trimmed}`;
}

export interface ReplyMimeInput {
  to: string;
  cc: string;
  subject: string;
  /** Message-ID of the message being replied to, angle brackets included */
  inReplyTo: string;
  /** Existing References header of that message, if any */
  references: string;
  body: string;
}

const LINE_BREAKS = /[\r\n]+/g;
const NON_ASCII = /[^\x20-\x7e]/;

function headerValue(value: string): string {
  return value.replace(LINE_BREAKS, " ").trim();
}

/**
 * Builds an RFC 2822 plain-text reply. `encodeBase64` is injected so this runs
 * under Node tests and Apps Script (`Utilities.base64Encode`) alike; it must
 * encode the string as UTF-8.
 */
export function buildReplyMime(
  input: ReplyMimeInput,
  encodeBase64: (text: string) => string,
): string {
  const subject = headerValue(replySubject(input.subject));
  const inReplyTo = headerValue(input.inReplyTo);
  const references = headerValue([input.references, inReplyTo].join(" "));

  const headers = [
    `To: ${headerValue(input.to)}`,
    input.cc.trim() ? `Cc: ${headerValue(input.cc)}` : "",
    `Subject: ${NON_ASCII.test(subject) ? `=?UTF-8?B?${encodeBase64(subject)}?=` : subject}`,
    inReplyTo ? `In-Reply-To: ${inReplyTo}` : "",
    references ? `References: ${references}` : "",
    "MIME-Version: 1.0",
    'Content-Type: text/plain; charset="UTF-8"',
    "Content-Transfer-Encoding: base64",
  ].filter(Boolean);

  const encoded = encodeBase64(input.body.replace(/\r?\n/g, "\r\n"));
  const wrapped = encoded.match(/.{1,76}/g)?.join("\r\n") ?? "";
  return `${headers.join("\r\n")}\r\n\r\n${wrapped}`;
}
