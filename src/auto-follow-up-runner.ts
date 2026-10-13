import { buildReplyMime, decideAutoFollowUp } from "./auto-follow-up-core.ts";
import {
  AUTO_FOLLOWUP_TRIGGER_HOURS,
  MAX_AUTO_FOLLOWUP_DRAFTS_PER_RUN,
} from "./constants.ts";
import { getEmailDomain, parseExcludedDomainList } from "./email-utils.ts";
import { log, perf } from "./log.ts";
import { ensureFollowUpLabels, listLabelThreads } from "./runtime/labels.ts";
import { getMyEmail } from "./runtime/session-email.ts";
import { readUserSettings } from "./runtime/user-props.ts";
import type {
  AutoFollowUpDecision,
  AutoFollowUpMessageInput,
  AutoFollowUpOptions,
} from "./types.ts";

export interface AutoFollowUpRunSummary {
  scanned: number;
  drafted: number;
  stopped: number;
  waiting: number;
  failed: number;
  skipped?: "disabled" | "no-steps" | "no-user-email" | "locked";
}

/** Thread IDs currently flagged for auto follow-up. */
export function getAutoFollowUpThreadIds(): Set<string> {
  const { auto } = ensureFollowUpLabels();
  return new Set(listLabelThreads(auto).map((thread) => thread.getId()));
}

export function setThreadAutoFollowUp(threadId: string, enabled: boolean): void {
  const thread = GmailApp.getThreadById(threadId);
  if (!thread) throw new Error("Thread not found");
  const { auto, done } = ensureFollowUpLabels();
  if (enabled) {
    thread.addLabel(auto);
    thread.removeLabel(done);
  } else {
    thread.removeLabel(auto);
  }
}

/**
 * Installs one hourly `_runAutoFollowUps` trigger when enabled (removing any
 * duplicates), or removes it when disabled.
 */
export function syncAutoFollowUpTrigger(enabled: boolean): void {
  const existing = ScriptApp.getProjectTriggers().filter(
    (trigger) => trigger.getHandlerFunction() === "_runAutoFollowUps",
  );
  const keep = enabled ? existing.slice(0, 1) : [];
  for (const trigger of existing) {
    if (!keep.includes(trigger)) ScriptApp.deleteTrigger(trigger);
  }
  if (enabled && keep.length === 0) {
    ScriptApp.newTrigger("_runAutoFollowUps")
      .timeBased()
      .everyHours(AUTO_FOLLOWUP_TRIGGER_HOURS)
      .create();
  }
}

function toMessageInput(
  message: GoogleAppsScript.Gmail.GmailMessage,
): AutoFollowUpMessageInput {
  return {
    from: message.getFrom(),
    to: message.getTo(),
    cc: message.getCc(),
    date: new Date(message.getDate().getTime()),
    subject: message.getSubject(),
    isDraft: message.isDraft(),
  };
}

function createFollowUpDraft(
  threadId: string,
  subject: string,
  replyTo: GoogleAppsScript.Gmail.GmailMessage,
  decision: Extract<AutoFollowUpDecision, { action: "draft" }>,
): void {
  if (typeof Gmail === "undefined" || !Gmail?.Users?.Drafts) {
    throw new Error("Advanced Gmail service is not enabled");
  }
  const mime = buildReplyMime(
    {
      to: decision.to,
      cc: decision.cc,
      subject,
      inReplyTo: replyTo.getHeader("Message-ID") || replyTo.getHeader("Message-Id"),
      references: replyTo.getHeader("References"),
      body: decision.body,
    },
    (text) => Utilities.base64Encode(text, Utilities.Charset.UTF_8),
  );
  Gmail.Users.Drafts.create(
    { message: { raw: Utilities.base64EncodeWebSafe(mime), threadId } },
    "me",
  );
}

/**
 * Creates due follow-up drafts on flagged threads and retires flagged threads
 * that got a reply or finished their sequence. Never sends mail.
 */
export function runAutoFollowUps(now: Date = new Date()): AutoFollowUpRunSummary {
  const t0 = Date.now();
  const summary: AutoFollowUpRunSummary = {
    scanned: 0,
    drafted: 0,
    stopped: 0,
    waiting: 0,
    failed: 0,
  };

  const settings = readUserSettings();
  if (!settings.autoFollowUpEnabled) {
    syncAutoFollowUpTrigger(false);
    return { ...summary, skipped: "disabled" };
  }
  if (settings.autoFollowUpSteps.length === 0) {
    return { ...summary, skipped: "no-steps" };
  }

  // Without the user's address every message looks self-sent, so bail out.
  const myEmail = getMyEmail();
  if (!myEmail) return { ...summary, skipped: "no-user-email" };

  const lock = LockService.getUserLock();
  if (!lock.tryLock(5_000)) return { ...summary, skipped: "locked" };

  try {
    const { auto, done } = ensureFollowUpLabels();
    const threads = listLabelThreads(auto);
    summary.scanned = threads.length;
    if (threads.length === 0) return summary;

    const allMessages = GmailApp.getMessagesForThreads(threads);
    const options: AutoFollowUpOptions = {
      now,
      steps: settings.autoFollowUpSteps,
      excludeInternal: settings.excludeInternal,
      excludedDomainList: parseExcludedDomainList(settings.excludedDomains),
      myEmail,
      myDomain: getEmailDomain(myEmail),
    };

    threads.forEach((thread, i) => {
      const threadId = thread.getId();
      const messages = allMessages[i].filter((message) => !message.isInTrash());
      const decision = decideAutoFollowUp(
        { threadId, messages: messages.map(toMessageInput) },
        options,
      );

      try {
        if (decision.action === "draft") {
          if (summary.drafted >= MAX_AUTO_FOLLOWUP_DRAFTS_PER_RUN) {
            summary.waiting++;
            return;
          }
          // A draft decision implies no drafts on the thread, so the last
          // message is the user's latest sent message.
          createFollowUpDraft(
            threadId,
            messages[0].getSubject(),
            messages[messages.length - 1],
            decision,
          );
          summary.drafted++;
          log.info("auto follow-up drafted", {
            threadId,
            step: decision.stepIndex + 1,
          });
        } else if (decision.action === "stop") {
          thread.removeLabel(auto);
          thread.addLabel(done);
          summary.stopped++;
          log.info("auto follow-up stopped", { threadId, reason: decision.reason });
        } else {
          summary.waiting++;
        }
      } catch (err) {
        summary.failed++;
        log.error("auto follow-up failed on thread", err, { threadId });
      }
    });
  } finally {
    lock.releaseLock();
  }

  perf("runAutoFollowUps", t0, { ...summary });
  return summary;
}
