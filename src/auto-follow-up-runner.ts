import { buildReplyMime, decideAutoFollowUp } from "./auto-follow-up-core.ts";
import {
  AUTO_FOLLOWUP_TRIGGER_HOURS,
  FOLLOWUP_AUTO_LABEL,
  FOLLOWUP_DONE_LABEL,
  MAX_AUTO_FOLLOWUP_DRAFTS_PER_RUN,
} from "./constants.ts";
import { getEmailDomain, parseExcludedDomainList } from "./email-utils.ts";
import type { AutoFollowUpThreadMetadata } from "./gmail-api-core.ts";
import { log } from "./log.ts";
import {
  createThreadDraft,
  fetchAutoFollowUpThreads,
  listThreadIdsWithLabel,
  modifyThreadLabels,
} from "./runtime/gmail-api.ts";
import { ensureFollowUpLabels } from "./runtime/labels.ts";
import { getMyEmail } from "./runtime/session-email.ts";
import { readUserSettings } from "./runtime/user-props.ts";
import type { AutoFollowUpDecision, AutoFollowUpOptions } from "./types.ts";

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
  ensureFollowUpLabels();
  return new Set(listThreadIdsWithLabel(FOLLOWUP_AUTO_LABEL).items);
}

export function setThreadAutoFollowUp(threadId: string, enabled: boolean): void {
  ensureFollowUpLabels();
  modifyThreadLabels(
    threadId,
    enabled
      ? { add: [FOLLOWUP_AUTO_LABEL], remove: [FOLLOWUP_DONE_LABEL] }
      : { remove: [FOLLOWUP_AUTO_LABEL] },
  );
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

function createFollowUpDraft(
  { thread, replyTo }: AutoFollowUpThreadMetadata,
  decision: Extract<AutoFollowUpDecision, { action: "draft" }>,
): void {
  const mime = buildReplyMime(
    {
      to: decision.to,
      cc: decision.cc,
      subject: thread.messages[0]?.subject ?? "",
      inReplyTo: replyTo.messageId,
      references: replyTo.references,
      body: decision.body,
    },
    (text) => Utilities.base64Encode(text, Utilities.Charset.UTF_8),
  );
  createThreadDraft(thread.threadId, mime);
}

/**
 * Creates due follow-up drafts on flagged threads and retires flagged threads
 * that got a reply or finished their sequence. Never sends mail.
 */
export function runAutoFollowUps(now: Date = new Date()): AutoFollowUpRunSummary {
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
    ensureFollowUpLabels();
    const threadIds = listThreadIdsWithLabel(FOLLOWUP_AUTO_LABEL).items;
    summary.scanned = threadIds.length;
    if (threadIds.length === 0) return summary;

    const fetched = fetchAutoFollowUpThreads(threadIds);
    if (fetched.failed > 0) {
      // Unfetched threads are simply retried on the next hourly run
      summary.failed += fetched.failed;
      log.warn("auto follow-up threads could not be fetched", {
        failed: fetched.failed,
        firstError: fetched.firstError,
      });
    }
    const options: AutoFollowUpOptions = {
      now,
      steps: settings.autoFollowUpSteps,
      excludeInternal: settings.excludeInternal,
      excludedDomainList: parseExcludedDomainList(settings.excludedDomains),
      myEmail,
      myDomain: getEmailDomain(myEmail),
    };

    for (const metadata of fetched.threads) {
      const { threadId } = metadata.thread;
      const decision = decideAutoFollowUp(metadata.thread, options);

      try {
        if (decision.action === "draft") {
          if (summary.drafted >= MAX_AUTO_FOLLOWUP_DRAFTS_PER_RUN) {
            summary.waiting++;
            continue;
          }
          // A draft decision implies no drafts on the thread, so the last
          // message (whose headers are in replyTo) is the user's latest sent.
          createFollowUpDraft(metadata, decision);
          summary.drafted++;
          log.info("auto follow-up drafted", {
            threadId,
            step: decision.stepIndex + 1,
          });
        } else if (decision.action === "stop") {
          modifyThreadLabels(threadId, {
            add: [FOLLOWUP_DONE_LABEL],
            remove: [FOLLOWUP_AUTO_LABEL],
          });
          summary.stopped++;
          log.info("auto follow-up stopped", { threadId, reason: decision.reason });
        } else {
          summary.waiting++;
        }
      } catch (err) {
        summary.failed++;
        log.error("auto follow-up failed on thread", { threadId, error: err });
      }
    }
  } finally {
    lock.releaseLock();
  }
  return summary;
}
