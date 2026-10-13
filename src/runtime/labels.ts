import { FOLLOWUP_LABELS, LABEL_BATCH_SIZE } from "../constants.ts";
import { paginateSearch } from "../follow-ups-core.ts";
import { missingLabelNames } from "../labels-core.ts";

export interface FollowUpLabels {
  waiting: GoogleAppsScript.Gmail.GmailLabel;
  auto: GoogleAppsScript.Gmail.GmailLabel;
  done: GoogleAppsScript.Gmail.GmailLabel;
}

let resolved: FollowUpLabels | null = null;

/** Creates any missing `Follow-ups` labels (parent first) and returns them. */
export function ensureFollowUpLabels(): FollowUpLabels {
  if (resolved) return resolved;

  const byName = new Map<string, GoogleAppsScript.Gmail.GmailLabel>();
  for (const label of GmailApp.getUserLabels()) {
    byName.set(label.getName().toLowerCase(), label);
  }
  const missing = missingLabelNames(FOLLOWUP_LABELS, [...byName.keys()]);
  for (const name of missing) {
    byName.set(name.toLowerCase(), GmailApp.createLabel(name));
  }

  const get = (name: string) =>
    byName.get(name.toLowerCase()) as GoogleAppsScript.Gmail.GmailLabel;
  const [, waiting, auto, done] = FOLLOWUP_LABELS;
  resolved = { waiting: get(waiting), auto: get(auto), done: get(done) };
  return resolved;
}

export function listLabelThreads(
  label: GoogleAppsScript.Gmail.GmailLabel,
): GoogleAppsScript.Gmail.GmailThread[] {
  const { items, truncated } = paginateSearch((start, max) =>
    label.getThreads(start, max),
  );
  if (truncated) console.log(`[sequel] ${label.getName()}: label threads truncated`);
  return items.filter((thread) => !thread.isInTrash());
}

function inBatches(
  threads: GoogleAppsScript.Gmail.GmailThread[],
  apply: (batch: GoogleAppsScript.Gmail.GmailThread[]) => void,
): void {
  for (let i = 0; i < threads.length; i += LABEL_BATCH_SIZE) {
    apply(threads.slice(i, i + LABEL_BATCH_SIZE));
  }
}

export function addLabelToThreads(
  label: GoogleAppsScript.Gmail.GmailLabel,
  threads: GoogleAppsScript.Gmail.GmailThread[],
): void {
  inBatches(threads, (batch) => label.addToThreads(batch));
}

export function removeLabelFromThreads(
  label: GoogleAppsScript.Gmail.GmailLabel,
  threads: GoogleAppsScript.Gmail.GmailThread[],
): void {
  inBatches(threads, (batch) => label.removeFromThreads(batch));
}
