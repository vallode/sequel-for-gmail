import { FOLLOWUP_LABELS, LABEL_BATCH_SIZE } from "../constants.ts";
import { paginateSearch } from "../follow-ups-core.ts";
import { missingLabelNames } from "../labels-core.ts";
import { log } from "../log.ts";
import { chunk } from "../util.ts";

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
  if (truncated) log.warn("label threads truncated", { label: label.getName() });
  return items.filter((thread) => !thread.isInTrash());
}

export function addLabelToThreads(
  label: GoogleAppsScript.Gmail.GmailLabel,
  threads: GoogleAppsScript.Gmail.GmailThread[],
): void {
  for (const batch of chunk(threads, LABEL_BATCH_SIZE)) label.addToThreads(batch);
}

export function removeLabelFromThreads(
  label: GoogleAppsScript.Gmail.GmailLabel,
  threads: GoogleAppsScript.Gmail.GmailThread[],
): void {
  for (const batch of chunk(threads, LABEL_BATCH_SIZE)) label.removeFromThreads(batch);
}
