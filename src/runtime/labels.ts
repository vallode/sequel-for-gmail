import { FOLLOWUP_LABELS } from "../constants.ts";
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
