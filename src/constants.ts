export const DEFAULT_DAYS = 3;
export const DEFAULT_EXCLUDED_DOMAINS = "";
export const DEFAULT_STALE_DAYS = 90;
export const DEFAULT_EXCLUDE_INTERNAL = false;

export const MIN_FOLLOWUP_DAYS = 1;
export const MAX_FOLLOWUP_DAYS = 365;

export const DISMISSED_CAP = 100;
export const DISMISSED_PROPERTY_KEY = "dismissed_ids";

export const CACHE_TTL_SECONDS = 300;
export const MAX_CACHE_VALUE_BYTES = 100_000;
export const MAX_CACHE_KEY_LENGTH = 250;

/** Gmail search page size */
export const GMAIL_SEARCH_PAGE_SIZE = 100;
/** Stop scanning after this many threads (Apps Script time limits) */
export const MAX_THREADS_TO_SCAN = 300;

export const HOMEPAGE_DISPLAY_LIMIT = 15;

export const SOURCE_REPO_URL = "https://github.com/vallode/sequel-for-gmail";

export const FOLLOWUP_PARENT_LABEL = "Follow-ups";
/** Threads currently in the follow-up list */
export const FOLLOWUP_WAITING_LABEL = `${FOLLOWUP_PARENT_LABEL}/Waiting`;
/** Threads the user flagged for auto follow-up drafts */
export const FOLLOWUP_AUTO_LABEL = `${FOLLOWUP_PARENT_LABEL}/Auto`;
/** Flagged threads that got a reply or finished their sequence */
export const FOLLOWUP_DONE_LABEL = `${FOLLOWUP_PARENT_LABEL}/Done`;
/** Parent first, so Gmail nests the children under it */
export const FOLLOWUP_LABELS = [
  FOLLOWUP_PARENT_LABEL,
  FOLLOWUP_WAITING_LABEL,
  FOLLOWUP_AUTO_LABEL,
  FOLLOWUP_DONE_LABEL,
] as const;
/** GmailLabel.addToThreads / removeFromThreads accept at most 100 threads */
export const LABEL_BATCH_SIZE = 100;
export const AUTO_FOLLOWUP_STEPS_PROPERTY_KEY = "auto_followup_steps";
/** Keeps the serialized sequence well under the 9KB UserProperties value limit */
export const MAX_AUTO_FOLLOWUP_STEPS = 3;
export const MAX_AUTO_FOLLOWUP_TEMPLATE_LENGTH = 2000;
export const AUTO_FOLLOWUP_NAME_FALLBACK = "there";
/** Bounds a single trigger run well inside the Apps Script execution limit */
export const MAX_AUTO_FOLLOWUP_DRAFTS_PER_RUN = 20;
export const AUTO_FOLLOWUP_TRIGGER_HOURS = 1;
export const DEFAULT_AUTO_FOLLOWUP_STEPS: ReadonlyArray<{
  waitDays: number;
  template: string;
}> = [
  {
    waitDays: 3,
    template:
      'Hi {{firstName}},\n\nJust following up on my note about "{{subject}}". Let me know if you have any questions.\n\nThanks',
  },
  {
    waitDays: 4,
    template:
      "Hi {{firstName}},\n\nBumping this in case it got buried. Happy to pick it up whenever suits.\n\nThanks",
  },
];
