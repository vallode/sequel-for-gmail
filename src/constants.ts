export const DEFAULT_DAYS = 3;
export const DEFAULT_AUTO_LABEL = "";
export const DEFAULT_EXCLUDED_DOMAINS = "";
export const DEFAULT_STALE_DAYS = 90;
export const DEFAULT_EXCLUDE_INTERNAL = false;

export const MIN_FOLLOWUP_DAYS = 1;
export const MAX_FOLLOWUP_DAYS = 365;

export const DISMISSED_CAP = 100;
export const DISMISSED_PROPERTY_KEY = "dismissed_ids";

/** Results change slowly; the homepage Refresh action bypasses the cache */
export const CACHE_TTL_SECONDS = 1800;
export const MAX_CACHE_VALUE_BYTES = 100_000;
export const MAX_CACHE_KEY_LENGTH = 250;

/** Gmail search page size */
export const GMAIL_SEARCH_PAGE_SIZE = 100;
/** Stop scanning after this many threads (Apps Script time limits) */
export const MAX_THREADS_TO_SCAN = 300;

export const HOMEPAGE_DISPLAY_LIMIT = 15;

/** GmailLabel.addToThreads accepts at most 100 threads per call */
export const LABEL_BATCH_SIZE = 100;

export const SOURCE_REPO_URL = "https://github.com/vallode/sequel-for-gmail";
