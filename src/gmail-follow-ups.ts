import {
  buildFollowUpCacheKey,
  parseCachedFollowUps,
  serializeCachedFollowUps,
} from "./cache-key.ts";
import {
  CACHE_TTL_SECONDS,
  FOLLOWUP_WAITING_LABEL,
  MAX_CACHE_VALUE_BYTES,
} from "./constants.ts";
import { buildSearchQuery } from "./dates.ts";
import { parseExcludedDomainList } from "./email-utils.ts";
import { filterThreadsToPending } from "./follow-ups-core.ts";
import type { ThreadMetadata } from "./gmail-api-core.ts";
import { diffLabelMembership } from "./labels-core.ts";
import { log, perf } from "./log.ts";
import {
  addLabelToMessages,
  fetchThreadMetadata,
  listThreadIdsWithLabel,
  removeLabelFromThreadIds,
  searchThreadIds,
} from "./runtime/gmail-api.ts";
import { ensureFollowUpLabels } from "./runtime/labels.ts";
import { getMyDomain, getMyEmail } from "./runtime/session-email.ts";
import type { FollowUpQueryResult, PendingEmail, UserSettings } from "./types.ts";

/**
 * Keeps `Follow-ups/Waiting` equal to the current follow-up list: labels new
 * pending threads and unlabels ones that dropped out. Unlabelling is skipped
 * when the search was truncated, since unscanned threads may still be pending.
 */
function syncWaitingLabel(
  scanned: ThreadMetadata[],
  emails: PendingEmail[],
  searchTruncated: boolean,
): { labelled: number; added: number; removed: number } {
  ensureFollowUpLabels();
  const labelled = listThreadIdsWithLabel(FOLLOWUP_WAITING_LABEL);
  if (labelled.truncated) log.warn("waiting label listing truncated");
  const changes = diffLabelMembership(
    new Set(emails.map((e) => e.threadId)),
    new Set(labelled.items),
    !searchTruncated,
  );
  const toAdd = new Set(changes.add);
  addLabelToMessages(
    FOLLOWUP_WAITING_LABEL,
    scanned.filter((t) => toAdd.has(t.thread.threadId)).flatMap((t) => t.messageIds),
  );
  removeLabelFromThreadIds(FOLLOWUP_WAITING_LABEL, changes.remove);
  return {
    labelled: labelled.items.length,
    added: changes.add.length,
    removed: changes.remove.length,
  };
}

function tryWriteCache(
  cacheKey: string,
  emails: PendingEmail[],
  searchTruncated: boolean,
  fetchedAt: Date,
): void {
  try {
    const payload = serializeCachedFollowUps(emails, searchTruncated, fetchedAt);
    if (payload.length > MAX_CACHE_VALUE_BYTES) return;
    CacheService.getUserCache().put(cacheKey, payload, CACHE_TTL_SECONDS);
  } catch (err) {
    log.warn("cache write skipped", { error: err });
  }
}

export function getPendingFollowUps(
  settings: UserSettings,
  ignoreCache = false,
): FollowUpQueryResult {
  const t0 = Date.now();
  const cacheKey = buildFollowUpCacheKey(settings);
  const cache = CacheService.getUserCache();

  if (!ignoreCache) {
    try {
      const hit = cache.get(cacheKey);
      if (hit) {
        const restored = parseCachedFollowUps(hit);
        if (restored) {
          perf("getPendingFollowUps", t0, { cacheHit: true });
          return {
            emails: restored.emails,
            threadsScanned: 0,
            searchTruncated: restored.searchTruncated,
            fetchedAt: restored.fetchedAt,
          };
        }
      }
    } catch (err) {
      log.warn("cache read skipped", { error: err });
    }
  }

  const fetchedAt = new Date();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - settings.days);

  const staleCutoff = new Date();
  staleCutoff.setDate(staleCutoff.getDate() - settings.staleDays);

  const query = buildSearchQuery(staleCutoff, cutoff);

  let threadIds: string[] = [];
  let searchTruncated = false;

  const t1 = Date.now();
  try {
    const result = searchThreadIds(query);
    threadIds = result.items;
    searchTruncated = result.truncated;
  } catch (err) {
    log.error("Gmail search failed", { error: err });
    const error = `Gmail search failed: ${err instanceof Error ? err.message : "unknown"}`;
    return { emails: [], threadsScanned: 0, searchTruncated: false, fetchedAt, error };
  }
  perf("Threads.list", t1, { threads: threadIds.length, searchTruncated });

  const myEmail = getMyEmail();
  const myDomain = getMyDomain();
  const domainList = parseExcludedDomainList(settings.excludedDomains);

  let scanned: ThreadMetadata[] = [];
  let fetchFailed = 0;
  const t2 = Date.now();
  try {
    const result = fetchThreadMetadata(threadIds);
    scanned = result.threads;
    fetchFailed = result.failed;
    if (fetchFailed > 0) {
      log.warn("some threads could not be fetched", {
        failed: fetchFailed,
        firstError: result.firstError,
      });
      // Nothing usable came back: report the error instead of an empty list
      if (scanned.length === 0) {
        return {
          emails: [],
          threadsScanned: threadIds.length,
          searchTruncated,
          fetchedAt,
          error: `Could not load threads (${result.firstError ?? "unknown error"})`,
        };
      }
      // Unfetched threads may still be pending; flag the list as incomplete
      // so the UI says so and the waiting label is not pruned.
      searchTruncated = true;
    }
  } catch (err) {
    log.error("thread metadata fetch failed", { error: err });
    const error = `Could not load threads: ${err instanceof Error ? err.message : "unknown"}`;
    return {
      emails: [],
      threadsScanned: threadIds.length,
      searchTruncated,
      fetchedAt,
      error,
    };
  }
  perf("threads.get metadata", t2, { threads: scanned.length, failed: fetchFailed });

  const emails = filterThreadsToPending(
    scanned.map((t) => t.thread),
    {
      cutoff,
      staleCutoff,
      excludeReplied: settings.excludeReplied,
      excludeInternal: settings.excludeInternal,
      excludedDomainList: domainList,
      myEmail,
      myDomain,
    },
  );

  const t3 = Date.now();
  try {
    perf("syncWaitingLabel", t3, syncWaitingLabel(scanned, emails, searchTruncated));
  } catch (err) {
    log.error("waiting label sync failed", { error: err });
  }

  // A partial fetch is worth showing but not caching; the next load retries it
  if (fetchFailed === 0) tryWriteCache(cacheKey, emails, searchTruncated, fetchedAt);

  perf("getPendingFollowUps", t0, {
    pending: emails.length,
    threads: threadIds.length,
    searchTruncated,
  });

  return {
    emails,
    threadsScanned: threadIds.length,
    searchTruncated,
    fetchedAt,
  };
}
