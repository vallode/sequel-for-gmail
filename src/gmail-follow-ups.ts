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
import {
  addLabelToMessages,
  fetchThreadMetadata,
  searchThreadIds,
} from "./runtime/gmail-api.ts";
import {
  ensureFollowUpLabels,
  listLabelThreads,
  removeLabelFromThreads,
} from "./runtime/labels.ts";
import { getMyDomain, getMyEmail } from "./runtime/session-email.ts";
import type { FollowUpQueryResult, PendingEmail, UserSettings } from "./types.ts";
import { perf } from "./util.ts";

/**
 * Keeps `Follow-ups/Waiting` equal to the current follow-up list: labels new
 * pending threads and unlabels ones that dropped out. Unlabelling is skipped
 * when the search was truncated, since unscanned threads may still be pending.
 */
function syncWaitingLabel(
  scanned: ThreadMetadata[],
  emails: PendingEmail[],
  searchTruncated: boolean,
): void {
  const { waiting } = ensureFollowUpLabels();
  const labelled = listLabelThreads(waiting);
  const changes = diffLabelMembership(
    new Set(emails.map((e) => e.threadId)),
    new Set(labelled.map((t) => t.getId())),
    !searchTruncated,
  );
  const toAdd = new Set(changes.add);
  const toRemove = new Set(changes.remove);
  addLabelToMessages(
    FOLLOWUP_WAITING_LABEL,
    scanned.filter((t) => toAdd.has(t.thread.threadId)).flatMap((t) => t.messageIds),
  );
  removeLabelFromThreads(
    waiting,
    labelled.filter((t) => toRemove.has(t.getId())),
  );
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
    console.log(
      `[sequel] cache write skipped: ${err instanceof Error ? err.message : "unknown"}`,
    );
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
          perf("getPendingFollowUps (cache hit)", t0);
          return {
            emails: restored.emails,
            threadsScanned: 0,
            searchTruncated: restored.searchTruncated,
            fetchedAt: restored.fetchedAt,
          };
        }
      }
    } catch (err) {
      console.log(
        `[sequel] cache read skipped: ${err instanceof Error ? err.message : "unknown"}`,
      );
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
    console.log(
      `[sequel] Gmail search failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
    return { emails: [], threadsScanned: 0, searchTruncated: false, fetchedAt };
  }
  perf(`Threads.list (${threadIds.length} threads)`, t1);

  const myEmail = getMyEmail();
  const myDomain = getMyDomain();
  const domainList = parseExcludedDomainList(settings.excludedDomains);

  let scanned: ThreadMetadata[] = [];
  const t2 = Date.now();
  try {
    const result = fetchThreadMetadata(threadIds);
    scanned = result.threads;
    // Threads that could not be fetched may still be pending; flag the list
    // as incomplete so the UI says so and the waiting label is not pruned.
    if (result.failed > 0) {
      console.log(`[sequel] ${result.failed} threads could not be fetched`);
      searchTruncated = true;
    }
  } catch (err) {
    console.log(
      `[sequel] thread fetch failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
    return {
      emails: [],
      threadsScanned: threadIds.length,
      searchTruncated,
      fetchedAt,
    };
  }
  perf(`threads.get metadata (${scanned.length} threads)`, t2);

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

  try {
    syncWaitingLabel(scanned, emails, searchTruncated);
  } catch (err) {
    console.log(
      `[sequel] waiting label sync failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
  }

  tryWriteCache(cacheKey, emails, searchTruncated, fetchedAt);

  perf(`getPendingFollowUps (${emails.length}/${threadIds.length} threads)`, t0);

  return {
    emails,
    threadsScanned: threadIds.length,
    searchTruncated,
    fetchedAt,
  };
}
