import {
  buildFollowUpCacheKey,
  parseCachedFollowUps,
  serializeCachedFollowUps,
} from "./cache-key.ts";
import { CACHE_TTL_SECONDS, MAX_CACHE_VALUE_BYTES } from "./constants.ts";
import { buildSearchQuery } from "./dates.ts";
import { parseExcludedDomainList } from "./email-utils.ts";
import { filterThreadsToPending, paginateSearch } from "./follow-ups-core.ts";
import { diffLabelMembership } from "./labels-core.ts";
import {
  addLabelToThreads,
  ensureFollowUpLabels,
  listLabelThreads,
  removeLabelFromThreads,
} from "./runtime/labels.ts";
import { getMyDomain, getMyEmail } from "./runtime/session-email.ts";
import type {
  FollowUpQueryResult,
  PendingEmail,
  ThreadInput,
  UserSettings,
} from "./types.ts";
import { perf } from "./util.ts";

function mapGmailThreads(
  threads: GoogleAppsScript.Gmail.GmailThread[],
  allMessages: GoogleAppsScript.Gmail.GmailMessage[][],
): ThreadInput[] {
  return threads.map((thread, i) => ({
    threadId: thread.getId(),
    messages: allMessages[i].map((msg) => ({
      from: msg.getFrom(),
      to: msg.getTo(),
      date: new Date(msg.getDate().getTime()),
      subject: msg.getSubject(),
    })),
  }));
}

/**
 * Keeps `Follow-ups/Waiting` equal to the current follow-up list: labels new
 * pending threads and unlabels ones that dropped out. Unlabelling is skipped
 * when the search was truncated, since unscanned threads may still be pending.
 */
function syncWaitingLabel(
  scanned: GoogleAppsScript.Gmail.GmailThread[],
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
  addLabelToThreads(
    waiting,
    scanned.filter((t) => toAdd.has(t.getId())),
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
): void {
  try {
    const payload = serializeCachedFollowUps(emails, searchTruncated);
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
          };
        }
      }
    } catch (err) {
      console.log(
        `[sequel] cache read skipped: ${err instanceof Error ? err.message : "unknown"}`,
      );
    }
  }

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - settings.days);

  const staleCutoff = new Date();
  staleCutoff.setDate(staleCutoff.getDate() - settings.staleDays);

  const query = buildSearchQuery(staleCutoff, cutoff);

  let collected: GoogleAppsScript.Gmail.GmailThread[] = [];
  let searchTruncated = false;

  const t1 = Date.now();
  try {
    const result = paginateSearch((start, max) => GmailApp.search(query, start, max));
    collected = result.items;
    searchTruncated = result.truncated;
  } catch (err) {
    console.log(
      `[sequel] Gmail search failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
    return { emails: [], threadsScanned: 0, searchTruncated: false };
  }
  perf(`GmailApp.search (${collected.length} threads)`, t1);

  const myEmail = getMyEmail();
  const myDomain = getMyDomain();
  const domainList = parseExcludedDomainList(settings.excludedDomains);

  let threadInputs: ThreadInput[] = [];
  const t2 = Date.now();
  try {
    const allMessages = GmailApp.getMessagesForThreads(collected);
    threadInputs = mapGmailThreads(collected, allMessages);
  } catch (err) {
    console.log(
      `[sequel] getMessagesForThreads failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
    return {
      emails: [],
      threadsScanned: collected.length,
      searchTruncated,
    };
  }
  perf(`GmailApp.getMessagesForThreads (${collected.length} threads)`, t2);

  const emails = filterThreadsToPending(threadInputs, {
    cutoff,
    staleCutoff,
    excludeReplied: settings.excludeReplied,
    excludeInternal: settings.excludeInternal,
    excludedDomainList: domainList,
    myEmail,
    myDomain,
  });

  try {
    syncWaitingLabel(collected, emails, searchTruncated);
  } catch (err) {
    console.log(
      `[sequel] waiting label sync failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
  }

  tryWriteCache(cacheKey, emails, searchTruncated);

  perf(`getPendingFollowUps (${emails.length}/${collected.length} threads)`, t0);

  return {
    emails,
    threadsScanned: collected.length,
    searchTruncated,
  };
}
