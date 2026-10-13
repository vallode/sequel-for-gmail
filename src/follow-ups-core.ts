import { GMAIL_SEARCH_PAGE_SIZE, MAX_THREADS_TO_SCAN } from "./constants.ts";
import {
  getEmailDomain,
  isAnyDomainExcluded,
  isInternalRecipient,
  parseFirstRecipient,
  senderIsUser,
} from "./email-utils.ts";
import type {
  FollowUpFilterOptions,
  FollowUpQueryResult,
  PendingEmail,
  ThreadInput,
} from "./types.ts";

export function sortPendingEmailsAsc(emails: PendingEmail[]): PendingEmail[] {
  return [...emails].sort((a, b) => a.date.getTime() - b.date.getTime());
}

export function applySortOrder(
  emails: PendingEmail[],
  sortAsc: boolean,
): PendingEmail[] {
  return sortAsc ? emails : [...emails].reverse();
}

export function evaluateThreadForFollowUp(
  thread: ThreadInput,
  options: FollowUpFilterOptions,
): PendingEmail | null {
  const messages = thread.messages;
  if (messages.length === 0) return null;

  const lastMessage = messages[messages.length - 1];
  const lastFromMe = senderIsUser(lastMessage.from, options.myEmail);

  let lastSent: (typeof messages)[number] | null = null;
  let hasReply = false;

  if (lastFromMe) {
    lastSent = lastMessage;
    hasReply = false;
  } else {
    hasReply = true;
  }

  if (!lastSent) return null;
  if (options.excludeReplied && hasReply) return null;

  const sentDate = lastSent.date;
  if (sentDate > options.cutoff || sentDate <= options.staleCutoff) return null;

  const recipient = parseFirstRecipient(lastSent.to);
  if (!recipient) return null;

  if (
    options.excludeInternal &&
    isInternalRecipient(recipient, options.myEmail, options.myDomain)
  ) {
    return null;
  }

  const toDomain = getEmailDomain(recipient);
  if (
    options.excludedDomainList.length > 0 &&
    isAnyDomainExcluded(toDomain, options.excludedDomainList)
  ) {
    return null;
  }

  const subject = messages[0]?.subject?.trim() || "(no subject)";

  return {
    threadId: thread.threadId,
    subject,
    to: recipient,
    date: new Date(sentDate.getTime()),
  };
}

export interface PaginatedSearchResult<T> {
  items: T[];
  truncated: boolean;
}

/**
 * Pages through a search until it runs dry or hits `maxItems`. When the cap is
 * reached, probes one more result so `truncated` is only set if more exist.
 */
export function paginateSearch<T>(
  search: (start: number, max: number) => T[],
  pageSize: number = GMAIL_SEARCH_PAGE_SIZE,
  maxItems: number = MAX_THREADS_TO_SCAN,
): PaginatedSearchResult<T> {
  const items: T[] = [];
  let start = 0;
  while (items.length < maxItems) {
    const size = Math.min(pageSize, maxItems - items.length);
    const page = search(start, size);
    items.push(...page);
    if (page.length < size) return { items, truncated: false };
    start += size;
  }
  return { items, truncated: search(start, 1).length > 0 };
}

export function filterThreadsToPending(
  threads: ThreadInput[],
  options: FollowUpFilterOptions,
): PendingEmail[] {
  const results: PendingEmail[] = [];
  for (const thread of threads) {
    const pending = evaluateThreadForFollowUp(thread, options);
    if (pending) results.push(pending);
  }
  return sortPendingEmailsAsc(results);
}

export function mergeFollowUpResult(
  emails: PendingEmail[],
  threadsScanned: number,
  searchTruncated: boolean,
): FollowUpQueryResult {
  return { emails, threadsScanned, searchTruncated };
}
