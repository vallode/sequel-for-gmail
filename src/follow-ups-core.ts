import { GMAIL_API_LIST_PAGE_SIZE, MAX_THREADS_TO_SCAN } from "./constants.ts";
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

export interface TokenPage<T> {
  items: T[];
  nextPageToken?: string;
}

/**
 * Pages through a Gmail API list via page tokens. Collects one item
 * past `maxItems` so `truncated` is only set if more results actually exist.
 */
export function paginateTokenSearch<T>(
  fetchPage: (pageToken: string | undefined, max: number) => TokenPage<T>,
  pageSize: number = GMAIL_API_LIST_PAGE_SIZE,
  maxItems: number = MAX_THREADS_TO_SCAN,
): PaginatedSearchResult<T> {
  const items: T[] = [];
  let pageToken: string | undefined;
  do {
    const page = fetchPage(pageToken, Math.min(pageSize, maxItems + 1 - items.length));
    items.push(...page.items);
    pageToken = page.nextPageToken;
  } while (pageToken && items.length <= maxItems);
  return { items: items.slice(0, maxItems), truncated: items.length > maxItems };
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
  fetchedAt: Date,
): FollowUpQueryResult {
  return { emails, threadsScanned, searchTruncated, fetchedAt };
}
