import {
  MESSAGE_MODIFY_BATCH_SIZE,
  THREAD_FETCH_BATCH_SIZE,
  THREAD_FETCH_RETRY_DELAYS_MS,
} from "../constants.ts";
import { type PaginatedSearchResult, paginateTokenSearch } from "../follow-ups-core.ts";
import {
  AUTO_FOLLOWUP_HEADERS,
  type AutoFollowUpThreadMetadata,
  boundaryFromContentType,
  buildBatchRequestBody,
  buildThreadMetadataPath,
  describeHttpError,
  FOLLOW_UP_HEADERS,
  GMAIL_BATCH_URL,
  type GmailApiThread,
  isRetryableStatus,
  parseAutoFollowUpThread,
  parseBatchResponse,
  parseThreadMetadata,
  type ThreadMetadata,
} from "../gmail-api-core.ts";
import { log } from "../log.ts";
import { chunk } from "../util.ts";

function gmailApi(): GoogleAppsScript.Gmail {
  if (typeof Gmail === "undefined" || !Gmail?.Users) {
    throw new Error("Advanced Gmail service is not enabled");
  }
  return Gmail;
}

export function searchThreadIds(query: string): PaginatedSearchResult<string> {
  return paginateTokenSearch((pageToken, maxResults) => {
    const args: { q: string; maxResults: number; pageToken?: string } = {
      q: query,
      maxResults,
    };
    if (pageToken) args.pageToken = pageToken;
    const res = gmailApi().Users.Threads.list("me", args);
    return {
      items: (res.threads ?? []).map((t) => t.id).filter((id): id is string => !!id),
      nextPageToken: res.nextPageToken,
    };
  });
}

export interface ThreadFetchResult<T> {
  threads: T[];
  /** Threads that still failed after the last retry */
  failed: number;
  /** First failure description, for logs and the error card */
  firstError?: string;
}

type ThreadRound<T> = ThreadFetchResult<T> & {
  retry: string[];
  /** Count of retryable responses by HTTP status, for logs */
  retryStatuses: Record<number, number>;
};

interface ThreadFetchSpec<T> {
  headers: readonly string[];
  parse: (raw: GmailApiThread) => T | null;
}

/**
 * One round of fetches: thread IDs are grouped into Gmail batch requests
 * (each carrying up to THREAD_FETCH_BATCH_SIZE threads.get calls) and the
 * batch requests are sent together with UrlFetchApp.fetchAll.
 */
function fetchRound<T>(
  ids: string[],
  token: string,
  spec: ThreadFetchSpec<T>,
): ThreadRound<T> {
  const round: ThreadRound<T> = {
    threads: [],
    retry: [],
    failed: 0,
    retryStatuses: {},
  };

  const record = (id: string, status: number, body: string) => {
    if (status === 200) {
      const parsed = spec.parse(JSON.parse(body) as GmailApiThread);
      if (parsed) round.threads.push(parsed);
      else round.failed++;
    } else if (isRetryableStatus(status)) {
      round.retry.push(id);
      round.retryStatuses[status] = (round.retryStatuses[status] ?? 0) + 1;
    } else {
      // 404 means the thread was deleted between search and fetch
      if (status !== 404 && !round.firstError) {
        round.firstError = describeHttpError(status, body);
        log.warn("threads.get failed", { threadId: id, status });
      }
      round.failed++;
    }
  };

  const groups = chunk(ids, THREAD_FETCH_BATCH_SIZE);
  const boundary = `sequel_batch_${Date.now()}`;
  const responses = UrlFetchApp.fetchAll(
    groups.map((group) => ({
      url: GMAIL_BATCH_URL,
      method: "post" as const,
      contentType: `multipart/mixed; boundary=${boundary}`,
      payload: buildBatchRequestBody(
        group.map((id) => buildThreadMetadataPath(id, spec.headers)),
        boundary,
      ),
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true,
    })),
  );

  responses.forEach((res, g) => {
    const group = groups[g];
    const status = res.getResponseCode();
    const text = res.getContentText();
    if (status !== 200) {
      // The whole batch request failed, so every thread in it did
      for (const id of group) record(id, status, text);
      return;
    }

    const headers = res.getHeaders() as Record<string, string>;
    const contentType = headers["Content-Type"] ?? headers["content-type"] ?? "";
    const responseBoundary =
      boundaryFromContentType(contentType) ?? text.match(/^--(\S+)/)?.[1] ?? null;
    const parts = responseBoundary ? parseBatchResponse(text, responseBoundary) : [];

    const answered = new Set<number>();
    for (const part of parts) {
      const id = group[part.index];
      if (id === undefined || answered.has(part.index)) continue;
      answered.add(part.index);
      record(id, part.status, part.body);
    }
    const missing = group.length - answered.size;
    if (missing > 0) {
      round.failed += missing;
      round.firstError ??= "Gmail batch response was missing results";
      log.warn("batch response missing parts", { missing });
    }
  });

  return round;
}

/**
 * Fetches thread metadata (the spec's headers plus dates and label IDs, no
 * bodies) via Gmail batch requests. Rate-limited or 5xx calls are retried
 * after each delay in THREAD_FETCH_RETRY_DELAYS_MS.
 */
function fetchThreads<T>(
  threadIds: string[],
  spec: ThreadFetchSpec<T>,
): ThreadFetchResult<T> {
  if (threadIds.length === 0) return { threads: [], failed: 0 };
  const token = ScriptApp.getOAuthToken();
  let round = fetchRound(threadIds, token, spec);
  const threads = round.threads;
  let failed = round.failed;
  let firstError = round.firstError;

  for (const delayMs of THREAD_FETCH_RETRY_DELAYS_MS) {
    if (round.retry.length === 0) break;
    log.warn("retrying threads.get requests", {
      count: round.retry.length,
      statuses: round.retryStatuses,
      delayMs,
    });
    Utilities.sleep(delayMs);
    round = fetchRound(round.retry, token, spec);
    threads.push(...round.threads);
    failed += round.failed;
    firstError ??= round.firstError;
  }
  if (round.retry.length > 0) {
    failed += round.retry.length;
    firstError ??= "Gmail API rate limit or server error";
  }

  return { threads, failed, firstError };
}

/** From/To/Subject/date for the follow-up list. */
export function fetchThreadMetadata(
  threadIds: string[],
): ThreadFetchResult<ThreadMetadata> {
  return fetchThreads(threadIds, {
    headers: FOLLOW_UP_HEADERS,
    parse: parseThreadMetadata,
  });
}

/** Sender, recipients, draft/trash state and reply headers for flagged threads. */
export function fetchAutoFollowUpThreads(
  threadIds: string[],
): ThreadFetchResult<AutoFollowUpThreadMetadata> {
  return fetchThreads(threadIds, {
    headers: AUTO_FOLLOWUP_HEADERS,
    parse: parseAutoFollowUpThread,
  });
}

const labelIds = new Map<string, string>();

/** Gmail API label ID for a user label name (case-insensitive), cached per execution. */
function getLabelId(labelName: string): string {
  const key = labelName.toLowerCase();
  const cached = labelIds.get(key);
  if (cached) return cached;
  for (const label of gmailApi().Users.Labels.list("me").labels ?? []) {
    if (label.name && label.id) labelIds.set(label.name.toLowerCase(), label.id);
  }
  const id = labelIds.get(key);
  if (!id) throw new Error(`Label not found: ${labelName}`);
  return id;
}

/**
 * IDs of threads carrying the label. Threads.list skips trash and spam by
 * default, so no per-thread trash check is needed.
 */
export function listThreadIdsWithLabel(
  labelName: string,
): PaginatedSearchResult<string> {
  const labelId = getLabelId(labelName);
  const result = paginateTokenSearch((pageToken, maxResults) => {
    const args: { labelIds: string[]; maxResults: number; pageToken?: string } = {
      labelIds: [labelId],
      maxResults,
    };
    if (pageToken) args.pageToken = pageToken;
    const res = gmailApi().Users.Threads.list("me", args);
    return {
      items: (res.threads ?? []).map((t) => t.id).filter((id): id is string => !!id),
      nextPageToken: res.nextPageToken,
    };
  });
  if (result.truncated) log.warn("label listing truncated", { label: labelName });
  return result;
}

/** Labels messages in bulk; `labelName` must already exist. */
export function addLabelToMessages(labelName: string, messageIds: string[]): void {
  if (messageIds.length === 0) return;
  const labelId = getLabelId(labelName);
  for (const ids of chunk(messageIds, MESSAGE_MODIFY_BATCH_SIZE)) {
    gmailApi().Users.Messages.batchModify({ ids, addLabelIds: [labelId] }, "me");
  }
}

/** Adds and removes labels on one thread in a single call; labels must exist. */
export function modifyThreadLabels(
  threadId: string,
  change: { add?: string[]; remove?: string[] },
): void {
  gmailApi().Users.Threads.modify(
    {
      addLabelIds: (change.add ?? []).map(getLabelId),
      removeLabelIds: (change.remove ?? []).map(getLabelId),
    },
    "me",
    threadId,
  );
}

/** Removes a label from whole threads; usually only a few per refresh. */
export function removeLabelFromThreadIds(labelName: string, threadIds: string[]): void {
  for (const id of threadIds) modifyThreadLabels(id, { remove: [labelName] });
}

/** Saves a raw RFC 2822 message as a draft in the thread. Never sends. */
export function createThreadDraft(threadId: string, mime: string): void {
  gmailApi().Users.Drafts.create(
    { message: { raw: Utilities.base64EncodeWebSafe(mime), threadId } },
    "me",
  );
}
