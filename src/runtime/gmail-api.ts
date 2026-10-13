import {
  MESSAGE_MODIFY_BATCH_SIZE,
  THREAD_FETCH_BATCH_SIZE,
  THREAD_FETCH_RETRY_DELAYS_MS,
} from "../constants.ts";
import { type PaginatedSearchResult, paginateTokenSearch } from "../follow-ups-core.ts";
import {
  boundaryFromContentType,
  buildBatchRequestBody,
  buildThreadMetadataPath,
  describeHttpError,
  GMAIL_BATCH_URL,
  type GmailApiThread,
  isRetryableStatus,
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

export interface ThreadMetadataResult {
  threads: ThreadMetadata[];
  /** Threads that still failed after the last retry */
  failed: number;
  /** First failure description, for logs and the error card */
  firstError?: string;
}

type ThreadRound = ThreadMetadataResult & {
  retry: string[];
  /** Count of retryable responses by HTTP status, for logs */
  retryStatuses: Record<number, number>;
};

/**
 * One round of fetches: thread IDs are grouped into Gmail batch requests
 * (each carrying up to THREAD_FETCH_BATCH_SIZE threads.get calls) and the
 * batch requests are sent together with UrlFetchApp.fetchAll.
 */
function fetchRound(ids: string[], token: string): ThreadRound {
  const round: ThreadRound = { threads: [], retry: [], failed: 0, retryStatuses: {} };

  const record = (id: string, status: number, body: string) => {
    if (status === 200) {
      const parsed = parseThreadMetadata(JSON.parse(body) as GmailApiThread);
      if (parsed) round.threads.push(parsed);
      else round.failed++;
    } else if (isRetryableStatus(status)) {
      round.retry.push(id);
      round.retryStatuses[status] = (round.retryStatuses[status] ?? 0) + 1;
    } else {
      // 404 means the thread was deleted between search and fetch
      if (status !== 404 && !round.firstError) {
        round.firstError = describeHttpError(status, body);
        log.warn("threads.get failed", undefined, { threadId: id, status });
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
      payload: buildBatchRequestBody(group.map(buildThreadMetadataPath), boundary),
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
      log.warn("batch response missing parts", undefined, { missing });
    }
  });

  return round;
}

/**
 * Fetches From/To/Subject/date for each thread via the Gmail API (metadata
 * format, no bodies) using batch requests. Rate-limited or 5xx calls are
 * retried after each delay in THREAD_FETCH_RETRY_DELAYS_MS.
 */
export function fetchThreadMetadata(threadIds: string[]): ThreadMetadataResult {
  const token = ScriptApp.getOAuthToken();
  let round = fetchRound(threadIds, token);
  const threads = round.threads;
  let failed = round.failed;
  let firstError = round.firstError;

  for (const delayMs of THREAD_FETCH_RETRY_DELAYS_MS) {
    if (round.retry.length === 0) break;
    log.warn("retrying threads.get requests", undefined, {
      count: round.retry.length,
      statuses: round.retryStatuses,
      delayMs,
    });
    Utilities.sleep(delayMs);
    round = fetchRound(round.retry, token);
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
  return paginateTokenSearch((pageToken, maxResults) => {
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
}

/** Labels messages in bulk; `labelName` must already exist. */
export function addLabelToMessages(labelName: string, messageIds: string[]): void {
  if (messageIds.length === 0) return;
  const labelId = getLabelId(labelName);
  for (const ids of chunk(messageIds, MESSAGE_MODIFY_BATCH_SIZE)) {
    gmailApi().Users.Messages.batchModify({ ids, addLabelIds: [labelId] }, "me");
  }
}

/** Removes a label from whole threads; usually only a few per refresh. */
export function removeLabelFromThreadIds(labelName: string, threadIds: string[]): void {
  if (threadIds.length === 0) return;
  const labelId = getLabelId(labelName);
  for (const id of threadIds) {
    gmailApi().Users.Threads.modify({ removeLabelIds: [labelId] }, "me", id);
  }
}
