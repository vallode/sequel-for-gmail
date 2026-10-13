import {
  MESSAGE_MODIFY_BATCH_SIZE,
  THREAD_FETCH_BATCH_SIZE,
  THREAD_FETCH_RETRY_DELAY_MS,
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
import { log, perf } from "../log.ts";
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
  /** Threads that still failed after one retry */
  failed: number;
  /** First failure description, for logs and the error card */
  firstError?: string;
}

type ThreadRound = ThreadMetadataResult & { retry: string[] };

/**
 * One round of fetches: thread IDs are grouped into Gmail batch requests
 * (each carrying up to THREAD_FETCH_BATCH_SIZE threads.get calls) and the
 * batch requests are sent together with UrlFetchApp.fetchAll.
 */
function fetchRound(ids: string[], token: string): ThreadRound {
  const round: ThreadRound = { threads: [], retry: [], failed: 0 };

  const record = (id: string, status: number, body: string) => {
    if (status === 200) {
      const parsed = parseThreadMetadata(JSON.parse(body) as GmailApiThread);
      if (parsed) round.threads.push(parsed);
      else round.failed++;
    } else if (isRetryableStatus(status)) {
      round.retry.push(id);
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
  const t0 = Date.now();
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
  perf("threads.get batch requests", t0, {
    requests: groups.length,
    threads: ids.length,
  });

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
 * retried once after a short pause.
 */
export function fetchThreadMetadata(threadIds: string[]): ThreadMetadataResult {
  const token = ScriptApp.getOAuthToken();
  const first = fetchRound(threadIds, token);
  const threads = first.threads;
  let failed = first.failed;
  let firstError = first.firstError;

  if (first.retry.length > 0) {
    log.warn("retrying threads.get requests", undefined, { count: first.retry.length });
    Utilities.sleep(THREAD_FETCH_RETRY_DELAY_MS);
    const second = fetchRound(first.retry, token);
    threads.push(...second.threads);
    failed += second.failed + second.retry.length;
    firstError ??= second.firstError;
    if (second.retry.length > 0) firstError ??= "Gmail API rate limit or server error";
  }

  return { threads, failed, firstError };
}

/** Labels messages in bulk; `labelName` must already exist. */
export function addLabelToMessages(labelName: string, messageIds: string[]): void {
  if (messageIds.length === 0) return;
  const api = gmailApi();
  const wanted = labelName.toLowerCase();
  const label = (api.Users.Labels.list("me").labels ?? []).find(
    (l) => l.name?.toLowerCase() === wanted,
  );
  if (!label?.id) throw new Error(`Label not found: ${labelName}`);
  for (const ids of chunk(messageIds, MESSAGE_MODIFY_BATCH_SIZE)) {
    api.Users.Messages.batchModify({ ids, addLabelIds: [label.id] }, "me");
  }
}
