import {
  MESSAGE_MODIFY_BATCH_SIZE,
  THREAD_FETCH_BATCH_SIZE,
  THREAD_FETCH_RETRY_DELAY_MS,
} from "../constants.ts";
import { type PaginatedSearchResult, paginateTokenSearch } from "../follow-ups-core.ts";
import {
  buildThreadMetadataUrl,
  describeHttpError,
  type GmailApiThread,
  isRetryableStatus,
  parseThreadMetadata,
  type ThreadMetadata,
} from "../gmail-api-core.ts";
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

function fetchBatch(
  ids: string[],
  token: string,
): ThreadMetadataResult & { retry: string[] } {
  const responses = UrlFetchApp.fetchAll(
    ids.map((id) => ({
      url: buildThreadMetadataUrl(id),
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true,
    })),
  );

  const threads: ThreadMetadata[] = [];
  const retry: string[] = [];
  let failed = 0;
  let firstError: string | undefined;
  responses.forEach((res, i) => {
    const status = res.getResponseCode();
    if (status === 200) {
      const parsed = parseThreadMetadata(
        JSON.parse(res.getContentText()) as GmailApiThread,
      );
      if (parsed) threads.push(parsed);
      else failed++;
    } else if (isRetryableStatus(status)) {
      retry.push(ids[i]);
    } else {
      // 404 means the thread was deleted between search and fetch
      if (status !== 404 && !firstError) {
        firstError = describeHttpError(status, res.getContentText());
        console.log(`[sequel] threads.get failed: ${firstError}`);
      }
      failed++;
    }
  });
  return { threads, retry, failed, firstError };
}

/**
 * Fetches From/To/Subject/date for each thread via the Gmail API (metadata
 * format, no bodies), issuing requests in parallel batches. Rate-limited or
 * 5xx requests are retried once after a short pause.
 */
export function fetchThreadMetadata(threadIds: string[]): ThreadMetadataResult {
  const token = ScriptApp.getOAuthToken();
  const threads: ThreadMetadata[] = [];
  const retry: string[] = [];
  let failed = 0;
  let firstError: string | undefined;

  for (const batch of chunk(threadIds, THREAD_FETCH_BATCH_SIZE)) {
    const result = fetchBatch(batch, token);
    threads.push(...result.threads);
    retry.push(...result.retry);
    failed += result.failed;
    firstError ??= result.firstError;
  }

  if (retry.length > 0) {
    console.log(`[sequel] retrying ${retry.length} threads.get requests`);
    Utilities.sleep(THREAD_FETCH_RETRY_DELAY_MS);
    for (const batch of chunk(retry, THREAD_FETCH_BATCH_SIZE)) {
      const result = fetchBatch(batch, token);
      threads.push(...result.threads);
      failed += result.failed + result.retry.length;
      firstError ??= result.firstError;
      if (result.retry.length > 0)
        firstError ??= "Gmail API rate limit or server error";
    }
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
