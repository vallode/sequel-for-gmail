import type { AutoFollowUpThreadInput, ThreadInput } from "./types.ts";

/** Gmail's batch endpoint: many API calls in one multipart HTTP request. */
export const GMAIL_BATCH_URL = "https://gmail.googleapis.com/batch/gmail/v1";

/** Headers the follow-up filter reads; bodies are never fetched. */
export const FOLLOW_UP_HEADERS = ["From", "To", "Subject"] as const;
/** Headers the auto follow-up runner reads, including what a threaded reply needs. */
export const AUTO_FOLLOWUP_HEADERS = [
  "From",
  "To",
  "Cc",
  "Subject",
  "Message-ID",
  "References",
] as const;
const THREAD_METADATA_FIELDS = "id,messages(id,internalDate,labelIds,payload/headers)";

export interface GmailApiHeader {
  name?: string;
  value?: string;
}

export interface GmailApiMessage {
  id?: string;
  internalDate?: string;
  labelIds?: string[];
  payload?: { headers?: GmailApiHeader[] };
}

export interface GmailApiThread {
  id?: string;
  messages?: GmailApiMessage[];
}

export interface ThreadMetadata {
  thread: ThreadInput;
  messageIds: string[];
}

/** Request path for one thread's metadata, relative to the API host. */
export function buildThreadMetadataPath(
  threadId: string,
  headers: readonly string[] = FOLLOW_UP_HEADERS,
): string {
  const params = [
    "format=metadata",
    ...headers.map((h) => `metadataHeaders=${h}`),
    `fields=${encodeURIComponent(THREAD_METADATA_FIELDS)}`,
  ];
  return `/gmail/v1/users/me/threads/${encodeURIComponent(threadId)}?${params.join("&")}`;
}

/**
 * multipart/mixed body for the batch endpoint: one GET per path, tagged with
 * its index as Content-ID so responses can be matched back.
 */
export function buildBatchRequestBody(paths: string[], boundary: string): string {
  const parts = paths.map(
    (path, i) =>
      `--${boundary}\r\nContent-Type: application/http\r\nContent-ID: <item-${i}>\r\n\r\nGET ${path}\r\n`,
  );
  return `${parts.join("\r\n")}\r\n--${boundary}--`;
}

export interface BatchPartResponse {
  /** Index of the request this answers, from its Content-ID */
  index: number;
  status: number;
  body: string;
}

/** Reads the multipart boundary from a Content-Type header, if present. */
export function boundaryFromContentType(contentType: string): string | null {
  const match = contentType.match(/boundary="?([^";]+)"?/i);
  return match ? match[1] : null;
}

/**
 * Splits a batch response into per-request status and body. Parts without a
 * recognisable Content-ID or status line are skipped; callers treat requests
 * with no matching part as failed.
 */
export function parseBatchResponse(
  body: string,
  boundary: string,
): BatchPartResponse[] {
  const results: BatchPartResponse[] = [];
  const normalized = body.replace(/\r\n/g, "\n");
  for (const part of normalized.split(`--${boundary}`)) {
    const id = part.match(/^Content-ID:\s*<response-item-(\d+)>/im);
    const status = part.match(/^HTTP\/[\d.]+\s+(\d{3})/m);
    if (!id || !status || status.index === undefined) continue;
    const afterStatus = part.slice(status.index);
    const bodyStart = afterStatus.indexOf("\n\n");
    results.push({
      index: Number(id[1]),
      status: Number(status[1]),
      body: bodyStart === -1 ? "" : afterStatus.slice(bodyStart + 2).trim(),
    });
  }
  return results;
}

/** Header names are case-insensitive; returns "" when the header is absent. */
export function findHeader(
  headers: GmailApiHeader[] | undefined,
  name: string,
): string {
  const wanted = name.toLowerCase();
  const match = headers?.find((h) => h.name?.toLowerCase() === wanted);
  return match?.value ?? "";
}

/** Returns null when the response has no thread ID, so callers treat it as failed. */
export function parseThreadMetadata(raw: GmailApiThread): ThreadMetadata | null {
  if (!raw.id) return null;
  const messages = raw.messages ?? [];
  return {
    thread: {
      threadId: raw.id,
      messages: messages.map((msg) => ({
        from: findHeader(msg.payload?.headers, "From"),
        to: findHeader(msg.payload?.headers, "To"),
        date: new Date(Number(msg.internalDate ?? 0)),
        subject: findHeader(msg.payload?.headers, "Subject"),
      })),
    },
    messageIds: messages.map((msg) => msg.id).filter((id): id is string => !!id),
  };
}

export interface AutoFollowUpThreadMetadata {
  thread: AutoFollowUpThreadInput;
  /** Headers of the last message, for threading a reply to it */
  replyTo: { messageId: string; references: string };
}

/**
 * Runner input for one flagged thread. Trashed messages are dropped (Gmail
 * still returns them in threads.get) and DRAFT-labelled ones are marked.
 */
export function parseAutoFollowUpThread(
  raw: GmailApiThread,
): AutoFollowUpThreadMetadata | null {
  if (!raw.id) return null;
  const messages = (raw.messages ?? []).filter((m) => !m.labelIds?.includes("TRASH"));
  const lastHeaders = messages[messages.length - 1]?.payload?.headers;
  return {
    thread: {
      threadId: raw.id,
      messages: messages.map((msg) => ({
        from: findHeader(msg.payload?.headers, "From"),
        to: findHeader(msg.payload?.headers, "To"),
        cc: findHeader(msg.payload?.headers, "Cc"),
        date: new Date(Number(msg.internalDate ?? 0)),
        subject: findHeader(msg.payload?.headers, "Subject"),
        isDraft: msg.labelIds?.includes("DRAFT") ?? false,
      })),
    },
    replyTo: {
      messageId: findHeader(lastHeaders, "Message-ID"),
      references: findHeader(lastHeaders, "References"),
    },
  };
}

/** Rate limits (429) and server errors (5xx) are worth one retry. */
export function isRetryableStatus(status: number): boolean {
  return status === 429 || status >= 500;
}

/** Short "status: message" from a Google API error body, for logs and the UI. */
export function describeHttpError(status: number, body: string): string {
  let message = "";
  try {
    message =
      (JSON.parse(body) as { error?: { message?: string } }).error?.message ?? "";
  } catch {
    // Non-JSON body; fall back to the status alone
  }
  const text = message ? `${status}: ${message}` : `HTTP ${status}`;
  return text.length > 200 ? `${text.slice(0, 199)}…` : text;
}
