import type { ThreadInput } from "./types.ts";

const GMAIL_API_BASE = "https://gmail.googleapis.com/gmail/v1/users/me";

/** Only the headers the follow-up filter reads; bodies are never fetched. */
const METADATA_HEADERS = ["From", "To", "Subject"];
const THREAD_METADATA_FIELDS = "id,messages(id,internalDate,payload/headers)";

export interface GmailApiHeader {
  name?: string;
  value?: string;
}

export interface GmailApiThread {
  id?: string;
  messages?: Array<{
    id?: string;
    internalDate?: string;
    payload?: { headers?: GmailApiHeader[] };
  }>;
}

export interface ThreadMetadata {
  thread: ThreadInput;
  messageIds: string[];
}

export function buildThreadMetadataUrl(threadId: string): string {
  const params = [
    "format=metadata",
    ...METADATA_HEADERS.map((h) => `metadataHeaders=${h}`),
    `fields=${encodeURIComponent(THREAD_METADATA_FIELDS)}`,
  ];
  return `${GMAIL_API_BASE}/threads/${encodeURIComponent(threadId)}?${params.join("&")}`;
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
