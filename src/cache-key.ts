import { MAX_CACHE_KEY_LENGTH } from "./constants.ts";
import type { PendingEmail, PendingEmailSerialized, UserSettings } from "./types.ts";

/** FNV-1a 32-bit — works in Apps Script and Node tests */
export function hashString(input: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < input.length; i++) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(36);
}

export function buildFollowUpCachePayload(settings: UserSettings): string {
  return [
    settings.days,
    settings.excludeReplied,
    settings.staleDays,
    settings.excludedDomains,
    settings.excludeInternal,
  ].join("|");
}

export function buildFollowUpCacheKey(settings: UserSettings): string {
  const payload = buildFollowUpCachePayload(settings);
  const key = `fu_${hashString(payload)}`;
  if (key.length <= MAX_CACHE_KEY_LENGTH) return key;
  return `fu_${hashString(key)}`;
}

export interface CachedFollowUps {
  emails: PendingEmail[];
  searchTruncated: boolean;
  fetchedAt: Date;
}

export function serializeCachedFollowUps(
  emails: PendingEmail[],
  searchTruncated: boolean,
  fetchedAt: Date,
): string {
  const serialized: PendingEmailSerialized[] = emails.map((e) => ({
    threadId: e.threadId,
    subject: e.subject,
    to: e.to,
    dateMs: e.date.getTime(),
  }));
  return JSON.stringify({
    emails: serialized,
    searchTruncated,
    fetchedAtMs: fetchedAt.getTime(),
  });
}

/** Returns null for malformed or legacy payloads (bare array, no fetch time) so callers refetch. */
export function parseCachedFollowUps(raw: string): CachedFollowUps | null {
  try {
    const parsed = JSON.parse(raw) as {
      emails?: PendingEmailSerialized[];
      searchTruncated?: boolean;
      fetchedAtMs?: number;
    };
    if (!parsed || !Array.isArray(parsed.emails)) return null;
    if (!Number.isFinite(parsed.fetchedAtMs)) return null;
    return {
      emails: parsed.emails.map((e) => ({
        threadId: e.threadId,
        subject: e.subject,
        to: e.to,
        date: new Date(e.dateMs),
      })),
      searchTruncated: parsed.searchTruncated === true,
      fetchedAt: new Date(parsed.fetchedAtMs as number),
    };
  } catch {
    return null;
  }
}
