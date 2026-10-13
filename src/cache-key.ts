import { MAX_CACHE_KEY_LENGTH } from "./constants.ts";
import type { UserSettings } from "./types.ts";

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
    settings.autoLabel,
    settings.excludeInternal,
  ].join("|");
}

export function buildFollowUpCacheKey(settings: UserSettings): string {
  const payload = buildFollowUpCachePayload(settings);
  const key = `fu_${hashString(payload)}`;
  if (key.length <= MAX_CACHE_KEY_LENGTH) return key;
  return `fu_${hashString(key)}`;
}
