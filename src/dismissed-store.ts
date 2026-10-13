import { DISMISSED_CAP } from "./constants.ts";
import type { DismissedEmail } from "./types.ts";

export function parseDismissedJson(raw: string | null | undefined): DismissedEmail[] {
  if (!raw || raw.trim() === "") return [];
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) return [];
    const result: DismissedEmail[] = [];
    for (const item of parsed) {
      if (item == null || typeof item !== "object") continue;
      const record = item as Record<string, unknown>;
      const threadId = record.threadId;
      if (typeof threadId !== "string" || !threadId.trim()) continue;
      const subject = typeof record.subject === "string" ? record.subject : "";
      const to = typeof record.to === "string" ? record.to : "";
      const dateMs =
        typeof record.dateMs === "number" && Number.isFinite(record.dateMs)
          ? record.dateMs
          : 0;
      result.push({ threadId, subject, to, dateMs });
    }
    return result;
  } catch {
    return [];
  }
}

export function capDismissedList(items: DismissedEmail[]): DismissedEmail[] {
  if (items.length <= DISMISSED_CAP) return items;
  return items.slice(items.length - DISMISSED_CAP);
}

export function addDismissedEntry(
  items: DismissedEmail[],
  entry: DismissedEmail,
): DismissedEmail[] {
  if (items.some((d) => d.threadId === entry.threadId)) return items;
  return capDismissedList([...items, entry]);
}

export function removeDismissedEntry(
  items: DismissedEmail[],
  threadId: string,
): DismissedEmail[] {
  if (!threadId) return items;
  return items.filter((d) => d.threadId !== threadId);
}

export function dismissedThreadIds(items: DismissedEmail[]): Set<string> {
  return new Set(items.map((d) => d.threadId));
}
