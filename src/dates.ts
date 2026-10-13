export function formatDateForQuery(date: Date): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}/${m}/${d}`;
}

export function getDayAge(date: Date): number {
  return Math.floor((Date.now() - date.getTime()) / 86400000);
}

export function buildSearchQuery(staleCutoff: Date, cutoff: Date): string {
  return `in:sent after:${formatDateForQuery(staleCutoff)} before:${formatDateForQuery(
    cutoff,
  )}`;
}

export function formatUpdatedLabel(fetchedAt: Date, now: Date = new Date()): string {
  const minutes = Math.floor((now.getTime() - fetchedAt.getTime()) / 60000);
  if (minutes < 1) return "Updated just now";
  if (minutes === 1) return "Updated 1 min ago";
  return `Updated ${minutes} min ago`;
}
