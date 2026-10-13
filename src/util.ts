export function truncate(str: string, max: number): string {
  return str.length > max ? `${str.slice(0, max - 1)}…` : str;
}

export function perf(label: string, startMs: number): void {
  console.log(`[sequel] ${label}: ${Date.now() - startMs}ms`);
}
