/**
 * Names from `wanted` with no existing label, in `wanted` order. Gmail label
 * names are case-insensitive, so `follow-ups` satisfies `Follow-ups`.
 */
export function missingLabelNames(
  wanted: readonly string[],
  existing: readonly string[],
): string[] {
  const have = new Set(existing.map((name) => name.toLowerCase()));
  return wanted.filter((name) => !have.has(name.toLowerCase()));
}

export interface LabelMembershipChanges {
  add: string[];
  remove: string[];
}

/**
 * Thread IDs to label and unlabel so `current` matches `desired`. Removal is
 * skipped when `desired` may be incomplete (e.g. a truncated search).
 */
export function diffLabelMembership(
  desired: ReadonlySet<string>,
  current: ReadonlySet<string>,
  desiredIsComplete: boolean,
): LabelMembershipChanges {
  return {
    add: [...desired].filter((id) => !current.has(id)),
    remove: desiredIsComplete ? [...current].filter((id) => !desired.has(id)) : [],
  };
}
