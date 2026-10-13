import {
  DEFAULT_AUTO_LABEL,
  DEFAULT_DAYS,
  DEFAULT_EXCLUDED_DOMAINS,
  DEFAULT_STALE_DAYS,
  MAX_FOLLOWUP_DAYS,
  MIN_FOLLOWUP_DAYS,
} from "./constants.ts";
import type { UserSettings } from "./types.ts";

export type SettingsPropertyBag = Record<string, string | null | undefined>;

function parsePositiveInt(raw: string | null | undefined, fallback: number): number {
  if (raw == null || raw.trim() === "") return fallback;
  const n = Number.parseInt(raw, 10);
  if (!Number.isFinite(n)) return fallback;
  return n;
}

export function clampFollowUpDays(days: number): number {
  return Math.max(MIN_FOLLOWUP_DAYS, Math.min(MAX_FOLLOWUP_DAYS, days));
}

export function clampStaleDays(staleDays: number, followUpDays: number): number {
  const minStale = Math.max(1, followUpDays + 1);
  return Math.max(minStale, staleDays);
}

export function normalizeUserSettings(partial: {
  days: number;
  excludeReplied: boolean;
  autoLabel: string;
  excludedDomains: string;
  staleDays: number;
  excludeInternal: boolean;
}): UserSettings {
  const days = clampFollowUpDays(partial.days);
  const staleDays = clampStaleDays(partial.staleDays, days);
  return {
    days,
    excludeReplied: partial.excludeReplied,
    autoLabel: partial.autoLabel.trim(),
    excludedDomains: partial.excludedDomains.trim(),
    staleDays,
    excludeInternal: partial.excludeInternal,
  };
}

export function parseSettingsFromProperties(props: SettingsPropertyBag): UserSettings {
  const daysRaw = parsePositiveInt(props.followup_days ?? null, DEFAULT_DAYS);
  const staleRaw = parsePositiveInt(props.stale_days ?? null, DEFAULT_STALE_DAYS);
  const excludeReplied = props.exclude_replied !== "false";
  const excludeInternal = props.exclude_internal === "true";
  const autoLabel = props.auto_label ?? DEFAULT_AUTO_LABEL;
  const excludedDomains = props.excluded_domains ?? DEFAULT_EXCLUDED_DOMAINS;

  return normalizeUserSettings({
    days: daysRaw,
    excludeReplied,
    autoLabel,
    excludedDomains,
    staleDays: staleRaw,
    excludeInternal,
  });
}

export function parseSettingsFromFormInput(
  formInput: Record<string, string>,
): UserSettings {
  const daysRaw = parsePositiveInt(formInput.followup_days, DEFAULT_DAYS);
  const staleRaw = parsePositiveInt(formInput.stale_days, DEFAULT_STALE_DAYS);
  return normalizeUserSettings({
    days: daysRaw,
    excludeReplied: formInput.exclude_replied === "true",
    autoLabel: formInput.auto_label ?? "",
    excludedDomains: formInput.excluded_domains ?? "",
    staleDays: staleRaw,
    excludeInternal: formInput.exclude_internal === "true",
  });
}

export function settingsToPropertyWrites(
  settings: UserSettings,
): Record<string, string> {
  return {
    followup_days: String(settings.days),
    stale_days: String(settings.staleDays),
    exclude_replied: settings.excludeReplied ? "true" : "false",
    exclude_internal: settings.excludeInternal ? "true" : "false",
    auto_label: settings.autoLabel,
    excluded_domains: settings.excludedDomains,
  };
}
