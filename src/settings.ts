import {
  normalizeAutoFollowUpSteps,
  parseAutoFollowUpSteps,
  serializeAutoFollowUpSteps,
} from "./auto-follow-up-core.ts";
import {
  AUTO_FOLLOWUP_STEPS_PROPERTY_KEY,
  DEFAULT_AUTO_FOLLOWUP_STEPS,
  DEFAULT_DAYS,
  DEFAULT_EXCLUDED_DOMAINS,
  DEFAULT_STALE_DAYS,
  MAX_AUTO_FOLLOWUP_STEPS,
  MAX_FOLLOWUP_DAYS,
  MIN_FOLLOWUP_DAYS,
} from "./constants.ts";
import type { AutoFollowUpStep, UserSettings } from "./types.ts";

export type SettingsPropertyBag = Record<string, string | null | undefined>;

/** Property keys read by `parseSettingsFromProperties` */
export const SETTINGS_PROPERTY_KEYS = [
  "followup_days",
  "stale_days",
  "exclude_replied",
  "exclude_internal",
  "excluded_domains",
  "auto_followup_enabled",
  AUTO_FOLLOWUP_STEPS_PROPERTY_KEY,
] as const;

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

export function defaultAutoFollowUpSteps(): AutoFollowUpStep[] {
  return DEFAULT_AUTO_FOLLOWUP_STEPS.map((step) => ({ ...step }));
}

export function normalizeUserSettings(partial: {
  days: number;
  excludeReplied: boolean;
  excludedDomains: string;
  staleDays: number;
  excludeInternal: boolean;
  autoFollowUpEnabled: boolean;
  autoFollowUpSteps: unknown[];
}): UserSettings {
  const days = clampFollowUpDays(partial.days);
  const staleDays = clampStaleDays(partial.staleDays, days);
  return {
    days,
    excludeReplied: partial.excludeReplied,
    excludedDomains: partial.excludedDomains.trim(),
    staleDays,
    excludeInternal: partial.excludeInternal,
    autoFollowUpEnabled: partial.autoFollowUpEnabled,
    autoFollowUpSteps: normalizeAutoFollowUpSteps(partial.autoFollowUpSteps),
  };
}

export function parseSettingsFromProperties(props: SettingsPropertyBag): UserSettings {
  const daysRaw = parsePositiveInt(props.followup_days ?? null, DEFAULT_DAYS);
  const staleRaw = parsePositiveInt(props.stale_days ?? null, DEFAULT_STALE_DAYS);
  const excludeReplied = props.exclude_replied !== "false";
  const excludeInternal = props.exclude_internal === "true";
  const excludedDomains = props.excluded_domains ?? DEFAULT_EXCLUDED_DOMAINS;
  // Never-saved sequences get the defaults; a saved empty list stays empty.
  const stepsRaw = props[AUTO_FOLLOWUP_STEPS_PROPERTY_KEY];
  const autoFollowUpSteps =
    stepsRaw == null ? defaultAutoFollowUpSteps() : parseAutoFollowUpSteps(stepsRaw);

  return normalizeUserSettings({
    days: daysRaw,
    excludeReplied,
    excludedDomains,
    staleDays: staleRaw,
    excludeInternal,
    autoFollowUpEnabled: props.auto_followup_enabled === "true",
    autoFollowUpSteps,
  });
}

/** Settings-form field names for sequence slot `index` (0-based). */
export function autoFollowUpFieldNames(index: number): {
  waitDays: string;
  template: string;
} {
  return {
    waitDays: `auto_followup_wait_${index + 1}`,
    template: `auto_followup_template_${index + 1}`,
  };
}

function parseAutoFollowUpStepsFromForm(
  formInput: Record<string, string>,
): AutoFollowUpStep[] {
  const fallback = defaultAutoFollowUpSteps();
  const raw: AutoFollowUpStep[] = [];
  for (let i = 0; i < MAX_AUTO_FOLLOWUP_STEPS; i++) {
    const fields = autoFollowUpFieldNames(i);
    raw.push({
      waitDays: parsePositiveInt(
        formInput[fields.waitDays],
        fallback[i]?.waitDays ?? DEFAULT_DAYS,
      ),
      template: formInput[fields.template] ?? "",
    });
  }
  return normalizeAutoFollowUpSteps(raw);
}

export function parseSettingsFromFormInput(
  formInput: Record<string, string>,
): UserSettings {
  const daysRaw = parsePositiveInt(formInput.followup_days, DEFAULT_DAYS);
  const staleRaw = parsePositiveInt(formInput.stale_days, DEFAULT_STALE_DAYS);
  return normalizeUserSettings({
    days: daysRaw,
    excludeReplied: formInput.exclude_replied === "true",
    excludedDomains: formInput.excluded_domains ?? "",
    staleDays: staleRaw,
    excludeInternal: formInput.exclude_internal === "true",
    autoFollowUpEnabled: formInput.auto_followup_enabled === "true",
    autoFollowUpSteps: parseAutoFollowUpStepsFromForm(formInput),
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
    excluded_domains: settings.excludedDomains,
    auto_followup_enabled: settings.autoFollowUpEnabled ? "true" : "false",
    [AUTO_FOLLOWUP_STEPS_PROPERTY_KEY]: serializeAutoFollowUpSteps(
      settings.autoFollowUpSteps,
    ),
  };
}
