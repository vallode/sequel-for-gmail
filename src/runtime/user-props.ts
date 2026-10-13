import { DISMISSED_PROPERTY_KEY } from "../constants.ts";
import {
  addDismissedEntry,
  parseDismissedJson,
  removeDismissedEntry,
} from "../dismissed-store.ts";
import { parseSettingsFromProperties } from "../settings.ts";
import type { DismissedEmail, UserSettings } from "../types.ts";

export function getUserPropertiesStore(): GoogleAppsScript.Properties.Properties {
  return PropertiesService.getUserProperties();
}

export function readUserSettings(
  props: GoogleAppsScript.Properties.Properties = getUserPropertiesStore(),
): UserSettings {
  return parseSettingsFromProperties({
    followup_days: props.getProperty("followup_days"),
    stale_days: props.getProperty("stale_days"),
    exclude_replied: props.getProperty("exclude_replied"),
    exclude_internal: props.getProperty("exclude_internal"),
    auto_label: props.getProperty("auto_label"),
    excluded_domains: props.getProperty("excluded_domains"),
  });
}

export function writeUserSettings(
  settings: UserSettings,
  props: GoogleAppsScript.Properties.Properties = getUserPropertiesStore(),
): void {
  try {
    props.setProperty("followup_days", String(settings.days));
    props.setProperty("stale_days", String(settings.staleDays));
    props.setProperty("exclude_replied", settings.excludeReplied ? "true" : "false");
    props.setProperty("exclude_internal", settings.excludeInternal ? "true" : "false");
    props.setProperty("auto_label", settings.autoLabel);
    props.setProperty("excluded_domains", settings.excludedDomains);
  } catch (err) {
    console.log(
      `[sequel] settings write failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
    throw err;
  }
}

export function readDismissedList(
  props: GoogleAppsScript.Properties.Properties = getUserPropertiesStore(),
): DismissedEmail[] {
  try {
    return parseDismissedJson(props.getProperty(DISMISSED_PROPERTY_KEY));
  } catch {
    return [];
  }
}

export function writeDismissedList(
  items: DismissedEmail[],
  props: GoogleAppsScript.Properties.Properties = getUserPropertiesStore(),
): void {
  try {
    props.setProperty(DISMISSED_PROPERTY_KEY, JSON.stringify(items));
  } catch (err) {
    console.log(
      `[sequel] dismissed write failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
    throw err;
  }
}

export function withDismissedLock<T>(fn: () => T): T {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(30_000)) {
    throw new Error("Could not acquire dismissed-list lock");
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

export function dismissEmailEntry(entry: DismissedEmail): void {
  withDismissedLock(() => {
    const props = getUserPropertiesStore();
    const current = readDismissedList(props);
    const next = addDismissedEntry(current, entry);
    writeDismissedList(next, props);
  });
}

export function restoreDismissedThread(threadId: string): void {
  if (!threadId?.trim()) return;
  withDismissedLock(() => {
    const props = getUserPropertiesStore();
    const current = readDismissedList(props);
    writeDismissedList(removeDismissedEntry(current, threadId), props);
  });
}

export function clearDismissedList(): void {
  withDismissedLock(() => {
    getUserPropertiesStore().deleteProperty(DISMISSED_PROPERTY_KEY);
  });
}
