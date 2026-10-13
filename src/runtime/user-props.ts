import { DISMISSED_PROPERTY_KEY } from "../constants.ts";
import {
  addDismissedEntry,
  parseDismissedJson,
  removeDismissedEntry,
} from "../dismissed-store.ts";
import { log } from "../log.ts";
import { parseSettingsFromProperties, settingsToPropertyWrites } from "../settings.ts";
import type { DismissedEmail, UserSettings } from "../types.ts";

export function getUserPropertiesStore(): GoogleAppsScript.Properties.Properties {
  return PropertiesService.getUserProperties();
}

// Each getProperty call is a separate round trip (~30-55 ms in logs), so
// readers take one getProperties() snapshot instead of reading keys one by one.

export function readUserSettings(
  props: GoogleAppsScript.Properties.Properties = getUserPropertiesStore(),
): UserSettings {
  return parseSettingsFromProperties(props.getProperties());
}

export interface HomepageProps {
  settings: UserSettings;
  sortAsc: boolean;
  dismissed: DismissedEmail[];
}

/** Everything the homepage reads from user properties, in one round trip. */
export function readHomepageProps(
  props: GoogleAppsScript.Properties.Properties = getUserPropertiesStore(),
): HomepageProps {
  const all = props.getProperties();
  return {
    settings: parseSettingsFromProperties(all),
    sortAsc: all.sort_order !== "desc",
    dismissed: parseDismissedJson(all[DISMISSED_PROPERTY_KEY]),
  };
}

export function writeUserSettings(
  settings: UserSettings,
  props: GoogleAppsScript.Properties.Properties = getUserPropertiesStore(),
): void {
  try {
    props.setProperties(settingsToPropertyWrites(settings));
  } catch (err) {
    log.error("settings write failed", err);
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
    log.error("dismissed write failed", err);
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
