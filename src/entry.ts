/// <reference path="./card-service-extensions.d.ts" />
import {
  runAutoFollowUps,
  setThreadAutoFollowUp,
  syncAutoFollowUpTrigger,
} from "./auto-follow-up-runner.ts";
import { buildDismissedCard, buildHomepageCard, buildSettingsCard } from "./cards.ts";
import {
  clearDismissedList,
  dismissEmailEntry,
  restoreDismissedThread,
  writeUserSettings,
} from "./runtime/user-props.ts";
import { parseSettingsFromFormInput } from "./settings.ts";
import type { ActionEvent, DismissedEmail } from "./types.ts";

export function _buildHomepage(
  e: ActionEvent,
  ignoreCache = false,
): GoogleAppsScript.Card_Service.Card {
  return buildHomepageCard(e, ignoreCache);
}

export { buildDismissedCard, buildSettingsCard };

function saveSettingsNotice(enabled: boolean, hasSteps: boolean, ok: boolean): string {
  if (!ok) return "Settings saved, but auto follow-ups could not be scheduled";
  if (enabled && !hasSteps) {
    return "Settings saved. Add a template to start auto follow-ups";
  }
  return "Settings saved";
}

export function _onSaveSettings(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const settings = parseSettingsFromFormInput(e.formInput ?? {});
  writeUserSettings(settings);

  let scheduled = true;
  try {
    syncAutoFollowUpTrigger(settings.autoFollowUpEnabled);
  } catch (err) {
    scheduled = false;
    console.log(
      `[sequel] trigger sync failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
  }

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildHomepage(e, true)))
    .setNotification(
      CardService.newNotification().setText(
        saveSettingsNotice(
          settings.autoFollowUpEnabled,
          settings.autoFollowUpSteps.length > 0,
          scheduled,
        ),
      ),
    )
    .build();
}

export function _onToggleSort(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const props = PropertiesService.getUserProperties();
  const current = props.getProperty("sort_order") !== "desc";
  props.setProperty("sort_order", current ? "desc" : "asc");

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildHomepage(e)))
    .build();
}

export function _openSettingsCard(
  _e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildSettingsCard()))
    .build();
}

export function _onBack(_e: ActionEvent): GoogleAppsScript.Card_Service.ActionResponse {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popCard())
    .build();
}

export function _onBackFromDismissed(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildHomepage(e)))
    .setStateChanged(true)
    .build();
}

function parseDismissParams(
  parameters: Record<string, string> | undefined,
): DismissedEmail | null {
  if (!parameters) return null;
  const threadId = parameters.threadId?.trim();
  if (!threadId) return null;
  const subject = parameters.subject ?? "";
  const to = parameters.to ?? "";
  const dateMs = Number(parameters.dateMs);
  if (!Number.isFinite(dateMs)) return null;
  return { threadId, subject, to, dateMs };
}

export function _onDismissEmail(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const entry = parseDismissParams(e.parameters);
  if (entry) {
    try {
      dismissEmailEntry(entry);
    } catch (err) {
      console.log(
        `[sequel] dismiss failed: ${err instanceof Error ? err.message : "unknown"}`,
      );
    }
  }

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildHomepage(e)))
    .setNotification(
      CardService.newNotification().setText(
        entry ? "Email dismissed" : "Could not dismiss email",
      ),
    )
    .build();
}

export function _openDismissedCard(
  _e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildDismissedCard()))
    .build();
}

export function _onRestoreEmail(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const threadId = e.parameters?.threadId?.trim();
  if (threadId) {
    try {
      restoreDismissedThread(threadId);
    } catch (err) {
      console.log(
        `[sequel] restore failed: ${err instanceof Error ? err.message : "unknown"}`,
      );
    }
  }

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildDismissedCard()))
    .setNotification(
      CardService.newNotification().setText(
        threadId ? "Email restored" : "Could not restore email",
      ),
    )
    .build();
}

export function _onClearDismissed(
  _e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  try {
    clearDismissedList();
  } catch (err) {
    console.log(
      `[sequel] clear dismissed failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
  }

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildDismissedCard()))
    .setNotification(
      CardService.newNotification().setText("All dismissed emails cleared"),
    )
    .build();
}

export function _onToggleAutoFollowUp(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const threadId = e.parameters?.threadId?.trim();
  const enable = e.parameters?.enable === "true";
  let ok = false;
  if (threadId) {
    try {
      setThreadAutoFollowUp(threadId, enable);
      ok = true;
    } catch (err) {
      console.log(
        `[sequel] auto follow-up toggle failed: ${err instanceof Error ? err.message : "unknown"}`,
      );
    }
  }

  const message = !ok
    ? "Could not update auto follow-up"
    : enable
      ? "Auto follow-up on for this thread"
      : "Auto follow-up off for this thread";

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildHomepage(e)))
    .setNotification(CardService.newNotification().setText(message))
    .build();
}

/** Time-driven trigger handler installed by `syncAutoFollowUpTrigger`. */
export function _runAutoFollowUps(): void {
  try {
    const summary = runAutoFollowUps();
    if (summary.skipped)
      console.log(`[sequel] auto follow-up skipped: ${summary.skipped}`);
  } catch (err) {
    console.log(
      `[sequel] auto follow-up run failed: ${err instanceof Error ? err.message : "unknown"}`,
    );
  }
}
