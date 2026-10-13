/// <reference path="./card-service-extensions.d.ts" />
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

export function _onSaveSettings(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const settings = parseSettingsFromFormInput(e.formInput ?? {});
  writeUserSettings(settings);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildHomepage(e, true)))
    .setNotification(CardService.newNotification().setText("Settings saved"))
    .build();
}

export function _onRefresh(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(_buildHomepage(e, true)))
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
