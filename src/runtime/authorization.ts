/**
 * Prompts the user to re-consent when the manifest gained scopes they have not
 * granted yet. Without this, calls needing a new scope fail inside our own
 * try/catch and Gmail never shows its authorization card.
 */
export function requireFullAuthorization(): void {
  const info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  if (info.getAuthorizationStatus() !== ScriptApp.AuthorizationStatus.REQUIRED) return;
  CardService.newAuthorizationException()
    .setAuthorizationUrl(info.getAuthorizationUrl())
    .setResourceDisplayName("Follow-Up Reminders")
    .throwException();
}
