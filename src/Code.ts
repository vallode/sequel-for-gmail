const PROPS = PropertiesService.getUserProperties();
const DEFAULT_DAYS = 3;
const DEFAULT_EXCLUDE_REPLIED = true;

interface PendingEmail {
  threadId: string;
  subject: string;
  to: string;
  date: GoogleAppsScript.Base.Date;
}

interface ActionEvent {
  formInput: Record<string, string>;
  parameters: Record<string, string>;
}

// ── Entry Points ─────────────────────────────────────────────

function buildAddOn(e: ActionEvent): GoogleAppsScript.Card_Service.Card {
  return buildHomepage(e);
}

function buildSettingsCard(): GoogleAppsScript.Card_Service.Card {
  return CardService.newCardBuilder().setHeader(
    CardService.newCardHeader().setTitle("Settings"),
  ).build();
}

function buildHomepage(_e: ActionEvent): GoogleAppsScript.Card_Service.Card {
  const days = parseInt(
    PROPS.getProperty("followup_days") || String(DEFAULT_DAYS),
  );
  const excludeReplied = PROPS.getProperty("exclude_replied") !== "false";
  const emails = getPendingFollowUps(days, excludeReplied);

  const card = CardService.newCardBuilder()
    .addCardAction(
      CardService.newCardAction().setText("Gmail").setOpenLink(
        CardService.newOpenLink().setUrl("https://mail.google.com/mail"),
      ),
    );

  // Settings section
  card.addSection(buildSettingsSection(days, excludeReplied));

  // Results section
  card.addSection(buildResultsSection(emails, days));

  return card.build();
}

// ── Card Sections ─────────────────────────────────────────────

function buildSettingsSection(
  days: number,
  excludeReplied: boolean,
): GoogleAppsScript.Card_Service.CardSection {
  const section = CardService.newCardSection()
    .setHeader("⚙️ Settings")
    .setCollapsible(true)
    .setNumUncollapsibleWidgets(0);

  const daysInput = CardService.newTextInput()
    .setFieldName("followup_days")
    .setTitle("Follow-up after (days)")
    .setValue(days.toString())
    .setHint("E.g. 3 means emails sent 3+ days ago");

  const excludeToggle = CardService.newDecoratedText()
    .setText("Exclude if they replied")
    .setBottomLabel("Skip emails where you already got a reply")
    .setSwitchControl(
      CardService.newSwitch()
        .setFieldName("exclude_replied")
        .setValue("true")
        .setSelected(excludeReplied)
        .setOnChangeAction(
          CardService.newAction().setFunctionName("onToggleExcludeReplied"),
        ),
    );

  const saveButton = CardService.newTextButton()
    .setText("Save & Refresh")
    .setOnClickAction(
      CardService.newAction().setFunctionName("onSaveSettings"),
    )
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  const dailyDigestButton = CardService.newTextButton()
    .setText("Enable Daily Digest Email")
    .setOnClickAction(
      CardService.newAction().setFunctionName("onEnableDailyDigest"),
    );

  const disableDigestButton = CardService.newTextButton()
    .setText("Disable Daily Digest")
    .setOnClickAction(
      CardService.newAction().setFunctionName("onDisableDailyDigest"),
    );

  const digestStatus = hasDailyTrigger()
    ? "✅ Daily digest is active (runs at 8 AM)"
    : "⭕ Daily digest is off";

  section
    .addWidget(daysInput)
    .addWidget(excludeToggle)
    .addWidget(
      CardService.newTextParagraph().setText(digestStatus),
    )
    .addWidget(
      CardService.newButtonSet()
        .addButton(dailyDigestButton)
        .addButton(disableDigestButton),
    )
    .addWidget(
      CardService.newButtonSet().addButton(saveButton),
    );

  return section;
}

function buildResultsSection(
  emails: PendingEmail[],
  days: number,
): GoogleAppsScript.Card_Service.CardSection {
  const section = CardService.newCardSection().setHeader(
    `Emails needing follow-up`,
  );

  if (emails.length === 0) {
    section.addWidget(
      CardService.newTextParagraph().setText(
        `🎉 You're all caught up! No sent emails older than ${days} day${
          days !== 1 ? "s" : ""
        } are waiting for a reply.`,
      ),
    );

    return section;
  }

  emails.slice(0, 15).forEach((email) => {
    const age = getDayAge(email.date);
    const label = age === 1 ? "1 day ago" : `${age} days ago`;

    const widget = CardService.newDecoratedText()
      .setText(truncate(email.subject, 45))
      .setBottomLabel(`To: ${email.to} · ${label}`)
      .setWrapText(true)
      .setOnClickAction(
        CardService.newAction()
          .setFunctionName("onOpenThread")
          .setParameters({ threadId: email.threadId }),
      );

    section.addWidget(widget);
  });

  if (emails.length > 15) {
    section.addWidget(
      CardService.newTextParagraph().setText(
        `… and ${
          emails.length - 15
        } more. Check your daily digest for the full list.`,
      ),
    );
  }

  return section;
}

// ── Action Handlers ──────────────────────────────────────────

function onSaveSettings(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const days = parseInt(e.formInput["followup_days"]) || DEFAULT_DAYS;
  const clamped = Math.max(1, Math.min(365, days));
  PROPS.setProperty("followup_days", clamped.toString());
  return CardService.newActionResponseBuilder()
    .setNavigation(
      CardService.newNavigation().updateCard(buildHomepage(e)),
    )
    .setNotification(
      CardService.newNotification().setText(
        `Settings saved: ${clamped}-day follow-up window`,
      ),
    )
    .build();
}

function onToggleExcludeReplied(e: ActionEvent): void {
  const val = e.formInput["exclude_replied"] === "true" ? "true" : "false";
  PROPS.setProperty("exclude_replied", val);
}

function onOpenThread(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const threadId = e.parameters["threadId"];
  const url = `https://mail.google.com/mail/#inbox/${threadId}`;
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink().setUrl(url))
    .build();
}

function onEnableDailyDigest(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  enableDailyDigest();
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildHomepage(e)))
    .setNotification(
      CardService.newNotification().setText(
        "Daily digest enabled — runs at 8 AM",
      ),
    )
    .build();
}

function onDisableDailyDigest(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  disableDailyDigest();
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildHomepage(e)))
    .setNotification(
      CardService.newNotification().setText("Daily digest disabled"),
    )
    .build();
}

// ── Core Logic ───────────────────────────────────────────────

/**
 * Returns sent emails older than `days` days that haven't received a reply.
 */
function getPendingFollowUps(
  days: number,
  excludeReplied: boolean,
): PendingEmail[] {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);

  // Search sent mail older than cutoff
  const query = `in:sent before:${formatDateForQuery(cutoff)}`;
  const threads = GmailApp.search(query, 0, 100);

  const results: PendingEmail[] = [];

  threads.forEach((thread) => {
    if (excludeReplied && threadHasReply(thread)) return;

    const messages = thread.getMessages();
    // Find the last sent message in this thread
    const sentMessages = messages.filter((m) => isSentByMe(m));
    if (sentMessages.length === 0) return;

    const lastSent = sentMessages[sentMessages.length - 1];
    const sentDate = lastSent.getDate();

    // Only include if the last sent message (not a reply to their reply) is old enough
    if (sentDate > cutoff) return;

    results.push({
      threadId: thread.getId(),
      subject: thread.getFirstMessageSubject() || "(no subject)",
      to: lastSent.getTo().split(",")[0].trim(),
      date: sentDate,
    });
  });

  // Sort oldest first
  results.sort((a, b) => a.date.getTime() - b.date.getTime());
  return results;
}

function threadHasReply(thread: GoogleAppsScript.Gmail.GmailThread): boolean {
  const messages = thread.getMessages();
  const myEmail = Session.getActiveUser().getEmail();
  // Check if any message was NOT sent by me (i.e., a reply from recipient)
  return messages.some((m) => {
    const from = m.getFrom();
    return !from.includes(myEmail);
  });
}

function isSentByMe(message: GoogleAppsScript.Gmail.GmailMessage): boolean {
  const myEmail = Session.getActiveUser().getEmail();
  return message.getFrom().includes(myEmail);
}

// ── Daily Digest ─────────────────────────────────────────────

function sendDailyDigest(): void {
  const days = parseInt(
    PROPS.getProperty("followup_days") || String(DEFAULT_DAYS),
  );
  const excludeReplied = PROPS.getProperty("exclude_replied") !== "false";
  const emails = getPendingFollowUps(days, excludeReplied);

  if (emails.length === 0) return; // Don't send if nothing to follow up on

  const userEmail = Session.getActiveUser().getEmail();

  const htmlRows = emails.map((e) => {
    const age = getDayAge(e.date);
    const url = `https://mail.google.com/mail/#all/${e.threadId}`;
    return `
      <tr style="border-bottom:1px solid #eee">
        <td style="padding:10px 8px">
          <a href="${url}" style="color:#1a73e8;text-decoration:none;font-weight:500">${
      escapeHtml(e.subject)
    }</a>
        </td>
        <td style="padding:10px 8px;color:#555;white-space:nowrap">${
      escapeHtml(e.to)
    }</td>
        <td style="padding:10px 8px;color:#d93025;white-space:nowrap;font-weight:500">${age}d ago</td>
      </tr>`;
  }).join("");

  const html = `
    <html><body style="font-family:sans-serif;color:#202124;max-width:600px;margin:0 auto">
      <h2 style="color:#1a73e8">📬 Follow-Up Reminder</h2>
      <p>You have <strong>${emails.length}</strong> email${
    emails.length !== 1 ? "s" : ""
  } sent more than <strong>${days} day${
    days !== 1 ? "s" : ""
  }</strong> ago that may need a follow-up:</p>
      <table style="width:100%;border-collapse:collapse;font-size:14px">
        <thead>
          <tr style="background:#f1f3f4;text-align:left">
            <th style="padding:10px 8px">Subject</th>
            <th style="padding:10px 8px">To</th>
            <th style="padding:10px 8px">Age</th>
          </tr>
        </thead>
        <tbody>${htmlRows}</tbody>
      </table>
      <p style="margin-top:24px;font-size:12px;color:#777">
        Sent by your Gmail Follow-Up Add-on ·
        <a href="https://mail.google.com/mail/#sent" style="color:#777">View Sent Mail</a>
      </p>
    </body></html>`;

  GmailApp.sendEmail(
    userEmail,
    `📬 ${emails.length} follow-up${emails.length !== 1 ? "s" : ""} needed`,
    "",
    {
      htmlBody: html,
      name: "Follow-Up Reminder",
    },
  );
}

// ── Trigger Management ───────────────────────────────────────

function enableDailyDigest(): void {
  disableDailyDigest(); // Remove any existing
  ScriptApp.newTrigger("sendDailyDigest")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

function disableDailyDigest(): void {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "sendDailyDigest")
    .forEach((t) => ScriptApp.deleteTrigger(t));
}

function hasDailyTrigger(): boolean {
  return ScriptApp.getProjectTriggers()
    .some((t) => t.getHandlerFunction() === "sendDailyDigest");
}

// ── Utilities ────────────────────────────────────────────────

function formatDateForQuery(date: Date): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}/${m}/${d}`;
}

function getDayAge(date: GoogleAppsScript.Base.Date): number {
  return Math.floor((Date.now() - date.getTime()) / 86400000);
}

function truncate(str: string, max: number): string {
  return str.length > max ? str.slice(0, max - 1) + "…" : str;
}

function escapeHtml(str: string): string {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
