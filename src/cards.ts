import { getAutoFollowUpThreadIds } from "./auto-follow-up-runner.ts";
import {
  HOMEPAGE_DISPLAY_LIMIT,
  MAX_AUTO_FOLLOWUP_STEPS,
  SOURCE_REPO_URL,
} from "./constants.ts";
import { getDayAge } from "./dates.ts";
import { dismissedThreadIds } from "./dismissed-store.ts";
import { applySortOrder } from "./follow-ups-core.ts";
import { getPendingFollowUps } from "./gmail-follow-ups.ts";
import { getMyDomain } from "./runtime/session-email.ts";
import { readDismissedList, readUserSettings } from "./runtime/user-props.ts";
import { autoFollowUpFieldNames, defaultAutoFollowUpSteps } from "./settings.ts";
import type {
  ActionEvent,
  DismissedEmail,
  PendingEmail,
  UserSettings,
} from "./types.ts";
import { truncate } from "./util.ts";

function buildSettingsSection(
  settings: UserSettings,
): GoogleAppsScript.Card_Service.CardSection {
  const section = CardService.newCardSection();

  const daysInput = CardService.newTextInput()
    .setFieldName("followup_days")
    .setTitle("Follow-up after (days)")
    .setValue(String(settings.days))
    .setHint("E.g. 3 means emails sent 3+ days ago");

  const excludeToggle = CardService.newDecoratedText()
    .setText("Exclude if they replied")
    .setBottomLabel("Skip emails where you already got a reply")
    .setSwitchControl(
      CardService.newSwitch()
        .setFieldName("exclude_replied")
        .setValue("true")
        .setSelected(settings.excludeReplied),
    );

  const excludeInternalToggle = CardService.newDecoratedText()
    .setText("Exclude self-sent & internal")
    .setBottomLabel(
      `Skip emails sent to yourself or your own domain (${getMyDomain()})`,
    )
    .setSwitchControl(
      CardService.newSwitch()
        .setFieldName("exclude_internal")
        .setValue("true")
        .setSelected(settings.excludeInternal),
    );

  const excludedDomainsInput = CardService.newTextInput()
    .setFieldName("excluded_domains")
    .setTitle("Excluded domains (optional)")
    .setValue(settings.excludedDomains)
    .setHint("Comma-separated, e.g. company.com,slack.com");

  const staleDaysInput = CardService.newTextInput()
    .setFieldName("stale_days")
    .setTitle("Ignore emails older than (days)")
    .setValue(String(settings.staleDays))
    .setHint("E.g. 90 hides anything sent 90+ days ago");

  section
    .addWidget(daysInput)
    .addWidget(staleDaysInput)
    .addWidget(CardService.newDivider())
    .addWidget(excludeToggle)
    .addWidget(excludeInternalToggle)
    .addWidget(CardService.newDivider())
    .addWidget(excludedDomainsInput);

  return section;
}

function buildAutoFollowUpSettingsSection(
  settings: UserSettings,
): GoogleAppsScript.Card_Service.CardSection {
  const section = CardService.newCardSection()
    .setHeader("Auto follow-up drafts")
    .addWidget(
      CardService.newDecoratedText()
        .setText("Create follow-up drafts")
        .setBottomLabel("Hourly, on Follow-ups/Auto threads. Never sends.")
        .setWrapText(true)
        .setSwitchControl(
          CardService.newSwitch()
            .setFieldName("auto_followup_enabled")
            .setValue("true")
            .setSelected(settings.autoFollowUpEnabled),
        ),
    );

  const fallback = defaultAutoFollowUpSteps();
  for (let i = 0; i < MAX_AUTO_FOLLOWUP_STEPS; i++) {
    const step = settings.autoFollowUpSteps[i];
    const fields = autoFollowUpFieldNames(i);
    section
      .addWidget(CardService.newDivider())
      .addWidget(
        CardService.newTextInput()
          .setFieldName(fields.waitDays)
          .setTitle(`Step ${i + 1}: wait (days)`)
          .setValue(String(step?.waitDays ?? fallback[i]?.waitDays ?? settings.days))
          .setHint("Days after your last email"),
      )
      .addWidget(
        CardService.newTextInput()
          .setFieldName(fields.template)
          .setTitle(`Step ${i + 1}: message`)
          .setMultiline(true)
          .setValue(step?.template ?? "")
          .setHint("Use {{firstName}} and {{subject}}"),
      );
  }

  return section;
}

export function buildSettingsCard(): GoogleAppsScript.Card_Service.Card {
  const settings = readUserSettings();

  const card = CardService.newCardBuilder().setHeader(
    CardService.newCardHeader().setTitle("Settings"),
  );

  const saveButton = CardService.newTextButton()
    .setText("Save & Refresh")
    .setOnClickAction(CardService.newAction().setFunctionName("_onSaveSettings"))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  const backButton = CardService.newTextButton()
    .setText("Go Back")
    .setOnClickAction(CardService.newAction().setFunctionName("_onBack"))
    .setTextButtonStyle(CardService.TextButtonStyle.OUTLINED);

  card.addSection(buildSettingsSection(settings));
  card.addSection(buildAutoFollowUpSettingsSection(settings));

  card.setFixedFooter(
    CardService.newFixedFooter()
      .setPrimaryButton(saveButton)
      .setSecondaryButton(backButton),
  );

  return card.build();
}

export function buildResultsSection(
  emails: PendingEmail[],
  days: number,
  sortAsc: boolean,
  searchTruncated: boolean,
  autoFollowUpThreadIds: Set<string> | null = null,
): GoogleAppsScript.Card_Service.CardSection {
  const section = CardService.newCardSection();

  if (emails.length === 0) {
    section.addWidget(
      CardService.newTextParagraph().setText(
        `You're all caught up!<br>No sent emails older than ${days} day${
          days !== 1 ? "s" : ""
        } are waiting for a reply`,
      ),
    );
    if (searchTruncated) {
      section.addWidget(
        CardService.newTextParagraph().setText(
          "Some matching threads were not scanned due to volume limits. Narrow your date window in Settings if results look incomplete.",
        ),
      );
    }
    return section;
  }

  section.addWidget(
    CardService.newDecoratedText().setText(
      `${emails.length} thread${emails.length !== 1 ? "s" : ""}`,
    ),
  );

  const sortButton = CardService.newTextButton()
    .setText(sortAsc ? "Oldest first" : "Newest first")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED_TONAL)
    .setMaterialIcon(
      CardService.newMaterialIcon().setName(
        sortAsc ? "arrow_upward" : "arrow_downward",
      ),
    )
    .setOnClickAction(CardService.newAction().setFunctionName("_onToggleSort"));

  section.addWidget(CardService.newButtonSet().addButton(sortButton));

  const sorted = applySortOrder(emails, sortAsc);
  const visible = sorted.slice(0, HOMEPAGE_DISPLAY_LIMIT);

  visible.forEach((email, index) => {
    const age = getDayAge(email.date);
    const label = age === 0 ? "Today" : age === 1 ? "1 day ago" : `${age} days ago`;

    const openLink = CardService.newOpenLink()
      .setUrl(`https://mail.google.com/mail/#all/${email.threadId}`)
      .setOpenAs(CardService.OpenAs.FULL_SIZE);

    const flagged = autoFollowUpThreadIds?.has(email.threadId) ?? false;

    const widget = CardService.newDecoratedText()
      .setTopLabel(truncate(email.to, 64))
      .setText(truncate(email.subject, 24))
      .setBottomLabel(flagged ? `${label} · Auto follow-up on` : label)
      .setWrapText(true)
      .setButton(
        CardService.newTextButton()
          .setAltText("Dismiss this reminder")
          .setMaterialIcon(CardService.newMaterialIcon().setName("close"))
          .setOnClickAction(
            CardService.newAction()
              .setFunctionName("_onDismissEmail")
              .setParameters({
                threadId: email.threadId,
                subject: email.subject,
                to: email.to,
                dateMs: String(email.date.getTime()),
              }),
          )
          .setTextButtonStyle(CardService.TextButtonStyle.BORDERLESS),
      )
      .setOpenLink(openLink);

    section.addWidget(widget);

    if (autoFollowUpThreadIds) {
      section.addWidget(
        CardService.newButtonSet().addButton(
          CardService.newTextButton()
            .setText(flagged ? "Stop auto follow-up" : "Auto follow-up")
            .setMaterialIcon(
              CardService.newMaterialIcon().setName(
                flagged ? "cancel_schedule_send" : "schedule_send",
              ),
            )
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("_onToggleAutoFollowUp")
                .setParameters({
                  threadId: email.threadId,
                  enable: flagged ? "false" : "true",
                }),
            )
            .setTextButtonStyle(CardService.TextButtonStyle.BORDERLESS),
        ),
      );
    }

    if (index < visible.length - 1) {
      section.addWidget(CardService.newDivider());
    }
  });

  if (emails.length > HOMEPAGE_DISPLAY_LIMIT) {
    section.addWidget(
      CardService.newTextParagraph().setText(
        `… and ${emails.length - HOMEPAGE_DISPLAY_LIMIT} more`,
      ),
    );
  }

  if (searchTruncated) {
    section.addWidget(
      CardService.newTextParagraph().setText(
        "Not all sent threads in this window were scanned. Results may be incomplete — try a shorter stale cutoff in Settings.",
      ),
    );
  }

  return section;
}

function buildAutoFollowUpStatusSection(
  settings: UserSettings,
  flaggedCount: number,
): GoogleAppsScript.Card_Service.CardSection {
  const steps = settings.autoFollowUpSteps.length;
  const text =
    steps === 0
      ? "Auto follow-up is on, but no message templates are set. Add one in Settings."
      : `Auto follow-up drafts on · ${flaggedCount} flagged thread${
          flaggedCount !== 1 ? "s" : ""
        } · ${steps}-step sequence`;
  return CardService.newCardSection().addWidget(
    CardService.newDecoratedText()
      .setText(text)
      .setWrapText(true)
      .setStartIcon(
        CardService.newIconImage().setMaterialIcon(
          CardService.newMaterialIcon().setName("schedule_send"),
        ),
      ),
  );
}

export function buildHomepageCard(
  _e: ActionEvent,
  ignoreCache = false,
): GoogleAppsScript.Card_Service.Card {
  const t0 = Date.now();
  const settings = readUserSettings();
  const { emails, searchTruncated } = getPendingFollowUps(settings, ignoreCache);
  const sortAsc =
    PropertiesService.getUserProperties().getProperty("sort_order") !== "desc";

  const dismissed = dismissedThreadIds(readDismissedList());
  const visibleEmails = emails.filter((e) => !dismissed.has(e.threadId));

  const card = CardService.newCardBuilder()
    .addCardAction(
      CardService.newCardAction()
        .setText("Dismissed")
        .setOnClickAction(
          CardService.newAction().setFunctionName("_openDismissedCard"),
        ),
    )
    .addCardAction(
      CardService.newCardAction()
        .setText("Settings")
        .setOnClickAction(CardService.newAction().setFunctionName("_openSettingsCard")),
    )
    .addCardAction(
      CardService.newCardAction()
        .setText("Source code")
        .setOpenLink(
          CardService.newOpenLink()
            .setUrl(SOURCE_REPO_URL)
            .setOpenAs(CardService.OpenAs.FULL_SIZE),
        ),
    );

  let autoFollowUpThreadIds: Set<string> | null = null;
  if (settings.autoFollowUpEnabled) {
    try {
      autoFollowUpThreadIds = getAutoFollowUpThreadIds();
    } catch (err) {
      console.log(
        `[sequel] auto follow-up lookup failed: ${err instanceof Error ? err.message : "unknown"}`,
      );
      autoFollowUpThreadIds = new Set();
    }
    card.addSection(
      buildAutoFollowUpStatusSection(settings, autoFollowUpThreadIds.size),
    );
  }

  card.addSection(
    buildResultsSection(
      visibleEmails,
      settings.days,
      sortAsc,
      searchTruncated,
      autoFollowUpThreadIds,
    ),
  );

  console.log(`[sequel] _buildHomepage: ${Date.now() - t0}ms`);
  return card.build();
}

export function buildDismissedCard(): GoogleAppsScript.Card_Service.Card {
  const dismissed: DismissedEmail[] = readDismissedList();
  const card = CardService.newCardBuilder().setHeader(
    CardService.newCardHeader().setTitle("Dismissed Emails"),
  );

  const section = CardService.newCardSection();

  if (dismissed.length === 0) {
    section.addWidget(CardService.newTextParagraph().setText("No dismissed emails."));
  } else {
    const reversed = [...dismissed].reverse();
    reversed.forEach((entry, index) => {
      const age = Math.floor((Date.now() - entry.dateMs) / 86400000);
      const ageLabel = age === 1 ? "1 day ago" : `${age} days ago`;

      const widget = CardService.newDecoratedText()
        .setTopLabel(truncate(entry.to, 64))
        .setText(truncate(entry.subject, 24))
        .setBottomLabel(ageLabel)
        .setWrapText(true)
        .setButton(
          CardService.newTextButton()
            .setText("Restore")
            .setAltText("Restore this email")
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName("_onRestoreEmail")
                .setParameters({ threadId: entry.threadId }),
            )
            .setTextButtonStyle(CardService.TextButtonStyle.BORDERLESS),
        );

      section.addWidget(widget);
      if (index < reversed.length - 1) {
        section.addWidget(CardService.newDivider());
      }
    });
  }

  card.addSection(section);

  const restoreAllButton = CardService.newTextButton()
    .setText("Restore All")
    .setOnClickAction(CardService.newAction().setFunctionName("_onClearDismissed"))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  const backButton = CardService.newTextButton()
    .setText("Go Back")
    .setOnClickAction(CardService.newAction().setFunctionName("_onBackFromDismissed"))
    .setTextButtonStyle(CardService.TextButtonStyle.OUTLINED);

  card.setFixedFooter(
    CardService.newFixedFooter()
      .setPrimaryButton(restoreAllButton)
      .setSecondaryButton(backButton),
  );

  return card.build();
}
