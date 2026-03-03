const PROPS = PropertiesService.getUserProperties();

const DEFAULT_DAYS = 3;
const DEFAULT_AUTO_LABEL = "";
const DEFAULT_EXCLUDED_DOMAINS = "";
const DEFAULT_STALE_DAYS = 90;

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

function _buildAddOn(e: ActionEvent): GoogleAppsScript.Card_Service.Card {
  return buildHomepage(e);
}

function buildSettingsCard(): GoogleAppsScript.Card_Service.Card {
  const days = parseInt(
    PROPS.getProperty("followup_days") || String(DEFAULT_DAYS),
  );
  const excludeReplied = PROPS.getProperty("exclude_replied") !== "false";
  const autoLabel = PROPS.getProperty("auto_label") || DEFAULT_AUTO_LABEL;
  const excludedDomains = PROPS.getProperty("excluded_domains") ||
    DEFAULT_EXCLUDED_DOMAINS;
  const staleDays = parseInt(
    PROPS.getProperty("stale_days") || String(DEFAULT_STALE_DAYS),
  );

  const card = CardService.newCardBuilder().setHeader(
    CardService.newCardHeader().setTitle("Settings"),
  );

  const saveButton = CardService.newTextButton()
    .setText("Save & Refresh")
    .setOnClickAction(
      CardService.newAction().setFunctionName("_onSaveSettings"),
    )
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED);

  const backButton = CardService.newTextButton().setText("Go Back")
    .setOnClickAction(CardService.newAction().setFunctionName("_onBack"))
    .setTextButtonStyle(CardService.TextButtonStyle.TEXT);

  card.addSection(
    buildSettingsSection(
      days,
      excludeReplied,
      autoLabel,
      excludedDomains,
      staleDays,
    ),
  );

  card.setFixedFooter(
    CardService.newFixedFooter().setPrimaryButton(saveButton)
      .setSecondaryButton(backButton),
  );

  return card.build();
}

function buildHomepage(_e: ActionEvent): GoogleAppsScript.Card_Service.Card {
  const days = parseInt(
    PROPS.getProperty("followup_days") || String(DEFAULT_DAYS),
  );
  const excludeReplied = PROPS.getProperty("exclude_replied") !== "false";
  const autoLabel = PROPS.getProperty("auto_label") || DEFAULT_AUTO_LABEL;
  const excludedDomains = PROPS.getProperty("excluded_domains") ||
    DEFAULT_EXCLUDED_DOMAINS;
  const staleDays = parseInt(
    PROPS.getProperty("stale_days") || String(DEFAULT_STALE_DAYS),
  );
  const emails = getPendingFollowUps(
    days,
    excludeReplied,
    excludedDomains,
    autoLabel,
    staleDays,
  );
  const sortAsc = PROPS.getProperty("sort_order") !== "desc";

  const card = CardService.newCardBuilder().addCardAction(
    CardService.newCardAction().setText("Settings").setOnClickAction(
      CardService.newAction().setFunctionName("_openSettingsCard"),
    ),
  );

  // Results section
  card.addSection(buildResultsSection(emails, days, sortAsc));

  return card.build();
}

// ── Card Sections ─────────────────────────────────────────────

function buildSettingsSection(
  days: number,
  excludeReplied: boolean,
  autoLabel: string,
  excludedDomains: string,
  staleDays: number,
): GoogleAppsScript.Card_Service.CardSection {
  const section = CardService.newCardSection();

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
          CardService.newAction().setFunctionName("_onToggleExcludeReplied"),
        ),
    );

  const autoLabelInput = CardService.newTextInput()
    .setFieldName("auto_label")
    .setTitle("Auto-label (optional)")
    .setValue(autoLabel)
    .setHint("Label name to apply to follow-up threads");

  const excludedDomainsInput = CardService.newTextInput()
    .setFieldName("excluded_domains")
    .setTitle("Excluded domains (optional)")
    .setValue(excludedDomains)
    .setHint("Comma-separated, e.g. company.com,slack.com");

  const staleDaysInput = CardService.newTextInput()
    .setFieldName("stale_days")
    .setTitle("Ignore emails older than (days)")
    .setValue(staleDays.toString())
    .setHint("E.g. 90 hides anything sent 90+ days ago");

  section
    .addWidget(daysInput)
    .addWidget(staleDaysInput)
    .addWidget(excludeToggle)
    .addWidget(autoLabelInput)
    .addWidget(excludedDomainsInput);

  return section;
}

function buildResultsSection(
  emails: PendingEmail[],
  days: number,
  sortAsc: boolean,
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

    return section;
  }

  const sortButton = CardService.newTextButton()
    .setText(sortAsc ? "Oldest first" : "Newest first")
    // @ts-ignore This does not show up in the type definitions but it is a valid method
    .setMaterialIcon(
      CardService.newMaterialIcon().setName(
        sortAsc ? "arrow_upward" : "arrow_downward",
      ),
    )
    .setOnClickAction(
      CardService.newAction().setFunctionName("_onToggleSort"),
    );
  section.addWidget(CardService.newButtonSet().addButton(sortButton));

  const sorted = sortAsc ? emails : [...emails].reverse();

  sorted.slice(0, 15).forEach((email) => {
    const age = getDayAge(email.date);
    const label = age === 1 ? "1 day ago" : `${age} days ago`;

    const openLink = CardService.newOpenLink()
      .setUrl(`https://mail.google.com/mail/#inbox/${email.threadId}`)
      .setOpenAs(CardService.OpenAs.FULL_SIZE);

    const widget = CardService.newDecoratedText()
      .setTopLabel(truncate(email.to, 64))
      .setText(truncate(email.subject, 28))
      .setBottomLabel(label)
      .setWrapText(true)
      .setButton(
        CardService.newTextButton()
          .setText("Open")
          .setOpenLink(openLink),
      )
      .setOpenLink(openLink);

    section.addWidget(widget);
    section.addWidget(CardService.newDivider());
  });

  if (emails.length > 15) {
    section.addWidget(
      CardService.newTextParagraph().setText(
        `… and ${emails.length - 15} more`,
      ),
    );
  }

  return section;
}

// ── Action Handlers ──────────────────────────────────────────

function _onSaveSettings(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const days = parseInt(e.formInput["followup_days"]) || DEFAULT_DAYS;
  const clamped = Math.max(1, Math.min(365, days));
  PROPS.setProperty("followup_days", clamped.toString());

  PROPS.setProperty(
    "excluded_domains",
    (e.formInput["excluded_domains"] || "").trim(),
  );
  PROPS.setProperty("auto_label", (e.formInput["auto_label"] || "").trim());

  const staleDays = parseInt(e.formInput["stale_days"]) || DEFAULT_STALE_DAYS;
  PROPS.setProperty("stale_days", Math.max(1, staleDays).toString());

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

function _onToggleExcludeReplied(e: ActionEvent): void {
  const val = e.formInput["exclude_replied"] === "true" ? "true" : "false";
  PROPS.setProperty("exclude_replied", val);
}

function _onToggleSort(
  e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const current = PROPS.getProperty("sort_order") !== "desc";
  PROPS.setProperty("sort_order", current ? "desc" : "asc");

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildHomepage(e)))
    .build();
}

function _openSettingsCard(
  _e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const navigation = CardService.newNavigation().pushCard(buildSettingsCard());

  return CardService.newActionResponseBuilder().setNavigation(navigation)
    .build();
}

function _onBack(
  _e: ActionEvent,
): GoogleAppsScript.Card_Service.ActionResponse {
  const navigation = CardService.newNavigation().popCard();

  return CardService.newActionResponseBuilder().setNavigation(navigation)
    .build();
}

// ── Core Logic ───────────────────────────────────────────────

/**
 * Returns sent emails older than `days` days that haven't received a reply.
 * Results are cached for 5 minutes to avoid repeated Gmail API calls.
 */
function getPendingFollowUps(
  days: number,
  excludeReplied: boolean,
  excludedDomains: string = "",
  autoLabel: string = "",
  staleDays: number = DEFAULT_STALE_DAYS,
): PendingEmail[] {
  const cache = CacheService.getUserCache();
  const cacheKey =
    `fu_${days}_${excludeReplied}_${staleDays}_${excludedDomains}_${autoLabel}`;

  const hit = cache.get(cacheKey);

  if (hit) {
    const parsed: Array<Omit<PendingEmail, "date"> & { dateMs: number }> = JSON
      .parse(hit);
    return parsed.map((e) => ({ ...e, date: new Date(e.dateMs) }));
  }

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);

  const staleCutoff = new Date();
  staleCutoff.setDate(staleCutoff.getDate() - staleDays);

  // Search sent mail within the active window (not too recent, not too old)
  const query = `in:sent after:${formatDateForQuery(staleCutoff)} before:${
    formatDateForQuery(cutoff)
  }`;
  const threads = GmailApp.search(query, 0, 100);

  const myEmail = getMyEmail();

  // Pre-process domain list once outside the loop
  const domainList = excludedDomains
    .split(",")
    .map((d) => d.trim().toLowerCase())
    .filter(Boolean);

  const results: PendingEmail[] = [];
  const threadMap = new Map<string, GoogleAppsScript.Gmail.GmailThread>();

  threads.forEach((thread) => {
    const messages = thread.getMessages();

    let lastSent: GoogleAppsScript.Gmail.GmailMessage | null = null;
    let hasReply = false;

    for (const m of messages) {
      if (m.getFrom().includes(myEmail)) {
        lastSent = m; // keep updating — last one wins
      } else {
        hasReply = true;
      }
    }

    if (!lastSent) return;
    if (excludeReplied && hasReply) return;

    const sentDate = lastSent.getDate();
    if (sentDate > cutoff || sentDate <= staleCutoff) return;

    const to = lastSent.getTo().split(",")[0].trim();

    if (domainList.length > 0) {
      const toDomain = getEmailDomain(to);
      if (
        domainList.some((d) => toDomain === d || toDomain.endsWith(`.${d}`))
      ) {
        return;
      }
    }

    const threadId = thread.getId();
    results.push({
      threadId,
      subject: thread.getFirstMessageSubject() || "(no subject)",
      to,
      date: sentDate,
    });

    if (autoLabel) threadMap.set(threadId, thread);
  });

  results.sort((a, b) => a.date.getTime() - b.date.getTime());

  if (autoLabel) {
    let label = GmailApp.getUserLabelByName(autoLabel);
    if (!label) label = GmailApp.createLabel(autoLabel);

    results.forEach((r) => {
      const thread = threadMap.get(r.threadId);
      if (thread) thread.addLabel(label!);
    });
  }

  try {
    const toCache = results.map((r) => ({ ...r, dateMs: r.date.getTime() }));
    cache.put(cacheKey, JSON.stringify(toCache), 300);
  } catch (_) {
    // Ignore — cache write failures are non-fatal
  }

  return results;
}

// ── Utilities ────────────────────────────────────────────────

function getMyEmail(): string {
  const cache = CacheService.getUserCache();
  const cached = cache.get("my_email");

  if (cached) return cached;

  const email = Session.getActiveUser().getEmail();
  cache.put("my_email", email, 3_600);

  return email;
}

function getEmailDomain(emailStr: string): string {
  const match = emailStr.match(/@([^>@\s]+)/);
  return match ? match[1].toLowerCase() : "";
}

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
