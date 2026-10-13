# Auto follow-up drafts

Status: implemented, not yet deployed. Pure logic in `src/auto-follow-up-core.ts`
(tests in `tests/auto-follow-up.test.ts`), Gmail side effects in
`src/auto-follow-up-runner.ts`, settings in `src/settings.ts`, UI in `src/cards.ts`.
The phase 0 checks below are still open and must pass before release.

## Behaviour

1. The user flags a thread by applying the Gmail label `Follow-ups/Auto`
   (manually, or with an "Auto follow-up" button on a homepage row).
2. An hourly time-driven trigger scans flagged threads.
3. For each thread, when the next step in the user's sequence is due, Sequel
   creates a **draft** reply in the thread from that step's template. The user
   reviews and sends it from Gmail. Nothing is sent automatically.
4. When the recipient replies, or the sequence is exhausted, the thread moves
   from `Follow-ups/Auto` to `Follow-ups/Done`.

Labels are created automatically, parent first (`Follow-ups`, then `Waiting`,
`Auto`, `Done`), so Gmail nests them. They aren't configurable.

## Sequence

Up to 3 steps, each `{ waitDays, template }`. `waitDays` counts from the
previous message the user sent (the original or the last nudge), so "day 3 then
day 7" is `[3, 4]`. Stored as JSON under the `auto_followup_steps` user property.

Template placeholders: `{{firstName}}` (from the recipient display name,
falling back to "there"), `{{subject}}`, `{{daysSinceSent}}` (days since the
original unanswered message). Unknown placeholders are left as written.

## State: derived from the thread, not stored

`decideAutoFollowUp` needs no per-thread storage:

| Thread state | Decision |
|---|---|
| Any open draft | `wait` (`draft-pending`): don't stack drafts; the user hasn't sent the last one |
| Last message not from user | `stop` (`replied`) |
| N trailing messages from user, N-1 ≥ steps | `stop` (`sequence-complete`) |
| Recipient missing / excluded / internal (when excluded) | `stop` |
| Step N-1 not yet due | `wait` (`not-due`) |
| Otherwise | `draft` step N-1 to the last message's To and Cc |

Manually sent nudges count as steps. A reply followed by a new message from the
user restarts the sequence.

## Runner (`src/auto-follow-up-runner.ts`)

- Global `_runAutoFollowUps`. Export it from `src/entry.ts` and list it in
  `scripts/global-callbacks.json`.
- `LockService.getUserLock()` so overlapping runs can't create duplicate drafts.
- List flagged thread IDs with `Gmail.Users.Threads.list` by label ID (skips
  trashed threads; avoids quoting nested label names in search syntax).
- Bail out if the user's email can't be resolved; otherwise every message looks self-sent.
- Fetch the threads through the Gmail batch endpoint (`fetchAutoFollowUpThreads`,
  `format=metadata`) with From/To/Cc/Subject/Message-ID/References and label IDs:
  `TRASH` messages are dropped, `DRAFT` ones are marked `isDraft`. Threads that
  fail to fetch count as `failed` and are retried on the next hourly run.
- Create the draft with the **Advanced Gmail Service** (`Gmail.Users.Drafts.create`)
  as raw MIME: To/Cc from the decision, `Subject: Re: …`, `In-Reply-To` and
  `References` set to the last message's `Message-ID`, and `threadId`. Avoid
  `GmailMessage.createDraftReply`: replying to your own sent message addresses
  the draft to yourself, and the method needs the full `mail.google.com` scope.
- On `stop`, swap the labels with one `Threads.modify`. Log every decision with `[sequel]` prefixes.
- Cap drafts per run (e.g. 20) to stay inside Apps Script execution time.

## Trigger lifecycle

On settings save: when the feature is enabled and no `_runAutoFollowUps` trigger
exists, create `ScriptApp.newTrigger("_runAutoFollowUps").timeBased().everyHours(1)`.
When it's disabled, delete it. Both run under a lock, and duplicates are
deduplicated by handler name.

## Settings

`autoFollowUpEnabled` (default off) and up to 3 step slots (wait days +
template textarea; an empty template disables the slot). A never-saved sequence
defaults to two steps (3 days, then 4). Excluded
domains and the internal-recipient setting reuse the existing settings.

## Scopes / manifest

- Add `https://www.googleapis.com/auth/script.scriptapp` (for the trigger).
- Enable the Advanced Gmail Service in `appsscript.json`
  (`dependencies.enabledAdvancedServices`).
- Drafts: the Gmail API lists `gmail.modify` as sufficient for
  `users.drafts.create`, so no new Gmail scope should be needed. **Confirm in the spike.**
- Update the OAuth scope notes in CLAUDE.md and README.

## Phases

0. **Test before release (½ day):** throwaway deploy that (a) installs an hourly trigger from
   the add-on and confirms it fires, and (b) creates a threaded draft via
   `Gmail.Users.Drafts.create` with only the current scopes plus `script.scriptapp`.
   Record the consent screen.
1. ~~**Runner + manifest:** `_runAutoFollowUps`, label swap, trigger lifecycle.~~ Done.
2. ~~**Settings + UI:** step editor, per-row "Auto follow-up" button, homepage status line.~~ Done.
3. **Later:** an optional send mode (needs the `gmail.send` scope, a send
   window, daily caps, and a re-check right before sending).

## Known adjacent issue

`threads.get` includes drafts. The existing `evaluateThreadForFollowUp`
doesn't filter them, so an open draft at the end of a thread is treated as a
sent message (wrong date, and a replied thread can look unanswered). This is
worth fixing: the follow-up list fetch already receives label IDs, so it can
skip `DRAFT` messages the same way the runner does.
