# Sequel for Gmail

Sequel is a Gmail add-on that surfaces **sent mail that may need a follow-up**: threads you sent at least *N* days ago that still look unanswered, with optional filters for replies, domain exclusions, internal mail, staleness, and auto-labeling. It runs on **Google Apps Script (V8)** and is deployed with [clasp](https://github.com/google/clasp).

**Status:** tester-only — not published on the Google Workspace Marketplace for general use.

**License:** [AGPL-3.0-only](LICENSE). If you run a modified version for users, AGPL section 13 requires offering them the corresponding source; the add-on homepage includes a **Source code** link to this repository for that purpose.

**Repository:** https://github.com/vallode/sequel-for-gmail

## What you see

Open the Sequel icon in Gmail's side panel to get the homepage card:

- **Follow-up list:** up to 15 sent threads that are at least *N* days old (default 3) and still waiting on a reply. Each row shows the recipient, subject and age, opens the thread on click, and has a **×** button to dismiss it.
- **Sort toggle:** switches between oldest first and newest first.
- **Truncation notice:** appears when there are more rows than fit on the card ("… and N more"), or when the Gmail search hit its scan cap.
- **Card actions:** **Dismissed** (review and restore dismissed threads), **Settings** (thresholds, filters, auto-label; **Save & Refresh** applies them right away) and **Source code**.

## Prerequisites

- **Node.js:** current Active Node LTS (Node 24 is expected in late 2026).
- **pnpm 10** via [Corepack](https://nodejs.org/api/corepack.html) (`corepack enable`).
- **clasp:** installed as a project dependency; you must be logged in (`clasp login`) with access to the Apps Script project referenced in `.clasp.json`.

Do not commit clasp credential files (for example `.clasprc.json`); see [CONTRIBUTING.md](CONTRIBUTING.md).

## Setup

```bash
git clone https://github.com/vallode/sequel-for-gmail.git
cd sequel-for-gmail
pnpm install
```

Configure clasp if you have not already:

```bash
pnpm exec clasp login
```

The script project ID and `rootDir` (`build/`) are already set in `.clasp.json`.

## Commands

| Command | Purpose |
|--------|---------|
| `pnpm install` | Install dependencies |
| `pnpm check` | Format check, lint, typecheck, unit tests, and clean build |
| `pnpm test` | Unit tests |
| `pnpm lint` | Lint |
| `pnpm format` | Format sources |
| `pnpm typecheck` | TypeScript typecheck |
| `pnpm clean` | Delete `build/` |
| `pnpm build` | esbuild bundles TypeScript into one global-scope `build/Code.js` and copies `appsscript.json` into `build/` |
| `pnpm push` | Build and push `build/` to the linked Apps Script project (**overwrites** remote project files) |
| `pnpm pull` | Pull the latest script from Apps Script |
| `pnpm open` | Open the script in the browser |
| `pnpm deploy` | Build, push, and **create a new Apps Script deployment** (see [Deployment](#deployment)) |

## Architecture

```
TypeScript (src/) → esbuild → build/Code.js + build/appsscript.json → clasp push → Google Apps Script
```

- **Runtime:** All add-on code executes inside Google Apps Script. There is no Node runtime in production. Exported functions must remain **globally scoped** so Gmail Card Service can invoke entry points and `_on*` action handlers.
- **UI:** [Card Service](https://developers.google.com/apps-script/add-ons/gmail/extending-gmail-ui) cards built in TypeScript; homepage, settings, and dismissed views are separate card flows.
- **Entry points:** Declared in `appsscript.json` (for example Gmail homepage trigger → `_buildHomepage`, settings card builder, dismissed card builder).
- **Core logic:** `getPendingFollowUps()` searches Gmail for sent threads in a date window, filters by settings (replies, domains, internal recipients, stale cutoff, dismiss list), optionally applies a user label, sorts results, and returns pending items. Search work is bounded to stay within Apps Script time limits; the UI indicates when the list is truncated.
- **State:** Per-user settings and dismissed reminders live in `PropertiesService.getUserProperties()`. Dismissed history is capped and pruned so storage stays bounded.
- **Cache:** Pending follow-up results are cached per user for about **30 minutes** (`CacheService.getUserCache()`), so dismissing, restoring, and re-sorting don't rescan Gmail. The homepage shows when results were last fetched; **Refresh** on the homepage and **Save & Refresh** on settings bypass the cache.

## Data stored

All persisted data is **per Google account** in that user’s Apps Script user properties (Google’s infrastructure). Sequel does not operate its own database.

Typical keys include:

- Follow-up threshold, stale cutoff, exclude-replied / exclude-internal toggles
- Optional excluded domains and auto-label name
- Sort order preference
- Dismissed reminders (thread id, subject, recipient snippet, sent date metadata), capped and pruned

Cached search results (thread metadata only, TTL ~30 minutes) live in user cache, not long-term storage.

## Privacy and permissions

Sequel **reads** your Gmail (sent threads and thread metadata needed for the list) and can **create or apply labels** when you configure an auto-label. It **does not send email**.

OAuth is limited to:

- Running the Gmail add-on
- Modifying Gmail (read access plus labels)
- Reading the active user’s email address (for filtering and display)

Mail content is processed inside Apps Script to build the follow-up list; maintainers should treat Stackdriver logs as potentially sensitive in error scenarios.

## Testing

There are **no automated end-to-end tests against Gmail**. After changes:

1. Run `pnpm check` locally.
2. Run the manual Gmail smoke checklist in [CONTRIBUTING.md](CONTRIBUTING.md).

## Deployment

**Target:** The Apps Script project ID in `.clasp.json` (already committed; not a secret, but treat deployments as production-affecting).

| Action | Effect |
|--------|--------|
| `pnpm push` | Updates the **head** project files users’ deployments may point at |
| `pnpm deploy` | Creates a **new deployment** version; testers on that deployment ID get the new code |

**Rollback:** If a deployment misbehaves:

1. List deployments: `pnpm exec clasp deployments`
2. List saved project versions: `pnpm exec clasp versions`
3. Point the live deployment at an earlier version: `pnpm exec clasp redeploy <deploymentId> --versionNumber <versionNumber>`

You can also inspect versions and deployments in the Apps Script web editor. Pushing bad code without redeploying may still affect `@HEAD` testing; prefer redeploying a known-good version for stable tester URLs.

## Known limits

- **Tester-only** distribution; marketplace listing and OAuth verification are out of scope for this README.
- **Gmail search** is paginated internally but **capped** to protect execution time; very large mailboxes may not scan every matching thread in one run. The card shows when results are truncated (including a cap on visible rows).
- **Apps Script quotas:** each homepage load runs inside a Card Service callback (30-second limit) and counts against the user's daily Gmail read quota. Sequel scans at most **300 threads** per load (`MAX_THREADS_TO_SCAN`, in pages of 100) and caches results to limit repeat reads. User properties are limited to 9 KB per value and 500 KB in total, which is why the dismissed list is capped at **100 entries**. Cache values are capped at 100 KB, and larger result sets are simply not cached.
- **Cache:** Homepage loads may show data up to ~30 minutes old (the "Updated X min ago" row shows how old) unless you use **Refresh** or **Save & Refresh**.
- **Dismissed list** retains a bounded history; oldest entries are dropped when the cap is exceeded.
- **Auto-label** creates or uses a Gmail label and applies it to matching threads when results are computed.
- **No send:** Sequel cannot compose or send mail on your behalf.

## Branding

Add-on logo asset: `assets/logo.png`. Public URL after merge to `main`:

`https://raw.githubusercontent.com/vallode/sequel-for-gmail/main/assets/logo.png`

## Contributing and security

- [CONTRIBUTING.md](CONTRIBUTING.md) — workflow and manual test checklist  
- [SECURITY.md](SECURITY.md) — vulnerability reporting and data handling  
