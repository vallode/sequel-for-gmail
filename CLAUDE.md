# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
pnpm install     # Install dev dependencies (Corepack uses packageManager pin)
pnpm clean       # Delete build/
pnpm build       # Bundle TypeScript → build/Code.js + copy appsscript.json
pnpm typecheck   # tsc --noEmit
pnpm lint        # biome check .
pnpm format      # biome format --write .
pnpm test        # Node built-in test runner (pure logic under tests/)
pnpm check       # lint/format check + typecheck + test + build
pnpm push        # Build + clasp push
pnpm deploy      # Build + push + clasp deploy
pnpm pull        # Pull latest from Google Apps Script
pnpm open        # Open the script in browser
```

Validation combines automated tests for pure logic with manual Gmail add-on testing after deploy.

## Architecture

Sequel is a Gmail add-on built on Google Apps Script (V8 runtime), written in TypeScript and deployed via `clasp`. Source lives under `src/` as ES modules; `scripts/build.mjs` uses esbuild to produce a single global-scope `build/Code.js` because Apps Script cannot load ES modules at runtime.

**Build pipeline:** TypeScript modules → esbuild (IIFE + `globalThis` exports) → `build/Code.js` + `build/appsscript.json` → `clasp push` → Google Apps Script.

**Runtime:** All Card Service entry points and `_on*` handlers must exist as global functions on `globalThis` in the deployed bundle. Interfaces and types are compile-time only.

**UI:** Google Card Service API. Entry functions declared in `src/appsscript.json`: `_buildHomepage`, `buildSettingsCard`, `buildDismissedCard`.

**State:** User settings via `PropertiesService.getUserProperties()`. Read them with one `getProperties()` snapshot (`readUserSettings`, `readHomepageProps`), not per-key `getProperty` calls: each call is a separate round trip (~30–55 ms). Dismissed emails are stored as JSON under `dismissed_ids` (capped, lock-protected updates). Follow-up results cached ~30 minutes in `CacheService` with hashed keys; Save & Refresh and the homepage Refresh action bypass cache.

**Core logic:** `getPendingFollowUps()` lists sent-thread IDs with `Gmail.Users.Threads.list` (bounded), fetches From/To/Subject/date for each through the Gmail batch endpoint (up to 50 `threads.get?format=metadata` calls per multipart request, sent together via `UrlFetchApp.fetchAll`; no bodies; `runtime/gmail-api.ts`), filters threads in `follow-ups-core.ts` (replies, domains, internal, stale window), and syncs the `Follow-ups/Waiting` label.

**Labels:** Fixed, not user-configurable: `Follow-ups` parent with `Waiting`, `Auto`, `Done` children (`FOLLOWUP_LABELS` in `constants.ts`). `ensureFollowUpLabels()` in `runtime/labels.ts` creates missing ones parent-first.

**Auto follow-up:** Hourly time-driven trigger `_runAutoFollowUps` (installed/removed by `syncAutoFollowUpTrigger` on settings save) scans threads labelled `Follow-ups/Auto`. Pure decisions live in `auto-follow-up-core.ts` (state derived from the thread: trailing sent-by-user count, open drafts); Gmail side effects in `auto-follow-up-runner.ts`. Creates drafts only, via the Advanced Gmail Service; never sends. See `docs/auto-follow-up.md`.

**OAuth scopes:** `gmail.addons.execute` (add-on runtime), `gmail.modify` (search threads, apply labels, and create drafts via the Advanced Gmail Service — covers all `GmailApp` usage), `userinfo.email` (`Session.getActiveUser().getEmail()`), `script.scriptapp` (auto follow-up time-driven trigger), `script.external_request` (`UrlFetchApp` calls to the Gmail API; also allowed by `urlFetchWhitelist` in the manifest). `gmail.readonly` is not requested (redundant with modify). Don't use `GmailApp` send/draft methods: they require the full `mail.google.com` scope.

**Action handlers:** Functions prefixed with `_on` or `_open` are Card Service callbacks and must remain globally accessible in the bundle. The same applies to trigger handlers (`_runAutoFollowUps`). Export them from `src/entry.ts` and list them in `scripts/global-callbacks.json` (used by the build footer); `tests/callbacks.test.ts` enforces both.

**Logging:** Use `log.{debug,info,warn,error}` and `perf()` from `src/log.ts`, not raw `console.*`. Entries are structured objects (`message` plus fields like `threadId`, `durationMs`, serialized `error` with stack) so they're filterable in Cloud Logging (`exceptionLogging: STACKDRIVER`). The build footer wraps every global callback in `withLogging`, so uncaught errors are logged with the handler name and stack before being rethrown. Never log email addresses, subjects or bodies.
