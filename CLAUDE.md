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

**State:** User settings via `PropertiesService.getUserProperties()`. Dismissed emails are stored as JSON under `dismissed_ids` (capped, lock-protected updates). Follow-up results cached ~5 minutes in `CacheService` with hashed keys; Save & Refresh bypasses cache.

**Core logic:** `getPendingFollowUps()` paginates Gmail sent-thread search (bounded), filters threads in `follow-ups-core.ts` (replies, domains, internal, stale window), optionally auto-labels.

**OAuth scopes:** `gmail.addons.execute` (add-on runtime), `gmail.modify` (search threads and apply labels — covers all `GmailApp` usage), `userinfo.email` (`Session.getActiveUser().getEmail()`). Removed `gmail.readonly` (redundant with modify) and `script.scriptapp` (no Script Apps API usage).

**Action handlers:** Functions prefixed with `_on` or `_open` are Card Service callbacks and must remain globally accessible in the bundle. Export them from `src/entry.ts` and list them in `scripts/global-callbacks.json` (used by the build footer); `tests/callbacks.test.ts` enforces both.
