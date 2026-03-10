# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
deno task build    # Compile TypeScript → build/
deno task push     # Build + push to Google Apps Script
deno task deploy   # Build + push + create new deployment
deno task pull     # Pull latest from Google Apps Script
deno task open     # Open the script in browser
```

There are no automated tests. Validation is done by deploying and testing manually in Gmail.

## Architecture

Sequel is a Gmail add-on built on Google Apps Script (V8 runtime), written in TypeScript and deployed via `clasp`. The entire application lives in `src/Code.ts`.

**Build pipeline:** TypeScript → `tsc` → `build/Code.js` + `build/appsscript.json` → `clasp push` → Google Apps Script.

**Runtime:** All code runs inside Google Apps Script — no Node.js, no bundler, no modules. All functions must be globally scoped to be callable by the Card Service framework. Interfaces and types are compile-time only.

**UI:** Built with the Google Card Service API. Cards are returned from entry point functions declared in `src/appsscript.json` (`buildSettingsCard`, `_buildHomepage`, `buildDismissedCard`).

**State:** User settings are persisted via `PropertiesService.getUserProperties()` (key-value string store). Dismissed emails are stored as a JSON array under the `dismissed_emails` property key. Results are cached for 5 minutes using `CacheService`.

**Core logic:** `getPendingFollowUps()` queries Gmail for sent threads older than N days, then filters out threads that have replies, match excluded domains, match an auto-label, are internal, or exceed the stale cutoff.

**Action handlers:** Functions prefixed with `_on` (e.g. `_onSaveSettings`, `_onDismiss`) are registered as Card Service action callbacks. They must be globally accessible.
