# Contributing to Sequel for Gmail

Thanks for helping improve Sequel. This is a small project: keep changes focused and verify behavior in Gmail before opening a pull request.

## Development setup

1. Install prerequisites from [README.md](README.md) (Node LTS, Corepack + pnpm 10, clasp login).
2. Clone the repo and run:

   ```bash
   pnpm install
   ```

3. Do **not** commit clasp login or credential files (for example `.clasprc.json`). They are gitignored; if you accidentally stage them, remove them from the commit.

## Before you open a PR

Run the full local gate:

```bash
pnpm check
```

Fix any format, lint, type, test, or build failures. Describe what you changed and how you tested in the pull request template.

## Manual Gmail smoke test

Automated tests do not cover Gmail. After `pnpm push` (or your tester deployment flow), exercise the add-on in Gmail:

- [ ] Open the add-on homepage; list loads without errors
- [ ] **Settings → Save & Refresh** persists values and refreshes the list (cache bypass)
- [ ] Toggle **sort** (oldest / newest first); order updates
- [ ] **Dismiss** a thread; it disappears from the homepage
- [ ] **Dismissed** view shows the entry; **Restore** returns it to the homepage
- [ ] **Restore All** clears dismissed state as expected
- [ ] With enough matching mail, confirm **truncation messaging** when results exceed UI/search limits
- [ ] **Excluded domains** skip intended recipients
- [ ] First load creates `Follow-ups` with `Waiting`, `Auto`, `Done` nested under it; pending threads get `Follow-ups/Waiting`, and a thread that gets a reply loses it on the next refresh
- [ ] **Settings → Auto follow-up drafts**: enable, edit a step, Save & Refresh; the homepage shows the status line and per-row **Auto follow-up** buttons, and one `_runAutoFollowUps` trigger exists (Apps Script editor → Triggers)
- [ ] Flag a thread, run `_runAutoFollowUps` from the editor (with a 1-day step on an older thread); a threaded draft appears addressed to the original recipients, and a second run does not add another draft
- [ ] A flagged thread that got a reply moves from `Follow-ups/Auto` to `Follow-ups/Done` on the next run
- [ ] Disable auto follow-up and save; the trigger is removed
- [ ] Re-open homepage within ~30 minutes; confirm **cache** behavior ("Updated X min ago" keeps counting, dismiss/sort don't rescan)
- [ ] **Refresh** on the homepage refetches and resets the label to "Updated just now"

## Releases and dependency maintenance

**Monthly, or when a security advisory lands:**

1. Run `pnpm outdated`, then `pnpm update` (or `pnpm update --latest` for majors, one at a time).
2. Run `pnpm check` and commit the updated `pnpm-lock.yaml` with a note on what changed.
3. Check the Node LTS version in `.github/workflows/ci.yml` and `package.json` `engines`, and bump them when a new LTS becomes active.

**Cutting a release:**

1. Make sure `main` is green in CI.
2. Run `pnpm push` and work through the manual Gmail smoke test above against `@HEAD`.
3. Run `pnpm deploy`, and record the new Apps Script version number.
4. Tag the commit and push the tag:

   ```bash
   git tag -a v1.2.0 -m "Apps Script version 12: <summary>"
   git push origin v1.2.0
   ```

5. Create a GitHub release from the tag with the deployment notes: user-visible changes, any OAuth scope changes (these make users re-authorize), and the Apps Script version to roll back to (see README → Rollback).

Use semantic versioning: bump major for scope changes or breaking settings changes, minor for features, patch for fixes.

## Code expectations

- Production code must remain compatible with Apps Script V8 and Card Service globals (see README architecture).
- New Card Service callbacks (`setFunctionName(...)`) must be exported from `src/entry.ts` and listed in `scripts/global-callbacks.json`; `pnpm test` fails otherwise.
- Prefer minimal diffs; match existing style (EditorConfig + project formatters).

## Questions

Open a [GitHub issue](https://github.com/vallode/sequel-for-gmail/issues) for bugs or feature ideas using the issue templates.
