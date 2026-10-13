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
- [ ] **Auto-label** applies (or creates) the configured label on matching threads
- [ ] Re-open homepage within ~5 minutes without Save & Refresh; confirm **cache** behavior (stale until refresh path bypasses cache)

## Code expectations

- Production code must remain compatible with Apps Script V8 and Card Service globals (see README architecture).
- Prefer minimal diffs; match existing style (EditorConfig + project formatters).

## Questions

Open a [GitHub issue](https://github.com/vallode/sequel-for-gmail/issues) for bugs or feature ideas using the issue templates.
