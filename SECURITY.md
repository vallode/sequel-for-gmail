# Security policy

## Reporting a vulnerability

**Please do not open a public GitHub issue for security problems.**

Report vulnerabilities privately by opening a [GitHub Security Advisory](https://github.com/vallode/sequel-for-gmail/security/advisories/new) or emailing the repository owner through GitHub (see the repo **About** / owner profile). Include:

- A short description of the issue
- Steps to reproduce (if applicable)
- Impact you believe it has on users or their mail data

We will acknowledge receipt when possible and work on a fix before public disclosure. We appreciate responsible disclosure.

Do not include exploit proof-of-concepts, weaponized payloads, or instructions that could harm users in public channels.

## What the add-on stores

Sequel persists **per-user** data in Google Apps Script **user properties** for the signed-in Google account:

- Settings (follow-up days, filters, optional domain list, auto follow-up sequence, sort preference)
- A bounded list of dismissed reminders (thread identifiers and display metadata such as subject and recipient)

Short-lived **cache** entries (roughly five minutes) may hold serialized follow-up search results and the user’s email address for performance. Cache is not intended as durable storage.

Sequel does **not** maintain a separate Sequel-operated server database for user mail.

## OAuth and Gmail access

The add-on requests OAuth scopes limited to:

- Running the Gmail add-on
- Modifying Gmail (**read** mail, **manage labels**, and **create drafts** — it does not send mail)
- Reading the active user’s **email address**
- Managing the add-on's own **time-driven trigger** (hourly auto follow-up check, only while enabled)
- Connecting to an **external service**, restricted by the manifest's `urlFetchWhitelist` to `https://gmail.googleapis.com/` (the Gmail API batch endpoint, for thread-metadata requests)

Mail is accessed only to build the follow-up list, apply its `Follow-ups/*` labels, and, when auto follow-up is on, create draft replies on threads you flagged. Drafts are never sent automatically. Dismissed state and settings stay in your Google account’s Apps Script user properties.

## AGPL

Sequel is licensed under AGPL-3.0-only. Network users of the add-on can obtain corresponding source via the homepage **Source code** link and this repository.
