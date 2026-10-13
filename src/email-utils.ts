const EMAIL_IN_ANGLE = /<([^>]+)>/;
const SIMPLE_MAILBOX = /^[^\s@]+@[^\s@]+$/;

export function parseMailboxAddress(raw: string): string | null {
  const trimmed = raw.trim();
  if (!trimmed) return null;

  const angle = trimmed.match(EMAIL_IN_ANGLE);
  const candidate = (angle ? angle[1] : trimmed).trim().toLowerCase();
  if (!SIMPLE_MAILBOX.test(candidate)) return null;
  return candidate;
}

export function parseFirstRecipient(rawTo: string): string | null {
  const parts = rawTo.split(",");
  for (const part of parts) {
    const mailbox = parseMailboxAddress(part);
    if (mailbox) return mailbox;
  }
  return null;
}

export function getEmailDomain(email: string): string {
  const at = email.lastIndexOf("@");
  if (at === -1) return "";
  return email.slice(at + 1).toLowerCase();
}

export function normalizeDomainToken(token: string): string {
  let d = token.trim().toLowerCase();
  d = d.replace(/^@+/, "");
  while (d.startsWith(".")) d = d.slice(1);
  return d;
}

export function parseExcludedDomainList(excludedDomains: string): string[] {
  return excludedDomains.split(",").map(normalizeDomainToken).filter(Boolean);
}

export function isDomainExcluded(recipientDomain: string, excluded: string): boolean {
  const domain = recipientDomain.toLowerCase();
  const rule = normalizeDomainToken(excluded);
  if (!rule) return false;
  return domain === rule || domain.endsWith(`.${rule}`);
}

export function isAnyDomainExcluded(
  recipientDomain: string,
  domainList: string[],
): boolean {
  return domainList.some((d) => isDomainExcluded(recipientDomain, d));
}

export function emailsMatch(a: string, b: string): boolean {
  return a.trim().toLowerCase() === b.trim().toLowerCase();
}

export function isInternalRecipient(
  recipientEmail: string,
  myEmail: string,
  myDomain: string,
): boolean {
  if (emailsMatch(recipientEmail, myEmail)) return true;
  const recipientDomain = getEmailDomain(recipientEmail);
  if (!recipientDomain || !myDomain) return false;
  return recipientDomain === myDomain;
}

export function senderIsUser(fromHeader: string, myEmail: string): boolean {
  const parsed = parseMailboxAddress(fromHeader);
  if (!parsed) {
    return fromHeader.toLowerCase().includes(myEmail.toLowerCase());
  }
  return emailsMatch(parsed, myEmail);
}
