import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  isAnyDomainExcluded,
  isInternalRecipient,
  parseExcludedDomainList,
  parseFirstRecipient,
  parseMailboxAddress,
  senderIsUser,
} from "../src/email-utils.ts";

describe("email and domain normalization", () => {
  it("parses mailbox addresses", () => {
    assert.equal(parseMailboxAddress("Alice <Alice@Example.com>"), "alice@example.com");
    assert.equal(parseMailboxAddress("bob@test.com"), "bob@test.com");
    assert.equal(parseMailboxAddress("not an email"), null);
  });

  it("parses first recipient from lists", () => {
    assert.equal(
      parseFirstRecipient('"Team" <team@corp.com>, bob@x.com'),
      "team@corp.com",
    );
  });

  it("detects internal mail case-insensitively", () => {
    assert.equal(isInternalRecipient("Me@Corp.com", "me@corp.com", "corp.com"), true);
    assert.equal(isInternalRecipient("peer@corp.com", "me@corp.com", "corp.com"), true);
    assert.equal(
      isInternalRecipient("client@other.com", "me@corp.com", "corp.com"),
      false,
    );
  });

  it("matches excluded domains and subdomains only", () => {
    const list = parseExcludedDomainList("@slack.com, company.com");
    assert.deepEqual(list, ["slack.com", "company.com"]);
    assert.equal(isAnyDomainExcluded("mail.slack.com", list), true);
    assert.equal(isAnyDomainExcluded("notslack.com", list), false);
    assert.equal(isAnyDomainExcluded("evilcompany.com", list), false);
  });

  it("detects sender as user with display names", () => {
    assert.equal(senderIsUser("Me <me@corp.com>", "me@corp.com"), true);
    assert.equal(senderIsUser("other@corp.com", "me@corp.com"), false);
  });
});
