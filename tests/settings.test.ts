import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  clampFollowUpDays,
  clampStaleDays,
  parseSettingsFromFormInput,
  parseSettingsFromProperties,
} from "../src/settings.ts";

describe("settings parsing and clamping", () => {
  it("uses defaults for missing properties", () => {
    const s = parseSettingsFromProperties({});
    assert.equal(s.days, 3);
    assert.equal(s.staleDays, 90);
    assert.equal(s.excludeReplied, true);
    assert.equal(s.excludeInternal, false);
  });

  it("clamps malformed follow-up days", () => {
    const s = parseSettingsFromProperties({ followup_days: "not-a-number" });
    assert.equal(s.days, 3);
    const high = parseSettingsFromProperties({ followup_days: "9999" });
    assert.equal(high.days, 365);
    const low = parseSettingsFromFormInput({ followup_days: "0" });
    assert.equal(low.days, 1);
  });

  it("clamps stale days above follow-up days", () => {
    assert.equal(clampStaleDays(10, 3), 10);
    assert.equal(clampStaleDays(3, 10), 11);
    assert.equal(clampStaleDays(0, 5), 6);
  });

  it("parses form toggles and trims strings", () => {
    const s = parseSettingsFromFormInput({
      followup_days: "7",
      stale_days: "120",
      exclude_replied: "false",
      exclude_internal: "true",
      auto_label: "  Follow up  ",
      excluded_domains: " example.com , slack.com ",
    });
    assert.equal(s.days, 7);
    assert.equal(s.staleDays, 120);
    assert.equal(s.excludeReplied, false);
    assert.equal(s.excludeInternal, true);
    assert.equal(s.autoLabel, "Follow up");
    assert.equal(s.excludedDomains, "example.com , slack.com");
  });

  it("clamps follow-up days helper", () => {
    assert.equal(clampFollowUpDays(-5), 1);
    assert.equal(clampFollowUpDays(500), 365);
  });
});
