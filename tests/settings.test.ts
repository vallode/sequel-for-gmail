import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  clampFollowUpDays,
  clampStaleDays,
  parseSettingsFromFormInput,
  parseSettingsFromProperties,
  settingsToPropertyWrites,
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
      excluded_domains: " example.com , slack.com ",
    });
    assert.equal(s.days, 7);
    assert.equal(s.staleDays, 120);
    assert.equal(s.excludeReplied, false);
    assert.equal(s.excludeInternal, true);
    assert.equal(s.excludedDomains, "example.com , slack.com");
  });

  it("clamps follow-up days helper", () => {
    assert.equal(clampFollowUpDays(-5), 1);
    assert.equal(clampFollowUpDays(500), 365);
  });
});

describe("auto follow-up settings", () => {
  it("defaults to disabled with the default sequence when never saved", () => {
    const s = parseSettingsFromProperties({});
    assert.equal(s.autoFollowUpEnabled, false);
    assert.deepEqual(
      s.autoFollowUpSteps.map((step) => step.waitDays),
      [3, 4],
    );
  });

  it("keeps a saved empty sequence empty", () => {
    const s = parseSettingsFromProperties({ auto_followup_steps: "[]" });
    assert.deepEqual(s.autoFollowUpSteps, []);
  });

  it("parses sequence slots from the form, skipping empty templates", () => {
    const s = parseSettingsFromFormInput({
      auto_followup_enabled: "true",
      auto_followup_wait_1: "2",
      auto_followup_template_1: "  First nudge ",
      auto_followup_wait_2: "5",
      auto_followup_template_2: "   ",
      auto_followup_wait_3: "not-a-number",
      auto_followup_template_3: "Last nudge",
    });
    assert.equal(s.autoFollowUpEnabled, true);
    assert.deepEqual(s.autoFollowUpSteps, [
      { waitDays: 2, template: "First nudge" },
      { waitDays: 3, template: "Last nudge" },
    ]);
  });

  it("round-trips through property writes", () => {
    const s = parseSettingsFromFormInput({
      auto_followup_enabled: "true",
      auto_followup_wait_1: "3",
      auto_followup_template_1: "Hi {{firstName}}",
    });
    const restored = parseSettingsFromProperties(settingsToPropertyWrites(s));
    assert.deepEqual(restored, s);
  });
});
