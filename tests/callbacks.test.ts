import assert from "node:assert/strict";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, it } from "node:test";
import * as entry from "../src/entry.ts";

const root = join(import.meta.dirname, "..");
const globalCallbacks: string[] = JSON.parse(
  readFileSync(join(root, "scripts/global-callbacks.json"), "utf8"),
);

function referencedCallbacks(): Set<string> {
  const names = new Set<string>();
  for (const file of readdirSync(join(root, "src"), { recursive: true })) {
    if (!String(file).endsWith(".ts")) continue;
    const source = readFileSync(join(root, "src", String(file)), "utf8");
    for (const match of source.matchAll(
      /(?:setFunctionName|newTrigger)\(\s*"([^"]+)"/g,
    )) {
      names.add(match[1]);
    }
  }

  const manifest = readFileSync(join(root, "src/appsscript.json"), "utf8");
  for (const match of manifest.matchAll(/"runFunction":\s*"([^"]+)"/g)) {
    names.add(match[1]);
  }
  return names;
}

describe("global callbacks", () => {
  it("exposes every callback referenced by cards and the manifest", () => {
    const missing = [...referencedCallbacks()].filter(
      (name) => !globalCallbacks.includes(name),
    );
    assert.deepEqual(missing, [], "add these to scripts/global-callbacks.json");
  });

  it("only lists functions exported from the entry module", () => {
    const exported = entry as Record<string, unknown>;
    const notExported = globalCallbacks.filter(
      (name) => typeof exported[name] !== "function",
    );
    assert.deepEqual(notExported, [], "export these from src/entry.ts");
  });
});
