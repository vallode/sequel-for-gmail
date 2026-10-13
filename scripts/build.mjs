import { cpSync, mkdirSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const outDir = join(root, "build");
const outFile = join(outDir, "Code.js");

const GLOBAL_CALLBACKS = [
  "_buildHomepage",
  "buildSettingsCard",
  "buildDismissedCard",
  "_onSaveSettings",
  "_onBack",
  "_onBackFromDismissed",
  "_onToggleSort",
  "_openSettingsCard",
  "_openDismissedCard",
  "_onDismissEmail",
  "_onRestoreEmail",
  "_onClearDismissed",
];

rmSync(outDir, { recursive: true, force: true });
mkdirSync(outDir, { recursive: true });

const footer = GLOBAL_CALLBACKS.map(
  (name) => `globalThis.${name} = Sequel.${name};`,
).join("\n");

await esbuild.build({
  entryPoints: [join(root, "src/entry.ts")],
  bundle: true,
  outfile: outFile,
  format: "iife",
  globalName: "Sequel",
  target: "es2019",
  platform: "browser",
  footer: { js: footer },
  logLevel: "info",
});

cpSync(join(root, "src/appsscript.json"), join(outDir, "appsscript.json"));

console.log(`Wrote ${outFile}`);
