import { cpSync, mkdirSync, readFileSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const outDir = join(root, "build");
const outFile = join(outDir, "Code.js");

// Shared with tests/callbacks.test.ts, which checks it against the source.
const GLOBAL_CALLBACKS = JSON.parse(
  readFileSync(join(root, "scripts/global-callbacks.json"), "utf8"),
);

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
