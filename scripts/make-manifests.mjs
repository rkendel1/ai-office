#!/usr/bin/env node
/**
 * Generate Office manifests from templates.
 *   dist/manifest.xml / dist/outlook-manifest.xml          → served by the companion (https://localhost:47615)
 *   dist/manifest.dev.xml / dist/outlook-manifest.dev.xml  → Vite dev server (https://localhost:3000)
 * Override the production base with LOAI_BASE_URL.
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const addin = path.join(root, "packages/addin");
const out = path.join(addin, "dist");
fs.mkdirSync(out, { recursive: true });
const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));
const version = pkg.version.split(".").concat(["0", "0", "0"]).slice(0, 4).join(".");

const variants = [
  { suffix: "", file: "", base: process.env.LOAI_BASE_URL ?? "https://localhost:47615" },
  { suffix: " (dev)", file: ".dev", base: "https://localhost:3000" },
];
for (const tpl of ["manifest", "outlook-manifest"]) {
  const src = fs.readFileSync(path.join(addin, "manifests", `${tpl}.template.xml`), "utf8");
  for (const v of variants) {
    let xml = src.replaceAll("{{BASE}}", v.base).replaceAll("{{VERSION}}", version).replaceAll("{{SUFFIX}}", v.suffix);
    // Separate add-in ids so dev and installed versions can coexist.
    if (v.file === ".dev") xml = xml.replace(/<Id>([0-9a-f-]+)<\/Id>/, (_, id) => `<Id>${id.slice(0, -4)}d0e0</Id>`);
    fs.writeFileSync(path.join(out, `${tpl}${v.file}.xml`), xml);
  }
}
console.log(`manifests written to ${path.relative(root, out)}`);
