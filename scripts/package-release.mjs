#!/usr/bin/env node
/**
 * Build a distributable release:
 *   release/local-office-ai-<version>/
 *     app/core, app/companion, app/addin      built code (no dev dependencies)
 *     install.sh, install.ps1, uninstall.*    installers
 *     SHA256SUMS (+ SHA256SUMS.sig when LOAI_SIGNING_KEY is set)
 *
 * Signing: LOAI_SIGNING_KEY = path to an Ed25519 private key (PEM).
 * Generate one with: node scripts/release-keys.mjs
 * Verify with:       node scripts/verify-release.mjs <release dir> <public key PEM>
 */
import { execSync } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));
const name = `local-office-ai-${pkg.version}`;
const out = path.join(root, "release", name);

const run = (cmd) => execSync(cmd, { cwd: root, stdio: "inherit" });

if (!process.argv.includes("--skip-build")) {
  run("npm run build");
  run("npx vitest run");
}

fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });

const copy = (from, to) => fs.cpSync(path.join(root, from), path.join(out, to), { recursive: true, filter: (p) => !p.endsWith(".map") });

// Runtime layout mirrors the workspace so the companion resolves @loai/core.
copy("packages/core/dist", "app/node_modules/@loai/core/dist");
fs.writeFileSync(
  path.join(out, "app/node_modules/@loai/core/package.json"),
  JSON.stringify({ name: "@loai/core", version: pkg.version, type: "module", main: "dist/index.js", types: "dist/index.d.ts" }, null, 2),
);
copy("packages/companion/dist", "app/companion/dist");
fs.writeFileSync(
  path.join(out, "app/companion/package.json"),
  JSON.stringify({ name: "@loai/companion", version: pkg.version, type: "module", bin: { "loai-companion": "dist/cli.js" } }, null, 2),
);
copy("packages/addin/dist", "app/addin/dist");
for (const f of ["install.sh", "install.ps1", "uninstall.sh", "uninstall.ps1"]) fs.copyFileSync(path.join(root, "installer", f), path.join(out, f));
fs.chmodSync(path.join(out, "install.sh"), 0o755);
fs.chmodSync(path.join(out, "uninstall.sh"), 0o755);
fs.copyFileSync(path.join(root, "README.md"), path.join(out, "README.md"));
// Docs (verification levels, limitations, acceptance and evaluation procedures) and the Outlook acceptance fixture.
copy("docs", "docs");
copy("acceptance/fixtures", "acceptance/fixtures");
fs.writeFileSync(path.join(out, "VERSION"), pkg.version + "\n");

// Checksums over every file (sorted, relative paths).
const files = [];
const walk = (d) => {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const p = path.join(d, e.name);
    if (e.isDirectory()) walk(p);
    else if (!/^SHA256SUMS/.test(e.name)) files.push(p);
  }
};
walk(out);
files.sort();
const sums = files.map((f) => `${crypto.createHash("sha256").update(fs.readFileSync(f)).digest("hex")}  ${path.relative(out, f).split(path.sep).join("/")}`).join("\n") + "\n";
fs.writeFileSync(path.join(out, "SHA256SUMS"), sums);

if (process.env.LOAI_SIGNING_KEY) {
  const key = crypto.createPrivateKey(fs.readFileSync(process.env.LOAI_SIGNING_KEY));
  const sig = crypto.sign(null, Buffer.from(sums), key).toString("base64");
  fs.writeFileSync(path.join(out, "SHA256SUMS.sig"), sig + "\n");
  console.log("Signed SHA256SUMS with", process.env.LOAI_SIGNING_KEY);
} else {
  console.warn("LOAI_SIGNING_KEY not set: release is checksummed but NOT signed.");
}

try {
  execSync(`tar -czf ${name}.tar.gz ${name}`, { cwd: path.join(root, "release") });
  console.log(`Release: release/${name}.tar.gz`);
} catch {
  console.log(`Release directory: release/${name}`);
}
