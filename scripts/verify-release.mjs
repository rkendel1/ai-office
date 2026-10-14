#!/usr/bin/env node
/**
 * Verify a release directory: every file matches SHA256SUMS, no unlisted
 * files exist, and (if a public key is given) SHA256SUMS.sig is a valid
 * Ed25519 signature. Exit code 0 only if everything checks out.
 *
 *   node scripts/verify-release.mjs release/local-office-ai-0.1.0 keys/release-public.pem
 */
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";

const [dir, pubKeyPath] = process.argv.slice(2);
if (!dir) {
  console.error("usage: verify-release.mjs <release dir> [public key PEM]");
  process.exit(2);
}
const sums = fs.readFileSync(path.join(dir, "SHA256SUMS"), "utf8");
let ok = true;

if (pubKeyPath) {
  const sigFile = path.join(dir, "SHA256SUMS.sig");
  if (!fs.existsSync(sigFile)) {
    console.error("✗ SHA256SUMS.sig missing (release is not signed)");
    process.exit(1);
  }
  const valid = crypto.verify(null, Buffer.from(sums), crypto.createPublicKey(fs.readFileSync(pubKeyPath)), Buffer.from(fs.readFileSync(sigFile, "utf8").trim(), "base64"));
  console.log(valid ? "✓ signature valid" : "✗ signature INVALID");
  ok &&= valid;
} else console.warn("! no public key given: checksums only, authenticity not verified");

const listed = new Set();
for (const line of sums.trim().split("\n")) {
  const [hash, rel] = line.split(/ {2}/);
  listed.add(rel);
  const file = path.join(dir, rel);
  if (!fs.existsSync(file)) {
    console.error(`✗ missing ${rel}`);
    ok = false;
    continue;
  }
  const actual = crypto.createHash("sha256").update(fs.readFileSync(file)).digest("hex");
  if (actual !== hash) {
    console.error(`✗ modified ${rel}`);
    ok = false;
  }
}
const walk = (d) => {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const p = path.join(d, e.name);
    if (e.isDirectory()) walk(p);
    else {
      const rel = path.relative(dir, p).split(path.sep).join("/");
      if (!/^SHA256SUMS/.test(rel) && !listed.has(rel)) {
        console.error(`✗ unexpected file ${rel}`);
        ok = false;
      }
    }
  }
};
walk(dir);
console.log(ok ? `✓ ${listed.size} files verified` : "✗ verification failed");
process.exit(ok ? 0 : 1);
