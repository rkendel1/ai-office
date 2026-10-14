#!/usr/bin/env node
/** Generate an Ed25519 key pair for signing releases. Keep the private key out of the repository. */
import crypto from "node:crypto";
import fs from "node:fs";

const dir = process.argv[2] ?? "keys";
fs.mkdirSync(dir, { recursive: true });
const { privateKey, publicKey } = crypto.generateKeyPairSync("ed25519");
fs.writeFileSync(`${dir}/release-private.pem`, privateKey.export({ type: "pkcs8", format: "pem" }), { mode: 0o600 });
fs.writeFileSync(`${dir}/release-public.pem`, publicKey.export({ type: "spki", format: "pem" }));
console.log(`Wrote ${dir}/release-private.pem (secret) and ${dir}/release-public.pem (publish this).`);
