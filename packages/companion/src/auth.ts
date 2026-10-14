import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import type { CompanionStore } from "./store.js";

export function hashToken(token: string): string {
  return crypto.createHash("sha256").update(token).digest("hex");
}

export function randomToken(bytes = 32): string {
  return crypto.randomBytes(bytes).toString("base64url");
}

function safeEqual(a: string, b: string): boolean {
  const x = Buffer.from(a);
  const y = Buffer.from(b);
  return x.length === y.length && crypto.timingSafeEqual(x, y);
}

/**
 * Admin token: a random secret in a 0600 file in the companion's home
 * directory. Only processes running as the user can read it (used by the CLI
 * to issue pairing codes). It is never sent to the browser.
 */
export function loadAdminToken(home: string): string {
  const file = path.join(home, "admin.token");
  if (fs.existsSync(file)) {
    const t = fs.readFileSync(file, "utf8").trim();
    if (t.length >= 32) return t;
  }
  const t = randomToken();
  fs.writeFileSync(file, t, { mode: 0o600 });
  try {
    fs.chmodSync(file, 0o600);
  } catch {
    /* Windows: ACLs of the user profile directory apply */
  }
  return t;
}

const PAIRING_TTL_MS = 10 * 60_000;
const MAX_ATTEMPTS = 5;
const ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // no 0/O/1/I

/**
 * Pairing: the user reads a short one-time code from the companion (console or
 * `loai-companion pair`) and enters it in the add-in. The add-in exchanges it
 * for a long-lived client token. Codes expire, are single use, and are
 * invalidated after a few wrong guesses.
 */
export class Pairing {
  private code: string | null = null;
  private expires = 0;
  private attempts = 0;

  constructor(
    private readonly store: CompanionStore,
    private readonly now: () => number = Date.now,
  ) {}

  issueCode(): { code: string; expiresAt: number } {
    const bytes = crypto.randomBytes(8);
    let c = "";
    for (const b of bytes) c += ALPHABET[b % ALPHABET.length];
    this.code = `${c.slice(0, 4)}-${c.slice(4)}`;
    this.expires = this.now() + PAIRING_TTL_MS;
    this.attempts = 0;
    return { code: this.code, expiresAt: this.expires };
  }

  /** Returns a new client token, or an error string. */
  redeem(code: string, clientName: string): { token: string } | { error: string } {
    if (!this.code || this.now() > this.expires) return { error: "No active pairing code. Run `loai-companion pair` to get one." };
    const normalized = code.toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (!safeEqual(normalized, this.code.replace("-", ""))) {
      this.attempts++;
      if (this.attempts >= MAX_ATTEMPTS) {
        this.code = null;
        return { error: "Too many wrong attempts. Get a new pairing code." };
      }
      return { error: "Incorrect pairing code." };
    }
    this.code = null;
    const token = randomToken();
    this.store.addClient(crypto.randomUUID(), clientName.slice(0, 80) || "Office add-in", hashToken(token));
    return { token };
  }

  get active(): boolean {
    return !!this.code && this.now() <= this.expires;
  }
}

export type Principal = { kind: "admin" } | { kind: "client"; id: string; name: string };

export function authenticate(header: string | undefined, adminToken: string, store: CompanionStore): Principal | null {
  if (!header?.startsWith("Bearer ")) return null;
  const token = header.slice(7).trim();
  if (!token || token.length > 200) return null;
  if (safeEqual(token, adminToken)) return { kind: "admin" };
  const c = store.findClient(hashToken(token));
  return c ? { kind: "client", ...c } : null;
}
