import { createRequire } from "node:module";
import type { HistoryStore, TaskRecord } from "@loai/core";

// node:sqlite is built into Node ≥ 22.13; loaded via require to keep the
// experimental warning out of module evaluation order issues.
const require = createRequire(import.meta.url);
const { DatabaseSync } = require("node:sqlite") as typeof import("node:sqlite");

type DB = InstanceType<typeof DatabaseSync>;

/**
 * Minimal local persistence: task history (references, not document copies),
 * settings, and paired client token hashes.
 */
export class CompanionStore implements HistoryStore {
  readonly db: DB;

  constructor(path: string) {
    this.db = new DatabaseSync(path);
    this.db.exec(`
      PRAGMA journal_mode = WAL;
      CREATE TABLE IF NOT EXISTS history (
        id TEXT PRIMARY KEY,
        started_at INTEGER NOT NULL,
        finished_at INTEGER NOT NULL,
        host TEXT NOT NULL,
        outcome TEXT NOT NULL,
        record TEXT NOT NULL
      );
      CREATE INDEX IF NOT EXISTS history_finished ON history(finished_at DESC);
      CREATE TABLE IF NOT EXISTS kv (key TEXT PRIMARY KEY, value TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS clients (
        id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        token_hash TEXT NOT NULL UNIQUE,
        created_at INTEGER NOT NULL,
        last_seen INTEGER
      );
    `);
  }

  async add(r: TaskRecord): Promise<void> {
    this.db
      .prepare("INSERT OR REPLACE INTO history (id, started_at, finished_at, host, outcome, record) VALUES (?, ?, ?, ?, ?, ?)")
      .run(r.id, r.startedAt, r.finishedAt, r.host, r.outcome, JSON.stringify(r));
  }

  async list(limit = 50): Promise<TaskRecord[]> {
    const rows = this.db.prepare("SELECT record FROM history ORDER BY finished_at DESC LIMIT ?").all(Math.min(500, Math.max(1, limit))) as { record: string }[];
    return rows.map((r) => JSON.parse(r.record));
  }

  async get(id: string): Promise<TaskRecord | undefined> {
    const row = this.db.prepare("SELECT record FROM history WHERE id = ?").get(id) as { record: string } | undefined;
    return row ? JSON.parse(row.record) : undefined;
  }

  async clear(): Promise<void> {
    this.db.exec("DELETE FROM history");
    // Reclaim space so deleted records are not left in the file.
    this.db.exec("VACUUM");
  }

  async prune(days: number): Promise<number> {
    const cutoff = days <= 0 ? Number.MAX_SAFE_INTEGER : Date.now() - days * 86_400_000;
    const res = this.db.prepare("DELETE FROM history WHERE finished_at < ?").run(cutoff);
    return Number(res.changes);
  }

  getSetting(key: string): unknown {
    const row = this.db.prepare("SELECT value FROM kv WHERE key = ?").get(key) as { value: string } | undefined;
    return row ? JSON.parse(row.value) : undefined;
  }

  setSetting(key: string, value: unknown): void {
    this.db.prepare("INSERT OR REPLACE INTO kv (key, value) VALUES (?, ?)").run(key, JSON.stringify(value));
  }

  addClient(id: string, name: string, tokenHash: string): void {
    this.db.prepare("INSERT INTO clients (id, name, token_hash, created_at) VALUES (?, ?, ?, ?)").run(id, name, tokenHash, Date.now());
  }

  findClient(tokenHash: string): { id: string; name: string } | undefined {
    const row = this.db.prepare("SELECT id, name FROM clients WHERE token_hash = ?").get(tokenHash) as { id: string; name: string } | undefined;
    if (row) this.db.prepare("UPDATE clients SET last_seen = ? WHERE id = ?").run(Date.now(), row.id);
    return row;
  }

  listClients(): { id: string; name: string; created_at: number; last_seen: number | null }[] {
    return this.db.prepare("SELECT id, name, created_at, last_seen FROM clients ORDER BY created_at").all() as any;
  }

  revokeClients(): number {
    return Number(this.db.prepare("DELETE FROM clients").run().changes);
  }

  close(): void {
    this.db.close();
  }
}
