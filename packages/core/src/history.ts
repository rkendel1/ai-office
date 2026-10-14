import type { ExecutionRecord, HostKind, Json } from "./types.js";

export type TaskOutcome = "completed" | "partial" | "failed" | "cancelled";

export interface ContextRef {
  id: string;
  label: string;
  scope: string;
  truncated: boolean;
  fingerprint: string;
}

/** Stored per task. By default contains references and short excerpts, not document contents. */
export interface TaskRecord {
  id: string;
  request: string;
  model: string;
  host: HostKind;
  itemTitle?: string;
  contextRefs: ContextRef[];
  actions: ExecutionRecord[];
  outcome: TaskOutcome;
  finalText: string;
  errors: string[];
  startedAt: number;
  finishedAt: number;
}

export interface HistoryStore {
  add(record: TaskRecord): Promise<void>;
  list(limit?: number): Promise<TaskRecord[]>;
  get(id: string): Promise<TaskRecord | undefined>;
  clear(): Promise<void>;
  /** Delete records older than `days`. 0 = keep nothing. */
  prune(days: number): Promise<number>;
}

export class MemoryHistoryStore implements HistoryStore {
  private records: TaskRecord[] = [];
  async add(r: TaskRecord) {
    this.records.unshift(r);
  }
  async list(limit = 50) {
    return this.records.slice(0, limit);
  }
  async get(id: string) {
    return this.records.find((r) => r.id === id);
  }
  async clear() {
    this.records = [];
  }
  async prune(days: number) {
    const cutoff = Date.now() - days * 86_400_000;
    const before = this.records.length;
    this.records = this.records.filter((r) => r.finishedAt >= cutoff && days > 0);
    return before - this.records.length;
  }
}

const EXCERPT = 160;

function redactJson(v: Json | undefined): Json | undefined {
  if (v === undefined || v === null) return v;
  if (typeof v === "string") return v.length > EXCERPT ? v.slice(0, EXCERPT) + `… [${v.length} chars]` : v;
  if (Array.isArray(v)) {
    const head = v.slice(0, 5).map((x) => redactJson(x) as Json);
    return v.length > 5 ? [...head, `… [${v.length} items]`] : head;
  }
  if (typeof v === "object") return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, redactJson(x) as Json]));
  return v;
}

/** Strip document content from a task record unless the user opted into full storage. */
export function minimizeRecord(r: TaskRecord, storeFull: boolean): TaskRecord {
  if (storeFull) return r;
  return {
    ...r,
    finalText: r.finalText.length > 2000 ? r.finalText.slice(0, 2000) + "…" : r.finalText,
    actions: r.actions.map((a) => ({
      ...a,
      args: redactJson(a.args) as Record<string, Json>,
      output: redactJson(a.output),
      verification: a.verification
        ? { ...a.verification, evidence: a.verification.evidence?.slice(0, EXCERPT) }
        : undefined,
      compensation: undefined, // contains prior content; only kept in-session
    })),
  };
}
