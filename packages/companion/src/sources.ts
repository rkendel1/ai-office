import crypto from "node:crypto";
import { SOURCE_TTL_MS, summarizeSource, validateSource, type SharedSource, type SourceSummary } from "@loai/core";

const MAX_SOURCES = 10;

/**
 * Shared cross-app content. Memory only: never written to disk, gone when the
 * companion stops, expires after SOURCE_TTL_MS, at most MAX_SOURCES at a time.
 */
export class SourceStore {
  private readonly items = new Map<string, SharedSource>();

  constructor(private readonly now: () => number = Date.now) {}

  private sweep() {
    const t = this.now();
    for (const [id, s] of this.items) if (s.expiresAt <= t) this.items.delete(id);
  }

  add(input: unknown): { source?: SourceSummary; error?: string } {
    const problem = validateSource(input);
    if (problem) return { error: problem };
    const s = input as Omit<SharedSource, "id" | "sharedAt" | "expiresAt">;
    this.sweep();
    while (this.items.size >= MAX_SOURCES) this.items.delete(this.items.keys().next().value!);
    const t = this.now();
    const source: SharedSource = {
      id: crypto.randomUUID(),
      app: s.app,
      documentTitle: s.documentTitle,
      documentId: s.documentId,
      scopeLabel: s.scopeLabel,
      content: s.content,
      truncated: !!s.truncated,
      retrievedAt: s.retrievedAt,
      sharedAt: t,
      expiresAt: t + SOURCE_TTL_MS,
    };
    this.items.set(source.id, source);
    return { source: summarizeSource(source) };
  }

  list(): SourceSummary[] {
    this.sweep();
    return [...this.items.values()].sort((a, b) => b.sharedAt - a.sharedAt).map(summarizeSource);
  }

  get(id: string): SharedSource | undefined {
    this.sweep();
    return this.items.get(id);
  }

  delete(id: string): boolean {
    return this.items.delete(id);
  }

  clear(): number {
    const n = this.items.size;
    this.items.clear();
    return n;
  }
}
