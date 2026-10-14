import type { ContextItem, ContextPackage, HostKind, ItemIdentity } from "./types.js";
import { truncate } from "./util.js";

export type Freshness = "fresh" | "stale" | "unknown";

/**
 * Tracks what the assistant has actually inspected, packages it within a
 * budget, and records freshness. The UI reads `inspected()` to show exactly
 * which content was used — a selection is never presented as the full document.
 */
export class ContextManager {
  private readonly items = new Map<string, ContextItem & { freshness: Freshness }>();
  private listeners = new Set<() => void>();

  register(items: ContextItem[]): void {
    for (const it of items) this.items.set(it.id, { ...it, freshness: "fresh" });
    this.emit();
  }

  inspected(): (ContextItem & { freshness: Freshness })[] {
    return [...this.items.values()].sort((a, b) => b.retrievedAt - a.retrievedAt);
  }

  get(id: string) {
    return this.items.get(id);
  }

  setFreshness(id: string, f: Freshness): void {
    const it = this.items.get(id);
    if (it) {
      it.freshness = f;
      this.emit();
    }
  }

  /** After the assistant modifies a host, everything previously read from it may be out of date. */
  markHostStale(host: HostKind): void {
    for (const it of this.items.values()) if (it.host === host && it.scope !== "metadata") it.freshness = "stale";
    this.emit();
  }

  clear(): void {
    this.items.clear();
    this.emit();
  }

  remove(pred: (item: ContextItem) => boolean): void {
    let changed = false;
    for (const [id, it] of this.items) if (pred(it)) {
      this.items.delete(id);
      changed = true;
    }
    if (changed) this.emit();
  }

  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  private emit() {
    for (const l of this.listeners) l();
  }

  /**
   * Build a bounded context package. Priority: selection, then fresh items
   * (most recent first), then stale items. Items that do not fit are
   * truncated or omitted, and that is reported.
   */
  buildPackage(host: HostKind, identity: ItemIdentity, budgetChars: number, onlyIds?: string[]): ContextPackage {
    const all = this.inspected().filter((i) => i.host === host && (!onlyIds || onlyIds.includes(i.id)));
    const rank = (i: (typeof all)[number]) =>
      (i.scope === "selection" ? 0 : i.scope === "source" ? 1 : i.scope === "metadata" ? 2 : 3) + (i.freshness === "stale" ? 10 : 0);
    const ordered = [...all].sort((a, b) => rank(a) - rank(b) || b.retrievedAt - a.retrievedAt);

    const included: ContextItem[] = [];
    const omitted: string[] = [];
    let used = 0;
    const OVERHEAD = 160; // envelope chars per item
    for (const it of ordered) {
      const remaining = budgetChars - used - OVERHEAD;
      if (remaining < 200) {
        omitted.push(it.id);
        continue;
      }
      const t = truncate(it.content, remaining);
      const { freshness: _f, ...plain } = it;
      included.push({ ...plain, content: t.text, truncated: it.truncated || t.truncated });
      used += t.text.length + OVERHEAD;
    }
    return {
      host,
      identity,
      items: included,
      usedChars: used,
      budgetChars,
      truncated: omitted.length > 0 || included.some((i) => i.truncated),
      omittedItemIds: omitted,
    };
  }
}

const OPEN = "<<<OFFICE_CONTENT";
const CLOSE = "<<<END_OFFICE_CONTENT>>>";

/** Neutralise anything in untrusted content that could imitate our envelope. */
export function escapeUntrusted(text: string): string {
  return text.replace(/<<</g, "‹‹‹").replace(/>>>/g, "›››");
}

function attr(v: string): string {
  return escapeUntrusted(v).replace(/"/g, "'").replace(/[\r\n]+/g, " ");
}

/** Render a context package for the model, with untrusted content clearly fenced. */
export function renderContextPackage(pkg: ContextPackage, staleIds: string[] = []): string {
  const lines: string[] = [];
  lines.push(
    `Active ${pkg.host} item: ${pkg.identity.title ? attr(pkg.identity.title) : "(untitled)"}.`,
    pkg.items.length
      ? `You have inspected ${pkg.items.length} context item(s)${pkg.truncated ? " (some content truncated or omitted to fit the budget)" : ""}.`
      : "You have not inspected any content yet. Use a read capability before making claims about the content.",
  );
  if (pkg.omittedItemIds.length) lines.push(`Omitted for budget: ${pkg.omittedItemIds.join(", ")}.`);
  for (const it of pkg.items) {
    const unsupported = it.unsupported.length
      ? ` unsupported="${it.unsupported.map((u) => `${u.count} ${u.kind}`).join(", ")}"`
      : "";
    const stale = staleIds.includes(it.id) ? ` stale="true"` : "";
    const p = it.provenance;
    const prov = p
      ? ` from_app="${p.app}"${p.document ? ` from_document="${attr(p.document)}"` : ""} from_scope="${attr(p.scope)}" retrieved="${new Date(p.retrievedAt).toISOString()}"`
      : "";
    lines.push(
      `${OPEN} id="${attr(it.id)}" label="${attr(it.label)}" scope="${it.scope}" truncated="${it.truncated}"${unsupported}${stale}${prov}>>>`,
      escapeUntrusted(it.content),
      CLOSE,
    );
  }
  return lines.join("\n");
}

/** Short reference list for history and the UI (no content). */
export function contextRefs(items: ContextItem[]) {
  return items.map((i) => ({
    id: i.id,
    label: i.label,
    scope: i.scope,
    truncated: i.truncated,
    fingerprint: i.fingerprint,
  }));
}
