import type { ContextItem, ContextScope, HostKind, UnsupportedContent } from "../types.js";
import { fingerprint, truncate } from "../util.js";

export const MAX_ITEM_CHARS = 60_000;

export function makeItem(p: {
  host: HostKind;
  id: string;
  label: string;
  scope: ContextScope;
  content: string;
  source: string;
  truncated?: boolean;
  unsupported?: UnsupportedContent[];
  /** Fingerprint of the full source (defaults to the content). */
  sourceFingerprint?: string;
}): ContextItem {
  const t = truncate(p.content, MAX_ITEM_CHARS);
  return {
    id: p.id,
    host: p.host,
    label: p.label,
    scope: p.scope,
    content: t.text,
    truncated: !!p.truncated || t.truncated,
    unsupported: p.unsupported ?? [],
    fingerprint: p.sourceFingerprint ?? fingerprint(p.content),
    retrievedAt: Date.now(),
    source: p.source,
  };
}

/** Compact before/after excerpt for previews. */
export function excerpt(s: string, max = 600): string {
  return s.length > max ? s.slice(0, max) + `… (${s.length} characters)` : s;
}
