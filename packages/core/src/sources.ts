/**
 * Cross-app source content: a read-only, explicitly authorised handoff of the
 * user's selection from one Office app to a task in another.
 *
 *   source app: user clicks "Share selection" → the host's own selection read
 *               capability runs through the registry → user previews and shares
 *   companion:  holds the shared content in memory only, for a limited time
 *   destination: user attaches a shared source to a request → it becomes a
 *               fenced, provenance-labelled context item for that task only
 *
 * There is no capability that lists, searches or fetches sources: the model
 * sees exactly the content the user attached, nothing else.
 */
import type { ContextItem, HostKind } from "./types.js";
import { fingerprint, truncate } from "./util.js";

export const SOURCE_MAX_CHARS = 50_000;
export const SOURCE_TTL_MS = 30 * 60_000;
export const SOURCE_APPS: HostKind[] = ["word", "excel", "powerpoint", "outlook"];

export const APP_NAMES: Record<HostKind, string> = { word: "Word", excel: "Excel", powerpoint: "PowerPoint", outlook: "Outlook" };

export interface SharedSource {
  id: string;
  /** Office app the content was read from. */
  app: HostKind;
  documentTitle?: string;
  documentId?: string;
  /** What was read: "Selection (paragraphs 3–5)", "Sheet Sales!A1:C10", "Message “Budget”". */
  scopeLabel: string;
  content: string;
  truncated: boolean;
  /** When the source app read the content. */
  retrievedAt: number;
  /** When the user shared it. */
  sharedAt: number;
  expiresAt: number;
}

export type SourceSummary = Omit<SharedSource, "content"> & { chars: number };

/** Content prepared for sharing, before the user confirms. */
export interface SourceCapture {
  app: HostKind;
  documentTitle?: string;
  documentId?: string;
  scopeLabel: string;
  content: string;
  truncated: boolean;
  retrievedAt: number;
}

export function sourceLabel(s: Pick<SharedSource, "app" | "documentTitle" | "scopeLabel">): string {
  return `From ${APP_NAMES[s.app]}${s.documentTitle ? ` “${s.documentTitle}”` : ""} — ${s.scopeLabel}`;
}

export function summarizeSource(s: SharedSource): SourceSummary {
  const { content, ...rest } = s;
  return { ...rest, chars: content.length };
}

/** A context item for an attached source, carrying its provenance. */
export function sourceContextItem(s: SharedSource, destination: HostKind): ContextItem {
  const t = truncate(s.content, SOURCE_MAX_CHARS);
  return {
    id: `source:${s.id}`,
    host: destination,
    label: sourceLabel(s),
    scope: "source",
    content: t.text,
    truncated: s.truncated || t.truncated,
    unsupported: [],
    fingerprint: fingerprint(s.content),
    retrievedAt: s.retrievedAt,
    source: "shared by user",
    provenance: { app: s.app, document: s.documentTitle, documentId: s.documentId, scope: s.scopeLabel, retrievedAt: s.retrievedAt, sharedAt: s.sharedAt },
  };
}

/** Validate a capture or a shared source received from elsewhere (e.g. the companion). */
export function validateSource(x: unknown): string | undefined {
  if (!x || typeof x !== "object") return "not an object";
  const s = x as Record<string, unknown>;
  if (!SOURCE_APPS.includes(s.app as HostKind)) return "unknown source app";
  if (typeof s.scopeLabel !== "string" || !s.scopeLabel || s.scopeLabel.length > 300) return "invalid scope";
  if (typeof s.content !== "string" || !s.content.trim()) return "empty content";
  if (s.content.length > SOURCE_MAX_CHARS) return `content longer than ${SOURCE_MAX_CHARS} characters`;
  if (typeof s.retrievedAt !== "number" || !Number.isFinite(s.retrievedAt)) return "missing retrieval time";
  for (const k of ["documentTitle", "documentId"]) if (s[k] !== undefined && (typeof s[k] !== "string" || (s[k] as string).length > 500)) return `invalid ${k}`;
  return undefined;
}
