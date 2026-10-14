/**
 * Deterministic checks of a specialist's output against what was actually
 * observed: numbers must come from computed or inspected data, citations must
 * refer to real ranges/messages, and extracted items must be attributable.
 */
import type { ContextItem, ExecutionRecord } from "../../types.js";

// ----------------------------------------------------------------- numbers

/** All numbers present in successful tool outputs and inspected content. */
export function evidenceNumbers(actions: ExecutionRecord[], context: ContextItem[]): number[] {
  const out: number[] = [];
  const add = (s: string) => {
    for (const m of s.matchAll(/-?\d[\d,]*(?:\.\d+)?/g)) {
      const n = Number(m[0].replace(/,/g, ""));
      if (Number.isFinite(n)) out.push(n);
    }
  };
  for (const a of actions) if (a.status === "succeeded" || a.status === "unverified" || a.status === "partial") add(JSON.stringify(a.output ?? null));
  for (const c of context) add(c.content);
  return out;
}

export interface TextNumber {
  raw: string;
  value: number;
  decimals: number;
  percent: boolean;
  scale: number;
}

/**
 * Numbers in prose that make factual claims. Skips references (A1, M2, Q3,
 * slide 4, paragraph 12, row 7), list markers, years, and small integers.
 */
export function numbersInText(text: string): TextNumber[] {
  const out: TextNumber[] = [];
  const re = /(?<![A-Za-z0-9_.!$:])(\$?)(-?\d[\d,]*(?:\.\d+)?)(\s?%|\s?[KMB]\b|\s?(?:thousand|million|billion)\b)?/g;
  for (const m of text.matchAll(re)) {
    const idx = m.index ?? 0;
    const before = text.slice(Math.max(0, idx - 14), idx).toLowerCase();
    const after = text.slice(idx + m[0].length, idx + m[0].length + 2);
    if (/(slide|paragraph|para|row|column|col|step|item|page|section|message|q|fy|h|m|#|¶)\s*$/.test(before)) continue;
    if (/^[A-Za-z]/.test(after) && !m[3]) continue; // 3rd, 2x, 10am
    if (/^[:/]\d/.test(after)) continue; // times and dates
    const lineStart = text.lastIndexOf("\n", idx) + 1;
    if (/^\s*$/.test(text.slice(lineStart, idx)) && /^\.\s/.test(after)) continue; // "1. " list marker
    const num = m[2]!.replace(/,/g, "");
    const value = Number(num);
    if (!Number.isFinite(value)) continue;
    const suffix = (m[3] ?? "").trim().toLowerCase();
    const percent = suffix === "%";
    const scale = suffix === "k" || suffix === "thousand" ? 1e3 : suffix === "m" || suffix === "million" ? 1e6 : suffix === "b" || suffix === "billion" ? 1e9 : 1;
    const decimals = num.includes(".") ? num.split(".")[1]!.length : 0;
    if (!percent && scale === 1 && decimals === 0 && Math.abs(value) <= 10) continue;
    if (!percent && scale === 1 && decimals === 0 && value >= 1900 && value <= 2100 && !m[1]) continue;
    out.push({ raw: m[0].trim(), value, decimals, percent, scale });
  }
  return out;
}

function matches(n: TextNumber, e: number): boolean {
  const tol = (d: number) => 0.5 * 10 ** -d + 1e-9;
  const candidates = [e, e / n.scale];
  if (n.percent) candidates.push(e * 100);
  return candidates.some((c) => Math.abs(c - n.value) <= tol(n.decimals) || (n.scale > 1 && Math.abs(c - n.value) <= Math.max(tol(n.decimals), 0.05 * Math.abs(n.value))));
}

/** Numbers stated in `text` that do not correspond to any observed number. */
export function ungroundedNumbers(text: string, evidence: number[]): string[] {
  const nums = numbersInText(text);
  return nums.filter((n) => !evidence.some((e) => matches(n, e))).map((n) => n.raw);
}

/** Sheet!A1-style references in text. */
export function cellReferences(text: string): string[] {
  return [...text.matchAll(/(?:'[^'\n]+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g)].map((m) => m[0]);
}

// ----------------------------------------------------------- email threads

export interface ThreadMessageRef {
  id: string;
  author: string;
  text: string;
}

/** Parse the rendered thread produced by outlook.split_thread. */
export function parseRenderedThread(content: string): ThreadMessageRef[] {
  const out: ThreadMessageRef[] = [];
  const parts = content.split(/\n(?=\[M\d+\] From: )/);
  for (const p of parts) {
    const m = /^\[(M\d+)\] From: (.+?)(?: — Sent: [^\n]*)?\n([\s\S]*)$/.exec(p.trim().replace(/^[\s\S]*?(?=\[M\d+\] From: )/, ""));
    if (m) out.push({ id: m[1]!, author: m[2]!.trim(), text: m[3]!.trim() });
  }
  return out;
}

export interface ExtractedItem {
  kind: "explicit" | "inferred";
  text: string;
  owner?: string;
  due?: string;
  source?: string;
  line: string;
}

/**
 * Parse structured items:
 *   - [explicit] Send the revised budget — owner: Dana — due: Friday — source: M2
 *   - [inferred] Book a room — owner: unknown — source: M1
 */
export function parseItems(text: string): ExtractedItem[] {
  const out: ExtractedItem[] = [];
  for (const line of text.split("\n")) {
    const m = /^\s*(?:[-*•]|\d+[.)])\s*\[(explicit|inferred)\]\s*(.+)$/i.exec(line);
    if (!m) continue;
    const segs = m[2]!.split(/\s+[—–]\s+|\s+--\s+|\s+\|\s+/);
    const item: ExtractedItem = { kind: m[1]!.toLowerCase() as "explicit" | "inferred", text: segs[0]!.trim(), line: line.trim() };
    for (const s of segs.slice(1)) {
      const kv = /^(owner|due|source|by|from)\s*:\s*(.+)$/i.exec(s.trim());
      if (!kv) continue;
      const k = kv[1]!.toLowerCase();
      const v = kv[2]!.trim().replace(/[.;]$/, "");
      if (k === "owner" || k === "by") item.owner = v;
      else if (k === "due") item.due = v;
      else if (k === "source" || k === "from") item.source = (/M\d+/i.exec(v)?.[0] ?? v).toUpperCase();
    }
    out.push(item);
  }
  return out;
}

const STOP = new Set(["the", "and", "for", "with", "that", "this", "will", "from", "your", "you", "our", "are", "can", "send", "please", "have", "has", "need", "needs", "to", "by", "on", "of", "a", "an", "in", "it", "be", "is", "we", "i"]);
const words = (s: string) => s.toLowerCase().match(/[\p{L}\p{N}]{3,}/gu)?.filter((w) => !STOP.has(w)) ?? [];
const sameName = (a: string, b: string) => {
  const x = a.toLowerCase().replace(/<.*?>/g, "").trim();
  const y = b.toLowerCase().replace(/<.*?>/g, "").trim();
  return !!x && !!y && (x === y || x.split(/\s+/)[0] === y.split(/\s+/)[0]);
};

export interface ItemCheck {
  item: ExtractedItem;
  problem?: string;
}

/**
 * Check items against the thread:
 *  - every explicit item cites an existing message and shares wording with it;
 *  - "commitment" mode: an explicit commitment's owner is the author of the cited message;
 *  - "assignment" mode: the owner is the author or is named in the cited message.
 */
export function checkItems(items: ExtractedItem[], thread: ThreadMessageRef[], mode: "commitment" | "assignment"): ItemCheck[] {
  return items.map((item) => {
    if (item.kind === "inferred") return { item };
    const src = thread.find((m) => m.id === item.source);
    if (!src) return { item, problem: `explicit item cites ${item.source ? `unknown message ${item.source}` : "no source message"}` };
    const shared = words(item.text).filter((w) => src.text.toLowerCase().includes(w));
    if (!shared.length) return { item, problem: `no wording from ${src.id} supports “${item.text}”` };
    if (!item.owner) return { item, problem: "explicit item has no owner" };
    if (mode === "commitment" && !sameName(item.owner, src.author)) return { item, problem: `${src.id} was written by ${src.author}, not ${item.owner}` };
    if (mode === "assignment" && !sameName(item.owner, src.author) && !src.text.toLowerCase().includes(item.owner.toLowerCase().split(/\s+/)[0]!))
      return { item, problem: `${item.owner} is neither the author of ${src.id} nor named in it` };
    return { item };
  });
}
