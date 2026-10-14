/**
 * Reusable completion criteria. Each is evaluated deterministically against
 * observed execution records, inspected context and host snapshots — never
 * against the model's own claims of success.
 */
import { parseRange } from "../../hosts/a1.js";
import type { ExecutionRecord } from "../../types.js";
import type { CompletionCriterion, CriterionResult, TaskEvidence } from "../types.js";
import { cellReferences, checkItems, evidenceNumbers, parseItems, parseRenderedThread, ungroundedNumbers } from "./evidence.js";

const EXECUTED = new Set(["succeeded", "unverified", "partial", "failed"]);

export const succeededActions = (e: TaskEvidence, names: string[]) => e.actions.filter((a) => names.includes(a.capability) && a.status === "succeeded");

/** At least `min` actions among `names` succeeded with verification. */
export function verifiedAction(id: string, description: string, names: string[], min = 1, required = true): CompletionCriterion {
  return {
    id,
    description,
    required,
    evaluate(e) {
      const ok = succeededActions(e, names);
      const bad = e.actions.filter((a) => names.includes(a.capability) && ["failed", "unverified", "partial"].includes(a.status));
      return {
        met: ok.length >= min,
        detail: ok.length >= min ? `${ok.length} verified` : `${ok.length} of ${min} verified${bad.length ? `; ${bad.length} failed or unverified` : ""}`,
      };
    },
  };
}

/** Some content matching the predicate was inspected. */
export function inspected(id: string, description: string, pred: (c: TaskEvidence["context"][number]) => boolean, required = true): CompletionCriterion {
  return {
    id,
    description,
    required,
    evaluate: (e) => {
      const hits = e.context.filter(pred);
      return { met: hits.length > 0, detail: hits.length ? hits.map((h) => h.label).slice(0, 3).join("; ") : "nothing relevant was read" };
    },
  };
}

export function noModifications(): CompletionCriterion {
  return {
    id: "no-modifications",
    description: "The source was not modified",
    required: true,
    evaluate: (e) => {
      const mods = e.actions.filter((a) => e.isModifying(a.capability) && EXECUTED.has(a.status));
      return { met: mods.length === 0, detail: mods.length ? `${mods.length} change(s) were made` : "read-only" };
    },
  };
}

export function textCriterion(id: string, description: string, test: (text: string, e: TaskEvidence) => CriterionResult, required = true): CompletionCriterion {
  return { id, description, required, evaluate: (e) => test(e.text, e) };
}

/** Numbers in the final answer must appear in computed results or inspected data. */
export function numbersGrounded(required = true): CompletionCriterion {
  return {
    id: "numbers-grounded",
    description: "Every number reported comes from calculated results or inspected data",
    required,
    evaluate(e) {
      const bad = ungroundedNumbers(e.text, evidenceNumbers(e.actions, e.context));
      return { met: bad.length === 0, detail: bad.length ? `not found in any computed or inspected data: ${bad.slice(0, 6).join(", ")}` : "all numbers traced to evidence" };
    },
  };
}

export function citesRanges(required = false): CompletionCriterion {
  return textCriterion(
    "cites-ranges",
    "Results cite worksheet ranges",
    (t) => {
      const refs = cellReferences(t);
      return { met: refs.length > 0, detail: refs.length ? refs.slice(0, 4).join(", ") : "no Sheet!Range references in the answer" };
    },
    required,
  );
}

/** If any content read was partial or truncated, the answer must say so. */
export function disclosesPartialRead(): CompletionCriterion {
  return {
    id: "discloses-partial",
    description: "States when only part of the content was inspected",
    required: true,
    evaluate(e) {
      const partial = e.context.filter((c) => c.scope !== "metadata" && c.source !== "initial" && (c.truncated || c.scope === "partial" || c.scope === "selection"));
      const full = e.context.some((c) => c.scope === "full" && !c.truncated);
      if (!partial.length || full) return { met: true, detail: full ? "full content inspected" : "no partial reads" };
      const says = /\b(partial|partly|truncated|only (the )?(first|part|selection|selected)|not (the )?(entire|whole|full)|incomplete|based on (the )?(first|selected|portion)|portion)\b/i.test(e.text);
      return { met: says, detail: says ? "partial read disclosed" : `only ${partial.map((p) => p.label).join("; ")} was read, but the answer does not say so` };
    },
  };
}

export function structuredItems(id: string, description: string, mode: "commitment" | "assignment", min = 1): CompletionCriterion {
  return {
    id,
    description,
    required: true,
    evaluate(e) {
      const items = parseItems(e.text);
      if (items.length < min) return { met: false, detail: `found ${items.length} structured item(s) ([explicit]/[inferred])` };
      const thread = e.context.filter((c) => c.id === "outlook:thread").flatMap((c) => parseRenderedThread(c.content));
      if (!thread.length) return { met: false, detail: "the thread was not split into messages, so attribution cannot be checked" };
      const problems = checkItems(items, thread, mode).filter((c) => c.problem);
      return {
        met: problems.length === 0,
        detail: problems.length ? problems.map((p) => p.problem).slice(0, 3).join("; ") : `${items.length} item(s): ${items.filter((i) => i.kind === "explicit").length} explicit, ${items.filter((i) => i.kind === "inferred").length} inferred`,
      };
    },
  };
}

// --------------------------------------------------------- preservation

function lcsMatchedIndices(a: string[], b: string[]): Set<number> {
  const n = a.length;
  const m = b.length;
  if (n * m > 4_000_000) {
    // Large documents: fall back to positional comparison of unchanged prefix/suffix.
    const matched = new Set<number>();
    let i = 0;
    while (i < n && i < m && a[i] === b[i]) matched.add(i++);
    let j = 0;
    while (j < n - i && j < m - i && a[n - 1 - j] === b[m - 1 - j]) matched.add(n - 1 - j++);
    return matched;
  }
  const dp = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) dp[i]![j] = a[i] === b[j] ? dp[i + 1]![j + 1]! + 1 : Math.max(dp[i + 1]![j]!, dp[i]![j + 1]!);
  const matched = new Set<number>();
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      matched.add(i);
      i++;
      j++;
    } else if (dp[i + 1]![j]! >= dp[i]![j + 1]!) i++;
    else j++;
  }
  return matched;
}

/** Word: replay executed edits to find which original paragraphs were targeted. */
function wordAllowed(beforeLen: number, actions: ExecutionRecord[]): { allowed: Set<number>; unlocatable: number } {
  let orig: number[] = Array.from({ length: beforeLen }, (_, i) => i);
  const allowed = new Set<number>();
  let unlocatable = 0;
  const removeRange = (start: number, end: number, insert: number) => {
    const removed = orig.splice(start, end - start + 1, ...Array(insert).fill(-1));
    for (const r of removed) if (r >= 0) allowed.add(r);
  };
  const range = (s: unknown) => {
    const m = /^(\d+)–(\d+)$/.exec(String(s ?? ""));
    return m ? [Number(m[1]), Number(m[2])] : undefined;
  };
  for (const a of actions) {
    if (!EXECUTED.has(a.status)) continue;
    const args = a.args as Record<string, any>;
    const out = (a.output ?? {}) as Record<string, any>;
    if (a.capability === "document.replace_text") {
      if (args.target === "selection") unlocatable++;
      else removeRange(args.start, args.end ?? args.start, String(args.newText).split("\n").length);
    } else if (a.capability === "document.write_blocks") {
      if (args.mode === "replace") removeRange(args.start, args.end ?? args.start, (args.blocks as unknown[]).length);
      else {
        const r = range(out.paragraphs);
        if (r) orig.splice(r[0]!, 0, ...Array(r[1]! - r[0]! + 1).fill(-1));
        else unlocatable++;
      }
    } else if (a.capability === "document.insert_text") {
      const r = range(out.insertedParagraphs);
      if (r) orig.splice(r[0]!, 0, ...Array(r[1]! - r[0]! + 1).fill(-1));
      else unlocatable++;
    } else if (a.capability === "document.delete_paragraphs") removeRange(args.start, args.end, 0);
    else if (a.capability === "document.format_text") {
      for (let i = args.start; i <= (args.end ?? args.start); i++) if (orig[i]! >= 0) allowed.add(orig[i]!);
    }
  }
  return { allowed, unlocatable };
}

/** Keys of map snapshots a record may legitimately change. */
function touches(a: ExecutionRecord, key: string): boolean {
  const args = a.args as Record<string, any>;
  if (a.capability === "workbook.write_range") {
    const [sheet, cell] = [key.slice(0, key.lastIndexOf("!")), key.slice(key.lastIndexOf("!") + 1)];
    if (sheet !== args.sheet) return false;
    if (cell === "*") return true;
    const r = parseRange(String(args.address));
    const c = parseRange(cell);
    if (!r || !c) return false;
    return c.start.row >= r.start.row && c.start.row <= r.end.row && c.start.col >= r.start.col && c.start.col <= r.end.col;
  }
  if (a.capability === "workbook.clean_to_sheet" || a.capability === "workbook.add_worksheet") return key.startsWith(`${args.targetSheet ?? args.name}!`);
  if (["presentation.edit_text", "presentation.format_text", "presentation.set_geometry"].includes(a.capability)) return key === `${args.slideId}/${args.shapeId}`;
  if (a.capability === "presentation.delete_shapes") return (args.shapeIds as string[]).some((id) => key === `${args.slideId}/${id}`);
  if (a.capability === "presentation.delete_slide") return key.startsWith(`${args.slideId}/`);
  if (a.capability === "presentation.create_slide" || a.capability === "presentation.add_slide") {
    const id = (a.output as Record<string, unknown> | undefined)?.slideId;
    return typeof id === "string" && key.startsWith(`${id}/`);
  }
  if (a.capability === "outlook.create_draft") return key === "body";
  return false;
}

/** Content the task did not target is unchanged (compares host snapshots). */
export function preservesUnrelatedContent(): CompletionCriterion {
  return {
    id: "preserves-unrelated",
    description: "Content outside the targeted range is unchanged",
    required: true,
    evaluate(e) {
      if (!e.before || !e.after) return { met: false, detail: "host state could not be captured, so preservation was not verified" };
      if (e.before.kind === "sequence" && e.after.kind === "sequence") {
        const { allowed, unlocatable } = wordAllowed(e.before.units.length, e.actions);
        const matched = lcsMatchedIndices(e.before.units, e.after.units);
        const changed = e.before.units.map((_, i) => i).filter((i) => !matched.has(i) && !allowed.has(i));
        // Edits whose position is unknown (selection replacements) may account for one contiguous run each.
        const runs = changed.reduce((n, i, k) => (k === 0 || changed[k - 1] !== i - 1 ? n + 1 : n), 0);
        const met = changed.length === 0 || runs <= unlocatable;
        return { met, detail: met ? `${e.before.units.length - matched.size} targeted paragraph(s) changed; all others unchanged` : `untargeted paragraph(s) changed: ${changed.slice(0, 8).join(", ")}` };
      }
      if (e.before.kind === "map" && e.after.kind === "map") {
        const after = e.after.units;
        const before = e.before.units;
        const changed = Object.keys(before).filter((k) => after[k] !== before[k]);
        const bad = changed.filter((k) => !e.actions.some((a) => EXECUTED.has(a.status) && touches(a, k)));
        return { met: bad.length === 0, detail: bad.length ? `untargeted content changed or removed: ${bad.slice(0, 6).join(", ")}` : `${changed.length} targeted item(s) changed; all others unchanged` };
      }
      return { met: false, detail: "snapshots are not comparable" };
    },
  };
}
