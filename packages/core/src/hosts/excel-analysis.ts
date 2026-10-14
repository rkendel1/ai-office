/**
 * Deterministic data analysis for Excel ranges. Runs inside the trusted
 * application, never inside the model: results produced here are the
 * authoritative numbers the assistant may report.
 */
import { colToIndex } from "./a1.js";
import type { CellValue } from "./excel.js";

export interface Table {
  headers: string[];
  rows: CellValue[][];
  /** Worksheet row number of rows[0] (1-based). */
  firstRow: number;
}

export function toTable(values: CellValue[][], hasHeaders: boolean, firstRow: number): Table {
  const width = Math.max(0, ...values.map((r) => r.length));
  const pad = (r: CellValue[]) => Array.from({ length: width }, (_, i) => r[i] ?? null);
  if (hasHeaders) {
    const [h = [], ...rest] = values;
    return { headers: pad(h).map((x, i) => (x === null || x === "" ? `Column${i + 1}` : String(x))), rows: rest.map(pad), firstRow: firstRow + 1 };
  }
  return { headers: Array.from({ length: width }, (_, i) => `Column${i + 1}`), rows: values.map(pad), firstRow };
}

/** Resolve a column by header name (case-insensitive) or by column letter within the range. */
export function columnIndex(t: Table, column: string, rangeStartCol: number): number {
  const byName = t.headers.findIndex((h) => h.toLowerCase() === column.toLowerCase());
  if (byName >= 0) return byName;
  if (/^[A-Z]{1,3}$/i.test(column)) {
    const i = colToIndex(column) - rangeStartCol;
    if (i >= 0 && i < t.headers.length) return i;
  }
  throw new Error(`Column "${column}" not found. Columns: ${t.headers.join(", ")}`);
}

const isEmpty = (v: CellValue) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");
const round = (n: number) => (Number.isFinite(n) ? Math.round(n * 1e6) / 1e6 : n);

export type StatOp = "count" | "sum" | "mean" | "median" | "min" | "max" | "stdev" | "distinct" | "missing";
export const STAT_OPS: StatOp[] = ["count", "sum", "mean", "median", "min", "max", "stdev", "distinct", "missing"];

export function stat(values: CellValue[], op: StatOp): number | null {
  const nums = values.filter((v): v is number => typeof v === "number" && Number.isFinite(v));
  switch (op) {
    case "count":
      return nums.length;
    case "missing":
      return values.filter(isEmpty).length;
    case "distinct":
      return new Set(values.filter((v) => !isEmpty(v)).map((v) => (typeof v === "string" ? v.trim() : v))).size;
    case "sum":
      return round(nums.reduce((a, b) => a + b, 0));
    case "mean":
      return nums.length ? round(nums.reduce((a, b) => a + b, 0) / nums.length) : null;
    case "min":
      return nums.length ? Math.min(...nums) : null;
    case "max":
      return nums.length ? Math.max(...nums) : null;
    case "median": {
      if (!nums.length) return null;
      const s = [...nums].sort((a, b) => a - b);
      const m = Math.floor(s.length / 2);
      return round(s.length % 2 ? s[m]! : (s[m - 1]! + s[m]!) / 2);
    }
    case "stdev": {
      if (nums.length < 2) return null;
      const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
      return round(Math.sqrt(nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1)));
    }
  }
}

export interface StatsResult {
  groupBy?: string;
  columns: string[];
  rows: (string | number | null)[][];
  nonNumericCells: Record<string, number>;
}

export function computeStats(t: Table, measures: { column: string; ops: StatOp[] }[], startCol: number, groupBy?: string): StatsResult {
  const cols = measures.map((m) => ({ ...m, idx: columnIndex(t, m.column, startCol) }));
  const nonNumericCells: Record<string, number> = {};
  for (const c of cols)
    nonNumericCells[t.headers[c.idx]!] = t.rows.filter((r) => !isEmpty(r[c.idx]!) && typeof r[c.idx] !== "number").length;
  const header = cols.flatMap((c) => c.ops.map((op) => `${op}(${t.headers[c.idx]})`));
  const calc = (rows: CellValue[][]) => cols.flatMap((c) => c.ops.map((op) => stat(rows.map((r) => r[c.idx] ?? null), op)));
  if (!groupBy) return { columns: header, rows: [calc(t.rows)], nonNumericCells };
  const g = columnIndex(t, groupBy, startCol);
  const groups = new Map<string, CellValue[][]>();
  for (const r of t.rows) {
    const key = isEmpty(r[g]!) ? "(blank)" : String(r[g]).trim();
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(r);
  }
  return {
    groupBy: t.headers[g],
    columns: [t.headers[g]!, ...header],
    rows: [...groups.entries()].map(([k, rows]) => [k, ...calc(rows)]),
    nonNumericCells,
  };
}

// --------------------------------------------------------------- profiling

export interface ColumnProfile {
  name: string;
  type: "number" | "text" | "boolean" | "mixed" | "empty";
  missing: number;
  distinct: number;
  numbersStoredAsText: number;
  /** IQR outliers: worksheet row and value. */
  outliers: { row: number; value: number }[];
  /** Groups of spellings that differ only in case/whitespace. */
  inconsistentCategories: string[][];
}

export interface DataProfile {
  rows: number;
  columns: ColumnProfile[];
  /** Worksheet rows that duplicate an earlier row (all columns equal). */
  duplicateRows: { row: number; duplicateOf: number }[];
}

const NUMERIC_TEXT = /^\s*[-+]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?\s*$/;

export function profile(t: Table): DataProfile {
  const columns: ColumnProfile[] = t.headers.map((name, i) => {
    const vals = t.rows.map((r) => r[i] ?? null);
    const present = vals.filter((v) => !isEmpty(v));
    const kinds = new Set(present.map((v) => typeof v));
    const type = !present.length ? "empty" : kinds.size > 1 ? "mixed" : kinds.has("number") ? "number" : kinds.has("boolean") ? "boolean" : "text";
    const nums = vals.map((v, r) => ({ v, row: t.firstRow + r })).filter((x): x is { v: number; row: number } => typeof x.v === "number");
    let outliers: { row: number; value: number }[] = [];
    if (nums.length >= 5) {
      const s = nums.map((x) => x.v).sort((a, b) => a - b);
      const q = (p: number) => {
        const idx = (s.length - 1) * p;
        const lo = Math.floor(idx);
        return s[lo]! + (s[Math.ceil(idx)]! - s[lo]!) * (idx - lo);
      };
      const q1 = q(0.25);
      const q3 = q(0.75);
      const iqr = q3 - q1;
      outliers = nums.filter((x) => x.v < q1 - 1.5 * iqr || x.v > q3 + 1.5 * iqr).map((x) => ({ row: x.row, value: x.v }));
    }
    const variants = new Map<string, Set<string>>();
    for (const v of present)
      if (typeof v === "string" && !NUMERIC_TEXT.test(v)) {
        const k = v.trim().replace(/\s+/g, " ").toLowerCase();
        if (!variants.has(k)) variants.set(k, new Set());
        variants.get(k)!.add(v);
      }
    return {
      name,
      type,
      missing: vals.length - present.length,
      distinct: new Set(present.map((v) => (typeof v === "string" ? v.trim() : v))).size,
      numbersStoredAsText: present.filter((v) => typeof v === "string" && NUMERIC_TEXT.test(v)).length,
      outliers: outliers.slice(0, 50),
      inconsistentCategories: [...variants.values()].filter((s) => s.size > 1).map((s) => [...s]),
    };
  });
  const seen = new Map<string, number>();
  const duplicateRows: { row: number; duplicateOf: number }[] = [];
  t.rows.forEach((r, i) => {
    if (r.every(isEmpty)) return;
    const key = JSON.stringify(r.map((v) => (typeof v === "string" ? v.trim() : v)));
    const row = t.firstRow + i;
    if (seen.has(key)) duplicateRows.push({ row, duplicateOf: seen.get(key)! });
    else seen.set(key, row);
  });
  return { rows: t.rows.length, columns, duplicateRows };
}

// ---------------------------------------------------------------- cleaning

export type CleanOp =
  | { op: "trim" }
  | { op: "dedupe"; columns?: string[] }
  | { op: "drop_missing"; columns: string[] }
  | { op: "fill_missing"; column: string; value: string | number }
  | { op: "normalize_case"; column: string; case: "lower" | "upper" | "title" }
  | { op: "canonicalize_categories"; column: string }
  | { op: "text_to_number"; column: string };

export interface CleanReport {
  inputRows: number;
  outputRows: number;
  removedDuplicates: number;
  droppedMissing: number;
  cellsChanged: number;
  steps: string[];
}

const title = (s: string) => s.toLowerCase().replace(/\b\p{L}/gu, (c) => c.toUpperCase());

/** Apply cleaning operations to a copy of the table. The input is never mutated. */
export function clean(input: Table, ops: CleanOp[], startCol: number): { table: Table; report: CleanReport } {
  let rows = input.rows.map((r) => [...r]);
  const t = { ...input, rows };
  const report: CleanReport = { inputRows: input.rows.length, outputRows: 0, removedDuplicates: 0, droppedMissing: 0, cellsChanged: 0, steps: [] };
  const set = (r: CellValue[], i: number, v: CellValue) => {
    if (r[i] !== v) {
      r[i] = v;
      report.cellsChanged++;
    }
  };
  for (const op of ops) {
    switch (op.op) {
      case "trim": {
        const before = report.cellsChanged;
        for (const r of rows) r.forEach((v, i) => typeof v === "string" && set(r, i, v.trim().replace(/\s+/g, " ")));
        report.steps.push(`Trimmed whitespace in ${report.cellsChanged - before} cell(s)`);
        break;
      }
      case "normalize_case": {
        const c = columnIndex(t, op.column, startCol);
        const f = op.case === "lower" ? (s: string) => s.toLowerCase() : op.case === "upper" ? (s: string) => s.toUpperCase() : title;
        const before = report.cellsChanged;
        for (const r of rows) if (typeof r[c] === "string") set(r, c, f(r[c] as string));
        report.steps.push(`Normalized case in ${t.headers[c]} (${report.cellsChanged - before} cell(s))`);
        break;
      }
      case "canonicalize_categories": {
        const c = columnIndex(t, op.column, startCol);
        const counts = new Map<string, Map<string, number>>();
        for (const r of rows) {
          const v = r[c];
          if (typeof v !== "string" || !v.trim()) continue;
          const k = v.trim().replace(/\s+/g, " ").toLowerCase();
          if (!counts.has(k)) counts.set(k, new Map());
          const m = counts.get(k)!;
          m.set(v, (m.get(v) ?? 0) + 1);
        }
        // Most frequent spelling wins; ties prefer a capitalised form, then alphabetical order.
        const capital = (s: string) => (/^\s*\p{Lu}/u.test(s) ? 0 : 1);
        const canonical = new Map([...counts].map(([k, m]) => [k, [...m.entries()].sort((a, b) => b[1] - a[1] || capital(a[0]) - capital(b[0]) || a[0].localeCompare(b[0]))[0]![0].trim()]));
        const before = report.cellsChanged;
        for (const r of rows) {
          const v = r[c];
          if (typeof v === "string" && v.trim()) set(r, c, canonical.get(v.trim().replace(/\s+/g, " ").toLowerCase())!);
        }
        report.steps.push(`Unified category spellings in ${t.headers[c]} (${report.cellsChanged - before} cell(s))`);
        break;
      }
      case "text_to_number": {
        const c = columnIndex(t, op.column, startCol);
        const before = report.cellsChanged;
        for (const r of rows) {
          const v = r[c];
          if (typeof v === "string" && NUMERIC_TEXT.test(v)) set(r, c, Number(v.replace(/,/g, "")));
        }
        report.steps.push(`Converted ${report.cellsChanged - before} numeric text value(s) in ${t.headers[c]}`);
        break;
      }
      case "fill_missing": {
        const c = columnIndex(t, op.column, startCol);
        const before = report.cellsChanged;
        for (const r of rows) if (isEmpty(r[c]!)) set(r, c, op.value);
        report.steps.push(`Filled ${report.cellsChanged - before} missing value(s) in ${t.headers[c]} with ${JSON.stringify(op.value)}`);
        break;
      }
      case "drop_missing": {
        const cs = op.columns.map((x) => columnIndex(t, x, startCol));
        const n = rows.length;
        rows = rows.filter((r) => !cs.some((c) => isEmpty(r[c]!)));
        report.droppedMissing += n - rows.length;
        report.steps.push(`Removed ${n - rows.length} row(s) missing ${cs.map((c) => t.headers[c]).join(", ")}`);
        break;
      }
      case "dedupe": {
        const cs = op.columns?.length ? op.columns.map((x) => columnIndex(t, x, startCol)) : t.headers.map((_, i) => i);
        const seen = new Set<string>();
        const n = rows.length;
        rows = rows.filter((r) => {
          const k = JSON.stringify(cs.map((c) => (typeof r[c] === "string" ? (r[c] as string).trim().toLowerCase() : r[c])));
          if (seen.has(k)) return false;
          seen.add(k);
          return true;
        });
        report.removedDuplicates += n - rows.length;
        report.steps.push(`Removed ${n - rows.length} duplicate row(s)`);
        break;
      }
    }
    t.rows = rows;
  }
  report.outputRows = rows.length;
  return { table: { headers: input.headers, rows, firstRow: input.firstRow }, report };
}
