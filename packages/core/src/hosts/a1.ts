/** A1-notation helpers for Excel ranges (single rectangular areas only). */

export interface CellRef {
  row: number; // 0-based
  col: number; // 0-based
}

export interface RangeRef {
  start: CellRef;
  end: CellRef;
}

export function colToIndex(col: string): number {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function indexToCol(i: number): string {
  let s = "";
  let n = i + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

const CELL = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;

export function parseCell(s: string): CellRef | undefined {
  const m = CELL.exec(s.trim());
  if (!m) return undefined;
  const col = colToIndex(m[1]!);
  const row = Number(m[2]) - 1;
  if (row < 0 || row >= 1_048_576 || col < 0 || col >= 16_384) return undefined;
  return { row, col };
}

/** Parse "A1" or "A1:C10" (no sheet prefix). */
export function parseRange(s: string): RangeRef | undefined {
  const parts = s.trim().split(":");
  if (parts.length < 1 || parts.length > 2) return undefined;
  const a = parseCell(parts[0]!);
  const b = parts.length === 2 ? parseCell(parts[1]!) : a;
  if (!a || !b) return undefined;
  return {
    start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
    end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
  };
}

export function formatRange(r: RangeRef): string {
  const a = `${indexToCol(r.start.col)}${r.start.row + 1}`;
  const b = `${indexToCol(r.end.col)}${r.end.row + 1}`;
  return a === b ? a : `${a}:${b}`;
}

export function rangeSize(r: RangeRef): { rows: number; cols: number; cells: number } {
  const rows = r.end.row - r.start.row + 1;
  const cols = r.end.col - r.start.col + 1;
  return { rows, cols, cells: rows * cols };
}

/** Strip "Sheet1!" or "'My Sheet'!" prefixes, returning [sheet?, address]. */
export function splitAddress(full: string): [string | undefined, string] {
  const i = full.lastIndexOf("!");
  if (i < 0) return [undefined, full];
  let sheet = full.slice(0, i);
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return [sheet, full.slice(i + 1)];
}
