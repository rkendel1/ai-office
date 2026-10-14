import type { Freshness } from "../context.js";
import { ScopeError, type CapabilityDefinition } from "../registry.js";
import type { ContextItem, ItemIdentity, Json, RiskClass } from "../types.js";
import { fingerprint } from "../util.js";
import { formatRange, parseRange, rangeSize, splitAddress } from "./a1.js";
import { allSupported, type ApiSupportEntry, type HostAdapter, type RequirementChecker, type StateSnapshot } from "./adapter.js";
import { makeItem } from "./common.js";
import { STAT_OPS, clean, computeStats, profile, toTable, type CleanOp, type StatOp } from "./excel-analysis.js";

export type CellValue = string | number | boolean | null;

export interface RangeData {
  /** Address without sheet, e.g. "A1:C3". */
  address: string;
  values: CellValue[][];
  formulas: CellValue[][];
}

export interface SheetInfo {
  name: string;
  visible: boolean;
  /** Used range address without sheet, or null if empty. */
  usedRange: string | null;
}

export interface TableInfo {
  name: string;
  sheet: string;
  address: string;
  headers: string[];
}

export type ChartKind = "column" | "bar" | "line" | "pie" | "scatter" | "area";
export const CHART_KINDS: ChartKind[] = ["column", "bar", "line", "pie", "scatter", "area"];

export interface ExcelDriver {
  workbookInfo(): Promise<{ id?: string; title?: string }>;
  worksheets(): Promise<SheetInfo[]>;
  activeSheet(): Promise<string>;
  selection(): Promise<{ sheet: string } & RangeData>;
  readRange(sheet: string, address: string): Promise<RangeData>;
  writeRange(sheet: string, address: string, data: { values?: CellValue[][]; formulas?: CellValue[][] }): Promise<void>;
  addWorksheet(name: string): Promise<void>;
  tables(): Promise<TableInfo[]>;
  createTable(sheet: string, address: string, hasHeaders: boolean, name?: string): Promise<string>;
  /** Apply a values filter on a table column. Returns the values now filtered on. */
  filterTable(table: string, column: string, values: string[]): Promise<void>;
  tableFilter(table: string, column: string): Promise<string[] | null>;
  charts(sheet: string): Promise<{ name: string; type: string; title?: string }[]>;
  createChart(p: { sheet: string; sourceSheet: string; sourceAddress: string; kind: ChartKind; title?: string; seriesBy: "auto" | "columns" | "rows" }): Promise<string>;
}

const MAX_READ_CELLS = 5_000;
const MAX_WRITE_CELLS = 10_000;
const SUBSTANTIAL_CELLS = 20;

function cellText(v: CellValue): string {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\t|\n/g, " ");
}

/** Tab-separated rendering with row numbers and column letters, formulas shown as "value [=formula]". */
export function renderRange(sheet: string, d: RangeData): string {
  const r = parseRange(d.address);
  if (!r) return "";
  const header = ["", ...Array.from({ length: r.end.col - r.start.col + 1 }, (_, i) => columnLetter(r.start.col + i))];
  const out = [header.join("\t")];
  d.values.forEach((row, ri) => {
    const cells = row.map((v, ci) => {
      const f = d.formulas[ri]?.[ci];
      return typeof f === "string" && f.startsWith("=") ? `${cellText(v)} [${f}]` : cellText(v);
    });
    out.push([String(r.start.row + ri + 1), ...cells].join("\t"));
  });
  return `Sheet ${sheet}, range ${d.address}\n` + out.join("\n");
}

function columnLetter(i: number) {
  return formatRange({ start: { row: 0, col: i }, end: { row: 0, col: i } }).replace(/\d+/g, "");
}

function sameValue(a: CellValue, b: CellValue): boolean {
  if (a === b) return true;
  if ((a === null || a === "") && (b === null || b === "")) return true;
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a));
  if (typeof a === "number" && typeof b === "string" && b.trim() !== "" && Number(b) === a) return true;
  if (typeof b === "number" && typeof a === "string" && a.trim() !== "" && Number(a) === b) return true;
  if (typeof a === "string" && typeof b === "string") return a.toUpperCase() === b.toUpperCase() && a.startsWith("=");
  return false;
}

function nonEmpty(m: CellValue[][]): number {
  return m.reduce((n, row) => n + row.filter((v) => v !== null && v !== "").length, 0);
}

export class ExcelAdapter implements HostAdapter {
  readonly host = "excel" as const;

  constructor(
    private readonly d: ExcelDriver,
    private readonly req: RequirementChecker = allSupported,
  ) {}

  async identity(): Promise<ItemIdentity> {
    const i = await this.d.workbookInfo();
    return { host: "excel", id: i.id, title: i.title };
  }

  apiSupport(): ApiSupportEntry[] {
    return [
      { requirement: "ExcelApi 1.1", supported: this.req.isSetSupported("ExcelApi", "1.1"), capabilities: ["read/write ranges", "tables", "charts"] },
      { requirement: "ExcelApi 1.2", supported: this.req.isSetSupported("ExcelApi", "1.2"), capabilities: ["workbook.filter_table"] },
    ];
  }

  private async sheetNames(): Promise<string[]> {
    return (await this.d.worksheets()).map((s) => s.name);
  }

  private async resolve(sheet: string | undefined, address: string): Promise<{ sheet: string; address: string; cells: number }> {
    const [prefix, addr] = splitAddress(address);
    const s = sheet ?? prefix ?? (await this.d.activeSheet());
    const names = await this.sheetNames();
    if (!names.includes(s)) throw new ScopeError(`Worksheet "${s}" does not exist. Sheets: ${names.join(", ")}`);
    const r = parseRange(addr);
    if (!r) throw new ScopeError(`"${addr}" is not a valid A1 range.`);
    return { sheet: s, address: formatRange(r), cells: rangeSize(r).cells };
  }

  private rangeItem(sheet: string, d: RangeData, source: string, truncated = false): ContextItem {
    return makeItem({
      host: "excel",
      id: `excel:range:${sheet}!${d.address}`,
      label: `${sheet}!${d.address}`,
      scope: "partial",
      content: renderRange(sheet, d),
      truncated,
      source,
    });
  }

  private async overviewItem(): Promise<ContextItem> {
    const sheets = await this.d.worksheets();
    const tables = await this.d.tables();
    const active = await this.d.activeSheet();
    const lines = [
      `Active sheet: ${active}`,
      "Sheets:",
      ...sheets.map((s) => `- ${s.name}${s.visible ? "" : " (hidden)"}: used range ${s.usedRange ?? "(empty)"}`),
    ];
    if (tables.length) lines.push("Tables:", ...tables.map((t) => `- ${t.name} on ${t.sheet}!${t.address} [${t.headers.join(", ")}]`));
    return makeItem({ host: "excel", id: "excel:meta", label: "Workbook overview", scope: "metadata", content: lines.join("\n"), source: "initial" });
  }

  async initialContext(): Promise<ContextItem[]> {
    const items = [await this.overviewItem()];
    const sel = await this.d.selection();
    const r = parseRange(sel.address);
    if (r && rangeSize(r).cells <= MAX_READ_CELLS && nonEmpty(sel.values) > 0) {
      const it = this.rangeItem(sel.sheet, sel, "selection.read");
      items.push({ ...it, id: "excel:selection", label: `Selection ${sel.sheet}!${sel.address}`, scope: "selection" });
    }
    return items;
  }

  async snapshot(): Promise<StateSnapshot> {
    const units: Record<string, string> = {};
    for (const sh of await this.d.worksheets()) {
      if (!sh.usedRange) continue;
      const r = parseRange(sh.usedRange)!;
      if (rangeSize(r).cells > 50_000) {
        const d = await this.d.readRange(sh.name, sh.usedRange);
        units[`${sh.name}!*`] = fingerprint(JSON.stringify(d.formulas));
        continue;
      }
      const d = await this.d.readRange(sh.name, sh.usedRange);
      d.formulas.forEach((row, i) =>
        row.forEach((v, j) => {
          if (v !== null && v !== "") units[`${sh.name}!${formatRange({ start: { row: r.start.row + i, col: r.start.col + j }, end: { row: r.start.row + i, col: r.start.col + j } })}`] = fingerprint(String(v));
        }),
      );
    }
    return { kind: "map", units };
  }

  /** Resolve a source range: explicit address, or the sheet's used range. */
  private async source(sheet: string, address?: string): Promise<{ sheet: string; address: string; startRow: number; startCol: number; values: CellValue[][] }> {
    let addr = address;
    if (!addr) {
      const used = (await this.d.worksheets()).find((s) => s.name === sheet)?.usedRange;
      if (!used) throw new ScopeError(`${sheet} has no data.`);
      addr = used;
    }
    const r = await this.resolve(sheet, addr);
    if (r.cells > 100_000) throw new ScopeError("Range is larger than 100,000 cells; choose a smaller range.");
    const data = await this.d.readRange(r.sheet, r.address);
    const pr = parseRange(r.address)!;
    return { sheet: r.sheet, address: r.address, startRow: pr.start.row + 1, startCol: pr.start.col, values: data.values };
  }

  async checkFreshness(item: ContextItem): Promise<Freshness> {
    const m = /^excel:range:(.+)!([A-Z]+\d+(?::[A-Z]+\d+)?)$/.exec(item.id);
    if (!m) return "unknown";
    try {
      const d = await this.d.readRange(m[1]!, m[2]!);
      return fingerprint(renderRange(m[1]!, d)) === item.fingerprint ? "fresh" : "stale";
    } catch {
      return "stale";
    }
  }

  capabilities(): CapabilityDefinition<any>[] {
    const d = this.d;
    const self = this;
    const caps: CapabilityDefinition<any>[] = [];

    caps.push({
      name: "selection.read",
      description: "Read the values and formulas of the user's selected cells.",
      hosts: ["excel"],
      risk: "read_selection",
      modifies: false,
      inputSchema: { type: "object", properties: {} },
      async execute() {
        const sel = await d.selection();
        const r = parseRange(sel.address);
        if (r && rangeSize(r).cells > MAX_READ_CELLS)
          return { output: { tooLarge: true, address: `${sel.sheet}!${sel.address}`, note: `Selection exceeds ${MAX_READ_CELLS} cells; read a sub-range with workbook.read_range.` } };
        const it = self.rangeItem(sel.sheet, sel, "selection.read");
        return { output: { item: it.id }, context: [it] };
      },
    });

    caps.push({
      name: "context.read",
      description: "List worksheets, used ranges, and tables in the workbook. Use workbook.read_range to read cell contents.",
      hosts: ["excel"],
      risk: "read_active",
      modifies: false,
      inputSchema: { type: "object", properties: {} },
      async execute() {
        const it = await self.overviewItem();
        return { output: { item: it.id }, context: [it] };
      },
    });

    caps.push({
      name: "workbook.read_range",
      description: `Read values and formulas of a range (A1 notation) on a worksheet. Max ${MAX_READ_CELLS} cells per call; omit address to read the sheet's used range (capped).`,
      hosts: ["excel"],
      risk: "read_active",
      modifies: false,
      inputSchema: {
        type: "object",
        required: ["sheet"],
        properties: {
          sheet: { type: "string", minLength: 1, maxLength: 31 },
          address: { type: "string", minLength: 2, maxLength: 32 },
        },
      },
      async resolveTarget(a: { sheet: string; address?: string }) {
        if (!a.address) {
          if (!(await self.sheetNames()).includes(a.sheet)) throw new ScopeError(`Worksheet "${a.sheet}" does not exist.`);
          return `${a.sheet} used range`;
        }
        const r = await self.resolve(a.sheet, a.address);
        return `${r.sheet}!${r.address}`;
      },
      async execute(a: { sheet: string; address?: string }) {
        let address = a.address;
        let truncated = false;
        if (!address) {
          const used = (await d.worksheets()).find((s) => s.name === a.sheet)?.usedRange;
          if (!used) return { output: { empty: true, note: `${a.sheet} has no used cells.` } };
          address = used;
        }
        const r = parseRange(splitAddress(address)[1])!;
        const size = rangeSize(r);
        if (size.cells > MAX_READ_CELLS) {
          const rows = Math.max(1, Math.floor(MAX_READ_CELLS / size.cols));
          r.end.row = r.start.row + rows - 1;
          if (size.cols > MAX_READ_CELLS) r.end.col = r.start.col + MAX_READ_CELLS - 1;
          truncated = true;
        }
        const data = await d.readRange(a.sheet, formatRange(r));
        const it = self.rangeItem(a.sheet, data, "workbook.read_range", truncated);
        return {
          output: { item: it.id, address: `${a.sheet}!${data.address}`, truncated, ...(truncated ? { fullRange: address } : {}) },
          context: [it],
        };
      },
    });

    type WriteArgs = { sheet: string; address: string; values?: CellValue[][]; formulas?: CellValue[][] };
    const matrix = { type: "array", maxItems: 2000, items: { type: "array", maxItems: 200, items: { anyOf: [{ type: "string", maxLength: 8192 }, { type: "number" }, { type: "boolean" }, { type: "null" }] } } } as const;
    caps.push({
      name: "workbook.write_range",
      description:
        "Write values or formulas (formulas start with '=') into a range. Provide exactly one of `values` or `formulas` as a 2-D array whose dimensions match the address.",
      hosts: ["excel"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["sheet", "address"],
        properties: {
          sheet: { type: "string", minLength: 1, maxLength: 31 },
          address: { type: "string", minLength: 2, maxLength: 32 },
          values: matrix as any,
          formulas: matrix as any,
        },
      },
      async resolveTarget(a: WriteArgs) {
        const data = a.values ?? a.formulas;
        if (!data || (a.values && a.formulas)) throw new ScopeError("Provide exactly one of `values` or `formulas`.");
        const r = await self.resolve(a.sheet, a.address);
        if (r.cells > MAX_WRITE_CELLS) throw new ScopeError(`At most ${MAX_WRITE_CELLS} cells can be written in one action.`);
        const pr = parseRange(r.address)!;
        const { rows, cols } = rangeSize(pr);
        if (data.length !== rows || data.some((row) => row.length !== cols))
          throw new ScopeError(`Data is ${data.length}×${data[0]?.length ?? 0} but ${r.address} is ${rows}×${cols}.`);
        return `${r.sheet}!${r.address}`;
      },
      async assessRisk(a: WriteArgs): Promise<RiskClass> {
        const r = await self.resolve(a.sheet, a.address);
        const existing = await d.readRange(r.sheet, r.address);
        return nonEmpty(existing.formulas) > SUBSTANTIAL_CELLS ? "destructive" : "modify";
      },
      async preview(a: WriteArgs) {
        const r = await self.resolve(a.sheet, a.address);
        const existing = await d.readRange(r.sheet, r.address);
        const occupied = nonEmpty(existing.formulas);
        const data = (a.values ?? a.formulas)!;
        return {
          summary: `Write ${a.formulas ? "formulas" : "values"} to ${r.sheet}!${r.address}`,
          target: `${r.sheet}!${r.address}`,
          before: occupied ? renderRange(r.sheet, existing).slice(0, 3000) : "(empty cells)",
          after: renderRange(r.sheet, { address: r.address, values: data, formulas: a.formulas ?? data }).slice(0, 3000),
          warnings: occupied ? [`Overwrites ${occupied} non-empty cell(s).`] : [],
        };
      },
      async execute(a: WriteArgs) {
        const r = await self.resolve(a.sheet, a.address);
        const before = await d.readRange(r.sheet, r.address);
        await d.writeRange(r.sheet, r.address, { values: a.values, formulas: a.formulas });
        const after = await d.readRange(r.sheet, r.address);
        const expected = (a.formulas ?? a.values)!;
        const actual = a.formulas ? after.formulas : after.values;
        let mismatches = 0;
        let firstBad = "";
        expected.forEach((row, i) =>
          row.forEach((v, j) => {
            const got = actual[i]?.[j] ?? null;
            const ok = a.formulas && typeof v === "string" && v.startsWith("=") ? sameValue(v, got) : sameValue(v, got) || sameValue(v, after.values[i]?.[j] ?? null);
            if (!ok) {
              mismatches++;
              if (!firstBad) firstBad = `row ${i + 1}, column ${j + 1}: expected ${JSON.stringify(v)}, found ${JSON.stringify(got)}`;
            }
          }),
        );
        const errors = after.values.flat().filter((v) => typeof v === "string" && /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!)/.test(v));
        const it = self.rangeItem(r.sheet, after, "workbook.write_range");
        return {
          output: {
            written: `${r.sheet}!${r.address}`,
            resultPreview: renderRange(r.sheet, after).slice(0, 1500),
            ...(errors.length ? { formulaErrors: errors.length } : {}),
          },
          context: [it],
          verification: {
            verified: mismatches === 0,
            method: "re-read range values and formulas",
            evidence: errors.length ? `${errors.length} cell(s) evaluate to an Excel error` : undefined,
            mismatch: mismatches ? `${mismatches} cell(s) differ (${firstBad}).` : undefined,
          },
          compensation: {
            description: `Restore previous contents of ${r.sheet}!${r.address}`,
            capability: "workbook.write_range",
            args: { sheet: r.sheet, address: r.address, formulas: before.formulas as Json },
          },
        };
      },
    });

    caps.push({
      name: "workbook.add_worksheet",
      description: "Add a new, empty worksheet with the given name.",
      hosts: ["excel"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["name"],
        properties: { name: { type: "string", minLength: 1, maxLength: 31, pattern: "^[^\\\\/?*\\[\\]:]+$" } },
      },
      async resolveTarget(a: { name: string }) {
        if ((await self.sheetNames()).some((n) => n.toLowerCase() === a.name.toLowerCase()))
          throw new ScopeError(`A worksheet named "${a.name}" already exists.`);
        return `new worksheet "${a.name}"`;
      },
      async preview(a: { name: string }) {
        return { summary: `Add worksheet "${a.name}"`, target: "Workbook" };
      },
      async execute(a: { name: string }) {
        await d.addWorksheet(a.name);
        const ok = (await self.sheetNames()).includes(a.name);
        return { output: { sheet: a.name }, verification: { verified: ok, method: "list worksheets" } };
      },
    });

    type TableArgs = { sheet: string; address: string; hasHeaders: boolean; name?: string };
    caps.push({
      name: "workbook.create_table",
      description: "Convert a range into an Excel table (with filter buttons).",
      hosts: ["excel"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["sheet", "address", "hasHeaders"],
        properties: {
          sheet: { type: "string", minLength: 1, maxLength: 31 },
          address: { type: "string", minLength: 2, maxLength: 32 },
          hasHeaders: { type: "boolean" },
          name: { type: "string", pattern: "^[A-Za-z_][A-Za-z0-9_.]{0,254}$" },
        },
      },
      async resolveTarget(a: TableArgs) {
        const r = await self.resolve(a.sheet, a.address);
        const tables = await d.tables();
        if (a.name && tables.some((t) => t.name.toLowerCase() === a.name!.toLowerCase())) throw new ScopeError(`Table "${a.name}" already exists.`);
        if (tables.some((t) => t.sheet === r.sheet && overlaps(t.address, r.address)))
          throw new ScopeError("The range overlaps an existing table.");
        return `${r.sheet}!${r.address}`;
      },
      async preview(a: TableArgs) {
        return { summary: `Create table${a.name ? ` "${a.name}"` : ""}${a.hasHeaders ? " with headers" : ""}`, target: `${a.sheet}!${a.address}` };
      },
      async execute(a: TableArgs) {
        const r = await self.resolve(a.sheet, a.address);
        const name = await d.createTable(r.sheet, r.address, a.hasHeaders, a.name);
        const t = (await d.tables()).find((x) => x.name === name);
        return {
          output: { table: name, headers: t?.headers ?? [] },
          verification: { verified: !!t && t.sheet === r.sheet, method: "list tables", evidence: t ? `${t.name} at ${t.sheet}!${t.address}` : undefined },
        };
      },
    });

    if (this.req.isSetSupported("ExcelApi", "1.2")) {
      type FilterArgs = { table: string; column: string; values: string[] };
      caps.push({
        name: "workbook.filter_table",
        description: "Filter a table column to show only rows whose value is one of `values`.",
        hosts: ["excel"],
        risk: "modify",
        modifies: true,
        inputSchema: {
          type: "object",
          required: ["table", "column", "values"],
          properties: {
            table: { type: "string", minLength: 1 },
            column: { type: "string", minLength: 1 },
            values: { type: "array", minItems: 1, maxItems: 500, items: { type: "string" } },
          },
        },
        async resolveTarget(a: FilterArgs) {
          const t = (await d.tables()).find((x) => x.name === a.table);
          if (!t) throw new ScopeError(`Table "${a.table}" does not exist.`);
          if (!t.headers.includes(a.column)) throw new ScopeError(`Table "${a.table}" has no column "${a.column}". Columns: ${t.headers.join(", ")}`);
          return `table ${a.table}, column ${a.column}`;
        },
        async preview(a: FilterArgs) {
          return { summary: `Filter ${a.table}[${a.column}] to ${a.values.length} value(s)`, target: a.table, after: a.values.join(", ") };
        },
        async execute(a: FilterArgs) {
          await d.filterTable(a.table, a.column, a.values);
          const got = await d.tableFilter(a.table, a.column);
          const ok = got !== null && got.length === a.values.length && a.values.every((v) => got.includes(v));
          return { output: { filtered: a.values.length }, verification: { verified: got === null ? null : ok, method: "read column filter criteria" } };
        },
      });
    }

    type ChartArgs = { sourceSheet: string; sourceAddress: string; kind: ChartKind; title?: string; targetSheet?: string; seriesBy?: "auto" | "columns" | "rows" };
    caps.push({
      name: "workbook.create_chart",
      description: "Create a chart from a source range. Put it on `targetSheet` (default: the source sheet).",
      hosts: ["excel"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["sourceSheet", "sourceAddress", "kind"],
        properties: {
          sourceSheet: { type: "string", minLength: 1, maxLength: 31 },
          sourceAddress: { type: "string", minLength: 2, maxLength: 32 },
          kind: { type: "string", enum: CHART_KINDS as unknown as Json[] },
          title: { type: "string", maxLength: 255 },
          targetSheet: { type: "string", minLength: 1, maxLength: 31 },
          seriesBy: { type: "string", enum: ["auto", "columns", "rows"] },
        },
      },
      async resolveTarget(a: ChartArgs) {
        const r = await self.resolve(a.sourceSheet, a.sourceAddress);
        const target = a.targetSheet ?? r.sheet;
        if (!(await self.sheetNames()).includes(target)) throw new ScopeError(`Worksheet "${target}" does not exist.`);
        return `${a.kind} chart of ${r.sheet}!${r.address} on ${target}`;
      },
      async preview(a: ChartArgs) {
        return {
          summary: `Create a ${a.kind} chart${a.title ? ` titled “${a.title}”` : ""}`,
          target: `Data ${a.sourceSheet}!${a.sourceAddress} → sheet ${a.targetSheet ?? a.sourceSheet}`,
        };
      },
      async execute(a: ChartArgs) {
        const r = await self.resolve(a.sourceSheet, a.sourceAddress);
        const target = a.targetSheet ?? r.sheet;
        const before = await d.charts(target);
        const name = await d.createChart({ sheet: target, sourceSheet: r.sheet, sourceAddress: r.address, kind: a.kind, title: a.title, seriesBy: a.seriesBy ?? "auto" });
        const after = await d.charts(target);
        const c = after.find((x) => x.name === name);
        const ok = !!c && after.length === before.length + 1 && (!a.title || c.title === a.title);
        return { output: { chart: name, sheet: target }, verification: { verified: ok, method: "list charts on sheet", evidence: c ? `${c.name} (${c.type})` : undefined } };
      },
    });

    const sourceProps = {
      sheet: { type: "string", minLength: 1, maxLength: 31 },
      address: { type: "string", minLength: 2, maxLength: 32, description: "Range including the header row; omit for the used range" },
      hasHeaders: { type: "boolean", description: "First row contains column names (default true)" },
    } as const;

    type StatsArgs = { sheet: string; address?: string; hasHeaders?: boolean; measures: { column: string; ops: StatOp[] }[]; groupBy?: string };
    caps.push({
      name: "workbook.compute_stats",
      description:
        "Calculate statistics (count, sum, mean, median, min, max, stdev, distinct, missing) for columns of a range, optionally grouped by a column. Calculated by the application, not estimated: use these results when reporting numbers.",
      hosts: ["excel"],
      risk: "read_active",
      modifies: false,
      inputSchema: {
        type: "object",
        required: ["sheet", "measures"],
        properties: {
          ...sourceProps,
          measures: {
            type: "array",
            minItems: 1,
            maxItems: 20,
            items: {
              type: "object",
              required: ["column", "ops"],
              properties: { column: { type: "string", minLength: 1 }, ops: { type: "array", minItems: 1, items: { type: "string", enum: STAT_OPS as unknown as Json[] } } },
            },
          },
          groupBy: { type: "string", minLength: 1 },
        },
      } as any,
      async resolveTarget(a: StatsArgs) {
        const src = await self.source(a.sheet, a.address);
        return `statistics for ${src.sheet}!${src.address}`;
      },
      async execute(a: StatsArgs) {
        const src = await self.source(a.sheet, a.address);
        const t = toTable(src.values, a.hasHeaders !== false, src.startRow);
        const res = computeStats(t, a.measures, src.startCol, a.groupBy);
        const lines = [res.columns.join("\t"), ...res.rows.map((r) => r.map((v) => (v === null ? "" : String(v))).join("\t"))];
        const nonNum = Object.entries(res.nonNumericCells).filter(([, n]) => n > 0);
        const item = makeItem({
          host: "excel",
          id: `excel:stats:${fingerprint(JSON.stringify(a))}`,
          label: `Computed statistics for ${src.sheet}!${src.address}`,
          scope: "partial",
          content: `Source: ${src.sheet}!${src.address} (${t.rows.length} data rows)\n${lines.join("\n")}${nonNum.length ? `\nNon-numeric cells ignored: ${nonNum.map(([k, n]) => `${k}: ${n}`).join(", ")}` : ""}`,
          source: "workbook.compute_stats",
        });
        return {
          output: { source: `${src.sheet}!${src.address}`, dataRows: t.rows.length, columns: res.columns, rows: res.rows as Json, nonNumericCells: res.nonNumericCells as Json },
          context: [item],
        };
      },
    });

    type ProfileArgs = { sheet: string; address?: string; hasHeaders?: boolean };
    caps.push({
      name: "workbook.profile_data",
      description:
        "Profile data quality of a range: per-column type, missing values, distinct values, numbers stored as text, IQR outliers (with row numbers), inconsistent category spellings, and duplicate rows.",
      hosts: ["excel"],
      risk: "read_active",
      modifies: false,
      inputSchema: { type: "object", required: ["sheet"], properties: { ...sourceProps } } as any,
      async resolveTarget(a: ProfileArgs) {
        const src = await self.source(a.sheet, a.address);
        return `data profile of ${src.sheet}!${src.address}`;
      },
      async execute(a: ProfileArgs) {
        const src = await self.source(a.sheet, a.address);
        const p = profile(toTable(src.values, a.hasHeaders !== false, src.startRow));
        const lines = [`Source: ${src.sheet}!${src.address}, ${p.rows} data rows`];
        for (const c of p.columns) {
          const issues = [
            c.missing && `${c.missing} missing`,
            c.numbersStoredAsText && `${c.numbersStoredAsText} numbers stored as text`,
            c.outliers.length && `outliers at rows ${c.outliers.map((o) => `${o.row} (${o.value})`).join(", ")}`,
            c.inconsistentCategories.length && `inconsistent spellings: ${c.inconsistentCategories.map((g) => g.map((x) => JSON.stringify(x)).join(" / ")).join("; ")}`,
            c.type === "mixed" && "mixed types",
          ].filter(Boolean);
          lines.push(`- ${c.name} [${c.type}, ${c.distinct} distinct]${issues.length ? ": " + issues.join("; ") : ""}`);
        }
        lines.push(p.duplicateRows.length ? `Duplicate rows: ${p.duplicateRows.map((d) => `${d.row} = ${d.duplicateOf}`).join(", ")}` : "Duplicate rows: none");
        const item = makeItem({ host: "excel", id: `excel:profile:${src.sheet}!${src.address}`, label: `Data profile of ${src.sheet}!${src.address}`, scope: "partial", content: lines.join("\n"), source: "workbook.profile_data" });
        return { output: { source: `${src.sheet}!${src.address}`, profile: p as unknown as Json }, context: [item] };
      },
    });

    type CleanArgs = { sheet: string; address?: string; hasHeaders?: boolean; operations: CleanOp[]; targetSheet: string };
    caps.push({
      name: "workbook.clean_to_sheet",
      description:
        "Write a cleaned copy of a range to a NEW worksheet, leaving the source untouched. Operations run in order: trim, dedupe {columns?}, drop_missing {columns}, fill_missing {column, value}, normalize_case {column, case}, canonicalize_categories {column}, text_to_number {column}. Row counts and the unchanged source are verified.",
      hosts: ["excel"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["sheet", "operations", "targetSheet"],
        properties: {
          ...sourceProps,
          targetSheet: { type: "string", minLength: 1, maxLength: 31, pattern: "^[^\\\\/?*\\[\\]:]+$" },
          operations: {
            type: "array",
            minItems: 1,
            maxItems: 30,
            items: {
              type: "object",
              required: ["op"],
              properties: {
                op: { type: "string", enum: ["trim", "dedupe", "drop_missing", "fill_missing", "normalize_case", "canonicalize_categories", "text_to_number"] },
                column: { type: "string" },
                columns: { type: "array", items: { type: "string" } },
                value: { anyOf: [{ type: "string" }, { type: "number" }] },
                case: { type: "string", enum: ["lower", "upper", "title"] },
              },
            },
          },
        },
      } as any,
      async resolveTarget(a: CleanArgs) {
        if ((await self.sheetNames()).some((n) => n.toLowerCase() === a.targetSheet.toLowerCase()))
          throw new ScopeError(`Worksheet "${a.targetSheet}" already exists; cleaned data must go to a new sheet.`);
        const src = await self.source(a.sheet, a.address);
        for (const op of a.operations) {
          if (["drop_missing"].includes(op.op) && !(op as any).columns?.length) throw new ScopeError("drop_missing needs `columns`.");
          if (["fill_missing", "normalize_case", "canonicalize_categories", "text_to_number"].includes(op.op) && !(op as any).column)
            throw new ScopeError(`${op.op} needs \`column\`.`);
          if (op.op === "fill_missing" && (op as any).value === undefined) throw new ScopeError("fill_missing needs `value`.");
          if (op.op === "normalize_case" && !(op as any).case) throw new ScopeError("normalize_case needs `case`.");
        }
        return `cleaned copy of ${src.sheet}!${src.address} → new sheet "${a.targetSheet}"`;
      },
      async preview(a: CleanArgs) {
        const src = await self.source(a.sheet, a.address);
        const { report } = clean(toTable(src.values, a.hasHeaders !== false, src.startRow), a.operations, src.startCol);
        return {
          summary: `Create sheet "${a.targetSheet}" with a cleaned copy (${report.inputRows} → ${report.outputRows} rows)`,
          target: `${src.sheet}!${src.address} (unchanged)`,
          after: report.steps.join("\n"),
        };
      },
      async execute(a: CleanArgs) {
        const src = await self.source(a.sheet, a.address);
        const sourceFp = fingerprint(JSON.stringify((await d.readRange(src.sheet, src.address)).formulas));
        const t = toTable(src.values, a.hasHeaders !== false, src.startRow);
        const { table, report } = clean(t, a.operations, src.startCol);
        const out: CellValue[][] = [...(a.hasHeaders !== false ? [table.headers] : []), ...table.rows];
        await d.addWorksheet(a.targetSheet);
        let written = "";
        if (out.length && out[0]!.length) {
          written = formatRange({ start: { row: 0, col: 0 }, end: { row: out.length - 1, col: out[0]!.length - 1 } });
          // Write as values: cleaned text must not be re-interpreted as formulas.
          await d.writeRange(a.targetSheet, written, { values: out.map((r) => r.map((v) => (typeof v === "string" && v.startsWith("=") ? `'${v}` : v))) });
        }
        const back = written ? await d.readRange(a.targetSheet, written) : { values: [] as CellValue[][] };
        const dataRowsBack = back.values.length - (a.hasHeaders !== false ? 1 : 0);
        let mismatched = 0;
        out.forEach((r, i) => r.forEach((v, j) => {
          const got = back.values[i]?.[j] ?? null;
          const ok = v === got || ((v === null || v === "") && (got === null || got === "")) || (typeof v === "string" && v.startsWith("=") && got === v);
          if (!ok) mismatched++;
        }));
        const sourceUnchanged = fingerprint(JSON.stringify((await d.readRange(src.sheet, src.address)).formulas)) === sourceFp;
        const invariant = report.outputRows === report.inputRows - report.removedDuplicates - report.droppedMissing;
        const ok = mismatched === 0 && sourceUnchanged && invariant && dataRowsBack === report.outputRows;
        return {
          output: { targetSheet: a.targetSheet, written: written ? `${a.targetSheet}!${written}` : null, report: report as unknown as Json, sourceUnchanged },
          verification: {
            verified: ok,
            method: "re-read cleaned sheet; check row-count invariant and that the source range is unchanged",
            evidence: `${report.inputRows} rows in → ${report.outputRows} out (−${report.removedDuplicates} duplicates, −${report.droppedMissing} incomplete); source ${sourceUnchanged ? "unchanged" : "CHANGED"}`,
            mismatch: ok ? undefined : [!sourceUnchanged && "source range changed", !invariant && "row counts do not add up", mismatched && `${mismatched} cell(s) differ`, dataRowsBack !== report.outputRows && "row count on the new sheet differs"].filter(Boolean).join("; "),
          },
        };
      },
    });

    return caps;
  }
}

function overlaps(a: string, b: string): boolean {
  const x = parseRange(a);
  const y = parseRange(b);
  if (!x || !y) return false;
  return !(x.end.row < y.start.row || y.end.row < x.start.row || x.end.col < y.start.col || y.end.col < x.start.col);
}
