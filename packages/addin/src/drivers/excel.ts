import { splitAddress, type CellValue, type ChartKind, type ExcelDriver, type RangeData, type SheetInfo, type TableInfo } from "@loai/core";
import { supports } from "./requirements";

const CHART_TYPES: Record<ChartKind, string> = {
  column: "ColumnClustered",
  bar: "BarClustered",
  line: "Line",
  pie: "Pie",
  scatter: "XYScatter",
  area: "Area",
};

const addr = (full: string) => splitAddress(full)[1];

/** Excel driver over Office.js (ExcelApi 1.1+; 1.2 for filters, 1.4 for null-safe used range). */
export class OfficeExcelDriver implements ExcelDriver {
  async workbookInfo() {
    return Excel.run(async (ctx) => {
      ctx.workbook.load("name");
      await ctx.sync();
      return { id: Office.context.document?.url || undefined, title: ctx.workbook.name || undefined };
    });
  }

  async worksheets(): Promise<SheetInfo[]> {
    return Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name,items/visibility");
      await ctx.sync();
      const used = sheets.items.map((s) => {
        const r = supports("ExcelApi", "1.4") ? s.getUsedRangeOrNullObject(true) : s.getUsedRange(true);
        r.load("address");
        return r;
      });
      await ctx.sync();
      return sheets.items.map((s, i) => {
        const r = used[i] as Excel.Range & { isNullObject?: boolean };
        return { name: s.name, visible: s.visibility === "Visible", usedRange: r.isNullObject ? null : addr(r.address) };
      });
    });
  }

  async activeSheet() {
    return Excel.run(async (ctx) => {
      const s = ctx.workbook.worksheets.getActiveWorksheet();
      s.load("name");
      await ctx.sync();
      return s.name;
    });
  }

  async selection() {
    return Excel.run(async (ctx) => {
      const r = ctx.workbook.getSelectedRange();
      r.load("address,cellCount");
      await ctx.sync();
      const [sheet, a] = splitAddress(r.address);
      if (r.cellCount > 5000) return { sheet: sheet!, address: a, values: [], formulas: [] };
      r.load("values,formulas");
      await ctx.sync();
      return { sheet: sheet!, address: a, values: r.values as CellValue[][], formulas: r.formulas as CellValue[][] };
    });
  }

  async readRange(sheet: string, address: string): Promise<RangeData> {
    return Excel.run(async (ctx) => {
      const r = ctx.workbook.worksheets.getItem(sheet).getRange(address);
      r.load("address,values,formulas");
      await ctx.sync();
      return { address: addr(r.address), values: r.values as CellValue[][], formulas: r.formulas as CellValue[][] };
    });
  }

  async writeRange(sheet: string, address: string, data: { values?: CellValue[][]; formulas?: CellValue[][] }) {
    await Excel.run(async (ctx) => {
      const r = ctx.workbook.worksheets.getItem(sheet).getRange(address);
      if (data.formulas) r.formulas = data.formulas as any;
      else r.values = data.values as any;
      await ctx.sync();
    });
  }

  async addWorksheet(name: string) {
    await Excel.run(async (ctx) => {
      ctx.workbook.worksheets.add(name);
      await ctx.sync();
    });
  }

  async tables(): Promise<TableInfo[]> {
    return Excel.run(async (ctx) => {
      const ts = ctx.workbook.tables;
      ts.load("items/name");
      await ctx.sync();
      const parts = ts.items.map((t) => {
        const ws = t.worksheet;
        ws.load("name");
        const r = t.getRange();
        r.load("address");
        const h = t.getHeaderRowRange();
        h.load("values");
        return { t, ws, r, h };
      });
      await ctx.sync();
      return parts.map((p) => ({ name: p.t.name, sheet: p.ws.name, address: addr(p.r.address), headers: (p.h.values[0] ?? []).map(String) }));
    });
  }

  async createTable(sheet: string, address: string, hasHeaders: boolean, name?: string) {
    return Excel.run(async (ctx) => {
      const t = ctx.workbook.worksheets.getItem(sheet).tables.add(address, hasHeaders);
      if (name) t.name = name;
      t.load("name");
      await ctx.sync();
      return t.name;
    });
  }

  async filterTable(table: string, column: string, values: string[]) {
    await Excel.run(async (ctx) => {
      ctx.workbook.tables.getItem(table).columns.getItem(column).filter.applyValuesFilter(values);
      await ctx.sync();
    });
  }

  async tableFilter(table: string, column: string) {
    if (!supports("ExcelApi", "1.2")) return null;
    return Excel.run(async (ctx) => {
      const f = ctx.workbook.tables.getItem(table).columns.getItem(column).filter;
      f.load("criteria");
      await ctx.sync();
      const v = f.criteria?.values;
      return Array.isArray(v) ? v.map(String) : [];
    });
  }

  async charts(sheet: string) {
    return Excel.run(async (ctx) => {
      const cs = ctx.workbook.worksheets.getItem(sheet).charts;
      cs.load("items/name,items/chartType,items/title/text");
      await ctx.sync();
      return cs.items.map((c) => ({ name: c.name, type: String(c.chartType), title: c.title?.text }));
    });
  }

  async createChart(p: { sheet: string; sourceSheet: string; sourceAddress: string; kind: ChartKind; title?: string; seriesBy: "auto" | "columns" | "rows" }) {
    return Excel.run(async (ctx) => {
      const src = ctx.workbook.worksheets.getItem(p.sourceSheet).getRange(p.sourceAddress);
      const seriesBy = p.seriesBy === "columns" ? "Columns" : p.seriesBy === "rows" ? "Rows" : "Auto";
      const chart = ctx.workbook.worksheets.getItem(p.sheet).charts.add(CHART_TYPES[p.kind] as any, src, seriesBy as any);
      if (p.title) chart.title.text = p.title;
      chart.load("name");
      await ctx.sync();
      return chart.name;
    });
  }
}
