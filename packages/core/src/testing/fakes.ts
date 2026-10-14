/**
 * In-memory Office hosts and a scripted model. Used by automated tests and by
 * the add-in's "demo mode" when running outside Office. They implement the
 * same driver interfaces as the Office.js drivers.
 */
import { formatRange, parseRange, rangeSize } from "../hosts/a1.js";
import type { CellValue, ChartKind, ExcelDriver, RangeData, SheetInfo, TableInfo } from "../hosts/excel.js";
import type { MailItem, OutlookDriver } from "../hosts/outlook.js";
import type { Box, PowerPointDriver, PptShape, PptSlide, ShapeKind, TextFormat } from "../hosts/powerpoint.js";
import type { WordDriver, WordParagraph, WordStyle } from "../hosts/word.js";
import type { ChatRequest, ChatResponse, InferenceProvider, ModelDetails, ModelInfo, ToolCall } from "../types.js";
import { newId } from "../util.js";

// --------------------------------- Word -----------------------------------

export class FakeWordDriver implements WordDriver {
  paragraphs: WordParagraph[];
  selectionRange: { start: number; end: number } | null = null;
  /** Set to make the next write silently not happen (tests verification). */
  dropNextWrite = false;
  unsupported: { kind: string; count: number }[] = [];

  constructor(
    paras: (string | WordParagraph)[],
    public title = "Test.docx",
  ) {
    this.paragraphs = paras.map((p) => (typeof p === "string" ? { text: p, style: "Normal", bold: false, italic: false } : p));
  }
  async documentInfo() {
    return { id: "file:///" + this.title, title: this.title };
  }
  async selection() {
    if (!this.selectionRange) return { text: "" };
    const { start, end } = this.selectionRange;
    return {
      text: this.paragraphs
        .slice(start, end + 1)
        .map((p) => p.text)
        .join("\n"),
      paragraphStart: start,
      paragraphEnd: end,
    };
  }
  async paragraphCount() {
    return this.paragraphs.length;
  }
  async readParagraphs(start: number, end: number) {
    return this.paragraphs.slice(start, end).map((p) => ({ ...p }));
  }
  async unsupportedContent() {
    return this.unsupported;
  }
  async search(q: string, matchCase: boolean, max: number) {
    const needle = matchCase ? q : q.toLowerCase();
    return this.paragraphs
      .map((p, i) => ({ paragraphIndex: i, text: p.text }))
      .filter((p) => (matchCase ? p.text : p.text.toLowerCase()).includes(needle))
      .slice(0, max);
  }
  private drop(): boolean {
    if (this.dropNextWrite) {
      this.dropNextWrite = false;
      return true;
    }
    return false;
  }
  async replaceParagraphs(start: number, end: number, text: string) {
    if (this.drop()) return;
    const style = this.paragraphs[start]?.style ?? "Normal";
    const lines = text.split("\n").map((t) => ({ text: t, style, bold: false, italic: false }));
    this.paragraphs.splice(start, end - start + 1, ...lines);
  }
  async replaceSelection(text: string) {
    if (!this.selectionRange) throw new Error("No selection");
    const { start, end } = this.selectionRange;
    if (!this.drop()) await this.replaceParagraphs(start, end, text);
    const n = text.split("\n").length;
    this.selectionRange = { start, end: start + n - 1 };
    return (await this.selection()).text;
  }
  async insertParagraphs(at: "start" | "end" | { before: number } | { after: number }, lines: string[], style?: WordStyle) {
    const idx = at === "start" ? 0 : at === "end" ? this.paragraphs.length : "before" in at ? at.before : at.after + 1;
    if (this.drop()) return idx;
    this.paragraphs.splice(idx, 0, ...lines.map((t) => ({ text: t, style: style ?? "Normal", bold: false, italic: false })));
    return idx;
  }
  async deleteParagraphs(start: number, end: number) {
    this.paragraphs.splice(start, end - start + 1);
  }
  async formatParagraphs(start: number, end: number, fmt: { style?: WordStyle; bold?: boolean; italic?: boolean }) {
    if (this.drop()) return;
    for (let i = start; i <= end; i++) {
      const p = this.paragraphs[i]!;
      if (fmt.style) p.style = fmt.style;
      if (fmt.bold !== undefined) p.bold = fmt.bold;
      if (fmt.italic !== undefined) p.italic = fmt.italic;
    }
  }
}

// --------------------------------- Excel ----------------------------------

type Sheet = { name: string; cells: Map<string, CellValue> };

export class FakeExcelDriver implements ExcelDriver {
  sheets: Sheet[] = [];
  tablesList: TableInfo[] = [];
  chartsList: { sheet: string; name: string; type: string; title?: string }[] = [];
  filters = new Map<string, string[]>();
  active = "";
  selectionAddr = { sheet: "", address: "A1" };

  constructor(data: Record<string, CellValue[][]>, public title = "Book.xlsx") {
    for (const [name, rows] of Object.entries(data)) {
      const s: Sheet = { name, cells: new Map() };
      rows.forEach((row, r) => row.forEach((v, c) => v !== null && v !== "" && s.cells.set(`${r},${c}`, v)));
      this.sheets.push(s);
    }
    this.active = this.sheets[0]?.name ?? "Sheet1";
    this.selectionAddr = { sheet: this.active, address: "A1" };
  }
  private sheet(name: string) {
    const s = this.sheets.find((x) => x.name === name);
    if (!s) throw new Error(`No sheet ${name}`);
    return s;
  }
  /**
   * Evaluate the formula subset used in tests: =SUM|AVERAGE|COUNT|MAX|MIN|MEDIAN(range)
   * and plain cell references, each optionally prefixed with a sheet name.
   */
  private evaluate(s: Sheet, v: CellValue, depth = 0): CellValue {
    if (typeof v === "string" && v.startsWith("'")) return v.slice(1);
    if (typeof v !== "string" || !v.startsWith("=")) return v;
    if (depth > 20) return "#REF!";
    const SHEET = "(?:'([^']+)'|([A-Za-z0-9_]+))!";
    const fn = new RegExp(`^=(SUM|AVERAGE|COUNT|MAX|MIN|MEDIAN)\\((?:${SHEET})?(\\$?[A-Z]+\\$?\\d+(?::\\$?[A-Z]+\\$?\\d+)?)\\)$`, "i").exec(v);
    const ref = new RegExp(`^=(?:${SHEET})?(\\$?[A-Z]+\\$?\\d+)$`, "i").exec(v);
    const target = (quoted?: string, plain?: string) => {
      const name = quoted ?? plain;
      return name ? this.sheets.find((x) => x.name === name) : s;
    };
    if (fn) {
      const sh = target(fn[2], fn[3]);
      if (!sh) return "#REF!";
      const r = parseRange(fn[4]!.replace(/\$/g, ""))!;
      const nums: number[] = [];
      for (let i = r.start.row; i <= r.end.row; i++)
        for (let j = r.start.col; j <= r.end.col; j++) {
          const x = this.evaluate(sh, sh.cells.get(`${i},${j}`) ?? null, depth + 1);
          if (typeof x === "number") nums.push(x);
        }
      const f = fn[1]!.toUpperCase();
      if (f === "SUM") return nums.reduce((a, b) => a + b, 0);
      if (f === "COUNT") return nums.length;
      if (f === "MAX") return nums.length ? Math.max(...nums) : 0;
      if (f === "MIN") return nums.length ? Math.min(...nums) : 0;
      if (f === "MEDIAN") {
        if (!nums.length) return "#NUM!";
        const so = [...nums].sort((a, b) => a - b);
        const m = Math.floor(so.length / 2);
        return so.length % 2 ? so[m]! : (so[m - 1]! + so[m]!) / 2;
      }
      return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : "#DIV/0!";
    }
    if (ref) {
      const sh = target(ref[1], ref[2]);
      if (!sh) return "#REF!";
      const c = parseRange(ref[3]!.replace(/\$/g, ""))!;
      return this.evaluate(sh, sh.cells.get(`${c.start.row},${c.start.col}`) ?? null, depth + 1) ?? 0;
    }
    return "#NAME?";
  }
  private used(s: Sheet): string | null {
    if (!s.cells.size) return null;
    let r0 = Infinity, c0 = Infinity, r1 = -1, c1 = -1;
    for (const k of s.cells.keys()) {
      const [r, c] = k.split(",").map(Number) as [number, number];
      r0 = Math.min(r0, r); c0 = Math.min(c0, c); r1 = Math.max(r1, r); c1 = Math.max(c1, c);
    }
    return formatRange({ start: { row: r0, col: c0 }, end: { row: r1, col: c1 } });
  }
  async workbookInfo() {
    return { id: "file:///" + this.title, title: this.title };
  }
  async worksheets(): Promise<SheetInfo[]> {
    return this.sheets.map((s) => ({ name: s.name, visible: true, usedRange: this.used(s) }));
  }
  async activeSheet() {
    return this.active;
  }
  async selection() {
    return { sheet: this.selectionAddr.sheet, ...(await this.readRange(this.selectionAddr.sheet, this.selectionAddr.address)) };
  }
  async readRange(sheet: string, address: string): Promise<RangeData> {
    const s = this.sheet(sheet);
    const r = parseRange(address)!;
    const values: CellValue[][] = [];
    const formulas: CellValue[][] = [];
    for (let i = r.start.row; i <= r.end.row; i++) {
      const vr: CellValue[] = [], fr: CellValue[] = [];
      for (let j = r.start.col; j <= r.end.col; j++) {
        const raw = s.cells.get(`${i},${j}`) ?? "";
        fr.push(raw);
        vr.push(this.evaluate(s, raw));
      }
      values.push(vr);
      formulas.push(fr);
    }
    return { address: formatRange(r), values, formulas };
  }
  async writeRange(sheet: string, address: string, data: { values?: CellValue[][]; formulas?: CellValue[][] }) {
    const s = this.sheet(sheet);
    const r = parseRange(address)!;
    const m = (data.formulas ?? data.values)!;
    const { rows, cols } = rangeSize(r);
    for (let i = 0; i < rows; i++)
      for (let j = 0; j < cols; j++) {
        const v = m[i]?.[j] ?? null;
        const key = `${r.start.row + i},${r.start.col + j}`;
        if (v === null || v === "") s.cells.delete(key);
        else s.cells.set(key, v);
      }
  }
  async addWorksheet(name: string) {
    this.sheets.push({ name, cells: new Map() });
  }
  async tables() {
    return this.tablesList.map((t) => ({ ...t }));
  }
  async createTable(sheet: string, address: string, hasHeaders: boolean, name?: string) {
    const n = name ?? `Table${this.tablesList.length + 1}`;
    const data = await this.readRange(sheet, address);
    const headers = hasHeaders ? (data.values[0] ?? []).map(String) : (data.values[0] ?? []).map((_, i) => `Column${i + 1}`);
    this.tablesList.push({ name: n, sheet, address: data.address, headers });
    return n;
  }
  async filterTable(table: string, column: string, values: string[]) {
    this.filters.set(`${table}/${column}`, [...values]);
  }
  async tableFilter(table: string, column: string) {
    return this.filters.get(`${table}/${column}`) ?? [];
  }
  async charts(sheet: string) {
    return this.chartsList.filter((c) => c.sheet === sheet).map(({ name, type, title }) => ({ name, type, title }));
  }
  async createChart(p: { sheet: string; sourceSheet: string; sourceAddress: string; kind: ChartKind; title?: string }) {
    const name = `Chart ${this.chartsList.length + 1}`;
    this.chartsList.push({ sheet: p.sheet, name, type: p.kind, title: p.title });
    return name;
  }
}

// ------------------------------- PowerPoint -------------------------------

export interface FakeShapeInit {
  name: string;
  text: string | null;
  type?: string;
  box?: Box;
  format?: TextFormat;
}

/**
 * In-memory presentation. Shapes carry geometry and formatting like the
 * Office.js driver reports them. Layouts mimic the default Office theme's
 * placeholder names. Notes and tables can be switched off to test the
 * "not supported" paths.
 */
export class FakePowerPointDriver implements PowerPointDriver {
  slidesList: PptSlide[];
  notes = new Map<string, string>();
  tables = new Map<string, string[][]>();
  selected: string[] | null = [];
  notesSupported = true;
  tablesSupported = true;
  size = { width: 960, height: 540 };
  private next = 1000;
  readonly layoutList = [
    { id: "L1", name: "Title Slide", masterId: "M1" },
    { id: "L2", name: "Title and Content", masterId: "M1" },
    { id: "L3", name: "Title Only", masterId: "M1" },
    { id: "L4", name: "Two Content", masterId: "M1" },
    { id: "L5", name: "Blank", masterId: "M1" },
  ];

  constructor(slides: { shapes: FakeShapeInit[]; notes?: string }[], public title = "Deck.pptx") {
    this.slidesList = slides.map((s, i) => {
      const id = `s${i + 1}`;
      if (s.notes) this.notes.set(id, s.notes);
      return {
        id,
        index: i,
        shapes: s.shapes.map((sh, j) => ({
          id: `s${i + 1}sh${j + 1}`,
          name: sh.name,
          type: sh.type ?? (sh.text === null ? "Image" : "TextBox"),
          text: sh.text,
          box: sh.box ?? { left: 48, top: 32 + j * 110, width: 864, height: j === 0 ? 80 : 360 },
          format: sh.format ?? (sh.text === null ? undefined : { size: j === 0 ? 32 : 20, bold: j === 0, italic: false, align: "left" as const, color: "#1B1B1F" }),
        })),
      };
    });
  }
  private slide(id: string) {
    const s = this.slidesList.find((x) => x.id === id);
    if (!s) throw new Error("slide not found");
    return s;
  }
  private shape(slideId: string, shapeId: string) {
    const sh = this.slide(slideId).shapes.find((x) => x.id === shapeId);
    if (!sh) throw new Error("shape not found");
    return sh;
  }
  async info() {
    return { id: "file:///" + this.title, title: this.title };
  }
  async slides() {
    return this.slidesList.map((s, i) => ({ ...s, index: i, shapes: s.shapes.map((x) => ({ ...x, box: x.box && { ...x.box }, format: x.format && { ...x.format } })) }));
  }
  async selectedSlideIds() {
    return this.selected;
  }
  async setShapeText(slideId: string, shapeId: string, text: string) {
    this.shape(slideId, shapeId).text = text;
  }
  async addSlide(layout?: { layoutId: string; masterId: string }) {
    const id = `slide${this.next++}`;
    const name = this.layoutList.find((l) => l.id === layout?.layoutId)?.name ?? "Blank";
    const ph = (n: string, box: Box, size: number) => ({ id: `${id}ph${n.replace(/\W/g, "")}`, name: n, type: "Placeholder", text: "", box, format: { size, bold: false, italic: false, align: "left" as const, color: "#1B1B1F" } });
    const shapes: PptShape[] = [];
    if (name === "Title Slide") shapes.push(ph("Title 1", { left: 120, top: 160, width: 720, height: 110 }, 44), ph("Subtitle 2", { left: 120, top: 290, width: 720, height: 70 }, 24));
    if (name === "Title and Content") shapes.push(ph("Title 1", { left: 48, top: 30, width: 864, height: 90 }, 36), ph("Content Placeholder 2", { left: 48, top: 135, width: 864, height: 360 }, 24));
    if (name === "Title Only") shapes.push(ph("Title 1", { left: 48, top: 30, width: 864, height: 90 }, 36));
    if (name === "Two Content")
      shapes.push(ph("Title 1", { left: 48, top: 30, width: 864, height: 90 }, 36), ph("Content Placeholder 2", { left: 48, top: 135, width: 420, height: 360 }, 24), ph("Content Placeholder 3", { left: 492, top: 135, width: 420, height: 360 }, 24));
    this.slidesList.push({ id, index: this.slidesList.length, shapes });
    return id;
  }
  async addTextBox(slideId: string, text: string, box: Box, format?: TextFormat) {
    const s = this.slide(slideId);
    const id = `${slideId}sh${this.next++}`;
    s.shapes.push({ id, name: `TextBox ${s.shapes.length + 1}`, type: "TextBox", text, box: { ...box }, format: { size: 18, bold: false, italic: false, align: "left", color: "#000000", ...format } });
    return id;
  }
  async addShape(slideId: string, kind: ShapeKind, box: Box, opts: { fill?: string; line?: string; text?: string; format?: TextFormat }) {
    const s = this.slide(slideId);
    const id = `${slideId}sh${this.next++}`;
    s.shapes.push({ id, name: `${kind} ${s.shapes.length + 1}`, type: "GeometricShape", text: opts.text ?? "", box: { ...box }, fill: opts.fill, format: { size: 18, bold: false, italic: false, align: "center", color: "#FFFFFF", ...opts.format } });
    return id;
  }
  async formatText(slideId: string, shapeId: string, format: TextFormat) {
    const sh = this.shape(slideId, shapeId);
    sh.format = { ...sh.format, ...format };
  }
  async setGeometry(slideId: string, shapeId: string, box: Partial<Box>) {
    const sh = this.shape(slideId, shapeId);
    sh.box = { ...sh.box!, ...Object.fromEntries(Object.entries(box).filter(([, v]) => v !== undefined)) };
  }
  async deleteShape(slideId: string, shapeId: string) {
    const s = this.slide(slideId);
    s.shapes = s.shapes.filter((x) => x.id !== shapeId);
  }
  async deleteSlide(slideId: string) {
    this.slidesList = this.slidesList.filter((s) => s.id !== slideId);
  }
  async slideSize() {
    return { ...this.size, reported: true };
  }
  async layouts() {
    return this.layoutList.map((l) => ({ ...l }));
  }
  supportsTables() {
    return this.tablesSupported;
  }
  async addTable(slideId: string, rows: string[][], box: Box) {
    const s = this.slide(slideId);
    const id = `${slideId}tbl${this.next++}`;
    this.tables.set(id, rows.map((r) => [...r]));
    s.shapes.push({ id, name: `Table ${s.shapes.length + 1}`, type: "Table", text: rows.map((r) => r.join("\t")).join("\n"), box: { ...box } });
    return id;
  }
  supportsNotes() {
    return this.notesSupported;
  }
  async getNotes(slideId: string) {
    return this.notes.get(slideId) ?? null;
  }
  async setNotes(slideId: string, text: string) {
    this.notes.set(slideId, text);
  }
}

// --------------------------------- Outlook --------------------------------

export class FakeOutlookDriver implements OutlookDriver {
  composeText = "";
  openedForms: { body: string; replyAll: boolean }[] = [];
  /** Outlook never sends from this driver; tests assert this stays empty. */
  readonly sent: never[] = [];

  constructor(
    public message: MailItem,
    public currentMode: "read" | "compose" = "read",
  ) {}
  mode() {
    return this.currentMode;
  }
  async item() {
    return { ...this.message };
  }
  async composeBody() {
    return this.composeText;
  }
  async insertComposeText(text: string) {
    this.composeText = text + this.composeText;
  }
  async displayReplyForm(body: string, replyAll: boolean) {
    this.openedForms.push({ body, replyAll });
  }
}

// ------------------------------ Scripted model ----------------------------

export type ScriptStep =
  | { text: string }
  | { call: { name: string; arguments: Record<string, any> }; text?: string }
  | ((req: ChatRequest) => ChatResponse | Promise<ChatResponse>);

/**
 * Deterministic stand-in for an Ollama model. Each `chat` call consumes the
 * next step. Records every request for assertions.
 */
export class ScriptedProvider implements InferenceProvider {
  readonly id = "scripted";
  readonly requests: ChatRequest[] = [];
  private i = 0;

  constructor(
    private readonly steps: ScriptStep[],
    private readonly details: Partial<ModelDetails> = {},
  ) {}

  async ping() {
    return { ok: true, version: "scripted" };
  }
  async listModels(): Promise<ModelInfo[]> {
    return [{ name: "scripted:latest", family: "test" }];
  }
  async describeModel(name: string): Promise<ModelDetails> {
    return { name, capabilities: ["completion", "tools"], contextLength: 32768, ...this.details };
  }
  async chat(req: ChatRequest): Promise<ChatResponse> {
    this.requests.push({ ...req, messages: req.messages.map((m) => ({ ...m })) });
    if (req.signal?.aborted) throw Object.assign(new Error("aborted"), { name: "AbortError" });
    const step = this.steps[this.i++];
    if (!step) return { message: { role: "assistant", content: "(script exhausted)" } };
    if (typeof step === "function") return step(req);
    if ("call" in step) {
      const tc: ToolCall = { id: newId("call"), name: step.call.name, arguments: step.call.arguments };
      return { message: { role: "assistant", content: step.text ?? "", toolCalls: [tc] } };
    }
    req.onDelta?.(step.text);
    return { message: { role: "assistant", content: step.text } };
  }
}
