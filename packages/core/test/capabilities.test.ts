/**
 * Tests for capabilities added for specialists: Word structured writing,
 * Excel statistics/profiling/cleaning, PowerPoint construction and layout
 * checks, Outlook thread splitting, and the evidence helpers.
 */
import { describe, expect, it } from "vitest";
import {
  ExcelAdapter,
  FakeExcelDriver,
  FakePowerPointDriver,
  FakeWordDriver,
  PowerPointAdapter,
  ScriptedProvider,
  WordAdapter,
  barChartGeometry,
  checkLayout,
  clean,
  computeStats,
  numbersInText,
  parseItems,
  profile,
  splitThread,
  toTable,
  ungroundedNumbers,
  type ToolCall,
} from "../src/index.js";
import { setup } from "./helpers.js";

const tc = (name: string, args: Record<string, unknown>): ToolCall => ({ id: "t", name, arguments: args as any });

describe("document.write_blocks", () => {
  it("inserts styled paragraphs, verifies them, and undoes", async () => {
    const doc = new FakeWordDriver(["Keep me."]);
    const s = setup(new WordAdapter(doc), new ScriptedProvider([]));
    const r = await s.runtime.executeDirect(tc("document.write_blocks", { mode: "insert", position: "start", blocks: [{ text: "Title", style: "Title" }, { text: "Body" }] }));
    expect(r.status).toBe("succeeded");
    expect(doc.paragraphs.map((p) => `${p.style}:${p.text}`)).toEqual(["Title:Title", "Normal:Body", "Normal:Keep me."]);
    const undo = await s.runtime.executeDirect(tc(r.compensation!.capability, r.compensation!.args));
    expect(undo.status).toBe("succeeded");
    expect(doc.paragraphs.map((p) => p.text)).toEqual(["Keep me."]);
  });

  it("replaces paragraphs only when the expected text matches, and can restore them", async () => {
    const doc = new FakeWordDriver(["A", "old one", "old two", "Z"]);
    const s = setup(new WordAdapter(doc), new ScriptedProvider([]));
    const stale = await s.runtime.executeDirect(tc("document.write_blocks", { mode: "replace", start: 1, end: 2, expectedText: "something else", blocks: [{ text: "x" }] }));
    expect(stale.status).toBe("failed");
    const r = await s.runtime.executeDirect(tc("document.write_blocks", { mode: "replace", start: 1, end: 2, expectedText: "old one\nold two", blocks: [{ text: "New", style: "Heading1" }, { text: "a" }, { text: "b" }] }));
    expect(r.status).toBe("succeeded");
    expect(doc.paragraphs.map((p) => p.text)).toEqual(["A", "New", "a", "b", "Z"]);
    await s.runtime.executeDirect(tc(r.compensation!.capability, r.compensation!.args));
    expect(doc.paragraphs.map((p) => p.text)).toEqual(["A", "old one", "old two", "Z"]);
  });

  it("rejects multi-line blocks and unknown styles", async () => {
    const s = setup(new WordAdapter(new FakeWordDriver(["a"])), new ScriptedProvider([]));
    expect((await s.runtime.executeDirect(tc("document.write_blocks", { mode: "insert", blocks: [{ text: "a\nb" }] }))).error).toMatch(/pattern/);
    expect((await s.runtime.executeDirect(tc("document.write_blocks", { mode: "insert", blocks: [{ text: "a", style: "Fancy" }] }))).error).toMatch(/must be one of/);
  });
});

describe("Excel analysis", () => {
  const values = [
    ["Region", "Sales", "Units"],
    ["North", 100, 3],
    ["north ", 110, 4],
    ["South", 90, ""],
    ["South", 90, ""],
    ["East", 1000, 5],
    ["West", 95, "7"],
  ];

  it("computes grouped statistics deterministically", () => {
    const t = toTable(values as any, true, 1);
    const r = computeStats(t, [{ column: "Sales", ops: ["sum", "mean", "median", "count", "min", "max"] }], 0, "Region");
    expect(r.columns).toEqual(["Region", "sum(Sales)", "mean(Sales)", "median(Sales)", "count(Sales)", "min(Sales)", "max(Sales)"]);
    expect(r.rows.find((x) => x[0] === "South")).toEqual(["South", 180, 90, 90, 2, 90, 90]);
    const all = computeStats(t, [{ column: "Units", ops: ["sum", "missing"] }], 0);
    expect(all.rows[0]).toEqual([12, 2]);
    expect(all.nonNumericCells.Units).toBe(1); // "7" stored as text is not silently counted
  });

  it("profiles duplicates, outliers, missing values, numbers stored as text and inconsistent categories", () => {
    const p = profile(toTable(values as any, true, 1));
    expect(p.duplicateRows).toEqual([{ row: 5, duplicateOf: 4 }]); // worksheet rows (header is row 1)
    const sales = p.columns.find((c) => c.name === "Sales")!;
    expect(sales.outliers).toEqual([{ row: 6, value: 1000 }]);
    const units = p.columns.find((c) => c.name === "Units")!;
    expect(units.missing).toBe(2);
    expect(units.numbersStoredAsText).toBe(1);
    expect(p.columns[0]!.inconsistentCategories).toEqual([["North", "north "]]);
  });

  it("cleans a copy with row-count invariants and never mutates the input", () => {
    const t = toTable(values as any, true, 1);
    const snapshot = JSON.stringify(t);
    const { table, report } = clean(t, [{ op: "trim" }, { op: "canonicalize_categories", column: "Region" }, { op: "drop_missing", columns: ["Units"] }, { op: "text_to_number", column: "Units" }, { op: "dedupe" }], 0);
    expect(JSON.stringify(t)).toBe(snapshot);
    expect(report).toMatchObject({ inputRows: 6, droppedMissing: 2, removedDuplicates: 0, outputRows: 4 });
    expect(table.rows.map((r) => r[0])).toEqual(["North", "North", "East", "West"]);
    expect(table.rows[3]![2]).toBe(7);
  });

  it("workbook.compute_stats and clean_to_sheet run through the registry", async () => {
    const xl = new FakeExcelDriver({ Data: values as any });
    const s = setup(new ExcelAdapter(xl), new ScriptedProvider([]));
    const stats = await s.runtime.executeDirect(tc("workbook.compute_stats", { sheet: "Data", measures: [{ column: "Sales", ops: ["sum"] }] }));
    expect((stats.output as any).rows).toEqual([[1485]]);
    expect(s.approver.requests).toHaveLength(0); // read-only
    const bad = await s.runtime.executeDirect(tc("workbook.clean_to_sheet", { sheet: "Data", targetSheet: "Data", operations: [{ op: "trim" }] }));
    expect(bad.error).toMatch(/already exists/);
    const ok = await s.runtime.executeDirect(tc("workbook.clean_to_sheet", { sheet: "Data", targetSheet: "Clean", operations: [{ op: "dedupe" }] }));
    expect(ok.status).toBe("succeeded");
    expect(ok.verification!.evidence).toMatch(/6 rows in → 5 out/);
  });
});

describe("PowerPoint construction", () => {
  const deck = () => new FakePowerPointDriver([{ shapes: [{ name: "Title 1", text: "Hello" }] }]);

  it("create_slide fills layout placeholders, removes unused ones, and adds notes", async () => {
    const ppt = deck();
    const s = setup(new PowerPointAdapter(ppt), new ScriptedProvider([]));
    const r = await s.runtime.executeDirect(tc("presentation.create_slide", { title: "Agenda", bullets: ["One", "Two"], notes: "Say hi", layout: "Two Content" }));
    expect(r.status).toBe("succeeded");
    const slide = ppt.slidesList[1]!;
    expect(slide.shapes.map((x) => x.text)).toEqual(["Agenda", "One\nTwo"]);
    expect(ppt.notes.get(slide.id)).toBe("Say hi");
    expect((await s.runtime.executeDirect(tc("presentation.create_slide", { title: "x", layout: "Nope" }))).error).toMatch(/Layout "Nope" does not exist/);
    const undo = await s.runtime.executeDirect(tc(r.compensation!.capability, r.compensation!.args));
    expect(undo.status).toBe("succeeded");
    expect(ppt.slidesList).toHaveLength(1);
  });

  it("text boxes, shapes, formatting and geometry are verified; off-slide geometry is refused", async () => {
    const ppt = deck();
    const s = setup(new PowerPointAdapter(ppt), new ScriptedProvider([]));
    const tb = await s.runtime.executeDirect(tc("presentation.add_text_box", { slideId: "s1", text: "Note", left: 100, top: 400, width: 300, height: 40, format: { size: 14, italic: true } }));
    expect(tb.status).toBe("succeeded");
    const sh = await s.runtime.executeDirect(tc("presentation.add_shape", { slideId: "s1", shape: "ellipse", left: 600, top: 300, width: 100, height: 100, fill: "#FF0000", text: "!" }));
    expect(sh.status).toBe("succeeded");
    const fmt = await s.runtime.executeDirect(tc("presentation.format_text", { slideId: "s1", shapeId: "s1sh1", format: { size: 40, color: "#112233", align: "center" } }));
    expect(fmt.status).toBe("succeeded");
    expect(ppt.slidesList[0]!.shapes[0]!.format).toMatchObject({ size: 40, color: "#112233", align: "center" });
    expect((await s.runtime.executeDirect(tc("presentation.set_geometry", { slideId: "s1", shapeId: "s1sh1", left: 900, width: 400 }))).error).toMatch(/beyond the slide width/);
    const geo = await s.runtime.executeDirect(tc("presentation.set_geometry", { slideId: "s1", shapeId: "s1sh1", top: 20 }));
    expect(geo.status).toBe("succeeded");
    expect((await s.runtime.executeDirect(tc("presentation.format_text", { slideId: "s1", shapeId: "s1sh1", format: { color: "red" } }))).error).toMatch(/pattern/);
  });

  it("diagram, bar chart and table produce editable, verified shapes", async () => {
    const ppt = deck();
    const s = setup(new PowerPointAdapter(ppt), new ScriptedProvider([]));
    const d = await s.runtime.executeDirect(tc("presentation.add_diagram", { slideId: "s1", kind: "process", items: [{ title: "Plan" }, { title: "Build", detail: "6 weeks" }, { title: "Ship" }] }));
    expect(d.status).toBe("succeeded");
    expect((d.output as any).shapeIds).toHaveLength(3);
    const c = await s.runtime.executeDirect(tc("presentation.add_bar_chart", { slideId: "s1", categories: ["Q1", "Q2", "Q3"], values: [10, 20, 40], unit: "%", title: "Growth" }));
    expect(c.status).toBe("succeeded");
    expect(ppt.slidesList[0]!.shapes.map((x) => x.text)).toEqual(expect.arrayContaining(["10%", "20%", "40%", "Growth"]));
    expect((await s.runtime.executeDirect(tc("presentation.add_bar_chart", { slideId: "s1", categories: ["a", "b", "c"], values: [1, 2] }))).error).toMatch(/3 categories but 2 values/);
    const t = await s.runtime.executeDirect(tc("presentation.add_table", { slideId: "s1", rows: [["Plan", "Price"], ["Basic", "$10"]] }));
    expect(t.status).toBe("succeeded");
    expect((await s.runtime.executeDirect(tc("presentation.add_table", { slideId: "s1", rows: [["a", "b"], ["c"]] }))).error).toMatch(/same number of cells/);
  });

  it("bar heights are proportional to the values", () => {
    const g = barChartGeometry([10, 20, 40], { left: 0, top: 0, width: 300, height: 256 });
    expect(g.bars.map((b) => Math.round((b.height / g.bars[2]!.height) * 100))).toEqual([25, 50, 100]);
  });

  it("tables and notes are not offered when the host does not support them", () => {
    const ppt = deck();
    ppt.tablesSupported = false;
    ppt.notesSupported = false;
    const names = new PowerPointAdapter(ppt).capabilities().map((c) => c.name);
    expect(names).not.toContain("presentation.add_table");
    expect(names).not.toContain("presentation.set_notes");
    expect(names).toContain("presentation.create_slide");
  });

  it("layout check finds overflow, overlap, off-slide shapes, missing titles and inconsistent sizes", () => {
    const slides = [
      { id: "a", index: 0, shapes: [{ id: "1", name: "Title 1", type: "Placeholder", text: "One", box: { left: 48, top: 30, width: 800, height: 80 }, format: { size: 36 } }] },
      {
        id: "b",
        index: 1,
        shapes: [
          { id: "1", name: "Title 1", type: "Placeholder", text: "Two", box: { left: 48, top: 30, width: 800, height: 80 }, format: { size: 28 } },
          { id: "2", name: "Body", type: "TextBox", text: "word ".repeat(100), box: { left: 48, top: 60, width: 300, height: 50 }, format: { size: 20 } },
          { id: "3", name: "Side", type: "TextBox", text: "x", box: { left: 800, top: 100, width: 300, height: 50 } },
        ],
      },
      { id: "c", index: 2, shapes: [{ id: "1", name: "Body", type: "TextBox", text: "no title here" }] },
    ];
    const kinds = checkLayout(slides, { width: 960, height: 540 }).map((i) => `${i.slide}:${i.kind}`);
    expect(kinds).toEqual(expect.arrayContaining(["2:overflow", "2:overlap", "2:off_slide", "2:dense", "3:no_title", "0:inconsistent_title_size"]));
  });

  it("check_layout says rendering was not inspected", async () => {
    const s = setup(new PowerPointAdapter(deck()), new ScriptedProvider([]));
    const r = await s.runtime.executeDirect(tc("presentation.check_layout", {}));
    expect((r.output as any).renderingVerified).toBe(false);
    expect((r.output as any).method).toMatch(/Rendering was not inspected/);
  });
});

describe("Outlook thread splitting", () => {
  it("splits Outlook- and Gmail-style quoted replies and attributes authors", () => {
    const body = [
      "Sounds good, I'll bring the slides.",
      "",
      "On Tue, Oct 13, 2026 at 10:02 AM Bob Jones <bob@example.com> wrote:",
      "> Can someone bring slides? I will book the room.",
      ">",
      "> From: Carol White <carol@example.com>",
      "> Sent: Monday, October 12, 2026 4:00 PM",
      "> To: Bob Jones",
      "> Subject: Kickoff",
      ">",
      "> Let's do the kickoff Thursday.",
    ].join("\n");
    const msgs = splitThread({ from: "Alice Green <alice@example.com>", body });
    expect(msgs.map((m) => [m.id, m.author])).toEqual([
      ["M1", "Alice Green"],
      ["M2", "Bob Jones"],
      ["M3", "Carol White"],
    ]);
    expect(msgs[1]!.text).toContain("I will book the room");
    expect(msgs[2]!.text).toBe("Let's do the kickoff Thursday.");
  });

  it("does not treat a lone 'From:' line as a header", () => {
    const msgs = splitThread({ from: "a@x.com", body: "From: the team's perspective this is fine.\nThanks" });
    expect(msgs).toHaveLength(1);
  });
});

describe("evidence helpers", () => {
  it("finds numeric claims but skips references, list markers, years and small counts", () => {
    const nums = numbersInText("1. Slide 3 shows Sales!B2:B4 totals 1,250 in 2026, up 12.5% with 3 regions and $4.2M; Q3 and M2 noted.").map((n) => n.raw);
    expect(nums).toEqual(["1,250", "12.5%", "$4.2M"]);
  });

  it("does not skip numbers after ordinary words ending in reference letters", () => {
    expect(numbersInText("North led with 150, and growth hit 37.5% in the month 12").map((n) => n.raw)).toEqual(["150", "37.5%", "12"]);
    expect(numbersInText("See slide 14, paragraph 22, row 120 and ¶ 31.").map((n) => n.raw)).toEqual([]);
  });

  it("matches rounded and scaled values to evidence", () => {
    expect(ungroundedNumbers("Total 1,250 (12.5%), revenue $4.2M, mean 33.3", [1250, 0.125, 4_210_000, 33.333])).toEqual([]);
    expect(ungroundedNumbers("Total 1,300", [1250])).toEqual(["1,300"]);
  });

  it("parses structured items", () => {
    const items = parseItems("- [explicit] Send deck — owner: Bob — due: Friday — source: M2\n* [inferred] Book room — owner: unknown — source: none\nnot an item");
    expect(items).toMatchObject([
      { kind: "explicit", text: "Send deck", owner: "Bob", due: "Friday", source: "M2" },
      { kind: "inferred", text: "Book room", owner: "unknown", source: "NONE" },
    ]);
  });
});
