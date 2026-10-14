/**
 * Complete specialist workflows on in-memory Office hosts with a scripted
 * model (PR spec §9 "Execution tests"). Each test checks Office state, the
 * verification of every action, and the specialist's completion criteria.
 */
import { describe, expect, it } from "vitest";
import {
  ExcelAdapter,
  FakeExcelDriver,
  FakeOutlookDriver,
  FakePowerPointDriver,
  FakeWordDriver,
  OutlookAdapter,
  PowerPointAdapter,
  ScriptedProvider,
  WordAdapter,
  type MailItem,
  type TaskReport,
} from "../src/index.js";
import { setup } from "./helpers.js";

const unmet = (r: TaskReport) => (r.completion ?? []).filter((c) => !c.met).map((c) => `${c.id}: ${c.detail}`);
const call = (name: string, args: Record<string, unknown>, text?: string) => ({ call: { name, arguments: args }, text });

// ------------------------------------------------------------------ Word

describe("Word Document Writer", () => {
  it("proposes an outline, then writes a structured document and verifies it", async () => {
    const doc = new FakeWordDriver(["Notes: launch in Q3 with 40 pilot users.", "Need pricing and support plan."]);
    const model = new ScriptedProvider([
      call("document.write_blocks", { mode: "insert", position: "end", blocks: [{ text: "Too early", style: "Title" }] }),
      call("plan.propose", { outcome: "Test outcome", target: "The open file", title: "Launch plan", steps: [{ title: "Overview" }, { title: "Pricing" }, { title: "Support" }] }),
      call("document.write_blocks", {
        mode: "insert",
        position: "end",
        blocks: [
          { text: "Launch Plan", style: "Title" },
          { text: "Overview", style: "Heading1" },
          { text: "We launch in Q3 with 40 pilot users." },
          { text: "Pricing", style: "Heading1" },
          { text: "Pricing will be finalised before launch." },
          { text: "Support", style: "Heading1" },
          { text: "Support runs through the existing help desk." },
        ],
      }),
      { text: "I wrote the launch plan below your notes." },
    ]);
    const s = setup(new WordAdapter(doc), model);
    const r = await s.runtime.run("Write a launch plan document from my notes", { specialist: s.specialists.get("word.writer")! });

    expect(r.actions[0]!.capability).toBe("context.read"); // required context acquired first
    const early = r.actions.find((a) => a.capability === "document.write_blocks")!;
    expect(early.status).toBe("rejected");
    expect(early.error).toMatch(/needs an approved plan/);
    expect(s.approver.requests.map((q) => q.capability)).toEqual(["plan.propose", "document.write_blocks"]);
    const write = r.actions.filter((a) => a.capability === "document.write_blocks").at(-1)!;
    expect(write.status).toBe("succeeded");
    expect(doc.paragraphs.map((p) => p.style)).toEqual(["Normal", "Normal", "Title", "Heading1", "Normal", "Heading1", "Normal", "Heading1", "Normal"]);
    expect(doc.paragraphs[0]!.text).toMatch(/^Notes/); // original content preserved
    expect(r.plan?.steps).toHaveLength(3);
    expect(unmet(r)).toEqual([]);
    expect(r.outcome).toBe("completed");
  });

  it("does not report success when the written content fails verification", async () => {
    const doc = new FakeWordDriver(["x"]);
    doc.dropNextWrite = true;
    const model = new ScriptedProvider([
      call("plan.propose", { outcome: "Test outcome", target: "The open file", title: "Doc", steps: [{ title: "Intro" }] }),
      call("document.write_blocks", { mode: "insert", position: "end", blocks: [{ text: "Doc", style: "Title" }, { text: "Intro", style: "Heading1" }, { text: "Body one." }, { text: "Body two." }] }),
      { text: "Done! The document is complete." },
    ]);
    const s = setup(new WordAdapter(doc), model);
    const r = await s.runtime.run("Write a short report", { specialist: s.specialists.get("word.writer")! });
    expect(r.outcome).not.toBe("completed");
    expect(unmet(r).join(" ")).toMatch(/content-written/);
  });
});

describe("Word Document Editor", () => {
  const paras = ["Intro paragraph stays.", "Revenue rose 12% to $4.2M in Q3, driven by Acme Corp renewals, according to Finance.", "Closing paragraph stays."];

  it("changes only the requested paragraph and keeps facts", async () => {
    const doc = new FakeWordDriver([...paras]);
    const model = new ScriptedProvider([
      call("document.replace_text", { target: "paragraphs", start: 1, expectedText: paras[1], newText: "Revenue rose 12% to $4.2M in Q3 on Acme Corp renewals (Finance)." }),
      { text: "Tightened paragraph 1." },
    ]);
    const s = setup(new WordAdapter(doc), model);
    const r = await s.runtime.run("Make paragraph 1 more concise", { specialist: s.specialists.get("word.editor")! });
    expect(doc.paragraphs[0]!.text).toBe(paras[0]);
    expect(doc.paragraphs[2]!.text).toBe(paras[2]);
    expect(doc.paragraphs[1]!.text).toMatch(/on Acme Corp renewals/);
    expect(s.approver.requests[0]!.preview.before).toContain("according to Finance");
    expect(unmet(r)).toEqual([]);
    expect(r.outcome).toBe("completed");
  });

  it("flags an edit that drops facts", async () => {
    const doc = new FakeWordDriver([...paras]);
    const model = new ScriptedProvider([call("document.replace_text", { target: "paragraphs", start: 1, expectedText: paras[1], newText: "Revenue rose in Q3." }), { text: "Done." }]);
    const s = setup(new WordAdapter(doc), model);
    const r = await s.runtime.run("Make paragraph 1 more concise", { specialist: s.specialists.get("word.editor")! });
    expect(r.outcome).toBe("partial");
    expect(unmet(r).join(" ")).toMatch(/facts-kept: missing after the edit: .*12%/);
  });

  it("detects changes outside the targeted paragraphs", async () => {
    const doc = new FakeWordDriver([...paras]);
    const model = new ScriptedProvider([
      call("document.replace_text", { target: "paragraphs", start: 1, expectedText: paras[1], newText: "Revenue rose 12% to $4.2M in Q3 on Acme Corp renewals (Finance)." }),
      () => {
        doc.paragraphs[2]!.text = "Something else changed this."; // e.g. a faulty driver
        return { message: { role: "assistant", content: "Done." } };
      },
    ]);
    const s = setup(new WordAdapter(doc), model);
    const r = await s.runtime.run("Make paragraph 1 more concise", { specialist: s.specialists.get("word.editor")! });
    expect(unmet(r).join(" ")).toMatch(/preserves-unrelated: untargeted paragraph\(s\) changed: 2/);
    expect(r.outcome).toBe("partial");
  });
});

describe("Word Document Reviewer and Summarizer", () => {
  it("reviewer reads, cites and never modifies", async () => {
    const doc = new FakeWordDriver(["Plan", "We will launch in March.", "Launch is scheduled for June.", "Budget: TBD."]);
    const model = new ScriptedProvider([{ text: "1. [Observation] ¶1 says March but ¶2 says June.\n   [Recommendation] Pick one date.\n2. [Observation] ¶3 has no budget.\n   [Recommendation] Add the figure." }]);
    const s = setup(new WordAdapter(doc), model);
    const r = await s.runtime.run("Review this plan for contradictions", { specialist: s.specialists.get("word.reviewer")! });
    expect(unmet(r)).toEqual([]);
    expect(model.requests[0]!.tools!.every((t) => !t.name.startsWith("document."))).toBe(true);
  });

  it("summarizer must disclose a partial read and not invent numbers", async () => {
    const paras = Array.from({ length: 260 }, (_, i) => `Paragraph ${i}: revenue line ${i * 3}.`);
    const doc = new FakeWordDriver(paras);
    const model = new ScriptedProvider([{ text: "Revenue lines grow steadily, reaching 9,999." }]);
    const s = setup(new WordAdapter(doc), model);
    const r = await s.runtime.run("Summarize this document", { specialist: s.specialists.get("word.summarizer")! });
    const failed = unmet(r).join(" ");
    expect(failed).toMatch(/discloses-partial/);
    expect(failed).toMatch(/numbers-grounded: .*9,999/);
    expect(r.outcome).toBe("partial");
  });
});

// ------------------------------------------------------------------ Excel

const salesData = () =>
  new FakeExcelDriver({
    Sales: [
      ["Region", "Q1", "Q2"],
      ["North", 120, 135],
      ["South", 90, 101],
      ["West", 75, 88],
    ],
  });

describe("Excel Data Analyst", () => {
  it("reports numbers calculated by the application with cited ranges", async () => {
    const xl = salesData();
    const model = new ScriptedProvider([
      call("workbook.compute_stats", { sheet: "Sales", measures: [{ column: "Q2", ops: ["sum", "max", "mean"] }] }),
      { text: "Calculated: Q2 revenue totals 324 across regions (sum of Sales!C2:C4); North is highest at 135; the mean is 108." },
    ]);
    const s = setup(new ExcelAdapter(xl), model);
    const r = await s.runtime.run("What is total Q2 revenue and which region is highest?", { specialist: s.specialists.get("excel.analyst")! });
    const stats = r.actions.find((a) => a.capability === "workbook.compute_stats")!;
    expect((stats.output as any).rows[0]).toEqual([324, 135, 108]);
    expect(unmet(r)).toEqual([]);
    expect(r.outcome).toBe("completed");
  });

  it("flags figures that were not calculated or read", async () => {
    const model = new ScriptedProvider([
      call("workbook.compute_stats", { sheet: "Sales", measures: [{ column: "Q2", ops: ["sum"] }] }),
      { text: "Q2 totals 324 (Sales!C2:C4), up 17.5% on last year." },
    ]);
    const s = setup(new ExcelAdapter(salesData()), model);
    const r = await s.runtime.run("Analyse Q2", { specialist: s.specialists.get("excel.analyst")! });
    expect(unmet(r).join(" ")).toMatch(/numbers-grounded: .*17\.5%/);
    expect(r.outcome).toBe("partial");
  });
});

describe("Excel Formula Expert", () => {
  it("writes formulas and verifies the formula and resulting values", async () => {
    const xl = salesData();
    const model = new ScriptedProvider([
      call("workbook.write_range", { sheet: "Sales", address: "D1:D4", formulas: [["Total"], ["=SUM(B2:C2)"], ["=SUM(B3:C3)"], ["=SUM(B4:C4)"]] }),
      { text: "D2:D4 now sum Q1 and Q2 for each row (Sales!B2:C4); D2 = 255." },
    ]);
    const s = setup(new ExcelAdapter(xl), model);
    const r = await s.runtime.run("Add a formula in column D totalling Q1 and Q2", { specialist: s.specialists.get("excel.formula")! });
    const w = r.actions.find((a) => a.capability === "workbook.write_range")!;
    expect(w.status).toBe("succeeded");
    expect((await xl.readRange("Sales", "D2:D4")).values).toEqual([[255], [191], [163]]);
    expect((await xl.readRange("Sales", "D2")).formulas).toEqual([["=SUM(B2:C2)"]]);
    expect(unmet(r)).toEqual([]);
  });

  it("does not report completion when the formula evaluates to an error", async () => {
    const model = new ScriptedProvider([call("workbook.write_range", { sheet: "Sales", address: "E2", formulas: [["=VLOOKUP(A2,Prices!A:B,2,FALSE)"]] }), { text: "E2 now looks up the price." }]);
    const s = setup(new ExcelAdapter(salesData()), model);
    const r = await s.runtime.run("Add a price lookup formula in E2", { specialist: s.specialists.get("excel.formula")! });
    expect(unmet(r).join(" ")).toMatch(/no-formula-errors/);
    expect(r.outcome).toBe("partial");
  });
});

describe("Excel Data Cleaner", () => {
  it("profiles, writes a cleaned copy to a new sheet, and preserves the source", async () => {
    const xl = new FakeExcelDriver({
      Orders: [
        ["Region", "Amount"],
        [" north", 10],
        ["North", 12],
        ["South ", 9],
        ["South", 9],
        ["north", 10],
      ],
    });
    const before = JSON.stringify(await xl.readRange("Orders", "A1:B6"));
    const model = new ScriptedProvider([
      call("workbook.profile_data", { sheet: "Orders" }),
      call("workbook.write_range", { sheet: "Orders", address: "A2", values: [["North"]] }),
      call("workbook.clean_to_sheet", { sheet: "Orders", targetSheet: "Orders (clean)", operations: [{ op: "trim" }, { op: "canonicalize_categories", column: "Region" }, { op: "dedupe" }] }),
      { text: "Cleaned copy written to 'Orders (clean)': 5 rows in, 3 out (2 duplicates removed)." },
    ]);
    const s = setup(new ExcelAdapter(xl), model);
    const r = await s.runtime.run("Clean this data: fix spacing, unify region names, remove duplicates", { specialist: s.specialists.get("excel.cleaner")! });

    const overwrite = r.actions.find((a) => a.capability === "workbook.write_range")!;
    expect(overwrite.status).toBe("rejected");
    expect(overwrite.error).toMatch(/only writes to sheets it created/);
    const profile = r.actions.find((a) => a.capability === "workbook.profile_data")!.output as any;
    expect(profile.profile.columns[0].inconsistentCategories.length).toBeGreaterThan(0);
    const clean = r.actions.find((a) => a.capability === "workbook.clean_to_sheet")!;
    expect(clean.status).toBe("succeeded");
    expect((clean.output as any).report).toMatchObject({ inputRows: 5, outputRows: 3, removedDuplicates: 2 });
    expect((await xl.readRange("Orders (clean)", "A1:B4")).values).toEqual([["Region", "Amount"], ["north", 10], ["north", 12], ["South", 9]]); // most frequent spelling wins
    expect(JSON.stringify(await xl.readRange("Orders", "A1:B6"))).toBe(before);
    expect(unmet(r)).toEqual([]);
  });
});

describe("Excel Dashboard Builder", () => {
  it("creates a dashboard sheet with traceable metrics and a verified chart", async () => {
    const xl = salesData();
    const model = new ScriptedProvider([
      call("workbook.compute_stats", { sheet: "Sales", measures: [{ column: "Q1", ops: ["sum"] }, { column: "Q2", ops: ["sum"] }] }),
      call("plan.propose", { outcome: "Test outcome", target: "The open file", title: "Sales dashboard", steps: [{ title: "Sheet Dashboard" }, { title: "Totals per quarter as formulas" }, { title: "Column chart" }] }),
      call("workbook.add_worksheet", { name: "Dashboard" }),
      call("workbook.write_range", { sheet: "Dashboard", address: "A1:B3", values: [["Quarter", "Total"], ["Q1", 285], ["Q2", 324]] }),
      call("workbook.write_range", { sheet: "Dashboard", address: "A1:B3", formulas: [["Quarter", "Total"], ["Q1", "=SUM(Sales!B2:B4)"], ["Q2", "=SUM(Sales!C2:C4)"]] }),
      call("workbook.create_chart", { sourceSheet: "Dashboard", sourceAddress: "A1:B3", kind: "column", title: "Revenue by quarter" }),
      { text: "Dashboard created: Q1 total 285 (=SUM(Sales!B2:B4)), Q2 total 324 (=SUM(Sales!C2:C4)), with a column chart." },
    ]);
    const s = setup(new ExcelAdapter(xl), model);
    const r = await s.runtime.run("Create a summary dashboard of revenue by quarter", { specialist: s.specialists.get("excel.dashboard")! });
    const typed = r.actions.find((a) => a.capability === "workbook.write_range" && (a.args as any).values)!;
    expect(typed.status).toBe("rejected");
    expect(typed.error).toMatch(/formulas that reference the source/);
    expect((await xl.readRange("Dashboard", "B2:B3")).values).toEqual([[285], [324]]);
    expect(xl.chartsList).toEqual([{ sheet: "Dashboard", name: "Chart 1", type: "column", title: "Revenue by quarter" }]);
    expect(unmet(r)).toEqual([]);
    expect(r.outcome).toBe("completed");
  });
});

// -------------------------------------------------------------- PowerPoint

describe("PowerPoint Presentation Author", () => {
  const steps = [
    call("plan.propose", {
      outcome: "Test outcome", target: "The open file", title: "Launch deck",
      steps: [{ title: "Lighthouse launch", detail: "title slide" }, { title: "The problem" }, { title: "Our approach", detail: "process diagram" }],
    }),
    call("presentation.list_layouts", {}),
    call("presentation.create_slide", { title: "Lighthouse launch", subtitle: "Q3 plan" }),
    call("presentation.create_slide", { title: "The problem", bullets: ["Reports take 3 days", "Data is scattered"], notes: "Open with the customer quote." }),
    call("presentation.create_slide", { title: "Our approach", layout: "Title Only" }),
    call("presentation.add_diagram", { slideId: "slide1002", kind: "process", items: [{ title: "Collect" }, { title: "Analyse" }, { title: "Report" }] }),
    call("presentation.check_layout", {}),
  ];

  it("creates multiple distinct, verified slides after an approved plan", async () => {
    const ppt = new FakePowerPointDriver([]);
    const model = new ScriptedProvider([...steps, { text: "Created 3 slides. The layout check found no serious issues." }]);
    const s = setup(new PowerPointAdapter(ppt), model);
    const r = await s.runtime.run("Create a 3-slide presentation about the Lighthouse launch", { specialist: s.specialists.get("powerpoint.author")! });

    expect(ppt.slidesList).toHaveLength(3);
    const texts = ppt.slidesList.map((sl) => sl.shapes.map((x) => x.text).join(" | "));
    expect(texts[0]).toBe("Lighthouse launch | Q3 plan");
    expect(texts[1]).toBe("The problem | Reports take 3 days\nData is scattered");
    expect(texts[2]).toContain("Collect");
    expect(ppt.slidesList.every((sl) => sl.shapes.every((x) => x.text !== ""))).toBe(true); // no empty placeholders left
    expect(ppt.notes.get("slide1001")).toBe("Open with the customer quote.");
    expect(r.actions.filter((a) => a.capability === "presentation.create_slide").every((a) => a.status === "succeeded")).toBe(true);
    expect(unmet(r)).toEqual([]);
    expect(r.outcome).toBe("completed");
  });

  it("reports speaker notes as a limitation when the host does not support them", async () => {
    const ppt = new FakePowerPointDriver([]);
    ppt.notesSupported = false;
    const model = new ScriptedProvider([...steps, { text: "Created 3 slides." }]);
    const s = setup(new PowerPointAdapter(ppt), model);
    const r = await s.runtime.run("Create a 3-slide presentation about the Lighthouse launch", { specialist: s.specialists.get("powerpoint.author")! });
    const withNotes = r.actions.find((a) => a.capability === "presentation.create_slide" && (a.args as any).notes)!;
    expect(withNotes.status).toBe("partial");
    expect((withNotes.output as any).notes).toBe("not_supported");
    expect(unmet(r).join(" ")).toMatch(/notes-limitation-reported/);
    expect(r.outcome).toBe("partial");
  });

  it("will not touch existing slides unless asked", async () => {
    const ppt = new FakePowerPointDriver([{ shapes: [{ name: "Title", text: "Existing" }] }]);
    const model = new ScriptedProvider([
      call("plan.propose", { outcome: "Test outcome", target: "The open file", title: "x", steps: [{ title: "a" }] }),
      call("presentation.format_text", { slideId: "s1", shapeId: "s1sh1", format: { size: 10 } }),
      { text: "ok" },
    ]);
    const s = setup(new PowerPointAdapter(ppt), model);
    const r = await s.runtime.run("Create a 1-slide deck", { specialist: s.specialists.get("powerpoint.author")! });
    expect(r.actions.find((a) => a.capability === "presentation.format_text")!.error).toMatch(/Only slides created in this task/);
    expect(ppt.slidesList[0]!.shapes[0]!.format!.size).toBe(32);
  });
});

describe("PowerPoint Slide Designer", () => {
  it("fixes layout problems without erasing unrelated content", async () => {
    const ppt = new FakePowerPointDriver([
      {
        shapes: [
          { name: "Title 1", text: "Results", format: { size: 20, bold: true } },
          { name: "Body", text: "Revenue grew 12% to $4.2M. ".repeat(12).trim(), box: { left: 48, top: 130, width: 400, height: 60 }, format: { size: 20 } },
          { name: "Footer", text: "Confidential", box: { left: 48, top: 500, width: 200, height: 20 }, format: { size: 10 } },
        ],
      },
    ]);
    const model = new ScriptedProvider([
      call("presentation.check_layout", {}),
      call("presentation.edit_text", { slideId: "s1", shapeId: "s1sh2", text: "short" }),
      call("presentation.format_text", { slideId: "s1", shapeId: "s1sh1", format: { size: 36 } }),
      call("presentation.set_geometry", { slideId: "s1", shapeId: "s1sh2", width: 864, height: 360 }),
      call("presentation.check_layout", {}),
      { text: "Slide 1: enlarged the title to 36pt and resized the body so the text fits. Rendering was not inspected." },
    ]);
    const s = setup(new PowerPointAdapter(ppt), model);
    const r = await s.runtime.run("Make this slide more readable", { specialist: s.specialists.get("powerpoint.designer")! });

    const first = r.actions.find((a) => a.capability === "presentation.check_layout")!.output as any;
    expect(first.issues.map((i: any) => i.kind)).toContain("overflow");
    expect(r.actions.find((a) => a.capability === "presentation.edit_text")!.error).toMatch(/must quote the shape's current text/);
    const shapes = ppt.slidesList[0]!.shapes;
    expect(shapes[0]!.format!.size).toBe(36);
    expect(shapes[1]!.text).toMatch(/^Revenue grew 12%/);
    expect(shapes[2]!.text).toBe("Confidential");
    const last = r.actions.filter((a) => a.capability === "presentation.check_layout").at(-1)!.output as any;
    expect(last.issues.filter((i: any) => i.kind === "overflow")).toHaveLength(0);
    expect(unmet(r)).toEqual([]);
  });
});

describe("PowerPoint Deck Reviewer", () => {
  const deck = () =>
    new FakePowerPointDriver([
      { shapes: [{ name: "Title 1", text: "Strategy" }, { name: "Body", text: "We will win." }] },
      { shapes: [{ name: "Body", text: "Our market share will triple next year. ".repeat(10), box: { left: 48, top: 40, width: 400, height: 80 }, format: { size: 20 } }] },
    ]);

  it("identifies a deliberately introduced problem", async () => {
    const model = new ScriptedProvider([
      call("presentation.check_layout", {}),
      { text: "High priority:\n- Slide 2 has no title and its body text overflows; the claim that share will triple has no evidence.\nLow priority:\n- Slide 1 needs supporting data. Layout findings come from slide objects; the rendered slides were not inspected." },
    ]);
    const s = setup(new PowerPointAdapter(deck()), model);
    const r = await s.runtime.run("Review this deck", { specialist: s.specialists.get("powerpoint.reviewer")! });
    const issues = (r.actions.find((a) => a.capability === "presentation.check_layout")!.output as any).issues;
    expect(issues.some((i: any) => i.slide === 2 && i.kind === "overflow" && i.severity === "high")).toBe(true);
    expect(issues.some((i: any) => i.slide === 2 && i.kind === "no_title")).toBe(true);
    expect(unmet(r)).toEqual([]);
  });

  it("is incomplete if it misses the problem or claims visual verification", async () => {
    const model = new ScriptedProvider([call("presentation.check_layout", {}), { text: "High: Slide 1 is fine. I visually verified every slide." }]);
    const s = setup(new PowerPointAdapter(deck()), model);
    const r = await s.runtime.run("Review this deck", { specialist: s.specialists.get("powerpoint.reviewer")! });
    const failed = unmet(r).join(" ");
    expect(failed).toMatch(/reports-issues: not reported: slide 2/);
    expect(failed).toMatch(/no-rendering-claims/);
    expect(r.outcome).toBe("partial");
  });
});

// ---------------------------------------------------------------- Outlook

const thread = (mode: "read" | "compose" = "read"): MailItem => ({
  itemType: "message",
  subject: "Budget review",
  from: "Alex Lee <alex@example.com>",
  to: ["dana@example.com"],
  cc: [],
  body: [
    "Thanks Dana, I'll send the revised budget by Friday.",
    "",
    "From: Dana Smith <dana@example.com>",
    "Sent: Monday, October 12, 2026 9:14 AM",
    "To: Alex Lee <alex@example.com>",
    "Subject: Budget review",
    "",
    "Hi Alex, can you send the revised budget? I will book the room for Thursday. Could you also confirm the vendor list?",
  ].join("\n"),
  attachments: [],
  ...(mode === "compose" ? { from: undefined } : {}),
});

describe("Outlook Thread Analyst", () => {
  it("extracts correctly attributed commitments", async () => {
    const model = new ScriptedProvider([
      {
        text: [
          "Summary: Dana asked Alex for the revised budget; Alex agreed.",
          "Commitments:",
          "- [explicit] Send the revised budget — owner: Alex Lee — due: Friday — source: M1",
          "- [explicit] Book the room for Thursday — owner: Dana Smith — due: Thursday — source: M2",
          "- [inferred] Confirm the vendor list — owner: Alex Lee — due: none — source: M2",
        ].join("\n"),
      },
    ]);
    const s = setup(new OutlookAdapter(new FakeOutlookDriver(thread())), model);
    const r = await s.runtime.run("Summarize this thread and who committed to what", { specialist: s.specialists.get("outlook.thread")! });
    expect(r.actions[0]!.capability).toBe("outlook.split_thread");
    expect(JSON.stringify(model.requests[0]!.messages)).toContain("[M2] From: Dana Smith");
    expect(unmet(r)).toEqual([]);
  });

  it("rejects a misattributed commitment", async () => {
    const model = new ScriptedProvider([{ text: "Summary: ...\nCommitments:\n- [explicit] Send the revised budget — owner: Dana Smith — due: Friday — source: M1" }]);
    const s = setup(new OutlookAdapter(new FakeOutlookDriver(thread())), model);
    const r = await s.runtime.run("Summarize the commitments", { specialist: s.specialists.get("outlook.thread")! });
    expect(unmet(r).join(" ")).toMatch(/M1 was written by Alex Lee, not Dana Smith/);
    expect(r.outcome).toBe("partial");
  });
});

describe("Outlook Reply Writer", () => {
  it("opens a reply draft without sending and reports it as unverified", async () => {
    const ol = new FakeOutlookDriver(thread());
    const model = new ScriptedProvider([
      call("outlook.create_draft", { body: "Hi Dana,\n\nThe revised budget will be with you by Friday. The vendor list is [confirm vendors].\n\nAlex" }),
      { text: "Draft opened (not sent). Fill in the placeholder [confirm vendors]. Outlook does not let me verify the reply window." },
    ]);
    const s = setup(new OutlookAdapter(ol), model);
    const r = await s.runtime.run("Draft a reply answering Dana's questions", { specialist: s.specialists.get("outlook.reply")! });
    expect(ol.openedForms).toHaveLength(1);
    expect(ol.sent).toHaveLength(0);
    expect(r.actions.find((a) => a.capability === "outlook.create_draft")!.status).toBe("unverified");
    expect(unmet(r)).toEqual([]);
    expect(r.outcome).toBe("partial"); // honest: the draft exists but could not be verified
  });

  it("verifies a draft inserted in compose mode", async () => {
    const ol = new FakeOutlookDriver(thread("compose"), "compose");
    const model = new ScriptedProvider([call("outlook.create_draft", { body: "Thanks, confirmed." }), { text: "Inserted the reply. Nothing was sent." }]);
    const s = setup(new OutlookAdapter(ol), model);
    const r = await s.runtime.run("Write a short reply confirming", { specialist: s.specialists.get("outlook.reply")! });
    expect(r.outcome).toBe("completed");
    expect(ol.composeText).toContain("Thanks, confirmed.");
  });
});

describe("Outlook Action Extractor", () => {
  it("distinguishes explicit assignments from inferred tasks", async () => {
    const model = new ScriptedProvider([
      {
        text: [
          "- [explicit] Send the revised budget — owner: Alex — due: Friday — source: M2",
          "- [explicit] Book the room for Thursday — owner: Dana Smith — due: Thursday — source: M2",
          "- [inferred] Share the agenda before the meeting — owner: unknown — due: none — source: none",
          "Dependencies: none",
          "Open questions: vendor list (M2)",
        ].join("\n"),
      },
    ]);
    const s = setup(new OutlookAdapter(new FakeOutlookDriver(thread())), model);
    const r = await s.runtime.run("Extract the action items with owners", { specialist: s.specialists.get("outlook.actions")! });
    expect(unmet(r)).toEqual([]);
  });

  it("rejects an explicit item that the cited message does not contain", async () => {
    const model = new ScriptedProvider([{ text: "- [explicit] Approve the hiring plan — owner: Alex — due: none — source: M2" }]);
    const s = setup(new OutlookAdapter(new FakeOutlookDriver(thread())), model);
    const r = await s.runtime.run("Extract the action items", { specialist: s.specialists.get("outlook.actions")! });
    expect(unmet(r).join(" ")).toMatch(/no wording from M2 supports/);
  });
});
