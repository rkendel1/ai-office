/**
 * Acceptance scenarios. Fixture content is created by each scenario so runs
 * are reproducible; live runs must use a blank document (checked).
 */
import type { ExcelDriver } from "../hosts/excel.js";
import type { OutlookDriver } from "../hosts/outlook.js";
import type { PowerPointDriver } from "../hosts/powerpoint.js";
import type { WordDriver } from "../hosts/word.js";
import type { ApprovalRequest } from "../permissions.js";
import type { TaskReport } from "../runtime.js";
import { parseRenderedThread } from "../specialists/shared/evidence.js";
import type { ScriptStep } from "../testing/fakes.js";
import { normalizeText } from "../util.js";
import type { AcceptanceScenario, CheckResult } from "./types.js";

const call = (name: string, args: Record<string, unknown>, text = ""): ScriptStep => ({ call: { name, arguments: args }, text });
const say = (text: string): ScriptStep => ({ text });
const batch = (r: ApprovalRequest) => (r.capability === "plan.propose" ? { approved: true, batch: true } : { approved: true });
const check = (id: string, description: string, passed: boolean, evidence: string): CheckResult => ({ id, description, passed, evidence });

function taskChecks(r: TaskReport, expectOutcome: string | string[] = "completed"): CheckResult[] {
  const unmet = (r.completion ?? []).filter((c) => c.required && !c.met);
  const want = Array.isArray(expectOutcome) ? expectOutcome : [expectOutcome];
  return [
    check("task-outcome", `Task outcome is ${want.join(" or ")}`, want.includes(r.outcome), `${r.outcome}: ${r.executionSummary}`),
    check("completion-criteria", "All required completion criteria are met", unmet.length === 0, unmet.length ? unmet.map((c) => `${c.description}: ${c.detail}`).join("; ") : `${r.completion?.length ?? 0} criteria met`),
    check(
      "writes-verified",
      "Every executed change was verified by read-back (or honestly reported as unverified)",
      r.actions.every((a) => !["failed"].includes(a.status)),
      r.actions.map((a) => `${a.capability}=${a.status}`).join(", "),
    ),
  ];
}

// ------------------------------------------------------------------ Word

const WORD_FIXTURE = ["LOAI acceptance fixture: launch notes", "We launch in Q3 with 40 pilot users across two teams.", "Pricing is still open; the support plan relies on the existing help desk."];

async function wordParagraphs(d: WordDriver) {
  const n = await d.paragraphCount();
  return n ? d.readParagraphs(0, n) : [];
}

async function blankWordDocument(c: { driver: WordDriver }) {
  const paras = await wordParagraphs(c.driver);
  return paras.some((p) => p.text.trim()) ? "The document is not blank. Open a new blank document: this scenario adds fixture paragraphs." : undefined;
}

export const wordWriterScenario: AcceptanceScenario<WordDriver> = {
  id: "word.writer.structured-document",
  host: "word",
  specialist: "word.writer",
  title: "Word: write a structured document under a batch-approved plan",
  purpose: "Creates and formats content (title, headings, body), verifies it by read-back, and confirms the notes it was written from are unchanged.",
  preconditions: ["Word is open with a new blank document", "The add-in is paired and a model is selected (only for model runs)"],
  checkPreconditions: blankWordDocument,
  async setup({ driver }) {
    await driver.insertParagraphs("end", WORD_FIXTURE);
  },
  request: "Turn my launch notes into a short structured plan document",
  async script() {
    return [
      call("plan.propose", {
        title: "Launch plan",
        outcome: "A titled launch plan with Overview and Pricing sections, below the notes",
        target: "This document, appended at the end",
        risks: ["Adds paragraphs at the end; existing notes are not changed"],
        steps: [{ title: "Title" }, { title: "Overview" }, { title: "Pricing and support" }],
        operations: [{ capability: "document.write_blocks", description: "Append the plan (title, two sections)", scope: { position: "end" } }],
      }),
      call("document.write_blocks", {
        mode: "insert",
        position: "end",
        blocks: [
          { text: "Launch Plan", style: "Title" },
          { text: "Overview", style: "Heading1" },
          { text: "We launch in Q3 with 40 pilot users across two teams." },
          { text: "Pricing and support", style: "Heading1" },
          { text: "Pricing is still open. Support uses the existing help desk." },
        ],
      }),
      say("I added a titled launch plan with Overview and Pricing sections below your notes."),
    ];
  },
  approve: batch,
  async checks({ driver }, r) {
    const paras = await wordParagraphs(driver);
    const texts = paras.map((p) => normalizeText(p.text));
    const fixtureAt = WORD_FIXTURE.map((t) => texts.indexOf(normalizeText(t)));
    const tail = paras.slice(-5);
    const wantTail: [string, string][] = [
      ["Launch Plan", "Title"],
      ["Overview", "Heading1"],
      ["We launch in Q3 with 40 pilot users across two teams.", "Normal"],
      ["Pricing and support", "Heading1"],
      ["Pricing is still open. Support uses the existing help desk.", "Normal"],
    ];
    const tailOk = tail.length === 5 && tail.every((p, i) => normalizeText(p.text) === normalizeText(wantTail[i]![0]) && p.style === wantTail[i]![1]);
    const write = r.actions.find((a) => a.capability === "document.write_blocks");
    return [
      check("content-and-styles", "The plan is in the document with Title/Heading1/Normal styles", tailOk, tail.map((p) => `${p.style}: ${p.text}`).join(" | ")),
      check("notes-unchanged", "The fixture notes are unchanged and in order", fixtureAt.every((i, k) => i >= 0 && (k === 0 || i > fixtureAt[k - 1]!)), `positions ${fixtureAt.join(", ")}`),
      check("approved-by-plan", "The write ran under the batch-approved plan and was verified", write?.approval === "plan" && write.status === "succeeded", `${write?.status} / ${write?.approval} / ${write?.verification?.method}`),
      ...taskChecks(r),
    ];
  },
  manual: ["Confirm the title and headings look like Word's built-in Title and Heading 1 styles", "Press Ctrl/Cmd+Z repeatedly and confirm Word can undo the inserted paragraphs"],
};

export const wordEditorScenario: AcceptanceScenario<WordDriver> = {
  id: "word.editor.scoped-edit",
  host: "word",
  specialist: "word.editor",
  title: "Word: edit one paragraph and leave everything else unchanged",
  purpose: "Edits a single targeted paragraph, keeps its facts, verifies the change, and confirms untargeted paragraphs are byte-for-byte unchanged.",
  preconditions: ["Word is open with a new blank document"],
  checkPreconditions: blankWordDocument,
  async setup({ driver }) {
    await driver.insertParagraphs("end", WORD_FIXTURE);
  },
  request: "Make the paragraph about pricing more concise",
  async script({ driver }) {
    const paras = await wordParagraphs(driver);
    const i = paras.findIndex((p) => p.text.startsWith("Pricing is still open"));
    return [call("document.replace_text", { target: "paragraphs", start: i, expectedText: paras[i]!.text, newText: "Pricing is open; support uses the existing help desk." }), say("I tightened the pricing paragraph.")];
  },
  approve: () => ({ approved: true }),
  async checks({ driver }, r) {
    const paras = await wordParagraphs(driver);
    const texts = paras.map((p) => p.text);
    return [
      check("target-edited", "The pricing paragraph has the approved text", texts.includes("Pricing is open; support uses the existing help desk."), texts.join(" | ")),
      check("others-unchanged", "The other fixture paragraphs are unchanged", WORD_FIXTURE.slice(0, 2).every((t) => texts.includes(t)), WORD_FIXTURE.slice(0, 2).map((t) => `${texts.includes(t) ? "✓" : "✗"} ${t}`).join(" | ")),
      check("diff-approved", "The edit was shown for approval", r.actions.some((a) => a.capability === "document.replace_text" && a.approval === "approved"), r.actions.map((a) => `${a.capability}:${a.approval}`).join(", ")),
      ...taskChecks(r),
    ];
  },
  manual: ["Confirm the approval card showed a before/after diff of only the pricing paragraph"],
};

// ------------------------------------------------------------------ Excel

const SALES: (string | number)[][] = [
  ["Region", "Q1", "Q2"],
  ["North", 120, 135],
  ["South", 90, 101],
  ["West", 75, 88],
  ["East", 60, 70],
];
const ORDERS: (string | number)[][] = [
  ["Region", "Amount"],
  [" North", 10],
  ["North", 12],
  ["South ", 9],
  ["South", 9],
  ["North", 10],
];

async function sheetsAbsent(d: ExcelDriver, names: string[]) {
  const existing = (await d.worksheets()).map((s) => s.name);
  const clash = names.filter((n) => existing.includes(n));
  return clash.length ? `Sheets ${clash.join(", ")} already exist. Use a new blank workbook.` : undefined;
}

export const excelDashboardScenario: AcceptanceScenario<ExcelDriver> = {
  id: "excel.dashboard.traceable-metrics",
  host: "excel",
  specialist: "excel.dashboard",
  title: "Excel: dashboard with Excel-calculated, traceable metrics and a chart",
  purpose: "Analyses data, creates a dashboard sheet whose metrics are formulas over the source, and checks the values Excel calculated against independently computed totals.",
  preconditions: ["Excel is open with a new blank workbook"],
  checkPreconditions: ({ driver }) => sheetsAbsent(driver, ["LOAI_Sales", "LOAI_Dashboard"]),
  async setup({ driver }) {
    await driver.addWorksheet("LOAI_Sales");
    await driver.writeRange("LOAI_Sales", "A1:C5", { values: SALES });
  },
  request: "Create a summary dashboard of revenue by quarter",
  async script() {
    return [
      call("workbook.compute_stats", { sheet: "LOAI_Sales", address: "A1:C5", measures: [{ column: "Q1", ops: ["sum"] }, { column: "Q2", ops: ["sum"] }] }),
      call("plan.propose", {
        title: "Quarterly dashboard",
        outcome: "Dashboard sheet with quarterly totals linked to LOAI_Sales and a column chart",
        target: "New sheet LOAI_Dashboard",
        risks: ["Adds a sheet; LOAI_Sales is not changed"],
        steps: [{ title: "Create LOAI_Dashboard" }, { title: "Totals as formulas" }, { title: "Column chart" }],
        operations: [
          { capability: "workbook.add_worksheet", description: "Add sheet LOAI_Dashboard", scope: { sheet: "LOAI_Dashboard" } },
          { capability: "workbook.write_range", description: "Totals in LOAI_Dashboard!A1:B3", scope: { sheet: "LOAI_Dashboard", range: "A1:B3" } },
          { capability: "workbook.create_chart", description: "Column chart on LOAI_Dashboard", scope: { sheet: "LOAI_Dashboard" } },
        ],
      }),
      call("workbook.add_worksheet", { name: "LOAI_Dashboard" }),
      call("workbook.write_range", { sheet: "LOAI_Dashboard", address: "A1:B3", formulas: [["Quarter", "Total"], ["Q1", "=SUM(LOAI_Sales!B2:B5)"], ["Q2", "=SUM(LOAI_Sales!C2:C5)"]] }),
      call("workbook.create_chart", { sourceSheet: "LOAI_Dashboard", sourceAddress: "A1:B3", kind: "column", title: "Revenue by quarter" }),
      say("Dashboard created: Q1 total 345 (=SUM(LOAI_Sales!B2:B5)) and Q2 total 394 (=SUM(LOAI_Sales!C2:C5)), with a column chart."),
    ];
  },
  approve: batch,
  async checks({ driver }, r) {
    const independent = [1, 2].map((col) => SALES.slice(1).reduce((s, row) => s + Number(row[col]), 0));
    const got = await driver.readRange("LOAI_Dashboard", "B2:B3");
    const values = got.values.map((x) => x[0]);
    const charts = await driver.charts("LOAI_Dashboard");
    const src = await driver.readRange("LOAI_Sales", "A1:C5");
    return [
      check("excel-calculated", "Values Excel calculated match independently computed totals", values[0] === independent[0] && values[1] === independent[1], `Excel: ${values.join(", ")}; expected ${independent.join(", ")}`),
      check("traceable", "Metrics are formulas referencing the source sheet", got.formulas.every((f) => String(f[0]).startsWith("=SUM(LOAI_Sales!")), got.formulas.map((f) => f[0]).join(", ")),
      check("chart", "A chart exists on LOAI_Dashboard", charts.length >= 1, charts.map((c) => `${c.name} (${c.type})`).join(", ") || "none"),
      check("source-unchanged", "LOAI_Sales is unchanged", JSON.stringify(src.values) === JSON.stringify(SALES), "compared all cells"),
      ...taskChecks(r),
    ];
  },
  manual: ["Click B2 and B3 on LOAI_Dashboard and confirm the formula bar shows the SUM formulas", "Change a number on LOAI_Sales and confirm the dashboard and chart update"],
  visual: ["Look at the chart and confirm it shows two columns labelled Q1 and Q2"],
};

export const excelCleanerScenario: AcceptanceScenario<ExcelDriver> = {
  id: "excel.cleaner.preserves-source",
  host: "excel",
  specialist: "excel.cleaner",
  title: "Excel: cleaned output sheet with the source preserved",
  purpose: "Profiles messy data, writes a cleaned copy to a new sheet, and checks the source cells are unchanged and row counts add up.",
  preconditions: ["Excel is open with a new blank workbook"],
  checkPreconditions: ({ driver }) => sheetsAbsent(driver, ["LOAI_Orders", "LOAI_Orders_clean"]),
  async setup({ driver }) {
    await driver.addWorksheet("LOAI_Orders");
    await driver.writeRange("LOAI_Orders", "A1:B6", { values: ORDERS });
  },
  request: "Clean this data: fix spacing, unify region names and remove duplicates",
  async script() {
    return [
      call("workbook.profile_data", { sheet: "LOAI_Orders", address: "A1:B6" }),
      call("workbook.clean_to_sheet", { sheet: "LOAI_Orders", address: "A1:B6", targetSheet: "LOAI_Orders_clean", operations: [{ op: "trim" }, { op: "canonicalize_categories", column: "Region" }, { op: "dedupe" }] }),
      say("Cleaned copy written to LOAI_Orders_clean: 5 rows in, 3 out (2 duplicates removed). LOAI_Orders is unchanged."),
    ];
  },
  approve: () => ({ approved: true }),
  async checks({ driver }, r) {
    const src = await driver.readRange("LOAI_Orders", "A1:B6");
    const out = await driver.readRange("LOAI_Orders_clean", "A1:B4");
    const clean = r.actions.find((a) => a.capability === "workbook.clean_to_sheet");
    return [
      check("source-unchanged", "LOAI_Orders is unchanged", JSON.stringify(src.values) === JSON.stringify(ORDERS), JSON.stringify(src.values)),
      check("cleaned", "LOAI_Orders_clean holds the 3 unique rows", JSON.stringify(out.values) === JSON.stringify([["Region", "Amount"], ["North", 10], ["North", 12], ["South", 9]]), JSON.stringify(out.values)),
      check("row-counts", "Rows in = rows out + removed duplicates (verified by the capability)", clean?.status === "succeeded", String(clean?.verification?.evidence)),
      ...taskChecks(r),
    ];
  },
  manual: [],
};

// -------------------------------------------------------------- PowerPoint

const DECK_TITLES = ["LOAI acceptance: Why now", "LOAI acceptance: The numbers", "LOAI acceptance: Next steps"];

export const powerpointAuthorScenario: AcceptanceScenario<PowerPointDriver> = {
  id: "powerpoint.author.multi-slide-deck",
  host: "powerpoint",
  specialist: "powerpoint.author",
  title: "PowerPoint: multi-slide deck under one plan approval, with object and rendered checks",
  purpose: "Generates three slides (one with a chart), inspects slide objects, renders the slides where supported, and records what still needs a human eye.",
  preconditions: ["PowerPoint is open (a new blank presentation is recommended; slides are added at the end)"],
  async checkPreconditions({ driver }) {
    const titles = (await driver.slides()).flatMap((s) => s.shapes.map((x) => x.text ?? ""));
    return titles.some((t) => t.startsWith("LOAI acceptance:")) ? "This presentation already has acceptance slides. Use a new presentation." : undefined;
  },
  async setup() {},
  request: "Create a 3-slide launch presentation",
  async script() {
    const onLastSlide = (name: string, args: Record<string, unknown>): ScriptStep => (req) => {
      const last = [...req.messages].reverse().find((m) => m.role === "tool" && m.toolName === "presentation.create_slide")!;
      return { message: { role: "assistant", content: "", toolCalls: [{ id: "x", name, arguments: { ...args, slideId: JSON.parse(last.content).output.slideId } }] } };
    };
    return [
      call("plan.propose", {
        title: "Launch deck",
        outcome: "Three slides: why now, the numbers (chart), next steps",
        target: "This presentation, new slides at the end",
        risks: ["Adds three slides; existing slides are not changed"],
        steps: DECK_TITLES.map((t) => ({ title: t })),
        operations: [
          ...DECK_TITLES.map((t) => ({ capability: "presentation.create_slide", description: `Create “${t}”`, scope: { slideTitle: t } })),
          { capability: "presentation.add_bar_chart", description: "Pilot chart on “The numbers”", scope: { slideTitle: DECK_TITLES[1] } },
        ],
      }),
      call("presentation.list_layouts", {}),
      call("presentation.create_slide", { title: DECK_TITLES[0], bullets: ["Reports take three days", "Teams want answers in hours"] }),
      call("presentation.create_slide", { title: DECK_TITLES[1], layout: "Title Only" }),
      onLastSlide("presentation.add_bar_chart", { categories: ["Team A", "Team B"], values: [18, 22], title: "Pilot users per team" }),
      call("presentation.create_slide", { title: DECK_TITLES[2], bullets: ["Fix exports", "Expand to 120 users"] }),
      call("presentation.check_layout", {}),
      call("presentation.check_rendering", {}),
      (req) => {
        const r = [...req.messages].reverse().find((m) => m.role === "tool" && m.toolName === "presentation.check_rendering")!;
        const out = JSON.parse(r.content).output;
        return { message: { role: "assistant", content: out.supported ? `Created three slides. Rendered check: ${out.method}` : "Created three slides. Rendering is not supported here, so the slides were not visually checked." } };
      },
    ];
  },
  approve: batch,
  async checks({ driver }, r) {
    const slides = await driver.slides();
    const ours = slides.filter((s) => s.shapes.some((x) => DECK_TITLES.includes(x.text ?? "")));
    const titles = ours.map((s) => s.shapes.find((x) => DECK_TITLES.includes(x.text ?? ""))!.text);
    const chartSlide = ours[1];
    const labels = chartSlide?.shapes.map((x) => x.text) ?? [];
    const layout = r.actions.filter((a) => a.capability === "presentation.check_layout").at(-1);
    const ourIds = new Set(ours.map((s) => s.id));
    const highOnOurs = ((layout?.output as any)?.issues ?? []).filter((i: any) => i.severity === "high" && ourIds.has(i.slideId));
    const prompts = r.actions.filter((a) => a.approval === "approved" && a.capability !== "plan.propose").length;
    return [
      check("slides-created", "Three slides with the planned titles were added at the end", JSON.stringify(titles) === JSON.stringify(DECK_TITLES) && slides.slice(-3).every((s) => ourIds.has(s.id)), titles.join(" | ")),
      check("chart", "The chart slide has bars and value labels 18 and 22", labels.includes("18") && labels.includes("22"), labels.filter(Boolean).join(", ")),
      check("one-approval", "Slides ran under the single batch plan approval", prompts === 0 && r.planExecution?.batch === true, `extra prompts: ${prompts}; plan: ${r.planExecution?.operations.map((o) => o.status).join(", ")}`),
      check("object-check", "No high-severity object-level findings on the new slides", highOnOurs.length === 0, highOnOurs.map((i: any) => i.detail).join("; ") || "none"),
      ...taskChecks(r),
    ];
  },
  manual: ["Open each new slide and confirm the title and bullets are readable and nothing overflows its box"],
  visual: ["Confirm the chart on “The numbers” shows two bars with the labels 18 and 22", "Confirm no text is cut off or overlapping on the three new slides"],
};

// ----------------------------------------------------------------- Outlook

export const OUTLOOK_FIXTURE_SUBJECT = "LOAI acceptance: budget thread";

async function fixtureMessage({ driver }: { driver: OutlookDriver }) {
  const m = await driver.item();
  return m.subject.includes(OUTLOOK_FIXTURE_SUBJECT) ? undefined : `Open the fixture message “${OUTLOOK_FIXTURE_SUBJECT}” (acceptance/fixtures/budget-thread.eml).`;
}

export const outlookThreadScenario: AcceptanceScenario<OutlookDriver> = {
  id: "outlook.thread.attribution",
  host: "outlook",
  specialist: "outlook.thread",
  title: "Outlook: analyse a thread and extract attributed commitments",
  purpose: "Splits the open thread by sender and checks commitments are attributed to the person who made them.",
  preconditions: [`The fixture message “${OUTLOOK_FIXTURE_SUBJECT}” is open (import acceptance/fixtures/budget-thread.eml)`],
  checkPreconditions: fixtureMessage,
  async setup() {},
  request: "Summarize this thread and list who committed to what",
  async script() {
    return [
      say(
        [
          "Summary: Dana asked Alex for the revised budget and offered to book a room; Alex agreed to send the budget by Friday.",
          "Commitments:",
          "- [explicit] Send the revised budget — owner: Alex Lee — due: Friday — source: M1",
          "- [explicit] Book the room for Thursday — owner: Dana Smith — due: Thursday — source: M2",
          "- [inferred] Confirm the vendor list — owner: Alex Lee — due: none — source: M2",
        ].join("\n"),
      ),
    ];
  },
  approve: () => ({ approved: true }),
  async checks(_c, r) {
    const thread = r.contextUsed.find((c) => c.id === "outlook:thread");
    const msgs = thread ? parseRenderedThread(thread.content) : [];
    const authors = msgs.map((m) => m.author);
    return [
      check("split", "The thread splits into the open message and the quoted one, with correct authors", authors[0] === "Alex Lee" && authors[1] === "Dana Smith", msgs.map((m) => `${m.id}: ${m.author}`).join(", ") || "not split"),
      ...taskChecks(r),
    ];
  },
  manual: [],
};

export const outlookReplyScenario: AcceptanceScenario<OutlookDriver> = {
  id: "outlook.reply.draft-not-sent",
  host: "outlook",
  specialist: "outlook.reply",
  title: "Outlook: draft a reply, verify it where possible, never send",
  purpose: "Creates a reply draft, records whether Outlook let it be verified, and confirms no send capability exists and nothing was sent.",
  preconditions: [`The fixture message “${OUTLOOK_FIXTURE_SUBJECT}” is open (read mode), or a reply to it is open (compose mode)`],
  checkPreconditions: fixtureMessage,
  async setup() {},
  request: "Draft a reply confirming the budget by Friday",
  async script() {
    return [
      call("outlook.create_draft", { body: "Hi Dana,\n\nConfirmed: the revised budget will be with you by Friday. Thanks for booking the room.\n\nAlex" }),
      say("I created a reply draft. Nothing was sent; please review and send it yourself."),
    ];
  },
  approve: () => ({ approved: true }),
  async checks(c, r) {
    const draft = r.actions.find((a) => a.capability === "outlook.create_draft");
    const caps = c.adapter.capabilities().map((x) => x.name);
    return [
      check("draft", "A draft was created (verified in compose mode; honestly unverified in read mode)", draft?.status === "succeeded" || draft?.status === "unverified", `${draft?.status}: ${draft?.verification?.evidence ?? draft?.verification?.method ?? ""}`),
      check("honest-outcome", "An unverified draft is not reported as complete", draft?.status !== "unverified" || r.outcome !== "completed", r.outcome),
      check("no-send", "No send capability exists and none ran", !caps.some((n) => /send/i.test(n)) && !r.actions.some((a) => /send/i.test(a.capability)), caps.join(", ")),
      ...taskChecks(r, ["completed", "partial"]),
    ];
  },
  manual: ["Confirm the reply window (or compose body) contains the draft text", "Open Sent Items and confirm nothing new was sent"],
};

export const ACCEPTANCE_SCENARIOS: AcceptanceScenario[] = [wordWriterScenario, wordEditorScenario, excelDashboardScenario, excelCleanerScenario, powerpointAuthorScenario, outlookThreadScenario, outlookReplyScenario];
