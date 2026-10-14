/**
 * Bounded plan approval and execution (hardening PR §1) and its regression gates.
 */
import { describe, expect, it } from "vitest";
import {
  DEFAULT_POLICY,
  ExcelAdapter,
  FakeExcelDriver,
  FakePowerPointDriver,
  FakeWordDriver,
  PowerPointAdapter,
  ScriptedProvider,
  WordAdapter,
  matchesOperation,
  type ApprovalRequest,
  type ScriptStep,
} from "../src/index.js";
import { setup } from "./helpers.js";

const call = (name: string, args: Record<string, unknown>, text?: string) => ({ call: { name, arguments: args }, text });
const batch = (r: ApprovalRequest) => (r.capability === "plan.propose" ? { approved: true, batch: true } : { approved: true });
const deckPlan = (titles: string[], extra: object[] = []) =>
  call("plan.propose", {
    title: "Launch deck",
    outcome: `A ${titles.length}-slide launch presentation`,
    target: "This presentation (new slides at the end)",
    risks: ["Adds slides; existing slides are not changed"],
    steps: titles.map((t) => ({ title: t })),
    operations: [...titles.map((t) => ({ capability: "presentation.create_slide", description: `Create slide “${t}”`, scope: { slideTitle: t } })), ...extra],
  });
/** A scripted step that uses the slideId returned by the most recent create_slide. */
const onLastSlide = (name: string, args: Record<string, unknown>): ScriptStep => (req) => {
  const last = [...req.messages].reverse().find((m) => m.role === "tool" && m.toolName === "presentation.create_slide")!;
  const slideId = JSON.parse(last.content).output.slideId;
  return { message: { role: "assistant", content: "", toolCalls: [{ id: "x", name, arguments: { ...args, slideId } }] } };
};
const deckRequest = "Create a 3-slide launch presentation";
const author = (s: ReturnType<typeof setup>) => s.specialists.get("powerpoint.author")!;

describe("bounded batch execution of an approved plan", () => {
  it("runs a multi-slide deck under one plan approval; every write is still validated and verified", async () => {
    const ppt = new FakePowerPointDriver([]);
    const model = new ScriptedProvider([
      deckPlan(["Why now", "The numbers", "Next steps"], [{ capability: "presentation.add_bar_chart", description: "Revenue chart on “The numbers”", scope: { slideTitle: "The numbers" } }]),
      call("presentation.create_slide", { title: "Why now", bullets: ["Demand doubled"] }),
      call("presentation.create_slide", { title: "The numbers", layout: "Title Only" }),
      onLastSlide("presentation.add_bar_chart", { categories: ["Q1", "Q2"], values: [10, 20] }),
      call("presentation.create_slide", { title: "Next steps", bullets: ["Hire", "Launch"] }),
      call("presentation.check_layout", {}),
      call("presentation.check_rendering", {}),
      { text: "Created 3 slides with a chart. No serious layout issues found." },
    ]);
    const s = setup(new PowerPointAdapter(ppt), model, { approve: batch });
    const r = await s.runtime.run(deckRequest, { specialist: author(s) });

    expect(s.approver.requests.map((q) => q.capability)).toEqual(["plan.propose"]);
    expect(s.approver.requests[0]!.planBatch).toEqual({ operations: 4, autoEligible: 4, alwaysAsk: 0 });
    const writes = r.actions.filter((a) => a.capability.startsWith("presentation.") && !["presentation.check_layout", "presentation.check_rendering", "presentation.read_slides"].includes(a.capability));
    expect(writes.map((a) => [a.status, a.approval])).toEqual(Array(4).fill(["succeeded", "plan"]));
    expect(writes.every((a) => a.verification?.verified === true)).toBe(true);
    expect(ppt.slidesList).toHaveLength(3);
    expect(r.planExecution).toMatchObject({ batch: true, suspended: undefined, undeclared: [] });
    expect(r.planExecution!.operations.map((o) => o.status)).toEqual(["done", "done", "done", "done"]);
    expect(r.outcome).toBe("completed");
    expect(r.executionSummary).toMatch(/4 of 4 planned change\(s\) completed \(approved as a batch\)/);
  });

  it("stops batch execution on a scope change; the undeclared change and every later one ask", async () => {
    const ppt = new FakePowerPointDriver([]);
    const model = new ScriptedProvider([
      deckPlan(["One", "Two"]),
      call("presentation.create_slide", { title: "One" }),
      call("presentation.create_slide", { title: "Surprise slide" }),
      call("presentation.create_slide", { title: "Two" }),
      call("presentation.check_layout", {}),
      { text: "Done." },
    ]);
    const s = setup(new PowerPointAdapter(ppt), model, { approve: batch });
    const r = await s.runtime.run("Create a 2-slide deck", { specialist: author(s) });
    const asked = s.approver.requests.map((q) => `${q.capability}:${(q.args as any)?.title ?? ""}`);
    expect(asked).toEqual(["plan.propose:Launch deck", "presentation.create_slide:Surprise slide", "presentation.create_slide:Two"]);
    expect(s.approver.requests[1]!.preview.warnings).toContain("This change is not part of the approved plan.");
    expect(r.planExecution!.suspended).toMatch(/not part of the approved plan/);
    expect(r.planExecution!.undeclared).toHaveLength(1);
    expect(r.planExecution!.operations.map((o) => o.calls[0]!.approvedBy)).toEqual(["plan", "user"]);
  });

  it("stops batch execution when verification fails and reports the step as failed", async () => {
    const doc = new FakeWordDriver(["Intro"]);
    const model = new ScriptedProvider([
      call("plan.propose", {
        title: "Report",
        outcome: "Two sections appended",
        target: "This document, end",
        steps: [{ title: "A" }, { title: "B" }],
        operations: [{ capability: "document.write_blocks", description: "Append both sections", scope: { position: "end" }, count: 2 }],
      }),
      () => {
        doc.dropNextWrite = true;
        return { message: { role: "assistant", content: "", toolCalls: [{ id: "1", name: "document.write_blocks", arguments: { mode: "insert", blocks: [{ text: "A", style: "Heading1" }, { text: "a" }] } }] } };
      },
      call("document.write_blocks", { mode: "insert", blocks: [{ text: "B", style: "Heading1" }, { text: "b" }] }),
      { text: "Both sections were added successfully." },
    ]);
    const s = setup(new WordAdapter(doc), model, { approve: batch });
    const r = await s.runtime.run("Write a two-section report", { specialist: s.specialists.get("word.writer")! });
    expect(s.approver.requests.map((q) => q.capability)).toEqual(["plan.propose", "document.write_blocks"]);
    expect(r.planExecution!.suspended).toMatch(/failed/);
    expect(r.planExecution!.operations[0]!.status).toBe("failed");
    expect(r.outcome).not.toBe("completed");
  });

  it("never auto-approves an operation that policy says must be confirmed", async () => {
    const big = "word ".repeat(400).trim();
    const doc = new FakeWordDriver(["Title", big]);
    const model = new ScriptedProvider([
      call("plan.propose", { title: "Rewrite", outcome: "Shorter section", target: "Paragraph 1", steps: [{ title: "Rewrite ¶1" }], operations: [{ capability: "document.replace_text", description: "Replace ¶1", scope: { paragraphs: { start: 1, end: 1 } } }] }),
      call("document.replace_text", { target: "paragraphs", start: 1, expectedText: big, newText: "Short." }),
      { text: "ok" },
    ]);
    const s = setup(new WordAdapter(doc), model, { approve: batch });
    const r = await s.runtime.run("Rewrite paragraph 1 as a structured document section", { specialist: s.specialists.get("word.writer")! });
    const q = s.approver.requests.find((x) => x.capability === "document.replace_text")!;
    expect(q.decision).toBe("confirm"); // escalated to destructive at execution time
    expect(r.actions.find((a) => a.capability === "document.replace_text")!.approval).toBe("approved");
  });

  it("counts policy-confirmed operations as always asking in the batch offer", async () => {
    const ppt = new FakePowerPointDriver([]);
    const s = setup(new PowerPointAdapter(ppt), new ScriptedProvider([deckPlan(["One"]), { text: "?" }]), {
      approve: batch,
      settings: { policy: { ...DEFAULT_POLICY, modify: "confirm" } },
    });
    await s.runtime.run("Create a 1-slide deck", { specialist: author(s) });
    expect(s.approver.requests[0]!.planBatch).toEqual({ operations: 1, autoEligible: 0, alwaysAsk: 1 });
  });

  it("stepwise approval (no batch) asks for every change and still tracks plan steps", async () => {
    const ppt = new FakePowerPointDriver([]);
    const model = new ScriptedProvider([deckPlan(["One", "Two"]), call("presentation.create_slide", { title: "One" }), call("presentation.create_slide", { title: "Two" }), call("presentation.check_layout", {}), { text: "Created two slides." }]);
    const s = setup(new PowerPointAdapter(ppt), model, { approve: true });
    const r = await s.runtime.run("Create a 2-slide deck", { specialist: author(s) });
    expect(s.approver.requests).toHaveLength(3);
    expect(r.planExecution).toMatchObject({ batch: false });
    expect(r.planExecution!.operations.map((o) => [o.status, o.calls[0]!.approvedBy])).toEqual([
      ["done", "user"],
      ["done", "user"],
    ]);
  });

  it("is disabled entirely by the planBatchApproval setting", async () => {
    const ppt = new FakePowerPointDriver([]);
    const model = new ScriptedProvider([deckPlan(["One"]), call("presentation.create_slide", { title: "One" }), { text: "ok" }]);
    const s = setup(new PowerPointAdapter(ppt), model, { approve: batch, settings: { planBatchApproval: false } });
    const r = await s.runtime.run("Create a 1-slide deck", { specialist: author(s) });
    expect(s.approver.requests[0]!.planBatch).toBeUndefined();
    expect(s.approver.requests.map((q) => q.capability)).toEqual(["plan.propose", "presentation.create_slide"]);
    expect(r.planExecution!.batch).toBe(false);
  });

  it("reports unexecuted plan operations as partial completion", async () => {
    const ppt = new FakePowerPointDriver([]);
    const model = new ScriptedProvider([deckPlan(["One", "Two", "Three"]), call("presentation.create_slide", { title: "One" }), call("presentation.check_layout", {}), { text: "All three slides are done!" }]);
    const s = setup(new PowerPointAdapter(ppt), model, { approve: batch });
    const r = await s.runtime.run("Create a 1-slide deck", { specialist: author(s) });
    expect(r.planExecution!.operations.map((o) => o.status)).toEqual(["done", "not_run", "not_run"]);
    expect(r.outcome).toBe("partial");
    expect(r.executionSummary).toMatch(/1 of 3 planned change/);
  });

  it("a write outside the declared Excel range asks and stops the batch", async () => {
    const xl = new FakeExcelDriver({ Data: [["a", 1]] });
    const model = new ScriptedProvider([
      call("plan.propose", {
        title: "Summary",
        outcome: "Totals on a new sheet",
        target: "Workbook, new sheet Summary",
        steps: [{ title: "Sheet" }, { title: "Totals" }],
        operations: [
          { capability: "workbook.add_worksheet", description: "Add sheet Summary", scope: { sheet: "Summary" } },
          { capability: "workbook.write_range", description: "Totals in Summary!A1:B2", scope: { sheet: "Summary", range: "A1:B2" } },
        ],
      }),
      call("workbook.add_worksheet", { name: "Summary" }),
      call("workbook.write_range", { sheet: "Summary", address: "A1:B3", formulas: [["x", "=SUM(Data!B1:B1)"], ["y", "1"], ["z", "2"]] }),
      { text: "ok" },
    ]);
    const s = setup(new ExcelAdapter(xl), model, { approve: batch });
    const r = await s.runtime.run("Create a summary dashboard sheet", { specialist: s.specialists.get("excel.dashboard")! });
    expect(s.approver.requests.map((q) => q.capability)).toEqual(["plan.propose", "workbook.write_range"]);
    expect(r.actions.find((a) => a.capability === "workbook.add_worksheet")!.approval).toBe("plan");
    expect(r.planExecution!.suspended).toBeDefined();
  });
});

describe("plan validation and forgery resistance", () => {
  const badPlan = (operations: object[]) =>
    new ScriptedProvider([call("plan.propose", { title: "P", outcome: "o", target: "t", steps: [{ title: "s" }], operations }), { text: "ok" }]);

  it.each([
    [[{ capability: "document.delete_paragraphs", description: "x", scope: {} }], /not an available capability/],
    [[{ capability: "shell.exec", description: "x", scope: {} }], /not an available capability/],
    [[{ capability: "context.read", description: "x", scope: {} }], /only reads/],
    [[{ capability: "document.replace_text", description: "x", scope: {} }], /needs scope paragraphs/],
    [[{ capability: "document.insert_text", description: "x", scope: { position: "end" } }], /not available in this task/],
  ])("rejects a plan declaring %j before asking the user", async (ops, msg) => {
    const s = setup(new WordAdapter(new FakeWordDriver(["a"])), badPlan(ops), { approve: batch });
    const r = await s.runtime.run("Write a report", { specialist: s.specialists.get("word.editor")! });
    const rec = r.actions.find((a) => a.capability === "plan.propose")!;
    expect(rec.status).toBe("rejected");
    expect(rec.error).toMatch(msg);
    expect(s.approver.requests).toHaveLength(0);
  });

  it("rejects operations for another Office host", async () => {
    const s = setup(new WordAdapter(new FakeWordDriver(["a"])), badPlan([{ capability: "workbook.write_range", description: "x", scope: { sheet: "S", range: "A1" } }]));
    const r = await s.runtime.run("anything");
    expect(r.actions[0]!.error).toMatch(/not an available capability/);
  });

  it("an approver response cannot forge plan coverage", async () => {
    const doc = new FakeWordDriver(["a"]);
    const s = setup(new WordAdapter(doc), new ScriptedProvider([call("document.insert_text", { position: "end", text: "b" }), { text: "ok" }]), {
      approve: () => ({ approved: true, via: "plan" }),
    });
    const r = await s.runtime.run("add b");
    expect(r.actions[0]!.approval).toBe("approved");
  });

  it("matches scope precisely", () => {
    const op = { capability: "workbook.write_range", description: "", scope: { sheet: "S", range: "B2:C5" }, count: 1 };
    const w = (address: string, sheet = "S") => ({ id: "x", name: "workbook.write_range", arguments: { sheet, address } });
    expect(matchesOperation(op, w("B2:C5"), [])).toBe(true);
    expect(matchesOperation(op, w("C3"), [])).toBe(true);
    expect(matchesOperation(op, w("B2:D5"), [])).toBe(false);
    expect(matchesOperation(op, w("B2", "Other"), [])).toBe(false);
    const slide = { capability: "presentation.add_diagram", description: "", scope: { slideTitle: "Plan" }, count: 1 };
    expect(matchesOperation(slide, { id: "x", name: "presentation.add_diagram", arguments: { slideId: "s1" } }, [])).toBe(false); // no slide titled "Plan" created in this task
  });
});
