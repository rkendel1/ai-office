/**
 * End-to-end acceptance scenarios from spec §10, run against in-memory Office
 * hosts and a scripted model. The Office.js drivers implement the same
 * interfaces; see docs/testing.md for running these flows in real Office.
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
} from "../src/index.js";
import { setup } from "./helpers.js";

describe("Word: read, answer grounded, propose rewrite, approve, apply, verify", () => {
  it("completes the full workflow", async () => {
    const doc = new FakeWordDriver([
      "Quarterly Report",
      "Revenue grew 12% year over year, driven by subscriptions.",
      "Costs were flat. We plan to hire two engineers.",
    ]);
    const adapter = new WordAdapter(doc);
    const model = new ScriptedProvider([
      { call: { name: "context.read", arguments: {} }, text: "I'll read the document." },
      { text: "Revenue grew 12% year over year (paragraph 1)." },
    ]);
    const s = setup(adapter, model);

    const q = await s.runtime.run("How much did revenue grow?");
    expect(q.outcome).toBe("completed");
    expect(q.text).toContain("12%");
    // Model saw the actual document content on the second call.
    const second = JSON.stringify(model.requests[1]!.messages);
    expect(second).toContain("Revenue grew 12% year over year");
    expect(q.contextUsed.some((c) => c.scope === "full")).toBe(true);

    const model2 = new ScriptedProvider([
      {
        call: {
          name: "document.replace_text",
          arguments: {
            target: "paragraphs",
            start: 2,
            expectedText: "Costs were flat. We plan to hire two engineers.",
            newText: "Costs held steady, and we plan to hire two engineers next quarter.",
          },
        },
      },
      { text: "Rewrote paragraph 2." },
    ]);
    const s2 = setup(adapter, model2);
    const r = await s2.runtime.run("Make paragraph 2 more polished");
    expect(s2.approver.requests).toHaveLength(1);
    expect(s2.approver.requests[0]!.preview.before).toContain("Costs were flat");
    expect(s2.approver.requests[0]!.preview.after).toContain("held steady");
    expect(r.actions[0]!.status).toBe("succeeded");
    expect(r.actions[0]!.verification?.verified).toBe(true);
    expect(r.outcome).toBe("completed");
    expect(doc.paragraphs[2]!.text).toBe("Costs held steady, and we plan to hire two engineers next quarter.");
    expect(r.actions[0]!.compensation?.args.newText).toBe("Costs were flat. We plan to hire two engineers.");

    // Undo via compensation goes through approval again.
    const undo = await s2.runtime.executeDirect({ id: "u", name: r.actions[0]!.compensation!.capability, arguments: r.actions[0]!.compensation!.args });
    expect(undo.status).toBe("succeeded");
    expect(doc.paragraphs[2]!.text).toBe("Costs were flat. We plan to hire two engineers.");
  });

  it("rejecting a proposal leaves the document unchanged", async () => {
    const doc = new FakeWordDriver(["Alpha", "Beta"]);
    const model = new ScriptedProvider([
      { call: { name: "document.replace_text", arguments: { target: "paragraphs", start: 0, expectedText: "Alpha", newText: "Gamma" } } },
      { call: { name: "document.replace_text", arguments: { target: "paragraphs", start: 0, expectedText: "Alpha", newText: "Gamma" } } },
      { text: "OK, I won't change it." },
    ]);
    const s = setup(new WordAdapter(doc), model, { approve: false });
    const r = await s.runtime.run("change alpha");
    expect(doc.paragraphs.map((p) => p.text)).toEqual(["Alpha", "Beta"]);
    expect(r.actions[0]!.approval).toBe("rejected");
    // Re-proposing the identical change is blocked without asking again.
    expect(s.approver.requests).toHaveLength(1);
    expect(r.actions[1]!.error).toMatch(/already rejected/);
    expect(r.executionSummary).toMatch(/Rejected by you/);
  });

  it("refuses stale edits when the document changed after reading", async () => {
    const doc = new FakeWordDriver(["Old text"]);
    const model = new ScriptedProvider([
      { call: { name: "document.replace_text", arguments: { target: "paragraphs", start: 0, expectedText: "Something else", newText: "New" } } },
      { text: "The paragraph changed; I did not edit it." },
    ]);
    const s = setup(new WordAdapter(doc), model);
    const r = await s.runtime.run("edit");
    expect(r.actions[0]!.status).toBe("failed");
    expect(doc.paragraphs[0]!.text).toBe("Old text");
    expect(r.outcome).toBe("failed");
  });

  it("reports failure when read-back verification does not match, even if the model claims success", async () => {
    const doc = new FakeWordDriver(["Hello"]);
    doc.dropNextWrite = true;
    const model = new ScriptedProvider([
      { call: { name: "document.insert_text", arguments: { position: "end", text: "World" } } },
      { text: "Done! I added 'World' to the end." },
    ]);
    const s = setup(new WordAdapter(doc), model);
    const r = await s.runtime.run("add World");
    expect(r.actions[0]!.status).toBe("failed");
    expect(r.outcome).toBe("failed");
    expect(r.executionSummary).toMatch(/Failed/);
  });

  it("distinguishes unsupported content from an empty document", async () => {
    const doc = new FakeWordDriver([]);
    doc.unsupported = [{ kind: "embedded objects", count: 2 }];
    const model = new ScriptedProvider([{ call: { name: "context.read", arguments: {} } }, { text: "It contains only embedded objects." }]);
    const s = setup(new WordAdapter(doc), model);
    const r = await s.runtime.run("summarize");
    const item = r.contextUsed.find((c) => c.id === "word:paragraphs:empty")!;
    expect(item.content).toMatch(/cannot be read/);
    expect(item.unsupported[0]!.count).toBe(2);
  });

  it("uses the selection without presenting it as the whole document", async () => {
    const doc = new FakeWordDriver(["A", "Selected sentence.", "C"]);
    doc.selectionRange = { start: 1, end: 1 };
    const model = new ScriptedProvider([{ text: "The selection says: Selected sentence." }]);
    const s = setup(new WordAdapter(doc), model);
    const r = await s.runtime.run("what does this say");
    const sel = r.contextUsed.find((c) => c.scope === "selection")!;
    expect(sel.label).toMatch(/Selection/);
    expect(r.contextUsed.some((c) => c.scope === "full")).toBe(false);
  });

  it("escalates substantial overwrites to explicit confirmation", async () => {
    const big = "word ".repeat(400).trim();
    const doc = new FakeWordDriver([big]);
    const model = new ScriptedProvider([
      { call: { name: "document.replace_text", arguments: { target: "paragraphs", start: 0, expectedText: big, newText: "short" } } },
      { text: "done" },
    ]);
    const s = setup(new WordAdapter(doc), model);
    await s.runtime.run("shorten");
    expect(s.approver.requests[0]!.risk).toBe("destructive");
    expect(s.approver.requests[0]!.decision).toBe("confirm");
  });
});

describe("Excel: inspect, summarize, new sheet + chart, verify", () => {
  it("completes the workflow", async () => {
    const xl = new FakeExcelDriver({
      Sales: [
        ["Region", "Q1", "Q2"],
        ["North", 100, 120],
        ["South", 80, 95],
        ["West", 60, 70],
      ],
    });
    const model = new ScriptedProvider([
      { call: { name: "workbook.read_range", arguments: { sheet: "Sales" } } },
      { call: { name: "workbook.add_worksheet", arguments: { name: "Summary" } } },
      {
        call: {
          name: "workbook.write_range",
          arguments: { sheet: "Summary", address: "A1:B3", formulas: [["Quarter", "Total"], ["Q1", "=SUM(Sales!B2:B4)"], ["Q2", "=SUM(Sales!C2:C4)"]] },
        },
      },
      { call: { name: "workbook.create_chart", arguments: { sourceSheet: "Summary", sourceAddress: "A1:B3", kind: "column", title: "Totals by quarter" } } },
      { text: "Created Summary with totals and a chart." },
    ]);
    const s = setup(new ExcelAdapter(xl), model);
    const r = await s.runtime.run("Summarize totals per quarter on a new sheet with a chart");
    expect(r.actions.map((a) => a.status)).toEqual(["succeeded", "succeeded", "succeeded", "succeeded"]);
    expect(r.outcome).toBe("completed");
    expect(xl.chartsList).toHaveLength(1);
    // three modifications → three approvals; the read needed none
    expect(s.approver.requests.map((q) => q.capability)).toEqual(["workbook.add_worksheet", "workbook.write_range", "workbook.create_chart"]);
    expect(JSON.stringify(model.requests[1]!.messages)).toContain("North");
  });

  it("verifies computed values and shows them with their formulas", async () => {
    const xl = new FakeExcelDriver({ S: [[1], [2], [3]] });
    const model = new ScriptedProvider([
      { call: { name: "workbook.write_range", arguments: { sheet: "S", address: "A4", formulas: [["=SUM(A1:A3)"]] } } },
      { text: "ok" },
    ]);
    const s = setup(new ExcelAdapter(xl), model);
    const r = await s.runtime.run("total");
    expect(r.actions[0]!.status).toBe("succeeded");
    expect((r.actions[0]!.output as any).resultPreview).toContain("6 [=SUM(A1:A3)]");
  });

  it("rejects writes whose shape does not match the range, and unknown sheets", async () => {
    const xl = new FakeExcelDriver({ S: [[1]] });
    const model = new ScriptedProvider([
      { call: { name: "workbook.write_range", arguments: { sheet: "S", address: "A1:B2", values: [[1]] } } },
      { call: { name: "workbook.read_range", arguments: { sheet: "Payroll", address: "A1" } } },
      { text: "Could not do it." },
    ]);
    const s = setup(new ExcelAdapter(xl), model);
    const r = await s.runtime.run("x");
    expect(r.actions[0]!.error).toMatch(/1×1 but A1:B2 is 2×2/);
    expect(r.actions[1]!.error).toMatch(/does not exist/);
    expect(s.approver.requests).toHaveLength(0);
  });

  it("requires confirmation to overwrite substantial data and offers restore", async () => {
    const rows = Array.from({ length: 30 }, (_, i) => [i]);
    const xl = new FakeExcelDriver({ S: rows });
    const model = new ScriptedProvider([
      { call: { name: "workbook.write_range", arguments: { sheet: "S", address: "A1:A30", values: rows.map(() => [0]) } } },
      { text: "zeroed" },
    ]);
    const s = setup(new ExcelAdapter(xl), model);
    const r = await s.runtime.run("zero it");
    expect(s.approver.requests[0]!.decision).toBe("confirm");
    expect(s.approver.requests[0]!.preview.warnings?.[0]).toMatch(/Overwrites 30/);
    const undo = await s.runtime.executeDirect({ id: "u", name: "workbook.write_range", arguments: r.actions[0]!.compensation!.args });
    expect(undo.status).toBe("succeeded");
    expect((await xl.readRange("S", "A30")).values[0]![0]).toBe(29);
  });
});

describe("PowerPoint: read deck, summarize, edit selected slide text, verify", () => {
  it("completes the workflow", async () => {
    const ppt = new FakePowerPointDriver([
      { shapes: [{ name: "Title", text: "Why local AI" }, { name: "Logo", text: null }] },
      { shapes: [{ name: "Title", text: "Privacy" }, { name: "Body", text: "Data stays on device" }] },
    ]);
    ppt.selected = ["s2"];
    const model = new ScriptedProvider([
      { call: { name: "presentation.read_slides", arguments: {} } },
      { text: "The deck argues local AI protects privacy." },
    ]);
    const s = setup(new PowerPointAdapter(ppt), model);
    const r1 = await s.runtime.run("Summarize the deck");
    expect(r1.contextUsed.find((c) => c.id === "ppt:slides:all")?.unsupported[0]?.count).toBe(1);

    const model2 = new ScriptedProvider([
      { call: { name: "presentation.edit_text", arguments: { slideId: "s2", shapeId: "s2sh2", text: "Your data never leaves your device", expectedText: "Data stays on device" } } },
      { text: "Updated slide 2." },
    ]);
    const s2 = setup(new PowerPointAdapter(ppt), model2);
    const r2 = await s2.runtime.run("Make the body on the selected slide punchier");
    expect(r2.actions[0]!.status).toBe("succeeded");
    expect(ppt.slidesList[1]!.shapes[1]!.text).toBe("Your data never leaves your device");
    // selection is in the initial context
    expect(JSON.stringify(model2.requests[0]!.messages)).toContain("Data stays on device");
  });

  it("refuses to edit shapes without text frames", async () => {
    const ppt = new FakePowerPointDriver([{ shapes: [{ name: "Pic", text: null }] }]);
    const model = new ScriptedProvider([{ call: { name: "presentation.edit_text", arguments: { slideId: "s1", shapeId: "s1sh1", text: "x" } } }, { text: "can't" }]);
    const s = setup(new PowerPointAdapter(ppt), model);
    const r = await s.runtime.run("x");
    expect(r.actions[0]!.error).toMatch(/no text frame/);
  });
});

describe("Outlook: summarize thread, draft reply, never send", () => {
  const msg = {
    itemType: "message" as const,
    subject: "Project kickoff",
    from: "Dana <dana@example.com>",
    to: ["me@example.com"],
    cc: [],
    body: "Can you confirm Thursday at 10? Please bring the budget.",
    attachments: [{ name: "agenda.pdf", size: 1000 }],
  };

  it("read mode: opens a reply draft, reports it honestly as unverified, sends nothing", async () => {
    const ol = new FakeOutlookDriver(msg, "read");
    const model = new ScriptedProvider([
      { call: { name: "outlook.read_message", arguments: {} } },
      { call: { name: "outlook.create_draft", arguments: { body: "Thursday at 10 works. I'll bring the budget." } } },
      { text: "Summary: Dana asks to confirm Thursday 10am and bring the budget. I opened a reply draft." },
    ]);
    const s = setup(new OutlookAdapter(ol), model);
    const r = await s.runtime.run("Summarize and draft a reply");
    expect(ol.openedForms).toHaveLength(1);
    expect(r.actions[1]!.status).toBe("unverified");
    expect(r.outcome).toBe("partial");
    expect(ol.sent).toHaveLength(0);
    expect(s.registry.list("outlook").some((c) => /send/.test(c.name))).toBe(false);
    expect(s.approver.requests[0]!.preview.warnings?.[0]).toMatch(/Nothing will be sent/);
  });

  it("compose mode: inserts and verifies the draft text", async () => {
    const ol = new FakeOutlookDriver(msg, "compose");
    const model = new ScriptedProvider([{ call: { name: "outlook.create_draft", arguments: { body: "Confirmed." } } }, { text: "Inserted." }]);
    const s = setup(new OutlookAdapter(ol), model);
    const r = await s.runtime.run("draft");
    expect(r.actions[0]!.status).toBe("succeeded");
    expect(ol.composeText).toContain("Confirmed.");
  });

  it("refuses a model attempt to send", async () => {
    const ol = new FakeOutlookDriver(msg, "read");
    const model = new ScriptedProvider([{ call: { name: "outlook.send", arguments: { to: "x@evil.test" } } }, { text: "I cannot send." }]);
    const s = setup(new OutlookAdapter(ol), model);
    const r = await s.runtime.run("send it");
    expect(r.actions[0]!.status).toBe("rejected");
    expect(r.actions[0]!.error).toMatch(/Unknown capability/);
  });
});
