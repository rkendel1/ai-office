/**
 * Specialist unit and routing tests (PR spec §9: specialist unit tests, routing tests).
 */
import { describe, expect, it } from "vitest";
import {
  ALL_SPECIALISTS,
  ExcelAdapter,
  FakeExcelDriver,
  FakeOutlookDriver,
  FakePowerPointDriver,
  FakeWordDriver,
  OutlookAdapter,
  PowerPointAdapter,
  ScriptedProvider,
  SpecialistRegistry,
  SpecialistRouter,
  WordAdapter,
  createSession,
  type HostAdapter,
  type HostKind,
} from "../src/index.js";
import { setup, settings } from "./helpers.js";

const adapters: Record<HostKind, () => HostAdapter> = {
  word: () => new WordAdapter(new FakeWordDriver(["a"])),
  excel: () => new ExcelAdapter(new FakeExcelDriver({ S: [[1]] })),
  powerpoint: () => new PowerPointAdapter(new FakePowerPointDriver([{ shapes: [{ name: "Title", text: "x" }] }])),
  outlook: () => new OutlookAdapter(new FakeOutlookDriver({ itemType: "message", subject: "s", from: "a@b.c", to: [], cc: [], body: "hi", attachments: [] })),
};
const sessionFor = (host: HostKind) => createSession({ adapter: adapters[host](), provider: new ScriptedProvider([]), approver: async () => ({ approved: true }), settings: () => settings() });

describe("specialist registration", () => {
  const reg = new SpecialistRegistry();

  it("registers concrete specialists for every Office host", () => {
    expect(ALL_SPECIALISTS).toHaveLength(14);
    expect(reg.forHost("word").map((s) => s.id)).toEqual(["word.writer", "word.editor", "word.reviewer", "word.summarizer"]);
    expect(reg.forHost("excel").map((s) => s.id)).toEqual(["excel.analyst", "excel.formula", "excel.cleaner", "excel.dashboard"]);
    expect(reg.forHost("powerpoint").map((s) => s.id)).toEqual(["powerpoint.author", "powerpoint.designer", "powerpoint.reviewer"]);
    expect(reg.forHost("outlook").map((s) => s.id)).toEqual(["outlook.thread", "outlook.reply", "outlook.actions"]);
  });

  it("rejects duplicate ids and ids that do not match the host", () => {
    const s = ALL_SPECIALISTS[0]!;
    expect(() => new SpecialistRegistry([s, s])).toThrow(/Duplicate/);
    expect(() => new SpecialistRegistry([{ ...s, id: "excel.fake" }])).toThrow(/must start with its host/);
  });

  for (const s of ALL_SPECIALISTS) {
    describe(s.name, () => {
      it("has a complete contract", () => {
        expect(s.description.length).toBeGreaterThan(30);
        expect(s.examples.length).toBeGreaterThanOrEqual(3);
        expect(s.abilities.length).toBeGreaterThan(1);
        expect(s.context.required.length).toBeGreaterThan(0);
        expect(s.instructions.length).toBeGreaterThan(200);
        expect(s.completion.length).toBeGreaterThan(1);
        expect(s.completion.some((c) => c.required)).toBe(true);
        expect(s.verification.length).toBeGreaterThan(0);
        expect(s.artifacts.length).toBeGreaterThan(0);
        expect(s.limitations.length).toBeGreaterThan(0);
        expect(s.approval.summary).toBeTruthy();
      });

      it("allowlists only capabilities registered for its host, never internal ones", () => {
        const session = sessionFor(s.host);
        const eff = session.specialists.effectiveCapabilities(s, session.registry);
        expect(eff.unavailable).toEqual([]);
        expect(session.specialists.validate(s, session.registry)).toEqual([]);
        for (const c of eff.available) expect(session.registry.get(c)!.internal).toBeFalsy();
      });

      it("read-only specialists have no modifying capability", () => {
        const session = sessionFor(s.host);
        const mods = s.capabilities.filter((c) => session.registry.get(c)?.modifies);
        expect(s.constraints.readOnly ? mods : ["n/a"]).toEqual(s.constraints.readOnly ? [] : ["n/a"]);
        if (!s.constraints.readOnly) expect(mods.length).toBeGreaterThan(0);
      });

      it("required context can be acquired through its own allowlist", () => {
        for (const r of s.context.required) if (r.acquire) expect(s.capabilities).toContain(r.acquire.capability);
      });

      it("routes its own example prompts to itself", () => {
        const router = new SpecialistRouter(new SpecialistRegistry());
        for (const ex of s.examples) {
          const r = router.route(ex, s.host);
          expect(r.kind === "specialist" ? r.specialist.id : `${r.kind}: ${"reason" in r ? r.reason : ""}`, ex).toBe(s.id);
        }
      });
    });
  }

  it("specialists on the same host have distinct instructions and allowlists", () => {
    for (const host of ["word", "excel", "powerpoint", "outlook"] as HostKind[]) {
      const list = reg.forHost(host);
      expect(new Set(list.map((s) => s.instructions)).size).toBe(list.length);
      expect(new Set(list.map((s) => [...s.capabilities].sort().join(","))).size).toBe(list.length);
    }
  });
});

describe("routing", () => {
  const router = new SpecialistRouter(new SpecialistRegistry());

  it.each([
    ["word", "Turn this rough draft into an executive-ready strategy document", "word.writer"],
    ["word", "Proofread the selected text", "word.editor"],
    ["excel", "Find anomalies in this workbook, explain them, and create a summary dashboard", "excel.dashboard"],
    ["excel", "Remove duplicate rows from the Orders sheet", "excel.cleaner"],
    ["powerpoint", "Create a complete 10-slide investor presentation from this document", "powerpoint.author"],
    ["outlook", "Summarize this thread, extract commitments, and draft a reply addressing each one", "outlook.reply"],
    ["outlook", "Who committed to what in this thread?", "outlook.thread"],
  ] as [HostKind, string, string][])("%s: %s → %s", (host, req, id) => {
    const r = router.route(req, host);
    expect(r.kind).toBe("specialist");
    if (r.kind === "specialist") expect(r.specialist.id).toBe(id);
  });

  it("only considers specialists for the active host", () => {
    const r = router.route("Create a summary dashboard with charts", "word");
    expect(r.kind === "specialist" ? r.specialist.host : "none").not.toBe("excel");
  });

  it("rejects an explicit selection for another host", () => {
    const r = router.route("anything", "word", "excel.analyst");
    expect(r).toMatchObject({ kind: "rejected" });
    expect((r as any).reason).toMatch(/works in excel, not in word/);
    expect(router.route("x", "word", "nope.nope").kind).toBe("rejected");
  });

  it("asks a clarifying question when the request is materially ambiguous", () => {
    const r = router.route("Review and rewrite this document", "word");
    expect(r.kind).toBe("clarify");
    if (r.kind === "clarify") {
      expect(r.options.map((o) => o.id).sort()).toEqual(["word.editor", "word.reviewer"]);
      expect(r.question).toMatch(/Which do you want/);
    }
  });

  it("falls back to the general assistant for unmatched requests", () => {
    expect(router.route("hello there", "word").kind).toBe("general");
    expect(router.route("What does EBITDA stand for?", "excel").kind).toBe("general");
    expect(router.route("anything", "word", "general").kind).toBe("general");
  });

  it("honours an explicit selection on the right host", () => {
    const r = router.route("hello", "powerpoint", "powerpoint.reviewer");
    expect(r).toMatchObject({ kind: "specialist", how: "selected" });
  });
});

describe("switching specialists", () => {
  it("preserves the conversation but revalidates capabilities for the new specialist", async () => {
    const doc = new FakeWordDriver(["Alpha beta.", "Gamma delta."]);
    const model = new ScriptedProvider([
      { call: { name: "document.replace_text", arguments: { target: "paragraphs", start: 0, expectedText: "Alpha beta.", newText: "Alpha, beta." } } },
      { text: "Edited paragraph 0." },
      // After switching to the reviewer, the model tries to edit again.
      { call: { name: "document.replace_text", arguments: { target: "paragraphs", start: 1, expectedText: "Gamma delta.", newText: "Gamma, delta." } } },
      { text: "1. [Observation] ¶1 is terse.\n   [Recommendation] Expand it." },
    ]);
    const s = setup(new WordAdapter(doc), model);
    const editor = s.specialists.get("word.editor")!;
    const reviewer = s.specialists.get("word.reviewer")!;

    const r1 = await s.runtime.run("Edit paragraph 0", { specialist: editor });
    expect(r1.actions.find((a) => a.capability === "document.replace_text")!.status).toBe("succeeded");

    const r2 = await s.runtime.run("Now review the document", { specialist: reviewer });
    // Conversation carried over …
    const req = model.requests[2]!;
    expect(req.messages.some((m) => m.role === "user" && m.content === "Edit paragraph 0")).toBe(true);
    // … but the reviewer's tools exclude editing, and the edit is refused.
    expect(req.tools!.map((t) => t.name)).not.toContain("document.replace_text");
    expect(req.messages[0]!.content).toContain("Active specialist: Document Reviewer");
    expect(req.messages[0]!.content).not.toContain("Active specialist: Document Editor");
    const attempt = r2.actions.find((a) => a.capability === "document.replace_text")!;
    expect(attempt.status).toBe("rejected");
    expect(attempt.error).toMatch(/not available to the Document Reviewer/);
    expect(doc.paragraphs[1]!.text).toBe("Gamma delta.");
    expect(r2.specialist?.id).toBe("word.reviewer");
  });

  it("refuses to run a specialist for another host", async () => {
    const s = setup(new WordAdapter(new FakeWordDriver(["a"])), new ScriptedProvider([{ text: "x" }]));
    const excel = new SpecialistRegistry().get("excel.analyst")!;
    const r = await s.runtime.run("analyse", { specialist: excel });
    expect(r.outcome).toBe("failed");
    expect(r.errors[0]).toMatch(/cannot run in word/);
  });

  it("the generic assistant remains available with all capabilities", async () => {
    const model = new ScriptedProvider([{ text: "Hello." }]);
    const s = setup(new WordAdapter(new FakeWordDriver(["a"])), model);
    const r = await s.runtime.run("hi");
    expect(r.outcome).toBe("completed");
    expect(r.specialist).toBeUndefined();
    expect(model.requests[0]!.tools!.map((t) => t.name)).toContain("document.insert_text");
  });
});
