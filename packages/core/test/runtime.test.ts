import { describe, expect, it } from "vitest";
import { FakeWordDriver, ScriptedProvider, WordAdapter, type RuntimeEvent } from "../src/index.js";
import { setup } from "./helpers.js";

const doc = () => new FakeWordDriver(["One", "Two", "Three"]);

describe("agent runtime", () => {
  it("enforces the capability call limit", async () => {
    const steps = Array.from({ length: 10 }, (_, i) => ({ call: { name: "context.read", arguments: { start: 0, count: i + 1 } } }));
    const s = setup(new WordAdapter(doc()), new ScriptedProvider(steps), { settings: { limits: { maxToolCalls: 3, maxModelCalls: 50, maxDurationMs: 60_000, maxContextChars: 1_000_000, retryLimit: 2 } } });
    const r = await s.runtime.run("loop forever");
    expect(r.actions).toHaveLength(3);
    expect(r.limitReached).toBe("capability call limit");
    expect(r.outcome).toBe("partial");
  });

  it("enforces the retry limit for repeatedly failing operations", async () => {
    const bad = { call: { name: "document.replace_text", arguments: { target: "paragraphs", start: 0, expectedText: "nope", newText: "x" } } };
    const s = setup(new WordAdapter(doc()), new ScriptedProvider([bad, bad, bad, bad, bad, { text: "giving up" }]), {
      settings: { limits: { maxToolCalls: 20, maxModelCalls: 50, maxDurationMs: 60_000, maxContextChars: 1_000_000, retryLimit: 1 } },
    });
    const r = await s.runtime.run("x");
    // attempts 1 and 2 execute (fail), later ones are blocked without dispatch/approval
    expect(r.actions.filter((a) => a.status === "failed")).toHaveLength(2);
    expect(r.actions.filter((a) => /Retry limit/.test(a.error ?? ""))).toHaveLength(3);
    expect(s.approver.requests).toHaveLength(2);
  });

  it("cancellation prevents new actions from starting", async () => {
    const ac = new AbortController();
    const d = doc();
    const model = new ScriptedProvider([
      { call: { name: "document.insert_text", arguments: { position: "end", text: "A" } } },
      () => {
        ac.abort();
        return { message: { role: "assistant", content: "", toolCalls: [{ id: "x", name: "document.insert_text", arguments: { position: "end", text: "B" } }] } };
      },
    ]);
    const s = setup(new WordAdapter(d), model);
    const r = await s.runtime.run("add A then B", { signal: ac.signal });
    expect(r.outcome).toBe("cancelled");
    expect(d.paragraphs.map((p) => p.text)).toEqual(["One", "Two", "Three", "A"]);
    expect(r.executionSummary).toMatch(/Changes already applied by Office remain/);
  });

  it("executes only one capability per model turn", async () => {
    const d = doc();
    const model = new ScriptedProvider([
      () => ({
        message: {
          role: "assistant",
          content: "",
          toolCalls: [
            { id: "1", name: "document.insert_text", arguments: { position: "end", text: "A" } },
            { id: "2", name: "document.insert_text", arguments: { position: "end", text: "B" } },
          ],
        },
      }),
      { text: "done" },
    ]);
    const s = setup(new WordAdapter(d), model);
    const r = await s.runtime.run("x");
    expect(r.actions).toHaveLength(1);
    const toolMsg = model.requests[1]!.messages.find((m) => m.role === "tool")!;
    expect(toolMsg.content).toMatch(/not run/);
  });

  it("falls back to the text tool protocol for models without native tool support", async () => {
    const model = new ScriptedProvider(
      [{ text: '```tool\n{"name":"context.read","arguments":{}}\n```' }, { text: "Three paragraphs." }],
      { capabilities: ["completion"] },
    );
    const s = setup(new WordAdapter(doc()), model);
    const r = await s.runtime.run("count");
    expect(model.requests[0]!.tools).toBeUndefined();
    expect(model.requests[0]!.messages[0]!.content).toContain("```tool");
    expect(r.actions[0]!.status).toBe("succeeded");
  });

  it("fails gracefully when the model's context window is too small", async () => {
    const big = new FakeWordDriver(["x".repeat(50_000)]);
    big.selectionRange = { start: 0, end: 0 };
    const model = new ScriptedProvider([{ text: "never" }], { contextLength: 4096 });
    const s = setup(new WordAdapter(big), model, { settings: { contextBudgetChars: 100_000 } });
    const r = await s.runtime.run("summarize");
    // Shrinks content first (tool definitions count against the window too); succeeds within a reduced budget.
    expect(r.outcome).toBe("completed");
    expect(JSON.stringify(model.requests[0]!.messages).length + JSON.stringify(model.requests[0]!.tools).length).toBeLessThan(4096 * 4);

    const tiny = new ScriptedProvider([{ text: "never" }], { contextLength: 1024 });
    const s2 = setup(new WordAdapter(big), tiny, { settings: { contextBudgetChars: 100_000, generation: { numCtx: 1024 } } });
    // System prompt alone nearly fills 1024 tokens → explicit failure, no model call.
    const r2 = await s2.runtime.run("x".repeat(4000));
    expect(r2.outcome).toBe("failed");
    expect(r2.errors[0]).toMatch(/context window|num_ctx/);
    expect(tiny.requests).toHaveLength(0);
  });

  it("reports when no model is selected", async () => {
    const s = setup(new WordAdapter(doc()), new ScriptedProvider([]), { settings: { model: "" } });
    const r = await s.runtime.run("hi");
    expect(r.outcome).toBe("failed");
    expect(r.errors[0]).toMatch(/No model selected/);
  });

  it("detects stale context on the next task", async () => {
    const d = doc();
    const model = new ScriptedProvider([{ call: { name: "context.read", arguments: {} } }, { text: "read" }, { text: "again" }]);
    const s = setup(new WordAdapter(d), model);
    await s.runtime.run("read");
    d.paragraphs[0]!.text = "Changed by the user";
    const events: RuntimeEvent[] = [];
    await s.runtime.run("follow-up", { onEvent: (e) => events.push(e) });
    expect(s.context.get("word:paragraphs:0-2")!.freshness).toBe("stale");
    expect(model.requests[2]!.messages.map((m) => m.content).join("\n")).toContain('stale="true"');
    // Follow-up includes the previous turn.
    expect(model.requests[2]!.messages.some((m) => m.role === "assistant" && m.content === "read")).toBe(true);
  });

  it("records history without document contents", async () => {
    const d = new FakeWordDriver(["Confidential: the merger closes on May 3."]);
    const model = new ScriptedProvider([{ call: { name: "context.read", arguments: {} } }, { text: "Read it." }]);
    const s = setup(new WordAdapter(d), model);
    await s.runtime.run("read");
    const [rec] = await s.history.list();
    expect(rec!.contextRefs.map((c) => c.id)).toContain("word:paragraphs:0-0");
    expect(JSON.stringify(rec)).not.toContain("merger");
  });

  it("emits state events including approval", async () => {
    const events: RuntimeEvent[] = [];
    const s = setup(new WordAdapter(doc()), new ScriptedProvider([{ call: { name: "document.insert_text", arguments: { position: "end", text: "x" } } }, { text: "ok" }]));
    await s.runtime.run("x", { onEvent: (e) => events.push(e) });
    const states = events.filter((e) => e.type === "state").map((e: any) => e.state);
    expect(states).toContain("awaiting_approval");
    expect(states.at(-1)).toBe("completed");
  });
});
