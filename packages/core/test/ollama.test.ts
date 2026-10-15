import { afterEach, describe, expect, it } from "vitest";
import { FakeWordDriver, InferenceError, OllamaProvider, WordAdapter, isLoopbackUrl } from "../src/index.js";
import { startFakeOllama } from "./fake-ollama.js";
import { setup } from "./helpers.js";

let close: (() => Promise<void>) | undefined;
afterEach(async () => {
  await close?.();
  close = undefined;
});

describe("Ollama adapter (integration with an HTTP Ollama stand-in)", () => {
  it("discovers models and their capabilities", async () => {
    const o = await startFakeOllama({ models: ["llama3.1:8b", "qwen2.5:7b"] });
    close = o.close;
    const p = new OllamaProvider({ baseUrl: o.url });
    expect((await p.ping()).ok).toBe(true);
    const models = await p.listModels();
    expect(models.map((m) => m.name)).toEqual(["llama3.1:8b", "qwen2.5:7b"]);
    expect(models[0]!.parameterSize).toBe("8B");
    const d = await p.describeModel("llama3.1:8b");
    expect(d.contextLength).toBe(131072);
    expect(d.capabilities).toContain("tools");
  });

  it("parses native tool calls and streams text", async () => {
    const o = await startFakeOllama({
      replies: [{ content: "Reading now. ", tool_calls: [{ function: { name: "context.read", arguments: { start: 0 } } }] }],
    });
    close = o.close;
    const p = new OllamaProvider({ baseUrl: o.url });
    const deltas: string[] = [];
    const r = await p.chat({ model: "llama3.1:8b", messages: [{ role: "user", content: "hi" }], onDelta: (t) => deltas.push(t) });
    expect(deltas.join("")).toBe("Reading now. ");
    expect(r.message.toolCalls?.[0]).toMatchObject({ name: "context.read", arguments: { start: 0 } });
  });

  it("sends tools, options and tool results in Ollama's format", async () => {
    const o = await startFakeOllama({ replies: [{ content: "done" }] });
    close = o.close;
    const p = new OllamaProvider({ baseUrl: o.url });
    await p.chat({
      model: "llama3.1:8b",
      messages: [
        { role: "assistant", content: "", toolCalls: [{ id: "1", name: "a.b", arguments: { x: 1 } }] },
        { role: "tool", content: "{}", toolName: "a.b" },
      ],
      tools: [{ name: "a.b", description: "d", parameters: { type: "object", properties: {} } }],
      options: { temperature: 0.1, numCtx: 4096 },
    });
    const body = o.requests.find((r) => r.path === "/api/chat")!.body;
    expect(body.tools[0].function.name).toBe("a.b");
    expect(body.options).toEqual({ temperature: 0.1, num_ctx: 4096 });
    expect(body.messages[0].tool_calls[0].function.arguments).toEqual({ x: 1 });
    expect(body.messages[1].tool_name).toBe("a.b");
  });

  it("classifies errors: missing model, unreachable, timeout", async () => {
    const o = await startFakeOllama({ replies: [{ delayMs: 500, content: "slow" }] });
    close = o.close;
    const p = new OllamaProvider({ baseUrl: o.url, timeoutMs: 100 });
    await expect(p.describeModel("missing:1b")).rejects.toMatchObject({ kind: "model_not_found" });
    await expect(p.chat({ model: "llama3.1:8b", messages: [] })).rejects.toMatchObject({ kind: "timeout" });
    const dead = new OllamaProvider({ baseUrl: "http://127.0.0.1:1" });
    await expect(dead.listModels()).rejects.toMatchObject({ kind: "unreachable" });
    expect((await dead.ping()).ok).toBe(false);
  });

  it("refuses non-loopback endpoints unless explicitly allowed", () => {
    expect(() => new OllamaProvider({ baseUrl: "https://api.example.com" })).toThrowError(InferenceError);
    expect(() => new OllamaProvider({ baseUrl: "http://192.168.1.5:11434" })).toThrowError(/not on this machine/);
    expect(() => new OllamaProvider({ baseUrl: "http://192.168.1.5:11434", allowRemote: true })).not.toThrow();
    expect(isLoopbackUrl("http://[::1]:11434")).toBe(true);
    expect(isLoopbackUrl("http://127.0.0.1.evil.test")).toBe(false);
    expect(isLoopbackUrl("http://localhost.evil.test")).toBe(false);
  });
});

describe("privacy: a full task only talks to the local endpoint", () => {
  it("never contacts any non-loopback host and never falls back", async () => {
    const o = await startFakeOllama({
      replies: [{ tool_calls: [{ function: { name: "context.read", arguments: {} } }] }, { error: "out of memory" }],
    });
    close = o.close;
    const urls: string[] = [];
    const spy: typeof fetch = (input, init) => {
      urls.push(String(input));
      return fetch(input, init);
    };
    const provider = new OllamaProvider({ baseUrl: o.url, fetchImpl: spy });
    const s = setup(new WordAdapter(new FakeWordDriver(["Secret plan: launch on Friday."])), provider, { settings: { model: "llama3.1:8b" } });
    const r = await s.runtime.run("summarize");
    expect(urls.length).toBeGreaterThan(0);
    for (const u of urls) expect(isLoopbackUrl(u)).toBe(true);
    // The second model call failed: reported as failure, no retry elsewhere.
    expect(r.outcome).toBe("failed");
    expect(r.errors.join(" ")).toMatch(/out of memory/);
    expect(urls.filter((u) => u.endsWith("/api/chat"))).toHaveLength(2);
  });
});

import { repetitionCut } from "../src/inference/ollama";
describe("repetition guard", () => {
  it("cuts a block repeated four times in a row", () => {
    const unit = "[Observation] The summary is vague about scope and ownership of the work. [Recommendation] Name the owner. ";
    expect(repetitionCut("Intro. " + unit.repeat(4))).toBe(("Intro. " + unit).length);
  });
  it("leaves normal text alone", () => {
    expect(repetitionCut("A normal paragraph about execution. ".repeat(3) + "and then something else entirely, with new words.")).toBeUndefined();
  });
});
