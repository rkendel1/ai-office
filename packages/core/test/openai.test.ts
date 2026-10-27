import { afterEach, describe, expect, it } from "vitest";
import { FakeWordDriver, InferenceError, OpenAIProvider, WordAdapter, isLoopbackUrl, normalizeBaseUrl } from "../src/index.js";
import { startFakeOpenAI } from "./fake-openai.js";
import { setup } from "./helpers.js";

let close: (() => Promise<void>) | undefined;
afterEach(async () => {
  await close?.();
  close = undefined;
});

describe("OpenAI-compatible adapter (Atomic Chat)", () => {
  it("discovers models from /models and treats them as tool-capable", async () => {
    const o = await startFakeOpenAI({ models: ["AtomicChat/Qwen3-8B", "AtomicChat/Llama-3.2-3B"] });
    close = o.close;
    const p = new OpenAIProvider({ baseUrl: o.url });
    expect((await p.ping()).ok).toBe(true);
    expect((await p.listModels()).map((m) => m.name)).toEqual(["AtomicChat/Qwen3-8B", "AtomicChat/Llama-3.2-3B"]);
    const d = await p.describeModel("AtomicChat/Qwen3-8B");
    expect(d.capabilities).toContain("tools");
    expect(d.family).toBe("llama.cpp-upstream");
  });

  it("adds /v1 to a bare origin", () => {
    expect(normalizeBaseUrl("http://127.0.0.1:1337")).toBe("http://127.0.0.1:1337/v1");
    expect(normalizeBaseUrl("http://127.0.0.1:1337/v1/")).toBe("http://127.0.0.1:1337/v1");
  });

  it("streams text, ignores reasoning_content, and assembles tool calls", async () => {
    const o = await startFakeOpenAI({
      replies: [{ content: "Reading now. ", reasoning: "I should call the tool first.", toolCalls: [{ id: "call_abc", name: "context.read", args: '{"start":0}' }] }],
    });
    close = o.close;
    const p = new OpenAIProvider({ baseUrl: o.url });
    const deltas: string[] = [];
    const r = await p.chat({ model: "AtomicChat/Qwen3-8B", messages: [{ role: "user", content: "hi" }], onDelta: (t) => deltas.push(t) });
    expect(deltas.join("")).toBe("Reading now. ");
    expect(r.message.content).toBe("Reading now. ");
    expect(r.message.toolCalls?.[0]).toMatchObject({ id: "call_abc", name: "context.read", arguments: { start: 0 } });
    expect(r.doneReason).toBe("tool_calls");
    expect(r.promptTokens).toBe(10);
  });

  it("sends OpenAI message, tool and generation formats", async () => {
    const o = await startFakeOpenAI({ replies: [{ content: "done" }] });
    close = o.close;
    const p = new OpenAIProvider({ baseUrl: o.url });
    await p.chat({
      model: "AtomicChat/Qwen3-8B",
      messages: [
        { role: "assistant", content: "", toolCalls: [{ id: "call_1", name: "a.b", arguments: { x: 1 } }] },
        { role: "tool", content: "{}", toolName: "a.b", toolCallId: "call_1" },
      ],
      tools: [{ name: "a.b", description: "d", parameters: { type: "object", properties: {} } }],
      options: { temperature: 0.1, numPredict: 512, topP: 0.9 },
    });
    const body = o.requests.find((r) => r.path === "/v1/chat/completions")!.body;
    expect(body.tools[0].function.name).toBe("a.b");
    expect(body.tool_choice).toBe("auto");
    expect(body.max_tokens).toBe(512);
    expect(body.temperature).toBe(0.1);
    expect(body.messages[0].tool_calls[0].function.arguments).toBe('{"x":1}');
    expect(body.messages[1]).toMatchObject({ role: "tool", tool_call_id: "call_1" });
  });

  it("classifies errors: missing model, timeout, unreachable", async () => {
    const o = await startFakeOpenAI({ replies: [{ delayMs: 500, content: "slow" }] });
    close = o.close;
    const p = new OpenAIProvider({ baseUrl: o.url, timeoutMs: 100 });
    await expect(p.describeModel("missing")).rejects.toMatchObject({ kind: "model_not_found" });
    await expect(p.chat({ model: "AtomicChat/Qwen3-8B", messages: [] })).rejects.toMatchObject({ kind: "timeout" });
    const dead = new OpenAIProvider({ baseUrl: "http://127.0.0.1:1" });
    await expect(dead.listModels()).rejects.toMatchObject({ kind: "unreachable" });
    expect((await dead.ping()).ok).toBe(false);
  });

  it("refuses non-loopback endpoints unless explicitly allowed", () => {
    expect(() => new OpenAIProvider({ baseUrl: "https://api.example.com/v1" })).toThrowError(InferenceError);
    expect(() => new OpenAIProvider({ baseUrl: "http://192.168.1.5:1337/v1" })).toThrowError(/not on this machine/);
    expect(() => new OpenAIProvider({ baseUrl: "http://192.168.1.5:1337/v1", allowRemote: true })).not.toThrow();
  });
});

describe("privacy: Atomic Chat tasks only talk to the local endpoint", () => {
  it("never contacts any non-loopback host and never falls back", async () => {
    const o = await startFakeOpenAI({
      models: ["AtomicChat/Qwen3-8B"],
      replies: [{ toolCalls: [{ id: "call_1", name: "context.read", args: "{}" }] }, { content: "The plan launches Friday." }],
    });
    close = o.close;
    const urls: string[] = [];
    const spy: typeof fetch = (input, init) => {
      urls.push(String(input));
      return fetch(input, init);
    };
    const provider = new OpenAIProvider({ baseUrl: o.url, fetchImpl: spy });
    const s = setup(new WordAdapter(new FakeWordDriver(["Secret plan: launch on Friday."])), provider, { settings: { model: "AtomicChat/Qwen3-8B" } });
    const r = await s.runtime.run("summarize");
    expect(urls.length).toBeGreaterThan(0);
    for (const u of urls) expect(isLoopbackUrl(u)).toBe(true);
    expect(r.outcome).toBe("completed");
    expect(urls.filter((u) => u.endsWith("/chat/completions"))).toHaveLength(2);
  });
});
