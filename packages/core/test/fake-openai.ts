import http from "node:http";
import type { AddressInfo } from "node:net";

export interface FakeOpenAIReply {
  content?: string;
  reasoning?: string;
  toolCalls?: Array<{ id?: string; name: string; args: string }>;
  error?: string;
  delayMs?: number;
}

export interface FakeOpenAIOptions {
  models?: string[];
  replies?: FakeOpenAIReply[];
}

/** Minimal HTTP stand-in for an OpenAI-compatible server (Atomic Chat/llama.cpp). */
export async function startFakeOpenAI(opts: FakeOpenAIOptions = {}) {
  const requests: { path: string; body: any }[] = [];
  const models = opts.models ?? ["AtomicChat/Qwen3-8B"];
  let i = 0;
  const server = http.createServer(async (req, res) => {
    let raw = "";
    for await (const c of req) raw += c;
    const body = raw ? JSON.parse(raw) : undefined;
    requests.push({ path: req.url!, body });
    const json = (code: number, v: unknown) => {
      res.writeHead(code, { "content-type": "application/json" });
      res.end(JSON.stringify(v));
    };
    if (req.url === "/v1/models")
      return json(200, { object: "list", data: models.map((m, n) => ({ id: m, object: "model", created: 1_700_000_000 + n, owned_by: "llama.cpp-upstream" })) });
    if (req.url === "/v1/chat/completions") {
      if (!models.includes(body.model)) return json(404, { error: { message: `model '${body.model}' not found`, type: "invalid_request_error" } });
      const r = opts.replies?.[i++] ?? { content: "ok", reasoning: "thinking" };
      if (r.delayMs) await new Promise((x) => setTimeout(x, r.delayMs));
      if (r.error) return json(500, { error: { message: r.error } });
      if (!body.stream)
        return json(200, {
          object: "chat.completion",
          model: body.model,
          choices: [
            {
              index: 0,
              finish_reason: r.toolCalls?.length ? "tool_calls" : "stop",
              message: {
                role: "assistant",
                content: r.content ?? "",
                ...(r.reasoning ? { reasoning_content: r.reasoning } : {}),
                ...(r.toolCalls?.length
                  ? { tool_calls: r.toolCalls.map((t, n) => ({ id: t.id ?? `call_${n}`, type: "function", function: { name: t.name, arguments: t.args } })) }
                  : {}),
              },
            },
          ],
          usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 },
        });
      res.writeHead(200, { "content-type": "text/event-stream" });
      const send = (chunk: unknown) => res.write(`data: ${JSON.stringify(chunk)}\n\n`);
      const base = { id: "chatcmpl-test", object: "chat.completion.chunk", model: body.model };
      if (r.reasoning) send({ ...base, choices: [{ index: 0, delta: { role: "assistant", reasoning_content: r.reasoning } }] });
      for (const word of (r.content ?? "").split(/(?<= )/)) if (word) send({ ...base, choices: [{ index: 0, delta: { content: word } }] });
      if (r.toolCalls)
        for (const [n, t] of r.toolCalls.entries()) {
          const half = Math.floor(t.args.length / 2);
          send({ ...base, choices: [{ index: 0, delta: { tool_calls: [{ index: n, id: t.id ?? `call_${n}`, type: "function", function: { name: t.name, arguments: t.args.slice(0, half) } }] } }] });
          send({ ...base, choices: [{ index: 0, delta: { tool_calls: [{ index: n, function: { arguments: t.args.slice(half) } }] } }] });
        }
      send({ ...base, choices: [{ index: 0, delta: {}, finish_reason: r.toolCalls?.length ? "tool_calls" : "stop" }], usage: { prompt_tokens: 10, completion_tokens: 5 } });
      res.end("data: [DONE]\n\n");
      return;
    }
    json(404, { error: { message: "not found" } });
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  const port = (server.address() as AddressInfo).port;
  return {
    url: `http://127.0.0.1:${port}/v1`,
    requests,
    close: () => new Promise<void>((r) => server.close(() => r())),
  };
}
