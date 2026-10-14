import http from "node:http";
import type { AddressInfo } from "node:net";

export interface FakeOllamaOptions {
  models?: string[];
  tools?: boolean;
  contextLength?: number;
  /** Responses for /api/chat in order. */
  replies?: Array<{ content?: string; tool_calls?: any[]; error?: string; delayMs?: number }>;
}

/** Minimal HTTP stand-in for the Ollama API (version, tags, show, chat incl. NDJSON streaming). */
export async function startFakeOllama(opts: FakeOllamaOptions = {}) {
  const requests: { path: string; body: any }[] = [];
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
    const models = opts.models ?? ["llama3.1:8b"];
    if (req.url === "/api/version") return json(200, { version: "0.12.0" });
    if (req.url === "/api/tags")
      return json(200, { models: models.map((m) => ({ name: m, model: m, size: 4_700_000_000, details: { family: "llama", parameter_size: "8B", quantization_level: "Q4_K_M" } })) });
    if (req.url === "/api/show") {
      if (!models.includes(body.model)) return json(404, { error: `model '${body.model}' not found` });
      return json(200, {
        details: { family: "llama" },
        model_info: { "llama.context_length": opts.contextLength ?? 131072 },
        capabilities: opts.tools === false ? ["completion"] : ["completion", "tools"],
      });
    }
    if (req.url === "/api/chat") {
      if (!models.includes(body.model)) return json(404, { error: `model '${body.model}' not found, try pulling it first` });
      const r = opts.replies?.[i++] ?? { content: "ok" };
      if (r.delayMs) await new Promise((x) => setTimeout(x, r.delayMs));
      if (r.error) return json(500, { error: r.error });
      if (!body.stream) return json(200, { model: body.model, message: { role: "assistant", content: r.content ?? "", tool_calls: r.tool_calls }, done: true, prompt_eval_count: 10, eval_count: 5 });
      res.writeHead(200, { "content-type": "application/x-ndjson" });
      const words = (r.content ?? "").split(/(?<= )/);
      for (const w of words) res.write(JSON.stringify({ message: { role: "assistant", content: w }, done: false }) + "\n");
      if (r.tool_calls) res.write(JSON.stringify({ message: { role: "assistant", content: "", tool_calls: r.tool_calls }, done: false }) + "\n");
      res.end(JSON.stringify({ message: { role: "assistant", content: "" }, done: true, done_reason: "stop", eval_count: 5 }) + "\n");
      return;
    }
    json(404, { error: "not found" });
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  const port = (server.address() as AddressInfo).port;
  return {
    url: `http://127.0.0.1:${port}`,
    requests,
    close: () => new Promise<void>((r) => server.close(() => r())),
  };
}
