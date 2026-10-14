import type {
  ChatMessage,
  ChatRequest,
  ChatResponse,
  InferenceProvider,
  Json,
  ModelDetails,
  ModelInfo,
  ToolCall,
} from "../types.js";
import { newId } from "../util.js";

export type InferenceErrorKind =
  | "unreachable"
  | "model_not_found"
  | "timeout"
  | "aborted"
  | "context_overflow"
  | "remote_blocked"
  | "bad_response"
  | "http";

export class InferenceError extends Error {
  constructor(
    public readonly kind: InferenceErrorKind,
    message: string,
    public readonly status?: number,
  ) {
    super(message);
    this.name = "InferenceError";
  }
}

export function isLoopbackUrl(url: string): boolean {
  let u: URL;
  try {
    u = new URL(url);
  } catch {
    return false;
  }
  const h = u.hostname.replace(/^\[|\]$/g, "").toLowerCase();
  return h === "localhost" || h === "::1" || /^127\.\d{1,3}\.\d{1,3}\.\d{1,3}$/.test(h);
}

export interface OllamaProviderOptions {
  /** e.g. http://127.0.0.1:11434, or the companion's Ollama proxy URL. */
  baseUrl: string;
  timeoutMs?: number;
  /**
   * Remote endpoints are refused unless the user explicitly opted in. There
   * is never an automatic fallback to any other endpoint.
   */
  allowRemote?: boolean;
  /** Extra headers, e.g. the companion's bearer token. */
  headers?: Record<string, string>;
  fetchImpl?: typeof fetch;
}

interface OllamaToolCall {
  function?: { name?: string; arguments?: unknown };
}

interface OllamaMessage {
  role?: string;
  content?: string;
  thinking?: string;
  tool_calls?: OllamaToolCall[];
}

/** Ollama HTTP API adapter (https://github.com/ollama/ollama/blob/main/docs/api.md). */
export class OllamaProvider implements InferenceProvider {
  readonly id = "ollama";
  private readonly base: string;
  private readonly timeoutMs: number;
  private readonly fetchImpl: typeof fetch;

  constructor(private readonly opts: OllamaProviderOptions) {
    this.base = opts.baseUrl.replace(/\/+$/, "");
    this.timeoutMs = opts.timeoutMs ?? 120_000;
    this.fetchImpl = opts.fetchImpl ?? ((...a) => fetch(...a));
    if (!opts.allowRemote && !isLoopbackUrl(this.base)) {
      throw new InferenceError(
        "remote_blocked",
        `Endpoint ${this.base} is not on this machine. Remote endpoints are disabled; enable them explicitly in Settings if you understand that document content will leave this computer.`,
      );
    }
  }

  private async request(path: string, init: RequestInit & { timeoutMs?: number }, signal?: AbortSignal): Promise<Response> {
    const ctl = new AbortController();
    const timeoutMs = init.timeoutMs ?? this.timeoutMs;
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      ctl.abort();
    }, timeoutMs);
    const onAbort = () => ctl.abort();
    signal?.addEventListener("abort", onAbort);
    try {
      const res = await this.fetchImpl(this.base + path, {
        ...init,
        headers: { "content-type": "application/json", ...this.opts.headers, ...(init.headers as Record<string, string>) },
        signal: ctl.signal,
      });
      return res;
    } catch (e) {
      if (timedOut) throw new InferenceError("timeout", `Request to the local model timed out after ${Math.round(timeoutMs / 1000)}s`);
      if (signal?.aborted) throw new InferenceError("aborted", "Request cancelled");
      throw new InferenceError("unreachable", `Cannot reach Ollama at ${this.base}. Is it running? (${(e as Error).message})`);
    } finally {
      clearTimeout(timer);
      signal?.removeEventListener("abort", onAbort);
    }
  }

  private async fail(res: Response, model?: string): Promise<never> {
    let detail = "";
    try {
      const body = (await res.json()) as { error?: string };
      detail = body.error ?? "";
    } catch {
      /* ignore */
    }
    if (res.status === 404 || /not found/i.test(detail))
      throw new InferenceError("model_not_found", `Model ${model ?? ""} is not installed in Ollama. ${detail}`.trim(), res.status);
    if (/context|too long|exceeds/i.test(detail))
      throw new InferenceError("context_overflow", `The request does not fit the model's context window: ${detail}`, res.status);
    throw new InferenceError("http", `Ollama returned HTTP ${res.status}${detail ? `: ${detail}` : ""}`, res.status);
  }

  async ping(signal?: AbortSignal) {
    try {
      const res = await this.request("/api/version", { method: "GET", timeoutMs: 5000 }, signal);
      if (!res.ok) return { ok: false, error: `HTTP ${res.status}` };
      const body = (await res.json()) as { version?: string };
      return { ok: true, version: body.version };
    } catch (e) {
      return { ok: false, error: (e as Error).message };
    }
  }

  async listModels(signal?: AbortSignal): Promise<ModelInfo[]> {
    const res = await this.request("/api/tags", { method: "GET", timeoutMs: 10_000 }, signal);
    if (!res.ok) await this.fail(res);
    const body = (await res.json()) as { models?: any[] };
    if (!Array.isArray(body.models)) throw new InferenceError("bad_response", "Unexpected response from /api/tags");
    return body.models.map((m) => ({
      name: String(m.name ?? m.model),
      sizeBytes: typeof m.size === "number" ? m.size : undefined,
      family: m.details?.family,
      parameterSize: m.details?.parameter_size,
      quantization: m.details?.quantization_level,
      modifiedAt: m.modified_at,
    }));
  }

  async describeModel(name: string, signal?: AbortSignal): Promise<ModelDetails> {
    const res = await this.request("/api/show", { method: "POST", body: JSON.stringify({ model: name }), timeoutMs: 15_000 }, signal);
    if (!res.ok) await this.fail(res, name);
    const body = (await res.json()) as any;
    let contextLength: number | undefined;
    const info = body.model_info ?? {};
    for (const [k, v] of Object.entries(info)) if (k.endsWith(".context_length") && typeof v === "number") contextLength = v;
    const caps: string[] = Array.isArray(body.capabilities)
      ? body.capabilities
      : // Older Ollama: infer tool support from the template.
        typeof body.template === "string" && body.template.includes(".Tools")
        ? ["completion", "tools"]
        : ["completion"];
    return {
      name,
      family: body.details?.family,
      parameterSize: body.details?.parameter_size,
      quantization: body.details?.quantization_level,
      contextLength,
      capabilities: caps,
    };
  }

  async chat(req: ChatRequest): Promise<ChatResponse> {
    const stream = !!req.onDelta;
    const body = {
      model: req.model,
      stream,
      messages: req.messages.map(toOllamaMessage),
      ...(req.tools?.length
        ? { tools: req.tools.map((t) => ({ type: "function", function: { name: t.name, description: t.description, parameters: t.parameters } })) }
        : {}),
      options: compactOptions({
        temperature: req.options?.temperature,
        top_p: req.options?.topP,
        num_ctx: req.options?.numCtx,
        num_predict: req.options?.numPredict,
        seed: req.options?.seed,
      }),
    };
    const res = await this.request("/api/chat", { method: "POST", body: JSON.stringify(body) }, req.signal);
    if (!res.ok) await this.fail(res, req.model);

    if (!stream) {
      const data = (await res.json()) as any;
      return fromOllama(data.message ?? {}, data);
    }

    // NDJSON stream
    if (!res.body) throw new InferenceError("bad_response", "Streaming response had no body");
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buf = "";
    let content = "";
    const toolCalls: OllamaToolCall[] = [];
    let last: any = {};
    try {
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        buf += decoder.decode(value, { stream: true });
        let nl: number;
        while ((nl = buf.indexOf("\n")) >= 0) {
          const line = buf.slice(0, nl).trim();
          buf = buf.slice(nl + 1);
          if (!line) continue;
          let chunk: any;
          try {
            chunk = JSON.parse(line);
          } catch {
            throw new InferenceError("bad_response", "Malformed streaming chunk from Ollama");
          }
          if (chunk.error) throw new InferenceError(/context/i.test(chunk.error) ? "context_overflow" : "http", chunk.error);
          const piece: string = chunk.message?.content ?? "";
          if (piece) {
            content += piece;
            req.onDelta?.(piece);
          }
          if (Array.isArray(chunk.message?.tool_calls)) toolCalls.push(...chunk.message.tool_calls);
          last = chunk;
        }
      }
    } catch (e) {
      if (req.signal?.aborted) throw new InferenceError("aborted", "Request cancelled");
      throw e;
    }
    return fromOllama({ role: "assistant", content, tool_calls: toolCalls }, last);
  }
}

function compactOptions(o: Record<string, unknown>) {
  return Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined));
}

function toOllamaMessage(m: ChatMessage) {
  const out: Record<string, unknown> = { role: m.role, content: m.content };
  if (m.toolCalls?.length)
    out.tool_calls = m.toolCalls.map((t) => ({ function: { name: t.name, arguments: t.arguments } }));
  if (m.role === "tool" && m.toolName) out.tool_name = m.toolName;
  return out;
}

function fromOllama(msg: OllamaMessage, meta: any): ChatResponse {
  const toolCalls: ToolCall[] = [];
  for (const tc of msg.tool_calls ?? []) {
    const name = tc.function?.name;
    if (!name) continue;
    let args = tc.function?.arguments;
    if (typeof args === "string") {
      try {
        args = JSON.parse(args);
      } catch {
        args = { __unparseable: args };
      }
    }
    toolCalls.push({
      id: newId("call"),
      name,
      arguments: (args && typeof args === "object" && !Array.isArray(args) ? args : {}) as Record<string, Json>,
    });
  }
  return {
    message: { role: "assistant", content: msg.content ?? "", ...(toolCalls.length ? { toolCalls } : {}) },
    promptTokens: meta?.prompt_eval_count,
    completionTokens: meta?.eval_count,
    doneReason: meta?.done_reason,
  };
}
