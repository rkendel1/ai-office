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
import { InferenceError, isLoopbackUrl, repetitionCut } from "./ollama.js";

export interface OpenAIProviderOptions {
  /** e.g. http://127.0.0.1:1337/v1 (Atomic Chat), or the companion's /v1/openai proxy URL. */
  baseUrl: string;
  timeoutMs?: number;
  /** Remote endpoints are refused unless the user explicitly opted in. */
  allowRemote?: boolean;
  /** Extra headers, e.g. the companion's bearer token. */
  headers?: Record<string, string>;
  /** Optional bearer token for OpenAI-compatible servers that require one. */
  apiKey?: string;
  /**
   * Whether the server supports OpenAI-style tool calling. Atomic Chat does;
   * turn this off to force the text tool protocol for a limited server.
   */
  nativeTools?: boolean;
  fetchImpl?: typeof fetch;
}

interface OpenAIToolCall {
  id?: string;
  index?: number;
  type?: string;
  function?: { name?: string; arguments?: unknown };
}

interface OpenAIMessage {
  role?: string;
  content?: string | null;
  reasoning_content?: string;
  tool_calls?: OpenAIToolCall[];
}

/**
 * Adapter for OpenAI-compatible local servers. Atomic Chat is the default; it
 * also works with llama.cpp's server, LM Studio and Ollama's /v1 API.
 */
const SIZE_BOUNDS = new Set(["minItems", "maxItems", "minLength", "maxLength"]);

/**
 * llama.cpp turns tool schemas into a sampling grammar and expands size bounds
 * into explicit rules, so large limits (e.g. maxItems: 2000) make the grammar
 * unparseable (HTTP 400). Limits are still enforced when a call is validated.
 */
export function stripBounds(schema: unknown): any {
  if (Array.isArray(schema)) return schema.map(stripBounds);
  if (!schema || typeof schema !== "object") return schema;
  return Object.fromEntries(Object.entries(schema).filter(([k]) => !SIZE_BOUNDS.has(k)).map(([k, v]) => [k, stripBounds(v)]));
}

export class OpenAIProvider implements InferenceProvider {
  readonly id = "openai";
  private readonly base: string;
  private readonly timeoutMs: number;
  private readonly fetchImpl: typeof fetch;
  private readonly nativeTools: boolean;

  constructor(private readonly opts: OpenAIProviderOptions) {
    this.base = normalizeBaseUrl(opts.baseUrl);
    this.timeoutMs = opts.timeoutMs ?? 120_000;
    this.fetchImpl = opts.fetchImpl ?? ((...a) => fetch(...a));
    this.nativeTools = opts.nativeTools !== false;
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
      return await this.fetchImpl(this.base + path, {
        ...init,
        headers: {
          "content-type": "application/json",
          ...(this.opts.apiKey ? { authorization: `Bearer ${this.opts.apiKey}` } : {}),
          ...this.opts.headers,
          ...(init.headers as Record<string, string>),
        },
        signal: ctl.signal,
      });
    } catch (e) {
      if (timedOut) throw new InferenceError("timeout", `Request to the local model timed out after ${Math.round(timeoutMs / 1000)}s`);
      if (signal?.aborted) throw new InferenceError("aborted", "Request cancelled");
      const page = (globalThis as { location?: { protocol?: string } }).location;
      if (page?.protocol === "https:" && this.base.startsWith("http:"))
        throw new InferenceError(
          "unreachable",
          `This pane is served over HTTPS, and Office (especially on Mac) usually blocks it from calling ${this.base} directly. Switch Connection to "Companion (recommended)", which relays to the local model app for you. (${(e as Error).message})`,
        );
      throw new InferenceError("unreachable", `Cannot reach the local model app at ${this.base}. Is it running? (${(e as Error).message})`);
    } finally {
      clearTimeout(timer);
      signal?.removeEventListener("abort", onAbort);
    }
  }

  private async fail(res: Response, model?: string): Promise<never> {
    let detail = "";
    try {
      const body = (await res.json()) as any;
      detail = typeof body?.error === "string" ? body.error : (body?.error?.message ?? body?.message ?? "");
    } catch {
      /* ignore */
    }
    if (model && (res.status === 404 || /model.*not found|no such model|not installed/i.test(detail)))
      throw new InferenceError("model_not_found", `Model ${model} is not available in the local model app. ${detail}`.trim(), res.status);
    if (/context|too long|exceeds/i.test(detail))
      throw new InferenceError("context_overflow", `The request does not fit the model's context window: ${detail}`, res.status);
    throw new InferenceError("http", `The local model app returned HTTP ${res.status}${detail ? `: ${detail}` : ""}`, res.status);
  }

  async ping(signal?: AbortSignal) {
    try {
      const res = await this.request("/models", { method: "GET", timeoutMs: 5000 }, signal);
      if (!res.ok) return { ok: false, error: `HTTP ${res.status}` };
      // Read (and discard) the body so the connection is released.
      await res.json().catch(() => undefined);
      return { ok: true };
    } catch (e) {
      return { ok: false, error: (e as Error).message };
    }
  }

  async listModels(signal?: AbortSignal): Promise<ModelInfo[]> {
    const res = await this.request("/models", { method: "GET", timeoutMs: 10_000 }, signal);
    if (!res.ok) await this.fail(res);
    const body = (await res.json()) as any;
    const rows: any[] = Array.isArray(body?.data) ? body.data : Array.isArray(body) ? body : [];
    if (!Array.isArray(rows)) throw new InferenceError("bad_response", "Unexpected response from /models");
    const models: ModelInfo[] = [];
    for (const m of rows) {
      const name = String(m?.id ?? m?.name ?? "");
      if (!name) continue;
      const created = typeof m?.created === "number" && m.created > 1_000_000_000 ? new Date(m.created * 1000).toISOString() : undefined;
      models.push({
        name,
        sizeBytes: typeof m?.size === "number" ? m.size : undefined,
        family: m?.owned_by ?? m?.family,
        modifiedAt: created,
      });
    }
    return models;
  }

  async describeModel(name: string, signal?: AbortSignal): Promise<ModelDetails> {
    const models = await this.listModels(signal);
    const found = models.find((m) => m.name === name);
    if (!found) throw new InferenceError("model_not_found", `Model ${name} is not available in the local model app.`, 404);
    return {
      ...found,
      // The OpenAI-compatible /models response carries no context window or
      // capability metadata. Atomic Chat (llama.cpp) supports tool calling.
      capabilities: this.nativeTools ? ["completion", "tools"] : ["completion"],
    };
  }

  async chat(req: ChatRequest): Promise<ChatResponse> {
    const stream = !!req.onDelta;
    const o = req.options ?? {};
    const body = {
      model: req.model,
      stream,
      messages: req.messages.map(toOpenAIMessage),
      ...(req.tools?.length
        ? {
            tools: req.tools.map((t) => ({ type: "function", function: { name: t.name, description: t.description, parameters: stripBounds(t.parameters) } })),
            tool_choice: "auto",
          }
        : {}),
      ...compact({
        temperature: o.temperature,
        top_p: o.topP,
        max_tokens: o.numPredict,
        seed: o.seed,
      }),
    };
    const res = await this.request("/chat/completions", { method: "POST", body: JSON.stringify(body) }, req.signal);
    if (!res.ok) await this.fail(res, req.model);

    if (!stream) {
      const data = (await res.json()) as any;
      const choice = data?.choices?.[0];
      return fromOpenAI(choice?.message ?? {}, choice?.finish_reason, data?.usage);
    }

    if (!res.body) throw new InferenceError("bad_response", "Streaming response had no body");
    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buf = "";
    let content = "";
    let doneReason: string | undefined;
    let promptTokens: number | undefined;
    let completionTokens: number | undefined;
    const calls = new Map<number, { id?: string; name: string; args: string }>();
    const aborted = new Promise<never>((_, reject) => {
      const fail = () => reject(new InferenceError("aborted", "Request cancelled"));
      if (req.signal?.aborted) fail();
      else req.signal?.addEventListener("abort", fail, { once: true });
    });
    aborted.catch(() => {});
    let sinceCheck = 0;
    let repeated = false;
    try {
      for (;;) {
        const { value, done } = await Promise.race([reader.read(), aborted]);
        if (done) break;
        buf += decoder.decode(value, { stream: true });
        let nl: number;
        while ((nl = buf.indexOf("\n")) >= 0) {
          const raw = buf.slice(0, nl).replace(/\r$/, "").trim();
          buf = buf.slice(nl + 1);
          if (!raw || raw.startsWith(":") || raw.startsWith("event:")) continue;
          if (!raw.startsWith("data:")) continue;
          const payload = raw.slice(5).trim();
          if (!payload || payload === "[DONE]") continue;
          let chunk: any;
          try {
            chunk = JSON.parse(payload);
          } catch {
            throw new InferenceError("bad_response", "Malformed streaming chunk from the local model app");
          }
          if (chunk.error) {
            const msg = typeof chunk.error === "string" ? chunk.error : (chunk.error.message ?? "stream error");
            throw new InferenceError(/context/i.test(msg) ? "context_overflow" : "http", msg);
          }
          if (chunk.usage) {
            promptTokens = chunk.usage.prompt_tokens ?? promptTokens;
            completionTokens = chunk.usage.completion_tokens ?? completionTokens;
          }
          const choice = chunk.choices?.[0];
          if (!choice) continue;
          if (choice.finish_reason) doneReason = choice.finish_reason;
          const delta = choice.delta ?? {};
          // Reasoning models stream hidden thinking in reasoning_content; it is
          // not shown to the user and is never part of the answer.
          const piece = typeof delta.content === "string" ? delta.content : "";
          if (piece) {
            content += piece;
            req.onDelta?.(piece);
            sinceCheck += piece.length;
            if (sinceCheck >= 200) {
              sinceCheck = 0;
              const cut = repetitionCut(content);
              if (cut !== undefined) {
                content = content.slice(0, cut);
                repeated = true;
              }
            }
          }
          for (const tc of delta.tool_calls ?? []) {
            const index = typeof tc.index === "number" ? tc.index : 0;
            const cur = calls.get(index) ?? { name: "", args: "" };
            if (tc.id) cur.id = tc.id;
            if (tc.function?.name) cur.name += tc.function.name;
            if (typeof tc.function?.arguments === "string") cur.args += tc.function.arguments;
            else if (tc.function?.arguments && typeof tc.function.arguments === "object") cur.args = JSON.stringify(tc.function.arguments);
            calls.set(index, cur);
          }
        }
        if (repeated) break;
      }
    } catch (e) {
      reader.cancel().catch(() => {});
      if (req.signal?.aborted) throw new InferenceError("aborted", "Request cancelled");
      throw e;
    }
    if (repeated) reader.cancel().catch(() => {});
    const out: ChatResponse = {
      message: {
        role: "assistant",
        content,
        ...(calls.size ? { toolCalls: [...calls.values()].filter((c) => c.name).map(toToolCall) } : {}),
      },
      promptTokens,
      completionTokens,
      doneReason: repeated ? "repetition" : doneReason,
    };
    return out;
  }
}

/** Add /v1 when a bare origin is given, and remove trailing slashes. */
export function normalizeBaseUrl(raw: string): string {
  let u: URL;
  try {
    u = new URL(raw.trim());
  } catch {
    throw new InferenceError("unreachable", `Invalid model endpoint: ${raw}`);
  }
  const path = u.pathname.replace(/\/+$/, "");
  u.pathname = path || "/v1";
  u.search = "";
  u.hash = "";
  return u.toString().replace(/\/+$/, "");
}

function compact(o: Record<string, unknown>) {
  return Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined));
}

function toOpenAIMessage(m: ChatMessage) {
  if (m.role === "tool")
    return {
      role: "tool",
      tool_call_id: m.toolCallId ?? m.toolName ?? "call",
      content: m.content,
    };
  if (m.role === "assistant" && m.toolCalls?.length)
    return {
      role: "assistant",
      content: m.content || null,
      tool_calls: m.toolCalls.map((t) => ({ id: t.id, type: "function", function: { name: t.name, arguments: JSON.stringify(t.arguments) } })),
    };
  return { role: m.role, content: m.content };
}

function toToolCall(c: { id?: string; name: string; args: string }): ToolCall {
  let args: unknown = {};
  if (c.args.trim()) {
    try {
      args = JSON.parse(c.args);
    } catch {
      args = { __unparseable: c.args };
    }
  }
  return {
    id: c.id ?? newId("call"),
    name: c.name,
    arguments: (args && typeof args === "object" && !Array.isArray(args) ? args : {}) as Record<string, Json>,
  };
}

function fromOpenAI(message: OpenAIMessage, finish?: string, usage?: any): ChatResponse {
  const toolCalls: ToolCall[] = [];
  for (const tc of message.tool_calls ?? []) {
    const name = tc.function?.name;
    if (!name) continue;
    let args = tc.function?.arguments;
    if (typeof args === "string") {
      try {
        args = args.trim() ? JSON.parse(args) : {};
      } catch {
        args = { __unparseable: args };
      }
    }
    toolCalls.push({
      id: tc.id ?? newId("call"),
      name,
      arguments: (args && typeof args === "object" && !Array.isArray(args) ? args : {}) as Record<string, Json>,
    });
  }
  return {
    message: { role: "assistant", content: message.content ?? "", ...(toolCalls.length ? { toolCalls } : {}) },
    promptTokens: usage?.prompt_tokens,
    completionTokens: usage?.completion_tokens,
    doneReason: finish,
  };
}
