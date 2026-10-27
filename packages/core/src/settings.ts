import { DEFAULT_POLICY, sanitizePolicy, type ApprovalPolicy } from "./permissions.js";
import type { GenerationOptions } from "./types.js";

/** Language of the local model app's HTTP API. */
export type LocalProviderKind = "openai" | "ollama";

/** Default endpoint for each supported local model API. */
export const DEFAULT_ENDPOINTS: Record<LocalProviderKind, string> = {
  openai: "http://127.0.0.1:1337/v1",
  ollama: "http://127.0.0.1:11434",
};

export interface RuntimeLimits {
  /** Max capability calls per task. */
  maxToolCalls: number;
  /** Max model round-trips per task. */
  maxModelCalls: number;
  /** Wall-clock budget per task. */
  maxDurationMs: number;
  /** Total characters sent to the model across all calls in one task. */
  maxContextChars: number;
  /** Max failed attempts of the same capability+arguments. */
  retryLimit: number;
}

export interface AssistantSettings {
  /** How the add-in reaches the local model app. */
  transport: "companion" | "direct";
  /**
   * Local model API. "openai" covers Atomic Chat and any other
   * OpenAI-compatible server; "ollama" uses Ollama's native API.
   */
  provider: LocalProviderKind;
  /** Base URL of the local model app (used directly, or by the companion). */
  endpoint: string;
  /** @deprecated Kept so settings saved by earlier versions still load. */
  ollamaEndpoint?: string;
  companionUrl: string;
  /** Explicit opt-in to non-loopback endpoints. Off by default; never automatic. */
  allowRemoteEndpoint: boolean;
  model: string;
  generation: GenerationOptions;
  requestTimeoutMs: number;
  /** Characters of Office content placed in one model request. */
  contextBudgetChars: number;
  limits: RuntimeLimits;
  policy: ApprovalPolicy;
  /** Persist task history locally. */
  historyEnabled: boolean;
  historyRetentionDays: number;
  /** Store full action arguments (may contain document text). Off by default. */
  historyStoreFullArgs: boolean;
  /** Use native tool calling when the model reports support. */
  preferNativeTools: boolean;
  /**
   * Allow the user to approve a plan as a bounded batch: declared operations
   * inside their scope then run without individual prompts (operations that
   * policy marks "confirm" still ask). Off = every change is approved separately.
   */
  planBatchApproval: boolean;
}

export const DEFAULT_SETTINGS: AssistantSettings = {
  transport: "companion",
  provider: "openai",
  endpoint: DEFAULT_ENDPOINTS.openai,
  companionUrl: "https://localhost:47615",
  allowRemoteEndpoint: false,
  model: "",
  generation: { temperature: 0.2, numCtx: 8192, numPredict: 4096 },
  requestTimeoutMs: 180_000,
  contextBudgetChars: 16_000,
  limits: {
    maxToolCalls: 40,
    maxModelCalls: 50,
    maxDurationMs: 10 * 60_000,
    maxContextChars: 400_000,
    retryLimit: 2,
  },
  policy: { ...DEFAULT_POLICY },
  historyEnabled: true,
  historyRetentionDays: 30,
  historyStoreFullArgs: false,
  preferNativeTools: true,
  planBatchApproval: true,
};

function num(v: unknown, fallback: number, min: number, max: number): number {
  return typeof v === "number" && Number.isFinite(v) ? Math.min(max, Math.max(min, v)) : fallback;
}

/** Merge persisted settings with defaults, clamping values and enforcing policy floors. */
export function normalizeSettings(raw: unknown): AssistantSettings {
  const r = (raw && typeof raw === "object" ? raw : {}) as Partial<AssistantSettings> & Record<string, any>;
  const d = DEFAULT_SETTINGS;
  const lim = (r.limits ?? {}) as Partial<RuntimeLimits>;
  const gen = (r.generation ?? {}) as GenerationOptions;
  // Earlier versions stored only ollamaEndpoint. A custom Ollama URL is
  // preserved; the old default (11434) migrates to the new default, Atomic Chat.
  const legacyEndpoint = typeof r.ollamaEndpoint === "string" && r.ollamaEndpoint ? r.ollamaEndpoint : undefined;
  const legacyCustom = !!legacyEndpoint && legacyEndpoint.replace(/\/+$/, "") !== DEFAULT_ENDPOINTS.ollama;
  const provider: LocalProviderKind = r.provider === "ollama" ? "ollama" : r.provider === "openai" ? "openai" : legacyCustom ? "ollama" : d.provider;
  const endpoint =
    typeof r.endpoint === "string" && r.endpoint
      ? r.endpoint
      : provider === "ollama"
        ? legacyEndpoint ?? DEFAULT_ENDPOINTS.ollama
        : d.endpoint;
  return {
    transport: r.transport === "direct" ? "direct" : "companion",
    provider,
    endpoint,
    companionUrl: typeof r.companionUrl === "string" && r.companionUrl ? r.companionUrl : d.companionUrl,
    allowRemoteEndpoint: r.allowRemoteEndpoint === true,
    model: typeof r.model === "string" ? r.model : "",
    generation: {
      temperature: num(gen.temperature, d.generation.temperature!, 0, 2),
      topP: gen.topP === undefined ? undefined : num(gen.topP, 0.9, 0, 1),
      numCtx: num(gen.numCtx, d.generation.numCtx!, 1024, 1_048_576),
      numPredict: gen.numPredict === undefined ? undefined : num(gen.numPredict, 1024, 16, 65_536),
      seed: typeof gen.seed === "number" ? Math.trunc(gen.seed) : undefined,
    },
    requestTimeoutMs: num(r.requestTimeoutMs, d.requestTimeoutMs, 5_000, 30 * 60_000),
    contextBudgetChars: num(r.contextBudgetChars, d.contextBudgetChars, 1_000, 2_000_000),
    limits: {
      maxToolCalls: num(lim.maxToolCalls, d.limits.maxToolCalls, 1, 100),
      maxModelCalls: num(lim.maxModelCalls, d.limits.maxModelCalls, 1, 200),
      maxDurationMs: num(lim.maxDurationMs, d.limits.maxDurationMs, 10_000, 60 * 60_000),
      maxContextChars: num(lim.maxContextChars, d.limits.maxContextChars, 10_000, 20_000_000),
      retryLimit: num(lim.retryLimit, d.limits.retryLimit, 0, 5),
    },
    policy: sanitizePolicy(r.policy as Record<string, unknown> | undefined),
    historyEnabled: r.historyEnabled !== false,
    historyRetentionDays: num(r.historyRetentionDays, d.historyRetentionDays, 0, 3650),
    historyStoreFullArgs: r.historyStoreFullArgs === true,
    preferNativeTools: r.preferNativeTools !== false,
    planBatchApproval: r.planBatchApproval !== false,
  };
}
