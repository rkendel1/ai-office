/**
 * Real-model evaluation harness: types.
 *
 * An evaluation case runs one request through the real runtime (routing,
 * specialists, approvals, verification) against an in-memory Office host, with
 * a local model as the only variable. Scores come from what the runtime and
 * the host observed — never from the model's description of what it did.
 */
import type { HostAdapter } from "../hosts/adapter.js";
import type { ApprovalRequest, ApprovalResponse } from "../permissions.js";
import type { TaskReport } from "../runtime.js";
import type { SharedSource } from "../sources.js";
import type { ScriptStep } from "../testing/fakes.js";
import type { TaskOutcome } from "../history.js";
import type { HostKind, ModelDetails } from "../types.js";

export type EvalMetric = "routing" | "capability_selection" | "plan" | "scope" | "grounding" | "honesty" | "task";
export const EVAL_METRICS: readonly EvalMetric[] = ["routing", "capability_selection", "plan", "scope", "grounding", "honesty", "task"];

export type EvalCategory =
  | "positive"
  | "negative.fabrication"
  | "negative.unsupported"
  | "negative.scope_expansion"
  | "negative.false_completion";

export interface EvalFixture<D = any> {
  driver: D;
  adapter: HostAdapter;
}

export interface EvalExpectations<D = any> {
  /** Expected router result: a specialist id, "general" or "clarify". */
  route: string;
  capabilities?: {
    /** Each must be called by the model at least once. */
    required?: string[];
    /** At least one capability from each group must be called. */
    anyOf?: string[][];
    /** Must never be called (attempts count, whatever their outcome). */
    forbidden?: string[];
    /** No modifying capability may be attempted. */
    readOnly?: boolean;
  };
  plan?: {
    /** plan.propose must succeed before the first change, with these capabilities declared. */
    declares?: string[];
  };
  scope?: {
    /** Application-state check of what must NOT have changed. Returns a problem or undefined. */
    unchanged?: (f: EvalFixture<D>, r: TaskReport) => Promise<string | undefined>;
  };
  grounding?: {
    /** The final text may state only numbers the runtime observed (default true). */
    numbers?: boolean;
    /** Facts the answer must contain. */
    mentions?: RegExp[];
  };
  honesty?: {
    /** Each pattern must appear in the final text (e.g. disclosing a failure). */
    discloses?: { pattern: RegExp; label: string }[];
    /** Each pattern must not appear in an un-negated sentence (e.g. "sent", "verified"). */
    forbidsClaims?: { pattern: RegExp; label: string }[];
    /** Acceptable runtime outcomes. */
    outcomes?: TaskOutcome[];
  };
  /** Task correctness against application state. Returns a problem or undefined. */
  task?: (f: EvalFixture<D>, r: TaskReport) => Promise<string | undefined>;
  /** Require every required specialist completion criterion to be met (default true for specialist routes). */
  completionCriteria?: boolean;
}

export interface EvalCase<D = any> {
  id: string;
  title: string;
  host: HostKind;
  category: EvalCategory;
  /** What this case is designed to catch. */
  probes: string;
  request: string;
  fixture(): EvalFixture<D>;
  sources?: SharedSource[];
  /** Approval policy for the run. Default: approve everything, batch-approve plans. */
  approve?: (req: ApprovalRequest) => ApprovalResponse;
  expect: EvalExpectations<D>;
  /**
   * Ideal behaviour as a deterministic script. Used to validate the harness:
   * the reference must pass every applicable metric.
   */
  reference: (f: EvalFixture<D>) => ScriptStep[];
  /** Known-bad behaviours. Each must fail the named metric (validates negative scoring). */
  flawed?: { name: string; fails: EvalMetric; script: (f: EvalFixture<D>) => ScriptStep[] }[];
}

export type MetricStatus = "pass" | "fail" | "n/a";

export interface MetricResult {
  metric: EvalMetric;
  status: MetricStatus;
  detail: string;
}

export interface CallTiming {
  ms: number;
  /** Time to the first streamed token, when streaming. */
  firstTokenMs?: number;
  promptTokens?: number;
  completionTokens?: number;
}

export interface CaseRunResult {
  caseId: string;
  category: EvalCategory;
  host: HostKind;
  repetition: number;
  /** "error" = the harness or model could not complete the run (inference error, timeout). */
  status: "passed" | "failed" | "timeout" | "error";
  error?: string;
  route: { expected: string; actual: string };
  metrics: MetricResult[];
  outcome?: TaskOutcome;
  finalText?: string;
  actions: { capability: string; status: string; approval?: string; error?: string }[];
  timing: {
    totalMs: number;
    modelCalls: number;
    calls: CallTiming[];
    promptTokens: number;
    completionTokens: number;
    /** Completion tokens per second of model time, when token counts were reported. */
    tokensPerSecond?: number;
  };
  resources?: { before?: ResourceSample; after?: ResourceSample };
}

/** A point-in-time resource reading. Fields are omitted when unavailable. */
export interface ResourceSample {
  at: string;
  /** From Ollama /api/ps: memory used by the loaded model. */
  modelBytes?: number;
  modelVramBytes?: number;
  hostFreeMemBytes?: number;
  hostTotalMemBytes?: number;
  /** Load average (1 min), when the OS reports it. */
  load1?: number;
}

export interface EvalEnvironment {
  harnessVersion: string;
  recordedAt: string;
  /** Always a loopback URL: remote endpoints are refused. */
  endpoint: string;
  ollamaVersion?: string;
  os?: string;
  arch?: string;
  cpu?: string;
  cpuCount?: number;
  totalMemBytes?: number;
  gpu?: string;
  node?: string;
}

export interface EvalRun {
  schema: 1;
  /** "ollama" = a real local model; "scripted" = the harness self-test (never a model result). */
  kind: "ollama" | "scripted";
  model: { name: string; details?: Partial<ModelDetails>; sizeBytes?: number };
  settings: { temperature?: number; seed?: number; numCtx?: number; timeoutMs: number; repetitions: number };
  environment: EvalEnvironment;
  results: CaseRunResult[];
  summary: EvalSummary;
}

export interface EvalSummary {
  cases: number;
  runs: number;
  passed: number;
  failed: number;
  timeouts: number;
  errors: number;
  metrics: Record<EvalMetric, { pass: number; fail: number; na: number; rate: number | null }>;
  categories: Record<string, { runs: number; passed: number }>;
  latency: { medianTaskMs: number | null; p90TaskMs: number | null; medianCallMs: number | null; medianFirstTokenMs: number | null; tokensPerSecond: number | null };
  peakModelBytes?: number;
  peakModelVramBytes?: number;
}
