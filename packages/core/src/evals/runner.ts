/**
 * Runs evaluation cases through the production runtime: the same router,
 * specialists, capability registry, approval checks and verification the
 * add-in uses. Only the model and the Office host differ (the host is an
 * in-memory fixture so results are reproducible).
 */
import { MemoryHistoryStore } from "../history.js";
import type { ApprovalRequest, ApprovalResponse } from "../permissions.js";
import type { TaskReport } from "../runtime.js";
import { createSession } from "../session.js";
import { DEFAULT_SETTINGS, normalizeSettings, type AssistantSettings } from "../settings.js";
import { ScriptedProvider, type ScriptStep } from "../testing/fakes.js";
import type { ChatRequest, ChatResponse, InferenceProvider } from "../types.js";
import { errorMessage } from "../util.js";
import { EVAL_CASES } from "./cases.js";
import { scoreCase } from "./scoring.js";
import { EVAL_METRICS, type CallTiming, type CaseRunResult, type EvalCase, type EvalEnvironment, type EvalFixture, type EvalRun, type EvalSummary, type ResourceSample } from "./types.js";

export const EVAL_HARNESS_VERSION = "1";
/** Deterministic generation by default so repeated runs are comparable. */
export const EVAL_GENERATION = { temperature: 0, seed: 42 } as const;
export const EVAL_TIMEOUT_MS = 5 * 60_000;

/** Wraps a provider to time each model call (total and first token) and collect token counts. */
export class MeteredProvider implements InferenceProvider {
  readonly id: string;
  readonly calls: CallTiming[] = [];
  constructor(
    private readonly inner: InferenceProvider,
    private readonly now: () => number = () => performance.now(),
  ) {
    this.id = inner.id;
  }
  listModels(signal?: AbortSignal) {
    return this.inner.listModels(signal);
  }
  describeModel(name: string, signal?: AbortSignal) {
    return this.inner.describeModel(name, signal);
  }
  ping(signal?: AbortSignal) {
    return this.inner.ping(signal);
  }
  async chat(req: ChatRequest): Promise<ChatResponse> {
    const start = this.now();
    let first: number | undefined;
    const onDelta = req.onDelta
      ? (t: string) => {
          if (first === undefined && t) first = this.now() - start;
          req.onDelta!(t);
        }
      : undefined;
    const timing: CallTiming = { ms: 0 };
    this.calls.push(timing);
    try {
      const res = await this.inner.chat({ ...req, onDelta });
      timing.promptTokens = res.promptTokens;
      timing.completionTokens = res.completionTokens;
      return res;
    } finally {
      timing.ms = Math.round(this.now() - start);
      if (first !== undefined) timing.firstTokenMs = Math.round(first);
    }
  }
}

export const defaultEvalApprover = (r: ApprovalRequest): ApprovalResponse => (r.capability === "plan.propose" ? { approved: true, batch: true } : { approved: true });

export interface EvalRunOptions {
  /** "ollama" for a real local model. "scripted" runs each case's reference script (harness self-test). */
  kind: "ollama" | "scripted";
  model: string;
  provider?: InferenceProvider;
  /** Override the script per case (scripted kind only), e.g. to replay a flawed behaviour. */
  scriptFor?: (c: EvalCase, f: EvalFixture) => ScriptStep[];
  cases?: EvalCase[];
  repetitions?: number;
  timeoutMs?: number;
  settings?: Partial<AssistantSettings>;
  environment: EvalEnvironment;
  /** Resource probe, sampled before and after each case (e.g. Ollama /api/ps and free memory). */
  resources?: () => Promise<ResourceSample | undefined>;
  onResult?: (r: CaseRunResult, done: number, total: number) => void;
  signal?: AbortSignal;
}

/** Inference failures that mean the run could not be evaluated (as opposed to the model doing badly). */
const INFRA = /cannot reach|not installed|timed out|aborted|cancelled|HTTP 5\d\d|does not fit the model's context/i;

export async function runEvalCase(c: EvalCase, repetition: number, o: EvalRunOptions): Promise<CaseRunResult> {
  const started = performance.now();
  const fixture = c.fixture();
  const base = o.kind === "scripted" || !o.provider ? new ScriptedProvider((o.scriptFor ?? ((x, f) => x.reference(f)))(c, fixture)) : o.provider;
  const meter = new MeteredProvider(base);
  const settings = normalizeSettings({
    ...DEFAULT_SETTINGS,
    model: o.model,
    requestTimeoutMs: o.timeoutMs ?? EVAL_TIMEOUT_MS,
    ...o.settings,
    generation: { ...DEFAULT_SETTINGS.generation, ...EVAL_GENERATION, ...o.settings?.generation },
  });
  const session = createSession({
    adapter: fixture.adapter,
    provider: meter,
    approver: async (req) => (c.approve ?? defaultEvalApprover)(req),
    settings: () => settings,
    history: new MemoryHistoryStore(),
  });
  const route = session.router.route(c.request, c.host);
  const routed = route.kind === "specialist" ? route.specialist.id : route.kind;
  const before = await o.resources?.().catch(() => undefined);

  const finish = async (p: Partial<CaseRunResult> & Pick<CaseRunResult, "status" | "metrics">): Promise<CaseRunResult> => {
    const after = await o.resources?.().catch(() => undefined);
    const promptTokens = meter.calls.reduce((n, x) => n + (x.promptTokens ?? 0), 0);
    const completionTokens = meter.calls.reduce((n, x) => n + (x.completionTokens ?? 0), 0);
    const modelMs = meter.calls.reduce((n, x) => n + x.ms, 0);
    return {
      caseId: c.id,
      category: c.category,
      host: c.host,
      repetition,
      route: { expected: c.expect.route, actual: routed },
      actions: [],
      ...p,
      timing: {
        totalMs: Math.round(performance.now() - started),
        modelCalls: meter.calls.length,
        calls: meter.calls,
        promptTokens,
        completionTokens,
        tokensPerSecond: completionTokens && modelMs ? Math.round((completionTokens / modelMs) * 1000 * 10) / 10 : undefined,
      },
      resources: before || after ? { before, after } : undefined,
    };
  };

  if (route.kind === "clarify" || route.kind === "rejected") {
    const ok = c.expect.route === route.kind;
    return finish({
      status: ok ? "passed" : "failed",
      metrics: EVAL_METRICS.map((metric) =>
        metric === "routing"
          ? { metric, status: ok ? "pass" : "fail", detail: ok ? `router returned ${route.kind}` : `expected ${c.expect.route}, router returned ${route.kind}` }
          : { metric, status: "n/a", detail: `not run: the router returned ${route.kind}` },
      ),
    });
  }

  const ctl = new AbortController();
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    ctl.abort();
  }, o.timeoutMs ?? EVAL_TIMEOUT_MS);
  const onAbort = () => ctl.abort();
  o.signal?.addEventListener("abort", onAbort);
  let report: TaskReport;
  try {
    report = await session.runtime.run(c.request, { specialist: route.kind === "specialist" ? route.specialist : undefined, signal: ctl.signal, sources: c.sources });
  } catch (e) {
    return finish({ status: timedOut ? "timeout" : "error", error: errorMessage(e), metrics: [] });
  } finally {
    clearTimeout(timer);
    o.signal?.removeEventListener("abort", onAbort);
  }

  const summary = {
    outcome: report.outcome,
    finalText: report.text.slice(0, 2000),
    actions: report.actions.map((a) => ({ capability: a.capability, status: a.status, approval: a.approval, error: a.error?.slice(0, 200) })),
  };
  if (timedOut) return finish({ ...summary, status: "timeout", error: `No result within ${Math.round((o.timeoutMs ?? EVAL_TIMEOUT_MS) / 1000)}s`, metrics: [] });
  const infra = report.errors.find((e) => INFRA.test(e));
  if (infra && !report.actions.length && !report.text) return finish({ ...summary, status: "error", error: infra, metrics: [] });

  const metrics = await scoreCase({ c, fixture, routed, report, isModifying: (n) => !!session.registry.get(n)?.modifies });
  return finish({ ...summary, status: metrics.some((m) => m.status === "fail") ? "failed" : "passed", metrics });
}

export async function runEvals(o: EvalRunOptions): Promise<EvalRun> {
  const cases = o.cases ?? EVAL_CASES;
  const reps = Math.max(1, Math.trunc(o.repetitions ?? 1));
  const results: CaseRunResult[] = [];
  const total = cases.length * reps;
  for (let rep = 1; rep <= reps; rep++)
    for (const c of cases) {
      if (o.signal?.aborted) break;
      const r = await runEvalCase(c, rep, o);
      results.push(r);
      o.onResult?.(r, results.length, total);
    }
  let details: EvalRun["model"]["details"];
  if (o.kind === "ollama" && o.provider) details = await o.provider.describeModel(o.model).catch(() => undefined);
  const gen = { ...EVAL_GENERATION, ...o.settings?.generation };
  return {
    schema: 1,
    kind: o.kind,
    model: { name: o.model, details },
    settings: { temperature: gen.temperature, seed: gen.seed, numCtx: o.settings?.generation?.numCtx ?? DEFAULT_SETTINGS.generation.numCtx, timeoutMs: o.timeoutMs ?? EVAL_TIMEOUT_MS, repetitions: reps },
    environment: o.environment,
    results,
    summary: summarize(cases.length, results),
  };
}

function quantile(xs: number[], q: number): number | null {
  if (!xs.length) return null;
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.min(s.length - 1, Math.max(0, Math.ceil(q * s.length) - 1))]!;
}

export function summarize(cases: number, results: CaseRunResult[]): EvalSummary {
  const metrics = Object.fromEntries(
    EVAL_METRICS.map((m) => {
      const xs = results.flatMap((r) => (r.metrics.length ? r.metrics.filter((x) => x.metric === m) : [{ metric: m, status: "fail" as const, detail: r.status }]));
      const pass = xs.filter((x) => x.status === "pass").length;
      const fail = xs.filter((x) => x.status === "fail").length;
      return [m, { pass, fail, na: xs.length - pass - fail, rate: pass + fail ? Math.round((pass / (pass + fail)) * 1000) / 1000 : null }];
    }),
  ) as EvalSummary["metrics"];
  const categories: EvalSummary["categories"] = {};
  for (const r of results) {
    const c = (categories[r.category] ??= { runs: 0, passed: 0 });
    c.runs++;
    if (r.status === "passed") c.passed++;
  }
  const calls = results.flatMap((r) => r.timing.calls);
  const tps = results.map((r) => r.timing.tokensPerSecond).filter((x): x is number => typeof x === "number");
  const peak = (k: "modelBytes" | "modelVramBytes") => {
    const xs = results.flatMap((r) => [r.resources?.before?.[k], r.resources?.after?.[k]]).filter((x): x is number => typeof x === "number");
    return xs.length ? Math.max(...xs) : undefined;
  };
  return {
    cases,
    runs: results.length,
    passed: results.filter((r) => r.status === "passed").length,
    failed: results.filter((r) => r.status === "failed").length,
    timeouts: results.filter((r) => r.status === "timeout").length,
    errors: results.filter((r) => r.status === "error").length,
    metrics,
    categories,
    latency: {
      medianTaskMs: quantile(results.map((r) => r.timing.totalMs), 0.5),
      p90TaskMs: quantile(results.map((r) => r.timing.totalMs), 0.9),
      medianCallMs: quantile(calls.map((c) => c.ms), 0.5),
      medianFirstTokenMs: quantile(calls.map((c) => c.firstTokenMs).filter((x): x is number => typeof x === "number"), 0.5),
      tokensPerSecond: quantile(tps, 0.5),
    },
    peakModelBytes: peak("modelBytes"),
    peakModelVramBytes: peak("modelVramBytes"),
  };
}
