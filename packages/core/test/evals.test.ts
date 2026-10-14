/**
 * The evaluation harness, validated against deterministic scripts: each
 * case's reference behaviour must pass every applicable metric, and each
 * known-bad behaviour must fail the metric it targets. This tests the
 * harness, not any model.
 */
import { describe, expect, it } from "vitest";
import { EVAL_CASES, compatibilityMatrix, evalMarkdown, runEvalCase, runEvals, summarize, unqualifiedClaims, validateEvalRun, type EvalRun } from "../src/index.js";

const env = { harnessVersion: "1", recordedAt: "2026-10-13T00:00:00.000Z", endpoint: "http://127.0.0.1:11434" };
const opts = { kind: "scripted" as const, model: "reference", environment: env };

describe("evaluation cases", () => {
  it("cover every metric and negative category the harness reports", () => {
    const cats = new Set(EVAL_CASES.map((c) => c.category));
    for (const k of ["negative.fabrication", "negative.unsupported", "negative.scope_expansion", "negative.false_completion"]) expect(cats.has(k as any)).toBe(true);
    expect(new Set(EVAL_CASES.map((c) => c.host))).toEqual(new Set(["word", "excel", "powerpoint", "outlook"]));
    expect(new Set(EVAL_CASES.map((c) => c.id)).size).toBe(EVAL_CASES.length);
    for (const c of EVAL_CASES.filter((x) => x.category.startsWith("negative"))) expect(c.flawed?.length, c.id).toBeGreaterThan(0);
  });

  for (const c of EVAL_CASES) {
    it(`${c.id}: the reference behaviour passes`, async () => {
      const r = await runEvalCase(c, 1, opts);
      expect(r.error).toBeUndefined();
      expect(r.metrics.filter((m) => m.status === "fail").map((m) => `${m.metric}: ${m.detail}`)).toEqual([]);
      expect(r.status).toBe("passed");
      expect(r.route).toEqual({ expected: c.expect.route, actual: c.expect.route });
    });
    for (const f of c.flawed ?? [])
      it(`${c.id}: catches “${f.name}” (${f.fails})`, async () => {
        const r = await runEvalCase(c, 1, { ...opts, scriptFor: (_, fx) => f.script(fx) });
        expect(r.status).toBe("failed");
        const m = r.metrics.find((x) => x.metric === f.fails)!;
        expect(m.status, m.detail).toBe("fail");
      });
  }
});

describe("scoring details", () => {
  it("treats negated statements as disclosures, not claims", () => {
    expect(unqualifiedClaims("The reply has not been sent. Review it first.", /\bsent\b/)).toEqual([]);
    expect(unqualifiedClaims("Done. I sent the reply.", /\bsent\b/)).toEqual(["I sent the reply."]);
    expect(unqualifiedClaims("I couldn't verify it.\nIt was verified.", /\bverified\b/)).toEqual(["It was verified."]);
  });

  it("a timeout is reported as a timeout, not as a model failure", async () => {
    const slow = { ...EVAL_CASES[0]! };
    const r = await runEvalCase(slow, 1, {
      ...opts,
      kind: "ollama",
      timeoutMs: 50,
      provider: {
        id: "slow",
        ping: async () => ({ ok: true }),
        listModels: async () => [],
        describeModel: async (name) => ({ name, capabilities: ["completion", "tools"], contextLength: 32768 }),
        chat: (req) => new Promise((_, rej) => req.signal?.addEventListener("abort", () => rej(Object.assign(new Error("aborted"), { name: "AbortError" })))),
      },
    });
    expect(r.status).toBe("timeout");
    expect(r.timing.modelCalls).toBe(1);
  });
});

describe("reports", () => {
  it("summarises a self-test run, labelled as not a model result", async () => {
    const run = await runEvals({ ...opts, cases: EVAL_CASES.slice(0, 3), repetitions: 2 });
    expect(run.summary).toMatchObject({ cases: 3, runs: 6, passed: 6 });
    expect(run.summary.metrics.routing.rate).toBe(1);
    expect(run.summary.metrics.plan.na).toBe(6);
    const md = evalMarkdown(run);
    expect(md).toMatch(/Harness self-test/);
    expect(md).toMatch(/\*\*6\/6 runs passed\*\*/);
    expect(validateEvalRun(JSON.parse(JSON.stringify(run)))).toBeUndefined();
  });

  it("the compatibility matrix lists only models that were actually run", async () => {
    const self = await runEvals({ ...opts, cases: EVAL_CASES.slice(0, 1) });
    expect(compatibilityMatrix([self])).toMatch(/No models have been evaluated yet/);
    const real: EvalRun = { ...self, kind: "ollama", model: { name: "qwen2.5:7b", details: { quantization: "Q4_K_M" } }, environment: { ...env, cpu: "Test CPU", ollamaVersion: "0.12.3" } };
    const md = compatibilityMatrix([self, real]);
    expect(md).toMatch(/\| qwen2\.5:7b \| Q4_K_M \| Test CPU/);
    expect(md).not.toMatch(/reference/);
  });

  it("summarize counts runs without metrics (errors, timeouts) as failures of every metric", () => {
    const s = summarize(1, [{ caseId: "x", category: "positive", host: "word", repetition: 1, status: "error", route: { expected: "general", actual: "general" }, metrics: [], actions: [], timing: { totalMs: 5, modelCalls: 0, calls: [], promptTokens: 0, completionTokens: 0 } }]);
    expect(s.metrics.grounding).toMatchObject({ pass: 0, fail: 1, rate: 0 });
    expect(s.errors).toBe(1);
  });
});
