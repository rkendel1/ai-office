/** Markdown reports and the model compatibility matrix, built only from recorded runs. */
import { EVAL_CASES } from "./cases.js";
import { EVAL_METRICS, type EvalMetric, type EvalRun } from "./types.js";

const METRIC_LABEL: Record<EvalMetric, string> = {
  routing: "Routing",
  capability_selection: "Capability selection",
  plan: "Plan correctness",
  scope: "Scope adherence",
  grounding: "Grounding",
  honesty: "Honest reporting",
  task: "Task state",
};

const pct = (x: number | null | undefined) => (x === null || x === undefined ? "—" : `${Math.round(x * 100)}%`);
const ms = (x: number | null | undefined) => (x === null || x === undefined ? "—" : x >= 10_000 ? `${(x / 1000).toFixed(0)} s` : x >= 1000 ? `${(x / 1000).toFixed(1)} s` : `${x} ms`);
const gb = (x?: number) => (typeof x === "number" ? `${(x / 1024 ** 3).toFixed(1)} GB` : "—");
const cell = (s: string) => s.replace(/\|/g, "\\|").replace(/\n/g, " ");

export function evalMarkdown(run: EvalRun): string {
  const s = run.summary;
  const env = run.environment;
  const d = run.model.details;
  const lines = [
    `# Evaluation: ${run.model.name}`,
    "",
    run.kind === "scripted"
      ? "> **Harness self-test.** These results come from the cases' reference scripts, not from a model. They show the harness scores correctly; they say nothing about any model."
      : `> Real local model through Ollama at ${env.endpoint}. Office hosts are in-memory fixtures; see docs/model-compatibility.md for what this does and does not show.`,
    "",
    `- Recorded: ${env.recordedAt}`,
    `- Model: ${run.model.name}${d ? ` (${[d.family, d.parameterSize, d.quantization, d.contextLength ? `${d.contextLength} ctx` : ""].filter(Boolean).join(", ")})` : ""}${d?.capabilities ? `; capabilities: ${d.capabilities.join(", ")}` : ""}`,
    `- Settings: temperature ${run.settings.temperature ?? "default"}, seed ${run.settings.seed ?? "none"}, num_ctx ${run.settings.numCtx ?? "default"}, timeout ${ms(run.settings.timeoutMs)}, ${run.settings.repetitions} repetition(s)`,
    `- Machine: ${[env.os, env.arch, env.cpu && `${env.cpu}${env.cpuCount ? ` ×${env.cpuCount}` : ""}`, env.totalMemBytes && `${gb(env.totalMemBytes)} RAM`, env.gpu].filter(Boolean).join("; ") || "not recorded"}`,
    `- Ollama ${env.ollamaVersion ?? "version unknown"}; Node ${env.node ?? "—"}; harness v${env.harnessVersion}`,
    "",
    `**${s.passed}/${s.runs} runs passed** (${s.failed} failed, ${s.timeouts} timed out, ${s.errors} errors).`,
    "",
    "| Metric | Pass | Fail | N/A | Rate |",
    "|---|---|---|---|---|",
    ...EVAL_METRICS.map((m) => `| ${METRIC_LABEL[m]} | ${s.metrics[m].pass} | ${s.metrics[m].fail} | ${s.metrics[m].na} | ${pct(s.metrics[m].rate)} |`),
    "",
    "| Category | Passed |",
    "|---|---|",
    ...Object.entries(s.categories).map(([k, v]) => `| ${k} | ${v.passed}/${v.runs} |`),
    "",
    `Latency: median task ${ms(s.latency.medianTaskMs)}, p90 task ${ms(s.latency.p90TaskMs)}, median model call ${ms(s.latency.medianCallMs)}, median first token ${ms(s.latency.medianFirstTokenMs)}, ${s.latency.tokensPerSecond ?? "—"} tokens/s. Peak model memory ${gb(s.peakModelBytes)} (VRAM ${gb(s.peakModelVramBytes)}).`,
    "",
    "## Cases",
    "",
    "| Case | Rep | Result | Route | Failed metrics | Time | Calls |",
    "|---|---|---|---|---|---|---|",
    ...run.results.map((r) => {
      const failed = r.metrics.filter((m) => m.status === "fail");
      return `| ${r.caseId} | ${r.repetition} | ${r.status} | ${r.route.actual}${r.route.actual !== r.route.expected ? ` (expected ${r.route.expected})` : ""} | ${cell(failed.length ? failed.map((m) => `**${m.metric}**: ${m.detail}`).join("; ").slice(0, 400) : r.error ?? "—")} | ${ms(r.timing.totalMs)} | ${r.timing.modelCalls} |`;
    }),
    "",
  ];
  return lines.join("\n");
}

/**
 * Model compatibility matrix. Lists only models with a recorded run of kind
 * "ollama"; it never infers results for models that were not evaluated.
 */
export function compatibilityMatrix(runs: EvalRun[]): string {
  const real = runs.filter((r) => r.kind === "ollama");
  const latest = new Map<string, EvalRun>();
  for (const r of real) {
    const key = `${r.model.name}|${r.environment.cpu ?? ""}|${r.environment.gpu ?? ""}`;
    const prev = latest.get(key);
    if (!prev || prev.environment.recordedAt < r.environment.recordedAt) latest.set(key, r);
  }
  const rows = [...latest.values()].sort((a, b) => a.model.name.localeCompare(b.model.name));
  if (!rows.length)
    return [
      "_No models have been evaluated yet._ No result in this table is estimated or carried over from another model: a model appears here only after the harness has run it on real hardware and the run file has been committed.",
      "",
      `The suite has ${EVAL_CASES.length} cases (${EVAL_CASES.filter((c) => c.category.startsWith("negative")).length} negative).`,
    ].join("\n");
  const header = `| Model | Quant | Machine | Ollama | Passed | ${EVAL_METRICS.map((m) => METRIC_LABEL[m]).join(" | ")} | Median task | Tokens/s | Model memory | Date |`;
  return [
    header,
    `|${"---|".repeat(9 + EVAL_METRICS.length)}`,
    ...rows.map((r) => {
      const s = r.summary;
      const machine = [r.environment.cpu, r.environment.gpu ?? "no GPU reported", r.environment.totalMemBytes && gb(r.environment.totalMemBytes)].filter(Boolean).join(", ");
      return `| ${r.model.name} | ${r.model.details?.quantization ?? "—"} | ${cell(machine)} | ${r.environment.ollamaVersion ?? "—"} | ${s.passed}/${s.runs} | ${EVAL_METRICS.map((m) => pct(s.metrics[m].rate)).join(" | ")} | ${ms(s.latency.medianTaskMs)} | ${s.latency.tokensPerSecond ?? "—"} | ${gb(s.peakModelBytes)} | ${r.environment.recordedAt.slice(0, 10)} |`;
    }),
  ].join("\n");
}

/** Structural check of a run file before it is used in the matrix. */
export function validateEvalRun(v: unknown): string | undefined {
  const r = v as Partial<EvalRun> | null;
  if (!r || typeof r !== "object") return "not an object";
  if (r.schema !== 1) return "unsupported schema";
  if (r.kind !== "ollama" && r.kind !== "scripted") return "kind must be ollama or scripted";
  if (!r.model || typeof r.model.name !== "string") return "missing model name";
  if (!r.environment || typeof r.environment.recordedAt !== "string" || typeof r.environment.endpoint !== "string") return "missing environment";
  if (!Array.isArray(r.results) || !r.summary) return "missing results";
  return undefined;
}
