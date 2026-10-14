/**
 * Real-model evaluation from the command line. Inference goes only to a local
 * Ollama endpoint: non-loopback URLs are refused regardless of configuration,
 * and there is no hosted-model fallback. Models are never pulled automatically.
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import {
  EVAL_CASES,
  EVAL_HARNESS_VERSION,
  OllamaProvider,
  compatibilityMatrix,
  evalMarkdown,
  isLoopbackUrl,
  runEvals,
  validateEvalRun,
  type CaseRunResult,
  type EvalCase,
  type EvalEnvironment,
  type EvalRun,
  type ResourceSample,
} from "@loai/core";

export interface ModelEvalOptions {
  endpoint: string;
  models: string[];
  /** Case ids, id prefixes or categories (e.g. "negative", "neg.scope", "excel."). Default: all. */
  select?: string[];
  repetitions?: number;
  timeoutMs?: number;
  outDir?: string;
  /** Free-text GPU description recorded with the run (the harness does not probe hardware). */
  gpu?: string;
  /** Run the reference scripts instead of a model (harness self-test; no Ollama needed). */
  selfTest?: boolean;
  fetchImpl?: typeof fetch;
  log?: (line: string) => void;
  now?: () => Date;
}

export function selectCases(select?: string[]): EvalCase[] {
  if (!select?.length) return EVAL_CASES;
  const out = EVAL_CASES.filter((c) => select.some((s) => c.id === s || c.id.startsWith(s) || c.category === s || c.category.startsWith(s) || c.host === s));
  if (!out.length) throw new Error(`No evaluation cases match ${select.join(", ")}. Cases: ${EVAL_CASES.map((c) => c.id).join(", ")}`);
  return out;
}

export function assertLocalEndpoint(url: string): void {
  if (!isLoopbackUrl(url)) throw new Error(`Refusing to evaluate against ${url}: evaluation runs only against a local Ollama endpoint (localhost / 127.0.0.1 / ::1).`);
}

/** Memory of the loaded model from Ollama's /api/ps, plus host memory and load. */
export function resourceProbe(endpoint: string, model: string, fetchImpl: typeof fetch = fetch, now = () => new Date()): () => Promise<ResourceSample> {
  return async () => {
    const s: ResourceSample = { at: now().toISOString(), hostFreeMemBytes: os.freemem(), hostTotalMemBytes: os.totalmem() };
    const load = os.loadavg()[0];
    if (load) s.load1 = Math.round(load * 100) / 100;
    try {
      const res = await fetchImpl(`${endpoint.replace(/\/+$/, "")}/api/ps`, { signal: AbortSignal.timeout(3000) });
      if (res.ok) {
        const body = (await res.json()) as { models?: { name?: string; model?: string; size?: number; size_vram?: number }[] };
        const m = body.models?.find((x) => x.name === model || x.model === model);
        if (m) {
          if (typeof m.size === "number") s.modelBytes = m.size;
          if (typeof m.size_vram === "number") s.modelVramBytes = m.size_vram;
        }
      }
    } catch {
      /* resource data is best-effort */
    }
    return s;
  };
}

export function captureEnvironment(endpoint: string, ollamaVersion: string | undefined, gpu: string | undefined, now: Date): EvalEnvironment {
  const cpus = os.cpus();
  return {
    harnessVersion: EVAL_HARNESS_VERSION,
    recordedAt: now.toISOString(),
    endpoint,
    ollamaVersion,
    os: `${os.type()} ${os.release()}`,
    arch: os.arch(),
    cpu: cpus[0]?.model?.trim(),
    cpuCount: cpus.length,
    totalMemBytes: os.totalmem(),
    gpu,
    node: process.version,
  };
}

const safe = (s: string) => s.replace(/[^A-Za-z0-9._-]+/g, "_");

function progressLine(r: CaseRunResult, done: number, total: number): string {
  const failed = r.metrics.filter((m) => m.status === "fail").map((m) => m.metric);
  return `[${done}/${total}] ${r.status.toUpperCase().padEnd(7)} ${r.caseId}${r.repetition > 1 ? ` #${r.repetition}` : ""} (${(r.timing.totalMs / 1000).toFixed(1)}s, ${r.timing.modelCalls} calls)${failed.length ? ` — ${failed.join(", ")}` : r.error ? ` — ${r.error}` : ""}`;
}

export async function runModelEvaluation(o: ModelEvalOptions): Promise<{ runs: EvalRun[]; files: string[] }> {
  const log = o.log ?? (() => {});
  const now = o.now ?? (() => new Date());
  const cases = selectCases(o.select);
  const files: string[] = [];
  const runs: EvalRun[] = [];
  const write = (run: EvalRun) => {
    if (!o.outDir) return;
    fs.mkdirSync(o.outDir, { recursive: true });
    const stamp = run.environment.recordedAt.replace(/[:.]/g, "-");
    const base = path.join(o.outDir, `${safe(run.kind === "scripted" ? "self-test" : run.model.name)}_${stamp}`);
    fs.writeFileSync(`${base}.json`, JSON.stringify(run, null, 2) + "\n");
    fs.writeFileSync(`${base}.md`, evalMarkdown(run));
    files.push(`${base}.json`, `${base}.md`);
  };

  if (o.selfTest) {
    const run = await runEvals({ kind: "scripted", model: "reference-scripts", cases, repetitions: 1, environment: captureEnvironment("(none: self-test)", undefined, o.gpu, now()), onResult: (r, d, t) => log(progressLine(r, d, t)) });
    write(run);
    return { runs: [run], files };
  }

  assertLocalEndpoint(o.endpoint);
  if (!o.models.length) throw new Error("Choose at least one model with --model <name>.");
  // allowRemote is never set: the provider refuses non-loopback endpoints too.
  const provider = new OllamaProvider({ baseUrl: o.endpoint, timeoutMs: o.timeoutMs, fetchImpl: o.fetchImpl });
  const ping = await provider.ping();
  if (!ping.ok) throw new Error(`Cannot reach Ollama at ${o.endpoint} (${ping.error}). Start Ollama and try again.`);
  const installed = (await provider.listModels()).map((m) => m.name);

  for (const model of o.models) {
    if (!installed.includes(model) && !installed.includes(`${model}:latest`)) {
      throw new Error(`Model ${model} is not installed in this Ollama. Install it yourself (ollama pull ${model}); the harness never downloads models. Installed: ${installed.join(", ") || "none"}`);
    }
    log(`Evaluating ${model} on ${cases.length} case(s) × ${o.repetitions ?? 1} …`);
    const run = await runEvals({
      kind: "ollama",
      model,
      provider,
      cases,
      repetitions: o.repetitions,
      timeoutMs: o.timeoutMs,
      environment: captureEnvironment(o.endpoint, ping.version, o.gpu, now()),
      resources: resourceProbe(o.endpoint, model, o.fetchImpl, now),
      onResult: (r, d, t) => log(progressLine(r, d, t)),
    });
    const sizes = await provider.listModels().catch(() => []);
    run.model.sizeBytes = sizes.find((m) => m.name === model)?.sizeBytes;
    write(run);
    runs.push(run);
    log(`${model}: ${run.summary.passed}/${run.summary.runs} passed`);
  }
  return { runs, files };
}

/** Read every valid run file in a directory (invalid files are reported, not used). */
export function readEvalRuns(dir: string): { runs: EvalRun[]; invalid: string[] } {
  const runs: EvalRun[] = [];
  const invalid: string[] = [];
  if (!fs.existsSync(dir)) return { runs, invalid };
  for (const f of fs.readdirSync(dir).filter((x) => x.endsWith(".json")).sort()) {
    try {
      const v = JSON.parse(fs.readFileSync(path.join(dir, f), "utf8"));
      const problem = validateEvalRun(v);
      if (problem) invalid.push(`${f}: ${problem}`);
      else runs.push(v);
    } catch (e) {
      invalid.push(`${f}: ${(e as Error).message}`);
    }
  }
  return { runs, invalid };
}

export function evalMatrix(dir: string): string {
  return compatibilityMatrix(readEvalRuns(dir).runs);
}
