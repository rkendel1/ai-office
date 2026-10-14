/**
 * The model evaluation CLI path, end to end against a stand-in Ollama server.
 * These tests exercise the harness plumbing (local-only endpoint, metering,
 * resources, environment, output files); they are not a model evaluation.
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { assertLocalEndpoint, evalMatrix, readEvalRuns, runModelEvaluation, selectCases } from "../src/evals.js";
import { startFakeOllama } from "../../core/test/fake-ollama.js";

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), "loai-eval-"));
let close: (() => Promise<void>) | undefined;
afterEach(async () => {
  await close?.();
  close = undefined;
});

const analystReplies = (answer: string) => [
  { content: "", tool_calls: [{ function: { name: "workbook.compute_stats", arguments: { sheet: "Sales", measures: [{ column: "Q2", ops: ["sum", "max"] }] } } }] },
  { content: answer },
];

describe("model evaluation (local only)", () => {
  it("refuses non-local endpoints before sending anything", async () => {
    for (const url of ["http://10.0.0.5:11434", "https://ollama.example.com", "http://192.168.1.20:11434"]) {
      expect(() => assertLocalEndpoint(url)).toThrow(/only against a local Ollama endpoint/);
      let called = false;
      await expect(
        runModelEvaluation({ endpoint: url, models: ["m"], fetchImpl: (async () => ((called = true), new Response("{}"))) as any }),
      ).rejects.toThrow(/local Ollama endpoint/);
      expect(called).toBe(false);
    }
  });

  it("never downloads a missing model", async () => {
    const f = await startFakeOllama({ models: ["llama3.1:8b"] });
    close = f.close;
    await expect(runModelEvaluation({ endpoint: f.url, models: ["qwen2.5:14b"] })).rejects.toThrow(/not installed.*never downloads/);
    expect(f.requests.some((r) => r.path.includes("pull"))).toBe(false);
  });

  it("runs a case through the real runtime and Ollama API, recording timing, resources and environment", async () => {
    const f = await startFakeOllama({ models: ["test-model:7b"], replies: analystReplies("Q2 revenue totals 324 (Sales!C2:C4); North is highest at 135.") });
    close = f.close;
    const out = tmp();
    const lines: string[] = [];
    const { runs, files } = await runModelEvaluation({ endpoint: f.url, models: ["test-model:7b"], select: ["excel.analyst.totals"], outDir: out, gpu: "none", log: (l) => lines.push(l) });
    const run = runs[0]!;
    expect(run.kind).toBe("ollama");
    expect(run.summary).toMatchObject({ runs: 1, passed: 1 });
    const r = run.results[0]!;
    expect(r.route).toEqual({ expected: "excel.analyst", actual: "excel.analyst" });
    expect(r.actions.map((a) => a.capability)).toContain("workbook.compute_stats");
    expect(r.timing.modelCalls).toBe(2);
    expect(r.timing.calls.every((c) => c.ms >= 0)).toBe(true);
    expect(r.resources?.after?.modelBytes).toBe(5_200_000_000);
    expect(run.environment).toMatchObject({ endpoint: f.url, ollamaVersion: "0.12.0", gpu: "none", node: process.version });
    expect(run.environment.cpuCount).toBeGreaterThan(0);
    expect(run.model.details?.capabilities).toContain("tools");
    expect(run.settings).toMatchObject({ temperature: 0, seed: 42 });
    // Deterministic generation options reach Ollama.
    const chat = f.requests.find((x) => x.path === "/api/chat")!.body;
    expect(chat.options).toMatchObject({ temperature: 0, seed: 42 });
    expect(f.requests.every((x) => x.path.startsWith("/api/"))).toBe(true);
    expect(files.map((x) => path.extname(x)).sort()).toEqual([".json", ".md"]);
    expect(fs.readFileSync(files.find((x) => x.endsWith(".md"))!, "utf8")).toMatch(/Real local model through Ollama/);
    expect(lines.join("\n")).toMatch(/PASSED\s+excel\.analyst\.totals/);
    // The saved run feeds the compatibility matrix.
    expect(evalMatrix(out)).toMatch(/\| test-model:7b \|/);
  });

  it("scores a model that fabricates a figure as failing grounding", async () => {
    const f = await startFakeOllama({ models: ["test-model:7b"], replies: analystReplies("Q2 revenue totals 412, led by North.") });
    close = f.close;
    const { runs } = await runModelEvaluation({ endpoint: f.url, models: ["test-model:7b"], select: ["excel.analyst.totals"] });
    const r = runs[0]!.results[0]!;
    expect(r.status).toBe("failed");
    expect(r.metrics.find((m) => m.metric === "grounding")).toMatchObject({ status: "fail" });
  });

  it("reports an unreachable model as an error, not a model failure", async () => {
    const f = await startFakeOllama({ models: ["test-model:7b"], replies: [{ error: "out of memory" }] });
    close = f.close;
    const { runs } = await runModelEvaluation({ endpoint: f.url, models: ["test-model:7b"], select: ["general.capabilities"] });
    expect(runs[0]!.results[0]!.status).toBe("error");
    expect(runs[0]!.summary.errors).toBe(1);
  });

  it("self-test runs without Ollama and never appears in the compatibility matrix", async () => {
    const out = tmp();
    const { runs } = await runModelEvaluation({ endpoint: "http://127.0.0.1:1", models: [], selfTest: true, outDir: out, select: ["negative"] });
    expect(runs[0]!.kind).toBe("scripted");
    expect(runs[0]!.summary.passed).toBe(runs[0]!.summary.runs);
    expect(readEvalRuns(out).runs).toHaveLength(1);
    expect(evalMatrix(out)).toMatch(/No models have been evaluated yet/);
  });

  it("selects cases by id, prefix, category or host", () => {
    expect(selectCases(["negative"]).every((c) => c.category.startsWith("negative"))).toBe(true);
    expect(selectCases(["outlook"]).every((c) => c.host === "outlook")).toBe(true);
    expect(() => selectCases(["nope"])).toThrow(/No evaluation cases match/);
  });
});
