import type { HostAdapter } from "../hosts/adapter.js";
import { MemoryHistoryStore } from "../history.js";
import type { TaskReport } from "../runtime.js";
import { createSession } from "../session.js";
import { DEFAULT_SETTINGS, normalizeSettings, type AssistantSettings } from "../settings.js";
import { ScriptedProvider } from "../testing/fakes.js";
import type { InferenceProvider } from "../types.js";
import { errorMessage } from "../util.js";
import type { AcceptanceEnvironment, AcceptanceMode, AcceptanceRecord, AcceptanceScenario } from "./types.js";

export interface RunOptions<D> {
  driver: D;
  adapter: HostAdapter;
  mode: AcceptanceMode;
  environment: AcceptanceEnvironment;
  /** Use a real local model instead of the scenario's deterministic script. */
  model?: { provider: InferenceProvider; name: string; settings?: Partial<AssistantSettings> };
}

/**
 * Run one scenario end to end: preconditions → fixture setup → task (scripted
 * or real model) → automated checks against application state → record.
 * Manual and visual steps are listed as pending; this runner never marks them passed.
 */
export async function runAcceptance<D>(sc: AcceptanceScenario<D>, o: RunOptions<D>): Promise<AcceptanceRecord> {
  const started = Date.now();
  const ctx = { driver: o.driver, adapter: o.adapter };
  const base = {
    schema: 1 as const,
    scenarioId: sc.id,
    title: sc.title,
    host: sc.host,
    specialist: sc.specialist,
    mode: o.mode,
    model: o.model ? { kind: "ollama" as const, name: o.model.name } : { kind: "scripted" as const },
    environment: o.environment,
    manual: [...sc.manual, ...(sc.visual ?? [])].map((step) => ({ step, status: "pending" as const })),
  };
  const done = (r: Omit<AcceptanceRecord, keyof typeof base | "durationMs" | "checks"> & { checks?: AcceptanceRecord["checks"] }): AcceptanceRecord => ({ ...base, checks: [], ...r, durationMs: Date.now() - started });

  if (o.adapter.host !== sc.host) return done({ outcome: "blocked", reason: `Scenario is for ${sc.host}, not ${o.adapter.host}.`, visual: "not-applicable" });
  try {
    const blocked = await sc.checkPreconditions(ctx);
    if (blocked) return done({ outcome: "blocked", reason: blocked, visual: "not-applicable" });
    await sc.setup(ctx);

    const provider = o.model?.provider ?? new ScriptedProvider(await sc.script(ctx));
    const settings = normalizeSettings({ ...DEFAULT_SETTINGS, model: o.model?.name ?? "scripted", ...o.model?.settings });
    const session = createSession({ adapter: o.adapter, provider, approver: async (req) => sc.approve(req), settings: () => settings, history: new MemoryHistoryStore() });
    const specialist = sc.specialist ? session.specialists.get(sc.specialist) : undefined;
    const report: TaskReport = await session.runtime.run(sc.request, { specialist });
    let checks: AcceptanceRecord["checks"];
    try {
      checks = await sc.checks(ctx, report);
    } catch (e) {
      // The checks read the state the task was meant to create; a missing sheet/range means the task did not produce it.
      checks = [{ id: "checks-ran", description: "Automated checks could read the expected result", passed: false, evidence: `${errorMessage(e)} — the task did not produce what the scenario expects. Task outcome: ${report.outcome}. ${report.errors?.length ? `Errors: ${report.errors.join(" | ")}. ` : ""}${report.executionSummary ?? ""}`.trim() }];
    }

    const render = report.actions.filter((a) => a.capability === "presentation.check_rendering").at(-1);
    const ro = render?.output as Record<string, any> | undefined;
    const visual: AcceptanceRecord["visual"] = !sc.visual?.length
      ? "not-applicable"
      : !ro
        ? "pending-manual"
        : !ro.supported
          ? "unsupported"
          : (ro.issues ?? []).some((i: any) => i.severity === "high")
            ? "rendered-check-failed"
            : "rendered-check-passed";

    return done({
      outcome: checks.every((c) => c.passed) ? "passed" : "failed",
      checks,
      visual,
      task: {
        outcome: report.outcome,
        summary: report.executionSummary,
        actions: report.actions.map((a) => ({ capability: a.capability, status: a.status, approval: a.approval, verified: a.verification?.verified, method: a.verification?.method, evidence: a.verification?.evidence?.slice(0, 200) })),
        completion: (report.completion ?? []).map((c) => ({ id: c.id, met: c.met, detail: c.detail })),
        limitations: report.limitations,
      },
    });
  } catch (e) {
    return done({ outcome: "error", reason: errorMessage(e), visual: "not-applicable" });
  }
}

const OUTCOME_LABEL: Record<AcceptanceRecord["outcome"], string> = { passed: "✅ passed", failed: "❌ failed", blocked: "⏸ blocked", error: "⚠️ error" };
const VISUAL_LABEL: Record<AcceptanceRecord["visual"], string> = {
  "rendered-check-passed": "rendered check passed",
  "rendered-check-failed": "rendered check failed",
  unsupported: "rendering unsupported",
  "pending-manual": "needs a person",
  "not-applicable": "—",
};

/**
 * Markdown status matrix from records. Live and mocked results are kept in
 * separate tables; a scenario with no live record is shown as "not run".
 */
export function acceptanceMatrix(records: AcceptanceRecord[], scenarios: { id: string; title: string; host: string }[]): string {
  const latest = (mode: AcceptanceMode) => {
    const m = new Map<string, AcceptanceRecord[]>();
    for (const r of records.filter((x) => x.mode === mode)) {
      const key = `${r.scenarioId}|${r.environment.officeApp}|${r.environment.platform ?? ""}|${r.environment.officeVersion ?? ""}|${r.model.kind}:${r.model.name ?? ""}`;
      const prev = m.get(key)?.[0];
      if (!prev || prev.environment.recordedAt < r.environment.recordedAt) m.set(key, [r]);
    }
    return [...m.values()].map((x) => x[0]!);
  };
  const live = latest("live");
  const mocked = latest("mocked");
  const row = (r: AcceptanceRecord) =>
    `| ${r.scenarioId} | ${r.environment.officeApp} ${r.environment.officeVersion ?? ""} | ${r.environment.platform ?? ""} | ${r.model.kind === "scripted" ? "scripted" : `${r.model.name}`} | ${r.environment.ollamaVersion ?? "—"} | ${OUTCOME_LABEL[r.outcome]}${r.reason ? ` (${r.reason})` : ""} | ${r.checks.filter((c) => c.passed).length}/${r.checks.length} | ${VISUAL_LABEL[r.visual]} | ${r.manual.filter((m) => m.status === "pending").length} pending | ${r.environment.recordedAt.slice(0, 10)} |`;
  const header = "| Scenario | Office | Platform | Model | Ollama | Automated outcome | Checks | Visual | Manual steps | Date |\n|---|---|---|---|---|---|---|---|---|---|";
  const notRun = scenarios.filter((s) => !live.some((r) => r.scenarioId === s.id));
  return [
    "### Live Office (real hosts)",
    live.length ? `${header}\n${live.map(row).join("\n")}` : "_No live runs recorded yet._",
    notRun.length ? `\nNot yet run live: ${notRun.map((s) => `\`${s.id}\``).join(", ")}.` : "",
    "",
    "### Mocked hosts (in-memory, CI)",
    mocked.length ? `${header}\n${mocked.map(row).join("\n")}` : "_No mocked runs recorded._",
    "",
    "A mocked pass says the runtime, specialists and checks behave as designed against in-memory hosts. It is not evidence that the scenario works in Office.",
  ].join("\n");
}
