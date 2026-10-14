import { useState } from "react";
import { ACCEPTANCE_SCENARIOS, runAcceptance, type AcceptanceEnvironment, type AcceptanceRecord, type AssistantSettings, type InferenceProvider } from "@loai/core";
import type { CompanionClient } from "../services/companion";
import type { HostInfo } from "../services/host";
import { ADDIN_VERSION } from "../services/version";

async function environment(host: HostInfo, provider: InferenceProvider, companion: CompanionClient): Promise<AcceptanceEnvironment> {
  const diag = typeof Office !== "undefined" ? Office.context?.diagnostics : undefined;
  const ollama = await provider.ping().catch(() => ({ ok: false, version: undefined }));
  const comp = companion.paired ? await companion.diagnostics().catch(() => undefined) : undefined;
  return {
    officeApp: host.demo ? `${host.label}` : String(diag?.host ?? host.label),
    officeVersion: diag?.version,
    platform: diag?.platform ? String(diag.platform) : host.demo ? "browser (demo)" : undefined,
    os: typeof navigator !== "undefined" ? navigator.userAgent : undefined,
    addinVersion: ADDIN_VERSION,
    companionVersion: comp?.companion,
    ollamaVersion: ollama.ok ? ollama.version : undefined,
    recordedAt: new Date().toISOString(),
  };
}

const OUTCOME: Record<AcceptanceRecord["outcome"], string> = { passed: "Passed (automated checks)", failed: "Failed", blocked: "Blocked", error: "Error" };

/**
 * Developer tool: run acceptance scenarios inside the real Office host and
 * record the result with its environment. Runs only on request, in a blank
 * test document; manual and visual steps are recorded by the person.
 */
export function AcceptancePanel({ host, companion, getProvider, settings }: { host: HostInfo; companion: CompanionClient; getProvider: () => InferenceProvider; settings: AssistantSettings }) {
  const scenarios = ACCEPTANCE_SCENARIOS.filter((s) => s.host === host.adapter.host);
  const [confirmed, setConfirmed] = useState(false);
  const [running, setRunning] = useState<string | null>(null);
  const [record, setRecord] = useState<AcceptanceRecord | null>(null);
  const [saved, setSaved] = useState<string | null>(null);

  const run = async (id: string, withModel: boolean) => {
    const sc = scenarios.find((s) => s.id === id)!;
    setRunning(id);
    setRecord(null);
    setSaved(null);
    try {
      const provider = getProvider();
      const r = await runAcceptance(sc, {
        driver: host.driver,
        adapter: host.adapter,
        mode: host.demo ? "mocked" : "live",
        environment: await environment(host, provider, companion),
        model: withModel ? { provider, name: settings.model, settings: { generation: settings.generation, requestTimeoutMs: settings.requestTimeoutMs } } : undefined,
      });
      setRecord(r);
    } finally {
      setRunning(null);
    }
  };

  const setManual = (i: number, status: "passed" | "failed" | "pending") =>
    record && setRecord({ ...record, manual: record.manual.map((m, k) => (k === i ? { ...m, status } : m)) });

  const download = () => {
    const blob = new Blob([JSON.stringify(record, null, 2)], { type: "application/json" });
    const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(blob), download: `${record!.scenarioId}-${record!.mode}.json` });
    a.click();
    URL.revokeObjectURL(a.href);
  };

  return (
    <section>
      <h3>Acceptance tests (developers)</h3>
      <p className="small muted">
        Runs a scripted scenario {host.demo ? "against the demo document (recorded as mocked)" : "in this Office document and records it as a live result"}. Scenarios add fixture content: use a new blank document.
      </p>
      <label className="check">
        <input type="checkbox" checked={confirmed} onChange={(e) => setConfirmed(e.target.checked)} /> This is a blank test document
      </label>
      {scenarios.map((s) => (
        <details key={s.id} className="small">
          <summary>{s.title}</summary>
          <p>{s.purpose}</p>
          <ul>
            {s.preconditions.map((p) => (
              <li key={p}>{p}</li>
            ))}
          </ul>
          <div className="row gap">
            <button disabled={!confirmed || !!running} onClick={() => run(s.id, false)}>
              {running === s.id ? "Running…" : "Run (scripted)"}
            </button>
            <button disabled={!confirmed || !!running || !settings.model} onClick={() => run(s.id, true)} title="Uses your local model instead of the script; results vary by model">
              Run with my model
            </button>
          </div>
        </details>
      ))}
      {record && (
        <div className="acc-result" aria-label="Acceptance result">
          <div className="small">
            <b>{OUTCOME[record.outcome]}</b> · {record.mode} · {record.model.kind === "scripted" ? "scripted responses" : record.model.name}
            {record.reason && <div className="warn">{record.reason}</div>}
          </div>
          <ul className="checks small">
            {record.checks.map((c) => (
              <li key={c.id} className={c.passed ? "ok" : "bad"}>
                <span aria-hidden>{c.passed ? "✓" : "✗"}</span> {c.description}
                <div className="muted">{c.evidence}</div>
              </li>
            ))}
          </ul>
          {record.visual !== "not-applicable" && <div className="small">Visual: {record.visual}</div>}
          {record.manual.length > 0 && (
            <div className="small">
              <b>Manual steps (record what you see)</b>
              {record.manual.map((m, i) => (
                <div key={m.step} className="row between">
                  <span>{m.step}</span>
                  <select value={m.status} onChange={(e) => setManual(i, e.target.value as any)} aria-label={`Result: ${m.step}`}>
                    <option value="pending">pending</option>
                    <option value="passed">passed</option>
                    <option value="failed">failed</option>
                  </select>
                </div>
              ))}
            </div>
          )}
          <div className="row gap">
            {companion.paired && (
              <button
                onClick={async () => {
                  try {
                    const r = await companion.call<{ stored: string }>("/v1/acceptance", { method: "POST", body: JSON.stringify(record) });
                    setSaved(`Saved as ${r.stored}`);
                  } catch (e) {
                    setSaved((e as Error).message);
                  }
                }}
              >
                Save record
              </button>
            )}
            <button onClick={download}>Download JSON</button>
          </div>
          {saved && <div className="small muted">{saved}</div>}
        </div>
      )}
    </section>
  );
}
