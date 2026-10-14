import { useEffect, useState } from "react";
import { DEFAULT_POLICY, isLoopbackUrl, type AssistantSettings, type InferenceProvider, type ModelDetails, type ModelInfo, type PolicyDecision, type RiskClass } from "@loai/core";
import type { CompanionClient } from "../services/companion";
import type { HostInfo } from "../services/host";
import type { Connection } from "./App";
import { Field } from "./components";

const RISK_LABEL: Record<RiskClass, string> = {
  read_selection: "Read the selection",
  read_active: "Read the open item",
  read_external: "Read other items",
  modify: "Change content",
  destructive: "Overwrite or delete substantial content",
  send: "Send or publish",
  prohibited: "Run code, shell, files, network",
};
/** Options offered per risk class; looser choices than the floor are never shown. */
const OPTIONS: Record<RiskClass, PolicyDecision[]> = {
  read_selection: ["allow", "approve"],
  read_active: ["allow", "approve"],
  read_external: ["confirm", "deny"],
  modify: ["approve", "allow"],
  destructive: ["confirm", "deny"],
  send: ["confirm", "deny"],
  prohibited: ["deny"],
};
const DECISION_LABEL: Record<PolicyDecision, string> = { allow: "Allow", approve: "Ask (preview)", confirm: "Always confirm", deny: "Never" };

type Props = {
  settings: AssistantSettings;
  setSettings: (s: AssistantSettings) => void;
  companion: CompanionClient;
  conn: Connection;
  onPairChanged: () => void;
  recheck: () => void;
  getProvider: () => InferenceProvider;
  host: HostInfo;
};

export function SettingsView({ settings: s, setSettings, companion, conn, onPairChanged, recheck, getProvider, host }: Props) {
  const set = (patch: Partial<AssistantSettings>) => setSettings({ ...s, ...patch });
  const [models, setModels] = useState<ModelInfo[] | null>(null);
  const [modelErr, setModelErr] = useState<string | null>(null);
  const [details, setDetails] = useState<ModelDetails | null>(null);
  const [code, setCode] = useState("");
  const [pairMsg, setPairMsg] = useState<string | null>(null);
  const [diag, setDiag] = useState<string[] | null>(null);
  const [pull, setPull] = useState({ name: "", confirm: false, progress: "" });

  const loadModels = async () => {
    setModelErr(null);
    try {
      setModels(await getProvider().listModels());
    } catch (e) {
      setModels(null);
      setModelErr((e as Error).message);
    }
  };
  useEffect(() => {
    if (conn.state === "ok") loadModels();
  }, [conn.state]);
  useEffect(() => {
    setDetails(null);
    if (s.model && conn.state === "ok") getProvider().describeModel(s.model).then(setDetails, () => setDetails(null));
  }, [s.model, conn.state]);

  const pair = async () => {
    setPairMsg(null);
    try {
      await companion.pair(code);
      setCode("");
      setPairMsg("Paired.");
      onPairChanged();
    } catch (e) {
      setPairMsg((e as Error).message);
    }
  };

  const runDiagnostics = async () => {
    const out: string[] = [];
    const p = getProvider();
    const ping = await p.ping();
    out.push(ping.ok ? `✓ Ollama reachable (${ping.version ?? "unknown version"})` : `✗ Ollama not reachable: ${ping.error}`);
    try {
      const ms = await p.listModels();
      out.push(`✓ ${ms.length} model(s) installed`);
    } catch (e) {
      out.push(`✗ Could not list models: ${(e as Error).message}`);
    }
    if (s.model) {
      try {
        const d = await p.describeModel(s.model);
        out.push(`✓ ${s.model}: context ${d.contextLength ?? "unknown"} tokens; ${d.capabilities.includes("tools") ? "native tool calling" : "text tool protocol"}`);
        if (d.contextLength && (s.generation.numCtx ?? 0) > d.contextLength) out.push(`! Context size ${s.generation.numCtx} exceeds the model maximum ${d.contextLength}; ${d.contextLength} will be used.`);
        const t0 = performance.now();
        // A fixed test prompt; no document content is sent.
        const r = await p.chat({ model: s.model, messages: [{ role: "user", content: "Reply with exactly: OK" }], options: { temperature: 0, numPredict: 8, numCtx: 2048 } });
        out.push(`${/ok/i.test(r.message.content) ? "✓" : "!"} Test response in ${Math.round(performance.now() - t0)} ms: “${r.message.content.trim().slice(0, 40)}”`);
      } catch (e) {
        out.push(`✗ Model check failed: ${(e as Error).message}`);
      }
    } else out.push("! No model selected");
    if (s.transport === "companion" && companion.paired) {
      try {
        const d = await companion.diagnostics();
        out.push(`✓ Companion ${d.companion} (${d.platform}, TLS ${d.tls ? "on" : "off"}), Ollama ${d.ollamaIsLocal ? "local" : "REMOTE"} at ${d.ollamaUrl}`);
      } catch (e) {
        out.push(`✗ Companion: ${(e as Error).message}`);
      }
    }
    for (const api of host.adapter.apiSupport()) out.push(`${api.supported ? "✓" : "✗"} ${api.requirement} — ${api.capabilities.join(", ")}`);
    setDiag(out);
  };

  const endpointRemote = s.transport === "direct" && !isLoopbackUrl(s.ollamaEndpoint);

  return (
    <div className="panel settings">
      <section>
        <h3>Connection</h3>
        <div className="row gap">
          <label className="radio">
            <input type="radio" checked={s.transport === "companion"} onChange={() => set({ transport: "companion" })} /> Companion (recommended)
          </label>
          <label className="radio">
            <input type="radio" checked={s.transport === "direct"} onChange={() => set({ transport: "direct" })} /> Ollama directly
          </label>
        </div>
        {s.transport === "companion" ? (
          <>
            <Field label="Companion URL">
              <input value={s.companionUrl} onChange={(e) => set({ companionUrl: e.target.value })} />
            </Field>
            {companion.paired ? (
              <div className="row between">
                <span className="small">Paired with this companion.</span>
                <button
                  onClick={() => {
                    companion.unpair();
                    onPairChanged();
                  }}
                >
                  Unpair
                </button>
              </div>
            ) : (
              <Field label="Pairing code" hint="Shown when the companion starts, or run: loai-companion pair">
                <div className="row gap">
                  <input value={code} onChange={(e) => setCode(e.target.value)} placeholder="ABCD-EFGH" />
                  <button className="primary" onClick={pair} disabled={code.replace(/\W/g, "").length < 8}>
                    Pair
                  </button>
                </div>
              </Field>
            )}
            {pairMsg && <p className="small">{pairMsg}</p>}
          </>
        ) : (
          <>
            <Field label="Ollama endpoint" hint="Ollama must allow this add-in's origin (OLLAMA_ORIGINS).">
              <input value={s.ollamaEndpoint} onChange={(e) => set({ ollamaEndpoint: e.target.value })} />
            </Field>
            {endpointRemote && (
              <label className="check warn">
                <input type="checkbox" checked={s.allowRemoteEndpoint} onChange={(e) => set({ allowRemoteEndpoint: e.target.checked })} />I understand this endpoint is not on this computer and document content will be sent to it.
              </label>
            )}
          </>
        )}
        <div className="row between">
          <span className="small">
            <span className={`dot dot-${conn.state}`} /> {conn.state === "ok" ? "Connected" : conn.state === "checking" ? "Checking…" : conn.state === "unpaired" ? "Not paired" : "Not connected"}
            {conn.detail ? ` — ${conn.detail}` : ""}
          </span>
          <button onClick={recheck}>Test</button>
        </div>
      </section>

      <section>
        <h3>Model</h3>
        <div className="row gap">
          <select value={s.model} onChange={(e) => set({ model: e.target.value })} aria-label="Model">
            <option value="">{models ? "Select a model" : "—"}</option>
            {s.model && !models?.some((m) => m.name === s.model) && <option value={s.model}>{s.model} (not found)</option>}
            {models?.map((m) => (
              <option key={m.name} value={m.name}>
                {m.name} {m.parameterSize ? `· ${m.parameterSize}` : ""} {m.quantization ? `· ${m.quantization}` : ""}
              </option>
            ))}
          </select>
          <button onClick={loadModels}>Refresh</button>
        </div>
        {modelErr && <p className="err small">{modelErr}</p>}
        {models?.length === 0 && <p className="small">No models installed. Install one below or with <code>ollama pull llama3.1:8b</code>.</p>}
        {details && (
          <p className="small muted">
            Context window {details.contextLength?.toLocaleString() ?? "unknown"} tokens · {details.capabilities.includes("tools") ? "native tool calling" : "no native tools (text protocol will be used)"}
          </p>
        )}
        {s.transport === "companion" && companion.paired && (
          <details>
            <summary className="small">Install a model…</summary>
            <div className="row gap">
              <input value={pull.name} placeholder="e.g. qwen2.5:7b" onChange={(e) => setPull({ ...pull, name: e.target.value, confirm: false })} />
              {!pull.confirm ? (
                <button onClick={() => setPull({ ...pull, confirm: true })} disabled={!pull.name.trim()}>
                  Install
                </button>
              ) : (
                <button
                  className="primary"
                  onClick={async () => {
                    try {
                      await companion.pullModel(pull.name.trim(), (st, pct) => setPull((p) => ({ ...p, progress: `${st}${pct !== undefined ? ` ${pct}%` : ""}` })));
                      setPull({ name: "", confirm: false, progress: "Installed." });
                      loadModels();
                    } catch (e) {
                      setPull((p) => ({ ...p, confirm: false, progress: (e as Error).message }));
                    }
                  }}
                >
                  Download {pull.name} now
                </button>
              )}
            </div>
            <p className="small muted">{pull.progress || "Downloads come from the Ollama registry and can be several GB."}</p>
          </details>
        )}
      </section>

      <section>
        <h3>Generation</h3>
        <div className="grid2">
          <Field label="Temperature">
            <input type="number" step="0.1" min="0" max="2" value={s.generation.temperature ?? 0.2} onChange={(e) => set({ generation: { ...s.generation, temperature: Number(e.target.value) } })} />
          </Field>
          <Field label="Context size (num_ctx)">
            <input type="number" step="1024" min="1024" value={s.generation.numCtx ?? 8192} onChange={(e) => set({ generation: { ...s.generation, numCtx: Number(e.target.value) } })} />
          </Field>
          <Field label="Max output tokens">
            <input type="number" min="16" placeholder="model default" value={s.generation.numPredict ?? ""} onChange={(e) => set({ generation: { ...s.generation, numPredict: e.target.value ? Number(e.target.value) : undefined } })} />
          </Field>
          <Field label="Request timeout (s)">
            <input type="number" min="5" value={Math.round(s.requestTimeoutMs / 1000)} onChange={(e) => set({ requestTimeoutMs: Number(e.target.value) * 1000 })} />
          </Field>
        </div>
        <label className="check">
          <input type="checkbox" checked={s.preferNativeTools} onChange={(e) => set({ preferNativeTools: e.target.checked })} /> Use native tool calling when the model supports it
        </label>
      </section>

      <section>
        <h3>Context and limits</h3>
        <div className="grid2">
          <Field label="Office content per request (chars)">
            <input type="number" step="1000" min="1000" value={s.contextBudgetChars} onChange={(e) => set({ contextBudgetChars: Number(e.target.value) })} />
          </Field>
          <Field label="Max actions per task">
            <input type="number" min="1" max="100" value={s.limits.maxToolCalls} onChange={(e) => set({ limits: { ...s.limits, maxToolCalls: Number(e.target.value) } })} />
          </Field>
          <Field label="Max minutes per task">
            <input type="number" min="1" value={Math.round(s.limits.maxDurationMs / 60000)} onChange={(e) => set({ limits: { ...s.limits, maxDurationMs: Number(e.target.value) * 60000 } })} />
          </Field>
          <Field label="Retries per failing action">
            <input type="number" min="0" max="5" value={s.limits.retryLimit} onChange={(e) => set({ limits: { ...s.limits, retryLimit: Number(e.target.value) } })} />
          </Field>
        </div>
      </section>

      <section>
        <h3>Approvals</h3>
        {(Object.keys(DEFAULT_POLICY) as RiskClass[]).map((r) => (
          <div key={r} className="row between policy">
            <span className="small">{RISK_LABEL[r]}</span>
            <select value={s.policy[r]} disabled={OPTIONS[r].length === 1} onChange={(e) => set({ policy: { ...s.policy, [r]: e.target.value as PolicyDecision } })} aria-label={RISK_LABEL[r]}>
              {OPTIONS[r].map((d) => (
                <option key={d} value={d}>
                  {DECISION_LABEL[d]}
                </option>
              ))}
            </select>
          </div>
        ))}
        {s.policy.modify === "allow" && <p className="warn small">Edits will be applied without a preview. Substantial overwrites still require confirmation.</p>}
        <label className="check">
          <input type="checkbox" checked={s.planBatchApproval} onChange={(e) => set({ planBatchApproval: e.target.checked })} /> Offer “approve plan as a batch” for multi-step plans (only the declared changes, on the declared targets; confirmations still ask)
        </label>
      </section>

      <section>
        <h3>History</h3>
        <label className="check">
          <input type="checkbox" checked={s.historyEnabled} onChange={(e) => set({ historyEnabled: e.target.checked })} /> Keep a local task history
        </label>
        <Field label="Keep for (days)">
          <input type="number" min="0" value={s.historyRetentionDays} onChange={(e) => set({ historyRetentionDays: Number(e.target.value) })} />
        </Field>
        <label className="check">
          <input type="checkbox" checked={s.historyStoreFullArgs} onChange={(e) => set({ historyStoreFullArgs: e.target.checked })} /> Store full action details (may include document text)
        </label>
        {companion.paired && (
          <button className="small-btn" onClick={() => companion.putSettings(s).catch(() => undefined)}>
            Apply history settings to companion
          </button>
        )}
      </section>

      <section>
        <h3>Privacy</h3>
        <ul className="small privacy">
          <li>Document content is sent only to the model endpoint shown above{s.transport === "companion" ? " (via the companion on this computer)" : ""}. {endpointRemote ? <b>This endpoint is NOT on this computer.</b> : "It is on this computer."}</li>
          <li>There is no cloud fallback. If the local model fails, the task fails.</li>
          <li>Office.js is loaded from Microsoft's CDN, as Office requires; it does not receive your content from this add-in.</li>
          <li>No telemetry is collected. Diagnostics never include document content.</li>
          <li>History stores task references and short excerpts unless you enable full details.</li>
        </ul>
      </section>

      <section>
        <h3>Diagnostics</h3>
        <button onClick={runDiagnostics}>Run diagnostics</button>
        {diag && (
          <pre className="diag">
            {diag.join("\n")}
          </pre>
        )}
      </section>
    </div>
  );
}
