import { useEffect, useState } from "react";
import { ALL_SPECIALISTS, type AssistantSettings, type HistoryStore, type TaskRecord } from "@loai/core";
import type { HostInfo } from "../services/host";
import type { CompanionClient } from "../services/companion";
import { SharePanel } from "./Sources";
import { ActionLine, ApprovalCard, StatusBadge } from "./components";
import type { useAssistant } from "./useAssistant";

type A = ReturnType<typeof useAssistant>;

export function ContextView({ a, host, companion, sharing }: { a: A; host: HostInfo; companion: CompanionClient; sharing: boolean }) {
  const items = a.session.context.inspected();
  const [checking, setChecking] = useState(false);
  void a.contextVersion; // re-render on context changes

  const refresh = async () => {
    setChecking(true);
    for (const it of items) {
      try {
        a.session.context.setFreshness(it.id, await host.adapter.checkFreshness(it));
      } catch {
        a.session.context.setFreshness(it.id, "unknown");
      }
    }
    setChecking(false);
  };

  return (
    <div className="panel">
      <SharePanel a={a} companion={companion} available={sharing} />
      <p className="muted small">
        Exactly what the assistant has read in this session. Content is used for answers only within the scope shown here; everything else has not been inspected.
      </p>
      <div className="row gap">
        <button onClick={refresh} disabled={checking || !items.length}>
          {checking ? "Checking…" : "Check for changes"}
        </button>
        <button onClick={() => a.session.context.clear()} disabled={!items.length}>
          Forget context
        </button>
      </div>
      {!items.length && <p className="muted">Nothing inspected yet.</p>}
      {items.map((it) => (
        <details key={it.id} className="ctx">
          <summary>
            <span className={`scope scope-${it.scope}`}>{it.scope === "source" ? "shared" : it.scope}</span> {it.label}
            {it.truncated && <span className="pill">partial</span>}
            {it.freshness !== "fresh" && <span className={`pill ${it.freshness === "stale" ? "pill-warn" : ""}`}>{it.freshness === "stale" ? "may have changed" : "unchecked"}</span>}
          </summary>
          {it.unsupported.length > 0 && <p className="small warn">Not readable: {it.unsupported.map((u) => `${u.count} ${u.kind}`).join(", ")}</p>}
          <pre className="ctx-body">{it.content.slice(0, 4000)}{it.content.length > 4000 ? "\n…" : ""}</pre>
          <div className="muted small">
            Read {new Date(it.retrievedAt).toLocaleTimeString()} via {it.source}
          </div>
        </details>
      ))}
    </div>
  );
}

export function ActionsView({ a }: { a: A }) {
  return (
    <div className="panel">
      {a.pending && <ApprovalCard pending={a.pending} />}
      {!a.actions.length && !a.pending && <p className="muted">No actions yet. Proposed and executed operations appear here with their verification result.</p>}
      {a.actions.map((r) => (
        <ActionLine
          key={r.id}
          record={r}
          onRetry={["failed", "unverified", "cancelled"].includes(r.status) ? () => a.execute({ id: r.id + "-retry", name: r.capability, arguments: r.args }) : undefined}
          onUndo={r.status === "succeeded" || r.status === "unverified" ? (r.compensation ? () => a.execute({ id: r.id + "-undo", name: r.compensation!.capability, arguments: r.compensation!.args }) : undefined) : undefined}
        />
      ))}
      {a.actions.some((r) => r.status === "succeeded" && !r.compensation) && (
        <p className="muted small">Some changes have no automatic undo. Use Office's own Undo (Ctrl/Cmd+Z) where available.</p>
      )}
    </div>
  );
}

export function HistoryView({ history, settings }: { history: HistoryStore | null; settings: AssistantSettings }) {
  const [records, setRecords] = useState<TaskRecord[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [confirmClear, setConfirmClear] = useState(false);

  const load = async () => {
    if (!history) return;
    try {
      setRecords(await history.list(100));
      setError(null);
    } catch (e) {
      setError((e as Error).message);
    }
  };
  useEffect(() => {
    load();
  }, [history]);

  if (!history)
    return (
      <div className="panel">
        <p className="muted">History is stored by the local companion. Pair the companion in Settings to keep a history across sessions.</p>
      </div>
    );
  return (
    <div className="panel">
      <div className="row between">
        <span className="muted small">
          {settings.historyEnabled ? `Kept ${settings.historyRetentionDays} days · ${settings.historyStoreFullArgs ? "full details" : "references only"}` : "History is off"}
        </span>
        {confirmClear ? (
          <span className="row gap">
            <button
              className="danger"
              onClick={async () => {
                await history.clear();
                setConfirmClear(false);
                load();
              }}
            >
              Delete all
            </button>
            <button onClick={() => setConfirmClear(false)}>Cancel</button>
          </span>
        ) : (
          <button onClick={() => setConfirmClear(true)} disabled={!records?.length}>
            Clear history
          </button>
        )}
      </div>
      {error && <p className="err small">{error}</p>}
      {records?.length === 0 && <p className="muted">No tasks yet.</p>}
      {records?.map((r) => (
        <details key={r.id} className="hist">
          <summary>
            <span className={`badge badge-${r.outcome === "completed" ? "succeeded" : r.outcome === "partial" ? "partial" : r.outcome === "cancelled" ? "cancelled" : "failed"}`}>{r.outcome}</span> {r.request.slice(0, 80)}
            <div className="muted small">
              {new Date(r.startedAt).toLocaleString()} · {r.host}
              {r.specialist ? ` · ${ALL_SPECIALISTS.find((x) => x.id === r.specialist)?.name ?? r.specialist}` : ""}
              {r.itemTitle ? ` · ${r.itemTitle}` : ""} · {r.model}
            </div>
          </summary>
          {r.contextRefs.length > 0 && <div className="small">Context: {r.contextRefs.map((c) => c.label).join("; ")}</div>}
          {r.actions.map((x) => (
            <div key={x.id} className="row between small">
              <span>{x.target || x.capability}</span>
              <StatusBadge status={x.status} approval={x.approval} />
            </div>
          ))}
          {r.completion?.some((c) => !c.met) && <div className="small warn">Not met: {r.completion.filter((c) => !c.met).map((c) => c.description).join("; ")}</div>}
          {r.errors.length > 0 && <div className="small err">{r.errors.join(" · ")}</div>}
          <div className="small muted hist-text">{r.finalText}</div>
        </details>
      ))}
    </div>
  );
}
