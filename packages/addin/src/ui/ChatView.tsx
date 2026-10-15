import { useEffect, useMemo, useRef, useState } from "react";
import type { AssistantSettings, SharedSource, SourceSummary, Specialist, TaskReport } from "@loai/core";
import type { CompanionClient } from "../services/companion";
import { SourcesBar } from "./Sources";
import type { Connection } from "./App";
import { ActionLine, ApprovalCard } from "./components";
import { AUTO, ClarifyCard, CompletionList, GENERAL, PlanCard, PlanProgress, RouteHint, SpecialistCard, SpecialistPicker } from "./Specialists";
import type { useAssistant } from "./useAssistant";

const STATE_TEXT: Record<string, string> = {
  gathering_context: "Reading context…",
  thinking: "Thinking…",
  awaiting_approval: "Waiting for your approval",
  executing: "Applying…",
};

export function ChatView({
  a,
  busy,
  settings,
  conn,
  onOpenSettings,
  companion,
  sharing,
}: {
  a: ReturnType<typeof useAssistant>;
  busy: boolean;
  settings: AssistantSettings;
  conn: Connection;
  onOpenSettings: () => void;
  companion: CompanionClient;
  sharing: boolean;
}) {
  const [attached, setAttached] = useState<SourceSummary[]>([]);
  const [text, setText] = useState("");
  const [choice, setChoice] = useState(AUTO);
  const [clarify, setClarify] = useState<{ request: string; question: string; options: Specialist[]; sources: SharedSource[] } | null>(null);
  const [routeError, setRouteError] = useState<string | null>(null);
  const end = useRef<HTMLDivElement>(null);
  const picker = useRef<HTMLDivElement>(null);
  const input = useRef<HTMLTextAreaElement>(null);
  useEffect(() => end.current?.scrollIntoView({ block: "end" }), [a.items.length, a.pending, a.items[a.items.length - 1], clarify]);

  const host = a.session.adapter.host;
  const specialists = useMemo(() => a.session.specialists.forHost(host), [a.session, host]);
  const selected = choice !== AUTO && choice !== GENERAL ? a.session.specialists.get(choice) : undefined;
  const preview = useMemo(() => (choice === AUTO && text.trim().length >= 6 ? a.session.router.route(text, host, AUTO) : null), [a.session, text, host, choice]);

  const fill = (t: string) => {
    setText(t);
    input.current?.focus();
  };
  /** Fetch the full content of attached sources; drop any that expired. */
  const loadSources = async (): Promise<SharedSource[] | null> => {
    const out: SharedSource[] = [];
    for (const s of attached) {
      try {
        out.push(await companion.getSource(s.id));
      } catch {
        setAttached(attached.filter((x) => x.id !== s.id));
        setRouteError(`“${s.scopeLabel}” is no longer shared (it expires after 30 minutes). Share it again or detach it.`);
        return null;
      }
    }
    return out;
  };
  const send = async (t = text) => {
    const v = t.trim();
    if (!v || busy) return;
    setRouteError(null);
    const route = a.session.router.route(v, host, choice);
    if (route.kind === "rejected") return setRouteError(route.reason);
    const sources = await loadSources();
    if (!sources) return;
    setText("");
    if (route.kind === "clarify") return setClarify({ request: v, question: route.question, options: route.options, sources });
    setClarify(null);
    if (route.kind === "specialist") a.run(v, route.specialist, route.how === "matched" ? route.reason : "You selected this specialist.", sources);
    else a.run(v, undefined, undefined, sources);
  };
  const lastReport = [...a.items].reverse().find((i) => i.kind === "report") as { report: TaskReport; request: string; specialist?: Specialist } | undefined;
  const notReady = !settings.model || conn.state === "error" || conn.state === "unpaired";

  return (
    <div className="chat">
      <div className="messages" aria-live="polite">
        {a.items.length === 0 && (
          <div className="empty">
            <p>
              I work only on this computer, using your local model. I can read what's open in {a.session.adapter.host === "powerpoint" ? "PowerPoint" : a.session.adapter.host[0]!.toUpperCase() + a.session.adapter.host.slice(1)}, propose changes for you to approve, and check
              that they were applied.
            </p>
            {notReady ? (
              <button className="primary" onClick={onOpenSettings}>
                {!settings.model ? "Choose a model" : conn.state === "unpaired" ? "Pair with the companion" : "Fix connection"}
              </button>
            ) : selected ? (
              <SpecialistCard specialist={selected} onExample={fill} />
            ) : (
              <div className="spec-grid">
                <p className="small">Describe what you want, or start from one of these (you can edit before sending):</p>
                {specialists.map((s) => (
                  <button key={s.id} className="spec-tile" onClick={() => fill(s.examples[0]!)}>
                    <b>{s.name}</b>
                    <span className="small muted">{s.examples[0]}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        )}
        {a.items.map((it) => {
          switch (it.kind) {
            case "user":
              return (
                <div key={it.id} className="msg user">
                  {it.text}
                </div>
              );
            case "assistant":
              return (
                <div key={it.id} className="msg assistant">
                  <div className="msg-text">{it.text}{it.streaming && <span className="caret" />}</div>
                  {!it.streaming && it.text && (
                    <button className="link small" onClick={() => navigator.clipboard?.writeText(it.text)}>
                      Copy
                    </button>
                  )}
                </div>
              );
            case "action":
              return <ActionLine key={it.id} record={it.record} />;
            case "notice":
              return (
                <div key={it.id} className={`notice notice-${it.level}`}>
                  {it.text}
                </div>
              );
            case "specialist":
              return (
                <div key={it.id} className="handled small">
                  <b>{it.specialist.name}</b> <span className="muted">· {it.specialist.description}</span>
                  {it.reason && <div className="muted">{it.reason}</div>}
                </div>
              );
            case "plan":
              return <PlanCard key={it.id} plan={it.plan} batch={it.batch} />;
            case "report":
              return (
                <ReportCard
                  key={it.id}
                  report={it.report}
                  onRetry={it === lastReport && it.report.outcome !== "completed" && !it.report.awaitingInput ? () => a.run(it.request, it.specialist) : undefined}
                  onProceed={
                    it === lastReport && !busy && it.specialist && !it.specialist.constraints.readOnly && /No changes were made/i.test(it.report.executionSummary)
                      ? () => a.run("Yes, apply these changes now.", it.specialist, "Approved from the Apply changes button.")
                      : undefined
                  }
                />
              );
          }
        })}
        {a.pending && <ApprovalCard pending={a.pending} />}
        {clarify && (
          <ClarifyCard
            question={clarify.question}
            options={clarify.options}
            onChoose={(sp) => {
              const { request: req, sources } = clarify;
              setClarify(null);
              a.run(req, sp ?? undefined, sp ? "You chose this specialist." : undefined, sources);
            }}
          />
        )}
        <div ref={end} />
      </div>
      <div className="composer">
        {busy && (
          <div className="row between working">
            <span>
              <span className="spinner" /> {STATE_TEXT[a.state] ?? "Working…"}
              {a.stateDetail ? <span className="muted"> · {a.stateDetail}</span> : null}
            </span>
            <button onClick={a.cancel}>Stop</button>
          </div>
        )}
        <div className="row between" ref={picker}>
          <SpecialistPicker list={specialists} value={choice} onChange={setChoice} disabled={busy} />
          {selected && (
            <details className="about small">
              <summary>About</summary>
              <SpecialistCard specialist={selected} onExample={fill} />
            </details>
          )}
        </div>
        <SourcesBar companion={companion} available={sharing} attached={attached} setAttached={setAttached} host={host} />
        <textarea
          ref={input}
          value={text}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              send();
            }
          }}
          placeholder={busy ? "Working…" : selected ? `Ask the ${selected.name}…` : "Ask about this item or describe a change"}
          rows={3}
          disabled={busy}
          aria-label="Message"
        />
        <RouteHint route={preview} onPick={() => picker.current?.querySelector("select")?.focus()} />
        {routeError && <div className="small err">{routeError}</div>}
        <div className="row between">
          <button className="link small" onClick={a.newConversation} disabled={busy}>
            New conversation
          </button>
          <button className="primary" onClick={() => send()} disabled={busy || !text.trim()}>
            Send
          </button>
        </div>
      </div>
    </div>
  );
}

const OUTCOME: Record<string, string> = { completed: "Done", partial: "Partly done", failed: "Did not complete", cancelled: "Stopped" };

function ReportCard({ report, onRetry, onProceed }: { report: TaskReport; onRetry?: () => void; onProceed?: () => void }) {
  const used = report.contextUsed.filter((c) => c.scope !== "metadata");
  return (
    <div className={`report report-${report.awaitingInput ? "partial" : report.outcome}`}>
      <div className="row between">
        <strong>
          {report.awaitingInput ? "Needs your answer" : OUTCOME[report.outcome]}
          {report.specialist && <span className="muted"> · {report.specialist.name}</span>}
        </strong>
        {onRetry && (
          <button className="small-btn" onClick={onRetry}>
            Retry task
          </button>
        )}
      </div>
      <div className="small">{report.executionSummary}</div>
      {onProceed && (
        <div className="row gap">
          <button className="primary small-btn" onClick={onProceed}>
            Proceed — apply changes
          </button>
        </div>
      )}
      {report.planExecution && <PlanProgress exec={report.planExecution} />}
      <CompletionList report={report} />
      {report.limitations.length > 0 && (
        <details className="small">
          <summary>Limitations ({report.limitations.length})</summary>
          <ul>
            {report.limitations.map((l) => (
              <li key={l}>{l}</li>
            ))}
          </ul>
        </details>
      )}
      {report.errors.length > 0 && report.outcome !== "completed" && <div className="small err">{report.errors[report.errors.length - 1]}</div>}
      <div className="small muted">
        Context used: {used.length ? used.map((c) => `${c.label}${c.truncated ? " (partial)" : ""}`).join("; ") : "overview only"}
      </div>
    </div>
  );
}
