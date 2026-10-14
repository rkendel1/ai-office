import { useEffect, useRef, useState } from "react";
import type { AssistantSettings, TaskReport } from "@loai/core";
import type { Connection } from "./App";
import { ActionLine, ApprovalCard } from "./components";
import type { useAssistant } from "./useAssistant";

const STATE_TEXT: Record<string, string> = {
  gathering_context: "Reading context…",
  thinking: "Thinking…",
  awaiting_approval: "Waiting for your approval",
  executing: "Applying…",
};

const SUGGESTIONS: Record<string, string[]> = {
  word: ["Summarize this document", "Improve the clarity of the selected text", "List the action items"],
  excel: ["Describe the data on this sheet", "Add a summary sheet with totals and a chart", "Explain the formulas in my selection"],
  powerpoint: ["Summarize the argument of this deck", "Tighten the text on the selected slide", "Add a closing slide"],
  outlook: ["Summarize this email", "Extract the tasks and dates", "Draft a polite reply"],
};

export function ChatView({ a, busy, settings, conn, onOpenSettings }: { a: ReturnType<typeof useAssistant>; busy: boolean; settings: AssistantSettings; conn: Connection; onOpenSettings: () => void }) {
  const [text, setText] = useState("");
  const end = useRef<HTMLDivElement>(null);
  useEffect(() => end.current?.scrollIntoView({ block: "end" }), [a.items.length, a.pending, a.items[a.items.length - 1]]);

  const send = (t = text) => {
    const v = t.trim();
    if (!v || busy) return;
    setText("");
    a.run(v);
  };
  const lastReport = [...a.items].reverse().find((i) => i.kind === "report") as { report: TaskReport; request: string } | undefined;
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
            ) : (
              <div className="suggestions">
                {(SUGGESTIONS[a.session.adapter.host] ?? []).map((s) => (
                  <button key={s} className="chip" onClick={() => send(s)}>
                    {s}
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
            case "report":
              return <ReportCard key={it.id} report={it.report} onRetry={it === lastReport && it.report.outcome !== "completed" ? () => a.run(it.request) : undefined} />;
          }
        })}
        {a.pending && <ApprovalCard pending={a.pending} />}
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
        <textarea
          value={text}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              send();
            }
          }}
          placeholder={busy ? "Working…" : "Ask about this item or describe a change"}
          rows={3}
          disabled={busy}
          aria-label="Message"
        />
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

function ReportCard({ report, onRetry }: { report: TaskReport; onRetry?: () => void }) {
  const used = report.contextUsed.filter((c) => c.scope !== "metadata");
  return (
    <div className={`report report-${report.outcome}`}>
      <div className="row between">
        <strong>{OUTCOME[report.outcome]}</strong>
        {onRetry && (
          <button className="small-btn" onClick={onRetry}>
            Retry task
          </button>
        )}
      </div>
      <div className="small">{report.executionSummary}</div>
      {report.errors.length > 0 && report.outcome !== "completed" && <div className="small err">{report.errors[report.errors.length - 1]}</div>}
      <div className="small muted">
        Context used: {used.length ? used.map((c) => `${c.label}${c.truncated ? " (partial)" : ""}`).join("; ") : "overview only"}
      </div>
    </div>
  );
}
