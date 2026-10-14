import type { Plan, PlanExecutionReport, RouteResult, Specialist, TaskReport } from "@loai/core";

export const AUTO = "auto";
export const GENERAL = "general";

/** Picker shown above the composer. One conversation hosts every specialist. */
export function SpecialistPicker({ list, value, onChange, disabled }: { list: Specialist[]; value: string; onChange: (id: string) => void; disabled?: boolean }) {
  return (
    <label className="picker">
      <span className="picker-label">Specialist</span>
      <select value={value} onChange={(e) => onChange(e.target.value)} disabled={disabled} aria-label="Specialist">
        <option value={AUTO}>Auto — pick from my request</option>
        {list.map((s) => (
          <option key={s.id} value={s.id}>
            {s.name}
          </option>
        ))}
        <option value={GENERAL}>General assistant</option>
      </select>
    </label>
  );
}

/** What the active specialist does, what it reads, and what needs approval. No prompts or schemas. */
export function SpecialistCard({ specialist, onExample }: { specialist: Specialist; onExample?: (text: string) => void }) {
  return (
    <div className="spec-card">
      <div className="spec-name">{specialist.name}</div>
      <p className="small">{specialist.description}</p>
      <ul className="spec-abilities small">
        {specialist.abilities.map((a) => (
          <li key={a}>{a}</li>
        ))}
      </ul>
      <div className="small muted">
        <b>Reads:</b> {specialist.context.required.map((r) => r.description).join("; ")}
      </div>
      <div className="small muted">
        <b>Approvals:</b> {specialist.approval.summary}
      </div>
      {onExample && (
        <div className="suggestions">
          {specialist.examples.map((ex) => (
            <button key={ex} className="chip" onClick={() => onExample(ex)} title="Insert into the message box to edit before sending">
              {ex}
            </button>
          ))}
        </div>
      )}
      <details className="small">
        <summary>Limitations</summary>
        <ul>
          {specialist.limitations.map((l) => (
            <li key={l}>{l}</li>
          ))}
        </ul>
      </details>
    </div>
  );
}

/** Live preview of routing while typing, so the user can switch before anything runs. */
export function RouteHint({ route, onPick }: { route: RouteResult | null; onPick: () => void }) {
  if (!route) return null;
  if (route.kind === "specialist")
    return (
      <div className="route small">
        Will use <b>{route.specialist.name}</b>
        {route.how === "matched" ? <span className="muted"> · {route.reason}</span> : null}{" "}
        <button className="link small" onClick={onPick}>
          change
        </button>
      </div>
    );
  if (route.kind === "clarify") return <div className="route small warn">I'll ask which specialist to use before starting.</div>;
  if (route.kind === "rejected") return <div className="route small err">{route.reason}</div>;
  return <div className="route small muted">Will use the general assistant.</div>;
}

export function ClarifyCard({ question, options, onChoose }: { question: string; options: Specialist[]; onChoose: (s: Specialist | null) => void }) {
  return (
    <div className="clarify" role="group" aria-label="Choose a specialist">
      <p className="small">{question}</p>
      <div className="row gap">
        {options.map((o) => (
          <button key={o.id} onClick={() => onChoose(o)}>
            {o.name}
          </button>
        ))}
        <button onClick={() => onChoose(null)}>General assistant</button>
      </div>
    </div>
  );
}

export function PlanCard({ plan, batch }: { plan: Plan; batch?: boolean }) {
  return (
    <div className="plan">
      <div className="small">
        <b>Approved plan{batch ? " (batch)" : ""}:</b> {plan.title}
      </div>
      {plan.outcome && <div className="small muted">Outcome: {plan.outcome}</div>}
      {plan.target && <div className="small muted">Target: {plan.target}</div>}
      <ol className="small">
        {plan.steps.map((s, i) => (
          <li key={i}>
            {s.title}
            {s.detail ? <span className="muted"> — {s.detail}</span> : null}
          </li>
        ))}
      </ol>
      <div className="small muted">
        {batch ? "The listed changes run without further prompts; anything else, or any failure, stops and asks." : "Each change in this plan is still shown to you for approval."}
      </div>
    </div>
  );
}

const STEP_LABEL: Record<string, string> = { done: "✓", failed: "✗ failed", unverified: "? unverified", partial: "◐ partial", rejected: "✗ rejected", not_run: "○ not run" };

export function PlanProgress({ exec }: { exec: PlanExecutionReport }) {
  return (
    <div className="plan-progress small" aria-label="Plan progress">
      <b>Plan steps</b>
      {exec.suspended && <div className="warn">Batch stopped: {exec.suspended}</div>}
      <ul>
        {exec.operations.map((o) => (
          <li key={o.index} className={`step-${o.status}`}>
            <span>{STEP_LABEL[o.status]}</span> {o.description}
            {o.calls.length > 0 && <span className="muted"> · {o.calls.map((c) => (c.approvedBy === "plan" ? "plan" : c.approvedBy === "user" ? "you" : c.approvedBy)).join(", ")}</span>}
          </li>
        ))}
      </ul>
      {exec.undeclared.length > 0 && <div className="muted">{exec.undeclared.length} change(s) outside the plan were approved separately.</div>}
    </div>
  );
}

export function CompletionList({ report }: { report: TaskReport }) {
  if (!report.completion?.length) return null;
  return (
    <ul className="checks small" aria-label="Completion checks">
      {report.completion.map((c) => (
        <li key={c.id} className={c.met ? "ok" : c.required ? "bad" : "soft"}>
          <span aria-hidden>{c.met ? "✓" : c.required ? "✗" : "–"}</span> {c.description}
          {c.detail && <span className="muted"> — {c.detail}</span>}
        </li>
      ))}
    </ul>
  );
}
