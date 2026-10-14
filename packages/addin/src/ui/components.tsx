import { useState } from "react";
import { wordDiff, type ExecutionRecord, type ExecutionStatus } from "@loai/core";
import type { PendingApproval } from "./useAssistant";

export const APPROVAL_LABEL: Record<string, string> = { plan: "approved in plan", approved: "approved by you", not_required: "no approval needed", rejected: "rejected" };

export const STATUS_LABEL: Record<ExecutionStatus, string> = {
  succeeded: "Verified",
  unverified: "Not verified",
  failed: "Failed",
  partial: "Partial",
  cancelled: "Cancelled",
  rejected: "Not run",
};

export function StatusBadge({ status, approval }: { status: ExecutionStatus; approval?: ExecutionRecord["approval"] }) {
  const label = status === "rejected" && approval === "rejected" ? "Rejected" : STATUS_LABEL[status];
  return <span className={`badge badge-${status}`}>{label}</span>;
}

export function Diff({ before, after }: { before?: string; after?: string }) {
  if (before === undefined && after === undefined) return null;
  if (before === undefined) return <pre className="diff">{<ins>{after}</ins>}</pre>;
  if (after === undefined) return <pre className="diff">{<del>{before}</del>}</pre>;
  const ops = wordDiff(before, after);
  return (
    <pre className="diff" aria-label="Proposed change">
      {ops.map((o, i) => (o.kind === "same" ? <span key={i}>{o.text}</span> : o.kind === "add" ? <ins key={i}>{o.text}</ins> : <del key={i}>{o.text}</del>))}
    </pre>
  );
}

export function ApprovalCard({ pending }: { pending: PendingApproval }) {
  const { request } = pending;
  const [showFull, setShowFull] = useState(false);
  const confirm = request.decision === "confirm";
  const batch = request.planBatch;
  return (
    <div className={`approval ${confirm ? "approval-confirm" : ""}`} role="dialog" aria-label="Approve change">
      <div className="approval-head">
        <strong>{confirm ? "Confirm this change" : "Review this change"}</strong>
        <span className="muted">{request.preview.target}</span>
      </div>
      <p>{request.preview.summary}</p>
      {request.preview.warnings?.map((w, i) => (
        <p key={i} className="warn">
          ⚠ {w}
        </p>
      ))}
      <div className={showFull ? "approval-body full" : "approval-body"}>
        <Diff before={request.preview.before} after={request.preview.after} />
      </div>
      {(request.preview.before?.length ?? 0) + (request.preview.after?.length ?? 0) > 600 && (
        <button className="link" onClick={() => setShowFull((v) => !v)}>
          {showFull ? "Show less" : "Show all"}
        </button>
      )}
      {batch ? (
        <>
          <div className="row gap">
            <button className="primary" onClick={() => pending.resolve({ approved: true, batch: true })} autoFocus>
              Approve plan and run {batch.autoEligible} change{batch.autoEligible === 1 ? "" : "s"}
            </button>
            <button onClick={() => pending.resolve({ approved: true, batch: false })}>Approve, ask me at each step</button>
            <button onClick={() => pending.resolve({ approved: false })}>Reject</button>
          </div>
          <p className="muted small">
            A batch approval covers only the {batch.operations} change{batch.operations === 1 ? "" : "s"} listed, on the targets listed. Each is still checked and verified.
            {batch.alwaysAsk ? ` ${batch.alwaysAsk} will still ask because your approval settings require confirmation.` : ""} Anything outside the plan, or any failure, stops the batch and asks you.
          </p>
        </>
      ) : (
        <>
          <div className="row gap">
            <button className={confirm ? "danger" : "primary"} onClick={() => pending.resolve({ approved: true })} autoFocus>
              {request.capability === "plan.propose" ? "Approve plan" : confirm ? "Confirm" : "Apply"}
            </button>
            <button onClick={() => pending.resolve({ approved: false })}>Reject</button>
          </div>
          <p className="muted small">
            {request.capability === "plan.propose" ? "Approving the plan applies nothing; each change will be shown for approval." : "Nothing changes unless you approve. The document text cannot approve actions."}
          </p>
        </>
      )}
    </div>
  );
}

export function ActionLine({ record, onRetry, onUndo }: { record: ExecutionRecord; onRetry?: () => void; onUndo?: () => void }) {
  const [open, setOpen] = useState(false);
  return (
    <div className="action">
      <div className="row between">
        <button className="link action-title" onClick={() => setOpen((v) => !v)} aria-expanded={open}>
          {record.target || record.capability}
        </button>
        <StatusBadge status={record.status} approval={record.approval} />
      </div>
      {record.error && record.status !== "succeeded" && <div className="small err">{record.error}</div>}
      {record.verification?.evidence && record.status === "unverified" && <div className="small muted">{record.verification.evidence}</div>}
      {open && (
        <div className="action-detail small">
          <div>
            <b>Capability</b> {record.capability} · <b>Risk</b> {record.risk} · <b>Approval</b> {record.approval ? APPROVAL_LABEL[record.approval] : "—"}
          </div>
          {record.verification && (
            <div>
              <b>Verification</b> {record.verification.method}:{" "}
              {record.verification.verified === null ? "not possible" : record.verification.verified ? "matched" : `mismatch — ${record.verification.mismatch ?? ""}`}
            </div>
          )}
          <pre className="args">{JSON.stringify(record.args, null, 1).slice(0, 1500)}</pre>
        </div>
      )}
      {(onRetry || onUndo) && (
        <div className="row gap">
          {onRetry && (
            <button className="small-btn" onClick={onRetry}>
              Retry
            </button>
          )}
          {onUndo && record.compensation && (
            <button className="small-btn" onClick={onUndo} title={record.compensation.description}>
              Undo
            </button>
          )}
        </div>
      )}
    </div>
  );
}

export function Field({ label, children, hint }: { label: string; children: React.ReactNode; hint?: string }) {
  return (
    <label className="field">
      <span className="field-label">{label}</span>
      {children}
      {hint && <span className="muted small">{hint}</span>}
    </label>
  );
}
