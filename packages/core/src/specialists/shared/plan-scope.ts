/**
 * Plan scope: the operations a plan declares, how calls are matched to them,
 * and the bounded grant that lets declared operations run without a prompt
 * when — and only when — the user approved the plan as a batch and policy
 * would otherwise ask for a simple approval ("approve", never "confirm").
 *
 * The grant never adds a capability, never widens a target, and is suspended
 * on the first scope change, failure or unverified result.
 */
import { parseRange } from "../../hosts/a1.js";
import type { CapabilityDefinition } from "../../registry.js";
import type { ExecutionRecord, Json, ToolCall } from "../../types.js";

export interface PlanScope {
  /** Excel: worksheet. For workbook.clean_to_sheet / add_worksheet: the new sheet. */
  sheet?: string;
  /** Excel: A1 range the operation may write inside. */
  range?: string;
  /** PowerPoint: title of the slide (created by this plan) the operation targets. */
  slideTitle?: string;
  /** Word: insertion position. */
  position?: "start" | "end";
  /** Word: paragraph range the operation may replace or format. */
  paragraphs?: { start: number; end: number };
}

export interface PlanOperation {
  capability: string;
  description: string;
  scope: PlanScope;
  /** How many calls this declaration covers (default 1). */
  count: number;
}

export type PlanStepStatus = "done" | "failed" | "unverified" | "partial" | "rejected" | "not_run";

export interface PlanOperationReport {
  index: number;
  capability: string;
  description: string;
  status: PlanStepStatus;
  /** Calls attributed to this operation, with how each was approved. */
  calls: { actionId: string; status: string; approvedBy: "plan" | "user" | "policy" | "none" }[];
  remaining: number;
}

export interface PlanExecutionReport {
  title: string;
  /** The user approved the plan for bounded batch execution. */
  batch: boolean;
  /** Why batch execution stopped early, if it did. */
  suspended?: string;
  operations: PlanOperationReport[];
  /** Modifying calls that were not part of the plan (each required its own approval). */
  undeclared: { actionId: string; capability: string; status: string }[];
}

const norm = (s: string) => s.trim().replace(/\s+/g, " ").toLowerCase();

/** Scope fields each capability needs before it can be declared as batch-eligible. */
const SCOPE_RULES: Record<string, (keyof PlanScope)[]> = {
  "presentation.create_slide": ["slideTitle"],
  "presentation.add_text_box": ["slideTitle"],
  "presentation.add_shape": ["slideTitle"],
  "presentation.add_diagram": ["slideTitle"],
  "presentation.add_bar_chart": ["slideTitle"],
  "presentation.add_table": ["slideTitle"],
  "presentation.set_notes": ["slideTitle"],
  "presentation.format_text": ["slideTitle"],
  "presentation.set_geometry": ["slideTitle"],
  "presentation.edit_text": ["slideTitle"],
  "workbook.add_worksheet": ["sheet"],
  "workbook.write_range": ["sheet", "range"],
  "workbook.create_table": ["sheet", "range"],
  "workbook.create_chart": ["sheet"],
  "workbook.clean_to_sheet": ["sheet"],
  "document.write_blocks": [],
  "document.insert_text": ["position"],
  "document.replace_text": ["paragraphs"],
  "document.format_text": ["paragraphs"],
  "outlook.create_draft": [],
};

export const PLAN_SCOPE_SCHEMA = {
  type: "object",
  properties: {
    sheet: { type: "string", minLength: 1, maxLength: 31 },
    range: { type: "string", minLength: 2, maxLength: 32 },
    slideTitle: { type: "string", minLength: 1, maxLength: 200 },
    position: { type: "string", enum: ["start", "end"] },
    paragraphs: {
      type: "object",
      required: ["start", "end"],
      properties: { start: { type: "integer", minimum: 0 }, end: { type: "integer", minimum: 0 } },
    },
  },
} as const;

export function parseOperations(args: Record<string, Json>): PlanOperation[] {
  const ops = Array.isArray(args.operations) ? args.operations : [];
  return ops
    .filter((o): o is Record<string, Json> => !!o && typeof o === "object" && !Array.isArray(o))
    .map((o) => ({
      capability: String(o.capability ?? ""),
      description: String(o.description ?? o.capability ?? ""),
      scope: (o.scope && typeof o.scope === "object" && !Array.isArray(o.scope) ? o.scope : {}) as PlanScope,
      count: typeof o.count === "number" && o.count >= 1 ? Math.floor(o.count) : 1,
    }));
}

/**
 * Problems that make a declared operation unacceptable in a plan: unknown or
 * internal capability, a read-only capability (nothing to authorise), a
 * capability outside the allowlist, or missing scope fields.
 */
export function operationProblem(op: PlanOperation, def: CapabilityDefinition | undefined, allowed: (name: string) => boolean): string | undefined {
  if (!def || def.internal) return `"${op.capability}" is not an available capability`;
  if (!allowed(op.capability)) return `"${op.capability}" is not available in this task`;
  if (!def.modifies) return `"${op.capability}" only reads; do not declare reads as plan operations`;
  const rule = SCOPE_RULES[op.capability];
  if (!rule) return `"${op.capability}" cannot be declared in a plan; it will be asked for separately`;
  const missing = rule.filter((k) => op.scope[k] === undefined);
  if (missing.length) return `"${op.capability}" needs scope ${missing.join(", ")}`;
  if (op.scope.range && !parseRange(op.scope.range)) return `"${op.scope.range}" is not a valid A1 range`;
  if (op.count > 50) return "count must be 50 or less";
  return undefined;
}

function inRange(outer: string | undefined, inner: unknown): boolean {
  if (!outer || typeof inner !== "string") return false;
  const o = parseRange(outer);
  const i = parseRange(inner.includes("!") ? inner.slice(inner.lastIndexOf("!") + 1) : inner);
  return !!o && !!i && i.start.row >= o.start.row && i.end.row <= o.end.row && i.start.col >= o.start.col && i.end.col <= o.end.col;
}

/** Slide ids created in this task for a planned slide title. */
function slidesTitled(title: string, actions: ExecutionRecord[]): string[] {
  return actions
    .filter((a) => a.capability === "presentation.create_slide" && (a.status === "succeeded" || a.status === "partial") && norm(String((a.args as any).title ?? "")) === norm(title))
    .map((a) => String((a.output as any)?.slideId));
}

/** Does `call` fall inside the declared operation's capability and scope? */
export function matchesOperation(op: PlanOperation, call: ToolCall, actions: ExecutionRecord[]): boolean {
  if (op.capability !== call.name) return false;
  const a = call.arguments as Record<string, any>;
  const s = op.scope;
  switch (call.name) {
    case "presentation.create_slide":
      return !!s.slideTitle && norm(String(a.title ?? "")) === norm(s.slideTitle);
    case "workbook.add_worksheet":
      return a.name === s.sheet;
    case "workbook.clean_to_sheet":
      return a.targetSheet === s.sheet;
    case "workbook.write_range":
    case "workbook.create_table":
      return a.sheet === s.sheet && inRange(s.range, a.address);
    case "workbook.create_chart":
      return (a.targetSheet ?? a.sourceSheet) === s.sheet;
    case "document.write_blocks":
      if (a.mode === "replace") return !!s.paragraphs && a.start === s.paragraphs.start && (a.end ?? a.start) === s.paragraphs.end;
      return (a.position ?? "end") === (s.position ?? "end") && a.paragraphIndex === undefined;
    case "document.insert_text":
      return a.position === s.position;
    case "document.replace_text":
      return a.target === "paragraphs" && !!s.paragraphs && a.start === s.paragraphs.start && (a.end ?? a.start) === s.paragraphs.end;
    case "document.format_text":
      return !!s.paragraphs && a.start >= s.paragraphs.start && (a.end ?? a.start) <= s.paragraphs.end;
    case "outlook.create_draft":
      return true;
    default:
      if (call.name.startsWith("presentation.") && s.slideTitle) return slidesTitled(s.slideTitle, actions).includes(String(a.slideId));
      return false;
  }
}

/** The bounded authorisation created when a plan is approved. */
export class PlanGrant {
  readonly operations: PlanOperation[];
  private used: number[];
  private calls: PlanOperationReport["calls"][];
  private readonly undeclared: PlanExecutionReport["undeclared"] = [];
  suspended?: string;

  constructor(
    readonly title: string,
    operations: PlanOperation[],
    /** User chose batch execution and policy permits it. */
    readonly batch: boolean,
  ) {
    this.operations = operations;
    this.used = operations.map(() => 0);
    this.calls = operations.map(() => []);
  }

  /** True while declared operations may run without a further prompt. */
  get active(): boolean {
    return this.batch && !this.suspended;
  }

  /** Index of the first declared operation with capacity left that covers `call`. */
  match(call: ToolCall, actions: ExecutionRecord[]): number | undefined {
    const i = this.operations.findIndex((op, k) => this.used[k]! < op.count && matchesOperation(op, call, actions));
    return i >= 0 ? i : undefined;
  }

  suspend(reason: string): void {
    if (this.batch && !this.suspended) this.suspended = reason;
  }

  /** Attribute an executed (or rejected) call to its operation, or record it as undeclared. */
  record(index: number | undefined, r: ExecutionRecord, modifies: boolean): void {
    const approvedBy = r.approval === "plan" ? "plan" : r.approval === "approved" ? "user" : r.approval === "not_required" ? "policy" : "none";
    if (index === undefined) {
      if (modifies && r.status !== "rejected") this.undeclared.push({ actionId: r.id, capability: r.capability, status: r.status });
      return;
    }
    this.calls[index]!.push({ actionId: r.id, status: r.status, approvedBy });
    if (r.status !== "rejected" || r.approval === "rejected") this.used[index]!++;
  }

  report(): PlanExecutionReport {
    return {
      title: this.title,
      batch: this.batch,
      suspended: this.suspended,
      undeclared: [...this.undeclared],
      operations: this.operations.map((op, i) => {
        const calls = this.calls[i]!;
        const statuses = calls.map((c) => c.status);
        const status: PlanStepStatus = !calls.length
          ? "not_run"
          : statuses.includes("failed")
            ? "failed"
            : statuses.includes("unverified")
              ? "unverified"
              : statuses.includes("partial")
                ? "partial"
                : calls.some((c) => c.status === "rejected") && !statuses.includes("succeeded")
                  ? "rejected"
                  : this.used[i]! < op.count
                    ? "partial"
                    : "done";
        return { index: i, capability: op.capability, description: op.description, status, calls, remaining: op.count - this.used[i]! };
      }),
    };
  }
}
