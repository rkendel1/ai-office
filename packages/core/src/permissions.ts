import type { ActionPreview, Json, PolicyDecision, RiskClass } from "./types.js";

/** Per-risk-class decision. See spec FR-07. */
export type ApprovalPolicy = Record<RiskClass, PolicyDecision>;

export const DEFAULT_POLICY: Readonly<ApprovalPolicy> = Object.freeze({
  read_selection: "allow", // permitted after the user invokes the assistant
  read_active: "allow", // permitted within the disclosed scope
  read_external: "confirm", // explicit request required
  modify: "approve", // preview + approval by default
  destructive: "confirm", // explicit confirmation
  send: "confirm", // explicit confirmation
  prohibited: "deny", // never
});

const STRICTNESS: Record<PolicyDecision, number> = { allow: 0, approve: 1, confirm: 2, deny: 3 };

/**
 * Minimum strictness per risk class. A user (or a corrupted settings blob) may
 * make policy stricter, but never looser than these floors.
 */
const FLOOR: Record<RiskClass, PolicyDecision> = {
  read_selection: "allow",
  read_active: "allow",
  read_external: "confirm",
  modify: "allow", // user may opt in to auto-apply edits (with a visible warning)
  destructive: "confirm",
  send: "confirm",
  prohibited: "deny",
};

export function stricter(a: PolicyDecision, b: PolicyDecision): PolicyDecision {
  return STRICTNESS[a] >= STRICTNESS[b] ? a : b;
}

/** Merge user overrides into the default policy and enforce floors. Unknown keys/values are ignored. */
export function sanitizePolicy(overrides?: Partial<Record<string, unknown>>): ApprovalPolicy {
  const out: ApprovalPolicy = { ...DEFAULT_POLICY };
  for (const risk of Object.keys(DEFAULT_POLICY) as RiskClass[]) {
    const v = overrides?.[risk];
    if (typeof v === "string" && v in STRICTNESS) out[risk] = v as PolicyDecision;
    out[risk] = stricter(out[risk], FLOOR[risk]);
  }
  return out;
}

export function decide(policy: ApprovalPolicy, risk: RiskClass): PolicyDecision {
  // Re-apply the floor at decision time so a mutated policy object cannot loosen it.
  return stricter(policy[risk] ?? "deny", FLOOR[risk] ?? "deny");
}

export interface ApprovalRequest {
  capability: string;
  risk: RiskClass;
  decision: "approve" | "confirm";
  preview: ActionPreview;
  /** Validated arguments of the call being approved (used for plan-scope matching; not shown raw). */
  args?: Record<string, Json>;
  /**
   * Present on plan approvals when bounded batch execution can be offered:
   * how many declared operations could run without further prompts and how
   * many will still ask because policy requires confirmation.
   */
  planBatch?: { operations: number; autoEligible: number; alwaysAsk: number };
}

export interface ApprovalResponse {
  approved: boolean;
  reason?: string;
  /** For plan approvals: the user chose bounded batch execution of the declared operations. */
  batch?: boolean;
  /** Set by the runtime when a declared plan operation was covered by an approved batch plan. */
  via?: "plan";
}

/**
 * Supplied by the UI. Only a human interaction may resolve it. The model has
 * no channel to call or influence this function.
 */
export type Approver = (req: ApprovalRequest, signal?: AbortSignal) => Promise<ApprovalResponse>;

/** Approver used in non-interactive contexts: rejects everything. */
export const denyAllApprover: Approver = async () => ({ approved: false, reason: "No interactive approver available" });
