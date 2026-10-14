import type { TaskOutcome } from "../history.js";
import type { StateSnapshot } from "../hosts/adapter.js";
import type { ApprovalPolicy } from "../permissions.js";
import type { ContextItem, ExecutionRecord, HostKind, Json, ToolCall } from "../types.js";

/**
 * A context the specialist needs before the model is called. If it is not
 * already inspected, the runtime acquires it through the capability registry
 * (so the normal policy applies) using `acquire`.
 */
export interface ContextRequirement {
  id: string;
  /** User-facing: "The document text", "The email thread". */
  description: string;
  satisfiedBy(items: ContextItem[]): boolean;
  acquire?: { capability: string; arguments: Record<string, Json> };
}

/** Everything observed during one task; completion criteria are evaluated against it. */
export interface TaskEvidence {
  request: string;
  actions: ExecutionRecord[];
  /** The model's final message. */
  text: string;
  /** All context inspected during the task (initial, acquired, read by capabilities). */
  context: ContextItem[];
  before?: StateSnapshot;
  after?: StateSnapshot;
  outcome: TaskOutcome;
  /** Whether a capability changes Office state (from the registry). */
  isModifying(name: string): boolean;
}

export interface CriterionResult {
  met: boolean;
  detail?: string;
}

export interface CompletionCriterion {
  id: string;
  /** User-facing statement of what "done" means. */
  description: string;
  /** Required criteria downgrade a "completed" task to "partial" when unmet. */
  required: boolean;
  evaluate(e: TaskEvidence): CriterionResult;
}

/** State visible to a specialist's call guard. */
export interface GuardState {
  request: string;
  actions: ExecutionRecord[];
  planApproved: boolean;
}

export interface RoutingPattern {
  re: RegExp;
  weight: number;
}

export interface Specialist {
  /** Stable id: "<host>.<name>". */
  id: string;
  name: string;
  host: HostKind;
  categories: string[];
  /** One-sentence purpose shown in the task pane. */
  description: string;
  /** Example prompts the user can start from (and edit). */
  examples: string[];
  /** Plain-language abilities shown to the user ("Reads the document", "Proposes edits for approval"). */
  abilities: string[];
  context: { required: ContextRequirement[]; optional: string[] };
  /** Task-specific instructions appended to the base system prompt. */
  instructions: string;
  /**
   * Capability allowlist. Narrows the capabilities registered for the host;
   * names that are not registered are simply unavailable. It cannot add any.
   */
  capabilities: string[];
  constraints: {
    /** Lower than the user's limit only; the runtime takes the minimum. */
    maxToolCalls?: number;
    /** Require an approved plan (plan.propose) before any modifying action. */
    requirePlan?: boolean | ((request: string) => boolean);
    /** No modifying capabilities in the allowlist (checked by the registry). */
    readOnly: boolean;
    /** Constraints stated to the model and shown to the user. */
    rules: string[];
  };
  approval: {
    /** User-facing statement of what needs approval. */
    summary: string;
    /** Optional stricter policy. Can only tighten the user's policy. */
    policy?: Partial<ApprovalPolicy>;
  };
  completion: CompletionCriterion[];
  /** How results are verified, user-facing. */
  verification: string[];
  /** What the specialist produces. */
  artifacts: string[];
  /** What it cannot do. Shown to the user and stated to the model. */
  limitations: string[];
  routing: {
    patterns: RoutingPattern[];
    /** Specialists whose work this one includes (e.g. a reply writer analyses the thread first). */
    subsumes?: string[];
  };
  /** Capture host state before/after the task for preservation checks. */
  snapshot?: boolean;
  /** Additional narrowing of individual calls. Returns a rejection reason, or undefined. */
  guard?(call: ToolCall, state: GuardState): string | undefined;
}
