import type { HostAdapter } from "../hosts/adapter.js";
import type { ApprovalRequest, ApprovalResponse } from "../permissions.js";
import type { TaskReport } from "../runtime.js";
import type { ScriptStep } from "../testing/fakes.js";
import type { HostKind } from "../types.js";

/**
 * How a record was produced. Only "live" records come from real Office. A
 * mocked pass never counts as a live pass.
 */
export type AcceptanceMode = "mocked" | "live";

export interface ScenarioContext<D> {
  driver: D;
  adapter: HostAdapter;
}

export interface CheckResult {
  id: string;
  description: string;
  passed: boolean;
  /** What was observed (read back from the application). */
  evidence: string;
}

/**
 * One acceptance scenario, usable three ways: automated against in-memory
 * hosts (CI), automated inside real Office (live runner), and as a manual
 * checklist when the host is not available.
 */
export interface AcceptanceScenario<D = any> {
  id: string;
  host: HostKind;
  title: string;
  specialist?: string;
  purpose: string;
  /** Human-readable preconditions (also the manual checklist's first steps). */
  preconditions: string[];
  /** Returns a reason when the scenario must not run here (e.g. the document is not blank). */
  checkPreconditions(c: ScenarioContext<D>): Promise<string | undefined>;
  /** Create the fixture content through the driver. */
  setup(c: ScenarioContext<D>): Promise<void>;
  request: string;
  /** Deterministic model responses (built after setup, so they can use real ids). */
  script(c: ScenarioContext<D>): Promise<ScriptStep[]>;
  /** Scripted user decisions on approval prompts. */
  approve(req: ApprovalRequest): ApprovalResponse;
  /** Automated checks against application state and the task report. */
  checks(c: ScenarioContext<D>, report: TaskReport): Promise<CheckResult[]>;
  /** Steps a person must do and record (visual inspection, Sent Items, ...). */
  manual: string[];
  /** Manual steps that are visual verification (shown separately in the matrix). */
  visual?: string[];
}

export interface AcceptanceEnvironment {
  officeApp: string;
  officeVersion?: string;
  platform?: string;
  os?: string;
  addinVersion: string;
  companionVersion?: string;
  ollamaVersion?: string;
  recordedAt: string;
}

export interface AcceptanceRecord {
  schema: 1;
  scenarioId: string;
  title: string;
  host: HostKind;
  specialist?: string;
  mode: AcceptanceMode;
  /** "scripted" = deterministic responses (tests the Office integration); "ollama" = a real local model. */
  model: { kind: "scripted" | "ollama"; name?: string };
  environment: AcceptanceEnvironment;
  /** passed = every automated check passed; manual checks are tracked separately. */
  outcome: "passed" | "failed" | "blocked" | "error";
  reason?: string;
  checks: CheckResult[];
  task?: {
    outcome: string;
    summary: string;
    actions: { capability: string; status: string; approval?: string; verified: boolean | null | undefined; method?: string; evidence?: string }[];
    completion: { id: string; met: boolean; detail?: string }[];
    limitations: string[];
  };
  /** Visual verification status: rendered checks done by the app, or still owed by a person. */
  visual: "rendered-check-passed" | "rendered-check-failed" | "unsupported" | "pending-manual" | "not-applicable";
  manual: { step: string; status: "pending" | "passed" | "failed"; note?: string }[];
  durationMs: number;
}
