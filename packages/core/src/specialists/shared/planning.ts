import type { CapabilityDefinition } from "../../registry.js";
import type { ExecutionRecord, HostKind, Json } from "../../types.js";
import { PLAN_SCOPE_SCHEMA, parseOperations, type PlanOperation } from "./plan-scope.js";

export interface PlanStep {
  title: string;
  detail?: string;
}

export interface Plan {
  title: string;
  /** The intended result, in the user's terms. */
  outcome?: string;
  /** The document, workbook, deck or message affected. */
  target?: string;
  risks?: string[];
  steps: PlanStep[];
  /** Declared write operations and their scope (what a batch approval covers). */
  operations: PlanOperation[];
}

const HOSTS: HostKind[] = ["word", "excel", "powerpoint", "outlook"];

/**
 * plan.propose: the model presents a plan (outline, slide plan, workbook
 * changes) for the user to approve before substantial work. Approving a plan
 * grants nothing: every change in it still goes through its own approval
 * according to policy. It always asks, even when policy would allow.
 */
type PlanArgs = { title: string; outcome: string; target: string; summary?: string; risks?: string[]; steps: PlanStep[]; operations?: { capability: string; description: string; scope?: Record<string, Json>; count?: number }[] };

export const planCapability: CapabilityDefinition<any> = {
  name: "plan.propose",
  description:
    "Present a plan for the user's approval before making substantial changes: title, the intended outcome, the target (document / workbook / presentation and where in it), risks, ordered steps (e.g. one per section or slide), and `operations`: every write you intend to make, each with its capability, a one-line description and its scope (slideTitle for slides you will create; sheet and range for cells; position or paragraphs for Word). The user may approve the plan as a bounded batch: then only those declared operations, within their scope, run without further prompts. Anything else will be asked separately and stops batch execution.",
  hosts: HOSTS,
  risk: "modify",
  modifies: false,
  alwaysAsk: true,
  inputSchema: {
    type: "object",
    required: ["title", "outcome", "target", "steps"],
    properties: {
      title: { type: "string", minLength: 1, maxLength: 200 },
      outcome: { type: "string", minLength: 1, maxLength: 1000 },
      target: { type: "string", minLength: 1, maxLength: 300 },
      summary: { type: "string", maxLength: 2000 },
      risks: { type: "array", maxItems: 10, items: { type: "string", minLength: 1, maxLength: 300 } },
      operations: {
        type: "array",
        maxItems: 60,
        items: {
          type: "object",
          required: ["capability", "description", "scope"],
          properties: {
            capability: { type: "string", minLength: 1, maxLength: 80 },
            description: { type: "string", minLength: 1, maxLength: 300 },
            scope: PLAN_SCOPE_SCHEMA as any,
            count: { type: "integer", minimum: 1, maximum: 50 },
          },
        },
      },
      steps: {
        type: "array",
        minItems: 1,
        maxItems: 40,
        items: {
          type: "object",
          required: ["title"],
          properties: { title: { type: "string", minLength: 1, maxLength: 200 }, detail: { type: "string", maxLength: 1000 } },
        },
      },
    },
  },
  resolveTarget: (a: PlanArgs) => `plan “${a.title}”`,
  async preview(a: PlanArgs) {
    const ops = parseOperations(a as unknown as Record<string, Json>);
    return {
      summary: `Plan: ${a.title}`,
      target: a.target,
      after: [
        `Outcome: ${a.outcome}`,
        a.summary,
        "Steps:",
        ...a.steps.map((s, i) => `${i + 1}. ${s.title}${s.detail ? ` — ${s.detail}` : ""}`),
        ...(ops.length ? ["Changes it will make:", ...ops.map((o) => `• ${o.description}${o.count > 1 ? ` (×${o.count})` : ""}`)] : []),
      ]
        .filter(Boolean)
        .join("\n"),
      warnings: [...(a.risks ?? []).map((r) => `Risk: ${r}`)],
    };
  },
  async execute(a: PlanArgs) {
    return { output: { approved: true, title: a.title, steps: a.steps.length } };
  },
};

export function isApprovedPlan(r: ExecutionRecord): boolean {
  return r.capability === "plan.propose" && r.status === "succeeded" && r.approval === "approved";
}

export function planFrom(r: ExecutionRecord): Plan | undefined {
  if (r.capability !== "plan.propose") return undefined;
  const a = r.args as { title?: Json; steps?: Json; outcome?: Json; target?: Json; risks?: Json };
  if (typeof a.title !== "string" || !Array.isArray(a.steps)) return undefined;
  return {
    title: a.title,
    outcome: typeof a.outcome === "string" ? a.outcome : undefined,
    target: typeof a.target === "string" ? a.target : undefined,
    risks: Array.isArray(a.risks) ? a.risks.map(String) : undefined,
    operations: parseOperations(r.args),
    steps: a.steps.filter((s): s is Record<string, Json> => !!s && typeof s === "object" && !Array.isArray(s)).map((s) => ({ title: String(s.title ?? ""), detail: typeof s.detail === "string" ? s.detail : undefined })),
  };
}
