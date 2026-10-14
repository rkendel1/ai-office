import type { CapabilityDefinition } from "../../registry.js";
import type { ExecutionRecord, HostKind, Json } from "../../types.js";

export interface PlanStep {
  title: string;
  detail?: string;
}

export interface Plan {
  title: string;
  steps: PlanStep[];
}

const HOSTS: HostKind[] = ["word", "excel", "powerpoint", "outlook"];

/**
 * plan.propose: the model presents a plan (outline, slide plan, workbook
 * changes) for the user to approve before substantial work. Approving a plan
 * grants nothing: every change in it still goes through its own approval
 * according to policy. It always asks, even when policy would allow.
 */
type PlanArgs = { title: string; summary?: string; steps: PlanStep[] };

export const planCapability: CapabilityDefinition<any> = {
  name: "plan.propose",
  description:
    "Present a plan for the user's approval before making substantial changes: a title and ordered steps (e.g. a document outline, or one step per slide with its title and content). Approval of the plan does not approve the changes themselves.",
  hosts: HOSTS,
  risk: "modify",
  modifies: false,
  alwaysAsk: true,
  inputSchema: {
    type: "object",
    required: ["title", "steps"],
    properties: {
      title: { type: "string", minLength: 1, maxLength: 200 },
      summary: { type: "string", maxLength: 2000 },
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
    return {
      summary: `Plan: ${a.title}`,
      target: `${a.steps.length} step(s)`,
      after: [a.summary, ...a.steps.map((s, i) => `${i + 1}. ${s.title}${s.detail ? ` — ${s.detail}` : ""}`)].filter(Boolean).join("\n"),
      warnings: ["Approving the plan does not apply any change. Each change will still be shown for approval."],
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
  const a = r.args as { title?: Json; steps?: Json };
  if (typeof a.title !== "string" || !Array.isArray(a.steps)) return undefined;
  return {
    title: a.title,
    steps: a.steps.filter((s): s is Record<string, Json> => !!s && typeof s === "object" && !Array.isArray(s)).map((s) => ({ title: String(s.title ?? ""), detail: typeof s.detail === "string" ? s.detail : undefined })),
  };
}
