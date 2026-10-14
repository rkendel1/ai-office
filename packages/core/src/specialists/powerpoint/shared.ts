import type { ExecutionRecord, ToolCall } from "../../types.js";
import type { CompletionCriterion } from "../types.js";

interface Issue {
  slide: number;
  severity: string;
  kind: string;
  detail: string;
}

export function lastLayoutCheck(actions: ExecutionRecord[]): { index: number; issues: Issue[] } | undefined {
  for (let i = actions.length - 1; i >= 0; i--) {
    const a = actions[i]!;
    if (a.capability === "presentation.check_layout" && a.status === "succeeded") return { index: i, issues: ((a.output as any)?.issues ?? []) as Issue[] };
  }
  return undefined;
}

/** The layout was checked after the last change. */
export function layoutCheckedAfterChanges(): CompletionCriterion {
  return {
    id: "layout-checked",
    description: "Slides were inspected (layout check) after the last change",
    required: true,
    evaluate(e) {
      const check = lastLayoutCheck(e.actions);
      let lastChange = -1;
      e.actions.forEach((a, i) => {
        if (e.isModifying(a.capability) && ["succeeded", "unverified", "partial"].includes(a.status)) lastChange = i;
      });
      if (!check) return { met: false, detail: "no layout check was run" };
      return { met: check.index > lastChange, detail: check.index > lastChange ? `${check.issues.length} issue(s) found in the final check` : "changes were made after the last check" };
    },
  };
}

/** Every high-severity issue found by the last layout check is reported to the user. */
export const reportsLayoutIssues: CompletionCriterion = {
  id: "reports-issues",
  description: "Serious layout or content problems found are reported",
  required: true,
  evaluate(e) {
    const check = lastLayoutCheck(e.actions);
    if (!check) return { met: false, detail: "no layout check" };
    const high = check.issues.filter((i) => i.severity === "high");
    const missing = high.filter((i) => i.slide > 0 && !new RegExp(`\\bslides?\\s*(\\d+\\s*(,|and|&)\\s*)*${i.slide}\\b`, "i").test(e.text));
    return { met: missing.length === 0, detail: high.length ? (missing.length ? `not reported: ${missing.map((i) => `slide ${i.slide} (${i.kind})`).join(", ")}` : `${high.length} serious issue(s) reported`) : "no serious issues" };
  },
};

/** Must not claim the rendered slides were checked: only objects are inspected. */
export const noRenderingClaims: CompletionCriterion = {
  id: "no-rendering-claims",
  description: "Does not claim the rendered slides were visually verified",
  required: true,
  evaluate(e) {
    const claim = /(?<!not |n't |never )\b(visually (verified|confirmed|checked)|rendering (was|has been|is) (verified|confirmed|checked)|i (looked at|viewed|saw) the (rendered )?slides?)\b/i.exec(e.text);
    return { met: !claim, detail: claim ? `claims “${claim[0]}”` : "no rendering claims" };
  },
};

/** Slides created by this task. */
export function createdSlideIds(actions: ExecutionRecord[]): string[] {
  return actions
    .filter((a) => a.capability === "presentation.create_slide" && (a.status === "succeeded" || a.status === "partial"))
    .map((a) => String((a.output as Record<string, unknown>)?.slideId));
}

const EDITS_EXISTING = /\b(edit|update|change|fix|revise|improve|redesign|rework|restyle|reformat)\b/i;

/** Changes may target only slides created in this task unless the user asked to change existing slides. */
export function onlyNewSlides(call: ToolCall, actions: ExecutionRecord[], request: string): string | undefined {
  const slideId = call.arguments.slideId;
  if (typeof slideId !== "string" || !call.name.startsWith("presentation.") || call.name === "presentation.read_slides" || call.name === "presentation.check_layout") return undefined;
  if (createdSlideIds(actions).includes(slideId) || EDITS_EXISTING.test(request)) return undefined;
  return "Only slides created in this task may be changed; the user did not ask to modify existing slides.";
}
