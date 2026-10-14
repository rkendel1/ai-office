export * from "./types.js";
export * from "./registry.js";
export * from "./router.js";
export * from "./shared/planning.js";
export * from "./shared/evidence.js";
export * from "./shared/completion.js";
export * from "./shared/requirements.js";
export * from "./word/index.js";
export * from "./excel/index.js";
export * from "./powerpoint/index.js";
export * from "./outlook/index.js";
export { requestedSlideCount } from "./powerpoint/presentation-author.js";
export { factTokens } from "./word/editor.js";

import type { Specialist } from "./types.js";

/** Specialist section of the system prompt. Internal: never shown in the UI. */
export function specialistPrompt(s: Specialist, available: string[], unavailable: string[], planRequired: boolean): string {
  return [
    `Active specialist: ${s.name}. ${s.description}`,
    s.instructions,
    "",
    "Rules for this task:",
    ...s.constraints.rules.map((r) => `- ${r}`),
    ...(planRequired ? ["- Call plan.propose and wait for approval before any change; changes before an approved plan are refused."] : []),
    `- You can use only these capabilities: ${available.join(", ")}.`,
    ...(unavailable.length ? [`- Not available in this Office version: ${unavailable.join(", ")}. Report affected parts as limitations.`] : []),
    "",
    "When you finish, the application checks (do not claim these yourself):",
    ...s.completion.map((c) => `- ${c.description}`),
    "",
    "Limitations to state when relevant:",
    ...s.limitations.map((l) => `- ${l}`),
  ].join("\n");
}
