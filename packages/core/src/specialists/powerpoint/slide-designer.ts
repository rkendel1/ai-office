import { preservesUnrelatedContent, verifiedAction } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import { factTokens } from "../word/editor.js";
import type { Specialist } from "../types.js";
import { layoutCheckedAfterChanges, lastLayoutCheck } from "./shared.js";

const CHANGES = ["presentation.format_text", "presentation.set_geometry", "presentation.edit_text", "presentation.add_shape", "presentation.add_text_box"];

export const slideDesigner: Specialist = {
  id: "powerpoint.designer",
  name: "Slide Designer",
  host: "powerpoint",
  categories: ["visual design", "layout", "readability", "consistency"],
  description: "Improves the visual quality of existing slides — hierarchy, text density, alignment and consistency — without changing what they say.",
  examples: [
    "Make the selected slide more readable",
    "Make title sizes and alignment consistent across the deck",
    "Fix text that overflows or overlaps on these slides",
  ],
  abilities: ["Checks layout: overflow, overlap, density, consistency", "Resizes, moves and formats shapes", "Tightens text while keeping figures", "Re-checks layout after changes"],
  context: { required: [REQUIRE.deckOverview, REQUIRE.deck], optional: ["Your selected slides"] },
  instructions: [
    "You are the Slide Designer. Improve how slides look and read; keep their meaning and every number unchanged.",
    "1. Run presentation.check_layout first (on the selected slides if the user selected some, else the whole deck) and decide what to fix: overflow, overlaps, off-slide shapes, dense text, small text, inconsistent title sizes or alignment.",
    "2. Use a consistent design system: titles 32–36pt bold, body 18–24pt, left-aligned body, consistent margins (48pt).",
    "3. Fix with presentation.format_text, presentation.set_geometry, and — only to reduce density — presentation.edit_text with expectedText set to the current text. Keep all numbers, names and qualifications.",
    "4. Change only the slides in scope. Do not delete content.",
    "5. Run presentation.check_layout again and report what improved and what could not be checked (rendering is not visible to you).",
  ].join("\n"),
  capabilities: ["context.read", "presentation.read_slides", "presentation.check_layout", "presentation.format_text", "presentation.set_geometry", "presentation.edit_text", "presentation.add_shape", "presentation.add_text_box"],
  constraints: { readOnly: false, rules: ["Keeps wording meaning and all numbers", "Text edits must quote the current text", "No content is deleted"] },
  approval: { summary: "Each formatting, position or text change is shown for approval." },
  completion: [
    {
      id: "checked-first",
      description: "Layout was checked before changes",
      required: true,
      evaluate: (e) => {
        const first = e.actions.findIndex((a) => a.capability === "presentation.check_layout" && a.status === "succeeded");
        const change = e.actions.findIndex((a) => CHANGES.includes(a.capability) && a.status !== "rejected");
        return { met: first >= 0 && (change < 0 || first < change), detail: first >= 0 ? "checked first" : "no initial check" };
      },
    },
    verifiedAction("changes-verified", "Design changes were applied and verified", CHANGES),
    {
      id: "numbers-kept",
      description: "Text edits keep every number and name",
      required: true,
      evaluate: (e) => {
        const lost: string[] = [];
        for (const a of e.actions)
          if (a.capability === "presentation.edit_text" && a.status === "succeeded") {
            const args = a.args as Record<string, string>;
            for (const t of factTokens(args.expectedText ?? "")) if (!args.text!.includes(t)) lost.push(t);
          }
        return { met: !lost.length, detail: lost.length ? `lost: ${lost.join(", ")}` : "kept" };
      },
    },
    layoutCheckedAfterChanges(),
    {
      id: "fewer-issues",
      description: "The final layout check finds no more problems than the first",
      required: false,
      evaluate: (e) => {
        const checks = e.actions.filter((a) => a.capability === "presentation.check_layout" && a.status === "succeeded").map((a) => ((a.output as any)?.issues ?? []).length);
        const last = lastLayoutCheck(e.actions);
        return { met: checks.length >= 1 && (last?.issues.length ?? 0) <= (checks[0] ?? 0), detail: checks.length ? `${checks[0]} → ${last?.issues.length}` : "no checks" };
      },
    },
    preservesUnrelatedContent(),
  ],
  verification: ["Changed shapes are re-read (text, geometry, formatting)", "Untouched shapes are compared before and after", "Layout is re-checked at object level; rendering is not inspected"],
  artifacts: ["Restyled slides in place"],
  limitations: ["Cannot change the theme, master or fonts installed", "Cannot see the rendered slide; checks are based on shape geometry and estimated text size", "Pictures and charts can be moved or resized but not restyled"],
  routing: {
    patterns: [
      { re: /\b(design|layout|visual(ly)?|look(s)? (better|cleaner|professional)|align(ment)?|consistent|consistency|font sizes?|restyle|declutter|readab\w*)\b/i, weight: 3 },
      { re: /\b(overflow\w*|overlap\w*|too (much|many) (text|words)|dense|crowded|cluttered)\b/i, weight: 2 },
    ],
  },
  snapshot: true,
  guard(call) {
    if (call.name === "presentation.edit_text" && typeof call.arguments.expectedText !== "string")
      return "The Slide Designer must quote the shape's current text in expectedText so meaning is preserved.";
    return undefined;
  },
};
