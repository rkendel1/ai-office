import { isApprovedPlan, planFrom } from "../shared/planning.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist, TaskEvidence } from "../types.js";
import { preservesUnrelatedContent } from "../shared/completion.js";
import { layoutCheckedAfterChanges, onlyNewSlides, reportsLayoutIssues } from "./shared.js";

/** Slide count the user asked for ("10-slide", "ten slides"), if any. */
export function requestedSlideCount(request: string): number | undefined {
  const words: Record<string, number> = { three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9, ten: 10, eleven: 11, twelve: 12, fifteen: 15, twenty: 20 };
  const m = /\b(\d{1,2}|three|four|five|six|seven|eight|nine|ten|eleven|twelve|fifteen|twenty)[- ](?:\w+[- ])?slides?\b/i.exec(request);
  if (!m) return undefined;
  return /^\d+$/.test(m[1]!) ? Number(m[1]) : words[m[1]!.toLowerCase()];
}

const created = (e: TaskEvidence) => e.actions.filter((a) => a.capability === "presentation.create_slide" && (a.status === "succeeded" || a.status === "partial"));

export const presentationAuthor: Specialist = {
  id: "powerpoint.author",
  name: "Presentation Author",
  host: "powerpoint",
  categories: ["create deck", "pitch", "narrative", "slides from brief"],
  description: "Builds a complete, editable presentation from a brief or pasted material: narrative arc, slide plan for approval, then real slides with text, diagrams and charts.",
  examples: [
    "Create a 10-slide executive presentation on our product strategy: customer problem, solution, market, differentiation, roadmap, risks and investment required",
    "Build a 6-slide investor update from the notes I paste below",
    "Add three slides summarising the roadmap at the end of this deck",
  ],
  abilities: ["Plans the narrative and every slide for your approval", "Creates slides from layouts with titles, bullets and notes", "Adds diagrams and charts built from editable shapes", "Re-reads every slide and checks layout"],
  context: { required: [REQUIRE.deckOverview, REQUIRE.deck], optional: ["Material you paste into the request"] },
  instructions: [
    "You are the Presentation Author. You produce an actual, editable PowerPoint presentation — never just an outline in chat.",
    "1. Clarify audience, purpose and length only if they are missing AND matter; otherwise state your assumptions in the plan.",
    "2. Source material is the request text, the open deck, and any content the user shared from another Office app and attached (OFFICE_CONTENT items with from_app). Use attached content as the primary source and cite it. You cannot open other documents yourself: if the user refers to one that is not attached, ask them to share its selection from that app (Share selection in its task pane) or paste it. Do not invent figures — use placeholders like \"[figure to confirm]\".",
    "3. Develop a narrative arc, then call plan.propose with one step per slide (title, key message, content and visual), the outcome, the target (this presentation, new slides at the end), risks, and `operations`: one presentation.create_slide per slide with scope {slideTitle}, plus each visual (add_diagram / add_bar_chart / add_table) with scope {slideTitle} of the slide it goes on. Use exactly these titles when creating the slides.",
    "4. After approval, call presentation.list_layouts once, then create each slide with presentation.create_slide (title, 3–5 concise bullets of ≤12 words, speaker notes when useful). Use a \"Title Slide\" layout with a subtitle for the opening slide.",
    "5. Add visuals where the plan calls for them: presentation.add_diagram for processes/roadmaps, presentation.add_bar_chart only with numbers from the source, presentation.add_table for comparisons. Do not put a diagram on a slide that already has long bullets; use the \"Title Only\" layout for visual slides.",
    "6. When all slides exist, run presentation.check_layout. Fix high-severity issues you can (shorten text with presentation.edit_text, move/resize with presentation.set_geometry), check again, and report anything that remains, including speaker notes that could not be added.",
  ].join("\n"),
  capabilities: [
    "context.read",
    "presentation.read_slides",
    "presentation.list_layouts",
    "presentation.check_layout",
    "plan.propose",
    "presentation.create_slide",
    "presentation.add_text_box",
    "presentation.add_shape",
    "presentation.add_diagram",
    "presentation.add_bar_chart",
    "presentation.add_table",
    "presentation.set_notes",
    "presentation.format_text",
    "presentation.edit_text",
    "presentation.set_geometry",
  ],
  constraints: {
    maxToolCalls: 60,
    requirePlan: true,
    readOnly: false,
    rules: ["The slide plan must be approved before any slide is created", "Slides are added at the end; existing slides are not changed unless you ask", "Only figures from your material are charted"],
  },
  approval: { summary: "You approve the slide plan, then each slide and visual as it is created." },
  completion: [
    { id: "plan-approved", description: "The slide plan was approved before slides were created", required: true, evaluate: (e) => ({ met: e.actions.some(isApprovedPlan), detail: e.actions.some(isApprovedPlan) ? "approved" : "no approved plan" }) },
    {
      id: "slides-created",
      description: "The requested number of slides was created and verified",
      required: true,
      evaluate: (e) => {
        const n = created(e).length;
        const want = requestedSlideCount(e.request) ?? planFrom(e.actions.find(isApprovedPlan) ?? ({} as any))?.steps.length ?? 2;
        const failed = e.actions.filter((a) => a.capability === "presentation.create_slide" && (a.status === "failed" || a.status === "unverified")).length;
        return { met: n >= want, detail: `${n} of ${want} slide(s) created${failed ? `; ${failed} failed or unverified` : ""}` };
      },
    },
    {
      id: "distinct-content",
      description: "Each slide has a distinct title and content",
      required: true,
      evaluate: (e) => {
        const titles = created(e).map((a) => String((a.args as any).title).trim().toLowerCase());
        const bodies = created(e).map((a) => JSON.stringify([(a.args as any).bullets ?? null, (a.args as any).body ?? null]));
        const dupT = titles.filter((t, i) => titles.indexOf(t) !== i);
        const dupB = bodies.filter((b, i) => b !== "[null,null]" && bodies.indexOf(b) !== i);
        return { met: titles.length > 0 && !dupT.length && !dupB.length, detail: dupT.length || dupB.length ? `repeated: ${[...dupT, ...(dupB.length ? ["identical slide content"] : [])].join(", ")}` : `${titles.length} distinct slide(s)` };
      },
    },
    layoutCheckedAfterChanges(),
    reportsLayoutIssues,
    {
      id: "notes-limitation-reported",
      description: "Speaker notes that could not be added are reported",
      required: true,
      evaluate: (e) => {
        const missing = created(e).filter((a) => (a.output as any)?.notes === "not_supported").length;
        if (!missing) return { met: true, detail: "no unsupported notes" };
        const said = /notes?\b[^.]*\b(not (be )?(supported|added|available)|could not|couldn't|unavailable)/i.test(e.text);
        return { met: said, detail: said ? `${missing} slide(s) without notes, reported` : `${missing} slide(s) could not get notes, but this was not reported` };
      },
    },
    preservesUnrelatedContent(),
  ],
  verification: ["Existing slides are compared before and after", "Each slide is re-read after creation (title, bullets, notes where supported)", "Charts are checked for correct labels and proportional bar heights", "A layout check inspects text fit, overlaps and off-slide shapes (objects only, not the rendered slide)"],
  artifacts: ["Editable slides in the open presentation", "An approved slide plan", "Speaker notes where the host supports them"],
  limitations: [
    "Cannot open Word documents or workbooks itself: share the selection from Word or Excel (via the companion) or paste it",
    "Charts are drawn with editable shapes, not native chart objects",
    "Speaker notes depend on host support (not available through Office.js today)",
    "Themes, images and animations are not created",
    "Visual rendering is not inspected; only slide objects are checked",
  ],
  routing: {
    patterns: [
      { re: /\b(create|build|generate|write|draft|prepare|put together|make (?:a|an|me an?|new))\b[^.?!]*\b(presentation|deck|slides?|pitch)\b/i, weight: 3 },
      { re: /\b(\d{1,2}|three|four|five|six|seven|eight|nine|ten|twelve)[- ](\w+[- ])?slides?\b/i, weight: 2 },
      { re: /\b(investor|pitch|executive|board|sales|kickoff)\s+(deck|presentation)\b/i, weight: 1 },
      { re: /\badd\b[^.?!]*\bslides?\b/i, weight: 2 },
    ],
  },
  snapshot: true,
  guard: (call, s) => onlyNewSlides(call, s.actions, s.request),
};
