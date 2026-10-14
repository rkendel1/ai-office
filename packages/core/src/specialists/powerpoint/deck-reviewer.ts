import { inspected, noModifications, textCriterion } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist } from "../types.js";
import { noRenderingClaims, reportsLayoutIssues } from "./shared.js";

export const deckReviewer: Specialist = {
  id: "powerpoint.reviewer",
  name: "Deck Reviewer",
  host: "powerpoint",
  categories: ["review", "narrative", "flow", "consistency"],
  description: "Reviews the whole presentation as a narrative — argument, flow, evidence, audience fit, consistency — and gives prioritised recommendations. Never edits.",
  examples: [
    "Review this deck for a board audience",
    "Does the story flow, and which claims lack evidence?",
    "Is the terminology used the same way throughout the narrative?",
  ],
  abilities: ["Reads every slide", "Runs an object-level layout check", "Prioritises recommendations", "Does not change the deck"],
  context: { required: [REQUIRE.deckOverview, REQUIRE.deck], optional: [] },
  instructions: [
    "You are the Deck Reviewer. You evaluate; you do not edit.",
    "1. Read every slide. Run presentation.check_layout to find text overflow, overlaps, missing titles, density and inconsistency.",
    "2. Assess: the argument and narrative flow, redundancy between slides, claims without supporting evidence, missing context for the stated audience, title and terminology consistency.",
    "3. Reply with prioritised recommendations, grouped as \"High\", \"Medium\" and \"Low\" priority, each citing the slide number (\"Slide 4\").",
    "4. Report every high-severity layout problem found by the check, citing its slide.",
    "5. Be explicit that layout findings come from slide objects; you have not seen the rendered slides.",
  ].join("\n"),
  capabilities: ["context.read", "presentation.read_slides", "presentation.check_layout"],
  constraints: { readOnly: true, rules: ["Read-only: the deck is never changed"] },
  approval: { summary: "No approvals needed: this specialist only reads." },
  completion: [
    inspected("read-deck", "Every slide was read", (c) => c.id === "ppt:slides:all"),
    {
      id: "layout-checked",
      description: "An object-level layout check was run",
      required: true,
      evaluate: (e) => ({ met: e.actions.some((a) => a.capability === "presentation.check_layout" && a.status === "succeeded"), detail: "presentation.check_layout" }),
    },
    textCriterion("prioritised", "Recommendations are prioritised and cite slides", (t) => {
      const pri = /\b(high|medium|low)\b/i.test(t) || /\bpriority\b/i.test(t);
      const cites = (t.match(/\bslides? \d+/gi) ?? []).length;
      return { met: pri && cites > 0, detail: `${pri ? "prioritised" : "no priorities"}, ${cites} slide reference(s)` };
    }),
    reportsLayoutIssues,
    noRenderingClaims,
    noModifications(),
  ],
  verification: ["Confirms the deck was read and the layout check ran", "Checks that serious problems found are reported", "Confirms nothing was modified"],
  artifacts: ["Prioritised recommendations in chat"],
  limitations: ["Cannot see the rendered slides, images or charts; layout findings come from object geometry and estimated text size", "Cannot verify facts against outside sources"],
  routing: {
    patterns: [
      { re: /\b(review|critique|evaluate|assess|feedback|audit)\b/i, weight: 3 },
      { re: /\b(narrative|flow|argument|story(line)?|redundan\w*|terminology|evidence|audience fit|consistency)\b/i, weight: 2 },
    ],
  },
};
