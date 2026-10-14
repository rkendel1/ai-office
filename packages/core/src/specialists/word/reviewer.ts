import { disclosesPartialRead, inspected, noModifications, textCriterion } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist } from "../types.js";

export const documentReviewer: Specialist = {
  id: "word.reviewer",
  name: "Document Reviewer",
  host: "word",
  categories: ["review", "critique", "quality"],
  description: "Critically reviews the document for ambiguity, contradictions, gaps and unsupported claims, and gives an actionable findings list. Never edits.",
  examples: [
    "Review this document for contradictions and unsupported claims",
    "What is missing from this proposal before it goes to the board?",
    "Critique the structure and argument of this report",
  ],
  abilities: ["Reads the whole document (not just a selection)", "Cites paragraphs for each finding", "Separates observations from recommendations", "Does not change the document"],
  context: { required: [REQUIRE.wordOverview, REQUIRE.wordDocument], optional: [] },
  instructions: [
    "You are the Document Reviewer. You evaluate; you do not edit.",
    "Read the document (continue with context.read if it has more paragraphs than you were given). Look for: ambiguity, internal contradictions, missing sections expected for this document type, unsupported assertions, unclear ownership or dates, and structural problems.",
    "Reply with a numbered findings list, most important first. For each finding write two lines:",
    "  [Observation] what the document says or lacks, citing the paragraph (¶n) or heading, quoting a few words where useful.",
    "  [Recommendation] the concrete change you suggest.",
    "Only state as an observation what you actually read. If you inspected only part of the document, say so first.",
  ].join("\n"),
  capabilities: ["context.read", "context.search"],
  constraints: { readOnly: true, rules: ["Read-only: the document is never changed"] },
  approval: { summary: "No approvals needed: this specialist only reads." },
  completion: [
    inspected("read-document", "The document text was read", (c) => c.id.startsWith("word:paragraphs:")),
    noModifications(),
    textCriterion("findings", "Findings separate observations from recommendations", (t) => {
      const obs = (t.match(/\[Observation\]/gi) ?? []).length;
      const rec = (t.match(/\[Recommendation\]/gi) ?? []).length;
      return { met: obs >= 1 && rec >= 1, detail: `${obs} observation(s), ${rec} recommendation(s)` };
    }),
    textCriterion("cites-locations", "Findings reference paragraphs or headings", (t) => {
      const refs = (t.match(/¶\s?\d+|\bparagraphs? \d+|\bheading\b|“[^”]{3,}”|"[^"]{3,}"/gi) ?? []).length;
      return { met: refs >= 1, detail: `${refs} location reference(s)` };
    }),
    disclosesPartialRead(),
  ],
  verification: ["Checks that the document was actually read and that nothing was modified"],
  artifacts: ["Prioritised findings list in chat"],
  limitations: ["Cannot verify facts against outside sources", "Does not see images, charts or embedded objects"],
  routing: {
    patterns: [
      { re: /\b(review|critique|evaluate|assess|audit|feedback on|poke holes)\b/i, weight: 3 },
      { re: /\b(contradict\w*|ambigu\w*|unsupported|gaps?|missing (sections?|pieces|information)|weak(ness|nesses)?|inconsisten\w*)\b/i, weight: 2 },
      { re: /\bwhat('s| is) missing\b/i, weight: 2 },
    ],
  },
};
