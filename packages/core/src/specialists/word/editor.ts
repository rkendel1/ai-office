import { preservesUnrelatedContent, verifiedAction } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import type { ContextRequirement, CriterionResult, Specialist, TaskEvidence } from "../types.js";

const EDITS = ["document.replace_text", "document.write_blocks", "document.format_text"];

const target: ContextRequirement = {
  id: "edit-target",
  description: "Your selection, or the document text when nothing is selected",
  satisfiedBy: (items) => items.some((c) => c.id === "word:selection" || c.id.startsWith("word:paragraphs:")),
  acquire: { capability: "context.read", arguments: { start: 0, count: 200 } },
};

/** Numbers and proper nouns in a passage (sentence-initial words are ignored). */
export function factTokens(text: string): string[] {
  const out = new Set<string>();
  for (const m of text.matchAll(/\d[\d,.]*%?/g)) out.add(m[0].replace(/[.,]$/, ""));
  for (const m of text.matchAll(/(?<![.!?]\s|^)\b([A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)*)/gm)) out.add(m[1]!);
  return [...out];
}

function factsKept(e: TaskEvidence): CriterionResult {
  if (/\b(remove|delete|cut|drop|omit|anonymi[sz]e|redact)\b/i.test(e.request)) return { met: true, detail: "removal was requested" };
  const lost: string[] = [];
  for (const a of e.actions) {
    if (a.status !== "succeeded") continue;
    const args = a.args as Record<string, any>;
    const before = a.capability === "document.replace_text" ? args.expectedText : a.capability === "document.write_blocks" && args.mode === "replace" ? args.expectedText : undefined;
    const after = a.capability === "document.replace_text" ? args.newText : a.capability === "document.write_blocks" ? (args.blocks as { text: string }[]).map((b) => b.text).join("\n") : undefined;
    if (typeof before !== "string" || typeof after !== "string") continue;
    for (const t of factTokens(before)) if (!after.includes(t)) lost.push(t);
  }
  return { met: lost.length === 0, detail: lost.length ? `missing after the edit: ${[...new Set(lost)].slice(0, 6).join(", ")}` : "numbers and names kept" };
}

export const documentEditor: Specialist = {
  id: "word.editor",
  name: "Document Editor",
  host: "word",
  categories: ["edit", "rewrite", "proofread", "tone", "concision"],
  description: "Improves clarity, tone, grammar or concision of the passage you point to, without changing its meaning or touching anything else.",
  examples: [
    "Make the selected paragraph more concise",
    "Fix grammar and spelling in this section without changing the wording more than necessary",
    "Rewrite paragraph 4 in a more formal tone",
    "Tighten the risks section and keep all the figures",
  ],
  abilities: ["Reads your selection or the paragraphs you name", "Shows a before/after diff", "Applies only approved edits", "Checks nothing else changed"],
  context: { required: [REQUIRE.wordOverview, target], optional: ["Surrounding paragraphs for consistency"] },
  instructions: [
    "You are the Document Editor. You improve existing text without losing its intended meaning.",
    "1. Identify exactly which text to edit: the user's selection (target \"selection\") or specific paragraphs (target \"paragraphs\" with start/end and the exact current text as expectedText). If it is unclear which passage is meant, ask.",
    "2. Edit only that text. Preserve facts, numbers, names, dates and qualifications (\"approximately\", \"subject to\") unless the user explicitly asks to change them.",
    "3. Make one replace call per contiguous passage. Do not rewrite paragraphs the user did not ask about.",
    "4. After the edit is applied, summarise what you changed in one or two sentences.",
  ].join("\n"),
  capabilities: ["selection.read", "context.read", "context.search", "document.replace_text", "document.write_blocks", "document.format_text"],
  constraints: { readOnly: false, rules: ["Edits only the selection or the paragraphs named", "Keeps numbers, names and qualifications unless told otherwise"] },
  approval: { summary: "Every edit is shown as a before/after diff for approval." },
  completion: [
    verifiedAction("edit-applied", "The requested edit was applied and verified", EDITS),
    { id: "facts-kept", description: "Numbers and names in the edited text are kept", required: true, evaluate: factsKept },
    preservesUnrelatedContent(),
  ],
  verification: ["The edited paragraphs are re-read and compared with the approved text", "All other paragraphs are compared before and after the task"],
  artifacts: ["Edited text in place"],
  limitations: ["Does not track changes (Word revision marks are not used)", "Cannot edit text inside text boxes, headers or footers"],
  routing: {
    patterns: [
      { re: /\b(edit|rewrite|revise|improve|tighten|shorten|simplify|proofread|polish|rephrase|reword|clean up)\b/i, weight: 3 },
      { re: /\bfix\b[^.?!]*\b(grammar|typos?|spelling|punctuation|wording)\b/i, weight: 3 },
      { re: /\bmake\b[^.?!]*\b(more |less )?(concise|clear|clearer|formal|friendly|shorter|readable|professional|persuasive)\b/i, weight: 3 },
      { re: /\b(selected|selection|this paragraph|paragraph \d+|this section|tone)\b/i, weight: 1 },
    ],
  },
  snapshot: true,
};
