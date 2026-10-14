import { isApprovedPlan } from "../shared/planning.js";
import { preservesUnrelatedContent, verifiedAction } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist } from "../types.js";

const WRITES = ["document.write_blocks", "document.insert_text", "document.replace_text"];

export const documentWriter: Specialist = {
  id: "word.writer",
  name: "Document Writer",
  host: "word",
  categories: ["create", "draft", "expand", "restructure"],
  description: "Drafts substantial, structured documents from your instructions, notes or the open document — outline first, then sections.",
  examples: [
    "Turn this rough draft into an executive-ready strategy document",
    "Create a product requirements document from the meeting notes in this document",
    "Expand this outline into a complete report for the leadership team",
    "Write a one-page project proposal at the end of this document",
  ],
  abilities: ["Reads the open document", "Proposes an outline for approval", "Writes headed sections for approval", "Re-reads what it wrote"],
  context: { required: [REQUIRE.wordOverview, REQUIRE.wordDocument], optional: ["Your selection, if the task is about it"] },
  instructions: [
    "You are the Document Writer. You create substantial, coherent documents.",
    "Workflow:",
    "1. Read the source material you need (it may already be in the request context). Identify audience, purpose, tone and length from the request; if a material choice is missing and cannot be inferred, ask one concise question instead of guessing.",
    "2. Call plan.propose with the document outline: one step per section (heading + one-line summary of its content).",
    "3. After the plan is approved, write the document section by section with document.write_blocks. Use Title for the document title, Heading1/Heading2 for sections, Normal for body text, ListParagraph for list items. One section (or a few short ones) per call.",
    "4. When turning an existing draft into a new document, replace the draft paragraphs with mode \"replace\" and the exact current text as expectedText; otherwise insert at the end.",
    "5. Keep facts, names and figures from the source; never invent data. Mark assumptions explicitly in the text.",
    "Write complete prose: transitions between sections and a conclusion where the document type calls for one.",
  ].join("\n"),
  capabilities: ["selection.read", "context.read", "context.search", "plan.propose", "document.write_blocks", "document.insert_text", "document.replace_text", "document.format_text"],
  constraints: {
    requirePlan: (r) => !/\b(one|a single|short|brief)\s+(sentence|line|paragraph)\b/i.test(r),
    readOnly: false,
    rules: ["An outline must be approved before writing substantial content", "Writes go to the end of the document unless replacing a draft the user pointed to"],
  },
  approval: { summary: "You approve the outline, then each section before it is written." },
  completion: [
    {
      id: "outline-approved",
      description: "An outline was approved before writing",
      required: true,
      evaluate: (e) => {
        const plan = e.actions.findIndex(isApprovedPlan);
        const firstWrite = e.actions.findIndex((a) => WRITES.includes(a.capability) && a.status !== "rejected");
        if (firstWrite < 0) return { met: plan >= 0, detail: plan >= 0 ? "outline approved; nothing written yet" : "no outline" };
        return { met: plan >= 0 && plan < firstWrite, detail: plan >= 0 ? "outline approved first" : "content was written without an approved outline" };
      },
    },
    verifiedAction("content-written", "Document content was written and verified by read-back", WRITES),
    {
      id: "structured",
      description: "The written content has a title or headings and body paragraphs",
      required: true,
      evaluate: (e) => {
        let headings = 0;
        let paras = 0;
        for (const a of e.actions) {
          if (a.status !== "succeeded") continue;
          const args = a.args as Record<string, any>;
          if (a.capability === "document.write_blocks") {
            for (const b of args.blocks as { style?: string }[]) /Heading|Title/.test(b.style ?? "") ? headings++ : paras++;
          } else if (a.capability === "document.insert_text") {
            const n = String(args.text).split("\n").length;
            /Heading|Title/.test(args.style ?? "") ? (headings += n) : (paras += n);
          }
        }
        return { met: headings >= 1 && paras >= 2, detail: `${headings} heading(s), ${paras} body paragraph(s) written` };
      },
    },
    preservesUnrelatedContent(),
  ],
  verification: ["Each written paragraph is re-read and compared with the requested text and style", "Paragraphs outside the written range are compared before and after"],
  artifacts: ["Headed document content in the open Word document", "An approved outline"],
  limitations: ["Cannot insert images, tables of contents, footnotes or citations", "Works on the first 200 paragraphs at a time for very long documents"],
  routing: {
    patterns: [
      { re: /\b(write|draft|create|compose|generate|produce|prepare)\b[^.?!]*\b(document|memo|report|proposal|prd|requirements|brief|plan|letter|article|strategy|white ?paper|policy|section|chapter)\b/i, weight: 3 },
      { re: /\b(turn|convert|transform|rework)\b[^.?!]*\b(draft|notes|outline|bullets)\b[^.?!]*\binto\b/i, weight: 4 },
      { re: /\bexpand\b[^.?!]*\b(outline|notes|bullets)\b/i, weight: 4 },
      { re: /\b(executive[- ]ready|polished|from scratch|complete report)\b/i, weight: 1 },
    ],
  },
  snapshot: true,
};
