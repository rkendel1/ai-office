import { disclosesPartialRead, inspected, noModifications, numbersGrounded } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist } from "../types.js";

export const documentSummarizer: Specialist = {
  id: "word.summarizer",
  name: "Document Summarizer",
  host: "word",
  categories: ["summary", "brief", "key points", "decisions", "risks", "action items"],
  description: "Produces audience-specific summaries — executive summary, brief, key findings, decisions, risks or action items — from what it actually read.",
  examples: [
    "Write an executive summary of this document for the CFO",
    "List the decisions, risks and action items in this document",
    "Give me a five-bullet brief of the key findings",
  ],
  abilities: ["Reads the document", "Summarises for the audience you name", "Says when only part was read", "Does not change the document"],
  context: { required: [REQUIRE.wordOverview, REQUIRE.wordDocument], optional: ["Your selection, to summarise just that part"] },
  instructions: [
    "You are the Document Summarizer. Summaries must reflect only inspected content.",
    "If the user selected text and the request is about the selection, summarise the selection; otherwise summarise the document. For long documents, continue reading with context.read until you have covered it or state which part you covered.",
    "Tailor the form to the request: executive summary (short paragraphs, outcome first), brief (bullets), key findings, decisions, risks, or action items (owner and date when stated).",
    "Reference paragraphs (¶n) for important points. Use numbers exactly as written in the document.",
    "If only part of the document was inspected, begin with: \"Based on paragraphs a–b of n:\".",
  ].join("\n"),
  capabilities: ["selection.read", "context.read", "context.search"],
  constraints: { readOnly: true, rules: ["Read-only", "No figures that are not in the document"] },
  approval: { summary: "No approvals needed: this specialist only reads." },
  completion: [
    inspected("read-content", "The document or selection was read", (c) => c.id.startsWith("word:paragraphs:") || c.id === "word:selection"),
    disclosesPartialRead(),
    numbersGrounded(),
    noModifications(),
  ],
  verification: ["Every number in the summary is checked against the inspected text", "Partial reads are disclosed"],
  artifacts: ["Summary in chat (ask the Document Writer to insert it into the document)"],
  limitations: ["Does not see images, charts or embedded objects"],
  routing: {
    patterns: [
      { re: /\b(summari[sz]e|summary|summaries|tl;?dr|recap|brief me|overview of)\b/i, weight: 3 },
      { re: /\b(key (points|findings|takeaways)|main points|decisions|risks|action items)\b/i, weight: 2 },
      { re: /\b(executive summary|write (a |an )?summary)\b/i, weight: 2 },
    ],
  },
};
