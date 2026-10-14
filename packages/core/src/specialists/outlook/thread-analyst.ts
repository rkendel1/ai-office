import { inspected, noModifications, structuredItems, textCriterion } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist } from "../types.js";

export const ITEM_FORMAT = [
  "Use exactly this line format for each item (it is checked):",
  "- [explicit] <what> — owner: <person> — due: <date or none> — source: M<n>",
  "- [inferred] <what> — owner: <person or unknown> — due: <date or none> — source: M<n or none>",
  "\"explicit\" = stated in so many words in the cited message; \"inferred\" = your interpretation.",
].join("\n");

export const threadAnalyst: Specialist = {
  id: "outlook.thread",
  name: "Thread Analyst",
  host: "outlook",
  categories: ["summary", "decisions", "commitments", "open questions"],
  description: "Summarises the email thread and lists decisions, open questions and commitments, each attributed to the person and message it came from.",
  examples: [
    "Summarize this thread and list who committed to what",
    "What was decided, and what is still open?",
    "Recap this conversation for someone who missed it",
  ],
  abilities: ["Splits the thread by sender", "Attributes statements to people and messages", "Separates explicit commitments from inferred ones", "Does not change or send anything"],
  context: { required: [REQUIRE.thread], optional: [] },
  instructions: [
    "You are the Thread Analyst. The thread is split into messages M1 (the open, newest message), M2 (the one it quotes), and so on, each with its author.",
    "Reply with these sections:",
    "Summary — 2–5 sentences, oldest to newest.",
    "Decisions — what was agreed, with the message id.",
    "Open questions — unanswered questions and who asked them.",
    "Commitments — one line per commitment, attributed to the person who made it in the cited message:",
    ITEM_FORMAT,
    "A commitment is explicit only if its author states it themselves in that message (\"I'll send…\", \"We will deliver…\").",
    "Never attribute a statement to someone who did not write it. If the thread could not be split reliably, say so.",
  ].join("\n"),
  capabilities: ["context.read", "outlook.read_message", "outlook.split_thread"],
  constraints: { readOnly: true, rules: ["Read-only", "Only the open item and its quoted thread are read"] },
  approval: { summary: "No approvals needed: this specialist only reads." },
  completion: [
    inspected("thread-split", "The thread was split into attributed messages", (c) => c.id === "outlook:thread"),
    structuredItems("attributed-commitments", "Commitments are attributed to the author of the cited message", "commitment", 0),
    textCriterion("sections", "Includes a summary and commitments section", (t) => {
      const ok = /summary/i.test(t) && /commitments?/i.test(t);
      return { met: ok, detail: ok ? "present" : "missing Summary or Commitments section" };
    }),
    noModifications(),
  ],
  verification: ["Each explicit commitment is checked against the cited message: it must exist, share wording, and be written by the named owner"],
  artifacts: ["Summary, decisions, open questions and attributed commitments in chat"],
  limitations: ["Sees only the open message and the quoted history in its body, not other messages in the mailbox", "Attachments are not read", "Unusual reply-header formats may not split; such text stays with the previous message"],
  routing: {
    patterns: [
      { re: /\b(summari[sz]e|summary|recap|catch me up|what('s| is| was) (this|the) (thread|email|conversation) about)\b/i, weight: 3 },
      { re: /\b(commit(ted|ments?)|decisions?|decided|open questions?|who (said|agreed|promised))\b/i, weight: 2 },
      { re: /\b(thread|conversation)\b/i, weight: 1 },
    ],
  },
};
