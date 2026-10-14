import { inspected, noModifications, structuredItems } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist } from "../types.js";
import { ITEM_FORMAT } from "./thread-analyst.js";

export const actionExtractor: Specialist = {
  id: "outlook.actions",
  name: "Action Extractor",
  host: "outlook",
  categories: ["tasks", "owners", "deadlines", "dependencies"],
  description: "Extracts tasks, owners, deadlines, dependencies and open questions from the thread as a structured list you can copy. Creates nothing outside this chat.",
  examples: [
    "Extract the action items from this thread with owners and deadlines",
    "What do I need to do after this email?",
    "List the next steps and who is responsible",
  ],
  abilities: ["Splits the thread by sender", "Separates explicit assignments from inferred tasks", "Links each item to its message", "Does not create tasks anywhere"],
  context: { required: [REQUIRE.thread], optional: [] },
  instructions: [
    "You are the Action Extractor.",
    "List every task in the thread, one per line:",
    ITEM_FORMAT,
    "An item is explicit only if the cited message assigns or accepts it in so many words (\"Bob, please send…\", \"I'll book…\"). The owner of an explicit item must be the author of that message or be named in it.",
    "Anything you deduce (a natural follow-up nobody asked for) is inferred.",
    "After the list, add \"Dependencies:\" and \"Open questions:\" sections (or \"none\").",
    "Do not create tasks, calendar entries or messages.",
  ].join("\n"),
  capabilities: ["outlook.split_thread"],
  constraints: { readOnly: true, rules: ["Read-only: no tasks, events or messages are created"] },
  approval: { summary: "No approvals needed: this specialist only reads." },
  completion: [
    inspected("thread-split", "The thread was split into attributed messages", (c) => c.id === "outlook:thread"),
    structuredItems("items", "Each task is marked explicit or inferred; explicit ones are traced to their message and owner", "assignment", 1),
    noModifications(),
  ],
  verification: ["Explicit items are checked against the cited message (wording and owner)"],
  artifacts: ["Structured task list in chat, ready to copy"],
  limitations: ["Does not create tasks in To Do, Planner or any other system", "Sees only the open message and its quoted history"],
  routing: {
    patterns: [
      { re: /\b(action items?|to-?dos?|next steps|tasks?|deliverables)\b/i, weight: 3 },
      { re: /\b(owners?|deadlines?|due dates?|who (needs|has|is supposed) to|responsib\w*|extract)\b/i, weight: 2 },
      { re: /\bwhat do i (need|have) to do\b/i, weight: 3 },
    ],
  },
};
