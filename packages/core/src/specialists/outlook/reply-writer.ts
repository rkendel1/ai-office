import { inspected } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist } from "../types.js";

export const replyWriter: Specialist = {
  id: "outlook.reply",
  name: "Reply Writer",
  host: "outlook",
  categories: ["reply", "response", "draft"],
  description: "Drafts a reply that answers what the thread asks, in the tone you want. It creates a draft for you to review and send — it never sends.",
  examples: [
    "Draft a reply accepting the meeting and answering their two questions",
    "Summarize this thread, extract commitments, and draft a reply addressing each one",
    "Write a polite reply declining, and suggest next week instead",
  ],
  abilities: ["Reads the thread", "Finds the questions that need an answer", "Drafts the reply for your approval", "Never sends — you send from Outlook"],
  context: { required: [REQUIRE.thread], optional: ["The tone and points you want made"] },
  instructions: [
    "You are the Reply Writer.",
    "1. From the thread, list (for yourself) every question or request addressed to the user, and any commitments already made.",
    "2. Draft a reply in the requested tone that addresses each one. Do not invent commitments, dates, prices or facts: if the user has not said how to answer something, write a clear placeholder like [confirm date] and mention it in your final message.",
    "3. Create the draft with outlook.create_draft (plain text; use replyAll only if the user asked or the thread clearly needs it).",
    "4. In your final message: what the draft answers, any placeholders the user must fill, and whether the draft could be verified. Remind the user that nothing was sent.",
    "If the request also asks for a summary or commitments, include them before the draft summary, attributed to message ids (M1, M2, …).",
  ].join("\n"),
  capabilities: ["context.read", "outlook.read_message", "outlook.split_thread", "outlook.create_draft"],
  constraints: { readOnly: false, rules: ["Drafts only: sending is not possible", "No invented commitments; unknowns become placeholders"] },
  approval: { summary: "The draft text and recipients are shown for approval before the draft is created." },
  completion: [
    inspected("thread-read", "The thread was read", (c) => c.id === "outlook:thread" || c.id === "outlook:message"),
    {
      id: "draft-created",
      description: "A reply draft was created (verified when Outlook allows)",
      required: true,
      evaluate: (e) => {
        const d = e.actions.filter((a) => a.capability === "outlook.create_draft");
        const ok = d.find((a) => a.status === "succeeded");
        const unverified = d.find((a) => a.status === "unverified");
        if (ok) return { met: true, detail: "draft inserted and verified" };
        if (unverified) return { met: true, detail: "reply window opened; Outlook does not allow reading it back (unverified)" };
        return { met: false, detail: "no draft was created" };
      },
    },
    {
      id: "not-sent",
      description: "No message was sent",
      required: true,
      evaluate: (e) => {
        const sent = e.actions.filter((a) => /send/i.test(a.capability) && a.status !== "rejected");
        return { met: sent.length === 0, detail: sent.length ? "a send operation ran" : "nothing sent (no send capability exists)" };
      },
    },
    {
      id: "placeholders-reported",
      description: "Placeholders left in the draft are pointed out",
      required: true,
      evaluate: (e) => {
        const body = e.actions.filter((a) => a.capability === "outlook.create_draft" && a.status !== "rejected").map((a) => String((a.args as any).body)).join("\n");
        const ph = body.match(/\[[^\]\n]{2,40}\]/g) ?? [];
        const said = ph.every((p) => e.text.includes(p)) || /placeholder/i.test(e.text);
        return { met: !ph.length || said, detail: ph.length ? `${ph.length} placeholder(s)${said ? " reported" : " not mentioned"}` : "none" };
      },
    },
  ],
  verification: ["In a compose window, the inserted text is re-read", "A reply window opened from a read message cannot be read back and is reported as unverified"],
  artifacts: ["A reply draft in Outlook (not sent)"],
  limitations: ["Cannot send email", "Cannot add attachments or new recipients", "Reply windows opened from a read message cannot be verified by add-ins"],
  routing: {
    patterns: [
      { re: /\b(draft|write|compose|prepare)\b[^.?!]*\b(reply|response|answer|email back)\b/i, weight: 4 },
      { re: /\b(reply|respond|write back|get back to)\b/i, weight: 3 },
    ],
    subsumes: ["outlook.thread", "outlook.actions"],
  },
  snapshot: true,
};
