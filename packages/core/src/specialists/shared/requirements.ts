import type { ContextItem } from "../../types.js";
import type { ContextRequirement } from "../types.js";

const has = (pred: (c: ContextItem) => boolean) => (items: ContextItem[]) => items.some(pred);

/** Common context requirements. Each is acquired through the registry when missing. */
export const REQUIRE = {
  wordDocument: {
    id: "document-text",
    description: "The document text (first 200 paragraphs; more is read on request)",
    satisfiedBy: has((c) => c.id.startsWith("word:paragraphs:")),
    acquire: { capability: "context.read", arguments: { start: 0, count: 200 } },
  },
  wordOverview: {
    id: "document-overview",
    description: "Document overview and your selection, if any",
    satisfiedBy: has((c) => c.id === "word:meta"),
  },
  workbookOverview: {
    id: "workbook-overview",
    description: "Sheet names, used ranges and tables",
    satisfiedBy: has((c) => c.id === "excel:meta"),
    acquire: { capability: "context.read", arguments: {} },
  },
  deck: {
    id: "deck",
    description: "All slide text, shapes and positions",
    satisfiedBy: has((c) => c.id === "ppt:slides:all"),
    acquire: { capability: "presentation.read_slides", arguments: {} },
  },
  deckOverview: {
    id: "deck-overview",
    description: "Slide count and selected slides",
    satisfiedBy: has((c) => c.id === "ppt:meta"),
  },
  thread: {
    id: "thread",
    description: "The open email and its quoted thread, split by sender",
    satisfiedBy: has((c) => c.id === "outlook:thread"),
    acquire: { capability: "outlook.split_thread", arguments: {} },
  },
} satisfies Record<string, ContextRequirement>;
