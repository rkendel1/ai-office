import type { Freshness } from "../context.js";
import { CapabilityUnsupportedError, ScopeError, type CapabilityDefinition } from "../registry.js";
import type { Compensation, ContextItem, ItemIdentity, Json, RiskClass, UnsupportedContent } from "../types.js";
import { fingerprint, normalizeText } from "../util.js";
import { allSupported, type ApiSupportEntry, type HostAdapter, type RequirementChecker, type StateSnapshot } from "./adapter.js";
import { excerpt, makeItem } from "./common.js";

// ---------------------------------------------------------------------------
// Driver: the thin Office.js boundary. Implemented in the add-in with Word.run;
// implemented in memory for tests.
// ---------------------------------------------------------------------------

export type WordStyle = "Normal" | "Heading1" | "Heading2" | "Heading3" | "Title" | "Quote" | "ListParagraph";
export const WORD_STYLES: WordStyle[] = ["Normal", "Heading1", "Heading2", "Heading3", "Title", "Quote", "ListParagraph"];

export interface WordParagraph {
  text: string;
  /** Built-in style name if known, else the style name. */
  style?: string;
  isListItem?: boolean;
  inTable?: boolean;
  bold?: boolean | null; // null = mixed
  italic?: boolean | null;
}

export interface WordSelection {
  text: string;
  /** Paragraph index range the selection spans, when determinable. */
  paragraphStart?: number;
  paragraphEnd?: number;
}

export interface WordDriver {
  documentInfo(): Promise<{ id?: string; title?: string }>;
  selection(): Promise<WordSelection>;
  paragraphCount(): Promise<number>;
  /** [start, end) */
  readParagraphs(start: number, end: number): Promise<WordParagraph[]>;
  /** Content present in the document that cannot be read as text. */
  unsupportedContent(): Promise<UnsupportedContent[]>;
  search(query: string, matchCase: boolean, max: number): Promise<{ paragraphIndex: number; text: string }[]>;
  /** Replace paragraphs [start, end] (inclusive) with text; "\n" separates new paragraphs. */
  replaceParagraphs(start: number, end: number, text: string): Promise<void>;
  /** Replace the current selection. Returns the text of the inserted range, read back. */
  replaceSelection(text: string): Promise<string>;
  /** Insert paragraphs; `index` = position of the first new paragraph. Returns that index. */
  insertParagraphs(at: "start" | "end" | { before: number } | { after: number }, lines: string[], style?: WordStyle): Promise<number>;
  deleteParagraphs(start: number, end: number): Promise<void>;
  formatParagraphs(start: number, end: number, fmt: { style?: WordStyle; bold?: boolean; italic?: boolean }): Promise<void>;
}

const MAX_READ = 200;
const SUBSTANTIAL_CHARS = 1500;

function renderParagraphs(start: number, paras: WordParagraph[]): string {
  return paras
    .map((p, i) => {
      const tags: string[] = [];
      if (p.style && p.style !== "Normal") tags.push(p.style);
      if (p.isListItem) tags.push("list");
      if (p.inTable) tags.push("table");
      return `[${start + i}]${tags.length ? ` (${tags.join(", ")})` : ""} ${p.text}`;
    })
    .join("\n");
}

function lines(text: string): string[] {
  return text.replace(/\r\n?/g, "\n").split("\n");
}

export class WordAdapter implements HostAdapter {
  readonly host = "word" as const;

  constructor(
    private readonly d: WordDriver,
    private readonly req: RequirementChecker = allSupported,
  ) {}

  async identity(): Promise<ItemIdentity> {
    const info = await this.d.documentInfo();
    return { host: "word", id: info.id, title: info.title };
  }

  apiSupport(): ApiSupportEntry[] {
    return [
      { requirement: "WordApi 1.1", supported: this.req.isSetSupported("WordApi", "1.1"), capabilities: ["all"] },
      {
        requirement: "WordApi 1.3",
        supported: this.req.isSetSupported("WordApi", "1.3"),
        capabilities: ["document.insert_text (positioned)", "table detection", "list detection"],
      },
    ];
  }

  private async selectionItem(): Promise<ContextItem | undefined> {
    const sel = await this.d.selection();
    if (!sel.text.trim()) return undefined;
    const where =
      sel.paragraphStart !== undefined
        ? ` (paragraphs ${sel.paragraphStart}${sel.paragraphEnd !== sel.paragraphStart ? `–${sel.paragraphEnd}` : ""})`
        : "";
    return makeItem({
      host: "word",
      id: "word:selection",
      label: `Selection${where}`,
      scope: "selection",
      content: sel.text,
      source: "selection.read",
    });
  }

  private async readItem(start: number, count: number): Promise<{ item: ContextItem; total: number; end: number }> {
    const total = await this.d.paragraphCount();
    if (total === 0) {
      const unsupported = await this.d.unsupportedContent();
      return {
        total,
        end: 0,
        item: makeItem({
          host: "word",
          id: "word:paragraphs:empty",
          label: "Document",
          scope: "full",
          content: unsupported.length
            ? "(No readable text. The document contains content that cannot be read through supported APIs.)"
            : "(The document is empty.)",
          unsupported,
          source: "context.read",
        }),
      };
    }
    if (start >= total) throw new ScopeError(`Paragraph ${start} does not exist; the document has ${total} paragraphs (0–${total - 1}).`);
    const end = Math.min(total, start + Math.min(count, MAX_READ));
    const paras = await this.d.readParagraphs(start, end);
    const unsupported = await this.d.unsupportedContent();
    const full = start === 0 && end === total;
    return {
      total,
      end,
      item: makeItem({
        host: "word",
        id: `word:paragraphs:${start}-${end - 1}`,
        label: full ? `Document (all ${total} paragraphs)` : `Paragraphs ${start}–${end - 1} of ${total}`,
        scope: full ? "full" : "partial",
        content: renderParagraphs(start, paras),
        truncated: !full,
        unsupported,
        source: "context.read",
      }),
    };
  }

  async initialContext(): Promise<ContextItem[]> {
    const items: ContextItem[] = [];
    const info = await this.d.documentInfo();
    const total = await this.d.paragraphCount();
    const unsupported = await this.d.unsupportedContent();
    items.push(
      makeItem({
        host: "word",
        id: "word:meta",
        label: "Document overview",
        scope: "metadata",
        content: `Title: ${info.title ?? "(untitled)"}\nParagraphs: ${total}${
          unsupported.length ? `\nNot readable as text: ${unsupported.map((u) => `${u.count} ${u.kind}`).join(", ")}` : ""
        }`,
        unsupported,
        source: "initial",
      }),
    );
    const sel = await this.selectionItem();
    if (sel) items.push(sel);
    return items;
  }

  async snapshot(): Promise<StateSnapshot> {
    const total = await this.d.paragraphCount();
    const paras = total ? await this.d.readParagraphs(0, total) : [];
    return { kind: "sequence", units: paras.map((p) => fingerprint(`${p.style ?? ""}|${p.text}`)) };
  }

  async checkFreshness(item: ContextItem): Promise<Freshness> {
    if (item.id === "word:selection") {
      const s = await this.d.selection();
      return fingerprint(s.text) === item.fingerprint ? "fresh" : "stale";
    }
    const m = /^word:paragraphs:(\d+)-(\d+)$/.exec(item.id);
    if (m) {
      const start = Number(m[1]);
      const end = Number(m[2]) + 1;
      const total = await this.d.paragraphCount();
      if (end > total) return "stale";
      const paras = await this.d.readParagraphs(start, end);
      return fingerprint(renderParagraphs(start, paras)) === item.fingerprint ? "fresh" : "stale";
    }
    return "unknown";
  }

  /** Current text of paragraphs [start, end] inclusive, joined by "\n". */
  private async textOf(start: number, end: number): Promise<string> {
    const paras = await this.d.readParagraphs(start, end + 1);
    return paras.map((p) => p.text).join("\n");
  }

  private async assertRange(start: number, end: number): Promise<number> {
    const total = await this.d.paragraphCount();
    if (start < 0 || end < start || end >= total)
      throw new ScopeError(`Paragraphs ${start}–${end} are outside the document (0–${total - 1}).`);
    if (end - start + 1 > MAX_READ) throw new ScopeError(`At most ${MAX_READ} paragraphs can be changed in one action.`);
    return total;
  }

  capabilities(): CapabilityDefinition<any>[] {
    const d = this.d;
    const self = this;
    const caps: CapabilityDefinition<any>[] = [];

    caps.push({
      name: "selection.read",
      description: "Read the text the user has currently selected in the document.",
      hosts: ["word"],
      risk: "read_selection",
      modifies: false,
      inputSchema: { type: "object", properties: {} },
      async execute() {
        const item = await self.selectionItem();
        if (!item) return { output: { empty: true, note: "Nothing is selected." } };
        return { output: { item: item.id, chars: item.content.length }, context: [item] };
      },
    });

    caps.push({
      name: "context.read",
      description: `Read document paragraphs (text with zero-based paragraph indices). Reads up to ${MAX_READ} paragraphs per call; call again with a later start to continue.`,
      hosts: ["word"],
      risk: "read_active",
      modifies: false,
      inputSchema: {
        type: "object",
        properties: {
          start: { type: "integer", minimum: 0, description: "First paragraph index (default 0)" },
          count: { type: "integer", minimum: 1, maximum: MAX_READ, description: "Number of paragraphs (default 80)" },
        },
      },
      resolveTarget: (a: { start?: number; count?: number }) => `Read paragraphs from ${a.start ?? 0}`,
      async execute(a: { start?: number; count?: number }) {
        const { item, total, end } = await self.readItem(a.start ?? 0, a.count ?? 80);
        return {
          output: { item: item.id, label: item.label, totalParagraphs: total, nextStart: end < total ? end : null },
          context: [item],
        };
      },
    });

    caps.push({
      name: "context.search",
      description: "Find paragraphs containing a phrase. Returns paragraph indices and text.",
      hosts: ["word"],
      risk: "read_active",
      modifies: false,
      inputSchema: {
        type: "object",
        required: ["query"],
        properties: {
          query: { type: "string", minLength: 1, maxLength: 255 },
          matchCase: { type: "boolean" },
          maxResults: { type: "integer", minimum: 1, maximum: 50 },
        },
      },
      async execute(a: { query: string; matchCase?: boolean; maxResults?: number }) {
        const hits = await d.search(a.query, !!a.matchCase, a.maxResults ?? 20);
        const item = makeItem({
          host: "word",
          id: `word:search:${fingerprint(a.query)}`,
          label: `Search “${a.query}” (${hits.length} match${hits.length === 1 ? "" : "es"})`,
          scope: "partial",
          content: hits.length ? hits.map((h) => `[${h.paragraphIndex}] ${h.text}`).join("\n") : "(no matches)",
          truncated: true,
          source: "context.search",
        });
        return { output: { item: item.id, matches: hits.length }, context: [item] };
      },
    });

    type ReplaceArgs = {
      target: "selection" | "paragraphs";
      start?: number;
      end?: number;
      expectedText?: string;
      newText: string;
    };
    caps.push({
      name: "document.replace_text",
      description:
        "Replace the user's selection, or paragraphs start..end (inclusive), with new text. For paragraphs, `expectedText` must equal the current text of those paragraphs joined by newlines (from context.read, without the [n] prefixes); the edit is refused if the document changed. Use \\n in newText to create separate paragraphs.",
      hosts: ["word"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["target", "newText"],
        properties: {
          target: { type: "string", enum: ["selection", "paragraphs"] },
          start: { type: "integer", minimum: 0 },
          end: { type: "integer", minimum: 0 },
          expectedText: { type: "string", maxLength: 200_000 },
          newText: { type: "string", maxLength: 100_000 },
        },
      },
      async resolveTarget(a: ReplaceArgs) {
        if (a.target === "selection") {
          const s = await d.selection();
          if (!s.text.trim()) throw new ScopeError("Nothing is selected.");
          return "the selected text";
        }
        if (a.start === undefined) throw new ScopeError("`start` is required when target is \"paragraphs\".");
        if (a.expectedText === undefined) throw new ScopeError("`expectedText` is required when target is \"paragraphs\".");
        const end = a.end ?? a.start;
        await self.assertRange(a.start, end);
        return end === a.start ? `paragraph ${a.start}` : `paragraphs ${a.start}–${end}`;
      },
      async assessRisk(a: ReplaceArgs): Promise<RiskClass> {
        const before = a.target === "selection" ? (await d.selection()).text : await self.textOf(a.start!, a.end ?? a.start!);
        const removed = before.length - Math.min(before.length, a.newText.length);
        if (before.length >= SUBSTANTIAL_CHARS) return "destructive";
        if (before.length > 300 && removed / before.length > 0.6) return "destructive";
        return "modify";
      },
      async preview(a: ReplaceArgs) {
        const before = a.target === "selection" ? (await d.selection()).text : await self.textOf(a.start!, a.end ?? a.start!);
        const warnings: string[] = [];
        if (a.target === "paragraphs" && normalizeText(before) !== normalizeText(a.expectedText ?? ""))
          warnings.push("The document text differs from what the assistant read. This change will be refused.");
        if (before.length >= SUBSTANTIAL_CHARS) warnings.push(`Replaces ${before.length} characters of existing text.`);
        return {
          summary: a.target === "selection" ? "Replace the selected text" : "Replace paragraph text",
          target: a.target === "selection" ? "Selection" : `Paragraphs ${a.start}–${a.end ?? a.start}`,
          before: excerpt(before, 4000),
          after: excerpt(a.newText, 4000),
          warnings,
        };
      },
      async execute(a: ReplaceArgs) {
        if (a.target === "selection") {
          const before = (await d.selection()).text;
          const readBack = await d.replaceSelection(a.newText);
          const ok = normalizeText(readBack) === normalizeText(a.newText);
          return {
            output: { replacedChars: before.length, insertedChars: a.newText.length },
            verification: {
              verified: ok,
              method: "read back inserted range",
              evidence: excerpt(readBack, 200),
              mismatch: ok ? undefined : "The text in the document after the edit does not match the requested text.",
            },
          };
        }
        const start = a.start!;
        const end = a.end ?? start;
        const before = await self.textOf(start, end);
        if (normalizeText(before) !== normalizeText(a.expectedText!)) {
          throw new Error(
            "The paragraphs no longer contain the expected text (the document changed or the wrong range was given). Re-read and try again.",
          );
        }
        await d.replaceParagraphs(start, end, a.newText);
        const newLines = lines(a.newText);
        const after = await self.textOf(start, start + newLines.length - 1).catch(() => "");
        const ok = normalizeText(after) === normalizeText(newLines.join("\n"));
        return {
          output: { paragraphs: `${start}–${start + newLines.length - 1}` },
          verification: {
            verified: ok,
            method: "re-read affected paragraphs",
            evidence: excerpt(after, 200),
            mismatch: ok ? undefined : "Re-read paragraphs do not match the requested text.",
          },
          compensation: {
            description: `Restore the previous text of paragraph${end > start ? "s" : ""} ${start}${end > start ? `–${end}` : ""}`,
            capability: "document.replace_text",
            args: { target: "paragraphs", start, end: start + newLines.length - 1, expectedText: a.newText, newText: before },
          },
        };
      },
    });

    type InsertArgs = {
      position: "start" | "end" | "before_paragraph" | "after_paragraph" | "replace_selection";
      paragraphIndex?: number;
      text: string;
      style?: WordStyle;
    };
    caps.push({
      name: "document.insert_text",
      description:
        "Insert new paragraphs at the start or end of the document, or before/after a paragraph index. Use \\n to separate paragraphs. Optional style applies to all inserted paragraphs.",
      hosts: ["word"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["position", "text"],
        properties: {
          position: { type: "string", enum: ["start", "end", "before_paragraph", "after_paragraph"] },
          paragraphIndex: { type: "integer", minimum: 0 },
          text: { type: "string", minLength: 1, maxLength: 100_000 },
          style: { type: "string", enum: WORD_STYLES as unknown as Json[] },
        },
      },
      async resolveTarget(a: InsertArgs) {
        if (a.position === "before_paragraph" || a.position === "after_paragraph") {
          if (a.paragraphIndex === undefined) throw new ScopeError("`paragraphIndex` is required for this position.");
          if (!self.req.isSetSupported("WordApi", "1.3"))
            throw new CapabilityUnsupportedError("Positioned insertion requires WordApi 1.3, which this Word version lacks.");
          await self.assertRange(a.paragraphIndex, a.paragraphIndex);
          return `${a.position === "before_paragraph" ? "before" : "after"} paragraph ${a.paragraphIndex}`;
        }
        return a.position === "start" ? "start of document" : "end of document";
      },
      async preview(a: InsertArgs) {
        return {
          summary: `Insert ${lines(a.text).length} paragraph(s)${a.style ? ` styled ${a.style}` : ""}`,
          target: a.position.replace("_", " ") + (a.paragraphIndex !== undefined ? ` ${a.paragraphIndex}` : ""),
          after: excerpt(a.text, 4000),
        };
      },
      async execute(a: InsertArgs) {
        const newLines = lines(a.text);
        const at =
          a.position === "start"
            ? "start"
            : a.position === "end"
              ? "end"
              : a.position === "before_paragraph"
                ? { before: a.paragraphIndex! }
                : { after: a.paragraphIndex! };
        const first = await d.insertParagraphs(at as any, newLines, a.style);
        const last = first + newLines.length - 1;
        const after = await self.textOf(first, last).catch(() => "");
        const ok = normalizeText(after) === normalizeText(newLines.join("\n"));
        return {
          output: { insertedParagraphs: `${first}–${last}` },
          verification: {
            verified: ok,
            method: "re-read inserted paragraphs",
            evidence: excerpt(after, 200),
            mismatch: ok ? undefined : "Inserted paragraphs could not be found with the requested text.",
          },
          compensation: {
            description: `Remove the ${newLines.length} inserted paragraph(s)`,
            capability: "document.delete_paragraphs",
            args: { start: first, end: last, expectedText: newLines.join("\n") },
          },
        };
      },
    });

    // Internal: used only for user-initiated undo. Not offered to the model.
    caps.push({
      name: "document.delete_paragraphs",
      description: "Delete paragraphs (undo only).",
      hosts: ["word"],
      risk: "destructive",
      modifies: true,
      internal: true,
      inputSchema: {
        type: "object",
        required: ["start", "end", "expectedText"],
        properties: {
          start: { type: "integer", minimum: 0 },
          end: { type: "integer", minimum: 0 },
          expectedText: { type: "string" },
        },
      },
      async resolveTarget(a: { start: number; end: number }) {
        await self.assertRange(a.start, a.end);
        return `paragraphs ${a.start}–${a.end}`;
      },
      async preview(a: { start: number; end: number }) {
        return { summary: "Delete paragraphs", target: `Paragraphs ${a.start}–${a.end}`, before: excerpt(await self.textOf(a.start, a.end)) };
      },
      async execute(a: { start: number; end: number; expectedText: string }) {
        const before = await self.textOf(a.start, a.end);
        if (normalizeText(before) !== normalizeText(a.expectedText))
          throw new Error("Those paragraphs changed since they were inserted; not deleting.");
        const countBefore = await d.paragraphCount();
        await d.deleteParagraphs(a.start, a.end);
        const countAfter = await d.paragraphCount();
        const ok = countBefore - countAfter === a.end - a.start + 1;
        return {
          output: { deleted: a.end - a.start + 1 },
          verification: { verified: ok, method: "paragraph count", evidence: `${countBefore} → ${countAfter}` },
        };
      },
    });

    type FormatArgs = { start: number; end?: number; style?: WordStyle; bold?: boolean; italic?: boolean };
    caps.push({
      name: "document.format_text",
      description: "Apply a built-in paragraph style and/or bold/italic to paragraphs start..end (inclusive).",
      hosts: ["word"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["start"],
        properties: {
          start: { type: "integer", minimum: 0 },
          end: { type: "integer", minimum: 0 },
          style: { type: "string", enum: WORD_STYLES as unknown as Json[] },
          bold: { type: "boolean" },
          italic: { type: "boolean" },
        },
      },
      async resolveTarget(a: FormatArgs) {
        if (a.style === undefined && a.bold === undefined && a.italic === undefined)
          throw new ScopeError("Specify at least one of style, bold, italic.");
        await self.assertRange(a.start, a.end ?? a.start);
        return `paragraphs ${a.start}–${a.end ?? a.start}`;
      },
      async preview(a: FormatArgs) {
        const fmt = [a.style && `style ${a.style}`, a.bold !== undefined && (a.bold ? "bold" : "not bold"), a.italic !== undefined && (a.italic ? "italic" : "not italic")]
          .filter(Boolean)
          .join(", ");
        return {
          summary: `Format text: ${fmt}`,
          target: `Paragraphs ${a.start}–${a.end ?? a.start}`,
          before: excerpt(await self.textOf(a.start, a.end ?? a.start), 800),
        };
      },
      async execute(a: FormatArgs) {
        const end = a.end ?? a.start;
        const prev = await d.readParagraphs(a.start, end + 1);
        await d.formatParagraphs(a.start, end, { style: a.style, bold: a.bold, italic: a.italic });
        const after = await d.readParagraphs(a.start, end + 1);
        const bad = after.filter(
          (p) =>
            (a.style !== undefined && p.style !== a.style) ||
            (a.bold !== undefined && p.bold !== a.bold) ||
            (a.italic !== undefined && p.italic !== a.italic),
        );
        const prevStyle = prev[0]?.style;
        const uniformPrev = prev.every((p) => p.style === prevStyle);
        return {
          output: { formatted: after.length },
          verification: {
            verified: bad.length === 0,
            method: "re-read paragraph formatting",
            mismatch: bad.length ? `${bad.length} paragraph(s) do not show the requested formatting.` : undefined,
          },
          compensation:
            a.style && uniformPrev && prevStyle && (WORD_STYLES as string[]).includes(prevStyle)
              ? {
                  description: `Restore style ${prevStyle}`,
                  capability: "document.format_text",
                  args: { start: a.start, end, style: prevStyle },
                }
              : undefined,
        };
      },
    });

    type Block = { text: string; style?: WordStyle };
    type BlocksArgs = {
      mode: "insert" | "replace";
      position?: "start" | "end" | "before_paragraph" | "after_paragraph";
      paragraphIndex?: number;
      start?: number;
      end?: number;
      expectedText?: string;
      blocks: Block[];
    };
    const renderBlocks = (blocks: Block[]) => blocks.map((b) => (b.style && b.style !== "Normal" ? `[${b.style}] ${b.text}` : b.text)).join("\n");
    caps.push({
      name: "document.write_blocks",
      description:
        "Write a structured section in one action: a list of paragraphs, each with its own built-in style (Title, Heading1-3, Normal, Quote, ListParagraph). mode \"insert\" adds them at a position; mode \"replace\" replaces paragraphs start..end (requires expectedText = their current text joined by newlines). Use for outlines, drafted sections and restructured documents.",
      hosts: ["word"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["mode", "blocks"],
        properties: {
          mode: { type: "string", enum: ["insert", "replace"] },
          position: { type: "string", enum: ["start", "end", "before_paragraph", "after_paragraph"] },
          paragraphIndex: { type: "integer", minimum: 0 },
          start: { type: "integer", minimum: 0 },
          end: { type: "integer", minimum: 0 },
          expectedText: { type: "string", maxLength: 200_000 },
          blocks: {
            type: "array",
            minItems: 1,
            maxItems: 300,
            items: {
              type: "object",
              required: ["text"],
              properties: {
                text: { type: "string", maxLength: 20_000, pattern: "^[^\\n\\r]*$" },
                style: { type: "string", enum: WORD_STYLES as unknown as Json[] },
              },
            },
          },
        },
      },
      async resolveTarget(a: BlocksArgs) {
        if (a.mode === "replace") {
          if (a.start === undefined || a.expectedText === undefined) throw new ScopeError("replace mode requires start and expectedText.");
          await self.assertRange(a.start, a.end ?? a.start);
          return `replace paragraphs ${a.start}–${a.end ?? a.start} with ${a.blocks.length} paragraph(s)`;
        }
        const pos = a.position ?? "end";
        if (pos === "before_paragraph" || pos === "after_paragraph") {
          if (a.paragraphIndex === undefined) throw new ScopeError("`paragraphIndex` is required for this position.");
          await self.assertRange(a.paragraphIndex, a.paragraphIndex);
          return `insert ${a.blocks.length} paragraph(s) ${pos === "before_paragraph" ? "before" : "after"} paragraph ${a.paragraphIndex}`;
        }
        return `insert ${a.blocks.length} paragraph(s) at the ${pos} of the document`;
      },
      async assessRisk(a: BlocksArgs): Promise<RiskClass> {
        if (a.mode !== "replace") return "modify";
        const before = await self.textOf(a.start!, a.end ?? a.start!);
        return before.length >= SUBSTANTIAL_CHARS ? "destructive" : "modify";
      },
      async preview(a: BlocksArgs) {
        const before = a.mode === "replace" ? await self.textOf(a.start!, a.end ?? a.start!) : undefined;
        const warnings: string[] = [];
        if (before !== undefined && normalizeText(before) !== normalizeText(a.expectedText ?? ""))
          warnings.push("The document text differs from what the assistant read. This change will be refused.");
        const headings = a.blocks.filter((b) => b.style && /Heading|Title/.test(b.style)).length;
        return {
          summary: `${a.mode === "replace" ? "Replace with" : "Insert"} ${a.blocks.length} paragraph(s)${headings ? `, ${headings} heading(s)` : ""}`,
          target: a.mode === "replace" ? `Paragraphs ${a.start}–${a.end ?? a.start}` : (a.position ?? "end").replace("_", " "),
          before: before !== undefined ? excerpt(before, 4000) : undefined,
          after: excerpt(renderBlocks(a.blocks), 4000),
          warnings,
        };
      },
      async execute(a: BlocksArgs) {
        let at: number;
        let oldParas: WordParagraph[] = [];
        if (a.mode === "replace") {
          const end = a.end ?? a.start!;
          const before = await self.textOf(a.start!, end);
          if (normalizeText(before) !== normalizeText(a.expectedText!))
            throw new Error("The paragraphs no longer contain the expected text. Re-read and try again.");
          oldParas = await d.readParagraphs(a.start!, end + 1);
          at = a.start!;
        } else {
          const count = await d.paragraphCount();
          const pos = a.position ?? "end";
          at = pos === "start" ? 0 : pos === "end" ? count : pos === "before_paragraph" ? a.paragraphIndex! : a.paragraphIndex! + 1;
        }
        // Insert style groups in order. In replace mode new text goes in before the old
        // paragraphs, which are then removed (Word always keeps at least one paragraph).
        let idx = at;
        let i = 0;
        while (i < a.blocks.length) {
          const style = a.blocks[i]!.style ?? "Normal";
          const group: string[] = [];
          while (i < a.blocks.length && (a.blocks[i]!.style ?? "Normal") === style) group.push(a.blocks[i++]!.text);
          const count = await d.paragraphCount();
          await d.insertParagraphs(idx >= count ? "end" : { before: idx }, group, style);
          idx += group.length;
        }
        const n = a.blocks.length;
        if (a.mode === "replace") await d.deleteParagraphs(at + n, at + n + oldParas.length - 1);
        const after = await d.readParagraphs(at, at + n);
        const mismatches = a.blocks.filter((b, k) => {
          const p = after[k];
          return !p || normalizeText(p.text) !== normalizeText(b.text) || (b.style !== undefined && p.style !== b.style);
        }).length;
        const total = await d.paragraphCount();
        return {
          output: { paragraphs: `${at}–${at + n - 1}`, written: n, headings: a.blocks.filter((b) => b.style && /Heading|Title/.test(b.style)).length, documentParagraphs: total },
          verification: {
            verified: mismatches === 0,
            method: "re-read written paragraphs (text and style)",
            evidence: excerpt(after.map((p) => p.text).join("\n"), 200),
            mismatch: mismatches ? `${mismatches} paragraph(s) do not match the requested text or style.` : undefined,
          },
          compensation: (a.mode === "replace"
              ? {
                  description: `Restore the previous ${oldParas.length} paragraph(s)`,
                  capability: "document.write_blocks",
                  args: {
                    mode: "replace",
                    start: at,
                    end: at + n - 1,
                    expectedText: a.blocks.map((b) => b.text).join("\n"),
                    blocks: oldParas.map((p) => ({ text: p.text, ...(p.style && (WORD_STYLES as string[]).includes(p.style) ? { style: p.style } : {}) })) as Json,
                  },
                }
              : {
                  description: `Remove the ${n} inserted paragraph(s)`,
                  capability: "document.delete_paragraphs",
                  args: { start: at, end: at + n - 1, expectedText: a.blocks.map((b) => b.text).join("\n") },
                }) as Compensation,
        };
      },
    });

    return caps;
  }
}
