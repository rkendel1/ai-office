import type { UnsupportedContent, WordDriver, WordParagraph, WordStyle } from "@loai/core";
import { supports } from "./requirements";

/** Word driver over Office.js (WordApi 1.1+, enhanced with 1.3 where available). */
export class OfficeWordDriver implements WordDriver {
  private get v13() {
    return supports("WordApi", "1.3");
  }

  async documentInfo() {
    return Word.run(async (ctx) => {
      const props = ctx.document.properties;
      props.load("title");
      await ctx.sync();
      const url = Office.context.document?.url || "";
      const file = decodeURIComponent(url.split(/[\\/]/).pop() ?? "");
      return { id: url || undefined, title: props.title || file || undefined };
    });
  }

  async selection() {
    return Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load("text");
      await ctx.sync();
      return { text: sel.text ?? "" };
    });
  }

  async paragraphCount() {
    return Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items/text");
      await ctx.sync();
      return ps.items.length;
    });
  }

  async readParagraphs(start: number, end: number): Promise<WordParagraph[]> {
    return Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      const fields = ["items/text", "items/style", "items/styleBuiltIn", "items/font/bold", "items/font/italic"];
      if (this.v13) fields.push("items/isListItem", "items/tableNestingLevel");
      ps.load(fields.join(","));
      await ctx.sync();
      return ps.items.slice(start, end).map((p) => ({
        text: p.text,
        style: p.styleBuiltIn && p.styleBuiltIn !== "Other" ? String(p.styleBuiltIn) : p.style,
        isListItem: this.v13 ? p.isListItem : undefined,
        inTable: this.v13 ? p.tableNestingLevel > 0 : undefined,
        bold: p.font.bold ?? null,
        italic: p.font.italic ?? null,
      }));
    });
  }

  async unsupportedContent(): Promise<UnsupportedContent[]> {
    return Word.run(async (ctx) => {
      const pics = ctx.document.body.inlinePictures;
      pics.load("items");
      await ctx.sync();
      const out: UnsupportedContent[] = [];
      if (pics.items.length) out.push({ kind: "images", count: pics.items.length, note: "Image content is not read" });
      return out;
    });
  }

  async search(query: string, matchCase: boolean, max: number) {
    // Done over paragraph text so results carry exact paragraph indices.
    const all = await this.readParagraphs(0, Number.MAX_SAFE_INTEGER);
    const needle = matchCase ? query : query.toLowerCase();
    return all
      .map((p, i) => ({ paragraphIndex: i, text: p.text }))
      .filter((p) => (matchCase ? p.text : p.text.toLowerCase()).includes(needle))
      .slice(0, max);
  }

  async replaceParagraphs(start: number, end: number, text: string) {
    await Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items");
      await ctx.sync();
      const first = ps.items[start]!;
      if (start === end) {
        first.insertText(text, "Replace");
      } else if (this.v13) {
        const range = first.getRange("Whole").expandTo(ps.items[end]!.getRange("Content"));
        range.insertText(text, "Replace");
      } else {
        // WordApi 1.1: replace the first, delete the rest (last to first).
        for (let i = end; i > start; i--) ps.items[i]!.delete();
        first.insertText(text, "Replace");
      }
      await ctx.sync();
    });
  }

  async replaceSelection(text: string) {
    return Word.run(async (ctx) => {
      const r = ctx.document.getSelection().insertText(text, "Replace");
      r.load("text");
      await ctx.sync();
      return r.text;
    });
  }

  async insertParagraphs(at: "start" | "end" | { before: number } | { after: number }, lines: string[], style?: WordStyle) {
    return Word.run(async (ctx) => {
      const body = ctx.document.body;
      const ps = body.paragraphs;
      ps.load("items");
      await ctx.sync();
      const total = ps.items.length;
      const apply = (p: Word.Paragraph) => {
        if (style) p.styleBuiltIn = style as unknown as Word.BuiltInStyleName;
      };
      let index: number;
      if (at === "start") {
        for (const line of [...lines].reverse()) apply(body.insertParagraph(line, "Start"));
        index = 0;
      } else if (at === "end") {
        for (const line of lines) apply(body.insertParagraph(line, "End"));
        index = total;
      } else if ("before" in at) {
        const anchor = ps.items[at.before]!;
        for (const line of lines) apply(anchor.insertParagraph(line, "Before"));
        index = at.before;
      } else {
        let anchor = ps.items[at.after]!;
        for (const line of lines) {
          anchor = anchor.insertParagraph(line, "After");
          apply(anchor);
        }
        index = at.after + 1;
      }
      await ctx.sync();
      return index;
    });
  }

  async deleteParagraphs(start: number, end: number) {
    await Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items");
      await ctx.sync();
      for (let i = end; i >= start; i--) ps.items[i]!.delete();
      await ctx.sync();
    });
  }

  async formatParagraphs(start: number, end: number, fmt: { style?: WordStyle; bold?: boolean; italic?: boolean }) {
    await Word.run(async (ctx) => {
      const ps = ctx.document.body.paragraphs;
      ps.load("items");
      await ctx.sync();
      for (let i = start; i <= end; i++) {
        const p = ps.items[i]!;
        if (fmt.style) p.styleBuiltIn = fmt.style as unknown as Word.BuiltInStyleName;
        if (fmt.bold !== undefined) p.font.bold = fmt.bold;
        if (fmt.italic !== undefined) p.font.italic = fmt.italic;
      }
      await ctx.sync();
    });
  }
}
