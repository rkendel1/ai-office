import type { PowerPointDriver, PptSlide } from "@loai/core";
import { supports } from "./requirements";

/** Shape types that carry a text frame. Others (pictures, charts, media) are reported as non-text. */
const TEXT_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder"]);

/** PowerPoint driver over Office.js (PowerPointApi 1.3–1.5). */
export class OfficePowerPointDriver implements PowerPointDriver {
  async info() {
    const url = Office.context.document?.url || "";
    return { id: url || undefined, title: decodeURIComponent(url.split(/[\\/]/).pop() ?? "") || undefined };
  }

  async slides(): Promise<PptSlide[]> {
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items/id");
      await ctx.sync();
      const shapeColls = slides.items.map((s) => {
        s.shapes.load("items/id,items/name,items/type");
        return s.shapes;
      });
      await ctx.sync();
      const texts: { slide: number; shape: number; range: PowerPoint.TextRange }[] = [];
      shapeColls.forEach((coll, si) =>
        coll.items.forEach((sh, shi) => {
          if (TEXT_TYPES.has(String(sh.type))) {
            const r = sh.textFrame.textRange;
            r.load("text");
            texts.push({ slide: si, shape: shi, range: r });
          }
        }),
      );
      await ctx.sync().catch(() => undefined); // a placeholder without text frame must not break the whole read
      const textOf = new Map<string, string>();
      for (const t of texts) {
        try {
          textOf.set(`${t.slide}/${t.shape}`, t.range.text);
        } catch {
          /* not loaded → treated as non-text */
        }
      }
      return slides.items.map((s, si) => ({
        id: s.id,
        index: si,
        shapes: shapeColls[si]!.items.map((sh, shi) => ({
          id: sh.id,
          name: sh.name,
          type: String(sh.type),
          text: textOf.has(`${si}/${shi}`) ? textOf.get(`${si}/${shi}`)! : null,
        })),
      }));
    });
  }

  async selectedSlideIds() {
    if (!supports("PowerPointApi", "1.5")) return null;
    return PowerPoint.run(async (ctx) => {
      const sel = ctx.presentation.getSelectedSlides();
      sel.load("items/id");
      await ctx.sync();
      return sel.items.map((s) => s.id);
    });
  }

  async setShapeText(slideId: string, shapeId: string, text: string) {
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.getItem(slideId).shapes.getItem(shapeId).textFrame.textRange.text = text;
      await ctx.sync();
    });
  }

  async addSlide() {
    return PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.add();
      await ctx.sync();
      const slides = ctx.presentation.slides;
      slides.load("items/id");
      await ctx.sync();
      return slides.items[slides.items.length - 1]!.id;
    });
  }

  async addTextBox(slideId: string, text: string, box: { left: number; top: number; width: number; height: number }) {
    return PowerPoint.run(async (ctx) => {
      const sh = ctx.presentation.slides.getItem(slideId).shapes.addTextBox(text, box);
      sh.load("id");
      await ctx.sync();
      return sh.id;
    });
  }

  async deleteSlide(slideId: string) {
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.getItem(slideId).delete();
      await ctx.sync();
    });
  }
}
