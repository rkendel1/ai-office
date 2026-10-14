import type { Box, PowerPointDriver, PptLayout, PptSlide, RenderedImage, ShapeKind, TextFormat } from "@loai/core";
import { supports } from "./requirements";

/** Shape types that carry a text frame. Others (pictures, charts, media) are reported as non-text. */
const TEXT_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder"]);

const GEOMETRIC: Record<ShapeKind, string> = {
  rectangle: "Rectangle",
  roundedRectangle: "RoundRectangle",
  ellipse: "Ellipse",
  triangle: "Triangle",
  rightArrow: "RightArrow",
  chevron: "Chevron",
  pentagon: "HomePlate",
};
const ALIGN: Record<NonNullable<TextFormat["align"]>, string> = { left: "Left", center: "Center", right: "Right" };

/**
 * PowerPoint driver over Office.js.
 *   1.3 slides from layouts, slide masters · 1.4 shapes, text boxes, geometry, fonts, fills
 *   1.5 selected slides · 1.8 tables.
 * Speaker notes are not exposed by Office.js, so supportsNotes() is false.
 */
export class OfficePowerPointDriver implements PowerPointDriver {
  private get v14() {
    return supports("PowerPointApi", "1.4");
  }

  async info() {
    const url = Office.context.document?.url || "";
    return { id: url || undefined, title: decodeURIComponent(url.split(/[\\/]/).pop() ?? "") || undefined };
  }

  async slides(): Promise<PptSlide[]> {
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items/id");
      await ctx.sync();
      const colls = slides.items.map((s) => {
        s.shapes.load(this.v14 ? "items/id,items/name,items/type,items/left,items/top,items/width,items/height" : "items/id,items/name,items/type");
        return s.shapes;
      });
      const layouts = supports("PowerPointApi", "1.3")
        ? slides.items.map((s) => {
            s.layout.load("name");
            return s.layout;
          })
        : [];
      await ctx.sync();

      // Text and formatting are loaded per shape so one shape without a text frame
      // (e.g. a picture placeholder) cannot break the whole read.
      const result: PptSlide[] = slides.items.map((s, si) => ({
        id: s.id,
        index: si,
        layoutName: layouts[si]?.name,
        shapes: colls[si]!.items.map((sh) => ({
          id: sh.id,
          name: sh.name,
          type: String(sh.type),
          text: null,
          box: this.v14 ? { left: sh.left, top: sh.top, width: sh.width, height: sh.height } : undefined,
        })),
      }));
      for (let si = 0; si < colls.length; si++)
        for (let k = 0; k < colls[si]!.items.length; k++) {
          const sh = colls[si]!.items[k]!;
          const out = result[si]!.shapes[k]!;
          try {
            if (String(sh.type) === "Table" && supports("PowerPointApi", "1.8")) {
              const t = sh.getTable();
              t.load("values");
              await ctx.sync();
              out.text = t.values.map((r) => r.join("\t")).join("\n");
            } else if (TEXT_TYPES.has(String(sh.type))) {
              const r = sh.textFrame.textRange;
              r.load("text");
              if (this.v14) {
                r.font.load("size,bold,italic,color,name");
                r.paragraphFormat.load("horizontalAlignment");
                if (String(sh.type) === "GeometricShape") sh.fill.load("foregroundColor,type");
              }
              await ctx.sync();
              out.text = r.text;
              if (this.v14) {
                const a = String(r.paragraphFormat.horizontalAlignment ?? "").toLowerCase();
                out.format = {
                  size: r.font.size ?? undefined,
                  bold: r.font.bold ?? undefined,
                  italic: r.font.italic ?? undefined,
                  color: r.font.color ?? undefined,
                  font: r.font.name ?? undefined,
                  align: a === "center" || a === "right" || a === "left" ? a : undefined,
                };
                if (String(sh.type) === "GeometricShape" && String(sh.fill.type) === "Solid") out.fill = sh.fill.foregroundColor;
              }
            }
          } catch {
            out.text = null; // no readable text frame
          }
        }
      return result;
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

  async layouts(): Promise<PptLayout[]> {
    if (!supports("PowerPointApi", "1.3")) return [];
    return PowerPoint.run(async (ctx) => {
      const masters = ctx.presentation.slideMasters;
      masters.load("items/id");
      await ctx.sync();
      const ls = masters.items.map((m) => {
        m.layouts.load("items/id,items/name");
        return { m, l: m.layouts };
      });
      await ctx.sync();
      return ls.flatMap(({ m, l }) => l.items.map((x) => ({ id: x.id, name: x.name, masterId: m.id })));
    });
  }

  async addSlide(layout?: { layoutId: string; masterId: string }) {
    return PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.add(layout ? { layoutId: layout.layoutId, slideMasterId: layout.masterId } : undefined);
      await ctx.sync();
      const slides = ctx.presentation.slides;
      slides.load("items/id");
      await ctx.sync();
      return slides.items[slides.items.length - 1]!.id;
    });
  }

  private applyFormat(r: PowerPoint.TextRange, f?: TextFormat) {
    if (!f) return;
    if (f.size !== undefined) r.font.size = f.size;
    if (f.bold !== undefined) r.font.bold = f.bold;
    if (f.italic !== undefined) r.font.italic = f.italic;
    if (f.color !== undefined) r.font.color = f.color;
    if (f.align !== undefined) r.paragraphFormat.horizontalAlignment = ALIGN[f.align] as PowerPoint.ParagraphHorizontalAlignment;
  }

  async addTextBox(slideId: string, text: string, box: Box, format?: TextFormat) {
    return PowerPoint.run(async (ctx) => {
      const sh = ctx.presentation.slides.getItem(slideId).shapes.addTextBox(text, box);
      this.applyFormat(sh.textFrame.textRange, format);
      sh.load("id");
      await ctx.sync();
      return sh.id;
    });
  }

  async addShape(slideId: string, kind: ShapeKind, box: Box, opts: { fill?: string; line?: string; text?: string; format?: TextFormat }) {
    return PowerPoint.run(async (ctx) => {
      const sh = ctx.presentation.slides.getItem(slideId).shapes.addGeometricShape(GEOMETRIC[kind] as PowerPoint.GeometricShapeType, box);
      if (opts.fill) sh.fill.setSolidColor(opts.fill);
      if (opts.line) sh.lineFormat.color = opts.line;
      if (opts.text !== undefined) sh.textFrame.textRange.text = opts.text;
      this.applyFormat(sh.textFrame.textRange, opts.format);
      sh.load("id");
      await ctx.sync();
      return sh.id;
    });
  }

  async formatText(slideId: string, shapeId: string, format: TextFormat) {
    await PowerPoint.run(async (ctx) => {
      this.applyFormat(ctx.presentation.slides.getItem(slideId).shapes.getItem(shapeId).textFrame.textRange, format);
      await ctx.sync();
    });
  }

  async setGeometry(slideId: string, shapeId: string, box: Partial<Box>) {
    await PowerPoint.run(async (ctx) => {
      const sh = ctx.presentation.slides.getItem(slideId).shapes.getItem(shapeId);
      if (box.left !== undefined) sh.left = box.left;
      if (box.top !== undefined) sh.top = box.top;
      if (box.width !== undefined) sh.width = box.width;
      if (box.height !== undefined) sh.height = box.height;
      await ctx.sync();
    });
  }

  async deleteShape(slideId: string, shapeId: string) {
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.getItem(slideId).shapes.getItem(shapeId).delete();
      await ctx.sync();
    });
  }

  async deleteSlide(slideId: string) {
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.getItem(slideId).delete();
      await ctx.sync();
    });
  }

  async slideSize() {
    // Office.js does not report slide size; 16:9 widescreen is assumed and reported as such.
    return { width: 960, height: 540, reported: false };
  }

  supportsTables() {
    return supports("PowerPointApi", "1.8");
  }

  async addTable(slideId: string, rows: string[][], box: Box) {
    return PowerPoint.run(async (ctx) => {
      const sh = ctx.presentation.slides.getItem(slideId).shapes.addTable(rows.length, rows[0]!.length, { values: rows, left: box.left, top: box.top, width: box.width });
      sh.load("id");
      await ctx.sync();
      return sh.id;
    });
  }

  supportsNotes() {
    return false;
  }

  supportsRendering() {
    return supports("PowerPointApi", "1.8") && typeof createImageBitmap === "function";
  }

  /** Render a slide with Slide.getImageAsBase64 (PowerPointApi 1.8) and decode the PNG to RGBA pixels. */
  async renderSlide(slideId: string, widthPx: number): Promise<RenderedImage> {
    const b64 = await PowerPoint.run(async (ctx) => {
      const img = ctx.presentation.slides.getItem(slideId).getImageAsBase64({ width: widthPx });
      await ctx.sync();
      return img.value;
    });
    const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    const bitmap = await createImageBitmap(new Blob([bytes], { type: "image/png" }));
    const canvas = typeof OffscreenCanvas !== "undefined" ? new OffscreenCanvas(bitmap.width, bitmap.height) : Object.assign(document.createElement("canvas"), { width: bitmap.width, height: bitmap.height });
    const g = canvas.getContext("2d") as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
    if (!g) throw new Error("Canvas is not available to decode the rendered slide.");
    g.drawImage(bitmap, 0, 0);
    const data = g.getImageData(0, 0, bitmap.width, bitmap.height);
    return { width: data.width, height: data.height, data: data.data };
  }

  async getNotes() {
    return null;
  }

  async setNotes(): Promise<void> {
    throw new Error("Speaker notes are not available through Office.js.");
  }
}

