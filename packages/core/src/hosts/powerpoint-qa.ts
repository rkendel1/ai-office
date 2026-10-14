/**
 * PowerPoint quality assurance.
 *
 * Two kinds of evidence, never mixed up:
 *  - object-level: Office object geometry, text and formatting ("geometry",
 *    "text-estimate"). Says nothing about how the slide actually looks.
 *  - rendered: pixels of the slide as PowerPoint rendered it (Slide.getImageAsBase64,
 *    PowerPointApi 1.8). Only these findings may be described as visual.
 */
import type { Box, PptShape, PptSlide } from "./powerpoint.js";

export type FindingBasis = "geometry" | "text-estimate" | "rendered";
export type Severity = "high" | "medium" | "low";

export type IssueKind =
  | "off_slide"
  | "overlap"
  | "overflow"
  | "dense"
  | "no_title"
  | "empty_placeholder"
  | "empty_slide"
  | "nearly_empty"
  | "small_text"
  | "inconsistent_title_size"
  | "inconsistent_title_position"
  | "inconsistent_body_size"
  | "inconsistent_fonts"
  | "rendered_overflow"
  | "rendered_text_missing"
  | "rendered_blank";

export interface LayoutIssue {
  /** 1-based slide number; 0 = whole deck. */
  slide: number;
  slideId: string;
  kind: IssueKind;
  severity: Severity;
  basis: FindingBasis;
  /** Objects involved. */
  objects: { id: string; name: string }[];
  /** First affected object (kept for compatibility). */
  shapeId?: string;
  detail: string;
  /** What to do about it. */
  remediation: string;
}

const area = (b: Box) => Math.max(0, b.width) * Math.max(0, b.height);
export function intersection(a: Box, b: Box): number {
  const w = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left);
  const h = Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top);
  return w > 0 && h > 0 ? w * h : 0;
}

export function isTitleShape(sh: PptShape): boolean {
  return sh.text !== null && /title/i.test(sh.name) && !/subtitle/i.test(sh.name);
}

/** Lines a text needs in a box, using average glyph width ≈ 0.5em. */
export function estimateLines(text: string, width: number, size = 18): number {
  const usableWidth = Math.max(10, width - 14.4);
  const charsPerLine = Math.max(1, Math.floor(usableWidth / (size * 0.5)));
  return text.split(/\r?\n|\u000b/).reduce((n, p) => n + Math.max(1, Math.ceil(p.length / charsPerLine)), 0);
}

/** Estimate whether text fits its box (line height ≈ 1.2em, 7.2pt insets). */
export function estimateOverflow(text: string, box: Box, size = 18): { needed: number; available: number; overflows: boolean } {
  const needed = estimateLines(text, box.width, size) * size * 1.2 + 7.2;
  return { needed, available: box.height, overflows: needed > box.height * 1.15 };
}

const LAYOUT_WITHOUT_TITLE = /^(blank|picture|content only)/i;
const words = (s: string | null) => (s?.trim() ? s.trim().split(/\s+/).length : 0);
const obj = (sh: PptShape) => ({ id: sh.id, name: sh.name });

function median(xs: number[]): number {
  const s = [...xs].sort((a, b) => a - b);
  return s.length ? s[Math.floor(s.length / 2)]! : 0;
}

/** Object-level checks of slide geometry, text and formatting. */
export function checkLayout(slides: PptSlide[], size: { width: number; height: number }): LayoutIssue[] {
  const issues: LayoutIssue[] = [];
  const titles: { slide: PptSlide; sh: PptShape }[] = [];
  const bodySizes = new Map<number, number[]>();
  const fonts = new Map<string, number[]>();

  for (const s of slides) {
    const n = s.index + 1;
    const add = (kind: IssueKind, severity: Severity, basis: FindingBasis, objects: PptShape[], detail: string, remediation: string) =>
      issues.push({ slide: n, slideId: s.id, kind, severity, basis, objects: objects.map(obj), shapeId: objects[0]?.id, detail, remediation });

    const textShapes = s.shapes.filter((x) => x.text !== null);
    const visuals = s.shapes.filter((x) => x.text === null || x.type === "Table" || (x.type === "GeometricShape" && !x.text?.trim()));
    const totalWords = textShapes.reduce((c, x) => c + words(x.text), 0);

    // Empty / nearly empty
    if (!s.shapes.length || (!totalWords && !visuals.length))
      add("empty_slide", "high", "geometry", [], "Slide has no content.", "Add content or delete the slide.");
    else if (totalWords < 4 && !visuals.length && textShapes.length <= 1)
      add("nearly_empty", "medium", "geometry", textShapes, `Slide has only ${totalWords} word(s) and no visuals.`, "Add the slide's key message and supporting content, or merge it with a neighbouring slide.");

    // Titles
    const title = textShapes.find(isTitleShape) ?? null;
    const expectsTitle = s.layoutName ? !LAYOUT_WITHOUT_TITLE.test(s.layoutName) : true;
    if (!title || !title.text?.trim()) {
      if (s.shapes.length)
        add(
          "no_title",
          expectsTitle ? "medium" : "low",
          "geometry",
          title ? [title] : [],
          expectsTitle ? `Slide has no title${s.layoutName ? ` although its layout (“${s.layoutName}”) has one` : ""}.` : "Slide has no title (its layout has none).",
          "Add a short title that states the slide's message; titles also drive navigation and accessibility.",
        );
    } else titles.push({ slide: s, sh: title });

    // Density
    if (totalWords > 90) add("dense", "high", "text-estimate", textShapes, `${totalWords} words on one slide.`, "Split into two slides or cut to the key points (aim for under 50 words).");
    else if (totalWords > 70) add("dense", "medium", "text-estimate", textShapes, `${totalWords} words; dense for presenting.`, "Trim supporting detail into speaker notes.");

    for (const sh of s.shapes) {
      const b = sh.box;
      if (sh.type === "Placeholder" && sh.text === "")
        add("empty_placeholder", "medium", "geometry", [sh], `Empty placeholder “${sh.name}” shows “Click to add text” while editing.`, "Fill it or delete it.");
      if (sh.text?.trim()) {
        if (sh.format?.font) fonts.set(sh.format.font, [...(fonts.get(sh.format.font) ?? []), n]);
        if (!isTitleShape(sh) && sh.format?.size) bodySizes.set(sh.format.size, [...(bodySizes.get(sh.format.size) ?? []), n]);
      }
      if (!b) continue;
      if (b.left < -1 || b.top < -1 || b.left + b.width > size.width + 1 || b.top + b.height > size.height + 1)
        add("off_slide", "high", "geometry", [sh], `“${sh.name}” extends beyond the slide edge.`, `Move or resize it to fit within ${size.width}×${size.height}pt.`);
      if (sh.text?.trim()) {
        const fs = sh.format?.size;
        if (fs && fs < 12) add("small_text", "low", "geometry", [sh], `“${sh.name}” uses ${fs}pt text.`, "Use at least 14pt for body text that will be presented.");
        const o = estimateOverflow(sh.text, b, fs ?? 18);
        if (o.overflows)
          add(
            "overflow",
            "high",
            "text-estimate",
            [sh],
            `“${sh.name}” likely overflows its box (needs ~${Math.round(o.needed)}pt of ${Math.round(o.available)}pt height).`,
            `Shorten the text, reduce the font size, or enlarge the box to about ${Math.ceil(o.needed)}pt tall.`,
          );
      }
    }

    // Overlaps: text–text (high) and picture/visual over text (medium).
    for (let i = 0; i < s.shapes.length; i++)
      for (let j = i + 1; j < s.shapes.length; j++) {
        const a = s.shapes[i]!;
        const c = s.shapes[j]!;
        if (!a.box || !c.box) continue;
        const aText = !!a.text?.trim();
        const cText = !!c.text?.trim();
        if (!aText && !cText) continue;
        const overlap = intersection(a.box, c.box) / Math.max(1, Math.min(area(a.box), area(c.box)));
        if (overlap <= 0.2) continue;
        if (aText && cText)
          add("overlap", "high", "geometry", [a, c], `“${a.name}” and “${c.name}” overlap (${Math.round(overlap * 100)}% of the smaller box).`, "Move one of them or reduce their size so the text does not collide.");
        else if (a.text === null || c.text === null) {
          const visual = a.text === null ? a : c;
          const text = visual === a ? c : a;
          add("overlap", "medium", "geometry", [visual, text], `“${visual.name}” covers ${Math.round(overlap * 100)}% of text “${text.name}”.`, "Move the picture/object or send it behind the text, and check the result visually.");
        }
      }
  }

  // Deck-wide typography and layout consistency
  const tSizes = new Map<number, number[]>();
  for (const t of titles) if (t.sh.format?.size) tSizes.set(t.sh.format.size, [...(tSizes.get(t.sh.format.size) ?? []), t.slide.index + 1]);
  const deck = (kind: IssueKind, detail: string, remediation: string) => issues.push({ slide: 0, slideId: "*", kind, severity: "low", basis: "geometry", objects: [], detail, remediation });
  const list = (m: Map<number | string, number[]>, unit = "") => [...m.entries()].map(([k, ns]) => `${k}${unit} on slide(s) ${[...new Set(ns)].join(", ")}`).join("; ");
  if (tSizes.size > 1) deck("inconsistent_title_size", `Title sizes vary: ${list(tSizes, "pt")}.`, `Use one title size throughout (e.g. ${median([...tSizes.keys()])}pt).`);
  const placed = titles.filter((t) => t.sh.box);
  if (placed.length >= 3) {
    const ml = median(placed.map((t) => t.sh.box!.left));
    const mt = median(placed.map((t) => t.sh.box!.top));
    const off = placed.filter((t) => Math.abs(t.sh.box!.left - ml) > 12 || Math.abs(t.sh.box!.top - mt) > 12).map((t) => t.slide.index + 1);
    if (off.length) deck("inconsistent_title_position", `Titles sit in a different place on slide(s) ${off.join(", ")}.`, `Align titles to the common position (${Math.round(ml)}, ${Math.round(mt)}pt) or use the same layout.`);
  }
  if (bodySizes.size > 2) deck("inconsistent_body_size", `Body text uses ${bodySizes.size} different sizes: ${list(bodySizes, "pt")}.`, "Use one or two body sizes across the deck.");
  if (fonts.size > 2) deck("inconsistent_fonts", `${fonts.size} different fonts are used: ${list(fonts)}.`, "Use the theme fonts (one heading font, one body font).");

  const order = { high: 0, medium: 1, low: 2 };
  return issues.sort((x, y) => order[x.severity] - order[y.severity] || x.slide - y.slide);
}

// --------------------------------------------------------------- rendering

export interface RenderedImage {
  width: number;
  height: number;
  /** RGBA, row-major. */
  data: Uint8ClampedArray;
}

function background(img: RenderedImage): [number, number, number] {
  const counts = new Map<number, number>();
  const sample = (x: number, y: number) => {
    const i = (y * img.width + x) * 4;
    const key = ((img.data[i]! >> 4) << 8) | ((img.data[i + 1]! >> 4) << 4) | (img.data[i + 2]! >> 4);
    counts.set(key, (counts.get(key) ?? 0) + 1);
  };
  for (let x = 0; x < img.width; x += 2) {
    sample(x, 0);
    sample(x, img.height - 1);
  }
  for (let y = 0; y < img.height; y += 2) {
    sample(0, y);
    sample(img.width - 1, y);
  }
  const key = [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? 0xfff;
  return [((key >> 8) & 15) * 16 + 8, ((key >> 4) & 15) * 16 + 8, (key & 15) * 16 + 8];
}

/**
 * Analyse a rendered slide image. Detects text that renders below its box
 * (overflow), text shapes whose box shows no rendered content, and slides
 * that render blank although they have text. Pixels are compared with the
 * slide background; this needs a rendered image, not object geometry.
 */
export function analyzeRendering(slide: PptSlide, img: RenderedImage, size: { width: number; height: number }): LayoutIssue[] {
  const bg = background(img);
  const sx = img.width / size.width;
  const sy = img.height / size.height;
  const ink = (x: number, y: number) => {
    const i = (y * img.width + x) * 4;
    return Math.abs(img.data[i]! - bg[0]) + Math.abs(img.data[i + 1]! - bg[1]) + Math.abs(img.data[i + 2]! - bg[2]) > 90;
  };
  const toPx = (b: Box) => ({
    x0: Math.max(0, Math.floor(b.left * sx)),
    y0: Math.max(0, Math.floor(b.top * sy)),
    x1: Math.min(img.width, Math.ceil((b.left + b.width) * sx)),
    y1: Math.min(img.height, Math.ceil((b.top + b.height) * sy)),
  });
  const boxes = slide.shapes.filter((s) => s.box).map((s) => ({ s, px: toPx(s.box!) }));
  const inAnyBox = (x: number, y: number, except?: PptShape) => boxes.some((b) => b.s !== except && x >= b.px.x0 - 1 && x < b.px.x1 + 1 && y >= b.px.y0 - 1 && y < b.px.y1 + 1);
  const n = slide.index + 1;
  const issues: LayoutIssue[] = [];
  const add = (kind: IssueKind, severity: Severity, objects: PptShape[], detail: string, remediation: string) =>
    issues.push({ slide: n, slideId: slide.id, kind, severity, basis: "rendered", objects: objects.map(obj), shapeId: objects[0]?.id, detail, remediation });

  let total = 0;
  for (let y = 0; y < img.height; y += 2) for (let x = 0; x < img.width; x += 2) if (ink(x, y)) total++;
  const textShapes = slide.shapes.filter((s) => s.text?.trim());
  if (total < (img.width * img.height) / 4 / 2000 && textShapes.length) {
    add("rendered_blank", "high", textShapes, "The slide renders blank although it has text.", "Check text colour against the background and that the shapes are not hidden.");
    return issues;
  }

  for (const sh of textShapes) {
    if (!sh.box) continue;
    const px = toPx(sh.box);
    // Ink inside the box
    let inside = 0;
    for (let y = px.y0; y < px.y1; y += 2) for (let x = px.x0; x < px.x1; x += 2) if (ink(x, y)) inside++;
    if (!inside)
      add("rendered_text_missing", "medium", [sh], `Text in “${sh.name}” is not visible in the rendered slide.`, "Check the text colour, size and that the shape is not covered by another object.");
    // Ink continuing directly below the box, outside every other object = text running past
    // the bottom edge. Scanning stops at the first full line of empty rows, so content that
    // merely sits further down (another shape's overflow, a footer) is not attributed here.
    const lineH = Math.max(4, Math.round((sh.format?.size ?? 18) * 1.2 * sy));
    let spill = 0;
    let band = 0;
    let emptyRun = 0;
    for (let y = px.y1 + 1; y < Math.min(img.height, px.y1 + 3 * lineH) && emptyRun < lineH; y++) {
      let rowInk = 0;
      for (let x = px.x0; x < px.x1; x++) {
        if (inAnyBox(x, y, sh)) continue;
        band++;
        if (ink(x, y)) rowInk++;
      }
      spill += rowInk;
      emptyRun = rowInk ? 0 : emptyRun + 1;
    }
    if (band && spill > Math.max(12, band * 0.01))
      add("rendered_overflow", "high", [sh], `Text in “${sh.name}” renders beyond the bottom of its box.`, "Shorten the text, reduce the font size, or make the box taller; then render again.");
  }
  return issues;
}

/**
 * Deterministic stand-in renderer used by the in-memory host: draws filled
 * shapes and text lines (bars) that run past the box like real unfitted text.
 */
export function simulateRender(slide: PptSlide, size: { width: number; height: number }, widthPx = 480, background: [number, number, number] = [255, 255, 255]): RenderedImage {
  const height = Math.round((widthPx * size.height) / size.width);
  const data = new Uint8ClampedArray(widthPx * height * 4);
  for (let i = 0; i < data.length; i += 4) {
    data[i] = background[0];
    data[i + 1] = background[1];
    data[i + 2] = background[2];
    data[i + 3] = 255;
  }
  const s = widthPx / size.width;
  const rect = (x0: number, y0: number, x1: number, y1: number, rgb: [number, number, number]) => {
    for (let y = Math.max(0, Math.floor(y0)); y < Math.min(height, Math.ceil(y1)); y++)
      for (let x = Math.max(0, Math.floor(x0)); x < Math.min(widthPx, Math.ceil(x1)); x++) {
        const i = (y * widthPx + x) * 4;
        data[i] = rgb[0];
        data[i + 1] = rgb[1];
        data[i + 2] = rgb[2];
      }
  };
  const hex = (c: string | undefined, d: [number, number, number]): [number, number, number] =>
    c && /^#[0-9a-f]{6}$/i.test(c) ? [parseInt(c.slice(1, 3), 16), parseInt(c.slice(3, 5), 16), parseInt(c.slice(5, 7), 16)] : d;
  for (const sh of slide.shapes) {
    if (!sh.box) continue;
    const b = sh.box;
    if (sh.fill || sh.text === null) rect(b.left * s, b.top * s, (b.left + b.width) * s, (b.top + b.height) * s, sh.text === null ? [150, 150, 150] : hex(sh.fill, [200, 200, 200]));
    if (!sh.text?.trim()) continue;
    const size18 = sh.format?.size ?? 18;
    const color = hex(sh.format?.color, [20, 20, 20]);
    const charsPerLine = Math.max(1, Math.floor(Math.max(10, b.width - 14.4) / (size18 * 0.5)));
    let line = 0;
    for (const para of sh.text.split(/\r?\n|\u000b/)) {
      let remaining = para.length;
      do {
        const chars = Math.min(remaining, charsPerLine);
        const y = b.top + 3.6 + line * size18 * 1.2;
        rect((b.left + 7.2) * s, y * s, (b.left + 7.2 + chars * size18 * 0.5) * s, (y + size18 * 0.7) * s, color);
        remaining -= chars;
        line++;
      } while (remaining > 0);
    }
  }
  return { width: widthPx, height, data };
}
