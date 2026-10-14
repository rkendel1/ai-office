/**
 * Presentation construction capabilities for PowerPoint: slides from layouts,
 * text boxes, shapes, formatting, geometry, native tables (where supported),
 * diagrams and bar charts drawn with editable shapes, speaker notes (where
 * supported), and an object-level layout check.
 *
 * Every capability validates its arguments, goes through the registry's
 * policy/approval path, and verifies by reading slide objects back.
 * Rendering is never claimed to be verified: only slide objects are inspected.
 */
import { ScopeError, type CapabilityDefinition } from "../registry.js";
import type { Compensation, Json } from "../types.js";
import { normalizeText } from "../util.js";
import type { RequirementChecker } from "./adapter.js";
import { excerpt, makeItem } from "./common.js";
import { SHAPE_KINDS, renderSlide, type Box, type PowerPointDriver, type PptShape, type PptSlide, type ShapeKind, type TextFormat } from "./powerpoint.js";

export const ACCENT = "#2F5BD3";
export const ACCENT_LIGHT = "#DCE4FA";
export const INK = "#1B1B1F";

// ------------------------------------------------------------- layout check

export type IssueKind = "off_slide" | "overlap" | "overflow" | "dense" | "no_title" | "empty_placeholder" | "small_text" | "inconsistent_title_size";

export interface LayoutIssue {
  slide: number;
  slideId: string;
  shapeId?: string;
  kind: IssueKind;
  severity: "high" | "medium" | "low";
  detail: string;
}

const area = (b: Box) => Math.max(0, b.width) * Math.max(0, b.height);
function intersection(a: Box, b: Box): number {
  const w = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left);
  const h = Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top);
  return w > 0 && h > 0 ? w * h : 0;
}

export function isTitleShape(sh: PptShape): boolean {
  return sh.text !== null && /title/i.test(sh.name) && !/subtitle/i.test(sh.name);
}

/** Estimate whether text fits its box (average glyph width ≈ 0.5em, line height ≈ 1.2em). */
export function estimateOverflow(text: string, box: Box, size = 18): { needed: number; available: number; overflows: boolean } {
  const usableWidth = Math.max(10, box.width - 14.4);
  const charsPerLine = Math.max(1, Math.floor(usableWidth / (size * 0.5)));
  const lines = text.split(/\r?\n|\u000b/).reduce((n, p) => n + Math.max(1, Math.ceil(p.length / charsPerLine)), 0);
  const needed = lines * size * 1.2 + 7.2;
  return { needed, available: box.height, overflows: needed > box.height * 1.15 };
}

export function checkLayout(slides: PptSlide[], size: { width: number; height: number }): LayoutIssue[] {
  const issues: LayoutIssue[] = [];
  const titleSizes = new Map<number, number[]>();
  for (const s of slides) {
    const n = s.index + 1;
    const add = (kind: IssueKind, severity: LayoutIssue["severity"], detail: string, shapeId?: string) => issues.push({ slide: n, slideId: s.id, shapeId, kind, severity, detail });
    const textShapes = s.shapes.filter((x) => x.text !== null);
    const title = textShapes.find(isTitleShape) ?? null;
    if (!title || !title.text?.trim()) add("no_title", "medium", "Slide has no title text.");
    else if (title.format?.size) titleSizes.set(title.format.size, [...(titleSizes.get(title.format.size) ?? []), n]);
    const words = textShapes.reduce((c, x) => c + (x.text?.trim() ? x.text.trim().split(/\s+/).length : 0), 0);
    if (words > 90) add("dense", "high", `${words} words on one slide; consider splitting or cutting.`);
    else if (words > 70) add("dense", "medium", `${words} words; dense for presenting.`);
    for (const sh of s.shapes) {
      const b = sh.box;
      if (sh.type === "Placeholder" && sh.text === "") add("empty_placeholder", "medium", `Empty placeholder "${sh.name}".`, sh.id);
      if (!b) continue;
      if (b.left < -1 || b.top < -1 || b.left + b.width > size.width + 1 || b.top + b.height > size.height + 1)
        add("off_slide", "high", `"${sh.name}" extends beyond the slide edge.`, sh.id);
      if (sh.text?.trim()) {
        const fs = sh.format?.size;
        if (fs && fs < 12) add("small_text", "low", `"${sh.name}" uses ${fs}pt text; hard to read when presented.`, sh.id);
        const o = estimateOverflow(sh.text, b, fs ?? 18);
        if (o.overflows) add("overflow", "high", `"${sh.name}" likely overflows its box (needs ~${Math.round(o.needed)}pt of ${Math.round(o.available)}pt height).`, sh.id);
      }
    }
    for (let i = 0; i < textShapes.length; i++)
      for (let j = i + 1; j < textShapes.length; j++) {
        const a = textShapes[i]!;
        const c = textShapes[j]!;
        if (!a.box || !c.box || !a.text?.trim() || !c.text?.trim()) continue;
        const overlap = intersection(a.box, c.box) / Math.max(1, Math.min(area(a.box), area(c.box)));
        if (overlap > 0.2) add("overlap", "high", `"${a.name}" and "${c.name}" overlap (${Math.round(overlap * 100)}% of the smaller box).`, c.id);
      }
  }
  if (titleSizes.size > 1)
    issues.push({
      slide: 0,
      slideId: "*",
      kind: "inconsistent_title_size",
      severity: "low",
      detail: `Title sizes vary: ${[...titleSizes.entries()].map(([sz, ns]) => `${sz}pt on slide(s) ${ns.join(", ")}`).join("; ")}.`,
    });
  const order = { high: 0, medium: 1, low: 2 };
  return issues.sort((x, y) => order[x.severity] - order[y.severity] || x.slide - y.slide);
}

// ------------------------------------------------------------ geometry

export function contentArea(size: { width: number; height: number }): Box {
  return { left: 48, top: 130, width: size.width - 96, height: size.height - 170 };
}

export type DiagramKind = "process" | "list" | "grid";

/** Deterministic placement for diagram items. Returns item boxes and connector boxes. */
export function diagramGeometry(kind: DiagramKind, n: number, a: Box): { items: Box[]; connectors: Box[] } {
  const gap = 16;
  if (kind === "process") {
    const w = (a.width - gap * (n - 1)) / n;
    const h = Math.min(a.height, Math.max(80, w * 0.75));
    const top = a.top + (a.height - h) / 2;
    return { items: Array.from({ length: n }, (_, i) => ({ left: a.left + i * (w + gap), top, width: w, height: h })), connectors: [] };
  }
  if (kind === "grid") {
    const cols = n <= 4 ? 2 : 3;
    const rows = Math.ceil(n / cols);
    const w = (a.width - gap * (cols - 1)) / cols;
    const h = (a.height - gap * (rows - 1)) / rows;
    return { items: Array.from({ length: n }, (_, i) => ({ left: a.left + (i % cols) * (w + gap), top: a.top + Math.floor(i / cols) * (h + gap), width: w, height: h })), connectors: [] };
  }
  const h = (a.height - gap * (n - 1)) / n;
  return { items: Array.from({ length: n }, (_, i) => ({ left: a.left, top: a.top + i * (h + gap), width: a.width, height: h })), connectors: [] };
}

export function barChartGeometry(values: number[], a: Box): { bars: Box[]; valueLabels: Box[]; categoryLabels: Box[]; plot: Box } {
  const labelH = 28;
  const plot = { left: a.left, top: a.top + labelH, width: a.width, height: a.height - labelH * 2 };
  const n = values.length;
  const slot = plot.width / n;
  const barW = slot * 0.6;
  const max = Math.max(...values, 0) || 1;
  const bars = values.map((v, i) => {
    const h = Math.max(1, (v / max) * plot.height);
    return { left: plot.left + i * slot + (slot - barW) / 2, top: plot.top + plot.height - h, width: barW, height: h };
  });
  return {
    plot,
    bars,
    valueLabels: bars.map((b) => ({ left: b.left - slot * 0.2, top: b.top - labelH, width: barW + slot * 0.4, height: labelH })),
    categoryLabels: values.map((_, i) => ({ left: plot.left + i * slot, top: plot.top + plot.height + 2, width: slot, height: labelH })),
  };
}

export function formatNumber(v: number, unit?: string): string {
  const abs = Math.abs(v);
  const s = abs >= 1e9 ? `${+(v / 1e9).toFixed(1)}B` : abs >= 1e6 ? `${+(v / 1e6).toFixed(1)}M` : abs >= 1e4 ? `${+(v / 1e3).toFixed(1)}K` : `${+v.toFixed(2)}`;
  return unit ? (unit === "$" ? `$${s}` : `${s}${unit}`) : s;
}

const near = (a: number, b: number, tol = 1.5) => Math.abs(a - b) <= tol;

function formatMatches(got: TextFormat | undefined, want: TextFormat): boolean | null {
  if (!got) return null;
  return (
    (want.size === undefined || got.size === want.size) &&
    (want.bold === undefined || got.bold === want.bold) &&
    (want.italic === undefined || got.italic === want.italic) &&
    (want.align === undefined || got.align === want.align) &&
    (want.color === undefined || got.color?.toUpperCase() === want.color.toUpperCase())
  );
}

// --------------------------------------------------------- capabilities

const boxProps = {
  left: { type: "number", minimum: -50, maximum: 2000 },
  top: { type: "number", minimum: -50, maximum: 2000 },
  width: { type: "number", minimum: 4, maximum: 2000 },
  height: { type: "number", minimum: 4, maximum: 2000 },
} as const;
const formatSchema = {
  type: "object",
  properties: {
    size: { type: "number", minimum: 6, maximum: 120 },
    bold: { type: "boolean" },
    italic: { type: "boolean" },
    color: { type: "string", pattern: "^#[0-9A-Fa-f]{6}$" },
    align: { type: "string", enum: ["left", "center", "right"] },
  },
} as const;
const colorSchema = { type: "string", pattern: "^#[0-9A-Fa-f]{6}$" } as const;

export function buildCapabilities(d: PowerPointDriver, req: RequirementChecker): CapabilityDefinition<any>[] {
  const caps: CapabilityDefinition<any>[] = [];
  const v13 = req.isSetSupported("PowerPointApi", "1.3");
  const v14 = req.isSetSupported("PowerPointApi", "1.4");

  const slide = async (id: string) => {
    const s = (await d.slides()).find((x) => x.id === id);
    if (!s) throw new ScopeError(`Slide id ${id} does not exist.`);
    return s;
  };
  const shapeOn = async (slideId: string, shapeId: string) => {
    const s = await slide(slideId);
    const sh = s.shapes.find((x) => x.id === shapeId);
    if (!sh) throw new ScopeError(`Shape ${shapeId} is not on slide ${s.index + 1}.`);
    return { s, sh };
  };
  const within = async (b: Partial<Box>) => {
    const size = await d.slideSize();
    if (b.left !== undefined && b.width !== undefined && b.left + b.width > size.width + 1) throw new ScopeError(`The box extends beyond the slide width (${size.width}pt).`);
    if (b.top !== undefined && b.height !== undefined && b.top + b.height > size.height + 1) throw new ScopeError(`The box extends beyond the slide height (${size.height}pt).`);
  };
  const removeShapes = (slideId: string, shapeIds: string[], what: string): Compensation => ({
    description: `Remove the ${what}`,
    capability: "presentation.delete_shapes",
    args: { slideId, shapeIds },
  });

  if (v13) {
    caps.push({
      name: "presentation.list_layouts",
      description: "List the slide layouts available in the presentation's slide master (e.g. Title Slide, Title and Content, Two Content).",
      hosts: ["powerpoint"],
      risk: "read_active",
      modifies: false,
      inputSchema: { type: "object", properties: {} },
      async execute() {
        const ls = await d.layouts();
        return { output: { layouts: ls.map((l) => l.name), slideSize: (await d.slideSize()) as unknown as Json } };
      },
    });

    type CreateArgs = { title: string; bullets?: string[]; body?: string; layout?: string; notes?: string; subtitle?: string };
    caps.push({
      name: "presentation.create_slide",
      description:
        "Create one complete slide at the end of the deck: title plus either bullets (preferred) or body text, optional subtitle (for title slides), optional layout name from presentation.list_layouts, and optional speaker notes. Fills layout placeholders when present, otherwise adds formatted text boxes.",
      hosts: ["powerpoint"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["title"],
        properties: {
          title: { type: "string", minLength: 1, maxLength: 200 },
          subtitle: { type: "string", maxLength: 300 },
          bullets: { type: "array", maxItems: 12, items: { type: "string", minLength: 1, maxLength: 400 } },
          body: { type: "string", maxLength: 4000 },
          layout: { type: "string", maxLength: 100 },
          notes: { type: "string", maxLength: 8000 },
        },
      },
      async resolveTarget(a: CreateArgs) {
        if (a.bullets && a.body) throw new ScopeError("Provide bullets or body, not both.");
        if (a.layout) {
          const ls = await d.layouts();
          if (!ls.some((l) => l.name.toLowerCase() === a.layout!.toLowerCase()))
            throw new ScopeError(`Layout "${a.layout}" does not exist. Available: ${ls.map((l) => l.name).join(", ") || "(none reported)"}`);
        }
        return `new slide “${a.title}”`;
      },
      async preview(a: CreateArgs) {
        const n = (await d.slides()).length + 1;
        const warnings = a.notes && !d.supportsNotes() ? ["Speaker notes are not supported by this PowerPoint version; they will be reported in chat instead."] : [];
        return {
          summary: `Add slide ${n}${a.layout ? ` (${a.layout})` : ""}`,
          target: "End of presentation",
          after: [a.title, a.subtitle, ...(a.bullets ?? []).map((b) => `• ${b}`), a.body, a.notes && `Notes: ${a.notes}`].filter(Boolean).join("\n"),
          warnings,
        };
      },
      async execute(a: CreateArgs) {
        const size = await d.slideSize();
        const ls = await d.layouts();
        const content = a.bullets?.length || a.body;
        const pick =
          (a.layout && ls.find((l) => l.name.toLowerCase() === a.layout!.toLowerCase())) ||
          (a.subtitle && !content ? ls.find((l) => /^title slide$/i.test(l.name)) : undefined) ||
          (content ? ls.find((l) => /^title and content$/i.test(l.name)) : ls.find((l) => /^title only$/i.test(l.name)));
        const id = await d.addSlide(pick ? { layoutId: pick.id, masterId: pick.masterId } : undefined);
        const fresh = await slide(id);
        const placeholders = fresh.shapes.filter((x) => x.text !== null && x.type === "Placeholder");
        const titlePh = placeholders.find(isTitleShape);
        const subtitlePh = placeholders.find((x) => /subtitle/i.test(x.name));
        const bodyPh = placeholders.find((x) => x !== titlePh && x !== subtitlePh);
        const bodyText = a.bullets?.length ? a.bullets.join("\n") : a.body;
        const used = new Set<string>();

        if (titlePh) {
          await d.setShapeText(id, titlePh.id, a.title);
          used.add(titlePh.id);
        } else used.add(await d.addTextBox(id, a.title, { left: 48, top: 32, width: size.width - 96, height: 80 }, { size: 32, bold: true, color: INK }));

        if (a.subtitle) {
          if (subtitlePh) {
            await d.setShapeText(id, subtitlePh.id, a.subtitle);
            used.add(subtitlePh.id);
          } else used.add(await d.addTextBox(id, a.subtitle, { left: 48, top: 120, width: size.width - 96, height: 50 }, { size: 20, color: INK }));
        }
        if (bodyText) {
          if (bodyPh) {
            await d.setShapeText(id, bodyPh.id, bodyText);
            used.add(bodyPh.id);
          } else {
            const text = a.bullets?.length ? a.bullets.map((b) => `• ${b}`).join("\n") : bodyText;
            used.add(await d.addTextBox(id, text, contentArea(size), { size: 20, color: INK }));
          }
        }
        // Remove placeholders the slide does not use, so no "Click to add text" boxes remain.
        for (const ph of placeholders) if (!used.has(ph.id)) await d.deleteShape(id, ph.id);

        let notes: "none" | "added" | "not_supported" = "none";
        let notesOk = true;
        if (a.notes) {
          if (d.supportsNotes()) {
            await d.setNotes(id, a.notes);
            notesOk = normalizeText((await d.getNotes(id)) ?? "") === normalizeText(a.notes);
            notes = "added";
          } else notes = "not_supported";
        }

        const after = await slide(id);
        const texts = after.shapes.map((x) => normalizeText(x.text ?? ""));
        const want = [a.title, a.subtitle, a.bullets?.length ? null : a.body].filter((x): x is string => !!x).map(normalizeText);
        const bulletsOk = !a.bullets?.length || a.bullets.every((b) => texts.some((t) => t.includes(normalizeText(b))));
        const ok = want.every((w) => texts.some((t) => t === w || t.includes(w))) && bulletsOk && notesOk;
        return {
          output: { slideId: id, slideNumber: after.index + 1, layout: pick?.name ?? "text boxes", notes },
          verification: {
            verified: ok,
            method: "re-read slide shapes" + (notes === "added" ? " and notes" : ""),
            evidence: excerpt(renderSlide(after), 300),
            mismatch: ok ? undefined : "Slide text does not match the requested content.",
          },
          partial: notes === "not_supported",
          message: notes === "not_supported" ? "Slide created, but speaker notes are not supported by this PowerPoint version and were not added." : undefined,
          compensation: { description: `Delete slide ${after.index + 1}`, capability: "presentation.delete_slide", args: { slideId: id } },
        };
      },
    });
  }

  if (v14) {
    type TextBoxArgs = Box & { slideId: string; text: string; format?: TextFormat };
    caps.push({
      name: "presentation.add_text_box",
      description: "Add a text box to a slide at a position (points; slide is typically 960×540) with optional formatting.",
      hosts: ["powerpoint"],
      risk: "modify",
      modifies: true,
      inputSchema: { type: "object", required: ["slideId", "text", "left", "top", "width", "height"], properties: { slideId: { type: "string", minLength: 1 }, text: { type: "string", minLength: 1, maxLength: 4000 }, ...boxProps, format: formatSchema } } as any,
      async resolveTarget(a: TextBoxArgs) {
        const s = await slide(a.slideId);
        await within(a);
        return `text box on slide ${s.index + 1}`;
      },
      async preview(a: TextBoxArgs) {
        return { summary: "Add a text box", target: `Slide ${(await slide(a.slideId)).index + 1}`, after: a.text };
      },
      async execute(a: TextBoxArgs) {
        const id = await d.addTextBox(a.slideId, a.text, a, a.format);
        const { sh } = await shapeOn(a.slideId, id);
        const textOk = normalizeText(sh.text ?? "") === normalizeText(a.text);
        const boxOk = !sh.box || (near(sh.box.left, a.left) && near(sh.box.top, a.top));
        const fmt = a.format ? formatMatches(sh.format, a.format) : true;
        return {
          output: { shapeId: id },
          verification: { verified: textOk && boxOk && fmt !== false, method: "re-read shape text, position and format", mismatch: textOk && boxOk && fmt !== false ? undefined : "Text box does not match the request." },
          compensation: removeShapes(a.slideId, [id], "text box"),
        };
      },
    });

    type ShapeArgs = Box & { slideId: string; shape: ShapeKind; fill?: string; line?: string; text?: string; format?: TextFormat };
    caps.push({
      name: "presentation.add_shape",
      description: `Add a shape (${SHAPE_KINDS.join(", ")}) with optional fill colour, outline colour and text.`,
      hosts: ["powerpoint"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["slideId", "shape", "left", "top", "width", "height"],
        properties: { slideId: { type: "string", minLength: 1 }, shape: { type: "string", enum: SHAPE_KINDS as unknown as Json[] }, ...boxProps, fill: colorSchema, line: colorSchema, text: { type: "string", maxLength: 1000 }, format: formatSchema },
      } as any,
      async resolveTarget(a: ShapeArgs) {
        const s = await slide(a.slideId);
        await within(a);
        return `${a.shape} on slide ${s.index + 1}`;
      },
      async preview(a: ShapeArgs) {
        return { summary: `Add a ${a.shape}${a.fill ? ` filled ${a.fill}` : ""}`, target: `Slide ${(await slide(a.slideId)).index + 1}`, after: a.text };
      },
      async execute(a: ShapeArgs) {
        const id = await d.addShape(a.slideId, a.shape, a, { fill: a.fill, line: a.line, text: a.text, format: a.format });
        const { sh } = await shapeOn(a.slideId, id);
        const ok = (a.text === undefined || normalizeText(sh.text ?? "") === normalizeText(a.text)) && (!a.fill || !sh.fill || sh.fill.toUpperCase() === a.fill.toUpperCase());
        return {
          output: { shapeId: id },
          verification: { verified: ok, method: "re-read shape", mismatch: ok ? undefined : "Shape does not match the request." },
          compensation: removeShapes(a.slideId, [id], a.shape),
        };
      },
    });

    type FormatArgs = { slideId: string; shapeId: string; format: TextFormat };
    caps.push({
      name: "presentation.format_text",
      description: "Apply text formatting (size, bold, italic, colour, alignment) to all text in a shape.",
      hosts: ["powerpoint"],
      risk: "modify",
      modifies: true,
      inputSchema: { type: "object", required: ["slideId", "shapeId", "format"], properties: { slideId: { type: "string", minLength: 1 }, shapeId: { type: "string", minLength: 1 }, format: formatSchema } } as any,
      async resolveTarget(a: FormatArgs) {
        const { s, sh } = await shapeOn(a.slideId, a.shapeId);
        if (sh.text === null) throw new ScopeError(`Shape ${a.shapeId} has no text.`);
        if (!Object.keys(a.format).length) throw new ScopeError("Specify at least one formatting property.");
        return `text format of “${sh.name}” on slide ${s.index + 1}`;
      },
      async preview(a: FormatArgs) {
        const { s, sh } = await shapeOn(a.slideId, a.shapeId);
        return { summary: `Format text: ${Object.entries(a.format).map(([k, v]) => `${k} ${v}`).join(", ")}`, target: `Slide ${s.index + 1} · ${sh.name}`, before: excerpt(sh.text ?? "", 300) };
      },
      async execute(a: FormatArgs) {
        const { sh: before } = await shapeOn(a.slideId, a.shapeId);
        await d.formatText(a.slideId, a.shapeId, a.format);
        const { sh } = await shapeOn(a.slideId, a.shapeId);
        const m = formatMatches(sh.format, a.format);
        const textKept = normalizeText(sh.text ?? "") === normalizeText(before.text ?? "");
        return {
          output: { formatted: a.shapeId },
          verification: {
            verified: m === null ? null : m && textKept,
            method: "re-read text format",
            mismatch: m === false ? "Formatting does not match after the change." : !textKept ? "Text changed unexpectedly." : undefined,
            evidence: m === null ? "This PowerPoint version does not report text formatting." : undefined,
          },
          compensation: before.format
            ? { description: "Restore previous formatting", capability: "presentation.format_text", args: { slideId: a.slideId, shapeId: a.shapeId, format: Object.fromEntries(Object.entries(before.format).filter(([, v]) => v !== undefined)) as Json } }
            : undefined,
        };
      },
    });

    type GeoArgs = Partial<Box> & { slideId: string; shapeId: string };
    caps.push({
      name: "presentation.set_geometry",
      description: "Move and/or resize a shape (points). The result must stay on the slide.",
      hosts: ["powerpoint"],
      risk: "modify",
      modifies: true,
      inputSchema: { type: "object", required: ["slideId", "shapeId"], properties: { slideId: { type: "string", minLength: 1 }, shapeId: { type: "string", minLength: 1 }, ...boxProps } } as any,
      async resolveTarget(a: GeoArgs) {
        const { s, sh } = await shapeOn(a.slideId, a.shapeId);
        if (!sh.box) throw new ScopeError("This PowerPoint version does not report shape geometry.");
        if ([a.left, a.top, a.width, a.height].every((x) => x === undefined)) throw new ScopeError("Specify left, top, width or height.");
        await within({ ...sh.box, ...Object.fromEntries(Object.entries(a).filter(([k, v]) => v !== undefined && ["left", "top", "width", "height"].includes(k))) });
        return `position/size of “${sh.name}” on slide ${s.index + 1}`;
      },
      async preview(a: GeoArgs) {
        const { sh } = await shapeOn(a.slideId, a.shapeId);
        const fmt = (b: Partial<Box>) => `${Math.round(b.left!)},${Math.round(b.top!)} ${Math.round(b.width!)}×${Math.round(b.height!)}`;
        return { summary: "Move/resize shape", target: sh.name, before: fmt(sh.box!), after: fmt({ ...sh.box!, ...Object.fromEntries(Object.entries(a).filter(([k, v]) => v !== undefined && k !== "slideId" && k !== "shapeId")) }) };
      },
      async execute(a: GeoArgs) {
        const { sh: before } = await shapeOn(a.slideId, a.shapeId);
        const box = { left: a.left, top: a.top, width: a.width, height: a.height };
        await d.setGeometry(a.slideId, a.shapeId, box);
        const { sh } = await shapeOn(a.slideId, a.shapeId);
        const ok = !!sh.box && (["left", "top", "width", "height"] as const).every((k) => box[k] === undefined || near(sh.box![k], box[k]!));
        const textKept = normalizeText(sh.text ?? "") === normalizeText(before.text ?? "");
        return {
          output: { moved: a.shapeId },
          verification: { verified: ok && textKept, method: "re-read shape geometry", mismatch: ok ? (textKept ? undefined : "Text changed unexpectedly.") : "Geometry does not match." },
          compensation: { description: "Restore previous position and size", capability: "presentation.set_geometry", args: { slideId: a.slideId, shapeId: a.shapeId, ...before.box! } },
        };
      },
    });

    type DiagramArgs = { slideId: string; kind: DiagramKind; items: { title: string; detail?: string }[]; fill?: string } & Partial<Box>;
    caps.push({
      name: "presentation.add_diagram",
      description:
        "Add an editable diagram built from shapes: process (left-to-right steps), list (stacked boxes) or grid. 2–8 items, each with a short title and optional detail line. Placed in the content area unless left/top/width/height are given.",
      hosts: ["powerpoint"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["slideId", "kind", "items"],
        properties: {
          slideId: { type: "string", minLength: 1 },
          kind: { type: "string", enum: ["process", "list", "grid"] },
          items: { type: "array", minItems: 2, maxItems: 8, items: { type: "object", required: ["title"], properties: { title: { type: "string", minLength: 1, maxLength: 80 }, detail: { type: "string", maxLength: 200 } } } },
          fill: colorSchema,
          ...boxProps,
        },
      } as any,
      async resolveTarget(a: DiagramArgs) {
        const s = await slide(a.slideId);
        if (a.left !== undefined) await within(a);
        return `${a.kind} diagram (${a.items.length} items) on slide ${s.index + 1}`;
      },
      async preview(a: DiagramArgs) {
        return { summary: `Add a ${a.kind} diagram`, target: `Slide ${(await slide(a.slideId)).index + 1}`, after: a.items.map((i, k) => `${k + 1}. ${i.title}${i.detail ? ` — ${i.detail}` : ""}`).join("\n") };
      },
      async execute(a: DiagramArgs) {
        const size = await d.slideSize();
        const region = a.left !== undefined && a.top !== undefined && a.width && a.height ? (a as Box) : contentArea(size);
        const g = diagramGeometry(a.kind, a.items.length, region);
        const kind: ShapeKind = a.kind === "process" ? "chevron" : "roundedRectangle";
        const ids: string[] = [];
        for (let i = 0; i < a.items.length; i++) {
          const it = a.items[i]!;
          const text = it.detail ? `${it.title}\n${it.detail}` : it.title;
          ids.push(await d.addShape(a.slideId, kind, g.items[i]!, { fill: a.fill ?? ACCENT, text, format: { size: a.items.length > 5 ? 14 : 16, color: "#FFFFFF", bold: false, align: "center" } }));
        }
        const s = await slide(a.slideId);
        const found = ids.map((id) => s.shapes.find((x) => x.id === id));
        const ok = found.every((sh, i) => sh && normalizeText(sh.text ?? "").startsWith(normalizeText(a.items[i]!.title)));
        return {
          output: { shapeIds: ids, kind: a.kind, note: "Diagram is made of individual editable shapes." },
          verification: { verified: ok, method: "re-read diagram shapes", mismatch: ok ? undefined : "Some diagram shapes are missing or have the wrong text." },
          compensation: removeShapes(a.slideId, ids, "diagram"),
        };
      },
    });

    type ChartArgs = { slideId: string; categories: string[]; values: number[]; unit?: string; title?: string; fill?: string } & Partial<Box>;
    caps.push({
      name: "presentation.add_bar_chart",
      description:
        "Add a column chart drawn with editable shapes (bars, value labels, category labels) from validated numbers. Use only numbers taken from inspected content. It is not a native chart object: Office.js cannot create native PowerPoint charts.",
      hosts: ["powerpoint"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["slideId", "categories", "values"],
        properties: {
          slideId: { type: "string", minLength: 1 },
          categories: { type: "array", minItems: 2, maxItems: 12, items: { type: "string", minLength: 1, maxLength: 40 } },
          values: { type: "array", minItems: 2, maxItems: 12, items: { type: "number", minimum: 0 } },
          unit: { type: "string", maxLength: 8 },
          title: { type: "string", maxLength: 120 },
          fill: colorSchema,
          ...boxProps,
        },
      } as any,
      async resolveTarget(a: ChartArgs) {
        if (a.categories.length !== a.values.length) throw new ScopeError(`${a.categories.length} categories but ${a.values.length} values.`);
        if (a.values.every((v) => v === 0)) throw new ScopeError("All values are zero.");
        const s = await slide(a.slideId);
        if (a.left !== undefined) await within(a);
        return `bar chart (${a.values.length} bars) on slide ${s.index + 1}`;
      },
      async preview(a: ChartArgs) {
        return {
          summary: `Add a bar chart${a.title ? ` “${a.title}”` : ""} (editable shapes, not a native chart)`,
          target: `Slide ${(await slide(a.slideId)).index + 1}`,
          after: a.categories.map((c, i) => `${c}: ${formatNumber(a.values[i]!, a.unit)}`).join("\n"),
        };
      },
      async execute(a: ChartArgs) {
        const size = await d.slideSize();
        let region = a.left !== undefined && a.top !== undefined && a.width && a.height ? (a as Box) : contentArea(size);
        const ids: string[] = [];
        if (a.title) {
          ids.push(await d.addTextBox(a.slideId, a.title, { left: region.left, top: region.top, width: region.width, height: 28 }, { size: 16, bold: true, color: INK }));
          region = { ...region, top: region.top + 32, height: region.height - 32 };
        }
        const g = barChartGeometry(a.values, region);
        const barIds: string[] = [];
        const labelIds: string[] = [];
        for (let i = 0; i < a.values.length; i++) {
          barIds.push(await d.addShape(a.slideId, "rectangle", g.bars[i]!, { fill: a.fill ?? ACCENT }));
          labelIds.push(await d.addTextBox(a.slideId, formatNumber(a.values[i]!, a.unit), g.valueLabels[i]!, { size: 12, align: "center", color: INK }));
          ids.push(barIds[i]!, labelIds[i]!, await d.addTextBox(a.slideId, a.categories[i]!, g.categoryLabels[i]!, { size: 12, align: "center", color: INK }));
        }
        const s = await slide(a.slideId);
        const bars = barIds.map((id) => s.shapes.find((x) => x.id === id));
        const labels = labelIds.map((id) => s.shapes.find((x) => x.id === id));
        const max = Math.max(...a.values);
        const tallest = Math.max(...bars.map((b) => b?.box?.height ?? 0));
        const proportional = bars.every((b, i) => b && (!b.box || Math.abs(b.box.height / tallest - a.values[i]! / max) < 0.02 || (a.values[i] === 0 && b.box.height <= 1)));
        const labelsOk = labels.every((l, i) => l && normalizeText(l.text ?? "") === formatNumber(a.values[i]!, a.unit));
        const ok = bars.every(Boolean) && proportional && labelsOk;
        return {
          output: { shapeIds: ids, bars: a.values.length, note: "Drawn with editable shapes; not a native chart object." },
          verification: {
            verified: ok,
            method: "re-read bars, labels and bar heights (proportional to values)",
            mismatch: ok ? undefined : !labelsOk ? "Value labels do not match the data." : "Bar heights are not proportional to the values.",
          },
          compensation: removeShapes(a.slideId, ids, "chart"),
        };
      },
    });
  }

  if (d.supportsTables()) {
    type TableArgs = { slideId: string; rows: string[][] } & Partial<Box>;
    caps.push({
      name: "presentation.add_table",
      description: "Add a native table (first row is the header). Up to 15 rows × 8 columns of short text.",
      hosts: ["powerpoint"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["slideId", "rows"],
        properties: { slideId: { type: "string", minLength: 1 }, rows: { type: "array", minItems: 2, maxItems: 15, items: { type: "array", minItems: 1, maxItems: 8, items: { type: "string", maxLength: 200 } } }, ...boxProps },
      } as any,
      async resolveTarget(a: TableArgs) {
        const w = a.rows[0]!.length;
        if (a.rows.some((r) => r.length !== w)) throw new ScopeError("All rows must have the same number of cells.");
        const s = await slide(a.slideId);
        return `table ${a.rows.length}×${w} on slide ${s.index + 1}`;
      },
      async preview(a: TableArgs) {
        return { summary: `Add a ${a.rows.length}×${a.rows[0]!.length} table`, target: `Slide ${(await slide(a.slideId)).index + 1}`, after: a.rows.map((r) => r.join(" | ")).join("\n") };
      },
      async execute(a: TableArgs) {
        const size = await d.slideSize();
        const box = a.left !== undefined && a.top !== undefined && a.width && a.height ? (a as Box) : { ...contentArea(size), height: Math.min(contentArea(size).height, 32 * a.rows.length) };
        const id = await d.addTable(a.slideId, a.rows, box);
        const { sh } = await shapeOn(a.slideId, id);
        const text = normalizeText(sh.text ?? "");
        const verified = sh.text === null ? null : a.rows.flat().every((c) => text.includes(normalizeText(c)));
        return {
          output: { shapeId: id },
          verification: { verified, method: "re-read table cells", evidence: verified === null ? "Table created; this PowerPoint version does not report cell text." : undefined, mismatch: verified === false ? "Table cells do not match." : undefined },
          compensation: removeShapes(a.slideId, [id], "table"),
        };
      },
    });
  }

  if (d.supportsNotes()) {
    caps.push({
      name: "presentation.set_notes",
      description: "Set the speaker notes of a slide (replaces existing notes).",
      hosts: ["powerpoint"],
      risk: "modify",
      modifies: true,
      inputSchema: { type: "object", required: ["slideId", "text"], properties: { slideId: { type: "string", minLength: 1 }, text: { type: "string", maxLength: 8000 } } },
      async resolveTarget(a: { slideId: string }) {
        return `speaker notes of slide ${(await slide(a.slideId)).index + 1}`;
      },
      async preview(a: { slideId: string; text: string }) {
        return { summary: "Set speaker notes", target: `Slide ${(await slide(a.slideId)).index + 1}`, before: (await d.getNotes(a.slideId)) ?? "", after: a.text };
      },
      async execute(a: { slideId: string; text: string }) {
        const before = (await d.getNotes(a.slideId)) ?? "";
        await d.setNotes(a.slideId, a.text);
        const ok = normalizeText((await d.getNotes(a.slideId)) ?? "") === normalizeText(a.text);
        return {
          output: { slideId: a.slideId },
          verification: { verified: ok, method: "re-read notes" },
          compensation: { description: "Restore previous notes", capability: "presentation.set_notes", args: { slideId: a.slideId, text: before } },
        };
      },
    });
  }

  caps.push({
    name: "presentation.check_layout",
    description:
      "Check slide objects for layout problems: text likely overflowing its box, overlapping text, shapes off the slide, missing titles, empty placeholders, dense slides, small text, inconsistent title sizes. Inspects object geometry and text only; it does not see the rendered slide.",
    hosts: ["powerpoint"],
    risk: "read_active",
    modifies: false,
    inputSchema: { type: "object", properties: { slideIds: { type: "array", maxItems: 200, items: { type: "string", minLength: 1 } } } },
    async execute(a: { slideIds?: string[] }) {
      const all = await d.slides();
      const chosen = a.slideIds ? all.filter((s) => a.slideIds!.includes(s.id)) : all;
      const size = await d.slideSize();
      const issues = checkLayout(chosen, size);
      const geometryReported = chosen.some((s) => s.shapes.some((x) => x.box));
      const method = `Object-level check of ${chosen.length} slide(s)${geometryReported ? "" : " (this PowerPoint version reports no geometry: position checks skipped)"}${size.reported ? "" : "; slide size assumed 960×540"}. Rendering was not inspected.`;
      const item = makeItem({
        host: "powerpoint",
        id: `ppt:layout:${a.slideIds?.join(",") ?? "all"}`,
        label: `Layout check (${issues.length} issue${issues.length === 1 ? "" : "s"})`,
        scope: "partial",
        content: `${method}\n${issues.map((i) => `- [${i.severity}] slide ${i.slide || "deck"}: ${i.detail}`).join("\n") || "No issues found."}`,
        source: "presentation.check_layout",
      });
      return { output: { issues: issues as unknown as Json, method, renderingVerified: false }, context: [item] };
    },
  });

  caps.push({
    name: "presentation.delete_shapes",
    description: "Delete shapes (undo only).",
    hosts: ["powerpoint"],
    risk: "destructive",
    modifies: true,
    internal: true,
    inputSchema: { type: "object", required: ["slideId", "shapeIds"], properties: { slideId: { type: "string" }, shapeIds: { type: "array", minItems: 1, maxItems: 100, items: { type: "string" } } } },
    async resolveTarget(a: { slideId: string; shapeIds: string[] }) {
      const s = await slide(a.slideId);
      return `${a.shapeIds.length} shape(s) on slide ${s.index + 1}`;
    },
    async preview(a: { slideId: string; shapeIds: string[] }) {
      const s = await slide(a.slideId);
      return { summary: `Delete ${a.shapeIds.length} shape(s)`, target: `Slide ${s.index + 1}`, before: s.shapes.filter((x) => a.shapeIds.includes(x.id)).map((x) => x.text ?? x.name).join("\n") };
    },
    async execute(a: { slideId: string; shapeIds: string[] }) {
      for (const id of a.shapeIds) await d.deleteShape(a.slideId, id).catch(() => undefined);
      const s = await slide(a.slideId);
      const gone = a.shapeIds.every((id) => !s.shapes.some((x) => x.id === id));
      return { output: { deleted: a.shapeIds.length }, verification: { verified: gone, method: "list shapes" } };
    },
  });

  return caps;
}
