import type { Freshness } from "../context.js";
import { ScopeError, type CapabilityDefinition } from "../registry.js";
import type { ContextItem, ItemIdentity } from "../types.js";
import { normalizeText } from "../util.js";
import { allSupported, type ApiSupportEntry, type HostAdapter, type RequirementChecker } from "./adapter.js";
import { excerpt, makeItem } from "./common.js";

export interface PptShape {
  id: string;
  name: string;
  type: string;
  /** null = the shape has no text frame (picture, chart, ...) */
  text: string | null;
}

export interface PptSlide {
  id: string;
  index: number; // 0-based
  shapes: PptShape[];
}

export interface PowerPointDriver {
  info(): Promise<{ id?: string; title?: string }>;
  slides(): Promise<PptSlide[]>;
  /** Selected slide ids, or null if the host cannot report selection. */
  selectedSlideIds(): Promise<string[] | null>;
  setShapeText(slideId: string, shapeId: string, text: string): Promise<void>;
  /** Adds a slide at the end; returns its id. */
  addSlide(): Promise<string>;
  addTextBox(slideId: string, text: string, box: { left: number; top: number; width: number; height: number }): Promise<string>;
  deleteSlide(slideId: string): Promise<void>;
}

function renderSlide(s: PptSlide): string {
  const lines = [`Slide ${s.index + 1} (id ${s.id})`];
  for (const sh of s.shapes) {
    if (sh.text === null) lines.push(`  - shape ${sh.id} "${sh.name}" [${sh.type}, no text]`);
    else lines.push(`  - shape ${sh.id} "${sh.name}": ${sh.text.replace(/\r?\n|\r|\u000b/g, " / ")}`);
  }
  return lines.join("\n");
}

function unsupportedOf(slides: PptSlide[]) {
  const n = slides.reduce((c, s) => c + s.shapes.filter((x) => x.text === null).length, 0);
  return n ? [{ kind: "non-text shapes (pictures, charts, media)", count: n }] : [];
}

export class PowerPointAdapter implements HostAdapter {
  readonly host = "powerpoint" as const;

  constructor(
    private readonly d: PowerPointDriver,
    private readonly req: RequirementChecker = allSupported,
  ) {}

  async identity(): Promise<ItemIdentity> {
    const i = await this.d.info();
    return { host: "powerpoint", id: i.id, title: i.title };
  }

  apiSupport(): ApiSupportEntry[] {
    return [
      { requirement: "PowerPointApi 1.3", supported: this.req.isSetSupported("PowerPointApi", "1.3"), capabilities: ["presentation.add_slide"] },
      { requirement: "PowerPointApi 1.4", supported: this.req.isSetSupported("PowerPointApi", "1.4"), capabilities: ["presentation.read_slides", "presentation.edit_text", "text boxes"] },
      { requirement: "PowerPointApi 1.5", supported: this.req.isSetSupported("PowerPointApi", "1.5"), capabilities: ["selected slides"] },
      { requirement: "Speaker notes", supported: false, capabilities: ["Not available through Office.js; not offered"] },
    ];
  }

  private async slidesItem(ids?: string[]): Promise<ContextItem> {
    const all = await this.d.slides();
    const chosen = ids ? all.filter((s) => ids.includes(s.id)) : all;
    const full = !ids || chosen.length === all.length;
    return makeItem({
      host: "powerpoint",
      id: full ? "ppt:slides:all" : `ppt:slides:${chosen.map((s) => s.id).join(",")}`,
      label: full ? `All ${all.length} slides` : `Slides ${chosen.map((s) => s.index + 1).join(", ")} of ${all.length}`,
      scope: full ? "full" : "partial",
      content: chosen.map(renderSlide).join("\n\n") || "(no slides)",
      unsupported: unsupportedOf(chosen),
      truncated: !full,
      source: "presentation.read_slides",
    });
  }

  async initialContext(): Promise<ContextItem[]> {
    const info = await this.d.info();
    const slides = await this.d.slides();
    const sel = await this.d.selectedSlideIds();
    const items: ContextItem[] = [
      makeItem({
        host: "powerpoint",
        id: "ppt:meta",
        label: "Presentation overview",
        scope: "metadata",
        content: `Title: ${info.title ?? "(untitled)"}\nSlides: ${slides.length}\nSelected slides: ${
          sel === null ? "(not reported by this PowerPoint version)" : sel.length ? sel.map((id) => (slides.find((s) => s.id === id)?.index ?? -1) + 1).join(", ") : "none"
        }`,
        source: "initial",
      }),
    ];
    if (sel && sel.length) {
      const it = await this.slidesItem(sel);
      items.push({ ...it, id: "ppt:selection", label: `Selected ${it.label}`, scope: "selection" });
    }
    return items;
  }

  async checkFreshness(item: ContextItem): Promise<Freshness> {
    const m = /^ppt:slides:(.+)$/.exec(item.id);
    if (!m) return "unknown";
    const now = await this.slidesItem(m[1] === "all" ? undefined : m[1]!.split(","));
    return now.fingerprint === item.fingerprint ? "fresh" : "stale";
  }

  private async findShape(slideId: string, shapeId: string) {
    const slide = (await this.d.slides()).find((s) => s.id === slideId);
    if (!slide) throw new ScopeError(`Slide id ${slideId} does not exist.`);
    const shape = slide.shapes.find((s) => s.id === shapeId);
    if (!shape) throw new ScopeError(`Shape ${shapeId} is not on slide ${slide.index + 1}.`);
    if (shape.text === null) throw new ScopeError(`Shape ${shapeId} has no text frame and cannot be edited as text.`);
    return { slide, shape };
  }

  capabilities(): CapabilityDefinition<any>[] {
    const d = this.d;
    const self = this;
    const caps: CapabilityDefinition<any>[] = [];

    caps.push({
      name: "selection.read",
      description: "Read the text of the slides the user has selected.",
      hosts: ["powerpoint"],
      risk: "read_selection",
      modifies: false,
      inputSchema: { type: "object", properties: {} },
      async execute() {
        const sel = await d.selectedSlideIds();
        if (sel === null) return { output: { unsupported: true, note: "This PowerPoint version does not report selected slides. Use presentation.read_slides." } };
        if (!sel.length) return { output: { empty: true, note: "No slides are selected." } };
        const it = await self.slidesItem(sel);
        return { output: { item: it.id }, context: [it] };
      },
    });

    const readSlides: CapabilityDefinition<any> = {
      name: "presentation.read_slides",
      description: "Read slide text and shape ids. Omit slideIds to read all slides.",
      hosts: ["powerpoint"],
      risk: "read_active",
      modifies: false,
      inputSchema: {
        type: "object",
        properties: { slideIds: { type: "array", maxItems: 200, items: { type: "string", minLength: 1 } } },
      },
      async resolveTarget(a: { slideIds?: string[] }) {
        if (!a.slideIds) return "all slides";
        const ids = new Set((await d.slides()).map((s) => s.id));
        const missing = a.slideIds.filter((x) => !ids.has(x));
        if (missing.length) throw new ScopeError(`Unknown slide id(s): ${missing.join(", ")}`);
        return `${a.slideIds.length} slide(s)`;
      },
      async execute(a: { slideIds?: string[] }) {
        const it = await self.slidesItem(a.slideIds);
        return { output: { item: it.id, label: it.label }, context: [it] };
      },
    };
    caps.push(readSlides, { ...readSlides, name: "context.read", description: "Read the text of all slides with slide and shape ids." });

    type EditArgs = { slideId: string; shapeId: string; text: string; expectedText?: string };
    caps.push({
      name: "presentation.edit_text",
      description: "Replace the text of a shape on a slide. Provide expectedText (the current text) to guard against stale edits.",
      hosts: ["powerpoint"],
      risk: "modify",
      modifies: true,
      inputSchema: {
        type: "object",
        required: ["slideId", "shapeId", "text"],
        properties: {
          slideId: { type: "string", minLength: 1 },
          shapeId: { type: "string", minLength: 1 },
          text: { type: "string", maxLength: 20_000 },
          expectedText: { type: "string", maxLength: 20_000 },
        },
      },
      async resolveTarget(a: EditArgs) {
        const { slide, shape } = await self.findShape(a.slideId, a.shapeId);
        return `slide ${slide.index + 1}, shape “${shape.name}”`;
      },
      async preview(a: EditArgs) {
        const { slide, shape } = await self.findShape(a.slideId, a.shapeId);
        const warnings =
          a.expectedText !== undefined && normalizeText(a.expectedText) !== normalizeText(shape.text!)
            ? ["The shape's text differs from what the assistant read. This change will be refused."]
            : [];
        return { summary: "Replace shape text", target: `Slide ${slide.index + 1} · ${shape.name}`, before: excerpt(shape.text!), after: excerpt(a.text), warnings };
      },
      async execute(a: EditArgs) {
        const { shape } = await self.findShape(a.slideId, a.shapeId);
        if (a.expectedText !== undefined && normalizeText(a.expectedText) !== normalizeText(shape.text!))
          throw new Error("The shape's text changed since it was read. Re-read the slide and try again.");
        await d.setShapeText(a.slideId, a.shapeId, a.text);
        const after = (await self.findShape(a.slideId, a.shapeId)).shape.text ?? "";
        const ok = normalizeText(after) === normalizeText(a.text);
        return {
          output: { updated: `${a.slideId}/${a.shapeId}` },
          verification: { verified: ok, method: "re-read shape text", evidence: excerpt(after, 200), mismatch: ok ? undefined : "Shape text does not match after the edit." },
          compensation: { description: "Restore the previous shape text", capability: "presentation.edit_text", args: { slideId: a.slideId, shapeId: a.shapeId, text: shape.text!, expectedText: a.text } },
        };
      },
    });

    if (this.req.isSetSupported("PowerPointApi", "1.3")) {
      type AddArgs = { title: string; body?: string };
      caps.push({
        name: "presentation.add_slide",
        description: "Add a new slide at the end with a title and optional body text (use \\n for separate lines).",
        hosts: ["powerpoint"],
        risk: "modify",
        modifies: true,
        inputSchema: {
          type: "object",
          required: ["title"],
          properties: { title: { type: "string", minLength: 1, maxLength: 500 }, body: { type: "string", maxLength: 10_000 } },
        },
        async preview(a: AddArgs) {
          return { summary: "Add a slide at the end", target: "Presentation", after: `${a.title}\n\n${a.body ?? ""}`.trim() };
        },
        async execute(a: AddArgs) {
          const id = await d.addSlide();
          await d.addTextBox(id, a.title, { left: 40, top: 30, width: 640, height: 70 });
          if (a.body) await d.addTextBox(id, a.body, { left: 40, top: 120, width: 640, height: 360 });
          const slide = (await d.slides()).find((s) => s.id === id);
          const texts = slide?.shapes.map((s) => normalizeText(s.text ?? "")) ?? [];
          const ok = !!slide && texts.includes(normalizeText(a.title)) && (!a.body || texts.includes(normalizeText(a.body)));
          return {
            output: { slideId: id, slideNumber: slide ? slide.index + 1 : null },
            verification: { verified: ok, method: "re-read new slide", evidence: slide ? renderSlide(slide).slice(0, 200) : undefined },
            compensation: { description: "Delete the added slide", capability: "presentation.delete_slide", args: { slideId: id } },
          };
        },
      });

      caps.push({
        name: "presentation.delete_slide",
        description: "Delete a slide (undo only).",
        hosts: ["powerpoint"],
        risk: "destructive",
        modifies: true,
        internal: true,
        inputSchema: { type: "object", required: ["slideId"], properties: { slideId: { type: "string" } } },
        async resolveTarget(a: { slideId: string }) {
          if (!(await d.slides()).some((s) => s.id === a.slideId)) throw new ScopeError("Slide no longer exists.");
          return `slide ${a.slideId}`;
        },
        async preview(a: { slideId: string }) {
          const s = (await d.slides()).find((x) => x.id === a.slideId)!;
          return { summary: "Delete slide", target: `Slide ${s.index + 1}`, before: renderSlide(s) };
        },
        async execute(a: { slideId: string }) {
          await d.deleteSlide(a.slideId);
          const gone = !(await d.slides()).some((s) => s.id === a.slideId);
          return { output: { deleted: a.slideId }, verification: { verified: gone, method: "list slides" } };
        },
      });
    }
    return caps;
  }
}

