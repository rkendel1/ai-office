/**
 * PowerPoint quality assurance (hardening PR §3): every object-level check and
 * every rendered check, with passing and failing fixtures, and the
 * "unsupported" path when rendering is unavailable.
 */
import { describe, expect, it } from "vitest";
import {
  FakePowerPointDriver,
  PowerPointAdapter,
  ScriptedProvider,
  analyzeRendering,
  checkLayout,
  simulateRender,
  type PptShape,
  type PptSlide,
} from "../src/index.js";
import { setup } from "./helpers.js";

const SIZE = { width: 960, height: 540 };
const sh = (id: string, name: string, text: string | null, box?: PptShape["box"], format?: PptShape["format"], type = text === null ? "Image" : "TextBox"): PptShape => ({ id, name, type, text, box, format });
const slide = (index: number, shapes: PptShape[], layoutName?: string): PptSlide => ({ id: `s${index + 1}`, index, shapes, layoutName });
const title = (t = "A clear title", box = { left: 48, top: 30, width: 864, height: 80 }, size = 36) => sh("t", "Title 1", t, box, { size, bold: true, font: "Aptos Display" }, "Placeholder");
const body = (t = "Point one\nPoint two\nPoint three", box = { left: 48, top: 130, width: 864, height: 360 }, format: PptShape["format"] = { size: 20, font: "Aptos" }) => sh("b", "Body", t, box, format);
const kinds = (s: PptSlide[]) => checkLayout(s, SIZE).map((i) => `${i.slide}:${i.kind}:${i.severity}`);

describe("object-level checks", () => {
  it("a clean slide has no findings", () => {
    expect(checkLayout([slide(0, [title(), body()], "Title and Content")], SIZE)).toEqual([]);
  });

  it("off-slide objects", () => {
    expect(kinds([slide(0, [title(), body("x", { left: 700, top: 130, width: 400, height: 100 })])])).toContain("1:off_slide:high");
  });

  it("estimated text overflow, with a sized remediation", () => {
    const f = checkLayout([slide(0, [title(), body("word ".repeat(80), { left: 48, top: 130, width: 400, height: 60 })])], SIZE).find((i) => i.kind === "overflow")!;
    expect(f).toMatchObject({ severity: "high", basis: "text-estimate", objects: [{ id: "b", name: "Body" }] });
    expect(f.remediation).toMatch(/enlarge the box to about \d+pt tall/);
  });

  it("empty and nearly empty slides", () => {
    expect(kinds([slide(0, [])])).toContain("1:empty_slide:high");
    expect(kinds([slide(0, [title("Hi")])])).toContain("1:nearly_empty:medium");
    expect(kinds([slide(0, [title("Hi"), sh("p", "Picture 3", null, { left: 100, top: 150, width: 300, height: 200 })])])).not.toContain("1:nearly_empty:medium");
  });

  it("overlaps: text on text is high, a picture over text is medium", () => {
    const overlapText = kinds([slide(0, [title(), body("a b c", { left: 48, top: 40, width: 864, height: 300 })])]);
    expect(overlapText).toContain("1:overlap:high");
    const picture = checkLayout([slide(0, [title(), body(), sh("p", "Picture 3", null, { left: 48, top: 130, width: 500, height: 300 })])], SIZE).find((i) => i.kind === "overlap")!;
    expect(picture).toMatchObject({ severity: "medium", objects: [{ name: "Picture 3" }, { name: "Body" }] });
  });

  it("missing titles depend on the layout", () => {
    expect(kinds([slide(0, [body()], "Title and Content")])).toContain("1:no_title:medium");
    expect(kinds([slide(0, [body()], "Blank")])).toContain("1:no_title:low");
    const f = checkLayout([slide(0, [body()], "Title Only")], SIZE).find((i) => i.kind === "no_title")!;
    expect(f.detail).toMatch(/layout \(“Title Only”\) has one/);
  });

  it("empty placeholders, dense slides and small text", () => {
    const ph = sh("x", "Content Placeholder 2", "", { left: 48, top: 130, width: 864, height: 300 }, undefined, "Placeholder");
    expect(kinds([slide(0, [title(), body(), ph])])).toContain("1:empty_placeholder:medium");
    expect(kinds([slide(0, [title(), body("word ".repeat(95), { left: 48, top: 130, width: 864, height: 400 }, { size: 14 })])])).toContain("1:dense:high");
    expect(kinds([slide(0, [title(), body("tiny", undefined, { size: 9 })])])).toContain("1:small_text:low");
  });

  it("deck consistency: title size and position, body sizes, fonts", () => {
    const deck = [
      slide(0, [title("One", undefined, 36), body("a b c d e", undefined, { size: 20, font: "Aptos" })]),
      slide(1, [title("Two", undefined, 28), body("a b c d e", undefined, { size: 16, font: "Arial" })]),
      slide(2, [title("Three", { left: 200, top: 200, width: 600, height: 80 }, 36), body("a b c d e", undefined, { size: 24, font: "Comic Sans MS" })]),
    ];
    const k = kinds(deck);
    expect(k).toEqual(expect.arrayContaining(["0:inconsistent_title_size:low", "0:inconsistent_title_position:low", "0:inconsistent_body_size:low", "0:inconsistent_fonts:low"]));
  });

  it("every finding names its basis, objects and a remediation", () => {
    const all = checkLayout([slide(0, [body("word ".repeat(100), { left: 700, top: 400, width: 400, height: 60 }, { size: 9 })], "Title and Content")], SIZE);
    for (const f of all) {
      expect(["geometry", "text-estimate"]).toContain(f.basis);
      expect(f.remediation.length).toBeGreaterThan(10);
    }
  });
});

describe("rendered checks", () => {
  it("detects text rendering past the bottom of its box", () => {
    const s = slide(0, [title(), body("word ".repeat(60), { left: 48, top: 130, width: 400, height: 60 })]);
    const f = analyzeRendering(s, simulateRender(s, SIZE, 640), SIZE);
    expect(f).toEqual([expect.objectContaining({ kind: "rendered_overflow", severity: "high", basis: "rendered", objects: [{ id: "b", name: "Body" }] })]);
  });

  it("a clean slide renders without findings, on light and dark backgrounds", () => {
    const s = slide(0, [title(), body()]);
    expect(analyzeRendering(s, simulateRender(s, SIZE, 640), SIZE)).toEqual([]);
    const dark = slide(0, [title("Dark", undefined, 36), body("Light text", undefined, { size: 20, color: "#F0F0F0" })]);
    dark.shapes[0]!.format!.color = "#FFFFFF";
    expect(analyzeRendering(dark, simulateRender(dark, SIZE, 640, [20, 24, 40]), SIZE)).toEqual([]);
  });

  it("detects text that does not render (same colour as the background)", () => {
    const s = slide(0, [title(), body("Invisible", undefined, { size: 20, color: "#FFFFFF" })]);
    const f = analyzeRendering(s, simulateRender(s, SIZE, 640), SIZE);
    expect(f.map((x) => x.kind)).toEqual(["rendered_text_missing"]);
  });

  it("detects a slide that renders blank although it has text", () => {
    const s = slide(0, [title(), body()]);
    const blank = simulateRender(slide(0, []), SIZE, 640);
    expect(analyzeRendering(s, blank, SIZE).map((x) => x.kind)).toEqual(["rendered_blank"]);
  });

  it("does not mistake an object below a text box for overflow", () => {
    const s = slide(0, [title(), body("short", { left: 48, top: 130, width: 400, height: 60 }), sh("p", "Picture", null, { left: 48, top: 195, width: 400, height: 100 })]);
    expect(analyzeRendering(s, simulateRender(s, SIZE, 640), SIZE).filter((x) => x.kind === "rendered_overflow")).toEqual([]);
  });
});

describe("check_rendering capability", () => {
  const deck = () => new FakePowerPointDriver([{ shapes: [{ name: "Title 1", text: "Big claims" }, { name: "Body", text: "word ".repeat(80), box: { left: 48, top: 130, width: 400, height: 60 }, format: { size: 20 } }] }]);

  it("renders slides and returns rendered findings when supported", async () => {
    const s = setup(new PowerPointAdapter(deck()), new ScriptedProvider([]));
    const r = await s.runtime.executeDirect({ id: "x", name: "presentation.check_rendering", arguments: {} });
    expect(r.output).toMatchObject({ supported: true, status: "checked", renderingVerified: true, slidesRendered: 1 });
    expect((r.output as any).issues.map((i: any) => i.kind)).toContain("rendered_overflow");
  });

  it("reports unsupported — never passed — when rendering is unavailable", async () => {
    const ppt = deck();
    ppt.renderingSupported = false;
    const s = setup(new PowerPointAdapter(ppt), new ScriptedProvider([]));
    const r = await s.runtime.executeDirect({ id: "x", name: "presentation.check_rendering", arguments: {} });
    expect(r.output).toMatchObject({ supported: false, status: "unsupported", renderingVerified: false });
    expect((r.output as any).issues).toBeUndefined();
  });

  it("reports slides that failed to render as not visually checked", async () => {
    const ppt = deck();
    ppt.renderOverride = () => {
      throw new Error("render failed");
    };
    const s = setup(new PowerPointAdapter(ppt), new ScriptedProvider([]));
    const r = await s.runtime.executeDirect({ id: "x", name: "presentation.check_rendering", arguments: {} });
    expect(r.output).toMatchObject({ supported: true, status: "partial", renderingVerified: false, notRendered: [1] });
  });

  it("the deck reviewer must disclose that slides were not visually checked when rendering is unsupported", async () => {
    const ppt = deck();
    ppt.renderingSupported = false;
    const steps = (text: string) => new ScriptedProvider([{ call: { name: "presentation.check_layout", arguments: {} } }, { call: { name: "presentation.check_rendering", arguments: {} } }, { text }]);
    const s1 = setup(new PowerPointAdapter(ppt), steps("High: Slide 1 text overflows its box. The slides look fine visually."));
    const r1 = await s1.runtime.run("Review this deck", { specialist: s1.specialists.get("powerpoint.reviewer")! });
    expect(r1.completion!.find((c) => c.id === "rendering-accounted")!.met).toBe(false);

    const s2 = setup(new PowerPointAdapter(ppt), steps("High: Slide 1 text likely overflows its box (object check). Rendering is not supported here, so the slides were not visually checked."));
    const r2 = await s2.runtime.run("Review this deck", { specialist: s2.specialists.get("powerpoint.reviewer")! });
    expect(r2.completion!.filter((c) => !c.met)).toEqual([]);
  });

  it("visual claims are accepted only with rendered evidence", async () => {
    const claim = "High: Slide 1 overflows. I visually verified every slide.";
    const mk = (supported: boolean) => {
      const ppt = deck();
      ppt.renderingSupported = supported;
      return setup(new PowerPointAdapter(ppt), new ScriptedProvider([{ call: { name: "presentation.check_layout", arguments: {} } }, { call: { name: "presentation.check_rendering", arguments: {} } }, { text: claim }]));
    };
    const yes = mk(true);
    const r = await yes.runtime.run("Review this deck", { specialist: yes.specialists.get("powerpoint.reviewer")! });
    expect(r.completion!.find((c) => c.id === "no-rendering-claims")!.met).toBe(true);
    const no = mk(false);
    const r2 = await no.runtime.run("Review this deck", { specialist: no.specialists.get("powerpoint.reviewer")! });
    expect(r2.completion!.find((c) => c.id === "no-rendering-claims")!.met).toBe(false);
  });
});
