/**
 * The live acceptance scenarios, run against in-memory hosts (mocked mode).
 * This validates the scenarios themselves; it is not a live Office result.
 */
import fs from "node:fs";
import { describe, expect, it } from "vitest";
import {
  ACCEPTANCE_SCENARIOS,
  ExcelAdapter,
  FakeExcelDriver,
  FakeOutlookDriver,
  FakePowerPointDriver,
  FakeWordDriver,
  OutlookAdapter,
  PowerPointAdapter,
  WordAdapter,
  acceptanceMatrix,
  runAcceptance,
  type AcceptanceRecord,
  type HostKind,
} from "../src/index.js";

const env = { officeApp: "in-memory", platform: "node", addinVersion: "test", recordedAt: "2026-10-13T00:00:00.000Z" };

/** Parse the fixture .eml (headers + plain-text body). */
function fixtureMail(mode: "read" | "compose" = "read") {
  const raw = fs.readFileSync(new URL("../../../acceptance/fixtures/budget-thread.eml", import.meta.url), "utf8").replace(/\r\n/g, "\n");
  const [head, ...rest] = raw.split("\n\n");
  const h = (k: string) => new RegExp(`^${k}: (.*)$`, "m").exec(head!)![1]!;
  const d = new FakeOutlookDriver({ itemType: "message", subject: h("Subject"), from: h("From"), to: [h("To")], cc: [], body: rest.join("\n\n").trim(), attachments: [] }, mode);
  return { driver: d, adapter: new OutlookAdapter(d) };
}

const hosts: Record<HostKind, () => { driver: any; adapter: any }> = {
  word: () => {
    const d = new FakeWordDriver([]);
    return { driver: d, adapter: new WordAdapter(d) };
  },
  excel: () => {
    const d = new FakeExcelDriver({ Sheet1: [] }, "Book1.xlsx");
    return { driver: d, adapter: new ExcelAdapter(d) };
  },
  powerpoint: () => {
    const d = new FakePowerPointDriver([]);
    return { driver: d, adapter: new PowerPointAdapter(d) };
  },
  outlook: () => fixtureMail(),
};

describe("acceptance scenarios (mocked hosts)", () => {
  const records: AcceptanceRecord[] = [];

  for (const sc of ACCEPTANCE_SCENARIOS) {
    it(`${sc.id} passes every automated check`, async () => {
      const r = await runAcceptance(sc, { ...hosts[sc.host](), mode: "mocked", environment: env });
      records.push(r);
      expect(r.reason).toBeUndefined();
      expect(r.checks.filter((c) => !c.passed)).toEqual([]);
      expect(r.outcome).toBe("passed");
      expect(r.mode).toBe("mocked");
      expect(r.manual.every((m) => m.status === "pending")).toBe(true);
    });
  }

  it("the reply scenario also passes in compose mode, with a verified draft", async () => {
    const sc = ACCEPTANCE_SCENARIOS.find((s) => s.id === "outlook.reply.draft-not-sent")!;
    const r = await runAcceptance(sc, { ...fixtureMail("compose"), mode: "mocked", environment: env });
    expect(r.outcome).toBe("passed");
    expect(r.task!.actions.find((a) => a.capability === "outlook.create_draft")!.status).toBe("succeeded");
  });

  it("refuses to run on a document that is not blank", async () => {
    const d = new FakeWordDriver(["My real report"]);
    const r = await runAcceptance(ACCEPTANCE_SCENARIOS[0]!, { driver: d, adapter: new WordAdapter(d), mode: "live", environment: env });
    expect(r).toMatchObject({ outcome: "blocked" });
    expect(r.reason).toMatch(/not blank/);
    expect(d.paragraphs.map((p) => p.text)).toEqual(["My real report"]);
  });

  it("is blocked when the Outlook fixture message is not open", async () => {
    const d = new FakeOutlookDriver({ itemType: "message", subject: "Lunch?", from: "a@b.c", to: [], cc: [], body: "hi", attachments: [] });
    const r = await runAcceptance(ACCEPTANCE_SCENARIOS.find((s) => s.host === "outlook")!, { driver: d, adapter: new OutlookAdapter(d), mode: "live", environment: env });
    expect(r.outcome).toBe("blocked");
  });

  it("records rendering status for the deck: rendered where supported, unsupported otherwise", async () => {
    const sc = ACCEPTANCE_SCENARIOS.find((s) => s.host === "powerpoint")!;
    const yes = await runAcceptance(sc, { ...hosts.powerpoint(), mode: "mocked", environment: env });
    expect(yes.visual).toBe("rendered-check-passed");
    const d = new FakePowerPointDriver([]);
    d.renderingSupported = false;
    const no = await runAcceptance(sc, { driver: d, adapter: new PowerPointAdapter(d), mode: "mocked", environment: env });
    expect(no.visual).toBe("unsupported");
    expect(no.outcome).toBe("passed"); // the scenario discloses it; a person still owes the visual check
    expect(no.manual.some((m) => /chart/.test(m.step) && m.status === "pending")).toBe(true);
  });

  it("detects a real failure: a dashboard whose values do not match", async () => {
    const sc = ACCEPTANCE_SCENARIOS.find((s) => s.id === "excel.dashboard.traceable-metrics")!;
    const d = new FakeExcelDriver({ Sheet1: [] });
    const orig = d.writeRange.bind(d);
    d.writeRange = async (sheet, address, data) => orig(sheet, address, sheet === "LOAI_Sales" ? { values: data.values!.map((r) => r.map((v) => (v === 120 ? 121 : v))) } : data);
    const r = await runAcceptance(sc, { driver: d, adapter: new ExcelAdapter(d), mode: "mocked", environment: env });
    expect(r.outcome).toBe("failed");
    expect(r.checks.find((c) => c.id === "excel-calculated")!.passed).toBe(false);
  });

  it("the status matrix keeps live and mocked results apart and lists scenarios never run live", () => {
    const md = acceptanceMatrix(records, ACCEPTANCE_SCENARIOS);
    expect(md).toMatch(/### Live Office \(real hosts\)\n_No live runs recorded yet._/);
    expect(md).toMatch(/Not yet run live: `word.writer.structured-document`/);
    expect(md).toMatch(/\| word.editor.scoped-edit \| in-memory/);
    expect(md).toMatch(/not evidence that the scenario works in Office/);
  });
});
