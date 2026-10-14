import {
  ExcelAdapter,
  FakeExcelDriver,
  FakeOutlookDriver,
  FakePowerPointDriver,
  FakeWordDriver,
  OllamaProvider,
  OutlookAdapter,
  PowerPointAdapter,
  WordAdapter,
  type AssistantSettings,
  type HostAdapter,
  type InferenceProvider,
} from "@loai/core";
import { OfficeExcelDriver } from "../drivers/excel";
import { OfficeOutlookDriver } from "../drivers/outlook";
import { OfficePowerPointDriver } from "../drivers/powerpoint";
import { officeRequirements } from "../drivers/requirements";
import { OfficeWordDriver } from "../drivers/word";
import type { CompanionClient } from "./companion";

export interface HostInfo {
  adapter: HostAdapter;
  /** True when running outside Office against an in-memory sample document. */
  demo: boolean;
  label: string;
}

/** Pick the adapter for the Office host we are running in. */
export function createHost(officeHost: Office.HostType | null, demoKind: string | null): HostInfo {
  if (officeHost === null || typeof Office === "undefined") return demoHost(demoKind ?? "word");
  switch (officeHost) {
    case Office.HostType.Word:
      return { adapter: new WordAdapter(new OfficeWordDriver(), officeRequirements), demo: false, label: "Word" };
    case Office.HostType.Excel:
      return { adapter: new ExcelAdapter(new OfficeExcelDriver(), officeRequirements), demo: false, label: "Excel" };
    case Office.HostType.PowerPoint:
      return { adapter: new PowerPointAdapter(new OfficePowerPointDriver(), officeRequirements), demo: false, label: "PowerPoint" };
    case Office.HostType.Outlook:
      return { adapter: new OutlookAdapter(new OfficeOutlookDriver(), officeRequirements), demo: false, label: "Outlook" };
  }
  return demoHost(demoKind ?? "word");
}

/** Sample documents for running the pane in a normal browser (development and demos). */
function demoHost(kind: string): HostInfo {
  if (kind === "excel")
    return {
      adapter: new ExcelAdapter(new FakeExcelDriver({ Sales: [["Region", "Q1", "Q2"], ["North", 120, 135], ["South", 90, 101], ["West", 75, 88]] }, "Demo workbook")),
      demo: true,
      label: "Excel (demo)",
    };
  if (kind === "powerpoint")
    return {
      adapter: new PowerPointAdapter(new FakePowerPointDriver([{ shapes: [{ name: "Title", text: "Local AI in Office" }] }, { shapes: [{ name: "Title", text: "Why" }, { name: "Body", text: "Privacy\nNo subscription\nWorks offline" }] }], "Demo deck")),
      demo: true,
      label: "PowerPoint (demo)",
    };
  if (kind === "outlook")
    return {
      adapter: new OutlookAdapter(new FakeOutlookDriver({ itemType: "message", subject: "Budget review", from: "Dana <dana@example.com>", to: ["you@example.com"], cc: [], body: "Hi — can we meet Thursday at 10 to review the Q3 budget? Please bring the vendor quotes.", attachments: [] })),
      demo: true,
      label: "Outlook (demo)",
    };
  const d = new FakeWordDriver(
    [
      { text: "Project Lighthouse — Status", style: "Title" },
      "The pilot launched in March with 40 users across two teams.",
      "Feedback has been positive; users cite faster report drafting as the main benefit.",
      "Risks: the data export takes too long and the onboarding guide is out of date.",
      "Next steps: fix exports, refresh onboarding, expand to 120 users in Q3.",
    ],
    "Demo document",
  );
  return { adapter: new WordAdapter(d), demo: true, label: "Word (demo)" };
}

/** Build the inference provider. No fallback: if this fails, the task fails. */
export function createProvider(s: AssistantSettings, companion: CompanionClient): InferenceProvider {
  if (s.transport === "companion") {
    return new OllamaProvider({
      baseUrl: companion.baseUrl + "/v1/ollama",
      headers: companion.authHeaders(),
      timeoutMs: s.requestTimeoutMs,
      // The companion itself enforces that Ollama is local; the companion URL must be local too.
      allowRemote: false,
    });
  }
  return new OllamaProvider({ baseUrl: s.ollamaEndpoint, timeoutMs: s.requestTimeoutMs, allowRemote: s.allowRemoteEndpoint });
}
