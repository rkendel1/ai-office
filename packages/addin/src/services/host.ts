import {
  ExcelAdapter,
  FakeExcelDriver,
  FakeOutlookDriver,
  FakePowerPointDriver,
  FakeWordDriver,
  OllamaProvider,
  OpenAIProvider,
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
  /** The driver behind the adapter (used by the acceptance runner to create fixtures and read state). */
  driver: unknown;
  /** True when running outside Office against an in-memory sample document. */
  demo: boolean;
  label: string;
}

/** Pick the adapter for the Office host we are running in. */
export function createHost(officeHost: Office.HostType | null, demoKind: string | null): HostInfo {
  if (officeHost === null || typeof Office === "undefined") return demoHost(demoKind ?? "word");
  switch (officeHost) {
    case Office.HostType.Word:
      return withDriver(new OfficeWordDriver(), (d) => new WordAdapter(d, officeRequirements), "Word");
    case Office.HostType.Excel:
      return withDriver(new OfficeExcelDriver(), (d) => new ExcelAdapter(d, officeRequirements), "Excel");
    case Office.HostType.PowerPoint:
      return withDriver(new OfficePowerPointDriver(), (d) => new PowerPointAdapter(d, officeRequirements), "PowerPoint");
    case Office.HostType.Outlook:
      return withDriver(new OfficeOutlookDriver(), (d) => new OutlookAdapter(d, officeRequirements), "Outlook");
  }
  return demoHost(demoKind ?? "word");
}

function withDriver<D>(driver: D, make: (d: D) => HostAdapter, label: string, demo = false): HostInfo {
  return { adapter: make(driver), driver, demo, label };
}

/** Sample documents for running the pane in a normal browser (development and demos). */
function demoHost(kind: string): HostInfo {
  if (kind === "excel")
    return withDriver(new FakeExcelDriver({ Sales: [["Region", "Q1", "Q2"], ["North", 120, 135], ["South", 90, 101], ["West", 75, 88]] }, "Demo workbook"), (d) => new ExcelAdapter(d), "Excel (demo)", true);
  if (kind === "powerpoint")
    return withDriver(
      new FakePowerPointDriver([{ shapes: [{ name: "Title 1", text: "Local AI in Office" }] }, { shapes: [{ name: "Title 1", text: "Why" }, { name: "Body", text: "Privacy\nNo subscription\nWorks offline" }] }], "Demo deck"),
      (d) => new PowerPointAdapter(d),
      "PowerPoint (demo)",
      true,
    );
  if (kind === "outlook")
    return withDriver(
      new FakeOutlookDriver({ itemType: "message", subject: "Budget review", from: "Dana <dana@example.com>", to: ["you@example.com"], cc: [], body: "Hi — can we meet Thursday at 10 to review the Q3 budget? Please bring the vendor quotes.", attachments: [] }),
      (d) => new OutlookAdapter(d),
      "Outlook (demo)",
      true,
    );
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
  return withDriver(d, (x) => new WordAdapter(x), "Word (demo)", true);
}

/** Build the inference provider. No fallback: if this fails, the task fails. */
export function createProvider(s: AssistantSettings, companion: CompanionClient): InferenceProvider {
  if (s.transport === "companion") {
    if (s.provider === "ollama")
      return new OllamaProvider({
        baseUrl: companion.baseUrl + "/v1/ollama",
        headers: companion.authHeaders(),
        timeoutMs: s.requestTimeoutMs,
        // The companion itself enforces that the model app is local; the companion URL must be local too.
        allowRemote: false,
      });
    return new OpenAIProvider({
      baseUrl: companion.baseUrl + "/v1/openai",
      headers: companion.authHeaders(),
      timeoutMs: s.requestTimeoutMs,
      allowRemote: false,
    });
  }
  if (s.provider === "ollama")
    return new OllamaProvider({ baseUrl: s.endpoint, timeoutMs: s.requestTimeoutMs, allowRemote: s.allowRemoteEndpoint });
  return new OpenAIProvider({ baseUrl: s.endpoint, timeoutMs: s.requestTimeoutMs, allowRemote: s.allowRemoteEndpoint });
}
