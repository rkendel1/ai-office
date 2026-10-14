import { isApprovedPlan } from "../shared/planning.js";
import { numbersGrounded, preservesUnrelatedContent, verifiedAction } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist } from "../types.js";
import { createdSheets, writesOnlyToNewSheets } from "./shared.js";

export const dashboardBuilder: Specialist = {
  id: "excel.dashboard",
  name: "Dashboard Builder",
  host: "excel",
  categories: ["dashboard", "summary", "kpi", "charts"],
  description: "Builds an editable dashboard sheet: metrics as formulas linked to your data, summary tables and charts.",
  examples: [
    "Create a summary dashboard of revenue by region and quarter",
    "Find anomalies in this workbook, explain them, and create a summary dashboard",
    "Build a KPI sheet with totals, averages and a chart for the leadership meeting",
  ],
  abilities: ["Profiles the data and picks metrics", "Proposes the dashboard layout for approval", "Writes metrics as formulas that reference the source", "Creates charts and verifies them"],
  context: { required: [REQUIRE.workbookOverview], optional: ["Audience and key questions, if you give them"] },
  instructions: [
    "You are the Dashboard Builder.",
    "1. Determine audience and key questions from the request (ask if unclear and it matters). Read headers and a sample of the source; use workbook.compute_stats / workbook.profile_data to understand the data (and anomalies, if asked).",
    "2. Call plan.propose with the dashboard design: the new sheet name, each metric (what it computes and from which range), each chart (type and source range).",
    "3. After approval: create the sheet with workbook.add_worksheet, write labels and metrics with workbook.write_range using `formulas` so every metric references the source range (e.g. =SUM(Sales!C2:C100)). Never type computed numbers as values.",
    "4. Create charts with workbook.create_chart from the summary range on the dashboard sheet.",
    "5. In your final answer, list each metric with its formula source, and report the values shown in the results you received.",
  ].join("\n"),
  capabilities: ["context.read", "workbook.read_range", "workbook.compute_stats", "workbook.profile_data", "plan.propose", "workbook.add_worksheet", "workbook.write_range", "workbook.create_table", "workbook.create_chart"],
  constraints: {
    requirePlan: true,
    readOnly: false,
    rules: ["Works on a new dashboard sheet; source data is not modified", "Metrics are formulas referencing source ranges (no typed numbers)"],
  },
  approval: { summary: "You approve the dashboard plan, then the sheet, each block of metrics and each chart." },
  completion: [
    { id: "plan-approved", description: "The dashboard plan was approved", required: true, evaluate: (e) => ({ met: e.actions.some(isApprovedPlan), detail: e.actions.some(isApprovedPlan) ? "approved" : "no approved plan" }) },
    { id: "sheet-created", description: "A dashboard sheet was created", required: true, evaluate: (e) => ({ met: createdSheets(e.actions).length > 0, detail: createdSheets(e.actions).join(", ") || "none" }) },
    {
      id: "traceable-metrics",
      description: "Metrics are verified formulas that reference the source data",
      required: true,
      evaluate: (e) => {
        const writes = e.actions.filter((a) => a.capability === "workbook.write_range" && a.status === "succeeded");
        const refs = writes.flatMap((a) => ((a.args as any).formulas ?? []).flat()).filter((f: unknown) => typeof f === "string" && /^=.*\b[A-Za-z0-9_']+!\$?[A-Z]+\$?\d+/.test(f));
        return { met: refs.length > 0, detail: `${refs.length} formula(s) linked to source ranges` };
      },
    },
    verifiedAction("chart", "At least one chart was created and verified", ["workbook.create_chart"]),
    numbersGrounded(),
    preservesUnrelatedContent(),
  ],
  verification: ["Formulas are re-read and their values checked for errors", "Charts are listed on the sheet after creation", "Source data is compared before and after"],
  artifacts: ["Dashboard sheet with linked metrics", "Charts", "Optional summary table"],
  limitations: ["Chart styling beyond type and title is not set", "Slicers, pivot tables and conditional formatting are not created"],
  routing: {
    patterns: [
      { re: /\bdashboards?\b/i, weight: 4 },
      { re: /\b(summary (sheet|tab|table|page)|kpis?|scorecard|metrics sheet)\b/i, weight: 3 },
      { re: /\b(charts?|graphs?|visuali[sz]\w*|plot)\b/i, weight: 2 },
    ],
    subsumes: ["excel.analyst"],
  },
  snapshot: true,
  guard(call, s) {
    const w = writesOnlyToNewSheets(call, s.actions, s.request, "Dashboard Builder");
    if (w) return w;
    if (call.name === "workbook.write_range" && Array.isArray(call.arguments.values) && call.arguments.values.flat().some((v) => typeof v === "number"))
      return "Write metrics as formulas that reference the source range (use `formulas`), not typed numbers, so every figure stays traceable.";
    return undefined;
  },
};
