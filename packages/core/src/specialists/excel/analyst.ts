import { citesRanges, numbersGrounded, textCriterion } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist } from "../types.js";
import { writesOnlyToNewSheets } from "./shared.js";

export const dataAnalyst: Specialist = {
  id: "excel.analyst",
  name: "Data Analyst",
  host: "excel",
  categories: ["analysis", "statistics", "comparison", "trends", "anomalies"],
  description: "Answers analytical questions about the workbook with calculations done by the application, citing the sheets and ranges used.",
  examples: [
    "Which region had the highest Q2 revenue, and by how much?",
    "Compare average order value by month and explain the trend",
    "Find anomalies in the sales data and explain them",
  ],
  abilities: ["Inspects sheets, tables and headers", "Calculates with the application (not the model)", "Cites Sheet!Range for every figure", "Can write results to a new sheet with approval"],
  context: { required: [REQUIRE.workbookOverview], optional: ["Your selected cells"] },
  instructions: [
    "You are the Data Analyst. Numbers you report must come from workbook.compute_stats results, cells you read, or formulas you wrote and verified — never from your own arithmetic.",
    "1. Use the workbook overview to decide which sheet and range answer the question. Read headers and a sample (workbook.read_range with a small address) before analysing large data.",
    "2. Compute with workbook.compute_stats (count, sum, mean, median, min, max, stdev, distinct, missing; optional groupBy). For anomalies use workbook.profile_data (IQR outliers with row numbers).",
    "3. Answer with each figure followed by its source, e.g. \"North: 255 (sum of Sales!C2:C4)\". Label conclusions drawn from the numbers as \"Calculated:\" and explanations you cannot verify from the data as \"Hypothesis:\".",
    "4. Only write to the workbook if asked; then put results on a new sheet as formulas referencing the source.",
  ].join("\n"),
  capabilities: ["selection.read", "context.read", "workbook.read_range", "workbook.compute_stats", "workbook.profile_data", "workbook.add_worksheet", "workbook.write_range"],
  constraints: { readOnly: false, rules: ["Reports only application-calculated or read numbers", "Writes only to new sheets unless you authorise changing existing data"] },
  approval: { summary: "Reading and calculating need no approval. Any write is shown for approval." },
  completion: [
    {
      id: "calculated",
      description: "Results were calculated by the application",
      required: true,
      evaluate: (e) => {
        const n = e.actions.filter((a) => (a.capability === "workbook.compute_stats" || a.capability === "workbook.profile_data") && a.status === "succeeded").length;
        const formulas = e.actions.filter((a) => a.capability === "workbook.write_range" && a.status === "succeeded" && (a.args as any).formulas).length;
        return { met: n + formulas > 0, detail: `${n} calculation(s), ${formulas} verified formula write(s)` };
      },
    },
    numbersGrounded(),
    citesRanges(true),
    textCriterion(
      "labels-hypotheses",
      "Explanations not provable from the data are labelled as hypotheses",
      (t) => {
        const speculative = /\b(likely|probably|may be|might|could be|possibly|suggests|because of)\b/i.test(t);
        return { met: !speculative || /hypothes[ie]s/i.test(t), detail: speculative ? "speculative language present" : "no speculative claims" };
      },
      false,
    ),
  ],
  verification: ["Every number in the answer is matched against calculated or read values", "Written results are re-read"],
  artifacts: ["Answer with cited ranges", "Optional results sheet with formulas"],
  limitations: ["No regression or forecasting models; statistics are descriptive", "Reads at most 5,000 cells per read and 100,000 per calculation"],
  routing: {
    patterns: [
      { re: /\b(analy[sz]e|analysis|insights?|trends?|compare|comparison|correlat\w*|anomal\w*|outliers?|growth|breakdown)\b/i, weight: 3 },
      { re: /\b(average|mean|median|total|sum of|how (many|much)|highest|lowest|top \d+|by (region|month|product|category|quarter))\b/i, weight: 2 },
    ],
  },
  snapshot: true,
  guard: (call, s) => writesOnlyToNewSheets(call, s.actions, s.request, "Data Analyst"),
};
