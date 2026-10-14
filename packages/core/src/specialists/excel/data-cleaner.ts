import { preservesUnrelatedContent } from "../shared/completion.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist } from "../types.js";
import { writesOnlyToNewSheets } from "./shared.js";

const WANTS_FIX = /\b(clean|fix|remove|dedupe|de-?duplicate|normali[sz]e|standardi[sz]e|fill|tidy|unify|correct)\b/i;

export const dataCleaner: Specialist = {
  id: "excel.cleaner",
  name: "Data Cleaner",
  host: "excel",
  categories: ["duplicates", "missing values", "formats", "outliers", "categories"],
  description: "Finds data-quality problems and writes a cleaned copy to a new sheet, leaving your original data untouched.",
  examples: [
    "Find duplicates and missing values in the Customers sheet",
    "Clean this table: trim spaces, unify the Region spellings and remove duplicate rows",
    "Check the data quality of the Orders table",
  ],
  abilities: ["Profiles each column", "Explains each proposed transformation", "Writes a cleaned copy to a new sheet", "Verifies row counts and that the source is unchanged"],
  context: { required: [REQUIRE.workbookOverview], optional: [] },
  instructions: [
    "You are the Data Cleaner.",
    "1. Run workbook.profile_data on the relevant range and report the issues found (duplicates, missing values, numbers stored as text, inconsistent category spellings, outliers) with row numbers.",
    "2. If the user asked for cleaning, explain each transformation in one line, then apply them with workbook.clean_to_sheet to a NEW sheet named like \"<Source> (clean)\". Choose only operations that address issues found.",
    "3. Outliers are reported, not removed or changed, unless the user explicitly asks.",
    "4. Report the row counts from the result (rows in, rows out, removed duplicates, dropped incomplete rows).",
    "The original data must never be overwritten unless the user explicitly authorises it.",
  ].join("\n"),
  capabilities: ["context.read", "workbook.read_range", "workbook.profile_data", "workbook.clean_to_sheet", "workbook.add_worksheet", "workbook.write_range"],
  constraints: { readOnly: false, rules: ["Original data is preserved; cleaned data goes to a new sheet", "Outliers are reported, not changed"] },
  approval: { summary: "The cleaned sheet is shown (operations and row counts) for approval before it is created." },
  completion: [
    {
      id: "profiled",
      description: "The data was profiled",
      required: true,
      evaluate: (e) => {
        const ok = e.actions.some((a) => a.capability === "workbook.profile_data" && a.status === "succeeded");
        return { met: ok, detail: ok ? "profile run" : "no profile" };
      },
    },
    {
      id: "cleaned-copy",
      description: "A cleaned copy was written to a new sheet and verified (row counts, unchanged source)",
      required: true,
      evaluate: (e) => {
        const c = e.actions.filter((a) => a.capability === "workbook.clean_to_sheet");
        if (!WANTS_FIX.test(e.request) && !c.length) return { met: true, detail: "only an assessment was requested" };
        const ok = c.find((a) => a.status === "succeeded");
        return { met: !!ok, detail: ok ? String(ok.verification?.evidence ?? "verified") : "no verified cleaned copy" };
      },
    },
    preservesUnrelatedContent(),
  ],
  verification: ["The cleaned sheet is re-read", "Rows in = rows out + removed duplicates + dropped rows", "The source range is compared before and after"],
  artifacts: ["Data-quality report in chat", "A new sheet with the cleaned copy"],
  limitations: ["Does not change data in place unless you explicitly authorise it", "Date-format normalisation is not supported", "Up to 100,000 cells per operation"],
  routing: {
    patterns: [
      { re: /\b(clean(ing|up)?|dedupe|de-?duplicate|duplicates?|missing values?|blanks?|data quality|tidy|normali[sz]e|standardi[sz]e)\b/i, weight: 3 },
      { re: /\b(inconsistent|misspell\w*|spellings?|trailing spaces|whitespace|stored as text)\b/i, weight: 2 },
      { re: /\boutliers?\b/i, weight: 1 },
    ],
  },
  snapshot: true,
  guard: (call, s) => writesOnlyToNewSheets(call, s.actions, s.request, "Data Cleaner"),
};
