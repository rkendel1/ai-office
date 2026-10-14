import { preservesUnrelatedContent, textCriterion } from "../shared/completion.js";
import { cellReferences } from "../shared/evidence.js";
import { REQUIRE } from "../shared/requirements.js";
import type { Specialist } from "../types.js";

const WANTS_WRITE = /\b(create|write|add|put|insert|enter|build|fix|repair|correct|replace|fill)\b/i;

export const formulaExpert: Specialist = {
  id: "excel.formula",
  name: "Formula Expert",
  host: "excel",
  categories: ["formulas", "functions", "errors", "lookup"],
  description: "Creates, explains, debugs and repairs formulas, then checks the written formula and the value it produces.",
  examples: [
    "Add a formula in D2:D10 that calculates margin from revenue and cost",
    "Why does E5 show #N/A? Fix it",
    "Explain the formula in the selected cell",
  ],
  abilities: ["Reads the target cells and their inputs", "Explains assumptions and references", "Writes formulas after approval", "Re-reads formulas and resulting values or errors"],
  context: { required: [REQUIRE.workbookOverview], optional: ["Your selected cells"] },
  instructions: [
    "You are the Formula Expert.",
    "1. Identify the target cells (selection or the address given) and read them and the cells they depend on, including headers, so you know the data types.",
    "2. Prefer functions available in all current Excel versions (SUM, SUMIFS, AVERAGEIFS, INDEX/MATCH, IFERROR, XLOOKUP only if the workbook already uses it).",
    "3. Write formulas with workbook.write_range using `formulas` (one formula per cell; relative references adjusted per row).",
    "4. After writing, check the resultPreview in the result: if any cell evaluates to an error (#N/A, #VALUE!, #REF!, #DIV/0!, #NAME?), explain and fix it.",
    "5. Explain each formula: what it computes, which ranges it references (Sheet!Range), and its assumptions.",
  ].join("\n"),
  capabilities: ["selection.read", "context.read", "workbook.read_range", "workbook.write_range"],
  constraints: { readOnly: false, rules: ["Only the target cells are written"] },
  approval: { summary: "Each formula write is shown with the current and new cell contents for approval." },
  completion: [
    {
      id: "formula-verified",
      description: "The formula was written and read back from the cell",
      required: true,
      evaluate: (e) => {
        const writes = e.actions.filter((a) => a.capability === "workbook.write_range" && (a.args as any).formulas);
        if (!WANTS_WRITE.test(e.request) && !writes.length) return { met: true, detail: "explanation only; nothing to write" };
        const ok = writes.filter((a) => a.status === "succeeded");
        return { met: ok.length > 0, detail: ok.length ? `${ok.length} formula write(s) verified` : "no verified formula write" };
      },
    },
    {
      id: "no-formula-errors",
      description: "Written formulas evaluate without Excel errors",
      required: true,
      evaluate: (e) => {
        const errs = e.actions.filter((a) => a.capability === "workbook.write_range" && a.status === "succeeded").reduce((n, a) => n + Number((a.output as any)?.formulaErrors ?? 0), 0);
        const last = [...e.actions].reverse().find((a) => a.capability === "workbook.write_range" && a.status === "succeeded");
        const lastErrs = Number((last?.output as any)?.formulaErrors ?? 0);
        return { met: lastErrs === 0, detail: errs ? `${errs} error value(s) seen${lastErrs ? " in the final write" : ", fixed by a later write"}` : "no error values" };
      },
    },
    textCriterion("explains-references", "The explanation names the cells and ranges involved", (t) => {
      const refs = cellReferences(t).length + (t.match(/\b[A-Z]{1,3}\$?\d+(?::[A-Z]{1,3}\$?\d+)?\b/g) ?? []).length;
      return { met: refs > 0, detail: refs ? `${refs} reference(s) explained` : "no cell references in the explanation" };
    }),
    preservesUnrelatedContent(),
  ],
  verification: ["Formulas are read back from the cells", "Resulting values are inspected for Excel errors", "Other cells are compared before and after"],
  artifacts: ["Formulas in the target cells", "Explanation in chat"],
  limitations: ["Cannot create named ranges, data validation or array-formula spill configuration", "Dynamic array and LAMBDA support depends on your Excel version"],
  routing: {
    patterns: [
      { re: /\bformulas?\b/i, weight: 3 },
      { re: /\b(vlookup|xlookup|index\s*\/?\s*match|sumifs?|countifs?|averageifs?|iferror|nested if)\b/i, weight: 3 },
      { re: /#(ref!?|value!?|n\/a|div\/0!?|name\??|spill!?)/i, weight: 3 },
    ],
  },
  snapshot: true,
};
