import type { ExecutionRecord, ToolCall } from "../../types.js";

/** Sheets created by this task (new sheets are safe to write to). */
export function createdSheets(actions: ExecutionRecord[]): string[] {
  return actions
    .filter((a) => a.status === "succeeded" && (a.capability === "workbook.add_worksheet" || a.capability === "workbook.clean_to_sheet"))
    .map((a) => String((a.args as Record<string, unknown>).name ?? (a.args as Record<string, unknown>).targetSheet));
}

export const OVERWRITE_AUTHORIZED = /\b(overwrite|in[- ]place|replace the (original|source|existing)|modify the (original|source)|write (it )?back)\b/i;

/** Writes may only go to sheets created in this task, unless the user authorised changing existing data. */
export function writesOnlyToNewSheets(call: ToolCall, actions: ExecutionRecord[], request: string, who: string): string | undefined {
  if (OVERWRITE_AUTHORIZED.test(request)) return undefined;
  const created = createdSheets(actions);
  const sheet =
    call.name === "workbook.write_range" ? call.arguments.sheet : call.name === "workbook.create_table" ? call.arguments.sheet : call.name === "workbook.create_chart" ? call.arguments.targetSheet ?? call.arguments.sourceSheet : undefined;
  if (sheet === undefined || typeof sheet !== "string") return undefined;
  if (!created.includes(sheet))
    return `The ${who} only writes to sheets it created in this task (${created.join(", ") || "none yet"}); "${sheet}" holds existing data. Create a new sheet with workbook.add_worksheet, or ask the user to authorise changing existing data.`;
  return undefined;
}
