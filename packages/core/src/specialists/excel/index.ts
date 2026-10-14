import { dataAnalyst } from "./analyst.js";
import { dashboardBuilder } from "./dashboard-builder.js";
import { dataCleaner } from "./data-cleaner.js";
import { formulaExpert } from "./formula-expert.js";

export { dashboardBuilder, dataAnalyst, dataCleaner, formulaExpert };
export const EXCEL_SPECIALISTS = [dataAnalyst, formulaExpert, dataCleaner, dashboardBuilder];
