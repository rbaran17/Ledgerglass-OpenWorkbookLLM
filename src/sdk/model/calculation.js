/** Calculation settings: validate them, attach them to a workbook, and read the defaults. */
export {
  defaultCalculationSettings,
  normalizeCalculationSettings,
  withCalculationSettings,
} from "../../app/03-workbook-model.js";

/** Running calculations: calculate with row filters, force a full recalculation, read the cached result. */
export {
  calculateWorkbookWithSettings,
  getCurrentCalculatorRecord,
  refreshCalculationReport,
} from "../../app/03-workbook-model.js";

/** Evidence: a JSON export of a calculation report, checked against the workbook it came from. */
export { buildCalculationEvidence } from "../../app/03-workbook-model.js";

/** Function catalogues: the names the calculator knows, grouped by family. */
export {
  builtinFunctionNames,
  databaseFunctionNames,
  daxFunctionNames,
  matrixResultFunctionNames,
  statisticalFunctionNames,
  textCleanupFunctionNames,
} from "../../app/03-workbook-model.js";

/** Date, text and financial function names, as the date and text modules define them. */
export { dateFunctionNames } from "../../app/01-schemas-date-functions-and-validation.js";
export {
  supportedFinancialFunctionNames,
  supportedTextFunctionNames,
} from "../../app/02-text-and-financial-functions.js";

/** Resource budgets: turn settings into byte and work limits, and estimate a workbook's working data. */
export {
  estimateWorkbookMemoryBytes,
  resolveCalculationBudget,
} from "../../app/01-schemas-date-functions-and-validation.js";

/** Limits: the fixed capacity of worksheet axes and of what-if analyses. */
export {
  whatIfLimits,
  worksheetCapacityLimits,
} from "../../app/03-workbook-model.js";

/** What-if analyses: validate a definition, then expand it into the cases it calculates. */
export {
  expandWhatIfScenarios,
  validateWhatIfDefinition,
} from "../../app/03-workbook-model.js";
