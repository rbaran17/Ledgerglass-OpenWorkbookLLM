/** Data validation: normalise rules, check cell values, and approve warnings. */
export {
  normalizeDataValidationRules,
  validateDataValidationRule,
  checkDataValidationBatch,
  findDataValidationsInRange,
  assertNoStopValidationIssues,
  DataValidationError,
  summarizeValidationInspection,
  resolveValidationInspection,
} from "../app/03-workbook-model.js";

/** Conditional formatting: normalise rules, evaluate them, and preview and apply changes. */
export {
  normalizeConditionalRules,
  evaluateConditionalFormatRules,
  updateConditionalFormatRules,
  assertConditionalFormatPreviewCurrent,
  sortConditionalRulesByPriority,
  conditionalRuleKinds,
} from "../app/03-workbook-model.js";

/** Defined names: preview and apply create, update, rename and delete, and expand names in formulas. */
export {
  applyDefinedNameOperation,
  assertNamePreviewCurrent,
  resolveDefinedNameValue,
} from "../app/03-workbook-model.js";

/** Review comments, hyperlinks and formula watches. */
export {
  addReviewThreadReply,
  setReviewThreadResolved,
  reviewThreadLimits,
  validateHyperlinkRecord,
  setCellHyperlink,
  validateFormulaWatches,
} from "../app/03-workbook-model.js";
