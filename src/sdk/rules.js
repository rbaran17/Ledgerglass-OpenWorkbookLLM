/**
 * The rules area of the SDK: the model functions in ./model/rules.js, with edits recorded
 * (see ./recording.js). Reference page: docs/api/rules.md.
 */
import * as model from "./model/rules.js";
import { recordEdits } from "./recording.js";

const area = recordEdits(model);

export const normalizeDataValidationRules = area.normalizeDataValidationRules;
export const validateDataValidationRule = area.validateDataValidationRule;
export const checkDataValidationBatch = area.checkDataValidationBatch;
export const findDataValidationsInRange = area.findDataValidationsInRange;
export const assertNoStopValidationIssues = area.assertNoStopValidationIssues;
export const DataValidationError = area.DataValidationError;
export const summarizeValidationInspection = area.summarizeValidationInspection;
export const resolveValidationInspection = area.resolveValidationInspection;
export const normalizeConditionalRules = area.normalizeConditionalRules;
export const evaluateConditionalFormatRules =
  area.evaluateConditionalFormatRules;
export const updateConditionalFormatRules = area.updateConditionalFormatRules;
export const assertConditionalFormatPreviewCurrent =
  area.assertConditionalFormatPreviewCurrent;
export const sortConditionalRulesByPriority =
  area.sortConditionalRulesByPriority;
export const conditionalRuleKinds = area.conditionalRuleKinds;
export const applyDefinedNameOperation = area.applyDefinedNameOperation;
export const assertNamePreviewCurrent = area.assertNamePreviewCurrent;
export const resolveDefinedNameValue = area.resolveDefinedNameValue;
export const addReviewThreadReply = area.addReviewThreadReply;
export const setReviewThreadResolved = area.setReviewThreadResolved;
export const reviewThreadLimits = area.reviewThreadLimits;
export const validateHyperlinkRecord = area.validateHyperlinkRecord;
export const validateFormulaWatches = area.validateFormulaWatches;
