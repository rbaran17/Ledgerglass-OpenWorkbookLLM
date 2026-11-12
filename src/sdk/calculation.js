/**
 * The calculation area of the SDK: the model functions in ./model/calculation.js, with edits recorded
 * (see ./recording.js). Reference page: docs/api/calculation.md.
 */
import * as model from "./model/calculation.js";
import { recordEdits } from "./recording.js";

const area = recordEdits(model);

export const defaultCalculationSettings = area.defaultCalculationSettings;
export const normalizeCalculationSettings = area.normalizeCalculationSettings;
export const withCalculationSettings = area.withCalculationSettings;
export const calculateWorkbookWithSettings = area.calculateWorkbookWithSettings;
export const getCurrentCalculatorRecord = area.getCurrentCalculatorRecord;
export const refreshCalculationReport = area.refreshCalculationReport;
export const buildCalculationEvidence = area.buildCalculationEvidence;
export const builtinFunctionNames = area.builtinFunctionNames;
export const databaseFunctionNames = area.databaseFunctionNames;
export const daxFunctionNames = area.daxFunctionNames;
export const matrixResultFunctionNames = area.matrixResultFunctionNames;
export const statisticalFunctionNames = area.statisticalFunctionNames;
export const textCleanupFunctionNames = area.textCleanupFunctionNames;
export const dateFunctionNames = area.dateFunctionNames;
export const supportedFinancialFunctionNames =
  area.supportedFinancialFunctionNames;
export const supportedTextFunctionNames = area.supportedTextFunctionNames;
export const estimateWorkbookMemoryBytes = area.estimateWorkbookMemoryBytes;
export const resolveCalculationBudget = area.resolveCalculationBudget;
export const whatIfLimits = area.whatIfLimits;
export const worksheetCapacityLimits = area.worksheetCapacityLimits;
export const expandWhatIfScenarios = area.expandWhatIfScenarios;
export const validateWhatIfDefinition = area.validateWhatIfDefinition;
