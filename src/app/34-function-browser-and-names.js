import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  parseFormulaTree,
  dateFunctionNames,
} from "./01-schemas-date-functions-and-validation.js";
import {
  databaseFunctionNames,
  builtinFunctionNames,
  statisticalFunctionNames,
  applyDefinedNameOperation,
  calculateWorkbookWithSettings,
  assertNamePreviewCurrent,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  functionDescriptions,
  functionArgumentSignatures,
} from "./33-page-setup-and-print.js";

const absFunctionSpec = { family: "local", signature: "number" },
  accrintFunctionSpec = {
    family: "financial",
    signature:
      "issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method]",
  },
  accrintmFunctionSpec = {
    family: "financial",
    signature: "issue, settlement, rate, par, [basis]",
  },
  acosFunctionSpec = { family: "local", signature: "number" },
  acoshFunctionSpec = { family: "local", signature: "number" },
  acotFunctionSpec = { family: "local", signature: "number" },
  acothFunctionSpec = { family: "local", signature: "number" },
  addressFunctionSpec = {
    family: "reference",
    signature: "rowNumber, columnNumber, [absNum], [a1], [sheetText]",
  },
  aggregateFunctionSpec = {
    family: "math",
    signature: "function_num, options, ref1 or array, [ref2 or k], …",
  },
  andFunctionSpec = { family: "logical", signature: "params1, [params2], …" },
  arabicFunctionSpec = { family: "local", signature: "text" },
  areasFunctionSpec = { family: "local", signature: "reference" },
  ascFunctionSpec = { family: "text", signature: "text" },
  asinFunctionSpec = { family: "local", signature: "number" },
  asinhFunctionSpec = { family: "local", signature: "number" },
  atanFunctionSpec = { family: "local", signature: "number" },
  atan2FunctionSpec = { family: "trigonometry", signature: "x, y" },
  atanhFunctionSpec = { family: "local", signature: "number" },
  avedevFunctionSpec = {
    family: "statistical",
    signature: "numbers1, [numbers2], …",
  },
  averageFunctionSpec = {
    family: "statistical",
    signature: "number1, [number2], …",
  },
  averageaFunctionSpec = {
    family: "statistical",
    signature: "numbers1, [numbers2], …",
  },
  averageifFunctionSpec = {
    family: "statistical",
    signature: "range, criteria, averageRange",
  },
  averageifsFunctionSpec = {
    family: "statistical",
    signature: "average_range, criteria_range1, criteria1, …",
  },
  bahttextFunctionSpec = { family: "text", signature: "number" },
  baseFunctionSpec = {
    family: "math",
    signature: "number, radix, [minLength]",
  },
  besseliFunctionSpec = { family: "engineering", signature: "x, n" },
  besseljFunctionSpec = { family: "engineering", signature: "x, n" },
  besselkFunctionSpec = { family: "engineering", signature: "x, n" },
  besselyFunctionSpec = { family: "engineering", signature: "x, n" },
  bin2decFunctionSpec = { family: "engineering", signature: "number" },
  bin2hexFunctionSpec = {
    family: "engineering",
    signature: "number, [places]",
  },
  bin2octFunctionSpec = {
    family: "engineering",
    signature: "number, [places]",
  },
  bitandFunctionSpec = { family: "engineering", signature: "number1, number2" },
  bitlshiftFunctionSpec = {
    family: "engineering",
    signature: "number, shiftAmount",
  },
  bitorFunctionSpec = { family: "engineering", signature: "number1, number2" },
  bitrshiftFunctionSpec = {
    family: "engineering",
    signature: "number, shiftAmount",
  },
  bitxorFunctionSpec = { family: "engineering", signature: "number1, number2" },
  ceilingFunctionSpec = { family: "math", signature: "number, significance" },
  charFunctionSpec = { family: "text", signature: "number" },
  chooseFunctionSpec = {
    family: "reference",
    signature: "index_num, value1, [value2], …",
  },
  choosecolsFunctionSpec = { family: "local", signature: "array, col_num1, …" },
  chooserowsFunctionSpec = { family: "local", signature: "array, row_num1, …" },
  cleanFunctionSpec = { family: "text", signature: "text" },
  codeFunctionSpec = { family: "text", signature: "text" },
  columnFunctionSpec = { family: "reference", signature: "obj" },
  columnsFunctionSpec = { family: "reference", signature: "obj" },
  combinFunctionSpec = { family: "math", signature: "number, numberChosen" },
  combinaFunctionSpec = { family: "math", signature: "number, numberChosen" },
  complexFunctionSpec = {
    family: "engineering",
    signature: "realNum, iNum, [suffix]",
  },
  concatFunctionSpec = { family: "text", signature: "text1, [text2], …" },
  concatenateFunctionSpec = {
    family: "text",
    signature: "params1, [params2], …",
  },
  convertFunctionSpec = {
    family: "local",
    signature: "number, from_unit, to_unit",
  },
  correlFunctionSpec = {
    family: "distribution",
    signature: "[array1], [array2]",
  },
  cosFunctionSpec = { family: "local", signature: "number" },
  coshFunctionSpec = { family: "local", signature: "number" },
  cotFunctionSpec = { family: "local", signature: "number" },
  cothFunctionSpec = { family: "local", signature: "number" },
  countFunctionSpec = { family: "statistical", signature: "number" },
  countaFunctionSpec = { family: "local", signature: "number" },
  countblankFunctionSpec = { family: "local", signature: "range" },
  countifFunctionSpec = { family: "statistical", signature: "range, criteria" },
  countifsFunctionSpec = {
    family: "local",
    signature: "criteria_range1, criteria1, …",
  },
  cscFunctionSpec = { family: "local", signature: "number" },
  cschFunctionSpec = { family: "local", signature: "number" },
  dateFunctionSpec = { family: "date", signature: "year, month, day" },
  datedifFunctionSpec = {
    family: "date",
    signature: "startDate, endDate, unit",
  },
  datevalueFunctionSpec = { family: "date", signature: "dateText" },
  daverageFunctionSpec = {
    family: "local",
    signature: "database, field, criteria",
  },
  dayFunctionSpec = { family: "local", signature: "serial_number" },
  daysFunctionSpec = { family: "date", signature: "endDate, startDate" },
  days360FunctionSpec = {
    family: "date",
    signature: "startDate, endDate, [method]",
  },
  dbFunctionSpec = {
    family: "financial",
    signature: "cost, salvage, life, period, [month]",
  },
  dbcsFunctionSpec = { family: "text", signature: "text" },
  dcountFunctionSpec = {
    family: "local",
    signature: "database, [field], criteria",
  },
  dcountaFunctionSpec = {
    family: "local",
    signature: "database, [field], criteria",
  },
  ddbFunctionSpec = {
    family: "local",
    signature: "cost, salvage, life, period, [factor]",
  },
  dec2binFunctionSpec = {
    family: "engineering",
    signature: "number, [places]",
  },
  dec2hexFunctionSpec = {
    family: "engineering",
    signature: "number, [places]",
  },
  dec2octFunctionSpec = {
    family: "engineering",
    signature: "number, [places]",
  },
  decimalFunctionSpec = { family: "math", signature: "text, radix" },
  degreesFunctionSpec = { family: "math", signature: "radians" },
  deltaFunctionSpec = {
    family: "engineering",
    signature: "number1, [number2]",
  },
  devsqFunctionSpec = {
    family: "distribution",
    signature: "numbers1, [numbers2], …",
  },
  dgetFunctionSpec = {
    family: "local",
    signature: "database, field, criteria",
  },
  dmaxFunctionSpec = {
    family: "local",
    signature: "database, field, criteria",
  },
  dminFunctionSpec = {
    family: "local",
    signature: "database, field, criteria",
  },
  dollarFunctionSpec = { family: "text", signature: "number, [decimals]" },
  dproductFunctionSpec = {
    family: "local",
    signature: "database, field, criteria",
  },
  dropFunctionSpec = { family: "local", signature: "array, rows, [columns]" },
  dstdevFunctionSpec = {
    family: "local",
    signature: "database, field, criteria",
  },
  dstdevpFunctionSpec = {
    family: "local",
    signature: "database, field, criteria",
  },
  dsumFunctionSpec = {
    family: "local",
    signature: "database, field, criteria",
  },
  dvarFunctionSpec = {
    family: "local",
    signature: "database, field, criteria",
  },
  dvarpFunctionSpec = {
    family: "local",
    signature: "database, field, criteria",
  },
  edateFunctionSpec = { family: "date", signature: "startDate, months" },
  encodeurlFunctionSpec = { family: "local", signature: "text" },
  eomonthFunctionSpec = { family: "date", signature: "startDate, months" },
  erfFunctionSpec = {
    family: "engineering",
    signature: "lowerLimit, [upperLimit]",
  },
  erfcFunctionSpec = { family: "engineering", signature: "x" },
  evenFunctionSpec = { family: "math", signature: "number" },
  exactFunctionSpec = { family: "text", signature: "text1, text2" },
  expFunctionSpec = { family: "math", signature: "number" },
  expandFunctionSpec = {
    family: "local",
    signature: "array, rows, [columns], [pad_with]",
  },
  factFunctionSpec = { family: "math", signature: "number" },
  factdoubleFunctionSpec = { family: "math", signature: "number" },
  falseFunctionSpec = { family: "logical", signature: "" },
  filterFunctionSpec = {
    family: "local",
    signature: "array, include, [if_empty]",
  },
  findFunctionSpec = {
    family: "text",
    signature: "findText, withinText, [startNum]",
  },
  findbFunctionSpec = { family: "text", signature: "params1, [params2], …" },
  fisherFunctionSpec = { family: "distribution", signature: "x" },
  fisherinvFunctionSpec = { family: "distribution", signature: "x" },
  fixedFunctionSpec = {
    family: "text",
    signature: "number, [decimals], [noCommas]",
  },
  floorFunctionSpec = { family: "math", signature: "number, significance" },
  forecastFunctionSpec = {
    family: "distribution",
    signature: "x, [knownYs], [knownXs]",
  },
  formulatextFunctionSpec = { family: "local", signature: "reference" },
  frequencyFunctionSpec = {
    family: "distribution",
    signature: "[dataArray], [binsArray]",
  },
  fvFunctionSpec = {
    family: "local",
    signature: "rate, nper, pmt, [pv], [type]",
  },
  gammaFunctionSpec = { family: "distribution", signature: "x" },
  gammalnFunctionSpec = { family: "distribution", signature: "x" },
  gaussFunctionSpec = { family: "distribution", signature: "z" },
  gcdFunctionSpec = { family: "math", signature: "params1, [params2], …" },
  geomeanFunctionSpec = {
    family: "distribution",
    signature: "numbers1, [numbers2], …",
  },
  gestepFunctionSpec = { family: "engineering", signature: "number, [step]" },
  growthFunctionSpec = {
    family: "regression",
    signature: "known_ys, [known_xs], [new_xs], [const]",
  },
  harmeanFunctionSpec = {
    family: "distribution",
    signature: "numbers1, [numbers2], …",
  },
  hex2binFunctionSpec = {
    family: "engineering",
    signature: "number, [places]",
  },
  hex2decFunctionSpec = { family: "engineering", signature: "number" },
  hex2octFunctionSpec = { family: "engineering", signature: "number, places" },
  hlookupFunctionSpec = {
    family: "reference",
    signature: "lookup_value, table_array, row_index_num, [range_lookup]",
  },
  hourFunctionSpec = { family: "local", signature: "serial_number" },
  hstackFunctionSpec = { family: "local", signature: "array1, [array2], …" },
  ifFunctionSpec = {
    family: "logical",
    signature: "logical_test, value_if_true, [value_if_false]",
  },
  iferrorFunctionSpec = {
    family: "logical",
    signature: "value, value_if_error",
  },
  ifnaFunctionSpec = { family: "local", signature: "value, value_if_na" },
  ifsFunctionSpec = {
    family: "logical",
    signature: "logical_test1, value_if_true1, …",
  },
  imabsFunctionSpec = { family: "engineering", signature: "iNumber" },
  imaginaryFunctionSpec = { family: "engineering", signature: "iNumber" },
  imargumentFunctionSpec = { family: "engineering", signature: "iNumber" },
  imconjugateFunctionSpec = { family: "engineering", signature: "iNumber" },
  imcosFunctionSpec = { family: "engineering", signature: "iNumber" },
  imcoshFunctionSpec = { family: "engineering", signature: "iNumber" },
  imcotFunctionSpec = { family: "engineering", signature: "iNumber" },
  imcscFunctionSpec = { family: "engineering", signature: "iNumber" },
  imcschFunctionSpec = { family: "engineering", signature: "iNumber" },
  imdivFunctionSpec = {
    family: "engineering",
    signature: "iNumber1, iNumber2",
  },
  imexpFunctionSpec = { family: "engineering", signature: "iNumber" },
  imlnFunctionSpec = { family: "engineering", signature: "iNumber" },
  imlog10FunctionSpec = { family: "engineering", signature: "iNumber" },
  imlog2FunctionSpec = { family: "engineering", signature: "iNumber" },
  impowerFunctionSpec = { family: "engineering", signature: "iNumber, number" },
  improductFunctionSpec = {
    family: "engineering",
    signature: "params1, [params2], …",
  },
  imrealFunctionSpec = { family: "engineering", signature: "iNumber" },
  imsecFunctionSpec = { family: "engineering", signature: "iNumber" },
  imsechFunctionSpec = { family: "engineering", signature: "iNumber" },
  imsinFunctionSpec = { family: "engineering", signature: "iNumber" },
  imsinhFunctionSpec = { family: "engineering", signature: "iNumber" },
  imsqrtFunctionSpec = { family: "engineering", signature: "iNumber" },
  imsubFunctionSpec = {
    family: "engineering",
    signature: "iNumber1, iNumber2",
  },
  imsumFunctionSpec = {
    family: "engineering",
    signature: "params1, [params2], …",
  },
  imtanFunctionSpec = { family: "engineering", signature: "iNumber" },
  indexFunctionSpec = {
    family: "reference",
    signature: "array, row_num, [column_num]",
  },
  intFunctionSpec = { family: "math", signature: "number" },
  interceptFunctionSpec = {
    family: "distribution",
    signature: "[knownYs], [knownXs]",
  },
  ipmtFunctionSpec = {
    family: "financial",
    signature: "rate, per, nper, pv, [fv], [type]",
  },
  irrFunctionSpec = { family: "local", signature: "values, [guess]" },
  isblankFunctionSpec = { family: "information", signature: "value" },
  iserrFunctionSpec = { family: "information", signature: "value" },
  iserrorFunctionSpec = { family: "information", signature: "value" },
  isevenFunctionSpec = { family: "local", signature: "number" },
  isformulaFunctionSpec = { family: "local", signature: "reference" },
  islogicalFunctionSpec = { family: "information", signature: "value" },
  isnaFunctionSpec = { family: "information", signature: "value" },
  isnontextFunctionSpec = { family: "information", signature: "value" },
  isnumberFunctionSpec = { family: "information", signature: "value" },
  isoweeknumFunctionSpec = { family: "date", signature: "serialOrString" },
  isrefFunctionSpec = { family: "information", signature: "value" },
  istextFunctionSpec = { family: "information", signature: "value" },
  kurtFunctionSpec = {
    family: "distribution",
    signature: "numbers1, [numbers2], …",
  },
  lambdaFunctionSpec = {
    family: "local",
    signature: "[parameter1, parameter2, …], calculation",
  },
  largeFunctionSpec = { family: "statistical", signature: "array, k" },
  lcmFunctionSpec = { family: "math", signature: "params1, [params2], …" },
  leftFunctionSpec = { family: "text", signature: "text, [num_chars]" },
  leftbFunctionSpec = { family: "text", signature: "params1, [params2], …" },
  lenFunctionSpec = { family: "text", signature: "text" },
  letFunctionSpec = {
    family: "local",
    signature: "name1, value1, [name2, value2], …, calculation",
  },
  linestFunctionSpec = {
    family: "regression",
    signature: "known_ys, [known_xs], [const], [stats]",
  },
  lnFunctionSpec = { family: "local", signature: "number" },
  logFunctionSpec = { family: "math", signature: "number, [base]" },
  log10FunctionSpec = { family: "local", signature: "number" },
  logestFunctionSpec = {
    family: "regression",
    signature: "known_ys, [known_xs], [const], [stats]",
  },
  lookupFunctionSpec = {
    family: "local",
    signature: "lookup_value, lookup_vector, [result_vector]",
  },
  lowerFunctionSpec = { family: "text", signature: "text" },
  matchFunctionSpec = {
    family: "reference",
    signature: "lookup_value, lookup_array, [match_type]",
  },
  maxFunctionSpec = {
    family: "statistical",
    signature: "number1, [number2], …",
  },
  maxaFunctionSpec = {
    family: "statistical",
    signature: "value1, [value2, ...]",
  },
  maxifsFunctionSpec = {
    family: "statistical",
    signature: "max_range, criteria_range1, criteria1, …",
  },
  mdetermFunctionSpec = { family: "math", signature: "array" },
  medianFunctionSpec = {
    family: "statistical",
    signature: "number1, [number2], …",
  },
  midFunctionSpec = { family: "text", signature: "text, start_num, num_chars" },
  midbFunctionSpec = { family: "text", signature: "params1, [params2], …" },
  minFunctionSpec = {
    family: "statistical",
    signature: "number1, [number2], …",
  },
  minaFunctionSpec = {
    family: "statistical",
    signature: "value1, [value2, ...]",
  },
  minifsFunctionSpec = {
    family: "statistical",
    signature: "min_range, criteria_range1, criteria1, …",
  },
  minuteFunctionSpec = { family: "local", signature: "serial_number" },
  minverseFunctionSpec = { family: "math", signature: "array" },
  mirrFunctionSpec = {
    family: "financial",
    signature: "values, finance_rate, reinvest_rate",
  },
  mmultFunctionSpec = { family: "math", signature: "array1, array2" },
  modFunctionSpec = { family: "math", signature: "number, divisor" },
  modeFunctionSpec = { family: "local", signature: "number1, [number2], …" },
  monthFunctionSpec = { family: "local", signature: "serial_number" },
  mroundFunctionSpec = { family: "math", signature: "number, multiple" },
  multinomialFunctionSpec = {
    family: "math",
    signature: "numbers1, [numbers2], …",
  },
  munitFunctionSpec = { family: "math", signature: "dimension" },
  nFunctionSpec = { family: "local", signature: "value" },
  naFunctionSpec = { family: "information", signature: "" },
  networkdaysFunctionSpec = {
    family: "date",
    signature: "startDate, endDate, holidays",
  },
  notFunctionSpec = { family: "logical", signature: "logical" },
  nowFunctionSpec = { family: "date", signature: "" },
  nperFunctionSpec = {
    family: "local",
    signature: "rate, pmt, pv, [fv], [type]",
  },
  npvFunctionSpec = { family: "local", signature: "rate, value1, [value2], …" },
  numbervalueFunctionSpec = {
    family: "text",
    signature: "text, [decimalSeparator], [groupSeparator]",
  },
  oct2binFunctionSpec = {
    family: "engineering",
    signature: "number, [places]",
  },
  oct2decFunctionSpec = { family: "engineering", signature: "number" },
  oct2hexFunctionSpec = {
    family: "engineering",
    signature: "number, [places]",
  },
  oddFunctionSpec = { family: "math", signature: "number" },
  orFunctionSpec = { family: "logical", signature: "params1, [params2], …" },
  pearsonFunctionSpec = { family: "regression", signature: "array1, array2" },
  percentileFunctionSpec = { family: "local", signature: "array, k" },
  permutFunctionSpec = {
    family: "statistical",
    signature: "number, number_chosen",
  },
  permutationaFunctionSpec = {
    family: "statistical",
    signature: "number, number_chosen",
  },
  phiFunctionSpec = { family: "distribution", signature: "x" },
  piFunctionSpec = { family: "math", signature: "" },
  pmtFunctionSpec = {
    family: "local",
    signature: "rate, nper, pv, [fv], [type]",
  },
  powerFunctionSpec = { family: "math", signature: "number, power" },
  ppmtFunctionSpec = {
    family: "financial",
    signature: "rate, per, nper, pv, [fv], [type]",
  },
  priceFunctionSpec = {
    family: "financial",
    signature:
      "settlement, maturity, rate, yld, redemption, frequency, [basis]",
  },
  productFunctionSpec = {
    family: "math",
    signature: "numbers1, [numbers2], …",
  },
  properFunctionSpec = { family: "text", signature: "text" },
  pvFunctionSpec = {
    family: "local",
    signature: "rate, nper, pmt, [fv], [type]",
  },
  quartileFunctionSpec = { family: "local", signature: "array, quart" },
  quotientFunctionSpec = {
    family: "math",
    signature: "numerator, denominator",
  },
  radiansFunctionSpec = { family: "math", signature: "degrees" },
  randFunctionSpec = { family: "math", signature: "" },
  randbetweenFunctionSpec = { family: "math", signature: "bottom, top" },
  rankFunctionSpec = { family: "local", signature: "number, ref, [order]" },
  rateFunctionSpec = {
    family: "local",
    signature: "nper, pmt, pv, [fv], [type], [guess]",
  },
  replaceFunctionSpec = {
    family: "text",
    signature: "old_text, start_num, num_chars, new_text",
  },
  replacebFunctionSpec = { family: "text", signature: "params1, [params2], …" },
  reptFunctionSpec = { family: "text", signature: "text, number_times" },
  rightFunctionSpec = { family: "text", signature: "text, [num_chars]" },
  rightbFunctionSpec = { family: "text", signature: "params1, [params2], …" },
  romanFunctionSpec = { family: "math", signature: "number, [form]" },
  roundFunctionSpec = { family: "math", signature: "number, num_digits" },
  rounddownFunctionSpec = { family: "math", signature: "number, num_digits" },
  roundupFunctionSpec = { family: "math", signature: "number, num_digits" },
  rowFunctionSpec = { family: "reference", signature: "obj" },
  rowsFunctionSpec = { family: "reference", signature: "obj" },
  rsqFunctionSpec = { family: "distribution", signature: "known_ys, known_xs" },
  searchFunctionSpec = {
    family: "text",
    signature: "findText, withinText, [startNum]",
  },
  searchbFunctionSpec = { family: "text", signature: "params1, [params2], …" },
  secFunctionSpec = { family: "local", signature: "number" },
  sechFunctionSpec = { family: "local", signature: "number" },
  secondFunctionSpec = { family: "date", signature: "serialOrString" },
  sequenceFunctionSpec = {
    family: "local",
    signature: "rows, [columns], [start], [step]",
  },
  seriessumFunctionSpec = {
    family: "math",
    signature: "x, n, m, coefficients",
  },
  signFunctionSpec = { family: "local", signature: "number" },
  sinFunctionSpec = { family: "local", signature: "number" },
  sinhFunctionSpec = { family: "local", signature: "number" },
  skewFunctionSpec = {
    family: "statistical",
    signature: "number1, [number2], …",
  },
  slnFunctionSpec = { family: "local", signature: "cost, salvage, life" },
  slopeFunctionSpec = {
    family: "distribution",
    signature: "known_ys, known_xs",
  },
  smallFunctionSpec = { family: "statistical", signature: "array, k" },
  sortFunctionSpec = {
    family: "local",
    signature: "array, [sort_index], [sort_order], [by_col]",
  },
  sortbyFunctionSpec = {
    family: "local",
    signature: "array, by_array1, [sort_order1], …",
  },
  sqrtFunctionSpec = { family: "local", signature: "number" },
  sqrtpiFunctionSpec = { family: "local", signature: "number" },
  standardizeFunctionSpec = {
    family: "distribution",
    signature: "x, mean, standard_dev",
  },
  stdevaFunctionSpec = {
    family: "statistical",
    signature: "value1, [value2, ...]",
  },
  stdevpaFunctionSpec = {
    family: "statistical",
    signature: "value1, [value2, ...]",
  },
  steyxFunctionSpec = {
    family: "distribution",
    signature: "known_ys, known_xs",
  },
  substituteFunctionSpec = {
    family: "text",
    signature: "text, old_text, new_text, [instance_num]",
  },
  subtotalFunctionSpec = {
    family: "math",
    signature: "function_num, ref1, [ref2], …",
  },
  sumFunctionSpec = { family: "math", signature: "number1, [number2], …" },
  sumifFunctionSpec = {
    family: "math",
    signature: "range, criteria, [sum_range]",
  },
  sumifsFunctionSpec = {
    family: "math",
    signature: "sum_range, criteria_range1, criteria1, …",
  },
  sumproductFunctionSpec = {
    family: "math",
    signature: "[array1], arrays1, [arrays2], …",
  },
  sumsqFunctionSpec = { family: "math", signature: "params1, [params2], …" },
  sumx2my2FunctionSpec = { family: "math", signature: "arrayX, arrayY" },
  sumx2py2FunctionSpec = { family: "math", signature: "arrayX, arrayY" },
  sumxmy2FunctionSpec = { family: "math", signature: "arrayX, arrayY" },
  switchFunctionSpec = {
    family: "logical",
    signature: "expression, value1, result1, …, [default]",
  },
  sydFunctionSpec = { family: "local", signature: "cost, salvage, life, per" },
  tFunctionSpec = { family: "text", signature: "value" },
  takeFunctionSpec = { family: "local", signature: "array, rows, [columns]" },
  tanFunctionSpec = { family: "local", signature: "number" },
  tanhFunctionSpec = { family: "local", signature: "number" },
  textFunctionSpec = { family: "text", signature: "value, format_text" },
  textafterFunctionSpec = {
    family: "local",
    signature:
      "text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]",
  },
  textbeforeFunctionSpec = {
    family: "local",
    signature:
      "text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]",
  },
  textjoinFunctionSpec = {
    family: "text",
    signature: "delimiter, ignore_empty, text1, …",
  },
  textsplitFunctionSpec = {
    family: "local",
    signature:
      "text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]",
  },
  timeFunctionSpec = { family: "date", signature: "hour, minute, second" },
  timevalueFunctionSpec = { family: "date", signature: "timeText" },
  tocolFunctionSpec = {
    family: "local",
    signature: "array, [ignore], [scan_by_column]",
  },
  todayFunctionSpec = { family: "date", signature: "" },
  torowFunctionSpec = {
    family: "local",
    signature: "array, [ignore], [scan_by_column]",
  },
  transposeFunctionSpec = { family: "reference", signature: "array" },
  trendFunctionSpec = {
    family: "regression",
    signature: "known_ys, [known_xs], [new_xs], [const]",
  },
  trimFunctionSpec = { family: "text", signature: "text" },
  trimmeanFunctionSpec = { family: "statistical", signature: "array, percent" },
  trueFunctionSpec = { family: "logical", signature: "" },
  truncFunctionSpec = { family: "math", signature: "number" },
  typeFunctionSpec = { family: "local", signature: "value" },
  unicharFunctionSpec = { family: "text", signature: "number" },
  unicodeFunctionSpec = { family: "text", signature: "text" },
  uniqueFunctionSpec = {
    family: "local",
    signature: "array, [by_col], [exactly_once]",
  },
  upperFunctionSpec = { family: "local", signature: "text" },
  valueFunctionSpec = { family: "local", signature: "text" },
  varaFunctionSpec = {
    family: "statistical",
    signature: "value1, [value2, ...]",
  },
  varpaFunctionSpec = {
    family: "statistical",
    signature: "value1, [value2, ...]",
  },
  vlookupFunctionSpec = {
    family: "reference",
    signature: "lookup_value, table_array, col_index_num, [range_lookup]",
  },
  vstackFunctionSpec = { family: "local", signature: "array1, [array2], …" },
  weekdayFunctionSpec = {
    family: "date",
    signature: "serialOrString, [returnType]",
  },
  weeknumFunctionSpec = {
    family: "date",
    signature: "serialOrString, [returnType]",
  },
  workdayFunctionSpec = {
    family: "date",
    signature: "startDate, days, holidays",
  },
  xirrFunctionSpec = { family: "local", signature: "values, dates, [guess]" },
  xlookupFunctionSpec = {
    family: "local",
    signature:
      "lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]",
  },
  xmatchFunctionSpec = {
    family: "local",
    signature: "lookup_value, lookup_array, [match_mode], [search_mode]",
  },
  xnpvFunctionSpec = { family: "local", signature: "rate, values, dates" },
  xorFunctionSpec = { family: "logical", signature: "params1, [params2], …" },
  yearFunctionSpec = { family: "date", signature: "serialOrString" },
  yearfracFunctionSpec = {
    family: "date",
    signature: "startDate, endDate, [basis]",
  },
  yieldFunctionSpec = {
    family: "financial",
    signature: "settlement, maturity, rate, pr, redemption, frequency, [basis]",
  },
  functionSpecsByName = {
    ABS: absFunctionSpec,
    ACCRINT: accrintFunctionSpec,
    ACCRINTM: accrintmFunctionSpec,
    ACOS: acosFunctionSpec,
    ACOSH: acoshFunctionSpec,
    ACOT: acotFunctionSpec,
    ACOTH: acothFunctionSpec,
    ADDRESS: addressFunctionSpec,
    AGGREGATE: aggregateFunctionSpec,
    AND: andFunctionSpec,
    ARABIC: arabicFunctionSpec,
    AREAS: areasFunctionSpec,
    ASC: ascFunctionSpec,
    ASIN: asinFunctionSpec,
    ASINH: asinhFunctionSpec,
    ATAN: atanFunctionSpec,
    ATAN2: atan2FunctionSpec,
    ATANH: atanhFunctionSpec,
    AVEDEV: avedevFunctionSpec,
    AVERAGE: averageFunctionSpec,
    AVERAGEA: averageaFunctionSpec,
    AVERAGEIF: averageifFunctionSpec,
    AVERAGEIFS: averageifsFunctionSpec,
    BAHTTEXT: bahttextFunctionSpec,
    BASE: baseFunctionSpec,
    BESSELI: besseliFunctionSpec,
    BESSELJ: besseljFunctionSpec,
    BESSELK: besselkFunctionSpec,
    BESSELY: besselyFunctionSpec,
    "BETA.DIST": {
      family: "distribution",
      signature: "x, alpha, beta, cumulative, [a], [b]",
    },
    "BETA.INV": {
      family: "distribution",
      signature: "probability, alpha, beta, [a], [b]",
    },
    BIN2DEC: bin2decFunctionSpec,
    BIN2HEX: bin2hexFunctionSpec,
    BIN2OCT: bin2octFunctionSpec,
    "BINOM.DIST": {
      family: "distribution",
      signature: "numberS, trials, probabilityS, cumulative",
    },
    "BINOM.DIST.RANGE": {
      family: "distribution",
      signature: "trials, probabilityS, numberS, [numberS2]",
    },
    "BINOM.INV": {
      family: "distribution",
      signature: "trials, probabilityS, alpha",
    },
    BITAND: bitandFunctionSpec,
    BITLSHIFT: bitlshiftFunctionSpec,
    BITOR: bitorFunctionSpec,
    BITRSHIFT: bitrshiftFunctionSpec,
    BITXOR: bitxorFunctionSpec,
    CEILING: ceilingFunctionSpec,
    "CEILING.MATH": {
      family: "math",
      signature: "number, [significance], [mode]",
    },
    "CEILING.PRECISE": { family: "math", signature: "number, [significance]" },
    CHAR: charFunctionSpec,
    "CHISQ.DIST": {
      family: "distribution",
      signature: "x, degFreedom, cumulative",
    },
    "CHISQ.DIST.RT": { family: "distribution", signature: "x, degFreedom" },
    "CHISQ.INV": {
      family: "distribution",
      signature: "probability, degFreedom",
    },
    "CHISQ.INV.RT": {
      family: "distribution",
      signature: "probability, degFreedom",
    },
    "CHISQ.TEST": {
      family: "distribution",
      signature: "[actualRange], [expectedRange]",
    },
    CHOOSE: chooseFunctionSpec,
    CHOOSECOLS: choosecolsFunctionSpec,
    CHOOSEROWS: chooserowsFunctionSpec,
    CLEAN: cleanFunctionSpec,
    CODE: codeFunctionSpec,
    COLUMN: columnFunctionSpec,
    COLUMNS: columnsFunctionSpec,
    COMBIN: combinFunctionSpec,
    COMBINA: combinaFunctionSpec,
    COMPLEX: complexFunctionSpec,
    CONCAT: concatFunctionSpec,
    CONCATENATE: concatenateFunctionSpec,
    "CONFIDENCE.NORM": {
      family: "distribution",
      signature: "alpha, std, size",
    },
    "CONFIDENCE.T": { family: "distribution", signature: "alpha, std, size" },
    CONVERT: convertFunctionSpec,
    CORREL: correlFunctionSpec,
    COS: cosFunctionSpec,
    COSH: coshFunctionSpec,
    COT: cotFunctionSpec,
    COTH: cothFunctionSpec,
    COUNT: countFunctionSpec,
    COUNTA: countaFunctionSpec,
    COUNTBLANK: countblankFunctionSpec,
    COUNTIF: countifFunctionSpec,
    COUNTIFS: countifsFunctionSpec,
    "COVARIANCE.P": { family: "distribution", signature: "[array1], [array2]" },
    "COVARIANCE.S": { family: "distribution", signature: "[array1], [array2]" },
    CSC: cscFunctionSpec,
    CSCH: cschFunctionSpec,
    DATE: dateFunctionSpec,
    DATEDIF: datedifFunctionSpec,
    DATEVALUE: datevalueFunctionSpec,
    DAVERAGE: daverageFunctionSpec,
    DAY: dayFunctionSpec,
    DAYS: daysFunctionSpec,
    DAYS360: days360FunctionSpec,
    DB: dbFunctionSpec,
    DBCS: dbcsFunctionSpec,
    DCOUNT: dcountFunctionSpec,
    DCOUNTA: dcountaFunctionSpec,
    DDB: ddbFunctionSpec,
    DEC2BIN: dec2binFunctionSpec,
    DEC2HEX: dec2hexFunctionSpec,
    DEC2OCT: dec2octFunctionSpec,
    DECIMAL: decimalFunctionSpec,
    DEGREES: degreesFunctionSpec,
    DELTA: deltaFunctionSpec,
    DEVSQ: devsqFunctionSpec,
    DGET: dgetFunctionSpec,
    DMAX: dmaxFunctionSpec,
    DMIN: dminFunctionSpec,
    DOLLAR: dollarFunctionSpec,
    DPRODUCT: dproductFunctionSpec,
    DROP: dropFunctionSpec,
    DSTDEV: dstdevFunctionSpec,
    DSTDEVP: dstdevpFunctionSpec,
    DSUM: dsumFunctionSpec,
    DVAR: dvarFunctionSpec,
    DVARP: dvarpFunctionSpec,
    EDATE: edateFunctionSpec,
    ENCODEURL: encodeurlFunctionSpec,
    EOMONTH: eomonthFunctionSpec,
    ERF: erfFunctionSpec,
    ERFC: erfcFunctionSpec,
    "ERROR.TYPE": { family: "information", signature: "value" },
    EVEN: evenFunctionSpec,
    EXACT: exactFunctionSpec,
    EXP: expFunctionSpec,
    EXPAND: expandFunctionSpec,
    "EXPON.DIST": {
      family: "distribution",
      signature: "x, lambda, cumulative",
    },
    "F.DIST": { family: "distribution", signature: "x, d1, d2, cumulative" },
    "F.DIST.RT": { family: "distribution", signature: "x, d1, d2" },
    "F.INV": { family: "distribution", signature: "probability, d1, d2" },
    "F.INV.RT": { family: "distribution", signature: "probability, d1, d2" },
    "F.TEST": { family: "distribution", signature: "[array1], [array2]" },
    FACT: factFunctionSpec,
    FACTDOUBLE: factdoubleFunctionSpec,
    FALSE: falseFunctionSpec,
    FILTER: filterFunctionSpec,
    FIND: findFunctionSpec,
    FINDB: findbFunctionSpec,
    FISHER: fisherFunctionSpec,
    FISHERINV: fisherinvFunctionSpec,
    FIXED: fixedFunctionSpec,
    FLOOR: floorFunctionSpec,
    "FLOOR.MATH": {
      family: "math",
      signature: "number, [significance], [mode]",
    },
    "FLOOR.PRECISE": { family: "math", signature: "number, [significance]" },
    FORECAST: forecastFunctionSpec,
    "FORECAST.LINEAR": {
      family: "distribution",
      signature: "params1, [params2], …",
    },
    FORMULATEXT: formulatextFunctionSpec,
    FREQUENCY: frequencyFunctionSpec,
    FV: fvFunctionSpec,
    GAMMA: gammaFunctionSpec,
    "GAMMA.DIST": {
      family: "distribution",
      signature: "x, alpha, beta, cumulative",
    },
    "GAMMA.INV": {
      family: "distribution",
      signature: "probability, alpha, beta",
    },
    GAMMALN: gammalnFunctionSpec,
    "GAMMALN.PRECISE": { family: "distribution", signature: "x" },
    GAUSS: gaussFunctionSpec,
    GCD: gcdFunctionSpec,
    GEOMEAN: geomeanFunctionSpec,
    GESTEP: gestepFunctionSpec,
    GROWTH: growthFunctionSpec,
    HARMEAN: harmeanFunctionSpec,
    HEX2BIN: hex2binFunctionSpec,
    HEX2DEC: hex2decFunctionSpec,
    HEX2OCT: hex2octFunctionSpec,
    HLOOKUP: hlookupFunctionSpec,
    HOUR: hourFunctionSpec,
    HSTACK: hstackFunctionSpec,
    "HYPGEOM.DIST": {
      family: "distribution",
      signature:
        "sample_s, number_sample, population_s, number_pop, cumulative",
    },
    IF: ifFunctionSpec,
    IFERROR: iferrorFunctionSpec,
    IFNA: ifnaFunctionSpec,
    IFS: ifsFunctionSpec,
    IMABS: imabsFunctionSpec,
    IMAGINARY: imaginaryFunctionSpec,
    IMARGUMENT: imargumentFunctionSpec,
    IMCONJUGATE: imconjugateFunctionSpec,
    IMCOS: imcosFunctionSpec,
    IMCOSH: imcoshFunctionSpec,
    IMCOT: imcotFunctionSpec,
    IMCSC: imcscFunctionSpec,
    IMCSCH: imcschFunctionSpec,
    IMDIV: imdivFunctionSpec,
    IMEXP: imexpFunctionSpec,
    IMLN: imlnFunctionSpec,
    IMLOG10: imlog10FunctionSpec,
    IMLOG2: imlog2FunctionSpec,
    IMPOWER: impowerFunctionSpec,
    IMPRODUCT: improductFunctionSpec,
    IMREAL: imrealFunctionSpec,
    IMSEC: imsecFunctionSpec,
    IMSECH: imsechFunctionSpec,
    IMSIN: imsinFunctionSpec,
    IMSINH: imsinhFunctionSpec,
    IMSQRT: imsqrtFunctionSpec,
    IMSUB: imsubFunctionSpec,
    IMSUM: imsumFunctionSpec,
    IMTAN: imtanFunctionSpec,
    INDEX: indexFunctionSpec,
    INT: intFunctionSpec,
    INTERCEPT: interceptFunctionSpec,
    IPMT: ipmtFunctionSpec,
    IRR: irrFunctionSpec,
    ISBLANK: isblankFunctionSpec,
    ISERR: iserrFunctionSpec,
    ISERROR: iserrorFunctionSpec,
    ISEVEN: isevenFunctionSpec,
    ISFORMULA: isformulaFunctionSpec,
    ISLOGICAL: islogicalFunctionSpec,
    ISNA: isnaFunctionSpec,
    ISNONTEXT: isnontextFunctionSpec,
    ISNUMBER: isnumberFunctionSpec,
    "ISO.CEILING": { family: "math", signature: "params1, [params2], …" },
    ISOWEEKNUM: isoweeknumFunctionSpec,
    ISREF: isrefFunctionSpec,
    ISTEXT: istextFunctionSpec,
    KURT: kurtFunctionSpec,
    LAMBDA: lambdaFunctionSpec,
    LARGE: largeFunctionSpec,
    LCM: lcmFunctionSpec,
    LEFT: leftFunctionSpec,
    LEFTB: leftbFunctionSpec,
    LEN: lenFunctionSpec,
    LET: letFunctionSpec,
    LINEST: linestFunctionSpec,
    LN: lnFunctionSpec,
    LOG: logFunctionSpec,
    LOG10: log10FunctionSpec,
    LOGEST: logestFunctionSpec,
    "LOGNORM.DIST": {
      family: "distribution",
      signature: "x, mean, standard_dev, cumulative",
    },
    "LOGNORM.INV": {
      family: "distribution",
      signature: "probability, mean, standard_dev",
    },
    LOOKUP: lookupFunctionSpec,
    LOWER: lowerFunctionSpec,
    MATCH: matchFunctionSpec,
    MAX: maxFunctionSpec,
    MAXA: maxaFunctionSpec,
    MAXIFS: maxifsFunctionSpec,
    MDETERM: mdetermFunctionSpec,
    MEDIAN: medianFunctionSpec,
    MID: midFunctionSpec,
    MIDB: midbFunctionSpec,
    MIN: minFunctionSpec,
    MINA: minaFunctionSpec,
    MINIFS: minifsFunctionSpec,
    MINUTE: minuteFunctionSpec,
    MINVERSE: minverseFunctionSpec,
    MIRR: mirrFunctionSpec,
    MMULT: mmultFunctionSpec,
    MOD: modFunctionSpec,
    MODE: modeFunctionSpec,
    "MODE.MULT": { family: "statistical", signature: "number1, [number2], …" },
    "MODE.SNGL": { family: "distribution", signature: "number1, [number2], …" },
    MONTH: monthFunctionSpec,
    MROUND: mroundFunctionSpec,
    MULTINOMIAL: multinomialFunctionSpec,
    MUNIT: munitFunctionSpec,
    N: nFunctionSpec,
    NA: naFunctionSpec,
    "NEGBINOM.DIST": {
      family: "distribution",
      signature: "number_f, number_s, probability_s, cumulative",
    },
    NETWORKDAYS: networkdaysFunctionSpec,
    "NETWORKDAYS.INTL": {
      family: "date",
      signature: "startDate, endDate, weekend, holidays",
    },
    "NORM.DIST": {
      family: "distribution",
      signature: "x, mean, standard_dev, cumulative",
    },
    "NORM.INV": {
      family: "distribution",
      signature: "probability, mean, standard_dev",
    },
    "NORM.S.DIST": { family: "distribution", signature: "z, cumulative" },
    "NORM.S.INV": { family: "distribution", signature: "probability" },
    NOT: notFunctionSpec,
    NOW: nowFunctionSpec,
    NPER: nperFunctionSpec,
    NPV: npvFunctionSpec,
    NUMBERVALUE: numbervalueFunctionSpec,
    OCT2BIN: oct2binFunctionSpec,
    OCT2DEC: oct2decFunctionSpec,
    OCT2HEX: oct2hexFunctionSpec,
    ODD: oddFunctionSpec,
    OR: orFunctionSpec,
    PEARSON: pearsonFunctionSpec,
    PERCENTILE: percentileFunctionSpec,
    "PERCENTILE.EXC": { family: "distribution", signature: "array, k" },
    "PERCENTILE.INC": { family: "distribution", signature: "array, k" },
    PERMUT: permutFunctionSpec,
    PERMUTATIONA: permutationaFunctionSpec,
    PHI: phiFunctionSpec,
    PI: piFunctionSpec,
    PMT: pmtFunctionSpec,
    "POISSON.DIST": {
      family: "distribution",
      signature: "x, mean, cumulative",
    },
    POWER: powerFunctionSpec,
    PPMT: ppmtFunctionSpec,
    PRICE: priceFunctionSpec,
    PRODUCT: productFunctionSpec,
    PROPER: properFunctionSpec,
    PV: pvFunctionSpec,
    QUARTILE: quartileFunctionSpec,
    "QUARTILE.EXC": { family: "distribution", signature: "array, quart" },
    "QUARTILE.INC": { family: "distribution", signature: "array, quart" },
    QUOTIENT: quotientFunctionSpec,
    RADIANS: radiansFunctionSpec,
    RAND: randFunctionSpec,
    RANDBETWEEN: randbetweenFunctionSpec,
    RANK: rankFunctionSpec,
    "RANK.AVG": { family: "distribution", signature: "number, ref, [order]" },
    "RANK.EQ": { family: "distribution", signature: "number, ref, [order]" },
    RATE: rateFunctionSpec,
    REPLACE: replaceFunctionSpec,
    REPLACEB: replacebFunctionSpec,
    REPT: reptFunctionSpec,
    RIGHT: rightFunctionSpec,
    RIGHTB: rightbFunctionSpec,
    ROMAN: romanFunctionSpec,
    ROUND: roundFunctionSpec,
    ROUNDDOWN: rounddownFunctionSpec,
    ROUNDUP: roundupFunctionSpec,
    ROW: rowFunctionSpec,
    ROWS: rowsFunctionSpec,
    RSQ: rsqFunctionSpec,
    SEARCH: searchFunctionSpec,
    SEARCHB: searchbFunctionSpec,
    SEC: secFunctionSpec,
    SECH: sechFunctionSpec,
    SECOND: secondFunctionSpec,
    SEQUENCE: sequenceFunctionSpec,
    SERIESSUM: seriessumFunctionSpec,
    SIGN: signFunctionSpec,
    SIN: sinFunctionSpec,
    SINH: sinhFunctionSpec,
    SKEW: skewFunctionSpec,
    "SKEW.P": { family: "statistical", signature: "number1, [number2], …" },
    SLN: slnFunctionSpec,
    SLOPE: slopeFunctionSpec,
    SMALL: smallFunctionSpec,
    SORT: sortFunctionSpec,
    SORTBY: sortbyFunctionSpec,
    SQRT: sqrtFunctionSpec,
    SQRTPI: sqrtpiFunctionSpec,
    STANDARDIZE: standardizeFunctionSpec,
    "STDEV.P": { family: "distribution", signature: "number1, [number2], …" },
    "STDEV.S": { family: "distribution", signature: "number1, [number2], …" },
    STDEVA: stdevaFunctionSpec,
    STDEVPA: stdevpaFunctionSpec,
    STEYX: steyxFunctionSpec,
    SUBSTITUTE: substituteFunctionSpec,
    SUBTOTAL: subtotalFunctionSpec,
    SUM: sumFunctionSpec,
    SUMIF: sumifFunctionSpec,
    SUMIFS: sumifsFunctionSpec,
    SUMPRODUCT: sumproductFunctionSpec,
    SUMSQ: sumsqFunctionSpec,
    SUMX2MY2: sumx2my2FunctionSpec,
    SUMX2PY2: sumx2py2FunctionSpec,
    SUMXMY2: sumxmy2FunctionSpec,
    SWITCH: switchFunctionSpec,
    SYD: sydFunctionSpec,
    T: tFunctionSpec,
    "T.DIST": {
      family: "distribution",
      signature: "x, deg_freedom, cumulative",
    },
    "T.DIST.2T": { family: "distribution", signature: "x, deg_freedom" },
    "T.DIST.RT": { family: "distribution", signature: "x, deg_freedom" },
    "T.INV": { family: "distribution", signature: "probability, deg_freedom" },
    "T.INV.2T": {
      family: "distribution",
      signature: "probability, deg_freedom",
    },
    "T.TEST": {
      family: "distribution",
      signature: "array1, array2, tails, type",
    },
    TAKE: takeFunctionSpec,
    TAN: tanFunctionSpec,
    TANH: tanhFunctionSpec,
    TEXT: textFunctionSpec,
    TEXTAFTER: textafterFunctionSpec,
    TEXTBEFORE: textbeforeFunctionSpec,
    TEXTJOIN: textjoinFunctionSpec,
    TEXTSPLIT: textsplitFunctionSpec,
    TIME: timeFunctionSpec,
    TIMEVALUE: timevalueFunctionSpec,
    TOCOL: tocolFunctionSpec,
    TODAY: todayFunctionSpec,
    TOROW: torowFunctionSpec,
    TRANSPOSE: transposeFunctionSpec,
    TREND: trendFunctionSpec,
    TRIM: trimFunctionSpec,
    TRIMMEAN: trimmeanFunctionSpec,
    TRUE: trueFunctionSpec,
    TRUNC: truncFunctionSpec,
    TYPE: typeFunctionSpec,
    UNICHAR: unicharFunctionSpec,
    UNICODE: unicodeFunctionSpec,
    UNIQUE: uniqueFunctionSpec,
    UPPER: upperFunctionSpec,
    VALUE: valueFunctionSpec,
    "VAR.P": { family: "distribution", signature: "number1, [number2], …" },
    "VAR.S": { family: "distribution", signature: "number1, [number2], …" },
    VARA: varaFunctionSpec,
    VARPA: varpaFunctionSpec,
    VLOOKUP: vlookupFunctionSpec,
    VSTACK: vstackFunctionSpec,
    WEEKDAY: weekdayFunctionSpec,
    WEEKNUM: weeknumFunctionSpec,
    "WEIBULL.DIST": {
      family: "distribution",
      signature: "x, alpha, beta, cumulative",
    },
    WORKDAY: workdayFunctionSpec,
    "WORKDAY.INTL": {
      family: "date",
      signature: "startDate, days, weekend, holidays",
    },
    XIRR: xirrFunctionSpec,
    XLOOKUP: xlookupFunctionSpec,
    XMATCH: xmatchFunctionSpec,
    XNPV: xnpvFunctionSpec,
    XOR: xorFunctionSpec,
    YEAR: yearFunctionSpec,
    YEARFRAC: yearfracFunctionSpec,
    YIELD: yieldFunctionSpec,
    "Z.TEST": { family: "distribution", signature: "array, x, [sigma]" },
  },
  localFunctionNameSet = new Set(builtinFunctionNames),
  arrayIncludesValue = (value, values) => values.includes(value),
  functionBrowserCategories = [
    {
      id: "database",
      title: "Database summaries",
      demand: "Summarize headed audit populations using explicit criteria.",
    },
    {
      id: "distribution",
      title: "Distributions",
      demand:
        "Sampling assumptions, probabilities and confidence calculations; auditor selects the appropriate model.",
    },
    {
      id: "regression",
      title: "Regression",
      demand: "Analytical review of paired observations and estimates.",
    },
    {
      id: "statistics",
      title: "Descriptive statistics",
      demand: "Population profiles, dispersion, ranking and outlier review.",
    },
    {
      id: "engineering",
      title: "Engineering and conversions",
      demand:
        "Reconcile inventory quantities, measurement units and specialist calculations.",
    },
    {
      id: "matrix",
      title: "Matrix operations",
      demand: "Review bounded allocation models and systems of equations.",
    },
    {
      id: "financial",
      title: "Financial models",
      demand:
        "Discount cash flows, annuities and depreciation; financial functions are not an actuarial model.",
    },
    {
      id: "cube",
      title: "CUBE / OLAP",
      demand:
        "Model queries require a separately approved host connector; no connector is implemented.",
    },
    {
      id: "network",
      title: "Web and linked data",
      demand:
        "External retrieval requires host authorization and an implemented connector; disabled locally.",
    },
    {
      id: "actuarial",
      title: "Actuarial and add-ins",
      demand:
        "Engagement-specific mortality, survival and reserving models require independently validated custom models; no dedicated actuarial engine is adopted.",
    },
    {
      id: "other",
      title: "Other workbook functions",
      demand:
        "General calculation, text, date, reference and reusable formulas.",
    },
  ],
  regressionFunctionNames = [
    "SLOPE",
    "RSQ",
    "STEYX",
    "CORREL",
    "INTERCEPT",
    "FORECAST",
    "FORECAST.LINEAR",
    "GROWTH",
    "LINEST",
    "LOGEST",
    "TREND",
    "PEARSON",
  ],
  matrixFunctionNames = ["MINVERSE", "MDETERM", "MMULT", "MUNIT"],
  financialFunctionNames = [
    "IPMT",
    "PPMT",
    "DB",
    "PRICE",
    "YIELD",
    "ACCRINT",
    "ACCRINTM",
    "MIRR",
    "PMT",
    "PV",
    "FV",
    "NPER",
    "RATE",
    "NPV",
    "IRR",
    "XNPV",
    "XIRR",
    "SLN",
    "SYD",
    "DDB",
  ],
  connectorFunctionFamilies = Object.freeze({
    CUBEVALUE: "cube",
    CUBEMEMBER: "cube",
    CUBEMEMBERPROPERTY: "cube",
    CUBERANKEDMEMBER: "cube",
    CUBESET: "cube",
    CUBESETCOUNT: "cube",
    CUBEKPIMEMBER: "cube",
    WEBSERVICE: "network",
    FILTERXML: "network",
    STOCKHISTORY: "network",
    RTD: "network",
    IMAGE: "network",
    FIELDVALUE: "network",
    HYPERLINK: "network",
    "ERF.PRECISE": "engineering",
    "ERFC.PRECISE": "engineering",
    DURATION: "financial",
    MDURATION: "financial",
    CALL: "actuarial",
    "REGISTER.ID": "actuarial",
  });
function classifyFunctionFamily(functionName) {
  if (
    ((functionName = functionName.toUpperCase()),
    arrayIncludesValue(functionName, databaseFunctionNames))
  )
    return "database";
  if (regressionFunctionNames.includes(functionName)) return "regression";
  if (matrixFunctionNames.includes(functionName)) return "matrix";
  if (financialFunctionNames.includes(functionName)) return "financial";
  if (functionName === "CONVERT") return "engineering";
  if (arrayIncludesValue(functionName, statisticalFunctionNames))
    return "statistics";
  if (connectorFunctionFamilies[functionName])
    return connectorFunctionFamilies[functionName];
  if (functionName.startsWith("CUBE")) return "cube";
  const specFamily = functionSpecsByName[functionName]?.family;
  return specFamily === "statistical"
    ? "statistics"
    : specFamily === "distribution" ||
        specFamily === "engineering" ||
        specFamily === "financial"
      ? specFamily
      : "other";
}
const familyResultContracts = {
  database: {
    output: "Selected-field aggregate or single DGET value",
    limits:
      "20,000 database cells; 1,000 criteria cells; 100,000 candidate comparisons; 10,000 per-record formula evaluations. Unknown/blank criteria headings require actual worksheet formula cells returning Boolean results. Relative rows rebase from the first record; absolute references stay fixed. Formula strings in arrays never execute.",
    errors: ["#VALUE!", "#DIV/0!", "#NUM!"],
    tests:
      "tests/database.test.ts;tests/database-boundaries.test.ts;tests/database-formula-criteria.test.ts",
  },
  distribution: {
    output:
      "Numeric probability, density, quantile or test statistic according to the function",
    limits:
      "Pinned distribution implementation; discrete work limits apply. Extreme tails and general numerical/argument parity are not verified.",
    errors: ["#VALUE!", "#NUM!", "#DIV/0!"],
    tests:
      "tests/formula-extended.test.ts;tests/statistical-specialist.test.ts",
  },
  regression: {
    output:
      "Numeric coefficient, prediction or fitted array according to the function",
    limits:
      "LINEST/LOGEST/TREND/GROWTH: rectangular numeric data, 10,000 observations, 32 predictors, 40,000 inspected values and 4,000,000 fit work units; centered pivoted QR with a relative 1e-12 rank threshold. Non-numeric regression matrices are rejected. Coefficients are in reverse predictor order; LOGEST statistics use log-y space. Paired SLOPE/RSQ/STEYX/PEARSON omit nonnumeric pairs. Exact Excel rank selection, extreme conditioning and coercion parity remain unverified.",
    errors: ["#VALUE!", "#N/A", "#REF!", "#DIV/0!", "#NUM!", "#SPILL!"],
    tests:
      "tests/specialist-functions.test.ts;tests/regression-acceptance.test.ts;tests/formula-extended.test.ts",
  },
  statistics: {
    output: "Numeric summary, count, rank or distribution summary",
    limits:
      "Bounded ranges and 100,000 argument values; exact coercion varies by function. A known answer is not full argument conformance.",
    errors: ["#VALUE!", "#N/A", "#DIV/0!", "#NUM!"],
    tests:
      "tests/formula-extended.test.ts;tests/statistical-acceptance.test.ts;tests/statistical-specialist.test.ts",
  },
  engineering: {
    output:
      "Number or text (including base notation and complex numbers), according to the function",
    limits:
      "CONVERT supports only its exported unit inventory and approved case-sensitive prefixes. Bessel order is limited to 1,000; generated text is bounded.",
    errors: ["#VALUE!", "#N/A", "#NUM!"],
    tests: "tests/specialist-functions.test.ts;tests/formula-extended.test.ts",
  },
  matrix: {
    output: "MDETERM: scalar determinant; MINVERSE/MMULT/MUNIT: numeric matrix",
    limits:
      "MINVERSE/MDETERM: square numeric matrices up to 32×32; MMULT: at most 1,000 argument values; MUNIT: dimension up to 100. Spill and result limits apply. Ill-conditioned numerical parity is not verified.",
    errors: ["#VALUE!", "#NUM!", "#SPILL!"],
    tests: "tests/specialist-functions.test.ts;tests/formula-extended.test.ts",
  },
  financial: {
    output:
      "Numeric financial result; units/sign convention follow the named function",
    limits:
      "10000 cash-flow positions/financial periods; 1000 coupon periods. Rates are per input interval. Dated returns use actual/365 Gregorian dates in the workbook's 1900/1904 system with one numeric amount per date. Residual-checked IRR/XIRR/RATE/YIELD search is bounded to 2048 evaluations and rates -0.999999999 through 1000000 (YIELD nonnegative); root discovery is not exhaustive. Regular security coupons and documented day-count conventions only. No actuarial or full Excel parity claim.",
    errors: ["#VALUE!", "#NUM!", "#DIV/0!"],
    tests:
      "tests/financial-functions.test.ts;tests/financial-schedules.test.ts;tests/formula-extended.test.ts",
  },
  cube: {
    output: "No local result",
    limits:
      "No CUBE connector is installed. Workbook text or a claimed approval cannot enable a connection.",
    errors: ["#NAME?", "#BLOCKED!"],
    tests: "tests/function-compatibility.test.ts",
  },
  network: {
    output: "No local result",
    limits:
      "No formula-driven network or linked-data connector is installed. Host connector approval and implementation are prerequisites, never workbook flags.",
    errors: ["#NAME?", "#BLOCKED!"],
    tests: "tests/function-compatibility.test.ts",
  },
  actuarial: {
    output: "No local result",
    limits:
      "No arbitrary add-in/executable model is run. Independently validated actuarial models and approved host integration are required.",
    errors: ["#NAME?", "#BLOCKED!"],
    tests: "tests/function-compatibility.test.ts",
  },
  other: {
    output: "Scalar, text, logical or array result according to the function",
    limits:
      "4,096 formula characters, 20,000 cells per reference, 100,000 dependency reads, 10,000 cells per result; function-specific documented limits also apply.",
    errors: [
      "#VALUE!",
      "#N/A",
      "#NUM!",
      "#DIV/0!",
      "#REF!",
      "#SPILL!",
      "#CALC!",
    ],
    tests: "tests/formula-extended.test.ts;tests/core.test.ts",
  },
};
function buildFunctionCatalogEntry(functionName) {
  const normalizedName = functionName
      .replace(/^_xlfn\.(?:_xlws\.)?/i, "")
      .toUpperCase(),
    family = classifyFunctionFamily(normalizedName),
    familyContract = familyResultContracts[family],
    isLocal = localFunctionNameSet.has(normalizedName);
  return {
    name: normalizedName,
    family: family,
    status: isLocal
      ? "local"
      : family === "cube" || family === "network"
        ? "connector-required"
        : "unimplemented",
    inputs:
      functionSpecsByName[normalizedName]?.signature ??
      "Not adopted; consult the source function or add-in contract.",
    outputs: isLocal
      ? (functionDescriptions[normalizedName] ?? familyContract.output)
      : "No local implementation",
    errors: [
      ...familyContract.errors,
      "#LIMIT!",
      "#CYCLE!",
      ...(isLocal ? ["#TIMEOUT!", "#BLOCKED!"] : []),
    ],
    numericTolerance: isLocal
      ? {
          absolute: 1e-6,
          relative: 1e-6,
          appliesTo:
            "Local known-answer comparison only; exact text/logical/errors. Not an accuracy guarantee across all inputs.",
        }
      : null,
    excelVerification: { status: "not-run", version: null, evidence: null },
    limits:
      (arrayIncludesValue(normalizedName, dateFunctionNames)
        ? "Gregorian 1900 serials 0–2958465 (compatibility serial 60) or 1904 serials 0–2957003; full-year ISO/configured MDY or DMY text; recorded IANA-zone clock shared per calculation. Up to 1000 holiday inputs, 366000 business-day offsets/date spans. The date system persists in native JSON and XLSX; text order and time zone are native JSON only. DATEDIF MD/YD and US day-count edge conventions follow the documented local policy; independent Excel parity unverified."
        : familyContract.limits) +
      (isLocal
        ? " Workbook work and accounted-memory budgets apply across a run; synchronous APIs use cooperative time checks and isolated worker execution supports deadline termination/cancellation."
        : "") +
      (["RAND", "RANDBETWEEN"].includes(normalizedName)
        ? " Unseeded runtime pseudorandom values are stable within the revision calculation and refresh on recalculation. They are not reproducible audit samples. RANDBETWEEN uses ceil(bottom)/floor(top) and refuses bounds or inclusive spans outside safe integers."
        : ""),
    evidence: isLocal
      ? arrayIncludesValue(normalizedName, dateFunctionNames)
        ? [
            "tests/date-calendar.test.ts",
            "tests/date-calendar-integration.test.ts",
            "docs/DATE-CALENDAR-ACCEPTANCE.md",
          ]
        : familyContract.tests.split(";")
      : [],
  };
}
function buildFunctionCatalog() {
  return [
    ...new Set([
      ...builtinFunctionNames,
      ...Object.keys(connectorFunctionFamilies),
    ]),
  ]
    .sort()
    .map(buildFunctionCatalogEntry);
}
function buildFunctionBrowserGroups() {
  const functionCatalog = buildFunctionCatalog();
  return functionBrowserCategories.map((functionBrowserCategory) => ({
    ...functionBrowserCategory,
    adopted: functionCatalog
      .filter(
        (adoptedCatalogEntry) =>
          adoptedCatalogEntry.family === functionBrowserCategory.id &&
          adoptedCatalogEntry.status === "local",
      )
      .map((adoptedEntry) => adoptedEntry.name),
    unavailable: functionCatalog
      .filter(
        (unavailableCatalogEntry) =>
          unavailableCatalogEntry.family === functionBrowserCategory.id &&
          unavailableCatalogEntry.status !== "local",
      )
      .map((unavailableEntry) => unavailableEntry.name),
    excelVersionVerified: false,
  }));
}
function FunctionAndNameDialog({
  mode,
  workbook,
  sheetId,
  range,
  readOnly,
  onClose,
  onWorkbookChange,
  onInsertFormula,
}) {
  const { text, language } = useWorkbookMessages(),
    sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    [searchQuery, setSearchQuery] = ne.useState(""),
    [functionName, setFunctionName] = ne.useState("SUM"),
    [functionArguments, setFunctionArguments] = ne.useState(range),
    [definedName, setDefinedName] = ne.useState(""),
    [refersTo, setRefersTo] = ne.useState(
      `='${(sheet?.name ?? "Sheet1").replace(/'/g, "''")}'!${range}`,
    ),
    [scopeSheetId, setScopeSheetId] = ne.useState(""),
    [comment, setComment] = ne.useState(""),
    [editingIndex, setEditingIndex] = ne.useState(null),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [workbookSnapshot, setWorkbookSnapshot] = ne.useState(
      JSON.stringify(workbook),
    ),
    [hasSnapshot, setHasSnapshot] = ne.useState(false),
    [renamePreview, setRenamePreview] = ne.useState(null),
    [testArguments, setTestArguments] = ne.useState("100,0.2"),
    [testResult, setTestResult] = ne.useState(""),
    formulaTree = parseFormulaTree(refersTo),
    updateDraft = (setDraftField, draftValue) => {
      (hasSnapshot || setWorkbookSnapshot(JSON.stringify(workbook)),
        setHasSnapshot(true),
        setRenamePreview(null),
        setTestResult(""),
        setDraftField(draftValue));
    },
    isLocked = readOnly || !sheet || !!workbook.structureProtection?.enabled,
    reportError = (se) =>
      setErrorMessage(
        se instanceof Error ? se.message : "The name could not be saved.",
      ),
    matchingFunctionNames = ne.useMemo(
      () =>
        builtinFunctionNames.filter((builtinFunctionName) =>
          builtinFunctionName.toLowerCase().includes(searchQuery.toLowerCase()),
        ),
      [searchQuery],
    ),
    raw = `=${functionName}(${functionArguments})`,
    functionCatalogEntry = buildFunctionCatalogEntry(functionName),
    previewResult = ne.useMemo(() => {
      if (mode !== "function") return "";
      const anchorAddress = range
          .split(":")[0]
          .replace(/\$/g, "")
          .toUpperCase(),
        previewWorkbook = {
          ...workbook,
          sheets: workbook.sheets.map((sheet) =>
            sheet.id === sheetId
              ? {
                  ...sheet,
                  cells: { ...sheet.cells, [anchorAddress]: { raw: raw } },
                }
              : sheet,
          ),
        };
      try {
        return String(
          calculateWorkbookWithSettings(previewWorkbook)(
            sheetId,
            anchorAddress,
          ) ?? "",
        );
      } catch {
        return "Check the function arguments.";
      }
    }, [raw, mode, workbook, sheetId, range]),
    onClick = () => {
      (setDefinedName(""),
        setComment(""),
        setEditingIndex(null),
        setErrorMessage(""),
        setTestResult(""),
        setHasSnapshot(false),
        setRenamePreview(null),
        setWorkbookSnapshot(JSON.stringify(workbook)),
        setScopeSheetId(""));
    },
    saveDefinedName = () => {
      var definedNames;
      try {
        if (isLocked) throw Error("This workbook is read-only or protected.");
        if (hasSnapshot && workbookSnapshot !== JSON.stringify(workbook))
          throw Error(
            "The workbook changed while this name was being edited. Choose New name or reopen Edit to review current content.",
          );
        const definition = {
            name: definedName.trim(),
            refersTo: refersTo.trim(),
            ...(scopeSheetId ? { scopeSheetId: scopeSheetId } : {}),
            ...(comment ? { comment: comment } : {}),
          },
          existingName =
            editingIndex === null ||
            (definedNames = workbook.definedNames) == null
              ? undefined
              : definedNames[editingIndex];
        if (editingIndex !== null && !existingName)
          throw Error("The defined name no longer exists.");
        if (existingName && existingName.name !== definition.name) {
          if (
            existingName.scopeSheetId !== definition.scopeSheetId ||
            existingName.refersTo !== definition.refersTo ||
            (existingName.comment ?? "") !== (definition.comment ?? "")
          )
            throw Error(
              "Rename separately from changing the definition or comment. Existing scope cannot be changed.",
            );
          (setRenamePreview(
            applyDefinedNameOperation(
              workbook,
              {
                kind: "rename",
                name: existingName.name,
                ...(existingName.scopeSheetId
                  ? { scopeSheetId: existingName.scopeSheetId }
                  : {}),
                newName: definition.name,
              },
              { readOnly },
            ),
          ),
            setErrorMessage(""));
          return;
        }
        const operationResult = applyDefinedNameOperation(
            workbook,
            existingName
              ? {
                  kind: "update",
                  name: existingName.name,
                  ...(existingName.scopeSheetId
                    ? { scopeSheetId: existingName.scopeSheetId }
                    : {}),
                  definition: definition,
                }
              : { kind: "create", definition: definition },
            { readOnly },
          ),
          nextWorkbook = assertNamePreviewCurrent(workbook, operationResult, {
            readOnly,
          });
        (nextWorkbook !== workbook &&
          onWorkbookChange(
            nextWorkbook,
            `${existingName ? "Edit" : "Define"} name: ${definition.name}`,
          ),
          onClick());
      } catch (error) {
        reportError(error);
      }
    },
    testDefinition = () => {
      try {
        const testAddress = range
            .split(":")[0]
            .replace(/\$/g, "")
            .toUpperCase(),
          testWorkbook = structuredClone(workbook),
          definition = {
            name: definedName.trim(),
            refersTo: refersTo.trim(),
            ...(scopeSheetId ? { scopeSheetId: scopeSheetId } : {}),
          };
        testWorkbook.definedNames = (testWorkbook.definedNames ?? []).filter(
          (candidateName) =>
            !(
              candidateName.name.toLowerCase() ===
                definition.name.toLowerCase() &&
              candidateName.scopeSheetId === definition.scopeSheetId
            ),
        );
        const next = applyDefinedNameOperation(testWorkbook, {
            kind: "create",
            definition: definition,
          }).next,
          targetSheetId = scopeSheetId || sheetId,
          sheet = next.sheets.find((sheet) => sheet.id === targetSheetId);
        sheet.cells[testAddress] = {
          raw: `=${definedName.trim()}${formulaTree ? `(${testArguments})` : ""}`,
        };
        const workbookWithSettings = calculateWorkbookWithSettings(next),
          testValue = workbookWithSettings(targetSheetId, testAddress),
          diagnostic = workbookWithSettings
            .getNamedFormulaDiagnostics()
            .find(
              (entry) =>
                entry.sheetId === targetSheetId &&
                entry.address === testAddress,
            ),
          spillRange = workbookWithSettings.getSpillRange(
            targetSheetId,
            testAddress,
          );
        (setTestResult(
          `Test result: ${String(testValue ?? "")}${spillRange ? ` · Array ${spillRange}` : ""}${diagnostic ? ` · ${diagnostic.message}` : ""}`,
        ),
          setErrorMessage(""));
      } catch (error) {
        (reportError(error), setTestResult(""));
      }
    };
  return i.jsx("div", {
    className: "formula-tool-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && onClose();
    },
    onKeyDown: (event) => {
      event.key === "Escape" && onClose();
    },
    children: i.jsxs("section", {
      className: "formula-tools",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": mode === "function" ? "Insert Function" : "Name Manager",
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              children:
                mode === "function" ? "Insert Function" : "Name Manager",
            }),
            i.jsx("button", {
              "aria-label": text("Close formula tools"),
              onClick: onClose,
              lang: language("Close formula tools"),
              "data-ui-label": "Close formula tools",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        mode === "function"
          ? i.jsxs(i.Fragment, {
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Search functions" }),
                    i.jsx("input", {
                      autoFocus: true,
                      "aria-label": text("Search functions"),
                      value: searchQuery,
                      onChange: (event) => setSearchQuery(event.target.value),
                      placeholder: text("SUM, lookup, text…"),
                      lang: language("Search functions"),
                      "data-ui-label": "Search functions",
                    }),
                  ],
                }),
                i.jsxs("div", {
                  className: "formula-function-layout",
                  children: [
                    i.jsx("select", {
                      "aria-label": text("Available functions"),
                      size: 11,
                      value: functionName,
                      onChange: (event) => {
                        (setFunctionName(event.target.value),
                          functionArgumentSignatures[event.target.value] ===
                            "" && setFunctionArguments(""));
                      },
                      lang: language("Available functions"),
                      "data-ui-label": "Available functions",
                      children: matchingFunctionNames.map((se) =>
                        i.jsx("option", { children: se }, se),
                      ),
                    }),
                    i.jsxs("div", {
                      children: [
                        i.jsx("h3", { children: functionName }),
                        i.jsxs("p", {
                          className: "formula-signature",
                          children: [
                            functionName,
                            "(",
                            functionCatalogEntry.inputs,
                            ")",
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Arguments" }),
                            i.jsx("input", {
                              "aria-label": text("Function arguments"),
                              value: functionArguments,
                              onChange: (event) =>
                                setFunctionArguments(event.target.value),
                              readOnly: isLocked,
                              lang: language("Function arguments"),
                              "data-ui-label": "Function arguments",
                            }),
                          ],
                        }),
                        i.jsx("p", {
                          children: i.jsx(LocalizedText, {
                            children:
                              "Separate arguments with commas. Put text inside double quotes. Use cell ranges, defined names or table columns.",
                          }),
                        }),
                        functionName === "SUBTOTAL" &&
                          i.jsx("p", {
                            children: i.jsx(LocalizedText, {
                              children:
                                "Function numbers 101–111 exclude hidden and collapsed outline rows. The quick text filter is view-only and does not change calculations. An external host can supply a separate filtered-row context.",
                            }),
                          }),
                        i.jsxs("output", {
                          "aria-label": text("Formula preview"),
                          lang: language("Formula preview"),
                          "data-ui-label": "Formula preview",
                          children: [
                            i.jsx("code", { children: raw }),
                            i.jsx("strong", { children: previewResult }),
                          ],
                        }),
                        i.jsxs("small", {
                          children: [
                            builtinFunctionNames.length,
                            " ",
                            i.jsx(LocalizedText, {
                              children:
                                "available functions. Compatibility and calculation limits are documented with the package.",
                            }),
                          ],
                        }),
                        i.jsxs("details", {
                          children: [
                            i.jsx("summary", {
                              children: i.jsx(LocalizedText, {
                                children: "Function support and limits",
                              }),
                            }),
                            i.jsx("p", {
                              children: functionCatalogEntry.outputs,
                            }),
                            i.jsxs("p", {
                              children: [
                                i.jsx(LocalizedText, { children: "Errors: " }),
                                functionCatalogEntry.errors.join(", "),
                              ],
                            }),
                            i.jsx("p", {
                              children: functionCatalogEntry.limits,
                            }),
                            i.jsx("p", {
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Local known-answer tolerance: absolute 0.000001; relative 0.000001. Verified Excel version: not run.",
                              }),
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("footer", {
                  children: [
                    i.jsx("button", {
                      onClick: onClose,
                      children: i.jsx(LocalizedText, { children: "Cancel" }),
                    }),
                    i.jsx("button", {
                      className: "formula-primary",
                      disabled: isLocked,
                      onClick: () => {
                        isLocked || (onInsertFormula(raw), onClose());
                      },
                      children: i.jsx(LocalizedText, {
                        children: "Insert formula",
                      }),
                    }),
                  ],
                }),
              ],
            })
          : i.jsxs(i.Fragment, {
              children: [
                i.jsxs("div", {
                  className: "formula-names-list",
                  children: [
                    i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Name",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Refers to",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Scope",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Actions",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: (workbook.definedNames ?? []).map(
                            (savedName, rowIndex) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsxs("td", {
                                      children: [
                                        savedName.name,
                                        parseFormulaTree(savedName.refersTo) &&
                                          i.jsxs("small", {
                                            children: [
                                              " ",
                                              i.jsx(LocalizedText, {
                                                children: "· Reusable formula",
                                              }),
                                            ],
                                          }),
                                      ],
                                    }),
                                    i.jsx("td", {
                                      children: i.jsx("code", {
                                        children: savedName.refersTo,
                                      }),
                                    }),
                                    i.jsx("td", {
                                      children: savedName.scopeSheetId
                                        ? workbook.sheets.find(
                                            (sheet) =>
                                              sheet.id ===
                                              savedName.scopeSheetId,
                                          )?.name
                                        : "Workbook",
                                    }),
                                    i.jsxs("td", {
                                      children: [
                                        i.jsx("button", {
                                          disabled: isLocked,
                                          onClick: () => {
                                            (setEditingIndex(rowIndex),
                                              setWorkbookSnapshot(
                                                JSON.stringify(workbook),
                                              ),
                                              setHasSnapshot(true),
                                              setRenamePreview(null),
                                              setErrorMessage(""),
                                              setTestResult(""),
                                              setDefinedName(savedName.name),
                                              setRefersTo(savedName.refersTo),
                                              setScopeSheetId(
                                                savedName.scopeSheetId ?? "",
                                              ),
                                              setComment(
                                                savedName.comment ?? "",
                                              ));
                                          },
                                          children: i.jsx(LocalizedText, {
                                            children: "Edit",
                                          }),
                                        }),
                                        i.jsx("button", {
                                          disabled: isLocked,
                                          "aria-label": `Delete name ${savedName.name}`,
                                          onClick: () => {
                                            try {
                                              const deleteResult =
                                                  applyDefinedNameOperation(
                                                    workbook,
                                                    {
                                                      kind: "delete",
                                                      name: savedName.name,
                                                      ...(savedName.scopeSheetId
                                                        ? {
                                                            scopeSheetId:
                                                              savedName.scopeSheetId,
                                                          }
                                                        : {}),
                                                    },
                                                    { readOnly: isLocked },
                                                  ),
                                                deletedWorkbook =
                                                  assertNamePreviewCurrent(
                                                    workbook,
                                                    deleteResult,
                                                    {
                                                      readOnly: isLocked,
                                                    },
                                                  );
                                              (deletedWorkbook !== workbook &&
                                                onWorkbookChange(
                                                  deletedWorkbook,
                                                  `Delete name: ${savedName.name}`,
                                                ),
                                                onClick());
                                            } catch (error) {
                                              reportError(error);
                                            }
                                          },
                                          children: i.jsx(LocalizedText, {
                                            children: "Delete",
                                          }),
                                        }),
                                      ],
                                    }),
                                  ],
                                },
                                `${savedName.scopeSheetId}:${savedName.name}`,
                              ),
                          ),
                        }),
                      ],
                    }),
                    !workbook.definedNames?.length &&
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "No names defined. Give a selected range a readable name below.",
                        }),
                      }),
                  ],
                }),
                i.jsxs("div", {
                  className: "formula-name-fields",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Name" }),
                        i.jsx("input", {
                          autoFocus: true,
                          "aria-label": text("Defined name"),
                          value: definedName,
                          onChange: (event) =>
                            updateDraft(setDefinedName, event.target.value),
                          readOnly: isLocked,
                          lang: language("Defined name"),
                          "data-ui-label": "Defined name",
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Scope" }),
                        i.jsxs("select", {
                          "aria-label": text("Name scope"),
                          value: scopeSheetId,
                          onChange: (event) =>
                            updateDraft(setScopeSheetId, event.target.value),
                          disabled: isLocked || editingIndex !== null,
                          lang: language("Name scope"),
                          "data-ui-label": "Name scope",
                          children: [
                            i.jsx("option", {
                              value: "",
                              children: "Workbook",
                            }),
                            workbook.sheets.map((sheet) =>
                              i.jsx(
                                "option",
                                { value: sheet.id, children: sheet.name },
                                sheet.id,
                              ),
                            ),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      className: "formula-wide",
                      children: [
                        i.jsx(LocalizedText, { children: "Refers to" }),
                        i.jsx("input", {
                          "aria-label": text("Name refers to"),
                          value: refersTo,
                          onChange: (event) =>
                            updateDraft(setRefersTo, event.target.value),
                          readOnly: isLocked,
                          lang: language("Name refers to"),
                          "data-ui-label": "Name refers to",
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      className: "formula-wide",
                      children: [
                        i.jsx(LocalizedText, { children: "Comment" }),
                        i.jsx("input", {
                          "aria-label": text("Name comment"),
                          value: comment,
                          onChange: (event) =>
                            updateDraft(setComment, event.target.value),
                          readOnly: isLocked,
                          lang: language("Name comment"),
                          "data-ui-label": "Name comment",
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsx("p", {
                  className: "formula-signature",
                  children: formulaTree
                    ? `Reusable formula: ${definedName || "Name"}(${formulaTree.join(", ")}). Parameters are local to each call.`
                    : "Define a range, constant or calculated formula. Use LET for local variables, or LAMBDA to create a reusable function.",
                }),
                i.jsxs("div", {
                  className: "formula-name-test",
                  children: [
                    i.jsx("button", {
                      disabled: isLocked,
                      onClick: () => {
                        (updateDraft(
                          setRefersTo,
                          "=LAMBDA(amount,rate,amount*(1-rate))",
                        ),
                          definedName || setDefinedName("NetAmount"));
                      },
                      children: i.jsx(LocalizedText, {
                        children: "Use LAMBDA template",
                      }),
                    }),
                    formulaTree &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Test arguments" }),
                          i.jsx("input", {
                            "aria-label": text("Name test arguments"),
                            value: testArguments,
                            onChange: (event) => {
                              (setTestArguments(event.target.value),
                                setTestResult(""));
                            },
                            lang: language("Name test arguments"),
                            "data-ui-label": "Name test arguments",
                          }),
                        ],
                      }),
                    i.jsx("button", {
                      disabled: !definedName.trim(),
                      onClick: testDefinition,
                      children: i.jsx(LocalizedText, {
                        children: "Test definition",
                      }),
                    }),
                  ],
                }),
                testResult &&
                  i.jsxs("p", {
                    role: "status",
                    "aria-label": text("Name test result"),
                    lang: language("Name test result"),
                    "data-ui-label": "Name test result",
                    children: [
                      testResult,
                      " ",
                      i.jsx(LocalizedText, {
                        children:
                          "· Preview only; the workbook has not changed.",
                      }),
                    ],
                  }),
                renamePreview &&
                  i.jsxs("section", {
                    "aria-label": text("Name migration preview"),
                    lang: language("Name migration preview"),
                    "data-ui-label": "Name migration preview",
                    children: [
                      i.jsxs("h3", {
                        children: [
                          i.jsx(LocalizedText, { children: "Rename " }),
                          renamePreview.before?.name,
                          " ",
                          i.jsx(LocalizedText, { children: "to " }),
                          renamePreview.after?.name,
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          renamePreview.changes.length,
                          " ",
                          i.jsx(LocalizedText, {
                            children:
                              "formulas or definitions will be rewritten. ",
                          }),
                          renamePreview.dependencies.length,
                          " ",
                          i.jsx(LocalizedText, {
                            children:
                              "dependent locations were found. Values, formatting and review annotations are retained. Apply commits the reviewed changes together.",
                          }),
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx("strong", {
                            children: i.jsx(LocalizedText, {
                              children: "Definition:",
                            }),
                          }),
                          " ",
                          i.jsx("code", {
                            children: renamePreview.before?.refersTo,
                          }),
                          " →",
                          " ",
                          i.jsx("code", {
                            children: renamePreview.after?.refersTo,
                          }),
                        ],
                      }),
                      i.jsx("div", {
                        className: "formula-names-list",
                        children: i.jsxs("table", {
                          children: [
                            i.jsx("thead", {
                              children: i.jsxs("tr", {
                                children: [
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Location",
                                    }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Before",
                                    }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "After",
                                    }),
                                  }),
                                ],
                              }),
                            }),
                            i.jsx("tbody", {
                              children: renamePreview.changes.map(
                                (change, changeIndex) =>
                                  i.jsxs(
                                    "tr",
                                    {
                                      children: [
                                        i.jsx("td", { children: change.label }),
                                        i.jsx("td", {
                                          children: i.jsx("code", {
                                            children: change.before,
                                          }),
                                        }),
                                        i.jsx("td", {
                                          children: i.jsx("code", {
                                            children: change.after,
                                          }),
                                        }),
                                      ],
                                    },
                                    changeIndex,
                                  ),
                              ),
                            }),
                          ],
                        }),
                      }),
                      i.jsxs("details", {
                        children: [
                          i.jsx("summary", {
                            children: i.jsx(LocalizedText, {
                              children: "Dependent locations",
                            }),
                          }),
                          i.jsx("ul", {
                            children: renamePreview.dependencies.map(
                              (dependency, dependencyIndex) =>
                                i.jsx(
                                  "li",
                                  { children: dependency.label },
                                  dependencyIndex,
                                ),
                            ),
                          }),
                        ],
                      }),
                      i.jsx("button", {
                        onClick: () => setRenamePreview(null),
                        children: i.jsx(LocalizedText, {
                          children: "Cancel rename",
                        }),
                      }),
                      i.jsx("button", {
                        disabled: isLocked,
                        onClick: () => {
                          try {
                            const renamedWorkbook = assertNamePreviewCurrent(
                              workbook,
                              renamePreview,
                              {
                                readOnly: isLocked,
                              },
                            );
                            (renamedWorkbook !== workbook &&
                              onWorkbookChange(
                                renamedWorkbook,
                                `Rename name: ${renamePreview.before?.name} → ${renamePreview.after?.name}`,
                              ),
                              onClick());
                          } catch (error) {
                            reportError(error);
                          }
                        },
                        children: i.jsx(LocalizedText, {
                          children: "Apply rename",
                        }),
                      }),
                    ],
                  }),
                errorMessage &&
                  i.jsx("p", {
                    role: "alert",
                    className: "formula-tool-error",
                    children: errorMessage,
                  }),
                i.jsxs("footer", {
                  children: [
                    editingIndex !== null &&
                      i.jsx("button", {
                        onClick: onClick,
                        children: i.jsx(LocalizedText, {
                          children: "New name",
                        }),
                      }),
                    i.jsx("button", {
                      onClick: onClose,
                      children: i.jsx(LocalizedText, { children: "Close" }),
                    }),
                    i.jsx("button", {
                      className: "formula-primary",
                      disabled:
                        isLocked || !definedName.trim() || !!renamePreview,
                      onClick: saveDefinedName,
                      children:
                        editingIndex === null ? "Add name" : "Save name",
                    }),
                  ],
                }),
              ],
            }),
      ],
    }),
  });
}
const nameScanLimits = Object.freeze({
  maxLocations: 1e4,
  maxCharacters: 2e6,
  maxFindings: 2e3,
  maxNameDepth: 32,
});

export {
  nameScanLimits,
  buildFunctionCatalogEntry,
  functionBrowserCategories,
  buildFunctionCatalog,
  FunctionAndNameDialog,
  buildFunctionBrowserGroups,
};
