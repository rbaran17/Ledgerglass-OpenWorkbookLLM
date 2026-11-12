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
  arrayIncludesValue = (e, r) => r.includes(e),
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
function classifyFunctionFamily(e) {
  if (((e = e.toUpperCase()), arrayIncludesValue(e, databaseFunctionNames)))
    return "database";
  if (regressionFunctionNames.includes(e)) return "regression";
  if (matrixFunctionNames.includes(e)) return "matrix";
  if (financialFunctionNames.includes(e)) return "financial";
  if (e === "CONVERT") return "engineering";
  if (arrayIncludesValue(e, statisticalFunctionNames)) return "statistics";
  if (connectorFunctionFamilies[e]) return connectorFunctionFamilies[e];
  if (e.startsWith("CUBE")) return "cube";
  const r = functionSpecsByName[e]?.family;
  return r === "statistical"
    ? "statistics"
    : r === "distribution" || r === "engineering" || r === "financial"
      ? r
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
function buildFunctionCatalogEntry(e) {
  const r = e.replace(/^_xlfn\.(?:_xlws\.)?/i, "").toUpperCase(),
    t = classifyFunctionFamily(r),
    n = familyResultContracts[t],
    o = localFunctionNameSet.has(r);
  return {
    name: r,
    family: t,
    status: o
      ? "local"
      : t === "cube" || t === "network"
        ? "connector-required"
        : "unimplemented",
    inputs:
      functionSpecsByName[r]?.signature ??
      "Not adopted; consult the source function or add-in contract.",
    outputs: o
      ? (functionDescriptions[r] ?? n.output)
      : "No local implementation",
    errors: [
      ...n.errors,
      "#LIMIT!",
      "#CYCLE!",
      ...(o ? ["#TIMEOUT!", "#BLOCKED!"] : []),
    ],
    numericTolerance: o
      ? {
          absolute: 1e-6,
          relative: 1e-6,
          appliesTo:
            "Local known-answer comparison only; exact text/logical/errors. Not an accuracy guarantee across all inputs.",
        }
      : null,
    excelVerification: { status: "not-run", version: null, evidence: null },
    limits:
      (arrayIncludesValue(r, dateFunctionNames)
        ? "Gregorian 1900 serials 0–2958465 (compatibility serial 60) or 1904 serials 0–2957003; full-year ISO/configured MDY or DMY text; recorded IANA-zone clock shared per calculation. Up to 1000 holiday inputs, 366000 business-day offsets/date spans. The date system persists in native JSON and XLSX; text order and time zone are native JSON only. DATEDIF MD/YD and US day-count edge conventions follow the documented local policy; independent Excel parity unverified."
        : n.limits) +
      (o
        ? " Workbook work and accounted-memory budgets apply across a run; synchronous APIs use cooperative time checks and isolated worker execution supports deadline termination/cancellation."
        : "") +
      (["RAND", "RANDBETWEEN"].includes(r)
        ? " Unseeded runtime pseudorandom values are stable within the revision calculation and refresh on recalculation. They are not reproducible audit samples. RANDBETWEEN uses ceil(bottom)/floor(top) and refuses bounds or inclusive spans outside safe integers."
        : ""),
    evidence: o
      ? arrayIncludesValue(r, dateFunctionNames)
        ? [
            "tests/date-calendar.test.ts",
            "tests/date-calendar-integration.test.ts",
            "docs/DATE-CALENDAR-ACCEPTANCE.md",
          ]
        : n.tests.split(";")
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
  const e = buildFunctionCatalog();
  return functionBrowserCategories.map((r) => ({
    ...r,
    adopted: e
      .filter((t) => t.family === r.id && t.status === "local")
      .map((t) => t.name),
    unavailable: e
      .filter((t) => t.family === r.id && t.status !== "local")
      .map((t) => t.name),
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
    c = workbook.sheets.find((se) => se.id === sheetId),
    [h, f] = ne.useState(""),
    [p, m] = ne.useState("SUM"),
    [v, g] = ne.useState(range),
    [b, y] = ne.useState(""),
    [w, E] = ne.useState(
      `='${(c?.name ?? "Sheet1").replace(/'/g, "''")}'!${range}`,
    ),
    [A, k] = ne.useState(""),
    [F, I] = ne.useState(""),
    [O, j] = ne.useState(null),
    [z, V] = ne.useState(""),
    [D, B] = ne.useState(JSON.stringify(workbook)),
    [_, C] = ne.useState(false),
    [T, N] = ne.useState(null),
    [R, L] = ne.useState("100,0.2"),
    [$, U] = ne.useState(""),
    G = parseFormulaTree(w),
    H = (se, we) => {
      (_ || B(JSON.stringify(workbook)), C(true), N(null), U(""), se(we));
    },
    Q = readOnly || !c || !!workbook.structureProtection?.enabled,
    X$1 = (se) =>
      V(se instanceof Error ? se.message : "The name could not be saved."),
    J = ne.useMemo(
      () =>
        builtinFunctionNames.filter((se) =>
          se.toLowerCase().includes(h.toLowerCase()),
        ),
      [h],
    ),
    q = `=${p}(${v})`,
    W = buildFunctionCatalogEntry(p),
    K = ne.useMemo(() => {
      if (mode !== "function") return "";
      const se = range.split(":")[0].replace(/\$/g, "").toUpperCase(),
        we = {
          ...workbook,
          sheets: workbook.sheets.map((Ee) =>
            Ee.id === sheetId
              ? { ...Ee, cells: { ...Ee.cells, [se]: { raw: q } } }
              : Ee,
          ),
        };
      try {
        return String(calculateWorkbookWithSettings(we)(sheetId, se) ?? "");
      } catch {
        return "Check the function arguments.";
      }
    }, [q, mode, workbook, sheetId, range]),
    te = () => {
      (y(""),
        I(""),
        j(null),
        V(""),
        U(""),
        C(false),
        N(null),
        B(JSON.stringify(workbook)),
        k(""));
    },
    ae = () => {
      var se;
      try {
        if (Q) throw Error("This workbook is read-only or protected.");
        if (_ && D !== JSON.stringify(workbook))
          throw Error(
            "The workbook changed while this name was being edited. Choose New name or reopen Edit to review current content.",
          );
        const we = {
            name: b.trim(),
            refersTo: w.trim(),
            ...(A ? { scopeSheetId: A } : {}),
            ...(F ? { comment: F } : {}),
          },
          Ee =
            O === null || (se = workbook.definedNames) == null
              ? undefined
              : se[O];
        if (O !== null && !Ee)
          throw Error("The defined name no longer exists.");
        if (Ee && Ee.name !== we.name) {
          if (
            Ee.scopeSheetId !== we.scopeSheetId ||
            Ee.refersTo !== we.refersTo ||
            (Ee.comment ?? "") !== (we.comment ?? "")
          )
            throw Error(
              "Rename separately from changing the definition or comment. Existing scope cannot be changed.",
            );
          (N(
            applyDefinedNameOperation(
              workbook,
              {
                kind: "rename",
                name: Ee.name,
                ...(Ee.scopeSheetId ? { scopeSheetId: Ee.scopeSheetId } : {}),
                newName: we.name,
              },
              { readOnly },
            ),
          ),
            V(""));
          return;
        }
        const Ae = applyDefinedNameOperation(
            workbook,
            Ee
              ? {
                  kind: "update",
                  name: Ee.name,
                  ...(Ee.scopeSheetId ? { scopeSheetId: Ee.scopeSheetId } : {}),
                  definition: we,
                }
              : { kind: "create", definition: we },
            { readOnly },
          ),
          ye = assertNamePreviewCurrent(workbook, Ae, { readOnly });
        (ye !== workbook &&
          onWorkbookChange(ye, `${Ee ? "Edit" : "Define"} name: ${we.name}`),
          te());
      } catch (we) {
        X$1(we);
      }
    },
    oe = () => {
      try {
        const se = range.split(":")[0].replace(/\$/g, "").toUpperCase(),
          we = structuredClone(workbook),
          Ee = {
            name: b.trim(),
            refersTo: w.trim(),
            ...(A ? { scopeSheetId: A } : {}),
          };
        we.definedNames = (we.definedNames ?? []).filter(
          (Me) =>
            !(
              Me.name.toLowerCase() === Ee.name.toLowerCase() &&
              Me.scopeSheetId === Ee.scopeSheetId
            ),
        );
        const Ae = applyDefinedNameOperation(we, {
            kind: "create",
            definition: Ee,
          }).next,
          ye = A || sheetId,
          Te = Ae.sheets.find((Me) => Me.id === ye);
        Te.cells[se] = { raw: `=${b.trim()}${G ? `(${R})` : ""}` };
        const Pe = calculateWorkbookWithSettings(Ae),
          qe = Pe(ye, se),
          ht = Pe.getNamedFormulaDiagnostics().find(
            (Me) => Me.sheetId === ye && Me.address === se,
          ),
          Re = Pe.getSpillRange(ye, se);
        (U(
          `Test result: ${String(qe ?? "")}${Re ? ` · Array ${Re}` : ""}${ht ? ` · ${ht.message}` : ""}`,
        ),
          V(""));
      } catch (se) {
        (X$1(se), U(""));
      }
    };
  return i.jsx("div", {
    className: "formula-tool-backdrop",
    onMouseDown: (se) => {
      se.target === se.currentTarget && onClose();
    },
    onKeyDown: (se) => {
      se.key === "Escape" && onClose();
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
                      value: h,
                      onChange: (se) => f(se.target.value),
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
                      value: p,
                      onChange: (se) => {
                        (m(se.target.value),
                          functionArgumentSignatures[se.target.value] === "" &&
                            g(""));
                      },
                      lang: language("Available functions"),
                      "data-ui-label": "Available functions",
                      children: J.map((se) =>
                        i.jsx("option", { children: se }, se),
                      ),
                    }),
                    i.jsxs("div", {
                      children: [
                        i.jsx("h3", { children: p }),
                        i.jsxs("p", {
                          className: "formula-signature",
                          children: [p, "(", W.inputs, ")"],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Arguments" }),
                            i.jsx("input", {
                              "aria-label": text("Function arguments"),
                              value: v,
                              onChange: (se) => g(se.target.value),
                              readOnly: Q,
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
                        p === "SUBTOTAL" &&
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
                            i.jsx("code", { children: q }),
                            i.jsx("strong", { children: K }),
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
                            i.jsx("p", { children: W.outputs }),
                            i.jsxs("p", {
                              children: [
                                i.jsx(LocalizedText, { children: "Errors: " }),
                                W.errors.join(", "),
                              ],
                            }),
                            i.jsx("p", { children: W.limits }),
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
                      disabled: Q,
                      onClick: () => {
                        Q || (onInsertFormula(q), onClose());
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
                            (se, we) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsxs("td", {
                                      children: [
                                        se.name,
                                        parseFormulaTree(se.refersTo) &&
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
                                        children: se.refersTo,
                                      }),
                                    }),
                                    i.jsx("td", {
                                      children: se.scopeSheetId
                                        ? workbook.sheets.find(
                                            (Ae) => Ae.id === se.scopeSheetId,
                                          )?.name
                                        : "Workbook",
                                    }),
                                    i.jsxs("td", {
                                      children: [
                                        i.jsx("button", {
                                          disabled: Q,
                                          onClick: () => {
                                            (j(we),
                                              B(JSON.stringify(workbook)),
                                              C(true),
                                              N(null),
                                              V(""),
                                              U(""),
                                              y(se.name),
                                              E(se.refersTo),
                                              k(se.scopeSheetId ?? ""),
                                              I(se.comment ?? ""));
                                          },
                                          children: i.jsx(LocalizedText, {
                                            children: "Edit",
                                          }),
                                        }),
                                        i.jsx("button", {
                                          disabled: Q,
                                          "aria-label": `Delete name ${se.name}`,
                                          onClick: () => {
                                            try {
                                              const Ae =
                                                  applyDefinedNameOperation(
                                                    workbook,
                                                    {
                                                      kind: "delete",
                                                      name: se.name,
                                                      ...(se.scopeSheetId
                                                        ? {
                                                            scopeSheetId:
                                                              se.scopeSheetId,
                                                          }
                                                        : {}),
                                                    },
                                                    { readOnly: Q },
                                                  ),
                                                ye = assertNamePreviewCurrent(
                                                  workbook,
                                                  Ae,
                                                  {
                                                    readOnly: Q,
                                                  },
                                                );
                                              (ye !== workbook &&
                                                onWorkbookChange(
                                                  ye,
                                                  `Delete name: ${se.name}`,
                                                ),
                                                te());
                                            } catch (Ae) {
                                              X$1(Ae);
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
                                `${se.scopeSheetId}:${se.name}`,
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
                          value: b,
                          onChange: (se) => H(y, se.target.value),
                          readOnly: Q,
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
                          value: A,
                          onChange: (se) => H(k, se.target.value),
                          disabled: Q || O !== null,
                          lang: language("Name scope"),
                          "data-ui-label": "Name scope",
                          children: [
                            i.jsx("option", {
                              value: "",
                              children: "Workbook",
                            }),
                            workbook.sheets.map((se) =>
                              i.jsx(
                                "option",
                                { value: se.id, children: se.name },
                                se.id,
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
                          value: w,
                          onChange: (se) => H(E, se.target.value),
                          readOnly: Q,
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
                          value: F,
                          onChange: (se) => H(I, se.target.value),
                          readOnly: Q,
                          lang: language("Name comment"),
                          "data-ui-label": "Name comment",
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsx("p", {
                  className: "formula-signature",
                  children: G
                    ? `Reusable formula: ${b || "Name"}(${G.join(", ")}). Parameters are local to each call.`
                    : "Define a range, constant or calculated formula. Use LET for local variables, or LAMBDA to create a reusable function.",
                }),
                i.jsxs("div", {
                  className: "formula-name-test",
                  children: [
                    i.jsx("button", {
                      disabled: Q,
                      onClick: () => {
                        (H(E, "=LAMBDA(amount,rate,amount*(1-rate))"),
                          b || y("NetAmount"));
                      },
                      children: i.jsx(LocalizedText, {
                        children: "Use LAMBDA template",
                      }),
                    }),
                    G &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Test arguments" }),
                          i.jsx("input", {
                            "aria-label": text("Name test arguments"),
                            value: R,
                            onChange: (se) => {
                              (L(se.target.value), U(""));
                            },
                            lang: language("Name test arguments"),
                            "data-ui-label": "Name test arguments",
                          }),
                        ],
                      }),
                    i.jsx("button", {
                      disabled: !b.trim(),
                      onClick: oe,
                      children: i.jsx(LocalizedText, {
                        children: "Test definition",
                      }),
                    }),
                  ],
                }),
                $ &&
                  i.jsxs("p", {
                    role: "status",
                    "aria-label": text("Name test result"),
                    lang: language("Name test result"),
                    "data-ui-label": "Name test result",
                    children: [
                      $,
                      " ",
                      i.jsx(LocalizedText, {
                        children:
                          "· Preview only; the workbook has not changed.",
                      }),
                    ],
                  }),
                T &&
                  i.jsxs("section", {
                    "aria-label": text("Name migration preview"),
                    lang: language("Name migration preview"),
                    "data-ui-label": "Name migration preview",
                    children: [
                      i.jsxs("h3", {
                        children: [
                          i.jsx(LocalizedText, { children: "Rename " }),
                          T.before?.name,
                          " ",
                          i.jsx(LocalizedText, { children: "to " }),
                          T.after?.name,
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          T.changes.length,
                          " ",
                          i.jsx(LocalizedText, {
                            children:
                              "formulas or definitions will be rewritten. ",
                          }),
                          T.dependencies.length,
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
                            children: T.before?.refersTo,
                          }),
                          " →",
                          " ",
                          i.jsx("code", {
                            children: T.after?.refersTo,
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
                              children: T.changes.map((se, we) =>
                                i.jsxs(
                                  "tr",
                                  {
                                    children: [
                                      i.jsx("td", { children: se.label }),
                                      i.jsx("td", {
                                        children: i.jsx("code", {
                                          children: se.before,
                                        }),
                                      }),
                                      i.jsx("td", {
                                        children: i.jsx("code", {
                                          children: se.after,
                                        }),
                                      }),
                                    ],
                                  },
                                  we,
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
                            children: T.dependencies.map((se, we) =>
                              i.jsx("li", { children: se.label }, we),
                            ),
                          }),
                        ],
                      }),
                      i.jsx("button", {
                        onClick: () => N(null),
                        children: i.jsx(LocalizedText, {
                          children: "Cancel rename",
                        }),
                      }),
                      i.jsx("button", {
                        disabled: Q,
                        onClick: () => {
                          try {
                            const Ee = assertNamePreviewCurrent(workbook, T, {
                              readOnly: Q,
                            });
                            (Ee !== workbook &&
                              onWorkbookChange(
                                Ee,
                                `Rename name: ${T.before?.name} → ${T.after?.name}`,
                              ),
                              te());
                          } catch (Ee) {
                            X$1(Ee);
                          }
                        },
                        children: i.jsx(LocalizedText, {
                          children: "Apply rename",
                        }),
                      }),
                    ],
                  }),
                z &&
                  i.jsx("p", {
                    role: "alert",
                    className: "formula-tool-error",
                    children: z,
                  }),
                i.jsxs("footer", {
                  children: [
                    O !== null &&
                      i.jsx("button", {
                        onClick: te,
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
                      disabled: Q || !b.trim() || !!T,
                      onClick: ae,
                      children: O === null ? "Add name" : "Save name",
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
