import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { $p, VS } from "./part-01.js";
import { AR, Fc, QN, Ty, Vn, q6 } from "./part-03.js";
import { Dr, M } from "./part-04.js";
import { $T, UT } from "./part-33.js";

const Ume = { family: "local", signature: "number" },
  $me = {
    family: "financial",
    signature:
      "issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method]",
  },
  zme = {
    family: "financial",
    signature: "issue, settlement, rate, par, [basis]",
  },
  Vme = { family: "local", signature: "number" },
  Wme = { family: "local", signature: "number" },
  Hme = { family: "local", signature: "number" },
  qme = { family: "local", signature: "number" },
  Xme = {
    family: "reference",
    signature: "rowNumber, columnNumber, [absNum], [a1], [sheetText]",
  },
  Gme = {
    family: "math",
    signature: "function_num, options, ref1 or array, [ref2 or k], …",
  },
  Zme = { family: "logical", signature: "params1, [params2], …" },
  Kme = { family: "local", signature: "text" },
  Yme = { family: "local", signature: "reference" },
  Jme = { family: "text", signature: "text" },
  Qme = { family: "local", signature: "number" },
  ege = { family: "local", signature: "number" },
  tge = { family: "local", signature: "number" },
  rge = { family: "trigonometry", signature: "x, y" },
  nge = { family: "local", signature: "number" },
  ige = { family: "statistical", signature: "numbers1, [numbers2], …" },
  oge = { family: "statistical", signature: "number1, [number2], …" },
  age = { family: "statistical", signature: "numbers1, [numbers2], …" },
  sge = { family: "statistical", signature: "range, criteria, averageRange" },
  lge = {
    family: "statistical",
    signature: "average_range, criteria_range1, criteria1, …",
  },
  uge = { family: "text", signature: "number" },
  cge = { family: "math", signature: "number, radix, [minLength]" },
  dge = { family: "engineering", signature: "x, n" },
  fge = { family: "engineering", signature: "x, n" },
  hge = { family: "engineering", signature: "x, n" },
  pge = { family: "engineering", signature: "x, n" },
  mge = { family: "engineering", signature: "number" },
  gge = { family: "engineering", signature: "number, [places]" },
  vge = { family: "engineering", signature: "number, [places]" },
  bge = { family: "engineering", signature: "number1, number2" },
  yge = { family: "engineering", signature: "number, shiftAmount" },
  wge = { family: "engineering", signature: "number1, number2" },
  xge = { family: "engineering", signature: "number, shiftAmount" },
  Ege = { family: "engineering", signature: "number1, number2" },
  Cge = { family: "math", signature: "number, significance" },
  Age = { family: "text", signature: "number" },
  kge = { family: "reference", signature: "index_num, value1, [value2], …" },
  Sge = { family: "local", signature: "array, col_num1, …" },
  Tge = { family: "local", signature: "array, row_num1, …" },
  Nge = { family: "text", signature: "text" },
  Rge = { family: "text", signature: "text" },
  Ige = { family: "reference", signature: "obj" },
  _ge = { family: "reference", signature: "obj" },
  Dge = { family: "math", signature: "number, numberChosen" },
  Oge = { family: "math", signature: "number, numberChosen" },
  Mge = { family: "engineering", signature: "realNum, iNum, [suffix]" },
  jge = { family: "text", signature: "text1, [text2], …" },
  Fge = { family: "text", signature: "params1, [params2], …" },
  Bge = { family: "local", signature: "number, from_unit, to_unit" },
  Lge = { family: "distribution", signature: "[array1], [array2]" },
  Pge = { family: "local", signature: "number" },
  Uge = { family: "local", signature: "number" },
  $ge = { family: "local", signature: "number" },
  zge = { family: "local", signature: "number" },
  Vge = { family: "statistical", signature: "number" },
  Wge = { family: "local", signature: "number" },
  Hge = { family: "local", signature: "range" },
  qge = { family: "statistical", signature: "range, criteria" },
  Xge = { family: "local", signature: "criteria_range1, criteria1, …" },
  Gge = { family: "local", signature: "number" },
  Zge = { family: "local", signature: "number" },
  Kge = { family: "date", signature: "year, month, day" },
  Yge = { family: "date", signature: "startDate, endDate, unit" },
  Jge = { family: "date", signature: "dateText" },
  Qge = { family: "local", signature: "database, field, criteria" },
  eve = { family: "local", signature: "serial_number" },
  tve = { family: "date", signature: "endDate, startDate" },
  rve = { family: "date", signature: "startDate, endDate, [method]" },
  nve = {
    family: "financial",
    signature: "cost, salvage, life, period, [month]",
  },
  ive = { family: "text", signature: "text" },
  ove = { family: "local", signature: "database, [field], criteria" },
  ave = { family: "local", signature: "database, [field], criteria" },
  sve = { family: "local", signature: "cost, salvage, life, period, [factor]" },
  lve = { family: "engineering", signature: "number, [places]" },
  uve = { family: "engineering", signature: "number, [places]" },
  cve = { family: "engineering", signature: "number, [places]" },
  dve = { family: "math", signature: "text, radix" },
  fve = { family: "math", signature: "radians" },
  hve = { family: "engineering", signature: "number1, [number2]" },
  pve = { family: "distribution", signature: "numbers1, [numbers2], …" },
  mve = { family: "local", signature: "database, field, criteria" },
  gve = { family: "local", signature: "database, field, criteria" },
  vve = { family: "local", signature: "database, field, criteria" },
  bve = { family: "text", signature: "number, [decimals]" },
  yve = { family: "local", signature: "database, field, criteria" },
  wve = { family: "local", signature: "array, rows, [columns]" },
  xve = { family: "local", signature: "database, field, criteria" },
  Eve = { family: "local", signature: "database, field, criteria" },
  Cve = { family: "local", signature: "database, field, criteria" },
  Ave = { family: "local", signature: "database, field, criteria" },
  kve = { family: "local", signature: "database, field, criteria" },
  Sve = { family: "date", signature: "startDate, months" },
  Tve = { family: "local", signature: "text" },
  Nve = { family: "date", signature: "startDate, months" },
  Rve = { family: "engineering", signature: "lowerLimit, [upperLimit]" },
  Ive = { family: "engineering", signature: "x" },
  _ve = { family: "math", signature: "number" },
  Dve = { family: "text", signature: "text1, text2" },
  Ove = { family: "math", signature: "number" },
  Mve = { family: "local", signature: "array, rows, [columns], [pad_with]" },
  jve = { family: "math", signature: "number" },
  Fve = { family: "math", signature: "number" },
  Bve = { family: "logical", signature: "" },
  Lve = { family: "local", signature: "array, include, [if_empty]" },
  Pve = { family: "text", signature: "findText, withinText, [startNum]" },
  Uve = { family: "text", signature: "params1, [params2], …" },
  $ve = { family: "distribution", signature: "x" },
  zve = { family: "distribution", signature: "x" },
  Vve = { family: "text", signature: "number, [decimals], [noCommas]" },
  Wve = { family: "math", signature: "number, significance" },
  Hve = { family: "distribution", signature: "x, [knownYs], [knownXs]" },
  qve = { family: "local", signature: "reference" },
  Xve = { family: "distribution", signature: "[dataArray], [binsArray]" },
  Gve = { family: "local", signature: "rate, nper, pmt, [pv], [type]" },
  Zve = { family: "distribution", signature: "x" },
  Kve = { family: "distribution", signature: "x" },
  Yve = { family: "distribution", signature: "z" },
  Jve = { family: "math", signature: "params1, [params2], …" },
  Qve = { family: "distribution", signature: "numbers1, [numbers2], …" },
  e1e = { family: "engineering", signature: "number, [step]" },
  t1e = {
    family: "regression",
    signature: "known_ys, [known_xs], [new_xs], [const]",
  },
  r1e = { family: "distribution", signature: "numbers1, [numbers2], …" },
  n1e = { family: "engineering", signature: "number, [places]" },
  i1e = { family: "engineering", signature: "number" },
  o1e = { family: "engineering", signature: "number, places" },
  a1e = {
    family: "reference",
    signature: "lookup_value, table_array, row_index_num, [range_lookup]",
  },
  s1e = { family: "local", signature: "serial_number" },
  l1e = { family: "local", signature: "array1, [array2], …" },
  u1e = {
    family: "logical",
    signature: "logical_test, value_if_true, [value_if_false]",
  },
  c1e = { family: "logical", signature: "value, value_if_error" },
  d1e = { family: "local", signature: "value, value_if_na" },
  f1e = { family: "logical", signature: "logical_test1, value_if_true1, …" },
  h1e = { family: "engineering", signature: "iNumber" },
  p1e = { family: "engineering", signature: "iNumber" },
  m1e = { family: "engineering", signature: "iNumber" },
  g1e = { family: "engineering", signature: "iNumber" },
  v1e = { family: "engineering", signature: "iNumber" },
  b1e = { family: "engineering", signature: "iNumber" },
  y1e = { family: "engineering", signature: "iNumber" },
  w1e = { family: "engineering", signature: "iNumber" },
  x1e = { family: "engineering", signature: "iNumber" },
  E1e = { family: "engineering", signature: "iNumber1, iNumber2" },
  C1e = { family: "engineering", signature: "iNumber" },
  A1e = { family: "engineering", signature: "iNumber" },
  k1e = { family: "engineering", signature: "iNumber" },
  S1e = { family: "engineering", signature: "iNumber" },
  T1e = { family: "engineering", signature: "iNumber, number" },
  N1e = { family: "engineering", signature: "params1, [params2], …" },
  R1e = { family: "engineering", signature: "iNumber" },
  I1e = { family: "engineering", signature: "iNumber" },
  _1e = { family: "engineering", signature: "iNumber" },
  D1e = { family: "engineering", signature: "iNumber" },
  O1e = { family: "engineering", signature: "iNumber" },
  M1e = { family: "engineering", signature: "iNumber" },
  j1e = { family: "engineering", signature: "iNumber1, iNumber2" },
  F1e = { family: "engineering", signature: "params1, [params2], …" },
  B1e = { family: "engineering", signature: "iNumber" },
  L1e = { family: "reference", signature: "array, row_num, [column_num]" },
  P1e = { family: "math", signature: "number" },
  U1e = { family: "distribution", signature: "[knownYs], [knownXs]" },
  $1e = { family: "financial", signature: "rate, per, nper, pv, [fv], [type]" },
  z1e = { family: "local", signature: "values, [guess]" },
  V1e = { family: "information", signature: "value" },
  W1e = { family: "information", signature: "value" },
  H1e = { family: "information", signature: "value" },
  q1e = { family: "local", signature: "number" },
  X1e = { family: "local", signature: "reference" },
  G1e = { family: "information", signature: "value" },
  Z1e = { family: "information", signature: "value" },
  K1e = { family: "information", signature: "value" },
  Y1e = { family: "information", signature: "value" },
  J1e = { family: "date", signature: "serialOrString" },
  Q1e = { family: "information", signature: "value" },
  ebe = { family: "information", signature: "value" },
  tbe = { family: "distribution", signature: "numbers1, [numbers2], …" },
  rbe = {
    family: "local",
    signature: "[parameter1, parameter2, …], calculation",
  },
  nbe = { family: "statistical", signature: "array, k" },
  ibe = { family: "math", signature: "params1, [params2], …" },
  obe = { family: "text", signature: "text, [num_chars]" },
  abe = { family: "text", signature: "params1, [params2], …" },
  sbe = { family: "text", signature: "text" },
  lbe = {
    family: "local",
    signature: "name1, value1, [name2, value2], …, calculation",
  },
  ube = {
    family: "regression",
    signature: "known_ys, [known_xs], [const], [stats]",
  },
  cbe = { family: "local", signature: "number" },
  dbe = { family: "math", signature: "number, [base]" },
  fbe = { family: "local", signature: "number" },
  hbe = {
    family: "regression",
    signature: "known_ys, [known_xs], [const], [stats]",
  },
  pbe = {
    family: "local",
    signature: "lookup_value, lookup_vector, [result_vector]",
  },
  mbe = { family: "text", signature: "text" },
  gbe = {
    family: "reference",
    signature: "lookup_value, lookup_array, [match_type]",
  },
  vbe = { family: "statistical", signature: "number1, [number2], …" },
  bbe = { family: "statistical", signature: "value1, [value2, ...]" },
  ybe = {
    family: "statistical",
    signature: "max_range, criteria_range1, criteria1, …",
  },
  wbe = { family: "math", signature: "array" },
  xbe = { family: "statistical", signature: "number1, [number2], …" },
  Ebe = { family: "text", signature: "text, start_num, num_chars" },
  Cbe = { family: "text", signature: "params1, [params2], …" },
  Abe = { family: "statistical", signature: "number1, [number2], …" },
  kbe = { family: "statistical", signature: "value1, [value2, ...]" },
  Sbe = {
    family: "statistical",
    signature: "min_range, criteria_range1, criteria1, …",
  },
  Tbe = { family: "local", signature: "serial_number" },
  Nbe = { family: "math", signature: "array" },
  Rbe = {
    family: "financial",
    signature: "values, finance_rate, reinvest_rate",
  },
  Ibe = { family: "math", signature: "array1, array2" },
  _be = { family: "math", signature: "number, divisor" },
  Dbe = { family: "local", signature: "number1, [number2], …" },
  Obe = { family: "local", signature: "serial_number" },
  Mbe = { family: "math", signature: "number, multiple" },
  jbe = { family: "math", signature: "numbers1, [numbers2], …" },
  Fbe = { family: "math", signature: "dimension" },
  Bbe = { family: "local", signature: "value" },
  Lbe = { family: "information", signature: "" },
  Pbe = { family: "date", signature: "startDate, endDate, holidays" },
  Ube = { family: "logical", signature: "logical" },
  $be = { family: "date", signature: "" },
  zbe = { family: "local", signature: "rate, pmt, pv, [fv], [type]" },
  Vbe = { family: "local", signature: "rate, value1, [value2], …" },
  Wbe = {
    family: "text",
    signature: "text, [decimalSeparator], [groupSeparator]",
  },
  Hbe = { family: "engineering", signature: "number, [places]" },
  qbe = { family: "engineering", signature: "number" },
  Xbe = { family: "engineering", signature: "number, [places]" },
  Gbe = { family: "math", signature: "number" },
  Zbe = { family: "logical", signature: "params1, [params2], …" },
  Kbe = { family: "regression", signature: "array1, array2" },
  Ybe = { family: "local", signature: "array, k" },
  Jbe = { family: "statistical", signature: "number, number_chosen" },
  Qbe = { family: "statistical", signature: "number, number_chosen" },
  eye = { family: "distribution", signature: "x" },
  tye = { family: "math", signature: "" },
  rye = { family: "local", signature: "rate, nper, pv, [fv], [type]" },
  nye = { family: "math", signature: "number, power" },
  iye = { family: "financial", signature: "rate, per, nper, pv, [fv], [type]" },
  oye = {
    family: "financial",
    signature:
      "settlement, maturity, rate, yld, redemption, frequency, [basis]",
  },
  aye = { family: "math", signature: "numbers1, [numbers2], …" },
  sye = { family: "text", signature: "text" },
  lye = { family: "local", signature: "rate, nper, pmt, [fv], [type]" },
  uye = { family: "local", signature: "array, quart" },
  cye = { family: "math", signature: "numerator, denominator" },
  dye = { family: "math", signature: "degrees" },
  fye = { family: "math", signature: "" },
  hye = { family: "math", signature: "bottom, top" },
  pye = { family: "local", signature: "number, ref, [order]" },
  mye = { family: "local", signature: "nper, pmt, pv, [fv], [type], [guess]" },
  gye = {
    family: "text",
    signature: "old_text, start_num, num_chars, new_text",
  },
  vye = { family: "text", signature: "params1, [params2], …" },
  bye = { family: "text", signature: "text, number_times" },
  yye = { family: "text", signature: "text, [num_chars]" },
  wye = { family: "text", signature: "params1, [params2], …" },
  xye = { family: "math", signature: "number, [form]" },
  Eye = { family: "math", signature: "number, num_digits" },
  Cye = { family: "math", signature: "number, num_digits" },
  Aye = { family: "math", signature: "number, num_digits" },
  kye = { family: "reference", signature: "obj" },
  Sye = { family: "reference", signature: "obj" },
  Tye = { family: "distribution", signature: "known_ys, known_xs" },
  Nye = { family: "text", signature: "findText, withinText, [startNum]" },
  Rye = { family: "text", signature: "params1, [params2], …" },
  Iye = { family: "local", signature: "number" },
  _ye = { family: "local", signature: "number" },
  Dye = { family: "date", signature: "serialOrString" },
  Oye = { family: "local", signature: "rows, [columns], [start], [step]" },
  Mye = { family: "math", signature: "x, n, m, coefficients" },
  jye = { family: "local", signature: "number" },
  Fye = { family: "local", signature: "number" },
  Bye = { family: "local", signature: "number" },
  Lye = { family: "statistical", signature: "number1, [number2], …" },
  Pye = { family: "local", signature: "cost, salvage, life" },
  Uye = { family: "distribution", signature: "known_ys, known_xs" },
  $ye = { family: "statistical", signature: "array, k" },
  zye = {
    family: "local",
    signature: "array, [sort_index], [sort_order], [by_col]",
  },
  Vye = { family: "local", signature: "array, by_array1, [sort_order1], …" },
  Wye = { family: "local", signature: "number" },
  Hye = { family: "local", signature: "number" },
  qye = { family: "distribution", signature: "x, mean, standard_dev" },
  Xye = { family: "statistical", signature: "value1, [value2, ...]" },
  Gye = { family: "statistical", signature: "value1, [value2, ...]" },
  Zye = { family: "distribution", signature: "known_ys, known_xs" },
  Kye = {
    family: "text",
    signature: "text, old_text, new_text, [instance_num]",
  },
  Yye = { family: "math", signature: "function_num, ref1, [ref2], …" },
  Jye = { family: "math", signature: "number1, [number2], …" },
  Qye = { family: "math", signature: "range, criteria, [sum_range]" },
  ewe = {
    family: "math",
    signature: "sum_range, criteria_range1, criteria1, …",
  },
  twe = { family: "math", signature: "[array1], arrays1, [arrays2], …" },
  rwe = { family: "math", signature: "params1, [params2], …" },
  nwe = { family: "math", signature: "arrayX, arrayY" },
  iwe = { family: "math", signature: "arrayX, arrayY" },
  owe = { family: "math", signature: "arrayX, arrayY" },
  awe = {
    family: "logical",
    signature: "expression, value1, result1, …, [default]",
  },
  swe = { family: "local", signature: "cost, salvage, life, per" },
  lwe = { family: "text", signature: "value" },
  uwe = { family: "local", signature: "array, rows, [columns]" },
  cwe = { family: "local", signature: "number" },
  dwe = { family: "local", signature: "number" },
  fwe = { family: "text", signature: "value, format_text" },
  hwe = {
    family: "local",
    signature:
      "text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]",
  },
  pwe = {
    family: "local",
    signature:
      "text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]",
  },
  mwe = { family: "text", signature: "delimiter, ignore_empty, text1, …" },
  gwe = {
    family: "local",
    signature:
      "text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]",
  },
  vwe = { family: "date", signature: "hour, minute, second" },
  bwe = { family: "date", signature: "timeText" },
  ywe = { family: "local", signature: "array, [ignore], [scan_by_column]" },
  wwe = { family: "date", signature: "" },
  xwe = { family: "local", signature: "array, [ignore], [scan_by_column]" },
  Ewe = { family: "reference", signature: "array" },
  Cwe = {
    family: "regression",
    signature: "known_ys, [known_xs], [new_xs], [const]",
  },
  Awe = { family: "text", signature: "text" },
  kwe = { family: "statistical", signature: "array, percent" },
  Swe = { family: "logical", signature: "" },
  Twe = { family: "math", signature: "number" },
  Nwe = { family: "local", signature: "value" },
  Rwe = { family: "text", signature: "number" },
  Iwe = { family: "text", signature: "text" },
  _we = { family: "local", signature: "array, [by_col], [exactly_once]" },
  Dwe = { family: "local", signature: "text" },
  Owe = { family: "local", signature: "text" },
  Mwe = { family: "statistical", signature: "value1, [value2, ...]" },
  jwe = { family: "statistical", signature: "value1, [value2, ...]" },
  Fwe = {
    family: "reference",
    signature: "lookup_value, table_array, col_index_num, [range_lookup]",
  },
  Bwe = { family: "local", signature: "array1, [array2], …" },
  Lwe = { family: "date", signature: "serialOrString, [returnType]" },
  Pwe = { family: "date", signature: "serialOrString, [returnType]" },
  Uwe = { family: "date", signature: "startDate, days, holidays" },
  $we = { family: "local", signature: "values, dates, [guess]" },
  zwe = {
    family: "local",
    signature:
      "lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]",
  },
  Vwe = {
    family: "local",
    signature: "lookup_value, lookup_array, [match_mode], [search_mode]",
  },
  Wwe = { family: "local", signature: "rate, values, dates" },
  Hwe = { family: "logical", signature: "params1, [params2], …" },
  qwe = { family: "date", signature: "serialOrString" },
  Xwe = { family: "date", signature: "startDate, endDate, [basis]" },
  Gwe = {
    family: "financial",
    signature: "settlement, maturity, rate, pr, redemption, frequency, [basis]",
  },
  dW = {
    ABS: Ume,
    ACCRINT: $me,
    ACCRINTM: zme,
    ACOS: Vme,
    ACOSH: Wme,
    ACOT: Hme,
    ACOTH: qme,
    ADDRESS: Xme,
    AGGREGATE: Gme,
    AND: Zme,
    ARABIC: Kme,
    AREAS: Yme,
    ASC: Jme,
    ASIN: Qme,
    ASINH: ege,
    ATAN: tge,
    ATAN2: rge,
    ATANH: nge,
    AVEDEV: ige,
    AVERAGE: oge,
    AVERAGEA: age,
    AVERAGEIF: sge,
    AVERAGEIFS: lge,
    BAHTTEXT: uge,
    BASE: cge,
    BESSELI: dge,
    BESSELJ: fge,
    BESSELK: hge,
    BESSELY: pge,
    "BETA.DIST": {
      family: "distribution",
      signature: "x, alpha, beta, cumulative, [a], [b]",
    },
    "BETA.INV": {
      family: "distribution",
      signature: "probability, alpha, beta, [a], [b]",
    },
    BIN2DEC: mge,
    BIN2HEX: gge,
    BIN2OCT: vge,
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
    BITAND: bge,
    BITLSHIFT: yge,
    BITOR: wge,
    BITRSHIFT: xge,
    BITXOR: Ege,
    CEILING: Cge,
    "CEILING.MATH": {
      family: "math",
      signature: "number, [significance], [mode]",
    },
    "CEILING.PRECISE": { family: "math", signature: "number, [significance]" },
    CHAR: Age,
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
    CHOOSE: kge,
    CHOOSECOLS: Sge,
    CHOOSEROWS: Tge,
    CLEAN: Nge,
    CODE: Rge,
    COLUMN: Ige,
    COLUMNS: _ge,
    COMBIN: Dge,
    COMBINA: Oge,
    COMPLEX: Mge,
    CONCAT: jge,
    CONCATENATE: Fge,
    "CONFIDENCE.NORM": {
      family: "distribution",
      signature: "alpha, std, size",
    },
    "CONFIDENCE.T": { family: "distribution", signature: "alpha, std, size" },
    CONVERT: Bge,
    CORREL: Lge,
    COS: Pge,
    COSH: Uge,
    COT: $ge,
    COTH: zge,
    COUNT: Vge,
    COUNTA: Wge,
    COUNTBLANK: Hge,
    COUNTIF: qge,
    COUNTIFS: Xge,
    "COVARIANCE.P": { family: "distribution", signature: "[array1], [array2]" },
    "COVARIANCE.S": { family: "distribution", signature: "[array1], [array2]" },
    CSC: Gge,
    CSCH: Zge,
    DATE: Kge,
    DATEDIF: Yge,
    DATEVALUE: Jge,
    DAVERAGE: Qge,
    DAY: eve,
    DAYS: tve,
    DAYS360: rve,
    DB: nve,
    DBCS: ive,
    DCOUNT: ove,
    DCOUNTA: ave,
    DDB: sve,
    DEC2BIN: lve,
    DEC2HEX: uve,
    DEC2OCT: cve,
    DECIMAL: dve,
    DEGREES: fve,
    DELTA: hve,
    DEVSQ: pve,
    DGET: mve,
    DMAX: gve,
    DMIN: vve,
    DOLLAR: bve,
    DPRODUCT: yve,
    DROP: wve,
    DSTDEV: xve,
    DSTDEVP: Eve,
    DSUM: Cve,
    DVAR: Ave,
    DVARP: kve,
    EDATE: Sve,
    ENCODEURL: Tve,
    EOMONTH: Nve,
    ERF: Rve,
    ERFC: Ive,
    "ERROR.TYPE": { family: "information", signature: "value" },
    EVEN: _ve,
    EXACT: Dve,
    EXP: Ove,
    EXPAND: Mve,
    "EXPON.DIST": {
      family: "distribution",
      signature: "x, lambda, cumulative",
    },
    "F.DIST": { family: "distribution", signature: "x, d1, d2, cumulative" },
    "F.DIST.RT": { family: "distribution", signature: "x, d1, d2" },
    "F.INV": { family: "distribution", signature: "probability, d1, d2" },
    "F.INV.RT": { family: "distribution", signature: "probability, d1, d2" },
    "F.TEST": { family: "distribution", signature: "[array1], [array2]" },
    FACT: jve,
    FACTDOUBLE: Fve,
    FALSE: Bve,
    FILTER: Lve,
    FIND: Pve,
    FINDB: Uve,
    FISHER: $ve,
    FISHERINV: zve,
    FIXED: Vve,
    FLOOR: Wve,
    "FLOOR.MATH": {
      family: "math",
      signature: "number, [significance], [mode]",
    },
    "FLOOR.PRECISE": { family: "math", signature: "number, [significance]" },
    FORECAST: Hve,
    "FORECAST.LINEAR": {
      family: "distribution",
      signature: "params1, [params2], …",
    },
    FORMULATEXT: qve,
    FREQUENCY: Xve,
    FV: Gve,
    GAMMA: Zve,
    "GAMMA.DIST": {
      family: "distribution",
      signature: "x, alpha, beta, cumulative",
    },
    "GAMMA.INV": {
      family: "distribution",
      signature: "probability, alpha, beta",
    },
    GAMMALN: Kve,
    "GAMMALN.PRECISE": { family: "distribution", signature: "x" },
    GAUSS: Yve,
    GCD: Jve,
    GEOMEAN: Qve,
    GESTEP: e1e,
    GROWTH: t1e,
    HARMEAN: r1e,
    HEX2BIN: n1e,
    HEX2DEC: i1e,
    HEX2OCT: o1e,
    HLOOKUP: a1e,
    HOUR: s1e,
    HSTACK: l1e,
    "HYPGEOM.DIST": {
      family: "distribution",
      signature:
        "sample_s, number_sample, population_s, number_pop, cumulative",
    },
    IF: u1e,
    IFERROR: c1e,
    IFNA: d1e,
    IFS: f1e,
    IMABS: h1e,
    IMAGINARY: p1e,
    IMARGUMENT: m1e,
    IMCONJUGATE: g1e,
    IMCOS: v1e,
    IMCOSH: b1e,
    IMCOT: y1e,
    IMCSC: w1e,
    IMCSCH: x1e,
    IMDIV: E1e,
    IMEXP: C1e,
    IMLN: A1e,
    IMLOG10: k1e,
    IMLOG2: S1e,
    IMPOWER: T1e,
    IMPRODUCT: N1e,
    IMREAL: R1e,
    IMSEC: I1e,
    IMSECH: _1e,
    IMSIN: D1e,
    IMSINH: O1e,
    IMSQRT: M1e,
    IMSUB: j1e,
    IMSUM: F1e,
    IMTAN: B1e,
    INDEX: L1e,
    INT: P1e,
    INTERCEPT: U1e,
    IPMT: $1e,
    IRR: z1e,
    ISBLANK: V1e,
    ISERR: W1e,
    ISERROR: H1e,
    ISEVEN: q1e,
    ISFORMULA: X1e,
    ISLOGICAL: G1e,
    ISNA: Z1e,
    ISNONTEXT: K1e,
    ISNUMBER: Y1e,
    "ISO.CEILING": { family: "math", signature: "params1, [params2], …" },
    ISOWEEKNUM: J1e,
    ISREF: Q1e,
    ISTEXT: ebe,
    KURT: tbe,
    LAMBDA: rbe,
    LARGE: nbe,
    LCM: ibe,
    LEFT: obe,
    LEFTB: abe,
    LEN: sbe,
    LET: lbe,
    LINEST: ube,
    LN: cbe,
    LOG: dbe,
    LOG10: fbe,
    LOGEST: hbe,
    "LOGNORM.DIST": {
      family: "distribution",
      signature: "x, mean, standard_dev, cumulative",
    },
    "LOGNORM.INV": {
      family: "distribution",
      signature: "probability, mean, standard_dev",
    },
    LOOKUP: pbe,
    LOWER: mbe,
    MATCH: gbe,
    MAX: vbe,
    MAXA: bbe,
    MAXIFS: ybe,
    MDETERM: wbe,
    MEDIAN: xbe,
    MID: Ebe,
    MIDB: Cbe,
    MIN: Abe,
    MINA: kbe,
    MINIFS: Sbe,
    MINUTE: Tbe,
    MINVERSE: Nbe,
    MIRR: Rbe,
    MMULT: Ibe,
    MOD: _be,
    MODE: Dbe,
    "MODE.MULT": { family: "statistical", signature: "number1, [number2], …" },
    "MODE.SNGL": { family: "distribution", signature: "number1, [number2], …" },
    MONTH: Obe,
    MROUND: Mbe,
    MULTINOMIAL: jbe,
    MUNIT: Fbe,
    N: Bbe,
    NA: Lbe,
    "NEGBINOM.DIST": {
      family: "distribution",
      signature: "number_f, number_s, probability_s, cumulative",
    },
    NETWORKDAYS: Pbe,
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
    NOT: Ube,
    NOW: $be,
    NPER: zbe,
    NPV: Vbe,
    NUMBERVALUE: Wbe,
    OCT2BIN: Hbe,
    OCT2DEC: qbe,
    OCT2HEX: Xbe,
    ODD: Gbe,
    OR: Zbe,
    PEARSON: Kbe,
    PERCENTILE: Ybe,
    "PERCENTILE.EXC": { family: "distribution", signature: "array, k" },
    "PERCENTILE.INC": { family: "distribution", signature: "array, k" },
    PERMUT: Jbe,
    PERMUTATIONA: Qbe,
    PHI: eye,
    PI: tye,
    PMT: rye,
    "POISSON.DIST": {
      family: "distribution",
      signature: "x, mean, cumulative",
    },
    POWER: nye,
    PPMT: iye,
    PRICE: oye,
    PRODUCT: aye,
    PROPER: sye,
    PV: lye,
    QUARTILE: uye,
    "QUARTILE.EXC": { family: "distribution", signature: "array, quart" },
    "QUARTILE.INC": { family: "distribution", signature: "array, quart" },
    QUOTIENT: cye,
    RADIANS: dye,
    RAND: fye,
    RANDBETWEEN: hye,
    RANK: pye,
    "RANK.AVG": { family: "distribution", signature: "number, ref, [order]" },
    "RANK.EQ": { family: "distribution", signature: "number, ref, [order]" },
    RATE: mye,
    REPLACE: gye,
    REPLACEB: vye,
    REPT: bye,
    RIGHT: yye,
    RIGHTB: wye,
    ROMAN: xye,
    ROUND: Eye,
    ROUNDDOWN: Cye,
    ROUNDUP: Aye,
    ROW: kye,
    ROWS: Sye,
    RSQ: Tye,
    SEARCH: Nye,
    SEARCHB: Rye,
    SEC: Iye,
    SECH: _ye,
    SECOND: Dye,
    SEQUENCE: Oye,
    SERIESSUM: Mye,
    SIGN: jye,
    SIN: Fye,
    SINH: Bye,
    SKEW: Lye,
    "SKEW.P": { family: "statistical", signature: "number1, [number2], …" },
    SLN: Pye,
    SLOPE: Uye,
    SMALL: $ye,
    SORT: zye,
    SORTBY: Vye,
    SQRT: Wye,
    SQRTPI: Hye,
    STANDARDIZE: qye,
    "STDEV.P": { family: "distribution", signature: "number1, [number2], …" },
    "STDEV.S": { family: "distribution", signature: "number1, [number2], …" },
    STDEVA: Xye,
    STDEVPA: Gye,
    STEYX: Zye,
    SUBSTITUTE: Kye,
    SUBTOTAL: Yye,
    SUM: Jye,
    SUMIF: Qye,
    SUMIFS: ewe,
    SUMPRODUCT: twe,
    SUMSQ: rwe,
    SUMX2MY2: nwe,
    SUMX2PY2: iwe,
    SUMXMY2: owe,
    SWITCH: awe,
    SYD: swe,
    T: lwe,
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
    TAKE: uwe,
    TAN: cwe,
    TANH: dwe,
    TEXT: fwe,
    TEXTAFTER: hwe,
    TEXTBEFORE: pwe,
    TEXTJOIN: mwe,
    TEXTSPLIT: gwe,
    TIME: vwe,
    TIMEVALUE: bwe,
    TOCOL: ywe,
    TODAY: wwe,
    TOROW: xwe,
    TRANSPOSE: Ewe,
    TREND: Cwe,
    TRIM: Awe,
    TRIMMEAN: kwe,
    TRUE: Swe,
    TRUNC: Twe,
    TYPE: Nwe,
    UNICHAR: Rwe,
    UNICODE: Iwe,
    UNIQUE: _we,
    UPPER: Dwe,
    VALUE: Owe,
    "VAR.P": { family: "distribution", signature: "number1, [number2], …" },
    "VAR.S": { family: "distribution", signature: "number1, [number2], …" },
    VARA: Mwe,
    VARPA: jwe,
    VLOOKUP: Fwe,
    VSTACK: Bwe,
    WEEKDAY: Lwe,
    WEEKNUM: Pwe,
    "WEIBULL.DIST": {
      family: "distribution",
      signature: "x, alpha, beta, cumulative",
    },
    WORKDAY: Uwe,
    "WORKDAY.INTL": {
      family: "date",
      signature: "startDate, days, weekend, holidays",
    },
    XIRR: $we,
    XLOOKUP: zwe,
    XMATCH: Vwe,
    XNPV: Wwe,
    XOR: Hwe,
    YEAR: qwe,
    YEARFRAC: Xwe,
    YIELD: Gwe,
    "Z.TEST": { family: "distribution", signature: "array, x, [sigma]" },
  },
  Zwe = new Set(Fc),
  S3 = (e, r) => r.includes(e),
  fW = [
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
  Kwe = [
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
  Ywe = ["MINVERSE", "MDETERM", "MMULT", "MUNIT"],
  Jwe = [
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
  zT = Object.freeze({
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
function Qwe(e) {
  var t;
  if (((e = e.toUpperCase()), S3(e, AR))) return "database";
  if (Kwe.includes(e)) return "regression";
  if (Ywe.includes(e)) return "matrix";
  if (Jwe.includes(e)) return "financial";
  if (e === "CONVERT") return "engineering";
  if (S3(e, QN)) return "statistics";
  if (zT[e]) return zT[e];
  if (e.startsWith("CUBE")) return "cube";
  const r = (t = dW[e]) == null ? void 0 : t.family;
  return r === "statistical"
    ? "statistics"
    : r === "distribution" || r === "engineering" || r === "financial"
      ? r
      : "other";
}
const exe = {
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
function d_(e) {
  var a;
  const r = e.replace(/^_xlfn\.(?:_xlws\.)?/i, "").toUpperCase(),
    t = Qwe(r),
    n = exe[t],
    o = Zwe.has(r);
  return {
    name: r,
    family: t,
    status: o
      ? "local"
      : t === "cube" || t === "network"
        ? "connector-required"
        : "unimplemented",
    inputs:
      ((a = dW[r]) == null ? void 0 : a.signature) ??
      "Not adopted; consult the source function or add-in contract.",
    outputs: o ? ($T[r] ?? n.output) : "No local implementation",
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
      (S3(r, VS)
        ? "Gregorian 1900 serials 0–2958465 (compatibility serial 60) or 1904 serials 0–2957003; full-year ISO/configured MDY or DMY text; recorded IANA-zone clock shared per calculation. Up to 1000 holiday inputs, 366000 business-day offsets/date spans. The date system persists in native JSON and XLSX; text order and time zone are native JSON only. DATEDIF MD/YD and US day-count edge conventions follow the documented local policy; independent Excel parity unverified."
        : n.limits) +
      (o
        ? " Workbook work and accounted-memory budgets apply across a run; synchronous APIs use cooperative time checks and isolated worker execution supports deadline termination/cancellation."
        : "") +
      (["RAND", "RANDBETWEEN"].includes(r)
        ? " Unseeded runtime pseudorandom values are stable within the revision calculation and refresh on recalculation. They are not reproducible audit samples. RANDBETWEEN uses ceil(bottom)/floor(top) and refuses bounds or inclusive spans outside safe integers."
        : ""),
    evidence: o
      ? S3(r, VS)
        ? [
            "tests/date-calendar.test.ts",
            "tests/date-calendar-integration.test.ts",
            "docs/DATE-CALENDAR-ACCEPTANCE.md",
          ]
        : n.tests.split(";")
      : [],
  };
}
function hW() {
  return [...new Set([...Fc, ...Object.keys(zT)])].sort().map(d_);
}
function txe() {
  const e = hW();
  return fW.map((r) => ({
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
function rxe({
  mode: e,
  workbook: r,
  sheetId: t,
  range: n,
  readOnly: o,
  onClose: a,
  onWorkbookChange: s,
  onInsertFormula: d,
}) {
  var de, he, Se, le, ce, ge;
  const { text: u, language: l } = Dr(),
    c = r.sheets.find((se) => se.id === t),
    [h, f] = ne.useState(""),
    [p, m] = ne.useState("SUM"),
    [v, g] = ne.useState(n),
    [b, y] = ne.useState(""),
    [w, E] = ne.useState(
      `='${((c == null ? void 0 : c.name) ?? "Sheet1").replace(/'/g, "''")}'!${n}`,
    ),
    [A, k] = ne.useState(""),
    [F, I] = ne.useState(""),
    [O, j] = ne.useState(null),
    [z, V] = ne.useState(""),
    [D, B] = ne.useState(JSON.stringify(r)),
    [_, C] = ne.useState(false),
    [T, N] = ne.useState(null),
    [R, L] = ne.useState("100,0.2"),
    [$, U] = ne.useState(""),
    G = $p(w),
    H = (se, we) => {
      (_ || B(JSON.stringify(r)), C(true), N(null), U(""), se(we));
    },
    Q = o || !c || !!((de = r.structureProtection) != null && de.enabled),
    X$1 = (se) =>
      V(se instanceof Error ? se.message : "The name could not be saved."),
    J = ne.useMemo(
      () => Fc.filter((se) => se.toLowerCase().includes(h.toLowerCase())),
      [h],
    ),
    q = `=${p}(${v})`,
    W = d_(p),
    K = ne.useMemo(() => {
      if (e !== "function") return "";
      const se = n.split(":")[0].replace(/\$/g, "").toUpperCase(),
        we = {
          ...r,
          sheets: r.sheets.map((Ee) =>
            Ee.id === t
              ? { ...Ee, cells: { ...Ee.cells, [se]: { raw: q } } }
              : Ee,
          ),
        };
      try {
        return String(Vn(we)(t, se) ?? "");
      } catch {
        return "Check the function arguments.";
      }
    }, [q, e, r, t, n]),
    te = () => {
      (y(""),
        I(""),
        j(null),
        V(""),
        U(""),
        C(false),
        N(null),
        B(JSON.stringify(r)),
        k(""));
    },
    ae = () => {
      var se;
      try {
        if (Q) throw Error("This workbook is read-only or protected.");
        if (_ && D !== JSON.stringify(r))
          throw Error(
            "The workbook changed while this name was being edited. Choose New name or reopen Edit to review current content.",
          );
        const we = {
            name: b.trim(),
            refersTo: w.trim(),
            ...(A ? { scopeSheetId: A } : {}),
            ...(F ? { comment: F } : {}),
          },
          Ee = O === null || (se = r.definedNames) == null ? void 0 : se[O];
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
            Ty(
              r,
              {
                kind: "rename",
                name: Ee.name,
                ...(Ee.scopeSheetId ? { scopeSheetId: Ee.scopeSheetId } : {}),
                newName: we.name,
              },
              { readOnly: o },
            ),
          ),
            V(""));
          return;
        }
        const Ae = Ty(
            r,
            Ee
              ? {
                  kind: "update",
                  name: Ee.name,
                  ...(Ee.scopeSheetId ? { scopeSheetId: Ee.scopeSheetId } : {}),
                  definition: we,
                }
              : { kind: "create", definition: we },
            { readOnly: o },
          ),
          ye = q6(r, Ae, { readOnly: o });
        (ye !== r && s(ye, `${Ee ? "Edit" : "Define"} name: ${we.name}`), te());
      } catch (we) {
        X$1(we);
      }
    },
    oe = () => {
      try {
        const se = n.split(":")[0].replace(/\$/g, "").toUpperCase(),
          we = structuredClone(r),
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
        const Ae = Ty(we, { kind: "create", definition: Ee }).next,
          ye = A || t,
          Te = Ae.sheets.find((Me) => Me.id === ye);
        Te.cells[se] = { raw: `=${b.trim()}${G ? `(${R})` : ""}` };
        const Pe = Vn(Ae),
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
      se.target === se.currentTarget && a();
    },
    onKeyDown: (se) => {
      se.key === "Escape" && a();
    },
    children: i.jsxs("section", {
      className: "formula-tools",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": e === "function" ? "Insert Function" : "Name Manager",
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              children: e === "function" ? "Insert Function" : "Name Manager",
            }),
            i.jsx("button", {
              "aria-label": u("Close formula tools"),
              onClick: a,
              lang: l("Close formula tools"),
              "data-ui-label": "Close formula tools",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        e === "function"
          ? i.jsxs(i.Fragment, {
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Search functions" }),
                    i.jsx("input", {
                      autoFocus: true,
                      "aria-label": u("Search functions"),
                      value: h,
                      onChange: (se) => f(se.target.value),
                      placeholder: u("SUM, lookup, text…"),
                      lang: l("Search functions"),
                      "data-ui-label": "Search functions",
                    }),
                  ],
                }),
                i.jsxs("div", {
                  className: "formula-function-layout",
                  children: [
                    i.jsx("select", {
                      "aria-label": u("Available functions"),
                      size: 11,
                      value: p,
                      onChange: (se) => {
                        (m(se.target.value),
                          UT[se.target.value] === "" && g(""));
                      },
                      lang: l("Available functions"),
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
                            i.jsx(M, { children: "Arguments" }),
                            i.jsx("input", {
                              "aria-label": u("Function arguments"),
                              value: v,
                              onChange: (se) => g(se.target.value),
                              readOnly: Q,
                              lang: l("Function arguments"),
                              "data-ui-label": "Function arguments",
                            }),
                          ],
                        }),
                        i.jsx("p", {
                          children: i.jsx(M, {
                            children:
                              "Separate arguments with commas. Put text inside double quotes. Use cell ranges, defined names or table columns.",
                          }),
                        }),
                        p === "SUBTOTAL" &&
                          i.jsx("p", {
                            children: i.jsx(M, {
                              children:
                                "Function numbers 101–111 exclude hidden and collapsed outline rows. The quick text filter is view-only and does not change calculations. An external host can supply a separate filtered-row context.",
                            }),
                          }),
                        i.jsxs("output", {
                          "aria-label": u("Formula preview"),
                          lang: l("Formula preview"),
                          "data-ui-label": "Formula preview",
                          children: [
                            i.jsx("code", { children: q }),
                            i.jsx("strong", { children: K }),
                          ],
                        }),
                        i.jsxs("small", {
                          children: [
                            Fc.length,
                            " ",
                            i.jsx(M, {
                              children:
                                "available functions. Compatibility and calculation limits are documented with the package.",
                            }),
                          ],
                        }),
                        i.jsxs("details", {
                          children: [
                            i.jsx("summary", {
                              children: i.jsx(M, {
                                children: "Function support and limits",
                              }),
                            }),
                            i.jsx("p", { children: W.outputs }),
                            i.jsxs("p", {
                              children: [
                                i.jsx(M, { children: "Errors: " }),
                                W.errors.join(", "),
                              ],
                            }),
                            i.jsx("p", { children: W.limits }),
                            i.jsx("p", {
                              children: i.jsx(M, {
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
                      onClick: a,
                      children: i.jsx(M, { children: "Cancel" }),
                    }),
                    i.jsx("button", {
                      className: "formula-primary",
                      disabled: Q,
                      onClick: () => {
                        Q || (d(q), a());
                      },
                      children: i.jsx(M, { children: "Insert formula" }),
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
                                children: i.jsx(M, { children: "Name" }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(M, { children: "Refers to" }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(M, { children: "Scope" }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(M, { children: "Actions" }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: (r.definedNames ?? []).map((se, we) => {
                            var Ee;
                            return i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("td", {
                                    children: [
                                      se.name,
                                      $p(se.refersTo) &&
                                        i.jsxs("small", {
                                          children: [
                                            " ",
                                            i.jsx(M, {
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
                                      ? (Ee = r.sheets.find(
                                          (Ae) => Ae.id === se.scopeSheetId,
                                        )) == null
                                        ? void 0
                                        : Ee.name
                                      : "Workbook",
                                  }),
                                  i.jsxs("td", {
                                    children: [
                                      i.jsx("button", {
                                        disabled: Q,
                                        onClick: () => {
                                          (j(we),
                                            B(JSON.stringify(r)),
                                            C(true),
                                            N(null),
                                            V(""),
                                            U(""),
                                            y(se.name),
                                            E(se.refersTo),
                                            k(se.scopeSheetId ?? ""),
                                            I(se.comment ?? ""));
                                        },
                                        children: i.jsx(M, {
                                          children: "Edit",
                                        }),
                                      }),
                                      i.jsx("button", {
                                        disabled: Q,
                                        "aria-label": `Delete name ${se.name}`,
                                        onClick: () => {
                                          try {
                                            const Ae = Ty(
                                                r,
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
                                              ye = q6(r, Ae, { readOnly: Q });
                                            (ye !== r &&
                                              s(ye, `Delete name: ${se.name}`),
                                              te());
                                          } catch (Ae) {
                                            X$1(Ae);
                                          }
                                        },
                                        children: i.jsx(M, {
                                          children: "Delete",
                                        }),
                                      }),
                                    ],
                                  }),
                                ],
                              },
                              `${se.scopeSheetId}:${se.name}`,
                            );
                          }),
                        }),
                      ],
                    }),
                    !((he = r.definedNames) != null && he.length) &&
                      i.jsx("p", {
                        children: i.jsx(M, {
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
                        i.jsx(M, { children: "Name" }),
                        i.jsx("input", {
                          autoFocus: true,
                          "aria-label": u("Defined name"),
                          value: b,
                          onChange: (se) => H(y, se.target.value),
                          readOnly: Q,
                          lang: l("Defined name"),
                          "data-ui-label": "Defined name",
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Scope" }),
                        i.jsxs("select", {
                          "aria-label": u("Name scope"),
                          value: A,
                          onChange: (se) => H(k, se.target.value),
                          disabled: Q || O !== null,
                          lang: l("Name scope"),
                          "data-ui-label": "Name scope",
                          children: [
                            i.jsx("option", {
                              value: "",
                              children: "Workbook",
                            }),
                            r.sheets.map((se) =>
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
                        i.jsx(M, { children: "Refers to" }),
                        i.jsx("input", {
                          "aria-label": u("Name refers to"),
                          value: w,
                          onChange: (se) => H(E, se.target.value),
                          readOnly: Q,
                          lang: l("Name refers to"),
                          "data-ui-label": "Name refers to",
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      className: "formula-wide",
                      children: [
                        i.jsx(M, { children: "Comment" }),
                        i.jsx("input", {
                          "aria-label": u("Name comment"),
                          value: F,
                          onChange: (se) => H(I, se.target.value),
                          readOnly: Q,
                          lang: l("Name comment"),
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
                      children: i.jsx(M, { children: "Use LAMBDA template" }),
                    }),
                    G &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Test arguments" }),
                          i.jsx("input", {
                            "aria-label": u("Name test arguments"),
                            value: R,
                            onChange: (se) => {
                              (L(se.target.value), U(""));
                            },
                            lang: l("Name test arguments"),
                            "data-ui-label": "Name test arguments",
                          }),
                        ],
                      }),
                    i.jsx("button", {
                      disabled: !b.trim(),
                      onClick: oe,
                      children: i.jsx(M, { children: "Test definition" }),
                    }),
                  ],
                }),
                $ &&
                  i.jsxs("p", {
                    role: "status",
                    "aria-label": u("Name test result"),
                    lang: l("Name test result"),
                    "data-ui-label": "Name test result",
                    children: [
                      $,
                      " ",
                      i.jsx(M, {
                        children:
                          "· Preview only; the workbook has not changed.",
                      }),
                    ],
                  }),
                T &&
                  i.jsxs("section", {
                    "aria-label": u("Name migration preview"),
                    lang: l("Name migration preview"),
                    "data-ui-label": "Name migration preview",
                    children: [
                      i.jsxs("h3", {
                        children: [
                          i.jsx(M, { children: "Rename " }),
                          (Se = T.before) == null ? void 0 : Se.name,
                          " ",
                          i.jsx(M, { children: "to " }),
                          (le = T.after) == null ? void 0 : le.name,
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          T.changes.length,
                          " ",
                          i.jsx(M, {
                            children:
                              "formulas or definitions will be rewritten. ",
                          }),
                          T.dependencies.length,
                          " ",
                          i.jsx(M, {
                            children:
                              "dependent locations were found. Values, formatting and review annotations are retained. Apply commits the reviewed changes together.",
                          }),
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx("strong", {
                            children: i.jsx(M, { children: "Definition:" }),
                          }),
                          " ",
                          i.jsx("code", {
                            children:
                              (ce = T.before) == null ? void 0 : ce.refersTo,
                          }),
                          " →",
                          " ",
                          i.jsx("code", {
                            children:
                              (ge = T.after) == null ? void 0 : ge.refersTo,
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
                                    children: i.jsx(M, {
                                      children: "Location",
                                    }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(M, { children: "Before" }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(M, { children: "After" }),
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
                            children: i.jsx(M, {
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
                        children: i.jsx(M, { children: "Cancel rename" }),
                      }),
                      i.jsx("button", {
                        disabled: Q,
                        onClick: () => {
                          var se, we;
                          try {
                            const Ee = q6(r, T, { readOnly: Q });
                            (Ee !== r &&
                              s(
                                Ee,
                                `Rename name: ${(se = T.before) == null ? void 0 : se.name} → ${(we = T.after) == null ? void 0 : we.name}`,
                              ),
                              te());
                          } catch (Ee) {
                            X$1(Ee);
                          }
                        },
                        children: i.jsx(M, { children: "Apply rename" }),
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
                        children: i.jsx(M, { children: "New name" }),
                      }),
                    i.jsx("button", {
                      onClick: a,
                      children: i.jsx(M, { children: "Close" }),
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
const dS = Object.freeze({
  maxLocations: 1e4,
  maxCharacters: 2e6,
  maxFindings: 2e3,
  maxNameDepth: 32,
});

export { dS, d_, fW, hW, rxe, txe };
