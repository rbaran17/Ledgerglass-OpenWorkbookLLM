import { buildCompositeKey } from "./11-local-workbook-store-and-versions.js";

import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import {
  assertWorkbookStructureEditable,
  parseAnalysisRequest,
  FormulaParser,
  createWorksheetRecord,
  sumDecimalPartsToNumber,
  exactAverageOfNumbers,
  _t,
  parseRangeAddress,
  analysisLimits,
  evaluateRegressionFunction,
  analysisFormatVersion,
  validateDataValidationRule,
  resolveDefinedNameValue,
  normalizeWorkbookInput,
  columnIndexToLetters,
} from "./03-workbook-model.js";

import {
  readSourceHeaders,
  inspectOutputNameSafety,
} from "./12-consolidation-and-dependency-inspection.js";
import {
  assertSolverMagnitude,
  parseRangeBounds,
  assertAutomaticCalculation,
  normalizeAnalysisPermissions,
  rejectRangeAsScalar,
  p_,
  validateAnalysisSource,
} from "./36-watch-window-and-goal-seek.js";

function linearizeSolverFormula(workbook, analysisRequest, t) {
  assertAutomaticCalculation(workbook);
  const n = analysisRequest.variables.length,
    o = new Map(),
    a = new Set();
  let s = 0;
  const d = (y) => ({
      constant: assertSolverMagnitude(y),
      coefficients: Array(n).fill(0),
    }),
    u = (y, w) => ({
      constant: assertSolverMagnitude(y.constant * w),
      coefficients: y.coefficients.map((coefficient) =>
        assertSolverMagnitude(coefficient * w),
      ),
    }),
    l = (y, w) => ({
      constant: assertSolverMagnitude(y.constant + w.constant),
      coefficients: y.coefficients.map((coefficient, A) =>
        assertSolverMagnitude(coefficient + w.coefficients[A]),
      ),
    }),
    c = (y, w) => {
      if (y.coefficients.some(Boolean) && w.coefficients.some(Boolean))
        throw Error(
          "Nonlinear multiplication is outside the linear Solver subset.",
        );
      return y.coefficients.some(Boolean) ? u(y, w.constant) : u(w, y.constant);
    },
    h = (sheetId, address, E = 0) => {
      const compositeKey = buildCompositeKey(sheetId, address);
      if (a.has(compositeKey)) throw Error("Solver formulas contain a cycle.");
      if (o.has(compositeKey)) return o.get(compositeKey);
      if (
        o.size + a.size >= analysisLimits.dependencyCells ||
        E > analysisLimits.depth
      )
        throw Error("Solver dependency budget exceeded.");
      const k = validateAnalysisSource(
          workbook,
          { sheetId: sheetId, address: address },
          t,
        ),
        cell = k.cells[address],
        I = analysisRequest.variables.findIndex(
          (variable) =>
            variable.cell.sheetId === sheetId &&
            variable.cell.address === address,
        );
      if (I >= 0) {
        const j = d(0);
        return ((j.coefficients[I] = 1), o.set(compositeKey, j), j);
      }
      if (!cell || cell.raw === null)
        throw Error(
          "Solver dependencies must be populated numeric cells; blanks and spill children are excluded.",
        );
      let O;
      if (typeof cell.raw == "number") O = d(cell.raw);
      else if (
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal
      ) {
        if (cell.raw.length > analysisLimits.expressionCharacters)
          throw Error("Solver expression budget exceeded.");
        const j = resolveDefinedNameValue(cell.raw, workbook, sheetId).slice(1);
        if (j.length > 4096)
          throw Error("Expanded Solver expression budget exceeded.");
        (a.add(compositeKey),
          (O = rejectRangeAsScalar(f(j, sheetId, E + 1))),
          a.delete(compositeKey));
      } else
        throw Error(
          "Solver dependencies must contain numbers or supported linear formulas.",
        );
      return (o.set(compositeKey, O), O);
    },
    f = (y, w, E) => {
      const A =
          /\s*(?:(\d*\.\d+(?:[Ee][+-]?\d+)?|\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?)|((?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?)(?![A-Za-z0-9_])|([A-Za-z_][A-Za-z0-9_.]*)|([()+\-*/,:%]))/gy,
        k = [];
      let F = 0;
      for (; F < y.length && y.slice(F).trim();) {
        A.lastIndex = F;
        const _ = A.exec(y);
        if (!_ || _.index !== F)
          throw Error(
            "Unsupported Solver syntax. Use numeric references, arithmetic, SUM, AVERAGE or SUMPRODUCT.",
          );
        if (
          (k.push({
            type: _[1] ? 1 : _[2] ? 2 : _[3] ? 3 : 4,
            text: _[1] ?? _[2] ?? _[3] ?? _[4],
          }),
          (F = A.lastIndex),
          ++s > analysisLimits.syntaxNodes)
        )
          throw Error("Solver syntax budget exceeded.");
      }
      let I = 0,
        O = 0;
      const j = () => k[I]?.text,
        z = () => {
          if (++O > 32)
            throw Error("Solver expression nesting limit exceeded.");
          let _;
          const C = k[I++];
          if (!C) throw Error("Incomplete Solver formula.");
          if (C.text === "+" || C.text === "-")
            _ = u(rejectRangeAsScalar(z()), C.text === "-" ? -1 : 1);
          else if (C.type === 1) _ = d(Number(C.text));
          else if (C.type === 2) {
            const $ = C.text.lastIndexOf("!"),
              U = $ < 0 ? null : C.text.slice(0, $),
              G = C.text
                .slice($ + 1)
                .replace(/\$/g, "")
                .toUpperCase(),
              H =
                U === null
                  ? w
                  : workbook.sheets.find(
                      (sheet) =>
                        sheet.name.toLowerCase() ===
                        (U.startsWith("'")
                          ? U.slice(1, -1).replace(/''/g, "'")
                          : U
                        ).toLowerCase(),
                    )?.id;
            if (!H) throw Error("A Solver worksheet reference is missing.");
            if (G.includes(":")) {
              const rangeAddress = parseRangeAddress(G);
              if (
                (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
                  (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
                500
              )
                throw Error("Solver range exceeds 500 cells.");
              const J = [];
              for (
                let firstRow = rangeAddress.firstRow;
                firstRow <= rangeAddress.lastRow;
                firstRow++
              )
                for (
                  let firstCol = rangeAddress.firstCol;
                  firstCol <= rangeAddress.lastCol;
                  firstCol++
                )
                  J.push(
                    h(H, `${columnIndexToLetters(firstCol)}${firstRow + 1}`, E),
                  );
              _ = Object.assign(J, {
                height: rangeAddress.lastRow - rangeAddress.firstRow + 1,
                width: rangeAddress.lastCol - rangeAddress.firstCol + 1,
              });
            } else _ = h(H, G, E);
          } else if (C.text === "(") {
            if (((_ = D()), k[I++]?.text !== ")"))
              throw Error("Unbalanced Solver parentheses.");
          } else if (C.type === 3) {
            const $ = C.text.toUpperCase();
            if (
              !["SUM", "AVERAGE", "SUMPRODUCT"].includes($) ||
              k[I++]?.text !== "("
            )
              throw Error(`Unsupported Solver function or name: ${C.text}.`);
            const U = [];
            if (j() !== ")") for (U.push(D()); j() === ",";) (I++, U.push(D()));
            if (k[I++]?.text !== ")" || !U.length)
              throw Error("Invalid Solver function arguments.");
            if ($ === "SUMPRODUCT") {
              const G = U.map((H) => (Array.isArray(H) ? H : [H]));
              if (
                G.some(
                  (H) =>
                    H.length !== G[0].length ||
                    (H.height ?? 1) !== (G[0].height ?? 1) ||
                    (H.width ?? 1) !== (G[0].width ?? 1),
                )
              )
                throw Error("SUMPRODUCT ranges must have matching dimensions.");
              _ = G[0].reduce(
                (H, Q, X) =>
                  l(
                    H,
                    G.reduce((J, q) => c(J, q[X]), d(1)),
                  ),
                d(0),
              );
            } else {
              const G = U.flatMap((H) => (Array.isArray(H) ? H : [H]));
              ((_ = G.reduce(l, d(0))),
                $ === "AVERAGE" && (_ = u(_, 1 / G.length)));
            }
          } else throw Error("Unsupported Solver expression.");
          for (; j() === "%";) (I++, (_ = u(rejectRangeAsScalar(_), 0.01)));
          return (O--, _);
        },
        V = () => {
          let _ = z();
          for (; j() === "*" || j() === "/";) {
            const text = k[I++].text,
              T = rejectRangeAsScalar(z());
            if (text === "*") _ = c(rejectRangeAsScalar(_), T);
            else {
              if (T.coefficients.some(Boolean) || T.constant === 0)
                throw Error(
                  "Solver division requires a nonzero constant denominator.",
                );
              _ = u(rejectRangeAsScalar(_), 1 / T.constant);
            }
          }
          return _;
        },
        D = () => {
          let _ = V();
          for (; j() === "+" || j() === "-";) {
            const text = k[I++].text,
              T = rejectRangeAsScalar(V());
            _ = l(rejectRangeAsScalar(_), u(T, text === "-" ? -1 : 1));
          }
          return _;
        },
        B = D();
      if (I !== k.length)
        throw Error("Unsupported or trailing Solver formula syntax.");
      return B;
    },
    baseline = analysisRequest.variables.map((variable) => {
      const w = validateAnalysisSource(workbook, variable.cell, t),
        cell = w.cells[variable.cell.address],
        A = _t(variable.cell.address);
      if (!cell || typeof cell.raw != "number" || !Number.isFinite(cell.raw))
        throw Error(
          "Changing cells must contain finite numeric inputs, not formulas or text.",
        );
      if (w.protection?.enabled)
        throw Error("A changing worksheet is protected.");
      if (
        [...(w.pivots ?? []), ...(w.queries ?? [])].some((I) =>
          parseRangeBounds(I.outputRange, variable.cell.address),
        )
      )
        throw Error("Generated results cannot be changing cells.");
      for (const I of w.tables ?? []) {
        if (!parseRangeBounds(I.range, variable.cell.address)) continue;
        const rangeAddress = parseRangeAddress(I.range);
        if (
          A.row === rangeAddress.firstRow ||
          (I.showTotals && A.row === rangeAddress.lastRow) ||
          I.columns[A.col - rangeAddress.firstCol]?.calculatedFormula ||
          Object.entries(w.cells).some(([j, cell]) => {
            const V = _t(j);
            return (
              V.col === A.col &&
              V.row > rangeAddress.firstRow &&
              V.row <= rangeAddress.lastRow - +!!I.showTotals &&
              !cell.literal &&
              typeof cell.raw == "string" &&
              cell.raw.startsWith("=")
            );
          })
        )
          throw Error(
            "Table headers, totals and calculated columns cannot be changing cells.",
          );
      }
      if (
        (w.validations ?? []).some((I) =>
          parseRangeBounds(I.range, variable.cell.address),
        )
      )
        throw Error(
          "Changing cells with validation rules are excluded. Use explicit Solver bounds and constraints on a reviewed input copy.",
        );
      return cell.raw;
    }),
    m = validateAnalysisSource(workbook, analysisRequest.objective, t),
    cell = m.cells[analysisRequest.objective.address];
  if (
    !cell ||
    typeof cell.raw != "string" ||
    !cell.raw.startsWith("=") ||
    cell.literal
  )
    throw Error("The objective must be a linear formula cell.");
  const objective = h(
      analysisRequest.objective.sheetId,
      analysisRequest.objective.address,
    ),
    constraints = analysisRequest.constraints.map((constraint) => ({
      form: l(
        h(constraint.left.sheetId, constraint.left.address),
        u(
          typeof constraint.right == "number"
            ? d(constraint.right)
            : h(constraint.right.sheetId, constraint.right.address),
          -1,
        ),
      ),
      operator: constraint.operator,
    }));
  return (
    analysisRequest.direction === "target" &&
      constraints.push({
        form: l(objective, d(-analysisRequest.target)),
        operator: "=",
      }),
    {
      objective: objective,
      constraints: constraints,
      baseline: baseline,
      dependencies: o.size,
    }
  );
}
const evaluateLinearExpression = (e, r) =>
  assertSolverMagnitude(
    e.constant + e.coefficients.reduce((t, n, o) => t + n * r[o], 0),
  );
function solveLinearSystem(e, r) {
  const t = r.length,
    n = e.map((o, a) => {
      const s = Math.max(...o.map(Math.abs));
      return [...o.map((d) => d / (s || 1)), r[a] / (s || 1)];
    });
  for (let o = 0; o < t; o++) {
    let a = o;
    for (let u = o + 1; u < t; u++)
      Math.abs(n[u][o]) > Math.abs(n[a][o]) && (a = u);
    const s = Math.abs(n[a][o]);
    if (s < 1e-14) return "singular";
    if (s < 1e-9) return "uncertain";
    [n[a], n[o]] = [n[o], n[a]];
    const d = n[o][o];
    for (let u = o; u <= t; u++) n[o][u] /= d;
    for (let u = 0; u < t; u++)
      if (u !== o) {
        const l = n[u][o];
        for (let c = o; c <= t; c++) n[u][c] -= l * n[o][c];
      }
  }
  return n.map((o) => o[t]);
}
function solveBoundedLinearModel(workbook, analysisRequest, t = {}) {
  const n = linearizeSolverFormula(workbook, analysisRequest, t),
    o = analysisRequest.variables.length,
    a = analysisRequest.variables.every(
      (variable) => variable.domain !== "continuous",
    ),
    s = {
      status: "infeasible",
      success: false,
      message:
        "No feasible solution satisfies the declared bounds and constraints.",
      algorithm: a
        ? "Exhaustive bounded integer enumeration"
        : "Exhaustive bounded linear vertex enumeration with pivoted elimination",
      assumptions: [
        "Symbolically proven linear formulas; explicit finite bounds on every changing cell.",
        "No nonlinear, evolutionary, mixed integer, unrestricted or proprietary Solver engine is used.",
        `Feasibility uses relative tolerance ${analysisRequest.tolerance}; all candidates must be exhausted before an optimal label.`,
        `Dependency cells: ${n.dependencies}. No sample exclusions; all declared constraints participate.`,
        "Independent references: analytic linear fixtures and Python Fraction vertex checks; see ANALYSIS-VERIFICATION.json.",
      ],
      exclusions: [],
      included: o,
      inspected: o,
      iterations: 0,
      rows: [],
      baseline: n.baseline,
    },
    d = [];
  for (const constraint of n.constraints) {
    const b = { a: constraint.form.coefficients, b: -constraint.form.constant };
    (constraint.operator !== ">=" && d.push(b),
      constraint.operator !== "<=" &&
        d.push({ a: b.a.map((y) => -y), b: -b.b }));
  }
  analysisRequest.variables.forEach((variable, b) => {
    const y = Array(o).fill(0);
    ((y[b] = 1),
      d.push(
        { a: y, b: variable.upper },
        { a: y.map((w) => -w), b: -variable.lower },
      ));
  });
  const u = (g) =>
    g.every(Number.isFinite) &&
    d.every((b) => {
      const y = b.a.reduce((E, A, k) => E + A * g[k], 0),
        w = Math.max(
          1,
          Math.abs(b.b),
          b.a.reduce((E, A, k) => E + Math.abs(A * g[k]), 0),
        );
      return y - b.b <= analysisRequest.tolerance * w;
    });
  let l,
    objective = 0,
    h = false,
    f = false;
  const p = (g) => {
      if (!u(g)) return;
      const b = evaluateLinearExpression(n.objective, g);
      (!l ||
        (analysisRequest.direction === "maximize" && b > objective) ||
        (analysisRequest.direction === "minimize" && b < objective)) &&
        ((l = [...g]), (objective = b));
    },
    m = () =>
      s.iterations >= analysisRequest.maxCandidates
        ? ((f = true), false)
        : (s.iterations++, true);
  if (a)
    if (
      analysisRequest.variables
        .map((variable) =>
          Math.max(
            0,
            Math.floor(variable.upper) - Math.ceil(variable.lower) + 1,
          ),
        )
        .reduce((b, y) => b * y, 1) > analysisRequest.maxCandidates
    )
      f = true;
    else {
      const b = (y) => {
        if (y.length === o) {
          m() && p(y);
          return;
        }
        const w = analysisRequest.variables[y.length];
        for (let E = Math.ceil(w.lower); E <= Math.floor(w.upper); E++)
          b([...y, E]);
      };
      b([]);
    }
  else {
    let g = 1;
    for (let b = 1; b <= o; b++) g = (g * (d.length - o + b)) / b;
    if (g > analysisRequest.maxCandidates) f = true;
    else {
      const b = (y, w) => {
        if (y.length === o) {
          if (!m()) return;
          const E = y.map((k) => d[k]),
            A = solveLinearSystem(
              E.map((k) => k.a),
              E.map((k) => k.b),
            );
          A === "uncertain" ? (h = true) : A !== "singular" && p(A);
          return;
        }
        for (let E = w; E <= d.length - (o - y.length); E++)
          b([...y, E], E + 1);
      };
      b([], 0);
    }
  }
  if (f)
    return {
      ...s,
      status: "work-limit",
      message:
        "The candidate budget is insufficient to certify a result. Reduce the model size or narrow integer bounds. Baseline unchanged.",
    };
  if (h)
    return {
      ...s,
      status: "precision-limit",
      message:
        "Near-dependent constraints prevent a reliable optimum certificate. Rescale or simplify the constraints. Baseline unchanged.",
    };
  if (!l) return s;
  const solution = l.map((g) => (Object.is(g, -0) ? 0 : g));
  return (
    solution.forEach((g, b) => {
      const cell = analysisRequest.variables[b].cell;
      if (
        workbook.sheets
          .find((sheet) => sheet.id === cell.sheetId)
          .validations?.some(
            (validation) =>
              parseRangeBounds(validation.range, cell.address) &&
              !validateDataValidationRule(validation, g),
          )
      )
        throw Error("The solution violates a changing-cell validation rule.");
    }),
    {
      ...s,
      status: "optimal",
      success: true,
      message:
        analysisRequest.direction === "target"
          ? "A feasible solution reaches the target within tolerance."
          : "All bounded candidates were examined; the best feasible objective is reported within tolerance.",
      solution: solution,
      objective: objective,
      rows: [
        ["Changing cell", "Baseline", "Solution"],
        ...analysisRequest.variables.map((variable, b) => [
          `${workbook.sheets.find((sheet) => sheet.id === variable.cell.sheetId).name}!${variable.cell.address}`,
          n.baseline[b],
          solution[b],
        ]),
        [
          "Objective",
          evaluateLinearExpression(n.objective, n.baseline),
          objective,
        ],
        ["Constraint", "Operator", "Residual (left − right)"],
        ...n.constraints.map((constraint, b) => [
          `Constraint ${b + 1}`,
          constraint.operator,
          evaluateLinearExpression(constraint.form, solution),
        ]),
      ],
    }
  );
}
const assertFiniteStatistic = (e) => {
    if (!Number.isFinite(e))
      throw Error(
        "The statistical result exceeds finite numeric precision. Rescale the sample.",
      );
    return e;
  },
  computeSampleVariance = (e) => {
    const r = exactAverageOfNumbers(e);
    return assertFiniteStatistic(
      sumDecimalPartsToNumber(e.map((t) => (t - r) ** 2)) / (e.length - 1),
    );
  },
  distributionPrimitiveHost = new FormulaParser({
    onCell: () => {
      throw Error("Distribution primitives cannot read workbook cells.");
    },
    onRange: () => {
      throw Error("Distribution primitives cannot read workbook ranges.");
    },
  });
function evaluateDistributionPrimitive(e, r) {
  const t = distributionPrimitiveHost.parse(
    `${e}(${r.map((n) => (typeof n == "boolean" ? String(n).toUpperCase() : String(n))).join(",")})`,
    { sheet: "Primitives", row: 1, col: 1 },
  );
  if (typeof t != "number" || !Number.isFinite(t))
    throw Error(`The ${e} distribution primitive failed.`);
  return t;
}
function fitLinearRegression(e, r) {
  const t = r[0].length;
  if (e.length <= t + 1)
    throw Error("Regression needs more observations than fitted coefficients.");
  if (computeSampleVariance(e) === 0)
    throw Error("Regression requires a nonconstant response.");
  const n = evaluateRegressionFunction(
      "LINEST",
      [e.map((l) => [l]), r, true, true],
      {
        retrieve: (l) => l,
        issue: (l) => new Error(l),
        isIssue: (l) => l instanceof Error,
      },
    ),
    o = n[3][1];
  if (o !== e.length - t - 1)
    throw Error(
      "Regression predictors are rank deficient or numerically dependent.",
    );
  const coefficient = [...n[0]].reverse(),
    standardErrors = [...n[1]].reverse(),
    fitted = evaluateRegressionFunction(
      "TREND",
      [e.map((l) => [l]), r, r, true],
      {
        retrieve: (l) => l,
        issue: (l) => new Error(l),
        isIssue: (l) => l instanceof Error,
      },
    ).map((l) => l[0]),
    residuals = e.map((l, c) => assertFiniteStatistic(l - fitted[c]));
  return {
    coefficient: coefficient,
    standardErrors: standardErrors,
    df: o,
    r2: n[2][0],
    standardError: n[2][1],
    f: n[3][0],
    ssRegression: n[4][0],
    ssError: n[4][1],
    fitted: fitted,
    residuals: residuals,
  };
}
function computeDescriptiveStatistics(workbook, analysisRequest, t = {}) {
  const rangeAddress = parseRangeAddress(analysisRequest.range),
    o = rangeAddress.lastCol - rangeAddress.firstCol + 1,
    a = rangeAddress.lastRow - rangeAddress.firstRow;
  if (
    !/^[A-Z]{1,3}[1-9]\d{0,6}:[A-Z]{1,3}[1-9]\d{0,6}$/.test(
      analysisRequest.range,
    ) ||
    a < 1 ||
    a > 500 ||
    o > 6 ||
    (a + 1) * o > 2e3
  )
    throw Error(
      "Select a headed rectangle with 1–500 data rows, 1–6 columns and at most 2,000 cells.",
    );
  (validateAnalysisSource(
    workbook,
    {
      sheetId: analysisRequest.sheetId,
      address: analysisRequest.range.split(":")[0],
    },
    t,
  ),
    assertAutomaticCalculation(workbook));
  const sourceHeaders = readSourceHeaders(
      workbook,
      [{ sheetId: analysisRequest.sheetId, range: analysisRequest.range }],
      t,
    ),
    d = sourceHeaders.snapshots[0];
  for (const dependencyCell of sourceHeaders.dependencyCells)
    validateAnalysisSource(
      workbook,
      { sheetId: dependencyCell.sheetId, address: dependencyCell.address },
      t,
    );
  const exclusions = [],
    l = [];
  for (const record of d.records)
    if (
      record.values.some(
        (value) => typeof value != "number" || !Number.isFinite(value),
      )
    ) {
      if (analysisRequest.missing === "reject")
        throw Error(
          `Row ${record.row} includes blank or nonnumeric values. Correct it or explicitly exclude incomplete rows.`,
        );
      exclusions.push({
        row: record.row,
        locator: record.locator,
        reason:
          "Complete-case exclusion: at least one blank, text or boolean value.",
      });
    } else l.push({ ...record, values: record.values });
  const included = l.length,
    h = l.map((b) => b.values),
    headers = d.headers,
    p = () => {
      if (o !== 1)
        throw Error("This analysis requires exactly one numeric column.");
      return h.map((b) => b[0]);
    };
  if (
    included <
    (analysisRequest.method === "histogram" ||
    analysisRequest.method === "distribution"
      ? 1
      : 2)
  )
    throw Error("Insufficient numeric observations after exclusions.");
  const assumptions = [
    "Headed numeric inputs; no text-to-number or date conversion. Complete-case exclusions apply to every selected column.",
    "Hidden cells require explicit authorization; calculation errors, dynamic/volatile formulas and generated report inputs are refused.",
    "Binary64 arithmetic; local MIT implementation and the existing pinned MIT fast-formula-parser distribution primitives. No proprietary engine or new license dependency.",
  ];
  let algorithm = "",
    rows = [];
  if (analysisRequest.method === "descriptive") {
    ((algorithm = "Two-pass sample moments and sorted inclusive quartiles"),
      assumptions.push(
        "Sample variance and standard deviation use n − 1. Median and quartiles use linear inclusive interpolation.",
        "Independent reference: Python statistics and analytic fixtures in ANALYSIS-VERIFICATION.json.",
      ));
    const b = [
        "Count",
        "Sum",
        "Mean",
        "Standard error",
        "Median",
        "Sample standard deviation",
        "Sample variance",
        "Minimum",
        "Q1 (inclusive)",
        "Q3 (inclusive)",
        "Maximum",
        "Range",
      ],
      y = headers.map((header, E) => {
        const A = h.map((O) => O[E]),
          k = [...A].sort((O, j) => O - j),
          sampleVariance = computeSampleVariance(A),
          I = (O) => {
            const j = (included - 1) * O,
              z = Math.floor(j);
            return k[z] + (k[Math.ceil(j)] - k[z]) * (j - z);
          };
        return [
          included,
          assertFiniteStatistic(sumDecimalPartsToNumber(A)),
          assertFiniteStatistic(exactAverageOfNumbers(A)),
          assertFiniteStatistic(Math.sqrt(sampleVariance / included)),
          I(0.5),
          assertFiniteStatistic(Math.sqrt(sampleVariance)),
          sampleVariance,
          k[0],
          I(0.25),
          I(0.75),
          k[included - 1],
          assertFiniteStatistic(k[included - 1] - k[0]),
        ];
      });
    rows = [
      ["Statistic", ...headers],
      ...b.map((w, E) => [w, ...y.map((A) => A[E])]),
    ];
  } else if (
    analysisRequest.method === "correlation" ||
    analysisRequest.method === "covariance"
  ) {
    if (o < 2) throw Error("Select at least two columns for a matrix.");
    ((algorithm =
      analysisRequest.method === "correlation"
        ? "Pearson product-moment correlation"
        : "Two-pass sample covariance"),
      assumptions.push(
        "Every matrix entry uses the same complete-case observations; covariance uses n − 1.",
        "Independent reference: analytic matrices and Python statistics in ANALYSIS-VERIFICATION.json.",
      ));
    const b = headers.map((header, E) =>
        exactAverageOfNumbers(h.map((A) => A[E])),
      ),
      y = headers.map((header, E) =>
        assertFiniteStatistic(
          sumDecimalPartsToNumber(h.map((A) => (A[E] - b[E]) ** 2)),
        ),
      );
    if (analysisRequest.method === "correlation" && y.some((w) => w === 0))
      throw Error("Correlation is undefined for a constant column.");
    rows = [
      [analysisRequest.method, ...headers],
      ...headers.map((header, E) => [
        header,
        ...headers.map((header, k) => {
          const F = assertFiniteStatistic(
            sumDecimalPartsToNumber(
              h.map((I) => (I[E] - b[E]) * (I[k] - b[k])),
            ),
          );
          return analysisRequest.method === "correlation"
            ? Math.max(
                -1,
                Math.min(
                  1,
                  assertFiniteStatistic(F / Math.sqrt(y[E]) / Math.sqrt(y[k])),
                ),
              )
            : assertFiniteStatistic(F / (included - 1));
        }),
      ]),
    ];
  } else if (analysisRequest.method === "regression") {
    if (o < 2)
      throw Error(
        "Regression requires Y in the first column and one or more X columns.",
      );
    const b = fitLinearRegression(
        h.map((A) => A[0]),
        h.map((A) => A.slice(1)),
      ),
      y = evaluateDistributionPrimitive("T.INV.2T", [0.05, b.df]);
    ((algorithm =
      "Ordinary least squares with intercept; centered, scaled pivoted QR"),
      assumptions.push(
        "First column is Y; remaining columns are X. Independent errors, linear mean and constant error variance; t/F inference additionally assumes normal errors.",
        "No automatic causal interpretation. Rank deficient and insufficient samples are refused. Perfect fits have undefined t/F values shown explicitly.",
        "Independent certified reference: NIST StRD Norris, 36 observations; see ANALYSIS-VERIFICATION.json.",
      ));
    const w = ["Intercept", ...headers.slice(1)].map((A, k) => {
        const F = b.standardErrors[k],
          I = b.coefficient[k],
          O = F === 0 ? null : assertFiniteStatistic(I / F);
        return [
          A,
          I,
          F,
          O,
          O === null
            ? null
            : evaluateDistributionPrimitive("BETA.DIST", [
                b.df / (b.df + O * O),
                b.df / 2,
                0.5,
                true,
              ]),
          I - y * F,
          I + y * F,
        ];
      }),
      E = typeof b.f == "number" ? b.f : null;
    rows = [
      ["Metric", "Value"],
      ["Observations", included],
      ["R squared", b.r2],
      ["Residual standard error", b.standardError],
      ["Residual df", b.df],
      ["Regression SS", b.ssRegression],
      ["Residual SS", b.ssError],
      ["F statistic", E],
      [
        "Significance F",
        E === null
          ? null
          : evaluateDistributionPrimitive("BETA.DIST", [
              b.df / (b.df + (o - 1) * E),
              b.df / 2,
              (o - 1) / 2,
              true,
            ]),
      ],
      [
        "Coefficient",
        "Estimate",
        "Standard error",
        "t",
        "Two-sided p",
        "95% lower",
        "95% upper",
      ],
      ...w,
      ["Source row", "Observed Y", "Fitted Y", "Residual"],
      ...l.map((A, k) => [A.row, A.values[0], b.fitted[k], b.residuals[k]]),
    ];
  } else if (analysisRequest.method === "histogram") {
    const b = p(),
      bins = analysisRequest.bins;
    if (!bins?.length || bins.some((bin, k) => k > 0 && bin <= bins[k - 1]))
      throw Error(
        "Histogram bin upper bounds must be finite and strictly increasing.",
      );
    ((algorithm = "Explicit right-closed histogram bins"),
      assumptions.push(
        "First bin includes all values at or below its upper bound; later bins are (previous, upper]; last bin is overflow.",
        "Independent reference: exact integer counting fixtures in ANALYSIS-VERIFICATION.json.",
      ));
    const w = Array(bins.length + 1).fill(0);
    b.forEach((A) => {
      const k = bins.findIndex((bin) => A <= bin);
      w[k < 0 ? bins.length : k]++;
    });
    let E = 0;
    rows = [
      [
        "Upper bound",
        "Frequency",
        "Relative frequency",
        "Cumulative frequency",
      ],
      ...w.map(
        (A, k) => (
          (E += A),
          [
            k === bins.length ? "Overflow" : bins[k],
            A,
            A / included,
            E / included,
          ]
        ),
      ),
    ];
  } else if (analysisRequest.method === "distribution") {
    const b = p();
    ((algorithm =
      "Pinned analytic probability density/mass and cumulative distribution"),
      assumptions.push(
        "Selected cells are evaluation points, not a fitted sample. Distribution parameters are supplied explicitly.",
        "Independent reference: Python NormalDist, math.comb and factorial probability sums in ANALYSIS-VERIFICATION.json.",
      ));
    const y = analysisRequest.distribution ?? "normal";
    let w, E;
    if (y === "normal") {
      if (
        analysisRequest.mean === undefined ||
        analysisRequest.standardDeviation === undefined
      )
        throw Error(
          "Normal distribution needs a finite mean and positive standard deviation.",
        );
      ((E = "NORM.DIST"),
        (w = (A, k) => [
          A,
          analysisRequest.mean,
          analysisRequest.standardDeviation,
          k,
        ]),
        assumptions.push(
          `Normal mean ${analysisRequest.mean}; standard deviation ${analysisRequest.standardDeviation}.`,
        ));
    } else if (y === "binomial") {
      if (
        analysisRequest.trials === undefined ||
        analysisRequest.probability === undefined ||
        b.some(
          (A) => !Number.isInteger(A) || A < 0 || A > analysisRequest.trials,
        )
      )
        throw Error(
          "Binomial points must be integers from zero through the explicit trial count; provide p in [0,1].",
        );
      ((E = "BINOM.DIST"),
        (w = (A, k) => [
          A,
          analysisRequest.trials,
          analysisRequest.probability,
          k,
        ]),
        assumptions.push(
          `Independent trials ${analysisRequest.trials}; success probability ${analysisRequest.probability}.`,
        ));
    } else {
      if (
        analysisRequest.lambda === undefined ||
        b.some((A) => !Number.isInteger(A) || A < 0 || A > 1e3)
      )
        throw Error(
          "Poisson points must be integers from zero through 1,000; provide a positive rate.",
        );
      ((E = "POISSON.DIST"),
        (w = (A, k) => [A, analysisRequest.lambda, k]),
        assumptions.push(
          `Poisson rate ${analysisRequest.lambda}; independent arrivals and constant rate.`,
        ));
    }
    rows = [
      [
        "Point",
        y === "normal" ? "Density" : "Probability mass",
        "Cumulative probability",
      ],
      ...b.map((A) => [
        A,
        evaluateDistributionPrimitive(E, w(A, false)),
        evaluateDistributionPrimitive(E, w(A, true)),
      ]),
    ];
  } else {
    if (o !== 2)
      throw Error(
        "Forecasting requires two columns: increasing equally spaced numeric period, then value.",
      );
    const b = h.map((j) => j[0]),
      y = h.map((j) => j[1]),
      w = b[1] - b[0];
    if (
      !(w > 0) ||
      b.some(
        (j, z) =>
          z > 0 && Math.abs(j - b[z - 1] - w) > 1e-9 * Math.max(1, Math.abs(w)),
      )
    )
      throw Error(
        "Forecast periods must be strictly increasing and equally spaced without gaps.",
      );
    const E = analysisRequest.horizon ?? 6,
      A = analysisRequest.forecast ?? "linear",
      k = Array.from({ length: E }, (j, z) =>
        assertFiniteStatistic(b[included - 1] + w * (z + 1)),
      ),
      F = [],
      I = [];
    if (A === "linear") {
      const j = fitLinearRegression(
        y,
        b.map((V) => [V]),
      );
      algorithm = "Linear trend forecast using centered QR least squares";
      const z = evaluateRegressionFunction(
        "TREND",
        [y.map((V) => [V]), b.map((V) => [V]), k.map((V) => [V]), true],
        {
          retrieve: (V) => V,
          issue: (V) => new Error(V),
          isIssue: (V) => V instanceof Error,
        },
      );
      (F.push(...z.map((V) => V[0])),
        j.fitted.forEach((V, D) => I.push({ index: D, value: V })));
    } else if (A === "moving-average") {
      const j = analysisRequest.window ?? 3;
      if (j >= included)
        throw Error("Moving average needs more observations than its window.");
      algorithm = `Trailing ${j}-period moving average with recursive forecasts`;
      for (let V = j; V < included; V++)
        I.push({ index: V, value: exactAverageOfNumbers(y.slice(V - j, V)) });
      const z = [...y];
      for (let V = 0; V < E; V++) {
        const D = assertFiniteStatistic(exactAverageOfNumbers(z.slice(-j)));
        (F.push(D), z.push(D));
      }
    } else {
      if (included < 3)
        throw Error("Exponential smoothing needs at least three observations.");
      const j = analysisRequest.alpha ?? 0.3;
      algorithm =
        "Simple exponential smoothing with fixed alpha and first-observation initialization";
      let z = y[0];
      for (let V = 1; V < included; V++)
        (I.push({ index: V, value: z }),
          (z = assertFiniteStatistic(j * y[V] + (1 - j) * z)));
      (F.push(...k.map(() => z)),
        assumptions.push(`Alpha ${j}; no trend or seasonal component.`));
    }
    const O = I.map((j) => assertFiniteStatistic(y[j.index] - j.value));
    (assumptions.push(
      "Numeric equally spaced periods; no seasonality, automatic tuning, ETS intervals or causal claims. Extrapolation can be unreliable.",
      "Error metrics use fitted/in-sample observations, not an independent holdout. No uncertainty interval is claimed.",
      "Independent reference: analytic forecasts and Python calculations in ANALYSIS-VERIFICATION.json.",
    ),
      (rows = [
        ["Metric", "Value"],
        ["Fitted observations", I.length],
        ["In-sample MAE", exactAverageOfNumbers(O.map(Math.abs))],
        [
          "In-sample RMSE",
          Math.sqrt(
            exactAverageOfNumbers(O.map((j) => assertFiniteStatistic(j * j))),
          ),
        ],
        ["Period", "Forecast"],
        ...k.map((j, z) => [j, assertFiniteStatistic(F[z])]),
        ["Source row", "Period", "Observed", "Fitted", "Residual"],
        ...I.map((j) => [
          l[j.index].row,
          b[j.index],
          y[j.index],
          j.value,
          y[j.index] - j.value,
        ]),
      ]));
  }
  for (const row of rows)
    for (const y of row) typeof y == "number" && assertFiniteStatistic(y);
  return {
    status: "complete",
    success: true,
    message:
      "Calculation completed for the declared method and included observations. Review the assumptions and exclusions before saving.",
    algorithm: algorithm,
    assumptions: assumptions,
    exclusions: exclusions,
    included: included,
    inspected: d.records.length,
    iterations: 0,
    rows: rows,
  };
}
const analysisResultCache = new WeakMap();
function runAnalysisRequest(workbook, request, t = {}) {
  const analysisRequest = parseAnalysisRequest(request),
    analysisPermissions = normalizeAnalysisPermissions(workbook, t),
    a =
      analysisRequest.kind === "solver"
        ? solveBoundedLinearModel(workbook, analysisRequest, t)
        : computeDescriptiveStatistics(workbook, analysisRequest, t),
    s = {
      request: analysisRequest,
      result: a,
      workbookHash: fingerprintCanonicalValue(workbook),
      contextHash: fingerprintCanonicalValue(analysisPermissions),
      sourceHash: p_(workbook),
    };
  return (analysisResultCache.set(s, fingerprintCanonicalValue(s)), s);
}
function buildAnalysisReport(e) {
  const result = e.result;
  return [
    [
      e.request.kind === "solver"
        ? "Solver report"
        : "Statistical analysis report",
    ],
    ["Outcome", result.status],
    ["Message", result.message],
    ["Algorithm", result.algorithm],
    ["Engine", analysisFormatVersion],
    [
      "License",
      "MIT; existing MIT distribution dependency; no proprietary engine",
    ],
    ["Inspected", result.inspected],
    ["Included", result.included],
    ["Excluded", result.exclusions.length],
    ["Candidates examined", result.iterations],
    ...result.assumptions.map((assumption, n) => [
      `Assumption ${n + 1}`,
      assumption,
    ]),
    [],
    ...result.rows,
    [],
    ["Excluded source row", "Reason", "Locator"],
    ...result.exclusions.map((exclusion) => [
      exclusion.row,
      exclusion.reason,
      exclusion.locator,
    ]),
  ];
}
function buildAnalysisReportCells(worksheetRecord, outputRange) {
  const rangeAddress = parseRangeAddress(outputRange);
  if (
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    2e4
  )
    throw Error("Analysis output exceeds 20,000 cells.");
  const n = [];
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  )
    for (
      let firstCol = rangeAddress.firstCol;
      firstCol <= rangeAddress.lastCol;
      firstCol++
    ) {
      const cell =
        worksheetRecord.cells[
          `${columnIndexToLetters(firstCol)}${firstRow + 1}`
        ];
      n.push({
        raw: cell?.raw ?? null,
        literal: !!cell?.literal,
      });
    }
  return fingerprintCanonicalValue(n);
}
function commitAnalysisResult(workbook, r, t, n = {}) {
  const analysisPermissions = normalizeAnalysisPermissions(workbook, n);
  if (analysisPermissions.readOnly) throw Error("This workbook is read-only.");
  if (
    (assertWorkbookStructureEditable(workbook),
    !["report", "apply-solution"].includes(t))
  )
    throw Error("Choose Save report or Apply solution and report.");
  if (
    analysisResultCache.get(r) !== fingerprintCanonicalValue(r) ||
    r.workbookHash !== fingerprintCanonicalValue(workbook) ||
    r.contextHash !== fingerprintCanonicalValue(analysisPermissions)
  )
    throw Error(
      "The preview, workbook or host permissions changed. Preview again.",
    );
  if (!r.result.success)
    throw Error("Only a successful reviewed analysis can be saved or applied.");
  if (
    t === "apply-solution" &&
    (r.request.kind !== "solver" || !r.result.solution)
  )
    throw Error("Only Solver can apply changing-cell values.");
  if (
    workbook.sheets.some(
      (sheet) =>
        sheet.name.toLowerCase() === r.request.outputName.toLowerCase(),
    )
  )
    throw Error("The output worksheet name is already in use.");
  inspectOutputNameSafety(workbook, r.request.outputName);
  const workbookInput = normalizeWorkbookInput(workbook);
  t === "apply-solution" &&
    r.request.kind === "solver" &&
    r.request.variables.forEach((variable, f) => {
      const sheet = workbookInput.sheets.find(
        (sheet) => sheet.id === variable.cell.sheetId,
      );
      sheet.cells[variable.cell.address] = {
        ...sheet.cells[variable.cell.address],
        raw: r.result.solution[f],
      };
    });
  const worksheetRecord = createWorksheetRecord(r.request.outputName),
    analysisReport = buildAnalysisReport(r),
    u = Math.max(...analysisReport.map((h) => h.length));
  if (analysisReport.length * u > 2e4)
    throw Error("Analysis report exceeds the output budget.");
  (analysisReport.forEach((h, f) =>
    h.forEach((raw, m) => {
      worksheetRecord.cells[`${columnIndexToLetters(m)}${f + 1}`] = {
        raw: raw,
        ...(typeof raw == "string" ? { literal: true } : {}),
        ...(f === 0
          ? { style: { bold: true, fill: "#362C63", color: "#FFFFFF" } }
          : {}),
      };
    }),
  ),
    (worksheetRecord.columnWidths = {
      0: 210,
      1: 260,
      2: 180,
      3: 160,
      4: 140,
      5: 140,
      6: 140,
    }),
    (worksheetRecord.frozenRows = 1));
  const outputRange = `A1:${columnIndexToLetters(u - 1)}${analysisReport.length}`;
  ((worksheetRecord.analysisReport = {
    engine: analysisFormatVersion,
    request: structuredClone(r.request),
    result: structuredClone(r.result),
    definitionHash: fingerprintCanonicalValue(r.request),
    sourceHash: p_(workbookInput),
    outputHash: buildAnalysisReportCells(worksheetRecord, outputRange),
    outputRange: outputRange,
    createdAt: new Date().toISOString(),
    sourceRevision: workbook.revision,
    includeHidden: !!n.includeHidden,
    applied: t === "apply-solution",
  }),
    workbookInput.sheets.push(worksheetRecord));
  const c = normalizeWorkbookInput(workbookInput);
  return (analysisResultCache.delete(r), c);
}
function readSavedAnalysisReport(workbook, sheetId, t = {}) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    o = sheet?.analysisReport;
  if (!sheet || !o)
    return { status: "unavailable", message: "No saved analysis report." };
  const analysisPermissions = normalizeAnalysisPermissions(workbook, t);
  if (
    !analysisPermissions.allowedSheetIds.includes(sheetId) ||
    (!analysisPermissions.includeHidden && (sheet.hidden || sheet.veryHidden))
  )
    return {
      status: "unavailable",
      message: "The report is outside the authorized context.",
    };
  if (
    (o.includeHidden && !analysisPermissions.includeHidden) ||
    analysisPermissions.allowedSheetIds.length < workbook.sheets.length
  )
    return {
      status: "unavailable",
      message:
        "Reopen the original analysis within the authorized source context before trusting this historical report.",
    };
  if (
    o.definitionHash !== fingerprintCanonicalValue(o.request) ||
    o.sourceHash !== p_(workbook) ||
    o.outputHash !== buildAnalysisReportCells(sheet, o.outputRange)
  )
    return {
      status: "stale",
      message:
        "Source data, calculation settings or report cells changed. Preview the analysis again; saved values are historical.",
    };
  const worksheetRecord = createWorksheetRecord("Expected");
  return (
    buildAnalysisReport(o).forEach((d, u) =>
      d.forEach(
        (raw, c) =>
          (worksheetRecord.cells[`${columnIndexToLetters(c)}${u + 1}`] = {
            raw: raw,
            ...(typeof raw == "string" ? { literal: true } : {}),
          }),
      ),
    ),
    buildAnalysisReportCells(worksheetRecord, o.outputRange) !== o.outputHash
      ? {
          status: "stale",
          message: "The saved result and report cells disagree. Preview again.",
        }
      : {
          status: "current",
          message:
            "Historical reviewed snapshot; source fingerprint is unchanged. Reopen to rerun or refresh.",
        }
  );
}
const parseNumericText = (e) => (e.trim() === "" ? NaN : Number(e));

export {
  readSavedAnalysisReport,
  commitAnalysisResult,
  runAnalysisRequest,
  parseNumericText,
};
