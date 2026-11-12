import {
  dateToSerial,
  evaluateDateFunction,
  daysInExcelMonth,
  financialSolverLimits,
  serialToDateParts,
  dateSystemMaxSerial,
  parseDateText,
} from "./01-core-schemas-and-hashing.js";

function solveRateCandidates(e, r, t = financialSolverLimits.minRate) {
  const n = Math.log1p(t),
    o = Math.log1p(financialSolverLimits.maxRate),
    a = Math.log1p(r);
  let s = 0;
  const d = [],
    u = (g) => {
      if (
        g < n ||
        g > o ||
        !Number.isFinite(g) ||
        s >= financialSolverLimits.evaluations
      )
        return;
      s++;
      const b = e(Math.expm1(g));
      if (!(
        !Number.isFinite(b.value) ||
        !Number.isFinite(b.scale) ||
        b.scale <= 0
      ))
        return { ...b, normalized: b.value / b.scale };
    },
    l = (g, b) => {
      if (!b || Math.abs(b.normalized) > financialSolverLimits.relativeResidual)
        return false;
      const y = Math.expm1(g);
      return (
        d.some(
          (w) => Math.abs(w.rate - y) <= 1e-7 * Math.max(1, Math.abs(y)),
        ) ||
          d.push({
            rate: y,
            residual: b.value,
            relativeResidual: Math.abs(b.normalized),
          }),
        true
      );
    };
  if ([0, Math.log1p(0.1), Math.log1p(0.5)].map(u).every((g) => g?.value === 0))
    return {
      status: "indeterminate",
      candidates: d,
      evaluations: s,
      exhaustive: false,
    };
  let h = a;
  for (let g = 0; g < 60; g++) {
    const b = u(h);
    if (!b || l(h, b)) break;
    const y = 1e-5,
      w = u(Math.max(n, h - y)),
      E = u(Math.min(o, h + y));
    if (!w || !E) break;
    const A =
      (E.normalized - w.normalized) / (Math.min(o, h + y) - Math.max(n, h - y));
    if (!Number.isFinite(A) || Math.abs(A) < 1e-15) break;
    const k = h - Math.max(-1, Math.min(1, b.normalized / A));
    if (k < n || k > o || k === h) break;
    h = k;
  }
  const f = [
    ...new Set([
      ...Array.from({ length: 193 }, (g, b) => n + ((o - n) * b) / 192),
      ...[
        -0.9,
        -0.5,
        -0.1,
        0,
        0.01,
        0.025,
        0.05,
        0.075,
        0.1,
        0.15,
        0.2,
        0.3,
        0.5,
        1,
        2,
        10,
        r,
      ]
        .filter((g) => g >= t)
        .map(Math.log1p),
    ]),
  ].sort((g, b) => g - b);
  let p;
  for (const g of f) {
    const b = u(g);
    if (!b) {
      p = undefined;
      continue;
    }
    if ((l(g, b), p && p.y.normalized * b.normalized < 0)) {
      let y = p.x,
        w = g,
        E = p.y.normalized;
      for (let A = 0; A < 100; A++) {
        const k = (y + w) / 2,
          F = u(k);
        if (!F || l(k, F) || k === y || k === w) break;
        E * F.normalized < 0 ? (w = k) : ((y = k), (E = F.normalized));
      }
    }
    p = { x: g, y: b };
  }
  d.sort((g, b) => g.rate - b.rate);
  const m = [...d].sort(
      (g, b) =>
        Math.abs(Math.log1p(g.rate) - a) - Math.abs(Math.log1p(b.rate) - a),
    )[0],
    v = s >= financialSolverLimits.evaluations;
  return {
    ...(m && !v ? { selected: m } : {}),
    candidates: d,
    evaluations: s,
    status: v ? "limit" : m ? "converged" : "no-convergence",
    exhaustive: false,
  };
}
const supportedFinancialFunctionNames = [
    "PV",
    "FV",
    "PMT",
    "IPMT",
    "PPMT",
    "RATE",
    "NPER",
    "NPV",
    "XNPV",
    "IRR",
    "XIRR",
    "MIRR",
    "SLN",
    "DB",
    "DDB",
    "SYD",
    "PRICE",
    "YIELD",
    "ACCRINT",
    "ACCRINTM",
  ],
  notFinancialSentinel = Symbol("not-financial"),
  financialArgumentRanges = {
    PV: [3, 5],
    FV: [3, 5],
    PMT: [3, 5],
    IPMT: [4, 6],
    PPMT: [4, 6],
    RATE: [3, 6],
    NPER: [3, 5],
    NPV: [2, 255],
    XNPV: [3, 3],
    IRR: [1, 2],
    XIRR: [2, 3],
    MIRR: [3, 3],
    SLN: [3, 3],
    DB: [4, 5],
    DDB: [4, 5],
    SYD: [4, 4],
    PRICE: [6, 7],
    YIELD: [6, 7],
    ACCRINT: [6, 8],
    ACCRINTM: [4, 5],
  },
  flattenNestedValues = (e) =>
    Array.isArray(e) ? e.flatMap(flattenNestedValues) : [e],
  compensatedSum = (e) => {
    let r = 0,
      t = 0;
    for (const n of e) {
      const o = n - t,
        a = r + o;
      ((t = a - r - o), (r = a));
    }
    return r;
  },
  countSignChanges = (e) => {
    const r = e.filter((t) => t !== 0).map(Math.sign);
    return r.slice(1).filter((t, n) => t !== r[n]).length;
  };
function evaluateFinancialFunction(e, r, t) {
  if (!supportedFinancialFunctionNames.includes(e)) return notFinancialSentinel;
  let n =
      "Cash paid out is negative; cash received is positive. Rates and periods must use the same interval. No automatic cent rounding.",
    o = false;
  const a = t.clock.dateSystem,
    s = (k, F, I) => dateToSerial(k, F, I, a),
    d = (k) => serialToDateParts(k, a),
    u = (k) => {
      ((o = true), t.report?.({ functionName: e, convention: n, ...k }));
    },
    l = (k, F) => {
      throw (u({ status: "error", error: k, message: F }), t.issue(k));
    },
    c = (k) =>
      Number.isFinite(k)
        ? k
        : l(
            "#NUM!",
            "The financial equation overflowed or has no finite result. Review rates, signs and periods.",
          ),
    h = (k) => {
      if (t.isIssue(k)) throw k;
      return Array.isArray(k)
        ? k.length === 1
          ? h(k[0])
          : l("#VALUE!", "A scalar argument must contain exactly one value.")
        : k;
    },
    f = (k, F) => (
      (k = h(k)),
      k == null
        ? (F ?? 0)
        : typeof k == "boolean"
          ? Number(k)
          : typeof k != "number" &&
              (typeof k != "string" ||
                !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(k))
            ? l(
                "#VALUE!",
                "Expected a finite number; nonnumeric text cannot be used as a financial input.",
              )
            : c(Number(k))
    ),
    p = (k) => {
      if (((k = h(k)), k == null || typeof k == "boolean"))
        return l(
          "#VALUE!",
          "A cash-flow or security date is missing or is not a date.",
        );
      const F = typeof k == "string" ? parseDateText(k, t.clock) : f(k);
      return F === undefined ||
        F < (a === "1904" ? 0 : 1) ||
        F >= dateSystemMaxSerial(a) + 1
        ? l(
            "#VALUE!",
            "Invalid cash-flow or security date. Use a valid Gregorian/workbook-system serial or explicit full-year date text.",
          )
        : Math.floor(F);
    },
    m = (k) => (
      k <= 0 && l("#NUM!", "Periods must be greater than zero."),
      k > financialSolverLimits.periods &&
        l("#LIMIT!", "Financial periods exceed the 10000-period budget."),
      k
    ),
    v = (k) =>
      k > -1 ? k : l("#NUM!", "The rate must be greater than -100%."),
    g = (k) =>
      [0, 1].includes(k)
        ? k
        : l("#NUM!", "Payment timing must be 0 (end) or 1 (beginning)."),
    b = (k, F) => Math.exp(F * Math.log1p(k)),
    y = (k, F) => (k === 0 ? F : Math.expm1(F * Math.log1p(k)) / k),
    w = (k, F, I, O, j) => -(I * b(k, F) + O) / ((1 + k * j) * y(k, F)),
    E = (k, F = false) => {
      const I = flattenNestedValues(k);
      I.length > financialSolverLimits.cashFlows &&
        l("#LIMIT!", "Cash-flow input exceeds 10000 positions.");
      const O = [];
      for (const j of I) {
        if (t.isIssue(j)) throw j;
        typeof j == "number"
          ? O.push(c(j))
          : F &&
            l(
              "#VALUE!",
              "Dated cash flows require one numeric amount for every date; blanks, booleans and text cannot be silently removed.",
            );
      }
      return (
        O.length || l("#NUM!", "No numeric cash flows were supplied."),
        O
      );
    },
    A = (k, F, I, O) => {
      (F < (O ?? financialSolverLimits.minRate) ||
        F > financialSolverLimits.maxRate) &&
        l("#NUM!", "The rate guess is outside the supported search interval.");
      const j = solveRateCandidates(k, F, O),
        z = j.status === "limit" ? "#LIMIT!" : "#NUM!";
      if (
        (u({
          status: j.selected ? "converged" : "error",
          ...(j.selected ? {} : { error: z }),
          guess: F,
          signChanges: I,
          solver: j,
          message: j.selected
            ? `Residual-checked rate. ${j.candidates.length} candidate root(s) found; selected nearest the guess in log-rate space. ${I !== undefined && I > 1 ? "Cash flows change sign more than once; multiple returns may exist. " : ""}The bounded search does not prove uniqueness.`
            : j.status === "indeterminate"
              ? "The equation does not identify a unique rate; no rate has been invented."
              : "No residual-checked rate converged within the supported search interval and evaluation budget. Review signs, timing and guess.",
        }),
        !j.selected)
      )
        throw t.issue(z);
      return j.selected.rate;
    };
  try {
    const [k, F] = financialArgumentRanges[e];
    (r.length < k || r.length > F) &&
      l("#VALUE!", `${e} requires ${k}–${F} arguments.`);
    const I = r.map(t.retrieve);
    let O;
    if (["PV", "FV", "PMT", "IPMT", "PPMT", "RATE", "NPER"].includes(e))
      if (e === "RATE") {
        const j = m(f(I[0])),
          z = f(I[1]),
          V = f(I[2]),
          D = f(I[3]),
          B = g(f(I[4]));
        ([V, z, D].every((_) => _ === 0) &&
          l("#NUM!", "All amounts are zero; no rate can be determined."),
          (O = A(
            (_) => {
              const C = [V * b(_, j), z * (1 + _ * B) * y(_, j), D];
              return {
                value: compensatedSum(C),
                scale: compensatedSum(C.map(Math.abs)),
              };
            },
            f(I[5], 0.1),
          )));
      } else if (e === "NPER") {
        const j = v(f(I[0])),
          z = f(I[1]),
          V = f(I[2]),
          D = f(I[3]),
          B = g(f(I[4]));
        if (j === 0)
          (z === 0 &&
            l(
              "#DIV/0!",
              "Zero rate and zero payment cannot determine a period count.",
            ),
            (O = -(V + D) / z));
        else {
          const _ = z * (1 + j * B),
            C = (_ - D * j) / (V * j + _);
          (C > 0 ||
            l(
              "#NUM!",
              "Cash-flow signs or payment size do not produce a finite real period count.",
            ),
            (O = Math.log(C) / Math.log1p(j)));
        }
        m(O);
      } else if (e === "IPMT" || e === "PPMT") {
        const j = v(f(I[0])),
          z = f(I[1]),
          V = m(f(I[2])),
          D = f(I[3]),
          B = f(I[4]),
          _ = g(f(I[5]));
        (z < 1 || z > V || !Number.isInteger(z)) &&
          l(
            "#NUM!",
            "Interest/principal period must be an integer from 1 through nper.",
          );
        const C = w(j, V, D, B, _),
          T =
            _ && z === 1
              ? 0
              : (-(D * b(j, z - 1) + C * (1 + j * _) * y(j, z - 1)) * j) /
                (1 + j * _);
        O = e === "IPMT" ? T : C - T;
      } else {
        const j = v(f(I[0])),
          z = m(f(I[1])),
          V = f(I[2]),
          D = f(I[3]),
          B = g(f(I[4]));
        O =
          e === "PMT"
            ? w(j, z, V, D, B)
            : e === "FV"
              ? -(D * b(j, z) + V * (1 + j * B) * y(j, z))
              : -(D + V * (1 + j * B) * y(j, z)) / b(j, z);
      }
    else if (["NPV", "XNPV", "IRR", "XIRR", "MIRR"].includes(e)) {
      const j = e.startsWith("X"),
        z = e === "IRR" || e === "XIRR",
        V = E(e === "NPV" ? I.slice(1) : I[z || e === "MIRR" ? 0 : 1], j);
      let D = V.map((_, C) => C + (e === "NPV" ? 1 : 0));
      if (j) {
        const _ = flattenNestedValues(I[z ? 1 : 2]);
        _.length !== V.length &&
          l("#NUM!", "Every cash-flow amount must have exactly one date.");
        const C = _.map(p);
        (C.some((T) => T < C[0]) &&
          l("#NUM!", "A cash-flow date precedes the first cash-flow date."),
          (D = C.map((T) => (T - C[0]) / 365)),
          (n =
            "Dated cash flows use truncated Gregorian/workbook-system dates and actual elapsed days / 365, including leap years. The first date is time zero; later dates may be unsorted or repeated. One numeric amount per date."));
      } else
        n =
          "Periodic cash-flow positions are equally spaced. NPV starts at period 1; IRR/MIRR start at period 0. Numeric zeros retain a period; blanks, text and booleans in cash-flow lists are ignored. No automatic cent rounding.";
      const B = (_) => {
        const C = V.map((T, N) => T / Math.exp(D[N] * Math.log1p(_)));
        return {
          value: compensatedSum(C),
          scale: compensatedSum(C.map(Math.abs)),
        };
      };
      if (z || e === "MIRR")
        if (
          ((!V.some((_) => _ < 0) || !V.some((_) => _ > 0)) &&
            l(
              "#NUM!",
              "At least one positive and one negative cash flow are required.",
            ),
          e === "MIRR")
        ) {
          const _ = v(f(I[1])),
            C = v(f(I[2])),
            T = V.length - 1,
            N = compensatedSum(
              V.map((L, $) => (L < 0 ? L / Math.exp($ * Math.log1p(_)) : 0)),
            ),
            R = compensatedSum(
              V.map((L, $) =>
                L > 0 ? L * Math.exp((T - $) * Math.log1p(C)) : 0,
              ),
            );
          O = Math.expm1(Math.log(-R / N) / T);
        } else O = A(B, f(I[j ? 2 : 1], 0.1), countSignChanges(V));
      else O = B(v(f(I[0]))).value;
    } else if (["SLN", "SYD", "DB", "DDB"].includes(e)) {
      const j = f(I[0]),
        z = f(I[1]),
        V = m(f(I[2]));
      if (
        ((n =
          "Depreciation uses matching life/period units and no automatic cent rounding. DB rounds its fixed rate to three decimals and prorates first/last years; DDB does not switch to straight line."),
        (j < 0 || z < 0) && l("#NUM!", "Cost and salvage must be nonnegative."),
        e === "SLN")
      )
        O = (j - z) / V;
      else {
        const D = f(I[3]);
        if (
          ((D < 1 || D > V + (e === "DB" ? 1 : 0) || !Number.isInteger(D)) &&
            l(
              "#NUM!",
              "Depreciation period must be a supported whole period within asset life (plus DB partial final year).",
            ),
          e === "SYD")
        )
          O = ((j - z) * (V - D + 1) * 2) / (V * (V + 1));
        else if (e === "DB") {
          const B = f(I[4], 12);
          if (
            ((!Number.isInteger(V) ||
              B < 1 ||
              B > 12 ||
              !Number.isInteger(B) ||
              D > V + (B < 12 ? 1 : 0)) &&
              l(
                "#NUM!",
                "DB requires whole life/period and 1–12 first-year months; a partial first year permits one partial final year.",
              ),
            j === 0 || z >= j)
          )
            O = 0;
          else {
            const _ = Math.round((1 - (z / j) ** (1 / V)) * 1e3) / 1e3;
            let C = j;
            O = 0;
            for (let T = 1; T <= D; T++)
              ((O = C * _ * (T === 1 ? B / 12 : T > V ? (12 - B) / 12 : 1)),
                (C -= O));
          }
        } else {
          const B = f(I[4], 2);
          B <= 0 && l("#NUM!", "DDB factor must be greater than zero.");
          let _ = j;
          O = 0;
          for (let C = 1; C <= D; C++)
            ((O = Math.max(0, Math.min((_ * B) / V, _ - z))), (_ -= O));
        }
      }
    } else {
      const j = e === "ACCRINT" ? 6 : e === "ACCRINTM" ? 4 : 6,
        z = Math.trunc(f(I[j]));
      ([0, 1, 2, 3, 4].includes(z) ||
        l(
          "#NUM!",
          "Day-count basis must be 0 (US 30/360), 1 (actual/actual), 2 (actual/360), 3 (actual/365), or 4 (European 30/360).",
        ),
        (n = `Security basis ${z}; annual coupon/yield rates, regular maturity-anchored coupons at frequency 1/2/4, end-of-month anchoring. PRICE/YIELD are per 100 face value. Gregorian/workbook-system dates; no business-day adjustment or automatic rounding.`));
      const V = (C, T) =>
          z === 0 || z === 4
            ? evaluateDateFunction("DAYS360", [C, T, z === 4], {
                ...t,
                retrieve: (N) => N,
              })
            : T - C,
        D = (C, T, N) => (z === 1 ? T - C : (z === 3 ? 365 : 360) / N),
        B = (C) => {
          const T = Math.trunc(f(C));
          return [1, 2, 4].includes(T)
            ? T
            : l("#NUM!", "Coupon frequency must be 1, 2 or 4.");
        },
        _ = (C, T) => {
          const N = d(C),
            R = N.year * 12 + N.month - 1 + T,
            L = Math.floor(R / 12),
            $ = (((R % 12) + 12) % 12) + 1;
          return s(
            L,
            $,
            N.day === daysInExcelMonth(N.year, N.month)
              ? daysInExcelMonth(L, $)
              : Math.min(N.day, daysInExcelMonth(L, $)),
          );
        };
      if (e === "PRICE" || e === "YIELD") {
        const C = p(I[0]),
          T = p(I[1]),
          N = f(I[2]),
          R = f(I[3]),
          L = f(I[4]),
          $ = B(I[5]);
        (C >= T || N < 0 || L <= 0 || R < 0 || (e === "YIELD" && R === 0)) &&
          l(
            "#NUM!",
            "Require settlement before maturity, nonnegative coupon/yield, and positive redemption/price.",
          );
        let U = 1,
          G = _(T, -12 / $);
        for (; G > C;)
          (U >= financialSolverLimits.couponPeriods &&
            l("#LIMIT!", "Security exceeds 1000 coupon periods."),
            U++,
            (G = _(T, (-U * 12) / $)));
        const H = _(T, (-(U - 1) * 12) / $),
          Q = D(G, H, $),
          X = V(G, C),
          J = z === 0 || z === 4 ? Q - X : V(C, H),
          q = J / Q,
          W = (100 * N) / $,
          K = (te) =>
            U === 1
              ? (L + W) / (1 + (te / $) * q) - (W * X) / Q
              : compensatedSum(
                  Array.from(
                    { length: U },
                    (ae, oe) => W / (1 + te / $) ** (oe + q),
                  ),
                ) +
                L / (1 + te / $) ** (U - 1 + q) -
                (W * X) / Q;
        O =
          e === "PRICE"
            ? K(R)
            : A(
                (te) => {
                  const ae = K(te);
                  return { value: ae - R, scale: Math.abs(ae) + R };
                },
                0.1,
                undefined,
                0,
              );
      } else if (e === "ACCRINTM") {
        const C = p(I[0]),
          T = p(I[1]),
          N = f(I[2]),
          R = f(I[3], 1e3);
        (C >= T || N <= 0 || R <= 0) &&
          l(
            "#NUM!",
            "Require issue before settlement, positive coupon rate and positive par.",
          );
        const L = evaluateDateFunction("YEARFRAC", [C, T, z], {
          ...t,
          retrieve: ($) => $,
        });
        ((O = R * N * L),
          (n = `Accrued interest at maturity: par × annual rate × YEARFRAC basis ${z}, using the published local day-count convention. Gregorian/workbook-system; no automatic rounding.`));
      } else {
        const C = p(I[0]),
          T = p(I[1]),
          N = p(I[2]),
          R = f(I[3]),
          L = f(I[4], 1e3),
          $ = B(I[5]),
          U = g(f(I[7], 1));
        (C >= N || T <= C || R <= 0 || L <= 0) &&
          l(
            "#NUM!",
            "Require issue before first interest and settlement, positive rate and positive par.",
          );
        const G = !U && N > T ? T : C;
        let H = Math.floor(
          (d(G).year * 12 + d(G).month - d(T).year * 12 - d(T).month) /
            (12 / $),
        );
        for (; _(T, (H * 12) / $) > G;) H--;
        let Q = _(T, (H * 12) / $),
          X = _(T, ((H + 1) * 12) / $),
          J = 0,
          q = 0;
        for (; Q < N;) {
          ++q > financialSolverLimits.couponPeriods &&
            l("#LIMIT!", "Accrual exceeds 1000 quasi-coupon periods.");
          const W = Math.max(G, Q),
            K = Math.min(N, X);
          (K > W && (J += V(W, K) / D(Q, X, $)),
            (Q = X),
            (X = _(T, ((++H + 1) * 12) / $)));
        }
        ((O = ((L * R) / $) * J),
          (n = `Accrued coupon interest: quasi-coupon periods anchored to first_interest, basis ${z}, frequency ${$}; calc_method ${U} starts at ${!U && N > T ? "first_interest" : "issue"}. No business-day adjustment or automatic rounding.`));
      }
    }
    return (
      (O = c(O)),
      o ||
        u({
          status: "calculated",
          message:
            "Calculated under the displayed sign, timing and day-count convention.",
        }),
      O
    );
  } catch (k) {
    throw (
      o ||
        u({
          status: "error",
          error: t.isIssue(k) ? (k.error ?? "#VALUE!") : "#VALUE!",
          message:
            "An input or dependent calculation failed; inspect the source values and dates.",
        }),
      k
    );
  }
}
const defaultColumnWidth = 84,
  minColumnWidth = 40,
  maxColumnWidth = 600,
  defaultRowHeight = 21,
  minRowHeight = 21,
  maxRowHeight = 546,
  supportedFontFamilies = [
    "Calibri",
    "Aptos",
    "Arial",
    "Times New Roman",
    "Courier New",
  ],
  workbookLimits = {
    maxRows: 1048576,
    maxCols: 16384,
    maxCells: 1e5,
    maxFileBytes: 10 * 1024 * 1024,
    maxNativeBytes: 50 * 1024 * 1024,
    maxLLMCells: 500,
    maxPatches: 500,
  },
  fileImportLimits = {
    ...workbookLimits,
    maxCellTextLength: 32767,
    maxFormulaLength: 4096,
    maxFormulaRangeCells: 2e4,
    maxFormulaDepth: 64,
    maxSpillCells: 1e4,
    maxXlsxExpandedBytes: 50 * 1024 * 1024,
    maxXlsxEntries: 2e3,
    maxXlsxExpansionRatio: 200,
    xlsxRatioThresholdBytes: 1024 * 1024,
    maxXlsxCellRecords: workbookLimits.maxCells,
  },
  workbookLimitTable = [
    ["Rows per worksheet", "1,048,576"],
    ["Columns per worksheet", "16,384 (A–XFD)"],
    ["Worksheets", "No fixed count; available memory and file budgets apply"],
    [
      "Stored cells per workbook",
      "100,000, including cells with only formatting or notes",
    ],
    [
      "Excluded population size",
      "100,000 rows × 20 columns (2,000,000 cells) exceeds this release; split the source before importing",
    ],
    [
      "Interactive performance",
      "100 ms edit / 50 ms scroll targets; local benchmark available, reference browser and host validation required",
    ],
    ["Cell text", "32,767 text units"],
    ["Formula text", "4,096 text units, including the leading ="],
    ["Formula range", "20,000 cells per referenced rectangle"],
    ["Calculation work", "100,000 cell reads; 64 dependent formula levels"],
    ["Dynamic array output", "10,000 cells per array"],
    [
      "Cell changes / assistant selection",
      "500 per patch batch / 500 selected cells",
    ],
    ["XLSX, CSV and TSV imports", "10 MiB (10,485,760 bytes)"],
    [
      "XLSX exports",
      "10 MiB; generated files pass the same archive and expansion limits as import",
    ],
    ["CSV exports", "10 MiB; 100,000 cells including blank gaps"],
    ["Native JSON save and reopen", "50 MiB (52,428,800 UTF-8 bytes)"],
    [
      "XLSX processing",
      "50 MiB expanded; 2,000 archive entries; 100,000 cell records",
    ],
    ["XLSX compression", "200:1 maximum for any entry larger than 1 MiB"],
  ];
function broadcastOverArrays(e, r, t) {
  if (!e.some(Array.isArray)) return r(e);
  const n = e.map((s) =>
      Array.isArray(s) ? (s.every(Array.isArray) ? s : [s]) : [[s]],
    ),
    o = Math.max(...n.map((s) => s.length)),
    a = Math.max(...n.map((s) => s[0]?.length ?? 0));
  return !o ||
    !a ||
    n.some(
      (s) =>
        !s.length ||
        !s[0].length ||
        s.some((d) => d.length !== s[0].length) ||
        (s.length !== o && s.length !== 1) ||
        (s[0].length !== a && s[0].length !== 1),
    )
    ? t("#VALUE!")
    : o * a > fileImportLimits.maxSpillCells
      ? t("#LIMIT!")
      : Array.from({ length: o }, (s, d) =>
          Array.from({ length: a }, (u, l) =>
            r(
              n.map(
                (c) => c[c.length === 1 ? 0 : d][c[0].length === 1 ? 0 : l],
              ),
            ),
          ),
        );
}
const supportedTextFunctionNames = [
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "TRIM",
    "CLEAN",
    "UPPER",
    "LOWER",
    "PROPER",
    "FIND",
    "SEARCH",
    "SUBSTITUTE",
    "REPLACE",
    "VALUE",
    "NUMBERVALUE",
    "TEXT",
    "CONCAT",
    "CONCATENATE",
    "TEXTJOIN",
    "EXACT",
    "TEXTBEFORE",
    "TEXTAFTER",
  ],
  notTextFunctionSentinel = Symbol("not-text-function"),
  coerceValueToText = (e) =>
    e == null ? "" : typeof e == "boolean" ? (e ? "TRUE" : "FALSE") : String(e);
function parseDateTimeText(e) {
  const r =
    /^(?:(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})\/(\d{1,2})\/(\d{4}))(?:[ T](.+))?$/.exec(
      e,
    );
  let t = 0,
    n = e;
  if (r) {
    const c = Number(r[1] ?? r[6]),
      h = Number(r[2] ?? r[4]),
      f = Number(r[3] ?? r[5]);
    if (c < 1900 || c > 9999 || h < 1 || h > 12 || f < 1) return;
    const p = c === 1900 && h === 2 && f === 29,
      m = Date.UTC(c, h - 1, f),
      v = new Date(m);
    if (!p && (v.getUTCMonth() !== h - 1 || v.getUTCDate() !== f)) return;
    if (
      ((t = p
        ? 60
        : (m - Date.UTC(1899, 11, 31)) / 864e5 +
          (m >= Date.UTC(1900, 2, 1) ? 1 : 0)),
      !r[7])
    )
      return t;
    n = r[7];
  }
  const o = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*(AM|PM)?$/i.exec(n);
  if (!o) return;
  let a = Number(o[1]);
  const s = Number(o[2]),
    d = Number(o[3] ?? 0),
    u = o[4]?.toUpperCase();
  if (!(s > 59 || d >= 60 || (u ? a < 1 || a > 12 : a > 23)))
    return (
      u && (a = (a % 12) + (u === "PM" ? 12 : 0)),
      t + (a * 3600 + s * 60 + d) / 86400
    );
}
function evaluateTextFunction(e, r, t) {
  if (!supportedTextFunctionNames.includes(e)) return notTextFunctionSentinel;
  const n = (b = "#VALUE!") => {
      throw t.issue(b);
    },
    o = {
      LEFT: [1, 2],
      RIGHT: [1, 2],
      MID: [3, 3],
      LEN: [1, 1],
      TRIM: [1, 1],
      CLEAN: [1, 1],
      UPPER: [1, 1],
      LOWER: [1, 1],
      PROPER: [1, 1],
      FIND: [2, 3],
      SEARCH: [2, 3],
      SUBSTITUTE: [3, 4],
      REPLACE: [4, 4],
      VALUE: [1, 1],
      NUMBERVALUE: [1, 3],
      TEXT: [2, 2],
      CONCAT: [1, 253],
      CONCATENATE: [1, 255],
      TEXTJOIN: [3, 254],
      EXACT: [2, 2],
      TEXTBEFORE: [2, 6],
      TEXTAFTER: [2, 6],
    },
    [a, s] = o[e];
  (r.length < a || r.length > s) && n();
  const d = r.map(t.retrieve),
    u = d.flatMap(flattenNestedValues).find(t.isIssue);
  if (u) return u;
  ((!(e === "CONCAT" || e === "TEXTJOIN") && d.some(Array.isArray)) ||
    (e === "TEXTJOIN" && d.slice(0, 2).some(Array.isArray))) &&
    n();
  const c = (b) => (b.length <= 32767 ? b : n("#LIMIT!")),
    h = (b, y) => {
      if (b === undefined && y !== undefined) return y;
      if (Array.isArray(b) || (typeof b == "string" && !b.trim())) return n();
      const w = Number(b ?? 0);
      return Number.isFinite(w) ? w : n();
    },
    f = (b, y) => Math.trunc(h(b, y)),
    p = (b) =>
      typeof b == "string"
        ? (/^(true|false)$/i.test(b) || n(), b.toLowerCase() === "true")
        : !!b,
    m = coerceValueToText(d[0]);
  if (e === "LEN") return [...m].length;
  if (e === "TRIM") return c(m.replace(/^ +| +$/g, "").replace(/ +/g, " "));
  if (e === "CLEAN") return c(m.replace(/[\u0000-\u001f]/g, ""));
  if (e === "UPPER") return c(m.toUpperCase());
  if (e === "LOWER") return c(m.toLowerCase());
  if (e === "PROPER")
    return c(
      m
        .toLowerCase()
        .replace(
          new RegExp("(^|[^\\p{L}\\p{M}])(\\p{L})", "gu"),
          (b, y, w) => y + w.toUpperCase(),
        ),
    );
  if (e === "EXACT") return m === coerceValueToText(d[1]);
  if (e === "LEFT" || e === "RIGHT") {
    const b = f(d[1], 1),
      y = [...m];
    return (
      b < 0 && n(),
      (e === "LEFT" ? y.slice(0, b) : b ? y.slice(-b) : []).join("")
    );
  }
  if (e === "MID" || e === "REPLACE") {
    const b = f(d[1]),
      y = f(d[2]),
      w = [...m];
    return (
      (b < 1 || y < 0) && n(),
      e === "MID"
        ? w.slice(b - 1, b - 1 + y).join("")
        : c(
            w.slice(0, b - 1).join("") +
              coerceValueToText(d[3]) +
              w.slice(b - 1 + y).join(""),
          )
    );
  }
  if (e === "FIND" || e === "SEARCH") {
    const b = [...coerceValueToText(d[1])],
      y = f(d[2], 1);
    (y < 1 || y > Math.max(1, b.length)) && n();
    const w = [...m],
      E = [];
    for (let F = 0; F < w.length; F++) {
      const I = w[F];
      e === "SEARCH" && I === "~" && ["*", "?", "~"].includes(w[F + 1])
        ? E.push({ char: w[++F], kind: "literal" })
        : E.push({
            char: I,
            kind:
              e === "SEARCH" && I === "*"
                ? "many"
                : e === "SEARCH" && I === "?"
                  ? "one"
                  : "literal",
          });
    }
    E.length * b.length > 2e6 && n("#LIMIT!");
    let A = Array(b.length + 1).fill(true);
    for (let F = E.length - 1; F >= 0; F--) {
      const I = E[F],
        O = Array(b.length + 1).fill(false);
      O[b.length] = I.kind === "many" && A[b.length];
      for (let j = b.length - 1; j >= y - 1; j--)
        O[j] =
          I.kind === "many"
            ? A[j] || O[j + 1]
            : A[j + 1] &&
              (I.kind === "one" ||
                (e === "SEARCH"
                  ? b[j].toLowerCase() === I.char.toLowerCase()
                  : b[j] === I.char));
      A = O;
    }
    const k = A.findIndex((F, I) => I >= y - 1 && F);
    return k < 0 ? n() : k + 1;
  }
  if (e === "SUBSTITUTE") {
    const b = coerceValueToText(d[1]),
      y = coerceValueToText(d[2]),
      w = d[3] === undefined ? undefined : f(d[3]);
    if ((w !== undefined && w < 1 && n(), !b)) return m;
    const E = m.split(b),
      A = E.length - 1;
    return (
      m.length +
        (w === undefined ? A : w <= A ? 1 : 0) * (y.length - b.length) >
        32767 && n("#LIMIT!"),
      E.reduce(
        (k, F, I) =>
          I === 0 ? F : k + (w === undefined || w === I ? y : b) + F,
        "",
      )
    );
  }
  if (e === "CONCAT" || e === "CONCATENATE" || e === "TEXTJOIN") {
    const b = e === "TEXTJOIN" ? coerceValueToText(d[0]) : "",
      y = e === "TEXTJOIN" && p(d[1]),
      w = d
        .slice(e === "TEXTJOIN" ? 2 : 0)
        .flatMap(flattenNestedValues)
        .filter((E) => !y || (E != null && E !== ""))
        .map(coerceValueToText);
    return (
      w.reduce((E, A) => E + A.length, 0) +
        Math.max(0, w.length - 1) * b.length >
        32767 && n("#LIMIT!"),
      w.join(b)
    );
  }
  if (e === "NUMBERVALUE") {
    const b = [...coerceValueToText(d[1] === undefined ? "." : d[1])][0],
      y = [...coerceValueToText(d[2] === undefined ? "," : d[2])][0];
    (!b || !y || b === y || /[0-9+\-eE%]/.test(b) || /[0-9+\-eE%]/.test(y)) &&
      n();
    let w = m.replace(/ /g, "");
    if (!w) return 0;
    const E = /%+$/.exec(w)?.[0].length ?? 0;
    w = w.slice(0, w.length - E);
    const A = w.split(b);
    ((A.length > 2 || A[1]?.includes(y)) && n(),
      (w = A[0].split(y).join("") + (A.length === 2 ? "." + A[1] : "")),
      /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(w) || n());
    let k = Number(w);
    Number.isFinite(k) || n("#NUM!");
    for (let F = 0; F < E; F++) k /= 100;
    return k;
  }
  if (e === "VALUE") {
    const b = t.temporal ? t.temporal(m.trim()) : parseDateTimeText(m.trim());
    if (b !== undefined) return b;
    let y = m.trim(),
      w = false;
    y.startsWith("(") &&
      y.endsWith(")") &&
      ((w = true), (y = y.slice(1, -1).trim()));
    const E = y.endsWith("%");
    (E && (y = y.slice(0, -1).trim()),
      (y = y.replace(/^\$\s*/, "")),
      w && /^[+-]/.test(y) && n(),
      /^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(
        y,
      ) || n());
    const A = (Number(y.replace(/,/g, "")) * (w ? -1 : 1)) / (E ? 100 : 1);
    return Number.isFinite(A) ? A : n("#NUM!");
  }
  if (e === "TEXT") {
    ((typeof d[1] != "string" || d[1].length > 256) && n(), h(d[0]));
    const b = t.format();
    return typeof b == "string" ? c(b) : b;
  }
  if (e === "TEXTBEFORE" || e === "TEXTAFTER") {
    const b = coerceValueToText(d[1]),
      y = f(d[2], 1),
      w = h(d[3], 0),
      E = h(d[4], 0);
    if (((!y || ![0, 1].includes(w) || ![0, 1].includes(E)) && n(), !m))
      return "";
    if ((Math.abs(y) > [...m].length && n(), !b))
      return (e === "TEXTBEFORE") == y > 0 ? "" : m;
    const A = b.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
      k = new RegExp(y < 0 ? `(?=${A})` : A, w ? "giu" : "gu");
    let F = [...m.matchAll(k)].map((O) => O.index);
    if (y < 0) {
      const O = [];
      let j = m.length;
      for (let z = F.length - 1; z >= 0; z--)
        F[z] + b.length <= j && (O.push(F[z]), (j = F[z]));
      F = O.reverse();
    }
    const I = F[y > 0 ? y - 1 : F.length + y];
    return I === undefined
      ? E && Math.abs(y) === F.length + 1
        ? (e === "TEXTBEFORE") == y > 0
          ? m
          : ""
        : d[5] !== undefined
          ? d[5]
          : n("#N/A")
      : e === "TEXTBEFORE"
        ? m.slice(0, I)
        : m.slice(I + b.length);
  }
  return notTextFunctionSentinel;
}
const maxSelectionAreas = 256;

export {
  workbookLimitTable,
  maxRowHeight,
  fileImportLimits,
  workbookLimits,
  evaluateFinancialFunction,
  maxSelectionAreas,
  notFinancialSentinel,
  minRowHeight,
  maxColumnWidth,
  supportedFontFamilies,
  evaluateTextFunction,
  defaultRowHeight,
  supportedFinancialFunctionNames,
  notTextFunctionSentinel,
  defaultColumnWidth,
  broadcastOverArrays,
  minColumnWidth,
};
export { coerceValueToText, flattenNestedValues };

/* Also used by the workers. */
export {
  compensatedSum,
  countSignChanges,
  financialArgumentRanges,
  parseDateTimeText,
  solveRateCandidates,
  supportedTextFunctionNames,
};
