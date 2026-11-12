import i from "react/jsx-runtime";
import { hr } from "./01-core-schemas-and-hashing.js";
import {
  $a,
  EU,
  JN,
  Ku,
  Rc,
  Us,
  _t,
  ct,
  ig,
  jE,
  jR,
  jw,
  pf,
  rn,
  ut,
} from "./03-workbook-model.js";
import { M } from "./04-cell-input-paste-and-fill.js";
import { X1, q1 } from "./12-formula-dependency-inspector.js";
import {
  Xv,
  bw,
  gW,
  gxe,
  lm,
  lp,
  p_,
  vw,
} from "./36-watch-window-and-goal-seek.js";

function vxe(e, r, t) {
  gW(e);
  const n = r.variables.length,
    o = new Map(),
    a = new Set();
  let s = 0;
  const d = (y) => ({ constant: Xv(y), coefficients: Array(n).fill(0) }),
    u = (y, w) => ({
      constant: Xv(y.constant * w),
      coefficients: y.coefficients.map((E) => Xv(E * w)),
    }),
    l = (y, w) => ({
      constant: Xv(y.constant + w.constant),
      coefficients: y.coefficients.map((E, A) => Xv(E + w.coefficients[A])),
    }),
    c = (y, w) => {
      if (y.coefficients.some(Boolean) && w.coefficients.some(Boolean))
        throw Error(
          "Nonlinear multiplication is outside the linear Solver subset.",
        );
      return y.coefficients.some(Boolean) ? u(y, w.constant) : u(w, y.constant);
    },
    h = (y, w, E = 0) => {
      const A = gxe(y, w);
      if (a.has(A)) throw Error("Solver formulas contain a cycle.");
      if (o.has(A)) return o.get(A);
      if (o.size + a.size >= ig.dependencyCells || E > ig.depth)
        throw Error("Solver dependency budget exceeded.");
      const k = vw(e, { sheetId: y, address: w }, t),
        F = k.cells[w],
        I = r.variables.findIndex(
          (j) => j.cell.sheetId === y && j.cell.address === w,
        );
      if (I >= 0) {
        const j = d(0);
        return ((j.coefficients[I] = 1), o.set(A, j), j);
      }
      if (!F || F.raw === null)
        throw Error(
          "Solver dependencies must be populated numeric cells; blanks and spill children are excluded.",
        );
      let O;
      if (typeof F.raw == "number") O = d(F.raw);
      else if (
        typeof F.raw == "string" &&
        F.raw.startsWith("=") &&
        !F.literal
      ) {
        if (F.raw.length > ig.expressionCharacters)
          throw Error("Solver expression budget exceeded.");
        const j = pf(F.raw, e, y).slice(1);
        if (j.length > 4096)
          throw Error("Expanded Solver expression budget exceeded.");
        (a.add(A), (O = lp(f(j, y, E + 1))), a.delete(A));
      } else
        throw Error(
          "Solver dependencies must contain numbers or supported linear formulas.",
        );
      return (o.set(A, O), O);
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
          ++s > ig.syntaxNodes)
        )
          throw Error("Solver syntax budget exceeded.");
      }
      let I = 0,
        O = 0;
      const j = () => {
          return k[I]?.text;
        },
        z = () => {
          if (++O > 32)
            throw Error("Solver expression nesting limit exceeded.");
          let _;
          const C = k[I++];
          if (!C) throw Error("Incomplete Solver formula.");
          if (C.text === "+" || C.text === "-")
            _ = u(lp(z()), C.text === "-" ? -1 : 1);
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
                  : e.sheets.find(
                      (Q) =>
                        Q.name.toLowerCase() ===
                        (U.startsWith("'")
                          ? U.slice(1, -1).replace(/''/g, "'")
                          : U
                        ).toLowerCase(),
                    )?.id;
            if (!H) throw Error("A Solver worksheet reference is missing.");
            if (G.includes(":")) {
              const Q = ct(G);
              if (
                (Q.lastRow - Q.firstRow + 1) * (Q.lastCol - Q.firstCol + 1) >
                500
              )
                throw Error("Solver range exceeds 500 cells.");
              const J = [];
              for (let q = Q.firstRow; q <= Q.lastRow; q++)
                for (let W = Q.firstCol; W <= Q.lastCol; W++)
                  J.push(h(H, `${ut(W)}${q + 1}`, E));
              _ = Object.assign(J, {
                height: Q.lastRow - Q.firstRow + 1,
                width: Q.lastCol - Q.firstCol + 1,
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
          for (; j() === "%";) (I++, (_ = u(lp(_), 0.01)));
          return (O--, _);
        },
        V = () => {
          let _ = z();
          for (; j() === "*" || j() === "/";) {
            const C = k[I++].text,
              T = lp(z());
            if (C === "*") _ = c(lp(_), T);
            else {
              if (T.coefficients.some(Boolean) || T.constant === 0)
                throw Error(
                  "Solver division requires a nonzero constant denominator.",
                );
              _ = u(lp(_), 1 / T.constant);
            }
          }
          return _;
        },
        D = () => {
          let _ = V();
          for (; j() === "+" || j() === "-";) {
            const C = k[I++].text,
              T = lp(V());
            _ = l(lp(_), u(T, C === "-" ? -1 : 1));
          }
          return _;
        },
        B = D();
      if (I !== k.length)
        throw Error("Unsupported or trailing Solver formula syntax.");
      return B;
    },
    p = r.variables.map((y) => {
      var k, F;
      const w = vw(e, y.cell, t),
        E = w.cells[y.cell.address],
        A = _t(y.cell.address);
      if (!E || typeof E.raw != "number" || !Number.isFinite(E.raw))
        throw Error(
          "Changing cells must contain finite numeric inputs, not formulas or text.",
        );
      if ((k = w.protection) != null && k.enabled)
        throw Error("A changing worksheet is protected.");
      if (
        [...(w.pivots ?? []), ...(w.queries ?? [])].some((I) =>
          bw(I.outputRange, y.cell.address),
        )
      )
        throw Error("Generated results cannot be changing cells.");
      for (const I of w.tables ?? []) {
        if (!bw(I.range, y.cell.address)) continue;
        const O = ct(I.range);
        if (
          A.row === O.firstRow ||
          (I.showTotals && A.row === O.lastRow) ||
          ((F = I.columns[A.col - O.firstCol]) != null &&
            F.calculatedFormula) ||
          Object.entries(w.cells).some(([j, z]) => {
            const V = _t(j);
            return (
              V.col === A.col &&
              V.row > O.firstRow &&
              V.row <= O.lastRow - +!!I.showTotals &&
              !z.literal &&
              typeof z.raw == "string" &&
              z.raw.startsWith("=")
            );
          })
        )
          throw Error(
            "Table headers, totals and calculated columns cannot be changing cells.",
          );
      }
      if ((w.validations ?? []).some((I) => bw(I.range, y.cell.address)))
        throw Error(
          "Changing cells with validation rules are excluded. Use explicit Solver bounds and constraints on a reviewed input copy.",
        );
      return E.raw;
    }),
    m = vw(e, r.objective, t),
    v = m.cells[r.objective.address];
  if (!v || typeof v.raw != "string" || !v.raw.startsWith("=") || v.literal)
    throw Error("The objective must be a linear formula cell.");
  const g = h(r.objective.sheetId, r.objective.address),
    b = r.constraints.map((y) => ({
      form: l(
        h(y.left.sheetId, y.left.address),
        u(
          typeof y.right == "number"
            ? d(y.right)
            : h(y.right.sheetId, y.right.address),
          -1,
        ),
      ),
      operator: y.operator,
    }));
  return (
    r.direction === "target" &&
      b.push({ form: l(g, d(-r.target)), operator: "=" }),
    { objective: g, constraints: b, baseline: p, dependencies: o.size }
  );
}
const fS = (e, r) =>
  Xv(e.constant + e.coefficients.reduce((t, n, o) => t + n * r[o], 0));
function bxe(e, r) {
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
function yxe(e, r, t = {}) {
  const n = vxe(e, r, t),
    o = r.variables.length,
    a = r.variables.every((g) => g.domain !== "continuous"),
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
        `Feasibility uses relative tolerance ${r.tolerance}; all candidates must be exhausted before an optimal label.`,
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
  for (const g of n.constraints) {
    const b = { a: g.form.coefficients, b: -g.form.constant };
    (g.operator !== ">=" && d.push(b),
      g.operator !== "<=" && d.push({ a: b.a.map((y) => -y), b: -b.b }));
  }
  r.variables.forEach((g, b) => {
    const y = Array(o).fill(0);
    ((y[b] = 1),
      d.push({ a: y, b: g.upper }, { a: y.map((w) => -w), b: -g.lower }));
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
      return y - b.b <= r.tolerance * w;
    });
  let l,
    c = 0,
    h = false,
    f = false;
  const p = (g) => {
      if (!u(g)) return;
      const b = fS(n.objective, g);
      (!l ||
        (r.direction === "maximize" && b > c) ||
        (r.direction === "minimize" && b < c)) &&
        ((l = [...g]), (c = b));
    },
    m = () =>
      s.iterations >= r.maxCandidates
        ? ((f = true), false)
        : (s.iterations++, true);
  if (a)
    if (
      r.variables
        .map((b) => Math.max(0, Math.floor(b.upper) - Math.ceil(b.lower) + 1))
        .reduce((b, y) => b * y, 1) > r.maxCandidates
    )
      f = true;
    else {
      const b = (y) => {
        if (y.length === o) {
          m() && p(y);
          return;
        }
        const w = r.variables[y.length];
        for (let E = Math.ceil(w.lower); E <= Math.floor(w.upper); E++)
          b([...y, E]);
      };
      b([]);
    }
  else {
    let g = 1;
    for (let b = 1; b <= o; b++) g = (g * (d.length - o + b)) / b;
    if (g > r.maxCandidates) f = true;
    else {
      const b = (y, w) => {
        if (y.length === o) {
          if (!m()) return;
          const E = y.map((k) => d[k]),
            A = bxe(
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
  const v = l.map((g) => (Object.is(g, -0) ? 0 : g));
  return (
    v.forEach((g, b) => {
      var E;
      const y = r.variables[b].cell;
      if (
        (E = e.sheets.find((A) => A.id === y.sheetId).validations) != null &&
        E.some((A) => bw(A.range, y.address) && !jw(A, g))
      )
        throw Error("The solution violates a changing-cell validation rule.");
    }),
    {
      ...s,
      status: "optimal",
      success: true,
      message:
        r.direction === "target"
          ? "A feasible solution reaches the target within tolerance."
          : "All bounded candidates were examined; the best feasible objective is reported within tolerance.",
      solution: v,
      objective: c,
      rows: [
        ["Changing cell", "Baseline", "Solution"],
        ...r.variables.map((g, b) => [
          `${e.sheets.find((y) => y.id === g.cell.sheetId).name}!${g.cell.address}`,
          n.baseline[b],
          v[b],
        ]),
        ["Objective", fS(n.objective, n.baseline), c],
        ["Constraint", "Operator", "Residual (left − right)"],
        ...n.constraints.map((g, b) => [
          `Constraint ${b + 1}`,
          g.operator,
          fS(g.form, v),
        ]),
      ],
    }
  );
}
const Ms = (e) => {
    if (!Number.isFinite(e))
      throw Error(
        "The statistical result exceeds finite numeric precision. Rescale the sample.",
      );
    return e;
  },
  vW = (e) => {
    const r = Us(e);
    return Ms(Rc(e.map((t) => (t - r) ** 2)) / (e.length - 1));
  },
  wxe = new JN({
    onCell: () => {
      throw Error("Distribution primitives cannot read workbook cells.");
    },
    onRange: () => {
      throw Error("Distribution primitives cannot read workbook ranges.");
    },
  });
function my(e, r) {
  const t = wxe.parse(
    `${e}(${r.map((n) => (typeof n == "boolean" ? String(n).toUpperCase() : String(n))).join(",")})`,
    { sheet: "Primitives", row: 1, col: 1 },
  );
  if (typeof t != "number" || !Number.isFinite(t))
    throw Error(`The ${e} distribution primitive failed.`);
  return t;
}
function oj(e, r) {
  const t = r[0].length;
  if (e.length <= t + 1)
    throw Error("Regression needs more observations than fitted coefficients.");
  if (vW(e) === 0) throw Error("Regression requires a nonconstant response.");
  const n = jE("LINEST", [e.map((l) => [l]), r, true, true], {
      retrieve: (l) => l,
      issue: (l) => new Error(l),
      isIssue: (l) => l instanceof Error,
    }),
    o = n[3][1];
  if (o !== e.length - t - 1)
    throw Error(
      "Regression predictors are rank deficient or numerically dependent.",
    );
  const a = [...n[0]].reverse(),
    s = [...n[1]].reverse(),
    d = jE("TREND", [e.map((l) => [l]), r, r, true], {
      retrieve: (l) => l,
      issue: (l) => new Error(l),
      isIssue: (l) => l instanceof Error,
    }).map((l) => l[0]),
    u = e.map((l, c) => Ms(l - d[c]));
  return {
    coefficient: a,
    standardErrors: s,
    df: o,
    r2: n[2][0],
    standardError: n[2][1],
    f: n[3][0],
    ssRegression: n[4][0],
    ssError: n[4][1],
    fitted: d,
    residuals: u,
  };
}
function xxe(e, r, t = {}) {
  const n = ct(r.range),
    o = n.lastCol - n.firstCol + 1,
    a = n.lastRow - n.firstRow;
  if (
    !/^[A-Z]{1,3}[1-9]\d{0,6}:[A-Z]{1,3}[1-9]\d{0,6}$/.test(r.range) ||
    a < 1 ||
    a > 500 ||
    o > 6 ||
    (a + 1) * o > 2e3
  )
    throw Error(
      "Select a headed rectangle with 1–500 data rows, 1–6 columns and at most 2,000 cells.",
    );
  (vw(e, { sheetId: r.sheetId, address: r.range.split(":")[0] }, t), gW(e));
  const s = X1(e, [{ sheetId: r.sheetId, range: r.range }], t),
    d = s.snapshots[0];
  for (const b of s.dependencyCells)
    vw(e, { sheetId: b.sheetId, address: b.address }, t);
  const u = [],
    l = [];
  for (const b of d.records)
    if (b.values.some((y) => typeof y != "number" || !Number.isFinite(y))) {
      if (r.missing === "reject")
        throw Error(
          `Row ${b.row} includes blank or nonnumeric values. Correct it or explicitly exclude incomplete rows.`,
        );
      u.push({
        row: b.row,
        locator: b.locator,
        reason:
          "Complete-case exclusion: at least one blank, text or boolean value.",
      });
    } else l.push({ ...b, values: b.values });
  const c = l.length,
    h = l.map((b) => b.values),
    f = d.headers,
    p = () => {
      if (o !== 1)
        throw Error("This analysis requires exactly one numeric column.");
      return h.map((b) => b[0]);
    };
  if (c < (r.method === "histogram" || r.method === "distribution" ? 1 : 2))
    throw Error("Insufficient numeric observations after exclusions.");
  const m = [
    "Headed numeric inputs; no text-to-number or date conversion. Complete-case exclusions apply to every selected column.",
    "Hidden cells require explicit authorization; calculation errors, dynamic/volatile formulas and generated report inputs are refused.",
    "Binary64 arithmetic; local MIT implementation and the existing pinned MIT fast-formula-parser distribution primitives. No proprietary engine or new license dependency.",
  ];
  let v = "",
    g = [];
  if (r.method === "descriptive") {
    ((v = "Two-pass sample moments and sorted inclusive quartiles"),
      m.push(
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
      y = f.map((w, E) => {
        const A = h.map((O) => O[E]),
          k = [...A].sort((O, j) => O - j),
          F = vW(A),
          I = (O) => {
            const j = (c - 1) * O,
              z = Math.floor(j);
            return k[z] + (k[Math.ceil(j)] - k[z]) * (j - z);
          };
        return [
          c,
          Ms(Rc(A)),
          Ms(Us(A)),
          Ms(Math.sqrt(F / c)),
          I(0.5),
          Ms(Math.sqrt(F)),
          F,
          k[0],
          I(0.25),
          I(0.75),
          k[c - 1],
          Ms(k[c - 1] - k[0]),
        ];
      });
    g = [["Statistic", ...f], ...b.map((w, E) => [w, ...y.map((A) => A[E])])];
  } else if (r.method === "correlation" || r.method === "covariance") {
    if (o < 2) throw Error("Select at least two columns for a matrix.");
    ((v =
      r.method === "correlation"
        ? "Pearson product-moment correlation"
        : "Two-pass sample covariance"),
      m.push(
        "Every matrix entry uses the same complete-case observations; covariance uses n − 1.",
        "Independent reference: analytic matrices and Python statistics in ANALYSIS-VERIFICATION.json.",
      ));
    const b = f.map((w, E) => Us(h.map((A) => A[E]))),
      y = f.map((w, E) => Ms(Rc(h.map((A) => (A[E] - b[E]) ** 2))));
    if (r.method === "correlation" && y.some((w) => w === 0))
      throw Error("Correlation is undefined for a constant column.");
    g = [
      [r.method, ...f],
      ...f.map((w, E) => [
        w,
        ...f.map((A, k) => {
          const F = Ms(Rc(h.map((I) => (I[E] - b[E]) * (I[k] - b[k]))));
          return r.method === "correlation"
            ? Math.max(
                -1,
                Math.min(1, Ms(F / Math.sqrt(y[E]) / Math.sqrt(y[k]))),
              )
            : Ms(F / (c - 1));
        }),
      ]),
    ];
  } else if (r.method === "regression") {
    if (o < 2)
      throw Error(
        "Regression requires Y in the first column and one or more X columns.",
      );
    const b = oj(
        h.map((A) => A[0]),
        h.map((A) => A.slice(1)),
      ),
      y = my("T.INV.2T", [0.05, b.df]);
    ((v = "Ordinary least squares with intercept; centered, scaled pivoted QR"),
      m.push(
        "First column is Y; remaining columns are X. Independent errors, linear mean and constant error variance; t/F inference additionally assumes normal errors.",
        "No automatic causal interpretation. Rank deficient and insufficient samples are refused. Perfect fits have undefined t/F values shown explicitly.",
        "Independent certified reference: NIST StRD Norris, 36 observations; see ANALYSIS-VERIFICATION.json.",
      ));
    const w = ["Intercept", ...f.slice(1)].map((A, k) => {
        const F = b.standardErrors[k],
          I = b.coefficient[k],
          O = F === 0 ? null : Ms(I / F);
        return [
          A,
          I,
          F,
          O,
          O === null
            ? null
            : my("BETA.DIST", [b.df / (b.df + O * O), b.df / 2, 0.5, true]),
          I - y * F,
          I + y * F,
        ];
      }),
      E = typeof b.f == "number" ? b.f : null;
    g = [
      ["Metric", "Value"],
      ["Observations", c],
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
          : my("BETA.DIST", [
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
  } else if (r.method === "histogram") {
    const b = p(),
      y = r.bins;
    if (!(y != null && y.length) || y.some((A, k) => k > 0 && A <= y[k - 1]))
      throw Error(
        "Histogram bin upper bounds must be finite and strictly increasing.",
      );
    ((v = "Explicit right-closed histogram bins"),
      m.push(
        "First bin includes all values at or below its upper bound; later bins are (previous, upper]; last bin is overflow.",
        "Independent reference: exact integer counting fixtures in ANALYSIS-VERIFICATION.json.",
      ));
    const w = Array(y.length + 1).fill(0);
    b.forEach((A) => {
      const k = y.findIndex((F) => A <= F);
      w[k < 0 ? y.length : k]++;
    });
    let E = 0;
    g = [
      [
        "Upper bound",
        "Frequency",
        "Relative frequency",
        "Cumulative frequency",
      ],
      ...w.map(
        (A, k) => (
          (E += A),
          [k === y.length ? "Overflow" : y[k], A, A / c, E / c]
        ),
      ),
    ];
  } else if (r.method === "distribution") {
    const b = p();
    ((v =
      "Pinned analytic probability density/mass and cumulative distribution"),
      m.push(
        "Selected cells are evaluation points, not a fitted sample. Distribution parameters are supplied explicitly.",
        "Independent reference: Python NormalDist, math.comb and factorial probability sums in ANALYSIS-VERIFICATION.json.",
      ));
    const y = r.distribution ?? "normal";
    let w, E;
    if (y === "normal") {
      if (r.mean === undefined || r.standardDeviation === undefined)
        throw Error(
          "Normal distribution needs a finite mean and positive standard deviation.",
        );
      ((E = "NORM.DIST"),
        (w = (A, k) => [A, r.mean, r.standardDeviation, k]),
        m.push(
          `Normal mean ${r.mean}; standard deviation ${r.standardDeviation}.`,
        ));
    } else if (y === "binomial") {
      if (
        r.trials === undefined ||
        r.probability === undefined ||
        b.some((A) => !Number.isInteger(A) || A < 0 || A > r.trials)
      )
        throw Error(
          "Binomial points must be integers from zero through the explicit trial count; provide p in [0,1].",
        );
      ((E = "BINOM.DIST"),
        (w = (A, k) => [A, r.trials, r.probability, k]),
        m.push(
          `Independent trials ${r.trials}; success probability ${r.probability}.`,
        ));
    } else {
      if (
        r.lambda === undefined ||
        b.some((A) => !Number.isInteger(A) || A < 0 || A > 1e3)
      )
        throw Error(
          "Poisson points must be integers from zero through 1,000; provide a positive rate.",
        );
      ((E = "POISSON.DIST"),
        (w = (A, k) => [A, r.lambda, k]),
        m.push(
          `Poisson rate ${r.lambda}; independent arrivals and constant rate.`,
        ));
    }
    g = [
      [
        "Point",
        y === "normal" ? "Density" : "Probability mass",
        "Cumulative probability",
      ],
      ...b.map((A) => [A, my(E, w(A, false)), my(E, w(A, true))]),
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
    const E = r.horizon ?? 6,
      A = r.forecast ?? "linear",
      k = Array.from({ length: E }, (j, z) => Ms(b[c - 1] + w * (z + 1))),
      F = [],
      I = [];
    if (A === "linear") {
      const j = oj(
        y,
        b.map((V) => [V]),
      );
      v = "Linear trend forecast using centered QR least squares";
      const z = jE(
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
      const j = r.window ?? 3;
      if (j >= c)
        throw Error("Moving average needs more observations than its window.");
      v = `Trailing ${j}-period moving average with recursive forecasts`;
      for (let V = j; V < c; V++)
        I.push({ index: V, value: Us(y.slice(V - j, V)) });
      const z = [...y];
      for (let V = 0; V < E; V++) {
        const D = Ms(Us(z.slice(-j)));
        (F.push(D), z.push(D));
      }
    } else {
      if (c < 3)
        throw Error("Exponential smoothing needs at least three observations.");
      const j = r.alpha ?? 0.3;
      v =
        "Simple exponential smoothing with fixed alpha and first-observation initialization";
      let z = y[0];
      for (let V = 1; V < c; V++)
        (I.push({ index: V, value: z }), (z = Ms(j * y[V] + (1 - j) * z)));
      (F.push(...k.map(() => z)),
        m.push(`Alpha ${j}; no trend or seasonal component.`));
    }
    const O = I.map((j) => Ms(y[j.index] - j.value));
    (m.push(
      "Numeric equally spaced periods; no seasonality, automatic tuning, ETS intervals or causal claims. Extrapolation can be unreliable.",
      "Error metrics use fitted/in-sample observations, not an independent holdout. No uncertainty interval is claimed.",
      "Independent reference: analytic forecasts and Python calculations in ANALYSIS-VERIFICATION.json.",
    ),
      (g = [
        ["Metric", "Value"],
        ["Fitted observations", I.length],
        ["In-sample MAE", Us(O.map(Math.abs))],
        ["In-sample RMSE", Math.sqrt(Us(O.map((j) => Ms(j * j))))],
        ["Period", "Forecast"],
        ...k.map((j, z) => [j, Ms(F[z])]),
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
  for (const b of g) for (const y of b) typeof y == "number" && Ms(y);
  return {
    status: "complete",
    success: true,
    message:
      "Calculation completed for the declared method and included observations. Review the assumptions and exclusions before saving.",
    algorithm: v,
    assumptions: m,
    exclusions: u,
    included: c,
    inspected: d.records.length,
    iterations: 0,
    rows: g,
  };
}
const WT = new WeakMap();
function Exe(e, r, t = {}) {
  const n = EU(r),
    o = lm(e, t),
    a = n.kind === "solver" ? yxe(e, n, t) : xxe(e, n, t),
    s = {
      request: n,
      result: a,
      workbookHash: hr(e),
      contextHash: hr(o),
      sourceHash: p_(e),
    };
  return (WT.set(s, hr(s)), s);
}
function bW(e) {
  const r = e.result;
  return [
    [
      e.request.kind === "solver"
        ? "Solver report"
        : "Statistical analysis report",
    ],
    ["Outcome", r.status],
    ["Message", r.message],
    ["Algorithm", r.algorithm],
    ["Engine", jR],
    [
      "License",
      "MIT; existing MIT distribution dependency; no proprietary engine",
    ],
    ["Inspected", r.inspected],
    ["Included", r.included],
    ["Excluded", r.exclusions.length],
    ["Candidates examined", r.iterations],
    ...r.assumptions.map((t, n) => [`Assumption ${n + 1}`, t]),
    [],
    ...r.rows,
    [],
    ["Excluded source row", "Reason", "Locator"],
    ...r.exclusions.map((t) => [t.row, t.reason, t.locator]),
  ];
}
function HT(e, r) {
  const t = ct(r);
  if ((t.lastRow - t.firstRow + 1) * (t.lastCol - t.firstCol + 1) > 2e4)
    throw Error("Analysis output exceeds 20,000 cells.");
  const n = [];
  for (let o = t.firstRow; o <= t.lastRow; o++)
    for (let a = t.firstCol; a <= t.lastCol; a++) {
      const s = e.cells[`${ut(a)}${o + 1}`];
      n.push({
        raw: s?.raw ?? null,
        literal: !!(s != null && s.literal),
      });
    }
  return hr(n);
}
function Cxe(e, r, t, n = {}) {
  const o = lm(e, n);
  if (o.readOnly) throw Error("This workbook is read-only.");
  if (($a(e), !["report", "apply-solution"].includes(t)))
    throw Error("Choose Save report or Apply solution and report.");
  if (
    WT.get(r) !== hr(r) ||
    r.workbookHash !== hr(e) ||
    r.contextHash !== hr(o)
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
    e.sheets.some(
      (h) => h.name.toLowerCase() === r.request.outputName.toLowerCase(),
    )
  )
    throw Error("The output worksheet name is already in use.");
  q1(e, r.request.outputName);
  const a = rn(e);
  t === "apply-solution" &&
    r.request.kind === "solver" &&
    r.request.variables.forEach((h, f) => {
      const p = a.sheets.find((m) => m.id === h.cell.sheetId);
      p.cells[h.cell.address] = {
        ...p.cells[h.cell.address],
        raw: r.result.solution[f],
      };
    });
  const s = Ku(r.request.outputName),
    d = bW(r),
    u = Math.max(...d.map((h) => h.length));
  if (d.length * u > 2e4)
    throw Error("Analysis report exceeds the output budget.");
  (d.forEach((h, f) =>
    h.forEach((p, m) => {
      s.cells[`${ut(m)}${f + 1}`] = {
        raw: p,
        ...(typeof p == "string" ? { literal: true } : {}),
        ...(f === 0
          ? { style: { bold: true, fill: "#362C63", color: "#FFFFFF" } }
          : {}),
      };
    }),
  ),
    (s.columnWidths = {
      0: 210,
      1: 260,
      2: 180,
      3: 160,
      4: 140,
      5: 140,
      6: 140,
    }),
    (s.frozenRows = 1));
  const l = `A1:${ut(u - 1)}${d.length}`;
  ((s.analysisReport = {
    engine: jR,
    request: structuredClone(r.request),
    result: structuredClone(r.result),
    definitionHash: hr(r.request),
    sourceHash: p_(a),
    outputHash: HT(s, l),
    outputRange: l,
    createdAt: new Date().toISOString(),
    sourceRevision: e.revision,
    includeHidden: !!n.includeHidden,
    applied: t === "apply-solution",
  }),
    a.sheets.push(s));
  const c = rn(a);
  return (WT.delete(r), c);
}
function Axe(e, r, t = {}) {
  const n = e.sheets.find((d) => d.id === r),
    o = n?.analysisReport;
  if (!n || !o)
    return { status: "unavailable", message: "No saved analysis report." };
  const a = lm(e, t);
  if (
    !a.allowedSheetIds.includes(r) ||
    (!a.includeHidden && (n.hidden || n.veryHidden))
  )
    return {
      status: "unavailable",
      message: "The report is outside the authorized context.",
    };
  if (
    (o.includeHidden && !a.includeHidden) ||
    a.allowedSheetIds.length < e.sheets.length
  )
    return {
      status: "unavailable",
      message:
        "Reopen the original analysis within the authorized source context before trusting this historical report.",
    };
  if (
    o.definitionHash !== hr(o.request) ||
    o.sourceHash !== p_(e) ||
    o.outputHash !== HT(n, o.outputRange)
  )
    return {
      status: "stale",
      message:
        "Source data, calculation settings or report cells changed. Preview the analysis again; saved values are historical.",
    };
  const s = Ku("Expected");
  return (
    bW(o).forEach((d, u) =>
      d.forEach(
        (l, c) =>
          (s.cells[`${ut(c)}${u + 1}`] = {
            raw: l,
            ...(typeof l == "string" ? { literal: true } : {}),
          }),
      ),
    ),
    HT(s, o.outputRange) !== o.outputHash
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
const Oi = (e, r) =>
    i.jsxs("label", { children: [i.jsx(M, { children: e }), r] }),
  aj = (e) => (e.trim() === "" ? NaN : Number(e));

export { Axe, Cxe, Exe, Oi, aj };
