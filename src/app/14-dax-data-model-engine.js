import i from "react/jsx-runtime";
import ne from "react";
import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import {
  assertWorkbookStructureEditable,
  daxFunctionNames,
  walkDaxSyntaxTree,
  createWorksheetRecord,
  DaxError,
  columnarModelFormatVersion,
  formatOccupiedRangeAddress,
  parseDaxExpression,
  validateModelQuery,
  validateDataModel,
  parseRangeAddress,
  throwDaxError,
  normalizeWorkbookInput,
  daxExpressionLimits,
  dataModelLimits,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  readSourceHeaders,
  inspectOutputNameSafety,
} from "./12-consolidation-and-dependency-inspection.js";
import {
  requireNumericOperand,
  requireLogicalOperand,
  buildTableColumnKey,
  requireScalarValue,
  requireFiniteNumber,
  sumWithCompensation,
  compareDaxValues,
  caseFoldText,
  buildTypedValueKey,
} from "./13-consolidation-dialog.js";

function evaluateDaxExpression(e) {
  const r = e.definition,
    t = new Map(r.tables.map((N) => [N.id, N])),
    n = new Map(),
    o = new Map(),
    a = new Set(),
    s = [],
    d = new WeakMap(),
    u = (N) =>
      r.tables.find((R) => caseFoldText(R.name) === caseFoldText(N)) ??
      throwDaxError("DAX_REFERENCE", `Unknown model table ${N}.`),
    l = (N) => {
      if (N.kind !== "field")
        return throwDaxError(
          "DAX_ARGUMENT",
          "A fully qualified model column is required.",
        );
      const R = u(N.table),
        L = R.columns.find(
          ($) => caseFoldText($.name) === caseFoldText(N.column),
        );
      return L
        ? { tableId: R.id, column: L.name }
        : throwDaxError(
            "DAX_REFERENCE",
            `Unknown model column ${N.table}[${N.column}].`,
          );
    },
    c = (N) =>
      r.measures.find((R) => caseFoldText(R.name) === caseFoldText(N)) ??
      throwDaxError(
        "DAX_REFERENCE",
        `Unknown measure [${N}]. Use qualified table[column] references for columns.`,
      );
  for (const N of r.tables)
    for (const R of N.columns)
      R.expression !== undefined &&
        n.set(
          buildTableColumnKey({ tableId: N.id, column: R.name }),
          parseDaxExpression(R.expression),
        );
  for (const N of r.measures)
    N.operation === "dax" &&
      n.set(`measure:${N.id}`, parseDaxExpression(N.expression));
  for (const N of n.values())
    walkDaxSyntaxTree(N, (R) => {
      (R.kind === "field" && l(R),
        R.kind === "table" && u(R.name),
        R.kind === "measure" && c(R.name));
    });
  const h = (N) =>
      N.kind === "table" ||
      (N.kind === "call" && ["FILTER", "ALL"].includes(N.name)),
    f = (N) =>
      (N.kind === "literal" && typeof N.value == "number") ||
      (N.kind === "unary" &&
        ["+", "-"].includes(N.op) &&
        N.value.kind === "literal" &&
        typeof N.value.value == "number") ||
      (N.kind === "call" && N.name === "BLANK");
  for (const N of n.values())
    walkDaxSyntaxTree(N, (R) => {
      if (R.kind !== "call") return;
      const L = ($) => {
        R.args[$].kind !== "field" &&
          throwDaxError(
            "DAX_ARGUMENT",
            `${R.name} requires a qualified table[column] reference.`,
          );
      };
      if (
        ([
          "SUM",
          "AVERAGE",
          "MIN",
          "MAX",
          "DISTINCTCOUNT",
          "SELECTEDVALUE",
          "RELATED",
        ].includes(R.name) && L(0),
        ["SUM", "AVERAGE", "MIN", "MAX"].includes(R.name))
      ) {
        const $ = l(R.args[0]);
        t.get($.tableId).columns.find((U) => U.name === $.column).type !==
          "number" &&
          throwDaxError(
            "DAX_TYPE",
            `${R.name} requires a number column in this subset.`,
          );
      }
      (R.name === "TOTALYTD" && L(1),
        R.name === "ALL" &&
          !["table", "field"].includes(R.args[0].kind) &&
          throwDaxError(
            "DAX_ARGUMENT",
            "ALL accepts one table or qualified column.",
          ),
        ["FILTER", "SUMX", "COUNTROWS"].includes(R.name) &&
          !h(R.args[0]) &&
          throwDaxError(
            "DAX_ARGUMENT",
            `${R.name} requires a table, FILTER or ALL expression.`,
          ),
        R.name === "DIVIDE" &&
          R.args[2] &&
          !f(R.args[2]) &&
          throwDaxError(
            "DAX_ARGUMENT",
            "DIVIDE alternate result must be a numeric literal or BLANK().",
          ),
        R.args.forEach(($, U) => {
          if (!h($)) return;
          (["FILTER", "SUMX", "COUNTROWS", "ALL"].includes(R.name) &&
            U === 0) ||
            (R.name === "CALCULATE" && U > 0) ||
            throwDaxError(
              "DAX_ARGUMENT",
              `${R.name} requires a scalar at argument ${U + 1}.`,
            );
        }));
    });
  const p = (N) => {
      const R = [];
      return (
        walkDaxSyntaxTree(N, (L) => {
          (L.kind === "field" && R.push(buildTableColumnKey(l(L))),
            L.kind === "measure" && R.push(`measure:${c(L.name).id}`));
        }),
        R
      );
    },
    m = new Set(),
    v = new Set(),
    g = (N) => {
      if (
        (v.has(N) &&
          throwDaxError(
            "DAX_CYCLE",
            "Calculated columns or measures contain a circular dependency.",
          ),
        m.has(N))
      )
        return;
      const R = n.get(N);
      R && (v.add(N), p(R).forEach(g), v.delete(N), m.add(N));
    };
  n.forEach((N, R) => g(R));
  const b = (N) => Array.from({ length: e.count(N) }, (R, L) => L),
    y = (N, R) => {
      e.tick();
      const L = t.get(N.tableId),
        $ = L.columns.find((Q) => Q.name === N.column);
      if ($.expression === undefined) return e.value(N.tableId, N.column, R);
      const U = buildTableColumnKey(N),
        G = `${U}:${R}`;
      let H = o.get(U);
      if (
        (H || ((H = Array(e.count(N.tableId))), o.set(U, H)),
        Object.prototype.hasOwnProperty.call(H, R))
      )
        return H[R];
      if (a.has(G) || a.size > 48)
        return throwDaxError(
          "DAX_CYCLE",
          "A calculated column has a circular or excessively deep evaluation dependency.",
        );
      a.add(G);
      try {
        const Q = requireScalarValue(
          C(n.get(U), {
            filters: [],
            rows: [{ tableId: N.tableId, index: R }],
          }),
        );
        return (
          isValidColumnValue(Q, $) ||
            throwDaxError(
              "DAX_TYPE",
              `${L.name}[${$.name}] row ${R + 1} does not match its declared ${$.type}${$.nullable ? " or BLANK" : ""} type.`,
            ),
          (H[R] = Q),
          Q
        );
      } finally {
        a.delete(G);
      }
    },
    w = (N) => {
      const R = d.get(N.filters);
      if (R) return R;
      const L = new Map(
        r.tables.map(($) => [$.id, new Uint8Array(e.count($.id)).fill(1)]),
      );
      for (const $ of N.filters) {
        const U = L.get($.tableId);
        for (let G = 0; G < U.length; G++)
          (e.tick(), U[G] && !$.test(G) && (U[G] = 0));
      }
      for (const $ of e.order)
        for (const U of r.relationships.filter((G) => G.one.tableId === $)) {
          const G = L.get($),
            H = L.get(U.many.tableId);
          for (let Q = 0; Q < H.length; Q++) {
            e.tick();
            const X = e.parentRow(U.many.tableId, $, Q);
            (X === undefined || !G[X]) && (H[Q] = 0);
          }
        }
      return (d.set(N.filters, L), L);
    },
    E = (N, R) => b(N).filter((L) => w(R).get(N)[L]),
    A = (N, R) =>
      N.flatMap((L) => {
        const $ = L.columns.filter(
          (U) => !R.has(buildTableColumnKey({ tableId: L.tableId, column: U })),
        );
        return $.length === L.columns.length
          ? [L]
          : $.length
            ? L.project
              ? [L.project($)]
              : throwDaxError(
                  "DAX_FILTER",
                  "This compound filter cannot be partially replaced. Use single-column filters.",
                )
            : [];
      }),
    k = (N, R) => {
      const L = new Set(
        R.flatMap(($) =>
          $.columns.map((U) =>
            buildTableColumnKey({ tableId: $.tableId, column: U }),
          ),
        ),
      );
      return [...A(N, L), ...R];
    },
    F = (N) => r.tables.filter((R) => e.canFilter(R.id, N)),
    I = (N, R, L) => {
      const $ = L.map((G) =>
          R.map((H) => buildTypedValueKey(y({ tableId: N, column: H }, G))),
        ),
        U = new Set($.map((G) => JSON.stringify(G)));
      return {
        tableId: N,
        columns: R,
        test: (G) =>
          U.has(
            JSON.stringify(
              R.map((H) => buildTypedValueKey(y({ tableId: N, column: H }, G))),
            ),
          ),
        project: (G) => I(N, G, L),
      };
    },
    O = (N) =>
      (N.columns ? [t.get(N.tableId)] : F(N.tableId)).map((L) => {
        const $ =
            N.columns ??
            L.columns
              .filter((G) => G.expression === undefined)
              .map((G) => G.name),
          U = N.rows
            .map((G) => e.parentRow(N.tableId, L.id, G))
            .filter((G) => G !== undefined);
        return I(L.id, $, U);
      }),
    j = (N) => {
      if (N.transitioned || !N.rows.length) return N;
      let R = [...N.filters];
      for (const L of N.rows)
        R = k(
          R,
          O({ tableId: L.tableId, rows: [L.index], columns: L.columns }),
        );
      return { ...N, filters: R, transitioned: true };
    },
    z = (N, R, L) => ({
      ...N,
      rows: [...N.rows, { tableId: R.tableId, index: L, columns: R.columns }],
      transitioned: false,
    }),
    V = (N, R) => {
      if (R.kind === "field") {
        const $ = l(R);
        return A(N, new Set([buildTableColumnKey($)]));
      }
      if (R.kind !== "table")
        return throwDaxError(
          "DAX_ARGUMENT",
          "ALL accepts one physical table or qualified column.",
        );
      const L = new Set(F(u(R.name).id).map(($) => $.id));
      return N.filter(($) => !L.has($.tableId));
    },
    D = (N) =>
      typeof N == "object" && N !== null
        ? N
        : throwDaxError("DAX_TYPE", "A model table expression is required."),
    B = (N, R) => {
      if (s.includes(N.id) || s.length >= 48)
        return throwDaxError("DAX_CYCLE", `Circular measure [${N.name}].`);
      s.push(N.id);
      try {
        const L = j(R);
        if (N.operation === "dax")
          return requireScalarValue(C(n.get(`measure:${N.id}`), L));
        const $ = E(N.tableId, L),
          U = $.map((H) =>
            N.operation === "countRows"
              ? 1
              : y({ tableId: N.tableId, column: N.column }, H),
          ),
          G = U.filter((H) => H !== null);
        return N.operation === "count" || N.operation === "countRows"
          ? G.length
          : N.operation === "distinctCount"
            ? new Set(G.map((H) => JSON.stringify([typeof H, H]))).size
            : G.length
              ? N.operation === "sum"
                ? sumWithCompensation(G)
                : N.operation === "average"
                  ? requireFiniteNumber(sumWithCompensation(G) / G.length)
                  : N.operation === "min"
                    ? Math.min(...G)
                    : Math.max(...G)
              : null;
      } finally {
        s.pop();
      }
    },
    _ = (N, R) => {
      const L = l(N);
      return { f: L, values: E(L.tableId, R).map(($) => y(L, $)) };
    };
  function C(N, R) {
    if ((e.tick(), N.kind === "literal")) return N.value;
    if (N.kind === "table") {
      const U = u(N.name);
      return { kind: "rows", tableId: U.id, rows: E(U.id, R) };
    }
    if (N.kind === "measure") return B(c(N.name), R);
    if (N.kind === "field") {
      const U = l(N),
        G = [...R.rows]
          .reverse()
          .find(
            (H) =>
              H.tableId === U.tableId &&
              (!H.columns || H.columns.includes(U.column)),
          );
      return G
        ? y(U, G.index)
        : throwDaxError(
            "DAX_ROW",
            `No row context for ${N.table}[${N.column}]. Use an aggregate or SELECTEDVALUE in a measure.`,
          );
    }
    if (N.kind === "unary") {
      const U = requireScalarValue(C(N.value, R));
      return N.op === "NOT"
        ? !requireLogicalOperand(U)
        : N.op === "+"
          ? U
          : -requireNumericOperand(U);
    }
    if (N.kind === "binary") {
      const U = requireScalarValue(C(N.left, R));
      if (N.op === "&&" && !requireLogicalOperand(U)) return false;
      if (N.op === "||" && requireLogicalOperand(U)) return true;
      const G = requireScalarValue(C(N.right, R));
      if (N.op === "&&" || N.op === "||") return requireLogicalOperand(G);
      if (N.op === "&")
        return `${U === null ? "" : typeof U == "boolean" ? String(U).toUpperCase() : U}${G === null ? "" : typeof G == "boolean" ? String(G).toUpperCase() : G}`;
      if (["=", "==", "<>", "<", ">", "<=", ">="].includes(N.op)) {
        const X = compareDaxValues(U, G, N.op === "==");
        return N.op === "=" || N.op === "=="
          ? X === 0
          : N.op === "<>"
            ? X !== 0
            : N.op === "<"
              ? X !== undefined && X < 0
              : N.op === ">"
                ? X !== undefined && X > 0
                : N.op === "<="
                  ? X !== undefined && X <= 0
                  : X !== undefined && X >= 0;
      }
      if (
        (N.op === "*" && (U === null || G === null)) ||
        (N.op === "/" && U === null) ||
        (U === null && G === null && ["+", "-"].includes(N.op))
      )
        return null;
      const H = requireNumericOperand(U),
        Q = requireNumericOperand(G);
      return requireFiniteNumber(
        N.op === "+"
          ? H + Q
          : N.op === "-"
            ? H - Q
            : N.op === "*"
              ? H * Q
              : N.op === "/"
                ? H / Q
                : H ** Q,
      );
    }
    const L = N.args,
      $ = (U) => requireScalarValue(C(L[U], R));
    switch (N.name) {
      case "BLANK":
        return null;
      case "TRUE":
        return true;
      case "FALSE":
        return false;
      case "ISBLANK":
        return $(0) === null;
      case "IF":
        return requireLogicalOperand($(0)) ? $(1) : L[2] ? $(2) : null;
      case "COALESCE":
        for (let U = 0; U < L.length; U++) {
          const G = $(U);
          if (G !== null) return G;
        }
        return null;
      case "DIVIDE": {
        if (L[2] && !f(L[2]))
          return throwDaxError(
            "DAX_ARGUMENT",
            "DIVIDE alternate result must be a numeric literal or BLANK().",
          );
        const U = $(0),
          G = $(1);
        return G === null || requireNumericOperand(G) === 0
          ? L[2]
            ? $(2)
            : null
          : U === null
            ? null
            : requireFiniteNumber(
                requireNumericOperand(U) / requireNumericOperand(G),
              );
      }
      case "SUM":
      case "AVERAGE":
      case "MIN":
      case "MAX": {
        const { f, values } = _(L[0], R);
        if (
          t.get(f.tableId).columns.find((Q) => Q.name === f.column).type !==
          "number"
        )
          return throwDaxError(
            "DAX_TYPE",
            `${N.name} requires a number column in this subset.`,
          );
        const H = values.filter((Q) => typeof Q == "number");
        return H.length
          ? N.name === "SUM"
            ? sumWithCompensation(H)
            : N.name === "AVERAGE"
              ? requireFiniteNumber(sumWithCompensation(H) / H.length)
              : N.name === "MIN"
                ? Math.min(...H)
                : Math.max(...H)
          : null;
      }
      case "DISTINCTCOUNT": {
        const { values } = _(L[0], R);
        return values.length
          ? new Set(values.map(buildTypedValueKey)).size
          : null;
      }
      case "SELECTEDVALUE": {
        const { values } = _(L[0], R),
          G = new Map(values.map((H) => [buildTypedValueKey(H), H]));
        return G.size === 1 ? [...G.values()][0] : L[1] ? $(1) : null;
      }
      case "COUNTROWS":
        return D(C(L[0], R)).rows.length || null;
      case "SUMX": {
        const U = D(C(L[0], R));
        return sumWithCompensation(
          U.rows.map((G) => requireScalarValue(C(L[1], z(R, U, G)))),
        );
      }
      case "FILTER": {
        const U = D(C(L[0], R));
        return {
          ...U,
          rows: U.rows.filter((G) =>
            requireLogicalOperand(requireScalarValue(C(L[1], z(R, U, G)))),
          ),
        };
      }
      case "ALL": {
        const U = L[0];
        if (U.kind === "table") {
          const X = u(U.name).id;
          return { kind: "rows", tableId: X, rows: b(X) };
        }
        const G = l(U),
          H = new Set(),
          Q = b(G.tableId).filter((X) => {
            const J = buildTypedValueKey(y(G, X));
            return H.has(J) ? false : (H.add(J), true);
          });
        return {
          kind: "rows",
          tableId: G.tableId,
          rows: Q,
          columns: [G.column],
        };
      }
      case "RELATED": {
        const U = l(L[0]),
          G = R.rows.at(-1);
        if (!G || G.columns || G.tableId === U.tableId)
          return throwDaxError(
            "DAX_ROW",
            "RELATED requires a full row context on the related many side.",
          );
        const H = e.parentRow(G.tableId, U.tableId, G.index);
        return H === undefined
          ? throwDaxError(
              "DAX_RELATIONSHIP",
              "RELATED requires a unique many-to-one relationship path.",
            )
          : y(U, H);
      }
      case "CALCULATE": {
        const U = L.slice(1).map((H) => {
          if (H.kind === "call" && H.name === "ALL")
            return { remove: H.args[0] };
          if (
            H.kind === "table" ||
            (H.kind === "call" && H.name === "FILTER")
          ) {
            const q = D(C(H, R)),
              W = F(q.tableId).filter((K) => K.id !== q.tableId);
            return !q.columns &&
              W.some((K) =>
                W.some(
                  (te) =>
                    !e.canFilter(K.id, te.id) && !e.canFilter(te.id, K.id),
                ),
              )
              ? throwDaxError(
                  "DAX_FILTER",
                  "CALCULATE table filters spanning independent dimensions are outside this subset. Use explicit column predicates; FILTER remains available inside SUMX or COUNTROWS.",
                )
              : { filters: O(q) };
          }
          const Q = [];
          if (
            (walkDaxSyntaxTree(H, (q) => {
              (q.kind === "field" && Q.push(l(q)),
                (q.kind === "measure" ||
                  q.kind === "table" ||
                  (q.kind === "call" &&
                    ![
                      "BLANK",
                      "TRUE",
                      "FALSE",
                      "ISBLANK",
                      "IF",
                      "COALESCE",
                      "DIVIDE",
                    ].includes(q.name))) &&
                  throwDaxError(
                    "DAX_FILTER",
                    "Boolean CALCULATE filters support same-table columns, scalar functions and literals; use FILTER for a table expression.",
                  ));
            }),
            !Q.length || new Set(Q.map((q) => q.tableId)).size !== 1)
          )
            return throwDaxError(
              "DAX_FILTER",
              "A Boolean CALCULATE filter must reference columns from one table.",
            );
          const X = Q[0].tableId,
            J = b(X).filter((q) =>
              requireLogicalOperand(
                requireScalarValue(
                  C(H, {
                    ...R,
                    rows: [...R.rows, { tableId: X, index: q }],
                    transitioned: false,
                  }),
                ),
              ),
            );
          return { filters: [I(X, [...new Set(Q.map((q) => q.column))], J)] };
        });
        let G = [...j(R).filters];
        for (const H of U) H.remove && (G = V(G, H.remove));
        return (
          (G = k(
            G,
            U.flatMap((H) => H.filters ?? []),
          )),
          C(L[0], { ...R, filters: G, transitioned: true })
        );
      }
      case "TOTALYTD": {
        const U = l(L[1]),
          G = t.get(U.tableId);
        if (G.dateColumn !== U.column)
          return throwDaxError(
            "DAX_CALENDAR",
            "TOTALYTD requires a marked physical calendar date column.",
          );
        const H = b(G.id).map((K) => y(U, K));
        validateMarkedCalendarDays(H);
        const Q = j(R),
          X = E(G.id, Q)
            .map((K) => String(y(U, K)))
            .sort();
        if (!X.length) return null;
        const J = X.at(-1),
          q = J.slice(0, 4) + "-01-01",
          W = Q.filters.filter((K) => K.tableId !== G.id);
        return (
          W.push({
            tableId: G.id,
            columns: G.columns.map((K) => K.name),
            test: (K) => {
              const te = String(y(U, K));
              return te >= q && te <= J;
            },
          }),
          C(L[0], { ...Q, filters: W, transitioned: true })
        );
      }
    }
  }
  const T = (N, R = new Set()) => {
    if (R.has(N.id)) return new Set();
    R.add(N.id);
    const L = new Set([N.tableId]);
    return (
      N.operation === "dax" &&
        walkDaxSyntaxTree(n.get(`measure:${N.id}`), ($) => {
          if (
            ($.kind === "call" &&
              [
                "SUM",
                "AVERAGE",
                "MIN",
                "MAX",
                "DISTINCTCOUNT",
                "SELECTEDVALUE",
              ].includes($.name) &&
              L.add(l($.args[0]).tableId),
            $.kind === "call" && ["SUMX", "COUNTROWS"].includes($.name))
          ) {
            let U = $.args[0];
            for (; U.kind === "call" && ["FILTER", "ALL"].includes(U.name);)
              U = U.args[0];
            U.kind === "table"
              ? L.add(u(U.name).id)
              : U.kind === "field" && L.add(l(U).tableId);
          }
          $.kind === "measure" && T(c($.name), R).forEach((U) => L.add(U));
        }),
      L
    );
  };
  return {
    calculatedColumns: () =>
      r.tables.flatMap((N) =>
        N.columns
          .filter((R) => R.expression !== undefined)
          .map((R) => ({
            tableId: N.id,
            name: R.name,
            type: R.type,
            values: b(N.id).map((L) => y({ tableId: N.id, column: R.name }, L)),
          })),
      ),
    query(N, R) {
      const L = N.measureIds.map((X) => r.measures.find((J) => J.id === X)),
        $ = new Set(L.flatMap((X) => [...T(X)])),
        U = N.filters.map((X) => {
          const J = t.get(X.tableId).columns.find((q) => q.name === X.column);
          return (
            !["blank", "notBlank"].includes(X.operator) &&
              (X.value === undefined ||
                !isValidColumnValue(X.value, { ...J, nullable: true })) &&
              throwDaxError(
                "DAX_FILTER",
                "Filter value does not match its declared column type.",
              ),
            ["gt", "gte", "lt", "lte"].includes(X.operator) &&
              (!["number", "date"].includes(J.type) || X.value == null) &&
              throwDaxError(
                "DAX_FILTER",
                "Ordered query filters require a nonblank number or date.",
              ),
            L.length &&
              ![...$].some((q) => e.canFilter(X.tableId, q)) &&
              throwDaxError(
                "DAX_FILTER",
                "A query filter cannot reach any selected measure.",
              ),
            {
              tableId: X.tableId,
              columns: [X.column],
              test: (q) =>
                matchesQueryFilter(
                  y({ tableId: X.tableId, column: X.column }, q),
                  X.operator,
                  X.value,
                ),
            }
          );
        }),
        G = { filters: U, rows: [] };
      for (const X of R.tables)
        ((X.included = E(X.id, G).length), (X.excluded = X.rows - X.included));
      for (const X of N.groupBy)
        for (const J of $)
          e.canFilter(X.tableId, J) ||
            throwDaxError(
              "DAX_FILTER",
              `Group ${X.column} cannot filter every table referenced by the selected measures.`,
            );
      ((R.headers = [
        ...N.groupBy.map((X) => `${t.get(X.tableId).name}.${X.column}`),
        ...L.map((X) => X.name),
      ]),
        new Set(R.headers.map(caseFoldText)).size !== R.headers.length &&
          throwDaxError(
            "DAX_REFERENCE",
            "Group and measure headings must be distinct.",
          ));
      const H = (X) =>
        X === null || typeof X == "number"
          ? X
          : throwDaxError(
              "DAX_TYPE",
              "Published model measures must return numbers or BLANK. Use typed calculated columns for text or boolean results.",
            );
      if (((R.totals = L.map((X) => H(B(X, G)))), !N.groupBy.length)) {
        R.rows = [{ keys: [], values: [...R.totals] }];
        return;
      }
      const Q = new Map();
      for (const X of new Set(L.map((J) => J.tableId)))
        for (const J of E(X, G)) {
          e.tick();
          const q = N.groupBy.map((W) => {
            const K = e.parentRow(X, W.tableId, J);
            return K === undefined
              ? throwDaxError(
                  "DAX_RELATIONSHIP",
                  "A group has no relationship path.",
                )
              : y(W, K);
          });
          (Q.set(JSON.stringify(q), q),
            Q.size > dataModelLimits.groups &&
              throwDaxError(
                "DAX_LIMIT",
                "The model query exceeds 1,000 groups.",
              ));
        }
      R.rows = [...Q.values()].map((X) => {
        const J = {
          rows: [],
          filters: [
            ...U,
            ...N.groupBy.map((q, W) => ({
              tableId: q.tableId,
              columns: [q.column],
              test: (K) => y(q, K) === X[W],
            })),
          ],
        };
        return { keys: X, values: L.map((q) => H(B(q, J))) };
      });
    },
  };
}
function isValidColumnValue(e, r) {
  return e === null
    ? r.nullable
    : r.type === "number"
      ? typeof e == "number" && Number.isFinite(e)
      : r.type === "boolean"
        ? typeof e == "boolean"
        : r.type === "text"
          ? typeof e == "string" && e.length <= 32767
          : typeof e == "string" &&
            /^\d{4}-\d{2}-\d{2}$/.test(e) &&
            Number(e.slice(0, 4)) >= 1 &&
            Number.isFinite(new Date(e + "T00:00:00Z").getTime()) &&
            new Date(e + "T00:00:00Z").toISOString().slice(0, 10) === e;
}
function validateMarkedCalendarDays(e) {
  const r = e.map((o) => String(o)).sort(),
    t = r[0]?.slice(0, 4);
  (!r.length ||
    new Set(r).size !== r.length ||
    r[0] !== `${t}-01-01` ||
    r.at(-1) !== `${t}-12-31`) &&
    throwDaxError(
      "DAX_CALENDAR",
      "A marked calendar must contain every unique day of one complete calendar year, without blanks.",
    );
  for (let o = 1; o < r.length; o++)
    new Date(r[o] + "T00:00:00Z").getTime() -
      new Date(r[o - 1] + "T00:00:00Z").getTime() !==
      864e5 &&
      throwDaxError(
        "DAX_CALENDAR",
        "The marked calendar has a missing or invalid day.",
      );
}
function matchesQueryFilter(e, r, t) {
  switch (r) {
    case "blank":
      return e === null || e === "";
    case "notBlank":
      return e !== null && e !== "";
    case "equals":
      return e === t;
    case "notEquals":
      return e !== t;
    case "gt":
      return e !== null && e > t;
    case "gte":
      return e !== null && e >= t;
    case "lt":
      return e !== null && e < t;
    case "lte":
      return e !== null && e <= t;
    default:
      throw new DaxError("DAX_FILTER", "Unsupported query filter.");
  }
}
class DaxModelError extends Error {
  constructor(r, t, n = {}) {
    (super(t), (this.code = r), (this.context = n));
  }
}
const throwDaxModelError = (e, r, t = {}) => {
    throw new DaxModelError(e, r, t);
  },
  typedValueKey = (e) => JSON.stringify([typeof e, e]);
function isValidColumnValueForType(e, r) {
  if (e === null) return r.nullable;
  if (r.type === "number") return typeof e == "number" && Number.isFinite(e);
  if (r.type === "boolean") return typeof e == "boolean";
  if (r.type === "text") return typeof e == "string";
  if (
    typeof e != "string" ||
    !/^\d{4}-\d{2}-\d{2}$/.test(e) ||
    Number(e.slice(0, 4)) < 1
  )
    return false;
  const t = new Date(e + "T00:00:00Z");
  return !isNaN(t.getTime()) && t.toISOString().slice(0, 10) === e;
}
const decodeColumnValue = (e, r) =>
  e.nulls[r]
    ? null
    : e.dictionary
      ? e.dictionary[e.data[r]]
      : e.definition.type === "boolean"
        ? !!e.data[r]
        : e.data[r];
function encodeColumnDictionary(e, r) {
  const t = new Uint8Array(r.length),
    n = [],
    o = new Map(),
    a =
      e.type === "number"
        ? new Float64Array(r.length)
        : e.type === "boolean"
          ? new Uint8Array(r.length)
          : new Uint32Array(r.length);
  return (
    r.forEach((s, d) => {
      if (s === null) t[d] = 1;
      else if (e.type === "number") a[d] = s;
      else if (e.type === "boolean") a[d] = s ? 1 : 0;
      else {
        const u = String(s);
        (o.has(u) || (o.set(u, n.length), n.push(u)), (a[d] = o.get(u)));
      }
    }),
    {
      definition: e,
      nulls: t,
      data: a,
      ...(["text", "date"].includes(e.type) ? { dictionary: n } : {}),
    }
  );
}
function resolveModelTableRange(e, r) {
  const t = e.sheets.find((s) => s.id === r.sheetId);
  if (!t)
    return throwDaxModelError(
      "SOURCE",
      `Model table ${r.name} has a missing worksheet.`,
      {
        tableId: r.id,
      },
    );
  if (!r.tableId) return r.range;
  const n = t.tables?.find((s) => s.id === r.tableId);
  if (!n)
    return throwDaxModelError(
      "SOURCE",
      `The structured table for ${r.name} is missing.`,
      {
        tableId: r.id,
      },
    );
  const o = parseRangeAddress(n.range);
  return `${columnIndexToLetters(o.firstCol)}${o.firstRow + 1}:${columnIndexToLetters(o.lastCol)}${o.lastRow + 1 - (n.showTotals ? 1 : 0)}`;
}
function buildModelTableFromRange(e, r, t, n, o = {}) {
  const a = e.sheets.find((u) => u.id === r);
  if (!a) throw Error("Choose a source worksheet.");
  const s = {
      id: crypto.randomUUID(),
      name: a.name,
      sheetId: r,
      range: t ?? formatOccupiedRangeAddress(a),
      ...(n ? { tableId: n } : {}),
      columns: [],
    },
    d = readSourceHeaders(
      e,
      [{ sheetId: r, range: resolveModelTableRange(e, s) }],
      o,
    ).snapshots[0];
  return (
    (s.range = d.source.range),
    (s.columns = d.headers.map((u, l) => {
      const c = d.records.map((p) => p.values[l]),
        h = c.filter((p) => p !== null);
      if (new Set(h.map((p) => typeof p)).size > 1)
        throw Error(
          `Column ${u} contains mixed types. Clean the source before adding it to the model.`,
        );
      return {
        name: u,
        type:
          h.length && typeof h[0] == "number"
            ? "number"
            : h.length && typeof h[0] == "boolean"
              ? "boolean"
              : "text",
        nullable: c.some((p) => p === null),
      };
    })),
    s
  );
}
function normalizeModelHostPermissions(e, r) {
  if (r.readOnly !== undefined && typeof r.readOnly != "boolean")
    throw Error("Host read-only must be boolean.");
  return {
    readOnly: !!r.readOnly,
    includeHidden: !!r.includeHidden,
    allowedSheetIds: [
      ...new Set(r.allowedSheetIds ?? e.sheets.map((t) => t.id)),
    ].sort(),
  };
}
function buildModelTableSnapshots(e, r, t, n) {
  normalizeModelHostPermissions(e, t);
  for (const c of r.tables)
    e.sheets.find((h) => h.id === c.sheetId)?.dataModelReport &&
      throwDaxModelError(
        "SNAPSHOT_SOURCE",
        "Model report snapshots cannot be calculation sources. Use the original typed source tables; retained measure caches are not trusted inputs.",
      );
  const o = r.tables.map((c) => ({
    sheetId: c.sheetId,
    range: resolveModelTableRange(e, c),
  }));
  let a = 0;
  for (const c of o) {
    const h = parseRangeAddress(c.range),
      f = (h.lastRow - h.firstRow + 1) * (h.lastCol - h.firstCol + 1);
    ((a += f),
      (f > dataModelLimits.cellsPerTable ||
        h.lastRow - h.firstRow > dataModelLimits.rowsPerTable ||
        h.lastCol - h.firstCol + 1 > dataModelLimits.columns) &&
        throwDaxModelError(
          "LIMIT",
          "A model table supports at most 500 data rows, 50 fields and 2,000 cells including headers.",
        ));
  }
  a > dataModelLimits.modelCells &&
    throwDaxModelError("LIMIT", "The data model exceeds 16,000 source cells.");
  const s = readSourceHeaders(e, o, t),
    d = new Map();
  return (
    e.calculation?.mode === "manual" &&
      s.dependencyCells.some(
        (c) =>
          typeof c.cell?.raw == "string" &&
          c.cell.raw.startsWith("=") &&
          !c.cell.literal,
      ) &&
      throwDaxModelError(
        "MANUAL",
        "Model source formulas require Automatic calculation. In Manual mode use an explicitly reviewed values snapshot; pending or historical formula values cannot enter the model.",
      ),
    s.snapshots.forEach((c, h) => {
      const f = r.tables[h],
        p = f.columns.filter((g) => g.expression === undefined);
      (c.headers.length !== p.length ||
        c.headers.some((g, b) => g !== p[b].name)) &&
        throwDaxModelError(
          "SCHEMA",
          `The source schema of ${f.name} changed. Review the model columns.`,
          { tableId: f.id },
        );
      const m = new Map();
      let v = 0;
      (p.forEach((g, b) => {
        const y = c.records.map((A) => {
            const k = A.values[b];
            return (
              isValidColumnValueForType(k, g) ||
                throwDaxModelError(
                  "TYPE",
                  `${f.name}.${g.name} has an invalid ${g.type} value on worksheet row ${A.row}.`,
                  { tableId: f.id, column: g.name, row: A.row },
                ),
              k
            );
          }),
          w = encodeColumnDictionary(g, y);
        (m.set(g.name, w),
          (v +=
            w.data.byteLength +
            w.nulls.byteLength +
            (w.dictionary?.reduce((A, k) => A + k.length * 2, 0) ?? 0)));
      }),
        d.set(f.id, {
          definition: f,
          columns: m,
          count: c.records.length,
          range: c.source.range,
          sheetName: c.sheetName,
          bytes: v,
        }));
    }),
    (n.sourceHash = fingerprintCanonicalValue({
      sources: s.snapshots,
      dependencies: s.dependencyCells,
      datePolicy: e.datePolicy,
    })),
    d
  );
}
function evaluateDataModelQuery(e, r, t, n = {}) {
  const o = {
    complete: false,
    diagnostics: [],
    tables: [],
    rows: [],
    totals: [],
    headers: [],
    sourceHash: "",
    operations: 0,
    engine: columnarModelFormatVersion,
  };
  try {
    const s = validateDataModel(r),
      d = validateModelQuery(t, s),
      u = buildModelTableSnapshots(e, s, n, o),
      l = s.tables.map((A) => A.id),
      c = (A = 1) => {
        ((o.operations += A),
          o.operations > dataModelLimits.operations &&
            throwDaxModelError(
              "LIMIT",
              "Data model execution exceeded 2,000,000 operations. Reduce the model or query.",
            ));
      },
      h = new Map(
        l.map((A) => [A, s.relationships.filter((k) => k.one.tableId === A)]),
      ),
      f = new Map(l.map((A) => [A, 0]));
    for (const A of s.relationships)
      f.set(A.many.tableId, f.get(A.many.tableId) + 1);
    const p = l.filter((A) => !f.get(A));
    for (let A = 0; A < p.length; A++)
      for (const k of h.get(p[A])) {
        const F = f.get(k.many.tableId) - 1;
        (f.set(k.many.tableId, F), F || p.push(k.many.tableId));
      }
    p.length !== l.length &&
      throwDaxModelError(
        "CYCLE",
        "Relationships form a cycle. Use a single-direction acyclic model.",
      );
    const m = new Map();
    for (const A of l) {
      const k = new Map([[A, []]]);
      for (const F of p) {
        const I = k.get(F);
        if (I)
          for (const O of h.get(F))
            (k.has(O.many.tableId) &&
              throwDaxModelError(
                "AMBIGUOUS",
                `More than one filter path connects ${u.get(A).definition.name} to ${u.get(O.many.tableId).definition.name}. Remove the ambiguous relationship.`,
              ),
              k.set(O.many.tableId, [...I, O.id]));
      }
      m.set(A, k);
    }
    const v = new Map();
    for (const A of s.relationships) {
      const k = u.get(A.one.tableId),
        F = u.get(A.many.tableId),
        I = new Map(),
        O = new Set(),
        j = new Int32Array(F.count).fill(-1);
      for (let V = 0; V < k.count; V++) {
        c();
        const D = decodeColumnValue(k.columns.get(A.one.column), V),
          B = typedValueKey(D);
        ((D === null || D === "") &&
          throwDaxModelError(
            "BLANK_KEY",
            `The one side of ${A.id} contains a blank key.`,
            {
              tableId: A.one.tableId,
              column: A.one.column,
              row: parseRangeAddress(k.range).firstRow + V + 2,
              relationshipId: A.id,
            },
          ),
          I.has(B) &&
            throwDaxModelError(
              "DUPLICATE_KEY",
              `The one side of ${A.id} contains a duplicate key.`,
              {
                tableId: A.one.tableId,
                column: A.one.column,
                row: parseRangeAddress(k.range).firstRow + V + 2,
                relationshipId: A.id,
              },
            ),
          I.set(B, V));
      }
      let z = 0;
      for (let V = 0; V < F.count; V++) {
        c();
        const D = decodeColumnValue(F.columns.get(A.many.column), V),
          B = typedValueKey(D),
          _ = I.get(B);
        (A.cardinality === "one-to-one" &&
          O.has(B) &&
          throwDaxModelError(
            "CARDINALITY",
            `The related side of ${A.id} is not unique. Choose one-to-many or correct its keys.`,
            { relationshipId: A.id },
          ),
          O.add(B),
          _ === undefined
            ? (z++,
              o.diagnostics.length < 100 &&
                o.diagnostics.push({
                  code: "MISSING_KEY",
                  message: `${F.definition.name}.${A.many.column} has no dimension key; correct the source before calculating.`,
                  tableId: A.many.tableId,
                  column: A.many.column,
                  row: parseRangeAddress(F.range).firstRow + V + 2,
                  relationshipId: A.id,
                }))
            : (j[V] = _));
      }
      (z > 100 &&
        o.diagnostics.push({
          code: "MISSING_KEY_LIMIT",
          message: `${A.id}: ${z} missing dimension keys; only the first 100 diagnostics are shown.`,
          relationshipId: A.id,
        }),
        v.set(A.id, j));
    }
    const g = new Map(
      [...u].map(([A, k]) => [A, new Uint8Array(k.count).fill(1)]),
    );
    if (
      ((o.tables = [...u.values()].map((A) => ({
        id: A.definition.id,
        name: A.definition.name,
        sheetId: A.definition.sheetId,
        range: A.range,
        rows: A.count,
        included: A.count,
        excluded: 0,
        columnarBytes: A.bytes,
      }))),
      o.diagnostics.length)
    )
      return o;
    if (
      s.measures.some((A) => A.operation === "dax") ||
      s.tables.some((A) => A.columns.some((k) => k.expression !== undefined))
    ) {
      const A = evaluateDaxExpression({
        definition: s,
        count: (k) => u.get(k).count,
        value: (k, F, I) => decodeColumnValue(u.get(k).columns.get(F), I),
        order: p,
        tick: c,
        canFilter: (k, F) => m.get(k)?.has(F) ?? false,
        parentRow: (k, F, I) => {
          const O = m.get(F)?.get(k);
          if (!O) return;
          let j = I;
          for (let V = O.length - 1; V >= 0; V--) (c(), (j = v.get(O[V])[j]));
          return j;
        },
      });
      o.calculatedColumns = A.calculatedColumns();
      for (const k of o.calculatedColumns) {
        const F = u.get(k.tableId),
          I = encodeColumnDictionary(
            F.definition.columns.find((O) => O.name === k.name),
            k.values,
          );
        (F.columns.set(k.name, I),
          (o.tables.find((O) => O.id === k.tableId).columnarBytes +=
            I.data.byteLength +
            I.nulls.byteLength +
            (I.dictionary?.reduce((O, j) => O + j.length * 2, 0) ?? 0)));
      }
      return (
        A.query(d, o),
        (o.rows.length + 2) * o.headers.length > dataModelLimits.outputCells &&
          throwDaxModelError(
            "LIMIT",
            "The model result exceeds 20,000 output cells.",
          ),
        (o.complete = true),
        o
      );
    }
    const b = d.measureIds.map((A) => s.measures.find((k) => k.id === A));
    for (const A of d.filters) {
      const k = u.get(A.tableId),
        F = k.columns.get(A.column),
        I = g.get(A.tableId);
      (!["blank", "notBlank"].includes(A.operator) &&
        (A.value === undefined ||
          !isValidColumnValueForType(A.value, {
            ...F.definition,
            nullable: true,
          })) &&
        throwDaxModelError(
          "FILTER",
          "Filter values must match the declared column type.",
        ),
        ["gt", "gte", "lt", "lte"].includes(A.operator) &&
          (!["number", "date"].includes(F.definition.type) ||
            A.value === null) &&
          throwDaxModelError(
            "FILTER",
            "Ordered filters require a nonblank number or ISO date.",
          ),
        b.length &&
          !b.some((O) => m.get(A.tableId).has(O.tableId)) &&
          throwDaxModelError(
            "DISCONNECTED",
            "A filter cannot reach any selected measure. Connect its table or remove the filter.",
          ));
      for (let O = 0; O < k.count; O++) {
        c();
        const j = decodeColumnValue(F, O);
        matchesComparisonOperator(j, A) || (I[O] = 0);
      }
    }
    for (const A of p)
      for (const k of h.get(A)) {
        const F = g.get(A),
          I = g.get(k.many.tableId),
          O = v.get(k.id);
        for (let j = 0; j < I.length; j++) (c(), F[O[j]] || (I[j] = 0));
      }
    for (const A of o.tables)
      ((A.included = g.get(A.id).reduce((k, F) => k + F, 0)),
        (A.excluded = A.rows - A.included));
    for (const A of d.groupBy)
      for (const k of b)
        m.get(A.tableId).has(k.tableId) ||
          throwDaxModelError(
            "DISCONNECTED",
            `Group field ${A.column} cannot filter measure ${k.name}. Use dimensions shared by every selected fact table.`,
          );
    ((o.headers = [
      ...d.groupBy.map(
        (A) => `${u.get(A.tableId).definition.name}.${A.column}`,
      ),
      ...b.map((A) => A.name),
    ]),
      new Set(o.headers.map((A) => A.toLowerCase())).size !==
        o.headers.length &&
        throwDaxModelError(
          "SCHEMA",
          "Group and measure headings must be distinct.",
        ));
    const y = (A, k) =>
        d.groupBy.map((F) => {
          let I = k;
          const O = m.get(F.tableId).get(A);
          for (let j = O.length - 1; j >= 0; j--) (c(), (I = v.get(O[j])[I]));
          return decodeColumnValue(u.get(F.tableId).columns.get(F.column), I);
        }),
      w = new Map(),
      E = b.map(() => createSumAccumulator());
    for (const [A, k] of b.entries()) {
      const F = u.get(k.tableId),
        I = g.get(k.tableId),
        O = k.column ? F.columns.get(k.column) : undefined;
      for (let j = 0; j < F.count; j++) {
        if ((c(), !I[j])) continue;
        const z = k.operation === "countRows" ? 1 : decodeColumnValue(O, j);
        if ((addAggregateValue(E[A], z), !d.groupBy.length)) continue;
        const V = y(k.tableId, j),
          D = JSON.stringify(V);
        (w.has(D) ||
          (w.size >= dataModelLimits.groups &&
            throwDaxModelError(
              "LIMIT",
              "The model query exceeds 1,000 groups.",
            ),
          w.set(D, { keys: V, states: b.map(() => createSumAccumulator()) })),
          addAggregateValue(w.get(D).states[A], z));
      }
    }
    return (
      (o.totals = E.map((A, k) => evaluateAggregation(A, b[k]))),
      (o.rows = d.groupBy.length
        ? [...w.values()].map((A) => ({
            keys: A.keys,
            values: A.states.map((k, F) => evaluateAggregation(k, b[F])),
          }))
        : [{ keys: [], values: [...o.totals] }]),
      (o.rows.length + 2) * o.headers.length > dataModelLimits.outputCells &&
        throwDaxModelError(
          "LIMIT",
          "The model result exceeds 20,000 output cells.",
        ),
      (o.complete = true),
      o
    );
  } catch (s) {
    return (
      (o.rows = []),
      (o.totals = []),
      o.diagnostics.push({
        code:
          s instanceof DaxModelError || s instanceof DaxError
            ? s.code
            : "MODEL",
        message: s instanceof Error ? s.message : "Model evaluation failed.",
        ...(s instanceof DaxModelError ? s.context : {}),
      }),
      o
    );
  }
}
function matchesComparisonOperator(e, r) {
  const t = r.value;
  switch (r.operator) {
    case "blank":
      return e === null || e === "";
    case "notBlank":
      return e !== null && e !== "";
    case "equals":
      return e === t;
    case "notEquals":
      return e !== t;
    case "gt":
      return e !== null && e > t;
    case "gte":
      return e !== null && e >= t;
    case "lt":
      return e !== null && e < t;
    case "lte":
      return e !== null && e <= t;
  }
}
function createSumAccumulator() {
  return {
    count: 0,
    sum: 0,
    correction: 0,
    min: 1 / 0,
    max: -1 / 0,
    distinct: new Set(),
  };
}
function addAggregateValue(e, r) {
  if (
    r !== null &&
    (e.count++, e.distinct.add(typedValueKey(r)), typeof r == "number")
  ) {
    const t = e.sum + r;
    ((e.correction +=
      Math.abs(e.sum) >= Math.abs(r) ? e.sum - t + r : r - t + e.sum),
      (e.sum = t),
      (e.min = Math.min(e.min, r)),
      (e.max = Math.max(e.max, r)));
  }
}
function evaluateAggregation(e, r) {
  const t = e.sum + e.correction;
  let n;
  switch (r.operation) {
    case "count":
    case "countRows":
      n = e.count;
      break;
    case "distinctCount":
      n = e.distinct.size;
      break;
    case "sum":
      n = e.count ? t : null;
      break;
    case "average":
      n = e.count ? t / e.count : null;
      break;
    case "min":
      n = e.count ? e.min : null;
      break;
    case "max":
      n = e.count ? e.max : null;
      break;
    case "dax":
      return throwDaxModelError(
        "DAX_ENGINE",
        "DAX measures require the separate model expression evaluator.",
      );
  }
  return (
    n !== null &&
      !Number.isFinite(n) &&
      throwDaxModelError(
        "OVERFLOW",
        `Measure ${r.name} exceeds finite numeric precision.`,
      ),
    n
  );
}
const dataModelResultCache = new WeakMap();
function previewDataModelOutput(e, r, t = {}) {
  const n = validateDataModel(r.definition),
    o = validateModelQuery(r.query, n),
    a = r.outputName.trim();
  if (!a || a.length > 31 || /[\\/?*\[\]:\u0000-\u001f]/.test(a))
    throw Error(
      "Enter a valid output worksheet name of at most 31 characters.",
    );
  const s = {
    request: { definition: n, query: o, outputName: a },
    workbookHash: fingerprintCanonicalValue(e),
    contextHash: fingerprintCanonicalValue(normalizeModelHostPermissions(e, t)),
    result: evaluateDataModelQuery(e, n, o, t),
  };
  return (dataModelResultCache.set(s, fingerprintCanonicalValue(s)), s);
}
function buildModelReportCells(e, r) {
  const t = parseRangeAddress(r),
    n = [];
  if (
    (t.lastRow - t.firstRow + 1) * (t.lastCol - t.firstCol + 1) >
    dataModelLimits.outputCells
  )
    throw Error("Model report range exceeds 20,000 cells.");
  for (let o = t.firstRow; o <= t.lastRow; o++) {
    const a = [];
    for (let s = t.firstCol; s <= t.lastCol; s++) {
      const d = e.cells[`${columnIndexToLetters(s)}${o + 1}`];
      a.push({
        raw: d?.raw ?? null,
        literal: !!d?.literal,
        note: d?.note ?? "",
      });
    }
    n.push(a);
  }
  return fingerprintCanonicalValue(n);
}
function publishDataModelReport(e, r, t, n = {}) {
  if (n.readOnly) throw Error("This workbook is read-only.");
  if ((assertWorkbookStructureEditable(e), !["save", "publish"].includes(t)))
    throw Error("Choose Save model or Publish result.");
  if (
    dataModelResultCache.get(r) !== fingerprintCanonicalValue(r) ||
    r.workbookHash !== fingerprintCanonicalValue(e) ||
    r.contextHash !==
      fingerprintCanonicalValue(normalizeModelHostPermissions(e, n))
  )
    throw Error(
      "The model, workbook or host permissions changed. Preview again.",
    );
  const o = evaluateDataModelQuery(e, r.request.definition, r.request.query, n);
  if (
    !o.complete ||
    fingerprintCanonicalValue(o) !== fingerprintCanonicalValue(r.result)
  )
    throw Error(
      "Resolve model diagnostics and preview again. No partial model result is published.",
    );
  const a = normalizeWorkbookInput(e);
  if (
    ((a.dataModel = structuredClone(r.request.definition)), t === "publish")
  ) {
    if (!r.request.query.measureIds.length)
      throw Error("Select at least one measure to publish.");
    if (
      a.sheets.some(
        (h) => h.name.toLowerCase() === r.request.outputName.toLowerCase(),
      )
    )
      throw Error("The output worksheet name is already in use.");
    inspectOutputNameSafety(e, r.request.outputName);
    const d = createWorksheetRecord(r.request.outputName),
      u = r.request.query.groupBy.length > 0,
      l = [o.headers, ...o.rows.map((h) => [...h.keys, ...h.values])];
    (u &&
      l.push([
        "Grand total",
        ...Array(r.request.query.groupBy.length - 1).fill(null),
        ...o.totals,
      ]),
      l.forEach((h, f) =>
        h.forEach((p, m) => {
          d.cells[`${columnIndexToLetters(m)}${f + 1}`] = {
            raw: p,
            ...(typeof p == "string" && p.startsWith("=")
              ? { literal: true }
              : {}),
            ...(f === 0
              ? { style: { bold: true, fill: "#362C63", color: "#FFFFFF" } }
              : {}),
            ...(f > 0
              ? {
                  note: `Reviewed model snapshot: ${r.request.definition.name}. Source hash ${o.sourceHash}. Reopen Data Model to inspect source freshness and refresh.`,
                }
              : {}),
          };
        }),
      ),
      (d.columnWidths = Object.fromEntries(
        o.headers.map((h, f) => [String(f), 160]),
      )),
      (d.frozenRows = 1));
    const c = `A1:${columnIndexToLetters(o.headers.length - 1)}${l.length}`;
    ((d.dataModelReport = {
      definition: structuredClone(r.request.definition),
      query: structuredClone(r.request.query),
      sourceHash: o.sourceHash,
      definitionHash: fingerprintCanonicalValue(r.request.definition),
      outputHash: buildModelReportCells(d, c),
      outputRange: c,
      createdAt: new Date().toISOString(),
      sourceRevision: e.revision,
      includeHidden: !!n.includeHidden,
      totals: [...o.totals],
    }),
      a.sheets.push(d));
  }
  const s = normalizeWorkbookInput(a);
  return (dataModelResultCache.delete(r), s);
}
function removeDataModelFromWorkbook(e, r = {}) {
  if (r.readOnly) throw Error("This workbook is read-only.");
  assertWorkbookStructureEditable(e);
  const t = normalizeWorkbookInput(e);
  return (delete t.dataModel, normalizeWorkbookInput(t));
}
function getModelReportStatus(e, r, t = {}) {
  const n = e.sheets.find((a) => a.id === r),
    o = n?.dataModelReport;
  if (!n || !o) return { status: "none", message: "No model report." };
  try {
    if (
      !e.dataModel ||
      fingerprintCanonicalValue(e.dataModel) !== o.definitionHash
    )
      return {
        status: "stale",
        message:
          "Model definition changed. Preview and publish a fresh result.",
      };
    const a = evaluateDataModelQuery(e, o.definition, o.query, {
      ...t,
      includeHidden: o.includeHidden && !!t.includeHidden,
    });
    return a.complete
      ? a.sourceHash !== o.sourceHash ||
        buildModelReportCells(n, o.outputRange) !== o.outputHash ||
        fingerprintCanonicalValue(a.totals) !==
          fingerprintCanonicalValue(o.totals)
        ? {
            status: "stale",
            message:
              "Model sources or output changed. Retained values are an earlier snapshot; preview and publish again.",
          }
        : {
            status: "current",
            message:
              "Model snapshot matches current sources and definition. Local fingerprints are not signed source evidence.",
          }
      : {
          status: "unavailable",
          message:
            "Model sources cannot be verified. Review relationships, source access and diagnostics.",
        };
  } catch {
    return {
      status: "unavailable",
      message:
        "Model source or output verification failed. Retained values are unverified.",
    };
  }
}
function DaxExpressionInput({ label, value, onChange }) {
  const { text } = useWorkbookMessages(),
    o = ne.useId(),
    a = ne.useRef(null);
  return i.jsxs("div", {
    className: "lg-dax-expression",
    children: [
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: label }),
          i.jsx("textarea", {
            ref: a,
            "aria-label": text(label),
            "aria-describedby": o,
            value,
            maxLength: daxExpressionLimits.expressionCharacters,
            rows: 3,
            spellCheck: false,
            onChange: (d) => onChange(d.target.value),
          }),
        ],
      }),
      i.jsx("div", {
        className: "lg-model-controls",
        children: i.jsxs("label", {
          children: [
            i.jsx(LocalizedText, { children: "Insert DAX function" }),
            i.jsxs("select", {
              "aria-label": text(`${label} insert function`),
              value: "",
              onChange: (d) => {
                if (!d.target.value) return;
                const u = a.current?.selectionStart ?? value.length,
                  l = a.current?.selectionEnd ?? u,
                  c = d.target.value + "(";
                (onChange(value.slice(0, u) + c + value.slice(l)),
                  requestAnimationFrame(() => {
                    (a.current?.focus(),
                      a.current?.setSelectionRange(u + c.length, u + c.length));
                  }));
              },
              children: [
                i.jsx("option", {
                  value: "",
                  children: text("Choose a function"),
                }),
                Object.keys(daxFunctionNames).map((d) =>
                  i.jsx("option", { value: d, children: d }, d),
                ),
              ],
            }),
          ],
        }),
      }),
      i.jsxs("p", {
        id: o,
        className: "lg-model-help",
        children: [
          i.jsx(LocalizedText, {
            children:
              "Use table[column] and [Measure name]. This is DAX, evaluated separately from worksheet formulas.",
          }),
          i.jsx("br", {}),
          `SUM('Sales'[Amount]) · [Gross sales] - [Net sales] · CALCULATE([Net sales], Customers[Region] = "North")`,
        ],
      }),
    ],
  });
}
function DataModelDefinitionEditor({ definition, onChange }) {
  const { text } = useWorkbookMessages(),
    n = definition.tables.reduce(
      (a, s) => a + s.columns.filter((d) => d.expression !== undefined).length,
      0,
    ),
    o = (a, s) =>
      onChange({
        ...definition,
        tables: definition.tables.map((d) => (d.id === a ? { ...d, ...s } : d)),
      });
  return i.jsxs(i.Fragment, {
    children: [
      i.jsx("p", {
        className: "lg-model-help",
        children: i.jsx(LocalizedText, {
          children:
            "Calculated columns are evaluated for every source row before query filters. Sources remain unchanged. Preview is required before saving or publishing.",
        }),
      }),
      definition.tables.map((a) =>
        i.jsxs(
          "fieldset",
          {
            children: [
              i.jsx("legend", { children: a.name }),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Calendar date column" }),
                  i.jsxs("select", {
                    "aria-label": text(`${a.name} calendar date column`),
                    value: a.dateColumn ?? "",
                    onChange: (s) =>
                      o(a.id, { dateColumn: s.target.value || undefined }),
                    children: [
                      i.jsx("option", {
                        value: "",
                        children: text("Not a marked calendar"),
                      }),
                      a.columns
                        .filter(
                          (s) =>
                            s.expression === undefined &&
                            s.type === "date" &&
                            !s.nullable,
                        )
                        .map((s) =>
                          i.jsx(
                            "option",
                            { value: s.name, children: s.name },
                            s.name,
                          ),
                        ),
                    ],
                  }),
                ],
              }),
              i.jsx("p", {
                className: "lg-model-help",
                children: i.jsx(LocalizedText, {
                  children:
                    "TOTALYTD requires a marked table containing every day of one complete calendar year, without duplicate or blank dates.",
                }),
              }),
              a.columns.map((s, d) => {
                const u = a.columns
                  .slice(0, d + 1)
                  .filter((l) => l.expression !== undefined).length;
                return s.expression === undefined
                  ? null
                  : i.jsxs(
                      "fieldset",
                      {
                        children: [
                          i.jsxs("legend", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Calculated column",
                              }),
                              " ",
                              u,
                            ],
                          }),
                          i.jsxs("div", {
                            className: "lg-model-controls",
                            children: [
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, { children: "Name" }),
                                  i.jsx("input", {
                                    "aria-label": text(
                                      `${a.name} calculated column ${u} name`,
                                    ),
                                    value: s.name,
                                    onChange: (l) =>
                                      o(a.id, {
                                        columns: a.columns.map((c, h) =>
                                          h === d
                                            ? { ...c, name: l.target.value }
                                            : c,
                                        ),
                                      }),
                                  }),
                                ],
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, { children: "Type" }),
                                  i.jsx("select", {
                                    "aria-label": text(
                                      `${a.name} calculated column ${u} type`,
                                    ),
                                    value: s.type,
                                    onChange: (l) =>
                                      o(a.id, {
                                        columns: a.columns.map((c, h) =>
                                          h === d
                                            ? { ...c, type: l.target.value }
                                            : c,
                                        ),
                                      }),
                                    children: [
                                      "number",
                                      "text",
                                      "boolean",
                                      "date",
                                    ].map((l) =>
                                      i.jsx(
                                        "option",
                                        { value: l, children: text(l) },
                                        l,
                                      ),
                                    ),
                                  }),
                                ],
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx("input", {
                                    type: "checkbox",
                                    "aria-label": text(
                                      `${a.name} calculated column ${u} nullable`,
                                    ),
                                    checked: s.nullable,
                                    onChange: (l) =>
                                      o(a.id, {
                                        columns: a.columns.map((c, h) =>
                                          h === d
                                            ? {
                                                ...c,
                                                nullable: l.target.checked,
                                              }
                                            : c,
                                        ),
                                      }),
                                  }),
                                  i.jsx(LocalizedText, {
                                    children: "Allow BLANK",
                                  }),
                                ],
                              }),
                              i.jsx("button", {
                                "aria-label": text(
                                  `Remove ${a.name} calculated column ${u}`,
                                ),
                                onClick: () =>
                                  o(a.id, {
                                    columns: a.columns.filter(
                                      (l, c) => c !== d,
                                    ),
                                  }),
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            ],
                          }),
                          i.jsx(DaxExpressionInput, {
                            label: `${a.name} calculated column ${u} expression`,
                            value: s.expression,
                            onChange: (l) =>
                              o(a.id, {
                                columns: a.columns.map((c, h) =>
                                  h === d ? { ...c, expression: l } : c,
                                ),
                              }),
                          }),
                        ],
                      },
                      d,
                    );
              }),
              i.jsx("button", {
                "aria-label": text(`Add calculated column to ${a.name}`),
                disabled:
                  n >= daxExpressionLimits.calculatedColumns ||
                  a.columns.length >= dataModelLimits.columns,
                onClick: () => {
                  let s = 1;
                  for (
                    ;
                    a.columns.some(
                      (d) => d.name.toLowerCase() === `calculated ${s}`,
                    );
                  )
                    s++;
                  o(a.id, {
                    columns: [
                      ...a.columns,
                      {
                        name: `Calculated ${s}`,
                        type: "number",
                        nullable: true,
                        expression: "BLANK()",
                      },
                    ],
                  });
                },
                children: i.jsx(LocalizedText, {
                  children: "Add calculated column",
                }),
              }),
            ],
          },
          a.id,
        ),
      ),
      !definition.tables.length &&
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children: "Add a source in Tables first.",
          }),
        }),
    ],
  });
}
const LabeledField = (e, r) =>
  i.jsxs("label", { children: [i.jsx(LocalizedText, { children: e }), r] });

export {
  previewDataModelOutput,
  DataModelDefinitionEditor,
  publishDataModelReport,
  getModelReportStatus,
  removeDataModelFromWorkbook,
  LabeledField,
  buildModelTableFromRange,
  DaxExpressionInput,
};
