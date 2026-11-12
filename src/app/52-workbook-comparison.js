import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { Mt } from "./02-financial-functions.js";
import {
  Af,
  Nr,
  Ta,
  Vn,
  _t,
  cn,
  ct,
  ff,
  hi,
  ul,
  ut,
  z1,
} from "./03-workbook-model.js";
import { Dr, M } from "./04-cell-input-paste-and-fill.js";
import { H1 } from "./07-xlsx-archive-reader.js";
import { yh } from "./09-workbook-links.js";
import {
  $h,
  Bj,
  Fj,
  Vf,
  a2e,
  i2e,
  n2e,
  nH,
  r2e,
  rH,
  ww,
  xc,
  yp,
} from "./51-pattern-checks-and-citations.js";

function s2e(e, r, t, n = {}) {
  var g;
  if (
    !t ||
    !Array.isArray(t.mappings) ||
    !t.mappings.length ||
    t.mappings.length > $h.mappings
  )
    throw Error("Choose 1–20 explicit worksheet/range mappings.");
  if (n.includeHidden !== undefined && typeof n.includeHidden != "boolean")
    throw Error("Include hidden must be boolean.");
  const o = (b, y) => {
      if (
        y !== undefined &&
        (!Array.isArray(y) ||
          y.some(
            (w) => typeof w != "string" || !b.sheets.some((E) => E.id === w),
          ))
      )
        throw Error("The comparison worksheet access context is invalid.");
      return [...new Set(y ?? b.sheets.map((w) => w.id))].sort();
    },
    a = {
      before: o(e, n.beforeAllowedSheetIds),
      after: o(r, n.afterAllowedSheetIds),
    },
    s = { before: e, after: r },
    d = {
      before: { id: e.id, name: e.name, revision: e.revision },
      after: { id: r.id, name: r.name, revision: r.revision },
      request: structuredClone(t),
      context: {
        beforeAllowedSheetIds: a.before,
        afterAllowedSheetIds: a.after,
        includeHidden: !!n.includeHidden,
      },
      findings: [],
      counts: {
        beforeRows: 0,
        afterRows: 0,
        matchedRows: 0,
        insertedRows: 0,
        deletedRows: 0,
        unresolvedBeforeRows: 0,
        unresolvedAfterRows: 0,
        scannedCells: 0,
        scannedFormulas: 0,
      },
      coverage: {
        selectedCells: 0,
        notes: [
          "Static formula text and stored cell metadata are compared; calculated values and imported caches are never used.",
          "Row keys use exact typed literal values; case and whitespace are significant. Duplicate, blank or formula keys remain unresolved.",
          "Literal input inventory is observed; formula-column hardcodes and pattern outliers are review candidates, not proof of error. Pattern analysis uses relative A1 signatures within each selected column; a dominant pattern requires at least two formulas and more than half the formulas.",
          "Only explicitly mapped rectangles receive row/cell comparison. Static links and hidden structure are inspected across permitted worksheets; unmapped sheets are listed.",
        ],
        formulaValuesCalculated: false,
        beforeUnmappedSheets: 0,
        afterUnmappedSheets: 0,
      },
    },
    u = (b) => {
      if (d.findings.length >= $h.findings)
        throw Error(
          "Comparison exceeds 30,000 findings; narrow the source workbooks.",
        );
      d.findings.push(b);
    },
    l = (b, y, w) => ({
      side: b,
      workbookId: s[b].id,
      revision: s[b].revision,
      sheetId: y.id,
      sheetName: y.name,
      address: w,
    }),
    c = (b, y, w) => ({ [b]: l(b, y, w) }),
    h = (b) => ({ rows: cn(b, "row"), cols: cn(b, "column") }),
    f = (b, y, w) => {
      if (!y || typeof y.sheetId != "string" || typeof y.range != "string")
        throw Error(
          "Choose explicit worksheet and rectangular range references.",
        );
      const E = s[b].sheets.find((j) => j.id === y.sheetId);
      if (!E || !a[b].includes(E.id))
        throw Error(
          "A selected source is outside the authorized comparison context.",
        );
      const A = ct(y.range),
        k = A.lastRow - A.firstRow + 1,
        F = A.lastCol - A.firstCol + 1;
      if (
        k > $h.rowsPerRange ||
        F > $h.columnsPerRange ||
        A.lastRow >= E.rowCount ||
        A.lastCol >= E.colCount
      )
        throw Error(
          "Each comparison range supports at most 1,000 rows and 100 columns within the worksheet.",
        );
      if (
        ((d.coverage.selectedCells += k * F),
        d.coverage.selectedCells > $h.selectedCells)
      )
        throw Error(
          "Comparison supports at most 10,000 selected cells across both workbooks.",
        );
      const I = h(E);
      if (
        !n.includeHidden &&
        (E.hidden ||
          E.veryHidden ||
          I.rows.some((j) => j.start <= A.lastRow && j.end >= A.firstRow) ||
          I.cols.some((j) => j.start <= A.lastCol && j.end >= A.firstCol))
      )
        throw Error(
          "A selected range contains hidden content. Explicitly include permitted hidden content or choose a visible range.",
        );
      const O = Array.from({ length: F }, (j, z) => {
        if (!w) return ut(z);
        const V = E.cells[xc(A.firstRow, A.firstCol + z)];
        if (Vf(V) || typeof V?.raw != "string" || !V.raw.trim())
          throw Error(
            "Header rows require nonempty literal text in every column. Turn off headers for an unstructured sheet.",
          );
        return V.raw;
      });
      if (new Set(O.map((j) => j.toLowerCase())).size !== O.length)
        throw Error(
          "Duplicate headers are ambiguous; resolve them or compare without headers.",
        );
      return {
        sheet: E,
        range: `${xc(A.firstRow, A.firstCol)}:${xc(A.lastRow, A.lastCol)}`,
        headers: O,
        firstCol: A.firstCol,
        rows: Array.from({ length: k - Number(w) }, (j, z) => {
          const V = A.firstRow + Number(w) + z;
          return {
            row: V,
            cells: O.map((D, B) => E.cells[xc(V, A.firstCol + B)]),
          };
        }),
      };
    },
    p = { before: new Set(), after: new Set() },
    m = { before: new Map(), after: new Map() };
  for (const b of t.mappings) {
    if (
      !b ||
      typeof b.headers != "boolean" ||
      !["keys", "position"].includes(b.rowMatch)
    )
      throw Error(
        "Each mapping requires explicit header and row-matching rules.",
      );
    for (const I of ["before", "after"]) {
      const O = b[I],
        j = ct(O.range),
        z = m[I].get(O.sheetId) ?? [];
      if (
        z.some(
          (V) =>
            V.firstRow <= j.lastRow &&
            V.lastRow >= j.firstRow &&
            V.firstCol <= j.lastCol &&
            V.lastCol >= j.firstCol,
        )
      )
        throw Error("Comparison mappings may not overlap on either side.");
      (m[I].set(O.sheetId, [...z, j]), p[I].add(O.sheetId));
    }
    const y = f("before", b.before, b.headers),
      w = f("after", b.after, b.headers);
    ((d.counts.beforeRows += y.rows.length),
      (d.counts.afterRows += w.rows.length));
    const E = (I, O, j) =>
        c(
          I,
          O.sheet,
          `${xc(j, O.firstCol)}:${xc(j, O.firstCol + O.headers.length - 1)}`,
        ),
      A = [],
      k = (I, O, j) => {
        (d.counts[I === "before" ? "deletedRows" : "insertedRows"]++,
          u({
            kind: I === "before" ? "deleted-row" : "inserted-row",
            confidence: "observed",
            message: `${I === "before" ? "Deleted" : "Inserted"} record under the declared ${b.rowMatch} match rule.`,
            ...E(I, O, O.rows[j].row),
          }));
      },
      F = (I, O, j, z) => {
        (d.counts[
          I === "before" ? "unresolvedBeforeRows" : "unresolvedAfterRows"
        ]++,
          u({
            kind: "uncertain",
            confidence: "unresolved",
            message: z,
            ...E(I, O, O.rows[j].row),
          }));
      };
    if (b.rowMatch === "position") {
      if ((g = b.keys) != null && g.length)
        throw Error("Positional matching does not use keys.");
      u({
        kind: "uncertain",
        confidence: "unresolved",
        message:
          "Positional alignment was explicitly chosen. Reordering cannot be distinguished from changed records; use unique literal keys for record identity.",
        ...c("before", y.sheet, y.range),
        ...c("after", w.sheet, w.range),
      });
      for (let I = 0; I < Math.max(y.rows.length, w.rows.length); I++)
        I >= y.rows.length
          ? k("after", w, I)
          : I >= w.rows.length
            ? k("before", y, I)
            : A.push([I, I]);
    } else {
      if (!Array.isArray(b.keys) || !b.keys.length || b.keys.length > $h.keys)
        throw Error(
          "Key matching requires 1–8 explicit header pairs (A, B, … without headers).",
        );
      const I = (V, D) => {
          const B = b.keys.map((_) => V.headers.indexOf(_[D]));
          if (B.some((_) => _ < 0) || new Set(B).size !== B.length)
            throw Error(
              "Key headers must exist exactly once and may not be repeated.",
            );
          return B;
        },
        O = (V, D) => {
          const B = I(V, D),
            _ = new Map();
          return (
            V.rows.forEach((C, T) => {
              if (B.some((L) => Fj(C.cells[L]) || Vf(C.cells[L]))) {
                F(
                  D,
                  V,
                  T,
                  "Blank or formula key: record identity is unresolved; formulas are not evaluated.",
                );
                return;
              }
              const N = JSON.stringify(
                  B.map((L) => [typeof yp(C.cells[L]), yp(C.cells[L])]),
                ),
                R = _.get(N) ?? [];
              (R.push(T), _.set(N, R));
            }),
            _
          );
        },
        j = O(y, "before"),
        z = O(w, "after");
      for (const V of new Set([...j.keys(), ...z.keys()])) {
        const D = j.get(V) ?? [],
          B = z.get(V) ?? [];
        if (D.length > 1 || B.length > 1) {
          for (const _ of D)
            F(
              "before",
              y,
              _,
              `Duplicate key ambiguity: ${D.length} before / ${B.length} after records; no automatic pairing.`,
            );
          for (const _ of B)
            F(
              "after",
              w,
              _,
              `Duplicate key ambiguity: ${D.length} before / ${B.length} after records; no automatic pairing.`,
            );
        } else
          D.length && B.length
            ? A.push([D[0], B[0]])
            : D.length
              ? k("before", y, D[0])
              : k("after", w, B[0]);
      }
    }
    for (const [I, O] of A) {
      d.counts.matchedRows++;
      const j = y.rows[I],
        z = w.rows[O];
      b.rowMatch === "keys" &&
        j.row !== z.row &&
        u({
          kind: "moved-row",
          confidence: "observed",
          message: "The unique typed key matches at a different worksheet row.",
          ...E("before", y, j.row),
          ...E("after", w, z.row),
        });
      for (const V of new Set([...y.headers, ...w.headers])) {
        const D = y.headers.indexOf(V),
          B = w.headers.indexOf(V),
          _ = D < 0 ? undefined : j.cells[D],
          C = B < 0 ? undefined : z.cells[B];
        (D >= 0 && B >= 0 && ww(_ ?? { raw: null }, C ?? { raw: null })) ||
          u({
            kind: Vf(_) && !Vf(C) ? "formula-to-value" : "cell-change",
            confidence: "observed",
            message: `${V}: ${D < 0 ? "column added" : B < 0 ? "column removed" : Vf(_) && !Vf(C) ? "formula replaced by a literal or blank" : "stored value, formula or cell metadata changed"}.`,
            ...(D < 0
              ? {}
              : {
                  ...c("before", y.sheet, xc(j.row, y.firstCol + D)),
                  beforeValue: yp(_),
                }),
            ...(B < 0
              ? {}
              : {
                  ...c("after", w.sheet, xc(z.row, w.firstCol + B)),
                  afterValue: yp(C),
                }),
          });
      }
    }
    for (const [I, O] of [
      ["before", y],
      ["after", w],
    ])
      for (let j = 0; j < O.headers.length; j++) {
        const z = new Map();
        let V = 0;
        O.rows.forEach((_, C) => {
          if (Vf(_.cells[j])) {
            V++;
            try {
              const T = a2e(_.cells[j].raw, xc(_.row, O.firstCol + j)),
                N = z.get(T) ?? [];
              (N.push(C), z.set(T, N));
            } catch {
              u({
                kind: "uncertain",
                confidence: "unresolved",
                message: "Formula pattern could not be normalized reliably.",
                ...c(I, O.sheet, xc(_.row, O.firstCol + j)),
              });
            }
          }
        });
        const D = [...z.entries()].sort((_, C) => C[1].length - _[1].length),
          B = D[0];
        for (const _ of O.rows) {
          const C = _.cells[j];
          if (Vf(C) || Fj(C)) continue;
          const T =
            V > O.rows.length / 2 &&
            !!B &&
            B[1].length >= 2 &&
            B[1].length > V / 2;
          u({
            kind: "hardcode",
            confidence: T ? "candidate" : "observed",
            message: T
              ? `${O.headers[j]}: literal value in a formula-dominant column; a deliberate exception is possible.`
              : `${O.headers[j]}: stored literal input, not a formula. This inventory entry is not an error finding.`,
            ...c(I, O.sheet, xc(_.row, O.firstCol + j)),
            [I === "before" ? "beforeValue" : "afterValue"]: yp(C),
          });
        }
        if (!(V < 2)) {
          if (!B || B[1].length < 2 || B[1].length <= V / 2) {
            z.size > 1 &&
              u({
                kind: "uncertain",
                confidence: "unresolved",
                message: `${O.headers[j]}: no majority formula pattern; ${z.size} patterns among ${V} formulas. No anomaly winner was selected.`,
                ...c(I, O.sheet, O.range),
              });
            continue;
          }
          O.rows.forEach((_, C) => {
            const T = _.cells[j],
              N = xc(_.row, O.firstCol + j);
            Vf(T) &&
              !B[1].includes(C) &&
              u({
                kind: "formula-pattern",
                confidence: "candidate",
                message: `${O.headers[j]}: formula differs from the dominant relative-reference pattern (${B[1].length}/${V} formulas); review against ${xc(O.rows[B[1][0]].row, O.firstCol + j)}.`,
                ...c(I, O.sheet, N),
                [I === "before" ? "beforeValue" : "afterValue"]: yp(T),
              });
          });
        }
      }
    (y.sheet.name !== w.sheet.name || !ww(y.headers, w.headers)) &&
      u({
        kind: "metadata",
        confidence: "observed",
        message:
          "Explicitly matched worksheet names, header text or column order differ.",
        ...c("before", y.sheet, y.range),
        ...c("after", w.sheet, w.range),
      });
    for (const I of new Set([...Object.keys(y.sheet), ...Object.keys(w.sheet)]))
      ["cells", "id", "name"].includes(I) ||
        ww(y.sheet[I], w.sheet[I]) ||
        u({
          kind: "metadata",
          confidence: "observed",
          message: `Worksheet property changed: ${I}. Property content is not evaluated.`,
          ...c("before", y.sheet, y.range),
          ...c("after", w.sheet, w.range),
        });
  }
  let v = 0;
  for (const b of ["before", "after"]) {
    const y = s[b];
    for (const w of y.sheets) {
      if (!a[b].includes(w.id)) continue;
      const E = h(w),
        A = n.includeHidden || (!w.hidden && !w.veryHidden);
      (w.hidden || w.veryHidden) &&
        u({
          kind: "hidden-content",
          confidence: "observed",
          message: `${w.veryHidden ? "Very hidden" : "Hidden"} worksheet; ${A ? "content included by explicit permission" : "cell content withheld"}.`,
          ...c(b, w, ""),
        });
      for (const [k, F] of [
        ["row", E.rows],
        ["column", E.cols],
      ])
        for (const I of F)
          u({
            kind: "hidden-content",
            confidence: "observed",
            message: `Hidden ${k} interval, including collapsed outlines; ${n.includeHidden ? "content included" : "cell content withheld"}.`,
            ...c(
              b,
              w,
              k === "row"
                ? `${I.start + 1}:${I.end + 1}`
                : `${ut(I.start)}:${ut(I.end)}`,
            ),
          });
      if (
        (p[b].has(w.id) ||
          (d.coverage[
            b === "before" ? "beforeUnmappedSheets" : "afterUnmappedSheets"
          ]++,
          u({
            kind: "unmapped-sheet",
            confidence: "unresolved",
            message:
              "Worksheet has no explicit counterpart mapping. No insertion, deletion or identity match was inferred.",
            ...c(b, w, ""),
          })),
        !!A)
      )
        for (const [k, F] of Object.entries(w.cells)) {
          const I = _t(k);
          if (
            !n.includeHidden &&
            (E.rows.some((z) => z.start <= I.row && z.end >= I.row) ||
              E.cols.some((z) => z.start <= I.col && z.end >= I.col))
          )
            continue;
          if (++d.counts.scannedCells > $h.scannedCells)
            throw Error(
              "Static comparison supports 20,000 permitted stored cells across both workbooks.",
            );
          if (!Vf(F)) continue;
          d.counts.scannedFormulas++;
          const O = (z, V) =>
              u({
                kind: z,
                confidence: z === "uncertain" ? "unresolved" : "observed",
                message: V,
                ...c(b, w, k),
              }),
            j = (z, V, D = []) => {
              if (((v += z.length), v > $h.formulaCharacters))
                throw Error(
                  "Static formula inspection exceeds the 2,000,000-character work budget.",
                );
              if (z.length > 4096) {
                O(
                  "uncertain",
                  "Formula exceeds the 4,096-character inspection bound.",
                );
                return;
              }
              if (D.length >= $h.nameDepth) {
                O("uncertain", "Defined-name expansion exceeds depth 16.");
                return;
              }
              let B = z.startsWith("=") ? 1 : 0;
              const _ = (T, N) => {
                  if (!T) {
                    O("broken-link", "Referenced worksheet is missing.");
                    return;
                  }
                  if (
                    !a[b].includes(T.id) ||
                    (!n.includeHidden && (T.hidden || T.veryHidden))
                  ) {
                    O(
                      "uncertain",
                      "A reference points outside the permitted visible worksheet context; target details are withheld.",
                    );
                    return;
                  }
                  try {
                    const R = N.replace(/[\s$]/g, ""),
                      L = ct(
                        /^[A-Za-z]+:[A-Za-z]+$/.test(R)
                          ? R.replace(":", "1:") + Mt.maxRows
                          : /^\d+:\d+$/.test(R)
                            ? "A" + R.replace(":", ":XFD")
                            : R,
                      );
                    if (L.lastRow >= T.rowCount || L.lastCol >= T.colCount)
                      throw Error();
                    const $ = h(T);
                    !n.includeHidden &&
                      ($.rows.some(
                        (U) => U.start <= L.lastRow && U.end >= L.firstRow,
                      ) ||
                        $.cols.some(
                          (U) => U.start <= L.lastCol && U.end >= L.firstCol,
                        )) &&
                      O(
                        "uncertain",
                        "A reference intersects hidden cells; their content was not inspected.",
                      );
                  } catch {
                    O(
                      "broken-link",
                      "Reference has invalid or out-of-bounds coordinates.",
                    );
                  }
                },
                C = (T, N) => {
                  const R =
                    y.definedNames?.find(
                      (G) =>
                        G.scopeSheetId === N &&
                        G.name.toLowerCase() === T.toLowerCase(),
                    ) ??
                    y.definedNames?.find(
                      (G) =>
                        !G.scopeSheetId &&
                        G.name.toLowerCase() === T.toLowerCase(),
                    );
                  if (!R) {
                    O(
                      "broken-link",
                      "An unresolved defined-name reference is present.",
                    );
                    return;
                  }
                  const L = `${R.scopeSheetId ?? ""}:${R.name.toLowerCase()}`;
                  if (D.includes(L)) {
                    O(
                      "uncertain",
                      "Defined-name cycle: references cannot be completely inspected.",
                    );
                    return;
                  }
                  j(R.refersTo, R.scopeSheetId ?? N, [...D, L]);
                };
              for (; B < z.length;) {
                const T = z.slice(B);
                if (T[0] === '"') {
                  const X = /^"(?:[^"]|"")*"/.exec(T);
                  if (!X) {
                    O(
                      "uncertain",
                      "Unterminated formula text; link inspection is incomplete.",
                    );
                    break;
                  }
                  B += X[0].length;
                  continue;
                }
                const N =
                  /^(?:'(?:[^']|'')*\[[^\]]+\](?:[^']|'')*'\s*!\s*[^+*/(),;\s]+|\[[^\]]+\][^!+*/(),;]*![^+*/(),;\s]+)/.exec(
                    T,
                  );
                if (N) {
                  (O(
                    "external-link",
                    "External workbook link is inventoried but never opened; target validity is unverified.",
                  ),
                    (B += N[0].length));
                  continue;
                }
                if (n2e.test(T) || /^'(?:[^']|'')*:(?:[^']|'')*'\s*!/.test(T)) {
                  O(
                    "uncertain",
                    "Three-dimensional references require separate review and are not expanded.",
                  );
                  break;
                }
                const R = /^#[A-Za-z0-9/?!]+/.exec(T);
                if (R) {
                  (R[0].toUpperCase() === "#REF!" &&
                    O(
                      "broken-link",
                      "Stored #REF! token identifies a broken reference.",
                    ),
                    (B += R[0].length));
                  continue;
                }
                const L = /^([A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(T);
                if (L) {
                  const X = B;
                  B += L[0].length;
                  let J = 1;
                  for (; B < z.length && J;)
                    (z[B] === "[" && J++, z[B] === "]" && J--, B++);
                  const q = z.slice(X, B);
                  try {
                    const W = !J && z1(y, V, q, k);
                    W
                      ? _(
                          y.sheets.find((K) => K.id === W.sheetId),
                          W.range,
                        )
                      : O(
                          "uncertain",
                          "Structured table reference could not be resolved; table/column or syntax may be unsupported.",
                        );
                  } catch {
                    O(
                      "uncertain",
                      "Structured reference could not be inspected.",
                    );
                  }
                  continue;
                }
                const $ = /^([A-Za-z_\\][A-Za-z0-9_.\\]*)\s*\(/.exec(T);
                if ($) {
                  (/^(?:_xlfn\.)?(?:INDIRECT|OFFSET)$/i.test($[1]) &&
                    O(
                      "uncertain",
                      "Dynamic reference function is not executed; actual target is unverified.",
                    ),
                    /^(?:_xlfn\.)?(?:WEBSERVICE|RTD|CUBE\w*)$/i.test($[1]) &&
                      O(
                        "external-link",
                        "External/network-enabled function is inventoried but never executed.",
                      ),
                    (B += $[0].length));
                  continue;
                }
                const U = rH.exec(T);
                if (U && nH(T[U[0].length] ?? "")) {
                  const X = U[1] ? Bj(U[1]) : undefined;
                  (_(
                    X === undefined
                      ? y.sheets.find((J) => J.id === V)
                      : y.sheets.find(
                          (J) => J.name.toLowerCase() === X.toLowerCase(),
                        ),
                    U[2],
                  ),
                    U[3] &&
                      O(
                        "uncertain",
                        "Spill reference is inspected at its anchor; runtime extent is not calculated.",
                      ),
                    (B += U[0].length));
                  continue;
                }
                const G = r2e.exec(T);
                if (G) {
                  const X = y.sheets.find(
                    (J) => J.name.toLowerCase() === Bj(G[1]).toLowerCase(),
                  );
                  (X
                    ? !a[b].includes(X.id) ||
                      (!n.includeHidden && (X.hidden || X.veryHidden))
                      ? O(
                          "uncertain",
                          "Qualified name is outside the permitted context.",
                        )
                      : C(G[2], X.id)
                    : O(
                        "broken-link",
                        "Defined-name worksheet qualifier is missing.",
                      ),
                    (B += G[0].length));
                  continue;
                }
                const H = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(T);
                if (H) {
                  B += H[0].length;
                  continue;
                }
                const Q = i2e.exec(T);
                if (Q) {
                  (/^(?:TRUE|FALSE)$/i.test(Q[0]) ||
                    (/^R(?:\d+|\[-?\d+\])C(?:\d+|\[-?\d+\])$/i.test(Q[0])
                      ? O("uncertain", "R1C1 syntax is not resolved.")
                      : C(Q[0], V)),
                    (B += Q[0].length));
                  continue;
                }
                (T[0] === "|" &&
                  O(
                    "external-link",
                    "DDE-style syntax is inventoried and never executed.",
                  ),
                  B++);
              }
            };
          j(F.raw, w.id);
        }
    }
  }
  return (
    e.name !== r.name &&
      u({
        kind: "metadata",
        confidence: "observed",
        message:
          "Workbook names differ; identities and revisions are retained separately.",
      }),
    structuredClone(d)
  );
}
const l2e = [
    "all",
    "inserted-row",
    "deleted-row",
    "moved-row",
    "cell-change",
    "formula-to-value",
    "formula-pattern",
    "hardcode",
    "broken-link",
    "external-link",
    "uncertain",
    "hidden-content",
    "metadata",
    "unmapped-sheet",
  ],
  Lj = (e) =>
    e === undefined
      ? "—"
      : e === null
        ? "(blank)"
        : typeof e == "boolean"
          ? String(e).toUpperCase()
          : String(e);
function u2e(e) {
  const { text, language } = Dr(),
    n = yh(),
    o = ne.useId(),
    a = ne.useRef(null),
    s = ne.useRef(0),
    d = ne.useRef(true),
    [u, l] = ne.useState(null),
    [c, h] = ne.useState([]),
    [f, p] = ne.useState(false),
    [m, v] = ne.useState(false),
    [g, b] = ne.useState([]),
    [y, w] = ne.useState(null),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(""),
    [j, z] = ne.useState("all"),
    [V, D] = ne.useState(0),
    B = u ?? e.baseline,
    _ = [
      ...new Set(e.allowedSheetIds ?? e.workbook.sheets.map((oe) => oe.id)),
    ].sort(),
    C = [
      ...new Set(
        (u ? undefined : e.baselineAllowedSheetIds) ??
          B?.sheets.map((oe) => oe.id) ??
          [],
      ),
    ].sort(),
    T = !!e.allowHidden && m,
    N = e.workbook.sheets.filter(
      (oe) => _.includes(oe.id) && (T || (!oe.hidden && !oe.veryHidden)),
    ),
    R =
      B?.sheets.filter(
        (oe) => C.includes(oe.id) && (T || (!oe.hidden && !oe.veryHidden)),
      ) ?? [],
    L = JSON.stringify([
      e.workbook,
      e.sheetId,
      e.range,
      e.readOnly,
      e.allowedSheetIds,
      e.allowHidden,
      e.baseline,
      e.baselineAllowedSheetIds,
    ]),
    $ = ne.useRef(L);
  $.current = L;
  const U = JSON.stringify([L, B, g, T, C, _]),
    G = y && U === E,
    H = () => {
      (w(null), F(""), D(0));
    },
    Q = (oe, de = !!u) => {
      const he = oe.sheets.find(
        (Se) =>
          (de ||
            !e.baselineAllowedSheetIds ||
            e.baselineAllowedSheetIds.includes(Se.id)) &&
          (T || (!Se.hidden && !Se.veryHidden)),
      );
      return {
        before: {
          sheetId: he?.id ?? "",
          range: he ? Ta(he) : "A1",
        },
        after: { sheetId: e.sheetId, range: e.range },
        headers: true,
        rowMatch: "keys",
        keys: [{ before: "", after: "" }],
      };
    };
  (ne.useEffect(() => {
    var he;
    const oe = document.activeElement;
    return (
      (he = a.current?.querySelector(
        '[data-ui-label="Comparison baseline file"]',
      )) == null || he.focus(),
      (d.current = true),
      () => {
        ((d.current = false),
          s.current++,
          oe != null && oe.isConnected && oe.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      B && !g.length && b([Q(B)]);
    }, [B]));
  const X$1 = (oe, de) => {
      (b(g.map((he, Se) => (Se === oe ? de : he))), H());
    },
    J = async (oe) => {
      if (!oe) return;
      const de = ++s.current,
        he = L;
      (p(true), F(""), w(null), l(null), h([]), b([]));
      try {
        if (!/\.(json|xlsx)$/i.test(oe.name))
          throw Error("Choose a native JSON or XLSX comparison baseline.");
        const Se = await H1(oe, n.begin());
        if (!d.current || de !== s.current || $.current !== he) return;
        (l(Se.workbook), h(Se.warnings), b([Q(Se.workbook, true)]));
      } catch (Se) {
        d.current &&
          de === s.current &&
          $.current === he &&
          (F(Se instanceof Error ? Se.message : "Baseline import failed."),
          O(he));
      } finally {
        d.current && de === s.current && p(false);
      }
    },
    q = () => {
      if (B)
        try {
          (w(
            s2e(
              B,
              e.workbook,
              { mappings: g },
              {
                beforeAllowedSheetIds: C,
                afterAllowedSheetIds: _,
                includeHidden: T,
              },
            ),
          ),
            A(U),
            D(0),
            z("all"),
            F(""));
        } catch (oe) {
          (w(null),
            F(
              oe instanceof Error
                ? oe.message
                : "Comparison could not complete.",
            ),
            O(L));
        }
    },
    W = G ? y.findings.filter((oe) => j === "all" || oe.kind === j) : [],
    K = Math.max(1, Math.ceil(W.length / 25)),
    te = Math.min(V, K - 1),
    ae = () => {
      (s.current++, e.onClose());
    };
  return i.jsx("div", {
    className: "wc-backdrop",
    onMouseDown: (oe) => {
      oe.target === oe.currentTarget && ae();
    },
    children: i.jsxs("section", {
      ref: a,
      className: "wc-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${o}-title`,
      onKeyDown: (oe) => {
        var he, Se;
        if (
          (oe.stopPropagation(),
          oe.key === "Escape" && (oe.preventDefault(), ae()),
          oe.key === "Tab")
        ) {
          const le = [
            ...(a.current?.querySelectorAll(
              'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
            ) ?? []),
          ].filter((ce) => !ce.hidden);
          oe.shiftKey && document.activeElement === le[0]
            ? (oe.preventDefault(), (he = le.at(-1)) == null || he.focus())
            : !oe.shiftKey &&
              document.activeElement === le.at(-1) &&
              (oe.preventDefault(), (Se = le[0]) == null || Se.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${o}-title`,
                  children: i.jsx(M, { children: "Compare workbooks" }),
                }),
                i.jsx("p", {
                  children: i.jsx(M, {
                    children:
                      "Review record changes, formula consistency, links and hidden content.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close workbook comparison"),
              onClick: ae,
              lang: language("Close workbook comparison"),
              "data-ui-label": "Close workbook comparison",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "wc-body",
          children: [
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Comparison is read-only. Formulas, external links and imported cached values are never calculated. Both source workbooks stay unchanged.",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Comparison baseline file" }),
                i.jsx("input", {
                  "aria-label": text("Comparison baseline file"),
                  type: "file",
                  accept: ".json,.xlsx",
                  onChange: (oe) => {
                    (J(oe.target.files?.[0]), (oe.target.value = ""));
                  },
                  lang: language("Comparison baseline file"),
                  "data-ui-label": "Comparison baseline file",
                }),
              ],
            }),
            f &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(M, {
                  children: "Reading the comparison baseline…",
                }),
              }),
            B &&
              i.jsxs("p", {
                children: [
                  i.jsx(M, { children: "Before: " }),
                  i.jsx("strong", { children: B.name }),
                  " ",
                  i.jsx(M, { children: "· revision" }),
                  " ",
                  B.revision,
                  i.jsx(M, { children: ". After: " }),
                  i.jsx("strong", { children: e.workbook.name }),
                  " ",
                  i.jsx(M, { children: "· revision " }),
                  e.workbook.revision,
                  ".",
                ],
              }),
            c.length > 0 &&
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(M, { children: "Baseline import notices (" }),
                      c.length,
                      ")",
                    ],
                  }),
                  i.jsx("ul", {
                    children: c.map((oe, de) =>
                      i.jsx("li", { children: oe }, de),
                    ),
                  }),
                ],
              }),
            i.jsxs("label", {
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: T,
                  disabled: !e.allowHidden,
                  onChange: (oe) => {
                    (v(oe.target.checked), H());
                  },
                }),
                i.jsx(M, { children: "Include permitted hidden content" }),
              ],
            }),
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Hidden structure is inventoried. Its cell content requires explicit permission. Imported files are separate user-selected baselines; host worksheet restrictions apply independently.",
              }),
            }),
            B &&
              g.map((oe, de) => {
                return i.jsxs(
                  "fieldset",
                  {
                    children: [
                      i.jsxs("legend", {
                        children: [
                          i.jsx(M, { children: "Worksheet mapping " }),
                          de + 1,
                        ],
                      }),
                      i.jsxs("div", {
                        className: "wc-fields",
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "Before worksheet " }),
                              de + 1,
                              i.jsxs("select", {
                                value: oe.before.sheetId,
                                onChange: (le) =>
                                  X$1(de, {
                                    ...oe,
                                    before: {
                                      ...oe.before,
                                      sheetId: le.target.value,
                                    },
                                  }),
                                children: [
                                  i.jsx("option", {
                                    value: "",
                                    children: "Choose worksheet",
                                  }),
                                  R.map((le) =>
                                    i.jsx(
                                      "option",
                                      { value: le.id, children: le.name },
                                      le.id,
                                    ),
                                  ),
                                ],
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "Before range " }),
                              de + 1,
                              i.jsx("input", {
                                value: oe.before.range,
                                onChange: (le) =>
                                  X$1(de, {
                                    ...oe,
                                    before: {
                                      ...oe.before,
                                      range: le.target.value,
                                    },
                                  }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "After worksheet " }),
                              de + 1,
                              i.jsxs("select", {
                                value: oe.after.sheetId,
                                onChange: (le) =>
                                  X$1(de, {
                                    ...oe,
                                    after: {
                                      ...oe.after,
                                      sheetId: le.target.value,
                                    },
                                  }),
                                children: [
                                  i.jsx("option", {
                                    value: "",
                                    children: "Choose worksheet",
                                  }),
                                  N.map((le) =>
                                    i.jsx(
                                      "option",
                                      { value: le.id, children: le.name },
                                      le.id,
                                    ),
                                  ),
                                ],
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "After range " }),
                              de + 1,
                              i.jsx("input", {
                                value: oe.after.range,
                                onChange: (le) =>
                                  X$1(de, {
                                    ...oe,
                                    after: {
                                      ...oe.after,
                                      range: le.target.value,
                                    },
                                  }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "Row matching " }),
                              de + 1,
                              i.jsxs("select", {
                                value: oe.rowMatch,
                                onChange: (le) =>
                                  X$1(de, {
                                    ...oe,
                                    rowMatch: le.target.value,
                                    keys:
                                      le.target.value === "keys"
                                        ? [{ before: "", after: "" }]
                                        : undefined,
                                  }),
                                children: [
                                  i.jsx("option", {
                                    value: "keys",
                                    children: "Exact typed unique keys",
                                  }),
                                  i.jsx("option", {
                                    value: "position",
                                    children: "Explicit row position",
                                  }),
                                ],
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: oe.headers,
                                onChange: (le) =>
                                  X$1(de, {
                                    ...oe,
                                    headers: le.target.checked,
                                  }),
                              }),
                              i.jsx(M, {
                                children: "First row contains headers ",
                              }),
                              de + 1,
                            ],
                          }),
                        ],
                      }),
                      oe.rowMatch === "keys" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsx("p", {
                              children: i.jsx(M, {
                                children:
                                  "Enter exact header text. Without headers, use relative column labels A, B, C. Values of different types, case or whitespace do not match.",
                              }),
                            }),
                            oe.keys?.map((le, ce) =>
                              i.jsxs(
                                "div",
                                {
                                  className: "wc-keys",
                                  children: [
                                    i.jsxs("label", {
                                      children: [
                                        i.jsx(M, {
                                          children: "Before key ",
                                        }),
                                        de + 1,
                                        ".",
                                        ce + 1,
                                        i.jsx("input", {
                                          value: le.before,
                                          onChange: (ge) =>
                                            X$1(de, {
                                              ...oe,
                                              keys: oe.keys.map((se, we) =>
                                                we === ce
                                                  ? {
                                                      ...se,
                                                      before: ge.target.value,
                                                    }
                                                  : se,
                                              ),
                                            }),
                                        }),
                                      ],
                                    }),
                                    i.jsxs("label", {
                                      children: [
                                        i.jsx(M, {
                                          children: "After key ",
                                        }),
                                        de + 1,
                                        ".",
                                        ce + 1,
                                        i.jsx("input", {
                                          value: le.after,
                                          onChange: (ge) =>
                                            X$1(de, {
                                              ...oe,
                                              keys: oe.keys.map((se, we) =>
                                                we === ce
                                                  ? {
                                                      ...se,
                                                      after: ge.target.value,
                                                    }
                                                  : se,
                                              ),
                                            }),
                                        }),
                                      ],
                                    }),
                                    i.jsxs("button", {
                                      disabled: oe.keys.length < 2,
                                      onClick: () =>
                                        X$1(de, {
                                          ...oe,
                                          keys: oe.keys.filter(
                                            (ge, se) => se !== ce,
                                          ),
                                        }),
                                      children: [
                                        i.jsx(M, {
                                          children: "Remove key ",
                                        }),
                                        de + 1,
                                        ".",
                                        ce + 1,
                                      ],
                                    }),
                                  ],
                                },
                                ce,
                              ),
                            ),
                            i.jsxs("button", {
                              disabled: (oe.keys?.length ?? 0) >= 8,
                              onClick: () =>
                                X$1(de, {
                                  ...oe,
                                  keys: [
                                    ...(oe.keys ?? []),
                                    { before: "", after: "" },
                                  ],
                                }),
                              children: [
                                i.jsx(M, { children: "Add key to mapping " }),
                                de + 1,
                              ],
                            }),
                          ],
                        }),
                      i.jsxs("button", {
                        disabled: g.length < 2,
                        onClick: () => {
                          (b(g.filter((le, ce) => ce !== de)), H());
                        },
                        children: [
                          i.jsx(M, { children: "Remove mapping " }),
                          de + 1,
                        ],
                      }),
                    ],
                  },
                  de,
                );
              }),
            i.jsxs("div", {
              className: "wc-actions",
              children: [
                i.jsx("button", {
                  disabled: !B || g.length >= 20,
                  onClick: () => {
                    B && (b([...g, Q(B)]), H());
                  },
                  children: i.jsx(M, { children: "Add worksheet mapping" }),
                }),
                i.jsx("button", {
                  disabled: !B || f,
                  onClick: q,
                  children: i.jsx(M, {
                    children: "Compare selected workbooks",
                  }),
                }),
              ],
            }),
            k && I === L && i.jsx("p", { role: "alert", children: k }),
            y &&
              !G &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(M, {
                  children:
                    "Comparison details are withheld because the workbook, selection, settings or access changed. Compare again.",
                }),
              }),
            G &&
              i.jsxs("section", {
                "aria-label": text("Workbook audit comparison report"),
                className: "wc-report",
                lang: language("Workbook audit comparison report"),
                "data-ui-label": "Workbook audit comparison report",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(M, { children: "Comparison report" }),
                  }),
                  i.jsxs("p", {
                    children: [
                      y.counts.beforeRows,
                      " ",
                      i.jsx(M, { children: "before records ·" }),
                      " ",
                      y.counts.afterRows,
                      " ",
                      i.jsx(M, { children: "after records ·" }),
                      " ",
                      y.counts.matchedRows,
                      " ",
                      i.jsx(M, { children: "matched ·" }),
                      " ",
                      y.counts.insertedRows,
                      " ",
                      i.jsx(M, { children: "inserted ·" }),
                      " ",
                      y.counts.deletedRows,
                      " ",
                      i.jsx(M, { children: "deleted ·" }),
                      " ",
                      y.counts.unresolvedBeforeRows,
                      "/",
                      y.counts.unresolvedAfterRows,
                      " ",
                      i.jsx(M, { children: "unresolved before/after." }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      y.coverage.selectedCells,
                      " ",
                      i.jsx(M, { children: "selected cells." }),
                      " ",
                      y.counts.scannedCells,
                      " ",
                      i.jsx(M, { children: "permitted stored cells /" }),
                      " ",
                      y.counts.scannedFormulas,
                      " ",
                      i.jsx(M, { children: "formulas inspected statically." }),
                      " ",
                      y.findings.length,
                      " ",
                      i.jsx(M, { children: "findings." }),
                    ],
                  }),
                  i.jsxs("details", {
                    children: [
                      i.jsx("summary", {
                        children: i.jsx(M, {
                          children: "Match rules and coverage",
                        }),
                      }),
                      i.jsx("ul", {
                        children: y.coverage.notes.map((oe) =>
                          i.jsx("li", { children: oe }, oe),
                        ),
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx(M, { children: "Before workbook: " }),
                          y.before.id,
                          i.jsx(M, { children: ", revision" }),
                          " ",
                          y.before.revision,
                          i.jsx(M, { children: ". After workbook: " }),
                          y.after.id,
                          i.jsx(M, { children: ", revision " }),
                          y.after.revision,
                          ".",
                        ],
                      }),
                      y.request.mappings.map((oe, de) => {
                        return i.jsxs(
                          "p",
                          {
                            children: [
                              i.jsx(M, { children: "Mapping " }),
                              de + 1,
                              ": ",
                              oe.before.sheetId,
                              "!",
                              oe.before.range,
                              " →",
                              " ",
                              oe.after.sheetId,
                              "!",
                              oe.after.range,
                              "; ",
                              oe.rowMatch,
                              ";",
                              " ",
                              oe.headers ? "headers" : "no headers",
                              ";",
                              " ",
                              oe.keys
                                ?.map((Se) => `${Se.before} → ${Se.after}`)
                                .join(", "),
                              ".",
                            ],
                          },
                          de,
                        );
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Comparison finding category" }),
                      i.jsx("select", {
                        value: j,
                        onChange: (oe) => {
                          (z(oe.target.value), D(0));
                        },
                        children: l2e.map((oe) =>
                          i.jsxs(
                            "option",
                            {
                              value: oe,
                              children: [
                                oe,
                                " (",
                                oe === "all"
                                  ? y.findings.length
                                  : y.findings.filter((de) => de.kind === oe)
                                      .length,
                                ")",
                              ],
                            },
                            oe,
                          ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "wc-scroll",
                    tabIndex: 0,
                    role: "region",
                    "aria-label": text("Scrollable comparison findings"),
                    lang: language("Scrollable comparison findings"),
                    "data-ui-label": "Scrollable comparison findings",
                    children: [
                      i.jsxs("table", {
                        children: [
                          i.jsx("thead", {
                            children: i.jsxs("tr", {
                              children: [
                                i.jsx("th", {
                                  children: i.jsx(M, { children: "Finding" }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(M, {
                                    children: "Before address",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(M, {
                                    children: "Before raw value",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(M, {
                                    children: "After address",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(M, {
                                    children: "After raw value",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(M, {
                                    children: "Explanation",
                                  }),
                                }),
                              ],
                            }),
                          }),
                          i.jsx("tbody", {
                            children: W.slice(te * 25, (te + 1) * 25).map(
                              (oe, de) =>
                                i.jsxs(
                                  "tr",
                                  {
                                    children: [
                                      i.jsxs("td", {
                                        children: [
                                          oe.kind,
                                          i.jsx("small", {
                                            children: oe.confidence,
                                          }),
                                        ],
                                      }),
                                      i.jsx("td", {
                                        children: oe.before
                                          ? `${oe.before.sheetName}!${oe.before.address}`
                                          : "—",
                                      }),
                                      i.jsx("td", {
                                        children: Lj(oe.beforeValue),
                                      }),
                                      i.jsx("td", {
                                        children: oe.after
                                          ? `${oe.after.sheetName}!${oe.after.address}`
                                          : "—",
                                      }),
                                      i.jsx("td", {
                                        children: Lj(oe.afterValue),
                                      }),
                                      i.jsx("td", { children: oe.message }),
                                    ],
                                  },
                                  `${j}-${te}-${de}`,
                                ),
                            ),
                          }),
                        ],
                      }),
                      !W.length &&
                        i.jsx("p", {
                          children: i.jsx(M, {
                            children:
                              "No findings in this category within the declared scope.",
                          }),
                        }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "wc-actions",
                    children: [
                      i.jsx("button", {
                        disabled: te === 0,
                        onClick: () => D(te - 1),
                        children: i.jsx(M, { children: "Previous findings" }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(M, { children: "Page " }),
                          te + 1,
                          " ",
                          i.jsx(M, { children: "of " }),
                          K,
                          " · ",
                          W.length,
                          " ",
                          i.jsx(M, { children: "findings" }),
                        ],
                      }),
                      i.jsx("button", {
                        disabled: te + 1 >= K,
                        onClick: () => D(te + 1),
                        children: i.jsx(M, { children: "Next findings" }),
                      }),
                    ],
                  }),
                ],
              }),
            i.jsx("p", {
              className: "wc-limit",
              children: i.jsx(M, {
                children:
                  "Limits: 20 mappings, 10,000 selected cells, 1,000 rows and 100 columns per range; 20,000 permitted stored cells inspected. Large or incomplete comparisons return an explicit diagnostic.",
              }),
            }),
          ],
        }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: ae,
            children: i.jsx(M, { children: "Close comparison" }),
          }),
        }),
      ],
    }),
  });
}
const r0 = {
    maxQueryLength: 256,
    maxResults: 1e5,
    maxCalculatedCells: 1e4,
    maxTextCharacters: 1e7,
    maxPatternWork: 3e7,
    maxSpecialCells: 500,
    maxSpecialAreas: 256,
    maxBlankScan: 1e5,
  },
  Eo = (e) => {
    throw new Error(e);
  },
  xw = (e, r) => `${ut(r)}${e + 1}`,
  E_ = (e) => typeof e?.raw == "string" && e.raw.startsWith("=") && !e.literal,
  c2e = (e) =>
    e === null
      ? ""
      : typeof e == "boolean"
        ? e
          ? "TRUE"
          : "FALSE"
        : String(e);
function oA(e) {
  return (
    e &&
      typeof e == "object" &&
      (Object.values(e).forEach(oA), Object.freeze(e)),
    e
  );
}
function iH(e, r) {
  const t = Nr(e, r.sheetId),
    n = ff(r.ranges ?? [r.range]),
    o = n.map(ct);
  ct(r.range);
  const a = r.scope ?? "sheet",
    s = r.order ?? "rows";
  return (
    (!["selection", "sheet", "workbook"].includes(a) ||
      !["rows", "columns"].includes(s)) &&
      Eo("Choose a valid search scope and order."),
    r.includeHidden !== undefined &&
      typeof r.includeHidden != "boolean" &&
      Eo("Include hidden must be boolean."),
    o.some((d) => d.lastRow >= t.rowCount || d.lastCol >= t.colCount) &&
      Eo("The search selection exceeds the worksheet."),
    {
      ...r,
      ranges: n,
      scope: a,
      order: s,
      includeHidden: r.includeHidden ?? false,
    }
  );
}
function d2e(e, r) {
  const t = iH(e, r),
    n = r.lookIn ?? "raw";
  ((typeof r.query != "string" ||
    !r.query.length ||
    r.query.length > r0.maxQueryLength) &&
    Eo(
      "Enter 1–256 characters to find. Use Go To Special to select blank cells.",
    ),
    ["raw", "formulas", "calculated", "formatted"].includes(n) ||
      Eo("Choose a supported Look in option."));
  for (const o of [r.matchCase, r.wholeCell, r.wildcards])
    o !== undefined &&
      typeof o != "boolean" &&
      Eo("Search switches must be boolean.");
  return {
    ...t,
    query: r.query,
    lookIn: n,
    matchCase: r.matchCase ?? false,
    wholeCell: r.wholeCell ?? false,
    wildcards: r.wildcards ?? true,
  };
}
const oH = () => ({ text: 0, work: 0, calculated: 0, scanned: 0 });
function aH(e) {
  const r = Array.from(e.query),
    t = [];
  for (let o = 0; o < r.length; o++) {
    const a = r[o];
    e.wildcards && a === "~" && ["~", "?", "*"].includes(r[o + 1])
      ? t.push({ kind: "literal", value: r[++o] })
      : e.wildcards && a === "*"
        ? t.at(-1)?.kind !== "star" && t.push({ kind: "star" })
        : e.wildcards && a === "?"
          ? t.push({ kind: "any" })
          : t.push({ kind: "literal", value: a });
  }
  return t;
}
function sH(e, r, t, n) {
  const o = Array.from(e),
    a = o.length;
  if (
    ((n.text += e.length),
    (n.work += (a + 1) * t.length),
    (n.text > r0.maxTextCharacters || n.work > r0.maxPatternWork) &&
      Eo(
        "This search exceeds the text-comparison budget. Narrow the selection or simplify the search.",
      ),
    !a)
  )
    return [];
  let s = Int32Array.from({ length: a + 1 }, (l, c) => c);
  for (let l = t.length - 1; l >= 0; l--) {
    const c = t[l],
      h = new Int32Array(a + 1).fill(-1);
    if (c.kind === "star") {
      h[a] = s[a];
      for (let f = a - 1; f >= 0; f--) h[f] = Math.max(s[f], h[f + 1]);
    } else
      for (let f = a - 1; f >= 0; f--)
        (c.kind === "any" ||
          (r.matchCase
            ? o[f] === c.value
            : o[f].toLowerCase() === c.value.toLowerCase())) &&
          (h[f] = s[f + 1]);
    s = h;
  }
  const d = [0];
  for (const l of o) d.push(d.at(-1) + l.length);
  if (r.wholeCell) return s[0] === a ? [[0, e.length]] : [];
  const u = [];
  for (let l = 0; l < a;)
    s[l] > l ? (u.push([d[l], d[s[l]]]), (l = s[l])) : l++;
  return u;
}
function lH(e) {
  const r = ul(e.rowCount, cn(e, "row")),
    t = ul(e.colCount, cn(e, "column"));
  return (n, o) =>
    e.hidden
      ? e.veryHidden
        ? "very hidden sheet"
        : "hidden sheet"
      : r.isHidden(n)
        ? "hidden row"
        : t.isHidden(o)
          ? "hidden column"
          : undefined;
}
const uH = (e, r, t) =>
  e >= t.firstRow && e <= t.lastRow && r >= t.firstCol && r <= t.lastCol;
function aA(e, r, t) {
  return e === "columns"
    ? r.col - t.col || r.row - t.row
    : r.row - t.row || r.col - t.col;
}
function cH(e, r) {
  var h;
  const t = d2e(e, r),
    n = oH(),
    o = aH(t),
    a = [],
    s = Vn(e),
    d = t.scope === "workbook" ? e.sheets : [Nr(e, t.sheetId)],
    u = t.ranges.map(ct),
    l = t.lookIn === "calculated" || t.lookIn === "formatted";
  for (const f of d) {
    if (f.hidden && !t.includeHidden) continue;
    const p = new Set(Object.keys(f.cells)),
      m = lH(f),
      v = [];
    if (l) for (const g of s.getSpillEntries(f.id)) p.add(g.address);
    for (const g of p) {
      ++n.scanned > Mt.maxCells &&
        Eo("This search exceeds 100,000 populated cells. Narrow its scope.");
      const { row, col } = _t(g),
        w = m(row, col),
        E = f.cells[g];
      if (
        (t.scope === "selection" && !u.some((O) => uH(row, col, O))) ||
        (w && !t.includeHidden)
      )
        continue;
      const A = E_(E);
      if (t.lookIn === "formulas" && !A) continue;
      let k = E?.raw ?? null;
      l &&
        (++n.calculated > r0.maxCalculatedCells &&
          Eo(
            "Calculated searches support at most 10,000 cells. Narrow the selection or use Raw contents.",
          ),
        (k = s(f.id, g)));
      const F =
          t.lookIn === "formatted"
            ? Af(k, E, e.datePolicy?.dateSystem)
            : c2e(k),
        I = sH(F, t, o, n);
      I.length &&
        (v.push({
          sheetId: f.id,
          sheetName: f.name,
          address: g,
          row,
          col,
          text: F,
          raw: E?.raw ?? null,
          isFormula: A,
          protected: !!((h = f.protection) != null && h.enabled),
          hidden: !!w,
          ...(w ? { hiddenReason: w } : {}),
          occurrences: I.length,
        }),
        a.length + v.length > r0.maxResults &&
          Eo("Too many matching cells. Narrow the search."));
    }
    (v.sort((g, b) => aA(t.order, g, b)), a.push(...v));
  }
  return oA({
    workbookId: e.id,
    revision: e.revision,
    snapshotGuard: JSON.stringify(e),
    sheetOrder: e.sheets.map((f) => f.id),
    options: { ...t },
    total: a.length,
    scannedCells: n.scanned,
    results: a,
  });
}
function f2e(e, r = 0, t = 25) {
  return (
    (!Number.isSafeInteger(r) ||
      r < 0 ||
      !Number.isSafeInteger(t) ||
      t < 1 ||
      t > 200) &&
      Eo("Search pages require a nonnegative page and 1–200 results per page."),
    e.results.slice(r * t, (r + 1) * t)
  );
}
function h2e(e, r, t = "next") {
  if (
    (["next", "previous"].includes(t) || Eo("Choose next or previous."),
    !e.total)
  )
    return null;
  const n = r
    ? e.results.findIndex(
        (o) => o.sheetId === r.sheetId && o.address === r.address.toUpperCase(),
      )
    : -1;
  if (n >= 0)
    return e.results[(n + (t === "next" ? 1 : e.total - 1)) % e.total];
  if (r) {
    const o = _t(r.address),
      a = new Map(e.sheetOrder.map((d, u) => [d, u])),
      s = a.get(r.sheetId);
    if (s !== undefined) {
      const d = (u) => a.get(u.sheetId) - s || aA(e.options.order, u, o);
      return t === "next"
        ? (e.results.find((u) => d(u) > 0) ?? e.results[0])
        : ([...e.results].reverse().find((u) => d(u) < 0) ?? e.results.at(-1));
    }
  }
  return t === "next" ? e.results[0] : e.results.at(-1);
}
function p2e(e, r) {
  let t = r;
  r.length
    ? typeof e.raw == "number" &&
      /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(r) &&
      Number.isFinite(Number(r))
      ? (t = Number(r))
      : typeof e.raw == "boolean" &&
        /^(true|false)$/i.test(r) &&
        (t = r.toLowerCase() === "true")
    : (t = null);
  const n = { ...e, raw: t };
  return (
    typeof t == "string" && t.startsWith("=") && !E_(e)
      ? (n.literal = true)
      : (typeof t == "string" && t.startsWith("=")) || delete n.literal,
    hi(n)
  );
}

export {
  E_,
  Eo,
  aA,
  aH,
  cH,
  f2e,
  h2e,
  iH,
  lH,
  oA,
  oH,
  p2e,
  r0,
  sH,
  u2e,
  uH,
  xw,
};
