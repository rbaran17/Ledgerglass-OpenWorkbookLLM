import { isFormulaInputValue } from "./11-local-workbook-store-and-versions.js";
import { eI, formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { throwWithMessage, isFormulaCellValue } from "./03-workbook-model.js";
import { coerceCellValueToText } from "./16-transformation-recipe-engine.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  formatNumberWithPattern,
  findSheetOrThrow,
  formatOccupiedRangeAddress,
  calculateWorkbookWithSettings,
  _t,
  cn,
  parseRangeAddress,
  parseSelectionAreas,
  normalizeCellRecordInput,
  createAxisVisibilityIndex,
  columnIndexToLetters,
  rewriteStructuredReferences,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { importWorkbookFile } from "./07-file-import-and-worker-dispatch.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import {
  patternCheckBudgets,
  Bj,
  isBlankFormulaCell,
  computeFormulaSignature,
  identifierPattern,
  sheetSpanPattern,
  isNonReferenceToken,
  definedNameReferencePattern,
  referenceTokenPattern,
  ww,
  readFormulaText,
} from "./51-pattern-checks-and-citations.js";

function runWorkbookComparison(before, workbook, t, n = {}) {
  if (
    !t ||
    !Array.isArray(t.mappings) ||
    !t.mappings.length ||
    t.mappings.length > patternCheckBudgets.mappings
  )
    throw Error("Choose 1–20 explicit worksheet/range mappings.");
  if (n.includeHidden !== undefined && typeof n.includeHidden != "boolean")
    throw Error("Include hidden must be boolean.");
  const o = (workbook, y) => {
      if (
        y !== undefined &&
        (!Array.isArray(y) ||
          y.some(
            (w) =>
              typeof w != "string" ||
              !workbook.sheets.some((sheet) => sheet.id === w),
          ))
      )
        throw Error("The comparison worksheet access context is invalid.");
      return [...new Set(y ?? workbook.sheets.map((sheet) => sheet.id))].sort();
    },
    a = {
      before: o(before, n.beforeAllowedSheetIds),
      after: o(workbook, n.afterAllowedSheetIds),
    },
    s = { before: before, after: workbook },
    d = {
      before: { id: before.id, name: before.name, revision: before.revision },
      after: {
        id: workbook.id,
        name: workbook.name,
        revision: workbook.revision,
      },
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
      if (d.findings.length >= patternCheckBudgets.findings)
        throw Error(
          "Comparison exceeds 30,000 findings; narrow the source workbooks.",
        );
      d.findings.push(b);
    },
    l = (side, y, address) => ({
      side: side,
      workbookId: s[side].id,
      revision: s[side].revision,
      sheetId: y.id,
      sheetName: y.name,
      address: address,
    }),
    c = (b, y, w) => ({ [b]: l(b, y, w) }),
    h = (b) => ({ rows: cn(b, "row"), cols: cn(b, "column") }),
    f = (b, y, w) => {
      if (!y || typeof y.sheetId != "string" || typeof y.range != "string")
        throw Error(
          "Choose explicit worksheet and rectangular range references.",
        );
      const sheet = s[b].sheets.find((sheet) => sheet.id === y.sheetId);
      if (!sheet || !a[b].includes(sheet.id))
        throw Error(
          "A selected source is outside the authorized comparison context.",
        );
      const rangeAddress = parseRangeAddress(y.range),
        k = rangeAddress.lastRow - rangeAddress.firstRow + 1,
        F = rangeAddress.lastCol - rangeAddress.firstCol + 1;
      if (
        k > patternCheckBudgets.rowsPerRange ||
        F > patternCheckBudgets.columnsPerRange ||
        rangeAddress.lastRow >= sheet.rowCount ||
        rangeAddress.lastCol >= sheet.colCount
      )
        throw Error(
          "Each comparison range supports at most 1,000 rows and 100 columns within the worksheet.",
        );
      if (
        ((d.coverage.selectedCells += k * F),
        d.coverage.selectedCells > patternCheckBudgets.selectedCells)
      )
        throw Error(
          "Comparison supports at most 10,000 selected cells across both workbooks.",
        );
      const I = h(sheet);
      if (
        !n.includeHidden &&
        (sheet.hidden ||
          sheet.veryHidden ||
          I.rows.some(
            (row) =>
              row.start <= rangeAddress.lastRow &&
              row.end >= rangeAddress.firstRow,
          ) ||
          I.cols.some(
            (col) =>
              col.start <= rangeAddress.lastCol &&
              col.end >= rangeAddress.firstCol,
          ))
      )
        throw Error(
          "A selected range contains hidden content. Explicitly include permitted hidden content or choose a visible range.",
        );
      const headers = Array.from({ length: F }, (j, z) => {
        if (!w) return columnIndexToLetters(z);
        const cell =
          sheet.cells[
            formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol + z)
          ];
        if (
          isFormulaInputValue(cell) ||
          typeof cell?.raw != "string" ||
          !cell.raw.trim()
        )
          throw Error(
            "Header rows require nonempty literal text in every column. Turn off headers for an unstructured sheet.",
          );
        return cell.raw;
      });
      if (
        new Set(headers.map((header) => header.toLowerCase())).size !==
        headers.length
      )
        throw Error(
          "Duplicate headers are ambiguous; resolve them or compare without headers.",
        );
      return {
        sheet: sheet,
        range: `${formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol)}`,
        headers: headers,
        firstCol: rangeAddress.firstCol,
        rows: Array.from({ length: k - Number(w) }, (j, z) => {
          const row = rangeAddress.firstRow + Number(w) + z;
          return {
            row: row,
            cells: headers.map(
              (header, B) =>
                sheet.cells[formatCellAddress(row, rangeAddress.firstCol + B)],
            ),
          };
        }),
      };
    },
    p = { before: new Set(), after: new Set() },
    m = { before: new Map(), after: new Map() };
  for (const mapping of t.mappings) {
    if (
      !mapping ||
      typeof mapping.headers != "boolean" ||
      !["keys", "position"].includes(mapping.rowMatch)
    )
      throw Error(
        "Each mapping requires explicit header and row-matching rules.",
      );
    for (const I of ["before", "after"]) {
      const O = mapping[I],
        rangeAddress = parseRangeAddress(O.range),
        z = m[I].get(O.sheetId) ?? [];
      if (
        z.some(
          (V) =>
            V.firstRow <= rangeAddress.lastRow &&
            V.lastRow >= rangeAddress.firstRow &&
            V.firstCol <= rangeAddress.lastCol &&
            V.lastCol >= rangeAddress.firstCol,
        )
      )
        throw Error("Comparison mappings may not overlap on either side.");
      (m[I].set(O.sheetId, [...z, rangeAddress]), p[I].add(O.sheetId));
    }
    const y = f("before", mapping.before, mapping.headers),
      w = f("after", mapping.after, mapping.headers);
    ((d.counts.beforeRows += y.rows.length),
      (d.counts.afterRows += w.rows.length));
    const E = (I, O, j) =>
        c(
          I,
          O.sheet,
          `${formatCellAddress(j, O.firstCol)}:${formatCellAddress(j, O.firstCol + O.headers.length - 1)}`,
        ),
      A = [],
      k = (I, O, j) => {
        (d.counts[I === "before" ? "deletedRows" : "insertedRows"]++,
          u({
            kind: I === "before" ? "deleted-row" : "inserted-row",
            confidence: "observed",
            message: `${I === "before" ? "Deleted" : "Inserted"} record under the declared ${mapping.rowMatch} match rule.`,
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
    if (mapping.rowMatch === "position") {
      if (mapping.keys?.length)
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
      if (
        !Array.isArray(mapping.keys) ||
        !mapping.keys.length ||
        mapping.keys.length > patternCheckBudgets.keys
      )
        throw Error(
          "Key matching requires 1–8 explicit header pairs (A, B, … without headers).",
        );
      const I = (V, D) => {
          const B = mapping.keys.map((key) => V.headers.indexOf(key[D]));
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
            V.rows.forEach((row, T) => {
              if (
                B.some(
                  (L) =>
                    isBlankFormulaCell(row.cells[L]) ||
                    isFormulaInputValue(row.cells[L]),
                )
              ) {
                F(
                  D,
                  V,
                  T,
                  "Blank or formula key: record identity is unresolved; formulas are not evaluated.",
                );
                return;
              }
              const N = JSON.stringify(
                  B.map((L) => [
                    typeof readFormulaText(row.cells[L]),
                    readFormulaText(row.cells[L]),
                  ]),
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
      mapping.rowMatch === "keys" &&
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
            kind:
              isFormulaInputValue(_) && !isFormulaInputValue(C)
                ? "formula-to-value"
                : "cell-change",
            confidence: "observed",
            message: `${V}: ${D < 0 ? "column added" : B < 0 ? "column removed" : isFormulaInputValue(_) && !isFormulaInputValue(C) ? "formula replaced by a literal or blank" : "stored value, formula or cell metadata changed"}.`,
            ...(D < 0
              ? {}
              : {
                  ...c(
                    "before",
                    y.sheet,
                    formatCellAddress(j.row, y.firstCol + D),
                  ),
                  beforeValue: readFormulaText(_),
                }),
            ...(B < 0
              ? {}
              : {
                  ...c(
                    "after",
                    w.sheet,
                    formatCellAddress(z.row, w.firstCol + B),
                  ),
                  afterValue: readFormulaText(C),
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
        O.rows.forEach((row, C) => {
          if (isFormulaInputValue(row.cells[j])) {
            V++;
            try {
              const formulaSignature = computeFormulaSignature(
                  row.cells[j].raw,
                  formatCellAddress(row.row, O.firstCol + j),
                ),
                N = z.get(formulaSignature) ?? [];
              (N.push(C), z.set(formulaSignature, N));
            } catch {
              u({
                kind: "uncertain",
                confidence: "unresolved",
                message: "Formula pattern could not be normalized reliably.",
                ...c(I, O.sheet, formatCellAddress(row.row, O.firstCol + j)),
              });
            }
          }
        });
        const D = [...z.entries()].sort((_, C) => C[1].length - _[1].length),
          B = D[0];
        for (const row of O.rows) {
          const C = row.cells[j];
          if (isFormulaInputValue(C) || isBlankFormulaCell(C)) continue;
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
            ...c(I, O.sheet, formatCellAddress(row.row, O.firstCol + j)),
            [I === "before" ? "beforeValue" : "afterValue"]: readFormulaText(C),
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
          O.rows.forEach((row, C) => {
            const T = row.cells[j],
              cellAddress = formatCellAddress(row.row, O.firstCol + j);
            isFormulaInputValue(T) &&
              !B[1].includes(C) &&
              u({
                kind: "formula-pattern",
                confidence: "candidate",
                message: `${O.headers[j]}: formula differs from the dominant relative-reference pattern (${B[1].length}/${V} formulas); review against ${formatCellAddress(O.rows[B[1][0]].row, O.firstCol + j)}.`,
                ...c(I, O.sheet, cellAddress),
                [I === "before" ? "beforeValue" : "afterValue"]:
                  readFormulaText(T),
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
    const workbook = s[b];
    for (const sheet of workbook.sheets) {
      if (!a[b].includes(sheet.id)) continue;
      const E = h(sheet),
        A = n.includeHidden || (!sheet.hidden && !sheet.veryHidden);
      (sheet.hidden || sheet.veryHidden) &&
        u({
          kind: "hidden-content",
          confidence: "observed",
          message: `${sheet.veryHidden ? "Very hidden" : "Hidden"} worksheet; ${A ? "content included by explicit permission" : "cell content withheld"}.`,
          ...c(b, sheet, ""),
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
              sheet,
              k === "row"
                ? `${I.start + 1}:${I.end + 1}`
                : `${columnIndexToLetters(I.start)}:${columnIndexToLetters(I.end)}`,
            ),
          });
      if (
        (p[b].has(sheet.id) ||
          (d.coverage[
            b === "before" ? "beforeUnmappedSheets" : "afterUnmappedSheets"
          ]++,
          u({
            kind: "unmapped-sheet",
            confidence: "unresolved",
            message:
              "Worksheet has no explicit counterpart mapping. No insertion, deletion or identity match was inferred.",
            ...c(b, sheet, ""),
          })),
        !!A)
      )
        for (const [k, cell] of Object.entries(sheet.cells)) {
          const I = _t(k);
          if (
            !n.includeHidden &&
            (E.rows.some((row) => row.start <= I.row && row.end >= I.row) ||
              E.cols.some((col) => col.start <= I.col && col.end >= I.col))
          )
            continue;
          if (++d.counts.scannedCells > patternCheckBudgets.scannedCells)
            throw Error(
              "Static comparison supports 20,000 permitted stored cells across both workbooks.",
            );
          if (!isFormulaInputValue(cell)) continue;
          d.counts.scannedFormulas++;
          const O = (z, V) =>
              u({
                kind: z,
                confidence: z === "uncertain" ? "unresolved" : "observed",
                message: V,
                ...c(b, sheet, k),
              }),
            j = (z, V, D = []) => {
              if (((v += z.length), v > patternCheckBudgets.formulaCharacters))
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
              if (D.length >= patternCheckBudgets.nameDepth) {
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
                      rangeAddress = parseRangeAddress(
                        /^[A-Za-z]+:[A-Za-z]+$/.test(R)
                          ? R.replace(":", "1:") + workbookLimits.maxRows
                          : /^\d+:\d+$/.test(R)
                            ? "A" + R.replace(":", ":XFD")
                            : R,
                      );
                    if (
                      rangeAddress.lastRow >= T.rowCount ||
                      rangeAddress.lastCol >= T.colCount
                    )
                      throw Error();
                    const $ = h(T);
                    !n.includeHidden &&
                      ($.rows.some(
                        (row) =>
                          row.start <= rangeAddress.lastRow &&
                          row.end >= rangeAddress.firstRow,
                      ) ||
                        $.cols.some(
                          (col) =>
                            col.start <= rangeAddress.lastCol &&
                            col.end >= rangeAddress.firstCol,
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
                    workbook.definedNames?.find(
                      (definedName) =>
                        definedName.scopeSheetId === N &&
                        definedName.name.toLowerCase() === T.toLowerCase(),
                    ) ??
                    workbook.definedNames?.find(
                      (definedName) =>
                        !definedName.scopeSheetId &&
                        definedName.name.toLowerCase() === T.toLowerCase(),
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
                if (
                  sheetSpanPattern.test(T) ||
                  /^'(?:[^']|'')*:(?:[^']|'')*'\s*!/.test(T)
                ) {
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
                    const W =
                      !J && rewriteStructuredReferences(workbook, V, q, k);
                    W
                      ? _(
                          workbook.sheets.find(
                            (sheet) => sheet.id === W.sheetId,
                          ),
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
                const U = referenceTokenPattern.exec(T);
                if (U && isNonReferenceToken(T[U[0].length] ?? "")) {
                  const X = U[1] ? Bj(U[1]) : undefined;
                  (_(
                    X === undefined
                      ? workbook.sheets.find((sheet) => sheet.id === V)
                      : workbook.sheets.find(
                          (sheet) =>
                            sheet.name.toLowerCase() === X.toLowerCase(),
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
                const G = definedNameReferencePattern.exec(T);
                if (G) {
                  const sheet = workbook.sheets.find(
                    (sheet) =>
                      sheet.name.toLowerCase() === Bj(G[1]).toLowerCase(),
                  );
                  (sheet
                    ? !a[b].includes(sheet.id) ||
                      (!n.includeHidden && (sheet.hidden || sheet.veryHidden))
                      ? O(
                          "uncertain",
                          "Qualified name is outside the permitted context.",
                        )
                      : C(G[2], sheet.id)
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
                const Q = identifierPattern.exec(T);
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
          j(cell.raw, sheet.id);
        }
    }
  }
  return (
    before.name !== workbook.name &&
      u({
        kind: "metadata",
        confidence: "observed",
        message:
          "Workbook names differ; identities and revisions are retained separately.",
      }),
    structuredClone(d)
  );
}
const comparisonFindingKinds = [
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
  formatComparisonValue = (e) =>
    e === undefined
      ? "—"
      : e === null
        ? "(blank)"
        : typeof e == "boolean"
          ? String(e).toUpperCase()
          : String(e);
function WorkbookComparisonDialog(e) {
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    o = ne.useId(),
    a = ne.useRef(null),
    s = ne.useRef(0),
    d = ne.useRef(true),
    [u, l] = ne.useState(null),
    [c, h] = ne.useState([]),
    [f, p] = ne.useState(false),
    [m, v] = ne.useState(false),
    [mappings, b] = ne.useState([]),
    [y, w] = ne.useState(null),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(""),
    [j, z] = ne.useState("all"),
    [V, D] = ne.useState(0),
    workbook = u ?? e.baseline,
    afterAllowedSheetIds = [
      ...new Set(
        e.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    beforeAllowedSheetIds = [
      ...new Set(
        (u ? undefined : e.baselineAllowedSheetIds) ??
          workbook?.sheets.map((sheet) => sheet.id) ??
          [],
      ),
    ].sort(),
    T = !!e.allowHidden && m,
    N = e.workbook.sheets.filter(
      (sheet) =>
        afterAllowedSheetIds.includes(sheet.id) &&
        (T || (!sheet.hidden && !sheet.veryHidden)),
    ),
    R =
      workbook?.sheets.filter(
        (sheet) =>
          beforeAllowedSheetIds.includes(sheet.id) &&
          (T || (!sheet.hidden && !sheet.veryHidden)),
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
  const U = JSON.stringify([
      L,
      workbook,
      mappings,
      T,
      beforeAllowedSheetIds,
      afterAllowedSheetIds,
    ]),
    G = y && U === E,
    H = () => {
      (w(null), F(""), D(0));
    },
    Q = (workbook, de = !!u) => {
      const sheet = workbook.sheets.find(
        (sheet) =>
          (de ||
            !e.baselineAllowedSheetIds ||
            e.baselineAllowedSheetIds.includes(sheet.id)) &&
          (T || (!sheet.hidden && !sheet.veryHidden)),
      );
      return {
        before: {
          sheetId: sheet?.id ?? "",
          range: sheet ? formatOccupiedRangeAddress(sheet) : "A1",
        },
        after: { sheetId: e.sheetId, range: e.range },
        headers: true,
        rowMatch: "keys",
        keys: [{ before: "", after: "" }],
      };
    };
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      a.current
        ?.querySelector('[data-ui-label="Comparison baseline file"]')
        ?.focus(),
      (d.current = true),
      () => {
        ((d.current = false),
          s.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      workbook && !mappings.length && b([Q(workbook)]);
    }, [workbook]));
  const X$1 = (oe, de) => {
      (b(mappings.map((mapping, Se) => (Se === oe ? de : mapping))), H());
    },
    J = async (oe) => {
      if (!oe) return;
      const de = ++s.current,
        he = L;
      (p(true), F(""), w(null), l(null), h([]), b([]));
      try {
        if (!/\.(json|xlsx)$/i.test(oe.name))
          throw Error("Choose a native JSON or XLSX comparison baseline.");
        const Se = await importWorkbookFile(oe, cancellableOperation.begin());
        if (!d.current || de !== s.current || $.current !== he) return;
        (l(Se.workbook), h(Se.warnings), b([Q(Se.workbook, true)]));
      } catch (error) {
        d.current &&
          de === s.current &&
          $.current === he &&
          (F(
            error instanceof Error ? error.message : "Baseline import failed.",
          ),
          O(he));
      } finally {
        d.current && de === s.current && p(false);
      }
    },
    onClick = () => {
      if (workbook)
        try {
          (w(
            runWorkbookComparison(
              workbook,
              e.workbook,
              { mappings: mappings },
              {
                beforeAllowedSheetIds: beforeAllowedSheetIds,
                afterAllowedSheetIds: afterAllowedSheetIds,
                includeHidden: T,
              },
            ),
          ),
            A(U),
            D(0),
            z("all"),
            F(""));
        } catch (error) {
          (w(null),
            F(
              error instanceof Error
                ? error.message
                : "Comparison could not complete.",
            ),
            O(L));
        }
    },
    W = G
      ? y.findings.filter((finding) => j === "all" || finding.kind === j)
      : [],
    K = Math.max(1, Math.ceil(W.length / 25)),
    te = Math.min(V, K - 1),
    ae = () => {
      (s.current++, e.onClose());
    };
  return i.jsx("div", {
    className: "wc-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && ae();
    },
    children: i.jsxs("section", {
      ref: a,
      className: "wc-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${o}-title`,
      onKeyDown: (event) => {
        var he, Se;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), ae()),
          event.key === "Tab")
        ) {
          const le = [
            ...(a.current?.querySelectorAll(
              'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
            ) ?? []),
          ].filter((ce) => !ce.hidden);
          event.shiftKey && document.activeElement === le[0]
            ? (event.preventDefault(), (he = le.at(-1)) == null || he.focus())
            : !event.shiftKey &&
              document.activeElement === le.at(-1) &&
              (event.preventDefault(), (Se = le[0]) == null || Se.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${o}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Compare workbooks",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
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
              children: i.jsx(LocalizedText, {
                children:
                  "Comparison is read-only. Formulas, external links and imported cached values are never calculated. Both source workbooks stay unchanged.",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Comparison baseline file" }),
                i.jsx("input", {
                  "aria-label": text("Comparison baseline file"),
                  type: "file",
                  accept: ".json,.xlsx",
                  onChange: (event) => {
                    (J(event.target.files?.[0]), (event.target.value = ""));
                  },
                  lang: language("Comparison baseline file"),
                  "data-ui-label": "Comparison baseline file",
                }),
              ],
            }),
            f &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children: "Reading the comparison baseline…",
                }),
              }),
            workbook &&
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Before: " }),
                  i.jsx("strong", { children: workbook.name }),
                  " ",
                  i.jsx(LocalizedText, { children: "· revision" }),
                  " ",
                  workbook.revision,
                  i.jsx(LocalizedText, { children: ". After: " }),
                  i.jsx("strong", { children: e.workbook.name }),
                  " ",
                  i.jsx(LocalizedText, { children: "· revision " }),
                  e.workbook.revision,
                  ".",
                ],
              }),
            c.length > 0 &&
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Baseline import notices (",
                      }),
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
                  onChange: (event) => {
                    (v(event.target.checked), H());
                  },
                }),
                i.jsx(LocalizedText, {
                  children: "Include permitted hidden content",
                }),
              ],
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Hidden structure is inventoried. Its cell content requires explicit permission. Imported files are separate user-selected baselines; host worksheet restrictions apply independently.",
              }),
            }),
            workbook &&
              mappings.map((mapping, de) =>
                i.jsxs(
                  "fieldset",
                  {
                    children: [
                      i.jsxs("legend", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Worksheet mapping ",
                          }),
                          de + 1,
                        ],
                      }),
                      i.jsxs("div", {
                        className: "wc-fields",
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Before worksheet ",
                              }),
                              de + 1,
                              i.jsxs("select", {
                                value: mapping.before.sheetId,
                                onChange: (event) =>
                                  X$1(de, {
                                    ...mapping,
                                    before: {
                                      ...mapping.before,
                                      sheetId: event.target.value,
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
                              i.jsx(LocalizedText, {
                                children: "Before range ",
                              }),
                              de + 1,
                              i.jsx("input", {
                                value: mapping.before.range,
                                onChange: (event) =>
                                  X$1(de, {
                                    ...mapping,
                                    before: {
                                      ...mapping.before,
                                      range: event.target.value,
                                    },
                                  }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "After worksheet ",
                              }),
                              de + 1,
                              i.jsxs("select", {
                                value: mapping.after.sheetId,
                                onChange: (event) =>
                                  X$1(de, {
                                    ...mapping,
                                    after: {
                                      ...mapping.after,
                                      sheetId: event.target.value,
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
                              i.jsx(LocalizedText, {
                                children: "After range ",
                              }),
                              de + 1,
                              i.jsx("input", {
                                value: mapping.after.range,
                                onChange: (event) =>
                                  X$1(de, {
                                    ...mapping,
                                    after: {
                                      ...mapping.after,
                                      range: event.target.value,
                                    },
                                  }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Row matching ",
                              }),
                              de + 1,
                              i.jsxs("select", {
                                value: mapping.rowMatch,
                                onChange: (event) =>
                                  X$1(de, {
                                    ...mapping,
                                    rowMatch: event.target.value,
                                    keys:
                                      event.target.value === "keys"
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
                                checked: mapping.headers,
                                onChange: (event) =>
                                  X$1(de, {
                                    ...mapping,
                                    headers: event.target.checked,
                                  }),
                              }),
                              i.jsx(LocalizedText, {
                                children: "First row contains headers ",
                              }),
                              de + 1,
                            ],
                          }),
                        ],
                      }),
                      mapping.rowMatch === "keys" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsx("p", {
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Enter exact header text. Without headers, use relative column labels A, B, C. Values of different types, case or whitespace do not match.",
                              }),
                            }),
                            mapping.keys?.map((key, ce) =>
                              i.jsxs(
                                "div",
                                {
                                  className: "wc-keys",
                                  children: [
                                    i.jsxs("label", {
                                      children: [
                                        i.jsx(LocalizedText, {
                                          children: "Before key ",
                                        }),
                                        de + 1,
                                        ".",
                                        ce + 1,
                                        i.jsx("input", {
                                          value: key.before,
                                          onChange: (event) =>
                                            X$1(de, {
                                              ...mapping,
                                              keys: mapping.keys.map(
                                                (key, we) =>
                                                  we === ce
                                                    ? {
                                                        ...key,
                                                        before:
                                                          event.target.value,
                                                      }
                                                    : key,
                                              ),
                                            }),
                                        }),
                                      ],
                                    }),
                                    i.jsxs("label", {
                                      children: [
                                        i.jsx(LocalizedText, {
                                          children: "After key ",
                                        }),
                                        de + 1,
                                        ".",
                                        ce + 1,
                                        i.jsx("input", {
                                          value: key.after,
                                          onChange: (event) =>
                                            X$1(de, {
                                              ...mapping,
                                              keys: mapping.keys.map(
                                                (key, we) =>
                                                  we === ce
                                                    ? {
                                                        ...key,
                                                        after:
                                                          event.target.value,
                                                      }
                                                    : key,
                                              ),
                                            }),
                                        }),
                                      ],
                                    }),
                                    i.jsxs("button", {
                                      disabled: mapping.keys.length < 2,
                                      onClick: () =>
                                        X$1(de, {
                                          ...mapping,
                                          keys: mapping.keys.filter(
                                            (key, se) => se !== ce,
                                          ),
                                        }),
                                      children: [
                                        i.jsx(LocalizedText, {
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
                              disabled: (mapping.keys?.length ?? 0) >= 8,
                              onClick: () =>
                                X$1(de, {
                                  ...mapping,
                                  keys: [
                                    ...(mapping.keys ?? []),
                                    { before: "", after: "" },
                                  ],
                                }),
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Add key to mapping ",
                                }),
                                de + 1,
                              ],
                            }),
                          ],
                        }),
                      i.jsxs("button", {
                        disabled: mappings.length < 2,
                        onClick: () => {
                          (b(mappings.filter((mapping, ce) => ce !== de)), H());
                        },
                        children: [
                          i.jsx(LocalizedText, { children: "Remove mapping " }),
                          de + 1,
                        ],
                      }),
                    ],
                  },
                  de,
                ),
              ),
            i.jsxs("div", {
              className: "wc-actions",
              children: [
                i.jsx("button", {
                  disabled: !workbook || mappings.length >= 20,
                  onClick: () => {
                    workbook && (b([...mappings, Q(workbook)]), H());
                  },
                  children: i.jsx(LocalizedText, {
                    children: "Add worksheet mapping",
                  }),
                }),
                i.jsx("button", {
                  disabled: !workbook || f,
                  onClick: onClick,
                  children: i.jsx(LocalizedText, {
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
                children: i.jsx(LocalizedText, {
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
                    children: i.jsx(LocalizedText, {
                      children: "Comparison report",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      y.counts.beforeRows,
                      " ",
                      i.jsx(LocalizedText, { children: "before records ·" }),
                      " ",
                      y.counts.afterRows,
                      " ",
                      i.jsx(LocalizedText, { children: "after records ·" }),
                      " ",
                      y.counts.matchedRows,
                      " ",
                      i.jsx(LocalizedText, { children: "matched ·" }),
                      " ",
                      y.counts.insertedRows,
                      " ",
                      i.jsx(LocalizedText, { children: "inserted ·" }),
                      " ",
                      y.counts.deletedRows,
                      " ",
                      i.jsx(LocalizedText, { children: "deleted ·" }),
                      " ",
                      y.counts.unresolvedBeforeRows,
                      "/",
                      y.counts.unresolvedAfterRows,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "unresolved before/after.",
                      }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      y.coverage.selectedCells,
                      " ",
                      i.jsx(LocalizedText, { children: "selected cells." }),
                      " ",
                      y.counts.scannedCells,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "permitted stored cells /",
                      }),
                      " ",
                      y.counts.scannedFormulas,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "formulas inspected statically.",
                      }),
                      " ",
                      y.findings.length,
                      " ",
                      i.jsx(LocalizedText, { children: "findings." }),
                    ],
                  }),
                  i.jsxs("details", {
                    children: [
                      i.jsx("summary", {
                        children: i.jsx(LocalizedText, {
                          children: "Match rules and coverage",
                        }),
                      }),
                      i.jsx("ul", {
                        children: y.coverage.notes.map((note) =>
                          i.jsx("li", { children: note }, note),
                        ),
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Before workbook: ",
                          }),
                          y.before.id,
                          i.jsx(LocalizedText, { children: ", revision" }),
                          " ",
                          y.before.revision,
                          i.jsx(LocalizedText, {
                            children: ". After workbook: ",
                          }),
                          y.after.id,
                          i.jsx(LocalizedText, { children: ", revision " }),
                          y.after.revision,
                          ".",
                        ],
                      }),
                      y.request.mappings.map((mapping, de) =>
                        i.jsxs(
                          "p",
                          {
                            children: [
                              i.jsx(LocalizedText, { children: "Mapping " }),
                              de + 1,
                              ": ",
                              mapping.before.sheetId,
                              "!",
                              mapping.before.range,
                              " →",
                              " ",
                              mapping.after.sheetId,
                              "!",
                              mapping.after.range,
                              "; ",
                              mapping.rowMatch,
                              ";",
                              " ",
                              mapping.headers ? "headers" : "no headers",
                              ";",
                              " ",
                              mapping.keys
                                ?.map((key) => `${key.before} → ${key.after}`)
                                .join(", "),
                              ".",
                            ],
                          },
                          de,
                        ),
                      ),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Comparison finding category",
                      }),
                      i.jsx("select", {
                        value: j,
                        onChange: (event) => {
                          (z(event.target.value), D(0));
                        },
                        children: comparisonFindingKinds.map(
                          (comparisonFindingKind) =>
                            i.jsxs(
                              "option",
                              {
                                value: comparisonFindingKind,
                                children: [
                                  comparisonFindingKind,
                                  " (",
                                  comparisonFindingKind === "all"
                                    ? y.findings.length
                                    : y.findings.filter(
                                        (finding) =>
                                          finding.kind ===
                                          comparisonFindingKind,
                                      ).length,
                                  ")",
                                ],
                              },
                              comparisonFindingKind,
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
                                  children: i.jsx(LocalizedText, {
                                    children: "Finding",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Before address",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Before raw value",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "After address",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "After raw value",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
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
                                        children: formatComparisonValue(
                                          oe.beforeValue,
                                        ),
                                      }),
                                      i.jsx("td", {
                                        children: oe.after
                                          ? `${oe.after.sheetName}!${oe.after.address}`
                                          : "—",
                                      }),
                                      i.jsx("td", {
                                        children: formatComparisonValue(
                                          oe.afterValue,
                                        ),
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
                          children: i.jsx(LocalizedText, {
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
                        children: i.jsx(LocalizedText, {
                          children: "Previous findings",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Page " }),
                          te + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          K,
                          " · ",
                          W.length,
                          " ",
                          i.jsx(LocalizedText, { children: "findings" }),
                        ],
                      }),
                      i.jsx("button", {
                        disabled: te + 1 >= K,
                        onClick: () => D(te + 1),
                        children: i.jsx(LocalizedText, {
                          children: "Next findings",
                        }),
                      }),
                    ],
                  }),
                ],
              }),
            i.jsx("p", {
              className: "wc-limit",
              children: i.jsx(LocalizedText, {
                children:
                  "Limits: 20 mappings, 10,000 selected cells, 1,000 rows and 100 columns per range; 20,000 permitted stored cells inspected. Large or incomplete comparisons return an explicit diagnostic.",
              }),
            }),
          ],
        }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: ae,
            children: i.jsx(LocalizedText, { children: "Close comparison" }),
          }),
        }),
      ],
    }),
  });
}
const comparisonLimits = {
  maxQueryLength: 256,
  maxResults: 1e5,
  maxCalculatedCells: 1e4,
  maxTextCharacters: 1e7,
  maxPatternWork: 3e7,
  maxSpecialCells: 500,
  maxSpecialAreas: 256,
  maxBlankScan: 1e5,
};

function normalizeSearchRequest(workbook, r) {
  const sheetOrThrow = findSheetOrThrow(workbook, r.sheetId),
    selectionAreas = parseSelectionAreas(r.ranges ?? [r.range]),
    o = selectionAreas.map(parseRangeAddress);
  parseRangeAddress(r.range);
  const scope = r.scope ?? "sheet",
    order = r.order ?? "rows";
  return (
    (!["selection", "sheet", "workbook"].includes(scope) ||
      !["rows", "columns"].includes(order)) &&
      throwWithMessage("Choose a valid search scope and order."),
    r.includeHidden !== undefined &&
      typeof r.includeHidden != "boolean" &&
      throwWithMessage("Include hidden must be boolean."),
    o.some(
      (d) =>
        d.lastRow >= sheetOrThrow.rowCount ||
        d.lastCol >= sheetOrThrow.colCount,
    ) && throwWithMessage("The search selection exceeds the worksheet."),
    {
      ...r,
      ranges: selectionAreas,
      scope: scope,
      order: order,
      includeHidden: r.includeHidden ?? false,
    }
  );
}
function normalizeSearchQuery(workbook, r) {
  const searchRequest = normalizeSearchRequest(workbook, r),
    lookIn = r.lookIn ?? "raw";
  ((typeof r.query != "string" ||
    !r.query.length ||
    r.query.length > comparisonLimits.maxQueryLength) &&
    throwWithMessage(
      "Enter 1–256 characters to find. Use Go To Special to select blank cells.",
    ),
    ["raw", "formulas", "calculated", "formatted"].includes(lookIn) ||
      throwWithMessage("Choose a supported Look in option."));
  for (const o of [r.matchCase, r.wholeCell, r.wildcards])
    o !== undefined &&
      typeof o != "boolean" &&
      throwWithMessage("Search switches must be boolean.");
  return {
    ...searchRequest,
    query: r.query,
    lookIn: lookIn,
    matchCase: r.matchCase ?? false,
    wholeCell: r.wholeCell ?? false,
    wildcards: r.wildcards ?? true,
  };
}
const createSearchBudget = () => ({
  text: 0,
  work: 0,
  calculated: 0,
  scanned: 0,
});
function compileSearchPattern(e) {
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
function matchTextWithBudget(e, r, t, searchBudget) {
  const o = Array.from(e),
    a = o.length;
  if (
    ((searchBudget.text += e.length),
    (searchBudget.work += (a + 1) * t.length),
    (searchBudget.text > comparisonLimits.maxTextCharacters ||
      searchBudget.work > comparisonLimits.maxPatternWork) &&
      throwWithMessage(
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
function describeHiddenState(sheetOrThrow) {
  const axisVisibilityIndex = createAxisVisibilityIndex(
      sheetOrThrow.rowCount,
      cn(sheetOrThrow, "row"),
    ),
    t = createAxisVisibilityIndex(
      sheetOrThrow.colCount,
      cn(sheetOrThrow, "column"),
    );
  return (n, o) =>
    sheetOrThrow.hidden
      ? sheetOrThrow.veryHidden
        ? "very hidden sheet"
        : "hidden sheet"
      : axisVisibilityIndex.isHidden(n)
        ? "hidden row"
        : t.isHidden(o)
          ? "hidden column"
          : undefined;
}
const rangeContainsCell = (e, r, t) =>
  e >= t.firstRow && e <= t.lastRow && r >= t.firstCol && r <= t.lastCol;
function compareCellPositions(order, r, t) {
  return order === "columns"
    ? r.col - t.col || r.row - t.row
    : r.row - t.row || r.col - t.col;
}
function searchWorkbookCells(workbook, r) {
  const searchQuery = normalizeSearchQuery(workbook, r),
    searchBudget = createSearchBudget(),
    o = compileSearchPattern(searchQuery),
    results = [],
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    d =
      searchQuery.scope === "workbook"
        ? workbook.sheets
        : [findSheetOrThrow(workbook, searchQuery.sheetId)],
    u = searchQuery.ranges.map(parseRangeAddress),
    l =
      searchQuery.lookIn === "calculated" || searchQuery.lookIn === "formatted";
  for (const sheet of d) {
    if (sheet.hidden && !searchQuery.includeHidden) continue;
    const p = new Set(Object.keys(sheet.cells)),
      m = describeHiddenState(sheet),
      v = [];
    if (l)
      for (const g of workbookWithSettings.getSpillEntries(sheet.id))
        p.add(g.address);
    for (const address of p) {
      ++searchBudget.scanned > workbookLimits.maxCells &&
        throwWithMessage(
          "This search exceeds 100,000 populated cells. Narrow its scope.",
        );
      const { row, col } = _t(address),
        hiddenReason = m(row, col),
        cell = sheet.cells[address];
      if (
        (searchQuery.scope === "selection" &&
          !u.some((O) => rangeContainsCell(row, col, O))) ||
        (hiddenReason && !searchQuery.includeHidden)
      )
        continue;
      const isFormula = isFormulaCellValue(cell);
      if (searchQuery.lookIn === "formulas" && !isFormula) continue;
      let k = cell?.raw ?? null;
      l &&
        (++searchBudget.calculated > comparisonLimits.maxCalculatedCells &&
          throwWithMessage(
            "Calculated searches support at most 10,000 cells. Narrow the selection or use Raw contents.",
          ),
        (k = workbookWithSettings(sheet.id, address)));
      const F =
          searchQuery.lookIn === "formatted"
            ? formatNumberWithPattern(k, cell, workbook.datePolicy?.dateSystem)
            : coerceCellValueToText(k),
        I = matchTextWithBudget(F, searchQuery, o, searchBudget);
      I.length &&
        (v.push({
          sheetId: sheet.id,
          sheetName: sheet.name,
          address: address,
          row,
          col,
          text: F,
          raw: cell?.raw ?? null,
          isFormula: isFormula,
          protected: !!sheet.protection?.enabled,
          hidden: !!hiddenReason,
          ...(hiddenReason ? { hiddenReason: hiddenReason } : {}),
          occurrences: I.length,
        }),
        results.length + v.length > comparisonLimits.maxResults &&
          throwWithMessage("Too many matching cells. Narrow the search."));
    }
    (v.sort((g, b) => compareCellPositions(searchQuery.order, g, b)),
      results.push(...v));
  }
  return eI({
    workbookId: workbook.id,
    revision: workbook.revision,
    snapshotGuard: JSON.stringify(workbook),
    sheetOrder: workbook.sheets.map((sheet) => sheet.id),
    options: { ...searchQuery },
    total: results.length,
    scannedCells: searchBudget.scanned,
    results: results,
  });
}
function normalizeSearchPage(e, r = 0, t = 25) {
  return (
    (!Number.isSafeInteger(r) ||
      r < 0 ||
      !Number.isSafeInteger(t) ||
      t < 1 ||
      t > 200) &&
      throwWithMessage(
        "Search pages require a nonnegative page and 1–200 results per page.",
      ),
    e.results.slice(r * t, (r + 1) * t)
  );
}
function findAdjacentSearchResult(e, r, t = "next") {
  if (
    (["next", "previous"].includes(t) ||
      throwWithMessage("Choose next or previous."),
    !e.total)
  )
    return null;
  const n = r
    ? e.results.findIndex(
        (result) =>
          result.sheetId === r.sheetId &&
          result.address === r.address.toUpperCase(),
      )
    : -1;
  if (n >= 0)
    return e.results[(n + (t === "next" ? 1 : e.total - 1)) % e.total];
  if (r) {
    const o = _t(r.address),
      a = new Map(e.sheetOrder.map((d, u) => [d, u])),
      s = a.get(r.sheetId);
    if (s !== undefined) {
      const d = (u) =>
        a.get(u.sheetId) - s || compareCellPositions(e.options.order, u, o);
      return t === "next"
        ? (e.results.find((result) => d(result) > 0) ?? e.results[0])
        : ([...e.results].reverse().find((u) => d(u) < 0) ?? e.results.at(-1));
    }
  }
  return t === "next" ? e.results[0] : e.results.at(-1);
}
function toComparableCellValue(cell, r) {
  let raw = r;
  r.length
    ? typeof cell.raw == "number" &&
      /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(r) &&
      Number.isFinite(Number(r))
      ? (raw = Number(r))
      : typeof cell.raw == "boolean" &&
        /^(true|false)$/i.test(r) &&
        (raw = r.toLowerCase() === "true")
    : (raw = null);
  const n = { ...cell, raw: raw };
  return (
    typeof raw == "string" && raw.startsWith("=") && !isFormulaCellValue(cell)
      ? (n.literal = true)
      : (typeof raw == "string" && raw.startsWith("=")) || delete n.literal,
    normalizeCellRecordInput(n)
  );
}

export {
  compareCellPositions,
  compileSearchPattern,
  searchWorkbookCells,
  normalizeSearchPage,
  findAdjacentSearchResult,
  normalizeSearchRequest,
  describeHiddenState,
  createSearchBudget,
  toComparableCellValue,
  comparisonLimits,
  matchTextWithBudget,
  WorkbookComparisonDialog,
  rangeContainsCell,
};
