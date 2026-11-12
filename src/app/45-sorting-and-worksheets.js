import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import {
  X,
  Plus,
  ArrowDown,
  ArrowUp,
  FileSpreadsheet,
  Copy,
  EyeOff,
  Eye as Eye$1,
  LockKeyhole,
  Trash2,
  LockKeyholeOpen,
} from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  describeSheetDeletionImpacts,
  addWorksheet,
  setWorksheetTabColor,
  unprotectWorkbookStructure,
  findSheetOrThrow,
  renameWorksheet,
  setWorksheetVisibility,
  deleteWorksheet,
  protectWorkbookStructure,
  _t,
  cn,
  parseRangeAddress,
  assertSheetUnprotected,
  summarizeWorksheets,
  normalizeWorkbookInput,
  columnIndexToLetters,
  checkDataValidationBatch,
  moveWorksheet,
  duplicateWorksheet,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import { applyStructuralEdit } from "./27-validation-and-structural-edits.js";
import {
  assertSortWorkbookAuthorized,
  buildAllowedSheetKey,
  buildSortPlan,
  MW,
  N3,
  stableSortPolicy,
  resolveSortSourceRange,
  assertSortSafeFormulas,
  isNonBlankCellEntry,
  normalizeSortSpecification,
  sortEngineLimits,
} from "./44-clean-and-reshape.js";

function buildSortPreview(workbook, r, t = {}) {
  const n = assertSortWorkbookAuthorized(workbook, r.sheetId, t),
    sortSpecification = normalizeSortSpecification(workbook, r),
    sortSourceRange = resolveSortSourceRange(
      workbook,
      sortSpecification.sheetId,
      sortSpecification.range,
      sortSpecification.headers,
    );
  if (sortSourceRange.expansionRequired && !sortSpecification.expandSelection)
    throw Error(
      `This selection is part of populated records. Expand the selection to ${sortSourceRange.expandedRange}, or cancel; partial-record sorting is blocked.`,
    );
  const rangeAddress = parseRangeAddress(sortSourceRange.expandedRange),
    d = sortSourceRange.lastRow - sortSourceRange.firstDataRow + 1;
  if (d < 2)
    throw Error("Choose at least two data rows. Check the header setting.");
  if (
    d > sortEngineLimits.rows ||
    d * (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      sortEngineLimits.cells
  )
    throw Error(
      "Sort supports at most 10,000 rows and 100,000 record cells per reviewed operation.",
    );
  const u = cn(n, "row"),
    hiddenRows = Array.from(
      { length: d },
      (B, _) => sortSourceRange.firstDataRow + _,
    ).filter((B) => u.some((_) => B >= _.start && B <= _.end)).length;
  if (hiddenRows && !sortSpecification.includeHidden)
    throw Error(
      `The source contains ${hiddenRows} hidden rows. Include hidden rows explicitly or cancel so records stay together.`,
    );
  assertSortSafeFormulas(workbook, d, t);
  const sortPlan = buildSortPlan(workbook, sortSpecification, sortSourceRange),
    keys = JSON.stringify(sortPlan),
    f = new Intl.Collator("en-US", {
      usage: "sort",
      sensitivity: sortSpecification.caseSensitive ? "variant" : "accent",
      numeric: false,
    }),
    p = (B) =>
      sortSpecification.caseSensitive
        ? B.normalize("NFC")
        : B.normalize("NFC").toLocaleLowerCase("en-US"),
    m = sortSpecification.keys.map(
      (key) => new Map(key.customOrder?.map((C, T) => [p(C), T])),
    ),
    v = (B) =>
      B.compare === null || B.compare === ""
        ? 4
        : B.error
          ? 3
          : typeof B.compare == "number"
            ? 0
            : typeof B.compare == "string"
              ? 1
              : 2,
    g = [...sortPlan].sort((B, _) => {
      for (let C = 0; C < sortSpecification.keys.length; C++) {
        const T = B.keys[C],
          N = _.keys[C],
          R = v(T),
          L = v(N),
          $ = sortSpecification.keys[C].direction === "ascending" ? 1 : -1;
        if (R === 4 || L === 4) {
          if (R !== L) return R === 4 ? 1 : -1;
          continue;
        }
        if (m[C].size) {
          const G = m[C].get(p(String(T.compare))),
            H = m[C].get(p(String(N.compare)));
          if (G !== undefined || H !== undefined) {
            if (G === undefined) return 1;
            if (H === undefined) return -1;
            if (G !== H) return (G - H) * $;
          }
        }
        const U =
          R !== L
            ? R - L
            : typeof T.compare == "number" && typeof N.compare == "number"
              ? T.compare - N.compare
              : typeof T.compare == "boolean" && typeof N.compare == "boolean"
                ? Number(T.compare) - Number(N.compare)
                : f.compare(String(T.compare), String(N.compare));
        if (U) return U * $;
      }
      return B.row - _.row;
    }),
    permutation = Array(d).fill(0);
  g.forEach(
    (B, _) =>
      (permutation[B.row - sortSourceRange.firstDataRow] =
        sortSourceRange.firstDataRow + _),
  );
  const movedRows = g.filter(
      (B, _) => B.row !== sortSourceRange.firstDataRow + _,
    ).length,
    w = {
      axis: "row",
      index: sortSourceRange.firstDataRow,
      count: d,
      remove: false,
      moveBefore: sortSourceRange.firstDataRow,
      permutation: permutation,
    };
  let candidate = movedRows
    ? applyStructuralEdit(workbook, n.id, w)
    : structuredClone(workbook);
  const A = [],
    formulaChanges = [],
    metadataChanges = [],
    I = (B) =>
      B >= sortSourceRange.firstDataRow && B <= sortSourceRange.lastRow
        ? permutation[B - sortSourceRange.firstDataRow]
        : B;
  for (const sheet of workbook.sheets) {
    const sheetOrThrow = findSheetOrThrow(candidate, sheet.id);
    for (const [source, cell] of Object.entries(sheet.cells))
      if (N3(cell)) {
        const N = _t(source),
          destination =
            sheet.id === n.id ? formatCellAddress(I(N.row), N.col) : source,
          L = sheetOrThrow.cells[destination]?.raw;
        L !== cell.raw &&
          formulaChanges.push({
            sheet: sheet.name,
            source: source,
            destination: destination,
            before: String(cell.raw),
            after: String(L),
          });
      }
    for (const address of new Set([
      ...Object.keys(sheet.cells),
      ...Object.keys(sheetOrThrow.cells),
    ]))
      JSON.stringify(sheet.cells[address]) !==
        JSON.stringify(sheetOrThrow.cells[address]) &&
        A.push({
          sheetId: sheet.id,
          address: address,
          cell: sheetOrThrow.cells[address] ?? { raw: null },
        });
    for (const field of Object.keys(sheetOrThrow).filter((T) => T !== "cells"))
      JSON.stringify(sheet[field]) !== JSON.stringify(sheetOrThrow[field]) &&
        metadataChanges.push({ sheet: sheet.name, field: field });
  }
  if (
    (JSON.stringify(workbook.definedNames) !==
      JSON.stringify(candidate.definedNames) &&
      metadataChanges.push({ sheet: "Workbook", field: "definedNames" }),
    JSON.stringify(workbook.formulaWatches) !==
      JSON.stringify(candidate.formulaWatches) &&
      metadataChanges.push({ sheet: "Workbook", field: "formulaWatches" }),
    A.length > sortEngineLimits.cells)
  )
    throw Error(
      "The sort would change more than 100,000 cells including dependent references.",
    );
  JSON.stringify(workbook.reviewThreads) !==
    JSON.stringify(candidate.reviewThreads) &&
    metadataChanges.push({ sheet: "Workbook", field: "reviewThreads" });
  for (const sheet of workbook.sheets)
    sheet.id !== n.id &&
      JSON.stringify(sheet) !==
        JSON.stringify(findSheetOrThrow(candidate, sheet.id)) &&
      assertSheetUnprotected(sheet);
  if ((checkDataValidationBatch(candidate, A), movedRows)) {
    const sheetOrThrow = findSheetOrThrow(candidate, n.id);
    ((candidate.rowOrigins = [...(candidate.rowOrigins ?? [])]),
      (sheetOrThrow.rowOriginIds = { ...sheetOrThrow.rowOriginIds }));
    const _ = new Set(
      Object.entries(n.cells)
        .filter(([, C]) => isNonBlankCellEntry(C))
        .map(([C]) => _t(C).row),
    );
    for (const C of g)
      if (_.has(C.row) && !n.rowOriginIds?.[C.row]) {
        if (candidate.rowOrigins.length >= 1e5)
          throw Error("Source row identity capacity has been reached.");
        const T = crypto.randomUUID();
        ((sheetOrThrow.rowOriginIds[I(C.row)] = T),
          candidate.rowOrigins.push({
            id: T,
            sourceId: workbook.id,
            sourceName: workbook.name,
            sourceSheet: n.name,
            originalRow: C.row + 1,
            kind: "snapshot",
          }));
      }
    candidate = normalizeWorkbookInput(candidate);
  }
  const O = new Map(
      candidate.rowOrigins?.map((rowOrigin) => [rowOrigin.id, rowOrigin]),
    ),
    j = {
      request: sortSpecification,
      selection: sortSourceRange,
      snapshot: JSON.stringify(workbook),
      rows: g.map((B, _) => {
        const sourceId = findSheetOrThrow(candidate, n.id).rowOriginIds?.[
            sortSourceRange.firstDataRow + _
          ],
          T = sourceId ? O.get(sourceId) : undefined;
        return {
          sourceRow: B.row + 1,
          destinationRow: sortSourceRange.firstDataRow + _ + 1,
          ...(sourceId
            ? {
                sourceId: sourceId,
                sourceName: T?.sourceName,
                originalRow: T?.originalRow,
              }
            : {}),
          keys: B.keys.map((key) => key.value),
        };
      }),
      movedRows: movedRows,
      changedCells: A.length,
      formulaChanges: formulaChanges,
      metadataChanges: metadataChanges,
      hiddenRows: hiddenRows,
      policy: stableSortPolicy,
    };
  return (
    MW.set(j, {
      signature: JSON.stringify(j),
      candidate: candidate,
      keys: keys,
      allowed: buildAllowedSheetKey(workbook, t),
    }),
    j
  );
}
function assertSortPreviewCurrent(workbook, r, t = {}) {
  const n = MW.get(r);
  if (!n || n.signature !== JSON.stringify(r))
    throw Error(
      "This sort preview changed or was restored from another session. Preview again.",
    );
  if (
    (assertSortWorkbookAuthorized(workbook, r.request.sheetId, t),
    n.allowed !== buildAllowedSheetKey(workbook, t) ||
      JSON.stringify(workbook) !== r.snapshot)
  )
    throw Error(
      "The workbook or host permissions changed after the sort preview. Preview again.",
    );
  if (
    n.keys !== JSON.stringify(buildSortPlan(workbook, r.request, r.selection))
  )
    throw Error(
      "Calculated sort keys changed after the preview. Preview and review the new order.",
    );
  return r.movedRows ? structuredClone(n.candidate) : workbook;
}
const buildSortKeySpec = (column) => ({
    column: column,
    direction: "ascending",
    type: "auto",
    customText: "",
    useCustom: false,
  }),
  expandSortSelectionToRecords = (e) => {
    try {
      const firstRow = parseRangeAddress(
        resolveSortSourceRange(e.workbook, e.sheetId, e.range).expandedRange,
      ).firstRow;
      return (
        firstRow === 0 ||
        !!e.workbook.sheets
          .find((sheet) => sheet.id === e.sheetId)
          ?.tables?.some(
            (table) => parseRangeAddress(table.range).firstRow === firstRow,
          )
      );
    } catch {
      return false;
    }
  };
function CustomSortDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(null),
    [range, d] = ne.useState(e.range),
    [headers, l] = ne.useState(() => expandSortSelectionToRecords(e)),
    [c, h] = ne.useState([
      buildSortKeySpec(parseRangeAddress(e.range).firstCol),
    ]),
    [f, p] = ne.useState(false),
    [m, v] = ne.useState(false),
    [g, b] = ne.useState(false),
    [y, w] = ne.useState(null),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(0),
    sheet = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    z = e.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
    V = z.includes(e.sheetId),
    D = JSON.stringify([
      e.workbook,
      e.sheetId,
      e.range,
      e.readOnly,
      [...z].sort(),
    ]),
    B = !!y && E === D && !e.readOnly;
  let _,
    C = "";
  try {
    if (!V) throw Error("The source is outside the host-approved worksheets.");
    _ = resolveSortSourceRange(e.workbook, e.sheetId, range, headers);
  } catch (error) {
    C = error.message;
  }
  const T = _ ? parseRangeAddress(_.expandedRange) : undefined,
    N = T
      ? Array.from(
          { length: T.lastCol - T.firstCol + 1 },
          (X, J) => T.firstCol + J,
        )
      : [],
    R = () => {
      (w(null), F(""), O(0));
    },
    L = (X, J) => {
      (R(), h(c.map((q, W) => (W === X ? { ...q, ...J } : q))));
    },
    $ = (X, J) => {
      R();
      const q = [...c];
      (([q[X], q[X + J]] = [q[X + J], q[X]]), h(q));
    },
    onClick = () => {
      try {
        const sortPreview = buildSortPreview(
          e.workbook,
          {
            sheetId: e.sheetId,
            range: range,
            headers: headers,
            keys: c.map(({ customText, useCustom, ...W }) => ({
              ...W,
              ...(useCustom ? { customOrder: customText.split(/\r?\n/) } : {}),
            })),
            expandSelection: f,
            includeHidden: m,
            caseSensitive: g,
          },
          { readOnly: e.readOnly, allowedSheetIds: e.allowedSheetIds },
        );
        (w(sortPreview), A(D), O(0), F(""));
      } catch (error) {
        (w(null), F(error.message));
      }
    },
    G = () => {
      try {
        if (!y || !B)
          throw Error("The workbook or selection changed. Preview again.");
        const X = assertSortPreviewCurrent(e.workbook, y, {
          readOnly: e.readOnly,
          allowedSheetIds: e.allowedSheetIds,
        });
        (e.onWorkbookChange(
          X,
          `Sort ${sheet?.name}: ${y.movedRows} records moved`,
        ),
          e.onClose());
      } catch (error) {
        (w(null), F(error.message));
      }
    };
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      o.current?.querySelector('input[data-ui-label="Source range"]')?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      var X;
      y &&
        (((X = a.current) == null ? undefined : X.scrollIntoView)?.call(X, {
          block: "start",
        }),
        a.current?.focus({ preventScroll: true }));
    }, [y]));
  const H = Math.max(
      1,
      Math.ceil((y?.rows.length ?? 0) / sortEngineLimits.pageSize),
    ),
    Q = Math.min(I, H - 1);
  return i.jsx("div", {
    className: "dup-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      className: "dup-dialog sort-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      ref: o,
      onKeyDown: (event) => {
        var q, W;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), e.onClose()),
          event.key === "Tab")
        ) {
          const K = [
            ...(o.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
            ) ?? []),
          ].filter(
            (te) =>
              !te.closest("details:not([open])") || te.tagName === "SUMMARY",
          );
          event.shiftKey && document.activeElement === K[0]
            ? (event.preventDefault(), (q = K.at(-1)) == null || q.focus())
            : !event.shiftKey &&
              document.activeElement === K.at(-1) &&
              (event.preventDefault(), (W = K[0]) == null || W.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
                  children: i.jsx(LocalizedText, { children: "Custom sort" }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children: "Review the new record order before applying it.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close custom sort"),
              onClick: e.onClose,
              lang: language("Close custom sort"),
              "data-ui-label": "Close custom sort",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "dup-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Worksheet" }),
                i.jsx("strong", {
                  children: V ? sheet?.name : "Outside current context",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Source range"),
                  value: range,
                  onChange: (event) => {
                    (R(), p(false), d(event.target.value));
                  },
                  lang: language("Source range"),
                  "data-ui-label": "Source range",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "dup-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: headers,
                  onChange: (event) => {
                    (R(), l(event.target.checked));
                  },
                }),
                i.jsx(LocalizedText, { children: "My data has headers" }),
              ],
            }),
          ],
        }),
        C
          ? i.jsx("p", { role: "alert", children: C })
          : i.jsxs("div", {
              className: "sort-scope",
              children: [
                i.jsxs("strong", {
                  children: [
                    i.jsx(LocalizedText, { children: "Complete records: " }),
                    _?.expandedRange,
                  ],
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Every populated column in these worksheet rows moves together, including fields beside the selection. Headers stay in place. Applying also clears temporary view sorting and search.",
                  }),
                }),
                _?.expansionRequired &&
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Expand selection",
                        }),
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "radio",
                            name: `${n}-expand`,
                            checked: !f,
                            onChange: () => {
                              (R(), p(false));
                            },
                          }),
                          i.jsx(LocalizedText, {
                            children:
                              "Keep current selection (sorting blocked to protect records)",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "radio",
                            name: `${n}-expand`,
                            checked: f,
                            onChange: () => {
                              (R(), p(true));
                            },
                          }),
                          i.jsx(LocalizedText, { children: "Expand to " }),
                          _.expandedRange,
                        ],
                      }),
                    ],
                  }),
              ],
            }),
        i.jsxs("div", {
          className: "sort-level-heading",
          children: [
            i.jsx("h3", {
              children: i.jsx(LocalizedText, { children: "Sort levels" }),
            }),
            i.jsxs("button", {
              disabled:
                c.length >= sortEngineLimits.keys || c.length >= N.length,
              onClick: () => {
                (R(),
                  h([
                    ...c,
                    buildSortKeySpec(
                      N.find((X) => !c.some((J) => J.column === X)),
                    ),
                  ]));
              },
              children: [
                i.jsx(Plus, { size: 15 }),
                i.jsx(LocalizedText, { children: "Add level" }),
              ],
            }),
          ],
        }),
        i.jsx("ol", {
          className: "sort-levels",
          children: c.map((X$1, J) =>
            i.jsxs(
              "li",
              {
                children: [
                  i.jsx("span", {
                    className: "sort-level-number",
                    children: J + 1,
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Column" }),
                      N.length > 500
                        ? i.jsx(LocalizedNumberField, {
                            "aria-label": `Column ${J + 1}`,
                            min: T.firstCol + 1,
                            max: T.lastCol + 1,
                            value: X$1.column + 1,
                            onValueChange: (q) =>
                              L(J, { column: Number(q) - 1 }),
                          })
                        : i.jsx("select", {
                            "aria-label": `Column ${J + 1}`,
                            value: X$1.column,
                            onChange: (event) =>
                              L(J, { column: Number(event.target.value) }),
                            children: N.map((q) =>
                              i.jsxs(
                                "option",
                                {
                                  value: q,
                                  children: [
                                    columnIndexToLetters(q),
                                    headers && T
                                      ? ` · ${String(sheet?.cells[`${columnIndexToLetters(q)}${T.firstRow + 1}`]?.raw ?? "(blank header)")}`
                                      : "",
                                  ],
                                },
                                q,
                              ),
                            ),
                          }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Order" }),
                      i.jsxs("select", {
                        "aria-label": `Order ${J + 1}`,
                        value: X$1.direction,
                        onChange: (event) =>
                          L(J, { direction: event.target.value }),
                        children: [
                          i.jsx("option", {
                            value: "ascending",
                            children: "Ascending",
                          }),
                          i.jsx("option", {
                            value: "descending",
                            children: "Descending",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Interpret as" }),
                      i.jsxs("select", {
                        "aria-label": `Interpretation ${J + 1}`,
                        value: X$1.type,
                        onChange: (event) =>
                          L(J, { type: event.target.value, useCustom: false }),
                        children: [
                          i.jsx("option", {
                            value: "auto",
                            children: "Automatic",
                          }),
                          i.jsx("option", {
                            value: "number",
                            children: "Number",
                          }),
                          i.jsx("option", {
                            value: "date",
                            children: "Date / time",
                          }),
                          i.jsx("option", { value: "text", children: "Text" }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "sort-level-actions",
                    children: [
                      i.jsx("button", {
                        "aria-label": `Move level ${J + 1} up`,
                        disabled: J === 0,
                        onClick: () => $(J, -1),
                        children: i.jsx(ArrowUp, { size: 15 }),
                      }),
                      i.jsx("button", {
                        "aria-label": `Move level ${J + 1} down`,
                        disabled: J === c.length - 1,
                        onClick: () => $(J, 1),
                        children: i.jsx(ArrowDown, { size: 15 }),
                      }),
                      i.jsx("button", {
                        "aria-label": `Remove level ${J + 1}`,
                        disabled: c.length === 1,
                        onClick: () => {
                          (R(), h(c.filter((q, W) => W !== J)));
                        },
                        children: i.jsx(X, { size: 15 }),
                      }),
                    ],
                  }),
                  X$1.type === "text" &&
                    i.jsxs("div", {
                      className: "sort-custom",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx("input", {
                              type: "checkbox",
                              checked: X$1.useCustom,
                              onChange: (event) =>
                                L(J, { useCustom: event.target.checked }),
                            }),
                            i.jsx(LocalizedText, {
                              children: "Use a custom order for level ",
                            }),
                            J + 1,
                          ],
                        }),
                        X$1.useCustom &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "One value per line",
                              }),
                              i.jsx("textarea", {
                                "aria-label": `Custom order ${J + 1}`,
                                rows: 3,
                                placeholder: `High
Medium
Low`,
                                value: X$1.customText,
                                onChange: (event) =>
                                  L(J, { customText: event.target.value }),
                              }),
                              i.jsx("small", {
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Unlisted values follow the list; blanks stay last. Descending reverses the list.",
                                }),
                              }),
                            ],
                          }),
                      ],
                    }),
                ],
              },
              J,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "sort-options",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: g,
                  onChange: (event) => {
                    (R(), b(event.target.checked));
                  },
                }),
                i.jsx(LocalizedText, { children: "Case-sensitive text" }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: m,
                  onChange: (event) => {
                    (R(), v(event.target.checked));
                  },
                }),
                i.jsx(LocalizedText, { children: "Include hidden rows" }),
              ],
            }),
          ],
        }),
        i.jsx("p", {
          className: "sort-policy",
          children: i.jsx(LocalizedText, {
            children:
              "Ties keep their original row order. Blanks stay last. Automatic keeps numbers and numeric text distinct. Limit: 10,000 rows and 100,000 cells.",
          }),
        }),
        i.jsxs("details", {
          children: [
            i.jsx("summary", {
              children: i.jsx(LocalizedText, {
                children: "Comparison and reference policy",
              }),
            }),
            Object.values(stableSortPolicy).map((X) =>
              i.jsx("p", { children: X }, X),
            ),
          ],
        }),
        e.readOnly &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children:
                "This host is read-only. Sorting cannot change records.",
            }),
          }),
        k && i.jsx("p", { className: "dup-error", role: "alert", children: k }),
        y &&
          !B &&
          i.jsx("p", {
            role: "alert",
            children: i.jsx(LocalizedText, {
              children:
                "The workbook, selection or permissions changed. Preview again.",
            }),
          }),
        y &&
          B &&
          i.jsxs("div", {
            className: "dup-preview",
            "aria-label": text("Sort preview"),
            ref: a,
            tabIndex: -1,
            lang: language("Sort preview"),
            "data-ui-label": "Sort preview",
            children: [
              i.jsxs("h3", {
                children: [
                  y.movedRows,
                  " ",
                  i.jsx(LocalizedText, { children: "records move · " }),
                  y.changedCells,
                  " ",
                  i.jsx(LocalizedText, { children: "cells change" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  y.rows.length,
                  " ",
                  i.jsx(LocalizedText, { children: "records reviewed · " }),
                  y.hiddenRows,
                  " ",
                  i.jsx(LocalizedText, { children: "hidden rows · " }),
                  y.formulaChanges.length,
                  " ",
                  i.jsx(LocalizedText, { children: "formula rewrites ·" }),
                  " ",
                  y.metadataChanges.length,
                  " ",
                  i.jsx(LocalizedText, { children: "metadata updates" }),
                ],
              }),
              i.jsx("div", {
                className: "dup-scroll",
                children: i.jsxs("table", {
                  children: [
                    i.jsx("thead", {
                      children: i.jsxs("tr", {
                        children: [
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Source row",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Destination row",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Sort values",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Source identity",
                            }),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: y.rows
                        .slice(
                          Q * sortEngineLimits.pageSize,
                          (Q + 1) * sortEngineLimits.pageSize,
                        )
                        .map((X) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: X.sourceRow }),
                                i.jsx("td", { children: X.destinationRow }),
                                i.jsx("td", {
                                  children: X.keys
                                    .map((key) =>
                                      key === null || key === ""
                                        ? "(blank)"
                                        : String(key),
                                    )
                                    .join(" · "),
                                }),
                                i.jsxs("td", {
                                  children: [
                                    i.jsxs("span", {
                                      children: [
                                        X.sourceName ?? "Unchanged source",
                                        X.originalRow
                                          ? ` · row ${X.originalRow}`
                                          : "",
                                      ],
                                    }),
                                    X.sourceId &&
                                      i.jsx("small", {
                                        title: X.sourceId,
                                        children: X.sourceId,
                                      }),
                                  ],
                                }),
                              ],
                            },
                            X.sourceRow,
                          ),
                        ),
                    }),
                  ],
                }),
              }),
              i.jsxs("nav", {
                "aria-label": text("Sort preview pages"),
                lang: language("Sort preview pages"),
                "data-ui-label": "Sort preview pages",
                children: [
                  i.jsx("button", {
                    disabled: Q === 0,
                    onClick: () => O(Q - 1),
                    children: i.jsx(LocalizedText, { children: "Previous" }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      Q + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      H,
                    ],
                  }),
                  i.jsx("button", {
                    disabled: Q === H - 1,
                    onClick: () => O(Q + 1),
                    children: i.jsx(LocalizedText, { children: "Next" }),
                  }),
                ],
              }),
              y.formulaChanges.length > 0 &&
                i.jsxs("details", {
                  children: [
                    i.jsxs("summary", {
                      children: [
                        i.jsx(LocalizedText, { children: "Review " }),
                        y.formulaChanges.length,
                        " ",
                        i.jsx(LocalizedText, { children: "formula changes" }),
                      ],
                    }),
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "First 50 changes shown; the SDK preview contains the complete list.",
                      }),
                    }),
                    y.formulaChanges.slice(0, 50).map((X, J) =>
                      i.jsxs(
                        "p",
                        {
                          children: [
                            i.jsxs("strong", {
                              children: [
                                X.sheet,
                                "!",
                                X.source,
                                " → ",
                                X.destination,
                              ],
                            }),
                            i.jsx("br", {}),
                            i.jsxs("code", {
                              children: [X.before, " → ", X.after],
                            }),
                          ],
                        },
                        J,
                      ),
                    ),
                  ],
                }),
              y.metadataChanges.length > 0 &&
                i.jsxs("details", {
                  children: [
                    i.jsx("summary", {
                      children: i.jsx(LocalizedText, {
                        children: "Review metadata updates",
                      }),
                    }),
                    y.metadataChanges.map((metadataChange, J) =>
                      i.jsxs(
                        "p",
                        {
                          children: [
                            metadataChange.sheet,
                            ": ",
                            metadataChange.field,
                          ],
                        },
                        J,
                      ),
                    ),
                  ],
                }),
            ],
          }),
        i.jsxs("footer", {
          className: "dup-actions",
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsxs("div", {
              children: [
                i.jsx("button", {
                  disabled: e.readOnly || !!C,
                  onClick: onClick,
                  children: i.jsx(LocalizedText, { children: "Preview sort" }),
                }),
                i.jsx("button", {
                  className: "dup-primary",
                  disabled: !B || !y?.movedRows,
                  onClick: G,
                  children: i.jsx(LocalizedText, { children: "Apply sort" }),
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  });
}
const worksheetVisibilityState = (e) =>
    e.veryHidden ? "veryHidden" : e.hidden ? "hidden" : "visible",
  worksheetVisibilityLabel = (e) =>
    e === "veryHidden" ? "Very hidden" : e === "hidden" ? "Hidden" : "Visible",
  describeWorksheetError = (e) =>
    e instanceof Error
      ? e.message
      : "The worksheet operation could not be completed.";
function WorksheetTabBar(e) {
  const { text, language } = useWorkbookMessages(),
    [n, o] = ne.useState(e.sheetId),
    a =
      e.workbook.sheets.find((sheet) => sheet.id === n) ?? e.workbook.sheets[0],
    [s, d] = ne.useState(a.name),
    [u, l] = ne.useState(worksheetVisibilityState(a)),
    [tabColor, h] = ne.useState(a.tabColor ?? "#107c41"),
    [f, p] = ne.useState(""),
    [m, v] = ne.useState(""),
    [g, b] = ne.useState(""),
    [y, w] = ne.useState(false),
    [E, A] = ne.useState(null),
    [checked, F] = ne.useState(false),
    [I, O] = ne.useState(0),
    j = ne.useRef(null),
    z = ne.useRef(null),
    V = ne.useRef(e),
    D = ne.useRef(true);
  V.current = e;
  const B = ne.useId(),
    _ = !!e.workbook.structureProtection?.enabled,
    T = ne
      .useMemo(() => summarizeWorksheets(e.workbook), [e.workbook])
      .find((ae) => ae.sheetId === a.id),
    disabled = e.readOnly || _ || y,
    R = e.workbook.sheets.findIndex((sheet) => sheet.id === a.id),
    L =
      !a.hidden &&
      e.workbook.sheets.filter((sheet) => !sheet.hidden).length === 1,
    $ = !disabled && e.workbook.sheets.length > 1 && !L;
  (ne.useEffect(() => {
    D.current = true;
    const activeElement = document.activeElement;
    return (
      (z.current ?? j.current)?.focus(),
      () => {
        ((D.current = false),
          activeElement instanceof HTMLElement &&
            activeElement.isConnected &&
            activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      (o(a.id),
        d(a.name),
        l(worksheetVisibilityState(a)),
        h(a.tabColor ?? "#107c41"));
    }, [a.id, a.name, a.hidden, a.veryHidden, a.tabColor]));
  function U(ae) {
    (o(ae), A(null), F(false), v(""), b(""));
  }
  function G(ae, oe, de = false) {
    (v(""), b(""));
    try {
      const he = V.current;
      if (he.readOnly) throw new Error("This workbook is read-only.");
      const workbook = oe(he.workbook);
      if (workbook === he.workbook) return;
      (he.onWorkbookChange(workbook, ae), A(null), F(false), O(0), b(ae));
      const le = de
        ? workbook.sheets.find(
            (sheet) => !he.workbook.sheets.some((ge) => ge.id === sheet.id),
          )
        : undefined;
      (le
        ? (o(le.id), le.hidden || he.onSelectSheet(le.id))
        : workbook.sheets.some((sheet) => sheet.id === a.id) ||
          o(workbook.sheets.find((sheet) => !sheet.hidden).id),
        !le &&
          !workbook.sheets.some(
            (sheet) => sheet.id === he.sheetId && !sheet.hidden,
          ) &&
          he.onSelectSheet(workbook.sheets.find((sheet) => !sheet.hidden).id));
    } catch (error) {
      v(describeWorksheetError(error));
    }
  }
  function onClick() {
    const ae = new Set(
      e.workbook.sheets.map((sheet) => sheet.name.toLowerCase()),
    );
    let oe = 1;
    for (; ae.has(`sheet${oe}`);) oe++;
    G(
      "Add worksheet",
      (activeWorkbook) => addWorksheet(activeWorkbook, `Sheet${oe}`),
      true,
    );
  }
  function Q() {
    (v(""), b(""), O(0));
    try {
      (A({
        workbook: e.workbook,
        sheetId: a.id,
        name: a.name,
        dependencies: describeSheetDeletionImpacts(e.workbook, a.id),
      }),
        F(false));
    } catch (error) {
      v(describeWorksheetError(error));
    }
  }
  function X$1() {
    if (!(!E || !checked)) {
      if (E.workbook !== V.current.workbook) {
        v("The workbook changed. Review current dependencies before deleting.");
        return;
      }
      G(`Delete worksheet: ${E.name}`, (workbook) =>
        deleteWorksheet(workbook, E.sheetId, true),
      );
    }
  }
  async function J() {
    const ae = V.current,
      oe = !!ae.workbook.structureProtection?.enabled;
    (v(""), b(""), w(true));
    try {
      if (ae.readOnly) throw new Error("This workbook is read-only.");
      const he = await (oe
        ? unprotectWorkbookStructure(ae.workbook, f)
        : protectWorkbookStructure(ae.workbook, f));
      if (!D.current) return;
      const Se = V.current;
      if (Se.readOnly)
        throw new Error(
          "The host changed this workbook to read-only. No structure change was applied.",
        );
      if (Se.workbook !== ae.workbook)
        throw new Error(
          "The workbook changed while checking the password. Try again with the current workbook.",
        );
      (Se.onWorkbookChange(
        he,
        oe ? "Unprotect workbook structure" : "Protect workbook structure",
      ),
        p(""),
        A(null),
        b(
          oe
            ? "Workbook structure is unlocked."
            : "Workbook structure is protected.",
        ));
    } catch (error) {
      D.current && v(describeWorksheetError(error));
    } finally {
      D.current && w(false);
    }
  }
  const q = !!(E && E.workbook !== e.workbook),
    W = E?.dependencies.some((dependency) =>
      /unsupported/i.test(dependency.kind),
    );
  return i.jsx("div", {
    className: "sm-backdrop",
    onPointerDown: (event) => {
      event.target === event.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      ref: j,
      className: "sm-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": B,
      tabIndex: -1,
      onKeyDown: (event) => {
        var de;
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(), E ? (A(null), F(false)) : e.onClose()),
          event.key === "Tab")
        ) {
          const he = Array.from(
              j.current?.querySelectorAll(
                'button,input,select,textarea,[tabindex="0"]',
              ) ?? [],
            ).filter(
              (ge) =>
                !ge.matches(":disabled") &&
                !ge.closest("[hidden]") &&
                ge.tabIndex >= 0,
            ),
            Se = he[0],
            le = he.at(-1),
            activeElement = document.activeElement;
          he.length
            ? event.shiftKey &&
              (activeElement === Se || !he.includes(activeElement))
              ? (event.preventDefault(), le == null || le.focus())
              : !event.shiftKey &&
                (activeElement === le || !he.includes(activeElement)) &&
                (event.preventDefault(), Se == null || Se.focus())
            : (event.preventDefault(), (de = j.current) == null || de.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "sm-header",
          children: [
            i.jsxs("div", {
              children: [
                i.jsx(FileSpreadsheet, { size: 19 }),
                i.jsx("h2", {
                  id: B,
                  children: i.jsx(LocalizedText, {
                    children: "Manage Worksheets",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": text("Close worksheet manager"),
              onClick: e.onClose,
              lang: language("Close worksheet manager"),
              "data-ui-label": "Close worksheet manager",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "sm-body",
          children: [
            i.jsxs("aside", {
              className: "sm-sheet-list",
              "aria-label": text("Workbook worksheets"),
              lang: language("Workbook worksheets"),
              "data-ui-label": "Workbook worksheets",
              children: [
                i.jsxs("div", {
                  className: "sm-list-heading",
                  children: [
                    i.jsxs("span", {
                      children: [
                        e.workbook.sheets.length,
                        " ",
                        i.jsx(LocalizedText, { children: "worksheets" }),
                      ],
                    }),
                    i.jsx("button", {
                      type: "button",
                      "aria-label": text("Add worksheet"),
                      title: _
                        ? "Unprotect workbook structure to add worksheets"
                        : "Add worksheet",
                      disabled: disabled,
                      onClick: onClick,
                      lang: language("Add worksheet"),
                      "data-ui-label": "Add worksheet",
                      children: i.jsx(Plus, { size: 16 }),
                    }),
                  ],
                }),
                i.jsx("ul", {
                  children: e.workbook.sheets.map((sheet) =>
                    i.jsx(
                      "li",
                      {
                        children: i.jsxs("button", {
                          type: "button",
                          ref: sheet.id === a.id ? z : undefined,
                          "aria-label": `Manage ${sheet.name}`,
                          "aria-pressed": sheet.id === a.id,
                          onClick: () => U(sheet.id),
                          children: [
                            i.jsx("span", {
                              className: "sm-tab-color",
                              style: {
                                background: sheet.tabColor ?? "#bcc7c0",
                              },
                            }),
                            i.jsxs("span", {
                              className: "sm-sheet-name",
                              title: sheet.name,
                              children: [
                                sheet.name,
                                i.jsxs("small", {
                                  children: [
                                    worksheetVisibilityLabel(
                                      worksheetVisibilityState(sheet),
                                    ),
                                    sheet.id === e.sheetId ? " · Current" : "",
                                  ],
                                }),
                              ],
                            }),
                            sheet.hidden && i.jsx(EyeOff, { size: 14 }),
                          ],
                        }),
                      },
                      sheet.id,
                    ),
                  ),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "sm-details",
              children: [
                i.jsxs("div", {
                  className: "sm-detail-title",
                  children: [
                    i.jsx("h3", { title: a.name, children: a.name }),
                    i.jsxs("button", {
                      type: "button",
                      className: "sm-view",
                      disabled: a.hidden,
                      title: a.hidden
                        ? "Make the worksheet visible before opening it"
                        : "Open this worksheet",
                      onClick: () => {
                        (e.onSelectSheet(a.id), e.onClose());
                      },
                      children: [
                        i.jsx(Eye$1, { size: 14 }),
                        i.jsx(LocalizedText, { children: "Go to sheet" }),
                      ],
                    }),
                  ],
                }),
                e.readOnly &&
                  i.jsx("p", {
                    className: "sm-hint",
                    children: i.jsx(LocalizedText, {
                      children:
                        "The host opened this workbook read-only. You can inspect its worksheets.",
                    }),
                  }),
                _ &&
                  i.jsxs("p", {
                    className: "sm-lock-note",
                    children: [
                      i.jsx(LockKeyhole, { size: 14 }),
                      i.jsx(LocalizedText, {
                        children:
                          "Workbook structure is protected. Unlock it below to make changes.",
                      }),
                    ],
                  }),
                i.jsxs("section", {
                  className: "sm-inspection",
                  "aria-label": text("Worksheet inspection metadata"),
                  lang: language("Worksheet inspection metadata"),
                  "data-ui-label": "Worksheet inspection metadata",
                  children: [
                    i.jsx("h4", {
                      children: i.jsx(LocalizedText, {
                        children: "Worksheet inspection",
                      }),
                    }),
                    i.jsxs("dl", {
                      children: [
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(LocalizedText, {
                                children: "Visibility",
                              }),
                            }),
                            i.jsx("dd", {
                              children: worksheetVisibilityLabel(T.visibility),
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(LocalizedText, {
                                children: "Stored cells",
                              }),
                            }),
                            i.jsx("dd", {
                              children: T.storedCellCount.toLocaleString(),
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(LocalizedText, {
                                children: "Formula cells",
                              }),
                            }),
                            i.jsx("dd", {
                              children: T.formulaCellCount.toLocaleString(),
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(LocalizedText, {
                                children: "Tab color",
                              }),
                            }),
                            i.jsx("dd", { children: T.tabColor ?? "None" }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(LocalizedText, {
                                children: "Calculation",
                              }),
                            }),
                            i.jsx("dd", {
                              children: i.jsx(LocalizedText, {
                                children: "Included when referenced",
                              }),
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Hidden and very hidden worksheets still participate in calculation. This inspection does not evaluate formulas.",
                      }),
                    }),
                  ],
                }),
                i.jsxs("fieldset", {
                  className: "sm-edit-fields",
                  disabled: disabled,
                  children: [
                    i.jsx("legend", {
                      className: "sm-sr-only",
                      children: i.jsx(LocalizedText, {
                        children: "Worksheet settings",
                      }),
                    }),
                    i.jsxs("form", {
                      className: "sm-name-form",
                      onSubmit: (event) => {
                        (event.preventDefault(),
                          G(`Rename worksheet: ${a.name}`, (oe) =>
                            renameWorksheet(oe, a.id, s),
                          ));
                      },
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Sheet name" }),
                            i.jsx("input", {
                              "aria-label": text("Sheet name"),
                              value: s,
                              maxLength: 31,
                              onChange: (event) => d(event.target.value),
                              lang: language("Sheet name"),
                              "data-ui-label": "Sheet name",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "submit",
                          disabled: !s.trim() || s.trim() === a.name,
                          children: i.jsx(LocalizedText, {
                            children: "Rename sheet",
                          }),
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "sm-row-actions",
                      children: [
                        i.jsxs("button", {
                          type: "button",
                          onClick: () =>
                            G(
                              `Duplicate worksheet: ${a.name}`,
                              (workbook) => duplicateWorksheet(workbook, a.id),
                              true,
                            ),
                          children: [
                            i.jsx(Copy, { size: 14 }),
                            i.jsx(LocalizedText, {
                              children: "Duplicate sheet",
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "sm-order",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              "aria-label": text("Move worksheet up"),
                              title: text("Move worksheet earlier"),
                              disabled: R === 0,
                              onClick: () =>
                                G(`Move worksheet: ${a.name}`, (workbook) =>
                                  moveWorksheet(workbook, a.id, R - 1),
                                ),
                              lang: language("Move worksheet up"),
                              "data-ui-label": "Move worksheet up",
                              children: i.jsx(ArrowUp, { size: 15 }),
                            }),
                            i.jsx("button", {
                              type: "button",
                              "aria-label": text("Move worksheet down"),
                              title: text("Move worksheet later"),
                              disabled: R === e.workbook.sheets.length - 1,
                              onClick: () =>
                                G(`Move worksheet: ${a.name}`, (workbook) =>
                                  moveWorksheet(workbook, a.id, R + 1),
                                ),
                              lang: language("Move worksheet down"),
                              "data-ui-label": "Move worksheet down",
                              children: i.jsx(ArrowDown, { size: 15 }),
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "sm-setting-row",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Visibility" }),
                            i.jsxs("select", {
                              "aria-label": text("Worksheet visibility"),
                              value: u,
                              onChange: (event) => l(event.target.value),
                              lang: language("Worksheet visibility"),
                              "data-ui-label": "Worksheet visibility",
                              children: [
                                i.jsx("option", {
                                  value: "visible",
                                  children: "Visible",
                                }),
                                i.jsx("option", {
                                  value: "hidden",
                                  children: "Hidden",
                                }),
                                i.jsx("option", {
                                  value: "veryHidden",
                                  children: "Very hidden",
                                }),
                              ],
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled: u === worksheetVisibilityState(a),
                          onClick: () =>
                            G(
                              `Change worksheet visibility: ${a.name}`,
                              (workbook) =>
                                setWorksheetVisibility(workbook, a.id, u),
                            ),
                          children: i.jsx(LocalizedText, {
                            children: "Apply visibility",
                          }),
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      className: "sm-hint",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Hidden worksheets remain available here. Very hidden records a separate Excel visibility state. At least one worksheet must remain visible.",
                      }),
                    }),
                    i.jsxs("div", {
                      className: "sm-setting-row sm-color-row",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Tab color" }),
                            i.jsx("input", {
                              "aria-label": text("Worksheet tab color"),
                              type: "color",
                              value: tabColor,
                              onChange: (event) => h(event.target.value),
                              lang: language("Worksheet tab color"),
                              "data-ui-label": "Worksheet tab color",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled:
                            tabColor.toLowerCase() ===
                            a.tabColor?.toLowerCase(),
                          onClick: () =>
                            G(`Change tab color: ${a.name}`, (workbook) =>
                              setWorksheetTabColor(workbook, a.id, tabColor),
                            ),
                          children: i.jsx(LocalizedText, {
                            children: "Apply tab color",
                          }),
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled: !a.tabColor,
                          onClick: () =>
                            G(`Remove tab color: ${a.name}`, (workbook) =>
                              setWorksheetTabColor(workbook, a.id),
                            ),
                          children: i.jsx(LocalizedText, {
                            children: "No color",
                          }),
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("div", {
                  className: "sm-delete-row",
                  children: [
                    i.jsxs("button", {
                      type: "button",
                      className: "sm-danger",
                      disabled: !$,
                      title: L
                        ? "Keep at least one visible worksheet"
                        : undefined,
                      onClick: Q,
                      children: [
                        i.jsx(Trash2, { size: 14 }),
                        i.jsx(LocalizedText, { children: "Delete worksheet…" }),
                      ],
                    }),
                    i.jsx("small", {
                      children: i.jsx(LocalizedText, {
                        children: "Review affected references before deleting.",
                      }),
                    }),
                  ],
                }),
                E &&
                  i.jsxs("section", {
                    className: "sm-deletion",
                    "aria-label": text("Delete worksheet confirmation"),
                    lang: language("Delete worksheet confirmation"),
                    "data-ui-label": "Delete worksheet confirmation",
                    children: [
                      i.jsxs("h4", {
                        children: [
                          i.jsx(LocalizedText, { children: "Delete “" }),
                          E.name,
                          "”?",
                        ],
                      }),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "The worksheet and its contents will be removed. Dependent formulas can become #REF!; dependent refresh definitions are removed while their result cells remain.",
                        }),
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx("strong", { children: E.dependencies.length }),
                          " ",
                          i.jsx(LocalizedText, { children: "affected" }),
                          " ",
                          E.dependencies.length === 1
                            ? "reference"
                            : "references",
                        ],
                      }),
                      E.dependencies.length
                        ? i.jsx("ul", {
                            children: E.dependencies
                              .slice(I * 200, (I + 1) * 200)
                              .map((ae, oe) =>
                                i.jsxs(
                                  "li",
                                  {
                                    children: [
                                      i.jsx("span", { children: ae.label }),
                                      ae.address &&
                                        i.jsx("code", { children: ae.address }),
                                    ],
                                  },
                                  `${ae.kind}:${oe}`,
                                ),
                              ),
                          })
                        : i.jsx("p", {
                            className: "sm-hint",
                            children: i.jsx(LocalizedText, {
                              children: "No dependent references were found.",
                            }),
                          }),
                      E.dependencies.length > 200 &&
                        i.jsxs("div", {
                          className: "sm-dependency-pages",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              "aria-label": text("Previous references"),
                              disabled: I === 0,
                              onClick: () => O((ae) => ae - 1),
                              lang: language("Previous references"),
                              "data-ui-label": "Previous references",
                              children: i.jsx(LocalizedText, {
                                children: "Previous",
                              }),
                            }),
                            i.jsxs("span", {
                              "aria-live": "polite",
                              children: [
                                I * 200 + 1,
                                "–",
                                Math.min(E.dependencies.length, (I + 1) * 200),
                                " ",
                                i.jsx(LocalizedText, { children: "of " }),
                                E.dependencies.length,
                                " ",
                                i.jsx(LocalizedText, {
                                  children: "references",
                                }),
                              ],
                            }),
                            i.jsx("button", {
                              type: "button",
                              "aria-label": text("Next references"),
                              disabled: (I + 1) * 200 >= E.dependencies.length,
                              onClick: () => O((ae) => ae + 1),
                              lang: language("Next references"),
                              "data-ui-label": "Next references",
                              children: i.jsx(LocalizedText, {
                                children: "Next",
                              }),
                            }),
                          ],
                        }),
                      W &&
                        i.jsx("p", {
                          className: "sm-error",
                          children: i.jsx(LocalizedText, {
                            children:
                              "Some references cannot be updated safely. Resolve the listed unsupported references first.",
                          }),
                        }),
                      q &&
                        i.jsx("p", {
                          className: "sm-error",
                          children: i.jsx(LocalizedText, {
                            children:
                              "The workbook changed. Review its current dependencies before deleting.",
                          }),
                        }),
                      i.jsxs("label", {
                        className: "sm-confirm",
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: checked,
                            disabled: disabled || q || W,
                            onChange: (event) => F(event.target.checked),
                          }),
                          i.jsx(LocalizedText, {
                            children:
                              "I understand that this removes the worksheet and may leave reference errors.",
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        className: "sm-confirm-actions",
                        children: [
                          i.jsx("button", {
                            type: "button",
                            onClick: () => {
                              (A(null), F(false));
                            },
                            children: i.jsx(LocalizedText, {
                              children: "Keep worksheet",
                            }),
                          }),
                          q
                            ? i.jsx("button", {
                                type: "button",
                                disabled: disabled,
                                onClick: Q,
                                children: i.jsx(LocalizedText, {
                                  children: "Review current dependencies",
                                }),
                              })
                            : i.jsx("button", {
                                type: "button",
                                className: "sm-danger sm-delete-final",
                                disabled: !$ || !checked || W,
                                onClick: X$1,
                                children: i.jsx(LocalizedText, {
                                  children: "Delete worksheet",
                                }),
                              }),
                        ],
                      }),
                    ],
                  }),
              ],
            }),
          ],
        }),
        i.jsxs("section", {
          className: "sm-protection",
          "aria-label": text("Workbook structure protection"),
          lang: language("Workbook structure protection"),
          "data-ui-label": "Workbook structure protection",
          children: [
            i.jsxs("div", {
              className: "sm-protection-title",
              children: [
                _
                  ? i.jsx(LockKeyhole, { size: 16 })
                  : i.jsx(LockKeyholeOpen, { size: 16 }),
                i.jsx("strong", {
                  children: i.jsx(LocalizedText, {
                    children: "Workbook structure",
                  }),
                }),
                i.jsx("span", { children: _ ? "Protected" : "Unlocked" }),
              ],
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Protect worksheet names, order, visibility, tab colors, and adding or deleting sheets.",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Visibility and structure locks control editing. Access to this workbook and its hidden content is granted separately by the host application.",
              }),
            }),
            i.jsxs("form", {
              onSubmit: (event) => {
                (event.preventDefault(), y || J());
              },
              children: [
                i.jsxs("label", {
                  children: [
                    _ ? "Password to unlock" : "Password (optional)",
                    i.jsx("input", {
                      "aria-label": text("Structure password"),
                      type: "password",
                      autoComplete: "off",
                      value: f,
                      maxLength: 200,
                      disabled: e.readOnly || y,
                      onChange: (event) => p(event.target.value),
                      lang: language("Structure password"),
                      "data-ui-label": "Structure password",
                    }),
                  ],
                }),
                i.jsx("button", {
                  type: "submit",
                  disabled: e.readOnly || y,
                  children: y
                    ? "Checking…"
                    : _
                      ? "Unprotect workbook structure"
                      : "Protect workbook structure",
                }),
              ],
            }),
          ],
        }),
        (m || g) &&
          i.jsx("div", {
            className: "sm-feedback",
            children: m
              ? i.jsx("p", {
                  role: "alert",
                  className: "sm-error",
                  children: m,
                })
              : i.jsx("p", { role: "status", children: g }),
          }),
        i.jsxs("footer", {
          className: "sm-footer",
          children: [
            i.jsx("span", {
              children: i.jsx(LocalizedText, {
                children: "Changes use the workbook’s undo history.",
              }),
            }),
            i.jsx("button", {
              type: "button",
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Done" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const worksheetPropertyLabels = {
    raw: "Value",
    literal: "Entered as text",
    style: "Formatting",
    numFmt: "Number format",
    format: "Format",
    note: "Note",
    tickmark: "Tickmark",
    createdAt: "Created",
    modifiedAt: "Modified",
    engagement: "Engagement",
    period: "Period",
    template: "Template",
    name: "Name",
    hidden: "Hidden",
    veryHidden: "Very hidden",
    tabColor: "Tab color",
    columnWidths: "Column widths",
    rowHeights: "Row heights",
    frozenRows: "Frozen rows",
    structureProtection: "Structure protection",
    protection: "Protection",
    definedNames: "Defined names",
    metadata: "Workbook properties",
    enabled: "Enabled",
    sourceRevision: "Source revision",
    version: "Version",
  },
  formatAuditDetailValue = (e) =>
    e === undefined
      ? "Not set"
      : e === null || e === ""
        ? "Blank"
        : typeof e == "boolean"
          ? e
            ? "Yes"
            : "No"
          : typeof e == "string"
            ? e.startsWith("data:image/")
              ? "Embedded picture"
              : /^\d{4}-\d{2}-\d{2}T.*Z$/.test(e) &&
                  Number.isFinite(Date.parse(e))
                ? new Date(e).toLocaleString()
                : e
            : typeof e != "object"
              ? String(e)
              : Array.isArray(e)
                ? e.length
                  ? e.map(formatAuditDetailValue).join("; ")
                  : "None"
                : Object.entries(e)
                    .filter(
                      ([r]) =>
                        ![
                          "id",
                          "sheetId",
                          "scopeSheetId",
                          "sourceWorkbookId",
                          "passwordHash",
                          "salt",
                        ].includes(r),
                    )
                    .map(
                      ([r, t]) =>
                        `${worksheetPropertyLabels[r] ?? r}: ${formatAuditDetailValue(t)}`,
                    )
                    .join(" · ") || "No details",
  formatAuditCellValue = (e) => (e ? formatAuditDetailValue(e) : "Empty cell"),
  formatChangeAttribution = (e) =>
    e?.actor && e.timestamp
      ? `${e.actor} · ${new Date(e.timestamp).toLocaleString()}${e.basis === "operation-history" ? " · from operation history" : ""}`
      : "Author and time not recorded for this change";

export {
  worksheetPropertyLabels,
  formatChangeAttribution,
  formatAuditCellValue,
  WorksheetTabBar,
  formatAuditDetailValue,
  CustomSortDialog,
};
