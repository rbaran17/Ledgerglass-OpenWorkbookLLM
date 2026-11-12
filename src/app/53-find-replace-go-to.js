import { throwWithMessage, isFormulaCellValue } from "./03-workbook-model.js";
import { eI, formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { Search, X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import A2e from "react-dom";
import {
  maxRowHeight,
  workbookLimits,
  minRowHeight,
  maxColumnWidth,
  minColumnWidth,
} from "./02-text-and-financial-functions.js";
import {
  formatNumberWithPattern,
  defaultRowHeightOf,
  columnWidthAt,
  enumerateSelectionCells,
  findSheetOrThrow,
  defaultColumnWidthOf,
  Pv,
  findOccupiedRangeBounds,
  calculateWorkbookWithSettings,
  _t,
  parseRangeAddress,
  indentToPx,
  measureCellBorderInsets,
  normalizeCellRecordInput,
  createAxisVisibilityIndex,
  columnIndexToLetters,
  vx,
  resolveEffectiveCellStyle,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  buildCellChangePreview,
  LocalizedText,
  applyCellChangePreview,
} from "./04-cell-input-paste-and-fill.js";
import {
  getFormulaCompletions,
  applyFormulaCompletion,
} from "./33-page-setup-and-print.js";
import {
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
  rangeContainsCell,
} from "./52-workbook-comparison.js";

function replaceSearchMatches(workbook, r, t, n = {}) {
  (n.readOnly && throwWithMessage("This workbook is read-only."),
    n.readOnly !== undefined &&
      typeof n.readOnly != "boolean" &&
      throwWithMessage("Read-only must be boolean."),
    (typeof t != "string" || t.length > 32767) &&
      throwWithMessage(
        "Replacement text must contain at most 32,767 characters.",
      ),
    r.lookIn &&
      !["raw", "formulas"].includes(r.lookIn) &&
      throwWithMessage(
        "Replace searches Raw contents or Formulas only. Calculated and formatted results are read-only search modes.",
      ));
  const o = searchWorkbookCells(workbook, r),
    options = o.options,
    s = compileSearchPattern(options),
    searchBudget = createSearchBudget(),
    u = n.match
      ? o.results.filter(
          (result) =>
            result.sheetId === n.match.sheetId &&
            result.address === n.match.address.toUpperCase(),
        )
      : o.results;
  n.match &&
    !u.length &&
    throwWithMessage(
      "The chosen cell does not match this search. Find it again before replacing.",
    );
  const l = [];
  let excludedProtected = 0,
    replacedOccurrences = 0;
  for (const g of u) {
    if (g.protected) {
      excludedProtected++;
      continue;
    }
    const b = matchTextWithBudget(g.text, options, s, searchBudget);
    let y = "",
      w = 0;
    g.text.length -
      b.reduce((F, [I, O]) => F + O - I, 0) +
      b.length * t.length >
      32767 &&
      throwWithMessage(
        `Replacement would exceed 32,767 characters at ${g.sheetName}!${g.address}.`,
      );
    for (const [F, I] of b) ((y += g.text.slice(w, F) + t), (w = I));
    y += g.text.slice(w);
    const cell = findSheetOrThrow(workbook, g.sheetId).cells[g.address],
      comparableCellValue = toComparableCellValue(cell, y);
    JSON.stringify(comparableCellValue) !==
      JSON.stringify(normalizeCellRecordInput(cell)) &&
      (l.push({
        sheetId: g.sheetId,
        address: g.address,
        cell: comparableCellValue,
      }),
      (replacedOccurrences += b.length),
      l.length > workbookLimits.maxPatches &&
        throwWithMessage(
          "Replace may change at most 500 cells, including dependent formula changes. Narrow the scope or replace one result.",
        ));
  }
  const f = {
    targetSheetId:
      l[0]?.sheetId ??
      workbook.sheets.find((sheet) => !sheet.protection?.enabled)?.id ??
      options.sheetId,
    targetRange: l[0]?.address ?? options.range,
    targetCount: l.length,
    matchedCells: u.length,
    replacedCells: l.length,
    replacedOccurrences: replacedOccurrences,
    excludedProtected: excludedProtected,
    summary: `Replace in ${options.scope}: ${l.length} cells; ${excludedProtected} protected matches excluded.`,
  };
  return l.length
    ? buildCellChangePreview(workbook, l, f, n.readOnly)
    : eI({
        ...f,
        workbookId: workbook.id,
        revision: workbook.revision,
        patches: [],
        previewPatches: [],
        count: 0,
        dependentCount: 0,
        snapshotGuard: JSON.stringify(workbook),
      });
}
function assertReplacePreviewCurrent(workbook, plan, t = {}) {
  return plan.count
    ? applyCellChangePreview(workbook, plan, t)
    : (t.readOnly !== undefined &&
        typeof t.readOnly != "boolean" &&
        throwWithMessage("Read-only must be boolean."),
      t.readOnly && throwWithMessage("This workbook is read-only."),
      (plan.workbookId !== workbook.id ||
        plan.revision !== workbook.revision ||
        plan.snapshotGuard !== JSON.stringify(workbook)) &&
        throwWithMessage(
          "The workbook changed after this preview. Create a new preview before applying it.",
        ),
      workbook);
}
function buildSelectionRanges(e) {
  const r = [];
  for (const o of [...e].sort((a, s) => compareCellPositions("rows", a, s))) {
    const a = r.at(-1);
    a && a.firstRow === o.row && a.lastCol + 1 === o.col
      ? (a.lastCol = o.col)
      : r.push({
          firstRow: o.row,
          lastRow: o.row,
          firstCol: o.col,
          lastCol: o.col,
        });
  }
  const t = [],
    n = new Map();
  for (const o of r) {
    const a = `${o.firstCol}:${o.lastCol}`,
      s = n.get(a);
    s && s.lastRow + 1 === o.firstRow
      ? (s.lastRow = o.lastRow)
      : (t.push(o), n.set(a, o));
  }
  return t.map((o) => {
    const cellAddress = formatCellAddress(o.firstRow, o.firstCol),
      s = formatCellAddress(o.lastRow, o.lastCol);
    return cellAddress === s ? cellAddress : `${cellAddress}:${s}`;
  });
}
function selectSpecialCells(workbook, r, t) {
  const searchRequest = normalizeSearchRequest(workbook, r),
    sheetOrThrow = findSheetOrThrow(workbook, searchRequest.sheetId),
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    s = describeHiddenState(sheetOrThrow);
  (searchRequest.scope === "workbook" &&
    throwWithMessage(
      "Go To Special selects cells on one sheet. Choose Selection or Sheet.",
    ),
    ["formulas", "errors", "constants", "blanks", "notes"].includes(t) ||
      throwWithMessage("Choose a supported Go To Special category."));
  const occupiedRangeBounds = findOccupiedRangeBounds(sheetOrThrow),
    u =
      searchRequest.scope === "selection"
        ? searchRequest.ranges
        : occupiedRangeBounds
          ? [
              formatCellAddress(
                occupiedRangeBounds.firstRow,
                occupiedRangeBounds.firstCol,
              ) +
                ":" +
                formatCellAddress(
                  occupiedRangeBounds.lastRow,
                  occupiedRangeBounds.lastCol,
                ),
            ]
          : [],
    l = u.map(parseRangeAddress),
    c = [],
    h = [],
    f = () => {
      h.sort((v, g) => compareCellPositions(searchRequest.order, v, g));
      const selectionRanges = buildSelectionRanges(h);
      return (
        selectionRanges.length > comparisonLimits.maxSpecialAreas &&
          throwWithMessage(
            "This selection contains more than 256 separate areas. Narrow the selection.",
          ),
        eI({
          sheetId: sheetOrThrow.id,
          kind: t,
          count: h.length,
          addresses: h.map((v) => formatCellAddress(v.row, v.col)),
          ranges: selectionRanges,
          workbookId: workbook.id,
          revision: workbook.revision,
          snapshotGuard: JSON.stringify(workbook),
        })
      );
    };
  if (!l.length || (sheetOrThrow.hidden && !searchRequest.includeHidden))
    return f();
  if (t === "blanks")
    try {
      c.push(...enumerateSelectionCells(u, comparisonLimits.maxBlankScan));
    } catch {
      throwWithMessage(
        "Blank selection scans at most 100,000 cells. Select a smaller rectangle.",
      );
    }
  else if ((c.push(...Object.keys(sheetOrThrow.cells)), t === "errors"))
    for (const m of workbookWithSettings.getSpillEntries(sheetOrThrow.id))
      Object.hasOwn(sheetOrThrow.cells, m.address) || c.push(m.address);
  c.length > workbookLimits.maxCells &&
    throwWithMessage("Go To Special scans at most 100,000 cells.");
  let p = 0;
  for (const m of c) {
    const v = _t(m);
    if (
      !l.some((w) => rangeContainsCell(v.row, v.col, w)) ||
      (!searchRequest.includeHidden && s(v.row, v.col))
    )
      continue;
    const cell = sheetOrThrow.cells[m],
      b = isFormulaCellValue(cell);
    let y = false;
    if (t === "formulas") y = b;
    else if (t === "notes") y = !!cell?.note;
    else if (t === "constants")
      y = !!cell && cell.raw !== null && cell.raw !== "" && !b;
    else {
      if (!b && cell && cell.raw !== null) continue;
      ++p > comparisonLimits.maxCalculatedCells &&
        t === "errors" &&
        throwWithMessage(
          "Error selection calculates at most 10,000 cells. Narrow the selection.",
        );
      const w = workbookWithSettings(sheetOrThrow.id, m);
      t === "blanks"
        ? (y = (cell?.raw === null || !cell) && w === null)
        : (y =
            (b || !!workbookWithSettings.getSpillAnchor(sheetOrThrow.id, m)) &&
            typeof w == "string" &&
            /^#(?:BLOCKED!|CALC!|CYCLE!|DIV\/0!|ERROR!|LIMIT!|N\/A|NAME\?|NULL!|NUM!|REF!|SPILL!|VALUE!)$/.test(
              w,
            ));
    }
    y &&
      (h.push(v),
      h.length > comparisonLimits.maxSpecialCells &&
        throwWithMessage(
          "Go To Special may select at most 500 cells. Narrow the selection.",
        ));
  }
  return f();
}
const describeOperationError = (e) =>
    e instanceof Error ? e.message : "This operation could not be completed.",
  formatSearchResultValue = (e) =>
    e == null
      ? "(blank)"
      : typeof e == "boolean"
        ? e
          ? "TRUE"
          : "FALSE"
        : String(e),
  Bv = 20;
function FindReplaceDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(null),
    s = ne.useRef(e);
  s.current = e;
  const d = ne.useRef({
      sheetId: e.sheetId,
      range: e.range,
      ranges: e.ranges ? [...e.ranges] : [e.range],
      sheetName:
        e.workbook.sheets.find((sheet) => sheet.id === e.sheetId)?.name ??
        "Worksheet",
    }),
    [u, l] = ne.useState(e.initialTab ?? "find"),
    [scope, h] = ne.useState("sheet"),
    [lookIn, p] = ne.useState("raw"),
    [order, v] = ne.useState("rows"),
    [query, b] = ne.useState(""),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(false),
    [k, F] = ne.useState(false),
    [I, O] = ne.useState(true),
    [j, z] = ne.useState(false),
    [V, D] = ne.useState("formulas"),
    [B, _] = ne.useState(null),
    [match, T] = ne.useState(null),
    [plan, R] = ne.useState(null),
    [L, $] = ne.useState(null),
    [U, G] = ne.useState(0),
    [H, Q] = ne.useState(0),
    [X$1, J] = ne.useState(""),
    [q, W] = ne.useState(""),
    [K, te] = ne.useState({}),
    ae = () => ({
      sheetId: d.current.sheetId,
      range: d.current.range,
      ranges: [...d.current.ranges],
      scope: scope,
      query: query,
      lookIn: lookIn,
      order: order,
      matchCase: E,
      wholeCell: k,
      wildcards: I,
      includeHidden: j,
    });
  function oe() {
    (_(null), T(null), R(null), $(null), G(0), Q(0), J(""), W(""));
  }
  function de(Re) {
    (l(Re),
      Re === "replace" &&
        !["raw", "formulas"].includes(lookIn ?? "raw") &&
        p("raw"),
      Re === "special" && scope === "workbook" && h("sheet"),
      oe());
  }
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      (
        a.current ??
        o.current?.querySelector("select,input") ??
        o.current?.querySelector("button")
      )?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const he = (Re) => {
    const workbook = s.current.workbook;
    if (
      Re.workbookId !== workbook.id ||
      Re.revision !== workbook.revision ||
      Re.snapshotGuard !== JSON.stringify(workbook)
    )
      throw new Error(
        "The workbook changed. Find or preview again before using these results.",
      );
  };
  function onClick() {
    try {
      (J(""), R(null), $(null));
      const Re = searchWorkbookCells(s.current.workbook, ae());
      return (
        _(Re),
        T(null),
        G(0),
        W(
          Re.total
            ? `${Re.total.toLocaleString()} matching cells.`
            : "No matches found.",
        ),
        Re
      );
    } catch (error) {
      return (J(describeOperationError(error)), _(null), null);
    }
  }
  function le(Re, Me = B) {
    try {
      if (!Me) return;
      if (
        (he(Me),
        J(""),
        T(Re),
        G(
          Math.floor(
            Me.results.findIndex(
              (result) =>
                result.sheetId === Re.sheetId && result.address === Re.address,
            ) / Bv,
          ),
        ),
        Re.hidden)
      ) {
        W(
          `${Re.sheetName}!${Re.address} is in a ${Re.hiddenReason}. Unhide it before navigating; its visibility has not changed.`,
        );
        return;
      }
      (s.current.onNavigate(Re.sheetId, Re.address),
        W(
          `${Re.sheetName}!${Re.address}${Re.protected ? " · protected" : ""}`,
        ));
    } catch (error) {
      J(describeOperationError(error));
    }
  }
  function ce(Re) {
    const Me = B ?? onClick();
    if (Me)
      try {
        he(Me);
        const rangeAddress = parseRangeAddress(d.current.range),
          je = match ?? {
            sheetId: d.current.sheetId,
            address:
              e.active ??
              `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
          },
          adjacentSearchResult = findAdjacentSearchResult(Me, je, Re);
        adjacentSearchResult
          ? le(adjacentSearchResult, Me)
          : W("No matches found.");
      } catch (error) {
        J(describeOperationError(error));
      }
  }
  function ge(Re) {
    try {
      if ((J(""), Re && !match))
        throw new Error("Select a matching result before replacing one cell.");
      Re && B && he(B);
      const Me = replaceSearchMatches(s.current.workbook, ae(), y, {
        readOnly: s.current.readOnly,
        ...(Re ? { match: match } : {}),
      });
      (R(Me),
        te(
          Object.fromEntries(
            Me.previewPatches.map((previewPatch) => {
              const sheetOrThrow = findSheetOrThrow(
                s.current.workbook,
                previewPatch.sheetId,
              );
              return [
                `${previewPatch.sheetId}:${previewPatch.address}`,
                {
                  sheetName: sheetOrThrow.name,
                  raw: sheetOrThrow.cells[previewPatch.address]?.raw ?? null,
                },
              ];
            }),
          ),
        ),
        $(null),
        Q(0),
        W(
          Me.matchedCells
            ? `${Me.matchedCells} matched; ${Me.replacedCells} cells to replace; ${Me.excludedProtected} protected matches excluded.`
            : "No matches found.",
        ));
    } catch (error) {
      (R(null), J(describeOperationError(error)));
    }
  }
  function se() {
    if (plan)
      try {
        J("");
        const workbook = s.current.workbook,
          Me = assertReplacePreviewCurrent(workbook, plan, {
            readOnly: s.current.readOnly,
          });
        (Me !== workbook &&
          s.current.onWorkbookChange(Me, `Replace ${plan.replacedCells} cells`),
          s.current.onClose());
      } catch (error) {
        J(describeOperationError(error));
      }
  }
  function we() {
    try {
      J("");
      const specialCells = selectSpecialCells(s.current.workbook, ae(), V);
      ($(specialCells),
        R(null),
        G(0),
        W(
          specialCells.count
            ? `${specialCells.count} cells in ${specialCells.ranges.length} areas.`
            : "No matching cells found.",
        ));
    } catch (error) {
      ($(null), J(describeOperationError(error)));
    }
  }
  function Ee() {
    if (L?.count)
      try {
        if ((J(""), he(L), !s.current.onSelectRanges))
          throw new Error(
            "This embedding host has not connected multiple-area selection.",
          );
        if (findSheetOrThrow(s.current.workbook, L.sheetId).hidden)
          throw new Error("Unhide this sheet before selecting its cells.");
        (s.current.onSelectRanges(L.sheetId, [...L.ranges]),
          s.current.onClose());
      } catch (error) {
        J(describeOperationError(error));
      }
  }
  const Ae = Math.max(1, Math.ceil((B?.total ?? 0) / Bv)),
    ye = Math.min(U, Ae - 1),
    Te = Math.max(1, Math.ceil((plan?.previewPatches.length ?? 0) / Bv)),
    Pe = Math.min(H, Te - 1),
    qe = !!(
      (B &&
        (B.workbookId !== e.workbook.id ||
          B.revision !== e.workbook.revision)) ||
      (plan &&
        (plan.workbookId !== e.workbook.id ||
          plan.revision !== e.workbook.revision)) ||
      (L &&
        (L.workbookId !== e.workbook.id || L.revision !== e.workbook.revision))
    );
  return i.jsx("div", {
    className: "st-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      className: "st-dialog",
      ref: o,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": n,
      onKeyDown: (event) => {
        if (event.key === "Escape")
          (event.preventDefault(),
            event.stopPropagation(),
            s.current.onClose());
        else if (event.key === "Tab") {
          const je = Array.from(
              o.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ).filter((ze) => !ze.hidden),
            $e = je[0],
            Je = je.at(-1);
          ((event.shiftKey && document.activeElement === $e) ||
            (!event.shiftKey && document.activeElement === Je)) &&
            (event.preventDefault(), (event.shiftKey ? Je : $e)?.focus());
        } else
          event.key === "Enter" &&
            event.target instanceof HTMLInputElement &&
            event.target.type !== "checkbox" &&
            (event.preventDefault(), ce(event.shiftKey ? "previous" : "next"));
      },
      children: [
        i.jsxs("header", {
          className: "st-header",
          children: [
            i.jsxs("h2", {
              id: n,
              children: [
                i.jsx(Search, { size: 18 }),
                i.jsx(LocalizedText, { children: "Find and Replace" }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": text("Close Find and Replace"),
              onClick: e.onClose,
              lang: language("Close Find and Replace"),
              "data-ui-label": "Close Find and Replace",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("div", {
          className: "st-tabs",
          role: "tablist",
          "aria-label": text("Search tools"),
          lang: language("Search tools"),
          "data-ui-label": "Search tools",
          children: [
            ["find", "Find"],
            ["replace", "Replace"],
            ["special", "Go To Special"],
          ].map(([Re, Me], ke) =>
            i.jsx(
              "button",
              {
                type: "button",
                role: "tab",
                "aria-selected": u === Re,
                tabIndex: u === Re ? 0 : -1,
                onClick: () => de(Re),
                onKeyDown: (event) => {
                  if (
                    ["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                      event.key,
                    )
                  ) {
                    event.preventDefault();
                    const ze = ["find", "replace", "special"],
                      at =
                        event.key === "Home"
                          ? 0
                          : event.key === "End"
                            ? 2
                            : (ke + (event.key === "ArrowRight" ? 1 : 2)) % 3;
                    (de(ze[at]),
                      event.currentTarget.parentElement?.children[at]?.focus());
                  }
                },
                children: Me,
              },
              Re,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "st-body",
          children: [
            i.jsxs("p", {
              className: "st-context",
              children: [
                i.jsx(LocalizedText, { children: "Selection:" }),
                " ",
                i.jsxs("strong", {
                  children: [
                    d.current.sheetName,
                    "!",
                    d.current.ranges.join(", "),
                  ],
                }),
              ],
            }),
            i.jsxs("div", {
              className: "st-options",
              children: [
                u !== "special"
                  ? i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("label", {
                          className: "st-wide",
                          children: [
                            i.jsx(LocalizedText, { children: "Find what" }),
                            i.jsx("input", {
                              ref: a,
                              value: query,
                              maxLength: 256,
                              onChange: (event) => {
                                (b(event.target.value), oe());
                              },
                            }),
                          ],
                        }),
                        u === "replace" &&
                          i.jsxs("label", {
                            className: "st-wide",
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Replace with",
                              }),
                              i.jsx("input", {
                                value: y,
                                maxLength: 32767,
                                onChange: (event) => {
                                  (w(event.target.value), R(null), J(""));
                                },
                              }),
                            ],
                          }),
                      ],
                    })
                  : i.jsxs("label", {
                      className: "st-wide",
                      children: [
                        i.jsx(LocalizedText, {
                          children: "Select cells containing",
                        }),
                        i.jsxs("select", {
                          value: V,
                          onChange: (event) => {
                            (D(event.target.value), oe());
                          },
                          children: [
                            i.jsx("option", {
                              value: "formulas",
                              children: "Formulas",
                            }),
                            i.jsx("option", {
                              value: "errors",
                              children: "Formula errors",
                            }),
                            i.jsx("option", {
                              value: "constants",
                              children: "Constants",
                            }),
                            i.jsx("option", {
                              value: "blanks",
                              children: "Blanks",
                            }),
                            i.jsx("option", {
                              value: "notes",
                              children: "Notes",
                            }),
                          ],
                        }),
                      ],
                    }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Within" }),
                    i.jsxs("select", {
                      value: scope,
                      onChange: (event) => {
                        (h(event.target.value), oe());
                      },
                      children: [
                        i.jsx("option", {
                          value: "selection",
                          children: "Selection",
                        }),
                        i.jsx("option", { value: "sheet", children: "Sheet" }),
                        u !== "special" &&
                          i.jsx("option", {
                            value: "workbook",
                            children: "Workbook",
                          }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Search order" }),
                    i.jsxs("select", {
                      value: order,
                      onChange: (event) => {
                        (v(event.target.value), oe());
                      },
                      children: [
                        i.jsx("option", { value: "rows", children: "By rows" }),
                        i.jsx("option", {
                          value: "columns",
                          children: "By columns",
                        }),
                      ],
                    }),
                  ],
                }),
                u !== "special" &&
                  i.jsxs("label", {
                    className: "st-wide",
                    children: [
                      i.jsx(LocalizedText, { children: "Look in" }),
                      i.jsxs("select", {
                        value: lookIn,
                        onChange: (event) => {
                          (p(event.target.value), oe());
                        },
                        children: [
                          i.jsx("option", {
                            value: "raw",
                            children:
                              "Raw contents (text, numbers and formulas)",
                          }),
                          i.jsx("option", {
                            value: "formulas",
                            children: "Formulas only",
                          }),
                          u === "find" &&
                            i.jsxs(i.Fragment, {
                              children: [
                                i.jsx("option", {
                                  value: "calculated",
                                  children: "Calculated values",
                                }),
                                i.jsx("option", {
                                  value: "formatted",
                                  children: "Formatted display",
                                }),
                              ],
                            }),
                        ],
                      }),
                    ],
                  }),
              ],
            }),
            i.jsxs("div", {
              className: "st-switches",
              children: [
                u !== "special" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: E,
                            onChange: (event) => {
                              (A(event.target.checked), oe());
                            },
                          }),
                          i.jsx(LocalizedText, { children: "Match case" }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: k,
                            onChange: (event) => {
                              (F(event.target.checked), oe());
                            },
                          }),
                          i.jsx(LocalizedText, {
                            children: "Match entire cell contents",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: I,
                            onChange: (event) => {
                              (O(event.target.checked), oe());
                            },
                          }),
                          i.jsx(LocalizedText, {
                            children: "Use wildcards (* ? ~)",
                          }),
                        ],
                      }),
                    ],
                  }),
                i.jsxs("label", {
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: j,
                      onChange: (event) => {
                        (z(event.target.checked), oe());
                      },
                    }),
                    i.jsx(LocalizedText, {
                      children: "Include hidden cells and sheets",
                    }),
                  ],
                }),
              ],
            }),
            i.jsx("p", {
              className: "st-help",
              children:
                u === "special"
                  ? "Select up to 500 cells in 256 separate areas. Sheet scope uses the occupied rectangle for blanks. Formula results of empty text are not blank cells."
                  : "? matches one character; * matches any sequence. Prefix *, ? or ~ with ~ to find it literally. Search follows workbook sheet order.",
            }),
            e.readOnly &&
              i.jsx("p", {
                className: "st-notice",
                children: i.jsx(LocalizedText, {
                  children:
                    "This workbook is read-only. Finding and selecting cells remain available.",
                }),
              }),
            qe &&
              i.jsx("p", {
                className: "st-notice",
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "The workbook changed. Find or preview again before continuing.",
                }),
              }),
            i.jsx("div", {
              className: "st-actions",
              children:
                u === "special"
                  ? i.jsx("button", {
                      type: "button",
                      onClick: we,
                      children: i.jsx(LocalizedText, {
                        children: "Preview selection",
                      }),
                    })
                  : i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("button", {
                          type: "button",
                          onClick: onClick,
                          children: i.jsx(LocalizedText, {
                            children: "Find all",
                          }),
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => ce("previous"),
                          children: i.jsx(LocalizedText, {
                            children: "Find previous",
                          }),
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => ce("next"),
                          children: i.jsx(LocalizedText, {
                            children: "Find next",
                          }),
                        }),
                        u === "replace" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("button", {
                                type: "button",
                                disabled: e.readOnly || !match,
                                onClick: () => ge(true),
                                children: i.jsx(LocalizedText, {
                                  children: "Preview replace selected",
                                }),
                              }),
                              i.jsx("button", {
                                type: "button",
                                disabled: e.readOnly,
                                onClick: () => ge(false),
                                children: i.jsx(LocalizedText, {
                                  children: "Preview replace all",
                                }),
                              }),
                            ],
                          }),
                      ],
                    }),
            }),
            X$1 &&
              i.jsx("p", {
                className: "st-error",
                role: "alert",
                children: X$1,
              }),
            q &&
              i.jsx("p", {
                className: "st-status",
                role: "status",
                children: q,
              }),
            B &&
              i.jsxs("div", {
                className: "st-results",
                "aria-label": text("Find results"),
                lang: language("Find results"),
                "data-ui-label": "Find results",
                children: [
                  i.jsx("div", {
                    className: "st-table-scroll",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Sheet",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Cell",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Matching contents",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "State",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: normalizeSearchPage(B, ye, Bv).map((Re) =>
                            i.jsxs(
                              "tr",
                              {
                                className:
                                  match?.sheetId === Re.sheetId &&
                                  match.address === Re.address
                                    ? "st-selected"
                                    : "",
                                children: [
                                  i.jsx("td", { children: Re.sheetName }),
                                  i.jsx("td", {
                                    children: i.jsx("button", {
                                      type: "button",
                                      "aria-label": `Go to ${Re.sheetName}!${Re.address}`,
                                      onClick: () => le(Re),
                                      children: Re.address,
                                    }),
                                  }),
                                  i.jsx("td", {
                                    className: "st-value",
                                    title: Re.text,
                                    children: Re.text,
                                  }),
                                  i.jsx("td", {
                                    children:
                                      [
                                        Re.protected ? "Protected" : "",
                                        Re.hidden ? Re.hiddenReason : "",
                                      ]
                                        .filter(Boolean)
                                        .join(", ") || "Visible",
                                  }),
                                ],
                              },
                              `${Re.sheetId}:${Re.address}`,
                            ),
                          ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "st-pagination",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        "aria-label": text("Previous results page"),
                        disabled: !ye,
                        onClick: () => G(ye - 1),
                        lang: language("Previous results page"),
                        "data-ui-label": "Previous results page",
                        children: i.jsx(LocalizedText, {
                          children: "Previous",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Results page " }),
                          ye + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          Ae,
                        ],
                      }),
                      i.jsx("button", {
                        type: "button",
                        "aria-label": text("Next results page"),
                        disabled: ye + 1 >= Ae,
                        onClick: () => G(ye + 1),
                        lang: language("Next results page"),
                        "data-ui-label": "Next results page",
                        children: i.jsx(LocalizedText, { children: "Next" }),
                      }),
                    ],
                  }),
                ],
              }),
            plan &&
              i.jsxs("div", {
                className: "st-preview",
                "aria-label": text("Replacement preview"),
                lang: language("Replacement preview"),
                "data-ui-label": "Replacement preview",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Replacement preview",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      plan.replacedCells,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "directly replaced cells · ",
                      }),
                      plan.count,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "total changed cells · ",
                      }),
                      plan.excludedProtected,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "protected matches excluded.",
                      }),
                    ],
                  }),
                  plan.dependentCount > 0 &&
                    i.jsxs("p", {
                      className: "st-notice",
                      children: [
                        i.jsx(LocalizedText, { children: "Includes " }),
                        plan.dependentCount,
                        " ",
                        i.jsx(LocalizedText, {
                          children:
                            "dependent cell changes outside the direct replacement targets. Every change is listed below.",
                        }),
                      ],
                    }),
                  i.jsx("div", {
                    className: "st-table-scroll",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Sheet / cell",
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
                          children: plan.previewPatches
                            .slice(Pe * Bv, (Pe + 1) * Bv)
                            .map((patch) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsxs("td", {
                                      children: [
                                        K[`${patch.sheetId}:${patch.address}`]
                                          ?.sheetName ?? patch.sheetId,
                                        "!",
                                        patch.address,
                                      ],
                                    }),
                                    i.jsx("td", {
                                      className: "st-value",
                                      children: formatSearchResultValue(
                                        K[`${patch.sheetId}:${patch.address}`]
                                          ?.raw,
                                      ),
                                    }),
                                    i.jsx("td", {
                                      className: "st-value",
                                      children: formatSearchResultValue(
                                        patch.cell.raw,
                                      ),
                                    }),
                                  ],
                                },
                                `${patch.sheetId}:${patch.address}`,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "st-pagination",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        "aria-label": text("Previous preview page"),
                        disabled: !Pe,
                        onClick: () => Q(Pe - 1),
                        lang: language("Previous preview page"),
                        "data-ui-label": "Previous preview page",
                        children: i.jsx(LocalizedText, {
                          children: "Previous",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Preview page " }),
                          Pe + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          Te,
                        ],
                      }),
                      i.jsx("button", {
                        type: "button",
                        "aria-label": text("Next preview page"),
                        disabled: Pe + 1 >= Te,
                        onClick: () => Q(Pe + 1),
                        lang: language("Next preview page"),
                        "data-ui-label": "Next preview page",
                        children: i.jsx(LocalizedText, { children: "Next" }),
                      }),
                    ],
                  }),
                ],
              }),
            L &&
              i.jsxs("div", {
                className: "st-preview",
                "aria-label": text("Selection preview"),
                lang: language("Selection preview"),
                "data-ui-label": "Selection preview",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Selection preview",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      L.count,
                      " ",
                      i.jsx(LocalizedText, { children: "cells in " }),
                      L.ranges.length,
                      " ",
                      i.jsx(LocalizedText, { children: "areas" }),
                    ],
                  }),
                  i.jsx("div", {
                    className: "st-ranges",
                    children: L.ranges.map((range) =>
                      i.jsx("code", { children: range }, range),
                    ),
                  }),
                  !e.onSelectRanges &&
                    i.jsx("p", {
                      className: "st-notice",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Multiple-area selection is not connected by this embedding host.",
                      }),
                    }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "st-footer",
          children: [
            i.jsx("button", {
              type: "button",
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Close" }),
            }),
            u === "replace" &&
              i.jsx("button", {
                type: "button",
                className: "st-primary",
                disabled: e.readOnly || !plan?.count || qe,
                onClick: se,
                children: i.jsx(LocalizedText, {
                  children: "Apply replacement",
                }),
              }),
            u === "special" &&
              i.jsx("button", {
                type: "button",
                className: "st-primary",
                disabled: !L?.count || !e.onSelectRanges || qe,
                onClick: Ee,
                children: i.jsx(LocalizedText, { children: "Select cells" }),
              }),
          ],
        }),
      ],
    }),
  });
}
function resolveCellNavigation(e, r, t, n) {
  if (!Number.isInteger(e) || !Number.isInteger(r) || r < 1 || e < 0 || e >= r)
    throw new Error("Navigation requires a valid visible cell.");
  if (n !== 1 && n !== -1) throw new Error("Choose a navigation direction.");
  const o = new Set(t.filter((d) => Number.isInteger(d) && d >= 0 && d < r)),
    a = e + n;
  if (a < 0 || a >= r) return e;
  if (o.has(e) && o.has(a)) {
    let d = a;
    for (; o.has(d + n);) d += n;
    return d;
  }
  let s = n === 1 ? r - 1 : 0;
  for (const d of o) (n === 1 ? d > e && d < s : d < e && d > s) && (s = d);
  return s;
}
const paneQuadrantNames = [
    "top-left",
    "top-right",
    "bottom-left",
    "bottom-right",
  ],
  createDefaultPaneRects = () =>
    Object.fromEntries(
      paneQuadrantNames.map((paneQuadrantName) => [
        paneQuadrantName,
        { top: 0, left: 0, width: 1e3, height: 600 },
      ]),
    );
function useSplitPaneSizes(splitPaneSizesKey) {
  const [views, t] = ne.useState(createDefaultPaneRects),
    n = ne.useRef(views),
    refs = ne.useRef({}),
    [activeId, s] = ne.useState("top-left"),
    active = ne.useRef("top-left");
  function patch(h) {
    let f = n.current;
    for (const paneQuadrantName of paneQuadrantNames)
      if (h[paneQuadrantName]) {
        const m = { ...f[paneQuadrantName], ...h[paneQuadrantName] };
        Object.entries(m).some(([v, g]) => f[paneQuadrantName][v] !== g) &&
          (f = { ...f, [paneQuadrantName]: m });
      }
    f !== n.current && ((n.current = f), t(f));
  }
  const bindings = ne.useMemo(
    () =>
      Object.fromEntries(
        paneQuadrantNames.map((paneQuadrantName) => [
          paneQuadrantName,
          (f) => {
            f
              ? (refs.current[paneQuadrantName] = f)
              : delete refs.current[paneQuadrantName];
          },
        ]),
      ),
    [],
  );
  ne.useEffect(() => {
    const h = () => {
      const p = {};
      for (const paneQuadrantName of paneQuadrantNames) {
        const v = refs.current[paneQuadrantName];
        v?.clientWidth &&
          v.clientHeight &&
          (p[paneQuadrantName] = {
            width: v.clientWidth,
            height: v.clientHeight,
          });
      }
      patch(p);
    };
    if ((h(), typeof ResizeObserver > "u")) return;
    const f = new ResizeObserver(h);
    return (
      Object.values(refs.current).forEach((p) => f.observe(p)),
      () => f.disconnect()
    );
  }, [splitPaneSizesKey]);
  function activate(h) {
    ((active.current = h), s(h));
  }
  return {
    views: views,
    current: n,
    refs: refs,
    bindings: bindings,
    activeId: activeId,
    active: active,
    activate: activate,
    patch: patch,
  };
}
function SplitPaneLayout(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useRef(null),
    o = ne.useRef(e);
  o.current = e;
  const [a, s] = ne.useState(null),
    d = ne.useRef(a);
  d.current = a;
  const [u, l] = ne.useState(""),
    sheetSplitPanes =
      a && a.snapshot === e.snapshot && !e.readOnly
        ? { ...e.split, [a.axis]: a.value }
        : e.split,
    h = (f, p) => {
      o.current.readOnly ||
        !o.current.split ||
        e.onResize({ ...o.current.split, [f]: Math.max(20, Math.min(80, p)) });
    };
  return i.jsxs("div", {
    ref: n,
    className: `lg-pane-layout ${sheetSplitPanes ? "is-split" : ""}`,
    style: {
      gridTemplateColumns:
        sheetSplitPanes?.vertical !== undefined
          ? `minmax(0,${sheetSplitPanes.vertical}fr) 7px minmax(0,${100 - sheetSplitPanes.vertical}fr)`
          : "minmax(0,1fr)",
      gridTemplateRows:
        sheetSplitPanes?.horizontal !== undefined
          ? `minmax(0,${sheetSplitPanes.horizontal}fr) 7px minmax(0,${100 - sheetSplitPanes.horizontal}fr)`
          : "minmax(0,1fr)",
    },
    children: [
      Pv(sheetSplitPanes).map((f) =>
        i.jsx(
          "div",
          {
            className: `lg-pane ${e.active === f ? "is-active" : ""}`,
            "data-pane": f,
            style: {
              gridRow: f.startsWith("bottom") ? 3 : 1,
              gridColumn: f.endsWith("right") ? 3 : 1,
            },
            children: e.renderPane(f),
          },
          f,
        ),
      ),
      ["horizontal", "vertical"].map((axis) =>
        sheetSplitPanes?.[axis] === undefined
          ? null
          : i.jsx(
              "div",
              {
                className: `lg-split-divider lg-split-${axis}`,
                role: "separator",
                tabIndex: e.readOnly ? -1 : 0,
                "aria-label": `Resize ${axis} split`,
                "aria-orientation": axis,
                "aria-valuenow": sheetSplitPanes[axis],
                "aria-valuemin": 20,
                "aria-valuemax": 80,
                "aria-disabled": e.readOnly,
                title: text(
                  "Drag or use arrow keys to resize; double-click to center",
                ),
                style:
                  axis === "horizontal"
                    ? { gridRow: 2, gridColumn: "1 / -1" }
                    : { gridColumn: 2, gridRow: "1 / -1" },
                onDoubleClick: () => h(axis, 50),
                onKeyDown: (event) => {
                  if (event.key === "Escape") {
                    ((d.current = null), s(null));
                    return;
                  }
                  const m = (
                    axis === "horizontal"
                      ? ["ArrowUp", "ArrowDown"]
                      : ["ArrowLeft", "ArrowRight"]
                  ).indexOf(event.key);
                  m >= 0
                    ? (event.preventDefault(),
                      h(
                        axis,
                        sheetSplitPanes[axis] +
                          (m ? 1 : -1) * (event.shiftKey ? 10 : 1),
                      ))
                    : (event.key === "Home" || event.key === "End") &&
                      (event.preventDefault(),
                      h(axis, event.key === "Home" ? 20 : 80));
                },
                onPointerDown: (event) => {
                  var v;
                  if (e.readOnly || event.button !== 0) return;
                  (event.preventDefault(), event.currentTarget.focus(), l(""));
                  const m = {
                    axis: axis,
                    value: sheetSplitPanes[axis],
                    snapshot: e.snapshot,
                    snapshotGuard: JSON.stringify(e.snapshot),
                    split: e.split,
                  };
                  ((d.current = m),
                    s(m),
                    (v = event.currentTarget).setPointerCapture?.call(
                      v,
                      event.pointerId,
                    ));
                },
                onPointerMove: (event) => {
                  const m = d.current,
                    v = n.current?.getBoundingClientRect();
                  if (!m || !v) return;
                  const g = axis === "horizontal" ? v.height : v.width;
                  if (g <= 0) return;
                  const b =
                      axis === "horizontal"
                        ? event.clientY - v.top
                        : event.clientX - v.left,
                    y = {
                      ...m,
                      value: Math.max(
                        20,
                        Math.min(80, Math.round((b / g) * 100)),
                      ),
                    };
                  ((d.current = y), s(y));
                },
                onPointerUp: (event) => {
                  var v;
                  const m = d.current;
                  if (((d.current = null), s(null), !!m)) {
                    if (
                      m.snapshot !== o.current.snapshot ||
                      m.snapshotGuard !== JSON.stringify(o.current.snapshot) ||
                      o.current.readOnly
                    ) {
                      l("The workbook changed. Resize the split again.");
                      return;
                    }
                    (m.value !== m.split[axis] && h(axis, m.value),
                      (v = event.currentTarget).hasPointerCapture?.call(
                        v,
                        event.pointerId,
                      ) &&
                        event.currentTarget.releasePointerCapture(
                          event.pointerId,
                        ));
                  }
                },
                onPointerCancel: () => {
                  ((d.current = null), s(null));
                },
                onLostPointerCapture: () => {
                  ((d.current = null), s(null));
                },
                lang: language(
                  "Drag or use arrow keys to resize; double-click to center",
                ),
              },
              axis,
            ),
      ),
      u &&
        i.jsx("span", {
          role: "alert",
          className: "lg-split-error",
          children: u,
        }),
    ],
  });
}

function isWithinFormulaSuggestions(currentTarget, relatedTarget) {
  if (!(relatedTarget instanceof Node)) return false;
  const t = document.getElementById(
    currentTarget.getAttribute("aria-controls") ?? "",
  );
  return !!t?.closest(".formula-suggestions")?.contains(relatedTarget);
}
const FormulaSuggestionList = ne.forwardRef(function (
    {
      value,
      caret = value.length,
      workbook,
      sheetId,
      inputRef,
      onAccept,
      disabled,
      onExit,
    },
    l,
  ) {
    const c = ne.useId(),
      h = ne.useRef(null),
      f = ne.useMemo(
        () => getFormulaCompletions(value, caret, workbook, sheetId),
        [value, caret, workbook, sheetId],
      ),
      [p, m] = ne.useState(-1),
      [v, g] = ne.useState(null),
      [b, y] = ne.useState(false),
      [w, E] = ne.useState({ left: 0, top: 0, width: 280 }),
      A = `${value}:${caret}`,
      k = !disabled && b && v !== A && (f.items.length > 0 || f.hint);
    (ne.useEffect(() => {
      const I = inputRef.current;
      if (!I) return;
      const O = ["aria-autocomplete", "aria-controls", "aria-activedescendant"],
        j = O.map((z) => I.getAttribute(z));
      return (
        k && f.items.length
          ? (I.setAttribute("aria-autocomplete", "list"),
            I.setAttribute("aria-controls", `${c}-list`),
            p >= 0 && f.items[p]
              ? I.setAttribute("aria-activedescendant", `${c}-${p}`)
              : I.removeAttribute("aria-activedescendant"))
          : (I.removeAttribute("aria-controls"),
            I.removeAttribute("aria-activedescendant")),
        () =>
          O.forEach((z, V) =>
            j[V] === null ? I.removeAttribute(z) : I.setAttribute(z, j[V]),
          )
      );
    }, [k, p, f.items.length, inputRef, c]),
      ne.useEffect(() => m(-1), [A]),
      ne.useEffect(() => {
        const I = inputRef.current;
        if (!I) return;
        const O = () => {
            y(
              document.activeElement === I ||
                !!h.current?.contains(document.activeElement),
            );
            const z = I.getBoundingClientRect();
            E({
              left: Math.max(8, Math.min(z.left, window.innerWidth - 350)),
              top: z.bottom + 3,
              width: Math.max(280, Math.min(400, z.width)),
            });
          },
          j = (z) => {
            h.current?.contains(z.relatedTarget) || y(false);
          };
        return (
          O(),
          I.addEventListener("focus", O),
          I.addEventListener("blur", j),
          window.addEventListener("resize", O),
          window.addEventListener("scroll", O, true),
          () => {
            (I.removeEventListener("focus", O),
              I.removeEventListener("blur", j),
              window.removeEventListener("resize", O),
              window.removeEventListener("scroll", O, true));
          }
        );
      }, [inputRef]),
      ne.useEffect(() => {
        if (!k || p < 0) return;
        const I = document.getElementById(`${c}-${p}`),
          O = h.current;
        if (!I || !O) return;
        const j = I.getBoundingClientRect(),
          z = O.getBoundingClientRect();
        j.bottom > z.bottom
          ? (O.scrollTop += j.bottom - z.bottom)
          : j.top < z.top && (O.scrollTop -= z.top - j.top);
      }, [k, p, c]));
    const F = (I) => {
      const O = f.items[I];
      if (!O) return;
      inputRef.current?.focus({ preventScroll: true });
      const j = applyFormulaCompletion(value, O);
      (onAccept(j.value, j.caret),
        m(-1),
        g(null),
        requestAnimationFrame(() => {
          (inputRef.current?.focus(),
            inputRef.current?.setSelectionRange(j.caret, j.caret));
        }));
    };
    return (
      ne.useImperativeHandle(l, () => ({
        handleKeyDown(I) {
          return I.nativeEvent.isComposing || I.defaultPrevented || !k
            ? false
            : (I.key === "ArrowDown" || I.key === "ArrowUp") && f.items.length
              ? (I.preventDefault(),
                m((O) =>
                  I.key === "ArrowDown"
                    ? (O + 1) % f.items.length
                    : O <= 0
                      ? f.items.length - 1
                      : O - 1,
                ),
                true)
              : I.key === "Tab" &&
                  !I.shiftKey &&
                  p < 0 &&
                  f.items.length &&
                  /[A-Za-z_][A-Za-z0-9_.]*(?:\[[^\]]*)?$/.test(
                    value.slice(0, caret),
                  )
                ? (I.preventDefault(), F(0), true)
                : (I.key === "Tab" || I.key === "Enter") && p >= 0 && f.items[p]
                  ? (I.preventDefault(), F(p), true)
                  : I.key === "Escape"
                    ? (I.preventDefault(), g(A), true)
                    : false;
        },
      })),
      !k || typeof document > "u"
        ? null
        : A2e.createPortal(
            i.jsxs("div", {
              className: "formula-suggestions",
              ref: h,
              role: "region",
              "aria-label": "Formula assistance",
              onBlur: (event) => {
                event.currentTarget.contains(event.relatedTarget) ||
                  event.relatedTarget === inputRef.current ||
                  (y(false), onExit?.());
              },
              style: { position: "fixed", ...w },
              children: [
                f.items.length > 0 &&
                  i.jsx("ul", {
                    id: `${c}-list`,
                    role: "listbox",
                    "aria-label": "Formula suggestions",
                    tabIndex: 0,
                    "aria-activedescendant":
                      p >= 0 && f.items[p] ? `${c}-${p}` : undefined,
                    onFocus: () => {
                      (y(true), p < 0 && m(0));
                    },
                    onKeyDown: (event) => {
                      var O;
                      event.nativeEvent.isComposing ||
                        (event.key === "ArrowDown" || event.key === "ArrowUp"
                          ? (event.preventDefault(),
                            m((j) =>
                              event.key === "ArrowDown"
                                ? (j + 1) % f.items.length
                                : (j <= 0 ? f.items.length : j) - 1,
                            ))
                          : event.key === "Enter" ||
                              (event.key === "Tab" && !event.shiftKey)
                            ? (event.preventDefault(), F(Math.max(0, p)))
                            : (event.key === "Escape" || event.key === "Tab") &&
                              (event.preventDefault(),
                              event.key === "Escape" && g(A),
                              (O = inputRef.current) == null || O.focus()));
                    },
                    children: f.items.map((item, O) =>
                      i.jsxs(
                        "li",
                        {
                          id: `${c}-${O}`,
                          role: "option",
                          "aria-selected": p === O,
                          onClick: (j) => {
                            j.detail === 0 && F(O);
                          },
                          onMouseDown: (event) => {
                            (event.preventDefault(), F(O));
                          },
                          onMouseEnter: () => m(O),
                          children: [
                            i.jsx("span", {
                              className: "formula-suggestion-kind",
                              children:
                                item.kind === "function"
                                  ? "ƒx"
                                  : item.kind === "table"
                                    ? "▦"
                                    : "N",
                            }),
                            i.jsxs("div", {
                              children: [
                                i.jsx("strong", { children: item.label }),
                                i.jsx("small", { children: item.detail }),
                              ],
                            }),
                          ],
                        },
                        `${item.kind}:${item.label}`,
                      ),
                    ),
                  }),
                f.hint &&
                  i.jsxs("div", {
                    className: "formula-argument-hint",
                    role: "status",
                    "aria-live": "polite",
                    "aria-atomic": "true",
                    children: [
                      i.jsx("strong", { children: f.hint.name }),
                      "(",
                      f.hint.signature,
                      ")",
                      i.jsxs("small", {
                        children: [
                          i.jsx(LocalizedText, { children: "Argument " }),
                          f.hint.argumentIndex + 1,
                        ],
                      }),
                    ],
                  }),
                f.items.length > 0 &&
                  i.jsx("div", {
                    className: "formula-suggestion-help",
                    children: i.jsx(LocalizedText, {
                      children: "↑ ↓ choose · Tab / Enter insert · Esc close",
                    }),
                  }),
              ],
            }),
            document.body,
          )
    );
  }),
  k2e = 6,
  S2e = 1.35,
  clampNumber = (e, r, t) => Math.max(r, Math.min(t, e));
function buildCellFontString(e) {
  const effectiveCellStyle = resolveEffectiveCellStyle(e?.style),
    t = effectiveCellStyle?.fontSize
      ? `${effectiveCellStyle.fontSize}pt`
      : "14px",
    n = effectiveCellStyle?.fontFamily
      ? JSON.stringify(effectiveCellStyle.fontFamily)
      : "Calibri, Aptos, Arial, sans-serif";
  return `${effectiveCellStyle?.italic ? "italic" : "normal"} ${effectiveCellStyle?.bold ? "700" : "400"} ${t} ${n}`;
}
function parseFontSizePx(e, cellFontString) {
  const t = cellFontString.match(/([\d.]+)(px|pt)/),
    n = t ? Number(t[1]) * (t[2] === "pt" ? 4 / 3 : 1) : 14,
    o = cellFontString.includes("700") ? 1.06 : 1;
  return (
    Array.from(e).reduce(
      (a, s) =>
        a +
        (/\s/u.test(s)
          ? 0.33
          : /[ilI.,'`!:;]/.test(s)
            ? 0.28
            : /[MW@#]/.test(s)
              ? 0.85
              : /[^\u0000-\u024f]/u.test(s)
                ? 1
                : 0.56),
      0,
    ) *
    n *
    o
  );
}
function createTextMeasurer() {
  if (typeof document > "u" || typeof CanvasRenderingContext2D > "u")
    return parseFontSizePx;
  try {
    const e = document.createElement("canvas").getContext("2d");
    if (!e) return parseFontSizePx;
    let r = "";
    return (t, cellFontString) => {
      r !== cellFontString && ((e.font = cellFontString), (r = cellFontString));
      const width = e.measureText(t).width;
      return Number.isFinite(width) && width >= 0
        ? width
        : parseFontSizePx(t, cellFontString);
    };
  } catch {
    return parseFontSizePx;
  }
}
function measureTextWidth(textMeasurer, r, cellFontString) {
  try {
    const n = textMeasurer(r, cellFontString);
    if (Number.isFinite(n) && n >= 0) return n;
  } catch {}
  return parseFontSizePx(r, cellFontString);
}
function measureTextWithTabs(e, cellFontString, textMeasurer) {
  if (!e.includes("	")) return measureTextWidth(textMeasurer, e, cellFontString);
  const n = Math.max(
    1,
    measureTextWidth(textMeasurer, " ", cellFontString) * 8,
  );
  let o = 0;
  return (
    e.split("	").forEach((a, s) => {
      (s && (o = (Math.floor(o / n) + 1) * n),
        (o += measureTextWidth(textMeasurer, a, cellFontString)));
    }),
    o
  );
}
function resolveDisplayedCellText(
  e,
  r,
  cell,
  workbookWithSettings,
  isShowingFormulas,
  a,
) {
  return isShowingFormulas &&
    typeof cell.raw == "string" &&
    cell.raw.startsWith("=") &&
    !cell.literal
    ? cell.raw
    : formatNumberWithPattern(workbookWithSettings(e.id, r), cell, a);
}
function measureNoteIndicatorWidth(e, textMeasurer) {
  return Math.max(
    e.note ? 5 : 0,
    e.tickmark
      ? Math.min(
          40,
          measureTextWidth(textMeasurer, e.tickmark, "normal 400 8px Arial") +
            4,
        )
      : 0,
  );
}
function computeCellTextAlignment(e, textMeasurer) {
  return (
    k2e +
    measureCellBorderInsets(e.style).horizontal +
    (e.style?.align === "center" ? 0 : indentToPx(e.style)) +
    measureNoteIndicatorWidth(e, textMeasurer)
  );
}
function validateAutoFitIndices(selectedAxisIndexes, r, t) {
  for (const selectedAxisIndex of selectedAxisIndexes)
    if (
      !Number.isInteger(selectedAxisIndex) ||
      selectedAxisIndex < 0 ||
      selectedAxisIndex >= r
    )
      throw new Error(`AutoFit ${t} index is outside the worksheet.`);
  return new Set(selectedAxisIndexes);
}
function autoFitColumnWidths(
  e,
  selectedAxisIndexes,
  workbookWithSettings,
  textMeasurer,
  isShowingFormulas = false,
  a = "1900",
) {
  const s = validateAutoFitIndices(selectedAxisIndexes, e.colCount, "column"),
    d = Object.fromEntries(
      [...s].map((l) => [String(l), defaultColumnWidthOf(e)]),
    ),
    u = new Set();
  for (const [l, c] of Object.entries(vx(e, workbookWithSettings))) {
    const { col } = _t(l);
    if (!s.has(col)) continue;
    const displayedCellText = resolveDisplayedCellText(
      e,
      l,
      c,
      workbookWithSettings,
      isShowingFormulas,
      a,
    );
    if (!displayedCellText && !c.note && !c.tickmark) continue;
    const cellFontString = buildCellFontString(c),
      m = resolveEffectiveCellStyle(c.style).wrap
        ? displayedCellText.replace(
            /\r\n?/g,
            `
`,
          ).split(`
`)
        : [displayedCellText.replace(/[\t\n\r\f ]+/g, " ")];
    let v = 0;
    for (const b of m)
      v = Math.max(v, measureTextWithTabs(b, cellFontString, textMeasurer));
    const g = clampNumber(
      Math.ceil(v + computeCellTextAlignment(c, textMeasurer)),
      minColumnWidth,
      maxColumnWidth,
    );
    ((d[String(col)] = u.has(col) ? Math.max(d[String(col)], g) : g),
      u.add(col));
  }
  return d;
}
function countWrappedTextLines(
  displayedCellText,
  r,
  cellFontString,
  textMeasurer,
  o,
) {
  let a = 0;
  for (const s of displayedCellText.replace(
    /\r\n?/g,
    `
`,
  ).split(`
`)) {
    const d = Array.from(s);
    if (!d.length) {
      if (++a >= o) return o;
      continue;
    }
    let u = 0;
    for (; u < d.length;) {
      if (++a >= o) return o;
      let l = u + 1,
        c = d.length;
      for (; l < c;) {
        const f = Math.ceil((l + c) / 2);
        measureTextWithTabs(
          d.slice(u, f).join(""),
          cellFontString,
          textMeasurer,
        ) <= r
          ? (l = f)
          : (c = f - 1);
      }
      let h = l;
      if (h >= d.length) break;
      for (let f = h - 1; f > u; f--)
        if (/[\t \-\u200b]/u.test(d[f])) {
          h = f + 1;
          break;
        }
      for (; h < d.length && /[\t ]/u.test(d[h]);) h++;
      u = h;
    }
  }
  return Math.max(1, a);
}
function autoFitRowHeights(
  e,
  selectedAxisIndexes,
  workbookWithSettings,
  textMeasurer,
  isShowingFormulas = false,
  a = "1900",
) {
  const s = validateAutoFitIndices(selectedAxisIndexes, e.rowCount, "row"),
    d = Object.fromEntries(
      [...s].map((l) => [String(l), defaultRowHeightOf(e)]),
    ),
    u = new Set();
  for (const [l, c] of Object.entries(vx(e, workbookWithSettings))) {
    const { row, col } = _t(l);
    if (!s.has(row)) continue;
    const displayedCellText = resolveDisplayedCellText(
      e,
      l,
      c,
      workbookWithSettings,
      isShowingFormulas,
      a,
    );
    if (!displayedCellText && !resolveEffectiveCellStyle(c.style).fontSize)
      continue;
    const cellFontString = buildCellFontString(c),
      g =
        (resolveEffectiveCellStyle(c.style).fontSize
          ? (resolveEffectiveCellStyle(c.style).fontSize * 4) / 3
          : 14) * S2e,
      vertical = measureCellBorderInsets(c.style).vertical,
      y = Math.ceil((maxRowHeight - vertical) / g),
      w = Math.max(
        1,
        columnWidthAt(e, col) - computeCellTextAlignment(c, textMeasurer),
      ),
      E = resolveEffectiveCellStyle(c.style).wrap
        ? countWrappedTextLines(
            displayedCellText,
            w,
            cellFontString,
            textMeasurer,
            y,
          )
        : 1,
      A = clampNumber(
        Math.ceil(Math.max(20, E * g) + vertical),
        minRowHeight,
        maxRowHeight,
      );
    ((d[String(row)] = u.has(row) ? Math.max(d[String(row)], A) : A),
      u.add(row));
  }
  return d;
}
const D2e = 16e6;
function Tp(e, row, col) {
  return (
    row >= e.firstRow &&
    row <= e.lastRow &&
    col >= e.firstCol &&
    col <= e.lastCol
  );
}

function lowerBoundIndex(e, r) {
  let t = 0,
    n = e.length;
  for (; t < n;) {
    const o = (t + n) >>> 1;
    e[o] < r ? (t = o + 1) : (n = o);
  }
  return t;
}
function buildVisibleAxisLayout(e, r, t = {}, n = []) {
  const axisVisibilityIndex = createAxisVisibilityIndex(e, n),
    a = Object.fromEntries(
      Object.entries(t).flatMap(([l, c]) => {
        const h = axisVisibilityIndex.viewIndex(Number(l));
        return h === null ? [] : [[h, c]];
      }),
    );
  e = axisVisibilityIndex.count;
  const s = Object.keys(a)
      .map(Number)
      .filter((l) => l >= 0 && l < e && a[l] > 0 && a[l] !== r)
      .sort((l, c) => l - c),
    d = [0];
  for (const l of s) d.push(d.at(-1) + a[l] - r);
  const offset = (l) => {
    const c = clampNumber(l, 0, e);
    return c * r + d[lowerBoundIndex(s, c)];
  };
  return {
    ...axisVisibilityIndex,
    total: offset(e),
    offset: offset,
    size: (l) => a[l] ?? r,
    indexAt: (l) => {
      let c = 0,
        h = e;
      for (; c < h;) {
        const f = Math.floor((c + h) / 2);
        offset(f + 1) <= l ? (c = f + 1) : (h = f);
      }
      return Math.min(Math.max(0, e - 1), c);
    },
  };
}
function sliceVisibleAxis(e, frozenRowCount) {
  if (
    !Number.isInteger(frozenRowCount) ||
    frozenRowCount < 0 ||
    frozenRowCount > e.count
  )
    throw Error("Choose a valid visible-axis slice.");
  const t = e.offset(frozenRowCount),
    count = e.count - frozenRowCount,
    logicalIndex = (d) =>
      Number.isInteger(d) && d >= 0 && d < count
        ? e.logicalIndex(d + frozenRowCount)
        : -1,
    viewIndex = (d) => {
      const u = e.viewIndex(d);
      return u === null || u < frozenRowCount ? null : u - frozenRowCount;
    },
    countBefore = (d) => Math.max(0, e.countBefore(d) - frozenRowCount);
  return {
    count: count,
    logicalCount: e.logicalCount,
    logicalIndex: logicalIndex,
    viewIndex: viewIndex,
    countBefore: countBefore,
    isHidden: (d) => viewIndex(d) === null,
    next: (d, u, l = false) => {
      const c = d + (l ? 0 : u),
        h =
          u === 1
            ? countBefore(Math.max(0, c))
            : countBefore(Math.min(e.logicalCount, c + 1)) - 1,
        f = logicalIndex(h);
      return f < 0 ? null : f;
    },
    total: e.total - t,
    size: (d) => e.size(d + frozenRowCount),
    offset: (d) =>
      e.offset(Math.max(0, Math.min(count, d)) + frozenRowCount) - t,
    indexAt: (d) =>
      count
        ? Math.max(0, Math.min(count - 1, e.indexAt(d + t) - frozenRowCount))
        : 0,
  };
}
function resolveAxisScale(total, r, zoomScale, n = D2e) {
  const viewport = Math.max(1, r),
    logicalPixels = total * zoomScale,
    physicalPixels = Math.min(logicalPixels, n);
  return {
    viewport: viewport,
    logicalPixels: logicalPixels,
    physicalPixels: physicalPixels,
    logicalMax: Math.max(0, logicalPixels - viewport),
    physicalMax: Math.max(0, physicalPixels - viewport),
  };
}
function physicalToLogicalPixels(e, total, t, zoomScale) {
  const axisScale = resolveAxisScale(total, t, zoomScale);
  return axisScale.physicalMax
    ? ((clampNumber(e, 0, axisScale.physicalMax) / axisScale.physicalMax) *
        axisScale.logicalMax) /
        zoomScale
    : 0;
}
function logicalToPhysicalPixels(logicalOffset, total, t, zoomScale) {
  const axisScale = resolveAxisScale(total, t, zoomScale);
  return axisScale.logicalMax
    ? (clampNumber(logicalOffset * zoomScale, 0, axisScale.logicalMax) /
        axisScale.logicalMax) *
        axisScale.physicalMax
    : 0;
}

export {
  SplitPaneLayout,
  buildVisibleAxisLayout,
  resolveAxisScale,
  useSplitPaneSizes,
  physicalToLogicalPixels,
  autoFitColumnWidths,
  createTextMeasurer,
  Tp,
  sliceVisibleAxis,
  autoFitRowHeights,
  isWithinFormulaSuggestions,
  FormulaSuggestionList,
  logicalToPhysicalPixels,
  resolveCellNavigation,
  FindReplaceDialog,
};
export { clampNumber };
