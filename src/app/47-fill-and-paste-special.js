import { unquoteSheetName, throwWorkbookError } from "./03-workbook-model.js";

import { formatSortRangeAddress } from "./44-clean-and-reshape.js";

import { X, Clipboard, Rows3, Grid2x2 } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
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
  readHiddenAxisIntervals,
  setFreezePanes,
  setSplitPanes,
  findSheetOrThrow,
  defaultColumnWidthOf,
  calculateWorkbookWithSettings,
  _t,
  parseRangeAddress,
  setHiddenAxisRange,
  buildSnapshotGuard,
  columnIndexToLetters,
  applySheetSizeEdits,
  setFreezePanesAtCell,
  yQ,
} from "./03-workbook-model.js";
import {
  buildFillPreview,
  useWorkbookMessages,
  buildFillHandlePreview,
  LocalizedText,
  buildSeriesPreview,
  LocalizedNumberField,
  buildPastePreview,
  applyCellPreview,
} from "./04-cell-input-paste-and-fill.js";
import {
  fillDirectionOptions,
  describeWorkbookError,
} from "./46-workbook-manager.js";

function PasteSpecialDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e);
  a.current = e;
  const s = ne.useRef({
      workbookId: e.workbook.id,
      revision: e.workbook.revision,
      sheetId: e.sheetId,
      range: e.range,
      clipboard: e.clipboard,
    }),
    sheetOrThrow = findSheetOrThrow(e.workbook, e.sheetId),
    rangeAddress = parseRangeAddress(e.range),
    l = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    [c, h] = ne.useState(""),
    [mode, p] = ne.useState("all"),
    [operation, v] = ne.useState("none"),
    [g, b] = ne.useState(false),
    [y, w] = ne.useState(false),
    [E, A] = ne.useState("series"),
    [k, F] = ne.useState("down"),
    [I, O] = ne.useState(
      rangeAddress.firstRow === rangeAddress.lastRow ? "rows" : "columns",
    ),
    [j, z] = ne.useState(false),
    [V, D] = ne.useState("number"),
    [B, _] = ne.useState(""),
    [C, T] = ne.useState("1"),
    [N, R] = ne.useState(""),
    [dateUnit, $] = ne.useState("day"),
    [monthPolicy, G] = ne.useState("preserveDay"),
    [H, Q] = ne.useState("row"),
    [X$1, J] = ne.useState(0),
    [q, W] = ne.useState(0),
    [expectedSnapshot] = ne.useState(() => buildSnapshotGuard(e.workbook)),
    [te, ae] = ne.useState(() => String(defaultRowHeightOf(sheetOrThrow))),
    [oe, de] = ne.useState(() => String(defaultColumnWidthOf(sheetOrThrow))),
    [he, Se] = ne.useState(false),
    [le, ce] = ne.useState(false),
    ge = Object.keys(sheetOrThrow.rowHeights ?? {}).length,
    se = Object.keys(sheetOrThrow.columnWidths ?? {}).length,
    we = e.readOnly || !!sheetOrThrow.protection?.enabled,
    Ee =
      (e.mode === "viewSettings" &&
        JSON.stringify(e.workbook) !== expectedSnapshot.snapshotGuard) ||
      s.current.workbookId !== e.workbook.id ||
      s.current.revision !== e.workbook.revision ||
      s.current.sheetId !== e.sheetId ||
      s.current.range !== e.range ||
      (e.mode === "pasteSpecial" && s.current.clipboard !== e.clipboard),
    Ae =
      "The workbook changed or the selection moved. Close this dialog and reopen it to review the current cells.",
    ye = ne.useMemo(
      () => calculateWorkbookWithSettings(e.workbook),
      [e.workbook],
    ),
    Te = ne.useMemo(() => {
      if (e.mode === "viewSettings") return { plan: null, error: "" };
      try {
        if (Ee) throw new Error(Ae);
        if (e.mode === "pasteSpecial")
          return e.clipboard
            ? {
                plan: buildPastePreview(
                  e.workbook,
                  e.sheetId,
                  e.range,
                  e.clipboard,
                  {
                    mode: mode,
                    operation: operation,
                    transpose: g,
                    skipBlanks: y,
                  },
                ),
                error: "",
              }
            : { plan: null, error: "" };
        if (E === "copy")
          return {
            plan: buildFillPreview(e.workbook, e.sheetId, e.range, k),
            error: "",
          };
        if (!C.trim()) throw new Error("Enter a step value.");
        const Ze = (lt) => (V === "date" ? lt : Number(lt));
        return {
          plan: buildSeriesPreview(e.workbook, e.sheetId, e.range, {
            kind: V,
            direction:
              I === "rows" ? (j ? "left" : "right") : j ? "up" : "down",
            ...(B.trim() ? { start: Ze(B) } : {}),
            step: Number(C),
            ...(N.trim() ? { stop: Ze(N) } : {}),
            dateUnit: dateUnit,
            monthPolicy: monthPolicy,
          }),
          error: "",
        };
      } catch (error) {
        return { plan: null, error: describeWorkbookError(error) };
      }
    }, [
      e.mode,
      e.workbook,
      e.sheetId,
      e.range,
      e.clipboard,
      Ee,
      mode,
      operation,
      g,
      y,
      E,
      k,
      I,
      j,
      V,
      B,
      C,
      N,
      dateUnit,
      monthPolicy,
    ]),
    Pe =
      e.mode === "viewSettings" ? readHiddenAxisIntervals(sheetOrThrow, H) : [],
    qe = Math.max(1, Math.ceil(Pe.length / 20)),
    ht = Math.max(1, Math.ceil((Te.plan?.previewPatches.length ?? 0) / 12)),
    Re = Math.min(q, ht - 1);
  (ne.useEffect(() => {
    W(0);
  }, [Te.plan]),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        o.current
          ?.querySelector("button:not(:disabled),input:not(:disabled)")
          ?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  function Me() {
    const Ze = a.current,
      lt = s.current;
    if (
      Ze.readOnly ||
      findSheetOrThrow(Ze.workbook, Ze.sheetId).protection?.enabled
    )
      throw new Error("This worksheet is read-only.");
    if (
      Ze.workbook.id !== lt.workbookId ||
      (Ze.mode === "viewSettings" &&
        JSON.stringify(Ze.workbook) !== expectedSnapshot.snapshotGuard) ||
      Ze.workbook.revision !== lt.revision ||
      Ze.sheetId !== lt.sheetId ||
      Ze.range !== lt.range ||
      (Ze.mode === "pasteSpecial" && Ze.clipboard !== lt.clipboard)
    )
      throw new Error(Ae);
  }
  function ke(Ze, lt, be) {
    var ft;
    try {
      (h(""), Me());
      const Fe = Ze();
      (JSON.stringify(Fe) !== JSON.stringify(a.current.workbook) &&
        a.current.onWorkbookChange(Fe, lt),
        be && (ft = a.current).onSelectionChange?.call(ft, be),
        a.current.onClose());
    } catch (error) {
      h(describeWorkbookError(error));
    }
  }
  function onClick() {
    const plan = Te.plan;
    if (!plan) {
      h(
        Te.error ||
          "Choose a valid source and target before applying this change.",
      );
      return;
    }
    "sourceRange" in plan
      ? ke(
          () =>
            applyCellPreview(a.current.workbook, plan, {
              readOnly: a.current.readOnly,
            }),
          "Paste Special",
          plan.targetRange,
        )
      : ke(
          () =>
            applyCellPreview(a.current.workbook, plan, {
              readOnly: a.current.readOnly,
            }),
          E === "series" ? "Fill series" : `Fill ${k}`,
          plan.targetRange,
        );
  }
  const $e = (Ze, lt) =>
      H === "row"
        ? `Rows ${Ze + 1}${lt === Ze ? "" : `–${lt + 1}`}`
        : `Columns ${columnIndexToLetters(Ze)}${lt === Ze ? "" : `–${columnIndexToLetters(lt)}`}`,
    Je =
      e.mode === "pasteSpecial"
        ? "Paste Special"
        : e.mode === "fillSeries"
          ? "Fill Series"
          : "Worksheet View",
    ze =
      e.mode === "pasteSpecial"
        ? Clipboard
        : e.mode === "fillSeries"
          ? Rows3
          : Grid2x2;
  return i.jsx("div", {
    className: "et-backdrop",
    children: i.jsxs("section", {
      ref: o,
      className: "et-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": n,
      tabIndex: -1,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(), a.current.onClose()),
          event.key === "Tab")
        ) {
          const be = Array.from(
              o.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ),
            ft = be[0],
            pt = be.at(-1);
          event.shiftKey && document.activeElement === ft
            ? (event.preventDefault(), pt == null || pt.focus())
            : !event.shiftKey &&
              document.activeElement === pt &&
              (event.preventDefault(), ft == null || ft.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "et-header",
          children: [
            i.jsxs("h2", { id: n, children: [i.jsx(ze, { size: 19 }), Je] }),
            i.jsx("button", {
              type: "button",
              "aria-label": text("Close editing tools"),
              onClick: e.onClose,
              lang: language("Close editing tools"),
              "data-ui-label": "Close editing tools",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "et-body",
          children: [
            i.jsxs("p", {
              className: "et-selection",
              children: [
                i.jsx(LocalizedText, { children: "Selection" }),
                " ",
                i.jsxs("strong", {
                  children: [sheetOrThrow.name, "!", e.range],
                }),
              ],
            }),
            we &&
              i.jsx("p", {
                className: "et-warning",
                children: i.jsx(LocalizedText, {
                  children:
                    "This worksheet is read-only. You can review the options.",
                }),
              }),
            Ee &&
              i.jsx("p", {
                className: "et-warning",
                role: "alert",
                children: Ae,
              }),
            e.mode === "pasteSpecial" &&
              i.jsxs(i.Fragment, {
                children: [
                  e.clipboard
                    ? i.jsxs("p", {
                        className: "et-source",
                        children: [
                          i.jsx(LocalizedText, { children: "Copied cells " }),
                          i.jsx("strong", {
                            children: e.clipboard.sourceRange,
                          }),
                          " ·",
                          " ",
                          e.clipboard.rowCount,
                          " ",
                          i.jsx(LocalizedText, { children: "rows × " }),
                          e.clipboard.colCount,
                          " ",
                          i.jsx(LocalizedText, { children: "columns" }),
                        ],
                      })
                    : i.jsx("p", {
                        className: "et-empty",
                        children: i.jsx(LocalizedText, {
                          children: "Copy cells before opening Paste Special.",
                        }),
                      }),
                  i.jsxs("fieldset", {
                    disabled: we || Ee || !e.clipboard,
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, { children: "Paste" }),
                      }),
                      i.jsx("div", {
                        className: "et-radio-grid",
                        children: [
                          ["all", "All", "Cell contents and formatting"],
                          ["values", "Values", "Calculated results"],
                          [
                            "formulas",
                            "Formulas",
                            "Formulas and cell contents",
                          ],
                          ["formats", "Formats", "Formatting only"],
                        ].map(([Ze, lt, be]) =>
                          i.jsxs(
                            "label",
                            {
                              className: "et-radio",
                              children: [
                                i.jsx("input", {
                                  type: "radio",
                                  name: `${n}-paste`,
                                  checked: mode === Ze,
                                  onChange: () => p(Ze),
                                }),
                                i.jsxs("span", {
                                  children: [
                                    lt,
                                    i.jsx("small", { children: be }),
                                  ],
                                }),
                              ],
                            },
                            Ze,
                          ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("fieldset", {
                    disabled: we || Ee || !e.clipboard,
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Operation",
                        }),
                      }),
                      i.jsx("div", {
                        className: "et-radio-line",
                        children: [
                          "none",
                          "add",
                          "subtract",
                          "multiply",
                          "divide",
                        ].map((Ze) =>
                          i.jsxs(
                            "label",
                            {
                              children: [
                                i.jsx("input", {
                                  type: "radio",
                                  name: `${n}-operation`,
                                  checked: operation === Ze,
                                  onChange: () => v(Ze),
                                }),
                                Ze[0].toUpperCase() + Ze.slice(1),
                              ],
                            },
                            Ze,
                          ),
                        ),
                      }),
                      i.jsxs("div", {
                        className: "et-checks",
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: y,
                                onChange: (event) => w(event.target.checked),
                              }),
                              i.jsx(LocalizedText, { children: "Skip blanks" }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: g,
                                onChange: (event) => b(event.target.checked),
                              }),
                              i.jsx(LocalizedText, { children: "Transpose" }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            e.mode === "fillSeries" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("div", {
                    className: "et-tabs",
                    role: "tablist",
                    "aria-label": text("Fill options"),
                    onKeyDown: (event) => {
                      if (
                        !["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                          event.key,
                        )
                      )
                        return;
                      event.preventDefault();
                      const lt =
                        event.key === "Home"
                          ? "series"
                          : event.key === "End" || E === "series"
                            ? "copy"
                            : "series";
                      (A(lt),
                        event.currentTarget
                          .querySelector(`[data-fill-tab="${lt}"]`)
                          ?.focus());
                    },
                    lang: language("Fill options"),
                    "data-ui-label": "Fill options",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        role: "tab",
                        id: `${n}-series`,
                        "data-fill-tab": "series",
                        "aria-controls": `${n}-fill-panel`,
                        tabIndex: E === "series" ? 0 : -1,
                        "aria-selected": E === "series",
                        onClick: () => A("series"),
                        children: i.jsx(LocalizedText, { children: "Series" }),
                      }),
                      i.jsx("button", {
                        type: "button",
                        role: "tab",
                        id: `${n}-copy`,
                        "data-fill-tab": "copy",
                        "aria-controls": `${n}-fill-panel`,
                        tabIndex: E === "copy" ? 0 : -1,
                        "aria-selected": E === "copy",
                        onClick: () => A("copy"),
                        children: i.jsx(LocalizedText, {
                          children: "Copy cells",
                        }),
                      }),
                    ],
                  }),
                  i.jsx("div", {
                    role: "tabpanel",
                    id: `${n}-fill-panel`,
                    "aria-labelledby": `${n}-${E}`,
                    children:
                      E === "copy"
                        ? i.jsxs("fieldset", {
                            disabled: we || Ee,
                            children: [
                              i.jsx("legend", {
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Fill from the edge of the selection",
                                }),
                              }),
                              i.jsx("div", {
                                className: "et-direction-grid",
                                children: fillDirectionOptions.map(
                                  ({ value, label, icon }) =>
                                    i.jsxs(
                                      "button",
                                      {
                                        type: "button",
                                        "aria-pressed": k === value,
                                        onClick: () => F(value),
                                        children: [
                                          i.jsx(icon, { size: 18 }),
                                          label,
                                        ],
                                      },
                                      value,
                                    ),
                                ),
                              }),
                              i.jsx("p", {
                                className: "et-muted",
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Copy the first row or column in the chosen direction. Relative formula references adjust for each destination.",
                                }),
                              }),
                            ],
                          })
                        : i.jsxs("fieldset", {
                            disabled: we || Ee,
                            children: [
                              i.jsx("legend", {
                                children: i.jsx(LocalizedText, {
                                  children: "Series settings",
                                }),
                              }),
                              i.jsxs("div", {
                                className: "et-fields",
                                children: [
                                  i.jsxs("label", {
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Series in",
                                      }),
                                      i.jsxs("select", {
                                        "aria-label": text("Series in"),
                                        value: I,
                                        onChange: (event) =>
                                          O(event.target.value),
                                        lang: language("Series in"),
                                        "data-ui-label": "Series in",
                                        children: [
                                          i.jsx("option", {
                                            value: "rows",
                                            children: "Rows",
                                          }),
                                          i.jsx("option", {
                                            value: "columns",
                                            children: "Columns",
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                                  i.jsxs("label", {
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Order",
                                      }),
                                      i.jsxs("select", {
                                        "aria-label": text("Series order"),
                                        value: j ? "backward" : "forward",
                                        onChange: (event) =>
                                          z(event.target.value === "backward"),
                                        lang: language("Series order"),
                                        "data-ui-label": "Series order",
                                        children: [
                                          i.jsxs("option", {
                                            value: "forward",
                                            children: [
                                              "Forward (",
                                              I === "rows" ? "right" : "down",
                                              ")",
                                            ],
                                          }),
                                          i.jsxs("option", {
                                            value: "backward",
                                            children: [
                                              "Backward (",
                                              I === "rows" ? "left" : "up",
                                              ")",
                                            ],
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                                  i.jsxs("label", {
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Type",
                                      }),
                                      i.jsxs("select", {
                                        "aria-label": text("Series type"),
                                        value: V,
                                        onChange: (event) => {
                                          (D(event.target.value), _(""), R(""));
                                        },
                                        lang: language("Series type"),
                                        "data-ui-label": "Series type",
                                        children: [
                                          i.jsx("option", {
                                            value: "number",
                                            children: "Number",
                                          }),
                                          i.jsx("option", {
                                            value: "date",
                                            children: "Date",
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                                  i.jsxs("label", {
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Start value",
                                      }),
                                      i.jsx("input", {
                                        "aria-label": text("Start value"),
                                        type: V === "date" ? "date" : "number",
                                        step: "any",
                                        value: B,
                                        placeholder: text("Use first cell"),
                                        onChange: (event) =>
                                          _(event.target.value),
                                        lang: language("Start value"),
                                        "data-ui-label": "Start value",
                                      }),
                                    ],
                                  }),
                                  i.jsxs("label", {
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Step value",
                                      }),
                                      i.jsx(LocalizedNumberField, {
                                        "aria-label": text("Step value"),
                                        step: V === "date" ? "1" : "any",
                                        value: C,
                                        onValueChange: (Ze) => T(Ze),
                                        lang: language("Step value"),
                                        "data-ui-label": "Step value",
                                      }),
                                    ],
                                  }),
                                  i.jsxs("label", {
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Stop value (optional)",
                                      }),
                                      i.jsx("input", {
                                        "aria-label": text("Stop value"),
                                        type: V === "date" ? "date" : "number",
                                        step: "any",
                                        value: N,
                                        onChange: (event) =>
                                          R(event.target.value),
                                        lang: language("Stop value"),
                                        "data-ui-label": "Stop value",
                                      }),
                                    ],
                                  }),
                                ],
                              }),
                              V === "date" &&
                                i.jsxs("div", {
                                  className: "et-fields",
                                  children: [
                                    i.jsxs("label", {
                                      children: [
                                        i.jsx(LocalizedText, {
                                          children: "Date unit",
                                        }),
                                        i.jsxs("select", {
                                          "aria-label": text("Date unit"),
                                          value: dateUnit,
                                          onChange: (event) =>
                                            $(event.target.value),
                                          lang: language("Date unit"),
                                          "data-ui-label": "Date unit",
                                          children: [
                                            i.jsx("option", {
                                              value: "day",
                                              children: "Day",
                                            }),
                                            i.jsx("option", {
                                              value: "weekday",
                                              children: "Weekday",
                                            }),
                                            i.jsx("option", {
                                              value: "month",
                                              children: "Month",
                                            }),
                                            i.jsx("option", {
                                              value: "year",
                                              children: "Year",
                                            }),
                                          ],
                                        }),
                                      ],
                                    }),
                                    (dateUnit === "month" ||
                                      dateUnit === "year") &&
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(LocalizedText, {
                                            children: "Month dates",
                                          }),
                                          i.jsxs("select", {
                                            "aria-label": text("Month dates"),
                                            value: monthPolicy,
                                            onChange: (event) =>
                                              G(event.target.value),
                                            lang: language("Month dates"),
                                            "data-ui-label": "Month dates",
                                            children: [
                                              i.jsx("option", {
                                                value: "preserveDay",
                                                children:
                                                  "Preserve day of month",
                                              }),
                                              i.jsx("option", {
                                                value: "endOfMonth",
                                                children: "End of month",
                                              }),
                                            ],
                                          }),
                                        ],
                                      }),
                                  ],
                                }),
                              i.jsx("p", {
                                className: "et-muted",
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Select one row or one column. Leave Start blank to use its first cell in the chosen direction. A stop value leaves the remaining cells unchanged.",
                                }),
                              }),
                            ],
                          }),
                  }),
                ],
              }),
            e.mode === "viewSettings" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("fieldset", {
                    disabled: we || Ee,
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Default row and column sizes",
                        }),
                      }),
                      i.jsx("p", {
                        className: "et-muted",
                        children: i.jsx(LocalizedText, {
                          children:
                            "Set worksheet-wide sizes, including blank rows and columns. Individual sizes stay unchanged unless you choose to reset them below.",
                        }),
                      }),
                      i.jsxs("div", {
                        className: "et-fields",
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Default row height (px)",
                              }),
                              i.jsx(LocalizedNumberField, {
                                "aria-label": text("Default row height (px)"),
                                min: minRowHeight,
                                max: maxRowHeight,
                                step: "any",
                                value: te,
                                onValueChange: (Ze) => ae(Ze),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Default column width (px)",
                              }),
                              i.jsx(LocalizedNumberField, {
                                "aria-label": text("Default column width (px)"),
                                min: minColumnWidth,
                                max: maxColumnWidth,
                                step: "any",
                                value: oe,
                                onValueChange: (Ze) => de(Ze),
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        className: "et-check",
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: he,
                            onChange: (event) => Se(event.target.checked),
                          }),
                          i.jsx(LocalizedText, { children: "Reset " }),
                          ge.toLocaleString(),
                          " ",
                          i.jsx(LocalizedText, {
                            children: "individual row heights to the default",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        className: "et-check",
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: le,
                            onChange: (event) => ce(event.target.checked),
                          }),
                          i.jsx(LocalizedText, { children: "Reset " }),
                          se.toLocaleString(),
                          " ",
                          i.jsx(LocalizedText, {
                            children: "individual column widths to the default",
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        className: "et-size-summary",
                        "aria-label": text("Default size preview"),
                        lang: language("Default size preview"),
                        "data-ui-label": "Default size preview",
                        children: [
                          i.jsxs("p", {
                            children: [
                              (
                                sheetOrThrow.rowCount - (he ? 0 : ge)
                              ).toLocaleString(),
                              " ",
                              i.jsx(LocalizedText, {
                                children: "rows inherit ",
                              }),
                              te || "—",
                              " ",
                              i.jsx(LocalizedText, { children: "px;" }),
                              " ",
                              he ? 0 : ge.toLocaleString(),
                              " ",
                              i.jsx(LocalizedText, {
                                children: "individual heights remain.",
                              }),
                            ],
                          }),
                          i.jsxs("p", {
                            children: [
                              (
                                sheetOrThrow.colCount - (le ? 0 : se)
                              ).toLocaleString(),
                              " ",
                              i.jsx(LocalizedText, {
                                children: "columns inherit ",
                              }),
                              oe || "—",
                              " ",
                              i.jsx(LocalizedText, { children: "px;" }),
                              " ",
                              le ? 0 : se.toLocaleString(),
                              " ",
                              i.jsx(LocalizedText, {
                                children: "individual widths remain.",
                              }),
                            ],
                          }),
                          i.jsx("p", {
                            className: "et-muted",
                            children: i.jsx(LocalizedText, {
                              children:
                                "Row content with wrapping or a larger font may need more display height. Hidden rows and columns remain hidden. Cell values and formulas stay unchanged.",
                            }),
                          }),
                        ],
                      }),
                      i.jsx("button", {
                        type: "button",
                        onClick: () =>
                          ke(
                            () =>
                              applySheetSizeEdits(
                                a.current.workbook,
                                a.current.sheetId,
                                {
                                  rowHeight: Number(te),
                                  columnWidth: Number(oe),
                                  resetRowOverrides: he,
                                  resetColumnOverrides: le,
                                },
                                {
                                  readOnly: a.current.readOnly,
                                  expectedSnapshot: expectedSnapshot,
                                },
                              ),
                            "Set default row and column sizes",
                          ),
                        children: i.jsx(LocalizedText, {
                          children: "Apply sizes",
                        }),
                      }),
                    ],
                  }),
                  i.jsxs("fieldset", {
                    disabled: we || Ee,
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Split panes",
                        }),
                      }),
                      i.jsxs("p", {
                        className: "et-muted",
                        children: [
                          i.jsx(LocalizedText, {
                            children:
                              "Compare separate areas of this worksheet. Lower panes start at row ",
                          }),
                          rangeAddress.firstRow + 1,
                          i.jsx(LocalizedText, {
                            children: "; right panes start at column",
                          }),
                          " ",
                          columnIndexToLetters(rangeAddress.firstCol),
                          i.jsx(LocalizedText, {
                            children:
                              ". Split replaces frozen panes. Drag the divider to resize, or focus it and use arrow keys. F6 switches panes.",
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        className: "et-freeze-grid",
                        children: [
                          ["horizontal", "vertical", "both"].map((Ze) =>
                            i.jsx(
                              "button",
                              {
                                type: "button",
                                onClick: () =>
                                  ke(
                                    () =>
                                      setSplitPanes(e.workbook, e.sheetId, {
                                        row: rangeAddress.firstRow,
                                        col: rangeAddress.firstCol,
                                        ...(Ze !== "vertical"
                                          ? { horizontal: 50 }
                                          : {}),
                                        ...(Ze !== "horizontal"
                                          ? { vertical: 50 }
                                          : {}),
                                      }),
                                    `Split worksheet ${Ze}`,
                                  ),
                                children:
                                  Ze === "horizontal"
                                    ? "Split top and bottom"
                                    : Ze === "vertical"
                                      ? "Split left and right"
                                      : "Split into four panes",
                              },
                              Ze,
                            ),
                          ),
                          i.jsx("button", {
                            type: "button",
                            disabled: !sheetOrThrow.splitPanes,
                            onClick: () =>
                              ke(
                                () => setSplitPanes(e.workbook, e.sheetId),
                                "Remove split panes",
                              ),
                            children: i.jsx(LocalizedText, {
                              children: "Remove split",
                            }),
                          }),
                        ],
                      }),
                      i.jsxs("p", {
                        className: "et-muted",
                        children: [
                          sheetOrThrow.splitPanes
                            ? "Split is active. Panes share cells, formulas, selection, and Undo."
                            : "No split is active.",
                          " ",
                          i.jsx(LocalizedText, {
                            children:
                              "Divider layout is saved in workbook data. Current scroll positions and zoom are local to this editor.",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("fieldset", {
                    disabled: we || Ee,
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Freeze panes",
                        }),
                      }),
                      i.jsx("p", {
                        className: "et-muted",
                        children: i.jsx(LocalizedText, {
                          children:
                            "Keep rows and columns in view while scrolling. The selected cell freezes everything above and to its left.",
                        }),
                      }),
                      i.jsxs("div", {
                        className: "et-freeze-grid",
                        children: [
                          i.jsxs("button", {
                            type: "button",
                            onClick: () =>
                              ke(
                                () =>
                                  setFreezePanesAtCell(
                                    e.workbook,
                                    e.sheetId,
                                    l,
                                  ),
                                `Freeze panes at ${l}`,
                              ),
                            children: [
                              i.jsx(LocalizedText, { children: "Freeze at " }),
                              l,
                            ],
                          }),
                          i.jsx("button", {
                            type: "button",
                            onClick: () =>
                              ke(
                                () =>
                                  setFreezePanes(e.workbook, e.sheetId, 1, 0),
                                "Freeze top row",
                              ),
                            children: i.jsx(LocalizedText, {
                              children: "Freeze top row",
                            }),
                          }),
                          i.jsx("button", {
                            type: "button",
                            onClick: () =>
                              ke(
                                () =>
                                  setFreezePanes(e.workbook, e.sheetId, 0, 1),
                                "Freeze first column",
                              ),
                            children: i.jsx(LocalizedText, {
                              children: "Freeze first column",
                            }),
                          }),
                          i.jsx("button", {
                            type: "button",
                            onClick: () =>
                              ke(
                                () =>
                                  setFreezePanes(e.workbook, e.sheetId, 0, 0),
                                "Unfreeze panes",
                              ),
                            disabled:
                              !sheetOrThrow.frozenRows &&
                              !sheetOrThrow.frozenCols,
                            children: i.jsx(LocalizedText, {
                              children: "Unfreeze panes",
                            }),
                          }),
                        ],
                      }),
                      i.jsxs("p", {
                        className: "et-muted et-current-view",
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Currently frozen: ",
                          }),
                          sheetOrThrow.frozenRows ?? 0,
                          " ",
                          i.jsx(LocalizedText, { children: "rows ·" }),
                          " ",
                          sheetOrThrow.frozenCols ?? 0,
                          " ",
                          i.jsx(LocalizedText, { children: "columns" }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("fieldset", {
                    disabled: we || Ee,
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Hide the selection",
                        }),
                      }),
                      i.jsxs("div", {
                        className: "et-freeze-grid",
                        children: [
                          i.jsx("button", {
                            type: "button",
                            onClick: () =>
                              ke(
                                () =>
                                  setHiddenAxisRange(
                                    e.workbook,
                                    e.sheetId,
                                    "row",
                                    rangeAddress.firstRow,
                                    rangeAddress.lastRow,
                                    true,
                                  ),
                                "Hide selected rows",
                              ),
                            children: i.jsx(LocalizedText, {
                              children: "Hide selected rows",
                            }),
                          }),
                          i.jsx("button", {
                            type: "button",
                            onClick: () =>
                              ke(
                                () =>
                                  setHiddenAxisRange(
                                    e.workbook,
                                    e.sheetId,
                                    "column",
                                    rangeAddress.firstCol,
                                    rangeAddress.lastCol,
                                    true,
                                  ),
                                "Hide selected columns",
                              ),
                            children: i.jsx(LocalizedText, {
                              children: "Hide selected columns",
                            }),
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("section", {
                    className: "et-hidden",
                    "aria-label": text("Hidden rows and columns"),
                    lang: language("Hidden rows and columns"),
                    "data-ui-label": "Hidden rows and columns",
                    children: [
                      i.jsxs("div", {
                        className: "et-hidden-heading",
                        children: [
                          i.jsx("h3", {
                            children: i.jsx(LocalizedText, {
                              children: "Hidden rows and columns",
                            }),
                          }),
                          i.jsxs("select", {
                            "aria-label": text("Hidden items"),
                            value: H,
                            onChange: (event) => {
                              (Q(event.target.value), J(0));
                            },
                            lang: language("Hidden items"),
                            "data-ui-label": "Hidden items",
                            children: [
                              i.jsx("option", {
                                value: "row",
                                children: "Rows",
                              }),
                              i.jsx("option", {
                                value: "column",
                                children: "Columns",
                              }),
                            ],
                          }),
                        ],
                      }),
                      Pe.length
                        ? i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("ul", {
                                children: Pe.slice(
                                  X$1 * 20,
                                  (X$1 + 1) * 20,
                                ).map((Ze) =>
                                  i.jsxs(
                                    "li",
                                    {
                                      children: [
                                        i.jsx("span", {
                                          children: $e(Ze.start, Ze.end),
                                        }),
                                        i.jsx("button", {
                                          type: "button",
                                          "aria-label": `Unhide ${$e(Ze.start, Ze.end).toLowerCase()}`,
                                          disabled: we || Ee,
                                          onClick: () =>
                                            ke(
                                              () =>
                                                setHiddenAxisRange(
                                                  e.workbook,
                                                  e.sheetId,
                                                  H,
                                                  Ze.start,
                                                  Ze.end,
                                                  false,
                                                ),
                                              `Unhide ${H === "row" ? "rows" : "columns"}`,
                                            ),
                                          children: i.jsx(LocalizedText, {
                                            children: "Unhide",
                                          }),
                                        }),
                                      ],
                                    },
                                    `${Ze.start}-${Ze.end}`,
                                  ),
                                ),
                              }),
                              qe > 1 &&
                                i.jsxs("div", {
                                  className: "et-pagination",
                                  children: [
                                    i.jsx("button", {
                                      type: "button",
                                      disabled: !X$1,
                                      onClick: () => J((Ze) => Ze - 1),
                                      children: i.jsx(LocalizedText, {
                                        children: "Previous hidden items",
                                      }),
                                    }),
                                    i.jsxs("span", {
                                      children: [
                                        i.jsx(LocalizedText, {
                                          children: "Page ",
                                        }),
                                        X$1 + 1,
                                        " ",
                                        i.jsx(LocalizedText, {
                                          children: "of ",
                                        }),
                                        qe,
                                      ],
                                    }),
                                    i.jsx("button", {
                                      type: "button",
                                      disabled: X$1 + 1 >= qe,
                                      onClick: () => J((Ze) => Ze + 1),
                                      children: i.jsx(LocalizedText, {
                                        children: "Next hidden items",
                                      }),
                                    }),
                                  ],
                                }),
                              i.jsxs("button", {
                                type: "button",
                                className: "et-unhide-all",
                                disabled: we || Ee,
                                onClick: () =>
                                  ke(
                                    () => yQ(e.workbook, e.sheetId, H),
                                    `Unhide all ${H === "row" ? "rows" : "columns"}`,
                                  ),
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Unhide all ",
                                  }),
                                  H === "row" ? "rows" : "columns",
                                ],
                              }),
                            ],
                          })
                        : i.jsxs("p", {
                            className: "et-muted",
                            children: [
                              i.jsx(LocalizedText, { children: "No " }),
                              H === "row" ? "rows" : "columns",
                              " ",
                              i.jsx(LocalizedText, { children: "are hidden." }),
                            ],
                          }),
                    ],
                  }),
                ],
              }),
            Te.plan &&
              i.jsxs("section", {
                className: "et-preview",
                "aria-label": text("Change preview"),
                lang: language("Change preview"),
                "data-ui-label": "Change preview",
                children: [
                  i.jsxs("h3", {
                    children: [
                      i.jsx(LocalizedText, { children: "Preview" }),
                      " ",
                      i.jsxs("span", {
                        children: [
                          Te.plan.count.toLocaleString(),
                          " ",
                          i.jsx(LocalizedText, { children: "cells to change" }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [sheetOrThrow.name, "!", Te.plan.targetRange],
                  }),
                  Te.plan.dependentCount > 0 &&
                    i.jsxs("p", {
                      className: "et-muted",
                      children: [
                        i.jsx(LocalizedText, {
                          children: "This operation also updates",
                        }),
                        " ",
                        Te.plan.dependentCount.toLocaleString(),
                        " ",
                        i.jsx(LocalizedText, {
                          children:
                            "related cells. Every changed cell is included in the preview pages below.",
                        }),
                      ],
                    }),
                  Te.plan.previewPatches.length > 0
                    ? i.jsxs(i.Fragment, {
                        children: [
                          i.jsx("div", {
                            className: "et-preview-scroll",
                            children: i.jsxs("table", {
                              children: [
                                i.jsx("thead", {
                                  children: i.jsxs("tr", {
                                    children: [
                                      i.jsx("th", {
                                        children: i.jsx(LocalizedText, {
                                          children: "Cell",
                                        }),
                                      }),
                                      i.jsx("th", {
                                        children: i.jsx(LocalizedText, {
                                          children: "Current",
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
                                  children: Te.plan.previewPatches
                                    .slice(Re * 12, (Re + 1) * 12)
                                    .map((patch) => {
                                      const sheetOrThrow = findSheetOrThrow(
                                        e.workbook,
                                        patch.sheetId,
                                      );
                                      return i.jsxs(
                                        "tr",
                                        {
                                          children: [
                                            i.jsx("th", {
                                              children:
                                                patch.sheetId === e.sheetId
                                                  ? patch.address
                                                  : `${sheetOrThrow.name}!${patch.address}`,
                                            }),
                                            i.jsx("td", {
                                              children:
                                                formatNumberWithPattern(
                                                  ye(
                                                    patch.sheetId,
                                                    patch.address,
                                                  ),
                                                  sheetOrThrow.cells[
                                                    patch.address
                                                  ],
                                                  e.workbook.datePolicy
                                                    ?.dateSystem,
                                                ) || "(blank)",
                                            }),
                                            i.jsxs("td", {
                                              children: [
                                                formatNumberWithPattern(
                                                  patch.cell.raw,
                                                  patch.cell,
                                                  e.workbook.datePolicy
                                                    ?.dateSystem,
                                                ) || "(blank)",
                                                JSON.stringify(
                                                  patch.cell.style,
                                                ) !==
                                                  JSON.stringify(
                                                    sheetOrThrow.cells[
                                                      patch.address
                                                    ]?.style,
                                                  ) &&
                                                  i.jsx("small", {
                                                    children: i.jsx(
                                                      LocalizedText,
                                                      {
                                                        children:
                                                          "Formatting changed",
                                                      },
                                                    ),
                                                  }),
                                              ],
                                            }),
                                          ],
                                        },
                                        `${patch.sheetId}:${patch.address}`,
                                      );
                                    }),
                                }),
                              ],
                            }),
                          }),
                          ht > 1 &&
                            i.jsxs("div", {
                              className: "et-pagination",
                              "aria-label": text("Preview pages"),
                              lang: language("Preview pages"),
                              "data-ui-label": "Preview pages",
                              children: [
                                i.jsx("button", {
                                  type: "button",
                                  disabled: Re === 0,
                                  onClick: () => W((Ze) => Ze - 1),
                                  children: i.jsx(LocalizedText, {
                                    children: "Previous changes",
                                  }),
                                }),
                                i.jsxs("span", {
                                  children: [
                                    i.jsx(LocalizedText, { children: "Page " }),
                                    Re + 1,
                                    " ",
                                    i.jsx(LocalizedText, { children: "of " }),
                                    ht,
                                  ],
                                }),
                                i.jsx("button", {
                                  type: "button",
                                  disabled: Re + 1 >= ht,
                                  onClick: () => W((Ze) => Ze + 1),
                                  children: i.jsx(LocalizedText, {
                                    children: "Next changes",
                                  }),
                                }),
                              ],
                            }),
                        ],
                      })
                    : i.jsx("p", {
                        className: "et-muted",
                        children: i.jsx(LocalizedText, {
                          children:
                            "The selected cells already match these settings.",
                        }),
                      }),
                ],
              }),
            !Ee &&
              (c || Te.error) &&
              i.jsx("p", {
                className: "et-error",
                role: "alert",
                children: c || Te.error,
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "et-footer",
          children: [
            i.jsx("button", {
              type: "button",
              onClick: e.onClose,
              children: e.mode === "viewSettings" ? "Done" : "Cancel",
            }),
            e.mode !== "viewSettings" &&
              i.jsx("button", {
                type: "button",
                className: "et-primary",
                disabled: we || Ee || !Te.plan || Te.plan.count === 0,
                onClick: onClick,
                children:
                  e.mode === "pasteSpecial"
                    ? "Paste"
                    : E === "series"
                      ? "Fill series"
                      : "Apply fill",
              }),
          ],
        }),
      ],
    }),
  });
}
function FillPreviewDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e.workbook),
    s = ne.useRef(e);
  s.current = e;
  const [mode, u] = ne.useState("copy"),
    [l, c] = ne.useState(""),
    [h, f] = ne.useState(0),
    p = e.workbook !== a.current,
    m = ne.useMemo(() => {
      try {
        return {
          plan: buildFillHandlePreview(
            e.workbook,
            e.sheetId,
            e.sourceRange,
            e.targetRange,
            {
              mode: mode,
            },
          ),
          error: "",
        };
      } catch (error) {
        return {
          plan: null,
          error:
            error instanceof Error
              ? error.message
              : "The fill cannot be applied.",
        };
      }
    }, [e.workbook, e.sheetId, e.sourceRange, e.targetRange, mode]),
    plan = m.plan;
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      o.current?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const disabled =
    e.readOnly ||
    !!findSheetOrThrow(e.workbook, e.sheetId).protection?.enabled ||
    p;
  return i.jsx("div", {
    className: "et-backdrop",
    children: i.jsxs("section", {
      ref: o,
      className: "et-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": n,
      tabIndex: -1,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.preventDefault(), e.onClose()),
          event.key === "Tab")
        ) {
          const E = Array.from(
              o.current?.querySelectorAll(
                "button:not(:disabled),select:not(:disabled),input:not(:disabled)",
              ) ?? [],
            ),
            A = E[0],
            k = E.at(-1);
          event.shiftKey &&
          (document.activeElement === A || document.activeElement === o.current)
            ? (event.preventDefault(), k == null || k.focus())
            : !event.shiftKey &&
              (document.activeElement === k ||
                document.activeElement === o.current) &&
              (event.preventDefault(), A == null || A.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "et-header",
          children: [
            i.jsx("h2", {
              id: n,
              children: i.jsx(LocalizedText, { children: "Fill selection" }),
            }),
            i.jsx("button", {
              "aria-label": text("Close fill preview"),
              onClick: e.onClose,
              lang: language("Close fill preview"),
              "data-ui-label": "Close fill preview",
              children: "×",
            }),
          ],
        }),
        i.jsxs("div", {
          className: "et-body",
          children: [
            i.jsxs("p", {
              children: [
                i.jsx(LocalizedText, { children: "Extend " }),
                i.jsx("strong", { children: e.sourceRange }),
                " ",
                i.jsx(LocalizedText, { children: "to" }),
                " ",
                i.jsx("strong", { children: e.targetRange }),
                i.jsx(LocalizedText, {
                  children:
                    ". Fill writes into the extension; dependent formula updates are included in the preview.",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Fill method" }),
                i.jsxs("select", {
                  "aria-label": text("Fill method"),
                  value: mode,
                  disabled: disabled,
                  onChange: (event) => {
                    (u(event.target.value), f(0), c(""));
                  },
                  lang: language("Fill method"),
                  "data-ui-label": "Fill method",
                  children: [
                    i.jsx("option", {
                      value: "copy",
                      children: "Copy cells and relative formulas",
                    }),
                    i.jsx("option", {
                      value: "series",
                      children: "Extend two number seeds",
                    }),
                  ],
                }),
              ],
            }),
            i.jsx("p", {
              className: "et-muted",
              children: i.jsx(LocalizedText, {
                children:
                  "Copy repeats the selected pattern. Number series requires two numbers in one row or column. Use Home → Fill Series for calendar dates.",
              }),
            }),
            plan &&
              i.jsxs("section", {
                className: "et-preview",
                "aria-label": text("Fill preview"),
                lang: language("Fill preview"),
                "data-ui-label": "Fill preview",
                children: [
                  i.jsxs("h3", {
                    children: [
                      plan.count,
                      " ",
                      i.jsx(LocalizedText, { children: "cells to change" }),
                    ],
                  }),
                  i.jsx("p", { children: plan.targetRange }),
                  i.jsx("div", {
                    className: "et-preview-scroll",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Cell",
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
                            .slice(h * 25, (h + 1) * 25)
                            .map((patch) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("th", {
                                      children:
                                        patch.sheetId === e.sheetId
                                          ? patch.address
                                          : `${findSheetOrThrow(e.workbook, patch.sheetId).name}!${patch.address}`,
                                    }),
                                    i.jsx("td", {
                                      children:
                                        formatNumberWithPattern(
                                          patch.cell.raw,
                                          patch.cell,
                                          e.workbook.datePolicy?.dateSystem,
                                        ) || "(blank)",
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
                  plan.count > 25 &&
                    i.jsxs("div", {
                      className: "et-pagination",
                      children: [
                        i.jsx("button", {
                          disabled: !h,
                          onClick: () => f(h - 1),
                          children: i.jsx(LocalizedText, {
                            children: "Previous changes",
                          }),
                        }),
                        i.jsxs("span", {
                          children: [
                            i.jsx(LocalizedText, { children: "Page " }),
                            h + 1,
                            " ",
                            i.jsx(LocalizedText, { children: "of " }),
                            Math.ceil(plan.count / 25),
                          ],
                        }),
                        i.jsx("button", {
                          disabled: (h + 1) * 25 >= plan.count,
                          onClick: () => f(h + 1),
                          children: i.jsx(LocalizedText, {
                            children: "Next changes",
                          }),
                        }),
                      ],
                    }),
                ],
              }),
            (l || m.error || p) &&
              i.jsx("p", {
                className: "et-error",
                role: "alert",
                children: p
                  ? "The workbook changed. Cancel and review the fill again."
                  : l || m.error,
              }),
            e.readOnly &&
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children: "This workbook is read-only.",
                }),
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "et-footer",
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "et-primary",
              disabled: disabled || !plan || plan.count === 0,
              onClick: () => {
                try {
                  const y = s.current;
                  if (y.workbook !== a.current)
                    throw Error("The workbook changed. Review the fill again.");
                  if (!plan) throw Error("Choose a valid fill first.");
                  const w = applyCellPreview(y.workbook, plan, {
                    readOnly: y.readOnly,
                  });
                  (y.onWorkbookChange(w, "Fill handle"),
                    y.onSelectionChange(y.targetRange),
                    y.onClose());
                } catch (error) {
                  c(
                    error instanceof Error
                      ? error.message
                      : "The fill could not be applied.",
                  );
                }
              },
              children: i.jsx(LocalizedText, { children: "Apply fill" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const pasteOperationLimits = {
  sourceCells: 500,
  changedCells: 500,
  metadataChanges: 500,
  formulaCells: 1e4,
  formulaCharacters: 1e7,
};

const rangeContainsRange = (e, r) =>
    e.firstRow <= r.firstRow &&
    e.lastRow >= r.lastRow &&
    e.firstCol <= r.firstCol &&
    e.lastCol >= r.lastCol,
  hasCellContent = (cell) =>
    !!cell && (cell.raw !== null || Object.keys(cell).some((r) => r !== "raw"));

function assertRangeWithinBoundary(sheetOrThrow, targetRange) {
  const rangeAddress = parseRangeAddress(targetRange);
  return (
    (rangeAddress.lastRow >= sheetOrThrow.rowCount ||
      rangeAddress.lastCol >= sheetOrThrow.colCount) &&
      throwWorkbookError("The range exceeds the worksheet boundary."),
    rangeAddress
  );
}
function assertCutSourceWithinLimit(sheetOrThrow, targetRange) {
  const t = assertRangeWithinBoundary(sheetOrThrow, targetRange);
  return (
    (t.lastRow - t.firstRow + 1) * (t.lastCol - t.firstCol + 1) >
      pasteOperationLimits.sourceCells &&
      throwWorkbookError("Cut and move supports at most 500 source cells."),
    t
  );
}
function createCutSourceSnapshot(
  activeWorkbook,
  sheetId,
  activeSelectionRange,
) {
  return (
    assertCutSourceWithinLimit(
      findSheetOrThrow(activeWorkbook, sheetId),
      activeSelectionRange,
    ),
    Object.freeze({
      version: 1,
      workbookId: activeWorkbook.id,
      revision: activeWorkbook.revision,
      snapshotGuard: JSON.stringify(activeWorkbook),
      sheetId: sheetId,
      range: formatSortRangeAddress(activeSelectionRange),
    })
  );
}
function verifyCutSourceSnapshot(book, r, sourceSheetId, sourceRange) {
  (!r ||
    r.version !== 1 ||
    r.workbookId !== book.id ||
    r.revision !== book.revision ||
    r.sheetId !== sourceSheetId ||
    r.range !== formatSortRangeAddress(sourceRange) ||
    r.snapshotGuard !== JSON.stringify(book)) &&
    throwWorkbookError(
      "The workbook changed after Cut. Cut the source again before moving it.",
    );
}
const isUnsafeStructuredReference = (e) =>
    e.kind === "structured" && (!e.name || /\[\s*@|#This\s+Row/i.test(e.text)),
  sheetNameReferencePattern = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*)`,
  sheetNamePrefixPattern = new RegExp(`^(${sheetNameReferencePattern})!`),
  sheetRangePrefixPattern = new RegExp(
    `^(${sheetNameReferencePattern}):(${sheetNameReferencePattern})!`,
  ),
  identifierCharPattern = /[A-Za-z0-9_.\\\u0080-\uffff]/;
function assertBalancedStructuredReference(refersTo, end) {
  let t = 0,
    n = end;
  do (refersTo[n] === "[" && t++, refersTo[n] === "]" && t--, n++);
  while (n < refersTo.length && t);
  return (
    t &&
      throwWorkbookError(
        "An unbalanced structured reference prevents a safe move.",
      ),
    n
  );
}
function assertFormulaReferencesMovable(refersTo) {
  const r = [];
  let t = 0;
  for (; t < refersTo.length;) {
    const start = t,
      o = refersTo[t];
    if (o === '"') {
      for (t++; t < refersTo.length;)
        if (refersTo[t++] === '"')
          if (refersTo[t] === '"') t++;
          else break;
      continue;
    }
    const a = refersTo.slice(t),
      s = t === 0 || !identifierCharPattern.test(refersTo[t - 1]);
    if (
      (s &&
        sheetRangePrefixPattern.test(a) &&
        throwWorkbookError(
          "Three-dimensional worksheet references must be replaced with explicit local references before moving cells.",
        ),
      o === "#")
    ) {
      const p = /^#[A-Za-z0-9/?]+[!?]?/.exec(a);
      if (p) {
        t += p[0].length;
        continue;
      }
    }
    let sheet,
      u = 0;
    const l = s ? sheetNamePrefixPattern.exec(a) : null;
    l &&
      ((sheet = unquoteSheetName(l[1])),
      /[\[\]:]/.test(sheet) &&
        throwWorkbookError(
          "External or three-dimensional references prevent a safe range move.",
        ),
      (u = l[0].length));
    const c = a.slice(u),
      h = s
        ? /^(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)/.exec(
            c,
          )
        : null;
    if (
      h &&
      !identifierCharPattern.test(c[h[0].length] ?? "") &&
      !/^\s*[(!]/.test(c.slice(h[0].length))
    ) {
      const end = t + u + h[0].length;
      (refersTo[end] === "#" &&
        throwWorkbookError(
          "Dynamic spill references require an explicit stable range before moving cells.",
        ),
        refersTo[end] === ":" &&
          throwWorkbookError(
            "Separately qualified range endpoints are not supported for a safe move.",
          ),
        r.push({
          start: start,
          end: end,
          text: refersTo.slice(start, end),
          kind: "reference",
          sheet: sheet,
          reference: h[0],
        }),
        (t = end));
      continue;
    }
    const f = s ? /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(c) : null;
    if (f) {
      const end = t + u + f[0].length,
        m = refersTo.slice(end);
      if (refersTo[end] === "[") {
        const v = assertBalancedStructuredReference(refersTo, end);
        (refersTo[v] === "!" &&
          throwWorkbookError(
            "External workbook references prevent a safe move.",
          ),
          r.push({
            start: start,
            end: v,
            text: refersTo.slice(start, v),
            kind: "structured",
            sheet: sheet,
            name: f[0],
          }),
          (t = v));
        continue;
      }
      if (/^\s*\(/.test(m)) {
        (/^(?:_xlfn\.)?(?:INDIRECT|OFFSET|EVALUATE|HYPERLINK)$/i.test(f[0]) &&
          r.push({
            start: start,
            end: end,
            text: f[0],
            kind: "dynamic",
            name: f[0],
          }),
          (t = end));
        continue;
      }
      (r.push({
        start: start,
        end: end,
        text: refersTo.slice(start, end),
        kind: "name",
        sheet: sheet,
        name: f[0],
      }),
        (t = end));
      continue;
    }
    if (o === "[") {
      const end = assertBalancedStructuredReference(refersTo, t);
      (/^\s*(?:[^!+*/(),;]+)!/.test(refersTo.slice(end)) &&
        throwWorkbookError("External workbook references prevent a safe move."),
        r.push({
          start: start,
          end: end,
          text: refersTo.slice(start, end),
          kind: "structured",
        }),
        (t = end));
      continue;
    }
    if (
      (l &&
        throwWorkbookError(
          "A worksheet-qualified expression is not a supported reference for a safe move.",
        ),
      o === "'")
    ) {
      const p = /^'(?:[^']|'')*'/.exec(a);
      if (p) {
        (/\[[^\]]+\]/.test(p[0]) &&
          refersTo[t + p[0].length] === "!" &&
          throwWorkbookError(
            "External workbook references prevent a safe move.",
          ),
          (t += p[0].length));
        continue;
      }
    }
    t++;
  }
  return r;
}
function parseWholeColumnReference(reference) {
  if (/^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/.test(reference)) {
    const [r, t] = reference.split(":"),
      n = _t(`${r}1`),
      o = _t(`${t}1`);
    return {
      firstRow: 0,
      lastRow: workbookLimits.maxRows - 1,
      firstCol: Math.min(n.col, o.col),
      lastCol: Math.max(n.col, o.col),
    };
  }
  if (/^\$?\d+:\$?\d+$/.test(reference)) {
    const [r, t] = reference.replace(/\$/g, "").split(":").map(Number);
    return (
      (r > workbookLimits.maxRows || t > workbookLimits.maxRows) &&
        throwWorkbookError("A reference exceeds Excel's row boundary."),
      {
        firstRow: Math.min(r, t) - 1,
        lastRow: Math.max(r, t) - 1,
        firstCol: 0,
        lastCol: workbookLimits.maxCols - 1,
      }
    );
  }
  return parseRangeAddress(reference);
}
function parseMovableA1Rectangle(reference, r, t) {
  return reference
    .split(":")
    .map((n) => {
      const o = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(n);
      o ||
        throwWorkbookError(
          "Only complete A1 cell rectangles can move to a new location.",
        );
      const a = _t(n);
      return `${o[1]}${columnIndexToLetters(a.col + t)}${o[3]}${a.row + r + 1}`;
    })
    .join(":");
}

export {
  pasteOperationLimits,
  hasCellContent,
  parseWholeColumnReference,
  assertCutSourceWithinLimit,
  createCutSourceSnapshot,
  rangeContainsRange,
  parseMovableA1Rectangle,
  FillPreviewDialog,
  assertFormulaReferencesMovable,
  verifyCutSourceSnapshot,
  PasteSpecialDialog,
  isUnsafeStructuredReference,
  assertRangeWithinBoundary,
};
