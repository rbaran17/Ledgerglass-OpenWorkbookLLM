import { rangesOverlap } from "./15-data-model-dialog.js";
import { countRangeCells } from "./08-xlsx-import-export-and-preservation.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { supportedFontFamilies } from "./02-text-and-financial-functions.js";
import {
  assertConditionalFormatPreviewCurrent,
  formatNumberWithPattern,
  conditionalRuleKinds,
  cellStyleToCss,
  normalizeConditionalRules,
  sortConditionalRulesByPriority,
  applyBorderEditTransaction,
  updateConditionalFormatRules,
  findSheetOrThrow,
  textWrapToCss,
  calculateWorkbookWithSettings,
  isCellAddressInRanges,
  parseRangeAddress,
  parseSelectionAreas,
  assertSheetUnprotected,
  mergeWorksheetCells,
  validateCellStyle,
  buildBorderPresetStyles,
  normalizeWorkbookInput,
  cellBorderEdges,
  borderLineStyles,
  columnIndexToLetters,
  countSelectionCells,
  resolveEffectiveCellStyle,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  expandShortHexColor,
  sampleGridAddresses,
  CellValuePreview,
  cellFormatTabs,
} from "./22-themes-and-cell-styles.js";

function FormatCellsDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const sheet = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    selectionAreas = parseSelectionAreas(e.ranges ?? [e.range]),
    rangeAddress = parseRangeAddress(e.range),
    l =
      e.active && isCellAddressInRanges(selectionAreas, e.active)
        ? e.active
        : `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    c = ne.useRef({
      book: e.workbook.id,
      revision: e.workbook.revision,
      content: JSON.stringify(e.workbook),
      sheet: e.sheetId,
      selection: JSON.stringify(e.ranges ?? [e.range]),
      active: e.active,
    }),
    [h] = ne.useState(() =>
      validateCellStyle(resolveEffectiveCellStyle(sheet?.cells[l]?.style)),
    ),
    [f, p] = ne.useState(
      e.mode === "alignment"
        ? "Alignment"
        : e.mode === "borders"
          ? "Border"
          : "Font",
    ),
    [m, v] = ne.useState({}),
    [g, b] = ne.useState(""),
    [preset, w] = ne.useState(),
    [line, A] = ne.useState({ style: "thin", color: "#000000" }),
    [k, F] = ne.useState(h.fontSize === undefined ? "" : String(h.fontSize)),
    [I, O] = ne.useState(String(h.indent ?? 0)),
    j = countSelectionCells(selectionAreas),
    z = ne.useMemo(
      () => calculateWorkbookWithSettings(e.workbook),
      [e.workbook],
    ),
    V = sheet ? z(sheet.id, l) : null,
    D =
      c.current.book !== e.workbook.id ||
      c.current.revision !== e.workbook.revision ||
      c.current.content !== JSON.stringify(e.workbook) ||
      c.current.sheet !== e.sheetId ||
      c.current.selection !== JSON.stringify(e.ranges ?? [e.range]) ||
      c.current.active !== e.active,
    B = e.readOnly || D || !!sheet?.protection?.enabled || !sheet || j > 500;
  let _ = h,
    C = "";
  try {
    _ = validateCellStyle({
      ...h,
      ...m,
      ...(m.borders ? { borders: { ...h.borders, ...m.borders } } : {}),
    });
  } catch (error) {
    C =
      error instanceof Error ? error.message : "Choose supported cell styles.";
  }
  function T(H) {
    (v((Q) => ({ ...Q, ...H })), b(""));
  }
  function N(H, Q = false) {
    (v((X) => ({
      ...X,
      borders: { ...X.borders, [H]: Q ? { style: "none" } : { ...line } },
    })),
      b(""));
  }
  function R(H) {
    (w(H),
      v((Q) => {
        const X = { ...Q };
        return (delete X.borders, X);
      }),
      b(""));
  }
  const L = (address) => ({
    ..._,
    borders: {
      ...h.borders,
      ...(preset
        ? buildBorderPresetStyles(address, sampleGridAddresses, preset, line)
        : {}),
      ...m.borders,
    },
  });
  return (
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        o.current?.querySelector('[role="tab"][aria-selected="true"]')?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []),
    i.jsx("div", {
      className: "cs-backdrop",
      children: i.jsxs("div", {
        ref: o,
        className: "cs-dialog",
        role: "dialog",
        "aria-modal": "true",
        "aria-labelledby": `${n}-title`,
        onKeyDown: (event) => {
          if (
            (event.stopPropagation(),
            event.key === "Escape" && (event.preventDefault(), a.current()),
            event.key === "Tab")
          ) {
            const X = [
                ...(o.current?.querySelectorAll(
                  'button:not(:disabled):not([tabindex="-1"]),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
                ) ?? []),
              ],
              J = X[0],
              q = X.at(-1);
            event.shiftKey && document.activeElement === J
              ? (event.preventDefault(), q == null || q.focus())
              : !event.shiftKey &&
                document.activeElement === q &&
                (event.preventDefault(), J == null || J.focus());
          }
        },
        children: [
          i.jsxs("header", {
            className: "cs-header",
            children: [
              i.jsx("h2", {
                id: `${n}-title`,
                children: i.jsx(LocalizedText, { children: "Format Cells" }),
              }),
              i.jsx("button", {
                "aria-label": text("Close cell styles"),
                onClick: e.onClose,
                lang: language("Close cell styles"),
                "data-ui-label": "Close cell styles",
                children: i.jsx(X, { size: 18 }),
              }),
            ],
          }),
          i.jsx("div", {
            className: "cs-tabs",
            role: "tablist",
            "aria-label": text("Cell style settings"),
            lang: language("Cell style settings"),
            "data-ui-label": "Cell style settings",
            children: cellFormatTabs.map((cellFormatTab, Q) =>
              i.jsx(
                "button",
                {
                  role: "tab",
                  id: `${n}-${cellFormatTab}`,
                  "aria-controls": `${n}-panel`,
                  "aria-selected": f === cellFormatTab,
                  tabIndex: f === cellFormatTab ? 0 : -1,
                  onClick: () => p(cellFormatTab),
                  onKeyDown: (event) => {
                    if (
                      ["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                        event.key,
                      )
                    ) {
                      event.preventDefault();
                      const W =
                        event.key === "Home"
                          ? 0
                          : event.key === "End"
                            ? cellFormatTabs.length - 1
                            : (Q +
                                (event.key === "ArrowRight"
                                  ? 1
                                  : cellFormatTabs.length - 1)) %
                              cellFormatTabs.length;
                      (p(cellFormatTabs[W]),
                        event.currentTarget.parentElement?.children[
                          W
                        ]?.focus());
                    }
                  },
                  children: cellFormatTab,
                },
                cellFormatTab,
              ),
            ),
          }),
          i.jsxs("form", {
            onSubmit: (event) => {
              (event.preventDefault(), b(""));
              try {
                if (e.readOnly) throw Error("This workbook is read-only.");
                if (D)
                  throw Error(
                    "The workbook or selection changed. Close and reopen Format Cells.",
                  );
                if (C) throw Error(C);
                const Q = applyBorderEditTransaction(
                  e.workbook,
                  e.sheetId,
                  selectionAreas,
                  m,
                  {
                    readOnly: e.readOnly,
                    ...(preset
                      ? { borders: { preset: preset, line: line } }
                      : {}),
                  },
                );
                (Q !== e.workbook &&
                  e.onWorkbookChange(Q, "Format cell styles"),
                  e.onClose());
              } catch (error) {
                b(
                  error instanceof Error
                    ? error.message
                    : "The cell styles could not be applied.",
                );
              }
            },
            children: [
              i.jsxs("div", {
                className: "cs-body",
                children: [
                  i.jsxs("p", {
                    className: "cs-selection",
                    children: [
                      j.toLocaleString(),
                      " ",
                      i.jsx(LocalizedText, { children: "selected cells · " }),
                      selectionAreas.length,
                      " ",
                      selectionAreas.length === 1 ? "area" : "areas",
                      i.jsx(LocalizedText, {
                        children:
                          ". Only changed controls apply; existing differences are preserved.",
                      }),
                    ],
                  }),
                  e.readOnly &&
                    i.jsx("p", {
                      className: "cs-notice",
                      children: i.jsx(LocalizedText, {
                        children:
                          "This workbook is read-only. You can inspect style previews.",
                      }),
                    }),
                  sheet?.protection?.enabled &&
                    i.jsx("p", {
                      className: "cs-notice",
                      children: i.jsx(LocalizedText, {
                        children:
                          "This worksheet is protected. Unprotect it before changing styles.",
                      }),
                    }),
                  D &&
                    i.jsx("p", {
                      className: "cs-error",
                      role: "alert",
                      children: i.jsx(LocalizedText, {
                        children:
                          "The workbook or selection changed. Close and reopen Format Cells.",
                      }),
                    }),
                  j > 500 &&
                    i.jsx("p", {
                      className: "cs-error",
                      role: "alert",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Select at most 500 unique cells to apply styles.",
                      }),
                    }),
                  i.jsxs("div", {
                    className: "cs-panel",
                    role: "tabpanel",
                    id: `${n}-panel`,
                    "aria-labelledby": `${n}-${f}`,
                    children: [
                      f === "Font" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("div", {
                              className: "cs-fields",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, { children: "Font" }),
                                    i.jsxs("select", {
                                      "aria-label": text("Font family"),
                                      value: _.fontFamily ?? "",
                                      onChange: (event) =>
                                        T({
                                          fontFamily:
                                            event.target.value || undefined,
                                        }),
                                      lang: language("Font family"),
                                      "data-ui-label": "Font family",
                                      children: [
                                        i.jsx("option", {
                                          value: "",
                                          children: "Default (Calibri)",
                                        }),
                                        supportedFontFamilies.map(
                                          (supportedFontFamily) =>
                                            i.jsx(
                                              "option",
                                              { children: supportedFontFamily },
                                              supportedFontFamily,
                                            ),
                                        ),
                                      ],
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, { children: "Size" }),
                                    i.jsx(LocalizedNumberField, {
                                      allowEmpty: true,
                                      "aria-label": text("Font size"),
                                      min: 8,
                                      max: 72,
                                      step: 1,
                                      placeholder: text("Default (14 px)"),
                                      value: k,
                                      onValueChange: (H) => {
                                        (F(H),
                                          T({
                                            fontSize:
                                              H === "" ? undefined : Number(H),
                                          }));
                                      },
                                      lang: language("Font size"),
                                      "data-ui-label": "Font size",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              className: "cs-effects",
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Effects",
                                  }),
                                }),
                                [
                                  ["bold", "Bold"],
                                  ["italic", "Italic"],
                                  ["underline", "Underline"],
                                  ["strikethrough", "Strikethrough"],
                                ].map(([H, Q]) =>
                                  i.jsxs(
                                    "label",
                                    {
                                      children: [
                                        i.jsx("input", {
                                          type: "checkbox",
                                          "aria-label": Q,
                                          checked: _[H] ?? false,
                                          onChange: (event) =>
                                            T({ [H]: event.target.checked }),
                                        }),
                                        Q,
                                      ],
                                    },
                                    H,
                                  ),
                                ),
                              ],
                            }),
                            i.jsxs("div", {
                              className: "cs-fields",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Font color",
                                    }),
                                    i.jsxs("span", {
                                      className: "cs-color",
                                      children: [
                                        i.jsx("input", {
                                          type: "color",
                                          "aria-label": text("Font color"),
                                          value: expandShortHexColor(
                                            _.color,
                                            "#000000",
                                          ),
                                          onChange: (event) =>
                                            T({ color: event.target.value }),
                                          lang: language("Font color"),
                                          "data-ui-label": "Font color",
                                        }),
                                        i.jsx("button", {
                                          type: "button",
                                          onClick: () =>
                                            T({ color: undefined }),
                                          children: i.jsx(LocalizedText, {
                                            children: "Automatic",
                                          }),
                                        }),
                                      ],
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Fill color",
                                    }),
                                    i.jsxs("span", {
                                      className: "cs-color",
                                      children: [
                                        i.jsx("input", {
                                          type: "color",
                                          "aria-label": text("Fill color"),
                                          value: expandShortHexColor(
                                            _.fill,
                                            "#ffffff",
                                          ),
                                          onChange: (event) =>
                                            T({ fill: event.target.value }),
                                          lang: language("Fill color"),
                                          "data-ui-label": "Fill color",
                                        }),
                                        i.jsx("button", {
                                          type: "button",
                                          onClick: () => T({ fill: undefined }),
                                          children: i.jsx(LocalizedText, {
                                            children: "No fill",
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
                      f === "Alignment" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("div", {
                              className: "cs-fields",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Horizontal",
                                    }),
                                    i.jsxs("select", {
                                      "aria-label": text(
                                        "Horizontal alignment",
                                      ),
                                      value: _.align ?? "",
                                      onChange: (event) =>
                                        T({
                                          align:
                                            event.target.value || undefined,
                                        }),
                                      lang: language("Horizontal alignment"),
                                      "data-ui-label": "Horizontal alignment",
                                      children: [
                                        i.jsx("option", {
                                          value: "",
                                          children: "General",
                                        }),
                                        i.jsx("option", {
                                          value: "left",
                                          children: "Left",
                                        }),
                                        i.jsx("option", {
                                          value: "center",
                                          children: "Center",
                                        }),
                                        i.jsx("option", {
                                          value: "right",
                                          children: "Right",
                                        }),
                                      ],
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Vertical",
                                    }),
                                    i.jsxs("select", {
                                      "aria-label": text("Vertical alignment"),
                                      value: _.verticalAlign ?? "",
                                      onChange: (event) =>
                                        T({
                                          verticalAlign:
                                            event.target.value || undefined,
                                        }),
                                      lang: language("Vertical alignment"),
                                      "data-ui-label": "Vertical alignment",
                                      children: [
                                        i.jsx("option", {
                                          value: "",
                                          children: "Default",
                                        }),
                                        i.jsx("option", {
                                          value: "top",
                                          children: "Top",
                                        }),
                                        i.jsx("option", {
                                          value: "middle",
                                          children: "Center",
                                        }),
                                        i.jsx("option", {
                                          value: "bottom",
                                          children: "Bottom",
                                        }),
                                      ],
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Indent",
                                    }),
                                    i.jsx(LocalizedNumberField, {
                                      "aria-label": text("Indent level"),
                                      min: 0,
                                      max: 15,
                                      step: 1,
                                      value: I,
                                      onValueChange: (H) => {
                                        (O(H),
                                          T({
                                            indent: H === "" ? NaN : Number(H),
                                          }));
                                      },
                                      lang: language("Indent level"),
                                      "data-ui-label": "Indent level",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              className: "cs-check",
                              children: [
                                i.jsx("input", {
                                  type: "checkbox",
                                  "aria-label": text("Wrap text"),
                                  checked: _.wrap ?? false,
                                  onChange: (event) =>
                                    T({ wrap: event.target.checked }),
                                  lang: language("Wrap text"),
                                  "data-ui-label": "Wrap text",
                                }),
                                i.jsx(LocalizedText, { children: "Wrap text" }),
                              ],
                            }),
                            i.jsx("p", {
                              className: "cs-help",
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Indent applies to the left or right edge according to alignment. Centered text ignores indent. Explicit row heights still clip content; use AutoFit when more space is needed.",
                              }),
                            }),
                          ],
                        }),
                      f === "Border" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("div", {
                              className: "cs-fields",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, { children: "Line" }),
                                    i.jsx("select", {
                                      "aria-label": text("Border line style"),
                                      value: line.style,
                                      onChange: (event) =>
                                        A({
                                          ...line,
                                          style: event.target.value,
                                        }),
                                      lang: language("Border line style"),
                                      "data-ui-label": "Border line style",
                                      children: borderLineStyles.map(
                                        (borderLineStyle) =>
                                          i.jsx(
                                            "option",
                                            {
                                              value: borderLineStyle,
                                              children:
                                                borderLineStyle
                                                  .charAt(0)
                                                  .toUpperCase() +
                                                borderLineStyle.slice(1),
                                            },
                                            borderLineStyle,
                                          ),
                                      ),
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, { children: "Color" }),
                                    i.jsx("input", {
                                      type: "color",
                                      "aria-label": text("Border color"),
                                      value: expandShortHexColor(
                                        line.color,
                                        "#000000",
                                      ),
                                      onChange: (event) =>
                                        A({
                                          ...line,
                                          color: event.target.value,
                                        }),
                                      lang: language("Border color"),
                                      "data-ui-label": "Border color",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              className: "cs-presets",
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Presets for selected cells",
                                  }),
                                }),
                                i.jsx("button", {
                                  type: "button",
                                  onClick: () => {
                                    (w(undefined),
                                      v((H) => {
                                        const Q = { ...H };
                                        return (delete Q.borders, Q);
                                      }),
                                      b(""));
                                  },
                                  children: i.jsx(LocalizedText, {
                                    children: "Keep existing borders",
                                  }),
                                }),
                                [
                                  ["clear", "Clear borders"],
                                  ["outline", "Outline borders"],
                                  ["inside", "Inside borders"],
                                  ["all", "All borders"],
                                ].map(([H, Q]) =>
                                  i.jsx(
                                    "button",
                                    {
                                      type: "button",
                                      "aria-pressed": preset === H,
                                      onClick: () => R(H),
                                      children: Q,
                                    },
                                    H,
                                  ),
                                ),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              className: "cs-presets",
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Individual edges",
                                  }),
                                }),
                                cellBorderEdges.map((cellBorderEdge) =>
                                  i.jsxs(
                                    "span",
                                    {
                                      children: [
                                        i.jsx("button", {
                                          type: "button",
                                          "aria-label": `Set ${cellBorderEdge} border`,
                                          onClick: () => N(cellBorderEdge),
                                          children:
                                            cellBorderEdge
                                              .charAt(0)
                                              .toUpperCase() +
                                            cellBorderEdge.slice(1),
                                        }),
                                        i.jsx("button", {
                                          type: "button",
                                          className: "cs-edge-clear",
                                          "aria-label": `Clear ${cellBorderEdge} border`,
                                          onClick: () =>
                                            N(cellBorderEdge, true),
                                          children: "×",
                                        }),
                                      ],
                                    },
                                    cellBorderEdge,
                                  ),
                                ),
                              ],
                            }),
                            i.jsx("p", {
                              className: "cs-help",
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Choose a line and color, then a preset or edge. Presets follow the selected cells, including gaps. Outline and Inside retain other existing edges; Clear removes every selected edge. Borders on unselected neighboring cells remain intact.",
                              }),
                            }),
                          ],
                        }),
                    ],
                  }),
                  i.jsxs("section", {
                    className: "cs-preview",
                    "aria-label": text("Cell style preview"),
                    lang: language("Cell style preview"),
                    "data-ui-label": "Cell style preview",
                    children: [
                      i.jsxs("h3", {
                        children: [
                          i.jsx(LocalizedText, { children: "Preview · " }),
                          sheet?.name,
                          "!",
                          l,
                        ],
                      }),
                      f === "Border"
                        ? i.jsx("table", {
                            className: "cs-border-preview",
                            "aria-label": text("Border preset preview"),
                            lang: language("Border preset preview"),
                            "data-ui-label": "Border preset preview",
                            children: i.jsx("tbody", {
                              children: [0, 1, 2].map((H) =>
                                i.jsx(
                                  "tr",
                                  {
                                    children: [0, 1, 2].map((Q) => {
                                      const X = `${columnIndexToLetters(Q)}${H + 1}`,
                                        J = L(X);
                                      return i.jsx(
                                        "td",
                                        {
                                          "data-preview-cell": X,
                                          style: cellStyleToCss(J, "", {
                                            defaultBorder: "1px dotted #d5d5d5",
                                          }),
                                          children:
                                            H === 1 && Q === 1 ? "Text" : "",
                                        },
                                        Q,
                                      );
                                    }),
                                  },
                                  H,
                                ),
                              ),
                            }),
                          })
                        : i.jsx("table", {
                            className: "cs-text-preview",
                            children: i.jsx("tbody", {
                              children: i.jsx("tr", {
                                children: i.jsx("td", {
                                  style: cellStyleToCss(L("B2"), V, {
                                    defaultBorder: "1px solid #d5d5d5",
                                  }),
                                  children: i.jsx("div", {
                                    style: {
                                      ...textWrapToCss(_),
                                      maxHeight: 94,
                                      overflow: "hidden",
                                      lineHeight: 1.35,
                                    },
                                    children: i.jsx(CellValuePreview, {
                                      dateSystem:
                                        e.workbook.datePolicy?.dateSystem,
                                      value: V === null ? "Audit evidence" : V,
                                      cell: {
                                        ...(sheet?.cells[l] ?? { raw: null }),
                                        style: _,
                                      },
                                    }),
                                  }),
                                }),
                              }),
                            }),
                          }),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Values, formulas, number formats and review evidence remain unchanged.",
                        }),
                      }),
                    ],
                  }),
                  (g || C) &&
                    i.jsx("p", {
                      className: "cs-error",
                      role: "alert",
                      children: g || C,
                    }),
                ],
              }),
              i.jsxs("footer", {
                className: "cs-actions",
                children: [
                  i.jsx("button", {
                    type: "button",
                    onClick: e.onClose,
                    children: i.jsx(LocalizedText, { children: "Cancel" }),
                  }),
                  i.jsx("button", {
                    className: "cs-primary",
                    disabled: B || !!C,
                    children: i.jsx(LocalizedText, {
                      children: "Apply styles",
                    }),
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    })
  );
}
const conditionalRuleTypeLabels = {
  greater: "Greater than",
  less: "Less than",
  equal: "Equal to",
  between: "Between",
  contains: "Text contains",
  duplicate: "Duplicate values",
  unique: "Unique values",
  blank: "Blank cells",
  notBlank: "Nonblank cells",
  error: "Errors",
  notError: "No errors",
  formula: "Formula is true",
  colorScale: "Two-color scale",
  dataBar: "Data bars",
  iconSet: "Three-level icons",
};
function createConditionalRuleDraft(range) {
  return {
    id: "",
    range: range,
    kind: "greater",
    value: "0",
    second: "100",
    fill: "#fff2cc",
    color: "#9c5700",
    stopIfTrue: false,
  };
}
function describeDataBarRule(e) {
  return [
    e.style.fill ? `Fill ${e.style.fill}` : "No fill",
    e.style.color ? `Text ${e.style.color}` : "",
    e.dataBar
      ? `Bar ${e.dataBar.width.toFixed(1)}% at ${e.dataBar.start.toFixed(1)}%${e.dataBar.negative ? " negative" : ""}`
      : "",
    e.icon ? `${e.icon.label} icon` : "",
    ...e.errors.map((error) => error.message),
  ]
    .filter(Boolean)
    .join("; ");
}
function ConditionalFormattingRulesManager(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef({
      book: e.workbook,
      sheetId: e.sheetId,
      range: e.range,
      guard: JSON.stringify(e.workbook),
    }),
    [rules, d] = ne.useState(() =>
      structuredClone(
        e.workbook.sheets.find((sheet) => sheet.id === e.sheetId)
          ?.conditionalFormats ?? [],
      ),
    ),
    [u, l] = ne.useState(() => createConditionalRuleDraft(e.range)),
    [c, h] = ne.useState(""),
    [f, p] = ne.useState(null),
    [m, v] = ne.useState(0),
    sheet = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    b =
      a.current.guard !== JSON.stringify(e.workbook) ||
      a.current.sheetId !== e.sheetId ||
      a.current.range !== e.range,
    disabled = e.readOnly || b || !sheet || !!sheet.protection?.enabled,
    w = sortConditionalRulesByPriority(rules),
    E = ["colorScale", "dataBar", "iconSet"].includes(u.kind);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      o.current?.querySelector("button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const A = (D) => {
      (d(D.map((B, _) => ({ ...B, priority: _ + 1 }))), p(null), h(""));
    },
    k = (D) => {
      (l({ ...u, ...D }), p(null), h(""));
    },
    F = () => {
      try {
        if (disabled) throw Error("Rules cannot be changed in this view.");
        const D = {
            ...u,
            id: u.id || crypto.randomUUID(),
            ...(E ? { stopIfTrue: false } : {}),
          },
          B = u.id ? w.map((C) => (C.id === u.id ? D : C)) : [D, ...w],
          conditionalRules = normalizeConditionalRules(
            B.map((C, T) => ({ ...C, priority: T + 1 })),
          );
        (A(conditionalRules), l(createConditionalRuleDraft(e.range)));
      } catch (error) {
        h(error instanceof Error ? error.message : "Rule could not be staged.");
      }
    },
    onClick = () => {
      try {
        if (disabled) throw Error("Rules cannot be changed in this view.");
        (p(
          updateConditionalFormatRules(e.workbook, e.sheetId, rules, {
            readOnly: e.readOnly,
          }),
        ),
          v(0),
          h(""));
      } catch (error) {
        h(error instanceof Error ? error.message : "Preview failed.");
      }
    };
  return i.jsx("div", {
    className: "cf-backdrop",
    children: i.jsxs("div", {
      ref: o,
      className: "cf-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), e.onClose()),
          event.key === "Tab")
        ) {
          const _ = [
              ...(o.current?.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
              ) ?? []),
            ],
            C = _[0],
            T = _.at(-1);
          event.shiftKey && document.activeElement === C
            ? (event.preventDefault(), T == null || T.focus())
            : !event.shiftKey &&
              document.activeElement === T &&
              (event.preventDefault(), C == null || C.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: `${n}-title`,
              children: i.jsx(LocalizedText, {
                children: "Conditional Formatting Rules Manager",
              }),
            }),
            i.jsx("button", {
              "aria-label": text("Close conditional formatting"),
              onClick: e.onClose,
              lang: language("Close conditional formatting"),
              "data-ui-label": "Close conditional formatting",
              children: "×",
            }),
          ],
        }),
        i.jsxs("p", {
          children: [
            sheet?.name ?? "Worksheet removed",
            " ",
            i.jsx(LocalizedText, {
              children:
                "· Highest priority first. Stage changes, inspect the preview, then Apply once.",
            }),
          ],
        }),
        b &&
          i.jsx("p", {
            role: "alert",
            children: i.jsx(LocalizedText, {
              children:
                "The workbook or selection changed. Close and reopen this manager to refresh.",
            }),
          }),
        e.readOnly &&
          i.jsx("p", {
            role: "alert",
            children: i.jsx(LocalizedText, {
              children: "This workbook is read-only.",
            }),
          }),
        sheet?.protection?.enabled &&
          i.jsx("p", {
            role: "alert",
            children: i.jsx(LocalizedText, {
              children: "Unprotect this worksheet before changing rules.",
            }),
          }),
        rules.some((rule) => rule.priority === undefined) &&
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "Legacy rules retain their last-wins behavior. Editing the list assigns explicit priorities that preserve the displayed order; the preview shows this migration.",
            }),
          }),
        i.jsxs("div", {
          className: "cf-workspace",
          children: [
            i.jsxs("section", {
              "aria-label": text("Conditional rules"),
              lang: language("Conditional rules"),
              "data-ui-label": "Conditional rules",
              children: [
                i.jsxs("h3", {
                  children: [
                    i.jsx(LocalizedText, { children: "Rules (" }),
                    rules.length,
                    "/100)",
                  ],
                }),
                w.length === 0
                  ? i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children: "No conditional rules.",
                      }),
                    })
                  : i.jsx("ol", {
                      className: "cf-rules",
                      children: w.map((D, B) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              i.jsx("strong", {
                                children: conditionalRuleTypeLabels[D.kind],
                              }),
                              " · ",
                              D.range,
                              i.jsxs("span", {
                                children: [
                                  D.formula ?? D.value,
                                  D.kind === "between" ? ` to ${D.second}` : "",
                                  D.stopIfTrue ? " · Stop if true" : "",
                                ],
                              }),
                              i.jsxs("div", {
                                children: [
                                  i.jsxs("button", {
                                    disabled: disabled,
                                    onClick: () => {
                                      (l(structuredClone(D)), h(""));
                                    },
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Edit rule ",
                                      }),
                                      B + 1,
                                    ],
                                  }),
                                  i.jsxs("button", {
                                    disabled: disabled,
                                    onClick: () => {
                                      (A(w.filter((_) => _.id !== D.id)),
                                        u.id === D.id &&
                                          l(
                                            createConditionalRuleDraft(e.range),
                                          ));
                                    },
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Delete rule ",
                                      }),
                                      B + 1,
                                    ],
                                  }),
                                  i.jsx("button", {
                                    "aria-label": `Move rule ${B + 1} up`,
                                    disabled: disabled || B === 0,
                                    onClick: () => {
                                      const _ = [...w];
                                      (([_[B - 1], _[B]] = [_[B], _[B - 1]]),
                                        A(_));
                                    },
                                    children: "↑",
                                  }),
                                  i.jsx("button", {
                                    "aria-label": `Move rule ${B + 1} down`,
                                    disabled: disabled || B === w.length - 1,
                                    onClick: () => {
                                      const _ = [...w];
                                      (([_[B + 1], _[B]] = [_[B], _[B + 1]]),
                                        A(_));
                                    },
                                    children: "↓",
                                  }),
                                ],
                              }),
                            ],
                          },
                          D.id,
                        ),
                      ),
                    }),
              ],
            }),
            i.jsxs("form", {
              className: "cf-editor",
              onSubmit: (event) => {
                (event.preventDefault(), F());
              },
              children: [
                i.jsx("h3", { children: u.id ? "Edit rule" : "New rule" }),
                i.jsxs("fieldset", {
                  disabled: disabled,
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Rule type" }),
                        i.jsx("select", {
                          value: u.kind,
                          onChange: (event) => {
                            const value = event.target.value;
                            k({
                              kind: value,
                              ...(value === "formula"
                                ? { formula: u.formula ?? "=A1>0" }
                                : {}),
                              ...(["colorScale", "dataBar", "iconSet"].includes(
                                value,
                              )
                                ? { stopIfTrue: false }
                                : {}),
                            });
                          },
                          children: conditionalRuleKinds.map(
                            (conditionalRuleKind) =>
                              i.jsx(
                                "option",
                                {
                                  value: conditionalRuleKind,
                                  children:
                                    conditionalRuleTypeLabels[
                                      conditionalRuleKind
                                    ],
                                },
                                conditionalRuleKind,
                              ),
                          ),
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Applies to" }),
                        i.jsx("input", {
                          value: u.range,
                          onChange: (event) => k({ range: event.target.value }),
                        }),
                      ],
                    }),
                    u.kind === "formula"
                      ? i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Condition formula",
                                }),
                                i.jsx("input", {
                                  value: u.formula ?? "",
                                  onChange: (event) =>
                                    k({ formula: event.target.value }),
                                }),
                              ],
                            }),
                            i.jsx("small", {
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Write relative references for the top-left cell of Applies to. Use $ to anchor rows or columns.",
                              }),
                            }),
                          ],
                        })
                      : [
                          "greater",
                          "less",
                          "equal",
                          "between",
                          "contains",
                        ].includes(u.kind) &&
                        i.jsxs("label", {
                          children: [
                            u.kind === "contains"
                              ? "Text to find"
                              : "Rule value",
                            i.jsx("input", {
                              value: u.value,
                              onChange: (event) =>
                                k({ value: event.target.value }),
                            }),
                          ],
                        }),
                    u.kind === "between" &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Second value" }),
                          i.jsx("input", {
                            value: u.second,
                            onChange: (event) =>
                              k({ second: event.target.value }),
                          }),
                        ],
                      }),
                    i.jsxs("div", {
                      className: "cf-colors",
                      children: [
                        i.jsxs("label", {
                          children: [
                            u.kind === "dataBar"
                              ? "Positive bar color"
                              : "Fill color",
                            i.jsx("input", {
                              type: "color",
                              value: u.fill,
                              onChange: (event) =>
                                k({ fill: event.target.value }),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Text color" }),
                            i.jsx("input", {
                              type: "color",
                              value: u.color,
                              onChange: (event) =>
                                k({ color: event.target.value }),
                            }),
                          ],
                        }),
                      ],
                    }),
                    u.kind === "colorScale" &&
                      i.jsxs("div", {
                        className: "cf-colors",
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Minimum color",
                              }),
                              i.jsx("input", {
                                type: "color",
                                value: u.minColor ?? "#f8e99b",
                                onChange: (event) =>
                                  k({ minColor: event.target.value }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Maximum color",
                              }),
                              i.jsx("input", {
                                type: "color",
                                value: u.maxColor ?? "#63be7b",
                                onChange: (event) =>
                                  k({ maxColor: event.target.value }),
                              }),
                            ],
                          }),
                        ],
                      }),
                    u.kind === "dataBar" &&
                      i.jsx("small", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Automatic numeric minimum and maximum, a zero axis, red negative bars. Values stay visible.",
                        }),
                      }),
                    u.kind === "iconSet" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Icon threshold units",
                              }),
                              i.jsxs("select", {
                                value: u.thresholdType ?? "percent",
                                onChange: (event) =>
                                  k({ thresholdType: event.target.value }),
                                children: [
                                  i.jsx("option", {
                                    value: "percent",
                                    children: "Percent of numeric span",
                                  }),
                                  i.jsx("option", {
                                    value: "number",
                                    children: "Number",
                                  }),
                                ],
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Middle icon starts at",
                              }),
                              i.jsx(LocalizedNumberField, {
                                "aria-label": text("Middle icon starts at"),
                                value: u.iconThresholds?.[0] ?? 33,
                                onValueChange: (D) =>
                                  k({
                                    iconThresholds: [
                                      Number(D),
                                      u.iconThresholds?.[1] ?? 67,
                                    ],
                                  }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "High icon starts at",
                              }),
                              i.jsx(LocalizedNumberField, {
                                "aria-label": text("High icon starts at"),
                                value: u.iconThresholds?.[1] ?? 67,
                                onValueChange: (D) =>
                                  k({
                                    iconThresholds: [
                                      u.iconThresholds?.[0] ?? 33,
                                      Number(D),
                                    ],
                                  }),
                              }),
                            ],
                          }),
                        ],
                      }),
                    i.jsxs("label", {
                      className: "cf-checkbox",
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: u.stopIfTrue ?? false,
                          disabled: E || disabled,
                          onChange: (event) =>
                            k({ stopIfTrue: event.target.checked }),
                        }),
                        i.jsx(LocalizedText, { children: "Stop if true" }),
                      ],
                    }),
                    i.jsx("button", {
                      type: "submit",
                      children: u.id ? "Stage rule edit" : "Stage new rule",
                    }),
                    i.jsx("button", {
                      type: "button",
                      onClick: () => l(createConditionalRuleDraft(e.range)),
                      children: i.jsx(LocalizedText, {
                        children: "Reset editor",
                      }),
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
        c && i.jsx("p", { role: "alert", children: c }),
        i.jsx("button", {
          disabled: disabled,
          onClick: onClick,
          children: i.jsx(LocalizedText, { children: "Preview rule changes" }),
        }),
        f &&
          i.jsxs("section", {
            className: "cf-preview",
            "aria-label": text("Conditional formatting preview"),
            lang: language("Conditional formatting preview"),
            "data-ui-label": "Conditional formatting preview",
            children: [
              i.jsxs("h3", {
                children: [
                  f.changes.length,
                  " ",
                  i.jsx(LocalizedText, { children: "rule changes" }),
                ],
              }),
              f.changes.length === 0 &&
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children: "No changes to apply.",
                  }),
                }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, {
                    children:
                      "Exact rule metadata is listed below. The visual sample shows up to",
                  }),
                  " ",
                  f.previewLimit,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "anchors and populated cells; it is not an exhaustive scan of the range.",
                  }),
                ],
              }),
              f.changes.map((change) =>
                i.jsxs(
                  "details",
                  {
                    children: [
                      i.jsxs("summary", {
                        children: [
                          change.before && change.after
                            ? "Edit"
                            : change.after
                              ? "Add"
                              : "Delete",
                          " ",
                          conditionalRuleTypeLabels[
                            (change.after ?? change.before).kind
                          ],
                          " ·",
                          " ",
                          (change.after ?? change.before).range,
                        ],
                      }),
                      i.jsxs("div", {
                        className: "cf-change",
                        children: [
                          i.jsxs("pre", {
                            children: [
                              i.jsx(LocalizedText, { children: "Before" }),
                              `
`,
                              change.before
                                ? JSON.stringify(change.before, null, 2)
                                : "None",
                            ],
                          }),
                          i.jsxs("pre", {
                            children: [
                              i.jsx(LocalizedText, { children: "After" }),
                              `
`,
                              change.after
                                ? JSON.stringify(change.after, null, 2)
                                : "None",
                            ],
                          }),
                        ],
                      }),
                    ],
                  },
                  change.id,
                ),
              ),
              f.errors.length > 0 &&
                i.jsxs("div", {
                  role: "alert",
                  children: [
                    i.jsx("strong", {
                      children: i.jsx(LocalizedText, {
                        children: "Rule errors",
                      }),
                    }),
                    i.jsx("ul", {
                      children: f.errors.map((error, B) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              error.address || "Range",
                              " · ",
                              error.ruleId,
                              ": ",
                              error.message,
                            ],
                          },
                          B,
                        ),
                      ),
                    }),
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Cells with a rule error receive no formatting from that rule. Apply keeps the rule available for later repair.",
                      }),
                    }),
                  ],
                }),
              i.jsx("div", {
                className: "cf-table",
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
                              children: "Value",
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
                      children: f.samples.slice(m * 10, m * 10 + 10).map((D) =>
                        i.jsxs(
                          "tr",
                          {
                            children: [
                              i.jsx("th", { children: D.address }),
                              i.jsx("td", {
                                children: String(D.value ?? ""),
                              }),
                              i.jsx("td", {
                                children: describeDataBarRule(D.before),
                              }),
                              i.jsx("td", {
                                children: describeDataBarRule(D.after),
                              }),
                            ],
                          },
                          D.address,
                        ),
                      ),
                    }),
                  ],
                }),
              }),
              f.samples.length > 10 &&
                i.jsxs("nav", {
                  "aria-label": text("Conditional preview pages"),
                  lang: language("Conditional preview pages"),
                  "data-ui-label": "Conditional preview pages",
                  children: [
                    i.jsx("button", {
                      disabled: m === 0,
                      onClick: () => v(m - 1),
                      children: i.jsx(LocalizedText, {
                        children: "Previous samples",
                      }),
                    }),
                    i.jsxs("span", {
                      children: [
                        i.jsx(LocalizedText, { children: "Page " }),
                        m + 1,
                        " ",
                        i.jsx(LocalizedText, { children: "of " }),
                        Math.ceil(f.samples.length / 10),
                      ],
                    }),
                    i.jsx("button", {
                      disabled: (m + 1) * 10 >= f.samples.length,
                      onClick: () => v(m + 1),
                      children: i.jsx(LocalizedText, {
                        children: "Next samples",
                      }),
                    }),
                  ],
                }),
            ],
          }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              disabled: disabled || !f || f.changes.length === 0,
              onClick: () => {
                try {
                  if (!f || disabled) return;
                  const D = assertConditionalFormatPreviewCurrent(
                    e.workbook,
                    f,
                    { readOnly: e.readOnly },
                  );
                  (D !== e.workbook &&
                    e.onWorkbookChange(
                      D,
                      "Update conditional formatting rules",
                    ),
                    e.onClose());
                } catch (error) {
                  h(
                    error instanceof Error
                      ? error.message
                      : "Rules could not be applied.",
                  );
                }
              },
              children: i.jsx(LocalizedText, { children: "Apply rules" }),
            }),
          ],
        }),
      ],
    }),
  });
}
function ConditionalFormatIndicator({ decoration, children }) {
  const { text, language } = useWorkbookMessages();
  if (!decoration?.dataBar && !decoration?.icon && !decoration?.errors.length)
    return i.jsx(i.Fragment, { children });
  const dataBar = decoration.dataBar;
  return i.jsxs("span", {
    style: { display: "block", position: "relative", minWidth: 0 },
    children: [
      dataBar &&
        i.jsxs(i.Fragment, {
          children: [
            i.jsx("span", {
              "data-testid": "conditional-data-bar",
              "aria-hidden": "true",
              style: {
                position: "absolute",
                pointerEvents: "none",
                left: `${dataBar.start}%`,
                width: `${dataBar.width}%`,
                top: "12%",
                bottom: "12%",
                background: dataBar.color,
                opacity: 0.28,
              },
            }),
            i.jsx("span", {
              "aria-hidden": "true",
              style: {
                position: "absolute",
                left: `${dataBar.axis}%`,
                top: 0,
                bottom: 0,
                borderLeft: "1px solid #777",
                pointerEvents: "none",
              },
            }),
          ],
        }),
      i.jsxs("span", {
        style: { position: "relative", zIndex: 1 },
        children: [
          decoration.icon &&
            i.jsx("span", {
              role: "img",
              "aria-label": `${decoration.icon.label} conditional indicator`,
              title: `${decoration.icon.label} conditional indicator`,
              style: { color: decoration.icon.color, marginRight: 4 },
              children: decoration.icon.symbol,
            }),
          children,
          decoration.errors.length > 0 &&
            i.jsx("span", {
              role: "img",
              "aria-label": text("Conditional formatting error"),
              title: decoration.errors.map(
                (error) => `${error.ruleId}: ${error.message}`,
              ).join(`
`),
              style: { color: "#a4262c", marginLeft: 3 },
              lang: language("Conditional formatting error"),
              "data-ui-label": "Conditional formatting error",
              children: "⚠",
            }),
        ],
      }),
    ],
  });
}
const rangeToA1Address = (e) => {
  const rangeAddress = parseRangeAddress(e);
  return `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
};
function previewMergeCells(book, sheetId, range, unmerge = false, o = {}) {
  if (o.readOnly) throw Error("This workbook is read-only.");
  const sheetOrThrow = findSheetOrThrow(book, sheetId);
  assertSheetUnprotected(sheetOrThrow);
  const s = rangeToA1Address(range);
  if (countRangeCells(s) > 500)
    throw Error("Preview a merge or unmerge of at most 500 cells.");
  if (!unmerge && countRangeCells(s) < 2)
    throw Error("Select at least two cells to merge.");
  const d = (sheetOrThrow.merges ?? []).filter((p) => rangesOverlap(p, s));
  if (!unmerge && d.length) throw Error("Unmerge overlapping areas first.");
  if (unmerge && !d.length)
    throw Error("The selected range does not intersect a merged area.");
  const u = unmerge ? d : [s];
  if (u.reduce((p, m) => p + countRangeCells(m), 0) > 500)
    throw Error(
      "Unmerge fewer areas; this preview supports at most 500 affected cells.",
    );
  if (!unmerge) {
    if ((sheetOrThrow.tables ?? []).some((p) => rangesOverlap(p.range, s)))
      throw Error(
        "Merging cells inside a table is unsupported. Convert the table first.",
      );
    if (
      [...(sheetOrThrow.pivots ?? []), ...(sheetOrThrow.queries ?? [])].some(
        (p) => rangesOverlap(p.outputRange, s),
      )
    )
      throw Error("Merging generated analysis output is unsupported.");
  }
  normalizeWorkbookInput(mergeWorksheetCells(book, sheetId, s, unmerge));
  const workbookWithSettings = calculateWorkbookWithSettings(book);
  if (!unmerge) {
    let p = 0;
    for (const [m, cell] of Object.entries(sheetOrThrow.cells))
      if (
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal
      ) {
        if (++p > 1e4)
          throw Error(
            "Merge checking supports at most 10,000 active formulas per worksheet.",
          );
        const g = workbookWithSettings.getSpillRange(sheetId, m);
        if (g && rangesOverlap(g, s))
          throw Error(
            `Unmergeable array output at ${g}. Move or remove its formula first.`,
          );
      }
  }
  const cells = [];
  for (const p of u) {
    const rangeAddress = parseRangeAddress(p),
      v = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`;
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
        const address = `${columnIndexToLetters(firstCol)}${firstRow + 1}`,
          cell = sheetOrThrow.cells[address],
          discussions = (book.reviewThreads ?? []).filter(
            (A) =>
              !A.orphaned &&
              A.anchor.sheetId === sheetId &&
              A.anchor.address === address,
          ).length;
        (cell || discussions || address === v) &&
          cells.push({
            address: address,
            cell: structuredClone(cell ?? { raw: null }),
            display: formatNumberWithPattern(
              workbookWithSettings(sheetId, address),
              cell,
              book.datePolicy?.dateSystem,
            ),
            discussions: discussions,
            effect:
              address === v
                ? "Anchor remains visible"
                : unmerge
                  ? "Revealed by unmerge"
                  : "Hidden until unmerge",
          });
      }
  }
  const hiddenValues = cells.filter(
    (cell) =>
      cell.effect === "Hidden until unmerge" &&
      cell.cell.raw !== null &&
      cell.cell.raw !== "",
  ).length;
  return {
    version: 1,
    snapshot: JSON.stringify(book),
    sheetId: sheetId,
    sheetName: sheetOrThrow.name,
    range: s,
    unmerge: unmerge,
    affectedRanges: [...u],
    cells: cells,
    hiddenValues: hiddenValues,
    xlsxBlocked: cells.some(
      (cell) =>
        cell.effect === "Hidden until unmerge" &&
        ((cell.cell.raw !== null && cell.cell.raw !== "") ||
          cell.cell.note ||
          cell.cell.tickmark ||
          cell.cell.hyperlink),
    ),
  };
}
function assertMergePreviewCurrent(workbook, plan, t = {}) {
  if (plan.snapshot !== JSON.stringify(workbook))
    throw Error(
      "The workbook changed. Close this preview and review the current cells again.",
    );
  const n = previewMergeCells(
      workbook,
      plan.sheetId,
      plan.range,
      plan.unmerge,
      t,
    ),
    o = (a) => ({ ...a, cells: a.cells.map(({ display, ...d }) => d) });
  if (JSON.stringify(o(n)) !== JSON.stringify(o(plan)))
    throw Error("The merge preview changed. Create a fresh preview.");
  return mergeWorksheetCells(workbook, plan.sheetId, plan.range, plan.unmerge);
}

export {
  previewMergeCells,
  ConditionalFormatIndicator,
  FormatCellsDialog,
  ConditionalFormattingRulesManager,
  assertMergePreviewCurrent,
};
