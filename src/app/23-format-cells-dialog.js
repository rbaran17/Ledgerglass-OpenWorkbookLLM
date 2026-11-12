import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { supportedFontFamilies } from "./02-financial-functions.js";
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
  var $;
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const s = e.workbook.sheets.find((H) => H.id === e.sheetId),
    d = parseSelectionAreas(e.ranges ?? [e.range]),
    u = parseRangeAddress(e.range),
    l =
      e.active && isCellAddressInRanges(d, e.active)
        ? e.active
        : `${columnIndexToLetters(u.firstCol)}${u.firstRow + 1}`,
    c = ne.useRef({
      book: e.workbook.id,
      revision: e.workbook.revision,
      content: JSON.stringify(e.workbook),
      sheet: e.sheetId,
      selection: JSON.stringify(e.ranges ?? [e.range]),
      active: e.active,
    }),
    [h] = ne.useState(() => {
      return validateCellStyle(resolveEffectiveCellStyle(s?.cells[l]?.style));
    }),
    [f, p] = ne.useState(
      e.mode === "alignment"
        ? "Alignment"
        : e.mode === "borders"
          ? "Border"
          : "Font",
    ),
    [m, v] = ne.useState({}),
    [g, b] = ne.useState(""),
    [y, w] = ne.useState(),
    [E, A] = ne.useState({ style: "thin", color: "#000000" }),
    [k, F] = ne.useState(h.fontSize === undefined ? "" : String(h.fontSize)),
    [I, O] = ne.useState(String(h.indent ?? 0)),
    j = countSelectionCells(d),
    z = ne.useMemo(
      () => calculateWorkbookWithSettings(e.workbook),
      [e.workbook],
    ),
    V = s ? z(s.id, l) : null,
    D =
      c.current.book !== e.workbook.id ||
      c.current.revision !== e.workbook.revision ||
      c.current.content !== JSON.stringify(e.workbook) ||
      c.current.sheet !== e.sheetId ||
      c.current.selection !== JSON.stringify(e.ranges ?? [e.range]) ||
      c.current.active !== e.active,
    B =
      e.readOnly ||
      D ||
      !!(($ = s?.protection) != null && $.enabled) ||
      !s ||
      j > 500;
  let _ = h,
    C = "";
  try {
    _ = validateCellStyle({
      ...h,
      ...m,
      ...(m.borders ? { borders: { ...h.borders, ...m.borders } } : {}),
    });
  } catch (H) {
    C = H instanceof Error ? H.message : "Choose supported cell styles.";
  }
  function T(H) {
    (v((Q) => ({ ...Q, ...H })), b(""));
  }
  function N(H, Q = false) {
    (v((X) => ({
      ...X,
      borders: { ...X.borders, [H]: Q ? { style: "none" } : { ...E } },
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
  const L = (H) => ({
    ..._,
    borders: {
      ...h.borders,
      ...(y ? buildBorderPresetStyles(H, sampleGridAddresses, y, E) : {}),
      ...m.borders,
    },
  });
  return (
    ne.useEffect(() => {
      var X;
      const H = document.activeElement;
      return (
        (X = o.current?.querySelector('[role="tab"][aria-selected="true"]')) ==
          null || X.focus(),
        () => {
          H != null && H.isConnected && H.focus();
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
        onKeyDown: (H) => {
          if (
            (H.stopPropagation(),
            H.key === "Escape" && (H.preventDefault(), a.current()),
            H.key === "Tab")
          ) {
            const X = [
                ...(o.current?.querySelectorAll(
                  'button:not(:disabled):not([tabindex="-1"]),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
                ) ?? []),
              ],
              J = X[0],
              q = X.at(-1);
            H.shiftKey && document.activeElement === J
              ? (H.preventDefault(), q == null || q.focus())
              : !H.shiftKey &&
                document.activeElement === q &&
                (H.preventDefault(), J == null || J.focus());
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
            children: cellFormatTabs.map((H, Q) =>
              i.jsx(
                "button",
                {
                  role: "tab",
                  id: `${n}-${H}`,
                  "aria-controls": `${n}-panel`,
                  "aria-selected": f === H,
                  tabIndex: f === H ? 0 : -1,
                  onClick: () => p(H),
                  onKeyDown: (X) => {
                    var q;
                    if (
                      ["ArrowLeft", "ArrowRight", "Home", "End"].includes(X.key)
                    ) {
                      X.preventDefault();
                      const W =
                        X.key === "Home"
                          ? 0
                          : X.key === "End"
                            ? cellFormatTabs.length - 1
                            : (Q +
                                (X.key === "ArrowRight"
                                  ? 1
                                  : cellFormatTabs.length - 1)) %
                              cellFormatTabs.length;
                      (p(cellFormatTabs[W]),
                        (q = X.currentTarget.parentElement?.children[W]) ==
                          null || q.focus());
                    }
                  },
                  children: H,
                },
                H,
              ),
            ),
          }),
          i.jsxs("form", {
            onSubmit: (H) => {
              (H.preventDefault(), b(""));
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
                  d,
                  m,
                  {
                    readOnly: e.readOnly,
                    ...(y ? { borders: { preset: y, line: E } } : {}),
                  },
                );
                (Q !== e.workbook &&
                  e.onWorkbookChange(Q, "Format cell styles"),
                  e.onClose());
              } catch (Q) {
                b(
                  Q instanceof Error
                    ? Q.message
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
                      d.length,
                      " ",
                      d.length === 1 ? "area" : "areas",
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
                  s?.protection?.enabled &&
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
                                      onChange: (H) =>
                                        T({
                                          fontFamily:
                                            H.target.value || undefined,
                                        }),
                                      lang: language("Font family"),
                                      "data-ui-label": "Font family",
                                      children: [
                                        i.jsx("option", {
                                          value: "",
                                          children: "Default (Calibri)",
                                        }),
                                        supportedFontFamilies.map((H) =>
                                          i.jsx("option", { children: H }, H),
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
                                          onChange: (X) =>
                                            T({ [H]: X.target.checked }),
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
                                          onChange: (H) =>
                                            T({ color: H.target.value }),
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
                                          onChange: (H) =>
                                            T({ fill: H.target.value }),
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
                                      onChange: (H) =>
                                        T({
                                          align: H.target.value || undefined,
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
                                      onChange: (H) =>
                                        T({
                                          verticalAlign:
                                            H.target.value || undefined,
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
                                  onChange: (H) =>
                                    T({ wrap: H.target.checked }),
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
                                      value: E.style,
                                      onChange: (H) =>
                                        A({ ...E, style: H.target.value }),
                                      lang: language("Border line style"),
                                      "data-ui-label": "Border line style",
                                      children: borderLineStyles.map((H) =>
                                        i.jsx(
                                          "option",
                                          {
                                            value: H,
                                            children:
                                              H.charAt(0).toUpperCase() +
                                              H.slice(1),
                                          },
                                          H,
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
                                        E.color,
                                        "#000000",
                                      ),
                                      onChange: (H) =>
                                        A({ ...E, color: H.target.value }),
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
                                      "aria-pressed": y === H,
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
                                cellBorderEdges.map((H) =>
                                  i.jsxs(
                                    "span",
                                    {
                                      children: [
                                        i.jsx("button", {
                                          type: "button",
                                          "aria-label": `Set ${H} border`,
                                          onClick: () => N(H),
                                          children:
                                            H.charAt(0).toUpperCase() +
                                            H.slice(1),
                                        }),
                                        i.jsx("button", {
                                          type: "button",
                                          className: "cs-edge-clear",
                                          "aria-label": `Clear ${H} border`,
                                          onClick: () => N(H, true),
                                          children: "×",
                                        }),
                                      ],
                                    },
                                    H,
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
                          s?.name,
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
                                        ...(s?.cells[l] ?? { raw: null }),
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
function createConditionalRuleDraft(e) {
  return {
    id: "",
    range: e,
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
    ...e.errors.map((r) => r.message),
  ]
    .filter(Boolean)
    .join("; ");
}
function ConditionalFormattingRulesManager(e) {
  var O;
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef({
      book: e.workbook,
      sheetId: e.sheetId,
      range: e.range,
      guard: JSON.stringify(e.workbook),
    }),
    [s, d] = ne.useState(() => {
      return structuredClone(
        e.workbook.sheets.find((B) => B.id === e.sheetId)?.conditionalFormats ??
          [],
      );
    }),
    [u, l] = ne.useState(() => createConditionalRuleDraft(e.range)),
    [c, h] = ne.useState(""),
    [f, p] = ne.useState(null),
    [m, v] = ne.useState(0),
    g = e.workbook.sheets.find((D) => D.id === e.sheetId),
    b =
      a.current.guard !== JSON.stringify(e.workbook) ||
      a.current.sheetId !== e.sheetId ||
      a.current.range !== e.range,
    y = e.readOnly || b || !g || !!((O = g.protection) != null && O.enabled),
    w = sortConditionalRulesByPriority(s),
    E = ["colorScale", "dataBar", "iconSet"].includes(u.kind);
  ne.useEffect(() => {
    var _;
    const D = document.activeElement;
    return (
      (_ = o.current?.querySelector("button")) == null || _.focus(),
      () => {
        D != null && D.isConnected && D.focus();
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
        if (y) throw Error("Rules cannot be changed in this view.");
        const D = {
            ...u,
            id: u.id || crypto.randomUUID(),
            ...(E ? { stopIfTrue: false } : {}),
          },
          B = u.id ? w.map((C) => (C.id === u.id ? D : C)) : [D, ...w],
          _ = normalizeConditionalRules(
            B.map((C, T) => ({ ...C, priority: T + 1 })),
          );
        (A(_), l(createConditionalRuleDraft(e.range)));
      } catch (D) {
        h(D instanceof Error ? D.message : "Rule could not be staged.");
      }
    },
    I = () => {
      try {
        if (y) throw Error("Rules cannot be changed in this view.");
        (p(
          updateConditionalFormatRules(e.workbook, e.sheetId, s, {
            readOnly: e.readOnly,
          }),
        ),
          v(0),
          h(""));
      } catch (D) {
        h(D instanceof Error ? D.message : "Preview failed.");
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
      onKeyDown: (D) => {
        if (
          (D.stopPropagation(),
          D.key === "Escape" && (D.preventDefault(), e.onClose()),
          D.key === "Tab")
        ) {
          const _ = [
              ...(o.current?.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
              ) ?? []),
            ],
            C = _[0],
            T = _.at(-1);
          D.shiftKey && document.activeElement === C
            ? (D.preventDefault(), T == null || T.focus())
            : !D.shiftKey &&
              document.activeElement === T &&
              (D.preventDefault(), C == null || C.focus());
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
            g?.name ?? "Worksheet removed",
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
        g?.protection?.enabled &&
          i.jsx("p", {
            role: "alert",
            children: i.jsx(LocalizedText, {
              children: "Unprotect this worksheet before changing rules.",
            }),
          }),
        s.some((D) => D.priority === undefined) &&
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
                    s.length,
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
                                    disabled: y,
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
                                    disabled: y,
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
                                    disabled: y || B === 0,
                                    onClick: () => {
                                      const _ = [...w];
                                      (([_[B - 1], _[B]] = [_[B], _[B - 1]]),
                                        A(_));
                                    },
                                    children: "↑",
                                  }),
                                  i.jsx("button", {
                                    "aria-label": `Move rule ${B + 1} down`,
                                    disabled: y || B === w.length - 1,
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
              onSubmit: (D) => {
                (D.preventDefault(), F());
              },
              children: [
                i.jsx("h3", { children: u.id ? "Edit rule" : "New rule" }),
                i.jsxs("fieldset", {
                  disabled: y,
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Rule type" }),
                        i.jsx("select", {
                          value: u.kind,
                          onChange: (D) => {
                            const B = D.target.value;
                            k({
                              kind: B,
                              ...(B === "formula"
                                ? { formula: u.formula ?? "=A1>0" }
                                : {}),
                              ...(["colorScale", "dataBar", "iconSet"].includes(
                                B,
                              )
                                ? { stopIfTrue: false }
                                : {}),
                            });
                          },
                          children: conditionalRuleKinds.map((D) =>
                            i.jsx(
                              "option",
                              {
                                value: D,
                                children: conditionalRuleTypeLabels[D],
                              },
                              D,
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
                          onChange: (D) => k({ range: D.target.value }),
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
                                  onChange: (D) =>
                                    k({ formula: D.target.value }),
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
                              onChange: (D) => k({ value: D.target.value }),
                            }),
                          ],
                        }),
                    u.kind === "between" &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Second value" }),
                          i.jsx("input", {
                            value: u.second,
                            onChange: (D) => k({ second: D.target.value }),
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
                              onChange: (D) => k({ fill: D.target.value }),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Text color" }),
                            i.jsx("input", {
                              type: "color",
                              value: u.color,
                              onChange: (D) => k({ color: D.target.value }),
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
                                onChange: (D) =>
                                  k({ minColor: D.target.value }),
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
                                onChange: (D) =>
                                  k({ maxColor: D.target.value }),
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
                                onChange: (D) =>
                                  k({ thresholdType: D.target.value }),
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
                                onValueChange: (D) => {
                                  return k({
                                    iconThresholds: [
                                      Number(D),
                                      u.iconThresholds?.[1] ?? 67,
                                    ],
                                  });
                                },
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
                                onValueChange: (D) => {
                                  return k({
                                    iconThresholds: [
                                      u.iconThresholds?.[0] ?? 33,
                                      Number(D),
                                    ],
                                  });
                                },
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
                          disabled: E || y,
                          onChange: (D) => k({ stopIfTrue: D.target.checked }),
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
          disabled: y,
          onClick: I,
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
              f.changes.map((D) =>
                i.jsxs(
                  "details",
                  {
                    children: [
                      i.jsxs("summary", {
                        children: [
                          D.before && D.after
                            ? "Edit"
                            : D.after
                              ? "Add"
                              : "Delete",
                          " ",
                          conditionalRuleTypeLabels[(D.after ?? D.before).kind],
                          " ·",
                          " ",
                          (D.after ?? D.before).range,
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
                              D.before
                                ? JSON.stringify(D.before, null, 2)
                                : "None",
                            ],
                          }),
                          i.jsxs("pre", {
                            children: [
                              i.jsx(LocalizedText, { children: "After" }),
                              `
`,
                              D.after
                                ? JSON.stringify(D.after, null, 2)
                                : "None",
                            ],
                          }),
                        ],
                      }),
                    ],
                  },
                  D.id,
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
                      children: f.errors.map((D, B) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              D.address || "Range",
                              " · ",
                              D.ruleId,
                              ": ",
                              D.message,
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
              disabled: y || !f || f.changes.length === 0,
              onClick: () => {
                try {
                  if (!f || y) return;
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
                } catch (D) {
                  h(
                    D instanceof Error
                      ? D.message
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
  if (
    !(decoration != null && decoration.dataBar) &&
    !(decoration != null && decoration.icon) &&
    !(decoration != null && decoration.errors.length)
  )
    return i.jsx(i.Fragment, { children });
  const o = decoration.dataBar;
  return i.jsxs("span", {
    style: { display: "block", position: "relative", minWidth: 0 },
    children: [
      o &&
        i.jsxs(i.Fragment, {
          children: [
            i.jsx("span", {
              "data-testid": "conditional-data-bar",
              "aria-hidden": "true",
              style: {
                position: "absolute",
                pointerEvents: "none",
                left: `${o.start}%`,
                width: `${o.width}%`,
                top: "12%",
                bottom: "12%",
                background: o.color,
                opacity: 0.28,
              },
            }),
            i.jsx("span", {
              "aria-hidden": "true",
              style: {
                position: "absolute",
                left: `${o.axis}%`,
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
              title: decoration.errors.map((a) => `${a.ruleId}: ${a.message}`)
                .join(`
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
const normalizeRangeBounds = (e, r) => {
    const t = parseRangeAddress(e),
      n = parseRangeAddress(r);
    return (
      t.firstRow <= n.lastRow &&
      t.lastRow >= n.firstRow &&
      t.firstCol <= n.lastCol &&
      t.lastCol >= n.firstCol
    );
  },
  rangeToA1Address = (e) => {
    const r = parseRangeAddress(e);
    return `${columnIndexToLetters(r.firstCol)}${r.firstRow + 1}:${columnIndexToLetters(r.lastCol)}${r.lastRow + 1}`;
  },
  rangeCellCount = (e) => {
    const r = parseRangeAddress(e);
    return (r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1);
  };
function previewMergeCells(e, r, t, n = false, o = {}) {
  if (o.readOnly) throw Error("This workbook is read-only.");
  const a = findSheetOrThrow(e, r);
  assertSheetUnprotected(a);
  const s = rangeToA1Address(t);
  if (rangeCellCount(s) > 500)
    throw Error("Preview a merge or unmerge of at most 500 cells.");
  if (!n && rangeCellCount(s) < 2)
    throw Error("Select at least two cells to merge.");
  const d = (a.merges ?? []).filter((p) => normalizeRangeBounds(p, s));
  if (!n && d.length) throw Error("Unmerge overlapping areas first.");
  if (n && !d.length)
    throw Error("The selected range does not intersect a merged area.");
  const u = n ? d : [s];
  if (u.reduce((p, m) => p + rangeCellCount(m), 0) > 500)
    throw Error(
      "Unmerge fewer areas; this preview supports at most 500 affected cells.",
    );
  if (!n) {
    if ((a.tables ?? []).some((p) => normalizeRangeBounds(p.range, s)))
      throw Error(
        "Merging cells inside a table is unsupported. Convert the table first.",
      );
    if (
      [...(a.pivots ?? []), ...(a.queries ?? [])].some((p) =>
        normalizeRangeBounds(p.outputRange, s),
      )
    )
      throw Error("Merging generated analysis output is unsupported.");
  }
  normalizeWorkbookInput(mergeWorksheetCells(e, r, s, n));
  const l = calculateWorkbookWithSettings(e);
  if (!n) {
    let p = 0;
    for (const [m, v] of Object.entries(a.cells))
      if (typeof v.raw == "string" && v.raw.startsWith("=") && !v.literal) {
        if (++p > 1e4)
          throw Error(
            "Merge checking supports at most 10,000 active formulas per worksheet.",
          );
        const g = l.getSpillRange(r, m);
        if (g && normalizeRangeBounds(g, s))
          throw Error(
            `Unmergeable array output at ${g}. Move or remove its formula first.`,
          );
      }
  }
  const c = [];
  for (const p of u) {
    const m = parseRangeAddress(p),
      v = `${columnIndexToLetters(m.firstCol)}${m.firstRow + 1}`;
    for (let g = m.firstRow; g <= m.lastRow; g++)
      for (let b = m.firstCol; b <= m.lastCol; b++) {
        const y = `${columnIndexToLetters(b)}${g + 1}`,
          w = a.cells[y],
          E = (e.reviewThreads ?? []).filter(
            (A) =>
              !A.orphaned && A.anchor.sheetId === r && A.anchor.address === y,
          ).length;
        (w || E || y === v) &&
          c.push({
            address: y,
            cell: structuredClone(w ?? { raw: null }),
            display: formatNumberWithPattern(
              l(r, y),
              w,
              e.datePolicy?.dateSystem,
            ),
            discussions: E,
            effect:
              y === v
                ? "Anchor remains visible"
                : n
                  ? "Revealed by unmerge"
                  : "Hidden until unmerge",
          });
      }
  }
  const h = c.filter(
    (p) =>
      p.effect === "Hidden until unmerge" &&
      p.cell.raw !== null &&
      p.cell.raw !== "",
  ).length;
  return {
    version: 1,
    snapshot: JSON.stringify(e),
    sheetId: r,
    sheetName: a.name,
    range: s,
    unmerge: n,
    affectedRanges: [...u],
    cells: c,
    hiddenValues: h,
    xlsxBlocked: c.some(
      (p) =>
        p.effect === "Hidden until unmerge" &&
        ((p.cell.raw !== null && p.cell.raw !== "") ||
          p.cell.note ||
          p.cell.tickmark ||
          p.cell.hyperlink),
    ),
  };
}
function assertMergePreviewCurrent(e, r, t = {}) {
  if (r.snapshot !== JSON.stringify(e))
    throw Error(
      "The workbook changed. Close this preview and review the current cells again.",
    );
  const n = previewMergeCells(e, r.sheetId, r.range, r.unmerge, t),
    o = (a) => ({ ...a, cells: a.cells.map(({ display, ...d }) => d) });
  if (JSON.stringify(o(n)) !== JSON.stringify(o(r)))
    throw Error("The merge preview changed. Create a fresh preview.");
  return mergeWorksheetCells(e, r.sheetId, r.range, r.unmerge);
}

export {
  previewMergeCells,
  ConditionalFormatIndicator,
  FormatCellsDialog,
  ConditionalFormattingRulesManager,
  assertMergePreviewCurrent,
};
