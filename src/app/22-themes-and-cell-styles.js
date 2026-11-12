import { cloneStructuredValue } from "./11-local-workbook-store-and-versions.js";
import i from "react/jsx-runtime";
import ne from "react";
import { X } from "lucide-react";
import {
  maxSelectionAreas,
  supportedFontFamilies,
} from "./02-text-and-financial-functions.js";
import {
  cellStyleToCss,
  getImportedNumberFormatWarning,
  isCellAddressInRanges,
  themePaletteRoles,
  collectStyleFallbackWarnings,
  builtInWorkbookThemes,
  parseRangeAddress,
  assertThemePreviewCurrent,
  parseSelectionAreas,
  createDefaultStyleCatalog,
  resolveBorderEdge,
  columnIndexToLetters,
  resolveNamedStyleAppearance,
  formatCellValueWithNumberFormat,
  countSelectionCells,
  applyNamedStyleChange,
  resolveEffectiveCellStyle,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  formatLocalizedNumberDisplay,
} from "./04-cell-input-paste-and-fill.js";

function ConnectionStatusNotice() {
  const { text, language } = useWorkbookMessages(),
    [t, n] = ne.useState(
      () => typeof navigator > "u" || navigator.onLine !== false,
    );
  return (
    ne.useEffect(() => {
      const o = () => n(true),
        a = () => n(false);
      return (
        window.addEventListener("online", o),
        window.addEventListener("offline", a),
        () => {
          (window.removeEventListener("online", o),
            window.removeEventListener("offline", a));
        }
      );
    }, []),
    t
      ? null
      : i.jsx("div", {
          className: "lg-notice",
          role: "status",
          "aria-label": text("Connection status"),
          lang: language("Connection status"),
          "data-ui-label": "Connection status",
          children: i.jsx(LocalizedText, {
            children:
              "You’re offline. Local editing, calculations, audit checks and demo assistance remain available. Connected host services may be unavailable.",
          }),
        })
  );
}
function parseGoToReference(activeWorkbook, r, goToReferenceText, n = {}) {
  if (
    typeof goToReferenceText != "string" ||
    !goToReferenceText.trim() ||
    goToReferenceText.length > 8192
  )
    throw Error(
      "Enter an address, range, or defined range name (up to 8,192 characters).",
    );
  const allowedSheetIds = n.allowedSheetIds,
    a = (c) => !allowedSheetIds || allowedSheetIds.includes(c),
    s = (c) => {
      const h = activeWorkbook.sheets.find(
        (sheet) => sheet.id === c && a(sheet.id),
      );
      if (!h)
        throw Error(
          "The destination is outside the available worksheet scope.",
        );
      if (h.hidden || h.veryHidden)
        throw Error(
          "The destination worksheet is hidden. Make it visible before navigating.",
        );
      return h;
    };
  s(r);
  const d = [],
    u = (c) => {
      const h = [];
      let f = false,
        p = 0;
      for (let m = 0; m < c.length; m++)
        c[m] === "'"
          ? f && c[m + 1] === "'"
            ? m++
            : (f = !f)
          : c[m] === "," && !f && (h.push(c.slice(p, m).trim()), (p = m + 1));
      if (
        (h.push(c.slice(p).trim()),
        f || h.some((m) => !m) || h.length > maxSelectionAreas)
      )
        throw Error(
          "Enter valid comma-separated addresses or defined range names.",
        );
      return h;
    },
    l = (c, h, f) => {
      if (f.length > 32)
        throw Error(
          "The defined-name chain exceeds the 32-name navigation limit.",
        );
      for (const g of u(c)) {
        const b = /^(?:'((?:[^']|'')+)'|([^'!,:\[\]()]+))!(.+)$/.exec(g);
        let sheetId = h,
          w = g;
        if (b) {
          const k = b[1]?.replace(/''/g, "'") ?? b[2].trim(),
            F = activeWorkbook.sheets.find(
              (sheet) =>
                a(sheet.id) && sheet.name.toLowerCase() === k.toLowerCase(),
            );
          if (!F)
            throw Error(
              "The destination is outside the available worksheet scope.",
            );
          ((sheetId = F.id), (w = b[3].trim()));
        }
        if (
          (s(sheetId),
          /^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/.test(
            w,
          ))
        ) {
          if (
            (d.push({ sheetId: sheetId, range: parseSelectionAreas(w)[0] }),
            d.length > maxSelectionAreas)
          )
            throw Error(
              `Select at most ${maxSelectionAreas} rectangular areas.`,
            );
          continue;
        }
        if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/.test(w))
          throw Error(
            "The Name Box accepts cell addresses and static range names. Computed, external, or three-dimensional destinations cannot be selected.",
          );
        const E =
          activeWorkbook.definedNames?.find(
            (definedName) =>
              definedName.scopeSheetId === sheetId &&
              definedName.name.toLowerCase() === w.toLowerCase(),
          ) ??
          activeWorkbook.definedNames?.find(
            (definedName) =>
              !definedName.scopeSheetId &&
              definedName.name.toLowerCase() === w.toLowerCase(),
          );
        if (!E)
          throw Error(
            "The address or defined range name was not found in the current worksheet scope.",
          );
        const A = `${E.scopeSheetId ?? "global"}:${E.name.toLowerCase()}`;
        if (f.includes(A))
          throw Error("The defined range name contains a circular reference.");
        l(E.refersTo.slice(1).trim(), E.scopeSheetId ?? sheetId, [...f, A]);
      }
    };
  if (
    (l(goToReferenceText.trim(), r, []),
    d.some((c) => c.sheetId !== d[0].sheetId))
  )
    throw Error("Select ranges on one worksheet at a time.");
  return {
    sheetId: d[0].sheetId,
    ranges: parseSelectionAreas(d.map((c) => c.range)),
  };
}

function describeStyleFormatting(e) {
  const effectiveCellStyle = resolveEffectiveCellStyle(e);
  return (
    [
      effectiveCellStyle.fontFamily,
      effectiveCellStyle.fontSize
        ? `${effectiveCellStyle.fontSize} pt`
        : undefined,
      effectiveCellStyle.bold ? "Bold" : undefined,
      effectiveCellStyle.italic ? "Italic" : undefined,
      effectiveCellStyle.fill ? `Fill ${effectiveCellStyle.fill}` : undefined,
      effectiveCellStyle.color ? `Text ${effectiveCellStyle.color}` : undefined,
      effectiveCellStyle.borders || effectiveCellStyle.border
        ? "Borders"
        : undefined,
      effectiveCellStyle.wrap ? "Wrap" : undefined,
    ]
      .filter(Boolean)
      .join(", ") || "Default formatting"
  );
}
const styleFormattingLabels = {
  fontFamily: "Font",
  fontSize: "Font size",
  bold: "Bold",
  italic: "Italic",
  underline: "Underline",
  strikethrough: "Strikethrough",
  fill: "Fill",
  color: "Text color",
  align: "Horizontal alignment",
  verticalAlign: "Vertical alignment",
  indent: "Indent",
  wrap: "Wrap",
  border: "Legacy border",
};
function CellStylePreview({ style }) {
  const effectiveCellStyle = resolveEffectiveCellStyle(style);
  return i.jsxs("dl", {
    className: "ts-detail-fields",
    children: [
      Object.entries(styleFormattingLabels).map(([t, n]) =>
        i.jsxs(
          "div",
          {
            children: [
              i.jsx("dt", { children: n }),
              i.jsx("dd", {
                children:
                  effectiveCellStyle[t] === undefined
                    ? "Default"
                    : typeof effectiveCellStyle[t] == "boolean"
                      ? effectiveCellStyle[t]
                        ? "On"
                        : "Off"
                      : String(effectiveCellStyle[t]) +
                        (t === "fontSize" ? " pt" : ""),
              }),
            ],
          },
          t,
        ),
      ),
      ["top", "right", "bottom", "left"].map((cellBorderEdge) => {
        const borderEdge = resolveBorderEdge(
          effectiveCellStyle,
          cellBorderEdge,
        );
        return i.jsxs(
          "div",
          {
            children: [
              i.jsxs("dt", {
                children: [
                  cellBorderEdge[0].toUpperCase() + cellBorderEdge.slice(1),
                  " ",
                  i.jsx(LocalizedText, { children: "border" }),
                ],
              }),
              i.jsx("dd", {
                children: borderEdge
                  ? `${borderEdge.style}${borderEdge.color ? ` ${borderEdge.color}` : ""}`
                  : "Default gridline",
              }),
            ],
          },
          cellBorderEdge,
        );
      }),
    ],
  });
}
function CellStyleSummary({ style, catalog }) {
  const t = style?.binding,
    { binding, ...o } = style ?? {};
  return i.jsxs("details", {
    children: [
      i.jsx("summary", {
        children: i.jsx(LocalizedText, { children: "Full formatting" }),
      }),
      i.jsx("h4", {
        children: i.jsx(LocalizedText, { children: "Displayed formatting" }),
      }),
      i.jsx(CellStylePreview, { style }),
      t &&
        i.jsxs(i.Fragment, {
          children: [
            i.jsxs("h4", {
              children: [
                i.jsx(LocalizedText, { children: "Named style:" }),
                " ",
                catalog?.styles.find((s) => s.id === t.styleId)?.name ??
                  "Missing style",
              ],
            }),
            t.themeFallback &&
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children: "Office theme fallback",
                }),
              }),
            i.jsx(CellStylePreview, { style: t.base }),
            i.jsx("h4", {
              children: i.jsx(LocalizedText, { children: "Direct overrides" }),
            }),
            Object.keys(o).length
              ? i.jsx("dl", {
                  className: "ts-detail-fields",
                  children: Object.entries(o).map(([s, d]) =>
                    i.jsxs(
                      "div",
                      {
                        children: [
                          i.jsx("dt", {
                            children: styleFormattingLabels[s] ?? s,
                          }),
                          i.jsx("dd", {
                            children:
                              s === "borders"
                                ? Object.entries(d)
                                    .map(
                                      ([u, l]) =>
                                        `${u}: ${l.style}${l.color ? ` ${l.color}` : ""}`,
                                    )
                                    .join("; ")
                                : typeof d == "boolean"
                                  ? d
                                    ? "On"
                                    : "Off"
                                  : String(d),
                          }),
                        ],
                      },
                      s,
                    ),
                  ),
                })
              : i.jsx("p", {
                  children: i.jsx(LocalizedText, { children: "None" }),
                }),
          ],
        }),
    ],
  });
}
function ThemeRoleBindings({ definition }) {
  return definition
    ? i.jsxs(i.Fragment, {
        children: [
          i.jsx(CellStylePreview, { style: definition.base }),
          i.jsxs("dl", {
            className: "ts-detail-fields",
            children: [
              i.jsxs("div", {
                children: [
                  i.jsx("dt", {
                    children: i.jsx(LocalizedText, {
                      children: "Theme font role",
                    }),
                  }),
                  i.jsx("dd", {
                    children: definition.fontRole ?? "Stored font",
                  }),
                ],
              }),
              i.jsxs("div", {
                children: [
                  i.jsx("dt", {
                    children: i.jsx(LocalizedText, {
                      children: "Theme fill role",
                    }),
                  }),
                  i.jsx("dd", {
                    children: definition.fillRole ?? "Stored fill",
                  }),
                ],
              }),
              i.jsxs("div", {
                children: [
                  i.jsx("dt", {
                    children: i.jsx(LocalizedText, {
                      children: "Theme text role",
                    }),
                  }),
                  i.jsx("dd", {
                    children: definition.colorRole ?? "Stored text color",
                  }),
                ],
              }),
              ["top", "right", "bottom", "left"].map((r) =>
                i.jsxs(
                  "div",
                  {
                    children: [
                      i.jsxs("dt", {
                        children: [
                          r,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "border color role",
                          }),
                        ],
                      }),
                      i.jsx("dd", {
                        children:
                          definition.borderRoles?.[r] ?? "Stored border color",
                      }),
                    ],
                  },
                  r,
                ),
              ),
            ],
          }),
        ],
      })
    : i.jsx("p", { children: i.jsx(LocalizedText, { children: "No style" }) });
}
function ThemeChangePreview({ plan }) {
  const catalogBefore = plan.catalogBefore,
    catalogAfter = plan.catalogAfter;
  if (JSON.stringify(catalogBefore) === JSON.stringify(catalogAfter))
    return null;
  const n = [
      {
        name: "Theme name",
        before: catalogBefore?.theme.name,
        after: catalogAfter?.theme.name,
      },
      {
        name: "Body font",
        before: catalogBefore?.theme.fonts.body,
        after: catalogAfter?.theme.fonts.body,
      },
      {
        name: "Heading font",
        before: catalogBefore?.theme.fonts.heading,
        after: catalogAfter?.theme.fonts.heading,
      },
      ...themePaletteRoles.map((themePaletteRole) => ({
        name: `${themePaletteRole} color`,
        before: catalogBefore?.theme.palette[themePaletteRole],
        after: catalogAfter?.theme.palette[themePaletteRole],
      })),
      {
        name: "Missing theme fallback",
        before: catalogBefore?.missingTheme ? "Office fallback" : "None",
        after: catalogAfter?.missingTheme ? "Office fallback" : "None",
      },
    ].filter((a) => a.before !== a.after),
    o = [
      ...new Set([
        ...(catalogBefore?.styles.map((style) => style.id) ?? []),
        ...(catalogAfter?.styles.map((style) => style.id) ?? []),
      ]),
    ];
  return i.jsxs("details", {
    className: "ts-catalog-preview",
    open: true,
    children: [
      i.jsx("summary", {
        children: i.jsx(LocalizedText, { children: "Workbook style settings" }),
      }),
      n.length > 0 &&
        i.jsxs("table", {
          children: [
            i.jsx("thead", {
              children: i.jsxs("tr", {
                children: [
                  i.jsx("th", {
                    children: i.jsx(LocalizedText, { children: "Setting" }),
                  }),
                  i.jsx("th", {
                    children: i.jsx(LocalizedText, { children: "Current" }),
                  }),
                  i.jsx("th", {
                    children: i.jsx(LocalizedText, {
                      children: "After change",
                    }),
                  }),
                ],
              }),
            }),
            i.jsx("tbody", {
              children: n.map((a) =>
                i.jsxs(
                  "tr",
                  {
                    children: [
                      i.jsx("th", { children: a.name }),
                      i.jsx("td", { children: a.before ?? "Not set" }),
                      i.jsx("td", { children: a.after ?? "Not set" }),
                    ],
                  },
                  a.name,
                ),
              ),
            }),
          ],
        }),
      o.map((a) => {
        const s = catalogBefore?.styles.find((style) => style.id === a),
          d = catalogAfter?.styles.find((style) => style.id === a);
        return JSON.stringify(s) === JSON.stringify(d)
          ? null
          : i.jsxs(
              "details",
              {
                className: "ts-style-definition",
                children: [
                  i.jsx("summary", {
                    children: s
                      ? d
                        ? `${s.name} → ${d.name}`
                        : `Delete ${s.name}`
                      : `Create ${d?.name}`,
                  }),
                  i.jsxs("div", {
                    className: "ts-definition-columns",
                    children: [
                      i.jsxs("div", {
                        children: [
                          i.jsx("h4", {
                            children: i.jsx(LocalizedText, {
                              children: "Current definition",
                            }),
                          }),
                          i.jsx(ThemeRoleBindings, {
                            definition: s?.definition,
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        children: [
                          i.jsx("h4", {
                            children: i.jsx(LocalizedText, {
                              children: "After change",
                            }),
                          }),
                          i.jsx(ThemeRoleBindings, {
                            definition: d?.definition,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              },
              a,
            );
      }),
    ],
  });
}
function ThemesAndCellStylesDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const [s] = ne.useState(() => ({
      snapshot: JSON.stringify(e.workbook),
      sheet: e.sheetId,
      ranges: JSON.stringify(e.ranges ?? [e.range]),
      active: e.active,
    })),
    [d] = ne.useState(() =>
      cloneStructuredValue(
        e.workbook.styleCatalog ?? createDefaultStyleCatalog(e.workbook.id),
      ),
    ),
    selectionAreas = parseSelectionAreas(e.ranges ?? [e.range]),
    rangeAddress = parseRangeAddress(e.range),
    c = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    h =
      e.active && isCellAddressInRanges(selectionAreas, e.active)
        ? e.active
        : `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    f = c?.cells[h]?.style?.binding?.styleId,
    [styleId, m] = ne.useState(
      d.styles.some((style) => style.id === f) ? f : (d.styles[0]?.id ?? ""),
    ),
    v = d.styles.find((style) => style.id === styleId),
    [g, b] = ne.useState("Theme"),
    [theme, w] = ne.useState(cloneStructuredValue(d.theme)),
    [definition, A] = ne.useState(
      cloneStructuredValue(v?.definition ?? { base: {} }),
    ),
    [k, F] = ne.useState("Custom style"),
    [I, O] = ne.useState(v?.name ?? ""),
    [plan, z] = ne.useState(),
    [V, D] = ne.useState(""),
    [B, _] = ne.useState(0),
    C =
      s.snapshot !== JSON.stringify(e.workbook) ||
      s.sheet !== e.sheetId ||
      s.ranges !== JSON.stringify(e.ranges ?? [e.range]) ||
      s.active !== e.active,
    disabled = e.readOnly || C || !c,
    N =
      disabled ||
      !!c?.protection?.enabled ||
      countSelectionCells(selectionAreas) > 500;
  function R() {
    (z(undefined), D(""), _(0));
  }
  function L(de) {
    (w(de), R());
  }
  function $(de) {
    (A(de), R());
  }
  function U(request) {
    (D(""), _(0));
    try {
      if (C)
        throw Error(
          "The workbook or selection changed. Close and reopen Themes & Cell Styles.",
        );
      z(applyNamedStyleChange(e.workbook, request, { readOnly: e.readOnly }));
    } catch (error) {
      (z(undefined),
        D(
          error instanceof Error
            ? error.message
            : "Unable to preview this style change.",
        ));
    }
  }
  function onClick() {
    try {
      if (!plan) return;
      if (C)
        throw Error(
          "The workbook or selection changed. Close and reopen Themes & Cell Styles.",
        );
      const de = assertThemePreviewCurrent(e.workbook, plan, {
        readOnly: e.readOnly,
      });
      (de !== e.workbook && e.onWorkbookChange(de, plan.summary), e.onClose());
    } catch (error) {
      D(
        error instanceof Error
          ? error.message
          : "Unable to apply this style change.",
      );
    }
  }
  function H(de) {
    m(de);
    const he = d.styles.find((style) => style.id === de);
    (A(cloneStructuredValue(he?.definition ?? { base: {} })),
      O(he?.name ?? ""),
      R());
  }
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      o.current?.querySelector('[role="tab"][aria-selected="true"]')?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const styleFallbackWarnings = collectStyleFallbackWarnings(e.workbook),
    X$1 = Math.max(1, Math.ceil((plan?.count ?? 0) / 20));
  let J = {};
  try {
    J = resolveNamedStyleAppearance(definition, theme);
  } catch {}
  return i.jsx("div", {
    className: "ts-backdrop",
    children: i.jsxs("div", {
      ref: o,
      className: "ts-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), a.current()),
          event.key === "Tab")
        ) {
          const Se = Array.from(
              o.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ).filter((ge) => !ge.hidden),
            le = Se[0],
            ce = Se.at(-1);
          event.shiftKey && document.activeElement === le
            ? (event.preventDefault(), ce == null || ce.focus())
            : !event.shiftKey &&
              document.activeElement === ce &&
              (event.preventDefault(), le == null || le.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "ts-header",
          children: [
            i.jsx("h2", {
              id: `${n}-title`,
              children: i.jsx(LocalizedText, {
                children: "Themes & Cell Styles",
              }),
            }),
            i.jsx("button", {
              "aria-label": text("Close Themes and Cell Styles"),
              onClick: e.onClose,
              lang: language("Close Themes and Cell Styles"),
              "data-ui-label": "Close Themes and Cell Styles",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "ts-location",
          children: [
            c?.name ?? "Missing sheet",
            " · ",
            selectionAreas.join(", "),
            " ·",
            " ",
            countSelectionCells(selectionAreas).toLocaleString(),
            " ",
            i.jsx(LocalizedText, { children: "selected cells" }),
          ],
        }),
        i.jsx("div", {
          className: "ts-tabs",
          role: "tablist",
          "aria-label": text("Theme settings"),
          lang: language("Theme settings"),
          "data-ui-label": "Theme settings",
          children: ["Theme", "Cell Styles"].map((de, he) =>
            i.jsx(
              "button",
              {
                role: "tab",
                "aria-selected": g === de,
                "aria-controls": `${n}-panel`,
                tabIndex: g === de ? 0 : -1,
                onClick: () => b(de),
                onKeyDown: (event) => {
                  if (
                    ["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                      event.key,
                    )
                  ) {
                    event.preventDefault();
                    const ge =
                      event.key === "Home"
                        ? 0
                        : event.key === "End"
                          ? 1
                          : 1 - he;
                    (b(ge === 0 ? "Theme" : "Cell Styles"),
                      event.currentTarget.parentElement?.children[ge]?.focus());
                  }
                },
                children: de,
              },
              de,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "ts-body",
          children: [
            (e.readOnly || C || styleFallbackWarnings.length > 0) &&
              i.jsx("div", {
                role: "status",
                className: "ts-notice",
                children: e.readOnly
                  ? "This workbook is read-only."
                  : C
                    ? "The workbook or selection changed. Close and reopen this dialog."
                    : styleFallbackWarnings.join(" "),
              }),
            i.jsx("section", {
              id: `${n}-panel`,
              role: "tabpanel",
              "aria-label": g,
              children:
                g === "Theme"
                  ? i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("p", {
                          children: i.jsx(LocalizedText, {
                            children:
                              "Theme colors and fonts update cells using named styles. Direct cell formatting remains in place.",
                          }),
                        }),
                        i.jsxs("div", {
                          className: "ts-fields",
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Theme preset",
                                }),
                                i.jsxs("select", {
                                  "aria-label": text("Theme preset"),
                                  value: builtInWorkbookThemes.some(
                                    (builtInWorkbookTheme) =>
                                      JSON.stringify(builtInWorkbookTheme) ===
                                      JSON.stringify(theme),
                                  )
                                    ? theme.name
                                    : "Custom",
                                  disabled: disabled,
                                  onChange: (event) => {
                                    const he = builtInWorkbookThemes.find(
                                      (builtInWorkbookTheme) =>
                                        builtInWorkbookTheme.name ===
                                        event.target.value,
                                    );
                                    he && L(cloneStructuredValue(he));
                                  },
                                  lang: language("Theme preset"),
                                  "data-ui-label": "Theme preset",
                                  children: [
                                    i.jsx("option", {
                                      value: "Custom",
                                      children: "Custom",
                                    }),
                                    builtInWorkbookThemes.map(
                                      (builtInWorkbookTheme) =>
                                        i.jsx(
                                          "option",
                                          {
                                            children: builtInWorkbookTheme.name,
                                          },
                                          builtInWorkbookTheme.name,
                                        ),
                                    ),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Theme name",
                                }),
                                i.jsx("input", {
                                  value: theme.name,
                                  maxLength: 80,
                                  disabled: disabled,
                                  onChange: (event) =>
                                    L({ ...theme, name: event.target.value }),
                                }),
                              ],
                            }),
                            ["body", "heading"].map((de) =>
                              i.jsxs(
                                "label",
                                {
                                  children: [
                                    de === "body"
                                      ? "Body font"
                                      : "Heading font",
                                    i.jsx("select", {
                                      "aria-label":
                                        de === "body"
                                          ? "Body font"
                                          : "Heading font",
                                      value: theme.fonts[de],
                                      disabled: disabled,
                                      onChange: (event) =>
                                        L({
                                          ...theme,
                                          fonts: {
                                            ...theme.fonts,
                                            [de]: event.target.value,
                                          },
                                        }),
                                      children: supportedFontFamilies.map(
                                        (supportedFontFamily) =>
                                          i.jsx(
                                            "option",
                                            { children: supportedFontFamily },
                                            supportedFontFamily,
                                          ),
                                      ),
                                    }),
                                  ],
                                },
                                de,
                              ),
                            ),
                          ],
                        }),
                        i.jsxs("fieldset", {
                          className: "ts-palette",
                          disabled: disabled,
                          children: [
                            i.jsx("legend", {
                              children: i.jsx(LocalizedText, {
                                children: "Theme colors",
                              }),
                            }),
                            themePaletteRoles.map((themePaletteRole) =>
                              i.jsxs(
                                "label",
                                {
                                  children: [
                                    i.jsx("span", {
                                      children: themePaletteRole
                                        .replace("accent", "Accent ")
                                        .replace(/^background$/, "Background")
                                        .replace(/^text$/, "Text"),
                                    }),
                                    i.jsx("input", {
                                      "aria-label": `Theme ${themePaletteRole} color`,
                                      type: "color",
                                      value: theme.palette[themePaletteRole],
                                      onChange: (event) =>
                                        L({
                                          ...theme,
                                          palette: {
                                            ...theme.palette,
                                            [themePaletteRole]:
                                              event.target.value.toUpperCase(),
                                          },
                                        }),
                                    }),
                                  ],
                                },
                                themePaletteRole,
                              ),
                            ),
                          ],
                        }),
                        i.jsx("button", {
                          className: "ts-command",
                          disabled: disabled,
                          onClick: () => U({ kind: "setTheme", theme: theme }),
                          children: i.jsx(LocalizedText, {
                            children: "Preview theme change",
                          }),
                        }),
                        i.jsx("p", {
                          className: "ts-help",
                          children: i.jsx(LocalizedText, {
                            children:
                              "A theme change may affect multiple sheets. The preview is limited to 500 changed cells and respects protected sheets.",
                          }),
                        }),
                      ],
                    })
                  : i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("div", {
                          className: "ts-fields",
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Named style",
                                }),
                                i.jsxs("select", {
                                  "aria-label": text("Named style"),
                                  value: styleId,
                                  disabled: disabled || !d.styles.length,
                                  onChange: (event) => H(event.target.value),
                                  lang: language("Named style"),
                                  "data-ui-label": "Named style",
                                  children: [
                                    !d.styles.length &&
                                      i.jsx("option", {
                                        value: "",
                                        children: "No named styles",
                                      }),
                                    d.styles.map((style) =>
                                      i.jsx(
                                        "option",
                                        {
                                          value: style.id,
                                          children: style.name,
                                        },
                                        style.id,
                                      ),
                                    ),
                                  ],
                                }),
                              ],
                            }),
                            i.jsx("div", {
                              className: "ts-sample",
                              style: cellStyleToCss(J),
                              children: i.jsx(LocalizedText, {
                                children: "Audit schedule · 1,234.50",
                              }),
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "ts-actions",
                          children: [
                            i.jsx("button", {
                              disabled: N || !v,
                              onClick: () =>
                                U({
                                  kind: "applyStyle",
                                  sheetId: e.sheetId,
                                  ranges: selectionAreas,
                                  styleId: styleId,
                                }),
                              children: i.jsx(LocalizedText, {
                                children: "Apply style to selection",
                              }),
                            }),
                            i.jsx("button", {
                              disabled: N || !v,
                              onClick: () =>
                                U({
                                  kind: "reapplyStyle",
                                  sheetId: e.sheetId,
                                  ranges: selectionAreas,
                                  styleId: styleId,
                                }),
                              children: i.jsx(LocalizedText, {
                                children: "Reapply and clear direct styles",
                              }),
                            }),
                            i.jsx("button", {
                              disabled: N,
                              onClick: () =>
                                U({
                                  kind: "detachStyle",
                                  sheetId: e.sheetId,
                                  ranges: selectionAreas,
                                }),
                              children: i.jsx(LocalizedText, {
                                children: "Detach and keep appearance",
                              }),
                            }),
                          ],
                        }),
                        i.jsx("p", {
                          className: "ts-help",
                          children: i.jsx(LocalizedText, {
                            children:
                              "Apply keeps direct formatting. Reapply clears direct font, fill, alignment and border overrides. Values, formulas, number formats and review evidence are preserved.",
                          }),
                        }),
                        c?.protection?.enabled &&
                          i.jsx("p", {
                            className: "ts-notice",
                            children: i.jsx(LocalizedText, {
                              children:
                                "This sheet is protected. Unprotect it before applying a style.",
                            }),
                          }),
                        i.jsxs("fieldset", {
                          disabled: disabled,
                          children: [
                            i.jsx("legend", {
                              children: i.jsx(LocalizedText, {
                                children: "Create or update a reusable style",
                              }),
                            }),
                            i.jsxs("div", {
                              className: "ts-fields",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "New style name",
                                    }),
                                    i.jsx("input", {
                                      value: k,
                                      maxLength: 80,
                                      onChange: (event) => {
                                        (F(event.target.value), R());
                                      },
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Style font",
                                    }),
                                    i.jsxs("select", {
                                      "aria-label": text("Style font"),
                                      value: definition.fontRole ?? "fixed",
                                      onChange: (event) =>
                                        $({
                                          ...definition,
                                          fontRole:
                                            event.target.value === "fixed"
                                              ? undefined
                                              : event.target.value,
                                        }),
                                      lang: language("Style font"),
                                      "data-ui-label": "Style font",
                                      children: [
                                        i.jsx("option", {
                                          value: "fixed",
                                          children: "Keep stored font",
                                        }),
                                        i.jsx("option", {
                                          value: "body",
                                          children: "Theme body font",
                                        }),
                                        i.jsx("option", {
                                          value: "heading",
                                          children: "Theme heading font",
                                        }),
                                      ],
                                    }),
                                  ],
                                }),
                                ["fillRole", "colorRole"].map((de) =>
                                  i.jsxs(
                                    "label",
                                    {
                                      children: [
                                        de === "fillRole"
                                          ? "Style fill"
                                          : "Style text color",
                                        i.jsxs("select", {
                                          "aria-label":
                                            de === "fillRole"
                                              ? "Style fill"
                                              : "Style text color",
                                          value: definition[de] ?? "fixed",
                                          onChange: (event) =>
                                            $({
                                              ...definition,
                                              [de]:
                                                event.target.value === "fixed"
                                                  ? undefined
                                                  : event.target.value,
                                            }),
                                          children: [
                                            i.jsx("option", {
                                              value: "fixed",
                                              children: "Keep stored color",
                                            }),
                                            themePaletteRoles.map(
                                              (themePaletteRole) =>
                                                i.jsx(
                                                  "option",
                                                  {
                                                    value: themePaletteRole,
                                                    children: themePaletteRole,
                                                  },
                                                  themePaletteRole,
                                                ),
                                            ),
                                          ],
                                        }),
                                      ],
                                    },
                                    de,
                                  ),
                                ),
                                i.jsxs("label", {
                                  className: "ts-check",
                                  children: [
                                    i.jsx("input", {
                                      type: "checkbox",
                                      checked: definition.base.bold ?? false,
                                      onChange: (event) =>
                                        $({
                                          ...definition,
                                          base: {
                                            ...definition.base,
                                            bold: event.target.checked,
                                          },
                                        }),
                                    }),
                                    i.jsx(LocalizedText, { children: "Bold" }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  className: "ts-check",
                                  children: [
                                    i.jsx("input", {
                                      type: "checkbox",
                                      checked: definition.base.italic ?? false,
                                      onChange: (event) =>
                                        $({
                                          ...definition,
                                          base: {
                                            ...definition.base,
                                            italic: event.target.checked,
                                          },
                                        }),
                                    }),
                                    i.jsx(LocalizedText, {
                                      children: "Italic",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("div", {
                              className: "ts-actions",
                              children: [
                                i.jsx("button", {
                                  onClick: () =>
                                    $({
                                      base: cloneStructuredValue(
                                        resolveEffectiveCellStyle(
                                          c?.cells[h]?.style,
                                        ),
                                      ),
                                    }),
                                  children: i.jsx(LocalizedText, {
                                    children: "Use active cell formatting",
                                  }),
                                }),
                                i.jsx("button", {
                                  onClick: () =>
                                    U({
                                      kind: "createStyle",
                                      name: k,
                                      definition: definition,
                                    }),
                                  children: i.jsx(LocalizedText, {
                                    children: "Preview new style",
                                  }),
                                }),
                                i.jsx("button", {
                                  disabled: !v,
                                  onClick: () =>
                                    U({
                                      kind: "updateStyle",
                                      styleId: styleId,
                                      definition: definition,
                                    }),
                                  children: i.jsx(LocalizedText, {
                                    children: "Preview style update",
                                  }),
                                }),
                              ],
                            }),
                            i.jsx("p", {
                              className: "ts-help",
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Stored formatting can include the active cell’s supported border edges. Theme roles above bind fonts and colors. Up to 50 workbook styles.",
                              }),
                            }),
                          ],
                        }),
                        i.jsxs("fieldset", {
                          disabled: disabled || !v,
                          children: [
                            i.jsx("legend", {
                              children: i.jsx(LocalizedText, {
                                children: "Manage selected style",
                              }),
                            }),
                            i.jsxs("div", {
                              className: "ts-fields",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Rename style",
                                    }),
                                    i.jsx("input", {
                                      value: I,
                                      maxLength: 80,
                                      onChange: (event) => {
                                        (O(event.target.value), R());
                                      },
                                    }),
                                  ],
                                }),
                                i.jsxs("div", {
                                  className: "ts-actions",
                                  children: [
                                    i.jsx("button", {
                                      onClick: () =>
                                        U({
                                          kind: "renameStyle",
                                          styleId: styleId,
                                          name: I,
                                        }),
                                      children: i.jsx(LocalizedText, {
                                        children: "Preview rename",
                                      }),
                                    }),
                                    i.jsx("button", {
                                      onClick: () =>
                                        U({
                                          kind: "deleteStyle",
                                          styleId: styleId,
                                        }),
                                      children: i.jsx(LocalizedText, {
                                        children: "Preview deletion",
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
            }),
            V &&
              i.jsx("div", {
                role: "alert",
                className: "ts-error",
                children: V,
              }),
            plan &&
              i.jsxs("section", {
                className: "ts-preview",
                "aria-label": text("Theme change preview"),
                lang: language("Theme change preview"),
                "data-ui-label": "Theme change preview",
                children: [
                  i.jsx("h3", { children: plan.summary }),
                  i.jsxs("p", {
                    children: [
                      plan.count,
                      " ",
                      i.jsx(LocalizedText, { children: "cell" }),
                      plan.count === 1 ? "" : "s",
                      " ",
                      i.jsx(LocalizedText, { children: "will change" }),
                      JSON.stringify(plan.catalogBefore) !==
                      JSON.stringify(plan.catalogAfter)
                        ? "; workbook style settings will also change"
                        : "",
                      ".",
                    ],
                  }),
                  i.jsx(ThemeChangePreview, { plan: plan }),
                  plan.warnings.map((warning) =>
                    i.jsx(
                      "p",
                      { className: "ts-notice", children: warning },
                      warning,
                    ),
                  ),
                  plan.count > 0 &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("div", {
                          className: "ts-preview-table",
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
                                        children: "Current appearance",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "After change",
                                      }),
                                    }),
                                  ],
                                }),
                              }),
                              i.jsx("tbody", {
                                children: plan.changes
                                  .slice(B * 20, B * 20 + 20)
                                  .map((de) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("th", {
                                            children: [
                                              e.workbook.sheets.find(
                                                (sheet) =>
                                                  sheet.id === de.sheetId,
                                              )?.name,
                                              "!",
                                              de.address,
                                            ],
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              describeStyleFormatting(
                                                de.before?.style,
                                              ),
                                              i.jsx(CellStyleSummary, {
                                                style: de.before?.style,
                                                catalog: plan.catalogBefore,
                                              }),
                                            ],
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              describeStyleFormatting(
                                                de.after.style,
                                              ),
                                              i.jsx(CellStyleSummary, {
                                                style: de.after.style,
                                                catalog: plan.catalogAfter,
                                              }),
                                            ],
                                          }),
                                        ],
                                      },
                                      `${de.sheetId}:${de.address}`,
                                    ),
                                  ),
                              }),
                            ],
                          }),
                        }),
                        i.jsxs("div", {
                          className: "ts-pager",
                          children: [
                            i.jsx("button", {
                              disabled: B === 0,
                              onClick: () => _((de) => de - 1),
                              children: i.jsx(LocalizedText, {
                                children: "Previous changes",
                              }),
                            }),
                            i.jsxs("span", {
                              children: [
                                i.jsx(LocalizedText, { children: "Page " }),
                                B + 1,
                                " ",
                                i.jsx(LocalizedText, { children: "of " }),
                                X$1,
                              ],
                            }),
                            i.jsx("button", {
                              disabled: B + 1 >= X$1,
                              onClick: () => _((de) => de + 1),
                              children: i.jsx(LocalizedText, {
                                children: "Next changes",
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
        i.jsxs("footer", {
          className: "ts-footer",
          children: [
            i.jsx("span", {
              children: plan
                ? "Review the preview before applying."
                : "Choose a change to preview.",
            }),
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "ts-primary",
              disabled: disabled || !plan,
              onClick: onClick,
              children: i.jsx(LocalizedText, { children: "Apply changes" }),
            }),
          ],
        }),
      ],
    }),
  });
}
function CellValuePreview({ value, cell, dateSystem }) {
  const workbookMessages = useWorkbookMessages(),
    o = cell?.numberFormat
      ? formatCellValueWithNumberFormat(value, cell.numberFormat, dateSystem)
      : undefined,
    a = cell ? getImportedNumberFormatWarning(cell) : undefined;
  if (o?.accounting) {
    const { symbol, amount, symbolPosition } = o.accounting;
    return i.jsxs("span", {
      title: a || undefined,
      "aria-label": o.text,
      style: {
        display: "inline-flex",
        width: "100%",
        minWidth: 0,
        justifyContent: "space-between",
        gap: 6,
        fontVariantNumeric: "tabular-nums",
      },
      children: [
        symbolPosition === "start" &&
          i.jsx("span", { "aria-hidden": "true", children: symbol }),
        i.jsx("span", {
          "aria-hidden": "true",
          style: { textAlign: "right", flex: 1 },
          children: amount,
        }),
        symbolPosition === "end" &&
          i.jsx("span", { "aria-hidden": "true", children: symbol }),
      ],
    });
  }
  return i.jsx("span", {
    title: a || undefined,
    children:
      o?.text ??
      formatLocalizedNumberDisplay(
        value,
        cell,
        dateSystem,
        workbookMessages.entryLocale,
      ),
  });
}
const cellFormatTabs = ["Font", "Alignment", "Border"],
  sampleGridAddresses = new Set(
    Array.from(
      { length: 9 },
      (e, r) => `${columnIndexToLetters(r % 3)}${Math.floor(r / 3) + 1}`,
    ),
  ),
  expandShortHexColor = (e, r) =>
    e?.length === 4
      ? "#" +
        e
          .slice(1)
          .split("")
          .map((t) => t + t)
          .join("")
      : (e ?? r);

export {
  expandShortHexColor,
  parseGoToReference,
  ConnectionStatusNotice,
  ThemesAndCellStylesDialog,
  sampleGridAddresses,
  CellValuePreview,
  cellFormatTabs,
};
