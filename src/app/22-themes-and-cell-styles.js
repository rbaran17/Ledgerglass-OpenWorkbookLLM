import { cloneStructuredValue } from "./11-workbook-formats-dialog.js";
import i from "react/jsx-runtime";
import ne from "react";
import { X } from "lucide-react";
import {
  maxSelectionAreas,
  supportedFontFamilies,
} from "./02-financial-functions.js";
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
function parseGoToReference(e, r, t, n = {}) {
  if (typeof t != "string" || !t.trim() || t.length > 8192)
    throw Error(
      "Enter an address, range, or defined range name (up to 8,192 characters).",
    );
  const o = n.allowedSheetIds,
    a = (c) => !o || o.includes(c),
    s = (c) => {
      const h = e.sheets.find((f) => f.id === c && a(f.id));
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
        let y = h,
          w = g;
        if (b) {
          const k = b[1]?.replace(/''/g, "'") ?? b[2].trim(),
            F = e.sheets.find(
              (I) => a(I.id) && I.name.toLowerCase() === k.toLowerCase(),
            );
          if (!F)
            throw Error(
              "The destination is outside the available worksheet scope.",
            );
          ((y = F.id), (w = b[3].trim()));
        }
        if (
          (s(y),
          /^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/.test(
            w,
          ))
        ) {
          if (
            (d.push({ sheetId: y, range: parseSelectionAreas(w)[0] }),
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
          e.definedNames?.find(
            (k) =>
              k.scopeSheetId === y && k.name.toLowerCase() === w.toLowerCase(),
          ) ??
          e.definedNames?.find(
            (k) => !k.scopeSheetId && k.name.toLowerCase() === w.toLowerCase(),
          );
        if (!E)
          throw Error(
            "The address or defined range name was not found in the current worksheet scope.",
          );
        const A = `${E.scopeSheetId ?? "global"}:${E.name.toLowerCase()}`;
        if (f.includes(A))
          throw Error("The defined range name contains a circular reference.");
        l(E.refersTo.slice(1).trim(), E.scopeSheetId ?? y, [...f, A]);
      }
    };
  if ((l(t.trim(), r, []), d.some((c) => c.sheetId !== d[0].sheetId)))
    throw Error("Select ranges on one worksheet at a time.");
  return {
    sheetId: d[0].sheetId,
    ranges: parseSelectionAreas(d.map((c) => c.range)),
  };
}

function describeStyleFormatting(e) {
  const r = resolveEffectiveCellStyle(e);
  return (
    [
      r.fontFamily,
      r.fontSize ? `${r.fontSize} pt` : undefined,
      r.bold ? "Bold" : undefined,
      r.italic ? "Italic" : undefined,
      r.fill ? `Fill ${r.fill}` : undefined,
      r.color ? `Text ${r.color}` : undefined,
      r.borders || r.border ? "Borders" : undefined,
      r.wrap ? "Wrap" : undefined,
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
  const r = resolveEffectiveCellStyle(style);
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
                  r[t] === undefined
                    ? "Default"
                    : typeof r[t] == "boolean"
                      ? r[t]
                        ? "On"
                        : "Off"
                      : String(r[t]) + (t === "fontSize" ? " pt" : ""),
              }),
            ],
          },
          t,
        ),
      ),
      ["top", "right", "bottom", "left"].map((t) => {
        const n = resolveBorderEdge(r, t);
        return i.jsxs(
          "div",
          {
            children: [
              i.jsxs("dt", {
                children: [
                  t[0].toUpperCase() + t.slice(1),
                  " ",
                  i.jsx(LocalizedText, { children: "border" }),
                ],
              }),
              i.jsx("dd", {
                children: n
                  ? `${n.style}${n.color ? ` ${n.color}` : ""}`
                  : "Default gridline",
              }),
            ],
          },
          t,
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
  const r = plan.catalogBefore,
    t = plan.catalogAfter;
  if (JSON.stringify(r) === JSON.stringify(t)) return null;
  const n = [
      {
        name: "Theme name",
        before: r?.theme.name,
        after: t?.theme.name,
      },
      {
        name: "Body font",
        before: r?.theme.fonts.body,
        after: t?.theme.fonts.body,
      },
      {
        name: "Heading font",
        before: r?.theme.fonts.heading,
        after: t?.theme.fonts.heading,
      },
      ...themePaletteRoles.map((a) => ({
        name: `${a} color`,
        before: r?.theme.palette[a],
        after: t?.theme.palette[a],
      })),
      {
        name: "Missing theme fallback",
        before: r?.missingTheme ? "Office fallback" : "None",
        after: t?.missingTheme ? "Office fallback" : "None",
      },
    ].filter((a) => a.before !== a.after),
    o = [
      ...new Set([
        ...(r?.styles.map((a) => a.id) ?? []),
        ...(t?.styles.map((a) => a.id) ?? []),
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
        const s = r?.styles.find((u) => u.id === a),
          d = t?.styles.find((u) => u.id === a);
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
    u = parseSelectionAreas(e.ranges ?? [e.range]),
    l = parseRangeAddress(e.range),
    c = e.workbook.sheets.find((de) => de.id === e.sheetId),
    h =
      e.active && isCellAddressInRanges(u, e.active)
        ? e.active
        : `${columnIndexToLetters(l.firstCol)}${l.firstRow + 1}`,
    f = c?.cells[h]?.style?.binding?.styleId,
    [p, m] = ne.useState(
      d.styles.some((de) => de.id === f) ? f : (d.styles[0]?.id ?? ""),
    ),
    v = d.styles.find((de) => de.id === p),
    [g, b] = ne.useState("Theme"),
    [y, w] = ne.useState(cloneStructuredValue(d.theme)),
    [E, A] = ne.useState(cloneStructuredValue(v?.definition ?? { base: {} })),
    [k, F] = ne.useState("Custom style"),
    [I, O] = ne.useState(v?.name ?? ""),
    [j, z] = ne.useState(),
    [V, D] = ne.useState(""),
    [B, _] = ne.useState(0),
    C =
      s.snapshot !== JSON.stringify(e.workbook) ||
      s.sheet !== e.sheetId ||
      s.ranges !== JSON.stringify(e.ranges ?? [e.range]) ||
      s.active !== e.active,
    T = e.readOnly || C || !c,
    N = T || !!c?.protection?.enabled || countSelectionCells(u) > 500;
  function R() {
    (z(undefined), D(""), _(0));
  }
  function L(de) {
    (w(de), R());
  }
  function $(de) {
    (A(de), R());
  }
  function U(de) {
    (D(""), _(0));
    try {
      if (C)
        throw Error(
          "The workbook or selection changed. Close and reopen Themes & Cell Styles.",
        );
      z(applyNamedStyleChange(e.workbook, de, { readOnly: e.readOnly }));
    } catch (he) {
      (z(undefined),
        D(
          he instanceof Error
            ? he.message
            : "Unable to preview this style change.",
        ));
    }
  }
  function G() {
    try {
      if (!j) return;
      if (C)
        throw Error(
          "The workbook or selection changed. Close and reopen Themes & Cell Styles.",
        );
      const de = assertThemePreviewCurrent(e.workbook, j, {
        readOnly: e.readOnly,
      });
      (de !== e.workbook && e.onWorkbookChange(de, j.summary), e.onClose());
    } catch (de) {
      D(
        de instanceof Error ? de.message : "Unable to apply this style change.",
      );
    }
  }
  function H(de) {
    m(de);
    const he = d.styles.find((Se) => Se.id === de);
    (A(cloneStructuredValue(he?.definition ?? { base: {} })),
      O(he?.name ?? ""),
      R());
  }
  ne.useEffect(() => {
    const de = document.activeElement;
    return (
      o.current?.querySelector('[role="tab"][aria-selected="true"]')?.focus(),
      () => {
        de?.isConnected && de.focus();
      }
    );
  }, []);
  const Q = collectStyleFallbackWarnings(e.workbook),
    X$1 = Math.max(1, Math.ceil((j?.count ?? 0) / 20));
  let J = {};
  try {
    J = resolveNamedStyleAppearance(E, y);
  } catch {}
  return i.jsx("div", {
    className: "ts-backdrop",
    children: i.jsxs("div", {
      ref: o,
      className: "ts-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (de) => {
        if (
          (de.stopPropagation(),
          de.key === "Escape" && (de.preventDefault(), a.current()),
          de.key === "Tab")
        ) {
          const Se = Array.from(
              o.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ).filter((ge) => !ge.hidden),
            le = Se[0],
            ce = Se.at(-1);
          de.shiftKey && document.activeElement === le
            ? (de.preventDefault(), ce == null || ce.focus())
            : !de.shiftKey &&
              document.activeElement === ce &&
              (de.preventDefault(), le == null || le.focus());
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
            u.join(", "),
            " ·",
            " ",
            countSelectionCells(u).toLocaleString(),
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
                onKeyDown: (Se) => {
                  if (
                    ["ArrowLeft", "ArrowRight", "Home", "End"].includes(Se.key)
                  ) {
                    Se.preventDefault();
                    const ge =
                      Se.key === "Home" ? 0 : Se.key === "End" ? 1 : 1 - he;
                    (b(ge === 0 ? "Theme" : "Cell Styles"),
                      Se.currentTarget.parentElement?.children[ge]?.focus());
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
            (e.readOnly || C || Q.length > 0) &&
              i.jsx("div", {
                role: "status",
                className: "ts-notice",
                children: e.readOnly
                  ? "This workbook is read-only."
                  : C
                    ? "The workbook or selection changed. Close and reopen this dialog."
                    : Q.join(" "),
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
                                    (de) =>
                                      JSON.stringify(de) === JSON.stringify(y),
                                  )
                                    ? y.name
                                    : "Custom",
                                  disabled: T,
                                  onChange: (de) => {
                                    const he = builtInWorkbookThemes.find(
                                      (Se) => Se.name === de.target.value,
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
                                    builtInWorkbookThemes.map((de) =>
                                      i.jsx(
                                        "option",
                                        { children: de.name },
                                        de.name,
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
                                  value: y.name,
                                  maxLength: 80,
                                  disabled: T,
                                  onChange: (de) =>
                                    L({ ...y, name: de.target.value }),
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
                                      value: y.fonts[de],
                                      disabled: T,
                                      onChange: (he) =>
                                        L({
                                          ...y,
                                          fonts: {
                                            ...y.fonts,
                                            [de]: he.target.value,
                                          },
                                        }),
                                      children: supportedFontFamilies.map(
                                        (he) =>
                                          i.jsx("option", { children: he }, he),
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
                          disabled: T,
                          children: [
                            i.jsx("legend", {
                              children: i.jsx(LocalizedText, {
                                children: "Theme colors",
                              }),
                            }),
                            themePaletteRoles.map((de) =>
                              i.jsxs(
                                "label",
                                {
                                  children: [
                                    i.jsx("span", {
                                      children: de
                                        .replace("accent", "Accent ")
                                        .replace(/^background$/, "Background")
                                        .replace(/^text$/, "Text"),
                                    }),
                                    i.jsx("input", {
                                      "aria-label": `Theme ${de} color`,
                                      type: "color",
                                      value: y.palette[de],
                                      onChange: (he) =>
                                        L({
                                          ...y,
                                          palette: {
                                            ...y.palette,
                                            [de]: he.target.value.toUpperCase(),
                                          },
                                        }),
                                    }),
                                  ],
                                },
                                de,
                              ),
                            ),
                          ],
                        }),
                        i.jsx("button", {
                          className: "ts-command",
                          disabled: T,
                          onClick: () => U({ kind: "setTheme", theme: y }),
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
                                  value: p,
                                  disabled: T || !d.styles.length,
                                  onChange: (de) => H(de.target.value),
                                  lang: language("Named style"),
                                  "data-ui-label": "Named style",
                                  children: [
                                    !d.styles.length &&
                                      i.jsx("option", {
                                        value: "",
                                        children: "No named styles",
                                      }),
                                    d.styles.map((de) =>
                                      i.jsx(
                                        "option",
                                        { value: de.id, children: de.name },
                                        de.id,
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
                                  ranges: u,
                                  styleId: p,
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
                                  ranges: u,
                                  styleId: p,
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
                                  ranges: u,
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
                          disabled: T,
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
                                      onChange: (de) => {
                                        (F(de.target.value), R());
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
                                      value: E.fontRole ?? "fixed",
                                      onChange: (de) =>
                                        $({
                                          ...E,
                                          fontRole:
                                            de.target.value === "fixed"
                                              ? undefined
                                              : de.target.value,
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
                                          value: E[de] ?? "fixed",
                                          onChange: (he) =>
                                            $({
                                              ...E,
                                              [de]:
                                                he.target.value === "fixed"
                                                  ? undefined
                                                  : he.target.value,
                                            }),
                                          children: [
                                            i.jsx("option", {
                                              value: "fixed",
                                              children: "Keep stored color",
                                            }),
                                            themePaletteRoles.map((he) =>
                                              i.jsx(
                                                "option",
                                                { value: he, children: he },
                                                he,
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
                                      checked: E.base.bold ?? false,
                                      onChange: (de) =>
                                        $({
                                          ...E,
                                          base: {
                                            ...E.base,
                                            bold: de.target.checked,
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
                                      checked: E.base.italic ?? false,
                                      onChange: (de) =>
                                        $({
                                          ...E,
                                          base: {
                                            ...E.base,
                                            italic: de.target.checked,
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
                                      definition: E,
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
                                      styleId: p,
                                      definition: E,
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
                          disabled: T || !v,
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
                                      onChange: (de) => {
                                        (O(de.target.value), R());
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
                                          styleId: p,
                                          name: I,
                                        }),
                                      children: i.jsx(LocalizedText, {
                                        children: "Preview rename",
                                      }),
                                    }),
                                    i.jsx("button", {
                                      onClick: () =>
                                        U({ kind: "deleteStyle", styleId: p }),
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
            j &&
              i.jsxs("section", {
                className: "ts-preview",
                "aria-label": text("Theme change preview"),
                lang: language("Theme change preview"),
                "data-ui-label": "Theme change preview",
                children: [
                  i.jsx("h3", { children: j.summary }),
                  i.jsxs("p", {
                    children: [
                      j.count,
                      " ",
                      i.jsx(LocalizedText, { children: "cell" }),
                      j.count === 1 ? "" : "s",
                      " ",
                      i.jsx(LocalizedText, { children: "will change" }),
                      JSON.stringify(j.catalogBefore) !==
                      JSON.stringify(j.catalogAfter)
                        ? "; workbook style settings will also change"
                        : "",
                      ".",
                    ],
                  }),
                  i.jsx(ThemeChangePreview, { plan: j }),
                  j.warnings.map((de) =>
                    i.jsx("p", { className: "ts-notice", children: de }, de),
                  ),
                  j.count > 0 &&
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
                                children: j.changes
                                  .slice(B * 20, B * 20 + 20)
                                  .map((de) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("th", {
                                            children: [
                                              e.workbook.sheets.find(
                                                (ce) => ce.id === de.sheetId,
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
                                                catalog: j.catalogBefore,
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
                                                catalog: j.catalogAfter,
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
              children: j
                ? "Review the preview before applying."
                : "Choose a change to preview.",
            }),
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "ts-primary",
              disabled: T || !j,
              onClick: G,
              children: i.jsx(LocalizedText, { children: "Apply changes" }),
            }),
          ],
        }),
      ],
    }),
  });
}
function CellValuePreview({ value, cell, dateSystem }) {
  const n = useWorkbookMessages(),
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
      formatLocalizedNumberDisplay(value, cell, dateSystem, n.entryLocale),
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
