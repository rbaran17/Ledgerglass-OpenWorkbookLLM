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
    [isOnline, setIsOnline] = ne.useState(
      () => typeof navigator > "u" || navigator.onLine !== false,
    );
  return (
    ne.useEffect(() => {
      const handleOnline = () => setIsOnline(true),
        handleOffline = () => setIsOnline(false);
      return (
        window.addEventListener("online", handleOnline),
        window.addEventListener("offline", handleOffline),
        () => {
          (window.removeEventListener("online", handleOnline),
            window.removeEventListener("offline", handleOffline));
        }
      );
    }, []),
    isOnline
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
function parseGoToReference(
  activeWorkbook,
  currentSheetId,
  goToReferenceText,
  options = {},
) {
  if (
    typeof goToReferenceText != "string" ||
    !goToReferenceText.trim() ||
    goToReferenceText.length > 8192
  )
    throw Error(
      "Enter an address, range, or defined range name (up to 8,192 characters).",
    );
  const allowedSheetIds = options.allowedSheetIds,
    isSheetAllowed = (candidateSheetId) =>
      !allowedSheetIds || allowedSheetIds.includes(candidateSheetId),
    getVisibleSheet = (requestedSheetId) => {
      const sheet = activeWorkbook.sheets.find(
        (sheet) => sheet.id === requestedSheetId && isSheetAllowed(sheet.id),
      );
      if (!sheet)
        throw Error(
          "The destination is outside the available worksheet scope.",
        );
      if (sheet.hidden || sheet.veryHidden)
        throw Error(
          "The destination worksheet is hidden. Make it visible before navigating.",
        );
      return sheet;
    };
  getVisibleSheet(currentSheetId);
  const targetAreas = [],
    splitReferenceItems = (listText) => {
      const referenceItems = [];
      let insideQuotes = false,
        itemStart = 0;
      for (let charIndex = 0; charIndex < listText.length; charIndex++)
        listText[charIndex] === "'"
          ? insideQuotes && listText[charIndex + 1] === "'"
            ? charIndex++
            : (insideQuotes = !insideQuotes)
          : listText[charIndex] === "," &&
            !insideQuotes &&
            (referenceItems.push(listText.slice(itemStart, charIndex).trim()),
            (itemStart = charIndex + 1));
      if (
        (referenceItems.push(listText.slice(itemStart).trim()),
        insideQuotes ||
          referenceItems.some((referenceItem) => !referenceItem) ||
          referenceItems.length > maxSelectionAreas)
      )
        throw Error(
          "Enter valid comma-separated addresses or defined range names.",
        );
      return referenceItems;
    },
    resolveReferenceText = (addressText, defaultSheetId, nameChain) => {
      if (nameChain.length > 32)
        throw Error(
          "The defined-name chain exceeds the 32-name navigation limit.",
        );
      for (const listItem of splitReferenceItems(addressText)) {
        const sheetPrefixMatch =
          /^(?:'((?:[^']|'')+)'|([^'!,:\[\]()]+))!(.+)$/.exec(listItem);
        let sheetId = defaultSheetId,
          addressPart = listItem;
        if (sheetPrefixMatch) {
          const sheetName =
              sheetPrefixMatch[1]?.replace(/''/g, "'") ??
              sheetPrefixMatch[2].trim(),
            sheet = activeWorkbook.sheets.find(
              (sheet) =>
                isSheetAllowed(sheet.id) &&
                sheet.name.toLowerCase() === sheetName.toLowerCase(),
            );
          if (!sheet)
            throw Error(
              "The destination is outside the available worksheet scope.",
            );
          ((sheetId = sheet.id), (addressPart = sheetPrefixMatch[3].trim()));
        }
        if (
          (getVisibleSheet(sheetId),
          /^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/.test(
            addressPart,
          ))
        ) {
          if (
            (targetAreas.push({
              sheetId: sheetId,
              range: parseSelectionAreas(addressPart)[0],
            }),
            targetAreas.length > maxSelectionAreas)
          )
            throw Error(
              `Select at most ${maxSelectionAreas} rectangular areas.`,
            );
          continue;
        }
        if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/.test(addressPart))
          throw Error(
            "The Name Box accepts cell addresses and static range names. Computed, external, or three-dimensional destinations cannot be selected.",
          );
        const matchedDefinedName =
          activeWorkbook.definedNames?.find(
            (definedName) =>
              definedName.scopeSheetId === sheetId &&
              definedName.name.toLowerCase() === addressPart.toLowerCase(),
          ) ??
          activeWorkbook.definedNames?.find(
            (definedName) =>
              !definedName.scopeSheetId &&
              definedName.name.toLowerCase() === addressPart.toLowerCase(),
          );
        if (!matchedDefinedName)
          throw Error(
            "The address or defined range name was not found in the current worksheet scope.",
          );
        const chainKey = `${matchedDefinedName.scopeSheetId ?? "global"}:${matchedDefinedName.name.toLowerCase()}`;
        if (nameChain.includes(chainKey))
          throw Error("The defined range name contains a circular reference.");
        resolveReferenceText(
          matchedDefinedName.refersTo.slice(1).trim(),
          matchedDefinedName.scopeSheetId ?? sheetId,
          [...nameChain, chainKey],
        );
      }
    };
  if (
    (resolveReferenceText(goToReferenceText.trim(), currentSheetId, []),
    targetAreas.some(
      (targetArea) => targetArea.sheetId !== targetAreas[0].sheetId,
    ))
  )
    throw Error("Select ranges on one worksheet at a time.");
  return {
    sheetId: targetAreas[0].sheetId,
    ranges: parseSelectionAreas(
      targetAreas.map((targetEntry) => targetEntry.range),
    ),
  };
}

function describeStyleFormatting(cellStyle) {
  const effectiveCellStyle = resolveEffectiveCellStyle(cellStyle);
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
      Object.entries(styleFormattingLabels).map(
        ([propertyName, propertyLabel]) =>
          i.jsxs(
            "div",
            {
              children: [
                i.jsx("dt", { children: propertyLabel }),
                i.jsx("dd", {
                  children:
                    effectiveCellStyle[propertyName] === undefined
                      ? "Default"
                      : typeof effectiveCellStyle[propertyName] == "boolean"
                        ? effectiveCellStyle[propertyName]
                          ? "On"
                          : "Off"
                        : String(effectiveCellStyle[propertyName]) +
                          (propertyName === "fontSize" ? " pt" : ""),
                }),
              ],
            },
            propertyName,
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
  const styleBinding = style?.binding,
    { binding, ...directOverrides } = style ?? {};
  return i.jsxs("details", {
    children: [
      i.jsx("summary", {
        children: i.jsx(LocalizedText, { children: "Full formatting" }),
      }),
      i.jsx("h4", {
        children: i.jsx(LocalizedText, { children: "Displayed formatting" }),
      }),
      i.jsx(CellStylePreview, { style }),
      styleBinding &&
        i.jsxs(i.Fragment, {
          children: [
            i.jsxs("h4", {
              children: [
                i.jsx(LocalizedText, { children: "Named style:" }),
                " ",
                catalog?.styles.find(
                  (style) => style.id === styleBinding.styleId,
                )?.name ?? "Missing style",
              ],
            }),
            styleBinding.themeFallback &&
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children: "Office theme fallback",
                }),
              }),
            i.jsx(CellStylePreview, { style: styleBinding.base }),
            i.jsx("h4", {
              children: i.jsx(LocalizedText, { children: "Direct overrides" }),
            }),
            Object.keys(directOverrides).length
              ? i.jsx("dl", {
                  className: "ts-detail-fields",
                  children: Object.entries(directOverrides).map(
                    ([overrideProperty, overrideValue]) =>
                      i.jsxs(
                        "div",
                        {
                          children: [
                            i.jsx("dt", {
                              children:
                                styleFormattingLabels[overrideProperty] ??
                                overrideProperty,
                            }),
                            i.jsx("dd", {
                              children:
                                overrideProperty === "borders"
                                  ? Object.entries(overrideValue)
                                      .map(
                                        ([borderEdgeName, borderEdgeValue]) =>
                                          `${borderEdgeName}: ${borderEdgeValue.style}${borderEdgeValue.color ? ` ${borderEdgeValue.color}` : ""}`,
                                      )
                                      .join("; ")
                                  : typeof overrideValue == "boolean"
                                    ? overrideValue
                                      ? "On"
                                      : "Off"
                                    : String(overrideValue),
                            }),
                          ],
                        },
                        overrideProperty,
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
              ["top", "right", "bottom", "left"].map((borderEdge) =>
                i.jsxs(
                  "div",
                  {
                    children: [
                      i.jsxs("dt", {
                        children: [
                          borderEdge,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "border color role",
                          }),
                        ],
                      }),
                      i.jsx("dd", {
                        children:
                          definition.borderRoles?.[borderEdge] ??
                          "Stored border color",
                      }),
                    ],
                  },
                  borderEdge,
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
  const themeSettingRows = [
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
    ].filter((changeRow) => changeRow.before !== changeRow.after),
    styleIds = [
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
      themeSettingRows.length > 0 &&
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
              children: themeSettingRows.map((settingRow) =>
                i.jsxs(
                  "tr",
                  {
                    children: [
                      i.jsx("th", { children: settingRow.name }),
                      i.jsx("td", { children: settingRow.before ?? "Not set" }),
                      i.jsx("td", { children: settingRow.after ?? "Not set" }),
                    ],
                  },
                  settingRow.name,
                ),
              ),
            }),
          ],
        }),
      styleIds.map((styleId) => {
        const styleBefore = catalogBefore?.styles.find(
            (style) => style.id === styleId,
          ),
          styleAfter = catalogAfter?.styles.find(
            (style) => style.id === styleId,
          );
        return JSON.stringify(styleBefore) === JSON.stringify(styleAfter)
          ? null
          : i.jsxs(
              "details",
              {
                className: "ts-style-definition",
                children: [
                  i.jsx("summary", {
                    children: styleBefore
                      ? styleAfter
                        ? `${styleBefore.name} → ${styleAfter.name}`
                        : `Delete ${styleBefore.name}`
                      : `Create ${styleAfter?.name}`,
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
                            definition: styleBefore?.definition,
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
                            definition: styleAfter?.definition,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              },
              styleId,
            );
      }),
    ],
  });
}
function ThemesAndCellStylesDialog(props) {
  const { text, language } = useWorkbookMessages(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    onCloseRef = ne.useRef(props.onClose);
  onCloseRef.current = props.onClose;
  const [openSnapshot] = ne.useState(() => ({
      snapshot: JSON.stringify(props.workbook),
      sheet: props.sheetId,
      ranges: JSON.stringify(props.ranges ?? [props.range]),
      active: props.active,
    })),
    [styleCatalog] = ne.useState(() =>
      cloneStructuredValue(
        props.workbook.styleCatalog ??
          createDefaultStyleCatalog(props.workbook.id),
      ),
    ),
    selectionAreas = parseSelectionAreas(props.ranges ?? [props.range]),
    rangeAddress = parseRangeAddress(props.range),
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    activeAddress =
      props.active && isCellAddressInRanges(selectionAreas, props.active)
        ? props.active
        : `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    activeStyleId = sheet?.cells[activeAddress]?.style?.binding?.styleId,
    [styleId, setStyleId] = ne.useState(
      styleCatalog.styles.some((style) => style.id === activeStyleId)
        ? activeStyleId
        : (styleCatalog.styles[0]?.id ?? ""),
    ),
    style = styleCatalog.styles.find((style) => style.id === styleId),
    [activeTab, setActiveTab] = ne.useState("Theme"),
    [theme, setTheme] = ne.useState(cloneStructuredValue(styleCatalog.theme)),
    [definition, setDefinition] = ne.useState(
      cloneStructuredValue(style?.definition ?? { base: {} }),
    ),
    [newStyleName, setNewStyleName] = ne.useState("Custom style"),
    [renamedStyleName, setRenamedStyleName] = ne.useState(style?.name ?? ""),
    [plan, setPlan] = ne.useState(),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [pageIndex, setPageIndex] = ne.useState(0),
    isStale =
      openSnapshot.snapshot !== JSON.stringify(props.workbook) ||
      openSnapshot.sheet !== props.sheetId ||
      openSnapshot.ranges !== JSON.stringify(props.ranges ?? [props.range]) ||
      openSnapshot.active !== props.active,
    disabled = props.readOnly || isStale || !sheet,
    isStyleChangeBlocked =
      disabled ||
      !!sheet?.protection?.enabled ||
      countSelectionCells(selectionAreas) > 500;
  function resetPreview() {
    (setPlan(undefined), setErrorMessage(""), setPageIndex(0));
  }
  function updateTheme(nextTheme) {
    (setTheme(nextTheme), resetPreview());
  }
  function updateDefinition(nextDefinition) {
    (setDefinition(nextDefinition), resetPreview());
  }
  function previewStyleChange(request) {
    (setErrorMessage(""), setPageIndex(0));
    try {
      if (isStale)
        throw Error(
          "The workbook or selection changed. Close and reopen Themes & Cell Styles.",
        );
      setPlan(
        applyNamedStyleChange(props.workbook, request, {
          readOnly: props.readOnly,
        }),
      );
    } catch (error) {
      (setPlan(undefined),
        setErrorMessage(
          error instanceof Error
            ? error.message
            : "Unable to preview this style change.",
        ));
    }
  }
  function onClick() {
    try {
      if (!plan) return;
      if (isStale)
        throw Error(
          "The workbook or selection changed. Close and reopen Themes & Cell Styles.",
        );
      const nextWorkbook = assertThemePreviewCurrent(props.workbook, plan, {
        readOnly: props.readOnly,
      });
      (nextWorkbook !== props.workbook &&
        props.onWorkbookChange(nextWorkbook, plan.summary),
        props.onClose());
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : "Unable to apply this style change.",
      );
    }
  }
  function selectStyle(nextStyleId) {
    setStyleId(nextStyleId);
    const style = styleCatalog.styles.find((style) => style.id === nextStyleId);
    (setDefinition(cloneStructuredValue(style?.definition ?? { base: {} })),
      setRenamedStyleName(style?.name ?? ""),
      resetPreview());
  }
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current
        ?.querySelector('[role="tab"][aria-selected="true"]')
        ?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const styleFallbackWarnings = collectStyleFallbackWarnings(props.workbook),
    X$1 = Math.max(1, Math.ceil((plan?.count ?? 0) / 20));
  let previewAppearance = {};
  try {
    previewAppearance = resolveNamedStyleAppearance(definition, theme);
  } catch {}
  return i.jsx("div", {
    className: "ts-backdrop",
    children: i.jsxs("div", {
      ref: dialogRef,
      className: "ts-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(), onCloseRef.current()),
          event.key === "Tab")
        ) {
          const focusableElements = Array.from(
              dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ).filter((focusableElement) => !focusableElement.hidden),
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1);
          event.shiftKey && document.activeElement === firstFocusable
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === lastFocusable &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "ts-header",
          children: [
            i.jsx("h2", {
              id: `${dialogId}-title`,
              children: i.jsx(LocalizedText, {
                children: "Themes & Cell Styles",
              }),
            }),
            i.jsx("button", {
              "aria-label": text("Close Themes and Cell Styles"),
              onClick: props.onClose,
              lang: language("Close Themes and Cell Styles"),
              "data-ui-label": "Close Themes and Cell Styles",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "ts-location",
          children: [
            sheet?.name ?? "Missing sheet",
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
          children: ["Theme", "Cell Styles"].map((tabLabel, tabIndex) =>
            i.jsx(
              "button",
              {
                role: "tab",
                "aria-selected": activeTab === tabLabel,
                "aria-controls": `${dialogId}-panel`,
                tabIndex: activeTab === tabLabel ? 0 : -1,
                onClick: () => setActiveTab(tabLabel),
                onKeyDown: (event) => {
                  if (
                    ["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                      event.key,
                    )
                  ) {
                    event.preventDefault();
                    const targetTabIndex =
                      event.key === "Home"
                        ? 0
                        : event.key === "End"
                          ? 1
                          : 1 - tabIndex;
                    (setActiveTab(
                      targetTabIndex === 0 ? "Theme" : "Cell Styles",
                    ),
                      event.currentTarget.parentElement?.children[
                        targetTabIndex
                      ]?.focus());
                  }
                },
                children: tabLabel,
              },
              tabLabel,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "ts-body",
          children: [
            (props.readOnly || isStale || styleFallbackWarnings.length > 0) &&
              i.jsx("div", {
                role: "status",
                className: "ts-notice",
                children: props.readOnly
                  ? "This workbook is read-only."
                  : isStale
                    ? "The workbook or selection changed. Close and reopen this dialog."
                    : styleFallbackWarnings.join(" "),
              }),
            i.jsx("section", {
              id: `${dialogId}-panel`,
              role: "tabpanel",
              "aria-label": activeTab,
              children:
                activeTab === "Theme"
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
                                    const builtInWorkbookTheme =
                                      builtInWorkbookThemes.find(
                                        (builtInWorkbookTheme) =>
                                          builtInWorkbookTheme.name ===
                                          event.target.value,
                                      );
                                    builtInWorkbookTheme &&
                                      updateTheme(
                                        cloneStructuredValue(
                                          builtInWorkbookTheme,
                                        ),
                                      );
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
                                    updateTheme({
                                      ...theme,
                                      name: event.target.value,
                                    }),
                                }),
                              ],
                            }),
                            ["body", "heading"].map((fontRole) =>
                              i.jsxs(
                                "label",
                                {
                                  children: [
                                    fontRole === "body"
                                      ? "Body font"
                                      : "Heading font",
                                    i.jsx("select", {
                                      "aria-label":
                                        fontRole === "body"
                                          ? "Body font"
                                          : "Heading font",
                                      value: theme.fonts[fontRole],
                                      disabled: disabled,
                                      onChange: (event) =>
                                        updateTheme({
                                          ...theme,
                                          fonts: {
                                            ...theme.fonts,
                                            [fontRole]: event.target.value,
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
                                fontRole,
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
                                        updateTheme({
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
                          onClick: () =>
                            previewStyleChange({
                              kind: "setTheme",
                              theme: theme,
                            }),
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
                                  disabled:
                                    disabled || !styleCatalog.styles.length,
                                  onChange: (event) =>
                                    selectStyle(event.target.value),
                                  lang: language("Named style"),
                                  "data-ui-label": "Named style",
                                  children: [
                                    !styleCatalog.styles.length &&
                                      i.jsx("option", {
                                        value: "",
                                        children: "No named styles",
                                      }),
                                    styleCatalog.styles.map((style) =>
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
                              style: cellStyleToCss(previewAppearance),
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
                              disabled: isStyleChangeBlocked || !style,
                              onClick: () =>
                                previewStyleChange({
                                  kind: "applyStyle",
                                  sheetId: props.sheetId,
                                  ranges: selectionAreas,
                                  styleId: styleId,
                                }),
                              children: i.jsx(LocalizedText, {
                                children: "Apply style to selection",
                              }),
                            }),
                            i.jsx("button", {
                              disabled: isStyleChangeBlocked || !style,
                              onClick: () =>
                                previewStyleChange({
                                  kind: "reapplyStyle",
                                  sheetId: props.sheetId,
                                  ranges: selectionAreas,
                                  styleId: styleId,
                                }),
                              children: i.jsx(LocalizedText, {
                                children: "Reapply and clear direct styles",
                              }),
                            }),
                            i.jsx("button", {
                              disabled: isStyleChangeBlocked,
                              onClick: () =>
                                previewStyleChange({
                                  kind: "detachStyle",
                                  sheetId: props.sheetId,
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
                        sheet?.protection?.enabled &&
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
                                      value: newStyleName,
                                      maxLength: 80,
                                      onChange: (event) => {
                                        (setNewStyleName(event.target.value),
                                          resetPreview());
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
                                        updateDefinition({
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
                                ["fillRole", "colorRole"].map((roleField) =>
                                  i.jsxs(
                                    "label",
                                    {
                                      children: [
                                        roleField === "fillRole"
                                          ? "Style fill"
                                          : "Style text color",
                                        i.jsxs("select", {
                                          "aria-label":
                                            roleField === "fillRole"
                                              ? "Style fill"
                                              : "Style text color",
                                          value:
                                            definition[roleField] ?? "fixed",
                                          onChange: (event) =>
                                            updateDefinition({
                                              ...definition,
                                              [roleField]:
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
                                    roleField,
                                  ),
                                ),
                                i.jsxs("label", {
                                  className: "ts-check",
                                  children: [
                                    i.jsx("input", {
                                      type: "checkbox",
                                      checked: definition.base.bold ?? false,
                                      onChange: (event) =>
                                        updateDefinition({
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
                                        updateDefinition({
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
                                    updateDefinition({
                                      base: cloneStructuredValue(
                                        resolveEffectiveCellStyle(
                                          sheet?.cells[activeAddress]?.style,
                                        ),
                                      ),
                                    }),
                                  children: i.jsx(LocalizedText, {
                                    children: "Use active cell formatting",
                                  }),
                                }),
                                i.jsx("button", {
                                  onClick: () =>
                                    previewStyleChange({
                                      kind: "createStyle",
                                      name: newStyleName,
                                      definition: definition,
                                    }),
                                  children: i.jsx(LocalizedText, {
                                    children: "Preview new style",
                                  }),
                                }),
                                i.jsx("button", {
                                  disabled: !style,
                                  onClick: () =>
                                    previewStyleChange({
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
                          disabled: disabled || !style,
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
                                      value: renamedStyleName,
                                      maxLength: 80,
                                      onChange: (event) => {
                                        (setRenamedStyleName(
                                          event.target.value,
                                        ),
                                          resetPreview());
                                      },
                                    }),
                                  ],
                                }),
                                i.jsxs("div", {
                                  className: "ts-actions",
                                  children: [
                                    i.jsx("button", {
                                      onClick: () =>
                                        previewStyleChange({
                                          kind: "renameStyle",
                                          styleId: styleId,
                                          name: renamedStyleName,
                                        }),
                                      children: i.jsx(LocalizedText, {
                                        children: "Preview rename",
                                      }),
                                    }),
                                    i.jsx("button", {
                                      onClick: () =>
                                        previewStyleChange({
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
            errorMessage &&
              i.jsx("div", {
                role: "alert",
                className: "ts-error",
                children: errorMessage,
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
                                  .slice(pageIndex * 20, pageIndex * 20 + 20)
                                  .map((planChange) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("th", {
                                            children: [
                                              props.workbook.sheets.find(
                                                (sheet) =>
                                                  sheet.id ===
                                                  planChange.sheetId,
                                              )?.name,
                                              "!",
                                              planChange.address,
                                            ],
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              describeStyleFormatting(
                                                planChange.before?.style,
                                              ),
                                              i.jsx(CellStyleSummary, {
                                                style: planChange.before?.style,
                                                catalog: plan.catalogBefore,
                                              }),
                                            ],
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              describeStyleFormatting(
                                                planChange.after.style,
                                              ),
                                              i.jsx(CellStyleSummary, {
                                                style: planChange.after.style,
                                                catalog: plan.catalogAfter,
                                              }),
                                            ],
                                          }),
                                        ],
                                      },
                                      `${planChange.sheetId}:${planChange.address}`,
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
                              disabled: pageIndex === 0,
                              onClick: () =>
                                setPageIndex((currentPage) => currentPage - 1),
                              children: i.jsx(LocalizedText, {
                                children: "Previous changes",
                              }),
                            }),
                            i.jsxs("span", {
                              children: [
                                i.jsx(LocalizedText, { children: "Page " }),
                                pageIndex + 1,
                                " ",
                                i.jsx(LocalizedText, { children: "of " }),
                                X$1,
                              ],
                            }),
                            i.jsx("button", {
                              disabled: pageIndex + 1 >= X$1,
                              onClick: () =>
                                setPageIndex(
                                  (currentPageIndex) => currentPageIndex + 1,
                                ),
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
              onClick: props.onClose,
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
    formattedValue = cell?.numberFormat
      ? formatCellValueWithNumberFormat(value, cell.numberFormat, dateSystem)
      : undefined,
    importWarning = cell ? getImportedNumberFormatWarning(cell) : undefined;
  if (formattedValue?.accounting) {
    const { symbol, amount, symbolPosition } = formattedValue.accounting;
    return i.jsxs("span", {
      title: importWarning || undefined,
      "aria-label": formattedValue.text,
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
    title: importWarning || undefined,
    children:
      formattedValue?.text ??
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
      (e, cellIndex) =>
        `${columnIndexToLetters(cellIndex % 3)}${Math.floor(cellIndex / 3) + 1}`,
    ),
  ),
  expandShortHexColor = (hexColor, fallbackColor) =>
    hexColor?.length === 4
      ? "#" +
        hexColor
          .slice(1)
          .split("")
          .map((hexDigit) => hexDigit + hexDigit)
          .join("")
      : (hexColor ?? fallbackColor);

export {
  expandShortHexColor,
  parseGoToReference,
  ConnectionStatusNotice,
  ThemesAndCellStylesDialog,
  sampleGridAddresses,
  CellValuePreview,
  cellFormatTabs,
};
