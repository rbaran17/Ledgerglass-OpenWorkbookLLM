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

function FormatCellsDialog(props) {
  const { text, language } = useWorkbookMessages(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    onCloseRef = ne.useRef(props.onClose);
  onCloseRef.current = props.onClose;
  const sheet = props.workbook.sheets.find(
      (sheet) => sheet.id === props.sheetId,
    ),
    selectionAreas = parseSelectionAreas(props.ranges ?? [props.range]),
    rangeAddress = parseRangeAddress(props.range),
    activeAddress =
      props.active && isCellAddressInRanges(selectionAreas, props.active)
        ? props.active
        : `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    snapshotRef = ne.useRef({
      book: props.workbook.id,
      revision: props.workbook.revision,
      content: JSON.stringify(props.workbook),
      sheet: props.sheetId,
      selection: JSON.stringify(props.ranges ?? [props.range]),
      active: props.active,
    }),
    [baseStyle] = ne.useState(() =>
      validateCellStyle(
        resolveEffectiveCellStyle(sheet?.cells[activeAddress]?.style),
      ),
    ),
    [activeTab, setActiveTab] = ne.useState(
      props.mode === "alignment"
        ? "Alignment"
        : props.mode === "borders"
          ? "Border"
          : "Font",
    ),
    [pendingStyle, setPendingStyle] = ne.useState({}),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [preset, setPreset] = ne.useState(),
    [line, setLine] = ne.useState({ style: "thin", color: "#000000" }),
    [fontSizeText, setFontSizeText] = ne.useState(
      baseStyle.fontSize === undefined ? "" : String(baseStyle.fontSize),
    ),
    [indentText, setIndentText] = ne.useState(String(baseStyle.indent ?? 0)),
    selectedCellCount = countSelectionCells(selectionAreas),
    getCalculatedCell = ne.useMemo(
      () => calculateWorkbookWithSettings(props.workbook),
      [props.workbook],
    ),
    activeCellValue = sheet ? getCalculatedCell(sheet.id, activeAddress) : null,
    isSnapshotStale =
      snapshotRef.current.book !== props.workbook.id ||
      snapshotRef.current.revision !== props.workbook.revision ||
      snapshotRef.current.content !== JSON.stringify(props.workbook) ||
      snapshotRef.current.sheet !== props.sheetId ||
      snapshotRef.current.selection !==
        JSON.stringify(props.ranges ?? [props.range]) ||
      snapshotRef.current.active !== props.active,
    isEditingDisabled =
      props.readOnly ||
      isSnapshotStale ||
      !!sheet?.protection?.enabled ||
      !sheet ||
      selectedCellCount > 500;
  let draftStyle = baseStyle,
    styleValidationError = "";
  try {
    draftStyle = validateCellStyle({
      ...baseStyle,
      ...pendingStyle,
      ...(pendingStyle.borders
        ? { borders: { ...baseStyle.borders, ...pendingStyle.borders } }
        : {}),
    });
  } catch (error) {
    styleValidationError =
      error instanceof Error ? error.message : "Choose supported cell styles.";
  }
  function updatePendingStyle(patch) {
    (setPendingStyle((previousPendingStyle) => ({
      ...previousPendingStyle,
      ...patch,
    })),
      setErrorMessage(""));
  }
  function updateBorderEdge(edge, shouldClearEdge = false) {
    (setPendingStyle((previousPendingState) => ({
      ...previousPendingState,
      borders: {
        ...previousPendingState.borders,
        [edge]: shouldClearEdge ? { style: "none" } : { ...line },
      },
    })),
      setErrorMessage(""));
  }
  function selectPreset(selectedPreset) {
    (setPreset(selectedPreset),
      setPendingStyle((previousPendingDraft) => {
        const pendingWithoutBorders = { ...previousPendingDraft };
        return (delete pendingWithoutBorders.borders, pendingWithoutBorders);
      }),
      setErrorMessage(""));
  }
  const getPreviewStyle = (address) => ({
    ...draftStyle,
    borders: {
      ...baseStyle.borders,
      ...(preset
        ? buildBorderPresetStyles(address, sampleGridAddresses, preset, line)
        : {}),
      ...pendingStyle.borders,
    },
  });
  return (
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
    }, []),
    i.jsx("div", {
      className: "cs-backdrop",
      children: i.jsxs("div", {
        ref: dialogRef,
        className: "cs-dialog",
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
            const focusableElements = [
                ...(dialogRef.current?.querySelectorAll(
                  'button:not(:disabled):not([tabindex="-1"]),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
                ) ?? []),
              ],
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
            className: "cs-header",
            children: [
              i.jsx("h2", {
                id: `${dialogId}-title`,
                children: i.jsx(LocalizedText, { children: "Format Cells" }),
              }),
              i.jsx("button", {
                "aria-label": text("Close cell styles"),
                onClick: props.onClose,
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
            children: cellFormatTabs.map((cellFormatTab, tabIndex) =>
              i.jsx(
                "button",
                {
                  role: "tab",
                  id: `${dialogId}-${cellFormatTab}`,
                  "aria-controls": `${dialogId}-panel`,
                  "aria-selected": activeTab === cellFormatTab,
                  tabIndex: activeTab === cellFormatTab ? 0 : -1,
                  onClick: () => setActiveTab(cellFormatTab),
                  onKeyDown: (event) => {
                    if (
                      ["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                        event.key,
                      )
                    ) {
                      event.preventDefault();
                      const nextTabIndex =
                        event.key === "Home"
                          ? 0
                          : event.key === "End"
                            ? cellFormatTabs.length - 1
                            : (tabIndex +
                                (event.key === "ArrowRight"
                                  ? 1
                                  : cellFormatTabs.length - 1)) %
                              cellFormatTabs.length;
                      (setActiveTab(cellFormatTabs[nextTabIndex]),
                        event.currentTarget.parentElement?.children[
                          nextTabIndex
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
              (event.preventDefault(), setErrorMessage(""));
              try {
                if (props.readOnly) throw Error("This workbook is read-only.");
                if (isSnapshotStale)
                  throw Error(
                    "The workbook or selection changed. Close and reopen Format Cells.",
                  );
                if (styleValidationError) throw Error(styleValidationError);
                const nextWorkbook = applyBorderEditTransaction(
                  props.workbook,
                  props.sheetId,
                  selectionAreas,
                  pendingStyle,
                  {
                    readOnly: props.readOnly,
                    ...(preset
                      ? { borders: { preset: preset, line: line } }
                      : {}),
                  },
                );
                (nextWorkbook !== props.workbook &&
                  props.onWorkbookChange(nextWorkbook, "Format cell styles"),
                  props.onClose());
              } catch (error) {
                setErrorMessage(
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
                      selectedCellCount.toLocaleString(),
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
                  props.readOnly &&
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
                  isSnapshotStale &&
                    i.jsx("p", {
                      className: "cs-error",
                      role: "alert",
                      children: i.jsx(LocalizedText, {
                        children:
                          "The workbook or selection changed. Close and reopen Format Cells.",
                      }),
                    }),
                  selectedCellCount > 500 &&
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
                    id: `${dialogId}-panel`,
                    "aria-labelledby": `${dialogId}-${activeTab}`,
                    children: [
                      activeTab === "Font" &&
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
                                      value: draftStyle.fontFamily ?? "",
                                      onChange: (event) =>
                                        updatePendingStyle({
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
                                      value: fontSizeText,
                                      onValueChange: (fontSizeValue) => {
                                        (setFontSizeText(fontSizeValue),
                                          updatePendingStyle({
                                            fontSize:
                                              fontSizeValue === ""
                                                ? undefined
                                                : Number(fontSizeValue),
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
                                ].map(([effectKey, effectLabel]) =>
                                  i.jsxs(
                                    "label",
                                    {
                                      children: [
                                        i.jsx("input", {
                                          type: "checkbox",
                                          "aria-label": effectLabel,
                                          checked:
                                            draftStyle[effectKey] ?? false,
                                          onChange: (event) =>
                                            updatePendingStyle({
                                              [effectKey]: event.target.checked,
                                            }),
                                        }),
                                        effectLabel,
                                      ],
                                    },
                                    effectKey,
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
                                            draftStyle.color,
                                            "#000000",
                                          ),
                                          onChange: (event) =>
                                            updatePendingStyle({
                                              color: event.target.value,
                                            }),
                                          lang: language("Font color"),
                                          "data-ui-label": "Font color",
                                        }),
                                        i.jsx("button", {
                                          type: "button",
                                          onClick: () =>
                                            updatePendingStyle({
                                              color: undefined,
                                            }),
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
                                            draftStyle.fill,
                                            "#ffffff",
                                          ),
                                          onChange: (event) =>
                                            updatePendingStyle({
                                              fill: event.target.value,
                                            }),
                                          lang: language("Fill color"),
                                          "data-ui-label": "Fill color",
                                        }),
                                        i.jsx("button", {
                                          type: "button",
                                          onClick: () =>
                                            updatePendingStyle({
                                              fill: undefined,
                                            }),
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
                      activeTab === "Alignment" &&
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
                                      value: draftStyle.align ?? "",
                                      onChange: (event) =>
                                        updatePendingStyle({
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
                                      value: draftStyle.verticalAlign ?? "",
                                      onChange: (event) =>
                                        updatePendingStyle({
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
                                      value: indentText,
                                      onValueChange: (indentValue) => {
                                        (setIndentText(indentValue),
                                          updatePendingStyle({
                                            indent:
                                              indentValue === ""
                                                ? NaN
                                                : Number(indentValue),
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
                                  checked: draftStyle.wrap ?? false,
                                  onChange: (event) =>
                                    updatePendingStyle({
                                      wrap: event.target.checked,
                                    }),
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
                      activeTab === "Border" &&
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
                                        setLine({
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
                                        setLine({
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
                                    (setPreset(undefined),
                                      setPendingStyle((previousStyleDraft) => {
                                        const clearedStyle = {
                                          ...previousStyleDraft,
                                        };
                                        return (
                                          delete clearedStyle.borders,
                                          clearedStyle
                                        );
                                      }),
                                      setErrorMessage(""));
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
                                ].map(([presetKey, presetLabel]) =>
                                  i.jsx(
                                    "button",
                                    {
                                      type: "button",
                                      "aria-pressed": preset === presetKey,
                                      onClick: () => selectPreset(presetKey),
                                      children: presetLabel,
                                    },
                                    presetKey,
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
                                          onClick: () =>
                                            updateBorderEdge(cellBorderEdge),
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
                                            updateBorderEdge(
                                              cellBorderEdge,
                                              true,
                                            ),
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
                          activeAddress,
                        ],
                      }),
                      activeTab === "Border"
                        ? i.jsx("table", {
                            className: "cs-border-preview",
                            "aria-label": text("Border preset preview"),
                            lang: language("Border preset preview"),
                            "data-ui-label": "Border preset preview",
                            children: i.jsx("tbody", {
                              children: [0, 1, 2].map((rowIndex) =>
                                i.jsx(
                                  "tr",
                                  {
                                    children: [0, 1, 2].map((columnIndex) => {
                                      const cellAddress = `${columnIndexToLetters(columnIndex)}${rowIndex + 1}`,
                                        cellPreviewStyle =
                                          getPreviewStyle(cellAddress);
                                      return i.jsx(
                                        "td",
                                        {
                                          "data-preview-cell": cellAddress,
                                          style: cellStyleToCss(
                                            cellPreviewStyle,
                                            "",
                                            {
                                              defaultBorder:
                                                "1px dotted #d5d5d5",
                                            },
                                          ),
                                          children:
                                            rowIndex === 1 && columnIndex === 1
                                              ? "Text"
                                              : "",
                                        },
                                        columnIndex,
                                      );
                                    }),
                                  },
                                  rowIndex,
                                ),
                              ),
                            }),
                          })
                        : i.jsx("table", {
                            className: "cs-text-preview",
                            children: i.jsx("tbody", {
                              children: i.jsx("tr", {
                                children: i.jsx("td", {
                                  style: cellStyleToCss(
                                    getPreviewStyle("B2"),
                                    activeCellValue,
                                    {
                                      defaultBorder: "1px solid #d5d5d5",
                                    },
                                  ),
                                  children: i.jsx("div", {
                                    style: {
                                      ...textWrapToCss(draftStyle),
                                      maxHeight: 94,
                                      overflow: "hidden",
                                      lineHeight: 1.35,
                                    },
                                    children: i.jsx(CellValuePreview, {
                                      dateSystem:
                                        props.workbook.datePolicy?.dateSystem,
                                      value:
                                        activeCellValue === null
                                          ? "Audit evidence"
                                          : activeCellValue,
                                      cell: {
                                        ...(sheet?.cells[activeAddress] ?? {
                                          raw: null,
                                        }),
                                        style: draftStyle,
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
                  (errorMessage || styleValidationError) &&
                    i.jsx("p", {
                      className: "cs-error",
                      role: "alert",
                      children: errorMessage || styleValidationError,
                    }),
                ],
              }),
              i.jsxs("footer", {
                className: "cs-actions",
                children: [
                  i.jsx("button", {
                    type: "button",
                    onClick: props.onClose,
                    children: i.jsx(LocalizedText, { children: "Cancel" }),
                  }),
                  i.jsx("button", {
                    className: "cs-primary",
                    disabled: isEditingDisabled || !!styleValidationError,
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
function describeDataBarRule(decoration) {
  return [
    decoration.style.fill ? `Fill ${decoration.style.fill}` : "No fill",
    decoration.style.color ? `Text ${decoration.style.color}` : "",
    decoration.dataBar
      ? `Bar ${decoration.dataBar.width.toFixed(1)}% at ${decoration.dataBar.start.toFixed(1)}%${decoration.dataBar.negative ? " negative" : ""}`
      : "",
    decoration.icon ? `${decoration.icon.label} icon` : "",
    ...decoration.errors.map((error) => error.message),
  ]
    .filter(Boolean)
    .join("; ");
}
function ConditionalFormattingRulesManager(props) {
  const { text, language } = useWorkbookMessages(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    snapshotRef = ne.useRef({
      book: props.workbook,
      sheetId: props.sheetId,
      range: props.range,
      guard: JSON.stringify(props.workbook),
    }),
    [rules, setRules] = ne.useState(() =>
      structuredClone(
        props.workbook.sheets.find((sheet) => sheet.id === props.sheetId)
          ?.conditionalFormats ?? [],
      ),
    ),
    [ruleDraft, setRuleDraft] = ne.useState(() =>
      createConditionalRuleDraft(props.range),
    ),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [pendingPreview, setPendingPreview] = ne.useState(null),
    [previewPage, setPreviewPage] = ne.useState(0),
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    isSnapshotStale =
      snapshotRef.current.guard !== JSON.stringify(props.workbook) ||
      snapshotRef.current.sheetId !== props.sheetId ||
      snapshotRef.current.range !== props.range,
    disabled =
      props.readOnly ||
      isSnapshotStale ||
      !sheet ||
      !!sheet.protection?.enabled,
    sortedRules = sortConditionalRulesByPriority(rules),
    isVisualRuleKind = ["colorScale", "dataBar", "iconSet"].includes(
      ruleDraft.kind,
    );
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const commitRules = (nextRules) => {
      (setRules(
        nextRules.map((rule, ruleIndex) => ({
          ...rule,
          priority: ruleIndex + 1,
        })),
      ),
        setPendingPreview(null),
        setErrorMessage(""));
    },
    patchRuleDraft = (patch) => {
      (setRuleDraft({ ...ruleDraft, ...patch }),
        setPendingPreview(null),
        setErrorMessage(""));
    },
    stageRule = () => {
      try {
        if (disabled) throw Error("Rules cannot be changed in this view.");
        const stagedRule = {
            ...ruleDraft,
            id: ruleDraft.id || crypto.randomUUID(),
            ...(isVisualRuleKind ? { stopIfTrue: false } : {}),
          },
          nextRuleList = ruleDraft.id
            ? sortedRules.map((existingRule) =>
                existingRule.id === ruleDraft.id ? stagedRule : existingRule,
              )
            : [stagedRule, ...sortedRules],
          conditionalRules = normalizeConditionalRules(
            nextRuleList.map((listedRule, listedRuleIndex) => ({
              ...listedRule,
              priority: listedRuleIndex + 1,
            })),
          );
        (commitRules(conditionalRules),
          setRuleDraft(createConditionalRuleDraft(props.range)));
      } catch (error) {
        setErrorMessage(
          error instanceof Error ? error.message : "Rule could not be staged.",
        );
      }
    },
    onClick = () => {
      try {
        if (disabled) throw Error("Rules cannot be changed in this view.");
        (setPendingPreview(
          updateConditionalFormatRules(props.workbook, props.sheetId, rules, {
            readOnly: props.readOnly,
          }),
        ),
          setPreviewPage(0),
          setErrorMessage(""));
      } catch (error) {
        setErrorMessage(
          error instanceof Error ? error.message : "Preview failed.",
        );
      }
    };
  return i.jsx("div", {
    className: "cf-backdrop",
    children: i.jsxs("div", {
      ref: dialogRef,
      className: "cf-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
              ...(dialogRef.current?.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
              ) ?? []),
            ],
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
          children: [
            i.jsx("h2", {
              id: `${dialogId}-title`,
              children: i.jsx(LocalizedText, {
                children: "Conditional Formatting Rules Manager",
              }),
            }),
            i.jsx("button", {
              "aria-label": text("Close conditional formatting"),
              onClick: props.onClose,
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
        isSnapshotStale &&
          i.jsx("p", {
            role: "alert",
            children: i.jsx(LocalizedText, {
              children:
                "The workbook or selection changed. Close and reopen this manager to refresh.",
            }),
          }),
        props.readOnly &&
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
                sortedRules.length === 0
                  ? i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children: "No conditional rules.",
                      }),
                    })
                  : i.jsx("ol", {
                      className: "cf-rules",
                      children: sortedRules.map((sortedRule, sortedIndex) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              i.jsx("strong", {
                                children:
                                  conditionalRuleTypeLabels[sortedRule.kind],
                              }),
                              " · ",
                              sortedRule.range,
                              i.jsxs("span", {
                                children: [
                                  sortedRule.formula ?? sortedRule.value,
                                  sortedRule.kind === "between"
                                    ? ` to ${sortedRule.second}`
                                    : "",
                                  sortedRule.stopIfTrue
                                    ? " · Stop if true"
                                    : "",
                                ],
                              }),
                              i.jsxs("div", {
                                children: [
                                  i.jsxs("button", {
                                    disabled: disabled,
                                    onClick: () => {
                                      (setRuleDraft(
                                        structuredClone(sortedRule),
                                      ),
                                        setErrorMessage(""));
                                    },
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Edit rule ",
                                      }),
                                      sortedIndex + 1,
                                    ],
                                  }),
                                  i.jsxs("button", {
                                    disabled: disabled,
                                    onClick: () => {
                                      (commitRules(
                                        sortedRules.filter(
                                          (ruleToKeep) =>
                                            ruleToKeep.id !== sortedRule.id,
                                        ),
                                      ),
                                        ruleDraft.id === sortedRule.id &&
                                          setRuleDraft(
                                            createConditionalRuleDraft(
                                              props.range,
                                            ),
                                          ));
                                    },
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Delete rule ",
                                      }),
                                      sortedIndex + 1,
                                    ],
                                  }),
                                  i.jsx("button", {
                                    "aria-label": `Move rule ${sortedIndex + 1} up`,
                                    disabled: disabled || sortedIndex === 0,
                                    onClick: () => {
                                      const movedRulesUp = [...sortedRules];
                                      (([
                                        movedRulesUp[sortedIndex - 1],
                                        movedRulesUp[sortedIndex],
                                      ] = [
                                        movedRulesUp[sortedIndex],
                                        movedRulesUp[sortedIndex - 1],
                                      ]),
                                        commitRules(movedRulesUp));
                                    },
                                    children: "↑",
                                  }),
                                  i.jsx("button", {
                                    "aria-label": `Move rule ${sortedIndex + 1} down`,
                                    disabled:
                                      disabled ||
                                      sortedIndex === sortedRules.length - 1,
                                    onClick: () => {
                                      const movedRulesDown = [...sortedRules];
                                      (([
                                        movedRulesDown[sortedIndex + 1],
                                        movedRulesDown[sortedIndex],
                                      ] = [
                                        movedRulesDown[sortedIndex],
                                        movedRulesDown[sortedIndex + 1],
                                      ]),
                                        commitRules(movedRulesDown));
                                    },
                                    children: "↓",
                                  }),
                                ],
                              }),
                            ],
                          },
                          sortedRule.id,
                        ),
                      ),
                    }),
              ],
            }),
            i.jsxs("form", {
              className: "cf-editor",
              onSubmit: (event) => {
                (event.preventDefault(), stageRule());
              },
              children: [
                i.jsx("h3", {
                  children: ruleDraft.id ? "Edit rule" : "New rule",
                }),
                i.jsxs("fieldset", {
                  disabled: disabled,
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Rule type" }),
                        i.jsx("select", {
                          value: ruleDraft.kind,
                          onChange: (event) => {
                            const value = event.target.value;
                            patchRuleDraft({
                              kind: value,
                              ...(value === "formula"
                                ? { formula: ruleDraft.formula ?? "=A1>0" }
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
                          value: ruleDraft.range,
                          onChange: (event) =>
                            patchRuleDraft({ range: event.target.value }),
                        }),
                      ],
                    }),
                    ruleDraft.kind === "formula"
                      ? i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Condition formula",
                                }),
                                i.jsx("input", {
                                  value: ruleDraft.formula ?? "",
                                  onChange: (event) =>
                                    patchRuleDraft({
                                      formula: event.target.value,
                                    }),
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
                        ].includes(ruleDraft.kind) &&
                        i.jsxs("label", {
                          children: [
                            ruleDraft.kind === "contains"
                              ? "Text to find"
                              : "Rule value",
                            i.jsx("input", {
                              value: ruleDraft.value,
                              onChange: (event) =>
                                patchRuleDraft({ value: event.target.value }),
                            }),
                          ],
                        }),
                    ruleDraft.kind === "between" &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Second value" }),
                          i.jsx("input", {
                            value: ruleDraft.second,
                            onChange: (event) =>
                              patchRuleDraft({ second: event.target.value }),
                          }),
                        ],
                      }),
                    i.jsxs("div", {
                      className: "cf-colors",
                      children: [
                        i.jsxs("label", {
                          children: [
                            ruleDraft.kind === "dataBar"
                              ? "Positive bar color"
                              : "Fill color",
                            i.jsx("input", {
                              type: "color",
                              value: ruleDraft.fill,
                              onChange: (event) =>
                                patchRuleDraft({ fill: event.target.value }),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Text color" }),
                            i.jsx("input", {
                              type: "color",
                              value: ruleDraft.color,
                              onChange: (event) =>
                                patchRuleDraft({ color: event.target.value }),
                            }),
                          ],
                        }),
                      ],
                    }),
                    ruleDraft.kind === "colorScale" &&
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
                                value: ruleDraft.minColor ?? "#f8e99b",
                                onChange: (event) =>
                                  patchRuleDraft({
                                    minColor: event.target.value,
                                  }),
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
                                value: ruleDraft.maxColor ?? "#63be7b",
                                onChange: (event) =>
                                  patchRuleDraft({
                                    maxColor: event.target.value,
                                  }),
                              }),
                            ],
                          }),
                        ],
                      }),
                    ruleDraft.kind === "dataBar" &&
                      i.jsx("small", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Automatic numeric minimum and maximum, a zero axis, red negative bars. Values stay visible.",
                        }),
                      }),
                    ruleDraft.kind === "iconSet" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Icon threshold units",
                              }),
                              i.jsxs("select", {
                                value: ruleDraft.thresholdType ?? "percent",
                                onChange: (event) =>
                                  patchRuleDraft({
                                    thresholdType: event.target.value,
                                  }),
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
                                value: ruleDraft.iconThresholds?.[0] ?? 33,
                                onValueChange: (middleThreshold) =>
                                  patchRuleDraft({
                                    iconThresholds: [
                                      Number(middleThreshold),
                                      ruleDraft.iconThresholds?.[1] ?? 67,
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
                                value: ruleDraft.iconThresholds?.[1] ?? 67,
                                onValueChange: (highThreshold) =>
                                  patchRuleDraft({
                                    iconThresholds: [
                                      ruleDraft.iconThresholds?.[0] ?? 33,
                                      Number(highThreshold),
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
                          checked: ruleDraft.stopIfTrue ?? false,
                          disabled: isVisualRuleKind || disabled,
                          onChange: (event) =>
                            patchRuleDraft({
                              stopIfTrue: event.target.checked,
                            }),
                        }),
                        i.jsx(LocalizedText, { children: "Stop if true" }),
                      ],
                    }),
                    i.jsx("button", {
                      type: "submit",
                      children: ruleDraft.id
                        ? "Stage rule edit"
                        : "Stage new rule",
                    }),
                    i.jsx("button", {
                      type: "button",
                      onClick: () =>
                        setRuleDraft(createConditionalRuleDraft(props.range)),
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
        errorMessage && i.jsx("p", { role: "alert", children: errorMessage }),
        i.jsx("button", {
          disabled: disabled,
          onClick: onClick,
          children: i.jsx(LocalizedText, { children: "Preview rule changes" }),
        }),
        pendingPreview &&
          i.jsxs("section", {
            className: "cf-preview",
            "aria-label": text("Conditional formatting preview"),
            lang: language("Conditional formatting preview"),
            "data-ui-label": "Conditional formatting preview",
            children: [
              i.jsxs("h3", {
                children: [
                  pendingPreview.changes.length,
                  " ",
                  i.jsx(LocalizedText, { children: "rule changes" }),
                ],
              }),
              pendingPreview.changes.length === 0 &&
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
                  pendingPreview.previewLimit,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "anchors and populated cells; it is not an exhaustive scan of the range.",
                  }),
                ],
              }),
              pendingPreview.changes.map((change) =>
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
              pendingPreview.errors.length > 0 &&
                i.jsxs("div", {
                  role: "alert",
                  children: [
                    i.jsx("strong", {
                      children: i.jsx(LocalizedText, {
                        children: "Rule errors",
                      }),
                    }),
                    i.jsx("ul", {
                      children: pendingPreview.errors.map((error, errorIndex) =>
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
                          errorIndex,
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
                      children: pendingPreview.samples
                        .slice(previewPage * 10, previewPage * 10 + 10)
                        .map((sample) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("th", { children: sample.address }),
                                i.jsx("td", {
                                  children: String(sample.value ?? ""),
                                }),
                                i.jsx("td", {
                                  children: describeDataBarRule(sample.before),
                                }),
                                i.jsx("td", {
                                  children: describeDataBarRule(sample.after),
                                }),
                              ],
                            },
                            sample.address,
                          ),
                        ),
                    }),
                  ],
                }),
              }),
              pendingPreview.samples.length > 10 &&
                i.jsxs("nav", {
                  "aria-label": text("Conditional preview pages"),
                  lang: language("Conditional preview pages"),
                  "data-ui-label": "Conditional preview pages",
                  children: [
                    i.jsx("button", {
                      disabled: previewPage === 0,
                      onClick: () => setPreviewPage(previewPage - 1),
                      children: i.jsx(LocalizedText, {
                        children: "Previous samples",
                      }),
                    }),
                    i.jsxs("span", {
                      children: [
                        i.jsx(LocalizedText, { children: "Page " }),
                        previewPage + 1,
                        " ",
                        i.jsx(LocalizedText, { children: "of " }),
                        Math.ceil(pendingPreview.samples.length / 10),
                      ],
                    }),
                    i.jsx("button", {
                      disabled:
                        (previewPage + 1) * 10 >= pendingPreview.samples.length,
                      onClick: () => setPreviewPage(previewPage + 1),
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
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              disabled:
                disabled ||
                !pendingPreview ||
                pendingPreview.changes.length === 0,
              onClick: () => {
                try {
                  if (!pendingPreview || disabled) return;
                  const nextWorkbook = assertConditionalFormatPreviewCurrent(
                    props.workbook,
                    pendingPreview,
                    { readOnly: props.readOnly },
                  );
                  (nextWorkbook !== props.workbook &&
                    props.onWorkbookChange(
                      nextWorkbook,
                      "Update conditional formatting rules",
                    ),
                    props.onClose());
                } catch (error) {
                  setErrorMessage(
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
const rangeToA1Address = (rangeText) => {
  const rangeAddress = parseRangeAddress(rangeText);
  return `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
};
function previewMergeCells(
  book,
  sheetId,
  range,
  unmerge = false,
  options = {},
) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  const sheetOrThrow = findSheetOrThrow(book, sheetId);
  assertSheetUnprotected(sheetOrThrow);
  const rangeA1 = rangeToA1Address(range);
  if (countRangeCells(rangeA1) > 500)
    throw Error("Preview a merge or unmerge of at most 500 cells.");
  if (!unmerge && countRangeCells(rangeA1) < 2)
    throw Error("Select at least two cells to merge.");
  const overlappingMerges = (sheetOrThrow.merges ?? []).filter((mergedRange) =>
    rangesOverlap(mergedRange, rangeA1),
  );
  if (!unmerge && overlappingMerges.length)
    throw Error("Unmerge overlapping areas first.");
  if (unmerge && !overlappingMerges.length)
    throw Error("The selected range does not intersect a merged area.");
  const affectedRanges = unmerge ? overlappingMerges : [rangeA1];
  if (
    affectedRanges.reduce(
      (total, affectedRange) => total + countRangeCells(affectedRange),
      0,
    ) > 500
  )
    throw Error(
      "Unmerge fewer areas; this preview supports at most 500 affected cells.",
    );
  if (!unmerge) {
    if (
      (sheetOrThrow.tables ?? []).some((table) =>
        rangesOverlap(table.range, rangeA1),
      )
    )
      throw Error(
        "Merging cells inside a table is unsupported. Convert the table first.",
      );
    if (
      [...(sheetOrThrow.pivots ?? []), ...(sheetOrThrow.queries ?? [])].some(
        (analysisOutput) => rangesOverlap(analysisOutput.outputRange, rangeA1),
      )
    )
      throw Error("Merging generated analysis output is unsupported.");
  }
  normalizeWorkbookInput(mergeWorksheetCells(book, sheetId, rangeA1, unmerge));
  const workbookWithSettings = calculateWorkbookWithSettings(book);
  if (!unmerge) {
    let formulaCount = 0;
    for (const [cellAddress, cell] of Object.entries(sheetOrThrow.cells))
      if (
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal
      ) {
        if (++formulaCount > 1e4)
          throw Error(
            "Merge checking supports at most 10,000 active formulas per worksheet.",
          );
        const spillRange = workbookWithSettings.getSpillRange(
          sheetId,
          cellAddress,
        );
        if (spillRange && rangesOverlap(spillRange, rangeA1))
          throw Error(
            `Unmergeable array output at ${spillRange}. Move or remove its formula first.`,
          );
      }
  }
  const cells = [];
  for (const affectedArea of affectedRanges) {
    const rangeAddress = parseRangeAddress(affectedArea),
      anchorAddress = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`;
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
            (thread) =>
              !thread.orphaned &&
              thread.anchor.sheetId === sheetId &&
              thread.anchor.address === address,
          ).length;
        (cell || discussions || address === anchorAddress) &&
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
              address === anchorAddress
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
    range: rangeA1,
    unmerge: unmerge,
    affectedRanges: [...affectedRanges],
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
function assertMergePreviewCurrent(workbook, plan, options = {}) {
  if (plan.snapshot !== JSON.stringify(workbook))
    throw Error(
      "The workbook changed. Close this preview and review the current cells again.",
    );
  const currentPreview = previewMergeCells(
      workbook,
      plan.sheetId,
      plan.range,
      plan.unmerge,
      options,
    ),
    comparablePreview = (preview) => ({
      ...preview,
      cells: preview.cells.map(
        ({ display, ...cellWithoutDisplay }) => cellWithoutDisplay,
      ),
    });
  if (
    JSON.stringify(comparablePreview(currentPreview)) !==
    JSON.stringify(comparablePreview(plan))
  )
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
