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

function PasteSpecialDialog(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    propsRef = ne.useRef(props);
  propsRef.current = props;
  const openedSnapshotRef = ne.useRef({
      workbookId: props.workbook.id,
      revision: props.workbook.revision,
      sheetId: props.sheetId,
      range: props.range,
      clipboard: props.clipboard,
    }),
    sheetOrThrow = findSheetOrThrow(props.workbook, props.sheetId),
    rangeAddress = parseRangeAddress(props.range),
    anchorAddress = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    [actionError, setActionError] = ne.useState(""),
    [mode, setMode] = ne.useState("all"),
    [operation, setOperation] = ne.useState("none"),
    [transpose, setTranspose] = ne.useState(false),
    [skipBlanks, setSkipBlanks] = ne.useState(false),
    [fillMode, setFillMode] = ne.useState("series"),
    [fillDirection, setFillDirection] = ne.useState("down"),
    [seriesAxis, setSeriesAxis] = ne.useState(
      rangeAddress.firstRow === rangeAddress.lastRow ? "rows" : "columns",
    ),
    [isBackward, setIsBackward] = ne.useState(false),
    [seriesKind, setSeriesKind] = ne.useState("number"),
    [startText, setStartText] = ne.useState(""),
    [stepText, setStepText] = ne.useState("1"),
    [stopText, setStopText] = ne.useState(""),
    [dateUnit, setDateUnit] = ne.useState("day"),
    [monthPolicy, setMonthPolicy] = ne.useState("preserveDay"),
    [viewAxis, setViewAxis] = ne.useState("row"),
    [X$1, setHiddenPage] = ne.useState(0),
    [previewPage, setPreviewPage] = ne.useState(0),
    [expectedSnapshot] = ne.useState(() => buildSnapshotGuard(props.workbook)),
    [defaultRowHeightText, setDefaultRowHeightText] = ne.useState(() =>
      String(defaultRowHeightOf(sheetOrThrow)),
    ),
    [defaultColumnWidthText, setDefaultColumnWidthText] = ne.useState(() =>
      String(defaultColumnWidthOf(sheetOrThrow)),
    ),
    [resetRowOverrides, setResetRowOverrides] = ne.useState(false),
    [resetColumnOverrides, setResetColumnOverrides] = ne.useState(false),
    customRowHeightCount = Object.keys(sheetOrThrow.rowHeights ?? {}).length,
    customColumnWidthCount = Object.keys(
      sheetOrThrow.columnWidths ?? {},
    ).length,
    isReadOnly = props.readOnly || !!sheetOrThrow.protection?.enabled,
    isStale =
      (props.mode === "viewSettings" &&
        JSON.stringify(props.workbook) !== expectedSnapshot.snapshotGuard) ||
      openedSnapshotRef.current.workbookId !== props.workbook.id ||
      openedSnapshotRef.current.revision !== props.workbook.revision ||
      openedSnapshotRef.current.sheetId !== props.sheetId ||
      openedSnapshotRef.current.range !== props.range ||
      (props.mode === "pasteSpecial" &&
        openedSnapshotRef.current.clipboard !== props.clipboard),
    staleMessage =
      "The workbook changed or the selection moved. Close this dialog and reopen it to review the current cells.",
    calculatedWorkbook = ne.useMemo(
      () => calculateWorkbookWithSettings(props.workbook),
      [props.workbook],
    ),
    previewResult = ne.useMemo(() => {
      if (props.mode === "viewSettings") return { plan: null, error: "" };
      try {
        if (isStale) throw new Error(staleMessage);
        if (props.mode === "pasteSpecial")
          return props.clipboard
            ? {
                plan: buildPastePreview(
                  props.workbook,
                  props.sheetId,
                  props.range,
                  props.clipboard,
                  {
                    mode: mode,
                    operation: operation,
                    transpose: transpose,
                    skipBlanks: skipBlanks,
                  },
                ),
                error: "",
              }
            : { plan: null, error: "" };
        if (fillMode === "copy")
          return {
            plan: buildFillPreview(
              props.workbook,
              props.sheetId,
              props.range,
              fillDirection,
            ),
            error: "",
          };
        if (!stepText.trim()) throw new Error("Enter a step value.");
        const toSeriesValue = (inputText) =>
          seriesKind === "date" ? inputText : Number(inputText);
        return {
          plan: buildSeriesPreview(props.workbook, props.sheetId, props.range, {
            kind: seriesKind,
            direction:
              seriesAxis === "rows"
                ? isBackward
                  ? "left"
                  : "right"
                : isBackward
                  ? "up"
                  : "down",
            ...(startText.trim() ? { start: toSeriesValue(startText) } : {}),
            step: Number(stepText),
            ...(stopText.trim() ? { stop: toSeriesValue(stopText) } : {}),
            dateUnit: dateUnit,
            monthPolicy: monthPolicy,
          }),
          error: "",
        };
      } catch (error) {
        return { plan: null, error: describeWorkbookError(error) };
      }
    }, [
      props.mode,
      props.workbook,
      props.sheetId,
      props.range,
      props.clipboard,
      isStale,
      mode,
      operation,
      transpose,
      skipBlanks,
      fillMode,
      fillDirection,
      seriesAxis,
      isBackward,
      seriesKind,
      startText,
      stepText,
      stopText,
      dateUnit,
      monthPolicy,
    ]),
    hiddenIntervals =
      props.mode === "viewSettings"
        ? readHiddenAxisIntervals(sheetOrThrow, viewAxis)
        : [],
    hiddenPageCount = Math.max(1, Math.ceil(hiddenIntervals.length / 20)),
    previewPageCount = Math.max(
      1,
      Math.ceil((previewResult.plan?.previewPatches.length ?? 0) / 12),
    ),
    clampedPreviewPage = Math.min(previewPage, previewPageCount - 1);
  (ne.useEffect(() => {
    setPreviewPage(0);
  }, [previewResult.plan]),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        dialogRef.current
          ?.querySelector("button:not(:disabled),input:not(:disabled)")
          ?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  function assertCanApply() {
    const currentProps = propsRef.current,
      openedSnapshot = openedSnapshotRef.current;
    if (
      currentProps.readOnly ||
      findSheetOrThrow(currentProps.workbook, currentProps.sheetId).protection
        ?.enabled
    )
      throw new Error("This worksheet is read-only.");
    if (
      currentProps.workbook.id !== openedSnapshot.workbookId ||
      (currentProps.mode === "viewSettings" &&
        JSON.stringify(currentProps.workbook) !==
          expectedSnapshot.snapshotGuard) ||
      currentProps.workbook.revision !== openedSnapshot.revision ||
      currentProps.sheetId !== openedSnapshot.sheetId ||
      currentProps.range !== openedSnapshot.range ||
      (currentProps.mode === "pasteSpecial" &&
        currentProps.clipboard !== openedSnapshot.clipboard)
    )
      throw new Error(staleMessage);
  }
  function applyChange(buildNextWorkbook, changeLabel, selectionRange) {
    var currentProps;
    try {
      (setActionError(""), assertCanApply());
      const nextWorkbook = buildNextWorkbook();
      (JSON.stringify(nextWorkbook) !==
        JSON.stringify(propsRef.current.workbook) &&
        propsRef.current.onWorkbookChange(nextWorkbook, changeLabel),
        selectionRange &&
          (currentProps = propsRef.current).onSelectionChange?.call(
            currentProps,
            selectionRange,
          ),
        propsRef.current.onClose());
    } catch (error) {
      setActionError(describeWorkbookError(error));
    }
  }
  function onClick() {
    const plan = previewResult.plan;
    if (!plan) {
      setActionError(
        previewResult.error ||
          "Choose a valid source and target before applying this change.",
      );
      return;
    }
    "sourceRange" in plan
      ? applyChange(
          () =>
            applyCellPreview(propsRef.current.workbook, plan, {
              readOnly: propsRef.current.readOnly,
            }),
          "Paste Special",
          plan.targetRange,
        )
      : applyChange(
          () =>
            applyCellPreview(propsRef.current.workbook, plan, {
              readOnly: propsRef.current.readOnly,
            }),
          fillMode === "series" ? "Fill series" : `Fill ${fillDirection}`,
          plan.targetRange,
        );
  }
  const formatAxisRange = (startIndex, endIndex) =>
      viewAxis === "row"
        ? `Rows ${startIndex + 1}${endIndex === startIndex ? "" : `–${endIndex + 1}`}`
        : `Columns ${columnIndexToLetters(startIndex)}${endIndex === startIndex ? "" : `–${columnIndexToLetters(endIndex)}`}`,
    dialogTitle =
      props.mode === "pasteSpecial"
        ? "Paste Special"
        : props.mode === "fillSeries"
          ? "Fill Series"
          : "Worksheet View",
    dialogIcon =
      props.mode === "pasteSpecial"
        ? Clipboard
        : props.mode === "fillSeries"
          ? Rows3
          : Grid2x2;
  return i.jsx("div", {
    className: "et-backdrop",
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "et-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      tabIndex: -1,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(), propsRef.current.onClose()),
          event.key === "Tab")
        ) {
          const focusables = Array.from(
              dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ),
            firstFocusable = focusables[0],
            lastFocusable = focusables.at(-1);
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
          className: "et-header",
          children: [
            i.jsxs("h2", {
              id: titleId,
              children: [i.jsx(dialogIcon, { size: 19 }), dialogTitle],
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": text("Close editing tools"),
              onClick: props.onClose,
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
                  children: [sheetOrThrow.name, "!", props.range],
                }),
              ],
            }),
            isReadOnly &&
              i.jsx("p", {
                className: "et-warning",
                children: i.jsx(LocalizedText, {
                  children:
                    "This worksheet is read-only. You can review the options.",
                }),
              }),
            isStale &&
              i.jsx("p", {
                className: "et-warning",
                role: "alert",
                children: staleMessage,
              }),
            props.mode === "pasteSpecial" &&
              i.jsxs(i.Fragment, {
                children: [
                  props.clipboard
                    ? i.jsxs("p", {
                        className: "et-source",
                        children: [
                          i.jsx(LocalizedText, { children: "Copied cells " }),
                          i.jsx("strong", {
                            children: props.clipboard.sourceRange,
                          }),
                          " ·",
                          " ",
                          props.clipboard.rowCount,
                          " ",
                          i.jsx(LocalizedText, { children: "rows × " }),
                          props.clipboard.colCount,
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
                    disabled: isReadOnly || isStale || !props.clipboard,
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
                        ].map(([optionValue, optionLabel, optionDescription]) =>
                          i.jsxs(
                            "label",
                            {
                              className: "et-radio",
                              children: [
                                i.jsx("input", {
                                  type: "radio",
                                  name: `${titleId}-paste`,
                                  checked: mode === optionValue,
                                  onChange: () => setMode(optionValue),
                                }),
                                i.jsxs("span", {
                                  children: [
                                    optionLabel,
                                    i.jsx("small", {
                                      children: optionDescription,
                                    }),
                                  ],
                                }),
                              ],
                            },
                            optionValue,
                          ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("fieldset", {
                    disabled: isReadOnly || isStale || !props.clipboard,
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
                        ].map((operationName) =>
                          i.jsxs(
                            "label",
                            {
                              children: [
                                i.jsx("input", {
                                  type: "radio",
                                  name: `${titleId}-operation`,
                                  checked: operation === operationName,
                                  onChange: () => setOperation(operationName),
                                }),
                                operationName[0].toUpperCase() +
                                  operationName.slice(1),
                              ],
                            },
                            operationName,
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
                                checked: skipBlanks,
                                onChange: (event) =>
                                  setSkipBlanks(event.target.checked),
                              }),
                              i.jsx(LocalizedText, { children: "Skip blanks" }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: transpose,
                                onChange: (event) =>
                                  setTranspose(event.target.checked),
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
            props.mode === "fillSeries" &&
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
                      const nextTab =
                        event.key === "Home"
                          ? "series"
                          : event.key === "End" || fillMode === "series"
                            ? "copy"
                            : "series";
                      (setFillMode(nextTab),
                        event.currentTarget
                          .querySelector(`[data-fill-tab="${nextTab}"]`)
                          ?.focus());
                    },
                    lang: language("Fill options"),
                    "data-ui-label": "Fill options",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        role: "tab",
                        id: `${titleId}-series`,
                        "data-fill-tab": "series",
                        "aria-controls": `${titleId}-fill-panel`,
                        tabIndex: fillMode === "series" ? 0 : -1,
                        "aria-selected": fillMode === "series",
                        onClick: () => setFillMode("series"),
                        children: i.jsx(LocalizedText, { children: "Series" }),
                      }),
                      i.jsx("button", {
                        type: "button",
                        role: "tab",
                        id: `${titleId}-copy`,
                        "data-fill-tab": "copy",
                        "aria-controls": `${titleId}-fill-panel`,
                        tabIndex: fillMode === "copy" ? 0 : -1,
                        "aria-selected": fillMode === "copy",
                        onClick: () => setFillMode("copy"),
                        children: i.jsx(LocalizedText, {
                          children: "Copy cells",
                        }),
                      }),
                    ],
                  }),
                  i.jsx("div", {
                    role: "tabpanel",
                    id: `${titleId}-fill-panel`,
                    "aria-labelledby": `${titleId}-${fillMode}`,
                    children:
                      fillMode === "copy"
                        ? i.jsxs("fieldset", {
                            disabled: isReadOnly || isStale,
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
                                        "aria-pressed": fillDirection === value,
                                        onClick: () => setFillDirection(value),
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
                            disabled: isReadOnly || isStale,
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
                                        value: seriesAxis,
                                        onChange: (event) =>
                                          setSeriesAxis(event.target.value),
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
                                        value: isBackward
                                          ? "backward"
                                          : "forward",
                                        onChange: (event) =>
                                          setIsBackward(
                                            event.target.value === "backward",
                                          ),
                                        lang: language("Series order"),
                                        "data-ui-label": "Series order",
                                        children: [
                                          i.jsxs("option", {
                                            value: "forward",
                                            children: [
                                              "Forward (",
                                              seriesAxis === "rows"
                                                ? "right"
                                                : "down",
                                              ")",
                                            ],
                                          }),
                                          i.jsxs("option", {
                                            value: "backward",
                                            children: [
                                              "Backward (",
                                              seriesAxis === "rows"
                                                ? "left"
                                                : "up",
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
                                        value: seriesKind,
                                        onChange: (event) => {
                                          (setSeriesKind(event.target.value),
                                            setStartText(""),
                                            setStopText(""));
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
                                        type:
                                          seriesKind === "date"
                                            ? "date"
                                            : "number",
                                        step: "any",
                                        value: startText,
                                        placeholder: text("Use first cell"),
                                        onChange: (event) =>
                                          setStartText(event.target.value),
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
                                        step:
                                          seriesKind === "date" ? "1" : "any",
                                        value: stepText,
                                        onValueChange: (stepValue) =>
                                          setStepText(stepValue),
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
                                        type:
                                          seriesKind === "date"
                                            ? "date"
                                            : "number",
                                        step: "any",
                                        value: stopText,
                                        onChange: (event) =>
                                          setStopText(event.target.value),
                                        lang: language("Stop value"),
                                        "data-ui-label": "Stop value",
                                      }),
                                    ],
                                  }),
                                ],
                              }),
                              seriesKind === "date" &&
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
                                            setDateUnit(event.target.value),
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
                                              setMonthPolicy(
                                                event.target.value,
                                              ),
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
            props.mode === "viewSettings" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("fieldset", {
                    disabled: isReadOnly || isStale,
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
                                value: defaultRowHeightText,
                                onValueChange: (nextRowHeight) =>
                                  setDefaultRowHeightText(nextRowHeight),
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
                                value: defaultColumnWidthText,
                                onValueChange: (nextColumnWidth) =>
                                  setDefaultColumnWidthText(nextColumnWidth),
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
                            checked: resetRowOverrides,
                            onChange: (event) =>
                              setResetRowOverrides(event.target.checked),
                          }),
                          i.jsx(LocalizedText, { children: "Reset " }),
                          customRowHeightCount.toLocaleString(),
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
                            checked: resetColumnOverrides,
                            onChange: (event) =>
                              setResetColumnOverrides(event.target.checked),
                          }),
                          i.jsx(LocalizedText, { children: "Reset " }),
                          customColumnWidthCount.toLocaleString(),
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
                                sheetOrThrow.rowCount -
                                (resetRowOverrides ? 0 : customRowHeightCount)
                              ).toLocaleString(),
                              " ",
                              i.jsx(LocalizedText, {
                                children: "rows inherit ",
                              }),
                              defaultRowHeightText || "—",
                              " ",
                              i.jsx(LocalizedText, { children: "px;" }),
                              " ",
                              resetRowOverrides
                                ? 0
                                : customRowHeightCount.toLocaleString(),
                              " ",
                              i.jsx(LocalizedText, {
                                children: "individual heights remain.",
                              }),
                            ],
                          }),
                          i.jsxs("p", {
                            children: [
                              (
                                sheetOrThrow.colCount -
                                (resetColumnOverrides
                                  ? 0
                                  : customColumnWidthCount)
                              ).toLocaleString(),
                              " ",
                              i.jsx(LocalizedText, {
                                children: "columns inherit ",
                              }),
                              defaultColumnWidthText || "—",
                              " ",
                              i.jsx(LocalizedText, { children: "px;" }),
                              " ",
                              resetColumnOverrides
                                ? 0
                                : customColumnWidthCount.toLocaleString(),
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
                          applyChange(
                            () =>
                              applySheetSizeEdits(
                                propsRef.current.workbook,
                                propsRef.current.sheetId,
                                {
                                  rowHeight: Number(defaultRowHeightText),
                                  columnWidth: Number(defaultColumnWidthText),
                                  resetRowOverrides: resetRowOverrides,
                                  resetColumnOverrides: resetColumnOverrides,
                                },
                                {
                                  readOnly: propsRef.current.readOnly,
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
                    disabled: isReadOnly || isStale,
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
                          ["horizontal", "vertical", "both"].map(
                            (splitOrientation) =>
                              i.jsx(
                                "button",
                                {
                                  type: "button",
                                  onClick: () =>
                                    applyChange(
                                      () =>
                                        setSplitPanes(
                                          props.workbook,
                                          props.sheetId,
                                          {
                                            row: rangeAddress.firstRow,
                                            col: rangeAddress.firstCol,
                                            ...(splitOrientation !== "vertical"
                                              ? { horizontal: 50 }
                                              : {}),
                                            ...(splitOrientation !==
                                            "horizontal"
                                              ? { vertical: 50 }
                                              : {}),
                                          },
                                        ),
                                      `Split worksheet ${splitOrientation}`,
                                    ),
                                  children:
                                    splitOrientation === "horizontal"
                                      ? "Split top and bottom"
                                      : splitOrientation === "vertical"
                                        ? "Split left and right"
                                        : "Split into four panes",
                                },
                                splitOrientation,
                              ),
                          ),
                          i.jsx("button", {
                            type: "button",
                            disabled: !sheetOrThrow.splitPanes,
                            onClick: () =>
                              applyChange(
                                () =>
                                  setSplitPanes(props.workbook, props.sheetId),
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
                    disabled: isReadOnly || isStale,
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
                              applyChange(
                                () =>
                                  setFreezePanesAtCell(
                                    props.workbook,
                                    props.sheetId,
                                    anchorAddress,
                                  ),
                                `Freeze panes at ${anchorAddress}`,
                              ),
                            children: [
                              i.jsx(LocalizedText, { children: "Freeze at " }),
                              anchorAddress,
                            ],
                          }),
                          i.jsx("button", {
                            type: "button",
                            onClick: () =>
                              applyChange(
                                () =>
                                  setFreezePanes(
                                    props.workbook,
                                    props.sheetId,
                                    1,
                                    0,
                                  ),
                                "Freeze top row",
                              ),
                            children: i.jsx(LocalizedText, {
                              children: "Freeze top row",
                            }),
                          }),
                          i.jsx("button", {
                            type: "button",
                            onClick: () =>
                              applyChange(
                                () =>
                                  setFreezePanes(
                                    props.workbook,
                                    props.sheetId,
                                    0,
                                    1,
                                  ),
                                "Freeze first column",
                              ),
                            children: i.jsx(LocalizedText, {
                              children: "Freeze first column",
                            }),
                          }),
                          i.jsx("button", {
                            type: "button",
                            onClick: () =>
                              applyChange(
                                () =>
                                  setFreezePanes(
                                    props.workbook,
                                    props.sheetId,
                                    0,
                                    0,
                                  ),
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
                    disabled: isReadOnly || isStale,
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
                              applyChange(
                                () =>
                                  setHiddenAxisRange(
                                    props.workbook,
                                    props.sheetId,
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
                              applyChange(
                                () =>
                                  setHiddenAxisRange(
                                    props.workbook,
                                    props.sheetId,
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
                            value: viewAxis,
                            onChange: (event) => {
                              (setViewAxis(event.target.value),
                                setHiddenPage(0));
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
                      hiddenIntervals.length
                        ? i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("ul", {
                                children: hiddenIntervals
                                  .slice(X$1 * 20, (X$1 + 1) * 20)
                                  .map((hiddenInterval) =>
                                    i.jsxs(
                                      "li",
                                      {
                                        children: [
                                          i.jsx("span", {
                                            children: formatAxisRange(
                                              hiddenInterval.start,
                                              hiddenInterval.end,
                                            ),
                                          }),
                                          i.jsx("button", {
                                            type: "button",
                                            "aria-label": `Unhide ${formatAxisRange(hiddenInterval.start, hiddenInterval.end).toLowerCase()}`,
                                            disabled: isReadOnly || isStale,
                                            onClick: () =>
                                              applyChange(
                                                () =>
                                                  setHiddenAxisRange(
                                                    props.workbook,
                                                    props.sheetId,
                                                    viewAxis,
                                                    hiddenInterval.start,
                                                    hiddenInterval.end,
                                                    false,
                                                  ),
                                                `Unhide ${viewAxis === "row" ? "rows" : "columns"}`,
                                              ),
                                            children: i.jsx(LocalizedText, {
                                              children: "Unhide",
                                            }),
                                          }),
                                        ],
                                      },
                                      `${hiddenInterval.start}-${hiddenInterval.end}`,
                                    ),
                                  ),
                              }),
                              hiddenPageCount > 1 &&
                                i.jsxs("div", {
                                  className: "et-pagination",
                                  children: [
                                    i.jsx("button", {
                                      type: "button",
                                      disabled: !X$1,
                                      onClick: () =>
                                        setHiddenPage(
                                          (currentPage) => currentPage - 1,
                                        ),
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
                                        hiddenPageCount,
                                      ],
                                    }),
                                    i.jsx("button", {
                                      type: "button",
                                      disabled: X$1 + 1 >= hiddenPageCount,
                                      onClick: () =>
                                        setHiddenPage(
                                          (currentPage) => currentPage + 1,
                                        ),
                                      children: i.jsx(LocalizedText, {
                                        children: "Next hidden items",
                                      }),
                                    }),
                                  ],
                                }),
                              i.jsxs("button", {
                                type: "button",
                                className: "et-unhide-all",
                                disabled: isReadOnly || isStale,
                                onClick: () =>
                                  applyChange(
                                    () =>
                                      yQ(
                                        props.workbook,
                                        props.sheetId,
                                        viewAxis,
                                      ),
                                    `Unhide all ${viewAxis === "row" ? "rows" : "columns"}`,
                                  ),
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Unhide all ",
                                  }),
                                  viewAxis === "row" ? "rows" : "columns",
                                ],
                              }),
                            ],
                          })
                        : i.jsxs("p", {
                            className: "et-muted",
                            children: [
                              i.jsx(LocalizedText, { children: "No " }),
                              viewAxis === "row" ? "rows" : "columns",
                              " ",
                              i.jsx(LocalizedText, { children: "are hidden." }),
                            ],
                          }),
                    ],
                  }),
                ],
              }),
            previewResult.plan &&
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
                          previewResult.plan.count.toLocaleString(),
                          " ",
                          i.jsx(LocalizedText, { children: "cells to change" }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      sheetOrThrow.name,
                      "!",
                      previewResult.plan.targetRange,
                    ],
                  }),
                  previewResult.plan.dependentCount > 0 &&
                    i.jsxs("p", {
                      className: "et-muted",
                      children: [
                        i.jsx(LocalizedText, {
                          children: "This operation also updates",
                        }),
                        " ",
                        previewResult.plan.dependentCount.toLocaleString(),
                        " ",
                        i.jsx(LocalizedText, {
                          children:
                            "related cells. Every changed cell is included in the preview pages below.",
                        }),
                      ],
                    }),
                  previewResult.plan.previewPatches.length > 0
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
                                  children: previewResult.plan.previewPatches
                                    .slice(
                                      clampedPreviewPage * 12,
                                      (clampedPreviewPage + 1) * 12,
                                    )
                                    .map((patch) => {
                                      const sheetOrThrow = findSheetOrThrow(
                                        props.workbook,
                                        patch.sheetId,
                                      );
                                      return i.jsxs(
                                        "tr",
                                        {
                                          children: [
                                            i.jsx("th", {
                                              children:
                                                patch.sheetId === props.sheetId
                                                  ? patch.address
                                                  : `${sheetOrThrow.name}!${patch.address}`,
                                            }),
                                            i.jsx("td", {
                                              children:
                                                formatNumberWithPattern(
                                                  calculatedWorkbook(
                                                    patch.sheetId,
                                                    patch.address,
                                                  ),
                                                  sheetOrThrow.cells[
                                                    patch.address
                                                  ],
                                                  props.workbook.datePolicy
                                                    ?.dateSystem,
                                                ) || "(blank)",
                                            }),
                                            i.jsxs("td", {
                                              children: [
                                                formatNumberWithPattern(
                                                  patch.cell.raw,
                                                  patch.cell,
                                                  props.workbook.datePolicy
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
                          previewPageCount > 1 &&
                            i.jsxs("div", {
                              className: "et-pagination",
                              "aria-label": text("Preview pages"),
                              lang: language("Preview pages"),
                              "data-ui-label": "Preview pages",
                              children: [
                                i.jsx("button", {
                                  type: "button",
                                  disabled: clampedPreviewPage === 0,
                                  onClick: () =>
                                    setPreviewPage(
                                      (currentPage) => currentPage - 1,
                                    ),
                                  children: i.jsx(LocalizedText, {
                                    children: "Previous changes",
                                  }),
                                }),
                                i.jsxs("span", {
                                  children: [
                                    i.jsx(LocalizedText, { children: "Page " }),
                                    clampedPreviewPage + 1,
                                    " ",
                                    i.jsx(LocalizedText, { children: "of " }),
                                    previewPageCount,
                                  ],
                                }),
                                i.jsx("button", {
                                  type: "button",
                                  disabled:
                                    clampedPreviewPage + 1 >= previewPageCount,
                                  onClick: () =>
                                    setPreviewPage(
                                      (currentPage) => currentPage + 1,
                                    ),
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
            !isStale &&
              (actionError || previewResult.error) &&
              i.jsx("p", {
                className: "et-error",
                role: "alert",
                children: actionError || previewResult.error,
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "et-footer",
          children: [
            i.jsx("button", {
              type: "button",
              onClick: props.onClose,
              children: props.mode === "viewSettings" ? "Done" : "Cancel",
            }),
            props.mode !== "viewSettings" &&
              i.jsx("button", {
                type: "button",
                className: "et-primary",
                disabled:
                  isReadOnly ||
                  isStale ||
                  !previewResult.plan ||
                  previewResult.plan.count === 0,
                onClick: onClick,
                children:
                  props.mode === "pasteSpecial"
                    ? "Paste"
                    : fillMode === "series"
                      ? "Fill series"
                      : "Apply fill",
              }),
          ],
        }),
      ],
    }),
  });
}
function FillPreviewDialog(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    initialWorkbookRef = ne.useRef(props.workbook),
    latestPropsRef = ne.useRef(props);
  latestPropsRef.current = props;
  const [mode, setMode] = ne.useState("copy"),
    [actionError, setActionError] = ne.useState(""),
    [previewPage, setPreviewPage] = ne.useState(0),
    isWorkbookChanged = props.workbook !== initialWorkbookRef.current,
    previewResult = ne.useMemo(() => {
      try {
        return {
          plan: buildFillHandlePreview(
            props.workbook,
            props.sheetId,
            props.sourceRange,
            props.targetRange,
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
    }, [
      props.workbook,
      props.sheetId,
      props.sourceRange,
      props.targetRange,
      mode,
    ]),
    plan = previewResult.plan;
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const disabled =
    props.readOnly ||
    !!findSheetOrThrow(props.workbook, props.sheetId).protection?.enabled ||
    isWorkbookChanged;
  return i.jsx("div", {
    className: "et-backdrop",
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "et-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      tabIndex: -1,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.preventDefault(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusables = Array.from(
              dialogRef.current?.querySelectorAll(
                "button:not(:disabled),select:not(:disabled),input:not(:disabled)",
              ) ?? [],
            ),
            firstFocusable = focusables[0],
            lastFocusable = focusables.at(-1);
          event.shiftKey &&
          (document.activeElement === firstFocusable ||
            document.activeElement === dialogRef.current)
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              (document.activeElement === lastFocusable ||
                document.activeElement === dialogRef.current) &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "et-header",
          children: [
            i.jsx("h2", {
              id: titleId,
              children: i.jsx(LocalizedText, { children: "Fill selection" }),
            }),
            i.jsx("button", {
              "aria-label": text("Close fill preview"),
              onClick: props.onClose,
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
                i.jsx("strong", { children: props.sourceRange }),
                " ",
                i.jsx(LocalizedText, { children: "to" }),
                " ",
                i.jsx("strong", { children: props.targetRange }),
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
                    (setMode(event.target.value),
                      setPreviewPage(0),
                      setActionError(""));
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
                            .slice(previewPage * 25, (previewPage + 1) * 25)
                            .map((patch) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("th", {
                                      children:
                                        patch.sheetId === props.sheetId
                                          ? patch.address
                                          : `${findSheetOrThrow(props.workbook, patch.sheetId).name}!${patch.address}`,
                                    }),
                                    i.jsx("td", {
                                      children:
                                        formatNumberWithPattern(
                                          patch.cell.raw,
                                          patch.cell,
                                          props.workbook.datePolicy?.dateSystem,
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
                          disabled: !previewPage,
                          onClick: () => setPreviewPage(previewPage - 1),
                          children: i.jsx(LocalizedText, {
                            children: "Previous changes",
                          }),
                        }),
                        i.jsxs("span", {
                          children: [
                            i.jsx(LocalizedText, { children: "Page " }),
                            previewPage + 1,
                            " ",
                            i.jsx(LocalizedText, { children: "of " }),
                            Math.ceil(plan.count / 25),
                          ],
                        }),
                        i.jsx("button", {
                          disabled: (previewPage + 1) * 25 >= plan.count,
                          onClick: () => setPreviewPage(previewPage + 1),
                          children: i.jsx(LocalizedText, {
                            children: "Next changes",
                          }),
                        }),
                      ],
                    }),
                ],
              }),
            (actionError || previewResult.error || isWorkbookChanged) &&
              i.jsx("p", {
                className: "et-error",
                role: "alert",
                children: isWorkbookChanged
                  ? "The workbook changed. Cancel and review the fill again."
                  : actionError || previewResult.error,
              }),
            props.readOnly &&
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
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "et-primary",
              disabled: disabled || !plan || plan.count === 0,
              onClick: () => {
                try {
                  const latestProps = latestPropsRef.current;
                  if (latestProps.workbook !== initialWorkbookRef.current)
                    throw Error("The workbook changed. Review the fill again.");
                  if (!plan) throw Error("Choose a valid fill first.");
                  const filledWorkbook = applyCellPreview(
                    latestProps.workbook,
                    plan,
                    {
                      readOnly: latestProps.readOnly,
                    },
                  );
                  (latestProps.onWorkbookChange(filledWorkbook, "Fill handle"),
                    latestProps.onSelectionChange(latestProps.targetRange),
                    latestProps.onClose());
                } catch (error) {
                  setActionError(
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

const rangeContainsRange = (outerRange, innerRange) =>
    outerRange.firstRow <= innerRange.firstRow &&
    outerRange.lastRow >= innerRange.lastRow &&
    outerRange.firstCol <= innerRange.firstCol &&
    outerRange.lastCol >= innerRange.lastCol,
  hasCellContent = (cell) =>
    !!cell &&
    (cell.raw !== null || Object.keys(cell).some((key) => key !== "raw"));

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
  const rangeAddress = assertRangeWithinBoundary(sheetOrThrow, targetRange);
  return (
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      pasteOperationLimits.sourceCells &&
      throwWorkbookError("Cut and move supports at most 500 source cells."),
    rangeAddress
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
function verifyCutSourceSnapshot(book, snapshot, sourceSheetId, sourceRange) {
  (!snapshot ||
    snapshot.version !== 1 ||
    snapshot.workbookId !== book.id ||
    snapshot.revision !== book.revision ||
    snapshot.sheetId !== sourceSheetId ||
    snapshot.range !== formatSortRangeAddress(sourceRange) ||
    snapshot.snapshotGuard !== JSON.stringify(book)) &&
    throwWorkbookError(
      "The workbook changed after Cut. Cut the source again before moving it.",
    );
}
const isUnsafeStructuredReference = (reference) =>
    reference.kind === "structured" &&
    (!reference.name || /\[\s*@|#This\s+Row/i.test(reference.text)),
  sheetNameReferencePattern = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*)`,
  sheetNamePrefixPattern = new RegExp(`^(${sheetNameReferencePattern})!`),
  sheetRangePrefixPattern = new RegExp(
    `^(${sheetNameReferencePattern}):(${sheetNameReferencePattern})!`,
  ),
  identifierCharPattern = /[A-Za-z0-9_.\\\u0080-\uffff]/;
function assertBalancedStructuredReference(refersTo, end) {
  let depth = 0,
    scanIndex = end;
  do
    (refersTo[scanIndex] === "[" && depth++,
      refersTo[scanIndex] === "]" && depth--,
      scanIndex++);
  while (scanIndex < refersTo.length && depth);
  return (
    depth &&
      throwWorkbookError(
        "An unbalanced structured reference prevents a safe move.",
      ),
    scanIndex
  );
}
function assertFormulaReferencesMovable(refersTo) {
  const tokens = [];
  let position = 0;
  for (; position < refersTo.length;) {
    const start = position,
      character = refersTo[position];
    if (character === '"') {
      for (position++; position < refersTo.length;)
        if (refersTo[position++] === '"')
          if (refersTo[position] === '"') position++;
          else break;
      continue;
    }
    const remainingText = refersTo.slice(position),
      isTokenBoundary =
        position === 0 || !identifierCharPattern.test(refersTo[position - 1]);
    if (
      (isTokenBoundary &&
        sheetRangePrefixPattern.test(remainingText) &&
        throwWorkbookError(
          "Three-dimensional worksheet references must be replaced with explicit local references before moving cells.",
        ),
      character === "#")
    ) {
      const errorMatch = /^#[A-Za-z0-9/?]+[!?]?/.exec(remainingText);
      if (errorMatch) {
        position += errorMatch[0].length;
        continue;
      }
    }
    let sheet,
      prefixLength = 0;
    const sheetPrefixMatch = isTokenBoundary
      ? sheetNamePrefixPattern.exec(remainingText)
      : null;
    sheetPrefixMatch &&
      ((sheet = unquoteSheetName(sheetPrefixMatch[1])),
      /[\[\]:]/.test(sheet) &&
        throwWorkbookError(
          "External or three-dimensional references prevent a safe range move.",
        ),
      (prefixLength = sheetPrefixMatch[0].length));
    const referenceText = remainingText.slice(prefixLength),
      a1Match = isTokenBoundary
        ? /^(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)/.exec(
            referenceText,
          )
        : null;
    if (
      a1Match &&
      !identifierCharPattern.test(referenceText[a1Match[0].length] ?? "") &&
      !/^\s*[(!]/.test(referenceText.slice(a1Match[0].length))
    ) {
      const end = position + prefixLength + a1Match[0].length;
      (refersTo[end] === "#" &&
        throwWorkbookError(
          "Dynamic spill references require an explicit stable range before moving cells.",
        ),
        refersTo[end] === ":" &&
          throwWorkbookError(
            "Separately qualified range endpoints are not supported for a safe move.",
          ),
        tokens.push({
          start: start,
          end: end,
          text: refersTo.slice(start, end),
          kind: "reference",
          sheet: sheet,
          reference: a1Match[0],
        }),
        (position = end));
      continue;
    }
    const nameMatch = isTokenBoundary
      ? /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(referenceText)
      : null;
    if (nameMatch) {
      const end = position + prefixLength + nameMatch[0].length,
        textAfterReference = refersTo.slice(end);
      if (refersTo[end] === "[") {
        const structuredEnd = assertBalancedStructuredReference(refersTo, end);
        (refersTo[structuredEnd] === "!" &&
          throwWorkbookError(
            "External workbook references prevent a safe move.",
          ),
          tokens.push({
            start: start,
            end: structuredEnd,
            text: refersTo.slice(start, structuredEnd),
            kind: "structured",
            sheet: sheet,
            name: nameMatch[0],
          }),
          (position = structuredEnd));
        continue;
      }
      if (/^\s*\(/.test(textAfterReference)) {
        (/^(?:_xlfn\.)?(?:INDIRECT|OFFSET|EVALUATE|HYPERLINK)$/i.test(
          nameMatch[0],
        ) &&
          tokens.push({
            start: start,
            end: end,
            text: nameMatch[0],
            kind: "dynamic",
            name: nameMatch[0],
          }),
          (position = end));
        continue;
      }
      (tokens.push({
        start: start,
        end: end,
        text: refersTo.slice(start, end),
        kind: "name",
        sheet: sheet,
        name: nameMatch[0],
      }),
        (position = end));
      continue;
    }
    if (character === "[") {
      const end = assertBalancedStructuredReference(refersTo, position);
      (/^\s*(?:[^!+*/(),;]+)!/.test(refersTo.slice(end)) &&
        throwWorkbookError("External workbook references prevent a safe move."),
        tokens.push({
          start: start,
          end: end,
          text: refersTo.slice(start, end),
          kind: "structured",
        }),
        (position = end));
      continue;
    }
    if (
      (sheetPrefixMatch &&
        throwWorkbookError(
          "A worksheet-qualified expression is not a supported reference for a safe move.",
        ),
      character === "'")
    ) {
      const quotedSheetMatch = /^'(?:[^']|'')*'/.exec(remainingText);
      if (quotedSheetMatch) {
        (/\[[^\]]+\]/.test(quotedSheetMatch[0]) &&
          refersTo[position + quotedSheetMatch[0].length] === "!" &&
          throwWorkbookError(
            "External workbook references prevent a safe move.",
          ),
          (position += quotedSheetMatch[0].length));
        continue;
      }
    }
    position++;
  }
  return tokens;
}
function parseWholeColumnReference(reference) {
  if (/^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/.test(reference)) {
    const [startColumn, endColumn] = reference.split(":"),
      startCell = _t(`${startColumn}1`),
      endCell = _t(`${endColumn}1`);
    return {
      firstRow: 0,
      lastRow: workbookLimits.maxRows - 1,
      firstCol: Math.min(startCell.col, endCell.col),
      lastCol: Math.max(startCell.col, endCell.col),
    };
  }
  if (/^\$?\d+:\$?\d+$/.test(reference)) {
    const [startRow, endRow] = reference
      .replace(/\$/g, "")
      .split(":")
      .map(Number);
    return (
      (startRow > workbookLimits.maxRows || endRow > workbookLimits.maxRows) &&
        throwWorkbookError("A reference exceeds Excel's row boundary."),
      {
        firstRow: Math.min(startRow, endRow) - 1,
        lastRow: Math.max(startRow, endRow) - 1,
        firstCol: 0,
        lastCol: workbookLimits.maxCols - 1,
      }
    );
  }
  return parseRangeAddress(reference);
}
function parseMovableA1Rectangle(reference, rowOffset, columnOffset) {
  return reference
    .split(":")
    .map((cellAddress) => {
      const addressMatch = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(
        cellAddress,
      );
      addressMatch ||
        throwWorkbookError(
          "Only complete A1 cell rectangles can move to a new location.",
        );
      const cellPosition = _t(cellAddress);
      return `${addressMatch[1]}${columnIndexToLetters(cellPosition.col + columnOffset)}${addressMatch[3]}${cellPosition.row + rowOffset + 1}`;
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
