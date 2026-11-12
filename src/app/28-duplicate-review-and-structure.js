import { isFormulaCellValue } from "./03-workbook-model.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  findRowOrigin,
  calculateWorkbookWithSettings,
  _t,
  parseRangeAddress,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  applyStructuralEdit,
  createDuplicateReviewSheet,
  describeCellValueType,
  assertInsertionWithinSheet,
  structuralEditLimits,
  qk,
  applyDuplicateReviewPreview,
} from "./27-validation-and-structural-edits.js";

function DuplicateReviewDialog(props) {
  const { text, language } = useWorkbookMessages(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    onCloseRef = ne.useRef(props.onClose);
  onCloseRef.current = props.onClose;
  const [range, setRange] = ne.useState(props.range),
    [keyColumns, setKeyColumns] = ne.useState(qk(props.range)),
    [headers, setHeaders] = ne.useState(true),
    [comparison, setComparison] = ne.useState("raw"),
    [caseSensitive, setCaseSensitive] = ne.useState(true),
    [trimText, setTrimText] = ne.useState(false),
    [blankPolicy, setBlankPolicy] = ne.useState("keep"),
    [outputSheetName, setOutputSheetName] = ne.useState("Unique review"),
    [checked, setChecked] = ne.useState(false),
    [preview, setPreview] = ne.useState(null),
    [previewSourceKey, setPreviewSourceKey] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [failedReviewKey, setFailedReviewKey] = ne.useState(""),
    [rowFilter, setRowFilter] = ne.useState("all"),
    [pageIndex, setPageIndex] = ne.useState(0),
    selectableColumns = qk(range),
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    allowedSheetIds = [
      ...new Set(
        props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    sheetContext = {
      includeHidden: !!props.allowHidden && checked,
      allowedSheetIds: props.allowedSheetIds,
    },
    isSourceInContext =
      !!sheet &&
      allowedSheetIds.includes(sheet.id) &&
      (sheetContext.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    sourceKey = JSON.stringify([props.sheetId, props.range]),
    X$1 = JSON.stringify([
      props.workbook,
      props.readOnly,
      allowedSheetIds,
      sheetContext.includeHidden,
      sourceKey,
    ]),
    isPreviewCurrent =
      !!preview &&
      !props.readOnly &&
      preview.snapshot === JSON.stringify(props.workbook) &&
      previewSourceKey === sourceKey &&
      preview.context.includeHidden === sheetContext.includeHidden &&
      JSON.stringify(preview.context.allowedSheetIds) ===
        JSON.stringify(allowedSheetIds),
    resetPreview = () => {
      (setPreview(null), setErrorMessage(""), setPageIndex(0));
    },
    reportError = (error) => {
      (setFailedReviewKey(X$1),
        setErrorMessage(
          error instanceof Error
            ? error.message
            : "Duplicate review could not complete the operation.",
        ));
    };
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current
        ?.querySelector('[data-ui-label="Source range"]')
        ?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const onClick = () => {
      try {
        const duplicateReviewSheet = createDuplicateReviewSheet(
          props.workbook,
          {
            sheetId: props.sheetId,
            range: range,
            keyColumns: keyColumns,
            headers: headers,
            comparison: comparison,
            caseSensitive: caseSensitive,
            trimText: trimText,
            blankPolicy: blankPolicy,
            outputSheetName: outputSheetName,
          },
          { ...sheetContext, readOnly: props.readOnly },
        );
        (setPreview(duplicateReviewSheet),
          setPreviewSourceKey(sourceKey),
          setErrorMessage(""),
          setPageIndex(0),
          setRowFilter("all"));
      } catch (error) {
        (setPreview(null), reportError(error));
      }
    },
    applyPreview = () => {
      if (preview)
        try {
          if (!isPreviewCurrent)
            throw Error(
              "The source selection or preview changed. Preview and review again.",
            );
          const updatedWorkbook = applyDuplicateReviewPreview(
            props.workbook,
            preview,
            {
              ...sheetContext,
              readOnly: props.readOnly,
            },
          );
          (props.onWorkbookChange(
            updatedWorkbook,
            `Create duplicate-reviewed population: ${preview.counts.retained} retained, ${preview.counts.excluded} excluded`,
          ),
            setPreview(null),
            setErrorMessage(""));
        } catch (error) {
          reportError(error);
        }
    },
    filteredRows =
      preview?.rows.filter(
        (row) =>
          rowFilter === "all" ||
          (rowFilter === "groups"
            ? !!row.groupId
            : row.disposition === rowFilter),
      ) ?? [],
    pageCount = Math.max(
      1,
      Math.ceil(filteredRows.length / structuralEditLimits.pageSize),
    ),
    currentPage = Math.min(pageIndex, pageCount - 1),
    pageRows = filteredRows.slice(currentPage * 20, currentPage * 20 + 20),
    renderCellDetails = (cell, calculatedValue) =>
      i.jsxs("details", {
        children: [
          i.jsx("summary", {
            children: describeCellValueType(calculatedValue),
          }),
          i.jsxs("dl", {
            children: [
              i.jsx("dt", {
                children: i.jsx(LocalizedText, { children: "Stored value" }),
              }),
              i.jsx("dd", { children: describeCellValueType(cell.raw) }),
              i.jsx("dt", {
                children: i.jsx(LocalizedText, {
                  children: "Stored as literal",
                }),
              }),
              i.jsx("dd", { children: cell.literal ? "Yes" : "No" }),
              i.jsx("dt", {
                children: i.jsx(LocalizedText, { children: "Note" }),
              }),
              i.jsx("dd", { children: cell.note ?? "None" }),
              i.jsx("dt", {
                children: i.jsx(LocalizedText, { children: "Tickmark" }),
              }),
              i.jsx("dd", { children: cell.tickmark ?? "None" }),
              i.jsx("dt", {
                children: i.jsx(LocalizedText, {
                  children: "Cell formatting / link",
                }),
              }),
              i.jsx("dd", {
                children: i.jsx("code", {
                  children: JSON.stringify(
                    Object.fromEntries(
                      Object.entries(cell).filter(
                        ([key]) =>
                          !["raw", "literal", "note", "tickmark"].includes(key),
                      ),
                    ),
                  ),
                }),
              }),
            ],
          }),
        ],
      });
  return i.jsx("div", {
    className: "dup-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      className: "dup-dialog",
      ref: dialogRef,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(), onCloseRef.current()),
          event.key === "Tab")
        ) {
          const focusableElements = [
            ...(dialogRef.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),summary",
            ) ?? []),
          ].filter(
            (element) =>
              !element.closest("details:not([open])") ||
              element.tagName === "SUMMARY",
          );
          event.shiftKey && document.activeElement === focusableElements[0]
            ? (event.preventDefault(),
              (lastFocusable = focusableElements.at(-1)) == null ||
                lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === focusableElements.at(-1) &&
              (event.preventDefault(),
              (firstFocusable = focusableElements[0]) == null ||
                firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${dialogId}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Remove duplicates",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Create a reviewed unique population on a new worksheet.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close duplicate review"),
              onClick: props.onClose,
              lang: language("Close duplicate review"),
              "data-ui-label": "Close duplicate review",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "The original worksheet stays intact. The new worksheet contains retained rows and an exclusion appendix with every later duplicate and its original source locator.",
          }),
        }),
        i.jsxs("div", {
          className: "dup-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source worksheet" }),
                i.jsx("strong", {
                  children: isSourceInContext
                    ? sheet.name
                    : "Outside current context",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Source range"),
                  value: isSourceInContext ? range : "",
                  readOnly: props.readOnly,
                  onChange: (event) => {
                    const value = event.target.value;
                    (setRange(value), setKeyColumns(qk(value)), resetPreview());
                  },
                  lang: language("Source range"),
                  "data-ui-label": "Source range",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "New worksheet name" }),
                i.jsx("input", {
                  "aria-label": text("New worksheet name"),
                  value: outputSheetName,
                  readOnly: props.readOnly,
                  onChange: (event) => {
                    (setOutputSheetName(event.target.value), resetPreview());
                  },
                  lang: language("New worksheet name"),
                  "data-ui-label": "New worksheet name",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Compare" }),
                i.jsxs("select", {
                  "aria-label": text("Comparison basis"),
                  value: comparison,
                  disabled: props.readOnly,
                  onChange: (event) => {
                    (setComparison(event.target.value), resetPreview());
                  },
                  lang: language("Comparison basis"),
                  "data-ui-label": "Comparison basis",
                  children: [
                    i.jsx("option", {
                      value: "raw",
                      children: "Typed raw values / exact formula text",
                    }),
                    i.jsx("option", {
                      value: "calculated",
                      children: "Typed calculated scalar values",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Blank key policy" }),
                i.jsxs("select", {
                  "aria-label": text("Blank key policy"),
                  value: blankPolicy,
                  disabled: props.readOnly,
                  onChange: (event) => {
                    (setBlankPolicy(event.target.value), resetPreview());
                  },
                  lang: language("Blank key policy"),
                  "data-ui-label": "Blank key policy",
                  children: [
                    i.jsx("option", {
                      value: "keep",
                      children: "Keep every row with any blank key",
                    }),
                    i.jsx("option", {
                      value: "match",
                      children: "Match blanks together",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              className: "dup-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: headers,
                  disabled: props.readOnly,
                  onChange: (event) => {
                    (setHeaders(event.target.checked), resetPreview());
                  },
                }),
                i.jsx(LocalizedText, {
                  children: "First row contains headers",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "dup-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: caseSensitive,
                  disabled: props.readOnly,
                  onChange: (event) => {
                    (setCaseSensitive(event.target.checked), resetPreview());
                  },
                }),
                i.jsx(LocalizedText, { children: "Case-sensitive text" }),
              ],
            }),
            i.jsxs("label", {
              className: "dup-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: trimText,
                  disabled: props.readOnly,
                  onChange: (event) => {
                    (setTrimText(event.target.checked), resetPreview());
                  },
                }),
                i.jsx(LocalizedText, {
                  children: "Trim outer text whitespace",
                }),
              ],
            }),
            props.allowHidden &&
              i.jsxs("label", {
                className: "dup-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => {
                      (setChecked(event.target.checked), resetPreview());
                    },
                  }),
                  i.jsx(LocalizedText, { children: "Include hidden sources" }),
                ],
              }),
          ],
        }),
        i.jsxs("fieldset", {
          disabled: props.readOnly || !isSourceInContext,
          children: [
            i.jsx("legend", {
              children: i.jsx(LocalizedText, {
                children: "Key columns — all selected keys must match",
              }),
            }),
            i.jsx("div", {
              className: "dup-keys",
              children:
                isSourceInContext &&
                selectableColumns.map((columnIndex) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          "aria-label": `Key column ${columnIndexToLetters(columnIndex)}`,
                          checked: keyColumns.includes(columnIndex),
                          onChange: (event) => {
                            (setKeyColumns(
                              event.target.checked
                                ? [...keyColumns, columnIndex]
                                : keyColumns.filter(
                                    (keyColumn) => keyColumn !== columnIndex,
                                  ),
                            ),
                              resetPreview());
                          },
                        }),
                        columnIndexToLetters(columnIndex),
                      ],
                    },
                    columnIndex,
                  ),
                ),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "First source row wins. Numbers, text and Booleans remain distinct. Matching rules affect keys only; exported values preserve their original text. Raw formulas compare exactly and become literal text in the derived snapshot. Calculated mode freezes supported scalar results; cell display formats do not determine equality.",
          }),
        }),
        errorMessage &&
          failedReviewKey === X$1 &&
          i.jsx("p", {
            className: "dup-error",
            role: "alert",
            children: errorMessage,
          }),
        i.jsx("button", {
          className: "dup-primary",
          disabled: props.readOnly || !isSourceInContext || !keyColumns.length,
          onClick: onClick,
          children: i.jsx(LocalizedText, {
            children: "Preview duplicate review",
          }),
        }),
        preview &&
          i.jsxs("section", {
            className: "dup-preview",
            "aria-label": text("Duplicate review preview"),
            lang: language("Duplicate review preview"),
            "data-ui-label": "Duplicate review preview",
            children: [
              isPreviewCurrent
                ? i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("h3", {
                        children: [
                          preview.counts.input,
                          " ",
                          i.jsx(LocalizedText, { children: "source rows · " }),
                          preview.counts.retained,
                          " ",
                          i.jsx(LocalizedText, { children: "retained · " }),
                          preview.counts.excluded,
                          " ",
                          i.jsx(LocalizedText, { children: "excluded" }),
                        ],
                      }),
                      i.jsxs("p", {
                        role: "status",
                        children: [
                          preview.counts.duplicateGroups,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "duplicate groups;",
                          }),
                          " ",
                          preview.counts.blankKeyRows,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "rows with a blank key. New worksheet: ",
                          }),
                          preview.request.outputSheetName,
                          i.jsx(LocalizedText, {
                            children: ". Unique population:",
                          }),
                          " ",
                          preview.populationRange,
                          ".",
                          " ",
                          preview.excludedRange
                            ? `Exclusion appendix: ${preview.excludedRange}.`
                            : "No duplicate rows were found; the derived population still includes source locators.",
                        ],
                      }),
                      preview.request.headers &&
                        i.jsxs("details", {
                          children: [
                            i.jsx("summary", {
                              children: i.jsx(LocalizedText, {
                                children: "Source header evidence",
                              }),
                            }),
                            i.jsx("div", {
                              className: "dup-headers",
                              children: preview.headerCells.map(
                                (headerCell, headerIndex) =>
                                  i.jsxs(
                                    "div",
                                    {
                                      children: [
                                        i.jsx("strong", {
                                          children: columnIndexToLetters(
                                            preview.request.keyColumns.length
                                              ? parseRangeAddress(
                                                  preview.request.range,
                                                ).firstCol + headerIndex
                                              : headerIndex,
                                          ),
                                        }),
                                        renderCellDetails(
                                          headerCell,
                                          headerCell.raw,
                                        ),
                                      ],
                                    },
                                    headerIndex,
                                  ),
                              ),
                            }),
                          ],
                        }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Preview view" }),
                          i.jsxs("select", {
                            "aria-label": text("Preview view"),
                            value: rowFilter,
                            onChange: (event) => {
                              (setRowFilter(event.target.value),
                                setPageIndex(0));
                            },
                            lang: language("Preview view"),
                            "data-ui-label": "Preview view",
                            children: [
                              i.jsx("option", {
                                value: "all",
                                children: "Before — every source row",
                              }),
                              i.jsx("option", {
                                value: "retained",
                                children: "After — retained population",
                              }),
                              i.jsx("option", {
                                value: "excluded",
                                children: "Excluded — every later duplicate",
                              }),
                              i.jsx("option", {
                                value: "groups",
                                children: "Duplicate groups — every member",
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        className: "dup-scroll",
                        children: [
                          i.jsxs("table", {
                            children: [
                              i.jsx("thead", {
                                children: i.jsxs("tr", {
                                  children: [
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Source row / range",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Decision / derived row",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Group / first row / members",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Normalized typed key",
                                      }),
                                    }),
                                    Array.from(
                                      {
                                        length:
                                          parseRangeAddress(
                                            preview.request.range,
                                          ).lastCol -
                                          parseRangeAddress(
                                            preview.request.range,
                                          ).firstCol +
                                          1,
                                      },
                                      (le, columnOffset) =>
                                        i.jsxs(
                                          "th",
                                          {
                                            children: [
                                              i.jsx(LocalizedText, {
                                                children: "Column",
                                              }),
                                              " ",
                                              columnIndexToLetters(
                                                parseRangeAddress(
                                                  preview.request.range,
                                                ).firstCol + columnOffset,
                                              ),
                                            ],
                                          },
                                          columnOffset,
                                        ),
                                    ),
                                  ],
                                }),
                              }),
                              i.jsx("tbody", {
                                children: pageRows.map((previewRow) =>
                                  i.jsxs(
                                    "tr",
                                    {
                                      children: [
                                        i.jsxs("th", {
                                          scope: "row",
                                          children: [
                                            previewRow.sourceRow,
                                            " · ",
                                            previewRow.sourceRange,
                                          ],
                                        }),
                                        i.jsxs("td", {
                                          children: [
                                            previewRow.disposition,
                                            " · ",
                                            previewRow.outputRow,
                                          ],
                                        }),
                                        i.jsxs("td", {
                                          children: [
                                            previewRow.groupId || "Unique",
                                            " · ",
                                            previewRow.firstSourceRow,
                                            " ·",
                                            " ",
                                            previewRow.groupSize,
                                          ],
                                        }),
                                        i.jsxs("td", {
                                          children: [
                                            i.jsx("code", {
                                              children:
                                                previewRow.normalizedKey,
                                            }),
                                            previewRow.blankKey &&
                                              i.jsxs("span", {
                                                children: [
                                                  " ",
                                                  i.jsx(LocalizedText, {
                                                    children: "· Blank key",
                                                  }),
                                                ],
                                              }),
                                          ],
                                        }),
                                        previewRow.cells.map(
                                          (cell, cellIndex) =>
                                            i.jsx(
                                              "td",
                                              {
                                                children: renderCellDetails(
                                                  cell,
                                                  previewRow.values[cellIndex],
                                                ),
                                              },
                                              cellIndex,
                                            ),
                                        ),
                                      ],
                                    },
                                    previewRow.sourceRow,
                                  ),
                                ),
                              }),
                            ],
                          }),
                          !filteredRows.length &&
                            i.jsx("p", {
                              children: i.jsx(LocalizedText, {
                                children: "No rows in this view.",
                              }),
                            }),
                        ],
                      }),
                      i.jsxs("nav", {
                        "aria-label": text("Duplicate preview pages"),
                        lang: language("Duplicate preview pages"),
                        "data-ui-label": "Duplicate preview pages",
                        children: [
                          i.jsx("button", {
                            disabled: currentPage === 0,
                            onClick: () => setPageIndex(currentPage - 1),
                            children: i.jsx(LocalizedText, {
                              children: "Previous page",
                            }),
                          }),
                          i.jsxs("span", {
                            role: "status",
                            children: [
                              i.jsx(LocalizedText, { children: "Page " }),
                              currentPage + 1,
                              " ",
                              i.jsx(LocalizedText, { children: "of " }),
                              pageCount,
                              " · ",
                              filteredRows.length,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "rows in this view",
                              }),
                            ],
                          }),
                          i.jsx("button", {
                            disabled: currentPage + 1 >= pageCount,
                            onClick: () => setPageIndex(currentPage + 1),
                            children: i.jsx(LocalizedText, {
                              children: "Next page",
                            }),
                          }),
                        ],
                      }),
                    ],
                  })
                : i.jsx("p", {
                    role: "status",
                    children: i.jsx(LocalizedText, {
                      children:
                        "This preview is stale. Source locations, values and groups are withheld. Preview again in the current workbook and access context.",
                    }),
                  }),
              i.jsxs("div", {
                className: "dup-actions",
                children: [
                  i.jsx("button", {
                    onClick: () => setPreview(null),
                    children: i.jsx(LocalizedText, {
                      children: "Discard preview",
                    }),
                  }),
                  i.jsx("button", {
                    className: "dup-primary",
                    disabled: props.readOnly || !isPreviewCurrent,
                    onClick: applyPreview,
                    children: i.jsx(LocalizedText, {
                      children: "Create reviewed population",
                    }),
                  }),
                ],
              }),
            ],
          }),
        i.jsxs("details", {
          children: [
            i.jsx("summary", {
              children: i.jsx(LocalizedText, {
                children: "Scope and evidence limits",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Maximum 500 selected cells. Calculated comparison requires a complete static scalar dependency closure within 1,000 cells, 5,000 scanned formulas and depth 8. Volatile, dynamic, array/spill and unresolved calculations are excluded. Formula work limits cannot be hidden by error handlers. Calculated sources and dependencies must always be populated to avoid mistaking runtime spill children for blanks. Use raw comparison for null or absent source cells.",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Null and empty normalized text are blanks. Case folding uses Unicode lowercase; trimming removes outer whitespace only. Formatting, locale-specific collation and numerical display rounding do not merge distinct values. Sheet objects, table styles, conditional rules and discussion threads remain on the original source; direct cell metadata is copied. Source locators describe this snapshot and are not immutable lineage or live links. A protected source may be read, but read-only hosts and workbook structure protection block creation.",
              }),
            }),
          ],
        }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: props.onClose,
            children: i.jsx(LocalizedText, { children: "Close" }),
          }),
        }),
      ],
    }),
  });
}
const structurePreviewRegistry = new WeakMap(),
  containsRefError = (formulaText) =>
    /#REF!/i.test(formulaText.replace(/"(?:[^"]|"")*"/g, '""')),
  formatAxisAddress = (edit, axisIndex) =>
    edit.axis === "row"
      ? String(axisIndex + 1)
      : columnIndexToLetters(axisIndex);
function buildStructurePreview(workbook, sheetId, edit, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  const workbookWithSettings = calculateWorkbookWithSettings(workbook),
    editedWorkbook = applyStructuralEdit(workbook, sheetId, edit),
    editedValueAt = calculateWorkbookWithSettings(editedWorkbook);
  if (
    workbook.sheets.reduce(
      (total, worksheet) =>
        total +
        Object.values(worksheet.cells).filter(isFormulaCellValue).length,
      0,
    ) > 5e3
  )
    throw Error(
      "Structural preview can verify up to 5,000 formula cells. Reduce or separate the formula population before changing worksheet structure.",
    );
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    action =
      edit.moveBefore !== undefined
        ? `Move ${edit.count} ${edit.axis}${edit.count === 1 ? "" : "s"} before ${edit.moveBefore === (edit.axis === "row" ? sheet.rowCount : sheet.colCount) ? "worksheet end" : formatAxisAddress(edit, edit.moveBefore)}`
        : `${edit.remove ? "Delete" : "Insert"} ${edit.count} ${edit.axis}${edit.count === 1 ? "" : "s"}`,
    preview = {
      sheetId: sheetId,
      edit: { ...edit },
      snapshot: JSON.stringify(workbook),
      action: action,
      deletedCells: 0,
      movedCells: 0,
      formulaChanges: [],
      referenceIssues: [],
      originChanges: [],
      metadataChanges: [],
    };
  for (const sheet of workbook.sheets) {
    const editedSheet = editedWorkbook.sheets.find(
      (entry) => entry.id === sheet.id,
    );
    for (const [beforeAddress, cell] of Object.entries(sheet.cells)) {
      const address =
        sheet.id === sheetId
          ? assertInsertionWithinSheet(beforeAddress, edit)
          : beforeAddress;
      if (!address) {
        preview.deletedCells++;
        continue;
      }
      address !== beforeAddress && preview.movedCells++;
      const editedCell = editedSheet.cells[address];
      isFormulaCellValue(cell) &&
        (editedCell?.raw !== cell.raw &&
          preview.formulaChanges.push({
            sheetId: sheet.id,
            sheetName: sheet.name,
            beforeAddress: beforeAddress,
            address: address,
            before: String(cell.raw),
            after: String(editedCell?.raw ?? ""),
          }),
        editedValueAt(sheet.id, address) === "#REF!" &&
          workbookWithSettings(sheet.id, beforeAddress) !== "#REF!" &&
          preview.referenceIssues.push({
            kind: "cell",
            sheetName: sheet.name,
            address: address,
            before: String(cell.raw),
            after: String(editedCell?.raw ?? ""),
          }));
    }
    for (const editedRule of editedSheet.conditionalFormats ?? []) {
      const beforeRule = sheet.conditionalFormats?.find(
        (conditionalFormat) => conditionalFormat.id === editedRule.id,
      );
      editedRule.formula &&
        containsRefError(editedRule.formula) &&
        !containsRefError(beforeRule?.formula ?? "") &&
        preview.referenceIssues.push({
          kind: "conditional rule",
          sheetName: editedSheet.name,
          address: editedRule.range,
          before: beforeRule?.formula ?? "",
          after: editedRule.formula,
        });
    }
    for (const field of [
      "rowHeights",
      "columnWidths",
      "hiddenRows",
      "hiddenCols",
      "frozenRows",
      "frozenCols",
      "merges",
      "validations",
      "conditionalFormats",
      "tables",
      "pivots",
      "queries",
      "objects",
      "pageLayout",
      "outlineGroups",
    ])
      JSON.stringify(sheet[field]) !== JSON.stringify(editedSheet[field]) &&
        preview.metadataChanges.push({ sheetName: sheet.name, field: field });
  }
  for (const editedName of editedWorkbook.definedNames ?? []) {
    const beforeName = workbook.definedNames?.find(
      (definedName) =>
        definedName.name === editedName.name &&
        definedName.scopeSheetId === editedName.scopeSheetId,
    );
    containsRefError(editedName.refersTo) &&
      !containsRefError(beforeName?.refersTo ?? "") &&
      preview.referenceIssues.push({
        kind: "name",
        sheetName:
          workbook.sheets.find((sheet) => sheet.id === editedName.scopeSheetId)
            ?.name ?? "Workbook",
        address: editedName.name,
        before: beforeName?.refersTo ?? "",
        after: editedName.refersTo,
      });
  }
  const rowOriginsById = new Map(
    (workbook.rowOrigins ?? []).map((rowOrigin) => [rowOrigin.id, rowOrigin]),
  );
  if (edit.axis === "row")
    for (const [rowKey, originId] of Object.entries(sheet.rowOriginIds ?? {})) {
      const shiftedAddress = assertInsertionWithinSheet(
          `A${Number(rowKey) + 1}`,
          edit,
        ),
        afterRow = shiftedAddress ? _t(shiftedAddress).row + 1 : null;
      if (afterRow === Number(rowKey) + 1) continue;
      const origin = rowOriginsById.get(originId);
      origin &&
        preview.originChanges.push({
          id: originId,
          sourceName: origin.sourceName,
          sourceSheet: origin.sourceSheet,
          originalRow: origin.originalRow,
          beforeRow: Number(rowKey) + 1,
          afterRow: afterRow,
        });
    }
  return (
    structurePreviewRegistry.set(preview, JSON.stringify(preview)),
    preview
  );
}
function assertStructurePreviewCurrent(workbook, preview, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  if (structurePreviewRegistry.get(preview) !== JSON.stringify(preview))
    throw Error(
      "This structural preview was changed or restored from another session. Create a fresh preview before applying.",
    );
  if (JSON.stringify(workbook) !== preview.snapshot)
    throw Error(
      "The workbook changed after this preview. Refresh the preview before applying.",
    );
  return applyStructuralEdit(workbook, preview.sheetId, preview.edit);
}
function StructureToolDialog({
  workbook,
  sheetId,
  range,
  mode,
  readOnly,
  onClose,
  onWorkbookChange,
}) {
  const { text, language } = useWorkbookMessages(),
    axis = mode === "insertRow" ? "row" : "column",
    rangeAddress = parseRangeAddress(range),
    axisLimit =
      axis === "row" ? workbookLimits.maxRows : workbookLimits.maxCols,
    [startValue, setStartValue] = ne.useState(
      axis === "row"
        ? String(rangeAddress.firstRow + 1)
        : columnIndexToLetters(rangeAddress.firstCol),
    ),
    [count, setCount] = ne.useState(1),
    [operation, setOperation] = ne.useState("insert"),
    [moveDestination, setMoveDestination] = ne.useState(""),
    [preview, setPreview] = ne.useState(null),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [pageIndex, setPageIndex] = ne.useState(0),
    [reviewCategory, setReviewCategory] = ne.useState("formulas"),
    [checked, setChecked] = ne.useState(false),
    dialogRef = ne.useRef(null),
    sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("select")?.focus(),
      () => activeElement?.focus()
    );
  }, []);
  const resetPreview = () => {
      (setPreview(null),
        setChecked(false),
        setErrorMessage(""),
        setPageIndex(0));
    },
    parseIndexInput = (inputText, allowEnd = false) => {
      if (allowEnd && inputText.trim().toLowerCase() === "end")
        return axisLimit;
      if (axis === "row") {
        const rowNumber = Number(inputText);
        if (
          !/^[1-9]\d*$/.test(inputText.trim()) ||
          !Number.isInteger(rowNumber) ||
          rowNumber > axisLimit
        )
          throw Error(
            "Enter an existing row number, or use end for the move destination.",
          );
        return rowNumber - 1;
      }
      if (!/^[a-z]{1,3}$/i.test(inputText.trim()))
        throw Error(
          "Enter a column letter such as A or XFD, or use end for the move destination.",
        );
      return parseRangeAddress(inputText.trim().toUpperCase() + "1").firstCol;
    };
  function onClick() {
    try {
      setErrorMessage("");
      const structurePreview = buildStructurePreview(
        workbook,
        sheetId,
        {
          axis: axis,
          index: parseIndexInput(startValue),
          count: count,
          remove: operation === "delete",
          ...(operation === "move"
            ? { moveBefore: parseIndexInput(moveDestination, true) }
            : {}),
        },
        { readOnly },
      );
      (setPreview(structurePreview), setChecked(false), setPageIndex(0));
    } catch (error) {
      (setErrorMessage(error.message), setPreview(null));
    }
  }
  function applyPreview() {
    try {
      if (!preview) throw Error("Create a preview first.");
      if (requiresConfirmation && !checked)
        throw Error(
          "Confirm the deletion and reference changes before applying.",
        );
      const nextWorkbook = assertStructurePreviewCurrent(workbook, preview, {
        readOnly,
      });
      (onWorkbookChange(
        nextWorkbook,
        `${preview.action}${preview.referenceIssues.length ? `; ${preview.referenceIssues.length} new reference error(s) — reviewed` : ""}`,
      ),
        onClose());
    } catch (error) {
      setErrorMessage(error.message);
    }
  }
  const requiresConfirmation =
      !!preview &&
      (preview.deletedCells > 0 ||
        preview.referenceIssues.length > 0 ||
        preview.originChanges.some(
          (originChange) => originChange.afterRow === null,
        )),
    reviewItems = preview
      ? reviewCategory === "errors"
        ? preview.referenceIssues.map((referenceIssue) => ({
            key: `${referenceIssue.kind}-${referenceIssue.sheetName}-${referenceIssue.address}`,
            title: `${referenceIssue.kind}: ${referenceIssue.sheetName}!${referenceIssue.address}`,
            detail: `${referenceIssue.before} → ${referenceIssue.after}`,
          }))
        : reviewCategory === "origins"
          ? preview.originChanges.map((originChange) => ({
              key: originChange.id + "-" + originChange.beforeRow,
              title: `Row ${originChange.beforeRow} → ${originChange.afterRow ?? "deleted"}`,
              detail: `${originChange.sourceName} · ${originChange.sourceSheet}, original row ${originChange.originalRow} · ID ${originChange.id}`,
            }))
          : preview.formulaChanges.map((formulaChange) => ({
              key: formulaChange.sheetId + formulaChange.beforeAddress,
              title: `${formulaChange.sheetName}!${formulaChange.beforeAddress} → ${formulaChange.address}`,
              detail: `${formulaChange.before} → ${formulaChange.after}`,
            }))
      : [];
  let rowOrigin = null;
  try {
    rowOrigin = findRowOrigin(
      workbook,
      sheetId,
      axis === "row" ? parseIndexInput(startValue) : rangeAddress.firstRow,
    );
  } catch {}
  return i.jsx("div", {
    className: "st-backdrop",
    onPointerDown: (event) => {
      event.target === event.currentTarget && onClose();
    },
    children: i.jsxs("section", {
      className: "st-dialog",
      ref: dialogRef,
      role: "dialog",
      "aria-modal": "true",
      "aria-label": `Insert / Delete / Move ${axis === "row" ? "Rows" : "Columns"}`,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.stopPropagation(), onClose()),
          event.key === "Tab")
        ) {
          const focusable = Array.from(
              dialogRef.current?.querySelectorAll(
                "button:not([disabled]),input:not([disabled]),select:not([disabled])",
              ) ?? [],
            ),
            firstFocusable = focusable[0],
            lastFocusable = focusable.at(-1);
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
          className: "st-header",
          children: [
            i.jsxs("h2", {
              children: [
                axis === "row" ? "Rows" : "Columns",
                " · ",
                sheet.name,
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close structure tool"),
              onClick: onClose,
              lang: language("Close structure tool"),
              "data-ui-label": "Close structure tool",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "st-body",
          children: [
            i.jsxs("div", {
              className: "st-controls",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Operation" }),
                    i.jsxs("select", {
                      "aria-label": text("Structure operation"),
                      value: operation,
                      onChange: (event) => {
                        (resetPreview(), setOperation(event.target.value));
                      },
                      lang: language("Structure operation"),
                      "data-ui-label": "Structure operation",
                      children: [
                        i.jsx("option", {
                          value: "insert",
                          children: "Insert",
                        }),
                        i.jsx("option", {
                          value: "delete",
                          children: "Delete",
                        }),
                        i.jsx("option", {
                          value: "move",
                          children: "Move before",
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    axis === "row" ? "First row" : "First column",
                    i.jsx("input", {
                      "aria-label": text("Structure start"),
                      value: startValue,
                      onChange: (event) => {
                        (resetPreview(), setStartValue(event.target.value));
                      },
                      lang: language("Structure start"),
                      "data-ui-label": "Structure start",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Count" }),
                    i.jsx(LocalizedNumberField, {
                      "aria-label": text("Structure count"),
                      min: 1,
                      max: 1e4,
                      value: count,
                      onValueChange: (nextValue) => {
                        (resetPreview(), setCount(Number(nextValue)));
                      },
                      lang: language("Structure count"),
                      "data-ui-label": "Structure count",
                    }),
                  ],
                }),
              ],
            }),
            operation === "move" &&
              i.jsxs("label", {
                className: "st-destination",
                children: [
                  i.jsx(LocalizedText, { children: "Before " }),
                  axis,
                  " ",
                  i.jsx(LocalizedText, { children: "(original position)" }),
                  i.jsx("input", {
                    "aria-label": text("Structure destination"),
                    value: moveDestination,
                    placeholder: axis === "row" ? "10, or end" : "F, or end",
                    onChange: (event) => {
                      (resetPreview(), setMoveDestination(event.target.value));
                    },
                    lang: language("Structure destination"),
                    "data-ui-label": "Structure destination",
                  }),
                ],
              }),
            i.jsxs("p", {
              children: [
                i.jsx(LocalizedText, { children: "Entire " }),
                axis === "row" ? "rows" : "columns",
                i.jsx(LocalizedText, {
                  children:
                    ", including hidden cells, move together. A move inserts the selection before the destination and shifts the intervening data.",
                }),
              ],
            }),
            i.jsxs("div", {
              className: "st-source",
              children: [
                i.jsx("strong", {
                  children: i.jsx(LocalizedText, { children: "Source row" }),
                }),
                i.jsx("p", {
                  children: rowOrigin
                    ? `${rowOrigin.sourceName} · ${rowOrigin.sourceSheet}, original row ${rowOrigin.originalRow} · ID ${rowOrigin.id}`
                    : "No source identity is recorded for this row. Inserted and editor-created rows do not acquire a fabricated file origin.",
                }),
              ],
            }),
            errorMessage &&
              i.jsx("p", {
                role: "alert",
                className: "st-error",
                children: errorMessage,
              }),
            preview &&
              i.jsxs("section", {
                "aria-label": text("Structure preview"),
                lang: language("Structure preview"),
                "data-ui-label": "Structure preview",
                children: [
                  i.jsx("h3", { children: preview.action }),
                  i.jsxs("p", {
                    children: [
                      preview.movedCells,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "stored cells relocate · ",
                      }),
                      preview.deletedCells,
                      " ",
                      i.jsx(LocalizedText, { children: "cells deleted · " }),
                      preview.formulaChanges.length,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "formulas rewritten ·",
                      }),
                      " ",
                      i.jsxs("strong", {
                        children: [
                          preview.referenceIssues.length,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "new reference errors",
                          }),
                        ],
                      }),
                      " ",
                      "· ",
                      preview.originChanges.length,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "source-row associations change.",
                      }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, {
                        children:
                          "Original source IDs and original row numbers stay fixed. Deleted-row origins remain in the native source registry.",
                      }),
                      " ",
                      preview.metadataChanges.length,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "worksheet settings update.",
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Review" }),
                      i.jsxs("select", {
                        "aria-label": text("Structure review category"),
                        value: reviewCategory,
                        onChange: (event) => {
                          (setReviewCategory(event.target.value),
                            setPageIndex(0));
                        },
                        lang: language("Structure review category"),
                        "data-ui-label": "Structure review category",
                        children: [
                          i.jsxs("option", {
                            value: "formulas",
                            children: [
                              "Rewritten formulas (",
                              preview.formulaChanges.length,
                              ")",
                            ],
                          }),
                          i.jsxs("option", {
                            value: "errors",
                            children: [
                              "New reference errors (",
                              preview.referenceIssues.length,
                              ")",
                            ],
                          }),
                          i.jsxs("option", {
                            value: "origins",
                            children: [
                              "Source rows (",
                              preview.originChanges.length,
                              ")",
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsx("div", {
                    className: "st-changes",
                    children:
                      reviewItems.length === 0
                        ? i.jsx("p", {
                            children: i.jsx(LocalizedText, {
                              children: "No changes in this category.",
                            }),
                          })
                        : reviewItems
                            .slice(pageIndex * 25, pageIndex * 25 + 25)
                            .map((reviewItem) =>
                              i.jsxs(
                                "article",
                                {
                                  children: [
                                    i.jsx("strong", {
                                      children: reviewItem.title,
                                    }),
                                    i.jsx("p", { children: reviewItem.detail }),
                                  ],
                                },
                                reviewItem.key,
                              ),
                            ),
                  }),
                  reviewItems.length > 25 &&
                    i.jsxs("nav", {
                      "aria-label": text("Structure review pages"),
                      lang: language("Structure review pages"),
                      "data-ui-label": "Structure review pages",
                      children: [
                        i.jsx("button", {
                          disabled: pageIndex === 0,
                          onClick: () => setPageIndex((page) => page - 1),
                          children: i.jsx(LocalizedText, {
                            children: "Previous",
                          }),
                        }),
                        i.jsxs("span", {
                          children: [
                            pageIndex * 25 + 1,
                            "–",
                            Math.min(reviewItems.length, (pageIndex + 1) * 25),
                            " ",
                            i.jsx(LocalizedText, { children: "of" }),
                            " ",
                            reviewItems.length,
                          ],
                        }),
                        i.jsx("button", {
                          disabled: (pageIndex + 1) * 25 >= reviewItems.length,
                          onClick: () => setPageIndex((page) => page + 1),
                          children: i.jsx(LocalizedText, { children: "Next" }),
                        }),
                      ],
                    }),
                  preview.metadataChanges.length > 0 &&
                    i.jsxs("details", {
                      children: [
                        i.jsx("summary", {
                          children: i.jsx(LocalizedText, {
                            children: "Changed worksheet settings",
                          }),
                        }),
                        i.jsx("ul", {
                          children: preview.metadataChanges.map(
                            (metadataChange) =>
                              i.jsxs(
                                "li",
                                {
                                  children: [
                                    metadataChange.sheetName,
                                    ":",
                                    " ",
                                    {
                                      rowHeights: "Row heights",
                                      columnWidths: "Column widths",
                                      hiddenRows: "Hidden rows",
                                      hiddenCols: "Hidden columns",
                                      frozenRows: "Frozen rows",
                                      frozenCols: "Frozen columns",
                                      merges: "Merged cells",
                                      validations: "Validation rules",
                                      conditionalFormats: "Conditional formats",
                                      tables: "Tables",
                                      pivots: "PivotTables",
                                      queries: "Saved queries",
                                      objects: "Pictures, charts and shapes",
                                      pageLayout: "Print settings",
                                      outlineGroups: "Outline groups",
                                    }[metadataChange.field],
                                  ],
                                },
                                metadataChange.sheetName + metadataChange.field,
                              ),
                          ),
                        }),
                      ],
                    }),
                ],
              }),
            i.jsx("p", {
              className: "st-hint",
              children: i.jsx(LocalizedText, {
                children:
                  "Up to 10,000 positions and 5,000 formula cells can be reviewed. Split references or structures are refused. Native JSON retains row provenance; XLSX and CSV do not preserve the source registry.",
              }),
            }),
          ],
        }),
        i.jsxs("footer", {
          className: "st-footer",
          children: [
            preview &&
              requiresConfirmation &&
              i.jsxs("label", {
                className: "st-consent",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => setChecked(event.target.checked),
                  }),
                  i.jsx(LocalizedText, {
                    children:
                      "I reviewed the deletions and resulting reference errors.",
                  }),
                ],
              }),
            i.jsxs("div", {
              children: [
                i.jsx("button", {
                  onClick: onClose,
                  children: i.jsx(LocalizedText, { children: "Cancel" }),
                }),
                i.jsx("button", {
                  disabled: readOnly,
                  onClick: onClick,
                  children: preview ? "Refresh preview" : "Preview changes",
                }),
                i.jsx("button", {
                  className: "st-primary",
                  disabled:
                    !preview || readOnly || (requiresConfirmation && !checked),
                  onClick: applyPreview,
                  children: i.jsx(LocalizedText, { children: "Apply changes" }),
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  });
}

export { DuplicateReviewDialog, StructureToolDialog };
