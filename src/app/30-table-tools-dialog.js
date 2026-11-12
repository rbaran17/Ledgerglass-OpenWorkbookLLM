import i from "react/jsx-runtime";
import ne from "react";
import {
  suggestTableHeaderName,
  applyTableOperationPreview,
  pivotAggregateOperations,
  applyTableOperation,
  runQueryDefinition,
  basicAggregateOperations,
  findTableForRange,
  createWorksheetTable,
  refreshPivotTable,
  getPivotSourceFields,
  buildPivotChartModel,
  buildPivotTableOutput,
  refreshQueryDefinition,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  PivotStatusBadge,
  formatAggregationLabel,
  nextDefaultTableName,
  PivotFieldOptionsPanel,
  toggleListValue,
  tryComputePivotRows,
  TableColumnTypesPanel,
} from "./29-worksheet-tools-and-pivot-panels.js";

function TableToolsDialog({
  mode,
  workbook,
  sheetId,
  range,
  readOnly,
  onClose,
  onWorkbookChange,
}) {
  const { text, language } = useWorkbookMessages(),
    dialogRef = ne.useRef(null);
  ne.useEffect(() => {
    const activeElement = document.activeElement,
      dialogElement = dialogRef.current,
      firstFocusable = dialogElement?.querySelector(
        "input:not(:disabled), select:not(:disabled), button:not(:disabled)",
      );
    return (
      firstFocusable?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const initialRange = ne.useMemo(() => {
      try {
        return findTableForRange(workbook, sheetId, range);
      } catch {
        return range;
      }
    }, []),
    initialFields = ne.useMemo(
      () => tryComputePivotRows(workbook, sheetId, initialRange),
      [],
    ),
    [sourceSheetId, setSourceSheetId] = ne.useState(sheetId),
    [sourceRange, setSourceRange] = ne.useState(initialRange),
    [fields, setFields] = ne.useState(initialFields),
    [tableName, setTableName] = ne.useState(
      mode === "table"
        ? nextDefaultTableName(workbook)
        : mode === "pivot"
          ? "Pivot summary"
          : "Extracted data",
    ),
    [tableId, setTableId] = ne.useState(""),
    [tableStyle, setTableStyle] = ne.useState("blue"),
    [showTotals, setShowTotals] = ne.useState(false),
    [totals, setTotals] = ne.useState({}),
    [rowFields, setRowFields] = ne.useState(initialFields.slice(0, 1)),
    [columnFields, setColumnFields] = ne.useState([]),
    [values, setValues] = ne.useState(
      initialFields.length
        ? [{ field: initialFields.at(-1), operation: "sum" }]
        : [],
    ),
    [grandTotals, setGrandTotals] = ne.useState(true),
    [subtotals, setSubtotals] = ne.useState(false),
    [dateGroups, setDateGroups] = ne.useState([]),
    [calculatedFields, setCalculatedFieldList] = ne.useState([]);
  function setCalculatedFields(nextCalculatedFields) {
    (setValues((previousValues) =>
      previousValues.flatMap((valueItem) => {
        const calculatedFieldIndex = calculatedFields.findIndex(
          (calculatedField) => calculatedField.name === valueItem.field,
        );
        return calculatedFieldIndex < 0
          ? [valueItem]
          : nextCalculatedFields.length === calculatedFields.length
            ? [
                {
                  ...valueItem,
                  field: nextCalculatedFields[calculatedFieldIndex].name,
                },
              ]
            : nextCalculatedFields.some(
                  (calculatedField) => calculatedField.name === valueItem.field,
                )
              ? [valueItem]
              : [];
      }),
    ),
      setCalculatedFieldList(nextCalculatedFields));
  }
  const [filterRules, setFilterRules] = ne.useState([]),
    [extractFields, setExtractFields] = ne.useState(initialFields),
    [rename, setRename] = ne.useState({}),
    [field, setField] = ne.useState(""),
    [direction, setDirection] = ne.useState("ascending"),
    [derived, setDerived] = ne.useState([]),
    [checked, setChecked] = ne.useState(false),
    [destinationSheetId, setDestinationSheetId] = ne.useState(sheetId),
    [address, setAddress] = ne.useState("A1"),
    [pivotId, setPivotId] = ne.useState(""),
    [columnId, setColumnId] = ne.useState(""),
    [columnFormula, setColumnFormula] = ne.useState(""),
    [appendCount, setAppendCount] = ne.useState("1"),
    [pendingPreview, setPendingPreview] = ne.useState(null),
    [pageIndex, setPageIndex] = ne.useState(0),
    context = JSON.stringify({
      sheetId,
      range,
      tableId: tableId,
      name: tableName,
      sourceRange: sourceRange,
      style: tableStyle,
      showTotals: showTotals,
      totals: totals,
      columnId: columnId,
      columnFormula: columnFormula,
      appendCount: appendCount,
    }),
    disabled =
      !!pendingPreview &&
      (pendingPreview.plan.snapshot !== JSON.stringify(workbook) ||
        pendingPreview.context !== context),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [statusMessage, setStatusMessage] = ne.useState(""),
    [headerPreview, setHeaderPreview] = ne.useState(null),
    sheetTables =
      workbook.sheets.find((sheet) => sheet.id === sheetId)?.tables ?? [],
    definitionEntries = workbook.sheets.flatMap((sheet) =>
      (mode === "pivot" ? (sheet.pivots ?? []) : (sheet.queries ?? [])).map(
        (definition) => ({
          sheetName: sheet.name,
          definition: definition,
        }),
      ),
    ),
    getDestination = () =>
      checked ? { sheetId: destinationSheetId, address: address } : undefined;
  function runOperation(operation, operationLabel) {
    if (!readOnly) {
      (setErrorMessage(""), setStatusMessage(""));
      try {
        const nextWorkbook = operation();
        (onWorkbookChange(nextWorkbook, operationLabel),
          setStatusMessage(`${operationLabel} completed.`));
      } catch (error) {
        setErrorMessage(
          error instanceof Error
            ? error.message
            : "The operation could not be completed.",
        );
      }
    }
  }
  function onClick() {
    setErrorMessage("");
    try {
      const pivotSourceFields = getPivotSourceFields(
        workbook,
        sourceSheetId,
        sourceRange,
      );
      (setFields(pivotSourceFields),
        setExtractFields(pivotSourceFields),
        setRowFields(pivotSourceFields.slice(0, 1)),
        setColumnFields([]),
        setValues(
          pivotSourceFields.length
            ? [{ field: pivotSourceFields.at(-1), operation: "sum" }]
            : [],
        ),
        setFilterRules([]),
        setRename({}),
        setDerived([]),
        setField(""),
        setDateGroups([]),
        setCalculatedFieldList([]));
    } catch (error) {
      setErrorMessage(error.message);
    }
  }
  function selectExistingTable(selectedTableId) {
    (setHeaderPreview(null),
      setTableId(selectedTableId),
      setPendingPreview(null));
    const selectedTable = sheetTables.find(
      (candidateTable) => candidateTable.id === selectedTableId,
    );
    if (!selectedTable) {
      (setTableName(nextDefaultTableName(workbook)),
        setSourceRange(initialRange),
        setTableStyle("blue"),
        setShowTotals(false),
        setTotals({}));
      return;
    }
    (setTableName(selectedTable.name),
      setColumnId(selectedTable.columns[0]?.id ?? ""),
      setColumnFormula(selectedTable.columns[0]?.calculatedFormula ?? ""),
      setSourceRange(selectedTable.range),
      setFields(selectedTable.columns.map((column) => column.name)),
      setTableStyle(selectedTable.style),
      setShowTotals(selectedTable.showTotals),
      setTotals(selectedTable.totals));
  }
  function buildFilters() {
    return filterRules.map((filterRule) => ({
      ...filterRule,
      value: ["gt", "gte", "lt", "lte"].includes(filterRule.operator)
        ? Number(filterRule.value)
        : filterRule.value,
    }));
  }
  function previewTableOperation(request) {
    (setErrorMessage(""),
      setStatusMessage(""),
      setPendingPreview(null),
      setPageIndex(0));
    try {
      setPendingPreview({
        plan: applyTableOperation(workbook, sheetId, tableId, request, {
          readOnly,
        }),
        context: context,
      });
    } catch (error) {
      setErrorMessage(error.message);
    }
  }
  function applyReviewedOperation() {
    if (pendingPreview) {
      setErrorMessage("");
      try {
        if (disabled)
          throw Error(
            "Table settings or the current selection changed. Preview again.",
          );
        const appliedWorkbook = applyTableOperationPreview(
          workbook,
          pendingPreview.plan,
          { readOnly },
        );
        (appliedWorkbook !== workbook &&
          onWorkbookChange(appliedWorkbook, "Apply reviewed table operation"),
          setSourceRange(pendingPreview.plan.afterRange),
          setPendingPreview(null),
          setStatusMessage(
            appliedWorkbook === workbook
              ? "No table changes were needed."
              : "Reviewed table operation completed.",
          ));
      } catch (error) {
        setErrorMessage(error.message);
      }
    }
  }
  function runPrimaryAction() {
    mode === "table"
      ? tableId
        ? previewTableOperation({
            kind: "updateTable",
            options: {
              name: tableName,
              range: sourceRange,
              style: tableStyle,
              showTotals: showTotals,
              totals: totals,
            },
          })
        : runOperation(
            () =>
              createWorksheetTable(workbook, sheetId, {
                name: tableName,
                range: sourceRange,
                style: tableStyle,
                showTotals: showTotals,
                totals: totals,
                ...(headerPreview
                  ? {
                      headers: (() => {
                        if (
                          headerPreview.snapshot !== JSON.stringify(workbook) ||
                          headerPreview.range !== sourceRange
                        )
                          throw Error(
                            "The workbook or source range changed. Review headers again.",
                          );
                        return headerPreview.headers.map(
                          (header) => header.suggested,
                        );
                      })(),
                    }
                  : {}),
              }),
            "Create table",
          )
      : mode === "pivot"
        ? runOperation(
            () =>
              buildPivotTableOutput(workbook, {
                name: tableName,
                source: { sheetId: sourceSheetId, range: sourceRange },
                rowFields: rowFields,
                columnFields: columnFields,
                values: values,
                filters: buildFilters(),
                grandTotals: grandTotals,
                subtotals: subtotals,
                dateGroups: dateGroups,
                calculatedFields: calculatedFields,
                destination: getDestination(),
              }),
            "Create pivot summary",
          )
        : runOperation(
            () =>
              runQueryDefinition(workbook, {
                name: tableName,
                source: { sheetId: sourceSheetId, range: sourceRange },
                fields: extractFields,
                rename: rename,
                filters: buildFilters(),
                ...(field
                  ? { sort: { field: field, direction: direction } }
                  : {}),
                derived: derived,
                destination: getDestination(),
              }),
            "Extract data",
          );
  }
  const renderFieldOptions = (placeholderLabel) =>
      i.jsxs(i.Fragment, {
        children: [
          placeholderLabel !== undefined &&
            i.jsx("option", { value: "", children: placeholderLabel }),
          fields.map((field) =>
            i.jsx("option", { value: field, children: field }, field),
          ),
        ],
      }),
    renderFieldCheckboxes = (legendLabel, selectedFields, onSelectionChange) =>
      i.jsxs("fieldset", {
        className: "lg-table-fields",
        children: [
          i.jsx("legend", { children: legendLabel }),
          fields.length
            ? fields.map((field) =>
                i.jsxs(
                  "label",
                  {
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: selectedFields.includes(field),
                        onChange: () =>
                          onSelectionChange(
                            toggleListValue(selectedFields, field),
                          ),
                      }),
                      field,
                    ],
                  },
                  field,
                ),
              )
            : i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children: "Read the source fields first.",
                }),
              }),
        ],
      }),
    dialogLabel =
      mode === "table"
        ? "Table tools"
        : mode === "pivot"
          ? "PivotTable builder"
          : "From Table / Range";
  return i.jsxs("div", {
    ref: dialogRef,
    className: "lg-table-tools",
    role: "dialog",
    "aria-modal": "true",
    "aria-label": dialogLabel,
    onKeyDown: (event) => {
      if (
        (event.key === "Escape" &&
          (event.preventDefault(), event.stopPropagation(), onClose()),
        event.key === "Tab")
      ) {
        const focusableElements = Array.from(
            dialogRef.current?.querySelectorAll(
              "button:not(:disabled), input:not(:disabled), select:not(:disabled), summary, [tabindex='0']",
            ) ?? [],
          ).filter(
            (focusableElement) =>
              (!focusableElement.closest("fieldset:disabled") &&
                !focusableElement.closest("details:not([open])")) ||
              focusableElement.tagName === "SUMMARY",
          ),
          firstElement = focusableElements[0],
          lastElement = focusableElements.at(-1);
        event.shiftKey && document.activeElement === firstElement && lastElement
          ? (event.preventDefault(), lastElement.focus())
          : !event.shiftKey &&
            document.activeElement === lastElement &&
            firstElement &&
            (event.preventDefault(), firstElement.focus());
      }
    },
    children: [
      i.jsxs("header", {
        children: [
          i.jsxs("div", {
            children: [
              i.jsx("h2", { children: dialogLabel }),
              i.jsx("p", {
                children:
                  mode === "table"
                    ? "Organize a headed range with a name, style and totals."
                    : mode === "pivot"
                      ? "Group source values into a summary with refreshable settings."
                      : "Select, filter, sort and transform values into a separate result.",
              }),
            ],
          }),
          i.jsx("button", {
            type: "button",
            "aria-label": `Close ${dialogLabel}`,
            onClick: onClose,
            children: "×",
          }),
        ],
      }),
      errorMessage &&
        i.jsx("p", {
          role: "alert",
          className: "lg-table-error",
          children: errorMessage,
        }),
      statusMessage && i.jsx("p", { role: "status", children: statusMessage }),
      readOnly &&
        i.jsx("p", {
          className: "lg-table-readonly",
          children: i.jsx(LocalizedText, {
            children: "This workbook is read-only.",
          }),
        }),
      i.jsxs("fieldset", {
        disabled: readOnly,
        className: "lg-table-form",
        children: [
          mode === "table" &&
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Existing table" }),
                i.jsxs("select", {
                  "aria-label": text("Existing table"),
                  value: tableId,
                  onChange: (event) => selectExistingTable(event.target.value),
                  lang: language("Existing table"),
                  "data-ui-label": "Existing table",
                  children: [
                    i.jsx("option", {
                      value: "",
                      children: "Create a new table",
                    }),
                    sheetTables.map((tableOption) =>
                      i.jsx(
                        "option",
                        { value: tableOption.id, children: tableOption.name },
                        tableOption.id,
                      ),
                    ),
                  ],
                }),
              ],
            }),
          i.jsxs("label", {
            children: [
              mode === "table" ? "Table name" : "Result name",
              i.jsx("input", {
                "aria-label": mode === "table" ? "Table name" : "Result name",
                value: tableName,
                onChange: (event) => setTableName(event.target.value),
              }),
            ],
          }),
          mode !== "table" &&
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source sheet" }),
                i.jsx("select", {
                  value: sourceSheetId,
                  onChange: (event) => setSourceSheetId(event.target.value),
                  children: workbook.sheets.map((sheet) =>
                    i.jsx(
                      "option",
                      { value: sheet.id, children: sheet.name },
                      sheet.id,
                    ),
                  ),
                }),
              ],
            }),
          i.jsxs("div", {
            className: "lg-table-inline",
            children: [
              i.jsxs("label", {
                children: [
                  mode === "table" ? "Table range" : "Source range",
                  i.jsx("input", {
                    "aria-label":
                      mode === "table" ? "Table range" : "Source range",
                    value: sourceRange,
                    onChange: (event) => setSourceRange(event.target.value),
                  }),
                ],
              }),
              i.jsx("button", {
                type: "button",
                onClick: onClick,
                children: i.jsx(LocalizedText, { children: "Read fields" }),
              }),
            ],
          }),
          mode === "table"
            ? i.jsxs(i.Fragment, {
                children: [
                  !tableId &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("button", {
                          type: "button",
                          onClick: () => {
                            setErrorMessage("");
                            try {
                              const headers = suggestTableHeaderName(
                                workbook,
                                sheetId,
                                sourceRange,
                              );
                              (setHeaderPreview({
                                snapshot: JSON.stringify(workbook),
                                range: sourceRange,
                                headers: headers,
                              }),
                                setFields(
                                  headers.map((header) => header.suggested),
                                ));
                            } catch (error) {
                              setErrorMessage(error.message);
                            }
                          },
                          children: i.jsx(LocalizedText, {
                            children: "Review headers",
                          }),
                        }),
                        headerPreview &&
                          i.jsxs("section", {
                            "aria-label": text("Table header resolution"),
                            className: "lg-table-list",
                            lang: language("Table header resolution"),
                            "data-ui-label": "Table header resolution",
                            children: [
                              i.jsx("h3", {
                                children: i.jsx(LocalizedText, {
                                  children: "Resolve table headers",
                                }),
                              }),
                              i.jsx("p", {
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Review each source value and replacement. Create table applies these names and retains changed originals in native source evidence.",
                                }),
                              }),
                              headerPreview.headers.map((header, headerIndex) =>
                                i.jsxs(
                                  "label",
                                  {
                                    children: [
                                      header.address,
                                      " ",
                                      i.jsx(LocalizedText, {
                                        children: "original: ",
                                      }),
                                      JSON.stringify(header.original),
                                      i.jsx("input", {
                                        "aria-label": `Reviewed header ${headerIndex + 1}`,
                                        value: header.suggested,
                                        onChange: (event) => {
                                          const headers =
                                            headerPreview.headers.map(
                                              (header, headerPosition) =>
                                                headerPosition === headerIndex
                                                  ? {
                                                      ...header,
                                                      suggested:
                                                        event.target.value,
                                                    }
                                                  : header,
                                            );
                                          (setHeaderPreview({
                                            ...headerPreview,
                                            headers: headers,
                                          }),
                                            setFields(
                                              headers.map(
                                                (header) => header.suggested,
                                              ),
                                            ));
                                        },
                                      }),
                                    ],
                                  },
                                  header.address,
                                ),
                              ),
                            ],
                          }),
                      ],
                    }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Table style" }),
                      i.jsxs("select", {
                        value: tableStyle,
                        onChange: (event) => setTableStyle(event.target.value),
                        children: [
                          i.jsx("option", { value: "blue", children: "Blue" }),
                          i.jsx("option", {
                            value: "green",
                            children: "Green",
                          }),
                          i.jsx("option", { value: "gray", children: "Gray" }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    className: "lg-table-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: showTotals,
                        onChange: (event) =>
                          setShowTotals(event.target.checked),
                      }),
                      i.jsx(LocalizedText, { children: "Show totals row" }),
                    ],
                  }),
                  showTotals &&
                    i.jsx("div", {
                      className: "lg-table-list",
                      children: fields.map((field) =>
                        i.jsxs(
                          "label",
                          {
                            children: [
                              field,
                              " ",
                              i.jsx(LocalizedText, { children: "total" }),
                              i.jsxs("select", {
                                "aria-label": `${field} total`,
                                value: Object.hasOwn(totals, field)
                                  ? totals[field]
                                  : "",
                                onChange: (event) =>
                                  setTotals((previousTotals) => {
                                    const nextTotals = { ...previousTotals };
                                    return (
                                      event.target.value
                                        ? Object.defineProperty(
                                            nextTotals,
                                            field,
                                            {
                                              value: event.target.value,
                                              enumerable: true,
                                              writable: true,
                                              configurable: true,
                                            },
                                          )
                                        : delete nextTotals[field],
                                      nextTotals
                                    );
                                  }),
                                children: [
                                  i.jsx("option", {
                                    value: "",
                                    children: "None",
                                  }),
                                  basicAggregateOperations.map(
                                    (basicAggregateOperation) =>
                                      i.jsx(
                                        "option",
                                        {
                                          value: basicAggregateOperation,
                                          children: formatAggregationLabel(
                                            basicAggregateOperation,
                                          ),
                                        },
                                        basicAggregateOperation,
                                      ),
                                  ),
                                ],
                              }),
                            ],
                          },
                          field,
                        ),
                      ),
                    }),
                  tableId &&
                    i.jsx(
                      TableColumnTypesPanel,
                      {
                        workbook,
                        sheetId,
                        tableId: tableId,
                        context: context,
                        readOnly,
                        onWorkbookChange,
                      },
                      tableId,
                    ),
                  tableId &&
                    i.jsxs("fieldset", {
                      className: "lg-table-list",
                      children: [
                        i.jsx("legend", {
                          children: i.jsx(LocalizedText, {
                            children: "Calculated columns and rows",
                          }),
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Calculated column",
                            }),
                            i.jsx("select", {
                              "aria-label": text("Calculated column"),
                              value: columnId,
                              onChange: (event) => {
                                (setColumnId(event.target.value),
                                  setColumnFormula(
                                    sheetTables
                                      .find(
                                        (columnTable) =>
                                          columnTable.id === tableId,
                                      )
                                      ?.columns.find(
                                        (column) =>
                                          column.id === event.target.value,
                                      )?.calculatedFormula ?? "",
                                  ),
                                  setPendingPreview(null));
                              },
                              lang: language("Calculated column"),
                              "data-ui-label": "Calculated column",
                              children: sheetTables
                                .find(
                                  (columnOptionTable) =>
                                    columnOptionTable.id === tableId,
                                )
                                ?.columns.map((column) =>
                                  i.jsxs(
                                    "option",
                                    {
                                      value: column.id,
                                      children: [
                                        column.name,
                                        column.calculatedFormula
                                          ? " · calculated"
                                          : "",
                                      ],
                                    },
                                    column.id,
                                  ),
                                ),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "First data-row formula",
                            }),
                            i.jsx("input", {
                              "aria-label": text("First data-row formula"),
                              value: columnFormula,
                              onChange: (event) =>
                                setColumnFormula(event.target.value),
                              placeholder: text("=[@Quantity]*[@Rate]"),
                              lang: language("First data-row formula"),
                              "data-ui-label": "First data-row formula",
                            }),
                          ],
                        }),
                        i.jsx("p", {
                          className: "lg-table-help",
                          children: i.jsx(LocalizedText, {
                            children:
                              "The formula is based on this column’s first data row. Preview lists each overwritten value. Other columns and existing evidence stay intact.",
                          }),
                        }),
                        i.jsxs("div", {
                          className: "lg-table-inline",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              onClick: () =>
                                previewTableOperation({
                                  kind: "setColumnFormula",
                                  columnId: columnId,
                                  formula: columnFormula,
                                }),
                              children: i.jsx(LocalizedText, {
                                children: "Preview column formula",
                              }),
                            }),
                            i.jsx("button", {
                              type: "button",
                              disabled: !sheetTables
                                .find(
                                  (clearRuleTable) =>
                                    clearRuleTable.id === tableId,
                                )
                                ?.columns.find(
                                  (column) => column.id === columnId,
                                )?.calculatedFormula,
                              onClick: () =>
                                previewTableOperation({
                                  kind: "clearColumnFormula",
                                  columnId: columnId,
                                }),
                              children: i.jsx(LocalizedText, {
                                children: "Preview clear rule",
                              }),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Rows to append",
                            }),
                            i.jsx(LocalizedNumberField, {
                              "aria-label": text("Rows to append"),
                              min: "1",
                              max: "100",
                              value: appendCount,
                              onValueChange: (nextAppendCount) =>
                                setAppendCount(nextAppendCount),
                              lang: language("Rows to append"),
                              "data-ui-label": "Rows to append",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () =>
                            previewTableOperation({
                              kind: "appendRows",
                              count: Number(appendCount),
                            }),
                          children: i.jsx(LocalizedText, {
                            children: "Preview append rows",
                          }),
                        }),
                        i.jsx("p", {
                          className: "lg-table-help",
                          children: i.jsx(LocalizedText, {
                            children:
                              "Append 1–100 rows within the 500-cell change limit. Saved rules fill added rows. Adjacent cells must be empty. Clearing a rule keeps its existing cell formulas. Ordinary typing does not create rules or expand tables.",
                          }),
                        }),
                      ],
                    }),
                  pendingPreview &&
                    i.jsxs("section", {
                      className: "lg-table-preview",
                      "aria-label": text("Table change preview"),
                      lang: language("Table change preview"),
                      "data-ui-label": "Table change preview",
                      children: [
                        i.jsx("h3", {
                          children: i.jsx(LocalizedText, {
                            children: "Review table changes",
                          }),
                        }),
                        disabled
                          ? i.jsx("p", {
                              role: "alert",
                              children: i.jsx(LocalizedText, {
                                children:
                                  "The workbook, selection or table settings changed. Preview again.",
                              }),
                            })
                          : i.jsxs(i.Fragment, {
                              children: [
                                i.jsxs("p", {
                                  children: [
                                    pendingPreview.plan.beforeRange,
                                    " →",
                                    " ",
                                    pendingPreview.plan.afterRange,
                                    ";",
                                    " ",
                                    pendingPreview.plan.changes.length,
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "cell changes.",
                                    }),
                                  ],
                                }),
                                i.jsxs("details", {
                                  children: [
                                    i.jsx("summary", {
                                      children: i.jsx(LocalizedText, {
                                        children:
                                          "Saved column rules before and after",
                                      }),
                                    }),
                                    pendingPreview.plan.afterColumns.map(
                                      (afterColumn, columnIndex) =>
                                        i.jsxs(
                                          "p",
                                          {
                                            children: [
                                              i.jsx("strong", {
                                                children: afterColumn.name,
                                              }),
                                              i.jsx("br", {}),
                                              i.jsx(LocalizedText, {
                                                children: "Before:",
                                              }),
                                              " ",
                                              pendingPreview.plan.beforeColumns[
                                                columnIndex
                                              ]?.calculatedFormula ?? "No rule",
                                              i.jsx("br", {}),
                                              i.jsx(LocalizedText, {
                                                children: "After: ",
                                              }),
                                              afterColumn.calculatedFormula ??
                                                "No rule",
                                            ],
                                          },
                                          afterColumn.id,
                                        ),
                                    ),
                                  ],
                                }),
                                i.jsx("ol", {
                                  start: pageIndex * 20 + 1,
                                  children: pendingPreview.plan.changes
                                    .slice(pageIndex * 20, pageIndex * 20 + 20)
                                    .map((change) =>
                                      i.jsxs(
                                        "li",
                                        {
                                          children: [
                                            i.jsxs("strong", {
                                              children: [
                                                workbook.sheets.find(
                                                  (sheet) =>
                                                    sheet.id === change.sheetId,
                                                )?.name,
                                                "!",
                                                change.address,
                                              ],
                                            }),
                                            i.jsx("br", {}),
                                            i.jsx(LocalizedText, {
                                              children: "Before: ",
                                            }),
                                            JSON.stringify(
                                              change.before?.raw ?? null,
                                            ),
                                            i.jsx("br", {}),
                                            i.jsx(LocalizedText, {
                                              children: "After: ",
                                            }),
                                            JSON.stringify(
                                              change.after?.raw ?? null,
                                            ),
                                            i.jsxs("details", {
                                              children: [
                                                i.jsx("summary", {
                                                  children: i.jsx(
                                                    LocalizedText,
                                                    {
                                                      children:
                                                        "Full cell metadata",
                                                    },
                                                  ),
                                                }),
                                                i.jsx("pre", {
                                                  children: JSON.stringify(
                                                    {
                                                      before:
                                                        change.before ?? null,
                                                      after:
                                                        change.after ?? null,
                                                    },
                                                    null,
                                                    2,
                                                  ),
                                                }),
                                              ],
                                            }),
                                          ],
                                        },
                                        `${change.sheetId}:${change.address}`,
                                      ),
                                    ),
                                }),
                                pendingPreview.plan.changes.length > 20 &&
                                  i.jsxs("div", {
                                    className: "lg-table-inline",
                                    children: [
                                      i.jsx("button", {
                                        type: "button",
                                        disabled: pageIndex === 0,
                                        onClick: () =>
                                          setPageIndex(
                                            (pageBeforePrevious) =>
                                              pageBeforePrevious - 1,
                                          ),
                                        children: i.jsx(LocalizedText, {
                                          children: "Previous changes",
                                        }),
                                      }),
                                      i.jsxs("span", {
                                        children: [
                                          i.jsx(LocalizedText, {
                                            children: "Page ",
                                          }),
                                          pageIndex + 1,
                                          " ",
                                          i.jsx(LocalizedText, {
                                            children: "of",
                                          }),
                                          " ",
                                          Math.ceil(
                                            pendingPreview.plan.changes.length /
                                              20,
                                          ),
                                        ],
                                      }),
                                      i.jsx("button", {
                                        type: "button",
                                        disabled:
                                          (pageIndex + 1) * 20 >=
                                          pendingPreview.plan.changes.length,
                                        onClick: () =>
                                          setPageIndex(
                                            (pageBeforeNext) =>
                                              pageBeforeNext + 1,
                                          ),
                                        children: i.jsx(LocalizedText, {
                                          children: "Next changes",
                                        }),
                                      }),
                                    ],
                                  }),
                                pendingPreview.plan.anchorChanges.map(
                                  (anchorChange) =>
                                    i.jsxs(
                                      "p",
                                      {
                                        children: [
                                          anchorChange.kind,
                                          " ",
                                          anchorChange.id,
                                          ": ",
                                          anchorChange.before,
                                          " →",
                                          " ",
                                          anchorChange.after,
                                        ],
                                      },
                                      anchorChange.kind + anchorChange.id,
                                    ),
                                ),
                              ],
                            }),
                        i.jsxs("div", {
                          className: "lg-table-inline",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              className: "lg-table-primary",
                              disabled: disabled,
                              onClick: applyReviewedOperation,
                              children: i.jsx(LocalizedText, {
                                children: "Apply table changes",
                              }),
                            }),
                            i.jsx("button", {
                              type: "button",
                              onClick: () => setPendingPreview(null),
                              children: i.jsx(LocalizedText, {
                                children: "Cancel table preview",
                              }),
                            }),
                          ],
                        }),
                      ],
                    }),
                  i.jsx("p", {
                    className: "lg-table-help",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Creating a totals row requires an empty row below the source. When resizing an existing table, include its totals row in the range.",
                    }),
                  }),
                ],
              })
            : i.jsxs(i.Fragment, {
                children: [
                  mode === "pivot"
                    ? i.jsxs(i.Fragment, {
                        children: [
                          renderFieldCheckboxes(
                            "Row groups",
                            rowFields,
                            setRowFields,
                          ),
                          renderFieldCheckboxes(
                            "Column groups",
                            columnFields,
                            setColumnFields,
                          ),
                          i.jsxs("fieldset", {
                            className: "lg-table-list",
                            children: [
                              i.jsx("legend", {
                                children: i.jsx(LocalizedText, {
                                  children: "Value fields",
                                }),
                              }),
                              values.map((value, valueIndex) =>
                                i.jsxs(
                                  "div",
                                  {
                                    className: "lg-table-value",
                                    children: [
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(LocalizedText, {
                                            children: "Value field ",
                                          }),
                                          valueIndex + 1,
                                          i.jsx("select", {
                                            value: value.field,
                                            onChange: (event) =>
                                              setValues((previousFieldValues) =>
                                                previousFieldValues.map(
                                                  (
                                                    valueFieldItem,
                                                    valueRowIndex,
                                                  ) =>
                                                    valueRowIndex === valueIndex
                                                      ? {
                                                          ...valueFieldItem,
                                                          field:
                                                            event.target.value,
                                                        }
                                                      : valueFieldItem,
                                                ),
                                              ),
                                            children: fields
                                              .concat(
                                                calculatedFields.map(
                                                  (calculatedField) =>
                                                    calculatedField.name,
                                                ),
                                              )
                                              .map((fieldOptionName) =>
                                                i.jsx(
                                                  "option",
                                                  {
                                                    value: fieldOptionName,
                                                    children: fieldOptionName,
                                                  },
                                                  fieldOptionName,
                                                ),
                                              ),
                                          }),
                                        ],
                                      }),
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(LocalizedText, {
                                            children: "Summarize by",
                                          }),
                                          i.jsx("select", {
                                            "aria-label": `Aggregation ${valueIndex + 1}`,
                                            value: value.operation,
                                            onChange: (event) =>
                                              setValues(
                                                (previousAggregateValues) =>
                                                  previousAggregateValues.map(
                                                    (
                                                      valueAggregateItem,
                                                      valueAggregateIndex,
                                                    ) =>
                                                      valueAggregateIndex ===
                                                      valueIndex
                                                        ? {
                                                            ...valueAggregateItem,
                                                            operation:
                                                              event.target
                                                                .value,
                                                          }
                                                        : valueAggregateItem,
                                                  ),
                                              ),
                                            children:
                                              pivotAggregateOperations.map(
                                                (pivotAggregateOperation) =>
                                                  i.jsx(
                                                    "option",
                                                    {
                                                      value:
                                                        pivotAggregateOperation,
                                                      children:
                                                        formatAggregationLabel(
                                                          pivotAggregateOperation,
                                                        ),
                                                    },
                                                    pivotAggregateOperation,
                                                  ),
                                              ),
                                          }),
                                        ],
                                      }),
                                      i.jsx("button", {
                                        type: "button",
                                        "aria-label": `Remove value ${valueIndex + 1}`,
                                        onClick: () =>
                                          setValues((valuesBeforeRemove) =>
                                            valuesBeforeRemove.filter(
                                              (
                                                valueRemoveItem,
                                                valueRemoveIndex,
                                              ) =>
                                                valueRemoveIndex !== valueIndex,
                                            ),
                                          ),
                                        children: i.jsx(LocalizedText, {
                                          children: "Remove",
                                        }),
                                      }),
                                    ],
                                  },
                                  valueIndex,
                                ),
                              ),
                              i.jsx("button", {
                                type: "button",
                                disabled:
                                  !fields.length || values.length >= 100,
                                onClick: () =>
                                  setValues((valuesBeforeAddButton) => [
                                    ...valuesBeforeAddButton,
                                    { field: fields.at(-1), operation: "sum" },
                                  ]),
                                children: i.jsx(LocalizedText, {
                                  children: "Add value field",
                                }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            className: "lg-table-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: grandTotals,
                                onChange: (event) =>
                                  setGrandTotals(event.target.checked),
                              }),
                              i.jsx(LocalizedText, {
                                children: "Grand totals",
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            className: "lg-table-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: subtotals,
                                onChange: (event) =>
                                  setSubtotals(event.target.checked),
                              }),
                              i.jsx(LocalizedText, {
                                children: "Row and column subtotals",
                              }),
                            ],
                          }),
                          i.jsx(PivotFieldOptionsPanel, {
                            fields: fields,
                            groups: [...rowFields, ...columnFields],
                            dateGroups: dateGroups,
                            setDateGroups: setDateGroups,
                            calculatedFields: calculatedFields,
                            setCalculatedFields: setCalculatedFields,
                            onAdd: (field) =>
                              setValues((valuesBeforeAddPanel) => [
                                ...valuesBeforeAddPanel,
                                { field: field, operation: "sum" },
                              ]),
                          }),
                        ],
                      })
                    : i.jsxs(i.Fragment, {
                        children: [
                          renderFieldCheckboxes(
                            "Output fields",
                            extractFields,
                            (nextOutputFields) => {
                              (setExtractFields(nextOutputFields),
                                setRename((renameBeforeFilter) =>
                                  Object.fromEntries(
                                    Object.entries(renameBeforeFilter).filter(
                                      ([renamedFieldName]) =>
                                        nextOutputFields.includes(
                                          renamedFieldName,
                                        ),
                                    ),
                                  ),
                                ));
                            },
                          ),
                          i.jsxs("details", {
                            children: [
                              i.jsx("summary", {
                                children: i.jsx(LocalizedText, {
                                  children: "Rename output fields",
                                }),
                              }),
                              i.jsx("div", {
                                className: "lg-table-list",
                                children: extractFields.map((fieldName) =>
                                  i.jsxs(
                                    "label",
                                    {
                                      children: [
                                        fieldName,
                                        i.jsx("input", {
                                          "aria-label": `Rename ${fieldName}`,
                                          value: Object.hasOwn(
                                            rename,
                                            fieldName,
                                          )
                                            ? rename[fieldName]
                                            : fieldName,
                                          onChange: (event) =>
                                            setRename((renameBeforeEdit) =>
                                              Object.defineProperty(
                                                { ...renameBeforeEdit },
                                                fieldName,
                                                {
                                                  value: event.target.value,
                                                  enumerable: true,
                                                  writable: true,
                                                  configurable: true,
                                                },
                                              ),
                                            ),
                                        }),
                                      ],
                                    },
                                    fieldName,
                                  ),
                                ),
                              }),
                            ],
                          }),
                          i.jsxs("div", {
                            className: "lg-table-inline",
                            children: [
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, { children: "Sort by" }),
                                  i.jsx("select", {
                                    value: field,
                                    onChange: (event) =>
                                      setField(event.target.value),
                                    children:
                                      renderFieldOptions("Source order"),
                                  }),
                                ],
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, { children: "Order" }),
                                  i.jsxs("select", {
                                    value: direction,
                                    onChange: (event) =>
                                      setDirection(event.target.value),
                                    children: [
                                      i.jsx("option", {
                                        value: "ascending",
                                        children: "Ascending",
                                      }),
                                      i.jsx("option", {
                                        value: "descending",
                                        children: "Descending",
                                      }),
                                    ],
                                  }),
                                ],
                              }),
                            ],
                          }),
                          i.jsxs("fieldset", {
                            className: "lg-table-list",
                            children: [
                              i.jsx("legend", {
                                children: i.jsx(LocalizedText, {
                                  children: "Derived fields",
                                }),
                              }),
                              derived.map((derivedField, derivedIndex) =>
                                i.jsxs(
                                  "div",
                                  {
                                    className: "lg-table-derived",
                                    children: [
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(LocalizedText, {
                                            children: "New field name",
                                          }),
                                          i.jsx("input", {
                                            value: derivedField.name,
                                            onChange: (event) =>
                                              setDerived(
                                                (previousDerivedName) =>
                                                  previousDerivedName.map(
                                                    (
                                                      derivedNameItem,
                                                      nameRowIndex,
                                                    ) =>
                                                      nameRowIndex ===
                                                      derivedIndex
                                                        ? {
                                                            ...derivedNameItem,
                                                            name: event.target
                                                              .value,
                                                          }
                                                        : derivedNameItem,
                                                  ),
                                              ),
                                          }),
                                        ],
                                      }),
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(LocalizedText, {
                                            children: "Source field",
                                          }),
                                          i.jsx("select", {
                                            value: derivedField.field,
                                            onChange: (event) =>
                                              setDerived(
                                                (previousDerivedSource) =>
                                                  previousDerivedSource.map(
                                                    (
                                                      derivedFieldItem,
                                                      fieldRowIndex,
                                                    ) =>
                                                      fieldRowIndex ===
                                                      derivedIndex
                                                        ? {
                                                            ...derivedFieldItem,
                                                            field:
                                                              event.target
                                                                .value,
                                                          }
                                                        : derivedFieldItem,
                                                  ),
                                              ),
                                            children: renderFieldOptions(),
                                          }),
                                        ],
                                      }),
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(LocalizedText, {
                                            children: "Operation",
                                          }),
                                          i.jsx("select", {
                                            value: derivedField.operation,
                                            onChange: (event) =>
                                              setDerived(
                                                (previousDerivedOperation) =>
                                                  previousDerivedOperation.map(
                                                    (
                                                      derivedOperationItem,
                                                      operationRowIndex,
                                                    ) =>
                                                      operationRowIndex ===
                                                      derivedIndex
                                                        ? {
                                                            ...derivedOperationItem,
                                                            operation:
                                                              event.target
                                                                .value,
                                                          }
                                                        : derivedOperationItem,
                                                  ),
                                              ),
                                            children: [
                                              "add",
                                              "subtract",
                                              "multiply",
                                              "divide",
                                              "concat",
                                              "upper",
                                              "lower",
                                              "trim",
                                            ].map((operationName) =>
                                              i.jsx(
                                                "option",
                                                {
                                                  value: operationName,
                                                  children:
                                                    formatAggregationLabel(
                                                      operationName,
                                                    ),
                                                },
                                                operationName,
                                              ),
                                            ),
                                          }),
                                        ],
                                      }),
                                      !["upper", "lower", "trim"].includes(
                                        derivedField.operation,
                                      ) &&
                                        i.jsxs(i.Fragment, {
                                          children: [
                                            i.jsxs("label", {
                                              children: [
                                                i.jsx(LocalizedText, {
                                                  children: "Second field",
                                                }),
                                                i.jsx("select", {
                                                  value:
                                                    derivedField.otherField ??
                                                    "",
                                                  onChange: (event) =>
                                                    setDerived(
                                                      (previousDerivedOther) =>
                                                        previousDerivedOther.map(
                                                          (
                                                            derivedOtherItem,
                                                            otherRowIndex,
                                                          ) =>
                                                            otherRowIndex ===
                                                            derivedIndex
                                                              ? {
                                                                  ...derivedOtherItem,
                                                                  otherField:
                                                                    event.target
                                                                      .value ||
                                                                    undefined,
                                                                }
                                                              : derivedOtherItem,
                                                        ),
                                                    ),
                                                  children:
                                                    renderFieldOptions(
                                                      "Use a constant",
                                                    ),
                                                }),
                                              ],
                                            }),
                                            !derivedField.otherField &&
                                              i.jsxs("label", {
                                                children: [
                                                  i.jsx(LocalizedText, {
                                                    children: "Constant",
                                                  }),
                                                  derivedField.operation ===
                                                  "concat"
                                                    ? i.jsx("input", {
                                                        value:
                                                          derivedField.constant ??
                                                          "",
                                                        onChange: (event) =>
                                                          setDerived(
                                                            (
                                                              previousDerivedConstant,
                                                            ) =>
                                                              previousDerivedConstant.map(
                                                                (
                                                                  derivedConstantItem,
                                                                  constantRowIndex,
                                                                ) =>
                                                                  constantRowIndex ===
                                                                  derivedIndex
                                                                    ? {
                                                                        ...derivedConstantItem,
                                                                        constant:
                                                                          event
                                                                            .target
                                                                            .value,
                                                                      }
                                                                    : derivedConstantItem,
                                                              ),
                                                          ),
                                                      })
                                                    : i.jsx(
                                                        LocalizedNumberField,
                                                        {
                                                          "aria-label":
                                                            text("Constant"),
                                                          value:
                                                            derivedField.constant ??
                                                            "",
                                                          onValueChange: (
                                                            constantValue,
                                                          ) =>
                                                            setDerived(
                                                              (
                                                                previousDerivedNumber,
                                                              ) =>
                                                                previousDerivedNumber.map(
                                                                  (
                                                                    derivedNumberItem,
                                                                    numberRowIndex,
                                                                  ) =>
                                                                    numberRowIndex ===
                                                                    derivedIndex
                                                                      ? {
                                                                          ...derivedNumberItem,
                                                                          constant:
                                                                            Number(
                                                                              constantValue,
                                                                            ),
                                                                        }
                                                                      : derivedNumberItem,
                                                                ),
                                                            ),
                                                        },
                                                      ),
                                                ],
                                              }),
                                          ],
                                        }),
                                      i.jsx("button", {
                                        type: "button",
                                        onClick: () =>
                                          setDerived((derivedBeforeRemove) =>
                                            derivedBeforeRemove.filter(
                                              (
                                                derivedRemoveItem,
                                                derivedRemoveIndex,
                                              ) =>
                                                derivedRemoveIndex !==
                                                derivedIndex,
                                            ),
                                          ),
                                        children: i.jsx(LocalizedText, {
                                          children: "Remove derived field",
                                        }),
                                      }),
                                    ],
                                  },
                                  derivedIndex,
                                ),
                              ),
                              i.jsx("button", {
                                type: "button",
                                disabled:
                                  !fields.length || derived.length >= 100,
                                onClick: () =>
                                  setDerived((derivedBeforeAdd) => [
                                    ...derivedBeforeAdd,
                                    {
                                      name: `Calculated${derivedBeforeAdd.length + 1}`,
                                      field: fields.at(-1),
                                      operation: "multiply",
                                      constant: 1,
                                    },
                                  ]),
                                children: i.jsx(LocalizedText, {
                                  children: "Add derived field",
                                }),
                              }),
                            ],
                          }),
                        ],
                      }),
                  i.jsxs("fieldset", {
                    className: "lg-table-list",
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Source filters",
                        }),
                      }),
                      filterRules.map((renderedFilterRule, filterIndex) =>
                        i.jsxs(
                          "div",
                          {
                            className: "lg-table-filter",
                            children: [
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Filter field ",
                                  }),
                                  filterIndex + 1,
                                  i.jsx("select", {
                                    value: renderedFilterRule.field,
                                    onChange: (event) =>
                                      setFilterRules((previousFilterFields) =>
                                        previousFilterFields.map(
                                          (
                                            filterFieldItem,
                                            filterFieldIndex,
                                          ) =>
                                            filterFieldIndex === filterIndex
                                              ? {
                                                  ...filterFieldItem,
                                                  field: event.target.value,
                                                }
                                              : filterFieldItem,
                                        ),
                                      ),
                                    children: renderFieldOptions(),
                                  }),
                                ],
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Condition",
                                  }),
                                  i.jsx("select", {
                                    value: renderedFilterRule.operator,
                                    onChange: (event) =>
                                      setFilterRules(
                                        (previousFilterOperators) =>
                                          previousFilterOperators.map(
                                            (
                                              filterOperatorItem,
                                              filterOperatorIndex,
                                            ) =>
                                              filterOperatorIndex ===
                                              filterIndex
                                                ? {
                                                    ...filterOperatorItem,
                                                    operator:
                                                      event.target.value,
                                                  }
                                                : filterOperatorItem,
                                          ),
                                      ),
                                    children: [
                                      ["equals", "Equals"],
                                      ["notEquals", "Does not equal"],
                                      ["contains", "Contains"],
                                      ["gt", "Greater than"],
                                      ["gte", "At least"],
                                      ["lt", "Less than"],
                                      ["lte", "At most"],
                                      ["blank", "Is blank"],
                                      ["notBlank", "Is not blank"],
                                    ].map(([operatorValue, operatorLabel]) =>
                                      i.jsx(
                                        "option",
                                        {
                                          value: operatorValue,
                                          children: operatorLabel,
                                        },
                                        operatorValue,
                                      ),
                                    ),
                                  }),
                                ],
                              }),
                              !["blank", "notBlank"].includes(
                                renderedFilterRule.operator,
                              ) &&
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Filter value",
                                    }),
                                    i.jsx("input", {
                                      value: String(
                                        renderedFilterRule.value ?? "",
                                      ),
                                      onChange: (event) =>
                                        setFilterRules((previousFilterValues) =>
                                          previousFilterValues.map(
                                            (
                                              filterValueItem,
                                              filterValueIndex,
                                            ) =>
                                              filterValueIndex === filterIndex
                                                ? {
                                                    ...filterValueItem,
                                                    value: event.target.value,
                                                  }
                                                : filterValueItem,
                                          ),
                                        ),
                                    }),
                                  ],
                                }),
                              i.jsx("button", {
                                type: "button",
                                "aria-label": `Remove filter ${filterIndex + 1}`,
                                onClick: () =>
                                  setFilterRules((filterRulesBeforeRemove) =>
                                    filterRulesBeforeRemove.filter(
                                      (filterRemoveItem, filterRemoveIndex) =>
                                        filterRemoveIndex !== filterIndex,
                                    ),
                                  ),
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            ],
                          },
                          filterIndex,
                        ),
                      ),
                      i.jsx("button", {
                        type: "button",
                        disabled: !fields.length || filterRules.length >= 100,
                        onClick: () =>
                          setFilterRules((filterRulesBeforeAdd) => [
                            ...filterRulesBeforeAdd,
                            { field: fields[0], operator: "equals", value: "" },
                          ]),
                        children: i.jsx(LocalizedText, {
                          children: "Add source filter",
                        }),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    className: "lg-table-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: checked,
                        onChange: (event) => setChecked(event.target.checked),
                      }),
                      i.jsx(LocalizedText, {
                        children: "Use an empty area in an existing sheet",
                      }),
                    ],
                  }),
                  checked &&
                    i.jsxs("div", {
                      className: "lg-table-inline",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Destination sheet",
                            }),
                            i.jsx("select", {
                              value: destinationSheetId,
                              onChange: (event) =>
                                setDestinationSheetId(event.target.value),
                              children: workbook.sheets.map((sheet) =>
                                i.jsx(
                                  "option",
                                  { value: sheet.id, children: sheet.name },
                                  sheet.id,
                                ),
                              ),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Destination cell",
                            }),
                            i.jsx("input", {
                              value: address,
                              onChange: (event) =>
                                setAddress(event.target.value),
                            }),
                          ],
                        }),
                      ],
                    }),
                  i.jsx("p", {
                    className: "lg-table-help",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Results preserve the source. An existing destination must be empty. Refresh creates a new result so earlier work remains available.",
                    }),
                  }),
                  definitionEntries.length > 0 &&
                    i.jsxs("fieldset", {
                      className: "lg-table-list",
                      children: [
                        i.jsx("legend", {
                          children: i.jsx(LocalizedText, {
                            children: "Refresh a saved result",
                          }),
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Saved result" }),
                            i.jsxs("select", {
                              value: pivotId,
                              onChange: (event) =>
                                setPivotId(event.target.value),
                              children: [
                                i.jsx("option", {
                                  value: "",
                                  children: "Choose a result",
                                }),
                                definitionEntries.map(
                                  ({ sheetName, definition }) =>
                                    i.jsxs(
                                      "option",
                                      {
                                        value: definition.id,
                                        children: [
                                          sheetName,
                                          ": ",
                                          definition.name,
                                        ],
                                      },
                                      definition.id,
                                    ),
                                ),
                              ],
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled: !pivotId,
                          onClick: () =>
                            runOperation(
                              () =>
                                mode === "pivot"
                                  ? refreshPivotTable(
                                      workbook,
                                      pivotId,
                                      getDestination(),
                                    )
                                  : refreshQueryDefinition(
                                      workbook,
                                      pivotId,
                                      getDestination(),
                                    ),
                              mode === "pivot"
                                ? "Refresh pivot summary"
                                : "Refresh extracted data",
                            ),
                          children: i.jsx(LocalizedText, {
                            children: "Refresh into a new result",
                          }),
                        }),
                        mode === "pivot" &&
                          pivotId &&
                          i.jsx(PivotStatusBadge, {
                            workbook,
                            pivotId: pivotId,
                          }),
                      ],
                    }),
                ],
              }),
          i.jsx("footer", {
            children: i.jsx("button", {
              className: "lg-table-primary",
              type: "button",
              onClick: runPrimaryAction,
              children:
                mode === "table"
                  ? tableId
                    ? "Update table"
                    : "Create table"
                  : mode === "pivot"
                    ? "Create pivot summary"
                    : "Extract data",
            }),
          }),
        ],
      }),
    ],
  });
}
const pivotShowAsLabels = {
    value: "No calculation",
    percentGrand: "% of grand total",
    percentRow: "% of row total",
    percentColumn: "% of column total",
  },
  formatPivotGroupLabel = (groupValue) =>
    groupValue === null
      ? "(blank)"
      : groupValue === ""
        ? "(empty text)"
        : `${typeof groupValue}: ${String(groupValue)}`;
function PivotChartFigure({ book, id: pivotId }) {
  const { text, language } = useWorkbookMessages();
  let chartModel;
  try {
    chartModel = buildPivotChartModel(book, pivotId);
  } catch (error) {
    return i.jsx("p", { role: "alert", children: error.message });
  }
  const values = chartModel.points.map((point) => point.value),
    nonNullValues = values.filter((m) => m !== null),
    segments = values
      .reduce(
        (runs, value, valueIndex) => (
          value === null
            ? runs.push([])
            : runs[runs.length - 1].push({ v: value, i: valueIndex }),
          runs
        ),
        [[]],
      )
      .filter((run) => run.length),
    axisMin = Math.min(0, ...nonNullValues),
    axisMax = Math.max(1, ...nonNullValues),
    valueToY = (pointValue) =>
      180 - ((pointValue - axisMin) / (axisMax - axisMin)) * 145,
    groupWidth = 420 / Math.max(1, values.length),
    zeroY = valueToY(0),
    isPercent = chartModel.showAs !== "value";
  return i.jsxs("figure", {
    className: "lg-pivot-chart",
    "aria-label": text("PivotChart"),
    lang: language("PivotChart"),
    "data-ui-label": "PivotChart",
    children: [
      i.jsxs("figcaption", {
        children: [
          chartModel.measure,
          isPercent ? ` — ${pivotShowAsLabels[chartModel.showAs]}` : "",
          " ",
          i.jsx(LocalizedText, { children: "· source revision" }),
          " ",
          chartModel.sourceRevision,
        ],
      }),
      values.length
        ? i.jsxs("svg", {
            role: "img",
            "aria-label": `${chartModel.kind} PivotChart with ${values.length} groups; values are listed below`,
            viewBox: "0 0 470 210",
            children: [
              chartModel.kind === "bar"
                ? i.jsx("line", {
                    x1: 35 - (axisMin / (axisMax - axisMin)) * 420,
                    x2: 35 - (axisMin / (axisMax - axisMin)) * 420,
                    y1: "28",
                    y2: "180",
                    stroke: "#607780",
                  })
                : i.jsx("line", {
                    x1: "35",
                    x2: "455",
                    y1: zeroY,
                    y2: zeroY,
                    stroke: "#607780",
                  }),
              i.jsx("text", {
                x: chartModel.kind === "bar" ? 455 : 30,
                y: chartModel.kind === "bar" ? 198 : valueToY(axisMax) + 4,
                textAnchor: "end",
                fontSize: "10",
                children: isPercent
                  ? `${(axisMax * 100).toFixed(0)}%`
                  : axisMax.toFixed(1),
              }),
              i.jsx("text", {
                x: chartModel.kind === "bar" ? 35 : 30,
                y: chartModel.kind === "bar" ? 198 : valueToY(axisMin) + 4,
                textAnchor: chartModel.kind === "bar" ? "start" : "end",
                fontSize: "10",
                children: isPercent
                  ? `${(axisMin * 100).toFixed(0)}%`
                  : axisMin.toFixed(1),
              }),
              chartModel.kind === "line"
                ? segments.map((segment, segmentIndex) =>
                    i.jsx(
                      "polyline",
                      {
                        fill: "none",
                        stroke: "#217346",
                        strokeWidth: "3",
                        points: segment
                          .map(
                            ({ v: value, i: valueIndex }) =>
                              `${35 + groupWidth * (valueIndex + 0.5)},${valueToY(value)}`,
                          )
                          .join(" "),
                      },
                      segmentIndex,
                    ),
                  )
                : values.map((pointValue, pointIndex) =>
                    pointValue === null
                      ? null
                      : chartModel.kind === "bar"
                        ? i.jsx(
                            "rect",
                            {
                              x:
                                35 +
                                ((Math.min(pointValue, 0) - axisMin) /
                                  (axisMax - axisMin)) *
                                  420,
                              y: 28 + pointIndex * (150 / values.length),
                              width: Math.max(
                                1,
                                (Math.abs(pointValue) / (axisMax - axisMin)) *
                                  420,
                              ),
                              height: Math.max(1, 120 / values.length),
                              fill: "#217346",
                              children: i.jsxs("title", {
                                children: [
                                  chartModel.points[pointIndex].label,
                                  ": ",
                                  pointValue,
                                ],
                              }),
                            },
                            pointIndex,
                          )
                        : i.jsx(
                            "rect",
                            {
                              x: 36 + groupWidth * pointIndex,
                              width: Math.max(1, groupWidth - 3),
                              y: Math.min(valueToY(pointValue), zeroY),
                              height: Math.max(
                                1,
                                Math.abs(valueToY(pointValue) - zeroY),
                              ),
                              fill: "#217346",
                              children: i.jsxs("title", {
                                children: [
                                  chartModel.points[pointIndex].label,
                                  ": ",
                                  pointValue,
                                ],
                              }),
                            },
                            pointIndex,
                          ),
                  ),
            ],
          })
        : i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children: "No records match these filters.",
            }),
          }),
      i.jsxs("details", {
        children: [
          i.jsx("summary", {
            children: i.jsx(LocalizedText, {
              children: "Chart values and categories",
            }),
          }),
          i.jsxs("table", {
            children: [
              i.jsx("thead", {
                children: i.jsxs("tr", {
                  children: [
                    i.jsx("th", {
                      children: i.jsx(LocalizedText, { children: "Group" }),
                    }),
                    i.jsx("th", {
                      children: i.jsx(LocalizedText, { children: "Value" }),
                    }),
                  ],
                }),
              }),
              i.jsx("tbody", {
                children: chartModel.points.map((point, pointIndex) =>
                  i.jsxs(
                    "tr",
                    {
                      children: [
                        i.jsx("td", { children: point.label }),
                        i.jsx("td", {
                          children:
                            point.value === null
                              ? "Blank"
                              : isPercent
                                ? `${(Number(point.value) * 100).toFixed(2)}%`
                                : String(point.value),
                        }),
                      ],
                    },
                    pointIndex,
                  ),
                ),
              }),
            ],
          }),
        ],
      }),
      i.jsx("p", {
        children: i.jsx(LocalizedText, {
          children:
            "Leaf row groups only; subtotals and grand totals are excluded. Column groups are combined using the selected aggregate.",
        }),
      }),
    ],
  });
}

export {
  pivotShowAsLabels,
  PivotChartFigure,
  formatPivotGroupLabel,
  TableToolsDialog,
};
