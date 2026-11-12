import {
  isNonArrayObject,
  createUuid,
  failWithMessage,
} from "./03-workbook-model.js";
import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { formatRangeCornerCells } from "./08-xlsx-import-export-and-preservation.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  columnarModelFormatVersion,
  formatOccupiedRangeAddress,
  _t,
  cn,
  parseRangeAddress,
  dataModelLimits,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  previewDataModelOutput,
  DataModelDefinitionEditor,
  publishDataModelReport,
  getModelReportStatus,
  removeDataModelFromWorkbook,
  LabeledField,
  buildModelTableFromRange,
  DaxExpressionInput,
} from "./14-dax-data-model-engine.js";

function DataModelDialog(props) {
  const { text } = useWorkbookMessages(),
    propsRef = ne.useRef(props);
  propsRef.current = props;
  const titleId = ne.useId(),
    savedReport = props.workbook.sheets.find(
      (sheet) => sheet.id === props.sheetId,
    )?.dataModelReport,
    [definition, setDefinition] = ne.useState(() =>
      structuredClone(
        savedReport?.definition ??
          props.workbook.dataModel ?? {
            engine: columnarModelFormatVersion,
            name: "Workbook model",
            tables: [],
            relationships: [],
            measures: [],
          },
      ),
    ),
    [query, setQuery] = ne.useState(() =>
      structuredClone(
        savedReport?.query ?? {
          groupBy: [],
          filters: [],
          measureIds: definition.measures.map((measure) => measure.id),
        },
      ),
    ),
    [activeSection, setActiveSection] = ne.useState("Tables"),
    [sheetId, setSheetId] = ne.useState(props.sheetId),
    [rangeAddress, setRangeAddress] = ne.useState(
      props.range.includes(":")
        ? props.range
        : formatOccupiedRangeAddress(
            props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
          ),
    ),
    [structuredSourceId, setStructuredSourceId] = ne.useState(""),
    [checked, setChecked] = ne.useState(false),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [previewResult, setPreviewResult] = ne.useState(null),
    [resultPage, setResultPage] = ne.useState(0),
    [outputName, setOutputName] = ne.useState(() => {
      let resultNumber = 1;
      for (
        ;
        props.workbook.sheets.some(
          (sheet) =>
            sheet.name.toLowerCase() === `model result ${resultNumber}`,
        );
      )
        resultNumber++;
      return `Model result ${resultNumber}`;
    }),
    dialogRef = ne.useRef(null),
    hasPublished = ne.useRef(false);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const modelOptions = {
      readOnly: props.readOnly,
      allowedSheetIds: props.allowedSheetIds,
      includeHidden: !!props.allowHidden && checked,
    },
    resetPreview = () => {
      (setPreviewResult(null),
        setErrorMessage(""),
        (hasPublished.current = false));
    },
    onChange = (nextDefinition) => {
      (setDefinition(nextDefinition), resetPreview());
    },
    updateQuery = (nextQuery) => {
      (setQuery(nextQuery), resetPreview());
    },
    findTable = (tableId) =>
      definition.tables.find((table) => table.id === tableId),
    defaultColumnName = (tableId) => findTable(tableId)?.columns[0]?.name ?? "",
    renderTextInput = (label, value, onValueChange) =>
      i.jsx("input", {
        "aria-label": text(label),
        value: value,
        onChange: (event) => onValueChange(event.target.value),
      }),
    renderSelect = (label, value, options, onValueChange) =>
      i.jsx("select", {
        "aria-label": text(label),
        value: value,
        onChange: (event) => onValueChange(event.target.value),
        children: options.map((option) =>
          i.jsx(
            "option",
            { value: option.value, children: text(option.label) },
            option.value,
          ),
        ),
      }),
    toStringOptions = (values) =>
      values.map((optionValue) => ({ value: optionValue, label: optionValue })),
    renderTableSelect = (label, value, onValueChange) =>
      renderSelect(
        label,
        value,
        definition.tables.map((table) => ({
          value: table.id,
          label: table.name,
        })),
        onValueChange,
      ),
    renderColumnSelect = (label, tableId, value, onValueChange) =>
      renderSelect(
        label,
        value,
        toStringOptions(
          findTable(tableId)?.columns.map((column) => column.name) ?? [],
        ),
        onValueChange,
      ),
    updateTable = (tableId, changes) =>
      onChange({
        ...definition,
        tables: definition.tables.map((table) =>
          table.id === tableId ? { ...table, ...changes } : table,
        ),
      }),
    deleteTable = (tableId) => {
      const measures = definition.measures.filter(
        (measure) => measure.tableId !== tableId,
      );
      (onChange({
        ...definition,
        tables: definition.tables.filter((table) => table.id !== tableId),
        relationships: definition.relationships.filter(
          (relationship) =>
            relationship.one.tableId !== tableId &&
            relationship.many.tableId !== tableId,
        ),
        measures: measures,
      }),
        updateQuery({
          ...query,
          groupBy: query.groupBy.filter(
            (groupField) => groupField.tableId !== tableId,
          ),
          filters: query.filters.filter((filter) => filter.tableId !== tableId),
          measureIds: query.measureIds.filter((measureId) =>
            measures.some((measure) => measure.id === measureId),
          ),
        }));
    };
  function X$1() {
    try {
      const modelTableFromRange = buildModelTableFromRange(
        props.workbook,
        sheetId,
        rangeAddress,
        structuredSourceId || undefined,
        modelOptions,
      );
      let tableName = modelTableFromRange.name,
        nameSuffix = 2;
      for (
        ;
        definition.tables.some(
          (table) => table.name.toLowerCase() === tableName.toLowerCase(),
        );
      )
        tableName = `${modelTableFromRange.name} ${nameSuffix++}`;
      onChange({
        ...definition,
        tables: [
          ...definition.tables,
          { ...modelTableFromRange, name: tableName },
        ],
      });
    } catch (error) {
      setErrorMessage(error.message);
    }
  }
  function onClick() {
    try {
      (setErrorMessage(""),
        setResultPage(0),
        (hasPublished.current = false),
        setPreviewResult(
          previewDataModelOutput(
            props.workbook,
            { definition: definition, query: query, outputName: outputName },
            modelOptions,
          ),
        ),
        setActiveSection("Preview"));
    } catch (error) {
      (setPreviewResult(null), setErrorMessage(error.message));
    }
  }
  function commitModelReport(mode) {
    if (!(!previewResult || hasPublished.current))
      try {
        const latestProps = propsRef.current,
          nextWorkbook = publishDataModelReport(
            latestProps.workbook,
            previewResult,
            mode,
            {
              readOnly: latestProps.readOnly,
              allowedSheetIds: latestProps.allowedSheetIds,
              includeHidden: !!latestProps.allowHidden && checked,
            },
          );
        ((hasPublished.current = true),
          latestProps.onWorkbookChange(
            nextWorkbook,
            mode === "save"
              ? "Save relational data model"
              : "Publish relational model result",
          ),
          latestProps.onClose());
      } catch (error) {
        setErrorMessage(error.message);
      }
  }
  const reportResult = previewResult?.result,
    canPublish =
      !!reportResult?.complete &&
      !props.readOnly &&
      !props.workbook.structureProtection?.enabled,
    selectableSheets = props.workbook.sheets.filter(
      (sheet) =>
        (!props.allowedSheetIds || props.allowedSheetIds.includes(sheet.id)) &&
        ((!sheet.hidden && !sheet.veryHidden) ||
          (props.allowHidden && checked)),
    );
  return i.jsx("div", {
    className: "lg-modal-backdrop",
    children: i.jsxs("section", {
      className: "lg-data-model",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      ref: dialogRef,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.stopPropagation(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
              ...dialogRef.current.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),select:not(:disabled),summary",
              ),
            ].filter(
              (element) =>
                element.getClientRects().length ||
                navigator.userAgent.includes("jsdom"),
            ),
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
            i.jsxs("div", {
              children: [
                i.jsx("p", {
                  className: "eyebrow",
                  children: i.jsx(LocalizedText, {
                    children: "Data · Relationships",
                  }),
                }),
                i.jsx("h2", {
                  id: titleId,
                  children: i.jsx(LocalizedText, { children: "Data Model" }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Connect typed tables and review measures across shared dimensions.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close data model"),
              onClick: props.onClose,
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsx("nav", {
          "aria-label": text("Data model sections"),
          children: [
            "Tables",
            "Calculated columns",
            "Relationships",
            "Measures",
            "Query",
            "Preview",
          ].map((sectionName) =>
            i.jsx(
              "button",
              {
                "aria-pressed": activeSection === sectionName,
                onClick: () => setActiveSection(sectionName),
                children: i.jsx(LocalizedText, { children: sectionName }),
              },
              sectionName,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "lg-model-body",
          children: [
            errorMessage &&
              i.jsx("p", {
                role: "alert",
                className: "lg-model-error",
                children: text(errorMessage),
              }),
            props.readOnly &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "This workbook is read-only. Preview remains available.",
                }),
              }),
            activeSection === "Tables" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("div", {
                    className: "lg-model-controls",
                    children: [
                      LabeledField(
                        "Model name",
                        renderTextInput(
                          "Model name",
                          definition.name,
                          (modelName) =>
                            onChange({ ...definition, name: modelName }),
                        ),
                      ),
                      LabeledField(
                        "Source worksheet",
                        renderSelect(
                          "Model source worksheet",
                          sheetId,
                          selectableSheets.map((sheet) => ({
                            value: sheet.id,
                            label: sheet.name,
                          })),
                          (nextSheetId) => {
                            (setSheetId(nextSheetId),
                              setStructuredSourceId(""),
                              setRangeAddress(
                                formatOccupiedRangeAddress(
                                  props.workbook.sheets.find(
                                    (sheet) => sheet.id === nextSheetId,
                                  ),
                                ),
                              ));
                          },
                        ),
                      ),
                      LabeledField(
                        "Source table or range",
                        renderSelect(
                          "Model structured source",
                          structuredSourceId,
                          [
                            { value: "", label: "Cell range" },
                            ...(
                              props.workbook.sheets.find(
                                (sheet) => sheet.id === sheetId,
                              )?.tables ?? []
                            ).map((table) => ({
                              value: table.id,
                              label: table.name,
                            })),
                          ],
                          setStructuredSourceId,
                        ),
                      ),
                      !structuredSourceId &&
                        LabeledField(
                          "Headed range",
                          renderTextInput(
                            "Model source range",
                            rangeAddress,
                            setRangeAddress,
                          ),
                        ),
                      i.jsx("button", {
                        disabled:
                          definition.tables.length >= dataModelLimits.tables,
                        onClick: X$1,
                        children: i.jsx(LocalizedText, {
                          children: "Add table",
                        }),
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "lg-model-help",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Values are never converted silently. Each table supports 500 records and 2,000 cells including headers.",
                    }),
                  }),
                  savedReport &&
                    props.workbook.dataModel &&
                    i.jsx("button", {
                      onClick: () => {
                        (onChange(structuredClone(props.workbook.dataModel)),
                          updateQuery({
                            groupBy: [],
                            filters: [],
                            measureIds: props.workbook.dataModel.measures.map(
                              (measure) => measure.id,
                            ),
                          }));
                      },
                      children: i.jsx(LocalizedText, {
                        children: "Use workbook model",
                      }),
                    }),
                  props.workbook.dataModel &&
                    i.jsx("button", {
                      disabled:
                        props.readOnly ||
                        props.workbook.structureProtection?.enabled,
                      onClick: () => {
                        try {
                          const latestProps = propsRef.current;
                          (latestProps.onWorkbookChange(
                            removeDataModelFromWorkbook(latestProps.workbook, {
                              readOnly: latestProps.readOnly,
                            }),
                            "Remove relational data model",
                          ),
                            latestProps.onClose());
                        } catch (error) {
                          setErrorMessage(error.message);
                        }
                      },
                      children: i.jsx(LocalizedText, {
                        children: "Remove saved model",
                      }),
                    }),
                  definition.tables.map((table) =>
                    i.jsxs(
                      "fieldset",
                      {
                        children: [
                          i.jsx("legend", { children: table.name }),
                          i.jsxs("div", {
                            className: "lg-model-controls",
                            children: [
                              LabeledField(
                                "Table name",
                                renderTextInput(
                                  `${table.name} table name`,
                                  table.name,
                                  (nextName) =>
                                    updateTable(table.id, { name: nextName }),
                                ),
                              ),
                              i.jsxs("span", {
                                children: [
                                  props.workbook.sheets.find(
                                    (sheet) => sheet.id === table.sheetId,
                                  )?.name ?? text("Missing worksheet"),
                                  " ",
                                  "· ",
                                  table.range,
                                  table.tableId
                                    ? " · Linked structured table"
                                    : "",
                                ],
                              }),
                              !table.tableId &&
                                LabeledField(
                                  "Headed range",
                                  renderTextInput(
                                    `${table.name} source range`,
                                    table.range,
                                    (range) =>
                                      updateTable(table.id, { range: range }),
                                  ),
                                ),
                              i.jsx("button", {
                                "aria-label": `Remove table ${table.name}`,
                                onClick: () => deleteTable(table.id),
                                children: i.jsx(LocalizedText, {
                                  children: "Remove table",
                                }),
                              }),
                            ],
                          }),
                          i.jsxs("details", {
                            children: [
                              i.jsxs("summary", {
                                children: [
                                  table.columns.length,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children: "typed columns",
                                  }),
                                ],
                              }),
                              i.jsxs("table", {
                                children: [
                                  i.jsx("thead", {
                                    children: i.jsx("tr", {
                                      children: [
                                        "Column",
                                        "Type",
                                        "Nullable",
                                      ].map((headerLabel) =>
                                        i.jsx(
                                          "th",
                                          {
                                            children: i.jsx(LocalizedText, {
                                              children: headerLabel,
                                            }),
                                          },
                                          headerLabel,
                                        ),
                                      ),
                                    }),
                                  }),
                                  i.jsx("tbody", {
                                    children: table.columns.map(
                                      (column, columnIndex) =>
                                        i.jsxs(
                                          "tr",
                                          {
                                            children: [
                                              i.jsx("th", {
                                                scope: "row",
                                                children: column.name,
                                              }),
                                              i.jsx("td", {
                                                children: renderSelect(
                                                  `${table.name}.${column.name} type`,
                                                  column.type,
                                                  toStringOptions([
                                                    "text",
                                                    "number",
                                                    "boolean",
                                                    "date",
                                                  ]),
                                                  (nextType) =>
                                                    updateTable(table.id, {
                                                      columns:
                                                        table.columns.map(
                                                          (
                                                            column,
                                                            itemIndex,
                                                          ) =>
                                                            itemIndex ===
                                                            columnIndex
                                                              ? {
                                                                  ...column,
                                                                  type: nextType,
                                                                }
                                                              : column,
                                                        ),
                                                    }),
                                                ),
                                              }),
                                              i.jsx("td", {
                                                children: i.jsx("input", {
                                                  type: "checkbox",
                                                  "aria-label": `${table.name}.${column.name} nullable`,
                                                  checked: column.nullable,
                                                  onChange: (event) =>
                                                    updateTable(table.id, {
                                                      columns:
                                                        table.columns.map(
                                                          (
                                                            column,
                                                            itemIndex,
                                                          ) =>
                                                            itemIndex ===
                                                            columnIndex
                                                              ? {
                                                                  ...column,
                                                                  nullable:
                                                                    event.target
                                                                      .checked,
                                                                }
                                                              : column,
                                                        ),
                                                    }),
                                                }),
                                              }),
                                            ],
                                          },
                                          column.name,
                                        ),
                                    ),
                                  }),
                                ],
                              }),
                            ],
                          }),
                        ],
                      },
                      table.id,
                    ),
                  ),
                ],
              }),
            activeSection === "Relationships" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("p", {
                    className: "lg-model-help",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Filters flow from the unique one side to the related side. Missing, blank or duplicate keys block calculation. Cycles and competing paths are refused.",
                    }),
                  }),
                  definition.relationships.map(
                    (relationship, relationshipIndex) => {
                      const updateRelationship = (nextRelationship) =>
                        onChange({
                          ...definition,
                          relationships: definition.relationships.map(
                            (existingRelationship) =>
                              existingRelationship.id === relationship.id
                                ? nextRelationship
                                : existingRelationship,
                          ),
                        });
                      return i.jsxs(
                        "fieldset",
                        {
                          children: [
                            i.jsxs("legend", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Relationship",
                                }),
                                " ",
                                relationshipIndex + 1,
                              ],
                            }),
                            i.jsxs("div", {
                              className: "lg-model-controls",
                              children: [
                                LabeledField(
                                  "One table",
                                  renderTableSelect(
                                    `Relationship ${relationshipIndex + 1} one table`,
                                    relationship.one.tableId,
                                    (tableId) =>
                                      updateRelationship({
                                        ...relationship,
                                        one: {
                                          tableId: tableId,
                                          column: defaultColumnName(tableId),
                                        },
                                      }),
                                  ),
                                ),
                                LabeledField(
                                  "Unique key",
                                  renderColumnSelect(
                                    `Relationship ${relationshipIndex + 1} one key`,
                                    relationship.one.tableId,
                                    relationship.one.column,
                                    (column) =>
                                      updateRelationship({
                                        ...relationship,
                                        one: {
                                          ...relationship.one,
                                          column: column,
                                        },
                                      }),
                                  ),
                                ),
                                i.jsx("span", {
                                  "aria-hidden": "true",
                                  children: "1 → *",
                                }),
                                LabeledField(
                                  "Related table",
                                  renderTableSelect(
                                    `Relationship ${relationshipIndex + 1} related table`,
                                    relationship.many.tableId,
                                    (tableId) =>
                                      updateRelationship({
                                        ...relationship,
                                        many: {
                                          tableId: tableId,
                                          column: defaultColumnName(tableId),
                                        },
                                      }),
                                  ),
                                ),
                                LabeledField(
                                  "Foreign key",
                                  renderColumnSelect(
                                    `Relationship ${relationshipIndex + 1} foreign key`,
                                    relationship.many.tableId,
                                    relationship.many.column,
                                    (column) =>
                                      updateRelationship({
                                        ...relationship,
                                        many: {
                                          ...relationship.many,
                                          column: column,
                                        },
                                      }),
                                  ),
                                ),
                                LabeledField(
                                  "Cardinality",
                                  renderSelect(
                                    `Relationship ${relationshipIndex + 1} cardinality`,
                                    relationship.cardinality,
                                    [
                                      {
                                        value: "one-to-many",
                                        label: "1 → many",
                                      },
                                      { value: "one-to-one", label: "1 → 1" },
                                    ],
                                    (cardinality) =>
                                      updateRelationship({
                                        ...relationship,
                                        cardinality: cardinality,
                                      }),
                                  ),
                                ),
                                i.jsx("button", {
                                  "aria-label": `Remove relationship ${relationshipIndex + 1}`,
                                  onClick: () =>
                                    onChange({
                                      ...definition,
                                      relationships:
                                        definition.relationships.filter(
                                          (otherRelationship) =>
                                            otherRelationship.id !==
                                            relationship.id,
                                        ),
                                    }),
                                  children: i.jsx(LocalizedText, {
                                    children: "Remove",
                                  }),
                                }),
                              ],
                            }),
                          ],
                        },
                        relationship.id,
                      );
                    },
                  ),
                  i.jsx("button", {
                    disabled:
                      definition.tables.length < 2 ||
                      definition.relationships.length >=
                        dataModelLimits.relationships,
                    onClick: () =>
                      onChange({
                        ...definition,
                        relationships: [
                          ...definition.relationships,
                          {
                            id: createUuid(),
                            one: {
                              tableId: definition.tables[0].id,
                              column: definition.tables[0].columns[0].name,
                            },
                            many: {
                              tableId: definition.tables[1].id,
                              column: definition.tables[1].columns[0].name,
                            },
                            direction: "single",
                            cardinality: "one-to-many",
                          },
                        ],
                      }),
                    children: i.jsx(LocalizedText, {
                      children: "Add relationship",
                    }),
                  }),
                ],
              }),
            activeSection === "Calculated columns" &&
              i.jsx(DataModelDefinitionEditor, {
                definition: definition,
                onChange: onChange,
              }),
            activeSection === "Measures" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("p", {
                    className: "lg-model-help",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Each measure aggregates its own filtered table, preventing fact-row multiplication. Choose DAX for expressions with row and filter context. Totals are recalculated in their own context.",
                    }),
                  }),
                  definition.measures.map((measure, measureIndex) => {
                    const updateMeasure = (nextMeasure) =>
                      onChange({
                        ...definition,
                        measures: definition.measures.map((existingMeasure) =>
                          existingMeasure.id === measure.id
                            ? nextMeasure
                            : existingMeasure,
                        ),
                      });
                    return i.jsxs(
                      "fieldset",
                      {
                        children: [
                          i.jsxs("legend", {
                            children: [
                              i.jsx(LocalizedText, { children: "Measure" }),
                              " ",
                              measureIndex + 1,
                            ],
                          }),
                          i.jsxs("div", {
                            className: "lg-model-controls",
                            children: [
                              LabeledField(
                                "Name",
                                renderTextInput(
                                  `Measure ${measureIndex + 1} name`,
                                  measure.name,
                                  (nextName) =>
                                    updateMeasure({
                                      ...measure,
                                      name: nextName,
                                    }),
                                ),
                              ),
                              LabeledField(
                                "Table",
                                renderTableSelect(
                                  `Measure ${measureIndex + 1} table`,
                                  measure.tableId,
                                  (tableId) =>
                                    updateMeasure({
                                      ...measure,
                                      tableId: tableId,
                                      ...(measure.operation === "countRows" ||
                                      measure.operation === "dax"
                                        ? {}
                                        : {
                                            column: defaultColumnName(tableId),
                                          }),
                                    }),
                                ),
                              ),
                              LabeledField(
                                "Operation",
                                renderSelect(
                                  `Measure ${measureIndex + 1} operation`,
                                  measure.operation,
                                  toStringOptions([
                                    "sum",
                                    "count",
                                    "countRows",
                                    "average",
                                    "min",
                                    "max",
                                    "distinctCount",
                                    "dax",
                                  ]),
                                  (operation) => {
                                    const updatedMeasure = {
                                      ...measure,
                                      operation: operation,
                                    };
                                    (operation === "countRows" ||
                                    operation === "dax"
                                      ? delete updatedMeasure.column
                                      : (updatedMeasure.column =
                                          updatedMeasure.column ??
                                          defaultColumnName(measure.tableId)),
                                      operation === "dax"
                                        ? (updatedMeasure.expression =
                                            updatedMeasure.expression ??
                                            "BLANK()")
                                        : delete updatedMeasure.expression,
                                      updateMeasure(updatedMeasure));
                                  },
                                ),
                              ),
                              measure.operation !== "countRows" &&
                                measure.operation !== "dax" &&
                                LabeledField(
                                  "Column",
                                  renderColumnSelect(
                                    `Measure ${measureIndex + 1} column`,
                                    measure.tableId,
                                    measure.column ?? "",
                                    (column) =>
                                      updateMeasure({
                                        ...measure,
                                        column: column,
                                      }),
                                  ),
                                ),
                              i.jsx("button", {
                                "aria-label": `Remove measure ${measureIndex + 1}`,
                                onClick: () => {
                                  (onChange({
                                    ...definition,
                                    measures: definition.measures.filter(
                                      (otherMeasure) =>
                                        otherMeasure.id !== measure.id,
                                    ),
                                  }),
                                    updateQuery({
                                      ...query,
                                      measureIds: query.measureIds.filter(
                                        (measureId) => measureId !== measure.id,
                                      ),
                                    }));
                                },
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            ],
                          }),
                          measure.operation === "dax" &&
                            i.jsx(DaxExpressionInput, {
                              label: `Measure ${measureIndex + 1} DAX expression`,
                              value: measure.expression ?? "",
                              onChange: (expression) =>
                                updateMeasure({
                                  ...measure,
                                  expression: expression,
                                }),
                            }),
                        ],
                      },
                      measure.id,
                    );
                  }),
                  i.jsx("button", {
                    disabled:
                      !definition.tables.length ||
                      definition.measures.length >= dataModelLimits.measures,
                    onClick: () => {
                      const newMeasure = {
                        id: createUuid(),
                        name: `Measure ${definition.measures.length + 1}`,
                        tableId: definition.tables.at(-1).id,
                        operation: "countRows",
                      };
                      (onChange({
                        ...definition,
                        measures: [...definition.measures, newMeasure],
                      }),
                        updateQuery({
                          ...query,
                          measureIds: [...query.measureIds, newMeasure.id],
                        }));
                    },
                    children: i.jsx(LocalizedText, { children: "Add measure" }),
                  }),
                ],
              }),
            activeSection === "Query" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Selected measures",
                        }),
                      }),
                      i.jsx("div", {
                        className: "lg-model-controls",
                        children: definition.measures.map((measure) =>
                          i.jsxs(
                            "label",
                            {
                              children: [
                                i.jsx("input", {
                                  type: "checkbox",
                                  checked: query.measureIds.includes(
                                    measure.id,
                                  ),
                                  onChange: (event) =>
                                    updateQuery({
                                      ...query,
                                      measureIds: event.target.checked
                                        ? [...query.measureIds, measure.id]
                                        : query.measureIds.filter(
                                            (measureId) =>
                                              measureId !== measure.id,
                                          ),
                                    }),
                                }),
                                measure.name,
                              ],
                            },
                            measure.id,
                          ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Group by shared dimensions",
                        }),
                      }),
                      query.groupBy.map((groupField, groupIndex) =>
                        i.jsxs(
                          "div",
                          {
                            className: "lg-model-controls",
                            children: [
                              renderTableSelect(
                                `Group ${groupIndex + 1} table`,
                                groupField.tableId,
                                (tableId) =>
                                  updateQuery({
                                    ...query,
                                    groupBy: query.groupBy.map(
                                      (existingGroup, candidateIndex) =>
                                        candidateIndex === groupIndex
                                          ? {
                                              tableId: tableId,
                                              column:
                                                defaultColumnName(tableId),
                                            }
                                          : existingGroup,
                                    ),
                                  }),
                              ),
                              renderColumnSelect(
                                `Group ${groupIndex + 1} column`,
                                groupField.tableId,
                                groupField.column,
                                (column) =>
                                  updateQuery({
                                    ...query,
                                    groupBy: query.groupBy.map(
                                      (existingGroup, candidateIndex) =>
                                        candidateIndex === groupIndex
                                          ? { ...existingGroup, column: column }
                                          : existingGroup,
                                    ),
                                  }),
                              ),
                              i.jsx("button", {
                                "aria-label": `Remove group ${groupIndex + 1}`,
                                onClick: () =>
                                  updateQuery({
                                    ...query,
                                    groupBy: query.groupBy.filter(
                                      (group, entryIndex) =>
                                        entryIndex !== groupIndex,
                                    ),
                                  }),
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            ],
                          },
                          groupIndex,
                        ),
                      ),
                      i.jsx("button", {
                        disabled:
                          !definition.tables.length ||
                          query.groupBy.length >= 3,
                        onClick: () =>
                          updateQuery({
                            ...query,
                            groupBy: [
                              ...query.groupBy,
                              {
                                tableId: definition.tables[0].id,
                                column: definition.tables[0].columns[0].name,
                              },
                            ],
                          }),
                        children: i.jsx(LocalizedText, {
                          children: "Add group field",
                        }),
                      }),
                    ],
                  }),
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Filters (all must match)",
                        }),
                      }),
                      query.filters.map((filter, filterIndex) => {
                        const updateFilter = (nextFilter) =>
                            updateQuery({
                              ...query,
                              filters: query.filters.map((filter, itemIndex) =>
                                itemIndex === filterIndex ? nextFilter : filter,
                              ),
                            }),
                          columnType = findTable(filter.tableId)?.columns.find(
                            (column) => column.name === filter.column,
                          )?.type;
                        return i.jsxs(
                          "div",
                          {
                            className: "lg-model-controls",
                            children: [
                              renderTableSelect(
                                `Filter ${filterIndex + 1} table`,
                                filter.tableId,
                                (tableId) =>
                                  updateFilter({
                                    ...filter,
                                    tableId: tableId,
                                    column: defaultColumnName(tableId),
                                    operator: "blank",
                                    value: undefined,
                                  }),
                              ),
                              renderColumnSelect(
                                `Filter ${filterIndex + 1} column`,
                                filter.tableId,
                                filter.column,
                                (column) =>
                                  updateFilter({
                                    ...filter,
                                    column: column,
                                    operator: "blank",
                                    value: undefined,
                                  }),
                              ),
                              renderSelect(
                                `Filter ${filterIndex + 1} operator`,
                                filter.operator,
                                toStringOptions([
                                  "equals",
                                  "notEquals",
                                  "gt",
                                  "gte",
                                  "lt",
                                  "lte",
                                  "blank",
                                  "notBlank",
                                ]),
                                (operator) =>
                                  updateFilter({
                                    ...filter,
                                    operator: operator,
                                    value: ["blank", "notBlank"].includes(
                                      operator,
                                    )
                                      ? undefined
                                      : columnType === "number"
                                        ? 0
                                        : columnType === "boolean"
                                          ? false
                                          : "",
                                  }),
                              ),
                              !["blank", "notBlank"].includes(
                                filter.operator,
                              ) &&
                                (columnType === "boolean"
                                  ? renderSelect(
                                      `Filter ${filterIndex + 1} value`,
                                      String(filter.value),
                                      toStringOptions(["true", "false"]),
                                      (nextValue) =>
                                        updateFilter({
                                          ...filter,
                                          value: nextValue === "true",
                                        }),
                                    )
                                  : renderTextInput(
                                      `Filter ${filterIndex + 1} value`,
                                      String(filter.value ?? ""),
                                      (nextText) =>
                                        updateFilter({
                                          ...filter,
                                          value:
                                            columnType === "number"
                                              ? nextText.trim()
                                                ? Number(nextText)
                                                : null
                                              : nextText,
                                        }),
                                    )),
                              i.jsx("button", {
                                "aria-label": `Remove filter ${filterIndex + 1}`,
                                onClick: () =>
                                  updateQuery({
                                    ...query,
                                    filters: query.filters.filter(
                                      (filter, candidateIndex) =>
                                        candidateIndex !== filterIndex,
                                    ),
                                  }),
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            ],
                          },
                          filterIndex,
                        );
                      }),
                      i.jsx("button", {
                        disabled:
                          !definition.tables.length ||
                          query.filters.length >= dataModelLimits.filters,
                        onClick: () =>
                          updateQuery({
                            ...query,
                            filters: [
                              ...query.filters,
                              {
                                tableId: definition.tables[0].id,
                                column: definition.tables[0].columns[0].name,
                                operator: "notBlank",
                              },
                            ],
                          }),
                        children: i.jsx(LocalizedText, {
                          children: "Add model filter",
                        }),
                      }),
                    ],
                  }),
                ],
              }),
            activeSection === "Preview" &&
              (reportResult
                ? i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("div", {
                        role: "status",
                        "aria-label": text("Model preview status"),
                        className: "lg-model-summary",
                        children: [
                          reportResult.complete
                            ? text("Model checks passed")
                            : text("Model blocked"),
                          " ",
                          "· ",
                          reportResult.tables.length,
                          " ",
                          text("tables"),
                          " ·",
                          " ",
                          reportResult.operations.toLocaleString(),
                          " ",
                          text("operations"),
                        ],
                      }),
                      reportResult.diagnostics.length > 0 &&
                        i.jsxs("table", {
                          "aria-label": text("Model diagnostics"),
                          children: [
                            i.jsx("thead", {
                              children: i.jsxs("tr", {
                                children: [
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Code",
                                    }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Source row",
                                    }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Diagnostic",
                                    }),
                                  }),
                                ],
                              }),
                            }),
                            i.jsx("tbody", {
                              children: reportResult.diagnostics.map(
                                (diagnostic, diagnosticIndex) =>
                                  i.jsxs(
                                    "tr",
                                    {
                                      children: [
                                        i.jsx("td", {
                                          children: diagnostic.code,
                                        }),
                                        i.jsxs("td", {
                                          children: [
                                            findTable(diagnostic.tableId ?? "")
                                              ?.name,
                                            " ",
                                            diagnostic.row ?? "",
                                          ],
                                        }),
                                        i.jsx("td", {
                                          children: text(diagnostic.message),
                                        }),
                                      ],
                                    },
                                    diagnosticIndex,
                                  ),
                              ),
                            }),
                          ],
                        }),
                      i.jsxs("table", {
                        "aria-label": text("Model source controls"),
                        children: [
                          i.jsx("thead", {
                            children: i.jsx("tr", {
                              children: [
                                "Table",
                                "Source",
                                "Records",
                                "Included",
                                "Excluded",
                                "Column storage (bytes)",
                              ].map((headerLabel) =>
                                i.jsx(
                                  "th",
                                  {
                                    children: i.jsx(LocalizedText, {
                                      children: headerLabel,
                                    }),
                                  },
                                  headerLabel,
                                ),
                              ),
                            }),
                          }),
                          i.jsx("tbody", {
                            children: reportResult.tables.map((table) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("th", {
                                      scope: "row",
                                      children: table.name,
                                    }),
                                    i.jsxs("td", {
                                      children: [
                                        props.workbook.sheets.find(
                                          (sheet) => sheet.id === table.sheetId,
                                        )?.name,
                                        "!",
                                        table.range,
                                      ],
                                    }),
                                    i.jsx("td", { children: table.rows }),
                                    i.jsx("td", { children: table.included }),
                                    i.jsx("td", { children: table.excluded }),
                                    i.jsx("td", {
                                      children: table.columnarBytes,
                                    }),
                                  ],
                                },
                                table.id,
                              ),
                            ),
                          }),
                        ],
                      }),
                      reportResult.complete &&
                        !!reportResult.calculatedColumns?.length &&
                        i.jsxs("details", {
                          children: [
                            i.jsx("summary", {
                              children: i.jsx(LocalizedText, {
                                children: "Calculated column preview",
                              }),
                            }),
                            i.jsx("p", {
                              className: "lg-model-help",
                              children: i.jsx(LocalizedText, {
                                children:
                                  "First 25 source records per column. All records are evaluated before publication.",
                              }),
                            }),
                            i.jsxs("table", {
                              "aria-label": text("Calculated column results"),
                              children: [
                                i.jsx("thead", {
                                  children: i.jsxs("tr", {
                                    children: [
                                      i.jsx("th", {
                                        children: i.jsx(LocalizedText, {
                                          children: "Table and column",
                                        }),
                                      }),
                                      i.jsx("th", {
                                        children: i.jsx(LocalizedText, {
                                          children: "Values",
                                        }),
                                      }),
                                    ],
                                  }),
                                }),
                                i.jsx("tbody", {
                                  children: reportResult.calculatedColumns.map(
                                    (calculatedColumn) =>
                                      i.jsxs(
                                        "tr",
                                        {
                                          children: [
                                            i.jsxs("th", {
                                              scope: "row",
                                              children: [
                                                findTable(
                                                  calculatedColumn.tableId,
                                                )?.name,
                                                ".",
                                                calculatedColumn.name,
                                              ],
                                            }),
                                            i.jsx("td", {
                                              children: calculatedColumn.values
                                                .slice(0, 25)
                                                .map((calculatedValue) =>
                                                  calculatedValue === null
                                                    ? text("Blank")
                                                    : String(calculatedValue),
                                                )
                                                .join(", "),
                                            }),
                                          ],
                                        },
                                        `${calculatedColumn.tableId}:${calculatedColumn.name}`,
                                      ),
                                  ),
                                }),
                              ],
                            }),
                          ],
                        }),
                      reportResult.complete &&
                        query.measureIds.length > 0 &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("table", {
                              "aria-label": text("Model measure results"),
                              children: [
                                i.jsx("thead", {
                                  children: i.jsx("tr", {
                                    children: reportResult.headers.map(
                                      (header) =>
                                        i.jsx(
                                          "th",
                                          { children: header },
                                          header,
                                        ),
                                    ),
                                  }),
                                }),
                                i.jsx("tbody", {
                                  children: reportResult.rows
                                    .slice(
                                      resultPage * 25,
                                      (resultPage + 1) * 25,
                                    )
                                    .map((row, rowIndex) =>
                                      i.jsx(
                                        "tr",
                                        {
                                          children: [
                                            ...row.keys,
                                            ...row.values,
                                          ].map((cellValue, cellIndex) =>
                                            i.jsx(
                                              "td",
                                              {
                                                children:
                                                  cellValue === null
                                                    ? text("Blank")
                                                    : String(cellValue),
                                              },
                                              cellIndex,
                                            ),
                                          ),
                                        },
                                        rowIndex,
                                      ),
                                    ),
                                }),
                                query.groupBy.length > 0 &&
                                  i.jsx("tfoot", {
                                    children: i.jsxs("tr", {
                                      children: [
                                        i.jsx("th", {
                                          colSpan: query.groupBy.length,
                                          children: i.jsx(LocalizedText, {
                                            children: "Grand total",
                                          }),
                                        }),
                                        reportResult.totals.map(
                                          (total, totalIndex) =>
                                            i.jsx(
                                              "td",
                                              {
                                                children:
                                                  total === null
                                                    ? text("Blank")
                                                    : total,
                                              },
                                              totalIndex,
                                            ),
                                        ),
                                      ],
                                    }),
                                  }),
                              ],
                            }),
                            i.jsxs("div", {
                              className: "lg-model-controls",
                              children: [
                                i.jsx("button", {
                                  disabled: !resultPage,
                                  onClick: () =>
                                    setResultPage(
                                      (currentPage) => currentPage - 1,
                                    ),
                                  children: i.jsx(LocalizedText, {
                                    children: "Previous results",
                                  }),
                                }),
                                i.jsxs("span", {
                                  children: [
                                    resultPage + 1,
                                    " /",
                                    " ",
                                    Math.max(
                                      1,
                                      Math.ceil(reportResult.rows.length / 25),
                                    ),
                                  ],
                                }),
                                i.jsx("button", {
                                  disabled:
                                    (resultPage + 1) * 25 >=
                                    reportResult.rows.length,
                                  onClick: () =>
                                    setResultPage(
                                      (currentPage) => currentPage + 1,
                                    ),
                                  children: i.jsx(LocalizedText, {
                                    children: "Next results",
                                  }),
                                }),
                              ],
                            }),
                          ],
                        }),
                      i.jsx("p", {
                        className: "lg-model-help",
                        children: i.jsx(LocalizedText, {
                          children:
                            "Refresh publishes a new worksheet. Earlier results retain a freshness label and remain available.",
                        }),
                      }),
                    ],
                  })
                : i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Preview the model to inspect integrity, source counts and measures.",
                    }),
                  })),
          ],
        }),
        i.jsxs("footer", {
          children: [
            props.allowHidden &&
              i.jsxs("label", {
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
            LabeledField(
              "Output worksheet",
              renderTextInput(
                "Model output worksheet",
                outputName,
                (nextOutputName) => {
                  (setOutputName(nextOutputName), resetPreview());
                },
              ),
            ),
            i.jsx("button", {
              onClick: onClick,
              children: i.jsx(LocalizedText, { children: "Preview model" }),
            }),
            i.jsx("button", {
              disabled: !canPublish,
              onClick: () => commitModelReport("save"),
              children: i.jsx(LocalizedText, { children: "Save model" }),
            }),
            i.jsx("button", {
              className: "lg-model-primary",
              disabled: !canPublish || !query.measureIds.length,
              onClick: () => commitModelReport("publish"),
              children: i.jsx(LocalizedText, {
                children: "Publish model result",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
function DataModelSnapshotNotice({ workbook, sheetId, options, onReview }) {
  const reportStatus = ne.useMemo(
    () => getModelReportStatus(workbook, sheetId, options),
    [workbook, sheetId, options?.includeHidden, options?.allowedSheetIds],
  );
  return reportStatus.status === "none"
    ? null
    : i.jsxs("div", {
        className: "lg-notice",
        role: "status",
        "aria-label": "Data model snapshot status",
        children: [
          i.jsxs("span", {
            children: [
              i.jsx("strong", {
                children: i.jsx(LocalizedText, {
                  children: `Model snapshot · ${reportStatus.status}`,
                }),
              }),
              " ",
              "· ",
              i.jsx(LocalizedText, { children: reportStatus.message }),
            ],
          }),
          i.jsx("button", {
            onClick: onReview,
            children: i.jsx(LocalizedText, { children: "Review model" }),
          }),
        ],
      });
}
const focusableElementSelector =
  'button:not(:disabled),a[href],input:not(:disabled):not([type="hidden"]),select:not(:disabled),textarea:not(:disabled),summary,[tabindex]:not([tabindex="-1"])';
function isElementVisibleForFocus(element) {
  for (let ancestor = element; ancestor; ancestor = ancestor.parentElement) {
    const computedStyle = getComputedStyle(ancestor);
    if (
      ancestor.hidden ||
      computedStyle.display === "none" ||
      computedStyle.visibility === "hidden" ||
      computedStyle.visibility === "collapse"
    )
      return false;
    if (ancestor.tagName === "DETAILS" && !ancestor.hasAttribute("open")) {
      const summary = Array.from(ancestor.children).find(
        (child) => child.tagName === "SUMMARY",
      );
      if (!summary?.contains(element)) return false;
    }
  }
  return true;
}
function getFocusableElements(container) {
  return Array.from(
    container.querySelectorAll(focusableElementSelector),
  ).filter(
    (element) =>
      !(
        element.tabIndex < 0 ||
        element.matches(":disabled") ||
        !isElementVisibleForFocus(element) ||
        element.closest('[hidden],[inert],[aria-hidden="true"]')
      ),
  );
}
function useDialogFocusTrap(appRootRef) {
  const syncFocusRef = ne.useRef(undefined);
  (ne.useLayoutEffect(() => {
    const appRoot = appRootRef.current;
    if (!appRoot) return;
    let activeDialog,
      focusBeforeDialog = null;
    const focusByDialog = new Map();
    let lastFocusedElement =
      document.activeElement instanceof HTMLElement
        ? document.activeElement
        : null;
    const inertedSiblings = new Map(),
      restoreInertSiblings = () => {
        for (const [siblingElement, savedAttributes] of inertedSiblings)
          (savedAttributes.inert === null
            ? siblingElement.removeAttribute("inert")
            : siblingElement.setAttribute("inert", savedAttributes.inert),
            savedAttributes.hidden === null
              ? siblingElement.removeAttribute("aria-hidden")
              : siblingElement.setAttribute(
                  "aria-hidden",
                  savedAttributes.hidden,
                ));
        inertedSiblings.clear();
      },
      syncFocusTrap = () => {
        const nextDialog = Array.from(
          appRoot.querySelectorAll('[role="dialog"][aria-modal="true"]'),
        )
          .filter(isElementVisibleForFocus)
          .at(-1);
        if (nextDialog === activeDialog) {
          nextDialog &&
            (!isElementVisibleForFocus(document.activeElement) ||
              document.activeElement?.matches(":disabled")) &&
            (getFocusableElements(nextDialog)[0] ?? nextDialog).focus({
              preventScroll: true,
            });
          return;
        }
        const previousDialog = activeDialog;
        if ((restoreInertSiblings(), (activeDialog = nextDialog), nextDialog)) {
          previousDialog || (focusBeforeDialog = lastFocusedElement);
          const hasSavedFocus = focusByDialog.has(nextDialog);
          hasSavedFocus || focusByDialog.set(nextDialog, lastFocusedElement);
          let ancestor = nextDialog;
          for (; ancestor !== appRoot && ancestor.parentElement;) {
            for (const sibling of Array.from(ancestor.parentElement.children))
              sibling instanceof HTMLElement &&
                sibling !== ancestor &&
                (inertedSiblings.set(sibling, {
                  inert: sibling.getAttribute("inert"),
                  hidden: sibling.getAttribute("aria-hidden"),
                }),
                sibling.setAttribute("inert", ""),
                sibling.setAttribute("aria-hidden", "true"));
            ancestor = ancestor.parentElement;
          }
          const previousDialogFocus = previousDialog
            ? focusByDialog.get(previousDialog)
            : undefined;
          if (!nextDialog.contains(document.activeElement)) {
            const focusTarget =
              hasSavedFocus &&
              previousDialogFocus &&
              nextDialog.contains(previousDialogFocus) &&
              isElementVisibleForFocus(previousDialogFocus) &&
              !previousDialogFocus.matches(":disabled")
                ? previousDialogFocus
                : (getFocusableElements(nextDialog)[0] ?? nextDialog);
            (focusTarget === nextDialog && (nextDialog.tabIndex = -1),
              focusTarget.focus({ preventScroll: true }));
          }
        } else if (
          previousDialog &&
          (document.activeElement === document.body ||
            previousDialog.contains(document.activeElement))
        ) {
          const restoreTarget =
            focusBeforeDialog?.isConnected &&
            isElementVisibleForFocus(focusBeforeDialog) &&
            !focusBeforeDialog.matches(":disabled")
              ? focusBeforeDialog
              : appRoot.querySelector(".lg-grid-scroll");
          restoreTarget?.focus({ preventScroll: true });
        }
        (previousDialog &&
          (!previousDialog.isConnected ||
            !isElementVisibleForFocus(previousDialog)) &&
          focusByDialog.delete(previousDialog),
          nextDialog || focusByDialog.clear());
      },
      handleFocusIn = (event) => {
        const target = event.target;
        if (target instanceof HTMLElement) {
          if (
            activeDialog &&
            appRoot.contains(target) &&
            !activeDialog.contains(target)
          ) {
            (getFocusableElements(activeDialog)[0] ?? activeDialog).focus({
              preventScroll: true,
            });
            return;
          }
          ((!activeDialog &&
            !target.closest('[role="dialog"][aria-modal="true"]')) ||
            activeDialog?.contains(target)) &&
            (lastFocusedElement = target);
        }
      },
      handleKeyDown = (event) => {
        if ((syncFocusTrap(), !activeDialog || event.key !== "Tab")) return;
        const focusableElements = getFocusableElements(activeDialog),
          firstFocusable = focusableElements[0],
          lastFocusable = focusableElements.at(-1);
        if (!firstFocusable) {
          (event.preventDefault(),
            (activeDialog.tabIndex = -1),
            activeDialog.focus());
          return;
        }
        (!activeDialog.contains(document.activeElement) ||
          (event.shiftKey && document.activeElement === firstFocusable) ||
          (!event.shiftKey && document.activeElement === lastFocusable)) &&
          (event.preventDefault(),
          event.stopPropagation(),
          (event.shiftKey ? lastFocusable : firstFocusable)?.focus());
      };
    (appRoot.addEventListener("keydown", handleKeyDown, true),
      document.addEventListener("focusin", handleFocusIn, true));
    const observer = new MutationObserver(syncFocusTrap);
    return (
      observer.observe(appRoot, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ["hidden", "style", "class", "open", "disabled"],
      }),
      syncFocusTrap(),
      (syncFocusRef.current = syncFocusTrap),
      () => {
        (observer.disconnect(),
          appRoot.removeEventListener("keydown", handleKeyDown, true),
          document.removeEventListener("focusin", handleFocusIn, true),
          restoreInertSiblings(),
          (syncFocusRef.current = undefined));
      }
    );
  }, [appRootRef]),
    ne.useLayoutEffect(() => syncFocusRef.current?.()));
}
const dataModelBudgets = {
  sourceCells: 500,
  outputCells: 500,
  fields: 50,
  sourceCharacters: 2e5,
  formulaCells: 1e3,
  formulaCharacters: 1e5,
  nameFormulaCharacters: 1e6,
  discoveryAnchors: 50,
  dependencyCells: 2e3,
};

function assertRecordKeys(destination, allowedKeys, errorMessage) {
  (!isNonArrayObject(destination) ||
    Object.keys(destination).some((key) => !allowedKeys.includes(key))) &&
    failWithMessage(errorMessage);
}
const rangesOverlap = (firstRange, secondRange) => {
  const rangeAddress = parseRangeAddress(firstRange),
    secondAddress = parseRangeAddress(secondRange);
  return (
    rangeAddress.firstRow <= secondAddress.lastRow &&
    rangeAddress.lastRow >= secondAddress.firstRow &&
    rangeAddress.firstCol <= secondAddress.lastCol &&
    rangeAddress.lastCol >= secondAddress.firstCol
  );
};
function assertDelimiterCharacter(delimiter) {
  (typeof delimiter != "string" ||
    [...delimiter].length !== 1 ||
    /["\r\n\0]/.test(delimiter)) &&
    failWithMessage(
      "Choose one delimiter character other than a quote, line break or NUL.",
    );
}
function splitDelimitedCell(cellTextValue, delimiter) {
  (assertDelimiterCharacter(delimiter),
    (typeof cellTextValue != "string" || cellTextValue.length > 32767) &&
      failWithMessage(
        "A source cell must contain text of at most 32,767 characters.",
      ));
  const fields = [];
  let fieldText = "",
    inQuotes = false,
    afterClosingQuote = false,
    fieldStarted = false;
  const characters = [...cellTextValue],
    endField = () => {
      (fields.push(fieldText),
        fields.length > dataModelBudgets.fields &&
          failWithMessage(
            "A source cell contains more than 50 fields. Choose a different delimiter or split a smaller record.",
          ),
        (fieldText = ""),
        (fieldStarted = false),
        (afterClosingQuote = false));
    };
  for (
    let characterIndex = 0;
    characterIndex < characters.length;
    characterIndex++
  ) {
    const character = characters[characterIndex];
    inQuotes
      ? character !== '"'
        ? (fieldText += character)
        : characters[characterIndex + 1] === '"'
          ? ((fieldText += '"'), characterIndex++)
          : ((inQuotes = false), (afterClosingQuote = true))
      : character === delimiter
        ? endField()
        : afterClosingQuote
          ? failWithMessage(
              "Unexpected text after a closing quote. A closing quote must be followed by the delimiter or end of the cell.",
            )
          : character === '"'
            ? (fieldStarted &&
                failWithMessage(
                  "A quote must begin a field; escape embedded quotes by doubling them inside a quoted field.",
                ),
              (inQuotes = true),
              (fieldStarted = true))
            : ((fieldText += character), (fieldStarted = true));
  }
  return (
    inQuotes &&
      failWithMessage(
        "A quoted field is not closed. Check the source cell before splitting.",
      ),
    endField(),
    fields
  );
}
function validateTextToColumnsRequest(request) {
  (assertRecordKeys(
    request,
    [
      "sourceSheetId",
      "sourceRange",
      "delimiter",
      "header",
      "formats",
      "destination",
    ],
    "Invalid Text to Columns request.",
  ),
    (typeof request.sourceSheetId != "string" ||
      typeof request.sourceRange != "string") &&
      failWithMessage("Choose a source worksheet and one-column range."),
    assertDelimiterCharacter(request.delimiter),
    ["none", "first-row"].includes(request.header) ||
      failWithMessage("Choose whether the first source row contains headings."),
    request.formats !== undefined &&
      (!Array.isArray(request.formats) ||
        request.formats.length > dataModelBudgets.fields ||
        request.formats.some(
          (format) => !["text", "general"].includes(format),
        )) &&
      failWithMessage(
        "Column conversion must be Text or General, with at most 50 columns.",
      ),
    assertRecordKeys(
      request.destination,
      ["kind", "name", "sheetId", "address"],
      "Choose a separate output range or a new worksheet.",
    ));
  let destination;
  if (request.destination.kind === "new-sheet") {
    (assertRecordKeys(
      request.destination,
      ["kind", "name"],
      "Invalid new worksheet destination.",
    ),
      typeof request.destination.name != "string" &&
        failWithMessage("Enter an output worksheet name."));
    const trimmedName = request.destination.name.trim();
    ((!trimmedName ||
      trimmedName.length > 31 ||
      /[\u0000-\u001f\u007f\\/\[\]*?:]/.test(trimmedName) ||
      trimmedName.startsWith("'") ||
      trimmedName.endsWith("'")) &&
      failWithMessage(
        "Output worksheet names need 1–31 characters without control/Excel-reserved characters or leading/trailing apostrophes.",
      ),
      (destination = { kind: "new-sheet", name: trimmedName }));
  } else if (request.destination.kind === "range") {
    (assertRecordKeys(
      request.destination,
      ["kind", "sheetId", "address"],
      "Invalid output range destination.",
    ),
      (typeof request.destination.sheetId != "string" ||
        typeof request.destination.address != "string") &&
        failWithMessage("Choose an output worksheet and one starting cell."));
    const startCell = _t(request.destination.address.trim());
    destination = {
      kind: "range",
      sheetId: request.destination.sheetId,
      address: formatCellAddress(startCell.row, startCell.col),
    };
  } else failWithMessage("Choose a separate output range or a new worksheet.");
  return {
    sourceSheetId: request.sourceSheetId,
    sourceRange: formatRangeCornerCells(request.sourceRange),
    delimiter: request.delimiter,
    header: request.header,
    ...(request.formats ? { formats: [...request.formats] } : {}),
    destination: destination,
  };
}
function validateTextToColumnsPermissions(workbook, permissions) {
  return (
    [
      permissions.readOnly,
      permissions.includeHidden,
      permissions.allowNewSheet,
    ].some((flag) => flag !== undefined && typeof flag != "boolean") &&
      failWithMessage("Text to Columns permissions must be boolean."),
    permissions.readOnly && failWithMessage("This workbook is read-only."),
    permissions.allowedSheetIds !== undefined &&
      (!Array.isArray(permissions.allowedSheetIds) ||
        permissions.allowedSheetIds.some(
          (allowedSheetId) =>
            typeof allowedSheetId != "string" ||
            !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
        )) &&
      failWithMessage("The authorized worksheet context is invalid."),
    {
      includeHidden: permissions.includeHidden ?? false,
      allowNewSheet: permissions.allowNewSheet ?? true,
      allowedSheetIds: [
        ...new Set(
          permissions.allowedSheetIds ??
            workbook.sheets.map((sheet) => sheet.id),
        ),
      ].sort(),
    }
  );
}
function assertSelectionInAuthorizedSheet(sheet, rangeText, context) {
  const rangeAddress = parseRangeAddress(rangeText);
  ((!context.allowedSheetIds.includes(sheet.id) ||
    (!context.includeHidden &&
      (sheet.hidden ||
        sheet.veryHidden ||
        cn(sheet, "row").some(
          (rowSpan) =>
            rowSpan.start <= rangeAddress.lastRow &&
            rowSpan.end >= rangeAddress.firstRow,
        ) ||
        cn(sheet, "column").some(
          (columnSpan) =>
            columnSpan.start <= rangeAddress.lastCol &&
            columnSpan.end >= rangeAddress.firstCol,
        )))) &&
    failWithMessage(
      "The selected cells are outside the visible, authorized worksheet context.",
    ),
    (rangeAddress.lastRow >= sheet.rowCount ||
      rangeAddress.lastCol >= sheet.colCount) &&
      failWithMessage("The selected cells exceed the worksheet boundary."));
}
function formatCellTextValue(cellValue) {
  return cellValue === null
    ? ""
    : typeof cellValue == "boolean"
      ? cellValue
        ? "TRUE"
        : "FALSE"
      : String(cellValue);
}
function stripFormulaLiterals(formula) {
  let result = "",
    position = 0;
  for (; position < formula.length;) {
    const tokenStart = position,
      character = formula[position++];
    if (character === '"' || character === "'") {
      for (; position < formula.length;)
        if (formula[position++] === character)
          if (formula[position] === character) position++;
          else break;
      result += character === '"' ? '""' : formula.slice(tokenStart, position);
    } else if (character === "[") {
      let bracketDepth = 1;
      for (; position < formula.length && bracketDepth;)
        (formula[position] === "[" && bracketDepth++,
          formula[position] === "]" && bracketDepth--,
          position++);
      result +=
        !bracketDepth &&
        !/^[\p{L}\p{M}\p{N}_.!\\]/u.test(formula.slice(position))
          ? "0"
          : formula.slice(tokenStart, position);
    } else result += character;
  }
  return result;
}
function assertOutputNameUnreferenced(workbook, outputName) {
  let characterCount = 0;
  const formulaReferencesOutput = (formulaText) => (
      (characterCount += formulaText.length) >
        dataModelBudgets.nameFormulaCharacters &&
        failWithMessage(
          "Output-name safety checking exceeds 1,000,000 formula characters. Use a smaller reviewed workbook.",
        ),
      [
        ...stripFormulaLiterals(formulaText).matchAll(
          /('(?:[^']|'')*'|[\p{L}\p{M}\p{N}_\\][\p{L}\p{M}\p{N}_.\\]*)\s*!/gu,
        ),
      ].some((match) => {
        const referencedSheetName = match[1];
        return (
          (referencedSheetName.startsWith("'")
            ? referencedSheetName.slice(1, -1).replace(/''/g, "'")
            : referencedSheetName
          ).toLowerCase() === outputName.toLowerCase()
        );
      })
    ),
    formulaTexts = [];
  for (const sheet of workbook.sheets) {
    for (const cell of Object.values(sheet.cells))
      typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal &&
        formulaTexts.push(cell.raw);
    for (const conditionalFormat of sheet.conditionalFormats ?? [])
      conditionalFormat.kind === "formula" &&
        conditionalFormat.formula &&
        formulaTexts.push(conditionalFormat.formula);
    for (const table of sheet.tables ?? [])
      for (const column of table.columns)
        column.calculatedFormula && formulaTexts.push(column.calculatedFormula);
  }
  for (const definedName of workbook.definedNames ?? [])
    formulaTexts.push(definedName.refersTo);
  formulaTexts.some(formulaReferencesOutput) &&
    failWithMessage(
      "The output worksheet name is already referenced by workbook formulas or formula metadata. Choose another name to avoid reconnecting an unresolved source reference.",
    );
}
function parseImportedCellValue(raw, format) {
  if (format === "general") {
    const isNumericText =
      /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw);
    if (/^(TRUE|FALSE)$/i.test(raw))
      return { raw: raw.toUpperCase() === "TRUE", precision: false };
    if (isNumericText) {
      const digitCount = raw
          .split(/[eE]/)[0]
          .replace(/[^0-9]/g, "")
          .replace(/^0+/, "").length,
        numericValue = Number(raw);
      return digitCount > 15 ||
        !Number.isFinite(numericValue) ||
        (numericValue === 0 && /[1-9]/.test(raw.split(/[eE]/)[0])) ||
        (Number.isInteger(numericValue) && !Number.isSafeInteger(numericValue))
        ? { raw: raw, precision: true }
        : { raw: numericValue, precision: false };
    }
  }
  return { raw: raw, precision: false };
}

export {
  validateTextToColumnsRequest,
  assertSelectionInAuthorizedSheet,
  splitDelimitedCell,
  parseImportedCellValue,
  DataModelSnapshotNotice,
  rangesOverlap,
  dataModelBudgets,
  formatCellTextValue,
  DataModelDialog,
  validateTextToColumnsPermissions,
  useDialogFocusTrap,
  assertOutputNameUnreferenced,
};
