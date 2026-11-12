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

function DataModelDialog(e) {
  const { text } = useWorkbookMessages(),
    t = ne.useRef(e);
  t.current = e;
  const n = ne.useId(),
    o = e.workbook.sheets.find(
      (sheet) => sheet.id === e.sheetId,
    )?.dataModelReport,
    [definition, s] = ne.useState(() =>
      structuredClone(
        o?.definition ??
          e.workbook.dataModel ?? {
            engine: columnarModelFormatVersion,
            name: "Workbook model",
            tables: [],
            relationships: [],
            measures: [],
          },
      ),
    ),
    [query, u] = ne.useState(() =>
      structuredClone(
        o?.query ?? {
          groupBy: [],
          filters: [],
          measureIds: definition.measures.map((measure) => measure.id),
        },
      ),
    ),
    [l, c] = ne.useState("Tables"),
    [sheetId, f] = ne.useState(e.sheetId),
    [p, m] = ne.useState(
      e.range.includes(":")
        ? e.range
        : formatOccupiedRangeAddress(
            e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
          ),
    ),
    [v, g] = ne.useState(""),
    [checked, y] = ne.useState(false),
    [w, E] = ne.useState(""),
    [A, k] = ne.useState(null),
    [F, I] = ne.useState(0),
    [outputName, j] = ne.useState(() => {
      let le = 1;
      for (
        ;
        e.workbook.sheets.some(
          (sheet) => sheet.name.toLowerCase() === `model result ${le}`,
        );
      )
        le++;
      return `Model result ${le}`;
    }),
    z = ne.useRef(null),
    V = ne.useRef(false);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      z.current?.querySelector("button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const D = {
      readOnly: e.readOnly,
      allowedSheetIds: e.allowedSheetIds,
      includeHidden: !!e.allowHidden && checked,
    },
    B = () => {
      (k(null), E(""), (V.current = false));
    },
    onChange = (le) => {
      (s(le), B());
    },
    C = (le) => {
      (u(le), B());
    },
    T = (le) => definition.tables.find((table) => table.id === le),
    N = (le) => T(le)?.columns[0]?.name ?? "",
    R = (le, ce, ge) =>
      i.jsx("input", {
        "aria-label": text(le),
        value: ce,
        onChange: (event) => ge(event.target.value),
      }),
    L = (le, ce, ge, se) =>
      i.jsx("select", {
        "aria-label": text(le),
        value: ce,
        onChange: (event) => se(event.target.value),
        children: ge.map((we) =>
          i.jsx(
            "option",
            { value: we.value, children: text(we.label) },
            we.value,
          ),
        ),
      }),
    $ = (le) => le.map((ce) => ({ value: ce, label: ce })),
    U = (le, ce, ge) =>
      L(
        le,
        ce,
        definition.tables.map((table) => ({
          value: table.id,
          label: table.name,
        })),
        ge,
      ),
    G = (le, ce, ge, se) =>
      L(le, ge, $(T(ce)?.columns.map((column) => column.name) ?? []), se),
    H = (le, ce) =>
      onChange({
        ...definition,
        tables: definition.tables.map((table) =>
          table.id === le ? { ...table, ...ce } : table,
        ),
      }),
    Q = (le) => {
      const measures = definition.measures.filter(
        (measure) => measure.tableId !== le,
      );
      (onChange({
        ...definition,
        tables: definition.tables.filter((table) => table.id !== le),
        relationships: definition.relationships.filter(
          (relationship) =>
            relationship.one.tableId !== le && relationship.many.tableId !== le,
        ),
        measures: measures,
      }),
        C({
          ...query,
          groupBy: query.groupBy.filter((ge) => ge.tableId !== le),
          filters: query.filters.filter((filter) => filter.tableId !== le),
          measureIds: query.measureIds.filter((measureId) =>
            measures.some((measure) => measure.id === measureId),
          ),
        }));
    };
  function X$1() {
    try {
      const modelTableFromRange = buildModelTableFromRange(
        e.workbook,
        sheetId,
        p,
        v || undefined,
        D,
      );
      let ce = modelTableFromRange.name,
        ge = 2;
      for (
        ;
        definition.tables.some(
          (table) => table.name.toLowerCase() === ce.toLowerCase(),
        );
      )
        ce = `${modelTableFromRange.name} ${ge++}`;
      onChange({
        ...definition,
        tables: [...definition.tables, { ...modelTableFromRange, name: ce }],
      });
    } catch (error) {
      E(error.message);
    }
  }
  function onClick() {
    try {
      (E(""),
        I(0),
        (V.current = false),
        k(
          previewDataModelOutput(
            e.workbook,
            { definition: definition, query: query, outputName: outputName },
            D,
          ),
        ),
        c("Preview"));
    } catch (error) {
      (k(null), E(error.message));
    }
  }
  function q(le) {
    if (!(!A || V.current))
      try {
        const ce = t.current,
          ge = publishDataModelReport(ce.workbook, A, le, {
            readOnly: ce.readOnly,
            allowedSheetIds: ce.allowedSheetIds,
            includeHidden: !!ce.allowHidden && checked,
          });
        ((V.current = true),
          ce.onWorkbookChange(
            ge,
            le === "save"
              ? "Save relational data model"
              : "Publish relational model result",
          ),
          ce.onClose());
      } catch (error) {
        E(error.message);
      }
  }
  const W = A?.result,
    K =
      !!W?.complete && !e.readOnly && !e.workbook.structureProtection?.enabled,
    te = e.workbook.sheets.filter(
      (sheet) =>
        (!e.allowedSheetIds || e.allowedSheetIds.includes(sheet.id)) &&
        ((!sheet.hidden && !sheet.veryHidden) || (e.allowHidden && checked)),
    );
  return i.jsx("div", {
    className: "lg-modal-backdrop",
    children: i.jsxs("section", {
      className: "lg-data-model",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": n,
      ref: z,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.stopPropagation(), e.onClose()),
          event.key === "Tab")
        ) {
          const ce = [
              ...z.current.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),select:not(:disabled),summary",
              ),
            ].filter(
              (we) =>
                we.getClientRects().length ||
                navigator.userAgent.includes("jsdom"),
            ),
            ge = ce[0],
            se = ce.at(-1);
          event.shiftKey && document.activeElement === ge
            ? (event.preventDefault(), se == null || se.focus())
            : !event.shiftKey &&
              document.activeElement === se &&
              (event.preventDefault(), ge == null || ge.focus());
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
                  id: n,
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
              onClick: e.onClose,
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
          ].map((le) =>
            i.jsx(
              "button",
              {
                "aria-pressed": l === le,
                onClick: () => c(le),
                children: i.jsx(LocalizedText, { children: le }),
              },
              le,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "lg-model-body",
          children: [
            w &&
              i.jsx("p", {
                role: "alert",
                className: "lg-model-error",
                children: text(w),
              }),
            e.readOnly &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "This workbook is read-only. Preview remains available.",
                }),
              }),
            l === "Tables" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("div", {
                    className: "lg-model-controls",
                    children: [
                      LabeledField(
                        "Model name",
                        R("Model name", definition.name, (le) =>
                          onChange({ ...definition, name: le }),
                        ),
                      ),
                      LabeledField(
                        "Source worksheet",
                        L(
                          "Model source worksheet",
                          sheetId,
                          te.map((le) => ({ value: le.id, label: le.name })),
                          (le) => {
                            (f(le),
                              g(""),
                              m(
                                formatOccupiedRangeAddress(
                                  e.workbook.sheets.find(
                                    (sheet) => sheet.id === le,
                                  ),
                                ),
                              ));
                          },
                        ),
                      ),
                      LabeledField(
                        "Source table or range",
                        L(
                          "Model structured source",
                          v,
                          [
                            { value: "", label: "Cell range" },
                            ...(
                              e.workbook.sheets.find(
                                (sheet) => sheet.id === sheetId,
                              )?.tables ?? []
                            ).map((le) => ({ value: le.id, label: le.name })),
                          ],
                          g,
                        ),
                      ),
                      !v &&
                        LabeledField(
                          "Headed range",
                          R("Model source range", p, m),
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
                  o &&
                    e.workbook.dataModel &&
                    i.jsx("button", {
                      onClick: () => {
                        (onChange(structuredClone(e.workbook.dataModel)),
                          C({
                            groupBy: [],
                            filters: [],
                            measureIds: e.workbook.dataModel.measures.map(
                              (measure) => measure.id,
                            ),
                          }));
                      },
                      children: i.jsx(LocalizedText, {
                        children: "Use workbook model",
                      }),
                    }),
                  e.workbook.dataModel &&
                    i.jsx("button", {
                      disabled:
                        e.readOnly || e.workbook.structureProtection?.enabled,
                      onClick: () => {
                        try {
                          const le = t.current;
                          (le.onWorkbookChange(
                            removeDataModelFromWorkbook(le.workbook, {
                              readOnly: le.readOnly,
                            }),
                            "Remove relational data model",
                          ),
                            le.onClose());
                        } catch (error) {
                          E(error.message);
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
                                R(
                                  `${table.name} table name`,
                                  table.name,
                                  (ge) => H(table.id, { name: ge }),
                                ),
                              ),
                              i.jsxs("span", {
                                children: [
                                  e.workbook.sheets.find(
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
                                  R(
                                    `${table.name} source range`,
                                    table.range,
                                    (range) => H(table.id, { range: range }),
                                  ),
                                ),
                              i.jsx("button", {
                                "aria-label": `Remove table ${table.name}`,
                                onClick: () => Q(table.id),
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
                                      ].map((ge) =>
                                        i.jsx(
                                          "th",
                                          {
                                            children: i.jsx(LocalizedText, {
                                              children: ge,
                                            }),
                                          },
                                          ge,
                                        ),
                                      ),
                                    }),
                                  }),
                                  i.jsx("tbody", {
                                    children: table.columns.map((column, se) =>
                                      i.jsxs(
                                        "tr",
                                        {
                                          children: [
                                            i.jsx("th", {
                                              scope: "row",
                                              children: column.name,
                                            }),
                                            i.jsx("td", {
                                              children: L(
                                                `${table.name}.${column.name} type`,
                                                column.type,
                                                $([
                                                  "text",
                                                  "number",
                                                  "boolean",
                                                  "date",
                                                ]),
                                                (we) =>
                                                  H(table.id, {
                                                    columns: table.columns.map(
                                                      (Ee, Ae) =>
                                                        Ae === se
                                                          ? { ...Ee, type: we }
                                                          : Ee,
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
                                                  H(table.id, {
                                                    columns: table.columns.map(
                                                      (Ee, Ae) =>
                                                        Ae === se
                                                          ? {
                                                              ...Ee,
                                                              nullable:
                                                                event.target
                                                                  .checked,
                                                            }
                                                          : Ee,
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
            l === "Relationships" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("p", {
                    className: "lg-model-help",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Filters flow from the unique one side to the related side. Missing, blank or duplicate keys block calculation. Cycles and competing paths are refused.",
                    }),
                  }),
                  definition.relationships.map((relationship, ce) => {
                    const ge = (se) =>
                      onChange({
                        ...definition,
                        relationships: definition.relationships.map((we) =>
                          we.id === relationship.id ? se : we,
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
                              ce + 1,
                            ],
                          }),
                          i.jsxs("div", {
                            className: "lg-model-controls",
                            children: [
                              LabeledField(
                                "One table",
                                U(
                                  `Relationship ${ce + 1} one table`,
                                  relationship.one.tableId,
                                  (tableId) =>
                                    ge({
                                      ...relationship,
                                      one: {
                                        tableId: tableId,
                                        column: N(tableId),
                                      },
                                    }),
                                ),
                              ),
                              LabeledField(
                                "Unique key",
                                G(
                                  `Relationship ${ce + 1} one key`,
                                  relationship.one.tableId,
                                  relationship.one.column,
                                  (column) =>
                                    ge({
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
                                U(
                                  `Relationship ${ce + 1} related table`,
                                  relationship.many.tableId,
                                  (tableId) =>
                                    ge({
                                      ...relationship,
                                      many: {
                                        tableId: tableId,
                                        column: N(tableId),
                                      },
                                    }),
                                ),
                              ),
                              LabeledField(
                                "Foreign key",
                                G(
                                  `Relationship ${ce + 1} foreign key`,
                                  relationship.many.tableId,
                                  relationship.many.column,
                                  (column) =>
                                    ge({
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
                                L(
                                  `Relationship ${ce + 1} cardinality`,
                                  relationship.cardinality,
                                  [
                                    { value: "one-to-many", label: "1 → many" },
                                    { value: "one-to-one", label: "1 → 1" },
                                  ],
                                  (cardinality) =>
                                    ge({
                                      ...relationship,
                                      cardinality: cardinality,
                                    }),
                                ),
                              ),
                              i.jsx("button", {
                                "aria-label": `Remove relationship ${ce + 1}`,
                                onClick: () =>
                                  onChange({
                                    ...definition,
                                    relationships:
                                      definition.relationships.filter(
                                        (se) => se.id !== relationship.id,
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
                  }),
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
            l === "Calculated columns" &&
              i.jsx(DataModelDefinitionEditor, {
                definition: definition,
                onChange: onChange,
              }),
            l === "Measures" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("p", {
                    className: "lg-model-help",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Each measure aggregates its own filtered table, preventing fact-row multiplication. Choose DAX for expressions with row and filter context. Totals are recalculated in their own context.",
                    }),
                  }),
                  definition.measures.map((measure, ce) => {
                    const ge = (se) =>
                      onChange({
                        ...definition,
                        measures: definition.measures.map((we) =>
                          we.id === measure.id ? se : we,
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
                              ce + 1,
                            ],
                          }),
                          i.jsxs("div", {
                            className: "lg-model-controls",
                            children: [
                              LabeledField(
                                "Name",
                                R(
                                  `Measure ${ce + 1} name`,
                                  measure.name,
                                  (se) => ge({ ...measure, name: se }),
                                ),
                              ),
                              LabeledField(
                                "Table",
                                U(
                                  `Measure ${ce + 1} table`,
                                  measure.tableId,
                                  (tableId) =>
                                    ge({
                                      ...measure,
                                      tableId: tableId,
                                      ...(measure.operation === "countRows" ||
                                      measure.operation === "dax"
                                        ? {}
                                        : { column: N(tableId) }),
                                    }),
                                ),
                              ),
                              LabeledField(
                                "Operation",
                                L(
                                  `Measure ${ce + 1} operation`,
                                  measure.operation,
                                  $([
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
                                    const we = {
                                      ...measure,
                                      operation: operation,
                                    };
                                    (operation === "countRows" ||
                                    operation === "dax"
                                      ? delete we.column
                                      : (we.column =
                                          we.column ?? N(measure.tableId)),
                                      operation === "dax"
                                        ? (we.expression =
                                            we.expression ?? "BLANK()")
                                        : delete we.expression,
                                      ge(we));
                                  },
                                ),
                              ),
                              measure.operation !== "countRows" &&
                                measure.operation !== "dax" &&
                                LabeledField(
                                  "Column",
                                  G(
                                    `Measure ${ce + 1} column`,
                                    measure.tableId,
                                    measure.column ?? "",
                                    (column) =>
                                      ge({ ...measure, column: column }),
                                  ),
                                ),
                              i.jsx("button", {
                                "aria-label": `Remove measure ${ce + 1}`,
                                onClick: () => {
                                  (onChange({
                                    ...definition,
                                    measures: definition.measures.filter(
                                      (se) => se.id !== measure.id,
                                    ),
                                  }),
                                    C({
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
                              label: `Measure ${ce + 1} DAX expression`,
                              value: measure.expression ?? "",
                              onChange: (expression) =>
                                ge({ ...measure, expression: expression }),
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
                      const le = {
                        id: createUuid(),
                        name: `Measure ${definition.measures.length + 1}`,
                        tableId: definition.tables.at(-1).id,
                        operation: "countRows",
                      };
                      (onChange({
                        ...definition,
                        measures: [...definition.measures, le],
                      }),
                        C({
                          ...query,
                          measureIds: [...query.measureIds, le.id],
                        }));
                    },
                    children: i.jsx(LocalizedText, { children: "Add measure" }),
                  }),
                ],
              }),
            l === "Query" &&
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
                                    C({
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
                      query.groupBy.map((le, ce) =>
                        i.jsxs(
                          "div",
                          {
                            className: "lg-model-controls",
                            children: [
                              U(
                                `Group ${ce + 1} table`,
                                le.tableId,
                                (tableId) =>
                                  C({
                                    ...query,
                                    groupBy: query.groupBy.map((se, we) =>
                                      we === ce
                                        ? {
                                            tableId: tableId,
                                            column: N(tableId),
                                          }
                                        : se,
                                    ),
                                  }),
                              ),
                              G(
                                `Group ${ce + 1} column`,
                                le.tableId,
                                le.column,
                                (column) =>
                                  C({
                                    ...query,
                                    groupBy: query.groupBy.map((se, we) =>
                                      we === ce
                                        ? { ...se, column: column }
                                        : se,
                                    ),
                                  }),
                              ),
                              i.jsx("button", {
                                "aria-label": `Remove group ${ce + 1}`,
                                onClick: () =>
                                  C({
                                    ...query,
                                    groupBy: query.groupBy.filter(
                                      (ge, se) => se !== ce,
                                    ),
                                  }),
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            ],
                          },
                          ce,
                        ),
                      ),
                      i.jsx("button", {
                        disabled:
                          !definition.tables.length ||
                          query.groupBy.length >= 3,
                        onClick: () =>
                          C({
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
                      query.filters.map((filter, ce) => {
                        const ge = (Ae) =>
                            C({
                              ...query,
                              filters: query.filters.map((ye, Te) =>
                                Te === ce ? Ae : ye,
                              ),
                            }),
                          se = T(filter.tableId)?.columns.find(
                            (column) => column.name === filter.column,
                          )?.type;
                        return i.jsxs(
                          "div",
                          {
                            className: "lg-model-controls",
                            children: [
                              U(
                                `Filter ${ce + 1} table`,
                                filter.tableId,
                                (tableId) =>
                                  ge({
                                    ...filter,
                                    tableId: tableId,
                                    column: N(tableId),
                                    operator: "blank",
                                    value: undefined,
                                  }),
                              ),
                              G(
                                `Filter ${ce + 1} column`,
                                filter.tableId,
                                filter.column,
                                (column) =>
                                  ge({
                                    ...filter,
                                    column: column,
                                    operator: "blank",
                                    value: undefined,
                                  }),
                              ),
                              L(
                                `Filter ${ce + 1} operator`,
                                filter.operator,
                                $([
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
                                  ge({
                                    ...filter,
                                    operator: operator,
                                    value: ["blank", "notBlank"].includes(
                                      operator,
                                    )
                                      ? undefined
                                      : se === "number"
                                        ? 0
                                        : se === "boolean"
                                          ? false
                                          : "",
                                  }),
                              ),
                              !["blank", "notBlank"].includes(
                                filter.operator,
                              ) &&
                                (se === "boolean"
                                  ? L(
                                      `Filter ${ce + 1} value`,
                                      String(filter.value),
                                      $(["true", "false"]),
                                      (Ae) =>
                                        ge({ ...filter, value: Ae === "true" }),
                                    )
                                  : R(
                                      `Filter ${ce + 1} value`,
                                      String(filter.value ?? ""),
                                      (Ae) =>
                                        ge({
                                          ...filter,
                                          value:
                                            se === "number"
                                              ? Ae.trim()
                                                ? Number(Ae)
                                                : null
                                              : Ae,
                                        }),
                                    )),
                              i.jsx("button", {
                                "aria-label": `Remove filter ${ce + 1}`,
                                onClick: () =>
                                  C({
                                    ...query,
                                    filters: query.filters.filter(
                                      (Ae, ye) => ye !== ce,
                                    ),
                                  }),
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            ],
                          },
                          ce,
                        );
                      }),
                      i.jsx("button", {
                        disabled:
                          !definition.tables.length ||
                          query.filters.length >= dataModelLimits.filters,
                        onClick: () =>
                          C({
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
            l === "Preview" &&
              (W
                ? i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("div", {
                        role: "status",
                        "aria-label": text("Model preview status"),
                        className: "lg-model-summary",
                        children: [
                          W.complete
                            ? text("Model checks passed")
                            : text("Model blocked"),
                          " ",
                          "· ",
                          W.tables.length,
                          " ",
                          text("tables"),
                          " ·",
                          " ",
                          W.operations.toLocaleString(),
                          " ",
                          text("operations"),
                        ],
                      }),
                      W.diagnostics.length > 0 &&
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
                              children: W.diagnostics.map((diagnostic, ce) =>
                                i.jsxs(
                                  "tr",
                                  {
                                    children: [
                                      i.jsx("td", {
                                        children: diagnostic.code,
                                      }),
                                      i.jsxs("td", {
                                        children: [
                                          T(diagnostic.tableId ?? "")?.name,
                                          " ",
                                          diagnostic.row ?? "",
                                        ],
                                      }),
                                      i.jsx("td", {
                                        children: text(diagnostic.message),
                                      }),
                                    ],
                                  },
                                  ce,
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
                              ].map((le) =>
                                i.jsx(
                                  "th",
                                  {
                                    children: i.jsx(LocalizedText, {
                                      children: le,
                                    }),
                                  },
                                  le,
                                ),
                              ),
                            }),
                          }),
                          i.jsx("tbody", {
                            children: W.tables.map((table) =>
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
                                        e.workbook.sheets.find(
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
                      W.complete &&
                        !!W.calculatedColumns?.length &&
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
                                  children: W.calculatedColumns.map(
                                    (calculatedColumn) =>
                                      i.jsxs(
                                        "tr",
                                        {
                                          children: [
                                            i.jsxs("th", {
                                              scope: "row",
                                              children: [
                                                T(calculatedColumn.tableId)
                                                  ?.name,
                                                ".",
                                                calculatedColumn.name,
                                              ],
                                            }),
                                            i.jsx("td", {
                                              children: calculatedColumn.values
                                                .slice(0, 25)
                                                .map((ge) =>
                                                  ge === null
                                                    ? text("Blank")
                                                    : String(ge),
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
                      W.complete &&
                        query.measureIds.length > 0 &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("table", {
                              "aria-label": text("Model measure results"),
                              children: [
                                i.jsx("thead", {
                                  children: i.jsx("tr", {
                                    children: W.headers.map((header) =>
                                      i.jsx("th", { children: header }, header),
                                    ),
                                  }),
                                }),
                                i.jsx("tbody", {
                                  children: W.rows
                                    .slice(F * 25, (F + 1) * 25)
                                    .map((le, ce) =>
                                      i.jsx(
                                        "tr",
                                        {
                                          children: [
                                            ...le.keys,
                                            ...le.values,
                                          ].map((ge, se) =>
                                            i.jsx(
                                              "td",
                                              {
                                                children:
                                                  ge === null
                                                    ? text("Blank")
                                                    : String(ge),
                                              },
                                              se,
                                            ),
                                          ),
                                        },
                                        ce,
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
                                        W.totals.map((total, ce) =>
                                          i.jsx(
                                            "td",
                                            {
                                              children:
                                                total === null
                                                  ? text("Blank")
                                                  : total,
                                            },
                                            ce,
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
                                  disabled: !F,
                                  onClick: () => I((le) => le - 1),
                                  children: i.jsx(LocalizedText, {
                                    children: "Previous results",
                                  }),
                                }),
                                i.jsxs("span", {
                                  children: [
                                    F + 1,
                                    " /",
                                    " ",
                                    Math.max(1, Math.ceil(W.rows.length / 25)),
                                  ],
                                }),
                                i.jsx("button", {
                                  disabled: (F + 1) * 25 >= W.rows.length,
                                  onClick: () => I((le) => le + 1),
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
            e.allowHidden &&
              i.jsxs("label", {
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => {
                      (y(event.target.checked), B());
                    },
                  }),
                  i.jsx(LocalizedText, { children: "Include hidden sources" }),
                ],
              }),
            LabeledField(
              "Output worksheet",
              R("Model output worksheet", outputName, (le) => {
                (j(le), B());
              }),
            ),
            i.jsx("button", {
              onClick: onClick,
              children: i.jsx(LocalizedText, { children: "Preview model" }),
            }),
            i.jsx("button", {
              disabled: !K,
              onClick: () => q("save"),
              children: i.jsx(LocalizedText, { children: "Save model" }),
            }),
            i.jsx("button", {
              className: "lg-model-primary",
              disabled: !K || !query.measureIds.length,
              onClick: () => q("publish"),
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
  const o = ne.useMemo(
    () => getModelReportStatus(workbook, sheetId, options),
    [workbook, sheetId, options?.includeHidden, options?.allowedSheetIds],
  );
  return o.status === "none"
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
                  children: `Model snapshot · ${o.status}`,
                }),
              }),
              " ",
              "· ",
              i.jsx(LocalizedText, { children: o.message }),
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
function isElementVisibleForFocus(e) {
  for (let r = e; r; r = r.parentElement) {
    const computedStyle = getComputedStyle(r);
    if (
      r.hidden ||
      computedStyle.display === "none" ||
      computedStyle.visibility === "hidden" ||
      computedStyle.visibility === "collapse"
    )
      return false;
    if (r.tagName === "DETAILS" && !r.hasAttribute("open")) {
      const n = Array.from(r.children).find((o) => o.tagName === "SUMMARY");
      if (!n?.contains(e)) return false;
    }
  }
  return true;
}
function getFocusableElements(e) {
  return Array.from(e.querySelectorAll(focusableElementSelector)).filter(
    (r) =>
      !(
        r.tabIndex < 0 ||
        r.matches(":disabled") ||
        !isElementVisibleForFocus(r) ||
        r.closest('[hidden],[inert],[aria-hidden="true"]')
      ),
  );
}
function useDialogFocusTrap(appRootRef) {
  const r = ne.useRef(undefined);
  (ne.useLayoutEffect(() => {
    const t = appRootRef.current;
    if (!t) return;
    let n,
      o = null;
    const a = new Map();
    let s =
      document.activeElement instanceof HTMLElement
        ? document.activeElement
        : null;
    const d = new Map(),
      u = () => {
        for (const [p, m] of d)
          (m.inert === null
            ? p.removeAttribute("inert")
            : p.setAttribute("inert", m.inert),
            m.hidden === null
              ? p.removeAttribute("aria-hidden")
              : p.setAttribute("aria-hidden", m.hidden));
        d.clear();
      },
      l = () => {
        const p = Array.from(
          t.querySelectorAll('[role="dialog"][aria-modal="true"]'),
        )
          .filter(isElementVisibleForFocus)
          .at(-1);
        if (p === n) {
          p &&
            (!isElementVisibleForFocus(document.activeElement) ||
              document.activeElement?.matches(":disabled")) &&
            (getFocusableElements(p)[0] ?? p).focus({ preventScroll: true });
          return;
        }
        const m = n;
        if ((u(), (n = p), p)) {
          m || (o = s);
          const g = a.has(p);
          g || a.set(p, s);
          let b = p;
          for (; b !== t && b.parentElement;) {
            for (const w of Array.from(b.parentElement.children))
              w instanceof HTMLElement &&
                w !== b &&
                (d.set(w, {
                  inert: w.getAttribute("inert"),
                  hidden: w.getAttribute("aria-hidden"),
                }),
                w.setAttribute("inert", ""),
                w.setAttribute("aria-hidden", "true"));
            b = b.parentElement;
          }
          const y = m ? a.get(m) : undefined;
          if (!p.contains(document.activeElement)) {
            const w =
              g &&
              y &&
              p.contains(y) &&
              isElementVisibleForFocus(y) &&
              !y.matches(":disabled")
                ? y
                : (getFocusableElements(p)[0] ?? p);
            (w === p && (p.tabIndex = -1), w.focus({ preventScroll: true }));
          }
        } else if (
          m &&
          (document.activeElement === document.body ||
            m.contains(document.activeElement))
        ) {
          const g =
            o?.isConnected &&
            isElementVisibleForFocus(o) &&
            !o.matches(":disabled")
              ? o
              : t.querySelector(".lg-grid-scroll");
          g?.focus({ preventScroll: true });
        }
        (m && (!m.isConnected || !isElementVisibleForFocus(m)) && a.delete(m),
          p || a.clear());
      },
      c = (event) => {
        const target = event.target;
        if (target instanceof HTMLElement) {
          if (n && t.contains(target) && !n.contains(target)) {
            (getFocusableElements(n)[0] ?? n).focus({ preventScroll: true });
            return;
          }
          ((!n && !target.closest('[role="dialog"][aria-modal="true"]')) ||
            n?.contains(target)) &&
            (s = target);
        }
      },
      h = (event) => {
        if ((l(), !n || event.key !== "Tab")) return;
        const focusableElements = getFocusableElements(n),
          v = focusableElements[0],
          g = focusableElements.at(-1);
        if (!v) {
          (event.preventDefault(), (n.tabIndex = -1), n.focus());
          return;
        }
        (!n.contains(document.activeElement) ||
          (event.shiftKey && document.activeElement === v) ||
          (!event.shiftKey && document.activeElement === g)) &&
          (event.preventDefault(),
          event.stopPropagation(),
          (event.shiftKey ? g : v)?.focus());
      };
    (t.addEventListener("keydown", h, true),
      document.addEventListener("focusin", c, true));
    const f = new MutationObserver(l);
    return (
      f.observe(t, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ["hidden", "style", "class", "open", "disabled"],
      }),
      l(),
      (r.current = l),
      () => {
        (f.disconnect(),
          t.removeEventListener("keydown", h, true),
          document.removeEventListener("focusin", c, true),
          u(),
          (r.current = undefined));
      }
    );
  }, [appRootRef]),
    ne.useLayoutEffect(() => r.current?.()));
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

function assertRecordKeys(destination, r, t) {
  (!isNonArrayObject(destination) ||
    Object.keys(destination).some((n) => !r.includes(n))) &&
    failWithMessage(t);
}
const rangesOverlap = (e, r) => {
  const rangeAddress = parseRangeAddress(e),
    n = parseRangeAddress(r);
  return (
    rangeAddress.firstRow <= n.lastRow &&
    rangeAddress.lastRow >= n.firstRow &&
    rangeAddress.firstCol <= n.lastCol &&
    rangeAddress.lastCol >= n.firstCol
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
  const t = [];
  let n = "",
    o = false,
    a = false,
    s = false;
  const d = [...cellTextValue],
    u = () => {
      (t.push(n),
        t.length > dataModelBudgets.fields &&
          failWithMessage(
            "A source cell contains more than 50 fields. Choose a different delimiter or split a smaller record.",
          ),
        (n = ""),
        (s = false),
        (a = false));
    };
  for (let l = 0; l < d.length; l++) {
    const c = d[l];
    o
      ? c !== '"'
        ? (n += c)
        : d[l + 1] === '"'
          ? ((n += '"'), l++)
          : ((o = false), (a = true))
      : c === delimiter
        ? u()
        : a
          ? failWithMessage(
              "Unexpected text after a closing quote. A closing quote must be followed by the delimiter or end of the cell.",
            )
          : c === '"'
            ? (s &&
                failWithMessage(
                  "A quote must begin a field; escape embedded quotes by doubling them inside a quoted field.",
                ),
              (o = true),
              (s = true))
            : ((n += c), (s = true));
  }
  return (
    o &&
      failWithMessage(
        "A quoted field is not closed. Check the source cell before splitting.",
      ),
    u(),
    t
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
    const t = request.destination.name.trim();
    ((!t ||
      t.length > 31 ||
      /[\u0000-\u001f\u007f\\/\[\]*?:]/.test(t) ||
      t.startsWith("'") ||
      t.endsWith("'")) &&
      failWithMessage(
        "Output worksheet names need 1–31 characters without control/Excel-reserved characters or leading/trailing apostrophes.",
      ),
      (destination = { kind: "new-sheet", name: t }));
  } else if (request.destination.kind === "range") {
    (assertRecordKeys(
      request.destination,
      ["kind", "sheetId", "address"],
      "Invalid output range destination.",
    ),
      (typeof request.destination.sheetId != "string" ||
        typeof request.destination.address != "string") &&
        failWithMessage("Choose an output worksheet and one starting cell."));
    const t = _t(request.destination.address.trim());
    destination = {
      kind: "range",
      sheetId: request.destination.sheetId,
      address: formatCellAddress(t.row, t.col),
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
function validateTextToColumnsPermissions(workbook, r) {
  return (
    [r.readOnly, r.includeHidden, r.allowNewSheet].some(
      (t) => t !== undefined && typeof t != "boolean",
    ) && failWithMessage("Text to Columns permissions must be boolean."),
    r.readOnly && failWithMessage("This workbook is read-only."),
    r.allowedSheetIds !== undefined &&
      (!Array.isArray(r.allowedSheetIds) ||
        r.allowedSheetIds.some(
          (allowedSheetId) =>
            typeof allowedSheetId != "string" ||
            !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
        )) &&
      failWithMessage("The authorized worksheet context is invalid."),
    {
      includeHidden: r.includeHidden ?? false,
      allowNewSheet: r.allowNewSheet ?? true,
      allowedSheetIds: [
        ...new Set(
          r.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
        ),
      ].sort(),
    }
  );
}
function assertSelectionInAuthorizedSheet(e, r, context) {
  const rangeAddress = parseRangeAddress(r);
  ((!context.allowedSheetIds.includes(e.id) ||
    (!context.includeHidden &&
      (e.hidden ||
        e.veryHidden ||
        cn(e, "row").some(
          (o) =>
            o.start <= rangeAddress.lastRow && o.end >= rangeAddress.firstRow,
        ) ||
        cn(e, "column").some(
          (o) =>
            o.start <= rangeAddress.lastCol && o.end >= rangeAddress.firstCol,
        )))) &&
    failWithMessage(
      "The selected cells are outside the visible, authorized worksheet context.",
    ),
    (rangeAddress.lastRow >= e.rowCount ||
      rangeAddress.lastCol >= e.colCount) &&
      failWithMessage("The selected cells exceed the worksheet boundary."));
}
function formatCellTextValue(e) {
  return e === null
    ? ""
    : typeof e == "boolean"
      ? e
        ? "TRUE"
        : "FALSE"
      : String(e);
}
function stripFormulaLiterals(e) {
  let r = "",
    t = 0;
  for (; t < e.length;) {
    const n = t,
      o = e[t++];
    if (o === '"' || o === "'") {
      for (; t < e.length;)
        if (e[t++] === o)
          if (e[t] === o) t++;
          else break;
      r += o === '"' ? '""' : e.slice(n, t);
    } else if (o === "[") {
      let a = 1;
      for (; t < e.length && a;)
        (e[t] === "[" && a++, e[t] === "]" && a--, t++);
      r +=
        !a && !/^[\p{L}\p{M}\p{N}_.!\\]/u.test(e.slice(t))
          ? "0"
          : e.slice(n, t);
    } else r += o;
  }
  return r;
}
function assertOutputNameUnreferenced(workbook, outputName) {
  let t = 0;
  const n = (a) => (
      (t += a.length) > dataModelBudgets.nameFormulaCharacters &&
        failWithMessage(
          "Output-name safety checking exceeds 1,000,000 formula characters. Use a smaller reviewed workbook.",
        ),
      [
        ...stripFormulaLiterals(a).matchAll(
          /('(?:[^']|'')*'|[\p{L}\p{M}\p{N}_\\][\p{L}\p{M}\p{N}_.\\]*)\s*!/gu,
        ),
      ].some((d) => {
        const u = d[1];
        return (
          (u.startsWith("'")
            ? u.slice(1, -1).replace(/''/g, "'")
            : u
          ).toLowerCase() === outputName.toLowerCase()
        );
      })
    ),
    o = [];
  for (const sheet of workbook.sheets) {
    for (const cell of Object.values(sheet.cells))
      typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal &&
        o.push(cell.raw);
    for (const s of sheet.conditionalFormats ?? [])
      s.kind === "formula" && s.formula && o.push(s.formula);
    for (const s of sheet.tables ?? [])
      for (const column of s.columns)
        column.calculatedFormula && o.push(column.calculatedFormula);
  }
  for (const a of workbook.definedNames ?? []) o.push(a.refersTo);
  o.some(n) &&
    failWithMessage(
      "The output worksheet name is already referenced by workbook formulas or formula metadata. Choose another name to avoid reconnecting an unresolved source reference.",
    );
}
function parseImportedCellValue(raw, format) {
  if (format === "general") {
    const t = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw);
    if (/^(TRUE|FALSE)$/i.test(raw))
      return { raw: raw.toUpperCase() === "TRUE", precision: false };
    if (t) {
      const n = raw
          .split(/[eE]/)[0]
          .replace(/[^0-9]/g, "")
          .replace(/^0+/, "").length,
        o = Number(raw);
      return n > 15 ||
        !Number.isFinite(o) ||
        (o === 0 && /[1-9]/.test(raw.split(/[eE]/)[0])) ||
        (Number.isInteger(o) && !Number.isSafeInteger(o))
        ? { raw: raw, precision: true }
        : { raw: o, precision: false };
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
