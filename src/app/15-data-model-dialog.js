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
    o = e.workbook.sheets.find((le) => le.id === e.sheetId)?.dataModelReport,
    [a, s] = ne.useState(() =>
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
    [d, u] = ne.useState(() =>
      structuredClone(
        o?.query ?? {
          groupBy: [],
          filters: [],
          measureIds: a.measures.map((le) => le.id),
        },
      ),
    ),
    [l, c] = ne.useState("Tables"),
    [h, f] = ne.useState(e.sheetId),
    [p, m] = ne.useState(
      e.range.includes(":")
        ? e.range
        : formatOccupiedRangeAddress(
            e.workbook.sheets.find((le) => le.id === e.sheetId),
          ),
    ),
    [v, g] = ne.useState(""),
    [b, y] = ne.useState(false),
    [w, E] = ne.useState(""),
    [A, k] = ne.useState(null),
    [F, I] = ne.useState(0),
    [O, j] = ne.useState(() => {
      let le = 1;
      for (
        ;
        e.workbook.sheets.some(
          (ce) => ce.name.toLowerCase() === `model result ${le}`,
        );
      )
        le++;
      return `Model result ${le}`;
    }),
    z = ne.useRef(null),
    V = ne.useRef(false);
  ne.useEffect(() => {
    const le = document.activeElement;
    return (
      z.current?.querySelector("button")?.focus(),
      () => {
        le?.isConnected && le.focus();
      }
    );
  }, []);
  const D = {
      readOnly: e.readOnly,
      allowedSheetIds: e.allowedSheetIds,
      includeHidden: !!e.allowHidden && b,
    },
    B = () => {
      (k(null), E(""), (V.current = false));
    },
    _ = (le) => {
      (s(le), B());
    },
    C = (le) => {
      (u(le), B());
    },
    T = (le) => a.tables.find((ce) => ce.id === le),
    N = (le) => T(le)?.columns[0]?.name ?? "",
    R = (le, ce, ge) =>
      i.jsx("input", {
        "aria-label": text(le),
        value: ce,
        onChange: (se) => ge(se.target.value),
      }),
    L = (le, ce, ge, se) =>
      i.jsx("select", {
        "aria-label": text(le),
        value: ce,
        onChange: (we) => se(we.target.value),
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
        a.tables.map((se) => ({ value: se.id, label: se.name })),
        ge,
      ),
    G = (le, ce, ge, se) =>
      L(le, ge, $(T(ce)?.columns.map((Ee) => Ee.name) ?? []), se),
    H = (le, ce) =>
      _({
        ...a,
        tables: a.tables.map((ge) => (ge.id === le ? { ...ge, ...ce } : ge)),
      }),
    Q = (le) => {
      const ce = a.measures.filter((ge) => ge.tableId !== le);
      (_({
        ...a,
        tables: a.tables.filter((ge) => ge.id !== le),
        relationships: a.relationships.filter(
          (ge) => ge.one.tableId !== le && ge.many.tableId !== le,
        ),
        measures: ce,
      }),
        C({
          ...d,
          groupBy: d.groupBy.filter((ge) => ge.tableId !== le),
          filters: d.filters.filter((ge) => ge.tableId !== le),
          measureIds: d.measureIds.filter((ge) =>
            ce.some((se) => se.id === ge),
          ),
        }));
    };
  function X$1() {
    try {
      const le = buildModelTableFromRange(e.workbook, h, p, v || undefined, D);
      let ce = le.name,
        ge = 2;
      for (; a.tables.some((se) => se.name.toLowerCase() === ce.toLowerCase());)
        ce = `${le.name} ${ge++}`;
      _({ ...a, tables: [...a.tables, { ...le, name: ce }] });
    } catch (le) {
      E(le.message);
    }
  }
  function J() {
    try {
      (E(""),
        I(0),
        (V.current = false),
        k(
          previewDataModelOutput(
            e.workbook,
            { definition: a, query: d, outputName: O },
            D,
          ),
        ),
        c("Preview"));
    } catch (le) {
      (k(null), E(le.message));
    }
  }
  function q(le) {
    if (!(!A || V.current))
      try {
        const ce = t.current,
          ge = publishDataModelReport(ce.workbook, A, le, {
            readOnly: ce.readOnly,
            allowedSheetIds: ce.allowedSheetIds,
            includeHidden: !!ce.allowHidden && b,
          });
        ((V.current = true),
          ce.onWorkbookChange(
            ge,
            le === "save"
              ? "Save relational data model"
              : "Publish relational model result",
          ),
          ce.onClose());
      } catch (ce) {
        E(ce.message);
      }
  }
  const W = A?.result,
    K =
      !!W?.complete && !e.readOnly && !e.workbook.structureProtection?.enabled,
    te = e.workbook.sheets.filter(
      (le) =>
        (!e.allowedSheetIds || e.allowedSheetIds.includes(le.id)) &&
        ((!le.hidden && !le.veryHidden) || (e.allowHidden && b)),
    );
  return i.jsx("div", {
    className: "lg-modal-backdrop",
    children: i.jsxs("section", {
      className: "lg-data-model",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": n,
      ref: z,
      onKeyDown: (le) => {
        if (
          (le.key === "Escape" && (le.stopPropagation(), e.onClose()),
          le.key === "Tab")
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
          le.shiftKey && document.activeElement === ge
            ? (le.preventDefault(), se == null || se.focus())
            : !le.shiftKey &&
              document.activeElement === se &&
              (le.preventDefault(), ge == null || ge.focus());
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
                        R("Model name", a.name, (le) => _({ ...a, name: le })),
                      ),
                      LabeledField(
                        "Source worksheet",
                        L(
                          "Model source worksheet",
                          h,
                          te.map((le) => ({ value: le.id, label: le.name })),
                          (le) => {
                            (f(le),
                              g(""),
                              m(
                                formatOccupiedRangeAddress(
                                  e.workbook.sheets.find((ce) => ce.id === le),
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
                              e.workbook.sheets.find((le) => le.id === h)
                                ?.tables ?? []
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
                        disabled: a.tables.length >= dataModelLimits.tables,
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
                        (_(structuredClone(e.workbook.dataModel)),
                          C({
                            groupBy: [],
                            filters: [],
                            measureIds: e.workbook.dataModel.measures.map(
                              (le) => le.id,
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
                        } catch (le) {
                          E(le.message);
                        }
                      },
                      children: i.jsx(LocalizedText, {
                        children: "Remove saved model",
                      }),
                    }),
                  a.tables.map((le) =>
                    i.jsxs(
                      "fieldset",
                      {
                        children: [
                          i.jsx("legend", { children: le.name }),
                          i.jsxs("div", {
                            className: "lg-model-controls",
                            children: [
                              LabeledField(
                                "Table name",
                                R(`${le.name} table name`, le.name, (ge) =>
                                  H(le.id, { name: ge }),
                                ),
                              ),
                              i.jsxs("span", {
                                children: [
                                  e.workbook.sheets.find(
                                    (ge) => ge.id === le.sheetId,
                                  )?.name ?? text("Missing worksheet"),
                                  " ",
                                  "· ",
                                  le.range,
                                  le.tableId
                                    ? " · Linked structured table"
                                    : "",
                                ],
                              }),
                              !le.tableId &&
                                LabeledField(
                                  "Headed range",
                                  R(`${le.name} source range`, le.range, (ge) =>
                                    H(le.id, { range: ge }),
                                  ),
                                ),
                              i.jsx("button", {
                                "aria-label": `Remove table ${le.name}`,
                                onClick: () => Q(le.id),
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
                                  le.columns.length,
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
                                    children: le.columns.map((ge, se) =>
                                      i.jsxs(
                                        "tr",
                                        {
                                          children: [
                                            i.jsx("th", {
                                              scope: "row",
                                              children: ge.name,
                                            }),
                                            i.jsx("td", {
                                              children: L(
                                                `${le.name}.${ge.name} type`,
                                                ge.type,
                                                $([
                                                  "text",
                                                  "number",
                                                  "boolean",
                                                  "date",
                                                ]),
                                                (we) =>
                                                  H(le.id, {
                                                    columns: le.columns.map(
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
                                                "aria-label": `${le.name}.${ge.name} nullable`,
                                                checked: ge.nullable,
                                                onChange: (we) =>
                                                  H(le.id, {
                                                    columns: le.columns.map(
                                                      (Ee, Ae) =>
                                                        Ae === se
                                                          ? {
                                                              ...Ee,
                                                              nullable:
                                                                we.target
                                                                  .checked,
                                                            }
                                                          : Ee,
                                                    ),
                                                  }),
                                              }),
                                            }),
                                          ],
                                        },
                                        ge.name,
                                      ),
                                    ),
                                  }),
                                ],
                              }),
                            ],
                          }),
                        ],
                      },
                      le.id,
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
                  a.relationships.map((le, ce) => {
                    const ge = (se) =>
                      _({
                        ...a,
                        relationships: a.relationships.map((we) =>
                          we.id === le.id ? se : we,
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
                                  le.one.tableId,
                                  (se) =>
                                    ge({
                                      ...le,
                                      one: { tableId: se, column: N(se) },
                                    }),
                                ),
                              ),
                              LabeledField(
                                "Unique key",
                                G(
                                  `Relationship ${ce + 1} one key`,
                                  le.one.tableId,
                                  le.one.column,
                                  (se) =>
                                    ge({
                                      ...le,
                                      one: { ...le.one, column: se },
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
                                  le.many.tableId,
                                  (se) =>
                                    ge({
                                      ...le,
                                      many: { tableId: se, column: N(se) },
                                    }),
                                ),
                              ),
                              LabeledField(
                                "Foreign key",
                                G(
                                  `Relationship ${ce + 1} foreign key`,
                                  le.many.tableId,
                                  le.many.column,
                                  (se) =>
                                    ge({
                                      ...le,
                                      many: { ...le.many, column: se },
                                    }),
                                ),
                              ),
                              LabeledField(
                                "Cardinality",
                                L(
                                  `Relationship ${ce + 1} cardinality`,
                                  le.cardinality,
                                  [
                                    { value: "one-to-many", label: "1 → many" },
                                    { value: "one-to-one", label: "1 → 1" },
                                  ],
                                  (se) => ge({ ...le, cardinality: se }),
                                ),
                              ),
                              i.jsx("button", {
                                "aria-label": `Remove relationship ${ce + 1}`,
                                onClick: () =>
                                  _({
                                    ...a,
                                    relationships: a.relationships.filter(
                                      (se) => se.id !== le.id,
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
                      le.id,
                    );
                  }),
                  i.jsx("button", {
                    disabled:
                      a.tables.length < 2 ||
                      a.relationships.length >= dataModelLimits.relationships,
                    onClick: () =>
                      _({
                        ...a,
                        relationships: [
                          ...a.relationships,
                          {
                            id: createUuid(),
                            one: {
                              tableId: a.tables[0].id,
                              column: a.tables[0].columns[0].name,
                            },
                            many: {
                              tableId: a.tables[1].id,
                              column: a.tables[1].columns[0].name,
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
              i.jsx(DataModelDefinitionEditor, { definition: a, onChange: _ }),
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
                  a.measures.map((le, ce) => {
                    const ge = (se) =>
                      _({
                        ...a,
                        measures: a.measures.map((we) =>
                          we.id === le.id ? se : we,
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
                                R(`Measure ${ce + 1} name`, le.name, (se) =>
                                  ge({ ...le, name: se }),
                                ),
                              ),
                              LabeledField(
                                "Table",
                                U(`Measure ${ce + 1} table`, le.tableId, (se) =>
                                  ge({
                                    ...le,
                                    tableId: se,
                                    ...(le.operation === "countRows" ||
                                    le.operation === "dax"
                                      ? {}
                                      : { column: N(se) }),
                                  }),
                                ),
                              ),
                              LabeledField(
                                "Operation",
                                L(
                                  `Measure ${ce + 1} operation`,
                                  le.operation,
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
                                  (se) => {
                                    const we = { ...le, operation: se };
                                    (se === "countRows" || se === "dax"
                                      ? delete we.column
                                      : (we.column =
                                          we.column ?? N(le.tableId)),
                                      se === "dax"
                                        ? (we.expression =
                                            we.expression ?? "BLANK()")
                                        : delete we.expression,
                                      ge(we));
                                  },
                                ),
                              ),
                              le.operation !== "countRows" &&
                                le.operation !== "dax" &&
                                LabeledField(
                                  "Column",
                                  G(
                                    `Measure ${ce + 1} column`,
                                    le.tableId,
                                    le.column ?? "",
                                    (se) => ge({ ...le, column: se }),
                                  ),
                                ),
                              i.jsx("button", {
                                "aria-label": `Remove measure ${ce + 1}`,
                                onClick: () => {
                                  (_({
                                    ...a,
                                    measures: a.measures.filter(
                                      (se) => se.id !== le.id,
                                    ),
                                  }),
                                    C({
                                      ...d,
                                      measureIds: d.measureIds.filter(
                                        (se) => se !== le.id,
                                      ),
                                    }));
                                },
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            ],
                          }),
                          le.operation === "dax" &&
                            i.jsx(DaxExpressionInput, {
                              label: `Measure ${ce + 1} DAX expression`,
                              value: le.expression ?? "",
                              onChange: (se) => ge({ ...le, expression: se }),
                            }),
                        ],
                      },
                      le.id,
                    );
                  }),
                  i.jsx("button", {
                    disabled:
                      !a.tables.length ||
                      a.measures.length >= dataModelLimits.measures,
                    onClick: () => {
                      const le = {
                        id: createUuid(),
                        name: `Measure ${a.measures.length + 1}`,
                        tableId: a.tables.at(-1).id,
                        operation: "countRows",
                      };
                      (_({ ...a, measures: [...a.measures, le] }),
                        C({ ...d, measureIds: [...d.measureIds, le.id] }));
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
                        children: a.measures.map((le) =>
                          i.jsxs(
                            "label",
                            {
                              children: [
                                i.jsx("input", {
                                  type: "checkbox",
                                  checked: d.measureIds.includes(le.id),
                                  onChange: (ce) =>
                                    C({
                                      ...d,
                                      measureIds: ce.target.checked
                                        ? [...d.measureIds, le.id]
                                        : d.measureIds.filter(
                                            (ge) => ge !== le.id,
                                          ),
                                    }),
                                }),
                                le.name,
                              ],
                            },
                            le.id,
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
                      d.groupBy.map((le, ce) =>
                        i.jsxs(
                          "div",
                          {
                            className: "lg-model-controls",
                            children: [
                              U(`Group ${ce + 1} table`, le.tableId, (ge) =>
                                C({
                                  ...d,
                                  groupBy: d.groupBy.map((se, we) =>
                                    we === ce
                                      ? { tableId: ge, column: N(ge) }
                                      : se,
                                  ),
                                }),
                              ),
                              G(
                                `Group ${ce + 1} column`,
                                le.tableId,
                                le.column,
                                (ge) =>
                                  C({
                                    ...d,
                                    groupBy: d.groupBy.map((se, we) =>
                                      we === ce ? { ...se, column: ge } : se,
                                    ),
                                  }),
                              ),
                              i.jsx("button", {
                                "aria-label": `Remove group ${ce + 1}`,
                                onClick: () =>
                                  C({
                                    ...d,
                                    groupBy: d.groupBy.filter(
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
                        disabled: !a.tables.length || d.groupBy.length >= 3,
                        onClick: () =>
                          C({
                            ...d,
                            groupBy: [
                              ...d.groupBy,
                              {
                                tableId: a.tables[0].id,
                                column: a.tables[0].columns[0].name,
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
                      d.filters.map((le, ce) => {
                        const ge = (Ae) =>
                            C({
                              ...d,
                              filters: d.filters.map((ye, Te) =>
                                Te === ce ? Ae : ye,
                              ),
                            }),
                          se = T(le.tableId)?.columns.find(
                            (Ae) => Ae.name === le.column,
                          )?.type;
                        return i.jsxs(
                          "div",
                          {
                            className: "lg-model-controls",
                            children: [
                              U(`Filter ${ce + 1} table`, le.tableId, (Ae) =>
                                ge({
                                  ...le,
                                  tableId: Ae,
                                  column: N(Ae),
                                  operator: "blank",
                                  value: undefined,
                                }),
                              ),
                              G(
                                `Filter ${ce + 1} column`,
                                le.tableId,
                                le.column,
                                (Ae) =>
                                  ge({
                                    ...le,
                                    column: Ae,
                                    operator: "blank",
                                    value: undefined,
                                  }),
                              ),
                              L(
                                `Filter ${ce + 1} operator`,
                                le.operator,
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
                                (Ae) =>
                                  ge({
                                    ...le,
                                    operator: Ae,
                                    value: ["blank", "notBlank"].includes(Ae)
                                      ? undefined
                                      : se === "number"
                                        ? 0
                                        : se === "boolean"
                                          ? false
                                          : "",
                                  }),
                              ),
                              !["blank", "notBlank"].includes(le.operator) &&
                                (se === "boolean"
                                  ? L(
                                      `Filter ${ce + 1} value`,
                                      String(le.value),
                                      $(["true", "false"]),
                                      (Ae) =>
                                        ge({ ...le, value: Ae === "true" }),
                                    )
                                  : R(
                                      `Filter ${ce + 1} value`,
                                      String(le.value ?? ""),
                                      (Ae) =>
                                        ge({
                                          ...le,
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
                                    ...d,
                                    filters: d.filters.filter(
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
                          !a.tables.length ||
                          d.filters.length >= dataModelLimits.filters,
                        onClick: () =>
                          C({
                            ...d,
                            filters: [
                              ...d.filters,
                              {
                                tableId: a.tables[0].id,
                                column: a.tables[0].columns[0].name,
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
                              children: W.diagnostics.map((le, ce) =>
                                i.jsxs(
                                  "tr",
                                  {
                                    children: [
                                      i.jsx("td", { children: le.code }),
                                      i.jsxs("td", {
                                        children: [
                                          T(le.tableId ?? "")?.name,
                                          " ",
                                          le.row ?? "",
                                        ],
                                      }),
                                      i.jsx("td", {
                                        children: text(le.message),
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
                            children: W.tables.map((le) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("th", {
                                      scope: "row",
                                      children: le.name,
                                    }),
                                    i.jsxs("td", {
                                      children: [
                                        e.workbook.sheets.find(
                                          (ge) => ge.id === le.sheetId,
                                        )?.name,
                                        "!",
                                        le.range,
                                      ],
                                    }),
                                    i.jsx("td", { children: le.rows }),
                                    i.jsx("td", { children: le.included }),
                                    i.jsx("td", { children: le.excluded }),
                                    i.jsx("td", { children: le.columnarBytes }),
                                  ],
                                },
                                le.id,
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
                                  children: W.calculatedColumns.map((le) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("th", {
                                            scope: "row",
                                            children: [
                                              T(le.tableId)?.name,
                                              ".",
                                              le.name,
                                            ],
                                          }),
                                          i.jsx("td", {
                                            children: le.values
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
                                      `${le.tableId}:${le.name}`,
                                    ),
                                  ),
                                }),
                              ],
                            }),
                          ],
                        }),
                      W.complete &&
                        d.measureIds.length > 0 &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("table", {
                              "aria-label": text("Model measure results"),
                              children: [
                                i.jsx("thead", {
                                  children: i.jsx("tr", {
                                    children: W.headers.map((le) =>
                                      i.jsx("th", { children: le }, le),
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
                                d.groupBy.length > 0 &&
                                  i.jsx("tfoot", {
                                    children: i.jsxs("tr", {
                                      children: [
                                        i.jsx("th", {
                                          colSpan: d.groupBy.length,
                                          children: i.jsx(LocalizedText, {
                                            children: "Grand total",
                                          }),
                                        }),
                                        W.totals.map((le, ce) =>
                                          i.jsx(
                                            "td",
                                            {
                                              children:
                                                le === null
                                                  ? text("Blank")
                                                  : le,
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
                    checked: b,
                    onChange: (le) => {
                      (y(le.target.checked), B());
                    },
                  }),
                  i.jsx(LocalizedText, { children: "Include hidden sources" }),
                ],
              }),
            LabeledField(
              "Output worksheet",
              R("Model output worksheet", O, (le) => {
                (j(le), B());
              }),
            ),
            i.jsx("button", {
              onClick: J,
              children: i.jsx(LocalizedText, { children: "Preview model" }),
            }),
            i.jsx("button", {
              disabled: !K,
              onClick: () => q("save"),
              children: i.jsx(LocalizedText, { children: "Save model" }),
            }),
            i.jsx("button", {
              className: "lg-model-primary",
              disabled: !K || !d.measureIds.length,
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
    const t = getComputedStyle(r);
    if (
      r.hidden ||
      t.display === "none" ||
      t.visibility === "hidden" ||
      t.visibility === "collapse"
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
function useDialogFocusTrap(e) {
  const r = ne.useRef(undefined);
  (ne.useLayoutEffect(() => {
    const t = e.current;
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
      c = (p) => {
        const m = p.target;
        if (m instanceof HTMLElement) {
          if (n && t.contains(m) && !n.contains(m)) {
            (getFocusableElements(n)[0] ?? n).focus({ preventScroll: true });
            return;
          }
          ((!n && !m.closest('[role="dialog"][aria-modal="true"]')) ||
            n?.contains(m)) &&
            (s = m);
        }
      },
      h = (p) => {
        if ((l(), !n || p.key !== "Tab")) return;
        const m = getFocusableElements(n),
          v = m[0],
          g = m.at(-1);
        if (!v) {
          (p.preventDefault(), (n.tabIndex = -1), n.focus());
          return;
        }
        (!n.contains(document.activeElement) ||
          (p.shiftKey && document.activeElement === v) ||
          (!p.shiftKey && document.activeElement === g)) &&
          (p.preventDefault(),
          p.stopPropagation(),
          (p.shiftKey ? g : v)?.focus());
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
  }, [e]),
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

function assertRecordKeys(e, r, t) {
  (!isNonArrayObject(e) || Object.keys(e).some((n) => !r.includes(n))) &&
    failWithMessage(t);
}
const rangesOverlap = (e, r) => {
  const t = parseRangeAddress(e),
    n = parseRangeAddress(r);
  return (
    t.firstRow <= n.lastRow &&
    t.lastRow >= n.firstRow &&
    t.firstCol <= n.lastCol &&
    t.lastCol >= n.firstCol
  );
};
function assertDelimiterCharacter(e) {
  (typeof e != "string" || [...e].length !== 1 || /["\r\n\0]/.test(e)) &&
    failWithMessage(
      "Choose one delimiter character other than a quote, line break or NUL.",
    );
}
function splitDelimitedCell(e, r) {
  (assertDelimiterCharacter(r),
    (typeof e != "string" || e.length > 32767) &&
      failWithMessage(
        "A source cell must contain text of at most 32,767 characters.",
      ));
  const t = [];
  let n = "",
    o = false,
    a = false,
    s = false;
  const d = [...e],
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
      : c === r
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
function validateTextToColumnsRequest(e) {
  (assertRecordKeys(
    e,
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
    (typeof e.sourceSheetId != "string" || typeof e.sourceRange != "string") &&
      failWithMessage("Choose a source worksheet and one-column range."),
    assertDelimiterCharacter(e.delimiter),
    ["none", "first-row"].includes(e.header) ||
      failWithMessage("Choose whether the first source row contains headings."),
    e.formats !== undefined &&
      (!Array.isArray(e.formats) ||
        e.formats.length > dataModelBudgets.fields ||
        e.formats.some((t) => !["text", "general"].includes(t))) &&
      failWithMessage(
        "Column conversion must be Text or General, with at most 50 columns.",
      ),
    assertRecordKeys(
      e.destination,
      ["kind", "name", "sheetId", "address"],
      "Choose a separate output range or a new worksheet.",
    ));
  let r;
  if (e.destination.kind === "new-sheet") {
    (assertRecordKeys(
      e.destination,
      ["kind", "name"],
      "Invalid new worksheet destination.",
    ),
      typeof e.destination.name != "string" &&
        failWithMessage("Enter an output worksheet name."));
    const t = e.destination.name.trim();
    ((!t ||
      t.length > 31 ||
      /[\u0000-\u001f\u007f\\/\[\]*?:]/.test(t) ||
      t.startsWith("'") ||
      t.endsWith("'")) &&
      failWithMessage(
        "Output worksheet names need 1–31 characters without control/Excel-reserved characters or leading/trailing apostrophes.",
      ),
      (r = { kind: "new-sheet", name: t }));
  } else if (e.destination.kind === "range") {
    (assertRecordKeys(
      e.destination,
      ["kind", "sheetId", "address"],
      "Invalid output range destination.",
    ),
      (typeof e.destination.sheetId != "string" ||
        typeof e.destination.address != "string") &&
        failWithMessage("Choose an output worksheet and one starting cell."));
    const t = _t(e.destination.address.trim());
    r = {
      kind: "range",
      sheetId: e.destination.sheetId,
      address: formatCellAddress(t.row, t.col),
    };
  } else failWithMessage("Choose a separate output range or a new worksheet.");
  return {
    sourceSheetId: e.sourceSheetId,
    sourceRange: formatRangeCornerCells(e.sourceRange),
    delimiter: e.delimiter,
    header: e.header,
    ...(e.formats ? { formats: [...e.formats] } : {}),
    destination: r,
  };
}
function validateTextToColumnsPermissions(e, r) {
  return (
    [r.readOnly, r.includeHidden, r.allowNewSheet].some(
      (t) => t !== undefined && typeof t != "boolean",
    ) && failWithMessage("Text to Columns permissions must be boolean."),
    r.readOnly && failWithMessage("This workbook is read-only."),
    r.allowedSheetIds !== undefined &&
      (!Array.isArray(r.allowedSheetIds) ||
        r.allowedSheetIds.some(
          (t) => typeof t != "string" || !e.sheets.some((n) => n.id === t),
        )) &&
      failWithMessage("The authorized worksheet context is invalid."),
    {
      includeHidden: r.includeHidden ?? false,
      allowNewSheet: r.allowNewSheet ?? true,
      allowedSheetIds: [
        ...new Set(r.allowedSheetIds ?? e.sheets.map((t) => t.id)),
      ].sort(),
    }
  );
}
function assertSelectionInAuthorizedSheet(e, r, t) {
  const n = parseRangeAddress(r);
  ((!t.allowedSheetIds.includes(e.id) ||
    (!t.includeHidden &&
      (e.hidden ||
        e.veryHidden ||
        cn(e, "row").some((o) => o.start <= n.lastRow && o.end >= n.firstRow) ||
        cn(e, "column").some(
          (o) => o.start <= n.lastCol && o.end >= n.firstCol,
        )))) &&
    failWithMessage(
      "The selected cells are outside the visible, authorized worksheet context.",
    ),
    (n.lastRow >= e.rowCount || n.lastCol >= e.colCount) &&
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
function assertOutputNameUnreferenced(e, r) {
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
          ).toLowerCase() === r.toLowerCase()
        );
      })
    ),
    o = [];
  for (const a of e.sheets) {
    for (const s of Object.values(a.cells))
      typeof s.raw == "string" &&
        s.raw.startsWith("=") &&
        !s.literal &&
        o.push(s.raw);
    for (const s of a.conditionalFormats ?? [])
      s.kind === "formula" && s.formula && o.push(s.formula);
    for (const s of a.tables ?? [])
      for (const d of s.columns)
        d.calculatedFormula && o.push(d.calculatedFormula);
  }
  for (const a of e.definedNames ?? []) o.push(a.refersTo);
  o.some(n) &&
    failWithMessage(
      "The output worksheet name is already referenced by workbook formulas or formula metadata. Choose another name to avoid reconnecting an unresolved source reference.",
    );
}
function parseImportedCellValue(e, r) {
  if (r === "general") {
    const t = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(e);
    if (/^(TRUE|FALSE)$/i.test(e))
      return { raw: e.toUpperCase() === "TRUE", precision: false };
    if (t) {
      const n = e
          .split(/[eE]/)[0]
          .replace(/[^0-9]/g, "")
          .replace(/^0+/, "").length,
        o = Number(e);
      return n > 15 ||
        !Number.isFinite(o) ||
        (o === 0 && /[1-9]/.test(e.split(/[eE]/)[0])) ||
        (Number.isInteger(o) && !Number.isSafeInteger(o))
        ? { raw: e, precision: true }
        : { raw: o, precision: false };
    }
  }
  return { raw: e, precision: false };
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
