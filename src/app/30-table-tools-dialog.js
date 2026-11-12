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
    l = ne.useRef(null);
  ne.useEffect(() => {
    const activeElement = document.activeElement,
      Le = l.current,
      et = Le?.querySelector(
        "input:not(:disabled), select:not(:disabled), button:not(:disabled)",
      );
    return (
      et?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const c = ne.useMemo(() => {
      try {
        return findTableForRange(workbook, sheetId, range);
      } catch {
        return range;
      }
    }, []),
    h = ne.useMemo(() => tryComputePivotRows(workbook, sheetId, c), []),
    [f, p] = ne.useState(sheetId),
    [m, v] = ne.useState(c),
    [fields, b] = ne.useState(h),
    [y, w] = ne.useState(
      mode === "table"
        ? nextDefaultTableName(workbook)
        : mode === "pivot"
          ? "Pivot summary"
          : "Extracted data",
    ),
    [tableId, A] = ne.useState(""),
    [k, F] = ne.useState("blue"),
    [I, O] = ne.useState(false),
    [totals, z] = ne.useState({}),
    [rowFields, D] = ne.useState(h.slice(0, 1)),
    [columnFields, _] = ne.useState([]),
    [values, T] = ne.useState(
      h.length ? [{ field: h.at(-1), operation: "sum" }] : [],
    ),
    [N, R] = ne.useState(true),
    [L, $] = ne.useState(false),
    [dateGroups, setDateGroups] = ne.useState([]),
    [calculatedFields, Q] = ne.useState([]);
  function setCalculatedFields(rt) {
    (T((Le) =>
      Le.flatMap((et) => {
        const st = calculatedFields.findIndex(
          (calculatedField) => calculatedField.name === et.field,
        );
        return st < 0
          ? [et]
          : rt.length === calculatedFields.length
            ? [{ ...et, field: rt[st].name }]
            : rt.some((kt) => kt.name === et.field)
              ? [et]
              : [];
      }),
    ),
      Q(rt));
  }
  const [J, q] = ne.useState([]),
    [W, K] = ne.useState(h),
    [rename, ae] = ne.useState({}),
    [field, de] = ne.useState(""),
    [direction, Se] = ne.useState("ascending"),
    [derived, ce] = ne.useState([]),
    [checked, se] = ne.useState(false),
    [we, Ee] = ne.useState(sheetId),
    [address, ye] = ne.useState("A1"),
    [pivotId, Pe] = ne.useState(""),
    [columnId, ht] = ne.useState(""),
    [Re, Me] = ne.useState(""),
    [appendCount, je] = ne.useState("1"),
    [$e, Je] = ne.useState(null),
    [ze, at] = ne.useState(0),
    context = JSON.stringify({
      sheetId,
      range,
      tableId: tableId,
      name: y,
      sourceRange: m,
      style: k,
      showTotals: I,
      totals: totals,
      columnId: columnId,
      columnFormula: Re,
      appendCount: appendCount,
    }),
    disabled =
      !!$e &&
      ($e.plan.snapshot !== JSON.stringify(workbook) || $e.context !== context),
    [lt, be] = ne.useState(""),
    [ft, pt] = ne.useState(""),
    [Fe, nt] = ne.useState(null),
    mt = workbook.sheets.find((sheet) => sheet.id === sheetId)?.tables ?? [],
    Wt = workbook.sheets.flatMap((sheet) =>
      (mode === "pivot" ? (sheet.pivots ?? []) : (sheet.queries ?? [])).map(
        (definition) => ({
          sheetName: sheet.name,
          definition: definition,
        }),
      ),
    ),
    Zt = () => (checked ? { sheetId: we, address: address } : undefined);
  function xt(rt, Le) {
    if (!readOnly) {
      (be(""), pt(""));
      try {
        const et = rt();
        (onWorkbookChange(et, Le), pt(`${Le} completed.`));
      } catch (error) {
        be(
          error instanceof Error
            ? error.message
            : "The operation could not be completed.",
        );
      }
    }
  }
  function onClick() {
    be("");
    try {
      const pivotSourceFields = getPivotSourceFields(workbook, f, m);
      (b(pivotSourceFields),
        K(pivotSourceFields),
        D(pivotSourceFields.slice(0, 1)),
        _([]),
        T(
          pivotSourceFields.length
            ? [{ field: pivotSourceFields.at(-1), operation: "sum" }]
            : [],
        ),
        q([]),
        ae({}),
        ce([]),
        de(""),
        setDateGroups([]),
        Q([]));
    } catch (error) {
      be(error.message);
    }
  }
  function Wr(rt) {
    (nt(null), A(rt), Je(null));
    const Le = mt.find((kt) => kt.id === rt);
    if (!Le) {
      (w(nextDefaultTableName(workbook)), v(c), F("blue"), O(false), z({}));
      return;
    }
    (w(Le.name),
      ht(Le.columns[0]?.id ?? ""),
      Me(Le.columns[0]?.calculatedFormula ?? ""),
      v(Le.range),
      b(Le.columns.map((column) => column.name)),
      F(Le.style),
      O(Le.showTotals),
      z(Le.totals));
  }
  function Rr() {
    return J.map((rt) => ({
      ...rt,
      value: ["gt", "gte", "lt", "lte"].includes(rt.operator)
        ? Number(rt.value)
        : rt.value,
    }));
  }
  function qt(request) {
    (be(""), pt(""), Je(null), at(0));
    try {
      Je({
        plan: applyTableOperation(workbook, sheetId, tableId, request, {
          readOnly,
        }),
        context: context,
      });
    } catch (error) {
      be(error.message);
    }
  }
  function Ye() {
    if ($e) {
      be("");
      try {
        if (disabled)
          throw Error(
            "Table settings or the current selection changed. Preview again.",
          );
        const rt = applyTableOperationPreview(workbook, $e.plan, { readOnly });
        (rt !== workbook &&
          onWorkbookChange(rt, "Apply reviewed table operation"),
          v($e.plan.afterRange),
          Je(null),
          pt(
            rt === workbook
              ? "No table changes were needed."
              : "Reviewed table operation completed.",
          ));
      } catch (error) {
        be(error.message);
      }
    }
  }
  function bt() {
    mode === "table"
      ? tableId
        ? qt({
            kind: "updateTable",
            options: {
              name: y,
              range: m,
              style: k,
              showTotals: I,
              totals: totals,
            },
          })
        : xt(
            () =>
              createWorksheetTable(workbook, sheetId, {
                name: y,
                range: m,
                style: k,
                showTotals: I,
                totals: totals,
                ...(Fe
                  ? {
                      headers: (() => {
                        if (
                          Fe.snapshot !== JSON.stringify(workbook) ||
                          Fe.range !== m
                        )
                          throw Error(
                            "The workbook or source range changed. Review headers again.",
                          );
                        return Fe.headers.map((header) => header.suggested);
                      })(),
                    }
                  : {}),
              }),
            "Create table",
          )
      : mode === "pivot"
        ? xt(
            () =>
              buildPivotTableOutput(workbook, {
                name: y,
                source: { sheetId: f, range: m },
                rowFields: rowFields,
                columnFields: columnFields,
                values: values,
                filters: Rr(),
                grandTotals: N,
                subtotals: L,
                dateGroups: dateGroups,
                calculatedFields: calculatedFields,
                destination: Zt(),
              }),
            "Create pivot summary",
          )
        : xt(
            () =>
              runQueryDefinition(workbook, {
                name: y,
                source: { sheetId: f, range: m },
                fields: W,
                rename: rename,
                filters: Rr(),
                ...(field
                  ? { sort: { field: field, direction: direction } }
                  : {}),
                derived: derived,
                destination: Zt(),
              }),
            "Extract data",
          );
  }
  const St = (rt) =>
      i.jsxs(i.Fragment, {
        children: [
          rt !== undefined && i.jsx("option", { value: "", children: rt }),
          fields.map((Le) => i.jsx("option", { value: Le, children: Le }, Le)),
        ],
      }),
    cr = (rt, Le, et) =>
      i.jsxs("fieldset", {
        className: "lg-table-fields",
        children: [
          i.jsx("legend", { children: rt }),
          fields.length
            ? fields.map((st) =>
                i.jsxs(
                  "label",
                  {
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: Le.includes(st),
                        onChange: () => et(toggleListValue(Le, st)),
                      }),
                      st,
                    ],
                  },
                  st,
                ),
              )
            : i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children: "Read the source fields first.",
                }),
              }),
        ],
      }),
    dr =
      mode === "table"
        ? "Table tools"
        : mode === "pivot"
          ? "PivotTable builder"
          : "From Table / Range";
  return i.jsxs("div", {
    ref: l,
    className: "lg-table-tools",
    role: "dialog",
    "aria-modal": "true",
    "aria-label": dr,
    onKeyDown: (event) => {
      if (
        (event.key === "Escape" &&
          (event.preventDefault(), event.stopPropagation(), onClose()),
        event.key === "Tab")
      ) {
        const et = Array.from(
            l.current?.querySelectorAll(
              "button:not(:disabled), input:not(:disabled), select:not(:disabled), summary, [tabindex='0']",
            ) ?? [],
          ).filter(
            (gr) =>
              (!gr.closest("fieldset:disabled") &&
                !gr.closest("details:not([open])")) ||
              gr.tagName === "SUMMARY",
          ),
          st = et[0],
          kt = et.at(-1);
        event.shiftKey && document.activeElement === st && kt
          ? (event.preventDefault(), kt.focus())
          : !event.shiftKey &&
            document.activeElement === kt &&
            st &&
            (event.preventDefault(), st.focus());
      }
    },
    children: [
      i.jsxs("header", {
        children: [
          i.jsxs("div", {
            children: [
              i.jsx("h2", { children: dr }),
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
            "aria-label": `Close ${dr}`,
            onClick: onClose,
            children: "×",
          }),
        ],
      }),
      lt &&
        i.jsx("p", {
          role: "alert",
          className: "lg-table-error",
          children: lt,
        }),
      ft && i.jsx("p", { role: "status", children: ft }),
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
                  onChange: (event) => Wr(event.target.value),
                  lang: language("Existing table"),
                  "data-ui-label": "Existing table",
                  children: [
                    i.jsx("option", {
                      value: "",
                      children: "Create a new table",
                    }),
                    mt.map((rt) =>
                      i.jsx(
                        "option",
                        { value: rt.id, children: rt.name },
                        rt.id,
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
                value: y,
                onChange: (event) => w(event.target.value),
              }),
            ],
          }),
          mode !== "table" &&
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source sheet" }),
                i.jsx("select", {
                  value: f,
                  onChange: (event) => p(event.target.value),
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
                    value: m,
                    onChange: (event) => v(event.target.value),
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
                            be("");
                            try {
                              const headers = suggestTableHeaderName(
                                workbook,
                                sheetId,
                                m,
                              );
                              (nt({
                                snapshot: JSON.stringify(workbook),
                                range: m,
                                headers: headers,
                              }),
                                b(headers.map((header) => header.suggested)));
                            } catch (error) {
                              be(error.message);
                            }
                          },
                          children: i.jsx(LocalizedText, {
                            children: "Review headers",
                          }),
                        }),
                        Fe &&
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
                              Fe.headers.map((header, Le) =>
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
                                        "aria-label": `Reviewed header ${Le + 1}`,
                                        value: header.suggested,
                                        onChange: (event) => {
                                          const headers = Fe.headers.map(
                                            (kt, gr) =>
                                              gr === Le
                                                ? {
                                                    ...kt,
                                                    suggested:
                                                      event.target.value,
                                                  }
                                                : kt,
                                          );
                                          (nt({ ...Fe, headers: headers }),
                                            b(
                                              headers.map((kt) => kt.suggested),
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
                        value: k,
                        onChange: (event) => F(event.target.value),
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
                        checked: I,
                        onChange: (event) => O(event.target.checked),
                      }),
                      i.jsx(LocalizedText, { children: "Show totals row" }),
                    ],
                  }),
                  I &&
                    i.jsx("div", {
                      className: "lg-table-list",
                      children: fields.map((rt) =>
                        i.jsxs(
                          "label",
                          {
                            children: [
                              rt,
                              " ",
                              i.jsx(LocalizedText, { children: "total" }),
                              i.jsxs("select", {
                                "aria-label": `${rt} total`,
                                value: Object.hasOwn(totals, rt)
                                  ? totals[rt]
                                  : "",
                                onChange: (event) =>
                                  z((et) => {
                                    const st = { ...et };
                                    return (
                                      event.target.value
                                        ? Object.defineProperty(st, rt, {
                                            value: event.target.value,
                                            enumerable: true,
                                            writable: true,
                                            configurable: true,
                                          })
                                        : delete st[rt],
                                      st
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
                          rt,
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
                                (ht(event.target.value),
                                  Me(
                                    mt
                                      .find((st) => st.id === tableId)
                                      ?.columns.find(
                                        (column) =>
                                          column.id === event.target.value,
                                      )?.calculatedFormula ?? "",
                                  ),
                                  Je(null));
                              },
                              lang: language("Calculated column"),
                              "data-ui-label": "Calculated column",
                              children: mt
                                .find((rt) => rt.id === tableId)
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
                              value: Re,
                              onChange: (event) => Me(event.target.value),
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
                                qt({
                                  kind: "setColumnFormula",
                                  columnId: columnId,
                                  formula: Re,
                                }),
                              children: i.jsx(LocalizedText, {
                                children: "Preview column formula",
                              }),
                            }),
                            i.jsx("button", {
                              type: "button",
                              disabled: !mt
                                .find((rt) => rt.id === tableId)
                                ?.columns.find(
                                  (column) => column.id === columnId,
                                )?.calculatedFormula,
                              onClick: () =>
                                qt({
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
                              onValueChange: (rt) => je(rt),
                              lang: language("Rows to append"),
                              "data-ui-label": "Rows to append",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () =>
                            qt({
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
                  $e &&
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
                                    $e.plan.beforeRange,
                                    " →",
                                    " ",
                                    $e.plan.afterRange,
                                    ";",
                                    " ",
                                    $e.plan.changes.length,
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
                                    $e.plan.afterColumns.map(
                                      (afterColumn, Le) =>
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
                                              $e.plan.beforeColumns[Le]
                                                ?.calculatedFormula ??
                                                "No rule",
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
                                  start: ze * 20 + 1,
                                  children: $e.plan.changes
                                    .slice(ze * 20, ze * 20 + 20)
                                    .map((rt) =>
                                      i.jsxs(
                                        "li",
                                        {
                                          children: [
                                            i.jsxs("strong", {
                                              children: [
                                                workbook.sheets.find(
                                                  (sheet) =>
                                                    sheet.id === rt.sheetId,
                                                )?.name,
                                                "!",
                                                rt.address,
                                              ],
                                            }),
                                            i.jsx("br", {}),
                                            i.jsx(LocalizedText, {
                                              children: "Before: ",
                                            }),
                                            JSON.stringify(
                                              rt.before?.raw ?? null,
                                            ),
                                            i.jsx("br", {}),
                                            i.jsx(LocalizedText, {
                                              children: "After: ",
                                            }),
                                            JSON.stringify(
                                              rt.after?.raw ?? null,
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
                                                      before: rt.before ?? null,
                                                      after: rt.after ?? null,
                                                    },
                                                    null,
                                                    2,
                                                  ),
                                                }),
                                              ],
                                            }),
                                          ],
                                        },
                                        `${rt.sheetId}:${rt.address}`,
                                      ),
                                    ),
                                }),
                                $e.plan.changes.length > 20 &&
                                  i.jsxs("div", {
                                    className: "lg-table-inline",
                                    children: [
                                      i.jsx("button", {
                                        type: "button",
                                        disabled: ze === 0,
                                        onClick: () => at((rt) => rt - 1),
                                        children: i.jsx(LocalizedText, {
                                          children: "Previous changes",
                                        }),
                                      }),
                                      i.jsxs("span", {
                                        children: [
                                          i.jsx(LocalizedText, {
                                            children: "Page ",
                                          }),
                                          ze + 1,
                                          " ",
                                          i.jsx(LocalizedText, {
                                            children: "of",
                                          }),
                                          " ",
                                          Math.ceil(
                                            $e.plan.changes.length / 20,
                                          ),
                                        ],
                                      }),
                                      i.jsx("button", {
                                        type: "button",
                                        disabled:
                                          (ze + 1) * 20 >=
                                          $e.plan.changes.length,
                                        onClick: () => at((rt) => rt + 1),
                                        children: i.jsx(LocalizedText, {
                                          children: "Next changes",
                                        }),
                                      }),
                                    ],
                                  }),
                                $e.plan.anchorChanges.map((anchorChange) =>
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
                              onClick: Ye,
                              children: i.jsx(LocalizedText, {
                                children: "Apply table changes",
                              }),
                            }),
                            i.jsx("button", {
                              type: "button",
                              onClick: () => Je(null),
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
                          cr("Row groups", rowFields, D),
                          cr("Column groups", columnFields, _),
                          i.jsxs("fieldset", {
                            className: "lg-table-list",
                            children: [
                              i.jsx("legend", {
                                children: i.jsx(LocalizedText, {
                                  children: "Value fields",
                                }),
                              }),
                              values.map((value, Le) =>
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
                                          Le + 1,
                                          i.jsx("select", {
                                            value: value.field,
                                            onChange: (event) =>
                                              T((st) =>
                                                st.map((kt, gr) =>
                                                  gr === Le
                                                    ? {
                                                        ...kt,
                                                        field:
                                                          event.target.value,
                                                      }
                                                    : kt,
                                                ),
                                              ),
                                            children: fields
                                              .concat(
                                                calculatedFields.map(
                                                  (calculatedField) =>
                                                    calculatedField.name,
                                                ),
                                              )
                                              .map((et) =>
                                                i.jsx(
                                                  "option",
                                                  { value: et, children: et },
                                                  et,
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
                                            "aria-label": `Aggregation ${Le + 1}`,
                                            value: value.operation,
                                            onChange: (event) =>
                                              T((st) =>
                                                st.map((kt, gr) =>
                                                  gr === Le
                                                    ? {
                                                        ...kt,
                                                        operation:
                                                          event.target.value,
                                                      }
                                                    : kt,
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
                                        "aria-label": `Remove value ${Le + 1}`,
                                        onClick: () =>
                                          T((et) =>
                                            et.filter((st, kt) => kt !== Le),
                                          ),
                                        children: i.jsx(LocalizedText, {
                                          children: "Remove",
                                        }),
                                      }),
                                    ],
                                  },
                                  Le,
                                ),
                              ),
                              i.jsx("button", {
                                type: "button",
                                disabled:
                                  !fields.length || values.length >= 100,
                                onClick: () =>
                                  T((rt) => [
                                    ...rt,
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
                                checked: N,
                                onChange: (event) => R(event.target.checked),
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
                                checked: L,
                                onChange: (event) => $(event.target.checked),
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
                            onAdd: (rt) =>
                              T((Le) => [
                                ...Le,
                                { field: rt, operation: "sum" },
                              ]),
                          }),
                        ],
                      })
                    : i.jsxs(i.Fragment, {
                        children: [
                          cr("Output fields", W, (rt) => {
                            (K(rt),
                              ae((Le) =>
                                Object.fromEntries(
                                  Object.entries(Le).filter(([et]) =>
                                    rt.includes(et),
                                  ),
                                ),
                              ));
                          }),
                          i.jsxs("details", {
                            children: [
                              i.jsx("summary", {
                                children: i.jsx(LocalizedText, {
                                  children: "Rename output fields",
                                }),
                              }),
                              i.jsx("div", {
                                className: "lg-table-list",
                                children: W.map((rt) =>
                                  i.jsxs(
                                    "label",
                                    {
                                      children: [
                                        rt,
                                        i.jsx("input", {
                                          "aria-label": `Rename ${rt}`,
                                          value: Object.hasOwn(rename, rt)
                                            ? rename[rt]
                                            : rt,
                                          onChange: (event) =>
                                            ae((et) =>
                                              Object.defineProperty(
                                                { ...et },
                                                rt,
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
                                    rt,
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
                                    onChange: (event) => de(event.target.value),
                                    children: St("Source order"),
                                  }),
                                ],
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, { children: "Order" }),
                                  i.jsxs("select", {
                                    value: direction,
                                    onChange: (event) => Se(event.target.value),
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
                              derived.map((rt, Le) =>
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
                                            value: rt.name,
                                            onChange: (event) =>
                                              ce((st) =>
                                                st.map((kt, gr) =>
                                                  gr === Le
                                                    ? {
                                                        ...kt,
                                                        name: event.target
                                                          .value,
                                                      }
                                                    : kt,
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
                                            value: rt.field,
                                            onChange: (event) =>
                                              ce((st) =>
                                                st.map((kt, gr) =>
                                                  gr === Le
                                                    ? {
                                                        ...kt,
                                                        field:
                                                          event.target.value,
                                                      }
                                                    : kt,
                                                ),
                                              ),
                                            children: St(),
                                          }),
                                        ],
                                      }),
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(LocalizedText, {
                                            children: "Operation",
                                          }),
                                          i.jsx("select", {
                                            value: rt.operation,
                                            onChange: (event) =>
                                              ce((st) =>
                                                st.map((kt, gr) =>
                                                  gr === Le
                                                    ? {
                                                        ...kt,
                                                        operation:
                                                          event.target.value,
                                                      }
                                                    : kt,
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
                                            ].map((et) =>
                                              i.jsx(
                                                "option",
                                                {
                                                  value: et,
                                                  children:
                                                    formatAggregationLabel(et),
                                                },
                                                et,
                                              ),
                                            ),
                                          }),
                                        ],
                                      }),
                                      !["upper", "lower", "trim"].includes(
                                        rt.operation,
                                      ) &&
                                        i.jsxs(i.Fragment, {
                                          children: [
                                            i.jsxs("label", {
                                              children: [
                                                i.jsx(LocalizedText, {
                                                  children: "Second field",
                                                }),
                                                i.jsx("select", {
                                                  value: rt.otherField ?? "",
                                                  onChange: (event) =>
                                                    ce((st) =>
                                                      st.map((kt, gr) =>
                                                        gr === Le
                                                          ? {
                                                              ...kt,
                                                              otherField:
                                                                event.target
                                                                  .value ||
                                                                undefined,
                                                            }
                                                          : kt,
                                                      ),
                                                    ),
                                                  children:
                                                    St("Use a constant"),
                                                }),
                                              ],
                                            }),
                                            !rt.otherField &&
                                              i.jsxs("label", {
                                                children: [
                                                  i.jsx(LocalizedText, {
                                                    children: "Constant",
                                                  }),
                                                  rt.operation === "concat"
                                                    ? i.jsx("input", {
                                                        value:
                                                          rt.constant ?? "",
                                                        onChange: (event) =>
                                                          ce((st) =>
                                                            st.map((kt, gr) =>
                                                              gr === Le
                                                                ? {
                                                                    ...kt,
                                                                    constant:
                                                                      event
                                                                        .target
                                                                        .value,
                                                                  }
                                                                : kt,
                                                            ),
                                                          ),
                                                      })
                                                    : i.jsx(
                                                        LocalizedNumberField,
                                                        {
                                                          "aria-label":
                                                            text("Constant"),
                                                          value:
                                                            rt.constant ?? "",
                                                          onValueChange: (et) =>
                                                            ce((st) =>
                                                              st.map(
                                                                (kt, gr) =>
                                                                  gr === Le
                                                                    ? {
                                                                        ...kt,
                                                                        constant:
                                                                          Number(
                                                                            et,
                                                                          ),
                                                                      }
                                                                    : kt,
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
                                          ce((et) =>
                                            et.filter((st, kt) => kt !== Le),
                                          ),
                                        children: i.jsx(LocalizedText, {
                                          children: "Remove derived field",
                                        }),
                                      }),
                                    ],
                                  },
                                  Le,
                                ),
                              ),
                              i.jsx("button", {
                                type: "button",
                                disabled:
                                  !fields.length || derived.length >= 100,
                                onClick: () =>
                                  ce((rt) => [
                                    ...rt,
                                    {
                                      name: `Calculated${rt.length + 1}`,
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
                      J.map((rt, Le) =>
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
                                  Le + 1,
                                  i.jsx("select", {
                                    value: rt.field,
                                    onChange: (event) =>
                                      q((st) =>
                                        st.map((kt, gr) =>
                                          gr === Le
                                            ? {
                                                ...kt,
                                                field: event.target.value,
                                              }
                                            : kt,
                                        ),
                                      ),
                                    children: St(),
                                  }),
                                ],
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Condition",
                                  }),
                                  i.jsx("select", {
                                    value: rt.operator,
                                    onChange: (event) =>
                                      q((st) =>
                                        st.map((kt, gr) =>
                                          gr === Le
                                            ? {
                                                ...kt,
                                                operator: event.target.value,
                                              }
                                            : kt,
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
                                    ].map(([et, st]) =>
                                      i.jsx(
                                        "option",
                                        { value: et, children: st },
                                        et,
                                      ),
                                    ),
                                  }),
                                ],
                              }),
                              !["blank", "notBlank"].includes(rt.operator) &&
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Filter value",
                                    }),
                                    i.jsx("input", {
                                      value: String(rt.value ?? ""),
                                      onChange: (event) =>
                                        q((st) =>
                                          st.map((kt, gr) =>
                                            gr === Le
                                              ? {
                                                  ...kt,
                                                  value: event.target.value,
                                                }
                                              : kt,
                                          ),
                                        ),
                                    }),
                                  ],
                                }),
                              i.jsx("button", {
                                type: "button",
                                "aria-label": `Remove filter ${Le + 1}`,
                                onClick: () =>
                                  q((et) => et.filter((st, kt) => kt !== Le)),
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            ],
                          },
                          Le,
                        ),
                      ),
                      i.jsx("button", {
                        type: "button",
                        disabled: !fields.length || J.length >= 100,
                        onClick: () =>
                          q((rt) => [
                            ...rt,
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
                        onChange: (event) => se(event.target.checked),
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
                              value: we,
                              onChange: (event) => Ee(event.target.value),
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
                              onChange: (event) => ye(event.target.value),
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
                  Wt.length > 0 &&
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
                              onChange: (event) => Pe(event.target.value),
                              children: [
                                i.jsx("option", {
                                  value: "",
                                  children: "Choose a result",
                                }),
                                Wt.map(({ sheetName, definition }) =>
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
                            xt(
                              () =>
                                mode === "pivot"
                                  ? refreshPivotTable(workbook, pivotId, Zt())
                                  : refreshQueryDefinition(
                                      workbook,
                                      pivotId,
                                      Zt(),
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
              onClick: bt,
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
  formatPivotGroupLabel = (e) =>
    e === null
      ? "(blank)"
      : e === ""
        ? "(empty text)"
        : `${typeof e}: ${String(e)}`;
function PivotChartFigure({ book, id: pivotId }) {
  const { text, language } = useWorkbookMessages();
  let o;
  try {
    o = buildPivotChartModel(book, pivotId);
  } catch (error) {
    return i.jsx("p", { role: "alert", children: error.message });
  }
  const a = o.points.map((point) => point.value),
    s = a.filter((m) => m !== null),
    d = a
      .reduce(
        (m, v, g) => (
          v === null ? m.push([]) : m[m.length - 1].push({ v, i: g }),
          m
        ),
        [[]],
      )
      .filter((m) => m.length),
    u = Math.min(0, ...s),
    l = Math.max(1, ...s),
    c = (m) => 180 - ((m - u) / (l - u)) * 145,
    h = 420 / Math.max(1, a.length),
    f = c(0),
    p = o.showAs !== "value";
  return i.jsxs("figure", {
    className: "lg-pivot-chart",
    "aria-label": text("PivotChart"),
    lang: language("PivotChart"),
    "data-ui-label": "PivotChart",
    children: [
      i.jsxs("figcaption", {
        children: [
          o.measure,
          p ? ` — ${pivotShowAsLabels[o.showAs]}` : "",
          " ",
          i.jsx(LocalizedText, { children: "· source revision" }),
          " ",
          o.sourceRevision,
        ],
      }),
      a.length
        ? i.jsxs("svg", {
            role: "img",
            "aria-label": `${o.kind} PivotChart with ${a.length} groups; values are listed below`,
            viewBox: "0 0 470 210",
            children: [
              o.kind === "bar"
                ? i.jsx("line", {
                    x1: 35 - (u / (l - u)) * 420,
                    x2: 35 - (u / (l - u)) * 420,
                    y1: "28",
                    y2: "180",
                    stroke: "#607780",
                  })
                : i.jsx("line", {
                    x1: "35",
                    x2: "455",
                    y1: f,
                    y2: f,
                    stroke: "#607780",
                  }),
              i.jsx("text", {
                x: o.kind === "bar" ? 455 : 30,
                y: o.kind === "bar" ? 198 : c(l) + 4,
                textAnchor: "end",
                fontSize: "10",
                children: p ? `${(l * 100).toFixed(0)}%` : l.toFixed(1),
              }),
              i.jsx("text", {
                x: o.kind === "bar" ? 35 : 30,
                y: o.kind === "bar" ? 198 : c(u) + 4,
                textAnchor: o.kind === "bar" ? "start" : "end",
                fontSize: "10",
                children: p ? `${(u * 100).toFixed(0)}%` : u.toFixed(1),
              }),
              o.kind === "line"
                ? d.map((m, v) =>
                    i.jsx(
                      "polyline",
                      {
                        fill: "none",
                        stroke: "#217346",
                        strokeWidth: "3",
                        points: m
                          .map(({ v, i }) => `${35 + h * (i + 0.5)},${c(v)}`)
                          .join(" "),
                      },
                      v,
                    ),
                  )
                : a.map((m, v) =>
                    m === null
                      ? null
                      : o.kind === "bar"
                        ? i.jsx(
                            "rect",
                            {
                              x: 35 + ((Math.min(m, 0) - u) / (l - u)) * 420,
                              y: 28 + v * (150 / a.length),
                              width: Math.max(1, (Math.abs(m) / (l - u)) * 420),
                              height: Math.max(1, 120 / a.length),
                              fill: "#217346",
                              children: i.jsxs("title", {
                                children: [o.points[v].label, ": ", m],
                              }),
                            },
                            v,
                          )
                        : i.jsx(
                            "rect",
                            {
                              x: 36 + h * v,
                              width: Math.max(1, h - 3),
                              y: Math.min(c(m), f),
                              height: Math.max(1, Math.abs(c(m) - f)),
                              fill: "#217346",
                              children: i.jsxs("title", {
                                children: [o.points[v].label, ": ", m],
                              }),
                            },
                            v,
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
                children: o.points.map((point, v) =>
                  i.jsxs(
                    "tr",
                    {
                      children: [
                        i.jsx("td", { children: point.label }),
                        i.jsx("td", {
                          children:
                            point.value === null
                              ? "Blank"
                              : p
                                ? `${(Number(point.value) * 100).toFixed(2)}%`
                                : String(point.value),
                        }),
                      ],
                    },
                    v,
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
