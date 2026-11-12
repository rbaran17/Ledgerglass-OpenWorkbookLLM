import i from "react/jsx-runtime";
import ne from "react";
import {
  buildSlicerItemKey,
  applyPivotViewChange,
  buildPivotDrillDown,
  inspectChartSource,
  inspectPivotControlFields,
  _t,
  extractPivotDrillThrough,
  parseRangeAddress,
  describePivotRefresh,
  getPivotSlicerItems,
  supportedChartKinds,
  buildTimelinePeriods,
  getPivotSourceFields,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import { PivotStatusBadge } from "./29-worksheet-tools-and-pivot-panels.js";
import {
  pivotShowAsLabels,
  PivotChartFigure,
  formatPivotGroupLabel,
} from "./30-table-tools-dialog.js";

function PivotExplorerDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = e.workbook.sheets.flatMap((u) =>
      (u.pivots ?? []).map((l) => ({ sheet: u, definition: l })),
    ),
    [o, a] = ne.useState(() => {
      const u = _t(e.range.split(":")[0]);
      return (
        n.find((f) => {
          const p = parseRangeAddress(f.definition.outputRange);
          return (
            f.sheet.id === e.sheetId &&
            u.row >= p.firstRow &&
            u.row <= p.lastRow &&
            u.col >= p.firstCol &&
            u.col <= p.lastCol
          );
        })?.definition.id ??
        n.find((f) => f.sheet.id === e.sheetId)?.definition.id ??
        n[0]?.definition.id ??
        ""
      );
    }),
    s = n.find((u) => u.definition.id === o)?.definition;
  return i.jsxs("div", {
    role: "dialog",
    "aria-modal": "false",
    "aria-label": text("Explore PivotTable"),
    className: "lg-table-tools lg-pivot-explore",
    onKeyDown: (u) => {
      u.key === "Escape" && e.onClose();
    },
    lang: language("Explore PivotTable"),
    "data-ui-label": "Explore PivotTable",
    children: [
      i.jsxs("header", {
        children: [
          i.jsxs("div", {
            children: [
              i.jsx("h2", {
                children: i.jsx(LocalizedText, {
                  children: "Explore PivotTable",
                }),
              }),
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "Filter, reconcile and inspect the records behind a result.",
                }),
              }),
            ],
          }),
          i.jsx("button", {
            "aria-label": text("Close Explore PivotTable"),
            onClick: e.onClose,
            lang: language("Close Explore PivotTable"),
            "data-ui-label": "Close Explore PivotTable",
            children: "×",
          }),
        ],
      }),
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: "Pivot result" }),
          i.jsx("select", {
            value: o,
            onChange: (u) => a(u.target.value),
            children: n.map((u) =>
              i.jsxs(
                "option",
                {
                  value: u.definition.id,
                  children: [u.sheet.name, ": ", u.definition.name],
                },
                u.definition.id,
              ),
            ),
          }),
        ],
      }),
      s
        ? i.jsx(PivotExplorerPanel, { ...e, definition: s }, o)
        : i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children: "Create a PivotTable to explore its results.",
            }),
          }),
    ],
  });
}
function PivotExplorerPanel(e) {
  const r = e.definition,
    t = r.id,
    n = ne.useRef(null),
    [o, a] = ne.useState(() =>
      structuredClone({
        slicers: r.slicers ?? [],
        timelines: r.timelines ?? [],
        sorts: r.sorts ?? [],
        values: r.values,
        chart: r.chart,
      }),
    ),
    [s, d] = ne.useState(""),
    [u, l] = ne.useState(e.range.split(":")[0]),
    [c, h] = ne.useState(0),
    [f, p] = ne.useState(null),
    [m, v] = ne.useState(() => {
      try {
        return {
          report: buildPivotDrillDown(e.workbook, t, u),
          snapshot: JSON.stringify(e.workbook),
        };
      } catch {
        return null;
      }
    }),
    g = JSON.stringify(e.workbook),
    b = JSON.stringify(o),
    y = f?.snapshot === g && f.request === b,
    w = m?.snapshot === g && m.report.address === u.toUpperCase(),
    E = ne.useMemo(() => {
      try {
        return getPivotSourceFields(
          e.workbook,
          r.source.sheetId,
          r.source.range,
        );
      } catch {
        return [];
      }
    }, [e.workbook, r.source.sheetId, r.source.range]);
  ne.useEffect(() => {
    n.current?.querySelector("input")?.focus();
  }, []);
  const A = (D) => {
      try {
        (D(), d(""));
      } catch (B) {
        d(B.message);
      }
    },
    k = (D, B) =>
      A(() => {
        if (e.onExport) e.onExport(structuredClone(D), B);
        else {
          const _ = URL.createObjectURL(
              new Blob([JSON.stringify(D, null, 2)], {
                type: "application/json",
              }),
            ),
            C = document.createElement("a");
          ((C.href = _),
            (C.download = B),
            C.click(),
            setTimeout(() => URL.revokeObjectURL(_), 1e3));
        }
      }),
    F = JSON.stringify([
      ...new Set(
        [...(o.slicers ?? []), ...(o.timelines ?? [])].map((D) => D.field),
      ),
    ]),
    I = ne.useMemo(() => {
      try {
        return inspectPivotControlFields(e.workbook, t, JSON.parse(F));
      } catch (D) {
        return new Map(
          JSON.parse(F).map((B) => [B, { items: [], error: D.message }]),
        );
      }
    }, [e.workbook, t, F]),
    O = (D) => I.get(D)?.items ?? [],
    j = (D) => I.get(D)?.error ?? "",
    z = (D, B) => a((_) => ({ ..._, [D]: B })),
    V = y ? f.book.sheets.flatMap((D) => D.pivots ?? []).at(-1) : null;
  return i.jsxs("div", {
    ref: n,
    className: "lg-table-form",
    children: [
      i.jsx(PivotStatusBadge, { workbook: e.workbook, pivotId: t }),
      s && i.jsx("p", { role: "alert", children: s }),
      e.readOnly &&
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Read-only: inspect and export existing results; changes and extraction are disabled.",
          }),
        }),
      i.jsx("button", {
        onClick: () =>
          k(describePivotRefresh(e.workbook, t), "Pivot-filter-context.json"),
        children: i.jsx(LocalizedText, { children: "Export pivot context" }),
      }),
      i.jsxs("fieldset", {
        disabled: e.readOnly,
        className: "lg-table-list",
        children: [
          i.jsx("legend", {
            children: i.jsx(LocalizedText, { children: "Pivot view" }),
          }),
          i.jsxs("details", {
            open: true,
            children: [
              i.jsx("summary", {
                children: i.jsx(LocalizedText, { children: "Slicers" }),
              }),
              (o.slicers ?? []).map((D, B) =>
                i.jsxs(
                  "fieldset",
                  {
                    children: [
                      i.jsxs("legend", {
                        children: [
                          i.jsx(LocalizedText, { children: "Slicer " }),
                          B + 1,
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Slicer field " }),
                          B + 1,
                          i.jsxs("select", {
                            value: D.field,
                            onChange: (_) =>
                              A(() =>
                                z(
                                  "slicers",
                                  o.slicers.map((C, T) =>
                                    B === T
                                      ? {
                                          field: _.target.value,
                                          selected: getPivotSlicerItems(
                                            e.workbook,
                                            t,
                                            _.target.value,
                                          ),
                                        }
                                      : C,
                                  ),
                                ),
                              ),
                            children: [
                              !E.includes(D.field) &&
                                i.jsx("option", { children: D.field }),
                              E.map((_) => i.jsx("option", { children: _ }, _)),
                            ],
                          }),
                        ],
                      }),
                      i.jsx("div", {
                        className: "lg-pivot-items",
                        children: O(D.field).map((_, C) =>
                          i.jsxs(
                            "label",
                            {
                              className: "lg-table-check",
                              children: [
                                i.jsx("input", {
                                  type: "checkbox",
                                  checked: D.selected.some(
                                    (T) =>
                                      buildSlicerItemKey(T) ===
                                      buildSlicerItemKey(_),
                                  ),
                                  onChange: (T) =>
                                    z(
                                      "slicers",
                                      o.slicers.map((N, R) =>
                                        B === R
                                          ? {
                                              ...N,
                                              selected: T.target.checked
                                                ? [...N.selected, _]
                                                : N.selected.filter(
                                                    (L) =>
                                                      buildSlicerItemKey(L) !==
                                                      buildSlicerItemKey(_),
                                                  ),
                                            }
                                          : N,
                                      ),
                                    ),
                                }),
                                formatPivotGroupLabel(_),
                              ],
                            },
                            C,
                          ),
                        ),
                      }),
                      j(D.field) &&
                        i.jsx("p", { role: "alert", children: j(D.field) }),
                      i.jsxs("p", {
                        children: [
                          D.selected.length,
                          " ",
                          i.jsx(LocalizedText, {
                            children:
                              "selected items (saved selections may include values no longer in the source).",
                          }),
                        ],
                      }),
                      i.jsxs("button", {
                        onClick: () =>
                          A(() =>
                            z(
                              "slicers",
                              o.slicers.map((_, C) =>
                                B === C
                                  ? {
                                      ..._,
                                      selected: getPivotSlicerItems(
                                        e.workbook,
                                        t,
                                        _.field,
                                      ),
                                    }
                                  : _,
                              ),
                            ),
                          ),
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Select all slicer ",
                          }),
                          B + 1,
                        ],
                      }),
                      i.jsxs("button", {
                        onClick: () =>
                          z(
                            "slicers",
                            o.slicers.map((_, C) =>
                              B === C ? { ..._, selected: [] } : _,
                            ),
                          ),
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Select none slicer ",
                          }),
                          B + 1,
                        ],
                      }),
                      i.jsxs("button", {
                        onClick: () =>
                          z(
                            "slicers",
                            o.slicers.filter((_, C) => B !== C),
                          ),
                        children: [
                          i.jsx(LocalizedText, { children: "Remove slicer " }),
                          B + 1,
                        ],
                      }),
                    ],
                  },
                  B,
                ),
              ),
              i.jsx("button", {
                disabled: !E.length,
                onClick: () =>
                  A(() => {
                    const D = E.find(
                      (B) => !o.slicers?.some((C) => C.field === B),
                    );
                    if (!D)
                      throw Error("Every source field already has a slicer.");
                    z("slicers", [
                      ...(o.slicers ?? []),
                      {
                        field: D,
                        selected: getPivotSlicerItems(e.workbook, t, D),
                      },
                    ]);
                  }),
                children: i.jsx(LocalizedText, { children: "Add slicer" }),
              }),
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "Selected items within one slicer are combined; different controls intersect. Select none produces zero records. Remove a slicer to include future new items.",
                }),
              }),
            ],
          }),
          i.jsxs("details", {
            children: [
              i.jsx("summary", {
                children: i.jsx(LocalizedText, { children: "Timelines" }),
              }),
              (o.timelines ?? []).map((D, B) => {
                const _ = (T) =>
                  z(
                    "timelines",
                    o.timelines.map((N, R) => (B === R ? { ...N, ...T } : N)),
                  );
                let C = [];
                try {
                  C = buildTimelinePeriods(O(D.field), D);
                } catch {}
                return i.jsxs(
                  "fieldset",
                  {
                    children: [
                      i.jsxs("legend", {
                        children: [
                          i.jsx(LocalizedText, { children: "Timeline " }),
                          B + 1,
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Timeline field " }),
                          B + 1,
                          i.jsxs("select", {
                            value: D.field,
                            onChange: (T) => _({ field: T.target.value }),
                            children: [
                              !E.includes(D.field) &&
                                i.jsx("option", { children: D.field }),
                              E.map((T) => i.jsx("option", { children: T }, T)),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Timeline date system ",
                          }),
                          B + 1,
                          i.jsxs("select", {
                            value: D.dateSystem,
                            onChange: (T) => _({ dateSystem: T.target.value }),
                            children: [
                              i.jsx("option", {
                                value: "iso",
                                children: "ISO date text",
                              }),
                              i.jsx("option", {
                                value: "1900",
                                children: "Excel 1900 serial",
                              }),
                              i.jsx("option", {
                                value: "1904",
                                children: "Excel 1904 serial",
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Timeline period ",
                          }),
                          B + 1,
                          i.jsxs("select", {
                            value: D.unit,
                            onChange: (T) => _({ unit: T.target.value }),
                            children: [
                              i.jsx("option", {
                                value: "year",
                                children: "Year",
                              }),
                              i.jsx("option", {
                                value: "quarter",
                                children: "Quarter",
                              }),
                              i.jsx("option", {
                                value: "month",
                                children: "Month",
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsx("div", {
                        className: "lg-pivot-periods",
                        children: C.slice(0, 120).map((T) =>
                          i.jsx(
                            "button",
                            {
                              type: "button",
                              "aria-pressed":
                                D.from === T.from && D.to === T.to,
                              onClick: () => _({ from: T.from, to: T.to }),
                              children: T.label,
                            },
                            T.label,
                          ),
                        ),
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Timeline from " }),
                          B + 1,
                          i.jsx("input", {
                            type: "date",
                            value: D.from,
                            onChange: (T) => _({ from: T.target.value }),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Timeline through ",
                          }),
                          B + 1,
                          i.jsx("input", {
                            type: "date",
                            value: D.to,
                            onChange: (T) => _({ to: T.target.value }),
                          }),
                        ],
                      }),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Inclusive dates; blanks are excluded. Use date bounds if no period buttons are available (maximum 500 distinct dates / 120 buttons).",
                        }),
                      }),
                      i.jsxs("button", {
                        onClick: () =>
                          z(
                            "timelines",
                            o.timelines.filter((T, N) => B !== N),
                          ),
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Remove timeline ",
                          }),
                          B + 1,
                        ],
                      }),
                    ],
                  },
                  B,
                );
              }),
              i.jsx("button", {
                disabled: !E.length,
                onClick: () => {
                  const D =
                    E.find(
                      (B) =>
                        /date/i.test(B) &&
                        !o.timelines?.some((C) => C.field === B),
                    ) ??
                    E.find((B) => !o.timelines?.some((C) => C.field === B));
                  D &&
                    z("timelines", [
                      ...(o.timelines ?? []),
                      {
                        field: D,
                        from: "2000-01-01",
                        to: "2099-12-31",
                        dateSystem: "iso",
                        unit: "month",
                      },
                    ]);
                },
                children: i.jsx(LocalizedText, { children: "Add timeline" }),
              }),
            ],
          }),
          i.jsxs("details", {
            children: [
              i.jsx("summary", {
                children: i.jsx(LocalizedText, {
                  children: "Sort pivot fields",
                }),
              }),
              [...new Set([...r.rowFields, ...r.columnFields])].map((D) =>
                i.jsxs(
                  "label",
                  {
                    children: [
                      i.jsx(LocalizedText, { children: "Sort " }),
                      D,
                      i.jsxs("select", {
                        value:
                          o.sorts?.find((C) => C.field === D)?.direction ?? "",
                        onChange: (C) =>
                          z("sorts", [
                            ...(o.sorts ?? []).filter((T) => T.field !== D),
                            ...(C.target.value
                              ? [{ field: D, direction: C.target.value }]
                              : []),
                          ]),
                        children: [
                          i.jsx("option", {
                            value: "",
                            children: "Source order",
                          }),
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
                  },
                  D,
                ),
              ),
            ],
          }),
          i.jsxs("details", {
            children: [
              i.jsx("summary", {
                children: i.jsx(LocalizedText, { children: "Show values as" }),
              }),
              o.values.map((D, B) =>
                i.jsxs(
                  "label",
                  {
                    children: [
                      D.name ?? `${D.operation} of ${D.field}`,
                      i.jsx("select", {
                        "aria-label": `Show values as ${B + 1}`,
                        value: D.showAs ?? "value",
                        onChange: (_) =>
                          z(
                            "values",
                            o.values.map((C, T) =>
                              T === B ? { ...C, showAs: _.target.value } : C,
                            ),
                          ),
                        children: Object.entries(pivotShowAsLabels).map(
                          ([_, C]) =>
                            i.jsx("option", { value: _, children: C }, _),
                        ),
                      }),
                    ],
                  },
                  B,
                ),
              ),
            ],
          }),
          i.jsxs("details", {
            children: [
              i.jsx("summary", {
                children: i.jsx(LocalizedText, {
                  children: "PivotChart settings",
                }),
              }),
              i.jsxs("label", {
                className: "lg-table-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: !!o.chart,
                    onChange: (D) =>
                      z(
                        "chart",
                        D.target.checked
                          ? { kind: "column", valueIndex: 0 }
                          : undefined,
                      ),
                  }),
                  i.jsx(LocalizedText, { children: "Show PivotChart" }),
                ],
              }),
              o.chart &&
                i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "PivotChart type" }),
                        i.jsxs("select", {
                          value: o.chart.kind,
                          onChange: (D) =>
                            z("chart", { ...o.chart, kind: D.target.value }),
                          children: [
                            i.jsx("option", {
                              value: "column",
                              children: "Column",
                            }),
                            i.jsx("option", { value: "bar", children: "Bar" }),
                            i.jsx("option", {
                              value: "line",
                              children: "Line",
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "PivotChart value" }),
                        i.jsx("select", {
                          value: o.chart.valueIndex,
                          onChange: (D) =>
                            z("chart", {
                              ...o.chart,
                              valueIndex: Number(D.target.value),
                            }),
                          children: o.values.map((D, B) =>
                            i.jsx(
                              "option",
                              {
                                value: B,
                                children:
                                  D.name ?? `${D.operation} of ${D.field}`,
                              },
                              B,
                            ),
                          ),
                        }),
                      ],
                    }),
                  ],
                }),
            ],
          }),
          i.jsx("button", {
            onClick: () =>
              A(() =>
                p({
                  book: applyPivotViewChange(e.workbook, t, o, {
                    snapshot: g,
                    readOnly: e.readOnly,
                  }),
                  snapshot: g,
                  request: b,
                }),
              ),
            children: i.jsx(LocalizedText, { children: "Preview pivot view" }),
          }),
          f &&
            !y &&
            i.jsx("p", {
              role: "status",
              children: i.jsx(LocalizedText, {
                children: "The workbook or controls changed. Preview again.",
              }),
            }),
          V &&
            i.jsxs("div", {
              role: "status",
              children: [
                i.jsx(LocalizedText, { children: "Preview: " }),
                V.refresh.includedRows,
                " ",
                i.jsx(LocalizedText, { children: "included +" }),
                " ",
                V.refresh.excludedRows,
                " ",
                i.jsx(LocalizedText, { children: "excluded =" }),
                " ",
                V.refresh.sourceRows,
                " ",
                i.jsx(LocalizedText, { children: "records." }),
                " ",
                V.refresh.measures
                  .map((D) => `${D.name}: ${D.result}`)
                  .join("; "),
                ".",
              ],
            }),
          i.jsx("button", {
            disabled: !y,
            onClick: () =>
              A(() => {
                if (!y) throw Error("Preview again before applying.");
                e.onWorkbookChange(
                  applyPivotViewChange(e.workbook, t, o, {
                    snapshot: f.snapshot,
                    readOnly: e.readOnly,
                  }),
                  "Apply pivot filters and view",
                );
              }),
            children: i.jsx(LocalizedText, {
              children: "Apply view to new result",
            }),
          }),
        ],
      }),
      (V ? V.chart : r.chart) &&
        i.jsx(PivotChartFigure, {
          book: V ? f.book : e.workbook,
          id: V?.id ?? t,
        }),
      i.jsxs("fieldset", {
        className: "lg-table-list",
        children: [
          i.jsx("legend", {
            children: i.jsx(LocalizedText, { children: "Drill into a result" }),
          }),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Pivot value cell" }),
              i.jsx("input", {
                value: u,
                onChange: (D) => l(D.target.value.toUpperCase()),
              }),
            ],
          }),
          i.jsx("button", {
            onClick: () =>
              A(() => {
                (v({
                  report: buildPivotDrillDown(e.workbook, t, u),
                  snapshot: g,
                }),
                  h(0));
              }),
            children: i.jsx(LocalizedText, { children: "Show source rows" }),
          }),
          m &&
            !w &&
            i.jsx("p", {
              role: "status",
              children: i.jsx(LocalizedText, {
                children:
                  "The workbook or selected value changed. Show source rows again.",
              }),
            }),
          w &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs("p", {
                  role: "status",
                  children: [
                    m.report.records.length,
                    " ",
                    i.jsx(LocalizedText, {
                      children: "underlying records. Aggregate:",
                    }),
                    " ",
                    String(m.report.underlyingValue),
                    i.jsx(LocalizedText, { children: "; displayed value:" }),
                    " ",
                    String(m.report.displayedValue),
                    i.jsx(LocalizedText, { children: ". Source revision" }),
                    " ",
                    m.report.context.sourceRevision,
                    ".",
                  ],
                }),
                i.jsx("div", {
                  className: "lg-pivot-records",
                  children: i.jsxs("table", {
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Source locator",
                              }),
                            }),
                            m.report.fields.map((D) =>
                              i.jsx("th", { children: D }, D),
                            ),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: m.report.records
                          .slice(c * 20, c * 20 + 20)
                          .map((D) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("td", {
                                    children: [
                                      D.locator.sheetName,
                                      "!",
                                      D.locator.range,
                                    ],
                                  }),
                                  D.values.map((B, _) =>
                                    i.jsx(
                                      "td",
                                      { children: String(B ?? "") },
                                      _,
                                    ),
                                  ),
                                ],
                              },
                              D.locator.row,
                            ),
                          ),
                      }),
                    ],
                  }),
                }),
                !m.report.records.length &&
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children: "No source rows underlie this cell.",
                    }),
                  }),
                i.jsx("button", {
                  disabled: !c,
                  onClick: () => h((D) => D - 1),
                  children: i.jsx(LocalizedText, {
                    children: "Previous source rows",
                  }),
                }),
                i.jsx("button", {
                  disabled: (c + 1) * 20 >= m.report.records.length,
                  onClick: () => h((D) => D + 1),
                  children: i.jsx(LocalizedText, {
                    children: "Next source rows",
                  }),
                }),
                i.jsx("button", {
                  onClick: () => k(m.report, "Pivot-drill-through.json"),
                  children: i.jsx(LocalizedText, {
                    children: "Export drilled rows and context",
                  }),
                }),
                i.jsx("button", {
                  disabled: e.readOnly,
                  onClick: () =>
                    A(() =>
                      e.onWorkbookChange(
                        extractPivotDrillThrough(e.workbook, t, u, {
                          snapshot: m.snapshot,
                          readOnly: e.readOnly,
                        }),
                        "Extract pivot source rows",
                      ),
                    ),
                  children: i.jsx(LocalizedText, {
                    children: "Extract source rows to new sheet",
                  }),
                }),
              ],
            }),
        ],
      }),
    ],
  });
}
const pivotSeriesPalette = [
    "#4472C4",
    "#ED7D31",
    "#548235",
    "#8064A2",
    "#9E480E",
    "#008387",
  ],
  formatPivotNumber = (e) =>
    Math.abs(e) >= 1e7 || (Math.abs(e) > 0 && Math.abs(e) < 0.001)
      ? e.toExponential(2)
      : Number(e.toPrecision(5)).toLocaleString("en-US");
function buildAxisTicks(e, r = false, t = true) {
  const n = e.length ? e : [r ? 1 : 0];
  let o = Math.min(...n),
    a = Math.max(...n);
  if (r)
    ((o = Math.log10(o)),
      (a = Math.log10(a)),
      o === a &&
        (o < -322 ? (a += 1) : a > 307 ? (o -= 1) : ((o -= 0.5), (a += 0.5))));
  else if ((t && ((o = Math.min(0, o)), (a = Math.max(0, a))), o === a)) {
    const s = Math.max(1, Math.abs(o) * 0.1);
    ((o = Number.isFinite(o - s) ? o - s : o),
      (a = Number.isFinite(a + s) ? a + s : a));
  }
  return {
    min: r ? 10 ** o : o,
    max: r ? 10 ** a : a,
    fraction: (s) => ((r ? Math.log10(s) : s) - o) / (a - o),
    ticks: Array.from({ length: 5 }, (s, d) =>
      r ? 10 ** (o + ((a - o) * d) / 4) : o + ((a - o) * d) / 4,
    ),
  };
}
function ChartPreviewFrame({ object, sheet, calc }) {
  let n;
  try {
    n = inspectChartSource(sheet, object, calc);
  } catch (o) {
    return i.jsxs("div", {
      role: "img",
      "aria-label": `${object.label}: chart unavailable`,
      style: {
        background: "white",
        padding: 8,
        font: "12px Arial",
        overflow: "auto",
        height: "100%",
      },
      children: [
        i.jsx(LocalizedText, { children: "Chart unavailable: " }),
        String(o instanceof Error ? o.message : o),
      ],
    });
  }
  return i.jsx(ChartReportSvg, { report: n });
}
function ChartReportSvg({ report }) {
  const r = report.diagnostics.filter((B) => B.severity === "error"),
    t = report.kind === "sparkline",
    n = t ? 200 : 640,
    o = t ? 40 : 380;
  if (r.length)
    return i.jsxs("div", {
      role: "img",
      "aria-label": `${report.title}: chart unavailable`,
      style: {
        background: "white",
        padding: 8,
        font: "12px Arial",
        height: "100%",
        overflow: "auto",
      },
      children: [
        i.jsxs("strong", {
          children: [
            report.title,
            i.jsx(LocalizedText, { children: ": chart unavailable" }),
          ],
        }),
        i.jsx("p", {
          children: r
            .slice(0, 5)
            .map((B) => B.message)
            .join(" "),
        }),
        i.jsx("p", { children: report.scope }),
      ],
    });
  const a = report.points.filter((B) => B.included),
    s = report.series
      .map((B, _) => ({ ...B, index: _ }))
      .filter((B) => B.included),
    d = s.flatMap((B) =>
      a.map((_) => _.values[B.index].value).filter((_) => typeof _ == "number"),
    ),
    u = t ? 3 : 80,
    l = t ? 197 : 610,
    c = t ? 3 : 60,
    h = t ? 37 : 270,
    f = l - u,
    p = h - c,
    m =
      report.kind === "histogram"
        ? report.histogram.map((B) => B.count)
        : report.kind === "waterfall"
          ? report.waterfall.flatMap((B) => [B.start, B.end])
          : d,
    v = buildAxisTicks(m, report.axes.y.scale === "log"),
    g = (B) => h - v.fraction(B) * p,
    b = report.axes.y.scale === "log" ? v.min : 0,
    y = buildAxisTicks(
      a.map((B) => Number(B.category.value)).filter(Number.isFinite),
      report.axes.x.scale === "log",
      false,
    ),
    w = (B, _ = a.length) => u + ((B + 0.5) * f) / Math.max(1, _),
    E = (B) => u + y.fraction(B) * f,
    A = [],
    k = (B, _) =>
      `${B}: ${formatPivotNumber(_)}${report.kind === "histogram" ? " records" : report.axes.units ? " " + report.axes.units : ""}`,
    F = (B, _, C, T, N = false) =>
      B.forEach((R, L) => {
        if (R === null) return;
        const $ = (N ? p : f) / Math.max(1, B.length),
          U = ($ * 0.72) / T,
          G = (N ? c : u) + L * $ + $ * 0.14 + C * U,
          H = v.fraction(b),
          Q = v.fraction(R);
        A.push(
          i.jsx(
            "rect",
            {
              "data-chart-mark": "bar",
              x: N ? u + Math.min(H, Q) * f : G,
              y: N ? G : Math.min(g(b), g(R)),
              width: N ? Math.abs(Q - H) * f : U,
              height: N ? U : Math.abs(g(R) - g(b)),
              fill: pivotSeriesPalette[C % pivotSeriesPalette.length],
              children: i.jsx("title", { children: k(_[L], R) }),
            },
            `bar-${C}-${L}`,
          ),
        );
      }),
    I = (B, _, C = false) => {
      let T = [];
      const N = () => {
        if (!T.length) return;
        const R = T[0],
          L = T.at(-1);
        (A.push(
          C
            ? i.jsx(
                "polygon",
                {
                  "data-chart-mark": "area",
                  points: [
                    `${R.x},${g(b)}`,
                    ...T.map(($) => `${$.x},${$.y}`),
                    `${L.x},${g(b)}`,
                  ].join(" "),
                  fill: pivotSeriesPalette[_ % pivotSeriesPalette.length],
                  fillOpacity: ".25",
                  stroke: pivotSeriesPalette[_ % pivotSeriesPalette.length],
                },
                `area-${_}-${R.i}`,
              )
            : i.jsx(
                "polyline",
                {
                  "data-chart-mark": "line",
                  points: T.map(($) => `${$.x},${$.y}`).join(" "),
                  fill: "none",
                  stroke: pivotSeriesPalette[_ % pivotSeriesPalette.length],
                  strokeWidth: t ? 2 : 2.5,
                },
                `line-${_}-${R.i}`,
              ),
        ),
          (T = []));
      };
      (B.forEach((R, L) => {
        if (R === null) {
          N();
          return;
        }
        const $ = w(L),
          U = g(R);
        (T.push({ x: $, y: U, v: R, i: L }),
          A.push(
            i.jsx(
              "circle",
              {
                cx: $,
                cy: U,
                r: t ? 1.7 : 3,
                fill: pivotSeriesPalette[_ % pivotSeriesPalette.length],
                children: i.jsx("title", { children: k(a[L].label, R) }),
              },
              `point-${_}-${L}`,
            ),
          ));
      }),
        N());
    };
  if (report.kind === "pie" || report.kind === "donut") {
    const B = a.map((T) => Number(T.values[s[0].index].value ?? 0)),
      _ = B.reduce((T, N) => T + N, 0);
    let C = -Math.PI / 2;
    (B.forEach((T, N) => {
      if (!T) return;
      const R = C;
      C += (T / _) * Math.PI * 2;
      const L = 315 + Math.cos(R) * 100,
        $ = 166 + Math.sin(R) * 100,
        U = 315 + Math.cos(C) * 100,
        G = 166 + Math.sin(C) * 100;
      A.push(
        T === _
          ? i.jsx(
              "circle",
              {
                "data-chart-mark": "slice",
                cx: "315",
                cy: "166",
                r: "100",
                fill: pivotSeriesPalette[N % pivotSeriesPalette.length],
                children: i.jsx("title", { children: k(a[N].label, T) }),
              },
              N,
            )
          : i.jsx(
              "path",
              {
                "data-chart-mark": "slice",
                d: `M315 166 L${L} ${$} A100 100 0 ${C - R > Math.PI ? 1 : 0} 1 ${U} ${G} Z`,
                fill: pivotSeriesPalette[N % pivotSeriesPalette.length],
                stroke: "white",
                children: i.jsx("title", { children: k(a[N].label, T) }),
              },
              N,
            ),
      );
    }),
      report.kind === "donut" &&
        A.push(
          i.jsx(
            "circle",
            { cx: "315", cy: "166", r: "55", fill: "white" },
            "hole",
          ),
        ));
  } else
    report.kind === "histogram"
      ? F(
          report.histogram.map((B) => B.count),
          report.histogram.map(
            (B) =>
              `${formatPivotNumber(B.lower)}–${formatPivotNumber(B.upper)}`,
          ),
          0,
          1,
        )
      : report.kind === "waterfall"
        ? report.waterfall.forEach((B, _) => {
            const C = w(_, report.waterfall.length),
              T = f / report.waterfall.length;
            (A.push(
              i.jsx(
                "rect",
                {
                  "data-chart-mark": "waterfall",
                  x: C - T * 0.34,
                  y: Math.min(g(B.start), g(B.end)),
                  width: T * 0.68,
                  height: Math.max(1, Math.abs(g(B.start) - g(B.end))),
                  fill:
                    B.row === null
                      ? "#4472C4"
                      : B.value >= 0
                        ? "#548235"
                        : "#C00000",
                  children: i.jsx("title", {
                    children: `${B.label}: change ${formatPivotNumber(B.value)}, from ${formatPivotNumber(B.start)} to ${formatPivotNumber(B.end)}`,
                  }),
                },
                _,
              ),
            ),
              _ < report.waterfall.length - 1 &&
                A.push(
                  i.jsx(
                    "line",
                    {
                      x1: C + T * 0.34,
                      x2: C + T * 0.66,
                      y1: g(B.end),
                      y2: g(B.end),
                      stroke: "#777",
                      strokeDasharray: "3 2",
                    },
                    `connector-${_}`,
                  ),
                ));
          })
        : s.forEach((B, _) => {
            const C = a.map((N) =>
                typeof N.values[B.index].value == "number"
                  ? N.values[B.index].value
                  : null,
              ),
              T =
                report.kind === "combo"
                  ? B.kind
                  : report.kind === "sparkline"
                    ? (report.options.sparklineKind ?? "line")
                    : report.kind;
            T === "scatter"
              ? C.forEach((N, R) => {
                  N !== null &&
                    A.push(
                      i.jsx(
                        "circle",
                        {
                          "data-chart-mark": "scatter",
                          cx: E(Number(a[R].category.value)),
                          cy: g(N),
                          r: "4",
                          fill: pivotSeriesPalette[
                            _ % pivotSeriesPalette.length
                          ],
                          children: i.jsx("title", {
                            children: `${a[R].category.address}=${a[R].category.value}; ${k(B.name, N)}`,
                          }),
                        },
                        `${_}-${R}`,
                      ),
                    );
                })
              : T === "line" || T === "area"
                ? I(C, _, T === "area")
                : T === "winLoss"
                  ? C.forEach((N, R) => {
                      if (N === null || N === 0) return;
                      const L = f / a.length;
                      A.push(
                        i.jsx(
                          "rect",
                          {
                            "data-chart-mark": "winLoss",
                            x: w(R) - L * 0.32,
                            y: N > 0 ? c : c + p / 2 + 1,
                            width: L * 0.64,
                            height: p / 2 - 1,
                            fill: N > 0 ? "#548235" : "#C00000",
                            children: i.jsx("title", {
                              children: k(a[R].label, N),
                            }),
                          },
                          R,
                        ),
                      );
                    })
                  : F(
                      C,
                      a.map((N) => N.label),
                      _,
                      s.length,
                      T === "bar",
                    );
          });
  const O = ["pie", "donut"].includes(report.kind),
    j = report.kind === "bar",
    z =
      report.kind === "histogram"
        ? report.histogram.map(
            (B) =>
              `${formatPivotNumber(B.lower)}–${formatPivotNumber(B.upper)}`,
          )
        : report.kind === "waterfall"
          ? report.waterfall.map((B) => B.label)
          : a.map((B) => B.label),
    V = O
      ? a
          .filter((B) => Number(B.values[s[0].index].value) > 0)
          .map((B) => B.label)
      : s.map((B) => B.name),
    D = O
      ? a
          .filter((B) => Number(B.values[s[0].index].value) > 0)
          .map(
            (B) => pivotSeriesPalette[a.indexOf(B) % pivotSeriesPalette.length],
          )
      : s.map((B, _) => pivotSeriesPalette[_ % pivotSeriesPalette.length]);
  return i.jsxs("svg", {
    viewBox: `0 0 ${n} ${o}`,
    preserveAspectRatio: t ? "none" : "xMidYMid meet",
    role: "img",
    "aria-label": report.altText,
    "data-chart-kind": report.kind,
    style: {
      width: "100%",
      height: "100%",
      display: "block",
      background: "white",
      fontFamily: "Arial,sans-serif",
    },
    children: [
      i.jsx("title", {
        children: `${report.title}${t ? `. ${report.scope}` : ""}`,
      }),
      i.jsxs("desc", {
        children: [
          report.altText,
          " ",
          report.scope,
          ". ",
          report.axes.units && `Units: ${report.axes.units}.`,
          " ",
          i.jsx(LocalizedText, { children: "X:" }),
          " ",
          report.axes.x.title,
          " (",
          report.axes.x.scale,
          i.jsx(LocalizedText, { children: "); Y: " }),
          report.axes.y.title,
          " (",
          report.axes.y.scale,
          "). ",
          report.population.policy,
          " ",
          V.join("; "),
        ],
      }),
      !t &&
        i.jsxs(i.Fragment, {
          children: [
            i.jsx("text", {
              x: "320",
              y: "24",
              textAnchor: "middle",
              fontSize: "17",
              fill: "#222",
              children: report.title.slice(0, 65),
            }),
            i.jsxs("text", {
              x: "320",
              y: "43",
              textAnchor: "middle",
              fontSize: "10",
              fill: "#444",
              children: [
                report.axes.units ? `Units: ${report.axes.units}` : "",
                report.axes.y.scale === "log"
                  ? " · Logarithmic value axis"
                  : "",
              ],
            }),
          ],
        }),
      !t &&
        !O &&
        i.jsxs(i.Fragment, {
          children: [
            v.ticks.map((B, _) =>
              i.jsxs(
                "g",
                {
                  children: [
                    i.jsx("line", {
                      x1: j ? u + v.fraction(B) * f : u,
                      x2: j ? u + v.fraction(B) * f : l,
                      y1: j ? c : g(B),
                      y2: j ? h : g(B),
                      stroke: "#E2E5E8",
                    }),
                    i.jsx("text", {
                      x: j ? u + v.fraction(B) * f : u - 6,
                      y: j ? h + 15 : g(B) + 3,
                      textAnchor: j ? "middle" : "end",
                      fontSize: "10",
                      fill: "#555",
                      children: formatPivotNumber(B),
                    }),
                  ],
                },
                _,
              ),
            ),
            report.kind === "scatter"
              ? y.ticks.map((B, _) =>
                  i.jsx(
                    "text",
                    {
                      x: E(B),
                      y: h + 15,
                      textAnchor: "middle",
                      fontSize: "10",
                      fill: "#555",
                      children: formatPivotNumber(B),
                    },
                    _,
                  ),
                )
              : z.map((B, _) =>
                  _ % Math.max(1, Math.ceil(z.length / (j ? 10 : 8))) === 0
                    ? i.jsx(
                        "text",
                        {
                          x: j ? u - 7 : w(_, z.length),
                          y: j ? c + ((_ + 0.5) * p) / z.length + 3 : h + 15,
                          textAnchor: j ? "end" : "middle",
                          fontSize: "10",
                          fill: "#555",
                          children: B.length > 14 ? B.slice(0, 12) + "…" : B,
                        },
                        _,
                      )
                    : null,
                ),
            i.jsx("text", {
              x: "340",
              y: "305",
              textAnchor: "middle",
              fontSize: "11",
              children: j ? report.axes.y.title : report.axes.x.title,
            }),
            i.jsx("text", {
              transform: "translate(14 168) rotate(-90)",
              textAnchor: "middle",
              fontSize: "11",
              children: j ? report.axes.x.title : report.axes.y.title,
            }),
          ],
        }),
      A,
      !t &&
        i.jsxs(i.Fragment, {
          children: [
            report.options.showLegend !== false &&
              V.slice(0, 6).map((B, _) =>
                i.jsxs(
                  "g",
                  {
                    children: [
                      i.jsx("rect", {
                        x: 25 + (_ % 3) * 205,
                        y: 317 + Math.floor(_ / 3) * 17,
                        width: "10",
                        height: "10",
                        fill: D[_],
                      }),
                      i.jsx("text", {
                        x: 39 + (_ % 3) * 205,
                        y: 326 + Math.floor(_ / 3) * 17,
                        fontSize: "10",
                        fill: "#333",
                        children: B.length > 24 ? B.slice(0, 22) + "…" : B,
                      }),
                    ],
                  },
                  _,
                ),
              ),
            i.jsxs("text", {
              x: "320",
              y: "361",
              textAnchor: "middle",
              fontSize: "10",
              fill: "#444",
              children: [
                report.source.sheetName,
                "!",
                report.source.range,
                " · ",
                report.population.includedRows,
                "/",
                report.population.sourceRows,
                " ",
                i.jsx(LocalizedText, { children: "rows · " }),
                report.population.selectedSeries,
                "/",
                report.population.availableSeries,
                " ",
                i.jsx(LocalizedText, { children: "series" }),
              ],
            }),
            i.jsxs("text", {
              x: "320",
              y: "375",
              textAnchor: "middle",
              fontSize: "9",
              fill: "#555",
              children: [
                report.population.hiddenRows,
                " ",
                i.jsx(LocalizedText, { children: "hidden source rows ·" }),
                " ",
                report.diagnostics.length
                  ? `${report.diagnostics.length} diagnostics · `
                  : "",
                V.length > 6 ? "Full legend and " : "",
                i.jsx(LocalizedText, {
                  children:
                    "Full source, filters and series in Chart properties / Export inspection",
                }),
              ],
            }),
          ],
        }),
    ],
  });
}
function downloadTextFile(e) {
  const r = URL.createObjectURL(new Blob([e.content], { type: e.mime })),
    t = document.createElement("a");
  ((t.href = r),
    (t.download = e.filename),
    t.click(),
    setTimeout(() => URL.revokeObjectURL(r), 1e3));
}
function ChartSettingsPanel({ object, report, update }) {
  const { text, language } = useWorkbookMessages(),
    a = object.chartOptions ?? {},
    s = (l) => update({ chartOptions: { ...a, ...l } }),
    d = report?.series ?? [],
    u = () =>
      d
        .filter((l) => l.included)
        .map((l) => ({ column: l.column, name: l.name, kind: l.kind }));
  return i.jsxs(i.Fragment, {
    children: [
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: "Chart type" }),
          i.jsx("select", {
            "aria-label": text("Chart type"),
            value: object.chart,
            onChange: (l) => {
              const c = l.target.value,
                h = [
                  "pie",
                  "donut",
                  "waterfall",
                  "histogram",
                  "sparkline",
                ].includes(c);
              update({
                chart: c,
                chartOptions: {
                  ...a,
                  yScale: "linear",
                  xScale: "linear",
                  ...(c === "histogram"
                    ? {
                        xTitle:
                          d.find((p) => p.included)?.name ?? "Source value",
                        yTitle: "Frequency (records)",
                      }
                    : {}),
                  ...(h && d.length
                    ? { series: [u()[0] ?? { column: d[0].column }] }
                    : {}),
                },
              });
            },
            lang: language("Chart type"),
            "data-ui-label": "Chart type",
            children: supportedChartKinds.map((l) =>
              i.jsx(
                "option",
                { value: l, children: l[0].toUpperCase() + l.slice(1) },
                l,
              ),
            ),
          }),
        ],
      }),
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: "Source range" }),
          i.jsx("input", {
            "aria-label": text("Chart source range"),
            value: object.range,
            onChange: (l) =>
              update({
                range: l.target.value.toUpperCase(),
                chartOptions: {
                  ...a,
                  series: undefined,
                  includedRows: undefined,
                },
              }),
            lang: language("Chart source range"),
            "data-ui-label": "Chart source range",
          }),
        ],
      }),
      i.jsxs("label", {
        className: "dg-check",
        children: [
          i.jsx("input", {
            type: "checkbox",
            checked: object.headers,
            onChange: (l) =>
              update({
                headers: l.target.checked,
                chartOptions: {
                  ...a,
                  series: undefined,
                  includedRows: undefined,
                },
              }),
          }),
          i.jsx(LocalizedText, { children: "First row contains series names" }),
        ],
      }),
      i.jsx("p", {
        className: "dg-muted",
        children: i.jsx(LocalizedText, {
          children:
            "The first column is the category, or numeric X for scatter. Remaining columns are value series; a single-column range uses cell addresses as labels. Up to 500 cells and 20 value series. Blank values are gaps. Every source row remains inspectable, including hidden rows.",
        }),
      }),
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: "Chart anchor cell" }),
          i.jsx("input", {
            "aria-label": text("Chart anchor cell"),
            value: object.anchor,
            onChange: (l) => update({ anchor: l.target.value.toUpperCase() }),
            lang: language("Chart anchor cell"),
            "data-ui-label": "Chart anchor cell",
          }),
        ],
      }),
      object.chart === "sparkline" &&
        i.jsxs(i.Fragment, {
          children: [
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "In-cell sparkline fits the anchor cell as its size changes. Use a separate empty destination cell. Inspect or export through Chart properties.",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Sparkline style" }),
                i.jsxs("select", {
                  "aria-label": text("Sparkline style"),
                  value: a.sparklineKind ?? "line",
                  onChange: (l) => s({ sparklineKind: l.target.value }),
                  lang: language("Sparkline style"),
                  "data-ui-label": "Sparkline style",
                  children: [
                    i.jsx("option", { value: "line", children: "Line" }),
                    i.jsx("option", { value: "column", children: "Column" }),
                    i.jsx("option", { value: "winLoss", children: "Win/loss" }),
                  ],
                }),
              ],
            }),
          ],
        }),
      i.jsxs("details", {
        children: [
          i.jsx("summary", {
            children: i.jsx(LocalizedText, {
              children: "Titles, axes and alt text",
            }),
          }),
          i.jsxs("div", {
            className: "dg-chart-stack",
            children: [
              ["altText", "xTitle", "yTitle", "units"].map((l, c) =>
                i.jsxs(
                  "label",
                  {
                    children: [
                      [
                        "Alt text",
                        "Category / X axis title",
                        "Value / Y axis title",
                        "Units (label; values are not rescaled)",
                      ][c],
                      i.jsx("input", {
                        "aria-label": [
                          "Chart alt text",
                          "Chart X axis title",
                          "Chart Y axis title",
                          "Chart units",
                        ][c],
                        maxLength: l === "altText" ? 2e3 : 200,
                        value: a[l] ?? "",
                        onChange: (h) => s({ [l]: h.target.value }),
                      }),
                    ],
                  },
                  l,
                ),
              ),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Value axis scale" }),
                  i.jsxs("select", {
                    "aria-label": text("Chart Y scale"),
                    value: a.yScale ?? "linear",
                    onChange: (l) => s({ yScale: l.target.value }),
                    lang: language("Chart Y scale"),
                    "data-ui-label": "Chart Y scale",
                    children: [
                      i.jsx("option", { value: "linear", children: "Linear" }),
                      i.jsx("option", {
                        value: "log",
                        children: "Logarithmic (positive values)",
                      }),
                    ],
                  }),
                ],
              }),
              object.chart === "scatter" &&
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "X axis scale" }),
                    i.jsxs("select", {
                      "aria-label": text("Chart X scale"),
                      value: a.xScale ?? "linear",
                      onChange: (l) => s({ xScale: l.target.value }),
                      lang: language("Chart X scale"),
                      "data-ui-label": "Chart X scale",
                      children: [
                        i.jsx("option", {
                          value: "linear",
                          children: "Linear",
                        }),
                        i.jsx("option", {
                          value: "log",
                          children: "Logarithmic (positive values)",
                        }),
                      ],
                    }),
                  ],
                }),
              i.jsxs("label", {
                className: "dg-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: a.showLegend !== false,
                    onChange: (l) => s({ showLegend: l.target.checked }),
                  }),
                  i.jsx(LocalizedText, { children: "Show legend" }),
                ],
              }),
            ],
          }),
        ],
      }),
      object.chart === "histogram" &&
        i.jsxs("label", {
          children: [
            i.jsx(LocalizedText, { children: "Histogram bins" }),
            i.jsx(LocalizedNumberField, {
              "aria-label": text("Histogram bins"),
              min: 1,
              max: 50,
              value:
                a.histogramBins ??
                Math.min(
                  50,
                  Math.max(1, Math.ceil(Math.sqrt(report?.points.length ?? 1))),
                ),
              onValueChange: (l) => s({ histogramBins: Number(l) }),
              lang: language("Histogram bins"),
              "data-ui-label": "Histogram bins",
            }),
          ],
        }),
      i.jsxs("details", {
        children: [
          i.jsx("summary", {
            children: i.jsx(LocalizedText, { children: "Chart series" }),
          }),
          i.jsx("button", {
            type: "button",
            onClick: () => s({ series: undefined }),
            children: i.jsx(LocalizedText, {
              children: "Use all source series",
            }),
          }),
          i.jsx("div", {
            className: "dg-chart-stack",
            children: d.map((l) =>
              i.jsxs(
                "div",
                {
                  className: "dg-chart-series",
                  children: [
                    i.jsxs("label", {
                      className: "dg-check",
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          "aria-label": `Include series ${l.column}`,
                          checked: l.included,
                          onChange: (c) =>
                            s({
                              series: c.target.checked
                                ? [
                                    ...u(),
                                    {
                                      column: l.column,
                                      name: l.name,
                                      kind: l.kind,
                                    },
                                  ]
                                : u().filter((h) => h.column !== l.column),
                            }),
                        }),
                        l.name,
                        " ",
                        i.jsx(LocalizedText, { children: "(column " }),
                        l.column,
                        ")",
                      ],
                    }),
                    l.included &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, { children: "Series name" }),
                              i.jsx("input", {
                                "aria-label": `Series ${l.column} name`,
                                value: l.name,
                                maxLength: 200,
                                onChange: (c) =>
                                  s({
                                    series: u().map((h) =>
                                      h.column === l.column
                                        ? {
                                            ...h,
                                            name: c.target.value || undefined,
                                          }
                                        : h,
                                    ),
                                  }),
                              }),
                            ],
                          }),
                          object.chart === "combo" &&
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Series rendering",
                                }),
                                i.jsxs("select", {
                                  "aria-label": `Series ${l.column} rendering`,
                                  value: l.kind,
                                  onChange: (c) =>
                                    s({
                                      series: u().map((h) =>
                                        h.column === l.column
                                          ? { ...h, kind: c.target.value }
                                          : h,
                                      ),
                                    }),
                                  children: [
                                    i.jsx("option", {
                                      value: "column",
                                      children: "Column",
                                    }),
                                    i.jsx("option", {
                                      value: "line",
                                      children: "Line",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                        ],
                      }),
                  ],
                },
                l.column,
              ),
            ),
          }),
        ],
      }),
      i.jsxs("details", {
        children: [
          i.jsx("summary", {
            children: i.jsx(LocalizedText, { children: "Chart row filter" }),
          }),
          i.jsx("p", { children: report?.scope }),
          i.jsx("button", {
            type: "button",
            onClick: () => s({ includedRows: undefined }),
            children: i.jsx(LocalizedText, {
              children: "Include all chart rows",
            }),
          }),
          i.jsx("button", {
            type: "button",
            onClick: () => s({ includedRows: [] }),
            children: i.jsx(LocalizedText, {
              children: "Exclude all chart rows",
            }),
          }),
          i.jsx("div", {
            className: "dg-chart-row-list",
            children: report?.points.map((l) =>
              i.jsxs(
                "label",
                {
                  className: "dg-check",
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      "aria-label": `Include chart row ${l.row}`,
                      checked: l.included,
                      onChange: (c) => {
                        const h =
                          a.includedRows ?? report.points.map((f) => f.row);
                        s({
                          includedRows: c.target.checked
                            ? [...h, l.row]
                            : h.filter((f) => f !== l.row),
                        });
                      },
                    }),
                    l.row,
                    ": ",
                    l.label,
                    l.hidden ? " (hidden worksheet row)" : "",
                  ],
                },
                l.row,
              ),
            ),
          }),
        ],
      }),
    ],
  });
}

export {
  ChartReportSvg,
  PivotExplorerDialog,
  downloadTextFile,
  ChartSettingsPanel,
  ChartPreviewFrame,
};
