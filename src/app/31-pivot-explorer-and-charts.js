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

function PivotExplorerDialog(props) {
  const { text, language } = useWorkbookMessages(),
    pivotEntries = props.workbook.sheets.flatMap((sheet) =>
      (sheet.pivots ?? []).map((definition) => ({
        sheet: sheet,
        definition: definition,
      })),
    ),
    [selectedPivotId, setSelectedPivotId] = ne.useState(() => {
      const anchorCell = _t(props.range.split(":")[0]);
      return (
        pivotEntries.find((entry) => {
          const rangeAddress = parseRangeAddress(entry.definition.outputRange);
          return (
            entry.sheet.id === props.sheetId &&
            anchorCell.row >= rangeAddress.firstRow &&
            anchorCell.row <= rangeAddress.lastRow &&
            anchorCell.col >= rangeAddress.firstCol &&
            anchorCell.col <= rangeAddress.lastCol
          );
        })?.definition.id ??
        pivotEntries.find((entry) => entry.sheet.id === props.sheetId)
          ?.definition.id ??
        pivotEntries[0]?.definition.id ??
        ""
      );
    }),
    definition = pivotEntries.find(
      (entry) => entry.definition.id === selectedPivotId,
    )?.definition;
  return i.jsxs("div", {
    role: "dialog",
    "aria-modal": "false",
    "aria-label": text("Explore PivotTable"),
    className: "lg-table-tools lg-pivot-explore",
    onKeyDown: (event) => {
      event.key === "Escape" && props.onClose();
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
            onClick: props.onClose,
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
            value: selectedPivotId,
            onChange: (event) => setSelectedPivotId(event.target.value),
            children: pivotEntries.map((entry) =>
              i.jsxs(
                "option",
                {
                  value: entry.definition.id,
                  children: [entry.sheet.name, ": ", entry.definition.name],
                },
                entry.definition.id,
              ),
            ),
          }),
        ],
      }),
      definition
        ? i.jsx(
            PivotExplorerPanel,
            { ...props, definition: definition },
            selectedPivotId,
          )
        : i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children: "Create a PivotTable to explore its results.",
            }),
          }),
    ],
  });
}
function PivotExplorerPanel(props) {
  const definition = props.definition,
    pivotId = definition.id,
    panelRef = ne.useRef(null),
    [draftView, setDraftView] = ne.useState(() =>
      structuredClone({
        slicers: definition.slicers ?? [],
        timelines: definition.timelines ?? [],
        sorts: definition.sorts ?? [],
        values: definition.values,
        chart: definition.chart,
      }),
    ),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [valueCellAddress, setValueCellAddress] = ne.useState(
      props.range.split(":")[0],
    ),
    [recordPage, setRecordPage] = ne.useState(0),
    [previewState, setPreviewState] = ne.useState(null),
    [drillDownState, setDrillDownState] = ne.useState(() => {
      try {
        return {
          report: buildPivotDrillDown(
            props.workbook,
            pivotId,
            valueCellAddress,
          ),
          snapshot: JSON.stringify(props.workbook),
        };
      } catch {
        return null;
      }
    }),
    snapshot = JSON.stringify(props.workbook),
    request = JSON.stringify(draftView),
    isPreviewCurrent =
      previewState?.snapshot === snapshot && previewState.request === request,
    isDrillDownCurrent =
      drillDownState?.snapshot === snapshot &&
      drillDownState.report.address === valueCellAddress.toUpperCase(),
    sourceFields = ne.useMemo(() => {
      try {
        return getPivotSourceFields(
          props.workbook,
          definition.source.sheetId,
          definition.source.range,
        );
      } catch {
        return [];
      }
    }, [props.workbook, definition.source.sheetId, definition.source.range]);
  ne.useEffect(() => {
    panelRef.current?.querySelector("input")?.focus();
  }, []);
  const runAction = (action) => {
      try {
        (action(), setErrorMessage(""));
      } catch (error) {
        setErrorMessage(error.message);
      }
    },
    exportJson = (payload, fileName) =>
      runAction(() => {
        if (props.onExport) props.onExport(structuredClone(payload), fileName);
        else {
          const objectUrl = URL.createObjectURL(
              new Blob([JSON.stringify(payload, null, 2)], {
                type: "application/json",
              }),
            ),
            downloadLink = document.createElement("a");
          ((downloadLink.href = objectUrl),
            (downloadLink.download = fileName),
            downloadLink.click(),
            setTimeout(() => URL.revokeObjectURL(objectUrl), 1e3));
        }
      }),
    controlFieldsKey = JSON.stringify([
      ...new Set(
        [...(draftView.slicers ?? []), ...(draftView.timelines ?? [])].map(
          (control) => control.field,
        ),
      ),
    ]),
    controlFieldInfo = ne.useMemo(() => {
      try {
        return inspectPivotControlFields(
          props.workbook,
          pivotId,
          JSON.parse(controlFieldsKey),
        );
      } catch (error) {
        return new Map(
          JSON.parse(controlFieldsKey).map((fieldName) => [
            fieldName,
            { items: [], error: error.message },
          ]),
        );
      }
    }, [props.workbook, pivotId, controlFieldsKey]),
    itemsForField = (fieldName) => controlFieldInfo.get(fieldName)?.items ?? [],
    errorForField = (fieldName) => controlFieldInfo.get(fieldName)?.error ?? "",
    updateDraftView = (key, value) =>
      setDraftView((previousView) => ({ ...previousView, [key]: value })),
    previewPivot = isPreviewCurrent
      ? previewState.book.sheets.flatMap((sheet) => sheet.pivots ?? []).at(-1)
      : null;
  return i.jsxs("div", {
    ref: panelRef,
    className: "lg-table-form",
    children: [
      i.jsx(PivotStatusBadge, { workbook: props.workbook, pivotId: pivotId }),
      errorMessage && i.jsx("p", { role: "alert", children: errorMessage }),
      props.readOnly &&
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Read-only: inspect and export existing results; changes and extraction are disabled.",
          }),
        }),
      i.jsx("button", {
        onClick: () =>
          exportJson(
            describePivotRefresh(props.workbook, pivotId),
            "Pivot-filter-context.json",
          ),
        children: i.jsx(LocalizedText, { children: "Export pivot context" }),
      }),
      i.jsxs("fieldset", {
        disabled: props.readOnly,
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
              (draftView.slicers ?? []).map((slicer, slicerIndex) =>
                i.jsxs(
                  "fieldset",
                  {
                    children: [
                      i.jsxs("legend", {
                        children: [
                          i.jsx(LocalizedText, { children: "Slicer " }),
                          slicerIndex + 1,
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Slicer field " }),
                          slicerIndex + 1,
                          i.jsxs("select", {
                            value: slicer.field,
                            onChange: (event) =>
                              runAction(() =>
                                updateDraftView(
                                  "slicers",
                                  draftView.slicers.map((slicer, index) =>
                                    slicerIndex === index
                                      ? {
                                          field: event.target.value,
                                          selected: getPivotSlicerItems(
                                            props.workbook,
                                            pivotId,
                                            event.target.value,
                                          ),
                                        }
                                      : slicer,
                                  ),
                                ),
                              ),
                            children: [
                              !sourceFields.includes(slicer.field) &&
                                i.jsx("option", { children: slicer.field }),
                              sourceFields.map((sourceField) =>
                                i.jsx(
                                  "option",
                                  { children: sourceField },
                                  sourceField,
                                ),
                              ),
                            ],
                          }),
                        ],
                      }),
                      i.jsx("div", {
                        className: "lg-pivot-items",
                        children: itemsForField(slicer.field).map(
                          (slicerItem, itemIndex) =>
                            i.jsxs(
                              "label",
                              {
                                className: "lg-table-check",
                                children: [
                                  i.jsx("input", {
                                    type: "checkbox",
                                    checked: slicer.selected.some(
                                      (selectedItem) =>
                                        buildSlicerItemKey(selectedItem) ===
                                        buildSlicerItemKey(slicerItem),
                                    ),
                                    onChange: (event) =>
                                      updateDraftView(
                                        "slicers",
                                        draftView.slicers.map(
                                          (slicer, index) =>
                                            slicerIndex === index
                                              ? {
                                                  ...slicer,
                                                  selected: event.target.checked
                                                    ? [
                                                        ...slicer.selected,
                                                        slicerItem,
                                                      ]
                                                    : slicer.selected.filter(
                                                        (selectedItem) =>
                                                          buildSlicerItemKey(
                                                            selectedItem,
                                                          ) !==
                                                          buildSlicerItemKey(
                                                            slicerItem,
                                                          ),
                                                      ),
                                                }
                                              : slicer,
                                        ),
                                      ),
                                  }),
                                  formatPivotGroupLabel(slicerItem),
                                ],
                              },
                              itemIndex,
                            ),
                        ),
                      }),
                      errorForField(slicer.field) &&
                        i.jsx("p", {
                          role: "alert",
                          children: errorForField(slicer.field),
                        }),
                      i.jsxs("p", {
                        children: [
                          slicer.selected.length,
                          " ",
                          i.jsx(LocalizedText, {
                            children:
                              "selected items (saved selections may include values no longer in the source).",
                          }),
                        ],
                      }),
                      i.jsxs("button", {
                        onClick: () =>
                          runAction(() =>
                            updateDraftView(
                              "slicers",
                              draftView.slicers.map((slicer, index) =>
                                slicerIndex === index
                                  ? {
                                      ...slicer,
                                      selected: getPivotSlicerItems(
                                        props.workbook,
                                        pivotId,
                                        slicer.field,
                                      ),
                                    }
                                  : slicer,
                              ),
                            ),
                          ),
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Select all slicer ",
                          }),
                          slicerIndex + 1,
                        ],
                      }),
                      i.jsxs("button", {
                        onClick: () =>
                          updateDraftView(
                            "slicers",
                            draftView.slicers.map((slicer, index) =>
                              slicerIndex === index
                                ? { ...slicer, selected: [] }
                                : slicer,
                            ),
                          ),
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Select none slicer ",
                          }),
                          slicerIndex + 1,
                        ],
                      }),
                      i.jsxs("button", {
                        onClick: () =>
                          updateDraftView(
                            "slicers",
                            draftView.slicers.filter(
                              (slicer, index) => slicerIndex !== index,
                            ),
                          ),
                        children: [
                          i.jsx(LocalizedText, { children: "Remove slicer " }),
                          slicerIndex + 1,
                        ],
                      }),
                    ],
                  },
                  slicerIndex,
                ),
              ),
              i.jsx("button", {
                disabled: !sourceFields.length,
                onClick: () =>
                  runAction(() => {
                    const field = sourceFields.find(
                      (sourceField) =>
                        !draftView.slicers?.some(
                          (slicer) => slicer.field === sourceField,
                        ),
                    );
                    if (!field)
                      throw Error("Every source field already has a slicer.");
                    updateDraftView("slicers", [
                      ...(draftView.slicers ?? []),
                      {
                        field: field,
                        selected: getPivotSlicerItems(
                          props.workbook,
                          pivotId,
                          field,
                        ),
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
              (draftView.timelines ?? []).map((timelineItem, timelineIndex) => {
                const updateTimeline = (patch) =>
                  updateDraftView(
                    "timelines",
                    draftView.timelines.map((timeline, index) =>
                      timelineIndex === index
                        ? { ...timeline, ...patch }
                        : timeline,
                    ),
                  );
                let periods = [];
                try {
                  periods = buildTimelinePeriods(
                    itemsForField(timelineItem.field),
                    timelineItem,
                  );
                } catch {}
                return i.jsxs(
                  "fieldset",
                  {
                    children: [
                      i.jsxs("legend", {
                        children: [
                          i.jsx(LocalizedText, { children: "Timeline " }),
                          timelineIndex + 1,
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Timeline field " }),
                          timelineIndex + 1,
                          i.jsxs("select", {
                            value: timelineItem.field,
                            onChange: (event) =>
                              updateTimeline({ field: event.target.value }),
                            children: [
                              !sourceFields.includes(timelineItem.field) &&
                                i.jsx("option", {
                                  children: timelineItem.field,
                                }),
                              sourceFields.map((sourceField) =>
                                i.jsx(
                                  "option",
                                  { children: sourceField },
                                  sourceField,
                                ),
                              ),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Timeline date system ",
                          }),
                          timelineIndex + 1,
                          i.jsxs("select", {
                            value: timelineItem.dateSystem,
                            onChange: (event) =>
                              updateTimeline({
                                dateSystem: event.target.value,
                              }),
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
                          timelineIndex + 1,
                          i.jsxs("select", {
                            value: timelineItem.unit,
                            onChange: (event) =>
                              updateTimeline({ unit: event.target.value }),
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
                        children: periods.slice(0, 120).map((period) =>
                          i.jsx(
                            "button",
                            {
                              type: "button",
                              "aria-pressed":
                                timelineItem.from === period.from &&
                                timelineItem.to === period.to,
                              onClick: () =>
                                updateTimeline({
                                  from: period.from,
                                  to: period.to,
                                }),
                              children: period.label,
                            },
                            period.label,
                          ),
                        ),
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Timeline from " }),
                          timelineIndex + 1,
                          i.jsx("input", {
                            type: "date",
                            value: timelineItem.from,
                            onChange: (event) =>
                              updateTimeline({ from: event.target.value }),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Timeline through ",
                          }),
                          timelineIndex + 1,
                          i.jsx("input", {
                            type: "date",
                            value: timelineItem.to,
                            onChange: (event) =>
                              updateTimeline({ to: event.target.value }),
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
                          updateDraftView(
                            "timelines",
                            draftView.timelines.filter(
                              (timeline, index) => timelineIndex !== index,
                            ),
                          ),
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Remove timeline ",
                          }),
                          timelineIndex + 1,
                        ],
                      }),
                    ],
                  },
                  timelineIndex,
                );
              }),
              i.jsx("button", {
                disabled: !sourceFields.length,
                onClick: () => {
                  const field =
                    sourceFields.find(
                      (sourceField) =>
                        /date/i.test(sourceField) &&
                        !draftView.timelines?.some(
                          (timeline) => timeline.field === sourceField,
                        ),
                    ) ??
                    sourceFields.find(
                      (sourceField) =>
                        !draftView.timelines?.some(
                          (timeline) => timeline.field === sourceField,
                        ),
                    );
                  field &&
                    updateDraftView("timelines", [
                      ...(draftView.timelines ?? []),
                      {
                        field: field,
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
              [
                ...new Set([
                  ...definition.rowFields,
                  ...definition.columnFields,
                ]),
              ].map((field) =>
                i.jsxs(
                  "label",
                  {
                    children: [
                      i.jsx(LocalizedText, { children: "Sort " }),
                      field,
                      i.jsxs("select", {
                        value:
                          draftView.sorts?.find((sort) => sort.field === field)
                            ?.direction ?? "",
                        onChange: (event) =>
                          updateDraftView("sorts", [
                            ...(draftView.sorts ?? []).filter(
                              (sort) => sort.field !== field,
                            ),
                            ...(event.target.value
                              ? [
                                  {
                                    field: field,
                                    direction: event.target.value,
                                  },
                                ]
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
                  field,
                ),
              ),
            ],
          }),
          i.jsxs("details", {
            children: [
              i.jsx("summary", {
                children: i.jsx(LocalizedText, { children: "Show values as" }),
              }),
              draftView.values.map((value, valueIndex) =>
                i.jsxs(
                  "label",
                  {
                    children: [
                      value.name ?? `${value.operation} of ${value.field}`,
                      i.jsx("select", {
                        "aria-label": `Show values as ${valueIndex + 1}`,
                        value: value.showAs ?? "value",
                        onChange: (event) =>
                          updateDraftView(
                            "values",
                            draftView.values.map((value, index) =>
                              index === valueIndex
                                ? { ...value, showAs: event.target.value }
                                : value,
                            ),
                          ),
                        children: Object.entries(pivotShowAsLabels).map(
                          ([showAsKey, showAsLabel]) =>
                            i.jsx(
                              "option",
                              { value: showAsKey, children: showAsLabel },
                              showAsKey,
                            ),
                        ),
                      }),
                    ],
                  },
                  valueIndex,
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
                    checked: !!draftView.chart,
                    onChange: (event) =>
                      updateDraftView(
                        "chart",
                        event.target.checked
                          ? { kind: "column", valueIndex: 0 }
                          : undefined,
                      ),
                  }),
                  i.jsx(LocalizedText, { children: "Show PivotChart" }),
                ],
              }),
              draftView.chart &&
                i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "PivotChart type" }),
                        i.jsxs("select", {
                          value: draftView.chart.kind,
                          onChange: (event) =>
                            updateDraftView("chart", {
                              ...draftView.chart,
                              kind: event.target.value,
                            }),
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
                          value: draftView.chart.valueIndex,
                          onChange: (event) =>
                            updateDraftView("chart", {
                              ...draftView.chart,
                              valueIndex: Number(event.target.value),
                            }),
                          children: draftView.values.map((value, valueIndex) =>
                            i.jsx(
                              "option",
                              {
                                value: valueIndex,
                                children:
                                  value.name ??
                                  `${value.operation} of ${value.field}`,
                              },
                              valueIndex,
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
              runAction(() =>
                setPreviewState({
                  book: applyPivotViewChange(
                    props.workbook,
                    pivotId,
                    draftView,
                    {
                      snapshot: snapshot,
                      readOnly: props.readOnly,
                    },
                  ),
                  snapshot: snapshot,
                  request: request,
                }),
              ),
            children: i.jsx(LocalizedText, { children: "Preview pivot view" }),
          }),
          previewState &&
            !isPreviewCurrent &&
            i.jsx("p", {
              role: "status",
              children: i.jsx(LocalizedText, {
                children: "The workbook or controls changed. Preview again.",
              }),
            }),
          previewPivot &&
            i.jsxs("div", {
              role: "status",
              children: [
                i.jsx(LocalizedText, { children: "Preview: " }),
                previewPivot.refresh.includedRows,
                " ",
                i.jsx(LocalizedText, { children: "included +" }),
                " ",
                previewPivot.refresh.excludedRows,
                " ",
                i.jsx(LocalizedText, { children: "excluded =" }),
                " ",
                previewPivot.refresh.sourceRows,
                " ",
                i.jsx(LocalizedText, { children: "records." }),
                " ",
                previewPivot.refresh.measures
                  .map((measure) => `${measure.name}: ${measure.result}`)
                  .join("; "),
                ".",
              ],
            }),
          i.jsx("button", {
            disabled: !isPreviewCurrent,
            onClick: () =>
              runAction(() => {
                if (!isPreviewCurrent)
                  throw Error("Preview again before applying.");
                props.onWorkbookChange(
                  applyPivotViewChange(props.workbook, pivotId, draftView, {
                    snapshot: previewState.snapshot,
                    readOnly: props.readOnly,
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
      (previewPivot ? previewPivot.chart : definition.chart) &&
        i.jsx(PivotChartFigure, {
          book: previewPivot ? previewState.book : props.workbook,
          id: previewPivot?.id ?? pivotId,
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
                value: valueCellAddress,
                onChange: (event) =>
                  setValueCellAddress(event.target.value.toUpperCase()),
              }),
            ],
          }),
          i.jsx("button", {
            onClick: () =>
              runAction(() => {
                (setDrillDownState({
                  report: buildPivotDrillDown(
                    props.workbook,
                    pivotId,
                    valueCellAddress,
                  ),
                  snapshot: snapshot,
                }),
                  setRecordPage(0));
              }),
            children: i.jsx(LocalizedText, { children: "Show source rows" }),
          }),
          drillDownState &&
            !isDrillDownCurrent &&
            i.jsx("p", {
              role: "status",
              children: i.jsx(LocalizedText, {
                children:
                  "The workbook or selected value changed. Show source rows again.",
              }),
            }),
          isDrillDownCurrent &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs("p", {
                  role: "status",
                  children: [
                    drillDownState.report.records.length,
                    " ",
                    i.jsx(LocalizedText, {
                      children: "underlying records. Aggregate:",
                    }),
                    " ",
                    String(drillDownState.report.underlyingValue),
                    i.jsx(LocalizedText, { children: "; displayed value:" }),
                    " ",
                    String(drillDownState.report.displayedValue),
                    i.jsx(LocalizedText, { children: ". Source revision" }),
                    " ",
                    drillDownState.report.context.sourceRevision,
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
                            drillDownState.report.fields.map((field) =>
                              i.jsx("th", { children: field }, field),
                            ),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: drillDownState.report.records
                          .slice(recordPage * 20, recordPage * 20 + 20)
                          .map((record) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("td", {
                                    children: [
                                      record.locator.sheetName,
                                      "!",
                                      record.locator.range,
                                    ],
                                  }),
                                  record.values.map((value, columnIndex) =>
                                    i.jsx(
                                      "td",
                                      { children: String(value ?? "") },
                                      columnIndex,
                                    ),
                                  ),
                                ],
                              },
                              record.locator.row,
                            ),
                          ),
                      }),
                    ],
                  }),
                }),
                !drillDownState.report.records.length &&
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children: "No source rows underlie this cell.",
                    }),
                  }),
                i.jsx("button", {
                  disabled: !recordPage,
                  onClick: () => setRecordPage((page) => page - 1),
                  children: i.jsx(LocalizedText, {
                    children: "Previous source rows",
                  }),
                }),
                i.jsx("button", {
                  disabled:
                    (recordPage + 1) * 20 >=
                    drillDownState.report.records.length,
                  onClick: () => setRecordPage((page) => page + 1),
                  children: i.jsx(LocalizedText, {
                    children: "Next source rows",
                  }),
                }),
                i.jsx("button", {
                  onClick: () =>
                    exportJson(
                      drillDownState.report,
                      "Pivot-drill-through.json",
                    ),
                  children: i.jsx(LocalizedText, {
                    children: "Export drilled rows and context",
                  }),
                }),
                i.jsx("button", {
                  disabled: props.readOnly,
                  onClick: () =>
                    runAction(() =>
                      props.onWorkbookChange(
                        extractPivotDrillThrough(
                          props.workbook,
                          pivotId,
                          valueCellAddress,
                          {
                            snapshot: drillDownState.snapshot,
                            readOnly: props.readOnly,
                          },
                        ),
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
  formatPivotNumber = (numericValue) =>
    Math.abs(numericValue) >= 1e7 ||
    (Math.abs(numericValue) > 0 && Math.abs(numericValue) < 0.001)
      ? numericValue.toExponential(2)
      : Number(numericValue.toPrecision(5)).toLocaleString("en-US");
function buildAxisTicks(values, isLogScale = false, shouldIncludeZero = true) {
  const valuesOrDefault = values.length ? values : [isLogScale ? 1 : 0];
  let axisMin = Math.min(...valuesOrDefault),
    axisMax = Math.max(...valuesOrDefault);
  if (isLogScale)
    ((axisMin = Math.log10(axisMin)),
      (axisMax = Math.log10(axisMax)),
      axisMin === axisMax &&
        (axisMin < -322
          ? (axisMax += 1)
          : axisMax > 307
            ? (axisMin -= 1)
            : ((axisMin -= 0.5), (axisMax += 0.5))));
  else if (
    (shouldIncludeZero &&
      ((axisMin = Math.min(0, axisMin)), (axisMax = Math.max(0, axisMax))),
    axisMin === axisMax)
  ) {
    const padding = Math.max(1, Math.abs(axisMin) * 0.1);
    ((axisMin = Number.isFinite(axisMin - padding)
      ? axisMin - padding
      : axisMin),
      (axisMax = Number.isFinite(axisMax + padding)
        ? axisMax + padding
        : axisMax));
  }
  return {
    min: isLogScale ? 10 ** axisMin : axisMin,
    max: isLogScale ? 10 ** axisMax : axisMax,
    fraction: (value) =>
      ((isLogScale ? Math.log10(value) : value) - axisMin) /
      (axisMax - axisMin),
    ticks: Array.from({ length: 5 }, (s, tickIndex) =>
      isLogScale
        ? 10 ** (axisMin + ((axisMax - axisMin) * tickIndex) / 4)
        : axisMin + ((axisMax - axisMin) * tickIndex) / 4,
    ),
  };
}
function ChartPreviewFrame({ object, sheet, calc }) {
  let report;
  try {
    report = inspectChartSource(sheet, object, calc);
  } catch (error) {
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
        String(error instanceof Error ? error.message : error),
      ],
    });
  }
  return i.jsx(ChartReportSvg, { report: report });
}
function ChartReportSvg({ report }) {
  const errorDiagnostics = report.diagnostics.filter(
      (diagnostic) => diagnostic.severity === "error",
    ),
    isSparkline = report.kind === "sparkline",
    viewWidth = isSparkline ? 200 : 640,
    viewHeight = isSparkline ? 40 : 380;
  if (errorDiagnostics.length)
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
          children: errorDiagnostics
            .slice(0, 5)
            .map((diagnostic) => diagnostic.message)
            .join(" "),
        }),
        i.jsx("p", { children: report.scope }),
      ],
    });
  const includedPoints = report.points.filter((point) => point.included),
    includedSeries = report.series
      .map((sery, seriesIndex) => ({ ...sery, index: seriesIndex }))
      .filter((series) => series.included),
    numericValues = includedSeries.flatMap((series) =>
      includedPoints
        .map((point) => point.values[series.index].value)
        .filter((value) => typeof value == "number"),
    ),
    plotLeft = isSparkline ? 3 : 80,
    plotRight = isSparkline ? 197 : 610,
    plotTop = isSparkline ? 3 : 60,
    plotBottom = isSparkline ? 37 : 270,
    plotWidth = plotRight - plotLeft,
    plotHeight = plotBottom - plotTop,
    axisValues =
      report.kind === "histogram"
        ? report.histogram.map((bin) => bin.count)
        : report.kind === "waterfall"
          ? report.waterfall.flatMap((step) => [step.start, step.end])
          : numericValues,
    axisTicks = buildAxisTicks(axisValues, report.axes.y.scale === "log"),
    valueToY = (value) => plotBottom - axisTicks.fraction(value) * plotHeight,
    baselineValue = report.axes.y.scale === "log" ? axisTicks.min : 0,
    xAxisTicks = buildAxisTicks(
      includedPoints
        .map((point) => Number(point.category.value))
        .filter(Number.isFinite),
      report.axes.x.scale === "log",
      false,
    ),
    pointCenterX = (pointIndex, pointCount = includedPoints.length) =>
      plotLeft + ((pointIndex + 0.5) * plotWidth) / Math.max(1, pointCount),
    scatterX = (xValue) => plotLeft + xAxisTicks.fraction(xValue) * plotWidth,
    marks = [],
    formatTooltip = (label, value) =>
      `${label}: ${formatPivotNumber(value)}${report.kind === "histogram" ? " records" : report.axes.units ? " " + report.axes.units : ""}`,
    drawBars = (
      values,
      labels,
      seriesIndex,
      seriesCount,
      isHorizontal = false,
    ) =>
      values.forEach((barValue, valueIndex) => {
        if (barValue === null) return;
        const slotSize =
            (isHorizontal ? plotHeight : plotWidth) /
            Math.max(1, values.length),
          barThickness = (slotSize * 0.72) / seriesCount,
          barStart =
            (isHorizontal ? plotTop : plotLeft) +
            valueIndex * slotSize +
            slotSize * 0.14 +
            seriesIndex * barThickness,
          baselineFraction = axisTicks.fraction(baselineValue),
          valueFraction = axisTicks.fraction(barValue);
        marks.push(
          i.jsx(
            "rect",
            {
              "data-chart-mark": "bar",
              x: isHorizontal
                ? plotLeft +
                  Math.min(baselineFraction, valueFraction) * plotWidth
                : barStart,
              y: isHorizontal
                ? barStart
                : Math.min(valueToY(baselineValue), valueToY(barValue)),
              width: isHorizontal
                ? Math.abs(valueFraction - baselineFraction) * plotWidth
                : barThickness,
              height: isHorizontal
                ? barThickness
                : Math.abs(valueToY(barValue) - valueToY(baselineValue)),
              fill: pivotSeriesPalette[seriesIndex % pivotSeriesPalette.length],
              children: i.jsx("title", {
                children: formatTooltip(labels[valueIndex], barValue),
              }),
            },
            `bar-${seriesIndex}-${valueIndex}`,
          ),
        );
      }),
    drawLineSeries = (values, seriesIndex, isArea = false) => {
      let segmentPoints = [];
      const flushSegment = () => {
        if (!segmentPoints.length) return;
        const firstPoint = segmentPoints[0],
          lastPoint = segmentPoints.at(-1);
        (marks.push(
          isArea
            ? i.jsx(
                "polygon",
                {
                  "data-chart-mark": "area",
                  points: [
                    `${firstPoint.x},${valueToY(baselineValue)}`,
                    ...segmentPoints.map(
                      (segmentPoint) => `${segmentPoint.x},${segmentPoint.y}`,
                    ),
                    `${lastPoint.x},${valueToY(baselineValue)}`,
                  ].join(" "),
                  fill: pivotSeriesPalette[
                    seriesIndex % pivotSeriesPalette.length
                  ],
                  fillOpacity: ".25",
                  stroke:
                    pivotSeriesPalette[seriesIndex % pivotSeriesPalette.length],
                },
                `area-${seriesIndex}-${firstPoint.i}`,
              )
            : i.jsx(
                "polyline",
                {
                  "data-chart-mark": "line",
                  points: segmentPoints
                    .map(
                      (segmentPoint) => `${segmentPoint.x},${segmentPoint.y}`,
                    )
                    .join(" "),
                  fill: "none",
                  stroke:
                    pivotSeriesPalette[seriesIndex % pivotSeriesPalette.length],
                  strokeWidth: isSparkline ? 2 : 2.5,
                },
                `line-${seriesIndex}-${firstPoint.i}`,
              ),
        ),
          (segmentPoints = []));
      };
      (values.forEach((value, valueIndex) => {
        if (value === null) {
          flushSegment();
          return;
        }
        const pointX = pointCenterX(valueIndex),
          pointY = valueToY(value);
        (segmentPoints.push({ x: pointX, y: pointY, v: value, i: valueIndex }),
          marks.push(
            i.jsx(
              "circle",
              {
                cx: pointX,
                cy: pointY,
                r: isSparkline ? 1.7 : 3,
                fill: pivotSeriesPalette[
                  seriesIndex % pivotSeriesPalette.length
                ],
                children: i.jsx("title", {
                  children: formatTooltip(
                    includedPoints[valueIndex].label,
                    value,
                  ),
                }),
              },
              `point-${seriesIndex}-${valueIndex}`,
            ),
          ));
      }),
        flushSegment());
    };
  if (report.kind === "pie" || report.kind === "donut") {
    const sliceValues = includedPoints.map((point) =>
        Number(point.values[includedSeries[0].index].value ?? 0),
      ),
      totalValue = sliceValues.reduce(
        (total, sliceValue) => total + sliceValue,
        0,
      );
    let runningAngle = -Math.PI / 2;
    (sliceValues.forEach((sliceValue, sliceIndex) => {
      if (!sliceValue) return;
      const sliceStartAngle = runningAngle;
      runningAngle += (sliceValue / totalValue) * Math.PI * 2;
      const startX = 315 + Math.cos(sliceStartAngle) * 100,
        startY = 166 + Math.sin(sliceStartAngle) * 100,
        endX = 315 + Math.cos(runningAngle) * 100,
        endY = 166 + Math.sin(runningAngle) * 100;
      marks.push(
        sliceValue === totalValue
          ? i.jsx(
              "circle",
              {
                "data-chart-mark": "slice",
                cx: "315",
                cy: "166",
                r: "100",
                fill: pivotSeriesPalette[
                  sliceIndex % pivotSeriesPalette.length
                ],
                children: i.jsx("title", {
                  children: formatTooltip(
                    includedPoints[sliceIndex].label,
                    sliceValue,
                  ),
                }),
              },
              sliceIndex,
            )
          : i.jsx(
              "path",
              {
                "data-chart-mark": "slice",
                d: `M315 166 L${startX} ${startY} A100 100 0 ${runningAngle - sliceStartAngle > Math.PI ? 1 : 0} 1 ${endX} ${endY} Z`,
                fill: pivotSeriesPalette[
                  sliceIndex % pivotSeriesPalette.length
                ],
                stroke: "white",
                children: i.jsx("title", {
                  children: formatTooltip(
                    includedPoints[sliceIndex].label,
                    sliceValue,
                  ),
                }),
              },
              sliceIndex,
            ),
      );
    }),
      report.kind === "donut" &&
        marks.push(
          i.jsx(
            "circle",
            { cx: "315", cy: "166", r: "55", fill: "white" },
            "hole",
          ),
        ));
  } else
    report.kind === "histogram"
      ? drawBars(
          report.histogram.map((bin) => bin.count),
          report.histogram.map(
            (bin) =>
              `${formatPivotNumber(bin.lower)}–${formatPivotNumber(bin.upper)}`,
          ),
          0,
          1,
        )
      : report.kind === "waterfall"
        ? report.waterfall.forEach((waterfallStep, stepIndex) => {
            const stepCenterX = pointCenterX(
                stepIndex,
                report.waterfall.length,
              ),
              stepWidth = plotWidth / report.waterfall.length;
            (marks.push(
              i.jsx(
                "rect",
                {
                  "data-chart-mark": "waterfall",
                  x: stepCenterX - stepWidth * 0.34,
                  y: Math.min(
                    valueToY(waterfallStep.start),
                    valueToY(waterfallStep.end),
                  ),
                  width: stepWidth * 0.68,
                  height: Math.max(
                    1,
                    Math.abs(
                      valueToY(waterfallStep.start) -
                        valueToY(waterfallStep.end),
                    ),
                  ),
                  fill:
                    waterfallStep.row === null
                      ? "#4472C4"
                      : waterfallStep.value >= 0
                        ? "#548235"
                        : "#C00000",
                  children: i.jsx("title", {
                    children: `${waterfallStep.label}: change ${formatPivotNumber(waterfallStep.value)}, from ${formatPivotNumber(waterfallStep.start)} to ${formatPivotNumber(waterfallStep.end)}`,
                  }),
                },
                stepIndex,
              ),
            ),
              stepIndex < report.waterfall.length - 1 &&
                marks.push(
                  i.jsx(
                    "line",
                    {
                      x1: stepCenterX + stepWidth * 0.34,
                      x2: stepCenterX + stepWidth * 0.66,
                      y1: valueToY(waterfallStep.end),
                      y2: valueToY(waterfallStep.end),
                      stroke: "#777",
                      strokeDasharray: "3 2",
                    },
                    `connector-${stepIndex}`,
                  ),
                ));
          })
        : includedSeries.forEach((series, seriesIndex) => {
            const seriesValues = includedPoints.map((point) =>
                typeof point.values[series.index].value == "number"
                  ? point.values[series.index].value
                  : null,
              ),
              markKind =
                report.kind === "combo"
                  ? series.kind
                  : report.kind === "sparkline"
                    ? (report.options.sparklineKind ?? "line")
                    : report.kind;
            markKind === "scatter"
              ? seriesValues.forEach((scatterValue, pointIndex) => {
                  scatterValue !== null &&
                    marks.push(
                      i.jsx(
                        "circle",
                        {
                          "data-chart-mark": "scatter",
                          cx: scatterX(
                            Number(includedPoints[pointIndex].category.value),
                          ),
                          cy: valueToY(scatterValue),
                          r: "4",
                          fill: pivotSeriesPalette[
                            seriesIndex % pivotSeriesPalette.length
                          ],
                          children: i.jsx("title", {
                            children: `${includedPoints[pointIndex].category.address}=${includedPoints[pointIndex].category.value}; ${formatTooltip(series.name, scatterValue)}`,
                          }),
                        },
                        `${seriesIndex}-${pointIndex}`,
                      ),
                    );
                })
              : markKind === "line" || markKind === "area"
                ? drawLineSeries(seriesValues, seriesIndex, markKind === "area")
                : markKind === "winLoss"
                  ? seriesValues.forEach((winLossValue, pointIndex) => {
                      if (winLossValue === null || winLossValue === 0) return;
                      const slotWidth = plotWidth / includedPoints.length;
                      marks.push(
                        i.jsx(
                          "rect",
                          {
                            "data-chart-mark": "winLoss",
                            x: pointCenterX(pointIndex) - slotWidth * 0.32,
                            y:
                              winLossValue > 0
                                ? plotTop
                                : plotTop + plotHeight / 2 + 1,
                            width: slotWidth * 0.64,
                            height: plotHeight / 2 - 1,
                            fill: winLossValue > 0 ? "#548235" : "#C00000",
                            children: i.jsx("title", {
                              children: formatTooltip(
                                includedPoints[pointIndex].label,
                                winLossValue,
                              ),
                            }),
                          },
                          pointIndex,
                        ),
                      );
                    })
                  : drawBars(
                      seriesValues,
                      includedPoints.map((point) => point.label),
                      seriesIndex,
                      includedSeries.length,
                      markKind === "bar",
                    );
          });
  const isPieOrDonut = ["pie", "donut"].includes(report.kind),
    isBarChart = report.kind === "bar",
    categoryLabels =
      report.kind === "histogram"
        ? report.histogram.map(
            (bin) =>
              `${formatPivotNumber(bin.lower)}–${formatPivotNumber(bin.upper)}`,
          )
        : report.kind === "waterfall"
          ? report.waterfall.map((step) => step.label)
          : includedPoints.map((point) => point.label),
    legendLabels = isPieOrDonut
      ? includedPoints
          .filter(
            (point) => Number(point.values[includedSeries[0].index].value) > 0,
          )
          .map((point) => point.label)
      : includedSeries.map((series) => series.name),
    legendColors = isPieOrDonut
      ? includedPoints
          .filter(
            (point) => Number(point.values[includedSeries[0].index].value) > 0,
          )
          .map(
            (point) =>
              pivotSeriesPalette[
                includedPoints.indexOf(point) % pivotSeriesPalette.length
              ],
          )
      : includedSeries.map(
          (series, seriesIndex) =>
            pivotSeriesPalette[seriesIndex % pivotSeriesPalette.length],
        );
  return i.jsxs("svg", {
    viewBox: `0 0 ${viewWidth} ${viewHeight}`,
    preserveAspectRatio: isSparkline ? "none" : "xMidYMid meet",
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
        children: `${report.title}${isSparkline ? `. ${report.scope}` : ""}`,
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
          legendLabels.join("; "),
        ],
      }),
      !isSparkline &&
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
      !isSparkline &&
        !isPieOrDonut &&
        i.jsxs(i.Fragment, {
          children: [
            axisTicks.ticks.map((tick, tickIndex) =>
              i.jsxs(
                "g",
                {
                  children: [
                    i.jsx("line", {
                      x1: isBarChart
                        ? plotLeft + axisTicks.fraction(tick) * plotWidth
                        : plotLeft,
                      x2: isBarChart
                        ? plotLeft + axisTicks.fraction(tick) * plotWidth
                        : plotRight,
                      y1: isBarChart ? plotTop : valueToY(tick),
                      y2: isBarChart ? plotBottom : valueToY(tick),
                      stroke: "#E2E5E8",
                    }),
                    i.jsx("text", {
                      x: isBarChart
                        ? plotLeft + axisTicks.fraction(tick) * plotWidth
                        : plotLeft - 6,
                      y: isBarChart ? plotBottom + 15 : valueToY(tick) + 3,
                      textAnchor: isBarChart ? "middle" : "end",
                      fontSize: "10",
                      fill: "#555",
                      children: formatPivotNumber(tick),
                    }),
                  ],
                },
                tickIndex,
              ),
            ),
            report.kind === "scatter"
              ? xAxisTicks.ticks.map((tick, tickIndex) =>
                  i.jsx(
                    "text",
                    {
                      x: scatterX(tick),
                      y: plotBottom + 15,
                      textAnchor: "middle",
                      fontSize: "10",
                      fill: "#555",
                      children: formatPivotNumber(tick),
                    },
                    tickIndex,
                  ),
                )
              : categoryLabels.map((categoryLabel, labelIndex) =>
                  labelIndex %
                    Math.max(
                      1,
                      Math.ceil(categoryLabels.length / (isBarChart ? 10 : 8)),
                    ) ===
                  0
                    ? i.jsx(
                        "text",
                        {
                          x: isBarChart
                            ? plotLeft - 7
                            : pointCenterX(labelIndex, categoryLabels.length),
                          y: isBarChart
                            ? plotTop +
                              ((labelIndex + 0.5) * plotHeight) /
                                categoryLabels.length +
                              3
                            : plotBottom + 15,
                          textAnchor: isBarChart ? "end" : "middle",
                          fontSize: "10",
                          fill: "#555",
                          children:
                            categoryLabel.length > 14
                              ? categoryLabel.slice(0, 12) + "…"
                              : categoryLabel,
                        },
                        labelIndex,
                      )
                    : null,
                ),
            i.jsx("text", {
              x: "340",
              y: "305",
              textAnchor: "middle",
              fontSize: "11",
              children: isBarChart ? report.axes.y.title : report.axes.x.title,
            }),
            i.jsx("text", {
              transform: "translate(14 168) rotate(-90)",
              textAnchor: "middle",
              fontSize: "11",
              children: isBarChart ? report.axes.x.title : report.axes.y.title,
            }),
          ],
        }),
      marks,
      !isSparkline &&
        i.jsxs(i.Fragment, {
          children: [
            report.options.showLegend !== false &&
              legendLabels.slice(0, 6).map((legendLabel, legendIndex) =>
                i.jsxs(
                  "g",
                  {
                    children: [
                      i.jsx("rect", {
                        x: 25 + (legendIndex % 3) * 205,
                        y: 317 + Math.floor(legendIndex / 3) * 17,
                        width: "10",
                        height: "10",
                        fill: legendColors[legendIndex],
                      }),
                      i.jsx("text", {
                        x: 39 + (legendIndex % 3) * 205,
                        y: 326 + Math.floor(legendIndex / 3) * 17,
                        fontSize: "10",
                        fill: "#333",
                        children:
                          legendLabel.length > 24
                            ? legendLabel.slice(0, 22) + "…"
                            : legendLabel,
                      }),
                    ],
                  },
                  legendIndex,
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
                legendLabels.length > 6 ? "Full legend and " : "",
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
function downloadTextFile(file) {
  const objectUrl = URL.createObjectURL(
      new Blob([file.content], { type: file.mime }),
    ),
    link = document.createElement("a");
  ((link.href = objectUrl),
    (link.download = file.filename),
    link.click(),
    setTimeout(() => URL.revokeObjectURL(objectUrl), 1e3));
}
function ChartSettingsPanel({ object, report, update }) {
  const { text, language } = useWorkbookMessages(),
    chartOptions = object.chartOptions ?? {},
    mergeChartOptions = (optionsPatch) =>
      update({ chartOptions: { ...chartOptions, ...optionsPatch } }),
    reportSeries = report?.series ?? [],
    includedSeries = () =>
      reportSeries
        .filter((seriesItem) => seriesItem.included)
        .map((seriesConfig) => ({
          column: seriesConfig.column,
          name: seriesConfig.name,
          kind: seriesConfig.kind,
        }));
  return i.jsxs(i.Fragment, {
    children: [
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: "Chart type" }),
          i.jsx("select", {
            "aria-label": text("Chart type"),
            value: object.chart,
            onChange: (event) => {
              const value = event.target.value,
                isSingleSeriesChart = [
                  "pie",
                  "donut",
                  "waterfall",
                  "histogram",
                  "sparkline",
                ].includes(value);
              update({
                chart: value,
                chartOptions: {
                  ...chartOptions,
                  yScale: "linear",
                  xScale: "linear",
                  ...(value === "histogram"
                    ? {
                        xTitle:
                          reportSeries.find(
                            (sourceSeries) => sourceSeries.included,
                          )?.name ?? "Source value",
                        yTitle: "Frequency (records)",
                      }
                    : {}),
                  ...(isSingleSeriesChart && reportSeries.length
                    ? {
                        series: [
                          includedSeries()[0] ?? {
                            column: reportSeries[0].column,
                          },
                        ],
                      }
                    : {}),
                },
              });
            },
            lang: language("Chart type"),
            "data-ui-label": "Chart type",
            children: supportedChartKinds.map((supportedChartKind) =>
              i.jsx(
                "option",
                {
                  value: supportedChartKind,
                  children:
                    supportedChartKind[0].toUpperCase() +
                    supportedChartKind.slice(1),
                },
                supportedChartKind,
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
            onChange: (event) =>
              update({
                range: event.target.value.toUpperCase(),
                chartOptions: {
                  ...chartOptions,
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
            onChange: (event) =>
              update({
                headers: event.target.checked,
                chartOptions: {
                  ...chartOptions,
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
            onChange: (event) =>
              update({ anchor: event.target.value.toUpperCase() }),
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
                  value: chartOptions.sparklineKind ?? "line",
                  onChange: (event) =>
                    mergeChartOptions({ sparklineKind: event.target.value }),
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
              ["altText", "xTitle", "yTitle", "units"].map(
                (optionKey, optionIndex) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        [
                          "Alt text",
                          "Category / X axis title",
                          "Value / Y axis title",
                          "Units (label; values are not rescaled)",
                        ][optionIndex],
                        i.jsx("input", {
                          "aria-label": [
                            "Chart alt text",
                            "Chart X axis title",
                            "Chart Y axis title",
                            "Chart units",
                          ][optionIndex],
                          maxLength: optionKey === "altText" ? 2e3 : 200,
                          value: chartOptions[optionKey] ?? "",
                          onChange: (event) =>
                            mergeChartOptions({
                              [optionKey]: event.target.value,
                            }),
                        }),
                      ],
                    },
                    optionKey,
                  ),
              ),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Value axis scale" }),
                  i.jsxs("select", {
                    "aria-label": text("Chart Y scale"),
                    value: chartOptions.yScale ?? "linear",
                    onChange: (event) =>
                      mergeChartOptions({ yScale: event.target.value }),
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
                      value: chartOptions.xScale ?? "linear",
                      onChange: (event) =>
                        mergeChartOptions({ xScale: event.target.value }),
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
                    checked: chartOptions.showLegend !== false,
                    onChange: (event) =>
                      mergeChartOptions({ showLegend: event.target.checked }),
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
                chartOptions.histogramBins ??
                Math.min(
                  50,
                  Math.max(1, Math.ceil(Math.sqrt(report?.points.length ?? 1))),
                ),
              onValueChange: (newBinCount) =>
                mergeChartOptions({ histogramBins: Number(newBinCount) }),
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
            onClick: () => mergeChartOptions({ series: undefined }),
            children: i.jsx(LocalizedText, {
              children: "Use all source series",
            }),
          }),
          i.jsx("div", {
            className: "dg-chart-stack",
            children: reportSeries.map((seriesEntry) =>
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
                          "aria-label": `Include series ${seriesEntry.column}`,
                          checked: seriesEntry.included,
                          onChange: (event) =>
                            mergeChartOptions({
                              series: event.target.checked
                                ? [
                                    ...includedSeries(),
                                    {
                                      column: seriesEntry.column,
                                      name: seriesEntry.name,
                                      kind: seriesEntry.kind,
                                    },
                                  ]
                                : includedSeries().filter(
                                    (keptSeries) =>
                                      keptSeries.column !== seriesEntry.column,
                                  ),
                            }),
                        }),
                        seriesEntry.name,
                        " ",
                        i.jsx(LocalizedText, { children: "(column " }),
                        seriesEntry.column,
                        ")",
                      ],
                    }),
                    seriesEntry.included &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, { children: "Series name" }),
                              i.jsx("input", {
                                "aria-label": `Series ${seriesEntry.column} name`,
                                value: seriesEntry.name,
                                maxLength: 200,
                                onChange: (event) =>
                                  mergeChartOptions({
                                    series: includedSeries().map(
                                      (includedEntry) =>
                                        includedEntry.column ===
                                        seriesEntry.column
                                          ? {
                                              ...includedEntry,
                                              name:
                                                event.target.value || undefined,
                                            }
                                          : includedEntry,
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
                                  "aria-label": `Series ${seriesEntry.column} rendering`,
                                  value: seriesEntry.kind,
                                  onChange: (event) =>
                                    mergeChartOptions({
                                      series: includedSeries().map(
                                        (styledSeries) =>
                                          styledSeries.column ===
                                          seriesEntry.column
                                            ? {
                                                ...styledSeries,
                                                kind: event.target.value,
                                              }
                                            : styledSeries,
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
                seriesEntry.column,
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
            onClick: () => mergeChartOptions({ includedRows: undefined }),
            children: i.jsx(LocalizedText, {
              children: "Include all chart rows",
            }),
          }),
          i.jsx("button", {
            type: "button",
            onClick: () => mergeChartOptions({ includedRows: [] }),
            children: i.jsx(LocalizedText, {
              children: "Exclude all chart rows",
            }),
          }),
          i.jsx("div", {
            className: "dg-chart-row-list",
            children: report?.points.map((point) =>
              i.jsxs(
                "label",
                {
                  className: "dg-check",
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      "aria-label": `Include chart row ${point.row}`,
                      checked: point.included,
                      onChange: (event) => {
                        const includedRows =
                          chartOptions.includedRows ??
                          report.points.map((point) => point.row);
                        mergeChartOptions({
                          includedRows: event.target.checked
                            ? [...includedRows, point.row]
                            : includedRows.filter(
                                (rowNumber) => rowNumber !== point.row,
                              ),
                        });
                      },
                    }),
                    point.row,
                    ": ",
                    point.label,
                    point.hidden ? " (hidden worksheet row)" : "",
                  ],
                },
                point.row,
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
