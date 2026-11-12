import i from "react/jsx-runtime";
import ne from "react";
import { normalizeDatePolicy } from "./01-schemas-date-functions-and-validation.js";
import {
  buildFilterFacetValues,
  resolveTableDataRange,
  createSavedFilterView,
  parseRangeAddress,
  comparableCellValueKey,
  addSavedFilterView,
  applyFilterConditions,
  removeSavedFilterView,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  createDefaultFilterCriterion,
  filterOperatorLabels,
} from "./18-transformation-recipes-dialog.js";

function FilterViewsDialog(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    previewRef = ne.useRef(null),
    [draftDefinition, setDraftDefinition] = ne.useState(() =>
      props.definition
        ? structuredClone(props.definition)
        : {
            range: resolveTableDataRange(
              props.workbook,
              props.sheetId,
              props.range,
            ),
            columns: [],
          },
    ),
    [selectedColumn, setSelectedColumn] = ne.useState(
      props.column ?? parseRangeAddress(draftDefinition.range).firstCol,
    ),
    [valueSearch, setValueSearch] = ne.useState(""),
    [preview, setPreview] = ne.useState(null),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [selectedViewId, setSelectedViewId] = ne.useState(""),
    [viewName, setViewName] = ne.useState(""),
    [statusMessage, setStatusMessage] = ne.useState(""),
    selectedView = props.workbook.savedFilterViews?.find(
      (savedFilterView) => savedFilterView.id === selectedViewId,
    ),
    canEditFilters =
      !!props.onWorkbookChange &&
      !props.readOnly &&
      !props.workbook.structureProtection?.enabled &&
      !props.workbook.sheets.find((sheet) => sheet.id === props.sheetId)
        ?.protection?.enabled,
    canUpdateSelectedView =
      !!props.onWorkbookChange &&
      !props.readOnly &&
      !props.workbook.structureProtection?.enabled &&
      !!selectedView &&
      !props.workbook.sheets.find((sheet) => sheet.id === selectedView.sheetId)
        ?.protection?.enabled,
    columnOptions = ne.useMemo(() => {
      try {
        const rangeAddress = parseRangeAddress(draftDefinition.range);
        return Array.from(
          {
            length: Math.min(
              64,
              rangeAddress.lastCol - rangeAddress.firstCol + 1,
            ),
          },
          (de, columnOffset) => {
            const column = rangeAddress.firstCol + columnOffset;
            return {
              column: column,
              label: `${columnIndexToLetters(column)} · ${String(props.workbook.sheets.find((sheet) => sheet.id === props.sheetId)?.cells[`${columnIndexToLetters(column)}${rangeAddress.firstRow + 1}`]?.raw ?? "Column")}`,
            };
          },
        );
      } catch {
        return [];
      }
    }, [draftDefinition.range, props.workbook, props.sheetId]),
    facetState = ne.useMemo(() => {
      try {
        return {
          result: buildFilterFacetValues(
            props.workbook,
            props.sheetId,
            draftDefinition,
            selectedColumn,
            valueSearch,
          ),
          error: "",
        };
      } catch (error) {
        return {
          result: undefined,
          error:
            error instanceof Error ? error.message : "Cannot inspect values.",
        };
      }
    }, [
      props.workbook,
      props.sheetId,
      draftDefinition,
      selectedColumn,
      valueSearch,
    ]),
    columnCriterion = draftDefinition.columns.find(
      (column) => column.column === selectedColumn,
    ) ?? { column: selectedColumn },
    currentPreview =
      preview?.book === props.workbook &&
      preview.definition === JSON.stringify(draftDefinition)
        ? preview.result
        : null;
  function getPreviewedFilter() {
    if (!currentPreview)
      throw Error(
        "Preview the current population before saving or applying a view.",
      );
    const filterResult = applyFilterConditions(
      props.workbook,
      props.sheetId,
      draftDefinition,
    );
    if (
      JSON.stringify(filterResult.context) !==
      JSON.stringify(currentPreview.context)
    )
      throw Error(
        "Population membership changed. Preview again before saving or applying.",
      );
    return filterResult;
  }
  function saveFilterView(replaceExisting = false) {
    try {
      if (
        !canEditFilters ||
        (replaceExisting &&
          (!selectedView || selectedView.sheetId !== props.sheetId))
      )
        throw Error("This saved view cannot be changed here.");
      const validatedFilter = getPreviewedFilter(),
        savedFilterView = createSavedFilterView(
          props.workbook,
          props.sheetId,
          validatedFilter.context.definition,
          viewName,
          replaceExisting ? selectedViewId : undefined,
        );
      (props.onWorkbookChange(
        savedFilterView,
        replaceExisting
          ? "Replace saved filter criteria"
          : "Save named filter view",
      ),
        setSelectedViewId(savedFilterView.savedFilterViews.at(-1).id),
        setPreview(null),
        setErrorMessage(""),
        setStatusMessage(
          "Saved in the native workbook. Preview again to apply; saved counts are historical and source cells are unchanged.",
        ));
    } catch (error) {
      setErrorMessage(
        error instanceof Error ? error.message : "Cannot save this view.",
      );
    }
  }
  ne.useEffect(() => {
    var previewElement;
    preview &&
      ((previewElement = previewRef.current) == null
        ? undefined
        : previewElement.scrollIntoView
      )?.call(previewElement, {
        block: "nearest",
      });
  }, [preview]);
  function upsertColumnCriterion(criterion) {
    (setDraftDefinition((previousDraft) => ({
      ...previousDraft,
      columns: [
        ...previousDraft.columns.filter(
          (column) => column.column !== selectedColumn,
        ),
        criterion,
      ],
    })),
      setErrorMessage(""));
  }
  function toggleFilterValue(cellValue) {
    const filterFacetValues = buildFilterFacetValues(
      props.workbook,
      props.sheetId,
      draftDefinition,
      selectedColumn,
    );
    if (columnCriterion.values === undefined && filterFacetValues.truncated) {
      setErrorMessage(
        "Choose Select no values first, then select specific values, or use a custom condition.",
      );
      return;
    }
    const selectedValues =
        columnCriterion.values ??
        filterFacetValues.values.map((value) => value.value),
      valueKey = comparableCellValueKey(
        cellValue,
        draftDefinition.caseSensitive,
      ),
      isValueSelected = selectedValues.some(
        (selectedValue) =>
          comparableCellValueKey(
            selectedValue,
            draftDefinition.caseSensitive,
          ) === valueKey,
      );
    upsertColumnCriterion({
      ...columnCriterion,
      values: isValueSelected
        ? selectedValues.filter(
            (selectedValue) =>
              comparableCellValueKey(
                selectedValue,
                draftDefinition.caseSensitive,
              ) !== valueKey,
          )
        : [...selectedValues, cellValue],
    });
  }
  function updateAdvancedCondition(groupIndex, conditionIndex, nextCondition) {
    setDraftDefinition((previousDraft) => ({
      ...previousDraft,
      advanced: previousDraft.advanced?.map(
        (conditionGroup, currentGroupIndex) =>
          currentGroupIndex === groupIndex
            ? conditionGroup.map((condition, currentConditionIndex) =>
                currentConditionIndex === conditionIndex
                  ? nextCondition
                  : condition,
              )
            : conditionGroup,
      ),
    }));
  }
  function renderConditionEditor(
    condition,
    onConditionChange,
    conditionLabel,
    showColumn = false,
  ) {
    return i.jsxs("div", {
      className: "lg-filter-condition",
      children: [
        showColumn &&
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Column" }),
              i.jsx("select", {
                "aria-label": `${conditionLabel} column`,
                value: condition.column,
                onChange: (event) =>
                  onConditionChange({
                    ...condition,
                    column: Number(event.target.value),
                  }),
                children: columnOptions.map((columnOption) =>
                  i.jsx(
                    "option",
                    {
                      value: columnOption.column,
                      children: columnOption.label,
                    },
                    columnOption.column,
                  ),
                ),
              }),
            ],
          }),
        i.jsxs("label", {
          children: [
            i.jsx(LocalizedText, { children: "Type" }),
            i.jsx("select", {
              "aria-label": `${conditionLabel} type`,
              value: condition.type,
              onChange: (event) =>
                onConditionChange({
                  ...condition,
                  type: event.target.value,
                  operator: "equals",
                  value: event.target.value === "number" ? 0 : "",
                }),
              children: ["value", "text", "number", "date"].map((typeOption) =>
                i.jsx(
                  "option",
                  {
                    value: typeOption,
                    children:
                      typeOption === "value"
                        ? "Typed value"
                        : typeOption[0].toUpperCase() + typeOption.slice(1),
                  },
                  typeOption,
                ),
              ),
            }),
          ],
        }),
        i.jsxs("label", {
          children: [
            i.jsx(LocalizedText, { children: "Comparison" }),
            i.jsx("select", {
              "aria-label": `${conditionLabel} comparison`,
              value: condition.operator,
              onChange: (event) =>
                onConditionChange({
                  ...condition,
                  operator: event.target.value,
                }),
              children: Object.entries(filterOperatorLabels)
                .filter(
                  ([operatorKey]) =>
                    condition.type === "text" ||
                    [
                      "equals",
                      "notEquals",
                      "blank",
                      "notBlank",
                      ...(condition.type === "value"
                        ? []
                        : ["gt", "gte", "lt", "lte"]),
                    ].includes(operatorKey),
                )
                .map(([operatorKey, operatorLabel]) =>
                  i.jsx(
                    "option",
                    { value: operatorKey, children: operatorLabel },
                    operatorKey,
                  ),
                ),
            }),
          ],
        }),
        !["blank", "notBlank"].includes(condition.operator) &&
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Value" }),
              condition.type === "number"
                ? i.jsx(LocalizedNumberField, {
                    "aria-label": `${conditionLabel} value`,
                    value:
                      typeof condition.value == "number" ? condition.value : "",
                    onValueChange: (numberValue) =>
                      onConditionChange({
                        ...condition,
                        value: Number(numberValue),
                      }),
                  })
                : i.jsx("input", {
                    "aria-label": `${conditionLabel} value`,
                    value: String(condition.value ?? ""),
                    type: "text",
                    placeholder:
                      condition.type === "date" ? "2026-09-17" : "Match value",
                    onChange: (event) =>
                      onConditionChange({
                        ...condition,
                        value: event.target.value,
                      }),
                  }),
            ],
          }),
      ],
    });
  }
  return (
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        dialogRef.current?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []),
    i.jsx("div", {
      className: "lg-filter-overlay",
      children: i.jsxs("section", {
        className: "lg-filter-dialog",
        role: "dialog",
        "aria-modal": "true",
        "aria-labelledby": titleId,
        ref: dialogRef,
        tabIndex: -1,
        onKeyDown: (event) => {
          var lastFocusable, firstFocusable;
          if (
            (event.stopPropagation(),
            event.key === "Escape" && (event.preventDefault(), props.onClose()),
            event.key === "Tab")
          ) {
            const focusableElements = [
              ...(dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
              ) ?? []),
            ];
            event.shiftKey &&
            [focusableElements[0], dialogRef.current].includes(
              document.activeElement,
            )
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
                    id: titleId,
                    children: i.jsx(LocalizedText, {
                      children: "Filter population",
                    }),
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Keep complete records and their original worksheet addresses.",
                    }),
                  }),
                ],
              }),
              i.jsx("button", {
                "aria-label": text("Close population filter"),
                onClick: props.onClose,
                lang: language("Close population filter"),
                "data-ui-label": "Close population filter",
                children: "×",
              }),
            ],
          }),
          i.jsxs("div", {
            className: "lg-filter-scroll",
            children: [
              i.jsx("p", {
                className: "lg-filter-disclosure",
                children: i.jsx(LocalizedText, {
                  children:
                    "View only · first range row is the header. Edits pause this view until you review and reapply it. Native JSON, XLSX, printing, other analysis tools and AI selections use full source data. CSV/TSV and Audit Checks offer an explicit full or visible population.",
                }),
              }),
              i.jsxs("details", {
                className: "lg-saved-filter-views",
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Saved filter views (",
                      }),
                      props.workbook.savedFilterViews?.length ?? 0,
                      ")",
                    ],
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Save named criteria in the native workbook and browser save. Loading never hides rows automatically: check the range, headers and date rules, then preview and apply. XLSX does not retain saved views.",
                    }),
                  }),
                  i.jsxs("div", {
                    className: "lg-filter-controls",
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Saved view" }),
                          i.jsxs("select", {
                            "aria-label": text("Saved filter view"),
                            value: selectedViewId,
                            onChange: (event) => {
                              (setSelectedViewId(event.target.value),
                                setViewName(
                                  props.workbook.savedFilterViews?.find(
                                    (savedFilterView) =>
                                      savedFilterView.id === event.target.value,
                                  )?.name ?? "",
                                ),
                                setStatusMessage(""),
                                setErrorMessage(""));
                            },
                            lang: language("Saved filter view"),
                            "data-ui-label": "Saved filter view",
                            children: [
                              i.jsx("option", {
                                value: "",
                                children: "Choose a saved view",
                              }),
                              props.workbook.savedFilterViews?.map(
                                (savedFilterView) =>
                                  i.jsxs(
                                    "option",
                                    {
                                      value: savedFilterView.id,
                                      children: [
                                        savedFilterView.name,
                                        " ·",
                                        " ",
                                        props.workbook.sheets.find(
                                          (sheet) =>
                                            sheet.id ===
                                            savedFilterView.sheetId,
                                        )?.name ??
                                          `${savedFilterView.sheetName} (sheet missing)`,
                                      ],
                                    },
                                    savedFilterView.id,
                                  ),
                              ),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "View name" }),
                          i.jsx("input", {
                            "aria-label": text("Filter view name"),
                            maxLength: 80,
                            value: viewName,
                            onChange: (event) =>
                              setViewName(event.target.value),
                            lang: language("Filter view name"),
                            "data-ui-label": "Filter view name",
                          }),
                        ],
                      }),
                    ],
                  }),
                  selectedView &&
                    i.jsxs("p", {
                      children: [
                        i.jsx(LocalizedText, { children: "Saved range " }),
                        selectedView.definition.range,
                        " ·",
                        " ",
                        selectedView.population.visibleRows,
                        " ",
                        i.jsx(LocalizedText, { children: "visible /" }),
                        " ",
                        selectedView.population.totalRows,
                        " ",
                        i.jsx(LocalizedText, {
                          children: "total at source revision",
                        }),
                        " ",
                        selectedView.sourceRevision,
                        i.jsx(LocalizedText, {
                          children: ". Current revision",
                        }),
                        " ",
                        props.workbook.revision,
                        i.jsx(LocalizedText, {
                          children:
                            ". These are historical counts, not current results.",
                        }),
                        selectedView.sourceWorkbookId !== props.workbook.id &&
                          " The preview was saved in the original workbook; this is an independent copy.",
                        selectedView.sheetId !== props.sheetId &&
                          " Open the source worksheet to load these criteria. A missing worksheet must be restored first.",
                        JSON.stringify(selectedView.datePolicy) !==
                          JSON.stringify(
                            normalizeDatePolicy(props.workbook.datePolicy),
                          ) &&
                          " Date policy changed. Check all date criteria against the current workbook policy.",
                      ],
                    }),
                  i.jsxs("div", {
                    className: "lg-filter-actions",
                    children: [
                      i.jsx("button", {
                        disabled:
                          !selectedView ||
                          selectedView.sheetId !== props.sheetId,
                        onClick: () => {
                          !selectedView ||
                            selectedView.sheetId !== props.sheetId ||
                            (setDraftDefinition(
                              structuredClone(selectedView.definition),
                            ),
                            setSelectedColumn(
                              parseRangeAddress(selectedView.definition.range)
                                .firstCol,
                            ),
                            setValueSearch(""),
                            setPreview(null),
                            setErrorMessage(""),
                            setStatusMessage(
                              "Criteria loaded for review. Saved coordinates do not move with inserted, deleted or sorted records. Check the range and headers, then preview current results.",
                            ));
                        },
                        children: i.jsx(LocalizedText, {
                          children: "Load criteria for review",
                        }),
                      }),
                      i.jsx("button", {
                        disabled:
                          !canEditFilters ||
                          !currentPreview ||
                          !viewName.trim(),
                        onClick: () => saveFilterView(),
                        children: i.jsx(LocalizedText, {
                          children: "Save as new view",
                        }),
                      }),
                      i.jsx("button", {
                        disabled:
                          !canEditFilters ||
                          !currentPreview ||
                          !selectedView ||
                          selectedView.sheetId !== props.sheetId ||
                          !viewName.trim(),
                        onClick: () => saveFilterView(true),
                        children: i.jsx(LocalizedText, {
                          children: "Replace saved criteria",
                        }),
                      }),
                      i.jsx("button", {
                        disabled: !canUpdateSelectedView || !viewName.trim(),
                        onClick: () => {
                          try {
                            if (!canUpdateSelectedView) return;
                            (props.onWorkbookChange(
                              addSavedFilterView(
                                props.workbook,
                                selectedViewId,
                                viewName,
                              ),
                              "Rename saved filter view",
                            ),
                              setStatusMessage(
                                "Saved view renamed. Its criteria and historical counts are unchanged.",
                              ),
                              setErrorMessage(""));
                          } catch (error) {
                            setErrorMessage(
                              error instanceof Error
                                ? error.message
                                : "Cannot rename this view.",
                            );
                          }
                        },
                        children: i.jsx(LocalizedText, {
                          children: "Rename view",
                        }),
                      }),
                      i.jsx("button", {
                        disabled: !canUpdateSelectedView,
                        onClick: () => {
                          try {
                            if (!canUpdateSelectedView) return;
                            (props.onWorkbookChange(
                              removeSavedFilterView(
                                props.workbook,
                                selectedViewId,
                              ),
                              "Delete saved filter view",
                            ),
                              setSelectedViewId(""),
                              setViewName(""),
                              setStatusMessage(
                                "Saved view deleted. Source cells and criteria in this editor are retained. Undo can restore the saved view.",
                              ),
                              setErrorMessage(""));
                          } catch (error) {
                            setErrorMessage(
                              error instanceof Error
                                ? error.message
                                : "Cannot delete this view.",
                            );
                          }
                        },
                        children: i.jsx(LocalizedText, {
                          children: "Delete view",
                        }),
                      }),
                    ],
                  }),
                  !canEditFilters &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Saving is unavailable in this read-only, protected or view-only host. Loading, previewing and applying session views remain available.",
                      }),
                    }),
                  statusMessage &&
                    i.jsx("p", { role: "status", children: statusMessage }),
                ],
              }),
              i.jsxs("div", {
                className: "lg-filter-controls",
                children: [
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Population range" }),
                      i.jsx("input", {
                        "aria-label": text("Population range"),
                        value: draftDefinition.range,
                        onChange: (event) =>
                          setDraftDefinition((previousDraft) => ({
                            ...previousDraft,
                            range: event.target.value,
                            columns: [],
                            advanced: [],
                            unique: false,
                          })),
                        lang: language("Population range"),
                        "data-ui-label": "Population range",
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Column" }),
                      i.jsx("select", {
                        "aria-label": text("Filter column"),
                        value: selectedColumn,
                        onChange: (event) => {
                          (setSelectedColumn(Number(event.target.value)),
                            setValueSearch(""));
                        },
                        lang: language("Filter column"),
                        "data-ui-label": "Filter column",
                        children: columnOptions.map((columnOption) =>
                          i.jsx(
                            "option",
                            {
                              value: columnOption.column,
                              children: columnOption.label,
                            },
                            columnOption.column,
                          ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: !!draftDefinition.caseSensitive,
                        onChange: (event) =>
                          setDraftDefinition((previousDraft) => ({
                            ...previousDraft,
                            caseSensitive: event.target.checked,
                          })),
                      }),
                      i.jsx(LocalizedText, { children: "Case sensitive" }),
                    ],
                  }),
                ],
              }),
              i.jsxs("div", {
                className: "lg-filter-columns",
                children: [
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Values / slicer",
                        }),
                      }),
                      i.jsx("input", {
                        "aria-label": text("Search filter values"),
                        placeholder: text("Search available values"),
                        value: valueSearch,
                        onChange: (event) => setValueSearch(event.target.value),
                        lang: language("Search filter values"),
                        "data-ui-label": "Search filter values",
                      }),
                      i.jsxs("div", {
                        className: "lg-filter-actions",
                        children: [
                          i.jsx("button", {
                            onClick: () =>
                              upsertColumnCriterion({
                                ...columnCriterion,
                                values: undefined,
                              }),
                            children: i.jsx(LocalizedText, {
                              children: "Select all values",
                            }),
                          }),
                          i.jsx("button", {
                            onClick: () =>
                              upsertColumnCriterion({
                                ...columnCriterion,
                                values: [],
                              }),
                            children: i.jsx(LocalizedText, {
                              children: "Select no values",
                            }),
                          }),
                        ],
                      }),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Other column and advanced criteria limit available values. Search narrows this list only. Manual hiding does not change the value list.",
                        }),
                      }),
                      facetState.error &&
                        i.jsx("p", {
                          role: "alert",
                          children: facetState.error,
                        }),
                      facetState.result?.truncated &&
                        i.jsxs("p", {
                          role: "status",
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Showing 1,000 of ",
                            }),
                            facetState.result.distinct,
                            " ",
                            i.jsx(LocalizedText, {
                              children:
                                "values. Search or use conditions. Clear this column to include all values.",
                            }),
                          ],
                        }),
                      i.jsx("div", {
                        className: "lg-filter-values",
                        role: "group",
                        "aria-label": text("Value slicer"),
                        lang: language("Value slicer"),
                        "data-ui-label": "Value slicer",
                        children: facetState.result?.values.map((value) => {
                          const isSelected =
                            columnCriterion.values === undefined ||
                            columnCriterion.values.some(
                              (selectedValue) =>
                                comparableCellValueKey(
                                  selectedValue,
                                  draftDefinition.caseSensitive,
                                ) === value.key,
                            );
                          return i.jsxs(
                            "button",
                            {
                              "aria-pressed": isSelected,
                              disabled:
                                columnCriterion.values === undefined &&
                                !!facetState.result?.truncated,
                              onClick: () => toggleFilterValue(value.value),
                              children: [
                                i.jsxs("span", {
                                  children: [
                                    value.value === null || value.value === ""
                                      ? "(Blanks)"
                                      : String(value.value),
                                    " ",
                                    i.jsx("small", {
                                      children:
                                        value.value === null ||
                                        value.value === ""
                                          ? ""
                                          : typeof value.value,
                                    }),
                                  ],
                                }),
                                i.jsx("span", { children: value.count }),
                              ],
                            },
                            value.key,
                          );
                        }),
                      }),
                      i.jsx("button", {
                        onClick: () =>
                          setDraftDefinition((previousDraft) => ({
                            ...previousDraft,
                            columns: previousDraft.columns.filter(
                              (column) => column.column !== selectedColumn,
                            ),
                          })),
                        children: i.jsx(LocalizedText, {
                          children: "Clear this column",
                        }),
                      }),
                    ],
                  }),
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Custom column filter",
                        }),
                      }),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Column filters intersect. Value selections and custom conditions both apply.",
                        }),
                      }),
                      (columnCriterion.conditions ?? []).map(
                        (condition, conditionIndex) =>
                          i.jsxs(
                            "div",
                            {
                              children: [
                                renderConditionEditor(
                                  condition,
                                  (nextCondition) =>
                                    upsertColumnCriterion({
                                      ...columnCriterion,
                                      conditions:
                                        columnCriterion.conditions?.map(
                                          (condition, candidateIndex) =>
                                            candidateIndex === conditionIndex
                                              ? nextCondition
                                              : condition,
                                        ),
                                    }),
                                  `Column condition ${conditionIndex + 1}`,
                                ),
                                i.jsxs("button", {
                                  onClick: () =>
                                    upsertColumnCriterion({
                                      ...columnCriterion,
                                      conditions:
                                        columnCriterion.conditions?.filter(
                                          (condition, candidateIndex) =>
                                            candidateIndex !== conditionIndex,
                                        ),
                                    }),
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Remove condition ",
                                    }),
                                    conditionIndex + 1,
                                  ],
                                }),
                              ],
                            },
                            conditionIndex,
                          ),
                      ),
                      i.jsx("button", {
                        disabled:
                          (columnCriterion.conditions?.length ?? 0) >= 2,
                        onClick: () =>
                          upsertColumnCriterion({
                            ...columnCriterion,
                            conditions: [
                              ...(columnCriterion.conditions ?? []),
                              createDefaultFilterCriterion(selectedColumn),
                            ],
                          }),
                        children: i.jsx(LocalizedText, {
                          children: "Add column condition",
                        }),
                      }),
                      (columnCriterion.conditions?.length ?? 0) > 1 &&
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Combine conditions",
                            }),
                            i.jsxs("select", {
                              "aria-label": text("Combine column conditions"),
                              value: columnCriterion.join ?? "and",
                              onChange: (event) =>
                                upsertColumnCriterion({
                                  ...columnCriterion,
                                  join: event.target.value,
                                }),
                              lang: language("Combine column conditions"),
                              "data-ui-label": "Combine column conditions",
                              children: [
                                i.jsx("option", {
                                  value: "and",
                                  children: "AND · both",
                                }),
                                i.jsx("option", {
                                  value: "or",
                                  children: "OR · either",
                                }),
                              ],
                            }),
                          ],
                        }),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Numbers match numeric cells only; text numbers remain text. Dates use the workbook's 1900/1904 and MDY/DMY policy. Blank means an empty cell or empty text.",
                        }),
                      }),
                    ],
                  }),
                ],
              }),
              i.jsxs("fieldset", {
                children: [
                  i.jsx("legend", {
                    children: i.jsx(LocalizedText, {
                      children: "Advanced criteria",
                    }),
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Meet every condition within a group (AND), and at least one group (OR). These groups also intersect with all column filters.",
                    }),
                  }),
                  draftDefinition.advanced?.map((conditionGroup, groupIndex) =>
                    i.jsxs(
                      "div",
                      {
                        className: "lg-filter-group",
                        children: [
                          i.jsxs("strong", {
                            children: [
                              i.jsx(LocalizedText, { children: "OR group " }),
                              groupIndex + 1,
                            ],
                          }),
                          conditionGroup.map((condition, conditionIndex) =>
                            i.jsxs(
                              "div",
                              {
                                children: [
                                  renderConditionEditor(
                                    condition,
                                    (nextCondition) =>
                                      updateAdvancedCondition(
                                        groupIndex,
                                        conditionIndex,
                                        nextCondition,
                                      ),
                                    `Group ${groupIndex + 1} condition ${conditionIndex + 1}`,
                                    true,
                                  ),
                                  i.jsxs("button", {
                                    onClick: () =>
                                      setDraftDefinition((previousDraft) => ({
                                        ...previousDraft,
                                        advanced: previousDraft.advanced
                                          ?.map((ge, candidateGroupIndex) =>
                                            candidateGroupIndex === groupIndex
                                              ? ge.filter(
                                                  (
                                                    we,
                                                    candidateConditionIndex,
                                                  ) =>
                                                    candidateConditionIndex !==
                                                    conditionIndex,
                                                )
                                              : ge,
                                          )
                                          .filter((ge) => ge.length),
                                      })),
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Remove group ",
                                      }),
                                      groupIndex + 1,
                                      " ",
                                      i.jsx(LocalizedText, {
                                        children: "condition ",
                                      }),
                                      conditionIndex + 1,
                                    ],
                                  }),
                                ],
                              },
                              conditionIndex,
                            ),
                          ),
                          i.jsxs("button", {
                            onClick: () =>
                              setDraftDefinition((previousDraft) => ({
                                ...previousDraft,
                                advanced: previousDraft.advanced?.map(
                                  (le, candidateGroupIndex) =>
                                    candidateGroupIndex === groupIndex
                                      ? [
                                          ...le,
                                          createDefaultFilterCriterion(
                                            selectedColumn,
                                          ),
                                        ]
                                      : le,
                                ),
                              })),
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Add AND condition to group ",
                              }),
                              groupIndex + 1,
                            ],
                          }),
                          i.jsxs("button", {
                            onClick: () =>
                              setDraftDefinition((previousDraft) => ({
                                ...previousDraft,
                                advanced: previousDraft.advanced?.filter(
                                  (le, candidateGroupIndex) =>
                                    candidateGroupIndex !== groupIndex,
                                ),
                              })),
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Remove group ",
                              }),
                              groupIndex + 1,
                            ],
                          }),
                        ],
                      },
                      groupIndex,
                    ),
                  ),
                  i.jsx("button", {
                    onClick: () =>
                      setDraftDefinition((previousDraft) => ({
                        ...previousDraft,
                        advanced: [
                          ...(previousDraft.advanced ?? []),
                          [createDefaultFilterCriterion(selectedColumn)],
                        ],
                      })),
                    children: i.jsx(LocalizedText, {
                      children: "Add OR group",
                    }),
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: !!draftDefinition.unique,
                        onChange: (event) =>
                          setDraftDefinition((previousDraft) => ({
                            ...previousDraft,
                            unique: event.target.checked,
                          })),
                      }),
                      i.jsx(LocalizedText, {
                        children:
                          "Unique complete records only (keep first matching source row)",
                      }),
                    ],
                  }),
                ],
              }),
              i.jsxs("div", {
                className: "lg-filter-actions",
                children: [
                  i.jsx("button", {
                    onClick: () =>
                      setDraftDefinition((previousDraft) => ({
                        ...previousDraft,
                        columns: [],
                        advanced: [],
                        unique: false,
                      })),
                    children: i.jsx(LocalizedText, {
                      children: "Clear all criteria",
                    }),
                  }),
                  i.jsx("button", {
                    onClick: () => {
                      try {
                        const filterResult = applyFilterConditions(
                          props.workbook,
                          props.sheetId,
                          draftDefinition,
                        );
                        (setPreview({
                          book: props.workbook,
                          definition: JSON.stringify(draftDefinition),
                          result: filterResult,
                        }),
                          setErrorMessage(""));
                      } catch (error) {
                        (setErrorMessage(
                          error instanceof Error
                            ? error.message
                            : "Filter preview failed.",
                        ),
                          setPreview(null));
                      }
                    },
                    children: i.jsx(LocalizedText, {
                      children: "Preview population",
                    }),
                  }),
                ],
              }),
              errorMessage &&
                i.jsx("p", { role: "alert", children: errorMessage }),
              currentPreview &&
                i.jsxs("div", {
                  className: "lg-filter-preview",
                  role: "status",
                  ref: previewRef,
                  children: [
                    i.jsxs("strong", {
                      children: [
                        currentPreview.context.visibleRows,
                        " ",
                        i.jsx(LocalizedText, { children: "visible / " }),
                        currentPreview.context.totalRows,
                        " ",
                        i.jsx(LocalizedText, { children: "total data rows" }),
                      ],
                    }),
                    i.jsxs("p", {
                      children: [
                        currentPreview.context.matchedRows,
                        " ",
                        i.jsx(LocalizedText, { children: "match criteria ·" }),
                        " ",
                        currentPreview.context.filteredRows,
                        " ",
                        i.jsx(LocalizedText, { children: "filtered out ·" }),
                        " ",
                        currentPreview.context.hiddenMatchingRows,
                        " ",
                        i.jsx(LocalizedText, {
                          children: "matching rows manually hidden.",
                        }),
                      ],
                    }),
                    currentPreview.context.visibleRows === 0
                      ? i.jsx("p", {
                          children: i.jsx(LocalizedText, {
                            children:
                              "No matching visible records. The full source population is unchanged.",
                          }),
                        })
                      : i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Original worksheet rows:",
                            }),
                            " ",
                            currentPreview.context.visibleSourceRows
                              .slice(0, 25)
                              .join(", "),
                            currentPreview.context.visibleRows > 25
                              ? " … (first 25 shown)"
                              : "",
                          ],
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
                className: "lg-filter-primary",
                disabled: !currentPreview,
                onClick: () => {
                  if (currentPreview)
                    try {
                      const validatedFilter = getPreviewedFilter();
                      props.onApply(validatedFilter.context.definition);
                    } catch (error) {
                      (setErrorMessage(
                        error instanceof Error
                          ? error.message
                          : "Filter apply failed.",
                      ),
                        setPreview(null));
                    }
                },
                children: i.jsx(LocalizedText, {
                  children: "Apply filter view",
                }),
              }),
            ],
          }),
        ],
      }),
    })
  );
}
function ValidationReviewDialog({
  error,
  readOnly,
  onAccept,
  onCancel,
  onRetry,
}) {
  const dialogRef = ne.useRef(null),
    hasStopIssue = error.issues.some(
      (issue) => issue.style === "stop" || issue.evaluationError,
    );
  return (
    ne.useEffect(() => {
      dialogRef.current?.querySelector("button")?.focus();
    }, []),
    i.jsx("div", {
      className: "wt-backdrop",
      style: { zIndex: 900 },
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" && onCancel(),
          event.key === "Tab")
        ) {
          const focusableButtons = [
              ...(dialogRef.current?.querySelectorAll(
                "button:not(:disabled)",
              ) ?? []),
            ],
            firstButton = focusableButtons[0],
            lastButton = focusableButtons.at(-1);
          event.shiftKey && document.activeElement === firstButton
            ? (event.preventDefault(), lastButton == null || lastButton.focus())
            : !event.shiftKey &&
              document.activeElement === lastButton &&
              (event.preventDefault(),
              firstButton == null || firstButton.focus());
        }
      },
      children: i.jsxs("div", {
        className: "wt-dialog",
        role: "alertdialog",
        "aria-modal": "true",
        "aria-label": hasStopIssue
          ? "Data validation stopped this change"
          : "Review validation warnings",
        ref: dialogRef,
        children: [
          i.jsx("header", {
            children: i.jsx("h2", {
              children: hasStopIssue
                ? "Correct invalid entries"
                : "Review validation warnings",
            }),
          }),
          i.jsxs("div", {
            className: "wt-body",
            children: [
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "No cells have changed." }),
                  " ",
                  hasStopIssue
                    ? "Correct these entries or discard the entire batch."
                    : "Accepting keeps the invalid entries and records your decision for the complete batch.",
                ],
              }),
              i.jsxs("p", {
                children: [
                  error.issues.length,
                  " ",
                  i.jsx(LocalizedText, { children: "rule violations" }),
                ],
              }),
              error.issues.slice(0, 500).map((issue, issueIndex) =>
                i.jsxs(
                  "div",
                  {
                    children: [
                      i.jsxs("strong", {
                        children: [
                          issue.sheetName,
                          "!",
                          issue.address,
                          " · ",
                          issue.style,
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsxs("strong", { children: [issue.title, ": "] }),
                          issue.message,
                          " ",
                          issue.reason,
                        ],
                      }),
                    ],
                  },
                  issueIndex,
                ),
              ),
            ],
          }),
          i.jsxs("footer", {
            children: [
              i.jsx("button", {
                onClick: onRetry,
                children: i.jsx(LocalizedText, { children: "Return to edit" }),
              }),
              i.jsx("button", {
                onClick: onCancel,
                children: i.jsx(LocalizedText, { children: "Discard changes" }),
              }),
              !hasStopIssue &&
                i.jsx("button", {
                  disabled: readOnly,
                  onClick: onAccept,
                  children: i.jsx(LocalizedText, {
                    children: "Accept warnings and apply batch",
                  }),
                }),
            ],
          }),
        ],
      }),
    })
  );
}

export { ValidationReviewDialog, FilterViewsDialog };
