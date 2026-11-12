import i from "react/jsx-runtime";
import ne from "react";
import { X } from "lucide-react";
import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  assertWorkbookStructureEditable,
  consolidationFormatVersion,
  serializeTypedKeyValues,
  createWorksheetRecord,
  sumConsolidationAmounts,
  formatOccupiedRangeAddress,
  calculateWorkbookWithSettings,
  parseRangeAddress,
  consolidationLimits,
  throwDaxError,
  normalizeWorkbookInput,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { importWorkbookFile } from "./07-file-import-and-worker-dispatch.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import {
  readSourceTableRecords,
  buildConsolidationOutputCells,
  createConsolidationPreview,
  buildAuthorizedContextSnapshot,
} from "./12-consolidation-and-dependency-inspection.js";

function publishConsolidatedSchedule(workbook, plan, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  if (
    (assertWorkbookStructureEditable(workbook),
    fingerprintCanonicalValue(workbook) !== plan.workbookHash ||
      buildAuthorizedContextSnapshot(workbook, options) !== plan.contextHash)
  )
    throw Error(
      "Workbook or authorization changed. Preview consolidation again.",
    );
  const consolidationPreview = createConsolidationPreview(
    workbook,
    plan.request,
    options,
  );
  if (
    fingerprintCanonicalValue(consolidationPreview) !==
    fingerprintCanonicalValue(plan)
  )
    throw Error(
      "Source data or the reviewed plan changed. Preview consolidation again.",
    );
  if (!consolidationPreview.review.complete)
    throw Error(
      "Resolve all consolidation errors before publishing. No partial result is permitted.",
    );
  const definition = consolidationPreview.request.definition,
    review = consolidationPreview.review,
    worksheetRecord = createWorksheetRecord(plan.request.outputName),
    amountColumnIndex = definition.categoryLabels.length + 4,
    amountColumnLetter = columnIndexToLetters(amountColumnIndex),
    columnCount = definition.categoryLabels.length + 7,
    headerRow = [
      "Row type",
      ...definition.categoryLabels,
      "Period",
      "Currency",
      "Unit",
      "Amount",
      "Source",
      "Link state",
    ];
  let nextRowNumber = 1;
  const subtotalChecks = [],
    appendRow = (values, formulaColumnIndex = -1, note = "") => {
      const rowNumber = nextRowNumber++;
      return (
        values.forEach((raw, columnIndex) => {
          worksheetRecord.cells[
            `${columnIndexToLetters(columnIndex)}${rowNumber}`
          ] = {
            raw: raw,
            ...(typeof raw == "string" &&
            raw.startsWith("=") &&
            columnIndex !== formulaColumnIndex
              ? { literal: true }
              : {}),
            ...(note && columnIndex === amountColumnIndex
              ? { note: note }
              : {}),
            ...(columnIndex === amountColumnIndex && rowNumber > 1
              ? { format: "number", numFmt: "#,##0.00;[Red](#,##0.00)" }
              : {}),
            ...(rowNumber === 1 || values[0] !== "Detail"
              ? {
                  style: {
                    bold: true,
                    fill: rowNumber === 1 ? "#362C63" : "#EEEAF8",
                    color: rowNumber === 1 ? "#FFFFFF" : "#29243C",
                  },
                }
              : {}),
          };
        }),
        rowNumber
      );
    };
  appendRow(headerRow);
  const appendDetailRow = (contribution) =>
      appendRow(
        [
          "Detail",
          ...contribution.categories,
          contribution.period,
          contribution.currency,
          contribution.unit,
          contribution.link === "live"
            ? `='${contribution.sheetName.replaceAll("'", "''")}'!$${contribution.address.replace(/\d+$/, "")}$${contribution.row}`
            : contribution.amount,
          contribution.locator,
          contribution.link === "live"
            ? "Live amount · reviewed context"
            : contribution.link === "approved snapshot"
              ? "Approved workbook snapshot"
              : "Reviewed snapshot",
        ],
        contribution.link === "live" ? amountColumnIndex : -1,
        `Reviewed amount ${contribution.amount}; period ${contribution.period}; currency ${contribution.currency}; unit ${contribution.unit}. ${contribution.locator}. Source ID ${contribution.sourceId}; workbook ${contribution.workbookId}; worksheet ${contribution.sheetId}. Source freshness is displayed in the native workbook banner.`,
      ),
    appendSubtotalGroups = (contributions, categoryDepth) => {
      const groupsByKey = new Map();
      for (const contribution of contributions) {
        const groupKey = serializeTypedKeyValues(
            contribution.categories.slice(0, categoryDepth),
          ),
          groupContributions = groupsByKey.get(groupKey) ?? [];
        (groupContributions.push(contribution),
          groupsByKey.set(groupKey, groupContributions));
      }
      for (const group of groupsByKey.values()) {
        const groupStartRow = nextRowNumber;
        (categoryDepth < definition.categoryLabels.length
          ? appendSubtotalGroups(group, categoryDepth + 1)
          : group.forEach(appendDetailRow),
          subtotalChecks.push({
            address: `${amountColumnLetter}${nextRowNumber}`,
            expected: sumConsolidationAmounts(
              group.map((contribution) => contribution.amount),
            ),
            count: group.length,
          }));
        const groupEndRow = nextRowNumber - 1,
          subtotalCategories = Array.from(
            { length: definition.categoryLabels.length },
            (_item, categoryIndex) =>
              categoryIndex < categoryDepth
                ? group[0].categories[categoryIndex]
                : null,
          );
        (appendRow(
          [
            `Subtotal ${categoryDepth}`,
            ...subtotalCategories,
            definition.period,
            definition.currency,
            definition.unit,
            `=SUBTOTAL(9,${amountColumnLetter}${groupStartRow}:${amountColumnLetter}${groupEndRow})`,
            `${group.length} contributions`,
            "Review source status",
          ],
          amountColumnIndex,
        ),
          (worksheetRecord.outlineGroups ??= []).push({
            id: crypto.randomUUID(),
            axis: "row",
            start: groupStartRow - 1,
            end: groupEndRow - 1,
            collapsed: definition.collapseDetails,
          }));
      }
    };
  if (
    (appendSubtotalGroups(review.contributions, 1),
    subtotalChecks.push({
      address: `${amountColumnLetter}${nextRowNumber}`,
      expected: review.grandTotal,
      count: review.contributions.length,
    }),
    appendRow(
      [
        "Grand total",
        ...definition.categoryLabels.map(() => null),
        definition.period,
        definition.currency,
        definition.unit,
        `=SUBTOTAL(9,${amountColumnLetter}2:${amountColumnLetter}${nextRowNumber - 1})`,
        `${review.contributions.length} contributions`,
        "Review source status",
      ],
      amountColumnIndex,
      `Reviewed grand total ${review.grandTotal}; source controls ${review.sourceTotal}; difference ${review.difference}. No implicit conversion.`,
    ),
    nextRowNumber - 1 > consolidationLimits.outputRows ||
      (nextRowNumber - 1) * columnCount > consolidationLimits.outputCells ||
      workbook.sheets.reduce(
        (total, sheet) => total + Object.keys(sheet.cells).length,
        0,
      ) +
        Object.keys(worksheetRecord.cells).length >
        workbookLimits.maxCells)
  )
    throw Error(
      "Consolidation exceeds the output or workbook cell budget. No worksheet was published.",
    );
  ((worksheetRecord.frozenRows = 1),
    (worksheetRecord.columnWidths = Object.fromEntries(
      headerRow.map((_header, columnIndex) => [
        columnIndex,
        columnIndex === columnCount - 2
          ? 340
          : columnIndex === columnCount - 1
            ? 260
            : 130,
      ]),
    )));
  const outputRange = `A1:${columnIndexToLetters(columnCount - 1)}${nextRowNumber - 1}`,
    origin = {
      workbookId: workbook.id,
      revision: workbook.revision,
      at: new Date().toISOString(),
    };
  worksheetRecord.consolidation = {
    engine: consolidationFormatVersion,
    definition: definition,
    review: review,
    origin: origin,
    evidenceHash: fingerprintCanonicalValue({
      definition: definition,
      review: review,
      origin: origin,
    }),
    outputRange: outputRange,
    outputHash: buildConsolidationOutputCells(worksheetRecord, outputRange),
  };
  const workbookInput = normalizeWorkbookInput({
      ...workbook,
      sheets: [...workbook.sheets, worksheetRecord],
    }),
    workbookWithSettings = calculateWorkbookWithSettings(workbookInput);
  for (const subtotalCheck of subtotalChecks) {
    const calculatedValue = workbookWithSettings(
        worksheetRecord.id,
        subtotalCheck.address,
      ),
      tolerance =
        Number.EPSILON *
        Math.max(1, Math.abs(subtotalCheck.expected)) *
        subtotalCheck.count *
        4;
    if (
      typeof calculatedValue != "number" ||
      !Number.isFinite(calculatedValue) ||
      Math.abs(calculatedValue - subtotalCheck.expected) > tolerance
    )
      throw Error(
        "Calculated subtotals do not reconcile with the reviewed amounts, or a calculation limit was reached. Reduce numeric cancellation or explicitly round source amounts before publishing.",
      );
  }
  return workbookInput;
}
function checkConsolidationFreshness(workbook, worksheetRecord, options = {}) {
  const consolidation = worksheetRecord.consolidation;
  if (!consolidation)
    throw Error("This worksheet has no consolidation record.");
  const messages = [];
  try {
    (consolidation.origin.workbookId !== workbook.id &&
      messages.push(
        "Historical result from another workbook identity. Review and rebind its sources before refreshing.",
      ),
      buildConsolidationOutputCells(
        worksheetRecord,
        consolidation.outputRange,
      ) !== consolidation.outputHash &&
        messages.push(
          "Published result cells or source notes were edited. Reconcile or publish a fresh result.",
        ));
    for (const source of consolidation.definition.sources)
      try {
        const sourceTableRecords = readSourceTableRecords(
            workbook,
            source,
            options,
          ),
          reviewedSource = consolidation.review.sources.find(
            (reviewSource) => reviewSource.sourceId === source.id,
          );
        sourceTableRecords.fingerprint !== reviewedSource?.fingerprint &&
          messages.push(
            `${source.id}: source data, context, layout or formula dependencies changed.`,
          );
      } catch (error) {
        messages.push(`${source.id}: ${error.message}`);
      }
  } catch (error) {
    messages.push(error.message);
  }
  return { state: messages.length ? "stale" : "current", messages: messages };
}
function saveBlobToFile(blob, fileName) {
  const objectUrl = URL.createObjectURL(blob),
    anchor = document.createElement("a");
  ((anchor.href = objectUrl),
    (anchor.download = fileName),
    anchor.click(),
    setTimeout(() => URL.revokeObjectURL(objectUrl), 5e3));
}
function findTableRangesForSource(workbook, source) {
  try {
    const sourceSheet = workbook?.sheets.find(
      (sheet) => sheet.id === source.sheetId,
    );
    if (!sourceSheet) return [];
    const rangeAddress = parseRangeAddress(
      sourceSheet.tables?.find((table) => table.id === source.tableId)?.range ??
        source.range,
    );
    return rangeAddress.lastCol - rangeAddress.firstCol >= 50
      ? []
      : Array.from(
          { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
          (_value, columnOffset) =>
            sourceSheet.cells[
              `${columnIndexToLetters(rangeAddress.firstCol + columnOffset)}${rangeAddress.firstRow + 1}`
            ]?.raw,
        ).filter((rawValue) => typeof rawValue == "string" && !!rawValue);
  } catch {
    return [];
  }
}
function createConsolidationDefinition(workbook, sheetId, range, definition) {
  const source = {
      id: crypto.randomUUID(),
      workbookId: workbook.id,
      sheetId: sheetId,
      range: range,
      categories: [...definition.categoryLabels],
      amountField: "Amount",
      period: { kind: "constant", value: definition.period },
      currency: { kind: "constant", value: definition.currency },
      unit: { kind: "constant", value: definition.unit },
      duplicatePolicy: "reject",
    },
    tableRangesForSource = findTableRangesForSource(workbook, source);
  ((source.categories = definition.categoryLabels.map(
    (categoryLabel, categoryIndex) =>
      tableRangesForSource.includes(categoryLabel)
        ? categoryLabel
        : (tableRangesForSource[categoryIndex] ?? categoryLabel),
  )),
    (source.amountField = tableRangesForSource.includes("Amount")
      ? "Amount"
      : (tableRangesForSource.at(-1) ?? "Amount")));
  for (const propertyName of ["period", "currency", "unit"]) {
    const field = propertyName[0].toUpperCase() + propertyName.slice(1);
    tableRangesForSource.includes(field) &&
      (source[propertyName] = { kind: "field", field: field });
  }
  return source;
}
function ConsolidationStatusBanner({
  workbook,
  sheetId,
  options = {},
  onReview,
}) {
  const { text, language } = useWorkbookMessages(),
    sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    dependencyKey = sheet?.consolidation
      ? JSON.stringify([workbook, options])
      : "",
    freshness = ne.useMemo(
      () =>
        sheet?.consolidation
          ? checkConsolidationFreshness(workbook, sheet, options)
          : null,
      [dependencyKey, sheetId],
    );
  return freshness
    ? i.jsxs("div", {
        className: `lg-consolidation-status ${freshness.state}`,
        role: "status",
        "aria-label": text("Consolidation source status"),
        lang: language("Consolidation source status"),
        "data-ui-label": "Consolidation source status",
        children: [
          i.jsxs("span", {
            children: [
              i.jsx("strong", {
                children:
                  freshness.state === "stale"
                    ? "Stale consolidation"
                    : "Consolidation sources current",
              }),
              " ",
              "·",
              " ",
              freshness.messages.length
                ? freshness.messages.join(" ")
                : "Amounts reconcile to the reviewed sources. External amounts are approved snapshots.",
            ],
          }),
          i.jsx("button", {
            onClick: onReview,
            children: i.jsx(LocalizedText, { children: "Review / refresh" }),
          }),
        ],
      })
    : null;
}
function ConsolidationDialog(props) {
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    previewRef = ne.useRef(null),
    dialogRef = ne.useRef(null),
    requestCounter = ne.useRef(0),
    latestInputStamp = ne.useRef(""),
    isMounted = ne.useRef(true),
    allowedSheetIds =
      props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
    inputStamp = JSON.stringify([
      props.workbook,
      props.sheetId,
      props.range,
      props.readOnly,
      allowedSheetIds,
      props.allowHidden,
      props.allowFileSources,
      props.approvedWorkbooks,
    ]);
  latestInputStamp.current = inputStamp;
  const [definition, setDefinition] = ne.useState(() => {
      const storedDefinition =
        props.initialDefinition ??
        props.workbook.sheets.find((sheet) => sheet.id === props.sheetId)
          ?.consolidation?.definition;
      if (storedDefinition) return structuredClone(storedDefinition);
      const definition = {
          engine: consolidationFormatVersion,
          categoryLabels: ["Category"],
          period: "2026-Q1",
          currency: "USD",
          unit: "ones",
          linkedAmounts: true,
          collapseDetails: false,
          sources: [],
        },
        sheet = props.workbook.sheets.find(
          (sheet) => sheet.id === props.sheetId,
        ),
        range = props.range.includes(":")
          ? props.range
          : formatOccupiedRangeAddress(sheet);
      return (
        (definition.sources = [
          createConsolidationDefinition(
            props.workbook,
            sheet.id,
            range,
            definition,
          ),
        ]),
        definition
      );
    }),
    [outputName, setOutputName] = ne.useState(() => {
      let candidateName = "Consolidated",
        nameSuffix = 2;
      for (
        ;
        props.workbook.sheets.some(
          (sheet) => sheet.name.toLowerCase() === candidateName.toLowerCase(),
        );
      )
        candidateName = `Consolidated ${nameSuffix++}`;
      return candidateName;
    }),
    [checked, setIncludeHidden] = ne.useState(false),
    [localApprovedWorkbooks, setLocalApprovedWorkbooks] = ne.useState([]),
    [pendingImport, setPendingImport] = ne.useState(null),
    [isInspecting, setIsInspecting] = ne.useState(false),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [preview, setPreview] = ne.useState(null),
    [requestedPage, setRequestedPage] = ne.useState(0),
    [reviewTab, setReviewTab] = ne.useState("contributions"),
    approvedWorkbooks = [
      ...(props.approvedWorkbooks ?? []),
      ...localApprovedWorkbooks.filter(
        (localApproval) =>
          !props.approvedWorkbooks?.some(
            (approvedWorkbook) =>
              approvedWorkbook.workbook.id === localApproval.workbook.id,
          ),
      ),
    ],
    options = {
      readOnly: props.readOnly,
      allowedSheetIds: allowedSheetIds,
      includeHidden: !!props.allowHidden && checked,
      approvedWorkbooks: approvedWorkbooks,
    },
    stamp = JSON.stringify([inputStamp, definition, outputName, options]),
    isPreviewCurrent = preview?.stamp === stamp,
    review = isPreviewCurrent ? preview.plan.review : null,
    workbookSources = [
      {
        workbook: props.workbook,
        allowedSheetIds: allowedSheetIds,
        includeHidden: !!props.allowHidden,
      },
      ...approvedWorkbooks,
    ];
  (ne.useEffect(() => {
    isMounted.current = true;
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("input,select,button")?.focus(),
      () => {
        ((isMounted.current = false),
          requestCounter.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      var previewElement;
      isPreviewCurrent &&
        ((previewElement = previewRef.current) == null
          ? undefined
          : previewElement.scrollIntoView
        )?.call(previewElement, {
          block: "start",
        });
    }, [preview, isPreviewCurrent]));
  const updateDefinition = (nextDefinition) => {
      (setDefinition(nextDefinition),
        setPreview(null),
        setErrorMessage(""),
        setRequestedPage(0));
    },
    updateSourceAt = (targetIndex, nextSource) =>
      updateDefinition({
        ...definition,
        sources: definition.sources.map((source, currentIndex) =>
          currentIndex === targetIndex ? nextSource : source,
        ),
      }),
    runGuarded = (action) => {
      try {
        (setErrorMessage(""), action());
      } catch (error) {
        setErrorMessage(error.message);
      }
    };
  async function inspectWorkbookFile(workbookFile) {
    if (!workbookFile) return;
    const requestId = ++requestCounter.current,
      stamp = latestInputStamp.current;
    (setPendingImport(null), setIsInspecting(true), setErrorMessage(""));
    try {
      if (!/\.(json|xlsx)$/i.test(workbookFile.name))
        throw Error("Choose a native workbook JSON or an XLSX file.");
      const importResult = await importWorkbookFile(
        workbookFile,
        cancellableOperation.begin(),
      );
      if (
        !isMounted.current ||
        requestId !== requestCounter.current ||
        stamp !== latestInputStamp.current
      )
        return;
      if (importResult.workbook.id === props.workbook.id)
        throw Error(
          "This is the current workbook identity. Use the current workbook source instead.",
        );
      setPendingImport({ ...importResult, stamp: stamp });
    } catch (error) {
      isMounted.current &&
        requestId === requestCounter.current &&
        stamp === latestInputStamp.current &&
        setErrorMessage(error.message);
    } finally {
      isMounted.current &&
        requestId === requestCounter.current &&
        setIsInspecting(false);
    }
  }
  function updateApprovedWorkbooks(nextApprovedWorkbooks) {
    (props.onApprovedWorkbooksChange
      ? props.onApprovedWorkbooksChange(nextApprovedWorkbooks)
      : setLocalApprovedWorkbooks(nextApprovedWorkbooks),
      setPreview(null));
  }
  const X$1 = review
      ? reviewTab === "contributions"
        ? review.contributions
        : reviewTab === "subtotals"
          ? review.subtotals
          : reviewTab === "sources"
            ? review.sources
            : review.exceptions
      : [],
    pageCount = Math.max(1, Math.ceil(X$1.length / 25)),
    pageIndex = Math.min(requestedPage, pageCount - 1);
  return i.jsx("div", {
    className: "lg-modal-backdrop lg-consolidation-backdrop",
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "lg-consolidation",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": "consolidation-title",
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.stopPropagation(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
              ...dialogRef.current.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),select:not(:disabled),a[href]",
              ),
            ].filter((element) => !element.hasAttribute("hidden")),
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
                    children: "Data · Consolidation",
                  }),
                }),
                i.jsx("h2", {
                  id: "consolidation-title",
                  children: i.jsx(LocalizedText, {
                    children: "Consolidate schedules & subtotals",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Align categories, verify source context, then publish a grouped schedule.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close consolidation"),
              onClick: props.onClose,
              lang: language("Close consolidation"),
              "data-ui-label": "Close consolidation",
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-consolidation-body",
          children: [
            i.jsxs("section", {
              className: "lg-consolidation-settings",
              "aria-label": text("Consolidation settings"),
              lang: language("Consolidation settings"),
              "data-ui-label": "Consolidation settings",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Output worksheet" }),
                    i.jsx("input", {
                      "aria-label": text("Output worksheet"),
                      value: outputName,
                      maxLength: 31,
                      onChange: (event) => {
                        (setOutputName(event.target.value), setPreview(null));
                      },
                      lang: language("Output worksheet"),
                      "data-ui-label": "Output worksheet",
                    }),
                  ],
                }),
                ["period", "currency", "unit"].map((fieldName) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx(LocalizedText, { children: "Target " }),
                        fieldName,
                        i.jsx("input", {
                          "aria-label": `Target ${fieldName}`,
                          value: definition[fieldName],
                          onChange: (event) =>
                            updateDefinition({
                              ...definition,
                              [fieldName]: event.target.value,
                            }),
                        }),
                      ],
                    },
                    fieldName,
                  ),
                ),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Category levels" }),
                    i.jsx("select", {
                      "aria-label": text("Category levels"),
                      value: definition.categoryLabels.length,
                      onChange: (event) => {
                        const levelCount = Number(event.target.value),
                          categoryLabels = Array.from(
                            { length: levelCount },
                            (oe, newLabelIndex) =>
                              definition.categoryLabels[newLabelIndex] ??
                              `Category ${newLabelIndex + 1}`,
                          );
                        updateDefinition({
                          ...definition,
                          categoryLabels: categoryLabels,
                          sources: definition.sources.map((source) => ({
                            ...source,
                            categories: categoryLabels.map(
                              (categoryLabel, categorySlot) =>
                                source.categories[categorySlot] ?? "",
                            ),
                          })),
                        });
                      },
                      lang: language("Category levels"),
                      "data-ui-label": "Category levels",
                      children: [1, 2, 3].map((levelOption) =>
                        i.jsx("option", { children: levelOption }, levelOption),
                      ),
                    }),
                  ],
                }),
                definition.categoryLabels.map((categoryLabel, labelIndex) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx(LocalizedText, { children: "Category " }),
                        labelIndex + 1,
                        " ",
                        i.jsx(LocalizedText, { children: "label" }),
                        i.jsx("input", {
                          "aria-label": `Category ${labelIndex + 1} label`,
                          value: categoryLabel,
                          onChange: (event) =>
                            updateDefinition({
                              ...definition,
                              categoryLabels: definition.categoryLabels.map(
                                (categoryLabel, mappedLabelIndex) =>
                                  mappedLabelIndex === labelIndex
                                    ? event.target.value
                                    : categoryLabel,
                              ),
                            }),
                        }),
                      ],
                    },
                    labelIndex,
                  ),
                ),
              ],
            }),
            i.jsxs("div", {
              className: "lg-consolidation-options",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: definition.linkedAmounts,
                      onChange: (event) =>
                        updateDefinition({
                          ...definition,
                          linkedAmounts: event.target.checked,
                        }),
                    }),
                    i.jsx(LocalizedText, {
                      children: "Link amounts from this workbook",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: definition.collapseDetails,
                      onChange: (event) =>
                        updateDefinition({
                          ...definition,
                          collapseDetails: event.target.checked,
                        }),
                    }),
                    i.jsx(LocalizedText, {
                      children: "Collapse detail groups",
                    }),
                  ],
                }),
                props.allowHidden &&
                  i.jsxs("label", {
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: checked,
                        onChange: (event) => {
                          (setIncludeHidden(event.target.checked),
                            setPreview(null));
                        },
                      }),
                      i.jsx(LocalizedText, {
                        children: "Include authorized hidden source cells",
                      }),
                    ],
                  }),
              ],
            }),
            i.jsx("p", {
              className: "lg-consolidation-hint",
              children: i.jsx(LocalizedText, {
                children:
                  "Periods, currencies and units must match exactly. Amounts are summed; no implicit conversion. Same-workbook links recalculate amounts; changed source context still requires review. Approved external workbooks use refreshable snapshots.",
              }),
            }),
            i.jsxs("section", {
              "aria-label": text("Consolidation sources"),
              lang: language("Consolidation sources"),
              "data-ui-label": "Consolidation sources",
              children: [
                definition.sources.map((source, sourceIndex) => {
                  const sourceApproval = workbookSources.find(
                      (approvalEntry) =>
                        approvalEntry.workbook.id === source.workbookId,
                    ),
                    workbook = sourceApproval?.workbook,
                    selectableSheets =
                      workbook?.sheets.filter(
                        (sheet) =>
                          sourceApproval.allowedSheetIds.includes(sheet.id) &&
                          ((!sheet.hidden && !sheet.veryHidden) ||
                            (checked && sourceApproval.includeHidden)),
                      ) ?? [],
                    selectedSheet = selectableSheets.find(
                      (candidateSheet) => candidateSheet.id === source.sheetId,
                    ),
                    sourceFields = selectedSheet
                      ? findTableRangesForSource(workbook, source)
                      : [],
                    renderFieldSelect = (
                      fieldLabel,
                      fieldValue,
                      onFieldChange,
                    ) =>
                      i.jsxs("label", {
                        children: [
                          fieldLabel,
                          i.jsxs("select", {
                            "aria-label": `Source ${sourceIndex + 1} ${fieldLabel}`,
                            value: fieldValue,
                            onChange: (event) =>
                              onFieldChange(event.target.value),
                            children: [
                              !sourceFields.includes(fieldValue) &&
                                i.jsx("option", {
                                  value: fieldValue,
                                  children: fieldValue || "Choose field",
                                }),
                              sourceFields.map((sourceField) =>
                                i.jsx(
                                  "option",
                                  { value: sourceField, children: sourceField },
                                  sourceField,
                                ),
                              ),
                            ],
                          }),
                        ],
                      });
                  return i.jsxs(
                    "article",
                    {
                      className: "lg-consolidation-source",
                      children: [
                        i.jsxs("div", {
                          className: "lg-consolidation-source-title",
                          children: [
                            i.jsxs("strong", {
                              children: [
                                i.jsx(LocalizedText, { children: "Source " }),
                                sourceIndex + 1,
                              ],
                            }),
                            i.jsxs("button", {
                              disabled: definition.sources.length === 1,
                              onClick: () =>
                                updateDefinition({
                                  ...definition,
                                  sources: definition.sources.filter(
                                    (source, candidateIndex) =>
                                      sourceIndex !== candidateIndex,
                                  ),
                                }),
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Remove source ",
                                }),
                                sourceIndex + 1,
                              ],
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "lg-consolidation-fields",
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, { children: "Workbook" }),
                                i.jsxs("select", {
                                  "aria-label": `Source ${sourceIndex + 1} Workbook`,
                                  value: source.workbookId,
                                  onChange: (event) => {
                                    const chosenApproval = workbookSources.find(
                                        (approvalMatch) =>
                                          approvalMatch.workbook.id ===
                                          event.target.value,
                                      ),
                                      sheet =
                                        chosenApproval.workbook.sheets.find(
                                          (sheet) =>
                                            chosenApproval.allowedSheetIds.includes(
                                              sheet.id,
                                            ) &&
                                            !sheet.hidden &&
                                            !sheet.veryHidden,
                                        );
                                    sheet &&
                                      updateSourceAt(sourceIndex, {
                                        ...createConsolidationDefinition(
                                          chosenApproval.workbook,
                                          sheet.id,
                                          formatOccupiedRangeAddress(sheet),
                                          definition,
                                        ),
                                        id: source.id,
                                      });
                                  },
                                  children: [
                                    !sourceApproval &&
                                      i.jsx("option", {
                                        value: source.workbookId,
                                        children:
                                          "Missing or unapproved workbook",
                                      }),
                                    workbookSources.map((approvalOption) =>
                                      i.jsxs(
                                        "option",
                                        {
                                          value: approvalOption.workbook.id,
                                          children: [
                                            approvalOption.workbook.name,
                                            approvalOption.workbook.id ===
                                            props.workbook.id
                                              ? " (current)"
                                              : " (approved)",
                                          ],
                                        },
                                        approvalOption.workbook.id,
                                      ),
                                    ),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, { children: "Worksheet" }),
                                i.jsxs("select", {
                                  "aria-label": `Source ${sourceIndex + 1} Worksheet`,
                                  value: source.sheetId,
                                  onChange: (event) => {
                                    const chosenSheet = selectableSheets.find(
                                      (matchingSheet) =>
                                        matchingSheet.id === event.target.value,
                                    );
                                    updateSourceAt(sourceIndex, {
                                      ...createConsolidationDefinition(
                                        workbook,
                                        chosenSheet.id,
                                        formatOccupiedRangeAddress(chosenSheet),
                                        definition,
                                      ),
                                      id: source.id,
                                    });
                                  },
                                  children: [
                                    !selectedSheet &&
                                      i.jsx("option", {
                                        value: source.sheetId,
                                        children:
                                          "Missing or unauthorized worksheet",
                                      }),
                                    selectableSheets.map((sheetOption) =>
                                      i.jsx(
                                        "option",
                                        {
                                          value: sheetOption.id,
                                          children: sheetOption.name,
                                        },
                                        sheetOption.id,
                                      ),
                                    ),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Source type",
                                }),
                                i.jsxs("select", {
                                  "aria-label": `Source ${sourceIndex + 1} Source type`,
                                  value: source.tableId ?? "",
                                  onChange: (event) => {
                                    const { tableId, ...sourceWithoutTable } =
                                      source;
                                    updateSourceAt(sourceIndex, {
                                      ...sourceWithoutTable,
                                      ...(event.target.value
                                        ? { tableId: event.target.value }
                                        : {}),
                                    });
                                  },
                                  children: [
                                    i.jsx("option", {
                                      value: "",
                                      children: "Explicit range",
                                    }),
                                    source.tableId &&
                                      !selectedSheet?.tables?.some(
                                        (table) => table.id === source.tableId,
                                      ) &&
                                      i.jsx("option", {
                                        value: source.tableId,
                                        children: "Missing table",
                                      }),
                                    selectedSheet?.tables?.map((table) =>
                                      i.jsx(
                                        "option",
                                        {
                                          value: table.id,
                                          children: table.name,
                                        },
                                        table.id,
                                      ),
                                    ),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Header & data range",
                                }),
                                i.jsx("input", {
                                  "aria-label": `Source ${sourceIndex + 1} Header & data range`,
                                  value: source.range,
                                  disabled: !!source.tableId,
                                  onChange: (event) =>
                                    updateSourceAt(sourceIndex, {
                                      ...source,
                                      range: event.target.value,
                                    }),
                                }),
                              ],
                            }),
                            definition.categoryLabels.map(
                              (categoryLabel, categoryIndex) =>
                                i.jsx(
                                  "div",
                                  {
                                    children: renderFieldSelect(
                                      `Category ${categoryIndex + 1}`,
                                      source.categories[categoryIndex],
                                      (categoryValue) =>
                                        updateSourceAt(sourceIndex, {
                                          ...source,
                                          categories: source.categories.map(
                                            (category, categoryPosition) =>
                                              categoryPosition === categoryIndex
                                                ? categoryValue
                                                : category,
                                          ),
                                        }),
                                    ),
                                  },
                                  categoryIndex,
                                ),
                            ),
                            renderFieldSelect(
                              "Amount",
                              source.amountField,
                              (amountField) =>
                                updateSourceAt(sourceIndex, {
                                  ...source,
                                  amountField: amountField,
                                }),
                            ),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Repeated categories",
                                }),
                                i.jsxs("select", {
                                  "aria-label": `Source ${sourceIndex + 1} Repeated categories`,
                                  value: source.duplicatePolicy,
                                  onChange: (event) =>
                                    updateSourceAt(sourceIndex, {
                                      ...source,
                                      duplicatePolicy: event.target.value,
                                    }),
                                  children: [
                                    i.jsx("option", {
                                      value: "reject",
                                      children: "Reject duplicates",
                                    }),
                                    i.jsx("option", {
                                      value: "sum",
                                      children: "Sum repeated categories",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                            ["period", "currency", "unit"].map((contextField) =>
                              i.jsxs(
                                "div",
                                {
                                  className: "lg-consolidation-context",
                                  children: [
                                    i.jsxs("label", {
                                      children: [
                                        contextField,
                                        " ",
                                        i.jsx(LocalizedText, {
                                          children: "source",
                                        }),
                                        i.jsxs("select", {
                                          "aria-label": `Source ${sourceIndex + 1} ${contextField} source`,
                                          value: source[contextField].kind,
                                          onChange: (event) =>
                                            updateSourceAt(sourceIndex, {
                                              ...source,
                                              [contextField]:
                                                event.target.value ===
                                                "constant"
                                                  ? {
                                                      kind: "constant",
                                                      value:
                                                        definition[
                                                          contextField
                                                        ],
                                                    }
                                                  : {
                                                      kind: "field",
                                                      field:
                                                        sourceFields[0] ?? "",
                                                    },
                                            }),
                                          children: [
                                            i.jsx("option", {
                                              value: "constant",
                                              children: "Declared constant",
                                            }),
                                            i.jsx("option", {
                                              value: "field",
                                              children: "Source column",
                                            }),
                                          ],
                                        }),
                                      ],
                                    }),
                                    source[contextField].kind === "field"
                                      ? renderFieldSelect(
                                          contextField,
                                          source[contextField].field,
                                          (field) =>
                                            updateSourceAt(sourceIndex, {
                                              ...source,
                                              [contextField]: {
                                                kind: "field",
                                                field: field,
                                              },
                                            }),
                                        )
                                      : i.jsxs("label", {
                                          children: [
                                            contextField,
                                            i.jsx("input", {
                                              "aria-label": `Source ${sourceIndex + 1} ${contextField}`,
                                              value: source[contextField].value,
                                              onChange: (event) =>
                                                updateSourceAt(sourceIndex, {
                                                  ...source,
                                                  [contextField]: {
                                                    kind: "constant",
                                                    value: event.target.value,
                                                  },
                                                }),
                                            }),
                                          ],
                                        }),
                                  ],
                                },
                                contextField,
                              ),
                            ),
                          ],
                        }),
                      ],
                    },
                    source.id,
                  );
                }),
                i.jsx("button", {
                  disabled: definition.sources.length >= 8,
                  onClick: () => {
                    const sheet = props.workbook.sheets.find(
                      (sheet) =>
                        allowedSheetIds.includes(sheet.id) &&
                        !sheet.hidden &&
                        !sheet.veryHidden,
                    );
                    sheet &&
                      updateDefinition({
                        ...definition,
                        sources: [
                          ...definition.sources,
                          createConsolidationDefinition(
                            props.workbook,
                            sheet.id,
                            formatOccupiedRangeAddress(sheet),
                            definition,
                          ),
                        ],
                      });
                  },
                  children: i.jsx(LocalizedText, { children: "Add source" }),
                }),
              ],
            }),
            props.allowFileSources &&
              i.jsxs("details", {
                className: "lg-consolidation-import",
                children: [
                  i.jsx("summary", {
                    children: i.jsx(LocalizedText, {
                      children: "Approved workbook files",
                    }),
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Choose a local native JSON or XLSX workbook, review the import, then explicitly approve its visible worksheets. Files stay in this session; nothing is fetched automatically.",
                    }),
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Inspect workbook file",
                      }),
                      i.jsx("input", {
                        type: "file",
                        accept: ".json,.xlsx",
                        "aria-label": text("Inspect consolidation workbook"),
                        disabled: isInspecting || approvedWorkbooks.length >= 8,
                        onChange: (event) => {
                          (inspectWorkbookFile(event.target.files?.[0]),
                            (event.target.value = ""));
                        },
                        lang: language("Inspect consolidation workbook"),
                        "data-ui-label": "Inspect consolidation workbook",
                      }),
                    ],
                  }),
                  isInspecting &&
                    i.jsx("p", {
                      role: "status",
                      children: i.jsx(LocalizedText, {
                        children: "Inspecting workbook…",
                      }),
                    }),
                  pendingImport?.stamp === inputStamp &&
                    i.jsxs("div", {
                      "aria-label": text("Pending workbook approval"),
                      lang: language("Pending workbook approval"),
                      "data-ui-label": "Pending workbook approval",
                      children: [
                        i.jsx("strong", {
                          children: pendingImport.workbook.name,
                        }),
                        i.jsx("p", {
                          children: pendingImport.workbook.sheets
                            .filter(
                              (sheet) => !sheet.hidden && !sheet.veryHidden,
                            )
                            .map((visibleSheet) => visibleSheet.name)
                            .join(", "),
                        }),
                        pendingImport.warnings.map((warning, warningIndex) =>
                          i.jsx("p", { children: warning }, warningIndex),
                        ),
                        i.jsx("button", {
                          onClick: () => {
                            (updateApprovedWorkbooks([
                              ...approvedWorkbooks.filter(
                                (approvedWorkbook) =>
                                  approvedWorkbook.workbook.id !==
                                  pendingImport.workbook.id,
                              ),
                              {
                                workbook: pendingImport.workbook,
                                allowedSheetIds: pendingImport.workbook.sheets
                                  .filter(
                                    (sheet) =>
                                      !sheet.hidden && !sheet.veryHidden,
                                  )
                                  .map((allowedSheetId) => allowedSheetId.id),
                                includeHidden: false,
                              },
                            ]),
                              setPendingImport(null));
                          },
                          children: i.jsx(LocalizedText, {
                            children: "Approve workbook",
                          }),
                        }),
                        i.jsx("button", {
                          onClick: () => setPendingImport(null),
                          children: i.jsx(LocalizedText, {
                            children: "Discard file",
                          }),
                        }),
                      ],
                    }),
                  approvedWorkbooks.map((approvedEntry) =>
                    i.jsxs(
                      "p",
                      {
                        children: [
                          approvedEntry.workbook.name,
                          " · ",
                          approvedEntry.allowedSheetIds.length,
                          " ",
                          i.jsx(LocalizedText, { children: "approved sheets" }),
                          " ",
                          i.jsxs("button", {
                            disabled:
                              !props.onApprovedWorkbooksChange &&
                              !!props.approvedWorkbooks?.some(
                                (approvedWorkbook) =>
                                  approvedWorkbook.workbook.id ===
                                  approvedEntry.workbook.id,
                              ),
                            onClick: () =>
                              updateApprovedWorkbooks(
                                approvedWorkbooks.filter(
                                  (approvedWorkbook) =>
                                    approvedWorkbook.workbook.id !==
                                    approvedEntry.workbook.id,
                                ),
                              ),
                            children: [
                              i.jsx(LocalizedText, { children: "Revoke " }),
                              approvedEntry.workbook.name,
                            ],
                          }),
                        ],
                      },
                      approvedEntry.workbook.id,
                    ),
                  ),
                ],
              }),
            errorMessage &&
              i.jsx("p", {
                role: "alert",
                className: "lg-consolidation-error",
                children: errorMessage,
              }),
            review &&
              i.jsxs("section", {
                ref: previewRef,
                "aria-label": text("Consolidation preview"),
                className: "lg-consolidation-preview",
                lang: language("Consolidation preview"),
                "data-ui-label": "Consolidation preview",
                children: [
                  i.jsx("h3", {
                    children: review.complete
                      ? "Ready for review"
                      : "Resolve exceptions before publishing",
                  }),
                  i.jsxs("div", {
                    className: "lg-consolidation-totals",
                    children: [
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Contributions " }),
                          i.jsx("strong", {
                            children: review.contributions.length,
                          }),
                        ],
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Grand total " }),
                          i.jsx("strong", {
                            children: review.grandTotal ?? "Blocked",
                          }),
                        ],
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Source controls" }),
                          " ",
                          i.jsx("strong", {
                            children: review.sourceTotal ?? "Blocked",
                          }),
                        ],
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Difference " }),
                          i.jsx("strong", {
                            children: review.difference ?? "—",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      definition.period,
                      " · ",
                      definition.currency,
                      " · ",
                      definition.unit,
                      ".",
                      " ",
                      review.exceptions.length,
                      " ",
                      i.jsx(LocalizedText, { children: "exceptions." }),
                      " ",
                      review.complete
                        ? "Review all contributions and source controls before publishing."
                        : "No totals or partial output can be published while errors remain.",
                    ],
                  }),
                  i.jsxs("div", {
                    className: "lg-consolidation-tabs",
                    children: [
                      [
                        "contributions",
                        "subtotals",
                        "sources",
                        "exceptions",
                      ].map((tabName) =>
                        i.jsx(
                          "button",
                          {
                            "aria-pressed": reviewTab === tabName,
                            onClick: () => {
                              (setReviewTab(tabName), setRequestedPage(0));
                            },
                            children:
                              tabName[0].toUpperCase() + tabName.slice(1),
                          },
                          tabName,
                        ),
                      ),
                      i.jsx("button", {
                        onClick: () =>
                          saveBlobToFile(
                            new Blob(
                              [
                                JSON.stringify(
                                  {
                                    definition: preview.plan.request.definition,
                                    review: review,
                                  },
                                  null,
                                  2,
                                ),
                              ],
                              { type: "application/json" },
                            ),
                            "Consolidation-review.json",
                          ),
                        children: i.jsx(LocalizedText, {
                          children: "Save review",
                        }),
                      }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "lg-consolidation-table",
                    children: [
                      i.jsxs("table", {
                        "aria-label": `Consolidation ${reviewTab}`,
                        children: [
                          i.jsx("thead", {
                            children: i.jsx("tr", {
                              children: (reviewTab === "contributions"
                                ? [
                                    "Category path",
                                    "Period / currency / unit",
                                    "Amount",
                                    "Source amount",
                                    "Link",
                                  ]
                                : reviewTab === "subtotals"
                                  ? [
                                      "Category path",
                                      "Level",
                                      "Contributions",
                                      "Subtotal",
                                    ]
                                  : reviewTab === "sources"
                                    ? [
                                        "Source",
                                        "Range",
                                        "Records / accepted",
                                        "Context",
                                        "Amount",
                                        "Fingerprint",
                                      ]
                                    : ["Source / row", "Severity", "Exception"]
                              ).map((headerLabel) =>
                                i.jsx(
                                  "th",
                                  { children: headerLabel },
                                  headerLabel,
                                ),
                              ),
                            }),
                          }),
                          i.jsxs("tbody", {
                            children: [
                              reviewTab === "contributions" &&
                                review.contributions
                                  .slice(pageIndex * 25, pageIndex * 25 + 25)
                                  .map((contribution, contributionRowIndex) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsx("td", {
                                            children: contribution.categories
                                              .map(
                                                (category) =>
                                                  `${typeof category}: ${String(category)}`,
                                              )
                                              .join(" / "),
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              contribution.period,
                                              " / ",
                                              contribution.currency,
                                              " / ",
                                              contribution.unit,
                                            ],
                                          }),
                                          i.jsx("td", {
                                            children: contribution.amount,
                                          }),
                                          i.jsx("td", {
                                            children: contribution.locator,
                                          }),
                                          i.jsx("td", {
                                            children: contribution.link,
                                          }),
                                        ],
                                      },
                                      contributionRowIndex,
                                    ),
                                  ),
                              reviewTab === "subtotals" &&
                                review.subtotals
                                  .slice(pageIndex * 25, pageIndex * 25 + 25)
                                  .map((subtotal, subtotalRowIndex) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsx("td", {
                                            children: subtotal.categories
                                              .map(String)
                                              .join(" / "),
                                          }),
                                          i.jsx("td", {
                                            children: subtotal.level,
                                          }),
                                          i.jsx("td", {
                                            children: subtotal.count,
                                          }),
                                          i.jsx("td", {
                                            children: subtotal.amount,
                                          }),
                                        ],
                                      },
                                      subtotalRowIndex,
                                    ),
                                  ),
                              reviewTab === "sources" &&
                                review.sources
                                  .slice(pageIndex * 25, pageIndex * 25 + 25)
                                  .map((sourceSummary) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [
                                              sourceSummary.workbookName,
                                              " / ",
                                              sourceSummary.sheetName,
                                            ],
                                          }),
                                          i.jsx("td", {
                                            children: sourceSummary.range,
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              sourceSummary.records,
                                              " / ",
                                              sourceSummary.acceptedRecords,
                                            ],
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              sourceSummary.periods.join(", "),
                                              " / ",
                                              sourceSummary.currencies.join(
                                                ", ",
                                              ),
                                              " ",
                                              "/ ",
                                              sourceSummary.units.join(", "),
                                            ],
                                          }),
                                          i.jsx("td", {
                                            children: sourceSummary.amount,
                                          }),
                                          i.jsx("td", {
                                            children: i.jsx("code", {
                                              children:
                                                sourceSummary.fingerprint,
                                            }),
                                          }),
                                        ],
                                      },
                                      sourceSummary.sourceId,
                                    ),
                                  ),
                              reviewTab === "exceptions" &&
                                review.exceptions
                                  .slice(pageIndex * 25, pageIndex * 25 + 25)
                                  .map((exception, exceptionRowIndex) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [
                                              i.jsx(LocalizedText, {
                                                children: "Source",
                                              }),
                                              " ",
                                              definition.sources.findIndex(
                                                (source) =>
                                                  source.id ===
                                                  exception.sourceId,
                                              ) + 1,
                                              exception.row
                                                ? ` / row ${exception.row}`
                                                : "",
                                            ],
                                          }),
                                          i.jsx("td", {
                                            children: exception.severity,
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              exception.code,
                                              ": ",
                                              exception.message,
                                            ],
                                          }),
                                        ],
                                      },
                                      exceptionRowIndex,
                                    ),
                                  ),
                            ],
                          }),
                        ],
                      }),
                      !X$1.length &&
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, { children: "No " }),
                            reviewTab,
                            ".",
                          ],
                        }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "lg-consolidation-pages",
                    children: [
                      i.jsx("button", {
                        disabled: pageIndex === 0,
                        onClick: () => setRequestedPage(pageIndex - 1),
                        children: i.jsx(LocalizedText, {
                          children: "Previous page",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Page " }),
                          pageIndex + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          pageCount,
                          " · ",
                          X$1.length,
                          " ",
                          i.jsx(LocalizedText, { children: "records" }),
                        ],
                      }),
                      i.jsx("button", {
                        disabled: pageIndex + 1 >= pageCount,
                        onClick: () => setRequestedPage(pageIndex + 1),
                        children: i.jsx(LocalizedText, {
                          children: "Next page",
                        }),
                      }),
                    ],
                  }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children: props.readOnly
                ? "Read-only: preview available"
                : "Separate output · originals retained · Undo / Redo",
            }),
            i.jsx("button", {
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              onClick: () =>
                runGuarded(() => {
                  const consolidationPreview = createConsolidationPreview(
                    props.workbook,
                    { definition: definition, outputName: outputName },
                    options,
                  );
                  (setPreview({ plan: consolidationPreview, stamp: stamp }),
                    setRequestedPage(0),
                    setReviewTab(
                      consolidationPreview.review.complete
                        ? "contributions"
                        : "exceptions",
                    ));
                }),
              children: i.jsx(LocalizedText, {
                children: "Preview consolidation",
              }),
            }),
            i.jsx("button", {
              className: "primary",
              disabled:
                !review?.complete ||
                props.readOnly ||
                !!props.workbook.structureProtection?.enabled,
              onClick: () =>
                runGuarded(() => {
                  if (!isPreviewCurrent) return;
                  const publishedWorkbook = publishConsolidatedSchedule(
                    props.workbook,
                    preview.plan,
                    options,
                  );
                  (setPreview(null),
                    props.onWorkbookChange(
                      publishedWorkbook,
                      "Publish consolidated schedule",
                    ),
                    props.onClose());
                }),
              children: i.jsx(LocalizedText, {
                children: "Publish consolidated schedule",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
const caseFoldText = (value) => value.toLowerCase(),
  buildTypedValueKey = (value) =>
    JSON.stringify([
      typeof value,
      typeof value == "string" ? caseFoldText(value) : value,
    ]),
  buildTableColumnKey = (tableColumn) =>
    `${tableColumn.tableId}\0${tableColumn.column}`,
  requireScalarValue = (value) =>
    typeof value == "object" && value !== null
      ? throwDaxError("DAX_TYPE", "A table cannot be used as a scalar.")
      : value,
  requireLogicalOperand = (value) =>
    value === null
      ? false
      : typeof value == "boolean"
        ? value
        : typeof value == "number"
          ? value !== 0
          : throwDaxError(
              "DAX_TYPE",
              "A logical condition requires a boolean, number or BLANK.",
            ),
  requireNumericOperand = (value) =>
    value === null
      ? 0
      : typeof value == "number"
        ? value
        : throwDaxError(
            "DAX_TYPE",
            "Numeric expressions require numbers or BLANK; implicit text/date conversion is unsupported.",
          ),
  requireFiniteNumber = (value) =>
    Number.isFinite(value)
      ? value
      : throwDaxError(
          "DAX_NUMBER",
          "The result is not finite. Use DIVIDE for a blank or zero denominator.",
        );
function compareDaxValues(left, right, shouldTreatBlankAsUnknown = false) {
  if (shouldTreatBlankAsUnknown && (left === null || right === null))
    return left === right ? 0 : undefined;
  if (left === null && right === null) return 0;
  if (
    (left === null &&
      (left =
        typeof right == "string" ? "" : typeof right == "boolean" ? false : 0),
    right === null &&
      (right =
        typeof left == "string" ? "" : typeof left == "boolean" ? false : 0),
    typeof left != typeof right)
  )
    return throwDaxError(
      "DAX_TYPE",
      "Comparison operands must have compatible types.",
    );
  const comparableLeft = typeof left == "string" ? caseFoldText(left) : left,
    comparableRight = typeof right == "string" ? caseFoldText(right) : right;
  return comparableLeft === comparableRight
    ? 0
    : comparableLeft < comparableRight
      ? -1
      : 1;
}
function sumWithCompensation(values) {
  let runningSum = 0,
    compensation = 0,
    numericCount = 0;
  for (const value of values) {
    if (typeof value != "number") continue;
    const nextSum = runningSum + value;
    ((compensation +=
      Math.abs(runningSum) >= Math.abs(value)
        ? runningSum - nextSum + value
        : value - nextSum + runningSum),
      (runningSum = nextSum),
      numericCount++);
  }
  return numericCount ? requireFiniteNumber(runningSum + compensation) : null;
}

export {
  requireNumericOperand,
  requireLogicalOperand,
  buildTableColumnKey,
  requireScalarValue,
  requireFiniteNumber,
  ConsolidationDialog,
  ConsolidationStatusBanner,
  sumWithCompensation,
  compareDaxValues,
  caseFoldText,
  buildTypedValueKey,
};
export { saveBlobToFile };
