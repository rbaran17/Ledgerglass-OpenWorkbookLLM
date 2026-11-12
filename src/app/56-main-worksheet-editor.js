import {
  usePersistenceState,
  useDialogState,
  useNavigationState,
  useFormulaEntryState,
  useSelectionState,
  useFilterSortState,
  useAuditState,
  useOperationStatusState,
  useImportExportState,
  useAssistantState,
  useViewDisplayState,
  useFindReplaceState,
  useDragFillState,
  useClipboardState,
} from "./56-main-worksheet-editor/editor-state.js";
import { createViewAndSelectionActions } from "./56-main-worksheet-editor/view-and-selection-actions.js";
import { createPersistenceActions } from "./56-main-worksheet-editor/persistence-actions.js";
import { createCellEditActions } from "./56-main-worksheet-editor/cell-edit-actions.js";
import { createNavigationActions } from "./56-main-worksheet-editor/navigation-actions.js";
import { createFormulaEntryActions } from "./56-main-worksheet-editor/formula-entry-actions.js";
import { createFileAndAnalysisActions } from "./56-main-worksheet-editor/file-and-analysis-actions.js";
import { createClipboardAndFindActions } from "./56-main-worksheet-editor/clipboard-and-find-actions.js";
import { createRowColumnAndFillActions } from "./56-main-worksheet-editor/row-column-and-fill-actions.js";
import { createGridInteractionActions } from "./56-main-worksheet-editor/grid-interaction-actions.js";
import { createGridPaneRenderer } from "./56-main-worksheet-editor/render-grid-pane.js";
import { EditorWorkbookGrid } from "./56-main-worksheet-editor/editor-workbook-grid.js";
import { EditorSidePanel } from "./56-main-worksheet-editor/editor-side-panel.js";
import { WorksheetGridRow } from "./56-main-worksheet-editor/worksheet-grid-row.js";
import { EditorDialogs } from "./56-main-worksheet-editor/editor-dialogs.js";
import { EditorRibbonBar } from "./56-main-worksheet-editor/editor-ribbon-bar.js";

import { saveBlobToFile } from "./13-consolidation-dialog.js";
import { isFiniteNumber } from "./03-workbook-model.js";
import i from "react/jsx-runtime";
import ne from "react";
import { X } from "lucide-react";

import {
  maxRowHeight,
  minRowHeight,
  maxColumnWidth,
  minColumnWidth,
} from "./02-text-and-financial-functions.js";
import {
  defaultRowHeightOf,
  evaluateConditionalFormatRules,
  enumerateSelectionCells,
  defaultColumnWidthOf,
  Pv,
  requireRetainedXlsxSource,
  _t,
  cn,
  parseRangeAddress,
  applyFilterConditions,
  assertNoStopValidationIssues,
  applyCellPatchTransaction,
  buildCanonicalWorkbookSnapshot,
  createBlankWorkbook,
  createAxisVisibilityIndex,
  columnIndexToLetters,
  vx,
  countSelectionCells,
} from "./03-workbook-model.js";
import {
  WorkbookMessagesProvider,
  useWorkbookMessages,
  withExtensionHostScope,
  LocalizedText,
  formatCellEntryForEdit,
  formatLocalizedNumberDisplay,
} from "./04-cell-input-paste-and-fill.js";
import { CustomFunctionReviewBanner } from "./05-add-ins-and-custom-functions.js";

import { exportWorksheetToCsv } from "./08-xlsx-import-export-and-preservation.js";
import {
  useOperationScope,
  WorkbookLinkStatusBadge,
  operationScopeContext,
} from "./09-workbook-links.js";
import { XlsxConversionReview } from "./10-source-import-dialog.js";
import {
  WorkbookRevisionConflictError,
  createWorkbookSession,
} from "./11-local-workbook-store-and-versions.js";
import { ConsolidationStatusBanner } from "./13-consolidation-dialog.js";
import {
  DataModelSnapshotNotice,
  useDialogFocusTrap,
} from "./15-data-model-dialog.js";

import { ValidationReviewDialog } from "./19-filter-views-dialog.js";
import {
  GridContextMenu,
  validateProposedCalculation,
  buildAssistantSelectionContext,
} from "./20-calculation-options-and-assistant.js";

import { ConnectionStatusNotice } from "./22-themes-and-cell-styles.js";

import { useCalculationRun } from "./24-merge-audit-and-ribbon-helpers.js";

import { AnalysisSnapshotNotice } from "./38-data-analysis-dialog.js";

import { DelimitedImportExportDialog } from "./40-what-if-and-delimited-dialogs.js";

import { FillPreviewDialog } from "./47-fill-and-paste-special.js";
import { MoveCellsDialog } from "./48-move-cells-and-reference-repair.js";

import { buildAnswerCitations } from "./51-pattern-checks-and-citations.js";

import {
  buildVisibleAxisLayout,
  useSplitPaneSizes,
  createTextMeasurer,
  Tp,
  sliceVisibleAxis,
  logicalToPhysicalPixels,
} from "./53-find-replace-go-to.js";
import {
  instrumentWorkbookPersistence,
  WorkbookSaveConflictDialog,
  WorkbookHealthDialog,
  createWorkbookDiagnosticsMonitor,
  saveWorkbookWithLock,
} from "./54-workbook-save-conflicts.js";
import {
  isPromiseLike,
  ExternalSourceStatusPanel,
  workbookInstanceKey,
  Vj,
  wy,
} from "./55-external-source-refresh.js";

function WorkbookSyncBoundary(props) {
  const diagnosticsMonitor = ne.useMemo(
      () => createWorkbookDiagnosticsMonitor(),
      [],
    ),
    reliability = props.reliability ?? diagnosticsMonitor,
    instanceRef = ne.useRef({
      propId: props.initialWorkbook?.id,
      emittedId: undefined,
      generation: 0,
    });
  instanceRef.current.propId !== props.initialWorkbook?.id &&
    (instanceRef.current.emittedId !== props.initialWorkbook?.id &&
      instanceRef.current.generation++,
    (instanceRef.current.propId = props.initialWorkbook?.id));
  const onChange = ne.useCallback(
      (workbook) => {
        ((instanceRef.current.emittedId = workbook.id),
          props.onChange?.(workbook));
      },
      [props.onChange],
    ),
    onWorkbookLoaded = ne.useCallback(
      (workbook) => {
        ((instanceRef.current.emittedId = workbook.id),
          props.onWorkbookLoaded?.(workbook));
      },
      [props.onWorkbookLoaded],
    );
  return i.jsx(WorkbookMessagesProvider, {
    options: props.localization,
    children: i.jsx(operationScopeContext.Provider, {
      value: { ...props.fileProcessingOptions, reliability: reliability },
      children: i.jsx(
        MainWorksheetEditor,
        {
          ...props,
          reliability: reliability,
          onChange: onChange,
          onWorkbookLoaded: onWorkbookLoaded,
        },
        (props.persistenceKey ?? "session") +
          ":" +
          instanceRef.current.generation +
          ":" +
          workbookInstanceKey(props.workbookStore),
      ),
    }),
  });
}
function MainWorksheetEditor({
  extensions,
  linkedWorkbooks,
  onLinkedWorkbooksChange,
  allowLinkedWorkbookFiles = true,
  documentExtractor,
  externalSourceAdapter,
  reliability,
  saveWarningMs = 15e3,
  consolidationWorkbooks,
  consolidationDefinition,
  onConsolidationWorkbooksChange,
  allowConsolidationFiles = true,
  calculationWorkerFactory,
  initialWorkbook,
  onChange,
  onBeforeChange,
  onCellEditPendingChange,
  onWorkbookLoaded,
  onSelectionChange,
  adapter = Vj,
  actor = "Preparer",
  readOnly = false,
  persistenceKey,
  workbookStore,
  workbookStorePersistent = false,
  onPersistenceStateChange,
  externalProposal,
  onExternalProposalStaged,
}) {
  const defaultDiagnosticsMonitor = ne.useMemo(
      () => createWorkbookDiagnosticsMonitor(),
      [],
    ),
    activeReliabilityMonitor = reliability ?? defaultDiagnosticsMonitor,
    dialogState = useDialogState(),
    saveWarningDelayMs =
      Number.isFinite(saveWarningMs) &&
      saveWarningMs >= 1 &&
      saveWarningMs <= 12e4
        ? saveWarningMs
        : 15e3,
    { text, language } = useWorkbookMessages(),
    workbookMessages = useWorkbookMessages(),
    getEditableCellText = (cell) =>
      formatCellEntryForEdit(cell, workbookMessages.entryLocale),
    editorIdPrefix = ne.useId(),
    [workbookSnapshot, setWorkbookSnapshot] = ne.useState(() =>
      buildCanonicalWorkbookSnapshot(initialWorkbook ?? createBlankWorkbook()),
    ),
    activeWorkbook = ne.useMemo(
      () =>
        workbookSnapshot.extensions || extensions?.length
          ? withExtensionHostScope(workbookSnapshot, extensions)
          : workbookSnapshot,
      [workbookSnapshot, extensions],
    ),
    [
      uncontrolledConsolidationWorkbooks,
      setUncontrolledConsolidationWorkbooks,
    ] = ne.useState([]),
    [uncontrolledLinkedWorkbooks, setUncontrolledLinkedWorkbooks] = ne.useState(
      [],
    ),
    effectiveLinkedWorkbooks = linkedWorkbooks ?? uncontrolledLinkedWorkbooks,
    setLinkedWorkbooks =
      linkedWorkbooks !== undefined
        ? onLinkedWorkbooksChange
        : setUncontrolledLinkedWorkbooks,
    effectiveConsolidationWorkbooks =
      consolidationWorkbooks ?? uncontrolledConsolidationWorkbooks,
    setConsolidationWorkbooks = consolidationWorkbooks
      ? onConsolidationWorkbooksChange
      : setUncontrolledConsolidationWorkbooks,
    [activeSheetId, setActiveSheetId] = ne.useState(
      activeWorkbook.sheets.find((sheet) => !sheet.hidden).id,
    ),
    activeSheet =
      activeWorkbook.sheets.find(
        (sheet) => sheet.id === activeSheetId && !sheet.hidden,
      ) ?? activeWorkbook.sheets.find((sheet) => !sheet.hidden),
    workbookPersistence = ne.useMemo(
      () =>
        instrumentWorkbookPersistence(
          workbookStore ??
            (persistenceKey
              ? saveWorkbookWithLock(
                  {
                    getItem: (storageKey) => localStorage.getItem(storageKey),
                    setItem: (storageKey, storedValue) =>
                      localStorage.setItem(storageKey, storedValue),
                  },
                  persistenceKey,
                )
              : createWorkbookSession()),
          activeReliabilityMonitor,
        ),
      [workbookStore, persistenceKey, activeReliabilityMonitor],
    ),
    persistenceState = usePersistenceState(),
    isEditableRef = ne.useRef(!readOnly);
  isEditableRef.current = !readOnly;
  const setSaveAbortControllerRef = ne.useRef(onPersistenceStateChange);
  setSaveAbortControllerRef.current = onPersistenceStateChange;
  const isEditingLocked = readOnly || !!activeSheet.protection?.enabled,
    selectionState = useSelectionState();
  useDialogFocusTrap(selectionState.appRootRef);
  const navigationState = useNavigationState(),
    formulaEntryState = useFormulaEntryState();

  const selectionAreas = ne.useMemo(
      () => [
        ...selectionState.completedSelectionAreas,
        selectionState.activeSelectionRange,
      ],
      [
        selectionState.completedSelectionAreas,
        selectionState.activeSelectionRange,
      ],
    ),
    selectionAreaRanges = ne.useMemo(
      () => selectionAreas.map(parseRangeAddress),
      [selectionAreas],
    ),
    selectionAreasKey = selectionAreas.join(", ");

  (ne.useEffect(() => {
    let animationFrameId = 0,
      pointerPosition;
    const autoScrollStep = () => {
        if (
          !formulaEntryState.formulaPointerStateRef.current.dragging ||
          !pointerPosition
        )
          return;
        const gridContainer = selectionState.gridContainerRef.current;
        if (gridContainer) {
          const gridBounds = gridContainer.getBoundingClientRect();
          (pointerPosition.y < gridBounds.top + 28
            ? (gridContainer.scrollTop -= 24)
            : pointerPosition.y > gridBounds.bottom - 28 &&
              (gridContainer.scrollTop += 24),
            pointerPosition.x < gridBounds.left + 28
              ? (gridContainer.scrollLeft -= 36)
              : pointerPosition.x > gridBounds.right - 28 &&
                (gridContainer.scrollLeft += 36));
          const cellElement = document
            .elementFromPoint?.(
              Math.max(
                gridBounds.left + 48,
                Math.min(gridBounds.right - 8, pointerPosition.x),
              ),
              Math.max(
                gridBounds.top + 28,
                Math.min(gridBounds.bottom - 8, pointerPosition.y),
              ),
            )
            ?.closest("[data-address]");
          cellElement &&
            gridContainer.contains(cellElement) &&
            formulaEntryState.formulaPointerMoveRef.current(
              cellElement.dataset.address,
            );
        }
        animationFrameId = requestAnimationFrame(autoScrollStep);
      },
      handlePointerMove = (event) => {
        formulaEntryState.formulaPointerStateRef.current.dragging &&
          ((pointerPosition = { x: event.clientX, y: event.clientY }),
          (animationFrameId ||= requestAnimationFrame(autoScrollStep)));
      },
      handlePointerEnd = () => {
        const dragging =
          formulaEntryState.formulaPointerStateRef.current.dragging;
        ((formulaEntryState.formulaPointerStateRef.current.dragging = false),
          cancelAnimationFrame(animationFrameId),
          (animationFrameId = 0),
          (pointerPosition = undefined),
          dragging && cellEditActions.restoreFormulaEditorFocus());
      };
    return (
      window.addEventListener("pointermove", handlePointerMove),
      window.addEventListener("pointerup", handlePointerEnd),
      window.addEventListener("pointercancel", handlePointerEnd),
      () => {
        (cancelAnimationFrame(animationFrameId),
          window.removeEventListener("pointermove", handlePointerMove),
          window.removeEventListener("pointerup", handlePointerEnd),
          window.removeEventListener("pointercancel", handlePointerEnd));
      }
    );
  }, []),
    ne.useEffect(() => {
      (formulaEntryState.activeCellEditRef.current ||
        cellEditActions.resetFormulaEditSession(),
        onCellEditPendingChange?.(
          !!formulaEntryState.activeCellEditRef.current,
        ));
    }, [
      formulaEntryState.cellEditText,
      formulaEntryState.cellEditState,
      activeWorkbook,
      onCellEditPendingChange,
    ]));
  const filterSortState = useFilterSortState(),
    auditState = useAuditState(),
    filterStateKey = JSON.stringify([activeWorkbook.id, activeSheet.id]),
    activeFilterState = filterSortState.filterStateByScope[filterStateKey],
    viewDisplayState = useViewDisplayState(),
    sheetSplitPanes = activeSheet.splitPanes,
    splitPaneSizesKey = [
      activeWorkbook.id,
      activeSheet.id,
      sheetSplitPanes?.horizontal !== undefined,
      sheetSplitPanes?.vertical !== undefined,
      sheetSplitPanes?.row,
      sheetSplitPanes?.col,
    ].join(":"),
    splitPaneSizes = useSplitPaneSizes(splitPaneSizesKey),
    activePaneLayout = splitPaneSizes.views[splitPaneSizes.activeId],
    activePaneScrollTop = activePaneLayout.top,
    activePaneScrollLeft = activePaneLayout.left;

  const [undoStack, setUndoStack] = ne.useState([]),
    [redoStack, setRedoStack] = ne.useState([]),
    operationStatusState = useOperationStatusState(),
    operationScope = useOperationScope();
  ne.useEffect(
    () => () => {
      (operationStatusState.latestOperationIdRef.current++,
        operationStatusState.activeOperationControllerRef.current?.abort());
    },
    [],
  );
  const importExportState = useImportExportState(),
    assistantState = useAssistantState(),
    findReplaceState = useFindReplaceState(),
    dragFillState = useDragFillState(),
    fillDragStateRef = ne.useRef(dragFillState.fillDragState);
  fillDragStateRef.current = dragFillState.fillDragState;
  const clipboardState = useClipboardState();

  const [, setRenderRefreshCount] = ne.useState(0),
    calculationRun = useCalculationRun(
      activeWorkbook,
      calculationWorkerFactory,
      handleCalculationCompleted,
      activeReliabilityMonitor,
    ),
    calculationState = calculationRun.calculation,
    workbookCalculator = calculationState.calculator,
    filterApplication = ne.useMemo(() => {
      if (!activeFilterState || activeFilterState.snapshot !== activeWorkbook)
        return { result: null, error: "" };
      try {
        return {
          result: applyFilterConditions(
            activeWorkbook,
            activeSheet.id,
            activeFilterState.definition,
            workbookCalculator,
          ),
          error: "",
        };
      } catch (error) {
        return {
          result: null,
          error:
            error instanceof Error
              ? error.message
              : "Filtering is unavailable.",
        };
      }
    }, [activeWorkbook, activeSheet.id, activeFilterState, workbookCalculator]),
    filterResult = filterApplication.result,
    isFilterStale =
      !!activeFilterState && activeFilterState.snapshot !== activeWorkbook,
    hasFrozenRows = !!activeSheet.frozenRows,
    hiddenRowIndexes = ne.useMemo(
      () => filterResult?.hiddenRows ?? cn(activeSheet, "row"),
      [activeSheet, filterResult],
    ),
    hiddenColumnIndexes = ne.useMemo(
      () => cn(activeSheet, "column"),
      [activeSheet],
    ),
    rowVisibilityIndex = ne.useMemo(
      () => createAxisVisibilityIndex(activeSheet.rowCount, hiddenRowIndexes),
      [activeSheet.rowCount, hiddenRowIndexes],
    ),
    columnVisibilityIndex = ne.useMemo(
      () =>
        createAxisVisibilityIndex(activeSheet.colCount, hiddenColumnIndexes),
      [activeSheet.colCount, hiddenColumnIndexes],
    ),
    pinnedRowCount = Math.max(
      activeSheet.frozenRows ?? 0,
      filterSortState.filterValueSearch || filterSortState.activeColumnSort
        ? 1
        : 0,
      filterResult &&
        parseRangeAddress(filterResult.context.range).firstRow === 0
        ? parseRangeAddress(filterResult.context.range).firstRow + 1
        : 0,
    ),
    frozenColumnCount = activeSheet.frozenCols ?? 0;
  ne.useEffect(() => {
    const activeCellPosition = _t(selectionState.activeCellAddress);
    if (
      !rowVisibilityIndex.isHidden(activeCellPosition.row) &&
      !columnVisibilityIndex.isHidden(activeCellPosition.col)
    )
      return;
    const nextVisibleRow =
        rowVisibilityIndex.next(activeCellPosition.row, 1, true) ??
        rowVisibilityIndex.next(activeCellPosition.row, -1, true),
      nextVisibleColumn =
        columnVisibilityIndex.next(activeCellPosition.col, 1, true) ??
        columnVisibilityIndex.next(activeCellPosition.col, -1, true);
    if (
      nextVisibleRow === null ||
      nextVisibleColumn === null ||
      nextVisibleRow === undefined ||
      nextVisibleColumn === undefined
    )
      return;
    const nextActiveAddress = `${columnIndexToLetters(nextVisibleColumn)}${nextVisibleRow + 1}`;
    (selectionState.setActiveCellAddress(nextActiveAddress),
      selectionState.setSelectionAnchorAddress(nextActiveAddress),
      (selectionState.activeCellAddressRef.current = nextActiveAddress),
      viewAndSelectionActions.selectSingleRange(nextActiveAddress),
      navigationState.setPendingScrollAddress(nextActiveAddress));
  }, [
    selectionState.activeCellAddress,
    rowVisibilityIndex,
    columnVisibilityIndex,
  ]);
  const defaultRowHeight = defaultRowHeightOf(activeSheet),
    defaultColumnWidth = defaultColumnWidthOf(activeSheet),
    zoomScale = viewDisplayState.zoomPercent / 100,
    rowHeaderWidth = viewDisplayState.isHeadingsVisible ? 58 : 0,
    textMeasurer = ne.useMemo(() => createTextMeasurer(), []),
    latestEditStateRef = ne.useRef({
      book: activeWorkbook,
      readOnly: isEditingLocked,
      sheetId: activeSheet.id,
      range: selectionState.activeSelectionRange,
      selectionKey: selectionAreasKey,
    });
  latestEditStateRef.current = {
    book: activeWorkbook,
    readOnly: isEditingLocked,
    sheetId: activeSheet.id,
    range: selectionState.activeSelectionRange,
    selectionKey: selectionAreasKey,
  };

  function handleCalculationCompleted() {
    (setRenderRefreshCount((previousCount) => previousCount + 1),
      assistantState.setAssistantProposal(null),
      assistantState.setAnswerCitations(null),
      auditState.setAuditAnalysisResult(null));
  }

  const commentedCellAddresses = ne.useMemo(
      () =>
        new Set(
          (activeWorkbook.reviewThreads ?? [])
            .filter(
              (thread) =>
                thread.anchor.sheetId === activeSheet.id && !thread.orphaned,
            )
            .map((thread) => thread.anchor.address),
        ),
      [activeWorkbook.reviewThreads, activeSheet.id],
    ),
    mergeRanges = ne.useMemo(
      () => (activeSheet.merges ?? []).map(parseRangeAddress),
      [activeSheet.merges],
    ),
    selectionRangeBounds = ne.useMemo(
      () => parseRangeAddress(selectionState.activeSelectionRange),
      [selectionState.activeSelectionRange],
    ),
    selectedCellCount = ne.useMemo(
      () => countSelectionCells(selectionAreas),
      [selectionAreas],
    ),
    assistantSelection = {
      sheetId: activeSheet.id,
      range: selectionState.activeSelectionRange,
    },
    arrayDisplayResult = ne.useMemo(() => {
      try {
        return { cells: vx(activeSheet, workbookCalculator), warning: "" };
      } catch (error) {
        return {
          cells: activeSheet.cells,
          warning:
            error instanceof Error
              ? error.message
              : "Array display exceeds this workbook's calculation budget.",
        };
      }
    }, [activeSheet, workbookCalculator]);
  ne.useEffect(() => {
    arrayDisplayResult.warning &&
      operationStatusState.setErrorNotice(arrayDisplayResult.warning);
  }, [arrayDisplayResult.warning]);
  const conditionalFormatEvaluator = ne.useMemo(
      () =>
        evaluateConditionalFormatRules(
          activeSheet,
          workbookCalculator,
          arrayDisplayResult.cells,
        ),
      [activeSheet, workbookCalculator, arrayDisplayResult],
    ),
    getConditionalRuleStyle = ne.useMemo(
      () => (cellAddress) => conditionalFormatEvaluator(cellAddress).style,
      [conditionalFormatEvaluator],
    ),
    activeCellFormulaErrors = conditionalFormatEvaluator(
      selectionState.activeCellAddress,
    ).errors,
    activeCellReferenceDiagnostics = ne.useMemo(
      () => (
        workbookCalculator(activeSheet.id, selectionState.activeCellAddress),
        workbookCalculator
          .getReferenceDiagnostics()
          .filter(
            (diagnostic) =>
              diagnostic.sheetId === activeSheet.id &&
              diagnostic.address === selectionState.activeCellAddress,
          )
      ),
      [workbookCalculator, activeSheet.id, selectionState.activeCellAddress],
    ),
    activeCellNamedFormulaDiagnostics = ne.useMemo(
      () => (
        workbookCalculator(activeSheet.id, selectionState.activeCellAddress),
        workbookCalculator
          .getNamedFormulaDiagnostics()
          .filter(
            (diagnostic) =>
              diagnostic.sheetId === activeSheet.id &&
              diagnostic.address === selectionState.activeCellAddress,
          )
      ),
      [workbookCalculator, activeSheet.id, selectionState.activeCellAddress],
    ),
    getAllCellEntries = ne.useMemo(() => {
      let cachedCellEntries;
      return () =>
        (cachedCellEntries ??= Object.entries(arrayDisplayResult.cells).map(
          ([address, cell]) => ({
            address: address,
            cell: cell,
            ..._t(address),
          }),
        ));
    }, [arrayDisplayResult.cells]),
    selectedNumericValues = ne.useMemo(
      () =>
        (selectedCellCount <= 500
          ? enumerateSelectionCells(selectionAreas).map((address) => ({
              address: address,
            }))
          : getAllCellEntries().filter(({ row, col }) =>
              selectionAreaRanges.some((selectionAreaRange) =>
                Tp(selectionAreaRange, row, col),
              ),
            )
        )
          .map(({ address }) => workbookCalculator(activeSheet.id, address))
          .filter(isFiniteNumber),
      [
        getAllCellEntries,
        selectionAreaRanges,
        selectionAreas,
        selectedCellCount,
        workbookCalculator,
        activeSheet.id,
      ],
    ),
    selectedNumericSum = selectedNumericValues.reduce(
      (total, numericValue) => total + numericValue,
      0,
    ),
    activeCell = activeSheet.cells[selectionState.activeCellAddress],
    sortedRowIndexes = ne.useMemo(() => {
      if (
        !filterSortState.filterValueSearch &&
        !filterSortState.activeColumnSort
      )
        return null;
      const addressesByRow = new Map();
      for (const { address, row, cell } of getAllCellEntries()) {
        if (
          row < pinnedRowCount ||
          rowVisibilityIndex.isHidden(row) ||
          cell.raw === null
        )
          continue;
        const rowAddresses = addressesByRow.get(row) ?? [];
        (rowAddresses.push(address), addressesByRow.set(row, rowAddresses));
      }
      const normalizedSearch = filterSortState.filterValueSearch.toLowerCase(),
        matchingRowIndexes = [...addressesByRow.keys()]
          .filter(
            (rowIndex) =>
              !normalizedSearch ||
              addressesByRow.get(rowIndex).some((cellAddress) =>
                String(workbookCalculator(activeSheet.id, cellAddress) ?? "")
                  .toLowerCase()
                  .includes(normalizedSearch),
              ),
          )
          .sort((leftRow, rightRow) => leftRow - rightRow);
      return (
        filterSortState.activeColumnSort &&
          matchingRowIndexes.sort((leftRow, rightRow) => {
            const leftValue = workbookCalculator(
                activeSheet.id,
                `${columnIndexToLetters(filterSortState.activeColumnSort.col)}${leftRow + 1}`,
              ),
              rightValue = workbookCalculator(
                activeSheet.id,
                `${columnIndexToLetters(filterSortState.activeColumnSort.col)}${rightRow + 1}`,
              );
            return (
              (isFiniteNumber(leftValue) && isFiniteNumber(rightValue)
                ? leftValue - rightValue
                : String(leftValue ?? "").localeCompare(
                    String(rightValue ?? ""),
                  )) * filterSortState.activeColumnSort.direction
            );
          }),
        matchingRowIndexes
      );
    }, [
      getAllCellEntries,
      workbookCalculator,
      activeSheet.id,
      filterSortState.filterValueSearch,
      filterSortState.activeColumnSort,
      pinnedRowCount,
      rowVisibilityIndex,
    ]),
    filteredRowPositionMap = ne.useMemo(
      () =>
        sortedRowIndexes
          ? new Map(
              sortedRowIndexes.map((sortedRowIndex, sortedPosition) => [
                sortedRowIndex,
                sortedPosition,
              ]),
            )
          : null,
      [sortedRowIndexes],
    );
  ne.useEffect(
    () => viewDisplayState.setAutoFitRowHeights({}),
    [
      activeSheet.id,
      activeSheet.cells,
      activeSheet.columnWidths,
      activeSheet.rowHeights,
      activeSheet.defaultRowHeight,
      activeSheet.defaultColumnWidth,
      viewDisplayState.isShowingFormulas,
    ],
  );
  const effectiveRowHeights = ne.useMemo(
      () => ({
        ...viewDisplayState.bn,
        ...activeSheet.rowHeights,
        ...(viewDisplayState.axisResizeState?.axis === "row"
          ? {
              [viewDisplayState.axisResizeState.index]:
                viewDisplayState.axisResizeState.size,
            }
          : {}),
      }),
      [
        viewDisplayState.bn,
        activeSheet.rowHeights,
        viewDisplayState.axisResizeState,
      ],
    ),
    rowAxisLayout = ne.useMemo(
      () =>
        buildVisibleAxisLayout(
          activeSheet.rowCount,
          defaultRowHeight,
          effectiveRowHeights,
          hiddenRowIndexes,
        ),
      [
        activeSheet.rowCount,
        defaultRowHeight,
        effectiveRowHeights,
        hiddenRowIndexes,
      ],
    ),
    columnAxisLayout = ne.useMemo(
      () =>
        buildVisibleAxisLayout(
          activeSheet.colCount,
          defaultColumnWidth,
          {
            ...activeSheet.columnWidths,
            ...(viewDisplayState.axisResizeState?.axis === "column"
              ? {
                  [viewDisplayState.axisResizeState.index]:
                    viewDisplayState.axisResizeState.size,
                }
              : {}),
          },
          hiddenColumnIndexes,
        ),
      [
        activeSheet.colCount,
        defaultColumnWidth,
        activeSheet.columnWidths,
        hiddenColumnIndexes,
        viewDisplayState.axisResizeState,
      ],
    ),
    frozenRowCount = rowVisibilityIndex.countBefore(pinnedRowCount),
    visibleFrozenColumnCount =
      columnVisibilityIndex.countBefore(frozenColumnCount),
    frozenRowIndexes = Array.from({ length: frozenRowCount }, (fe, viewIndex) =>
      rowVisibilityIndex.logicalIndex(viewIndex),
    ),
    frozenColumnIndexes = Array.from(
      { length: visibleFrozenColumnCount },
      (fe, viewIndex) => columnVisibilityIndex.logicalIndex(viewIndex),
    ),
    scrollableRowLayout = ne.useMemo(() => {
      if (!sortedRowIndexes)
        return sliceVisibleAxis(rowAxisLayout, frozenRowCount);
      const filteredRowHeights = {};
      for (const [rowKey, rowHeight] of Object.entries(effectiveRowHeights)) {
        const filteredPosition = filteredRowPositionMap?.get(Number(rowKey));
        filteredPosition !== undefined &&
          filteredPosition >= 0 &&
          (filteredRowHeights[filteredPosition] = rowHeight);
      }
      return buildVisibleAxisLayout(
        sortedRowIndexes.length,
        defaultRowHeight,
        filteredRowHeights,
      );
    }, [
      sortedRowIndexes,
      filteredRowPositionMap,
      rowAxisLayout,
      frozenRowCount,
      effectiveRowHeights,
      defaultRowHeight,
    ]),
    scrollableColumnLayout = ne.useMemo(
      () => sliceVisibleAxis(columnAxisLayout, visibleFrozenColumnCount),
      [columnAxisLayout, visibleFrozenColumnCount],
    ),
    matchingRowCount = scrollableRowLayout.count,
    logicalRowAt = (scrollableIndex) =>
      sortedRowIndexes
        ? sortedRowIndexes[scrollableIndex]
        : rowVisibilityIndex.logicalIndex(scrollableIndex + frozenRowCount),
    scrollableRowPositionOf = (logicalRow) => {
      if (!(logicalRow < pinnedRowCount))
        return filteredRowPositionMap
          ? filteredRowPositionMap.get(logicalRow)
          : (rowVisibilityIndex.viewIndex(logicalRow) ?? -1) - frozenRowCount;
    },
    frozenRowsHeight = rowAxisLayout.offset(frozenRowCount),
    frozenColumnsWidth = columnAxisLayout.offset(visibleFrozenColumnCount);
  const viewAndSelectionActions = createViewAndSelectionActions({
    selectionState,
    filterSortState,
    navigationState,
    operationStatusState,
    viewDisplayState,
    clipboardState,
    activeReliabilityMonitor,
    setActiveSheetId,
    activeSheet,
    selectionAreas,
    sheetSplitPanes,
    splitPaneSizes,
    rowVisibilityIndex,
    columnVisibilityIndex,
    zoomScale,
    rowHeaderWidth,
    getAllCellEntries,
    sortedRowIndexes,
    yn: visibleFrozenColumnCount,
    frozenRowIndexes,
    frozenColumnIndexes,
    scrollableRowLayout,
    scrollableColumnLayout,
    logicalRowAt,
    frozenRowsHeight,
    frozenColumnsWidth,
  });

  const persistenceActions = createPersistenceActions({
    viewAndSelectionActions,
    persistenceState,
    formulaEntryState,
    operationStatusState,
    auditState,
    assistantState,
    onChange,
    onBeforeChange,
    readOnly,
    persistenceKey,
    workbookStore,
    activeReliabilityMonitor,
    saveWarningDelayMs,
    setWorkbookSnapshot,
    workbookPersistence,
    isEditableRef,
    setSaveAbortControllerRef,
    isEditingLocked,
    setUndoStack,
    setRedoStack,
    calculationRun,
    latestEditStateRef,
  });

  const cellEditActions = createCellEditActions({
    viewAndSelectionActions,
    persistenceActions,
    persistenceState,
    formulaEntryState,
    selectionState,
    navigationState,
    filterSortState,
    operationStatusState,
    actor,
    readOnly,
    workbookMessages,
    getEditableCellText,
    setActiveSheetId,
    activeSheet,
    workbookCalculator,
    latestEditStateRef,
  });

  const navigationActions = createNavigationActions({
    viewAndSelectionActions,
    persistenceActions,
    cellEditActions,
    dialogState,
    selectionState,
    navigationState,
    formulaEntryState,
    filterSortState,
    operationStatusState,
    auditState,
    viewDisplayState,
    findReplaceState,
    actor,
    readOnly,
    activeWorkbook,
    activeSheetId,
    setActiveSheetId,
    activeSheet,
    isEditingLocked,
    selectionAreas,
    sheetSplitPanes,
    splitPaneSizes,
    undoStack,
    setUndoStack,
    redoStack,
    setRedoStack,
    hasFrozenRows,
    rowVisibilityIndex,
    columnVisibilityIndex,
    rowHeaderWidth,
    latestEditStateRef,
    scrollableRowLayout,
    scrollableColumnLayout,
    frozenRowsHeight,
    frozenColumnsWidth,
  });

  const formulaEntryActions = createFormulaEntryActions({
    viewAndSelectionActions,
    cellEditActions,
    navigationActions,
    formulaEntryState,
    selectionState,
    operationStatusState,
    dragFillState,
    activeWorkbook,
    activeSheet,
  });

  const fileAndAnalysisActions = createFileAndAnalysisActions({
    viewAndSelectionActions,
    persistenceActions,
    navigationActions,
    dialogState,
    auditState,
    operationStatusState,
    importExportState,
    assistantState,
    adapter,
    actor,
    activeReliabilityMonitor,
    activeWorkbook,
    activeSheet,
    isEditingLocked,
    selectionAreas,
    activeFilterState,
    operationScope,
    filterResult,
    latestEditStateRef,
    assistantSelection,
  });

  const clipboardAndFindActions = createClipboardAndFindActions({
    viewAndSelectionActions,
    cellEditActions,
    navigationActions,
    selectionState,
    operationStatusState,
    findReplaceState,
    clipboardState,
    activeWorkbook,
    activeSheet,
    isEditingLocked,
    selectionAreasKey,
    workbookCalculator,
    rowVisibilityIndex,
    columnVisibilityIndex,
    latestEditStateRef,
    selectionRangeBounds,
  });

  const {
    rowViewport,
    columnViewport,
    rowWindow,
    columnWindow,
    shownRows,
    shownColumns,
  } = viewAndSelectionActions.computeGridViewport(activePaneLayout);
  const rowColumnAndFillActions = createRowColumnAndFillActions({
    viewAndSelectionActions,
    persistenceActions,
    cellEditActions,
    selectionState,
    viewDisplayState,
    dragFillState,
    actor,
    activeSheet,
    isEditingLocked,
    fillDragStateRef,
    columnVisibilityIndex,
    defaultRowHeight,
    textMeasurer,
    latestEditStateRef,
    selectionRangeBounds,
    getAllCellEntries,
    effectiveRowHeights,
    columnAxisLayout,
    frozenRowIndexes,
    frozenColumnIndexes,
    shownRows,
    shownColumns,
  });

  const gridInteractionActions = createGridInteractionActions({
    viewAndSelectionActions,
    persistenceActions,
    cellEditActions,
    navigationActions,
    dialogState,
    selectionState,
    navigationState,
    formulaEntryState,
    filterSortState,
    operationStatusState,
    importExportState,
    findReplaceState,
    clipboardState,
    getEditableCellText,
    activeWorkbook,
    activeSheet,
    isEditingLocked,
    selectionAreas,
    splitPaneSizes,
    rowVisibilityIndex,
    columnVisibilityIndex,
    defaultRowHeight,
    zoomScale,
    mergeRanges,
    activeCell,
    rowViewport,
  });

  (ne.useEffect(() => {
    const panePositions = {};
    for (const paneId of Pv(sheetSplitPanes)) {
      const paneViewport = viewAndSelectionActions.computeGridViewport(
          splitPaneSizes.current.current[paneId],
        ),
        splitRowIndex = sheetSplitPanes
          ? (rowVisibilityIndex.next(sheetSplitPanes.row, 1, true) ??
            rowVisibilityIndex.next(sheetSplitPanes.row, -1, true) ??
            0)
          : 0,
        splitColumnIndex = sheetSplitPanes
          ? (columnVisibilityIndex.next(sheetSplitPanes.col, 1, true) ??
            columnVisibilityIndex.next(sheetSplitPanes.col, -1, true) ??
            0)
          : 0,
        paneTop = paneId.startsWith("bottom")
          ? logicalToPhysicalPixels(
              scrollableRowLayout.offset(
                Math.max(0, scrollableRowPositionOf(splitRowIndex) ?? 0),
              ),
              scrollableRowLayout.total,
              paneViewport.rowViewport,
              zoomScale,
            )
          : 0,
        left = paneId.endsWith("right")
          ? logicalToPhysicalPixels(
              scrollableColumnLayout.offset(
                Math.max(
                  0,
                  (columnVisibilityIndex.viewIndex(splitColumnIndex) ?? 0) -
                    visibleFrozenColumnCount,
                ),
              ),
              scrollableColumnLayout.total,
              paneViewport.columnViewport,
              zoomScale,
            )
          : 0;
      panePositions[paneId] = { top: paneTop, left: left };
      const paneElement = splitPaneSizes.refs.current[paneId];
      paneElement &&
        ((paneElement.scrollTop = paneTop), (paneElement.scrollLeft = left));
    }
    splitPaneSizes.patch(panePositions);
    const activePaneId = Pv(sheetSplitPanes).at(-1);
    (splitPaneSizes.activate(activePaneId),
      (selectionState.gridContainerRef.current =
        splitPaneSizes.refs.current[activePaneId] ?? null));
  }, [splitPaneSizesKey]),
    ne.useEffect(() => {
      if (
        !selectionState.gridContainerRef.current ||
        typeof ResizeObserver > "u"
      )
        return;
      const resizeObserver = new ResizeObserver((resizeEntries) => {
        const measuredHeights = {};
        for (const entry of resizeEntries) {
          const rowIndex = Number(entry.target.dataset.row),
            offsetHeight = entry.target.offsetHeight;
          Number.isInteger(rowIndex) &&
            activeSheet.rowHeights?.[rowIndex] === undefined &&
            viewDisplayState.axisResizeState?.axis !== "row" &&
            offsetHeight > defaultRowHeight &&
            (measuredHeights[rowIndex] = offsetHeight);
        }
        Object.keys(measuredHeights).length &&
          viewDisplayState.setAutoFitRowHeights((previousHeights) =>
            Object.entries(measuredHeights).every(
              ([rowKey, measuredHeight]) =>
                previousHeights[Number(rowKey)] === measuredHeight,
            )
              ? previousHeights
              : { ...previousHeights, ...measuredHeights },
          );
      });
      return (
        Object.values(splitPaneSizes.refs.current).forEach((paneElement) =>
          paneElement
            .querySelectorAll("tr[data-row]")
            .forEach((rowElement) => resizeObserver.observe(rowElement)),
        ),
        () => resizeObserver.disconnect()
      );
    }),
    ne.useEffect(() => {
      if (
        !navigationState.pendingScrollAddress ||
        !selectionState.gridContainerRef.current
      )
        return;
      const pendingPosition = _t(navigationState.pendingScrollAddress),
        gridContainer = selectionState.gridContainerRef.current;
      if (
        rowVisibilityIndex.isHidden(pendingPosition.row) ||
        columnVisibilityIndex.isHidden(pendingPosition.col)
      ) {
        navigationState.setPendingScrollAddress(null);
        return;
      }
      const scrollRowPosition = scrollableRowPositionOf(pendingPosition.row);
      if (
        scrollRowPosition === undefined &&
        pendingPosition.row >= pinnedRowCount
      ) {
        (filterSortState.setFilterValueSearch(""),
          filterSortState.setActiveColumnSort(null));
        return;
      }
      let paneScrollTop = activePaneScrollTop,
        paneScrollLeft = activePaneScrollLeft;
      if (scrollRowPosition !== undefined && scrollRowPosition >= 0) {
        const rowStartOffset = scrollableRowLayout.offset(scrollRowPosition),
          rowEndOffset = scrollableRowLayout.offset(scrollRowPosition + 1);
        let logicalOffset = rowWindow.logicalOffset;
        (rowStartOffset < logicalOffset
          ? (logicalOffset = rowStartOffset)
          : rowEndOffset > logicalOffset + rowViewport / zoomScale &&
            (logicalOffset = rowEndOffset - rowViewport / zoomScale),
          (paneScrollTop = logicalToPhysicalPixels(
            logicalOffset,
            scrollableRowLayout.total,
            rowViewport,
            zoomScale,
          )));
      }
      const scrollColumnPosition =
        (columnVisibilityIndex.viewIndex(pendingPosition.col) ?? 0) -
        visibleFrozenColumnCount;
      if (scrollColumnPosition >= 0) {
        const columnStartOffset =
            scrollableColumnLayout.offset(scrollColumnPosition),
          columnEndOffset = scrollableColumnLayout.offset(
            scrollColumnPosition + 1,
          );
        let logicalOffset = columnWindow.logicalOffset;
        (columnStartOffset < logicalOffset
          ? (logicalOffset = columnStartOffset)
          : columnEndOffset > logicalOffset + columnViewport / zoomScale &&
            (logicalOffset = columnEndOffset - columnViewport / zoomScale),
          (paneScrollLeft = logicalToPhysicalPixels(
            logicalOffset,
            scrollableColumnLayout.total,
            columnViewport,
            zoomScale,
          )));
      }
      ((gridContainer.scrollTop = paneScrollTop),
        (gridContainer.scrollLeft = paneScrollLeft),
        viewAndSelectionActions.scrollActivePaneVertically(paneScrollTop),
        viewAndSelectionActions.scrollActivePaneHorizontally(paneScrollLeft),
        navigationState.setPendingScrollAddress(null));
    }, [
      navigationState.pendingScrollAddress,
      filteredRowPositionMap,
      scrollableRowLayout,
      scrollableColumnLayout,
      pinnedRowCount,
      visibleFrozenColumnCount,
      rowVisibilityIndex,
      columnVisibilityIndex,
      rowViewport,
      columnViewport,
      zoomScale,
      rowWindow.logicalOffset,
      columnWindow.logicalOffset,
      activePaneScrollTop,
      activePaneScrollLeft,
    ]));

  const selectionReferenceLabel = `${activeSheet.name}!${selectionAreasKey}`;
  (ne.useEffect(() => {
    formulaEntryState.activeCellEditRef.current ||
      formulaEntryState.setCellEditText(
        getEditableCellText(
          activeSheet.cells[selectionState.activeCellAddress],
        ),
      );
  }, [
    activeSheet,
    selectionState.activeCellAddress,
    workbookMessages.entryLocale,
  ]),
    ne.useEffect(() => {
      (navigationState.setGoToReferenceText(selectionAreasKey),
        viewAndSelectionActions.notifyHostCallback("selection", () =>
          onSelectionChange?.({
            sheetId: activeSheet.id,
            range: selectionState.activeSelectionRange,
            ...(selectionAreas.length > 1 ? { ranges: selectionAreas } : {}),
          }),
        ));
    }, [
      selectionAreasKey,
      selectionState.activeSelectionRange,
      activeSheet.id,
      onSelectionChange,
    ]),
    ne.useEffect(() => {
      persistenceState.isWorkbookLoadedRef.current = false;
      const persistenceSession = { active: true, controllers: new Set() };
      persistenceState.persistenceSessionRef.current = persistenceSession;
      try {
        const persistedRecord = workbookPersistence.load();
        if (
          ((persistenceState.persistedWorkbookRecordRef.current =
            persistedRecord),
          persistenceState.setPersistedWorkbookRecord(persistedRecord),
          persistedRecord)
        ) {
          (setWorkbookSnapshot(
            buildCanonicalWorkbookSnapshot(persistedRecord.workbook),
          ),
            (latestEditStateRef.current.book = persistedRecord.workbook),
            viewAndSelectionActions.applyWorkbookView(persistedRecord.view),
            operationStatusState.setStatusNotice(
              persistedRecord.warning ??
                "Restored your locally saved workbook.",
            ),
            persistenceActions.reportSaveStatus(
              "saved",
              persistedRecord.workbook,
            ));
          try {
            viewAndSelectionActions.notifyHostCallback("loaded", () =>
              onWorkbookLoaded?.(persistedRecord.workbook),
            );
          } catch {
            operationStatusState.setErrorNotice(
              "The host could not acknowledge the loaded workbook.",
            );
          }
        }
        persistenceState.isWorkbookLoadedRef.current = true;
      } catch (error) {
        (persistenceActions.reportSaveError(error),
          (persistenceState.persistenceMessageRef.current =
            "The saved workbook could not be restored. Import a backup to recover it. Existing saved data will be preserved."),
          operationStatusState.setErrorNotice(
            persistenceState.persistenceMessageRef.current,
          ));
      }
      return () => {
        persistenceSession.active = false;
        for (const controller of persistenceSession.controllers)
          controller.abort();
        const currentSession = persistenceState.persistenceSessionRef.current;
        currentSession.active = false;
        for (const controller of currentSession.controllers) controller.abort();
        persistenceState.saveAbortControllerRef.current?.abort();
      };
    }, [workbookPersistence]),
    ne.useEffect(() => {
      if (readOnly) {
        for (const controller of persistenceState.persistenceSessionRef.current
          .controllers)
          controller.abort();
        persistenceState.saveAbortControllerRef.current?.abort();
      }
    }, [readOnly]),
    ne.useEffect(() => {
      if (!persistenceKey && !("coordination" in workbookPersistence)) return;
      const handleStorageEvent = (event) => {
        const knownGuard =
            persistenceState.persistedWorkbookRecordRef.current?.guard,
          storageNamespace = persistenceKey ?? knownGuard?.namespace;
        if (!(
          !storageNamespace ||
          (event.key !== null && event.key !== storageNamespace)
        ))
          try {
            const reloadedRecord = workbookPersistence.load();
            if (
              JSON.stringify(reloadedRecord?.guard ?? null) ===
              JSON.stringify(knownGuard ?? null)
            )
              return;
            persistenceActions.reportSaveError(
              new WorkbookRevisionConflictError(
                "conflict",
                "Another workbook window changed the saved version. Your local work has been retained.",
                reloadedRecord?.guard,
              ),
            );
          } catch (error) {
            persistenceActions.reportSaveError(error);
          }
      };
      return (
        window.addEventListener("storage", handleStorageEvent),
        () => window.removeEventListener("storage", handleStorageEvent)
      );
    }, [workbookPersistence, persistenceKey]),
    ne.useEffect(() => {
      if (!persistenceKey && !workbookStorePersistent) return;
      const handleBeforeUnload = (beforeUnloadEvent) => {
        (formulaEntryState.activeCellEditRef.current ||
          persistenceState.persistenceSessionRef.current.controllers.size ||
          JSON.stringify(
            persistenceState.persistedWorkbookRecordRef.current?.workbook,
          ) !== JSON.stringify(latestEditStateRef.current.book)) &&
          (beforeUnloadEvent.preventDefault(),
          (beforeUnloadEvent.returnValue = ""));
      };
      return (
        window.addEventListener("beforeunload", handleBeforeUnload),
        () => window.removeEventListener("beforeunload", handleBeforeUnload)
      );
    }, [persistenceKey, workbookStorePersistent]),
    ne.useEffect(() => {
      if (!persistenceState.isWorkbookLoadedRef.current || readOnly) return;
      const saveTimer = setTimeout(() => {
        try {
          const latestEditState = latestEditStateRef.current;
          if (
            latestEditState.readOnly ||
            persistenceState.persistenceError ||
            persistenceState.persistenceSessionRef.current.controllers.size
          )
            return;
          const viewState = viewAndSelectionActions.buildWorkbookViewState(
              latestEditState.book,
              latestEditState.sheetId,
              latestEditState.range,
            ),
            savedRecord = persistenceState.persistedWorkbookRecordRef.current;
          if (
            savedRecord &&
            JSON.stringify(savedRecord.view) === JSON.stringify(viewState) &&
            (savedRecord.workbook === latestEditState.book ||
              JSON.stringify(savedRecord.workbook) ===
                JSON.stringify(latestEditState.book))
          )
            return;
          const saveResult = persistenceActions.persistWorkbookSnapshot(
            latestEditState.book,
            viewState,
          );
          isPromiseLike(saveResult) && saveResult.catch(() => {});
        } catch (error) {
          operationStatusState.setErrorNotice(
            `${error instanceof Error ? error.message : "The current selection could not be saved."} Changes remain in this editor; export JSON to keep a copy.`,
          );
        }
      }, 400);
      return () => clearTimeout(saveTimer);
    }, [
      activeSheet.id,
      selectionState.activeSelectionRange,
      workbookPersistence,
      readOnly,
      persistenceState.persistenceError,
    ]),
    ne.useEffect(() => {
      (auditState.setAuditAnalysisResult(null),
        auditState.setAgingReportPayload(null));
    }, [
      auditState.auditSourceRange,
      auditState.materialityThreshold,
      auditState.auditSampleSize,
      auditState.samplingSeed,
      auditState.agingAsOfDate,
    ]),
    ne.useEffect(
      () => (
        assistantState.assistantRequestIdRef.current++,
        assistantState.assistantAbortControllerRef.current?.abort(),
        assistantState.setIsAssistantBusy(false),
        assistantState.setAssistantProposal(null),
        assistantState.setAnswerCitations(null),
        () => assistantState.assistantAbortControllerRef.current?.abort()
      ),
      [activeWorkbook, activeSheet.id, selectionAreasKey, adapter],
    ),
    ne.useEffect(() => {
      if (
        !externalProposal ||
        externalProposal.requestId ===
          assistantState.lastExternalProposalIdRef.current
      )
        return;
      assistantState.lastExternalProposalIdRef.current =
        externalProposal.requestId;
      let proposalResult;
      try {
        if (
          JSON.stringify(externalProposal.context) !==
          JSON.stringify(
            buildAssistantSelectionContext(activeWorkbook, {
              sheetId: activeSheet.id,
              range: selectionState.activeSelectionRange,
              ranges: selectionAreas,
            }),
          )
        )
          throw new Error(
            "The workbook or selected context changed. Request a new external proposal.",
          );
        const validatedProposal = validateProposedCalculation(
          activeWorkbook,
          {
            sheetId: activeSheet.id,
            range: selectionState.activeSelectionRange,
          },
          externalProposal.proposal,
        );
        (assistantState.setAnswerCitations(
          buildAnswerCitations(
            activeWorkbook,
            {
              sheetId: activeSheet.id,
              range: selectionState.activeSelectionRange,
            },
            validatedProposal,
            externalProposal.context,
          ),
        ),
          assistantState.setAssistantProposal(validatedProposal),
          dialogState.setSidePanelKind("assistant"),
          operationStatusState.setErrorNotice(""),
          (proposalResult = { accepted: true }));
      } catch (error) {
        const errorMessage =
          error instanceof Error ? error.message : "Invalid external proposal.";
        (operationStatusState.setErrorNotice(errorMessage),
          (proposalResult = { accepted: false, error: errorMessage }));
      }
      try {
        viewAndSelectionActions.notifyHostCallback("proposal", () =>
          onExternalProposalStaged?.(
            externalProposal.requestId,
            proposalResult,
          ),
        );
      } catch {
        operationStatusState.setErrorNotice(
          "The host could not acknowledge the staged proposal.",
        );
      }
    }, [
      externalProposal,
      activeWorkbook,
      activeSheet.id,
      selectionAreasKey,
      onExternalProposalStaged,
    ]));

  formulaEntryState.formulaPointerMoveRef.current = (cellAddress) => {
    formulaEntryState.formulaPointerStateRef.current.head !== cellAddress &&
      formulaEntryActions.insertFormulaReference(cellAddress, true);
  };

  ne.useEffect(() => {
    if (!viewDisplayState.axisResizeState) return;
    const handleAxisPointerMove = (event) =>
        viewDisplayState.setAxisResizeState(
          (previousResizeState) =>
            previousResizeState && {
              ...previousResizeState,
              size: Math.round(
                Math.max(
                  previousResizeState.axis === "column"
                    ? minColumnWidth
                    : minRowHeight,
                  Math.min(
                    previousResizeState.axis === "column"
                      ? maxColumnWidth
                      : maxRowHeight,
                    previousResizeState.initial +
                      ((previousResizeState.axis === "column"
                        ? event.clientX
                        : event.clientY) -
                        previousResizeState.start) /
                        zoomScale,
                  ),
                ),
              ),
            },
        ),
      handleAxisPointerUp = () => {
        (viewDisplayState.axisResizeState.size !==
          viewDisplayState.axisResizeState.initial &&
          viewAndSelectionActions.runGuardedAction(() =>
            rowColumnAndFillActions.applyAxisSizes(
              viewDisplayState.axisResizeState.axis,
              {
                [viewDisplayState.axisResizeState.index]:
                  viewDisplayState.axisResizeState.size,
              },
              `Resize ${viewDisplayState.axisResizeState.axis} ${viewDisplayState.axisResizeState.axis === "column" ? columnIndexToLetters(viewDisplayState.axisResizeState.index) : viewDisplayState.axisResizeState.index + 1}`,
            ),
          ),
          viewDisplayState.setAxisResizeState(null));
      },
      cancelAxisResize = () => viewDisplayState.setAxisResizeState(null),
      handleAxisKeyDown = (keyboardEvent) => {
        keyboardEvent.key === "Escape" &&
          (!keyboardEvent.defaultPrevented &&
            keyboardEvent.target instanceof Element &&
            keyboardEvent.target.closest(".lg-app") ===
              selectionState.appRootRef.current &&
            keyboardEvent.preventDefault(),
          cancelAxisResize());
      };
    return (
      window.addEventListener("pointermove", handleAxisPointerMove),
      window.addEventListener("pointerup", handleAxisPointerUp),
      window.addEventListener("pointercancel", cancelAxisResize),
      window.addEventListener("keydown", handleAxisKeyDown),
      () => {
        (window.removeEventListener("pointermove", handleAxisPointerMove),
          window.removeEventListener("pointerup", handleAxisPointerUp),
          window.removeEventListener("pointercancel", cancelAxisResize),
          window.removeEventListener("keydown", handleAxisKeyDown));
      }
    );
  }, [
    viewDisplayState.axisResizeState,
    activeWorkbook,
    activeSheet.id,
    zoomScale,
    isEditingLocked,
  ]);

  ne.useEffect(() => {
    if (!dragFillState.fillDragState) return;
    const clearFillDrag = () => {
        ((fillDragStateRef.current = null),
          dragFillState.setFillDragState(null));
      },
      handleFillPointerUp = () => {
        const fillDrag = fillDragStateRef.current;
        if (
          (clearFillDrag(),
          !(!fillDrag || fillDrag.targetRange === fillDrag.sourceRange))
        ) {
          if (
            latestEditStateRef.current.readOnly ||
            latestEditStateRef.current.book !== fillDrag.book ||
            latestEditStateRef.current.sheetId !== fillDrag.sheetId
          ) {
            operationStatusState.setErrorNotice(
              "The workbook changed while filling. Select the cells and try again.",
            );
            return;
          }
          dragFillState.setFillPreviewState(fillDrag);
        }
      },
      handleFillKeyDown = (keyboardEvent) => {
        keyboardEvent.key === "Escape" &&
          (!keyboardEvent.defaultPrevented &&
            keyboardEvent.target instanceof Element &&
            keyboardEvent.target.closest(".lg-app") ===
              selectionState.appRootRef.current &&
            keyboardEvent.preventDefault(),
          clearFillDrag());
      };
    return (
      window.addEventListener("pointerup", handleFillPointerUp),
      window.addEventListener("pointercancel", clearFillDrag),
      window.addEventListener("keydown", handleFillKeyDown),
      () => {
        (window.removeEventListener("pointerup", handleFillPointerUp),
          window.removeEventListener("pointercancel", clearFillDrag),
          window.removeEventListener("keydown", handleFillKeyDown));
      }
    );
  }, [!!dragFillState.fillDragState]);

  const frozenColumnLeftOffsets = Object.fromEntries(
      frozenColumnIndexes.map((frozenColumnIndex, frozenColumnPosition) => [
        frozenColumnIndex,
        rowHeaderWidth + columnAxisLayout.offset(frozenColumnPosition),
      ]),
    ),
    renderColumnHeader = (columnIndex, isFrozenColumn = false) =>
      i.jsxs(
        "th",
        {
          role: "columnheader",
          scope: "col",
          "data-column": columnIndex,
          "aria-colindex": columnIndex + 2,
          className: `${_t(selectionState.activeCellAddress).col === columnIndex ? "lg-selected-header" : ""} ${isFrozenColumn ? "lg-frozen-column" : ""}`,
          style: isFrozenColumn
            ? {
                position: "sticky",
                left: frozenColumnLeftOffsets[columnIndex],
                zIndex: 7,
              }
            : undefined,
          onClick: (mouseEvent) => {
            formulaEntryActions.insertFormulaReference(
              `${columnIndexToLetters(columnIndex)}1`,
              false,
              {
                range: `${columnIndexToLetters(columnIndex)}:${columnIndexToLetters(columnIndex)}`,
                bounds: {
                  firstRow: 0,
                  lastRow: activeSheet.rowCount - 1,
                  firstCol: columnIndex,
                  lastCol: columnIndex,
                },
              },
            ) ||
              navigationActions.selectRowOrColumn(
                "column",
                columnIndex,
                mouseEvent.shiftKey,
              );
          },
          children: [
            columnIndexToLetters(columnIndex),
            i.jsx("span", {
              className: "lg-column-resize",
              role: "separator",
              "aria-label": `Resize column ${columnIndexToLetters(columnIndex)}`,
              "aria-orientation": "vertical",
              "aria-disabled": isEditingLocked,
              title: text(
                "Drag to resize; double-click to AutoFit column width",
              ),
              onPointerDown: (pointerEvent) =>
                rowColumnAndFillActions.beginAxisResize(
                  "column",
                  columnIndex,
                  pointerEvent,
                ),
              onDoubleClick: (event) => {
                (event.stopPropagation(),
                  rowColumnAndFillActions.autoFitSelectedAxis(
                    "column",
                    columnIndex,
                  ));
              },
              onClick: (event) => event.stopPropagation(),
              lang: language(
                "Drag to resize; double-click to AutoFit column width",
              ),
            }),
          ],
        },
        columnIndex,
      ),
    renderGridRow = (rowIndex, pinned, renderWindow, paneId) =>
      i.jsx(
        WorksheetGridRow,
        {
          idPrefix: `${editorIdPrefix}-${paneId}`,
          row: rowIndex,
          pinned: pinned,
          pinnedTop:
            (viewDisplayState.isHeadingsVisible ? 22 : 0) +
            rowAxisLayout.offset(rowVisibilityIndex.viewIndex(rowIndex) ?? 0),
          frozenColumns: frozenColumnIndexes,
          frozenColumnOffsets: frozenColumnLeftOffsets,
          headings: viewDisplayState.isHeadingsVisible,
          columns: renderWindow.shownColumns,
          leftGap: renderWindow.leftGap,
          rightGap: renderWindow.rightGap,
          minimumHeight: effectiveRowHeights[rowIndex] ?? defaultRowHeight,
          fixedHeight:
            activeSheet.rowHeights?.[rowIndex] !== undefined ||
            (viewDisplayState.axisResizeState?.axis === "row" &&
              viewDisplayState.axisResizeState.index === rowIndex),
          onRowSelect: (mouseEvent) => {
            formulaEntryActions.insertFormulaReference(
              `A${rowIndex + 1}`,
              false,
              {
                range: `${rowIndex + 1}:${rowIndex + 1}`,
                bounds: {
                  firstRow: rowIndex,
                  lastRow: rowIndex,
                  firstCol: 0,
                  lastCol: activeSheet.colCount - 1,
                },
              },
            ) ||
              navigationActions.selectRowOrColumn("row", rowIndex, mouseEvent);
          },
          onResize: (pointerEvent) =>
            rowColumnAndFillActions.beginAxisResize(
              "row",
              rowIndex,
              pointerEvent,
            ),
          onAutoFit: () =>
            rowColumnAndFillActions.autoFitSelectedAxis("row", rowIndex),
          onCancelEdit: cellEditActions.cancelCellEdit,
          onFormulaInput: cellEditActions.syncFormulaEditorSelection,
          onFormulaKey: formulaEntryActions.handleFormulaKeyDown,
          formulaHelpId: `${editorIdPrefix}-formula-help`,
          onFormulaPick: formulaEntryActions.insertFormulaReference,
          formulaReference:
            formulaEntryState.activeFormulaReference?.sheetId === activeSheet.id
              ? formulaEntryState.activeFormulaReference.bounds
              : null,
          formulaDraftActive:
            !!formulaEntryState.activeCellEditRef.current?.value.startsWith(
              "=",
            ),
          onDragStart: (anchorAddress, shouldKeepSelection) => {
            ((dragFillState.dragAnchorRef.current = anchorAddress),
              (dragFillState.dragBaseSelectionsRef.current = shouldKeepSelection
                ? [...selectionAreas]
                : []));
          },
          onDragEnter: (targetAddress, pressedButtons) => {
            if (
              formulaEntryState.formulaPointerStateRef.current.dragging &&
              pressedButtons === 1
            ) {
              formulaEntryActions.insertFormulaReference(targetAddress, true);
              return;
            }
            fillDragStateRef.current && pressedButtons === 1
              ? rowColumnAndFillActions.extendFillDragTo(targetAddress)
              : dragFillState.dragAnchorRef.current &&
                pressedButtons === 1 &&
                ((selectionState.activeCellAddressRef.current = targetAddress),
                viewAndSelectionActions.runGuardedAction(() =>
                  viewAndSelectionActions.applySelectionAreas([
                    ...dragFillState.dragBaseSelectionsRef.current,
                    wy(dragFillState.dragAnchorRef.current, targetAddress),
                  ]),
                ),
                selectionState.setActiveCellAddress(
                  dragFillState.dragAnchorRef.current,
                ));
          },
          onNavigate: navigationActions.moveSelectionAfterCommit,
          onPivotDrill: (cellAddress) => {
            const position = _t(cellAddress);
            return activeSheet.pivots?.some((pivot) => {
              const rangeAddress = parseRangeAddress(pivot.outputRange);
              return (
                position.row > rangeAddress.firstRow &&
                position.row <= rangeAddress.lastRow &&
                position.col >=
                  rangeAddress.firstCol + Math.max(1, pivot.rowFields.length) &&
                position.col <= rangeAddress.lastCol
              );
            })
              ? (navigationActions.selectCellAddress(cellAddress),
                gridInteractionActions.runEditorCommand("pivotExplore"),
                true)
              : false;
          },
          commentAddresses: commentedCellAddresses,
          onReview: (cellAddress) => {
            (navigationActions.selectCellAddress(cellAddress),
              gridInteractionActions.runEditorCommand("comments"));
          },
          onLink: (Nt, wt) =>
            viewAndSelectionActions.runGuardedAction(() =>
              navigationActions.navigateToSheetCell(Nt, wt),
            ),
          onFillStart: rowColumnAndFillActions.beginFillDrag,
          onFillOptions: () =>
            gridInteractionActions.runEditorCommand("fillSeries"),
          fillDragging:
            !!dragFillState.fillDragState ||
            selectionState.completedSelectionAreas.length > 0,
          workbook: activeWorkbook,
          mergeBounds: mergeRanges,
          displayedRows: pinned ? frozenRowIndexes : renderWindow.shownRows,
          filterHeader:
            filterResult &&
            rowIndex === parseRangeAddress(filterResult.context.range).firstRow
              ? filterResult.fields
              : undefined,
          filteredColumns:
            filterResult?.context.definition.columns.map(
              (column) => column.column,
            ) ?? [],
          onColumnFilter: (columnIndex) => {
            (gridInteractionActions.runEditorCommand("populationFilter"),
              filterSortState.setActiveFilterColumn(columnIndex));
          },
          sortedView: !!(
            filterSortState.filterValueSearch ||
            filterSortState.activeColumnSort
          ),
          ruleStyles: getConditionalRuleStyle,
          conditionalFormatter: conditionalFormatEvaluator,
          sheetId: activeSheet.id,
          cells: activeSheet.cells,
          active: selectionState.activeCellAddress,
          selectedBounds: dragFillState.fillDragState
            ? parseRangeAddress(dragFillState.fillDragState.targetRange)
            : selectionRangeBounds,
          selectionBounds: dragFillState.fillDragState
            ? [parseRangeAddress(dragFillState.fillDragState.targetRange)]
            : selectionAreaRanges,
          calc: workbookCalculator,
          showFormulas: viewDisplayState.isShowingFormulas,
          editing:
            splitPaneSizes.activeId === paneId &&
            (!formulaEntryState.activeCellEditRef.current ||
              formulaEntryState.activeCellEditRef.current.sheetId ===
                activeSheet.id)
              ? formulaEntryState.cellEditState
              : null,
          setEditing: (nextCellEditState) =>
            nextCellEditState === null
              ? cellEditActions.cancelCellEdit()
              : cellEditActions.beginCellEdit(nextCellEditState, true),
          select: (cellAddress, wt, jt) => {
            (navigationActions.selectCellAddress(cellAddress, wt, jt),
              selectionState.gridContainerRef.current?.focus({
                preventScroll: true,
              }));
          },
          saveCell: cellEditActions.commitCellEdit,
          readOnly: isEditingLocked,
        },
        rowIndex,
      ),
    renderVirtualSpacer = (height, spacerKey, colSpan) =>
      height > 0
        ? i.jsx(
            "tr",
            {
              "aria-hidden": "true",
              className: "lg-virtual-spacer",
              style: { height: height },
              children: i.jsx("td", {
                colSpan: colSpan,
                style: { height: height },
              }),
            },
            spacerKey,
          )
        : null;
  const { renderGridPane } = createGridPaneRenderer({
    viewAndSelectionActions,
    cellEditActions,
    navigationActions,
    formulaEntryActions,
    clipboardAndFindActions,
    gridInteractionActions,
    selectionState,
    filterSortState,
    operationStatusState,
    viewDisplayState,
    dragFillState,
    text,
    language,
    editorIdPrefix,
    activeWorkbook,
    activeSheet,
    isEditingLocked,
    selectionAreas,
    selectionAreaRanges,
    sheetSplitPanes,
    splitPaneSizes,
    workbookCalculator,
    rowVisibilityIndex,
    columnVisibilityIndex,
    pinnedRowCount,
    frozenColumnCount,
    zoomScale,
    rowHeaderWidth,
    mergeRanges,
    rowAxisLayout,
    columnAxisLayout,
    yn: visibleFrozenColumnCount,
    frozenRowIndexes,
    frozenColumnIndexes,
    scrollableRowLayout,
    scrollableColumnLayout,
    scrollableRowPositionOf,
    frozenRowsHeight,
    frozenColumnsWidth,
    renderColumnHeader,
    renderGridRow,
    renderVirtualSpacer,
  });

  const closeContextMenu = ne.useCallback((shouldRestoreFocus) => {
      (selectionState.setContextMenuState(null),
        shouldRestoreFocus &&
          selectionState.gridContainerRef.current?.focus({
            preventScroll: true,
          }));
    }, []),
    editBlockReason = operationStatusState.isOperationBusy
      ? "Wait for the current workbook operation to finish."
      : isEditingLocked
        ? readOnly
          ? "The host has opened this workbook read-only."
          : "Unprotect the worksheet before editing."
        : undefined,
    multiRangeBlockReason =
      selectionAreas.length > 1
        ? "Select one rectangular range for this command."
        : undefined;
  let selectionReadError;
  try {
    selectionState.contextMenuState &&
      viewAndSelectionActions.getSelectedCellAddresses();
  } catch (error) {
    selectionReadError = error.message;
  }
  const contextMenuActions = [
    {
      name: "Cut",
      run: () => void clipboardAndFindActions.copyOrCutSelection(true),
      reason: editBlockReason ?? multiRangeBlockReason ?? selectionReadError,
    },
    {
      name: "Copy",
      run: () => void clipboardAndFindActions.copyOrCutSelection(false),
      reason: multiRangeBlockReason ?? selectionReadError,
    },
    {
      name: "Paste",
      run: () => void clipboardAndFindActions.pasteFromClipboard(),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "Paste Special",
      run: () => gridInteractionActions.runEditorCommand("pasteSpecial"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "Format Cells",
      run: () => gridInteractionActions.runEditorCommand("numberFormat"),
      reason: editBlockReason,
    },
    {
      name: "Cell Styles",
      run: () => gridInteractionActions.runEditorCommand("cellStyle"),
      reason: editBlockReason,
    },
    {
      name: "Insert / Delete / Move Rows",
      run: () => gridInteractionActions.runEditorCommand("insertRow"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "Insert / Delete / Move Columns",
      run: () => gridInteractionActions.runEditorCommand("insertColumn"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "AutoFit Column Width",
      run: () => rowColumnAndFillActions.autoFitSelectedAxis("column"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "AutoFit Row Height",
      run: () => rowColumnAndFillActions.autoFitSelectedAxis("row"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "Hide / Unhide / Freeze",
      run: () => gridInteractionActions.runEditorCommand("viewSettings"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "Clear Contents",
      run: () =>
        viewAndSelectionActions.runGuardedAction(() =>
          cellEditActions.applyCellPatches(
            viewAndSelectionActions
              .getSelectedCellAddresses()
              .map((address) => ({
                sheetId: activeSheet.id,
                address: address,
                cell: { ...activeSheet.cells[address], raw: null },
              })),
            "Clear selection",
          ),
        ),
      reason: editBlockReason ?? selectionReadError,
    },
    {
      name: "Add / Edit Note",
      run: () =>
        dialogState.setPromptDialogState({
          kind: "note",
          value: activeCell?.note ?? "",
        }),
      reason: editBlockReason,
    },
    {
      name: "Comments",
      run: () => gridInteractionActions.runEditorCommand("comments"),
    },
    {
      name: "Link",
      run: () => gridInteractionActions.runEditorCommand("hyperlink"),
    },
  ];
  return i.jsxs("div", {
    className: "lg-app",
    ref: selectionState.appRootRef,
    lang: workbookMessages.locale,
    role: "region",
    "aria-label": activeWorkbook.name,
    onKeyDown: (keyboardEvent) => {
      if (
        keyboardEvent.key !== "F6" ||
        keyboardEvent.defaultPrevented ||
        keyboardEvent.ctrlKey ||
        keyboardEvent.metaKey ||
        keyboardEvent.altKey ||
        keyboardEvent.nativeEvent.isComposing ||
        dialogState.activeDialogMode ||
        dialogState.promptDialogState ||
        (formulaEntryState.activeCellEditRef.current &&
          (formulaEntryState.activeCellEditRef.current.value !==
            formulaEntryState.activeCellEditRef.current.originalText ||
            !cellEditActions.commitCellEdit()))
      )
        return;
      keyboardEvent.preventDefault();
      const focusTargets = [
          selectionState.appRootRef.current?.querySelector(
            '[role="tab"][aria-selected="true"]',
          ),
          formulaEntryState.formulaInputRef.current,
          selectionState.gridContainerRef.current,
          selectionState.appRootRef.current?.querySelector(
            '[data-sheet-id][aria-current="true"]',
          ),
        ].filter((target) => !!target),
        activeElement = document.activeElement,
        currentFocusIndex = activeElement?.closest(".xl-ribbon")
          ? 0
          : activeElement?.closest(".lg-sheet-tabs")
            ? 3
            : focusTargets.findIndex(
                (target) =>
                  target === activeElement || target.contains(activeElement),
              ),
        focusDirection = keyboardEvent.shiftKey ? -1 : 1;
      focusTargets[
        currentFocusIndex < 0
          ? keyboardEvent.shiftKey
            ? focusTargets.length - 1
            : 0
          : (currentFocusIndex + focusDirection + focusTargets.length) %
            focusTargets.length
      ]?.focus({ preventScroll: true });
    },
    children: [
      i.jsxs("div", {
        className: "lg-skip-links",
        children: [
          i.jsx("button", {
            onClick: () => selectionState.gridContainerRef.current?.focus(),
            children: i.jsx(LocalizedText, { children: "Skip to worksheet" }),
          }),
          i.jsx("button", {
            onClick: () => formulaEntryState.formulaInputRef.current?.focus(),
            children: i.jsx(LocalizedText, { children: "Skip to formula bar" }),
          }),
        ],
      }),
      i.jsx("p", {
        id: `${editorIdPrefix}-grid-help`,
        className: "lg-sr-only",
        children: i.jsx(LocalizedText, {
          children:
            "Arrow keys move between cells. F2 edits. Escape returns to the ribbon. Use the selected range field to jump to a cell. Formula names, references and decimal points inside formulas use invariant syntax; enter dates as YYYY-MM-DD.",
        }),
      }),
      i.jsx("p", {
        className: "lg-sr-only",
        role: "status",
        "aria-live": "polite",
        "aria-atomic": "true",
        "aria-label": workbookMessages.text("Cell position"),
        children: workbookMessages.message(
          "{sheet}, cell {address}, row {row}, column {column}. {value}{formula}{access}",
          {
            sheet: activeSheet.name,
            address: selectionState.activeCellAddress,
            row: _t(selectionState.activeCellAddress).row + 1,
            column: _t(selectionState.activeCellAddress).col + 1,
            value: (workbookCalculator(
              activeSheet.id,
              selectionState.activeCellAddress,
            ) == null
              ? workbookMessages.text("Blank")
              : formatLocalizedNumberDisplay(
                  workbookCalculator(
                    activeSheet.id,
                    selectionState.activeCellAddress,
                  ),
                  activeSheet.cells[selectionState.activeCellAddress],
                  activeWorkbook.datePolicy?.dateSystem,
                  workbookMessages.entryLocale,
                )
            ).slice(0, 250),
            formula:
              typeof activeSheet.cells[selectionState.activeCellAddress]?.raw ==
                "string" &&
              String(
                activeSheet.cells[selectionState.activeCellAddress].raw,
              ).startsWith("=") &&
              !activeSheet.cells[selectionState.activeCellAddress].literal
                ? workbookMessages.message(" Formula: {formula}.", {
                    formula: String(
                      activeSheet.cells[selectionState.activeCellAddress].raw,
                    ).slice(0, 250),
                  })
                : "",
            access: isEditingLocked ? workbookMessages.text(" Read only.") : "",
          },
        ),
      }),
      selectionState.contextMenuState &&
        i.jsx(
          GridContextMenu,
          {
            ...selectionState.contextMenuState,
            label: `${selectionState.contextMenuState.kind === "cell" ? "Cells" : selectionState.contextMenuState.kind === "row" ? "Rows" : "Columns"} ${selectionState.activeSelectionRange}`,
            actions: contextMenuActions,
            workbook: activeWorkbook,
            selection: `${activeSheet.id}:${selectionAreas.join(";")}`,
            onClose: closeContextMenu,
          },
          `${selectionState.contextMenuState.x}:${selectionState.contextMenuState.y}:${selectionState.contextMenuState.kind}`,
        ),
      dragFillState.fillPreviewState &&
        i.jsx(FillPreviewDialog, {
          workbook: activeWorkbook,
          sheetId: dragFillState.fillPreviewState.sheetId,
          sourceRange: dragFillState.fillPreviewState.sourceRange,
          targetRange: dragFillState.fillPreviewState.targetRange,
          readOnly: isEditingLocked,
          onWorkbookChange: navigationActions.commitWorkbookEdit,
          onSelectionChange: viewAndSelectionActions.selectSingleRange,
          onClose: () => {
            (dragFillState.setFillPreviewState(null),
              viewAndSelectionActions.focusGridSurface());
          },
        }),
      clipboardState.moveCellsSource &&
        i.jsx(MoveCellsDialog, {
          workbook: activeWorkbook,
          source: clipboardState.moveCellsSource,
          sheetId: activeSheet.id,
          range: selectionState.activeSelectionRange,
          readOnly,
          onWorkbookChange: navigationActions.replaceWorkbookSnapshot,
          onSelectionChange: viewAndSelectionActions.selectSingleRange,
          onApplied: viewAndSelectionActions.clearCutSession,
          onClose: () => {
            (clipboardState.setMoveCellsSource(null),
              viewAndSelectionActions.focusGridSurface());
          },
        }),
      i.jsx(EditorDialogs, {
        viewAndSelectionActions,
        persistenceActions,
        cellEditActions,
        navigationActions,
        persistenceState,
        dialogState,
        navigationState,
        formulaEntryState,
        selectionState,
        filterSortState,
        auditState,
        operationStatusState,
        importExportState,
        clipboardState,
        extensions,
        allowLinkedWorkbookFiles,
        documentExtractor,
        externalSourceAdapter,
        consolidationDefinition,
        allowConsolidationFiles,
        actor,
        readOnly,
        persistenceKey,
        workbookStore,
        workbookStorePersistent,
        activeReliabilityMonitor,
        activeWorkbook,
        effectiveLinkedWorkbooks,
        setLinkedWorkbooks,
        effectiveConsolidationWorkbooks,
        setConsolidationWorkbooks,
        activeSheet,
        workbookPersistence,
        isEditingLocked,
        selectionAreas,
        filterStateKey,
        activeFilterState,
        operationScope,
        calculationRun,
        workbookCalculator,
        filterResult,
        latestEditStateRef,
        handleCalculationCompleted,
      }),
      importExportState.delimitedImportExportSource &&
        i.jsx(DelimitedImportExportDialog, {
          source: importExportState.delimitedImportExportSource,
          blocked:
            importExportState.delimitedImportExportSource.snapshot !==
            activeWorkbook
              ? "The workbook changed. Cancel and review again for the current revision."
              : "bytes" in importExportState.delimitedImportExportSource &&
                  isEditingLocked
                ? "This workbook is read-only. Import is unavailable."
                : undefined,
          onClose: () => importExportState.setDelimitedImportExportSource(null),
          onReport: (report) =>
            saveBlobToFile(
              new Blob([JSON.stringify(report, null, 2)], {
                type: "application/json",
              }),
              `${report.fileName}.schema.json`,
            ),
          onImport: (workbook, importSummary) => {
            if (
              latestEditStateRef.current.book !==
              importExportState.delimitedImportExportSource.snapshot
            )
              throw new Error("The workbook changed. Review the file again.");
            if (latestEditStateRef.current.readOnly)
              throw new Error("This workbook is read-only.");
            (persistenceActions.commitWorkbookChange(workbook),
              navigationActions.switchActiveWorksheet(workbook.sheets[0].id),
              operationStatusState.setStatusNotice(
                `Imported ${importSummary.records} records using the reviewed schema. Formula-like source text remains literal.`,
              ),
              importExportState.setDelimitedImportExportSource(null));
          },
          onExport: (exportBytes, exportReview) => {
            if (
              latestEditStateRef.current.book !==
              importExportState.delimitedImportExportSource.snapshot
            )
              throw new Error("The workbook changed. Review the export again.");
            if (
              exportReview.population &&
              !("bytes" in importExportState.delimitedImportExportSource)
            ) {
              const csvExport = exportWorksheetToCsv(
                latestEditStateRef.current.book,
                importExportState.delimitedImportExportSource.sheetId,
                exportReview.options,
              );
              if (
                JSON.stringify(csvExport.report.population) !==
                  JSON.stringify(exportReview.population) ||
                csvExport.bytes.length !== exportBytes.length ||
                csvExport.bytes.some(
                  (byte, byteIndex) => byte !== exportBytes[byteIndex],
                )
              )
                throw Error(
                  "Calculated values or population membership changed. Preview the export again.",
                );
            }
            (saveBlobToFile(
              new Blob([Uint8Array.from(exportBytes)], {
                type: "application/octet-stream",
              }),
              exportReview.fileName,
            ),
              operationStatusState.setStatusNotice(
                `Delimited data exported; ${exportReview.safeTextChanges} formula-like text values neutralized under the reviewed safe-export policy.`,
              ),
              importExportState.setDelimitedImportExportSource(null));
          },
        }),
      importExportState.xlsxConversionReview &&
        i.jsx(XlsxConversionReview, {
          report: importExportState.xlsxConversionReview.report,
          inspectOnly: importExportState.xlsxConversionReview.inspectOnly,
          onOriginal: () =>
            viewAndSelectionActions.runGuardedAction(() => {
              const retainedXlsxSource = requireRetainedXlsxSource(
                importExportState.xlsxConversionReview.imported ??
                  importExportState.xlsxConversionReview.snapshot,
              );
              saveBlobToFile(
                new Blob([new Uint8Array(retainedXlsxSource.bytes)], {
                  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                }),
                retainedXlsxSource.fileName,
              );
            }),
          blocked:
            importExportState.xlsxConversionReview.snapshot !== activeWorkbook
              ? "The workbook changed. Cancel and review the file again for the current revision."
              : importExportState.xlsxConversionReview.imported &&
                  isEditingLocked
                ? "This workbook is read-only. Import is unavailable."
                : undefined,
          onClose: () => importExportState.setXlsxConversionReview(null),
          onReport: () =>
            saveBlobToFile(
              new Blob(
                [
                  JSON.stringify(
                    importExportState.xlsxConversionReview.report,
                    null,
                    2,
                  ),
                ],
                {
                  type: "application/json",
                },
              ),
              `${importExportState.xlsxConversionReview.fileName}.compatibility.json`,
            ),
          onConfirm: () =>
            viewAndSelectionActions.runGuardedAction(() => {
              if (
                latestEditStateRef.current.book !==
                importExportState.xlsxConversionReview.snapshot
              )
                throw new Error("The workbook changed. Review the file again.");
              if (importExportState.xlsxConversionReview.imported) {
                if (latestEditStateRef.current.readOnly)
                  throw new Error("This workbook is read-only.");
                (persistenceActions.commitWorkbookChange(
                  importExportState.xlsxConversionReview.imported,
                ),
                  navigationActions.switchActiveWorksheet(
                    importExportState.xlsxConversionReview.imported.sheets.find(
                      (sheet) => !sheet.hidden,
                    ).id,
                  ),
                  operationStatusState.setStatusNotice(
                    "Derived XLSX workbook opened. The exact original is available from File → Original XLSX and compatibility. Save native JSON to retain it across sessions.",
                  ));
              } else
                importExportState.xlsxConversionReview.blob &&
                  (saveBlobToFile(
                    importExportState.xlsxConversionReview.blob,
                    importExportState.xlsxConversionReview.fileName,
                  ),
                  operationStatusState.setStatusNotice(
                    "Derived XLSX downloaded. Keep a native JSON backup for full application metadata.",
                  ));
              importExportState.setXlsxConversionReview(null);
            }),
        }),
      i.jsxs("div", {
        className: "lg-main",
        children: [
          i.jsx("input", {
            type: "file",
            "aria-label": text("Import workbook file"),
            ref: importExportState.importFileInputRef,
            accept:
              ".xlsx,.xltx,.xls,.xlsb,.xlsm,.xltm,.ods,.csv,.tsv,.json,.bas,.vba,.cls,.ts,.osts,.xlam",
            hidden: true,
            onChange: (event) =>
              void fileAndAnalysisActions.importWorkbookFromFile(
                event.target.files?.[0],
              ),
            lang: language("Import workbook file"),
            "data-ui-label": "Import workbook file",
          }),
          i.jsx(EditorRibbonBar, {
            viewAndSelectionActions,
            cellEditActions,
            navigationActions,
            fileAndAnalysisActions,
            clipboardAndFindActions,
            rowColumnAndFillActions,
            gridInteractionActions,
            selectionState,
            dialogState,
            filterSortState,
            viewDisplayState,
            operationStatusState,
            findReplaceState,
            importExportState,
            readOnly,
            activeWorkbook,
            activeSheet,
            isEditingLocked,
            undoStack,
            redoStack,
            hasFrozenRows,
            activeCell,
          }),
          i.jsx(ConsolidationStatusBanner, {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            options: { approvedWorkbooks: effectiveConsolidationWorkbooks },
            onReview: () =>
              gridInteractionActions.runEditorCommand("consolidation"),
          }),
          i.jsx(CustomFunctionReviewBanner, {
            workbook: activeWorkbook,
            extensions,
            onReview: () =>
              gridInteractionActions.runEditorCommand("extensions"),
          }),
          i.jsx(WorkbookLinkStatusBadge, {
            workbook: activeWorkbook,
            sources: effectiveLinkedWorkbooks,
            onReview: () =>
              gridInteractionActions.runEditorCommand("workbookLinks"),
          }),
          i.jsx(ExternalSourceStatusPanel, {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            onReview: () =>
              gridInteractionActions.runEditorCommand("externalSources"),
          }),
          i.jsx(AnalysisSnapshotNotice, {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            onReview: () =>
              gridInteractionActions.runEditorCommand(
                activeSheet.analysisReport?.request.kind === "solver"
                  ? "solver"
                  : "statistics",
              ),
          }),
          i.jsx(DataModelSnapshotNotice, {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            onReview: () =>
              gridInteractionActions.runEditorCommand("dataModel"),
          }),
          (calculationState.report.settings.mode === "manual" ||
            calculationState.report.pending > 0 ||
            calculationState.report.execution?.state === "timeout" ||
            calculationState.report.execution?.state === "failed") &&
            i.jsxs("div", {
              className: "lg-notice",
              role: "status",
              "aria-label": text("Calculation status"),
              lang: language("Calculation status"),
              "data-ui-label": "Calculation status",
              children: [
                i.jsxs("span", {
                  children: [
                    calculationState.report.settings.mode === "manual"
                      ? "Manual calculation"
                      : "Calculation",
                    " ",
                    "· ",
                    calculationState.report.pending,
                    " ",
                    i.jsx(LocalizedText, { children: "pending formulas" }),
                    calculationState.report.execution?.state === "timeout" ||
                    calculationState.report.execution?.state === "failed"
                      ? ` · ${calculationState.report.execution.message}`
                      : "",
                  ],
                }),
                i.jsx("button", {
                  onClick: () =>
                    gridInteractionActions.runEditorCommand("calculateNow"),
                  children: i.jsx(LocalizedText, { children: "Calculate Now" }),
                }),
                calculationState.report.execution?.mode === "worker" &&
                  calculationState.report.execution.state === "pending" &&
                  i.jsx("button", {
                    onClick: calculationRun.cancel,
                    children: i.jsx(LocalizedText, {
                      children: "Cancel calculation",
                    }),
                  }),
                i.jsx("button", {
                  onClick: () =>
                    gridInteractionActions.runEditorCommand("calculation"),
                  children: i.jsx(LocalizedText, {
                    children: "Calculation Options",
                  }),
                }),
              ],
            }),
          (operationStatusState.errorNotice ||
            operationStatusState.statusNotice) &&
            i.jsxs("div", {
              className: `lg-notice ${operationStatusState.errorNotice ? "error" : ""}`,
              role: operationStatusState.errorNotice ? "alert" : "status",
              children: [
                i.jsx(LocalizedText, {
                  children:
                    operationStatusState.errorNotice ||
                    operationStatusState.statusNotice,
                }),
                i.jsx("button", {
                  "aria-label": text("Dismiss notification"),
                  onClick: () => {
                    (operationStatusState.setErrorNotice(""),
                      operationStatusState.setStatusNotice(""));
                  },
                  lang: language("Dismiss notification"),
                  "data-ui-label": "Dismiss notification",
                  children: i.jsx(X, { size: 14 }),
                }),
              ],
            }),
          formulaEntryState.activeCellEditRef.current &&
            formulaEntryState.activeCellEditRef.current.locale !==
              workbookMessages.entryLocale &&
            i.jsx("p", {
              role: "status",
              className: "lg-notice",
              children: i.jsx(LocalizedText, {
                children:
                  "The current edit keeps its original numeric locale until you apply or cancel it.",
              }),
            }),
          persistenceState.persistenceError &&
            i.jsxs("div", {
              className: "lg-notice",
              role: "status",
              "aria-label": text("Unsaved workbook changes"),
              lang: language("Unsaved workbook changes"),
              "data-ui-label": "Unsaved workbook changes",
              children: [
                i.jsx("span", {
                  children: i.jsx(LocalizedText, {
                    children: "Your local changes are retained in this editor.",
                  }),
                }),
                persistenceState.persistenceError instanceof
                  WorkbookRevisionConflictError &&
                persistenceState.persistenceError.code === "conflict"
                  ? i.jsx("button", {
                      onClick: () =>
                        persistenceState.setIsSaveConflictDialogOpen(true),
                      children: i.jsx(LocalizedText, {
                        children: "Review save conflict",
                      }),
                    })
                  : i.jsx("button", {
                      disabled:
                        readOnly || persistenceState.pendingSaveCount > 0,
                      onClick: persistenceActions.recoverLocalWorkbook,
                      children: i.jsx(LocalizedText, {
                        children: "Retry saving",
                      }),
                    }),
                persistenceState.pendingSaveCount > 0 &&
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
                      children: "Waiting for save acknowledgement",
                    }),
                  }),
                i.jsx("button", {
                  onClick: persistenceActions.exportLocalChangesJson,
                  children: i.jsx(LocalizedText, {
                    children: "Export my changes",
                  }),
                }),
              ],
            }),
          persistenceState.isSaveConflictDialogOpen &&
            i.jsx(WorkbookSaveConflictDialog, {
              workbook: activeWorkbook,
              loadSaved: () => workbookPersistence.load(),
              onExport: persistenceActions.exportLocalChangesJson,
              onReload: persistenceActions.reloadSavedWorkbook,
              onClose: () =>
                persistenceState.setIsSaveConflictDialogOpen(false),
            }),
          i.jsx(ConnectionStatusNotice, {}),
          i.jsxs("div", {
            className: "lg-workarea",
            children: [
              i.jsx(EditorWorkbookGrid, {
                viewAndSelectionActions,
                cellEditActions,
                navigationActions,
                formulaEntryActions,
                clipboardAndFindActions,
                gridInteractionActions,
                dialogState,
                selectionState,
                navigationState,
                formulaEntryState,
                filterSortState,
                auditState,
                operationStatusState,
                viewDisplayState,
                findReplaceState,
                dragFillState,
                clipboardState,
                readOnly,
                text,
                language,
                editorIdPrefix,
                activeWorkbook,
                activeSheetId,
                activeSheet,
                isEditingLocked,
                selectionAreas,
                filterStateKey,
                activeFilterState,
                sheetSplitPanes,
                splitPaneSizes,
                workbookCalculator,
                filterApplication,
                filterResult,
                isFilterStale,
                latestEditStateRef,
                selectedCellCount,
                activeCellFormulaErrors,
                activeCellReferenceDiagnostics,
                activeCellNamedFormulaDiagnostics,
                selectedNumericValues,
                selectedNumericSum,
                activeCell,
                matchingRowCount,
                renderGridPane,
              }),
              i.jsx(EditorSidePanel, {
                viewAndSelectionActions,
                cellEditActions,
                navigationActions,
                fileAndAnalysisActions,
                gridInteractionActions,
                selectionState,
                formulaEntryState,
                dialogState,
                filterSortState,
                auditState,
                navigationState,
                operationStatusState,
                assistantState,
                adapter,
                text,
                language,
                activeWorkbook,
                activeSheet,
                isEditingLocked,
                selectionAreasKey,
                filterStateKey,
                activeFilterState,
                filterResult,
                latestEditStateRef,
                selectedCellCount,
                assistantSelection,
                selectionReferenceLabel,
              }),
            ],
          }),
        ],
      }),
      dialogState.isWorkbookHealthDialogOpen &&
        i.jsx(WorkbookHealthDialog, {
          monitor: activeReliabilityMonitor,
          onClose: () => dialogState.setIsWorkbookHealthDialogOpen(false),
        }),
      formulaEntryState.validationReviewState &&
        i.jsx(ValidationReviewDialog, {
          error: formulaEntryState.validationReviewState.error,
          readOnly: isEditingLocked,
          onRetry: () => {
            (formulaEntryState.setValidationReviewState(null),
              requestAnimationFrame(() =>
                formulaEntryState.formulaInputRef.current?.focus(),
              ));
          },
          onCancel: () => {
            (formulaEntryState.setValidationReviewState(null),
              operationStatusState.setErrorNotice(""),
              cellEditActions.cancelCellEdit());
          },
          onAccept: () =>
            viewAndSelectionActions.runGuardedAction(() => {
              if (latestEditStateRef.current.readOnly)
                throw Error("This workbook is read-only.");
              if (
                formulaEntryState.validationReviewState.selectionGuard &&
                formulaEntryState.validationReviewState.selectionGuard !==
                  latestEditStateRef.current.selectionKey
              )
                throw Error(
                  "Approved selection changed. Review the proposal again.",
                );
              if (
                formulaEntryState.validationReviewState.snapshot !==
                JSON.stringify(latestEditStateRef.current.book)
              )
                throw Error("Workbook changed. Review the batch again.");
              (persistenceActions.commitWorkbookChange(
                applyCellPatchTransaction(
                  latestEditStateRef.current.book,
                  formulaEntryState.validationReviewState.patches,
                  actor,
                  formulaEntryState.validationReviewState.action,
                  {
                    validationApproval: assertNoStopValidationIssues(
                      formulaEntryState.validationReviewState.error,
                    ),
                  },
                ),
              ),
                formulaEntryState.setValidationReviewState(null),
                (formulaEntryState.activeCellEditRef.current = null),
                formulaEntryState.setCellEditState(null),
                operationStatusState.setStatusNotice(
                  "Validation warnings accepted for the complete batch.",
                ),
                viewAndSelectionActions.focusGridSurface());
            }),
        }),
      dialogState.promptDialogState &&
        i.jsx("div", {
          className: "lg-modal-backdrop",
          onKeyDown: (event) => {
            event.key === "Escape" && dialogState.setPromptDialogState(null);
          },
          onMouseDown: (event) => {
            event.target === event.currentTarget &&
              dialogState.setPromptDialogState(null);
          },
          children: i.jsxs("section", {
            className: "lg-modal",
            role: "dialog",
            "aria-modal": "true",
            "aria-labelledby": `${editorIdPrefix}-dialog-title`,
            children: [
              i.jsx("button", {
                className: "lg-modal-close",
                "aria-label": text("Close dialog"),
                onClick: () => dialogState.setPromptDialogState(null),
                lang: language("Close dialog"),
                "data-ui-label": "Close dialog",
                children: i.jsx(X, { size: 18 }),
              }),
              i.jsx("h2", {
                id: `${editorIdPrefix}-dialog-title`,
                children:
                  dialogState.promptDialogState.kind === "note"
                    ? "Add a cell note"
                    : dialogState.promptDialogState.kind === "rename"
                      ? "Rename workbook"
                      : "Add worksheet",
              }),
              dialogState.promptDialogState.kind === "note"
                ? i.jsx("textarea", {
                    autoFocus: true,
                    "aria-label": text("Cell note"),
                    rows: 5,
                    value: dialogState.promptDialogState.value,
                    onChange: (event) =>
                      dialogState.setPromptDialogState({
                        ...dialogState.promptDialogState,
                        value: event.target.value,
                      }),
                    lang: language("Cell note"),
                    "data-ui-label": "Cell note",
                  })
                : i.jsx("input", {
                    autoFocus: true,
                    "aria-label": text("Name"),
                    value: dialogState.promptDialogState.value,
                    onChange: (event) =>
                      dialogState.setPromptDialogState({
                        ...dialogState.promptDialogState,
                        value: event.target.value,
                      }),
                    onKeyDown: (event) => {
                      (event.key === "Enter" &&
                        gridInteractionActions.commitDialogInput(),
                        event.key === "Escape" &&
                          dialogState.setPromptDialogState(null));
                    },
                    lang: language("Name"),
                    "data-ui-label": "Name",
                  }),
              i.jsxs("div", {
                className: "lg-modal-actions",
                children: [
                  i.jsx("button", {
                    className: "lg-button",
                    onClick: () => dialogState.setPromptDialogState(null),
                    children: i.jsx(LocalizedText, { children: "Cancel" }),
                  }),
                  i.jsx("button", {
                    className: "lg-button dark",
                    onClick: gridInteractionActions.commitDialogInput,
                    children: i.jsx(LocalizedText, { children: "Save" }),
                  }),
                ],
              }),
            ],
          }),
        }),
    ],
  });
}

export { WorkbookSyncBoundary };
