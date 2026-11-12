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

function WorkbookSyncBoundary(e) {
  const r = ne.useMemo(() => createWorkbookDiagnosticsMonitor(), []),
    t = e.reliability ?? r,
    n = ne.useRef({
      propId: e.initialWorkbook?.id,
      emittedId: undefined,
      generation: 0,
    });
  n.current.propId !== e.initialWorkbook?.id &&
    (n.current.emittedId !== e.initialWorkbook?.id && n.current.generation++,
    (n.current.propId = e.initialWorkbook?.id));
  const o = ne.useCallback(
      (c) => {
        ((n.current.emittedId = c.id), e.onChange?.(c));
      },
      [e.onChange],
    ),
    a = ne.useCallback(
      (c) => {
        ((n.current.emittedId = c.id), e.onWorkbookLoaded?.(c));
      },
      [e.onWorkbookLoaded],
    );
  return i.jsx(WorkbookMessagesProvider, {
    options: e.localization,
    children: i.jsx(operationScopeContext.Provider, {
      value: { ...e.fileProcessingOptions, reliability: t },
      children: i.jsx(
        MainWorksheetEditor,
        { ...e, reliability: t, onChange: o, onWorkbookLoaded: a },
        (e.persistenceKey ?? "session") +
          ":" +
          n.current.generation +
          ":" +
          workbookInstanceKey(e.workbookStore),
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
    getEditableCellText = (fe) =>
      formatCellEntryForEdit(fe, workbookMessages.entryLocale),
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
      activeWorkbook.sheets.find((fe) => !fe.hidden).id,
    ),
    activeSheet =
      activeWorkbook.sheets.find(
        (fe) => fe.id === activeSheetId && !fe.hidden,
      ) ?? activeWorkbook.sheets.find((fe) => !fe.hidden),
    workbookPersistence = ne.useMemo(
      () =>
        instrumentWorkbookPersistence(
          workbookStore ??
            (persistenceKey
              ? saveWorkbookWithLock(
                  {
                    getItem: (fe) => localStorage.getItem(fe),
                    setItem: (fe, Oe) => localStorage.setItem(fe, Oe),
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
    let fe = 0,
      Oe;
    const Be = () => {
        if (!formulaEntryState.formulaPointerStateRef.current.dragging || !Oe)
          return;
        const gt = selectionState.gridContainerRef.current;
        if (gt) {
          const jt = gt.getBoundingClientRect();
          (Oe.y < jt.top + 28
            ? (gt.scrollTop -= 24)
            : Oe.y > jt.bottom - 28 && (gt.scrollTop += 24),
            Oe.x < jt.left + 28
              ? (gt.scrollLeft -= 36)
              : Oe.x > jt.right - 28 && (gt.scrollLeft += 36));
          const rr = document
            .elementFromPoint?.(
              Math.max(jt.left + 48, Math.min(jt.right - 8, Oe.x)),
              Math.max(jt.top + 28, Math.min(jt.bottom - 8, Oe.y)),
            )
            ?.closest("[data-address]");
          rr &&
            gt.contains(rr) &&
            formulaEntryState.formulaPointerMoveRef.current(rr.dataset.address);
        }
        fe = requestAnimationFrame(Be);
      },
      Ue = (gt) => {
        formulaEntryState.formulaPointerStateRef.current.dragging &&
          ((Oe = { x: gt.clientX, y: gt.clientY }),
          (fe ||= requestAnimationFrame(Be)));
      },
      Xe = () => {
        const gt = formulaEntryState.formulaPointerStateRef.current.dragging;
        ((formulaEntryState.formulaPointerStateRef.current.dragging = false),
          cancelAnimationFrame(fe),
          (fe = 0),
          (Oe = undefined),
          gt && cellEditActions.restoreFormulaEditorFocus());
      };
    return (
      window.addEventListener("pointermove", Ue),
      window.addEventListener("pointerup", Xe),
      window.addEventListener("pointercancel", Xe),
      () => {
        (cancelAnimationFrame(fe),
          window.removeEventListener("pointermove", Ue),
          window.removeEventListener("pointerup", Xe),
          window.removeEventListener("pointercancel", Xe));
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
      } catch (fe) {
        return {
          result: null,
          error: fe instanceof Error ? fe.message : "Filtering is unavailable.",
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
    const fe = _t(selectionState.activeCellAddress);
    if (
      !rowVisibilityIndex.isHidden(fe.row) &&
      !columnVisibilityIndex.isHidden(fe.col)
    )
      return;
    const Oe =
        rowVisibilityIndex.next(fe.row, 1, true) ??
        rowVisibilityIndex.next(fe.row, -1, true),
      Be =
        columnVisibilityIndex.next(fe.col, 1, true) ??
        columnVisibilityIndex.next(fe.col, -1, true);
    if (Oe === null || Be === null || Oe === undefined || Be === undefined)
      return;
    const Ue = `${columnIndexToLetters(Be)}${Oe + 1}`;
    (selectionState.setActiveCellAddress(Ue),
      selectionState.setSelectionAnchorAddress(Ue),
      (selectionState.activeCellAddressRef.current = Ue),
      viewAndSelectionActions.selectSingleRange(Ue),
      navigationState.setPendingScrollAddress(Ue));
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
    (setRenderRefreshCount((fe) => fe + 1),
      assistantState.setAssistantProposal(null),
      assistantState.setAnswerCitations(null),
      auditState.setAuditAnalysisResult(null));
  }

  const commentedCellAddresses = ne.useMemo(
      () =>
        new Set(
          (activeWorkbook.reviewThreads ?? [])
            .filter(
              (fe) => fe.anchor.sheetId === activeSheet.id && !fe.orphaned,
            )
            .map((fe) => fe.anchor.address),
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
      } catch (fe) {
        return {
          cells: activeSheet.cells,
          warning:
            fe instanceof Error
              ? fe.message
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
      () => (fe) => conditionalFormatEvaluator(fe).style,
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
            (fe) =>
              fe.sheetId === activeSheet.id &&
              fe.address === selectionState.activeCellAddress,
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
            (fe) =>
              fe.sheetId === activeSheet.id &&
              fe.address === selectionState.activeCellAddress,
          )
      ),
      [workbookCalculator, activeSheet.id, selectionState.activeCellAddress],
    ),
    getAllCellEntries = ne.useMemo(() => {
      let fe;
      return () =>
        (fe ??= Object.entries(arrayDisplayResult.cells).map(([Oe, Be]) => ({
          address: Oe,
          cell: Be,
          ..._t(Oe),
        })));
    }, [arrayDisplayResult.cells]),
    selectedNumericValues = ne.useMemo(
      () =>
        (selectedCellCount <= 500
          ? enumerateSelectionCells(selectionAreas).map((fe) => ({
              address: fe,
            }))
          : getAllCellEntries().filter(({ row, col }) =>
              selectionAreaRanges.some((Be) => Tp(Be, row, col)),
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
    selectedNumericSum = selectedNumericValues.reduce((fe, Oe) => fe + Oe, 0),
    activeCell = activeSheet.cells[selectionState.activeCellAddress],
    sortedRowIndexes = ne.useMemo(() => {
      if (
        !filterSortState.filterValueSearch &&
        !filterSortState.activeColumnSort
      )
        return null;
      const fe = new Map();
      for (const { address, row, cell } of getAllCellEntries()) {
        if (
          row < pinnedRowCount ||
          rowVisibilityIndex.isHidden(row) ||
          cell.raw === null
        )
          continue;
        const Nt = fe.get(row) ?? [];
        (Nt.push(address), fe.set(row, Nt));
      }
      const Oe = filterSortState.filterValueSearch.toLowerCase(),
        Be = [...fe.keys()]
          .filter(
            (Ue) =>
              !Oe ||
              fe.get(Ue).some((Xe) =>
                String(workbookCalculator(activeSheet.id, Xe) ?? "")
                  .toLowerCase()
                  .includes(Oe),
              ),
          )
          .sort((Ue, Xe) => Ue - Xe);
      return (
        filterSortState.activeColumnSort &&
          Be.sort((Ue, Xe) => {
            const gt = workbookCalculator(
                activeSheet.id,
                `${columnIndexToLetters(filterSortState.activeColumnSort.col)}${Ue + 1}`,
              ),
              Nt = workbookCalculator(
                activeSheet.id,
                `${columnIndexToLetters(filterSortState.activeColumnSort.col)}${Xe + 1}`,
              );
            return (
              (isFiniteNumber(gt) && isFiniteNumber(Nt)
                ? gt - Nt
                : String(gt ?? "").localeCompare(String(Nt ?? ""))) *
              filterSortState.activeColumnSort.direction
            );
          }),
        Be
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
          ? new Map(sortedRowIndexes.map((fe, Oe) => [fe, Oe]))
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
    yn = columnVisibilityIndex.countBefore(frozenColumnCount),
    frozenRowIndexes = Array.from({ length: frozenRowCount }, (fe, Oe) =>
      rowVisibilityIndex.logicalIndex(Oe),
    ),
    frozenColumnIndexes = Array.from({ length: yn }, (fe, Oe) =>
      columnVisibilityIndex.logicalIndex(Oe),
    ),
    scrollableRowLayout = ne.useMemo(() => {
      if (!sortedRowIndexes)
        return sliceVisibleAxis(rowAxisLayout, frozenRowCount);
      const fe = {};
      for (const [Oe, Be] of Object.entries(effectiveRowHeights)) {
        const Ue = filteredRowPositionMap?.get(Number(Oe));
        Ue !== undefined && Ue >= 0 && (fe[Ue] = Be);
      }
      return buildVisibleAxisLayout(
        sortedRowIndexes.length,
        defaultRowHeight,
        fe,
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
      () => sliceVisibleAxis(columnAxisLayout, yn),
      [columnAxisLayout, yn],
    ),
    matchingRowCount = scrollableRowLayout.count,
    logicalRowAt = (fe) =>
      sortedRowIndexes
        ? sortedRowIndexes[fe]
        : rowVisibilityIndex.logicalIndex(fe + frozenRowCount),
    scrollableRowPositionOf = (fe) => {
      if (!(fe < pinnedRowCount))
        return filteredRowPositionMap
          ? filteredRowPositionMap.get(fe)
          : (rowVisibilityIndex.viewIndex(fe) ?? -1) - frozenRowCount;
    },
    frozenRowsHeight = rowAxisLayout.offset(frozenRowCount),
    frozenColumnsWidth = columnAxisLayout.offset(yn);
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
    yn,
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
    const fe = {};
    for (const Be of Pv(sheetSplitPanes)) {
      const Ue = viewAndSelectionActions.computeGridViewport(
          splitPaneSizes.current.current[Be],
        ),
        Xe = sheetSplitPanes
          ? (rowVisibilityIndex.next(sheetSplitPanes.row, 1, true) ??
            rowVisibilityIndex.next(sheetSplitPanes.row, -1, true) ??
            0)
          : 0,
        gt = sheetSplitPanes
          ? (columnVisibilityIndex.next(sheetSplitPanes.col, 1, true) ??
            columnVisibilityIndex.next(sheetSplitPanes.col, -1, true) ??
            0)
          : 0,
        Nt = Be.startsWith("bottom")
          ? logicalToPhysicalPixels(
              scrollableRowLayout.offset(
                Math.max(0, scrollableRowPositionOf(Xe) ?? 0),
              ),
              scrollableRowLayout.total,
              Ue.rowViewport,
              zoomScale,
            )
          : 0,
        wt = Be.endsWith("right")
          ? logicalToPhysicalPixels(
              scrollableColumnLayout.offset(
                Math.max(0, (columnVisibilityIndex.viewIndex(gt) ?? 0) - yn),
              ),
              scrollableColumnLayout.total,
              Ue.columnViewport,
              zoomScale,
            )
          : 0;
      fe[Be] = { top: Nt, left: wt };
      const jt = splitPaneSizes.refs.current[Be];
      jt && ((jt.scrollTop = Nt), (jt.scrollLeft = wt));
    }
    splitPaneSizes.patch(fe);
    const Oe = Pv(sheetSplitPanes).at(-1);
    (splitPaneSizes.activate(Oe),
      (selectionState.gridContainerRef.current =
        splitPaneSizes.refs.current[Oe] ?? null));
  }, [splitPaneSizesKey]),
    ne.useEffect(() => {
      if (
        !selectionState.gridContainerRef.current ||
        typeof ResizeObserver > "u"
      )
        return;
      const Oe = new ResizeObserver((Be) => {
        const Ue = {};
        for (const gt of Be) {
          const Nt = Number(gt.target.dataset.row),
            wt = gt.target.offsetHeight;
          Number.isInteger(Nt) &&
            activeSheet.rowHeights?.[Nt] === undefined &&
            viewDisplayState.axisResizeState?.axis !== "row" &&
            wt > defaultRowHeight &&
            (Ue[Nt] = wt);
        }
        Object.keys(Ue).length &&
          viewDisplayState.setAutoFitRowHeights((gt) =>
            Object.entries(Ue).every(([Nt, wt]) => gt[Number(Nt)] === wt)
              ? gt
              : { ...gt, ...Ue },
          );
      });
      return (
        Object.values(splitPaneSizes.refs.current).forEach((Be) =>
          Be.querySelectorAll("tr[data-row]").forEach((Ue) => Oe.observe(Ue)),
        ),
        () => Oe.disconnect()
      );
    }),
    ne.useEffect(() => {
      if (
        !navigationState.pendingScrollAddress ||
        !selectionState.gridContainerRef.current
      )
        return;
      const fe = _t(navigationState.pendingScrollAddress),
        Oe = selectionState.gridContainerRef.current;
      if (
        rowVisibilityIndex.isHidden(fe.row) ||
        columnVisibilityIndex.isHidden(fe.col)
      ) {
        navigationState.setPendingScrollAddress(null);
        return;
      }
      const Be = scrollableRowPositionOf(fe.row);
      if (Be === undefined && fe.row >= pinnedRowCount) {
        (filterSortState.setFilterValueSearch(""),
          filterSortState.setActiveColumnSort(null));
        return;
      }
      let Ue = activePaneScrollTop,
        Xe = activePaneScrollLeft;
      if (Be !== undefined && Be >= 0) {
        const Nt = scrollableRowLayout.offset(Be),
          wt = scrollableRowLayout.offset(Be + 1);
        let jt = rowWindow.logicalOffset;
        (Nt < jt
          ? (jt = Nt)
          : wt > jt + rowViewport / zoomScale &&
            (jt = wt - rowViewport / zoomScale),
          (Ue = logicalToPhysicalPixels(
            jt,
            scrollableRowLayout.total,
            rowViewport,
            zoomScale,
          )));
      }
      const gt = (columnVisibilityIndex.viewIndex(fe.col) ?? 0) - yn;
      if (gt >= 0) {
        const Nt = scrollableColumnLayout.offset(gt),
          wt = scrollableColumnLayout.offset(gt + 1);
        let jt = columnWindow.logicalOffset;
        (Nt < jt
          ? (jt = Nt)
          : wt > jt + columnViewport / zoomScale &&
            (jt = wt - columnViewport / zoomScale),
          (Xe = logicalToPhysicalPixels(
            jt,
            scrollableColumnLayout.total,
            columnViewport,
            zoomScale,
          )));
      }
      ((Oe.scrollTop = Ue),
        (Oe.scrollLeft = Xe),
        viewAndSelectionActions.scrollActivePaneVertically(Ue),
        viewAndSelectionActions.scrollActivePaneHorizontally(Xe),
        navigationState.setPendingScrollAddress(null));
    }, [
      navigationState.pendingScrollAddress,
      filteredRowPositionMap,
      scrollableRowLayout,
      scrollableColumnLayout,
      pinnedRowCount,
      yn,
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
      const fe = { active: true, controllers: new Set() };
      persistenceState.persistenceSessionRef.current = fe;
      try {
        const Oe = workbookPersistence.load();
        if (
          ((persistenceState.persistedWorkbookRecordRef.current = Oe),
          persistenceState.setPersistedWorkbookRecord(Oe),
          Oe)
        ) {
          (setWorkbookSnapshot(buildCanonicalWorkbookSnapshot(Oe.workbook)),
            (latestEditStateRef.current.book = Oe.workbook),
            viewAndSelectionActions.applyWorkbookView(Oe.view),
            operationStatusState.setStatusNotice(
              Oe.warning ?? "Restored your locally saved workbook.",
            ),
            persistenceActions.reportSaveStatus("saved", Oe.workbook));
          try {
            viewAndSelectionActions.notifyHostCallback("loaded", () =>
              onWorkbookLoaded?.(Oe.workbook),
            );
          } catch {
            operationStatusState.setErrorNotice(
              "The host could not acknowledge the loaded workbook.",
            );
          }
        }
        persistenceState.isWorkbookLoadedRef.current = true;
      } catch (Oe) {
        (persistenceActions.reportSaveError(Oe),
          (persistenceState.persistenceMessageRef.current =
            "The saved workbook could not be restored. Import a backup to recover it. Existing saved data will be preserved."),
          operationStatusState.setErrorNotice(
            persistenceState.persistenceMessageRef.current,
          ));
      }
      return () => {
        fe.active = false;
        for (const Ue of fe.controllers) Ue.abort();
        const Oe = persistenceState.persistenceSessionRef.current;
        Oe.active = false;
        for (const Ue of Oe.controllers) Ue.abort();
        persistenceState.saveAbortControllerRef.current?.abort();
      };
    }, [workbookPersistence]),
    ne.useEffect(() => {
      if (readOnly) {
        for (const Oe of persistenceState.persistenceSessionRef.current
          .controllers)
          Oe.abort();
        persistenceState.saveAbortControllerRef.current?.abort();
      }
    }, [readOnly]),
    ne.useEffect(() => {
      if (!persistenceKey && !("coordination" in workbookPersistence)) return;
      const fe = (Oe) => {
        const Be = persistenceState.persistedWorkbookRecordRef.current?.guard,
          Ue = persistenceKey ?? Be?.namespace;
        if (!(!Ue || (Oe.key !== null && Oe.key !== Ue)))
          try {
            const gt = workbookPersistence.load();
            if (
              JSON.stringify(gt?.guard ?? null) === JSON.stringify(Be ?? null)
            )
              return;
            persistenceActions.reportSaveError(
              new WorkbookRevisionConflictError(
                "conflict",
                "Another workbook window changed the saved version. Your local work has been retained.",
                gt?.guard,
              ),
            );
          } catch (gt) {
            persistenceActions.reportSaveError(gt);
          }
      };
      return (
        window.addEventListener("storage", fe),
        () => window.removeEventListener("storage", fe)
      );
    }, [workbookPersistence, persistenceKey]),
    ne.useEffect(() => {
      if (!persistenceKey && !workbookStorePersistent) return;
      const fe = (Oe) => {
        (formulaEntryState.activeCellEditRef.current ||
          persistenceState.persistenceSessionRef.current.controllers.size ||
          JSON.stringify(
            persistenceState.persistedWorkbookRecordRef.current?.workbook,
          ) !== JSON.stringify(latestEditStateRef.current.book)) &&
          (Oe.preventDefault(), (Oe.returnValue = ""));
      };
      return (
        window.addEventListener("beforeunload", fe),
        () => window.removeEventListener("beforeunload", fe)
      );
    }, [persistenceKey, workbookStorePersistent]),
    ne.useEffect(() => {
      if (!persistenceState.isWorkbookLoadedRef.current || readOnly) return;
      const fe = setTimeout(() => {
        try {
          const Oe = latestEditStateRef.current;
          if (
            Oe.readOnly ||
            persistenceState.persistenceError ||
            persistenceState.persistenceSessionRef.current.controllers.size
          )
            return;
          const Be = viewAndSelectionActions.buildWorkbookViewState(
              Oe.book,
              Oe.sheetId,
              Oe.range,
            ),
            Ue = persistenceState.persistedWorkbookRecordRef.current;
          if (
            Ue &&
            JSON.stringify(Ue.view) === JSON.stringify(Be) &&
            (Ue.workbook === Oe.book ||
              JSON.stringify(Ue.workbook) === JSON.stringify(Oe.book))
          )
            return;
          const Xe = persistenceActions.persistWorkbookSnapshot(Oe.book, Be);
          isPromiseLike(Xe) && Xe.catch(() => {});
        } catch (Oe) {
          operationStatusState.setErrorNotice(
            `${Oe instanceof Error ? Oe.message : "The current selection could not be saved."} Changes remain in this editor; export JSON to keep a copy.`,
          );
        }
      }, 400);
      return () => clearTimeout(fe);
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
      let fe;
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
        const Oe = validateProposedCalculation(
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
            Oe,
            externalProposal.context,
          ),
        ),
          assistantState.setAssistantProposal(Oe),
          dialogState.setSidePanelKind("assistant"),
          operationStatusState.setErrorNotice(""),
          (fe = { accepted: true }));
      } catch (Oe) {
        const Be =
          Oe instanceof Error ? Oe.message : "Invalid external proposal.";
        (operationStatusState.setErrorNotice(Be),
          (fe = { accepted: false, error: Be }));
      }
      try {
        viewAndSelectionActions.notifyHostCallback("proposal", () =>
          onExternalProposalStaged?.(externalProposal.requestId, fe),
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

  formulaEntryState.formulaPointerMoveRef.current = (fe) => {
    formulaEntryState.formulaPointerStateRef.current.head !== fe &&
      formulaEntryActions.insertFormulaReference(fe, true);
  };

  ne.useEffect(() => {
    if (!viewDisplayState.axisResizeState) return;
    const fe = (Xe) =>
        viewDisplayState.setAxisResizeState(
          (gt) =>
            gt && {
              ...gt,
              size: Math.round(
                Math.max(
                  gt.axis === "column" ? minColumnWidth : minRowHeight,
                  Math.min(
                    gt.axis === "column" ? maxColumnWidth : maxRowHeight,
                    gt.initial +
                      ((gt.axis === "column" ? Xe.clientX : Xe.clientY) -
                        gt.start) /
                        zoomScale,
                  ),
                ),
              ),
            },
        ),
      Oe = () => {
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
      Be = () => viewDisplayState.setAxisResizeState(null),
      Ue = (Xe) => {
        Xe.key === "Escape" &&
          (!Xe.defaultPrevented &&
            Xe.target instanceof Element &&
            Xe.target.closest(".lg-app") ===
              selectionState.appRootRef.current &&
            Xe.preventDefault(),
          Be());
      };
    return (
      window.addEventListener("pointermove", fe),
      window.addEventListener("pointerup", Oe),
      window.addEventListener("pointercancel", Be),
      window.addEventListener("keydown", Ue),
      () => {
        (window.removeEventListener("pointermove", fe),
          window.removeEventListener("pointerup", Oe),
          window.removeEventListener("pointercancel", Be),
          window.removeEventListener("keydown", Ue));
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
    const fe = () => {
        ((fillDragStateRef.current = null),
          dragFillState.setFillDragState(null));
      },
      Oe = () => {
        const Ue = fillDragStateRef.current;
        if ((fe(), !(!Ue || Ue.targetRange === Ue.sourceRange))) {
          if (
            latestEditStateRef.current.readOnly ||
            latestEditStateRef.current.book !== Ue.book ||
            latestEditStateRef.current.sheetId !== Ue.sheetId
          ) {
            operationStatusState.setErrorNotice(
              "The workbook changed while filling. Select the cells and try again.",
            );
            return;
          }
          dragFillState.setFillPreviewState(Ue);
        }
      },
      Be = (Ue) => {
        Ue.key === "Escape" &&
          (!Ue.defaultPrevented &&
            Ue.target instanceof Element &&
            Ue.target.closest(".lg-app") ===
              selectionState.appRootRef.current &&
            Ue.preventDefault(),
          fe());
      };
    return (
      window.addEventListener("pointerup", Oe),
      window.addEventListener("pointercancel", fe),
      window.addEventListener("keydown", Be),
      () => {
        (window.removeEventListener("pointerup", Oe),
          window.removeEventListener("pointercancel", fe),
          window.removeEventListener("keydown", Be));
      }
    );
  }, [!!dragFillState.fillDragState]);

  const frozenColumnLeftOffsets = Object.fromEntries(
      frozenColumnIndexes.map((fe, Oe) => [
        fe,
        rowHeaderWidth + columnAxisLayout.offset(Oe),
      ]),
    ),
    renderColumnHeader = (fe, Oe = false) =>
      i.jsxs(
        "th",
        {
          role: "columnheader",
          scope: "col",
          "data-column": fe,
          "aria-colindex": fe + 2,
          className: `${_t(selectionState.activeCellAddress).col === fe ? "lg-selected-header" : ""} ${Oe ? "lg-frozen-column" : ""}`,
          style: Oe
            ? {
                position: "sticky",
                left: frozenColumnLeftOffsets[fe],
                zIndex: 7,
              }
            : undefined,
          onClick: (Be) => {
            formulaEntryActions.insertFormulaReference(
              `${columnIndexToLetters(fe)}1`,
              false,
              {
                range: `${columnIndexToLetters(fe)}:${columnIndexToLetters(fe)}`,
                bounds: {
                  firstRow: 0,
                  lastRow: activeSheet.rowCount - 1,
                  firstCol: fe,
                  lastCol: fe,
                },
              },
            ) || navigationActions.selectRowOrColumn("column", fe, Be.shiftKey);
          },
          children: [
            columnIndexToLetters(fe),
            i.jsx("span", {
              className: "lg-column-resize",
              role: "separator",
              "aria-label": `Resize column ${columnIndexToLetters(fe)}`,
              "aria-orientation": "vertical",
              "aria-disabled": isEditingLocked,
              title: text(
                "Drag to resize; double-click to AutoFit column width",
              ),
              onPointerDown: (Be) =>
                rowColumnAndFillActions.beginAxisResize("column", fe, Be),
              onDoubleClick: (Be) => {
                (Be.stopPropagation(),
                  rowColumnAndFillActions.autoFitSelectedAxis("column", fe));
              },
              onClick: (Be) => Be.stopPropagation(),
              lang: language(
                "Drag to resize; double-click to AutoFit column width",
              ),
            }),
          ],
        },
        fe,
      ),
    renderGridRow = (fe, Oe, Be, Ue) =>
      i.jsx(
        WorksheetGridRow,
        {
          idPrefix: `${editorIdPrefix}-${Ue}`,
          row: fe,
          pinned: Oe,
          pinnedTop:
            (viewDisplayState.isHeadingsVisible ? 22 : 0) +
            rowAxisLayout.offset(rowVisibilityIndex.viewIndex(fe) ?? 0),
          frozenColumns: frozenColumnIndexes,
          frozenColumnOffsets: frozenColumnLeftOffsets,
          headings: viewDisplayState.isHeadingsVisible,
          columns: Be.shownColumns,
          leftGap: Be.leftGap,
          rightGap: Be.rightGap,
          minimumHeight: effectiveRowHeights[fe] ?? defaultRowHeight,
          fixedHeight:
            activeSheet.rowHeights?.[fe] !== undefined ||
            (viewDisplayState.axisResizeState?.axis === "row" &&
              viewDisplayState.axisResizeState.index === fe),
          onRowSelect: (Nt) => {
            formulaEntryActions.insertFormulaReference(`A${fe + 1}`, false, {
              range: `${fe + 1}:${fe + 1}`,
              bounds: {
                firstRow: fe,
                lastRow: fe,
                firstCol: 0,
                lastCol: activeSheet.colCount - 1,
              },
            }) || navigationActions.selectRowOrColumn("row", fe, Nt);
          },
          onResize: (Nt) =>
            rowColumnAndFillActions.beginAxisResize("row", fe, Nt),
          onAutoFit: () =>
            rowColumnAndFillActions.autoFitSelectedAxis("row", fe),
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
          onDragStart: (Nt, wt) => {
            ((dragFillState.dragAnchorRef.current = Nt),
              (dragFillState.dragBaseSelectionsRef.current = wt
                ? [...selectionAreas]
                : []));
          },
          onDragEnter: (Nt, wt) => {
            if (
              formulaEntryState.formulaPointerStateRef.current.dragging &&
              wt === 1
            ) {
              formulaEntryActions.insertFormulaReference(Nt, true);
              return;
            }
            fillDragStateRef.current && wt === 1
              ? rowColumnAndFillActions.extendFillDragTo(Nt)
              : dragFillState.dragAnchorRef.current &&
                wt === 1 &&
                ((selectionState.activeCellAddressRef.current = Nt),
                viewAndSelectionActions.runGuardedAction(() =>
                  viewAndSelectionActions.applySelectionAreas([
                    ...dragFillState.dragBaseSelectionsRef.current,
                    wy(dragFillState.dragAnchorRef.current, Nt),
                  ]),
                ),
                selectionState.setActiveCellAddress(
                  dragFillState.dragAnchorRef.current,
                ));
          },
          onNavigate: navigationActions.moveSelectionAfterCommit,
          onPivotDrill: (Nt) => {
            const wt = _t(Nt);
            return activeSheet.pivots?.some((rr) => {
              const Mr = parseRangeAddress(rr.outputRange);
              return (
                wt.row > Mr.firstRow &&
                wt.row <= Mr.lastRow &&
                wt.col >= Mr.firstCol + Math.max(1, rr.rowFields.length) &&
                wt.col <= Mr.lastCol
              );
            })
              ? (navigationActions.selectCellAddress(Nt),
                gridInteractionActions.runEditorCommand("pivotExplore"),
                true)
              : false;
          },
          commentAddresses: commentedCellAddresses,
          onReview: (Nt) => {
            (navigationActions.selectCellAddress(Nt),
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
          displayedRows: Oe ? frozenRowIndexes : Be.shownRows,
          filterHeader:
            filterResult &&
            fe === parseRangeAddress(filterResult.context.range).firstRow
              ? filterResult.fields
              : undefined,
          filteredColumns:
            filterResult?.context.definition.columns.map((Nt) => Nt.column) ??
            [],
          onColumnFilter: (Nt) => {
            (gridInteractionActions.runEditorCommand("populationFilter"),
              filterSortState.setActiveFilterColumn(Nt));
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
            splitPaneSizes.activeId === Ue &&
            (!formulaEntryState.activeCellEditRef.current ||
              formulaEntryState.activeCellEditRef.current.sheetId ===
                activeSheet.id)
              ? formulaEntryState.cellEditState
              : null,
          setEditing: (Nt) =>
            Nt === null
              ? cellEditActions.cancelCellEdit()
              : cellEditActions.beginCellEdit(Nt, true),
          select: (Nt, wt, jt) => {
            (navigationActions.selectCellAddress(Nt, wt, jt),
              selectionState.gridContainerRef.current?.focus({
                preventScroll: true,
              }));
          },
          saveCell: cellEditActions.commitCellEdit,
          readOnly: isEditingLocked,
        },
        fe,
      ),
    renderVirtualSpacer = (fe, Oe, Be) =>
      fe > 0
        ? i.jsx(
            "tr",
            {
              "aria-hidden": "true",
              className: "lg-virtual-spacer",
              style: { height: fe },
              children: i.jsx("td", { colSpan: Be, style: { height: fe } }),
            },
            Oe,
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
    yn,
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

  const closeContextMenu = ne.useCallback((fe) => {
      (selectionState.setContextMenuState(null),
        fe &&
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
  } catch (fe) {
    selectionReadError = fe.message;
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
            viewAndSelectionActions.getSelectedCellAddresses().map((fe) => ({
              sheetId: activeSheet.id,
              address: fe,
              cell: { ...activeSheet.cells[fe], raw: null },
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
    onKeyDown: (fe) => {
      if (
        fe.key !== "F6" ||
        fe.defaultPrevented ||
        fe.ctrlKey ||
        fe.metaKey ||
        fe.altKey ||
        fe.nativeEvent.isComposing ||
        dialogState.activeDialogMode ||
        dialogState.promptDialogState ||
        (formulaEntryState.activeCellEditRef.current &&
          (formulaEntryState.activeCellEditRef.current.value !==
            formulaEntryState.activeCellEditRef.current.originalText ||
            !cellEditActions.commitCellEdit()))
      )
        return;
      fe.preventDefault();
      const Oe = [
          selectionState.appRootRef.current?.querySelector(
            '[role="tab"][aria-selected="true"]',
          ),
          formulaEntryState.formulaInputRef.current,
          selectionState.gridContainerRef.current,
          selectionState.appRootRef.current?.querySelector(
            '[data-sheet-id][aria-current="true"]',
          ),
        ].filter((jt) => !!jt),
        Be = document.activeElement,
        Ue = Be?.closest(".xl-ribbon")
          ? 0
          : Be?.closest(".lg-sheet-tabs")
            ? 3
            : Oe.findIndex((jt) => jt === Be || jt.contains(Be)),
        Xe = fe.shiftKey ? -1 : 1;
      Oe[
        Ue < 0
          ? fe.shiftKey
            ? Oe.length - 1
            : 0
          : (Ue + Xe + Oe.length) % Oe.length
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
          onReport: (fe) =>
            saveBlobToFile(
              new Blob([JSON.stringify(fe, null, 2)], {
                type: "application/json",
              }),
              `${fe.fileName}.schema.json`,
            ),
          onImport: (fe, Oe) => {
            if (
              latestEditStateRef.current.book !==
              importExportState.delimitedImportExportSource.snapshot
            )
              throw new Error("The workbook changed. Review the file again.");
            if (latestEditStateRef.current.readOnly)
              throw new Error("This workbook is read-only.");
            (persistenceActions.commitWorkbookChange(fe),
              navigationActions.switchActiveWorksheet(fe.sheets[0].id),
              operationStatusState.setStatusNotice(
                `Imported ${Oe.records} records using the reviewed schema. Formula-like source text remains literal.`,
              ),
              importExportState.setDelimitedImportExportSource(null));
          },
          onExport: (fe, Oe) => {
            if (
              latestEditStateRef.current.book !==
              importExportState.delimitedImportExportSource.snapshot
            )
              throw new Error("The workbook changed. Review the export again.");
            if (
              Oe.population &&
              !("bytes" in importExportState.delimitedImportExportSource)
            ) {
              const Be = exportWorksheetToCsv(
                latestEditStateRef.current.book,
                importExportState.delimitedImportExportSource.sheetId,
                Oe.options,
              );
              if (
                JSON.stringify(Be.report.population) !==
                  JSON.stringify(Oe.population) ||
                Be.bytes.length !== fe.length ||
                Be.bytes.some((Ue, Xe) => Ue !== fe[Xe])
              )
                throw Error(
                  "Calculated values or population membership changed. Preview the export again.",
                );
            }
            (saveBlobToFile(
              new Blob([Uint8Array.from(fe)], {
                type: "application/octet-stream",
              }),
              Oe.fileName,
            ),
              operationStatusState.setStatusNotice(
                `Delimited data exported; ${Oe.safeTextChanges} formula-like text values neutralized under the reviewed safe-export policy.`,
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
              const fe = requireRetainedXlsxSource(
                importExportState.xlsxConversionReview.imported ??
                  importExportState.xlsxConversionReview.snapshot,
              );
              saveBlobToFile(
                new Blob([new Uint8Array(fe.bytes)], {
                  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                }),
                fe.fileName,
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
                      (fe) => !fe.hidden,
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
            onChange: (fe) =>
              void fileAndAnalysisActions.importWorkbookFromFile(
                fe.target.files?.[0],
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
          onKeyDown: (fe) => {
            fe.key === "Escape" && dialogState.setPromptDialogState(null);
          },
          onMouseDown: (fe) => {
            fe.target === fe.currentTarget &&
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
                    onChange: (fe) =>
                      dialogState.setPromptDialogState({
                        ...dialogState.promptDialogState,
                        value: fe.target.value,
                      }),
                    lang: language("Cell note"),
                    "data-ui-label": "Cell note",
                  })
                : i.jsx("input", {
                    autoFocus: true,
                    "aria-label": text("Name"),
                    value: dialogState.promptDialogState.value,
                    onChange: (fe) =>
                      dialogState.setPromptDialogState({
                        ...dialogState.promptDialogState,
                        value: fe.target.value,
                      }),
                    onKeyDown: (fe) => {
                      (fe.key === "Enter" &&
                        gridInteractionActions.commitDialogInput(),
                        fe.key === "Escape" &&
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
