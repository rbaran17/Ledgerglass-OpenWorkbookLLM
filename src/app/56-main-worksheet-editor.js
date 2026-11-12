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
import { renderWorkbookGrid } from "./56-main-worksheet-editor/render-workbook-grid.js";
import { renderSidePanel } from "./56-main-worksheet-editor/render-side-panel.js";
import { WorksheetGridRow } from "./56-main-worksheet-editor/worksheet-grid-row.js";
import { renderEditorDialogs } from "./56-main-worksheet-editor/render-dialogs.js";
import { renderEditorRibbon } from "./56-main-worksheet-editor/render-ribbon.js";

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
} from "./02-financial-functions.js";
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
  createEmptyAssistantConversation,
  LocalizedText,
  formatCellEntryForEdit,
  formatLocalizedNumberDisplay,
} from "./04-cell-input-paste-and-fill.js";
import { CustomFunctionReviewBanner } from "./05-add-ins-and-custom-functions.js";

import { exportWorksheetToCsv } from "./08-xlsx-export-and-preservation.js";
import {
  useOperationScope,
  WorkbookLinkStatusBadge,
  operationScopeContext,
} from "./09-workbook-links.js";
import { XlsxConversionReview } from "./10-source-import-dialog.js";
import {
  WorkbookRevisionConflictError,
  createWorkbookSession,
} from "./11-workbook-formats-dialog.js";
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
} from "./20-calculation-options-and-sources.js";

import { ConnectionStatusNotice } from "./22-themes-and-cell-styles.js";

import { useCalculationRun } from "./24-merge-and-audit-analysis.js";

import { AnalysisSnapshotNotice } from "./38-data-analysis-dialog.js";

import { DelimitedImportExportDialog } from "./40-what-if-analysis.js";

import { FillPreviewDialog } from "./47-fill-and-paste-special.js";
import { MoveCellsDialog } from "./48-format-cells-and-move.js";

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
        var h;
        ((n.current.emittedId = c.id),
          (h = e.onChange) == null || h.call(e, c));
      },
      [e.onChange],
    ),
    a = ne.useCallback(
      (c) => {
        var h;
        ((n.current.emittedId = c.id),
          (h = e.onWorkbookLoaded) == null || h.call(e, c));
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
  var sheetProtectionDetails;
  const defaultDiagnosticsMonitor = ne.useMemo(
      () => createWorkbookDiagnosticsMonitor(),
      [],
    ),
    activeReliabilityMonitor = reliability ?? defaultDiagnosticsMonitor,
    [isWorkbookHealthDialogOpen, setIsWorkbookHealthDialogOpen] =
      ne.useState(false),
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
        workbookSnapshot.extensions || (extensions != null && extensions.length)
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
    persistedWorkbookRecordRef = ne.useRef(null),
    [persistedWorkbookRecord, setPersistedWorkbookRecord] = ne.useState(null),
    isWorkbookLoadedRef = ne.useRef(false),
    isEditableRef = ne.useRef(!readOnly);
  isEditableRef.current = !readOnly;
  const persistenceSessionRef = ne.useRef({
      active: false,
      controllers: new Set(),
    }),
    saveAbortControllerRef = ne.useRef(null),
    setSaveAbortControllerRef = ne.useRef(onPersistenceStateChange);
  setSaveAbortControllerRef.current = onPersistenceStateChange;
  const [persistenceError, setPersistenceError] = ne.useState(null),
    persistenceMessageRef = ne.useRef(null),
    [isSaveConflictDialogOpen, setIsSaveConflictDialogOpen] =
      ne.useState(false),
    [pendingSaveCount, setPendingSaveCount] = ne.useState(0),
    isEditingLocked =
      readOnly ||
      !!(
        (sheetProtectionDetails = activeSheet.protection) != null &&
        sheetProtectionDetails.enabled
      ),
    [activeDialogMode, setActiveDialogMode] = ne.useState(null),
    [automationInspection, setAutomationInspection] = ne.useState(),
    [contextMenuState, setContextMenuState] = ne.useState(null),
    appRootRef = ne.useRef(null);
  useDialogFocusTrap(appRootRef);
  const [tabBarSheetId, setTabBarSheetId] = ne.useState(null),
    formulaInputRef = ne.useRef(null),
    formulaSuggestionListRef = ne.useRef(null),
    [formulaCaretPosition, setFormulaCaretPosition] = ne.useState(0),
    [isFormulaBarFocused, setIsFormulaBarFocused] = ne.useState(false),
    [activeCellAddress, setActiveCellAddress] = ne.useState("A1"),
    [selectionAnchorAddress, setSelectionAnchorAddress] = ne.useState("A1"),
    activeCellAddressRef = ne.useRef("A1"),
    [activeSelectionRange, setActiveSelectionRange] = ne.useState("A1"),
    [completedSelectionAreas, setCompletedSelectionAreas] = ne.useState([]);

  const selectionAreas = ne.useMemo(
      () => [...completedSelectionAreas, activeSelectionRange],
      [completedSelectionAreas, activeSelectionRange],
    ),
    selectionAreaRanges = ne.useMemo(
      () => selectionAreas.map(parseRangeAddress),
      [selectionAreas],
    ),
    selectionAreasKey = selectionAreas.join(", ");

  const [goToReferenceText, setGoToReferenceText] = ne.useState("A1"),
    [cellEditText, setCellEditText] = ne.useState(""),
    [cellEditState, setCellEditState] = ne.useState(null),
    activeCellEditRef = ne.useRef(null),
    formulaPointerStateRef = ne.useRef({
      input: null,
      start: 0,
      end: 0,
      dragging: false,
      mode: true,
    }),
    [activeFormulaReference, setActiveFormulaReference] = ne.useState(null),
    [isFormulaPointMode, setIsFormulaPointMode] = ne.useState(true),
    pendingCrossSheetTargetRef = ne.useRef(null),
    formulaPointerMoveRef = ne.useRef(() => {});
  (ne.useEffect(() => {
    let fe = 0,
      Oe;
    const Be = () => {
        if (!formulaPointerStateRef.current.dragging || !Oe) return;
        const gt = gridContainerRef.current;
        if (gt) {
          const jt = gt.getBoundingClientRect();
          (Oe.y < jt.top + 28
            ? (gt.scrollTop -= 24)
            : Oe.y > jt.bottom - 28 && (gt.scrollTop += 24),
            Oe.x < jt.left + 28
              ? (gt.scrollLeft -= 36)
              : Oe.x > jt.right - 28 && (gt.scrollLeft += 36));
          const rr = document.elementFromPoint
            ?.call(
              document,
              Math.max(jt.left + 48, Math.min(jt.right - 8, Oe.x)),
              Math.max(jt.top + 28, Math.min(jt.bottom - 8, Oe.y)),
            )
            ?.closest("[data-address]");
          rr &&
            gt.contains(rr) &&
            formulaPointerMoveRef.current(rr.dataset.address);
        }
        fe = requestAnimationFrame(Be);
      },
      Ue = (gt) => {
        formulaPointerStateRef.current.dragging &&
          ((Oe = { x: gt.clientX, y: gt.clientY }),
          fe || (fe = requestAnimationFrame(Be)));
      },
      Xe = () => {
        const gt = formulaPointerStateRef.current.dragging;
        ((formulaPointerStateRef.current.dragging = false),
          cancelAnimationFrame(fe),
          (fe = 0),
          (Oe = undefined),
          gt && restoreFormulaEditorFocus());
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
      (activeCellEditRef.current || resetFormulaEditSession(),
        onCellEditPendingChange == null ||
          onCellEditPendingChange(!!activeCellEditRef.current));
    }, [cellEditText, cellEditState, activeWorkbook, onCellEditPendingChange]));
  const [sidePanelKind, setSidePanelKind] = ne.useState(null),
    [filterValueSearch, setFilterValueSearch] = ne.useState(""),
    [filterStateByScope, setFilterStateByScope] = ne.useState({}),
    [activeFilterColumn, setActiveFilterColumn] = ne.useState(),
    [auditPopulationScope, setAuditPopulationScope] = ne.useState("full"),
    filterStateKey = JSON.stringify([activeWorkbook.id, activeSheet.id]),
    activeFilterState = filterStateByScope[filterStateKey],
    [activeColumnSort, setActiveColumnSort] = ne.useState(null),
    [isShowingFormulas, setIsShowingFormulas] = ne.useState(false),
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

  const [pendingScrollAddress, setPendingScrollAddress] = ne.useState(null),
    [undoStack, setUndoStack] = ne.useState([]),
    [redoStack, setRedoStack] = ne.useState([]),
    [statusNotice, setStatusNotice] = ne.useState(""),
    [errorNotice, setErrorNotice] = ne.useState(""),
    [isOperationBusy, setIsOperationBusy] = ne.useState(false),
    operationScope = useOperationScope(),
    latestOperationIdRef = ne.useRef(0),
    activeOperationControllerRef = ne.useRef(null);
  ne.useEffect(
    () => () => {
      var fe;
      (latestOperationIdRef.current++,
        (fe = activeOperationControllerRef.current) == null || fe.abort());
    },
    [],
  );
  const [delimitedImportExportSource, setDelimitedImportExportSource] =
      ne.useState(null),
    [pendingImportFile, setPendingImportFile] = ne.useState(),
    [xlsxConversionReview, setXlsxConversionReview] = ne.useState(null),
    [selectedAuditKind, setSelectedAuditKind] = ne.useState("duplicates"),
    [auditAnalysisResult, setAuditAnalysisResult] = ne.useState(null),
    [materialityThreshold, setMaterialityThreshold] = ne.useState(25e3),
    [auditSampleSize, setAuditSampleSize] = ne.useState(5),
    [samplingSeed, setSamplingSeed] = ne.useState("FY25-01"),
    [agingReportPayload, setAgingReportPayload] = ne.useState(null),
    [agingAsOfDate, setAgingAsOfDate] = ne.useState("2025-12-31"),
    [auditSourceRange, setAuditSourceRange] = ne.useState(""),
    [assistantSession, setAssistantSession] = ne.useState(
      createEmptyAssistantConversation,
    ),
    [assistantProposal, setAssistantProposal] = ne.useState(null),
    [validationReviewState, setValidationReviewState] = ne.useState(null),
    [answerCitations, setAnswerCitations] = ne.useState(null),
    [isAssistantBusy, setIsAssistantBusy] = ne.useState(false),
    [promptDialogState, setPromptDialogState] = ne.useState(null),
    [isGridlinesVisible, setIsGridlinesVisible] = ne.useState(true),
    [isHeadingsVisible, setIsHeadingsVisible] = ne.useState(true),
    [zoomPercent, setZoomPercent] = ne.useState(100),
    [isFilterBarVisible, setIsFilterBarVisible] = ne.useState(false),
    [isFindBarVisible, setIsFindBarVisible] = ne.useState(false),
    [findQuery, setFindQuery] = ne.useState(""),
    [replaceWithText, setReplaceWithText] = ne.useState(""),
    dragAnchorRef = ne.useRef(null),
    dragBaseSelectionsRef = ne.useRef([]),
    [fillDragState, setFillDragState] = ne.useState(null),
    fillDragStateRef = ne.useRef(fillDragState);
  fillDragStateRef.current = fillDragState;
  const [fillPreviewState, setFillPreviewState] = ne.useState(null),
    [axisResizeState, setAxisResizeState] = ne.useState(null),
    internalClipboardRef = ne.useRef(null),
    [clipboardSnapshotState, setClipboardSnapshotState] = ne.useState(null),
    [moveCellsSource, setMoveCellsSource] = ne.useState(null);

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
      filterValueSearch || activeColumnSort ? 1 : 0,
      filterResult &&
        parseRangeAddress(filterResult.context.range).firstRow === 0
        ? parseRangeAddress(filterResult.context.range).firstRow + 1
        : 0,
    ),
    frozenColumnCount = activeSheet.frozenCols ?? 0;
  ne.useEffect(() => {
    const fe = _t(activeCellAddress);
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
    (setActiveCellAddress(Ue),
      setSelectionAnchorAddress(Ue),
      (activeCellAddressRef.current = Ue),
      selectSingleRange(Ue),
      setPendingScrollAddress(Ue));
  }, [activeCellAddress, rowVisibilityIndex, columnVisibilityIndex]);
  const defaultRowHeight = defaultRowHeightOf(activeSheet),
    defaultColumnWidth = defaultColumnWidthOf(activeSheet),
    zoomScale = zoomPercent / 100,
    rowHeaderWidth = isHeadingsVisible ? 58 : 0,
    importFileInputRef = ne.useRef(null),
    gridContainerRef = ne.useRef(null),
    findInputRef = ne.useRef(null),
    textMeasurer = ne.useMemo(() => createTextMeasurer(), []),
    assistantAbortControllerRef = ne.useRef(null),
    assistantRequestIdRef = ne.useRef(0),
    latestEditStateRef = ne.useRef({
      book: activeWorkbook,
      readOnly: isEditingLocked,
      sheetId: activeSheet.id,
      range: activeSelectionRange,
      selectionKey: selectionAreasKey,
    });
  latestEditStateRef.current = {
    book: activeWorkbook,
    readOnly: isEditingLocked,
    sheetId: activeSheet.id,
    range: activeSelectionRange,
    selectionKey: selectionAreasKey,
  };
  const lastExternalProposalIdRef = ne.useRef(undefined);
  function handleCalculationCompleted() {
    (setRenderRefreshCount((fe) => fe + 1),
      setAssistantProposal(null),
      setAnswerCitations(null),
      setAuditAnalysisResult(null));
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
      () => parseRangeAddress(activeSelectionRange),
      [activeSelectionRange],
    ),
    selectedCellCount = ne.useMemo(
      () => countSelectionCells(selectionAreas),
      [selectionAreas],
    ),
    assistantSelection = {
      sheetId: activeSheet.id,
      range: activeSelectionRange,
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
    arrayDisplayResult.warning && setErrorNotice(arrayDisplayResult.warning);
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
    activeCellFormulaErrors =
      conditionalFormatEvaluator(activeCellAddress).errors,
    activeCellReferenceDiagnostics = ne.useMemo(
      () => (
        workbookCalculator(activeSheet.id, activeCellAddress),
        workbookCalculator
          .getReferenceDiagnostics()
          .filter(
            (fe) =>
              fe.sheetId === activeSheet.id && fe.address === activeCellAddress,
          )
      ),
      [workbookCalculator, activeSheet.id, activeCellAddress],
    ),
    activeCellNamedFormulaDiagnostics = ne.useMemo(
      () => (
        workbookCalculator(activeSheet.id, activeCellAddress),
        workbookCalculator
          .getNamedFormulaDiagnostics()
          .filter(
            (fe) =>
              fe.sheetId === activeSheet.id && fe.address === activeCellAddress,
          )
      ),
      [workbookCalculator, activeSheet.id, activeCellAddress],
    ),
    getAllCellEntries = ne.useMemo(() => {
      let fe;
      return () =>
        fe ??
        (fe = Object.entries(arrayDisplayResult.cells).map(([Oe, Be]) => ({
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
    activeCell = activeSheet.cells[activeCellAddress],
    sortedRowIndexes = ne.useMemo(() => {
      if (!filterValueSearch && !activeColumnSort) return null;
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
      const Oe = filterValueSearch.toLowerCase(),
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
        activeColumnSort &&
          Be.sort((Ue, Xe) => {
            const gt = workbookCalculator(
                activeSheet.id,
                `${columnIndexToLetters(activeColumnSort.col)}${Ue + 1}`,
              ),
              Nt = workbookCalculator(
                activeSheet.id,
                `${columnIndexToLetters(activeColumnSort.col)}${Xe + 1}`,
              );
            return (
              (isFiniteNumber(gt) && isFiniteNumber(Nt)
                ? gt - Nt
                : String(gt ?? "").localeCompare(String(Nt ?? ""))) *
              activeColumnSort.direction
            );
          }),
        Be
      );
    }, [
      getAllCellEntries,
      workbookCalculator,
      activeSheet.id,
      filterValueSearch,
      activeColumnSort,
      pinnedRowCount,
      rowVisibilityIndex,
    ]),
    filteredRowPositionMap = ne.useMemo(
      () =>
        sortedRowIndexes
          ? new Map(sortedRowIndexes.map((fe, Oe) => [fe, Oe]))
          : null,
      [sortedRowIndexes],
    ),
    [bn, setAutoFitRowHeights] = ne.useState({});
  ne.useEffect(
    () => setAutoFitRowHeights({}),
    [
      activeSheet.id,
      activeSheet.cells,
      activeSheet.columnWidths,
      activeSheet.rowHeights,
      activeSheet.defaultRowHeight,
      activeSheet.defaultColumnWidth,
      isShowingFormulas,
    ],
  );
  const effectiveRowHeights = ne.useMemo(
      () => ({
        ...bn,
        ...activeSheet.rowHeights,
        ...(axisResizeState?.axis === "row"
          ? { [axisResizeState.index]: axisResizeState.size }
          : {}),
      }),
      [bn, activeSheet.rowHeights, axisResizeState],
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
            ...(axisResizeState?.axis === "column"
              ? { [axisResizeState.index]: axisResizeState.size }
              : {}),
          },
          hiddenColumnIndexes,
        ),
      [
        activeSheet.colCount,
        defaultColumnWidth,
        activeSheet.columnWidths,
        hiddenColumnIndexes,
        axisResizeState,
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
  const {
    selectSingleRange,
    applySelectionAreas,
    requireSingleSelectionRange,
    scrollActivePaneVertically,
    scrollActivePaneHorizontally,
    clearCutSession,
    computeGridViewport,
    applyPaneScrollOffsets,
    getSelectedCellAddresses,
    buildWorkbookViewState,
    applyWorkbookView,
    notifyHostCallback,
    runGuardedAction,
    focusGridSurface,
    requestCellScroll,
    stepVisibleRowIndex,
    stepVisibleColumnIndex,
    jumpToDataEdge,
  } = createViewAndSelectionActions({
    activeReliabilityMonitor,
    setActiveSheetId,
    activeSheet,
    setActiveCellAddress,
    setSelectionAnchorAddress,
    activeCellAddressRef,
    activeSelectionRange,
    setActiveSelectionRange,
    setCompletedSelectionAreas,
    selectionAreas,
    setFilterValueSearch,
    setActiveColumnSort,
    sheetSplitPanes,
    splitPaneSizes,
    setPendingScrollAddress,
    setErrorNotice,
    isHeadingsVisible,
    internalClipboardRef,
    setClipboardSnapshotState,
    setMoveCellsSource,
    rowVisibilityIndex,
    columnVisibilityIndex,
    zoomScale,
    rowHeaderWidth,
    gridContainerRef,
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

  const {
    runCalculationNow,
    reportSaveStatus,
    reportSaveError,
    enqueueSaveOperation,
    acknowledgeSavedWorkbook,
    recoverLocalWorkbook,
    persistWorkbookSnapshot,
    exportLocalChangesJson,
    reloadSavedWorkbook,
    commitWorkbookChange,
  } = createPersistenceActions({
    onChange,
    onBeforeChange,
    readOnly,
    persistenceKey,
    workbookStore,
    activeReliabilityMonitor,
    saveWarningDelayMs,
    setWorkbookSnapshot,
    workbookPersistence,
    persistedWorkbookRecordRef,
    setPersistedWorkbookRecord,
    isWorkbookLoadedRef,
    isEditableRef,
    persistenceSessionRef,
    saveAbortControllerRef,
    setSaveAbortControllerRef,
    setPersistenceError,
    persistenceMessageRef,
    setIsSaveConflictDialogOpen,
    setPendingSaveCount,
    isEditingLocked,
    setCellEditState,
    activeCellEditRef,
    setUndoStack,
    setRedoStack,
    setStatusNotice,
    setErrorNotice,
    setAuditAnalysisResult,
    setAssistantProposal,
    setAnswerCitations,
    calculationRun,
    latestEditStateRef,
    buildWorkbookViewState,
    applyWorkbookView,
    notifyHostCallback,
  });

  const {
    clearSelectedCellContents,
    applyCellPatches,
    resetFormulaEditSession,
    syncFormulaEditorSelection,
    restoreFormulaEditorFocus,
    isFormulaEditingLocked,
    beginCellEdit,
    cancelCellEdit,
    commitCellEdit,
    applyCellFormatPatch,
  } = createCellEditActions({
    actor,
    readOnly,
    workbookMessages,
    getEditableCellText,
    setActiveSheetId,
    activeSheet,
    saveAbortControllerRef,
    formulaInputRef,
    activeCellAddress,
    setActiveCellAddress,
    setSelectionAnchorAddress,
    selectSingleRange,
    setCellEditText,
    setCellEditState,
    activeCellEditRef,
    formulaPointerStateRef,
    setActiveFormulaReference,
    setIsFormulaPointMode,
    pendingCrossSheetTargetRef,
    setFilterValueSearch,
    setActiveColumnSort,
    setPendingScrollAddress,
    setErrorNotice,
    setValidationReviewState,
    workbookCalculator,
    latestEditStateRef,
    getSelectedCellAddresses,
    commitWorkbookChange,
    runGuardedAction,
    focusGridSurface,
  });

  const {
    commitWorkbookEdit,
    replaceWorkbookSnapshot,
    commitWorkbookWithNewSheet,
    applyZoomPercent,
    selectCellAddress,
    switchActiveWorksheet,
    navigateToSheetCell,
    selectSheetRanges,
    commitNameBoxReference,
    undoOrRedo,
    replaceActiveWorkbook,
    toggleFreezeTopRow,
    moveSelectionAfterCommit,
    selectAllCells,
    selectRowOrColumn,
  } = createNavigationActions({
    actor,
    readOnly,
    activeWorkbook,
    activeSheetId,
    setActiveSheetId,
    activeSheet,
    isEditingLocked,
    setActiveDialogMode,
    activeCellAddress,
    setActiveCellAddress,
    selectionAnchorAddress,
    setSelectionAnchorAddress,
    activeCellAddressRef,
    selectSingleRange,
    applySelectionAreas,
    selectionAreas,
    goToReferenceText,
    setCellEditState,
    activeCellEditRef,
    formulaPointerStateRef,
    setActiveFormulaReference,
    pendingCrossSheetTargetRef,
    setSidePanelKind,
    filterValueSearch,
    setFilterValueSearch,
    setFilterStateByScope,
    activeColumnSort,
    setActiveColumnSort,
    sheetSplitPanes,
    splitPaneSizes,
    scrollActivePaneVertically,
    scrollActivePaneHorizontally,
    setPendingScrollAddress,
    undoStack,
    setUndoStack,
    redoStack,
    setRedoStack,
    setStatusNotice,
    setErrorNotice,
    setAuditAnalysisResult,
    setAuditSourceRange,
    isHeadingsVisible,
    setZoomPercent,
    setIsFilterBarVisible,
    setIsFindBarVisible,
    hasFrozenRows,
    rowVisibilityIndex,
    columnVisibilityIndex,
    rowHeaderWidth,
    gridContainerRef,
    latestEditStateRef,
    scrollableRowLayout,
    scrollableColumnLayout,
    frozenRowsHeight,
    frozenColumnsWidth,
    computeGridViewport,
    buildWorkbookViewState,
    applyWorkbookView,
    commitWorkbookChange,
    runGuardedAction,
    restoreFormulaEditorFocus,
    focusGridSurface,
    commitCellEdit,
    requestCellScroll,
    stepVisibleRowIndex,
    stepVisibleColumnIndex,
  });

  const {
    insertFormulaReference,
    handleFormulaPointerDown,
    handleFormulaKeyDown,
  } = createFormulaEntryActions({
    activeWorkbook,
    activeSheet,
    setFormulaCaretPosition,
    activeCellAddress,
    activeCellEditRef,
    formulaPointerStateRef,
    setActiveFormulaReference,
    setIsFormulaPointMode,
    isOperationBusy,
    dragAnchorRef,
    syncFormulaEditorSelection,
    restoreFormulaEditorFocus,
    isFormulaEditingLocked,
    beginCellEdit,
    requestCellScroll,
    switchActiveWorksheet,
    stepVisibleRowIndex,
    stepVisibleColumnIndex,
    jumpToDataEdge,
  });

  const {
    importWorkbookFromFile,
    exportWorkbookAs,
    runCurrentAuditAnalysis,
    askAssistantQuestion,
  } = createFileAndAnalysisActions({
    adapter,
    actor,
    activeReliabilityMonitor,
    activeWorkbook,
    activeSheet,
    isEditingLocked,
    setActiveDialogMode,
    setAutomationInspection,
    selectionAreas,
    auditPopulationScope,
    activeFilterState,
    setStatusNotice,
    setErrorNotice,
    setIsOperationBusy,
    operationScope,
    latestOperationIdRef,
    activeOperationControllerRef,
    setDelimitedImportExportSource,
    setPendingImportFile,
    setXlsxConversionReview,
    selectedAuditKind,
    setAuditAnalysisResult,
    materialityThreshold,
    auditSampleSize,
    samplingSeed,
    setAgingReportPayload,
    agingAsOfDate,
    auditSourceRange,
    setAssistantProposal,
    setAnswerCitations,
    setIsAssistantBusy,
    filterResult,
    importFileInputRef,
    assistantAbortControllerRef,
    assistantRequestIdRef,
    latestEditStateRef,
    assistantSelection,
    commitWorkbookChange,
    runGuardedAction,
    switchActiveWorksheet,
  });

  const {
    selectionToTsv,
    copyOrCutSelection,
    pasteClipboardText,
    pasteFromClipboard,
    findNextTextMatch,
    replaceTextInSelection,
  } = createClipboardAndFindActions({
    activeWorkbook,
    activeSheet,
    isEditingLocked,
    activeCellAddress,
    activeSelectionRange,
    selectSingleRange,
    selectionAreasKey,
    requireSingleSelectionRange,
    setStatusNotice,
    setErrorNotice,
    findQuery,
    replaceWithText,
    internalClipboardRef,
    clipboardSnapshotState,
    setClipboardSnapshotState,
    setMoveCellsSource,
    workbookCalculator,
    rowVisibilityIndex,
    columnVisibilityIndex,
    gridContainerRef,
    latestEditStateRef,
    selectionRangeBounds,
    getSelectedCellAddresses,
    runGuardedAction,
    applyCellPatches,
    selectCellAddress,
  });

  const {
    rowViewport,
    columnViewport,
    rowWindow,
    columnWindow,
    shownRows,
    shownColumns,
  } = computeGridViewport(activePaneLayout);
  const {
    applyAxisSizes,
    autoFitSelectedAxis,
    beginAxisResize,
    beginFillDrag,
    extendFillDragTo,
  } = createRowColumnAndFillActions({
    actor,
    activeSheet,
    isEditingLocked,
    activeCellAddress,
    activeSelectionRange,
    requireSingleSelectionRange,
    isShowingFormulas,
    dragAnchorRef,
    setFillDragState,
    fillDragStateRef,
    setAxisResizeState,
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
    getSelectedCellAddresses,
    commitWorkbookChange,
    runGuardedAction,
    focusGridSurface,
    commitCellEdit,
    requestCellScroll,
  });

  const {
    runEditorCommand,
    handleGridKeyDown,
    commitDialogInput,
    activateGridPane,
  } = createGridInteractionActions({
    getEditableCellText,
    activeWorkbook,
    activeSheet,
    isEditingLocked,
    setActiveDialogMode,
    setAutomationInspection,
    setContextMenuState,
    appRootRef,
    setTabBarSheetId,
    activeCellAddress,
    setActiveCellAddress,
    setSelectionAnchorAddress,
    activeCellAddressRef,
    activeSelectionRange,
    completedSelectionAreas,
    selectSingleRange,
    selectionAreas,
    requireSingleSelectionRange,
    activeCellEditRef,
    formulaPointerStateRef,
    setIsFormulaPointMode,
    filterValueSearch,
    setActiveFilterColumn,
    activeColumnSort,
    splitPaneSizes,
    setStatusNotice,
    setErrorNotice,
    isOperationBusy,
    setXlsxConversionReview,
    promptDialogState,
    setPromptDialogState,
    setIsFindBarVisible,
    clipboardSnapshotState,
    setMoveCellsSource,
    clearCutSession,
    rowVisibilityIndex,
    columnVisibilityIndex,
    defaultRowHeight,
    zoomScale,
    gridContainerRef,
    findInputRef,
    runCalculationNow,
    mergeRanges,
    activeCell,
    rowViewport,
    getSelectedCellAddresses,
    commitWorkbookEdit,
    replaceWorkbookSnapshot,
    runGuardedAction,
    applyCellPatches,
    beginCellEdit,
    commitCellEdit,
    requestCellScroll,
    selectCellAddress,
    switchActiveWorksheet,
    applyCellFormatPatch,
    undoOrRedo,
    stepVisibleRowIndex,
    stepVisibleColumnIndex,
    jumpToDataEdge,
    selectAllCells,
    selectRowOrColumn,
  });

  (ne.useEffect(() => {
    const fe = {};
    for (const Be of Pv(sheetSplitPanes)) {
      const Ue = computeGridViewport(splitPaneSizes.current.current[Be]),
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
      (gridContainerRef.current = splitPaneSizes.refs.current[Oe] ?? null));
  }, [splitPaneSizesKey]),
    ne.useEffect(() => {
      if (!gridContainerRef.current || typeof ResizeObserver > "u") return;
      const Oe = new ResizeObserver((Be) => {
        const Ue = {};
        for (const gt of Be) {
          const Nt = Number(gt.target.dataset.row),
            wt = gt.target.offsetHeight;
          Number.isInteger(Nt) &&
            activeSheet.rowHeights?.[Nt] === undefined &&
            axisResizeState?.axis !== "row" &&
            wt > defaultRowHeight &&
            (Ue[Nt] = wt);
        }
        Object.keys(Ue).length &&
          setAutoFitRowHeights((gt) =>
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
      if (!pendingScrollAddress || !gridContainerRef.current) return;
      const fe = _t(pendingScrollAddress),
        Oe = gridContainerRef.current;
      if (
        rowVisibilityIndex.isHidden(fe.row) ||
        columnVisibilityIndex.isHidden(fe.col)
      ) {
        setPendingScrollAddress(null);
        return;
      }
      const Be = scrollableRowPositionOf(fe.row);
      if (Be === undefined && fe.row >= pinnedRowCount) {
        (setFilterValueSearch(""), setActiveColumnSort(null));
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
        scrollActivePaneVertically(Ue),
        scrollActivePaneHorizontally(Xe),
        setPendingScrollAddress(null));
    }, [
      pendingScrollAddress,
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
    activeCellEditRef.current ||
      setCellEditText(
        getEditableCellText(activeSheet.cells[activeCellAddress]),
      );
  }, [activeSheet, activeCellAddress, workbookMessages.entryLocale]),
    ne.useEffect(() => {
      (setGoToReferenceText(selectionAreasKey),
        notifyHostCallback("selection", () =>
          onSelectionChange?.({
            sheetId: activeSheet.id,
            range: activeSelectionRange,
            ...(selectionAreas.length > 1 ? { ranges: selectionAreas } : {}),
          }),
        ));
    }, [
      selectionAreasKey,
      activeSelectionRange,
      activeSheet.id,
      onSelectionChange,
    ]),
    ne.useEffect(() => {
      isWorkbookLoadedRef.current = false;
      const fe = { active: true, controllers: new Set() };
      persistenceSessionRef.current = fe;
      try {
        const Oe = workbookPersistence.load();
        if (
          ((persistedWorkbookRecordRef.current = Oe),
          setPersistedWorkbookRecord(Oe),
          Oe)
        ) {
          (setWorkbookSnapshot(buildCanonicalWorkbookSnapshot(Oe.workbook)),
            (latestEditStateRef.current.book = Oe.workbook),
            applyWorkbookView(Oe.view),
            setStatusNotice(
              Oe.warning ?? "Restored your locally saved workbook.",
            ),
            reportSaveStatus("saved", Oe.workbook));
          try {
            notifyHostCallback("loaded", () => onWorkbookLoaded?.(Oe.workbook));
          } catch {
            setErrorNotice(
              "The host could not acknowledge the loaded workbook.",
            );
          }
        }
        isWorkbookLoadedRef.current = true;
      } catch (Oe) {
        (reportSaveError(Oe),
          (persistenceMessageRef.current =
            "The saved workbook could not be restored. Import a backup to recover it. Existing saved data will be preserved."),
          setErrorNotice(persistenceMessageRef.current));
      }
      return () => {
        var Be;
        fe.active = false;
        for (const Ue of fe.controllers) Ue.abort();
        const Oe = persistenceSessionRef.current;
        Oe.active = false;
        for (const Ue of Oe.controllers) Ue.abort();
        (Be = saveAbortControllerRef.current) == null || Be.abort();
      };
    }, [workbookPersistence]),
    ne.useEffect(() => {
      var fe;
      if (readOnly) {
        for (const Oe of persistenceSessionRef.current.controllers) Oe.abort();
        (fe = saveAbortControllerRef.current) == null || fe.abort();
      }
    }, [readOnly]),
    ne.useEffect(() => {
      if (!persistenceKey && !("coordination" in workbookPersistence)) return;
      const fe = (Oe) => {
        const Be = persistedWorkbookRecordRef.current?.guard,
          Ue = persistenceKey ?? Be?.namespace;
        if (!(!Ue || (Oe.key !== null && Oe.key !== Ue)))
          try {
            const gt = workbookPersistence.load();
            if (
              JSON.stringify(gt?.guard ?? null) === JSON.stringify(Be ?? null)
            )
              return;
            reportSaveError(
              new WorkbookRevisionConflictError(
                "conflict",
                "Another workbook window changed the saved version. Your local work has been retained.",
                gt?.guard,
              ),
            );
          } catch (gt) {
            reportSaveError(gt);
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
        (activeCellEditRef.current ||
          persistenceSessionRef.current.controllers.size ||
          JSON.stringify(persistedWorkbookRecordRef.current?.workbook) !==
            JSON.stringify(latestEditStateRef.current.book)) &&
          (Oe.preventDefault(), (Oe.returnValue = ""));
      };
      return (
        window.addEventListener("beforeunload", fe),
        () => window.removeEventListener("beforeunload", fe)
      );
    }, [persistenceKey, workbookStorePersistent]),
    ne.useEffect(() => {
      if (!isWorkbookLoadedRef.current || readOnly) return;
      const fe = setTimeout(() => {
        try {
          const Oe = latestEditStateRef.current;
          if (
            Oe.readOnly ||
            persistenceError ||
            persistenceSessionRef.current.controllers.size
          )
            return;
          const Be = buildWorkbookViewState(Oe.book, Oe.sheetId, Oe.range),
            Ue = persistedWorkbookRecordRef.current;
          if (
            Ue &&
            JSON.stringify(Ue.view) === JSON.stringify(Be) &&
            (Ue.workbook === Oe.book ||
              JSON.stringify(Ue.workbook) === JSON.stringify(Oe.book))
          )
            return;
          const Xe = persistWorkbookSnapshot(Oe.book, Be);
          isPromiseLike(Xe) && Xe.catch(() => {});
        } catch (Oe) {
          setErrorNotice(
            `${Oe instanceof Error ? Oe.message : "The current selection could not be saved."} Changes remain in this editor; export JSON to keep a copy.`,
          );
        }
      }, 400);
      return () => clearTimeout(fe);
    }, [
      activeSheet.id,
      activeSelectionRange,
      workbookPersistence,
      readOnly,
      persistenceError,
    ]),
    ne.useEffect(() => {
      (setAuditAnalysisResult(null), setAgingReportPayload(null));
    }, [
      auditSourceRange,
      materialityThreshold,
      auditSampleSize,
      samplingSeed,
      agingAsOfDate,
    ]),
    ne.useEffect(() => {
      var fe;
      return (
        assistantRequestIdRef.current++,
        (fe = assistantAbortControllerRef.current) == null || fe.abort(),
        setIsAssistantBusy(false),
        setAssistantProposal(null),
        setAnswerCitations(null),
        () => {
          return assistantAbortControllerRef.current?.abort();
        }
      );
    }, [activeWorkbook, activeSheet.id, selectionAreasKey, adapter]),
    ne.useEffect(() => {
      if (
        !externalProposal ||
        externalProposal.requestId === lastExternalProposalIdRef.current
      )
        return;
      lastExternalProposalIdRef.current = externalProposal.requestId;
      let fe;
      try {
        if (
          JSON.stringify(externalProposal.context) !==
          JSON.stringify(
            buildAssistantSelectionContext(activeWorkbook, {
              sheetId: activeSheet.id,
              range: activeSelectionRange,
              ranges: selectionAreas,
            }),
          )
        )
          throw new Error(
            "The workbook or selected context changed. Request a new external proposal.",
          );
        const Oe = validateProposedCalculation(
          activeWorkbook,
          { sheetId: activeSheet.id, range: activeSelectionRange },
          externalProposal.proposal,
        );
        (setAnswerCitations(
          buildAnswerCitations(
            activeWorkbook,
            { sheetId: activeSheet.id, range: activeSelectionRange },
            Oe,
            externalProposal.context,
          ),
        ),
          setAssistantProposal(Oe),
          setSidePanelKind("assistant"),
          setErrorNotice(""),
          (fe = { accepted: true }));
      } catch (Oe) {
        const Be =
          Oe instanceof Error ? Oe.message : "Invalid external proposal.";
        (setErrorNotice(Be), (fe = { accepted: false, error: Be }));
      }
      try {
        notifyHostCallback("proposal", () =>
          onExternalProposalStaged?.(externalProposal.requestId, fe),
        );
      } catch {
        setErrorNotice("The host could not acknowledge the staged proposal.");
      }
    }, [
      externalProposal,
      activeWorkbook,
      activeSheet.id,
      selectionAreasKey,
      onExternalProposalStaged,
    ]));

  formulaPointerMoveRef.current = (fe) => {
    formulaPointerStateRef.current.head !== fe &&
      insertFormulaReference(fe, true);
  };

  ne.useEffect(() => {
    if (!axisResizeState) return;
    const fe = (Xe) =>
        setAxisResizeState(
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
        (axisResizeState.size !== axisResizeState.initial &&
          runGuardedAction(() =>
            applyAxisSizes(
              axisResizeState.axis,
              { [axisResizeState.index]: axisResizeState.size },
              `Resize ${axisResizeState.axis} ${axisResizeState.axis === "column" ? columnIndexToLetters(axisResizeState.index) : axisResizeState.index + 1}`,
            ),
          ),
          setAxisResizeState(null));
      },
      Be = () => setAxisResizeState(null),
      Ue = (Xe) => {
        Xe.key === "Escape" &&
          (!Xe.defaultPrevented &&
            Xe.target instanceof Element &&
            Xe.target.closest(".lg-app") === appRootRef.current &&
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
    axisResizeState,
    activeWorkbook,
    activeSheet.id,
    zoomScale,
    isEditingLocked,
  ]);

  ne.useEffect(() => {
    if (!fillDragState) return;
    const fe = () => {
        ((fillDragStateRef.current = null), setFillDragState(null));
      },
      Oe = () => {
        const Ue = fillDragStateRef.current;
        if ((fe(), !(!Ue || Ue.targetRange === Ue.sourceRange))) {
          if (
            latestEditStateRef.current.readOnly ||
            latestEditStateRef.current.book !== Ue.book ||
            latestEditStateRef.current.sheetId !== Ue.sheetId
          ) {
            setErrorNotice(
              "The workbook changed while filling. Select the cells and try again.",
            );
            return;
          }
          setFillPreviewState(Ue);
        }
      },
      Be = (Ue) => {
        Ue.key === "Escape" &&
          (!Ue.defaultPrevented &&
            Ue.target instanceof Element &&
            Ue.target.closest(".lg-app") === appRootRef.current &&
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
  }, [!!fillDragState]);

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
          className: `${_t(activeCellAddress).col === fe ? "lg-selected-header" : ""} ${Oe ? "lg-frozen-column" : ""}`,
          style: Oe
            ? {
                position: "sticky",
                left: frozenColumnLeftOffsets[fe],
                zIndex: 7,
              }
            : undefined,
          onClick: (Be) => {
            insertFormulaReference(`${columnIndexToLetters(fe)}1`, false, {
              range: `${columnIndexToLetters(fe)}:${columnIndexToLetters(fe)}`,
              bounds: {
                firstRow: 0,
                lastRow: activeSheet.rowCount - 1,
                firstCol: fe,
                lastCol: fe,
              },
            }) || selectRowOrColumn("column", fe, Be.shiftKey);
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
              onPointerDown: (Be) => beginAxisResize("column", fe, Be),
              onDoubleClick: (Be) => {
                (Be.stopPropagation(), autoFitSelectedAxis("column", fe));
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
    renderGridRow = (fe, Oe, Be, Ue) => {
      var gt;
      return i.jsx(
        WorksheetGridRow,
        {
          idPrefix: `${editorIdPrefix}-${Ue}`,
          row: fe,
          pinned: Oe,
          pinnedTop:
            (isHeadingsVisible ? 22 : 0) +
            rowAxisLayout.offset(rowVisibilityIndex.viewIndex(fe) ?? 0),
          frozenColumns: frozenColumnIndexes,
          frozenColumnOffsets: frozenColumnLeftOffsets,
          headings: isHeadingsVisible,
          columns: Be.shownColumns,
          leftGap: Be.leftGap,
          rightGap: Be.rightGap,
          minimumHeight: effectiveRowHeights[fe] ?? defaultRowHeight,
          fixedHeight:
            activeSheet.rowHeights?.[fe] !== undefined ||
            (axisResizeState?.axis === "row" && axisResizeState.index === fe),
          onRowSelect: (Nt) => {
            insertFormulaReference(`A${fe + 1}`, false, {
              range: `${fe + 1}:${fe + 1}`,
              bounds: {
                firstRow: fe,
                lastRow: fe,
                firstCol: 0,
                lastCol: activeSheet.colCount - 1,
              },
            }) || selectRowOrColumn("row", fe, Nt);
          },
          onResize: (Nt) => beginAxisResize("row", fe, Nt),
          onAutoFit: () => autoFitSelectedAxis("row", fe),
          onCancelEdit: cancelCellEdit,
          onFormulaInput: syncFormulaEditorSelection,
          onFormulaKey: handleFormulaKeyDown,
          formulaHelpId: `${editorIdPrefix}-formula-help`,
          onFormulaPick: insertFormulaReference,
          formulaReference:
            activeFormulaReference?.sheetId === activeSheet.id
              ? activeFormulaReference.bounds
              : null,
          formulaDraftActive: !!(
            (gt = activeCellEditRef.current) != null && gt.value.startsWith("=")
          ),
          onDragStart: (Nt, wt) => {
            ((dragAnchorRef.current = Nt),
              (dragBaseSelectionsRef.current = wt ? [...selectionAreas] : []));
          },
          onDragEnter: (Nt, wt) => {
            if (formulaPointerStateRef.current.dragging && wt === 1) {
              insertFormulaReference(Nt, true);
              return;
            }
            fillDragStateRef.current && wt === 1
              ? extendFillDragTo(Nt)
              : dragAnchorRef.current &&
                wt === 1 &&
                ((activeCellAddressRef.current = Nt),
                runGuardedAction(() =>
                  applySelectionAreas([
                    ...dragBaseSelectionsRef.current,
                    wy(dragAnchorRef.current, Nt),
                  ]),
                ),
                setActiveCellAddress(dragAnchorRef.current));
          },
          onNavigate: moveSelectionAfterCommit,
          onPivotDrill: (Nt) => {
            var jt;
            const wt = _t(Nt);
            return (jt = activeSheet.pivots) != null &&
              jt.some((rr) => {
                const Mr = parseRangeAddress(rr.outputRange);
                return (
                  wt.row > Mr.firstRow &&
                  wt.row <= Mr.lastRow &&
                  wt.col >= Mr.firstCol + Math.max(1, rr.rowFields.length) &&
                  wt.col <= Mr.lastCol
                );
              })
              ? (selectCellAddress(Nt), runEditorCommand("pivotExplore"), true)
              : false;
          },
          commentAddresses: commentedCellAddresses,
          onReview: (Nt) => {
            (selectCellAddress(Nt), runEditorCommand("comments"));
          },
          onLink: (Nt, wt) =>
            runGuardedAction(() => navigateToSheetCell(Nt, wt)),
          onFillStart: beginFillDrag,
          onFillOptions: () => runEditorCommand("fillSeries"),
          fillDragging: !!fillDragState || completedSelectionAreas.length > 0,
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
            (runEditorCommand("populationFilter"), setActiveFilterColumn(Nt));
          },
          sortedView: !!(filterValueSearch || activeColumnSort),
          ruleStyles: getConditionalRuleStyle,
          conditionalFormatter: conditionalFormatEvaluator,
          sheetId: activeSheet.id,
          cells: activeSheet.cells,
          active: activeCellAddress,
          selectedBounds: fillDragState
            ? parseRangeAddress(fillDragState.targetRange)
            : selectionRangeBounds,
          selectionBounds: fillDragState
            ? [parseRangeAddress(fillDragState.targetRange)]
            : selectionAreaRanges,
          calc: workbookCalculator,
          showFormulas: isShowingFormulas,
          editing:
            splitPaneSizes.activeId === Ue &&
            (!activeCellEditRef.current ||
              activeCellEditRef.current.sheetId === activeSheet.id)
              ? cellEditState
              : null,
          setEditing: (Nt) =>
            Nt === null ? cancelCellEdit() : beginCellEdit(Nt, true),
          select: (Nt, wt, jt) => {
            var rr;
            (selectCellAddress(Nt, wt, jt),
              (rr = gridContainerRef.current) == null ||
                rr.focus({ preventScroll: true }));
          },
          saveCell: commitCellEdit,
          readOnly: isEditingLocked,
        },
        fe,
      );
    },
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
    text,
    language,
    editorIdPrefix,
    activeWorkbook,
    activeSheet,
    isEditingLocked,
    setContextMenuState,
    activeCellAddress,
    selectionAreas,
    selectionAreaRanges,
    filterValueSearch,
    activeColumnSort,
    sheetSplitPanes,
    splitPaneSizes,
    setStatusNotice,
    isOperationBusy,
    isGridlinesVisible,
    isHeadingsVisible,
    dragAnchorRef,
    workbookCalculator,
    rowVisibilityIndex,
    columnVisibilityIndex,
    pinnedRowCount,
    frozenColumnCount,
    zoomScale,
    rowHeaderWidth,
    gridContainerRef,
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
    computeGridViewport,
    applyPaneScrollOffsets,
    commitWorkbookEdit,
    runGuardedAction,
    handleFormulaPointerDown,
    commitCellEdit,
    selectCellAddress,
    selectionToTsv,
    pasteClipboardText,
    selectAllCells,
    selectRowOrColumn,
    handleGridKeyDown,
    renderColumnHeader,
    renderGridRow,
    renderVirtualSpacer,
    activateGridPane,
  });

  const closeContextMenu = ne.useCallback((fe) => {
      var Oe;
      (setContextMenuState(null),
        fe &&
          ((Oe = gridContainerRef.current) == null ||
            Oe.focus({ preventScroll: true })));
    }, []),
    editBlockReason = isOperationBusy
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
    contextMenuState && getSelectedCellAddresses();
  } catch (fe) {
    selectionReadError = fe.message;
  }
  const contextMenuActions = [
    {
      name: "Cut",
      run: () => void copyOrCutSelection(true),
      reason: editBlockReason ?? multiRangeBlockReason ?? selectionReadError,
    },
    {
      name: "Copy",
      run: () => void copyOrCutSelection(false),
      reason: multiRangeBlockReason ?? selectionReadError,
    },
    {
      name: "Paste",
      run: () => void pasteFromClipboard(),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "Paste Special",
      run: () => runEditorCommand("pasteSpecial"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "Format Cells",
      run: () => runEditorCommand("numberFormat"),
      reason: editBlockReason,
    },
    {
      name: "Cell Styles",
      run: () => runEditorCommand("cellStyle"),
      reason: editBlockReason,
    },
    {
      name: "Insert / Delete / Move Rows",
      run: () => runEditorCommand("insertRow"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "Insert / Delete / Move Columns",
      run: () => runEditorCommand("insertColumn"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "AutoFit Column Width",
      run: () => autoFitSelectedAxis("column"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "AutoFit Row Height",
      run: () => autoFitSelectedAxis("row"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "Hide / Unhide / Freeze",
      run: () => runEditorCommand("viewSettings"),
      reason: editBlockReason ?? multiRangeBlockReason,
    },
    {
      name: "Clear Contents",
      run: () =>
        runGuardedAction(() =>
          applyCellPatches(
            getSelectedCellAddresses().map((fe) => ({
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
        setPromptDialogState({ kind: "note", value: activeCell?.note ?? "" }),
      reason: editBlockReason,
    },
    { name: "Comments", run: () => runEditorCommand("comments") },
    { name: "Link", run: () => runEditorCommand("hyperlink") },
  ];
  return i.jsxs("div", {
    className: "lg-app",
    ref: appRootRef,
    lang: workbookMessages.locale,
    role: "region",
    "aria-label": activeWorkbook.name,
    onKeyDown: (fe) => {
      var wt;
      if (
        fe.key !== "F6" ||
        fe.defaultPrevented ||
        fe.ctrlKey ||
        fe.metaKey ||
        fe.altKey ||
        fe.nativeEvent.isComposing ||
        activeDialogMode ||
        promptDialogState ||
        (activeCellEditRef.current &&
          (activeCellEditRef.current.value !==
            activeCellEditRef.current.originalText ||
            !commitCellEdit()))
      )
        return;
      fe.preventDefault();
      const Oe = [
          appRootRef.current?.querySelector(
            '[role="tab"][aria-selected="true"]',
          ),
          formulaInputRef.current,
          gridContainerRef.current,
          appRootRef.current?.querySelector(
            '[data-sheet-id][aria-current="true"]',
          ),
        ].filter((jt) => !!jt),
        Be = document.activeElement,
        Ue =
          Be != null && Be.closest(".xl-ribbon")
            ? 0
            : Be != null && Be.closest(".lg-sheet-tabs")
              ? 3
              : Oe.findIndex((jt) => jt === Be || jt.contains(Be)),
        Xe = fe.shiftKey ? -1 : 1;
      (wt =
        Oe[
          Ue < 0
            ? fe.shiftKey
              ? Oe.length - 1
              : 0
            : (Ue + Xe + Oe.length) % Oe.length
        ]) == null || wt.focus({ preventScroll: true });
    },
    children: [
      i.jsxs("div", {
        className: "lg-skip-links",
        children: [
          i.jsx("button", {
            onClick: () => {
              return gridContainerRef.current?.focus();
            },
            children: i.jsx(LocalizedText, { children: "Skip to worksheet" }),
          }),
          i.jsx("button", {
            onClick: () => {
              return formulaInputRef.current?.focus();
            },
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
            address: activeCellAddress,
            row: _t(activeCellAddress).row + 1,
            column: _t(activeCellAddress).col + 1,
            value: (workbookCalculator(activeSheet.id, activeCellAddress) ==
            null
              ? workbookMessages.text("Blank")
              : formatLocalizedNumberDisplay(
                  workbookCalculator(activeSheet.id, activeCellAddress),
                  activeSheet.cells[activeCellAddress],
                  activeWorkbook.datePolicy?.dateSystem,
                  workbookMessages.entryLocale,
                )
            ).slice(0, 250),
            formula:
              typeof activeSheet.cells[activeCellAddress]?.raw == "string" &&
              String(activeSheet.cells[activeCellAddress].raw).startsWith(
                "=",
              ) &&
              !activeSheet.cells[activeCellAddress].literal
                ? workbookMessages.message(" Formula: {formula}.", {
                    formula: String(
                      activeSheet.cells[activeCellAddress].raw,
                    ).slice(0, 250),
                  })
                : "",
            access: isEditingLocked ? workbookMessages.text(" Read only.") : "",
          },
        ),
      }),
      contextMenuState &&
        i.jsx(
          GridContextMenu,
          {
            ...contextMenuState,
            label: `${contextMenuState.kind === "cell" ? "Cells" : contextMenuState.kind === "row" ? "Rows" : "Columns"} ${activeSelectionRange}`,
            actions: contextMenuActions,
            workbook: activeWorkbook,
            selection: `${activeSheet.id}:${selectionAreas.join(";")}`,
            onClose: closeContextMenu,
          },
          `${contextMenuState.x}:${contextMenuState.y}:${contextMenuState.kind}`,
        ),
      fillPreviewState &&
        i.jsx(FillPreviewDialog, {
          workbook: activeWorkbook,
          sheetId: fillPreviewState.sheetId,
          sourceRange: fillPreviewState.sourceRange,
          targetRange: fillPreviewState.targetRange,
          readOnly: isEditingLocked,
          onWorkbookChange: commitWorkbookEdit,
          onSelectionChange: selectSingleRange,
          onClose: () => {
            (setFillPreviewState(null), focusGridSurface());
          },
        }),
      moveCellsSource &&
        i.jsx(MoveCellsDialog, {
          workbook: activeWorkbook,
          source: moveCellsSource,
          sheetId: activeSheet.id,
          range: activeSelectionRange,
          readOnly,
          onWorkbookChange: replaceWorkbookSnapshot,
          onSelectionChange: selectSingleRange,
          onApplied: clearCutSession,
          onClose: () => {
            (setMoveCellsSource(null), focusGridSurface());
          },
        }),
      renderEditorDialogs({
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
        persistedWorkbookRecordRef,
        persistedWorkbookRecord,
        persistenceSessionRef,
        saveAbortControllerRef,
        isEditingLocked,
        activeDialogMode,
        setActiveDialogMode,
        automationInspection,
        setAutomationInspection,
        tabBarSheetId,
        formulaInputRef,
        setFormulaCaretPosition,
        setIsFormulaBarFocused,
        activeCellAddress,
        activeSelectionRange,
        selectSingleRange,
        selectionAreas,
        activeCellEditRef,
        setFilterValueSearch,
        setFilterStateByScope,
        activeFilterColumn,
        setAuditPopulationScope,
        filterStateKey,
        activeFilterState,
        setActiveColumnSort,
        scrollActivePaneVertically,
        setStatusNotice,
        setErrorNotice,
        operationScope,
        pendingImportFile,
        setPendingImportFile,
        auditSourceRange,
        internalClipboardRef,
        calculationRun,
        workbookCalculator,
        filterResult,
        latestEditStateRef,
        handleCalculationCompleted,
        runCalculationNow,
        applyWorkbookView,
        reportSaveStatus,
        enqueueSaveOperation,
        acknowledgeSavedWorkbook,
        persistWorkbookSnapshot,
        commitWorkbookChange,
        commitWorkbookEdit,
        replaceWorkbookSnapshot,
        commitWorkbookWithNewSheet,
        beginCellEdit,
        switchActiveWorksheet,
        navigateToSheetCell,
        selectSheetRanges,
      }),
      delimitedImportExportSource &&
        i.jsx(DelimitedImportExportDialog, {
          source: delimitedImportExportSource,
          blocked:
            delimitedImportExportSource.snapshot !== activeWorkbook
              ? "The workbook changed. Cancel and review again for the current revision."
              : "bytes" in delimitedImportExportSource && isEditingLocked
                ? "This workbook is read-only. Import is unavailable."
                : undefined,
          onClose: () => setDelimitedImportExportSource(null),
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
              delimitedImportExportSource.snapshot
            )
              throw new Error("The workbook changed. Review the file again.");
            if (latestEditStateRef.current.readOnly)
              throw new Error("This workbook is read-only.");
            (commitWorkbookChange(fe),
              switchActiveWorksheet(fe.sheets[0].id),
              setStatusNotice(
                `Imported ${Oe.records} records using the reviewed schema. Formula-like source text remains literal.`,
              ),
              setDelimitedImportExportSource(null));
          },
          onExport: (fe, Oe) => {
            if (
              latestEditStateRef.current.book !==
              delimitedImportExportSource.snapshot
            )
              throw new Error("The workbook changed. Review the export again.");
            if (Oe.population && !("bytes" in delimitedImportExportSource)) {
              const Be = exportWorksheetToCsv(
                latestEditStateRef.current.book,
                delimitedImportExportSource.sheetId,
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
              setStatusNotice(
                `Delimited data exported; ${Oe.safeTextChanges} formula-like text values neutralized under the reviewed safe-export policy.`,
              ),
              setDelimitedImportExportSource(null));
          },
        }),
      xlsxConversionReview &&
        i.jsx(XlsxConversionReview, {
          report: xlsxConversionReview.report,
          inspectOnly: xlsxConversionReview.inspectOnly,
          onOriginal: () =>
            runGuardedAction(() => {
              const fe = requireRetainedXlsxSource(
                xlsxConversionReview.imported ?? xlsxConversionReview.snapshot,
              );
              saveBlobToFile(
                new Blob([new Uint8Array(fe.bytes)], {
                  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                }),
                fe.fileName,
              );
            }),
          blocked:
            xlsxConversionReview.snapshot !== activeWorkbook
              ? "The workbook changed. Cancel and review the file again for the current revision."
              : xlsxConversionReview.imported && isEditingLocked
                ? "This workbook is read-only. Import is unavailable."
                : undefined,
          onClose: () => setXlsxConversionReview(null),
          onReport: () =>
            saveBlobToFile(
              new Blob([JSON.stringify(xlsxConversionReview.report, null, 2)], {
                type: "application/json",
              }),
              `${xlsxConversionReview.fileName}.compatibility.json`,
            ),
          onConfirm: () =>
            runGuardedAction(() => {
              if (
                latestEditStateRef.current.book !==
                xlsxConversionReview.snapshot
              )
                throw new Error("The workbook changed. Review the file again.");
              if (xlsxConversionReview.imported) {
                if (latestEditStateRef.current.readOnly)
                  throw new Error("This workbook is read-only.");
                (commitWorkbookChange(xlsxConversionReview.imported),
                  switchActiveWorksheet(
                    xlsxConversionReview.imported.sheets.find(
                      (fe) => !fe.hidden,
                    ).id,
                  ),
                  setStatusNotice(
                    "Derived XLSX workbook opened. The exact original is available from File → Original XLSX and compatibility. Save native JSON to retain it across sessions.",
                  ));
              } else
                xlsxConversionReview.blob &&
                  (saveBlobToFile(
                    xlsxConversionReview.blob,
                    xlsxConversionReview.fileName,
                  ),
                  setStatusNotice(
                    "Derived XLSX downloaded. Keep a native JSON backup for full application metadata.",
                  ));
              setXlsxConversionReview(null);
            }),
        }),
      i.jsxs("div", {
        className: "lg-main",
        children: [
          i.jsx("input", {
            type: "file",
            "aria-label": text("Import workbook file"),
            ref: importFileInputRef,
            accept:
              ".xlsx,.xltx,.xls,.xlsb,.xlsm,.xltm,.ods,.csv,.tsv,.json,.bas,.vba,.cls,.ts,.osts,.xlam",
            hidden: true,
            onChange: (fe) => {
              return void importWorkbookFromFile(fe.target.files?.[0]);
            },
            lang: language("Import workbook file"),
            "data-ui-label": "Import workbook file",
          }),
          renderEditorRibbon({
            readOnly,
            activeWorkbook,
            activeSheet,
            isEditingLocked,
            activeCellAddress,
            activeSelectionRange,
            requireSingleSelectionRange,
            setSidePanelKind,
            setFilterValueSearch,
            setActiveColumnSort,
            isShowingFormulas,
            setIsShowingFormulas,
            scrollActivePaneVertically,
            undoStack,
            redoStack,
            isOperationBusy,
            setPromptDialogState,
            isGridlinesVisible,
            setIsGridlinesVisible,
            isHeadingsVisible,
            setIsHeadingsVisible,
            zoomPercent,
            setIsFilterBarVisible,
            setIsFindBarVisible,
            hasFrozenRows,
            importFileInputRef,
            gridContainerRef,
            activeCell,
            clearSelectedCellContents,
            runEditorCommand,
            commitWorkbookEdit,
            runGuardedAction,
            applyCellPatches,
            applyZoomPercent,
            applyCellFormatPatch,
            undoOrRedo,
            exportWorkbookAs,
            replaceActiveWorkbook,
            toggleFreezeTopRow,
            copyOrCutSelection,
            pasteFromClipboard,
            autoFitSelectedAxis,
          }),
          i.jsx(ConsolidationStatusBanner, {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            options: { approvedWorkbooks: effectiveConsolidationWorkbooks },
            onReview: () => runEditorCommand("consolidation"),
          }),
          i.jsx(CustomFunctionReviewBanner, {
            workbook: activeWorkbook,
            extensions,
            onReview: () => runEditorCommand("extensions"),
          }),
          i.jsx(WorkbookLinkStatusBadge, {
            workbook: activeWorkbook,
            sources: effectiveLinkedWorkbooks,
            onReview: () => runEditorCommand("workbookLinks"),
          }),
          i.jsx(ExternalSourceStatusPanel, {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            onReview: () => runEditorCommand("externalSources"),
          }),
          i.jsx(AnalysisSnapshotNotice, {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            onReview: () => {
              return runEditorCommand(
                activeSheet.analysisReport?.request.kind === "solver"
                  ? "solver"
                  : "statistics",
              );
            },
          }),
          i.jsx(DataModelSnapshotNotice, {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            onReview: () => runEditorCommand("dataModel"),
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
                  onClick: () => runEditorCommand("calculateNow"),
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
                  onClick: () => runEditorCommand("calculation"),
                  children: i.jsx(LocalizedText, {
                    children: "Calculation Options",
                  }),
                }),
              ],
            }),
          (errorNotice || statusNotice) &&
            i.jsxs("div", {
              className: `lg-notice ${errorNotice ? "error" : ""}`,
              role: errorNotice ? "alert" : "status",
              children: [
                i.jsx(LocalizedText, { children: errorNotice || statusNotice }),
                i.jsx("button", {
                  "aria-label": text("Dismiss notification"),
                  onClick: () => {
                    (setErrorNotice(""), setStatusNotice(""));
                  },
                  lang: language("Dismiss notification"),
                  "data-ui-label": "Dismiss notification",
                  children: i.jsx(X, { size: 14 }),
                }),
              ],
            }),
          activeCellEditRef.current &&
            activeCellEditRef.current.locale !== workbookMessages.entryLocale &&
            i.jsx("p", {
              role: "status",
              className: "lg-notice",
              children: i.jsx(LocalizedText, {
                children:
                  "The current edit keeps its original numeric locale until you apply or cancel it.",
              }),
            }),
          persistenceError &&
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
                persistenceError instanceof WorkbookRevisionConflictError &&
                persistenceError.code === "conflict"
                  ? i.jsx("button", {
                      onClick: () => setIsSaveConflictDialogOpen(true),
                      children: i.jsx(LocalizedText, {
                        children: "Review save conflict",
                      }),
                    })
                  : i.jsx("button", {
                      disabled: readOnly || pendingSaveCount > 0,
                      onClick: recoverLocalWorkbook,
                      children: i.jsx(LocalizedText, {
                        children: "Retry saving",
                      }),
                    }),
                pendingSaveCount > 0 &&
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
                      children: "Waiting for save acknowledgement",
                    }),
                  }),
                i.jsx("button", {
                  onClick: exportLocalChangesJson,
                  children: i.jsx(LocalizedText, {
                    children: "Export my changes",
                  }),
                }),
              ],
            }),
          isSaveConflictDialogOpen &&
            i.jsx(WorkbookSaveConflictDialog, {
              workbook: activeWorkbook,
              loadSaved: () => workbookPersistence.load(),
              onExport: exportLocalChangesJson,
              onReload: reloadSavedWorkbook,
              onClose: () => setIsSaveConflictDialogOpen(false),
            }),
          i.jsx(ConnectionStatusNotice, {}),
          i.jsxs("div", {
            className: "lg-workarea",
            children: [
              renderWorkbookGrid({
                readOnly,
                setIsWorkbookHealthDialogOpen,
                text,
                language,
                editorIdPrefix,
                activeWorkbook,
                activeSheetId,
                activeSheet,
                isEditingLocked,
                setActiveDialogMode,
                appRootRef,
                setTabBarSheetId,
                formulaInputRef,
                formulaSuggestionListRef,
                formulaCaretPosition,
                setFormulaCaretPosition,
                isFormulaBarFocused,
                setIsFormulaBarFocused,
                activeCellAddress,
                completedSelectionAreas,
                selectionAreas,
                goToReferenceText,
                setGoToReferenceText,
                cellEditText,
                activeCellEditRef,
                formulaPointerStateRef,
                activeFormulaReference,
                isFormulaPointMode,
                filterValueSearch,
                setFilterValueSearch,
                setFilterStateByScope,
                setAuditPopulationScope,
                filterStateKey,
                activeFilterState,
                activeColumnSort,
                sheetSplitPanes,
                splitPaneSizes,
                scrollActivePaneVertically,
                isOperationBusy,
                activeOperationControllerRef,
                setPromptDialogState,
                zoomPercent,
                isFilterBarVisible,
                setIsFilterBarVisible,
                isFindBarVisible,
                setIsFindBarVisible,
                findQuery,
                setFindQuery,
                replaceWithText,
                setReplaceWithText,
                fillDragState,
                clipboardSnapshotState,
                clearCutSession,
                workbookCalculator,
                filterApplication,
                filterResult,
                isFilterStale,
                gridContainerRef,
                findInputRef,
                latestEditStateRef,
                selectedCellCount,
                activeCellFormulaErrors,
                activeCellReferenceDiagnostics,
                activeCellNamedFormulaDiagnostics,
                selectedNumericValues,
                selectedNumericSum,
                activeCell,
                matchingRowCount,
                runEditorCommand,
                commitWorkbookEdit,
                runGuardedAction,
                applyCellPatches,
                syncFormulaEditorSelection,
                isFormulaEditingLocked,
                handleFormulaKeyDown,
                beginCellEdit,
                focusGridSurface,
                cancelCellEdit,
                commitCellEdit,
                applyZoomPercent,
                switchActiveWorksheet,
                commitNameBoxReference,
                findNextTextMatch,
                replaceTextInSelection,
                moveSelectionAfterCommit,
                renderGridPane,
              }),
              renderSidePanel({
                adapter,
                text,
                language,
                activeWorkbook,
                activeSheet,
                isEditingLocked,
                setActiveCellAddress,
                selectionAreasKey,
                activeCellEditRef,
                sidePanelKind,
                setSidePanelKind,
                filterValueSearch,
                setFilterValueSearch,
                setFilterStateByScope,
                auditPopulationScope,
                setAuditPopulationScope,
                filterStateKey,
                activeFilterState,
                activeColumnSort,
                setActiveColumnSort,
                setPendingScrollAddress,
                setStatusNotice,
                setErrorNotice,
                selectedAuditKind,
                setSelectedAuditKind,
                auditAnalysisResult,
                setAuditAnalysisResult,
                materialityThreshold,
                setMaterialityThreshold,
                auditSampleSize,
                setAuditSampleSize,
                samplingSeed,
                setSamplingSeed,
                agingReportPayload,
                agingAsOfDate,
                setAgingAsOfDate,
                auditSourceRange,
                setAuditSourceRange,
                assistantSession,
                setAssistantSession,
                assistantProposal,
                setAssistantProposal,
                answerCitations,
                setAnswerCitations,
                isAssistantBusy,
                setIsAssistantBusy,
                filterResult,
                assistantAbortControllerRef,
                assistantRequestIdRef,
                latestEditStateRef,
                selectedCellCount,
                assistantSelection,
                selectionReferenceLabel,
                runEditorCommand,
                runGuardedAction,
                applyCellPatches,
                focusGridSurface,
                selectCellAddress,
                switchActiveWorksheet,
                runCurrentAuditAnalysis,
                askAssistantQuestion,
              }),
            ],
          }),
        ],
      }),
      isWorkbookHealthDialogOpen &&
        i.jsx(WorkbookHealthDialog, {
          monitor: activeReliabilityMonitor,
          onClose: () => setIsWorkbookHealthDialogOpen(false),
        }),
      validationReviewState &&
        i.jsx(ValidationReviewDialog, {
          error: validationReviewState.error,
          readOnly: isEditingLocked,
          onRetry: () => {
            (setValidationReviewState(null),
              requestAnimationFrame(() => {
                return formulaInputRef.current?.focus();
              }));
          },
          onCancel: () => {
            (setValidationReviewState(null),
              setErrorNotice(""),
              cancelCellEdit());
          },
          onAccept: () =>
            runGuardedAction(() => {
              if (latestEditStateRef.current.readOnly)
                throw Error("This workbook is read-only.");
              if (
                validationReviewState.selectionGuard &&
                validationReviewState.selectionGuard !==
                  latestEditStateRef.current.selectionKey
              )
                throw Error(
                  "Approved selection changed. Review the proposal again.",
                );
              if (
                validationReviewState.snapshot !==
                JSON.stringify(latestEditStateRef.current.book)
              )
                throw Error("Workbook changed. Review the batch again.");
              (commitWorkbookChange(
                applyCellPatchTransaction(
                  latestEditStateRef.current.book,
                  validationReviewState.patches,
                  actor,
                  validationReviewState.action,
                  {
                    validationApproval: assertNoStopValidationIssues(
                      validationReviewState.error,
                    ),
                  },
                ),
              ),
                setValidationReviewState(null),
                (activeCellEditRef.current = null),
                setCellEditState(null),
                setStatusNotice(
                  "Validation warnings accepted for the complete batch.",
                ),
                focusGridSurface());
            }),
        }),
      promptDialogState &&
        i.jsx("div", {
          className: "lg-modal-backdrop",
          onKeyDown: (fe) => {
            fe.key === "Escape" && setPromptDialogState(null);
          },
          onMouseDown: (fe) => {
            fe.target === fe.currentTarget && setPromptDialogState(null);
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
                onClick: () => setPromptDialogState(null),
                lang: language("Close dialog"),
                "data-ui-label": "Close dialog",
                children: i.jsx(X, { size: 18 }),
              }),
              i.jsx("h2", {
                id: `${editorIdPrefix}-dialog-title`,
                children:
                  promptDialogState.kind === "note"
                    ? "Add a cell note"
                    : promptDialogState.kind === "rename"
                      ? "Rename workbook"
                      : "Add worksheet",
              }),
              promptDialogState.kind === "note"
                ? i.jsx("textarea", {
                    autoFocus: true,
                    "aria-label": text("Cell note"),
                    rows: 5,
                    value: promptDialogState.value,
                    onChange: (fe) =>
                      setPromptDialogState({
                        ...promptDialogState,
                        value: fe.target.value,
                      }),
                    lang: language("Cell note"),
                    "data-ui-label": "Cell note",
                  })
                : i.jsx("input", {
                    autoFocus: true,
                    "aria-label": text("Name"),
                    value: promptDialogState.value,
                    onChange: (fe) =>
                      setPromptDialogState({
                        ...promptDialogState,
                        value: fe.target.value,
                      }),
                    onKeyDown: (fe) => {
                      (fe.key === "Enter" && commitDialogInput(),
                        fe.key === "Escape" && setPromptDialogState(null));
                    },
                    lang: language("Name"),
                    "data-ui-label": "Name",
                  }),
              i.jsxs("div", {
                className: "lg-modal-actions",
                children: [
                  i.jsx("button", {
                    className: "lg-button",
                    onClick: () => setPromptDialogState(null),
                    children: i.jsx(LocalizedText, { children: "Cancel" }),
                  }),
                  i.jsx("button", {
                    className: "lg-button dark",
                    onClick: commitDialogInput,
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
