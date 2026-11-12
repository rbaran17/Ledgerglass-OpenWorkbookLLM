import { applyCellPreview } from "./04-cell-input-paste-and-fill.js";
import { saveBlobToFile } from "./13-consolidation-dialog.js";
import { createUuid, isFiniteNumber } from "./03-workbook-model.js";
import i from "react/jsx-runtime";
import ne from "react";
import {
  Check,
  FileSpreadsheet,
  Search,
  ArrowUpRight,
  X,
  Filter,
  Plus,
  Layers,
  ClipboardCheck,
  Sparkles,
  History,
  CircleHelp,
  PanelRightClose,
  ListFilter,
  Copy,
  FlaskConical,
  ChartColumn,
  Clock3,
  Download,
  ShieldCheck,
} from "lucide-react";
import { isDeepFrozenValue } from "./01-core-schemas-and-hashing.js";
import {
  maxRowHeight,
  workbookLimits,
  minRowHeight,
  maxColumnWidth,
  minColumnWidth,
} from "./02-financial-functions.js";
import {
  formatNumberWithPattern,
  addWorksheet,
  findDataValidationsInRange,
  defaultRowHeightOf,
  cellStyleToCss,
  evaluateConditionalFormatRules,
  setFreezePanes,
  setSplitPanes,
  describeNamedStyleFallback,
  renameWorkbook,
  enumerateSelectionCells,
  defaultColumnWidthOf,
  Pv,
  findOccupiedRangeBounds,
  textWrapToCss,
  getImportedNumberFormatWarning,
  carryCalculatorRecordForward,
  formatOccupiedRangeAddress,
  calculateWorkbookWithSettings,
  requireRetainedXlsxSource,
  _t,
  anchorsMatchOnAxis,
  collectStyleFallbackWarnings,
  cn,
  parseRangeAddress,
  applyFilterConditions,
  parseSelectionAreas,
  touchWorkbookModifiedAt,
  assertNoStopValidationIssues,
  DataValidationError,
  applyAutoSum,
  lD,
  applyCellPatchTransaction,
  buildCanonicalWorkbookSnapshot,
  createBlankWorkbook,
  createAxisVisibilityIndex,
  columnIndexToLetters,
  vx,
  countSelectionCells,
  resolveEffectiveCellStyle,
} from "./03-workbook-model.js";
import {
  buildFillPreview,
  $ae,
  buildClipboardSnapshot,
  WorkbookMessagesProvider,
  useWorkbookMessages,
  parseDateCellEntry,
  withExtensionHostScope,
  findFormulaReferenceSpan,
  createEmptyAssistantConversation,
  LocalizedText,
  parseLocalizedCellEntry,
  replaceFormulaReferenceSpan,
  formatCellEntryForEdit,
  LocalizedNumberField,
  buildPastePreview,
  formatLocalizedNumberDisplay,
  autoCloseFormulaParens,
} from "./04-cell-input-paste-and-fill.js";
import {
  CustomFunctionReviewBanner,
  AddInsAndCustomFunctionsDialog,
} from "./05-add-ins-and-custom-functions.js";
import {
  importWorkbookFile,
  FileProcessingError,
  toFileProcessingError,
  importDelimitedFile,
  importAutomationFile,
} from "./07-xlsx-archive-reader.js";
import {
  encodeNativeWorkbook,
  exportWorksheetToCsv,
  exportWorkbookToXlsx,
  x7,
} from "./08-xlsx-export-and-preservation.js";
import {
  useOperationScope,
  WorkbookLinksDialog,
  buildSourceImportRecord,
  WorkbookLinkStatusBadge,
  operationScopeContext,
} from "./09-workbook-links.js";
import {
  SourceImportDialog,
  XlsxConversionReview,
} from "./10-source-import-dialog.js";
import {
  FormatConversionDialog,
  WorkbookRevisionConflictError,
  createWorkbookSession,
} from "./11-workbook-formats-dialog.js";
import {
  ConsolidationDialog,
  ConsolidationStatusBanner,
} from "./13-consolidation-dialog.js";
import {
  DataModelSnapshotNotice,
  DataModelDialog,
  useDialogFocusTrap,
} from "./15-data-model-dialog.js";
import { TransformationRecipesDialog } from "./18-transformation-recipes-dialog.js";
import {
  ValidationReviewDialog,
  FilterViewsDialog,
} from "./19-filter-views-dialog.js";
import {
  GridContextMenu,
  validateProposedCalculation,
  WorkbookCapacityLimits,
  CalculationOptionsDialog,
  buildAssistantSelectionContext,
} from "./20-calculation-options-and-sources.js";
import {
  AgingScheduleScreen,
  requestAssistantAnswer,
} from "./21-aging-and-materiality-screens.js";
import {
  parseGoToReference,
  ConnectionStatusNotice,
  ThemesAndCellStylesDialog,
  CellValuePreview,
} from "./22-themes-and-cell-styles.js";
import {
  ConditionalFormatIndicator,
  FormatCellsDialog,
  ConditionalFormattingRulesManager,
} from "./23-format-cells-dialog.js";
import {
  useCalculationRun,
  MergeCellsPreviewDialog,
  buildSampleLedgerWorkbook,
  runAuditAnalysis,
  recordCrossWorkbookImport,
  applyWorkbookChange,
} from "./24-merge-and-audit-analysis.js";
import { WorkbookRibbon } from "./26-workbook-ribbon.js";
import { DuplicateReviewDialog } from "./28-duplicate-review-and-structure.js";
import {
  PivotResultsStatusBar,
  DataToolDialogRouter,
} from "./29-conditional-formatting-dialog.js";
import { TableToolsDialog } from "./30-table-and-pivot-builder.js";
import { PivotExplorerDialog } from "./31-pivot-explore-and-slicers.js";
import {
  DrawingObjectOverlay,
  PicturePropertiesDialog,
} from "./32-drawing-and-chart-export.js";
import { PageSetupDialog } from "./33-page-setup-and-print.js";
import { FunctionAndNameDialog } from "./34-function-browser-and-names.js";
import {
  FormulaCompatibilityDialog,
  CalculationInspectorDialog,
  EvaluateFormulaDialog,
} from "./35-formula-compatibility-report.js";
import {
  FormulaWatchWindow,
  GoalSeekDialog,
} from "./36-watch-window-and-goal-seek.js";
import {
  AnalysisSnapshotNotice,
  DataAnalysisDialog,
} from "./38-data-analysis-dialog.js";
import { ReportOutputsDialog } from "./39-report-outputs.js";
import {
  WhatIfAnalysisDialog,
  DelimitedImportExportDialog,
} from "./40-what-if-analysis.js";
import {
  AutomationScriptsDialog,
  removeAutomationArtifact,
  runApprovedMigration,
  preserveAutomationArtifact,
} from "./41-automation-and-scripts.js";
import {
  CombinePopulationsDialog,
  TextToColumnsDialog,
} from "./42-combine-populations.js";
import { TextCleanupDialog } from "./43-text-cleanup.js";
import { CleanAndReshapeDialog } from "./44-clean-and-reshape.js";
import {
  WorksheetTabBar,
  CustomSortDialog,
} from "./45-sorting-and-worksheets.js";
import { WorkbookManagerDialog } from "./46-workbook-manager.js";
import {
  createCutSourceSnapshot,
  FillPreviewDialog,
  PasteSpecialDialog,
} from "./47-fill-and-paste-special.js";
import { MoveCellsDialog, M4e } from "./48-format-cells-and-move.js";
import {
  ReviewToolsPanel,
  OutlineToolsPanel,
} from "./49-outlines-and-comments.js";
import { MaterialityRiskScreen } from "./50-materiality-screen.js";
import {
  buildAnswerCitations,
  AssistantAnswerEvidence,
  PatternCheckDialog,
  evaluateAnswerEvidence,
} from "./51-pattern-checks-and-citations.js";
import { WorkbookComparisonDialog } from "./52-workbook-comparison.js";
import {
  SplitPaneLayout,
  buildVisibleAxisLayout,
  useSplitPaneSizes,
  autoFitColumnWidths,
  createTextMeasurer,
  Tp,
  sliceVisibleAxis,
  autoFitRowHeights,
  isWithinFormulaSuggestions,
  FormulaSuggestionList,
  logicalToPhysicalPixels,
  resolveCellNavigation,
  FindReplaceDialog,
} from "./53-find-replace-go-to.js";
import {
  instrumentWorkbookPersistence,
  computeVirtualizedRange,
  WorkbookSaveConflictDialog,
  WorkbookHealthDialog,
  createWorkbookDiagnosticsMonitor,
  saveWorkbookWithLock,
  findRecoverableCheckpoint,
} from "./54-workbook-save-conflicts.js";
import {
  isPromiseLike,
  ExternalSourceStatusPanel,
  workbookInstanceKey,
  chainMaybePromise,
  Vj,
  ExternalDataRefreshDialog,
  LazyAssistantChat,
  LazyAssistantChatMarkdown,
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
  var sheetProtectionDetails,
    ribbonSheetProtectionDetails,
    workbookStructureDetails,
    formulaInputNode,
    workbookStructureLock;
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
  function selectSingleRange(fe) {
    (setActiveSelectionRange(fe), setCompletedSelectionAreas([]));
  }
  function applySelectionAreas(fe) {
    const Oe = parseSelectionAreas(fe);
    (setCompletedSelectionAreas(Oe.slice(0, -1)),
      setActiveSelectionRange(Oe.at(-1)));
  }
  const selectionAreas = ne.useMemo(
      () => [...completedSelectionAreas, activeSelectionRange],
      [completedSelectionAreas, activeSelectionRange],
    ),
    selectionAreaRanges = ne.useMemo(
      () => selectionAreas.map(parseRangeAddress),
      [selectionAreas],
    ),
    selectionAreasKey = selectionAreas.join(", ");
  function requireSingleSelectionRange() {
    if (selectionAreas.length > 1)
      throw Error(
        "This command requires one rectangular range. Click a cell or enter a single range to continue.",
      );
  }
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
  function scrollActivePaneVertically(fe) {
    applyPaneScrollOffsets(
      splitPaneSizes.active.current,
      fe,
      splitPaneSizes.current.current[splitPaneSizes.active.current].left,
    );
  }
  function scrollActivePaneHorizontally(fe) {
    applyPaneScrollOffsets(
      splitPaneSizes.active.current,
      splitPaneSizes.current.current[splitPaneSizes.active.current].top,
      fe,
    );
  }
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
  function clearCutSession() {
    (setClipboardSnapshotState(null),
      setMoveCellsSource(null),
      (internalClipboardRef.current = null));
  }
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
  function runCalculationNow() {
    (calculationRun.run(),
      setStatusNotice(
        calculationRun.isolated
          ? "Isolated calculation requested. Source cells and audit history are unchanged."
          : "Calculation refreshed. Source cells and audit history are unchanged.",
      ));
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
  function computeGridViewport(fe) {
    const Oe = Math.max(
        1,
        fe.height -
          ((isHeadingsVisible ? 22 : 0) + frozenRowsHeight) * zoomScale,
      ),
      Be = Math.max(
        1,
        fe.width - (rowHeaderWidth + frozenColumnsWidth) * zoomScale,
      ),
      Ue = computeVirtualizedRange(
        scrollableRowLayout,
        fe.top,
        Oe,
        zoomScale,
        6,
      ),
      Xe = computeVirtualizedRange(
        scrollableColumnLayout,
        fe.left,
        Be,
        zoomScale,
        3,
      ),
      gt = Array.from({ length: Ue.last - Ue.first }, (Mr, Br) =>
        logicalRowAt(Ue.first + Br),
      ),
      Nt = Array.from({ length: Xe.last - Xe.first }, (Mr, Br) =>
        columnVisibilityIndex.logicalIndex(Xe.first + Br + yn),
      ),
      wt = Xe.leading,
      jt = Xe.trailing,
      rr =
        Nt.length +
        frozenColumnIndexes.length +
        Number(isHeadingsVisible) +
        +(wt > 0) +
        +(jt > 0);
    return {
      rowViewport: Oe,
      columnViewport: Be,
      rowWindow: Ue,
      columnWindow: Xe,
      shownRows: gt,
      shownColumns: Nt,
      leftGap: wt,
      rightGap: jt,
      visibleColumnCount: rr,
    };
  }
  const {
    rowViewport,
    columnViewport,
    rowWindow,
    columnWindow,
    shownRows,
    shownColumns,
  } = computeGridViewport(activePaneLayout);
  function applyPaneScrollOffsets(fe, Oe, Be) {
    const Ue = computeGridViewport(splitPaneSizes.current.current[fe]),
      Xe = computeVirtualizedRange(
        scrollableRowLayout,
        Oe,
        Ue.rowViewport,
        zoomScale,
        0,
      ).logicalOffset,
      gt = computeVirtualizedRange(
        scrollableColumnLayout,
        Be,
        Ue.columnViewport,
        zoomScale,
        0,
      ).logicalOffset,
      Nt = {};
    for (const wt of Pv(sheetSplitPanes)) {
      const jt = computeGridViewport(splitPaneSizes.current.current[wt]),
        rr = splitPaneSizes.current.current[wt],
        Mr = anchorsMatchOnAxis(fe, wt, "row")
          ? logicalToPhysicalPixels(
              Xe,
              scrollableRowLayout.total,
              jt.rowViewport,
              zoomScale,
            )
          : rr.top,
        Br = anchorsMatchOnAxis(fe, wt, "column")
          ? logicalToPhysicalPixels(
              gt,
              scrollableColumnLayout.total,
              jt.columnViewport,
              zoomScale,
            )
          : rr.left;
      Nt[wt] = { top: Mr, left: Br };
      const xr = splitPaneSizes.refs.current[wt];
      xr &&
        (Math.abs(xr.scrollTop - Mr) > 0.5 && (xr.scrollTop = Mr),
        Math.abs(xr.scrollLeft - Br) > 0.5 && (xr.scrollLeft = Br));
    }
    splitPaneSizes.patch(Nt);
  }
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
  function getSelectedCellAddresses() {
    return enumerateSelectionCells(selectionAreas);
  }
  function clearSelectedCellContents() {
    applyCellPatches(
      getSelectedCellAddresses().map((fe) => ({
        sheetId: activeSheet.id,
        address: fe,
        cell: { ...activeSheet.cells[fe], raw: null, literal: false },
      })),
      "Clear contents",
    );
  }
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
  function buildWorkbookViewState(
    fe,
    Oe = activeSheet.id,
    Be = activeSelectionRange,
  ) {
    const Ue =
      fe.sheets.find((Xe) => Xe.id === Oe && !Xe.hidden) ??
      fe.sheets.find((Xe) => !Xe.hidden);
    return {
      activeSheetId: Ue.id,
      selection: { sheetId: Ue.id, range: Ue.id === Oe ? Be : "A1" },
    };
  }
  function applyWorkbookView(fe) {
    const [Oe, Be = Oe] = fe.selection.range.split(":");
    (setActiveSheetId(fe.activeSheetId),
      setActiveCellAddress(Oe),
      setSelectionAnchorAddress(Oe),
      selectSingleRange(fe.selection.range),
      (activeCellAddressRef.current = Be),
      setActiveColumnSort(null),
      setFilterValueSearch(""),
      scrollActivePaneVertically(0),
      scrollActivePaneHorizontally(0),
      setPendingScrollAddress(Oe));
  }
  function reportSaveStatus(fe, Oe, Be) {
    try {
      notifyHostCallback("save-state", () => {
        return setSaveAbortControllerRef.current?.call(
          setSaveAbortControllerRef,
          {
            status: fe,
            workbookId: Oe.id,
            revision: Oe.revision,
            ...(Be ? { message: Be } : {}),
          },
        );
      });
    } catch {
      setErrorNotice(
        "The host could not acknowledge the save status. Export JSON to retain your work.",
      );
    }
  }
  function notifyHostCallback(fe, Oe) {
    const Be = () =>
      setErrorNotice(
        "The application could not acknowledge this operation. Your local workbook is retained.",
      );
    try {
      const Ue = activeReliabilityMonitor.run("host", fe, Oe);
      Ue && typeof Ue.then == "function" && Promise.resolve(Ue).catch(Be);
    } catch {
      Be();
    }
  }
  function reportSaveError(fe) {
    const Oe =
        fe instanceof WorkbookRevisionConflictError
          ? fe
          : new WorkbookRevisionConflictError(
              "unavailable",
              "The storage adapter is unavailable. The workbook could not be saved.",
            ),
      Be =
        Oe instanceof WorkbookRevisionConflictError && Oe.code === "conflict"
          ? Oe
          : (persistenceSessionRef.current.conflict ?? Oe);
    (Be instanceof WorkbookRevisionConflictError &&
      Be.code === "conflict" &&
      (persistenceSessionRef.current.conflict = Be),
      setPersistenceError(Be),
      setStatusNotice(""),
      (persistenceMessageRef.current =
        Be.message +
        " Changes remain in this editor. Export JSON to keep a copy."),
      setErrorNotice(persistenceMessageRef.current),
      reportSaveStatus(
        Be instanceof WorkbookRevisionConflictError && Be.code === "conflict"
          ? "conflict"
          : "error",
        latestEditStateRef.current.book,
        Be.message,
      ));
  }
  function enqueueSaveOperation(fe) {
    const Oe = persistenceSessionRef.current;
    if (Oe.controllers.size >= 20) {
      const gt = new WorkbookRevisionConflictError(
        "limit",
        "Twenty saves are already waiting. Your latest changes remain in this editor; retry saving after the pending saves finish, or export your work.",
      );
      throw (reportSaveError(gt), gt);
    }
    const Be = new AbortController();
    Oe.controllers.add(Be);
    const Ue = () => {
        if (
          !Oe.active ||
          Oe !== persistenceSessionRef.current ||
          Be.signal.aborted ||
          !isEditableRef.current
        )
          throw new WorkbookRevisionConflictError(
            "unavailable",
            "The save was cancelled because this editor or its access changed.",
          );
        if (Oe.conflict) throw Oe.conflict;
        if (Oe.stalled)
          throw new WorkbookRevisionConflictError(
            "unavailable",
            "A previous save is still unresolved. Wait for it to finish, then retry or export your local changes.",
          );
        return fe(Be.signal);
      },
      Xe = (gt) => {
        throw (
          Oe.controllers.delete(Be),
          Oe.active &&
            Oe === persistenceSessionRef.current &&
            (setPendingSaveCount(Oe.controllers.size),
            Oe.controllers.size || (Oe.stalled = false),
            reportSaveError(gt)),
          gt
        );
      };
    try {
      const gt = Oe.tail ? Oe.tail.then(Ue) : Ue();
      if (!isPromiseLike(gt)) return (Oe.controllers.delete(Be), gt);
      setPendingSaveCount(Oe.controllers.size);
      const Nt = setTimeout(() => {
          if (!Oe.active || Oe !== persistenceSessionRef.current) return;
          ((Oe.stalled = true),
            activeReliabilityMonitor
              .start("persistence", "save")
              .finish("failed", "timeout"),
            reportSaveError(
              new WorkbookRevisionConflictError(
                "unavailable",
                "The save has not been acknowledged. Keep this editor open or export your local changes. Retrying is disabled until pending work finishes.",
              ),
            ));
        }, saveWarningDelayMs),
        wt = gt.then(
          (rr) => (
            clearTimeout(Nt),
            Oe.controllers.delete(Be),
            Oe.active &&
              Oe === persistenceSessionRef.current &&
              (setPendingSaveCount(Oe.controllers.size),
              Oe.controllers.size || (Oe.stalled = false)),
            rr
          ),
          (rr) => (clearTimeout(Nt), Xe(rr)),
        ),
        jt = wt.then(
          () => {},
          () => {},
        );
      return (
        (Oe.tail = jt),
        jt.then(() => {
          Oe.tail === jt && (Oe.tail = undefined);
        }),
        wt
      );
    } catch (gt) {
      return Xe(gt);
    }
  }
  function acknowledgeSavedWorkbook(fe, Oe) {
    if (
      !Oe.active ||
      Oe !== persistenceSessionRef.current ||
      ((persistedWorkbookRecordRef.current = fe),
      setPersistedWorkbookRecord(fe),
      !isEditableRef.current || Oe.conflict)
    )
      return fe;
    if (
      fe.workbook === latestEditStateRef.current.book ||
      JSON.stringify(fe.workbook) ===
        JSON.stringify(latestEditStateRef.current.book)
    ) {
      setPersistenceError(null);
      const Be = persistenceMessageRef.current;
      (setErrorNotice((Ue) => (Ue === Be ? "" : Ue)),
        (persistenceMessageRef.current = null),
        Oe.stalled &&
          activeReliabilityMonitor
            .start("persistence", "save")
            .finish("recovered"),
        setStatusNotice(
          persistenceKey
            ? "Saved to this browser"
            : workbookStore
              ? "Saved to the application"
              : "Changes kept in this session",
        ),
        reportSaveStatus("saved", fe.workbook));
    }
    return fe;
  }
  function recoverLocalWorkbook() {
    if (!(
      persistenceSessionRef.current.controllers.size || !isEditableRef.current
    ))
      try {
        const fe = latestEditStateRef.current.book,
          Oe = buildWorkbookViewState(fe),
          Be = findRecoverableCheckpoint(workbookPersistence, fe);
        if (Be) {
          ((isWorkbookLoadedRef.current = true),
            acknowledgeSavedWorkbook(Be, persistenceSessionRef.current),
            activeReliabilityMonitor
              .start("persistence", "save")
              .finish("recovered"));
          return;
        }
        if (!isWorkbookLoadedRef.current) {
          const Xe = workbookPersistence.load();
          if (Xe)
            throw new WorkbookRevisionConflictError(
              "conflict",
              "Saved data is available again. Compare it before replacing your local changes.",
              Xe.guard,
            );
          isWorkbookLoadedRef.current = true;
        }
        const Ue = persistWorkbookSnapshot(fe, Oe);
        isPromiseLike(Ue) && Ue.catch(() => {});
      } catch (fe) {
        reportSaveError(fe);
      }
  }
  function persistWorkbookSnapshot(fe, Oe = buildWorkbookViewState(fe)) {
    if (!isWorkbookLoadedRef.current) {
      const gt = new WorkbookRevisionConflictError(
        "corrupt",
        "Saved data needs recovery before this session can overwrite it.",
      );
      throw (reportSaveError(gt), gt);
    }
    const Be = workbookStore || persistenceKey ? structuredClone(fe) : fe,
      Ue = structuredClone(Oe),
      Xe = persistenceSessionRef.current;
    return (
      reportSaveStatus("saving", Be),
      enqueueSaveOperation((gt) => {
        return chainMaybePromise(
          workbookPersistence.save(
            Be,
            Ue,
            persistedWorkbookRecordRef.current?.guard ?? null,
            {
              allowIdentityChange: !!(
                persistedWorkbookRecordRef.current &&
                persistedWorkbookRecordRef.current.workbook.id !== Be.id
              ),
              signal: gt,
            },
          ),
          (wt) => acknowledgeSavedWorkbook(wt, Xe),
        );
      })
    );
  }
  function exportLocalChangesJson() {
    saveBlobToFile(
      new Blob([encodeNativeWorkbook(latestEditStateRef.current.book)], {
        type: "application/json",
      }),
      latestEditStateRef.current.book.name.replace(/[^a-zA-Z0-9 _.-]/g, "_") +
        " - local changes.json",
    );
  }
  function reloadSavedWorkbook(fe) {
    var Ue;
    const Oe = workbookPersistence.load();
    if (!Oe || JSON.stringify(Oe.guard) !== JSON.stringify(fe.guard))
      throw new Error(
        "The saved version changed again. Refresh the comparison before reloading.",
      );
    const Be = persistenceSessionRef.current;
    Be.active = false;
    for (const Xe of Be.controllers) Xe.abort();
    ((persistenceSessionRef.current = { active: true, controllers: new Set() }),
      (Ue = saveAbortControllerRef.current) == null || Ue.abort(),
      (persistedWorkbookRecordRef.current = Oe),
      setPersistedWorkbookRecord(Oe),
      (latestEditStateRef.current.book = Oe.workbook),
      setWorkbookSnapshot(buildCanonicalWorkbookSnapshot(Oe.workbook)),
      (activeCellEditRef.current = null),
      setCellEditState(null),
      setUndoStack([]),
      setRedoStack([]),
      setAssistantProposal(null),
      setAnswerCitations(null),
      setAuditAnalysisResult(null),
      applyWorkbookView(Oe.view),
      setIsSaveConflictDialogOpen(false),
      setPersistenceError(null),
      setErrorNotice(""),
      setStatusNotice(
        "Loaded the saved workbook. The discarded local undo history has been cleared.",
      ));
    try {
      notifyHostCallback("change", () => onChange?.(Oe.workbook));
    } catch {
      setErrorNotice("The host could not acknowledge the reloaded workbook.");
    }
    reportSaveStatus("saved", Oe.workbook);
  }
  function commitWorkbookChange(fe, Oe = true, Be = false, Ue, Xe = false, gt) {
    var Nt;
    if (
      readOnly ||
      ((isEditingLocked || latestEditStateRef.current.readOnly) &&
        !Be &&
        !(
          Xe &&
          gt &&
          latestEditStateRef.current.book.sheets.some((wt) => {
            var jt;
            return (
              wt.id === gt && !((jt = wt.protection) != null && jt.enabled)
            );
          })
        ))
    )
      throw new Error("This workbook is read-only.");
    if (!(
      !Ue &&
      (fe === latestEditStateRef.current.book ||
        (fe.revision === latestEditStateRef.current.book.revision &&
          JSON.stringify(fe) ===
            JSON.stringify(latestEditStateRef.current.book)))
    )) {
      if (
        (Ue || (Nt = saveAbortControllerRef.current) == null || Nt.abort(),
        (fe = buildCanonicalWorkbookSnapshot(
          Ue || Xe ? fe : touchWorkbookModifiedAt(fe),
        )),
        onBeforeChange == null ||
          onBeforeChange(fe, latestEditStateRef.current.book),
        Oe)
      ) {
        const wt = latestEditStateRef.current.book;
        (setUndoStack((jt) => [...jt.slice(-49), wt]), setRedoStack([]));
      }
      if (
        (carryCalculatorRecordForward(latestEditStateRef.current.book, fe),
        (latestEditStateRef.current.book = fe),
        setWorkbookSnapshot(fe),
        setAssistantProposal(null),
        setAnswerCitations(null),
        setAuditAnalysisResult(null),
        setErrorNotice(""),
        Ue)
      )
        ((persistedWorkbookRecordRef.current = Ue),
          setPersistedWorkbookRecord(Ue));
      else
        try {
          const wt = persistWorkbookSnapshot(fe);
          isPromiseLike(wt) && wt.catch(() => {});
        } catch (wt) {
          reportSaveError(wt);
        }
      try {
        notifyHostCallback("change", () => onChange?.(fe));
      } catch {
        setErrorNotice(
          "The host could not acknowledge the change. Export JSON to retain this revision.",
        );
      }
    }
  }
  function runEditorCommand(fe) {
    if (fe === "xlsxSource") {
      runGuardedAction(() =>
        setXlsxConversionReview({
          snapshot: activeWorkbook,
          report: buildSourceImportRecord(activeWorkbook),
          fileName: activeWorkbook.xlsxSource.fileName,
          inspectOnly: true,
        }),
      );
      return;
    }
    if (commitCellEdit()) {
      if (
        (fe === "populationFilter" && setActiveFilterColumn(undefined),
        selectionAreas.length > 1 &&
          ![
            "search",
            "numberFormat",
            "cellStyle",
            "themes",
            "goToSpecial",
            "comments",
            "hyperlink",
            "sheets",
            "workbookProperties",
            "versions",
            "tracePrecedents",
            "traceDependents",
            "formulaWatches",
            "evaluateFormula",
            "functionCompatibility",
            "calculation",
            "calculateNow",
            "automation",
            "externalSources",
            "workbookFormats",
            "importSources",
            "workbookLinks",
            "extensions",
          ].includes(fe))
      ) {
        setErrorNotice(
          "This command requires one rectangular range. Click a cell or enter a single range to continue.",
        );
        return;
      }
      if (fe === "calculateNow") {
        runCalculationNow();
        return;
      }
      if (fe === "pasteSpecial" && clipboardSnapshotState) {
        setMoveCellsSource(clipboardSnapshotState.snapshot);
        return;
      }
      (fe === "sheets" && setTabBarSheetId(activeSheet.id),
        fe === "automation" && setAutomationInspection(undefined),
        setActiveDialogMode(fe));
    }
  }
  function commitWorkbookEdit(fe, Oe) {
    var Nt;
    const Be = latestEditStateRef.current.book,
      Ue = applyWorkbookChange(Be, fe, actor, Oe),
      Xe = fe.sheets.find((wt) => !Be.sheets.some((jt) => jt.id === wt.id));
    (commitWorkbookChange(
      Ue,
      true,
      Oe === "Unprotect sheet" &&
        !(
          (Nt = fe.sheets.find((wt) => wt.id === activeSheet.id)?.protection) !=
            null && Nt.enabled
        ),
    ),
      setStatusNotice(Oe),
      Xe && (setActiveDialogMode(null), switchActiveWorksheet(Xe.id)));
  }
  function replaceWorkbookSnapshot(fe, Oe) {
    if (readOnly) throw new Error("This workbook is read-only.");
    const Be = latestEditStateRef.current.book;
    if (JSON.stringify(Be) === JSON.stringify(fe)) return;
    const Ue =
      JSON.stringify(Be.structureProtection) !==
      JSON.stringify(fe.structureProtection);
    (commitWorkbookChange(applyWorkbookChange(Be, fe, actor, Oe), true, true),
      Ue && (setUndoStack([]), setRedoStack([])),
      fe.sheets.some((Xe) => Xe.id === activeSheet.id && !Xe.hidden) ||
        applyWorkbookView(buildWorkbookViewState(fe)),
      setStatusNotice(Oe));
  }
  function commitWorkbookWithNewSheet(fe, Oe) {
    const Be = fe.sheets.find(
      (Ue) =>
        !latestEditStateRef.current.book.sheets.some((Xe) => Xe.id === Ue.id),
    );
    (replaceWorkbookSnapshot(fe, Oe),
      Be && (setActiveDialogMode(null), switchActiveWorksheet(Be.id)));
  }
  function runGuardedAction(fe) {
    try {
      fe();
    } catch (Oe) {
      setErrorNotice(
        Oe instanceof Error
          ? Oe.message
          : "The operation could not be completed.",
      );
    }
  }
  function applyCellPatches(fe, Oe = "Edit cells", Be, Ue) {
    try {
      if (Ue && fe.some((Xe) => Xe.sheetId !== Ue))
        throw Error("The formula destination changed.");
      commitWorkbookChange(
        applyCellPatchTransaction(
          latestEditStateRef.current.book,
          fe,
          actor,
          Oe,
        ),
        true,
        false,
        undefined,
        true,
        Ue,
      );
    } catch (Xe) {
      throw (
        Xe instanceof DataValidationError &&
          setValidationReviewState({
            error: Xe,
            patches: structuredClone(fe),
            action: Oe,
            snapshot: JSON.stringify(latestEditStateRef.current.book),
            selectionGuard: Be,
          }),
        Xe
      );
    }
  }
  function resetFormulaEditSession() {
    ((formulaPointerStateRef.current = {
      input: null,
      start: 0,
      end: 0,
      dragging: false,
      mode: true,
    }),
      setActiveFormulaReference(null),
      setIsFormulaPointMode(true));
  }
  function syncFormulaEditorSelection(fe, Oe = false) {
    const Be = formulaPointerStateRef.current,
      Ue = fe.selectionStart ?? fe.value.length,
      Xe = fe.selectionEnd ?? Ue;
    ((Be.input = fe),
      !(Be.restoring && !Oe) &&
        ((Be.restoring = false),
        (Be.start = Ue),
        (Be.end = Xe),
        (Oe || (Be.span && (Ue !== Be.span.end || Xe !== Ue))) &&
          ((Be.span = undefined),
          (Be.anchor = undefined),
          (Be.head = undefined),
          setActiveFormulaReference(null))));
  }
  function restoreFormulaEditorFocus() {
    ((formulaPointerStateRef.current.restoring = true),
      requestAnimationFrame(() => {
        var Be;
        if (!activeCellEditRef.current) return;
        const fe = formulaPointerStateRef.current,
          Oe =
            (Be = fe.input) != null && Be.isConnected
              ? fe.input
              : formulaInputRef.current;
        (Oe &&
          ((fe.input = Oe),
          Oe.focus({ preventScroll: true }),
          Oe.setSelectionRange(fe.start, fe.end)),
          (fe.restoring = false));
      }));
  }
  function isFormulaEditingLocked() {
    var Oe;
    const fe = latestEditStateRef.current.book.sheets.find((Be) => {
      return Be.id === (activeCellEditRef.current?.sheetId ?? activeSheet.id);
    });
    return readOnly || !fe || !!((Oe = fe.protection) != null && Oe.enabled);
  }
  function insertFormulaReference(fe, Oe = false, Be) {
    var rr;
    const Ue = activeCellEditRef.current,
      Xe = formulaPointerStateRef.current;
    if (
      !(Ue != null && Ue.value.startsWith("=")) ||
      isFormulaEditingLocked() ||
      isOperationBusy
    )
      return false;
    const gt = Xe.span ?? findFormulaReferenceSpan(Ue.value, Xe.start, Xe.end);
    if (!gt) return false;
    const Nt =
        Oe && Xe.sheetId === activeSheet.id && Xe.anchor ? Xe.anchor : fe,
      wt = Be?.range ?? (Nt === fe ? fe : wy(Nt, fe)),
      jt = replaceFormulaReferenceSpan(
        Ue.value,
        gt,
        wt,
        Ue.sheetId === activeSheet.id ? undefined : activeSheet.name,
      );
    return (
      beginCellEdit(
        jt.value,
        !!((rr = Xe.input) != null && rr.classList.contains("lg-cell-editor")),
      ),
      (Xe.span = jt.span),
      (Xe.start = Xe.end = jt.span.end),
      (Xe.anchor = Nt),
      (Xe.head = fe),
      (Xe.sheetId = activeSheet.id),
      setFormulaCaretPosition(jt.span.end),
      setActiveFormulaReference({
        sheetId: activeSheet.id,
        range: wt,
        bounds: Be?.bounds ?? parseRangeAddress(wt),
      }),
      restoreFormulaEditorFocus(),
      true
    );
  }
  formulaPointerMoveRef.current = (fe) => {
    formulaPointerStateRef.current.head !== fe &&
      insertFormulaReference(fe, true);
  };
  function handleFormulaPointerDown(fe) {
    var Nt;
    if (
      fe.button !== 0 ||
      !((Nt = activeCellEditRef.current) != null && Nt.value.startsWith("="))
    )
      return false;
    const Oe = fe.target;
    if (Oe.closest('input,button,a,[role="separator"]')) return false;
    const Be = Oe.closest("[data-address]"),
      Ue = Oe.closest("[data-column]"),
      Xe = Oe.closest("[data-row-header]");
    let gt = false;
    if (Be) gt = insertFormulaReference(Be.dataset.address, fe.shiftKey);
    else if (Ue) {
      const wt = Number(Ue.dataset.column);
      gt = insertFormulaReference(`${columnIndexToLetters(wt)}1`, false, {
        range: `${columnIndexToLetters(wt)}:${columnIndexToLetters(wt)}`,
        bounds: {
          firstRow: 0,
          lastRow: activeSheet.rowCount - 1,
          firstCol: wt,
          lastCol: wt,
        },
      });
    } else if (Xe) {
      const wt = Number(Xe.dataset.rowHeader);
      gt = insertFormulaReference(`A${wt + 1}`, false, {
        range: `${wt + 1}:${wt + 1}`,
        bounds: {
          firstRow: wt,
          lastRow: wt,
          firstCol: 0,
          lastCol: activeSheet.colCount - 1,
        },
      });
    }
    return gt
      ? (fe.preventDefault(),
        fe.stopPropagation(),
        (formulaPointerStateRef.current.dragging = !!Be),
        (dragAnchorRef.current = null),
        true)
      : false;
  }
  function handleFormulaKeyDown(fe) {
    if (
      fe.nativeEvent.isComposing ||
      fe.altKey ||
      isFormulaEditingLocked() ||
      isOperationBusy
    )
      return false;
    const Oe = fe.currentTarget;
    if (!Oe.value.startsWith("=")) return false;
    syncFormulaEditorSelection(Oe);
    const Be = formulaPointerStateRef.current;
    if (fe.key === "F2" && !fe.ctrlKey && !fe.metaKey)
      return (
        fe.preventDefault(),
        activeCellEditRef.current || beginCellEdit(Oe.value),
        (Be.mode = !Be.mode),
        setIsFormulaPointMode(Be.mode),
        true
      );
    if (fe.key === "F4" && !fe.ctrlKey && !fe.metaKey) {
      const jt = $ae(Oe.value, Be.start, Be.end);
      return jt
        ? (fe.preventDefault(),
          beginCellEdit(jt.value, Oe.classList.contains("lg-cell-editor")),
          (Be.span = jt.span),
          (Be.start = Be.end = jt.span.end),
          setFormulaCaretPosition(jt.span.end),
          restoreFormulaEditorFocus(),
          true)
        : false;
    }
    if ((fe.ctrlKey || fe.metaKey) && ["PageUp", "PageDown"].includes(fe.key)) {
      fe.preventDefault();
      const jt = activeWorkbook.sheets.filter((Br) => !Br.hidden),
        rr = jt.findIndex((Br) => Br.id === activeSheet.id),
        Mr = jt[rr + (fe.key === "PageUp" ? -1 : 1)];
      return (Mr && switchActiveWorksheet(Mr.id), true);
    }
    if (
      !Be.mode ||
      !fe.key.startsWith("Arrow") ||
      !activeCellEditRef.current ||
      !(Be.span ?? findFormulaReferenceSpan(Oe.value, Be.start, Be.end))
    )
      return false;
    fe.preventDefault();
    const Ue = _t(
        Be.sheetId === activeSheet.id && Be.head ? Be.head : activeCellAddress,
      ),
      Xe = ["ArrowUp", "ArrowDown"].includes(fe.key),
      gt = ["ArrowUp", "ArrowLeft"].includes(fe.key) ? -1 : 1,
      Nt =
        fe.ctrlKey || fe.metaKey
          ? jumpToDataEdge(Ue.row, Ue.col, Xe, gt)
          : {
              row: Xe ? stepVisibleRowIndex(Ue.row, gt) : Ue.row,
              col: Xe ? Ue.col : stepVisibleColumnIndex(Ue.col, gt),
            },
      wt = `${columnIndexToLetters(Nt.col)}${Nt.row + 1}`;
    return (
      insertFormulaReference(wt, fe.shiftKey),
      requestCellScroll(wt),
      true
    );
  }
  function switchToEditSheet(fe) {
    fe.sheetId !== activeSheet.id &&
      (setActiveSheetId(fe.sheetId),
      setActiveCellAddress(fe.address),
      setSelectionAnchorAddress(fe.address),
      selectSingleRange(fe.address),
      setActiveColumnSort(null),
      setFilterValueSearch(""),
      setPendingScrollAddress(fe.address));
  }
  function beginCellEdit(fe, Oe = false) {
    var Ue;
    const Be = activeCellEditRef.current
      ? undefined
      : workbookCalculator.getSpillAnchor(activeSheet.id, activeCellAddress);
    if (Be && Be !== activeCellAddress) {
      setErrorNotice(
        `This cell belongs to the array formula in ${Be}. Edit that cell instead.`,
      );
      return;
    }
    ((Ue = saveAbortControllerRef.current) == null || Ue.abort(),
      (activeCellEditRef.current = {
        sheetId: activeCellEditRef.current?.sheetId ?? activeSheet.id,
        address: activeCellEditRef.current?.address ?? activeCellAddress,
        value: fe,
        locale: activeCellEditRef.current
          ? activeCellEditRef.current.locale
          : workbookMessages.entryLocale,
        originalText:
          activeCellEditRef.current?.originalText ??
          getEditableCellText(activeSheet.cells[activeCellAddress]),
      }),
      setCellEditText(fe),
      Oe && setCellEditState(fe));
  }
  function focusGridSurface() {
    requestAnimationFrame(() => {
      return gridContainerRef.current?.focus({ preventScroll: true });
    });
  }
  function cancelCellEdit() {
    const fe = activeCellEditRef.current;
    (fe && switchToEditSheet(fe),
      resetFormulaEditSession(),
      (activeCellEditRef.current = null),
      setCellEditState(null),
      setCellEditText(
        fe?.originalText ??
          getEditableCellText(activeSheet.cells[activeCellAddress]),
      ),
      focusGridSurface());
  }
  function commitCellEdit(fe) {
    const Oe = activeCellEditRef.current;
    if (!Oe) return true;
    try {
      const Xe = autoCloseFormulaParens(Oe.value),
        gt = latestEditStateRef.current.book.sheets.find(
          (Nt) => Nt.id === Oe.sheetId,
        )?.cells[Oe.address];
      return (
        Xe !== Oe.originalText &&
          applyCellPatches(
            [
              {
                sheetId: Oe.sheetId,
                address: Oe.address,
                cell: parseLocalizedCellEntry(
                  Xe,
                  gt,
                  latestEditStateRef.current.book.datePolicy?.dateSystem,
                  Oe.locale,
                ),
              },
            ],
            "Edit cells",
            undefined,
            Oe.sheetId,
          ),
        (pendingCrossSheetTargetRef.current =
          Oe.sheetId !== activeSheet.id
            ? { sheetId: Oe.sheetId, address: Oe.address }
            : null),
        switchToEditSheet(Oe),
        resetFormulaEditSession(),
        (activeCellEditRef.current = null),
        setCellEditState(null),
        setCellEditText(Xe),
        true
      );
    } catch (Xe) {
      return (
        setErrorNotice(
          Xe instanceof Error ? Xe.message : "The cell could not be saved.",
        ),
        false
      );
    }
  }
  function requestCellScroll(fe) {
    setPendingScrollAddress(fe);
  }
  function applyZoomPercent(fe) {
    if (!Number.isFinite(fe) || fe < 50 || fe > 200) return;
    const Oe = fe / 100,
      Be = {};
    for (const Ue of Pv(sheetSplitPanes)) {
      const Xe = splitPaneSizes.current.current[Ue],
        gt = computeGridViewport(Xe),
        Nt = Math.max(
          1,
          Xe.height - ((isHeadingsVisible ? 22 : 0) + frozenRowsHeight) * Oe,
        ),
        wt = Math.max(1, Xe.width - (rowHeaderWidth + frozenColumnsWidth) * Oe),
        jt = logicalToPhysicalPixels(
          gt.rowWindow.logicalOffset,
          scrollableRowLayout.total,
          Nt,
          Oe,
        ),
        rr = logicalToPhysicalPixels(
          gt.columnWindow.logicalOffset,
          scrollableColumnLayout.total,
          wt,
          Oe,
        );
      Be[Ue] = { top: jt, left: rr };
      const Mr = splitPaneSizes.refs.current[Ue];
      Mr && ((Mr.scrollTop = jt), (Mr.scrollLeft = rr));
    }
    (splitPaneSizes.patch(Be),
      setZoomPercent(fe),
      setPendingScrollAddress(activeCellAddress));
  }
  function selectCellAddress(fe, Oe = false, Be = false) {
    if (!commitCellEdit()) return;
    if (!Oe && !filterValueSearch && !activeColumnSort) {
      const gt = (activeSheet.merges ?? []).find((Nt) => lD(Nt, fe));
      gt && (fe = gt.split(":")[0]);
    }
    const Ue = Oe ? wy(selectionAnchorAddress, fe) : fe;
    let Xe;
    try {
      Xe = parseSelectionAreas(Be ? [...selectionAreas, Ue] : [Ue]);
    } catch (gt) {
      setErrorNotice(gt.message);
      return;
    }
    (requestCellScroll(fe),
      (activeCellAddressRef.current = fe),
      setActiveCellAddress(Oe ? selectionAnchorAddress : fe),
      applySelectionAreas(Xe),
      Oe || setSelectionAnchorAddress(fe),
      setCellEditState(null));
  }
  function switchActiveWorksheet(fe) {
    var Oe;
    if (
      (Oe = activeCellEditRef.current) != null &&
      Oe.value.startsWith("=") &&
      fe !== activeSheet.id
    ) {
      (setActiveSheetId(fe),
        setActiveColumnSort(null),
        setFilterValueSearch(""),
        scrollActivePaneVertically(0),
        scrollActivePaneHorizontally(0),
        setCellEditState(null),
        (formulaPointerStateRef.current.anchor = undefined),
        (formulaPointerStateRef.current.head = undefined),
        setActiveFormulaReference(null),
        restoreFormulaEditorFocus());
      return;
    }
    commitCellEdit() &&
      ((activeCellAddressRef.current = "A1"),
      setActiveSheetId(fe),
      setActiveCellAddress("A1"),
      setSelectionAnchorAddress("A1"),
      selectSingleRange("A1"),
      setActiveColumnSort(null),
      setFilterValueSearch(""),
      scrollActivePaneVertically(0),
      scrollActivePaneHorizontally(0),
      setAuditAnalysisResult(null),
      setAuditSourceRange(""),
      gridContainerRef.current &&
        ((gridContainerRef.current.scrollTop = 0),
        (gridContainerRef.current.scrollLeft = 0)),
      focusGridSurface());
  }
  function navigateToSheetCell(fe, Oe) {
    const Be = latestEditStateRef.current.book.sheets.find(
      (Xe) => Xe.id === fe,
    );
    if (!Be || Be.hidden)
      throw Error(
        "Unhide the target worksheet with Manage worksheets before navigating to it.",
      );
    const Ue = _t(Oe);
    if (
      createAxisVisibilityIndex(Be.rowCount, cn(Be, "row")).isHidden(Ue.row) ||
      createAxisVisibilityIndex(Be.colCount, cn(Be, "column")).isHidden(Ue.col)
    )
      throw Error(
        "The target cell is hidden. Unhide its row/column or expand its outline before navigating.",
      );
    commitCellEdit() &&
      (setFilterStateByScope((Xe) => {
        const gt = { ...Xe };
        return (delete gt[JSON.stringify([activeWorkbook.id, fe])], gt);
      }),
      setActiveSheetId(fe),
      setActiveCellAddress(Oe),
      setSelectionAnchorAddress(Oe),
      (activeCellAddressRef.current = Oe),
      selectSingleRange(Oe),
      setFilterValueSearch(""),
      setActiveColumnSort(null),
      setPendingScrollAddress(Oe));
  }
  function selectSheetRanges(fe, Oe) {
    const Be = parseSelectionAreas(Oe);
    (navigateToSheetCell(fe, Be.at(-1).split(":")[0]),
      applySelectionAreas(Be),
      setActiveDialogMode(null),
      focusGridSurface());
  }
  function applyCellFormatPatch(fe) {
    runGuardedAction(() => {
      const Oe = getSelectedCellAddresses();
      if (
        fe.style &&
        (Object.hasOwn(fe.style, "border") ||
          Object.hasOwn(fe.style, "borders")) &&
        Oe.some((Be) => (activeSheet.merges ?? []).some((Ue) => lD(Ue, Be)))
      )
        throw Error("Unmerge selected cells before changing border edges.");
      applyCellPatches(
        Oe.map((Be) => {
          return {
            sheetId: activeSheet.id,
            address: Be,
            cell: {
              ...activeSheet.cells[Be],
              raw: activeSheet.cells[Be]?.raw ?? null,
              ...fe,
              ...(fe.style
                ? {
                    style: {
                      ...activeSheet.cells[Be]?.style,
                      ...fe.style,
                    },
                  }
                : {}),
            },
          };
        }),
        "Format selection",
      );
    });
  }
  function commitNameBoxReference() {
    runGuardedAction(() => {
      if (!commitCellEdit()) return;
      const fe = parseGoToReference(
          activeWorkbook,
          activeSheet.id,
          goToReferenceText,
        ),
        Oe = fe.ranges,
        Be = parseRangeAddress(Oe.at(-1)),
        Ue = `${columnIndexToLetters(Be.firstCol)}${Be.firstRow + 1}`,
        Xe = `${columnIndexToLetters(Be.lastCol)}${Be.lastRow + 1}`;
      if (
        fe.sheetId === activeSheet.id &&
        (rowVisibilityIndex.isHidden(Be.firstRow) ||
          columnVisibilityIndex.isHidden(Be.firstCol))
      )
        throw Error(
          `${Ue} is hidden. Use View → Hide & Unhide or Data → Group & Outline to reveal it.`,
        );
      (fe.sheetId !== activeSheet.id && navigateToSheetCell(fe.sheetId, Ue),
        setActiveCellAddress(Ue),
        setSelectionAnchorAddress(Ue),
        (activeCellAddressRef.current = Xe),
        applySelectionAreas(Oe),
        fe.sheetId === activeSheet.id && requestCellScroll(Ue));
    });
  }
  function undoOrRedo(fe = false) {
    runGuardedAction(() => {
      if (isEditingLocked || latestEditStateRef.current.readOnly)
        throw new Error("This workbook is read-only.");
      const Oe = fe
        ? redoStack[redoStack.length - 1]
        : undoStack[undoStack.length - 1];
      if (!Oe) return;
      if (
        Oe.id === activeWorkbook.id &&
        JSON.stringify(Oe.structureProtection) !==
          JSON.stringify(activeWorkbook.structureProtection)
      )
        throw new Error(
          "Undo cannot change workbook structure protection. Use Manage worksheets to unlock it.",
        );
      const Be =
        Oe.id !== activeWorkbook.id
          ? recordCrossWorkbookImport(
              activeWorkbook,
              Oe,
              actor,
              fe ? "redo" : "undo",
            )
          : applyWorkbookChange(
              activeWorkbook,
              {
                ...(isDeepFrozenValue(Oe) ? Oe : structuredClone(Oe)),
                ...(activeWorkbook.rowOrigins
                  ? { rowOrigins: structuredClone(activeWorkbook.rowOrigins) }
                  : {}),
                revision: activeWorkbook.revision,
                events: activeWorkbook.events,
              },
              actor,
              fe ? "Redo workbook change" : "Undo workbook change",
              { restoringHistory: true },
            );
      (commitWorkbookChange(Be, false),
        fe
          ? (setRedoStack(redoStack.slice(0, -1)),
            setUndoStack([...undoStack, activeWorkbook]))
          : (setUndoStack(undoStack.slice(0, -1)),
            setRedoStack([...redoStack, activeWorkbook])),
        Be.sheets.some((Ue) => Ue.id === activeSheetId && !Ue.hidden) ||
          switchActiveWorksheet(Be.sheets.find((Ue) => !Ue.hidden).id));
    });
  }
  async function importWorkbookFromFile(fe) {
    var gt;
    if (!fe) return;
    (gt = activeOperationControllerRef.current) == null || gt.abort();
    const Oe = new AbortController(),
      Be = ++latestOperationIdRef.current;
    activeOperationControllerRef.current = Oe;
    const Ue = {
        ...operationScope,
        signal: operationScope.signal
          ? AbortSignal.any([Oe.signal, operationScope.signal])
          : Oe.signal,
      },
      Xe = () => Be === latestOperationIdRef.current;
    (setIsOperationBusy(true), setErrorNotice(""));
    try {
      if (isEditingLocked) throw new Error("This workbook is read-only.");
      if (/\.(xls|xlsb|xltx|ods)$/i.test(fe.name)) {
        (setPendingImportFile(fe), setActiveDialogMode("workbookFormats"));
        return;
      }
      if (/\.(bas|vba|cls|ts|osts|xlsm|xlam|xltm)$/i.test(fe.name)) {
        if (fe.size > workbookLimits.maxFileBytes)
          throw new Error("Files must be 10 MB or smaller.");
        const wt = await importAutomationFile(fe, Ue);
        if (!Xe()) return;
        if (
          latestEditStateRef.current.book !== activeWorkbook ||
          latestEditStateRef.current.readOnly
        )
          throw new Error(
            "The workbook changed while the file was being read. Open the file again.",
          );
        (setAutomationInspection(wt), setActiveDialogMode("automation"));
        return;
      }
      if (/\.(csv|tsv)$/i.test(fe.name)) {
        if (fe.size > workbookLimits.maxFileBytes)
          throw new Error("Files must be 10 MB or smaller.");
        const wt = await importDelimitedFile(fe, {}, Ue);
        if (!Xe()) return;
        if (
          latestEditStateRef.current.book !== activeWorkbook ||
          latestEditStateRef.current.readOnly
        )
          throw new Error(
            "The workbook changed while the file was being read. Open the file again.",
          );
        setDelimitedImportExportSource({
          snapshot: activeWorkbook,
          bytes: new Uint8Array(),
          file: fe,
          initial: wt,
          fileName: fe.name,
        });
        return;
      }
      const Nt = await importWorkbookFile(fe, Ue);
      if (!Xe()) return;
      if (
        latestEditStateRef.current.book !== activeWorkbook ||
        latestEditStateRef.current.readOnly
      )
        throw new Error(
          "The workbook changed while the file was being read. Import again to replace the current revision.",
        );
      if (Nt.report) {
        setXlsxConversionReview({
          snapshot: activeWorkbook,
          report: Nt.report,
          imported: Nt.workbook,
          fileName: fe.name,
        });
        return;
      }
      (commitWorkbookChange(Nt.workbook),
        switchActiveWorksheet(Nt.workbook.sheets.find((wt) => !wt.hidden).id),
        setStatusNotice(
          Nt.warnings.length ? Nt.warnings.join(" ") : "Workbook imported.",
        ));
    } catch (Nt) {
      if (!Xe()) return;
      if (
        !Oe.signal.aborted &&
        (!(Nt instanceof FileProcessingError) ||
          ["FILE_ACTIVE", "FILE_INVALID"].includes(Nt.code)) &&
        /\.xlsx$/i.test(fe.name) &&
        fe.size <= workbookLimits.maxFileBytes &&
        latestEditStateRef.current.book === activeWorkbook
      )
        try {
          const wt = await importAutomationFile(fe, Ue);
          if (
            Xe() &&
            !latestEditStateRef.current.readOnly &&
            wt.items.length &&
            latestEditStateRef.current.book === activeWorkbook
          ) {
            (setAutomationInspection(wt), setActiveDialogMode("automation"));
            return;
          }
        } catch {}
      Xe() && setErrorNotice(toFileProcessingError(Nt).message);
    } finally {
      Xe() &&
        ((activeOperationControllerRef.current = null),
        setIsOperationBusy(false),
        importFileInputRef.current && (importFileInputRef.current.value = ""));
    }
  }
  async function exportWorkbookAs(fe) {
    (setIsOperationBusy(true), setErrorNotice(""));
    try {
      const Oe =
        activeWorkbook.name.replace(/[^a-zA-Z0-9 _-]/g, "").slice(0, 80) ||
        "Workbook";
      if (fe === "xlsx") {
        const Be = await exportWorkbookToXlsx(activeWorkbook);
        (Be.report.warnings.push(
          "Full workbook export includes all filtered and hidden rows. Session filter views are not saved in XLSX. Use CSV/TSV with Visible population for a filtered extract.",
        ),
          setXlsxConversionReview({
            snapshot: activeWorkbook,
            report: { ...Be.report, fileName: `${Oe}-derived.xlsx` },
            blob: Be.blob,
            fileName: `${Oe}-derived.xlsx`,
          }));
      }
      (fe === "csv" &&
        setDelimitedImportExportSource({
          snapshot: activeWorkbook,
          sheetId: activeSheet.id,
          ...(filterResult
            ? { population: filterResult.context.definition }
            : {}),
        }),
        fe === "json" &&
          (setStatusNotice(
            "Native JSON exports the full workbook, including all filtered and hidden rows and any named filter views. Active session masks are not saved; load named criteria, preview and apply after reopening.",
          ),
          saveBlobToFile(
            new Blob([encodeNativeWorkbook(activeWorkbook)], {
              type: "application/json",
            }),
            `${Oe}.json`,
          )));
    } catch (Oe) {
      setErrorNotice(Oe instanceof Error ? Oe.message : "Export failed.");
    } finally {
      setIsOperationBusy(false);
    }
  }
  function runCurrentAuditAnalysis() {
    runGuardedAction(() => {
      if (auditPopulationScope === "visible" && !filterResult)
        throw Error(
          "Review and apply a current population filter before visible-only analysis.",
        );
      const fe = runAuditAnalysis(
        activeWorkbook,
        activeSheet.id,
        selectedAuditKind,
        {
          range:
            auditPopulationScope === "visible" && filterResult
              ? filterResult.context.range
              : auditSourceRange || formatOccupiedRangeAddress(activeSheet),
          ...(auditPopulationScope === "visible"
            ? {
                population: {
                  definition:
                    filterResult?.context.definition ??
                    activeFilterState?.definition,
                  scope: "visible",
                },
              }
            : {}),
          threshold: materialityThreshold,
          sampleSize: auditSampleSize,
          seed: samplingSeed,
          asOf: agingAsOfDate,
        },
      );
      (setAuditAnalysisResult(fe),
        setAgingReportPayload({
          schemaVersion: 1,
          workbookId: activeWorkbook.id,
          workbookName: activeWorkbook.name,
          revision: activeWorkbook.revision,
          sheetId: activeSheet.id,
          sheetName: activeSheet.name,
          generatedAt: new Date().toISOString(),
          actor,
          populationScope: auditPopulationScope,
          ...(filterResult ? { filterContext: filterResult.context } : {}),
          kind: selectedAuditKind,
          parameters: {
            range:
              auditPopulationScope === "visible" && filterResult
                ? filterResult.context.range
                : auditSourceRange || formatOccupiedRangeAddress(activeSheet),
            threshold: materialityThreshold,
            sampleSize: auditSampleSize,
            seed: samplingSeed,
            asOf: agingAsOfDate,
          },
          result: fe,
        }),
        setStatusNotice(
          auditPopulationScope === "visible"
            ? "Analysis completed on visible population rows; citations retain original worksheet addresses."
            : "Analysis completed on the full selected source range, including hidden and filtered rows.",
        ));
    });
  }
  async function askAssistantQuestion(fe, Oe) {
    var Xe;
    if (!fe.trim()) return null;
    if (selectionAreas.length > 1)
      throw (
        setErrorNotice(
          "Select one rectangular range before asking the assistant.",
        ),
        new Error("Select one rectangular range before asking the assistant.")
      );
    (Xe = assistantAbortControllerRef.current) == null || Xe.abort();
    const Be = ++assistantRequestIdRef.current;
    (setIsAssistantBusy(true),
      setAssistantProposal(null),
      setAnswerCitations(null),
      setErrorNotice(""),
      setStatusNotice(""));
    const Ue = new AbortController();
    assistantAbortControllerRef.current = Ue;
    try {
      const gt = buildAssistantSelectionContext(
          activeWorkbook,
          assistantSelection,
        ),
        Nt = await requestAssistantAnswer(adapter, fe, gt, {
          conversation: Oe,
          signal: Ue.signal,
          reliability: activeReliabilityMonitor,
        });
      if (Ue.signal.aborted || Be !== assistantRequestIdRef.current)
        return null;
      const wt = validateProposedCalculation(
        activeWorkbook,
        assistantSelection,
        Nt,
      );
      return (
        setAnswerCitations(
          buildAnswerCitations(activeWorkbook, assistantSelection, wt, gt),
        ),
        setAssistantProposal(wt),
        wt
      );
    } catch (gt) {
      if (
        (Be === assistantRequestIdRef.current &&
          !Ue.signal.aborted &&
          setErrorNotice(
            gt instanceof Error
              ? gt.message
              : "The assistant could not complete the request.",
          ),
        Ue.signal.aborted || Be !== assistantRequestIdRef.current)
      )
        return null;
      throw gt;
    } finally {
      Be === assistantRequestIdRef.current && setIsAssistantBusy(false);
    }
  }
  function replaceActiveWorkbook(fe) {
    runGuardedAction(() => {
      (commitWorkbookChange(fe),
        setUndoStack([]),
        setRedoStack([]),
        switchActiveWorksheet(fe.sheets.find((Oe) => !Oe.hidden).id),
        setIsFilterBarVisible(false),
        setIsFindBarVisible(false),
        setSidePanelKind(null));
    });
  }
  function toggleFreezeTopRow() {
    runGuardedAction(() =>
      commitWorkbookEdit(
        setFreezePanes(
          activeWorkbook,
          activeSheet.id,
          hasFrozenRows ? 0 : 1,
          activeSheet.frozenCols ?? 0,
        ),
        hasFrozenRows ? "Unfreeze rows" : "Freeze top row",
      ),
    );
  }
  function selectionToTsv(fe = false) {
    (requireSingleSelectionRange(), getSelectedCellAddresses());
    const Oe = {
        row: selectionRangeBounds.firstRow,
        col: selectionRangeBounds.firstCol,
      },
      Be = {
        row: selectionRangeBounds.lastRow,
        col: selectionRangeBounds.lastCol,
      },
      Ue = [];
    for (let Nt = Oe.row; Nt <= Be.row; Nt++) {
      const wt = [];
      for (let jt = Oe.col; jt <= Be.col; jt++) {
        const rr = `${columnIndexToLetters(jt)}${Nt + 1}`,
          Mr = workbookCalculator(activeSheet.id, rr);
        wt.push(
          typeof Mr == "string" && /^[=+@-]/.test(Mr)
            ? `'${Mr}`
            : String(Mr ?? ""),
        );
      }
      Ue.push(wt);
    }
    const Xe = x7.unparse(Ue, {
        delimiter: "	",
        newline: `
`,
      }),
      gt = fe
        ? createCutSourceSnapshot(
            activeWorkbook,
            activeSheet.id,
            activeSelectionRange,
          )
        : null;
    return (
      setClipboardSnapshotState(gt ? { snapshot: gt, text: Xe } : null),
      (internalClipboardRef.current = {
        text: Xe,
        snapshot: buildClipboardSnapshot(
          activeWorkbook,
          activeSheet.id,
          activeSelectionRange,
        ),
      }),
      Xe
    );
  }
  async function copyOrCutSelection(fe) {
    var Oe;
    try {
      if (fe && isEditingLocked) throw new Error("This workbook is read-only.");
      const Be = activeWorkbook,
        Ue = selectionToTsv(fe);
      if ((await navigator.clipboard.writeText(Ue), fe)) {
        if (
          latestEditStateRef.current.readOnly ||
          latestEditStateRef.current.book !== Be ||
          latestEditStateRef.current.sheetId !== activeSheet.id ||
          latestEditStateRef.current.selectionKey !== selectionAreasKey
        )
          throw new Error("Workbook changed while copying. Copy again.");
        setStatusNotice(
          "Cut is ready. Select a destination and paste to preview the move.",
        );
      }
      fe || setStatusNotice("Selection copied.");
    } catch (Be) {
      (fe && setClipboardSnapshotState(null),
        setErrorNotice(
          Be instanceof Error && Be.name !== "NotAllowedError"
            ? Be.message
            : "Clipboard access is unavailable. Use Ctrl/⌘ C or X with the grid focused.",
        ),
        (Oe = gridContainerRef.current) == null || Oe.focus());
    }
  }
  function pasteClipboardText(fe) {
    runGuardedAction(() => {
      if (
        (requireSingleSelectionRange(), fe.length > workbookLimits.maxFileBytes)
      )
        throw new Error("Clipboard text is too large.");
      const Oe =
        internalClipboardRef.current?.text === fe
          ? internalClipboardRef.current
          : null;
      if (Oe && clipboardSnapshotState?.text === fe) {
        setMoveCellsSource(clipboardSnapshotState.snapshot);
        return;
      }
      if (Oe) {
        const wt = buildPastePreview(
          activeWorkbook,
          activeSheet.id,
          activeSelectionRange,
          Oe.snapshot,
          {
            mode: "all",
            deferValidationWarnings: true,
            readOnly: latestEditStateRef.current.readOnly,
          },
        );
        (applyCellPatches(wt.patches, "Paste cells"),
          selectSingleRange(wt.targetRange));
        return;
      }
      const Be = x7.parse(fe.replace(/\r?\n$/, ""), {
        delimiter: "	",
        skipEmptyLines: false,
      });
      if (Be.errors.length)
        throw new Error("Clipboard text contains malformed quoted cells.");
      const Ue = Be.data,
        Xe = _t(activeCellAddress),
        gt = [];
      if (!Ue.length || Ue.some((wt) => wt.length !== Ue[0].length))
        throw new Error("Paste requires a rectangular range.");
      if (Ue.length * Ue[0].length > workbookLimits.maxPatches)
        throw new Error(
          `Paste at most ${workbookLimits.maxPatches} cells at a time.`,
        );
      (Ue.forEach((wt, jt) =>
        wt.forEach((rr, Mr) => {
          if (
            Xe.row + jt >= activeSheet.rowCount ||
            Xe.col + Mr >= activeSheet.colCount
          )
            throw new Error(
              "Paste exceeds the Excel worksheet boundary (XFD1048576).",
            );
          const Br = `${columnIndexToLetters(Xe.col + Mr)}${Xe.row + jt + 1}`,
            xr = parseDateCellEntry(
              rr,
              activeSheet.cells[Br],
              latestEditStateRef.current.book.datePolicy?.dateSystem,
            );
          gt.push({ sheetId: activeSheet.id, address: Br, cell: xr });
        }),
      ),
        applyCellPatches(gt, "Paste cells"),
        selectSingleRange(
          wy(
            activeCellAddress,
            `${columnIndexToLetters(Xe.col + Ue[0].length - 1)}${Xe.row + Ue.length}`,
          ),
        ));
    });
  }
  async function pasteFromClipboard() {
    var fe;
    try {
      const Oe = activeWorkbook,
        Be = await navigator.clipboard.readText();
      if (
        latestEditStateRef.current.book !== Oe ||
        latestEditStateRef.current.sheetId !== activeSheet.id ||
        latestEditStateRef.current.selectionKey !== selectionAreasKey
      )
        throw new Error("Workbook changed while pasting. Try again.");
      pasteClipboardText(Be);
    } catch {
      (setErrorNotice(
        "Clipboard access is unavailable. Use Ctrl/⌘ V with the grid focused.",
      ),
        (fe = gridContainerRef.current) == null || fe.focus());
    }
  }
  function findNextTextMatch() {
    runGuardedAction(() => {
      if (!findQuery) return;
      const fe = Object.keys(activeSheet.cells)
        .filter((Ue) => {
          const Xe = _t(Ue);
          return (
            !rowVisibilityIndex.isHidden(Xe.row) &&
            !columnVisibilityIndex.isHidden(Xe.col) &&
            String(activeSheet.cells[Ue]?.raw ?? "")
              .toLowerCase()
              .includes(findQuery.toLowerCase())
          );
        })
        .sort((Ue, Xe) => {
          const gt = _t(Ue),
            Nt = _t(Xe);
          return gt.row - Nt.row || gt.col - Nt.col;
        });
      if (!fe.length) {
        setStatusNotice(`We couldn't find “${findQuery}”.`);
        return;
      }
      const Oe = _t(activeCellAddress),
        Be =
          fe.find((Ue) => {
            const Xe = _t(Ue);
            return Xe.row > Oe.row || (Xe.row === Oe.row && Xe.col > Oe.col);
          }) ?? fe[0];
      (selectCellAddress(Be),
        setStatusNotice(
          `${fe.length} matching cell${fe.length === 1 ? "" : "s"}.`,
        ));
    });
  }
  function replaceTextInSelection() {
    runGuardedAction(() => {
      if (!findQuery) throw new Error("Enter text to find.");
      const fe = new RegExp(
          findQuery.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
          "gi",
        ),
        Oe = getSelectedCellAddresses()
          .filter((Be) => {
            return (
              typeof activeSheet.cells[Be]?.raw == "string" &&
              String(activeSheet.cells[Be].raw)
                .toLowerCase()
                .includes(findQuery.toLowerCase())
            );
          })
          .map((Be) => ({
            sheetId: activeSheet.id,
            address: Be,
            cell: {
              ...activeSheet.cells[Be],
              raw: String(activeSheet.cells[Be].raw).replace(
                fe,
                () => replaceWithText,
              ),
            },
          }));
      (Oe.length && applyCellPatches(Oe, "Replace text in selection"),
        setStatusNotice(`Replaced text in ${Oe.length} cells.`));
    });
  }
  function stepVisibleRowIndex(fe, Oe) {
    if (!sortedRowIndexes) {
      const Xe =
        rowVisibilityIndex.viewIndex(fe) ?? rowVisibilityIndex.countBefore(fe);
      return rowVisibilityIndex.logicalIndex(
        Math.max(0, Math.min(rowVisibilityIndex.count - 1, Xe + Oe)),
      );
    }
    const Be = [...frozenRowIndexes, ...sortedRowIndexes],
      Ue = Be.indexOf(fe);
    return Be[Math.max(0, Math.min(Be.length - 1, Ue + Oe))] ?? fe;
  }
  function stepVisibleColumnIndex(fe, Oe) {
    const Be =
      columnVisibilityIndex.viewIndex(fe) ??
      columnVisibilityIndex.countBefore(fe);
    return columnVisibilityIndex.logicalIndex(
      Math.max(0, Math.min(columnVisibilityIndex.count - 1, Be + Oe)),
    );
  }
  function jumpToDataEdge(fe, Oe, Be, Ue) {
    const Xe =
        Be && sortedRowIndexes
          ? [...frozenRowIndexes, ...sortedRowIndexes]
          : null,
      gt = Be ? rowVisibilityIndex : columnVisibilityIndex,
      Nt = Xe ? new Map(Xe.map((xr, Sr) => [xr, Sr])) : null,
      wt = (xr) => (Nt ? (Nt.get(xr) ?? null) : gt.viewIndex(xr)),
      jt = wt(Be ? fe : Oe);
    if (jt === null) return { row: fe, col: Oe };
    const rr = getAllCellEntries().flatMap((xr) => {
        if (
          xr.cell.raw === null ||
          xr.cell.raw === "" ||
          (Be ? xr.col !== Oe : xr.row !== fe)
        )
          return [];
        const Sr = wt(Be ? xr.row : xr.col);
        return Sr === null ? [] : [Sr];
      }),
      Mr = resolveCellNavigation(jt, Xe?.length ?? gt.count, rr, Ue),
      Br = Xe ? Xe[Mr] : gt.logicalIndex(Mr);
    return Be ? { row: Br, col: Oe } : { row: fe, col: Br };
  }
  function moveSelectionAfterCommit(fe, Oe) {
    const Be = pendingCrossSheetTargetRef.current;
    if (((pendingCrossSheetTargetRef.current = null), Be)) {
      const Nt = activeWorkbook.sheets.find((xr) => xr.id === Be.sheetId),
        wt = _t(Be.address),
        jt = fe === "Enter" ? "row" : "column",
        Mr =
          createAxisVisibilityIndex(
            jt === "row" ? Nt.rowCount : Nt.colCount,
            cn(Nt, jt),
          ).next(jt === "row" ? wt.row : wt.col, Oe ? -1 : 1) ??
          (jt === "row" ? wt.row : wt.col),
        Br = `${columnIndexToLetters(jt === "column" ? Mr : wt.col)}${(jt === "row" ? Mr : wt.row) + 1}`;
      (setActiveSheetId(Be.sheetId),
        setActiveCellAddress(Br),
        setSelectionAnchorAddress(Br),
        selectSingleRange(Br),
        setPendingScrollAddress(Br),
        focusGridSurface());
      return;
    }
    const Ue = _t(activeCellAddress),
      Xe = fe === "Enter" ? stepVisibleRowIndex(Ue.row, Oe ? -1 : 1) : Ue.row,
      gt = fe === "Tab" ? stepVisibleColumnIndex(Ue.col, Oe ? -1 : 1) : Ue.col;
    (selectCellAddress(`${columnIndexToLetters(gt)}${Xe + 1}`),
      focusGridSurface());
  }
  function selectAllCells() {
    if (!commitCellEdit()) return;
    selectSingleRange(
      `A1:${columnIndexToLetters(activeSheet.colCount - 1)}${activeSheet.rowCount}`,
    );
    const fe = `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${rowVisibilityIndex.logicalIndex(0) + 1}`;
    (setActiveCellAddress(fe),
      setSelectionAnchorAddress(fe),
      (activeCellAddressRef.current = fe),
      focusGridSurface());
  }
  function selectRowOrColumn(fe, Oe, Be) {
    if (!commitCellEdit()) return;
    const Ue = _t(selectionAnchorAddress)[fe === "column" ? "col" : "row"],
      Xe = Be ? Math.min(Ue, Oe) : Oe,
      gt = Be ? Math.max(Ue, Oe) : Oe,
      Nt =
        fe === "column"
          ? `${columnIndexToLetters(Oe)}${rowVisibilityIndex.logicalIndex(0) + 1}`
          : `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${Oe + 1}`;
    (selectSingleRange(
      fe === "column"
        ? `${columnIndexToLetters(Xe)}1:${columnIndexToLetters(gt)}${activeSheet.rowCount}`
        : `A${Xe + 1}:${columnIndexToLetters(activeSheet.colCount - 1)}${gt + 1}`,
    ),
      (activeCellAddressRef.current = Nt),
      setActiveCellAddress(Be ? selectionAnchorAddress : Nt),
      Be || setSelectionAnchorAddress(Nt),
      focusGridSurface());
  }
  function resolveSelectedAxisIndexes(fe, Oe) {
    requireSingleSelectionRange();
    const { firstRow, lastRow, firstCol, lastCol } = selectionRangeBounds,
      Nt =
        fe === "column"
          ? firstRow === 0 && lastRow === activeSheet.rowCount - 1
          : firstCol === 0 && lastCol === activeSheet.colCount - 1,
      wt = fe === "column" ? firstCol : firstRow,
      jt = fe === "column" ? lastCol : lastRow;
    if (Oe !== undefined && (!Nt || Oe < wt || Oe > jt)) return [Oe];
    const rr = new Set();
    for (const Br of getAllCellEntries()) {
      const xr = fe === "column" ? Br.col : Br.row;
      xr >= wt && xr <= jt && rr.add(xr);
    }
    const Mr =
      fe === "column" ? activeSheet.columnWidths : activeSheet.rowHeights;
    for (const Br of Object.keys(Mr ?? {})) {
      const xr = Number(Br);
      xr >= wt && xr <= jt && rr.add(xr);
    }
    if (jt - wt < 500) for (let Br = wt; Br <= jt; Br++) rr.add(Br);
    return [...rr].sort((Br, xr) => Br - xr);
  }
  function applyAxisSizes(fe, Oe, Be) {
    const Ue = latestEditStateRef.current.book,
      Xe = Ue.sheets.find((rr) => rr.id === activeSheet.id),
      gt = fe === "column" ? "columnWidths" : "rowHeights";
    if (
      !Object.entries(Oe).some(([rr, Mr]) => {
        return (
          (Xe[gt]?.[rr] ??
            (fe === "column"
              ? defaultColumnWidthOf(Xe)
              : defaultRowHeightOf(Xe))) !== Mr
        );
      })
    )
      return;
    const wt = _t(activeCellAddress),
      jt =
        (shownColumns.includes(wt.col) ||
          frozenColumnIndexes.includes(wt.col)) &&
        (shownRows.includes(wt.row) || frozenRowIndexes.includes(wt.row));
    (commitWorkbookChange({
      ...Ue,
      revision: Ue.revision + 1,
      sheets: Ue.sheets.map((rr) =>
        rr.id === activeSheet.id ? { ...rr, [gt]: { ...rr[gt], ...Oe } } : rr,
      ),
      events: [
        ...Ue.events,
        {
          id: createUuid(),
          timestamp: new Date().toISOString(),
          actor,
          action: Be,
          changes: [],
        },
      ],
    }),
      jt && requestCellScroll(activeCellAddress));
  }
  function autoFitSelectedAxis(fe, Oe) {
    runGuardedAction(() => {
      if (
        isEditingLocked ||
        latestEditStateRef.current.readOnly ||
        !commitCellEdit()
      )
        return;
      const Be = resolveSelectedAxisIndexes(fe, Oe),
        Ue = latestEditStateRef.current.book.sheets.find(
          (jt) => jt.id === activeSheet.id,
        ),
        Xe = calculateWorkbookWithSettings(latestEditStateRef.current.book),
        gt =
          fe === "column"
            ? autoFitColumnWidths(
                Ue,
                Be,
                Xe,
                textMeasurer,
                isShowingFormulas,
                latestEditStateRef.current.book.datePolicy?.dateSystem,
              )
            : autoFitRowHeights(
                Ue,
                Be,
                Xe,
                textMeasurer,
                isShowingFormulas,
                latestEditStateRef.current.book.datePolicy?.dateSystem,
              );
      (applyAxisSizes(
        fe,
        gt,
        `AutoFit ${fe === "column" ? "column widths" : "row heights"}`,
      ),
        setAxisResizeState(null),
        focusGridSurface());
    });
  }
  function beginAxisResize(fe, Oe, Be) {
    if (
      (Be.preventDefault(),
      Be.stopPropagation(),
      Be.button !== 0 || isEditingLocked || !commitCellEdit())
    )
      return;
    const Ue =
      fe === "column"
        ? columnAxisLayout.size(columnVisibilityIndex.viewIndex(Oe) ?? 0)
        : (effectiveRowHeights[Oe] ?? defaultRowHeight);
    setAxisResizeState({
      axis: fe,
      index: Oe,
      start: fe === "column" ? Be.clientX : Be.clientY,
      initial: Ue,
      size: Ue,
    });
  }
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
  function beginFillDrag(fe) {
    (fe.preventDefault(),
      fe.stopPropagation(),
      !(fe.button !== 0 || isEditingLocked || !commitCellEdit()) &&
        runGuardedAction(() => {
          getSelectedCellAddresses();
          const Oe = {
            book: latestEditStateRef.current.book,
            sheetId: activeSheet.id,
            sourceRange: activeSelectionRange,
            targetRange: activeSelectionRange,
          };
          ((dragAnchorRef.current = null),
            (fillDragStateRef.current = Oe),
            setFillDragState(Oe));
        }));
  }
  function extendFillDragTo(fe) {
    const Oe = fillDragStateRef.current;
    if (!Oe) return;
    const Be = parseRangeAddress(Oe.sourceRange),
      Ue = _t(fe),
      Xe = Math.max(Be.firstRow - Ue.row, Ue.row - Be.lastRow, 0),
      gt = Math.max(Be.firstCol - Ue.col, Ue.col - Be.lastCol, 0),
      Nt = Xe >= gt ? Math.min(Be.firstRow, Ue.row) : Be.firstRow,
      wt = Xe >= gt ? Math.max(Be.lastRow, Ue.row) : Be.lastRow,
      jt = gt > Xe ? Math.min(Be.firstCol, Ue.col) : Be.firstCol,
      rr = gt > Xe ? Math.max(Be.lastCol, Ue.col) : Be.lastCol,
      Mr = wy(
        `${columnIndexToLetters(jt)}${Nt + 1}`,
        `${columnIndexToLetters(rr)}${wt + 1}`,
      ),
      Br = { ...Oe, targetRange: Mr };
    ((fillDragStateRef.current = Br), setFillDragState(Br));
  }
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
  function handleGridKeyDown(fe) {
    var wt;
    if (
      fe.target !== fe.currentTarget ||
      fe.defaultPrevented ||
      fe.nativeEvent.isComposing ||
      (fe.altKey &&
        !((fe.ctrlKey || fe.metaKey) && fe.key.toLowerCase() === "v")) ||
      fe.getModifierState("AltGraph")
    )
      return;
    const Oe = fe.ctrlKey || fe.metaKey,
      Be = fe.key.toLowerCase();
    if (
      (Oe &&
        ["tab", "enter", "f2", "f5", "f6", "delete", "backspace"].includes(
          Be,
        )) ||
      (Oe &&
        fe.shiftKey &&
        ![
          "z",
          "f",
          "home",
          "end",
          " ",
          "arrowleft",
          "arrowright",
          "arrowup",
          "arrowdown",
        ].includes(Be))
    )
      return;
    if (
      isOperationBusy &&
      ((Oe && ["z", "y", "b", "i", "u", "d", "r", "v"].includes(Be)) ||
        (!Oe &&
          (fe.key.length === 1 ||
            ["F2", "Delete", "Backspace"].includes(fe.key))))
    ) {
      (fe.preventDefault(),
        setStatusNotice("Wait for the current workbook operation to finish."));
      return;
    }
    if (
      (fe.key === "ContextMenu" && !fe.shiftKey && !Oe) ||
      (fe.key === "F10" && fe.shiftKey && !fe.ctrlKey && !fe.metaKey)
    ) {
      fe.preventDefault();
      const rr = (
        fe.currentTarget.querySelector(
          `[data-address="${activeCellAddress}"]`,
        ) ?? fe.currentTarget
      ).getBoundingClientRect();
      setContextMenuState({
        x: rr.left + Math.min(40, rr.width),
        y: rr.top + Math.min(24, rr.height),
        kind:
          parseRangeAddress(activeSelectionRange).lastRow -
            parseRangeAddress(activeSelectionRange).firstRow +
            1 ===
          activeSheet.rowCount
            ? "column"
            : parseRangeAddress(activeSelectionRange).lastCol -
                  parseRangeAddress(activeSelectionRange).firstCol +
                  1 ===
                activeSheet.colCount
              ? "row"
              : "cell",
      });
      return;
    }
    if (fe.key === "Escape" && clipboardSnapshotState) {
      (fe.preventDefault(),
        clearCutSession(),
        setStatusNotice("Cut canceled; source cells remain unchanged."));
      return;
    }
    if (fe.key === "Escape" && completedSelectionAreas.length) {
      (fe.preventDefault(),
        selectSingleRange(activeCellAddress),
        setSelectionAnchorAddress(activeCellAddress),
        (activeCellAddressRef.current = activeCellAddress));
      return;
    }
    if (fe.key === "Escape" && !fe.ctrlKey && !fe.metaKey && !fe.shiftKey) {
      (fe.preventDefault(),
        (wt = appRootRef.current?.querySelector(
          '[role="tab"][aria-selected="true"]',
        )) == null || wt.focus());
      return;
    }
    if (
      (fe.key === "F5" && !fe.shiftKey && !Oe) ||
      ((fe.ctrlKey || fe.metaKey) && fe.key.toLowerCase() === "g")
    ) {
      (fe.preventDefault(), runEditorCommand("goToSpecial"));
      return;
    }
    if (
      (fe.ctrlKey || fe.metaKey) &&
      fe.shiftKey &&
      fe.key.toLowerCase() === "f"
    ) {
      (fe.preventDefault(), runEditorCommand("search"));
      return;
    }
    if (completedSelectionAreas.length && ["Tab", "Enter"].includes(fe.key)) {
      (fe.preventDefault(),
        runGuardedAction(() => {
          const jt = enumerateSelectionCells(selectionAreas).filter((xr) => {
            const Sr = _t(xr);
            return (
              !rowVisibilityIndex.isHidden(Sr.row) &&
              !columnVisibilityIndex.isHidden(Sr.col)
            );
          });
          if (!jt.length) return;
          const rr = jt.indexOf(activeCellAddress),
            Mr = fe.shiftKey ? -1 : 1,
            Br = jt[(rr + Mr + jt.length) % jt.length];
          commitCellEdit() &&
            (setActiveCellAddress(Br),
            setSelectionAnchorAddress(Br),
            (activeCellAddressRef.current = Br),
            requestCellScroll(Br));
        }));
      return;
    }
    const Ue = _t(
      fe.shiftKey && fe.key.startsWith("Arrow")
        ? activeCellAddressRef.current
        : activeCellAddress,
    );
    let Xe = Ue.row,
      gt = Ue.col;
    if ((fe.ctrlKey || fe.metaKey) && fe.key.toLowerCase() === "z") {
      (fe.preventDefault(), undoOrRedo(fe.shiftKey));
      return;
    }
    if (fe.ctrlKey || fe.metaKey) {
      const jt = fe.key.toLowerCase();
      if (jt === "1") {
        (fe.preventDefault(), runEditorCommand("numberFormat"));
        return;
      }
      if (jt === "v" && fe.altKey) {
        (fe.preventDefault(), runEditorCommand("pasteSpecial"));
        return;
      }
      if (jt === " ") {
        (fe.preventDefault(), selectRowOrColumn("column", Ue.col, fe.shiftKey));
        return;
      }
      if (jt === "pageup" || jt === "pagedown") {
        fe.preventDefault();
        const rr = activeWorkbook.sheets.filter((xr) => !xr.hidden),
          Mr = rr.findIndex((xr) => xr.id === activeSheet.id),
          Br = rr[Mr + (jt === "pageup" ? -1 : 1)];
        Br && commitCellEdit() && switchActiveWorksheet(Br.id);
        return;
      }
      if (jt.startsWith("arrow")) {
        fe.preventDefault();
        const rr = jumpToDataEdge(
          Ue.row,
          Ue.col,
          jt === "arrowup" || jt === "arrowdown",
          jt === "arrowup" || jt === "arrowleft" ? -1 : 1,
        );
        selectCellAddress(
          `${columnIndexToLetters(rr.col)}${rr.row + 1}`,
          fe.shiftKey,
        );
        return;
      }
      if (jt === "a") {
        (fe.preventDefault(), selectAllCells());
        return;
      }
      if (["b", "i", "u"].includes(jt)) {
        if ((fe.preventDefault(), !isEditingLocked)) {
          const rr = jt === "b" ? "bold" : jt === "i" ? "italic" : "underline";
          applyCellFormatPatch({
            style: { [rr]: !resolveEffectiveCellStyle(activeCell?.style)[rr] },
          });
        }
        return;
      }
      if (jt === "d") {
        (fe.preventDefault(),
          isEditingLocked ||
            runGuardedAction(() => {
              (requireSingleSelectionRange(),
                commitWorkbookEdit(
                  applyCellPreview(
                    activeWorkbook,
                    buildFillPreview(
                      activeWorkbook,
                      activeSheet.id,
                      activeSelectionRange,
                      "down",
                    ),
                  ),
                  "Fill down",
                ));
            }));
        return;
      }
      if (jt === "r") {
        (fe.preventDefault(),
          isEditingLocked ||
            runGuardedAction(() => {
              (requireSingleSelectionRange(),
                commitWorkbookEdit(
                  applyCellPreview(
                    activeWorkbook,
                    buildFillPreview(
                      activeWorkbook,
                      activeSheet.id,
                      activeSelectionRange,
                      "right",
                    ),
                  ),
                  "Fill right",
                ));
            }));
        return;
      }
      if (jt === "y") {
        (fe.preventDefault(), undoOrRedo(true));
        return;
      }
      if (jt === "f" || jt === "h") {
        (fe.preventDefault(),
          setIsFindBarVisible(true),
          requestAnimationFrame(() => {
            return findInputRef.current?.focus();
          }));
        return;
      }
      if (jt === "home") {
        (fe.preventDefault(),
          selectCellAddress(
            `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${rowVisibilityIndex.logicalIndex(0) + 1}`,
            fe.shiftKey,
          ));
        return;
      }
      if (jt === "end") {
        fe.preventDefault();
        const rr = findOccupiedRangeBounds(activeSheet);
        selectCellAddress(
          `${columnIndexToLetters(columnVisibilityIndex.next(rr?.lastCol ?? 0, -1, true) ?? columnVisibilityIndex.logicalIndex(0))}${(rowVisibilityIndex.next(rr?.lastRow ?? 0, -1, true) ?? rowVisibilityIndex.logicalIndex(0)) + 1}`,
          fe.shiftKey,
        );
        return;
      }
      return;
    }
    if (fe.shiftKey && fe.key === " ") {
      (fe.preventDefault(), selectRowOrColumn("row", Ue.row, false));
      return;
    }
    if (["Home", "PageUp", "PageDown"].includes(fe.key)) {
      (fe.preventDefault(),
        fe.key === "Home"
          ? (gt = columnVisibilityIndex.logicalIndex(0))
          : (Xe = stepVisibleRowIndex(
              Xe,
              Math.max(
                1,
                Math.floor(rowViewport / zoomScale / defaultRowHeight) - 1,
              ) * (fe.key === "PageUp" ? -1 : 1),
            )),
        selectCellAddress(`${columnIndexToLetters(gt)}${Xe + 1}`, fe.shiftKey));
      return;
    }
    if (
      [
        "ArrowDown",
        "ArrowUp",
        "ArrowLeft",
        "ArrowRight",
        "Tab",
        "Enter",
      ].includes(fe.key)
    ) {
      if (
        (fe.preventDefault(),
        fe.key === "ArrowDown" && (Xe = stepVisibleRowIndex(Xe, 1)),
        fe.key === "Enter" &&
          (Xe = stepVisibleRowIndex(Xe, fe.shiftKey ? -1 : 1)),
        fe.key === "ArrowUp" && (Xe = stepVisibleRowIndex(Xe, -1)),
        fe.key === "ArrowLeft" && (gt = stepVisibleColumnIndex(gt, -1)),
        (fe.key === "ArrowRight" || fe.key === "Tab") &&
          (gt = stepVisibleColumnIndex(
            gt,
            fe.shiftKey && fe.key === "Tab" ? -1 : 1,
          )),
        !fe.shiftKey && !filterValueSearch && !activeColumnSort)
      ) {
        const rr = mergeRanges.find((Mr) => Tp(Mr, Ue.row, Ue.col));
        rr &&
          (gt > Ue.col &&
            (gt = columnVisibilityIndex.next(rr.lastCol, 1) ?? Ue.col),
          gt < Ue.col &&
            (gt = columnVisibilityIndex.next(rr.firstCol, -1) ?? Ue.col),
          Xe > Ue.row &&
            (Xe = rowVisibilityIndex.next(rr.lastRow, 1) ?? Ue.row),
          Xe < Ue.row &&
            (Xe = rowVisibilityIndex.next(rr.firstRow, -1) ?? Ue.row));
      }
      const jt = `${columnIndexToLetters(Math.max(0, Math.min(activeSheet.colCount - 1, gt)))}${Math.max(1, Math.min(activeSheet.rowCount, Xe + 1))}`;
      selectCellAddress(jt, fe.shiftKey && !["Tab", "Enter"].includes(fe.key));
      return;
    }
    (!isEditingLocked &&
      (fe.key === "Delete" || fe.key === "Backspace") &&
      (fe.preventDefault(),
      runGuardedAction(() =>
        applyCellPatches(
          getSelectedCellAddresses().map((jt) => ({
            sheetId: activeSheet.id,
            address: jt,
            cell: { ...activeSheet.cells[jt], raw: null },
          })),
          "Clear selection",
        ),
      )),
      !isEditingLocked &&
        (fe.key === "F2" ||
          (fe.key.length === 1 && !fe.ctrlKey && !fe.metaKey && !fe.altKey)) &&
        (fe.preventDefault(),
        (formulaPointerStateRef.current.mode = fe.key !== "F2"),
        setIsFormulaPointMode(formulaPointerStateRef.current.mode),
        beginCellEdit(
          fe.key === "F2" ? getEditableCellText(activeCell) : fe.key,
          true,
        )));
  }
  function commitDialogInput() {
    promptDialogState &&
      runGuardedAction(() => {
        if (promptDialogState.kind === "note")
          applyCellFormatPatch({ note: promptDialogState.value });
        else if (promptDialogState.kind === "rename")
          replaceWorkbookSnapshot(
            renameWorkbook(activeWorkbook, promptDialogState.value),
            "Rename workbook",
          );
        else {
          const fe = addWorksheet(activeWorkbook, promptDialogState.value),
            Oe = fe.sheets.at(-1);
          (replaceWorkbookSnapshot(fe, `Add sheet: ${Oe.name}`),
            switchActiveWorksheet(Oe.id));
        }
        setPromptDialogState(null);
      });
  }
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
  function activateGridPane(fe) {
    var Oe;
    return splitPaneSizes.active.current === fe
      ? true
      : !(
            (Oe = activeCellEditRef.current) != null && Oe.value.startsWith("=")
          ) && !commitCellEdit()
        ? false
        : (splitPaneSizes.activate(fe),
          (gridContainerRef.current = splitPaneSizes.refs.current[fe] ?? null),
          true);
  }
  function renderGridPane(fe) {
    const Oe = _t(activeCellAddress),
      Be =
        filterValueSearch || activeColumnSort
          ? undefined
          : mergeRanges.find((Sr) => Tp(Sr, Oe.row, Oe.col)),
      Ue = Be
        ? `${columnIndexToLetters(Be.firstCol)}${Be.firstRow + 1}`
        : activeCellAddress,
      Xe = splitPaneSizes.views[fe],
      gt = computeGridViewport(Xe),
      {
        rowWindow,
        columnWindow,
        shownRows,
        shownColumns,
        leftGap,
        rightGap,
        visibleColumnCount,
      } = gt;
    return i.jsxs("div", {
      className: "lg-grid-stage",
      onPointerDownCapture: (Sr) => {
        handleFormulaPointerDown(Sr) ||
          activateGridPane(fe) ||
          (Sr.preventDefault(), Sr.stopPropagation());
      },
      onFocusCapture: () => activateGridPane(fe),
      children: [
        i.jsx("div", {
          className: "lg-grid-scroll",
          id: `${editorIdPrefix}-${fe}-grid`,
          ref: (Sr) => {
            (splitPaneSizes.bindings[fe](Sr),
              splitPaneSizes.active.current === fe &&
                (gridContainerRef.current = Sr));
          },
          tabIndex: 0,
          role: "grid",
          "aria-rowcount": activeSheet.rowCount + 1,
          "aria-colcount": activeSheet.colCount + 1,
          "aria-readonly": isEditingLocked,
          "aria-multiselectable": "true",
          "aria-describedby": `${editorIdPrefix}-grid-help`,
          "aria-activedescendant":
            [...frozenRowIndexes, ...shownRows].includes(
              _t(activeCellAddress).row,
            ) &&
            [...frozenColumnIndexes, ...shownColumns].includes(
              _t(activeCellAddress).col,
            )
              ? `${editorIdPrefix}-${fe}-${Ue}`
              : undefined,
          "aria-label": sheetSplitPanes
            ? `Worksheet pane ${fe.replace("-", " ")}`
            : text("Editable worksheet grid"),
          onScroll: (Sr) => {
            applyPaneScrollOffsets(
              fe,
              Sr.currentTarget.scrollTop,
              Sr.currentTarget.scrollLeft,
            );
          },
          onKeyDown: (Sr) => {
            var Wi;
            if (
              sheetSplitPanes &&
              Sr.key === "F6" &&
              Sr.target === Sr.currentTarget &&
              !Sr.defaultPrevented &&
              !Sr.ctrlKey &&
              !Sr.metaKey &&
              !Sr.altKey
            ) {
              Sr.preventDefault();
              const Di = Pv(sheetSplitPanes),
                $o = Di.indexOf(fe),
                Ds = Di[($o + (Sr.shiftKey ? Di.length - 1 : 1)) % Di.length];
              activateGridPane(Ds) &&
                ((Wi = splitPaneSizes.refs.current[Ds]) == null ||
                  Wi.focus({ preventScroll: true }));
            } else handleGridKeyDown(Sr);
          },
          onContextMenu: (Sr) => {
            const Wi = Sr.target;
            if (Wi.closest("input,textarea,button,a,[contenteditable=true]"))
              return;
            const Di = Wi.closest("[data-address]"),
              $o = Wi.closest("[data-column]"),
              Ds = Wi.closest("[data-row-header]");
            if (!(!Di && !$o && !Ds) && commitCellEdit()) {
              if ((Sr.preventDefault(), $o)) {
                const qo = Number($o.dataset.column);
                selectionAreaRanges.some(
                  (yo) =>
                    yo.firstRow === 0 &&
                    yo.lastRow === activeSheet.rowCount - 1 &&
                    qo >= yo.firstCol &&
                    qo <= yo.lastCol,
                ) || selectRowOrColumn("column", qo, false);
              } else if (Ds) {
                const qo = Number(Ds.dataset.rowHeader);
                selectionAreaRanges.some(
                  (yo) =>
                    yo.firstCol === 0 &&
                    yo.lastCol === activeSheet.colCount - 1 &&
                    qo >= yo.firstRow &&
                    qo <= yo.lastRow,
                ) || selectRowOrColumn("row", qo, false);
              } else if (Di) {
                const qo = _t(Di.dataset.address);
                selectionAreaRanges.some((yo) => Tp(yo, qo.row, qo.col)) ||
                  selectCellAddress(Di.dataset.address);
              }
              setContextMenuState({
                x: Sr.clientX,
                y: Sr.clientY,
                kind: $o ? "column" : Ds ? "row" : "cell",
              });
            }
          },
          onCopy: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(),
              runGuardedAction(() =>
                Sr.clipboardData.setData("text/plain", selectionToTsv()),
              ));
          },
          onCut: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              isEditingLocked ||
              isOperationBusy ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(),
              runGuardedAction(() => {
                (Sr.clipboardData.setData("text/plain", selectionToTsv(true)),
                  setStatusNotice(
                    "Cut is ready. Select a destination and paste to preview the move.",
                  ));
              }));
          },
          onPaste: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              isEditingLocked ||
              isOperationBusy ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(),
              pasteClipboardText(Sr.clipboardData.getData("text/plain")));
          },
          onPointerUp: () => {
            dragAnchorRef.current = null;
          },
          onPointerLeave: (Sr) => {
            Sr.buttons === 0 && (dragAnchorRef.current = null);
          },
          children: i.jsxs("table", {
            className: `lg-grid ${isGridlinesVisible ? "" : "no-gridlines"} ${isHeadingsVisible ? "" : "no-headings"}`,
            style: {
              zoom: zoomScale,
              width: columnWindow.total + frozenColumnsWidth + rowHeaderWidth,
              minWidth: 0,
            },
            role: "presentation",
            children: [
              i.jsxs("colgroup", {
                children: [
                  isHeadingsVisible &&
                    i.jsx("col", { style: { width: rowHeaderWidth } }),
                  frozenColumnIndexes.map((Sr) =>
                    i.jsx(
                      "col",
                      {
                        style: {
                          width: columnAxisLayout.size(
                            columnVisibilityIndex.viewIndex(Sr),
                          ),
                        },
                      },
                      Sr,
                    ),
                  ),
                  leftGap > 0 && i.jsx("col", { style: { width: leftGap } }),
                  shownColumns.map((Sr) =>
                    i.jsx(
                      "col",
                      {
                        style: {
                          width: columnAxisLayout.size(
                            columnVisibilityIndex.viewIndex(Sr),
                          ),
                        },
                      },
                      Sr,
                    ),
                  ),
                  rightGap > 0 && i.jsx("col", { style: { width: rightGap } }),
                ],
              }),
              i.jsx("thead", {
                children: i.jsxs("tr", {
                  role: "row",
                  "aria-rowindex": 1,
                  children: [
                    isHeadingsVisible &&
                      i.jsx("th", {
                        role: "columnheader",
                        className: "lg-corner",
                        "aria-label": text("Select all cells"),
                        onClick: selectAllCells,
                        lang: language("Select all cells"),
                        "data-ui-label": "Select all cells",
                        children: i.jsx("span", {
                          className: "lg-sr-only",
                          children: i.jsx(LocalizedText, {
                            children: "Select all cells",
                          }),
                        }),
                      }),
                    frozenColumnIndexes.map((Sr) =>
                      renderColumnHeader(Sr, true),
                    ),
                    leftGap > 0 &&
                      i.jsx("th", {
                        className: "lg-column-spacer",
                        "aria-hidden": "true",
                      }),
                    shownColumns.map((Sr) => renderColumnHeader(Sr)),
                    rightGap > 0 &&
                      i.jsx("th", {
                        className: "lg-column-spacer",
                        "aria-hidden": "true",
                      }),
                  ],
                }),
              }),
              i.jsxs("tbody", {
                children: [
                  frozenRowIndexes.map((Sr) => renderGridRow(Sr, true, gt, fe)),
                  renderVirtualSpacer(
                    rowWindow.leading,
                    "before",
                    visibleColumnCount,
                  ),
                  shownRows.map((Sr) => renderGridRow(Sr, false, gt, fe)),
                  renderVirtualSpacer(
                    rowWindow.trailing,
                    "after",
                    visibleColumnCount,
                  ),
                ],
              }),
            ],
          }),
        }),
        i.jsx(DrawingObjectOverlay, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          selection: JSON.stringify(selectionAreas),
          calc: workbookCalculator,
          readOnly: isEditingLocked,
          scale: zoomScale,
          onWorkbookChange: commitWorkbookEdit,
          positionForCell: (Sr) => {
            const Wi = _t(Sr);
            if (
              rowVisibilityIndex.isHidden(Wi.row) ||
              columnVisibilityIndex.isHidden(Wi.col)
            )
              return null;
            const Di = scrollableRowPositionOf(Wi.row),
              $o = Wi.row < pinnedRowCount,
              Ds = Wi.col < frozenColumnCount;
            if (!$o && Di === undefined) return null;
            const qo =
                (rowHeaderWidth +
                  (Ds
                    ? columnAxisLayout.offset(
                        columnVisibilityIndex.viewIndex(Wi.col),
                      )
                    : frozenColumnsWidth +
                      scrollableColumnLayout.offset(
                        columnVisibilityIndex.viewIndex(Wi.col) - yn,
                      ) -
                      columnWindow.logicalOffset)) *
                zoomScale,
              yo =
                ((isHeadingsVisible ? 22 : 0) +
                  ($o
                    ? rowAxisLayout.offset(rowVisibilityIndex.viewIndex(Wi.row))
                    : frozenRowsHeight +
                      scrollableRowLayout.offset(Math.max(0, Di)) -
                      rowWindow.logicalOffset)) *
                zoomScale;
            return qo < -5e3 || yo < -5e3 || qo > Xe.width || yo > Xe.height
              ? null
              : { left: qo, top: yo };
          },
        }),
      ],
    });
  }
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
      activeDialogMode &&
        (() => {
          const fe = {
            mode: activeDialogMode,
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            range: activeSelectionRange,
            readOnly:
              activeDialogMode === "protection" ? readOnly : isEditingLocked,
            onClose: () => setActiveDialogMode(null),
            onWorkbookChange: commitWorkbookEdit,
          };
          if (activeDialogMode === "automation")
            return i.jsx(
              AutomationScriptsDialog,
              {
                workbook: activeWorkbook,
                sheetId: activeSheet.id,
                range: activeSelectionRange,
                readOnly: isEditingLocked,
                initialInspection: automationInspection,
                onClose: fe.onClose,
                onPreserve: (Ue) => {
                  const Xe = latestEditStateRef.current.book,
                    gt = preserveAutomationArtifact(Xe, Ue, {
                      actor,
                      readOnly: latestEditStateRef.current.readOnly,
                    });
                  (gt !== Xe && commitWorkbookChange(gt),
                    setStatusNotice(
                      "Original automation bytes preserved in native JSON. No source code was executed.",
                    ));
                },
                onRemove: (Ue) =>
                  commitWorkbookChange(
                    removeAutomationArtifact(
                      latestEditStateRef.current.book,
                      Ue,
                      {
                        actor,
                        readOnly: latestEditStateRef.current.readOnly,
                      },
                    ),
                  ),
                onApply: (Ue, Xe, gt) => {
                  const Nt = runApprovedMigration(
                    latestEditStateRef.current.book,
                    Ue,
                    {
                      actor,
                      scope: Xe,
                      readOnly: latestEditStateRef.current.readOnly,
                      retainSource: gt,
                    },
                  );
                  (commitWorkbookChange(Nt),
                    setStatusNotice(
                      "Reviewed migration applied in one Undo step; source hashes and run provenance recorded.",
                    ));
                },
                onDownload: (Ue, Xe) =>
                  saveBlobToFile(
                    new Blob([Uint8Array.from(Ue)], {
                      type: "application/octet-stream",
                    }),
                    Xe,
                  ),
              },
              automationInspection?.sha256 ?? "automation",
            );
          if (
            activeDialogMode === "search" ||
            activeDialogMode === "goToSpecial"
          )
            return i.jsx(
              FindReplaceDialog,
              {
                ...fe,
                ranges: selectionAreas,
                active: activeCellAddress,
                readOnly,
                initialTab:
                  activeDialogMode === "goToSpecial" ? "special" : "find",
                onWorkbookChange: replaceWorkbookSnapshot,
                onNavigate: navigateToSheetCell,
                onSelectRanges: selectSheetRanges,
              },
              activeDialogMode,
            );
          if (activeDialogMode === "cellStyle")
            return i.jsx(FormatCellsDialog, {
              ...fe,
              ranges: selectionAreas,
              active: activeCellAddress,
            });
          if (activeDialogMode === "themes")
            return i.jsx(ThemesAndCellStylesDialog, {
              ...fe,
              ranges: selectionAreas,
            });
          if (activeDialogMode === "conditional")
            return i.jsx(ConditionalFormattingRulesManager, { ...fe });
          if (activeDialogMode === "merge")
            return i.jsx(MergeCellsPreviewDialog, { ...fe });
          if (activeDialogMode === "numberFormat")
            return i.jsx(M4e, {
              ...fe,
              ranges: selectionAreas,
              active: activeCellAddress,
            });
          if (activeDialogMode === "outlines")
            return i.jsx(OutlineToolsPanel, { ...fe });
          if (
            activeDialogMode === "comments" ||
            activeDialogMode === "hyperlink"
          )
            return i.jsx(ReviewToolsPanel, {
              ...fe,
              range: activeCellAddress,
              mode: activeDialogMode,
              active: activeCellAddress,
              actor,
              onNavigate: navigateToSheetCell,
            });
          if (
            ["pasteSpecial", "fillSeries", "viewSettings"].includes(
              activeDialogMode,
            )
          )
            return i.jsx(
              PasteSpecialDialog,
              {
                ...fe,
                mode: activeDialogMode,
                clipboard: internalClipboardRef.current?.snapshot,
                onSelectionChange: selectSingleRange,
              },
              activeDialogMode,
            );
          if (activeDialogMode === "sheets")
            return i.jsx(WorksheetTabBar, {
              workbook: activeWorkbook,
              sheetId: tabBarSheetId ?? activeSheet.id,
              readOnly,
              onWorkbookChange: replaceWorkbookSnapshot,
              onSelectSheet: (Ue) => {
                (setActiveDialogMode(null), switchActiveWorksheet(Ue));
              },
              onClose: fe.onClose,
            });
          if (
            ["workbookProperties", "versions", "template"].includes(
              activeDialogMode,
            )
          )
            return i.jsx(WorkbookManagerDialog, {
              mode:
                activeDialogMode === "workbookProperties"
                  ? "properties"
                  : activeDialogMode,
              workbook: activeWorkbook,
              readOnly,
              persistent: !!(
                persistenceKey ||
                (workbookStore && workbookStorePersistent)
              ),
              checkpoints: (persistedWorkbookRecord?.checkpoints ?? []).filter(
                (Ue) => Ue.workbookId === activeWorkbook.id,
              ),
              onClose: fe.onClose,
              onWorkbookChange: replaceWorkbookSnapshot,
              onCreateWorkbook: (Ue) => {
                if (readOnly) throw new Error("This workbook is read-only.");
                (saveBlobToFile(
                  new Blob([encodeNativeWorkbook(Ue)], {
                    type: "application/json",
                  }),
                  `${Ue.name.replace(/[^a-zA-Z0-9 _.-]/g, "_")}.json`,
                ),
                  setStatusNotice(
                    "Created an independent workbook download. The template remains open.",
                  ));
              },
              onCheckpoint: (Ue) => {
                if (readOnly) throw new Error("This workbook is read-only.");
                const Xe = persistenceSessionRef.current,
                  gt = latestEditStateRef.current.book;
                return chainMaybePromise(persistWorkbookSnapshot(gt), () =>
                  enqueueSaveOperation((Nt) => {
                    if (latestEditStateRef.current.book !== gt)
                      throw new Error(
                        "The workbook changed. Save the checkpoint again for the current revision.",
                      );
                    return chainMaybePromise(
                      workbookPersistence.checkpoint(
                        Ue,
                        persistedWorkbookRecordRef.current.guard,
                        { signal: Nt },
                      ),
                      (wt) => {
                        acknowledgeSavedWorkbook(wt, Xe);
                      },
                    );
                  }),
                );
              },
              onCompare: (Ue) =>
                workbookPersistence.compareCheckpoint(
                  Ue,
                  latestEditStateRef.current.book,
                ),
              onRestore: (Ue) => {
                var wt;
                if (readOnly) throw new Error("This workbook is read-only.");
                if (!persistedWorkbookRecordRef.current)
                  throw new Error(
                    "Save this workbook before restoring a checkpoint.",
                  );
                const Xe = latestEditStateRef.current.book,
                  gt = persistenceSessionRef.current;
                (wt = saveAbortControllerRef.current) == null || wt.abort();
                const Nt = new AbortController();
                return (
                  (saveAbortControllerRef.current = Nt),
                  enqueueSaveOperation((jt) => {
                    if (
                      latestEditStateRef.current.book !== Xe ||
                      Nt.signal.aborted
                    )
                      throw new Error(
                        "The workbook changed. Compare the checkpoint again before restoring.",
                      );
                    const rr = () => Nt.abort();
                    jt.addEventListener("abort", rr, { once: true });
                    const Mr = (xr) => {
                        if ((jt.removeEventListener("abort", rr), gt.conflict))
                          throw gt.conflict;
                        if (
                          !gt.active ||
                          gt !== persistenceSessionRef.current ||
                          Nt.signal.aborted ||
                          latestEditStateRef.current.book !== Xe ||
                          activeCellEditRef.current ||
                          latestEditStateRef.current.readOnly
                        )
                          throw new Error(
                            "The editor changed before the restore completed. Your current local changes were retained.",
                          );
                        (commitWorkbookChange(xr.workbook, true, false, xr),
                          applyWorkbookView(xr.view),
                          setStatusNotice(
                            "Restored checkpoint as a new revision.",
                          ),
                          reportSaveStatus("saved", xr.workbook));
                      },
                      Br = workbookPersistence.restoreCheckpoint(
                        Ue,
                        Xe,
                        persistedWorkbookRecordRef.current.guard,
                        actor,
                        {
                          signal: Nt.signal,
                        },
                      );
                    return isPromiseLike(Br)
                      ? Br.then(Mr).finally(() =>
                          jt.removeEventListener("abort", rr),
                        )
                      : Mr(Br);
                  })
                );
              },
            });
          if (["table", "pivot", "query"].includes(activeDialogMode))
            return i.jsx(
              TableToolsDialog,
              { ...fe, mode: activeDialogMode },
              activeDialogMode,
            );
          if (activeDialogMode === "pivotExplore")
            return i.jsx(PivotExplorerDialog, { ...fe });
          if (["picture", "shape", "chart"].includes(activeDialogMode))
            return i.jsx(PicturePropertiesDialog, { ...fe }, activeDialogMode);
          if (
            [
              "margins",
              "orientation",
              "size",
              "printArea",
              "printTitles",
              "pageBreaks",
              "fit",
              "pageLayout",
            ].includes(activeDialogMode)
          )
            return i.jsx(PageSetupDialog, { ...fe }, activeDialogMode);
          if (activeDialogMode === "functionCompatibility")
            return i.jsx(FormulaCompatibilityDialog, {
              workbook: activeWorkbook,
              onClose: () => setActiveDialogMode(null),
              onNavigate: navigateToSheetCell,
            });
          if (["function", "names"].includes(activeDialogMode))
            return i.jsx(
              FunctionAndNameDialog,
              {
                ...fe,
                mode: activeDialogMode,
                onInsertFormula: (Ue) => {
                  (setActiveDialogMode(null),
                    beginCellEdit(Ue),
                    setFormulaCaretPosition(Ue.length),
                    setIsFormulaBarFocused(true),
                    requestAnimationFrame(() => {
                      var Xe, gt;
                      ((Xe = formulaInputRef.current) == null || Xe.focus(),
                        (gt = formulaInputRef.current) == null ||
                          gt.setSelectionRange(Ue.length, Ue.length));
                    }));
                },
              },
              activeDialogMode,
            );
          if (activeDialogMode === "populationFilter")
            return i.jsx(FilterViewsDialog, {
              workbook: activeWorkbook,
              sheetId: activeSheet.id,
              range: activeSelectionRange,
              column: activeFilterColumn,
              definition: activeFilterState?.definition,
              readOnly,
              onWorkbookChange: (Ue, Xe) => {
                if (latestEditStateRef.current.book !== activeWorkbook)
                  throw Error(
                    "The workbook changed. Reopen the filter tool before saving a view.",
                  );
                replaceWorkbookSnapshot(Ue, Xe);
              },
              onClose: fe.onClose,
              onApply: (Ue) => {
                try {
                  (applyFilterConditions(
                    activeWorkbook,
                    activeSheet.id,
                    Ue,
                    workbookCalculator,
                  ),
                    setFilterStateByScope((Xe) => ({
                      ...Xe,
                      [filterStateKey]: {
                        snapshot: activeWorkbook,
                        definition: Ue,
                      },
                    })),
                    setFilterValueSearch(""),
                    setActiveColumnSort(null),
                    setAuditPopulationScope("full"),
                    scrollActivePaneVertically(0),
                    setActiveDialogMode(null),
                    setStatusNotice(
                      "Filter view applied. Source rows, formulas and history are unchanged.",
                    ));
                } catch (Xe) {
                  setErrorNotice(
                    Xe instanceof Error
                      ? Xe.message
                      : "Recalculate before filtering.",
                  );
                }
              },
            });
          if (activeDialogMode === "externalSources")
            return i.jsx(ExternalDataRefreshDialog, {
              workbook: activeWorkbook,
              sheetId: activeSheet.id,
              adapter: externalSourceAdapter,
              reliability: activeReliabilityMonitor,
              readOnly,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (latestEditStateRef.current.book !== activeWorkbook)
                  throw Error(
                    "The workbook changed. Preview the source again.",
                  );
                commitWorkbookWithNewSheet(Ue, Xe);
              },
            });
          if (
            activeDialogMode === "solver" ||
            activeDialogMode === "statistics"
          )
            return i.jsx(DataAnalysisDialog, {
              workbook: activeWorkbook,
              sheetId: activeSheet.id,
              range: activeSelectionRange,
              initialKind: activeDialogMode,
              readOnly,
              allowHidden: true,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (latestEditStateRef.current.book !== activeWorkbook)
                  throw Error(
                    "The workbook changed. Preview the analysis again.",
                  );
                commitWorkbookWithNewSheet(Ue, Xe);
              },
            });
          if (activeDialogMode === "extensions")
            return i.jsx(AddInsAndCustomFunctionsDialog, {
              workbook: activeWorkbook,
              extensions,
              readOnly,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (
                  latestEditStateRef.current.book !== activeWorkbook ||
                  latestEditStateRef.current.readOnly
                )
                  throw Error(
                    "The workbook or permissions changed. Review extensions again.",
                  );
                replaceWorkbookSnapshot(Ue, Xe);
              },
            });
          if (activeDialogMode === "workbookLinks")
            return i.jsx(WorkbookLinksDialog, {
              workbook: activeWorkbook,
              readOnly,
              approvedSources: effectiveLinkedWorkbooks,
              onApprovedSourcesChange: setLinkedWorkbooks,
              allowFileSources: allowLinkedWorkbookFiles,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (
                  latestEditStateRef.current.book !== activeWorkbook ||
                  latestEditStateRef.current.readOnly
                )
                  throw Error(
                    "The workbook or permissions changed. Preview links again.",
                  );
                replaceWorkbookSnapshot(Ue, Xe);
              },
            });
          if (activeDialogMode === "importSources")
            return i.jsx(SourceImportDialog, {
              workbook: activeWorkbook,
              readOnly,
              extractor: documentExtractor,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (
                  latestEditStateRef.current.book !== activeWorkbook ||
                  latestEditStateRef.current.readOnly
                )
                  throw Error(
                    "The workbook or permissions changed. Preview the source again.",
                  );
                commitWorkbookWithNewSheet(Ue, Xe);
              },
            });
          if (activeDialogMode === "workbookFormats")
            return i.jsx(FormatConversionDialog, {
              workbook: activeWorkbook,
              readOnly: isEditingLocked,
              initialFile: pendingImportFile,
              onClose: () => {
                (setPendingImportFile(undefined), fe.onClose());
              },
              onWorkbookChange: (Ue) => {
                if (
                  latestEditStateRef.current.book !== activeWorkbook ||
                  latestEditStateRef.current.readOnly
                )
                  throw Error(
                    "The workbook or permissions changed. Review the conversion again.",
                  );
                (commitWorkbookChange(Ue),
                  switchActiveWorksheet(Ue.sheets.find((Xe) => !Xe.hidden).id));
              },
              onAutomation: async (Ue) => {
                const Xe = await importAutomationFile(Ue, operationScope);
                (setAutomationInspection(Xe),
                  setPendingImportFile(undefined),
                  setActiveDialogMode("automation"));
              },
            });
          if (activeDialogMode === "reportOutputs")
            return i.jsx(ReportOutputsDialog, {
              workbook: activeWorkbook,
              sheetId: activeSheet.id,
              range: activeSelectionRange,
              filter: filterResult?.context.definition,
              readOnly,
              allowHidden: true,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (latestEditStateRef.current.book !== activeWorkbook)
                  throw Error(
                    "The workbook changed. Preview the report again.",
                  );
                replaceWorkbookSnapshot(Ue, Xe);
              },
            });
          if (activeDialogMode === "dataModel")
            return i.jsx(DataModelDialog, {
              workbook: activeWorkbook,
              sheetId: activeSheet.id,
              range: activeSelectionRange,
              readOnly,
              allowHidden: true,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (latestEditStateRef.current.book !== activeWorkbook)
                  throw Error("The workbook changed. Preview the model again.");
                commitWorkbookWithNewSheet(Ue, Xe);
              },
            });
          if (activeDialogMode === "sort")
            return i.jsx(CustomSortDialog, {
              workbook: activeWorkbook,
              sheetId: activeSheet.id,
              range: activeSelectionRange,
              readOnly,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                (replaceWorkbookSnapshot(Ue, Xe),
                  setActiveColumnSort(null),
                  setFilterValueSearch(""));
              },
            });
          if (
            [
              "combinePopulations",
              "transformationRecipes",
              "consolidation",
              "textToColumns",
              "textCleanup",
              "fieldTransform",
              "duplicates",
            ].includes(activeDialogMode)
          ) {
            const Ue = {
              workbook: activeWorkbook,
              sheetId: activeSheet.id,
              range:
                activeDialogMode === "duplicates" &&
                !activeSelectionRange.includes(":")
                  ? formatOccupiedRangeAddress(activeSheet)
                  : activeSelectionRange,
              readOnly,
              allowHidden: true,
              onClose: fe.onClose,
              onWorkbookChange: commitWorkbookWithNewSheet,
            };
            return activeDialogMode === "consolidation"
              ? i.jsx(ConsolidationDialog, {
                  ...Ue,
                  initialDefinition:
                    activeSheet.consolidation?.definition ??
                    consolidationDefinition,
                  approvedWorkbooks: effectiveConsolidationWorkbooks,
                  onApprovedWorkbooksChange: setConsolidationWorkbooks,
                  allowFileSources:
                    allowConsolidationFiles && !!setConsolidationWorkbooks,
                })
              : activeDialogMode === "transformationRecipes"
                ? i.jsx(TransformationRecipesDialog, { ...Ue })
                : activeDialogMode === "combinePopulations"
                  ? i.jsx(CombinePopulationsDialog, { ...Ue })
                  : activeDialogMode === "textToColumns"
                    ? i.jsx(TextToColumnsDialog, { ...Ue })
                    : activeDialogMode === "textCleanup"
                      ? i.jsx(TextCleanupDialog, { ...Ue })
                      : activeDialogMode === "fieldTransform"
                        ? i.jsx(CleanAndReshapeDialog, { ...Ue })
                        : i.jsx(DuplicateReviewDialog, { ...Ue });
          }
          return activeDialogMode === "patterns"
            ? i.jsx(PatternCheckDialog, {
                workbook: activeWorkbook,
                sheetId: activeSheet.id,
                range:
                  auditSourceRange ||
                  (activeSelectionRange.includes(":")
                    ? activeSelectionRange
                    : formatOccupiedRangeAddress(activeSheet)),
                readOnly,
                allowHidden: true,
                onClose: fe.onClose,
              })
            : activeDialogMode === "aging"
              ? i.jsx(AgingScheduleScreen, {
                  workbook: activeWorkbook,
                  sheetId: activeSheet.id,
                  range:
                    auditSourceRange ||
                    (activeSelectionRange.includes(":")
                      ? activeSelectionRange
                      : formatOccupiedRangeAddress(activeSheet)),
                  readOnly,
                  allowHidden: true,
                  onClose: fe.onClose,
                })
              : activeDialogMode === "materiality"
                ? i.jsx(MaterialityRiskScreen, {
                    workbook: activeWorkbook,
                    sheetId: activeSheet.id,
                    range:
                      auditSourceRange ||
                      (activeSelectionRange.includes(":")
                        ? activeSelectionRange
                        : formatOccupiedRangeAddress(activeSheet)),
                    readOnly,
                    allowHidden: true,
                    onClose: fe.onClose,
                  })
                : activeDialogMode === "workbookComparison"
                  ? i.jsx(WorkbookComparisonDialog, {
                      workbook: activeWorkbook,
                      sheetId: activeSheet.id,
                      range: activeSelectionRange.includes(":")
                        ? activeSelectionRange
                        : formatOccupiedRangeAddress(activeSheet),
                      readOnly,
                      allowHidden: true,
                      onClose: fe.onClose,
                    })
                  : activeDialogMode === "scenarios" ||
                      activeDialogMode === "dataTable"
                    ? i.jsx(WhatIfAnalysisDialog, {
                        workbook: activeWorkbook,
                        sheetId: activeSheet.id,
                        address: activeCellAddress,
                        initialKind:
                          activeDialogMode === "scenarios"
                            ? "scenarios"
                            : "one-variable",
                        readOnly,
                        allowHidden: true,
                        onClose: fe.onClose,
                        onWorkbookChange: commitWorkbookWithNewSheet,
                      })
                    : activeDialogMode === "goalSeek"
                      ? i.jsx(GoalSeekDialog, {
                          workbook: activeWorkbook,
                          sheetId: activeSheet.id,
                          address: activeCellAddress,
                          readOnly,
                          allowHidden: true,
                          onClose: fe.onClose,
                          onWorkbookChange: replaceWorkbookSnapshot,
                        })
                      : activeDialogMode === "calculation"
                        ? i.jsx(CalculationOptionsDialog, {
                            workbook: activeWorkbook,
                            readOnly,
                            onClose: fe.onClose,
                            onWorkbookChange: replaceWorkbookSnapshot,
                            onCalculated: handleCalculationCompleted,
                            onCalculate: runCalculationNow,
                            onCancelCalculation: calculationRun.cancel,
                            onNavigate: navigateToSheetCell,
                          })
                        : activeDialogMode === "formulaWatches"
                          ? i.jsx(FormulaWatchWindow, {
                              workbook: activeWorkbook,
                              sheetId: activeSheet.id,
                              range: activeSelectionRange,
                              ranges: selectionAreas,
                              readOnly,
                              allowHidden: true,
                              onClose: fe.onClose,
                              onWorkbookChange: replaceWorkbookSnapshot,
                              onNavigate: navigateToSheetCell,
                            })
                          : activeDialogMode === "evaluateFormula"
                            ? i.jsx(EvaluateFormulaDialog, {
                                workbook: activeWorkbook,
                                sheetId: activeSheet.id,
                                address: activeCellAddress,
                                allowHidden: true,
                                onClose: fe.onClose,
                                onNavigate: navigateToSheetCell,
                              })
                            : activeDialogMode === "tracePrecedents" ||
                                activeDialogMode === "traceDependents"
                              ? i.jsx(CalculationInspectorDialog, {
                                  workbook: activeWorkbook,
                                  sheetId: activeSheet.id,
                                  address: activeCellAddress,
                                  initialDirection:
                                    activeDialogMode === "tracePrecedents"
                                      ? "precedents"
                                      : "dependents",
                                  allowHidden: true,
                                  onClose: fe.onClose,
                                  onNavigate: navigateToSheetCell,
                                })
                              : i.jsx(
                                  DataToolDialogRouter,
                                  { ...fe, onNavigate: navigateToSheetCell },
                                  activeDialogMode,
                                );
        })(),
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
          i.jsx(WorkbookRibbon, {
            onTool: runEditorCommand,
            sheetProtected: !!(
              (ribbonSheetProtectionDetails = activeSheet.protection) != null &&
              ribbonSheetProtectionDetails.enabled
            ),
            allowProtectionChange: !readOnly,
            workbookStructureProtected: !!(
              (workbookStructureDetails = activeWorkbook.structureProtection) !=
                null && workbookStructureDetails.enabled
            ),
            readOnly: isEditingLocked,
            busy: isOperationBusy,
            canUndo: undoStack.length > 0,
            canRedo: redoStack.length > 0,
            cell: activeCell,
            showFormulas: isShowingFormulas,
            gridlines: isGridlinesVisible,
            headings: isHeadingsVisible,
            freezeRows: hasFrozenRows,
            zoom: zoomPercent,
            onFormat: applyCellFormatPatch,
            onUndo: () => undoOrRedo(),
            onRedo: () => undoOrRedo(true),
            onImport: () => {
              return importFileInputRef.current?.click();
            },
            onExport: (fe) => void exportWorkbookAs(fe),
            onNew: () => replaceActiveWorkbook(createBlankWorkbook()),
            onDemo: () => replaceActiveWorkbook(buildSampleLedgerWorkbook()),
            onRename: () =>
              setPromptDialogState({
                kind: "rename",
                value: activeWorkbook.name,
              }),
            onAddSheet: () =>
              setPromptDialogState({
                kind: "sheet",
                value: `Sheet${activeWorkbook.sheets.length + 1}`,
              }),
            onAddRow: () => runEditorCommand("insertRow"),
            onAddColumn: () => runEditorCommand("insertColumn"),
            onAutoFitColumns: () => autoFitSelectedAxis("column"),
            onAutoFitRows: () => autoFitSelectedAxis("row"),
            onAutosum: () =>
              runGuardedAction(() =>
                applyCellPatches(
                  [
                    applyAutoSum(
                      activeWorkbook,
                      activeSheet.id,
                      activeCellAddress,
                    ),
                  ],
                  "AutoSum",
                ),
              ),
            onFillDown: () =>
              runGuardedAction(() => {
                (requireSingleSelectionRange(),
                  commitWorkbookEdit(
                    applyCellPreview(
                      activeWorkbook,
                      buildFillPreview(
                        activeWorkbook,
                        activeSheet.id,
                        activeSelectionRange,
                        "down",
                      ),
                    ),
                    "Fill down",
                  ));
              }),
            onClear: () => runGuardedAction(clearSelectedCellContents),
            onCopy: () => void copyOrCutSelection(false),
            onCut: () => void copyOrCutSelection(true),
            onPaste: () => void pasteFromClipboard(),
            onFind: () => setIsFindBarVisible((fe) => !fe),
            onSort: (fe) => {
              (setActiveColumnSort(
                fe
                  ? {
                      col: _t(activeCellAddress).col,
                      direction: fe === "asc" ? 1 : -1,
                    }
                  : null,
              ),
                scrollActivePaneVertically(0),
                gridContainerRef.current &&
                  (gridContainerRef.current.scrollTop = 0));
            },
            onToggleFilter: () => {
              (setIsFilterBarVisible((fe) => !fe), setFilterValueSearch(""));
            },
            onShowFormulas: () => setIsShowingFormulas((fe) => !fe),
            onGridlines: () => setIsGridlinesVisible((fe) => !fe),
            onHeadings: () => setIsHeadingsVisible((fe) => !fe),
            onFreeze: toggleFreezeTopRow,
            onZoom: applyZoomPercent,
            onPanel: setSidePanelKind,
            onNote: () =>
              setPromptDialogState({
                kind: "note",
                value: activeCell?.note ?? "",
              }),
            onReviewed: () =>
              applyCellFormatPatch({
                tickmark: activeCell != null && activeCell.tickmark ? "" : "✓",
              }),
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
              i.jsxs("section", {
                className: "lg-workbook",
                "aria-label": text("Spreadsheet workbook"),
                lang: language("Spreadsheet workbook"),
                "data-ui-label": "Spreadsheet workbook",
                children: [
                  i.jsxs("div", {
                    className: "lg-formula-bar",
                    children: [
                      i.jsx("input", {
                        className: "lg-name-box",
                        "aria-label": text("Selected range"),
                        value: goToReferenceText,
                        onChange: (fe) => setGoToReferenceText(fe.target.value),
                        onBlur: commitNameBoxReference,
                        onKeyDown: (fe) => {
                          var Oe;
                          fe.key === "Enter" &&
                            (commitNameBoxReference(),
                            (Oe = gridContainerRef.current) == null ||
                              Oe.focus());
                        },
                        lang: language("Selected range"),
                        "data-ui-label": "Selected range",
                      }),
                      i.jsx("button", {
                        className: "lg-formula-cancel",
                        "aria-label": text("Cancel cell edit"),
                        onPointerDown: (fe) => fe.preventDefault(),
                        onMouseDown: (fe) => fe.preventDefault(),
                        onClick: cancelCellEdit,
                        lang: language("Cancel cell edit"),
                        "data-ui-label": "Cancel cell edit",
                        children: i.jsx(X, { size: 15 }),
                      }),
                      i.jsx("button", {
                        "aria-label": text("Apply cell value"),
                        disabled: isFormulaEditingLocked(),
                        onPointerDown: (fe) => fe.preventDefault(),
                        onMouseDown: (fe) => fe.preventDefault(),
                        onClick: () => {
                          commitCellEdit() && focusGridSurface();
                        },
                        lang: language("Apply cell value"),
                        "data-ui-label": "Apply cell value",
                        children: i.jsx(Check, { size: 15 }),
                      }),
                      i.jsx("span", {
                        className: "lg-fx",
                        children: i.jsx(LocalizedText, { children: "ƒx" }),
                      }),
                      i.jsx("input", {
                        ref: formulaInputRef,
                        "aria-label": text("Formula bar"),
                        readOnly: isFormulaEditingLocked(),
                        value: cellEditText,
                        "aria-describedby":
                          (formulaInputNode = activeCellEditRef.current) !=
                            null && formulaInputNode.value.startsWith("=")
                            ? `${editorIdPrefix}-formula-help`
                            : undefined,
                        onFocus: (fe) => {
                          (!activeCellEditRef.current &&
                            cellEditText.startsWith("=") &&
                            !isEditingLocked &&
                            beginCellEdit(cellEditText),
                            setIsFormulaBarFocused(true),
                            syncFormulaEditorSelection(fe.currentTarget));
                        },
                        onSelect: (fe) => {
                          (syncFormulaEditorSelection(fe.currentTarget),
                            setFormulaCaretPosition(
                              fe.currentTarget.selectionStart ??
                                cellEditText.length,
                            ));
                        },
                        onChange: (fe) => {
                          (beginCellEdit(fe.target.value),
                            syncFormulaEditorSelection(fe.currentTarget, true),
                            setFormulaCaretPosition(
                              fe.target.selectionStart ??
                                fe.target.value.length,
                            ));
                        },
                        onBlur: (fe) => {
                          isWithinFormulaSuggestions(
                            fe.currentTarget,
                            fe.relatedTarget,
                          ) ||
                            (setIsFormulaBarFocused(false), commitCellEdit());
                        },
                        onKeyDown: (fe) => {
                          var Oe, Be;
                          if (!fe.nativeEvent.isComposing) {
                            if (fe.key === "Escape") {
                              if (
                                !formulaPointerStateRef.current.span &&
                                cellEditText.length > 1 &&
                                (Oe = formulaSuggestionListRef.current) !=
                                  null &&
                                Oe.handleKeyDown(fe)
                              )
                                return;
                              (fe.preventDefault(), cancelCellEdit());
                              return;
                            }
                            handleFormulaKeyDown(fe) ||
                              ((Be = formulaSuggestionListRef.current) !=
                                null &&
                                Be.handleKeyDown(fe)) ||
                              ((fe.key === "Enter" || fe.key === "Tab") &&
                                (fe.preventDefault(),
                                commitCellEdit() &&
                                  moveSelectionAfterCommit(
                                    fe.key,
                                    fe.shiftKey,
                                  )));
                          }
                        },
                        lang: language("Formula bar"),
                        "data-ui-label": "Formula bar",
                      }),
                      i.jsx(FormulaSuggestionList, {
                        ref: formulaSuggestionListRef,
                        onExit: () => {
                          (setIsFormulaBarFocused(false), commitCellEdit());
                        },
                        value: cellEditText,
                        caret: formulaCaretPosition,
                        workbook: activeWorkbook,
                        sheetId: activeSheet.id,
                        inputRef: formulaInputRef,
                        disabled:
                          isFormulaEditingLocked() || !isFormulaBarFocused,
                        onAccept: (fe, Oe) => {
                          (beginCellEdit(fe),
                            setFormulaCaretPosition(Oe),
                            requestAnimationFrame(() => {
                              var Be, Ue;
                              ((Be = formulaInputRef.current) == null ||
                                Be.focus(),
                                (Ue = formulaInputRef.current) == null ||
                                  Ue.setSelectionRange(Oe, Oe));
                            }));
                        },
                      }),
                    ],
                  }),
                  activeCellEditRef.current?.value.startsWith("=") &&
                    i.jsxs("div", {
                      className: "lg-formula-entry-help",
                      id: `${editorIdPrefix}-formula-help`,
                      lang: "en",
                      children: [
                        i.jsxs("span", {
                          children: [
                            "Editing",
                            " ",
                            activeWorkbook.sheets.find((fe) => {
                              return (
                                fe.id === activeCellEditRef.current?.sheetId
                              );
                            })?.name,
                            "!",
                            activeCellEditRef.current?.address,
                            " ·",
                            " ",
                            isFormulaPointMode ? "Point mode" : "Edit mode",
                            " · Select cells to insert a reference. F2 switches arrow keys; F4 locks references. Enter applies; Esc cancels.",
                          ],
                        }),
                        i.jsx("span", {
                          role: "status",
                          "aria-live": "polite",
                          "aria-atomic": "true",
                          "aria-label": "Formula reference",
                          children: activeFormulaReference
                            ? `Reference: ${activeWorkbook.sheets.find((fe) => fe.id === activeFormulaReference.sheetId)?.name}!${activeFormulaReference.range}`
                            : "Formula not yet applied",
                        }),
                      ],
                    }),
                  findDataValidationsInRange(
                    activeWorkbook,
                    activeSheet.id,
                    activeCellAddress,
                  ).map((fe) =>
                    i.jsxs(
                      "div",
                      {
                        className: "lg-view-note",
                        "aria-label": text("Validation input guidance"),
                        lang: language("Validation input guidance"),
                        "data-ui-label": "Validation input guidance",
                        children: [
                          fe.showInputMessage !== false &&
                            (fe.inputTitle || fe.inputMessage) &&
                            i.jsxs("span", {
                              children: [
                                i.jsx("strong", { children: fe.inputTitle }),
                                " ",
                                fe.inputMessage,
                                " ",
                              ],
                            }),
                          fe.kind === "list" &&
                            fe.showDropdown !== false &&
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Choose a value for ",
                                }),
                                activeCellAddress,
                                " ",
                                i.jsxs("select", {
                                  "aria-label": `Validation dropdown for ${activeCellAddress}`,
                                  disabled: isEditingLocked,
                                  value: "",
                                  onChange: (Oe) => {
                                    const Be = fe.list[Number(Oe.target.value)];
                                    Be !== undefined &&
                                      runGuardedAction(() => {
                                        (cancelCellEdit(),
                                          applyCellPatches(
                                            [
                                              {
                                                sheetId: activeSheet.id,
                                                address: activeCellAddress,
                                                cell: {
                                                  ...activeSheet.cells[
                                                    activeCellAddress
                                                  ],
                                                  raw: Be,
                                                  literal: true,
                                                },
                                              },
                                            ],
                                            "Choose validation list value",
                                          ));
                                      });
                                  },
                                  children: [
                                    i.jsx("option", {
                                      value: "",
                                      disabled: true,
                                      children: "Select…",
                                    }),
                                    fe.list.map((Oe, Be) =>
                                      i.jsx(
                                        "option",
                                        { value: Be, children: Oe },
                                        Be,
                                      ),
                                    ),
                                  ],
                                }),
                              ],
                            }),
                        ],
                      },
                      fe.id,
                    ),
                  ),
                  workbookCalculator.getSpillDiagnostic(
                    activeSheet.id,
                    activeCellAddress,
                  ) &&
                    i.jsxs("div", {
                      className: "lg-format-warning",
                      role: "status",
                      "aria-label": text("Array spill diagnostic"),
                      lang: language("Array spill diagnostic"),
                      "data-ui-label": "Array spill diagnostic",
                      children: [
                        activeCellAddress,
                        ": ",
                        workbookCalculator.getSpillDiagnostic(
                          activeSheet.id,
                          activeCellAddress,
                        ).message,
                      ],
                    }),
                  (workbookCalculator.getSpillAnchor(
                    activeSheet.id,
                    activeCellAddress,
                  ) ||
                    workbookCalculator.getSpillRange(
                      activeSheet.id,
                      activeCellAddress,
                    )) &&
                    i.jsxs("div", {
                      className: "lg-view-note",
                      role: "status",
                      "aria-label": text("Dynamic array range"),
                      lang: language("Dynamic array range"),
                      "data-ui-label": "Dynamic array range",
                      children: [
                        i.jsx(LocalizedText, { children: "Dynamic array" }),
                        " ",
                        workbookCalculator.getSpillRange(
                          activeSheet.id,
                          workbookCalculator.getSpillAnchor(
                            activeSheet.id,
                            activeCellAddress,
                          ) ?? activeCellAddress,
                        ),
                        ".",
                        workbookCalculator.getSpillAnchor(
                          activeSheet.id,
                          activeCellAddress,
                        )
                          ? ` Edit the formula in ${workbookCalculator.getSpillAnchor(activeSheet.id, activeCellAddress)}; this cell is a calculated result.`
                          : " Results resize when source data changes.",
                      ],
                    }),
                  activeCellNamedFormulaDiagnostics.length > 0 &&
                    i.jsxs("div", {
                      className: "lg-format-warning",
                      role: "status",
                      "aria-label": text("Named formula diagnostic"),
                      lang: language("Named formula diagnostic"),
                      "data-ui-label": "Named formula diagnostic",
                      children: [
                        activeCellAddress,
                        ":",
                        " ",
                        activeCellNamedFormulaDiagnostics
                          .map((fe) => fe.message)
                          .join(" "),
                      ],
                    }),
                  activeCellReferenceDiagnostics.length > 0 &&
                    i.jsxs("div", {
                      className: "lg-format-warning",
                      role: "status",
                      "aria-label": text("Formula reference diagnostic"),
                      lang: language("Formula reference diagnostic"),
                      "data-ui-label": "Formula reference diagnostic",
                      children: [
                        activeCellAddress,
                        ":",
                        " ",
                        activeCellReferenceDiagnostics
                          .map((fe) => fe.message)
                          .join(" "),
                      ],
                    }),
                  activeCellFormulaErrors.length > 0 &&
                    i.jsxs("div", {
                      className: "lg-format-warning",
                      role: "status",
                      children: [
                        activeCellAddress,
                        ":",
                        " ",
                        activeCellFormulaErrors
                          .map((fe) => fe.message)
                          .join(" "),
                        " ",
                        i.jsx("button", {
                          type: "button",
                          onClick: () => runEditorCommand("conditional"),
                          children: i.jsx(LocalizedText, {
                            children: "Review conditional rules",
                          }),
                        }),
                      ],
                    }),
                  (describeNamedStyleFallback(activeCell?.style) ||
                    activeWorkbook.styleCatalog?.missingTheme) &&
                    i.jsxs("div", {
                      className: "lg-format-warning",
                      role: "status",
                      children: [
                        describeNamedStyleFallback(activeCell?.style) ||
                          collectStyleFallbackWarnings(activeWorkbook).join(
                            " ",
                          ),
                        i.jsx("button", {
                          onClick: () => runEditorCommand("themes"),
                          children: i.jsx(LocalizedText, {
                            children: "Review themes and styles",
                          }),
                        }),
                      ],
                    }),
                  getImportedNumberFormatWarning(
                    activeSheet.cells[activeCellAddress],
                  ) &&
                    i.jsxs("div", {
                      className: "lg-view-note",
                      role: "status",
                      children: [
                        activeCellAddress,
                        ": ",
                        getImportedNumberFormatWarning(
                          activeSheet.cells[activeCellAddress],
                        ),
                      ],
                    }),
                  activeFilterState &&
                    i.jsxs("div", {
                      className: "lg-population-context",
                      role: "status",
                      "aria-label": text("Population filter context"),
                      lang: language("Population filter context"),
                      "data-ui-label": "Population filter context",
                      children: [
                        i.jsx(Filter, { size: 14 }),
                        i.jsx("strong", {
                          children: isFilterStale
                            ? "Filter paused — workbook changed"
                            : filterApplication.error
                              ? "Filter unavailable — source rows shown"
                              : `${filterResult?.context.visibleRows} visible / ${filterResult?.context.totalRows} total data rows`,
                        }),
                        filterResult &&
                          i.jsxs("span", {
                            children: [
                              filterResult.context.range,
                              " ·",
                              " ",
                              filterResult.context.filteredRows,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "filtered out ·",
                              }),
                              " ",
                              filterResult.context.hiddenMatchingRows,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "matching rows manually hidden",
                              }),
                            ],
                          }),
                        i.jsx("button", {
                          onClick: () => runEditorCommand("populationFilter"),
                          children: isFilterStale
                            ? "Review and reapply"
                            : "Edit filters",
                        }),
                        i.jsx("button", {
                          onClick: () => {
                            (setFilterStateByScope((fe) => {
                              const Oe = { ...fe };
                              return (delete Oe[filterStateKey], Oe);
                            }),
                              setAuditPopulationScope("full"));
                          },
                          children: i.jsx(LocalizedText, {
                            children: "Clear population filters",
                          }),
                        }),
                        filterResult &&
                          i.jsx("button", {
                            onClick: () =>
                              saveBlobToFile(
                                new Blob(
                                  [
                                    JSON.stringify(
                                      filterResult.context,
                                      null,
                                      2,
                                    ),
                                  ],
                                  { type: "application/json" },
                                ),
                                "Population-filter-context.json",
                              ),
                            children: i.jsx(LocalizedText, {
                              children: "Save filter context",
                            }),
                          }),
                        i.jsxs("details", {
                          children: [
                            i.jsx("summary", {
                              children: i.jsx(LocalizedText, {
                                children: "Population scope",
                              }),
                            }),
                            i.jsx("p", {
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Column filters intersect; advanced groups use OR. These are session views; named criteria can be saved in the native workbook through Edit filters. Native JSON, XLSX, print, other analysis tools and AI selections include full source data in their chosen ranges. CSV/TSV and Audit Checks offer full or visible scope. View filtering does not change SUM or ordinary formulas; SUBTOTAL/AGGREGATE retain their existing worksheet visibility semantics.",
                              }),
                            }),
                          ],
                        }),
                        filterApplication.error &&
                          i.jsx("span", {
                            role: "alert",
                            children: filterApplication.error,
                          }),
                      ],
                    }),
                  filterResult?.context.visibleRows === 0 &&
                    i.jsxs("div", {
                      className: "lg-population-empty",
                      role: "status",
                      children: [
                        i.jsx(LocalizedText, {
                          children: "No matching visible records. All",
                        }),
                        " ",
                        filterResult.context.totalRows,
                        " ",
                        i.jsx(LocalizedText, {
                          children:
                            "source data rows are retained. Clear population filters to show them again; manually hidden rows remain hidden.",
                        }),
                      ],
                    }),
                  isFilterBarVisible &&
                    i.jsxs("div", {
                      className: "lg-find-bar",
                      children: [
                        i.jsx(Filter, { size: 14 }),
                        i.jsx("input", {
                          "aria-label": text("Filter rows"),
                          placeholder: text(
                            "Filter rows (first row is a header)",
                          ),
                          value: filterValueSearch,
                          onChange: (fe) => {
                            (setFilterValueSearch(fe.target.value),
                              setFilterStateByScope((Oe) => {
                                const Be = { ...Oe };
                                return (delete Be[filterStateKey], Be);
                              }),
                              scrollActivePaneVertically(0),
                              gridContainerRef.current &&
                                (gridContainerRef.current.scrollTop = 0));
                          },
                          lang: language("Filter rows"),
                          "data-ui-label": "Filter rows",
                        }),
                        i.jsx("button", {
                          "aria-label": text("Close filter"),
                          onClick: () => {
                            (setIsFilterBarVisible(false),
                              setFilterValueSearch(""),
                              focusGridSurface());
                          },
                          lang: language("Close filter"),
                          "data-ui-label": "Close filter",
                          children: i.jsx(X, { size: 14 }),
                        }),
                      ],
                    }),
                  i.jsx(PivotResultsStatusBar, {
                    workbook: activeWorkbook,
                    sheetId: activeSheetId,
                    onOpen: () => runEditorCommand("pivot"),
                    onExplore: () => runEditorCommand("pivotExplore"),
                  }),
                  isFindBarVisible &&
                    i.jsxs("div", {
                      className: "lg-find-bar",
                      children: [
                        i.jsx(Search, { size: 14 }),
                        i.jsx("input", {
                          ref: findInputRef,
                          "aria-label": text("Find text"),
                          placeholder: text("Find"),
                          value: findQuery,
                          onChange: (fe) => setFindQuery(fe.target.value),
                          onKeyDown: (fe) => {
                            fe.key === "Enter" && findNextTextMatch();
                          },
                          lang: language("Find text"),
                          "data-ui-label": "Find text",
                        }),
                        i.jsx("button", {
                          onClick: findNextTextMatch,
                          children: i.jsx(LocalizedText, {
                            children: "Find Next",
                          }),
                        }),
                        i.jsx("input", {
                          "aria-label": text("Replace text"),
                          placeholder: text("Replace with"),
                          value: replaceWithText,
                          onChange: (fe) => setReplaceWithText(fe.target.value),
                          lang: language("Replace text"),
                          "data-ui-label": "Replace text",
                        }),
                        i.jsx("button", {
                          disabled: isEditingLocked,
                          onClick: replaceTextInSelection,
                          children: i.jsx(LocalizedText, {
                            children: "Replace in selection",
                          }),
                        }),
                        i.jsx("button", {
                          "aria-label": text("Close find"),
                          onClick: () => {
                            (setIsFindBarVisible(false), focusGridSurface());
                          },
                          lang: language("Close find"),
                          "data-ui-label": "Close find",
                          children: i.jsx(X, { size: 14 }),
                        }),
                      ],
                    }),
                  fillDragState &&
                    i.jsxs("div", {
                      className: "lg-view-note",
                      role: "status",
                      children: [
                        i.jsx(LocalizedText, { children: "Fill preview · " }),
                        fillDragState.targetRange,
                        " ",
                        i.jsx(LocalizedText, {
                          children:
                            "· Release to review, or press Escape to cancel",
                        }),
                      ],
                    }),
                  filterValueSearch &&
                    matchingRowCount === 0 &&
                    i.jsx("div", {
                      className: "lg-population-empty",
                      role: "status",
                      children: i.jsx(LocalizedText, {
                        children:
                          "No rows match this search. Source data is unchanged; clear the search to restore the view.",
                      }),
                    }),
                  (filterValueSearch || activeColumnSort) &&
                    i.jsxs("div", {
                      className: "lg-view-note",
                      children: [
                        i.jsx(Filter, { size: 12 }),
                        " ",
                        i.jsx(LocalizedText, {
                          children:
                            "View only · source row numbers and formula references stay fixed",
                        }),
                        filterValueSearch &&
                          ` · ${matchingRowCount} matching rows`,
                      ],
                    }),
                  clipboardSnapshotState &&
                    i.jsxs("div", {
                      className: "lg-view-note",
                      role: "status",
                      children: [
                        i.jsx(LocalizedText, { children: "Cut ready:" }),
                        " ",
                        activeWorkbook.sheets.find(
                          (fe) =>
                            fe.id === clipboardSnapshotState.snapshot.sheetId,
                        )?.name ?? "Deleted worksheet",
                        "!",
                        clipboardSnapshotState.snapshot.range,
                        " ",
                        i.jsx(LocalizedText, {
                          children: "· Paste to review the move.",
                        }),
                        " ",
                        i.jsx("button", {
                          type: "button",
                          onClick: clearCutSession,
                          children: i.jsx(LocalizedText, {
                            children: "Cancel cut",
                          }),
                        }),
                      ],
                    }),
                  i.jsx(SplitPaneLayout, {
                    split: sheetSplitPanes,
                    snapshot: activeWorkbook,
                    readOnly: isEditingLocked,
                    active: splitPaneSizes.activeId,
                    onResize: (fe) =>
                      runGuardedAction(() =>
                        commitWorkbookEdit(
                          setSplitPanes(
                            latestEditStateRef.current.book,
                            activeSheet.id,
                            fe,
                          ),
                          "Resize split panes",
                        ),
                      ),
                    renderPane: renderGridPane,
                  }),
                  i.jsxs("div", {
                    className: "lg-sheet-tabs",
                    "aria-label": text("Worksheet navigation"),
                    role: "group",
                    lang: language("Worksheet navigation"),
                    "data-ui-label": "Worksheet navigation",
                    children: [
                      i.jsxs("div", {
                        children: [
                          activeWorkbook.sheets
                            .filter((fe) => !fe.hidden)
                            .map((fe) =>
                              i.jsx(
                                "button",
                                {
                                  "aria-current":
                                    fe.id === activeSheet.id
                                      ? "true"
                                      : undefined,
                                  className:
                                    fe.id === activeSheet.id ? "active" : "",
                                  "data-sheet-id": fe.id,
                                  tabIndex: fe.id === activeSheet.id ? 0 : -1,
                                  onPointerDown: (Oe) => {
                                    var Be;
                                    (Be = activeCellEditRef.current) != null &&
                                      Be.value.startsWith("=") &&
                                      Oe.preventDefault();
                                  },
                                  onMouseDown: (Oe) => {
                                    var Be;
                                    (Be = activeCellEditRef.current) != null &&
                                      Be.value.startsWith("=") &&
                                      Oe.preventDefault();
                                  },
                                  onKeyDown: (Oe) => {
                                    if (
                                      ![
                                        "ArrowLeft",
                                        "ArrowRight",
                                        "Home",
                                        "End",
                                      ].includes(Oe.key) ||
                                      Oe.altKey ||
                                      Oe.ctrlKey ||
                                      Oe.metaKey
                                    )
                                      return;
                                    Oe.preventDefault();
                                    const Be = activeWorkbook.sheets.filter(
                                        (gt) => !gt.hidden,
                                      ),
                                      Ue = Be.findIndex(
                                        (gt) => gt.id === fe.id,
                                      ),
                                      Xe =
                                        Oe.key === "Home"
                                          ? Be[0]
                                          : Oe.key === "End"
                                            ? Be.at(-1)
                                            : Be[
                                                (Ue +
                                                  (Oe.key === "ArrowRight"
                                                    ? 1
                                                    : Be.length - 1)) %
                                                  Be.length
                                              ];
                                    (switchActiveWorksheet(Xe.id),
                                      requestAnimationFrame(() => {
                                        return Array.from(
                                          appRootRef.current?.querySelectorAll(
                                            "[data-sheet-id]",
                                          ) ?? [],
                                        )
                                          .find(
                                            (wt) =>
                                              wt.dataset.sheetId === Xe.id,
                                          )
                                          ?.focus();
                                      }));
                                  },
                                  onClick: () => switchActiveWorksheet(fe.id),
                                  style: fe.tabColor
                                    ? {
                                        boxShadow: `inset 0 -3px ${fe.tabColor}`,
                                      }
                                    : undefined,
                                  onDoubleClick: () => {
                                    commitCellEdit() &&
                                      (setTabBarSheetId(fe.id),
                                      setActiveDialogMode("sheets"));
                                  },
                                  onContextMenu: (Oe) => {
                                    (Oe.preventDefault(),
                                      commitCellEdit() &&
                                        (setTabBarSheetId(fe.id),
                                        setActiveDialogMode("sheets")));
                                  },
                                  children: fe.name,
                                },
                                fe.id,
                              ),
                            ),
                          i.jsx("button", {
                            className: "lg-add-sheet",
                            "aria-label": text("Add worksheet"),
                            disabled:
                              readOnly ||
                              !!(
                                (workbookStructureLock =
                                  activeWorkbook.structureProtection) != null &&
                                workbookStructureLock.enabled
                              ),
                            onClick: () =>
                              setPromptDialogState({
                                kind: "sheet",
                                value: `Sheet ${activeWorkbook.sheets.length + 1}`,
                              }),
                            lang: language("Add worksheet"),
                            "data-ui-label": "Add worksheet",
                            children: i.jsx(Plus, { size: 16 }),
                          }),
                          i.jsx("button", {
                            "aria-label": text("Manage worksheets"),
                            title: text("Manage worksheets"),
                            onClick: () => runEditorCommand("sheets"),
                            lang: language("Manage worksheets"),
                            "data-ui-label": "Manage worksheets",
                            children: i.jsx(Layers, { size: 16 }),
                          }),
                        ],
                      }),
                      i.jsx("span", {
                        className: "lg-book-name",
                        title: text("Workbook name"),
                        lang: language("Workbook name"),
                        children: activeWorkbook.name,
                      }),
                    ],
                  }),
                  i.jsxs("footer", {
                    role: "group",
                    className: "lg-grid-footer",
                    "aria-label": text("Selection summary"),
                    lang: language("Selection summary"),
                    "data-ui-label": "Selection summary",
                    children: [
                      i.jsxs("span", {
                        children: [
                          i.jsx("span", { className: "lg-ready-dot" }),
                          i.jsx(LocalizedText, {
                            children: isOperationBusy ? "Working…" : "Ready",
                          }),
                          i.jsx("button", {
                            type: "button",
                            onClick: () => setIsWorkbookHealthDialogOpen(true),
                            children: i.jsx(LocalizedText, {
                              children: "Workbook health",
                            }),
                          }),
                          isOperationBusy &&
                            activeOperationControllerRef.current &&
                            i.jsx("button", {
                              onClick: () => {
                                return activeOperationControllerRef.current?.abort();
                              },
                              children: i.jsx(LocalizedText, {
                                children: "Cancel file processing",
                              }),
                            }),
                          i.jsx("span", { className: "lg-footer-separator" }),
                          i.jsxs("span", {
                            title: text(
                              "Selection count, numeric count, sum and average include hidden and filtered cells within the selected range.",
                            ),
                            lang: language(
                              "Selection count, numeric count, sum and average include hidden and filtered cells within the selected range.",
                            ),
                            children: [
                              i.jsx(LocalizedText, { children: "Count: " }),
                              selectedCellCount,
                              " ",
                              completedSelectionAreas.length > 0
                                ? `· ${selectionAreas.length} areas `
                                : "",
                              " ",
                              i.jsx(LocalizedText, { children: "· Numbers: " }),
                              selectedNumericValues.length,
                              i.jsxs("span", {
                                className: "lg-selection-inclusion",
                                children: [
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children:
                                      "· Includes hidden and filtered cells",
                                  }),
                                ],
                              }),
                            ],
                          }),
                          isEditingLocked ? " · Read only" : "",
                        ],
                      }),
                      i.jsxs("div", {
                        children: [
                          i.jsxs("span", {
                            children: [
                              i.jsx(LocalizedText, { children: "Average" }),
                              " ",
                              i.jsx("b", {
                                children: formatNumberWithPattern(
                                  selectedNumericValues.length
                                    ? selectedNumericSum /
                                        selectedNumericValues.length
                                    : 0,
                                ),
                              }),
                            ],
                          }),
                          i.jsxs("span", {
                            children: [
                              i.jsx(LocalizedText, { children: "Sum" }),
                              " ",
                              i.jsx("b", {
                                children: formatNumberWithPattern(
                                  selectedNumericSum,
                                  {
                                    raw: 0,
                                    format: "number",
                                  },
                                ),
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        className: "lg-zoom",
                        children: [
                          i.jsx("button", {
                            "aria-label": text("Zoom out"),
                            onClick: () =>
                              applyZoomPercent(Math.max(50, zoomPercent - 10)),
                            lang: language("Zoom out"),
                            "data-ui-label": "Zoom out",
                            children: "−",
                          }),
                          i.jsx("input", {
                            "aria-label": text("Zoom"),
                            type: "range",
                            min: "50",
                            max: "200",
                            step: "10",
                            value: zoomPercent,
                            onChange: (fe) =>
                              applyZoomPercent(Number(fe.target.value)),
                            lang: language("Zoom"),
                            "data-ui-label": "Zoom",
                          }),
                          i.jsx("button", {
                            "aria-label": text("Zoom in"),
                            onClick: () =>
                              applyZoomPercent(Math.min(200, zoomPercent + 10)),
                            lang: language("Zoom in"),
                            "data-ui-label": "Zoom in",
                            children: "+",
                          }),
                          i.jsxs("span", { children: [zoomPercent, "%"] }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              sidePanelKind &&
                i.jsxs("section", {
                  "aria-label":
                    sidePanelKind === "assistant"
                      ? "Workbook assistant panel"
                      : "Workbook side panel",
                  className: `lg-panel ${sidePanelKind === "assistant" ? "lg-chat-panel" : ""}`,
                  children: [
                    i.jsxs("div", {
                      className: "lg-panel-heading",
                      children: [
                        i.jsxs("div", {
                          children: [
                            sidePanelKind === "checks"
                              ? i.jsx(ClipboardCheck, { size: 18 })
                              : sidePanelKind === "assistant"
                                ? i.jsx(Sparkles, { size: 18 })
                                : sidePanelKind === "history"
                                  ? i.jsx(History, { size: 18 })
                                  : i.jsx(CircleHelp, { size: 18 }),
                            i.jsx("h2", {
                              children:
                                sidePanelKind === "checks"
                                  ? "Audit toolkit"
                                  : sidePanelKind === "assistant"
                                    ? "Workbook assistant"
                                    : sidePanelKind === "history"
                                      ? "Activity"
                                      : "About this workspace",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          "aria-label": text("Close side panel"),
                          onClick: () => setSidePanelKind(null),
                          lang: language("Close side panel"),
                          "data-ui-label": "Close side panel",
                          children: i.jsx(PanelRightClose, { size: 16 }),
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "lg-panel-body",
                      children: [
                        sidePanelKind === "checks" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("span", {
                                className: "lg-eyebrow",
                                children: i.jsx(LocalizedText, {
                                  children: "AUDIT",
                                }),
                              }),
                              i.jsx("h3", {
                                children: i.jsx(LocalizedText, {
                                  children: "Audit checks",
                                }),
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Analysis population",
                                  }),
                                  i.jsxs("select", {
                                    "aria-label": text("Analysis population"),
                                    value: auditPopulationScope,
                                    onChange: (fe) => {
                                      (setAuditPopulationScope(fe.target.value),
                                        setAuditAnalysisResult(null));
                                    },
                                    lang: language("Analysis population"),
                                    "data-ui-label": "Analysis population",
                                    children: [
                                      i.jsx("option", {
                                        value: "full",
                                        children:
                                          "Full selected source range (including hidden/filtered)",
                                      }),
                                      i.jsx("option", {
                                        value: "visible",
                                        disabled: !filterResult,
                                        children: "Visible filter population",
                                      }),
                                    ],
                                  }),
                                ],
                              }),
                              i.jsxs("p", {
                                className: "lg-muted",
                                children: [
                                  i.jsx(LocalizedText, {
                                    children:
                                      "Applies to the five checks below. Pattern checks and other tool dialogs use their own complete source ranges and explicit source filters.",
                                  }),
                                  " ",
                                  auditPopulationScope === "visible" &&
                                  filterResult
                                    ? `Using ${filterResult.context.range}, ${filterResult.context.visibleRows} visible / ${filterResult.context.totalRows} total rows. The custom check range is ignored.`
                                    : "Full scope includes hidden and filtered rows.",
                                ],
                              }),
                              i.jsx("p", {
                                className: "lg-muted",
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Run a reproducible check, then trace each finding directly to its source.",
                                }),
                              }),
                              i.jsxs("div", {
                                className: "lg-test-choices",
                                children: [
                                  i.jsxs("button", {
                                    onClick: () => runEditorCommand("patterns"),
                                    children: [
                                      i.jsx(ListFilter, { size: 17 }),
                                      i.jsxs("span", {
                                        children: [
                                          i.jsx("strong", {
                                            children: i.jsx(LocalizedText, {
                                              children: "Pattern checks",
                                            }),
                                          }),
                                          i.jsx("small", {
                                            children: i.jsx(LocalizedText, {
                                              children:
                                                "Duplicates, sequence gaps and unusual postings",
                                            }),
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                                  [
                                    {
                                      kind: "duplicates",
                                      title: "Duplicate entries",
                                      icon: Copy,
                                      detail: "Find repeated transaction IDs",
                                    },
                                    {
                                      kind: "materiality",
                                      title: "Materiality scan",
                                      icon: ListFilter,
                                      detail:
                                        "Surface amounts above a threshold",
                                    },
                                    {
                                      kind: "sample",
                                      title: "Select a sample",
                                      icon: FlaskConical,
                                      detail:
                                        "Repeatable selection with a seed",
                                    },
                                    {
                                      kind: "summary",
                                      title: "Account summary",
                                      icon: ChartColumn,
                                      detail: "Group and total the population",
                                    },
                                    {
                                      kind: "aging",
                                      title: "Aging analysis",
                                      icon: Clock3,
                                      detail: "Bucket transactions by age",
                                    },
                                  ].map((fe) =>
                                    i.jsxs(
                                      "button",
                                      {
                                        onClick: () => {
                                          if (
                                            fe.kind === "materiality" ||
                                            fe.kind === "aging"
                                          ) {
                                            runEditorCommand(fe.kind);
                                            return;
                                          }
                                          (setSelectedAuditKind(fe.kind),
                                            setAuditAnalysisResult(null));
                                        },
                                        className:
                                          selectedAuditKind === fe.kind
                                            ? "active"
                                            : "",
                                        children: [
                                          i.jsx(fe.icon, { size: 17 }),
                                          i.jsxs("span", {
                                            children: [
                                              i.jsx("strong", {
                                                children: fe.title,
                                              }),
                                              i.jsx("small", {
                                                children: fe.detail,
                                              }),
                                            ],
                                          }),
                                          i.jsx("span", {
                                            className: "lg-radio",
                                            children:
                                              selectedAuditKind === fe.kind &&
                                              i.jsx("span", {}),
                                          }),
                                        ],
                                      },
                                      fe.kind,
                                    ),
                                  ),
                                ],
                              }),
                              i.jsxs("label", {
                                className: "lg-field",
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Source range",
                                  }),
                                  " ",
                                  i.jsx("input", {
                                    "aria-label": text("Audit source range"),
                                    placeholder: `A1:${columnIndexToLetters(activeSheet.colCount - 1)}${activeSheet.rowCount} (entire sheet)`,
                                    value: auditSourceRange,
                                    onChange: (fe) =>
                                      setAuditSourceRange(fe.target.value),
                                    lang: language("Audit source range"),
                                    "data-ui-label": "Audit source range",
                                  }),
                                ],
                              }),
                              i.jsxs("span", {
                                className: "lg-field-hint",
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "First row contains field names.",
                                  }),
                                  " ",
                                  auditPopulationScope === "visible"
                                    ? "Uses the visible rows of the applied filter range."
                                    : "Includes all source rows.",
                                ],
                              }),
                              selectedAuditKind === "materiality" &&
                                i.jsxs("label", {
                                  className: "lg-field",
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Absolute amount threshold",
                                    }),
                                    i.jsx(LocalizedNumberField, {
                                      "aria-label": text(
                                        "Materiality threshold",
                                      ),
                                      min: "0",
                                      value: materialityThreshold,
                                      onValueChange: (fe) =>
                                        setMaterialityThreshold(Number(fe)),
                                      lang: language("Materiality threshold"),
                                      "data-ui-label": "Materiality threshold",
                                    }),
                                  ],
                                }),
                              selectedAuditKind === "sample" &&
                                i.jsxs("div", {
                                  className: "lg-two-fields",
                                  children: [
                                    i.jsxs("label", {
                                      className: "lg-field",
                                      children: [
                                        i.jsx(LocalizedText, {
                                          children: "Sample size",
                                        }),
                                        i.jsx(LocalizedNumberField, {
                                          "aria-label": text("Sample size"),
                                          min: "1",
                                          value: auditSampleSize,
                                          onValueChange: (fe) =>
                                            setAuditSampleSize(Number(fe)),
                                          lang: language("Sample size"),
                                          "data-ui-label": "Sample size",
                                        }),
                                      ],
                                    }),
                                    i.jsxs("label", {
                                      className: "lg-field",
                                      children: [
                                        i.jsx(LocalizedText, {
                                          children: "Seed",
                                        }),
                                        i.jsx("input", {
                                          "aria-label": text("Sampling seed"),
                                          value: samplingSeed,
                                          onChange: (fe) =>
                                            setSamplingSeed(fe.target.value),
                                          lang: language("Sampling seed"),
                                          "data-ui-label": "Sampling seed",
                                        }),
                                      ],
                                    }),
                                  ],
                                }),
                              selectedAuditKind === "aging" &&
                                i.jsxs("label", {
                                  className: "lg-field",
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "As of date",
                                    }),
                                    i.jsx("input", {
                                      "aria-label": text("Aging as of date"),
                                      type: "date",
                                      value: agingAsOfDate,
                                      onChange: (fe) =>
                                        setAgingAsOfDate(fe.target.value),
                                      lang: language("Aging as of date"),
                                      "data-ui-label": "Aging as of date",
                                    }),
                                  ],
                                }),
                              i.jsxs("button", {
                                className: "lg-button dark full",
                                onClick: runCurrentAuditAnalysis,
                                children: [
                                  i.jsx(ClipboardCheck, { size: 15 }),
                                  i.jsx(LocalizedText, {
                                    children: "Run check",
                                  }),
                                  i.jsx(ArrowUpRight, { size: 15 }),
                                ],
                              }),
                              auditAnalysisResult
                                ? i.jsxs("div", {
                                    className: "lg-results",
                                    children: [
                                      i.jsxs("div", {
                                        className: "lg-results-heading",
                                        children: [
                                          i.jsx("h4", {
                                            children: auditAnalysisResult.title,
                                          }),
                                          i.jsx("span", {
                                            children:
                                              auditAnalysisResult.findings
                                                .length,
                                          }),
                                        ],
                                      }),
                                      i.jsx("p", {
                                        children: auditAnalysisResult.summary,
                                      }),
                                      i.jsxs("button", {
                                        className: "lg-button full",
                                        onClick: () =>
                                          saveBlobToFile(
                                            new Blob(
                                              [
                                                JSON.stringify(
                                                  agingReportPayload,
                                                  null,
                                                  2,
                                                ),
                                              ],
                                              { type: "application/json" },
                                            ),
                                            `Audit-${selectedAuditKind}-revision-${activeWorkbook.revision}.json`,
                                          ),
                                        children: [
                                          i.jsx(Download, { size: 13 }),
                                          i.jsx(LocalizedText, {
                                            children: "Export evidence report",
                                          }),
                                        ],
                                      }),
                                      auditAnalysisResult.findings.map((fe) =>
                                        i.jsxs(
                                          "article",
                                          {
                                            className: `lg-finding ${fe.severity}`,
                                            children: [
                                              i.jsx("h4", {
                                                children: fe.title,
                                              }),
                                              i.jsx("p", {
                                                children: fe.detail,
                                              }),
                                              i.jsx("div", {
                                                className: "lg-citations",
                                                children: fe.citations.map(
                                                  (Oe, Be) =>
                                                    i.jsxs(
                                                      "button",
                                                      {
                                                        onClick: () => {
                                                          (setFilterStateByScope(
                                                            (Ue) => {
                                                              const Xe = {
                                                                ...Ue,
                                                              };
                                                              return (
                                                                delete Xe[
                                                                  JSON.stringify(
                                                                    [
                                                                      activeWorkbook.id,
                                                                      Oe.sheetId,
                                                                    ],
                                                                  )
                                                                ],
                                                                Xe
                                                              );
                                                            },
                                                          ),
                                                            switchActiveWorksheet(
                                                              Oe.sheetId,
                                                            ),
                                                            selectCellAddress(
                                                              Oe.address,
                                                            ),
                                                            setAuditAnalysisResult(
                                                              auditAnalysisResult,
                                                            ),
                                                            setStatusNotice(
                                                              `${Oe.sheetName}!${Oe.address}: ${String(Oe.value ?? "")}`,
                                                            ));
                                                        },
                                                        children: [
                                                          Oe.sheetName,
                                                          "!",
                                                          Oe.address,
                                                          i.jsx(ArrowUpRight, {
                                                            size: 10,
                                                          }),
                                                        ],
                                                      },
                                                      `${Oe.address}-${Be}`,
                                                    ),
                                                ),
                                              }),
                                            ],
                                          },
                                          fe.id,
                                        ),
                                      ),
                                    ],
                                  })
                                : i.jsxs("div", {
                                    className: "lg-info-card",
                                    children: [
                                      i.jsx(ShieldCheck, { size: 18 }),
                                      i.jsxs("div", {
                                        children: [
                                          i.jsx("strong", {
                                            children: i.jsx(LocalizedText, {
                                              children:
                                                "Evidence stays connected",
                                            }),
                                          }),
                                          i.jsx("p", {
                                            children: i.jsx(LocalizedText, {
                                              children:
                                                "Every finding includes a worksheet and cell reference. Checks leave source data unchanged.",
                                            }),
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                            ],
                          }),
                        sidePanelKind === "assistant" &&
                          i.jsx(ne.Suspense, {
                            fallback: i.jsx("p", {
                              role: "status",
                              className: "lg-muted",
                              children: "Loading chat…",
                            }),
                            children: i.jsx(LazyAssistantChat, {
                              session: assistantSession,
                              setSession: setAssistantSession,
                              adapter,
                              local:
                                adapter === Vj || adapter.localOnly === true,
                              scope: JSON.stringify([
                                activeWorkbook.id,
                                activeWorkbook.revision,
                                activeSheet.id,
                                selectionAreasKey,
                                workbookInstanceKey(adapter),
                              ]),
                              selectionLabel: selectionReferenceLabel,
                              selectedCount: selectedCellCount,
                              busy: isAssistantBusy,
                              proposalId: assistantProposal?.id,
                              onSend: askAssistantQuestion,
                              onCancel: () => {
                                var fe;
                                (assistantRequestIdRef.current++,
                                  (fe = assistantAbortControllerRef.current) ==
                                    null || fe.abort(),
                                  setIsAssistantBusy(false),
                                  isAssistantBusy &&
                                    setStatusNotice(
                                      "Assistant request cancelled. No workbook changes were made.",
                                    ));
                              },
                              onReset: () => {
                                (setAssistantProposal(null),
                                  setAnswerCitations(null),
                                  setErrorNotice(""));
                              },
                              review:
                                assistantProposal && answerCitations
                                  ? i.jsxs(i.Fragment, {
                                      children: [
                                        assistantProposal &&
                                          answerCitations &&
                                          !evaluateAnswerEvidence(
                                            activeWorkbook,
                                            assistantSelection,
                                            answerCitations,
                                          ).current &&
                                          i.jsx("p", {
                                            role: "status",
                                            children: i.jsx(LocalizedText, {
                                              children:
                                                "The answer evidence is stale or unavailable. Request a fresh answer before review.",
                                            }),
                                          }),
                                        assistantProposal &&
                                          answerCitations &&
                                          evaluateAnswerEvidence(
                                            activeWorkbook,
                                            assistantSelection,
                                            answerCitations,
                                          ).current &&
                                          i.jsxs("div", {
                                            className: "lg-proposal",
                                            children: [
                                              i.jsx("span", {
                                                className: "lg-eyebrow",
                                                children: assistantProposal
                                                  .patches.length
                                                  ? "REVIEW PROPOSED CHANGES"
                                                  : "RESPONSE",
                                              }),
                                              i.jsx(
                                                AssistantAnswerEvidence,
                                                {
                                                  workbook: activeWorkbook,
                                                  selection: assistantSelection,
                                                  evidence: answerCitations,
                                                  renderText: (fe) =>
                                                    i.jsx(
                                                      LazyAssistantChatMarkdown,
                                                      {
                                                        children: fe,
                                                      },
                                                    ),
                                                  onNavigate: (fe, Oe) => {
                                                    const Be =
                                                      latestEditStateRef.current;
                                                    if (
                                                      activeCellEditRef.current
                                                    )
                                                      throw Error(
                                                        "Finish the current cell edit and request fresh evidence before following a citation.",
                                                      );
                                                    if (
                                                      fe !== Be.sheetId ||
                                                      !evaluateAnswerEvidence(
                                                        Be.book,
                                                        {
                                                          sheetId: Be.sheetId,
                                                          range: Be.range,
                                                        },
                                                        answerCitations,
                                                      ).current
                                                    )
                                                      throw Error(
                                                        "The current approved source changed. Request a fresh answer.",
                                                      );
                                                    ((filterValueSearch ||
                                                      activeColumnSort ||
                                                      activeFilterState) &&
                                                      (setFilterStateByScope(
                                                        (Ue) => {
                                                          const Xe = { ...Ue };
                                                          return (
                                                            delete Xe[
                                                              filterStateKey
                                                            ],
                                                            Xe
                                                          );
                                                        },
                                                      ),
                                                      setFilterValueSearch(""),
                                                      setActiveColumnSort(null),
                                                      setStatusNotice(
                                                        "The row view was cleared to reveal the cited cell. The approved shared range is unchanged.",
                                                      )),
                                                      setActiveCellAddress(Oe),
                                                      setPendingScrollAddress(
                                                        Oe,
                                                      ),
                                                      focusGridSurface());
                                                  },
                                                },
                                                assistantProposal.id,
                                              ),
                                              assistantProposal.patches.map(
                                                (fe) => {
                                                  var Be;
                                                  return i.jsxs(
                                                    "div",
                                                    {
                                                      className: "lg-diff",
                                                      children: [
                                                        i.jsx("b", {
                                                          children: fe.address,
                                                        }),
                                                        i.jsx("del", {
                                                          children: String(
                                                            activeSheet.cells[
                                                              fe.address
                                                            ]?.raw ?? "(blank)",
                                                          ),
                                                        }),
                                                        i.jsxs("span", {
                                                          children: [
                                                            String(
                                                              fe.cell.raw ??
                                                                "(blank)",
                                                            ),
                                                            (Be =
                                                              fe.cell.style) !=
                                                              null && Be.fill
                                                              ? " · highlight"
                                                              : "",
                                                          ],
                                                        }),
                                                        i.jsx("dl", {
                                                          className:
                                                            "lg-proposal-fields",
                                                          "aria-label":
                                                            fe.address +
                                                            " proposed field changes",
                                                          children: [
                                                            [
                                                              "Value type",
                                                              activeSheet.cells[
                                                                fe.address
                                                              ]?.raw == null
                                                                ? "blank"
                                                                : typeof activeSheet
                                                                    .cells[
                                                                    fe.address
                                                                  ]?.raw,
                                                              fe.cell.raw ===
                                                              null
                                                                ? "blank"
                                                                : typeof fe.cell
                                                                    .raw,
                                                            ],
                                                            ...Array.from(
                                                              new Set([
                                                                ...Object.keys(
                                                                  activeSheet
                                                                    .cells[
                                                                    fe.address
                                                                  ] ?? {},
                                                                ),
                                                                ...Object.keys(
                                                                  fe.cell,
                                                                ),
                                                              ]),
                                                            )
                                                              .filter(
                                                                (gt) =>
                                                                  gt !== "raw",
                                                              )
                                                              .map((gt) => {
                                                                return [
                                                                  gt,
                                                                  activeSheet
                                                                    .cells[
                                                                    fe.address
                                                                  ]?.[gt],
                                                                  fe.cell[gt],
                                                                ];
                                                              }),
                                                          ]
                                                            .filter(
                                                              ([, gt, Nt]) =>
                                                                JSON.stringify(
                                                                  gt,
                                                                ) !==
                                                                JSON.stringify(
                                                                  Nt,
                                                                ),
                                                            )
                                                            .map(
                                                              ([gt, Nt, wt]) =>
                                                                i.jsxs(
                                                                  ne.Fragment,
                                                                  {
                                                                    children: [
                                                                      i.jsx(
                                                                        "dt",
                                                                        {
                                                                          children:
                                                                            String(
                                                                              gt,
                                                                            ),
                                                                        },
                                                                      ),
                                                                      i.jsxs(
                                                                        "dd",
                                                                        {
                                                                          children:
                                                                            [
                                                                              i.jsx(
                                                                                "del",
                                                                                {
                                                                                  children:
                                                                                    Nt ===
                                                                                    undefined
                                                                                      ? "(not set)"
                                                                                      : JSON.stringify(
                                                                                          Nt,
                                                                                        ),
                                                                                },
                                                                              ),
                                                                              i.jsx(
                                                                                "ins",
                                                                                {
                                                                                  children:
                                                                                    wt ===
                                                                                    undefined
                                                                                      ? "(not set)"
                                                                                      : JSON.stringify(
                                                                                          wt,
                                                                                        ),
                                                                                },
                                                                              ),
                                                                            ],
                                                                        },
                                                                      ),
                                                                    ],
                                                                  },
                                                                  String(gt),
                                                                ),
                                                            ),
                                                        }),
                                                      ],
                                                    },
                                                    fe.address,
                                                  );
                                                },
                                              ),
                                              assistantProposal.patches.length >
                                                0 &&
                                                i.jsxs(i.Fragment, {
                                                  children: [
                                                    i.jsxs("button", {
                                                      className:
                                                        "lg-button dark full",
                                                      disabled: isEditingLocked,
                                                      onClick: () =>
                                                        runGuardedAction(() => {
                                                          const fe =
                                                              latestEditStateRef.current,
                                                            Oe = {
                                                              sheetId:
                                                                fe.sheetId,
                                                              range: fe.range,
                                                            };
                                                          if (
                                                            fe.readOnly ||
                                                            activeCellEditRef.current ||
                                                            !evaluateAnswerEvidence(
                                                              fe.book,
                                                              Oe,
                                                              answerCitations,
                                                            ).current
                                                          )
                                                            throw Error(
                                                              "The current approved source changed or editing is unavailable. Request a fresh answer before applying changes.",
                                                            );
                                                          const Be =
                                                            validateProposedCalculation(
                                                              fe.book,
                                                              Oe,
                                                              assistantProposal,
                                                            );
                                                          (applyCellPatches(
                                                            Be.patches,
                                                            `Accepted assistant proposal ${Be.id}`,
                                                            fe.selectionKey,
                                                          ),
                                                            setStatusNotice(
                                                              "Reviewed proposal applied and recorded in activity.",
                                                            ));
                                                        }),
                                                      children: [
                                                        i.jsx(Check, {
                                                          size: 15,
                                                        }),
                                                        i.jsx(LocalizedText, {
                                                          children: "Apply ",
                                                        }),
                                                        assistantProposal
                                                          .patches.length,
                                                        " ",
                                                        i.jsx(LocalizedText, {
                                                          children: "changes",
                                                        }),
                                                      ],
                                                    }),
                                                    i.jsx("button", {
                                                      className:
                                                        "lg-button full",
                                                      onClick: () => {
                                                        (setAssistantProposal(
                                                          null,
                                                        ),
                                                          setAnswerCitations(
                                                            null,
                                                          ));
                                                      },
                                                      children: i.jsx(
                                                        LocalizedText,
                                                        {
                                                          children:
                                                            "Discard proposal",
                                                        },
                                                      ),
                                                    }),
                                                  ],
                                                }),
                                            ],
                                          }),
                                      ],
                                    })
                                  : null,
                            }),
                          }),
                        sidePanelKind === "history" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("span", {
                                className: "lg-eyebrow",
                                children: i.jsx(LocalizedText, {
                                  children: "WORKBOOK ACTIVITY",
                                }),
                              }),
                              i.jsx("h3", {
                                children: i.jsx(LocalizedText, {
                                  children: "Change history",
                                }),
                              }),
                              i.jsx("p", {
                                className: "lg-muted",
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Local activity is included in JSON exports. This log is editable with the file and is not an immutable audit record.",
                                }),
                              }),
                              activeWorkbook.events.length === 0
                                ? i.jsxs("div", {
                                    className: "lg-empty",
                                    children: [
                                      i.jsx(History, { size: 26 }),
                                      i.jsx("p", {
                                        children: i.jsx(LocalizedText, {
                                          children:
                                            "Your first edit starts the activity trail.",
                                        }),
                                      }),
                                    ],
                                  })
                                : i.jsx("div", {
                                    className: "lg-timeline",
                                    children: [...activeWorkbook.events]
                                      .reverse()
                                      .map((fe) =>
                                        i.jsxs(
                                          "article",
                                          {
                                            children: [
                                              i.jsx("span", {
                                                className: "lg-timeline-dot",
                                              }),
                                              i.jsx("h4", {
                                                children: fe.action,
                                              }),
                                              i.jsxs("p", {
                                                children: [
                                                  fe.actor,
                                                  " ·",
                                                  " ",
                                                  new Date(
                                                    fe.timestamp,
                                                  ).toLocaleString(),
                                                ],
                                              }),
                                              fe.changes
                                                .slice(0, 5)
                                                .map((Oe, Be) => {
                                                  return i.jsxs(
                                                    "small",
                                                    {
                                                      children: [
                                                        activeWorkbook.sheets.find(
                                                          (Nt) =>
                                                            Nt.id ===
                                                            Oe.sheetId,
                                                        )?.name ?? "Sheet",
                                                        "!",
                                                        Oe.address,
                                                        ":",
                                                        " ",
                                                        String(
                                                          Oe.before?.raw ?? "∅",
                                                        ),
                                                        " →",
                                                        " ",
                                                        String(
                                                          Oe.after?.raw ?? "∅",
                                                        ),
                                                      ],
                                                    },
                                                    Be,
                                                  );
                                                }),
                                              fe.changes.length > 5 &&
                                                i.jsxs("small", {
                                                  children: [
                                                    "+ ",
                                                    fe.changes.length - 5,
                                                    " ",
                                                    i.jsx(LocalizedText, {
                                                      children:
                                                        "more cell changes",
                                                    }),
                                                  ],
                                                }),
                                            ],
                                          },
                                          fe.id,
                                        ),
                                      ),
                                  }),
                            ],
                          }),
                        sidePanelKind === "help" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("span", {
                                className: "lg-eyebrow",
                                children: i.jsx(LocalizedText, {
                                  children: "SPREADSHEET HELP",
                                }),
                              }),
                              i.jsx("h3", {
                                children: i.jsx(LocalizedText, {
                                  children: "Workbook help",
                                }),
                              }),
                              i.jsx("p", {
                                className: "lg-muted",
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "An embeddable React workbook for audit preparation and review.",
                                }),
                              }),
                              i.jsx("h4", {
                                children: i.jsx(LocalizedText, {
                                  children: "Working with cells",
                                }),
                              }),
                              i.jsxs("ul", {
                                className: "lg-help-list",
                                children: [
                                  i.jsx("li", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Click a cell, then type or use the formula bar. Enter commits the value.",
                                    }),
                                  }),
                                  i.jsx("li", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Enter dates as YYYY-MM-DD using the workbook's selected 1900 or 1904 date system. Other date formats stay text; prefix an ISO date with an apostrophe to keep it as text. The formula bar shows the stored serial and the grid shows the date.",
                                    }),
                                  }),
                                  i.jsx("li", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Ctrl/Command-click to add areas; Shift-click extends a range. Enter A1:D10, F1:F10 in the name box.",
                                    }),
                                  }),
                                  i.jsx("li", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Copy and paste rectangular tab-separated values with Ctrl/⌘ C and V.",
                                    }),
                                  }),
                                  i.jsx("li", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Prefix text with an apostrophe to prevent formula evaluation.",
                                    }),
                                  }),
                                  i.jsx("li", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Sort and search change the view; source row numbers remain fixed.",
                                    }),
                                  }),
                                ],
                              }),
                              i.jsx("h4", {
                                children: i.jsx(LocalizedText, {
                                  children: "Compatibility",
                                }),
                              }),
                              i.jsx("p", {
                                className: "lg-muted",
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "XLSX retains a documented subset of cells, formulas and formatting. JSON retains the complete native workpaper. Macros, Power Query, DAX, advanced PivotTables, and full Excel layout fidelity are target requirements, not features of this release.",
                                }),
                              }),
                              i.jsx(WorkbookCapacityLimits, {}),
                              i.jsxs("button", {
                                className: "lg-button full",
                                onClick: () => window.print(),
                                children: [
                                  i.jsx(FileSpreadsheet, { size: 15 }),
                                  i.jsx(LocalizedText, {
                                    children: "Print current view",
                                  }),
                                ],
                              }),
                            ],
                          }),
                      ],
                    }),
                  ],
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
function WorksheetGridRow({
  idPrefix,
  pinned,
  pinnedTop,
  frozenColumns,
  frozenColumnOffsets,
  headings,
  onRowSelect,
  onDragStart,
  onDragEnter,
  onNavigate,
  onPivotDrill,
  onFillStart,
  onFillOptions,
  fillDragging,
  commentAddresses,
  onReview,
  onLink,
  row,
  columns,
  leftGap,
  rightGap,
  minimumHeight,
  fixedHeight,
  onResize,
  onAutoFit,
  onCancelEdit,
  onFormulaInput,
  onFormulaKey,
  formulaHelpId,
  onFormulaPick,
  formulaReference,
  formulaDraftActive,
  workbook,
  mergeBounds,
  conditionalFormatter,
  displayedRows,
  sortedView,
  filterHeader,
  filteredColumns,
  onColumnFilter,
  ruleStyles,
  sheetId,
  cells,
  active,
  selectedBounds,
  selectionBounds,
  calc,
  showFormulas,
  editing,
  setEditing,
  select,
  saveCell,
  readOnly,
}) {
  const { text, language } = useWorkbookMessages(),
    ge = useWorkbookMessages(),
    se = (Te) => formatCellEntryForEdit(Te, ge.entryLocale),
    we = ne.useRef(null),
    Ee = ne.useRef(null),
    [Ae, ye] = ne.useState(0);
  return i.jsx(i.Fragment, {
    children: i.jsxs("tr", {
      role: "row",
      "data-row": row,
      style: { height: minimumHeight, ...(pinned ? { top: pinnedTop } : {}) },
      className: pinned ? "lg-pinned-row" : "",
      "aria-rowindex": row + 2,
      children: [
        headings &&
          i.jsxs("th", {
            role: "rowheader",
            scope: "row",
            "data-row-header": row,
            onClick: (Te) => onRowSelect(Te.shiftKey),
            "aria-label": `Select row ${row + 1}`,
            className: _t(active).row === row ? "lg-selected-header" : "",
            children: [
              row + 1,
              i.jsx("span", {
                className: "lg-row-resize",
                role: "separator",
                "aria-label": `Resize row ${row + 1}`,
                "aria-orientation": "horizontal",
                "aria-disabled": readOnly,
                title: text(
                  "Drag to resize; double-click to AutoFit row height",
                ),
                onPointerDown: onResize,
                onClick: (Te) => Te.stopPropagation(),
                onDoubleClick: (Te) => {
                  (Te.stopPropagation(), onAutoFit());
                },
                lang: language(
                  "Drag to resize; double-click to AutoFit row height",
                ),
              }),
            ],
          }),
        [...frozenColumns, ...columns].map((Te, Pe) => {
          const qe = mergeBounds.find((lt) => Tp(lt, row, Te));
          let ht = 1,
            Re = 1,
            Me = `${columnIndexToLetters(Te)}${row + 1}`;
          if (qe && !sortedView) {
            const lt = (
                frozenColumns.includes(Te) ? frozenColumns : columns
              ).filter((ft) => ft >= qe.firstCol && ft <= qe.lastCol),
              be = displayedRows.filter(
                (ft) => ft >= qe.firstRow && ft <= qe.lastRow,
              );
            if (Te !== lt[0] || row !== be[0]) return null;
            ((ht = lt.length),
              (Re = be.length),
              (Me = `${columnIndexToLetters(qe.firstCol)}${qe.firstRow + 1}`));
          }
          const ke = cells[Me],
            je = {
              ...ke,
              raw: ke?.raw ?? null,
              style: { ...ke?.style, ...ruleStyles(Me) },
            },
            $e = calc(sheetId, Me),
            Je =
              typeof je?.raw == "string" &&
              je.raw.startsWith("=") &&
              !je.literal,
            ze = active === Me,
            at = {
              ...cellStyleToCss(je.style, $e),
              ...(frozenColumns.includes(Te)
                ? {
                    position: "sticky",
                    left: frozenColumnOffsets[Te],
                    zIndex: 3,
                  }
                : {}),
            };
          return i.jsxs(
            ne.Fragment,
            {
              children: [
                Pe === frozenColumns.length &&
                  leftGap > 0 &&
                  i.jsx("td", {
                    className: "lg-column-spacer",
                    "aria-hidden": "true",
                  }),
                i.jsxs(
                  "td",
                  {
                    id: `${idPrefix}-${Me}`,
                    colSpan: ht,
                    rowSpan: Re,
                    "aria-colspan": ht > 1 ? ht : undefined,
                    "aria-rowspan": Re > 1 ? Re : undefined,
                    role: "gridcell",
                    "aria-label": `${Me} ${ge.entryLocale ? formatLocalizedNumberDisplay($e, je, workbook.datePolicy?.dateSystem, ge.entryLocale) : String($e ?? "")}`,
                    "aria-colindex": Te + 2,
                    "aria-selected": selectionBounds.some((lt) =>
                      Tp(lt, row, Te),
                    ),
                    "data-address": Me,
                    "data-formula-reference":
                      formulaReference && Tp(formulaReference, row, Te)
                        ? "true"
                        : undefined,
                    className: `${ze ? "lg-cell-active" : ""} ${selectionBounds.some((lt) => Tp(lt, row, Te)) ? "lg-cell-selected" : ""} ${Je ? "lg-formula-cell" : ""} ${typeof $e == "string" && $e.startsWith("#") ? "lg-cell-error" : ""} ${frozenColumns.includes(Te) ? "lg-frozen-column" : ""}`,
                    style: at,
                    title: `${Me}${je != null && je.note ? ` · ${je.note}` : ""}${Je ? ` · ${je.raw}` : ""}`,
                    onPointerDown: (lt) => {
                      lt.button !== 0 ||
                        editing !== null ||
                        (lt.preventDefault(),
                        onDragStart(Me, lt.ctrlKey || lt.metaKey),
                        select(Me, lt.shiftKey, lt.ctrlKey || lt.metaKey));
                    },
                    onPointerEnter: (lt) => onDragEnter(Me, lt.buttons),
                    onClick: (lt) => {
                      lt.detail === 0 &&
                        ((formulaDraftActive &&
                          onFormulaPick(Me, lt.shiftKey)) ||
                          select(Me, lt.shiftKey, lt.ctrlKey || lt.metaKey));
                    },
                    onDoubleClick: () => {
                      formulaDraftActive ||
                        onPivotDrill(Me) ||
                        readOnly ||
                        setEditing(se(je));
                    },
                    children: [
                      ze && editing !== null
                        ? i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("input", {
                                ref: we,
                                className: "lg-cell-editor",
                                onFocus: (lt) =>
                                  onFormulaInput(lt.currentTarget),
                                autoFocus: true,
                                readOnly,
                                "aria-label": `Edit ${Me}`,
                                "aria-describedby": formulaDraftActive
                                  ? formulaHelpId
                                  : undefined,
                                value: editing,
                                onSelect: (lt) => {
                                  (onFormulaInput(lt.currentTarget),
                                    ye(
                                      lt.currentTarget.selectionStart ??
                                        editing.length,
                                    ));
                                },
                                onChange: (lt) => {
                                  (setEditing(lt.target.value),
                                    onFormulaInput(lt.currentTarget, true),
                                    ye(
                                      lt.target.selectionStart ??
                                        lt.target.value.length,
                                    ));
                                },
                                onClick: (lt) => lt.stopPropagation(),
                                onBlur: (lt) => {
                                  isWithinFormulaSuggestions(
                                    lt.currentTarget,
                                    lt.relatedTarget,
                                  ) || saveCell(editing);
                                },
                                onKeyDown: (lt) => {
                                  var be, ft;
                                  if (
                                    (lt.stopPropagation(),
                                    !lt.nativeEvent.isComposing)
                                  ) {
                                    if (lt.key === "Escape") {
                                      if (
                                        !formulaReference &&
                                        editing.length > 1 &&
                                        (be = Ee.current) != null &&
                                        be.handleKeyDown(lt)
                                      )
                                        return;
                                      (lt.preventDefault(), onCancelEdit());
                                      return;
                                    }
                                    onFormulaKey(lt) ||
                                      ((ft = Ee.current) != null &&
                                        ft.handleKeyDown(lt)) ||
                                      ((lt.key === "Enter" ||
                                        lt.key === "Tab") &&
                                        (lt.preventDefault(),
                                        saveCell(editing) &&
                                          onNavigate(lt.key, lt.shiftKey)),
                                      lt.key === "Escape" &&
                                        (lt.preventDefault(), onCancelEdit()));
                                  }
                                },
                              }),
                              i.jsx(FormulaSuggestionList, {
                                ref: Ee,
                                onExit: () => {
                                  saveCell(editing);
                                },
                                value: editing,
                                caret: Ae,
                                workbook,
                                sheetId,
                                inputRef: we,
                                disabled: readOnly,
                                onAccept: (lt, be) => {
                                  (setEditing(lt),
                                    ye(be),
                                    requestAnimationFrame(() => {
                                      var ft, pt;
                                      ((ft = we.current) == null || ft.focus(),
                                        (pt = we.current) == null ||
                                          pt.setSelectionRange(be, be));
                                    }));
                                },
                              }),
                            ],
                          })
                        : i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("span", {
                                style: {
                                  ...textWrapToCss(je.style),
                                  ...(fixedHeight && Re === 1
                                    ? {
                                        maxHeight: minimumHeight - 1,
                                        minHeight: 0,
                                      }
                                    : {}),
                                },
                                children:
                                  showFormulas && Je
                                    ? String(je.raw)
                                    : i.jsx(ConditionalFormatIndicator, {
                                        decoration: conditionalFormatter(Me),
                                        children: i.jsx(CellValuePreview, {
                                          value: $e,
                                          cell: je,
                                          dateSystem:
                                            workbook.datePolicy?.dateSystem,
                                        }),
                                      }),
                              }),
                              filterHeader?.some((lt) => lt.column === Te) &&
                                i.jsx("button", {
                                  className: "lg-header-filter",
                                  "aria-label": `Filter ${filterHeader.find((lt) => lt.column === Te).label} column`,
                                  "aria-pressed": filteredColumns.includes(Te),
                                  onPointerDown: (lt) => lt.stopPropagation(),
                                  onClick: (lt) => {
                                    (lt.stopPropagation(), onColumnFilter(Te));
                                  },
                                  title: text("Filter this column"),
                                  lang: language("Filter this column"),
                                  children: "▾",
                                }),
                              je?.hyperlink &&
                                (je.hyperlink.kind === "external"
                                  ? i.jsx("a", {
                                      className: "lg-cell-link",
                                      href: je.hyperlink.url,
                                      target: "_blank",
                                      rel: "noopener noreferrer",
                                      "aria-label": `Open link at ${Me}`,
                                      title:
                                        je.hyperlink.label || je.hyperlink.url,
                                      onPointerDown: (lt) =>
                                        lt.stopPropagation(),
                                      onClick: (lt) => lt.stopPropagation(),
                                      children: "↗",
                                    })
                                  : i.jsx("button", {
                                      className: "lg-cell-link",
                                      "aria-label": `Open link at ${Me}`,
                                      title:
                                        je.hyperlink.label ||
                                        `Go to ${je.hyperlink.address}`,
                                      onPointerDown: (lt) =>
                                        lt.stopPropagation(),
                                      onClick: (lt) => {
                                        lt.stopPropagation();
                                        const be = je.hyperlink;
                                        be?.kind === "internal" &&
                                          onLink(be.sheetId, be.address);
                                      },
                                      children: "↗",
                                    })),
                              commentAddresses.has(Me) &&
                                i.jsx("button", {
                                  className: "lg-comment-marker",
                                  style: {
                                    right: je != null && je.note ? 8 : 0,
                                  },
                                  "aria-label": `Review comments at ${Me}`,
                                  title: text("Open review discussion"),
                                  onPointerDown: (lt) => lt.stopPropagation(),
                                  onClick: (lt) => {
                                    (lt.stopPropagation(), onReview(Me));
                                  },
                                  lang: language("Open review discussion"),
                                  children: "●",
                                }),
                              je?.note &&
                                i.jsx("i", { className: "lg-note-marker" }),
                              je?.tickmark &&
                                i.jsx("span", {
                                  className: "lg-tick",
                                  children: je.tickmark,
                                }),
                            ],
                          }),
                      !readOnly &&
                        !fillDragging &&
                        editing === null &&
                        row === selectedBounds.lastRow &&
                        Te === selectedBounds.lastCol &&
                        i.jsx("button", {
                          className: "lg-fill-handle",
                          "aria-label": text("Fill handle"),
                          title: text(
                            "Drag to extend the selection; use Fill Series for keyboard options",
                          ),
                          onPointerDown: onFillStart,
                          onClick: (lt) => {
                            (lt.stopPropagation(),
                              lt.detail === 0 && onFillOptions());
                          },
                          onDoubleClick: (lt) => lt.stopPropagation(),
                          lang: language("Fill handle"),
                          "data-ui-label": "Fill handle",
                        }),
                    ],
                  },
                  `${Me}:${row}:${Te}`,
                ),
              ],
            },
            `${Me}:${row}:${Te}`,
          );
        }),
        rightGap > 0 &&
          i.jsx("td", { className: "lg-column-spacer", "aria-hidden": "true" }),
      ],
    }),
  });
}

export { WorkbookSyncBoundary };
