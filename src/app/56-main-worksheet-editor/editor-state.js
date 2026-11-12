import ne from "react";
import { createEmptyAssistantConversation } from "../04-cell-input-paste-and-fill.js";

/* State of the worksheet editor, grouped by subject. Each hook is called once per render by
   MainWorksheetEditor and returns the current values and setters as one object. */

/** Saved-record, save session, abort, error and conflict tracking for persistence */
function usePersistenceState() {
  const persistedWorkbookRecordRef = ne.useRef(null),
    [persistedWorkbookRecord, setPersistedWorkbookRecord] = ne.useState(null),
    isWorkbookLoadedRef = ne.useRef(false),
    persistenceSessionRef = ne.useRef({
      active: false,
      controllers: new Set(),
    }),
    saveAbortControllerRef = ne.useRef(null),
    [persistenceError, setPersistenceError] = ne.useState(null),
    persistenceMessageRef = ne.useRef(null),
    [isSaveConflictDialogOpen, setIsSaveConflictDialogOpen] =
      ne.useState(false),
    [pendingSaveCount, setPendingSaveCount] = ne.useState(0);
  return {
    persistedWorkbookRecordRef,
    persistedWorkbookRecord,
    setPersistedWorkbookRecord,
    isWorkbookLoadedRef,
    persistenceSessionRef,
    saveAbortControllerRef,
    persistenceError,
    setPersistenceError,
    persistenceMessageRef,
    isSaveConflictDialogOpen,
    setIsSaveConflictDialogOpen,
    pendingSaveCount,
    setPendingSaveCount,
  };
}

/** Which dialog, side panel, prompt or inspector overlay is open */
function useDialogState() {
  const [isWorkbookHealthDialogOpen, setIsWorkbookHealthDialogOpen] =
      ne.useState(false),
    [activeDialogMode, setActiveDialogMode] = ne.useState(null),
    [automationInspection, setAutomationInspection] = ne.useState(),
    [sidePanelKind, setSidePanelKind] = ne.useState(null),
    [promptDialogState, setPromptDialogState] = ne.useState(null);
  return {
    isWorkbookHealthDialogOpen,
    setIsWorkbookHealthDialogOpen,
    activeDialogMode,
    setActiveDialogMode,
    automationInspection,
    setAutomationInspection,
    sidePanelKind,
    setSidePanelKind,
    promptDialogState,
    setPromptDialogState,
  };
}

/** Sheet tab, go-to reference and cross-sheet or scroll navigation targets */
function useNavigationState() {
  const [tabBarSheetId, setTabBarSheetId] = ne.useState(null),
    [goToReferenceText, setGoToReferenceText] = ne.useState("A1"),
    pendingCrossSheetTargetRef = ne.useRef(null),
    [pendingScrollAddress, setPendingScrollAddress] = ne.useState(null);
  return {
    tabBarSheetId,
    setTabBarSheetId,
    goToReferenceText,
    setGoToReferenceText,
    pendingCrossSheetTargetRef,
    pendingScrollAddress,
    setPendingScrollAddress,
  };
}

/** Cell editing and formula entry: caret, pointer references, edit text and validation */
function useFormulaEntryState() {
  const formulaInputRef = ne.useRef(null),
    [formulaCaretPosition, setFormulaCaretPosition] = ne.useState(0),
    [isFormulaBarFocused, setIsFormulaBarFocused] = ne.useState(false),
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
    formulaPointerMoveRef = ne.useRef(() => {}),
    [validationReviewState, setValidationReviewState] = ne.useState(null);
  return {
    formulaInputRef,
    formulaCaretPosition,
    setFormulaCaretPosition,
    isFormulaBarFocused,
    setIsFormulaBarFocused,
    cellEditText,
    setCellEditText,
    cellEditState,
    setCellEditState,
    activeCellEditRef,
    formulaPointerStateRef,
    activeFormulaReference,
    setActiveFormulaReference,
    isFormulaPointMode,
    setIsFormulaPointMode,
    formulaPointerMoveRef,
    validationReviewState,
    setValidationReviewState,
  };
}

/** Cell and range selection plus grid pointer context and root refs */
function useSelectionState() {
  const [contextMenuState, setContextMenuState] = ne.useState(null),
    appRootRef = ne.useRef(null),
    [activeCellAddress, setActiveCellAddress] = ne.useState("A1"),
    [selectionAnchorAddress, setSelectionAnchorAddress] = ne.useState("A1"),
    activeCellAddressRef = ne.useRef("A1"),
    [activeSelectionRange, setActiveSelectionRange] = ne.useState("A1"),
    [completedSelectionAreas, setCompletedSelectionAreas] = ne.useState([]),
    gridContainerRef = ne.useRef(null);
  return {
    contextMenuState,
    setContextMenuState,
    appRootRef,
    activeCellAddress,
    setActiveCellAddress,
    selectionAnchorAddress,
    setSelectionAnchorAddress,
    activeCellAddressRef,
    activeSelectionRange,
    setActiveSelectionRange,
    completedSelectionAreas,
    setCompletedSelectionAreas,
    gridContainerRef,
  };
}

/** Column filters, filter search, sort state and filter bar visibility */
function useFilterSortState() {
  const [filterValueSearch, setFilterValueSearch] = ne.useState(""),
    [filterStateByScope, setFilterStateByScope] = ne.useState({}),
    [activeFilterColumn, setActiveFilterColumn] = ne.useState(),
    [activeColumnSort, setActiveColumnSort] = ne.useState(null),
    [isFilterBarVisible, setIsFilterBarVisible] = ne.useState(false);
  return {
    filterValueSearch,
    setFilterValueSearch,
    filterStateByScope,
    setFilterStateByScope,
    activeFilterColumn,
    setActiveFilterColumn,
    activeColumnSort,
    setActiveColumnSort,
    isFilterBarVisible,
    setIsFilterBarVisible,
  };
}

/** Audit population, sampling, materiality, and aging report inputs and results */
function useAuditState() {
  const [auditPopulationScope, setAuditPopulationScope] = ne.useState("full"),
    [selectedAuditKind, setSelectedAuditKind] = ne.useState("duplicates"),
    [auditAnalysisResult, setAuditAnalysisResult] = ne.useState(null),
    [materialityThreshold, setMaterialityThreshold] = ne.useState(25e3),
    [auditSampleSize, setAuditSampleSize] = ne.useState(5),
    [samplingSeed, setSamplingSeed] = ne.useState("FY25-01"),
    [agingReportPayload, setAgingReportPayload] = ne.useState(null),
    [agingAsOfDate, setAgingAsOfDate] = ne.useState("2025-12-31"),
    [auditSourceRange, setAuditSourceRange] = ne.useState("");
  return {
    auditPopulationScope,
    setAuditPopulationScope,
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
    setAgingReportPayload,
    agingAsOfDate,
    setAgingAsOfDate,
    auditSourceRange,
    setAuditSourceRange,
  };
}

/** Status and error notices, busy flag and async operation tracking */
function useOperationStatusState() {
  const [statusNotice, setStatusNotice] = ne.useState(""),
    [errorNotice, setErrorNotice] = ne.useState(""),
    [isOperationBusy, setIsOperationBusy] = ne.useState(false),
    latestOperationIdRef = ne.useRef(0),
    activeOperationControllerRef = ne.useRef(null);
  return {
    statusNotice,
    setStatusNotice,
    errorNotice,
    setErrorNotice,
    isOperationBusy,
    setIsOperationBusy,
    latestOperationIdRef,
    activeOperationControllerRef,
  };
}

/** Delimited import and export source, pending file and xlsx conversion review */
function useImportExportState() {
  const [delimitedImportExportSource, setDelimitedImportExportSource] =
      ne.useState(null),
    [pendingImportFile, setPendingImportFile] = ne.useState(),
    [xlsxConversionReview, setXlsxConversionReview] = ne.useState(null),
    importFileInputRef = ne.useRef(null);
  return {
    delimitedImportExportSource,
    setDelimitedImportExportSource,
    pendingImportFile,
    setPendingImportFile,
    xlsxConversionReview,
    setXlsxConversionReview,
    importFileInputRef,
  };
}

/** Assistant conversation, proposals, citations, busy flag and request tracking */
function useAssistantState() {
  const [assistantSession, setAssistantSession] = ne.useState(
      createEmptyAssistantConversation,
    ),
    [assistantProposal, setAssistantProposal] = ne.useState(null),
    [answerCitations, setAnswerCitations] = ne.useState(null),
    [isAssistantBusy, setIsAssistantBusy] = ne.useState(false),
    assistantAbortControllerRef = ne.useRef(null),
    assistantRequestIdRef = ne.useRef(0),
    lastExternalProposalIdRef = ne.useRef(undefined);
  return {
    assistantSession,
    setAssistantSession,
    assistantProposal,
    setAssistantProposal,
    answerCitations,
    setAnswerCitations,
    isAssistantBusy,
    setIsAssistantBusy,
    assistantAbortControllerRef,
    assistantRequestIdRef,
    lastExternalProposalIdRef,
  };
}

/** Grid display toggles, zoom level, formula view and axis sizing */
function useViewDisplayState() {
  const [isShowingFormulas, setIsShowingFormulas] = ne.useState(false),
    [isGridlinesVisible, setIsGridlinesVisible] = ne.useState(true),
    [isHeadingsVisible, setIsHeadingsVisible] = ne.useState(true),
    [zoomPercent, setZoomPercent] = ne.useState(100),
    [axisResizeState, setAxisResizeState] = ne.useState(null),
    [bn, setAutoFitRowHeights] = ne.useState({});
  return {
    isShowingFormulas,
    setIsShowingFormulas,
    isGridlinesVisible,
    setIsGridlinesVisible,
    isHeadingsVisible,
    setIsHeadingsVisible,
    zoomPercent,
    setZoomPercent,
    axisResizeState,
    setAxisResizeState,
    bn,
    setAutoFitRowHeights,
  };
}

/** Find bar visibility, find and replace text, and find input ref */
function useFindReplaceState() {
  const [isFindBarVisible, setIsFindBarVisible] = ne.useState(false),
    [findQuery, setFindQuery] = ne.useState(""),
    [replaceWithText, setReplaceWithText] = ne.useState(""),
    findInputRef = ne.useRef(null);
  return {
    isFindBarVisible,
    setIsFindBarVisible,
    findQuery,
    setFindQuery,
    replaceWithText,
    setReplaceWithText,
    findInputRef,
  };
}

/** Fill-handle drag anchors, base selections, fill state and preview */
function useDragFillState() {
  const dragAnchorRef = ne.useRef(null),
    dragBaseSelectionsRef = ne.useRef([]),
    [fillDragState, setFillDragState] = ne.useState(null),
    [fillPreviewState, setFillPreviewState] = ne.useState(null);
  return {
    dragAnchorRef,
    dragBaseSelectionsRef,
    fillDragState,
    setFillDragState,
    fillPreviewState,
    setFillPreviewState,
  };
}

/** Internal clipboard, clipboard snapshot and pending move-cells source */
function useClipboardState() {
  const internalClipboardRef = ne.useRef(null),
    [clipboardSnapshotState, setClipboardSnapshotState] = ne.useState(null),
    [moveCellsSource, setMoveCellsSource] = ne.useState(null);
  return {
    internalClipboardRef,
    clipboardSnapshotState,
    setClipboardSnapshotState,
    moveCellsSource,
    setMoveCellsSource,
  };
}

export {
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
};
