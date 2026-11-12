import { WorkbookRevisionConflictError } from "../11-local-workbook-store-and-versions.js";
import {
  chainMaybePromise,
  isPromiseLike,
} from "../55-external-source-refresh.js";
import { findRecoverableCheckpoint } from "../54-workbook-save-conflicts.js";
import { saveBlobToFile } from "../13-consolidation-dialog.js";
import { encodeNativeWorkbook } from "../08-xlsx-import-export-and-preservation.js";
import {
  buildCanonicalWorkbookSnapshot,
  carryCalculatorRecordForward,
  touchWorkbookModifiedAt,
} from "../03-workbook-model.js";

function createPersistenceActions({
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
}) {
  function runCalculationNow() {
    (calculationRun.run(),
      operationStatusState.setStatusNotice(
        calculationRun.isolated
          ? "Isolated calculation requested. Source cells and audit history are unchanged."
          : "Calculation refreshed. Source cells and audit history are unchanged.",
      ));
  }
  function reportSaveStatus(saveStatus, workbook, message) {
    try {
      viewAndSelectionActions.notifyHostCallback("save-state", () =>
        setSaveAbortControllerRef.current?.({
          status: saveStatus,
          workbookId: workbook.id,
          revision: workbook.revision,
          ...(message ? { message: message } : {}),
        }),
      );
    } catch {
      operationStatusState.setErrorNotice(
        "The host could not acknowledge the save status. Export JSON to retain your work.",
      );
    }
  }
  function reportSaveError(error) {
    const normalizedError =
        error instanceof WorkbookRevisionConflictError
          ? error
          : new WorkbookRevisionConflictError(
              "unavailable",
              "The storage adapter is unavailable. The workbook could not be saved.",
            ),
      reportedError =
        normalizedError instanceof WorkbookRevisionConflictError &&
        normalizedError.code === "conflict"
          ? normalizedError
          : (persistenceState.persistenceSessionRef.current.conflict ??
            normalizedError);
    (reportedError instanceof WorkbookRevisionConflictError &&
      reportedError.code === "conflict" &&
      (persistenceState.persistenceSessionRef.current.conflict = reportedError),
      persistenceState.setPersistenceError(reportedError),
      operationStatusState.setStatusNotice(""),
      (persistenceState.persistenceMessageRef.current =
        reportedError.message +
        " Changes remain in this editor. Export JSON to keep a copy."),
      operationStatusState.setErrorNotice(
        persistenceState.persistenceMessageRef.current,
      ),
      reportSaveStatus(
        reportedError instanceof WorkbookRevisionConflictError &&
          reportedError.code === "conflict"
          ? "conflict"
          : "error",
        latestEditStateRef.current.book,
        reportedError.message,
      ));
  }
  function enqueueSaveOperation(saveOperation) {
    const persistenceSession = persistenceState.persistenceSessionRef.current;
    if (persistenceSession.controllers.size >= 20) {
      const limitError = new WorkbookRevisionConflictError(
        "limit",
        "Twenty saves are already waiting. Your latest changes remain in this editor; retry saving after the pending saves finish, or export your work.",
      );
      throw (reportSaveError(limitError), limitError);
    }
    const abortController = new AbortController();
    persistenceSession.controllers.add(abortController);
    const runSaveOperation = () => {
        if (
          !persistenceSession.active ||
          persistenceSession !==
            persistenceState.persistenceSessionRef.current ||
          abortController.signal.aborted ||
          !isEditableRef.current
        )
          throw new WorkbookRevisionConflictError(
            "unavailable",
            "The save was cancelled because this editor or its access changed.",
          );
        if (persistenceSession.conflict) throw persistenceSession.conflict;
        if (persistenceSession.stalled)
          throw new WorkbookRevisionConflictError(
            "unavailable",
            "A previous save is still unresolved. Wait for it to finish, then retry or export your local changes.",
          );
        return saveOperation(abortController.signal);
      },
      handleSaveFailure = (saveError) => {
        throw (
          persistenceSession.controllers.delete(abortController),
          persistenceSession.active &&
            persistenceSession ===
              persistenceState.persistenceSessionRef.current &&
            (persistenceState.setPendingSaveCount(
              persistenceSession.controllers.size,
            ),
            persistenceSession.controllers.size ||
              (persistenceSession.stalled = false),
            reportSaveError(saveError)),
          saveError
        );
      };
    try {
      const saveResult = persistenceSession.tail
        ? persistenceSession.tail.then(runSaveOperation)
        : runSaveOperation();
      if (!isPromiseLike(saveResult))
        return (
          persistenceSession.controllers.delete(abortController),
          saveResult
        );
      persistenceState.setPendingSaveCount(persistenceSession.controllers.size);
      const stallTimer = setTimeout(() => {
          if (
            !persistenceSession.active ||
            persistenceSession !==
              persistenceState.persistenceSessionRef.current
          )
            return;
          ((persistenceSession.stalled = true),
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
        savePromise = saveResult.then(
          (savedResult) => (
            clearTimeout(stallTimer),
            persistenceSession.controllers.delete(abortController),
            persistenceSession.active &&
              persistenceSession ===
                persistenceState.persistenceSessionRef.current &&
              (persistenceState.setPendingSaveCount(
                persistenceSession.controllers.size,
              ),
              persistenceSession.controllers.size ||
                (persistenceSession.stalled = false)),
            savedResult
          ),
          (saveFailure) => (
            clearTimeout(stallTimer),
            handleSaveFailure(saveFailure)
          ),
        ),
        settledPromise = savePromise.then(
          () => {},
          () => {},
        );
      return (
        (persistenceSession.tail = settledPromise),
        settledPromise.then(() => {
          persistenceSession.tail === settledPromise &&
            (persistenceSession.tail = undefined);
        }),
        savePromise
      );
    } catch (error) {
      return handleSaveFailure(error);
    }
  }
  function acknowledgeSavedWorkbook(savedRecord, persistenceSession) {
    if (
      !persistenceSession.active ||
      persistenceSession !== persistenceState.persistenceSessionRef.current ||
      ((persistenceState.persistedWorkbookRecordRef.current = savedRecord),
      persistenceState.setPersistedWorkbookRecord(savedRecord),
      !isEditableRef.current || persistenceSession.conflict)
    )
      return savedRecord;
    if (
      savedRecord.workbook === latestEditStateRef.current.book ||
      JSON.stringify(savedRecord.workbook) ===
        JSON.stringify(latestEditStateRef.current.book)
    ) {
      persistenceState.setPersistenceError(null);
      const persistenceMessage = persistenceState.persistenceMessageRef.current;
      (operationStatusState.setErrorNotice((previousNotice) =>
        previousNotice === persistenceMessage ? "" : previousNotice,
      ),
        (persistenceState.persistenceMessageRef.current = null),
        persistenceSession.stalled &&
          activeReliabilityMonitor
            .start("persistence", "save")
            .finish("recovered"),
        operationStatusState.setStatusNotice(
          persistenceKey
            ? "Saved to this browser"
            : workbookStore
              ? "Saved to the application"
              : "Changes kept in this session",
        ),
        reportSaveStatus("saved", savedRecord.workbook));
    }
    return savedRecord;
  }
  function recoverLocalWorkbook() {
    if (!(
      persistenceState.persistenceSessionRef.current.controllers.size ||
      !isEditableRef.current
    ))
      try {
        const book = latestEditStateRef.current.book,
          viewState = viewAndSelectionActions.buildWorkbookViewState(book),
          recoverableCheckpoint = findRecoverableCheckpoint(
            workbookPersistence,
            book,
          );
        if (recoverableCheckpoint) {
          ((persistenceState.isWorkbookLoadedRef.current = true),
            acknowledgeSavedWorkbook(
              recoverableCheckpoint,
              persistenceState.persistenceSessionRef.current,
            ),
            activeReliabilityMonitor
              .start("persistence", "save")
              .finish("recovered"));
          return;
        }
        if (!persistenceState.isWorkbookLoadedRef.current) {
          const loadedRecord = workbookPersistence.load();
          if (loadedRecord)
            throw new WorkbookRevisionConflictError(
              "conflict",
              "Saved data is available again. Compare it before replacing your local changes.",
              loadedRecord.guard,
            );
          persistenceState.isWorkbookLoadedRef.current = true;
        }
        const savePromise = persistWorkbookSnapshot(book, viewState);
        isPromiseLike(savePromise) && savePromise.catch(() => {});
      } catch (error) {
        reportSaveError(error);
      }
  }
  function persistWorkbookSnapshot(
    workbook,
    viewState = viewAndSelectionActions.buildWorkbookViewState(workbook),
  ) {
    if (!persistenceState.isWorkbookLoadedRef.current) {
      const corruptError = new WorkbookRevisionConflictError(
        "corrupt",
        "Saved data needs recovery before this session can overwrite it.",
      );
      throw (reportSaveError(corruptError), corruptError);
    }
    const workbookCopy =
        workbookStore || persistenceKey ? structuredClone(workbook) : workbook,
      viewStateCopy = structuredClone(viewState),
      persistenceSession = persistenceState.persistenceSessionRef.current;
    return (
      reportSaveStatus("saving", workbookCopy),
      enqueueSaveOperation((signal) =>
        chainMaybePromise(
          workbookPersistence.save(
            workbookCopy,
            viewStateCopy,
            persistenceState.persistedWorkbookRecordRef.current?.guard ?? null,
            {
              allowIdentityChange: !!(
                persistenceState.persistedWorkbookRecordRef.current &&
                persistenceState.persistedWorkbookRecordRef.current.workbook
                  .id !== workbookCopy.id
              ),
              signal: signal,
            },
          ),
          (savedRecord) =>
            acknowledgeSavedWorkbook(savedRecord, persistenceSession),
        ),
      )
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
  function reloadSavedWorkbook(expectedRecord) {
    const loadedRecord = workbookPersistence.load();
    if (
      !loadedRecord ||
      JSON.stringify(loadedRecord.guard) !==
        JSON.stringify(expectedRecord.guard)
    )
      throw new Error(
        "The saved version changed again. Refresh the comparison before reloading.",
      );
    const previousSession = persistenceState.persistenceSessionRef.current;
    previousSession.active = false;
    for (const controller of previousSession.controllers) controller.abort();
    ((persistenceState.persistenceSessionRef.current = {
      active: true,
      controllers: new Set(),
    }),
      persistenceState.saveAbortControllerRef.current?.abort(),
      (persistenceState.persistedWorkbookRecordRef.current = loadedRecord),
      persistenceState.setPersistedWorkbookRecord(loadedRecord),
      (latestEditStateRef.current.book = loadedRecord.workbook),
      setWorkbookSnapshot(
        buildCanonicalWorkbookSnapshot(loadedRecord.workbook),
      ),
      (formulaEntryState.activeCellEditRef.current = null),
      formulaEntryState.setCellEditState(null),
      setUndoStack([]),
      setRedoStack([]),
      assistantState.setAssistantProposal(null),
      assistantState.setAnswerCitations(null),
      auditState.setAuditAnalysisResult(null),
      viewAndSelectionActions.applyWorkbookView(loadedRecord.view),
      persistenceState.setIsSaveConflictDialogOpen(false),
      persistenceState.setPersistenceError(null),
      operationStatusState.setErrorNotice(""),
      operationStatusState.setStatusNotice(
        "Loaded the saved workbook. The discarded local undo history has been cleared.",
      ));
    try {
      viewAndSelectionActions.notifyHostCallback("change", () =>
        onChange?.(loadedRecord.workbook),
      );
    } catch {
      operationStatusState.setErrorNotice(
        "The host could not acknowledge the reloaded workbook.",
      );
    }
    reportSaveStatus("saved", loadedRecord.workbook);
  }
  function commitWorkbookChange(
    nextWorkbook,
    shouldRecordUndo = true,
    shouldBypassLock = false,
    persistedRecord,
    Xe = false,
    sheetId,
  ) {
    if (
      readOnly ||
      ((isEditingLocked || latestEditStateRef.current.readOnly) &&
        !shouldBypassLock &&
        !(
          Xe &&
          sheetId &&
          latestEditStateRef.current.book.sheets.some(
            (sheet) => sheet.id === sheetId && !sheet.protection?.enabled,
          )
        ))
    )
      throw new Error("This workbook is read-only.");
    if (!(
      !persistedRecord &&
      (nextWorkbook === latestEditStateRef.current.book ||
        (nextWorkbook.revision === latestEditStateRef.current.book.revision &&
          JSON.stringify(nextWorkbook) ===
            JSON.stringify(latestEditStateRef.current.book)))
    )) {
      if (
        (persistedRecord ||
          persistenceState.saveAbortControllerRef.current?.abort(),
        (nextWorkbook = buildCanonicalWorkbookSnapshot(
          persistedRecord || Xe
            ? nextWorkbook
            : touchWorkbookModifiedAt(nextWorkbook),
        )),
        onBeforeChange?.(nextWorkbook, latestEditStateRef.current.book),
        shouldRecordUndo)
      ) {
        const book = latestEditStateRef.current.book;
        (setUndoStack((previousUndoStack) => [
          ...previousUndoStack.slice(-49),
          book,
        ]),
          setRedoStack([]));
      }
      if (
        (carryCalculatorRecordForward(
          latestEditStateRef.current.book,
          nextWorkbook,
        ),
        (latestEditStateRef.current.book = nextWorkbook),
        setWorkbookSnapshot(nextWorkbook),
        assistantState.setAssistantProposal(null),
        assistantState.setAnswerCitations(null),
        auditState.setAuditAnalysisResult(null),
        operationStatusState.setErrorNotice(""),
        persistedRecord)
      )
        ((persistenceState.persistedWorkbookRecordRef.current =
          persistedRecord),
          persistenceState.setPersistedWorkbookRecord(persistedRecord));
      else
        try {
          const savePromise = persistWorkbookSnapshot(nextWorkbook);
          isPromiseLike(savePromise) && savePromise.catch(() => {});
        } catch (error) {
          reportSaveError(error);
        }
      try {
        viewAndSelectionActions.notifyHostCallback("change", () =>
          onChange?.(nextWorkbook),
        );
      } catch {
        operationStatusState.setErrorNotice(
          "The host could not acknowledge the change. Export JSON to retain this revision.",
        );
      }
    }
  }
  return {
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
  };
}

export { createPersistenceActions };
