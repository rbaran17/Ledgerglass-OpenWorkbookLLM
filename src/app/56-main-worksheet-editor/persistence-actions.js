import { WorkbookRevisionConflictError } from "../11-workbook-formats-dialog.js";
import {
  chainMaybePromise,
  isPromiseLike,
} from "../55-external-source-refresh.js";
import { findRecoverableCheckpoint } from "../54-workbook-save-conflicts.js";
import { saveBlobToFile } from "../13-consolidation-dialog.js";
import { encodeNativeWorkbook } from "../08-xlsx-export-and-preservation.js";
import {
  buildCanonicalWorkbookSnapshot,
  carryCalculatorRecordForward,
  touchWorkbookModifiedAt,
} from "../03-workbook-model.js";

function createPersistenceActions({
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
}) {
  function runCalculationNow() {
    (calculationRun.run(),
      setStatusNotice(
        calculationRun.isolated
          ? "Isolated calculation requested. Source cells and audit history are unchanged."
          : "Calculation refreshed. Source cells and audit history are unchanged.",
      ));
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
