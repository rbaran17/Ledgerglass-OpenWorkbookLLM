import i from "react/jsx-runtime";
import ne from "react";
import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import {
  isExternalSourceError,
  ExternalSourceError,
  _t,
  columnIndexToLetters,
  computeSourceControlTotals,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { buildLocalDemoProposal } from "./20-calculation-options-and-assistant.js";
import {
  snapshotSheetContent,
  EH,
  G2e,
  applySourceGrantPlans,
  authorizeSourceAccess,
  refreshSourceSnapshot,
  grantSheetSnapshotSource,
  runSourceRefreshBatch,
  normalizeExternalSourceState,
  fetchSourceCatalog,
} from "./54-workbook-save-conflicts.js";

function ExternalSourceBatchRefreshPanel(props) {
  const workbookMessages = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    latestProps = ne.useRef(props);
  latestProps.current = props;
  const connectedSheets = props.workbook.sheets.filter(
      (sheet) =>
        sheet.externalConnection &&
        !sheet.hidden &&
        !sheet.veryHidden &&
        (!props.allowedSheetIds || props.allowedSheetIds.includes(sheet.id)),
    ),
    [selectedSheetIds, setSelectedSheetIds] = ne.useState(
      connectedSheets
        .filter((unprotectedSheet) => !unprotectedSheet.protection?.enabled)
        .map((candidateSheet) => candidateSheet.id),
    ),
    [previewState, setPreviewState] = ne.useState(null),
    [disabled, setDisabled] = ne.useState(false),
    [statusMessage, setStatusMessage] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [progressMessage, setProgressMessage] = ne.useState(""),
    [reviewSheetId, setReviewSheetId] = ne.useState(""),
    [rowPageIndex, setRowPageIndex] = ne.useState(0),
    rowsSectionRef = ne.useRef(null);
  ne.useEffect(() => {
    rowsSectionRef.current?.focus();
  }, [reviewSheetId]);
  const isMountedRef = ne.useRef(true),
    refreshControllerRef = ne.useRef(null),
    scope = JSON.stringify([props.readOnly, props.allowedSheetIds]),
    workbookFingerprint = ne.useMemo(
      () =>
        fingerprintCanonicalValue([props.workbook.id, props.workbook.sheets]),
      [props.workbook],
    ),
    reviewResult =
      previewState &&
      (previewState.book === props.workbook ||
        (!previewState.result.plan &&
          previewState.contentHash === workbookFingerprint)) &&
      previewState.adapter === props.adapter &&
      previewState.scope === scope
        ? previewState.result
        : undefined;
  (ne.useEffect(() => {
    isMountedRef.current = true;
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button")?.focus(),
      () => {
        ((isMountedRef.current = false),
          refreshControllerRef.current?.abort(),
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      (refreshControllerRef.current?.abort(),
        setSelectedSheetIds((previousSelectedIds) =>
          previousSelectedIds.filter((sheetId) =>
            connectedSheets.some(
              (matchingSheet) =>
                matchingSheet.id === sheetId &&
                !matchingSheet.protection?.enabled,
            ),
          ),
        ),
        setPreviewState((previousPreviewState) =>
          previousPreviewState &&
          (previousPreviewState.book === props.workbook ||
            (!previousPreviewState.result.plan &&
              previousPreviewState.contentHash === workbookFingerprint)) &&
          previousPreviewState.adapter === props.adapter &&
          previousPreviewState.scope === scope
            ? previousPreviewState
            : null,
        ));
    }, [props.workbook, props.adapter, scope]));
  function onClick() {
    (refreshControllerRef.current?.abort(), props.onClose());
  }
  async function runBatchRefresh(shouldApply = false) {
    if (
      refreshControllerRef.current ||
      !props.adapter ||
      props.readOnly ||
      !selectedSheetIds.length ||
      (shouldApply && !reviewResult?.plan)
    )
      return;
    const propsAtStart = props,
      adapter = props.adapter,
      refreshController = new AbortController(),
      isCurrentRun = () =>
        isMountedRef.current &&
        latestProps.current.workbook === propsAtStart.workbook &&
        latestProps.current.adapter === adapter &&
        JSON.stringify([
          latestProps.current.readOnly,
          latestProps.current.allowedSheetIds,
        ]) === scope;
    ((refreshControllerRef.current = refreshController),
      setDisabled(true),
      setErrorMessage(""),
      setStatusMessage(""),
      setReviewSheetId(""),
      setProgressMessage(
        workbookMessages.text(
          shouldApply
            ? "Checking access for every source…"
            : "Preparing refresh previews…",
        ),
      ));
    const refreshOptions = {
      readOnly: props.readOnly,
      allowedSheetIds: props.allowedSheetIds,
      signal: refreshController.signal,
      reliability: props.reliability,
    };
    try {
      if (shouldApply) {
        const grantResults = await G2e(
          adapter,
          reviewResult.plan,
          refreshOptions,
        );
        if (!isCurrentRun()) return;
        if (refreshController.signal.aborted)
          throw new ExternalSourceError("cancelled");
        const nextWorkbook = applySourceGrantPlans(
          propsAtStart.workbook,
          reviewResult.plan,
          grantResults,
          refreshOptions,
        );
        (propsAtStart.onWorkbookChange(
          nextWorkbook,
          "Refresh all external sources",
        ),
          setPreviewState(null),
          setStatusMessage(
            workbookMessages.text(
              "All reviewed connections were refreshed together. Undo restores the complete previous batch.",
            ),
          ));
      } else {
        const batchResult = await runSourceRefreshBatch(
          propsAtStart.workbook,
          adapter,
          selectedSheetIds,
          {
            ...refreshOptions,
            onProgress: (completed, total) => {
              isCurrentRun() &&
                setProgressMessage(
                  workbookMessages.message(
                    "Reviewed {completed} of {total} connections.",
                    {
                      completed: completed,
                      total: total,
                    },
                  ),
                );
            },
          },
        );
        if (!isCurrentRun()) return;
        let workbook = propsAtStart.workbook;
        for (const entry of batchResult.entries) {
          const code = entry.error ?? batchResult.error;
          code &&
            (workbook = normalizeExternalSourceState(
              workbook,
              entry.sheetId,
              entry.sourceId,
              code,
              refreshOptions,
            ));
        }
        (workbook !== propsAtStart.workbook &&
          propsAtStart.onWorkbookChange(
            workbook,
            "External source batch failed; previous data retained",
          ),
          setPreviewState({
            result: batchResult,
            book: workbook,
            contentHash: fingerprintCanonicalValue([
              workbook.id,
              workbook.sheets,
            ]),
            adapter: adapter,
            scope: scope,
          }),
          batchResult.plan ||
            setErrorMessage(
              workbookMessages.text(
                "The batch is not ready. No source values were replaced. Resolve the failed connections, or select fewer connections and preview again.",
              ),
            ));
      }
    } catch (error) {
      if (!isCurrentRun()) return;
      const sourceError = isExternalSourceError(error);
      (setErrorMessage(sourceError.message), setPreviewState(null));
      let workbook = propsAtStart.workbook;
      for (const selectedSheetId of selectedSheetIds) {
        const sourceId = propsAtStart.workbook.sheets.find(
          (sheet) => sheet.id === selectedSheetId,
        )?.externalConnection?.source.id;
        sourceId &&
          (workbook = normalizeExternalSourceState(
            workbook,
            selectedSheetId,
            sourceId,
            sourceError.code,
            refreshOptions,
          ));
      }
      if (workbook !== propsAtStart.workbook)
        try {
          propsAtStart.onWorkbookChange(
            workbook,
            "External source batch failed; previous data retained",
          );
        } catch {}
    } finally {
      refreshControllerRef.current === refreshController &&
        ((refreshControllerRef.current = null),
        isMountedRef.current && setDisabled(false));
    }
  }
  const reviewSnapshot = reviewResult?.entries.find(
    (entry) => entry.sheetId === reviewSheetId,
  )?.snapshot;
  return i.jsx("div", {
    className: "lg-modal-backdrop",
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "lg-external-sources",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.key === "Escape" && (event.stopPropagation(), onClick()),
          event.key === "Tab")
        ) {
          const focusableElements = [
            ...event.currentTarget.querySelectorAll(
              'button:not(:disabled),input:not(:disabled),[tabindex="0"]',
            ),
          ].filter((element) => !element.closest("[hidden]"));
          event.shiftKey && document.activeElement === focusableElements[0]
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
                i.jsx("p", {
                  className: "lg-source-eyebrow",
                  children: i.jsx(LocalizedText, {
                    children: "Connections & Refresh",
                  }),
                }),
                i.jsx("h2", {
                  id: titleId,
                  children: i.jsx(LocalizedText, { children: "Refresh All" }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Compare connected worksheets, then apply all selected refreshes together.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              onClick: onClick,
              children: i.jsx(LocalizedText, {
                children: "Back to connections",
              }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-source-body",
          children: [
            i.jsx("p", {
              className: "lg-source-help",
              children: i.jsx(LocalizedText, {
                children:
                  "Only visible connections permitted by your application are listed. Protected worksheets cannot be refreshed. Each source is read once per batch; source snapshots may have different timestamps.",
              }),
            }),
            i.jsxs("fieldset", {
              className: "lg-source-batch-list",
              disabled: disabled || !!props.readOnly,
              children: [
                i.jsx("legend", {
                  children: i.jsx(LocalizedText, {
                    children: "Connections to refresh",
                  }),
                }),
                connectedSheets.map((connectedSheet) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: selectedSheetIds.includes(connectedSheet.id),
                          disabled: !!connectedSheet.protection?.enabled,
                          onChange: (event) => {
                            (setSelectedSheetIds((currentSelectedIds) =>
                              event.target.checked
                                ? [...currentSelectedIds, connectedSheet.id]
                                : currentSelectedIds.filter(
                                    (selectedId) =>
                                      selectedId !== connectedSheet.id,
                                  ),
                            ),
                              setPreviewState(null),
                              setErrorMessage(""),
                              setStatusMessage(""),
                              setReviewSheetId(""));
                          },
                        }),
                        i.jsxs("span", {
                          children: [
                            i.jsx("strong", { children: connectedSheet.name }),
                            " ·",
                            " ",
                            connectedSheet.externalConnection.source.label,
                            connectedSheet.protection?.enabled
                              ? ` · ${workbookMessages.text("Protected")}`
                              : "",
                            i.jsxs("small", {
                              children: [
                                connectedSheet.externalConnection.source.id,
                                " ·",
                                " ",
                                connectedSheet.externalConnection.lastSuccess
                                  .rows,
                                " ",
                                workbookMessages.text("records"),
                                " ·",
                                " ",
                                connectedSheet.externalConnection.lastSuccess
                                  .at,
                                connectedSheet.externalConnection.lastAttempt
                                  .outcome !== "succeeded"
                                  ? ` · ${workbookMessages.text("Stale source data")}`
                                  : "",
                              ],
                            }),
                          ],
                        }),
                      ],
                    },
                    connectedSheet.id,
                  ),
                ),
                !connectedSheets.length &&
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children: "No connected worksheets are available.",
                    }),
                  }),
              ],
            }),
            errorMessage &&
              i.jsx("p", {
                role: "alert",
                className: "lg-source-error",
                children: errorMessage,
              }),
            statusMessage &&
              i.jsx("p", { role: "status", children: statusMessage }),
            disabled &&
              i.jsxs("p", {
                role: "status",
                children: [
                  progressMessage,
                  " ",
                  i.jsx("button", {
                    type: "button",
                    onClick: () => refreshControllerRef.current?.abort(),
                    children: i.jsx(LocalizedText, {
                      children: "Cancel refresh",
                    }),
                  }),
                ],
              }),
            reviewResult &&
              i.jsx("section", {
                "aria-label": workbookMessages.text("Refresh All results"),
                className: "lg-source-preview",
                children: reviewResult.entries.map((entry) => {
                  const sheet = props.workbook.sheets.find(
                      (sheet) => sheet.id === entry.sheetId,
                    ),
                    lastSuccess = sheet.externalConnection.lastSuccess,
                    errorCode = entry.error ?? reviewResult.error;
                  return i.jsxs(
                    "article",
                    {
                      className: "lg-source-history",
                      children: [
                        i.jsxs("h3", {
                          children: [
                            sheet.name,
                            " ·",
                            " ",
                            workbookMessages.text(
                              errorCode ? "Not ready" : "Ready for review",
                            ),
                          ],
                        }),
                        i.jsxs("p", {
                          children: [
                            entry.sourceId,
                            " · ",
                            sheet.externalConnection.source.kind,
                          ],
                        }),
                        errorCode &&
                          i.jsx("p", {
                            children: isExternalSourceError(
                              new ExternalSourceError(errorCode),
                            ).message,
                          }),
                        entry.snapshot &&
                          !errorCode &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("p", {
                                children: [
                                  workbookMessages.text("Source snapshot"),
                                  ":",
                                  " ",
                                  entry.snapshot.sourceTimestamp,
                                  " ·",
                                  " ",
                                  entry.snapshot.version,
                                ],
                              }),
                              i.jsxs("p", {
                                children: [
                                  workbookMessages.text("Rows"),
                                  ": ",
                                  lastSuccess.rows,
                                  " →",
                                  " ",
                                  entry.snapshot.rows.length,
                                  " · ",
                                  workbookMessages.text("Change"),
                                  ":",
                                  " ",
                                  entry.snapshot.rows.length - lastSuccess.rows,
                                ],
                              }),
                              i.jsx("div", {
                                className: "lg-source-scroll",
                                children: i.jsxs("table", {
                                  "aria-label": workbookMessages.message(
                                    "Control changes for {sheet}",
                                    { sheet: sheet.name },
                                  ),
                                  children: [
                                    i.jsx("thead", {
                                      children: i.jsx("tr", {
                                        children: [
                                          "Control field",
                                          "Previous total",
                                          "New total",
                                          "Change",
                                          "Numbers / blanks",
                                        ].map((headerLabel) =>
                                          i.jsx(
                                            "th",
                                            {
                                              scope: "col",
                                              children:
                                                workbookMessages.text(
                                                  headerLabel,
                                                ),
                                            },
                                            headerLabel,
                                          ),
                                        ),
                                      }),
                                    }),
                                    i.jsx("tbody", {
                                      children: computeSourceControlTotals(
                                        entry.snapshot,
                                      ).map((controlTotal) => {
                                        const previousTotal =
                                          lastSuccess.controls.find(
                                            (control) =>
                                              control.field ===
                                              controlTotal.field,
                                          )?.sum ?? 0;
                                        return i.jsxs(
                                          "tr",
                                          {
                                            children: [
                                              i.jsx("th", {
                                                scope: "row",
                                                children: controlTotal.field,
                                              }),
                                              i.jsx("td", {
                                                children: previousTotal,
                                              }),
                                              i.jsx("td", {
                                                children: controlTotal.sum,
                                              }),
                                              i.jsx("td", {
                                                children:
                                                  controlTotal.sum -
                                                  previousTotal,
                                              }),
                                              i.jsxs("td", {
                                                children: [
                                                  controlTotal.numeric,
                                                  " / ",
                                                  controlTotal.blanks,
                                                ],
                                              }),
                                            ],
                                          },
                                          controlTotal.field,
                                        );
                                      }),
                                    }),
                                  ],
                                }),
                              }),
                              i.jsx("button", {
                                type: "button",
                                disabled: disabled,
                                onClick: () => {
                                  (setReviewSheetId(entry.sheetId),
                                    setRowPageIndex(0));
                                },
                                children: workbookMessages.message(
                                  "Review rows for {sheet}",
                                  {
                                    sheet: sheet.name,
                                  },
                                ),
                              }),
                            ],
                          }),
                      ],
                    },
                    entry.sheetId,
                  );
                }),
              }),
            reviewSnapshot &&
              i.jsxs("section", {
                ref: rowsSectionRef,
                tabIndex: -1,
                "aria-label": workbookMessages.text("Batch source rows"),
                children: [
                  i.jsx("h3", {
                    children: props.workbook.sheets.find(
                      (sheet) => sheet.id === reviewSheetId,
                    )?.name,
                  }),
                  i.jsx("div", {
                    className: "lg-source-scroll",
                    tabIndex: 0,
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsx("tr", {
                            children: reviewSnapshot.source.columns.map(
                              (column) =>
                                i.jsxs(
                                  "th",
                                  {
                                    scope: "col",
                                    children: [
                                      column.name,
                                      " (",
                                      column.type,
                                      ")",
                                    ],
                                  },
                                  column.name,
                                ),
                            ),
                          }),
                        }),
                        i.jsx("tbody", {
                          children: reviewSnapshot.rows
                            .slice(rowPageIndex * 25, (rowPageIndex + 1) * 25)
                            .map((row, rowIndex) =>
                              i.jsx(
                                "tr",
                                {
                                  children: row.map((cell, cellIndex) =>
                                    i.jsx(
                                      "td",
                                      {
                                        children:
                                          cell === null
                                            ? workbookMessages.text("(blank)")
                                            : String(cell),
                                      },
                                      cellIndex,
                                    ),
                                  ),
                                },
                                rowIndex,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("nav", {
                    "aria-label": workbookMessages.text("Batch row pages"),
                    children: [
                      i.jsx("button", {
                        type: "button",
                        disabled: rowPageIndex === 0,
                        onClick: () =>
                          setRowPageIndex(
                            (previousPageIndex) => previousPageIndex - 1,
                          ),
                        children: i.jsx(LocalizedText, {
                          children: "Previous records",
                        }),
                      }),
                      i.jsx("span", {
                        children: workbookMessages.message(
                          "Page {page} of {pages}",
                          {
                            page: rowPageIndex + 1,
                            pages: Math.max(
                              1,
                              Math.ceil(reviewSnapshot.rows.length / 25),
                            ),
                          },
                        ),
                      }),
                      i.jsx("button", {
                        type: "button",
                        disabled:
                          (rowPageIndex + 1) * 25 >= reviewSnapshot.rows.length,
                        onClick: () =>
                          setRowPageIndex(
                            (currentPageIndex) => currentPageIndex + 1,
                          ),
                        children: i.jsx(LocalizedText, {
                          children: "Next records",
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
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Apply checks every source permission again. One failed connection prevents the entire batch from replacing data.",
              }),
            }),
            i.jsx("button", {
              type: "button",
              disabled:
                disabled ||
                !!props.readOnly ||
                !props.adapter ||
                !selectedSheetIds.length,
              onClick: () => void runBatchRefresh(),
              children: i.jsx(LocalizedText, {
                children: "Preview selected connections",
              }),
            }),
            i.jsx("button", {
              type: "button",
              className: "primary",
              disabled: disabled || !!props.readOnly || !reviewResult?.plan,
              onClick: () => void runBatchRefresh(true),
              children: i.jsx(LocalizedText, {
                children: "Apply all reviewed refreshes",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
function ExternalSourceStatusPanel({ workbook, sheetId, onReview }) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    connection = sheet?.externalConnection;
  if (!sheet || !connection) return null;
  let hasLocalChanges = true;
  try {
    hasLocalChanges = snapshotSheetContent(sheet) !== connection.outputHash;
  } catch {}
  return i.jsxs("div", {
    className: "lg-source-status",
    role: "status",
    "aria-label": "External source status",
    children: [
      i.jsx("strong", {
        children:
          connection.lastAttempt.outcome !== "succeeded" || hasLocalChanges
            ? "Stale source data"
            : "Source snapshot",
      }),
      i.jsxs("span", {
        children: [
          connection.source.label,
          " · ",
          connection.lastSuccess.rows,
          " records · last successful refresh ",
          connection.lastSuccess.at,
        ],
      }),
      i.jsx("span", {
        children: hasLocalChanges
          ? "Local values changed; refresh cannot overwrite them."
          : connection.lastAttempt.code
            ? errorMessageOf(connection.lastAttempt.code)
            : "Refresh to check for newer source data.",
      }),
      i.jsx("button", {
        type: "button",
        onClick: onReview,
        children: i.jsx(LocalizedText, { children: "Connections & Refresh" }),
      }),
    ],
  });
}
function errorMessageOf(code) {
  return new ExternalSourceError(code).message;
}
function ExternalDataRefreshDialog(props) {
  const [isBatchRefreshOpen, setBatchRefreshOpen] = ne.useState(false),
    workbookMessages = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    propsRef = ne.useRef(props);
  propsRef.current = props;
  const isMounted = ne.useRef(true),
    abortControllerRef = ne.useRef(null),
    isRequestPending = ne.useRef(false),
    previousInputs = ne.useRef([
      props.workbook,
      props.readOnly,
      JSON.stringify(props.allowedSheetIds),
    ]),
    initialConnection = props.workbook.sheets.find(
      (sheet) => sheet.id === props.sheetId,
    )?.externalConnection,
    [selectedSheetId, setSelectedSheetId] = ne.useState(
      initialConnection ? props.sheetId : "",
    ),
    [selectedSourceId, setSelectedSourceId] = ne.useState(
      initialConnection?.source.id ?? "",
    ),
    [outputName, setOutputName] = ne.useState("External data"),
    [approvedSources, setApprovedSources] = ne.useState([]),
    [busyMessage, setBusyMessage] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [successMessage, setSuccessMessage] = ne.useState(""),
    [refreshPreview, setRefreshPreview] = ne.useState(null),
    [previewPage, setPreviewPage] = ne.useState(0),
    sheet = props.workbook.sheets.find((sheet) => sheet.id === selectedSheetId),
    selectedConnection = sheet?.externalConnection,
    applicablePlan =
      refreshPreview?.book === props.workbook &&
      refreshPreview.adapter === props.adapter &&
      !props.readOnly
        ? refreshPreview.plan
        : undefined,
    accessPolicy = {
      readOnly: props.readOnly,
      allowedSheetIds: props.allowedSheetIds,
    },
    canRefreshSheet =
      !sheet ||
      (!sheet.hidden &&
        !sheet.veryHidden &&
        !sheet.protection?.enabled &&
        (!props.allowedSheetIds || props.allowedSheetIds.includes(sheet.id))),
    isRefreshBlocked =
      !!props.readOnly || !canRefreshSheet || !props.adapter || !!busyMessage;
  (ne.useEffect(() => {
    isMounted.current = true;
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button")?.focus(),
      () => {
        ((isMounted.current = false),
          abortControllerRef.current?.abort(),
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      (abortControllerRef.current?.abort(),
        (abortControllerRef.current = null),
        (isRequestPending.current = false),
        setBusyMessage(""),
        setApprovedSources([]),
        setRefreshPreview(null),
        props.adapter && loadApprovedSources());
    }, [props.adapter]),
    ne.useEffect(() => {
      const currentInputs = [
        props.workbook,
        props.readOnly,
        JSON.stringify(props.allowedSheetIds),
      ];
      (currentInputs.some(
        (inputValue, inputIndex) =>
          inputValue !== previousInputs.current[inputIndex],
      ) &&
        (isRequestPending.current && abortControllerRef.current?.abort(),
        setRefreshPreview(null)),
        (previousInputs.current = currentInputs));
    }, [
      props.workbook,
      props.readOnly,
      JSON.stringify(props.allowedSheetIds),
    ]));
  function onClick() {
    (abortControllerRef.current?.abort(), props.onClose());
  }
  function resetRefreshState() {
    (setRefreshPreview(null),
      setPreviewPage(0),
      setErrorMessage(""),
      setSuccessMessage(""));
  }
  async function loadApprovedSources() {
    const adapter = propsRef.current.adapter;
    if (!adapter || isRequestPending.current) return;
    ((isRequestPending.current = true),
      setBusyMessage("Loading approved sources…"),
      setErrorMessage(""));
    const loadController = new AbortController();
    abortControllerRef.current = loadController;
    try {
      const sourceCatalog = await fetchSourceCatalog(adapter, {
        signal: loadController.signal,
        reliability: propsRef.current.reliability,
      });
      if (!isMounted.current || adapter !== propsRef.current.adapter) return;
      (setApprovedSources(sourceCatalog),
        setSelectedSourceId(
          (previousSourceId) =>
            previousSourceId || sourceCatalog[0]?.source.id || "",
        ));
    } catch (error) {
      isMounted.current &&
        setErrorMessage(isExternalSourceError(error).message);
    } finally {
      abortControllerRef.current === loadController &&
        ((abortControllerRef.current = null),
        (isRequestPending.current = false),
        isMounted.current && setBusyMessage(""));
    }
  }
  async function runSourceRefresh(isApplyRequested = false) {
    if (
      isRefreshBlocked ||
      isRequestPending.current ||
      !props.adapter ||
      !selectedSourceId ||
      (isApplyRequested && !applicablePlan)
    )
      return;
    const currentProps = propsRef.current,
      adapter = props.adapter,
      sourceId = selectedSourceId,
      sheetId = selectedSheetId;
    ((isRequestPending.current = true),
      setBusyMessage(
        isApplyRequested
          ? "Checking access and applying…"
          : "Refreshing preview…",
      ),
      setErrorMessage(""),
      setSuccessMessage(""));
    const refreshController = new AbortController();
    abortControllerRef.current = refreshController;
    try {
      if (isApplyRequested) {
        const accessGrant = await authorizeSourceAccess(adapter, sourceId, {
          signal: refreshController.signal,
          reliability: props.reliability,
        });
        if (!isMounted.current || refreshController.signal.aborted) return;
        if (
          propsRef.current.workbook !== currentProps.workbook ||
          propsRef.current.adapter !== adapter ||
          propsRef.current.readOnly
        )
          throw new ExternalSourceError("stale");
        const grantedWorkbook = grantSheetSnapshotSource(
          propsRef.current.workbook,
          applicablePlan,
          accessGrant,
          {
            readOnly: propsRef.current.readOnly,
            allowedSheetIds: propsRef.current.allowedSheetIds,
          },
        );
        (propsRef.current.onWorkbookChange(
          grantedWorkbook,
          "Refresh external source",
        ),
          setRefreshPreview(null),
          setSuccessMessage(
            "Source data refreshed. The reviewed population and control totals were applied together.",
          ));
      } else {
        const refreshResult = await refreshSourceSnapshot(adapter, sourceId, {
          signal: refreshController.signal,
          reliability: props.reliability,
        });
        if (!isMounted.current || refreshController.signal.aborted) return;
        if (
          propsRef.current.workbook !== currentProps.workbook ||
          propsRef.current.adapter !== adapter ||
          propsRef.current.readOnly
        )
          throw new ExternalSourceError("stale");
        const plan = EH(
          currentProps.workbook,
          refreshResult.snapshot,
          refreshResult.grant,
          { outputName: outputName, sheetId: sheetId || undefined },
          accessPolicy,
        );
        (setRefreshPreview({
          plan: plan,
          book: currentProps.workbook,
          adapter: adapter,
        }),
          setPreviewPage(0));
      }
    } catch (error) {
      if (!isMounted.current) return;
      const sourceError = isExternalSourceError(error);
      if (
        (setErrorMessage(sourceError.message),
        setRefreshPreview(null),
        sheetId &&
          propsRef.current.adapter === adapter &&
          propsRef.current.workbook === currentProps.workbook)
      ) {
        const latestProps = propsRef.current,
          externalSourceState = normalizeExternalSourceState(
            latestProps.workbook,
            sheetId,
            sourceId,
            sourceError.code,
            {
              readOnly: latestProps.readOnly,
              allowedSheetIds: latestProps.allowedSheetIds,
            },
          );
        if (externalSourceState !== latestProps.workbook)
          try {
            latestProps.onWorkbookChange(
              externalSourceState,
              "External source refresh failed; previous data retained",
            );
          } catch {
            setErrorMessage(
              "The host changed while refresh failed. Previous source data remains unchanged; reopen Connections & Refresh.",
            );
          }
      }
    } finally {
      abortControllerRef.current === refreshController &&
        ((abortControllerRef.current = null),
        (isRequestPending.current = false),
        isMounted.current && setBusyMessage(""));
    }
  }
  const connectedSheets = props.workbook.sheets.filter(
    (sheet) =>
      sheet.externalConnection &&
      !sheet.hidden &&
      !sheet.veryHidden &&
      (!props.allowedSheetIds || props.allowedSheetIds.includes(sheet.id)),
  );
  return isBatchRefreshOpen
    ? i.jsx(ExternalSourceBatchRefreshPanel, {
        ...props,
        onClose: () => setBatchRefreshOpen(false),
      })
    : i.jsx("div", {
        className: "lg-modal-backdrop",
        children: i.jsxs("section", {
          className: "lg-external-sources",
          ref: dialogRef,
          role: "dialog",
          "aria-modal": "true",
          "aria-labelledby": titleId,
          onKeyDown: (event) => {
            var lastFocusable, firstFocusable;
            if (
              (event.key === "Escape" && (event.stopPropagation(), onClick()),
              event.key === "Tab")
            ) {
              const focusableElements = [
                ...event.currentTarget.querySelectorAll(
                  'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
                ),
              ];
              event.shiftKey && document.activeElement === focusableElements[0]
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
                    i.jsx("p", {
                      className: "lg-source-eyebrow",
                      children: i.jsx(LocalizedText, {
                        children: "Get & Transform Data",
                      }),
                    }),
                    i.jsx("h2", {
                      id: titleId,
                      children: i.jsx(LocalizedText, {
                        children: "Connections & Refresh",
                      }),
                    }),
                    i.jsx("p", {
                      children:
                        "Review an approved source before replacing its last successful population.",
                    }),
                  ],
                }),
                i.jsx("button", {
                  type: "button",
                  "aria-label": workbookMessages.text(
                    "Close source connections",
                  ),
                  onClick: onClick,
                  children: i.jsx(LocalizedText, { children: "Close" }),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "lg-source-body",
              children: [
                !props.adapter &&
                  i.jsx("p", {
                    className: "lg-source-message",
                    children:
                      "Your application has not connected a source service. Ask the host developer to supply approved sources. Existing snapshots remain available.",
                  }),
                props.readOnly &&
                  i.jsx("p", {
                    className: "lg-source-message",
                    children:
                      "This workbook is read-only. Source history can be inspected; refresh cannot change it.",
                  }),
                i.jsxs("div", {
                  className: "lg-source-settings",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Connection" }),
                        i.jsxs("select", {
                          "aria-label": "Connection",
                          value: selectedSheetId,
                          disabled: !!busyMessage,
                          onChange: (event) => {
                            (resetRefreshState(),
                              setSelectedSheetId(event.target.value));
                            const selectedSheetConnection =
                              props.workbook.sheets.find(
                                (sheet) => sheet.id === event.target.value,
                              )?.externalConnection;
                            setSelectedSourceId(
                              selectedSheetConnection?.source.id ??
                                approvedSources[0]?.source.id ??
                                "",
                            );
                          },
                          children: [
                            i.jsx("option", {
                              value: "",
                              children: "New connection",
                            }),
                            connectedSheets.map((connectedSheet) =>
                              i.jsxs(
                                "option",
                                {
                                  value: connectedSheet.id,
                                  children: [
                                    connectedSheet.name,
                                    " · ",
                                    connectedSheet.externalConnection.source
                                      .label,
                                  ],
                                },
                                connectedSheet.id,
                              ),
                            ),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Approved source" }),
                        i.jsxs("select", {
                          "aria-label": "Approved source",
                          disabled:
                            !!busyMessage ||
                            !!selectedSheetId ||
                            !props.adapter,
                          value: selectedSourceId,
                          onChange: (event) => {
                            (resetRefreshState(),
                              setSelectedSourceId(event.target.value));
                          },
                          children: [
                            i.jsx("option", {
                              value: "",
                              children: "Choose a source",
                            }),
                            selectedSourceId &&
                              !approvedSources.some(
                                (approvedSource) =>
                                  approvedSource.source.id === selectedSourceId,
                              ) &&
                              i.jsxs("option", {
                                value: selectedSourceId,
                                children: [
                                  selectedConnection?.source.label ??
                                    selectedSourceId,
                                  " · approval unavailable",
                                ],
                              }),
                            approvedSources.map((approvedSource) =>
                              i.jsxs(
                                "option",
                                {
                                  value: approvedSource.source.id,
                                  children: [
                                    approvedSource.source.label,
                                    " · ",
                                    approvedSource.source.kind,
                                  ],
                                },
                                approvedSource.source.id,
                              ),
                            ),
                          ],
                        }),
                      ],
                    }),
                    !selectedSheetId &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Output worksheet",
                          }),
                          i.jsx("input", {
                            "aria-label": "Output worksheet",
                            value: outputName,
                            disabled: !!busyMessage,
                            onChange: (event) => {
                              (resetRefreshState(),
                                setOutputName(event.target.value));
                            },
                          }),
                        ],
                      }),
                    i.jsx("button", {
                      type: "button",
                      onClick: () => void loadApprovedSources(),
                      disabled: !props.adapter || !!busyMessage,
                      children: i.jsx(LocalizedText, {
                        children: "Reload approved sources",
                      }),
                    }),
                  ],
                }),
                i.jsx("p", {
                  className: "lg-source-help",
                  children:
                    "Only the host’s approved catalog is available here. Data-source credentials, endpoints and queries stay in its backend. Refresh is manual; opening a workbook never fetches source data.",
                }),
                i.jsx("button", {
                  type: "button",
                  disabled:
                    !!busyMessage ||
                    !!props.readOnly ||
                    !props.adapter ||
                    !connectedSheets.length,
                  onClick: () => {
                    (resetRefreshState(), setBatchRefreshOpen(true));
                  },
                  children: i.jsx(LocalizedText, { children: "Refresh All" }),
                }),
                selectedConnection &&
                  i.jsxs("section", {
                    className: "lg-source-history",
                    "aria-label": "Last successful source data",
                    children: [
                      i.jsx("h3", {
                        children:
                          selectedConnection.lastAttempt.outcome === "succeeded"
                            ? "Last successful snapshot"
                            : "Stale · previous successful data retained",
                      }),
                      i.jsxs("p", {
                        children: [
                          selectedConnection.source.label,
                          " · ",
                          selectedConnection.source.kind,
                          " · source ID",
                          " ",
                          selectedConnection.source.id,
                          " · schema ",
                          selectedConnection.source.schemaVersion,
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          selectedConnection.lastSuccess.rows,
                          " records · refreshed",
                          " ",
                          selectedConnection.lastSuccess.at,
                          " · source as of",
                          " ",
                          selectedConnection.lastSuccess.sourceTimestamp,
                          " · version",
                          " ",
                          selectedConnection.lastSuccess.version,
                        ],
                      }),
                      i.jsx("p", {
                        children:
                          selectedConnection.lastSuccess.controls
                            .map(
                              (control) =>
                                `${control.field}: ${control.sum} (${control.numeric} numbers; ${control.blanks} blanks)`,
                            )
                            .join(" · ") ||
                          "No numeric control fields declared.",
                      }),
                      selectedConnection.lastAttempt.code &&
                        i.jsxs("p", {
                          children: [
                            errorMessageOf(selectedConnection.lastAttempt.code),
                            " · attempted",
                            " ",
                            selectedConnection.lastAttempt.at,
                          ],
                        }),
                    ],
                  }),
                errorMessage &&
                  i.jsx("p", {
                    className: "lg-source-error",
                    role: "alert",
                    children: errorMessage,
                  }),
                successMessage &&
                  i.jsx("p", { role: "status", children: successMessage }),
                busyMessage &&
                  i.jsxs("p", {
                    role: "status",
                    children: [
                      busyMessage,
                      " ",
                      i.jsx("button", {
                        type: "button",
                        onClick: () => abortControllerRef.current?.abort(),
                        children: i.jsx(LocalizedText, {
                          children: "Cancel refresh",
                        }),
                      }),
                    ],
                  }),
                refreshPreview &&
                  !applicablePlan &&
                  i.jsx("p", {
                    className: "lg-source-message",
                    children:
                      "The workbook or access changed. Preview again before applying.",
                  }),
                applicablePlan &&
                  i.jsxs("section", {
                    className: "lg-source-preview",
                    "aria-label": "External source refresh preview",
                    children: [
                      i.jsxs("h3", {
                        children: [
                          "Refresh preview · ",
                          applicablePlan.snapshot.rows.length,
                          " records",
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          applicablePlan.snapshot.source.label,
                          " · ",
                          applicablePlan.snapshot.source.kind,
                          " ",
                          "· ",
                          applicablePlan.snapshot.source.id,
                          " · as of",
                          " ",
                          applicablePlan.snapshot.sourceTimestamp,
                          " · version",
                          " ",
                          applicablePlan.snapshot.version,
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          "Rows: ",
                          applicablePlan.previous?.rows ?? 0,
                          " →",
                          " ",
                          applicablePlan.snapshot.rows.length,
                          " · change",
                          " ",
                          applicablePlan.snapshot.rows.length -
                            (applicablePlan.previous?.rows ?? 0),
                        ],
                      }),
                      i.jsx("div", {
                        className: "lg-source-scroll",
                        children: i.jsxs("table", {
                          "aria-label": "Refresh control changes",
                          children: [
                            i.jsx("thead", {
                              children: i.jsxs("tr", {
                                children: [
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "Control field",
                                  }),
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "Previous total",
                                  }),
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "New total",
                                  }),
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "Change",
                                  }),
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "Numbers / blanks",
                                  }),
                                ],
                              }),
                            }),
                            i.jsx("tbody", {
                              children: applicablePlan.controls.map(
                                (control) => {
                                  const previousTotal =
                                    applicablePlan.previous?.controls.find(
                                      (previousControl) =>
                                        previousControl.field === control.field,
                                    )?.sum ?? 0;
                                  return i.jsxs(
                                    "tr",
                                    {
                                      children: [
                                        i.jsx("th", {
                                          scope: "row",
                                          children: control.field,
                                        }),
                                        i.jsx("td", {
                                          children: previousTotal,
                                        }),
                                        i.jsx("td", { children: control.sum }),
                                        i.jsx("td", {
                                          children: control.sum - previousTotal,
                                        }),
                                        i.jsxs("td", {
                                          children: [
                                            control.numeric,
                                            " / ",
                                            control.blanks,
                                          ],
                                        }),
                                      ],
                                    },
                                    control.field,
                                  );
                                },
                              ),
                            }),
                          ],
                        }),
                      }),
                      !applicablePlan.controls.length &&
                        i.jsx("p", {
                          children:
                            "No numeric control fields were declared by the host.",
                        }),
                      i.jsxs("p", {
                        children: [
                          "Schema:",
                          " ",
                          applicablePlan.snapshot.source.columns
                            .map(
                              (column) =>
                                `${column.name} (${column.type}${column.nullable ? ", nullable" : ""})`,
                            )
                            .join(" · "),
                        ],
                      }),
                      i.jsx("div", {
                        className: "lg-source-scroll",
                        tabIndex: 0,
                        "aria-label": "Source result rows",
                        children: i.jsxs("table", {
                          children: [
                            i.jsx("thead", {
                              children: i.jsxs("tr", {
                                children: [
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "Record",
                                  }),
                                  applicablePlan.snapshot.source.columns.map(
                                    (column) =>
                                      i.jsx(
                                        "th",
                                        { scope: "col", children: column.name },
                                        column.name,
                                      ),
                                  ),
                                ],
                              }),
                            }),
                            i.jsx("tbody", {
                              children: applicablePlan.snapshot.rows
                                .slice(previewPage * 25, (previewPage + 1) * 25)
                                .map((row, pageRowIndex) =>
                                  i.jsxs(
                                    "tr",
                                    {
                                      children: [
                                        i.jsx("th", {
                                          scope: "row",
                                          children:
                                            previewPage * 25 + pageRowIndex + 1,
                                        }),
                                        row.map((cellValue, columnIndex) =>
                                          i.jsx(
                                            "td",
                                            {
                                              children:
                                                cellValue === null
                                                  ? "(blank)"
                                                  : String(cellValue),
                                            },
                                            columnIndex,
                                          ),
                                        ),
                                      ],
                                    },
                                    pageRowIndex,
                                  ),
                                ),
                            }),
                          ],
                        }),
                      }),
                      i.jsxs("nav", {
                        "aria-label": "Source result pages",
                        children: [
                          i.jsx("button", {
                            type: "button",
                            disabled: previewPage === 0,
                            onClick: () =>
                              setPreviewPage((currentPage) => currentPage - 1),
                            children: "Previous records",
                          }),
                          i.jsxs("span", {
                            children: [
                              "Page ",
                              previewPage + 1,
                              " of",
                              " ",
                              Math.max(
                                1,
                                Math.ceil(
                                  applicablePlan.snapshot.rows.length / 25,
                                ),
                              ),
                              " · all records available",
                            ],
                          }),
                          i.jsx("button", {
                            type: "button",
                            disabled:
                              (previewPage + 1) * 25 >=
                              applicablePlan.snapshot.rows.length,
                            onClick: () =>
                              setPreviewPage((currentPage) => currentPage + 1),
                            children: "Next records",
                          }),
                        ],
                      }),
                    ],
                  }),
              ],
            }),
            i.jsxs("footer", {
              children: [
                i.jsx("p", {
                  children:
                    "Apply checks access again. Failed or cancelled refreshes retain prior data.",
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: isRefreshBlocked || !selectedSourceId,
                  onClick: () => void runSourceRefresh(),
                  children: i.jsx(LocalizedText, {
                    children: "Preview refresh",
                  }),
                }),
                i.jsx("button", {
                  type: "button",
                  className: "primary",
                  disabled: isRefreshBlocked || !applicablePlan,
                  onClick: () => void runSourceRefresh(true),
                  children: i.jsx(LocalizedText, {
                    children: "Apply reviewed refresh",
                  }),
                }),
              ],
            }),
          ],
        }),
      });
}
const workbookInstanceIds = new WeakMap();
let workbookInstanceCounter = 0;
function workbookInstanceKey(workbook) {
  return workbook
    ? (workbookInstanceIds.has(workbook) ||
        workbookInstanceIds.set(workbook, ++workbookInstanceCounter),
      workbookInstanceIds.get(workbook))
    : "default";
}
const isPromiseLike = (value) => !!value && typeof value.then == "function";
function chainMaybePromise(valueOrPromise, callback) {
  return isPromiseLike(valueOrPromise)
    ? valueOrPromise.then(callback)
    : callback(valueOrPromise);
}
const Vj = buildLocalDemoProposal();

function wy(startAddress, endAddress) {
  const startCell = _t(startAddress),
    endCell = _t(endAddress);
  return `${columnIndexToLetters(Math.min(startCell.col, endCell.col))}${Math.min(startCell.row, endCell.row) + 1}:${columnIndexToLetters(Math.max(startCell.col, endCell.col))}${Math.max(startCell.row, endCell.row) + 1}`;
}
const LazyAssistantChat = ne.lazy(() =>
    import("../assistant/AssistantChat.js").then((loadedModule) => ({
      default: loadedModule.AssistantChat,
    })),
  ),
  LazyAssistantChatMarkdown = ne.lazy(() =>
    import("../assistant/AssistantChatMarkdown.js").then((loadedModule) => ({
      default: loadedModule.AssistantChatMarkdown,
    })),
  );

export {
  isPromiseLike,
  ExternalSourceStatusPanel,
  workbookInstanceKey,
  chainMaybePromise,
  Vj,
  ExternalDataRefreshDialog,
  LazyAssistantChat,
  LazyAssistantChatMarkdown,
  wy,
};
