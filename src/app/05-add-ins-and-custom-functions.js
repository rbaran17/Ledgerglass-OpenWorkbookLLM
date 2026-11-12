import {
  formatInspectionRecords,
  recordSourceInspection,
  sourceInspectionRecords,
  validateIsolatedFormatInspection,
} from "./06c-package-and-worksheet-parsing.js";

import i from "react/jsx-runtime";
import ne from "react";
import { X } from "lucide-react";
import {
  fingerprintExtensionCall,
  extensionRuntimeLimits,
  hashExtensionManifest,
  fingerprintCanonicalValue,
  validateExtensionValue,
  normalizeDatePolicy,
} from "./01-schemas-date-functions-and-validation.js";
import {
  formatExternalSheetReference,
  rewriteFormulaReferences,
  evaluateWorkbookFormulas,
  encodeWorkbookFingerprint,
  parseExternalLinkRange,
  workbookQualifierSchema,
  tokenizeExternalReferences,
  assertSheetUnprotected,
  buildExternalLinkAliases,
  normalizeWorkbookInput,
  workbookLinkLimits,
  checkDataValidationBatch,
  rewriteExternalReferences,
} from "./03-workbook-model.js";
import {
  resolveExtensionCachedResult,
  evaluateExtensionFunctionCall,
  requireApprovedExtension,
  withExtensionHostScope,
  runExtensionOperation,
  applyExtensionPreview,
  resolveExtensionSheetScope,
  k$,
  validateExtensionHostPolicy,
  reviewExtensionFormulas,
  assertExtensionWriteScope,
  getFormulaScanList,
} from "./04-cell-input-paste-and-fill.js";

function buildExtensionCallReview(workbook, extensionId, options = {}) {
  const hostPolicy = validateExtensionHostPolicy(workbook, options),
    binding = requireApprovedExtension(workbook, extensionId, options),
    manifest = binding.manifest,
    hostScope = withExtensionHostScope(workbook, hostPolicy.extensions),
    requests = [],
    manifestHash = hashExtensionManifest(manifest);
  let inputCellCount = 0;
  for (const formulaCell of reviewExtensionFormulas(workbook))
    for (const call of formulaCell.calls.filter((call) =>
      call.name.startsWith(manifest.namespace + "."),
    )) {
      if (
        (assertExtensionWriteScope(workbook, [formulaCell.sheetId], options),
        formulaCell.literal)
      )
        throw Error(
          "Inert imported formulas must be deliberately entered as active formulas before extension execution.",
        );
      const callArguments = evaluateExtensionFunctionCall(
        hostScope,
        formulaCell,
        call,
        binding,
      );
      if (
        ((inputCellCount += callArguments.reduce(
          (total, argument) =>
            total +
            (Array.isArray(argument)
              ? argument.length * (argument[0]?.length ?? 0)
              : 1),
          0,
        )),
        inputCellCount > extensionRuntimeLimits.inputCells ||
          requests.length >= extensionRuntimeLimits.calls)
      )
        throw Error(
          "Extension argument or call count exceeds its review limit.",
        );
      const callFingerprint = fingerprintExtensionCall(
        formulaCell.sheetId,
        formulaCell.address,
        formulaCell.formula,
        call.name,
        callArguments,
        manifestHash,
      );
      requests.some((request) => request.cacheKey === callFingerprint) ||
        requests.push({
          sheetId: formulaCell.sheetId,
          address: formulaCell.address,
          formula: formulaCell.formula,
          cacheKey: callFingerprint,
          invocation: {
            requestId: callFingerprint,
            extensionId: extensionId,
            version: manifest.version,
            functionName: call.name,
            arguments: callArguments,
          },
        });
    }
  if (!requests.length)
    throw Error("No active formula calls use this extension.");
  const review = {
    extensionId: extensionId,
    manifest: structuredClone(manifest),
    requests: requests,
    disclosure:
      "Only the listed argument values, function identity, extension version and request ID will be sent to the host adapter. Confirm processing separately; results require another review before changing the workbook.",
  };
  return (
    k$.set(review, {
      bookHash: fingerprintCanonicalValue(workbook),
      policyHash: hostPolicy.hash(),
      planHash: fingerprintCanonicalValue(review),
      binding: binding,
      used: false,
    }),
    review
  );
}
const extensionResultCandidates = new WeakMap();
async function runReviewedExtensionCalls(
  workbook,
  review,
  options = {},
  runOptions = { approved: false },
) {
  const reviewRecord = k$.get(review),
    hostPolicy = validateExtensionHostPolicy(workbook, options);
  if (!reviewRecord || reviewRecord.used)
    throw Error("Unknown or already executed extension input review.");
  if (runOptions.approved !== true)
    throw Error(
      "Explicit approval is required before sending extension arguments.",
    );
  if (
    options.readOnly ||
    reviewRecord.bookHash !== fingerprintCanonicalValue(workbook) ||
    reviewRecord.policyHash !== hostPolicy.hash() ||
    reviewRecord.planHash !== fingerprintCanonicalValue(review)
  )
    throw Error(
      "Workbook, extension or permissions changed. Review inputs again.",
    );
  const timeoutMs = runOptions.timeoutMs ?? extensionRuntimeLimits.timeoutMs;
  if (
    !Number.isInteger(timeoutMs) ||
    timeoutMs < 1 ||
    timeoutMs > extensionRuntimeLimits.maxTimeoutMs
  )
    throw Error("Extension timeout must be between 1 and 60,000 milliseconds.");
  if (runOptions.signal?.aborted)
    throw Error("Extension processing cancelled.");
  reviewRecord.used = true;
  const abortController = new AbortController(),
    abortHandler = () => abortController.abort();
  runOptions.signal?.addEventListener("abort", abortHandler, { once: true });
  let timeoutHandle;
  try {
    const timeoutPromise = new Promise((resolve, reject) => {
        (abortController.signal.addEventListener(
          "abort",
          () => reject(Error("Extension processing cancelled or timed out.")),
          { once: true },
        ),
          (timeoutHandle = setTimeout(
            () => abortController.abort(),
            timeoutMs,
          )));
      }),
      results = [];
    let outputCellCount = 0;
    const invokeRequests = async () => {
      for (const request of review.requests) {
        if (
          reviewRecord.bookHash !== fingerprintCanonicalValue(workbook) ||
          reviewRecord.policyHash !==
            validateExtensionHostPolicy(workbook, options).hash()
        )
          throw Error(
            "Inputs or permissions changed during extension processing.",
          );
        if (abortController.signal.aborted)
          throw Error("Extension processing cancelled.");
        let response;
        try {
          response = await reviewRecord.binding.invoke(
            structuredClone(request.invocation),
            abortController.signal,
          );
        } catch {
          throw Error(
            "The extension adapter failed. No workbook changes were applied.",
          );
        }
        if (abortController.signal.aborted)
          throw Error("Extension processing cancelled.");
        if (
          !response ||
          Object.keys(response).sort().join("|") !== "revision|value" ||
          typeof response.revision != "string" ||
          !response.revision.trim() ||
          response.revision.length > 200 ||
          /[\u0000-\u001f\u007f]/.test(response.revision)
        )
          throw Error("The extension returned an invalid response envelope.");
        const functionDescriptor = review.manifest.functions.find(
            (manifestFunction) =>
              `${review.manifest.namespace}.${manifestFunction.id}` ===
              request.invocation.functionName,
          ),
          validatedValue = validateExtensionValue(
            response.value,
            functionDescriptor.result,
          );
        if (
          ((outputCellCount += Array.isArray(validatedValue)
            ? validatedValue.length * validatedValue[0].length
            : validatedValue && typeof validatedValue == "object"
              ? Object.keys(validatedValue.fields).length
              : 1),
          outputCellCount > extensionRuntimeLimits.outputCells)
        )
          throw Error("Extension results exceed the total output cell budget.");
        results.push({
          key: request.cacheKey,
          sheetId: request.sheetId,
          address: request.address,
          formula: request.formula,
          functionName: request.invocation.functionName,
          arguments: structuredClone(request.invocation.arguments),
          manifestHash: hashExtensionManifest(review.manifest),
          scopeHash: resolveExtensionSheetScope(workbook, reviewRecord.binding),
          result: validatedValue,
          calculatedAt: new Date().toISOString(),
          providerRevision: response.revision,
        });
      }
      return results;
    };
    if (
      (await Promise.race([invokeRequests(), timeoutPromise]),
      reviewRecord.bookHash !== fingerprintCanonicalValue(workbook) ||
        reviewRecord.policyHash !==
          validateExtensionHostPolicy(workbook, options).hash() ||
        reviewRecord.planHash !== fingerprintCanonicalValue(review))
    )
      throw Error("Inputs or permissions changed during extension processing.");
    const workbookInput = normalizeWorkbookInput(workbook),
      nextExtensions = structuredClone(workbookInput.extensions);
    ((nextExtensions.caches = nextExtensions.caches.filter(
      (cach) =>
        !results.some(
          (resultRecord) =>
            resultRecord.sheetId === cach.sheetId &&
            resultRecord.address === cach.address &&
            resultRecord.functionName === cach.functionName,
        ),
    )),
      nextExtensions.caches.push(...results),
      (workbookInput.extensions = nextExtensions));
    const candidateWorkbook = normalizeWorkbookInput(workbookInput),
      scopedWorkbook = withExtensionHostScope(
        candidateWorkbook,
        hostPolicy.extensions,
      ),
      evaluatedWorkbook = evaluateWorkbookFormulas(scopedWorkbook),
      preview = [],
      changedCells = [];
    for (const cachedResult of results) {
      const typedValue = evaluatedWorkbook.getTypedValue(
        cachedResult.sheetId,
        cachedResult.address,
      );
      if (typedValue.type === "error")
        throw Error(
          "The resulting formula has an error or obstructed array. No results were applied.",
        );
      if (
        preview.some(
          (previewEntry) =>
            previewEntry.sheetId === cachedResult.sheetId &&
            previewEntry.address === cachedResult.address,
        )
      )
        continue;
      preview.push({
        sheetId: cachedResult.sheetId,
        address: cachedResult.address,
        value: typedValue.value,
      });
      const cell = candidateWorkbook.sheets.find(
        (sheet) => sheet.id === cachedResult.sheetId,
      ).cells[cachedResult.address];
      changedCells.push({
        sheetId: cachedResult.sheetId,
        address: cachedResult.address,
        before: cell,
        after: cell,
      });
    }
    checkDataValidationBatch(scopedWorkbook, changedCells);
    const resultReview = {
      extensionId: review.extensionId,
      results: structuredClone(results),
      preview: preview,
      warning:
        "These are reviewed cached results, not a live service. Changed arguments require another approved run. Provider revision and entity source labels are host assertions, not proof of authenticity.",
    };
    return (
      extensionResultCandidates.set(resultReview, {
        bookHash: fingerprintCanonicalValue(workbook),
        policyHash: hostPolicy.hash(),
        planHash: fingerprintCanonicalValue(resultReview),
        candidate: candidateWorkbook,
      }),
      resultReview
    );
  } finally {
    (clearTimeout(timeoutHandle),
      runOptions.signal?.removeEventListener("abort", abortHandler),
      abortController.abort());
  }
}
function applyReviewedExtensionResults(workbook, resultReview, options = {}) {
  const candidateRecord = extensionResultCandidates.get(resultReview);
  if (!candidateRecord)
    throw Error("Unknown or already applied extension results.");
  if (
    options.readOnly ||
    candidateRecord.bookHash !== fingerprintCanonicalValue(workbook) ||
    candidateRecord.policyHash !==
      validateExtensionHostPolicy(workbook, options).hash() ||
    candidateRecord.planHash !== fingerprintCanonicalValue(resultReview)
  )
    throw Error(
      "Workbook, extension or permissions changed. Run and review again.",
    );
  return (
    extensionResultCandidates.delete(resultReview),
    normalizeWorkbookInput(candidateRecord.candidate)
  );
}
const emptyExtensionList = [];
function CustomFunctionReviewBanner({ workbook, extensions, onReview }) {
  const statusText = ne.useMemo(() => {
    try {
      const extensionCachedResult = resolveExtensionCachedResult(workbook, {
        extensions,
      });
      return extensionCachedResult.cells.length
        ? `${extensionCachedResult.cells.length} custom-function cells · ${extensionCachedResult.cells.filter((cell) => cell.functions.some((cellFunction) => cellFunction.state !== "reviewed-cache")).length} need review. Extension results refresh only after approval.`
        : null;
    } catch {
      return "Custom-function inventory needs review.";
    }
  }, [workbook, extensions]);
  return statusText
    ? i.jsxs("div", {
        className: "lg-workbook-link-status",
        role: "status",
        children: [
          i.jsx("span", { children: statusText }),
          i.jsx("button", {
            onClick: onReview,
            children: "Review custom functions",
          }),
        ],
      })
    : null;
}
function AddInsAndCustomFunctionsDialog(props) {
  const dialogRef = ne.useRef(null),
    latestPropsRef = ne.useRef(props),
    abortControllerRef = ne.useRef(null),
    requestIdRef = ne.useRef(0);
  latestPropsRef.current = props;
  const extensions = props.extensions ?? emptyExtensionList,
    [selectedExtensionId, setSelectedExtensionId] = ne.useState(""),
    [configurationPreview, setConfigurationPreview] = ne.useState(null),
    [inputReview, setInputReview] = ne.useState(null),
    [resultReview, setResultReview] = ne.useState(null),
    [isApproved, setIsApproved] = ne.useState(false),
    [disabled, setDisabled] = ne.useState(false),
    [actionError, setActionError] = ne.useState(""),
    [statusMessage, setStatusMessage] = ne.useState(""),
    [formulaPage, setFormulaPage] = ne.useState(0),
    inventoryState = ne.useMemo(() => {
      try {
        return {
          data: resolveExtensionCachedResult(props.workbook, {
            extensions: extensions,
          }),
          error: "",
        };
      } catch (error) {
        return { data: null, error: error.message };
      }
    }, [props.workbook, extensions]),
    manifestById = new Map(
      (inventoryState.data?.available ?? []).map((availableEntry) => [
        availableEntry.manifest.id,
        availableEntry.manifest,
      ]),
    );
  for (const installedEntry of inventoryState.data?.installed ?? [])
    manifestById.set(installedEntry.manifest.id, installedEntry.manifest);
  const extensionId = manifestById.has(selectedExtensionId)
      ? selectedExtensionId
      : (manifestById.keys().next().value ?? ""),
    selectedManifest = manifestById.get(extensionId),
    installedExtension = props.workbook.extensions?.installed.find(
      (installedRecord) => installedRecord.manifest.id === extensionId,
    ),
    buildOptions = () => ({
      extensions: latestPropsRef.current.extensions,
      readOnly: latestPropsRef.current.readOnly,
      allowedSheetIds: latestPropsRef.current.allowedSheetIds,
    }),
    onClick = () => {
      (requestIdRef.current++,
        abortControllerRef.current?.abort(),
        (abortControllerRef.current = null),
        setDisabled(false),
        setInputReview(null),
        setResultReview(null),
        setConfigurationPreview(null),
        setIsApproved(false));
    };
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.focus(),
      () => {
        (requestIdRef.current++,
          abortControllerRef.current?.abort(),
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      onClick();
    }, [props.workbook, extensions, props.readOnly, props.allowedSheetIds]));
  const runGuarded = (operation) => {
      (setActionError(""), setStatusMessage(""));
      try {
        operation();
      } catch (error) {
        setActionError(error.message);
      }
    },
    previewOperation = (action) =>
      runGuarded(() => {
        (onClick(),
          setConfigurationPreview(
            runExtensionOperation(
              props.workbook,
              extensionId,
              action,
              buildOptions(),
            ),
          ));
      }),
    reviewFunctionInputs = () =>
      runGuarded(() => {
        (onClick(),
          setInputReview(
            buildExtensionCallReview(
              props.workbook,
              extensionId,
              buildOptions(),
            ),
          ));
      }),
    runApprovedFunctions = async () => {
      if (!inputReview || !isApproved) return;
      const requestId = ++requestIdRef.current,
        abortController = new AbortController();
      ((abortControllerRef.current = abortController),
        setDisabled(true),
        setActionError(""));
      try {
        const reviewedResults = await runReviewedExtensionCalls(
          latestPropsRef.current.workbook,
          inputReview,
          buildOptions(),
          {
            approved: isApproved,
            signal: abortController.signal,
          },
        );
        requestId === requestIdRef.current &&
          (setResultReview(reviewedResults),
          setInputReview(null),
          setIsApproved(false));
      } catch (error) {
        requestId === requestIdRef.current &&
          (setActionError(error.message), setInputReview(null));
      } finally {
        requestId === requestIdRef.current &&
          (setDisabled(false), (abortControllerRef.current = null));
      }
    },
    applyReviewedChanges = () =>
      runGuarded(() => {
        const workbook = latestPropsRef.current.workbook,
          nextWorkbook = configurationPreview
            ? applyExtensionPreview(
                workbook,
                configurationPreview,
                buildOptions(),
              )
            : resultReview
              ? applyReviewedExtensionResults(
                  workbook,
                  resultReview,
                  buildOptions(),
                )
              : null;
        nextWorkbook &&
          (latestPropsRef.current.onWorkbookChange(
            nextWorkbook,
            configurationPreview
              ? `${configurationPreview.action} custom-function extension`
              : "Apply reviewed custom-function results",
          ),
          onClick(),
          setStatusMessage(
            "Reviewed changes applied. Raw formulas are preserved.",
          ));
      }),
    formulaCells = inventoryState.data?.cells ?? [],
    pageCount = Math.max(1, Math.ceil(formulaCells.length / 10)),
    currentPage = Math.min(formulaPage, pageCount - 1);
  return i.jsx("div", {
    className: "lg-reports-overlay",
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "lg-reports-dialog lg-workbook-links lg-extensions",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": "Add-ins and custom functions",
      tabIndex: -1,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.stopPropagation(), onClick(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusable = Array.from(
              dialogRef.current.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
              ),
            ),
            firstFocusable = focusable[0],
            lastFocusable = focusable.at(-1);
          event.shiftKey &&
          (document.activeElement === firstFocusable ||
            document.activeElement === dialogRef.current)
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              (document.activeElement === lastFocusable ||
                document.activeElement === dialogRef.current) &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("p", { children: "LEDGERGLASS · WORKBOOK EXTENSIONS" }),
                i.jsx("h2", { children: "Add-ins & custom functions" }),
              ],
            }),
            i.jsx("button", {
              "aria-label": "Close extensions",
              onClick: () => {
                (onClick(), props.onClose());
              },
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-reports-scroll",
          children: [
            i.jsxs("div", {
              className: "lg-links-summary",
              children: [
                i.jsx("strong", { children: "Manual, reviewed execution" }),
                i.jsxs("span", {
                  children: [
                    inventoryState.data?.installed.length ?? 0,
                    " installed extensions",
                  ],
                }),
                i.jsxs("span", {
                  children: [formulaCells.length, " formula cells"],
                }),
              ],
            }),
            i.jsx("p", {
              children:
                inventoryState.data?.policy ??
                "Only host-approved adapters can process reviewed inputs.",
            }),
            i.jsx("p", {
              children:
                "Office add-in, Python and R integrations use a host adapter. This component does not run those environments or install workbook-supplied code.",
            }),
            props.readOnly &&
              i.jsx("p", {
                className: "lg-links-warning",
                children:
                  "Read-only: inspection is available; configuration and execution are disabled.",
              }),
            (actionError || inventoryState.error) &&
              i.jsx("p", {
                role: "alert",
                className: "lg-links-error",
                children: actionError || inventoryState.error,
              }),
            statusMessage &&
              i.jsx("p", { role: "status", children: statusMessage }),
            i.jsxs("section", {
              className: "lg-links-card",
              "aria-label": "Extension manifest",
              children: [
                i.jsx("h3", { children: "1. Choose a sanctioned extension" }),
                manifestById.size
                  ? i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("label", {
                          children: [
                            "Extension",
                            i.jsx("select", {
                              value: extensionId,
                              disabled: disabled,
                              onChange: (event) => {
                                (onClick(),
                                  setSelectedExtensionId(event.target.value));
                              },
                              children: [...manifestById].map(
                                ([manifestId, manifest]) =>
                                  i.jsxs(
                                    "option",
                                    {
                                      value: manifestId,
                                      children: [
                                        manifest.label,
                                        " · ",
                                        manifest.version,
                                      ],
                                    },
                                    manifestId,
                                  ),
                              ),
                            }),
                          ],
                        }),
                        selectedManifest &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("dl", {
                                children: [
                                  i.jsx("dt", {
                                    children: "Namespace / version",
                                  }),
                                  i.jsxs("dd", {
                                    children: [
                                      selectedManifest.namespace,
                                      " / ",
                                      selectedManifest.version,
                                    ],
                                  }),
                                  i.jsx("dt", { children: "Runtime" }),
                                  i.jsx("dd", {
                                    children: selectedManifest.runtime,
                                  }),
                                  i.jsx("dt", { children: "Permissions" }),
                                  i.jsx("dd", {
                                    children:
                                      selectedManifest.permissions.join(", "),
                                  }),
                                  i.jsx("dt", { children: "Recalculation" }),
                                  i.jsx("dd", {
                                    children:
                                      "Manual · input values invalidate caches · nonvolatile",
                                  }),
                                  i.jsx("dt", { children: "State" }),
                                  i.jsx("dd", {
                                    children: installedExtension
                                      ? installedExtension.enabled
                                        ? "Enabled"
                                        : "Disabled"
                                      : "Not installed",
                                  }),
                                ],
                              }),
                              i.jsxs("details", {
                                children: [
                                  i.jsx("summary", {
                                    children: "Function signatures and schemas",
                                  }),
                                  i.jsx("pre", {
                                    children: JSON.stringify(
                                      selectedManifest.functions,
                                      null,
                                      2,
                                    ),
                                  }),
                                ],
                              }),
                              i.jsxs("div", {
                                className: "lg-links-actions",
                                children: [
                                  i.jsx("button", {
                                    disabled: props.readOnly || disabled,
                                    onClick: () => previewOperation("install"),
                                    children: installedExtension
                                      ? "Preview install / update"
                                      : "Preview installation",
                                  }),
                                  installedExtension &&
                                    i.jsxs(i.Fragment, {
                                      children: [
                                        i.jsxs("button", {
                                          disabled: props.readOnly || disabled,
                                          onClick: () =>
                                            previewOperation(
                                              installedExtension.enabled
                                                ? "disable"
                                                : "enable",
                                            ),
                                          children: [
                                            "Preview ",
                                            installedExtension.enabled
                                              ? "disable"
                                              : "enable",
                                          ],
                                        }),
                                        i.jsx("button", {
                                          disabled: props.readOnly || disabled,
                                          onClick: () =>
                                            previewOperation("remove"),
                                          children: "Preview removal",
                                        }),
                                      ],
                                    }),
                                ],
                              }),
                            ],
                          }),
                      ],
                    })
                  : i.jsx("p", {
                      children:
                        "No host adapters are configured. Ask your application developer to register an approved extension. Existing formulas stay intact.",
                    }),
              ],
            }),
            configurationPreview &&
              i.jsxs("section", {
                className: "lg-links-confirm",
                "aria-label": "Extension configuration preview",
                children: [
                  i.jsxs("h3", {
                    children: ["Review ", configurationPreview.action],
                  }),
                  i.jsxs("p", {
                    children: [
                      configurationPreview.manifest.label,
                      " · version ",
                      configurationPreview.manifest.version,
                    ],
                  }),
                  i.jsx("p", { children: configurationPreview.warning }),
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: props.readOnly,
                    onClick: applyReviewedChanges,
                    children: "Apply reviewed configuration",
                  }),
                  i.jsx("button", {
                    onClick: onClick,
                    children: "Cancel extension preview",
                  }),
                ],
              }),
            i.jsxs("section", {
              className: "lg-links-card",
              children: [
                i.jsx("h3", {
                  children: "2. Review inputs and request results",
                }),
                i.jsx("p", {
                  children:
                    "Review every argument before allowing this host adapter to process it. Source workbooks and credentials are not included in the request.",
                }),
                i.jsx("button", {
                  disabled:
                    !installedExtension?.enabled ||
                    props.readOnly ||
                    disabled ||
                    !!inventoryState.error,
                  onClick: reviewFunctionInputs,
                  children: "Review function inputs",
                }),
              ],
            }),
            inputReview &&
              i.jsxs("section", {
                className: "lg-links-confirm",
                "aria-label": "Extension input review",
                children: [
                  i.jsxs("h3", {
                    children: [
                      "Inputs for ",
                      inputReview.requests.length,
                      " calls",
                    ],
                  }),
                  i.jsx("p", { children: inputReview.disclosure }),
                  i.jsx("div", {
                    className: "lg-links-table",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", { children: "Cell" }),
                              i.jsx("th", { children: "Function" }),
                              i.jsx("th", { children: "Exact arguments" }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: inputReview.requests.map((request) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("td", {
                                    children: [
                                      props.workbook.sheets.find(
                                        (sheet) => sheet.id === request.sheetId,
                                      )?.name,
                                      "!",
                                      request.address,
                                    ],
                                  }),
                                  i.jsx("td", {
                                    children: request.invocation.functionName,
                                  }),
                                  i.jsx("td", {
                                    children: i.jsx("pre", {
                                      children: JSON.stringify(
                                        request.invocation.arguments,
                                        null,
                                        2,
                                      ),
                                    }),
                                  }),
                                ],
                              },
                              request.cacheKey,
                            ),
                          ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("label", {
                    className: "lg-links-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        disabled: disabled,
                        checked: isApproved,
                        onChange: (event) =>
                          setIsApproved(event.target.checked),
                      }),
                      "I approve sending these argument values to this host adapter.",
                    ],
                  }),
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: !isApproved || disabled || props.readOnly,
                    onClick: () => void runApprovedFunctions(),
                    children: "Run approved functions",
                  }),
                  i.jsx("button", {
                    onClick: onClick,
                    children: disabled
                      ? "Cancel processing"
                      : "Cancel extension preview",
                  }),
                ],
              }),
            disabled &&
              i.jsx("p", {
                role: "status",
                children: "Processing approved inputs…",
              }),
            resultReview &&
              i.jsxs("section", {
                className: "lg-links-confirm",
                "aria-label": "Extension result review",
                children: [
                  i.jsx("h3", { children: "3. Review and apply results" }),
                  i.jsx("p", { children: resultReview.warning }),
                  i.jsx("div", {
                    className: "lg-links-table",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", { children: "Cell / function" }),
                              i.jsx("th", {
                                children: "Returned value / linked fields",
                              }),
                              i.jsx("th", { children: "Provider revision" }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: resultReview.results.map((result) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("td", {
                                    children: [
                                      result.address,
                                      " · ",
                                      result.functionName,
                                    ],
                                  }),
                                  i.jsx("td", {
                                    children: i.jsx("pre", {
                                      children: JSON.stringify(
                                        result.result,
                                        null,
                                        2,
                                      ),
                                    }),
                                  }),
                                  i.jsx("td", {
                                    children: result.providerRevision,
                                  }),
                                ],
                              },
                              result.key,
                            ),
                          ),
                        }),
                      ],
                    }),
                  }),
                  i.jsx("h4", { children: "Calculated formula results" }),
                  i.jsx("ul", {
                    children: resultReview.preview.map((previewEntry) =>
                      i.jsxs(
                        "li",
                        {
                          children: [
                            previewEntry.address,
                            ": ",
                            String(previewEntry.value),
                          ],
                        },
                        previewEntry.sheetId + previewEntry.address,
                      ),
                    ),
                  }),
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: props.readOnly,
                    onClick: applyReviewedChanges,
                    children: "Apply reviewed results",
                  }),
                  i.jsx("button", {
                    onClick: onClick,
                    children: "Cancel extension preview",
                  }),
                ],
              }),
            i.jsxs("section", {
              className: "lg-links-card",
              "aria-label": "Custom function inventory",
              children: [
                i.jsx("h3", { children: "Formula and cache inventory" }),
                formulaCells.length
                  ? i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("div", {
                          className: "lg-links-table",
                          children: i.jsxs("table", {
                            children: [
                              i.jsx("thead", {
                                children: i.jsxs("tr", {
                                  children: [
                                    i.jsx("th", {
                                      children: "Cell / raw formula",
                                    }),
                                    i.jsx("th", {
                                      children: "Availability / current value",
                                    }),
                                    i.jsx("th", {
                                      children: "Retained evidence",
                                    }),
                                  ],
                                }),
                              }),
                              i.jsx("tbody", {
                                children: formulaCells
                                  .slice(
                                    currentPage * 10,
                                    currentPage * 10 + 10,
                                  )
                                  .map((formulaCell) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [
                                              i.jsxs("strong", {
                                                children: [
                                                  formulaCell.sheetName,
                                                  "!",
                                                  formulaCell.address,
                                                ],
                                              }),
                                              i.jsx("pre", {
                                                children: formulaCell.formula,
                                              }),
                                              formulaCell.literal &&
                                                "Inert imported text",
                                            ],
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              formulaCell.functions.map(
                                                (cellFunction) =>
                                                  i.jsxs(
                                                    "div",
                                                    {
                                                      children: [
                                                        cellFunction.name,
                                                        ": ",
                                                        cellFunction.state,
                                                      ],
                                                    },
                                                    cellFunction.name,
                                                  ),
                                              ),
                                              i.jsx("strong", {
                                                children: String(
                                                  formulaCell.result.value,
                                                ),
                                              }),
                                            ],
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              formulaCell.cacheLabel,
                                              formulaCell.caches.map((cach) =>
                                                i.jsxs(
                                                  "details",
                                                  {
                                                    children: [
                                                      i.jsxs("summary", {
                                                        children: [
                                                          cach.functionName,
                                                          " ·",
                                                          " ",
                                                          cach.providerRevision,
                                                          " ·",
                                                          " ",
                                                          cach.calculatedAt,
                                                        ],
                                                      }),
                                                      i.jsx("pre", {
                                                        children:
                                                          JSON.stringify(
                                                            cach.result,
                                                            null,
                                                            2,
                                                          ),
                                                      }),
                                                    ],
                                                  },
                                                  cach.key,
                                                ),
                                              ),
                                              formulaCell.importedCache &&
                                                i.jsxs("details", {
                                                  children: [
                                                    i.jsx("summary", {
                                                      children:
                                                        "Historical imported result — unverified",
                                                    }),
                                                    i.jsx("pre", {
                                                      children: JSON.stringify(
                                                        formulaCell.importedCache,
                                                        null,
                                                        2,
                                                      ),
                                                    }),
                                                  ],
                                                }),
                                            ],
                                          }),
                                        ],
                                      },
                                      formulaCell.sheetId + formulaCell.address,
                                    ),
                                  ),
                              }),
                            ],
                          }),
                        }),
                        i.jsxs("div", {
                          className: "lg-links-actions",
                          children: [
                            i.jsx("button", {
                              disabled: currentPage === 0,
                              onClick: () => setFormulaPage(currentPage - 1),
                              children: "Previous formula page",
                            }),
                            i.jsxs("span", {
                              children: [
                                "Page ",
                                currentPage + 1,
                                " of ",
                                pageCount,
                              ],
                            }),
                            i.jsx("button", {
                              disabled: currentPage + 1 === pageCount,
                              onClick: () => setFormulaPage(currentPage + 1),
                              children: "Next formula page",
                            }),
                          ],
                        }),
                      ],
                    })
                  : i.jsx("p", { children: "No custom function calls found." }),
              ],
            }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children:
                "Native JSON retains formulas and evidence. Host execution permissions are session-only.",
            }),
            i.jsx("button", {
              onClick: () => {
                (onClick(), props.onClose());
              },
              children: "Done",
            }),
          ],
        }),
      ],
    }),
  });
}
function buildWorkbookLinkPolicy(workbook, options) {
  if (options.readOnly !== undefined && typeof options.readOnly != "boolean")
    throw Error("Read-only policy must be boolean.");
  const sources = options.sources ?? [],
    allowed =
      options.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id);
  if (
    !Array.isArray(sources) ||
    sources.length > workbookLinkLimits.sources ||
    new Set(sources.map((event) => event.target)).size !== sources.length
  )
    throw Error("Approve distinct source qualifiers, at most eight workbooks.");
  for (const source of sources)
    if (
      (workbookQualifierSchema.parse(source.target),
      source.workbook.id === workbook.id ||
        !Array.isArray(source.allowedSheetIds) ||
        source.allowedSheetIds.some(
          (allowedSheetId) =>
            !source.workbook.sheets.some(
              (sheet) => sheet.id === allowedSheetId,
            ),
        ) ||
        (source.includeHidden !== undefined &&
          typeof source.includeHidden != "boolean"))
    )
      throw Error("Invalid approved workbook identity or worksheet scope.");
  if (
    !Array.isArray(allowed) ||
    allowed.some(
      (sheetId) => !workbook.sheets.some((sheet) => sheet.id === sheetId),
    )
  )
    throw Error("Invalid destination worksheet scope.");
  return {
    sources: sources,
    allowed: allowed,
    hash: () =>
      fingerprintCanonicalValue([
        !!options.readOnly,
        allowed,
        sources.map((source) => [
          source.target,
          encodeWorkbookFingerprint(source.workbook),
          source.allowedSheetIds,
          !!source.includeHidden,
        ]),
      ]),
  };
}
function collectWorkbookLinkFormulas(workbookInput) {
  const linkFormulas = [];
  for (const sheet of workbookInput.sheets)
    for (const [address, cell] of getFormulaScanList(sheet.cells, "external")) {
      if (typeof cell.raw != "string" || !cell.raw.startsWith("=")) continue;
      const references = tokenizeExternalReferences(cell.raw),
        zeroedFormula = rewriteExternalReferences(cell.raw, () => "0"),
        rewrittenFormula = rewriteFormulaReferences(
          zeroedFormula,
          workbookInput,
          sheet.id,
          address,
          false,
        );
      if (
        ((references.length || rewrittenFormula?.kind === "external") &&
          linkFormulas.push({
            sheetId: sheet.id,
            sheetName: sheet.name,
            address: address,
            formula: cell.raw,
            literal: !!cell.literal,
            references: references,
            ...(rewrittenFormula?.kind === "external"
              ? {
                  unsupported:
                    "Unsupported external syntax. Static A1 cells and rectangles only; DDE, whole axes, external names and 3-D links remain blocked.",
                }
              : {}),
          }),
        linkFormulas.length > workbookLinkLimits.formulas)
      )
        throw Error(
          `Workbook link review is limited to ${workbookLinkLimits.formulas} formula cells.`,
        );
    }
  return linkFormulas;
}
function buildWorkbookLinkReview(workbook, options = {}) {
  const workbookLinkPolicy = buildWorkbookLinkPolicy(workbook, options),
    workbookLinkFormulas = collectWorkbookLinkFormulas(workbook);
  return {
    links: [
      ...new Set(
        workbookLinkFormulas.flatMap((workbookLinkFormula) =>
          workbookLinkFormula.references.map((reference) => reference.target),
        ),
      ),
    ].map((target) => {
      const source = workbookLinkPolicy.sources.find(
          (source) => source.target === target,
        ),
        cachedSource = workbook.workbookLinks?.sources.find(
          (source) => source.target === target,
        ),
        references = workbookLinkFormulas.flatMap((workbookLinkFormula) =>
          workbookLinkFormula.references
            .filter((reference) => reference.target === target)
            .map((matchedReference) => ({
              ...matchedReference,
              sheetId: workbookLinkFormula.sheetId,
              address: workbookLinkFormula.address,
            })),
        ),
        hasMissingWorksheet =
          source &&
          references.some(
            (reference) =>
              !source.workbook.sheets.some(
                (sheet) =>
                  sheet.name.toLowerCase() === reference.sheet.toLowerCase(),
              ),
          ),
        hasUnauthorizedSheet =
          source &&
          references.some((reference) => {
            const sheet = source.workbook.sheets.find(
              (sheet) =>
                sheet.name.toLowerCase() === reference.sheet.toLowerCase(),
            );
            return (
              sheet &&
              (!source.allowedSheetIds.includes(sheet.id) ||
                ((sheet.hidden || sheet.veryHidden) && !source.includeHidden))
            );
          }),
        linkStatus = source
          ? hasMissingWorksheet
            ? "missing-worksheet"
            : hasUnauthorizedSheet
              ? "unauthorized-worksheet"
              : cachedSource
                ? cachedSource.sourceFingerprint !==
                  encodeWorkbookFingerprint(source.workbook)
                  ? "stale"
                  : references.some(
                        (reference) =>
                          !cachedSource.sheets.find(
                            (sheet) =>
                              sheet.name.toLowerCase() ===
                                reference.sheet.toLowerCase() &&
                              parseExternalLinkRange(reference.range).every(
                                (cellAddress) =>
                                  Object.hasOwn(sheet.cells, cellAddress),
                              ),
                          ),
                      )
                    ? "incomplete-cache"
                    : "matches-approved-snapshot"
                : "unresolved"
          : "unapproved-or-missing";
      return {
        target: target,
        packageTarget: workbook.workbookLinks?.packageTargets?.find(
          (packageTarget) => packageTarget.key === target,
        )?.target,
        status: linkStatus,
        references: references,
        lastResolvedRevision: cachedSource?.sourceRevision,
        lastResolvedAt: cachedSource?.resolvedAt,
        sourceName: cachedSource?.sourceName,
        sourceFingerprint: cachedSource?.sourceFingerprint,
        cached: !!cachedSource,
      };
    }),
    cells: workbookLinkFormulas,
    unsupported: workbookLinkFormulas.filter(
      (workbookLinkFormula) => workbookLinkFormula.unsupported,
    ),
    cachePolicy:
      "Values are retained snapshots. Only explicit refresh reads approved workbooks. No path or URL is fetched; matching a supplied revision does not verify a live file.",
  };
}
const workbookLinkPreviewRecords = new WeakMap();
function resolveLinkedSourceCalculation(workbook, references, target, options) {
  const source = buildWorkbookLinkPolicy(workbook, options).sources.find(
    (source) => source.target === target,
  );
  if (!source)
    throw Error(
      `Source ${target} is missing or not approved. Attach the exact source and approve its worksheets.`,
    );
  const workbookInput = normalizeWorkbookInput(source.workbook),
    allowedSheetIdSet = new Set(source.allowedSheetIds);
  if (
    normalizeDatePolicy(workbookInput.datePolicy).dateSystem !==
    normalizeDatePolicy(workbook.datePolicy).dateSystem
  )
    throw Error(
      "Source and destination date systems differ. Convert and review the source date system before linking.",
    );
  const allowRead = (sheetId) => {
      const sheet = workbookInput.sheets.find((sheet) => sheet.id === sheetId);
      return (
        !!sheet &&
        allowedSheetIdSet.has(sheetId) &&
        (!(sheet.hidden || sheet.veryHidden) || !!source.includeHidden)
      );
    },
    evaluator = evaluateWorkbookFormulas(
      workbookInput,
      {},
      { allowRead: allowRead, allowName: () => true, step: () => {} },
    ),
    sheets = [];
  let cellCount = 0;
  for (const reference of references) {
    const sourceSheet = workbookInput.sheets.find(
      (sheet) => sheet.name.toLowerCase() === reference.sheet.toLowerCase(),
    );
    if (!sourceSheet)
      throw Error(`Source worksheet ${reference.sheet} is missing.`);
    if (!allowRead(sourceSheet.id))
      throw Error(
        `Source worksheet ${reference.sheet} is outside the approved scope.`,
      );
    let sheet = sheets.find((sheet) => sheet.id === sourceSheet.id);
    sheet ||
      ((sheet = { id: sourceSheet.id, name: sourceSheet.name, cells: {} }),
      sheets.push(sheet));
    for (const address of parseExternalLinkRange(reference.range))
      if (!Object.hasOwn(sheet.cells, address)) {
        if (++cellCount > workbookLinkLimits.cells)
          throw Error("External link refresh exceeds 5,000 source cells.");
        const typedValue = evaluator.getTypedValue(sourceSheet.id, address);
        if (
          typedValue.type === "error" &&
          [
            "#BLOCKED!",
            "#LIMIT!",
            "#PENDING!",
            "#TIMEOUT!",
            "#CYCLE!",
          ].includes(String(typedValue.value))
        )
          throw Error(
            "Source calculation is unavailable, unauthorized, cyclic or over its resource limit.",
          );
        sheet.cells[address] = typedValue;
      }
  }
  if (evaluator.getOperationalErrors().length)
    throw Error(
      "A source calculation attempted an unavailable or unauthorized dependency.",
    );
  return {
    target: target,
    sourceId: workbookInput.id,
    sourceName: workbookInput.name,
    sourceRevision: workbookInput.revision,
    sourceFingerprint: encodeWorkbookFingerprint(workbookInput),
    resolvedAt: new Date().toISOString(),
    sheets: sheets,
  };
}
function assertLinkDestinationPublishable(evaluator, destination) {
  const typedValue = evaluator.getTypedValue(
    destination.sheetId,
    destination.address,
  );
  if (
    typedValue.type === "error" ||
    evaluator.getSpillRange(destination.sheetId, destination.address)
  )
    throw Error(
      `Cannot publish ${destination.sheetName}!${destination.address}: resolve its error or array result first.`,
    );
  if (evaluator.getOperationalErrors().length)
    throw Error(
      "The destination calculation includes an unresolved or unavailable dependency.",
    );
  return typedValue;
}
function previewWorkbookLinkOperation(workbook, operation, options = {}) {
  const workbookLinkPolicy = buildWorkbookLinkPolicy(workbook, options);
  if (options.readOnly) throw Error("The workbook is read-only.");
  if (!["refresh", "replace", "break"].includes(operation.action))
    throw Error("Unknown link operation.");
  if (
    (operation.target !== undefined &&
      workbookQualifierSchema.parse(operation.target),
    operation.action !== "refresh" && !operation.target)
  )
    throw Error("Choose a target workbook link.");
  if (
    operation.action === "replace" &&
    (workbookQualifierSchema.parse(operation.replacementTarget),
    operation.replacementTarget === operation.target)
  )
    throw Error("Choose a different source qualifier.");
  const workbookLinkFormulas = collectWorkbookLinkFormulas(workbook),
    matchingFormulas = workbookLinkFormulas.filter((workbookLinkFormula) =>
      operation.target
        ? workbookLinkFormula.references.some(
            (reference) => reference.target === operation.target,
          )
        : workbookLinkFormula.references.length ||
          workbookLinkFormula.unsupported,
    );
  if (!matchingFormulas.length) throw Error("No matching workbook link cells.");
  for (const linkFormula of matchingFormulas) {
    if (linkFormula.unsupported) throw Error(linkFormula.unsupported);
    if (!workbookLinkPolicy.allowed.includes(linkFormula.sheetId))
      throw Error("A destination worksheet is outside the host write scope.");
    if (
      (assertSheetUnprotected(
        workbook.sheets.find((sheet) => sheet.id === linkFormula.sheetId),
      ),
      linkFormula.literal && !operation.activateLiteral)
    )
      throw Error(
        "Explicitly approve activation of inert imported formulas before resolving them.",
      );
  }
  let workbookInput = normalizeWorkbookInput(workbook);
  workbookInput.workbookLinks = structuredClone(
    workbookInput.workbookLinks ?? { schemaVersion: 1, sources: [] },
  );
  for (const formulaToRewrite of matchingFormulas) {
    const sheet = workbookInput.sheets.find(
      (sheet) => sheet.id === formulaToRewrite.sheetId,
    );
    ((sheet.cells[formulaToRewrite.address] = {
      ...sheet.cells[formulaToRewrite.address],
      raw:
        operation.action === "replace"
          ? rewriteExternalReferences(
              formulaToRewrite.formula,
              (externalReference) =>
                externalReference.target === operation.target
                  ? formatExternalSheetReference(
                      operation.replacementTarget,
                      externalReference.sheet,
                      externalReference.range,
                    )
                  : undefined,
            )
          : formulaToRewrite.formula,
    }),
      delete sheet.cells[formulaToRewrite.address].literal);
  }
  if (operation.action !== "break") {
    const changedReferences = collectWorkbookLinkFormulas(workbookInput)
        .filter((collectedFormula) =>
          matchingFormulas.some(
            (matchingFormula) =>
              matchingFormula.sheetId === collectedFormula.sheetId &&
              matchingFormula.address === collectedFormula.address,
          ),
        )
        .flatMap((affectedFormula) => affectedFormula.references),
      uniqueTargets = [
        ...new Set(changedReferences.map((event) => event.target)),
      ];
    for (const target of uniqueTargets) {
      const targetReferences = collectWorkbookLinkFormulas(workbookInput)
          .flatMap((candidateFormula) => candidateFormula.references)
          .filter((event) => event.target === target),
        linkedSourceCalculation = resolveLinkedSourceCalculation(
          workbook,
          targetReferences,
          target,
          options,
        );
      ((workbookInput.workbookLinks.sources =
        workbookInput.workbookLinks.sources.filter(
          (source) => source.target !== target,
        )),
        workbookInput.workbookLinks.sources.push(linkedSourceCalculation));
    }
  } else if (!operation.acknowledgeCached)
    throw Error(
      "Acknowledge that breaking links keeps the previewed cached values and removes the complete formulas.",
    );
  workbookInput = normalizeWorkbookInput(workbookInput);
  const readableSheetIds = new Set([
      ...workbookLinkPolicy.allowed,
      ...buildExternalLinkAliases(workbookInput).sheets.map(
        (sheet) => sheet.id,
      ),
    ]),
    evaluator = evaluateWorkbookFormulas(
      workbookInput,
      {},
      { allowRead: (sheetId) => readableSheetIds.has(sheetId), step: () => {} },
    ),
    cells = matchingFormulas.map((matchedFormula) => {
      const typedValue = assertLinkDestinationPublishable(
        evaluator,
        matchedFormula,
      );
      return {
        sheetId: matchedFormula.sheetId,
        sheetName: matchedFormula.sheetName,
        address: matchedFormula.address,
        before: matchedFormula.formula,
        after:
          operation.action === "break"
            ? typedValue.value
            : workbookInput.sheets.find(
                (sheet) => sheet.id === matchedFormula.sheetId,
              ).cells[matchedFormula.address].raw,
        result: typedValue,
      };
    }),
    cellChanges = [];
  for (const cell of cells) {
    const sheet = workbookInput.sheets.find(
        (sheet) => sheet.id === cell.sheetId,
      ),
      before = workbook.sheets.find((sheet) => sheet.id === cell.sheetId).cells[
        cell.address
      ];
    (operation.action === "break" &&
      (sheet.cells[cell.address] = {
        ...sheet.cells[cell.address],
        raw: cell.result.value,
        literal: typeof cell.result.value == "string",
      }),
      cellChanges.push({
        sheetId: sheet.id,
        address: cell.address,
        before: before,
        after: sheet.cells[cell.address],
      }));
  }
  const referencedTargets = new Set(
    collectWorkbookLinkFormulas(workbookInput).flatMap((workbookFormula) =>
      workbookFormula.references.map((reference) => reference.target),
    ),
  );
  ((workbookInput.workbookLinks = {
    ...workbookInput.workbookLinks,
    sources: workbookInput.workbookLinks.sources.filter((source) =>
      referencedTargets.has(source.target),
    ),
  }),
    checkDataValidationBatch(workbookInput, cellChanges),
    (workbookInput = normalizeWorkbookInput(workbookInput)));
  const preview = {
    request: structuredClone(operation),
    cells: cells,
    sources: structuredClone(workbookInput.workbookLinks.sources),
    warning:
      operation.action === "break"
        ? "Complete formulas are replaced with these scalar values, including any other references in those formulas. Undo can restore the prior workbook. Retained original files are unchanged."
        : "These are reviewed snapshots of approved workbooks, not live network connections. Future changes require another preview and refresh.",
  };
  return (
    workbookLinkPreviewRecords.set(preview, {
      bookHash: fingerprintCanonicalValue(workbook),
      contextHash: workbookLinkPolicy.hash(),
      planHash: fingerprintCanonicalValue(preview),
      candidate: workbookInput,
    }),
    preview
  );
}
function applyWorkbookLinkPreview(workbook, preview, options = {}) {
  const previewRecord = workbookLinkPreviewRecords.get(preview);
  if (!previewRecord)
    throw Error("Unknown or already applied workbook link preview.");
  if (options.readOnly) throw Error("The workbook is read-only.");
  if (
    previewRecord.bookHash !== fingerprintCanonicalValue(workbook) ||
    previewRecord.contextHash !==
      buildWorkbookLinkPolicy(workbook, options).hash() ||
    previewRecord.planHash !== fingerprintCanonicalValue(preview)
  )
    throw Error(
      "Workbook, source, review or permissions changed. Preview links again.",
    );
  return (
    workbookLinkPreviewRecords.delete(preview),
    normalizeWorkbookInput(previewRecord.candidate)
  );
}

function assertSourceInspectionCurrent(inspection) {
  const inspectionRecord = sourceInspectionRecords.get(inspection);
  if (
    !inspectionRecord ||
    inspectionRecord.signature !== JSON.stringify(inspection)
  )
    throw Error(
      "The source inspection changed or is untrusted. Inspect the original again.",
    );
  return inspectionRecord.bytes.slice();
}

function assertFormatInspectionCurrent(inspection) {
  const inspectionRecord = formatInspectionRecords.get(inspection);
  if (
    !inspectionRecord ||
    inspectionRecord.signature !== JSON.stringify(inspection)
  )
    throw Error(
      "The format inspection changed or is untrusted. Inspect the original again.",
    );
  return inspectionRecord.bytes.slice();
}

export {
  assertSourceInspectionCurrent,
  CustomFunctionReviewBanner,
  applyWorkbookLinkPreview,
  recordSourceInspection,
  validateIsolatedFormatInspection,
  previewWorkbookLinkOperation,
  buildWorkbookLinkReview,
  assertFormatInspectionCurrent,
  AddInsAndCustomFunctionsDialog,
};
