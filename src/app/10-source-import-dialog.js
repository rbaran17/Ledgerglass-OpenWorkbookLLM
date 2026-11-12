import { X, Download, CircleCheck, FileInput } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  sourceImportLimits,
  retainAlternateFormatSource,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import { assertFormatInspectionCurrent } from "./05-add-ins-and-custom-functions.js";
import {
  importWorkbookFile,
  importNativeJsonFile,
} from "./07-file-import-and-worker-dispatch.js";
import {
  commitSourceImport,
  toSourceFileDescriptor,
  runDocumentExtraction,
  resolveImportedSourceCell,
  findSourceImport,
  Rue,
  buildSourceReview,
  downloadBlobAsFile,
  parseExtractionResult,
  useCancellableOperation,
} from "./09-workbook-links.js";

function SourceImportDialog(props) {
  const titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    latestProps = ne.useRef(props),
    operationCounter = ne.useRef(0),
    isMounted = ne.useRef(true),
    cancellableOperation = useCancellableOperation();
  latestProps.current = props;
  const [inspection, setInspection] = ne.useState(null),
    [tableId, setTableId] = ne.useState(""),
    [sheetName, setSheetName] = ne.useState("Imported data"),
    [ambiguityReviewed, setAmbiguityReviewed] = ne.useState(false),
    [checked, setChecked] = ne.useState(false),
    [resolutions, setResolutions] = ne.useState({}),
    [reviewPreview, setReviewPreview] = ne.useState(null),
    [previewedWorkbook, setPreviewedWorkbook] = ne.useState(null),
    [disabled, setDisabled] = ne.useState(false),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [statusMessage, setStatusMessage] = ne.useState(""),
    [tablePage, setTablePage] = ne.useState(0),
    [uncertainPage, setUncertainPage] = ne.useState(0),
    [correctionErrors, setCorrectionErrors] = ne.useState({}),
    [retainedImportId, setRetainedImportId] = ne.useState(""),
    [selectedCell, setSelectedCell] = ne.useState("0:0");
  ne.useEffect(() => {
    isMounted.current = true;
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.focus(),
      () => {
        ((isMounted.current = false),
          operationCounter.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []);
  const runOperation = async (operation) => {
      const operationId = ++operationCounter.current;
      (setDisabled(true), setErrorMessage(""), setStatusMessage(""));
      try {
        await operation(operationId);
      } catch (error) {
        isMounted.current &&
          operationId === operationCounter.current &&
          setErrorMessage(
            error instanceof Error
              ? error.message
              : "The source could not be processed.",
          );
      } finally {
        isMounted.current &&
          operationId === operationCounter.current &&
          setDisabled(false);
      }
    },
    applyInspection = (nextInspection) => {
      (setInspection(nextInspection),
        setTableId(nextInspection.tables[0]?.id ?? ""),
        setReviewPreview(null),
        setResolutions({}),
        setCorrectionErrors({}),
        setAmbiguityReviewed(false),
        setTablePage(0),
        setUncertainPage(0));
    },
    importSourceFile = (file) => (
      setInspection(null),
      setReviewPreview(null),
      setChecked(false),
      runOperation(async (requestId) => {
        const loadedInspection = await Rue(file, cancellableOperation.begin());
        isMounted.current &&
          operationCounter.current === requestId &&
          applyInspection(loadedInspection);
      })
    );
  ne.useEffect(() => {
    props.initialFile && importSourceFile(props.initialFile);
  }, [props.initialFile]);
  const X$1 = (jsonFile) =>
      runOperation(async (requestId) => {
        const currentInspection = inspection,
          extractionResult = await importNativeJsonFile(
            jsonFile,
            cancellableOperation.begin(),
          );
        isMounted.current &&
          operationCounter.current === requestId &&
          applyInspection(
            parseExtractionResult(currentInspection, extractionResult),
          );
      }),
    runHostExtraction = () =>
      runOperation(async (requestId) => {
        const extractedInspection = await runDocumentExtraction(
          inspection,
          props.extractor,
          checked,
          cancellableOperation.begin(),
        );
        isMounted.current &&
          operationCounter.current === requestId &&
          applyInspection(extractedInspection);
      }),
    downloadFile = (filename, blob) =>
      (latestProps.current.onDownload ?? downloadBlobAsFile)({
        filename: filename,
        blob: blob,
      }),
    downloadSource = (isRetainedSource) =>
      runOperation(async () => {
        const sourceFile = isRetainedSource
          ? findSourceImport(latestProps.current.workbook, retainedImportId)
          : toSourceFileDescriptor(inspection);
        await downloadFile(
          sourceFile.fileName,
          new Blob([new Uint8Array(sourceFile.bytes)], {
            type: "application/octet-stream",
          }),
        );
      }),
    clearPreview = () => setReviewPreview(null),
    selectedTable = inspection?.tables.find((table) => table.id === tableId),
    uncertainCells =
      selectedTable?.rows
        .flatMap((row, rowIndex) =>
          row.map((cell, column) => ({
            cell: cell,
            row: rowIndex,
            column: column,
            key: `${rowIndex}:${column}`,
          })),
        )
        .filter(
          (entry) => entry.cell.confidence < sourceImportLimits.confidence,
        ) ?? [],
    unresolvedCount = uncertainCells.filter(
      (event) => !resolutions[event.key],
    ).length,
    retainedImport = props.workbook.sourceImports?.find(
      (sourceImport) => sourceImport.id === retainedImportId,
    ),
    [selectedRow, selectedColumn] = selectedCell.split(":").map(Number),
    cellEvidence = retainedImport
      ? resolveImportedSourceCell(
          props.workbook,
          retainedImport.id,
          selectedRow,
          selectedColumn,
        )
      : null,
    onClick = () => {
      setErrorMessage("");
      try {
        if (Object.values(correctionErrors).some(Boolean))
          throw Error("Finish each corrected value before previewing.");
        const sourceReview = buildSourceReview(props.workbook, inspection, {
          tableId: tableId,
          sheetName: sheetName,
          ambiguityReviewed: ambiguityReviewed,
          resolutions: resolutions,
        });
        (setReviewPreview(sourceReview), setPreviewedWorkbook(props.workbook));
      } catch (error) {
        setErrorMessage(error.message);
      }
    },
    commitImport = () => {
      setErrorMessage("");
      try {
        if (latestProps.current.workbook !== previewedWorkbook)
          throw Error("The workbook changed. Preview the import again.");
        const nextWorkbook = commitSourceImport(
          latestProps.current.workbook,
          reviewPreview,
          {
            readOnly: latestProps.current.readOnly,
          },
        );
        (props.onWorkbookChange(nextWorkbook, "Import reviewed source"),
          setRetainedImportId(reviewPreview.record.id),
          setSelectedCell("0:0"),
          setReviewPreview(null),
          setStatusMessage(
            "Reviewed source imported into a new worksheet. The original and cell evidence are retained in native JSON.",
          ));
      } catch (error) {
        setErrorMessage(error.message);
      }
    },
    setCellResolution = (cellKey, resolution) => {
      (clearPreview(),
        setCorrectionErrors((previousErrors) => ({
          ...previousErrors,
          [cellKey]: false,
        })),
        setResolutions((previousResolutions) => ({
          ...previousResolutions,
          [cellKey]: resolution,
        })));
    };
  return i.jsx("div", {
    className: "lg-reports-overlay",
    children: i.jsxs("section", {
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      lang: "en",
      tabIndex: -1,
      ref: dialogRef,
      className: "lg-reports-dialog lg-formats-dialog lg-source-dialog",
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.key === "Escape" && (event.stopPropagation(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
            ...dialogRef.current.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
            ),
          ].filter((element) => element.getClientRects().length > 0);
          event.shiftKey &&
          (document.activeElement === focusableElements[0] ||
            document.activeElement === dialogRef.current)
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
                i.jsx("p", { children: "DATA & SOURCE EVIDENCE" }),
                i.jsx("h2", { id: titleId, children: "Import sources" }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              onClick: props.onClose,
              "aria-label": "Close import sources",
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-reports-body",
          children: [
            i.jsx("p", {
              children:
                "Bring JSON or XML tables into a new worksheet. Document extraction can be supplied by your host or a matching extraction JSON file. Review first; the current workbook stays unchanged until import.",
            }),
            errorMessage &&
              i.jsx("p", {
                role: "alert",
                className: "lg-formats-error",
                children: errorMessage,
              }),
            statusMessage &&
              i.jsxs("p", {
                role: "status",
                className: "lg-source-success",
                children: [i.jsx(CircleCheck, { size: 17 }), statusMessage],
              }),
            i.jsxs("div", {
              className: "lg-formats-actions",
              children: [
                i.jsxs("label", {
                  className: "lg-source-file",
                  children: [
                    i.jsx(FileInput, { size: 18 }),
                    "Choose source file",
                    i.jsx("input", {
                      "aria-label": "Import source file",
                      type: "file",
                      accept: ".json,.xml,.pdf,.png,.jpg,.jpeg,.html,.htm",
                      disabled: disabled,
                      onChange: (event) => {
                        const selectedFile = event.target.files?.[0];
                        ((event.target.value = ""),
                          selectedFile && importSourceFile(selectedFile));
                      },
                    }),
                  ],
                }),
                disabled &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsx("span", {
                        role: "status",
                        children: "Processing source…",
                      }),
                      i.jsx("button", {
                        type: "button",
                        onClick: () => {
                          (cancellableOperation.cancel(),
                            operationCounter.current++,
                            setDisabled(false),
                            setStatusMessage("Source processing cancelled."));
                        },
                        children: "Cancel processing",
                      }),
                    ],
                  }),
              ],
            }),
            inspection &&
              i.jsxs("section", {
                className: "lg-formats-source",
                children: [
                  i.jsxs("div", {
                    className: "lg-source-heading",
                    children: [
                      i.jsxs("div", {
                        children: [
                          i.jsx("h3", { children: inspection.fileName }),
                          i.jsxs("p", {
                            children: [
                              inspection.format.toUpperCase(),
                              " ·",
                              " ",
                              inspection.byteLength.toLocaleString(),
                              " bytes ·",
                              " ",
                              inspection.tables.length,
                              " table candidates",
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("button", {
                        type: "button",
                        disabled: disabled,
                        onClick: () => void downloadSource(false),
                        children: [
                          i.jsx(Download, { size: 16 }),
                          "Download original source",
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("p", {
                    className: "lg-formats-hash",
                    children: ["SHA-256: ", inspection.sha256],
                  }),
                  i.jsx("ul", {
                    children: inspection.warnings.map((warning, index) =>
                      i.jsx("li", { children: warning }, index),
                    ),
                  }),
                  inspection.method === "document-adapter" &&
                    i.jsxs("div", {
                      className: "lg-source-extractor",
                      children: [
                        i.jsx("h4", { children: "Document extraction" }),
                        i.jsx("p", {
                          children:
                            "No OCR or document provider runs automatically. Regions use a 1-based page and normalized top-left coordinates.",
                        }),
                        props.extractor
                          ? i.jsxs(i.Fragment, {
                              children: [
                                i.jsxs("label", {
                                  className: "lg-formats-ack",
                                  children: [
                                    i.jsx("input", {
                                      type: "checkbox",
                                      checked: checked,
                                      disabled: disabled,
                                      onChange: (event) =>
                                        setChecked(event.target.checked),
                                    }),
                                    "I authorize ",
                                    props.extractor.label,
                                    " to process this selected source.",
                                  ],
                                }),
                                i.jsx("button", {
                                  type: "button",
                                  disabled: !checked || disabled,
                                  onClick: () => void runHostExtraction(),
                                  children: "Run host extraction",
                                }),
                              ],
                            })
                          : i.jsx("p", {
                              children:
                                "Your host has not connected a document extractor. Load a matching extraction JSON result below.",
                            }),
                        i.jsxs("label", {
                          children: [
                            "Load extraction JSON",
                            i.jsx("input", {
                              type: "file",
                              accept: ".json",
                              disabled: disabled,
                              "aria-label": "Load extraction JSON",
                              onChange: (event) => {
                                const selectedFile = event.target.files?.[0];
                                ((event.target.value = ""),
                                  selectedFile && X$1(selectedFile));
                              },
                            }),
                          ],
                        }),
                      ],
                    }),
                ],
              }),
            selectedTable &&
              i.jsxs("section", {
                className: "lg-source-review",
                children: [
                  i.jsxs("div", {
                    className: "lg-source-options",
                    children: [
                      i.jsxs("label", {
                        children: [
                          "Source table",
                          i.jsx("select", {
                            "aria-label": "Source table",
                            value: tableId,
                            disabled: disabled,
                            onChange: (event) => {
                              (setTableId(event.target.value),
                                setReviewPreview(null),
                                setResolutions({}),
                                setAmbiguityReviewed(false),
                                setTablePage(0));
                            },
                            children: inspection.tables.map((table) =>
                              i.jsxs(
                                "option",
                                {
                                  value: table.id,
                                  children: [
                                    table.name,
                                    " · ",
                                    table.rows.length,
                                    " rows",
                                  ],
                                },
                                table.id,
                              ),
                            ),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "New worksheet name",
                          i.jsx("input", {
                            "aria-label": "New worksheet name",
                            value: sheetName,
                            maxLength: 31,
                            onChange: (event) => {
                              (setSheetName(event.target.value),
                                clearPreview());
                            },
                          }),
                        ],
                      }),
                    ],
                  }),
                  (selectedTable.ambiguous || inspection.tables.length > 1) &&
                    i.jsxs("label", {
                      className: "lg-formats-ack",
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: ambiguityReviewed,
                          onChange: (event) => {
                            (setAmbiguityReviewed(event.target.checked),
                              clearPreview());
                          },
                        }),
                        "I reviewed the candidates and confirm this table selection.",
                      ],
                    }),
                  i.jsxs("p", {
                    children: [
                      selectedTable.rows.length,
                      " rows · ",
                      selectedTable.columns.length,
                      " fields ·",
                      " ",
                      unresolvedCount,
                      " low-confidence cells awaiting review. Missing fields remain distinguishable from null in source evidence.",
                    ],
                  }),
                  i.jsx("div", {
                    className: "lg-formats-table lg-source-preview",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("caption", {
                          children:
                            "Source preview — values are not yet in the workbook",
                        }),
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", { scope: "col", children: "Row" }),
                              selectedTable.columns.map((column, columnIndex) =>
                                i.jsxs(
                                  "th",
                                  {
                                    scope: "col",
                                    children: [
                                      column.name,
                                      i.jsx("small", {
                                        children: column.path || "(root)",
                                      }),
                                    ],
                                  },
                                  columnIndex,
                                ),
                              ),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: selectedTable.rows
                            .slice(tablePage * 20, tablePage * 20 + 20)
                            .map((rowCells, rowIndex) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("th", {
                                      scope: "row",
                                      children: tablePage * 20 + rowIndex + 1,
                                    }),
                                    rowCells.map((cell, columnIndex) =>
                                      i.jsxs(
                                        "td",
                                        {
                                          className:
                                            cell.confidence < 0.9
                                              ? "lg-source-low"
                                              : "",
                                          children: [
                                            String(
                                              cell.value ??
                                                (cell.missing
                                                  ? "(missing)"
                                                  : "(null)"),
                                            ),
                                            i.jsxs("small", {
                                              children: [
                                                cell.type,
                                                cell.declaredType
                                                  ? " · " + cell.declaredType
                                                  : "",
                                                " · ",
                                                Math.round(
                                                  cell.confidence * 100,
                                                ),
                                                "%",
                                                cell.region
                                                  ? " · page " +
                                                    cell.region.page
                                                  : "",
                                              ],
                                            }),
                                          ],
                                        },
                                        columnIndex,
                                      ),
                                    ),
                                  ],
                                },
                                rowIndex,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "lg-formats-actions",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        disabled: tablePage === 0,
                        onClick: () =>
                          setTablePage((currentPage) => currentPage - 1),
                        children: "Previous source rows",
                      }),
                      i.jsxs("span", {
                        children: [
                          "Page ",
                          tablePage + 1,
                          " of ",
                          Math.ceil(selectedTable.rows.length / 20),
                        ],
                      }),
                      i.jsx("button", {
                        type: "button",
                        disabled:
                          (tablePage + 1) * 20 >= selectedTable.rows.length,
                        onClick: () =>
                          setTablePage((currentPage) => currentPage + 1),
                        children: "Next source rows",
                      }),
                    ],
                  }),
                  uncertainCells.length > 0 &&
                    i.jsxs("section", {
                      className: "lg-source-flags",
                      children: [
                        i.jsx("h3", { children: "Review uncertain cells" }),
                        i.jsx("p", {
                          children:
                            "Accept after comparison with the source, correct the value, or exclude it. Excluded cells become blanks and stay documented in the review record.",
                        }),
                        uncertainCells
                          .slice(uncertainPage * 10, uncertainPage * 10 + 10)
                          .map(({ cell, row, column, key }) =>
                            i.jsxs(
                              "div",
                              {
                                className: "lg-source-flag",
                                children: [
                                  i.jsxs("div", {
                                    children: [
                                      i.jsxs("strong", {
                                        children: [
                                          columnIndexToLetters(column),
                                          row + 2,
                                          " · ",
                                          Math.round(cell.confidence * 100),
                                          "% confidence",
                                        ],
                                      }),
                                      i.jsx("p", {
                                        children: String(
                                          cell.value ?? "(null)",
                                        ),
                                      }),
                                      i.jsxs("small", {
                                        children: [
                                          cell.path,
                                          " ·",
                                          " ",
                                          cell.region
                                            ? `Page ${cell.region.page} · x ${cell.region.x}, y ${cell.region.y}, w ${cell.region.width}, h ${cell.region.height}`
                                            : "Structured field",
                                        ],
                                      }),
                                    ],
                                  }),
                                  i.jsxs("label", {
                                    children: [
                                      "Review ",
                                      key,
                                      i.jsxs("select", {
                                        "aria-label": "Review " + key,
                                        value: resolutions[key]?.decision ?? "",
                                        onChange: (event) => {
                                          event.target.value
                                            ? setCellResolution(key, {
                                                decision: event.target.value,
                                                ...(event.target.value ===
                                                "correct"
                                                  ? { value: cell.value }
                                                  : {}),
                                              })
                                            : (setResolutions(
                                                (previousResolutions) => {
                                                  const nextResolutions = {
                                                    ...previousResolutions,
                                                  };
                                                  return (
                                                    delete nextResolutions[key],
                                                    nextResolutions
                                                  );
                                                },
                                              ),
                                              clearPreview());
                                        },
                                        children: [
                                          i.jsx("option", {
                                            value: "",
                                            children: "Review required",
                                          }),
                                          i.jsx("option", {
                                            value: "accept",
                                            children: "Accept source value",
                                          }),
                                          i.jsx("option", {
                                            value: "exclude",
                                            children: "Exclude from worksheet",
                                          }),
                                          i.jsx("option", {
                                            value: "correct",
                                            children: "Correct value",
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                                  resolutions[key]?.decision === "correct" &&
                                    i.jsxs("label", {
                                      children: [
                                        "Corrected value (JSON scalar)",
                                        i.jsx("input", {
                                          "aria-label":
                                            "Corrected value " + key,
                                          "aria-invalid":
                                            correctionErrors[key] || undefined,
                                          defaultValue: JSON.stringify(
                                            cell.value,
                                          ),
                                          onChange: (event) => {
                                            try {
                                              const parsedValue = JSON.parse(
                                                event.target.value,
                                              );
                                              if (
                                                parsedValue !== null &&
                                                ![
                                                  "string",
                                                  "number",
                                                  "boolean",
                                                ].includes(typeof parsedValue)
                                              )
                                                throw Error();
                                              (setCellResolution(key, {
                                                decision: "correct",
                                                value: parsedValue,
                                              }),
                                                event.target.setCustomValidity(
                                                  "",
                                                ));
                                            } catch {
                                              (setCorrectionErrors(
                                                (previousErrors) => ({
                                                  ...previousErrors,
                                                  [key]: true,
                                                }),
                                              ),
                                                setReviewPreview(null),
                                                event.target.setCustomValidity(
                                                  "Enter a JSON string, number, boolean or null.",
                                                ));
                                            }
                                          },
                                        }),
                                      ],
                                    }),
                                ],
                              },
                              key,
                            ),
                          ),
                        i.jsxs("div", {
                          className: "lg-formats-actions",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              disabled: uncertainPage === 0,
                              onClick: () =>
                                setUncertainPage(
                                  (currentPage) => currentPage - 1,
                                ),
                              children: "Previous uncertain cells",
                            }),
                            i.jsxs("span", {
                              children: [
                                "Review page ",
                                uncertainPage + 1,
                                " of",
                                " ",
                                Math.ceil(uncertainCells.length / 10),
                              ],
                            }),
                            i.jsx("button", {
                              type: "button",
                              disabled:
                                (uncertainPage + 1) * 10 >=
                                uncertainCells.length,
                              onClick: () =>
                                setUncertainPage(
                                  (currentPage) => currentPage + 1,
                                ),
                              children: "Next uncertain cells",
                            }),
                          ],
                        }),
                      ],
                    }),
                  i.jsx("div", {
                    className: "lg-formats-actions",
                    children: i.jsx("button", {
                      type: "button",
                      className: "lg-source-primary",
                      disabled:
                        disabled ||
                        props.readOnly ||
                        Object.values(correctionErrors).some(Boolean) ||
                        unresolvedCount > 0 ||
                        ((selectedTable.ambiguous ||
                          inspection.tables.length > 1) &&
                          !ambiguityReviewed),
                      onClick: onClick,
                      children: "Preview import",
                    }),
                  }),
                  reviewPreview &&
                    i.jsxs("section", {
                      className: "lg-source-confirm",
                      "aria-label": "Import confirmation",
                      children: [
                        i.jsxs("h3", {
                          children: ["Ready for ", reviewPreview.sheet.name],
                        }),
                        i.jsxs("p", {
                          children: [
                            reviewPreview.accepted,
                            " accepted · ",
                            reviewPreview.corrected,
                            " corrected ·",
                            " ",
                            reviewPreview.excluded,
                            " excluded cells. Source evidence will be retained in native JSON.",
                          ],
                        }),
                        previewedWorkbook !== props.workbook &&
                          i.jsx("p", {
                            role: "alert",
                            children: "The workbook changed. Preview again.",
                          }),
                        i.jsx("button", {
                          type: "button",
                          className: "lg-source-primary",
                          disabled:
                            props.readOnly ||
                            previewedWorkbook !== props.workbook ||
                            disabled,
                          onClick: commitImport,
                          children: "Import reviewed table",
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => setReviewPreview(null),
                          children: "Cancel import preview",
                        }),
                      ],
                    }),
                ],
              }),
            !!props.workbook.sourceImports?.length &&
              i.jsxs("section", {
                className: "lg-formats-source",
                children: [
                  i.jsx("h3", { children: "Retained sources & cell evidence" }),
                  i.jsxs("label", {
                    children: [
                      "Retained import",
                      i.jsxs("select", {
                        "aria-label": "Retained import",
                        value: retainedImportId,
                        onChange: (event) => {
                          (setRetainedImportId(event.target.value),
                            setSelectedCell("0:0"));
                        },
                        children: [
                          i.jsx("option", {
                            value: "",
                            children: "Select an imported source",
                          }),
                          props.workbook.sourceImports.map((sourceImport) =>
                            i.jsxs(
                              "option",
                              {
                                value: sourceImport.id,
                                children: [
                                  sourceImport.sheetName,
                                  " · ",
                                  sourceImport.original.fileName,
                                ],
                              },
                              sourceImport.id,
                            ),
                          ),
                        ],
                      }),
                    ],
                  }),
                  retainedImport &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("p", {
                          className: "lg-formats-hash",
                          children: [
                            "SHA-256: ",
                            retainedImport.original.sha256,
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => void downloadSource(true),
                          children: "Download retained source",
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () =>
                            void runOperation(async () => {
                              const { base64, ...original } =
                                retainedImport.original;
                              await downloadFile(
                                retainedImport.original.fileName +
                                  ".import-review.json",
                                new Blob(
                                  [
                                    JSON.stringify(
                                      { ...retainedImport, original: original },
                                      null,
                                      2,
                                    ),
                                  ],
                                  { type: "application/json" },
                                ),
                              );
                            }),
                          children: "Download source review",
                        }),
                        i.jsxs("div", {
                          className: "lg-source-options",
                          children: [
                            i.jsxs("label", {
                              children: [
                                "Source record",
                                i.jsx(LocalizedNumberField, {
                                  "aria-label": "Source record",
                                  min: 1,
                                  max: retainedImport.table.rows.length,
                                  value: selectedRow + 1,
                                  onValueChange: (inputValue) => {
                                    const recordNumber = Number(inputValue);
                                    Number.isInteger(recordNumber) &&
                                      recordNumber >= 1 &&
                                      recordNumber <=
                                        retainedImport.table.rows.length &&
                                      setSelectedCell(
                                        recordNumber - 1 + ":" + selectedColumn,
                                      );
                                  },
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                "Source field",
                                i.jsx("select", {
                                  "aria-label": "Source field",
                                  value: selectedColumn,
                                  onChange: (event) =>
                                    setSelectedCell(
                                      selectedRow + ":" + event.target.value,
                                    ),
                                  children: retainedImport.table.columns.map(
                                    (column, columnIndex) =>
                                      i.jsxs(
                                        "option",
                                        {
                                          value: columnIndex,
                                          children: [
                                            columnIndexToLetters(columnIndex),
                                            " · ",
                                            column.name,
                                          ],
                                        },
                                        columnIndex,
                                      ),
                                  ),
                                }),
                              ],
                            }),
                          ],
                        }),
                        cellEvidence &&
                          i.jsxs("dl", {
                            className: "lg-source-evidence",
                            children: [
                              i.jsx("dt", { children: "Link status" }),
                              i.jsx("dd", { children: cellEvidence.status }),
                              i.jsx("dt", { children: "Source field" }),
                              i.jsx("dd", {
                                children: cellEvidence.source.path || "(root)",
                              }),
                              i.jsx("dt", {
                                children: "Original type & value",
                              }),
                              i.jsxs("dd", {
                                children: [
                                  cellEvidence.source.type,
                                  " ·",
                                  " ",
                                  String(cellEvidence.source.value ?? "(null)"),
                                ],
                              }),
                              i.jsx("dt", { children: "Review decision" }),
                              i.jsx("dd", { children: cellEvidence.decision }),
                              i.jsx("dt", { children: "Source region" }),
                              i.jsx("dd", {
                                children: cellEvidence.source.region
                                  ? `Page ${cellEvidence.source.region.page} · x ${cellEvidence.source.region.x}, y ${cellEvidence.source.region.y}, width ${cellEvidence.source.region.width}, height ${cellEvidence.source.region.height}`
                                  : "Structured field path",
                              }),
                            ],
                          }),
                        i.jsx("p", {
                          children:
                            "Evidence records the import position. Edits, moves or sorts can change the link status; compare the original before relying on a changed cell.",
                        }),
                      ],
                    }),
                ],
              }),
            i.jsx("p", {
              className: "lg-source-limits",
              children:
                "Limits: 4 MiB per source; 20 table candidates; 1,000 rows, 100 columns and 20,000 cells across candidates; 10 retained imports and 12 MiB original bytes. Native JSON preserves originals and provenance; XLSX and delimited exports carry worksheet values without this evidence.",
            }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children: props.readOnly
                ? "Read-only: inspection and original downloads are available."
                : "Import appends a new worksheet. Undo restores the prior workbook.",
            }),
            i.jsx("button", {
              type: "button",
              onClick: props.onClose,
              children: "Close",
            }),
          ],
        }),
      ],
    }),
  });
}
async function validateConvertedSource(inspection, convertedFile, checked, n) {
  const o = assertFormatInspectionCurrent(inspection);
  if (checked !== true)
    throw Error(
      "Acknowledge conversion losses and compare the source before importing the converted workbook.",
    );
  if (!/\.xlsx$/i.test(convertedFile.name))
    throw Error(
      "Choose the converted XLSX copy; renaming the original does not convert it.",
    );
  const importResult = await importWorkbookFile(convertedFile, n);
  if (
    (assertFormatInspectionCurrent(inspection),
    !importResult.workbook.xlsxSource || !importResult.report)
  )
    throw Error(
      "A screened XLSX conversion and compatibility report are required.",
    );
  importResult.workbook.alternateSource = retainAlternateFormatSource(
    o,
    inspection.fileName,
    inspection.format,
    {
      method: "external-conversion",
      convertedFileName: convertedFile.name,
      convertedSha256: importResult.workbook.xlsxSource.sha256,
      reviewedAt: new Date().toISOString(),
      equivalence: "not-verified",
    },
  );
  const conversionWarning =
    "An externally converted XLSX was selected and conversion losses acknowledged. This application has not verified that its formulas, dates, styles or special parts match the retained original. Compare them in the source application. The original is retained under File > Workbook formats and in native JSON.";
  return (
    (importResult.warnings = [
      ...new Set([...importResult.warnings, conversionWarning]),
    ]),
    (importResult.report.warnings = importResult.warnings),
    (importResult.report.alternateOriginal = {
      fileName: inspection.fileName,
      byteLength: inspection.fileBytes,
      sha256: inspection.sha256,
      format: inspection.format,
    }),
    importResult
  );
}
function describeSourceFile(inspection) {
  return {
    bytes: assertFormatInspectionCurrent(inspection),
    fileName: inspection.fileName,
    sha256: inspection.sha256,
  };
}
const partDispositionLabels = {
  editable: "Editable",
  "preserved-opaque": "Preserved opaque",
  flattened: "Flattened",
  dropped: "Dropped",
};
function SourcePackageInventory({ inventory }) {
  const { text, language } = useWorkbookMessages(),
    [searchText, setSearchText] = ne.useState(""),
    [dispositionFilter, setDispositionFilter] = ne.useState("all"),
    [pageIndex, setPageIndex] = ne.useState(0),
    matchingParts = inventory.parts.filter(
      (part) =>
        (dispositionFilter === "all" ||
          part.disposition === dispositionFilter) &&
        [
          part.path,
          part.family,
          ...part.locations.map((sourceLocation) => sourceLocation.sheet ?? ""),
          ...part.features.map(
            (feature) =>
              `${feature.feature} ${feature.location.range ?? ""} ${feature.location.detail ?? ""}`,
          ),
        ]
          .join(" ")
          .toLowerCase()
          .includes(searchText.toLowerCase()),
    ),
    lastPageIndex = Math.max(0, Math.ceil(matchingParts.length / 25) - 1),
    currentPageIndex = Math.min(pageIndex, lastPageIndex);
  return i.jsxs("section", {
    className: "lg-xlsx-inventory",
    "aria-label": text("Source package inventory"),
    lang: language("Source package inventory"),
    "data-ui-label": "Source package inventory",
    children: [
      i.jsx("h3", {
        children: i.jsx(LocalizedText, {
          children: "Source parts and export outcome",
        }),
      }),
      i.jsxs("p", {
        children: [
          inventory.parts.length,
          " ",
          i.jsx(LocalizedText, { children: "entries inventoried ·" }),
          " ",
          inventory.counts.editable,
          " ",
          i.jsx(LocalizedText, { children: "editable ·" }),
          " ",
          inventory.counts["preserved-opaque"],
          " ",
          i.jsx(LocalizedText, { children: "preserved opaque ·" }),
          " ",
          inventory.counts.flattened,
          " ",
          i.jsx(LocalizedText, { children: "flattened · " }),
          inventory.counts.dropped,
          " ",
          i.jsx(LocalizedText, { children: "dropped from the derived copy." }),
        ],
      }),
      i.jsx("p", { children: inventory.note }),
      i.jsx("p", {
        children:
          inventory.phase === "export-result"
            ? "These are the checked outcomes for this generated file."
            : "Opaque parts shown here are eligible for export. Export checks the current package and limits again before reporting the final outcome.",
      }),
      i.jsxs("div", {
        className: "lg-xlsx-filters",
        children: [
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, {
                children: "Find a part, sheet or feature",
              }),
              i.jsx("input", {
                value: searchText,
                onChange: (event) => {
                  (setSearchText(event.target.value), setPageIndex(0));
                },
                type: "search",
              }),
            ],
          }),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Export outcome" }),
              i.jsxs("select", {
                value: dispositionFilter,
                onChange: (event) => {
                  (setDispositionFilter(event.target.value), setPageIndex(0));
                },
                children: [
                  i.jsx("option", { value: "all", children: "All outcomes" }),
                  Object.entries(partDispositionLabels).map(
                    ([disposition, dispositionLabel]) =>
                      i.jsx(
                        "option",
                        { value: disposition, children: dispositionLabel },
                        disposition,
                      ),
                  ),
                ],
              }),
            ],
          }),
        ],
      }),
      i.jsxs("table", {
        "aria-label": text("Workbook part inventory"),
        lang: language("Workbook part inventory"),
        "data-ui-label": "Workbook part inventory",
        children: [
          i.jsx("thead", {
            children: i.jsxs("tr", {
              children: [
                i.jsx("th", {
                  children: i.jsx(LocalizedText, {
                    children: "Part and source location",
                  }),
                }),
                i.jsx("th", {
                  children: i.jsx(LocalizedText, { children: "Derived copy" }),
                }),
                i.jsx("th", {
                  children: i.jsx(LocalizedText, { children: "Details" }),
                }),
              ],
            }),
          }),
          i.jsx("tbody", {
            children: matchingParts
              .slice(currentPageIndex * 25, currentPageIndex * 25 + 25)
              .map((partEntry) =>
                i.jsxs(
                  "tr",
                  {
                    children: [
                      i.jsxs("td", {
                        children: [
                          i.jsx("strong", { children: partEntry.path }),
                          i.jsxs("small", {
                            children: [
                              partEntry.family,
                              " · ",
                              partEntry.bytes.toLocaleString(),
                              " ",
                              i.jsx(LocalizedText, { children: "bytes" }),
                            ],
                          }),
                          partEntry.locations.map((location, locationIndex) =>
                            i.jsxs(
                              "small",
                              {
                                children: [
                                  location.sheet ??
                                    (location.scope === "unresolved"
                                      ? "Unresolved location"
                                      : "Workbook"),
                                  location.range ? `!${location.range}` : "",
                                  location.detail
                                    ? ` — ${location.detail}`
                                    : "",
                                ],
                              },
                              locationIndex,
                            ),
                          ),
                        ],
                      }),
                      i.jsxs("td", {
                        children: [
                          i.jsx("span", {
                            className: `lg-xlsx-outcome ${partEntry.disposition}`,
                            children:
                              partDispositionLabels[partEntry.disposition],
                          }),
                          i.jsx("small", {
                            children: i.jsx(LocalizedText, {
                              children: "Retained in original",
                            }),
                          }),
                        ],
                      }),
                      i.jsxs("td", {
                        children: [
                          partEntry.reason,
                          i.jsxs("details", {
                            children: [
                              i.jsxs("summary", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Part evidence",
                                  }),
                                  partEntry.features.length
                                    ? ` (${partEntry.features.length} features)`
                                    : "",
                                ],
                              }),
                              i.jsx("p", { children: partEntry.contentType }),
                              i.jsxs("p", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "SHA-256: ",
                                  }),
                                  i.jsx("code", { children: partEntry.sha256 }),
                                ],
                              }),
                              partEntry.features
                                .slice(0, 200)
                                .map((partFeature, featureIndex) =>
                                  i.jsxs(
                                    "p",
                                    {
                                      children: [
                                        partFeature.location.sheet
                                          ? `${partFeature.location.sheet}!`
                                          : "",
                                        partFeature.location.range ??
                                          partFeature.location.detail ??
                                          "Workbook",
                                        " · ",
                                        partFeature.feature,
                                        " · ",
                                        partDispositionLabels[
                                          partFeature.disposition
                                        ],
                                        " —",
                                        " ",
                                        partFeature.reason,
                                      ],
                                    },
                                    featureIndex,
                                  ),
                                ),
                              partEntry.features.length > 200 &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children:
                                        "Showing 200 features. Save the compatibility report for all ",
                                    }),
                                    partEntry.features.length,
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "entries.",
                                    }),
                                  ],
                                }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  },
                  partEntry.path,
                ),
              ),
          }),
        ],
      }),
      !matchingParts.length &&
        i.jsx("p", {
          role: "status",
          children: i.jsx(LocalizedText, {
            children: "No matching source parts.",
          }),
        }),
      i.jsxs("div", {
        className: "lg-xlsx-pages",
        children: [
          i.jsx("button", {
            disabled: currentPageIndex === 0,
            onClick: () => setPageIndex(currentPageIndex - 1),
            children: i.jsx(LocalizedText, { children: "Previous parts" }),
          }),
          i.jsxs("span", {
            "aria-live": "polite",
            children: [
              matchingParts.length,
              " ",
              i.jsx(LocalizedText, { children: "matching parts · Page " }),
              currentPageIndex + 1,
              " ",
              i.jsx(LocalizedText, { children: "of " }),
              lastPageIndex + 1,
            ],
          }),
          i.jsx("button", {
            disabled: currentPageIndex === lastPageIndex,
            onClick: () => setPageIndex(currentPageIndex + 1),
            children: i.jsx(LocalizedText, { children: "Next parts" }),
          }),
        ],
      }),
    ],
  });
}
function XlsxConversionReview({
  report,
  blocked,
  onConfirm,
  onClose,
  onReport,
  onOriginal,
  originalLabel = "Download original XLSX",
  inspectOnly = false,
}) {
  const { text, language } = useWorkbookMessages(),
    dialogTitleId = ne.useId(),
    dialogRef = ne.useRef(null);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const isImport = report.direction === "import",
    caches = report.caches;
  return i.jsx("div", {
    className: "lg-xlsx-overlay",
    children: i.jsxs("section", {
      ref: dialogRef,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": dialogTitleId,
      tabIndex: -1,
      className: "lg-xlsx-review",
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.preventDefault(), event.stopPropagation(), onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
              ...(dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),a[href],summary,[tabindex="0"]',
              ) ?? []),
            ],
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1);
          event.shiftKey &&
          (document.activeElement === firstFocusable ||
            document.activeElement === dialogRef.current)
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === lastFocusable &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
        }
      },
      children: [
        i.jsx("h2", {
          id: dialogTitleId,
          children: inspectOnly
            ? "Original XLSX and compatibility"
            : `Review ${report.format === "xltx" ? "XLTX" : "XLSX"} ${isImport ? "import" : "export"}`,
        }),
        i.jsxs("p", {
          children: [
            i.jsx("strong", { children: report.fileName }),
            " ·",
            " ",
            (report.fileBytes / 1024).toFixed(1),
            " ",
            i.jsx(LocalizedText, { children: "KiB · " }),
            report.sheets.length,
            " ",
            i.jsx(LocalizedText, {
              children: "worksheets in current conversion ·",
            }),
            " ",
            report.sheets
              .reduce((total, sheet) => total + sheet.storedCells, 0)
              .toLocaleString(),
            " ",
            i.jsx(LocalizedText, { children: "stored cells" }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "This is a derived copy with supported content. It is not a lossless Excel conversion. Keep the original file and a native JSON backup.",
          }),
        }),
        inspectOnly
          ? i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "This is the retained source snapshot. It does not include your subsequent edits. Source part locations refer to the original sheet names and addresses.",
              }),
            })
          : isImport
            ? i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, {
                    children:
                      "Opening replaces the current workbook. Cancel keeps it unchanged; Undo can restore it after opening.",
                  }),
                  " ",
                  report.originalFileRetained
                    ? "The exact original is retained in this workbook and native JSON; access it later from the File menu."
                    : "No source bytes are attached to this report.",
                ],
              })
            : i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, {
                    children: "The download contains revision ",
                  }),
                  report.revision,
                  " ",
                  i.jsx(LocalizedText, { children: "of" }),
                  " ",
                  report.workbookName,
                  i.jsx(LocalizedText, {
                    children:
                      ". Unsupported native features remain in your current workbook and JSON backup.",
                  }),
                ],
              }),
        report.originalSource &&
          i.jsxs("p", {
            className: "lg-xlsx-source",
            children: [
              i.jsx(LocalizedText, {
                children: report.alternateOriginal
                  ? "Converted XLSX retained: "
                  : "Original retained: ",
              }),
              i.jsx("strong", { children: report.originalSource.fileName }),
              " ·",
              " ",
              report.originalSource.byteLength.toLocaleString(),
              " ",
              i.jsx(LocalizedText, { children: "bytes" }),
              i.jsx("br", {}),
              i.jsxs("small", {
                children: [
                  i.jsx(LocalizedText, { children: "SHA-256: " }),
                  i.jsx("code", { children: report.originalSource.sha256 }),
                ],
              }),
            ],
          }),
        report.alternateOriginal &&
          i.jsxs("p", {
            className: "lg-xlsx-source",
            children: [
              i.jsxs("strong", {
                children: [
                  "Original ",
                  report.alternateOriginal.format.toUpperCase(),
                  ": ",
                  report.alternateOriginal.fileName,
                ],
              }),
              " · ",
              report.alternateOriginal.byteLength.toLocaleString(),
              " bytes",
              i.jsx("br", {}),
              i.jsxs("small", {
                children: [
                  "SHA-256: ",
                  i.jsx("code", { children: report.alternateOriginal.sha256 }),
                ],
              }),
            ],
          }),
        i.jsxs("div", {
          className: "lg-xlsx-scroll",
          children: [
            report.partInventory &&
              i.jsx(
                SourcePackageInventory,
                { inventory: report.partInventory },
                report.partInventory.sourceSha256,
              ),
            !inspectOnly &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, { children: "Worksheets" }),
                  }),
                  i.jsx("ul", {
                    children: report.sheets.map((sheet) =>
                      i.jsxs(
                        "li",
                        {
                          children: [
                            sheet.name,
                            " — ",
                            sheet.storedCells.toLocaleString(),
                            " ",
                            i.jsx(LocalizedText, { children: "cells" }),
                            sheet.hidden ? " · hidden" : "",
                          ],
                        },
                        sheet.name,
                      ),
                    ),
                  }),
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Formula results",
                    }),
                  }),
                  isImport
                    ? i.jsxs("p", {
                        children: [
                          caches.pending
                            ? `${caches.pending} pending comparisons; `
                            : "",
                          caches.total,
                          " ",
                          i.jsx(LocalizedText, { children: "formulas: " }),
                          caches.matched,
                          " ",
                          i.jsx(LocalizedText, { children: "caches match, " }),
                          caches.different,
                          " ",
                          i.jsx(LocalizedText, { children: "differ, " }),
                          caches.missing,
                          " ",
                          i.jsx(LocalizedText, { children: "missing, " }),
                          caches.neutralized,
                          " ",
                          i.jsx(LocalizedText, {
                            children:
                              "converted to text. Local recalculation is used; a saved source cache is historical evidence, not proof of correctness.",
                          }),
                        ],
                      })
                    : i.jsxs("p", {
                        children: [
                          caches.written,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "formula caches written; ",
                          }),
                          caches.omitted,
                          " ",
                          i.jsx(LocalizedText, { children: "omitted;" }),
                          " ",
                          caches.neutralized,
                          " ",
                          i.jsx(LocalizedText, {
                            children:
                              "formulas converted to text. Excel is asked to recalculate on open.",
                          }),
                        ],
                      }),
                  !!caches.issues.length &&
                    i.jsxs("table", {
                      "aria-label": text("Formula cache differences"),
                      lang: language("Formula cache differences"),
                      "data-ui-label": "Formula cache differences",
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Cell",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Saved result",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Local result",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Status",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: caches.issues.map((issue) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("td", {
                                    children: [issue.sheet, "!", issue.address],
                                  }),
                                  i.jsx("td", {
                                    children: issue.source
                                      ? `${issue.source.type}: ${String(issue.source.value)}`
                                      : "No cache",
                                  }),
                                  i.jsxs("td", {
                                    children: [
                                      issue.calculated.type,
                                      ": ",
                                      String(issue.calculated.value),
                                    ],
                                  }),
                                  i.jsx("td", { children: issue.status }),
                                ],
                              },
                              `${issue.sheet}!${issue.address}`,
                            ),
                          ),
                        }),
                      ],
                    }),
                  !!caches.detailsOmitted &&
                    i.jsxs("p", {
                      children: [
                        caches.detailsOmitted,
                        " ",
                        i.jsx(LocalizedText, {
                          children:
                            "additional cache details omitted from the bounded report; counts include all formulas.",
                        }),
                      ],
                    }),
                ],
              }),
            i.jsx("h3", {
              children: i.jsx(LocalizedText, {
                children: "Compatibility notices",
              }),
            }),
            i.jsxs("ul", {
              "aria-label": text("Compatibility notices"),
              lang: language("Compatibility notices"),
              "data-ui-label": "Compatibility notices",
              children: [
                report.embeddedObjectLosses?.length
                  ? i.jsxs("li", {
                      children: [
                        i.jsxs("strong", {
                          children: [
                            report.embeddedObjectLosses.length,
                            " unsupported embedded-object item(s) omitted",
                          ],
                        }),
                        i.jsx("ul", {
                          children: report.embeddedObjectLosses.map(
                            (embeddedObjectLoss, lossIndex) =>
                              i.jsxs(
                                "li",
                                {
                                  children: [
                                    embeddedObjectLoss.sourceFile,
                                    " · ",
                                    embeddedObjectLoss.path,
                                    " · ",
                                    embeddedObjectLoss.kind,
                                    " · dropped · never executed",
                                  ],
                                },
                                lossIndex,
                              ),
                          ),
                        }),
                      ],
                    })
                  : null,
                report.warnings.map((warning, warningIndex) =>
                  i.jsx("li", { children: warning }, warningIndex),
                ),
              ],
            }),
          ],
        }),
        blocked && i.jsx("p", { role: "alert", children: blocked }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: onClose,
              children: inspectOnly ? "Close" : "Cancel",
            }),
            report.originalFileRetained &&
              onOriginal &&
              i.jsx("button", {
                onClick: onOriginal,
                children: i.jsx(LocalizedText, { children: originalLabel }),
              }),
            i.jsx("button", {
              onClick: onReport,
              children: i.jsx(LocalizedText, {
                children: "Save compatibility report",
              }),
            }),
            !inspectOnly &&
              i.jsx("button", {
                disabled: !!blocked,
                onClick: onConfirm,
                children: isImport
                  ? "Open derived workbook"
                  : `Download derived ${report.format === "xltx" ? "XLTX" : "XLSX"}`,
              }),
          ],
        }),
      ],
    }),
  });
}

export {
  validateConvertedSource,
  SourceImportDialog,
  describeSourceFile,
  XlsxConversionReview,
};
