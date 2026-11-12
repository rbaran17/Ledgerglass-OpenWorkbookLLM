import { buildSnapshotGuard } from "./03-workbook-model.js";
import { X, Palette, FileText, Printer } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  parseFormulaTree,
  listInstalledExtensionFunctions,
} from "./01-schemas-date-functions-and-validation.js";
import {
  verifyPrintSnapshotGuard,
  builtinFunctionNames,
  evaluateConditionalFormatRules,
  validatePrintPageBreaks,
  assertRangeCellBudget,
  calculateWorkbookWithSettings,
  defaultPrintHeaderFooter,
  createDefaultPageSetup,
  parseRangeAddress,
  assertSheetUnprotected,
  applyPrintLayoutChange,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  renderSchedulePdf,
  printDocumentInWindow,
  PrintPagePreview,
  renderPrintDocument,
  waitForPrintAssets,
  summarizePrintIssues,
  measurePrintLayoutIssues,
  printPaperStyles,
  buildPrintPages,
} from "./32-drawing-objects-and-printing.js";

function PageSetupDialog(props) {
  const { text, language } = useWorkbookMessages(),
    propsRef = ne.useRef(props);
  propsRef.current = props;
  const printAreaRef = ne.useRef(null),
    abortRef = ne.useRef(null),
    [printIssues, setPrintIssues] = ne.useState(null),
    [isChecking, setIsChecking] = ne.useState(false),
    [isBusy, setIsBusy] = ne.useState(false),
    [progressMessage, setProgressMessage] = ne.useState(""),
    [pdfDocument, setPdfDocument] = ne.useState(null);
  (ne.useEffect(
    () => () => {
      pdfDocument && URL.revokeObjectURL(pdfDocument.url);
    },
    [pdfDocument],
  ),
    ne.useEffect(() => () => abortRef.current?.abort(), []));
  const [openingState] = ne.useState(() => ({
      snapshot: buildSnapshotGuard(props.workbook),
      sheetId: props.sheetId,
      range: props.range,
      sheet: structuredClone(
        props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
      ),
    })),
    sheetOrThrow =
      props.workbook.sheets.find((sheet) => sheet.id === props.sheetId) ??
      openingState.sheet;
  function verifySheetCurrent(shouldCheckEditable = false) {
    const latestProps = propsRef.current;
    if (
      (verifyPrintSnapshotGuard(latestProps.workbook, openingState.snapshot),
      latestProps.sheetId !== openingState.sheetId ||
        latestProps.range !== openingState.range)
    )
      throw Error(
        "The worksheet or selection changed. Close and reopen page setup.",
      );
    if (shouldCheckEditable) {
      if (latestProps.readOnly) throw Error("This workbook is read-only.");
      assertSheetUnprotected(
        latestProps.workbook.sheets.find(
          (sheet) => sheet.id === openingState.sheetId,
        ),
      );
    }
  }
  const [layout, setLayout] = ne.useState(() =>
      structuredClone(
        sheetOrThrow.pageLayout ?? {
          ...createDefaultPageSetup(sheetOrThrow),
          ...defaultPrintHeaderFooter,
        },
      ),
    ),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [themeId, setThemeId] = ne.useState("office"),
    [rowBreakText, setRowBreakText] = ne.useState(() =>
      (sheetOrThrow.pageLayout?.rowBreaks ?? [])
        .map((rowIndex) => String(rowIndex + 1))
        .join(", "),
    ),
    [columnBreakText, setColumnBreakText] = ne.useState(() =>
      (sheetOrThrow.pageLayout?.columnBreaks ?? [])
        .map(columnIndexToLetters)
        .join(", "),
    );
  function buildLayoutDraft() {
    const parseBreakList = (breakText, breakKind) => {
      if (!breakText.trim()) return;
      const breakIndexes = breakText
        .split(",")
        .map((segment) => {
          const segmentText = segment.trim();
          if (
            !(breakKind === "row" ? /^[1-9]\d*$/ : /^[A-Za-z]{1,3}$/).test(
              segmentText,
            )
          )
            throw Error(
              "Enter comma-separated row numbers or column letters for manual breaks.",
            );
          return breakKind === "row"
            ? Number(segmentText) - 1
            : parseRangeAddress(`${segmentText}1`).firstCol;
        })
        .sort((leftIndex, rightIndex) => leftIndex - rightIndex);
      return validatePrintPageBreaks(
        breakIndexes,
        breakKind === "row" ? sheetOrThrow.rowCount : sheetOrThrow.colCount,
      );
    };
    return {
      ...layout,
      rowBreaks: parseBreakList(rowBreakText, "row"),
      columnBreaks: parseBreakList(columnBreakText, "column"),
    };
  }
  const isThemeMode = /theme/i.test(props.mode),
    calc = ne.useMemo(
      () => calculateWorkbookWithSettings(props.workbook),
      [props.workbook],
    ),
    printPreview = ne.useMemo(() => {
      try {
        if (!props.workbook.sheets.some((sheet) => sheet.id === props.sheetId))
          throw Error(
            "The worksheet changed or was deleted. Close and reopen page setup.",
          );
        return {
          ...buildPrintPages(props.workbook, props.sheetId, buildLayoutDraft()),
          error: "",
        };
      } catch (error) {
        return {
          pages: [],
          notices: [],
          scale: 1,
          layout: layout,
          ruleStyles: evaluateConditionalFormatRules(
            { ...sheetOrThrow, conditionalFormats: [] },
            calc,
          ),
          error: error instanceof Error ? error.message : "Invalid page setup.",
        };
      }
    }, [
      props.workbook,
      props.sheetId,
      sheetOrThrow,
      layout,
      calc,
      rowBreakText,
      columnBreakText,
    ]),
    layoutSignature = JSON.stringify([layout, rowBreakText, columnBreakText]),
    layoutSignatureRef = ne.useRef(layoutSignature);
  ((layoutSignatureRef.current = layoutSignature),
    ne.useEffect(() => {
      (setPrintIssues(null), setErrorMessage(""), setPdfDocument(null));
    }, [layoutSignature, props.workbook, props.sheetId, props.range]));
  async function onClick() {
    const printElement = printAreaRef.current;
    if (!(!printElement || printPreview.error)) {
      (setIsChecking(true), setErrorMessage(""));
      try {
        if (
          (verifySheetCurrent(),
          await waitForPrintAssets(printElement),
          verifySheetCurrent(),
          layoutSignatureRef.current !== layoutSignature)
        )
          throw Error("Page settings changed. Run the print check again.");
        setPrintIssues(measurePrintLayoutIssues(printElement));
      } catch (error) {
        setErrorMessage(
          error instanceof Error ? error.message : "Print check failed.",
        );
      } finally {
        setIsChecking(false);
      }
    }
  }
  const updateLayout = (patch) =>
    setLayout((previous) => ({ ...previous, ...patch }));
  function savePageSetup() {
    try {
      if ((verifySheetCurrent(true), printPreview.error))
        throw Error(printPreview.error);
      const updatedWorkbook = applyPrintLayoutChange(
        props.workbook,
        props.sheetId,
        printPreview.layout,
        {
          readOnly: props.readOnly,
          expectedSnapshot: openingState.snapshot,
        },
      );
      (updatedWorkbook !== props.workbook &&
        props.onWorkbookChange(updatedWorkbook, "Update page layout"),
        props.onClose());
    } catch (error) {
      setErrorMessage(
        error instanceof Error ? error.message : "Could not save page setup.",
      );
    }
  }
  function applyTheme() {
    try {
      verifySheetCurrent(true);
      const rangeAddress = parseRangeAddress(props.range);
      if (
        (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
          (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
        500
      )
        throw new Error("Select at most 500 cells to apply a theme.");
      const themeColors =
          themeId === "mint"
            ? ["#548235", "#E2EFDA"]
            : themeId === "slate"
              ? ["#44546A", "#D6DCE4"]
              : themeId === "amber"
                ? ["#9C6500", "#FFF2CC"]
                : ["#4472C4", "#D9E2F3"],
        cells = { ...sheetOrThrow.cells };
      for (const cellAddress of assertRangeCellBudget(props.range)) {
        const bold =
          parseRangeAddress(cellAddress).firstRow === rangeAddress.firstRow;
        cells[cellAddress] = {
          ...cells[cellAddress],
          raw: cells[cellAddress]?.raw ?? null,
          style: {
            ...cells[cellAddress]?.style,
            fontFamily: "Calibri",
            fill: bold ? themeColors[0] : themeColors[1],
            color: bold ? "#FFFFFF" : "#242424",
            bold: bold,
          },
        };
      }
      (props.onWorkbookChange(
        {
          ...props.workbook,
          sheets: props.workbook.sheets.map((sheet) =>
            sheet.id === sheetOrThrow.id ? { ...sheet, cells: cells } : sheet,
          ),
        },
        `Apply ${themeId} theme to selection`,
      ),
        props.onClose());
    } catch (error) {
      setErrorMessage(
        error instanceof Error ? error.message : "Theme could not be applied.",
      );
    }
  }
  async function openPrintPreview() {
    if (abortRef.current) return;
    let previewWindow = null;
    const abortController = new AbortController();
    ((abortRef.current = abortController),
      setIsBusy(true),
      setErrorMessage(""));
    try {
      if ((verifySheetCurrent(), printPreview.error))
        throw Error(printPreview.error);
      const html = renderPrintDocument(
        props.workbook,
        props.sheetId,
        printPreview.layout,
      ).html;
      if (((previewWindow = window.open("", "_blank")), !previewWindow))
        throw Error(
          "Allow the print preview window in your browser, then try again.",
        );
      const printAreaElement = printAreaRef.current;
      await waitForPrintAssets(printAreaElement, {
        signal: abortController.signal,
      });
      const layoutIssues = measurePrintLayoutIssues(printAreaElement);
      (setPrintIssues(layoutIssues), summarizePrintIssues(layoutIssues));
      const previewInspection = await printDocumentInWindow(
        previewWindow,
        html,
        {
          signal: abortController.signal,
          assertCurrent: () => {
            if ((verifySheetCurrent(), abortController.signal.aborted))
              throw Error("Print preparation was cancelled.");
            if (layoutSignatureRef.current !== layoutSignature)
              throw Error("Page settings changed. Run the print check again.");
          },
        },
      );
      setPrintIssues(previewInspection);
    } catch (error) {
      (previewWindow && !previewWindow.closed && previewWindow.close(),
        abortController.signal.aborted ||
          setErrorMessage(
            error instanceof Error
              ? error.message
              : "Print preview could not open.",
          ));
    } finally {
      ((abortRef.current = null), setIsBusy(false));
    }
  }
  async function createPdf() {
    if (abortRef.current) return;
    const pdfAbort = new AbortController();
    ((abortRef.current = pdfAbort),
      setIsBusy(true),
      setProgressMessage("Preparing PDF…"),
      setErrorMessage(""),
      setPdfDocument(null));
    try {
      if ((verifySheetCurrent(), printPreview.error))
        throw Error(printPreview.error);
      await waitForPrintAssets(printAreaRef.current, {
        signal: pdfAbort.signal,
      });
      const pdfLayoutIssues = measurePrintLayoutIssues(printAreaRef.current);
      (setPrintIssues(pdfLayoutIssues), summarizePrintIssues(pdfLayoutIssues));
      const pdfResult = await renderSchedulePdf(props.workbook, props.sheetId, {
        layout: printPreview.layout,
        signal: pdfAbort.signal,
        assertCurrent: () => {
          if (
            (verifySheetCurrent(),
            layoutSignatureRef.current !== layoutSignature)
          )
            throw Error(
              "Page settings changed. Create the PDF again from the updated preview.",
            );
        },
        onProgress: (completedPages, totalPages) =>
          setProgressMessage(
            `Creating PDF: ${completedPages} of ${totalPages} pages…`,
          ),
      });
      if (pdfAbort.signal.aborted) return;
      (setPrintIssues(pdfResult.inspection),
        setPdfDocument({
          url: URL.createObjectURL(
            new Blob([new Uint8Array(pdfResult.bytes)], {
              type: "application/pdf",
            }),
          ),
          filename: pdfResult.filename,
          pages: pdfResult.pageCount,
        }));
    } catch (error) {
      pdfAbort.signal.aborted ||
        setErrorMessage(
          error instanceof Error
            ? error.message
            : "PDF creation failed. Try browser printing.",
        );
    } finally {
      ((abortRef.current = null), setIsBusy(false), setProgressMessage(""));
    }
  }
  function closeDialog() {
    (abortRef.current?.abort(), props.onClose());
  }
  return i.jsx("div", {
    className: "dg-backdrop",
    onKeyDown: (event) => {
      event.key === "Escape" && closeDialog();
    },
    children: i.jsxs("section", {
      className: `dg-dialog ${isThemeMode ? "" : "dg-page-dialog"}`,
      role: "dialog",
      "aria-modal": "true",
      "aria-label": isThemeMode
        ? "Workbook themes"
        : "Page setup and print preview",
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("span", {
              children: [
                isThemeMode
                  ? i.jsx(Palette, { size: 19 })
                  : i.jsx(FileText, { size: 19 }),
                isThemeMode ? "Themes" : "Page Setup",
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close page setup"),
              onClick: closeDialog,
              lang: language("Close page setup"),
              "data-ui-label": "Close page setup",
              children: i.jsx(X, { size: 17 }),
            }),
          ],
        }),
        isThemeMode
          ? i.jsxs("div", {
              className: "dg-form",
              children: [
                i.jsx("p", {
                  className: "dg-muted",
                  children: i.jsx(LocalizedText, {
                    children:
                      "Apply a coordinated font and header/body colors to the selected cells. Values, formulas, notes, and review marks are preserved.",
                  }),
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Theme" }),
                    i.jsxs("select", {
                      "aria-label": text("Theme"),
                      value: themeId,
                      onChange: (event) => setThemeId(event.target.value),
                      lang: language("Theme"),
                      "data-ui-label": "Theme",
                      children: [
                        i.jsx("option", {
                          value: "office",
                          children: "Office Blue",
                        }),
                        i.jsx("option", { value: "mint", children: "Mint" }),
                        i.jsx("option", { value: "slate", children: "Slate" }),
                        i.jsx("option", { value: "amber", children: "Amber" }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("p", {
                  children: [
                    i.jsx(LocalizedText, { children: "Selection: " }),
                    props.range,
                  ],
                }),
                errorMessage &&
                  i.jsx("p", {
                    role: "alert",
                    className: "dg-error",
                    children: errorMessage,
                  }),
              ],
            })
          : i.jsxs("div", {
              className: "dg-page-body",
              children: [
                i.jsxs("div", {
                  className: "dg-form dg-page-settings",
                  children: [
                    i.jsxs("div", {
                      className: "dg-two",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Orientation" }),
                            i.jsxs("select", {
                              "aria-label": text("Page orientation"),
                              value: layout.orientation,
                              onChange: (event) =>
                                updateLayout({
                                  orientation: event.target.value,
                                }),
                              lang: language("Page orientation"),
                              "data-ui-label": "Page orientation",
                              children: [
                                i.jsx("option", {
                                  value: "portrait",
                                  children: "Portrait",
                                }),
                                i.jsx("option", {
                                  value: "landscape",
                                  children: "Landscape",
                                }),
                              ],
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Paper" }),
                            i.jsxs("select", {
                              "aria-label": text("Paper size"),
                              value: layout.paper,
                              onChange: (event) =>
                                updateLayout({ paper: event.target.value }),
                              lang: language("Paper size"),
                              "data-ui-label": "Paper size",
                              children: [
                                i.jsx("option", {
                                  value: "letter",
                                  children: "Letter",
                                }),
                                i.jsx("option", {
                                  value: "a4",
                                  children: "A4",
                                }),
                                i.jsx("option", {
                                  value: "legal",
                                  children: "Legal",
                                }),
                              ],
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Print area" }),
                        i.jsx("input", {
                          "aria-label": text("Print area"),
                          value: layout.printArea,
                          onChange: (event) =>
                            updateLayout({
                              printArea: event.target.value.toUpperCase(),
                            }),
                          lang: language("Print area"),
                          "data-ui-label": "Print area",
                        }),
                      ],
                    }),
                    i.jsx("button", {
                      className: "dg-inline",
                      onClick: () => updateLayout({ printArea: props.range }),
                      children: i.jsx(LocalizedText, {
                        children: "Use current selection",
                      }),
                    }),
                    i.jsxs("fieldset", {
                      children: [
                        i.jsx("legend", {
                          children: i.jsx(LocalizedText, {
                            children: "Repeat print titles",
                          }),
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Rows to repeat at top",
                            }),
                            i.jsx("input", {
                              "aria-label": text("Rows to repeat at top"),
                              placeholder: "1:2",
                              value: layout.repeatRows ?? "",
                              onChange: (event) =>
                                updateLayout({
                                  repeatRows: event.target.value,
                                }),
                              lang: language("Rows to repeat at top"),
                              "data-ui-label": "Rows to repeat at top",
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Columns to repeat at left",
                            }),
                            i.jsx("input", {
                              "aria-label": text("Columns to repeat at left"),
                              placeholder: text("A:B"),
                              value: layout.repeatColumns ?? "",
                              onChange: (event) =>
                                updateLayout({
                                  repeatColumns:
                                    event.target.value.toUpperCase(),
                                }),
                              lang: language("Columns to repeat at left"),
                              "data-ui-label": "Columns to repeat at left",
                            }),
                          ],
                        }),
                        i.jsx("p", {
                          className: "dg-muted",
                          children: i.jsx(LocalizedText, {
                            children:
                              "Titles may be outside the print area. Hidden title rows/columns are omitted. Up to 100 rows and 50 columns.",
                          }),
                        }),
                      ],
                    }),
                    i.jsxs("fieldset", {
                      children: [
                        i.jsx("legend", {
                          children: i.jsx(LocalizedText, {
                            children: "Manual page breaks",
                          }),
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Page order" }),
                            i.jsxs("select", {
                              "aria-label": text("Page order"),
                              value: layout.pageOrder ?? "overThenDown",
                              onChange: (event) =>
                                updateLayout({ pageOrder: event.target.value }),
                              lang: language("Page order"),
                              "data-ui-label": "Page order",
                              children: [
                                i.jsx("option", {
                                  value: "overThenDown",
                                  children: "Over, then down",
                                }),
                                i.jsx("option", {
                                  value: "downThenOver",
                                  children: "Down, then over",
                                }),
                              ],
                            }),
                          ],
                        }),
                        i.jsx("p", {
                          className: "dg-muted",
                          children: i.jsx(LocalizedText, {
                            children:
                              "Controls the order pages are numbered and printed. Page contents and the total page count stay the same.",
                          }),
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Break before rows",
                            }),
                            i.jsx("input", {
                              "aria-label": text("Break before rows"),
                              placeholder: "25, 50",
                              value: rowBreakText,
                              onChange: (event) =>
                                setRowBreakText(event.target.value),
                              lang: language("Break before rows"),
                              "data-ui-label": "Break before rows",
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Break before columns",
                            }),
                            i.jsx("input", {
                              "aria-label": text("Break before columns"),
                              placeholder: text("D, H"),
                              value: columnBreakText,
                              onChange: (event) =>
                                setColumnBreakText(
                                  event.target.value.toUpperCase(),
                                ),
                              lang: language("Break before columns"),
                              "data-ui-label": "Break before columns",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => {
                            (setRowBreakText(""), setColumnBreakText(""));
                          },
                          children: i.jsx(LocalizedText, {
                            children: "Reset manual breaks",
                          }),
                        }),
                        i.jsx("p", {
                          className: "dg-muted",
                          children: i.jsx(LocalizedText, {
                            children:
                              "Enter row numbers or column letters separated by commas. Edit a boundary to move it; remove an entry to delete it. Fit to pages ignores manual breaks.",
                          }),
                        }),
                      ],
                    }),
                    i.jsxs("fieldset", {
                      children: [
                        i.jsx("legend", {
                          children: i.jsx(LocalizedText, {
                            children: "Margins (inches)",
                          }),
                        }),
                        i.jsx("div", {
                          className: "dg-two",
                          children: [
                            "top",
                            "bottom",
                            "left",
                            "right",
                            "header",
                            "footer",
                          ].map((marginName) =>
                            i.jsxs(
                              "label",
                              {
                                children: [
                                  marginName,
                                  i.jsx(LocalizedNumberField, {
                                    min: 0,
                                    max: 3,
                                    step: 0.05,
                                    "aria-label": `${marginName} margin`,
                                    value: layout.margins[marginName],
                                    onValueChange: (nextValue) =>
                                      updateLayout({
                                        margins: {
                                          ...layout.margins,
                                          [marginName]: Number(nextValue),
                                        },
                                      }),
                                  }),
                                ],
                              },
                              marginName,
                            ),
                          ),
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      className: "dg-check",
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: layout.fitToPage,
                          onChange: (event) =>
                            updateLayout({ fitToPage: event.target.checked }),
                        }),
                        i.jsx(LocalizedText, { children: "Fit to pages" }),
                      ],
                    }),
                    layout.fitToPage
                      ? i.jsxs("div", {
                          className: "dg-two",
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Pages wide",
                                }),
                                i.jsx(LocalizedNumberField, {
                                  min: 0,
                                  max: 10,
                                  "aria-label": text("Pages wide"),
                                  value: layout.fitToWidth,
                                  onValueChange: (widthValue) =>
                                    updateLayout({
                                      fitToWidth: Number(widthValue),
                                    }),
                                  lang: language("Pages wide"),
                                  "data-ui-label": "Pages wide",
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Pages tall",
                                }),
                                i.jsx(LocalizedNumberField, {
                                  min: 0,
                                  max: 10,
                                  "aria-label": text("Pages tall"),
                                  value: layout.fitToHeight,
                                  onValueChange: (heightValue) =>
                                    updateLayout({
                                      fitToHeight: Number(heightValue),
                                    }),
                                  lang: language("Pages tall"),
                                  "data-ui-label": "Pages tall",
                                }),
                              ],
                            }),
                          ],
                        })
                      : i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Scale (%)" }),
                            i.jsx(LocalizedNumberField, {
                              min: 10,
                              max: 400,
                              "aria-label": text("Print scale"),
                              value: layout.scale,
                              onValueChange: (scaleValue) =>
                                updateLayout({ scale: Number(scaleValue) }),
                              lang: language("Print scale"),
                              "data-ui-label": "Print scale",
                            }),
                          ],
                        }),
                    i.jsx("p", {
                      className: "dg-muted",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Zero pages means automatic. Fit to pages reduces the content as needed.",
                      }),
                    }),
                    i.jsxs("label", {
                      className: "dg-check",
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: layout.gridlines,
                          onChange: (event) =>
                            updateLayout({ gridlines: event.target.checked }),
                        }),
                        i.jsx(LocalizedText, { children: "Print gridlines" }),
                      ],
                    }),
                    i.jsxs("label", {
                      className: "dg-check",
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: layout.headings,
                          onChange: (event) =>
                            updateLayout({ headings: event.target.checked }),
                        }),
                        i.jsx(LocalizedText, {
                          children: "Print row and column headings",
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Header" }),
                        i.jsx("textarea", {
                          "aria-label": text("Print header"),
                          rows: 3,
                          maxLength: 255,
                          value: layout.header,
                          onChange: (event) =>
                            updateLayout({ header: event.target.value }),
                          lang: language("Print header"),
                          "data-ui-label": "Print header",
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Footer" }),
                        i.jsx("textarea", {
                          "aria-label": text("Print footer"),
                          rows: 2,
                          maxLength: 255,
                          value: layout.footer,
                          onChange: (event) =>
                            updateLayout({ footer: event.target.value }),
                          lang: language("Print footer"),
                          "data-ui-label": "Print footer",
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      className: "dg-muted",
                      children: i.jsx(LocalizedText, {
                        children:
                          "&P page · &N total pages · &A sheet name · && literal ampersand",
                      }),
                    }),
                    i.jsx("button", {
                      type: "button",
                      onClick: () => updateLayout(defaultPrintHeaderFooter),
                      children: i.jsx(LocalizedText, {
                        children: "Insert workbook identifiers",
                      }),
                    }),
                    i.jsx("p", {
                      className: "dg-muted",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Replaces the header and footer with workbook name, engagement, period, revision and page numbers. Fields: &[Workbook], &[Engagement], &[Period], &[Revision]. Missing properties display a “not set” label. Values come from this workbook; revision is local history, not an approval or certification.",
                      }),
                    }),
                    !!sheetOrThrow.merges?.length &&
                      i.jsx("p", {
                        className: "dg-muted",
                        children: i.jsx(LocalizedText, {
                          children:
                            "Merged areas crossing a page or print-area boundary repeat their top-left content in each clipped fragment.",
                        }),
                      }),
                    errorMessage &&
                      i.jsx("p", {
                        role: "alert",
                        className: "dg-error",
                        children: errorMessage,
                      }),
                  ],
                }),
                i.jsxs("div", {
                  className: "dg-print-preview",
                  ref: printAreaRef,
                  children: [
                    i.jsx("style", { children: printPaperStyles }),
                    i.jsxs("div", {
                      className: "dg-preview-label",
                      children: [
                        i.jsx(LocalizedText, { children: "PRINT PREVIEW · " }),
                        printPreview.pages.length,
                        " ",
                        printPreview.pages.length === 1 ? "PAGE" : "PAGES",
                        " ·",
                        " ",
                        Math.round(printPreview.scale * 100),
                        "%",
                      ],
                    }),
                    i.jsxs("div", {
                      className: "dg-print-check",
                      "aria-label": text("Print output checks"),
                      lang: language("Print output checks"),
                      "data-ui-label": "Print output checks",
                      children: [
                        i.jsx("p", {
                          children: i.jsx(LocalizedText, {
                            children:
                              "Create PDF produces a local image-based snapshot at 192 dpi. Text in this download is not searchable, selectable or accessibility-tagged. All pages are checked before the download is offered.",
                          }),
                        }),
                        i.jsxs("details", {
                          children: [
                            i.jsx("summary", {
                              children: i.jsx(LocalizedText, {
                                children: "Browser printing settings",
                              }),
                            }),
                            i.jsx("p", {
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Print / Save as PDF uses your browser’s PDF destination. Match this paper size and orientation, use 100% scale and no browser margins, and turn off browser headers/footers. Keep background graphics on for cell fills.",
                              }),
                            }),
                            i.jsx("p", {
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Printing rechecks the final document. A print dialog opening does not confirm that a PDF was saved.",
                              }),
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: onClick,
                          disabled:
                            isChecking || isBusy || !!printPreview.error,
                          children: isChecking
                            ? "Checking…"
                            : "Check print output",
                        }),
                        printIssues &&
                          i.jsxs("div", {
                            role: "status",
                            children: [
                              i.jsx("p", {
                                children: printIssues.issues.length
                                  ? printIssues.issues.filter(
                                      (issue) => issue.severity === "error",
                                    ).length +
                                    " blocking issue(s), " +
                                    printIssues.issues.filter(
                                      (issue) => issue.severity === "warning",
                                    ).length +
                                    " warning(s)."
                                  : "Print check passed: " +
                                    printIssues.pages +
                                    " page(s); no measured clipping or missing images.",
                              }),
                              printIssues.issues.length > 0 &&
                                i.jsx("ul", {
                                  children: printIssues.issues
                                    .slice(0, 100)
                                    .map((issue, issueIndex) =>
                                      i.jsxs(
                                        "li",
                                        {
                                          children: [
                                            issue.severity === "error"
                                              ? "Fix: "
                                              : "Review: ",
                                            issue.message,
                                          ],
                                        },
                                        issueIndex,
                                      ),
                                    ),
                                }),
                              printIssues.issues.length > 100 &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Showing the first 100 of ",
                                    }),
                                    printIssues.issues.length,
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children:
                                        "issues. Correct the affected sizes and check again.",
                                    }),
                                  ],
                                }),
                            ],
                          }),
                        progressMessage &&
                          i.jsx("p", {
                            role: "status",
                            children: progressMessage,
                          }),
                        pdfDocument &&
                          i.jsxs("p", {
                            role: "status",
                            children: [
                              i.jsxs("a", {
                                href: pdfDocument.url,
                                download: pdfDocument.filename,
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Download PDF (",
                                  }),
                                  pdfDocument.pages,
                                  " ",
                                  pdfDocument.pages === 1 ? "page" : "pages",
                                  ")",
                                ],
                              }),
                              " ",
                              "· ",
                              pdfDocument.filename,
                            ],
                          }),
                      ],
                    }),
                    printPreview.notices?.map((notice) =>
                      i.jsx(
                        "p",
                        { className: "dg-muted", children: notice },
                        notice,
                      ),
                    ),
                    printPreview.ruleStyles.errors().length > 0 &&
                      i.jsxs("p", {
                        role: "alert",
                        className: "dg-error",
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Conditional formatting: ",
                          }),
                          printPreview.ruleStyles.errors().length,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "evaluation error(s). ",
                          }),
                          printPreview.ruleStyles.errors()[0].message,
                        ],
                      }),
                    printPreview.error
                      ? i.jsx("p", {
                          className: "dg-error",
                          role: "alert",
                          children: printPreview.error,
                        })
                      : printPreview.pages.map((page, pageIndex) =>
                          i.jsx(
                            PrintPagePreview,
                            {
                              workbook: props.workbook,
                              sheet: sheetOrThrow,
                              layout: printPreview.layout,
                              page: page,
                              index: pageIndex,
                              total: printPreview.pages.length,
                              scale: printPreview.scale,
                              calc: calc,
                              ruleStyles: printPreview.ruleStyles,
                            },
                            pageIndex,
                          ),
                        ),
                  ],
                }),
              ],
            }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: closeDialog,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            !isThemeMode &&
              i.jsx("button", {
                onClick: createPdf,
                disabled: !!printPreview.error || isBusy || isChecking,
                children: i.jsx(LocalizedText, { children: "Create PDF" }),
              }),
            !isThemeMode &&
              i.jsxs("button", {
                onClick: openPrintPreview,
                disabled: !!printPreview.error || isBusy || isChecking,
                children: [
                  i.jsx(Printer, { size: 14 }),
                  isBusy && !progressMessage
                    ? "Preparing print…"
                    : "Print / Save as PDF",
                ],
              }),
            i.jsx("button", {
              className: "dg-primary",
              onClick: isThemeMode ? applyTheme : savePageSetup,
              disabled:
                props.readOnly ||
                isBusy ||
                isChecking ||
                !!sheetOrThrow.protection?.enabled ||
                (!isThemeMode && !!printPreview.error),
              children: isThemeMode ? "Apply theme" : "Save page setup",
            }),
          ],
        }),
      ],
    }),
  });
}
const functionArgumentSignatures = {
    PERMUTATIONA: "number, number_chosen",
    PERMUT: "number, number_chosen",
    VARPA: "value1, [value2, ...]",
    VARA: "value1, [value2, ...]",
    STDEVPA: "value1, [value2, ...]",
    STDEVA: "value1, [value2, ...]",
    MINA: "value1, [value2, ...]",
    MAXA: "value1, [value2, ...]",
    "MODE.MULT": "number1, [number2], …",
    TRIMMEAN: "array, percent",
    "SKEW.P": "number1, [number2], …",
    SKEW: "number1, [number2], …",
    "Z.TEST": "array, x, [sigma]",
    "T.TEST": "array1, array2, tails, type",
    PEARSON: "array1, array2",
    GROWTH: "known_ys, [known_xs], [new_xs], [const]",
    TREND: "known_ys, [known_xs], [new_xs], [const]",
    LOGEST: "known_ys, [known_xs], [const], [stats]",
    LINEST: "known_ys, [known_xs], [const], [stats]",
    CONVERT: "number, from_unit, to_unit",
    SLOPE: "known_ys, known_xs",
    RSQ: "known_ys, known_xs",
    STEYX: "known_ys, known_xs",
    LET: "name1, value1, [name2, value2], …, calculation",
    LAMBDA: "[parameter1, parameter2, …], calculation",
    MODE: "number1, [number2], …",
    PERCENTILE: "array, k",
    QUARTILE: "array, quart",
    RANK: "number, ref, [order]",
    LOOKUP: "lookup_value, lookup_vector, [result_vector]",
    CHOOSE: "index_num, value1, [value2], …",
    DSUM: "database, field, criteria",
    DAVERAGE: "database, field, criteria",
    DCOUNT: "database, [field], criteria",
    DCOUNTA: "database, [field], criteria",
    DMAX: "database, field, criteria",
    DMIN: "database, field, criteria",
    DPRODUCT: "database, field, criteria",
    DSTDEV: "database, field, criteria",
    DSTDEVP: "database, field, criteria",
    DVAR: "database, field, criteria",
    DVARP: "database, field, criteria",
    DGET: "database, field, criteria",
    AGGREGATE: "function_num, options, ref1 or array, [ref2 or k], …",
    SUBTOTAL: "function_num, ref1, [ref2], …",
    ISFORMULA: "reference",
    FORMULATEXT: "reference",
    SUM: "number1, [number2], …",
    AVERAGE: "number1, [number2], …",
    COUNT: "value1, [value2], …",
    COUNTA: "value1, [value2], …",
    IF: "logical_test, value_if_true, [value_if_false]",
    IFERROR: "value, value_if_error",
    IFNA: "value, value_if_na",
    IFS: "logical_test1, value_if_true1, …",
    SWITCH: "expression, value1, result1, …, [default]",
    SUMIF: "range, criteria, [sum_range]",
    SUMIFS: "sum_range, criteria_range1, criteria1, …",
    COUNTIF: "range, criteria",
    COUNTIFS: "criteria_range1, criteria1, …",
    MAXIFS: "max_range, criteria_range1, criteria1, …",
    MINIFS: "min_range, criteria_range1, criteria1, …",
    XLOOKUP:
      "lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]",
    XMATCH: "lookup_value, lookup_array, [match_mode], [search_mode]",
    VLOOKUP: "lookup_value, table_array, col_index_num, [range_lookup]",
    HLOOKUP: "lookup_value, table_array, row_index_num, [range_lookup]",
    INDEX: "array, row_num, [column_num]",
    MATCH: "lookup_value, lookup_array, [match_type]",
    FILTER: "array, include, [if_empty]",
    SORT: "array, [sort_index], [sort_order], [by_col]",
    SORTBY: "array, by_array1, [sort_order1], …",
    UNIQUE: "array, [by_col], [exactly_once]",
    SEQUENCE: "rows, [columns], [start], [step]",
    TRANSPOSE: "array",
    TAKE: "array, rows, [columns]",
    DROP: "array, rows, [columns]",
    CHOOSECOLS: "array, col_num1, …",
    CHOOSEROWS: "array, row_num1, …",
    HSTACK: "array1, [array2], …",
    VSTACK: "array1, [array2], …",
    TOCOL: "array, [ignore], [scan_by_column]",
    TOROW: "array, [ignore], [scan_by_column]",
    EXPAND: "array, rows, [columns], [pad_with]",
    TEXTJOIN: "delimiter, ignore_empty, text1, …",
    CONCAT: "text1, [text2], …",
    TEXTSPLIT:
      "text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]",
    TEXTBEFORE:
      "text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]",
    TEXTAFTER:
      "text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]",
    ROUND: "number, num_digits",
    LEFT: "text, [num_chars]",
    RIGHT: "text, [num_chars]",
    MID: "text, start_num, num_chars",
    LEN: "text",
    TRIM: "text",
    TEXT: "value, format_text",
    DATE: "year, month, day",
    TODAY: "",
    NOW: "",
    PI: "",
    TRUE: "",
    FALSE: "",
    RAND: "",
    RANDBETWEEN: "bottom, top",
    PMT: "rate, nper, pv, [fv], [type]",
    IPMT: "rate, per, nper, pv, [fv], [type]",
    PPMT: "rate, per, nper, pv, [fv], [type]",
    DB: "cost, salvage, life, period, [month]",
    PRICE: "settlement, maturity, rate, yld, redemption, frequency, [basis]",
    YIELD: "settlement, maturity, rate, pr, redemption, frequency, [basis]",
    ACCRINT:
      "issue, first_interest, settlement, rate, par, frequency, [basis], [calc_method]",
    ACCRINTM: "issue, settlement, rate, par, [basis]",
    MIRR: "values, finance_rate, reinvest_rate",
    PV: "rate, nper, pmt, [fv], [type]",
    FV: "rate, nper, pmt, [pv], [type]",
    NPER: "rate, pmt, pv, [fv], [type]",
    RATE: "nper, pmt, pv, [fv], [type], [guess]",
    NPV: "rate, value1, [value2], …",
    IRR: "values, [guess]",
    XNPV: "rate, values, dates",
    XIRR: "values, dates, [guess]",
    "STDEV.P": "number1, [number2], …",
    "VAR.P": "number1, [number2], …",
    "PERCENTILE.INC": "array, k",
    "PERCENTILE.EXC": "array, k",
    "QUARTILE.INC": "array, quart",
    "QUARTILE.EXC": "array, quart",
    "RANK.EQ": "number, ref, [order]",
    "RANK.AVG": "number, ref, [order]",
    "MODE.SNGL": "number1, [number2], …",
    MDETERM: "array",
    MINVERSE: "array",
    SLN: "cost, salvage, life",
    SYD: "cost, salvage, life, per",
    DDB: "cost, salvage, life, period, [factor]",
  },
  functionDescriptions = {
    PERMUTATIONA: "Counts ordered selections with replacement",
    PERMUT: "Counts ordered selections without replacement",
    VARPA:
      "Calculates population variance including logical values and text as zero",
    VARA: "Calculates sample variance including logical values and text as zero",
    STDEVPA:
      "Calculates population deviation including logical values and text as zero",
    STDEVA:
      "Calculates sample deviation including logical values and text as zero",
    MINA: "Returns the minimum including logical values and text as zero",
    MAXA: "Returns the maximum including logical values and text as zero",
    "MODE.MULT": "Returns all most-frequent numeric values as a vertical array",
    TRIMMEAN: "Averages data after removing equal counts from both tails",
    "SKEW.P": "Calculates population skewness",
    SKEW: "Calculates adjusted sample skewness",
    "Z.TEST":
      "Returns a one-tailed normal-test probability using known or sample deviation",
    "T.TEST": "Returns paired, pooled or Welch t-test probabilities",
    PEARSON:
      "Calculates the signed Pearson correlation of paired numeric observations",
    GROWTH: "Predicts positive values from an exponential regression",
    TREND:
      "Predicts values from a linear regression with one or more predictors",
    LOGEST: "Fits exponential coefficients with optional regression statistics",
    LINEST:
      "Fits linear coefficients and optionally returns complete regression statistics",
    CONVERT:
      "Converts documented quantity units; unit names and prefixes are case-sensitive",
    SLOPE: "Calculates the slope from paired numeric observations",
    RSQ: "Returns the squared correlation of paired numeric observations",
    STEYX: "Calculates the standard error of the fitted regression",
    LET: "Calculates local named values once and uses them in a final calculation",
    LAMBDA:
      "Creates a reusable formula; call it immediately or save it in Name Manager",
    MODE: "Returns the most frequent number; ties use the smallest value",
    PERCENTILE: "Returns an inclusive percentile with linear interpolation",
    QUARTILE: "Returns an inclusive quartile of numeric values",
    RANK: "Returns a number’s rank with equal values sharing a rank",
    LOOKUP: "Finds the last ascending vector key at or below the lookup value",
    CHOOSE: "Selects one value or reference by a one-based index",
    DSUM: "Adds numeric fields in records matching the criteria range",
    DAVERAGE: "Averages numeric fields in records matching the criteria range",
    DCOUNT:
      "Counts numeric matching fields; omit field to count matching records",
    DCOUNTA:
      "Counts nonempty matching fields; omit field to count matching records",
    DMAX: "Largest numeric field in matching records",
    DMIN: "Smallest numeric field in matching records",
    DPRODUCT: "Multiplies numeric fields in matching records",
    DSTDEV: "Sample standard deviation of matching numeric fields",
    DSTDEVP: "Population standard deviation of matching numeric fields",
    DVAR: "Sample variance of matching numeric fields",
    DVARP: "Population variance of matching numeric fields",
    DGET: "Returns one matching field; no matches or multiple matches are errors",
    AGGREGATE:
      "Uses 19 summary operations; options 0–7 control hidden rows, nested totals, and errors; 14–19 require scalar k",
    SUBTOTAL:
      "Summarizes references, excluding nested subtotals and supplied filter rows; 101–111 also exclude hidden rows",
    ISFORMULA:
      "Checks whether referenced cells store formulas, including formulas with errors",
    FORMULATEXT: "Returns stored formula text from the first referenced cell",
    SUM: "Adds numbers and ranges",
    AVERAGE: "Returns the arithmetic mean",
    IF: "Chooses a result when a condition is true or false",
    COUNT: "Counts numeric values",
    COUNTA: "Counts nonempty values",
    SUMIFS: "Adds values meeting every condition",
    COUNTIFS: "Counts rows meeting every condition",
    MAXIFS: "Largest value meeting every condition",
    MINIFS: "Smallest value meeting every condition",
    XLOOKUP: "Finds a match and returns its corresponding value or row",
    XMATCH: "Returns the position of a matching value",
    IFS: "Returns the first result whose condition is true",
    SWITCH: "Matches an expression to a list of results",
    FILTER: "Returns rows or columns that meet a condition",
    SORT: "Sorts an array by a row or column",
    SORTBY: "Sorts an array using corresponding values",
    UNIQUE: "Returns distinct rows or columns",
    SEQUENCE: "Creates a grid of sequential numbers",
    TRANSPOSE: "Exchanges the rows and columns of an array",
    TEXTJOIN: "Joins text with a delimiter",
    CONCAT: "Joins text from values and ranges",
    TEXTSPLIT: "Splits text into rows and columns",
    TEXTBEFORE: "Returns text before a delimiter",
    TEXTAFTER: "Returns text after a delimiter",
    TAKE: "Keeps rows or columns from an array edge",
    DROP: "Removes rows or columns from an array edge",
    HSTACK: "Combines arrays horizontally",
    VSTACK: "Combines arrays vertically",
    EXPAND: "Pads an array to the requested dimensions",
    NPV: "Discounts periodic cash flows",
    IRR: "Finds the return rate of periodic cash flows",
    XNPV: "Discounts cash flows on specified dates",
    XIRR: "Finds the return rate of dated cash flows",
    PMT: "Calculates a constant loan or annuity payment",
    IPMT: "Calculates the interest portion of a periodic payment",
    PPMT: "Calculates the principal portion of a periodic payment",
    DB: "Calculates fixed declining-balance depreciation",
    PRICE:
      "Calculates the clean price per 100 face value of a regular-coupon security",
    YIELD:
      "Solves a regular-coupon security yield with a checked price residual",
    ACCRINT: "Calculates interest accrued across coupon periods",
    ACCRINTM: "Calculates interest accrued to maturity",
    MIRR: "Calculates modified return using explicit finance and reinvestment rates",
  };
function getFormulaCompletions(value, caret, workbook, sheetId) {
  if (!value.startsWith("=")) return { items: [] };
  const textBeforeCaret = value.slice(0, caret);
  let isInString = false;
  const callStack = [];
  for (let charIndex = 1; charIndex < textBeforeCaret.length; charIndex++) {
    if (textBeforeCaret[charIndex] === '"') {
      if (isInString && textBeforeCaret[charIndex + 1] === '"') {
        charIndex++;
        continue;
      }
      isInString = !isInString;
      continue;
    }
    if (!isInString)
      if (textBeforeCaret[charIndex] === "(") {
        const functionName =
          /([A-Za-z][A-Za-z0-9_.]*)\s*$/
            .exec(textBeforeCaret.slice(0, charIndex))?.[1]
            ?.toUpperCase() ?? "";
        callStack.push({ name: functionName, argumentIndex: 0 });
      } else
        textBeforeCaret[charIndex] === ")"
          ? callStack.pop()
          : textBeforeCaret[charIndex] === "," &&
            callStack.length &&
            callStack[callStack.length - 1].argumentIndex++;
  }
  if (isInString) return { items: [] };
  const installedExtensionFunctions = listInstalledExtensionFunctions(workbook),
    innermostCall = callStack.at(-1),
    installedExtensionFunction = installedExtensionFunctions.find(
      (installedExtensionFunction) =>
        installedExtensionFunction.name === innermostCall?.name,
    ),
    matchedDefinedName = innermostCall?.name
      ? [
          ...(workbook.definedNames ?? []).filter(
            (definedName) => definedName.scopeSheetId === sheetId,
          ),
          ...(workbook.definedNames ?? []).filter(
            (definedName) => !definedName.scopeSheetId,
          ),
        ].find(
          (definedName) =>
            definedName.name.toUpperCase() === innermostCall.name,
        )
      : undefined,
    definedNameParameters = matchedDefinedName
      ? parseFormulaTree(matchedDefinedName.refersTo)
      : undefined,
    hint =
      innermostCall?.name &&
      (builtinFunctionNames.includes(innermostCall.name) ||
        definedNameParameters ||
        installedExtensionFunction)
        ? {
            ...innermostCall,
            signature:
              (installedExtensionFunction
                ? installedExtensionFunction.fn.parameters
                    .map((parameter) => parameter.name)
                    .join(", ")
                : definedNameParameters
                  ? definedNameParameters.join(", ")
                  : functionArgumentSignatures[innermostCall.name]) ??
              "value1, [value2], …",
          }
        : undefined,
    identifierMatch = /([A-Za-z_][A-Za-z0-9_.]*(?:\[[^\]]*)?)$/.exec(
      textBeforeCaret,
    ),
    isAtFormulaStart = /^=\s*$/.test(textBeforeCaret);
  if (!identifierMatch && !isAtFormulaStart) return { items: [], hint: hint };
  const typedPrefix = identifierMatch?.[1] ?? "",
    start = identifierMatch ? caret - typedPrefix.length : caret;
  if (start > 0 && /[A-Za-z0-9_$!']/.test(value[start - 1]))
    return { items: [], hint: hint };
  const prefixLower = typedPrefix.toLowerCase(),
    candidates = [],
    addCandidate = (candidateLabel, candidateKind, insertText, detail) => {
      candidateLabel.toLowerCase().startsWith(prefixLower) &&
        candidates.push({
          label: candidateLabel,
          kind: candidateKind,
          insertText: insertText,
          start: start,
          end: caret,
          detail: detail,
        });
    },
    defaultFunctionNames = [
      "SUM",
      "AVERAGE",
      "IF",
      "XLOOKUP",
      "COUNT",
      "SUMIFS",
      "FILTER",
      "UNIQUE",
    ];
  for (const builtinName of typedPrefix
    ? builtinFunctionNames
    : defaultFunctionNames)
    addCandidate(
      builtinName,
      "function",
      `${builtinName}(`,
      `${builtinName}(${functionArgumentSignatures[builtinName] ?? "value1, …"})${functionDescriptions[builtinName] ? ` · ${functionDescriptions[builtinName]}` : ""}`,
    );
  const seenDefinedNames = new Set();
  for (const installedExtensionFunction of installedExtensionFunctions)
    addCandidate(
      installedExtensionFunction.name,
      "function",
      `${installedExtensionFunction.name}(`,
      `${installedExtensionFunction.name}(${installedExtensionFunction.fn.parameters.map((parameter) => parameter.name).join(", ")}) · ${installedExtensionFunction.entry.enabled ? "Custom function · manual refresh" : "Disabled extension"} · ${installedExtensionFunction.fn.description}`,
    );
  for (const definedName of [
    ...(workbook.definedNames ?? []).filter((O) => O.scopeSheetId === sheetId),
    ...(workbook.definedNames ?? []).filter((O) => !O.scopeSheetId),
  ]) {
    const definedNameKey = definedName.name.toLowerCase(),
      formulaTree = parseFormulaTree(definedName.refersTo);
    (seenDefinedNames.has(definedNameKey) ||
      addCandidate(
        definedName.name,
        "name",
        `${definedName.name}${formulaTree ? "(" : ""}`,
        formulaTree
          ? `${definedName.name}(${formulaTree.join(", ")})${definedName.comment ? ` · ${definedName.comment}` : ""}`
          : definedName.refersTo,
      ),
      seenDefinedNames.add(definedNameKey));
  }
  for (const sheet of workbook.sheets)
    for (const table of sheet.tables ?? []) {
      addCandidate(
        table.name,
        "table",
        `${table.name}[`,
        `${sheet.name}!${table.range}`,
      );
      for (const column of table.columns)
        addCandidate(
          `${table.name}[${column.name}]`,
          "table",
          `${table.name}[${column.name}]`,
          `Table column · ${column.name}`,
        );
    }
  return (
    typedPrefix &&
      candidates.sort((leftCandidate, rightCandidate) => {
        const candidateRank = (candidate) =>
          candidate.label.toLowerCase() === prefixLower
            ? -2
            : candidate.kind === "function" &&
                defaultFunctionNames.includes(candidate.label)
              ? defaultFunctionNames.indexOf(candidate.label)
              : 100;
        return (
          candidateRank(leftCandidate) - candidateRank(rightCandidate) ||
          leftCandidate.label.localeCompare(rightCandidate.label)
        );
      }),
    { items: candidates.slice(0, 12), hint: hint }
  );
}
function applyFormulaCompletion(value, completion) {
  const textAfterCaret = value.slice(completion.end),
    insertedText =
      completion.insertText.endsWith("(") && textAfterCaret.startsWith("(")
        ? completion.insertText.slice(0, -1)
        : completion.insertText;
  return {
    value: value.slice(0, completion.start) + insertedText + textAfterCaret,
    caret: completion.start + insertedText.length,
  };
}

export {
  functionDescriptions,
  PageSetupDialog,
  getFormulaCompletions,
  applyFormulaCompletion,
  functionArgumentSignatures,
};
