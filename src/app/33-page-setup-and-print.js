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

function PageSetupDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useRef(e);
  n.current = e;
  const o = ne.useRef(null),
    a = ne.useRef(null),
    [s, d] = ne.useState(null),
    [u, l] = ne.useState(false),
    [c, h] = ne.useState(false),
    [f, p] = ne.useState(""),
    [m, v] = ne.useState(null);
  (ne.useEffect(
    () => () => {
      m && URL.revokeObjectURL(m.url);
    },
    [m],
  ),
    ne.useEffect(() => () => a.current?.abort(), []));
  const [g] = ne.useState(() => ({
      snapshot: buildSnapshotGuard(e.workbook),
      sheetId: e.sheetId,
      range: e.range,
      sheet: structuredClone(
        e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
      ),
    })),
    sheetOrThrow =
      e.workbook.sheets.find((sheet) => sheet.id === e.sheetId) ?? g.sheet;
  function y(W = false) {
    const K = n.current;
    if (
      (verifyPrintSnapshotGuard(K.workbook, g.snapshot),
      K.sheetId !== g.sheetId || K.range !== g.range)
    )
      throw Error(
        "The worksheet or selection changed. Close and reopen page setup.",
      );
    if (W) {
      if (K.readOnly) throw Error("This workbook is read-only.");
      assertSheetUnprotected(
        K.workbook.sheets.find((sheet) => sheet.id === g.sheetId),
      );
    }
  }
  const [layout, E] = ne.useState(() =>
      structuredClone(
        sheetOrThrow.pageLayout ?? {
          ...createDefaultPageSetup(sheetOrThrow),
          ...defaultPrintHeaderFooter,
        },
      ),
    ),
    [A, k] = ne.useState(""),
    [F, I] = ne.useState("office"),
    [O, j] = ne.useState(() =>
      (sheetOrThrow.pageLayout?.rowBreaks ?? [])
        .map((K) => String(K + 1))
        .join(", "),
    ),
    [z, V] = ne.useState(() =>
      (sheetOrThrow.pageLayout?.columnBreaks ?? [])
        .map(columnIndexToLetters)
        .join(", "),
    );
  function D() {
    const W = (K, te) => {
      if (!K.trim()) return;
      const ae = K.split(",")
        .map((oe) => {
          const de = oe.trim();
          if (!(te === "row" ? /^[1-9]\d*$/ : /^[A-Za-z]{1,3}$/).test(de))
            throw Error(
              "Enter comma-separated row numbers or column letters for manual breaks.",
            );
          return te === "row"
            ? Number(de) - 1
            : parseRangeAddress(`${de}1`).firstCol;
        })
        .sort((oe, de) => oe - de);
      return validatePrintPageBreaks(
        ae,
        te === "row" ? sheetOrThrow.rowCount : sheetOrThrow.colCount,
      );
    };
    return { ...layout, rowBreaks: W(O, "row"), columnBreaks: W(z, "column") };
  }
  const B = /theme/i.test(e.mode),
    calc = ne.useMemo(
      () => calculateWorkbookWithSettings(e.workbook),
      [e.workbook],
    ),
    C = ne.useMemo(() => {
      try {
        if (!e.workbook.sheets.some((sheet) => sheet.id === e.sheetId))
          throw Error(
            "The worksheet changed or was deleted. Close and reopen page setup.",
          );
        return { ...buildPrintPages(e.workbook, e.sheetId, D()), error: "" };
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
    }, [e.workbook, e.sheetId, sheetOrThrow, layout, calc, O, z]),
    T = JSON.stringify([layout, O, z]),
    N = ne.useRef(T);
  ((N.current = T),
    ne.useEffect(() => {
      (d(null), k(""), v(null));
    }, [T, e.workbook, e.sheetId, e.range]));
  async function onClick() {
    const W = o.current;
    if (!(!W || C.error)) {
      (l(true), k(""));
      try {
        if ((y(), await waitForPrintAssets(W), y(), N.current !== T))
          throw Error("Page settings changed. Run the print check again.");
        d(measurePrintLayoutIssues(W));
      } catch (error) {
        k(error instanceof Error ? error.message : "Print check failed.");
      } finally {
        l(false);
      }
    }
  }
  const L = (W) => E((K) => ({ ...K, ...W }));
  function $() {
    try {
      if ((y(true), C.error)) throw Error(C.error);
      const W = applyPrintLayoutChange(e.workbook, e.sheetId, C.layout, {
        readOnly: e.readOnly,
        expectedSnapshot: g.snapshot,
      });
      (W !== e.workbook && e.onWorkbookChange(W, "Update page layout"),
        e.onClose());
    } catch (error) {
      k(error instanceof Error ? error.message : "Could not save page setup.");
    }
  }
  function U() {
    try {
      y(true);
      const rangeAddress = parseRangeAddress(e.range);
      if (
        (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
          (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
        500
      )
        throw new Error("Select at most 500 cells to apply a theme.");
      const ae =
          F === "mint"
            ? ["#548235", "#E2EFDA"]
            : F === "slate"
              ? ["#44546A", "#D6DCE4"]
              : F === "amber"
                ? ["#9C6500", "#FFF2CC"]
                : ["#4472C4", "#D9E2F3"],
        cells = { ...sheetOrThrow.cells };
      for (const de of assertRangeCellBudget(e.range)) {
        const bold = parseRangeAddress(de).firstRow === rangeAddress.firstRow;
        cells[de] = {
          ...cells[de],
          raw: cells[de]?.raw ?? null,
          style: {
            ...cells[de]?.style,
            fontFamily: "Calibri",
            fill: bold ? ae[0] : ae[1],
            color: bold ? "#FFFFFF" : "#242424",
            bold: bold,
          },
        };
      }
      (e.onWorkbookChange(
        {
          ...e.workbook,
          sheets: e.workbook.sheets.map((sheet) =>
            sheet.id === sheetOrThrow.id ? { ...sheet, cells: cells } : sheet,
          ),
        },
        `Apply ${F} theme to selection`,
      ),
        e.onClose());
    } catch (error) {
      k(error instanceof Error ? error.message : "Theme could not be applied.");
    }
  }
  async function G() {
    if (a.current) return;
    let W = null;
    const K = new AbortController();
    ((a.current = K), h(true), k(""));
    try {
      if ((y(), C.error)) throw Error(C.error);
      const html = renderPrintDocument(e.workbook, e.sheetId, C.layout).html;
      if (((W = window.open("", "_blank")), !W))
        throw Error(
          "Allow the print preview window in your browser, then try again.",
        );
      const ae = o.current;
      await waitForPrintAssets(ae, { signal: K.signal });
      const oe = measurePrintLayoutIssues(ae);
      (d(oe), summarizePrintIssues(oe));
      const de = await printDocumentInWindow(W, html, {
        signal: K.signal,
        assertCurrent: () => {
          if ((y(), K.signal.aborted))
            throw Error("Print preparation was cancelled.");
          if (N.current !== T)
            throw Error("Page settings changed. Run the print check again.");
        },
      });
      d(de);
    } catch (error) {
      (W && !W.closed && W.close(),
        K.signal.aborted ||
          k(
            error instanceof Error
              ? error.message
              : "Print preview could not open.",
          ));
    } finally {
      ((a.current = null), h(false));
    }
  }
  async function H() {
    if (a.current) return;
    const W = new AbortController();
    ((a.current = W), h(true), p("Preparing PDF…"), k(""), v(null));
    try {
      if ((y(), C.error)) throw Error(C.error);
      await waitForPrintAssets(o.current, { signal: W.signal });
      const K = measurePrintLayoutIssues(o.current);
      (d(K), summarizePrintIssues(K));
      const te = await renderSchedulePdf(e.workbook, e.sheetId, {
        layout: C.layout,
        signal: W.signal,
        assertCurrent: () => {
          if ((y(), N.current !== T))
            throw Error(
              "Page settings changed. Create the PDF again from the updated preview.",
            );
        },
        onProgress: (ae, oe) => p(`Creating PDF: ${ae} of ${oe} pages…`),
      });
      if (W.signal.aborted) return;
      (d(te.inspection),
        v({
          url: URL.createObjectURL(
            new Blob([new Uint8Array(te.bytes)], { type: "application/pdf" }),
          ),
          filename: te.filename,
          pages: te.pageCount,
        }));
    } catch (error) {
      W.signal.aborted ||
        k(
          error instanceof Error
            ? error.message
            : "PDF creation failed. Try browser printing.",
        );
    } finally {
      ((a.current = null), h(false), p(""));
    }
  }
  function Q() {
    (a.current?.abort(), e.onClose());
  }
  return i.jsx("div", {
    className: "dg-backdrop",
    onKeyDown: (event) => {
      event.key === "Escape" && Q();
    },
    children: i.jsxs("section", {
      className: `dg-dialog ${B ? "" : "dg-page-dialog"}`,
      role: "dialog",
      "aria-modal": "true",
      "aria-label": B ? "Workbook themes" : "Page setup and print preview",
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("span", {
              children: [
                B
                  ? i.jsx(Palette, { size: 19 })
                  : i.jsx(FileText, { size: 19 }),
                B ? "Themes" : "Page Setup",
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close page setup"),
              onClick: Q,
              lang: language("Close page setup"),
              "data-ui-label": "Close page setup",
              children: i.jsx(X, { size: 17 }),
            }),
          ],
        }),
        B
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
                      value: F,
                      onChange: (event) => I(event.target.value),
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
                    e.range,
                  ],
                }),
                A &&
                  i.jsx("p", {
                    role: "alert",
                    className: "dg-error",
                    children: A,
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
                                L({ orientation: event.target.value }),
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
                                L({ paper: event.target.value }),
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
                            L({ printArea: event.target.value.toUpperCase() }),
                          lang: language("Print area"),
                          "data-ui-label": "Print area",
                        }),
                      ],
                    }),
                    i.jsx("button", {
                      className: "dg-inline",
                      onClick: () => L({ printArea: e.range }),
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
                                L({ repeatRows: event.target.value }),
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
                                L({
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
                                L({ pageOrder: event.target.value }),
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
                              value: O,
                              onChange: (event) => j(event.target.value),
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
                              value: z,
                              onChange: (event) =>
                                V(event.target.value.toUpperCase()),
                              lang: language("Break before columns"),
                              "data-ui-label": "Break before columns",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => {
                            (j(""), V(""));
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
                          ].map((W) =>
                            i.jsxs(
                              "label",
                              {
                                children: [
                                  W,
                                  i.jsx(LocalizedNumberField, {
                                    min: 0,
                                    max: 3,
                                    step: 0.05,
                                    "aria-label": `${W} margin`,
                                    value: layout.margins[W],
                                    onValueChange: (K) =>
                                      L({
                                        margins: {
                                          ...layout.margins,
                                          [W]: Number(K),
                                        },
                                      }),
                                  }),
                                ],
                              },
                              W,
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
                            L({ fitToPage: event.target.checked }),
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
                                  onValueChange: (W) =>
                                    L({ fitToWidth: Number(W) }),
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
                                  onValueChange: (W) =>
                                    L({ fitToHeight: Number(W) }),
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
                              onValueChange: (W) => L({ scale: Number(W) }),
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
                            L({ gridlines: event.target.checked }),
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
                            L({ headings: event.target.checked }),
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
                            L({ header: event.target.value }),
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
                            L({ footer: event.target.value }),
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
                      onClick: () => L(defaultPrintHeaderFooter),
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
                    A &&
                      i.jsx("p", {
                        role: "alert",
                        className: "dg-error",
                        children: A,
                      }),
                  ],
                }),
                i.jsxs("div", {
                  className: "dg-print-preview",
                  ref: o,
                  children: [
                    i.jsx("style", { children: printPaperStyles }),
                    i.jsxs("div", {
                      className: "dg-preview-label",
                      children: [
                        i.jsx(LocalizedText, { children: "PRINT PREVIEW · " }),
                        C.pages.length,
                        " ",
                        C.pages.length === 1 ? "PAGE" : "PAGES",
                        " ·",
                        " ",
                        Math.round(C.scale * 100),
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
                          disabled: u || c || !!C.error,
                          children: u ? "Checking…" : "Check print output",
                        }),
                        s &&
                          i.jsxs("div", {
                            role: "status",
                            children: [
                              i.jsx("p", {
                                children: s.issues.length
                                  ? s.issues.filter(
                                      (issue) => issue.severity === "error",
                                    ).length +
                                    " blocking issue(s), " +
                                    s.issues.filter(
                                      (issue) => issue.severity === "warning",
                                    ).length +
                                    " warning(s)."
                                  : "Print check passed: " +
                                    s.pages +
                                    " page(s); no measured clipping or missing images.",
                              }),
                              s.issues.length > 0 &&
                                i.jsx("ul", {
                                  children: s.issues.slice(0, 100).map((W, K) =>
                                    i.jsxs(
                                      "li",
                                      {
                                        children: [
                                          W.severity === "error"
                                            ? "Fix: "
                                            : "Review: ",
                                          W.message,
                                        ],
                                      },
                                      K,
                                    ),
                                  ),
                                }),
                              s.issues.length > 100 &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Showing the first 100 of ",
                                    }),
                                    s.issues.length,
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children:
                                        "issues. Correct the affected sizes and check again.",
                                    }),
                                  ],
                                }),
                            ],
                          }),
                        f && i.jsx("p", { role: "status", children: f }),
                        m &&
                          i.jsxs("p", {
                            role: "status",
                            children: [
                              i.jsxs("a", {
                                href: m.url,
                                download: m.filename,
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Download PDF (",
                                  }),
                                  m.pages,
                                  " ",
                                  m.pages === 1 ? "page" : "pages",
                                  ")",
                                ],
                              }),
                              " ",
                              "· ",
                              m.filename,
                            ],
                          }),
                      ],
                    }),
                    C.notices?.map((notice) =>
                      i.jsx(
                        "p",
                        { className: "dg-muted", children: notice },
                        notice,
                      ),
                    ),
                    C.ruleStyles.errors().length > 0 &&
                      i.jsxs("p", {
                        role: "alert",
                        className: "dg-error",
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Conditional formatting: ",
                          }),
                          C.ruleStyles.errors().length,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "evaluation error(s). ",
                          }),
                          C.ruleStyles.errors()[0].message,
                        ],
                      }),
                    C.error
                      ? i.jsx("p", {
                          className: "dg-error",
                          role: "alert",
                          children: C.error,
                        })
                      : C.pages.map((page, K) =>
                          i.jsx(
                            PrintPagePreview,
                            {
                              workbook: e.workbook,
                              sheet: sheetOrThrow,
                              layout: C.layout,
                              page: page,
                              index: K,
                              total: C.pages.length,
                              scale: C.scale,
                              calc: calc,
                              ruleStyles: C.ruleStyles,
                            },
                            K,
                          ),
                        ),
                  ],
                }),
              ],
            }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: Q,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            !B &&
              i.jsx("button", {
                onClick: H,
                disabled: !!C.error || c || u,
                children: i.jsx(LocalizedText, { children: "Create PDF" }),
              }),
            !B &&
              i.jsxs("button", {
                onClick: G,
                disabled: !!C.error || c || u,
                children: [
                  i.jsx(Printer, { size: 14 }),
                  c && !f ? "Preparing print…" : "Print / Save as PDF",
                ],
              }),
            i.jsx("button", {
              className: "dg-primary",
              onClick: B ? U : $,
              disabled:
                e.readOnly ||
                c ||
                u ||
                !!sheetOrThrow.protection?.enabled ||
                (!B && !!C.error),
              children: B ? "Apply theme" : "Save page setup",
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
  const o = value.slice(0, caret);
  let a = false;
  const s = [];
  for (let I = 1; I < o.length; I++) {
    if (o[I] === '"') {
      if (a && o[I + 1] === '"') {
        I++;
        continue;
      }
      a = !a;
      continue;
    }
    if (!a)
      if (o[I] === "(") {
        const O =
          /([A-Za-z][A-Za-z0-9_.]*)\s*$/
            .exec(o.slice(0, I))?.[1]
            ?.toUpperCase() ?? "";
        s.push({ name: O, argumentIndex: 0 });
      } else
        o[I] === ")"
          ? s.pop()
          : o[I] === "," && s.length && s[s.length - 1].argumentIndex++;
  }
  if (a) return { items: [] };
  const installedExtensionFunctions = listInstalledExtensionFunctions(workbook),
    u = s.at(-1),
    l = installedExtensionFunctions.find(
      (installedExtensionFunction) =>
        installedExtensionFunction.name === u?.name,
    ),
    c = u?.name
      ? [
          ...(workbook.definedNames ?? []).filter(
            (I) => I.scopeSheetId === sheetId,
          ),
          ...(workbook.definedNames ?? []).filter((I) => !I.scopeSheetId),
        ].find((I) => I.name.toUpperCase() === u.name)
      : undefined,
    h = c ? parseFormulaTree(c.refersTo) : undefined,
    hint =
      u?.name && (builtinFunctionNames.includes(u.name) || h || l)
        ? {
            ...u,
            signature:
              (l
                ? l.fn.parameters.map((parameter) => parameter.name).join(", ")
                : h
                  ? h.join(", ")
                  : functionArgumentSignatures[u.name]) ??
              "value1, [value2], …",
          }
        : undefined,
    p = /([A-Za-z_][A-Za-z0-9_.]*(?:\[[^\]]*)?)$/.exec(o),
    m = /^=\s*$/.test(o);
  if (!p && !m) return { items: [], hint: hint };
  const v = p?.[1] ?? "",
    start = p ? caret - v.length : caret;
  if (start > 0 && /[A-Za-z0-9_$!']/.test(value[start - 1]))
    return { items: [], hint: hint };
  const b = v.toLowerCase(),
    y = [],
    w = (I, O, insertText, detail) => {
      I.toLowerCase().startsWith(b) &&
        y.push({
          label: I,
          kind: O,
          insertText: insertText,
          start: start,
          end: caret,
          detail: detail,
        });
    },
    E = [
      "SUM",
      "AVERAGE",
      "IF",
      "XLOOKUP",
      "COUNT",
      "SUMIFS",
      "FILTER",
      "UNIQUE",
    ];
  for (const I of v ? builtinFunctionNames : E)
    w(
      I,
      "function",
      `${I}(`,
      `${I}(${functionArgumentSignatures[I] ?? "value1, …"})${functionDescriptions[I] ? ` · ${functionDescriptions[I]}` : ""}`,
    );
  const A = new Set();
  for (const installedExtensionFunction of installedExtensionFunctions)
    w(
      installedExtensionFunction.name,
      "function",
      `${installedExtensionFunction.name}(`,
      `${installedExtensionFunction.name}(${installedExtensionFunction.fn.parameters.map((parameter) => parameter.name).join(", ")}) · ${installedExtensionFunction.entry.enabled ? "Custom function · manual refresh" : "Disabled extension"} · ${installedExtensionFunction.fn.description}`,
    );
  for (const I of [
    ...(workbook.definedNames ?? []).filter((O) => O.scopeSheetId === sheetId),
    ...(workbook.definedNames ?? []).filter((O) => !O.scopeSheetId),
  ]) {
    const O = I.name.toLowerCase(),
      formulaTree = parseFormulaTree(I.refersTo);
    (A.has(O) ||
      w(
        I.name,
        "name",
        `${I.name}${formulaTree ? "(" : ""}`,
        formulaTree
          ? `${I.name}(${formulaTree.join(", ")})${I.comment ? ` · ${I.comment}` : ""}`
          : I.refersTo,
      ),
      A.add(O));
  }
  for (const sheet of workbook.sheets)
    for (const O of sheet.tables ?? []) {
      w(O.name, "table", `${O.name}[`, `${sheet.name}!${O.range}`);
      for (const column of O.columns)
        w(
          `${O.name}[${column.name}]`,
          "table",
          `${O.name}[${column.name}]`,
          `Table column · ${column.name}`,
        );
    }
  return (
    v &&
      y.sort((I, O) => {
        const j = (z) =>
          z.label.toLowerCase() === b
            ? -2
            : z.kind === "function" && E.includes(z.label)
              ? E.indexOf(z.label)
              : 100;
        return j(I) - j(O) || I.label.localeCompare(O.label);
      }),
    { items: y.slice(0, 12), hint: hint }
  );
}
function applyFormulaCompletion(value, r) {
  const t = value.slice(r.end),
    n =
      r.insertText.endsWith("(") && t.startsWith("(")
        ? r.insertText.slice(0, -1)
        : r.insertText;
  return { value: value.slice(0, r.start) + n + t, caret: r.start + n.length };
}

export {
  functionDescriptions,
  PageSetupDialog,
  getFormulaCompletions,
  applyFormulaCompletion,
  functionArgumentSignatures,
};
