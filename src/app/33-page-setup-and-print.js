import { X, Palette, FileText, Printer } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { $p, z3 } from "./01-core-schemas-and-hashing.js";
import {
  $U,
  Fc,
  Fw,
  GE,
  Ix,
  Vn,
  XE,
  ZR,
  ct,
  jn,
  nae,
  oae,
  ut,
} from "./03-workbook-model.js";
import { Dr, M, ci } from "./04-cell-input-paste-and-fill.js";
import {
  Fme,
  Mme,
  Xz,
  cW,
  gw,
  k3,
  mw,
  qz,
  uW,
} from "./32-drawing-and-chart-export.js";

function Bme(e) {
  var X$1, q;
  const { text, language } = Dr(),
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
    ne.useEffect(
      () => () => {
        return a.current?.abort();
      },
      [],
    ));
  const [g] = ne.useState(() => ({
      snapshot: nae(e.workbook),
      sheetId: e.sheetId,
      range: e.range,
      sheet: structuredClone(e.workbook.sheets.find((W) => W.id === e.sheetId)),
    })),
    b = e.workbook.sheets.find((W) => W.id === e.sheetId) ?? g.sheet;
  function y(W = false) {
    const K = n.current;
    if (
      ($U(K.workbook, g.snapshot),
      K.sheetId !== g.sheetId || K.range !== g.range)
    )
      throw Error(
        "The worksheet or selection changed. Close and reopen page setup.",
      );
    if (W) {
      if (K.readOnly) throw Error("This workbook is read-only.");
      jn(K.workbook.sheets.find((te) => te.id === g.sheetId));
    }
  }
  const [w, E] = ne.useState(() =>
      structuredClone(b.pageLayout ?? { ...ZR(b), ...XE }),
    ),
    [A, k] = ne.useState(""),
    [F, I] = ne.useState("office"),
    [O, j] = ne.useState(() => {
      return (b.pageLayout?.rowBreaks ?? [])
        .map((K) => String(K + 1))
        .join(", ");
    }),
    [z, V] = ne.useState(() => {
      return (b.pageLayout?.columnBreaks ?? []).map(ut).join(", ");
    });
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
          return te === "row" ? Number(de) - 1 : ct(`${de}1`).firstCol;
        })
        .sort((oe, de) => oe - de);
      return GE(ae, te === "row" ? b.rowCount : b.colCount);
    };
    return { ...w, rowBreaks: W(O, "row"), columnBreaks: W(z, "column") };
  }
  const B = /theme/i.test(e.mode),
    _ = ne.useMemo(() => Vn(e.workbook), [e.workbook]),
    C = ne.useMemo(() => {
      try {
        if (!e.workbook.sheets.some((W) => W.id === e.sheetId))
          throw Error(
            "The worksheet changed or was deleted. Close and reopen page setup.",
          );
        return { ...uW(e.workbook, e.sheetId, D()), error: "" };
      } catch (W) {
        return {
          pages: [],
          notices: [],
          scale: 1,
          layout: w,
          ruleStyles: Fw({ ...b, conditionalFormats: [] }, _),
          error: W instanceof Error ? W.message : "Invalid page setup.",
        };
      }
    }, [e.workbook, e.sheetId, b, w, _, O, z]),
    T = JSON.stringify([w, O, z]),
    N = ne.useRef(T);
  ((N.current = T),
    ne.useEffect(() => {
      (d(null), k(""), v(null));
    }, [T, e.workbook, e.sheetId, e.range]));
  async function R() {
    const W = o.current;
    if (!(!W || C.error)) {
      (l(true), k(""));
      try {
        if ((y(), await gw(W), y(), N.current !== T))
          throw Error("Page settings changed. Run the print check again.");
        d(mw(W));
      } catch (K) {
        k(K instanceof Error ? K.message : "Print check failed.");
      } finally {
        l(false);
      }
    }
  }
  const L = (W) => E((K) => ({ ...K, ...W }));
  function $() {
    try {
      if ((y(true), C.error)) throw Error(C.error);
      const W = oae(e.workbook, e.sheetId, C.layout, {
        readOnly: e.readOnly,
        expectedSnapshot: g.snapshot,
      });
      (W !== e.workbook && e.onWorkbookChange(W, "Update page layout"),
        e.onClose());
    } catch (W) {
      k(W instanceof Error ? W.message : "Could not save page setup.");
    }
  }
  function U() {
    try {
      y(true);
      const te = ct(e.range);
      if ((te.lastRow - te.firstRow + 1) * (te.lastCol - te.firstCol + 1) > 500)
        throw new Error("Select at most 500 cells to apply a theme.");
      const ae =
          F === "mint"
            ? ["#548235", "#E2EFDA"]
            : F === "slate"
              ? ["#44546A", "#D6DCE4"]
              : F === "amber"
                ? ["#9C6500", "#FFF2CC"]
                : ["#4472C4", "#D9E2F3"],
        oe = { ...b.cells };
      for (const de of Ix(e.range)) {
        const he = ct(de).firstRow === te.firstRow;
        oe[de] = {
          ...oe[de],
          raw: oe[de]?.raw ?? null,
          style: {
            ...oe[de]?.style,
            fontFamily: "Calibri",
            fill: he ? ae[0] : ae[1],
            color: he ? "#FFFFFF" : "#242424",
            bold: he,
          },
        };
      }
      (e.onWorkbookChange(
        {
          ...e.workbook,
          sheets: e.workbook.sheets.map((de) =>
            de.id === b.id ? { ...de, cells: oe } : de,
          ),
        },
        `Apply ${F} theme to selection`,
      ),
        e.onClose());
    } catch (te) {
      k(te instanceof Error ? te.message : "Theme could not be applied.");
    }
  }
  async function G() {
    if (a.current) return;
    let W = null;
    const K = new AbortController();
    ((a.current = K), h(true), k(""));
    try {
      if ((y(), C.error)) throw Error(C.error);
      const te = cW(e.workbook, e.sheetId, C.layout).html;
      if (((W = window.open("", "_blank")), !W))
        throw Error(
          "Allow the print preview window in your browser, then try again.",
        );
      const ae = o.current;
      await gw(ae, { signal: K.signal });
      const oe = mw(ae);
      (d(oe), k3(oe));
      const de = await Mme(W, te, {
        signal: K.signal,
        assertCurrent: () => {
          if ((y(), K.signal.aborted))
            throw Error("Print preparation was cancelled.");
          if (N.current !== T)
            throw Error("Page settings changed. Run the print check again.");
        },
      });
      d(de);
    } catch (te) {
      (W && !W.closed && W.close(),
        K.signal.aborted ||
          k(
            te instanceof Error ? te.message : "Print preview could not open.",
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
      await gw(o.current, { signal: W.signal });
      const K = mw(o.current);
      (d(K), k3(K));
      const te = await Fme(e.workbook, e.sheetId, {
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
    } catch (K) {
      W.signal.aborted ||
        k(
          K instanceof Error
            ? K.message
            : "PDF creation failed. Try browser printing.",
        );
    } finally {
      ((a.current = null), h(false), p(""));
    }
  }
  function Q() {
    var W;
    ((W = a.current) == null || W.abort(), e.onClose());
  }
  return i.jsx("div", {
    className: "dg-backdrop",
    onKeyDown: (W) => {
      W.key === "Escape" && Q();
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
                  children: i.jsx(M, {
                    children:
                      "Apply a coordinated font and header/body colors to the selected cells. Values, formulas, notes, and review marks are preserved.",
                  }),
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Theme" }),
                    i.jsxs("select", {
                      "aria-label": text("Theme"),
                      value: F,
                      onChange: (W) => I(W.target.value),
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
                  children: [i.jsx(M, { children: "Selection: " }), e.range],
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
                            i.jsx(M, { children: "Orientation" }),
                            i.jsxs("select", {
                              "aria-label": text("Page orientation"),
                              value: w.orientation,
                              onChange: (W) =>
                                L({ orientation: W.target.value }),
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
                            i.jsx(M, { children: "Paper" }),
                            i.jsxs("select", {
                              "aria-label": text("Paper size"),
                              value: w.paper,
                              onChange: (W) => L({ paper: W.target.value }),
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
                        i.jsx(M, { children: "Print area" }),
                        i.jsx("input", {
                          "aria-label": text("Print area"),
                          value: w.printArea,
                          onChange: (W) =>
                            L({ printArea: W.target.value.toUpperCase() }),
                          lang: language("Print area"),
                          "data-ui-label": "Print area",
                        }),
                      ],
                    }),
                    i.jsx("button", {
                      className: "dg-inline",
                      onClick: () => L({ printArea: e.range }),
                      children: i.jsx(M, { children: "Use current selection" }),
                    }),
                    i.jsxs("fieldset", {
                      children: [
                        i.jsx("legend", {
                          children: i.jsx(M, {
                            children: "Repeat print titles",
                          }),
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Rows to repeat at top" }),
                            i.jsx("input", {
                              "aria-label": text("Rows to repeat at top"),
                              placeholder: "1:2",
                              value: w.repeatRows ?? "",
                              onChange: (W) =>
                                L({ repeatRows: W.target.value }),
                              lang: language("Rows to repeat at top"),
                              "data-ui-label": "Rows to repeat at top",
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Columns to repeat at left" }),
                            i.jsx("input", {
                              "aria-label": text("Columns to repeat at left"),
                              placeholder: text("A:B"),
                              value: w.repeatColumns ?? "",
                              onChange: (W) =>
                                L({
                                  repeatColumns: W.target.value.toUpperCase(),
                                }),
                              lang: language("Columns to repeat at left"),
                              "data-ui-label": "Columns to repeat at left",
                            }),
                          ],
                        }),
                        i.jsx("p", {
                          className: "dg-muted",
                          children: i.jsx(M, {
                            children:
                              "Titles may be outside the print area. Hidden title rows/columns are omitted. Up to 100 rows and 50 columns.",
                          }),
                        }),
                      ],
                    }),
                    i.jsxs("fieldset", {
                      children: [
                        i.jsx("legend", {
                          children: i.jsx(M, {
                            children: "Manual page breaks",
                          }),
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Page order" }),
                            i.jsxs("select", {
                              "aria-label": text("Page order"),
                              value: w.pageOrder ?? "overThenDown",
                              onChange: (W) => L({ pageOrder: W.target.value }),
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
                          children: i.jsx(M, {
                            children:
                              "Controls the order pages are numbered and printed. Page contents and the total page count stay the same.",
                          }),
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Break before rows" }),
                            i.jsx("input", {
                              "aria-label": text("Break before rows"),
                              placeholder: "25, 50",
                              value: O,
                              onChange: (W) => j(W.target.value),
                              lang: language("Break before rows"),
                              "data-ui-label": "Break before rows",
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Break before columns" }),
                            i.jsx("input", {
                              "aria-label": text("Break before columns"),
                              placeholder: text("D, H"),
                              value: z,
                              onChange: (W) => V(W.target.value.toUpperCase()),
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
                          children: i.jsx(M, {
                            children: "Reset manual breaks",
                          }),
                        }),
                        i.jsx("p", {
                          className: "dg-muted",
                          children: i.jsx(M, {
                            children:
                              "Enter row numbers or column letters separated by commas. Edit a boundary to move it; remove an entry to delete it. Fit to pages ignores manual breaks.",
                          }),
                        }),
                      ],
                    }),
                    i.jsxs("fieldset", {
                      children: [
                        i.jsx("legend", {
                          children: i.jsx(M, { children: "Margins (inches)" }),
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
                                  i.jsx(ci, {
                                    min: 0,
                                    max: 3,
                                    step: 0.05,
                                    "aria-label": `${W} margin`,
                                    value: w.margins[W],
                                    onValueChange: (K) =>
                                      L({
                                        margins: {
                                          ...w.margins,
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
                          checked: w.fitToPage,
                          onChange: (W) => L({ fitToPage: W.target.checked }),
                        }),
                        i.jsx(M, { children: "Fit to pages" }),
                      ],
                    }),
                    w.fitToPage
                      ? i.jsxs("div", {
                          className: "dg-two",
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(M, { children: "Pages wide" }),
                                i.jsx(ci, {
                                  min: 0,
                                  max: 10,
                                  "aria-label": text("Pages wide"),
                                  value: w.fitToWidth,
                                  onValueChange: (W) =>
                                    L({ fitToWidth: Number(W) }),
                                  lang: language("Pages wide"),
                                  "data-ui-label": "Pages wide",
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(M, { children: "Pages tall" }),
                                i.jsx(ci, {
                                  min: 0,
                                  max: 10,
                                  "aria-label": text("Pages tall"),
                                  value: w.fitToHeight,
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
                            i.jsx(M, { children: "Scale (%)" }),
                            i.jsx(ci, {
                              min: 10,
                              max: 400,
                              "aria-label": text("Print scale"),
                              value: w.scale,
                              onValueChange: (W) => L({ scale: Number(W) }),
                              lang: language("Print scale"),
                              "data-ui-label": "Print scale",
                            }),
                          ],
                        }),
                    i.jsx("p", {
                      className: "dg-muted",
                      children: i.jsx(M, {
                        children:
                          "Zero pages means automatic. Fit to pages reduces the content as needed.",
                      }),
                    }),
                    i.jsxs("label", {
                      className: "dg-check",
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: w.gridlines,
                          onChange: (W) => L({ gridlines: W.target.checked }),
                        }),
                        i.jsx(M, { children: "Print gridlines" }),
                      ],
                    }),
                    i.jsxs("label", {
                      className: "dg-check",
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: w.headings,
                          onChange: (W) => L({ headings: W.target.checked }),
                        }),
                        i.jsx(M, { children: "Print row and column headings" }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Header" }),
                        i.jsx("textarea", {
                          "aria-label": text("Print header"),
                          rows: 3,
                          maxLength: 255,
                          value: w.header,
                          onChange: (W) => L({ header: W.target.value }),
                          lang: language("Print header"),
                          "data-ui-label": "Print header",
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Footer" }),
                        i.jsx("textarea", {
                          "aria-label": text("Print footer"),
                          rows: 2,
                          maxLength: 255,
                          value: w.footer,
                          onChange: (W) => L({ footer: W.target.value }),
                          lang: language("Print footer"),
                          "data-ui-label": "Print footer",
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      className: "dg-muted",
                      children: i.jsx(M, {
                        children:
                          "&P page · &N total pages · &A sheet name · && literal ampersand",
                      }),
                    }),
                    i.jsx("button", {
                      type: "button",
                      onClick: () => L(XE),
                      children: i.jsx(M, {
                        children: "Insert workbook identifiers",
                      }),
                    }),
                    i.jsx("p", {
                      className: "dg-muted",
                      children: i.jsx(M, {
                        children:
                          "Replaces the header and footer with workbook name, engagement, period, revision and page numbers. Fields: &[Workbook], &[Engagement], &[Period], &[Revision]. Missing properties display a “not set” label. Values come from this workbook; revision is local history, not an approval or certification.",
                      }),
                    }),
                    !!((X$1 = b.merges) != null && X$1.length) &&
                      i.jsx("p", {
                        className: "dg-muted",
                        children: i.jsx(M, {
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
                    i.jsx("style", { children: qz }),
                    i.jsxs("div", {
                      className: "dg-preview-label",
                      children: [
                        i.jsx(M, { children: "PRINT PREVIEW · " }),
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
                          children: i.jsx(M, {
                            children:
                              "Create PDF produces a local image-based snapshot at 192 dpi. Text in this download is not searchable, selectable or accessibility-tagged. All pages are checked before the download is offered.",
                          }),
                        }),
                        i.jsxs("details", {
                          children: [
                            i.jsx("summary", {
                              children: i.jsx(M, {
                                children: "Browser printing settings",
                              }),
                            }),
                            i.jsx("p", {
                              children: i.jsx(M, {
                                children:
                                  "Print / Save as PDF uses your browser’s PDF destination. Match this paper size and orientation, use 100% scale and no browser margins, and turn off browser headers/footers. Keep background graphics on for cell fills.",
                              }),
                            }),
                            i.jsx("p", {
                              children: i.jsx(M, {
                                children:
                                  "Printing rechecks the final document. A print dialog opening does not confirm that a PDF was saved.",
                              }),
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: R,
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
                                      (W) => W.severity === "error",
                                    ).length +
                                    " blocking issue(s), " +
                                    s.issues.filter(
                                      (W) => W.severity === "warning",
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
                                    i.jsx(M, {
                                      children: "Showing the first 100 of ",
                                    }),
                                    s.issues.length,
                                    " ",
                                    i.jsx(M, {
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
                                  i.jsx(M, { children: "Download PDF (" }),
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
                    C.notices?.map((W) =>
                      i.jsx("p", { className: "dg-muted", children: W }, W),
                    ),
                    C.ruleStyles.errors().length > 0 &&
                      i.jsxs("p", {
                        role: "alert",
                        className: "dg-error",
                        children: [
                          i.jsx(M, { children: "Conditional formatting: " }),
                          C.ruleStyles.errors().length,
                          " ",
                          i.jsx(M, { children: "evaluation error(s). " }),
                          C.ruleStyles.errors()[0].message,
                        ],
                      }),
                    C.error
                      ? i.jsx("p", {
                          className: "dg-error",
                          role: "alert",
                          children: C.error,
                        })
                      : C.pages.map((W, K) =>
                          i.jsx(
                            Xz,
                            {
                              workbook: e.workbook,
                              sheet: b,
                              layout: C.layout,
                              page: W,
                              index: K,
                              total: C.pages.length,
                              scale: C.scale,
                              calc: _,
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
              children: i.jsx(M, { children: "Cancel" }),
            }),
            !B &&
              i.jsx("button", {
                onClick: H,
                disabled: !!C.error || c || u,
                children: i.jsx(M, { children: "Create PDF" }),
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
                !!((q = b.protection) != null && q.enabled) ||
                (!B && !!C.error),
              children: B ? "Apply theme" : "Save page setup",
            }),
          ],
        }),
      ],
    }),
  });
}
const UT = {
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
  $T = {
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
function Lme(e, r, t, n) {
  if (!e.startsWith("=")) return { items: [] };
  const o = e.slice(0, r);
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
  const d = z3(t),
    u = s.at(-1),
    l = d.find((I) => I.name === u?.name),
    c =
      u != null && u.name
        ? [
            ...(t.definedNames ?? []).filter((I) => I.scopeSheetId === n),
            ...(t.definedNames ?? []).filter((I) => !I.scopeSheetId),
          ].find((I) => I.name.toUpperCase() === u.name)
        : undefined,
    h = c ? $p(c.refersTo) : undefined,
    f =
      u != null && u.name && (Fc.includes(u.name) || h || l)
        ? {
            ...u,
            signature:
              (l
                ? l.fn.parameters.map((I) => I.name).join(", ")
                : h
                  ? h.join(", ")
                  : UT[u.name]) ?? "value1, [value2], …",
          }
        : undefined,
    p = /([A-Za-z_][A-Za-z0-9_.]*(?:\[[^\]]*)?)$/.exec(o),
    m = /^=\s*$/.test(o);
  if (!p && !m) return { items: [], hint: f };
  const v = p?.[1] ?? "",
    g = p ? r - v.length : r;
  if (g > 0 && /[A-Za-z0-9_$!']/.test(e[g - 1])) return { items: [], hint: f };
  const b = v.toLowerCase(),
    y = [],
    w = (I, O, j, z) => {
      I.toLowerCase().startsWith(b) &&
        y.push({
          label: I,
          kind: O,
          insertText: j,
          start: g,
          end: r,
          detail: z,
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
  for (const I of v ? Fc : E)
    w(
      I,
      "function",
      `${I}(`,
      `${I}(${UT[I] ?? "value1, …"})${$T[I] ? ` · ${$T[I]}` : ""}`,
    );
  const A = new Set();
  for (const I of d)
    w(
      I.name,
      "function",
      `${I.name}(`,
      `${I.name}(${I.fn.parameters.map((O) => O.name).join(", ")}) · ${I.entry.enabled ? "Custom function · manual refresh" : "Disabled extension"} · ${I.fn.description}`,
    );
  for (const I of [
    ...(t.definedNames ?? []).filter((O) => O.scopeSheetId === n),
    ...(t.definedNames ?? []).filter((O) => !O.scopeSheetId),
  ]) {
    const O = I.name.toLowerCase(),
      j = $p(I.refersTo);
    (A.has(O) ||
      w(
        I.name,
        "name",
        `${I.name}${j ? "(" : ""}`,
        j
          ? `${I.name}(${j.join(", ")})${I.comment ? ` · ${I.comment}` : ""}`
          : I.refersTo,
      ),
      A.add(O));
  }
  for (const I of t.sheets)
    for (const O of I.tables ?? []) {
      w(O.name, "table", `${O.name}[`, `${I.name}!${O.range}`);
      for (const j of O.columns)
        w(
          `${O.name}[${j.name}]`,
          "table",
          `${O.name}[${j.name}]`,
          `Table column · ${j.name}`,
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
    { items: y.slice(0, 12), hint: f }
  );
}
function Pme(e, r) {
  const t = e.slice(r.end),
    n =
      r.insertText.endsWith("(") && t.startsWith("(")
        ? r.insertText.slice(0, -1)
        : r.insertText;
  return { value: e.slice(0, r.start) + n + t, caret: r.start + n.length };
}

export { $T, Bme, Lme, Pme, UT };
