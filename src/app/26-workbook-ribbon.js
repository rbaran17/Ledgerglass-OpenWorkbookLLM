import {
  FileSpreadsheet,
  LayoutGrid,
  FolderOpen,
  Search,
  ChevronDown,
  BookOpen,
  Filter,
  ClipboardCheck,
  Sparkles,
  History,
  ListFilter,
  Copy,
  ChartColumn,
  Download,
  ShieldCheck,
  MessageSquare,
  Clipboard,
  Rows3,
  Grid2x2,
  FilePlus2,
  Sheet,
  FileJson,
  Type,
  FileCode,
  ChevronUp,
  Undo2,
  Redo2,
  Paintbrush,
  CheckCheck,
  Scissors,
  Bold,
  Italic,
  Underline,
  AlignLeft,
  AlignCenter,
  AlignRight,
  WrapText,
  DollarSign,
  Percent,
  Grid3x3,
  Table2,
  Maximize,
  Sigma,
  ArrowDownToLine,
  Eraser,
  Image as Image$1,
  Square,
  Calculator,
  ArrowLeftRight,
  ArrowDownAZ,
  ArrowUpAZ,
  Columns3,
  PanelTop,
} from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { $B, Q3, l0, ys } from "./03-workbook-model.js";
import { Dr, M } from "./04-cell-input-paste-and-fill.js";
import { G7, Mn, Uh, Vde, Xl, lr } from "./24-merge-and-audit-analysis.js";
import { Hk } from "./25-color-picker.js";

function Hde(e) {
  var U, G, H, Q, X;
  const { text: r, language: t } = Dr(),
    [n, o] = ne.useState("Home"),
    [a, s] = ne.useState(false),
    d = ne.useRef(null),
    [u, l] = ne.useState(false),
    [c, h] = ne.useState(false),
    [f, p] = ne.useState(false),
    m = ne.useRef(null),
    v = ne.useRef(null),
    g = ne.useRef(null),
    b = ne.useId(),
    y = !c || f,
    w = () => {
      var J, q;
      return (q =
        (J = m.current) == null
          ? void 0
          : J.querySelector('[role="tab"][aria-selected="true"]')) == null
        ? void 0
        : q.focus();
    },
    E = () => {
      var J;
      ((J = v.current) != null && J.contains(document.activeElement) && w(),
        h((q) => !q),
        p(false),
        l(false));
    },
    A = (J) => {
      (o(J), l(false), c && p(true));
    };
  (ne.useEffect(() => {
    if (!f) return;
    const J = (q) => {
      var W;
      ((W = m.current) != null && W.contains(q.target)) || p(false);
    };
    return (
      document.addEventListener("pointerdown", J),
      document.addEventListener("focusin", J),
      () => {
        (document.removeEventListener("pointerdown", J),
          document.removeEventListener("focusin", J));
      }
    );
  }, [f]),
    ne.useEffect(() => {
      const J = (q) => {
        if (
          q.defaultPrevented ||
          q.repeat ||
          q.altKey ||
          q.shiftKey ||
          q.key !== "F1" ||
          !(q.ctrlKey || q.metaKey)
        )
          return;
        const W = q.target;
        if (!(W instanceof Element) || W.closest('[role="dialog"]')) return;
        const K = m.current,
          te = (K == null ? void 0 : K.closest(".lg-app")) ?? K;
        !te ||
          (te.classList.contains("lg-app")
            ? W.closest(".lg-app") !== te
            : !te.contains(W)) ||
          (q.preventDefault(), q.stopPropagation(), E());
      };
      return (
        document.addEventListener("keydown", J),
        () => document.removeEventListener("keydown", J)
      );
    }, []));
  const k = e.readOnly || e.busy,
    F = ys((U = e.cell) == null ? void 0 : U.style),
    I = (J) => e.onFormat({ style: J }),
    O = (J) => e.onFormat({ format: J, numFmt: void 0, numberFormat: void 0 }),
    j =
      ((G = e.cell) == null ? void 0 : G.numberFormat) &&
      !["number", "currency", "accounting", "percent"].includes(
        e.cell.numberFormat.category,
      ),
    z = (J) => {
      var K, te, ae, oe, de, he, Se;
      if ((K = e.cell) != null && K.numberFormat && !j) {
        const le = l0({
          ...e.cell.numberFormat,
          decimals: Math.max(
            0,
            Math.min(12, (e.cell.numberFormat.decimals ?? 2) + J),
          ),
        });
        e.onFormat({ numberFormat: le, numFmt: Q3(le) });
        return;
      }
      const q =
          $B((te = e.cell) == null ? void 0 : te.numFmt) ??
          (((ae = e.cell) == null ? void 0 : ae.format) === "currency" ||
          ((oe = e.cell) == null ? void 0 : oe.format) === "number" ||
          ((de = e.cell) == null ? void 0 : de.format) === "percent"
            ? 2
            : 0),
        W = Math.max(0, Math.min(10, q + J));
      e.onFormat({
        numberFormat: void 0,
        numFmt: `${((he = e.cell) == null ? void 0 : he.format) === "currency" ? "$" : ""}#,##0${W ? `.${"0".repeat(W)}` : ""}${((Se = e.cell) == null ? void 0 : Se.format) === "percent" ? "%" : ""}`,
      });
    };
  ne.useEffect(() => {
    if (!u) return;
    const J = (q) => {
      var W;
      ((W = g.current) != null && W.contains(q.target)) || l(false);
    };
    return (
      document.addEventListener("pointerdown", J),
      () => document.removeEventListener("pointerdown", J)
    );
  }, [u]);
  const V = (J) => {
      (l(false), J());
    },
    D = i.jsxs("select", {
      "aria-label": r("Number format"),
      value:
        (H = e.cell) != null && H.numberFormat
          ? "advanced"
          : (((Q = e.cell) == null ? void 0 : Q.format) ?? "general"),
      disabled: k,
      onChange: (J) =>
        J.target.value === "advanced"
          ? e.onTool("numberFormat")
          : O(J.target.value),
      lang: t("Number format"),
      "data-ui-label": "Number format",
      children: [
        ((X = e.cell) == null ? void 0 : X.numberFormat) &&
          i.jsxs("option", {
            value: "advanced",
            children: [
              e.cell.numberFormat.category.charAt(0).toUpperCase() +
                e.cell.numberFormat.category.slice(1),
              " ",
              "(",
              e.cell.numberFormat.locale,
              ")",
            ],
          }),
        i.jsx("option", { value: "general", children: "General" }),
        i.jsx("option", { value: "number", children: "Number" }),
        i.jsx("option", { value: "currency", children: "Currency" }),
        i.jsx("option", { value: "percent", children: "Percentage" }),
        i.jsx("option", { value: "date", children: "Short Date" }),
      ],
    }),
    B = i.jsx("select", {
      "aria-label": r("Zoom"),
      value: e.zoom,
      onChange: (J) => e.onZoom(Number(J.target.value)),
      lang: t("Zoom"),
      "data-ui-label": "Zoom",
      children: Array.from(
        new Set([50, 75, 90, 100, 110, 125, 150, 175, 200, e.zoom]),
      )
        .sort((J, q) => J - q)
        .map((J) => i.jsxs("option", { value: J, children: [J, "%"] }, J)),
    }),
    _ = (J) =>
      e.busy
        ? "Wait for the current workbook operation to finish."
        : !e.allowProtectionChange && /Protect|New Sheet|New from/.test(J)
          ? "The host has opened this workbook read-only."
          : e.readOnly
            ? e.sheetProtected
              ? "Unprotect the worksheet before editing."
              : "The host has opened this workbook read-only."
            : J === "Undo"
              ? "There is no earlier edit to undo."
              : J === "Redo"
                ? "There is no undone edit to redo."
                : /Decimal/.test(J)
                  ? "Choose a number, currency, accounting or percent format to adjust decimal places."
                  : /New Sheet/.test(J) && e.workbookStructureProtected
                    ? "Unprotect workbook structure in Manage worksheets before adding a sheet."
                    : "This command is unavailable for the current workbook state.",
    C = { group: Mn, command: lr, icon: Xl, color: Hk },
    T = (J) =>
      i.jsxs(i.Fragment, {
        children: [
          J === "Home" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs(Mn, {
                  label: r("Clipboard"),
                  className: "xl-clipboard-group",
                  children: [
                    i.jsxs("div", {
                      className: "xl-paste-commands",
                      children: [
                        i.jsx(lr, {
                          icon: i.jsx(Clipboard, {}),
                          label: r("Paste"),
                          large: true,
                          onClick: e.onPaste,
                          disabled: k,
                          className: "xl-gold-icon xl-paste-primary",
                        }),
                        i.jsxs("button", {
                          type: "button",
                          className: "xl-paste-special",
                          "aria-label": r("Paste Special"),
                          title: r(
                            "Choose values, formulas, formatting and other paste options",
                          ),
                          disabled: k,
                          onClick: () => e.onTool("pasteSpecial"),
                          lang: t("Paste Special"),
                          "data-ui-label": "Paste Special",
                          children: [
                            i.jsx(M, { children: "Paste Special" }),
                            i.jsx(ChevronDown, { size: 10 }),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "xl-command-stack",
                      children: [
                        i.jsx(lr, {
                          icon: i.jsx(Scissors, {}),
                          label: r("Cut"),
                          onClick: e.onCut,
                          disabled: k,
                        }),
                        i.jsx(lr, {
                          icon: i.jsx(Copy, {}),
                          label: r("Copy"),
                          onClick: e.onCopy,
                        }),
                        i.jsx(lr, {
                          icon: i.jsx(Paintbrush, {}),
                          label: r("Format Painter"),
                          onClick: () => e.onTool("paint"),
                          disabled: k,
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Font"),
                  className: "xl-font-group",
                  onDialog: () => e.onTool("cellStyle"),
                  dialogLabel: r("Cell Styles"),
                  dialogDisabled: k,
                  children: [
                    i.jsxs("div", {
                      className: "xl-font-selects",
                      children: [
                        i.jsx("select", {
                          "aria-label": r("Font family"),
                          value: F.fontFamily ?? "Calibri",
                          disabled: k,
                          onChange: (q) => I({ fontFamily: q.target.value }),
                          lang: t("Font family"),
                          "data-ui-label": "Font family",
                          children: [
                            "Calibri",
                            "Aptos",
                            "Arial",
                            "Times New Roman",
                            "Courier New",
                          ].map((q) => i.jsx("option", { children: q }, q)),
                        }),
                        i.jsx("select", {
                          "aria-label": r("Font size"),
                          value: F.fontSize ?? 11,
                          disabled: k,
                          onChange: (q) =>
                            I({ fontSize: Number(q.target.value) }),
                          lang: t("Font size"),
                          "data-ui-label": "Font size",
                          children: Array.from(
                            new Set([
                              8,
                              9,
                              10,
                              11,
                              12,
                              14,
                              16,
                              18,
                              20,
                              22,
                              24,
                              28,
                              36,
                              48,
                              72,
                              F.fontSize ?? 11,
                            ]),
                          )
                            .sort((q, W) => q - W)
                            .map((q) => i.jsx("option", { children: q }, q)),
                        }),
                        i.jsx(Xl, {
                          icon: i.jsxs("span", {
                            className: "xl-grow-font",
                            children: [
                              i.jsx(M, { children: "A" }),
                              i.jsx("sup", { children: "⌃" }),
                            ],
                          }),
                          label: r("Increase Font Size"),
                          disabled: k,
                          onClick: () =>
                            I({
                              fontSize: Math.min(72, (F.fontSize ?? 11) + 1),
                            }),
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "xl-font-buttons",
                      children: [
                        i.jsx(Xl, {
                          icon: i.jsx(Bold, { size: 15 }),
                          label: r("Bold (Ctrl+B)"),
                          active: !!F.bold,
                          onClick: () => I({ bold: !F.bold }),
                          disabled: k,
                        }),
                        i.jsx(Xl, {
                          icon: i.jsx(Italic, { size: 15 }),
                          label: r("Italic (Ctrl+I)"),
                          active: !!F.italic,
                          onClick: () => I({ italic: !F.italic }),
                          disabled: k,
                        }),
                        i.jsx(Xl, {
                          icon: i.jsx(Underline, { size: 15 }),
                          label: r("Underline (Ctrl+U)"),
                          active: !!F.underline,
                          onClick: () => I({ underline: !F.underline }),
                          disabled: k,
                        }),
                        i.jsx("span", { className: "xl-control-divider" }),
                        i.jsxs("select", {
                          className: "xl-border-select",
                          "aria-label": r("Borders"),
                          title: r("Borders"),
                          value: F.border ?? "none",
                          disabled: k,
                          onChange: (q) =>
                            I({ border: q.target.value, borders: void 0 }),
                          lang: t("Borders"),
                          "data-ui-label": "Borders",
                          children: [
                            i.jsx("option", { value: "none", children: "▦" }),
                            i.jsx("option", {
                              value: "all",
                              children: "All Borders",
                            }),
                            i.jsx("option", {
                              value: "bottom",
                              children: "Bottom Border",
                            }),
                          ],
                        }),
                        i.jsx(Hk, {
                          kind: "fill",
                          value: F.fill ?? "#FFFF00",
                          onChange: (q) => I({ fill: q }),
                          disabled: k,
                        }),
                        i.jsx(Hk, {
                          kind: "color",
                          value: F.color ?? "#C00000",
                          onChange: (q) => I({ color: q }),
                          disabled: k,
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Alignment"),
                  className: "xl-alignment-group",
                  children: [
                    i.jsxs("div", {
                      className: "xl-alignment-top",
                      children: [
                        i.jsx(Xl, {
                          icon: i.jsx(AlignLeft, { size: 16 }),
                          label: r("Align Left"),
                          active: F.align === "left",
                          onClick: () => I({ align: "left" }),
                          disabled: k,
                        }),
                        i.jsx(Xl, {
                          icon: i.jsx(AlignCenter, { size: 16 }),
                          label: r("Center"),
                          active: F.align === "center",
                          onClick: () => I({ align: "center" }),
                          disabled: k,
                        }),
                        i.jsx(Xl, {
                          icon: i.jsx(AlignRight, { size: 16 }),
                          label: r("Align Right"),
                          active: F.align === "right",
                          onClick: () => I({ align: "right" }),
                          disabled: k,
                        }),
                        i.jsx(lr, {
                          icon: i.jsx(WrapText, {}),
                          label: r("Wrap Text"),
                          active: !!F.wrap,
                          onClick: () => I({ wrap: !F.wrap }),
                          disabled: k,
                        }),
                      ],
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Grid2x2, {}),
                      label: r("Merge & Center"),
                      onClick: () => e.onTool("merge"),
                      disabled: k,
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Number"),
                  className: "xl-number-group",
                  children: [
                    D,
                    i.jsx(lr, {
                      icon: i.jsx(DollarSign, { size: 15 }),
                      label: r("Format Cells"),
                      onClick: () => e.onTool("numberFormat"),
                      disabled: k,
                    }),
                    i.jsxs("div", {
                      className: "xl-number-buttons",
                      children: [
                        i.jsx(Xl, {
                          icon: i.jsx(DollarSign, { size: 16 }),
                          label: r("Currency"),
                          onClick: () => O("currency"),
                          disabled: k,
                        }),
                        i.jsx(Xl, {
                          icon: i.jsx(Percent, { size: 15 }),
                          label: r("Percent Style"),
                          onClick: () => O("percent"),
                          disabled: k,
                        }),
                        i.jsx(Xl, {
                          icon: i.jsx("span", {
                            className: "xl-comma-icon",
                            children: ",",
                          }),
                          label: r("Comma Style"),
                          onClick: () =>
                            e.onFormat({
                              format: "number",
                              numFmt: "#,##0.00",
                              numberFormat: void 0,
                            }),
                          disabled: k,
                        }),
                        i.jsx(Xl, {
                          icon: i.jsxs("span", {
                            className: "xl-decimal-icon",
                            children: [
                              "←",
                              i.jsx("small", { children: ".00" }),
                            ],
                          }),
                          label: r("Increase Decimal"),
                          onClick: () => z(1),
                          disabled: k || !!j,
                        }),
                        i.jsx(Xl, {
                          icon: i.jsxs("span", {
                            className: "xl-decimal-icon",
                            children: ["→", i.jsx("small", { children: ".0" })],
                          }),
                          label: r("Decrease Decimal"),
                          onClick: () => z(-1),
                          disabled: k || !!j,
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Styles"),
                  className: "xl-styles-group",
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Grid3x3, {}),
                      label: r("Conditional Formatting"),
                      large: true,
                      onClick: () => e.onTool("conditional"),
                      disabled: k,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Table2, {}),
                      label: r("Format as Table"),
                      large: true,
                      onClick: () => e.onTool("table"),
                      disabled: k,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(LayoutGrid, {}),
                      label: r("Heading Style"),
                      large: true,
                      onClick: () =>
                        I({ fill: "#E2EFDA", color: "#375623", bold: true }),
                      disabled: k,
                      title: r(
                        "Apply green heading cell style to the selection",
                      ),
                      className: "xl-green-icon",
                    }),
                  ],
                }),
                i.jsx(Mn, {
                  label: r("Cells"),
                  className: "xl-cells-group",
                  children: i.jsxs("div", {
                    className: "xl-command-stack",
                    children: [
                      i.jsx(lr, {
                        icon: i.jsx(Rows3, {}),
                        label: r("Insert Row"),
                        onClick: e.onAddRow,
                        disabled: k,
                        className: "xl-green-icon",
                      }),
                      i.jsx(lr, {
                        icon: i.jsx(Table2, {}),
                        label: r("Insert Column"),
                        onClick: e.onAddColumn,
                        disabled: k,
                        className: "xl-green-icon",
                      }),
                      i.jsx(lr, {
                        icon: i.jsx(Sheet, {}),
                        label: r("New Sheet"),
                        onClick: e.onAddSheet,
                        disabled:
                          e.busy ||
                          !e.allowProtectionChange ||
                          e.workbookStructureProtected,
                      }),
                    ],
                  }),
                }),
                i.jsx(Mn, {
                  label: r("Cell Size"),
                  children: i.jsxs("div", {
                    className: "xl-command-stack",
                    children: [
                      i.jsx(lr, {
                        icon: i.jsx(Maximize, {}),
                        label: r("AutoFit Column Width"),
                        onClick: e.onAutoFitColumns,
                        disabled: k,
                        title: r("Fit the selected columns to their contents"),
                      }),
                      i.jsx(lr, {
                        icon: i.jsx(Rows3, {}),
                        label: r("AutoFit Row Height"),
                        onClick: e.onAutoFitRows,
                        disabled: k,
                        title: r(
                          "Fit the selected rows to their contents and wrapping",
                        ),
                      }),
                      i.jsx(lr, {
                        icon: i.jsx(Rows3, {}),
                        label: r("Hide & Unhide"),
                        onClick: () => e.onTool("viewSettings"),
                        title: r("Hide or restore selected rows and columns"),
                      }),
                    ],
                  }),
                }),
                i.jsxs(Mn, {
                  label: r("Editing"),
                  className: "xl-editing-group",
                  children: [
                    i.jsxs("div", {
                      className: "xl-command-stack",
                      children: [
                        i.jsx(lr, {
                          icon: i.jsx(Sigma, {}),
                          label: r("AutoSum"),
                          onClick: e.onAutosum,
                          disabled: k,
                        }),
                        i.jsx(lr, {
                          icon: i.jsx(ArrowDownToLine, {}),
                          label: r("Fill Down"),
                          onClick: e.onFillDown,
                          disabled: k,
                          className: "xl-blue-icon",
                        }),
                        i.jsx(lr, {
                          icon: i.jsx(Eraser, {}),
                          label: r("Clear"),
                          onClick: e.onClear,
                          disabled: k,
                        }),
                      ],
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Rows3, {}),
                      label: r("Fill Series"),
                      large: true,
                      onClick: () => e.onTool("fillSeries"),
                      disabled: k,
                      title: r(
                        "Fill in any direction or create a number or date series",
                      ),
                      className: "xl-blue-icon",
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(ListFilter, {}),
                      label: r("Sort & Filter"),
                      large: true,
                      onClick: e.onToggleFilter,
                      className: "xl-blue-icon",
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Search, {}),
                      label: r("Find & Select"),
                      large: true,
                      onClick: e.onFind,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Search, {}),
                      label: r("Advanced Find"),
                      onClick: () => e.onTool("search"),
                      title: r(
                        "Find all matches, preview replacements, or inspect special cells (Ctrl/⌘+Shift+F)",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Search, {}),
                      label: r("Go To Special"),
                      onClick: () => e.onTool("goToSpecial"),
                      title: r(
                        "Select formulas, errors, constants, notes or blanks (F5)",
                      ),
                    }),
                  ],
                }),
              ],
            }),
          J === "Insert" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs(Mn, {
                  label: r("Tables"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Table2, {}),
                      label: r("PivotTable"),
                      large: true,
                      onClick: () => e.onTool("pivot"),
                      disabled: k,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Table2, {}),
                      label: r("Table"),
                      large: true,
                      onClick: () => e.onTool("table"),
                      disabled: k,
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Illustrations"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Image$1, {}),
                      label: r("Pictures"),
                      large: true,
                      onClick: () => e.onTool("picture"),
                      disabled: k,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Square, {}),
                      label: r("Shapes"),
                      large: true,
                      onClick: () => e.onTool("shape"),
                      disabled: k,
                    }),
                  ],
                }),
                i.jsx(Mn, {
                  label: r("Charts"),
                  children: i.jsx(lr, {
                    icon: i.jsx(ChartColumn, {}),
                    label: r("Recommended Charts"),
                    large: true,
                    onClick: () => e.onTool("chart"),
                    disabled: k,
                  }),
                }),
                i.jsx(Mn, {
                  label: r("Comments"),
                  children: i.jsx(lr, {
                    icon: i.jsx(MessageSquare, {}),
                    label: r("New Note"),
                    large: true,
                    onClick: e.onNote,
                    disabled: k,
                    className: "xl-gold-icon",
                  }),
                }),
                i.jsxs(Mn, {
                  label: r("Worksheet"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Sheet, {}),
                      label: r("New Sheet"),
                      large: true,
                      onClick: e.onAddSheet,
                      disabled:
                        e.busy ||
                        !e.allowProtectionChange ||
                        e.workbookStructureProtected,
                      className: "xl-green-icon",
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Rows3, {}),
                      label: r("Insert Row"),
                      large: true,
                      onClick: e.onAddRow,
                      disabled: k,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Table2, {}),
                      label: r("Insert Column"),
                      large: true,
                      onClick: e.onAddColumn,
                      disabled: k,
                    }),
                  ],
                }),
                i.jsx(Mn, {
                  label: r("Files"),
                  children: i.jsx(lr, {
                    icon: i.jsx(FolderOpen, {}),
                    label: r("Open Workbook"),
                    large: true,
                    onClick: e.onImport,
                    disabled: k,
                    className: "xl-gold-icon",
                  }),
                }),
              ],
            }),
          J === "Page Layout" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsx(Mn, {
                  label: r("Themes"),
                  children: i.jsx(lr, {
                    icon: i.jsx(LayoutGrid, {}),
                    label: r("Themes"),
                    large: true,
                    onClick: () => e.onTool("themes"),
                    disabled: k,
                  }),
                }),
                i.jsxs(Mn, {
                  label: r("Page Setup"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Sheet, {}),
                      label: r("Margins"),
                      large: true,
                      onClick: () => e.onTool("margins"),
                      disabled: false,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Rows3, {}),
                      label: r("Orientation"),
                      large: true,
                      onClick: () => e.onTool("orientation"),
                      disabled: false,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Square, {}),
                      label: r("Size"),
                      large: true,
                      onClick: () => e.onTool("size"),
                      disabled: false,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Grid2x2, {}),
                      label: r("Print Area"),
                      large: true,
                      onClick: () => e.onTool("printArea"),
                      disabled: false,
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Print titles and breaks"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Rows3, {}),
                      label: r("Print Titles"),
                      onClick: () => e.onTool("printTitles"),
                      disabled: false,
                      large: true,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Grid2x2, {}),
                      label: r("Page Breaks"),
                      onClick: () => e.onTool("pageBreaks"),
                      disabled: false,
                      large: true,
                    }),
                  ],
                }),
                i.jsx(Mn, {
                  label: r("Scale to Fit"),
                  children: i.jsx(lr, {
                    icon: i.jsx(Maximize, {}),
                    label: r("Fit to Page"),
                    large: true,
                    onClick: () => e.onTool("fit"),
                    disabled: false,
                  }),
                }),
                i.jsxs(Mn, {
                  label: r("Sheet Options"),
                  className: "xl-options-group",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: e.gridlines,
                          onChange: e.onGridlines,
                        }),
                        i.jsx(M, { children: "Gridlines" }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: e.headings,
                          onChange: e.onHeadings,
                        }),
                        i.jsx(M, { children: "Headings" }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          J === "Formulas" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs(Mn, {
                  label: r("Function Library"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Sigma, {}),
                      label: r("AutoSum"),
                      large: true,
                      onClick: e.onAutosum,
                      disabled: k,
                      className: "xl-blue-icon",
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(BookOpen, {}),
                      label: r("Function Reference"),
                      large: true,
                      onClick: () => e.onPanel("help"),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Calculator, {}),
                      label: r("Insert Function"),
                      large: true,
                      onClick: () => e.onTool("function"),
                      disabled: false,
                    }),
                  ],
                }),
                i.jsx(Mn, {
                  label: r("Extensions"),
                  children: i.jsx(lr, {
                    icon: i.jsx(Calculator, {}),
                    label: r("Add-ins & Functions"),
                    large: true,
                    onClick: () => e.onTool("extensions"),
                    disabled: e.busy,
                  }),
                }),
                i.jsx(Mn, {
                  label: r("Defined Names"),
                  children: i.jsx(lr, {
                    icon: i.jsx(Table2, {}),
                    label: r("Name Manager"),
                    large: true,
                    onClick: () => e.onTool("names"),
                    disabled: false,
                  }),
                }),
                i.jsxs(Mn, {
                  label: r("Formula Auditing"),
                  children: [
                    i.jsxs("div", {
                      className: "xl-command-stack",
                      children: [
                        i.jsx(lr, {
                          icon: i.jsx(ArrowLeftRight, {}),
                          label: r("Trace Precedents"),
                          onClick: () => e.onTool("tracePrecedents"),
                          title: r(
                            "Inspect the cells referenced by the active formula",
                          ),
                        }),
                        i.jsx(lr, {
                          icon: i.jsx(ArrowLeftRight, {}),
                          label: r("Trace Dependents"),
                          onClick: () => e.onTool("traceDependents"),
                          title: r(
                            "Inspect formulas that reference the active cell",
                          ),
                        }),
                        i.jsx(lr, {
                          icon: i.jsx(Sheet, {}),
                          label: r("Watch Window"),
                          onClick: () => e.onTool("formulaWatches"),
                          title: r(
                            "Track selected cells and their current calculated values",
                          ),
                        }),
                      ],
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Calculator, {}),
                      label: r("Evaluate Formula"),
                      large: true,
                      onClick: () => e.onTool("evaluateFormula"),
                      title: r(
                        "Replay actual calculation steps and inspect errors",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(ShieldCheck, {}),
                      label: r("Compatibility"),
                      large: true,
                      onClick: () => e.onTool("functionCompatibility"),
                      title: r(
                        "Review function contracts and unsupported formulas",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Sheet, {}),
                      label: r("Show Formulas"),
                      large: true,
                      active: e.showFormulas,
                      onClick: e.onShowFormulas,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(ShieldCheck, {}),
                      label: r("Audit Checks"),
                      large: true,
                      onClick: () => e.onPanel("checks"),
                      className: "xl-green-icon",
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Calculation"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Calculator, {}),
                      label: r("Calculation Options"),
                      large: true,
                      onClick: () => e.onTool("calculation"),
                      title: r(
                        "Automatic or manual mode, iteration and calculation evidence",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Calculator, {}),
                      label: r("Calculate Now"),
                      large: true,
                      onClick: () => e.onTool("calculateNow"),
                      title: r(
                        "Recalculate the workbook without changing source cells",
                      ),
                    }),
                  ],
                }),
                i.jsx(Mn, {
                  label: r("Formula Assistance"),
                  children: i.jsx(lr, {
                    icon: i.jsx(Sparkles, {}),
                    label: r("Ask AI Assistant"),
                    large: true,
                    onClick: () => e.onPanel("assistant"),
                    className: "xl-purple-icon",
                  }),
                }),
              ],
            }),
          J === "Data" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs(Mn, {
                  label: r("Get & Transform Data"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(FolderOpen, {}),
                      label: r("Get Data"),
                      large: true,
                      onClick: e.onImport,
                      disabled: k,
                      className: "xl-gold-icon",
                      title: r("Import Excel, CSV, TSV, or workbook JSON"),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Table2, {}),
                      label: r("From Table / Range"),
                      large: true,
                      onClick: () => e.onTool("query"),
                      disabled: k,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(ArrowLeftRight, {}),
                      label: r("Append / Join"),
                      large: true,
                      onClick: () => e.onTool("combinePopulations"),
                      disabled: e.busy,
                      title: r(
                        "Review combined populations, unmatched records and match counts",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Table2, {}),
                      label: r("Transformation Recipes"),
                      large: true,
                      onClick: () => e.onTool("transformationRecipes"),
                      disabled: e.busy,
                      title: r(
                        "Build versioned steps with source fingerprints, failure details and reviewed outputs",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(FolderOpen, {}),
                      label: r("Connections & Refresh"),
                      large: true,
                      onClick: () => e.onTool("externalSources"),
                      disabled: e.busy,
                      title: r(
                        "Review approved external sources, control changes and last successful data",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Table2, {}),
                      label: r("Data Model"),
                      large: true,
                      onClick: () => e.onTool("dataModel"),
                      disabled: e.busy,
                      title: r(
                        "Build typed tables, relationships and measures across shared dimensions",
                      ),
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Sort & Filter"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Filter, {}),
                      label: r("Column Filters"),
                      large: true,
                      onClick: () => e.onTool("populationFilter"),
                      disabled: e.busy,
                      title: r(
                        "Filter values, text, numbers, dates and advanced criteria while retaining source rows",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(ArrowDownAZ, {}),
                      label: r("Custom Sort"),
                      large: true,
                      onClick: () => e.onTool("sort"),
                      disabled: e.busy,
                      title: r(
                        "Review a stable multi-level sort of complete worksheet records",
                      ),
                    }),
                    i.jsxs("div", {
                      className: "xl-command-stack",
                      children: [
                        i.jsx(lr, {
                          icon: i.jsx(ArrowDownAZ, {}),
                          label: r("Sort A to Z"),
                          onClick: () => e.onSort("asc"),
                          className: "xl-blue-icon",
                          title: r(
                            "Sort worksheet view ascending by the selected column",
                          ),
                        }),
                        i.jsx(lr, {
                          icon: i.jsx(ArrowUpAZ, {}),
                          label: r("Sort Z to A"),
                          onClick: () => e.onSort("desc"),
                          className: "xl-blue-icon",
                          title: r(
                            "Sort worksheet view descending by the selected column",
                          ),
                        }),
                        i.jsx(lr, {
                          icon: i.jsx(Eraser, {}),
                          label: r("Clear Sort"),
                          onClick: () => e.onSort(null),
                        }),
                      ],
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Filter, {}),
                      label: r("Filter"),
                      large: true,
                      onClick: e.onToggleFilter,
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Data Tools"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(ClipboardCheck, {}),
                      label: r("Audit Checks"),
                      large: true,
                      onClick: () => e.onPanel("checks"),
                      className: "xl-green-icon",
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(CheckCheck, {}),
                      label: r("Data Validation"),
                      large: true,
                      onClick: () => e.onTool("validation"),
                      disabled: e.busy,
                      title: r(
                        "Review validation rules and inspect existing values; editing requires write access",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Copy, {}),
                      label: r("Remove Duplicates"),
                      large: true,
                      onClick: () => e.onTool("duplicates"),
                      disabled: e.busy,
                      title: r(
                        "Review duplicate groups and create a separate deduplicated population",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Columns3, {}),
                      label: r("Text to Columns"),
                      large: true,
                      onClick: () => e.onTool("textToColumns"),
                      disabled: e.busy,
                      title: r(
                        "Preview split fields and types before creating a separate output",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Columns3, {}),
                      label: r("Clean up text"),
                      large: true,
                      onClick: () => e.onTool("textCleanup"),
                      disabled: e.busy,
                      title: r(
                        "Review cleanup and extraction steps while retaining original evidence",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Columns3, {}),
                      label: r("Clean & Reshape"),
                      large: true,
                      onClick: () => e.onTool("fieldTransform"),
                      disabled: e.busy,
                      title: r(
                        "Review splitting, conversions, gap filling and inferred patterns with retained exceptions",
                      ),
                    }),
                  ],
                }),
                i.jsx(Mn, {
                  label: r("Workbook Links"),
                  children: i.jsx(lr, {
                    icon: i.jsx(FileSpreadsheet, {}),
                    label: r("Workbook Links"),
                    large: true,
                    onClick: () => e.onTool("workbookLinks"),
                    disabled: e.busy,
                    title: r(
                      "Inspect, resolve, refresh, replace and break approved workbook links",
                    ),
                  }),
                }),
                i.jsxs(Mn, {
                  label: r("What-If Analysis"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Calculator, {}),
                      label: r("Scenario Manager"),
                      large: true,
                      onClick: () => e.onTool("scenarios"),
                      disabled: e.busy,
                      title: r(
                        "Compare named assumptions and review a scenario before applying it",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Table2, {}),
                      label: r("Data Table"),
                      large: true,
                      onClick: () => e.onTool("dataTable"),
                      disabled: e.busy,
                      title: r(
                        "Build a one- or two-variable sensitivity table with retained baseline assumptions",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Calculator, {}),
                      label: r("Goal Seek"),
                      large: true,
                      onClick: () => e.onTool("goalSeek"),
                      disabled: e.busy,
                      title: r(
                        "Find a numeric input that reaches a target formula result",
                      ),
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Analysis"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Calculator, {}),
                      label: r("Solver"),
                      large: true,
                      onClick: () => e.onTool("solver"),
                      disabled: e.busy,
                      title: r(
                        "Review bounded linear optimization with explicit constraints",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Table2, {}),
                      label: r("Data Analysis"),
                      large: true,
                      onClick: () => e.onTool("statistics"),
                      disabled: e.busy,
                      title: r(
                        "Descriptive statistics, regression, distributions and forecasts",
                      ),
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Outline"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Table2, {}),
                      label: r("Consolidate / Subtotals"),
                      large: true,
                      onClick: () => e.onTool("consolidation"),
                      disabled: e.busy,
                      title: r(
                        "Review cross-sheet or approved-workbook rollups, source context and grouped subtotals",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Rows3, {}),
                      label: r("Group & Outline"),
                      large: true,
                      onClick: () => e.onTool("outlines"),
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Export"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Copy, {}),
                      label: r("Report Outputs"),
                      large: true,
                      onClick: () => e.onTool("reportOutputs"),
                      disabled: e.busy,
                      title: r(
                        "Publish fixed snapshots, review linked refreshes and inspect embedded objects",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(FileSpreadsheet, {}),
                      label: r("Save as Excel"),
                      large: true,
                      onClick: () => e.onExport("xlsx"),
                      disabled: e.busy,
                      className: "xl-green-icon",
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Download, {}),
                      label: r("Export CSV"),
                      large: true,
                      onClick: () => e.onExport("csv"),
                      disabled: e.busy,
                    }),
                  ],
                }),
              ],
            }),
          J === "Review" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsx(Mn, {
                  label: r("Proofing"),
                  children: i.jsx(lr, {
                    icon: i.jsx(BookOpen, {}),
                    label: r("Spelling"),
                    large: true,
                    onClick: () => e.onTool("spelling"),
                    disabled: false,
                  }),
                }),
                i.jsx(Mn, {
                  label: r("Notes"),
                  children: i.jsx(lr, {
                    icon: i.jsx(MessageSquare, {}),
                    label: r("New / Edit Note"),
                    large: true,
                    onClick: e.onNote,
                    disabled: k,
                    className: "xl-gold-icon",
                  }),
                }),
                i.jsxs(Mn, {
                  label: r("Comments"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(MessageSquare, {}),
                      label: r("Review Comments"),
                      large: true,
                      onClick: () => e.onTool("comments"),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(MessageSquare, {}),
                      label: r("Hyperlink"),
                      onClick: () => e.onTool("hyperlink"),
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Audit Review"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(ListFilter, {}),
                      label: r("Pattern Checks"),
                      large: true,
                      onClick: () => e.onTool("patterns"),
                      disabled: e.busy,
                      title: r(
                        "Test explicit duplicate keys, sequence gaps, round amounts, weekends, holidays and combined rules",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(ListFilter, {}),
                      label: r("Aging Schedule"),
                      large: true,
                      onClick: () => e.onTool("aging"),
                      disabled: e.busy,
                      title: r(
                        "Reconcile configurable aging buckets using a fixed as-of date",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(ListFilter, {}),
                      label: r("Risk Screening"),
                      large: true,
                      onClick: () => e.onTool("materiality"),
                      disabled: e.busy,
                      title: r(
                        "Review explicit thresholds, currency and effective period",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(CheckCheck, {}),
                      label: r("Mark Reviewed"),
                      large: true,
                      onClick: e.onReviewed,
                      disabled: k,
                      className: "xl-green-icon",
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(ClipboardCheck, {}),
                      label: r("Audit Checks"),
                      large: true,
                      onClick: () => e.onPanel("checks"),
                      className: "xl-green-icon",
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Changes"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(ArrowLeftRight, {}),
                      label: r("Compare Workbooks"),
                      large: true,
                      onClick: () => e.onTool("workbookComparison"),
                      disabled: e.busy,
                      title: r(
                        "Compare keyed rows, formulas, links and hidden content",
                      ),
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(History, {}),
                      label: r("Show Changes"),
                      large: true,
                      onClick: () => e.onPanel("history"),
                      className: "xl-blue-icon",
                    }),
                  ],
                }),
                i.jsx(Mn, {
                  label: r("Protection"),
                  children: i.jsx(lr, {
                    icon: i.jsx(ShieldCheck, {}),
                    label: e.sheetProtected
                      ? "Unprotect Sheet"
                      : "Protect Sheet",
                    large: true,
                    onClick: () => e.onTool("protection"),
                    disabled: !e.allowProtectionChange,
                  }),
                }),
                i.jsx(Mn, {
                  label: r("AI Review"),
                  children: i.jsx(lr, {
                    icon: i.jsx(Sparkles, {}),
                    label: r("AI Assistant"),
                    large: true,
                    onClick: () => e.onPanel("assistant"),
                    className: "xl-purple-icon",
                  }),
                }),
                i.jsx(Mn, {
                  label: r("Automation"),
                  children: i.jsx(lr, {
                    icon: i.jsx(FileCode, {}),
                    label: r("Automation & Scripts"),
                    large: true,
                    onClick: () => e.onTool("automation"),
                    disabled: e.busy,
                    title: r(
                      "Inspect and preserve VBA/Office Scripts; review bounded migrations",
                    ),
                  }),
                }),
              ],
            }),
          J === "View" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs(Mn, {
                  label: r("Workbook Views"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Grid3x3, {}),
                      label: r("Normal"),
                      large: true,
                      active: true,
                      onClick: () => {
                        (e.gridlines || e.onGridlines(),
                          e.headings || e.onHeadings());
                      },
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Sheet, {}),
                      label: r("Page Layout"),
                      large: true,
                      onClick: () => e.onTool("pageLayout"),
                      disabled: false,
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Show"),
                  className: "xl-options-group",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: e.gridlines,
                          onChange: e.onGridlines,
                        }),
                        i.jsx(M, { children: "Gridlines" }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: e.headings,
                          onChange: e.onHeadings,
                        }),
                        i.jsx(M, { children: "Headings" }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: e.showFormulas,
                          onChange: e.onShowFormulas,
                        }),
                        i.jsx(M, { children: "Formulas" }),
                      ],
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Zoom"),
                  className: "xl-zoom-group",
                  children: [
                    i.jsxs("div", {
                      className: "xl-zoom-selector",
                      children: [i.jsx(Search, { size: 22 }), B],
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Maximize, {}),
                      label: "100%",
                      large: true,
                      onClick: () => e.onZoom(100),
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Window"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Grid2x2, {}),
                      label: r("Split"),
                      onClick: () => e.onTool("viewSettings"),
                      title: r(
                        "Split the worksheet into two or four scrolling panes",
                      ),
                      disabled: k,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Rows3, {}),
                      label: r("Default Sizes"),
                      onClick: () => e.onTool("viewSettings"),
                      title: r(
                        "Set worksheet-wide row height and column width",
                      ),
                      disabled: k,
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Grid2x2, {}),
                      label: r("Freeze Panes"),
                      large: true,
                      onClick: () => e.onTool("viewSettings"),
                      title: r(
                        "Freeze rows and columns above and to the left of the selected cell",
                      ),
                      className: "xl-blue-icon",
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(PanelTop, {}),
                      label: r("Freeze Top Row"),
                      large: true,
                      active: e.freezeRows,
                      onClick: e.onFreeze,
                      className: "xl-blue-icon",
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(Rows3, {}),
                      label: r("Hide & Unhide"),
                      large: true,
                      onClick: () => e.onTool("viewSettings"),
                      title: r("Hide or restore selected rows and columns"),
                    }),
                  ],
                }),
                i.jsxs(Mn, {
                  label: r("Task Panes"),
                  children: [
                    i.jsx(lr, {
                      icon: i.jsx(Sparkles, {}),
                      label: r("AI Assistant"),
                      large: true,
                      onClick: () => e.onPanel("assistant"),
                      className: "xl-purple-icon",
                    }),
                    i.jsx(lr, {
                      icon: i.jsx(History, {}),
                      label: r("Show Changes"),
                      large: true,
                      onClick: () => e.onPanel("history"),
                    }),
                  ],
                }),
              ],
            }),
        ],
      }),
    N = Object.fromEntries(Uh.map((J) => [J, G7(T(J), J, C, _)])),
    R = G7(
      i.jsxs("div", {
        className: "xl-file-menu",
        role: "menu",
        "aria-label": r("File"),
        lang: t("File"),
        "data-ui-label": "File",
        children: [
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            onClick: () => V(e.onNew),
            disabled: k,
            children: [
              i.jsx(FilePlus2, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "New" }),
                  i.jsx("span", {
                    children: i.jsx(M, { children: "Blank workbook" }),
                  }),
                ],
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            onClick: () => V(e.onImport),
            disabled: k,
            children: [
              i.jsx(FolderOpen, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "Open" }),
                  i.jsx("span", {
                    children: i.jsx(M, {
                      children: "Excel, CSV, TSV, or JSON",
                    }),
                  }),
                ],
              }),
            ],
          }),
          i.jsx("div", { className: "xl-menu-divider" }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            onClick: () => V(() => e.onTool("sheets")),
            children: [
              i.jsx(Sheet, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "Manage worksheets" }),
                  i.jsx("span", {
                    children: i.jsx(M, {
                      children:
                        "Names, order, visibility and structure protection",
                    }),
                  }),
                ],
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            onClick: () => V(() => e.onTool("workbookProperties")),
            children: [
              i.jsx(FileSpreadsheet, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "Workbook properties" }),
                  i.jsx("span", {
                    children: i.jsx(M, {
                      children: "Engagement, period and document details",
                    }),
                  }),
                ],
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            onClick: () => V(() => e.onTool("versions")),
            children: [
              i.jsx(History, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "Version history" }),
                  i.jsx("span", {
                    children: i.jsx(M, {
                      children: "Save, compare and restore checkpoints",
                    }),
                  }),
                ],
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            disabled: !e.allowProtectionChange || e.busy,
            onClick: () => V(() => e.onTool("template")),
            children: [
              i.jsx(FilePlus2, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "New from current template" }),
                  i.jsx("span", {
                    children: i.jsx(M, {
                      children: "Create an independent workbook download",
                    }),
                  }),
                ],
              }),
            ],
          }),
          i.jsx("div", { className: "xl-menu-divider" }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            onClick: () => V(() => e.onExport("xlsx")),
            disabled: e.busy,
            children: [
              i.jsx(FileSpreadsheet, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "Save as Excel" }),
                  i.jsx("span", {
                    children: i.jsx(M, { children: "Excel Workbook (.xlsx)" }),
                  }),
                ],
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            onClick: () => V(() => e.onExport("csv")),
            disabled: e.busy,
            children: [
              i.jsx(Download, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "Save as CSV" }),
                  i.jsx("span", {
                    children: i.jsx(M, {
                      children: "Current worksheet (.csv)",
                    }),
                  }),
                ],
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            onClick: () => V(() => e.onExport("json")),
            disabled: e.busy,
            children: [
              i.jsx(FileJson, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "Save workbook data" }),
                  i.jsx("span", {
                    children: i.jsx(M, {
                      children: "Workbook, notes, and history (.json)",
                    }),
                  }),
                ],
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            disabled: e.busy,
            onClick: () => V(() => e.onTool("xlsxSource")),
            children: [
              i.jsx(FileSpreadsheet, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "Original XLSX and compatibility" }),
                  i.jsx("span", {
                    children: i.jsx(M, {
                      children: "Source file and part inventory",
                    }),
                  }),
                ],
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            disabled: e.busy,
            onClick: () => V(() => e.onTool("importSources")),
            children: [
              i.jsx(FileSpreadsheet, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "Import sources" }),
                  i.jsx("span", {
                    children: i.jsx(M, {
                      children: "JSON, XML and document extraction",
                    }),
                  }),
                ],
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            disabled: e.busy,
            onClick: () => V(() => e.onTool("workbookFormats")),
            children: [
              i.jsx(FileSpreadsheet, { size: 18 }),
              i.jsxs("span", {
                children: [
                  i.jsx(M, { children: "Workbook formats" }),
                  i.jsx("span", {
                    children: i.jsx(M, {
                      children: "Templates and legacy conversion",
                    }),
                  }),
                ],
              }),
            ],
          }),
          i.jsx("div", { className: "xl-menu-divider" }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            onClick: () => V(e.onRename),
            disabled: k,
            children: [
              i.jsx(Type, { size: 18 }),
              i.jsx("span", {
                children: i.jsx(M, { children: "Rename workbook" }),
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            onClick: () => V(e.onDemo),
            disabled: k,
            children: [
              i.jsx(ClipboardCheck, { size: 18 }),
              i.jsx("span", {
                children: i.jsx(M, { children: "Open audit example" }),
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            disabled: e.busy,
            onClick: () => V(() => e.onTool("automation")),
            children: [
              i.jsx(FileCode, { size: 18 }),
              i.jsx("span", {
                children: i.jsx(M, { children: "Automation & Scripts" }),
              }),
            ],
          }),
          i.jsxs("button", {
            type: "button",
            role: "menuitem",
            onClick: () => V(() => e.onPanel("help")),
            children: [
              i.jsx(BookOpen, { size: 18 }),
              i.jsx("span", {
                children: i.jsx(M, { children: "Workbook help" }),
              }),
            ],
          }),
        ],
      }),
      "File",
      C,
      _,
    ),
    L = [
      {
        id: "quick-undo",
        name: "Undo",
        location: "Quick access",
        reason: k || !e.canUndo ? _("Undo") : void 0,
      },
      {
        id: "quick-redo",
        name: "Redo",
        location: "Quick access",
        reason: k || !e.canRedo ? _("Redo") : void 0,
      },
    ],
    $ = (J) => {
      J.reason ||
        (s(false),
        (d.current = J.id),
        J.id.startsWith("File:")
          ? (l(true), p(false))
          : J.id.startsWith("quick-") || A(J.id.split(":")[0]));
    };
  return (
    ne.useEffect(() => {
      var W;
      const J = d.current;
      if (!J) return;
      const q = Array.from(
        ((W = m.current) == null
          ? void 0
          : W.querySelectorAll("[data-workbook-command]")) ?? [],
      ).find((K) => K.dataset.workbookCommand === J);
      q && (q.focus(), (d.current = null));
    }),
    ne.useEffect(() => {
      const J = (q) => {
        if (
          q.defaultPrevented ||
          q.repeat ||
          q.key !== "F1" ||
          q.ctrlKey ||
          q.metaKey ||
          q.altKey ||
          q.shiftKey
        )
          return;
        const W = q.target,
          K = m.current;
        if (!(W instanceof Element) || !K || W.closest('[role="dialog"]'))
          return;
        const te = K.closest(".lg-app") ?? K;
        (te.classList.contains("lg-app")
          ? W.closest(".lg-app") !== te
          : !te.contains(W)) ||
          (q.preventDefault(), q.stopPropagation(), s(true));
      };
      return (
        document.addEventListener("keydown", J),
        () => document.removeEventListener("keydown", J)
      );
    }, []),
    i.jsxs(i.Fragment, {
      children: [
        a &&
          i.jsx(Vde, {
            commands: [
              ...L,
              ...R.commands,
              ...Uh.flatMap((J) => N[J].commands),
            ],
            onClose: () => s(false),
            onLocate: $,
          }),
        i.jsxs("div", {
          ref: m,
          role: "group",
          className: `xl-ribbon${c ? " xl-ribbon-collapsed" : ""}`,
          "aria-label": r("Workbook ribbon"),
          onKeyDown: (J) => {
            J.key === "Escape" &&
              f &&
              !J.defaultPrevented &&
              (J.preventDefault(), J.stopPropagation(), w(), p(false));
          },
          lang: t("Workbook ribbon"),
          "data-ui-label": "Workbook ribbon",
          children: [
            i.jsxs("div", {
              className: "xl-ribbon-tabbar",
              children: [
                i.jsxs("div", {
                  className: "xl-file-control",
                  ref: g,
                  onKeyDown: (J) => {
                    var q, W, K, te;
                    if (!(
                      J.altKey ||
                      J.ctrlKey ||
                      J.metaKey ||
                      J.defaultPrevented
                    ))
                      if (
                        ["ArrowDown", "ArrowUp", "Home", "End"].includes(J.key)
                      ) {
                        if ((J.preventDefault(), !u)) {
                          l(true);
                          const he = R.commands.find((Se) => !Se.reason);
                          he && (d.current = he.id);
                          return;
                        }
                        const ae = Array.from(
                            ((q = g.current) == null
                              ? void 0
                              : q.querySelectorAll(
                                  '[role="menuitem"]:not(:disabled)',
                                )) ?? [],
                          ),
                          oe = ae.indexOf(document.activeElement),
                          de =
                            J.key === "Home"
                              ? 0
                              : J.key === "End"
                                ? ae.length - 1
                                : (oe +
                                    (J.key === "ArrowDown"
                                      ? 1
                                      : ae.length - 1)) %
                                  ae.length;
                        (W = ae[de]) == null || W.focus();
                      } else
                        J.key === "Escape" &&
                          (J.preventDefault(),
                          J.stopPropagation(),
                          l(false),
                          (te =
                            (K = g.current) == null
                              ? void 0
                              : K.querySelector("button")) == null ||
                            te.focus());
                  },
                  children: [
                    i.jsx("button", {
                      type: "button",
                      className: `xl-file-tab ${u ? "xl-tab-open" : ""}`,
                      "aria-expanded": u,
                      "aria-haspopup": "menu",
                      onClick: () => {
                        (l(!u), p(false));
                      },
                      children: i.jsx(M, { children: "File" }),
                    }),
                    u && R.content,
                  ],
                }),
                i.jsx("div", {
                  className: "xl-ribbon-tabs",
                  role: "tablist",
                  "aria-label": r("Ribbon tabs"),
                  lang: t("Ribbon tabs"),
                  "data-ui-label": "Ribbon tabs",
                  children: Uh.map((J, q) =>
                    i.jsx(
                      "button",
                      {
                        type: "button",
                        id: `${b}-${q}`,
                        role: "tab",
                        "aria-selected": n === J,
                        "aria-controls": `${b}-panel`,
                        "aria-expanded": n === J && y,
                        tabIndex: n === J ? 0 : -1,
                        className: n === J ? "xl-tab-active" : "",
                        title: r(
                          "Double-click to collapse or expand the ribbon",
                        ),
                        onClick: () => A(J),
                        onDoubleClick: E,
                        onKeyDown: (W) => {
                          var K, te;
                          if (!(
                            W.defaultPrevented ||
                            W.altKey ||
                            W.ctrlKey ||
                            W.metaKey
                          )) {
                            if (
                              W.key === "ArrowRight" ||
                              W.key === "ArrowLeft"
                            ) {
                              W.preventDefault();
                              const ae =
                                (q +
                                  (W.key === "ArrowRight"
                                    ? 1
                                    : Uh.length - 1)) %
                                Uh.length;
                              (A(Uh[ae]),
                                (K = document.getElementById(`${b}-${ae}`)) ==
                                  null || K.focus());
                            } else if (W.key === "Home" || W.key === "End") {
                              W.preventDefault();
                              const ae = W.key === "Home" ? 0 : Uh.length - 1;
                              (A(Uh[ae]),
                                (te = document.getElementById(`${b}-${ae}`)) ==
                                  null || te.focus());
                            } else if (W.key === "ArrowDown") {
                              (W.preventDefault(), A(J));
                              const ae = N[J].commands.find((oe) => !oe.reason);
                              ae && (d.current = ae.id);
                            }
                          }
                        },
                        lang: t(
                          "Double-click to collapse or expand the ribbon",
                        ),
                        children: i.jsx(M, { children: J }),
                      },
                      J,
                    ),
                  ),
                }),
                i.jsxs("div", {
                  className: "xl-ribbon-quick-actions",
                  children: [
                    i.jsx("button", {
                      type: "button",
                      className: "xl-icon-button xl-ribbon-toggle",
                      "aria-label": r(c ? "Expand ribbon" : "Collapse ribbon"),
                      "aria-expanded": y,
                      "aria-controls": `${b}-panel`,
                      title: c
                        ? "Expand and keep ribbon open (Ctrl+F1)"
                        : "Collapse ribbon (Ctrl+F1)",
                      onClick: E,
                      children: c
                        ? i.jsx(ChevronDown, { size: 16 })
                        : i.jsx(ChevronUp, { size: 16 }),
                    }),
                    i.jsx(Xl, {
                      icon: i.jsx(Undo2, { size: 16 }),
                      label: r("Undo (Ctrl+Z)"),
                      commandId: "quick-undo",
                      reason: L[0].reason,
                      onClick: e.onUndo,
                      disabled: k || !e.canUndo,
                    }),
                    i.jsx(Xl, {
                      icon: i.jsx(Redo2, { size: 16 }),
                      label: r("Redo (Ctrl+Y)"),
                      commandId: "quick-redo",
                      reason: L[1].reason,
                      onClick: e.onRedo,
                      disabled: k || !e.canRedo,
                    }),
                    i.jsx(Xl, {
                      icon: i.jsx(BookOpen, { size: 16 }),
                      label: r("Commands & shortcuts"),
                      onClick: () => s(true),
                    }),
                    i.jsx("span", { className: "xl-tabbar-divider" }),
                    i.jsxs("button", {
                      type: "button",
                      className: "xl-ai-button",
                      "aria-label": r("AI assistant"),
                      onClick: () => e.onPanel("assistant"),
                      lang: t("AI assistant"),
                      "data-ui-label": "AI assistant",
                      children: [
                        i.jsx(Sparkles, { size: 15 }),
                        i.jsx("span", {
                          children: i.jsx(M, { children: "AI Assistant" }),
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
            i.jsx("div", {
              ref: v,
              id: `${b}-panel`,
              className: `xl-ribbon-body${c && f ? " xl-ribbon-temporary" : ""}`,
              hidden: !y,
              role: "tabpanel",
              "aria-labelledby": `${b}-${Uh.indexOf(n)}`,
              children: y && i.jsx(i.Fragment, { children: N[n].content }),
            }),
          ],
        }),
      ],
    })
  );
}

export { Hde };
