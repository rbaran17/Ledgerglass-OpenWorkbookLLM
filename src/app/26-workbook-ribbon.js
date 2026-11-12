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
import {
  summarizeNumberFormat,
  buildNumberFormatCode,
  validateNumberFormatSettings,
  resolveEffectiveCellStyle,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  collectRibbonCommands,
  RibbonGroup,
  Uh,
  CommandReferenceDialog,
  RibbonIconButton,
  RibbonCommandButton,
} from "./24-merge-audit-and-ribbon-helpers.js";
import { ColorPickerButton } from "./25-color-picker.js";

function WorkbookRibbon(e) {
  const { text, language } = useWorkbookMessages(),
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
    w = () =>
      m.current?.querySelector('[role="tab"][aria-selected="true"]')?.focus(),
    E = () => {
      (v.current?.contains(document.activeElement) && w(),
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
      m.current?.contains(q.target) || p(false);
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
          te = K?.closest(".lg-app") ?? K;
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
    F = resolveEffectiveCellStyle(e.cell?.style),
    I = (J) => e.onFormat({ style: J }),
    O = (J) =>
      e.onFormat({ format: J, numFmt: undefined, numberFormat: undefined }),
    j =
      e.cell?.numberFormat &&
      !["number", "currency", "accounting", "percent"].includes(
        e.cell.numberFormat.category,
      ),
    z = (J) => {
      if (e.cell?.numberFormat && !j) {
        const le = validateNumberFormatSettings({
          ...e.cell.numberFormat,
          decimals: Math.max(
            0,
            Math.min(12, (e.cell.numberFormat.decimals ?? 2) + J),
          ),
        });
        e.onFormat({ numberFormat: le, numFmt: buildNumberFormatCode(le) });
        return;
      }
      const q =
          summarizeNumberFormat(e.cell?.numFmt) ??
          (e.cell?.format === "currency" ||
          e.cell?.format === "number" ||
          e.cell?.format === "percent"
            ? 2
            : 0),
        W = Math.max(0, Math.min(10, q + J));
      e.onFormat({
        numberFormat: undefined,
        numFmt: `${e.cell?.format === "currency" ? "$" : ""}#,##0${W ? `.${"0".repeat(W)}` : ""}${e.cell?.format === "percent" ? "%" : ""}`,
      });
    };
  ne.useEffect(() => {
    if (!u) return;
    const J = (q) => {
      g.current?.contains(q.target) || l(false);
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
      "aria-label": text("Number format"),
      value: e.cell?.numberFormat ? "advanced" : (e.cell?.format ?? "general"),
      disabled: k,
      onChange: (J) =>
        J.target.value === "advanced"
          ? e.onTool("numberFormat")
          : O(J.target.value),
      lang: language("Number format"),
      "data-ui-label": "Number format",
      children: [
        e.cell?.numberFormat &&
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
      "aria-label": text("Zoom"),
      value: e.zoom,
      onChange: (J) => e.onZoom(Number(J.target.value)),
      lang: language("Zoom"),
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
    C = {
      group: RibbonGroup,
      command: RibbonCommandButton,
      icon: RibbonIconButton,
      color: ColorPickerButton,
    },
    T = (J) =>
      i.jsxs(i.Fragment, {
        children: [
          J === "Home" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs(RibbonGroup, {
                  label: text("Clipboard"),
                  className: "xl-clipboard-group",
                  children: [
                    i.jsxs("div", {
                      className: "xl-paste-commands",
                      children: [
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Clipboard, {}),
                          label: text("Paste"),
                          large: true,
                          onClick: e.onPaste,
                          disabled: k,
                          className: "xl-gold-icon xl-paste-primary",
                        }),
                        i.jsxs("button", {
                          type: "button",
                          className: "xl-paste-special",
                          "aria-label": text("Paste Special"),
                          title: text(
                            "Choose values, formulas, formatting and other paste options",
                          ),
                          disabled: k,
                          onClick: () => e.onTool("pasteSpecial"),
                          lang: language("Paste Special"),
                          "data-ui-label": "Paste Special",
                          children: [
                            i.jsx(LocalizedText, { children: "Paste Special" }),
                            i.jsx(ChevronDown, { size: 10 }),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "xl-command-stack",
                      children: [
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Scissors, {}),
                          label: text("Cut"),
                          onClick: e.onCut,
                          disabled: k,
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Copy, {}),
                          label: text("Copy"),
                          onClick: e.onCopy,
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Paintbrush, {}),
                          label: text("Format Painter"),
                          onClick: () => e.onTool("paint"),
                          disabled: k,
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Font"),
                  className: "xl-font-group",
                  onDialog: () => e.onTool("cellStyle"),
                  dialogLabel: text("Cell Styles"),
                  dialogDisabled: k,
                  children: [
                    i.jsxs("div", {
                      className: "xl-font-selects",
                      children: [
                        i.jsx("select", {
                          "aria-label": text("Font family"),
                          value: F.fontFamily ?? "Calibri",
                          disabled: k,
                          onChange: (q) => I({ fontFamily: q.target.value }),
                          lang: language("Font family"),
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
                          "aria-label": text("Font size"),
                          value: F.fontSize ?? 11,
                          disabled: k,
                          onChange: (q) =>
                            I({ fontSize: Number(q.target.value) }),
                          lang: language("Font size"),
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
                        i.jsx(RibbonIconButton, {
                          icon: i.jsxs("span", {
                            className: "xl-grow-font",
                            children: [
                              i.jsx(LocalizedText, { children: "A" }),
                              i.jsx("sup", { children: "⌃" }),
                            ],
                          }),
                          label: text("Increase Font Size"),
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
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(Bold, { size: 15 }),
                          label: text("Bold (Ctrl+B)"),
                          active: !!F.bold,
                          onClick: () => I({ bold: !F.bold }),
                          disabled: k,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(Italic, { size: 15 }),
                          label: text("Italic (Ctrl+I)"),
                          active: !!F.italic,
                          onClick: () => I({ italic: !F.italic }),
                          disabled: k,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(Underline, { size: 15 }),
                          label: text("Underline (Ctrl+U)"),
                          active: !!F.underline,
                          onClick: () => I({ underline: !F.underline }),
                          disabled: k,
                        }),
                        i.jsx("span", { className: "xl-control-divider" }),
                        i.jsxs("select", {
                          className: "xl-border-select",
                          "aria-label": text("Borders"),
                          title: text("Borders"),
                          value: F.border ?? "none",
                          disabled: k,
                          onChange: (q) =>
                            I({ border: q.target.value, borders: undefined }),
                          lang: language("Borders"),
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
                        i.jsx(ColorPickerButton, {
                          kind: "fill",
                          value: F.fill ?? "#FFFF00",
                          onChange: (q) => I({ fill: q }),
                          disabled: k,
                        }),
                        i.jsx(ColorPickerButton, {
                          kind: "color",
                          value: F.color ?? "#C00000",
                          onChange: (q) => I({ color: q }),
                          disabled: k,
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Alignment"),
                  className: "xl-alignment-group",
                  children: [
                    i.jsxs("div", {
                      className: "xl-alignment-top",
                      children: [
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(AlignLeft, { size: 16 }),
                          label: text("Align Left"),
                          active: F.align === "left",
                          onClick: () => I({ align: "left" }),
                          disabled: k,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(AlignCenter, { size: 16 }),
                          label: text("Center"),
                          active: F.align === "center",
                          onClick: () => I({ align: "center" }),
                          disabled: k,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(AlignRight, { size: 16 }),
                          label: text("Align Right"),
                          active: F.align === "right",
                          onClick: () => I({ align: "right" }),
                          disabled: k,
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(WrapText, {}),
                          label: text("Wrap Text"),
                          active: !!F.wrap,
                          onClick: () => I({ wrap: !F.wrap }),
                          disabled: k,
                        }),
                      ],
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid2x2, {}),
                      label: text("Merge & Center"),
                      onClick: () => e.onTool("merge"),
                      disabled: k,
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Number"),
                  className: "xl-number-group",
                  children: [
                    D,
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(DollarSign, { size: 15 }),
                      label: text("Format Cells"),
                      onClick: () => e.onTool("numberFormat"),
                      disabled: k,
                    }),
                    i.jsxs("div", {
                      className: "xl-number-buttons",
                      children: [
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(DollarSign, { size: 16 }),
                          label: text("Currency"),
                          onClick: () => O("currency"),
                          disabled: k,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(Percent, { size: 15 }),
                          label: text("Percent Style"),
                          onClick: () => O("percent"),
                          disabled: k,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx("span", {
                            className: "xl-comma-icon",
                            children: ",",
                          }),
                          label: text("Comma Style"),
                          onClick: () =>
                            e.onFormat({
                              format: "number",
                              numFmt: "#,##0.00",
                              numberFormat: undefined,
                            }),
                          disabled: k,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsxs("span", {
                            className: "xl-decimal-icon",
                            children: [
                              "←",
                              i.jsx("small", { children: ".00" }),
                            ],
                          }),
                          label: text("Increase Decimal"),
                          onClick: () => z(1),
                          disabled: k || !!j,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsxs("span", {
                            className: "xl-decimal-icon",
                            children: ["→", i.jsx("small", { children: ".0" })],
                          }),
                          label: text("Decrease Decimal"),
                          onClick: () => z(-1),
                          disabled: k || !!j,
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Styles"),
                  className: "xl-styles-group",
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid3x3, {}),
                      label: text("Conditional Formatting"),
                      large: true,
                      onClick: () => e.onTool("conditional"),
                      disabled: k,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Format as Table"),
                      large: true,
                      onClick: () => e.onTool("table"),
                      disabled: k,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(LayoutGrid, {}),
                      label: text("Heading Style"),
                      large: true,
                      onClick: () =>
                        I({ fill: "#E2EFDA", color: "#375623", bold: true }),
                      disabled: k,
                      title: text(
                        "Apply green heading cell style to the selection",
                      ),
                      className: "xl-green-icon",
                    }),
                  ],
                }),
                i.jsx(RibbonGroup, {
                  label: text("Cells"),
                  className: "xl-cells-group",
                  children: i.jsxs("div", {
                    className: "xl-command-stack",
                    children: [
                      i.jsx(RibbonCommandButton, {
                        icon: i.jsx(Rows3, {}),
                        label: text("Insert Row"),
                        onClick: e.onAddRow,
                        disabled: k,
                        className: "xl-green-icon",
                      }),
                      i.jsx(RibbonCommandButton, {
                        icon: i.jsx(Table2, {}),
                        label: text("Insert Column"),
                        onClick: e.onAddColumn,
                        disabled: k,
                        className: "xl-green-icon",
                      }),
                      i.jsx(RibbonCommandButton, {
                        icon: i.jsx(Sheet, {}),
                        label: text("New Sheet"),
                        onClick: e.onAddSheet,
                        disabled:
                          e.busy ||
                          !e.allowProtectionChange ||
                          e.workbookStructureProtected,
                      }),
                    ],
                  }),
                }),
                i.jsx(RibbonGroup, {
                  label: text("Cell Size"),
                  children: i.jsxs("div", {
                    className: "xl-command-stack",
                    children: [
                      i.jsx(RibbonCommandButton, {
                        icon: i.jsx(Maximize, {}),
                        label: text("AutoFit Column Width"),
                        onClick: e.onAutoFitColumns,
                        disabled: k,
                        title: text(
                          "Fit the selected columns to their contents",
                        ),
                      }),
                      i.jsx(RibbonCommandButton, {
                        icon: i.jsx(Rows3, {}),
                        label: text("AutoFit Row Height"),
                        onClick: e.onAutoFitRows,
                        disabled: k,
                        title: text(
                          "Fit the selected rows to their contents and wrapping",
                        ),
                      }),
                      i.jsx(RibbonCommandButton, {
                        icon: i.jsx(Rows3, {}),
                        label: text("Hide & Unhide"),
                        onClick: () => e.onTool("viewSettings"),
                        title: text(
                          "Hide or restore selected rows and columns",
                        ),
                      }),
                    ],
                  }),
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Editing"),
                  className: "xl-editing-group",
                  children: [
                    i.jsxs("div", {
                      className: "xl-command-stack",
                      children: [
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Sigma, {}),
                          label: text("AutoSum"),
                          onClick: e.onAutosum,
                          disabled: k,
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(ArrowDownToLine, {}),
                          label: text("Fill Down"),
                          onClick: e.onFillDown,
                          disabled: k,
                          className: "xl-blue-icon",
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Eraser, {}),
                          label: text("Clear"),
                          onClick: e.onClear,
                          disabled: k,
                        }),
                      ],
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Fill Series"),
                      large: true,
                      onClick: () => e.onTool("fillSeries"),
                      disabled: k,
                      title: text(
                        "Fill in any direction or create a number or date series",
                      ),
                      className: "xl-blue-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ListFilter, {}),
                      label: text("Sort & Filter"),
                      large: true,
                      onClick: e.onToggleFilter,
                      className: "xl-blue-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Search, {}),
                      label: text("Find & Select"),
                      large: true,
                      onClick: e.onFind,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Search, {}),
                      label: text("Advanced Find"),
                      onClick: () => e.onTool("search"),
                      title: text(
                        "Find all matches, preview replacements, or inspect special cells (Ctrl/⌘+Shift+F)",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Search, {}),
                      label: text("Go To Special"),
                      onClick: () => e.onTool("goToSpecial"),
                      title: text(
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
                i.jsxs(RibbonGroup, {
                  label: text("Tables"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("PivotTable"),
                      large: true,
                      onClick: () => e.onTool("pivot"),
                      disabled: k,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Table"),
                      large: true,
                      onClick: () => e.onTool("table"),
                      disabled: k,
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Illustrations"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Image$1, {}),
                      label: text("Pictures"),
                      large: true,
                      onClick: () => e.onTool("picture"),
                      disabled: k,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Square, {}),
                      label: text("Shapes"),
                      large: true,
                      onClick: () => e.onTool("shape"),
                      disabled: k,
                    }),
                  ],
                }),
                i.jsx(RibbonGroup, {
                  label: text("Charts"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(ChartColumn, {}),
                    label: text("Recommended Charts"),
                    large: true,
                    onClick: () => e.onTool("chart"),
                    disabled: k,
                  }),
                }),
                i.jsx(RibbonGroup, {
                  label: text("Comments"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(MessageSquare, {}),
                    label: text("New Note"),
                    large: true,
                    onClick: e.onNote,
                    disabled: k,
                    className: "xl-gold-icon",
                  }),
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Worksheet"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Sheet, {}),
                      label: text("New Sheet"),
                      large: true,
                      onClick: e.onAddSheet,
                      disabled:
                        e.busy ||
                        !e.allowProtectionChange ||
                        e.workbookStructureProtected,
                      className: "xl-green-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Insert Row"),
                      large: true,
                      onClick: e.onAddRow,
                      disabled: k,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Insert Column"),
                      large: true,
                      onClick: e.onAddColumn,
                      disabled: k,
                    }),
                  ],
                }),
                i.jsx(RibbonGroup, {
                  label: text("Files"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(FolderOpen, {}),
                    label: text("Open Workbook"),
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
                i.jsx(RibbonGroup, {
                  label: text("Themes"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(LayoutGrid, {}),
                    label: text("Themes"),
                    large: true,
                    onClick: () => e.onTool("themes"),
                    disabled: k,
                  }),
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Page Setup"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Sheet, {}),
                      label: text("Margins"),
                      large: true,
                      onClick: () => e.onTool("margins"),
                      disabled: false,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Orientation"),
                      large: true,
                      onClick: () => e.onTool("orientation"),
                      disabled: false,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Square, {}),
                      label: text("Size"),
                      large: true,
                      onClick: () => e.onTool("size"),
                      disabled: false,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid2x2, {}),
                      label: text("Print Area"),
                      large: true,
                      onClick: () => e.onTool("printArea"),
                      disabled: false,
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Print titles and breaks"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Print Titles"),
                      onClick: () => e.onTool("printTitles"),
                      disabled: false,
                      large: true,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid2x2, {}),
                      label: text("Page Breaks"),
                      onClick: () => e.onTool("pageBreaks"),
                      disabled: false,
                      large: true,
                    }),
                  ],
                }),
                i.jsx(RibbonGroup, {
                  label: text("Scale to Fit"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(Maximize, {}),
                    label: text("Fit to Page"),
                    large: true,
                    onClick: () => e.onTool("fit"),
                    disabled: false,
                  }),
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Sheet Options"),
                  className: "xl-options-group",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: e.gridlines,
                          onChange: e.onGridlines,
                        }),
                        i.jsx(LocalizedText, { children: "Gridlines" }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: e.headings,
                          onChange: e.onHeadings,
                        }),
                        i.jsx(LocalizedText, { children: "Headings" }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          J === "Formulas" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs(RibbonGroup, {
                  label: text("Function Library"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Sigma, {}),
                      label: text("AutoSum"),
                      large: true,
                      onClick: e.onAutosum,
                      disabled: k,
                      className: "xl-blue-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(BookOpen, {}),
                      label: text("Function Reference"),
                      large: true,
                      onClick: () => e.onPanel("help"),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Calculator, {}),
                      label: text("Insert Function"),
                      large: true,
                      onClick: () => e.onTool("function"),
                      disabled: false,
                    }),
                  ],
                }),
                i.jsx(RibbonGroup, {
                  label: text("Extensions"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(Calculator, {}),
                    label: text("Add-ins & Functions"),
                    large: true,
                    onClick: () => e.onTool("extensions"),
                    disabled: e.busy,
                  }),
                }),
                i.jsx(RibbonGroup, {
                  label: text("Defined Names"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(Table2, {}),
                    label: text("Name Manager"),
                    large: true,
                    onClick: () => e.onTool("names"),
                    disabled: false,
                  }),
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Formula Auditing"),
                  children: [
                    i.jsxs("div", {
                      className: "xl-command-stack",
                      children: [
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(ArrowLeftRight, {}),
                          label: text("Trace Precedents"),
                          onClick: () => e.onTool("tracePrecedents"),
                          title: text(
                            "Inspect the cells referenced by the active formula",
                          ),
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(ArrowLeftRight, {}),
                          label: text("Trace Dependents"),
                          onClick: () => e.onTool("traceDependents"),
                          title: text(
                            "Inspect formulas that reference the active cell",
                          ),
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Sheet, {}),
                          label: text("Watch Window"),
                          onClick: () => e.onTool("formulaWatches"),
                          title: text(
                            "Track selected cells and their current calculated values",
                          ),
                        }),
                      ],
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Calculator, {}),
                      label: text("Evaluate Formula"),
                      large: true,
                      onClick: () => e.onTool("evaluateFormula"),
                      title: text(
                        "Replay actual calculation steps and inspect errors",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ShieldCheck, {}),
                      label: text("Compatibility"),
                      large: true,
                      onClick: () => e.onTool("functionCompatibility"),
                      title: text(
                        "Review function contracts and unsupported formulas",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Sheet, {}),
                      label: text("Show Formulas"),
                      large: true,
                      active: e.showFormulas,
                      onClick: e.onShowFormulas,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ShieldCheck, {}),
                      label: text("Audit Checks"),
                      large: true,
                      onClick: () => e.onPanel("checks"),
                      className: "xl-green-icon",
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Calculation"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Calculator, {}),
                      label: text("Calculation Options"),
                      large: true,
                      onClick: () => e.onTool("calculation"),
                      title: text(
                        "Automatic or manual mode, iteration and calculation evidence",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Calculator, {}),
                      label: text("Calculate Now"),
                      large: true,
                      onClick: () => e.onTool("calculateNow"),
                      title: text(
                        "Recalculate the workbook without changing source cells",
                      ),
                    }),
                  ],
                }),
                i.jsx(RibbonGroup, {
                  label: text("Formula Assistance"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(Sparkles, {}),
                    label: text("Ask AI Assistant"),
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
                i.jsxs(RibbonGroup, {
                  label: text("Get & Transform Data"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(FolderOpen, {}),
                      label: text("Get Data"),
                      large: true,
                      onClick: e.onImport,
                      disabled: k,
                      className: "xl-gold-icon",
                      title: text("Import Excel, CSV, TSV, or workbook JSON"),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("From Table / Range"),
                      large: true,
                      onClick: () => e.onTool("query"),
                      disabled: k,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ArrowLeftRight, {}),
                      label: text("Append / Join"),
                      large: true,
                      onClick: () => e.onTool("combinePopulations"),
                      disabled: e.busy,
                      title: text(
                        "Review combined populations, unmatched records and match counts",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Transformation Recipes"),
                      large: true,
                      onClick: () => e.onTool("transformationRecipes"),
                      disabled: e.busy,
                      title: text(
                        "Build versioned steps with source fingerprints, failure details and reviewed outputs",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(FolderOpen, {}),
                      label: text("Connections & Refresh"),
                      large: true,
                      onClick: () => e.onTool("externalSources"),
                      disabled: e.busy,
                      title: text(
                        "Review approved external sources, control changes and last successful data",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Data Model"),
                      large: true,
                      onClick: () => e.onTool("dataModel"),
                      disabled: e.busy,
                      title: text(
                        "Build typed tables, relationships and measures across shared dimensions",
                      ),
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Sort & Filter"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Filter, {}),
                      label: text("Column Filters"),
                      large: true,
                      onClick: () => e.onTool("populationFilter"),
                      disabled: e.busy,
                      title: text(
                        "Filter values, text, numbers, dates and advanced criteria while retaining source rows",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ArrowDownAZ, {}),
                      label: text("Custom Sort"),
                      large: true,
                      onClick: () => e.onTool("sort"),
                      disabled: e.busy,
                      title: text(
                        "Review a stable multi-level sort of complete worksheet records",
                      ),
                    }),
                    i.jsxs("div", {
                      className: "xl-command-stack",
                      children: [
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(ArrowDownAZ, {}),
                          label: text("Sort A to Z"),
                          onClick: () => e.onSort("asc"),
                          className: "xl-blue-icon",
                          title: text(
                            "Sort worksheet view ascending by the selected column",
                          ),
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(ArrowUpAZ, {}),
                          label: text("Sort Z to A"),
                          onClick: () => e.onSort("desc"),
                          className: "xl-blue-icon",
                          title: text(
                            "Sort worksheet view descending by the selected column",
                          ),
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Eraser, {}),
                          label: text("Clear Sort"),
                          onClick: () => e.onSort(null),
                        }),
                      ],
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Filter, {}),
                      label: text("Filter"),
                      large: true,
                      onClick: e.onToggleFilter,
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Data Tools"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ClipboardCheck, {}),
                      label: text("Audit Checks"),
                      large: true,
                      onClick: () => e.onPanel("checks"),
                      className: "xl-green-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(CheckCheck, {}),
                      label: text("Data Validation"),
                      large: true,
                      onClick: () => e.onTool("validation"),
                      disabled: e.busy,
                      title: text(
                        "Review validation rules and inspect existing values; editing requires write access",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Copy, {}),
                      label: text("Remove Duplicates"),
                      large: true,
                      onClick: () => e.onTool("duplicates"),
                      disabled: e.busy,
                      title: text(
                        "Review duplicate groups and create a separate deduplicated population",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Columns3, {}),
                      label: text("Text to Columns"),
                      large: true,
                      onClick: () => e.onTool("textToColumns"),
                      disabled: e.busy,
                      title: text(
                        "Preview split fields and types before creating a separate output",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Columns3, {}),
                      label: text("Clean up text"),
                      large: true,
                      onClick: () => e.onTool("textCleanup"),
                      disabled: e.busy,
                      title: text(
                        "Review cleanup and extraction steps while retaining original evidence",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Columns3, {}),
                      label: text("Clean & Reshape"),
                      large: true,
                      onClick: () => e.onTool("fieldTransform"),
                      disabled: e.busy,
                      title: text(
                        "Review splitting, conversions, gap filling and inferred patterns with retained exceptions",
                      ),
                    }),
                  ],
                }),
                i.jsx(RibbonGroup, {
                  label: text("Workbook Links"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(FileSpreadsheet, {}),
                    label: text("Workbook Links"),
                    large: true,
                    onClick: () => e.onTool("workbookLinks"),
                    disabled: e.busy,
                    title: text(
                      "Inspect, resolve, refresh, replace and break approved workbook links",
                    ),
                  }),
                }),
                i.jsxs(RibbonGroup, {
                  label: text("What-If Analysis"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Calculator, {}),
                      label: text("Scenario Manager"),
                      large: true,
                      onClick: () => e.onTool("scenarios"),
                      disabled: e.busy,
                      title: text(
                        "Compare named assumptions and review a scenario before applying it",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Data Table"),
                      large: true,
                      onClick: () => e.onTool("dataTable"),
                      disabled: e.busy,
                      title: text(
                        "Build a one- or two-variable sensitivity table with retained baseline assumptions",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Calculator, {}),
                      label: text("Goal Seek"),
                      large: true,
                      onClick: () => e.onTool("goalSeek"),
                      disabled: e.busy,
                      title: text(
                        "Find a numeric input that reaches a target formula result",
                      ),
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Analysis"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Calculator, {}),
                      label: text("Solver"),
                      large: true,
                      onClick: () => e.onTool("solver"),
                      disabled: e.busy,
                      title: text(
                        "Review bounded linear optimization with explicit constraints",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Data Analysis"),
                      large: true,
                      onClick: () => e.onTool("statistics"),
                      disabled: e.busy,
                      title: text(
                        "Descriptive statistics, regression, distributions and forecasts",
                      ),
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Outline"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Consolidate / Subtotals"),
                      large: true,
                      onClick: () => e.onTool("consolidation"),
                      disabled: e.busy,
                      title: text(
                        "Review cross-sheet or approved-workbook rollups, source context and grouped subtotals",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Group & Outline"),
                      large: true,
                      onClick: () => e.onTool("outlines"),
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Export"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Copy, {}),
                      label: text("Report Outputs"),
                      large: true,
                      onClick: () => e.onTool("reportOutputs"),
                      disabled: e.busy,
                      title: text(
                        "Publish fixed snapshots, review linked refreshes and inspect embedded objects",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(FileSpreadsheet, {}),
                      label: text("Save as Excel"),
                      large: true,
                      onClick: () => e.onExport("xlsx"),
                      disabled: e.busy,
                      className: "xl-green-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Download, {}),
                      label: text("Export CSV"),
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
                i.jsx(RibbonGroup, {
                  label: text("Proofing"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(BookOpen, {}),
                    label: text("Spelling"),
                    large: true,
                    onClick: () => e.onTool("spelling"),
                    disabled: false,
                  }),
                }),
                i.jsx(RibbonGroup, {
                  label: text("Notes"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(MessageSquare, {}),
                    label: text("New / Edit Note"),
                    large: true,
                    onClick: e.onNote,
                    disabled: k,
                    className: "xl-gold-icon",
                  }),
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Comments"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(MessageSquare, {}),
                      label: text("Review Comments"),
                      large: true,
                      onClick: () => e.onTool("comments"),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(MessageSquare, {}),
                      label: text("Hyperlink"),
                      onClick: () => e.onTool("hyperlink"),
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Audit Review"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ListFilter, {}),
                      label: text("Pattern Checks"),
                      large: true,
                      onClick: () => e.onTool("patterns"),
                      disabled: e.busy,
                      title: text(
                        "Test explicit duplicate keys, sequence gaps, round amounts, weekends, holidays and combined rules",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ListFilter, {}),
                      label: text("Aging Schedule"),
                      large: true,
                      onClick: () => e.onTool("aging"),
                      disabled: e.busy,
                      title: text(
                        "Reconcile configurable aging buckets using a fixed as-of date",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ListFilter, {}),
                      label: text("Risk Screening"),
                      large: true,
                      onClick: () => e.onTool("materiality"),
                      disabled: e.busy,
                      title: text(
                        "Review explicit thresholds, currency and effective period",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(CheckCheck, {}),
                      label: text("Mark Reviewed"),
                      large: true,
                      onClick: e.onReviewed,
                      disabled: k,
                      className: "xl-green-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ClipboardCheck, {}),
                      label: text("Audit Checks"),
                      large: true,
                      onClick: () => e.onPanel("checks"),
                      className: "xl-green-icon",
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Changes"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ArrowLeftRight, {}),
                      label: text("Compare Workbooks"),
                      large: true,
                      onClick: () => e.onTool("workbookComparison"),
                      disabled: e.busy,
                      title: text(
                        "Compare keyed rows, formulas, links and hidden content",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(History, {}),
                      label: text("Show Changes"),
                      large: true,
                      onClick: () => e.onPanel("history"),
                      className: "xl-blue-icon",
                    }),
                  ],
                }),
                i.jsx(RibbonGroup, {
                  label: text("Protection"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(ShieldCheck, {}),
                    label: e.sheetProtected
                      ? "Unprotect Sheet"
                      : "Protect Sheet",
                    large: true,
                    onClick: () => e.onTool("protection"),
                    disabled: !e.allowProtectionChange,
                  }),
                }),
                i.jsx(RibbonGroup, {
                  label: text("AI Review"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(Sparkles, {}),
                    label: text("AI Assistant"),
                    large: true,
                    onClick: () => e.onPanel("assistant"),
                    className: "xl-purple-icon",
                  }),
                }),
                i.jsx(RibbonGroup, {
                  label: text("Automation"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(FileCode, {}),
                    label: text("Automation & Scripts"),
                    large: true,
                    onClick: () => e.onTool("automation"),
                    disabled: e.busy,
                    title: text(
                      "Inspect and preserve VBA/Office Scripts; review bounded migrations",
                    ),
                  }),
                }),
              ],
            }),
          J === "View" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs(RibbonGroup, {
                  label: text("Workbook Views"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid3x3, {}),
                      label: text("Normal"),
                      large: true,
                      active: true,
                      onClick: () => {
                        (e.gridlines || e.onGridlines(),
                          e.headings || e.onHeadings());
                      },
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Sheet, {}),
                      label: text("Page Layout"),
                      large: true,
                      onClick: () => e.onTool("pageLayout"),
                      disabled: false,
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Show"),
                  className: "xl-options-group",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: e.gridlines,
                          onChange: e.onGridlines,
                        }),
                        i.jsx(LocalizedText, { children: "Gridlines" }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: e.headings,
                          onChange: e.onHeadings,
                        }),
                        i.jsx(LocalizedText, { children: "Headings" }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: e.showFormulas,
                          onChange: e.onShowFormulas,
                        }),
                        i.jsx(LocalizedText, { children: "Formulas" }),
                      ],
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Zoom"),
                  className: "xl-zoom-group",
                  children: [
                    i.jsxs("div", {
                      className: "xl-zoom-selector",
                      children: [i.jsx(Search, { size: 22 }), B],
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Maximize, {}),
                      label: "100%",
                      large: true,
                      onClick: () => e.onZoom(100),
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Window"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid2x2, {}),
                      label: text("Split"),
                      onClick: () => e.onTool("viewSettings"),
                      title: text(
                        "Split the worksheet into two or four scrolling panes",
                      ),
                      disabled: k,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Default Sizes"),
                      onClick: () => e.onTool("viewSettings"),
                      title: text(
                        "Set worksheet-wide row height and column width",
                      ),
                      disabled: k,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid2x2, {}),
                      label: text("Freeze Panes"),
                      large: true,
                      onClick: () => e.onTool("viewSettings"),
                      title: text(
                        "Freeze rows and columns above and to the left of the selected cell",
                      ),
                      className: "xl-blue-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(PanelTop, {}),
                      label: text("Freeze Top Row"),
                      large: true,
                      active: e.freezeRows,
                      onClick: e.onFreeze,
                      className: "xl-blue-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Hide & Unhide"),
                      large: true,
                      onClick: () => e.onTool("viewSettings"),
                      title: text("Hide or restore selected rows and columns"),
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Task Panes"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Sparkles, {}),
                      label: text("AI Assistant"),
                      large: true,
                      onClick: () => e.onPanel("assistant"),
                      className: "xl-purple-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(History, {}),
                      label: text("Show Changes"),
                      large: true,
                      onClick: () => e.onPanel("history"),
                    }),
                  ],
                }),
              ],
            }),
        ],
      }),
    N = Object.fromEntries(
      Uh.map((J) => [J, collectRibbonCommands(T(J), J, C, _)]),
    ),
    R = collectRibbonCommands(
      i.jsxs("div", {
        className: "xl-file-menu",
        role: "menu",
        "aria-label": text("File"),
        lang: language("File"),
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
                  i.jsx(LocalizedText, { children: "New" }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
                      children: "Blank workbook",
                    }),
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
                  i.jsx(LocalizedText, { children: "Open" }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
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
                  i.jsx(LocalizedText, { children: "Manage worksheets" }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
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
                  i.jsx(LocalizedText, { children: "Workbook properties" }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
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
                  i.jsx(LocalizedText, { children: "Version history" }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
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
                  i.jsx(LocalizedText, {
                    children: "New from current template",
                  }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
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
                  i.jsx(LocalizedText, { children: "Save as Excel" }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
                      children: "Excel Workbook (.xlsx)",
                    }),
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
                  i.jsx(LocalizedText, { children: "Save as CSV" }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
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
                  i.jsx(LocalizedText, { children: "Save workbook data" }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
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
                  i.jsx(LocalizedText, {
                    children: "Original XLSX and compatibility",
                  }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
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
                  i.jsx(LocalizedText, { children: "Import sources" }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
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
                  i.jsx(LocalizedText, { children: "Workbook formats" }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
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
                children: i.jsx(LocalizedText, { children: "Rename workbook" }),
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
                children: i.jsx(LocalizedText, {
                  children: "Open audit example",
                }),
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
                children: i.jsx(LocalizedText, {
                  children: "Automation & Scripts",
                }),
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
                children: i.jsx(LocalizedText, { children: "Workbook help" }),
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
        reason: k || !e.canUndo ? _("Undo") : undefined,
      },
      {
        id: "quick-redo",
        name: "Redo",
        location: "Quick access",
        reason: k || !e.canRedo ? _("Redo") : undefined,
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
      const J = d.current;
      if (!J) return;
      const q = Array.from(
        m.current?.querySelectorAll("[data-workbook-command]") ?? [],
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
          i.jsx(CommandReferenceDialog, {
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
          "aria-label": text("Workbook ribbon"),
          onKeyDown: (J) => {
            J.key === "Escape" &&
              f &&
              !J.defaultPrevented &&
              (J.preventDefault(), J.stopPropagation(), w(), p(false));
          },
          lang: language("Workbook ribbon"),
          "data-ui-label": "Workbook ribbon",
          children: [
            i.jsxs("div", {
              className: "xl-ribbon-tabbar",
              children: [
                i.jsxs("div", {
                  className: "xl-file-control",
                  ref: g,
                  onKeyDown: (J) => {
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
                            g.current?.querySelectorAll(
                              '[role="menuitem"]:not(:disabled)',
                            ) ?? [],
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
                        ae[de]?.focus();
                      } else
                        J.key === "Escape" &&
                          (J.preventDefault(),
                          J.stopPropagation(),
                          l(false),
                          g.current?.querySelector("button")?.focus());
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
                      children: i.jsx(LocalizedText, { children: "File" }),
                    }),
                    u && R.content,
                  ],
                }),
                i.jsx("div", {
                  className: "xl-ribbon-tabs",
                  role: "tablist",
                  "aria-label": text("Ribbon tabs"),
                  lang: language("Ribbon tabs"),
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
                        title: text(
                          "Double-click to collapse or expand the ribbon",
                        ),
                        onClick: () => A(J),
                        onDoubleClick: E,
                        onKeyDown: (W) => {
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
                                document.getElementById(`${b}-${ae}`)?.focus());
                            } else if (W.key === "Home" || W.key === "End") {
                              W.preventDefault();
                              const ae = W.key === "Home" ? 0 : Uh.length - 1;
                              (A(Uh[ae]),
                                document.getElementById(`${b}-${ae}`)?.focus());
                            } else if (W.key === "ArrowDown") {
                              (W.preventDefault(), A(J));
                              const ae = N[J].commands.find((oe) => !oe.reason);
                              ae && (d.current = ae.id);
                            }
                          }
                        },
                        lang: language(
                          "Double-click to collapse or expand the ribbon",
                        ),
                        children: i.jsx(LocalizedText, { children: J }),
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
                      "aria-label": text(
                        c ? "Expand ribbon" : "Collapse ribbon",
                      ),
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
                    i.jsx(RibbonIconButton, {
                      icon: i.jsx(Undo2, { size: 16 }),
                      label: text("Undo (Ctrl+Z)"),
                      commandId: "quick-undo",
                      reason: L[0].reason,
                      onClick: e.onUndo,
                      disabled: k || !e.canUndo,
                    }),
                    i.jsx(RibbonIconButton, {
                      icon: i.jsx(Redo2, { size: 16 }),
                      label: text("Redo (Ctrl+Y)"),
                      commandId: "quick-redo",
                      reason: L[1].reason,
                      onClick: e.onRedo,
                      disabled: k || !e.canRedo,
                    }),
                    i.jsx(RibbonIconButton, {
                      icon: i.jsx(BookOpen, { size: 16 }),
                      label: text("Commands & shortcuts"),
                      onClick: () => s(true),
                    }),
                    i.jsx("span", { className: "xl-tabbar-divider" }),
                    i.jsxs("button", {
                      type: "button",
                      className: "xl-ai-button",
                      "aria-label": text("AI assistant"),
                      onClick: () => e.onPanel("assistant"),
                      lang: language("AI assistant"),
                      "data-ui-label": "AI assistant",
                      children: [
                        i.jsx(Sparkles, { size: 15 }),
                        i.jsx("span", {
                          children: i.jsx(LocalizedText, {
                            children: "AI Assistant",
                          }),
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

export { WorkbookRibbon };
