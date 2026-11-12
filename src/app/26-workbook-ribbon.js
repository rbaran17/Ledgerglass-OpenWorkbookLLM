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

function WorkbookRibbon(props) {
  const { text, language } = useWorkbookMessages(),
    [activeTab, setActiveTab] = ne.useState("Home"),
    [isCommandDialogOpen, setCommandDialogOpen] = ne.useState(false),
    pendingFocusId = ne.useRef(null),
    [isFileMenuOpen, setFileMenuOpen] = ne.useState(false),
    [isRibbonCollapsed, setRibbonCollapsed] = ne.useState(false),
    [isRibbonPeekOpen, setRibbonPeekOpen] = ne.useState(false),
    ribbonRef = ne.useRef(null),
    panelRef = ne.useRef(null),
    fileControlRef = ne.useRef(null),
    ribbonId = ne.useId(),
    isRibbonBodyVisible = !isRibbonCollapsed || isRibbonPeekOpen,
    focusActiveTab = () =>
      ribbonRef.current
        ?.querySelector('[role="tab"][aria-selected="true"]')
        ?.focus(),
    toggleRibbonCollapsed = () => {
      (panelRef.current?.contains(document.activeElement) && focusActiveTab(),
        setRibbonCollapsed((previousCollapsed) => !previousCollapsed),
        setRibbonPeekOpen(false),
        setFileMenuOpen(false));
    },
    selectTab = (tabName) => {
      (setActiveTab(tabName),
        setFileMenuOpen(false),
        isRibbonCollapsed && setRibbonPeekOpen(true));
    };
  (ne.useEffect(() => {
    if (!isRibbonPeekOpen) return;
    const handleOutsidePointer = (event) => {
      ribbonRef.current?.contains(event.target) || setRibbonPeekOpen(false);
    };
    return (
      document.addEventListener("pointerdown", handleOutsidePointer),
      document.addEventListener("focusin", handleOutsidePointer),
      () => {
        (document.removeEventListener("pointerdown", handleOutsidePointer),
          document.removeEventListener("focusin", handleOutsidePointer));
      }
    );
  }, [isRibbonPeekOpen]),
    ne.useEffect(() => {
      const handleRibbonShortcut = (keyboardEvent) => {
        if (
          keyboardEvent.defaultPrevented ||
          keyboardEvent.repeat ||
          keyboardEvent.altKey ||
          keyboardEvent.shiftKey ||
          keyboardEvent.key !== "F1" ||
          !(keyboardEvent.ctrlKey || keyboardEvent.metaKey)
        )
          return;
        const target = keyboardEvent.target;
        if (!(target instanceof Element) || target.closest('[role="dialog"]'))
          return;
        const ribbonElement = ribbonRef.current,
          shortcutScope = ribbonElement?.closest(".lg-app") ?? ribbonElement;
        !shortcutScope ||
          (shortcutScope.classList.contains("lg-app")
            ? target.closest(".lg-app") !== shortcutScope
            : !shortcutScope.contains(target)) ||
          (keyboardEvent.preventDefault(),
          keyboardEvent.stopPropagation(),
          toggleRibbonCollapsed());
      };
      return (
        document.addEventListener("keydown", handleRibbonShortcut),
        () => document.removeEventListener("keydown", handleRibbonShortcut)
      );
    }, []));
  const isDisabled = props.readOnly || props.busy,
    effectiveCellStyle = resolveEffectiveCellStyle(props.cell?.style),
    applyStyle = (style) => props.onFormat({ style: style }),
    applyFormat = (format) =>
      props.onFormat({
        format: format,
        numFmt: undefined,
        numberFormat: undefined,
      }),
    hasAdvancedNumberFormat =
      props.cell?.numberFormat &&
      !["number", "currency", "accounting", "percent"].includes(
        props.cell.numberFormat.category,
      ),
    adjustDecimalPlaces = (decimalDelta) => {
      if (props.cell?.numberFormat && !hasAdvancedNumberFormat) {
        const numberFormat = validateNumberFormatSettings({
          ...props.cell.numberFormat,
          decimals: Math.max(
            0,
            Math.min(
              12,
              (props.cell.numberFormat.decimals ?? 2) + decimalDelta,
            ),
          ),
        });
        props.onFormat({
          numberFormat: numberFormat,
          numFmt: buildNumberFormatCode(numberFormat),
        });
        return;
      }
      const currentDecimals =
          summarizeNumberFormat(props.cell?.numFmt) ??
          (props.cell?.format === "currency" ||
          props.cell?.format === "number" ||
          props.cell?.format === "percent"
            ? 2
            : 0),
        nextDecimals = Math.max(
          0,
          Math.min(10, currentDecimals + decimalDelta),
        );
      props.onFormat({
        numberFormat: undefined,
        numFmt: `${props.cell?.format === "currency" ? "$" : ""}#,##0${nextDecimals ? `.${"0".repeat(nextDecimals)}` : ""}${props.cell?.format === "percent" ? "%" : ""}`,
      });
    };
  ne.useEffect(() => {
    if (!isFileMenuOpen) return;
    const handleFileMenuPointer = (event) => {
      fileControlRef.current?.contains(event.target) || setFileMenuOpen(false);
    };
    return (
      document.addEventListener("pointerdown", handleFileMenuPointer),
      () => document.removeEventListener("pointerdown", handleFileMenuPointer)
    );
  }, [isFileMenuOpen]);
  const runFileMenuAction = (action) => {
      (setFileMenuOpen(false), action());
    },
    numberFormatSelect = i.jsxs("select", {
      "aria-label": text("Number format"),
      value: props.cell?.numberFormat
        ? "advanced"
        : (props.cell?.format ?? "general"),
      disabled: isDisabled,
      onChange: (event) =>
        event.target.value === "advanced"
          ? props.onTool("numberFormat")
          : applyFormat(event.target.value),
      lang: language("Number format"),
      "data-ui-label": "Number format",
      children: [
        props.cell?.numberFormat &&
          i.jsxs("option", {
            value: "advanced",
            children: [
              props.cell.numberFormat.category.charAt(0).toUpperCase() +
                props.cell.numberFormat.category.slice(1),
              " ",
              "(",
              props.cell.numberFormat.locale,
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
    zoomSelect = i.jsx("select", {
      "aria-label": text("Zoom"),
      value: props.zoom,
      onChange: (event) => props.onZoom(Number(event.target.value)),
      lang: language("Zoom"),
      "data-ui-label": "Zoom",
      children: Array.from(
        new Set([50, 75, 90, 100, 110, 125, 150, 175, 200, props.zoom]),
      )
        .sort((leftZoom, rightZoom) => leftZoom - rightZoom)
        .map((zoomOption) =>
          i.jsxs(
            "option",
            { value: zoomOption, children: [zoomOption, "%"] },
            zoomOption,
          ),
        ),
    }),
    getDisabledReason = (commandLabel) =>
      props.busy
        ? "Wait for the current workbook operation to finish."
        : !props.allowProtectionChange &&
            /Protect|New Sheet|New from/.test(commandLabel)
          ? "The host has opened this workbook read-only."
          : props.readOnly
            ? props.sheetProtected
              ? "Unprotect the worksheet before editing."
              : "The host has opened this workbook read-only."
            : commandLabel === "Undo"
              ? "There is no earlier edit to undo."
              : commandLabel === "Redo"
                ? "There is no undone edit to redo."
                : /Decimal/.test(commandLabel)
                  ? "Choose a number, currency, accounting or percent format to adjust decimal places."
                  : /New Sheet/.test(commandLabel) &&
                      props.workbookStructureProtected
                    ? "Unprotect workbook structure in Manage worksheets before adding a sheet."
                    : "This command is unavailable for the current workbook state.",
    ribbonComponents = {
      group: RibbonGroup,
      command: RibbonCommandButton,
      icon: RibbonIconButton,
      color: ColorPickerButton,
    },
    renderRibbonTab = (tab) =>
      i.jsxs(i.Fragment, {
        children: [
          tab === "Home" &&
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
                          onClick: props.onPaste,
                          disabled: isDisabled,
                          className: "xl-gold-icon xl-paste-primary",
                        }),
                        i.jsxs("button", {
                          type: "button",
                          className: "xl-paste-special",
                          "aria-label": text("Paste Special"),
                          title: text(
                            "Choose values, formulas, formatting and other paste options",
                          ),
                          disabled: isDisabled,
                          onClick: () => props.onTool("pasteSpecial"),
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
                          onClick: props.onCut,
                          disabled: isDisabled,
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Copy, {}),
                          label: text("Copy"),
                          onClick: props.onCopy,
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Paintbrush, {}),
                          label: text("Format Painter"),
                          onClick: () => props.onTool("paint"),
                          disabled: isDisabled,
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Font"),
                  className: "xl-font-group",
                  onDialog: () => props.onTool("cellStyle"),
                  dialogLabel: text("Cell Styles"),
                  dialogDisabled: isDisabled,
                  children: [
                    i.jsxs("div", {
                      className: "xl-font-selects",
                      children: [
                        i.jsx("select", {
                          "aria-label": text("Font family"),
                          value: effectiveCellStyle.fontFamily ?? "Calibri",
                          disabled: isDisabled,
                          onChange: (event) =>
                            applyStyle({ fontFamily: event.target.value }),
                          lang: language("Font family"),
                          "data-ui-label": "Font family",
                          children: [
                            "Calibri",
                            "Aptos",
                            "Arial",
                            "Times New Roman",
                            "Courier New",
                          ].map((family) =>
                            i.jsx("option", { children: family }, family),
                          ),
                        }),
                        i.jsx("select", {
                          "aria-label": text("Font size"),
                          value: effectiveCellStyle.fontSize ?? 11,
                          disabled: isDisabled,
                          onChange: (event) =>
                            applyStyle({
                              fontSize: Number(event.target.value),
                            }),
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
                              effectiveCellStyle.fontSize ?? 11,
                            ]),
                          )
                            .sort((leftSize, rightSize) => leftSize - rightSize)
                            .map((size) =>
                              i.jsx("option", { children: size }, size),
                            ),
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
                          disabled: isDisabled,
                          onClick: () =>
                            applyStyle({
                              fontSize: Math.min(
                                72,
                                (effectiveCellStyle.fontSize ?? 11) + 1,
                              ),
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
                          active: !!effectiveCellStyle.bold,
                          onClick: () =>
                            applyStyle({ bold: !effectiveCellStyle.bold }),
                          disabled: isDisabled,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(Italic, { size: 15 }),
                          label: text("Italic (Ctrl+I)"),
                          active: !!effectiveCellStyle.italic,
                          onClick: () =>
                            applyStyle({ italic: !effectiveCellStyle.italic }),
                          disabled: isDisabled,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(Underline, { size: 15 }),
                          label: text("Underline (Ctrl+U)"),
                          active: !!effectiveCellStyle.underline,
                          onClick: () =>
                            applyStyle({
                              underline: !effectiveCellStyle.underline,
                            }),
                          disabled: isDisabled,
                        }),
                        i.jsx("span", { className: "xl-control-divider" }),
                        i.jsxs("select", {
                          className: "xl-border-select",
                          "aria-label": text("Borders"),
                          title: text("Borders"),
                          value: effectiveCellStyle.border ?? "none",
                          disabled: isDisabled,
                          onChange: (event) =>
                            applyStyle({
                              border: event.target.value,
                              borders: undefined,
                            }),
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
                          value: effectiveCellStyle.fill ?? "#FFFF00",
                          onChange: (fill) => applyStyle({ fill: fill }),
                          disabled: isDisabled,
                        }),
                        i.jsx(ColorPickerButton, {
                          kind: "color",
                          value: effectiveCellStyle.color ?? "#C00000",
                          onChange: (color) => applyStyle({ color: color }),
                          disabled: isDisabled,
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
                          active: effectiveCellStyle.align === "left",
                          onClick: () => applyStyle({ align: "left" }),
                          disabled: isDisabled,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(AlignCenter, { size: 16 }),
                          label: text("Center"),
                          active: effectiveCellStyle.align === "center",
                          onClick: () => applyStyle({ align: "center" }),
                          disabled: isDisabled,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(AlignRight, { size: 16 }),
                          label: text("Align Right"),
                          active: effectiveCellStyle.align === "right",
                          onClick: () => applyStyle({ align: "right" }),
                          disabled: isDisabled,
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(WrapText, {}),
                          label: text("Wrap Text"),
                          active: !!effectiveCellStyle.wrap,
                          onClick: () =>
                            applyStyle({ wrap: !effectiveCellStyle.wrap }),
                          disabled: isDisabled,
                        }),
                      ],
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid2x2, {}),
                      label: text("Merge & Center"),
                      onClick: () => props.onTool("merge"),
                      disabled: isDisabled,
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Number"),
                  className: "xl-number-group",
                  children: [
                    numberFormatSelect,
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(DollarSign, { size: 15 }),
                      label: text("Format Cells"),
                      onClick: () => props.onTool("numberFormat"),
                      disabled: isDisabled,
                    }),
                    i.jsxs("div", {
                      className: "xl-number-buttons",
                      children: [
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(DollarSign, { size: 16 }),
                          label: text("Currency"),
                          onClick: () => applyFormat("currency"),
                          disabled: isDisabled,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx(Percent, { size: 15 }),
                          label: text("Percent Style"),
                          onClick: () => applyFormat("percent"),
                          disabled: isDisabled,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsx("span", {
                            className: "xl-comma-icon",
                            children: ",",
                          }),
                          label: text("Comma Style"),
                          onClick: () =>
                            props.onFormat({
                              format: "number",
                              numFmt: "#,##0.00",
                              numberFormat: undefined,
                            }),
                          disabled: isDisabled,
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
                          onClick: () => adjustDecimalPlaces(1),
                          disabled: isDisabled || !!hasAdvancedNumberFormat,
                        }),
                        i.jsx(RibbonIconButton, {
                          icon: i.jsxs("span", {
                            className: "xl-decimal-icon",
                            children: ["→", i.jsx("small", { children: ".0" })],
                          }),
                          label: text("Decrease Decimal"),
                          onClick: () => adjustDecimalPlaces(-1),
                          disabled: isDisabled || !!hasAdvancedNumberFormat,
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
                      onClick: () => props.onTool("conditional"),
                      disabled: isDisabled,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Format as Table"),
                      large: true,
                      onClick: () => props.onTool("table"),
                      disabled: isDisabled,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(LayoutGrid, {}),
                      label: text("Heading Style"),
                      large: true,
                      onClick: () =>
                        applyStyle({
                          fill: "#E2EFDA",
                          color: "#375623",
                          bold: true,
                        }),
                      disabled: isDisabled,
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
                        onClick: props.onAddRow,
                        disabled: isDisabled,
                        className: "xl-green-icon",
                      }),
                      i.jsx(RibbonCommandButton, {
                        icon: i.jsx(Table2, {}),
                        label: text("Insert Column"),
                        onClick: props.onAddColumn,
                        disabled: isDisabled,
                        className: "xl-green-icon",
                      }),
                      i.jsx(RibbonCommandButton, {
                        icon: i.jsx(Sheet, {}),
                        label: text("New Sheet"),
                        onClick: props.onAddSheet,
                        disabled:
                          props.busy ||
                          !props.allowProtectionChange ||
                          props.workbookStructureProtected,
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
                        onClick: props.onAutoFitColumns,
                        disabled: isDisabled,
                        title: text(
                          "Fit the selected columns to their contents",
                        ),
                      }),
                      i.jsx(RibbonCommandButton, {
                        icon: i.jsx(Rows3, {}),
                        label: text("AutoFit Row Height"),
                        onClick: props.onAutoFitRows,
                        disabled: isDisabled,
                        title: text(
                          "Fit the selected rows to their contents and wrapping",
                        ),
                      }),
                      i.jsx(RibbonCommandButton, {
                        icon: i.jsx(Rows3, {}),
                        label: text("Hide & Unhide"),
                        onClick: () => props.onTool("viewSettings"),
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
                          onClick: props.onAutosum,
                          disabled: isDisabled,
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(ArrowDownToLine, {}),
                          label: text("Fill Down"),
                          onClick: props.onFillDown,
                          disabled: isDisabled,
                          className: "xl-blue-icon",
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Eraser, {}),
                          label: text("Clear"),
                          onClick: props.onClear,
                          disabled: isDisabled,
                        }),
                      ],
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Fill Series"),
                      large: true,
                      onClick: () => props.onTool("fillSeries"),
                      disabled: isDisabled,
                      title: text(
                        "Fill in any direction or create a number or date series",
                      ),
                      className: "xl-blue-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ListFilter, {}),
                      label: text("Sort & Filter"),
                      large: true,
                      onClick: props.onToggleFilter,
                      className: "xl-blue-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Search, {}),
                      label: text("Find & Select"),
                      large: true,
                      onClick: props.onFind,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Search, {}),
                      label: text("Advanced Find"),
                      onClick: () => props.onTool("search"),
                      title: text(
                        "Find all matches, preview replacements, or inspect special cells (Ctrl/⌘+Shift+F)",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Search, {}),
                      label: text("Go To Special"),
                      onClick: () => props.onTool("goToSpecial"),
                      title: text(
                        "Select formulas, errors, constants, notes or blanks (F5)",
                      ),
                    }),
                  ],
                }),
              ],
            }),
          tab === "Insert" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs(RibbonGroup, {
                  label: text("Tables"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("PivotTable"),
                      large: true,
                      onClick: () => props.onTool("pivot"),
                      disabled: isDisabled,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Table"),
                      large: true,
                      onClick: () => props.onTool("table"),
                      disabled: isDisabled,
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
                      onClick: () => props.onTool("picture"),
                      disabled: isDisabled,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Square, {}),
                      label: text("Shapes"),
                      large: true,
                      onClick: () => props.onTool("shape"),
                      disabled: isDisabled,
                    }),
                  ],
                }),
                i.jsx(RibbonGroup, {
                  label: text("Charts"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(ChartColumn, {}),
                    label: text("Recommended Charts"),
                    large: true,
                    onClick: () => props.onTool("chart"),
                    disabled: isDisabled,
                  }),
                }),
                i.jsx(RibbonGroup, {
                  label: text("Comments"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(MessageSquare, {}),
                    label: text("New Note"),
                    large: true,
                    onClick: props.onNote,
                    disabled: isDisabled,
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
                      onClick: props.onAddSheet,
                      disabled:
                        props.busy ||
                        !props.allowProtectionChange ||
                        props.workbookStructureProtected,
                      className: "xl-green-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Insert Row"),
                      large: true,
                      onClick: props.onAddRow,
                      disabled: isDisabled,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Insert Column"),
                      large: true,
                      onClick: props.onAddColumn,
                      disabled: isDisabled,
                    }),
                  ],
                }),
                i.jsx(RibbonGroup, {
                  label: text("Files"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(FolderOpen, {}),
                    label: text("Open Workbook"),
                    large: true,
                    onClick: props.onImport,
                    disabled: isDisabled,
                    className: "xl-gold-icon",
                  }),
                }),
              ],
            }),
          tab === "Page Layout" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsx(RibbonGroup, {
                  label: text("Themes"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(LayoutGrid, {}),
                    label: text("Themes"),
                    large: true,
                    onClick: () => props.onTool("themes"),
                    disabled: isDisabled,
                  }),
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Page Setup"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Sheet, {}),
                      label: text("Margins"),
                      large: true,
                      onClick: () => props.onTool("margins"),
                      disabled: false,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Orientation"),
                      large: true,
                      onClick: () => props.onTool("orientation"),
                      disabled: false,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Square, {}),
                      label: text("Size"),
                      large: true,
                      onClick: () => props.onTool("size"),
                      disabled: false,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid2x2, {}),
                      label: text("Print Area"),
                      large: true,
                      onClick: () => props.onTool("printArea"),
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
                      onClick: () => props.onTool("printTitles"),
                      disabled: false,
                      large: true,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid2x2, {}),
                      label: text("Page Breaks"),
                      onClick: () => props.onTool("pageBreaks"),
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
                    onClick: () => props.onTool("fit"),
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
                          checked: props.gridlines,
                          onChange: props.onGridlines,
                        }),
                        i.jsx(LocalizedText, { children: "Gridlines" }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: props.headings,
                          onChange: props.onHeadings,
                        }),
                        i.jsx(LocalizedText, { children: "Headings" }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          tab === "Formulas" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs(RibbonGroup, {
                  label: text("Function Library"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Sigma, {}),
                      label: text("AutoSum"),
                      large: true,
                      onClick: props.onAutosum,
                      disabled: isDisabled,
                      className: "xl-blue-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(BookOpen, {}),
                      label: text("Function Reference"),
                      large: true,
                      onClick: () => props.onPanel("help"),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Calculator, {}),
                      label: text("Insert Function"),
                      large: true,
                      onClick: () => props.onTool("function"),
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
                    onClick: () => props.onTool("extensions"),
                    disabled: props.busy,
                  }),
                }),
                i.jsx(RibbonGroup, {
                  label: text("Defined Names"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(Table2, {}),
                    label: text("Name Manager"),
                    large: true,
                    onClick: () => props.onTool("names"),
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
                          onClick: () => props.onTool("tracePrecedents"),
                          title: text(
                            "Inspect the cells referenced by the active formula",
                          ),
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(ArrowLeftRight, {}),
                          label: text("Trace Dependents"),
                          onClick: () => props.onTool("traceDependents"),
                          title: text(
                            "Inspect formulas that reference the active cell",
                          ),
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Sheet, {}),
                          label: text("Watch Window"),
                          onClick: () => props.onTool("formulaWatches"),
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
                      onClick: () => props.onTool("evaluateFormula"),
                      title: text(
                        "Replay actual calculation steps and inspect errors",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ShieldCheck, {}),
                      label: text("Compatibility"),
                      large: true,
                      onClick: () => props.onTool("functionCompatibility"),
                      title: text(
                        "Review function contracts and unsupported formulas",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Sheet, {}),
                      label: text("Show Formulas"),
                      large: true,
                      active: props.showFormulas,
                      onClick: props.onShowFormulas,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ShieldCheck, {}),
                      label: text("Audit Checks"),
                      large: true,
                      onClick: () => props.onPanel("checks"),
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
                      onClick: () => props.onTool("calculation"),
                      title: text(
                        "Automatic or manual mode, iteration and calculation evidence",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Calculator, {}),
                      label: text("Calculate Now"),
                      large: true,
                      onClick: () => props.onTool("calculateNow"),
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
                    onClick: () => props.onPanel("assistant"),
                    className: "xl-purple-icon",
                  }),
                }),
              ],
            }),
          tab === "Data" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsxs(RibbonGroup, {
                  label: text("Get & Transform Data"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(FolderOpen, {}),
                      label: text("Get Data"),
                      large: true,
                      onClick: props.onImport,
                      disabled: isDisabled,
                      className: "xl-gold-icon",
                      title: text("Import Excel, CSV, TSV, or workbook JSON"),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("From Table / Range"),
                      large: true,
                      onClick: () => props.onTool("query"),
                      disabled: isDisabled,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ArrowLeftRight, {}),
                      label: text("Append / Join"),
                      large: true,
                      onClick: () => props.onTool("combinePopulations"),
                      disabled: props.busy,
                      title: text(
                        "Review combined populations, unmatched records and match counts",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Transformation Recipes"),
                      large: true,
                      onClick: () => props.onTool("transformationRecipes"),
                      disabled: props.busy,
                      title: text(
                        "Build versioned steps with source fingerprints, failure details and reviewed outputs",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(FolderOpen, {}),
                      label: text("Connections & Refresh"),
                      large: true,
                      onClick: () => props.onTool("externalSources"),
                      disabled: props.busy,
                      title: text(
                        "Review approved external sources, control changes and last successful data",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Data Model"),
                      large: true,
                      onClick: () => props.onTool("dataModel"),
                      disabled: props.busy,
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
                      onClick: () => props.onTool("populationFilter"),
                      disabled: props.busy,
                      title: text(
                        "Filter values, text, numbers, dates and advanced criteria while retaining source rows",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ArrowDownAZ, {}),
                      label: text("Custom Sort"),
                      large: true,
                      onClick: () => props.onTool("sort"),
                      disabled: props.busy,
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
                          onClick: () => props.onSort("asc"),
                          className: "xl-blue-icon",
                          title: text(
                            "Sort worksheet view ascending by the selected column",
                          ),
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(ArrowUpAZ, {}),
                          label: text("Sort Z to A"),
                          onClick: () => props.onSort("desc"),
                          className: "xl-blue-icon",
                          title: text(
                            "Sort worksheet view descending by the selected column",
                          ),
                        }),
                        i.jsx(RibbonCommandButton, {
                          icon: i.jsx(Eraser, {}),
                          label: text("Clear Sort"),
                          onClick: () => props.onSort(null),
                        }),
                      ],
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Filter, {}),
                      label: text("Filter"),
                      large: true,
                      onClick: props.onToggleFilter,
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
                      onClick: () => props.onPanel("checks"),
                      className: "xl-green-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(CheckCheck, {}),
                      label: text("Data Validation"),
                      large: true,
                      onClick: () => props.onTool("validation"),
                      disabled: props.busy,
                      title: text(
                        "Review validation rules and inspect existing values; editing requires write access",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Copy, {}),
                      label: text("Remove Duplicates"),
                      large: true,
                      onClick: () => props.onTool("duplicates"),
                      disabled: props.busy,
                      title: text(
                        "Review duplicate groups and create a separate deduplicated population",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Columns3, {}),
                      label: text("Text to Columns"),
                      large: true,
                      onClick: () => props.onTool("textToColumns"),
                      disabled: props.busy,
                      title: text(
                        "Preview split fields and types before creating a separate output",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Columns3, {}),
                      label: text("Clean up text"),
                      large: true,
                      onClick: () => props.onTool("textCleanup"),
                      disabled: props.busy,
                      title: text(
                        "Review cleanup and extraction steps while retaining original evidence",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Columns3, {}),
                      label: text("Clean & Reshape"),
                      large: true,
                      onClick: () => props.onTool("fieldTransform"),
                      disabled: props.busy,
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
                    onClick: () => props.onTool("workbookLinks"),
                    disabled: props.busy,
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
                      onClick: () => props.onTool("scenarios"),
                      disabled: props.busy,
                      title: text(
                        "Compare named assumptions and review a scenario before applying it",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Data Table"),
                      large: true,
                      onClick: () => props.onTool("dataTable"),
                      disabled: props.busy,
                      title: text(
                        "Build a one- or two-variable sensitivity table with retained baseline assumptions",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Calculator, {}),
                      label: text("Goal Seek"),
                      large: true,
                      onClick: () => props.onTool("goalSeek"),
                      disabled: props.busy,
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
                      onClick: () => props.onTool("solver"),
                      disabled: props.busy,
                      title: text(
                        "Review bounded linear optimization with explicit constraints",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Table2, {}),
                      label: text("Data Analysis"),
                      large: true,
                      onClick: () => props.onTool("statistics"),
                      disabled: props.busy,
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
                      onClick: () => props.onTool("consolidation"),
                      disabled: props.busy,
                      title: text(
                        "Review cross-sheet or approved-workbook rollups, source context and grouped subtotals",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Group & Outline"),
                      large: true,
                      onClick: () => props.onTool("outlines"),
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
                      onClick: () => props.onTool("reportOutputs"),
                      disabled: props.busy,
                      title: text(
                        "Publish fixed snapshots, review linked refreshes and inspect embedded objects",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(FileSpreadsheet, {}),
                      label: text("Save as Excel"),
                      large: true,
                      onClick: () => props.onExport("xlsx"),
                      disabled: props.busy,
                      className: "xl-green-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Download, {}),
                      label: text("Export CSV"),
                      large: true,
                      onClick: () => props.onExport("csv"),
                      disabled: props.busy,
                    }),
                  ],
                }),
              ],
            }),
          tab === "Review" &&
            i.jsxs(i.Fragment, {
              children: [
                i.jsx(RibbonGroup, {
                  label: text("Proofing"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(BookOpen, {}),
                    label: text("Spelling"),
                    large: true,
                    onClick: () => props.onTool("spelling"),
                    disabled: false,
                  }),
                }),
                i.jsx(RibbonGroup, {
                  label: text("Notes"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(MessageSquare, {}),
                    label: text("New / Edit Note"),
                    large: true,
                    onClick: props.onNote,
                    disabled: isDisabled,
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
                      onClick: () => props.onTool("comments"),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(MessageSquare, {}),
                      label: text("Hyperlink"),
                      onClick: () => props.onTool("hyperlink"),
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
                      onClick: () => props.onTool("patterns"),
                      disabled: props.busy,
                      title: text(
                        "Test explicit duplicate keys, sequence gaps, round amounts, weekends, holidays and combined rules",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ListFilter, {}),
                      label: text("Aging Schedule"),
                      large: true,
                      onClick: () => props.onTool("aging"),
                      disabled: props.busy,
                      title: text(
                        "Reconcile configurable aging buckets using a fixed as-of date",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ListFilter, {}),
                      label: text("Risk Screening"),
                      large: true,
                      onClick: () => props.onTool("materiality"),
                      disabled: props.busy,
                      title: text(
                        "Review explicit thresholds, currency and effective period",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(CheckCheck, {}),
                      label: text("Mark Reviewed"),
                      large: true,
                      onClick: props.onReviewed,
                      disabled: isDisabled,
                      className: "xl-green-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(ClipboardCheck, {}),
                      label: text("Audit Checks"),
                      large: true,
                      onClick: () => props.onPanel("checks"),
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
                      onClick: () => props.onTool("workbookComparison"),
                      disabled: props.busy,
                      title: text(
                        "Compare keyed rows, formulas, links and hidden content",
                      ),
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(History, {}),
                      label: text("Show Changes"),
                      large: true,
                      onClick: () => props.onPanel("history"),
                      className: "xl-blue-icon",
                    }),
                  ],
                }),
                i.jsx(RibbonGroup, {
                  label: text("Protection"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(ShieldCheck, {}),
                    label: props.sheetProtected
                      ? "Unprotect Sheet"
                      : "Protect Sheet",
                    large: true,
                    onClick: () => props.onTool("protection"),
                    disabled: !props.allowProtectionChange,
                  }),
                }),
                i.jsx(RibbonGroup, {
                  label: text("AI Review"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(Sparkles, {}),
                    label: text("AI Assistant"),
                    large: true,
                    onClick: () => props.onPanel("assistant"),
                    className: "xl-purple-icon",
                  }),
                }),
                i.jsx(RibbonGroup, {
                  label: text("Automation"),
                  children: i.jsx(RibbonCommandButton, {
                    icon: i.jsx(FileCode, {}),
                    label: text("Automation & Scripts"),
                    large: true,
                    onClick: () => props.onTool("automation"),
                    disabled: props.busy,
                    title: text(
                      "Inspect and preserve VBA/Office Scripts; review bounded migrations",
                    ),
                  }),
                }),
              ],
            }),
          tab === "View" &&
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
                        (props.gridlines || props.onGridlines(),
                          props.headings || props.onHeadings());
                      },
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Sheet, {}),
                      label: text("Page Layout"),
                      large: true,
                      onClick: () => props.onTool("pageLayout"),
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
                          checked: props.gridlines,
                          onChange: props.onGridlines,
                        }),
                        i.jsx(LocalizedText, { children: "Gridlines" }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: props.headings,
                          onChange: props.onHeadings,
                        }),
                        i.jsx(LocalizedText, { children: "Headings" }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: props.showFormulas,
                          onChange: props.onShowFormulas,
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
                      children: [i.jsx(Search, { size: 22 }), zoomSelect],
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Maximize, {}),
                      label: "100%",
                      large: true,
                      onClick: () => props.onZoom(100),
                    }),
                  ],
                }),
                i.jsxs(RibbonGroup, {
                  label: text("Window"),
                  children: [
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid2x2, {}),
                      label: text("Split"),
                      onClick: () => props.onTool("viewSettings"),
                      title: text(
                        "Split the worksheet into two or four scrolling panes",
                      ),
                      disabled: isDisabled,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Default Sizes"),
                      onClick: () => props.onTool("viewSettings"),
                      title: text(
                        "Set worksheet-wide row height and column width",
                      ),
                      disabled: isDisabled,
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Grid2x2, {}),
                      label: text("Freeze Panes"),
                      large: true,
                      onClick: () => props.onTool("viewSettings"),
                      title: text(
                        "Freeze rows and columns above and to the left of the selected cell",
                      ),
                      className: "xl-blue-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(PanelTop, {}),
                      label: text("Freeze Top Row"),
                      large: true,
                      active: props.freezeRows,
                      onClick: props.onFreeze,
                      className: "xl-blue-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(Rows3, {}),
                      label: text("Hide & Unhide"),
                      large: true,
                      onClick: () => props.onTool("viewSettings"),
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
                      onClick: () => props.onPanel("assistant"),
                      className: "xl-purple-icon",
                    }),
                    i.jsx(RibbonCommandButton, {
                      icon: i.jsx(History, {}),
                      label: text("Show Changes"),
                      large: true,
                      onClick: () => props.onPanel("history"),
                    }),
                  ],
                }),
              ],
            }),
        ],
      }),
    commandsByTab = Object.fromEntries(
      Uh.map((tabLabel) => [
        tabLabel,
        collectRibbonCommands(
          renderRibbonTab(tabLabel),
          tabLabel,
          ribbonComponents,
          getDisabledReason,
        ),
      ]),
    ),
    ribbonCommands = collectRibbonCommands(
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
            onClick: () => runFileMenuAction(props.onNew),
            disabled: isDisabled,
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
            onClick: () => runFileMenuAction(props.onImport),
            disabled: isDisabled,
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
            onClick: () => runFileMenuAction(() => props.onTool("sheets")),
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
            onClick: () =>
              runFileMenuAction(() => props.onTool("workbookProperties")),
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
            onClick: () => runFileMenuAction(() => props.onTool("versions")),
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
            disabled: !props.allowProtectionChange || props.busy,
            onClick: () => runFileMenuAction(() => props.onTool("template")),
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
            onClick: () => runFileMenuAction(() => props.onExport("xlsx")),
            disabled: props.busy,
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
            onClick: () => runFileMenuAction(() => props.onExport("csv")),
            disabled: props.busy,
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
            onClick: () => runFileMenuAction(() => props.onExport("json")),
            disabled: props.busy,
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
            disabled: props.busy,
            onClick: () => runFileMenuAction(() => props.onTool("xlsxSource")),
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
            disabled: props.busy,
            onClick: () =>
              runFileMenuAction(() => props.onTool("importSources")),
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
            disabled: props.busy,
            onClick: () =>
              runFileMenuAction(() => props.onTool("workbookFormats")),
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
            onClick: () => runFileMenuAction(props.onRename),
            disabled: isDisabled,
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
            onClick: () => runFileMenuAction(props.onDemo),
            disabled: isDisabled,
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
            disabled: props.busy,
            onClick: () => runFileMenuAction(() => props.onTool("automation")),
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
            onClick: () => runFileMenuAction(() => props.onPanel("help")),
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
      ribbonComponents,
      getDisabledReason,
    ),
    quickAccessCommands = [
      {
        id: "quick-undo",
        name: "Undo",
        location: "Quick access",
        reason:
          isDisabled || !props.canUndo ? getDisabledReason("Undo") : undefined,
      },
      {
        id: "quick-redo",
        name: "Redo",
        location: "Quick access",
        reason:
          isDisabled || !props.canRedo ? getDisabledReason("Redo") : undefined,
      },
    ],
    onLocate = (commandEntry) => {
      commandEntry.reason ||
        (setCommandDialogOpen(false),
        (pendingFocusId.current = commandEntry.id),
        commandEntry.id.startsWith("File:")
          ? (setFileMenuOpen(true), setRibbonPeekOpen(false))
          : commandEntry.id.startsWith("quick-") ||
            selectTab(commandEntry.id.split(":")[0]));
    };
  return (
    ne.useEffect(() => {
      const pendingCommandId = pendingFocusId.current;
      if (!pendingCommandId) return;
      const focusTarget = Array.from(
        ribbonRef.current?.querySelectorAll("[data-workbook-command]") ?? [],
      ).find((element) => element.dataset.workbookCommand === pendingCommandId);
      focusTarget && (focusTarget.focus(), (pendingFocusId.current = null));
    }),
    ne.useEffect(() => {
      const handleHelpShortcut = (helpKeyEvent) => {
        if (
          helpKeyEvent.defaultPrevented ||
          helpKeyEvent.repeat ||
          helpKeyEvent.key !== "F1" ||
          helpKeyEvent.ctrlKey ||
          helpKeyEvent.metaKey ||
          helpKeyEvent.altKey ||
          helpKeyEvent.shiftKey
        )
          return;
        const target = helpKeyEvent.target,
          ribbonNode = ribbonRef.current;
        if (
          !(target instanceof Element) ||
          !ribbonNode ||
          target.closest('[role="dialog"]')
        )
          return;
        const f1Scope = ribbonNode.closest(".lg-app") ?? ribbonNode;
        (f1Scope.classList.contains("lg-app")
          ? target.closest(".lg-app") !== f1Scope
          : !f1Scope.contains(target)) ||
          (helpKeyEvent.preventDefault(),
          helpKeyEvent.stopPropagation(),
          setCommandDialogOpen(true));
      };
      return (
        document.addEventListener("keydown", handleHelpShortcut),
        () => document.removeEventListener("keydown", handleHelpShortcut)
      );
    }, []),
    i.jsxs(i.Fragment, {
      children: [
        isCommandDialogOpen &&
          i.jsx(CommandReferenceDialog, {
            commands: [
              ...quickAccessCommands,
              ...ribbonCommands.commands,
              ...Uh.flatMap((ribbonTab) => commandsByTab[ribbonTab].commands),
            ],
            onClose: () => setCommandDialogOpen(false),
            onLocate: onLocate,
          }),
        i.jsxs("div", {
          ref: ribbonRef,
          role: "group",
          className: `xl-ribbon${isRibbonCollapsed ? " xl-ribbon-collapsed" : ""}`,
          "aria-label": text("Workbook ribbon"),
          onKeyDown: (ribbonKeyEvent) => {
            ribbonKeyEvent.key === "Escape" &&
              isRibbonPeekOpen &&
              !ribbonKeyEvent.defaultPrevented &&
              (ribbonKeyEvent.preventDefault(),
              ribbonKeyEvent.stopPropagation(),
              focusActiveTab(),
              setRibbonPeekOpen(false));
          },
          lang: language("Workbook ribbon"),
          "data-ui-label": "Workbook ribbon",
          children: [
            i.jsxs("div", {
              className: "xl-ribbon-tabbar",
              children: [
                i.jsxs("div", {
                  className: "xl-file-control",
                  ref: fileControlRef,
                  onKeyDown: (fileKeyEvent) => {
                    if (!(
                      fileKeyEvent.altKey ||
                      fileKeyEvent.ctrlKey ||
                      fileKeyEvent.metaKey ||
                      fileKeyEvent.defaultPrevented
                    ))
                      if (
                        ["ArrowDown", "ArrowUp", "Home", "End"].includes(
                          fileKeyEvent.key,
                        )
                      ) {
                        if ((fileKeyEvent.preventDefault(), !isFileMenuOpen)) {
                          setFileMenuOpen(true);
                          const command = ribbonCommands.commands.find(
                            (command) => !command.reason,
                          );
                          command && (pendingFocusId.current = command.id);
                          return;
                        }
                        const menuItems = Array.from(
                            fileControlRef.current?.querySelectorAll(
                              '[role="menuitem"]:not(:disabled)',
                            ) ?? [],
                          ),
                          activeMenuIndex = menuItems.indexOf(
                            document.activeElement,
                          ),
                          nextMenuIndex =
                            fileKeyEvent.key === "Home"
                              ? 0
                              : fileKeyEvent.key === "End"
                                ? menuItems.length - 1
                                : (activeMenuIndex +
                                    (fileKeyEvent.key === "ArrowDown"
                                      ? 1
                                      : menuItems.length - 1)) %
                                  menuItems.length;
                        menuItems[nextMenuIndex]?.focus();
                      } else
                        fileKeyEvent.key === "Escape" &&
                          (fileKeyEvent.preventDefault(),
                          fileKeyEvent.stopPropagation(),
                          setFileMenuOpen(false),
                          fileControlRef.current
                            ?.querySelector("button")
                            ?.focus());
                  },
                  children: [
                    i.jsx("button", {
                      type: "button",
                      className: `xl-file-tab ${isFileMenuOpen ? "xl-tab-open" : ""}`,
                      "aria-expanded": isFileMenuOpen,
                      "aria-haspopup": "menu",
                      onClick: () => {
                        (setFileMenuOpen(!isFileMenuOpen),
                          setRibbonPeekOpen(false));
                      },
                      children: i.jsx(LocalizedText, { children: "File" }),
                    }),
                    isFileMenuOpen && ribbonCommands.content,
                  ],
                }),
                i.jsx("div", {
                  className: "xl-ribbon-tabs",
                  role: "tablist",
                  "aria-label": text("Ribbon tabs"),
                  lang: language("Ribbon tabs"),
                  "data-ui-label": "Ribbon tabs",
                  children: Uh.map((tabTitle, tabIndex) =>
                    i.jsx(
                      "button",
                      {
                        type: "button",
                        id: `${ribbonId}-${tabIndex}`,
                        role: "tab",
                        "aria-selected": activeTab === tabTitle,
                        "aria-controls": `${ribbonId}-panel`,
                        "aria-expanded":
                          activeTab === tabTitle && isRibbonBodyVisible,
                        tabIndex: activeTab === tabTitle ? 0 : -1,
                        className:
                          activeTab === tabTitle ? "xl-tab-active" : "",
                        title: text(
                          "Double-click to collapse or expand the ribbon",
                        ),
                        onClick: () => selectTab(tabTitle),
                        onDoubleClick: toggleRibbonCollapsed,
                        onKeyDown: (tabKeyEvent) => {
                          if (!(
                            tabKeyEvent.defaultPrevented ||
                            tabKeyEvent.altKey ||
                            tabKeyEvent.ctrlKey ||
                            tabKeyEvent.metaKey
                          )) {
                            if (
                              tabKeyEvent.key === "ArrowRight" ||
                              tabKeyEvent.key === "ArrowLeft"
                            ) {
                              tabKeyEvent.preventDefault();
                              const nextTabIndex =
                                (tabIndex +
                                  (tabKeyEvent.key === "ArrowRight"
                                    ? 1
                                    : Uh.length - 1)) %
                                Uh.length;
                              (selectTab(Uh[nextTabIndex]),
                                document
                                  .getElementById(`${ribbonId}-${nextTabIndex}`)
                                  ?.focus());
                            } else if (
                              tabKeyEvent.key === "Home" ||
                              tabKeyEvent.key === "End"
                            ) {
                              tabKeyEvent.preventDefault();
                              const edgeTabIndex =
                                tabKeyEvent.key === "Home" ? 0 : Uh.length - 1;
                              (selectTab(Uh[edgeTabIndex]),
                                document
                                  .getElementById(`${ribbonId}-${edgeTabIndex}`)
                                  ?.focus());
                            } else if (tabKeyEvent.key === "ArrowDown") {
                              (tabKeyEvent.preventDefault(),
                                selectTab(tabTitle));
                              const command = commandsByTab[
                                tabTitle
                              ].commands.find((command) => !command.reason);
                              command && (pendingFocusId.current = command.id);
                            }
                          }
                        },
                        lang: language(
                          "Double-click to collapse or expand the ribbon",
                        ),
                        children: i.jsx(LocalizedText, { children: tabTitle }),
                      },
                      tabTitle,
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
                        isRibbonCollapsed ? "Expand ribbon" : "Collapse ribbon",
                      ),
                      "aria-expanded": isRibbonBodyVisible,
                      "aria-controls": `${ribbonId}-panel`,
                      title: isRibbonCollapsed
                        ? "Expand and keep ribbon open (Ctrl+F1)"
                        : "Collapse ribbon (Ctrl+F1)",
                      onClick: toggleRibbonCollapsed,
                      children: isRibbonCollapsed
                        ? i.jsx(ChevronDown, { size: 16 })
                        : i.jsx(ChevronUp, { size: 16 }),
                    }),
                    i.jsx(RibbonIconButton, {
                      icon: i.jsx(Undo2, { size: 16 }),
                      label: text("Undo (Ctrl+Z)"),
                      commandId: "quick-undo",
                      reason: quickAccessCommands[0].reason,
                      onClick: props.onUndo,
                      disabled: isDisabled || !props.canUndo,
                    }),
                    i.jsx(RibbonIconButton, {
                      icon: i.jsx(Redo2, { size: 16 }),
                      label: text("Redo (Ctrl+Y)"),
                      commandId: "quick-redo",
                      reason: quickAccessCommands[1].reason,
                      onClick: props.onRedo,
                      disabled: isDisabled || !props.canRedo,
                    }),
                    i.jsx(RibbonIconButton, {
                      icon: i.jsx(BookOpen, { size: 16 }),
                      label: text("Commands & shortcuts"),
                      onClick: () => setCommandDialogOpen(true),
                    }),
                    i.jsx("span", { className: "xl-tabbar-divider" }),
                    i.jsxs("button", {
                      type: "button",
                      className: "xl-ai-button",
                      "aria-label": text("AI assistant"),
                      onClick: () => props.onPanel("assistant"),
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
              ref: panelRef,
              id: `${ribbonId}-panel`,
              className: `xl-ribbon-body${isRibbonCollapsed && isRibbonPeekOpen ? " xl-ribbon-temporary" : ""}`,
              hidden: !isRibbonBodyVisible,
              role: "tabpanel",
              "aria-labelledby": `${ribbonId}-${Uh.indexOf(activeTab)}`,
              children:
                isRibbonBodyVisible &&
                i.jsx(i.Fragment, {
                  children: commandsByTab[activeTab].content,
                }),
            }),
          ],
        }),
      ],
    })
  );
}

export { WorkbookRibbon };
