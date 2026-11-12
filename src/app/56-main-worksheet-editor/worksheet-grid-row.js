import {
  formatCellEntryForEdit,
  formatLocalizedNumberDisplay,
  useWorkbookMessages,
} from "../04-cell-input-paste-and-fill.js";
import ne from "react";
import i from "react/jsx-runtime";
import {
  _t,
  cellStyleToCss,
  columnIndexToLetters,
  textWrapToCss,
} from "../03-workbook-model.js";
import {
  FormulaSuggestionList,
  Tp,
  isWithinFormulaSuggestions,
} from "../53-find-replace-go-to.js";
import { ConditionalFormatIndicator } from "../23-format-cells-and-conditional-formatting.js";
import { CellValuePreview } from "../22-themes-and-cell-styles.js";

function WorksheetGridRow({
  idPrefix,
  pinned,
  pinnedTop,
  frozenColumns,
  frozenColumnOffsets,
  headings,
  onRowSelect,
  onDragStart,
  onDragEnter,
  onNavigate,
  onPivotDrill,
  onFillStart,
  onFillOptions,
  fillDragging,
  commentAddresses,
  onReview,
  onLink,
  row,
  columns,
  leftGap,
  rightGap,
  minimumHeight,
  fixedHeight,
  onResize,
  onAutoFit,
  onCancelEdit,
  onFormulaInput,
  onFormulaKey,
  formulaHelpId,
  onFormulaPick,
  formulaReference,
  formulaDraftActive,
  workbook,
  mergeBounds,
  conditionalFormatter,
  displayedRows,
  sortedView,
  filterHeader,
  filteredColumns,
  onColumnFilter,
  ruleStyles,
  sheetId,
  cells,
  active,
  selectedBounds,
  selectionBounds,
  calc,
  showFormulas,
  editing,
  setEditing,
  select,
  saveCell,
  readOnly,
}) {
  const { text, language } = useWorkbookMessages(),
    workbookMessages = useWorkbookMessages(),
    se = (cell) => formatCellEntryForEdit(cell, workbookMessages.entryLocale),
    inputRef = ne.useRef(null),
    Ee = ne.useRef(null),
    [caret, ye] = ne.useState(0);
  return i.jsx(i.Fragment, {
    children: i.jsxs("tr", {
      role: "row",
      "data-row": row,
      style: { height: minimumHeight, ...(pinned ? { top: pinnedTop } : {}) },
      className: pinned ? "lg-pinned-row" : "",
      "aria-rowindex": row + 2,
      children: [
        headings &&
          i.jsxs("th", {
            role: "rowheader",
            scope: "row",
            "data-row-header": row,
            onClick: (Te) => onRowSelect(Te.shiftKey),
            "aria-label": `Select row ${row + 1}`,
            className: _t(active).row === row ? "lg-selected-header" : "",
            children: [
              row + 1,
              i.jsx("span", {
                className: "lg-row-resize",
                role: "separator",
                "aria-label": `Resize row ${row + 1}`,
                "aria-orientation": "horizontal",
                "aria-disabled": readOnly,
                title: text(
                  "Drag to resize; double-click to AutoFit row height",
                ),
                onPointerDown: onResize,
                onClick: (event) => event.stopPropagation(),
                onDoubleClick: (event) => {
                  (event.stopPropagation(), onAutoFit());
                },
                lang: language(
                  "Drag to resize; double-click to AutoFit row height",
                ),
              }),
            ],
          }),
        [...frozenColumns, ...columns].map((col, Pe) => {
          const mergeBound = mergeBounds.find((mergeBound) =>
            Tp(mergeBound, row, col),
          );
          let colSpan = 1,
            rowSpan = 1,
            Me = `${columnIndexToLetters(col)}${row + 1}`;
          if (mergeBound && !sortedView) {
            const lt = (
                frozenColumns.includes(col) ? frozenColumns : columns
              ).filter(
                (ft) => ft >= mergeBound.firstCol && ft <= mergeBound.lastCol,
              ),
              be = displayedRows.filter(
                (displayedRow) =>
                  displayedRow >= mergeBound.firstRow &&
                  displayedRow <= mergeBound.lastRow,
              );
            if (col !== lt[0] || row !== be[0]) return null;
            ((colSpan = lt.length),
              (rowSpan = be.length),
              (Me = `${columnIndexToLetters(mergeBound.firstCol)}${mergeBound.firstRow + 1}`));
          }
          const cell = cells[Me],
            je = {
              ...cell,
              raw: cell?.raw ?? null,
              style: { ...cell?.style, ...ruleStyles(Me) },
            },
            $e = calc(sheetId, Me),
            Je =
              typeof je?.raw == "string" &&
              je.raw.startsWith("=") &&
              !je.literal,
            ze = active === Me,
            at = {
              ...cellStyleToCss(je.style, $e),
              ...(frozenColumns.includes(col)
                ? {
                    position: "sticky",
                    left: frozenColumnOffsets[col],
                    zIndex: 3,
                  }
                : {}),
            };
          return i.jsxs(
            ne.Fragment,
            {
              children: [
                Pe === frozenColumns.length &&
                  leftGap > 0 &&
                  i.jsx("td", {
                    className: "lg-column-spacer",
                    "aria-hidden": "true",
                  }),
                i.jsxs(
                  "td",
                  {
                    id: `${idPrefix}-${Me}`,
                    colSpan: colSpan,
                    rowSpan: rowSpan,
                    "aria-colspan": colSpan > 1 ? colSpan : undefined,
                    "aria-rowspan": rowSpan > 1 ? rowSpan : undefined,
                    role: "gridcell",
                    "aria-label": `${Me} ${workbookMessages.entryLocale ? formatLocalizedNumberDisplay($e, je, workbook.datePolicy?.dateSystem, workbookMessages.entryLocale) : String($e ?? "")}`,
                    "aria-colindex": col + 2,
                    "aria-selected": selectionBounds.some((selectionBound) =>
                      Tp(selectionBound, row, col),
                    ),
                    "data-address": Me,
                    "data-formula-reference":
                      formulaReference && Tp(formulaReference, row, col)
                        ? "true"
                        : undefined,
                    className: `${ze ? "lg-cell-active" : ""} ${selectionBounds.some((selectionBound) => Tp(selectionBound, row, col)) ? "lg-cell-selected" : ""} ${Je ? "lg-formula-cell" : ""} ${typeof $e == "string" && $e.startsWith("#") ? "lg-cell-error" : ""} ${frozenColumns.includes(col) ? "lg-frozen-column" : ""}`,
                    style: at,
                    title: `${Me}${je?.note ? ` · ${je.note}` : ""}${Je ? ` · ${je.raw}` : ""}`,
                    onPointerDown: (event) => {
                      event.button !== 0 ||
                        editing !== null ||
                        (event.preventDefault(),
                        onDragStart(Me, event.ctrlKey || event.metaKey),
                        select(
                          Me,
                          event.shiftKey,
                          event.ctrlKey || event.metaKey,
                        ));
                    },
                    onPointerEnter: (lt) => onDragEnter(Me, lt.buttons),
                    onClick: (lt) => {
                      lt.detail === 0 &&
                        ((formulaDraftActive &&
                          onFormulaPick(Me, lt.shiftKey)) ||
                          select(Me, lt.shiftKey, lt.ctrlKey || lt.metaKey));
                    },
                    onDoubleClick: () => {
                      formulaDraftActive ||
                        onPivotDrill(Me) ||
                        readOnly ||
                        setEditing(se(je));
                    },
                    children: [
                      ze && editing !== null
                        ? i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("input", {
                                ref: inputRef,
                                className: "lg-cell-editor",
                                onFocus: (event) =>
                                  onFormulaInput(event.currentTarget),
                                autoFocus: true,
                                readOnly,
                                "aria-label": `Edit ${Me}`,
                                "aria-describedby": formulaDraftActive
                                  ? formulaHelpId
                                  : undefined,
                                value: editing,
                                onSelect: (event) => {
                                  (onFormulaInput(event.currentTarget),
                                    ye(
                                      event.currentTarget.selectionStart ??
                                        editing.length,
                                    ));
                                },
                                onChange: (event) => {
                                  (setEditing(event.target.value),
                                    onFormulaInput(event.currentTarget, true),
                                    ye(
                                      event.target.selectionStart ??
                                        event.target.value.length,
                                    ));
                                },
                                onClick: (event) => event.stopPropagation(),
                                onBlur: (event) => {
                                  isWithinFormulaSuggestions(
                                    event.currentTarget,
                                    event.relatedTarget,
                                  ) || saveCell(editing);
                                },
                                onKeyDown: (event) => {
                                  if (
                                    (event.stopPropagation(),
                                    !event.nativeEvent.isComposing)
                                  ) {
                                    if (event.key === "Escape") {
                                      if (
                                        !formulaReference &&
                                        editing.length > 1 &&
                                        Ee.current?.handleKeyDown(event)
                                      )
                                        return;
                                      (event.preventDefault(), onCancelEdit());
                                      return;
                                    }
                                    onFormulaKey(event) ||
                                      Ee.current?.handleKeyDown(event) ||
                                      ((event.key === "Enter" ||
                                        event.key === "Tab") &&
                                        (event.preventDefault(),
                                        saveCell(editing) &&
                                          onNavigate(
                                            event.key,
                                            event.shiftKey,
                                          )),
                                      event.key === "Escape" &&
                                        (event.preventDefault(),
                                        onCancelEdit()));
                                  }
                                },
                              }),
                              i.jsx(FormulaSuggestionList, {
                                ref: Ee,
                                onExit: () => {
                                  saveCell(editing);
                                },
                                value: editing,
                                caret: caret,
                                workbook,
                                sheetId,
                                inputRef: inputRef,
                                disabled: readOnly,
                                onAccept: (lt, be) => {
                                  (setEditing(lt),
                                    ye(be),
                                    requestAnimationFrame(() => {
                                      (inputRef.current?.focus(),
                                        inputRef.current?.setSelectionRange(
                                          be,
                                          be,
                                        ));
                                    }));
                                },
                              }),
                            ],
                          })
                        : i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("span", {
                                style: {
                                  ...textWrapToCss(je.style),
                                  ...(fixedHeight && rowSpan === 1
                                    ? {
                                        maxHeight: minimumHeight - 1,
                                        minHeight: 0,
                                      }
                                    : {}),
                                },
                                children:
                                  showFormulas && Je
                                    ? String(je.raw)
                                    : i.jsx(ConditionalFormatIndicator, {
                                        decoration: conditionalFormatter(Me),
                                        children: i.jsx(CellValuePreview, {
                                          value: $e,
                                          cell: je,
                                          dateSystem:
                                            workbook.datePolicy?.dateSystem,
                                        }),
                                      }),
                              }),
                              filterHeader?.some((lt) => lt.column === col) &&
                                i.jsx("button", {
                                  className: "lg-header-filter",
                                  "aria-label": `Filter ${filterHeader.find((lt) => lt.column === col).label} column`,
                                  "aria-pressed": filteredColumns.includes(col),
                                  onPointerDown: (event) =>
                                    event.stopPropagation(),
                                  onClick: (event) => {
                                    (event.stopPropagation(),
                                      onColumnFilter(col));
                                  },
                                  title: text("Filter this column"),
                                  lang: language("Filter this column"),
                                  children: "▾",
                                }),
                              je?.hyperlink &&
                                (je.hyperlink.kind === "external"
                                  ? i.jsx("a", {
                                      className: "lg-cell-link",
                                      href: je.hyperlink.url,
                                      target: "_blank",
                                      rel: "noopener noreferrer",
                                      "aria-label": `Open link at ${Me}`,
                                      title:
                                        je.hyperlink.label || je.hyperlink.url,
                                      onPointerDown: (event) =>
                                        event.stopPropagation(),
                                      onClick: (event) =>
                                        event.stopPropagation(),
                                      children: "↗",
                                    })
                                  : i.jsx("button", {
                                      className: "lg-cell-link",
                                      "aria-label": `Open link at ${Me}`,
                                      title:
                                        je.hyperlink.label ||
                                        `Go to ${je.hyperlink.address}`,
                                      onPointerDown: (event) =>
                                        event.stopPropagation(),
                                      onClick: (event) => {
                                        event.stopPropagation();
                                        const hyperlink = je.hyperlink;
                                        hyperlink?.kind === "internal" &&
                                          onLink(
                                            hyperlink.sheetId,
                                            hyperlink.address,
                                          );
                                      },
                                      children: "↗",
                                    })),
                              commentAddresses.has(Me) &&
                                i.jsx("button", {
                                  className: "lg-comment-marker",
                                  style: {
                                    right: je?.note ? 8 : 0,
                                  },
                                  "aria-label": `Review comments at ${Me}`,
                                  title: text("Open review discussion"),
                                  onPointerDown: (event) =>
                                    event.stopPropagation(),
                                  onClick: (event) => {
                                    (event.stopPropagation(), onReview(Me));
                                  },
                                  lang: language("Open review discussion"),
                                  children: "●",
                                }),
                              je?.note &&
                                i.jsx("i", { className: "lg-note-marker" }),
                              je?.tickmark &&
                                i.jsx("span", {
                                  className: "lg-tick",
                                  children: je.tickmark,
                                }),
                            ],
                          }),
                      !readOnly &&
                        !fillDragging &&
                        editing === null &&
                        row === selectedBounds.lastRow &&
                        col === selectedBounds.lastCol &&
                        i.jsx("button", {
                          className: "lg-fill-handle",
                          "aria-label": text("Fill handle"),
                          title: text(
                            "Drag to extend the selection; use Fill Series for keyboard options",
                          ),
                          onPointerDown: onFillStart,
                          onClick: (lt) => {
                            (lt.stopPropagation(),
                              lt.detail === 0 && onFillOptions());
                          },
                          onDoubleClick: (event) => event.stopPropagation(),
                          lang: language("Fill handle"),
                          "data-ui-label": "Fill handle",
                        }),
                    ],
                  },
                  `${Me}:${row}:${col}`,
                ),
              ],
            },
            `${Me}:${row}:${col}`,
          );
        }),
        rightGap > 0 &&
          i.jsx("td", { className: "lg-column-spacer", "aria-hidden": "true" }),
      ],
    }),
  });
}

export { WorksheetGridRow };
