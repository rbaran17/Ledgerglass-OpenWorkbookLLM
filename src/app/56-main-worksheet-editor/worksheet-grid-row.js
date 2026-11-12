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
    formatCellForEdit = (cell) =>
      formatCellEntryForEdit(cell, workbookMessages.entryLocale),
    inputRef = ne.useRef(null),
    suggestionListRef = ne.useRef(null),
    [caret, setCaret] = ne.useState(0);
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
            onClick: (headerClickEvent) =>
              onRowSelect(headerClickEvent.shiftKey),
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
        [...frozenColumns, ...columns].map((col, columnPosition) => {
          const mergeBound = mergeBounds.find((mergeBound) =>
            Tp(mergeBound, row, col),
          );
          let colSpan = 1,
            rowSpan = 1,
            cellAddress = `${columnIndexToLetters(col)}${row + 1}`;
          if (mergeBound && !sortedView) {
            const mergeColumns = (
                frozenColumns.includes(col) ? frozenColumns : columns
              ).filter(
                (mergeColumn) =>
                  mergeColumn >= mergeBound.firstCol &&
                  mergeColumn <= mergeBound.lastCol,
              ),
              mergeRows = displayedRows.filter(
                (displayedRow) =>
                  displayedRow >= mergeBound.firstRow &&
                  displayedRow <= mergeBound.lastRow,
              );
            if (col !== mergeColumns[0] || row !== mergeRows[0]) return null;
            ((colSpan = mergeColumns.length),
              (rowSpan = mergeRows.length),
              (cellAddress = `${columnIndexToLetters(mergeBound.firstCol)}${mergeBound.firstRow + 1}`));
          }
          const cell = cells[cellAddress],
            displayCell = {
              ...cell,
              raw: cell?.raw ?? null,
              style: { ...cell?.style, ...ruleStyles(cellAddress) },
            },
            cellValue = calc(sheetId, cellAddress),
            isFormula =
              typeof displayCell?.raw == "string" &&
              displayCell.raw.startsWith("=") &&
              !displayCell.literal,
            isActive = active === cellAddress,
            cellStyle = {
              ...cellStyleToCss(displayCell.style, cellValue),
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
                columnPosition === frozenColumns.length &&
                  leftGap > 0 &&
                  i.jsx("td", {
                    className: "lg-column-spacer",
                    "aria-hidden": "true",
                  }),
                i.jsxs(
                  "td",
                  {
                    id: `${idPrefix}-${cellAddress}`,
                    colSpan: colSpan,
                    rowSpan: rowSpan,
                    "aria-colspan": colSpan > 1 ? colSpan : undefined,
                    "aria-rowspan": rowSpan > 1 ? rowSpan : undefined,
                    role: "gridcell",
                    "aria-label": `${cellAddress} ${workbookMessages.entryLocale ? formatLocalizedNumberDisplay(cellValue, displayCell, workbook.datePolicy?.dateSystem, workbookMessages.entryLocale) : String(cellValue ?? "")}`,
                    "aria-colindex": col + 2,
                    "aria-selected": selectionBounds.some((selectionBound) =>
                      Tp(selectionBound, row, col),
                    ),
                    "data-address": cellAddress,
                    "data-formula-reference":
                      formulaReference && Tp(formulaReference, row, col)
                        ? "true"
                        : undefined,
                    className: `${isActive ? "lg-cell-active" : ""} ${selectionBounds.some((selectionBound) => Tp(selectionBound, row, col)) ? "lg-cell-selected" : ""} ${isFormula ? "lg-formula-cell" : ""} ${typeof cellValue == "string" && cellValue.startsWith("#") ? "lg-cell-error" : ""} ${frozenColumns.includes(col) ? "lg-frozen-column" : ""}`,
                    style: cellStyle,
                    title: `${cellAddress}${displayCell?.note ? ` · ${displayCell.note}` : ""}${isFormula ? ` · ${displayCell.raw}` : ""}`,
                    onPointerDown: (event) => {
                      event.button !== 0 ||
                        editing !== null ||
                        (event.preventDefault(),
                        onDragStart(
                          cellAddress,
                          event.ctrlKey || event.metaKey,
                        ),
                        select(
                          cellAddress,
                          event.shiftKey,
                          event.ctrlKey || event.metaKey,
                        ));
                    },
                    onPointerEnter: (pointerEnterEvent) =>
                      onDragEnter(cellAddress, pointerEnterEvent.buttons),
                    onClick: (clickEvent) => {
                      clickEvent.detail === 0 &&
                        ((formulaDraftActive &&
                          onFormulaPick(cellAddress, clickEvent.shiftKey)) ||
                          select(
                            cellAddress,
                            clickEvent.shiftKey,
                            clickEvent.ctrlKey || clickEvent.metaKey,
                          ));
                    },
                    onDoubleClick: () => {
                      formulaDraftActive ||
                        onPivotDrill(cellAddress) ||
                        readOnly ||
                        setEditing(formatCellForEdit(displayCell));
                    },
                    children: [
                      isActive && editing !== null
                        ? i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("input", {
                                ref: inputRef,
                                className: "lg-cell-editor",
                                onFocus: (event) =>
                                  onFormulaInput(event.currentTarget),
                                autoFocus: true,
                                readOnly,
                                "aria-label": `Edit ${cellAddress}`,
                                "aria-describedby": formulaDraftActive
                                  ? formulaHelpId
                                  : undefined,
                                value: editing,
                                onSelect: (event) => {
                                  (onFormulaInput(event.currentTarget),
                                    setCaret(
                                      event.currentTarget.selectionStart ??
                                        editing.length,
                                    ));
                                },
                                onChange: (event) => {
                                  (setEditing(event.target.value),
                                    onFormulaInput(event.currentTarget, true),
                                    setCaret(
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
                                        suggestionListRef.current?.handleKeyDown(
                                          event,
                                        )
                                      )
                                        return;
                                      (event.preventDefault(), onCancelEdit());
                                      return;
                                    }
                                    onFormulaKey(event) ||
                                      suggestionListRef.current?.handleKeyDown(
                                        event,
                                      ) ||
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
                                ref: suggestionListRef,
                                onExit: () => {
                                  saveCell(editing);
                                },
                                value: editing,
                                caret: caret,
                                workbook,
                                sheetId,
                                inputRef: inputRef,
                                disabled: readOnly,
                                onAccept: (nextValue, nextCaret) => {
                                  (setEditing(nextValue),
                                    setCaret(nextCaret),
                                    requestAnimationFrame(() => {
                                      (inputRef.current?.focus(),
                                        inputRef.current?.setSelectionRange(
                                          nextCaret,
                                          nextCaret,
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
                                  ...textWrapToCss(displayCell.style),
                                  ...(fixedHeight && rowSpan === 1
                                    ? {
                                        maxHeight: minimumHeight - 1,
                                        minHeight: 0,
                                      }
                                    : {}),
                                },
                                children:
                                  showFormulas && isFormula
                                    ? String(displayCell.raw)
                                    : i.jsx(ConditionalFormatIndicator, {
                                        decoration:
                                          conditionalFormatter(cellAddress),
                                        children: i.jsx(CellValuePreview, {
                                          value: cellValue,
                                          cell: displayCell,
                                          dateSystem:
                                            workbook.datePolicy?.dateSystem,
                                        }),
                                      }),
                              }),
                              filterHeader?.some(
                                (filterHeaderEntry) =>
                                  filterHeaderEntry.column === col,
                              ) &&
                                i.jsx("button", {
                                  className: "lg-header-filter",
                                  "aria-label": `Filter ${filterHeader.find((columnFilterEntry) => columnFilterEntry.column === col).label} column`,
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
                              displayCell?.hyperlink &&
                                (displayCell.hyperlink.kind === "external"
                                  ? i.jsx("a", {
                                      className: "lg-cell-link",
                                      href: displayCell.hyperlink.url,
                                      target: "_blank",
                                      rel: "noopener noreferrer",
                                      "aria-label": `Open link at ${cellAddress}`,
                                      title:
                                        displayCell.hyperlink.label ||
                                        displayCell.hyperlink.url,
                                      onPointerDown: (event) =>
                                        event.stopPropagation(),
                                      onClick: (event) =>
                                        event.stopPropagation(),
                                      children: "↗",
                                    })
                                  : i.jsx("button", {
                                      className: "lg-cell-link",
                                      "aria-label": `Open link at ${cellAddress}`,
                                      title:
                                        displayCell.hyperlink.label ||
                                        `Go to ${displayCell.hyperlink.address}`,
                                      onPointerDown: (event) =>
                                        event.stopPropagation(),
                                      onClick: (event) => {
                                        event.stopPropagation();
                                        const hyperlink = displayCell.hyperlink;
                                        hyperlink?.kind === "internal" &&
                                          onLink(
                                            hyperlink.sheetId,
                                            hyperlink.address,
                                          );
                                      },
                                      children: "↗",
                                    })),
                              commentAddresses.has(cellAddress) &&
                                i.jsx("button", {
                                  className: "lg-comment-marker",
                                  style: {
                                    right: displayCell?.note ? 8 : 0,
                                  },
                                  "aria-label": `Review comments at ${cellAddress}`,
                                  title: text("Open review discussion"),
                                  onPointerDown: (event) =>
                                    event.stopPropagation(),
                                  onClick: (event) => {
                                    (event.stopPropagation(),
                                      onReview(cellAddress));
                                  },
                                  lang: language("Open review discussion"),
                                  children: "●",
                                }),
                              displayCell?.note &&
                                i.jsx("i", { className: "lg-note-marker" }),
                              displayCell?.tickmark &&
                                i.jsx("span", {
                                  className: "lg-tick",
                                  children: displayCell.tickmark,
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
                          onClick: (fillClickEvent) => {
                            (fillClickEvent.stopPropagation(),
                              fillClickEvent.detail === 0 && onFillOptions());
                          },
                          onDoubleClick: (event) => event.stopPropagation(),
                          lang: language("Fill handle"),
                          "data-ui-label": "Fill handle",
                        }),
                    ],
                  },
                  `${cellAddress}:${row}:${col}`,
                ),
              ],
            },
            `${cellAddress}:${row}:${col}`,
          );
        }),
        rightGap > 0 &&
          i.jsx("td", { className: "lg-column-spacer", "aria-hidden": "true" }),
      ],
    }),
  });
}

export { WorksheetGridRow };
