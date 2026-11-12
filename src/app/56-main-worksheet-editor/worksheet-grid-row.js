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
import { ConditionalFormatIndicator } from "../23-format-cells-dialog.js";
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
    ge = useWorkbookMessages(),
    se = (Te) => formatCellEntryForEdit(Te, ge.entryLocale),
    we = ne.useRef(null),
    Ee = ne.useRef(null),
    [Ae, ye] = ne.useState(0);
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
                onClick: (Te) => Te.stopPropagation(),
                onDoubleClick: (Te) => {
                  (Te.stopPropagation(), onAutoFit());
                },
                lang: language(
                  "Drag to resize; double-click to AutoFit row height",
                ),
              }),
            ],
          }),
        [...frozenColumns, ...columns].map((Te, Pe) => {
          const qe = mergeBounds.find((lt) => Tp(lt, row, Te));
          let ht = 1,
            Re = 1,
            Me = `${columnIndexToLetters(Te)}${row + 1}`;
          if (qe && !sortedView) {
            const lt = (
                frozenColumns.includes(Te) ? frozenColumns : columns
              ).filter((ft) => ft >= qe.firstCol && ft <= qe.lastCol),
              be = displayedRows.filter(
                (ft) => ft >= qe.firstRow && ft <= qe.lastRow,
              );
            if (Te !== lt[0] || row !== be[0]) return null;
            ((ht = lt.length),
              (Re = be.length),
              (Me = `${columnIndexToLetters(qe.firstCol)}${qe.firstRow + 1}`));
          }
          const ke = cells[Me],
            je = {
              ...ke,
              raw: ke?.raw ?? null,
              style: { ...ke?.style, ...ruleStyles(Me) },
            },
            $e = calc(sheetId, Me),
            Je =
              typeof je?.raw == "string" &&
              je.raw.startsWith("=") &&
              !je.literal,
            ze = active === Me,
            at = {
              ...cellStyleToCss(je.style, $e),
              ...(frozenColumns.includes(Te)
                ? {
                    position: "sticky",
                    left: frozenColumnOffsets[Te],
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
                    colSpan: ht,
                    rowSpan: Re,
                    "aria-colspan": ht > 1 ? ht : undefined,
                    "aria-rowspan": Re > 1 ? Re : undefined,
                    role: "gridcell",
                    "aria-label": `${Me} ${ge.entryLocale ? formatLocalizedNumberDisplay($e, je, workbook.datePolicy?.dateSystem, ge.entryLocale) : String($e ?? "")}`,
                    "aria-colindex": Te + 2,
                    "aria-selected": selectionBounds.some((lt) =>
                      Tp(lt, row, Te),
                    ),
                    "data-address": Me,
                    "data-formula-reference":
                      formulaReference && Tp(formulaReference, row, Te)
                        ? "true"
                        : undefined,
                    className: `${ze ? "lg-cell-active" : ""} ${selectionBounds.some((lt) => Tp(lt, row, Te)) ? "lg-cell-selected" : ""} ${Je ? "lg-formula-cell" : ""} ${typeof $e == "string" && $e.startsWith("#") ? "lg-cell-error" : ""} ${frozenColumns.includes(Te) ? "lg-frozen-column" : ""}`,
                    style: at,
                    title: `${Me}${je != null && je.note ? ` · ${je.note}` : ""}${Je ? ` · ${je.raw}` : ""}`,
                    onPointerDown: (lt) => {
                      lt.button !== 0 ||
                        editing !== null ||
                        (lt.preventDefault(),
                        onDragStart(Me, lt.ctrlKey || lt.metaKey),
                        select(Me, lt.shiftKey, lt.ctrlKey || lt.metaKey));
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
                                ref: we,
                                className: "lg-cell-editor",
                                onFocus: (lt) =>
                                  onFormulaInput(lt.currentTarget),
                                autoFocus: true,
                                readOnly,
                                "aria-label": `Edit ${Me}`,
                                "aria-describedby": formulaDraftActive
                                  ? formulaHelpId
                                  : undefined,
                                value: editing,
                                onSelect: (lt) => {
                                  (onFormulaInput(lt.currentTarget),
                                    ye(
                                      lt.currentTarget.selectionStart ??
                                        editing.length,
                                    ));
                                },
                                onChange: (lt) => {
                                  (setEditing(lt.target.value),
                                    onFormulaInput(lt.currentTarget, true),
                                    ye(
                                      lt.target.selectionStart ??
                                        lt.target.value.length,
                                    ));
                                },
                                onClick: (lt) => lt.stopPropagation(),
                                onBlur: (lt) => {
                                  isWithinFormulaSuggestions(
                                    lt.currentTarget,
                                    lt.relatedTarget,
                                  ) || saveCell(editing);
                                },
                                onKeyDown: (lt) => {
                                  var be, ft;
                                  if (
                                    (lt.stopPropagation(),
                                    !lt.nativeEvent.isComposing)
                                  ) {
                                    if (lt.key === "Escape") {
                                      if (
                                        !formulaReference &&
                                        editing.length > 1 &&
                                        (be = Ee.current) != null &&
                                        be.handleKeyDown(lt)
                                      )
                                        return;
                                      (lt.preventDefault(), onCancelEdit());
                                      return;
                                    }
                                    onFormulaKey(lt) ||
                                      ((ft = Ee.current) != null &&
                                        ft.handleKeyDown(lt)) ||
                                      ((lt.key === "Enter" ||
                                        lt.key === "Tab") &&
                                        (lt.preventDefault(),
                                        saveCell(editing) &&
                                          onNavigate(lt.key, lt.shiftKey)),
                                      lt.key === "Escape" &&
                                        (lt.preventDefault(), onCancelEdit()));
                                  }
                                },
                              }),
                              i.jsx(FormulaSuggestionList, {
                                ref: Ee,
                                onExit: () => {
                                  saveCell(editing);
                                },
                                value: editing,
                                caret: Ae,
                                workbook,
                                sheetId,
                                inputRef: we,
                                disabled: readOnly,
                                onAccept: (lt, be) => {
                                  (setEditing(lt),
                                    ye(be),
                                    requestAnimationFrame(() => {
                                      var ft, pt;
                                      ((ft = we.current) == null || ft.focus(),
                                        (pt = we.current) == null ||
                                          pt.setSelectionRange(be, be));
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
                                  ...(fixedHeight && Re === 1
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
                              filterHeader?.some((lt) => lt.column === Te) &&
                                i.jsx("button", {
                                  className: "lg-header-filter",
                                  "aria-label": `Filter ${filterHeader.find((lt) => lt.column === Te).label} column`,
                                  "aria-pressed": filteredColumns.includes(Te),
                                  onPointerDown: (lt) => lt.stopPropagation(),
                                  onClick: (lt) => {
                                    (lt.stopPropagation(), onColumnFilter(Te));
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
                                      onPointerDown: (lt) =>
                                        lt.stopPropagation(),
                                      onClick: (lt) => lt.stopPropagation(),
                                      children: "↗",
                                    })
                                  : i.jsx("button", {
                                      className: "lg-cell-link",
                                      "aria-label": `Open link at ${Me}`,
                                      title:
                                        je.hyperlink.label ||
                                        `Go to ${je.hyperlink.address}`,
                                      onPointerDown: (lt) =>
                                        lt.stopPropagation(),
                                      onClick: (lt) => {
                                        lt.stopPropagation();
                                        const be = je.hyperlink;
                                        be?.kind === "internal" &&
                                          onLink(be.sheetId, be.address);
                                      },
                                      children: "↗",
                                    })),
                              commentAddresses.has(Me) &&
                                i.jsx("button", {
                                  className: "lg-comment-marker",
                                  style: {
                                    right: je != null && je.note ? 8 : 0,
                                  },
                                  "aria-label": `Review comments at ${Me}`,
                                  title: text("Open review discussion"),
                                  onPointerDown: (lt) => lt.stopPropagation(),
                                  onClick: (lt) => {
                                    (lt.stopPropagation(), onReview(Me));
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
                        Te === selectedBounds.lastCol &&
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
                          onDoubleClick: (lt) => lt.stopPropagation(),
                          lang: language("Fill handle"),
                          "data-ui-label": "Fill handle",
                        }),
                    ],
                  },
                  `${Me}:${row}:${Te}`,
                ),
              ],
            },
            `${Me}:${row}:${Te}`,
          );
        }),
        rightGap > 0 &&
          i.jsx("td", { className: "lg-column-spacer", "aria-hidden": "true" }),
      ],
    }),
  });
}

export { WorksheetGridRow };
