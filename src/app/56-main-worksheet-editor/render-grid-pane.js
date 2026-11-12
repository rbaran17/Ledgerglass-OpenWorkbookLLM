import { Pv, _t, columnIndexToLetters } from "../03-workbook-model.js";
import i from "react/jsx-runtime";
import { LocalizedText } from "../04-cell-input-paste-and-fill.js";
import { DrawingObjectOverlay } from "../32-drawing-and-chart-export.js";
import { Tp } from "../53-find-replace-go-to.js";

function createGridPaneRenderer({
  text,
  language,
  editorIdPrefix,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  setContextMenuState,
  activeCellAddress,
  selectionAreas,
  selectionAreaRanges,
  filterValueSearch,
  activeColumnSort,
  sheetSplitPanes,
  splitPaneSizes,
  setStatusNotice,
  isOperationBusy,
  isGridlinesVisible,
  isHeadingsVisible,
  dragAnchorRef,
  workbookCalculator,
  rowVisibilityIndex,
  columnVisibilityIndex,
  pinnedRowCount,
  frozenColumnCount,
  zoomScale,
  rowHeaderWidth,
  gridContainerRef,
  mergeRanges,
  rowAxisLayout,
  columnAxisLayout,
  yn,
  frozenRowIndexes,
  frozenColumnIndexes,
  scrollableRowLayout,
  scrollableColumnLayout,
  scrollableRowPositionOf,
  frozenRowsHeight,
  frozenColumnsWidth,
  computeGridViewport,
  applyPaneScrollOffsets,
  commitWorkbookEdit,
  runGuardedAction,
  handleFormulaPointerDown,
  commitCellEdit,
  selectCellAddress,
  selectionToTsv,
  pasteClipboardText,
  selectAllCells,
  selectRowOrColumn,
  handleGridKeyDown,
  renderColumnHeader,
  renderGridRow,
  renderVirtualSpacer,
  activateGridPane,
}) {
  function renderGridPane(fe) {
    const Oe = _t(activeCellAddress),
      Be =
        filterValueSearch || activeColumnSort
          ? undefined
          : mergeRanges.find((Sr) => Tp(Sr, Oe.row, Oe.col)),
      Ue = Be
        ? `${columnIndexToLetters(Be.firstCol)}${Be.firstRow + 1}`
        : activeCellAddress,
      Xe = splitPaneSizes.views[fe],
      gt = computeGridViewport(Xe),
      {
        rowWindow,
        columnWindow,
        shownRows,
        shownColumns,
        leftGap,
        rightGap,
        visibleColumnCount,
      } = gt;
    return i.jsxs("div", {
      className: "lg-grid-stage",
      onPointerDownCapture: (Sr) => {
        handleFormulaPointerDown(Sr) ||
          activateGridPane(fe) ||
          (Sr.preventDefault(), Sr.stopPropagation());
      },
      onFocusCapture: () => activateGridPane(fe),
      children: [
        i.jsx("div", {
          className: "lg-grid-scroll",
          id: `${editorIdPrefix}-${fe}-grid`,
          ref: (Sr) => {
            (splitPaneSizes.bindings[fe](Sr),
              splitPaneSizes.active.current === fe &&
                (gridContainerRef.current = Sr));
          },
          tabIndex: 0,
          role: "grid",
          "aria-rowcount": activeSheet.rowCount + 1,
          "aria-colcount": activeSheet.colCount + 1,
          "aria-readonly": isEditingLocked,
          "aria-multiselectable": "true",
          "aria-describedby": `${editorIdPrefix}-grid-help`,
          "aria-activedescendant":
            [...frozenRowIndexes, ...shownRows].includes(
              _t(activeCellAddress).row,
            ) &&
            [...frozenColumnIndexes, ...shownColumns].includes(
              _t(activeCellAddress).col,
            )
              ? `${editorIdPrefix}-${fe}-${Ue}`
              : undefined,
          "aria-label": sheetSplitPanes
            ? `Worksheet pane ${fe.replace("-", " ")}`
            : text("Editable worksheet grid"),
          onScroll: (Sr) => {
            applyPaneScrollOffsets(
              fe,
              Sr.currentTarget.scrollTop,
              Sr.currentTarget.scrollLeft,
            );
          },
          onKeyDown: (Sr) => {
            var Wi;
            if (
              sheetSplitPanes &&
              Sr.key === "F6" &&
              Sr.target === Sr.currentTarget &&
              !Sr.defaultPrevented &&
              !Sr.ctrlKey &&
              !Sr.metaKey &&
              !Sr.altKey
            ) {
              Sr.preventDefault();
              const Di = Pv(sheetSplitPanes),
                $o = Di.indexOf(fe),
                Ds = Di[($o + (Sr.shiftKey ? Di.length - 1 : 1)) % Di.length];
              activateGridPane(Ds) &&
                ((Wi = splitPaneSizes.refs.current[Ds]) == null ||
                  Wi.focus({ preventScroll: true }));
            } else handleGridKeyDown(Sr);
          },
          onContextMenu: (Sr) => {
            const Wi = Sr.target;
            if (Wi.closest("input,textarea,button,a,[contenteditable=true]"))
              return;
            const Di = Wi.closest("[data-address]"),
              $o = Wi.closest("[data-column]"),
              Ds = Wi.closest("[data-row-header]");
            if (!(!Di && !$o && !Ds) && commitCellEdit()) {
              if ((Sr.preventDefault(), $o)) {
                const qo = Number($o.dataset.column);
                selectionAreaRanges.some(
                  (yo) =>
                    yo.firstRow === 0 &&
                    yo.lastRow === activeSheet.rowCount - 1 &&
                    qo >= yo.firstCol &&
                    qo <= yo.lastCol,
                ) || selectRowOrColumn("column", qo, false);
              } else if (Ds) {
                const qo = Number(Ds.dataset.rowHeader);
                selectionAreaRanges.some(
                  (yo) =>
                    yo.firstCol === 0 &&
                    yo.lastCol === activeSheet.colCount - 1 &&
                    qo >= yo.firstRow &&
                    qo <= yo.lastRow,
                ) || selectRowOrColumn("row", qo, false);
              } else if (Di) {
                const qo = _t(Di.dataset.address);
                selectionAreaRanges.some((yo) => Tp(yo, qo.row, qo.col)) ||
                  selectCellAddress(Di.dataset.address);
              }
              setContextMenuState({
                x: Sr.clientX,
                y: Sr.clientY,
                kind: $o ? "column" : Ds ? "row" : "cell",
              });
            }
          },
          onCopy: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(),
              runGuardedAction(() =>
                Sr.clipboardData.setData("text/plain", selectionToTsv()),
              ));
          },
          onCut: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              isEditingLocked ||
              isOperationBusy ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(),
              runGuardedAction(() => {
                (Sr.clipboardData.setData("text/plain", selectionToTsv(true)),
                  setStatusNotice(
                    "Cut is ready. Select a destination and paste to preview the move.",
                  ));
              }));
          },
          onPaste: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              isEditingLocked ||
              isOperationBusy ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(),
              pasteClipboardText(Sr.clipboardData.getData("text/plain")));
          },
          onPointerUp: () => {
            dragAnchorRef.current = null;
          },
          onPointerLeave: (Sr) => {
            Sr.buttons === 0 && (dragAnchorRef.current = null);
          },
          children: i.jsxs("table", {
            className: `lg-grid ${isGridlinesVisible ? "" : "no-gridlines"} ${isHeadingsVisible ? "" : "no-headings"}`,
            style: {
              zoom: zoomScale,
              width: columnWindow.total + frozenColumnsWidth + rowHeaderWidth,
              minWidth: 0,
            },
            role: "presentation",
            children: [
              i.jsxs("colgroup", {
                children: [
                  isHeadingsVisible &&
                    i.jsx("col", { style: { width: rowHeaderWidth } }),
                  frozenColumnIndexes.map((Sr) =>
                    i.jsx(
                      "col",
                      {
                        style: {
                          width: columnAxisLayout.size(
                            columnVisibilityIndex.viewIndex(Sr),
                          ),
                        },
                      },
                      Sr,
                    ),
                  ),
                  leftGap > 0 && i.jsx("col", { style: { width: leftGap } }),
                  shownColumns.map((Sr) =>
                    i.jsx(
                      "col",
                      {
                        style: {
                          width: columnAxisLayout.size(
                            columnVisibilityIndex.viewIndex(Sr),
                          ),
                        },
                      },
                      Sr,
                    ),
                  ),
                  rightGap > 0 && i.jsx("col", { style: { width: rightGap } }),
                ],
              }),
              i.jsx("thead", {
                children: i.jsxs("tr", {
                  role: "row",
                  "aria-rowindex": 1,
                  children: [
                    isHeadingsVisible &&
                      i.jsx("th", {
                        role: "columnheader",
                        className: "lg-corner",
                        "aria-label": text("Select all cells"),
                        onClick: selectAllCells,
                        lang: language("Select all cells"),
                        "data-ui-label": "Select all cells",
                        children: i.jsx("span", {
                          className: "lg-sr-only",
                          children: i.jsx(LocalizedText, {
                            children: "Select all cells",
                          }),
                        }),
                      }),
                    frozenColumnIndexes.map((Sr) =>
                      renderColumnHeader(Sr, true),
                    ),
                    leftGap > 0 &&
                      i.jsx("th", {
                        className: "lg-column-spacer",
                        "aria-hidden": "true",
                      }),
                    shownColumns.map((Sr) => renderColumnHeader(Sr)),
                    rightGap > 0 &&
                      i.jsx("th", {
                        className: "lg-column-spacer",
                        "aria-hidden": "true",
                      }),
                  ],
                }),
              }),
              i.jsxs("tbody", {
                children: [
                  frozenRowIndexes.map((Sr) => renderGridRow(Sr, true, gt, fe)),
                  renderVirtualSpacer(
                    rowWindow.leading,
                    "before",
                    visibleColumnCount,
                  ),
                  shownRows.map((Sr) => renderGridRow(Sr, false, gt, fe)),
                  renderVirtualSpacer(
                    rowWindow.trailing,
                    "after",
                    visibleColumnCount,
                  ),
                ],
              }),
            ],
          }),
        }),
        i.jsx(DrawingObjectOverlay, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          selection: JSON.stringify(selectionAreas),
          calc: workbookCalculator,
          readOnly: isEditingLocked,
          scale: zoomScale,
          onWorkbookChange: commitWorkbookEdit,
          positionForCell: (Sr) => {
            const Wi = _t(Sr);
            if (
              rowVisibilityIndex.isHidden(Wi.row) ||
              columnVisibilityIndex.isHidden(Wi.col)
            )
              return null;
            const Di = scrollableRowPositionOf(Wi.row),
              $o = Wi.row < pinnedRowCount,
              Ds = Wi.col < frozenColumnCount;
            if (!$o && Di === undefined) return null;
            const qo =
                (rowHeaderWidth +
                  (Ds
                    ? columnAxisLayout.offset(
                        columnVisibilityIndex.viewIndex(Wi.col),
                      )
                    : frozenColumnsWidth +
                      scrollableColumnLayout.offset(
                        columnVisibilityIndex.viewIndex(Wi.col) - yn,
                      ) -
                      columnWindow.logicalOffset)) *
                zoomScale,
              yo =
                ((isHeadingsVisible ? 22 : 0) +
                  ($o
                    ? rowAxisLayout.offset(rowVisibilityIndex.viewIndex(Wi.row))
                    : frozenRowsHeight +
                      scrollableRowLayout.offset(Math.max(0, Di)) -
                      rowWindow.logicalOffset)) *
                zoomScale;
            return qo < -5e3 || yo < -5e3 || qo > Xe.width || yo > Xe.height
              ? null
              : { left: qo, top: yo };
          },
        }),
      ],
    });
  }
  return { renderGridPane };
}

export { createGridPaneRenderer };
