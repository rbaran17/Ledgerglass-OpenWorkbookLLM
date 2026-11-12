import { Pv, _t, columnIndexToLetters } from "../03-workbook-model.js";
import i from "react/jsx-runtime";
import { LocalizedText } from "../04-cell-input-paste-and-fill.js";
import { DrawingObjectOverlay } from "../32-drawing-objects-and-printing.js";
import { Tp } from "../53-find-replace-go-to.js";

function createGridPaneRenderer({
  viewAndSelectionActions,
  navigationActions,
  formulaEntryActions,
  cellEditActions,
  clipboardAndFindActions,
  gridInteractionActions,
  selectionState,
  filterSortState,
  operationStatusState,
  viewDisplayState,
  dragFillState,
  text,
  language,
  editorIdPrefix,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  selectionAreas,
  selectionAreaRanges,
  sheetSplitPanes,
  splitPaneSizes,
  workbookCalculator,
  rowVisibilityIndex,
  columnVisibilityIndex,
  pinnedRowCount,
  frozenColumnCount,
  zoomScale,
  rowHeaderWidth,
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
  renderColumnHeader,
  renderGridRow,
  renderVirtualSpacer,
}) {
  function renderGridPane(paneId) {
    const activeCell = _t(selectionState.activeCellAddress),
      activeMergeRange =
        filterSortState.filterValueSearch || filterSortState.activeColumnSort
          ? undefined
          : mergeRanges.find((mergeRange) =>
              Tp(mergeRange, activeCell.row, activeCell.col),
            ),
      activeCellLabel = activeMergeRange
        ? `${columnIndexToLetters(activeMergeRange.firstCol)}${activeMergeRange.firstRow + 1}`
        : selectionState.activeCellAddress,
      paneSize = splitPaneSizes.views[paneId],
      gridViewport = viewAndSelectionActions.computeGridViewport(paneSize),
      {
        rowWindow,
        columnWindow,
        shownRows,
        shownColumns,
        leftGap,
        rightGap,
        visibleColumnCount,
      } = gridViewport;
    return i.jsxs("div", {
      className: "lg-grid-stage",
      onPointerDownCapture: (event) => {
        formulaEntryActions.handleFormulaPointerDown(event) ||
          gridInteractionActions.activateGridPane(paneId) ||
          (event.preventDefault(), event.stopPropagation());
      },
      onFocusCapture: () => gridInteractionActions.activateGridPane(paneId),
      children: [
        i.jsx("div", {
          className: "lg-grid-scroll",
          id: `${editorIdPrefix}-${paneId}-grid`,
          ref: (gridElement) => {
            (splitPaneSizes.bindings[paneId](gridElement),
              splitPaneSizes.active.current === paneId &&
                (selectionState.gridContainerRef.current = gridElement));
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
              _t(selectionState.activeCellAddress).row,
            ) &&
            [...frozenColumnIndexes, ...shownColumns].includes(
              _t(selectionState.activeCellAddress).col,
            )
              ? `${editorIdPrefix}-${paneId}-${activeCellLabel}`
              : undefined,
          "aria-label": sheetSplitPanes
            ? `Worksheet pane ${paneId.replace("-", " ")}`
            : text("Editable worksheet grid"),
          onScroll: (event) => {
            viewAndSelectionActions.applyPaneScrollOffsets(
              paneId,
              event.currentTarget.scrollTop,
              event.currentTarget.scrollLeft,
            );
          },
          onKeyDown: (keyboardEvent) => {
            if (
              sheetSplitPanes &&
              keyboardEvent.key === "F6" &&
              keyboardEvent.target === keyboardEvent.currentTarget &&
              !keyboardEvent.defaultPrevented &&
              !keyboardEvent.ctrlKey &&
              !keyboardEvent.metaKey &&
              !keyboardEvent.altKey
            ) {
              keyboardEvent.preventDefault();
              const paneIds = Pv(sheetSplitPanes),
                currentPaneIndex = paneIds.indexOf(paneId),
                nextPaneId =
                  paneIds[
                    (currentPaneIndex +
                      (keyboardEvent.shiftKey ? paneIds.length - 1 : 1)) %
                      paneIds.length
                  ];
              gridInteractionActions.activateGridPane(nextPaneId) &&
                splitPaneSizes.refs.current[nextPaneId]?.focus({
                  preventScroll: true,
                });
            } else gridInteractionActions.handleGridKeyDown(keyboardEvent);
          },
          onContextMenu: (event) => {
            const target = event.target;
            if (
              target.closest("input,textarea,button,a,[contenteditable=true]")
            )
              return;
            const cellElement = target.closest("[data-address]"),
              columnHeaderElement = target.closest("[data-column]"),
              rowHeaderElement = target.closest("[data-row-header]");
            if (
              !(!cellElement && !columnHeaderElement && !rowHeaderElement) &&
              cellEditActions.commitCellEdit()
            ) {
              if ((event.preventDefault(), columnHeaderElement)) {
                const columnIndex = Number(columnHeaderElement.dataset.column);
                selectionAreaRanges.some(
                  (selectionAreaRange) =>
                    selectionAreaRange.firstRow === 0 &&
                    selectionAreaRange.lastRow === activeSheet.rowCount - 1 &&
                    columnIndex >= selectionAreaRange.firstCol &&
                    columnIndex <= selectionAreaRange.lastCol,
                ) ||
                  navigationActions.selectRowOrColumn(
                    "column",
                    columnIndex,
                    false,
                  );
              } else if (rowHeaderElement) {
                const rowIndex = Number(rowHeaderElement.dataset.rowHeader);
                selectionAreaRanges.some(
                  (selectionAreaRange) =>
                    selectionAreaRange.firstCol === 0 &&
                    selectionAreaRange.lastCol === activeSheet.colCount - 1 &&
                    rowIndex >= selectionAreaRange.firstRow &&
                    rowIndex <= selectionAreaRange.lastRow,
                ) ||
                  navigationActions.selectRowOrColumn("row", rowIndex, false);
              } else if (cellElement) {
                const clickedCell = _t(cellElement.dataset.address);
                selectionAreaRanges.some((selectionAreaRange) =>
                  Tp(selectionAreaRange, clickedCell.row, clickedCell.col),
                ) ||
                  navigationActions.selectCellAddress(
                    cellElement.dataset.address,
                  );
              }
              selectionState.setContextMenuState({
                x: event.clientX,
                y: event.clientY,
                kind: columnHeaderElement
                  ? "column"
                  : rowHeaderElement
                    ? "row"
                    : "cell",
              });
            }
          },
          onCopy: (clipboardEvent) => {
            clipboardEvent.target !== clipboardEvent.currentTarget ||
              clipboardEvent.defaultPrevented ||
              (clipboardEvent.preventDefault(),
              viewAndSelectionActions.runGuardedAction(() =>
                clipboardEvent.clipboardData.setData(
                  "text/plain",
                  clipboardAndFindActions.selectionToTsv(),
                ),
              ));
          },
          onCut: (clipboardEvent) => {
            clipboardEvent.target !== clipboardEvent.currentTarget ||
              isEditingLocked ||
              operationStatusState.isOperationBusy ||
              clipboardEvent.defaultPrevented ||
              (clipboardEvent.preventDefault(),
              viewAndSelectionActions.runGuardedAction(() => {
                (clipboardEvent.clipboardData.setData(
                  "text/plain",
                  clipboardAndFindActions.selectionToTsv(true),
                ),
                  operationStatusState.setStatusNotice(
                    "Cut is ready. Select a destination and paste to preview the move.",
                  ));
              }));
          },
          onPaste: (clipboardEvent) => {
            clipboardEvent.target !== clipboardEvent.currentTarget ||
              isEditingLocked ||
              operationStatusState.isOperationBusy ||
              clipboardEvent.defaultPrevented ||
              (clipboardEvent.preventDefault(),
              clipboardAndFindActions.pasteClipboardText(
                clipboardEvent.clipboardData.getData("text/plain"),
              ));
          },
          onPointerUp: () => {
            dragFillState.dragAnchorRef.current = null;
          },
          onPointerLeave: (pointerEvent) => {
            pointerEvent.buttons === 0 &&
              (dragFillState.dragAnchorRef.current = null);
          },
          children: i.jsxs("table", {
            className: `lg-grid ${viewDisplayState.isGridlinesVisible ? "" : "no-gridlines"} ${viewDisplayState.isHeadingsVisible ? "" : "no-headings"}`,
            style: {
              zoom: zoomScale,
              width: columnWindow.total + frozenColumnsWidth + rowHeaderWidth,
              minWidth: 0,
            },
            role: "presentation",
            children: [
              i.jsxs("colgroup", {
                children: [
                  viewDisplayState.isHeadingsVisible &&
                    i.jsx("col", { style: { width: rowHeaderWidth } }),
                  frozenColumnIndexes.map((frozenColumnIndex) =>
                    i.jsx(
                      "col",
                      {
                        style: {
                          width: columnAxisLayout.size(
                            columnVisibilityIndex.viewIndex(frozenColumnIndex),
                          ),
                        },
                      },
                      frozenColumnIndex,
                    ),
                  ),
                  leftGap > 0 && i.jsx("col", { style: { width: leftGap } }),
                  shownColumns.map((shownColumn) =>
                    i.jsx(
                      "col",
                      {
                        style: {
                          width: columnAxisLayout.size(
                            columnVisibilityIndex.viewIndex(shownColumn),
                          ),
                        },
                      },
                      shownColumn,
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
                    viewDisplayState.isHeadingsVisible &&
                      i.jsx("th", {
                        role: "columnheader",
                        className: "lg-corner",
                        "aria-label": text("Select all cells"),
                        onClick: navigationActions.selectAllCells,
                        lang: language("Select all cells"),
                        "data-ui-label": "Select all cells",
                        children: i.jsx("span", {
                          className: "lg-sr-only",
                          children: i.jsx(LocalizedText, {
                            children: "Select all cells",
                          }),
                        }),
                      }),
                    frozenColumnIndexes.map((frozenColumnIndex) =>
                      renderColumnHeader(frozenColumnIndex, true),
                    ),
                    leftGap > 0 &&
                      i.jsx("th", {
                        className: "lg-column-spacer",
                        "aria-hidden": "true",
                      }),
                    shownColumns.map((shownColumn) =>
                      renderColumnHeader(shownColumn),
                    ),
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
                  frozenRowIndexes.map((frozenRowIndex) =>
                    renderGridRow(frozenRowIndex, true, gridViewport, paneId),
                  ),
                  renderVirtualSpacer(
                    rowWindow.leading,
                    "before",
                    visibleColumnCount,
                  ),
                  shownRows.map((shownRow) =>
                    renderGridRow(shownRow, false, gridViewport, paneId),
                  ),
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
          onWorkbookChange: navigationActions.commitWorkbookEdit,
          positionForCell: (address) => {
            const cell = _t(address);
            if (
              rowVisibilityIndex.isHidden(cell.row) ||
              columnVisibilityIndex.isHidden(cell.col)
            )
              return null;
            const scrollableRowPosition = scrollableRowPositionOf(cell.row),
              isPinnedRow = cell.row < pinnedRowCount,
              isFrozenColumn = cell.col < frozenColumnCount;
            if (!isPinnedRow && scrollableRowPosition === undefined)
              return null;
            const left =
                (rowHeaderWidth +
                  (isFrozenColumn
                    ? columnAxisLayout.offset(
                        columnVisibilityIndex.viewIndex(cell.col),
                      )
                    : frozenColumnsWidth +
                      scrollableColumnLayout.offset(
                        columnVisibilityIndex.viewIndex(cell.col) - yn,
                      ) -
                      columnWindow.logicalOffset)) *
                zoomScale,
              topOffset =
                ((viewDisplayState.isHeadingsVisible ? 22 : 0) +
                  (isPinnedRow
                    ? rowAxisLayout.offset(
                        rowVisibilityIndex.viewIndex(cell.row),
                      )
                    : frozenRowsHeight +
                      scrollableRowLayout.offset(
                        Math.max(0, scrollableRowPosition),
                      ) -
                      rowWindow.logicalOffset)) *
                zoomScale;
            return left < -5e3 ||
              topOffset < -5e3 ||
              left > paneSize.width ||
              topOffset > paneSize.height
              ? null
              : { left: left, top: topOffset };
          },
        }),
      ],
    });
  }
  return { renderGridPane };
}

export { createGridPaneRenderer };
