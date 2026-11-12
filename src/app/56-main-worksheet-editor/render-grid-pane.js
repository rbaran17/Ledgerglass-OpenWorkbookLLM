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
  function renderGridPane(fe) {
    const Oe = _t(selectionState.activeCellAddress),
      Be =
        filterSortState.filterValueSearch || filterSortState.activeColumnSort
          ? undefined
          : mergeRanges.find((mergeRange) => Tp(mergeRange, Oe.row, Oe.col)),
      Ue = Be
        ? `${columnIndexToLetters(Be.firstCol)}${Be.firstRow + 1}`
        : selectionState.activeCellAddress,
      Xe = splitPaneSizes.views[fe],
      gt = viewAndSelectionActions.computeGridViewport(Xe),
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
      onPointerDownCapture: (event) => {
        formulaEntryActions.handleFormulaPointerDown(event) ||
          gridInteractionActions.activateGridPane(fe) ||
          (event.preventDefault(), event.stopPropagation());
      },
      onFocusCapture: () => gridInteractionActions.activateGridPane(fe),
      children: [
        i.jsx("div", {
          className: "lg-grid-scroll",
          id: `${editorIdPrefix}-${fe}-grid`,
          ref: (Sr) => {
            (splitPaneSizes.bindings[fe](Sr),
              splitPaneSizes.active.current === fe &&
                (selectionState.gridContainerRef.current = Sr));
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
              ? `${editorIdPrefix}-${fe}-${Ue}`
              : undefined,
          "aria-label": sheetSplitPanes
            ? `Worksheet pane ${fe.replace("-", " ")}`
            : text("Editable worksheet grid"),
          onScroll: (event) => {
            viewAndSelectionActions.applyPaneScrollOffsets(
              fe,
              event.currentTarget.scrollTop,
              event.currentTarget.scrollLeft,
            );
          },
          onKeyDown: (Sr) => {
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
              gridInteractionActions.activateGridPane(Ds) &&
                splitPaneSizes.refs.current[Ds]?.focus({ preventScroll: true });
            } else gridInteractionActions.handleGridKeyDown(Sr);
          },
          onContextMenu: (event) => {
            const target = event.target;
            if (
              target.closest("input,textarea,button,a,[contenteditable=true]")
            )
              return;
            const Di = target.closest("[data-address]"),
              $o = target.closest("[data-column]"),
              Ds = target.closest("[data-row-header]");
            if (!(!Di && !$o && !Ds) && cellEditActions.commitCellEdit()) {
              if ((event.preventDefault(), $o)) {
                const qo = Number($o.dataset.column);
                selectionAreaRanges.some(
                  (selectionAreaRange) =>
                    selectionAreaRange.firstRow === 0 &&
                    selectionAreaRange.lastRow === activeSheet.rowCount - 1 &&
                    qo >= selectionAreaRange.firstCol &&
                    qo <= selectionAreaRange.lastCol,
                ) || navigationActions.selectRowOrColumn("column", qo, false);
              } else if (Ds) {
                const qo = Number(Ds.dataset.rowHeader);
                selectionAreaRanges.some(
                  (selectionAreaRange) =>
                    selectionAreaRange.firstCol === 0 &&
                    selectionAreaRange.lastCol === activeSheet.colCount - 1 &&
                    qo >= selectionAreaRange.firstRow &&
                    qo <= selectionAreaRange.lastRow,
                ) || navigationActions.selectRowOrColumn("row", qo, false);
              } else if (Di) {
                const qo = _t(Di.dataset.address);
                selectionAreaRanges.some((selectionAreaRange) =>
                  Tp(selectionAreaRange, qo.row, qo.col),
                ) || navigationActions.selectCellAddress(Di.dataset.address);
              }
              selectionState.setContextMenuState({
                x: event.clientX,
                y: event.clientY,
                kind: $o ? "column" : Ds ? "row" : "cell",
              });
            }
          },
          onCopy: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(),
              viewAndSelectionActions.runGuardedAction(() =>
                Sr.clipboardData.setData(
                  "text/plain",
                  clipboardAndFindActions.selectionToTsv(),
                ),
              ));
          },
          onCut: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              isEditingLocked ||
              operationStatusState.isOperationBusy ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(),
              viewAndSelectionActions.runGuardedAction(() => {
                (Sr.clipboardData.setData(
                  "text/plain",
                  clipboardAndFindActions.selectionToTsv(true),
                ),
                  operationStatusState.setStatusNotice(
                    "Cut is ready. Select a destination and paste to preview the move.",
                  ));
              }));
          },
          onPaste: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              isEditingLocked ||
              operationStatusState.isOperationBusy ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(),
              clipboardAndFindActions.pasteClipboardText(
                Sr.clipboardData.getData("text/plain"),
              ));
          },
          onPointerUp: () => {
            dragFillState.dragAnchorRef.current = null;
          },
          onPointerLeave: (Sr) => {
            Sr.buttons === 0 && (dragFillState.dragAnchorRef.current = null);
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
                    renderGridRow(frozenRowIndex, true, gt, fe),
                  ),
                  renderVirtualSpacer(
                    rowWindow.leading,
                    "before",
                    visibleColumnCount,
                  ),
                  shownRows.map((shownRow) =>
                    renderGridRow(shownRow, false, gt, fe),
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
            const left =
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
                ((viewDisplayState.isHeadingsVisible ? 22 : 0) +
                  ($o
                    ? rowAxisLayout.offset(rowVisibilityIndex.viewIndex(Wi.row))
                    : frozenRowsHeight +
                      scrollableRowLayout.offset(Math.max(0, Di)) -
                      rowWindow.logicalOffset)) *
                zoomScale;
            return left < -5e3 || yo < -5e3 || left > Xe.width || yo > Xe.height
              ? null
              : { left: left, top: yo };
          },
        }),
      ],
    });
  }
  return { renderGridPane };
}

export { createGridPaneRenderer };
