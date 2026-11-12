import i from "react/jsx-runtime";
import { Layers, Plus } from "lucide-react";

function EditorSheetTabs({
  gridInteractionActions,
  cellEditActions,
  navigationActions,
  dialogState,
  selectionState,
  navigationState,
  formulaEntryState,
  readOnly,
  text,
  language,
  activeWorkbook,
  activeSheet,
}) {
  return i.jsxs("div", {
    className: "lg-sheet-tabs",
    "aria-label": text("Worksheet navigation"),
    role: "group",
    lang: language("Worksheet navigation"),
    "data-ui-label": "Worksheet navigation",
    children: [
      i.jsxs("div", {
        children: [
          activeWorkbook.sheets
            .filter((sheet) => !sheet.hidden)
            .map((fe) =>
              i.jsx(
                "button",
                {
                  "aria-current": fe.id === activeSheet.id ? "true" : undefined,
                  className: fe.id === activeSheet.id ? "active" : "",
                  "data-sheet-id": fe.id,
                  tabIndex: fe.id === activeSheet.id ? 0 : -1,
                  onPointerDown: (event) => {
                    formulaEntryState.activeCellEditRef.current?.value.startsWith(
                      "=",
                    ) && event.preventDefault();
                  },
                  onMouseDown: (event) => {
                    formulaEntryState.activeCellEditRef.current?.value.startsWith(
                      "=",
                    ) && event.preventDefault();
                  },
                  onKeyDown: (event) => {
                    if (
                      !["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                        event.key,
                      ) ||
                      event.altKey ||
                      event.ctrlKey ||
                      event.metaKey
                    )
                      return;
                    event.preventDefault();
                    const Be = activeWorkbook.sheets.filter(
                        (sheet) => !sheet.hidden,
                      ),
                      Ue = Be.findIndex((gt) => gt.id === fe.id),
                      Xe =
                        event.key === "Home"
                          ? Be[0]
                          : event.key === "End"
                            ? Be.at(-1)
                            : Be[
                                (Ue +
                                  (event.key === "ArrowRight"
                                    ? 1
                                    : Be.length - 1)) %
                                  Be.length
                              ];
                    (navigationActions.switchActiveWorksheet(Xe.id),
                      requestAnimationFrame(() =>
                        Array.from(
                          selectionState.appRootRef.current?.querySelectorAll(
                            "[data-sheet-id]",
                          ) ?? [],
                        )
                          .find((wt) => wt.dataset.sheetId === Xe.id)
                          ?.focus(),
                      ));
                  },
                  onClick: () => navigationActions.switchActiveWorksheet(fe.id),
                  style: fe.tabColor
                    ? {
                        boxShadow: `inset 0 -3px ${fe.tabColor}`,
                      }
                    : undefined,
                  onDoubleClick: () => {
                    cellEditActions.commitCellEdit() &&
                      (navigationState.setTabBarSheetId(fe.id),
                      dialogState.setActiveDialogMode("sheets"));
                  },
                  onContextMenu: (event) => {
                    (event.preventDefault(),
                      cellEditActions.commitCellEdit() &&
                        (navigationState.setTabBarSheetId(fe.id),
                        dialogState.setActiveDialogMode("sheets")));
                  },
                  children: fe.name,
                },
                fe.id,
              ),
            ),
          i.jsx("button", {
            className: "lg-add-sheet",
            "aria-label": text("Add worksheet"),
            disabled: readOnly || !!activeWorkbook.structureProtection?.enabled,
            onClick: () =>
              dialogState.setPromptDialogState({
                kind: "sheet",
                value: `Sheet ${activeWorkbook.sheets.length + 1}`,
              }),
            lang: language("Add worksheet"),
            "data-ui-label": "Add worksheet",
            children: i.jsx(Plus, { size: 16 }),
          }),
          i.jsx("button", {
            "aria-label": text("Manage worksheets"),
            title: text("Manage worksheets"),
            onClick: () => gridInteractionActions.runEditorCommand("sheets"),
            lang: language("Manage worksheets"),
            "data-ui-label": "Manage worksheets",
            children: i.jsx(Layers, { size: 16 }),
          }),
        ],
      }),
      i.jsx("span", {
        className: "lg-book-name",
        title: text("Workbook name"),
        lang: language("Workbook name"),
        children: activeWorkbook.name,
      }),
    ],
  });
}

export { EditorSheetTabs };
