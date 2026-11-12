import i from "react/jsx-runtime";
import { Layers, Plus } from "lucide-react";

function renderSheetTabs({
  dialogState,
  selectionState,
  navigationState,
  formulaEntryState,
  readOnly,
  text,
  language,
  activeWorkbook,
  activeSheet,
  runEditorCommand,
  commitCellEdit,
  switchActiveWorksheet,
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
            .filter((fe) => !fe.hidden)
            .map((fe) =>
              i.jsx(
                "button",
                {
                  "aria-current": fe.id === activeSheet.id ? "true" : undefined,
                  className: fe.id === activeSheet.id ? "active" : "",
                  "data-sheet-id": fe.id,
                  tabIndex: fe.id === activeSheet.id ? 0 : -1,
                  onPointerDown: (Oe) => {
                    formulaEntryState.activeCellEditRef.current?.value.startsWith(
                      "=",
                    ) && Oe.preventDefault();
                  },
                  onMouseDown: (Oe) => {
                    formulaEntryState.activeCellEditRef.current?.value.startsWith(
                      "=",
                    ) && Oe.preventDefault();
                  },
                  onKeyDown: (Oe) => {
                    if (
                      !["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                        Oe.key,
                      ) ||
                      Oe.altKey ||
                      Oe.ctrlKey ||
                      Oe.metaKey
                    )
                      return;
                    Oe.preventDefault();
                    const Be = activeWorkbook.sheets.filter((gt) => !gt.hidden),
                      Ue = Be.findIndex((gt) => gt.id === fe.id),
                      Xe =
                        Oe.key === "Home"
                          ? Be[0]
                          : Oe.key === "End"
                            ? Be.at(-1)
                            : Be[
                                (Ue +
                                  (Oe.key === "ArrowRight"
                                    ? 1
                                    : Be.length - 1)) %
                                  Be.length
                              ];
                    (switchActiveWorksheet(Xe.id),
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
                  onClick: () => switchActiveWorksheet(fe.id),
                  style: fe.tabColor
                    ? {
                        boxShadow: `inset 0 -3px ${fe.tabColor}`,
                      }
                    : undefined,
                  onDoubleClick: () => {
                    commitCellEdit() &&
                      (navigationState.setTabBarSheetId(fe.id),
                      dialogState.setActiveDialogMode("sheets"));
                  },
                  onContextMenu: (Oe) => {
                    (Oe.preventDefault(),
                      commitCellEdit() &&
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
            onClick: () => runEditorCommand("sheets"),
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

export { renderSheetTabs };
