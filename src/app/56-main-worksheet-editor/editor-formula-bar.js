import ne from "react";
import i from "react/jsx-runtime";
import { Check, X } from "lucide-react";
import { LocalizedText } from "../04-cell-input-paste-and-fill.js";
import {
  FormulaSuggestionList,
  isWithinFormulaSuggestions,
} from "../53-find-replace-go-to.js";

function EditorFormulaBar({
  cellEditActions,
  formulaEntryActions,
  viewAndSelectionActions,
  navigationActions,
  formulaEntryState,
  navigationState,
  selectionState,
  text,
  language,
  editorIdPrefix,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
}) {
  // The suggestion list belongs to the formula bar alone, so its ref lives here.
  const formulaSuggestionListRef = ne.useRef(null);
  return i.jsxs("div", {
    className: "lg-formula-bar",
    children: [
      i.jsx("input", {
        className: "lg-name-box",
        "aria-label": text("Selected range"),
        value: navigationState.goToReferenceText,
        onChange: (fe) => navigationState.setGoToReferenceText(fe.target.value),
        onBlur: navigationActions.commitNameBoxReference,
        onKeyDown: (fe) => {
          fe.key === "Enter" &&
            (navigationActions.commitNameBoxReference(),
            selectionState.gridContainerRef.current?.focus());
        },
        lang: language("Selected range"),
        "data-ui-label": "Selected range",
      }),
      i.jsx("button", {
        className: "lg-formula-cancel",
        "aria-label": text("Cancel cell edit"),
        onPointerDown: (fe) => fe.preventDefault(),
        onMouseDown: (fe) => fe.preventDefault(),
        onClick: cellEditActions.cancelCellEdit,
        lang: language("Cancel cell edit"),
        "data-ui-label": "Cancel cell edit",
        children: i.jsx(X, { size: 15 }),
      }),
      i.jsx("button", {
        "aria-label": text("Apply cell value"),
        disabled: cellEditActions.isFormulaEditingLocked(),
        onPointerDown: (fe) => fe.preventDefault(),
        onMouseDown: (fe) => fe.preventDefault(),
        onClick: () => {
          cellEditActions.commitCellEdit() &&
            viewAndSelectionActions.focusGridSurface();
        },
        lang: language("Apply cell value"),
        "data-ui-label": "Apply cell value",
        children: i.jsx(Check, { size: 15 }),
      }),
      i.jsx("span", {
        className: "lg-fx",
        children: i.jsx(LocalizedText, { children: "ƒx" }),
      }),
      i.jsx("input", {
        ref: formulaEntryState.formulaInputRef,
        "aria-label": text("Formula bar"),
        readOnly: cellEditActions.isFormulaEditingLocked(),
        value: formulaEntryState.cellEditText,
        "aria-describedby":
          formulaEntryState.activeCellEditRef.current?.value.startsWith("=")
            ? `${editorIdPrefix}-formula-help`
            : undefined,
        onFocus: (fe) => {
          (!formulaEntryState.activeCellEditRef.current &&
            formulaEntryState.cellEditText.startsWith("=") &&
            !isEditingLocked &&
            cellEditActions.beginCellEdit(formulaEntryState.cellEditText),
            formulaEntryState.setIsFormulaBarFocused(true),
            cellEditActions.syncFormulaEditorSelection(fe.currentTarget));
        },
        onSelect: (fe) => {
          (cellEditActions.syncFormulaEditorSelection(fe.currentTarget),
            formulaEntryState.setFormulaCaretPosition(
              fe.currentTarget.selectionStart ??
                formulaEntryState.cellEditText.length,
            ));
        },
        onChange: (fe) => {
          (cellEditActions.beginCellEdit(fe.target.value),
            cellEditActions.syncFormulaEditorSelection(fe.currentTarget, true),
            formulaEntryState.setFormulaCaretPosition(
              fe.target.selectionStart ?? fe.target.value.length,
            ));
        },
        onBlur: (fe) => {
          isWithinFormulaSuggestions(fe.currentTarget, fe.relatedTarget) ||
            (formulaEntryState.setIsFormulaBarFocused(false),
            cellEditActions.commitCellEdit());
        },
        onKeyDown: (fe) => {
          if (!fe.nativeEvent.isComposing) {
            if (fe.key === "Escape") {
              if (
                !formulaEntryState.formulaPointerStateRef.current.span &&
                formulaEntryState.cellEditText.length > 1 &&
                formulaSuggestionListRef.current?.handleKeyDown(fe)
              )
                return;
              (fe.preventDefault(), cellEditActions.cancelCellEdit());
              return;
            }
            formulaEntryActions.handleFormulaKeyDown(fe) ||
              formulaSuggestionListRef.current?.handleKeyDown(fe) ||
              ((fe.key === "Enter" || fe.key === "Tab") &&
                (fe.preventDefault(),
                cellEditActions.commitCellEdit() &&
                  navigationActions.moveSelectionAfterCommit(
                    fe.key,
                    fe.shiftKey,
                  )));
          }
        },
        lang: language("Formula bar"),
        "data-ui-label": "Formula bar",
      }),
      i.jsx(FormulaSuggestionList, {
        ref: formulaSuggestionListRef,
        onExit: () => {
          (formulaEntryState.setIsFormulaBarFocused(false),
            cellEditActions.commitCellEdit());
        },
        value: formulaEntryState.cellEditText,
        caret: formulaEntryState.formulaCaretPosition,
        workbook: activeWorkbook,
        sheetId: activeSheet.id,
        inputRef: formulaEntryState.formulaInputRef,
        disabled:
          cellEditActions.isFormulaEditingLocked() ||
          !formulaEntryState.isFormulaBarFocused,
        onAccept: (fe, Oe) => {
          (cellEditActions.beginCellEdit(fe),
            formulaEntryState.setFormulaCaretPosition(Oe),
            requestAnimationFrame(() => {
              (formulaEntryState.formulaInputRef.current?.focus(),
                formulaEntryState.formulaInputRef.current?.setSelectionRange(
                  Oe,
                  Oe,
                ));
            }));
        },
      }),
    ],
  });
}

export { EditorFormulaBar };
