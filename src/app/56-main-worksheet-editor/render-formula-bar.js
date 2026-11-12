import i from "react/jsx-runtime";
import { Check, X } from "lucide-react";
import { LocalizedText } from "../04-cell-input-paste-and-fill.js";
import {
  FormulaSuggestionList,
  isWithinFormulaSuggestions,
} from "../53-find-replace-go-to.js";

function renderFormulaBar({
  formulaEntryState,
  navigationState,
  selectionState,
  text,
  language,
  editorIdPrefix,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  syncFormulaEditorSelection,
  isFormulaEditingLocked,
  handleFormulaKeyDown,
  beginCellEdit,
  focusGridSurface,
  cancelCellEdit,
  commitCellEdit,
  commitNameBoxReference,
  moveSelectionAfterCommit,
}) {
  return i.jsxs("div", {
    className: "lg-formula-bar",
    children: [
      i.jsx("input", {
        className: "lg-name-box",
        "aria-label": text("Selected range"),
        value: navigationState.goToReferenceText,
        onChange: (fe) => navigationState.setGoToReferenceText(fe.target.value),
        onBlur: commitNameBoxReference,
        onKeyDown: (fe) => {
          fe.key === "Enter" &&
            (commitNameBoxReference(),
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
        onClick: cancelCellEdit,
        lang: language("Cancel cell edit"),
        "data-ui-label": "Cancel cell edit",
        children: i.jsx(X, { size: 15 }),
      }),
      i.jsx("button", {
        "aria-label": text("Apply cell value"),
        disabled: isFormulaEditingLocked(),
        onPointerDown: (fe) => fe.preventDefault(),
        onMouseDown: (fe) => fe.preventDefault(),
        onClick: () => {
          commitCellEdit() && focusGridSurface();
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
        readOnly: isFormulaEditingLocked(),
        value: formulaEntryState.cellEditText,
        "aria-describedby":
          formulaEntryState.activeCellEditRef.current?.value.startsWith("=")
            ? `${editorIdPrefix}-formula-help`
            : undefined,
        onFocus: (fe) => {
          (!formulaEntryState.activeCellEditRef.current &&
            formulaEntryState.cellEditText.startsWith("=") &&
            !isEditingLocked &&
            beginCellEdit(formulaEntryState.cellEditText),
            formulaEntryState.setIsFormulaBarFocused(true),
            syncFormulaEditorSelection(fe.currentTarget));
        },
        onSelect: (fe) => {
          (syncFormulaEditorSelection(fe.currentTarget),
            formulaEntryState.setFormulaCaretPosition(
              fe.currentTarget.selectionStart ??
                formulaEntryState.cellEditText.length,
            ));
        },
        onChange: (fe) => {
          (beginCellEdit(fe.target.value),
            syncFormulaEditorSelection(fe.currentTarget, true),
            formulaEntryState.setFormulaCaretPosition(
              fe.target.selectionStart ?? fe.target.value.length,
            ));
        },
        onBlur: (fe) => {
          isWithinFormulaSuggestions(fe.currentTarget, fe.relatedTarget) ||
            (formulaEntryState.setIsFormulaBarFocused(false), commitCellEdit());
        },
        onKeyDown: (fe) => {
          if (!fe.nativeEvent.isComposing) {
            if (fe.key === "Escape") {
              if (
                !formulaEntryState.formulaPointerStateRef.current.span &&
                formulaEntryState.cellEditText.length > 1 &&
                formulaEntryState.formulaSuggestionListRef.current?.handleKeyDown(
                  fe,
                )
              )
                return;
              (fe.preventDefault(), cancelCellEdit());
              return;
            }
            handleFormulaKeyDown(fe) ||
              formulaEntryState.formulaSuggestionListRef.current?.handleKeyDown(
                fe,
              ) ||
              ((fe.key === "Enter" || fe.key === "Tab") &&
                (fe.preventDefault(),
                commitCellEdit() &&
                  moveSelectionAfterCommit(fe.key, fe.shiftKey)));
          }
        },
        lang: language("Formula bar"),
        "data-ui-label": "Formula bar",
      }),
      i.jsx(FormulaSuggestionList, {
        ref: formulaEntryState.formulaSuggestionListRef,
        onExit: () => {
          (formulaEntryState.setIsFormulaBarFocused(false), commitCellEdit());
        },
        value: formulaEntryState.cellEditText,
        caret: formulaEntryState.formulaCaretPosition,
        workbook: activeWorkbook,
        sheetId: activeSheet.id,
        inputRef: formulaEntryState.formulaInputRef,
        disabled:
          isFormulaEditingLocked() || !formulaEntryState.isFormulaBarFocused,
        onAccept: (fe, Oe) => {
          (beginCellEdit(fe),
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

export { renderFormulaBar };
