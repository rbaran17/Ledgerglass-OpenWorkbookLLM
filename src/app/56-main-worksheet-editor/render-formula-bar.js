import i from "react/jsx-runtime";
import { Check, X } from "lucide-react";
import { LocalizedText } from "../04-cell-input-paste-and-fill.js";
import {
  FormulaSuggestionList,
  isWithinFormulaSuggestions,
} from "../53-find-replace-go-to.js";

function renderFormulaBar({
  text,
  language,
  editorIdPrefix,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  formulaInputRef,
  formulaSuggestionListRef,
  formulaCaretPosition,
  setFormulaCaretPosition,
  isFormulaBarFocused,
  setIsFormulaBarFocused,
  goToReferenceText,
  setGoToReferenceText,
  cellEditText,
  activeCellEditRef,
  formulaPointerStateRef,
  gridContainerRef,
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
        value: goToReferenceText,
        onChange: (fe) => setGoToReferenceText(fe.target.value),
        onBlur: commitNameBoxReference,
        onKeyDown: (fe) => {
          fe.key === "Enter" &&
            (commitNameBoxReference(), gridContainerRef.current?.focus());
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
        ref: formulaInputRef,
        "aria-label": text("Formula bar"),
        readOnly: isFormulaEditingLocked(),
        value: cellEditText,
        "aria-describedby": activeCellEditRef.current?.value.startsWith("=")
          ? `${editorIdPrefix}-formula-help`
          : undefined,
        onFocus: (fe) => {
          (!activeCellEditRef.current &&
            cellEditText.startsWith("=") &&
            !isEditingLocked &&
            beginCellEdit(cellEditText),
            setIsFormulaBarFocused(true),
            syncFormulaEditorSelection(fe.currentTarget));
        },
        onSelect: (fe) => {
          (syncFormulaEditorSelection(fe.currentTarget),
            setFormulaCaretPosition(
              fe.currentTarget.selectionStart ?? cellEditText.length,
            ));
        },
        onChange: (fe) => {
          (beginCellEdit(fe.target.value),
            syncFormulaEditorSelection(fe.currentTarget, true),
            setFormulaCaretPosition(
              fe.target.selectionStart ?? fe.target.value.length,
            ));
        },
        onBlur: (fe) => {
          isWithinFormulaSuggestions(fe.currentTarget, fe.relatedTarget) ||
            (setIsFormulaBarFocused(false), commitCellEdit());
        },
        onKeyDown: (fe) => {
          if (!fe.nativeEvent.isComposing) {
            if (fe.key === "Escape") {
              if (
                !formulaPointerStateRef.current.span &&
                cellEditText.length > 1 &&
                formulaSuggestionListRef.current?.handleKeyDown(fe)
              )
                return;
              (fe.preventDefault(), cancelCellEdit());
              return;
            }
            handleFormulaKeyDown(fe) ||
              formulaSuggestionListRef.current?.handleKeyDown(fe) ||
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
        ref: formulaSuggestionListRef,
        onExit: () => {
          (setIsFormulaBarFocused(false), commitCellEdit());
        },
        value: cellEditText,
        caret: formulaCaretPosition,
        workbook: activeWorkbook,
        sheetId: activeSheet.id,
        inputRef: formulaInputRef,
        disabled: isFormulaEditingLocked() || !isFormulaBarFocused,
        onAccept: (fe, Oe) => {
          (beginCellEdit(fe),
            setFormulaCaretPosition(Oe),
            requestAnimationFrame(() => {
              (formulaInputRef.current?.focus(),
                formulaInputRef.current?.setSelectionRange(Oe, Oe));
            }));
        },
      }),
    ],
  });
}

export { renderFormulaBar };
