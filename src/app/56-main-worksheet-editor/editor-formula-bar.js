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
        onChange: (event) =>
          navigationState.setGoToReferenceText(event.target.value),
        onBlur: navigationActions.commitNameBoxReference,
        onKeyDown: (event) => {
          event.key === "Enter" &&
            (navigationActions.commitNameBoxReference(),
            selectionState.gridContainerRef.current?.focus());
        },
        lang: language("Selected range"),
        "data-ui-label": "Selected range",
      }),
      i.jsx("button", {
        className: "lg-formula-cancel",
        "aria-label": text("Cancel cell edit"),
        onPointerDown: (event) => event.preventDefault(),
        onMouseDown: (event) => event.preventDefault(),
        onClick: cellEditActions.cancelCellEdit,
        lang: language("Cancel cell edit"),
        "data-ui-label": "Cancel cell edit",
        children: i.jsx(X, { size: 15 }),
      }),
      i.jsx("button", {
        "aria-label": text("Apply cell value"),
        disabled: cellEditActions.isFormulaEditingLocked(),
        onPointerDown: (event) => event.preventDefault(),
        onMouseDown: (event) => event.preventDefault(),
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
        onFocus: (event) => {
          (!formulaEntryState.activeCellEditRef.current &&
            formulaEntryState.cellEditText.startsWith("=") &&
            !isEditingLocked &&
            cellEditActions.beginCellEdit(formulaEntryState.cellEditText),
            formulaEntryState.setIsFormulaBarFocused(true),
            cellEditActions.syncFormulaEditorSelection(event.currentTarget));
        },
        onSelect: (event) => {
          (cellEditActions.syncFormulaEditorSelection(event.currentTarget),
            formulaEntryState.setFormulaCaretPosition(
              event.currentTarget.selectionStart ??
                formulaEntryState.cellEditText.length,
            ));
        },
        onChange: (event) => {
          (cellEditActions.beginCellEdit(event.target.value),
            cellEditActions.syncFormulaEditorSelection(
              event.currentTarget,
              true,
            ),
            formulaEntryState.setFormulaCaretPosition(
              event.target.selectionStart ?? event.target.value.length,
            ));
        },
        onBlur: (event) => {
          isWithinFormulaSuggestions(
            event.currentTarget,
            event.relatedTarget,
          ) ||
            (formulaEntryState.setIsFormulaBarFocused(false),
            cellEditActions.commitCellEdit());
        },
        onKeyDown: (event) => {
          if (!event.nativeEvent.isComposing) {
            if (event.key === "Escape") {
              if (
                !formulaEntryState.formulaPointerStateRef.current.span &&
                formulaEntryState.cellEditText.length > 1 &&
                formulaSuggestionListRef.current?.handleKeyDown(event)
              )
                return;
              (event.preventDefault(), cellEditActions.cancelCellEdit());
              return;
            }
            formulaEntryActions.handleFormulaKeyDown(event) ||
              formulaSuggestionListRef.current?.handleKeyDown(event) ||
              ((event.key === "Enter" || event.key === "Tab") &&
                (event.preventDefault(),
                cellEditActions.commitCellEdit() &&
                  navigationActions.moveSelectionAfterCommit(
                    event.key,
                    event.shiftKey,
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
