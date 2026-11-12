import i from "react/jsx-runtime";
import ne from "react";
import {
  LazyAssistantChat,
  LazyAssistantChatMarkdown,
  Vj,
  workbookInstanceKey,
} from "../55-external-source-refresh.js";
import {
  AssistantAnswerEvidence,
  evaluateAnswerEvidence,
} from "../51-pattern-checks-and-citations.js";
import { LocalizedText } from "../04-cell-input-paste-and-fill.js";
import { Check } from "lucide-react";
import { validateProposedCalculation } from "../20-calculation-options-and-assistant.js";

function EditorAssistantPanel({
  viewAndSelectionActions,
  cellEditActions,
  fileAndAnalysisActions,
  selectionState,
  formulaEntryState,
  dialogState,
  filterSortState,
  navigationState,
  operationStatusState,
  assistantState,
  adapter,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  selectionAreasKey,
  filterStateKey,
  activeFilterState,
  latestEditStateRef,
  selectedCellCount,
  assistantSelection,
  selectionReferenceLabel,
}) {
  return (
    dialogState.sidePanelKind === "assistant" &&
    i.jsx(ne.Suspense, {
      fallback: i.jsx("p", {
        role: "status",
        className: "lg-muted",
        children: "Loading chat…",
      }),
      children: i.jsx(LazyAssistantChat, {
        session: assistantState.assistantSession,
        setSession: assistantState.setAssistantSession,
        adapter,
        local: adapter === Vj || adapter.localOnly === true,
        scope: JSON.stringify([
          activeWorkbook.id,
          activeWorkbook.revision,
          activeSheet.id,
          selectionAreasKey,
          workbookInstanceKey(adapter),
        ]),
        selectionLabel: selectionReferenceLabel,
        selectedCount: selectedCellCount,
        busy: assistantState.isAssistantBusy,
        proposalId: assistantState.assistantProposal?.id,
        onSend: fileAndAnalysisActions.askAssistantQuestion,
        onCancel: () => {
          (assistantState.assistantRequestIdRef.current++,
            assistantState.assistantAbortControllerRef.current?.abort(),
            assistantState.setIsAssistantBusy(false),
            assistantState.isAssistantBusy &&
              operationStatusState.setStatusNotice(
                "Assistant request cancelled. No workbook changes were made.",
              ));
        },
        onReset: () => {
          (assistantState.setAssistantProposal(null),
            assistantState.setAnswerCitations(null),
            operationStatusState.setErrorNotice(""));
        },
        review:
          assistantState.assistantProposal && assistantState.answerCitations
            ? i.jsxs(i.Fragment, {
                children: [
                  assistantState.assistantProposal &&
                    assistantState.answerCitations &&
                    !evaluateAnswerEvidence(
                      activeWorkbook,
                      assistantSelection,
                      assistantState.answerCitations,
                    ).current &&
                    i.jsx("p", {
                      role: "status",
                      children: i.jsx(LocalizedText, {
                        children:
                          "The answer evidence is stale or unavailable. Request a fresh answer before review.",
                      }),
                    }),
                  assistantState.assistantProposal &&
                    assistantState.answerCitations &&
                    evaluateAnswerEvidence(
                      activeWorkbook,
                      assistantSelection,
                      assistantState.answerCitations,
                    ).current &&
                    i.jsxs("div", {
                      className: "lg-proposal",
                      children: [
                        i.jsx("span", {
                          className: "lg-eyebrow",
                          children: assistantState.assistantProposal.patches
                            .length
                            ? "REVIEW PROPOSED CHANGES"
                            : "RESPONSE",
                        }),
                        i.jsx(
                          AssistantAnswerEvidence,
                          {
                            workbook: activeWorkbook,
                            selection: assistantSelection,
                            evidence: assistantState.answerCitations,
                            renderText: (markdownText) =>
                              i.jsx(LazyAssistantChatMarkdown, {
                                children: markdownText,
                              }),
                            onNavigate: (sheetId, cellAddress) => {
                              const latestEdit = latestEditStateRef.current;
                              if (formulaEntryState.activeCellEditRef.current)
                                throw Error(
                                  "Finish the current cell edit and request fresh evidence before following a citation.",
                                );
                              if (
                                sheetId !== latestEdit.sheetId ||
                                !evaluateAnswerEvidence(
                                  latestEdit.book,
                                  {
                                    sheetId: latestEdit.sheetId,
                                    range: latestEdit.range,
                                  },
                                  assistantState.answerCitations,
                                ).current
                              )
                                throw Error(
                                  "The current approved source changed. Request a fresh answer.",
                                );
                              ((filterSortState.filterValueSearch ||
                                filterSortState.activeColumnSort ||
                                activeFilterState) &&
                                (filterSortState.setFilterStateByScope(
                                  (previousFilterState) => {
                                    const nextFilterState = {
                                      ...previousFilterState,
                                    };
                                    return (
                                      delete nextFilterState[filterStateKey],
                                      nextFilterState
                                    );
                                  },
                                ),
                                filterSortState.setFilterValueSearch(""),
                                filterSortState.setActiveColumnSort(null),
                                operationStatusState.setStatusNotice(
                                  "The row view was cleared to reveal the cited cell. The approved shared range is unchanged.",
                                )),
                                selectionState.setActiveCellAddress(
                                  cellAddress,
                                ),
                                navigationState.setPendingScrollAddress(
                                  cellAddress,
                                ),
                                viewAndSelectionActions.focusGridSurface());
                            },
                          },
                          assistantState.assistantProposal.id,
                        ),
                        assistantState.assistantProposal.patches.map((patch) =>
                          i.jsxs(
                            "div",
                            {
                              className: "lg-diff",
                              children: [
                                i.jsx("b", {
                                  children: patch.address,
                                }),
                                i.jsx("del", {
                                  children: String(
                                    activeSheet.cells[patch.address]?.raw ??
                                      "(blank)",
                                  ),
                                }),
                                i.jsxs("span", {
                                  children: [
                                    String(patch.cell.raw ?? "(blank)"),
                                    patch.cell.style?.fill
                                      ? " · highlight"
                                      : "",
                                  ],
                                }),
                                i.jsx("dl", {
                                  className: "lg-proposal-fields",
                                  "aria-label":
                                    patch.address + " proposed field changes",
                                  children: [
                                    [
                                      "Value type",
                                      activeSheet.cells[patch.address]?.raw ==
                                      null
                                        ? "blank"
                                        : typeof activeSheet.cells[
                                            patch.address
                                          ]?.raw,
                                      patch.cell.raw === null
                                        ? "blank"
                                        : typeof patch.cell.raw,
                                    ],
                                    ...Array.from(
                                      new Set([
                                        ...Object.keys(
                                          activeSheet.cells[patch.address] ??
                                            {},
                                        ),
                                        ...Object.keys(patch.cell),
                                      ]),
                                    )
                                      .filter(
                                        (propertyName) =>
                                          propertyName !== "raw",
                                      )
                                      .map((cellField) => [
                                        cellField,
                                        activeSheet.cells[patch.address]?.[
                                          cellField
                                        ],
                                        patch.cell[cellField],
                                      ]),
                                  ]
                                    .filter(
                                      ([, currentValue, proposedValue]) =>
                                        JSON.stringify(currentValue) !==
                                        JSON.stringify(proposedValue),
                                    )
                                    .map(
                                      ([
                                        fieldName,
                                        currentFieldValue,
                                        proposedFieldValue,
                                      ]) =>
                                        i.jsxs(
                                          ne.Fragment,
                                          {
                                            children: [
                                              i.jsx("dt", {
                                                children: String(fieldName),
                                              }),
                                              i.jsxs("dd", {
                                                children: [
                                                  i.jsx("del", {
                                                    children:
                                                      currentFieldValue ===
                                                      undefined
                                                        ? "(not set)"
                                                        : JSON.stringify(
                                                            currentFieldValue,
                                                          ),
                                                  }),
                                                  i.jsx("ins", {
                                                    children:
                                                      proposedFieldValue ===
                                                      undefined
                                                        ? "(not set)"
                                                        : JSON.stringify(
                                                            proposedFieldValue,
                                                          ),
                                                  }),
                                                ],
                                              }),
                                            ],
                                          },
                                          String(fieldName),
                                        ),
                                    ),
                                }),
                              ],
                            },
                            patch.address,
                          ),
                        ),
                        assistantState.assistantProposal.patches.length > 0 &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("button", {
                                className: "lg-button dark full",
                                disabled: isEditingLocked,
                                onClick: () =>
                                  viewAndSelectionActions.runGuardedAction(
                                    () => {
                                      const currentEdit =
                                          latestEditStateRef.current,
                                        assistantSelection = {
                                          sheetId: currentEdit.sheetId,
                                          range: currentEdit.range,
                                        };
                                      if (
                                        currentEdit.readOnly ||
                                        formulaEntryState.activeCellEditRef
                                          .current ||
                                        !evaluateAnswerEvidence(
                                          currentEdit.book,
                                          assistantSelection,
                                          assistantState.answerCitations,
                                        ).current
                                      )
                                        throw Error(
                                          "The current approved source changed or editing is unavailable. Request a fresh answer before applying changes.",
                                        );
                                      const validatedProposal =
                                        validateProposedCalculation(
                                          currentEdit.book,
                                          assistantSelection,
                                          assistantState.assistantProposal,
                                        );
                                      (cellEditActions.applyCellPatches(
                                        validatedProposal.patches,
                                        `Accepted assistant proposal ${validatedProposal.id}`,
                                        currentEdit.selectionKey,
                                      ),
                                        operationStatusState.setStatusNotice(
                                          "Reviewed proposal applied and recorded in activity.",
                                        ));
                                    },
                                  ),
                                children: [
                                  i.jsx(Check, {
                                    size: 15,
                                  }),
                                  i.jsx(LocalizedText, {
                                    children: "Apply ",
                                  }),
                                  assistantState.assistantProposal.patches
                                    .length,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children: "changes",
                                  }),
                                ],
                              }),
                              i.jsx("button", {
                                className: "lg-button full",
                                onClick: () => {
                                  (assistantState.setAssistantProposal(null),
                                    assistantState.setAnswerCitations(null));
                                },
                                children: i.jsx(LocalizedText, {
                                  children: "Discard proposal",
                                }),
                              }),
                            ],
                          }),
                      ],
                    }),
                ],
              })
            : null,
      }),
    })
  );
}

export { EditorAssistantPanel };
