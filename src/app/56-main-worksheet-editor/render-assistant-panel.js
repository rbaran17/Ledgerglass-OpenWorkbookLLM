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
import { validateProposedCalculation } from "../20-calculation-options-and-sources.js";

function renderAssistantPanel({
  adapter,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  setActiveCellAddress,
  selectionAreasKey,
  activeCellEditRef,
  sidePanelKind,
  filterValueSearch,
  setFilterValueSearch,
  setFilterStateByScope,
  filterStateKey,
  activeFilterState,
  activeColumnSort,
  setActiveColumnSort,
  setPendingScrollAddress,
  setStatusNotice,
  setErrorNotice,
  assistantSession,
  setAssistantSession,
  assistantProposal,
  setAssistantProposal,
  answerCitations,
  setAnswerCitations,
  isAssistantBusy,
  setIsAssistantBusy,
  assistantAbortControllerRef,
  assistantRequestIdRef,
  latestEditStateRef,
  selectedCellCount,
  assistantSelection,
  selectionReferenceLabel,
  runGuardedAction,
  applyCellPatches,
  focusGridSurface,
  askAssistantQuestion,
}) {
  return (
    sidePanelKind === "assistant" &&
    i.jsx(ne.Suspense, {
      fallback: i.jsx("p", {
        role: "status",
        className: "lg-muted",
        children: "Loading chat…",
      }),
      children: i.jsx(LazyAssistantChat, {
        session: assistantSession,
        setSession: setAssistantSession,
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
        busy: isAssistantBusy,
        proposalId: assistantProposal?.id,
        onSend: askAssistantQuestion,
        onCancel: () => {
          var fe;
          (assistantRequestIdRef.current++,
            (fe = assistantAbortControllerRef.current) == null || fe.abort(),
            setIsAssistantBusy(false),
            isAssistantBusy &&
              setStatusNotice(
                "Assistant request cancelled. No workbook changes were made.",
              ));
        },
        onReset: () => {
          (setAssistantProposal(null),
            setAnswerCitations(null),
            setErrorNotice(""));
        },
        review:
          assistantProposal && answerCitations
            ? i.jsxs(i.Fragment, {
                children: [
                  assistantProposal &&
                    answerCitations &&
                    !evaluateAnswerEvidence(
                      activeWorkbook,
                      assistantSelection,
                      answerCitations,
                    ).current &&
                    i.jsx("p", {
                      role: "status",
                      children: i.jsx(LocalizedText, {
                        children:
                          "The answer evidence is stale or unavailable. Request a fresh answer before review.",
                      }),
                    }),
                  assistantProposal &&
                    answerCitations &&
                    evaluateAnswerEvidence(
                      activeWorkbook,
                      assistantSelection,
                      answerCitations,
                    ).current &&
                    i.jsxs("div", {
                      className: "lg-proposal",
                      children: [
                        i.jsx("span", {
                          className: "lg-eyebrow",
                          children: assistantProposal.patches.length
                            ? "REVIEW PROPOSED CHANGES"
                            : "RESPONSE",
                        }),
                        i.jsx(
                          AssistantAnswerEvidence,
                          {
                            workbook: activeWorkbook,
                            selection: assistantSelection,
                            evidence: answerCitations,
                            renderText: (fe) =>
                              i.jsx(LazyAssistantChatMarkdown, {
                                children: fe,
                              }),
                            onNavigate: (fe, Oe) => {
                              const Be = latestEditStateRef.current;
                              if (activeCellEditRef.current)
                                throw Error(
                                  "Finish the current cell edit and request fresh evidence before following a citation.",
                                );
                              if (
                                fe !== Be.sheetId ||
                                !evaluateAnswerEvidence(
                                  Be.book,
                                  {
                                    sheetId: Be.sheetId,
                                    range: Be.range,
                                  },
                                  answerCitations,
                                ).current
                              )
                                throw Error(
                                  "The current approved source changed. Request a fresh answer.",
                                );
                              ((filterValueSearch ||
                                activeColumnSort ||
                                activeFilterState) &&
                                (setFilterStateByScope((Ue) => {
                                  const Xe = { ...Ue };
                                  return (delete Xe[filterStateKey], Xe);
                                }),
                                setFilterValueSearch(""),
                                setActiveColumnSort(null),
                                setStatusNotice(
                                  "The row view was cleared to reveal the cited cell. The approved shared range is unchanged.",
                                )),
                                setActiveCellAddress(Oe),
                                setPendingScrollAddress(Oe),
                                focusGridSurface());
                            },
                          },
                          assistantProposal.id,
                        ),
                        assistantProposal.patches.map((fe) => {
                          var Be;
                          return i.jsxs(
                            "div",
                            {
                              className: "lg-diff",
                              children: [
                                i.jsx("b", {
                                  children: fe.address,
                                }),
                                i.jsx("del", {
                                  children: String(
                                    activeSheet.cells[fe.address]?.raw ??
                                      "(blank)",
                                  ),
                                }),
                                i.jsxs("span", {
                                  children: [
                                    String(fe.cell.raw ?? "(blank)"),
                                    (Be = fe.cell.style) != null && Be.fill
                                      ? " · highlight"
                                      : "",
                                  ],
                                }),
                                i.jsx("dl", {
                                  className: "lg-proposal-fields",
                                  "aria-label":
                                    fe.address + " proposed field changes",
                                  children: [
                                    [
                                      "Value type",
                                      activeSheet.cells[fe.address]?.raw == null
                                        ? "blank"
                                        : typeof activeSheet.cells[fe.address]
                                            ?.raw,
                                      fe.cell.raw === null
                                        ? "blank"
                                        : typeof fe.cell.raw,
                                    ],
                                    ...Array.from(
                                      new Set([
                                        ...Object.keys(
                                          activeSheet.cells[fe.address] ?? {},
                                        ),
                                        ...Object.keys(fe.cell),
                                      ]),
                                    )
                                      .filter((gt) => gt !== "raw")
                                      .map((gt) => {
                                        return [
                                          gt,
                                          activeSheet.cells[fe.address]?.[gt],
                                          fe.cell[gt],
                                        ];
                                      }),
                                  ]
                                    .filter(
                                      ([, gt, Nt]) =>
                                        JSON.stringify(gt) !==
                                        JSON.stringify(Nt),
                                    )
                                    .map(([gt, Nt, wt]) =>
                                      i.jsxs(
                                        ne.Fragment,
                                        {
                                          children: [
                                            i.jsx("dt", {
                                              children: String(gt),
                                            }),
                                            i.jsxs("dd", {
                                              children: [
                                                i.jsx("del", {
                                                  children:
                                                    Nt === undefined
                                                      ? "(not set)"
                                                      : JSON.stringify(Nt),
                                                }),
                                                i.jsx("ins", {
                                                  children:
                                                    wt === undefined
                                                      ? "(not set)"
                                                      : JSON.stringify(wt),
                                                }),
                                              ],
                                            }),
                                          ],
                                        },
                                        String(gt),
                                      ),
                                    ),
                                }),
                              ],
                            },
                            fe.address,
                          );
                        }),
                        assistantProposal.patches.length > 0 &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("button", {
                                className: "lg-button dark full",
                                disabled: isEditingLocked,
                                onClick: () =>
                                  runGuardedAction(() => {
                                    const fe = latestEditStateRef.current,
                                      Oe = {
                                        sheetId: fe.sheetId,
                                        range: fe.range,
                                      };
                                    if (
                                      fe.readOnly ||
                                      activeCellEditRef.current ||
                                      !evaluateAnswerEvidence(
                                        fe.book,
                                        Oe,
                                        answerCitations,
                                      ).current
                                    )
                                      throw Error(
                                        "The current approved source changed or editing is unavailable. Request a fresh answer before applying changes.",
                                      );
                                    const Be = validateProposedCalculation(
                                      fe.book,
                                      Oe,
                                      assistantProposal,
                                    );
                                    (applyCellPatches(
                                      Be.patches,
                                      `Accepted assistant proposal ${Be.id}`,
                                      fe.selectionKey,
                                    ),
                                      setStatusNotice(
                                        "Reviewed proposal applied and recorded in activity.",
                                      ));
                                  }),
                                children: [
                                  i.jsx(Check, {
                                    size: 15,
                                  }),
                                  i.jsx(LocalizedText, {
                                    children: "Apply ",
                                  }),
                                  assistantProposal.patches.length,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children: "changes",
                                  }),
                                ],
                              }),
                              i.jsx("button", {
                                className: "lg-button full",
                                onClick: () => {
                                  (setAssistantProposal(null),
                                    setAnswerCitations(null));
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

export { renderAssistantPanel };
