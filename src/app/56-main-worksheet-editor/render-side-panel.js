import i from "react/jsx-runtime";
import {
  CircleHelp,
  ClipboardCheck,
  FileSpreadsheet,
  History,
  PanelRightClose,
  Sparkles,
} from "lucide-react";
import { renderChecksPanel } from "./render-checks-panel.js";
import { renderAssistantPanel } from "./render-assistant-panel.js";
import { LocalizedText } from "../04-cell-input-paste-and-fill.js";
import { WorkbookCapacityLimits } from "../20-calculation-options-and-assistant.js";

function renderSidePanel({
  adapter,
  text,
  language,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  setActiveCellAddress,
  selectionAreasKey,
  activeCellEditRef,
  sidePanelKind,
  setSidePanelKind,
  filterValueSearch,
  setFilterValueSearch,
  setFilterStateByScope,
  auditPopulationScope,
  setAuditPopulationScope,
  filterStateKey,
  activeFilterState,
  activeColumnSort,
  setActiveColumnSort,
  setPendingScrollAddress,
  setStatusNotice,
  setErrorNotice,
  selectedAuditKind,
  setSelectedAuditKind,
  auditAnalysisResult,
  setAuditAnalysisResult,
  materialityThreshold,
  setMaterialityThreshold,
  auditSampleSize,
  setAuditSampleSize,
  samplingSeed,
  setSamplingSeed,
  agingReportPayload,
  agingAsOfDate,
  setAgingAsOfDate,
  auditSourceRange,
  setAuditSourceRange,
  assistantSession,
  setAssistantSession,
  assistantProposal,
  setAssistantProposal,
  answerCitations,
  setAnswerCitations,
  isAssistantBusy,
  setIsAssistantBusy,
  filterResult,
  assistantAbortControllerRef,
  assistantRequestIdRef,
  latestEditStateRef,
  selectedCellCount,
  assistantSelection,
  selectionReferenceLabel,
  runEditorCommand,
  runGuardedAction,
  applyCellPatches,
  focusGridSurface,
  selectCellAddress,
  switchActiveWorksheet,
  runCurrentAuditAnalysis,
  askAssistantQuestion,
}) {
  return (
    sidePanelKind &&
    i.jsxs("section", {
      "aria-label":
        sidePanelKind === "assistant"
          ? "Workbook assistant panel"
          : "Workbook side panel",
      className: `lg-panel ${sidePanelKind === "assistant" ? "lg-chat-panel" : ""}`,
      children: [
        i.jsxs("div", {
          className: "lg-panel-heading",
          children: [
            i.jsxs("div", {
              children: [
                sidePanelKind === "checks"
                  ? i.jsx(ClipboardCheck, { size: 18 })
                  : sidePanelKind === "assistant"
                    ? i.jsx(Sparkles, { size: 18 })
                    : sidePanelKind === "history"
                      ? i.jsx(History, { size: 18 })
                      : i.jsx(CircleHelp, { size: 18 }),
                i.jsx("h2", {
                  children:
                    sidePanelKind === "checks"
                      ? "Audit toolkit"
                      : sidePanelKind === "assistant"
                        ? "Workbook assistant"
                        : sidePanelKind === "history"
                          ? "Activity"
                          : "About this workspace",
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close side panel"),
              onClick: () => setSidePanelKind(null),
              lang: language("Close side panel"),
              "data-ui-label": "Close side panel",
              children: i.jsx(PanelRightClose, { size: 16 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-panel-body",
          children: [
            renderChecksPanel({
              text,
              language,
              activeWorkbook,
              activeSheet,
              sidePanelKind,
              setFilterStateByScope,
              auditPopulationScope,
              setAuditPopulationScope,
              setStatusNotice,
              selectedAuditKind,
              setSelectedAuditKind,
              auditAnalysisResult,
              setAuditAnalysisResult,
              materialityThreshold,
              setMaterialityThreshold,
              auditSampleSize,
              setAuditSampleSize,
              samplingSeed,
              setSamplingSeed,
              agingReportPayload,
              agingAsOfDate,
              setAgingAsOfDate,
              auditSourceRange,
              setAuditSourceRange,
              filterResult,
              runEditorCommand,
              selectCellAddress,
              switchActiveWorksheet,
              runCurrentAuditAnalysis,
            }),
            renderAssistantPanel({
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
            }),
            sidePanelKind === "history" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("span", {
                    className: "lg-eyebrow",
                    children: i.jsx(LocalizedText, {
                      children: "WORKBOOK ACTIVITY",
                    }),
                  }),
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Change history",
                    }),
                  }),
                  i.jsx("p", {
                    className: "lg-muted",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Local activity is included in JSON exports. This log is editable with the file and is not an immutable audit record.",
                    }),
                  }),
                  activeWorkbook.events.length === 0
                    ? i.jsxs("div", {
                        className: "lg-empty",
                        children: [
                          i.jsx(History, { size: 26 }),
                          i.jsx("p", {
                            children: i.jsx(LocalizedText, {
                              children:
                                "Your first edit starts the activity trail.",
                            }),
                          }),
                        ],
                      })
                    : i.jsx("div", {
                        className: "lg-timeline",
                        children: [...activeWorkbook.events]
                          .reverse()
                          .map((fe) =>
                            i.jsxs(
                              "article",
                              {
                                children: [
                                  i.jsx("span", {
                                    className: "lg-timeline-dot",
                                  }),
                                  i.jsx("h4", {
                                    children: fe.action,
                                  }),
                                  i.jsxs("p", {
                                    children: [
                                      fe.actor,
                                      " ·",
                                      " ",
                                      new Date(fe.timestamp).toLocaleString(),
                                    ],
                                  }),
                                  fe.changes.slice(0, 5).map((Oe, Be) =>
                                    i.jsxs(
                                      "small",
                                      {
                                        children: [
                                          activeWorkbook.sheets.find(
                                            (Nt) => Nt.id === Oe.sheetId,
                                          )?.name ?? "Sheet",
                                          "!",
                                          Oe.address,
                                          ":",
                                          " ",
                                          String(Oe.before?.raw ?? "∅"),
                                          " →",
                                          " ",
                                          String(Oe.after?.raw ?? "∅"),
                                        ],
                                      },
                                      Be,
                                    ),
                                  ),
                                  fe.changes.length > 5 &&
                                    i.jsxs("small", {
                                      children: [
                                        "+ ",
                                        fe.changes.length - 5,
                                        " ",
                                        i.jsx(LocalizedText, {
                                          children: "more cell changes",
                                        }),
                                      ],
                                    }),
                                ],
                              },
                              fe.id,
                            ),
                          ),
                      }),
                ],
              }),
            sidePanelKind === "help" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("span", {
                    className: "lg-eyebrow",
                    children: i.jsx(LocalizedText, {
                      children: "SPREADSHEET HELP",
                    }),
                  }),
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Workbook help",
                    }),
                  }),
                  i.jsx("p", {
                    className: "lg-muted",
                    children: i.jsx(LocalizedText, {
                      children:
                        "An embeddable React workbook for audit preparation and review.",
                    }),
                  }),
                  i.jsx("h4", {
                    children: i.jsx(LocalizedText, {
                      children: "Working with cells",
                    }),
                  }),
                  i.jsxs("ul", {
                    className: "lg-help-list",
                    children: [
                      i.jsx("li", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Click a cell, then type or use the formula bar. Enter commits the value.",
                        }),
                      }),
                      i.jsx("li", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Enter dates as YYYY-MM-DD using the workbook's selected 1900 or 1904 date system. Other date formats stay text; prefix an ISO date with an apostrophe to keep it as text. The formula bar shows the stored serial and the grid shows the date.",
                        }),
                      }),
                      i.jsx("li", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Ctrl/Command-click to add areas; Shift-click extends a range. Enter A1:D10, F1:F10 in the name box.",
                        }),
                      }),
                      i.jsx("li", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Copy and paste rectangular tab-separated values with Ctrl/⌘ C and V.",
                        }),
                      }),
                      i.jsx("li", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Prefix text with an apostrophe to prevent formula evaluation.",
                        }),
                      }),
                      i.jsx("li", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Sort and search change the view; source row numbers remain fixed.",
                        }),
                      }),
                    ],
                  }),
                  i.jsx("h4", {
                    children: i.jsx(LocalizedText, {
                      children: "Compatibility",
                    }),
                  }),
                  i.jsx("p", {
                    className: "lg-muted",
                    children: i.jsx(LocalizedText, {
                      children:
                        "XLSX retains a documented subset of cells, formulas and formatting. JSON retains the complete native workpaper. Macros, Power Query, DAX, advanced PivotTables, and full Excel layout fidelity are target requirements, not features of this release.",
                    }),
                  }),
                  i.jsx(WorkbookCapacityLimits, {}),
                  i.jsxs("button", {
                    className: "lg-button full",
                    onClick: () => window.print(),
                    children: [
                      i.jsx(FileSpreadsheet, { size: 15 }),
                      i.jsx(LocalizedText, {
                        children: "Print current view",
                      }),
                    ],
                  }),
                ],
              }),
          ],
        }),
      ],
    })
  );
}

export { renderSidePanel };
