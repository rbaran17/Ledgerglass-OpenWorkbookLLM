import i from "react/jsx-runtime";
import {
  LocalizedNumberField,
  LocalizedText,
} from "../04-cell-input-paste-and-fill.js";
import {
  ArrowUpRight,
  ChartColumn,
  ClipboardCheck,
  Clock3,
  Copy,
  Download,
  FlaskConical,
  ListFilter,
  ShieldCheck,
} from "lucide-react";
import { columnIndexToLetters } from "../03-workbook-model.js";
import { saveBlobToFile } from "../13-consolidation-dialog.js";

function EditorChecksPanel({
  gridInteractionActions,
  navigationActions,
  fileAndAnalysisActions,
  dialogState,
  filterSortState,
  auditState,
  operationStatusState,
  text,
  language,
  activeWorkbook,
  activeSheet,
  filterResult,
}) {
  return (
    dialogState.sidePanelKind === "checks" &&
    i.jsxs(i.Fragment, {
      children: [
        i.jsx("span", {
          className: "lg-eyebrow",
          children: i.jsx(LocalizedText, {
            children: "AUDIT",
          }),
        }),
        i.jsx("h3", {
          children: i.jsx(LocalizedText, {
            children: "Audit checks",
          }),
        }),
        i.jsxs("label", {
          children: [
            i.jsx(LocalizedText, {
              children: "Analysis population",
            }),
            i.jsxs("select", {
              "aria-label": text("Analysis population"),
              value: auditState.auditPopulationScope,
              onChange: (event) => {
                (auditState.setAuditPopulationScope(event.target.value),
                  auditState.setAuditAnalysisResult(null));
              },
              lang: language("Analysis population"),
              "data-ui-label": "Analysis population",
              children: [
                i.jsx("option", {
                  value: "full",
                  children:
                    "Full selected source range (including hidden/filtered)",
                }),
                i.jsx("option", {
                  value: "visible",
                  disabled: !filterResult,
                  children: "Visible filter population",
                }),
              ],
            }),
          ],
        }),
        i.jsxs("p", {
          className: "lg-muted",
          children: [
            i.jsx(LocalizedText, {
              children:
                "Applies to the five checks below. Pattern checks and other tool dialogs use their own complete source ranges and explicit source filters.",
            }),
            " ",
            auditState.auditPopulationScope === "visible" && filterResult
              ? `Using ${filterResult.context.range}, ${filterResult.context.visibleRows} visible / ${filterResult.context.totalRows} total rows. The custom check range is ignored.`
              : "Full scope includes hidden and filtered rows.",
          ],
        }),
        i.jsx("p", {
          className: "lg-muted",
          children: i.jsx(LocalizedText, {
            children:
              "Run a reproducible check, then trace each finding directly to its source.",
          }),
        }),
        i.jsxs("div", {
          className: "lg-test-choices",
          children: [
            i.jsxs("button", {
              onClick: () =>
                gridInteractionActions.runEditorCommand("patterns"),
              children: [
                i.jsx(ListFilter, { size: 17 }),
                i.jsxs("span", {
                  children: [
                    i.jsx("strong", {
                      children: i.jsx(LocalizedText, {
                        children: "Pattern checks",
                      }),
                    }),
                    i.jsx("small", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Duplicates, sequence gaps and unusual postings",
                      }),
                    }),
                  ],
                }),
              ],
            }),
            [
              {
                kind: "duplicates",
                title: "Duplicate entries",
                icon: Copy,
                detail: "Find repeated transaction IDs",
              },
              {
                kind: "materiality",
                title: "Materiality scan",
                icon: ListFilter,
                detail: "Surface amounts above a threshold",
              },
              {
                kind: "sample",
                title: "Select a sample",
                icon: FlaskConical,
                detail: "Repeatable selection with a seed",
              },
              {
                kind: "summary",
                title: "Account summary",
                icon: ChartColumn,
                detail: "Group and total the population",
              },
              {
                kind: "aging",
                title: "Aging analysis",
                icon: Clock3,
                detail: "Bucket transactions by age",
              },
            ].map((fe) =>
              i.jsxs(
                "button",
                {
                  onClick: () => {
                    if (fe.kind === "materiality" || fe.kind === "aging") {
                      gridInteractionActions.runEditorCommand(fe.kind);
                      return;
                    }
                    (auditState.setSelectedAuditKind(fe.kind),
                      auditState.setAuditAnalysisResult(null));
                  },
                  className:
                    auditState.selectedAuditKind === fe.kind ? "active" : "",
                  children: [
                    i.jsx(fe.icon, { size: 17 }),
                    i.jsxs("span", {
                      children: [
                        i.jsx("strong", {
                          children: fe.title,
                        }),
                        i.jsx("small", {
                          children: fe.detail,
                        }),
                      ],
                    }),
                    i.jsx("span", {
                      className: "lg-radio",
                      children:
                        auditState.selectedAuditKind === fe.kind &&
                        i.jsx("span", {}),
                    }),
                  ],
                },
                fe.kind,
              ),
            ),
          ],
        }),
        i.jsxs("label", {
          className: "lg-field",
          children: [
            i.jsx(LocalizedText, {
              children: "Source range",
            }),
            " ",
            i.jsx("input", {
              "aria-label": text("Audit source range"),
              placeholder: `A1:${columnIndexToLetters(activeSheet.colCount - 1)}${activeSheet.rowCount} (entire sheet)`,
              value: auditState.auditSourceRange,
              onChange: (event) =>
                auditState.setAuditSourceRange(event.target.value),
              lang: language("Audit source range"),
              "data-ui-label": "Audit source range",
            }),
          ],
        }),
        i.jsxs("span", {
          className: "lg-field-hint",
          children: [
            i.jsx(LocalizedText, {
              children: "First row contains field names.",
            }),
            " ",
            auditState.auditPopulationScope === "visible"
              ? "Uses the visible rows of the applied filter range."
              : "Includes all source rows.",
          ],
        }),
        auditState.selectedAuditKind === "materiality" &&
          i.jsxs("label", {
            className: "lg-field",
            children: [
              i.jsx(LocalizedText, {
                children: "Absolute amount threshold",
              }),
              i.jsx(LocalizedNumberField, {
                "aria-label": text("Materiality threshold"),
                min: "0",
                value: auditState.materialityThreshold,
                onValueChange: (fe) =>
                  auditState.setMaterialityThreshold(Number(fe)),
                lang: language("Materiality threshold"),
                "data-ui-label": "Materiality threshold",
              }),
            ],
          }),
        auditState.selectedAuditKind === "sample" &&
          i.jsxs("div", {
            className: "lg-two-fields",
            children: [
              i.jsxs("label", {
                className: "lg-field",
                children: [
                  i.jsx(LocalizedText, {
                    children: "Sample size",
                  }),
                  i.jsx(LocalizedNumberField, {
                    "aria-label": text("Sample size"),
                    min: "1",
                    value: auditState.auditSampleSize,
                    onValueChange: (fe) =>
                      auditState.setAuditSampleSize(Number(fe)),
                    lang: language("Sample size"),
                    "data-ui-label": "Sample size",
                  }),
                ],
              }),
              i.jsxs("label", {
                className: "lg-field",
                children: [
                  i.jsx(LocalizedText, {
                    children: "Seed",
                  }),
                  i.jsx("input", {
                    "aria-label": text("Sampling seed"),
                    value: auditState.samplingSeed,
                    onChange: (event) =>
                      auditState.setSamplingSeed(event.target.value),
                    lang: language("Sampling seed"),
                    "data-ui-label": "Sampling seed",
                  }),
                ],
              }),
            ],
          }),
        auditState.selectedAuditKind === "aging" &&
          i.jsxs("label", {
            className: "lg-field",
            children: [
              i.jsx(LocalizedText, {
                children: "As of date",
              }),
              i.jsx("input", {
                "aria-label": text("Aging as of date"),
                type: "date",
                value: auditState.agingAsOfDate,
                onChange: (event) =>
                  auditState.setAgingAsOfDate(event.target.value),
                lang: language("Aging as of date"),
                "data-ui-label": "Aging as of date",
              }),
            ],
          }),
        i.jsxs("button", {
          className: "lg-button dark full",
          onClick: fileAndAnalysisActions.runCurrentAuditAnalysis,
          children: [
            i.jsx(ClipboardCheck, { size: 15 }),
            i.jsx(LocalizedText, {
              children: "Run check",
            }),
            i.jsx(ArrowUpRight, { size: 15 }),
          ],
        }),
        auditState.auditAnalysisResult
          ? i.jsxs("div", {
              className: "lg-results",
              children: [
                i.jsxs("div", {
                  className: "lg-results-heading",
                  children: [
                    i.jsx("h4", {
                      children: auditState.auditAnalysisResult.title,
                    }),
                    i.jsx("span", {
                      children: auditState.auditAnalysisResult.findings.length,
                    }),
                  ],
                }),
                i.jsx("p", {
                  children: auditState.auditAnalysisResult.summary,
                }),
                i.jsxs("button", {
                  className: "lg-button full",
                  onClick: () =>
                    saveBlobToFile(
                      new Blob(
                        [
                          JSON.stringify(
                            auditState.agingReportPayload,
                            null,
                            2,
                          ),
                        ],
                        {
                          type: "application/json",
                        },
                      ),
                      `Audit-${auditState.selectedAuditKind}-revision-${activeWorkbook.revision}.json`,
                    ),
                  children: [
                    i.jsx(Download, { size: 13 }),
                    i.jsx(LocalizedText, {
                      children: "Export evidence report",
                    }),
                  ],
                }),
                auditState.auditAnalysisResult.findings.map((finding) =>
                  i.jsxs(
                    "article",
                    {
                      className: `lg-finding ${finding.severity}`,
                      children: [
                        i.jsx("h4", {
                          children: finding.title,
                        }),
                        i.jsx("p", {
                          children: finding.detail,
                        }),
                        i.jsx("div", {
                          className: "lg-citations",
                          children: finding.citations.map((citation, Be) =>
                            i.jsxs(
                              "button",
                              {
                                onClick: () => {
                                  (filterSortState.setFilterStateByScope(
                                    (Ue) => {
                                      const Xe = {
                                        ...Ue,
                                      };
                                      return (
                                        delete Xe[
                                          JSON.stringify([
                                            activeWorkbook.id,
                                            citation.sheetId,
                                          ])
                                        ],
                                        Xe
                                      );
                                    },
                                  ),
                                    navigationActions.switchActiveWorksheet(
                                      citation.sheetId,
                                    ),
                                    navigationActions.selectCellAddress(
                                      citation.address,
                                    ),
                                    auditState.setAuditAnalysisResult(
                                      auditState.auditAnalysisResult,
                                    ),
                                    operationStatusState.setStatusNotice(
                                      `${citation.sheetName}!${citation.address}: ${String(citation.value ?? "")}`,
                                    ));
                                },
                                children: [
                                  citation.sheetName,
                                  "!",
                                  citation.address,
                                  i.jsx(ArrowUpRight, {
                                    size: 10,
                                  }),
                                ],
                              },
                              `${citation.address}-${Be}`,
                            ),
                          ),
                        }),
                      ],
                    },
                    finding.id,
                  ),
                ),
              ],
            })
          : i.jsxs("div", {
              className: "lg-info-card",
              children: [
                i.jsx(ShieldCheck, { size: 18 }),
                i.jsxs("div", {
                  children: [
                    i.jsx("strong", {
                      children: i.jsx(LocalizedText, {
                        children: "Evidence stays connected",
                      }),
                    }),
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Every finding includes a worksheet and cell reference. Checks leave source data unchanged.",
                      }),
                    }),
                  ],
                }),
              ],
            }),
      ],
    })
  );
}

export { EditorChecksPanel };
