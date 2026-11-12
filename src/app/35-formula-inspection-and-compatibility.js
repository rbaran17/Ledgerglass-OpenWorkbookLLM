import { formatCellValueForDisplay } from "./17-recipe-step-editors.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  parseFormulaTree,
  walkFormulaNodes,
  normalizeFunctionName,
  tokenizeFormulaText,
  normalizeBindingName,
} from "./01-schemas-date-functions-and-validation.js";
import {
  rewriteFormulaReferences,
  frozenConversionUnits,
  calculateWorkbookWithSettings,
  _t,
  cn,
  parseRangeAddress,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { inspectFormulaDependencies } from "./12-consolidation-and-dependency-inspection.js";
import {
  nameScanLimits,
  buildFunctionCatalogEntry,
  functionBrowserCategories,
  buildFunctionCatalog,
  buildFunctionBrowserGroups,
} from "./34-function-browser-and-names.js";

function buildFormulaCompatibilityReport(workbook, options = {}) {
  for (const optionName of ["includeHidden", "resolveNames", "includeMetadata"])
    if (
      options[optionName] !== undefined &&
      typeof options[optionName] != "boolean"
    )
      throw Error(`${optionName} must be boolean.`);
  if (
    options.allowedSheetIds !== undefined &&
    (!Array.isArray(options.allowedSheetIds) ||
      options.allowedSheetIds.some(
        (allowedSheetId) =>
          typeof allowedSheetId != "string" ||
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ))
  )
    throw Error("Invalid authorized worksheet scope.");
  const limits = { ...nameScanLimits, ...options.limits };
  for (const limitName of ["maxLocations", "maxCharacters", "maxFindings"])
    if (
      !Number.isInteger(limits[limitName]) ||
      limits[limitName] < 1 ||
      limits[limitName] > nameScanLimits[limitName]
    )
      throw Error(`Invalid ${limitName}; limits can only be narrowed.`);
  const includeHidden = options.includeHidden ?? true,
    resolvesNames = options.resolveNames ?? true,
    includesMetadata = options.includeMetadata ?? true,
    allowedSheetIdSet = new Set(
      options.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
    ),
    visibleSheets = workbook.sheets.filter(
      (sheet) =>
        allowedSheetIdSet.has(sheet.id) && (includeHidden || !sheet.hidden),
    ),
    visibleSheetIds = new Set(visibleSheets.map((sheet) => sheet.id)),
    report = {
      workbookId: workbook.id,
      revision: workbook.revision,
      scope: {
        sheetIds: [...visibleSheetIds],
        includeHidden: includeHidden,
        resolvesNames: resolvesNames,
        includesMetadata: includesMetadata,
      },
      coverage: {
        locationsScanned: 0,
        charactersScanned: 0,
        findings: 0,
        complete: true,
        notes: [
          "Static function inventory only: does not certify argument accuracy, compute downstream cell dependencies, execute formulas or contact services.",
          "Unsupported function inventory is representative; unknown custom names are also reported.",
        ],
      },
      findings: [],
      families: buildFunctionBrowserGroups(),
      connectorPolicy: {
        enabled: false,
        required:
          "CUBE, network and linked-data functions stay disabled. An implemented connector and explicit trusted-host approval are required; workbook data and LLM proposals cannot authorize access.",
      },
      excelVerification: { status: "not-run", version: null },
    };
  if (
    (resolvesNames ||
      report.coverage.notes.push(
        "Named definitions are outside this report scope; unresolved names require host review.",
      ),
    includeHidden ||
      report.coverage.notes.push(
        "Hidden worksheets, rows and columns are excluded by the chosen visibility scope.",
      ),
    !visibleSheetIds.size)
  )
    return (
      report.coverage.notes.push(
        "No worksheets are in the authorized visible scope.",
      ),
      report
    );
  const markIncomplete = (note) => {
    ((report.coverage.complete = false),
      report.coverage.notes.includes(note) || report.coverage.notes.push(note));
  };
  let isScanStopped = false;
  const definedNames = resolvesNames ? (workbook.definedNames ?? []) : [],
    scanFormulaText = (formulaText, sheetId, nameChain = []) => {
      if (
        formulaText.length > 4096 ||
        report.coverage.charactersScanned + formulaText.length >
          limits.maxCharacters
      )
        return (
          markIncomplete(
            "Formula character budget reached; uninspected content remains.",
          ),
          [{ kind: "limit", message: "Formula character budget reached." }]
        );
      report.coverage.charactersScanned += formulaText.length;
      const findings = [];
      rewriteFormulaReferences(
        formulaText,
        workbook,
        sheetId ?? visibleSheets[0]?.id ?? "",
        "A1",
      ) &&
        findings.push({
          kind: "external-reference",
          message:
            "A reference form is external or unsupported. Inspect the original formula in the authorized workbook.",
        });
      try {
        walkFormulaNodes(
          tokenizeFormulaText(formulaText),
          (node, localBindingNames, isFunctionCall, I) => {
            if (
              I ||
              (!node.qualifier &&
                localBindingNames.has(normalizeBindingName(node.name)))
            )
              return;
            const functionName = normalizeFunctionName(node.name);
            if (!isFunctionCall && ["TRUE", "FALSE"].includes(functionName))
              return;
            const functionCatalogEntry =
              buildFunctionCatalogEntry(functionName);
            if (
              isFunctionCall &&
              !node.qualifier &&
              functionCatalogEntry.status === "local"
            )
              return;
            const qualifiedSheet = node.qualifier
              ? workbook.sheets.find(
                  (sheet) =>
                    sheet.name.toLowerCase() === node.qualifier.toLowerCase(),
                )
              : undefined;
            if (
              node.qualifier &&
              (!qualifiedSheet || !visibleSheetIds.has(qualifiedSheet.id))
            ) {
              findings.push({
                kind: "excluded",
                message:
                  "A qualified name refers to a missing or excluded worksheet. Its definition was not inspected.",
              });
              return;
            }
            const scopeSheetId = qualifiedSheet?.id ?? sheetId,
              definedName =
                definedNames.find(
                  (entry) =>
                    entry.scopeSheetId === scopeSheetId &&
                    entry.name.toLowerCase() === node.name.toLowerCase(),
                ) ??
                (node.qualifier
                  ? undefined
                  : definedNames.find(
                      (entry) =>
                        !entry.scopeSheetId &&
                        entry.name.toLowerCase() === node.name.toLowerCase(),
                    ));
            if (definedName) {
              const cycleKey = JSON.stringify([
                definedName.scopeSheetId ?? scopeSheetId,
                definedName.name.toLowerCase(),
              ]);
              if (nameChain.includes(cycleKey)) {
                parseFormulaTree(definedName.refersTo) ||
                  findings.push({
                    kind: "name-cycle",
                    viaName: definedName.name,
                    message:
                      "Circular named definition; inspection stopped at the cycle.",
                  });
                return;
              }
              if (nameChain.length >= nameScanLimits.maxNameDepth) {
                (markIncomplete(
                  "Named-definition depth reached; uninspected definitions remain.",
                ),
                  findings.push({
                    kind: "limit",
                    message:
                      "Named-definition inspection is limited to 32 levels.",
                  }));
                return;
              }
              const nestedFindings = scanFormulaText(
                definedName.refersTo,
                definedName.scopeSheetId ?? scopeSheetId,
                [...nameChain, cycleKey],
              );
              findings.push(
                ...nestedFindings.map((nestedFinding) => ({
                  ...nestedFinding,
                  viaName: definedName.name,
                })),
              );
              return;
            }
            if (!isFunctionCall) {
              findings.push({
                kind: "unresolved-name",
                message:
                  "A name is unresolved or its definition is outside this report scope.",
              });
              return;
            }
            findings.push({
              kind:
                functionCatalogEntry.status === "connector-required"
                  ? "connector-required"
                  : "unsupported-function",
              functionName: functionName,
              family: functionCatalogEntry.family,
              message:
                functionCatalogEntry.status === "connector-required"
                  ? "Disabled: an approved, implemented host connector is required. No data was requested."
                  : "This function has no local implementation. Review its source or add-in contract.",
            });
          },
        );
      } catch {
        (findings.push({
          kind: "syntax",
          message:
            "Formula syntax or lexical bindings could not be inspected. Review the original formula.",
        }),
          (report.coverage.complete = false));
      }
      return [
        ...new Map(
          findings.map((finding) => [JSON.stringify(finding), finding]),
        ).values(),
      ];
    },
    scanLocation = (formulaText, source, location, sheet) => {
      if (isScanStopped) return;
      if (++report.coverage.locationsScanned > limits.maxLocations) {
        (report.coverage.locationsScanned--,
          (isScanStopped = true),
          markIncomplete(
            "Location budget reached; uninspected cells or definitions remain.",
          ));
        return;
      }
      const formulaFindings = scanFormulaText(formulaText, sheet?.id);
      for (const finding of formulaFindings) {
        if (report.findings.length >= limits.maxFindings) {
          ((isScanStopped = true),
            markIncomplete(
              "Finding budget reached; additional findings were not included.",
            ));
          break;
        }
        report.findings.push({
          ...finding,
          source: source,
          location: location,
          ...(sheet ? { sheetId: sheet.id, sheetName: sheet.name } : {}),
          message:
            (source === "literal"
              ? "Stored as literal text; will not execute. "
              : "") + finding.message,
        });
      }
      report.coverage.charactersScanned >= limits.maxCharacters &&
        (isScanStopped = true);
    };
  for (const sheet of visibleSheets) {
    const hiddenRowRanges = cn(sheet, "row"),
      hiddenColumnRanges = cn(sheet, "column");
    for (const [cellAddress, cell] of Object.entries(sheet.cells)) {
      if (isScanStopped) break;
      if (!(typeof cell.raw != "string" || !cell.raw.startsWith("="))) {
        if (!includeHidden) {
          const cellPosition = _t(cellAddress);
          if (
            hiddenRowRanges.some(
              ({ start, end }) =>
                cellPosition.row >= start && cellPosition.row <= end,
            ) ||
            hiddenColumnRanges.some(
              ({ start, end }) =>
                cellPosition.col >= start && cellPosition.col <= end,
            )
          )
            continue;
        }
        scanLocation(
          cell.raw,
          cell.literal ? "literal" : "cell",
          cellAddress,
          sheet,
        );
      }
    }
    if (includesMetadata && includeHidden) {
      for (const conditionalFormat of sheet.conditionalFormats ?? [])
        conditionalFormat.formula &&
          scanLocation(
            conditionalFormat.formula,
            "conditional-format",
            conditionalFormat.range,
            sheet,
          );
      for (const table of sheet.tables ?? [])
        for (const column of table.columns)
          column.calculatedFormula &&
            scanLocation(
              column.calculatedFormula,
              "table-template",
              `${table.name}[${column.name}]`,
              sheet,
            );
    }
  }
  if (includesMetadata && resolvesNames)
    for (const definedName of definedNames)
      (definedName.scopeSheetId &&
        !visibleSheetIds.has(definedName.scopeSheetId)) ||
        scanLocation(
          definedName.refersTo,
          "name",
          definedName.name,
          visibleSheets.find((sheet) => sheet.id === definedName.scopeSheetId),
        );
  return (
    includesMetadata &&
      !includeHidden &&
      report.coverage.notes.push(
        "Conditional and table templates are excluded when hidden content is excluded; named definitions remain in scope.",
      ),
    (report.coverage.findings = report.findings.length),
    report
  );
}
function FormulaCompatibilityDialog(props) {
  const { text, language } = useWorkbookMessages(),
    [activeView, setActiveView] = ne.useState("report"),
    [shouldIncludeHidden, setShouldIncludeHidden] = ne.useState(true),
    [searchText, setSearchText] = ne.useState(""),
    [familyFilter, setFamilyFilter] = ne.useState(""),
    [selectedFunctionName, setSelectedFunctionName] = ne.useState("DSUM"),
    reportResult = ne.useMemo(() => {
      try {
        return {
          report: buildFormulaCompatibilityReport(props.workbook, {
            allowedSheetIds: props.allowedSheetIds,
            includeHidden: shouldIncludeHidden,
          }),
        };
      } catch (error) {
        return {
          error:
            error instanceof Error
              ? error.message
              : "Could not inspect this workbook.",
        };
      }
    }, [props.workbook, props.allowedSheetIds, shouldIncludeHidden]),
    filteredFunctions = ne
      .useMemo(buildFunctionCatalog, [])
      .filter(
        (catalogEntry) =>
          (!familyFilter || catalogEntry.family === familyFilter) &&
          `${catalogEntry.name} ${catalogEntry.outputs}`
            .toLowerCase()
            .includes(searchText.toLowerCase()),
      ),
    selectedFunction =
      filteredFunctions.find(
        (catalogEntry) => catalogEntry.name === selectedFunctionName,
      ) ?? filteredFunctions[0],
    report = reportResult.report,
    onClick = () => {
      if (!report) return;
      if (props.onExport) return props.onExport(report);
      const reportBlob = new Blob([JSON.stringify(report, null, 2)], {
          type: "application/json",
        }),
        reportUrl = URL.createObjectURL(reportBlob),
        downloadLink = document.createElement("a");
      ((downloadLink.href = reportUrl),
        (downloadLink.download = "Formula-compatibility.json"),
        downloadLink.click(),
        setTimeout(() => URL.revokeObjectURL(reportUrl), 1e3));
    };
  return i.jsx("div", {
    className: "formula-tool-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    onKeyDown: (event) => {
      event.key === "Escape" && props.onClose();
    },
    children: i.jsxs("section", {
      className: "formula-tools function-compatibility",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": text("Formula compatibility"),
      lang: language("Formula compatibility"),
      "data-ui-label": "Formula compatibility",
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              children: i.jsx(LocalizedText, {
                children: "Formula compatibility",
              }),
            }),
            i.jsx("button", {
              "aria-label": text("Close compatibility"),
              onClick: props.onClose,
              lang: language("Close compatibility"),
              "data-ui-label": "Close compatibility",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Review support limits and affected locations before relying on calculations. Excel version comparison has not been run.",
          }),
        }),
        i.jsxs("div", {
          role: "tablist",
          "aria-label": text("Compatibility views"),
          lang: language("Compatibility views"),
          "data-ui-label": "Compatibility views",
          children: [
            i.jsx("button", {
              autoFocus: true,
              role: "tab",
              "aria-selected": activeView === "report",
              onClick: () => setActiveView("report"),
              children: i.jsx(LocalizedText, {
                children: "Compatibility report",
              }),
            }),
            i.jsx("button", {
              role: "tab",
              "aria-selected": activeView === "registry",
              onClick: () => setActiveView("registry"),
              children: i.jsx(LocalizedText, { children: "Function registry" }),
            }),
          ],
        }),
        activeView === "report"
          ? i.jsxs("div", {
              role: "tabpanel",
              "aria-label": text("Compatibility report"),
              lang: language("Compatibility report"),
              "data-ui-label": "Compatibility report",
              children: [
                i.jsxs("label", {
                  className: "compatibility-checkbox",
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: shouldIncludeHidden,
                      onChange: (event) =>
                        setShouldIncludeHidden(event.target.checked),
                    }),
                    i.jsx(LocalizedText, {
                      children: "Include hidden worksheets, rows and columns",
                    }),
                  ],
                }),
                reportResult.error &&
                  i.jsx("p", { role: "alert", children: reportResult.error }),
                report &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("p", {
                        role: "status",
                        children: [
                          i.jsx(LocalizedText, { children: "Revision " }),
                          report.revision,
                          " · ",
                          report.coverage.locationsScanned,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "locations inspected ·",
                          }),
                          " ",
                          report.findings.length,
                          " ",
                          i.jsx(LocalizedText, { children: "findings ·" }),
                          " ",
                          report.coverage.complete
                            ? "Scan complete within chosen scope"
                            : "Incomplete scan — review limits below",
                        ],
                      }),
                      i.jsx("div", {
                        className: "compatibility-findings",
                        children: i.jsxs("table", {
                          children: [
                            i.jsx("thead", {
                              children: i.jsxs("tr", {
                                children: [
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Location",
                                    }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Function / source",
                                    }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Finding",
                                    }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Action",
                                    }),
                                  }),
                                ],
                              }),
                            }),
                            i.jsxs("tbody", {
                              children: [
                                report.findings
                                  .slice(0, 200)
                                  .map((finding, findingIndex) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [
                                              finding.sheetName
                                                ? `${finding.sheetName}!`
                                                : "",
                                              finding.location,
                                              i.jsx("small", {
                                                children: finding.source,
                                              }),
                                            ],
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              finding.functionName ??
                                                "Reference or name",
                                              finding.viaName &&
                                                i.jsxs("small", {
                                                  children: [
                                                    i.jsx(LocalizedText, {
                                                      children: "Via ",
                                                    }),
                                                    finding.viaName,
                                                  ],
                                                }),
                                            ],
                                          }),
                                          i.jsx("td", {
                                            children: finding.message,
                                          }),
                                          i.jsx("td", {
                                            children:
                                              (finding.source === "cell" ||
                                                finding.source === "literal") &&
                                              finding.sheetId &&
                                              i.jsx("button", {
                                                onClick: () => {
                                                  (props.onNavigate(
                                                    finding.sheetId,
                                                    finding.location,
                                                  ),
                                                    props.onClose());
                                                },
                                                children: i.jsx(LocalizedText, {
                                                  children: "Go to cell",
                                                }),
                                              }),
                                          }),
                                        ],
                                      },
                                      findingIndex,
                                    ),
                                  ),
                                !report.findings.length &&
                                  i.jsx("tr", {
                                    children: i.jsx("td", {
                                      colSpan: 4,
                                      children: i.jsx(LocalizedText, {
                                        children:
                                          "No unsupported function calls found in the inspected scope. This does not establish calculation accuracy.",
                                      }),
                                    }),
                                  }),
                              ],
                            }),
                          ],
                        }),
                      }),
                      report.findings.length > 200 &&
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, {
                              children:
                                "Showing 200 findings. Export the report for all",
                            }),
                            " ",
                            report.findings.length,
                            ".",
                          ],
                        }),
                      i.jsxs("details", {
                        children: [
                          i.jsx("summary", {
                            children: i.jsx(LocalizedText, {
                              children: "Family inventory and engagement uses",
                            }),
                          }),
                          report.families.map((family) =>
                            i.jsxs(
                              "p",
                              {
                                children: [
                                  i.jsx("strong", { children: family.title }),
                                  " · ",
                                  family.adopted.length,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children: "adopted ·",
                                  }),
                                  " ",
                                  family.unavailable.length,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children: "inventoried as unavailable",
                                  }),
                                  i.jsx("br", {}),
                                  family.demand,
                                ],
                              },
                              family.id,
                            ),
                          ),
                        ],
                      }),
                      i.jsx("p", { children: report.connectorPolicy.required }),
                      i.jsx("ul", {
                        children: report.coverage.notes.map((note) =>
                          i.jsx("li", { children: note }, note),
                        ),
                      }),
                    ],
                  }),
                i.jsxs("footer", {
                  children: [
                    i.jsx("button", {
                      onClick: props.onClose,
                      children: i.jsx(LocalizedText, { children: "Close" }),
                    }),
                    i.jsx("button", {
                      disabled: !report,
                      onClick: onClick,
                      children: i.jsx(LocalizedText, {
                        children: "Export compatibility report",
                      }),
                    }),
                  ],
                }),
              ],
            })
          : i.jsxs("div", {
              role: "tabpanel",
              "aria-label": text("Function registry"),
              lang: language("Function registry"),
              "data-ui-label": "Function registry",
              children: [
                i.jsxs("div", {
                  className: "compatibility-filters",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, {
                          children: "Search function registry",
                        }),
                        i.jsx("input", {
                          value: searchText,
                          onChange: (event) =>
                            setSearchText(event.target.value),
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Function family" }),
                        i.jsxs("select", {
                          value: familyFilter,
                          onChange: (event) =>
                            setFamilyFilter(event.target.value),
                          children: [
                            i.jsx("option", {
                              value: "",
                              children: "All families",
                            }),
                            functionBrowserCategories.map(
                              (functionBrowserCategory) =>
                                i.jsx(
                                  "option",
                                  {
                                    value: functionBrowserCategory.id,
                                    children: functionBrowserCategory.title,
                                  },
                                  functionBrowserCategory.id,
                                ),
                            ),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
                selectedFunction
                  ? i.jsxs("div", {
                      className: "formula-function-layout",
                      children: [
                        i.jsx("select", {
                          "aria-label": text("Function contracts"),
                          size: 12,
                          value: selectedFunction.name,
                          onChange: (event) =>
                            setSelectedFunctionName(event.target.value),
                          lang: language("Function contracts"),
                          "data-ui-label": "Function contracts",
                          children: filteredFunctions.map((catalogEntry) =>
                            i.jsxs(
                              "option",
                              {
                                value: catalogEntry.name,
                                children: [
                                  catalogEntry.name,
                                  " ·",
                                  " ",
                                  catalogEntry.status === "local"
                                    ? "Local"
                                    : catalogEntry.status ===
                                        "connector-required"
                                      ? "Connector required"
                                      : "Unavailable",
                                ],
                              },
                              catalogEntry.name,
                            ),
                          ),
                        }),
                        i.jsxs("div", {
                          "aria-label": text("Selected function contract"),
                          lang: language("Selected function contract"),
                          "data-ui-label": "Selected function contract",
                          children: [
                            i.jsx("h3", { children: selectedFunction.name }),
                            i.jsxs("p", {
                              className: "formula-signature",
                              children: [
                                selectedFunction.name,
                                "(",
                                selectedFunction.inputs,
                                ")",
                              ],
                            }),
                            i.jsxs("dl", {
                              children: [
                                i.jsx("dt", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Output",
                                  }),
                                }),
                                i.jsx("dd", {
                                  children: selectedFunction.outputs,
                                }),
                                i.jsx("dt", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Errors",
                                  }),
                                }),
                                i.jsx("dd", {
                                  children: selectedFunction.errors.join(", "),
                                }),
                                i.jsx("dt", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Numeric tolerance",
                                  }),
                                }),
                                i.jsx("dd", {
                                  children: selectedFunction.numericTolerance
                                    ? `Absolute ${selectedFunction.numericTolerance.absolute}; relative ${selectedFunction.numericTolerance.relative}. ${selectedFunction.numericTolerance.appliesTo}`
                                    : "Not applicable — no local implementation.",
                                }),
                                i.jsx("dt", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Verified Excel version",
                                  }),
                                }),
                                i.jsx("dd", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Not run; no verified version.",
                                  }),
                                }),
                                i.jsx("dt", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Limits",
                                  }),
                                }),
                                i.jsx("dd", {
                                  children: selectedFunction.limits,
                                }),
                              ],
                            }),
                            selectedFunction.name === "CONVERT" &&
                              i.jsxs("details", {
                                children: [
                                  i.jsx("summary", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Supported base unit symbols",
                                    }),
                                  }),
                                  i.jsx("p", {
                                    children: frozenConversionUnits
                                      .map(
                                        (frozenConversionUnit) =>
                                          frozenConversionUnit.symbol,
                                      )
                                      .join(", "),
                                  }),
                                  i.jsx("p", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Decimal prefixes apply only to metric entries. Binary prefixes apply only to bit and byte. Unknown or incompatible units return #N/A.",
                                    }),
                                  }),
                                ],
                              }),
                          ],
                        }),
                      ],
                    })
                  : i.jsx("p", {
                      role: "status",
                      children: i.jsx(LocalizedText, {
                        children: "No matching functions.",
                      }),
                    }),
              ],
            }),
      ],
    }),
  });
}
function describeEvaluatedValue(value) {
  if (value instanceof Error && "error" in value)
    return { kind: "error", text: String(value.error) };
  if (Array.isArray(value)) {
    const gridRows = value.every(Array.isArray) ? value : [value],
      preview = [];
    for (const row of gridRows) {
      for (const cellValue of row) {
        if (preview.length === 8) break;
        preview.push(describeEvaluatedValue(cellValue));
      }
      if (preview.length === 8) break;
    }
    return {
      kind: "array",
      text: `${gridRows.length} × ${gridRows[0]?.length ?? 0} array`,
      rows: gridRows.length,
      columns: gridRows[0]?.length ?? 0,
      preview: preview,
      truncated: gridRows.length * (gridRows[0]?.length ?? 0) > 8,
    };
  }
  if (value && typeof value == "object") {
    if ("kind" in value && value.kind === "formula-closure")
      return { kind: "function", text: "LAMBDA function" };
    if ("ref" in value && value.ref && typeof value.ref == "object") {
      const ref = value.ref,
        formatCellReference = (position) =>
          Number.isInteger(position.row) &&
          Number.isInteger(position.col) &&
          position.col > 0
            ? `${columnIndexToLetters(position.col - 1)}${position.row}`
            : "axis reference";
      return {
        kind: "reference",
        text: `${ref.sheet ? `${ref.sheet}!` : ""}${ref.from && ref.to ? `${formatCellReference(ref.from)}:${formatCellReference(ref.to)}` : formatCellReference(ref)}`,
      };
    }
    return { kind: "unavailable", text: "Internal value" };
  }
  const valueText =
    value == null
      ? "(blank)"
      : typeof value == "boolean"
        ? String(value).toUpperCase()
        : String(value);
  return {
    kind: "value",
    text: valueText.slice(0, 256),
    ...(valueText.length <= 256
      ? { value: value ?? null }
      : { truncated: true }),
  };
}
const validateInspectionLimit = (limit, maximum) => {
  if (limit === undefined) return maximum;
  if (!Number.isInteger(limit) || limit < 1 || limit > maximum)
    throw Error(`Inspection limits must be integers between 1 and ${maximum}.`);
  return limit;
};
function inspectFormulaCalculation(workbook, options) {
  const maxSteps = validateInspectionLimit(options.limits?.maxSteps, 500),
    maxReads = validateInspectionLimit(options.limits?.maxReads, 2e3);
  if (
    options.includeHidden !== undefined &&
    typeof options.includeHidden != "boolean"
  )
    throw Error("includeHidden must be boolean.");
  if (
    options.includeNames !== undefined &&
    typeof options.includeNames != "boolean"
  )
    throw Error("includeNames must be boolean.");
  if (
    options.allowedSheetIds !== undefined &&
    (!Array.isArray(options.allowedSheetIds) ||
      options.allowedSheetIds.some(
        (allowedSheetId) =>
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ))
  )
    throw Error("Invalid authorized worksheet context.");
  const allowedSheetIdSet = new Set(
      options.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
    ),
    position = _t(options.address),
    address = `${columnIndexToLetters(position.col)}${position.row + 1}`,
    allowedRanges = options.allowedRanges?.map((allowedRange) => {
      if (!allowedSheetIdSet.has(allowedRange.sheetId))
        throw Error(
          "Inspection ranges must remain within authorized worksheets.",
        );
      return { ...allowedRange, bounds: parseRangeAddress(allowedRange.range) };
    }),
    sheetContextById = new Map(
      workbook.sheets.map((sheet) => [
        sheet.id,
        { sheet: sheet, rows: cn(sheet, "row"), cols: cn(sheet, "column") },
      ]),
    ),
    isCellVisible = (sheetId, cellAddress) => {
      const sheetContext = sheetContextById.get(sheetId);
      if (!sheetContext || !allowedSheetIdSet.has(sheetId)) return false;
      const { row, col } = _t(cellAddress);
      return row >= sheetContext.sheet.rowCount ||
        col >= sheetContext.sheet.colCount ||
        (allowedRanges &&
          !allowedRanges.some(
            (allowedRange) =>
              allowedRange.sheetId === sheetId &&
              row >= allowedRange.bounds.firstRow &&
              row <= allowedRange.bounds.lastRow &&
              col >= allowedRange.bounds.firstCol &&
              col <= allowedRange.bounds.lastCol,
          ))
        ? false
        : !!options.includeHidden ||
            (!sheetContext.sheet.hidden &&
              !sheetContext.rows.some(
                (hiddenRowRange) =>
                  row >= hiddenRowRange.start && row <= hiddenRowRange.end,
              ) &&
              !sheetContext.cols.some(
                (hiddenColumnRange) =>
                  col >= hiddenColumnRange.start &&
                  col <= hiddenColumnRange.end,
              ));
    };
  if (!isCellVisible(options.sheetId, address))
    throw Error(
      "The selected cell is outside the visible, authorized inspection context.",
    );
  const steps = [],
    issues = [];
  let reads = 0,
    stepsObserved = 0,
    hasExcludedRead = false,
    isReadLimitReached = false;
  const addIssue = (kind, message) => {
      !issues.some(
        (issue) => issue.kind === kind && issue.message === message,
      ) &&
        issues.length < 50 &&
        issues.push({ kind: kind, message: message });
    },
    calculationWorkbook =
      options.includeNames === false
        ? { ...workbook, definedNames: [] }
        : workbook,
    workbookWithSettings = calculateWorkbookWithSettings(
      calculationWorkbook,
      {},
      {
        allowName(sheetId) {
          return !allowedSheetIdSet.has(sheetId) ||
            (!options.includeHidden &&
              sheetContextById.get(sheetId)?.sheet.hidden)
            ? ((hasExcludedRead = true),
              addIssue(
                "excluded",
                "A defined name is outside the visible, authorized context. Its definition was not evaluated and the final result is withheld.",
              ),
              false)
            : true;
        },
        allowRead(sheetId, cellAddress) {
          if (++reads > maxReads)
            return (
              (isReadLimitReached = true),
              addIssue(
                "limit",
                "The calculation read budget was reached; the final result is withheld.",
              ),
              false
            );
          const sheet = sheetContextById.get(sheetId)?.sheet;
          if (!sheet) return true;
          const readPosition = _t(cellAddress);
          return readPosition.row >= sheet.rowCount ||
            readPosition.col >= sheet.colCount ||
            isCellVisible(sheetId, cellAddress)
            ? true
            : ((hasExcludedRead = true),
              addIssue(
                "excluded",
                "A calculation attempted to read outside the visible, authorized context. The source was not read and the final result is withheld.",
              ),
              false);
        },
        step(calculationStep) {
          isCellVisible(calculationStep.sheetId, calculationStep.address) &&
            (stepsObserved++,
            !(steps.length >= maxSteps) &&
              steps.push({
                index: steps.length + 1,
                kind: calculationStep.kind,
                sheetId: calculationStep.sheetId,
                address: calculationStep.address,
                operation: calculationStep.operation.slice(0, 1024),
                inputs: (calculationStep.inputs ?? [])
                  .slice(0, 10)
                  .map(describeEvaluatedValue),
                result: describeEvaluatedValue(calculationStep.result),
              }));
        },
      },
    ),
    cell = workbook.sheets.find((sheet) => sheet.id === options.sheetId).cells[
      address
    ],
    raw = cell?.raw ?? null,
    isFormulaCell =
      typeof raw == "string" && raw.startsWith("=") && !cell?.literal,
    calculatedValue = workbookWithSettings(options.sheetId, address),
    { range, diagnostic } = workbookWithSettings.peekSpill(
      options.sheetId,
      address,
    );
  for (const referenceDiagnostic of workbookWithSettings.getReferenceDiagnostics())
    addIssue("reference", referenceDiagnostic.message);
  for (const namedFormulaDiagnostic of workbookWithSettings.getNamedFormulaDiagnostics())
    addIssue("name", namedFormulaDiagnostic.message);
  for (const financialDiagnostic of workbookWithSettings.getFinancialDiagnostics())
    (financialDiagnostic.status === "error" ||
      (financialDiagnostic.solver &&
        (financialDiagnostic.solver.candidates.length > 1 ||
          (financialDiagnostic.signChanges ?? 0) > 1))) &&
      addIssue(
        "calculation",
        `${financialDiagnostic.functionName}: ${financialDiagnostic.message}`,
      );
  const operationalErrors = workbookWithSettings.getOperationalErrors();
  operationalErrors.length &&
    addIssue(
      "calculation",
      `Local calculation diagnostics: ${operationalErrors.join(", ")}.`,
    );
  const truncated = stepsObserved > maxSteps;
  truncated &&
    addIssue(
      "limit",
      `Only the first ${maxSteps} steps were retained; calculation continued within its own limits.`,
    );
  const isResultWithheld = hasExcludedRead || isReadLimitReached;
  return {
    workbookId: workbook.id,
    dateContext: workbookWithSettings.getDateContext(),
    ...(!isResultWithheld &&
    workbookWithSettings.getFinancialDiagnostics().length
      ? { financial: workbookWithSettings.getFinancialDiagnostics() }
      : {}),
    revision: workbook.revision,
    sheetId: options.sheetId,
    address: address,
    raw: raw,
    result: isResultWithheld ? null : calculatedValue,
    resultStatus: isResultWithheld
      ? "withheld"
      : isFormulaCell
        ? "calculated"
        : "literal",
    steps: steps,
    firstErrorStep: steps.find((step) => step.result.kind === "error")?.index,
    ...(!isResultWithheld && range ? { spillRange: range } : {}),
    ...(!isResultWithheld && diagnostic ? { spillDiagnostic: diagnostic } : {}),
    issues: issues,
    coverage: {
      complete: !truncated && !isResultWithheld && !operationalErrors.length,
      truncated: truncated,
      reads: reads,
      stepsObserved: stepsObserved,
      maxSteps: maxSteps,
      maxReads: maxReads,
    },
    notes: [
      "Steps record one actual local calculation, in execution order. Function internals are a single step; references are shown as references without a second read.",
      "Restart and Previous replay the captured result. Reopen or change the workbook revision to recalculate; volatile values may then change.",
      "Array previews contain at most 8 values; text previews contain at most 256 characters and function inputs at most 10 arguments. Full Excel step-order and native tracer-arrow parity are unverified.",
      "Static precedents include potential branches; this log shows only work performed by the local engine, including spill discovery and cache reads. Host authorization remains the embedding application's responsibility.",
    ],
  };
}
const calculationErrorHints = {
  "#DIV/0!": "Review the denominator and blank inputs.",
  "#REF!":
    "A reference is missing, invalid, or no longer points to an available spill.",
  "#NAME?":
    "Review function spelling, defined names, and the supported function catalog.",
  "#VALUE!": "Review argument types and array dimensions.",
  "#N/A": "No matching value was found or a requested value is unavailable.",
  "#NUM!": "Review numeric domains, convergence, and overflow.",
  "#CYCLE!":
    "A calculation depends on itself. Review the recorded cell reads and static precedents.",
  "#SPILL!":
    "The array cannot occupy its output range. Review the obstruction below.",
  "#BLOCKED!":
    "External execution or a source outside inspection scope was refused.",
  "#LIMIT!": "The bounded local calculator reached a work or size limit.",
  "#CALC!": "Review empty arrays and uninvoked reusable functions.",
};
function EvaluateFormulaDialog(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    [checked, setChecked] = ne.useState(false),
    [stepIndex, setStepIndex] = ne.useState(0),
    [exportError, setExportError] = ne.useState(""),
    inspection = ne.useMemo(() => {
      try {
        return {
          report: inspectFormulaCalculation(props.workbook, {
            sheetId: props.sheetId,
            address: props.address,
            allowedSheetIds: props.allowedSheetIds,
            includeHidden: !!props.allowHidden && checked,
          }),
        };
      } catch (error) {
        return {
          error:
            error instanceof Error
              ? error.message
              : "Calculation could not be inspected.",
        };
      }
    }, [
      props.workbook,
      props.sheetId,
      props.address,
      props.allowedSheetIds,
      props.allowHidden,
      checked,
    ]);
  (ne.useEffect(() => {
    (setStepIndex(0), setExportError(""));
  }, [inspection]),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        dialogRef.current?.querySelector("button")?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  const report = inspection.report,
    currentStepIndex = report
      ? Math.min(stepIndex, report.steps.length - 1)
      : -1,
    currentStep = report?.steps[currentStepIndex],
    formatSheetAddress = (sheetId, address) =>
      `${props.workbook.sheets.find((sheet) => sheet.id === sheetId)?.name ?? "Unavailable sheet"}!${address}`,
    onClick = () => {
      if (report)
        try {
          if (props.onExport) props.onExport(structuredClone(report));
          else {
            const objectUrl = URL.createObjectURL(
                new Blob([JSON.stringify(report, null, 2)], {
                  type: "application/json",
                }),
              ),
              downloadLink = document.createElement("a");
            ((downloadLink.href = objectUrl),
              (downloadLink.download = `Calculation-${report.address}-revision-${report.revision}.json`),
              downloadLink.click(),
              setTimeout(() => URL.revokeObjectURL(objectUrl), 1e3));
          }
          setExportError("");
        } catch (error) {
          setExportError(
            error instanceof Error ? error.message : "Export failed.",
          );
        }
    };
  return i.jsx("div", {
    className: "fi-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      className: "fi-dialog fe-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${titleId}-title`,
      ref: dialogRef,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
              ...dialogRef.current.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),summary",
              ),
            ],
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1);
          event.shiftKey && document.activeElement === firstFocusable
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === lastFocusable &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${titleId}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Evaluate Formula",
                  }),
                }),
                i.jsxs("p", {
                  children: [
                    i.jsx(LocalizedText, {
                      children: "Read-only calculation replay · revision ",
                    }),
                    props.workbook.revision,
                  ],
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close formula evaluation"),
              onClick: props.onClose,
              lang: language("Close formula evaluation"),
              "data-ui-label": "Close formula evaluation",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "fi-controls",
          children: [
            props.allowHidden &&
              i.jsxs("label", {
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => setChecked(event.target.checked),
                  }),
                  i.jsx(LocalizedText, { children: "Include hidden cells" }),
                ],
              }),
            i.jsx("button", {
              disabled: !report,
              onClick: onClick,
              children: i.jsx(LocalizedText, {
                children: "Export calculation steps",
              }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "fe-body",
          children: [
            inspection.error &&
              i.jsx("p", { role: "alert", children: inspection.error }),
            exportError && i.jsx("p", { role: "alert", children: exportError }),
            report &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("section", {
                    className: "fi-target",
                    "aria-label": text("Evaluation target"),
                    lang: language("Evaluation target"),
                    "data-ui-label": "Evaluation target",
                    children: [
                      i.jsx("h3", {
                        children: formatSheetAddress(
                          report.sheetId,
                          report.address,
                        ),
                      }),
                      i.jsx("code", {
                        children: String(report.raw ?? "(blank)"),
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx("strong", {
                            children: i.jsx(LocalizedText, {
                              children: "Final result: ",
                            }),
                          }),
                          report.resultStatus === "withheld"
                            ? "Withheld"
                            : String(report.result ?? "(blank)"),
                          report.resultStatus === "literal"
                            ? " · literal input"
                            : "",
                        ],
                      }),
                      report.spillRange &&
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Spill output: ",
                            }),
                            report.spillRange,
                          ],
                        }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: `fi-coverage ${report.coverage.complete ? "" : "fi-warning"}`,
                    role: "status",
                    children: [
                      report.coverage.complete
                        ? "Calculation captured"
                        : "Partial calculation capture",
                      " ",
                      "· ",
                      report.coverage.reads,
                      " ",
                      i.jsx(LocalizedText, { children: "reads · " }),
                      report.steps.length,
                      " ",
                      i.jsx(LocalizedText, { children: "retained steps" }),
                    ],
                  }),
                  report.issues.length > 0 &&
                    i.jsx("ul", {
                      "aria-label": text("Evaluation diagnostics"),
                      lang: language("Evaluation diagnostics"),
                      "data-ui-label": "Evaluation diagnostics",
                      children: report.issues.map((issue, issueIndex) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              i.jsxs("strong", {
                                children: [issue.kind, ": "],
                              }),
                              issue.message,
                            ],
                          },
                          issueIndex,
                        ),
                      ),
                    }),
                  report.spillDiagnostic &&
                    i.jsx("p", {
                      className: "fi-warning",
                      children: report.spillDiagnostic.message,
                    }),
                  i.jsxs("div", {
                    className: "fi-controls fe-playback",
                    children: [
                      i.jsx("button", {
                        onClick: () => setStepIndex(0),
                        disabled: currentStepIndex <= 0,
                        children: i.jsx(LocalizedText, { children: "Restart" }),
                      }),
                      i.jsx("button", {
                        onClick: () =>
                          setStepIndex((previousIndex) => previousIndex - 1),
                        disabled: currentStepIndex <= 0,
                        children: i.jsx(LocalizedText, {
                          children: "Previous step",
                        }),
                      }),
                      i.jsx("button", {
                        onClick: () =>
                          setStepIndex((currentIndex) => currentIndex + 1),
                        disabled:
                          !report.steps.length ||
                          currentStepIndex >= report.steps.length - 1,
                        children: i.jsx(LocalizedText, {
                          children: "Next step",
                        }),
                      }),
                      i.jsx("button", {
                        onClick: () => setStepIndex(report.steps.length - 1),
                        disabled:
                          !report.steps.length ||
                          currentStepIndex >= report.steps.length - 1,
                        children: i.jsx(LocalizedText, {
                          children: "Last step",
                        }),
                      }),
                      i.jsx("button", {
                        onClick: () => setStepIndex(report.firstErrorStep - 1),
                        disabled: !report.firstErrorStep,
                        children: i.jsx(LocalizedText, {
                          children: "First error",
                        }),
                      }),
                    ],
                  }),
                  currentStep &&
                    i.jsxs("section", {
                      className: "fe-step",
                      "aria-label": text("Current calculation step"),
                      "aria-live": "polite",
                      lang: language("Current calculation step"),
                      "data-ui-label": "Current calculation step",
                      children: [
                        i.jsxs("h3", {
                          children: [
                            i.jsx(LocalizedText, { children: "Step " }),
                            currentStep.index,
                            " ",
                            i.jsx(LocalizedText, { children: "of " }),
                            report.steps.length,
                            " · ",
                            currentStep.kind,
                          ],
                        }),
                        i.jsx("p", {
                          children: i.jsx("button", {
                            onClick: () => {
                              try {
                                (props.onNavigate(
                                  currentStep.sheetId,
                                  currentStep.address,
                                ),
                                  setExportError(""));
                              } catch (error) {
                                setExportError(
                                  error instanceof Error
                                    ? error.message
                                    : "Navigation failed.",
                                );
                              }
                            },
                            children: formatSheetAddress(
                              currentStep.sheetId,
                              currentStep.address,
                            ),
                          }),
                        }),
                        i.jsx("code", { children: currentStep.operation }),
                        currentStep.inputs.length > 0 &&
                          i.jsxs("p", {
                            children: [
                              i.jsx(LocalizedText, { children: "Inputs: " }),
                              currentStep.inputs
                                .map((input) => input.text)
                                .join("; "),
                            ],
                          }),
                        i.jsxs("p", {
                          children: [
                            i.jsx("strong", {
                              children: i.jsx(LocalizedText, {
                                children: "Result: ",
                              }),
                            }),
                            i.jsx("code", {
                              children: currentStep.result.text,
                            }),
                            currentStep.result.truncated &&
                              " (preview truncated)",
                          ],
                        }),
                        currentStep.result.preview &&
                          i.jsxs("p", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Array preview:",
                              }),
                              " ",
                              currentStep.result.preview
                                .map((previewItem) => previewItem.text)
                                .join(", "),
                            ],
                          }),
                        currentStep.result.kind === "error" &&
                          i.jsx("p", {
                            className: "fi-warning",
                            children:
                              calculationErrorHints[currentStep.result.text] ??
                              "Review the source formula and arguments.",
                          }),
                      ],
                    }),
                  i.jsxs("details", {
                    children: [
                      i.jsxs("summary", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Captured steps (",
                          }),
                          report.steps.length,
                          ")",
                        ],
                      }),
                      i.jsx("ol", {
                        className: "fe-log",
                        children: report.steps.map((step) =>
                          i.jsx(
                            "li",
                            {
                              children: i.jsxs("button", {
                                "aria-current":
                                  step.index === currentStep?.index
                                    ? "step"
                                    : undefined,
                                onClick: () => setStepIndex(step.index - 1),
                                children: [
                                  step.index,
                                  ". ",
                                  formatSheetAddress(
                                    step.sheetId,
                                    step.address,
                                  ),
                                  " · ",
                                  step.operation,
                                  " ",
                                  "→ ",
                                  step.result.text,
                                ],
                              }),
                            },
                            step.index,
                          ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("details", {
                    children: [
                      i.jsx("summary", {
                        children: i.jsx(LocalizedText, {
                          children: "Scope and calculation limits",
                        }),
                      }),
                      i.jsx("ul", {
                        children: report.notes.map((note) =>
                          i.jsx("li", { children: note }, note),
                        ),
                      }),
                    ],
                  }),
                ],
              }),
          ],
        }),
        i.jsx("footer", {
          children: i.jsx(LocalizedText, {
            children:
              "Inspection and replay do not edit formulas, history, or calculation settings.",
          }),
        }),
      ],
    }),
  });
}

function CalculationInspectorDialog(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    onCloseRef = ne.useRef(props.onClose);
  onCloseRef.current = props.onClose;
  const [direction, setDirection] = ne.useState(
      props.initialDirection ?? "precedents",
    ),
    [depth, setDepth] = ne.useState(1),
    [checked, setChecked] = ne.useState(false),
    [requestedPage, setRequestedPage] = ne.useState(0),
    [navigationError, setNavigationError] = ne.useState(""),
    inspection = ne.useMemo(() => {
      try {
        return {
          result: inspectFormulaDependencies(props.workbook, {
            sheetId: props.sheetId,
            address: props.address,
            direction: direction,
            depth: depth,
            includeHidden: !!props.allowHidden && checked,
            allowedSheetIds: props.allowedSheetIds,
          }),
        };
      } catch (error) {
        return {
          error:
            error instanceof Error
              ? error.message
              : "The selected cell could not be inspected.",
        };
      }
    }, [
      props.workbook,
      props.sheetId,
      props.address,
      props.allowedSheetIds,
      props.allowHidden,
      direction,
      depth,
      checked,
    ]);
  (ne.useEffect(() => {
    (setRequestedPage(0), setNavigationError(""));
  }, [inspection]),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        dialogRef.current
          ?.querySelector('[role="tab"][aria-selected="true"]')
          ?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  const result = inspection.result,
    formatCellKey = (cellKey) => {
      const [sheetId, address] = JSON.parse(cellKey);
      return `${props.workbook.sheets.find((sheet) => sheet.id === sheetId)?.name ?? "Unavailable sheet"}!${address}`;
    },
    pageCount = Math.max(1, Math.ceil((result?.nodes.length ?? 0) / 50)),
    currentPage = Math.min(requestedPage, pageCount - 1),
    navigateToNode = (node) => {
      try {
        (inspectFormulaDependencies(props.workbook, {
          sheetId: node.sheetId,
          address: node.address,
          direction: "precedents",
          allowedSheetIds: props.allowedSheetIds,
          includeHidden: !!props.allowHidden && checked,
          limits: { maxFormulaCells: 1, maxNodes: 1, maxReferences: 1 },
        }),
          props.onNavigate(node.sheetId, node.address),
          setNavigationError(""));
      } catch (error) {
        setNavigationError(
          error instanceof Error
            ? error.message
            : "The host could not navigate to this cell.",
        );
      }
    };
  return i.jsx("div", {
    className: "fi-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "fi-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${titleId}-title`,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(), onCloseRef.current()),
          event.key === "Tab")
        ) {
          const focusableElements = [
              ...(dialogRef.current?.querySelectorAll(
                'button:not(:disabled):not([tabindex="-1"]),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
              ) ?? []),
            ].filter((element) => !element.hidden),
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1);
          event.shiftKey && document.activeElement === firstFocusable
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === lastFocusable &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${titleId}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Inspect calculations",
                  }),
                }),
                i.jsxs("p", {
                  children: [
                    i.jsx(LocalizedText, {
                      children: "Read-only formula relationships · revision ",
                    }),
                    props.workbook.revision,
                  ],
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close calculation inspection"),
              onClick: props.onClose,
              lang: language("Close calculation inspection"),
              "data-ui-label": "Close calculation inspection",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "fi-controls",
          children: [
            i.jsx("div", {
              role: "tablist",
              "aria-label": text("Formula relationships"),
              lang: language("Formula relationships"),
              "data-ui-label": "Formula relationships",
              children: ["precedents", "dependents"].map((tabDirection) =>
                i.jsx(
                  "button",
                  {
                    role: "tab",
                    id: `${titleId}-${tabDirection}`,
                    "aria-controls": `${titleId}-panel`,
                    "aria-selected": direction === tabDirection,
                    tabIndex: direction === tabDirection ? 0 : -1,
                    onClick: () => setDirection(tabDirection),
                    onKeyDown: (event) => {
                      if (
                        ["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                          event.key,
                        )
                      ) {
                        event.preventDefault();
                        const nextDirection =
                          event.key === "Home"
                            ? "precedents"
                            : event.key === "End" ||
                                tabDirection === "precedents"
                              ? "dependents"
                              : "precedents";
                        (setDirection(nextDirection),
                          document
                            .getElementById(`${titleId}-${nextDirection}`)
                            ?.focus());
                      }
                    },
                    children:
                      tabDirection === "precedents"
                        ? "Precedents"
                        : "Dependents",
                  },
                  tabDirection,
                ),
              ),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Traversal depth" }),
                i.jsx("select", {
                  "aria-label": text("Traversal depth"),
                  value: depth,
                  onChange: (event) => setDepth(Number(event.target.value)),
                  lang: language("Traversal depth"),
                  "data-ui-label": "Traversal depth",
                  children: Array.from({ length: 8 }, (k, optionIndex) =>
                    i.jsx(
                      "option",
                      {
                        value: optionIndex + 1,
                        children:
                          optionIndex === 0
                            ? "1 — Direct"
                            : `${optionIndex + 1} — Transitive`,
                      },
                      optionIndex,
                    ),
                  ),
                }),
              ],
            }),
            props.allowHidden &&
              i.jsxs("label", {
                className: "fi-hidden",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => setChecked(event.target.checked),
                  }),
                  i.jsx(LocalizedText, {
                    children: "Include hidden cells in the allowed context",
                  }),
                ],
              }),
          ],
        }),
        inspection.error &&
          i.jsx("p", {
            role: "alert",
            className: "fi-warning",
            children: inspection.error,
          }),
        navigationError &&
          i.jsx("p", {
            role: "alert",
            className: "fi-warning",
            children: navigationError,
          }),
        result &&
          i.jsxs("div", {
            role: "tabpanel",
            id: `${titleId}-panel`,
            "aria-labelledby": `${titleId}-${direction}`,
            children: [
              i.jsxs("section", {
                className: "fi-target",
                "aria-label": text("Selected cell"),
                lang: language("Selected cell"),
                "data-ui-label": "Selected cell",
                children: [
                  i.jsxs("h3", {
                    children: [
                      result.target.sheetName,
                      "!",
                      result.target.address,
                    ],
                  }),
                  i.jsxs("dl", {
                    children: [
                      i.jsx("dt", {
                        children: i.jsx(LocalizedText, {
                          children: "Stored input",
                        }),
                      }),
                      i.jsx("dd", {
                        children: i.jsx("code", {
                          children: formatCellValueForDisplay(
                            result.target.raw,
                          ),
                        }),
                      }),
                      i.jsx("dt", {
                        children: i.jsx(LocalizedText, {
                          children: "Current local result",
                        }),
                      }),
                      i.jsxs("dd", {
                        children: [
                          result.target.resultStatus === "unavailable"
                            ? "Unavailable"
                            : formatCellValueForDisplay(result.target.result),
                          result.target.resultNote &&
                            i.jsx("small", {
                              children: result.target.resultNote,
                            }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              i.jsxs("div", {
                role: "status",
                className: result.coverage.complete
                  ? "fi-coverage"
                  : "fi-coverage fi-warning",
                children: [
                  i.jsxs("strong", {
                    children: [
                      result.nodes.length,
                      " ",
                      direction,
                      " ",
                      i.jsx(LocalizedText, { children: "listed ·" }),
                      " ",
                      result.coverage.complete
                        ? "Complete within the selected static scope"
                        : "Partial static inspection",
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      result.coverage.formulaCellsScanned,
                      " ",
                      i.jsx(LocalizedText, { children: "of" }),
                      " ",
                      result.coverage.formulaCellsAvailable,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "eligible formula cells scanned; ",
                      }),
                      result.coverage.referencesScanned,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "references inspected. ",
                      }),
                      result.coverage.excludedCells,
                      " ",
                      i.jsx(LocalizedText, {
                        children:
                          "populated cells excluded by visibility or host context.",
                      }),
                    ],
                  }),
                  result.coverage.notes.map((note) =>
                    i.jsx("p", { children: note }, note),
                  ),
                ],
              }),
              i.jsxs("p", {
                className: "fi-explanation",
                children: [
                  direction === "precedents"
                    ? "These cells are referenced by the selected cell, directly or through other formulas."
                    : "These formula cells contain a potential reference path to the selected cell.",
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "Results are freshly evaluated by the local calculator; volatile functions can change between inspections.",
                  }),
                ],
              }),
              result.nodes.length
                ? i.jsx("div", {
                    className: "fi-table-wrap",
                    children: i.jsxs("table", {
                      "aria-label": `${direction} cells`,
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Cell",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Depth",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Stored input",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Local result",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: result.nodes
                            .slice(currentPage * 50, (currentPage + 1) * 50)
                            .map((cell) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("td", {
                                      children: i.jsxs("button", {
                                        onClick: () => navigateToNode(cell),
                                        "aria-label": `Go to ${cell.sheetName}!${cell.address}`,
                                        children: [
                                          cell.sheetName,
                                          "!",
                                          cell.address,
                                        ],
                                      }),
                                    }),
                                    i.jsx("td", { children: cell.depth }),
                                    i.jsx("td", {
                                      children: i.jsx("code", {
                                        children: formatCellValueForDisplay(
                                          cell.raw,
                                        ),
                                      }),
                                    }),
                                    i.jsx("td", {
                                      children:
                                        cell.resultStatus === "unavailable"
                                          ? i.jsx("span", {
                                              title: cell.resultNote,
                                              children: i.jsx(LocalizedText, {
                                                children: "Unavailable",
                                              }),
                                            })
                                          : formatCellValueForDisplay(
                                              cell.result,
                                            ),
                                    }),
                                  ],
                                },
                                cell.key,
                              ),
                            ),
                        }),
                      ],
                    }),
                  })
                : i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, { children: "No " }),
                      direction,
                      " ",
                      i.jsx(LocalizedText, {
                        children:
                          "were found in the inspected subset. Review coverage and unresolved issues before concluding there are none.",
                      }),
                    ],
                  }),
              pageCount > 1 &&
                i.jsxs("nav", {
                  className: "fi-pages",
                  "aria-label": text("Inspection result pages"),
                  lang: language("Inspection result pages"),
                  "data-ui-label": "Inspection result pages",
                  children: [
                    i.jsx("button", {
                      disabled: currentPage === 0,
                      onClick: () => setRequestedPage(currentPage - 1),
                      children: i.jsx(LocalizedText, {
                        children: "Previous results",
                      }),
                    }),
                    i.jsxs("span", {
                      children: [
                        i.jsx(LocalizedText, { children: "Page " }),
                        currentPage + 1,
                        " ",
                        i.jsx(LocalizedText, { children: "of " }),
                        pageCount,
                      ],
                    }),
                    i.jsx("button", {
                      disabled: currentPage + 1 >= pageCount,
                      onClick: () => setRequestedPage(currentPage + 1),
                      children: i.jsx(LocalizedText, {
                        children: "Next results",
                      }),
                    }),
                  ],
                }),
              result.cycles.length > 0 &&
                i.jsxs("section", {
                  className: "fi-warning",
                  "aria-label": text("Detected cycles"),
                  lang: language("Detected cycles"),
                  "data-ui-label": "Detected cycles",
                  children: [
                    i.jsx("h3", {
                      children: i.jsx(LocalizedText, {
                        children: "Detected static cycles",
                      }),
                    }),
                    i.jsx("ul", {
                      children: result.cycles.map((cycle, cycleIndex) =>
                        i.jsx(
                          "li",
                          { children: cycle.map(formatCellKey).join(" → ") },
                          cycleIndex,
                        ),
                      ),
                    }),
                  ],
                }),
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, { children: "Reference ranges (" }),
                      result.references.length,
                      ")",
                    ],
                  }),
                  i.jsx("ul", {
                    children: result.references.map(
                      (reference, referenceIndex) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              i.jsx("code", {
                                children: formatCellKey(reference.from),
                              }),
                              " →",
                              " ",
                              props.workbook.sheets.find(
                                (sheet) => sheet.id === reference.sheetId,
                              )?.name,
                              "!",
                              reference.range,
                              " · ",
                              reference.cellCount.toLocaleString(),
                              " ",
                              i.jsx(LocalizedText, { children: "cells ·" }),
                              " ",
                              reference.kind,
                              " ·",
                              " ",
                              i.jsx("code", { children: reference.token }),
                            ],
                          },
                          referenceIndex,
                        ),
                    ),
                  }),
                ],
              }),
              i.jsxs("details", {
                open: result.issues.length > 0,
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Unresolved or excluded dependencies (",
                      }),
                      result.issues.length,
                      ")",
                    ],
                  }),
                  result.issues.length
                    ? i.jsx("ul", {
                        children: result.issues.map((issue, issueIndex) =>
                          i.jsxs(
                            "li",
                            {
                              children: [
                                i.jsx("strong", { children: issue.kind }),
                                " · ",
                                formatCellKey(issue.from),
                                issue.token &&
                                  i.jsxs(i.Fragment, {
                                    children: [
                                      " ",
                                      "· ",
                                      i.jsx("code", { children: issue.token }),
                                    ],
                                  }),
                                " ",
                                "— ",
                                issue.message,
                              ],
                            },
                            issueIndex,
                          ),
                        ),
                      })
                    : i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "No unresolved references were identified by the static scanner.",
                        }),
                      }),
                ],
              }),
            ],
          }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children: i.jsx(LocalizedText, {
                children:
                  "Inspection does not change cells, history, or calculation settings.",
              }),
            }),
            i.jsx("button", {
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Close" }),
            }),
          ],
        }),
      ],
    }),
  });
}
function normalizeFormulaWatchContext(workbook, hostContext) {
  if (
    hostContext.includeHidden !== undefined &&
    typeof hostContext.includeHidden != "boolean"
  )
    throw Error("Include hidden must be boolean.");
  if (
    hostContext.allowedSheetIds !== undefined &&
    (!Array.isArray(hostContext.allowedSheetIds) ||
      hostContext.allowedSheetIds.some(
        (allowedSheetId) =>
          typeof allowedSheetId != "string" ||
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ))
  )
    throw Error(
      "The formula watch host context contains an unknown worksheet.",
    );
  return {
    includeHidden: !!hostContext.includeHidden,
    allowedSheetIds: [
      ...new Set(
        hostContext.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
  };
}
function classifyWatchedCell(workbook, watchedCell, formulaWatchContext) {
  const sheet = workbook.sheets.find(
    (sheet) => sheet.id === watchedCell.anchor.sheetId,
  );
  if (!sheet) return "orphaned";
  if (!formulaWatchContext.allowedSheetIds.includes(sheet.id))
    return "excluded";
  if (!formulaWatchContext.includeHidden) {
    const cellPosition = _t(watchedCell.anchor.address);
    if (
      sheet.hidden ||
      sheet.veryHidden ||
      cn(sheet, "row").some(
        (rowRange) =>
          rowRange.start <= cellPosition.row &&
          rowRange.end >= cellPosition.row,
      ) ||
      cn(sheet, "column").some(
        (columnRange) =>
          columnRange.start <= cellPosition.col &&
          columnRange.end >= cellPosition.col,
      )
    )
      return "hidden";
  }
  return watchedCell.orphaned ? "orphaned" : "ready";
}

export {
  normalizeFormulaWatchContext,
  classifyWatchedCell,
  FormulaCompatibilityDialog,
  CalculationInspectorDialog,
  EvaluateFormulaDialog,
};
