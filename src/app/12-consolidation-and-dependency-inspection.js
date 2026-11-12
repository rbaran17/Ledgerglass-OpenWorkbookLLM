import { hasOnlyAllowedKeys } from "./19b-assistant-protocol.js";

import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { isNullOrEmptyValue, failWithMessage } from "./03-workbook-model.js";
import {
  parseFormulaTree,
  walkFormulaNodes,
  tokenizeFormulaText,
  fingerprintCanonicalValue,
  normalizeBindingName,
  normalizeDatePolicy,
} from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  assertWorkbookStructureEditable,
  serializeTypedKeyValues,
  createWorksheetRecord,
  sumConsolidationAmounts,
  calculateWorkbookWithSettings,
  _t,
  cn,
  parseRangeAddress,
  buildConsolidationSubtotals,
  consolidationLimits,
  parseConsolidationDefinition,
  resolveDefinedNameValue,
  normalizeWorkbookInput,
  columnIndexToLetters,
  matrixResultFunctionNames,
  rewriteStructuredReferences,
} from "./03-workbook-model.js";
import {
  requireIntegerInRange,
  singleRangeReferencePattern,
  buildCompositeKey,
  parseCellReference,
  bareSheetNamePattern,
  knownFunctionNameSet,
  isFormulaInputValue,
  sheetRangeReferencePattern,
} from "./11-local-workbook-store-and-versions.js";

function inspectFormulaDependencies(workbook, request) {
  if (!["precedents", "dependents"].includes(request.direction))
    throw Error("Choose precedents or dependents.");
  const depth = requireIntegerInRange(request.depth, 1, 8, "Inspection depth"),
    formulaScanLimit = requireIntegerInRange(
      request.limits?.maxFormulaCells,
      5e3,
      1e4,
      "Formula scan limit",
    ),
    resultCellLimit = requireIntegerInRange(
      request.limits?.maxNodes,
      500,
      1e3,
      "Result cell limit",
    ),
    referenceScanLimit = requireIntegerInRange(
      request.limits?.maxReferences,
      1e4,
      5e4,
      "Reference scan limit",
    );
  if (
    request.includeHidden !== undefined &&
    typeof request.includeHidden != "boolean"
  )
    throw Error("Include hidden must be boolean.");
  if (
    request.allowedSheetIds &&
    (!Array.isArray(request.allowedSheetIds) ||
      request.allowedSheetIds.some(
        (allowedSheetId) =>
          typeof allowedSheetId != "string" ||
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ))
  )
    throw Error("The allowed worksheet context is invalid.");
  const sheetsById = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet])),
    allowedSheetIdSet = new Set(
      request.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
    ),
    hiddenRangesBySheetId = new Map(
      workbook.sheets.map((sheet) => [
        sheet.id,
        { rows: cn(sheet, "row"), cols: cn(sheet, "column") },
      ]),
    ),
    isCellVisible = (worksheet, addressText) => {
      if (!allowedSheetIdSet.has(worksheet.id)) return false;
      if (request.includeHidden) return true;
      const position = _t(addressText);
      return (
        !worksheet.hidden &&
        !hiddenRangesBySheetId
          .get(worksheet.id)
          .rows.some(
            (row) => position.row >= row.start && position.row <= row.end,
          ) &&
        !hiddenRangesBySheetId
          .get(worksheet.id)
          .cols.some(
            (col) => position.col >= col.start && position.col <= col.end,
          )
      );
    },
    cellReference = parseCellReference(request.address);
  if (cellReference.includes(":")) throw Error("Choose one cell to inspect.");
  const sheet = sheetsById.get(request.sheetId);
  if (!sheet || !isCellVisible(sheet, cellReference))
    throw Error(
      "The selected cell is outside the visible, authorized inspection context.",
    );
  const compositeKey = buildCompositeKey(sheet.id, cellReference),
    issues = [],
    references = [],
    referencesByCellKey = new Map(),
    blockedCellKeys = new Set();
  let referencesScanned = 0,
    truncated = false,
    depthLimited = false,
    excludedCells = 0;
  const addIssue = (from, issueKind, message, token) => {
      (blockedCellKeys.add(from),
        issues.length < 200 &&
        !issues.some(
          (issue) =>
            issue.from === from &&
            issue.kind === issueKind &&
            issue.token === token &&
            issue.message === message,
        )
          ? issues.push({
              from: from,
              kind: issueKind,
              message: message,
              ...(token ? { token: token } : {}),
            })
          : issues.length >= 200 && (truncated = true));
    },
    formulaCells = [];
  for (const sheet of workbook.sheets)
    for (const entry of Object.entries(sheet.cells).sort(
      ([leftAddress], [rightAddress]) => {
        const leftPosition = _t(leftAddress),
          rightPosition = _t(rightAddress);
        return (
          leftPosition.row - rightPosition.row ||
          leftPosition.col - rightPosition.col
        );
      },
    )) {
      if (!isCellVisible(sheet, entry[0])) {
        excludedCells++;
        continue;
      }
      isFormulaInputValue(entry[1]) &&
        formulaCells.push({
          key: buildCompositeKey(sheet.id, entry[0]),
          sheet: sheet,
          address: entry[0],
          raw: entry[1].raw,
        });
    }
  formulaCells.sort((event, otherCell) =>
    event.key === compositeKey ? -1 : otherCell.key === compositeKey ? 1 : 0,
  );
  const scannedFormulaCells = formulaCells.slice(0, formulaScanLimit);
  formulaCells.length > scannedFormulaCells.length &&
    ((truncated = true),
    addIssue(
      compositeKey,
      "limit",
      "The formula-cell scan limit was reached; additional dependents may exist.",
    ));
  for (const cell of scannedFormulaCells) {
    const cellReferences = [];
    referencesByCellKey.set(cell.key, cellReferences);
    let expansionCount = 0;
    const addReference = (sheetId, address, token, referenceKind) => {
        if (referencesScanned >= referenceScanLimit) {
          ((truncated = true),
            addIssue(
              cell.key,
              "limit",
              "The workbook reference scan limit was reached.",
            ));
          return;
        }
        referencesScanned++;
        const referencedSheet = sheetsById.get(sheetId);
        if (!referencedSheet) {
          addIssue(
            cell.key,
            "missing",
            "The referenced worksheet does not exist.",
            token,
          );
          return;
        }
        if (
          !allowedSheetIdSet.has(sheetId) ||
          (referencedSheet.hidden && !request.includeHidden)
        ) {
          addIssue(
            cell.key,
            "excluded",
            "A referenced worksheet is outside the selected inspection context; its cells and values were excluded.",
          );
          return;
        }
        try {
          const cellReference = parseCellReference(address),
            rangeAddress = parseRangeAddress(cellReference);
          (cellReferences.push({
            from: cell.key,
            token: token,
            sheetId: sheetId,
            range: cellReference,
            kind: referenceKind,
            cellCount:
              (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
              (rangeAddress.lastCol - rangeAddress.firstCol + 1),
          }),
            !request.includeHidden &&
              (hiddenRangesBySheetId
                .get(referencedSheet.id)
                .rows.some(
                  (row) =>
                    row.start <= rangeAddress.lastRow &&
                    row.end >= rangeAddress.firstRow,
                ) ||
                hiddenRangesBySheetId
                  .get(referencedSheet.id)
                  .cols.some(
                    (col) =>
                      col.start <= rangeAddress.lastCol &&
                      col.end >= rangeAddress.firstCol,
                  )) &&
              addIssue(
                cell.key,
                "excluded",
                "A referenced range intersects hidden rows or columns; hidden cell details are excluded.",
                token,
              ));
        } catch {
          addIssue(
            cell.key,
            "missing",
            "The reference is outside Excel's address bounds or is invalid.",
            token,
          );
        }
      },
      expandFormulaText = (
        formulaText,
        contextSheetId,
        nameChain = [],
        originToken,
      ) => {
        if (++expansionCount > 64 || nameChain.length > 16) {
          ((truncated = true),
            addIssue(
              cell.key,
              "limit",
              "The defined-name expansion limit was reached.",
              originToken,
            ));
          return;
        }
        let scanIndex = formulaText.startsWith("=") ? 1 : 0;
        const localBindingSpans = new Map();
        if (/\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(formulaText))
          try {
            walkFormulaNodes(
              tokenizeFormulaText(formulaText),
              (node, localBindingNames, Te, Pe) => {
                (Pe ||
                  (!node.qualifier &&
                    localBindingNames.has(normalizeBindingName(node.name)))) &&
                  localBindingSpans.set(node.start, node.end);
              },
            );
          } catch {
            addIssue(
              cell.key,
              "unsupported",
              "Invalid LET/LAMBDA bindings prevent complete dependency inspection.",
            );
            return;
          }
        for (; scanIndex < formulaText.length;) {
          if (localBindingSpans.has(scanIndex)) {
            scanIndex = localBindingSpans.get(scanIndex);
            continue;
          }
          if (referencesScanned >= referenceScanLimit) {
            ((truncated = true),
              addIssue(
                cell.key,
                "limit",
                "The workbook reference scan limit was reached.",
              ));
            return;
          }
          const remainingText = formulaText.slice(scanIndex),
            currentChar = formulaText[scanIndex];
          if (currentChar === '"') {
            scanIndex++;
            let isTerminated = false;
            for (; scanIndex < formulaText.length;)
              if (formulaText[scanIndex++] === '"')
                if (formulaText[scanIndex] === '"') scanIndex++;
                else {
                  isTerminated = true;
                  break;
                }
            isTerminated ||
              addIssue(
                cell.key,
                "unsupported",
                "Unterminated formula text cannot be inspected completely.",
              );
            continue;
          }
          const externalMatch =
            /^(?:'(?:[^']|'')*\[[^\]]+\](?:[^']|'')*'\s*!\s*[^+*/(),;\s]+|\[[^\]]+\][^!+*/(),;]*![^+*/(),;\s]+)/.exec(
              remainingText,
            );
          if (externalMatch) {
            (addIssue(
              cell.key,
              "external",
              "External workbook references are not opened or fetched.",
              externalMatch[0],
            ),
              (scanIndex += externalMatch[0].length));
            continue;
          }
          const sheetRangeMatch =
            sheetRangeReferencePattern.exec(remainingText);
          if (
            sheetRangeMatch ||
            /^'(?:[^']|'')*:(?:[^']|'')*'\s*!/.test(remainingText)
          ) {
            const sheetRangeText =
              sheetRangeMatch?.[0] ?? /^[^+*/(),;]+/.exec(remainingText)[0];
            (addIssue(
              cell.key,
              "unsupported",
              "Three-dimensional worksheet references are not expanded.",
              sheetRangeText,
            ),
              (scanIndex += sheetRangeText.length));
            continue;
          }
          const errorMatch = /^#[A-Z0-9/?!]+/i.exec(remainingText);
          if (errorMatch) {
            (addIssue(
              cell.key,
              errorMatch[0].toUpperCase() === "#REF!"
                ? "missing"
                : "unsupported",
              "The formula contains an error token.",
              errorMatch[0],
            ),
              (scanIndex += errorMatch[0].length));
            continue;
          }
          const tableMatch = /^([A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(
            remainingText,
          );
          if (tableMatch) {
            const structuredStart = scanIndex;
            scanIndex += tableMatch[0].length;
            let bracketDepth = 1;
            for (; scanIndex < formulaText.length && bracketDepth;)
              (formulaText[scanIndex] === "[" && bracketDepth++,
                formulaText[scanIndex] === "]" && bracketDepth--,
                scanIndex++);
            const structuredReferenceText = formulaText.slice(
                structuredStart,
                scanIndex,
              ),
              resolvedStructuredReference =
                !bracketDepth &&
                rewriteStructuredReferences(
                  workbook,
                  contextSheetId,
                  structuredReferenceText,
                  cell.address,
                );
            resolvedStructuredReference
              ? addReference(
                  resolvedStructuredReference.sheetId,
                  resolvedStructuredReference.range,
                  originToken ?? structuredReferenceText,
                  originToken ? "name" : "table",
                )
              : addIssue(
                  cell.key,
                  "unsupported",
                  "The table, field, or structured-reference form could not be resolved.",
                  structuredReferenceText,
                );
            continue;
          }
          const sheetQualifiedMatch =
            /^(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))\s*!\s*([A-Za-z_\\][A-Za-z0-9_.\\]*)/.exec(
              remainingText,
            );
          if (
            sheetQualifiedMatch &&
            !/^[A-Za-z]{1,3}[1-9]\d*$/i.test(sheetQualifiedMatch[3])
          ) {
            const qualifiedSheetName = (
                sheetQualifiedMatch[1] ?? sheetQualifiedMatch[2]
              ).replace(/''/g, "'"),
              sheet = workbook.sheets.find(
                (sheet) =>
                  sheet.name.toLowerCase() === qualifiedSheetName.toLowerCase(),
              ),
              qualifiedReferenceText = sheetQualifiedMatch[0],
              qualifiedName = sheetQualifiedMatch[3];
            if (!sheet)
              addIssue(
                cell.key,
                "missing",
                "The referenced worksheet does not exist.",
                qualifiedReferenceText,
              );
            else if (
              !allowedSheetIdSet.has(sheet.id) ||
              (sheet.hidden && !request.includeHidden)
            )
              addIssue(
                cell.key,
                "excluded",
                "The defined name belongs to a worksheet outside the selected inspection context.",
              );
            else {
              const qualifiedDefinedName =
                workbook.definedNames?.find(
                  (definedName) =>
                    definedName.scopeSheetId === sheet.id &&
                    definedName.name.toLowerCase() ===
                      qualifiedName.toLowerCase(),
                ) ??
                workbook.definedNames?.find(
                  (definedName) =>
                    !definedName.scopeSheetId &&
                    definedName.name.toLowerCase() ===
                      qualifiedName.toLowerCase(),
                );
              if (!qualifiedDefinedName)
                addIssue(
                  cell.key,
                  "missing",
                  "The qualified defined name does not exist.",
                  qualifiedReferenceText,
                );
              else {
                const nameChainKey = `${qualifiedDefinedName.scopeSheetId ?? ""}:${qualifiedDefinedName.name.toLowerCase()}`;
                nameChain.includes(nameChainKey)
                  ? addIssue(
                      cell.key,
                      parseFormulaTree(qualifiedDefinedName.refersTo)
                        ? "dynamic"
                        : "cycle",
                      "The reusable definition recurs; references are shown once.",
                      qualifiedReferenceText,
                    )
                  : expandFormulaText(
                      qualifiedDefinedName.refersTo,
                      qualifiedDefinedName.scopeSheetId ?? sheet.id,
                      [...nameChain, nameChainKey],
                      originToken ?? qualifiedReferenceText,
                    );
              }
            }
            scanIndex += qualifiedReferenceText.length;
            continue;
          }
          const functionCallMatch = /^([A-Za-z_\\][A-Za-z0-9_.\\]*)\s*\(/.exec(
            remainingText,
          );
          if (functionCallMatch) {
            const functionName = functionCallMatch[1]
              .toUpperCase()
              .replace(/^_XLFN\./, "")
              .replace(/^_XLWS\./, "");
            if (["INDIRECT", "OFFSET"].includes(functionName))
              addIssue(
                cell.key,
                "dynamic",
                "Runtime reference selection is not resolved; static argument references are only potential dependencies.",
                functionCallMatch[1],
              );
            else if (!knownFunctionNameSet.has(functionName)) {
              const functionDefinedName =
                workbook.definedNames?.find(
                  (definedName) =>
                    definedName.scopeSheetId === contextSheetId &&
                    definedName.name.toUpperCase() === functionName,
                ) ??
                workbook.definedNames?.find(
                  (definedName) =>
                    !definedName.scopeSheetId &&
                    definedName.name.toUpperCase() === functionName,
                );
              if (functionDefinedName) {
                const functionNameKey = `${functionDefinedName.scopeSheetId ?? ""}:${functionDefinedName.name.toLowerCase()}`;
                nameChain.includes(functionNameKey)
                  ? addIssue(
                      cell.key,
                      parseFormulaTree(functionDefinedName.refersTo)
                        ? "dynamic"
                        : "cycle",
                      "Recursive reusable formula: static references are shown once; execution depth is bounded by calculation.",
                      functionCallMatch[1],
                    )
                  : expandFormulaText(
                      functionDefinedName.refersTo,
                      functionDefinedName.scopeSheetId ?? contextSheetId,
                      [...nameChain, functionNameKey],
                      originToken ?? functionCallMatch[1],
                    );
              } else
                addIssue(
                  cell.key,
                  "unsupported",
                  "The function is outside the local supported catalog; its static references may be incomplete.",
                  functionCallMatch[1],
                );
            }
            scanIndex += functionCallMatch[0].length;
            continue;
          }
          const singleRangeMatch =
            singleRangeReferencePattern.exec(remainingText);
          if (
            singleRangeMatch &&
            !/[A-Za-z0-9_.!:]/.test(
              remainingText[singleRangeMatch[0].length] ?? "",
            )
          ) {
            const sheetPrefix = singleRangeMatch[1]
                ?.trim()
                .replace(/!\s*$/, "")
                .trim(),
              unquotedSheetName = sheetPrefix
                ?.replace(/^'|'$/g, "")
                .replace(/''/g, "'"),
              targetSheet =
                unquotedSheetName === undefined
                  ? sheetsById.get(contextSheetId)
                  : workbook.sheets.find(
                      (sheet) =>
                        sheet.name.toLowerCase() ===
                        unquotedSheetName.toLowerCase(),
                    );
            (targetSheet
              ? addReference(
                  targetSheet.id,
                  singleRangeMatch[2],
                  originToken ?? singleRangeMatch[0],
                  originToken
                    ? "name"
                    : singleRangeMatch[3]
                      ? "spill-anchor"
                      : "a1",
                )
              : addIssue(
                  cell.key,
                  "missing",
                  "The referenced worksheet does not exist.",
                  singleRangeMatch[0],
                ),
              singleRangeMatch[3] &&
                addIssue(
                  cell.key,
                  "dynamic",
                  "Only the spill anchor is included; runtime spill output dependencies are not expanded.",
                  singleRangeMatch[0],
                ),
              (scanIndex += singleRangeMatch[0].length));
            continue;
          }
          const numberMatch = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(
            remainingText,
          );
          if (numberMatch) {
            scanIndex += numberMatch[0].length;
            continue;
          }
          const bareSheetMatch = bareSheetNamePattern.exec(remainingText);
          if (bareSheetMatch) {
            const bareSheetName = bareSheetMatch[0],
              textAfterSheetName = remainingText.slice(bareSheetName.length);
            if (/^\s*!/.test(textAfterSheetName)) {
              const sheetQualifiedNameText = /^[^+*/(),;\s]+/.exec(
                remainingText,
              )[0];
              (addIssue(
                cell.key,
                "unsupported",
                "Worksheet-qualified defined names are not resolved by this inspector.",
                sheetQualifiedNameText,
              ),
                (scanIndex += sheetQualifiedNameText.length));
              continue;
            }
            const bareDefinedName =
              workbook.definedNames?.find(
                (definedName) =>
                  definedName.scopeSheetId === contextSheetId &&
                  definedName.name.toLowerCase() ===
                    bareSheetName.toLowerCase(),
              ) ??
              workbook.definedNames?.find(
                (definedName) =>
                  !definedName.scopeSheetId &&
                  definedName.name.toLowerCase() ===
                    bareSheetName.toLowerCase(),
              );
            if (bareDefinedName) {
              const bareNameKey = `${bareDefinedName.scopeSheetId ?? ""}:${bareDefinedName.name.toLowerCase()}`;
              nameChain.includes(bareNameKey)
                ? addIssue(
                    cell.key,
                    "cycle",
                    "Defined names contain a circular reference.",
                    bareSheetName,
                  )
                : expandFormulaText(
                    bareDefinedName.refersTo,
                    bareDefinedName.scopeSheetId ?? contextSheetId,
                    [...nameChain, bareNameKey],
                    originToken ?? bareSheetName,
                  );
            } else
              ["TRUE", "FALSE"].includes(bareSheetName.toUpperCase()) ||
                addIssue(
                  cell.key,
                  "missing",
                  "The identifier is not a defined name or supported reference.",
                  bareSheetName,
                );
            scanIndex += bareSheetName.length;
            continue;
          }
          ((currentChar === "'" ||
            currentChar === "!" ||
            currentChar === "[" ||
            currentChar === "]" ||
            currentChar === "@" ||
            currentChar === "|") &&
            addIssue(
              cell.key,
              "unsupported",
              "The reference syntax is outside the inspected subset.",
              currentChar,
            ),
            scanIndex++);
        }
      };
    (expandFormulaText(cell.raw, cell.sheet.id),
      references.push(...cellReferences));
  }
  const graphNodes = new Map([
      [compositeKey, { sheet: sheet, address: cellReference, depth: 0 }],
    ]),
    edgesByKey = new Map();
  let frontierKeys = [compositeKey];
  const addDependencyEdge = (sheet, address, depth, from, toKey) => {
    if (!isCellVisible(sheet, address)) return;
    const compositeKey = buildCompositeKey(sheet.id, address);
    if (!graphNodes.has(compositeKey)) {
      if (graphNodes.size - 1 >= resultCellLimit) {
        truncated = true;
        return;
      }
      graphNodes.set(compositeKey, {
        sheet: sheet,
        address: address,
        depth: depth,
      });
    }
    edgesByKey.set(JSON.stringify([from, toKey]), { from: from, to: toKey });
  };
  for (
    let traversalLevel = 1;
    traversalLevel <= depth && frontierKeys.length;
    traversalLevel++
  ) {
    const previousNodeKeys = new Set(graphNodes.keys());
    for (const frontierKey of frontierKeys) {
      const frontierNode = graphNodes.get(frontierKey);
      if (request.direction === "precedents") {
        isFormulaInputValue(frontierNode.sheet.cells[frontierNode.address]) &&
          !referencesByCellKey.has(frontierKey) &&
          ((truncated = true),
          addIssue(
            frontierKey,
            "limit",
            "This formula was outside the formula scan budget.",
          ));
        for (const reference of referencesByCellKey.get(frontierKey) ?? []) {
          const rangeAddress = parseRangeAddress(reference.range),
            referenceSheet = sheetsById.get(reference.sheetId);
          let expandedCellCount = 0;
          e: for (
            let firstRow = rangeAddress.firstRow;
            firstRow <= rangeAddress.lastRow;
            firstRow++
          )
            for (
              let firstCol = rangeAddress.firstCol;
              firstCol <= rangeAddress.lastCol;
              firstCol++
            ) {
              if (
                ++expandedCellCount > resultCellLimit + 1 ||
                graphNodes.size - 1 >= resultCellLimit
              ) {
                truncated = true;
                break e;
              }
              const cellAddress = formatCellAddress(firstRow, firstCol);
              addDependencyEdge(
                referenceSheet,
                cellAddress,
                traversalLevel,
                frontierKey,
                buildCompositeKey(referenceSheet.id, cellAddress),
              );
            }
        }
      } else {
        const dependentPosition = _t(frontierNode.address);
        for (const formulaCell of scannedFormulaCells)
          (referencesByCellKey.get(formulaCell.key) ?? []).some(
            (candidateReference) => {
              const rangeAddress = parseRangeAddress(candidateReference.range);
              return (
                candidateReference.sheetId === frontierNode.sheet.id &&
                dependentPosition.row >= rangeAddress.firstRow &&
                dependentPosition.row <= rangeAddress.lastRow &&
                dependentPosition.col >= rangeAddress.firstCol &&
                dependentPosition.col <= rangeAddress.lastCol
              );
            },
          ) &&
            addDependencyEdge(
              formulaCell.sheet,
              formulaCell.address,
              traversalLevel,
              formulaCell.key,
              frontierKey,
            );
      }
    }
    frontierKeys = [...graphNodes.keys()].filter(
      (newNodeKey) => !previousNodeKeys.has(newNodeKey),
    );
  }
  for (const terminalNodeKey of frontierKeys) {
    const terminalNode = graphNodes.get(terminalNodeKey);
    (request.direction === "precedents"
      ? isFormulaInputValue(terminalNode.sheet.cells[terminalNode.address]) &&
        ((referencesByCellKey.get(terminalNodeKey)?.length ?? 0) > 0 ||
          !referencesByCellKey.has(terminalNodeKey))
      : references.some((referenceRecord) => {
          const nodePosition = _t(terminalNode.address),
            rangeAddress = parseRangeAddress(referenceRecord.range);
          return (
            referenceRecord.sheetId === terminalNode.sheet.id &&
            nodePosition.row >= rangeAddress.firstRow &&
            nodePosition.row <= rangeAddress.lastRow &&
            nodePosition.col >= rangeAddress.firstCol &&
            nodePosition.col <= rangeAddress.lastCol
          );
        })) && (depthLimited = true);
  }
  const edges = [...edgesByKey.values()],
    cycles = [],
    visitState = new Map(),
    visitNode = (currentNodeKey, visitPath) => {
      if (visitState.get(currentNodeKey) === 1) {
        const cyclePath = [
          ...visitPath.slice(visitPath.indexOf(currentNodeKey)),
          currentNodeKey,
        ];
        cycles.length < 50 && cycles.push(cyclePath);
        return;
      }
      if (!(visitState.get(currentNodeKey) === 2 || visitPath.length > 1e3)) {
        visitState.set(currentNodeKey, 1);
        for (const edge of edges)
          edge.from === currentNodeKey &&
            visitNode(edge.to, [...visitPath, currentNodeKey]);
        visitState.set(currentNodeKey, 2);
      }
    };
  for (const graphNodeKey of graphNodes.keys()) visitNode(graphNodeKey, []);
  for (const cycle of cycles)
    for (const cycleNodeKey of cycle) blockedCellKeys.add(cycleNodeKey);
  const canResolveStatically = (
      cellKey,
      ancestorKeys = new Set(),
      budget = { cells: 0 },
    ) => {
      if (
        ancestorKeys.has(cellKey) ||
        ancestorKeys.size >= 32 ||
        ++budget.cells > 2e3 ||
        blockedCellKeys.has(cellKey)
      )
        return false;
      const [keySheetId, keyAddress] = JSON.parse(cellKey),
        keySheet = sheetsById.get(keySheetId);
      if (!isCellVisible(keySheet, keyAddress)) return false;
      if (!isFormulaInputValue(keySheet.cells[keyAddress])) return true;
      const keyReferences = referencesByCellKey.get(cellKey);
      if (!keyReferences) return false;
      const nextAncestorKeys = new Set(ancestorKeys).add(cellKey);
      for (const keyReference of keyReferences) {
        if (keyReference.cellCount > 2e3) return false;
        const rangeAddress = parseRangeAddress(keyReference.range);
        for (
          let firstRow = rangeAddress.firstRow;
          firstRow <= rangeAddress.lastRow;
          firstRow++
        )
          for (
            let firstCol = rangeAddress.firstCol;
            firstCol <= rangeAddress.lastCol;
            firstCol++
          )
            if (
              !canResolveStatically(
                buildCompositeKey(
                  keyReference.sheetId,
                  formatCellAddress(firstRow, firstCol),
                ),
                nextAncestorKeys,
                budget,
              )
            )
              return false;
      }
      return true;
    },
    inspectionWorkbook = {
      ...workbook,
      sheets: workbook.sheets
        .filter(
          (sheet) =>
            allowedSheetIdSet.has(sheet.id) &&
            (request.includeHidden || !sheet.hidden),
        )
        .map((inspectionSheet) => ({
          ...inspectionSheet,
          cells: Object.fromEntries(
            Object.entries(inspectionSheet.cells).filter(([entryAddress]) =>
              isCellVisible(inspectionSheet, entryAddress),
            ),
          ),
        })),
    },
    workbookWithSettings = calculateWorkbookWithSettings(inspectionWorkbook),
    describeNode = (describedNodeKey) => {
      const graphNode = graphNodes.get(describedNodeKey),
        cell = graphNode.sheet.cells[graphNode.address],
        nodeSummary = {
          key: describedNodeKey,
          sheetId: graphNode.sheet.id,
          sheetName: graphNode.sheet.name,
          address: graphNode.address,
          raw: cell?.raw ?? null,
          depth: graphNode.depth,
        };
      if (!isFormulaInputValue(cell))
        return {
          ...nodeSummary,
          result: cell?.raw ?? null,
          resultStatus: "literal",
        };
      if (!canResolveStatically(describedNodeKey))
        return {
          ...nodeSummary,
          result: null,
          resultStatus: "unavailable",
          resultNote:
            "Result withheld: the static dependency closure is cyclic, unresolved, excluded, or beyond the inspection budget.",
        };
      try {
        return {
          ...nodeSummary,
          result: workbookWithSettings(graphNode.sheet.id, graphNode.address),
          resultStatus: "calculated",
        };
      } catch {
        return {
          ...nodeSummary,
          result: null,
          resultStatus: "unavailable",
          resultNote: "The local calculator could not evaluate this cell.",
        };
      }
    },
    notes = [
      "Static potential dependencies include all formula branches; they do not describe the executed branch or evaluation steps.",
      "Blank referenced cells are valid local precedents. Runtime spill outputs are not expanded.",
      "Visibility preferences are local UI controls. The embedding host must supply authorized workbook data and enforce access independently.",
    ];
  return (
    excludedCells &&
      notes.push(
        "Hidden or host-excluded populated cells were omitted from scanning and result details.",
      ),
    truncated &&
      notes.push(
        "One or more budgets were reached. Additional references, cells, issues, or dependents may exist.",
      ),
    depthLimited &&
      notes.push(
        "The requested traversal depth ended at cells with further potential relationships.",
      ),
    cycles.length &&
      notes.push(
        "Cycles were found in the displayed static graph; other cycles may exist beyond the selected depth or scan limits.",
      ),
    issues.length >= 200 &&
      notes.push("Only the first 200 distinct issues are listed."),
    {
      workbookId: workbook.id,
      revision: workbook.revision,
      direction: request.direction,
      depth: depth,
      target: describeNode(compositeKey),
      nodes: [...graphNodes.keys()]
        .filter((nodeKey) => nodeKey !== compositeKey)
        .map(describeNode),
      edges: edges,
      references: references.filter((referenceEntry) =>
        graphNodes.has(referenceEntry.from),
      ),
      issues: issues,
      cycles: cycles,
      coverage: {
        formulaCellsAvailable: formulaCells.length,
        formulaCellsScanned: scannedFormulaCells.length,
        referencesScanned: referencesScanned,
        excludedCells: excludedCells,
        truncated: truncated,
        depthLimited: depthLimited,
        complete:
          !truncated &&
          !depthLimited &&
          !issues.length &&
          !cycles.length &&
          !excludedCells,
        notes: notes,
      },
    }
  );
}
const fuzzyMatchLimits = Object.freeze({
  fuzzyCharacters: 512,
  comparisonWork: 1e7,
  matchPairs: 1e3,
});
function parseKeyNormalization(normalization, fuzzy, keyCount) {
  if (
    normalization !== undefined &&
    (!hasOnlyAllowedKeys(normalization, [
      "trim",
      "collapseWhitespace",
      "ignoreCase",
      "unicode",
    ]) ||
      [
        normalization.trim,
        normalization.collapseWhitespace,
        normalization.ignoreCase,
      ].some((option) => option !== undefined && typeof option != "boolean") ||
      (normalization.unicode !== undefined &&
        !["NFC", "NFKC"].includes(normalization.unicode)))
  )
    throw Error("Invalid explicit key normalization options.");
  if (
    fuzzy !== undefined &&
    (!hasOnlyAllowedKeys(fuzzy, ["key", "threshold"]) ||
      !Number.isInteger(fuzzy.key) ||
      fuzzy.key < 0 ||
      fuzzy.key >= keyCount ||
      typeof fuzzy.threshold != "number" ||
      !Number.isFinite(fuzzy.threshold) ||
      fuzzy.threshold < 0 ||
      fuzzy.threshold > 1)
  )
    throw Error(
      "Choose a fuzzy key pair and a similarity threshold from 0 to 1.",
    );
  return {
    ...(normalization ? { normalization: { ...normalization } } : {}),
    ...(fuzzy ? { fuzzy: { ...fuzzy } } : {}),
  };
}
function normalizeKeyText(text, normalization = {}) {
  if (typeof text != "string") return text;
  let normalizedText = normalization.unicode
    ? text.normalize(normalization.unicode)
    : text;
  return (
    normalization.collapseWhitespace &&
      (normalizedText = normalizedText.replace(/\s+/gu, " ")),
    normalization.trim && (normalizedText = normalizedText.trim()),
    normalization.ignoreCase &&
      (normalizedText = normalizedText.toLocaleLowerCase("en-US")),
    normalizedText
  );
}

function matchKeyPopulations(
  records,
  rightRecords,
  leftKeyColumns,
  rightKeyColumns,
  normalization,
  fuzzy,
) {
  parseKeyNormalization(normalization, fuzzy, leftKeyColumns.length);
  const diagnostics = {
    policy: fuzzy
      ? `Normalized Unicode code-point Levenshtein similarity: 1 - edit distance / maximum text length. Fuzzy key ${fuzzy.key + 1}; inclusive threshold ${fuzzy.threshold}. Other keys compare typed values exactly after selected text normalization. All qualifying matches are retained; scores are not probabilities or Power Query equivalents.`
      : "Exact typed composite keys after explicitly selected text normalization. Blank keys never match; numbers and booleans are not coerced.",
    normalizedKeys: [],
    exceptions: [],
    matches: [],
    ambiguousLeftRows: 0,
    ambiguousRightRows: 0,
    normalizationCollisionGroups: { left: 0, right: 0 },
    comparisonWork: 0,
  };
  let totalKeyCharacters = 0;
  const buildKeyedRecords = (sourceRecords, keyColumns, side) =>
      sourceRecords.map((record) => {
        const original = keyColumns.map(
            (keyColumn) => record.values[keyColumn],
          ),
          normalizedValues = original.map((originalValue) =>
            normalizeKeyText(originalValue, normalization),
          );
        if (
          ((totalKeyCharacters += normalizedValues.reduce(
            (total, normalizedValue) =>
              total +
              (typeof normalizedValue == "string" ? normalizedValue.length : 0),
            0,
          )),
          totalKeyCharacters > 2e6)
        )
          throw Error("Normalized key text exceeds 2,000,000 characters.");
        serializeTypedKeyValues(original) !==
          serializeTypedKeyValues(normalizedValues) &&
          diagnostics.normalizedKeys.push({
            side: side,
            locator: record.locator,
            original: original,
            normalized: normalizedValues,
          });
        const isBlank = normalizedValues.some(isNullOrEmptyValue),
          fuzzyKeyValue = fuzzy ? normalizedValues[fuzzy.key] : undefined,
          isNonTextFuzzyKey =
            !!fuzzy && !isBlank && typeof fuzzyKeyValue != "string";
        isNonTextFuzzyKey &&
          diagnostics.exceptions.push({
            side: side,
            locator: record.locator,
            reason:
              "Fuzzy matching requires text in the selected key. This record is unmatched; no numeric or Boolean coercion was applied.",
          });
        const chars =
          typeof fuzzyKeyValue == "string" ? Array.from(fuzzyKeyValue) : [];
        if (fuzzy && chars.length > fuzzyMatchLimits.fuzzyCharacters)
          throw Error(
            "Fuzzy keys are limited to 512 Unicode characters. Shorten the reviewed key or use exact matching.",
          );
        return {
          record: record,
          original: original,
          values: normalizedValues,
          valid: !isBlank && !isNonTextFuzzyKey,
          isBlank: isBlank,
          key: serializeTypedKeyValues(normalizedValues),
          block: serializeTypedKeyValues(
            normalizedValues.filter(
              (X, valueIndex) => valueIndex !== fuzzy?.key,
            ),
          ),
          chars: chars,
        };
      }),
    leftKeyedRecords = buildKeyedRecords(records, leftKeyColumns, "left"),
    rightKeyedRecords = buildKeyedRecords(
      rightRecords,
      rightKeyColumns,
      "right",
    ),
    groupByKey = (keyedRecords, sideName) => {
      const groupsByKey = new Map();
      for (const keyedRecord of keyedRecords)
        keyedRecord.isBlank ||
          groupsByKey.set(keyedRecord.key, [
            ...(groupsByKey.get(keyedRecord.key) ?? []),
            keyedRecord,
          ]);
      return (
        (diagnostics.normalizationCollisionGroups[sideName] = [
          ...groupsByKey.values(),
        ].filter(
          (keyGroup) =>
            new Set(
              keyGroup.map((groupMember) =>
                serializeTypedKeyValues(groupMember.original),
              ),
            ).size > 1,
        ).length),
        groupsByKey
      );
    },
    leftKeyGroups = groupByKey(leftKeyedRecords, "left"),
    rightKeyGroups = groupByKey(rightKeyedRecords, "right"),
    byLeft = new Map(),
    rightMatchCounts = new Map(),
    bestLeftScores = new Map(),
    bestRightScores = new Map(),
    addComparisonWork = (workUnits) => {
      if (
        ((diagnostics.comparisonWork += workUnits),
        diagnostics.comparisonWork > fuzzyMatchLimits.comparisonWork)
      )
        throw Error(
          "Fuzzy comparison exceeds 10,000,000 work units. Narrow the populations or add an exact key before previewing.",
        );
    };
  function textSimilarity(leftChars, rightChars) {
    if (
      (addComparisonWork(leftChars.length + rightChars.length + 1),
      leftChars.join("") === rightChars.join(""))
    )
      return 1;
    addComparisonWork(leftChars.length * rightChars.length);
    let previousRow = Array.from(
      { length: rightChars.length + 1 },
      (R, columnIndex) => columnIndex,
    );
    for (
      let leftPosition = 1;
      leftPosition <= leftChars.length;
      leftPosition++
    ) {
      const currentRow = [leftPosition];
      for (
        let rightPosition = 1;
        rightPosition <= rightChars.length;
        rightPosition++
      )
        currentRow[rightPosition] = Math.min(
          previousRow[rightPosition] + 1,
          currentRow[rightPosition - 1] + 1,
          previousRow[rightPosition - 1] +
            (leftChars[leftPosition - 1] === rightChars[rightPosition - 1]
              ? 0
              : 1),
        );
      previousRow = currentRow;
    }
    return (
      1 -
      previousRow[rightChars.length] /
        Math.max(leftChars.length, rightChars.length)
    );
  }
  const rightRecordsByKey = new Map();
  for (const rightKeyedRecord of rightKeyedRecords)
    if (rightKeyedRecord.valid) {
      const joinKey = fuzzy ? rightKeyedRecord.block : rightKeyedRecord.key;
      rightRecordsByKey.set(joinKey, [
        ...(rightRecordsByKey.get(joinKey) ?? []),
        rightKeyedRecord,
      ]);
    }
  for (const leftKeyedRecord of leftKeyedRecords) {
    if (!leftKeyedRecord.valid) continue;
    const candidates =
      rightRecordsByKey.get(
        fuzzy ? leftKeyedRecord.block : leftKeyedRecord.key,
      ) ?? [];
    for (const candidate of candidates) {
      const score = fuzzy
        ? textSimilarity(leftKeyedRecord.chars, candidate.chars)
        : 1;
      if (
        (bestLeftScores.set(
          leftKeyedRecord.record.row,
          Math.max(bestLeftScores.get(leftKeyedRecord.record.row) ?? -1, score),
        ),
        bestRightScores.set(
          candidate.record.row,
          Math.max(bestRightScores.get(candidate.record.row) ?? -1, score),
        ),
        fuzzy && score < fuzzy.threshold)
      )
        continue;
      const explanation = fuzzy
          ? `Text similarity ${score.toFixed(6)} meets threshold ${fuzzy.threshold}; other key pairs match exactly after selected normalization.`
          : "Exact typed keys match after selected normalization.",
        existingMatches = byLeft.get(leftKeyedRecord.record.row) ?? [];
      if (
        (existingMatches.push({
          record: candidate.record,
          score: score,
          explanation: explanation,
        }),
        byLeft.set(leftKeyedRecord.record.row, existingMatches),
        rightMatchCounts.set(
          candidate.record.row,
          (rightMatchCounts.get(candidate.record.row) ?? 0) + 1,
        ),
        diagnostics.matches.push({
          left: leftKeyedRecord.record.locator,
          right: candidate.record.locator,
          score: score,
          explanation: explanation,
        }),
        diagnostics.matches.length > fuzzyMatchLimits.matchPairs)
      )
        throw Error(
          "Matching duplicate keys expand beyond 1,000 result rows or reviewable match pairs. Narrow or reconcile the source keys first; no pairs were truncated.",
        );
    }
  }
  ((diagnostics.ambiguousLeftRows = [...byLeft.values()].filter(
    (leftMatches) => leftMatches.length > 1,
  ).length),
    (diagnostics.ambiguousRightRows = [...rightMatchCounts.values()].filter(
      (rightMatchCount) => rightMatchCount > 1,
    ).length));
  const serializeDisplayedKey = (keyValues) =>
      JSON.stringify(keyValues.map((keyValue) => String(keyValue))),
    leftDisplayedKeys = new Set(
      leftKeyedRecords
        .filter((leftKeyed) => leftKeyed.valid)
        .map((keyedLeft) => serializeDisplayedKey(keyedLeft.values)),
    ),
    rightDisplayedKeys = new Set(
      rightKeyedRecords
        .filter((rightKeyed) => rightKeyed.valid)
        .map((keyedRight) => serializeDisplayedKey(keyedRight.values)),
    ),
    explainUnmatched = (keyedEntry, bestScore, displayedKeys) =>
      keyedEntry.isBlank
        ? "Blank join key; blanks do not match."
        : keyedEntry.valid
          ? fuzzy
            ? bestScore === undefined
              ? "No eligible text candidate with matching exact keys."
              : `Best eligible similarity ${bestScore.toFixed(6)} is below threshold ${fuzzy.threshold}.`
            : displayedKeys.has(serializeDisplayedKey(keyedEntry.values))
              ? "No exact typed composite-key match; equivalent displayed keys have mismatched types. No coercion was applied."
              : "No exact typed composite-key match."
          : "Non-text fuzzy key; no implicit type conversion.",
    unmatchedLeft = leftKeyedRecords
      .filter((leftKeyedItem) => !byLeft.has(leftKeyedItem.record.row))
      .map((leftEntry) => ({
        ...leftEntry.record,
        unmatchedReason: explainUnmatched(
          leftEntry,
          bestLeftScores.get(leftEntry.record.row),
          rightDisplayedKeys,
        ),
      })),
    unmatchedRight = rightKeyedRecords
      .filter(
        (rightKeyedItem) => !rightMatchCounts.has(rightKeyedItem.record.row),
      )
      .map((rightEntry) => ({
        ...rightEntry.record,
        unmatchedReason: explainUnmatched(
          rightEntry,
          bestRightScores.get(rightEntry.record.row),
          leftDisplayedKeys,
        ),
      }));
  for (const [side, unmatchedRows] of [
    ["left", unmatchedLeft],
    ["right", unmatchedRight],
  ])
    for (const unmatchedEntry of unmatchedRows)
      diagnostics.exceptions.some(
        (exception) =>
          exception.side === side &&
          exception.locator === unmatchedEntry.locator,
      ) ||
        diagnostics.exceptions.push({
          side: side,
          locator: unmatchedEntry.locator,
          reason: unmatchedEntry.unmatchedReason,
        });
  const rightToLeftMatches = new Map(),
    leftToRightMatches = new Map();
  for (const match of diagnostics.matches)
    (leftToRightMatches.set(match.left, [
      ...(leftToRightMatches.get(match.left) ?? []),
      match.right,
    ]),
      rightToLeftMatches.set(match.right, [
        ...(rightToLeftMatches.get(match.right) ?? []),
        match.left,
      ]));
  const visitedLocators = new Set();
  let manyToManyKeys = 0,
    expandedMatchRows = 0;
  for (const startLocator of leftToRightMatches.keys())
    if (!visitedLocators.has(startLocator)) {
      const componentLeft = new Set(),
        componentRight = new Set(),
        pendingLocators = [startLocator];
      let componentPairs = 0;
      for (; pendingLocators.length;) {
        const currentLocator = pendingLocators.pop();
        if (visitedLocators.has(currentLocator)) continue;
        (visitedLocators.add(currentLocator),
          componentLeft.add(currentLocator));
        const matchedRights = leftToRightMatches.get(currentLocator);
        componentPairs += matchedRights.length;
        for (const rightLocator of matchedRights) {
          componentRight.add(rightLocator);
          for (const leftLocator of rightToLeftMatches.get(rightLocator))
            visitedLocators.has(leftLocator) ||
              pendingLocators.push(leftLocator);
        }
      }
      componentLeft.size > 1 &&
        componentRight.size > 1 &&
        (manyToManyKeys++,
        (expandedMatchRows += Math.max(
          0,
          componentPairs - Math.max(componentLeft.size, componentRight.size),
        )));
    }
  return {
    byLeft: byLeft,
    unmatchedLeft: unmatchedLeft,
    unmatchedRight: unmatchedRight,
    diagnostics: diagnostics,
    duplicateLeftKeys: [...leftKeyGroups.values()].filter(
      (leftGroup) => leftGroup.length > 1,
    ).length,
    duplicateRightKeys: [...rightKeyGroups.values()].filter(
      (rightGroup) => rightGroup.length > 1,
    ).length,
    matchedLeftRows: byLeft.size,
    matchedRightRows: rightMatchCounts.size,
    matchPairs: diagnostics.matches.length,
    manyToManyKeys: manyToManyKeys,
    expandedMatchRows: expandedMatchRows,
  };
}
const consolidationBudgets = {
  sourceCells: 2e3,
  sourceRows: 500,
  sourceColumns: 50,
  outputRows: 1e3,
  outputCells: 2e4,
  keyColumns: 8,
  sourceFormulas: 50,
  scannedFormulas: 1e3,
  dependencyCells: 2e3,
  headerCharacters: 128,
  textCharacters: 32767,
  totalTextCharacters: 1e6,
};

function parseSourceRectangle(source) {
  (!hasOnlyAllowedKeys(source, ["sheetId", "range"]) ||
    typeof source.sheetId != "string" ||
    typeof source.range != "string" ||
    !/^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/.test(
      source.range,
    )) &&
    failWithMessage(
      "Choose a worksheet and one explicit headed cell rectangle for each source.",
    );
  const normalizedRange = source.range.replace(/\$/g, "").toUpperCase(),
    rangeParts = normalizedRange.split(":"),
    startCell = _t(rangeParts[0]),
    endCell = _t(rangeParts[1] ?? rangeParts[0]);
  (startCell.row > endCell.row ||
    startCell.col > endCell.col ||
    endCell.row >= workbookLimits.maxRows ||
    endCell.col >= workbookLimits.maxCols) &&
    failWithMessage(
      "A source rectangle is reversed or outside Excel's address bounds.",
    );
  const rowCount = endCell.row - startCell.row + 1,
    columnCount = endCell.col - startCell.col + 1;
  return (
    (rowCount > consolidationBudgets.sourceRows + 1 ||
      columnCount > consolidationBudgets.sourceColumns ||
      rowCount * columnCount > consolidationBudgets.sourceCells) &&
      failWithMessage(
        "Each source supports at most 500 data rows, 50 columns and 2,000 cells including headers.",
      ),
    {
      sheetId: source.sheetId,
      range: `${formatCellAddress(startCell.row, startCell.col)}:${formatCellAddress(endCell.row, endCell.col)}`,
    }
  );
}
function isCellVisible(sheet, firstRow, firstCol) {
  return (
    !sheet.hidden &&
    !sheet.veryHidden &&
    !cn(sheet, "row").some(
      (hiddenRowRange) =>
        firstRow >= hiddenRowRange.start && firstRow <= hiddenRowRange.end,
    ) &&
    !cn(sheet, "column").some(
      (hiddenColumnRange) =>
        firstCol >= hiddenColumnRange.start &&
        firstCol <= hiddenColumnRange.end,
    )
  );
}
function inspectOutputNameSafety(workbook, outputSheetName) {
  let totalFormulaCharacters = 0;
  const scanFormula = (formulaText, sheetId, address = "A1") => {
    if (!formulaText) return;
    ((totalFormulaCharacters += formulaText.length),
      totalFormulaCharacters > consolidationBudgets.totalTextCharacters &&
        failWithMessage(
          "Output-name safety inspection exceeds 1,000,000 formula characters. Use a smaller reviewed workbook.",
        ));
    const formulaWithoutStrings = formulaText.replace(/"(?:[^"]|"")*"/g, "");
    let rewrittenFormula = "",
      scanPosition = 0;
    for (; scanPosition < formulaWithoutStrings.length;) {
      if (formulaWithoutStrings[scanPosition] === "'") {
        const quoteStart = scanPosition++;
        for (; scanPosition < formulaWithoutStrings.length;)
          if (formulaWithoutStrings[scanPosition++] === "'")
            if (formulaWithoutStrings[scanPosition] === "'") scanPosition++;
            else break;
        rewrittenFormula += formulaWithoutStrings.slice(
          quoteStart,
          scanPosition,
        );
        continue;
      }
      const referenceStart = scanPosition,
        structuredMatch = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(
          formulaWithoutStrings.slice(scanPosition),
        );
      if (structuredMatch) {
        scanPosition += structuredMatch[0].length;
        let bracketDepth = 1;
        for (; scanPosition < formulaWithoutStrings.length && bracketDepth;)
          (formulaWithoutStrings[scanPosition] === "[" && bracketDepth++,
            formulaWithoutStrings[scanPosition] === "]" && bracketDepth--,
            scanPosition++);
        const structuredReferenceText = formulaWithoutStrings.slice(
          referenceStart,
          scanPosition,
        );
        rewrittenFormula +=
          !bracketDepth &&
          !/^[A-Za-z0-9_!]/.test(formulaWithoutStrings.slice(scanPosition)) &&
          rewriteStructuredReferences(
            workbook,
            sheetId,
            structuredReferenceText,
            address,
          )
            ? "0"
            : structuredReferenceText;
      } else rewrittenFormula += formulaWithoutStrings[scanPosition++];
    }
    for (const referenceMatch of rewrittenFormula.matchAll(
      /('(?:[^']|'')*'|[A-Za-z_.\d\u007F-\uFFFF]+)\s*!/g,
    )) {
      const sheetToken = referenceMatch[1];
      (sheetToken.startsWith("'")
        ? sheetToken.slice(1, -1).replace(/''/g, "'")
        : sheetToken
      ).toLowerCase() === outputSheetName.toLowerCase() &&
        failWithMessage(
          "The output name would reconnect an existing unresolved worksheet reference. Choose a different name.",
        );
    }
  };
  for (const sheet of workbook.sheets) {
    for (const [cellAddress, cell] of Object.entries(sheet.cells))
      isFormulaInputValue(cell) &&
        scanFormula(String(cell.raw), sheet.id, cellAddress);
    for (const conditionalFormat of sheet.conditionalFormats ?? [])
      conditionalFormat.kind === "formula" &&
        scanFormula(
          conditionalFormat.formula,
          sheet.id,
          conditionalFormat.range.split(":")[0],
        );
    for (const table of sheet.tables ?? [])
      for (const [columnIndex, column] of table.columns.entries()) {
        const rangeAddress = parseRangeAddress(table.range);
        scanFormula(
          column.calculatedFormula,
          sheet.id,
          formatCellAddress(
            rangeAddress.firstRow + 1,
            rangeAddress.firstCol + columnIndex,
          ),
        );
      }
  }
  for (const definedName of workbook.definedNames ?? [])
    scanFormula(
      definedName.refersTo,
      definedName.scopeSheetId ?? workbook.sheets[0].id,
    );
}
function readSourceRectangles(workbookInput, sourceRectangles, options = {}) {
  (!Array.isArray(sourceRectangles) ||
    !sourceRectangles.length ||
    sourceRectangles.length > 8) &&
    failWithMessage("Read one to eight explicit source rectangles.");
  const parsedRectangles = sourceRectangles.map(parseSourceRectangle);
  (options.includeHidden !== undefined &&
    typeof options.includeHidden != "boolean" &&
    failWithMessage("Include hidden must be boolean."),
    options.allowedSheetIds !== undefined &&
      (!Array.isArray(options.allowedSheetIds) ||
        options.allowedSheetIds.some(
          (allowedSheetId) =>
            typeof allowedSheetId != "string" ||
            !workbookInput.sheets.some((sheet) => sheet.id === allowedSheetId),
        )) &&
      failWithMessage("The host scope contains an unknown worksheet."));
  const scope = {
      includeHidden: !!options.includeHidden,
      allowedSheetIds: [
        ...new Set(
          options.allowedSheetIds ??
            workbookInput.sheets.map((sheet) => sheet.id),
        ),
      ].sort(),
    },
    sourceSheets = parsedRectangles.map((parsedRectangle) => {
      const sheet = workbookInput.sheets.find(
        (sheet) => sheet.id === parsedRectangle.sheetId,
      );
      (!sheet || !scope.allowedSheetIds.includes(sheet.id)) &&
        failWithMessage("A source is outside the authorized host context.");
      const rangeAddress = parseRangeAddress(parsedRectangle.range);
      if (!scope.includeHidden)
        for (
          let firstRow = rangeAddress.firstRow;
          firstRow <= rangeAddress.lastRow;
          firstRow++
        )
          for (
            let firstCol = rangeAddress.firstCol;
            firstCol <= rangeAddress.lastCol;
            firstCol++
          )
            isCellVisible(sheet, firstRow, firstCol) ||
              failWithMessage(
                "A source includes hidden cells. Explicitly include hidden records before previewing.",
              );
      return (
        sheet.merges?.some((merge) => {
          const mergeAddress = parseRangeAddress(merge);
          return (
            mergeAddress.firstRow <= rangeAddress.lastRow &&
            mergeAddress.lastRow >= rangeAddress.firstRow &&
            mergeAddress.firstCol <= rangeAddress.lastCol &&
            mergeAddress.lastCol >= rangeAddress.firstCol
          );
        }) &&
          failWithMessage(
            "Merged source cells are ambiguous; use an unmerged headed population.",
          ),
        sheet
      );
    }),
    capturedKeys = new Set(),
    formulaKeys = new Set(),
    formulaSourceCells = [];
  for (const [rectangleIndex, rectangle] of parsedRectangles.entries()) {
    const sheet = sourceSheets[rectangleIndex],
      rangeAddress = parseRangeAddress(rectangle.range);
    for (
      let firstRow = rangeAddress.firstRow;
      firstRow <= rangeAddress.lastRow;
      firstRow++
    )
      for (
        let firstCol = rangeAddress.firstCol;
        firstCol <= rangeAddress.lastCol;
        firstCol++
      ) {
        const cellAddress = formatCellAddress(firstRow, firstCol),
          compositeKey = buildCompositeKey(sheet.id, cellAddress);
        (capturedKeys.add(compositeKey),
          isFormulaInputValue(sheet.cells[cellAddress]) &&
            !formulaKeys.has(compositeKey) &&
            (formulaKeys.add(compositeKey),
            formulaSourceCells.push({ sheet: sheet, address: cellAddress })));
      }
  }
  formulaSourceCells.length > consolidationBudgets.sourceFormulas &&
    failWithMessage(
      "Source capture supports at most 50 formula source cells. Use a reviewed values snapshot for a larger calculated population.",
    );
  for (const sourceSheet of new Set(sourceSheets)) {
    const blankCellPositions = [];
    for (const capturedKey of capturedKeys) {
      const [keySheetId, keyAddress] = JSON.parse(capturedKey);
      keySheetId === sourceSheet.id &&
        (!sourceSheet.cells[keyAddress] ||
          sourceSheet.cells[keyAddress].raw === null) &&
        blankCellPositions.push(_t(keyAddress));
    }
    if (!blankCellPositions.length) continue;
    const formulaAnchorCandidates = Object.entries(sourceSheet.cells).filter(
      ([candidateAddress, candidateCell]) =>
        isFormulaInputValue(candidateCell) &&
        !formulaKeys.has(buildCompositeKey(sourceSheet.id, candidateAddress)),
    );
    formulaAnchorCandidates.length > consolidationBudgets.scannedFormulas &&
      failWithMessage(
        "Blank source cells require a bounded spill-ambiguity check; this sheet exceeds 1,000 possible formula anchors.",
      );
    for (const [anchorAddress] of formulaAnchorCandidates) {
      const anchorPosition = _t(anchorAddress);
      blankCellPositions.some(
        (blankPosition) =>
          blankPosition.row >= anchorPosition.row &&
          blankPosition.col >= anchorPosition.col &&
          (blankPosition.row - anchorPosition.row + 1) *
            (blankPosition.col - anchorPosition.col + 1) <=
            1e4,
      ) &&
        failWithMessage(
          "A blank source cell could be a spill child of a formula outside the source. Use a reviewed values snapshot or include and validate the scalar source formula.",
        );
    }
  }
  const volatileFunctionNames = new Set([
      ...matrixResultFunctionNames,
      "ISFORMULA",
      "INDEX",
      "CHOOSE",
      "INDIRECT",
      "OFFSET",
      "LET",
      "LAMBDA",
      "RAND",
      "RANDBETWEEN",
      "NOW",
      "TODAY",
      "CELL",
      "INFO",
      "RTD",
      "WEBSERVICE",
      "STOCKHISTORY",
    ]),
    workbook = {
      ...workbookInput,
      sheets: workbookInput.sheets.map((sheet) =>
        sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
      ),
    },
    dependencyKeys = new Set();
  for (const formulaSourceCell of formulaSourceCells) {
    const dependencyReport = inspectFormulaDependencies(workbook, {
        sheetId: formulaSourceCell.sheet.id,
        address: formulaSourceCell.address,
        direction: "precedents",
        depth: 8,
        includeHidden: scope.includeHidden,
        allowedSheetIds: scope.allowedSheetIds,
        limits: { maxFormulaCells: 1e3, maxNodes: 500, maxReferences: 2e3 },
      }),
      targetAndNodes = [dependencyReport.target, ...dependencyReport.nodes],
      nodeKeys = new Set(targetAndNodes.map((event) => event.key));
    (dependencyReport.coverage.truncated ||
      dependencyReport.coverage.depthLimited ||
      dependencyReport.cycles.length ||
      dependencyReport.issues.some((issue) => nodeKeys.has(issue.from))) &&
      failWithMessage(
        "A source formula has incomplete, cyclic, hidden, unauthorized or unsupported dependencies.",
      );
    for (const dependencyNode of targetAndNodes) {
      (capturedKeys.add(dependencyNode.key),
        dependencyKeys.add(dependencyNode.key),
        dependencyKeys.size > consolidationBudgets.dependencyCells &&
          failWithMessage(
            "Combined source formula dependencies exceed 2,000 cells.",
          ));
      const sheet = workbookInput.sheets.find(
          (sheet) => sheet.id === dependencyNode.sheetId,
        ),
        cell = sheet.cells[dependencyNode.address];
      if (
        ((!cell || cell.raw === null) &&
          failWithMessage(
            "Formula dependencies must be populated static cells; blank or possible spill-child dependencies are outside this subset.",
          ),
        isFormulaInputValue(cell))
      ) {
        const resolvedFormula = resolveDefinedNameValue(
          String(cell.raw),
          workbookInput,
          sheet.id,
        ).replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, "");
        ([...resolvedFormula.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)]
          .map((functionMatch) =>
            functionMatch[1]
              .toUpperCase()
              .replace(/^_XLFN\./, "")
              .replace(/^_XLWS\./, ""),
          )
          .some((functionName) => volatileFunctionNames.has(functionName)) ||
          /[{}#]/.test(resolvedFormula)) &&
          failWithMessage(
            "Source formulas must have deterministic scalar dependencies; volatile, dynamic and array formulas are excluded.",
          );
      }
    }
  }
  const capturedWorkbook = {
      ...workbookInput,
      sheets: workbookInput.sheets
        .filter((sheet) => scope.allowedSheetIds.includes(sheet.id))
        .map((capturedSheet) => ({
          ...capturedSheet,
          cells: Object.fromEntries(
            Object.entries(capturedSheet.cells).filter(([keptAddress]) =>
              capturedKeys.has(
                buildCompositeKey(capturedSheet.id, keptAddress),
              ),
            ),
          ),
        })),
    },
    workbookWithSettings = calculateWorkbookWithSettings(capturedWorkbook),
    errorValuePattern =
      /^#(?:DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|CYCLE!|LIMIT!|BLOCKED!|ERROR!)$/;
  let combinedTextLength = 0;
  const grids = parsedRectangles.map((source, sourceIndex) => {
    const gridSheet = sourceSheets[sourceIndex],
      rangeAddress = parseRangeAddress(source.range),
      values = [];
    for (
      let firstRow = rangeAddress.firstRow;
      firstRow <= rangeAddress.lastRow;
      firstRow++
    ) {
      const rowValues = [];
      for (
        let firstCol = rangeAddress.firstCol;
        firstCol <= rangeAddress.lastCol;
        firstCol++
      ) {
        const cellAddress = formatCellAddress(firstRow, firstCol),
          cell = gridSheet.cells[cellAddress],
          cellValue = isFormulaInputValue(cell)
            ? workbookWithSettings(gridSheet.id, cellAddress)
            : (cell?.raw ?? null);
        (((isFormulaInputValue(cell) &&
          typeof cellValue == "string" &&
          errorValuePattern.test(cellValue)) ||
          (typeof cellValue == "number" && !Number.isFinite(cellValue)) ||
          (typeof cellValue == "string" &&
            cellValue.length > consolidationBudgets.textCharacters)) &&
          failWithMessage(
            "A source contains a calculation error, nonfinite number or oversized text. Repair it before capturing source values.",
          ),
          (combinedTextLength +=
            typeof cellValue == "string" ? cellValue.length : 0),
          combinedTextLength > consolidationBudgets.totalTextCharacters &&
            failWithMessage(
              "Combined source text exceeds 1,000,000 characters.",
            ),
          rowValues.push(cellValue));
      }
      values.push(rowValues);
    }
    return { source: source, sheetName: gridSheet.name, values: values };
  });
  for (const dependencyKey of dependencyKeys) {
    const [dependencySheetId, dependencyAddress] = JSON.parse(dependencyKey);
    workbookWithSettings.getSpillRange(dependencySheetId, dependencyAddress) &&
      failWithMessage(
        "Source formula dependencies produced an array; use a reviewed scalar population.",
      );
  }
  workbookWithSettings.getOperationalErrors().length &&
    failWithMessage(
      "A source encountered an operational calculation failure, including an error hidden by a formula handler.",
    );
  const dependencyCells = [...capturedKeys].sort().map((cellKey) => {
    const [sheetId, address] = JSON.parse(cellKey);
    return {
      sheetId: sheetId,
      address: address,
      cell:
        workbookInput.sheets.find((sheet) => sheet.id === sheetId)?.cells[
          address
        ] ?? null,
    };
  });
  return { grids: grids, dependencyCells: dependencyCells };
}
function readSourceHeaders(workbookInput, sourceRectangles, options = {}) {
  const { grids, dependencyCells } = readSourceRectangles(
    workbookInput,
    sourceRectangles,
    options,
  );
  return {
    snapshots: grids.map(({ source, sheetName, values }) => {
      const rangeAddress = parseRangeAddress(source.range);
      values[0].some(
        (headerValue) =>
          typeof headerValue != "string" ||
          !headerValue.trim() ||
          headerValue.length > consolidationBudgets.headerCharacters,
      ) &&
        failWithMessage(
          "Every source column needs a nonblank text header of at most 128 characters.",
        );
      const headers = values[0];
      return (
        new Set(headers.map((header) => header.toLowerCase())).size !==
          headers.length &&
          failWithMessage(
            "Source headers must be unique, including case variants.",
          ),
        {
          source,
          sheetName,
          headers: headers,
          records: values.slice(1).map((values, recordIndex) => ({
            row: rangeAddress.firstRow + recordIndex + 2,
            locator: `'${sheetName.replace(/'/g, "''")}'!${formatCellAddress(rangeAddress.firstRow + recordIndex + 1, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.firstRow + recordIndex + 1, rangeAddress.lastCol)}`,
            values: values,
          })),
        }
      );
    }),
    dependencyCells,
  };
}
function buildConsolidationOutput(workbook, request, options = {}) {
  (options.readOnly && failWithMessage("This workbook is read-only."),
    assertWorkbookStructureEditable(workbook),
    (!hasOnlyAllowedKeys(request, [
      "left",
      "right",
      "mode",
      "keys",
      "outputSheetName",
      "normalization",
      "fuzzy",
    ]) ||
      ![
        "append",
        "inner",
        "left",
        "right",
        "full",
        "leftAnti",
        "rightAnti",
      ].includes(request.mode)) &&
      failWithMessage("Choose append or a supported join operation."),
    (typeof request.outputSheetName != "string" ||
      request.outputSheetName !== request.outputSheetName.trim() ||
      !request.outputSheetName ||
      request.outputSheetName.length > 31 ||
      /[\\/\[\]*?:\x00-\x1f]/.test(request.outputSheetName) ||
      /^'|'$/.test(request.outputSheetName)) &&
      failWithMessage(
        "Output sheet name must be a valid, unpadded name of 1–31 characters.",
      ),
    workbook.sheets.some(
      (sheet) =>
        sheet.name.toLowerCase() === request.outputSheetName.toLowerCase(),
    ) &&
      failWithMessage(
        "The output sheet name already exists. Choose a new sheet name.",
      ),
    inspectOutputNameSafety(workbook, request.outputSheetName),
    options.includeHidden !== undefined &&
      typeof options.includeHidden != "boolean" &&
      failWithMessage("Include hidden must be boolean."),
    options.allowedSheetIds !== undefined &&
      (!Array.isArray(options.allowedSheetIds) ||
        options.allowedSheetIds.some(
          (allowedSheetId) =>
            typeof allowedSheetId != "string" ||
            !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
        )) &&
      failWithMessage("The host scope contains an unknown worksheet."));
  const context = {
      includeHidden: !!options.includeHidden,
      allowedSheetIds: [
        ...new Set(
          options.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
        ),
      ].sort(),
    },
    parsedRequest = {
      left: parseSourceRectangle(request.left),
      right: parseSourceRectangle(request.right),
      mode: request.mode,
      outputSheetName: request.outputSheetName,
    };
  parsedRequest.mode === "append"
    ? ((request.normalization !== undefined || request.fuzzy !== undefined) &&
        failWithMessage(
          "Append aligns headers; key normalization and fuzzy matching apply only to joins.",
        ),
      request.keys !== undefined &&
        (!Array.isArray(request.keys) || request.keys.length) &&
        failWithMessage(
          "Append matches headers; it does not accept join keys.",
        ))
    : ((!Array.isArray(request.keys) ||
        !request.keys.length ||
        request.keys.length > consolidationBudgets.keyColumns ||
        request.keys.some(
          (key) =>
            !hasOnlyAllowedKeys(key, ["left", "right"]) ||
            typeof key.left != "string" ||
            typeof key.right != "string",
        )) &&
        failWithMessage("Choose from one to eight pairs of join headers."),
      (parsedRequest.keys = request.keys.map((key) => ({ ...key }))),
      Object.assign(
        parsedRequest,
        parseKeyNormalization(
          request.normalization,
          request.fuzzy,
          parsedRequest.keys.length,
        ),
      ),
      (new Set(parsedRequest.keys.map((key) => key.left)).size !==
        parsedRequest.keys.length ||
        new Set(parsedRequest.keys.map((key) => key.right)).size !==
          parsedRequest.keys.length) &&
        failWithMessage("A join header cannot be repeated in the key pairs."));
  const { snapshots } = readSourceHeaders(
      workbook,
      [parsedRequest.left, parsedRequest.right],
      options,
    ),
    [left, right] = snapshots,
    rows = [],
    unmatchedLeft = [],
    unmatchedRight = [],
    keyProfiles = [];
  let matching,
    headers,
    matchedLeftRows = 0,
    matchedRightRows = 0,
    matchPairs = 0,
    duplicateLeftKeys = 0,
    duplicateRightKeys = 0,
    manyToManyKeys = 0,
    expandedMatchRows = 0;
  if (parsedRequest.mode === "append") {
    ((left.headers.length !== right.headers.length ||
      left.headers.some((header) => !right.headers.includes(header))) &&
      failWithMessage(
        "Append requires exactly matching header names on both sources; column order may differ.",
      ),
      (headers = [...left.headers, "Source left row", "Source right row"]));
    const rightHeaderIndexes = left.headers.map((header) =>
      right.headers.indexOf(header),
    );
    (left.records.forEach((record) =>
      rows.push({
        values: [...record.values],
        left: record.locator,
        right: null,
      }),
    ),
      right.records.forEach((record) =>
        rows.push({
          values: rightHeaderIndexes.map(
            (rightColumnIndex) => record.values[rightColumnIndex],
          ),
          left: null,
          right: record.locator,
        }),
      ));
  } else {
    headers = [
      ...left.headers.map((header) => `Left: ${header}`),
      ...right.headers.map((header) => `Right: ${header}`),
      "Source left row",
      "Source right row",
    ];
    const leftKeyIndexes = parsedRequest.keys.map((key) =>
        left.headers.indexOf(key.left),
      ),
      rightKeyIndexes = parsedRequest.keys.map((key) =>
        right.headers.indexOf(key.right),
      );
    (leftKeyIndexes.includes(-1) || rightKeyIndexes.includes(-1)) &&
      failWithMessage(
        "Every join key must name an existing source header exactly.",
      );
    const typeProfile = (records, columnIndex) =>
      [
        ...new Set(
          records.map((profiledRecord) =>
            profiledRecord.values[columnIndex] === null ||
            profiledRecord.values[columnIndex] === ""
              ? "blank"
              : typeof profiledRecord.values[columnIndex],
          ),
        ),
      ].sort();
    parsedRequest.keys.forEach((key, keyIndex) =>
      keyProfiles.push({
        leftHeader: key.left,
        rightHeader: key.right,
        leftTypes: typeProfile(left.records, leftKeyIndexes[keyIndex]),
        rightTypes: typeProfile(right.records, rightKeyIndexes[keyIndex]),
      }),
    );
    const keyMatchResult = matchKeyPopulations(
      left.records,
      right.records,
      leftKeyIndexes,
      rightKeyIndexes,
      parsedRequest.normalization,
      parsedRequest.fuzzy,
    );
    ((matching = keyMatchResult.diagnostics),
      ({
        matchedLeftRows: matchedLeftRows,
        matchedRightRows: matchedRightRows,
        matchPairs: matchPairs,
        duplicateLeftKeys: duplicateLeftKeys,
        duplicateRightKeys: duplicateRightKeys,
        manyToManyKeys: manyToManyKeys,
        expandedMatchRows: expandedMatchRows,
      } = keyMatchResult),
      unmatchedLeft.push(...keyMatchResult.unmatchedLeft),
      unmatchedRight.push(...keyMatchResult.unmatchedRight));
    for (const record of left.records) {
      const leftMatches = keyMatchResult.byLeft.get(record.row);
      if (leftMatches) {
        if (
          parsedRequest.mode !== "leftAnti" &&
          parsedRequest.mode !== "rightAnti"
        )
          for (const rightMatch of leftMatches)
            rows.push({
              values: [...record.values, ...rightMatch.record.values],
              left: record.locator,
              right: rightMatch.record.locator,
              ...(parsedRequest.fuzzy || parsedRequest.normalization
                ? {
                    match: {
                      score: rightMatch.score,
                      explanation: rightMatch.explanation,
                    },
                  }
                : {}),
            });
      } else
        ["left", "full", "leftAnti"].includes(parsedRequest.mode) &&
          rows.push({
            values: [...record.values, ...right.headers.map(() => null)],
            left: record.locator,
            right: null,
          });
    }
    if (["right", "full", "rightAnti"].includes(parsedRequest.mode))
      for (const unmatchedRecord of unmatchedRight)
        rows.push({
          values: [...left.headers.map(() => null), ...unmatchedRecord.values],
          left: null,
          right: unmatchedRecord.locator,
        });
  }
  (new Set(headers.map((header) => header.toLowerCase())).size !==
    headers.length &&
    failWithMessage(
      "A source header conflicts with a generated provenance column name.",
    ),
    (rows.length > consolidationBudgets.outputRows ||
      (rows.length + 1) * headers.length > consolidationBudgets.outputCells) &&
      failWithMessage(
        "The output exceeds 1,000 data rows or 20,000 cells including headers and provenance.",
      ),
    rows.reduce(
      (total, outputRow) =>
        total +
        outputRow.values.reduce(
          (rowTotal, value) =>
            rowTotal + (typeof value == "string" ? value.length : 0),
          0,
        ),
      headers.join("").length,
    ) > consolidationBudgets.totalTextCharacters &&
      failWithMessage("Expanded output text exceeds 1,000,000 characters."),
    workbook.sheets.reduce(
      (storedCells, existingSheet) =>
        storedCells + Object.keys(existingSheet.cells).length,
      0,
    ) +
      (rows.length + 1) * headers.length >
      workbookLimits.maxCells &&
      failWithMessage(
        "The new output would exceed the workbook's stored-cell budget.",
      ));
  const worksheetRecord = createWorksheetRecord(parsedRequest.outputSheetName),
    snapshot = JSON.stringify(workbook),
    existingCombineIds = new Set(
      snapshot.match(/"population-combine-[1-9]\d*"/g) ?? [],
    );
  let combineNumber = 1;
  for (
    ;
    existingCombineIds.has(
      JSON.stringify(`population-combine-${combineNumber}`),
    );
  )
    combineNumber++;
  ((worksheetRecord.id = `population-combine-${combineNumber}`),
    (worksheetRecord.frozenRows = 1),
    (worksheetRecord.columnWidths = Object.fromEntries(
      headers.map((header, headerIndex) => [
        headerIndex,
        Math.min(240, Math.max(110, header.length * 7 + 20)),
      ]),
    )),
    headers.forEach((raw, headerColumn) => {
      worksheetRecord.cells[formatCellAddress(0, headerColumn)] = {
        raw: raw,
        literal: raw.startsWith("="),
        style: { bold: true, fill: "#E2F0D9" },
      };
    }),
    (worksheetRecord.cells.A1.note = `Reviewed ${parsedRequest.mode} values snapshot; source workbook ${workbook.id}, revision ${workbook.revision}. Left: '${left.sheetName}'!${left.source.range} (${left.records.length} records); right: '${right.sheetName}'!${right.source.range} (${right.records.length} records). Keys: ${parsedRequest.keys ? JSON.stringify(parsedRequest.keys) : "exact matching headers"}. ${matching?.policy ?? "Append by exact headers."} Normalization: ${JSON.stringify(parsedRequest.normalization ?? {})}; blank keys do not match. Matched left/right: ${matchedLeftRows}/${matchedRightRows}; match pairs: ${matchPairs}; unmatched left/right: ${unmatchedLeft.length}/${unmatchedRight.length}; output records: ${rows.length}. Duplicate key groups left/right: ${duplicateLeftKeys}/${duplicateRightKeys}; many-to-many groups: ${manyToManyKeys}; extra many-to-many pairs: ${expandedMatchRows}. Hidden sources explicitly included: ${context.includeHidden}. Filter views do not exclude records. This snapshot does not automatically refresh.`),
    rows.forEach((row, rowIndex) =>
      [...row.values, row.left, row.right].forEach((raw, outputColumn) => {
        worksheetRecord.cells[formatCellAddress(rowIndex + 1, outputColumn)] = {
          raw: raw,
          ...(typeof raw == "string" && raw.startsWith("=")
            ? { literal: true }
            : {}),
          note: `Combined values snapshot. Left source: ${row.left ?? "unmatched / not applicable"}. Right source: ${row.right ?? "unmatched / not applicable"}. Workbook ${workbook.id}; revision ${workbook.revision}. ${row.match?.explanation ?? "No matched pair / append source record."}`,
        };
      }),
    ));
  const next = structuredClone(workbook);
  return (
    next.sheets.push(worksheetRecord),
    normalizeWorkbookInput(next),
    {
      snapshot: snapshot,
      request: parsedRequest,
      context: context,
      left: left,
      right: right,
      headers: headers,
      rows: rows,
      unmatchedLeft: unmatchedLeft,
      unmatchedRight: unmatchedRight,
      keyProfiles: keyProfiles,
      ...(matching ? { matching: matching } : {}),
      counts: {
        leftRows: left.records.length,
        rightRows: right.records.length,
        matchedLeftRows: matchedLeftRows,
        matchedRightRows: matchedRightRows,
        matchPairs: matchPairs,
        duplicateLeftKeys: duplicateLeftKeys,
        duplicateRightKeys: duplicateRightKeys,
        manyToManyKeys: manyToManyKeys,
        expandedMatchRows: expandedMatchRows,
        outputRows: rows.length,
      },
      outputSheetId: worksheetRecord.id,
      outputRange: `A1:${formatCellAddress(rows.length, headers.length - 1)}`,
      next: next,
    }
  );
}
function assertConsolidationPreviewCurrent(workbook, preview, options = {}) {
  (options.readOnly && failWithMessage("This workbook is read-only."),
    preview.snapshot !== JSON.stringify(workbook) &&
      failWithMessage(
        "The workbook changed after preview. Preview and review the populations again.",
      ));
  const consolidationOutput = buildConsolidationOutput(
    workbook,
    preview.request,
    options,
  );
  return (
    JSON.stringify(consolidationOutput) !== JSON.stringify(preview) &&
      failWithMessage(
        "The preview or host context changed. Preview and review again.",
      ),
    consolidationOutput.next
  );
}
function normalizeAuthorizedContext(workbook, options) {
  if (
    (options.readOnly !== undefined && typeof options.readOnly != "boolean") ||
    (options.includeHidden !== undefined &&
      typeof options.includeHidden != "boolean")
  )
    throw Error("Host read-only and hidden-source policies must be boolean.");
  const allowed =
    options.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id);
  if (
    !Array.isArray(allowed) ||
    allowed.some(
      (sheetId) =>
        typeof sheetId != "string" ||
        !workbook.sheets.some((sheet) => sheet.id === sheetId),
    )
  )
    throw Error("Invalid authorized worksheet context.");
  const external = options.approvedWorkbooks ?? [];
  if (
    !Array.isArray(external) ||
    external.length > 8 ||
    new Set(external.map((approvedWorkbook) => approvedWorkbook.workbook.id))
      .size !== external.length ||
    external.some(
      (approvedEntry) =>
        approvedEntry.workbook.id === workbook.id ||
        !Array.isArray(approvedEntry.allowedSheetIds) ||
        approvedEntry.allowedSheetIds.some(
          (allowedSheetId) =>
            !approvedEntry.workbook.sheets.some(
              (sheet) => sheet.id === allowedSheetId,
            ),
        ) ||
        (approvedEntry.includeHidden !== undefined &&
          typeof approvedEntry.includeHidden != "boolean"),
    )
  )
    throw Error(
      "Approved workbook identities and worksheet permissions must be explicit, distinct and valid (at most 8 workbooks).",
    );
  return { allowed: allowed, external: external };
}
function readSourceTableRecords(workbook, source, options) {
  const authorizedContext = normalizeAuthorizedContext(workbook, options),
    approvedWorkbook = authorizedContext.external.find(
      (approvedCandidate) =>
        approvedCandidate.workbook.id === source.workbookId,
    ),
    isSameWorkbook = source.workbookId === workbook.id,
    book = isSameWorkbook ? workbook : approvedWorkbook?.workbook;
  if (!book)
    throw Error(
      "Source workbook is missing or has not been approved. Attach and approve the exact workbook identity before refreshing.",
    );
  const allowedSheetIds = isSameWorkbook
      ? authorizedContext.allowed
      : approvedWorkbook.allowedSheetIds,
    sheet = book.sheets.find((sheet) => sheet.id === source.sheetId);
  if (!sheet || !allowedSheetIds.includes(source.sheetId))
    throw Error(
      "Source worksheet is missing or outside the approved worksheet scope.",
    );
  let range = source.range;
  if (source.tableId) {
    const sourceTable = sheet.tables?.find(
      (table) => table.id === source.tableId,
    );
    if (!sourceTable)
      throw Error("Source table is missing. Review the source mapping.");
    const rangeAddress = parseRangeAddress(sourceTable.range);
    range = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1 - (sourceTable.showTotals ? 1 : 0)}`;
  }
  const sourceHeaders = readSourceHeaders(
      book,
      [{ sheetId: sheet.id, range: range }],
      {
        allowedSheetIds: allowedSheetIds,
        includeHidden: isSameWorkbook
          ? options.includeHidden
          : !!options.includeHidden && !!approvedWorkbook.includeHidden,
      },
    ),
    snapshot = sourceHeaders.snapshots[0];
  return {
    book: book,
    sheet: sheet,
    snapshot: snapshot,
    fingerprint: fingerprintCanonicalValue({
      workbookId: book.id,
      workbookName: book.name,
      source: source,
      snapshot: snapshot,
      dependencies: sourceHeaders.dependencyCells,
      datePolicy: normalizeDatePolicy(book.datePolicy),
    }),
  };
}
function reconcileConsolidationSources(workbook, definition, options = {}) {
  const consolidationDefinition = parseConsolidationDefinition(definition);
  normalizeAuthorizedContext(workbook, options);
  const review = {
      sources: [],
      contributions: [],
      subtotals: [],
      exceptions: [],
      grandTotal: null,
      sourceTotal: null,
      difference: null,
      complete: false,
    },
    seenAddresses = new Set(),
    addException = (sourceId, code, message, row, severity = "error") =>
      review.exceptions.push({
        sourceId: sourceId,
        code: code,
        message: message,
        ...(row ? { row: row } : {}),
        severity: severity,
      });
  for (const source of consolidationDefinition.sources)
    try {
      const sourceTableRecords = readSourceTableRecords(
          workbook,
          source,
          options,
        ),
        { snapshot } = sourceTableRecords,
        rangeAddress = parseRangeAddress(snapshot.source.range),
        requireHeaderIndex = (headerName) => {
          const headerIndex = snapshot.headers.indexOf(headerName);
          if (headerIndex < 0)
            throw Error(
              `Required field “${headerName}” is missing. Review the source schema.`,
            );
          return headerIndex;
        },
        categoryIndexes = source.categories.map(requireHeaderIndex),
        amountIndex = requireHeaderIndex(source.amountField),
        readDimensionValue = (dimensionKey, recordValues) =>
          source[dimensionKey].kind === "constant"
            ? source[dimensionKey].value
            : recordValues[requireHeaderIndex(source[dimensionKey].field)];
      for (const dimensionName of ["period", "currency", "unit"])
        source[dimensionName].kind === "field" &&
          requireHeaderIndex(source[dimensionName].field);
      const sourceSummary = {
        sourceId: source.id,
        workbookId: sourceTableRecords.book.id,
        workbookName: sourceTableRecords.book.name,
        sheetId: sourceTableRecords.sheet.id,
        sheetName: sourceTableRecords.sheet.name,
        range: snapshot.source.range,
        fingerprint: sourceTableRecords.fingerprint,
        records: snapshot.records.length,
        acceptedRecords: 0,
        amount: 0,
        periods: [],
        currencies: [],
        units: [],
      };
      review.sources.push(sourceSummary);
      const firstRowByCategories = new Map(),
        sourceContributions = [];
      for (const record of snapshot.records) {
        let isRejected = false;
        const addRecordException = (exceptionCode, exceptionMessage) => {
            (addException(
              source.id,
              exceptionCode,
              exceptionMessage,
              record.row,
            ),
              (isRejected = true));
          },
          categories = categoryIndexes.map(
            (categoryIndex) => record.values[categoryIndex],
          ),
          amount = record.values[amountIndex],
          address = `${columnIndexToLetters(rangeAddress.firstCol + amountIndex)}${record.row}`,
          overlapKey = JSON.stringify([
            sourceTableRecords.book.id,
            sourceTableRecords.sheet.id,
            address,
          ]);
        (seenAddresses.has(overlapKey) &&
          addRecordException(
            "overlap",
            `Amount ${sourceTableRecords.sheet.name}!${address} is included more than once. Remove overlapping sources to prevent double counting.`,
          ),
          seenAddresses.add(overlapKey),
          categories.some(
            (category) =>
              category === null ||
              (typeof category == "string" && !category.trim()),
          ) &&
            addRecordException(
              "blank-category",
              "Every contributing row needs nonblank categories.",
            ),
          (typeof amount != "number" || !Number.isFinite(amount)) &&
            addRecordException(
              "amount",
              `Amount ${sourceTableRecords.sheet.name}!${address} must be a finite number; text is not silently converted.`,
            ));
        const recordDimensions = {
          period: readDimensionValue("period", record.values),
          currency: readDimensionValue("currency", record.values),
          unit: readDimensionValue("unit", record.values),
        };
        for (const dimension of ["period", "currency", "unit"]) {
          const dimensionValue = recordDimensions[dimension];
          if (
            typeof dimensionValue == "string" &&
            dimensionValue.trim() &&
            dimensionValue.length <= 128 &&
            !/[\u0000-\u001f\u007f]/.test(dimensionValue)
          ) {
            const distinctValues =
              sourceSummary[
                dimension === "currency"
                  ? "currencies"
                  : dimension === "period"
                    ? "periods"
                    : "units"
              ];
            distinctValues.includes(dimensionValue) ||
              distinctValues.push(dimensionValue);
          }
          dimensionValue !== consolidationDefinition[dimension] &&
            addRecordException(
              dimension,
              `Source ${dimension} ${JSON.stringify(dimensionValue)} does not match ${JSON.stringify(consolidationDefinition[dimension])}. No currency, unit or period conversion is implicit.`,
            );
        }
        const categoryKey = serializeTypedKeyValues(categories),
          firstRowForKey = firstRowByCategories.get(categoryKey);
        (firstRowForKey !== undefined
          ? (addException(
              source.id,
              "duplicate-category",
              `Repeated category also occurs at row ${firstRowForKey}. ${source.duplicatePolicy === "sum" ? "Explicit sum policy retains both contributions." : "Choose Sum repeated categories only after reviewing these records."}`,
              record.row,
              source.duplicatePolicy === "sum" ? "warning" : "error",
            ),
            source.duplicatePolicy === "reject" && (isRejected = true))
          : firstRowByCategories.set(categoryKey, record.row),
          isRejected ||
            sourceContributions.push({
              sourceId: source.id,
              workbookId: sourceTableRecords.book.id,
              workbookName: sourceTableRecords.book.name,
              sheetId: sourceTableRecords.sheet.id,
              sheetName: sourceTableRecords.sheet.name,
              row: record.row,
              address: address,
              locator: `${sourceTableRecords.book.name} / ${snapshot.sheetName}!${address}`,
              categories: categories,
              amount: amount,
              period: recordDimensions.period,
              currency: recordDimensions.currency,
              unit: recordDimensions.unit,
              link:
                sourceTableRecords.book.id !== workbook.id
                  ? "approved snapshot"
                  : consolidationDefinition.linkedAmounts
                    ? "live"
                    : "snapshot",
            }));
      }
      ((sourceSummary.acceptedRecords = sourceContributions.length),
        (sourceSummary.amount = sumConsolidationAmounts(
          sourceContributions.map((entry) => entry.amount),
        )),
        review.contributions.push(...sourceContributions));
    } catch (error) {
      addException(source.id, "source", error.message.slice(0, 4e3));
    }
  if (review.contributions.length > consolidationLimits.contributions)
    throw Error(
      "Consolidation exceeds 2,000 contributions. Narrow the sources; no output was truncated.",
    );
  if (
    (review.contributions.length ||
      addException(
        consolidationDefinition.sources[0].id,
        "empty",
        "No valid contributing amounts. Review the sources and exceptions.",
      ),
    !review.exceptions.some((exception) => exception.severity === "error"))
  ) {
    ((review.subtotals = buildConsolidationSubtotals(review.contributions)),
      (review.grandTotal = sumConsolidationAmounts(
        review.contributions.map((contribution) => contribution.amount),
      )),
      (review.sourceTotal = sumConsolidationAmounts(
        review.sources.map((source) => source.amount),
      )),
      (review.difference = review.grandTotal - review.sourceTotal));
    const amountTolerance =
      Number.EPSILON *
      Math.max(1, Math.abs(review.grandTotal), Math.abs(review.sourceTotal)) *
      review.contributions.length *
      4;
    (Math.abs(review.difference) > amountTolerance &&
      addException(
        consolidationDefinition.sources[0].id,
        "reconciliation",
        "Source control totals do not reconcile to the contributing amounts.",
      ),
      (review.complete = !review.exceptions.some(
        (exception) => exception.severity === "error",
      )));
  }
  return review;
}
function buildAuthorizedContextSnapshot(workbook, options) {
  const authorizedContext = normalizeAuthorizedContext(workbook, options);
  return fingerprintCanonicalValue({
    allowed: [...authorizedContext.allowed].sort(),
    hidden: !!options.includeHidden,
    readOnly: !!options.readOnly,
    approved: authorizedContext.external
      .map((externalEntry) => ({
        id: externalEntry.workbook.id,
        sheets: [...externalEntry.allowedSheetIds].sort(),
        hidden: !!externalEntry.includeHidden,
      }))
      .sort((leftEntry, rightEntry) =>
        leftEntry.id.localeCompare(rightEntry.id),
      ),
  });
}
function validateOutputSheetName(workbook, outputName) {
  if (
    typeof outputName != "string" ||
    !outputName.trim() ||
    outputName !== outputName.trim() ||
    outputName.length > 31 ||
    /[\\/?*\[\]:\u0000-\u001f\u007f]/.test(outputName) ||
    outputName.startsWith("'") ||
    outputName.endsWith("'") ||
    workbook.sheets.some(
      (sheet) => sheet.name.toLowerCase() === outputName.toLowerCase(),
    )
  )
    throw Error(
      "Choose a distinct worksheet name of 1–31 characters without invalid characters.",
    );
  inspectOutputNameSafety(workbook, outputName);
}
function createConsolidationPreview(workbook, request, options = {}) {
  validateOutputSheetName(workbook, request.outputName);
  const consolidationDefinition = parseConsolidationDefinition(
      request.definition,
    ),
    review = reconcileConsolidationSources(
      workbook,
      consolidationDefinition,
      options,
    );
  return {
    request: {
      definition: consolidationDefinition,
      outputName: request.outputName,
    },
    workbookHash: fingerprintCanonicalValue(workbook),
    contextHash: buildAuthorizedContextSnapshot(workbook, options),
    review: review,
  };
}
function buildConsolidationOutputCells(worksheetRecord, outputRange) {
  const rangeAddress = parseRangeAddress(outputRange);
  if (
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    consolidationLimits.outputCells
  )
    throw Error("Consolidation output range exceeds its cell budget.");
  const outputRows = [];
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  ) {
    const rowCells = [];
    for (
      let firstCol = rangeAddress.firstCol;
      firstCol <= rangeAddress.lastCol;
      firstCol++
    ) {
      const cell =
        worksheetRecord.cells[
          `${columnIndexToLetters(firstCol)}${firstRow + 1}`
        ];
      rowCells.push({
        raw: cell?.raw ?? null,
        literal: !!cell?.literal,
        note: cell?.note ?? "",
      });
    }
    outputRows.push(rowCells);
  }
  return fingerprintCanonicalValue(outputRows);
}

export {
  readSourceRectangles,
  buildConsolidationOutput,
  readSourceHeaders,
  matchKeyPopulations,
  assertConsolidationPreviewCurrent,
  readSourceTableRecords,
  inspectFormulaDependencies,
  buildConsolidationOutputCells,
  inspectOutputNameSafety,
  createConsolidationPreview,
  buildAuthorizedContextSnapshot,
};
export { hasOnlyAllowedKeys };
