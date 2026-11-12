import { parseDelimitedCellValue } from "./06c-package-and-worksheet-parsing.js";

import {
  F$,
  Hle,
  Wle,
  buildContainsTextFormula,
  buildDelimitedImportReport,
  buildExchangeSchemaReport,
  byteCharacterTable,
  computeCrc32,
  convertTemplateToXlsx,
  createEmptyExportLedger,
  createWorkbookParseState,
  dataValidationOperators,
  decodeXmlEntities,
  delimitedImportDefaults,
  formatCellAddress,
  formatRangeCornerCells,
  getXmlAttribute,
  inspectXlsxPackage,
  isExternalActionFormula,
  isExternalReferenceFormula,
  maxXlsxExpandedByteLimit,
  normalizeDelimitedOptions,
  numericPrecisionPolicy,
  parseCalculationProperties,
  parseWorksheetCells,
  parseWorksheetXml,
  parseXlsxDataValidations,
  parseXmlPartStrict,
  readPackageRelationships,
  recordDelimitedChange,
  recordDelimitedIssue,
  relationshipsPathFor,
  resolvePackagePartPath,
  scanXlsxNumericTokens,
  summarizePreservedParts,
  validateConditionalFormatRules,
  validateWorkbookPackage,
  validateXlsxNumericValue,
  xlsxExchangeSupportNotice,
} from "./06c-package-and-worksheet-parsing.js";

import {
  strFromU8,
  strToU8,
  zipSync,
  Unzip,
  UnzipInflate,
  unzipSync,
} from "fflate";
import iue from "exceljs";
import lue from "papaparse";
import {
  unwrapDefaultExport,
  sha256HexDigest,
  normalizeDatePolicy,
  dateSystemMaxSerial,
} from "./01-schemas-date-functions-and-validation.js";
import {
  fileImportLimits,
  workbookLimits,
} from "./02-text-and-financial-functions.js";
import {
  assertWorkbookStructureEditable,
  defaultRowHeightOf,
  readHiddenAxisIntervals,
  defaultCalculationSettings,
  validatePrintPageBreaks,
  sortConditionalRulesByPriority,
  normalizeCalculationSettings,
  defaultColumnWidthOf,
  buildNumberFormatCode,
  findOccupiedRangeBounds,
  calculateWorkbookWithSettings,
  requireRetainedXlsxSource,
  encodeBytesToBase64,
  _t,
  aP,
  updateSheetPageSetup,
  cn,
  parseRangeAddress,
  serializeDefinedNames,
  buildDrawingExportNotes,
  applyFilterConditions,
  resolveBorderEdge,
  normalizeWorkbookInput,
  cellBorderEdges,
  columnIndexToLetters,
  buildTableCellEntries,
  worksheetCapacityLimits,
  resolveEffectiveCellStyle,
} from "./03-workbook-model.js";

import {
  inspectionSignatureCache,
  describeImportSource,
  readGeneratedTemplateBuffer,
} from "./07-file-import-and-worker-dispatch.js";

async function exportXltxTemplate(workbook) {
  const exportResult = await exportWorkbookToXlsx(workbook),
    templateBytes = convertTemplateToXlsx(
      new Uint8Array(await readGeneratedTemplateBuffer(exportResult.blob)),
      "xlsx",
      "xltx",
    ),
    warnings = [
      ...exportResult.warnings,
      "Derived XLTX template: supported worksheet content only. Macro projects, native automation and unsupported special parts are not transferred. Reopening creates a new workbook; keep native JSON and original source files.",
    ];
  return {
    blob: new Blob([new Uint8Array(templateBytes)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
    }),
    warnings: warnings,
    report: {
      ...exportResult.report,
      format: "xltx",
      fileName: describeImportSource(workbook, "xltx"),
      fileBytes: templateBytes.length,
      warnings: warnings,
    },
  };
}
const g7 = "http://schemas.openxmlformats.org/package/2006/relationships",
  contentTypesNamespace =
    "http://schemas.openxmlformats.org/package/2006/content-types",
  escapeXmlAttribute = (value) =>
    value.replace(
      /[&<>"']/g,
      (matchedCharacter) =>
        ({
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
          "'": "&apos;",
        })[matchedCharacter],
    );

function rebuildPreservedPackage(
  originalParts,
  generatedParts,
  preservationInventory,
) {
  const inventory = structuredClone(preservationInventory),
    packageRelationships = readPackageRelationships(originalParts),
    generatedRelationships = readPackageRelationships(generatedParts),
    packageFiles = { ...generatedParts };
  inventory.phase = "export-result";
  for (const family of ["custom XML", "custom properties"]) {
    const preservedOpaqueParts = inventory.parts
      .filter(
        (part) =>
          part.family === family ||
          (family === "custom XML" && part.path.startsWith("customXml/")),
      )
      .filter(
        (candidatePart) => candidatePart.disposition === "preserved-opaque",
      );
    if (!preservedOpaqueParts.length) continue;
    const opaquePaths = new Set(
        preservedOpaqueParts.map((opaquePart) => opaquePart.path),
      ),
      inboundRelationships = packageRelationships.relationships.filter(
        (relationship) =>
          !opaquePaths.has(relationship.owner) &&
          opaquePaths.has(relationship.target),
      );
    if (
      preservedOpaqueParts.some(
        (collidingPart) =>
          packageFiles[collidingPart.path] &&
          sha256HexDigest(packageFiles[collidingPart.path]) !==
            collidingPart.sha256,
      ) ||
      inboundRelationships.some((inboundRelationship) =>
        generatedRelationships.relationships.some(
          (relationship) =>
            relationship.owner === inboundRelationship.owner &&
            relationship.type === inboundRelationship.type,
        ),
      )
    ) {
      for (const droppedPart of preservedOpaqueParts)
        ((droppedPart.disposition = "dropped"),
          (droppedPart.reason =
            "Generated package has a conflicting part or relationship. The entire metadata island remains only in the original."));
      continue;
    }
    for (const copiedPart of preservedOpaqueParts)
      packageFiles[copiedPart.path] = originalParts[copiedPart.path];
    for (const owner of new Set(
      inboundRelationships.map((relationshipEntry) => relationshipEntry.owner),
    )) {
      const relsPath = relationshipsPathFor(owner),
        relsXml = packageFiles[relsPath]
          ? strFromU8(packageFiles[relsPath])
          : `<Relationships xmlns="${g7}"/>`,
        existingIds = new Set();
      packageFiles[relsPath] &&
        parseXmlPartStrict(packageFiles[relsPath], (xmlElement) => {
          xmlElement.local === "Relationship" &&
            existingIds.add(xmlElement.attributes.Id?.value ?? "");
        });
      let preserveIndex = 1;
      const addedRelationships = inboundRelationships
        .filter((candidateEntry) => candidateEntry.owner === owner)
        .map((ownedEntry) => {
          for (; existingIds.has(`lgPreserve${preserveIndex}`);)
            preserveIndex++;
          const preserveId = `lgPreserve${preserveIndex++}`;
          return (
            existingIds.add(preserveId),
            `<Relationship xmlns="${g7}" Id="${preserveId}" Type="${escapeXmlAttribute(ownedEntry.type)}" Target="/${escapeXmlAttribute(ownedEntry.target)}"/>`
          );
        })
        .join("");
      packageFiles[relsPath] = strToU8(
        expandEmptyXmlElements(relsXml, "Relationships", addedRelationships),
      );
    }
    const contentTypeOverrides = preservedOpaqueParts
      .filter(
        (unregisteredPart) =>
          !unregisteredPart.path.endsWith(".rels") &&
          !generatedRelationships.types.has(unregisteredPart.path),
      )
      .map(
        (overridePart) =>
          `<Override xmlns="${contentTypesNamespace}" PartName="/${escapeXmlAttribute(overridePart.path)}" ContentType="${escapeXmlAttribute(overridePart.contentType)}"/>`,
      )
      .join("");
    packageFiles["[Content_Types].xml"] = strToU8(
      expandEmptyXmlElements(
        strFromU8(packageFiles["[Content_Types].xml"]),
        "Types",
        contentTypeOverrides,
      ),
    );
  }
  return {
    bytes: zipSync(packageFiles),
    inventory: summarizePreservedParts(inventory),
  };
}
function expandEmptyXmlElements(xmlText, elementName, insertedXml) {
  if (!insertedXml) return xmlText;
  const closingTagPattern = new RegExp(
    `</((?:[\\w.-]+:)?${elementName})\\s*>\\s*$`,
  );
  if (closingTagPattern.test(xmlText))
    return xmlText.replace(closingTagPattern, `${insertedXml}$&`);
  const selfClosingPattern = new RegExp(
    `<((?:[\\w.-]+:)?${elementName})([^>]*)/>(\\s*)$`,
  );
  if (selfClosingPattern.test(xmlText))
    return xmlText.replace(selfClosingPattern, `<$1$2>${insertedXml}</$1>$3`);
  throw Error("Unable to preserve XLSX package metadata safely.");
}
function describePreservationDisposition(inventory, dropReason) {
  const droppedInventory = structuredClone(inventory);
  droppedInventory.phase = "export-result";
  for (const part of droppedInventory.parts)
    part.disposition === "preserved-opaque" &&
      ((part.disposition = "dropped"), (part.reason = dropReason));
  return summarizePreservedParts(droppedInventory);
}

const countRangeCells = (rangeText) => {
    const rangeAddress = parseRangeAddress(rangeText);
    return (
      (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1)
    );
  },
  hexColorToArgb = (hexColor) => `FF${hexColor.slice(1).toUpperCase()}`;

function applyValidationDropdownFlags(sheetXml, sheet) {
  return sheetXml.replace(/<dataValidation\b[^>]*>/g, (validationTag) => {
    const sqref = /\bsqref="([^"]+)"/.exec(validationTag)?.[1],
      matchingValidation = sheet.validations?.find(
        (validation) =>
          sqref &&
          formatRangeCornerCells(validation.range) ===
            formatRangeCornerCells(sqref),
      );
    return matchingValidation?.kind === "list" &&
      matchingValidation.showDropdown === false
      ? validationTag
          .replace(/\s+showDropDown="[^"]*"/, "")
          .replace(/>$/, ' showDropDown="1">')
      : validationTag;
  });
}

function serializeDataValidationRule(validationRule) {
  const excelType = {
    list: "list",
    number: "decimal",
    integer: "whole",
    textLength: "textLength",
    date: "date",
    time: "time",
    custom: "custom",
  }[validationRule.kind];
  if (
    validationRule.kind === "list" &&
    (validationRule.list.some((listItem) => /[",\r\n]/.test(listItem)) ||
      validationRule.list.join(",").length > 253)
  )
    throw Error(
      "This list validation cannot be represented by an inline XLSX list. Retain the native workbook or use shorter values without commas, quotes or line breaks.",
    );
  if (
    validationRule.kind === "custom" &&
    isExternalReferenceFormula(validationRule.formula)
  )
    throw Error(
      "An unsafe custom validation formula cannot be exported. Keep the native workbook and review the rule.",
    );
  return {
    type: excelType,
    ...(["list", "custom"].includes(validationRule.kind)
      ? {}
      : { operator: dataValidationOperators[validationRule.operator] }),
    allowBlank: validationRule.allowBlank,
    showErrorMessage: true,
    errorStyle: validationRule.errorStyle ?? "stop",
    errorTitle: validationRule.errorTitle ?? "Invalid value",
    error: validationRule.message,
    showInputMessage: validationRule.showInputMessage !== false,
    promptTitle: validationRule.inputTitle ?? "",
    prompt: validationRule.inputMessage ?? "",
    formulae:
      validationRule.kind === "list"
        ? [JSON.stringify(validationRule.list.join(","))]
        : validationRule.kind === "custom"
          ? [validationRule.formula.slice(1)]
          : [
              validationRule.min,
              ...(["between", "notBetween"].includes(validationRule.operator)
                ? [validationRule.max]
                : []),
            ].map((formulaValue) =>
              validationRule.kind === "date"
                ? new Date(Date.UTC(1899, 11, 30) + formulaValue * 864e5)
                : formulaValue,
            ),
  };
}

function appendStopIfTrueAttribute(xmlText, sheet) {
  const sortedRules = sortConditionalRulesByPriority(
    sheet.conditionalFormats ?? [],
  );
  return xmlText.replace(/<cfRule\b[^>]*>/g, (ruleTag) => {
    const priority = Number(/\bpriority="(\d+)"/.exec(ruleTag)?.[1]),
      matchedRule = sortedRules[priority - 1];
    return matchedRule?.stopIfTrue
      ? ruleTag
          .replace(/\s+stopIfTrue="[^"]*"/, "")
          .replace(/>$/, ' stopIfTrue="1">')
      : ruleTag;
  });
}
function serializeConditionalFormatRules(conditionalFormat, priority) {
  const dxfStyle = {
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: hexColorToArgb(conditionalFormat.fill) },
      },
      font: { color: { argb: hexColorToArgb(conditionalFormat.color) } },
    },
    rangeAddress = parseRangeAddress(conditionalFormat.range),
    cellAddress = formatCellAddress(
      rangeAddress.firstRow,
      rangeAddress.firstCol,
    ),
    buildExpressionRule = (formulaText) => ({
      type: "expression",
      priority: priority,
      formulae: [formulaText.replace(/^=/, "")],
      style: dxfStyle,
    });
  if (conditionalFormat.kind === "formula") {
    if (isExternalReferenceFormula(conditionalFormat.formula ?? ""))
      throw Error(
        "External-action conditional formulas cannot be exported. Remove that rule or use native JSON.",
      );
    return buildExpressionRule(conditionalFormat.formula);
  }
  if (
    ["blank", "notBlank", "error", "notError"].includes(conditionalFormat.kind)
  )
    return buildExpressionRule(
      conditionalFormat.kind === "blank"
        ? `LEN(${cellAddress})=0`
        : conditionalFormat.kind === "notBlank"
          ? `LEN(${cellAddress})>0`
          : conditionalFormat.kind === "error"
            ? `ISERROR(${cellAddress})`
            : `NOT(ISERROR(${cellAddress}))`,
    );
  if (
    conditionalFormat.kind === "unique" ||
    conditionalFormat.kind === "duplicate"
  )
    return buildExpressionRule(
      Wle(conditionalFormat.kind, conditionalFormat.range),
    );
  if (conditionalFormat.kind === "dataBar")
    return {
      type: "dataBar",
      priority: priority,
      cfvo: [{ type: "min" }, { type: "max" }],
      color: { argb: hexColorToArgb(conditionalFormat.fill) },
      gradient: false,
      showValue: true,
      negativeBarColorSameAsPositive: false,
      negativeFillColor: { argb: "FFD13438" },
      axisColor: { argb: "FF777777" },
      axisPosition: "auto",
      direction: "leftToRight",
      minLength: 0,
      maxLength: 100,
    };
  if (conditionalFormat.kind === "iconSet")
    return {
      type: "iconSet",
      priority: priority,
      iconSet: "3Arrows",
      showValue: true,
      reverse: false,
      cfvo: [
        { type: "min" },
        ...(conditionalFormat.iconThresholds ?? [33, 67]).map((threshold) => ({
          type:
            conditionalFormat.thresholdType === "number" ? "num" : "percent",
          value: threshold,
        })),
      ],
    };
  if (conditionalFormat.kind === "contains")
    return buildExpressionRule(
      buildContainsTextFormula(conditionalFormat.value, cellAddress),
    );
  if (conditionalFormat.kind === "colorScale")
    return {
      type: "colorScale",
      priority: priority,
      cfvo: [{ type: "min" }, { type: "max" }],
      color: [
        { argb: hexColorToArgb(conditionalFormat.minColor ?? "#f8e99b") },
        { argb: hexColorToArgb(conditionalFormat.maxColor ?? "#63be7b") },
      ],
    };
  const operator =
      conditionalFormat.kind === "greater"
        ? "greaterThan"
        : conditionalFormat.kind === "less"
          ? "lessThan"
          : conditionalFormat.kind === "equal"
            ? "equal"
            : "between",
    toFormulaLiteral = (rawValue) =>
      rawValue.trim() !== "" && Number.isFinite(Number(rawValue))
        ? Number(rawValue)
        : `"${rawValue.replace(/"/g, '""')}"`;
  return {
    type: "cellIs",
    operator: operator,
    formulae: [
      toFormulaLiteral(conditionalFormat.value),
      ...(conditionalFormat.kind === "between"
        ? [toFormulaLiteral(conditionalFormat.second)]
        : []),
    ],
    priority: priority,
    style: dxfStyle,
  };
}
function serializeWorksheetRuleXml(worksheet, sheet) {
  let mergedCellTotal = 0,
    validationCellTotal = 0;
  for (const mergeRange of sheet.merges ?? []) {
    if (
      (mergedCellTotal += countRangeCells(mergeRange)) > workbookLimits.maxCells
    )
      throw Error("Merged ranges exceed the 100,000-cell interchange budget.");
    const rangeAddress = parseRangeAddress(mergeRange);
    for (const [cellAddress, cell] of Object.entries(sheet.cells)) {
      const cellPosition = _t(cellAddress);
      if (
        cellPosition.row >= rangeAddress.firstRow &&
        cellPosition.row <= rangeAddress.lastRow &&
        cellPosition.col >= rangeAddress.firstCol &&
        cellPosition.col <= rangeAddress.lastCol &&
        (cellPosition.row !== rangeAddress.firstRow ||
          cellPosition.col !== rangeAddress.firstCol) &&
        ((cell.raw !== null && cell.raw !== "") ||
          cell.note ||
          cell.tickmark ||
          cell.hyperlink)
      )
        throw Error(
          "Exporting this merged area would discard populated cells or annotations. Unmerge it or preserve those values before XLSX export.",
        );
    }
    worksheet.mergeCells(mergeRange);
  }
  const validationRanges = [];
  for (const validation of sheet.validations ?? []) {
    const rangeAddress = parseRangeAddress(validation.range);
    if (
      validationRanges.some(
        (pushedRange) =>
          pushedRange.firstRow <= rangeAddress.lastRow &&
          pushedRange.lastRow >= rangeAddress.firstRow &&
          pushedRange.firstCol <= rangeAddress.lastCol &&
          pushedRange.lastCol >= rangeAddress.firstCol,
      )
    )
      throw Error(
        "Overlapping validation rules cannot be represented by one Excel rule per cell. Use native JSON or resolve overlapping rules before export.",
      );
    if (
      (validationRanges.push(rangeAddress),
      (validationCellTotal += countRangeCells(validation.range)) >
        workbookLimits.maxCells)
    )
      throw Error(
        "XLSX validation ranges exceed the 100,000-cell interchange budget. Use native JSON to preserve larger rules.",
      );
    worksheet.dataValidations.add(
      validation.range,
      serializeDataValidationRule(validation),
    );
  }
  const priorityById = new Map(
    sortConditionalRulesByPriority(sheet.conditionalFormats ?? []).map(
      (sortedRule, sortedIndex) => [sortedRule.id, sortedIndex + 1],
    ),
  );
  let populationChecks = 0;
  for (const conditionalRule of sheet.conditionalFormats ?? []) {
    if (
      conditionalRule.kind !== "duplicate" &&
      conditionalRule.kind !== "unique"
    )
      continue;
    const rangeAddress = parseRangeAddress(conditionalRule.range);
    for (const [candidateAddress, cell] of Object.entries(sheet.cells)) {
      if (++populationChecks > 1e6)
        throw Error(
          "Duplicate/unique conditional export exceeds 1,000,000 population checks. Use native JSON or fewer rules.",
        );
      const candidatePosition = _t(candidateAddress);
      if (
        candidatePosition.row < rangeAddress.firstRow ||
        candidatePosition.row > rangeAddress.lastRow ||
        candidatePosition.col < rangeAddress.firstCol ||
        candidatePosition.col > rangeAddress.lastCol
      )
        continue;
      let raw = cell.raw;
      if (
        (typeof cell.raw == "string" &&
          cell.raw.startsWith("=") &&
          !cell.literal &&
          (raw = worksheet.getCell(candidateAddress).result),
        !Hle(raw))
      )
        throw Error(
          `Duplicate/unique conditional rule at ${candidateAddress} cannot preserve Excel comparison semantics for this value. Use native JSON; XLSX supports finite numbers, canonical numeric text and short printable ASCII text without date-like numeric prefixes.`,
        );
    }
  }
  (sheet.conditionalFormats ?? []).forEach((formatRule) =>
    worksheet.addConditionalFormatting({
      ref: formatRule.range,
      rules: [
        serializeConditionalFormatRules(
          formatRule,
          priorityById.get(formatRule.id),
        ),
      ],
    }),
  );
}

function serializeCalculationProperties(workbookXml, workbookInput) {
  const calculationSettings = normalizeCalculationSettings(
      workbookInput.calculation ?? defaultCalculationSettings,
    ),
    calcPrXml = `<calcPr fullPrecision="1" calcId="0" calcMode="${calculationSettings.mode === "manual" ? "manual" : "auto"}" fullCalcOnLoad="1" forceFullCalc="1" calcOnSave="${calculationSettings.mode === "manual" ? "0" : "1"}" iterate="${calculationSettings.iterative ? "1" : "0"}" iterateCount="${calculationSettings.maxIterations}" iterateDelta="${calculationSettings.maxChange}"/>`;
  return /<(?:\w+:)?calcPr\b[^>]*\/>/.test(workbookXml)
    ? workbookXml.replace(/<(?:\w+:)?calcPr\b[^>]*\/>/, calcPrXml)
    : /<(?:\w+:)?calcPr\b[^>]*>[\s\S]*?<\/(?:\w+:)?calcPr>/.test(workbookXml)
      ? workbookXml.replace(
          /<(?:\w+:)?calcPr\b[^>]*>[\s\S]*?<\/(?:\w+:)?calcPr>/,
          calcPrXml,
        )
      : workbookXml.replace(/<\/(?:\w+:)?workbook>/, `${calcPrXml}$&`);
}
function verifyAutomationInspection(inspection) {
  const cachedInspection = inspectionSignatureCache.get(inspection);
  if (
    !cachedInspection ||
    cachedInspection.signature !== JSON.stringify(inspection)
  )
    throw Error(
      "Automation inspection was modified or is untrusted. Inspect the original bytes again.",
    );
  return cachedInspection.bytes.slice();
}
function retainSourceArtifact(inspection, actor) {
  const inspectionBytes = verifyAutomationInspection(inspection);
  if (!inspection.retainedEligible)
    throw Error(
      "This source exceeds the 2 MB artifact-retention limit. Keep the original file separately.",
    );
  if (!actor.trim() || actor.length > 100)
    throw Error("A bounded reviewer name is required.");
  return {
    id: crypto.randomUUID(),
    fileName: inspection.fileName,
    kind: inspection.kind,
    sha256: inspection.sha256,
    byteLength: inspectionBytes.length,
    base64: encodeBytesToBase64(inspectionBytes),
    retainedAt: new Date().toISOString(),
    retainedBy: actor,
  };
}
function inspectEmbeddedObjects(inspection) {
  verifyAutomationInspection(inspection);
  const embeddedItems = inspection.items.filter((item) =>
    /embedded|activeX|OLE|External automation/i.test(item.kind),
  );
  if (!embeddedItems.length)
    throw Error(
      "No known OLE, ActiveX or embedded-object references were detected. Separately stored objects remain unknown.",
    );
  return {
    id: crypto.randomUUID(),
    fileName: inspection.fileName,
    fileBytes: inspection.fileBytes,
    sourceSha256: inspection.sha256,
    inspectedAt: new Date().toISOString(),
    execution: "never",
    items: embeddedItems.map((embeddedItem) => ({
      ...embeddedItem,
      execution: "never",
    })),
    warnings: [
      ...inspection.warnings,
      "Inventory only: no payload is opened, fetched or executed. Missing parts and external references remain unresolved. Keep the original source separately.",
    ],
  };
}
function retainEmbeddedObjectInventory(workbook, inspection, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  assertWorkbookStructureEditable(workbook);
  const inventory = inspectEmbeddedObjects(inspection);
  if (
    workbook.embeddedObjectInventories?.some(
      (embeddedObjectInventory) =>
        embeddedObjectInventory.sourceSha256 === inventory.sourceSha256,
    )
  )
    throw Error("This source inventory has already been retained.");
  return normalizeWorkbookInput({
    ...workbook,
    embeddedObjectInventories: [
      ...(workbook.embeddedObjectInventories ?? []),
      inventory,
    ],
  });
}

const excelErrorLiterals = new Set([
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
]);
function toExportedCellValue(cell) {
  if (cell.type !== "blank")
    return cell.type === "error"
      ? excelErrorLiterals.has(String(cell.value))
        ? { error: cell.value }
        : undefined
      : cell.value;
}
function classifyArrayFormula(xmlText, sheetName) {
  const readAttribute = (tagText, attributeName) => {
    const attributeValue = new RegExp(
      `(?:^|\\s)${attributeName}\\s*=\\s*(["'])(.*?)\\1`,
      "i",
    ).exec(tagText)?.[2];
    return attributeValue?.replace(
      /&#(x[0-9a-f]+|[0-9]+);/gi,
      (entityMatch, entityCode) =>
        String.fromCodePoint(
          entityCode[0].toLowerCase() === "x"
            ? parseInt(entityCode.slice(1), 16)
            : Number(entityCode),
        ),
    );
  };
  let cellAttributes;
  for (const xmlMatch of xmlText.matchAll(
    /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<(\/?)(?:\w+:)?(c|f)\b([^>]*)>/gi,
  )) {
    if (!xmlMatch[2]) continue;
    if (xmlMatch[2].toLowerCase() === "c") {
      cellAttributes =
        xmlMatch[1] || /\/\s*$/.test(xmlMatch[3]) ? undefined : xmlMatch[3];
      continue;
    }
    if (
      cellAttributes === undefined ||
      xmlMatch[1] ||
      readAttribute(xmlMatch[3], "t") !== "array"
    )
      continue;
    const cellAddress = readAttribute(cellAttributes, "r") ?? "unknown address",
      formulaRange = readAttribute(xmlMatch[3], "ref") ?? cellAddress,
      arrayKindLabel =
        readAttribute(cellAttributes, "cm") !== undefined
          ? "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)"
          : "Legacy CSE array formula";
    throw new Error(
      `${arrayKindLabel} at ${sheetName}!${cellAddress}, range ${formulaRange}. Import stopped before changing the workbook. Open the original in Excel and convert the array to ordinary supported formulas or an explicitly reviewed values snapshot, then import that copy. Fixed CSE ranges and Excel array metadata are not silently treated as dynamic spills.`,
    );
  }
}

const oue = unwrapDefaultExport(iue);

const x7 = unwrapDefaultExport(lue);

function dateSerialToIsoText(raw, dateOrder, dateSystem = "1900") {
  if (
    !Number.isFinite(raw) ||
    raw < 0 ||
    raw >= dateSystemMaxSerial(dateSystem) + 1
  )
    throw Error(
      `Date serial is outside 0–${dateSystemMaxSerial(dateSystem)} for the ${dateSystem} system.`,
    );
  let wholeDays = Math.floor(raw),
    secondsOfDay = Math.round((raw - wholeDays) * 86400);
  if (
    (secondsOfDay === 86400 && (wholeDays++, (secondsOfDay = 0)),
    wholeDays > dateSystemMaxSerial(dateSystem))
  )
    throw Error("Rounded date/time exceeds year 9999.");
  const isoDate =
      dateSystem === "1900" && wholeDays === 60
        ? "1900-02-29"
        : new Date(
            dateSystem === "1904"
              ? Date.UTC(1904, 0, 1) + wholeDays * 864e5
              : Date.UTC(1899, 11, 31) +
                  (wholeDays > 60 ? wholeDays - 1 : wholeDays) * 864e5,
          )
            .toISOString()
            .slice(0, 10),
    [isoYear, isoMonth, isoDay] = isoDate.split("-");
  return (
    (dateOrder === "ymd"
      ? isoDate
      : dateOrder === "dmy"
        ? `${isoDay}/${isoMonth}/${isoYear}`
        : `${isoMonth}/${isoDay}/${isoYear}`) +
    (secondsOfDay
      ? ` ${String(Math.floor(secondsOfDay / 3600)).padStart(2, "0")}:${String(Math.floor(secondsOfDay / 60) % 60).padStart(2, "0")}:${String(secondsOfDay % 60).padStart(2, "0")}`
      : "")
  );
}

function getWorksheetUsedRange(workbook, sheetId) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Worksheet not found.");
  const occupiedRangeBounds = findOccupiedRangeBounds(sheet, false);
  return `A1:${columnIndexToLetters(Math.max(occupiedRangeBounds?.lastCol ?? 0, (sheet.delimitedSource?.columns ?? 1) - 1))}${Math.max(occupiedRangeBounds?.lastRow ?? 0, (sheet.delimitedSource?.records ?? 1) - 1) + 1}`;
}
function encodeExportTextBytes(text, encoding, bom) {
  if (
    !["utf-8", "utf-16le", "utf-16be", "windows-1252"].includes(encoding) ||
    typeof bom != "boolean"
  )
    throw Error("Invalid output encoding or BOM setting.");
  for (const character of text)
    if (
      character.length === 1 &&
      character.charCodeAt(0) >= 55296 &&
      character.charCodeAt(0) <= 57343
    )
      throw Error(
        "Unpaired Unicode surrogate cannot be exported; correct the text first.",
      );
  if (encoding === "utf-8")
    return new TextEncoder().encode((bom ? "\uFEFF" : "") + text);
  if (encoding === "windows-1252") {
    if (bom)
      throw Error(
        "Windows-1252 does not support a BOM; disable BOM or choose UTF-8.",
      );
    const byteByCharacter = new Map(
        byteCharacterTable.map((tableCharacter, byteIndex) => [
          tableCharacter,
          byteIndex,
        ]),
      ),
      bytes = [];
    for (const textCharacter of text) {
      const byteValue = byteByCharacter.get(textCharacter);
      if (byteValue === undefined)
        throw Error(
          "A character cannot be represented in Windows-1252. Choose UTF-8; text is never replaced.",
        );
      bytes.push(byteValue);
    }
    return Uint8Array.from(bytes);
  }
  const textWithBom = (bom ? "\uFEFF" : "") + text,
    utf16Bytes = new Uint8Array(textWithBom.length * 2),
    dataView = new DataView(utf16Bytes.buffer);
  for (
    let codeUnitIndex = 0;
    codeUnitIndex < textWithBom.length;
    codeUnitIndex++
  )
    dataView.setUint16(
      codeUnitIndex * 2,
      textWithBom.charCodeAt(codeUnitIndex),
      encoding === "utf-16le",
    );
  return utf16Bytes;
}
const needsFormulaInjectionGuard = (text) =>
  /^[\s\u0000-\u001f]*[=+\-@＝＋－＠]/u.test(text) || /^[\t\r\n]/.test(text);
function quoteDelimitedField(field, delimitedOptions) {
  const shouldQuote =
    field.includes(delimitedOptions.delimiter) ||
    /[\r\n\uFEFF]/.test(field) ||
    (delimitedOptions.quote !== "none" &&
      field.includes(delimitedOptions.quote)) ||
    /^[ ]|[ ]$/.test(field) ||
    delimitedOptions.quoteAll;
  if (delimitedOptions.quote === "none") {
    if (field.includes(delimitedOptions.delimiter) || /[\r\n]/.test(field))
      throw Error(
        "A field needs quoting; choose a quote character to preserve delimiters/newlines.",
      );
    return field;
  }
  if (!shouldQuote) return field;
  const escapedField =
    delimitedOptions.escape === "double"
      ? field
          .split(delimitedOptions.quote)
          .join(delimitedOptions.quote + delimitedOptions.quote)
      : field
          .replaceAll("\\", "\\\\")
          .split(delimitedOptions.quote)
          .join("\\" + delimitedOptions.quote);
  return delimitedOptions.quote + escapedField + delimitedOptions.quote;
}
function exportWorksheetToCsv(workbook, sheetId, options = {}) {
  const workbookInput = normalizeWorkbookInput(workbook),
    sheet = workbookInput.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Worksheet not found.");
  const delimitedOptions = {
    ...normalizeDelimitedOptions({
      ...options,
      delimiter: options.delimiter ?? ",",
      encoding: options.encoding ?? "utf-8",
      newline:
        options.newline ??
        `\r
`,
    }),
    includeHeader: options.includeHeader ?? true,
    bom: options.bom ?? false,
    quoteAll: options.quoteAll ?? false,
    dateOutput: options.dateOutput ?? "calendar",
    range:
      options.range ??
      options.population?.definition.range ??
      getWorksheetUsedRange(workbookInput, sheetId),
    ...(options.population ? { population: options.population } : {}),
  };
  for (const optionName of ["includeHeader", "bom", "quoteAll"])
    if (typeof delimitedOptions[optionName] != "boolean")
      throw Error(`Invalid export ${optionName} control.`);
  if (!["calendar", "serial"].includes(delimitedOptions.dateOutput))
    throw Error("Invalid date export mode.");
  (delimitedOptions.delimiter === "auto" && (delimitedOptions.delimiter = ","),
    delimitedOptions.encoding === "auto" &&
      (delimitedOptions.encoding = "utf-8"),
    delimitedOptions.newline === "auto" &&
      (delimitedOptions.newline = `\r
`));
  const rangeAddress = parseRangeAddress(delimitedOptions.range),
    rowCount = rangeAddress.lastRow - rangeAddress.firstRow + 1,
    columnCount = rangeAddress.lastCol - rangeAddress.firstCol + 1;
  if (rowCount * columnCount > workbookLimits.maxCells)
    throw Error(
      "Delimited/CSV export would expand more than 100,000 cells including blank gaps. Choose a smaller range or use native JSON.",
    );
  const delimitedImportReport = buildDelimitedImportReport(
      "export",
      `${sheet.name}.${delimitedOptions.delimiter === "	" ? "tsv" : "csv"}`,
      0,
      delimitedOptions,
    ),
    workbookWithSettings = calculateWorkbookWithSettings(workbookInput),
    recordTexts = [];
  if (delimitedOptions.population) {
    if (
      !["full", "visible"].includes(delimitedOptions.population.scope) ||
      !delimitedOptions.hasHeader
    )
      throw Error(
        "Filtered exports require a header and an explicit full or visible scope.",
      );
    const context = applyFilterConditions(
      workbookInput,
      sheetId,
      delimitedOptions.population.definition,
      workbookWithSettings,
    ).context;
    if (
      delimitedOptions.range !== context.range &&
      delimitedOptions.range !== delimitedOptions.population.definition.range
    )
      throw Error(
        "The export range must match the declared filter population.",
      );
    ((delimitedImportReport.population = {
      ...context,
      scope: delimitedOptions.population.scope,
      exportedSourceRows:
        delimitedOptions.population.scope === "visible"
          ? context.visibleSourceRows
          : context.sourceRows,
    }),
      delimitedImportReport.warnings.push(
        `Population scope: ${delimitedOptions.population.scope}; exporting ${delimitedImportReport.population.exportedSourceRows.length} of ${context.totalRows} source data rows. Save the schema report for filter criteria, row locators and source IDs. Session filter criteria are not encoded in CSV/TSV.`,
      ));
  } else
    delimitedImportReport.warnings.push(
      "Population scope: full selected rectangle, including hidden and filtered rows.",
    );
  const exportedSourceRowSet = delimitedImportReport.population
    ? new Set(delimitedImportReport.population.exportedSourceRows)
    : null;
  ((delimitedImportReport.columns = columnCount),
    delimitedImportReport.warnings.push(
      "Values-only export: formulas are evaluated; formatting, notes, history and workbook metadata are not carried by CSV/TSV. Reimport with the visible schema; the file itself has no type schema.",
      "Safe export always prefixes an apostrophe to formula-like text, including whitespace/control-prefixed and full-width triggers. Numeric negatives stay numeric. Quotes alone do not neutralize spreadsheet formulas. Changed text is listed below; it is not automatically unescaped on reimport.",
    ),
    sheet.delimitedSource &&
      delimitedImportReport.warnings.push(
        "The default range includes the historical imported footprint so trailing blank records survive. Review the range after structural edits; this is not a live table binding.",
      ));
  let totalBytes = 0;
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  ) {
    if (
      firstRow !== rangeAddress.firstRow &&
      exportedSourceRowSet &&
      !exportedSourceRowSet.has(firstRow + 1)
    )
      continue;
    const record = firstRow - rangeAddress.firstRow + 1,
      isHeaderRow =
        delimitedOptions.hasHeader && firstRow === rangeAddress.firstRow,
      fields = [],
      raw = [],
      values = [];
    let isBlankRecord = true;
    for (let columnOffset = 0; columnOffset < columnCount; columnOffset++) {
      const cellAddress = `${columnIndexToLetters(rangeAddress.firstCol + columnOffset)}${firstRow + 1}`,
        storedCell = sheet.cells[cellAddress],
        cellValue = workbookWithSettings(sheetId, cellAddress),
        columnType = delimitedOptions.columns[columnOffset] ?? "auto";
      if (
        (firstRow === rangeAddress.firstRow &&
          delimitedImportReport.schema.push({
            index: columnOffset,
            label: delimitedOptions.hasHeader
              ? String(cellValue ?? "")
              : columnIndexToLetters(rangeAddress.firstCol + columnOffset),
            type: columnType,
            observed: [],
          }),
        isHeaderRow && !delimitedOptions.includeHeader)
      )
        continue;
      let cellText = String(cellValue ?? ""),
        isTextValue = typeof cellValue == "string";
      try {
        const cell =
          isHeaderRow || columnType === "auto"
            ? { raw: cellValue }
            : typeof cellValue == "number" &&
                ["number", "integer", "date"].includes(columnType)
              ? { raw: cellValue }
              : parseDelimitedCellValue(
                  cellText,
                  columnType,
                  delimitedOptions,
                  workbookInput.datePolicy?.dateSystem,
                );
        if (
          !isHeaderRow &&
          columnType === "integer" &&
          typeof cell.raw == "number" &&
          !Number.isSafeInteger(cell.raw)
        )
          throw Error("Integer column requires a safe whole number.");
        if (
          ((isTextValue = typeof cell.raw == "string"),
          (cellText = String(cell.raw ?? "")),
          typeof cell.raw == "number" &&
            (!isHeaderRow &&
            (columnType === "date" ||
              (columnType === "auto" && storedCell?.format === "date")) &&
            delimitedOptions.dateOutput === "calendar"
              ? (cellText = dateSerialToIsoText(
                  cell.raw,
                  delimitedOptions.dateOrder,
                  workbookInput.datePolicy?.dateSystem,
                ))
              : (cellText = cellText.replace(".", delimitedOptions.decimal))),
          isTextValue && needsFormulaInjectionGuard(cellText))
        ) {
          const originalText = cellText;
          ((cellText = "'" + cellText),
            delimitedImportReport.safeTextChanges++,
            recordDelimitedChange(
              delimitedImportReport,
              record,
              columnOffset + 1,
              cellAddress,
              "Formula-like text prefixed with apostrophe",
              originalText,
              cellText,
            ));
        } else
          cellText !== String(cellValue ?? "") &&
            recordDelimitedChange(
              delimitedImportReport,
              record,
              columnOffset + 1,
              cellAddress,
              "Selected type, number or date representation",
              String(cellValue ?? ""),
              cellText,
            );
        const observedType = cell.raw === null ? "blank" : typeof cell.raw;
        if (
          (delimitedImportReport.schema[columnOffset].observed.includes(
            observedType,
          ) ||
            delimitedImportReport.schema[columnOffset].observed.push(
              observedType,
            ),
          cellText.length > fileImportLimits.maxCellTextLength)
        )
          throw Error(
            "Safe-export prefix would exceed the 32,767-character field limit. Shorten the text before exporting.",
          );
        (cellText !== "" && (isBlankRecord = false),
          fields.push(quoteDelimitedField(cellText, delimitedOptions)));
      } catch (error) {
        (recordDelimitedIssue(
          delimitedImportReport,
          record,
          columnOffset + 1,
          error instanceof Error ? error.message : "Export conversion failed.",
          cellText,
        ),
          fields.push(""));
      }
      (cellValue !== null && delimitedImportReport.storedCells++,
        columnOffset < 200 &&
          (raw.push(String(cellValue ?? "")), values.push(cellText)));
    }
    if (isHeaderRow && !delimitedOptions.includeHeader) continue;
    const joinedFields = fields.join(delimitedOptions.delimiter);
    isBlankRecord && delimitedImportReport.blankRecords++;
    const recordText =
      joinedFields === "" && delimitedOptions.quote !== "none"
        ? delimitedOptions.quote + delimitedOptions.quote
        : joinedFields;
    if (
      ((totalBytes +=
        new TextEncoder().encode(recordText).length +
        (recordTexts.length ? delimitedOptions.newline.length : 0)),
      totalBytes > workbookLimits.maxFileBytes * 3)
    )
      throw Error(
        "Delimited export exceeds the 10 MB output limit. Choose a smaller range.",
      );
    (recordTexts.push(recordText),
      delimitedImportReport.records++,
      delimitedImportReport.preview.length < 25 &&
        delimitedImportReport.preview.push({
          record: record,
          raw: raw,
          values: values,
        }));
  }
  let csvText = recordTexts.join(delimitedOptions.newline);
  recordTexts.at(-1) === "" &&
    delimitedOptions.quote === "none" &&
    (csvText += delimitedOptions.newline);
  let encodedBytes = new Uint8Array();
  try {
    encodedBytes = encodeExportTextBytes(
      csvText,
      delimitedOptions.encoding,
      delimitedOptions.bom,
    );
  } catch (error) {
    recordDelimitedIssue(
      delimitedImportReport,
      0,
      undefined,
      error instanceof Error ? error.message : "Encoding failure.",
    );
  }
  if (encodedBytes.length > workbookLimits.maxFileBytes)
    throw Error(
      "CSV/TSV export exceeds the 10 MB output limit. Choose a smaller range or native JSON.",
    );
  return (
    (delimitedImportReport.fileBytes = encodedBytes.length),
    (delimitedImportReport.valid = delimitedImportReport.issueCount === 0),
    (delimitedImportReport.complete = true),
    {
      text: csvText,
      bytes: delimitedImportReport.valid ? encodedBytes : new Uint8Array(),
      report: delimitedImportReport,
    }
  );
}

function importXlsxWorkbook(bytes, importResult) {
  const packageEntries = inspectXlsxPackage(bytes),
    entryNames = new Set(packageEntries.keys()),
    size = packageEntries.size,
    workbookParseState = createWorkbookParseState();
  if (
    [...entryNames].some((entryName) =>
      /(?:vbaProject\.bin|xl\/macrosheets\/|xl\/activeX\/|xl\/embeddings\/)/i.test(
        entryName,
      ),
    )
  )
    throw new Error(
      "Workbooks containing macros, ActiveX or embedded objects are not accepted. Save a clean XLSX copy first.",
    );
  let expandedBytes = 0,
    archiveError,
    entryCount = 0,
    completedEntries = 0,
    mergedCellTotal = 0,
    validationCellTotal = 0,
    mergeCount = 0,
    hiddenRowCount = 0,
    m = 0,
    hasInvalidFontSize = false;
  const seenEntryNames = new Set(),
    unzipper = new Unzip((entry) => {
      if (
        (entryCount++,
        !entryNames.has(entry.name) || seenEntryNames.has(entry.name))
      ) {
        archiveError = new Error("Archive entries do not match its directory.");
        return;
      }
      seenEntryNames.add(entry.name);
      let entryBytes = 0;
      const isXmlEntry = /\.(?:xml|rels|vml)$/i.test(entry.name),
        xmlDecoder = isXmlEntry
          ? new TextDecoder("utf-8", { fatal: true })
          : undefined;
      let crc = -1,
        xmlText = "";
      ((entry.ondata = (chunkError, chunk, isFinalChunk) => {
        if (chunkError) {
          archiveError = new Error(
            `Corrupt XLSX compressed content in ${entry.name}. Repair the original workbook and save a clean XLSX copy.`,
          );
          return;
        }
        if (
          ((expandedBytes += chunk.length),
          (entryBytes += chunk.length),
          expandedBytes > maxXlsxExpandedByteLimit)
        ) {
          (entry.terminate(),
            (archiveError = new Error(
              "The XLSX archive expands beyond 50 MB.",
            )));
          return;
        }
        if (entryBytes > packageEntries.get(entry.name).size) {
          (entry.terminate(),
            (archiveError = new Error(
              "The XLSX archive expansion does not match its directory.",
            )));
          return;
        }
        crc = computeCrc32(crc, chunk);
        try {
          xmlDecoder &&
            (xmlText += xmlDecoder.decode(chunk, { stream: !isFinalChunk }));
        } catch {
          archiveError = new Error(
            `Unsupported XML encoding in ${entry.name}. Save a standard UTF-8 XLSX copy.`,
          );
          return;
        }
        if (isFinalChunk) {
          if (entryBytes !== packageEntries.get(entry.name).size) {
            archiveError = new Error(
              "The XLSX archive expansion does not match its directory.",
            );
            return;
          }
          if ((crc ^ -1) >>> 0 !== packageEntries.get(entry.name).crc) {
            archiveError = new Error(
              `XLSX checksum mismatch in ${entry.name}. The archive is corrupt; repair the original and save a clean XLSX copy.`,
            );
            return;
          }
          try {
            if (
              (isXmlEntry &&
                parseWorksheetXml(
                  xmlText,
                  entry.name,
                  packageEntries,
                  workbookParseState,
                ),
              /^xl\/styles\.xml$/i.test(entry.name))
            )
              for (const fontSizeMatch of xmlText.matchAll(
                /<(?:\w+:)?sz\b([^>]*)>/g,
              )) {
                const fontSize = Number(
                  getXmlAttribute(fontSizeMatch[1], "val"),
                );
                (!Number.isInteger(fontSize) ||
                  fontSize < 8 ||
                  fontSize > 72) &&
                  (hasInvalidFontSize = true);
              }
            if (/^xl\/worksheets\//i.test(entry.name)) {
              (classifyArrayFormula(xmlText, entry.name),
                (m = parseWorksheetCells(xmlText, entry.name, m)),
                /^xl\/worksheets\/[^/]+\.xml$/i.test(entry.name) &&
                  importResult.numericTokens.set(
                    entry.name,
                    scanXlsxNumericTokens(xmlText),
                  ));
              for (const pageSetupMatch of xmlText.matchAll(
                /<(?:\w+:)?pageSetup\b([^>]*)>/gi,
              )) {
                const xmlAttribute = getXmlAttribute(
                  pageSetupMatch[1],
                  "pageOrder",
                );
                if (
                  xmlAttribute !== undefined &&
                  !["overThenDown", "downThenOver"].includes(
                    decodeXmlEntities(xmlAttribute),
                  )
                )
                  throw new Error(
                    "Unsupported XLSX page order. Import stopped to preserve the source layout.",
                  );
              }
              const printBreaks = {};
              for (const [
                breakElementName,
                breakKey,
                breakLimit,
                crossAxisLimit,
              ] of [
                [
                  "rowBreaks",
                  "rowBreaks",
                  workbookLimits.maxRows,
                  workbookLimits.maxCols,
                ],
                [
                  "colBreaks",
                  "columnBreaks",
                  workbookLimits.maxCols,
                  workbookLimits.maxRows,
                ],
              ]) {
                const breakSectionMatches = [
                  ...xmlText.matchAll(
                    new RegExp(
                      "<(?:\\w+:)?" +
                        breakElementName +
                        "\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?" +
                        breakElementName +
                        ">",
                      "gi",
                    ),
                  ),
                ];
                if (breakSectionMatches.length > 1)
                  throw Error(
                    "Duplicate XLSX page-break sections are not supported.",
                  );
                const breakBoundaries = [];
                let breakCount = 0;
                for (const breakMatch of (
                  breakSectionMatches[0]?.[1] ?? ""
                ).matchAll(/<(?:\w+:)?brk\b([^>]*)\/?\s*>/gi)) {
                  if (++breakCount > 100)
                    throw Error(
                      "XLSX page breaks exceed 100 boundaries per axis.",
                    );
                  const xmlAttribute = getXmlAttribute(breakMatch[1], "id"),
                    manualAttribute = getXmlAttribute(breakMatch[1], "man");
                  if (
                    !/^[1-9]\d*$/.test(xmlAttribute ?? "") ||
                    !Number.isSafeInteger(Number(xmlAttribute)) ||
                    Number(xmlAttribute) >= breakLimit
                  )
                    throw Error("Invalid XLSX page-break boundary.");
                  if (manualAttribute !== "1" && manualAttribute !== "true")
                    continue;
                  const spanStart = Number(
                      getXmlAttribute(breakMatch[1], "min") ?? 0,
                    ),
                    spanEnd = Number(
                      getXmlAttribute(breakMatch[1], "max") ??
                        crossAxisLimit - 1,
                    );
                  if (
                    spanStart !== 0 ||
                    !Number.isSafeInteger(spanEnd) ||
                    spanEnd < crossAxisLimit - 1
                  )
                    throw Error(
                      "Partial-span XLSX page breaks are not supported. Use full-row or full-column breaks.",
                    );
                  breakBoundaries.push(Number(xmlAttribute));
                }
                breakBoundaries.sort(
                  (leftBoundary, rightBoundary) => leftBoundary - rightBoundary,
                );
                const validatedBreaks = validatePrintPageBreaks(
                  breakBoundaries,
                  breakLimit,
                );
                validatedBreaks && (printBreaks[breakKey] = validatedBreaks);
              }
              (importResult.printBreaks.set(entry.name, printBreaks),
                importResult.conditionalStops.set(
                  entry.name,
                  validateConditionalFormatRules(xmlText),
                ),
                importResult.validationAreas.set(
                  entry.name,
                  parseXlsxDataValidations(xmlText),
                ));
              const hiddenRowIntervals = [],
                columnMetadata = [];
              for (const columnMatch of xmlText.matchAll(
                /<(?:\w+:)?col\b([^>]*)>/gi,
              )) {
                const xmlAttribute = getXmlAttribute(columnMatch[1], "min"),
                  maxAttribute = getXmlAttribute(columnMatch[1], "max"),
                  widthAttribute = getXmlAttribute(columnMatch[1], "width"),
                  start = Number(xmlAttribute) - 1,
                  end = Number(maxAttribute) - 1;
                if (
                  !/^[1-9]\d*$/.test(xmlAttribute ?? "") ||
                  !/^[1-9]\d*$/.test(maxAttribute ?? "") ||
                  !Number.isSafeInteger(start) ||
                  !Number.isSafeInteger(end) ||
                  start < 0 ||
                  end < start ||
                  end >= workbookLimits.maxCols ||
                  columnMetadata.length >= workbookLimits.maxCols
                )
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
                const width =
                  widthAttribute === undefined
                    ? undefined
                    : Number(widthAttribute);
                if (
                  width !== undefined &&
                  (!widthAttribute?.trim() ||
                    !Number.isFinite(width) ||
                    width < 0)
                )
                  throw Error(
                    "XLSX column widths must be finite nonnegative values.",
                  );
                columnMetadata.push({
                  start: start,
                  end: end,
                  ...(width === undefined ? {} : { width: width }),
                });
              }
              columnMetadata.sort(
                (leftRange, rightRange) => leftRange.start - rightRange.start,
              );
              for (
                let columnIndex = 1;
                columnIndex < columnMetadata.length;
                columnIndex++
              )
                if (
                  columnMetadata[columnIndex].start <=
                  columnMetadata[columnIndex - 1].end
                )
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
              if (
                (importResult.columns.set(entry.name, columnMetadata),
                /<(?:\w+:)?sheetFormatPr\b[^>]*\bzeroHeight\s*=\s*["'](?:1|true)["']/i.test(
                  xmlText,
                ))
              )
                throw new Error(
                  "Default-hidden worksheet rows cannot be imported safely. Replace the default with explicit hidden rows before importing.",
                );
              for (const rowMatch of xmlText.matchAll(
                /<(?:\w+:)?row\b([^>]*)>/gi,
              )) {
                if (!/\bhidden\s*=\s*["'](?:1|true)["']/i.test(rowMatch[1]))
                  continue;
                const rowIndex = Number(getXmlAttribute(rowMatch[1], "r")) - 1;
                if (
                  !Number.isInteger(rowIndex) ||
                  rowIndex < 0 ||
                  rowIndex >= workbookLimits.maxRows
                )
                  throw new Error(
                    "Hidden XLSX rows must have valid row numbers.",
                  );
                const previousInterval = hiddenRowIntervals.at(-1);
                if (previousInterval && rowIndex <= previousInterval.end)
                  throw new Error(
                    "Hidden XLSX row numbers must be distinct and ascending.",
                  );
                if (
                  (previousInterval && rowIndex === previousInterval.end + 1
                    ? (previousInterval.end = rowIndex)
                    : hiddenRowIntervals.push({
                        start: rowIndex,
                        end: rowIndex,
                      }),
                  hiddenRowIntervals.length >
                    worksheetCapacityLimits.maxIntervalsPerAxis)
                )
                  throw new Error(
                    "Hidden metadata exceeds 10,000 intervals per worksheet axis.",
                  );
                if (
                  ++hiddenRowCount > worksheetCapacityLimits.maxXlsxHiddenRows
                )
                  throw new Error(
                    "XLSX hidden rows exceed the 100,000-row processing budget. Reduce hidden rows in the source; native JSON preserves larger visibility intervals.",
                  );
              }
              hiddenRowIntervals.length &&
                importResult.hiddenRows.set(entry.name, hiddenRowIntervals);
            }
            if (entry.name === "xl/workbook.xml")
              for (const sheetMatch of xmlText.matchAll(
                /<(?:\w+:)?sheet\b([^>]*)>/g,
              )) {
                const xmlAttribute = getXmlAttribute(sheetMatch[1], "name"),
                  relationshipId = getXmlAttribute(sheetMatch[1], "r:id");
                xmlAttribute &&
                  relationshipId &&
                  importResult.sheetNames.set(
                    decodeXmlEntities(xmlAttribute),
                    decodeXmlEntities(relationshipId),
                  );
              }
            if (entry.name === "xl/_rels/workbook.xml.rels")
              for (const relationshipMatch of xmlText.matchAll(
                /<(?:\w+:)?Relationship\b([^>]*)>/g,
              )) {
                const xmlAttribute = getXmlAttribute(
                    relationshipMatch[1],
                    "Id",
                  ),
                  targetPath = getXmlAttribute(relationshipMatch[1], "Target");
                xmlAttribute &&
                  targetPath &&
                  getXmlAttribute(relationshipMatch[1], "TargetMode") !==
                    "External" &&
                  importResult.sheetRelationships.set(
                    decodeXmlEntities(xmlAttribute),
                    resolvePackagePartPath(decodeXmlEntities(targetPath)),
                  );
              }
            for (const rangeMatch of xmlText.matchAll(
              /<(?:\w+:)?(mergeCell|dataValidation)\s[^>]*?\b(ref|sqref)\s*=\s*["']([^"']+)["'][^>]*>/g,
            ))
              for (const rangeText of decodeXmlEntities(rangeMatch[3])
                .trim()
                .split(/\s+/)) {
                const rangeAddress = parseRangeAddress(rangeText),
                  cellCount =
                    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
                    (rangeAddress.lastCol - rangeAddress.firstCol + 1);
                if (rangeMatch[1] === "mergeCell") {
                  if (
                    cellCount > 500 ||
                    ++mergeCount > 1e3 ||
                    (mergedCellTotal += cellCount) > workbookLimits.maxCells
                  )
                    throw new Error(
                      "Merged ranges exceed the supported import budget (500 cells per merge and 100,000 cells overall).",
                    );
                } else if (
                  (validationCellTotal += cellCount) > workbookLimits.maxCells
                )
                  throw new Error(
                    "XLSX validation ranges exceed the 100,000-cell import budget. Narrow the validation area in the source workbook.",
                  );
              }
            if (entry.name === "xl/workbook.xml") {
              importResult.calculation = parseCalculationProperties(xmlText);
              const parsedNames = aP(xmlText);
              ((importResult.names = parsedNames.names),
                (importResult.nameWarnings = parsedNames.warnings),
                (importResult.builtinNameXml = parsedNames.builtinXml));
            }
          } catch (error) {
            archiveError =
              error instanceof Error
                ? error
                : new Error("Invalid XLSX worksheet metadata.");
            return;
          }
          completedEntries++;
        }
      }),
        entry.start());
    });
  unzipper.register(UnzipInflate);
  for (let byteOffset = 0; byteOffset < bytes.length; byteOffset += 2048) {
    if (archiveError) throw archiveError;
    try {
      unzipper.push(
        bytes.subarray(byteOffset, Math.min(byteOffset + 2048, bytes.length)),
        byteOffset + 2048 >= bytes.length,
      );
    } catch {
      throw new Error(
        "Corrupt XLSX compressed content. Repair the original workbook and save a clean XLSX copy.",
      );
    }
  }
  if (archiveError) throw archiveError;
  if (entryCount !== size || completedEntries !== size)
    throw new Error("The XLSX archive contains incomplete entries.");
  const importWarnings = [
    xlsxExchangeSupportNotice,
    ...importResult.nameWarnings,
    ...validateWorkbookPackage(workbookParseState, entryNames),
  ];
  return (
    (importResult.emptyStringCaches = workbookParseState.emptyStringCaches),
    hasInvalidFontSize &&
      importWarnings.push(
        "Unsupported fractional or out-of-range font sizes were found in the source. The XLSX reader may normalize them; the editor supports integer font sizes from 8 to 72 points.",
      ),
    [...entryNames].some((partName) =>
      /externalLinks|connections\.xml/i.test(partName),
    ) &&
      importWarnings.push(
        "External workbook links or connections were detected. They will not be refreshed or included in the derived copy; the original retains their source parts.",
      ),
    importWarnings
  );
}
const expandHexColorToArgb = (hexColor) => {
    const hexDigits = hexColor.slice(1);
    return `FF${hexDigits.length === 3 ? [...hexDigits].map((digit) => digit + digit).join("") : hexDigits}`;
  },
  builtinNumberFormatCodes = {
    general: "General",
    number: "#,##0.00",
    currency: "$#,##0.00;[Red]($#,##0.00)",
    percent: "0.00%",
    date: "yyyy-mm-dd",
  };
function serializeWorksheetXml(xlsxBytes, sheets) {
  if (
    !sheets.some(
      (sheet) => sheet.hiddenRows?.length || sheet.hiddenCols?.length,
    )
  )
    return xlsxBytes;
  const zipEntries = unzipSync(xlsxBytes);
  return (
    sheets.forEach((sheet, sheetIndex) => {
      const hiddenAxisIntervals = readHiddenAxisIntervals(sheet, "row");
      if (!hiddenAxisIntervals.length && !sheet.hiddenCols?.length) return;
      const sheetPath = "xl/worksheets/sheet" + (sheetIndex + 1) + ".xml";
      let sheetXml = strFromU8(zipEntries[sheetPath]);
      if (
        ((sheetXml = sheetXml.replace(
          /<col\b([^>]*)\/>/g,
          (colTag, colAttributes) => {
            const firstColumnIndex =
                Number(getXmlAttribute(colAttributes, "min")) - 1,
              lastColumnIndex =
                Number(getXmlAttribute(colAttributes, "max")) - 1;
            if (
              !Number.isInteger(firstColumnIndex) ||
              !Number.isInteger(lastColumnIndex) ||
              firstColumnIndex < 0 ||
              lastColumnIndex < firstColumnIndex ||
              lastColumnIndex >= workbookLimits.maxCols
            )
              throw Error(
                "The generated worksheet has invalid column metadata.",
              );
            const retainedColAttributes = colAttributes
                .replace(
                  /\s+(?:min|max|width|customWidth)\s*=\s*(["'])[\s\S]*?\1/g,
                  "",
                )
                .replace(/\/\s*$/, ""),
              colElements = [];
            for (
              let runStart = firstColumnIndex;
              runStart <= lastColumnIndex;
            ) {
              const columnWidth = sheet.columnWidths?.[runStart];
              let runEnd = runStart;
              for (
                ;
                runEnd < lastColumnIndex &&
                sheet.columnWidths?.[runEnd + 1] === columnWidth;
              )
                runEnd++;
              (colElements.push(
                "<col" +
                  retainedColAttributes +
                  ' min="' +
                  (runStart + 1) +
                  '" max="' +
                  (runEnd + 1) +
                  '"' +
                  (columnWidth === undefined
                    ? ' customWidth="0"'
                    : ' width="' + columnWidth / 7 + '" customWidth="1"') +
                  "/>",
              ),
                (runStart = runEnd + 1));
            }
            return colElements.join("");
          },
        )),
        !hiddenAxisIntervals.length)
      ) {
        zipEntries[sheetPath] = strToU8(sheetXml);
        return;
      }
      const sheetDataMatch =
        /<sheetData\b[^>]*\/>|<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(
          sheetXml,
        );
      if (!sheetDataMatch)
        throw new Error("The generated worksheet is missing its data section.");
      const rowXmlByNumber = new Map();
      for (const rowMatch of (sheetDataMatch[1] ?? "").matchAll(
        /<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/g,
      )) {
        const rowNumber = Number(
          getXmlAttribute(/^<row\b([^>]*)>/.exec(rowMatch[0])?.[1] ?? "", "r"),
        );
        if (
          !Number.isInteger(rowNumber) ||
          rowNumber < 1 ||
          rowNumber > workbookLimits.maxRows
        )
          throw new Error("The generated worksheet has an invalid row number.");
        rowXmlByNumber.set(rowNumber, rowMatch[0]);
      }
      for (const hiddenAxisInterval of hiddenAxisIntervals)
        for (
          let start = hiddenAxisInterval.start;
          start <= hiddenAxisInterval.end;
          start++
        ) {
          const oneBasedRowNumber = start + 1,
            existingRowXml = rowXmlByNumber.get(oneBasedRowNumber);
          rowXmlByNumber.set(
            oneBasedRowNumber,
            existingRowXml
              ? existingRowXml.replace(
                  /^<row\b([^>]*)>/,
                  (rowTag, rowAttributes) =>
                    "<row" +
                    rowAttributes
                      .replace(/\s+hidden\s*=\s*(["'])[\s\S]*?\1/g, "")
                      .replace(/\/\s*$/, "") +
                    ' hidden="1"' +
                    (/\/\s*$/.test(rowAttributes) ? "/>" : ">"),
                )
              : '<row r="' + oneBasedRowNumber + '" hidden="1"/>',
          );
        }
      const sortedRowsXml = [...rowXmlByNumber]
        .sort(
          ([leftRowNumber], [rightRowNumber]) => leftRowNumber - rightRowNumber,
        )
        .map(([, rowXml]) => rowXml)
        .join("");
      zipEntries[sheetPath] = strToU8(
        sheetXml.replace(
          sheetDataMatch[0],
          "<sheetData>" + sortedRowsXml + "</sheetData>",
        ),
      );
    }),
    zipSync(zipEntries)
  );
}
function applyPageBreakXml(xlsxBytes, sheets) {
  if (
    !sheets.some(
      (sheet) =>
        sheet.pageLayout?.rowBreaks?.length ||
        sheet.pageLayout?.columnBreaks?.length,
    )
  )
    return xlsxBytes;
  const zipEntries = unzipSync(xlsxBytes);
  return (
    sheets.forEach((sheet, sheetIndex) => {
      const pageLayout = sheet.pageLayout;
      if (!pageLayout?.rowBreaks?.length && !pageLayout?.columnBreaks?.length)
        return;
      const sheetPath = `xl/worksheets/sheet${sheetIndex + 1}.xml`;
      let sheetXml = strFromU8(zipEntries[sheetPath]);
      sheetXml = sheetXml.replace(
        /<(?:rowBreaks|colBreaks)\b[^>]*(?:\/>|>[\s\S]*?<\/(?:rowBreaks|colBreaks)>)/g,
        "",
      );
      const breakXml = [
        ["rowBreaks", pageLayout.rowBreaks, 16383],
        ["colBreaks", pageLayout.columnBreaks, 1048575],
      ]
        .map(([breakElementName, breakIds, maxBreakId]) =>
          breakIds?.length
            ? `<${breakElementName} count="${breakIds.length}" manualBreakCount="${breakIds.length}">${breakIds.map((breakId) => `<brk id="${breakId}" min="0" max="${maxBreakId}" man="1"/>`).join("")}</${breakElementName}>`
            : "",
        )
        .join("");
      ((sheetXml = sheetXml.replace(
        /<(?:customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|<\/worksheet>/,
        (insertionAnchor) => breakXml + insertionAnchor,
      )),
        (zipEntries[sheetPath] = strToU8(sheetXml)));
    }),
    zipSync(zipEntries)
  );
}
function applySheetRuleFlags(xlsxBytes, sheets) {
  if (
    !sheets.some(
      (sheet) =>
        sheet.conditionalFormats?.some(
          (conditionalFormat) => conditionalFormat.stopIfTrue,
        ) ||
        sheet.validations?.some(
          (validation) => validation.showDropdown === false,
        ),
    )
  )
    return xlsxBytes;
  const zipEntries = unzipSync(xlsxBytes);
  return (
    sheets.forEach((sheet, sheetIndex) => {
      const sheetPath = `xl/worksheets/sheet${sheetIndex + 1}.xml`;
      zipEntries[sheetPath] = strToU8(
        applyValidationDropdownFlags(
          appendStopIfTrueAttribute(strFromU8(zipEntries[sheetPath]), sheet),
          sheet,
        ),
      );
    }),
    zipSync(zipEntries)
  );
}
async function exportWorkbookToXlsx(workbook) {
  const xlsxWorkbookPackage = [
    "Dynamic-array formulas are exported as ordinary anchor formulas without Excel dynamic-array metadata. Native JSON retains local spill behavior. Legacy CSE and metadata-qualified Excel arrays are identified and refused on import; convert a reviewed copy in Excel before importing. Independent Excel round-trip acceptance is required.",
    `Date calculations use the Gregorian ${normalizeDatePolicy(workbook.datePolicy).dateSystem} serial system, preserved in XLSX. Native JSON retains the explicit text-date order and calculation time zone; XLSX does not retain that policy or the calculation clock evidence. DATEVALUE/VALUE text interpretation and TODAY/NOW results may change in Excel. Export calculation evidence separately and review calendar-sensitive formulas.`,
    "XLSX exported with supported values, formulas, styles and page layout. Keep native JSON for editable review history, outline groups, links, pivots, charts, shapes and application metadata. Number formats may differ in Excel. External-action formulas are saved as text.",
  ];
  (workbook.styleCatalog ||
    workbook.sheets.some((sheet) =>
      Object.values(sheet.cells).some((cell) => cell.style?.binding),
    )) &&
    xlsxWorkbookPackage.push(
      "Named styles and theme bindings are flattened to their current effective formatting in XLSX. Keep native JSON to preserve reusable styles, inherited values and explicit overrides.",
    );
  const workbookInput = normalizeWorkbookInput(workbook),
    workbookWithSettings = calculateWorkbookWithSettings(workbookInput);
  (workbookInput.workbookLinks &&
    xlsxWorkbookPackage.push(
      "Workbook link targets, reviewed source caches and last-resolved revision evidence are retained in native JSON only. External formulas export as inert text; use Data > Workbook Links to preview breaking them into values before XLSX export. Retained originals are unchanged.",
    ),
    workbookInput.extensions &&
      xlsxWorkbookPackage.push(
        "Custom-function manifests, reviewed result caches and linked entity fields are retained in native JSON only. XLSX does not install or authorize extensions. Custom function formulas require a compatible, separately approved implementation in the receiving application; cached results are historical and not proof of live calculation.",
      ),
    workbookInput.sourceImports?.length &&
      xlsxWorkbookPackage.push(
        "Structured/document source originals, field paths, native-type evidence, page regions and review decisions are native JSON only and omitted from derived XLSX. Keep a native backup and use File > Import sources to retrieve the originals.",
      ),
    workbookInput.alternateSource &&
      xlsxWorkbookPackage.push(
        "The retained alternate-format original and conversion provenance are native JSON only and are omitted from derived XLSX. Keep the original from File > Workbook formats. Conversion equivalence has not been verified.",
      ),
    workbookInput.reportOutputs?.length &&
      xlsxWorkbookPackage.push(
        "Reusable report outputs, fixed snapshots and host-reviewed links are native JSON only and are omitted from XLSX. Export each report as SVG/PNG/HTML/JSON for downstream reuse; XLSX does not create Office OLE links or refresh authority.",
      ),
    workbookInput.embeddedObjectInventories?.length &&
      xlsxWorkbookPackage.push(
        "Unsupported OLE/ActiveX/embedded-object metadata and payloads are omitted from derived XLSX. Every retained inventory item is listed as dropped in embeddedObjectLosses. No object was executed or fetched; preserve the original files separately and native JSON for the inventory.",
      ),
    workbookInput.sheets.some((sheet) => sheet.analysisReport) &&
      xlsxWorkbookPackage.push(
        "Solver and statistical analysis definitions, methods, assumptions, exclusions and freshness evidence require native JSON. XLSX preserves historical report cells only; native Solver models and Analysis ToolPak settings are not generated. Reopen native JSON to review and rerun.",
      ),
    (workbookInput.dataModel ||
      workbookInput.sheets.some((sheet) => sheet.dataModelReport)) &&
      xlsxWorkbookPackage.push(
        "Relational model tables, relationships, typed contracts, measures, queries and source-freshness evidence require native JSON. XLSX exports supported worksheet values only; no Excel Data Model or DAX is generated. Published model results are snapshots and are not recalculated by Excel from these definitions.",
      ),
    workbookInput.importedCalculationCaches?.entries.length &&
      xlsxWorkbookPackage.push(
        "Original imported formula caches, source file hash and cache comparisons are native JSON evidence only. XLSX writes current local results; it never writes imported historical caches as current values.",
      ),
    workbookInput.calculation?.mode === "manual" &&
      xlsxWorkbookPackage.push(
        "Manual calculation mode is exported. Pending formula caches are omitted; choose Calculate Now before relying on formula results in another application.",
      ),
    (workbookInput.automationArtifacts?.length ||
      workbookInput.automationRuns?.length) &&
      xlsxWorkbookPackage.push(
        "Preserved VBA/Office Scripts source bytes and migration run provenance are native JSON only. XLSX contains the resulting cells; it does not embed executable macros or Office Scripts. Download original artifacts separately from Automation & Scripts.",
      ),
    workbookInput.whatIfAnalyses?.length &&
      xlsxWorkbookPackage.push(
        "What-if scenario and sensitivity definitions, retained assumptions and baseline evidence are native JSON metadata. XLSX preserves report cells only; native Excel Scenario Manager and data-table formulas are not created. Keep native JSON to reopen and refresh the analysis.",
      ),
    workbookInput.sheets.some((sheet) => sheet.pivots?.length) &&
      xlsxWorkbookPackage.push(
        "Native PivotTable definitions, slicers, timelines, Show values as settings, field sorting, PivotCharts, calculated fields, grouping, refresh fingerprints and freshness labels are not preserved in XLSX. Only the current output cells are exported; keep native JSON and export pivot context to verify and refresh the result.",
      ),
    (workbookInput.tableEvidence?.length ||
      workbookInput.sheets.some((sheet) =>
        sheet.tables?.some((table) =>
          table.columns.some((column) => column.dataType),
        ),
      )) &&
      xlsxWorkbookPackage.push(
        "Table column types and original header/conversion evidence are native JSON only. XLSX retains supported tables and current cell values; save native JSON to retain the original source values and type contracts.",
      ),
    workbookInput.textCleanups?.length &&
      xlsxWorkbookPackage.push(
        "Text cleanup history, original evidence snapshots and cleanup recipes are retained in native JSON only. XLSX keeps the output values and cell notes but cannot restore the cleanup history.",
      ),
    workbookInput.savedFilterViews?.length &&
      xlsxWorkbookPackage.push(
        "Named filter views and their historical preview counts are retained in native JSON only. XLSX exports the full source population and does not write Excel AutoFilter criteria. Keep native JSON to reopen saved views and review fresh results.",
      ),
    (workbookInput.transformationRecipes?.length ||
      workbookInput.sheets.some((sheet) => sheet.transformationRun)) &&
      xlsxWorkbookPackage.push(
        "Versioned transformation recipes, source/output hashes and complete step execution evidence are native JSON only. XLSX retains current output values and source-row notes; keep the native workbook for repeatable recipe refresh and lineage.",
      ),
    workbookInput.sheets.some((sheet) => sheet.consolidation) &&
      xlsxWorkbookPackage.push(
        "Consolidation definitions, approved external source identities, review evidence and live source-freshness checks require native JSON. XLSX retains supported cells, same-workbook formulas and source notes, but no consolidation refresh or external workbook links. External amounts remain reviewed snapshots; keep native JSON and reapprove source workbooks to refresh.",
      ),
    workbookInput.sheets.some((sheet) => sheet.externalConnection) &&
      xlsxWorkbookPackage.push(
        "External source connections, refresh attempts, source versions, stale state and control-total evidence are native JSON only. XLSX contains a values snapshot and source notes, with no refresh or credential authority. Keep the native workbook and reauthorize the host source before refreshing.",
      ),
    workbookInput.fieldTransforms?.length &&
      xlsxWorkbookPackage.push(
        "Clean & Reshape recipes, reviewed inference examples and original source snapshots are native JSON only. XLSX keeps current output values, derivation notes and the exceptions appendix; save native JSON to retain reversible transformation evidence.",
      ),
    workbookInput.sheets.some((sheet) =>
      sheet.tables?.some((table) =>
        table.columns.some((column) => column.calculatedFormula),
      ),
    ) &&
      xlsxWorkbookPackage.push(
        "Calculated table column rules are retained in native JSON only. XLSX retains the individual cell formulas and values, but does not preserve the reusable calculated-column rule for future appends.",
      ),
    workbookInput.formulaWatches?.length &&
      xlsxWorkbookPackage.push(
        "The formula Watch Window list is retained in native JSON only; XLSX cellWatches are not exported.",
      ),
    workbookInput.sheets.some((sheet) => sheet.splitPanes) &&
      xlsxWorkbookPackage.push(
        "Split panes are retained in native JSON only and are not preserved in XLSX. Save workbook data to keep divider layouts and initial scroll anchors.",
      ),
    workbookInput.rowOrigins?.length &&
      xlsxWorkbookPackage.push(
        "Original source row identities and their current associations are retained in native JSON only. XLSX export does not preserve the source registry; reopening that XLSX creates a new file-source identity.",
      ),
    workbookInput.definedNames?.length &&
      xlsxWorkbookPackage.push(
        "Defined-name formula text, workbook/worksheet scope and supported comments are preserved. Recalculation uses the supported local engine; relative, volatile and dynamic name formulas require independent Excel review.",
      ),
    workbookInput.sheets.some((sheet) => sheet.delimitedSource) &&
      xlsxWorkbookPackage.push(
        "Delimited source footprints and schemas are not carried by XLSX. Keep native JSON for the historical source footprint and the separate delimited schema report for type choices.",
      ));
  for (const sheet of workbookInput.sheets)
    ((sheet.hiddenRows = cn(sheet, "row")),
      (sheet.hiddenCols = cn(sheet, "column")));
  if (
    workbookInput.sheets.reduce(
      (hiddenRowTotal, sheetEntry) =>
        hiddenRowTotal +
        readHiddenAxisIntervals(sheetEntry, "row").reduce(
          (hiddenRowCount, hiddenInterval) =>
            hiddenRowCount + hiddenInterval.end - hiddenInterval.start + 1,
          0,
        ),
      0,
    ) > worksheetCapacityLimits.maxXlsxHiddenRows
  )
    throw new Error(
      "XLSX export exceeds the 100,000-hidden-row processing budget. Use native JSON to preserve larger hidden ranges without expanding them.",
    );
  const excelWorkbook = new oue.Workbook();
  ((excelWorkbook.creator = "LedgerGlass"),
    (excelWorkbook.properties.date1904 =
      normalizeDatePolicy(workbookInput.datePolicy).dateSystem === "1904"));
  const numericCells = workbookInput.sheets
    .flatMap((sheet) => Object.values(sheet.cells))
    .filter(
      (cell) =>
        typeof cell.raw == "number" &&
        validateXlsxNumericValue(String(cell.raw)),
    );
  (numericCells.length &&
    xlsxWorkbookPackage.push(
      `Numeric precision risk: ${numericCells.length} stored numeric value(s) exceed Excel's 15-digit precision or numeric range. Export preserves stored numbers, not precision as displayed. Use explicit ROUND for monetary amounts and Text for identifiers; retain native JSON.`,
    ),
    workbookInput.metadata &&
      ((excelWorkbook.created = new Date(workbookInput.metadata.createdAt)),
      (excelWorkbook.modified = new Date(workbookInput.metadata.modifiedAt))));
  const exchangeSchemaReport = buildExchangeSchemaReport(
    "export",
    `${workbookInput.name}.xlsx`,
    0,
    workbookInput,
    xlsxWorkbookPackage,
  );
  for (const sheet of workbookInput.sheets) {
    const worksheet = excelWorkbook.addWorksheet(sheet.name, {
      state: sheet.veryHidden
        ? "veryHidden"
        : sheet.hidden
          ? "hidden"
          : "visible",
      properties: {
        defaultRowHeight: defaultRowHeightOf(sheet) * 0.75,
        defaultColWidth: defaultColumnWidthOf(sheet) / 7,
        ...(sheet.tabColor
          ? { tabColor: { argb: `FF${sheet.tabColor.slice(1)}` } }
          : {}),
      },
      views:
        sheet.frozenRows || sheet.frozenCols
          ? [
              {
                state: "frozen",
                ySplit: sheet.frozenRows ?? 0,
                xSplit: sheet.frozenCols ?? 0,
              },
            ]
          : [],
    });
    for (const [columnKey, columnWidth] of Object.entries(
      sheet.columnWidths ?? {},
    ))
      worksheet.getColumn(Number(columnKey) + 1).width = columnWidth / 7;
    for (const [rowKey, rowHeight] of Object.entries(sheet.rowHeights ?? {}))
      worksheet.getRow(Number(rowKey) + 1).height = rowHeight * 0.75;
    for (const hiddenColumnInterval of readHiddenAxisIntervals(sheet, "column"))
      for (
        let start = hiddenColumnInterval.start;
        start <= hiddenColumnInterval.end;
        start++
      )
        worksheet.getColumn(start + 1).hidden = true;
    for (const [cellAddress, cell] of Object.entries(sheet.cells)) {
      const worksheetCell = worksheet.getCell(cellAddress);
      if (
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal &&
        !isExternalActionFormula(cell.raw)
      ) {
        const exportedCellValue = toExportedCellValue(
          workbookWithSettings.getTypedValue(sheet.id, cellAddress),
        );
        (exchangeSchemaReport.caches.total++,
          exportedCellValue === undefined
            ? exchangeSchemaReport.caches.omitted++
            : exchangeSchemaReport.caches.written++,
          (worksheetCell.value = {
            formula: cell.raw.slice(1),
            ...(exportedCellValue !== undefined
              ? { result: exportedCellValue }
              : {}),
          }));
      } else
        ((worksheetCell.value = cell.raw),
          typeof cell.raw == "string" &&
            cell.raw.startsWith("=") &&
            !cell.literal &&
            (exchangeSchemaReport.caches.total++,
            exchangeSchemaReport.caches.neutralized++));
      if (
        ((worksheetCell.numFmt = cell.numberFormat
          ? (buildNumberFormatCode(cell.numberFormat) ?? "General")
          : (cell.numFmt ??
            builtinNumberFormatCodes[cell.format ?? "general"])),
        cell.style)
      ) {
        const effectiveCellStyle = resolveEffectiveCellStyle(cell.style);
        ((worksheetCell.font = {
          bold: effectiveCellStyle.bold,
          italic: effectiveCellStyle.italic,
          strike: effectiveCellStyle.strikethrough,
          underline: effectiveCellStyle.underline,
          name: effectiveCellStyle.fontFamily,
          size: effectiveCellStyle.fontSize,
          ...(effectiveCellStyle.color
            ? {
                color: { argb: expandHexColorToArgb(effectiveCellStyle.color) },
              }
            : {}),
        }),
          effectiveCellStyle.fill &&
            (worksheetCell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: expandHexColorToArgb(effectiveCellStyle.fill) },
            }),
          (effectiveCellStyle.align ||
            effectiveCellStyle.wrap !== undefined ||
            effectiveCellStyle.verticalAlign ||
            effectiveCellStyle.indent !== undefined) &&
            (worksheetCell.alignment = {
              horizontal: effectiveCellStyle.align,
              wrapText: effectiveCellStyle.wrap,
              vertical: effectiveCellStyle.verticalAlign,
              indent: effectiveCellStyle.indent,
            }));
        const cellBorders = {};
        for (const cellBorderEdge of cellBorderEdges) {
          const borderEdge = resolveBorderEdge(
            effectiveCellStyle,
            cellBorderEdge,
          );
          borderEdge &&
            borderEdge.style !== "none" &&
            (cellBorders[cellBorderEdge] = {
              style: borderEdge.style,
              color: {
                argb: expandHexColorToArgb(borderEdge.color ?? "#000000"),
              },
            });
        }
        (effectiveCellStyle.border || effectiveCellStyle.borders) &&
          (worksheetCell.border = cellBorders);
      }
      cell.note && (worksheetCell.note = cell.note);
    }
    (buildTableCellEntries(worksheet, sheet),
      serializeWorksheetRuleXml(worksheet, sheet),
      xlsxWorkbookPackage.push(
        ...buildDrawingExportNotes(excelWorkbook, worksheet, sheet),
      ),
      xlsxWorkbookPackage.push(
        ...updateSheetPageSetup(worksheet, sheet, workbookInput),
      ));
  }
  excelWorkbook.calcProperties.fullCalcOnLoad = true;
  const xlsxBuffer = await excelWorkbook.xlsx.writeBuffer(),
    packageEntries = unzipSync(new Uint8Array(xlsxBuffer));
  packageEntries["xl/workbook.xml"] = strToU8(
    serializeCalculationProperties(
      serializeDefinedNames(
        strFromU8(packageEntries["xl/workbook.xml"]),
        workbookInput,
      ),
      workbookInput,
    ),
  );
  const packageBytes = zipSync(packageEntries);
  let bytes = new Uint8Array(
    applySheetRuleFlags(
      applyPageBreakXml(
        serializeWorksheetXml(packageBytes, workbookInput.sheets),
        workbookInput.sheets,
      ),
      workbookInput.sheets,
    ),
  );
  if (workbookInput.xlsxSource) {
    const retainedSource = requireRetainedXlsxSource(workbookInput);
    importXlsxWorkbook(retainedSource.bytes, createEmptyExportLedger());
    const originalEntries = unzipSync(retainedSource.bytes),
      sourcePreservation = F$(originalEntries, retainedSource.sha256),
      preservedPackage = rebuildPreservedPackage(
        originalEntries,
        unzipSync(bytes),
        sourcePreservation,
      );
    try {
      if (preservedPackage.bytes.length > workbookLimits.maxFileBytes)
        throw Error("Preserved metadata exceeds the XLSX reopen budget.");
      (importXlsxWorkbook(preservedPackage.bytes, createEmptyExportLedger()),
        (bytes = new Uint8Array(preservedPackage.bytes)),
        (exchangeSchemaReport.partInventory = preservedPackage.inventory));
    } catch {
      ((exchangeSchemaReport.partInventory = describePreservationDisposition(
        sourcePreservation,
        "Opaque metadata could not pass the final output integrity or capacity checks. Retained only in the original; the derived copy excludes it.",
      )),
        xlsxWorkbookPackage.push(
          "Passive metadata could not be safely included within the export budgets. Inspect the part inventory; download the original for all source content.",
        ));
    }
    xlsxWorkbookPackage.push(
      `${exchangeSchemaReport.partInventory.counts["preserved-opaque"]} source part(s) preserved byte-for-byte as passive metadata. All source bytes remain available through File → Original XLSX and compatibility and native JSON. Derived XLSX does not embed the original source attachment.`,
    );
  }
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  if (blob.size > workbookLimits.maxFileBytes)
    throw new Error(
      "XLSX export exceeds the 10 MB reopen limit. Reduce the workbook or use native JSON.",
    );
  return (
    importXlsxWorkbook(bytes, {
      names: [],
      nameWarnings: [],
      numericTokens: new Map(),
      builtinNameXml: "",
      hiddenRows: new Map(),
      conditionalStops: new Map(),
      validationAreas: new Map(),
      printBreaks: new Map(),
      columns: new Map(),
      sheetRelationships: new Map(),
      sheetNames: new Map(),
    }),
    exchangeSchemaReport.caches.omitted &&
      xlsxWorkbookPackage.push(
        `${exchangeSchemaReport.caches.omitted} formula cache(s) omitted because the local result is blank or an application-specific error. Excel must recalculate them; unsupported results are never written as successful text caches.`,
      ),
    exchangeSchemaReport.caches.neutralized &&
      xlsxWorkbookPackage.push(
        `${exchangeSchemaReport.caches.neutralized} external/action formula(s) exported as literal text.`,
      ),
    (exchangeSchemaReport.fileBytes = blob.size),
    (exchangeSchemaReport.warnings = [...new Set(xlsxWorkbookPackage)]),
    {
      blob: blob,
      warnings: exchangeSchemaReport.warnings,
      report: exchangeSchemaReport,
    }
  );
}
function encodeNativeWorkbook(workbook) {
  const nativeJson = JSON.stringify(normalizeWorkbookInput(workbook), null, 2);
  if (
    new TextEncoder().encode(nativeJson).length > workbookLimits.maxNativeBytes
  )
    throw new Error(
      "Native workbook exceeds the 50 MB save limit. Reduce the workbook size before saving.",
    );
  return nativeJson;
}

export {
  numericPrecisionPolicy,
  F$,
  retainEmbeddedObjectInventory,
  retainSourceArtifact,
  buildExchangeSchemaReport,
  exportXltxTemplate,
  encodeNativeWorkbook,
  delimitedImportDefaults,
  exportWorksheetToCsv,
  importXlsxWorkbook,
  createEmptyExportLedger,
  exportWorkbookToXlsx,
  verifyAutomationInspection,
  inspectEmbeddedObjects,
  x7,
  getWorksheetUsedRange,
};
export { formatRangeCornerCells, countRangeCells };
