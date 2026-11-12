import {
  _t,
  alternateFormatExtensions,
  columnIndexToLetters,
  parseRangeAddress,
  stripFormulaStringLiterals,
  throwWorkbookError,
  validateExtractedSourceTables,
  xlsxInspectionLimits,
} from "./03-workbook-model.js";
import {
  dateSystemMaxSerial,
  dateSystemSerialOffset,
  sha256HexDigest,
} from "./01-schemas-date-functions-and-validation.js";
import {
  excelMessagePatterns,
  fileProcessingErrorMessages,
  unsupportedXlsxMarkers,
} from "./06-unsupported-marker-and-message-catalog.js";
import {
  fileImportLimits,
  workbookLimits,
} from "./02-text-and-financial-functions.js";
import Ox from "saxes";
import { Unzip, UnzipInflate, strFromU8, strToU8, zipSync } from "fflate";

/* From 04-cell-input-paste-and-fill.js */
const formatCellAddress = (rowIndex, columnIndex) =>
  `${columnIndexToLetters(columnIndex)}${rowIndex + 1}`;
const maxExcelDateSerial = 2958465;
function utcDateFromParts(year, monthIndex, dayOfMonth) {
  const date = new Date(0);
  return (
    date.setUTCHours(0, 0, 0, 0),
    date.setUTCFullYear(year, monthIndex, dayOfMonth),
    date
  );
}

/* From 05-add-ins-and-custom-functions.js */
const sourceInspectionRecords = new WeakMap();
function recordSourceInspection(sourceInspection, bytes, fileName) {
  if (
    (validateExtractedSourceTables(sourceInspection),
    sourceInspection.fileName !== fileName ||
      sourceInspection.byteLength !== bytes.length ||
      sourceInspection.sha256 !== sha256HexDigest(bytes))
  )
    throw Error("Source inspection does not match its original.");
  sourceInspectionRecords.set(sourceInspection, {
    signature: JSON.stringify(sourceInspection),
    bytes: bytes.slice(),
  });
}
const formatInspectionRecords = new WeakMap();
function validateIsolatedFormatInspection(inspection, bytes, fileName) {
  if (
    !inspection ||
    inspection.schemaVersion !== 1 ||
    inspection.fileName !== fileName ||
    inspection.fileBytes !== bytes.length ||
    inspection.sha256 !== sha256HexDigest(bytes) ||
    !alternateFormatExtensions.includes(inspection.format) ||
    !fileName.toLowerCase().endsWith("." + inspection.format) ||
    inspection.execution !== "never" ||
    !Array.isArray(inspection.parts) ||
    inspection.parts.length > 2e3
  )
    throw Error("Invalid isolated format inspection.");
  formatInspectionRecords.set(inspection, {
    signature: JSON.stringify(inspection),
    bytes: bytes.slice(),
  });
}

/* From 07-file-import-and-worker-dispatch.js */
const unsupportedArrayFormulaReasons = new Set([
  ...unsupportedXlsxMarkers,
  "Legacy CSE array formula",
  "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)",
]);
const excelRepairSuffix =
  " Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.";
const sanitizedMessagePatterns = excelMessagePatterns.map(
  (excelMessagePattern) => ({
    parts: excelMessagePattern,
    expression: new RegExp(
      "^" +
        excelMessagePattern
          .map((patternPart) =>
            patternPart.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
          )
          .join("([\\s\\S]*?)") +
        "$",
    ),
  }),
);
function sanitizeFileMessage(message, depth = 0) {
  if (!(message.length > 4096 || depth > 3)) {
    if (unsupportedArrayFormulaReasons.has(message)) return message;
    if (message.endsWith(excelRepairSuffix)) {
      const sanitizedBody = sanitizeFileMessage(
        message.slice(0, -excelRepairSuffix.length),
        depth + 1,
      );
      if (sanitizedBody) return sanitizedBody + excelRepairSuffix;
    }
    for (const { parts, expression } of sanitizedMessagePatterns) {
      const match = expression.exec(message);
      if (match)
        return parts
          .map(
            (part, partIndex) =>
              part +
              (partIndex < parts.length - 1
                ? (sanitizeFileMessage(match[partIndex + 1], depth + 1) ??
                  "[source detail omitted]")
                : ""),
          )
          .join("");
    }
  }
}
const u1 = Object.freeze({
  maxFileBytes: workbookLimits.maxFileBytes,
  maxNativeBytes: workbookLimits.maxNativeBytes,
  maxConfigurationBytes: 2e6,
  maxPictureBytes: 2 * 1024 * 1024,
  maxConcurrent: 2,
  defaultTimeoutMs: 3e4,
  maxTimeoutMs: 6e4,
  maxJsonDepth: 128,
  maxJsonNodes: 2e6,
  nodeHeapMiB: 256,
  maxResultBytes: 100 * 1024 * 1024,
});
function fileByteLimitFor(operation, fileName) {
  return operation === "source"
    ? 4 * 1024 * 1024
    : operation === "json"
      ? u1.maxConfigurationBytes
      : operation === "picture"
        ? u1.maxPictureBytes
        : operation === "workbook" && /\.json$/i.test(fileName)
          ? workbookLimits.maxNativeBytes
          : workbookLimits.maxFileBytes;
}
function assertXlsxExpansionLimit(compressedSize, size, expandedTotal) {
  if (
    expandedTotal > fileImportLimits.maxXlsxExpandedBytes ||
    (size > fileImportLimits.xlsxRatioThresholdBytes &&
      size / Math.max(compressedSize, 1) >
        fileImportLimits.maxXlsxExpansionRatio)
  )
    throw new Error(
      "The XLSX archive exceeds the safe expansion limit (50 MB or 200:1 ratio).",
    );
}
function parseWorksheetCells(xmlText, worksheetPath, initialCellCount = 0) {
  const readAttribute = (attributeText, attributeName) =>
      new RegExp(`(?:^|\\s)${attributeName}\\s*=\\s*(["'])(.*?)\\1`).exec(
        attributeText,
      )?.[2],
    strippedXml = xmlText.replace(
      /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g,
      "",
    );
  let currentRow = 0,
    previousRow = 0,
    lastColumn = -1,
    cellRecordCount = initialCellCount;
  const stopImport = (message) => {
    throw new Error(
      `Worksheet ${worksheetPath}: ${message} Import stopped; no cells were truncated.`,
    );
  };
  for (const match of strippedXml.matchAll(
    /<(\/?)(?:[\w.-]+:)?(dimension|row|c)\b([^>]*)>/g,
  )) {
    const [, isClosingTag, tagName, attributes] = match;
    if (isClosingTag) {
      tagName === "row" && (currentRow = 0);
      continue;
    }
    if (tagName === "dimension") {
      const dimensionRef = readAttribute(attributes, "ref");
      dimensionRef ||
        stopImport("The declared dimension has no valid cell range.");
      try {
        parseRangeAddress(dimensionRef);
      } catch {
        stopImport(
          `Declared dimension ${dimensionRef} exceeds or violates the supported worksheet bounds.`,
        );
      }
    } else if (tagName === "row") {
      const rowNumberText = readAttribute(attributes, "r"),
        rowNumber = Number(rowNumberText);
      ((!rowNumberText ||
        !/^[1-9]\d*$/.test(rowNumberText) ||
        rowNumber > fileImportLimits.maxRows) &&
        stopImport(
          `Row ${rowNumberText ?? "(missing)"} must be between 1 and ${fileImportLimits.maxRows}.`,
        ),
        (currentRow || rowNumber <= previousRow) &&
          stopImport("Rows must have distinct, ascending row numbers."),
        (previousRow = rowNumber),
        (currentRow = /\/\s*$/.test(attributes) ? 0 : rowNumber),
        (lastColumn = -1));
    } else {
      const cellAddress = readAttribute(attributes, "r");
      (!cellAddress || !currentRow) &&
        stopImport(
          "Each cell must have an explicit address inside its source row.",
        );
      let cellPosition;
      try {
        cellPosition = _t(cellAddress);
      } catch {
        stopImport(
          `Cell ${cellAddress} exceeds or violates the supported worksheet bounds.`,
        );
      }
      ((cellPosition.row + 1 !== currentRow ||
        cellPosition.col <= lastColumn) &&
        stopImport(
          `Cell ${cellAddress} has a duplicate, unordered or mismatched row address.`,
        ),
        (lastColumn = cellPosition.col),
        ++cellRecordCount > fileImportLimits.maxXlsxCellRecords &&
          stopImport(
            `The XLSX exceeds ${fileImportLimits.maxXlsxCellRecords.toLocaleString("en-US")} source cell records, including explicit blank cells.`,
          ));
    }
  }
  return cellRecordCount;
}
const excelRepairAdvice =
  "Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.";
const throwInvalidXlsx = (message) => {
  throw Error(`${message} ${excelRepairAdvice}`);
};
const crc32Table = new Uint32Array(256).map((element, crcValue) => {
  for (let bitIndex = 0; bitIndex < 8; bitIndex++)
    crcValue = crcValue & 1 ? 3988292384 ^ (crcValue >>> 1) : crcValue >>> 1;
  return crcValue >>> 0;
});
function computeCrc32(crc, bytes) {
  for (const byte of bytes) crc = crc32Table[(crc ^ byte) & 255] ^ (crc >>> 8);
  return crc;
}
function inspectXlsxPackage(bytes, options = {}) {
  const zipDirectory = parseZipDirectory(bytes);
  return (
    (!zipDirectory.has("[Content_Types].xml") ||
      (!zipDirectory.has("xl/workbook.xml") &&
        !(
          options.allowBinaryWorkbookInventory &&
          zipDirectory.has("xl/workbook.bin")
        ))) &&
      throwInvalidXlsx("This archive is not a supported Excel workbook."),
    zipDirectory
  );
}
function parseZipDirectory(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  (bytes.length >= 8 &&
    view.getUint32(0, true) === 3759263696 &&
    view.getUint32(4, true) === 3776623009 &&
    throwInvalidXlsx(
      "Encrypted Office or legacy binary workbooks are unsupported. Remove the open password in the source application.",
    ),
    (bytes.length < 22 || view.getUint32(0, true) !== 67324752) &&
      throwInvalidXlsx("This file is not a supported XLSX archive."));
  let endRecordOffset = -1;
  for (
    let scanOffset = bytes.length - 22;
    scanOffset >= Math.max(0, bytes.length - 65557);
    scanOffset--
  )
    if (
      view.getUint32(scanOffset, true) === 101010256 &&
      scanOffset + 22 + view.getUint16(scanOffset + 20, true) === bytes.length
    ) {
      endRecordOffset = scanOffset;
      break;
    }
  endRecordOffset < 0 &&
    throwInvalidXlsx("The XLSX archive is incomplete or has trailing data.");
  const entryCount = view.getUint16(endRecordOffset + 10, true),
    centralDirectorySize = view.getUint32(endRecordOffset + 12, true),
    centralDirectoryOffset = view.getUint32(endRecordOffset + 16, true);
  (view.getUint16(endRecordOffset + 4, true) ||
    view.getUint16(endRecordOffset + 6, true) ||
    entryCount !== view.getUint16(endRecordOffset + 8, true) ||
    entryCount === 65535 ||
    entryCount > fileImportLimits.maxXlsxEntries ||
    centralDirectoryOffset + centralDirectorySize !== endRecordOffset) &&
    throwInvalidXlsx(
      "Multi-volume, ZIP64 or oversized XLSX archives are not supported.",
    );
  const entriesByPath = new Map(),
    seenPaths = new Set(),
    entryRanges = [],
    decodePath = (pathBytes) => {
      try {
        return new TextDecoder("utf-8", { fatal: true }).decode(pathBytes);
      } catch {
        return throwInvalidXlsx("Unsupported XLSX path encoding.");
      }
    };
  let recordOffset = centralDirectoryOffset,
    expandedTotal = 0;
  for (let entryIndex = 0; entryIndex < entryCount; entryIndex++) {
    (recordOffset + 46 > endRecordOffset ||
      view.getUint32(recordOffset, true) !== 33639248) &&
      throwInvalidXlsx("Invalid XLSX archive directory.");
    const generalFlags = view.getUint16(recordOffset + 8, true),
      compressionMethod = view.getUint16(recordOffset + 10, true),
      crc = view.getUint32(recordOffset + 16, true),
      compressedSize = view.getUint32(recordOffset + 20, true),
      size = view.getUint32(recordOffset + 24, true),
      fileNameLength = view.getUint16(recordOffset + 28, true),
      extraLength = view.getUint16(recordOffset + 30, true),
      commentLength = view.getUint16(recordOffset + 32, true),
      start = view.getUint32(recordOffset + 42, true);
    (generalFlags & 8257 &&
      throwInvalidXlsx("Encrypted XLSX archive entries are unsupported."),
      (generalFlags & -2063 ||
        ![0, 8].includes(compressionMethod) ||
        [compressedSize, size, start].includes(4294967295) ||
        view.getUint16(recordOffset + 34, true)) &&
        throwInvalidXlsx("Unsupported XLSX compression or ZIP64 metadata."),
      (recordOffset + 46 + fileNameLength + extraLength + commentLength >
        endRecordOffset ||
        start + 30 > centralDirectoryOffset ||
        view.getUint32(start, true) !== 67324752) &&
        throwInvalidXlsx("Invalid XLSX archive offsets."));
    const entryPath = decodePath(
      bytes.subarray(recordOffset + 46, recordOffset + 46 + fileNameLength),
    );
    ((!entryPath ||
      seenPaths.has(entryPath.toLowerCase()) ||
      /^[\/]/.test(entryPath) ||
      /[\\\x00-\x1f\x7f%:]/.test(entryPath) ||
      entryPath
        .split("/")
        .some(
          (segment, segmentIndex, segments) =>
            segment === ".." ||
            segment === "." ||
            (!segment && segmentIndex < segments.length - 1),
        )) &&
      throwInvalidXlsx("Unsafe or duplicate XLSX archive path."),
      seenPaths.add(entryPath.toLowerCase()));
    const localNameLength = view.getUint16(start + 26, true),
      localExtraLength = view.getUint16(start + 28, true),
      dataStart = start + 30 + localNameLength + localExtraLength,
      dataEnd = dataStart + compressedSize;
    (dataEnd > centralDirectoryOffset ||
      view.getUint16(start + 6, true) !== generalFlags ||
      view.getUint16(start + 8, true) !== compressionMethod ||
      decodePath(bytes.subarray(start + 30, start + 30 + localNameLength)) !==
        entryPath) &&
      throwInvalidXlsx(
        "XLSX local headers do not match the archive directory.",
      );
    let end = dataEnd;
    if (generalFlags & 8) {
      const descriptorOffset =
        dataEnd +
        (dataEnd + 4 <= centralDirectoryOffset &&
        view.getUint32(dataEnd, true) === 134695760
          ? 4
          : 0);
      ((descriptorOffset + 12 > centralDirectoryOffset ||
        view.getUint32(descriptorOffset, true) !== crc ||
        view.getUint32(descriptorOffset + 4, true) !== compressedSize ||
        view.getUint32(descriptorOffset + 8, true) !== size) &&
        throwInvalidXlsx("Invalid XLSX streaming data descriptor."),
        (end = descriptorOffset + 12));
    } else
      (view.getUint32(start + 14, true) !== crc ||
        view.getUint32(start + 18, true) !== compressedSize ||
        view.getUint32(start + 22, true) !== size) &&
        throwInvalidXlsx(
          "XLSX local sizes or checksum do not match the archive directory.",
        );
    (compressionMethod === 0 &&
      compressedSize !== size &&
      throwInvalidXlsx("Invalid stored XLSX entry size."),
      entryRanges.push({ start: start, end: end }),
      (expandedTotal += size),
      assertXlsxExpansionLimit(compressedSize, size, expandedTotal),
      entriesByPath.set(entryPath, { size: size, crc: crc }),
      (recordOffset += 46 + fileNameLength + extraLength + commentLength));
  }
  return (
    entryRanges.sort(
      (leftRange, rightRange) => leftRange.start - rightRange.start,
    ),
    (recordOffset !== endRecordOffset ||
      entryRanges.some(
        (range, rangeIndex) =>
          range.start !== (rangeIndex ? entryRanges[rangeIndex - 1].end : 0),
      ) ||
      entryRanges.at(-1)?.end !== centralDirectoryOffset) &&
      throwInvalidXlsx("Overlapping or unlisted XLSX archive entries."),
    entriesByPath
  );
}
const xlsxMainContentType =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const packageRelationshipsNamespace =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const officeDocumentRelationshipPrefix =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const spreadsheetMainNamespace =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const createWorkbookParseState = () => ({
  warnings: new Set(),
  hasMainType: false,
  hasOfficeDocument: false,
  strings: 0,
  maxString: -1,
  styles: 0,
  maxStyle: -1,
  emptyStringCaches: new Map(),
});
function resolveRelationshipTarget(sourcePartName, relationshipTarget) {
  let decodedTarget;
  try {
    decodedTarget = decodeURIComponent(relationshipTarget);
  } catch {
    return throwInvalidXlsx(
      `Invalid relationship target in ${sourcePartName}.`,
    );
  }
  /[\\\x00-\x1f\x7f:?#]/.test(decodedTarget) &&
    throwInvalidXlsx(`Unsafe internal relationship in ${sourcePartName}.`);
  const baseSegments =
      sourcePartName === "_rels/.rels"
        ? []
        : sourcePartName
            .replace(/_rels\/[^/]+\.rels$/, "")
            .split("/")
            .filter(Boolean),
    resolvedSegments = decodedTarget.startsWith("/") ? [] : baseSegments;
  for (const segment of decodedTarget.split("/"))
    segment === ".."
      ? (resolvedSegments.length ||
          throwInvalidXlsx(
            `Relationship escapes the XLSX package in ${sourcePartName}.`,
          ),
        resolvedSegments.pop())
      : segment && segment !== "." && resolvedSegments.push(segment);
  return resolvedSegments.join("/");
}
function parseWorksheetXml(xmlText, partName, partNames, workbookParseState) {
  const parser = new Ox.SaxesParser({ xmlns: true });
  let elementDepth = 0,
    elementCount = 0;
  const seenDeclarations = new Set(),
    openElementNames = [];
  let currentCell;
  const isWorksheetPart = /^xl\/worksheets\/[^/]+\.xml$/.test(partName);
  (parser.on("error", () => throwInvalidXlsx(`Malformed XML in ${partName}.`)),
    parser.on("doctype", () =>
      throwInvalidXlsx(
        `DTD/entity declarations are unsupported in ${partName}.`,
      ),
    ),
    parser.on("cdata", () =>
      throwInvalidXlsx(
        `CDATA content is unsupported in ${partName}. Resave as standard XLSX text encoding.`,
      ),
    ),
    parser.on("opentag", (startTag) => {
      (++elementDepth > 128 || ++elementCount > 1e6) &&
        throwInvalidXlsx(
          `XML complexity exceeds the supported budget in ${partName}.`,
        );
      const readAttribute = (attributeName) =>
          startTag.attributes[attributeName]?.value,
        parentName = openElementNames.at(-1);
      if (
        (openElementNames.push(startTag.local),
        partName === "xl/sharedStrings.xml" &&
          startTag.local === "si" &&
          workbookParseState.strings++,
        partName === "xl/styles.xml" &&
          startTag.local === "xf" &&
          parentName === "cellXfs" &&
          workbookParseState.styles++,
        isWorksheetPart && startTag.local === "c" && parentName === "row")
      ) {
        const cellType = readAttribute("t") ?? "n";
        (["n", "s", "str", "inlineStr", "b", "e"].includes(cellType) ||
          throwInvalidXlsx(`Unsupported cell type ${cellType} in ${partName}.`),
          (currentCell = {
            address: readAttribute("r") ?? "unknown cell",
            type: cellType,
            value: "",
            inValue: false,
            formula: false,
          }));
        const styleIndex = readAttribute("s");
        styleIndex !== undefined &&
          ((!/^\d+$/.test(styleIndex) ||
            !Number.isSafeInteger(Number(styleIndex))) &&
            throwInvalidXlsx(`Invalid cell style index in ${partName}.`),
          (workbookParseState.maxStyle = Math.max(
            workbookParseState.maxStyle,
            Number(styleIndex),
          )));
      }
      if (
        (currentCell &&
          parentName === "c" &&
          startTag.local === "v" &&
          ((currentCell.value = ""), (currentCell.inValue = true)),
        currentCell &&
          parentName === "c" &&
          startTag.local === "f" &&
          ((currentCell.formula = true),
          readAttribute("t") &&
            !["normal", "shared", "array"].includes(readAttribute("t")) &&
            throwInvalidXlsx(
              `Unsupported formula kind at ${currentCell.address} in ${partName}. Convert data-table formulas to reviewed ordinary formulas or values first.`,
            )),
        isWorksheetPart &&
          ["sheetProtection", "autoFilter", "extLst", "hyperlinks"].includes(
            startTag.local,
          ) &&
          workbookParseState.warnings.add(
            `Worksheet ${startTag.local} metadata was detected. Protection, filter criteria, interactive links and unsupported extensions are not preserved as editable Excel features.`,
          ),
        elementDepth === 1 &&
          ((partName === "[Content_Types].xml" &&
            (startTag.local !== "Types" ||
              startTag.uri !==
                "http://schemas.openxmlformats.org/package/2006/content-types")) ||
            (partName.endsWith(".rels") &&
              (startTag.local !== "Relationships" ||
                startTag.uri !== packageRelationshipsNamespace)) ||
            (partName === "xl/workbook.xml" &&
              (startTag.local !== "workbook" ||
                startTag.uri !== spreadsheetMainNamespace))) &&
          throwInvalidXlsx(
            `Unsupported XLSX package type or Strict XML namespace in ${partName}. Save as standard Excel Workbook (.xlsx).`,
          ),
        partName === "[Content_Types].xml" && startTag.local === "Override")
      ) {
        const overridePartName = readAttribute("PartName"),
          overrideContentType = readAttribute("ContentType");
        ((!overridePartName ||
          !overrideContentType ||
          seenDeclarations.has(overridePartName)) &&
          throwInvalidXlsx(
            "Invalid or duplicate XLSX content type declaration.",
          ),
          seenDeclarations.add(overridePartName),
          /model\+data|dataModel/i.test(overrideContentType) &&
            workbookParseState.warnings.add(
              "Imported Excel Data Model and DAX content is not executed or translated; model caches are historical and require independent review. Exact bytes remain only in the retained original.",
            ),
          /macroEnabled|vbaProject|activeX|oleObject/i.test(
            overrideContentType,
          ) &&
            throwInvalidXlsx(
              "Macro-enabled or executable workbook content is unsupported; renaming a file does not convert its format.",
            ),
          overridePartName === "/xl/workbook.xml" &&
            (overrideContentType !== xlsxMainContentType &&
              throwInvalidXlsx(
                "Unsupported workbook content type. Templates and alternate formats must be saved as XLSX.",
              ),
            (workbookParseState.hasMainType = true)),
          partNames.has(
            resolveRelationshipTarget("_rels/.rels", overridePartName),
          ) || throwInvalidXlsx(`Missing XLSX part ${overridePartName}.`));
      }
      if (partName.endsWith(".rels") && startTag.local === "Relationship") {
        const relationshipId = readAttribute("Id"),
          targetAttribute = readAttribute("Target"),
          relationshipType = readAttribute("Type"),
          targetMode = readAttribute("TargetMode");
        if (
          ((!relationshipId ||
            !targetAttribute ||
            !relationshipType ||
            seenDeclarations.has(relationshipId) ||
            (targetMode && !["External", "Internal"].includes(targetMode))) &&
            throwInvalidXlsx(
              `Invalid or duplicate relationship in ${partName}.`,
            ),
          seenDeclarations.add(relationshipId),
          /vbaProject|activeX|oleObject|attachedTemplate/i.test(
            relationshipType,
          ) &&
            throwInvalidXlsx(
              "Executable or embedded-object relationships are unsupported.",
            ),
          targetMode === "External")
        )
          (relationshipType !==
            `${officeDocumentRelationshipPrefix}hyperlink` &&
            relationshipType !==
              `${officeDocumentRelationshipPrefix}externalLinkPath` &&
            throwInvalidXlsx(
              `Unsupported external relationship in ${partName}.`,
            ),
            workbookParseState.warnings.add(
              "External hyperlinks/workbook links were found. No external target is fetched; interactive links and refresh connections are omitted.",
            ));
        else {
          const relationshipTarget = resolveRelationshipTarget(
            partName,
            targetAttribute,
          );
          (partNames.has(relationshipTarget) ||
            throwInvalidXlsx(
              `Missing relationship target ${relationshipTarget} in ${partName}.`,
            ),
            partName === "_rels/.rels" &&
              relationshipType ===
                `${officeDocumentRelationshipPrefix}officeDocument` &&
              ((workbookParseState.hasOfficeDocument ||
                relationshipTarget !== "xl/workbook.xml") &&
                throwInvalidXlsx(
                  "Unsupported or duplicate workbook relationship.",
                ),
              (workbookParseState.hasOfficeDocument = true)));
        }
      }
    }),
    parser.on("text", (textChunk) => {
      currentCell?.inValue && (currentCell.value += textChunk);
    }),
    parser.on("closetag", (endTag) => {
      if (currentCell && endTag.local === "v") {
        const value = currentCell.value;
        if (
          (currentCell.type === "n" &&
            value !== "" &&
            (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value) ||
              !Number.isFinite(Number(value))) &&
            throwInvalidXlsx(
              `Invalid numeric value at ${currentCell.address} in ${partName}.`,
            ),
          currentCell.type === "b" &&
            !["0", "1"].includes(value) &&
            throwInvalidXlsx(
              `Unsupported boolean value at ${currentCell.address} in ${partName}; save numeric Boolean encoding in a standard XLSX copy.`,
            ),
          currentCell.type === "s" &&
            ((!/^\d+$/.test(value) || !Number.isSafeInteger(Number(value))) &&
              throwInvalidXlsx(
                `Invalid shared-string index at ${currentCell.address} in ${partName}.`,
              ),
            (workbookParseState.maxString = Math.max(
              workbookParseState.maxString,
              Number(value),
            ))),
          currentCell.type === "e" &&
            !/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA|CONNECT!|BLOCKED!|FIELD!|UNKNOWN!|BUSY!)$/.test(
              value,
            ) &&
            throwInvalidXlsx(
              `Unsupported error value at ${currentCell.address} in ${partName}.`,
            ),
          currentCell.type === "e" &&
            !currentCell.formula &&
            workbookParseState.warnings.add(
              "Source constant error cells are flattened to displayed text. Error propagation from those cells is not preserved; review dependent calculations.",
            ),
          currentCell.type === "str" && currentCell.formula && value === "")
        ) {
          const emptyStringAddresses =
            workbookParseState.emptyStringCaches.get(partName) ?? new Set();
          (emptyStringAddresses.add(currentCell.address),
            workbookParseState.emptyStringCaches.set(
              partName,
              emptyStringAddresses,
            ));
        }
        currentCell.inValue = false;
      }
      (endTag.local === "c" && (currentCell = undefined),
        openElementNames.pop(),
        elementDepth--);
    }),
    parser.write(xmlText).close());
}
function validateWorkbookPackage(workbookParseState, partPaths) {
  ((!workbookParseState.hasMainType || !workbookParseState.hasOfficeDocument) &&
    throwInvalidXlsx(
      "Missing standard XLSX workbook content type or relationship.",
    ),
    workbookParseState.maxString >= workbookParseState.strings &&
      throwInvalidXlsx("A worksheet references a missing shared string."),
    workbookParseState.maxStyle >= workbookParseState.styles &&
      throwInvalidXlsx("A worksheet references a missing cell style."));
  const partWarningRules = [
    [
      /^xl\/(?:model|dataModel)\//i,
      "Excel Data Model / Power Pivot content is retained only in the original source and is not executed or translated. Relationships, DAX measures and model caches are unavailable; worksheet results are historical snapshots. Rebuild an explicit local Data Model from reviewed worksheet tables before relying on recalculated results.",
    ],
    [/^xl\/charts\//i, "Charts are not imported as editable native charts."],
    [
      /^xl\/(?:pivotTables|pivotCache)\//i,
      "Pivot definitions and caches are omitted; supported worksheet result cells remain.",
    ],
    [
      /^xl\/(?:queryTables|externalLinks)\/|^xl\/connections\.xml$/i,
      "Queries, external workbook links and connections are omitted and never refreshed.",
    ],
    [
      /^customXml\/|^docProps\/custom\.xml$/i,
      "Custom XML and custom document properties are not editable; the part inventory identifies passive metadata preserved unchanged and content retained only in the original.",
    ],
    [
      /^xl\/(?:threadedComments|persons|slicers|slicerCaches|timelines)\//i,
      "Threaded discussions, people, slicers and timelines are omitted.",
    ],
  ];
  for (const partPath of partPaths)
    for (const [warningPattern, warningMessage] of partWarningRules)
      warningPattern.test(partPath) &&
        workbookParseState.warnings.add(warningMessage);
  return [...workbookParseState.warnings];
}
const templateConversionRoutes = Object.freeze(
  alternateFormatExtensions.map((alternateFormatExtension) =>
    Object.freeze({
      format: alternateFormatExtension,
      read:
        alternateFormatExtension === "xltx"
          ? "Native template adapter: supported XLSX subset"
          : "External conversion to XLSX, then reviewed import",
      write:
        alternateFormatExtension === "xltx"
          ? "Derived XLTX template with loss review"
          : "Export reviewed XLSX, then Save As in the source application",
      preserve:
        "Exact original bytes in native JSON; derived output is not lossless",
      formulas:
        alternateFormatExtension === "xltx"
          ? "Supported formulas and cache comparison"
          : "Verify formulas and caches after external conversion",
      dates: "Verify date system, serials and displayed dates",
      styles: "Supported formatting only; compare with original",
      specialParts:
        "Original only unless supported by the XLSX codec; never execute macros or objects",
    }),
  ),
);
const macroEnabledContentTypes = Object.freeze({
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
  xltm: "application/vnd.ms-excel.template.macroEnabled.main+xml",
  xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
});
function inspectAlternateWorkbookArchive(bytes) {
  if (bytes.length > workbookLimits.maxFileBytes)
    throw Error("Files must be 10 MB or smaller.");
  const zipDirectory = parseZipDirectory(bytes),
    entriesByName = Object.create(null),
    seenEntryNames = new Set();
  let totalExpandedBytes = 0,
    archiveError;
  const unzipper = new Unzip((fileEntry) => {
    if (
      !zipDirectory.has(fileEntry.name) ||
      seenEntryNames.has(fileEntry.name)
    ) {
      archiveError = Error("Alternate workbook archive directory mismatch.");
      return;
    }
    seenEntryNames.add(fileEntry.name);
    let entryByteCount = 0,
      runningCrc = -1;
    const dataChunks = [];
    ((fileEntry.ondata = (chunkError, chunk, isFinalChunk) => {
      if (!archiveError) {
        if (chunkError) {
          archiveError = Error("Corrupt alternate workbook archive.");
          return;
        }
        if (
          ((entryByteCount += chunk.length),
          (totalExpandedBytes += chunk.length),
          entryByteCount > zipDirectory.get(fileEntry.name).size ||
            totalExpandedBytes > 50 * 1024 * 1024)
        ) {
          ((archiveError = Error(
            "Alternate workbook archive exceeds its expansion budget.",
          )),
            fileEntry.terminate());
          return;
        }
        if (
          ((runningCrc = computeCrc32(runningCrc, chunk)),
          dataChunks.push(chunk),
          isFinalChunk)
        ) {
          if (
            entryByteCount !== zipDirectory.get(fileEntry.name).size ||
            (runningCrc ^ -1) >>> 0 !== zipDirectory.get(fileEntry.name).crc
          ) {
            archiveError = Error(
              "Alternate workbook archive checksum mismatch.",
            );
            return;
          }
          const entryData = new Uint8Array(entryByteCount);
          let copyOffset = 0;
          for (const dataChunk of dataChunks)
            (entryData.set(dataChunk, copyOffset),
              (copyOffset += dataChunk.length));
          entriesByName[fileEntry.name] = entryData;
        }
      }
    }),
      fileEntry.start());
  });
  unzipper.register(UnzipInflate);
  try {
    unzipper.push(bytes, true);
  } catch {
    throw Error("Unable to inspect corrupt alternate workbook archive.");
  }
  if (archiveError) throw archiveError;
  if (Object.keys(entriesByName).length !== zipDirectory.size)
    throw Error("Alternate workbook archive is incomplete.");
  return entriesByName;
}
function parseAlternateWorkbookXml(xmlBytes, onElement) {
  const xmlText = new TextDecoder("utf-8", { fatal: true }).decode(xmlBytes);
  if (
    xmlText.includes("\0") ||
    /<\?xml[^>]*encoding\s*=\s*["'](?!utf-8|utf8)[^"']+/i.test(xmlText)
  )
    throw Error("Alternate workbook XML must use UTF-8.");
  const parser = new Ox.SaxesParser({ xmlns: true });
  let elementDepth = 0,
    elementCount = 0;
  (parser.on("doctype", () => {
    throw Error("Alternate workbook XML declarations are unsupported.");
  }),
    parser.on("opentag", (element) => {
      if (++elementDepth > 128 || ++elementCount > 1e6)
        throw Error("Alternate workbook XML exceeds its inspection budget.");
      onElement(element);
    }),
    parser.on("closetag", () => {
      elementDepth--;
    }),
    parser.write(xmlText).close());
}
function readOpenDocumentPackage(packageEntries) {
  const readAttributes = (element) =>
    Object.fromEntries(
      Object.values(element.attributes).map((attribute) => [
        attribute.local,
        attribute.value,
      ]),
    );
  if (packageEntries.mimetype) {
    if (
      new TextDecoder().decode(packageEntries.mimetype) !==
        "application/vnd.oasis.opendocument.spreadsheet" ||
      !packageEntries["META-INF/manifest.xml"] ||
      !packageEntries["content.xml"] ||
      packageEntries["[Content_Types].xml"]
    )
      throw Error("Unsupported or mismatched OpenDocument workbook identity.");
    let hasRootManifestEntry = false,
      hasSpreadsheetBody = false;
    if (
      (parseAlternateWorkbookXml(
        packageEntries["META-INF/manifest.xml"],
        (manifestElement) => {
          const manifestAttributes = readAttributes(manifestElement);
          if (
            manifestElement.uri ===
            "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
          ) {
            if (manifestElement.local === "encryption-data")
              throw Error(
                "Encrypted OpenDocument workbooks require an unencrypted XLSX conversion.",
              );
            manifestElement.local === "file-entry" &&
              manifestAttributes["full-path"] === "/" &&
              manifestAttributes["media-type"] ===
                "application/vnd.oasis.opendocument.spreadsheet" &&
              (hasRootManifestEntry = true);
          }
        },
      ),
      parseAlternateWorkbookXml(
        packageEntries["content.xml"],
        (contentElement) => {
          contentElement.local === "spreadsheet" &&
            contentElement.uri ===
              "urn:oasis:names:tc:opendocument:xmlns:office:1.0" &&
            (hasSpreadsheetBody = true);
        },
      ),
      !hasRootManifestEntry || !hasSpreadsheetBody)
    )
      throw Error("Unsupported or mismatched OpenDocument workbook identity.");
    return "ods";
  }
  if (!packageEntries["[Content_Types].xml"] || !packageEntries["_rels/.rels"])
    throw Error("Unsupported alternate workbook package identity.");
  let workbookContentType,
    workbookPartPath = "",
    mainRelationshipCount = 0;
  const seenPartNames = new Set();
  parseAlternateWorkbookXml(
    packageEntries["[Content_Types].xml"],
    (contentTypeElement) => {
      if (
        contentTypeElement.uri !==
        "http://schemas.openxmlformats.org/package/2006/content-types"
      )
        throw Error("Unsupported alternate workbook content-type namespace.");
      const contentTypeAttributes = readAttributes(contentTypeElement);
      if (contentTypeElement.local === "Override") {
        if (
          !contentTypeAttributes.PartName ||
          seenPartNames.has(contentTypeAttributes.PartName)
        )
          throw Error("Duplicate or missing alternate workbook content type.");
        if (
          (seenPartNames.add(contentTypeAttributes.PartName),
          contentTypeAttributes.PartName === "/xl/workbook.xml" ||
            contentTypeAttributes.PartName === "/xl/workbook.bin")
        ) {
          if (workbookContentType)
            throw Error("Ambiguous alternate workbook identity.");
          ((workbookContentType = contentTypeAttributes.ContentType),
            (workbookPartPath = contentTypeAttributes.PartName.slice(1)));
        }
      }
    },
  );
  const workbookExtension = Object.entries(macroEnabledContentTypes).find(
    ([, knownContentType]) => knownContentType === workbookContentType,
  )?.[0];
  if (
    !workbookExtension ||
    !packageEntries[workbookPartPath] ||
    workbookPartPath !==
      (workbookExtension === "xlsb" ? "xl/workbook.bin" : "xl/workbook.xml")
  )
    throw Error("Unsupported alternate workbook content type.");
  if (
    (parseAlternateWorkbookXml(
      packageEntries["_rels/.rels"],
      (relationshipElement) => {
        const relationshipAttributes = readAttributes(relationshipElement);
        if (
          relationshipElement.uri ===
            "http://schemas.openxmlformats.org/package/2006/relationships" &&
          relationshipElement.local === "Relationship" &&
          relationshipAttributes.Type ===
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        ) {
          if (
            relationshipAttributes.TargetMode ||
            resolveRelationshipTarget(
              "_rels/.rels",
              relationshipAttributes.Target,
            ) !== workbookPartPath
          )
            throw Error("Invalid alternate workbook main relationship.");
          mainRelationshipCount++;
        }
      },
    ),
    workbookExtension === "xlsb")
  ) {
    const workbookBytes = packageEntries[workbookPartPath];
    if (
      workbookBytes.length < 3 ||
      workbookBytes[0] !== 131 ||
      workbookBytes[1] !== 1 ||
      workbookBytes[2] !== 0
    )
      throw Error(
        "Invalid binary workbook beginning record. Convert the original in its source application.",
      );
  } else {
    let isFirstElement = true;
    parseAlternateWorkbookXml(
      packageEntries[workbookPartPath],
      (workbookElement) => {
        if (
          isFirstElement &&
          ((isFirstElement = false),
          workbookElement.local !== "workbook" ||
            ![
              "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
              "http://purl.oclc.org/ooxml/spreadsheetml/main",
            ].includes(workbookElement.uri))
        )
          throw Error("Invalid alternate workbook XML identity.");
      },
    );
  }
  if (mainRelationshipCount !== 1)
    throw Error("Invalid alternate workbook main relationship.");
  return workbookExtension;
}
function convertTemplateToXlsx(bytes, sourceFormat, targetFormat) {
  const archiveEntries = inspectAlternateWorkbookArchive(bytes);
  if (readOpenDocumentPackage(archiveEntries) !== sourceFormat)
    throw Error(
      "The workbook content does not match the template conversion route.",
    );
  if (
    ((archiveEntries["[Content_Types].xml"] = strToU8(
      strFromU8(archiveEntries["[Content_Types].xml"]).replace(
        macroEnabledContentTypes[sourceFormat],
        macroEnabledContentTypes[targetFormat],
      ),
    )),
    readOpenDocumentPackage(archiveEntries) !== targetFormat)
  )
    throw Error("The template content type could not be converted safely.");
  return zipSync(archiveEntries);
}

/* From 08-xlsx-import-export-and-preservation.js */
const z2 =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const customPropertiesContentType =
  "application/vnd.openxmlformats-officedocument.custom-properties+xml";
const customXmlPropertiesContentType =
  "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
const Dle = { scope: "workbook" };
function parseXmlPartStrict(xmlBytes, onOpenTag, onCloseTag, onText) {
  const parser = new Ox.SaxesParser({ xmlns: true });
  (parser.on("opentag", onOpenTag),
    onCloseTag && parser.on("closetag", onCloseTag),
    onText && parser.on("text", onText),
    parser.on("doctype", () => {
      throw Error("DTD is not permitted in preservation metadata.");
    }),
    parser.write(strFromU8(xmlBytes)).close());
}
function ownerPartForRelsPath(relsPath) {
  return relsPath === "_rels/.rels"
    ? ""
    : relsPath.replace(/(^|\/)\_rels\/([^/]+)\.rels$/, "$1$2");
}
function relationshipsPathFor(partPath) {
  if (!partPath) return "_rels/.rels";
  const lastSlashIndex = partPath.lastIndexOf("/");
  return `${partPath.slice(0, lastSlashIndex + 1)}_rels/${partPath.slice(lastSlashIndex + 1)}.rels`;
}
function readPackageRelationships(packageEntries) {
  const types = new Map(),
    defaultContentTypes = new Map(),
    relationships = [],
    sheets = new Map();
  parseXmlPartStrict(
    packageEntries["[Content_Types].xml"],
    (contentTypesElement) => {
      (contentTypesElement.local === "Override" &&
        types.set(
          (contentTypesElement.attributes.PartName?.value ?? "").slice(1),
          contentTypesElement.attributes.ContentType?.value ?? "",
        ),
        contentTypesElement.local === "Default" &&
          defaultContentTypes.set(
            contentTypesElement.attributes.Extension?.value ?? "",
            contentTypesElement.attributes.ContentType?.value ?? "",
          ));
    },
  );
  for (const [rels, relsBytes] of Object.entries(packageEntries))
    rels.endsWith(".rels") &&
      parseXmlPartStrict(relsBytes, (relationshipElement) => {
        if (relationshipElement.local !== "Relationship") return;
        const external =
            relationshipElement.attributes.TargetMode?.value === "External",
          targetValue = relationshipElement.attributes.Target?.value ?? "";
        relationships.push({
          owner: ownerPartForRelsPath(rels),
          rels: rels,
          id: relationshipElement.attributes.Id?.value ?? "",
          type: relationshipElement.attributes.Type?.value ?? "",
          external: external,
          target: external
            ? targetValue
            : resolveRelationshipTarget(rels, targetValue),
        });
      });
  return (
    parseXmlPartStrict(packageEntries["xl/workbook.xml"], (workbookElement) => {
      if (workbookElement.local === "sheet") {
        const relationship = relationships.find(
          (relationship) =>
            relationship.owner === "xl/workbook.xml" &&
            relationship.id ===
              Object.values(workbookElement.attributes).find(
                (attribute) =>
                  attribute.local === "id" && attribute.uri === z2.slice(0, -1),
              )?.value,
        );
        relationship &&
          sheets.set(
            relationship.target,
            workbookElement.attributes.name?.value ?? "Unknown worksheet",
          );
      }
    }),
    {
      relationships: relationships,
      sheets: sheets,
      contentType: (partPath) =>
        types.get(partPath) ??
        defaultContentTypes.get(partPath.split(".").at(-1)) ??
        "undeclared",
      types: types,
    }
  );
}
function classifyPreservedPart(path, contentType) {
  return path.endsWith("/")
    ? "directory"
    : path === "[Content_Types].xml" || path.endsWith(".rels")
      ? "package relationships"
      : /vba|macrosheet|activeX|embedding/i.test(path + contentType)
        ? "active content"
        : /^customXml\//.test(path)
          ? "custom XML"
          : path === "docProps/custom.xml"
            ? "custom properties"
            : /^xl\/charts\//.test(path)
              ? "charts"
              : /^xl\/(?:model|dataModel)\//i.test(path) ||
                  /model\+data|dataModel/i.test(contentType)
                ? "Excel Data Model"
                : /^xl\/(pivotTables|pivotCache)\//.test(path)
                  ? "pivots"
                  : /connections|queryTables|externalLinks/.test(path)
                    ? "connections and external links"
                    : /slicers|slicerCaches|timelines/.test(path)
                      ? "slicers and timelines"
                      : /threadedComments|persons/.test(path)
                        ? "threaded comments"
                        : /^_xmlsignatures\//.test(path)
                          ? "signatures"
                          : /^xl\/worksheets\/[^/]+\.xml$/.test(path)
                            ? "worksheet"
                            : path === "xl/workbook.xml"
                              ? "workbook and names"
                              : /^xl\/tables\//.test(path)
                                ? "tables"
                                : /^xl\/drawings\//.test(path)
                                  ? "drawings"
                                  : /^xl\/media\//.test(path)
                                    ? "media"
                                    : /^xl\/comments/.test(path)
                                      ? "notes"
                                      : path === "xl/sharedStrings.xml"
                                        ? "shared strings"
                                        : path === "xl/styles.xml" ||
                                            /^xl\/theme\//.test(path)
                                          ? "formatting"
                                          : /^docProps\//.test(path)
                                            ? "document properties"
                                            : "unknown part";
}
function F$(packageEntries, sha256) {
  const packageRelationships = readPackageRelationships(packageEntries),
    relationshipsByOwner = new Map(),
    relationshipsByTarget = new Map();
  for (const relationship of packageRelationships.relationships)
    (relationshipsByOwner.set(relationship.owner, [
      ...(relationshipsByOwner.get(relationship.owner) ?? []),
      relationship,
    ]),
      relationship.external ||
        relationshipsByTarget.set(relationship.target, [
          ...(relationshipsByTarget.get(relationship.target) ?? []),
          relationship,
        ]));
  const sheetsByPart = new Map();
  for (const [sheetPartPath, sheetInfo] of packageRelationships.sheets) {
    const pendingPartPaths = [sheetPartPath],
      visitedParts = new Set();
    for (; pendingPartPaths.length;) {
      const partPath = pendingPartPaths.pop();
      if (visitedParts.has(partPath)) continue;
      visitedParts.add(partPath);
      const partSheets = sheetsByPart.get(partPath) ?? new Set();
      (partSheets.add(sheetInfo), sheetsByPart.set(partPath, partSheets));
      for (const childRelationship of relationshipsByOwner.get(partPath) ?? [])
        childRelationship.external ||
          pendingPartPaths.push(childRelationship.target);
    }
  }
  const opaqueParts = new Set(),
    retentionReasons = new Map(),
    customXmlPaths = Object.keys(packageEntries).filter(
      (entryPath) =>
        entryPath.startsWith("customXml/") && !entryPath.endsWith("/"),
    );
  if (customXmlPaths.length) {
    const customXmlPathSet = new Set(customXmlPaths),
      customXmlRelationships = packageRelationships.relationships.filter(
        (relationship) =>
          customXmlPathSet.has(relationship.owner) ||
          (!relationship.external && customXmlPathSet.has(relationship.target)),
      );
    let isPassiveCustomXml = !Object.keys(packageEntries).some((partName) =>
      /xmlMaps|singleXmlCells/i.test(partName),
    );
    ((isPassiveCustomXml &&= customXmlPaths.every(
      (customXmlPart) =>
        customXmlPart.endsWith(".rels") ||
        (/\.xml$/.test(customXmlPart) &&
          [
            "application/xml",
            "text/xml",
            customXmlPropertiesContentType,
          ].includes(packageRelationships.contentType(customXmlPart))),
    )),
      (isPassiveCustomXml &&= customXmlRelationships.every(
        (customXmlLink) =>
          !customXmlLink.external &&
          ((customXmlLink.owner === "xl/workbook.xml" &&
            customXmlLink.type === z2 + "customXml" &&
            customXmlPathSet.has(customXmlLink.target)) ||
            (customXmlPathSet.has(customXmlLink.owner) &&
              customXmlPathSet.has(customXmlLink.target) &&
              customXmlLink.type === z2 + "customXmlProps")),
      )),
      (isPassiveCustomXml &&= customXmlRelationships.every((xmlRelationship) =>
        xmlRelationship.owner === "xl/workbook.xml"
          ? ["application/xml", "text/xml"].includes(
              packageRelationships.contentType(xmlRelationship.target),
            )
          : packageRelationships.contentType(xmlRelationship.owner) !==
              customXmlPropertiesContentType &&
            packageRelationships.contentType(xmlRelationship.target) ===
              customXmlPropertiesContentType,
      )));
    const workbookRelationships = customXmlRelationships.filter(
        (workbookRelationship) =>
          workbookRelationship.owner === "xl/workbook.xml",
      ),
      reachableParts = new Set(),
      pendingTargets = workbookRelationships.map((event) => event.target);
    for (; pendingTargets.length;) {
      const targetPath = pendingTargets.pop();
      if (!reachableParts.has(targetPath)) {
        (reachableParts.add(targetPath),
          packageEntries[relationshipsPathFor(targetPath)] &&
            reachableParts.add(relationshipsPathFor(targetPath)));
        for (const outgoingRelationship of relationshipsByOwner.get(
          targetPath,
        ) ?? [])
          outgoingRelationship.external ||
            pendingTargets.push(outgoingRelationship.target);
      }
    }
    isPassiveCustomXml &&=
      workbookRelationships.length > 0 &&
      customXmlPaths.every((xmlPartPath) => reachableParts.has(xmlPartPath));
    for (const customXmlPath of customXmlPaths)
      isPassiveCustomXml
        ? opaqueParts.add(customXmlPath)
        : retentionReasons.set(
            customXmlPath,
            "Custom XML has bindings, unrecognized types, orphaned parts or relationships outside the passive metadata island. Retained only in the original.",
          );
  }
  const customPropsPath = "docProps/custom.xml";
  if (packageEntries[customPropsPath]) {
    const customPropsRelationships =
      relationshipsByTarget.get(customPropsPath) ?? [];
    let isPassiveCustomProps =
      packageRelationships.contentType(customPropsPath) ===
        customPropertiesContentType &&
      customPropsRelationships.length === 1 &&
      customPropsRelationships[0].owner === "" &&
      customPropsRelationships[0].type === z2 + "custom-properties" &&
      !relationshipsByOwner.get(customPropsPath)?.length;
    (parseXmlPartStrict(packageEntries[customPropsPath], (customPropsTag) => {
      (customPropsTag.local === "property" &&
        customPropsTag.attributes.linkTarget &&
        (isPassiveCustomProps = false),
        (customPropsTag.uri ===
          "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" &&
          ["Properties", "property"].includes(customPropsTag.local)) ||
          (customPropsTag.uri ===
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" &&
            [
              "lpwstr",
              "lpstr",
              "bstr",
              "i1",
              "i2",
              "i4",
              "i8",
              "int",
              "ui1",
              "ui2",
              "ui4",
              "ui8",
              "uint",
              "r4",
              "r8",
              "decimal",
              "bool",
              "date",
              "filetime",
              "empty",
              "null",
            ].includes(customPropsTag.local)) ||
          (isPassiveCustomProps = false));
    }),
      isPassiveCustomProps
        ? opaqueParts.add(customPropsPath)
        : retentionReasons.set(
            customPropsPath,
            "Custom properties contain links, unsupported value types or relationships; retained only in the original.",
          ));
  }
  const richTextStringIndexes = new Set();
  if (packageEntries["xl/sharedStrings.xml"]) {
    let currentStringIndex = -1;
    parseXmlPartStrict(packageEntries["xl/sharedStrings.xml"], (xmlTag) => {
      (xmlTag.local === "si" && currentStringIndex++,
        xmlTag.local === "r" && richTextStringIndexes.add(currentStringIndex));
    });
  }
  let featureCount = 0;
  const parts = Object.entries(packageEntries)
    .sort(([leftPath], [rightPath]) => leftPath.localeCompare(rightPath))
    .map(([path, partBytes]) => {
      const family = classifyPreservedPart(
          path,
          packageRelationships.contentType(path),
        ),
        locations = [
          ...(sheetsByPart.get(
            path.endsWith(".rels") ? ownerPartForRelsPath(path) : path,
          ) ?? []),
        ].map((sheet) => ({
          sheet: sheet,
          scope: "sheet",
        }));
      locations.length ||
        locations.push(
          [
            "charts",
            "pivots",
            "drawings",
            "notes",
            "unknown part",
            "media",
          ].includes(family)
            ? {
                scope: "unresolved",
                detail:
                  "No worksheet relationship identifies the affected cells.",
              }
            : Dle,
        );
      let disposition = "dropped",
        reason =
          "No safe editable mapping. Exact source bytes remain in the original file.";
      (family === "Excel Data Model" &&
        (reason =
          "Imported Excel Data Model, relationships and DAX are not executed or translated. Worksheet results are historical snapshots. Exact model bytes remain only in the original source; rebuild reviewed local typed tables and relationships to calculate new results."),
        ["worksheet", "workbook and names", "tables", "notes"].includes(
          family,
        ) &&
          ((disposition = "editable"),
          (reason =
            "Supported content is editable and regenerated; inspect the feature details and conversion notices for losses. XML bytes are not preserved.")),
        [
          "package relationships",
          "shared strings",
          "formatting",
          "document properties",
          "drawings",
        ].includes(family) &&
          ((disposition = "flattened"),
          (reason =
            "Supported content is regenerated or flattened. Unsupported settings and original XML are retained only in the source.")),
        family === "media" &&
          /^image\/(png|jpeg)$/.test(packageRelationships.contentType(path)) &&
          ((disposition = "flattened"),
          (reason =
            "Supported worksheet pictures are re-encoded with their supported placement; orphaned images and chart assets are not carried over.")),
        opaqueParts.has(path) &&
          ((disposition = "preserved-opaque"),
          (reason =
            "Passive metadata snapshot copied byte-for-byte, without interpreting or refreshing its contents. Its original meaning may be stale after edits.")),
        retentionReasons.has(path) && (reason = retentionReasons.get(path)),
        family === "signatures" &&
          (reason =
            "Signatures cannot authenticate an edited workbook; retained only in the original file."));
      const features = [],
        addFeature = (feature, disposition, range, detail, reasonOverride) => {
          if (++featureCount > xlsxInspectionLimits.maxDetails)
            throw Error(
              "XLSX compatibility details exceed the 100,000-entry budget. Split the source workbook before importing.",
            );
          features.push({
            feature: feature,
            disposition: disposition,
            location: {
              ...locations[0],
              ...(range ? { range: range, scope: "cell" } : {}),
              ...(detail ? { detail: detail } : {}),
            },
            reason:
              reasonOverride ??
              (disposition === "editable"
                ? "Supported subset is editable; review conversion notices for unsupported attributes."
                : "This source feature is not preserved as an editable feature in the derived workbook."),
          });
        };
      if (
        family === "worksheet" ||
        family === "workbook and names" ||
        family === "tables" ||
        family === "drawings" ||
        family === "pivots" ||
        family === "charts"
      ) {
        let cellRef = "",
          cellType = "",
          definedName,
          formulaText = "",
          isInFormula = false,
          isInValue = false,
          valueText = "";
        parseXmlPartStrict(
          partBytes,
          (openTag) => {
            const readAttribute = (attributeName) =>
              openTag.attributes[attributeName]?.value;
            (openTag.local === "c" &&
              ((cellRef = readAttribute("r") ?? ""),
              (cellType = readAttribute("t") ?? ""),
              cellType === "e" &&
                addFeature("error cell", "flattened", cellRef)),
              openTag.local === "v" && ((isInValue = true), (valueText = "")),
              openTag.local === "f" &&
                ((isInFormula = true), (formulaText = "")),
              openTag.local === "definedName" &&
                ((definedName = readAttribute("name")),
                addFeature("defined name", "editable", undefined, definedName)),
              openTag.local === "dataValidation" &&
                addFeature(
                  "data validation",
                  "editable",
                  readAttribute("sqref"),
                  readAttribute("type"),
                ),
              openTag.local === "conditionalFormatting" &&
                addFeature(
                  "conditional formatting",
                  "editable",
                  readAttribute("sqref"),
                ),
              openTag.local === "mergeCell" &&
                addFeature("merged cells", "editable", readAttribute("ref")),
              openTag.local === "table" &&
                addFeature(
                  "structured table",
                  "editable",
                  readAttribute("ref"),
                  readAttribute("name"),
                ),
              openTag.local === "location" &&
                addFeature("pivot output", "dropped", readAttribute("ref")),
              openTag.local === "worksheetSource" &&
                addFeature(
                  "pivot source",
                  "dropped",
                  readAttribute("ref"),
                  readAttribute("sheet"),
                ),
              openTag.local === "r" &&
                cellRef &&
                cellType === "inlineStr" &&
                addFeature("rich text", "flattened", cellRef),
              [
                "sheetProtection",
                "workbookProtection",
                "autoFilter",
                "extLst",
                "hyperlink",
                "legacyDrawingHF",
                "webPublishItems",
                "controls",
                "oleObjects",
                "pivotTable",
                "pivotTablePart",
                "slicer",
                "timeline",
                "sp",
                "graphicFrame",
              ].includes(openTag.local) &&
                addFeature(
                  openTag.local,
                  "dropped",
                  readAttribute("ref"),
                  readAttribute("name"),
                ));
          },
          (closeTag) => {
            (closeTag.local === "v" &&
              ((isInValue = false),
              cellRef &&
                cellType === "s" &&
                richTextStringIndexes.has(Number(valueText)) &&
                addFeature("rich text", "flattened", cellRef)),
              closeTag.local === "c" && (cellRef = ""),
              closeTag.local === "f" &&
                ((isInFormula = false),
                family === "charts"
                  ? addFeature(
                      "chart source reference",
                      "dropped",
                      undefined,
                      formulaText.slice(0, 1024),
                    )
                  : /(?:\[|https?:|file:|WEBSERVICE|HYPERLINK|RTD|DDE|CALL\s*\(|EXEC\s*\()/i.test(
                      formulaText,
                    ) &&
                    addFeature(
                      "external/action formula",
                      "flattened",
                      cellRef,
                      undefined,
                      "External/action references may be neutralized to literal text; the formula cache report provides the actual conversion outcome.",
                    )));
          },
          (textChunk) => {
            (isInFormula && (formulaText += textChunk),
              isInValue && (valueText += textChunk));
          },
        );
      }
      for (const partRelationship of relationshipsByOwner.get(path) ?? [])
        partRelationship.external &&
          addFeature(
            "external relationship",
            "dropped",
            undefined,
            partRelationship.type.split("/").at(-1),
            "External target is never fetched or carried into the derived copy; its URI remains only in the original.",
          );
      return {
        path: path,
        contentType: path.endsWith("/")
          ? "directory"
          : packageRelationships.contentType(path),
        bytes: partBytes.length,
        sha256: sha256HexDigest(partBytes),
        family: family,
        disposition: disposition,
        reason: reason,
        locations: locations,
        features: features,
        retainedInOriginal: true,
      };
    });
  return summarizePreservedParts({
    schemaVersion: 1,
    policy: "passive-metadata-v1",
    complete: true,
    phase: "eligible-for-export",
    sourceSha256: sha256,
    parts: parts,
    counts: { editable: 0, "preserved-opaque": 0, flattened: 0, dropped: 0 },
    opaqueBytes: 0,
    note: "All package entries are inventoried. Locations come from source relationships and XML; unresolved locations are explicit. Dispositions describe the derived copy, while all parts remain byte-exact in the retained original. Editable parts can also contain flattened or dropped features. No source part is executed or fetched.",
  });
}
function summarizePreservedParts(summary) {
  ((summary.counts = {
    editable: 0,
    "preserved-opaque": 0,
    flattened: 0,
    dropped: 0,
  }),
    (summary.opaqueBytes = 0));
  for (const part of summary.parts)
    (summary.counts[part.disposition]++,
      part.disposition === "preserved-opaque" &&
        (summary.opaqueBytes += part.bytes));
  return summary;
}
const numericPrecisionPolicy = Object.freeze({
  storage: "IEEE-754 binary64",
  excelSignificantDigits: 15,
  rounding:
    "ROUND uses decimal ties away from zero; monetary rounding is explicit, never automatic.",
  display:
    "General displays 15 significant digits. Display precision does not round stored values or calculations.",
});
function validateXlsxNumericValue(rawValue) {
  const numericMatch = /^[+-]?(\d+\.?\d*|\.\d+)(?:e([+-]?\d+))?$/i.exec(
    rawValue.trim(),
  );
  if (!numericMatch) return;
  const significantDigits = numericMatch[1]
      .replace(".", "")
      .replace(/^0+/, "")
      .replace(/0+$/, ""),
    numericValue = Number(rawValue);
  if (!Number.isFinite(numericValue)) return "outside the finite numeric range";
  if (
    significantDigits &&
    (numericValue === 0 || Math.abs(numericValue) < 22250738585072014e-324)
  )
    return "below Excel's normal floating-point range";
  if (significantDigits.length > 15) return "more than 15 significant digits";
  if (Number.isInteger(numericValue) && !Number.isSafeInteger(numericValue))
    return "outside the exact integer range";
}
function scanXlsxNumericTokens(xmlText) {
  const numericTokens = new Map(),
    parser = new Ox.SaxesParser({ xmlns: true });
  let cellAddress = "",
    cellType = "",
    valueText = "",
    isInValue = false;
  parser.on("opentag", (openTag) => {
    openTag.local === "c"
      ? ((cellAddress = String(openTag.attributes.r?.value ?? "")),
        (cellType = String(openTag.attributes.t?.value ?? "n")))
      : openTag.local === "v" &&
        cellAddress &&
        ((isInValue = true), (valueText = ""));
  });
  const collectText = (textChunk) => {
    if (isInValue && ((valueText += textChunk), valueText.length > 1024))
      throw Error("An XLSX numeric token exceeds 1,024 characters.");
  };
  return (
    parser.on("text", collectText),
    parser.on("cdata", collectText),
    parser.on("closetag", (closeTag) => {
      if (closeTag.local === "v") {
        if (cellType === "n" && valueText.trim()) {
          const numericText = valueText.trim();
          if (
            !/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(numericText) ||
            !Number.isFinite(Number(numericText))
          )
            throw Error(`Invalid finite numeric value at ${cellAddress}.`);
          numericTokens.set(cellAddress, numericText);
        }
        isInValue = false;
      } else closeTag.local === "c" && (cellAddress = "");
    }),
    parser.write(xmlText).close(),
    numericTokens
  );
}
const formatRangeCornerCells = (rangeText) => {
  const rangeAddress = parseRangeAddress(rangeText);
  return `${formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol)}`;
};
function parseXlsxDataValidations(xmlText) {
  const validations = [];
  for (const validationMatch of xmlText.matchAll(
    /<(?:\w+:)?dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?dataValidation>)/g,
  )) {
    const readAttribute = (attributeName) =>
        new RegExp("\\b" + attributeName + `=["\\']([^"\\']*)["\\']`).exec(
          validationMatch[1],
        )?.[1],
      sqrefRanges = (readAttribute("sqref") ?? "")
        .trim()
        .split(/\s+/)
        .filter(Boolean);
    if (!sqrefRanges.length) continue;
    const anchor = formatRangeCornerCells(sqrefRanges[0]).split(":")[0];
    if (validations.length + sqrefRanges.length > 1e3)
      throw Error(
        "XLSX has more than 1,000 validation areas; narrow the source rules before import.",
      );
    const operands = [
      ...(validationMatch[2] ?? "").matchAll(
        /<(?:\w+:)?formula[12]\b[^>]*>([\s\S]*?)<\/(?:\w+:)?formula[12]>/g,
      ),
    ].map((operandMatch) => operandMatch[1].trim());
    for (const sqrefRange of sqrefRanges)
      validations.push({
        range: formatRangeCornerCells(sqrefRange),
        anchor: anchor,
        operands: operands,
        dropdown: !["1", "true"].includes(readAttribute("showDropDown") ?? "0"),
        alert: ["1", "true"].includes(readAttribute("showErrorMessage") ?? "0"),
      });
  }
  return validations;
}
const dataValidationOperators = {
  between: "between",
  notBetween: "notBetween",
  greater: "greaterThan",
  greaterEqual: "greaterThanOrEqual",
  less: "lessThan",
  lessEqual: "lessThanOrEqual",
  equal: "equal",
  notEqual: "notEqual",
};
const isExternalReferenceFormula = (formula) => {
  const sanitizedFormula = stripFormulaStringLiterals(formula).replace(
    /"(?:[^"]|"")*"/g,
    '""',
  );
  return (
    /[\[\]|]/.test(sanitizedFormula) ||
    /(?:https?:|file:|javascript:|data:|\\\\)/i.test(sanitizedFormula) ||
    /\b(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY)\s*\(/i.test(
      sanitizedFormula.replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet"),
    )
  );
};
const quoteFormulaString = (literalText) =>
  '"' + literalText.replace(/"/g, '""') + '"';
function buildContainsTextFormula(searchText, cellAddress) {
  return (
    "ISNUMBER(FIND(LOWER(" +
    quoteFormulaString(searchText) +
    "),LOWER(" +
    cellAddress +
    '&"")))'
  );
}
function Wle(ruleType, rangeRef) {
  const rangeAddress = parseRangeAddress(rangeRef),
    cellAddress = formatCellAddress(
      rangeAddress.firstRow,
      rangeAddress.firstCol,
    ),
    absoluteRange = formatRangeCornerCells(rangeRef).replace(
      /([A-Z]+)(\d+)/g,
      "$$$1$$$2",
    ),
    criteriaCell =
      "SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" +
      cellAddress +
      ',"~","~~"),"*","~*"),"?","~?")';
  return (
    "AND(LEN(" +
    cellAddress +
    ")>0,COUNTIF(" +
    absoluteRange +
    ',"="&' +
    criteriaCell +
    ")" +
    (ruleType === "duplicate" ? ">1" : "=1") +
    ")"
  );
}
function Hle(raw) {
  if (raw === null || raw === "") return true;
  if (typeof raw == "number") return Number.isFinite(raw);
  if (typeof raw != "string") return false;
  const isNumericString = raw.trim() !== "" && Number.isFinite(Number(raw));
  return !(
    /[^\x20-\x7e]/.test(raw) ||
    /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|ERROR!|CYCLE!|SPILL!|CALC!|LIMIT!|BLOCKED!)/.test(
      raw,
    ) ||
    raw.replace(/[~*?]/g, "~~").length > 254 ||
    (isNumericString && String(Number(raw)) !== raw) ||
    (!isNumericString && /^[\s+-]*\d/.test(raw))
  );
}
function validateConditionalFormatRules(xmlText) {
  const stopIfTrueByPriority = {},
    seenPriorities = new Set();
  for (const ruleMatch of xmlText.matchAll(/<(?:\w+:)?cfRule\b([^>]*)>/g)) {
    const priorityText = /\bpriority=["']([^"']*)["']/.exec(ruleMatch[1])?.[1],
      stopIfTrueText = /\bstopIfTrue=["']([^"']+)["']/.exec(ruleMatch[1])?.[1];
    if (priorityText === undefined && /\bid=["']/.test(ruleMatch[1])) continue;
    const priority = Number(priorityText);
    if (
      !priorityText ||
      !/^\d+$/.test(priorityText) ||
      !Number.isSafeInteger(priority) ||
      priority < 1 ||
      seenPriorities.has(priority)
    )
      throw Error("Invalid or duplicate conditional-format priority.");
    if (
      (seenPriorities.add(priority),
      priorityText && stopIfTrueText !== undefined)
    ) {
      if (!["1", "0", "true", "false"].includes(stopIfTrueText))
        throw Error("Invalid conditional-format stop flag.");
      stopIfTrueByPriority[String(priority)] =
        stopIfTrueText === "1" || stopIfTrueText === "true";
    }
  }
  for (const cfvoMatch of xmlText.matchAll(/<(?:\w+:)?cfvo\b([^>]*)>/g)) {
    const gteText = /\bgte=["']([^"']*)["']/.exec(cfvoMatch[1])?.[1];
    if (gteText !== undefined && gteText !== "1" && gteText !== "true")
      throw Error(
        "Strict or malformed conditional icon thresholds are not supported. Use inclusive thresholds or native JSON.",
      );
  }
  return stopIfTrueByPriority;
}
function parseCalculationProperties(xmlText) {
  let calculationProperties = {
    mode: "auto",
    iterative: false,
    maxIterations: null,
    maxChange: null,
  };
  const parser = new Ox.SaxesParser({ xmlns: true });
  let hasCalculationProperties = false;
  return (
    parser.on("opentag", (tag) => {
      if (tag.local !== "calcPr" || !/spreadsheetml/.test(tag.uri)) return;
      if (hasCalculationProperties)
        throw Error(
          "Duplicate XLSX calculation properties. Repair the source workbook.",
        );
      hasCalculationProperties = true;
      const readAttribute = (attributeName) => {
          const attribute = tag.attributes[attributeName];
          return attribute && typeof attribute != "string"
            ? attribute.value
            : attribute;
        },
        mode = readAttribute("calcMode") ?? "auto",
        iterateValue = readAttribute("iterate"),
        iterateCount = readAttribute("iterateCount"),
        iterateDelta = readAttribute("iterateDelta"),
        fullPrecisionValue = readAttribute("fullPrecision");
      if (
        !["auto", "autoNoTable", "manual"].includes(mode) ||
        (fullPrecisionValue !== undefined &&
          !["true", "false", "1", "0"].includes(fullPrecisionValue)) ||
        (iterateValue !== undefined &&
          !["true", "false", "1", "0"].includes(iterateValue)) ||
        (iterateCount !== undefined &&
          (!/^\d+$/.test(iterateCount) ||
            !Number.isSafeInteger(Number(iterateCount)) ||
            Number(iterateCount) < 1)) ||
        (iterateDelta !== undefined &&
          (!iterateDelta.trim() ||
            !Number.isFinite(Number(iterateDelta)) ||
            Number(iterateDelta) < 0))
      )
        throw Error(
          "Invalid XLSX calculation properties. Repair the source workbook.",
        );
      calculationProperties = {
        mode: mode,
        iterative: iterateValue === "true" || iterateValue === "1",
        maxIterations: iterateCount === undefined ? null : Number(iterateCount),
        maxChange: iterateDelta === undefined ? null : Number(iterateDelta),
        ...(fullPrecisionValue === undefined
          ? {}
          : {
              fullPrecision:
                fullPrecisionValue === "true" || fullPrecisionValue === "1",
            }),
      };
    }),
    parser.on("doctype", () => {
      throw Error("XLSX calculation properties cannot contain a DTD.");
    }),
    parser.write(xmlText).close(),
    calculationProperties
  );
}
function flattenEmbeddedObjectItems(workbook) {
  return (workbook.embeddedObjectInventories ?? []).flatMap((inventory) =>
    inventory.items.map((item) => ({
      sourceFile: inventory.fileName,
      sourceSha256: inventory.sourceSha256,
      path: item.path,
      kind: item.kind,
      bytes: item.bytes,
      ...(item.sha256 ? { sha256: item.sha256 } : {}),
      disposition: "dropped",
      execution: "never",
      reason:
        "Inventory retained in native JSON only; unsupported OLE/ActiveX payload was not imported and is not emitted in derived XLSX.",
    })),
  );
}
function buildExchangeSchemaReport(
  direction,
  fileName,
  fileBytes,
  workbook,
  xlsxWorkbookPackage,
) {
  return {
    schemaVersion: 1,
    direction: direction,
    fileName: fileName,
    fileBytes: fileBytes,
    workbookName: workbook.name,
    revision: workbook.revision,
    compatibility: "derived-subset",
    lossless: false,
    originalFileRetained: !!workbook.xlsxSource,
    ...(workbook.alternateSource
      ? {
          alternateOriginal: {
            fileName: workbook.alternateSource.fileName,
            byteLength: workbook.alternateSource.byteLength,
            sha256: workbook.alternateSource.sha256,
            format: workbook.alternateSource.format,
          },
        }
      : {}),
    ...(workbook.xlsxSource
      ? {
          originalSource: {
            fileName: workbook.xlsxSource.fileName,
            byteLength: workbook.xlsxSource.byteLength,
            sha256: workbook.xlsxSource.sha256,
          },
        }
      : {}),
    sheets: workbook.sheets.map((sheet) => ({
      name: sheet.name,
      storedCells: Object.keys(sheet.cells).length,
      hidden: !!(sheet.hidden || sheet.veryHidden),
    })),
    warnings: [...new Set(xlsxWorkbookPackage)],
    ...(direction === "export" && workbook.embeddedObjectInventories?.length
      ? { embeddedObjectLosses: flattenEmbeddedObjectItems(workbook) }
      : {}),
    caches: {
      total: 0,
      pending: 0,
      matched: 0,
      different: 0,
      missing: 0,
      neutralized: 0,
      written: 0,
      omitted: 0,
      issues: [],
      matches: [],
      detailsOmitted: 0,
    },
  };
}
const delimitedImportDefaults = {
  delimiter: "auto",
  encoding: "auto",
  quote: '"',
  escape: "double",
  newline: "auto",
  hasHeader: true,
  decimal: ".",
  grouping: "none",
  dateOrder: "ymd",
  empty: "blank",
  trim: false,
  columns: {},
};
const delimitedColumnTypes = [
  "auto",
  "text",
  "number",
  "integer",
  "boolean",
  "date",
];
const byteCharacterTable = Array.from({ length: 256 }, (e, byteValue) =>
  String.fromCharCode(byteValue),
);
[
  8364, 129, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 352, 8249, 338, 141,
  381, 143, 144, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732, 8482, 353, 8250,
  339, 157, 382, 376,
].forEach(
  (codePoint, offset) =>
    (byteCharacterTable[128 + offset] = String.fromCharCode(codePoint)),
);
const recordDelimitedIssue = (report, record, column, message, value) => {
  (report.issueCount++,
    report.issues.length < 100 &&
      report.issues.push({
        record: record,
        column: column,
        message: message,
        ...(value !== undefined ? { value: value.slice(0, 200) } : {}),
      }));
};
const recordDelimitedChange = (
  report,
  record,
  column,
  address,
  reason,
  before,
  after,
) => {
  (report.changeCount++,
    report.changes.push({
      record: record,
      column: column,
      address: address,
      reason: reason,
      before: before,
      after: after,
    }));
};
function normalizeDelimitedOptions(options = {}) {
  const normalized = {
    ...delimitedImportDefaults,
    ...options,
    columns: { ...(options.columns ?? {}) },
  };
  for (const [settingName, allowedValues] of Object.entries({
    delimiter: ["auto", ",", ";", "	", "|"],
    encoding: ["auto", "utf-8", "utf-16le", "utf-16be", "windows-1252"],
    quote: ['"', "'", "none"],
    escape: ["double", "backslash"],
    newline: [
      "auto",
      `\r
`,
      `
`,
      "\r",
    ],
    decimal: [".", ","],
    grouping: ["none", ".", ",", " "],
    dateOrder: ["ymd", "dmy", "mdy"],
    empty: ["blank", "text"],
  }))
    if (!allowedValues.includes(String(normalized[settingName])))
      throw Error(`Unsupported delimited ${settingName} setting.`);
  if (
    typeof normalized.hasHeader != "boolean" ||
    typeof normalized.trim != "boolean"
  )
    throw Error("Header and whitespace controls must be boolean.");
  if (normalized.decimal === normalized.grouping)
    throw Error("Decimal and grouping separators must differ.");
  for (const [columnIndex, columnType] of Object.entries(normalized.columns))
    if (
      !/^(0|[1-9]\d*)$/.test(columnIndex) ||
      Number(columnIndex) >= workbookLimits.maxCols ||
      !delimitedColumnTypes.includes(columnType)
    )
      throw Error("Invalid column type selection.");
  return normalized;
}
function buildDelimitedImportReport(
  direction,
  fileName,
  fileBytes,
  delimitedOptions,
) {
  return {
    schemaVersion: 1,
    direction: direction,
    fileName: fileName,
    fileBytes: fileBytes,
    options: delimitedOptions,
    records: 0,
    columns: 0,
    storedCells: 0,
    blankRecords: 0,
    schema: [],
    preview: [],
    issueCount: 0,
    issues: [],
    changeCount: 0,
    changes: [],
    safeTextChanges: 0,
    complete: true,
    valid: false,
    warnings: [],
    safeExportPolicy: "apostrophe-prefix-formula-like-text",
  };
}
const escapeRegExpSource = (text) =>
  text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
function parseDelimitedNumber(raw, delimitedOptions) {
  const groupingSeparator =
      delimitedOptions.grouping === "none" ? "" : delimitedOptions.grouping,
    integerPattern = groupingSeparator
      ? `(?:\\d+|\\d{1,3}(?:${escapeRegExpSource(groupingSeparator)}\\d{3})+)`
      : "\\d+";
  if (
    !new RegExp(
      `^[+-]?(?:${integerPattern}(?:${escapeRegExpSource(delimitedOptions.decimal)}\\d*)?|${escapeRegExpSource(delimitedOptions.decimal)}\\d+)(?:[eE][+-]?\\d+)?$`,
    ).test(raw)
  )
    throw Error("Invalid number for the selected decimal/grouping separators.");
  const normalizedText = (
      groupingSeparator ? raw.split(groupingSeparator).join("") : raw
    ).replace(delimitedOptions.decimal, "."),
    significantDigits = normalizedText
      .split(/[eE]/)[0]
      .replace(/[^0-9]/g, "")
      .replace(/^0+/, "")
      .replace(/0+$/, "").length,
    numericValue = Number(normalizedText);
  if (
    !Number.isFinite(numericValue) ||
    significantDigits > 15 ||
    (numericValue === 0 && /[1-9]/.test(normalizedText.split(/[eE]/)[0]))
  )
    throw Error(
      "Number is outside supported precision/range; select Text to retain every digit.",
    );
  return numericValue;
}
const maxXlsxExpandedByteLimit = fileImportLimits.maxXlsxExpandedBytes;
const xlsxExchangeSupportNotice =
  "XLSX exchange supports cell values, formulas, common formats, supported fonts/strike/alignment/indent/per-edge borders, notes, hidden sheets, column widths, explicit row heights, worksheet default dimensions, bounded hidden rows/columns, frozen rows and columns, basic tables, safe merges, supported validation/conditional-format rules, PNG/JPEG pictures, repeated print titles, full-axis manual print breaks and page setup. XLSX processes at most 100,000 hidden rows per workbook; native JSON keeps larger visibility intervals compactly. Workbook and worksheet-scoped defined names, constant/formula definitions and supported name comments are preserved. Unsupported name attributes are reported; unsafe or malformed definitions stop import/export. Enhanced number formats export supported masks; locale, scaling labels, duration/date-system and accounting alignment fidelity may differ, and enhanced format metadata remains native JSON only. Reusable calculated-column rules are native JSON only; XLSX imports retain supported cell formulas but do not infer rules for future appends. Editable outline groups, threaded discussions and interactive hyperlink metadata remain native-JSON features; XLSX retains only collapsed outline visibility. Native pivot/query refresh definitions, charts/shapes, filters, split panes, default-hidden-row settings, password protection and other unsupported Excel metadata are not preserved. External or executable formulas are converted to literal text on import and export. Keep the original workbook; use native JSON for a lossless application backup.";
const isExternalActionFormula = (formulaText) =>
  /[\[\]|]/.test(
    stripFormulaStringLiterals(formulaText).replace(/"(?:[^"]|"")*"/g, '""'),
  ) ||
  /(?:https?:|file:|javascript:|data:|\\\\)/i.test(formulaText) ||
  /\b(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY)\s*\(/i.test(
    formulaText,
  );
function getXmlAttribute(xmlText, attributeName) {
  return new RegExp(
    "(?:^|\\s)" + attributeName + `\\s*=\\s*(["'])([\\s\\S]*?)\\1`,
  ).exec(xmlText)?.[2];
}
function resolvePackagePartPath(partPath) {
  const segments = [];
  for (const segment of (partPath.startsWith("/")
    ? partPath.slice(1)
    : "xl/" + partPath
  ).split("/"))
    segment === ".."
      ? segments.pop()
      : segment !== "." && segments.push(segment);
  return segments.join("/");
}
const decodeXmlEntities = (text) =>
  text.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[a-f\d]+);/gi, (entity) =>
    entity === "&amp;"
      ? "&"
      : entity === "&lt;"
        ? "<"
        : entity === "&gt;"
          ? ">"
          : entity === "&quot;"
            ? '"'
            : entity === "&apos;"
              ? "'"
              : String.fromCodePoint(
                  Number.parseInt(
                    entity.slice(entity[2]?.toLowerCase() === "x" ? 3 : 2, -1),
                    entity[2]?.toLowerCase() === "x" ? 16 : 10,
                  ),
                ),
  );
function createEmptyExportLedger() {
  return {
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
  };
}

/* From 04-cell-input-paste-and-fill.js */
const s$ = 864e5;
const excelEpochUtc = Date.UTC(1899, 11, 31);
class DateSeriesRangeError extends Error {}
function throwDateSeriesRangeError() {
  throw new DateSeriesRangeError(
    "Date series must stay between 1899-12-31 and 9999-12-31.",
  );
}
function dateToExcelSerial(date) {
  const rawSerial = (date.getTime() - excelEpochUtc) / s$,
    serial = rawSerial >= 60 ? rawSerial + 1 : rawSerial;
  return (
    (!Number.isSafeInteger(serial) ||
      serial < 0 ||
      serial > maxExcelDateSerial) &&
      throwDateSeriesRangeError(),
    serial
  );
}
function parseSeriesSeedDate(seedValue, dateSystem = "1900") {
  if (typeof seedValue == "number")
    return (
      (!Number.isSafeInteger(seedValue) ||
        seedValue < 0 ||
        seedValue > dateSystemMaxSerial(dateSystem)) &&
        throwWorkbookError(
          `Date values must be whole Excel ${dateSystem} serials between 0 and ${dateSystemMaxSerial(dateSystem)}.`,
        ),
      seedValue
    );
  if (
    ((typeof seedValue != "string" || !/^\d{4}-\d{2}-\d{2}$/.test(seedValue)) &&
      throwWorkbookError(
        "Enter a date as YYYY-MM-DD or a whole Excel date serial.",
      ),
    seedValue === "1900-02-29" && dateSystem === "1900")
  )
    return 60;
  const [year, month, day] = seedValue.split("-").map(Number),
    utcDate = utcDateFromParts(year, month - 1, day);
  return (
    (utcDate.getUTCFullYear() !== year ||
      utcDate.getUTCMonth() !== month - 1 ||
      utcDate.getUTCDate() !== day) &&
      throwWorkbookError("The series date is not a valid calendar date."),
    parseSeriesSeedDate(
      dateToExcelSerial(utcDate) - dateSystemSerialOffset(dateSystem),
      dateSystem,
    )
  );
}

/* From 07-file-import-and-worker-dispatch.js */
class FileProcessingError extends Error {
  constructor(code, message = fileProcessingErrorMessages[code]) {
    (super(message), (this.code = code), (this.name = "FileProcessingError"));
  }
}
function toFileProcessingError(error) {
  if (error instanceof FileProcessingError)
    return new FileProcessingError(
      error.code,
      sanitizeFileMessage(error.message) ??
        fileProcessingErrorMessages[error.code],
    );
  const message = error instanceof Error ? error.message : "",
    code =
      /exceed|budget|limit|too (?:large|many)|at most|smaller|depth|nesting/i.test(
        message,
      )
        ? "FILE_LIMIT"
        : /macro|active.?content|activex|executable|unsafe|document type|doctype|entity|dangerous|DDE/i.test(
              message,
            )
          ? "FILE_ACTIVE"
          : "FILE_INVALID";
  return new FileProcessingError(
    code,
    sanitizeFileMessage(message) ?? fileProcessingErrorMessages[code],
  );
}
function assertImportFileAllowed(operation, fileName, fileBytes) {
  if (
    typeof fileName != "string" ||
    !fileName.length ||
    fileName.length > 255 ||
    /[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(fileName)
  )
    throw new FileProcessingError("FILE_TYPE");
  if (
    operation === "workbook" &&
    /\.(xlsm|xlsb|xls|xltm|xlam)$/i.test(fileName)
  )
    throw new FileProcessingError(
      "FILE_TYPE",
      "Legacy or macro-enabled Excel formats are not supported. Save as XLSX, CSV or native JSON.",
    );
  if (
    !{
      workbook: /\.(xlsx|xltx|csv|tsv|json)$/i,
      source: /\.(json|xml|pdf|png|jpe?g|html?)$/i,
      formats: /\.(xls|xlsb|xlsm|xltx|xltm|ods)$/i,
      delimited: /\.(csv|tsv)$/i,
      automation: /\.(bas|vba|cls|ts|osts|xlsm|xlam|xltm|xlsx|xltx|xls|xlsb)$/i,
      json: /\.json$/i,
      picture: /\.(png|jpe?g)$/i,
    }[operation]?.test(operation === "picture" ? fileName.trim() : fileName)
  )
    throw new FileProcessingError("FILE_TYPE");
  if (
    !Number.isSafeInteger(fileBytes) ||
    fileBytes < 0 ||
    fileBytes > fileByteLimitFor(operation, fileName)
  )
    throw new FileProcessingError(
      "FILE_LIMIT",
      fileByteLimitFor(operation, fileName) === workbookLimits.maxNativeBytes
        ? "Native workbook exceeds the 50 MB import limit."
        : fileByteLimitFor(operation, fileName) === workbookLimits.maxFileBytes
          ? "Files must be 10 MB or smaller."
          : undefined,
    );
}

/* From 08-xlsx-import-export-and-preservation.js */
function parseDelimitedDateText(raw, dateOrder, dateSystem = "1900") {
  const dateMatch =
    /^(\d{1,4})([-/])(\d{1,2})\2(\d{1,4})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(
      raw,
    );
  if (!dateMatch)
    throw Error(
      "Date must match the selected year/month/day order, optionally HH:mm:ss.",
    );
  const rawParts = [dateMatch[1], dateMatch[3], dateMatch[4]],
    [year, month, day] =
      dateOrder === "ymd"
        ? rawParts
        : dateOrder === "dmy"
          ? [rawParts[2], rawParts[1], rawParts[0]]
          : [rawParts[2], rawParts[0], rawParts[1]];
  if (year.length !== 4) throw Error("Dates require a four-digit year.");
  const hours = Number(dateMatch[5] ?? 0),
    minutes = Number(dateMatch[6] ?? 0),
    seconds = Number(dateMatch[7] ?? 0);
  if (hours > 23 || minutes > 59 || seconds > 59)
    throw Error("Invalid date/time component.");
  return (
    parseSeriesSeedDate(
      `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`,
      dateSystem,
    ) +
    (hours * 3600 + minutes * 60 + seconds) / 86400
  );
}
function parseDelimitedCellValue(
  cellText,
  columnType,
  delimitedOptions,
  dateSystem = "1900",
) {
  if (cellText === "")
    return { raw: delimitedOptions.empty === "text" ? "" : null };
  const raw = delimitedOptions.trim ? cellText.trim() : cellText;
  if (columnType === "auto") {
    if (
      /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(raw) &&
      raw.split(/[eE]/)[0].replace(/[^0-9]/g, "").length <= 15 &&
      delimitedOptions.decimal === "." &&
      delimitedOptions.grouping === "none"
    )
      try {
        return { raw: parseDelimitedNumber(raw, delimitedOptions) };
      } catch {}
    return { raw: raw, ...(raw.startsWith("=") ? { literal: true } : {}) };
  }
  if (columnType === "text")
    return { raw: raw, ...(raw.startsWith("=") ? { literal: true } : {}) };
  if (columnType === "date")
    return {
      raw: parseDelimitedDateText(raw, delimitedOptions.dateOrder, dateSystem),
      format: "date",
      numFmt: /[ T]\d\d:/.test(raw) ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd",
    };
  if (columnType === "boolean") {
    if (/^(true|false)$/i.test(raw))
      return { raw: raw.toLowerCase() === "true" };
    throw Error("Boolean values must be TRUE or FALSE (case-insensitive).");
  }
  const delimitedNumber = parseDelimitedNumber(raw, delimitedOptions);
  if (columnType === "integer" && !Number.isSafeInteger(delimitedNumber))
    throw Error("Integer column requires a safe whole number.");
  return { raw: delimitedNumber };
}

export {
  DateSeriesRangeError,
  Dle,
  F$,
  FileProcessingError,
  Hle,
  Wle,
  assertImportFileAllowed,
  assertXlsxExpansionLimit,
  buildContainsTextFormula,
  buildDelimitedImportReport,
  buildExchangeSchemaReport,
  byteCharacterTable,
  classifyPreservedPart,
  computeCrc32,
  convertTemplateToXlsx,
  crc32Table,
  createEmptyExportLedger,
  createWorkbookParseState,
  customPropertiesContentType,
  customXmlPropertiesContentType,
  dataValidationOperators,
  dateToExcelSerial,
  decodeXmlEntities,
  delimitedColumnTypes,
  delimitedImportDefaults,
  escapeRegExpSource,
  excelEpochUtc,
  excelRepairAdvice,
  excelRepairSuffix,
  fileByteLimitFor,
  flattenEmbeddedObjectItems,
  formatCellAddress,
  formatInspectionRecords,
  formatRangeCornerCells,
  getXmlAttribute,
  inspectAlternateWorkbookArchive,
  inspectXlsxPackage,
  isExternalActionFormula,
  isExternalReferenceFormula,
  macroEnabledContentTypes,
  maxExcelDateSerial,
  maxXlsxExpandedByteLimit,
  normalizeDelimitedOptions,
  numericPrecisionPolicy,
  officeDocumentRelationshipPrefix,
  ownerPartForRelsPath,
  packageRelationshipsNamespace,
  parseAlternateWorkbookXml,
  parseCalculationProperties,
  parseDelimitedCellValue,
  parseDelimitedDateText,
  parseDelimitedNumber,
  parseSeriesSeedDate,
  parseWorksheetCells,
  parseWorksheetXml,
  parseXlsxDataValidations,
  parseXmlPartStrict,
  parseZipDirectory,
  quoteFormulaString,
  readOpenDocumentPackage,
  readPackageRelationships,
  recordDelimitedChange,
  recordDelimitedIssue,
  recordSourceInspection,
  relationshipsPathFor,
  resolvePackagePartPath,
  resolveRelationshipTarget,
  s$,
  sanitizeFileMessage,
  sanitizedMessagePatterns,
  scanXlsxNumericTokens,
  sourceInspectionRecords,
  spreadsheetMainNamespace,
  summarizePreservedParts,
  templateConversionRoutes,
  throwDateSeriesRangeError,
  throwInvalidXlsx,
  toFileProcessingError,
  u1,
  unsupportedArrayFormulaReasons,
  utcDateFromParts,
  validateConditionalFormatRules,
  validateIsolatedFormatInspection,
  validateWorkbookPackage,
  validateXlsxNumericValue,
  xlsxExchangeSupportNotice,
  xlsxMainContentType,
  z2,
};
