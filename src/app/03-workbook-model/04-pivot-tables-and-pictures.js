import {
  isNonArrayObject,
  throwWorkbookError,
} from "./01-cell-styles-and-themes.js";
import { isNullOrEmptyValue } from "./03-defined-names.js";
import { boundedTextSchema } from "../01-schemas-date-functions-and-validation.js";
import { flattenNestedValues } from "../02-text-and-financial-functions.js";
import {
  workbookLimits,
  supportedFinancialFunctionNames,
} from "../02-text-and-financial-functions.js";
import { strToU8 } from "fflate";
import {
  string,
  number,
  boolean,
  null as _null,
  array,
  strictObject,
  discriminatedUnion,
  literal,
  record,
} from "zod";
import {
  isoDateTimeSchemaFactory,
  deepFreezeValue,
  sha256HexDigest,
  supportedFunctionNames,
  normalizeFunctionName,
  parseFormulaTree,
} from "../01-schemas-date-functions-and-validation.js";
import { assertSheetUnprotected } from "./01-cell-styles-and-themes.js";
import {
  FormulaParser,
  statisticalFunctionNames,
  XL,
  specialistFunctionNames,
  significanceTestFunctionNames,
  ownedLookupFunctionNames,
  modernFunctionNames,
} from "./02-formula-function-evaluators.js";
import {
  definedNameScopeKey,
  hashPictureBytes,
  maxPictureHashBytes,
  stripFormulaStringLiterals,
  replaceTextSpans,
  localBindingFunctionNames,
  m5,
  validateNameFormulaText,
  formatQuotedSheetPrefix,
  validateDrawingText,
} from "./03-defined-names.js";

function inspectPictureDataUrl(dataURL) {
  if (
    typeof dataURL != "string" ||
    dataURL.length > Math.ceil(maxPictureHashBytes / 3) * 4 + 30
  )
    throw Error("Pictures must be local PNG/JPEG data of at most 2 MB.");
  const dataUrlMatch =
    /^data:image\/(png|jpeg);base64,([A-Za-z0-9+/]+={0,2})$/.exec(dataURL);
  if (!dataUrlMatch || dataUrlMatch[2].length % 4 !== 0)
    throw Error("Pictures require canonical local PNG/JPEG base64 data.");
  let bytes;
  try {
    const binaryString = atob(dataUrlMatch[2]);
    if (btoa(binaryString) !== dataUrlMatch[2]) throw Error();
    bytes = Uint8Array.from(binaryString, (character) =>
      character.charCodeAt(0),
    );
  } catch {
    throw Error("The picture data is invalid.");
  }
  if (bytes.length > maxPictureHashBytes)
    throw Error("Pictures must be at most 2 MB.");
  let width = 0,
    height = 0;
  const assertPictureDimensions = () => {
    if (
      !width ||
      !height ||
      width > 1e4 ||
      height > 1e4 ||
      width * height > 4e7
    )
      throw Error(
        "Picture dimensions are invalid or exceed 40 megapixels / 10,000 pixels per side.",
      );
  };
  if (dataUrlMatch[1] === "png") {
    if (
      bytes.length < 45 ||
      ![137, 80, 78, 71, 13, 10, 26, 10].every(
        (expectedByte, byteIndex) => bytes[byteIndex] === expectedByte,
      )
    )
      throw Error("The file contents do not match a PNG picture.");
    const pngView = new DataView(
      bytes.buffer,
      bytes.byteOffset,
      bytes.byteLength,
    );
    if (
      pngView.getUint32(8) !== 13 ||
      String.fromCharCode(...bytes.subarray(12, 16)) !== "IHDR"
    )
      throw Error("PNG requires an intact first IHDR chunk.");
    if (
      ((width = pngView.getUint32(16)),
      (height = pngView.getUint32(20)),
      assertPictureDimensions(),
      !{
        0: [1, 2, 4, 8, 16],
        2: [8, 16],
        3: [1, 2, 4, 8],
        4: [8, 16],
        6: [8, 16],
      }[bytes[25]]?.includes(bytes[24]) ||
        bytes[26] !== 0 ||
        bytes[27] !== 0 ||
        bytes[28] > 1)
    )
      throw Error("Unsupported PNG pixel encoding.");
    let chunkOffset = 8,
      chunkCount = 0,
      hasImageData = false,
      hasEndMarker = false;
    for (; chunkOffset < bytes.length;) {
      if (++chunkCount > 4096 || chunkOffset + 12 > bytes.length)
        throw Error("PNG chunk structure is invalid or excessive.");
      const chunkLength = pngView.getUint32(chunkOffset),
        chunkType = String.fromCharCode(
          ...bytes.subarray(chunkOffset + 4, chunkOffset + 8),
        );
      if (
        chunkLength > bytes.length - chunkOffset - 12 ||
        !/^[A-Za-z]{4}$/.test(chunkType)
      )
        throw Error("PNG chunk length or type is invalid.");
      if (chunkType === "IHDR" && chunkOffset !== 8)
        throw Error("PNG has duplicate dimensions.");
      if (
        (chunkType === "IDAT" && chunkLength > 0 && (hasImageData = true),
        chunkType === "IEND")
      ) {
        if (chunkLength !== 0 || chunkOffset + 12 !== bytes.length)
          throw Error("PNG has invalid trailing data.");
        hasEndMarker = true;
      }
      chunkOffset += chunkLength + 12;
    }
    if (!hasImageData || !hasEndMarker)
      throw Error("PNG data or end marker is missing.");
  } else {
    if (bytes.length < 12 || bytes[0] !== 255 || bytes[1] !== 216)
      throw Error("The file contents do not match a JPEG picture.");
    let segmentOffset = 2,
      segmentCount = 0,
      hasScanData = false,
      hasEndOfImage = false;
    for (; segmentOffset < bytes.length;) {
      if (++segmentCount > 65536) throw Error("JPEG has too many segments.");
      if (hasScanData)
        for (; segmentOffset < bytes.length && bytes[segmentOffset] !== 255;)
          segmentOffset++;
      if (segmentOffset >= bytes.length || bytes[segmentOffset++] !== 255)
        throw Error("JPEG marker structure is invalid.");
      for (; bytes[segmentOffset] === 255;) segmentOffset++;
      const markerCode = bytes[segmentOffset++];
      if (
        hasScanData &&
        (markerCode === 0 || (markerCode >= 208 && markerCode <= 215))
      )
        continue;
      if (markerCode === 217) {
        if (segmentOffset !== bytes.length)
          throw Error("JPEG has trailing data.");
        hasEndOfImage = true;
        break;
      }
      if (markerCode === 1) continue;
      if (segmentOffset + 2 > bytes.length)
        throw Error("JPEG segment is truncated.");
      const segmentLength =
        bytes[segmentOffset] * 256 + bytes[segmentOffset + 1];
      if (segmentLength < 2 || segmentOffset + segmentLength > bytes.length)
        throw Error("JPEG segment length is invalid.");
      if (
        markerCode >= 192 &&
        markerCode <= 207 &&
        ![196, 200, 204].includes(markerCode)
      ) {
        if (
          ![192, 193, 194].includes(markerCode) ||
          segmentLength < 8 ||
          bytes[segmentOffset + 2] !== 8
        )
          throw Error("Unsupported JPEG pixel encoding.");
        const componentCount = bytes[segmentOffset + 7];
        if (
          componentCount < 1 ||
          componentCount > 4 ||
          segmentLength !== 8 + componentCount * 3
        )
          throw Error("JPEG frame components are invalid.");
        const frameHeight =
            bytes[segmentOffset + 3] * 256 + bytes[segmentOffset + 4],
          frameWidth =
            bytes[segmentOffset + 5] * 256 + bytes[segmentOffset + 6];
        if (width && (width !== frameWidth || height !== frameHeight))
          throw Error("JPEG dimensions conflict.");
        ((width = frameWidth),
          (height = frameHeight),
          assertPictureDimensions());
      }
      if (markerCode === 218) {
        if (
          !width ||
          segmentLength < 6 ||
          segmentLength !== 6 + bytes[segmentOffset + 2] * 2
        )
          throw Error("JPEG scan is invalid.");
        hasScanData = true;
      }
      segmentOffset += segmentLength;
    }
    if (!hasScanData || !hasEndOfImage)
      throw Error("JPEG scan or end marker is missing.");
    assertPictureDimensions();
  }
  return {
    bytes: bytes,
    extension: dataUrlMatch[1],
    width: width,
    height: height,
  };
}
function assertPlainTextLabel(text, maxLength, label) {
  const validatedText = validateDrawingText(text, maxLength, label);
  if (!validatedText.trim() || /[\u0000-\u001f\u007f]/.test(validatedText))
    throw Error(
      `Invalid picture ${label}: use nonblank plain text without control characters.`,
    );
  return validatedText;
}
function validatePictureEvidence(evidence, bytes) {
  if (
    !isNonArrayObject(evidence) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(evidence)) ||
    Reflect.ownKeys(evidence).some(
      (key) =>
        typeof key != "string" ||
        ![
          "version",
          "origin",
          "sourceReference",
          "sha256",
          "byteLength",
          "filename",
          "sourceLocator",
          "containerFilename",
          "mediaReference",
        ].includes(key),
    )
  )
    throw Error("Invalid picture evidence metadata.");
  if (
    evidence.version !== 1 ||
    typeof evidence.origin != "string" ||
    !["local-file", "xlsx-media", "legacy-native"].includes(evidence.origin)
  )
    throw Error("Invalid picture evidence origin.");
  const sha256 = hashPictureBytes(bytes);
  if (
    typeof evidence.sha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(evidence.sha256) ||
    evidence.sha256 !== sha256 ||
    evidence.sourceReference !== `sha256:${sha256}` ||
    evidence.byteLength !== bytes.length
  )
    throw Error(
      "Picture evidence digest or byte length does not match the embedded image.",
    );
  const normalizedEvidence = {
    version: 1,
    origin: evidence.origin,
    sourceReference: `sha256:${sha256}`,
    sha256: sha256,
    byteLength: bytes.length,
  };
  for (const metadataKey of [
    "filename",
    "sourceLocator",
    "containerFilename",
    "mediaReference",
  ])
    evidence[metadataKey] !== undefined &&
      (normalizedEvidence[metadataKey] = assertPlainTextLabel(
        evidence[metadataKey],
        metadataKey === "sourceLocator" ? 2e3 : 512,
        metadataKey,
      ));
  if (
    normalizedEvidence.origin === "local-file" &&
    !normalizedEvidence.filename
  )
    throw Error("A local evidence picture requires its original filename.");
  if (
    normalizedEvidence.origin !== "xlsx-media" &&
    (normalizedEvidence.containerFilename || normalizedEvidence.mediaReference)
  )
    throw Error("Only observed XLSX media may have container metadata.");
  if (
    normalizedEvidence.origin === "legacy-native" &&
    normalizedEvidence.filename
  )
    throw Error(
      "An original filename cannot be inferred for a legacy picture.",
    );
  return normalizedEvidence;
}
function createPictureEvidence(dataURL, metadata) {
  const { bytes } = inspectPictureDataUrl(dataURL),
    sha256 = hashPictureBytes(bytes);
  return validatePictureEvidence(
    {
      ...metadata,
      version: 1,
      sourceReference: `sha256:${sha256}`,
      sha256: sha256,
      byteLength: bytes.length,
    },
    bytes,
  );
}
function F2(workbook, sheetId, selection) {
  if (!workbook.sheets.some((sheet) => sheet.id === sheetId))
    throw Error("The worksheet no longer exists.");
  return {
    workbookId: workbook.id,
    revision: workbook.revision,
    sheetId: sheetId,
    ...(selection !== undefined ? { selection: selection } : {}),
    guard: JSON.stringify(workbook),
  };
}
function assertDrawingPreviewCurrent(workbook, sheetId, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("The worksheet no longer exists.");
  assertSheetUnprotected(sheet);
  const snapshot = options.snapshot;
  if (
    snapshot &&
    (snapshot.workbookId !== workbook.id ||
      snapshot.revision !== workbook.revision ||
      snapshot.sheetId !== sheetId ||
      snapshot.guard !== JSON.stringify(workbook) ||
      snapshot.selection !== options.selection)
  )
    throw Error(
      "The drawing preview is stale. Reopen it or select the object again.",
    );
}
const throwErrorMessage = (message) => {
    throw Error(message);
  },
  isFiniteNumber = (value) =>
    typeof value == "number" && Number.isFinite(value);
function validatePivotRefreshEvidence(refresh) {
  isNonArrayObject(refresh) ||
    throwErrorMessage("Invalid pivot refresh evidence.");
  const evidence = refresh;
  (typeof evidence.at != "string" ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(evidence.at) ||
    !Number.isFinite(Date.parse(evidence.at)) ||
    new Date(evidence.at).toISOString() !== evidence.at ||
    typeof evidence.sourceSheetName != "string" ||
    !evidence.sourceSheetName ||
    evidence.sourceSheetName.length > 31 ||
    typeof evidence.sourceRange != "string" ||
    !/^[A-Z]+\d+:[A-Z]+\d+$/.test(evidence.sourceRange)) &&
    throwErrorMessage("Invalid pivot refresh time or source label.");
  for (const fingerprintField of ["sourceFingerprint", "outputFingerprint"])
    (typeof evidence[fingerprintField] != "string" ||
      !/^[a-f0-9]{64}$/.test(evidence[fingerprintField])) &&
      throwErrorMessage("Invalid pivot content fingerprint.");
  for (const countField of ["sourceRows", "includedRows", "excludedRows"])
    (!Number.isSafeInteger(evidence[countField]) ||
      Number(evidence[countField]) < 0 ||
      Number(evidence[countField]) > 1e5) &&
      throwErrorMessage("Invalid pivot population counts.");
  (Number(evidence.sourceRows) !==
    Number(evidence.includedRows) + Number(evidence.excludedRows) ||
    !Array.isArray(evidence.measures) ||
    !evidence.measures.length ||
    evidence.measures.length > 100) &&
    throwErrorMessage("Pivot population evidence does not reconcile.");
  const measures = evidence.measures.map((measure) => {
    (!isNonArrayObject(measure) ||
      typeof measure.name != "string" ||
      !measure.name ||
      measure.name.length > 512 ||
      typeof measure.field != "string" ||
      !measure.field ||
      measure.field.length > 255 ||
      !["sum", "count", "average", "min", "max", "distinctCount"].includes(
        String(measure.operation),
      ) ||
      !(
        measure.result === null ||
        typeof measure.result == "boolean" ||
        isFiniteNumber(measure.result) ||
        (typeof measure.result == "string" && measure.result.length <= 32767)
      )) &&
      throwErrorMessage("Invalid pivot measure evidence.");
    const measureEvidence = measure;
    for (const countKey of ["numericCount", "nonblankCount"])
      (!Number.isSafeInteger(measureEvidence[countKey]) ||
        Number(measureEvidence[countKey]) < 0 ||
        Number(measureEvidence[countKey]) > Number(evidence.includedRows)) &&
        throwErrorMessage("Invalid pivot measure count.");
    return (
      Number(measureEvidence.numericCount) >
        Number(measureEvidence.nonblankCount) &&
        throwErrorMessage("Invalid pivot numeric count."),
      {
        name: measureEvidence.name,
        field: measureEvidence.field,
        operation: measureEvidence.operation,
        result: measureEvidence.result,
        numericCount: measureEvidence.numericCount,
        nonblankCount: measureEvidence.nonblankCount,
      }
    );
  });
  return {
    at: evidence.at,
    sourceSheetName: evidence.sourceSheetName,
    sourceRange: evidence.sourceRange,
    sourceFingerprint: evidence.sourceFingerprint,
    outputFingerprint: evidence.outputFingerprint,
    sourceRows: evidence.sourceRows,
    includedRows: evidence.includedRows,
    excludedRows: evidence.excludedRows,
    measures: measures,
  };
}
function bucketPivotDateValue(dateValue, dateGroup, rowIndex) {
  if (isNullOrEmptyValue(dateValue)) return null;
  let date;
  if (dateGroup.dateSystem === "iso") {
    if (typeof dateValue != "string" || !/^\d{4}-\d{2}-\d{2}$/.test(dateValue))
      return throwErrorMessage(
        `Date grouping for “${dateGroup.field}” needs YYYY-MM-DD text on source row ${rowIndex + 1}; choose the correct date system or clean the source.`,
      );
    if (
      ((date = new Date(dateValue + "T00:00:00.000Z")),
      !Number.isFinite(date.getTime()) ||
        date.toISOString().slice(0, 10) !== dateValue ||
        dateValue.startsWith("0000"))
    )
      return throwErrorMessage(
        `Invalid date in “${dateGroup.field}” on source row ${rowIndex + 1}.`,
      );
  } else {
    if (
      !isFiniteNumber(dateValue) ||
      dateValue < (dateGroup.dateSystem === "1900" ? 1 : 0) ||
      dateValue > 2958465
    )
      return throwErrorMessage(
        `Date grouping for “${dateGroup.field}” requires a valid ${dateGroup.dateSystem} numeric date serial on source row ${rowIndex + 1}.`,
      );
    const serialDay = Math.floor(dateValue),
      epochBaseMs =
        dateGroup.dateSystem === "1900"
          ? Date.UTC(1899, 11, 31)
          : Date.UTC(1904, 0, 1);
    if (
      ((date = new Date(
        epochBaseMs +
          (serialDay -
            (dateGroup.dateSystem === "1900" && serialDay >= 60 ? 1 : 0)) *
            864e5,
      )),
      date.getUTCFullYear() > 9999)
    )
      return throwErrorMessage("Grouped date exceeds year 9999.");
  }
  const yearText = String(date.getUTCFullYear()).padStart(4, "0"),
    monthNumber = date.getUTCMonth() + 1;
  return dateGroup.unit === "year"
    ? yearText
    : dateGroup.unit === "quarter"
      ? `${yearText} Q${Math.ceil(monthNumber / 3)}`
      : `${yearText}-${String(monthNumber).padStart(2, "0")}`;
}
const createAggregateAccumulator = (shouldTrackDistinct = false) => ({
  count: 0,
  numeric: 0,
  sum: 0,
  min: 1 / 0,
  max: -1 / 0,
  ...(shouldTrackDistinct ? { distinct: new Set() } : {}),
});
function SO(right, value) {
  (isNullOrEmptyValue(value) ||
    (right.count++,
    right.distinct?.add(
      JSON.stringify([
        typeof value,
        typeof value == "string" ? value.toLocaleLowerCase("en-US") : value,
      ]),
    )),
    isFiniteNumber(value) &&
      (right.numeric++,
      (right.sum += value),
      (right.min = Math.min(right.min, value)),
      (right.max = Math.max(right.max, value)),
      Number.isFinite(right.sum) ||
        throwErrorMessage("Aggregation exceeds finite numeric bounds.")));
}
function finalizeAggregateState(aggregateState, operation) {
  return operation === "count"
    ? aggregateState.count
    : operation === "distinctCount"
      ? (aggregateState.distinct?.size ?? 0)
      : operation === "sum"
        ? aggregateState.sum
        : aggregateState.numeric
          ? operation === "average"
            ? aggregateState.sum / aggregateState.numeric
            : operation === "min"
              ? aggregateState.min
              : aggregateState.max
          : null;
}
function computePivotTable(pivotDefinition, fields, rows) {
  const resolveFieldIndex = (fieldName) => {
      const fieldIndex = fields.indexOf(fieldName);
      return fieldIndex < 0
        ? throwErrorMessage(
            `Source field “${fieldName}” was not found. Refresh the field selection.`,
          )
        : fieldIndex;
    },
    rowFieldIndexes = pivotDefinition.rowFields.map(resolveFieldIndex),
    columnFieldIndexes = pivotDefinition.columnFields.map(resolveFieldIndex);
  for (const dateGroup of pivotDefinition.dateGroups ?? [])
    [...pivotDefinition.rowFields, ...pivotDefinition.columnFields].includes(
      dateGroup.field,
    ) ||
      throwErrorMessage(
        `Date group “${dateGroup.field}” must be selected as a row or column field.`,
      );
  for (const calculatedField of pivotDefinition.calculatedFields ?? [])
    (fields.some(
      (field) => field.toLowerCase() === calculatedField.name.toLowerCase(),
    ) &&
      throwErrorMessage(
        `Calculated field “${calculatedField.name}” duplicates a source header.`,
      ),
      resolveFieldIndex(calculatedField.field),
      calculatedField.otherField &&
        resolveFieldIndex(calculatedField.otherField));
  const valueSpecs = pivotDefinition.values.map((value) => {
      const calculated = pivotDefinition.calculatedFields?.find(
        (calculatedField) => calculatedField.name === value.field,
      );
      return (
        calculated &&
          value.operation !== "sum" &&
          throwErrorMessage(
            `Calculated field “${value.field}” requires sum aggregation; its arithmetic is applied to source-field sums.`,
          ),
        {
          ...value,
          index: resolveFieldIndex(calculated?.field ?? value.field),
          calculated: calculated,
          other: calculated?.otherField
            ? resolveFieldIndex(calculated.otherField)
            : undefined,
        }
      );
    }),
    createGroupAccumulators = () =>
      valueSpecs.map((valueSpec) => ({
        ...createAggregateAccumulator(valueSpec.operation === "distinctCount"),
        ...(valueSpec.other !== undefined
          ? { right: createAggregateAccumulator() }
          : {}),
      }));
  rows.length *
    valueSpecs.length *
    Math.max(1, pivotDefinition.subtotals ? rowFieldIndexes.length : 1) *
    Math.max(1, pivotDefinition.subtotals ? columnFieldIndexes.length : 1) >
    2e6 &&
    throwErrorMessage(
      "Pivot aggregation exceeds the 2,000,000 contribution budget. Reduce value fields, hierarchy depth or source rows.",
    );
  let retainedDistinctEntries = 0;
  const accumulateRowValues = (accumulators, rowValues) =>
      valueSpecs.forEach((valueSpec, valueIndex) => {
        const previousDistinctCount =
          accumulators[valueIndex].distinct?.size ?? 0;
        (SO(accumulators[valueIndex], rowValues[valueSpec.index]),
          (retainedDistinctEntries +=
            (accumulators[valueIndex].distinct?.size ?? 0) -
            previousDistinctCount),
          retainedDistinctEntries > 25e4 &&
            throwErrorMessage(
              "Pivot distinct counts exceed 250,000 retained group entries. Reduce groups, value fields or source rows.",
            ),
          valueSpec.other !== undefined &&
            SO(accumulators[valueIndex].right, rowValues[valueSpec.other]));
      }),
    measureResult = (accumulators, valueIndex) => {
      const valueSpec = valueSpecs[valueIndex],
        accumulator = accumulators?.[valueIndex];
      if (!valueSpec.calculated)
        return finalizeAggregateState(
          accumulator ??
            createAggregateAccumulator(valueSpec.operation === "distinctCount"),
          valueSpec.operation,
        );
      if (!accumulator) return null;
      const sum = accumulator.sum,
        operand =
          valueSpec.other !== undefined
            ? accumulator.right.sum
            : valueSpec.calculated.constant;
      valueSpec.calculated.operation === "divide" &&
        operand === 0 &&
        throwErrorMessage(
          `Calculated field “${valueSpec.field}” divides by zero in a group or total. Filter, clean or change the calculation; no result was written.`,
        );
      const calculatedValue =
        valueSpec.calculated.operation === "add"
          ? sum + operand
          : valueSpec.calculated.operation === "subtract"
            ? sum - operand
            : valueSpec.calculated.operation === "multiply"
              ? sum * operand
              : sum / operand;
      return Number.isFinite(calculatedValue)
        ? calculatedValue
        : throwErrorMessage(
            `Calculated field “${valueSpec.field}” exceeds finite numeric bounds.`,
          );
    },
    readFieldValues = (fieldIndexes, sourceRow) =>
      fieldIndexes.map((fieldIndex) => {
        const matchingDateGroup = pivotDefinition.dateGroups?.find(
            (dateGroup) => dateGroup.field === fields[fieldIndex],
          ),
          rawValue = sourceRow.values[fieldIndex];
        return matchingDateGroup
          ? bucketPivotDateValue(rawValue, matchingDateGroup, sourceRow.index)
          : isNullOrEmptyValue(rawValue)
            ? null
            : rawValue;
      }),
    buildGroupEntries = (keyParts) =>
      [
        keyParts,
        ...(pivotDefinition.subtotals
          ? Array.from(
              { length: Math.max(0, keyParts.length - 1) },
              (T, prefixIndex) => keyParts.slice(0, prefixIndex + 1),
            )
          : []),
      ].map((parts) => ({
        key: JSON.stringify(parts),
        parts: parts,
        subtotal: parts.length < keyParts.length,
      })),
    rowGroups = new Map(),
    columnGroups = new Map(),
    cellGroups = new Map(),
    rowAccumulators = new Map(),
    columnAccumulators = new Map(),
    grandTotalAccumulators = createGroupAccumulators(),
    addToGroup = (groupMap, groupKey, rowValues) => {
      (groupMap.has(groupKey) ||
        groupMap.set(groupKey, createGroupAccumulators()),
        accumulateRowValues(groupMap.get(groupKey), rowValues));
    };
  for (const row of rows) {
    const rowEntries = buildGroupEntries(readFieldValues(rowFieldIndexes, row)),
      columnEntries = buildGroupEntries(
        readFieldValues(columnFieldIndexes, row),
      );
    (rowEntries.forEach((event) => rowGroups.set(event.key, event)),
      columnEntries.forEach((event) => columnGroups.set(event.key, event)));
    const cellsPerRow =
      Math.max(1, rowFieldIndexes.length) +
      (columnGroups.size +
        +(!!pivotDefinition.grandTotals && !!columnFieldIndexes.length)) *
        valueSpecs.length;
    (rowGroups.size + 2) * cellsPerRow > 1e5 &&
      throwErrorMessage(
        "The pivot grouping would exceed the 100,000-cell output budget. Reduce group fields or filter the source.",
      );
    for (const rowEntry of rowEntries) {
      addToGroup(rowAccumulators, rowEntry.key, row.values);
      for (const columnEntry of columnEntries)
        addToGroup(
          cellGroups,
          JSON.stringify([rowEntry.key, columnEntry.key]),
          row.values,
        );
    }
    for (const columnEntry of columnEntries)
      addToGroup(columnAccumulators, columnEntry.key, row.values);
    accumulateRowValues(grandTotalAccumulators, row.values);
  }
  (columnGroups.size ||
    columnGroups.set("[]", { key: "[]", parts: [], subtotal: false }),
    !rowGroups.size &&
      !rowFieldIndexes.length &&
      rowGroups.set("[]", { key: "[]", parts: [], subtotal: false }));
  function orderGroupEntries(groupMap, fieldNames) {
    if (!pivotDefinition.subtotals && !pivotDefinition.sorts?.length)
      return [...groupMap.values()];
    if (!fieldNames.length) return [...groupMap.values()];
    const childrenByParent = new Map();
    for (const entry of groupMap.values())
      if (!entry.subtotal)
        for (let partIndex = 0; partIndex < entry.parts.length; partIndex++) {
          const parentKey = JSON.stringify(entry.parts.slice(0, partIndex)),
            childParts = entry.parts.slice(0, partIndex + 1),
            childKey = JSON.stringify(childParts);
          (childrenByParent.has(parentKey) ||
            childrenByParent.set(parentKey, new Map()),
            childrenByParent.get(parentKey).set(childKey, childParts));
        }
    const orderedEntries = [],
      visitPrefix = (prefixParts) => {
        const sortRule = pivotDefinition.sorts?.find(
            (sort) => sort.field === fieldNames[prefixParts.length],
          ),
          childPartsList = [
            ...(childrenByParent.get(JSON.stringify(prefixParts))?.values() ??
              []),
          ];
        sortRule &&
          childPartsList.sort((leftParts, rightParts) => {
            const leftValue = leftParts.at(-1),
              rightValue = rightParts.at(-1),
              comparison =
                leftValue === rightValue
                  ? 0
                  : leftValue === null
                    ? -1
                    : rightValue === null
                      ? 1
                      : typeof leftValue == "number" &&
                          typeof rightValue == "number"
                        ? leftValue - rightValue
                        : String(leftValue).localeCompare(
                            String(rightValue),
                            "en-US",
                            {
                              numeric: true,
                            },
                          );
            return sortRule.direction === "ascending"
              ? comparison
              : -comparison;
          });
        for (const nextParts of childPartsList) visitPrefix(nextParts);
        const groupEntry = groupMap.get(JSON.stringify(prefixParts));
        groupEntry && orderedEntries.push(groupEntry);
      };
    return (visitPrefix([]), orderedEntries);
  }
  const sortedColumnEntries = orderGroupEntries(
      columnGroups,
      pivotDefinition.columnFields,
    ),
    sortedRowEntries = orderGroupEntries(rowGroups, pivotDefinition.rowFields),
    measureLabel = (valueSpec) =>
      valueSpec.name ??
      (valueSpec.calculated
        ? `Calculated ${valueSpec.field}`
        : `${valueSpec.operation === "distinctCount" ? "Distinct count" : valueSpec.operation[0].toUpperCase() + valueSpec.operation.slice(1)} of ${valueSpec.field}`),
    formatKeyPart = (keyPart) =>
      isNullOrEmptyValue(keyPart) ? "(blank)" : String(keyPart),
    headerRow = pivotDefinition.rowFields.length
      ? pivotDefinition.rowFields.map((rowField) => {
          const rowDateGroup = pivotDefinition.dateGroups?.find(
            (dateGroup) => dateGroup.field === rowField,
          );
          return rowDateGroup ? `${rowField} (${rowDateGroup.unit})` : rowField;
        })
      : ["Row labels"];
  for (const columnEntry of sortedColumnEntries)
    for (const valueSpec of valueSpecs)
      headerRow.push(
        [
          ...columnEntry.parts.map(formatKeyPart),
          ...(columnEntry.subtotal ? ["Subtotal"] : []),
          measureLabel(valueSpec),
        ].join(" / "),
      );
  if (pivotDefinition.grandTotals && columnFieldIndexes.length)
    for (const valueSpec of valueSpecs)
      headerRow.push(`Grand Total / ${measureLabel(valueSpec)}`);
  valueSpecs.forEach((valueSpec, valueIndex) => {
    if (valueSpec.showAs && valueSpec.showAs !== "value")
      for (
        let headerIndex = headerRow.length - 1;
        headerIndex >= Math.max(1, rowFieldIndexes.length);
        headerIndex--
      )
        (headerIndex - Math.max(1, rowFieldIndexes.length)) %
          valueSpecs.length ===
          valueIndex &&
          (headerRow[headerIndex] =
            `${headerRow[headerIndex]} (% of ${valueSpec.showAs === "percentGrand" ? "grand" : valueSpec.showAs === "percentRow" ? "row" : "column"} total)`);
  });
  const rowAxes = [],
    columnAxes = [
      ...sortedColumnEntries,
      ...(pivotDefinition.grandTotals && columnFieldIndexes.length
        ? [null]
        : []),
    ],
    baseMatrix = [headerRow],
    applyShowAs = (cellValue, valueIndex, rowTotalState, columnTotalState) => {
      const showAs = valueSpecs[valueIndex].showAs;
      if (!showAs || showAs === "value" || cellValue === null) return cellValue;
      const baseTotal = measureResult(
        showAs === "percentGrand"
          ? grandTotalAccumulators
          : showAs === "percentRow"
            ? rowTotalState
            : columnTotalState,
        valueIndex,
      );
      return typeof cellValue == "number" &&
        typeof baseTotal == "number" &&
        baseTotal !== 0
        ? cellValue / baseTotal
        : "#DIV/0!";
    },
    matrix = [headerRow],
    subtotalRows = [];
  for (const rowEntry of sortedRowEntries) {
    const rowLabels = rowEntry.subtotal
        ? [
            ...rowEntry.parts,
            "Subtotal",
            ...Array(rowFieldIndexes.length - rowEntry.parts.length - 1).fill(
              null,
            ),
          ]
        : rowEntry.parts.length
          ? [...rowEntry.parts]
          : ["Total"],
      baseRow = [...rowLabels];
    for (const columnAx of columnAxes)
      valueSpecs.forEach((L, valueIndex) => {
        const cellValue = measureResult(
          columnAx
            ? cellGroups.get(JSON.stringify([rowEntry.key, columnAx.key]))
            : rowAccumulators.get(rowEntry.key),
          valueIndex,
        );
        (baseRow.push(cellValue),
          rowLabels.push(
            applyShowAs(
              cellValue,
              valueIndex,
              rowAccumulators.get(rowEntry.key),
              columnAx
                ? columnAccumulators.get(columnAx.key)
                : grandTotalAccumulators,
            ),
          ));
      });
    (rowEntry.subtotal && subtotalRows.push(matrix.length),
      matrix.push(rowLabels),
      baseMatrix.push(baseRow),
      rowAxes.push(rowEntry));
  }
  if (pivotDefinition.grandTotals && rowFieldIndexes.length) {
    const grandTotalRow = [
        "Grand Total",
        ...Array(rowFieldIndexes.length - 1).fill(null),
      ],
      grandTotalBaseRow = [...grandTotalRow];
    for (const columnAx of columnAxes)
      valueSpecs.forEach((R, valueIndex) => {
        const grandTotalValue = measureResult(
          columnAx
            ? columnAccumulators.get(columnAx.key)
            : grandTotalAccumulators,
          valueIndex,
        );
        (grandTotalBaseRow.push(grandTotalValue),
          grandTotalRow.push(
            applyShowAs(
              grandTotalValue,
              valueIndex,
              grandTotalAccumulators,
              columnAx
                ? columnAccumulators.get(columnAx.key)
                : grandTotalAccumulators,
            ),
          ));
      });
    (matrix.push(grandTotalRow),
      baseMatrix.push(grandTotalBaseRow),
      rowAxes.push(null),
      subtotalRows.push(matrix.length - 1));
  }
  return {
    matrix: matrix,
    baseMatrix: baseMatrix,
    rowAxes: rowAxes,
    columnAxes: columnAxes,
    chartRows: sortedRowEntries
      .filter((rowEntry) => !rowEntry.subtotal)
      .map((rowEntry) => ({
        label: rowEntry.parts.map(formatKeyPart).join(" / ") || "Total",
        values: valueSpecs.map((T, valueIndex) =>
          applyShowAs(
            measureResult(rowAccumulators.get(rowEntry.key), valueIndex),
            valueIndex,
            rowAccumulators.get(rowEntry.key),
            grandTotalAccumulators,
          ),
        ),
      })),
    subtotalRows: subtotalRows,
    measures: valueSpecs.map((valueSpec, valueIndex) => ({
      name: measureLabel(valueSpec),
      field: valueSpec.field,
      operation: valueSpec.operation,
      result: measureResult(grandTotalAccumulators, valueIndex),
      numericCount: grandTotalAccumulators[valueIndex].numeric,
      nonblankCount: grandTotalAccumulators[valueIndex].count,
    })),
  };
}
const pivotShowValuesAsModes = [
    "value",
    "percentGrand",
    "percentRow",
    "percentColumn",
  ],
  requirePivotFieldName = (fieldName) =>
    typeof fieldName == "string" && fieldName.trim() && fieldName.length <= 255
      ? fieldName
      : throwErrorMessage("Choose a valid pivot field."),
  isIsoDateString = (value) =>
    typeof value == "string" &&
    /^\d{4}-\d{2}-\d{2}$/.test(value) &&
    !value.startsWith("0000") &&
    Number.isFinite(Date.parse(value)) &&
    new Date(value).toISOString().slice(0, 10) === value,
  requireBoundedList = (entries, maxEntries) =>
    Array.isArray(entries) && entries.length <= maxEntries
      ? entries
      : throwErrorMessage(
          `Pivot controls require a list of at most ${maxEntries} entries.`,
        ),
  buildSlicerItemKey = (itemValue) =>
    JSON.stringify([
      typeof itemValue,
      typeof itemValue == "string"
        ? itemValue.toLocaleLowerCase("en-US")
        : itemValue,
    ]);
function normalizePivotControls(controls) {
  const normalized = {};
  (controls.slicers !== undefined &&
    (normalized.slicers = requireBoundedList(controls.slicers, 20).map(
      (slicer) => {
        if (!isNonArrayObject(slicer))
          return throwErrorMessage("Invalid pivot slicer.");
        const selected = requireBoundedList(slicer.selected, 500).map(
          (item) => (
            item !== null &&
              typeof item != "boolean" &&
              !(typeof item == "number" && Number.isFinite(item)) &&
              !(typeof item == "string" && item.length <= 32767) &&
              throwErrorMessage("Invalid slicer item."),
            item
          ),
        );
        return (
          new Set(selected.map(buildSlicerItemKey)).size !== selected.length &&
            throwErrorMessage("Slicer selections must be unique."),
          { field: requirePivotFieldName(slicer.field), selected: selected }
        );
      },
    )),
    controls.timelines !== undefined &&
      (normalized.timelines = requireBoundedList(controls.timelines, 10).map(
        (timeline) =>
          !isNonArrayObject(timeline) ||
          !isIsoDateString(timeline.from) ||
          !isIsoDateString(timeline.to) ||
          timeline.from > timeline.to ||
          !["year", "quarter", "month"].includes(String(timeline.unit)) ||
          !["iso", "1900", "1904"].includes(String(timeline.dateSystem))
            ? throwErrorMessage(
                "Timeline requires ordered YYYY-MM-DD bounds, a period and an explicit date system.",
              )
            : {
                field: requirePivotFieldName(timeline.field),
                from: timeline.from,
                to: timeline.to,
                unit: timeline.unit,
                dateSystem: timeline.dateSystem,
              },
      )),
    controls.sorts !== undefined &&
      (normalized.sorts = requireBoundedList(controls.sorts, 100).map(
        (sortRule) =>
          !isNonArrayObject(sortRule) ||
          !["ascending", "descending"].includes(String(sortRule.direction))
            ? throwErrorMessage(
                "Pivot field sorting requires ascending or descending order.",
              )
            : {
                field: requirePivotFieldName(sortRule.field),
                direction: sortRule.direction,
              },
      )));
  for (const controlGroup of [
    normalized.slicers,
    normalized.timelines,
    normalized.sorts,
  ])
    controlGroup &&
      new Set(controlGroup.map((control) => control.field)).size !==
        controlGroup.length &&
      throwErrorMessage("Choose only one control of each kind per field.");
  if (controls.chart !== undefined) {
    const chart = controls.chart;
    if (
      !isNonArrayObject(chart) ||
      !["column", "bar", "line"].includes(String(chart.kind)) ||
      !Number.isInteger(chart.valueIndex) ||
      Number(chart.valueIndex) < 0 ||
      Number(chart.valueIndex) >= Number(controls.values?.length)
    )
      return throwErrorMessage(
        "PivotChart requires a supported chart and an existing value field.",
      );
    normalized.chart = { kind: chart.kind, valueIndex: chart.valueIndex };
  }
  return normalized;
}
function validatePivotFieldReferences(definition, fields) {
  const calculatedFieldNames = new Set(
      definition.calculatedFields?.map(
        (calculatedField) => calculatedField.name,
      ),
    ),
    referencedFields = [
      ...definition.rowFields,
      ...definition.columnFields,
      ...(definition.filters ?? []).map((filter) => filter.field),
      ...(definition.slicers ?? []).map((slicer) => slicer.field),
      ...(definition.timelines ?? []).map((timeline) => timeline.field),
      ...(definition.dateGroups ?? []).map((dateGroup) => dateGroup.field),
      ...(definition.sorts ?? []).map((sortEntry) => sortEntry.field),
      ...definition.values
        .filter((value) => !calculatedFieldNames.has(value.field))
        .map((valueEntry) => valueEntry.field),
      ...(definition.calculatedFields ?? []).flatMap((calculatedEntry) => [
        calculatedEntry.field,
        ...(calculatedEntry.otherField ? [calculatedEntry.otherField] : []),
      ]),
    ];
  for (const referencedField of referencedFields)
    fields.includes(referencedField) ||
      throwErrorMessage(
        `Broken pivot field “${referencedField}”: source field was not found. Restore the header or remove/reselect this field before refreshing. Historical totals are not current.`,
      );
  for (const sortRule of definition.sorts ?? [])
    [...definition.rowFields, ...definition.columnFields].includes(
      sortRule.field,
    ) ||
      throwErrorMessage(
        `Sort field “${sortRule.field}” must be a row or column group.`,
      );
}
function normalizeTimelineDate(dateValue, dateSystem, field, rowIndex) {
  if (dateValue === null || dateValue === "") return null;
  if (dateSystem === "iso")
    return isIsoDateString(dateValue)
      ? dateValue
      : throwErrorMessage(
          `Timeline field “${field}” has an invalid ISO date on source row ${rowIndex + 1}.`,
        );
  if (
    typeof dateValue != "number" ||
    !Number.isFinite(dateValue) ||
    dateValue < (dateSystem === "1900" ? 1 : 0) ||
    dateValue > 2958465
  )
    return throwErrorMessage(
      `Timeline field “${field}” requires ${dateSystem} date serials on source row ${rowIndex + 1}.`,
    );
  const serialDay = Math.floor(dateValue);
  if (dateSystem === "1900" && serialDay === 60) return "1900-02-29";
  const serialDate = new Date(
    (dateSystem === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1)) +
      (serialDay - (dateSystem === "1900" && serialDay >= 60 ? 1 : 0)) * 864e5,
  );
  return serialDate.getUTCFullYear() > 9999
    ? throwErrorMessage("Timeline date exceeds year 9999.")
    : serialDate.toISOString().slice(0, 10);
}
function filterPivotSourceRows(pivotDefinition, fields, sourceRows) {
  validatePivotFieldReferences(pivotDefinition, fields);
  const slicerFilters = (pivotDefinition.slicers ?? []).map((slicer) => ({
      index: fields.indexOf(slicer.field),
      keys: new Set(slicer.selected.map(buildSlicerItemKey)),
    })),
    timelineFilters = (pivotDefinition.timelines ?? []).map((timeline) => ({
      ...timeline,
      index: fields.indexOf(timeline.field),
    }));
  return sourceRows.filter((row) => {
    const timelineDates = timelineFilters.map((timelineFilter) =>
      normalizeTimelineDate(
        row.values[timelineFilter.index],
        timelineFilter.dateSystem,
        timelineFilter.field,
        row.index,
      ),
    );
    return (
      slicerFilters.every((slicerFilter) =>
        slicerFilter.keys.has(
          buildSlicerItemKey(row.values[slicerFilter.index]),
        ),
      ) &&
      timelineFilters.every(
        (timelineRange, dateIndex) =>
          timelineDates[dateIndex] !== null &&
          timelineDates[dateIndex] >= timelineRange.from &&
          timelineDates[dateIndex] <= timelineRange.to,
      )
    );
  });
}
function buildTimelinePeriods(dates, timeline) {
  const periodsByLabel = new Map();
  for (const rawDate of dates) {
    const timelineDate = normalizeTimelineDate(
      rawDate,
      timeline.dateSystem,
      timeline.field,
      0,
    );
    if (!timelineDate) continue;
    const year = Number(timelineDate.slice(0, 4)),
      month = Number(timelineDate.slice(5, 7)),
      periodStartMonth =
        timeline.unit === "year"
          ? 1
          : timeline.unit === "quarter"
            ? Math.floor((month - 1) / 3) * 3 + 1
            : month,
      periodEndMonth =
        timeline.unit === "year"
          ? 12
          : timeline.unit === "quarter"
            ? periodStartMonth + 2
            : periodStartMonth,
      yearText = String(year).padStart(4, "0"),
      periodLabel =
        timeline.unit === "year"
          ? yearText
          : timeline.unit === "quarter"
            ? `${yearText} Q${Math.ceil(month / 3)}`
            : `${yearText}-${String(month).padStart(2, "0")}`,
      periodEndDate = new Date(0);
    (periodEndDate.setUTCFullYear(year, periodEndMonth, 0),
      periodEndDate.setUTCHours(0, 0, 0, 0),
      periodsByLabel.set(periodLabel, {
        label: periodLabel,
        from: `${yearText}-${String(periodStartMonth).padStart(2, "0")}-01`,
        to: periodEndDate.toISOString().slice(0, 10),
      }));
  }
  return [...periodsByLabel.values()].sort((leftPeriod, rightPeriod) =>
    leftPeriod.from.localeCompare(rightPeriod.from),
  );
}
const isPlainDataObject = (value) =>
    !!value &&
    typeof value == "object" &&
    !Array.isArray(value) &&
    (Object.getPrototypeOf(value) === Object.prototype ||
      Object.getPrototypeOf(value) === null),
  isBoundedString = (value, maxLength) =>
    typeof value == "string" && value.length > 0 && value.length <= maxLength;
function normalizeRowProvenance(workbook, normalizedWorkbook) {
  const sheets = workbook.sheets;
  if (
    workbook.rowOrigins === undefined &&
    !sheets.some((sheet) => sheet.rowOriginIds !== undefined)
  )
    return;
  if (
    !Array.isArray(workbook.rowOrigins) ||
    workbook.rowOrigins.length > workbookLimits.maxCells
  )
    throw Error(
      "Row provenance requires a registry of at most 100,000 source rows.",
    );
  const seenIds = new Set();
  normalizedWorkbook.rowOrigins = workbook.rowOrigins.map((rowOrigin) => {
    if (
      !isPlainDataObject(rowOrigin) ||
      !isBoundedString(rowOrigin.id, 100) ||
      seenIds.has(rowOrigin.id) ||
      !isBoundedString(rowOrigin.sourceId, 100) ||
      !isBoundedString(rowOrigin.sourceName, 255) ||
      !isBoundedString(rowOrigin.sourceSheet, 31) ||
      !["file", "snapshot"].includes(String(rowOrigin.kind)) ||
      !Number.isInteger(rowOrigin.originalRow) ||
      Number(rowOrigin.originalRow) < 1 ||
      Number(rowOrigin.originalRow) > workbookLimits.maxRows
    )
      throw Error("Invalid, duplicated or out-of-bounds original source row.");
    return (
      seenIds.add(rowOrigin.id),
      {
        id: rowOrigin.id,
        sourceId: rowOrigin.sourceId,
        sourceName: rowOrigin.sourceName,
        sourceSheet: rowOrigin.sourceSheet,
        originalRow: rowOrigin.originalRow,
        kind: rowOrigin.kind,
      }
    );
  });
  let associationCount = 0;
  for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
    const rowOriginIds = sheets[sheetIndex].rowOriginIds;
    if (rowOriginIds === undefined) continue;
    if (!isPlainDataObject(rowOriginIds))
      throw Error("Source row associations must be a plain index map.");
    const normalizedRowOriginIds = {};
    for (const [rowKey, originId] of Object.entries(rowOriginIds)) {
      if (
        !/^(0|[1-9]\d*)$/.test(rowKey) ||
        Number(rowKey) >= Number(sheets[sheetIndex].rowCount) ||
        typeof originId != "string" ||
        !seenIds.has(originId) ||
        ++associationCount > workbookLimits.maxCells
      )
        throw Error("Invalid or over-capacity source row association.");
      normalizedRowOriginIds[rowKey] = originId;
    }
    normalizedWorkbook.sheets[sheetIndex].rowOriginIds = normalizedRowOriginIds;
  }
}
function findRowOrigin(workbook, sheetId, rowIndex) {
  const rowOriginId = workbook.sheets.find((sheet) => sheet.id === sheetId)
      ?.rowOriginIds?.[rowIndex],
    matchedRowOrigin =
      rowOriginId &&
      workbook.rowOrigins?.find((rowOrigin) => rowOrigin.id === rowOriginId);
  return matchedRowOrigin ? { ...matchedRowOrigin } : null;
}
function validateRowOrigins(workbook, normalizedWorkbook) {
  if (!workbook.rowOrigins?.length) return;
  const normalizedRowOriginsById = new Map(
    (normalizedWorkbook.rowOrigins ?? []).map((normalizedOrigin) => [
      normalizedOrigin.id,
      normalizedOrigin,
    ]),
  );
  for (const rowOrigin of workbook.rowOrigins)
    if (
      !normalizedRowOriginsById.has(rowOrigin.id) ||
      [
        "id",
        "sourceId",
        "sourceName",
        "sourceSheet",
        "originalRow",
        "kind",
      ].some(
        (identityField) =>
          normalizedRowOriginsById.get(rowOrigin.id)[identityField] !==
          rowOrigin[identityField],
      )
    )
      throw Error(
        "Original source row identities cannot be changed or removed. Keep the source registry and change only current row associations.",
      );
}
const basicAggregateOperations = ["sum", "count", "average", "min", "max"],
  pivotAggregateOperations = [...basicAggregateOperations, "distinctCount"];

function assertUniqueTableNames(workbook) {
  const tableNames = new Set(),
    tableIds = new Set();
  for (const sheet of workbook.sheets)
    for (const table of sheet.tables ?? [])
      ((tableNames.has(table.name.toLowerCase()) || tableIds.has(table.id)) &&
        throwWorkbookError(
          "Table names and IDs must be unique across the workbook.",
        ),
        tableNames.add(table.name.toLowerCase()),
        tableIds.add(table.id));
}
function assertWithinCellBudget(workbook) {
  return (
    workbook.sheets.reduce(
      (total, sheet) => total + Object.keys(sheet.cells).length,
      0,
    ) > workbookLimits.maxCells &&
      throwWorkbookError(
        "The result would exceed the 100,000 populated-cell workbook budget.",
      ),
    workbook
  );
}
function findSourceFieldIndex(fields, fieldName) {
  const fieldIndex = fields.indexOf(fieldName);
  return fieldIndex >= 0
    ? fieldIndex
    : throwWorkbookError(
        `Source field “${fieldName}” was not found. Refresh the field selection.`,
      );
}
function buildUniqueSheetName(workbook, requestedName) {
  const baseName =
    requestedName
      .replace(/[\\/\[\]*?:]/g, " ")
      .trim()
      .slice(0, 31) || "Analysis";
  let candidateName = baseName,
    suffix = 2;
  for (
    ;
    workbook.sheets.some(
      (sheet) => sheet.name.toLowerCase() === candidateName.toLowerCase(),
    );
  )
    candidateName = `${baseName.slice(0, 25)} (${suffix++})`;
  return candidateName;
}
const computePivotRefreshFingerprint = (evidence) => {
  const evidenceJson = JSON.stringify(evidence);
  evidenceJson.length > 16e6 &&
    throwWorkbookError(
      "Pivot refresh evidence exceeds the 16-million-character fingerprint budget. Use a smaller source.",
    );
  const chunkHashes = [];
  for (let offset = 0; offset < evidenceJson.length; offset += 5e5)
    chunkHashes.push(
      hashPictureBytes(
        new Uint8Array(strToU8(evidenceJson.slice(offset, offset + 5e5))),
      ),
    );
  return hashPictureBytes(
    new Uint8Array(
      strToU8(JSON.stringify([evidenceJson.length, ...chunkHashes])),
    ),
  );
};
function findPivotDefinition(workbook, pivotId) {
  const sheet = workbook.sheets.find((sheet) =>
      sheet.pivots?.some((pivot) => pivot.id === pivotId),
    ),
    definition = sheet?.pivots?.find((pivot) => pivot.id === pivotId);
  return (
    (!sheet || !definition) &&
      throwWorkbookError("Pivot definition not found."),
    { owner: sheet, definition: definition }
  );
}
function buildPivotStatusRecord(workbook, definition, pivotRefreshState) {
  return {
    workbookId: workbook.id,
    revision: workbook.revision,
    pivotId: definition.id,
    name: definition.name,
    source: structuredClone(definition.source),
    sourceRevision: definition.sourceRevision,
    refreshedAt: definition.refresh?.at ?? null,
    sourceFingerprint: definition.refresh?.sourceFingerprint ?? null,
    status: pivotRefreshState.state,
    diagnostic: pivotRefreshState.reason,
    outputRange: definition.outputRange,
    rowFields: [...definition.rowFields],
    columnFields: [...definition.columnFields],
    values: structuredClone(definition.values),
    filters: structuredClone(definition.filters ?? []),
    slicers: structuredClone(definition.slicers ?? []),
    timelines: structuredClone(definition.timelines ?? []),
    sorts: structuredClone(definition.sorts ?? []),
    dateGroups: structuredClone(definition.dateGroups ?? []),
    calculatedFields: structuredClone(definition.calculatedFields ?? []),
    chart: structuredClone(definition.chart ?? null),
    population: definition.refresh
      ? {
          source: definition.refresh.sourceRows,
          included: definition.refresh.includedRows,
          excluded: definition.refresh.excludedRows,
        }
      : null,
    measures: structuredClone(definition.refresh?.measures ?? []),
    semantics:
      "Controls intersect. Slicer items use type-sensitive, case-insensitive equality; an empty selection includes no records. Hidden source rows are included. Percentages use filtered aggregate totals. Drill-through returns numerator records, including blank/nonnumeric measure inputs; it does not expand denominator populations.",
  };
}
const tableMutationLimits = {
  changedCells: 500,
  appendedRows: 100,
  formulaLength: 4096,
};
function validateCalculatedColumnFormula(refersTo, shouldSkipRefCheck = false) {
  (typeof refersTo != "string" ||
    refersTo.length > tableMutationLimits.formulaLength ||
    refersTo !== refersTo.trim() ||
    !refersTo.startsWith("=")) &&
    throwWorkbookError(
      "Calculated column formulas must begin with = and contain at most 4,096 characters, without outer whitespace.",
    );
  try {
    validateNameFormulaText(refersTo);
  } catch (error) {
    throwWorkbookError(
      `Calculated column formula: ${error.message.replace(/^XLSX defined names: /, "").replace(/ Import\/export stopped.*$/, "")}`,
    );
  }
  return (
    !shouldSkipRefCheck &&
      /#REF!/.test(
        stripFormulaStringLiterals(refersTo.replace(/"(?:[^"]|"")*"/g, '""')),
      ) &&
      throwWorkbookError(
        "Repair broken #REF! references before saving a calculated column rule.",
      ),
    refersTo
  );
}
const a1CellPattern = "\\$?[A-Za-z]{1,3}\\$?[1-9]\\d{0,6}",
  a1ReferenceAnchorPattern = new RegExp(
    `^(${a1CellPattern}(?::${a1CellPattern})?)(?![A-Za-z0-9_.$:#\\[])`,
  );
function parseExternalLinkRange(rangeText, maxCells = 5e3) {
  const parseCellAddress = (address) => {
      const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(address);
      if (!match)
        throw Error("External links require bounded A1 cells or rectangles.");
      let col = 0;
      for (const letter of match[1].toUpperCase())
        col = col * 26 + letter.charCodeAt(0) - 64;
      const row = Number(match[2]);
      if (col > 16384 || row > 1048576)
        throw Error("External reference exceeds worksheet bounds.");
      return { row: row, col: col };
    },
    [startAddress, endAddress = startAddress, extraPart] = rangeText.split(":");
  if (extraPart) throw Error("Invalid external range.");
  const startCell = parseCellAddress(startAddress),
    endCell = parseCellAddress(endAddress),
    cellAddresses = [],
    minRow = Math.min(startCell.row, endCell.row),
    maxRow = Math.max(startCell.row, endCell.row),
    minCol = Math.min(startCell.col, endCell.col),
    maxCol = Math.max(startCell.col, endCell.col);
  if ((maxRow - minRow + 1) * (maxCol - minCol + 1) > maxCells)
    throw Error(
      `External link ranges are limited to ${maxCells.toLocaleString()} cells.`,
    );
  for (let rowNumber = minRow; rowNumber <= maxRow; rowNumber++)
    for (let columnNumber = minCol; columnNumber <= maxCol; columnNumber++) {
      let remainingColumn = columnNumber,
        columnLetters = "";
      for (; remainingColumn;)
        ((columnLetters =
          String.fromCharCode(65 + ((remainingColumn - 1) % 26)) +
          columnLetters),
          (remainingColumn = Math.floor((remainingColumn - 1) / 26)));
      cellAddresses.push(columnLetters + rowNumber);
    }
  return cellAddresses;
}
function tokenizeExternalReferences(raw) {
  const tokens = [];
  for (let end = 0; end < raw.length;) {
    if (raw[end] === '"') {
      for (end++; end < raw.length;)
        if (raw[end++] === '"')
          if (raw[end] === '"') end++;
          else break;
      continue;
    }
    const start = end,
      quotedSheetMatch = /^'((?:[^']|'')*)'\s*!/.exec(raw.slice(end)),
      bracketedSheetMatch =
        /^\[([^\[\]\r\n]+)\]([A-Za-z_][A-Za-z0-9_.]*)!/.exec(raw.slice(end));
    let sheetReference,
      prefixLength = 0;
    if (
      (quotedSheetMatch
        ? ((sheetReference = quotedSheetMatch[1].replace(/''/g, "'")),
          (prefixLength = quotedSheetMatch[0].length))
        : bracketedSheetMatch &&
          ((sheetReference = `[${bracketedSheetMatch[1]}]${bracketedSheetMatch[2]}`),
          (prefixLength = bracketedSheetMatch[0].length)),
      sheetReference)
    ) {
      const sheetParts = /^(.*\[[^\[\]\r\n]+\])([^\[\]:\r\n]+)$/.exec(
          sheetReference,
        ),
        rangeMatch = a1ReferenceAnchorPattern.exec(
          raw.slice(end + prefixLength),
        );
      if (
        sheetParts &&
        rangeMatch &&
        (!end || !/[A-Za-z0-9_.\]]/.test(raw[end - 1]))
      )
        try {
          (parseExternalLinkRange(rangeMatch[1]),
            (end += prefixLength + rangeMatch[0].length),
            tokens.push({
              start: start,
              end: end,
              token: raw.slice(start, end),
              target: sheetParts[1],
              sheet: sheetParts[2],
              range: rangeMatch[1].toUpperCase(),
            }));
          continue;
        } catch {}
      end += prefixLength;
      continue;
    }
    if (raw[end] === "[") {
      let bracketDepth = 0;
      do
        (raw[end] === "[" && bracketDepth++,
          raw[end] === "]" && bracketDepth--,
          end++);
      while (end < raw.length && bracketDepth > 0);
    } else if (raw[end] === "'") {
      for (end++; end < raw.length;)
        if (raw[end++] === "'")
          if (raw[end] === "'") end++;
          else break;
    } else end++;
  }
  return tokens;
}
function rewriteExternalReferences(raw, replacer) {
  const tokens = tokenizeExternalReferences(raw);
  for (const token of tokens.reverse()) {
    const replacement = replacer(token);
    replacement !== undefined &&
      (raw = raw.slice(0, token.start) + replacement + raw.slice(token.end));
  }
  return raw;
}
const formatExternalSheetReference = (target, sheetName, range) =>
    `'${(target + sheetName).replace(/'/g, "''")}'!${range}`,
  workbookLinkLimits = Object.freeze({
    sources: 8,
    sheets: 32,
    cells: 5e3,
    formulas: 500,
    metadataCharacters: 2e6,
  }),
  workbookQualifierSchema = boundedTextSchema(2048).refine(
    (qualifier) => /^[^\[\]]*\[[^\[\]]+\]$/.test(qualifier),
    "Use an exact workbook qualifier such as [Source.xlsx].",
  ),
  tU = string().regex(/^[a-f0-9]{64}$/),
  linkedCellValueSchema = discriminatedUnion("type", [
    strictObject({ type: literal("number"), value: number().finite() }),
    strictObject({ type: literal("text"), value: string().max(32767) }),
    strictObject({ type: literal("boolean"), value: boolean() }),
    strictObject({ type: literal("blank"), value: _null() }),
    strictObject({
      type: literal("error"),
      value: string()
        .regex(/^#[A-Z0-9/?!_.]+$/)
        .max(30),
    }),
  ]),
  workbookLinkSourceSchema = strictObject({
    target: workbookQualifierSchema,
    sourceId: boundedTextSchema(100),
    sourceName: boundedTextSchema(200),
    sourceRevision: number().int().nonnegative(),
    sourceFingerprint: tU,
    resolvedAt: isoDateTimeSchemaFactory(),
    sheets: array(
      strictObject({
        id: boundedTextSchema(100),
        name: boundedTextSchema(31),
        cells: record(
          string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
          linkedCellValueSchema,
        ),
      }),
    )
      .min(1)
      .max(32),
  }),
  workbookLinkPackageSchema = strictObject({
    schemaVersion: literal(1),
    sources: array(workbookLinkSourceSchema).max(8),
    packageTargets: array(
      strictObject({
        key: workbookQualifierSchema,
        target: boundedTextSchema(4096),
        part: boundedTextSchema(1024),
        sourceSha256: tU,
      }),
    )
      .max(100)
      .optional(),
  }),
  validatedWorkbookLinkRegistry = new WeakSet();
function attachWorkbookLinkMetadata(workbook, destination) {
  if (workbook.workbookLinks === undefined) return;
  if (
    workbook.workbookLinks &&
    typeof workbook.workbookLinks == "object" &&
    validatedWorkbookLinkRegistry.has(workbook.workbookLinks)
  ) {
    destination.workbookLinks = workbook.workbookLinks;
    return;
  }
  if (
    JSON.stringify(workbook.workbookLinks).length >
    workbookLinkLimits.metadataCharacters
  )
    throw Error("Workbook link metadata exceeds 2 million characters.");
  const linkPackage = workbookLinkPackageSchema.parse(workbook.workbookLinks);
  if (
    new Set(linkPackage.sources.map((source) => source.target)).size !==
    linkPackage.sources.length
  )
    throw Error("Workbook link targets must be unique.");
  if (
    new Set(
      linkPackage.packageTargets?.map((packageTarget) => packageTarget.key),
    ).size !== (linkPackage.packageTargets?.length ?? 0)
  )
    throw Error("Package link identities must be unique.");
  let totalCells = 0,
    totalSheets = 0;
  for (const source of linkPackage.sources) {
    if (
      new Set(source.sheets.map((sheet) => sheet.name.toLowerCase())).size !==
      source.sheets.length
    )
      throw Error("Linked worksheet names must be unique.");
    for (const sheet of source.sheets) {
      totalSheets++;
      for (const cell of Object.keys(sheet.cells))
        (parseExternalLinkRange(cell), totalCells++);
    }
  }
  if (
    totalCells > workbookLinkLimits.cells ||
    totalSheets > workbookLinkLimits.sheets
  )
    throw Error("Workbook link cache exceeds its cell or sheet limit.");
  ((destination.workbookLinks = deepFreezeValue(linkPackage)),
    validatedWorkbookLinkRegistry.add(destination.workbookLinks));
}
function encodeWorkbookFingerprint(workbook) {
  const fingerprintBytes = new TextEncoder().encode(
    JSON.stringify([
      workbook.id,
      workbook.name,
      workbook.revision,
      workbook.calculation,
      workbook.datePolicy,
      workbook.definedNames,
      workbook.workbookLinks,
      workbook.sheets.map((sheet) => [
        sheet.id,
        sheet.name,
        sheet.rowCount,
        sheet.colCount,
        sheet.hidden,
        sheet.veryHidden,
        sheet.hiddenRows,
        sheet.hiddenCols,
        sheet.tables,
        sheet.merges,
        sheet.outlineGroups,
        Object.entries(sheet.cells).map(([cellAddress, cell]) => [
          cellAddress,
          cell.raw,
          !!cell.literal,
        ]),
      ]),
    ]),
  );
  return sha256HexDigest(fingerprintBytes);
}
function buildExternalLinkAliases(workbookInput) {
  const sheets = [],
    values = new Map(),
    aliasNameByKey = new Map(),
    takenSheetNames = new Set(
      workbookInput.sheets.map((sheet) => sheet.name.toLowerCase()),
    );
  for (const [sourceIndex, workbook] of (
    workbookInput.workbookLinks?.sources ?? []
  ).entries())
    for (const [sheetIndex, linkedSheet] of workbook.sheets.entries()) {
      let aliasName = `_LG_LINK_${sourceIndex}_${sheetIndex}`;
      for (; takenSheetNames.has(aliasName.toLowerCase());) aliasName += "_";
      takenSheetNames.add(aliasName.toLowerCase());
      let linkedSheetId = `linked:${sourceIndex}:${sheetIndex}`;
      for (; workbookInput.sheets.some((sheet) => sheet.id === linkedSheetId);)
        linkedSheetId += "_";
      const cells = {};
      for (const [cellAddress, linkedValue] of Object.entries(
        linkedSheet.cells,
      ))
        ((cells[cellAddress] = { raw: linkedValue.value, literal: true }),
          values.set(`${linkedSheetId}:${cellAddress}`, linkedValue));
      (sheets.push({
        id: linkedSheetId,
        name: aliasName,
        rowCount: 1048576,
        colCount: 16384,
        cells: cells,
      }),
        aliasNameByKey.set(
          JSON.stringify([workbook.target, linkedSheet.name.toLowerCase()]),
          aliasName,
        ));
    }
  return {
    sheets: sheets,
    values: values,
    rewrite: (raw) =>
      rewriteExternalReferences(raw, (token) => {
        const workbook = workbookInput.workbookLinks?.sources.find(
            (source) => source.target === token.target,
          ),
          matchedSheet = workbook?.sheets.find(
            (sheet) => sheet.name.toLowerCase() === token.sheet.toLowerCase(),
          );
        if (!(
          !matchedSheet ||
          !parseExternalLinkRange(token.range).every((rangeCell) =>
            Object.hasOwn(matchedSheet.cells, rangeCell),
          )
        ))
          return `'${aliasNameByKey.get(JSON.stringify([token.target, token.sheet.toLowerCase()]))}'!${token.range}`;
      }),
  };
}
const databaseFunctionNames = [
    "DSUM",
    "DAVERAGE",
    "DCOUNT",
    "DCOUNTA",
    "DMAX",
    "DMIN",
    "DPRODUCT",
    "DSTDEV",
    "DSTDEVP",
    "DVAR",
    "DVARP",
    "DGET",
  ],
  notDatabaseFunctionMarker = Symbol("not a database function"),
  databaseFunctionLimits = {
    databaseCells: 2e4,
    criteriaCells: 1e3,
    predicateComparisons: 1e5,
    textComparisonSteps: 5e6,
    criterionCharacters: 255,
    headerCharacters: 255,
    formulaEvaluations: 1e4,
  };
function evaluateDatabaseFunction(functionName, args, context) {
  if (!databaseFunctionNames.includes(functionName))
    return notDatabaseFunctionMarker;
  const raiseError = (errorCode = "#VALUE!") => {
    throw context.issue(errorCode);
  };
  if (args.length !== 3 || args[0] == null || args[2] == null)
    return raiseError();
  const readTable = (reference, cellLimit) => {
      const retrieved = context.retrieve(reference);
      if (context.isIssue(retrieved)) throw retrieved;
      if (!Array.isArray(retrieved) || retrieved.length < 2)
        return raiseError();
      if (retrieved.length > cellLimit) return raiseError("#LIMIT!");
      if (!Array.isArray(retrieved[0])) return raiseError();
      const table = retrieved,
        columnCount = table[0].length;
      if (!columnCount) return raiseError();
      if (table.length * columnCount > cellLimit) return raiseError("#LIMIT!");
      for (let rowIndex = 0; rowIndex < table.length; rowIndex++)
        if (
          !Array.isArray(table[rowIndex]) ||
          table[rowIndex].length !== columnCount
        )
          return raiseError();
      for (const row of table)
        for (const cellValue of row)
          if (context.isIssue(cellValue)) {
            if (
              [
                "#LIMIT!",
                "#CYCLE!",
                "#BLOCKED!",
                "#TIMEOUT!",
                "#ITER!",
                "#ERROR!",
              ].includes(cellValue.error)
            )
              throw cellValue;
          } else {
            if (
              cellValue != null &&
              !["number", "string", "boolean"].includes(typeof cellValue)
            )
              return raiseError();
            if (typeof cellValue == "number" && !Number.isFinite(cellValue))
              return raiseError("#NUM!");
          }
      return table;
    },
    databaseTable = readTable(args[0], databaseFunctionLimits.databaseCells),
    criteriaTable = readTable(args[2], databaseFunctionLimits.criteriaCells);
  if (
    (databaseTable.length - 1) *
      (criteriaTable.length - 1) *
      criteriaTable[0].length >
    databaseFunctionLimits.predicateComparisons
  )
    return raiseError("#LIMIT!");
  const normalizeHeader = (header) => {
      if (context.isIssue(header)) throw header;
      return typeof header != "string" ||
        !header.trim() ||
        header.length > databaseFunctionLimits.headerCharacters
        ? raiseError()
        : header.toLowerCase();
    },
    normalizedHeaders = databaseTable[0].map(normalizeHeader);
  if (new Set(normalizedHeaders).size !== normalizedHeaders.length)
    return raiseError();
  const criteriaColumnIndexes = criteriaTable[0].map((criteriaHeader) => {
      if (context.isIssue(criteriaHeader)) throw criteriaHeader;
      return criteriaHeader == null || criteriaHeader === ""
        ? -1
        : normalizedHeaders.indexOf(normalizeHeader(criteriaHeader));
    }),
    isCountFunction = functionName === "DCOUNT" || functionName === "DCOUNTA";
  let fieldIndex;
  if (args[1] == null) {
    if (!isCountFunction) return raiseError();
  } else {
    const fieldArgument = context.retrieve(args[1]);
    if (context.isIssue(fieldArgument)) throw fieldArgument;
    if (typeof fieldArgument == "number") {
      if (
        !Number.isInteger(fieldArgument) ||
        fieldArgument < 1 ||
        fieldArgument > normalizedHeaders.length
      )
        return raiseError();
      fieldIndex = fieldArgument - 1;
    } else if (typeof fieldArgument == "string" && fieldArgument.length) {
      if (
        ((fieldIndex = normalizedHeaders.indexOf(fieldArgument.toLowerCase())),
        fieldIndex < 0)
      )
        return raiseError();
    } else return raiseError();
  }
  let comparisonSteps = 0;
  const addComparisonSteps = (stepCount) => {
      ((comparisonSteps += stepCount),
        comparisonSteps > databaseFunctionLimits.textComparisonSteps &&
          raiseError("#LIMIT!"));
    },
    compareValues = (left, right, operator) =>
      operator === "<>"
        ? left !== right
        : operator === "<"
          ? left < right
          : operator === ">"
            ? left > right
            : operator === "<="
              ? left <= right
              : operator === ">="
                ? left >= right
                : left === right,
    buildCriterionPredicate = (criterion) => {
      if (context.isIssue(criterion)) throw criterion;
      if (criterion == null || criterion === "") return null;
      if (typeof criterion == "number" || typeof criterion == "boolean")
        return (candidateValue) => candidateValue === criterion;
      if (typeof criterion != "string") return raiseError();
      if (criterion.length > databaseFunctionLimits.criterionCharacters)
        return raiseError("#LIMIT!");
      const criterionMatch = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(criterion),
        criterionOperator = criterionMatch[1] ?? "=",
        criterionOperand = criterionMatch[2];
      if (criterionMatch[1] && /^[<>=]/.test(criterionOperand))
        return raiseError();
      if (!criterionOperand && criterionMatch[1])
        return criterionOperator !== "=" && criterionOperator !== "<>"
          ? raiseError()
          : (blankTestCell) =>
              criterionOperator === "="
                ? blankTestCell == null || blankTestCell === ""
                : blankTestCell != null && blankTestCell !== "";
      if (
        /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(
          criterionOperand.trim(),
        )
      ) {
        const numericCriterion = Number(criterionOperand);
        return Number.isFinite(numericCriterion)
          ? (candidateItem) =>
              typeof candidateItem == "number"
                ? compareValues(
                    candidateItem,
                    numericCriterion,
                    criterionOperator,
                  )
                : criterionOperator === "<>"
          : raiseError("#NUM!");
      }
      if (/^(TRUE|FALSE)$/i.test(criterionOperand)) {
        if (criterionOperator !== "=" && criterionOperator !== "<>")
          return raiseError();
        const booleanCriterion = criterionOperand.toUpperCase() === "TRUE";
        return (candidateCell) =>
          criterionOperator === "="
            ? candidateCell === booleanCriterion
            : candidateCell !== booleanCriterion;
      }
      const lowerCaseOperand = criterionOperand.toLowerCase(),
        patternTokens = [];
      let hasWildcard = false;
      for (
        let patternIndex = 0;
        patternIndex < lowerCaseOperand.length;
        patternIndex++
      ) {
        const char = lowerCaseOperand[patternIndex];
        if (char === "~") {
          const nextPatternChar = lowerCaseOperand[patternIndex + 1];
          nextPatternChar && ["*", "?", "~"].includes(nextPatternChar)
            ? (patternTokens.push({ kind: "literal", char: nextPatternChar }),
              patternIndex++,
              (hasWildcard = true))
            : patternTokens.push({ kind: "literal", char: char });
        } else
          char === "*" || char === "?"
            ? ((hasWildcard = true),
              (char === "?" || patternTokens.at(-1)?.kind !== "many") &&
                patternTokens.push({ kind: char === "*" ? "many" : "one" }))
            : patternTokens.push({ kind: "literal", char: char });
      }
      return ["<", ">", "<=", ">="].includes(criterionOperator)
        ? hasWildcard
          ? raiseError()
          : (candidateText) =>
              typeof candidateText != "string"
                ? false
                : (addComparisonSteps(
                    candidateText.length + lowerCaseOperand.length,
                  ),
                  compareValues(
                    candidateText.toLowerCase(),
                    lowerCaseOperand,
                    criterionOperator,
                  ))
        : (!criterionMatch[1] &&
            !hasWildcard &&
            patternTokens.push({ kind: "many" }),
          (candidateString) => {
            if (typeof candidateString != "string")
              return criterionOperator === "<>";
            addComparisonSteps(
              (candidateString.length + 1) * (patternTokens.length + 1),
            );
            const lowerCaseCandidate = candidateString.toLowerCase();
            let previousMatchRow = new Uint8Array(
              lowerCaseCandidate.length + 1,
            );
            previousMatchRow[0] = 1;
            for (const patternToken of patternTokens) {
              const nextMatchRow = new Uint8Array(
                lowerCaseCandidate.length + 1,
              );
              if (patternToken.kind === "many") {
                nextMatchRow[0] = previousMatchRow[0];
                for (
                  let candidateIndex = 1;
                  candidateIndex <= lowerCaseCandidate.length;
                  candidateIndex++
                )
                  nextMatchRow[candidateIndex] =
                    previousMatchRow[candidateIndex] ||
                    nextMatchRow[candidateIndex - 1];
              } else
                for (
                  let candidateOffset = 1;
                  candidateOffset <= lowerCaseCandidate.length;
                  candidateOffset++
                )
                  nextMatchRow[candidateOffset] =
                    previousMatchRow[candidateOffset - 1] &&
                    (patternToken.kind === "one" ||
                      patternToken.char ===
                        lowerCaseCandidate[candidateOffset - 1])
                      ? 1
                      : 0;
              previousMatchRow = nextMatchRow;
            }
            const isPatternMatch =
              !!previousMatchRow[lowerCaseCandidate.length];
            return criterionOperator === "<>"
              ? !isPatternMatch
              : isPatternMatch;
          });
    };
  let formulaEvaluationCount = 0;
  const criterionFormulas = criteriaTable
    .slice(1)
    .map((criteriaRow, criteriaRowIndex) =>
      criteriaRow.map((criteriaCell, criteriaColumnIndex) =>
        criteriaColumnIndexes[criteriaColumnIndex] < 0
          ? context.criterionFormula?.(
              args[2],
              criteriaRowIndex + 1,
              criteriaColumnIndex,
            )
          : undefined,
      ),
    );
  criteriaColumnIndexes.forEach(
    (databaseColumnIndex, criteriaColumnPosition) => {
      databaseColumnIndex < 0 &&
        !criterionFormulas.some(
          (criterionFormulaRow) => criterionFormulaRow[criteriaColumnPosition],
        ) &&
        raiseError();
    },
  );
  const criteriaPredicates = criteriaTable
    .slice(1)
    .map((criteriaRowValues, predicateRowIndex) =>
      criteriaRowValues.map((criteriaValue, predicateColumnIndex) => {
        if (criteriaColumnIndexes[predicateColumnIndex] >= 0)
          return buildCriterionPredicate(criteriaValue);
        if (
          !criterionFormulas[predicateRowIndex][predicateColumnIndex] &&
          !(criteriaValue == null || criteriaValue === "")
        )
          return raiseError();
      }),
    );
  criterionFormulas.flat().filter(Boolean).length * (databaseTable.length - 1) >
    databaseFunctionLimits.formulaEvaluations && raiseError("#LIMIT!");
  let matchedRecordCount = 0,
    numericValueCount = 0,
    nonBlankCount = 0,
    valueSum = 0,
    valueProduct = 1,
    minimumValue = 1 / 0,
    maximumValue = -1 / 0,
    runningMean = 0,
    squaredDeviationSum = 0,
    lastFieldValue;
  for (const [recordIndex, record] of databaseTable.slice(1).entries()) {
    let isRecordMatched = false;
    for (const [
      alternativeIndex,
      alternativePredicates,
    ] of criteriaPredicates.entries()) {
      let isAlternativeMatch = true;
      for (
        let predicateIndex = 0;
        predicateIndex < alternativePredicates.length;
        predicateIndex++
      ) {
        const criterionFormula =
          criterionFormulas[alternativeIndex][predicateIndex];
        if (criterionFormula) {
          ++formulaEvaluationCount >
            databaseFunctionLimits.formulaEvaluations && raiseError("#LIMIT!");
          const formulaResult = criterionFormula(recordIndex);
          if (context.isIssue(formulaResult)) throw formulaResult;
          (typeof formulaResult != "boolean" && raiseError(),
            formulaResult || (isAlternativeMatch = false));
          continue;
        }
        const predicate = alternativePredicates[predicateIndex];
        if (!predicate) continue;
        const recordCell = record[criteriaColumnIndexes[predicateIndex]];
        if (context.isIssue(recordCell)) throw recordCell;
        predicate(recordCell) || (isAlternativeMatch = false);
      }
      isAlternativeMatch && (isRecordMatched = true);
    }
    if (!isRecordMatched || (matchedRecordCount++, fieldIndex === undefined))
      continue;
    const fieldValue = record[fieldIndex];
    if (
      ((lastFieldValue = fieldValue),
      fieldValue != null && nonBlankCount++,
      context.isIssue(fieldValue))
    ) {
      if (!["DCOUNT", "DCOUNTA", "DGET"].includes(functionName))
        throw fieldValue;
      continue;
    }
    if (typeof fieldValue != "number") continue;
    (numericValueCount++,
      (valueSum += fieldValue),
      (valueProduct *= fieldValue),
      (minimumValue = Math.min(minimumValue, fieldValue)),
      (maximumValue = Math.max(maximumValue, fieldValue)));
    const valueDelta = fieldValue - runningMean;
    ((runningMean += valueDelta / numericValueCount),
      (squaredDeviationSum += valueDelta * (fieldValue - runningMean)));
  }
  if (functionName === "DGET") {
    if (matchedRecordCount !== 1)
      return raiseError(matchedRecordCount ? "#NUM!" : "#VALUE!");
    if (context.isIssue(lastFieldValue)) throw lastFieldValue;
    return lastFieldValue ?? 0;
  }
  if (functionName === "DCOUNT")
    return fieldIndex === undefined ? matchedRecordCount : numericValueCount;
  if (functionName === "DCOUNTA")
    return fieldIndex === undefined ? matchedRecordCount : nonBlankCount;
  const toFiniteResult = (result) =>
    Number.isFinite(result) ? result : raiseError("#NUM!");
  if (functionName === "DSUM") return toFiniteResult(valueSum);
  if (functionName === "DMAX") return numericValueCount ? maximumValue : 0;
  if (functionName === "DMIN") return numericValueCount ? minimumValue : 0;
  if (functionName === "DPRODUCT")
    return numericValueCount ? toFiniteResult(valueProduct) : 0;
  if (functionName === "DAVERAGE")
    return numericValueCount
      ? toFiniteResult(valueSum / numericValueCount)
      : raiseError("#DIV/0!");
  const isSampleStatistic =
    functionName === "DSTDEV" || functionName === "DVAR";
  if (numericValueCount < (isSampleStatistic ? 2 : 1))
    return raiseError("#DIV/0!");
  const variance = Math.max(
    0,
    squaredDeviationSum / (numericValueCount - (isSampleStatistic ? 1 : 0)),
  );
  return toFiniteResult(
    functionName === "DSTDEV" || functionName === "DSTDEVP"
      ? Math.sqrt(variance)
      : variance,
  );
}
const referenceSensitiveFunctionNames = [
    "SUBTOTAL",
    "AGGREGATE",
    "ISFORMULA",
    "FORMULATEXT",
  ],
  notReferenceFunctionMarker = Symbol("not-reference-function"),
  isFormulaCellValue = (cell) =>
    typeof cell?.raw == "string" && cell.raw.startsWith("=") && !cell.literal;
function containsSubtotalCall(formulaText) {
  return /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?SUBTOTAL\s*\(/i.test(
    formulaText.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, ""),
  );
}
function containsSubtotalOrAggregateCall(formulaText) {
  return /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?(?:SUBTOTAL|AGGREGATE)\s*\(/i.test(
    formulaText.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, ""),
  );
}
const aggregateFunctionNames = [
    "SUM",
    "AVERAGE",
    "MIN",
    "MAX",
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "SUMIF",
    "SUMIFS",
    "COUNTIF",
    "COUNTIFS",
    "AVERAGEIF",
    "AVERAGEIFS",
    "IF",
    "IFERROR",
    "IFNA",
    "AND",
    "OR",
    "NOT",
    "TRUE",
    "FALSE",
    "ABS",
    "ROUND",
    "ROUNDDOWN",
    "ROUNDUP",
    "INT",
    "MOD",
    "SIGN",
    "SQRT",
    "POWER",
    "SUMPRODUCT",
    "VLOOKUP",
    "HLOOKUP",
    "INDEX",
    "MATCH",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "TRIM",
    "CLEAN",
    "UPPER",
    "LOWER",
    "CONCATENATE",
    "TEXT",
    "VALUE",
    "SUBSTITUTE",
    "REPLACE",
    "FIND",
    "SEARCH",
    "DATE",
    "DAY",
    "MONTH",
    "YEAR",
    "TODAY",
    "EDATE",
    "EOMONTH",
    "DAYS",
    "NETWORKDAYS",
    "WORKDAY",
    "YEARFRAC",
    "ISBLANK",
    "ISNUMBER",
    "ISTEXT",
    "ISERROR",
    "ISNA",
    "ROW",
    "COLUMN",
    "ROWS",
    "COLUMNS",
    "LARGE",
    "SMALL",
    "MEDIAN",
    "STDEV.S",
    "VAR.S",
  ],
  builtinFunctionNames = [
    ...new Set([
      ...aggregateFunctionNames,
      ...supportedFunctionNames,
      ...modernFunctionNames,
      ...supportedFinancialFunctionNames,
      ...referenceSensitiveFunctionNames,
      ...databaseFunctionNames,
      ...statisticalFunctionNames,
      ...ownedLookupFunctionNames,
      ...localBindingFunctionNames,
      ...specialistFunctionNames,
      ...XL,
      ...significanceTestFunctionNames,
    ]),
  ].sort(),
  builtinFunctionNameSet = new Set(builtinFunctionNames),
  FormulaErrorClass = FormulaParser.FormulaError,
  createFormulaError = (e) => new FormulaErrorClass(e),
  formulaErrorFactory = createFormulaError,
  isFormulaError = (value) => value instanceof FormulaErrorClass,
  findFormulaErrorArg = (argValues) =>
    argValues.flatMap(flattenNestedValues).find(isFormulaError);
function findScopedDefinedName(nameReference, workbook, currentSheetId) {
  if (
    nameReference.cell ||
    (nameReference.call &&
      !nameReference.qualifier &&
      builtinFunctionNames.includes(normalizeFunctionName(nameReference.name)))
  )
    return;
  const scopeSheetId = nameReference.qualifier
    ? workbook.sheets.find(
        (sheet) =>
          sheet.name.toLowerCase() === nameReference.qualifier.toLowerCase(),
      )?.id
    : currentSheetId;
  if (scopeSheetId)
    return (
      workbook.definedNames?.find(
        (definedName) =>
          definedName.scopeSheetId === scopeSheetId &&
          definedName.name.toLowerCase() === nameReference.name.toLowerCase(),
      ) ??
      workbook.definedNames?.find(
        (definedName) =>
          !definedName.scopeSheetId &&
          definedName.name.toLowerCase() === nameReference.name.toLowerCase(),
      )
    );
}
function resolveDefinedNameValue(
  formulaText,
  workbook,
  currentSheetId,
  activeNameKeys = [],
  canReadSheet,
) {
  if (activeNameKeys.length > 32) throw new Error("#LIMIT!");
  const resolvedText = replaceTextSpans(
    formulaText,
    m5(formulaText),
    (textSpan) => {
      if (textSpan.call) return textSpan.text;
      const scopedDefinedName = findScopedDefinedName(
        textSpan,
        workbook,
        currentSheetId,
      );
      if (!scopedDefinedName || parseFormulaTree(scopedDefinedName.refersTo))
        return textSpan.text;
      const scopeKey = definedNameScopeKey(scopedDefinedName);
      if (activeNameKeys.includes(scopeKey)) throw new Error("#CYCLE!");
      const evaluationSheetId =
        scopedDefinedName.scopeSheetId ??
        (textSpan.qualifier
          ? workbook.sheets.find(
              (sheet) =>
                sheet.name.toLowerCase() === textSpan.qualifier.toLowerCase(),
            ).id
          : currentSheetId);
      if (canReadSheet && !canReadSheet(evaluationSheetId))
        throw new Error("#BLOCKED!");
      let definedNameValue = resolveDefinedNameValue(
        scopedDefinedName.refersTo.slice(1),
        workbook,
        evaluationSheetId,
        [...activeNameKeys, scopeKey],
        canReadSheet,
      );
      if (evaluationSheetId !== currentSheetId) {
        const sheet = workbook.sheets.find(
            (sheet) => sheet.id === evaluationSheetId,
          ),
          valueSpans = m5(definedNameValue);
        definedNameValue = replaceTextSpans(
          definedNameValue,
          valueSpans,
          (valueSpan) => {
            const spanIndex = valueSpans.indexOf(valueSpan),
              previousSpan = valueSpans[spanIndex - 1],
              isRangeEnd =
                previousSpan?.cell &&
                definedNameValue
                  .slice(previousSpan.end, valueSpan.start)
                  .trim() === ":";
            return valueSpan.cell && !valueSpan.qualifier && !isRangeEnd
              ? formatQuotedSheetPrefix(sheet.name) + valueSpan.text
              : valueSpan.text;
          },
        );
      }
      return `(${definedNameValue})`;
    },
  );
  if (resolvedText.length > 16384) throw new Error("#LIMIT!");
  return resolvedText;
}

export {
  databaseFunctionNames,
  formatExternalSheetReference,
  assertDrawingPreviewCurrent,
  buildSlicerItemKey,
  bucketPivotDateValue,
  encodeWorkbookFingerprint,
  F2,
  pivotAggregateOperations,
  builtinFunctionNames,
  containsSubtotalCall,
  tableMutationLimits,
  findFormulaErrorArg,
  builtinFunctionNameSet,
  attachWorkbookLinkMetadata,
  computePivotTable,
  inspectPictureDataUrl,
  evaluateDatabaseFunction,
  findPivotDefinition,
  parseExternalLinkRange,
  findRowOrigin,
  assertUniqueTableNames,
  computePivotRefreshFingerprint,
  workbookQualifierSchema,
  findSourceFieldIndex,
  validatePictureEvidence,
  containsSubtotalOrAggregateCall,
  normalizeRowProvenance,
  basicAggregateOperations,
  findScopedDefinedName,
  tokenizeExternalReferences,
  createPictureEvidence,
  referenceSensitiveFunctionNames,
  normalizePivotControls,
  filterPivotSourceRows,
  formulaErrorFactory,
  createFormulaError,
  notDatabaseFunctionMarker,
  pivotShowValuesAsModes,
  notReferenceFunctionMarker,
  buildTimelinePeriods,
  resolveDefinedNameValue,
  buildUniqueSheetName,
  buildPivotStatusRecord,
  buildExternalLinkAliases,
  validateRowOrigins,
  validatePivotRefreshEvidence,
  validatePivotFieldReferences,
  workbookLinkLimits,
  validateCalculatedColumnFormula,
  isFormulaError,
  assertWithinCellBudget,
  rewriteExternalReferences,
  isFormulaCellValue,
};
export { throwErrorMessage, isPlainDataObject, isFiniteNumber };

/* Also used by the workers. */
export {
  FormulaErrorClass,
  a1CellPattern,
  a1ReferenceAnchorPattern,
  aggregateFunctionNames,
  databaseFunctionLimits,
  linkedCellValueSchema,
  tU,
  workbookLinkPackageSchema,
  workbookLinkSourceSchema,
};

/* Also used by the workers. */
export {
  assertPlainTextLabel,
  isBoundedString,
  isIsoDateString,
  requireBoundedList,
  requirePivotFieldName,
};
