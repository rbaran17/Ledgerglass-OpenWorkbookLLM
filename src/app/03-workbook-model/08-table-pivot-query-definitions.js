import {
  isNonArrayObject,
  failWithMessage,
  throwWorkbookError,
} from "./01-cell-styles-and-themes.js";
import {
  isPlainDataObject,
  isFiniteNumber,
} from "./04-pivot-tables-and-pictures.js";
import { isNullOrEmptyValue } from "./03-defined-names.js";
import { workbookLimits } from "../02-text-and-financial-functions.js";
import { cn } from "./01-cell-styles-and-themes.js";
import {
  validateDrawingColor,
  maxPictureHashBytes,
  validateChartSettings,
  Kne,
  validateDrawingNumber,
  throwXlsxDefinedNameError,
  Yne,
  validateXmlSafeName,
  validateNameSheetReferences,
  validateDefinedName,
  insertDefinedNamesIntoWorkbookXml,
  validateTableColumnType,
  supportedChartKinds,
  validateNameFormulaText,
  validateDrawingText,
} from "./03-defined-names.js";
import {
  assertDrawingPreviewCurrent,
  pivotAggregateOperations,
  inspectPictureDataUrl,
  computePivotRefreshFingerprint,
  findSourceFieldIndex,
  validatePictureEvidence,
  basicAggregateOperations,
  normalizePivotControls,
  pivotShowValuesAsModes,
  validatePivotRefreshEvidence,
  validateCalculatedColumnFormula,
} from "./04-pivot-tables-and-pictures.js";
import { columnIndexToLetters } from "./06-data-model-and-analysis.js";
import {
  _t,
  aP,
  parseRangeAddress,
  oP,
} from "./07-conditional-rules-worksheet-features.js";

function serializeDefinedNames(workbookXml, workbookInput) {
  const names = validateDefinedName(workbookInput.definedNames, workbookInput);
  oP(names);
  const { builtinXml } = aP(workbookXml);
  let builtinNameXml = builtinXml;
  for (const definedName of names) {
    if (definedName.name.toLowerCase().startsWith("_xlnm."))
      return throwXlsxDefinedNameError(
        `Reserved built-in name "${definedName.name}" cannot be supplied as a user name.`,
      );
    (validateNameFormulaText(definedName.refersTo),
      validateNameSheetReferences(definedName.refersTo, workbookInput));
    const scopeSheetIndex = definedName.scopeSheetId
      ? workbookInput.sheets.findIndex(
          (sheet) => sheet.id === definedName.scopeSheetId,
        )
      : undefined;
    builtinNameXml += `<definedName name="${validateXmlSafeName(definedName.name, true)}"${scopeSheetIndex !== undefined ? ` localSheetId="${scopeSheetIndex}"` : ""}${definedName.comment !== undefined ? ` comment="${validateXmlSafeName(definedName.comment, true)}"` : ""}>${validateXmlSafeName(definedName.refersTo.slice(1))}</definedName>`;
  }
  return insertDefinedNamesIntoWorkbookXml(workbookXml, builtinNameXml);
}
function inspectChartSource(sheet, object, getCellValue) {
  (object.kind !== "chart" || !supportedChartKinds.includes(object.chart)) &&
    failWithMessage("Choose a supported chart type.");
  const chart = object.chart,
    range = object.range ?? "",
    rangeAddress = parseRangeAddress(range),
    chartSettings = validateChartSettings(object.chartOptions);
  (rangeAddress.lastRow >= sheet.rowCount ||
    rangeAddress.lastCol >= sheet.colCount ||
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      500) &&
    failWithMessage(
      "Select at most 500 cells inside this worksheet for a chart.",
    );
  const firstDataRow = rangeAddress.firstRow + (object.headers ? 1 : 0),
    firstValueCol =
      rangeAddress.firstCol +
      (rangeAddress.lastCol > rangeAddress.firstCol ? 1 : 0);
  rangeAddress.lastCol - firstValueCol + 1 > 20 &&
    failWithMessage(
      "Chart inspection supports at most 20 value columns; narrow the source range.",
    );
  const readCell = (rowIndex, columnIndex) => {
      const address = `${columnIndexToLetters(columnIndex)}${rowIndex + 1}`,
        cell = sheet.cells[address] ?? { raw: null };
      return {
        address: address,
        raw: cell.raw,
        value: getCellValue(sheet.id, address),
        ...(cell.literal ? { literal: true } : {}),
        ...(cell.note ? { note: cell.note } : {}),
      };
    },
    series = Array.from(
      { length: rangeAddress.lastCol - firstValueCol + 1 },
      (z, seriesIndex) => {
        const sourceColumnIndex = firstValueCol + seriesIndex,
          seriesOption = chartSettings.series?.find(
            (sery) => sery.column === sourceColumnIndex + 1,
          ),
          headerLabel = object.headers
            ? readCell(rangeAddress.firstRow, sourceColumnIndex).value
            : columnIndexToLetters(sourceColumnIndex);
        return {
          column: sourceColumnIndex + 1,
          name:
            seriesOption?.name ??
            String(headerLabel ?? columnIndexToLetters(sourceColumnIndex)),
          kind:
            seriesOption?.kind ??
            (chart === "combo" && seriesIndex > 0 ? "line" : "column"),
          included: chartSettings.series === undefined || !!seriesOption,
        };
      },
    );
  (chartSettings.series?.some(
    (sery) =>
      sery.column < firstValueCol + 1 || sery.column > rangeAddress.lastCol + 1,
  ) &&
    failWithMessage(
      "A selected chart series no longer exists in its source range. Re-select series.",
    ),
    chartSettings.includedRows?.some(
      (includedRow) =>
        includedRow < firstDataRow + 1 ||
        includedRow > rangeAddress.lastRow + 1,
    ) &&
      failWithMessage(
        "A chart filter row is outside its source range. Reset the row filter.",
      ));
  const hiddenRowRanges = cn(sheet, "row"),
    isHiddenRow = (rowIndex) =>
      hiddenRowRanges.some(
        (hiddenRange) =>
          rowIndex >= hiddenRange.start && rowIndex <= hiddenRange.end,
      ),
    points = Array.from(
      { length: Math.max(0, rangeAddress.lastRow - firstDataRow + 1) },
      (z, rowOffset) => {
        const sourceRowIndex = firstDataRow + rowOffset,
          category = readCell(sourceRowIndex, rangeAddress.firstCol);
        return {
          row: sourceRowIndex + 1,
          label:
            rangeAddress.firstCol < firstValueCol
              ? String(category.value ?? "")
              : category.address,
          category: category,
          included:
            chartSettings.includedRows === undefined ||
            chartSettings.includedRows.includes(sourceRowIndex + 1),
          hidden: isHiddenRow(sourceRowIndex),
          values: series.map((sery) =>
            readCell(sourceRowIndex, sery.column - 1),
          ),
        };
      },
    ),
    selectedSeriesIndexes = series
      .map((sery, seriesPosition) => (sery.included ? seriesPosition : -1))
      .filter((selectedIndex) => selectedIndex >= 0),
    includedPoints = points.filter((point) => point.included),
    diagnostics = [],
    addDiagnostic = (message, address) =>
      diagnostics.push({
        severity: "error",
        message: message,
        ...(address ? { address: address } : {}),
      });
  (points.length ||
    addDiagnostic("Select at least one data row below the chart header."),
    includedPoints.length ||
      addDiagnostic("The chart filter includes no source rows."),
    selectedSeriesIndexes.length ||
      addDiagnostic("Select at least one value series."),
    ["pie", "donut", "waterfall", "histogram", "sparkline"].includes(chart) &&
      selectedSeriesIndexes.length !== 1 &&
      addDiagnostic(
        `${chart} requires exactly one selected series. Other source series remain available in inspection.`,
      ));
  const isEmptyValue = (value) => value === null || value === "";
  for (let seriesIndex = 0; seriesIndex < series.length; seriesIndex++) {
    let plottedValueCount = 0;
    for (const point of points) {
      const valueCell = point.values[seriesIndex];
      if (isEmptyValue(valueCell.value)) continue;
      const isNumeric =
          typeof valueCell.value == "number" &&
          Number.isFinite(valueCell.value),
        isPlotted = point.included && series[seriesIndex].included;
      if (!isNumeric) {
        diagnostics.push({
          severity: isPlotted ? "error" : "warning",
          message: `${valueCell.address}: ${series[seriesIndex].name} contains a nonnumeric or invalid calculated value${isPlotted ? "" : " outside the plotted selection"}.`,
          address: valueCell.address,
        });
        continue;
      }
      isPlotted &&
        (plottedValueCount++,
        chartSettings.yScale === "log" &&
          Number(valueCell.value) <= 0 &&
          addDiagnostic(
            `${valueCell.address}: log-axis values must be positive.`,
            valueCell.address,
          ),
        ["pie", "donut"].includes(chart) &&
          Number(valueCell.value) < 0 &&
          addDiagnostic(
            `${valueCell.address}: pie/donut values cannot be negative.`,
            valueCell.address,
          ));
    }
    series[seriesIndex].included &&
      !plottedValueCount &&
      addDiagnostic(
        `${series[seriesIndex].name}: the selected series is empty or has no numeric values.`,
      );
  }
  if (chart === "scatter") {
    rangeAddress.firstCol === firstValueCol &&
      addDiagnostic(
        "Scatter requires a numeric X column followed by at least one Y column.",
      );
    for (const includedPoint of includedPoints) {
      if (
        !includedPoint.values.some(
          (value, seriesIndex) =>
            series[seriesIndex].included && !isEmptyValue(value.value),
        )
      )
        continue;
      const value = includedPoint.category.value;
      typeof value != "number" || !Number.isFinite(value)
        ? addDiagnostic(
            `${includedPoint.category.address}: scatter X values must be numeric.`,
            includedPoint.category.address,
          )
        : chartSettings.xScale === "log" &&
          value <= 0 &&
          addDiagnostic(
            `${includedPoint.category.address}: log X-axis values must be positive.`,
            includedPoint.category.address,
          );
    }
    const xValues = includedPoints
      .map((point) => point.category.value)
      .filter((xValue) => typeof xValue == "number" && Number.isFinite(xValue));
    xValues.length &&
      !Number.isFinite(Math.max(...xValues) - Math.min(...xValues)) &&
      addDiagnostic("Scatter X magnitude exceeds the finite plotting range.");
  } else
    chartSettings.xScale === "log" &&
      addDiagnostic(
        "A logarithmic X axis is available only for scatter charts.",
      );
  (chartSettings.yScale === "log" &&
    ["pie", "donut", "waterfall", "histogram", "sparkline"].includes(chart) &&
    addDiagnostic(
      `Log axes are not supported for ${chart}; choose a linear scale.`,
    ),
    ["pie", "donut"].includes(chart) &&
      selectedSeriesIndexes.length === 1 &&
      !includedPoints.some(
        (point) => Number(point.values[selectedSeriesIndexes[0]].value) > 0,
      ) &&
      addDiagnostic("Pie/donut requires at least one positive value."));
  const population = {
      sourceRows: points.length,
      includedRows: includedPoints.length,
      excludedRows: points.length - includedPoints.length,
      hiddenRows: points.filter((point) => point.hidden).length,
      availableSeries: series.length,
      selectedSeries: selectedSeriesIndexes.length,
      policy:
        "All source rows, including worksheet-hidden or filtered rows, are inspected. Only explicit chart row/series selections affect the plotted population. Blank cells are gaps; invalid values are diagnostics, never zero.",
    },
    scope = `${sheet.name}!${range} · ${includedPoints.length}/${points.length} rows · ${selectedSeriesIndexes.length}/${series.length} series · ${population.hiddenRows} hidden rows in source`,
    chartModel = {
      version: 1,
      objectId: object.id,
      kind: chart,
      title: object.label,
      altText:
        chartSettings.altText?.trim() ||
        `${object.label || chart}. ${scope}. ${population.policy}`,
      source: { sheetId: sheet.id, sheetName: sheet.name, range: range },
      axes: {
        x: {
          title: chartSettings.xTitle ?? "",
          scale: chartSettings.xScale ?? "linear",
        },
        y: {
          title:
            chartSettings.yTitle ??
            (chart === "histogram" ? "Frequency (records)" : ""),
          scale: chartSettings.yScale ?? "linear",
        },
        units: chartSettings.units ?? "",
      },
      options: chartSettings,
      population: population,
      scope: scope,
      series: series,
      points: points,
      diagnostics: diagnostics,
    },
    plottedValues =
      selectedSeriesIndexes.length === 1
        ? includedPoints
            .filter(
              (point) =>
                typeof point.values[selectedSeriesIndexes[0]].value ==
                  "number" &&
                Number.isFinite(point.values[selectedSeriesIndexes[0]].value),
            )
            .map((point) => ({
              label: point.label,
              row: point.row,
              value: point.values[selectedSeriesIndexes[0]].value,
            }))
        : [];
  if (chart === "histogram" && plottedValues.length) {
    const minValue = Math.min(...plottedValues.map((entry) => entry.value)),
      maxValue = Math.max(...plottedValues.map((entry) => entry.value)),
      binCount =
        minValue === maxValue
          ? 1
          : (chartSettings.histogramBins ??
            Math.min(50, Math.ceil(Math.sqrt(plottedValues.length)))),
      binWidth = (maxValue - minValue) / binCount;
    if (!Number.isFinite(binWidth))
      addDiagnostic(
        "Histogram magnitude exceeds the finite calculation range.",
      );
    else {
      chartModel.histogram = Array.from(
        { length: binCount },
        (_, binIndex) => ({
          lower: minValue + binIndex * binWidth,
          upper:
            binIndex === binCount - 1
              ? maxValue
              : minValue + (binIndex + 1) * binWidth,
          count: 0,
          rows: [],
          upperInclusive: binIndex === binCount - 1,
        }),
      );
      for (const entry of plottedValues) {
        const bin =
          chartModel.histogram[
            binWidth
              ? Math.min(
                  binCount - 1,
                  Math.floor((entry.value - minValue) / binWidth),
                )
              : 0
          ];
        (bin.count++, bin.rows.push(entry.row));
      }
    }
  }
  if (chart === "waterfall" && plottedValues.length) {
    let end = 0;
    ((chartModel.waterfall = plottedValues.map((plottedValue) => {
      const start = end;
      return (
        (end += plottedValue.value),
        { ...plottedValue, start: start, end: end }
      );
    })),
      chartModel.waterfall.push({
        label: "Total",
        row: null,
        start: 0,
        end: end,
        value: end,
      }),
      (!Number.isFinite(end) ||
        chartModel.waterfall.some(
          (waterfallEntry) => !Number.isFinite(waterfallEntry.end),
        )) &&
        addDiagnostic(
          "Waterfall cumulative values exceed the finite calculation range.",
        ));
  }
  ["pie", "donut"].includes(chart) &&
    selectedSeriesIndexes.length === 1 &&
    !Number.isFinite(
      plottedValues.reduce(
        (total, plottedValue) => total + plottedValue.value,
        0,
      ),
    ) &&
    addDiagnostic("Pie/donut total exceeds the finite calculation range.");
  const selectedValues = selectedSeriesIndexes.flatMap((selectedSeriesIndex) =>
    includedPoints
      .map((point) => point.values[selectedSeriesIndex].value)
      .filter((value) => typeof value == "number" && Number.isFinite(value)),
  );
  return (
    selectedValues.length &&
      !Number.isFinite(
        Math.max(...selectedValues) - Math.min(0, ...selectedValues),
      ) &&
      addDiagnostic("Chart magnitude exceeds the finite plotting range."),
    chartModel
  );
}
function measureChartSelection(range, sheet) {
  const rangeAddress = parseRangeAddress(range);
  if (
    rangeAddress.lastRow >= sheet.rowCount ||
    rangeAddress.lastCol >= sheet.colCount ||
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      Yne
  )
    throw new Error(
      "Select at most 500 cells inside this worksheet for a chart.",
    );
  return rangeAddress;
}
function validateDrawingObjects(drawingObjects, sheet) {
  if (drawingObjects === undefined) return;
  if (!Array.isArray(drawingObjects) || drawingObjects.length > Kne)
    throw new Error("A worksheet supports at most 100 drawing objects.");
  const usedDrawingIds = new Set();
  let pictureBytes = 0,
    estimatedPictureBytes = 0;
  for (const drawingObject of drawingObjects)
    if (
      isNonArrayObject(drawingObject) &&
      drawingObject.kind === "picture" &&
      typeof drawingObject.dataURL == "string"
    ) {
      if (
        drawingObject.dataURL.length >
        Math.ceil(maxPictureHashBytes / 3) * 4 + 100
      )
        throw Error("Pictures must be at most 2 MB.");
      const base64Data = drawingObject.dataURL.slice(
        drawingObject.dataURL.indexOf(",") + 1,
      );
      if (
        ((estimatedPictureBytes += Math.max(
          0,
          Math.floor((base64Data.length * 3) / 4) -
            (base64Data.endsWith("==") ? 2 : base64Data.endsWith("=") ? 1 : 0),
        )),
        estimatedPictureBytes > 10 * 1024 * 1024)
      )
        throw Error("Pictures in one worksheet may use at most 10 MB.");
    }
  const validatedObjects = drawingObjects.map((drawingObject) => {
    if (
      !isNonArrayObject(drawingObject) ||
      !["picture", "shape", "chart"].includes(String(drawingObject.kind))
    )
      throw new Error("Invalid worksheet drawing object.");
    const drawingId = validateDrawingText(drawingObject.id, 100, "id"),
      anchorText = validateDrawingText(drawingObject.anchor, 20, "anchor");
    if (!drawingId || usedDrawingIds.has(drawingId))
      throw new Error("Drawing IDs must be unique.");
    usedDrawingIds.add(drawingId);
    const anchorCell = _t(anchorText);
    if (anchorCell.row >= sheet.rowCount || anchorCell.col >= sheet.colCount)
      throw new Error("Drawing anchor exceeds the worksheet.");
    const normalizedObject = {
      id: drawingId,
      kind: drawingObject.kind,
      anchor: `${columnIndexToLetters(anchorCell.col)}${anchorCell.row + 1}`,
      offsetX: validateDrawingNumber(
        drawingObject.offsetX,
        -4e3,
        4e3,
        "horizontal offset",
      ),
      offsetY: validateDrawingNumber(
        drawingObject.offsetY,
        -4e3,
        4e3,
        "vertical offset",
      ),
      width: validateDrawingNumber(drawingObject.width, 24, 4e3, "width"),
      height: validateDrawingNumber(drawingObject.height, 24, 4e3, "height"),
      label: validateDrawingText(drawingObject.label, 500, "label"),
      fill: validateDrawingColor(drawingObject.fill),
      stroke: validateDrawingColor(drawingObject.stroke),
    };
    if (normalizedObject.kind === "picture") {
      normalizedObject.dataURL = validateDrawingText(
        drawingObject.dataURL,
        Math.ceil(maxPictureHashBytes / 3) * 4 + 100,
        "picture",
      );
      const pictureInfo = inspectPictureDataUrl(normalizedObject.dataURL);
      ((pictureBytes += pictureInfo.bytes.length),
        drawingObject.evidence !== undefined &&
          (normalizedObject.evidence = validatePictureEvidence(
            drawingObject.evidence,
            pictureInfo.bytes,
          )));
    } else if (drawingObject.evidence !== undefined)
      throw Error("Evidence metadata belongs only to pictures.");
    if (normalizedObject.kind === "shape") {
      if (
        !["rectangle", "ellipse", "arrow", "text"].includes(
          String(drawingObject.shape),
        )
      )
        throw new Error("Unsupported shape.");
      normalizedObject.shape = drawingObject.shape;
    }
    if (normalizedObject.kind === "chart") {
      if (!supportedChartKinds.includes(drawingObject.chart))
        throw new Error("Unsupported chart type.");
      if (
        ((normalizedObject.chart = drawingObject.chart),
        (normalizedObject.range = validateDrawingText(
          drawingObject.range,
          50,
          "chart range",
        )),
        measureChartSelection(normalizedObject.range, sheet),
        typeof drawingObject.headers != "boolean")
      )
        throw new Error("Invalid chart header setting.");
      ((normalizedObject.headers = drawingObject.headers),
        drawingObject.chartOptions !== undefined &&
          (normalizedObject.chartOptions = validateChartSettings(
            drawingObject.chartOptions,
          )));
    }
    return normalizedObject;
  });
  if (pictureBytes > 10 * 1024 * 1024)
    throw new Error("Pictures in one worksheet may use at most 10 MB.");
  sheet.objects = validatedObjects;
}
function setSheetDrawingObjects(workbook, sheetId, drawingObjects, n = {}) {
  assertDrawingPreviewCurrent(workbook, sheetId, n);
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    nextSheet = { ...sheet };
  return (
    validateDrawingObjects(drawingObjects, nextSheet),
    JSON.stringify(sheet.objects ?? []) ===
    JSON.stringify(nextSheet.objects ?? [])
      ? workbook
      : {
          ...workbook,
          sheets: workbook.sheets.map((sheet) =>
            sheet.id !== sheetId ? sheet : nextSheet,
          ),
        }
  );
}
const xf = (value, label, maxLength = 255) =>
    typeof value == "string" && value.trim() && value.length <= maxLength
      ? value
      : throwWorkbookError(`${label} must contain 1–${maxLength} characters.`),
  requireMetadataId = (value) => xf(value, "Metadata ID", 100),
  toA1RangeAddress = (range) => {
    const rangeAddress = parseRangeAddress(range);
    return `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
  },
  Cn = (rowIndex, columnIndex) =>
    `${columnIndexToLetters(columnIndex)}${rowIndex + 1}`,
  Dp = (range, row, col) =>
    row >= range.firstRow &&
    row <= range.lastRow &&
    col >= range.firstCol &&
    col <= range.lastCol,
  rangesIntersect = (firstRange, secondRange) =>
    firstRange.firstRow <= secondRange.lastRow &&
    firstRange.lastRow >= secondRange.firstRow &&
    firstRange.firstCol <= secondRange.lastCol &&
    firstRange.lastCol >= secondRange.firstCol,
  measureRangeArea = (rangeAddress) =>
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
    (rangeAddress.lastCol - rangeAddress.firstCol + 1),
  tablePaletteColors = {
    blue: ["#174a7e", "#e8f1fa"],
    green: ["#176344", "#e9f3ed"],
    gray: ["#4b5563", "#f0f1f3"],
  },
  tablePaletteStyleNames = {
    blue: "TableStyleMedium2",
    green: "TableStyleMedium4",
    gray: "TableStyleMedium9",
  };
function resolveSourceRange(rangeText, sheetOrThrow) {
  const rangeAddress = parseRangeAddress(xf(rangeText, "Source range", 50));
  return (
    measureRangeArea(rangeAddress) > workbookLimits.maxCells &&
      throwWorkbookError(
        "Table, query and pivot ranges are limited to 100,000 cells including blank gaps. Select the used data range.",
      ),
    sheetOrThrow &&
      (rangeAddress.lastRow >= sheetOrThrow.rowCount ||
        rangeAddress.lastCol >= sheetOrThrow.colCount) &&
      throwWorkbookError("Range exceeds worksheet bounds."),
    rangeAddress
  );
}
function validateColumnName(value) {
  const columnName = xf(value, "Column name");
  return (
    /[\[\]\r\n]/.test(columnName) &&
      throwWorkbookError(
        "Column names cannot contain brackets or line breaks.",
      ),
    columnName
  );
}
function validateTableName(value) {
  const tableName = xf(value, "Table name");
  return (
    (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(tableName) ||
      /^[RC]$/i.test(tableName) ||
      /^[A-Z]{1,3}[1-9]\d*$/i.test(tableName) ||
      /^R\d+C\d+$/i.test(tableName)) &&
      throwWorkbookError(
        "Use a table name beginning with a letter or underscore, without spaces or a cell-address name.",
      ),
    tableName
  );
}
function validateFieldList(fields, label) {
  (!Array.isArray(fields) || fields.length > workbookLimits.maxCols) &&
    throwWorkbookError(`${label} must be a bounded list of column names.`);
  const validatedFields = fields.map(validateColumnName);
  return (
    new Set(validatedFields.map((field) => field.toLowerCase())).size !==
      validatedFields.length &&
      throwWorkbookError(`${label} contains duplicate fields.`),
    validatedFields
  );
}
function normalizeSourceReference(source) {
  isPlainDataObject(source) ||
    throwWorkbookError("A source worksheet and range are required.");
  const normalizedSource = {
    sheetId: requireMetadataId(source.sheetId),
    range: toA1RangeAddress(xf(source.range, "Source range", 50)),
    ...(source.tableId !== undefined
      ? { tableId: requireMetadataId(source.tableId) }
      : {}),
  };
  return (resolveSourceRange(normalizedSource.range), normalizedSource);
}
function normalizeSourceFilters(filters) {
  return filters === undefined
    ? []
    : ((!Array.isArray(filters) || filters.length > 100) &&
        throwWorkbookError("Use at most 100 source filters."),
      filters.map((filter) => {
        (!isPlainDataObject(filter) ||
          ![
            "equals",
            "notEquals",
            "contains",
            "gt",
            "gte",
            "lt",
            "lte",
            "blank",
            "notBlank",
          ].includes(String(filter.operator))) &&
          throwWorkbookError("Invalid source filter.");
        const value = filter.value;
        return (
          value != null &&
            !(typeof value == "string" && value.length <= 32767) &&
            typeof value != "boolean" &&
            !(typeof value == "number" && Number.isFinite(value)) &&
            throwWorkbookError(
              "Filter values must be finite numbers, text, booleans or blank.",
            ),
          {
            field: validateColumnName(filter.field),
            operator: filter.operator,
            ...(value !== undefined ? { value: value } : {}),
          }
        );
      }));
}
function normalizeOutputDestination(destination) {
  if (destination === undefined) return;
  isPlainDataObject(destination) ||
    throwWorkbookError("Invalid output destination.");
  const addressText =
      destination.address === undefined
        ? "A1"
        : xf(destination.address, "Output address", 20),
    addressCell = _t(addressText);
  return {
    address: Cn(addressCell.row, addressCell.col),
    ...(destination.sheetId !== undefined
      ? { sheetId: requireMetadataId(destination.sheetId) }
      : {}),
    ...(destination.sheetName !== undefined
      ? { sheetName: xf(destination.sheetName, "Output sheet name", 31) }
      : {}),
  };
}
function normalizePivotDefinition(definition) {
  (isPlainDataObject(definition) ||
    throwWorkbookError("Invalid pivot definition."),
    (!Array.isArray(definition.values) ||
      !definition.values.length ||
      definition.values.length > 100) &&
      throwWorkbookError("Choose between 1 and 100 pivot value fields."));
  const values = definition.values.map(
    (value) => (
      (!isPlainDataObject(value) ||
        !pivotAggregateOperations.includes(value.operation)) &&
        throwWorkbookError(
          "Invalid pivot aggregation. Choose sum, count, average, min, max or distinctCount.",
        ),
      value.showAs !== undefined &&
        !pivotShowValuesAsModes.includes(value.showAs) &&
        throwWorkbookError("Choose a supported Show values as setting."),
      {
        field: validateColumnName(value.field),
        operation: value.operation,
        ...(value.name !== undefined
          ? { name: validateColumnName(value.name) }
          : {}),
        ...(value.showAs !== undefined ? { showAs: value.showAs } : {}),
      }
    ),
  );
  (definition.grandTotals !== undefined &&
    typeof definition.grandTotals != "boolean" &&
    throwWorkbookError("Grand totals must be boolean."),
    definition.subtotals !== undefined &&
      typeof definition.subtotals != "boolean" &&
      throwWorkbookError("Subtotals must be boolean."),
    definition.dateGroups !== undefined &&
      (!Array.isArray(definition.dateGroups) ||
        definition.dateGroups.length > 100) &&
      throwWorkbookError("Use at most 100 date groups."));
  const dateGroups = (definition.dateGroups ?? []).map(
    (dateGroupInput) => (
      (!isPlainDataObject(dateGroupInput) ||
        !["year", "quarter", "month"].includes(String(dateGroupInput.unit)) ||
        !["iso", "1900", "1904"].includes(String(dateGroupInput.dateSystem))) &&
        throwWorkbookError(
          "Date grouping requires a year, quarter or month and an explicit ISO/1900/1904 date system.",
        ),
      {
        field: validateColumnName(dateGroupInput.field),
        unit: dateGroupInput.unit,
        dateSystem: dateGroupInput.dateSystem,
      }
    ),
  );
  (new Set(dateGroups.map((dateGroup) => dateGroup.field)).size !==
    dateGroups.length &&
    throwWorkbookError("Choose one date grouping per source field."),
    definition.calculatedFields !== undefined &&
      (!Array.isArray(definition.calculatedFields) ||
        definition.calculatedFields.length > 100) &&
      throwWorkbookError("Use at most 100 calculated fields."));
  const calculatedFields = (definition.calculatedFields ?? []).map(
    (calculatedFieldInput) => (
      (!isPlainDataObject(calculatedFieldInput) ||
        !["add", "subtract", "multiply", "divide"].includes(
          String(calculatedFieldInput.operation),
        ) ||
        (calculatedFieldInput.otherField === undefined) ==
          (calculatedFieldInput.constant === undefined) ||
        (calculatedFieldInput.constant !== undefined &&
          (typeof calculatedFieldInput.constant != "number" ||
            !Number.isFinite(calculatedFieldInput.constant)))) &&
        throwWorkbookError(
          "A calculated field needs add/subtract/multiply/divide and exactly one numeric constant or other source field.",
        ),
      {
        name: validateColumnName(calculatedFieldInput.name),
        field: validateColumnName(calculatedFieldInput.field),
        operation: calculatedFieldInput.operation,
        ...(calculatedFieldInput.otherField !== undefined
          ? { otherField: validateColumnName(calculatedFieldInput.otherField) }
          : { constant: calculatedFieldInput.constant }),
      }
    ),
  );
  return (
    new Set(
      calculatedFields.map((calculatedField) =>
        calculatedField.name.toLowerCase(),
      ),
    ).size !== calculatedFields.length &&
      throwWorkbookError("Calculated field names must be unique."),
    {
      name: xf(definition.name, "Pivot name", 100),
      source: normalizeSourceReference(definition.source),
      rowFields: validateFieldList(definition.rowFields, "Row fields"),
      columnFields: validateFieldList(definition.columnFields, "Column fields"),
      values: values,
      filters: normalizeSourceFilters(definition.filters),
      grandTotals: definition.grandTotals !== false,
      ...(definition.subtotals !== undefined
        ? { subtotals: definition.subtotals }
        : {}),
      ...(definition.dateGroups !== undefined
        ? { dateGroups: dateGroups }
        : {}),
      ...(definition.calculatedFields !== undefined
        ? { calculatedFields: calculatedFields }
        : {}),
      ...normalizePivotControls(definition),
      destination: normalizeOutputDestination(definition.destination),
    }
  );
}
function normalizeQueryDefinition(definition) {
  isPlainDataObject(definition) ||
    throwWorkbookError("Invalid query definition.");
  const fields = validateFieldList(definition.fields, "Selected fields");
  fields.length || throwWorkbookError("Choose at least one output field.");
  const rename = {};
  if (definition.rename !== undefined) {
    isPlainDataObject(definition.rename) ||
      throwWorkbookError("Rename fields must be a name map.");
    for (const [fieldName, newColumnName] of Object.entries(definition.rename))
      (fields.includes(fieldName) ||
        throwWorkbookError("Only selected fields can be renamed."),
        Object.defineProperty(rename, fieldName, {
          value: validateColumnName(newColumnName),
          enumerable: true,
          writable: true,
          configurable: true,
        }));
  }
  let sort;
  (definition.sort !== undefined &&
    ((!isPlainDataObject(definition.sort) ||
      !["ascending", "descending"].includes(
        String(definition.sort.direction),
      )) &&
      throwWorkbookError("Invalid query sort."),
    (sort = {
      field: validateColumnName(definition.sort.field),
      direction: definition.sort.direction,
    })),
    definition.derived !== undefined &&
      (!Array.isArray(definition.derived) || definition.derived.length > 100) &&
      throwWorkbookError("Use at most 100 derived fields."));
  const derived = (definition.derived ?? []).map(
      (derivation) => (
        (!isPlainDataObject(derivation) ||
          ![
            "add",
            "subtract",
            "multiply",
            "divide",
            "concat",
            "upper",
            "lower",
            "trim",
          ].includes(String(derivation.operation))) &&
          throwWorkbookError("Invalid derived operation."),
        derivation.constant !== undefined &&
          !(
            typeof derivation.constant == "string" &&
            derivation.constant.length <= 32767
          ) &&
          !(
            typeof derivation.constant == "number" &&
            Number.isFinite(derivation.constant)
          ) &&
          throwWorkbookError(
            "A derived constant must be bounded text or a finite number.",
          ),
        {
          name: validateColumnName(derivation.name),
          field: validateColumnName(derivation.field),
          operation: derivation.operation,
          ...(derivation.otherField !== undefined
            ? { otherField: validateColumnName(derivation.otherField) }
            : {}),
          ...(derivation.constant !== undefined
            ? { constant: derivation.constant }
            : {}),
        }
      ),
    ),
    outputNames = [
      ...fields.map((field) =>
        Object.hasOwn(rename, field) ? rename[field] : field,
      ),
      ...derived.map((derivedRule) => derivedRule.name),
    ];
  return (
    new Set(outputNames.map((columnName) => columnName.toLowerCase())).size !==
      outputNames.length &&
      throwWorkbookError("Output column names must be unique."),
    {
      name: xf(definition.name, "Query name", 100),
      source: normalizeSourceReference(definition.source),
      fields: fields,
      rename: rename,
      filters: normalizeSourceFilters(definition.filters),
      ...(sort ? { sort: sort } : {}),
      derived: derived,
      destination: normalizeOutputDestination(definition.destination),
    }
  );
}
function normalizeTableDefinitions(sheetDefinitions, sheetOrThrow) {
  if (isPlainDataObject(sheetDefinitions)) {
    if (sheetDefinitions.tables !== undefined) {
      (!Array.isArray(sheetDefinitions.tables) ||
        sheetDefinitions.tables.length > workbookLimits.maxCells) &&
        throwWorkbookError("Tables must be a bounded list.");
      const tableIds = new Set(),
        tableNames = new Set();
      sheetOrThrow.tables = sheetDefinitions.tables.map((table) => {
        (!isPlainDataObject(table) ||
          !Array.isArray(table.columns) ||
          typeof table.showTotals != "boolean" ||
          !["blue", "green", "gray"].includes(String(table.style)) ||
          !isPlainDataObject(table.totals)) &&
          throwWorkbookError("Invalid table metadata.");
        const tableId = requireMetadataId(table.id),
          tableName = validateTableName(table.name),
          a1RangeAddress = toA1RangeAddress(xf(table.range, "Table range", 50)),
          sourceRange = resolveSourceRange(a1RangeAddress, sheetOrThrow);
        (sheetOrThrow.merges?.some((merge) =>
          rangesIntersect(sourceRange, parseRangeAddress(merge)),
        ) &&
          throwWorkbookError(
            "A table cannot overlap merged cells. Unmerge the source first.",
          ),
          (tableIds.has(tableId) || tableNames.has(tableName.toLowerCase())) &&
            throwWorkbookError("Table IDs and names must be unique."),
          tableIds.add(tableId),
          tableNames.add(tableName.toLowerCase()),
          (table.columns.length !==
            sourceRange.lastCol - sourceRange.firstCol + 1 ||
            (table.showTotals &&
              sourceRange.lastRow <= sourceRange.firstRow)) &&
            throwWorkbookError(
              "Table columns or totals do not fit the table range.",
            ));
        const columns = table.columns.map(
          (column) => (
            isPlainDataObject(column) ||
              throwWorkbookError("Invalid table column."),
            {
              id: requireMetadataId(column.id),
              name: validateColumnName(column.name),
              ...(column.dataType !== undefined
                ? { dataType: validateTableColumnType(column.dataType) }
                : {}),
              ...(column.calculatedFormula !== undefined
                ? {
                    calculatedFormula: validateCalculatedColumnFormula(
                      column.calculatedFormula,
                      true,
                    ),
                  }
                : {}),
            }
          ),
        );
        (new Set(columns.map((column) => column.id)).size !== columns.length ||
          new Set(columns.map((column) => column.name.toLowerCase())).size !==
            columns.length) &&
          throwWorkbookError("Table column IDs and names must be unique.");
        const totals = {};
        for (const [columnName, totalOperation] of Object.entries(table.totals))
          ((!columns.some((column) => column.name === columnName) ||
            !basicAggregateOperations.includes(totalOperation)) &&
            throwWorkbookError("Invalid table total field or operation."),
            Object.defineProperty(totals, columnName, {
              value: totalOperation,
              enumerable: true,
              writable: true,
              configurable: true,
            }));
        return {
          id: tableId,
          name: tableName,
          range: a1RangeAddress,
          columns: columns,
          style: table.style,
          showTotals: table.showTotals,
          totals: totals,
        };
      });
      const occupiedCells = new Set();
      let totalFootprint = 0;
      for (const table of sheetOrThrow.tables) {
        const rangeAddress = parseRangeAddress(table.range);
        (totalFootprint += measureRangeArea(rangeAddress)) >
          workbookLimits.maxCells &&
          throwWorkbookError(
            "The total table footprint exceeds 100,000 cells.",
          );
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
            const cellAddress = Cn(firstRow, firstCol);
            (occupiedCells.has(cellAddress) &&
              throwWorkbookError("Tables cannot overlap."),
              occupiedCells.add(cellAddress));
          }
      }
    }
    for (const collectionName of ["pivots", "queries"]) {
      if (sheetDefinitions[collectionName] === undefined) continue;
      (!Array.isArray(sheetDefinitions[collectionName]) ||
        sheetDefinitions[collectionName].length > workbookLimits.maxCells) &&
        throwWorkbookError("Analysis definitions must be a bounded list.");
      const analysisIds = new Set(),
        analyses = sheetDefinitions[collectionName].map((analysis) => {
          isPlainDataObject(analysis) ||
            throwWorkbookError("Invalid analysis metadata.");
          const normalizedDefinition =
              collectionName === "pivots"
                ? normalizePivotDefinition(analysis)
                : normalizeQueryDefinition(analysis),
            analysisId = requireMetadataId(analysis.id),
            a1RangeAddress = toA1RangeAddress(
              xf(analysis.outputRange, "Output range", 50),
            );
          (resolveSourceRange(a1RangeAddress, sheetOrThrow),
            (analysisIds.has(analysisId) ||
              !Number.isSafeInteger(analysis.sourceRevision) ||
              Number(analysis.sourceRevision) < 0) &&
              throwWorkbookError(
                "Invalid analysis identity or source revision.",
              ),
            analysisIds.add(analysisId));
          const { destination, ...withoutDestination } = normalizedDefinition;
          return {
            ...withoutDestination,
            id: analysisId,
            outputRange: a1RangeAddress,
            sourceRevision: Number(analysis.sourceRevision),
            ...(collectionName === "pivots" && analysis.refresh !== undefined
              ? { refresh: validatePivotRefreshEvidence(analysis.refresh) }
              : {}),
          };
        });
      collectionName === "pivots"
        ? (sheetOrThrow.pivots = analyses)
        : (sheetOrThrow.queries = analyses);
    }
  }
}
function checkStructuralEditImpact(
  workbook,
  sheetId,
  axis,
  startIndex,
  count,
  remove,
) {
  const axisLimit =
    axis === "row" ? workbookLimits.maxRows : workbookLimits.maxCols;
  ((!Number.isInteger(startIndex) ||
    startIndex < 0 ||
    startIndex >= axisLimit ||
    !Number.isInteger(count) ||
    count < 1 ||
    startIndex + count > axisLimit) &&
    throwWorkbookError("Invalid structural edit bounds."),
    workbook.sheets.some((sheet) =>
      sheet.tables?.some((table) =>
        table.columns.some((column) => column.calculatedFormula),
      ),
    ) &&
      throwWorkbookError(
        "Calculated table formulas require explicit table append/resize. Clear their saved rules (keeping cell formulas) before inserting or deleting worksheet rows/columns, then reapply the rules.",
      ));
  const updatedWorkbook = structuredClone(workbook);
  function shiftRangeAddress(rangeText, kindLabel, hasTotalsRow = false) {
    const sourceRange = resolveSourceRange(rangeText),
      firstAxisIndex =
        axis === "row" ? sourceRange.firstRow : sourceRange.firstCol,
      lastAxisIndex =
        axis === "row" ? sourceRange.lastRow : sourceRange.lastCol;
    let shiftedFirst = firstAxisIndex,
      shiftedLast = lastAxisIndex;
    if (remove) {
      const removedLastIndex = startIndex + count - 1;
      (startIndex <= firstAxisIndex &&
        removedLastIndex >= firstAxisIndex &&
        throwWorkbookError(
          `This edit would remove a ${kindLabel} header. Move or recreate that structure first.`,
        ),
        removedLastIndex < firstAxisIndex
          ? ((shiftedFirst -= count), (shiftedLast -= count))
          : startIndex <= lastAxisIndex &&
            (axis === "column" &&
              throwWorkbookError(
                `Deleting columns inside a ${kindLabel} would invalidate its fields. Resize or recreate it first.`,
              ),
            hasTotalsRow &&
              removedLastIndex >= lastAxisIndex &&
              throwWorkbookError(
                "This edit would remove a table totals row. Turn off table totals first.",
              ),
            (shiftedLast -=
              Math.min(lastAxisIndex, removedLastIndex) - startIndex + 1)));
    } else
      startIndex <= firstAxisIndex
        ? ((shiftedFirst += count), (shiftedLast += count))
        : startIndex <= lastAxisIndex &&
          (axis === "column" &&
            throwWorkbookError(
              `Inserting columns inside a ${kindLabel} requires explicit field reconciliation. Resize or recreate it first.`,
            ),
          (shiftedLast += count));
    return (
      shiftedLast >= axisLimit &&
        throwWorkbookError(
          `The structural edit would push a ${kindLabel} beyond the worksheet boundary.`,
        ),
      axis === "row"
        ? ((sourceRange.firstRow = shiftedFirst),
          (sourceRange.lastRow = shiftedLast))
        : ((sourceRange.firstCol = shiftedFirst),
          (sourceRange.lastCol = shiftedLast)),
      measureRangeArea(sourceRange) > workbookLimits.maxCells &&
        throwWorkbookError(
          "The shifted table or analysis exceeds the 100,000-cell range budget.",
        ),
      `${Cn(sourceRange.firstRow, sourceRange.firstCol)}:${Cn(sourceRange.lastRow, sourceRange.lastCol)}`
    );
  }
  for (const sheet of updatedWorkbook.sheets) {
    if (sheet.id === sheetId) {
      for (const sheetTable of sheet.tables ?? [])
        sheetTable.range = shiftRangeAddress(
          sheetTable.range,
          "table",
          sheetTable.showTotals,
        );
      for (const sheetAnalysis of [
        ...(sheet.pivots ?? []),
        ...(sheet.queries ?? []),
      ])
        sheetAnalysis.outputRange = shiftRangeAddress(
          sheetAnalysis.outputRange,
          "analysis result",
        );
    }
    for (const analysisDefinition of [
      ...(sheet.pivots ?? []),
      ...(sheet.queries ?? []),
    ])
      analysisDefinition.source.sheetId === sheetId &&
        (analysisDefinition.source.range = shiftRangeAddress(
          analysisDefinition.source.range,
          "analysis source",
        ));
  }
  return updatedWorkbook;
}
function tableDataBodyRange(table) {
  const rangeAddress = parseRangeAddress(table.range);
  return {
    ...rangeAddress,
    firstRow: rangeAddress.firstRow + 1,
    lastRow: rangeAddress.lastRow - (table.showTotals ? 1 : 0),
  };
}
function rewriteStructuredReferences(
  workbook,
  sheetId,
  referenceText,
  address,
) {
  const referenceMatch = /^([A-Za-z_][A-Za-z0-9_.]*)?(\[.+\])$/.exec(
    referenceText,
  );
  if (!referenceMatch) return null;
  let resolvedTarget;
  if (referenceMatch[1])
    for (const sheet of workbook.sheets) {
      const table = sheet.tables?.find(
        (table) => table.name.toLowerCase() === referenceMatch[1].toLowerCase(),
      );
      if (table) {
        resolvedTarget = { sheet: sheet, table: table };
        break;
      }
    }
  else if (address) {
    const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
      cellPosition = _t(address),
      table = sheet?.tables?.find((table) =>
        Dp(parseRangeAddress(table.range), cellPosition.row, cellPosition.col),
      );
    sheet && table && (resolvedTarget = { sheet: sheet, table: table });
  }
  if (!resolvedTarget) return null;
  const { sheet, table } = resolvedTarget,
    rangeAddress = parseRangeAddress(table.range),
    dataRange = tableDataBodyRange(table),
    specifiers = referenceMatch[2]
      .replace(/^\[|\]$/g, "")
      .replace(/[\[\]]/g, "")
      .split(/[:,]/)
      .map((rawSpecifier) => rawSpecifier.trim()),
    specialKeyword = specifiers
      .find((specifier) => specifier.startsWith("#"))
      ?.toLowerCase(),
    hasThisRowSpecifier =
      specifiers.some((atSpecifier) => atSpecifier.startsWith("@")) ||
      specialKeyword === "#this row";
  if (
    specifiers.some(
      (keywordSpecifier) =>
        keywordSpecifier.startsWith("#") &&
        !["#all", "#data", "#headers", "#totals", "#this row"].includes(
          keywordSpecifier.toLowerCase(),
        ),
    )
  )
    return null;
  if (
    (specialKeyword === "#all" && Object.assign(dataRange, rangeAddress),
    specialKeyword === "#headers" &&
      Object.assign(dataRange, rangeAddress, {
        lastRow: rangeAddress.firstRow,
      }),
    specialKeyword === "#totals")
  ) {
    if (!table.showTotals) return null;
    Object.assign(dataRange, rangeAddress, { firstRow: rangeAddress.lastRow });
  }
  if (hasThisRowSpecifier) {
    if (!address || sheet.id !== sheetId) return null;
    const currentCell = _t(address);
    if (
      currentCell.row < dataRange.firstRow ||
      currentCell.row > dataRange.lastRow
    )
      return null;
    ((dataRange.firstRow = currentCell.row),
      (dataRange.lastRow = currentCell.row));
  }
  const columnNames = specifiers
    .filter((plainSpecifier) => !plainSpecifier.startsWith("#"))
    .map((bareSpecifier) => bareSpecifier.replace(/^@/, ""));
  if (columnNames.length) {
    if (columnNames.length > 2) return null;
    const columnIndexes = columnNames.map((columnName) =>
      table.columns.findIndex(
        (column) => column.name.toLowerCase() === columnName.toLowerCase(),
      ),
    );
    if (columnIndexes.some((columnIndex) => columnIndex < 0)) return null;
    ((dataRange.firstCol = rangeAddress.firstCol + Math.min(...columnIndexes)),
      (dataRange.lastCol = rangeAddress.firstCol + Math.max(...columnIndexes)));
  }
  return dataRange.firstRow > dataRange.lastRow
    ? null
    : {
        sheetId: sheet.id,
        range: `${Cn(dataRange.firstRow, dataRange.firstCol)}:${Cn(dataRange.lastRow, dataRange.lastCol)}`,
      };
}
function buildTableTotalCell(table, columnIndex) {
  const dataRange = tableDataBodyRange(table),
    columnName = table.columns[columnIndex].name,
    operation = Object.hasOwn(table.totals, columnName)
      ? table.totals[columnName]
      : undefined;
  return operation
    ? dataRange.lastRow < dataRange.firstRow
      ? { raw: operation === "count" || operation === "sum" ? 0 : null }
      : {
          raw: `=${operation === "count" ? "COUNTA" : operation.toUpperCase()}(${Cn(dataRange.firstRow, dataRange.firstCol + columnIndex)}:${Cn(dataRange.lastRow, dataRange.firstCol + columnIndex)})`,
        }
    : { raw: columnIndex === 0 ? "Total" : null };
}
function applyTableStyleToCells(sheetOrThrow, table) {
  const sourceRange = resolveSourceRange(table.range, sheetOrThrow),
    [headerFill, bandFill] = tablePaletteColors[table.style];
  for (
    let firstRow = sourceRange.firstRow;
    firstRow <= sourceRange.lastRow;
    firstRow++
  )
    for (
      let firstCol = sourceRange.firstCol;
      firstCol <= sourceRange.lastCol;
      firstCol++
    ) {
      const cellAddress = Cn(firstRow, firstCol),
        existingCell = sheetOrThrow.cells[cellAddress] ?? { raw: null },
        isHeaderRow = firstRow === sourceRange.firstRow,
        isTotalsRow = table.showTotals && firstRow === sourceRange.lastRow,
        nextCell = isHeaderRow
          ? {
              ...existingCell,
              raw: table.columns[firstCol - sourceRange.firstCol].name,
              literal: table.columns[
                firstCol - sourceRange.firstCol
              ].name.startsWith("=")
                ? true
                : undefined,
            }
          : isTotalsRow
            ? {
                ...existingCell,
                ...buildTableTotalCell(table, firstCol - sourceRange.firstCol),
                literal: undefined,
              }
            : existingCell;
      sheetOrThrow.cells[cellAddress] = {
        ...nextCell,
        style: {
          ...nextCell.style,
          bold: isHeaderRow || isTotalsRow,
          color: isHeaderRow ? "#ffffff" : "#222222",
          fill: isHeaderRow
            ? headerFill
            : isTotalsRow || (firstRow - sourceRange.firstRow) % 2 === 0
              ? bandFill
              : "#ffffff",
          ...(isTotalsRow ? { border: "bottom" } : {}),
        },
      };
    }
}
function filterSourceRows(pivotSource, filters) {
  const resolvedFilters = filters.map((filterSpec) => ({
    ...filterSpec,
    index: findSourceFieldIndex(pivotSource.fields, filterSpec.field),
  }));
  return pivotSource.rows.filter((row) =>
    resolvedFilters.every((resolvedFilter) => {
      const cellValue = row.values[resolvedFilter.index],
        filterValue = resolvedFilter.value ?? null;
      switch (resolvedFilter.operator) {
        case "blank":
          return isNullOrEmptyValue(cellValue);
        case "notBlank":
          return !isNullOrEmptyValue(cellValue);
        case "contains":
          return String(cellValue ?? "")
            .toLocaleLowerCase()
            .includes(String(filterValue ?? "").toLocaleLowerCase());
        case "equals":
          return (
            String(cellValue ?? "").toLocaleLowerCase() ===
            String(filterValue ?? "").toLocaleLowerCase()
          );
        case "notEquals":
          return (
            String(cellValue ?? "").toLocaleLowerCase() !==
            String(filterValue ?? "").toLocaleLowerCase()
          );
        default:
          return (
            isFiniteNumber(filterValue) ||
              throwWorkbookError(
                "Numeric comparisons require a numeric filter value.",
              ),
            isFiniteNumber(cellValue)
              ? resolvedFilter.operator === "gt"
                ? cellValue > filterValue
                : resolvedFilter.operator === "gte"
                  ? cellValue >= filterValue
                  : resolvedFilter.operator === "lt"
                    ? cellValue < filterValue
                    : cellValue <= filterValue
              : false
          );
      }
    }),
  );
}
function collectRangeValues(sheet, rangeText) {
  const rangeAddress = parseRangeAddress(rangeText);
  return computePivotRefreshFingerprint(
    Object.entries(sheet.cells)
      .filter(([cellAddress]) => {
        const cellPosition = _t(cellAddress);
        return Dp(rangeAddress, cellPosition.row, cellPosition.col);
      })
      .sort(([leftAddress], [rightAddress]) =>
        leftAddress.localeCompare(rightAddress),
      )
      .map(([entryAddress, cell]) => [entryAddress, cell.raw, !!cell.literal]),
  );
}
function buildTableCellEntries(worksheet, sheet) {
  for (const table of sheet.tables ?? []) {
    const sourceRange = resolveSourceRange(table.range, sheet),
      dataRange = tableDataBodyRange(table),
      rows = [];
    for (
      let firstRow = dataRange.firstRow;
      firstRow <= dataRange.lastRow;
      firstRow++
    ) {
      const rowValues = [];
      for (
        let firstCol = dataRange.firstCol;
        firstCol <= dataRange.lastCol;
        firstCol++
      )
        rowValues.push(worksheet.getCell(Cn(firstRow, firstCol)).value);
      rows.push(rowValues);
    }
    const totalsValues = table.showTotals
      ? table.columns.map(
          (column, totalsColumnIndex) =>
            worksheet.getCell(
              Cn(sourceRange.lastRow, sourceRange.firstCol + totalsColumnIndex),
            ).value,
        )
      : [];
    (worksheet.addTable({
      name: table.name,
      ref: Cn(sourceRange.firstRow, sourceRange.firstCol),
      headerRow: true,
      totalsRow: table.showTotals,
      style: {
        theme: tablePaletteStyleNames[table.style],
        showRowStripes: true,
      },
      columns: table.columns.map((column, tableColumnIndex) => ({
        name: column.name,
        filterButton: true,
        ...(table.showTotals
          ? {
              totalsRowFunction: Object.hasOwn(table.totals, column.name)
                ? table.totals[column.name]
                : "none",
              ...(tableColumnIndex === 0 &&
              !Object.hasOwn(table.totals, column.name)
                ? {
                    totalsRowLabel: String(
                      buildTableTotalCell(table, 0).raw ?? "",
                    ),
                  }
                : {}),
            }
          : {}),
      })),
      rows: rows,
    }),
      totalsValues.forEach((totalValue, totalIndex) => {
        worksheet.getCell(
          Cn(sourceRange.lastRow, sourceRange.firstCol + totalIndex),
        ).value = totalValue;
      }));
  }
}
function rewriteFormulaReferences(
  formula,
  workbook,
  sheetId,
  address,
  shouldCheckR1C1 = true,
) {
  let rewrittenFormula = "";
  for (let scanIndex = 0; scanIndex < formula.length;) {
    const tokenStart = scanIndex,
      startChar = formula[scanIndex];
    if (startChar === '"' || startChar === "'") {
      for (scanIndex++; scanIndex < formula.length;)
        if (formula[scanIndex++] === startChar)
          if (formula[scanIndex] === startChar) scanIndex++;
          else break;
      rewrittenFormula +=
        startChar === '"' ? '""' : formula.slice(tokenStart, scanIndex);
      continue;
    }
    const structuredMatch = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(
      formula.slice(scanIndex),
    );
    if (structuredMatch) {
      scanIndex += structuredMatch[0].length;
      let bracketDepth = 1;
      for (; scanIndex < formula.length && bracketDepth;)
        (formula[scanIndex] === "[" && bracketDepth++,
          formula[scanIndex] === "]" && bracketDepth--,
          scanIndex++);
      const structuredText = formula.slice(tokenStart, scanIndex);
      rewrittenFormula +=
        !bracketDepth &&
        rewriteStructuredReferences(workbook, sheetId, structuredText, address)
          ? "TABLETOKEN"
          : structuredText;
      continue;
    }
    rewrittenFormula += formula[scanIndex++];
  }
  const makeIssue = (kind, errorCode, message) => ({
    sheetId: sheetId,
    address: address,
    kind: kind,
    error: errorCode,
    message: message,
  });
  if (
    /\[[^\]]+\][^!+*/^&=<>(),;{}]*!/.test(rewrittenFormula) ||
    /\|/.test(rewrittenFormula) ||
    /(?:https?:|file:|\\\\)/i.test(rewrittenFormula)
  )
    return makeIssue(
      "external",
      "#BLOCKED!",
      "External workbook, path, network and DDE references are not supported. Use an authorized local worksheet reference.",
    );
  const sheetNamePattern = `(?:'(?:[^']|'')*'|[^\\s!+*/^&=<>(),;:{}\\[\\]\\"'%\\-]+)`;
  if (
    /'(?:[^']|'')*:(?:[^']|'')*'\s*!/.test(rewrittenFormula) ||
    new RegExp(`${sheetNamePattern}\\s*:\\s*${sheetNamePattern}\\s*!`).test(
      rewrittenFormula,
    )
  )
    return makeIssue(
      "three-dimensional",
      "#VALUE!",
      "3-D references across a worksheet interval are not supported. List the individual worksheet references explicitly.",
    );
  const maskedFormula = rewrittenFormula
    .replace(/'(?:[^']|'')*'\s*!/g, "Sheet!")
    .replace(
      new RegExp(
        "(?<![\\p{L}\\p{N}_.\\]])\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*(?![\\p{L}\\p{N}_.])",
        "gu",
      ),
      "CELLREF",
    );
  if (
    shouldCheckR1C1 &&
    (new RegExp(
      "(?<![\\p{L}\\p{N}_.])R(?:\\d+|\\[[+-]?\\d+\\])?C(?:\\d+|\\[[+-]?\\d+\\])?(?![\\p{L}\\p{N}_.]|\\s*[!(])",
      "iu",
    ).test(maskedFormula) ||
      new RegExp(
        "(?<![\\p{L}\\p{N}_.])[RC]\\[[+-]?\\d+\\](?![\\p{L}\\p{N}_.])",
        "iu",
      ).test(maskedFormula))
  )
    return makeIssue(
      "r1c1",
      "#REF!",
      "R1C1-style references are not supported. Use A1 references, including $A$1, $A1 or A$1 as needed.",
    );
}

export {
  measureRangeArea,
  normalizeTableDefinitions,
  rewriteFormulaReferences,
  Cn,
  Dp,
  resolveSourceRange,
  tableDataBodyRange,
  setSheetDrawingObjects,
  normalizeQueryDefinition,
  rangesIntersect,
  applyTableStyleToCells,
  inspectChartSource,
  toA1RangeAddress,
  collectRangeValues,
  checkStructuralEditImpact,
  validateTableName,
  serializeDefinedNames,
  normalizeSourceReference,
  buildTableTotalCell,
  filterSourceRows,
  measureChartSelection,
  validateColumnName,
  normalizePivotDefinition,
  buildTableCellEntries,
  validateDrawingObjects,
  rewriteStructuredReferences,
};

/* Also used by the workers. */
export {
  normalizeOutputDestination,
  normalizeSourceFilters,
  requireMetadataId,
  tablePaletteStyleNames,
  validateFieldList,
  xf,
};
