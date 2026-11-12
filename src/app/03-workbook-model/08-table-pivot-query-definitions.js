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

function serializeDefinedNames(e, workbookInput) {
  const names = validateDefinedName(workbookInput.definedNames, workbookInput);
  oP(names);
  const { builtinXml } = aP(e);
  let builtinNameXml = builtinXml;
  for (const a of names) {
    if (a.name.toLowerCase().startsWith("_xlnm."))
      return throwXlsxDefinedNameError(
        `Reserved built-in name "${a.name}" cannot be supplied as a user name.`,
      );
    (validateNameFormulaText(a.refersTo),
      validateNameSheetReferences(a.refersTo, workbookInput));
    const s = a.scopeSheetId
      ? workbookInput.sheets.findIndex((sheet) => sheet.id === a.scopeSheetId)
      : undefined;
    builtinNameXml += `<definedName name="${validateXmlSafeName(a.name, true)}"${s !== undefined ? ` localSheetId="${s}"` : ""}${a.comment !== undefined ? ` comment="${validateXmlSafeName(a.comment, true)}"` : ""}>${validateXmlSafeName(a.refersTo.slice(1))}</definedName>`;
  }
  return insertDefinedNamesIntoWorkbookXml(e, builtinNameXml);
}
function inspectChartSource(sheet, object, t) {
  (object.kind !== "chart" || !supportedChartKinds.includes(object.chart)) &&
    failWithMessage("Choose a supported chart type.");
  const chart = object.chart,
    range = object.range ?? "",
    rangeAddress = parseRangeAddress(range),
    s = validateChartSettings(object.chartOptions);
  (rangeAddress.lastRow >= sheet.rowCount ||
    rangeAddress.lastCol >= sheet.colCount ||
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      500) &&
    failWithMessage(
      "Select at most 500 cells inside this worksheet for a chart.",
    );
  const d = rangeAddress.firstRow + (object.headers ? 1 : 0),
    u =
      rangeAddress.firstCol +
      (rangeAddress.lastCol > rangeAddress.firstCol ? 1 : 0);
  rangeAddress.lastCol - u + 1 > 20 &&
    failWithMessage(
      "Chart inspection supports at most 20 value columns; narrow the source range.",
    );
  const l = (z, V) => {
      const address = `${columnIndexToLetters(V)}${z + 1}`,
        cell = sheet.cells[address] ?? { raw: null };
      return {
        address: address,
        raw: cell.raw,
        value: t(sheet.id, address),
        ...(cell.literal ? { literal: true } : {}),
        ...(cell.note ? { note: cell.note } : {}),
      };
    },
    series = Array.from({ length: rangeAddress.lastCol - u + 1 }, (z, V) => {
      const D = u + V,
        B = s.series?.find((sery) => sery.column === D + 1),
        _ = object.headers
          ? l(rangeAddress.firstRow, D).value
          : columnIndexToLetters(D);
      return {
        column: D + 1,
        name: B?.name ?? String(_ ?? columnIndexToLetters(D)),
        kind: B?.kind ?? (chart === "combo" && V > 0 ? "line" : "column"),
        included: s.series === undefined || !!B,
      };
    });
  (s.series?.some(
    (sery) => sery.column < u + 1 || sery.column > rangeAddress.lastCol + 1,
  ) &&
    failWithMessage(
      "A selected chart series no longer exists in its source range. Re-select series.",
    ),
    s.includedRows?.some(
      (includedRow) =>
        includedRow < d + 1 || includedRow > rangeAddress.lastRow + 1,
    ) &&
      failWithMessage(
        "A chart filter row is outside its source range. Reset the row filter.",
      ));
  const h = cn(sheet, "row"),
    f = (z) => h.some((V) => z >= V.start && z <= V.end),
    points = Array.from(
      { length: Math.max(0, rangeAddress.lastRow - d + 1) },
      (z, V) => {
        const D = d + V,
          category = l(D, rangeAddress.firstCol);
        return {
          row: D + 1,
          label:
            rangeAddress.firstCol < u
              ? String(category.value ?? "")
              : category.address,
          category: category,
          included:
            s.includedRows === undefined || s.includedRows.includes(D + 1),
          hidden: f(D),
          values: series.map((sery) => l(D, sery.column - 1)),
        };
      },
    ),
    m = series.map((sery, V) => (sery.included ? V : -1)).filter((z) => z >= 0),
    v = points.filter((point) => point.included),
    diagnostics = [],
    b = (z, address) =>
      diagnostics.push({
        severity: "error",
        message: z,
        ...(address ? { address: address } : {}),
      });
  (points.length || b("Select at least one data row below the chart header."),
    v.length || b("The chart filter includes no source rows."),
    m.length || b("Select at least one value series."),
    ["pie", "donut", "waterfall", "histogram", "sparkline"].includes(chart) &&
      m.length !== 1 &&
      b(
        `${chart} requires exactly one selected series. Other source series remain available in inspection.`,
      ));
  const y = (z) => z === null || z === "";
  for (let z = 0; z < series.length; z++) {
    let V = 0;
    for (const point of points) {
      const B = point.values[z];
      if (y(B.value)) continue;
      const _ = typeof B.value == "number" && Number.isFinite(B.value),
        C = point.included && series[z].included;
      if (!_) {
        diagnostics.push({
          severity: C ? "error" : "warning",
          message: `${B.address}: ${series[z].name} contains a nonnumeric or invalid calculated value${C ? "" : " outside the plotted selection"}.`,
          address: B.address,
        });
        continue;
      }
      C &&
        (V++,
        s.yScale === "log" &&
          Number(B.value) <= 0 &&
          b(`${B.address}: log-axis values must be positive.`, B.address),
        ["pie", "donut"].includes(chart) &&
          Number(B.value) < 0 &&
          b(`${B.address}: pie/donut values cannot be negative.`, B.address));
    }
    series[z].included &&
      !V &&
      b(
        `${series[z].name}: the selected series is empty or has no numeric values.`,
      );
  }
  if (chart === "scatter") {
    rangeAddress.firstCol === u &&
      b(
        "Scatter requires a numeric X column followed by at least one Y column.",
      );
    for (const V of v) {
      if (!V.values.some((value, _) => series[_].included && !y(value.value)))
        continue;
      const value = V.category.value;
      typeof value != "number" || !Number.isFinite(value)
        ? b(
            `${V.category.address}: scatter X values must be numeric.`,
            V.category.address,
          )
        : s.xScale === "log" &&
          value <= 0 &&
          b(
            `${V.category.address}: log X-axis values must be positive.`,
            V.category.address,
          );
    }
    const z = v
      .map((V) => V.category.value)
      .filter((V) => typeof V == "number" && Number.isFinite(V));
    z.length &&
      !Number.isFinite(Math.max(...z) - Math.min(...z)) &&
      b("Scatter X magnitude exceeds the finite plotting range.");
  } else
    s.xScale === "log" &&
      b("A logarithmic X axis is available only for scatter charts.");
  (s.yScale === "log" &&
    ["pie", "donut", "waterfall", "histogram", "sparkline"].includes(chart) &&
    b(`Log axes are not supported for ${chart}; choose a linear scale.`),
    ["pie", "donut"].includes(chart) &&
      m.length === 1 &&
      !v.some((z) => Number(z.values[m[0]].value) > 0) &&
      b("Pie/donut requires at least one positive value."));
  const population = {
      sourceRows: points.length,
      includedRows: v.length,
      excludedRows: points.length - v.length,
      hiddenRows: points.filter((point) => point.hidden).length,
      availableSeries: series.length,
      selectedSeries: m.length,
      policy:
        "All source rows, including worksheet-hidden or filtered rows, are inspected. Only explicit chart row/series selections affect the plotted population. Blank cells are gaps; invalid values are diagnostics, never zero.",
    },
    scope = `${sheet.name}!${range} · ${v.length}/${points.length} rows · ${m.length}/${series.length} series · ${population.hiddenRows} hidden rows in source`,
    A = {
      version: 1,
      objectId: object.id,
      kind: chart,
      title: object.label,
      altText:
        s.altText?.trim() ||
        `${object.label || chart}. ${scope}. ${population.policy}`,
      source: { sheetId: sheet.id, sheetName: sheet.name, range: range },
      axes: {
        x: { title: s.xTitle ?? "", scale: s.xScale ?? "linear" },
        y: {
          title:
            s.yTitle ?? (chart === "histogram" ? "Frequency (records)" : ""),
          scale: s.yScale ?? "linear",
        },
        units: s.units ?? "",
      },
      options: s,
      population: population,
      scope: scope,
      series: series,
      points: points,
      diagnostics: diagnostics,
    },
    k =
      m.length === 1
        ? v
            .filter(
              (z) =>
                typeof z.values[m[0]].value == "number" &&
                Number.isFinite(z.values[m[0]].value),
            )
            .map((z) => ({
              label: z.label,
              row: z.row,
              value: z.values[m[0]].value,
            }))
        : [];
  if (chart === "histogram" && k.length) {
    const z = Math.min(...k.map((_) => _.value)),
      V = Math.max(...k.map((_) => _.value)),
      D =
        z === V
          ? 1
          : (s.histogramBins ?? Math.min(50, Math.ceil(Math.sqrt(k.length)))),
      B = (V - z) / D;
    if (!Number.isFinite(B))
      b("Histogram magnitude exceeds the finite calculation range.");
    else {
      A.histogram = Array.from({ length: D }, (_, C) => ({
        lower: z + C * B,
        upper: C === D - 1 ? V : z + (C + 1) * B,
        count: 0,
        rows: [],
        upperInclusive: C === D - 1,
      }));
      for (const _ of k) {
        const C =
          A.histogram[B ? Math.min(D - 1, Math.floor((_.value - z) / B)) : 0];
        (C.count++, C.rows.push(_.row));
      }
    }
  }
  if (chart === "waterfall" && k.length) {
    let end = 0;
    ((A.waterfall = k.map((V) => {
      const start = end;
      return ((end += V.value), { ...V, start: start, end: end });
    })),
      A.waterfall.push({
        label: "Total",
        row: null,
        start: 0,
        end: end,
        value: end,
      }),
      (!Number.isFinite(end) ||
        A.waterfall.some((V) => !Number.isFinite(V.end))) &&
        b("Waterfall cumulative values exceed the finite calculation range."));
  }
  ["pie", "donut"].includes(chart) &&
    m.length === 1 &&
    !Number.isFinite(k.reduce((z, V) => z + V.value, 0)) &&
    b("Pie/donut total exceeds the finite calculation range.");
  const F = m.flatMap((z) =>
    v
      .map((V) => V.values[z].value)
      .filter((V) => typeof V == "number" && Number.isFinite(V)),
  );
  return (
    F.length &&
      !Number.isFinite(Math.max(...F) - Math.min(0, ...F)) &&
      b("Chart magnitude exceeds the finite plotting range."),
    A
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
function validateDrawingObjects(e, r) {
  if (e === undefined) return;
  if (!Array.isArray(e) || e.length > Kne)
    throw new Error("A worksheet supports at most 100 drawing objects.");
  const t = new Set();
  let n = 0,
    o = 0;
  for (const s of e)
    if (
      isNonArrayObject(s) &&
      s.kind === "picture" &&
      typeof s.dataURL == "string"
    ) {
      if (s.dataURL.length > Math.ceil(maxPictureHashBytes / 3) * 4 + 100)
        throw Error("Pictures must be at most 2 MB.");
      const d = s.dataURL.slice(s.dataURL.indexOf(",") + 1);
      if (
        ((o += Math.max(
          0,
          Math.floor((d.length * 3) / 4) -
            (d.endsWith("==") ? 2 : d.endsWith("=") ? 1 : 0),
        )),
        o > 10 * 1024 * 1024)
      )
        throw Error("Pictures in one worksheet may use at most 10 MB.");
    }
  const a = e.map((s) => {
    if (
      !isNonArrayObject(s) ||
      !["picture", "shape", "chart"].includes(String(s.kind))
    )
      throw new Error("Invalid worksheet drawing object.");
    const d = validateDrawingText(s.id, 100, "id"),
      u = validateDrawingText(s.anchor, 20, "anchor");
    if (!d || t.has(d)) throw new Error("Drawing IDs must be unique.");
    t.add(d);
    const l = _t(u);
    if (l.row >= r.rowCount || l.col >= r.colCount)
      throw new Error("Drawing anchor exceeds the worksheet.");
    const c = {
      id: d,
      kind: s.kind,
      anchor: `${columnIndexToLetters(l.col)}${l.row + 1}`,
      offsetX: validateDrawingNumber(s.offsetX, -4e3, 4e3, "horizontal offset"),
      offsetY: validateDrawingNumber(s.offsetY, -4e3, 4e3, "vertical offset"),
      width: validateDrawingNumber(s.width, 24, 4e3, "width"),
      height: validateDrawingNumber(s.height, 24, 4e3, "height"),
      label: validateDrawingText(s.label, 500, "label"),
      fill: validateDrawingColor(s.fill),
      stroke: validateDrawingColor(s.stroke),
    };
    if (c.kind === "picture") {
      c.dataURL = validateDrawingText(
        s.dataURL,
        Math.ceil(maxPictureHashBytes / 3) * 4 + 100,
        "picture",
      );
      const h = inspectPictureDataUrl(c.dataURL);
      ((n += h.bytes.length),
        s.evidence !== undefined &&
          (c.evidence = validatePictureEvidence(s.evidence, h.bytes)));
    } else if (s.evidence !== undefined)
      throw Error("Evidence metadata belongs only to pictures.");
    if (c.kind === "shape") {
      if (!["rectangle", "ellipse", "arrow", "text"].includes(String(s.shape)))
        throw new Error("Unsupported shape.");
      c.shape = s.shape;
    }
    if (c.kind === "chart") {
      if (!supportedChartKinds.includes(s.chart))
        throw new Error("Unsupported chart type.");
      if (
        ((c.chart = s.chart),
        (c.range = validateDrawingText(s.range, 50, "chart range")),
        measureChartSelection(c.range, r),
        typeof s.headers != "boolean")
      )
        throw new Error("Invalid chart header setting.");
      ((c.headers = s.headers),
        s.chartOptions !== undefined &&
          (c.chartOptions = validateChartSettings(s.chartOptions)));
    }
    return c;
  });
  if (n > 10 * 1024 * 1024)
    throw new Error("Pictures in one worksheet may use at most 10 MB.");
  r.objects = a;
}
function setSheetDrawingObjects(workbook, sheetId, t, n = {}) {
  assertDrawingPreviewCurrent(workbook, sheetId, n);
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    a = { ...sheet };
  return (
    validateDrawingObjects(t, a),
    JSON.stringify(sheet.objects ?? []) === JSON.stringify(a.objects ?? [])
      ? workbook
      : {
          ...workbook,
          sheets: workbook.sheets.map((sheet) =>
            sheet.id !== sheetId ? sheet : a,
          ),
        }
  );
}
const xf = (e, r, t = 255) =>
    typeof e == "string" && e.trim() && e.length <= t
      ? e
      : throwWorkbookError(`${r} must contain 1–${t} characters.`),
  requireMetadataId = (e) => xf(e, "Metadata ID", 100),
  toA1RangeAddress = (e) => {
    const rangeAddress = parseRangeAddress(e);
    return `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
  },
  Cn = (e, r) => `${columnIndexToLetters(r)}${e + 1}`,
  Dp = (e, r, t) =>
    r >= e.firstRow && r <= e.lastRow && t >= e.firstCol && t <= e.lastCol,
  rangesIntersect = (e, r) =>
    e.firstRow <= r.lastRow &&
    e.lastRow >= r.firstRow &&
    e.firstCol <= r.lastCol &&
    e.lastCol >= r.firstCol,
  measureRangeArea = (e) =>
    (e.lastRow - e.firstRow + 1) * (e.lastCol - e.firstCol + 1),
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
function resolveSourceRange(e, sheetOrThrow) {
  const rangeAddress = parseRangeAddress(xf(e, "Source range", 50));
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
function validateColumnName(e) {
  const r = xf(e, "Column name");
  return (
    /[\[\]\r\n]/.test(r) &&
      throwWorkbookError(
        "Column names cannot contain brackets or line breaks.",
      ),
    r
  );
}
function validateTableName(e) {
  const r = xf(e, "Table name");
  return (
    (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(r) ||
      /^[RC]$/i.test(r) ||
      /^[A-Z]{1,3}[1-9]\d*$/i.test(r) ||
      /^R\d+C\d+$/i.test(r)) &&
      throwWorkbookError(
        "Use a table name beginning with a letter or underscore, without spaces or a cell-address name.",
      ),
    r
  );
}
function validateFieldList(e, r) {
  (!Array.isArray(e) || e.length > workbookLimits.maxCols) &&
    throwWorkbookError(`${r} must be a bounded list of column names.`);
  const t = e.map(validateColumnName);
  return (
    new Set(t.map((n) => n.toLowerCase())).size !== t.length &&
      throwWorkbookError(`${r} contains duplicate fields.`),
    t
  );
}
function normalizeSourceReference(source) {
  isPlainDataObject(source) ||
    throwWorkbookError("A source worksheet and range are required.");
  const r = {
    sheetId: requireMetadataId(source.sheetId),
    range: toA1RangeAddress(xf(source.range, "Source range", 50)),
    ...(source.tableId !== undefined
      ? { tableId: requireMetadataId(source.tableId) }
      : {}),
  };
  return (resolveSourceRange(r.range), r);
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
  const r =
      destination.address === undefined
        ? "A1"
        : xf(destination.address, "Output address", 20),
    t = _t(r);
  return {
    address: Cn(t.row, t.col),
    ...(destination.sheetId !== undefined
      ? { sheetId: requireMetadataId(destination.sheetId) }
      : {}),
    ...(destination.sheetName !== undefined
      ? { sheetName: xf(destination.sheetName, "Output sheet name", 31) }
      : {}),
  };
}
function normalizePivotDefinition(e) {
  (isPlainDataObject(e) || throwWorkbookError("Invalid pivot definition."),
    (!Array.isArray(e.values) || !e.values.length || e.values.length > 100) &&
      throwWorkbookError("Choose between 1 and 100 pivot value fields."));
  const values = e.values.map(
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
  (e.grandTotals !== undefined &&
    typeof e.grandTotals != "boolean" &&
    throwWorkbookError("Grand totals must be boolean."),
    e.subtotals !== undefined &&
      typeof e.subtotals != "boolean" &&
      throwWorkbookError("Subtotals must be boolean."),
    e.dateGroups !== undefined &&
      (!Array.isArray(e.dateGroups) || e.dateGroups.length > 100) &&
      throwWorkbookError("Use at most 100 date groups."));
  const dateGroups = (e.dateGroups ?? []).map(
    (o) => (
      (!isPlainDataObject(o) ||
        !["year", "quarter", "month"].includes(String(o.unit)) ||
        !["iso", "1900", "1904"].includes(String(o.dateSystem))) &&
        throwWorkbookError(
          "Date grouping requires a year, quarter or month and an explicit ISO/1900/1904 date system.",
        ),
      {
        field: validateColumnName(o.field),
        unit: o.unit,
        dateSystem: o.dateSystem,
      }
    ),
  );
  (new Set(dateGroups.map((dateGroup) => dateGroup.field)).size !==
    dateGroups.length &&
    throwWorkbookError("Choose one date grouping per source field."),
    e.calculatedFields !== undefined &&
      (!Array.isArray(e.calculatedFields) || e.calculatedFields.length > 100) &&
      throwWorkbookError("Use at most 100 calculated fields."));
  const calculatedFields = (e.calculatedFields ?? []).map(
    (o) => (
      (!isPlainDataObject(o) ||
        !["add", "subtract", "multiply", "divide"].includes(
          String(o.operation),
        ) ||
        (o.otherField === undefined) == (o.constant === undefined) ||
        (o.constant !== undefined &&
          (typeof o.constant != "number" || !Number.isFinite(o.constant)))) &&
        throwWorkbookError(
          "A calculated field needs add/subtract/multiply/divide and exactly one numeric constant or other source field.",
        ),
      {
        name: validateColumnName(o.name),
        field: validateColumnName(o.field),
        operation: o.operation,
        ...(o.otherField !== undefined
          ? { otherField: validateColumnName(o.otherField) }
          : { constant: o.constant }),
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
      name: xf(e.name, "Pivot name", 100),
      source: normalizeSourceReference(e.source),
      rowFields: validateFieldList(e.rowFields, "Row fields"),
      columnFields: validateFieldList(e.columnFields, "Column fields"),
      values: values,
      filters: normalizeSourceFilters(e.filters),
      grandTotals: e.grandTotals !== false,
      ...(e.subtotals !== undefined ? { subtotals: e.subtotals } : {}),
      ...(e.dateGroups !== undefined ? { dateGroups: dateGroups } : {}),
      ...(e.calculatedFields !== undefined
        ? { calculatedFields: calculatedFields }
        : {}),
      ...normalizePivotControls(e),
      destination: normalizeOutputDestination(e.destination),
    }
  );
}
function normalizeQueryDefinition(e) {
  isPlainDataObject(e) || throwWorkbookError("Invalid query definition.");
  const fields = validateFieldList(e.fields, "Selected fields");
  fields.length || throwWorkbookError("Choose at least one output field.");
  const rename = {};
  if (e.rename !== undefined) {
    isPlainDataObject(e.rename) ||
      throwWorkbookError("Rename fields must be a name map.");
    for (const [s, d] of Object.entries(e.rename))
      (fields.includes(s) ||
        throwWorkbookError("Only selected fields can be renamed."),
        Object.defineProperty(rename, s, {
          value: validateColumnName(d),
          enumerable: true,
          writable: true,
          configurable: true,
        }));
  }
  let sort;
  (e.sort !== undefined &&
    ((!isPlainDataObject(e.sort) ||
      !["ascending", "descending"].includes(String(e.sort.direction))) &&
      throwWorkbookError("Invalid query sort."),
    (sort = {
      field: validateColumnName(e.sort.field),
      direction: e.sort.direction,
    })),
    e.derived !== undefined &&
      (!Array.isArray(e.derived) || e.derived.length > 100) &&
      throwWorkbookError("Use at most 100 derived fields."));
  const derived = (e.derived ?? []).map(
      (s) => (
        (!isPlainDataObject(s) ||
          ![
            "add",
            "subtract",
            "multiply",
            "divide",
            "concat",
            "upper",
            "lower",
            "trim",
          ].includes(String(s.operation))) &&
          throwWorkbookError("Invalid derived operation."),
        s.constant !== undefined &&
          !(typeof s.constant == "string" && s.constant.length <= 32767) &&
          !(typeof s.constant == "number" && Number.isFinite(s.constant)) &&
          throwWorkbookError(
            "A derived constant must be bounded text or a finite number.",
          ),
        {
          name: validateColumnName(s.name),
          field: validateColumnName(s.field),
          operation: s.operation,
          ...(s.otherField !== undefined
            ? { otherField: validateColumnName(s.otherField) }
            : {}),
          ...(s.constant !== undefined ? { constant: s.constant } : {}),
        }
      ),
    ),
    a = [
      ...fields.map((field) =>
        Object.hasOwn(rename, field) ? rename[field] : field,
      ),
      ...derived.map((s) => s.name),
    ];
  return (
    new Set(a.map((s) => s.toLowerCase())).size !== a.length &&
      throwWorkbookError("Output column names must be unique."),
    {
      name: xf(e.name, "Query name", 100),
      source: normalizeSourceReference(e.source),
      fields: fields,
      rename: rename,
      filters: normalizeSourceFilters(e.filters),
      ...(sort ? { sort: sort } : {}),
      derived: derived,
      destination: normalizeOutputDestination(e.destination),
    }
  );
}
function normalizeTableDefinitions(e, sheetOrThrow) {
  if (isPlainDataObject(e)) {
    if (e.tables !== undefined) {
      (!Array.isArray(e.tables) || e.tables.length > workbookLimits.maxCells) &&
        throwWorkbookError("Tables must be a bounded list.");
      const t = new Set(),
        n = new Set();
      sheetOrThrow.tables = e.tables.map((table) => {
        (!isPlainDataObject(table) ||
          !Array.isArray(table.columns) ||
          typeof table.showTotals != "boolean" ||
          !["blue", "green", "gray"].includes(String(table.style)) ||
          !isPlainDataObject(table.totals)) &&
          throwWorkbookError("Invalid table metadata.");
        const d = requireMetadataId(table.id),
          u = validateTableName(table.name),
          a1RangeAddress = toA1RangeAddress(xf(table.range, "Table range", 50)),
          sourceRange = resolveSourceRange(a1RangeAddress, sheetOrThrow);
        (sheetOrThrow.merges?.some((merge) =>
          rangesIntersect(sourceRange, parseRangeAddress(merge)),
        ) &&
          throwWorkbookError(
            "A table cannot overlap merged cells. Unmerge the source first.",
          ),
          (t.has(d) || n.has(u.toLowerCase())) &&
            throwWorkbookError("Table IDs and names must be unique."),
          t.add(d),
          n.add(u.toLowerCase()),
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
        for (const [m, v] of Object.entries(table.totals))
          ((!columns.some((column) => column.name === m) ||
            !basicAggregateOperations.includes(v)) &&
            throwWorkbookError("Invalid table total field or operation."),
            Object.defineProperty(totals, m, {
              value: v,
              enumerable: true,
              writable: true,
              configurable: true,
            }));
        return {
          id: d,
          name: u,
          range: a1RangeAddress,
          columns: columns,
          style: table.style,
          showTotals: table.showTotals,
          totals: totals,
        };
      });
      const o = new Set();
      let a = 0;
      for (const table of sheetOrThrow.tables) {
        const rangeAddress = parseRangeAddress(table.range);
        (a += measureRangeArea(rangeAddress)) > workbookLimits.maxCells &&
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
            const c = Cn(firstRow, firstCol);
            (o.has(c) && throwWorkbookError("Tables cannot overlap."),
              o.add(c));
          }
      }
    }
    for (const t of ["pivots", "queries"]) {
      if (e[t] === undefined) continue;
      (!Array.isArray(e[t]) || e[t].length > workbookLimits.maxCells) &&
        throwWorkbookError("Analysis definitions must be a bounded list.");
      const n = new Set(),
        o = e[t].map((a) => {
          isPlainDataObject(a) ||
            throwWorkbookError("Invalid analysis metadata.");
          const s =
              t === "pivots"
                ? normalizePivotDefinition(a)
                : normalizeQueryDefinition(a),
            d = requireMetadataId(a.id),
            a1RangeAddress = toA1RangeAddress(
              xf(a.outputRange, "Output range", 50),
            );
          (resolveSourceRange(a1RangeAddress, sheetOrThrow),
            (n.has(d) ||
              !Number.isSafeInteger(a.sourceRevision) ||
              Number(a.sourceRevision) < 0) &&
              throwWorkbookError(
                "Invalid analysis identity or source revision.",
              ),
            n.add(d));
          const { destination, ...c } = s;
          return {
            ...c,
            id: d,
            outputRange: a1RangeAddress,
            sourceRevision: Number(a.sourceRevision),
            ...(t === "pivots" && a.refresh !== undefined
              ? { refresh: validatePivotRefreshEvidence(a.refresh) }
              : {}),
          };
        });
      t === "pivots" ? (sheetOrThrow.pivots = o) : (sheetOrThrow.queries = o);
    }
  }
}
function checkStructuralEditImpact(workbook, sheetId, axis, n, count, remove) {
  const s = axis === "row" ? workbookLimits.maxRows : workbookLimits.maxCols;
  ((!Number.isInteger(n) ||
    n < 0 ||
    n >= s ||
    !Number.isInteger(count) ||
    count < 1 ||
    n + count > s) &&
    throwWorkbookError("Invalid structural edit bounds."),
    workbook.sheets.some((sheet) =>
      sheet.tables?.some((table) =>
        table.columns.some((column) => column.calculatedFormula),
      ),
    ) &&
      throwWorkbookError(
        "Calculated table formulas require explicit table append/resize. Clear their saved rules (keeping cell formulas) before inserting or deleting worksheet rows/columns, then reapply the rules.",
      ));
  const d = structuredClone(workbook);
  function u(l, c, h = false) {
    const sourceRange = resolveSourceRange(l),
      p = axis === "row" ? sourceRange.firstRow : sourceRange.firstCol,
      m = axis === "row" ? sourceRange.lastRow : sourceRange.lastCol;
    let v = p,
      g = m;
    if (remove) {
      const b = n + count - 1;
      (n <= p &&
        b >= p &&
        throwWorkbookError(
          `This edit would remove a ${c} header. Move or recreate that structure first.`,
        ),
        b < p
          ? ((v -= count), (g -= count))
          : n <= m &&
            (axis === "column" &&
              throwWorkbookError(
                `Deleting columns inside a ${c} would invalidate its fields. Resize or recreate it first.`,
              ),
            h &&
              b >= m &&
              throwWorkbookError(
                "This edit would remove a table totals row. Turn off table totals first.",
              ),
            (g -= Math.min(m, b) - n + 1)));
    } else
      n <= p
        ? ((v += count), (g += count))
        : n <= m &&
          (axis === "column" &&
            throwWorkbookError(
              `Inserting columns inside a ${c} requires explicit field reconciliation. Resize or recreate it first.`,
            ),
          (g += count));
    return (
      g >= s &&
        throwWorkbookError(
          `The structural edit would push a ${c} beyond the worksheet boundary.`,
        ),
      axis === "row"
        ? ((sourceRange.firstRow = v), (sourceRange.lastRow = g))
        : ((sourceRange.firstCol = v), (sourceRange.lastCol = g)),
      measureRangeArea(sourceRange) > workbookLimits.maxCells &&
        throwWorkbookError(
          "The shifted table or analysis exceeds the 100,000-cell range budget.",
        ),
      `${Cn(sourceRange.firstRow, sourceRange.firstCol)}:${Cn(sourceRange.lastRow, sourceRange.lastCol)}`
    );
  }
  for (const sheet of d.sheets) {
    if (sheet.id === sheetId) {
      for (const c of sheet.tables ?? [])
        c.range = u(c.range, "table", c.showTotals);
      for (const c of [...(sheet.pivots ?? []), ...(sheet.queries ?? [])])
        c.outputRange = u(c.outputRange, "analysis result");
    }
    for (const c of [...(sheet.pivots ?? []), ...(sheet.queries ?? [])])
      c.source.sheetId === sheetId &&
        (c.source.range = u(c.source.range, "analysis source"));
  }
  return d;
}
function tableDataBodyRange(e) {
  const rangeAddress = parseRangeAddress(e.range);
  return {
    ...rangeAddress,
    firstRow: rangeAddress.firstRow + 1,
    lastRow: rangeAddress.lastRow - (e.showTotals ? 1 : 0),
  };
}
function rewriteStructuredReferences(workbook, r, t, address) {
  const o = /^([A-Za-z_][A-Za-z0-9_.]*)?(\[.+\])$/.exec(t);
  if (!o) return null;
  let a;
  if (o[1])
    for (const sheet of workbook.sheets) {
      const table = sheet.tables?.find(
        (table) => table.name.toLowerCase() === o[1].toLowerCase(),
      );
      if (table) {
        a = { sheet: sheet, table: table };
        break;
      }
    }
  else if (address) {
    const sheet = workbook.sheets.find((sheet) => sheet.id === r),
      y = _t(address),
      table = sheet?.tables?.find((table) =>
        Dp(parseRangeAddress(table.range), y.row, y.col),
      );
    sheet && table && (a = { sheet: sheet, table: table });
  }
  if (!a) return null;
  const { sheet, table } = a,
    rangeAddress = parseRangeAddress(table.range),
    l = tableDataBodyRange(table),
    c = o[2]
      .replace(/^\[|\]$/g, "")
      .replace(/[\[\]]/g, "")
      .split(/[:,]/)
      .map((b) => b.trim()),
    h = c.find((b) => b.startsWith("#"))?.toLowerCase(),
    f = c.some((b) => b.startsWith("@")) || h === "#this row";
  if (
    c.some(
      (b) =>
        b.startsWith("#") &&
        !["#all", "#data", "#headers", "#totals", "#this row"].includes(
          b.toLowerCase(),
        ),
    )
  )
    return null;
  if (
    (h === "#all" && Object.assign(l, rangeAddress),
    h === "#headers" &&
      Object.assign(l, rangeAddress, { lastRow: rangeAddress.firstRow }),
    h === "#totals")
  ) {
    if (!table.showTotals) return null;
    Object.assign(l, rangeAddress, { firstRow: rangeAddress.lastRow });
  }
  if (f) {
    if (!address || sheet.id !== r) return null;
    const b = _t(address);
    if (b.row < l.firstRow || b.row > l.lastRow) return null;
    ((l.firstRow = b.row), (l.lastRow = b.row));
  }
  const p = c.filter((b) => !b.startsWith("#")).map((b) => b.replace(/^@/, ""));
  if (p.length) {
    if (p.length > 2) return null;
    const b = p.map((y) =>
      table.columns.findIndex(
        (column) => column.name.toLowerCase() === y.toLowerCase(),
      ),
    );
    if (b.some((y) => y < 0)) return null;
    ((l.firstCol = rangeAddress.firstCol + Math.min(...b)),
      (l.lastCol = rangeAddress.firstCol + Math.max(...b)));
  }
  return l.firstRow > l.lastRow
    ? null
    : {
        sheetId: sheet.id,
        range: `${Cn(l.firstRow, l.firstCol)}:${Cn(l.lastRow, l.lastCol)}`,
      };
}
function buildTableTotalCell(e, r) {
  const t = tableDataBodyRange(e),
    n = e.columns[r].name,
    o = Object.hasOwn(e.totals, n) ? e.totals[n] : undefined;
  return o
    ? t.lastRow < t.firstRow
      ? { raw: o === "count" || o === "sum" ? 0 : null }
      : {
          raw: `=${o === "count" ? "COUNTA" : o.toUpperCase()}(${Cn(t.firstRow, t.firstCol + r)}:${Cn(t.lastRow, t.firstCol + r)})`,
        }
    : { raw: r === 0 ? "Total" : null };
}
function applyTableStyleToCells(sheetOrThrow, r) {
  const sourceRange = resolveSourceRange(r.range, sheetOrThrow),
    [n, o] = tablePaletteColors[r.style];
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
      const d = Cn(firstRow, firstCol),
        u = sheetOrThrow.cells[d] ?? { raw: null },
        l = firstRow === sourceRange.firstRow,
        c = r.showTotals && firstRow === sourceRange.lastRow,
        h = l
          ? {
              ...u,
              raw: r.columns[firstCol - sourceRange.firstCol].name,
              literal: r.columns[
                firstCol - sourceRange.firstCol
              ].name.startsWith("=")
                ? true
                : undefined,
            }
          : c
            ? {
                ...u,
                ...buildTableTotalCell(r, firstCol - sourceRange.firstCol),
                literal: undefined,
              }
            : u;
      sheetOrThrow.cells[d] = {
        ...h,
        style: {
          ...h.style,
          bold: l || c,
          color: l ? "#ffffff" : "#222222",
          fill: l
            ? n
            : c || (firstRow - sourceRange.firstRow) % 2 === 0
              ? o
              : "#ffffff",
          ...(c ? { border: "bottom" } : {}),
        },
      };
    }
}
function filterSourceRows(pivotSource, r) {
  const t = r.map((n) => ({
    ...n,
    index: findSourceFieldIndex(pivotSource.fields, n.field),
  }));
  return pivotSource.rows.filter((row) =>
    t.every((o) => {
      const a = row.values[o.index],
        s = o.value ?? null;
      switch (o.operator) {
        case "blank":
          return isNullOrEmptyValue(a);
        case "notBlank":
          return !isNullOrEmptyValue(a);
        case "contains":
          return String(a ?? "")
            .toLocaleLowerCase()
            .includes(String(s ?? "").toLocaleLowerCase());
        case "equals":
          return (
            String(a ?? "").toLocaleLowerCase() ===
            String(s ?? "").toLocaleLowerCase()
          );
        case "notEquals":
          return (
            String(a ?? "").toLocaleLowerCase() !==
            String(s ?? "").toLocaleLowerCase()
          );
        default:
          return (
            isFiniteNumber(s) ||
              throwWorkbookError(
                "Numeric comparisons require a numeric filter value.",
              ),
            isFiniteNumber(a)
              ? o.operator === "gt"
                ? a > s
                : o.operator === "gte"
                  ? a >= s
                  : o.operator === "lt"
                    ? a < s
                    : a <= s
              : false
          );
      }
    }),
  );
}
function collectRangeValues(e, r) {
  const rangeAddress = parseRangeAddress(r);
  return computePivotRefreshFingerprint(
    Object.entries(e.cells)
      .filter(([n]) => {
        const o = _t(n);
        return Dp(rangeAddress, o.row, o.col);
      })
      .sort(([n], [o]) => n.localeCompare(o))
      .map(([n, cell]) => [n, cell.raw, !!cell.literal]),
  );
}
function buildTableCellEntries(e, sheet) {
  for (const t of sheet.tables ?? []) {
    const sourceRange = resolveSourceRange(t.range, sheet),
      o = tableDataBodyRange(t),
      rows = [];
    for (let firstRow = o.firstRow; firstRow <= o.lastRow; firstRow++) {
      const u = [];
      for (let firstCol = o.firstCol; firstCol <= o.lastCol; firstCol++)
        u.push(e.getCell(Cn(firstRow, firstCol)).value);
      rows.push(u);
    }
    const s = t.showTotals
      ? t.columns.map(
          (column, u) =>
            e.getCell(Cn(sourceRange.lastRow, sourceRange.firstCol + u)).value,
        )
      : [];
    (e.addTable({
      name: t.name,
      ref: Cn(sourceRange.firstRow, sourceRange.firstCol),
      headerRow: true,
      totalsRow: t.showTotals,
      style: { theme: tablePaletteStyleNames[t.style], showRowStripes: true },
      columns: t.columns.map((column, u) => ({
        name: column.name,
        filterButton: true,
        ...(t.showTotals
          ? {
              totalsRowFunction: Object.hasOwn(t.totals, column.name)
                ? t.totals[column.name]
                : "none",
              ...(u === 0 && !Object.hasOwn(t.totals, column.name)
                ? {
                    totalsRowLabel: String(buildTableTotalCell(t, 0).raw ?? ""),
                  }
                : {}),
            }
          : {}),
      })),
      rows: rows,
    }),
      s.forEach((d, u) => {
        e.getCell(Cn(sourceRange.lastRow, sourceRange.firstCol + u)).value = d;
      }));
  }
}
function rewriteFormulaReferences(e, workbook, sheetId, address, o = true) {
  let a = "";
  for (let c = 0; c < e.length;) {
    const h = c,
      f = e[c];
    if (f === '"' || f === "'") {
      for (c++; c < e.length;)
        if (e[c++] === f)
          if (e[c] === f) c++;
          else break;
      a += f === '"' ? '""' : e.slice(h, c);
      continue;
    }
    const p = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(e.slice(c));
    if (p) {
      c += p[0].length;
      let m = 1;
      for (; c < e.length && m;)
        (e[c] === "[" && m++, e[c] === "]" && m--, c++);
      const v = e.slice(h, c);
      a +=
        !m && rewriteStructuredReferences(workbook, sheetId, v, address)
          ? "TABLETOKEN"
          : v;
      continue;
    }
    a += e[c++];
  }
  const s = (c, h, f) => ({
    sheetId: sheetId,
    address: address,
    kind: c,
    error: h,
    message: f,
  });
  if (
    /\[[^\]]+\][^!+*/^&=<>(),;{}]*!/.test(a) ||
    /\|/.test(a) ||
    /(?:https?:|file:|\\\\)/i.test(a)
  )
    return s(
      "external",
      "#BLOCKED!",
      "External workbook, path, network and DDE references are not supported. Use an authorized local worksheet reference.",
    );
  const d = `(?:'(?:[^']|'')*'|[^\\s!+*/^&=<>(),;:{}\\[\\]\\"'%\\-]+)`;
  if (
    /'(?:[^']|'')*:(?:[^']|'')*'\s*!/.test(a) ||
    new RegExp(`${d}\\s*:\\s*${d}\\s*!`).test(a)
  )
    return s(
      "three-dimensional",
      "#VALUE!",
      "3-D references across a worksheet interval are not supported. List the individual worksheet references explicitly.",
    );
  const l = a
    .replace(/'(?:[^']|'')*'\s*!/g, "Sheet!")
    .replace(
      new RegExp(
        "(?<![\\p{L}\\p{N}_.\\]])\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*(?![\\p{L}\\p{N}_.])",
        "gu",
      ),
      "CELLREF",
    );
  if (
    o &&
    (new RegExp(
      "(?<![\\p{L}\\p{N}_.])R(?:\\d+|\\[[+-]?\\d+\\])?C(?:\\d+|\\[[+-]?\\d+\\])?(?![\\p{L}\\p{N}_.]|\\s*[!(])",
      "iu",
    ).test(l) ||
      new RegExp(
        "(?<![\\p{L}\\p{N}_.])[RC]\\[[+-]?\\d+\\](?![\\p{L}\\p{N}_.])",
        "iu",
      ).test(l))
  )
    return s(
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
