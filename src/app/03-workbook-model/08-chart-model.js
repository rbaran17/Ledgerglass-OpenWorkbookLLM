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
import { workbookLimits } from "../02-financial-functions.js";
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
import { _t, aP, parseRangeAddress, oP } from "./07-conditional-formatting.js";

function serializeDefinedNames(e, r) {
  const t = validateDefinedName(r.definedNames, r);
  oP(t);
  const { builtinXml } = aP(e);
  let o = builtinXml;
  for (const a of t) {
    if (a.name.toLowerCase().startsWith("_xlnm."))
      return throwXlsxDefinedNameError(
        `Reserved built-in name "${a.name}" cannot be supplied as a user name.`,
      );
    (validateNameFormulaText(a.refersTo),
      validateNameSheetReferences(a.refersTo, r));
    const s = a.scopeSheetId
      ? r.sheets.findIndex((d) => d.id === a.scopeSheetId)
      : undefined;
    o += `<definedName name="${validateXmlSafeName(a.name, true)}"${s !== undefined ? ` localSheetId="${s}"` : ""}${a.comment !== undefined ? ` comment="${validateXmlSafeName(a.comment, true)}"` : ""}>${validateXmlSafeName(a.refersTo.slice(1))}</definedName>`;
  }
  return insertDefinedNamesIntoWorkbookXml(e, o);
}
function inspectChartSource(e, r, t) {
  var I, O;
  (r.kind !== "chart" || !supportedChartKinds.includes(r.chart)) &&
    failWithMessage("Choose a supported chart type.");
  const n = r.chart,
    o = r.range ?? "",
    a = parseRangeAddress(o),
    s = validateChartSettings(r.chartOptions);
  (a.lastRow >= e.rowCount ||
    a.lastCol >= e.colCount ||
    (a.lastRow - a.firstRow + 1) * (a.lastCol - a.firstCol + 1) > 500) &&
    failWithMessage(
      "Select at most 500 cells inside this worksheet for a chart.",
    );
  const d = a.firstRow + (r.headers ? 1 : 0),
    u = a.firstCol + (a.lastCol > a.firstCol ? 1 : 0);
  a.lastCol - u + 1 > 20 &&
    failWithMessage(
      "Chart inspection supports at most 20 value columns; narrow the source range.",
    );
  const l = (z, V) => {
      const D = `${columnIndexToLetters(V)}${z + 1}`,
        B = e.cells[D] ?? { raw: null };
      return {
        address: D,
        raw: B.raw,
        value: t(e.id, D),
        ...(B.literal ? { literal: true } : {}),
        ...(B.note ? { note: B.note } : {}),
      };
    },
    c = Array.from({ length: a.lastCol - u + 1 }, (z, V) => {
      const D = u + V,
        B = s.series?.find((T) => T.column === D + 1),
        _ = r.headers ? l(a.firstRow, D).value : columnIndexToLetters(D);
      return {
        column: D + 1,
        name: B?.name ?? String(_ ?? columnIndexToLetters(D)),
        kind: B?.kind ?? (n === "combo" && V > 0 ? "line" : "column"),
        included: s.series === undefined || !!B,
      };
    });
  ((I = s.series) != null &&
    I.some((z) => z.column < u + 1 || z.column > a.lastCol + 1) &&
    failWithMessage(
      "A selected chart series no longer exists in its source range. Re-select series.",
    ),
    (O = s.includedRows) != null &&
      O.some((z) => z < d + 1 || z > a.lastRow + 1) &&
      failWithMessage(
        "A chart filter row is outside its source range. Reset the row filter.",
      ));
  const h = cn(e, "row"),
    f = (z) => h.some((V) => z >= V.start && z <= V.end),
    p = Array.from({ length: Math.max(0, a.lastRow - d + 1) }, (z, V) => {
      const D = d + V,
        B = l(D, a.firstCol);
      return {
        row: D + 1,
        label: a.firstCol < u ? String(B.value ?? "") : B.address,
        category: B,
        included:
          s.includedRows === undefined || s.includedRows.includes(D + 1),
        hidden: f(D),
        values: c.map((_) => l(D, _.column - 1)),
      };
    }),
    m = c.map((z, V) => (z.included ? V : -1)).filter((z) => z >= 0),
    v = p.filter((z) => z.included),
    g = [],
    b = (z, V) =>
      g.push({ severity: "error", message: z, ...(V ? { address: V } : {}) });
  (p.length || b("Select at least one data row below the chart header."),
    v.length || b("The chart filter includes no source rows."),
    m.length || b("Select at least one value series."),
    ["pie", "donut", "waterfall", "histogram", "sparkline"].includes(n) &&
      m.length !== 1 &&
      b(
        `${n} requires exactly one selected series. Other source series remain available in inspection.`,
      ));
  const y = (z) => z === null || z === "";
  for (let z = 0; z < c.length; z++) {
    let V = 0;
    for (const D of p) {
      const B = D.values[z];
      if (y(B.value)) continue;
      const _ = typeof B.value == "number" && Number.isFinite(B.value),
        C = D.included && c[z].included;
      if (!_) {
        g.push({
          severity: C ? "error" : "warning",
          message: `${B.address}: ${c[z].name} contains a nonnumeric or invalid calculated value${C ? "" : " outside the plotted selection"}.`,
          address: B.address,
        });
        continue;
      }
      C &&
        (V++,
        s.yScale === "log" &&
          Number(B.value) <= 0 &&
          b(`${B.address}: log-axis values must be positive.`, B.address),
        ["pie", "donut"].includes(n) &&
          Number(B.value) < 0 &&
          b(`${B.address}: pie/donut values cannot be negative.`, B.address));
    }
    c[z].included &&
      !V &&
      b(`${c[z].name}: the selected series is empty or has no numeric values.`);
  }
  if (n === "scatter") {
    a.firstCol === u &&
      b(
        "Scatter requires a numeric X column followed by at least one Y column.",
      );
    for (const V of v) {
      if (!V.values.some((B, _) => c[_].included && !y(B.value))) continue;
      const D = V.category.value;
      typeof D != "number" || !Number.isFinite(D)
        ? b(
            `${V.category.address}: scatter X values must be numeric.`,
            V.category.address,
          )
        : s.xScale === "log" &&
          D <= 0 &&
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
    ["pie", "donut", "waterfall", "histogram", "sparkline"].includes(n) &&
    b(`Log axes are not supported for ${n}; choose a linear scale.`),
    ["pie", "donut"].includes(n) &&
      m.length === 1 &&
      !v.some((z) => Number(z.values[m[0]].value) > 0) &&
      b("Pie/donut requires at least one positive value."));
  const w = {
      sourceRows: p.length,
      includedRows: v.length,
      excludedRows: p.length - v.length,
      hiddenRows: p.filter((z) => z.hidden).length,
      availableSeries: c.length,
      selectedSeries: m.length,
      policy:
        "All source rows, including worksheet-hidden or filtered rows, are inspected. Only explicit chart row/series selections affect the plotted population. Blank cells are gaps; invalid values are diagnostics, never zero.",
    },
    E = `${e.name}!${o} · ${v.length}/${p.length} rows · ${m.length}/${c.length} series · ${w.hiddenRows} hidden rows in source`,
    A = {
      version: 1,
      objectId: r.id,
      kind: n,
      title: r.label,
      altText: s.altText?.trim() || `${r.label || n}. ${E}. ${w.policy}`,
      source: { sheetId: e.id, sheetName: e.name, range: o },
      axes: {
        x: { title: s.xTitle ?? "", scale: s.xScale ?? "linear" },
        y: {
          title: s.yTitle ?? (n === "histogram" ? "Frequency (records)" : ""),
          scale: s.yScale ?? "linear",
        },
        units: s.units ?? "",
      },
      options: s,
      population: w,
      scope: E,
      series: c,
      points: p,
      diagnostics: g,
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
  if (n === "histogram" && k.length) {
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
  if (n === "waterfall" && k.length) {
    let z = 0;
    ((A.waterfall = k.map((V) => {
      const D = z;
      return ((z += V.value), { ...V, start: D, end: z });
    })),
      A.waterfall.push({
        label: "Total",
        row: null,
        start: 0,
        end: z,
        value: z,
      }),
      (!Number.isFinite(z) ||
        A.waterfall.some((V) => !Number.isFinite(V.end))) &&
        b("Waterfall cumulative values exceed the finite calculation range."));
  }
  ["pie", "donut"].includes(n) &&
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
function measureChartSelection(e, r) {
  const t = parseRangeAddress(e);
  if (
    t.lastRow >= r.rowCount ||
    t.lastCol >= r.colCount ||
    (t.lastRow - t.firstRow + 1) * (t.lastCol - t.firstCol + 1) > Yne
  )
    throw new Error(
      "Select at most 500 cells inside this worksheet for a chart.",
    );
  return t;
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
function setSheetDrawingObjects(e, r, t, n = {}) {
  assertDrawingPreviewCurrent(e, r, n);
  const o = e.sheets.find((s) => s.id === r),
    a = { ...o };
  return (
    validateDrawingObjects(t, a),
    JSON.stringify(o.objects ?? []) === JSON.stringify(a.objects ?? [])
      ? e
      : { ...e, sheets: e.sheets.map((s) => (s.id !== r ? s : a)) }
  );
}
const xf = (e, r, t = 255) =>
    typeof e == "string" && e.trim() && e.length <= t
      ? e
      : throwWorkbookError(`${r} must contain 1–${t} characters.`),
  requireMetadataId = (e) => xf(e, "Metadata ID", 100),
  toA1RangeAddress = (e) => {
    const r = parseRangeAddress(e);
    return `${columnIndexToLetters(r.firstCol)}${r.firstRow + 1}:${columnIndexToLetters(r.lastCol)}${r.lastRow + 1}`;
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
function resolveSourceRange(e, r) {
  const t = parseRangeAddress(xf(e, "Source range", 50));
  return (
    measureRangeArea(t) > workbookLimits.maxCells &&
      throwWorkbookError(
        "Table, query and pivot ranges are limited to 100,000 cells including blank gaps. Select the used data range.",
      ),
    r &&
      (t.lastRow >= r.rowCount || t.lastCol >= r.colCount) &&
      throwWorkbookError("Range exceeds worksheet bounds."),
    t
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
function normalizeSourceReference(e) {
  isPlainDataObject(e) ||
    throwWorkbookError("A source worksheet and range are required.");
  const r = {
    sheetId: requireMetadataId(e.sheetId),
    range: toA1RangeAddress(xf(e.range, "Source range", 50)),
    ...(e.tableId !== undefined
      ? { tableId: requireMetadataId(e.tableId) }
      : {}),
  };
  return (resolveSourceRange(r.range), r);
}
function normalizeSourceFilters(e) {
  return e === undefined
    ? []
    : ((!Array.isArray(e) || e.length > 100) &&
        throwWorkbookError("Use at most 100 source filters."),
      e.map((r) => {
        (!isPlainDataObject(r) ||
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
          ].includes(String(r.operator))) &&
          throwWorkbookError("Invalid source filter.");
        const t = r.value;
        return (
          t != null &&
            !(typeof t == "string" && t.length <= 32767) &&
            typeof t != "boolean" &&
            !(typeof t == "number" && Number.isFinite(t)) &&
            throwWorkbookError(
              "Filter values must be finite numbers, text, booleans or blank.",
            ),
          {
            field: validateColumnName(r.field),
            operator: r.operator,
            ...(t !== undefined ? { value: t } : {}),
          }
        );
      }));
}
function normalizeOutputDestination(e) {
  if (e === undefined) return;
  isPlainDataObject(e) || throwWorkbookError("Invalid output destination.");
  const r =
      e.address === undefined ? "A1" : xf(e.address, "Output address", 20),
    t = _t(r);
  return {
    address: Cn(t.row, t.col),
    ...(e.sheetId !== undefined
      ? { sheetId: requireMetadataId(e.sheetId) }
      : {}),
    ...(e.sheetName !== undefined
      ? { sheetName: xf(e.sheetName, "Output sheet name", 31) }
      : {}),
  };
}
function normalizePivotDefinition(e) {
  (isPlainDataObject(e) || throwWorkbookError("Invalid pivot definition."),
    (!Array.isArray(e.values) || !e.values.length || e.values.length > 100) &&
      throwWorkbookError("Choose between 1 and 100 pivot value fields."));
  const r = e.values.map(
    (o) => (
      (!isPlainDataObject(o) ||
        !pivotAggregateOperations.includes(o.operation)) &&
        throwWorkbookError(
          "Invalid pivot aggregation. Choose sum, count, average, min, max or distinctCount.",
        ),
      o.showAs !== undefined &&
        !pivotShowValuesAsModes.includes(o.showAs) &&
        throwWorkbookError("Choose a supported Show values as setting."),
      {
        field: validateColumnName(o.field),
        operation: o.operation,
        ...(o.name !== undefined ? { name: validateColumnName(o.name) } : {}),
        ...(o.showAs !== undefined ? { showAs: o.showAs } : {}),
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
  const t = (e.dateGroups ?? []).map(
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
  (new Set(t.map((o) => o.field)).size !== t.length &&
    throwWorkbookError("Choose one date grouping per source field."),
    e.calculatedFields !== undefined &&
      (!Array.isArray(e.calculatedFields) || e.calculatedFields.length > 100) &&
      throwWorkbookError("Use at most 100 calculated fields."));
  const n = (e.calculatedFields ?? []).map(
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
    new Set(n.map((o) => o.name.toLowerCase())).size !== n.length &&
      throwWorkbookError("Calculated field names must be unique."),
    {
      name: xf(e.name, "Pivot name", 100),
      source: normalizeSourceReference(e.source),
      rowFields: validateFieldList(e.rowFields, "Row fields"),
      columnFields: validateFieldList(e.columnFields, "Column fields"),
      values: r,
      filters: normalizeSourceFilters(e.filters),
      grandTotals: e.grandTotals !== false,
      ...(e.subtotals !== undefined ? { subtotals: e.subtotals } : {}),
      ...(e.dateGroups !== undefined ? { dateGroups: t } : {}),
      ...(e.calculatedFields !== undefined ? { calculatedFields: n } : {}),
      ...normalizePivotControls(e),
      destination: normalizeOutputDestination(e.destination),
    }
  );
}
function normalizeQueryDefinition(e) {
  isPlainDataObject(e) || throwWorkbookError("Invalid query definition.");
  const r = validateFieldList(e.fields, "Selected fields");
  r.length || throwWorkbookError("Choose at least one output field.");
  const t = {};
  if (e.rename !== undefined) {
    isPlainDataObject(e.rename) ||
      throwWorkbookError("Rename fields must be a name map.");
    for (const [s, d] of Object.entries(e.rename))
      (r.includes(s) ||
        throwWorkbookError("Only selected fields can be renamed."),
        Object.defineProperty(t, s, {
          value: validateColumnName(d),
          enumerable: true,
          writable: true,
          configurable: true,
        }));
  }
  let n;
  (e.sort !== undefined &&
    ((!isPlainDataObject(e.sort) ||
      !["ascending", "descending"].includes(String(e.sort.direction))) &&
      throwWorkbookError("Invalid query sort."),
    (n = {
      field: validateColumnName(e.sort.field),
      direction: e.sort.direction,
    })),
    e.derived !== undefined &&
      (!Array.isArray(e.derived) || e.derived.length > 100) &&
      throwWorkbookError("Use at most 100 derived fields."));
  const o = (e.derived ?? []).map(
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
      ...r.map((s) => (Object.hasOwn(t, s) ? t[s] : s)),
      ...o.map((s) => s.name),
    ];
  return (
    new Set(a.map((s) => s.toLowerCase())).size !== a.length &&
      throwWorkbookError("Output column names must be unique."),
    {
      name: xf(e.name, "Query name", 100),
      source: normalizeSourceReference(e.source),
      fields: r,
      rename: t,
      filters: normalizeSourceFilters(e.filters),
      ...(n ? { sort: n } : {}),
      derived: o,
      destination: normalizeOutputDestination(e.destination),
    }
  );
}
function normalizeTableDefinitions(e, r) {
  if (isPlainDataObject(e)) {
    if (e.tables !== undefined) {
      (!Array.isArray(e.tables) || e.tables.length > workbookLimits.maxCells) &&
        throwWorkbookError("Tables must be a bounded list.");
      const t = new Set(),
        n = new Set();
      r.tables = e.tables.map((s) => {
        var p;
        (!isPlainDataObject(s) ||
          !Array.isArray(s.columns) ||
          typeof s.showTotals != "boolean" ||
          !["blue", "green", "gray"].includes(String(s.style)) ||
          !isPlainDataObject(s.totals)) &&
          throwWorkbookError("Invalid table metadata.");
        const d = requireMetadataId(s.id),
          u = validateTableName(s.name),
          l = toA1RangeAddress(xf(s.range, "Table range", 50)),
          c = resolveSourceRange(l, r);
        ((p = r.merges) != null &&
          p.some((m) => rangesIntersect(c, parseRangeAddress(m))) &&
          throwWorkbookError(
            "A table cannot overlap merged cells. Unmerge the source first.",
          ),
          (t.has(d) || n.has(u.toLowerCase())) &&
            throwWorkbookError("Table IDs and names must be unique."),
          t.add(d),
          n.add(u.toLowerCase()),
          (s.columns.length !== c.lastCol - c.firstCol + 1 ||
            (s.showTotals && c.lastRow <= c.firstRow)) &&
            throwWorkbookError(
              "Table columns or totals do not fit the table range.",
            ));
        const h = s.columns.map(
          (m) => (
            isPlainDataObject(m) || throwWorkbookError("Invalid table column."),
            {
              id: requireMetadataId(m.id),
              name: validateColumnName(m.name),
              ...(m.dataType !== undefined
                ? { dataType: validateTableColumnType(m.dataType) }
                : {}),
              ...(m.calculatedFormula !== undefined
                ? {
                    calculatedFormula: validateCalculatedColumnFormula(
                      m.calculatedFormula,
                      true,
                    ),
                  }
                : {}),
            }
          ),
        );
        (new Set(h.map((m) => m.id)).size !== h.length ||
          new Set(h.map((m) => m.name.toLowerCase())).size !== h.length) &&
          throwWorkbookError("Table column IDs and names must be unique.");
        const f = {};
        for (const [m, v] of Object.entries(s.totals))
          ((!h.some((g) => g.name === m) ||
            !basicAggregateOperations.includes(v)) &&
            throwWorkbookError("Invalid table total field or operation."),
            Object.defineProperty(f, m, {
              value: v,
              enumerable: true,
              writable: true,
              configurable: true,
            }));
        return {
          id: d,
          name: u,
          range: l,
          columns: h,
          style: s.style,
          showTotals: s.showTotals,
          totals: f,
        };
      });
      const o = new Set();
      let a = 0;
      for (const s of r.tables) {
        const d = parseRangeAddress(s.range);
        (a += measureRangeArea(d)) > workbookLimits.maxCells &&
          throwWorkbookError(
            "The total table footprint exceeds 100,000 cells.",
          );
        for (let u = d.firstRow; u <= d.lastRow; u++)
          for (let l = d.firstCol; l <= d.lastCol; l++) {
            const c = Cn(u, l);
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
            u = toA1RangeAddress(xf(a.outputRange, "Output range", 50));
          (resolveSourceRange(u, r),
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
            outputRange: u,
            sourceRevision: Number(a.sourceRevision),
            ...(t === "pivots" && a.refresh !== undefined
              ? { refresh: validatePivotRefreshEvidence(a.refresh) }
              : {}),
          };
        });
      t === "pivots" ? (r.pivots = o) : (r.queries = o);
    }
  }
}
function checkStructuralEditImpact(e, r, t, n, o, a) {
  const s = t === "row" ? workbookLimits.maxRows : workbookLimits.maxCols;
  ((!Number.isInteger(n) ||
    n < 0 ||
    n >= s ||
    !Number.isInteger(o) ||
    o < 1 ||
    n + o > s) &&
    throwWorkbookError("Invalid structural edit bounds."),
    e.sheets.some((l) => {
      return l.tables?.some((h) => h.columns.some((f) => f.calculatedFormula));
    }) &&
      throwWorkbookError(
        "Calculated table formulas require explicit table append/resize. Clear their saved rules (keeping cell formulas) before inserting or deleting worksheet rows/columns, then reapply the rules.",
      ));
  const d = structuredClone(e);
  function u(l, c, h = false) {
    const f = resolveSourceRange(l),
      p = t === "row" ? f.firstRow : f.firstCol,
      m = t === "row" ? f.lastRow : f.lastCol;
    let v = p,
      g = m;
    if (a) {
      const b = n + o - 1;
      (n <= p &&
        b >= p &&
        throwWorkbookError(
          `This edit would remove a ${c} header. Move or recreate that structure first.`,
        ),
        b < p
          ? ((v -= o), (g -= o))
          : n <= m &&
            (t === "column" &&
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
        ? ((v += o), (g += o))
        : n <= m &&
          (t === "column" &&
            throwWorkbookError(
              `Inserting columns inside a ${c} requires explicit field reconciliation. Resize or recreate it first.`,
            ),
          (g += o));
    return (
      g >= s &&
        throwWorkbookError(
          `The structural edit would push a ${c} beyond the worksheet boundary.`,
        ),
      t === "row"
        ? ((f.firstRow = v), (f.lastRow = g))
        : ((f.firstCol = v), (f.lastCol = g)),
      measureRangeArea(f) > workbookLimits.maxCells &&
        throwWorkbookError(
          "The shifted table or analysis exceeds the 100,000-cell range budget.",
        ),
      `${Cn(f.firstRow, f.firstCol)}:${Cn(f.lastRow, f.lastCol)}`
    );
  }
  for (const l of d.sheets) {
    if (l.id === r) {
      for (const c of l.tables ?? [])
        c.range = u(c.range, "table", c.showTotals);
      for (const c of [...(l.pivots ?? []), ...(l.queries ?? [])])
        c.outputRange = u(c.outputRange, "analysis result");
    }
    for (const c of [...(l.pivots ?? []), ...(l.queries ?? [])])
      c.source.sheetId === r &&
        (c.source.range = u(c.source.range, "analysis source"));
  }
  return d;
}
function tableDataBodyRange(e) {
  const r = parseRangeAddress(e.range);
  return {
    ...r,
    firstRow: r.firstRow + 1,
    lastRow: r.lastRow - (e.showTotals ? 1 : 0),
  };
}
function rewriteStructuredReferences(e, r, t, n) {
  const o = /^([A-Za-z_][A-Za-z0-9_.]*)?(\[.+\])$/.exec(t);
  if (!o) return null;
  let a;
  if (o[1])
    for (const b of e.sheets) {
      const y = b.tables?.find(
        (w) => w.name.toLowerCase() === o[1].toLowerCase(),
      );
      if (y) {
        a = { sheet: b, table: y };
        break;
      }
    }
  else if (n) {
    const b = e.sheets.find((E) => E.id === r),
      y = _t(n),
      w = b?.tables?.find((E) => Dp(parseRangeAddress(E.range), y.row, y.col));
    b && w && (a = { sheet: b, table: w });
  }
  if (!a) return null;
  const { sheet, table } = a,
    u = parseRangeAddress(table.range),
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
    (h === "#all" && Object.assign(l, u),
    h === "#headers" && Object.assign(l, u, { lastRow: u.firstRow }),
    h === "#totals")
  ) {
    if (!table.showTotals) return null;
    Object.assign(l, u, { firstRow: u.lastRow });
  }
  if (f) {
    if (!n || sheet.id !== r) return null;
    const b = _t(n);
    if (b.row < l.firstRow || b.row > l.lastRow) return null;
    ((l.firstRow = b.row), (l.lastRow = b.row));
  }
  const p = c.filter((b) => !b.startsWith("#")).map((b) => b.replace(/^@/, ""));
  if (p.length) {
    if (p.length > 2) return null;
    const b = p.map((y) =>
      table.columns.findIndex((w) => w.name.toLowerCase() === y.toLowerCase()),
    );
    if (b.some((y) => y < 0)) return null;
    ((l.firstCol = u.firstCol + Math.min(...b)),
      (l.lastCol = u.firstCol + Math.max(...b)));
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
function applyTableStyleToCells(e, r) {
  const t = resolveSourceRange(r.range, e),
    [n, o] = tablePaletteColors[r.style];
  for (let a = t.firstRow; a <= t.lastRow; a++)
    for (let s = t.firstCol; s <= t.lastCol; s++) {
      const d = Cn(a, s),
        u = e.cells[d] ?? { raw: null },
        l = a === t.firstRow,
        c = r.showTotals && a === t.lastRow,
        h = l
          ? {
              ...u,
              raw: r.columns[s - t.firstCol].name,
              literal: r.columns[s - t.firstCol].name.startsWith("=")
                ? true
                : undefined,
            }
          : c
            ? {
                ...u,
                ...buildTableTotalCell(r, s - t.firstCol),
                literal: undefined,
              }
            : u;
      e.cells[d] = {
        ...h,
        style: {
          ...h.style,
          bold: l || c,
          color: l ? "#ffffff" : "#222222",
          fill: l ? n : c || (a - t.firstRow) % 2 === 0 ? o : "#ffffff",
          ...(c ? { border: "bottom" } : {}),
        },
      };
    }
}
function filterSourceRows(e, r) {
  const t = r.map((n) => ({
    ...n,
    index: findSourceFieldIndex(e.fields, n.field),
  }));
  return e.rows.filter((n) =>
    t.every((o) => {
      const a = n.values[o.index],
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
  const t = parseRangeAddress(r);
  return computePivotRefreshFingerprint(
    Object.entries(e.cells)
      .filter(([n]) => {
        const o = _t(n);
        return Dp(t, o.row, o.col);
      })
      .sort(([n], [o]) => n.localeCompare(o))
      .map(([n, o]) => [n, o.raw, !!o.literal]),
  );
}
function buildTableCellEntries(e, r) {
  for (const t of r.tables ?? []) {
    const n = resolveSourceRange(t.range, r),
      o = tableDataBodyRange(t),
      a = [];
    for (let d = o.firstRow; d <= o.lastRow; d++) {
      const u = [];
      for (let l = o.firstCol; l <= o.lastCol; l++)
        u.push(e.getCell(Cn(d, l)).value);
      a.push(u);
    }
    const s = t.showTotals
      ? t.columns.map((d, u) => e.getCell(Cn(n.lastRow, n.firstCol + u)).value)
      : [];
    (e.addTable({
      name: t.name,
      ref: Cn(n.firstRow, n.firstCol),
      headerRow: true,
      totalsRow: t.showTotals,
      style: { theme: tablePaletteStyleNames[t.style], showRowStripes: true },
      columns: t.columns.map((d, u) => ({
        name: d.name,
        filterButton: true,
        ...(t.showTotals
          ? {
              totalsRowFunction: Object.hasOwn(t.totals, d.name)
                ? t.totals[d.name]
                : "none",
              ...(u === 0 && !Object.hasOwn(t.totals, d.name)
                ? {
                    totalsRowLabel: String(buildTableTotalCell(t, 0).raw ?? ""),
                  }
                : {}),
            }
          : {}),
      })),
      rows: a,
    }),
      s.forEach((d, u) => {
        e.getCell(Cn(n.lastRow, n.firstCol + u)).value = d;
      }));
  }
}
function rewriteFormulaReferences(e, r, t, n, o = true) {
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
      a += !m && rewriteStructuredReferences(r, t, v, n) ? "TABLETOKEN" : v;
      continue;
    }
    a += e[c++];
  }
  const s = (c, h, f) => ({
    sheetId: t,
    address: n,
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
