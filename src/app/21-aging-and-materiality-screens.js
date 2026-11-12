import { throwErrorMessage, failWithMessage } from "./03-workbook-model.js";
import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import {
  buildCompositeKey,
  isFormulaInputValue,
} from "./11-local-workbook-store-and-versions.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  calculateWorkbookWithSettings,
  cn,
  parseRangeAddress,
  resolveDefinedNameValue,
  columnIndexToLetters,
  matrixResultFunctionNames,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  importNativeJsonFile,
  toFileProcessingError,
} from "./07-file-import-and-worker-dispatch.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import { inspectFormulaDependencies } from "./12-consolidation-and-dependency-inspection.js";
import {
  parseAssistantChatRequest,
  validateAssistantProposal,
  LLMRequestError,
  raceWithAbortSignal,
  validateCalculationClock,
} from "./20-calculation-options-and-assistant.js";

async function requestAssistantAnswer(
  adapter,
  r,
  assistantSelectionContext,
  n = {},
) {
  if (n.reliability) {
    const { reliability, ...c } = n;
    return reliability.run("provider", "request", () =>
      requestAssistantAnswer(adapter, r, assistantSelectionContext, c),
    );
  }
  if (
    typeof r != "string" ||
    !r.trim() ||
    r.length > 1e4 ||
    !Array.isArray(assistantSelectionContext?.cells) ||
    assistantSelectionContext.cells.length > workbookLimits.maxLLMCells
  )
    throw new LLMRequestError("invalid_request");
  validateCalculationClock(assistantSelectionContext);
  const o = structuredClone(assistantSelectionContext),
    a = n.conversation ? parseAssistantChatRequest(n.conversation) : undefined;
  if (
    (a?.attachments.length ||
      a?.messages.some((message) => message.attachments?.length)) &&
    !adapter.chat?.attachments
  )
    throw new Error(
      "This provider does not support file attachments. Remove the files or use a chat-enabled host adapter.",
    );
  if (a?.messages.length && !adapter.chat?.history)
    throw new Error(
      "This provider does not support conversation history. Start a new chat or use a chat-enabled host adapter.",
    );
  const s = await raceWithAbortSignal(
    (l) =>
      a
        ? adapter.propose(r, structuredClone(o), l, structuredClone(a))
        : adapter.propose(r, structuredClone(o), l),
    n.signal,
    n.timeoutMs,
  );
  try {
    return validateAssistantProposal(s, o);
  } catch {
    throw new LLMRequestError("malformed_response");
  }
}
const materialityScreenLimits = { sourceCells: 500, pageSize: 20 },
  Tz = new WeakMap();

const cde = 864e5;
function parseStrictIsoDate(e) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(e)) return null;
  const r = Date.parse(`${e}T00:00:00.000Z`);
  return Number.isFinite(r) && new Date(r).toISOString().slice(0, 10) === e
    ? r
    : null;
}
function normalizeDateToIso(e) {
  return typeof e == "string"
    ? parseStrictIsoDate(e) === null
      ? null
      : e
    : typeof e != "number" ||
        !Number.isInteger(e) ||
        e < 1 ||
        e > 2958465 ||
        e === 60
      ? null
      : new Date(Date.UTC(1899, 11, 31) + (e > 60 ? e - 1 : e) * cde)
          .toISOString()
          .slice(0, 10);
}
function normalizeScreenHostScope(workbook, r) {
  return (
    r.includeHidden !== undefined &&
      typeof r.includeHidden != "boolean" &&
      failWithMessage("Include hidden must be boolean."),
    r.readOnly !== undefined &&
      typeof r.readOnly != "boolean" &&
      failWithMessage("Read-only must be boolean."),
    r.allowedSheetIds !== undefined &&
      (!Array.isArray(r.allowedSheetIds) ||
        r.allowedSheetIds.some(
          (allowedSheetId) =>
            typeof allowedSheetId != "string" ||
            !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
        )) &&
      failWithMessage("The materiality screen host context is invalid."),
    {
      includeHidden: !!r.includeHidden,
      allowedSheetIds: [
        ...new Set(
          r.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
        ),
      ].sort(),
    }
  );
}
function normalizeMaterialityConfig(workbook, r) {
  const t = [
    "sheetId",
    "range",
    "headers",
    "amountColumn",
    "currencyColumn",
    "dateColumn",
    "currency",
    "threshold",
    "amountMode",
    "comparator",
    "effectiveFrom",
    "effectiveTo",
  ];
  ((!r ||
    typeof r != "object" ||
    Object.keys(r).some((a) => !t.includes(a)) ||
    t.some((a) => !(a in r))) &&
    failWithMessage(
      "Provide a complete, supported materiality screen configuration.",
    ),
    (typeof r.sheetId != "string" ||
      !workbook.sheets.some((sheet) => sheet.id === r.sheetId)) &&
      failWithMessage("The source worksheet no longer exists."),
    (typeof r.range != "string" ||
      !/^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/.test(
        r.range,
      )) &&
      failWithMessage("Choose one rectangular A1 source range."));
  const rangeAddress = parseRangeAddress(r.range);
  ((rangeAddress.lastRow - rangeAddress.firstRow + 1) *
    (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    materialityScreenLimits.sourceCells &&
    failWithMessage(
      "The materiality screen supports at most 500 selected source cells.",
    ),
    (typeof r.headers != "boolean" ||
      rangeAddress.lastRow - rangeAddress.firstRow + 1 <= Number(r.headers)) &&
      failWithMessage(
        "Choose a source range with at least one data row and valid header handling.",
      ));
  const o = [r.amountColumn, r.currencyColumn, r.dateColumn];
  return (
    (o.some(
      (a) =>
        !Number.isInteger(a) ||
        a < rangeAddress.firstCol ||
        a > rangeAddress.lastCol,
    ) ||
      new Set(o).size !== 3) &&
      failWithMessage(
        "Choose three distinct amount, currency and date columns inside the source range.",
      ),
    (typeof r.threshold != "number" ||
      !Number.isFinite(r.threshold) ||
      r.threshold < 0) &&
      failWithMessage("Threshold must be a finite, nonnegative number."),
    (!["absolute", "signed"].includes(r.amountMode) ||
      !["gt", "gte", "lt", "lte", "eq"].includes(r.comparator)) &&
      failWithMessage(
        "Choose an absolute or signed amount and a supported comparison rule.",
      ),
    (typeof r.currency != "string" || !/^[A-Za-z]{3}$/.test(r.currency)) &&
      failWithMessage(
        "Declare a three-letter currency identifier; no currency conversion is performed.",
      ),
    (typeof r.effectiveFrom != "string" ||
      typeof r.effectiveTo != "string" ||
      parseStrictIsoDate(r.effectiveFrom) === null ||
      parseStrictIsoDate(r.effectiveTo) === null ||
      r.effectiveFrom > r.effectiveTo) &&
      failWithMessage(
        "Choose valid ISO start and end dates in chronological order.",
      ),
    {
      sheetId: r.sheetId,
      range: `${formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol)}`,
      headers: r.headers,
      amountColumn: r.amountColumn,
      currencyColumn: r.currencyColumn,
      dateColumn: r.dateColumn,
      currency: r.currency.toUpperCase(),
      threshold: r.threshold,
      amountMode: r.amountMode,
      comparator: r.comparator,
      effectiveFrom: r.effectiveFrom,
      effectiveTo: r.effectiveTo,
    }
  );
}
function resolveMaterialitySourceRange(workbook, r, screenHostScope) {
  const n = workbook.sheets.find((sheet) => sheet.id === r.sheetId);
  (!n || !screenHostScope.allowedSheetIds.includes(n.id)) &&
    failWithMessage("The source is outside the authorized host context.");
  const rangeAddress = parseRangeAddress(r.range);
  return (
    !screenHostScope.includeHidden &&
      (n.hidden ||
        n.veryHidden ||
        cn(n, "row").some(
          (a) =>
            a.start <= rangeAddress.lastRow && a.end >= rangeAddress.firstRow,
        ) ||
        cn(n, "column").some(
          (a) =>
            a.start <= rangeAddress.lastCol && a.end >= rangeAddress.firstCol,
        )) &&
      failWithMessage(
        "The source range includes hidden content. Explicitly include hidden content only within authorized host scope.",
      ),
    (n.merges ?? []).some((a) => {
      const s = parseRangeAddress(a);
      return (
        s.firstRow <= rangeAddress.lastRow &&
        s.lastRow >= rangeAddress.firstRow &&
        s.firstCol <= rangeAddress.lastCol &&
        s.lastCol >= rangeAddress.firstCol
      );
    }) &&
      failWithMessage(
        "Unmerge the source range before screening so every record has unambiguous fields.",
      ),
    n
  );
}
function inspectScreenFormulaCells(
  workbook,
  materialitySourceRange,
  t,
  screenHostScope,
) {
  const o = t.filter((g) =>
      isFormulaInputValue(materialitySourceRange.cells[g]),
    ),
    a = new Map(
      t.map((g) => [g, materialitySourceRange.cells[g]?.raw ?? null]),
    );
  if (!o.length) return a;
  const s = JSON.stringify(workbook);
  let d = "";
  for (
    let g = 0;
    g < 4 &&
    ((d = `_ms_${crypto.randomUUID().replace(/-/g, "").slice(0, 24)}`),
    !!s.toLowerCase().includes(d.toLowerCase()));
    g++
  )
    d = "";
  d || failWithMessage("A private calculation context could not be isolated.");
  const u = {
    id: d,
    name: d,
    rowCount: workbookLimits.maxRows,
    colCount: workbookLimits.maxCols,
    cells: {
      A1: {
        raw: `=SUM(${o.map((g) => `'${materialitySourceRange.name.replace(/'/g, "''")}'!${g}`).join(",")})`,
      },
    },
  };
  String(u.cells.A1.raw).length > 8192 &&
    failWithMessage(
      "The selected formulas exceed the bounded calculation inspection length.",
    );
  const l = inspectFormulaDependencies(
      {
        ...workbook,
        sheets: [
          u,
          ...workbook.sheets.map((sheet) =>
            sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
          ),
        ],
      },
      {
        sheetId: u.id,
        address: "A1",
        direction: "precedents",
        depth: 8,
        includeHidden: screenHostScope.includeHidden,
        allowedSheetIds: [...screenHostScope.allowedSheetIds, u.id],
        limits: { maxNodes: 1e3, maxFormulaCells: 5e3, maxReferences: 1e4 },
      },
    ),
    c = new Set(l.nodes.map((node) => node.key));
  (l.coverage.truncated ||
    l.coverage.depthLimited ||
    l.cycles.length ||
    l.issues.some(
      (issue) => issue.from === l.target.key || c.has(issue.from),
    )) &&
    failWithMessage(
      "Formula fields require a complete supported static dependency closure within the visible authorized context (1,000 cells, depth 8, 5,000 scanned formulas).",
    );
  const h = new Set([
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
    "EVALUATE",
    "GET.CELL",
  ]);
  for (const node of l.nodes) {
    const cell = workbook.sheets.find((sheet) => sheet.id === node.sheetId)
      ?.cells[node.address];
    if (
      ((!cell || cell.raw === null) &&
        failWithMessage(
          "Formula dependencies must be populated scalar cells; ambiguous blank or spill-child dependencies cannot be screened.",
        ),
      isFormulaInputValue(cell))
    ) {
      const y = resolveDefinedNameValue(
        String(cell.raw),
        workbook,
        node.sheetId,
      ).replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, "");
      (/[{}#]/.test(y) ||
        [...y.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)].some((w) =>
          h.has(
            w[1]
              .toUpperCase()
              .replace(/^_XLFN\./, "")
              .replace(/^_XLWS\./, ""),
          ),
        )) &&
        failWithMessage(
          "The screen excludes volatile, dynamic-reference and array/spill formulas. Use supported scalar values or formulas.",
        );
    }
  }
  const f = {
      ...workbook,
      sheets: workbook.sheets
        .filter((sheet) => screenHostScope.allowedSheetIds.includes(sheet.id))
        .map((g) => ({
          ...g,
          cells: Object.fromEntries(
            Object.entries(g.cells).filter(([b]) =>
              c.has(buildCompositeKey(g.id, b)),
            ),
          ),
        })),
    },
    workbookWithSettings = calculateWorkbookWithSettings(f);
  for (const node of l.nodes) {
    const b = workbookWithSettings(node.sheetId, node.address);
    (workbookWithSettings.getOperationalErrors().length &&
      failWithMessage(
        "An operational calculation limit or cycle prevents screening, even when a formula handles the visible error.",
      ),
      workbookWithSettings.getSpillRange(node.sheetId, node.address) &&
        failWithMessage(
          "The screen requires scalar fields; a formula produced an array/spill result.",
        ),
      isFormulaInputValue(
        workbook.sheets.find((sheet) => sheet.id === node.sheetId)?.cells[
          node.address
        ],
      ) &&
        typeof b == "string" &&
        /^#(?:[A-Z0-9_/.]+[!?]?)/.test(b) &&
        failWithMessage(
          "A selected formula or dependency has a calculation error. Correct it before screening.",
        ),
      node.sheetId === materialitySourceRange.id &&
        o.includes(node.address) &&
        a.set(node.address, b));
  }
  return a;
}
const compareAmountToThreshold = (e, r, t) =>
  t === "gt"
    ? e > r
    : t === "gte"
      ? e >= r
      : t === "lt"
        ? e < r
        : t === "lte"
          ? e <= r
          : e === r;
function normalizeMaterialityRequest(workbook, r, t = {}) {
  const materialityConfig = normalizeMaterialityConfig(workbook, r);
  return {
    ...buildMaterialitySourceEvidence(
      workbook,
      {
        sheetId: materialityConfig.sheetId,
        range: materialityConfig.range,
        headers: materialityConfig.headers,
        columns: [
          materialityConfig.amountColumn,
          materialityConfig.currencyColumn,
          materialityConfig.dateColumn,
        ],
      },
      t,
    ),
    request: materialityConfig,
  };
}
function buildMaterialitySourceEvidence(workbook, r, t = {}) {
  (!r ||
    typeof r.sheetId != "string" ||
    typeof r.range != "string" ||
    typeof r.headers != "boolean" ||
    !Array.isArray(r.columns) ||
    r.columns.length > 500) &&
    failWithMessage(
      "Provide a bounded source range and scalar field selection.",
    );
  const rangeAddress = parseRangeAddress(r.range),
    o = workbook.sheets.find((sheet) => sheet.id === r.sheetId);
  (!o ||
    rangeAddress.lastRow >= o.rowCount ||
    rangeAddress.lastCol >= o.colCount ||
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      materialityScreenLimits.sourceCells ||
    r.columns.some(
      (column) =>
        !Number.isInteger(column) ||
        column < rangeAddress.firstCol ||
        column > rangeAddress.lastCol,
    )) &&
    failWithMessage(
      "Choose scalar fields inside one source rectangle of at most 500 cells and within worksheet dimensions.",
    );
  const a = {
      ...r,
      range: `${formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol)}`,
    },
    screenHostScope = normalizeScreenHostScope(workbook, t),
    materialitySourceRange = resolveMaterialitySourceRange(
      workbook,
      a,
      screenHostScope,
    ),
    u = parseRangeAddress(a.range);
  for (let firstRow = u.firstRow; firstRow <= u.lastRow; firstRow++)
    for (let firstCol = u.firstCol; firstCol <= u.lastCol; firstCol++) {
      const g =
        materialitySourceRange.cells[formatCellAddress(firstRow, firstCol)]
          ?.raw;
      typeof g == "number" &&
        !Number.isFinite(g) &&
        failWithMessage(
          "Selected raw evidence contains a nonfinite number and cannot be preserved in JSON. Correct it before screening.",
        );
    }
  const l = [];
  for (let m = u.firstRow + Number(a.headers); m <= u.lastRow; m++)
    for (const v of new Set(a.columns)) l.push(formatCellAddress(m, v));
  const c = inspectScreenFormulaCells(
      workbook,
      materialitySourceRange,
      l,
      screenHostScope,
    ),
    rows = [];
  for (let m = u.firstRow + Number(a.headers); m <= u.lastRow; m++)
    rows.push({
      sourceRow: m + 1,
      sourceRange: `${formatCellAddress(m, u.firstCol)}:${formatCellAddress(m, u.lastCol)}`,
      cells: Array.from({ length: u.lastCol - u.firstCol + 1 }, (v, g) =>
        structuredClone(
          materialitySourceRange.cells[
            formatCellAddress(m, u.firstCol + g)
          ] ?? { raw: null },
        ),
      ),
      values: Array.from({ length: u.lastCol - u.firstCol + 1 }, (v, g) => {
        const cellAddress = formatCellAddress(m, u.firstCol + g);
        return c.has(cellAddress)
          ? c.get(cellAddress)
          : (materialitySourceRange.cells[cellAddress]?.raw ?? null);
      }),
    });
  const f = {
    context: screenHostScope,
    sheetId: materialitySourceRange.id,
    sheetName: materialitySourceRange.name,
    rows: rows,
    headers: Array.from({ length: u.lastCol - u.firstCol + 1 }, (m, v) =>
      a.headers
        ? String(
            materialitySourceRange.cells[
              formatCellAddress(u.firstRow, u.firstCol + v)
            ]?.raw ?? `Column ${columnIndexToLetters(u.firstCol + v)}`,
          )
        : `Column ${columnIndexToLetters(u.firstCol + v)}`,
    ),
  };
  return (
    new TextEncoder().encode(JSON.stringify(f, null, 2)).length > 2e6 &&
      failWithMessage(
        "The complete source evidence exceeds 2,000,000 bytes. Select a smaller population.",
      ),
    f
  );
}
function runMaterialityScreen(workbook, r, t = {}) {
  const materialityRequest = normalizeMaterialityRequest(workbook, r, t),
    { request, context } = materialityRequest,
    rangeAddress = parseRangeAddress(request.range),
    rows = [],
    counts = {
      population: 0,
      eligible: 0,
      screened: 0,
      flagged: 0,
      notFlagged: 0,
      excluded: 0,
      notScreened: 0,
      currencyExceptions: 0,
      invalidAmounts: 0,
      invalidDates: 0,
      outsidePeriod: 0,
    };
  let knownSourceAmount = 0,
    eligibleAmount = 0,
    unknownOrOtherCurrencyRows = 0;
  for (const row of materialityRequest.rows) {
    const v = (I) => row.values[I - rangeAddress.firstCol],
      g = v(request.amountColumn),
      b = v(request.currencyColumn),
      y = v(request.dateColumn),
      amount = typeof g == "number" && Number.isFinite(g) ? g : null,
      currency =
        typeof b == "string" && /^[A-Za-z]{3}$/.test(b)
          ? b.toUpperCase()
          : null,
      dateToIso = normalizeDateToIso(y),
      reasons = [];
    (amount === null &&
      (reasons.push(
        "Amount is missing or is not a finite numeric value (numeric text is not coerced).",
      ),
      counts.invalidAmounts++),
      (!currency || currency !== request.currency) &&
        (reasons.push(
          currency
            ? `Currency ${currency} differs from the declared ${request.currency}; conversion is unsupported.`
            : "Currency is missing or is not a three-letter identifier.",
        ),
        counts.currencyExceptions++),
      dateToIso
        ? (dateToIso < request.effectiveFrom ||
            dateToIso > request.effectiveTo) &&
          (reasons.push("Date is outside the inclusive effective period."),
          counts.outsidePeriod++)
        : (reasons.push(
            "Date is missing or invalid; use ISO YYYY-MM-DD or an integer Excel 1900-system date serial (except 60).",
          ),
          counts.invalidDates++));
    const comparedAmount =
      amount === null
        ? null
        : request.amountMode === "absolute"
          ? Math.abs(amount)
          : amount;
    (amount !== null && currency === request.currency
      ? (knownSourceAmount += amount)
      : unknownOrOtherCurrencyRows++,
      reasons.length || (counts.eligible++, (eligibleAmount += amount)),
      rows.push({
        ...row,
        amount: amount,
        currency: currency,
        date: dateToIso,
        comparedAmount: comparedAmount,
        disposition: reasons.length
          ? "excluded"
          : compareAmountToThreshold(
                comparedAmount,
                request.threshold,
                request.comparator,
              )
            ? "flagged"
            : "not-flagged",
        reasons: reasons,
      }));
  }
  if (counts.currencyExceptions)
    for (const row of rows)
      row.disposition !== "excluded" &&
        ((row.disposition = "not-screened"),
        row.reasons.push(
          "Screen blocked by currency exceptions; no pass or exception conclusion is available.",
        ));
  counts.population = rows.length;
  for (const row of rows)
    row.disposition === "flagged"
      ? counts.flagged++
      : row.disposition === "not-flagged"
        ? counts.notFlagged++
        : row.disposition === "excluded"
          ? counts.excluded++
          : counts.notScreened++;
  counts.screened = counts.flagged + counts.notFlagged;
  const flaggedAmount = counts.currencyExceptions
    ? null
    : rows
        .filter((row) => row.disposition === "flagged")
        .reduce((m, v) => m + v.amount, 0);
  [
    knownSourceAmount,
    eligibleAmount,
    knownSourceAmount - eligibleAmount,
    flaggedAmount,
  ].some((m) => m !== null && !Number.isFinite(m)) &&
    failWithMessage(
      "The control total exceeds finite numeric capacity. Reduce the population or magnitude.",
    );
  const p = {
    schemaVersion: 1,
    kind: "materiality-screen",
    algorithm: "ledgerglass-materiality-1",
    workbookId: workbook.id,
    revision: workbook.revision,
    sheetId: materialityRequest.sheetId,
    sheetName: materialityRequest.sheetName,
    generatedAt: new Date().toISOString(),
    request: structuredClone(request),
    context,
    status: counts.currencyExceptions
      ? "blocked"
      : counts.excluded
        ? "complete-with-exclusions"
        : "complete",
    headers: materialityRequest.headers,
    rows: rows,
    counts: counts,
    totals: {
      currency: request.currency,
      knownSourceAmount: knownSourceAmount,
      eligibleAmount: eligibleAmount,
      excludedKnownAmount: knownSourceAmount - eligibleAmount,
      flaggedAmount: flaggedAmount,
      unknownOrOtherCurrencyRows: unknownOrOtherCurrencyRows,
    },
    notes: [
      "User-configured quantitative screen; it does not determine appropriate audit materiality or identify fraud.",
      "Period endpoints are inclusive. Comparisons use exact stored/computed binary numbers with no display rounding, epsilon or currency conversion.",
      "Every data row is accounted for. Headers alone are excluded automatically; total and subtotal rows are ordinary records unless the selected population excludes them.",
      "Amount/currency/date formulas use a bounded supported static scalar closure. Other evidence columns retain stored values/formula text. Blank fields are reported as missing, never inferred from spill results.",
      "Control totals retain signed amounts and cover only valid numeric amounts in the declared currency; exclusions and unknown or other-currency rows are disclosed.",
      "This report is a local snapshot with source locators; it is not immutable evidence, a signature or a source-lineage guarantee.",
    ],
  };
  return (
    new TextEncoder().encode(JSON.stringify(p, null, 2)).length > 2e6 &&
      failWithMessage(
        "The complete evidence report exceeds 2,000,000 bytes. Select a smaller source population.",
      ),
    Tz.set(p, {
      snapshot: JSON.stringify(workbook),
      request: JSON.stringify(request),
      context: JSON.stringify(context),
      content: JSON.stringify(p),
    }),
    p
  );
}
function checkMaterialityReportFreshness(workbook, r, t, n = {}) {
  const o = Tz.get(r);
  if (!o)
    return {
      current: false,
      disclose: false,
      reason: "This report is untrusted or changed. Run the screen again.",
    };
  try {
    if (o.content !== JSON.stringify(r))
      return {
        current: false,
        disclose: false,
        reason: "This report is untrusted or changed. Run the screen again.",
      };
    const screenHostScope = normalizeScreenHostScope(workbook, n);
    if (
      (resolveMaterialitySourceRange(workbook, r.request, screenHostScope),
      JSON.stringify(screenHostScope) !== o.context)
    )
      return {
        current: false,
        disclose: false,
        reason:
          "The authorized or hidden-content scope changed. Run the screen again.",
      };
    if (JSON.stringify(workbook) !== o.snapshot)
      return {
        current: false,
        disclose: false,
        reason: "The source workbook changed. Run the screen again.",
      };
    const materialityConfig = normalizeMaterialityConfig(workbook, t);
    return JSON.stringify(materialityConfig) !== o.request
      ? {
          current: false,
          disclose: true,
          reason:
            "The screen parameters changed. This is the previous result; run the screen again.",
        }
      : {
          current: true,
          disclose: true,
          reason: "Current result for this source and configuration.",
        };
  } catch {
    return {
      current: false,
      disclose: false,
      reason:
        "The source, settings or authorized scope is no longer valid. Run the screen again.",
    };
  }
}
function exportMaterialityReportJson(workbook, r, t, n = {}) {
  const o = checkMaterialityReportFreshness(workbook, r, t, n);
  return (o.current || failWithMessage(o.reason), JSON.stringify(r, null, 2));
}
function parseMaterialityReportText(workbook, r, t = {}) {
  (typeof r != "string" ||
    r.length > 2e6 ||
    new TextEncoder().encode(r).length > 2e6) &&
    failWithMessage(
      "The screen configuration/report must be at most 2,000,000 bytes.",
    );
  let n;
  try {
    n = JSON.parse(r);
  } catch {
    failWithMessage("The screen configuration/report is not valid JSON.");
  }
  return loadMaterialityConfiguration(workbook, n, t);
}
function loadMaterialityConfiguration(workbook, r, t = {}) {
  (!r || typeof r != "object" || Array.isArray(r)) &&
    failWithMessage("Choose a supported screen configuration or report.");
  const n = r;
  (n.schemaVersion !== 1 || n.kind !== "materiality-screen") &&
    failWithMessage(
      "The materiality screen configuration version is unsupported.",
    );
  const materialityConfig = normalizeMaterialityConfig(workbook, n.request);
  return (
    resolveMaterialitySourceRange(
      workbook,
      materialityConfig,
      normalizeScreenHostScope(workbook, t),
    ),
    materialityConfig
  );
}
const agingScheduleLimits = {
    sourceCells: 500,
    buckets: 24,
    pageSize: 20,
    reportBytes: 2e6,
  },
  _z = new WeakMap(),
  dayInMilliseconds = 864e5;
function parseAgingIsoDate(e) {
  if (
    typeof e != "string" ||
    !/^\d{4}-\d{2}-\d{2}$/.test(e) ||
    e.startsWith("0000-")
  )
    return null;
  const r = Date.parse(`${e}T00:00:00.000Z`);
  return Number.isFinite(r) && new Date(r).toISOString().slice(0, 10) === e
    ? r / dayInMilliseconds
    : null;
}
function normalizeAgingDateInput(e, dateSystem) {
  if (typeof e == "string") return parseAgingIsoDate(e) === null ? null : e;
  if (
    typeof e != "number" ||
    !Number.isInteger(e) ||
    (dateSystem === "1900" && (e < 1 || e > 2958465 || e === 60)) ||
    (dateSystem === "1904" && (e < 0 || e > 2957003))
  )
    return null;
  const t =
    dateSystem === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  return new Date(
    t + (dateSystem === "1900" && e > 60 ? e - 1 : e) * dayInMilliseconds,
  )
    .toISOString()
    .slice(0, 10);
}
function defaultMaterialityConfig(agingConfiguration) {
  return {
    sheetId: agingConfiguration.sheetId,
    range: agingConfiguration.range,
    headers: agingConfiguration.headers,
    amountColumn: agingConfiguration.amountColumn,
    currencyColumn: agingConfiguration.currencyColumn,
    dateColumn: agingConfiguration.dateColumn,
    currency: agingConfiguration.currency,
    threshold: 0,
    amountMode: "absolute",
    comparator: "gte",
    effectiveFrom: "0000-01-01",
    effectiveTo: "9999-12-31",
  };
}
function normalizeAgingConfiguration(workbook, agingConfiguration, t) {
  const n = [
    "sheetId",
    "range",
    "headers",
    "amountColumn",
    "currencyColumn",
    "dateColumn",
    "currency",
    "dateBasis",
    "dateSystem",
    "asOf",
    "buckets",
  ];
  ((!agingConfiguration ||
    typeof agingConfiguration != "object" ||
    Array.isArray(agingConfiguration) ||
    Object.keys(agingConfiguration).length !== n.length ||
    n.some((d) => !Object.hasOwn(agingConfiguration, d))) &&
    throwErrorMessage("Provide a complete, supported aging configuration."),
    (!["due", "document"].includes(agingConfiguration.dateBasis) ||
      !["1900", "1904"].includes(agingConfiguration.dateSystem)) &&
      throwErrorMessage(
        "Choose due or document date and an explicit 1900 or 1904 serial date system.",
      ),
    parseAgingIsoDate(agingConfiguration.asOf) === null &&
      throwErrorMessage(
        "Choose a valid fixed as-of date in ISO YYYY-MM-DD format.",
      ),
    (!Array.isArray(agingConfiguration.buckets) ||
      agingConfiguration.buckets.length < 1 ||
      agingConfiguration.buckets.length > agingScheduleLimits.buckets) &&
      throwErrorMessage("Configure between 1 and 24 aging buckets."));
  let o = 0;
  const a = new Set();
  for (const [d, u] of agingConfiguration.buckets.entries())
    ((!u ||
      typeof u != "object" ||
      Array.isArray(u) ||
      Object.keys(u).length !== 3 ||
      !["label", "from", "to"].every((l) => Object.hasOwn(u, l)) ||
      typeof u.label != "string" ||
      !u.label.trim() ||
      u.label.trim().length > 80 ||
      /[\u0000-\u001f\u007f]/.test(u.label) ||
      a.has(u.label.trim().toLowerCase())) &&
      throwErrorMessage(
        "Each bucket needs a unique label of 1–80 characters and inclusive from/to bounds.",
      ),
      a.add(u.label.trim().toLowerCase()),
      (!Number.isSafeInteger(u.from) ||
        u.from !== o ||
        (d === agingConfiguration.buckets.length - 1
          ? u.to !== null
          : !Number.isSafeInteger(u.to) || u.to < u.from || u.to > 3652424)) &&
        throwErrorMessage(
          "Buckets must cover every nonnegative day exactly once: start at 0, use contiguous inclusive integer bounds, and leave only the final upper bound open.",
        ),
      (o = u.to === null ? 0 : u.to + 1));
  const materialityReportText = parseMaterialityReportText(
    workbook,
    JSON.stringify({
      schemaVersion: 1,
      kind: "materiality-screen",
      request: defaultMaterialityConfig(agingConfiguration),
    }),
    t,
  );
  return {
    sheetId: materialityReportText.sheetId,
    range: materialityReportText.range,
    headers: materialityReportText.headers,
    amountColumn: materialityReportText.amountColumn,
    currencyColumn: materialityReportText.currencyColumn,
    dateColumn: materialityReportText.dateColumn,
    currency: materialityReportText.currency,
    dateBasis: agingConfiguration.dateBasis,
    dateSystem: agingConfiguration.dateSystem,
    asOf: agingConfiguration.asOf,
    buckets: agingConfiguration.buckets.map((bucket) => ({
      ...bucket,
      label: bucket.label.trim(),
    })),
  };
}
function addDecimalAmounts(e) {
  const r = e.map((s) => {
      const [d, u = "0"] = String(s).toLowerCase().split("e"),
        [l, c = ""] = d.split(".");
      return { n: BigInt(l + c), scale: c.length - Number(u) };
    }),
    t = Math.max(0, ...r.map((s) => s.scale));
  let n = r.reduce((s, d) => s + d.n * 10n ** BigInt(t - d.scale), 0n);
  if (n === 0n) return "0";
  const o = n < 0n ? "-" : "";
  n < 0n && (n = -n);
  const a = n.toString().padStart(t + 1, "0");
  return o + (t ? `${a.slice(0, -t)}.${a.slice(-t)}`.replace(/\.?0+$/, "") : a);
}
function assertAgingControlTotal(e) {
  const r = Number(e);
  return Number.isFinite(r)
    ? r
    : throwErrorMessage(
        "The aging control total exceeds finite numeric capacity. Reduce the selected population or magnitude.",
      );
}
function buildAgingSchedule(workbook, r, t = {}) {
  const agingConfiguration = normalizeAgingConfiguration(workbook, r, t),
    materialityRequest = normalizeMaterialityRequest(
      workbook,
      defaultMaterialityConfig(agingConfiguration),
      t,
    ),
    firstCol = parseRangeAddress(agingConfiguration.range).firstCol,
    agingIsoDate = parseAgingIsoDate(agingConfiguration.asOf),
    rows = materialityRequest.rows.map((row) => {
      const w = row.values[agingConfiguration.amountColumn - firstCol],
        E = row.values[agingConfiguration.currencyColumn - firstCol],
        amount = typeof w == "number" && Number.isFinite(w) ? w : null,
        currency =
          typeof E == "string" && /^[A-Za-z]{3}$/.test(E)
            ? E.toUpperCase()
            : null,
        F = row.values[agingConfiguration.dateColumn - firstCol],
        agingDateInput = normalizeAgingDateInput(
          F,
          agingConfiguration.dateSystem,
        ),
        ageDays =
          agingDateInput === null
            ? null
            : agingIsoDate - parseAgingIsoDate(agingDateInput),
        exceptions = [],
        reasons = [];
      (F === null || F === ""
        ? (exceptions.push("missing-date"),
          reasons.push("Date is missing; no date is inferred."))
        : agingDateInput === null
          ? (exceptions.push("invalid-date"),
            reasons.push(
              `Date must be strict ISO text or an integer ${agingConfiguration.dateSystem}-system calendar serial; time fractions and fictitious 1900-02-29 are unsupported.`,
            ))
          : ageDays < 0 &&
            (exceptions.push("future-date"),
            reasons.push(
              "Date is after the fixed as-of date; future items are excluded and separately counted.",
            )),
        amount === null &&
          (exceptions.push("invalid-amount"),
          reasons.push(
            "Amount is missing or is not a finite number; numeric text is not coerced.",
          )),
        currency !== agingConfiguration.currency &&
          (exceptions.push("unsupported-currency"),
          reasons.push(
            `Currency is missing, invalid or differs from declared ${agingConfiguration.currency}; no conversion or cross-currency aggregation is performed.`,
          )));
      const V = exceptions.length
        ? null
        : agingConfiguration.buckets.findIndex(
            (bucket) =>
              ageDays >= bucket.from &&
              (bucket.to === null || ageDays <= bucket.to),
          );
      return (
        V === -1 &&
          throwErrorMessage(
            "The configured aging buckets do not cover a valid source row.",
          ),
        {
          sourceRow: row.sourceRow,
          sourceRange: row.sourceRange,
          values: structuredClone(row.values),
          cells: structuredClone(row.cells),
          amount: amount,
          currency: currency,
          date: agingDateInput,
          ageDays: ageDays,
          bucket: V,
          disposition: exceptions.length ? "excluded" : "aged",
          exceptions: exceptions,
          reasons: reasons,
        }
      );
    }),
    u = rows.filter(
      (row) =>
        row.amount !== null && row.currency === agingConfiguration.currency,
    ),
    l = rows.filter((row) => row.disposition === "aged"),
    c = u.filter((y) => y.disposition === "excluded"),
    h = (y) => addDecimalAmounts(y.map((w) => w.amount)),
    exactKnownSourceAmount = h(u),
    exactAgedAmount = h(l),
    exactExcludedKnownAmount = h(c),
    buckets = agingConfiguration.buckets.map((bucket, w) => {
      const E = l.filter((k) => k.bucket === w),
        exactAmount = h(E);
      return {
        ...bucket,
        count: E.length,
        amount: assertAgingControlTotal(exactAmount),
        exactAmount: exactAmount,
      };
    }),
    g = (y) => rows.filter((row) => row.exceptions.includes(y)).length,
    b = {
      schemaVersion: 1,
      kind: "aging-schedule",
      algorithm: "ledgerglass-aging-1",
      workbookId: workbook.id,
      revision: workbook.revision,
      sheetId: materialityRequest.sheetId,
      sheetName: materialityRequest.sheetName,
      generatedAt: new Date().toISOString(),
      request: structuredClone(agingConfiguration),
      context: structuredClone(materialityRequest.context),
      status:
        l.length === rows.length ? "complete" : "complete-with-exclusions",
      headers: structuredClone(materialityRequest.headers),
      rows: rows,
      buckets: buckets,
      counts: {
        population: rows.length,
        aged: l.length,
        excluded: rows.length - l.length,
        missingDates: g("missing-date"),
        invalidDates: g("invalid-date"),
        futureDates: g("future-date"),
        unsupportedCurrencies: g("unsupported-currency"),
        missingCurrencies: rows.filter((row) => {
          const w = row.values[agingConfiguration.currencyColumn - firstCol];
          return w === null || w === "";
        }).length,
        invalidAmounts: g("invalid-amount"),
      },
      totals: {
        currency: agingConfiguration.currency,
        knownSourceAmount: assertAgingControlTotal(exactKnownSourceAmount),
        agedAmount: assertAgingControlTotal(exactAgedAmount),
        excludedKnownAmount: assertAgingControlTotal(exactExcludedKnownAmount),
        unknownOrOtherCurrencyRows: rows.length - u.length,
        exactKnownSourceAmount: exactKnownSourceAmount,
        exactAgedAmount: exactAgedAmount,
        exactExcludedKnownAmount: exactExcludedKnownAmount,
        reconciled: true,
      },
      notes: [
        "Inclusive bucket endpoints cover every nonnegative day exactly once; the final bucket is open-ended. Age is the UTC calendar-day difference from the chosen due/document date to the fixed as-of date.",
        "Every source data row is either aged once or excluded once. Exception categories may overlap. Future, missing and invalid dates are distinct categories. Negative and zero balances retain their sign.",
        "Only the explicitly declared currency is summed. Missing, invalid and other currencies are separately counted exceptions; no foreign-exchange conversion is performed.",
        "Exact decimal controls sum each stored/computed number's shortest decimal representation. Bucket amounts equal the aged population; aged plus known-currency exclusions equals the known-currency source. Unknown amounts and other currencies are outside these totals.",
        "Application dates and imported calendar dates use the 1900 system. Select 1904 only for raw, unconverted 1904 serials; choosing a different serial system never changes ISO dates. ISO years 0001–9999 are accepted; year 0000 is an invalid calendar date. Fractional serials are exceptions, not silently truncated.",
        "Selected amount/currency/date formulas use the existing authorized static scalar dependency closure (1,000 cells, depth 8, 5,000 scanned formulas); volatile, array, dynamic, hidden or unauthorized dependencies are rejected. Other evidence formulas remain stored text.",
        "Headers alone are omitted automatically. Total/subtotal lines are ordinary data unless excluded from the selected population. Source cells and history remain unchanged. This local snapshot is not signed or immutable evidence.",
      ],
    };
  return (
    new TextEncoder().encode(JSON.stringify(b, null, 2)).length >
      agingScheduleLimits.reportBytes &&
      throwErrorMessage(
        "The aging report exceeds 2,000,000 bytes. Select a smaller population.",
      ),
    _z.set(b, {
      snapshot: JSON.stringify(workbook),
      context: JSON.stringify(materialityRequest.context),
      request: JSON.stringify(agingConfiguration),
      content: JSON.stringify(b),
    }),
    b
  );
}
function checkAgingReportFreshness(workbook, r, t, n = {}) {
  const o = _z.get(r);
  if (!o || o.content !== JSON.stringify(r))
    return {
      current: false,
      disclose: false,
      reason:
        "This aging report is untrusted or changed. Run the schedule again.",
    };
  try {
    const agingConfiguration = normalizeAgingConfiguration(workbook, t, n),
      s = {
        includeHidden: !!n.includeHidden,
        allowedSheetIds: [
          ...new Set(
            n.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
          ),
        ].sort(),
      };
    return o.snapshot !== JSON.stringify(workbook) ||
      o.context !== JSON.stringify(s)
      ? {
          current: false,
          disclose: false,
          reason:
            "The source or authorized scope changed. Run the aging schedule again.",
        }
      : JSON.stringify(agingConfiguration) !== o.request
        ? {
            current: false,
            disclose: true,
            reason:
              "The aging settings changed. This is the previous result; run the schedule again.",
          }
        : {
            current: true,
            disclose: true,
            reason: "Current aging result for this source and configuration.",
          };
  } catch {
    return {
      current: false,
      disclose: false,
      reason:
        "The source, settings or authorized scope is no longer valid. Run the aging schedule again.",
    };
  }
}
function exportAgingReportJson(workbook, r, t, n = {}) {
  const o = checkAgingReportFreshness(workbook, r, t, n);
  return (o.current || throwErrorMessage(o.reason), JSON.stringify(r, null, 2));
}
function loadAgingConfiguration(workbook, r, t = {}) {
  return (
    (!r ||
      typeof r != "object" ||
      Array.isArray(r) ||
      r.schemaVersion !== 1 ||
      r.kind !== "aging-schedule") &&
      throwErrorMessage("Choose a supported aging configuration or report."),
    normalizeAgingConfiguration(workbook, r.request, t)
  );
}
const listRangeColumnIndexes = (e) => {
    try {
      const rangeAddress = parseRangeAddress(e);
      return rangeAddress.lastCol - rangeAddress.firstCol < 500
        ? Array.from(
            { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
            (t, n) => rangeAddress.firstCol + n,
          )
        : [];
    } catch {
      return [];
    }
  },
  defaultAgingBuckets = [
    { label: "0–30 days", from: 0, to: 30 },
    { label: "31–60 days", from: 31, to: 60 },
    { label: "61–90 days", from: 61, to: 90 },
    { label: "91+ days", from: 91, to: null },
  ];
function AgingScheduleScreen(e) {
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    o = ne.useId(),
    a = ne.useRef(null),
    s = ne.useRef(true),
    d = ne.useRef(0),
    u = ne.useRef(""),
    rangeColumnIndexes = listRangeColumnIndexes(e.range),
    [range, h] = ne.useState(e.range),
    [f, p] = ne.useState(true),
    [amountColumn, v] = ne.useState(rangeColumnIndexes[0] ?? 0),
    [currencyColumn, b] = ne.useState(rangeColumnIndexes[1] ?? 1),
    [dateColumn, w] = ne.useState(rangeColumnIndexes[2] ?? 2),
    [currency, A] = ne.useState("USD"),
    [dateBasis, F] = ne.useState("due"),
    [dateSystem, O] = ne.useState(e.workbook.datePolicy?.dateSystem ?? "1900"),
    [asOf, z] = ne.useState(new Date().toISOString().slice(0, 10)),
    [buckets, D] = ne.useState(structuredClone(defaultAgingBuckets)),
    [checked, _] = ne.useState(false),
    [C, T] = ne.useState(null),
    [N, R] = ne.useState(""),
    [L, $] = ne.useState(""),
    [U, G] = ne.useState(""),
    [H, Q] = ne.useState(""),
    [X$1, J] = ne.useState(0),
    [q, W] = ne.useState("all"),
    K = {
      sheetId: e.sheetId,
      range: range,
      headers: f,
      amountColumn: amountColumn,
      currencyColumn: currencyColumn,
      dateColumn: dateColumn,
      currency: currency,
      dateBasis: dateBasis,
      dateSystem: dateSystem,
      asOf: asOf,
      buckets: buckets,
    },
    te = {
      readOnly: e.readOnly,
      includeHidden: !!e.allowHidden && checked,
      allowedSheetIds: e.allowedSheetIds,
    },
    ae = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    oe = [
      ...new Set(
        e.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    de =
      !!ae &&
      oe.includes(ae.id) &&
      (te.includeHidden || (!ae.hidden && !ae.veryHidden)),
    he = JSON.stringify([e.sheetId, e.range]),
    Se = JSON.stringify([e.workbook, oe, te.includeHidden, he]),
    le = JSON.stringify([Se, K]);
  u.current = le;
  const ce = C
      ? N !== he
        ? {
            current: false,
            disclose: false,
            reason:
              "The source selection changed. Run the aging schedule again.",
          }
        : checkAgingReportFreshness(e.workbook, C, K, te)
      : null,
    ge = ce?.disclose ? C : null,
    se =
      ge?.rows.filter(
        (row) =>
          q === "all" || q === row.disposition || row.exceptions.includes(q),
      ) ?? [],
    we = Math.max(1, Math.ceil(se.length / agingScheduleLimits.pageSize)),
    Ee = Math.min(X$1, we - 1),
    Ae = listRangeColumnIndexes(range),
    ye = (Me) => {
      ($(
        Me instanceof Error
          ? Me.message
          : "The aging schedule could not complete. Review the settings and source.",
      ),
        G(Se),
        Q(""));
    };
  ne.useEffect(() => {
    s.current = true;
    const activeElement = document.activeElement;
    return (
      a.current?.querySelector('[data-ui-label="Aging source range"]')?.focus(),
      () => {
        ((s.current = false),
          d.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []);
  const onClick = () => {
      d.current++;
      try {
        (T(buildAgingSchedule(e.workbook, K, te)),
          R(he),
          $(""),
          Q(""),
          W("all"),
          J(0));
      } catch (error) {
        (T(null), ye(error));
      }
    },
    Pe = () => {
      try {
        if (!C || !ce?.current)
          throw Error("Run a current aging schedule before downloading.");
        const Me = exportAgingReportJson(e.workbook, C, K, te),
          ke = URL.createObjectURL(
            new Blob([Me], { type: "application/json" }),
          ),
          je = document.createElement("a");
        ((je.href = ke),
          (je.download = "aging-schedule.json"),
          je.click(),
          URL.revokeObjectURL(ke),
          Q(
            "Aging report downloaded with settings, all source rows, exceptions and exact control totals.",
          ),
          $(""));
      } catch (error) {
        ye(error);
      }
    },
    qe = async (Me) => {
      if (!Me) return;
      const ke = le,
        je = ++d.current,
        $e = () => s.current && je === d.current && u.current === ke;
      try {
        if (Me.size > agingScheduleLimits.reportBytes)
          throw Error(
            "Choose an aging configuration/report no larger than 2,000,000 bytes.",
          );
        const Je = await importNativeJsonFile(Me, cancellableOperation.begin());
        if (!$e()) return;
        const agingConfiguration = loadAgingConfiguration(e.workbook, Je, te);
        if (agingConfiguration.sheetId !== e.sheetId)
          throw Error(
            "Open the saved configuration's source worksheet before loading it.",
          );
        (h(agingConfiguration.range),
          p(agingConfiguration.headers),
          v(agingConfiguration.amountColumn),
          b(agingConfiguration.currencyColumn),
          w(agingConfiguration.dateColumn),
          A(agingConfiguration.currency),
          F(agingConfiguration.dateBasis),
          O(agingConfiguration.dateSystem),
          z(agingConfiguration.asOf),
          D(agingConfiguration.buckets),
          T(null),
          $(""),
          Q(
            "Aging configuration loaded. Run it against the current source; saved conclusions are not reused.",
          ));
      } catch (error) {
        $e() && ye(toFileProcessingError(error));
      }
    },
    ht = (Me, ke) =>
      D((je) => je.map(($e, Je) => (Je === Me ? { ...$e, ...ke } : $e)));
  return i.jsx("div", {
    className: "aging-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      className: "aging-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${o}-title`,
      ref: a,
      onKeyDown: (event) => {
        var je, $e;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), e.onClose()),
          event.key === "Tab")
        ) {
          const Je = Array.from(
            a.current?.querySelectorAll(
              'button:not(:disabled), input:not(:disabled), select:not(:disabled), summary, [tabindex="0"]',
            ) ?? [],
          ).filter(
            (ze) =>
              !ze.closest("details:not([open])") || ze.tagName === "SUMMARY",
          );
          event.shiftKey && document.activeElement === Je[0]
            ? (event.preventDefault(), (je = Je.at(-1)) == null || je.focus())
            : !event.shiftKey &&
              document.activeElement === Je.at(-1) &&
              (event.preventDefault(), ($e = Je[0]) == null || $e.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${o}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Aging schedule",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Age receivables or payables from a due or document date using your fixed as-of date and inclusive buckets.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close aging schedule"),
              onClick: e.onClose,
              lang: language("Close aging schedule"),
              "data-ui-label": "Close aging schedule",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "This schedule reads the source without editing cells or history. Protected and read-only workbooks support inspection.",
          }),
        }),
        i.jsxs("div", {
          className: "aging-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source worksheet" }),
                i.jsx("strong", {
                  children: de ? ae.name : "Outside current context",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Aging source range"),
                  value: de ? range : "",
                  onChange: (event) => h(event.target.value),
                  lang: language("Aging source range"),
                  "data-ui-label": "Aging source range",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "aging-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: f,
                  onChange: (event) => p(event.target.checked),
                }),
                i.jsx(LocalizedText, {
                  children: "First row contains headers",
                }),
              ],
            }),
            [
              ["Amount column", amountColumn, v],
              ["Currency column", currencyColumn, b],
              ["Date column", dateColumn, w],
            ].map(([Me, ke, je]) =>
              i.jsxs(
                "label",
                {
                  children: [
                    Me,
                    i.jsx("select", {
                      "aria-label": Me,
                      value: ke,
                      onChange: (event) => je(Number(event.target.value)),
                      children: Ae.map(($e) =>
                        i.jsx(
                          "option",
                          { value: $e, children: columnIndexToLetters($e) },
                          $e,
                        ),
                      ),
                    }),
                  ],
                },
                Me,
              ),
            ),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Declared currency" }),
                i.jsx("input", {
                  "aria-label": text("Declared currency"),
                  maxLength: 3,
                  value: currency,
                  onChange: (event) => A(event.target.value),
                  lang: language("Declared currency"),
                  "data-ui-label": "Declared currency",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Date basis" }),
                i.jsxs("select", {
                  "aria-label": text("Date basis"),
                  value: dateBasis,
                  onChange: (event) => F(event.target.value),
                  lang: language("Date basis"),
                  "data-ui-label": "Date basis",
                  children: [
                    i.jsx("option", { value: "due", children: "Due date" }),
                    i.jsx("option", {
                      value: "document",
                      children: "Document date",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Fixed as-of date" }),
                i.jsx("input", {
                  "aria-label": text("Fixed as-of date"),
                  type: "date",
                  value: asOf,
                  onChange: (event) => z(event.target.value),
                  lang: language("Fixed as-of date"),
                  "data-ui-label": "Fixed as-of date",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, {
                  children: "Numeric serial date system",
                }),
                i.jsxs("select", {
                  "aria-label": text("Numeric serial date system"),
                  value: dateSystem,
                  onChange: (event) => O(event.target.value),
                  lang: language("Numeric serial date system"),
                  "data-ui-label": "Numeric serial date system",
                  children: [
                    i.jsx("option", {
                      value: "1900",
                      children: "1900 — application and imported dates",
                    }),
                    i.jsx("option", {
                      value: "1904",
                      children: "1904 — raw unconverted serials",
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
        i.jsxs("fieldset", {
          children: [
            i.jsx("legend", {
              children: i.jsx(LocalizedText, {
                children: "Inclusive aging buckets",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Start at day 0. Bounds must be contiguous without overlap; the last bucket is open-ended. Future dates are separate exceptions.",
              }),
            }),
            i.jsx("div", {
              className: "aging-scroll",
              children: i.jsxs("table", {
                children: [
                  i.jsx("thead", {
                    children: i.jsxs("tr", {
                      children: [
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, { children: "Label" }),
                        }),
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, {
                            children: "From day (inclusive)",
                          }),
                        }),
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, {
                            children: "To day (inclusive)",
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
                  i.jsx("tbody", {
                    children: buckets.map((bucket, ke) =>
                      i.jsxs(
                        "tr",
                        {
                          children: [
                            i.jsx("td", {
                              children: i.jsx("input", {
                                "aria-label": `Bucket ${ke + 1} label`,
                                value: bucket.label,
                                maxLength: 80,
                                onChange: (event) =>
                                  ht(ke, { label: event.target.value }),
                              }),
                            }),
                            i.jsx("td", {
                              children: i.jsx(LocalizedNumberField, {
                                "aria-label": `Bucket ${ke + 1} from`,
                                min: "0",
                                step: "1",
                                value: Number.isFinite(bucket.from)
                                  ? bucket.from
                                  : "",
                                onValueChange: (je) =>
                                  ht(ke, {
                                    from: je === "" ? NaN : Number(je),
                                  }),
                              }),
                            }),
                            i.jsx("td", {
                              children: i.jsx(LocalizedNumberField, {
                                allowEmpty: true,
                                "aria-label": `Bucket ${ke + 1} to`,
                                min: "0",
                                step: "1",
                                placeholder: text("Open-ended"),
                                value:
                                  bucket.to === null
                                    ? ""
                                    : Number.isFinite(bucket.to)
                                      ? bucket.to
                                      : "",
                                onValueChange: (je) =>
                                  ht(ke, { to: je === "" ? null : Number(je) }),
                                lang: language("Open-ended"),
                              }),
                            }),
                            i.jsx("td", {
                              children: i.jsx("button", {
                                "aria-label": `Remove bucket ${ke + 1}`,
                                disabled: buckets.length === 1,
                                onClick: () =>
                                  D((je) => je.filter(($e, Je) => Je !== ke)),
                                children: i.jsx(LocalizedText, {
                                  children: "Remove",
                                }),
                              }),
                            }),
                          ],
                        },
                        ke,
                      ),
                    ),
                  }),
                ],
              }),
            }),
            i.jsx("button", {
              disabled: buckets.length >= agingScheduleLimits.buckets,
              onClick: () =>
                D((Me) => {
                  const ke = Me.at(-1),
                    je = ke.to ?? ke.from + 29;
                  return [
                    ...Me.slice(0, -1),
                    { ...ke, to: je },
                    { label: `${je + 1}+ days`, from: je + 1, to: null },
                  ];
                }),
              children: i.jsx(LocalizedText, { children: "Add bucket" }),
            }),
            i.jsx("button", {
              onClick: () => D(structuredClone(defaultAgingBuckets)),
              children: i.jsx(LocalizedText, { children: "Reset buckets" }),
            }),
          ],
        }),
        e.allowHidden &&
          i.jsxs("label", {
            className: "aging-check",
            children: [
              i.jsx("input", {
                "aria-label": text("Include hidden content"),
                type: "checkbox",
                checked: checked,
                onChange: (event) => _(event.target.checked),
                lang: language("Include hidden content"),
                "data-ui-label": "Include hidden content",
              }),
              i.jsx(LocalizedText, {
                children:
                  "Explicitly include hidden source and dependency cells within permitted host context",
              }),
            ],
          }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Up to 500 selected cells. Dates use strict ISO calendar text or whole-day serials in the selected system. Invalid, missing, future and other-currency rows are counted separately. The declared currency alone is totaled; no conversion is performed.",
          }),
        }),
        i.jsxs("div", {
          className: "aging-actions",
          children: [
            i.jsx("button", {
              className: "aging-primary",
              disabled: !de,
              onClick: onClick,
              children: i.jsx(LocalizedText, {
                children: "Run aging schedule",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, {
                  children: "Load saved configuration/report",
                }),
                i.jsx("input", {
                  "aria-label": text("Load aging configuration"),
                  type: "file",
                  accept: "application/json,.json",
                  onChange: (event) => {
                    const ke = event.target.files?.[0];
                    ((event.target.value = ""), qe(ke));
                  },
                  lang: language("Load aging configuration"),
                  "data-ui-label": "Load aging configuration",
                }),
              ],
            }),
            i.jsx("button", {
              disabled: !ce?.current,
              onClick: Pe,
              children: i.jsx(LocalizedText, {
                children: "Download current aging report",
              }),
            }),
          ],
        }),
        L &&
          U === Se &&
          i.jsx("p", { role: "alert", className: "aging-error", children: L }),
        H && i.jsx("p", { role: "status", children: H }),
        C &&
          ce &&
          i.jsx("p", {
            role: "status",
            className: ce.current ? "" : "aging-stale",
            children: ce.current ? ce.reason : `Stale result. ${ce.reason}`,
          }),
        ge &&
          i.jsxs("section", {
            className: "aging-result",
            "aria-label": text("Aging schedule results"),
            lang: language("Aging schedule results"),
            "data-ui-label": "Aging schedule results",
            children: [
              i.jsx("h3", {
                children:
                  ge.status === "complete"
                    ? "Schedule complete"
                    : "Complete with exclusions — review every excluded row",
              }),
              i.jsxs("p", {
                children: [
                  ge.request.dateBasis === "due" ? "Due date" : "Document date",
                  " ",
                  i.jsx(LocalizedText, { children: "basis · as of " }),
                  ge.request.asOf,
                  " · ",
                  ge.request.dateSystem,
                  " ",
                  i.jsx(LocalizedText, { children: "serial system ·" }),
                  " ",
                  ge.request.currency,
                ],
              }),
              i.jsxs("p", {
                children: [
                  ge.counts.population,
                  " ",
                  i.jsx(LocalizedText, { children: "source rows ·" }),
                  " ",
                  ge.counts.aged,
                  " ",
                  i.jsx(LocalizedText, { children: "aged · " }),
                  ge.counts.excluded,
                  " ",
                  i.jsx(LocalizedText, { children: "excluded" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  ge.counts.missingDates,
                  " ",
                  i.jsx(LocalizedText, { children: "missing dates ·" }),
                  " ",
                  ge.counts.invalidDates,
                  " ",
                  i.jsx(LocalizedText, { children: "invalid dates ·" }),
                  " ",
                  ge.counts.futureDates,
                  " ",
                  i.jsx(LocalizedText, { children: "future dates ·" }),
                  " ",
                  ge.counts.unsupportedCurrencies,
                  " ",
                  i.jsx(LocalizedText, {
                    children: "unsupported currencies (including ",
                  }),
                  ge.counts.missingCurrencies,
                  " ",
                  i.jsx(LocalizedText, { children: "missing currencies) · " }),
                  ge.counts.invalidAmounts,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "invalid amounts. Exception categories may overlap.",
                  }),
                ],
              }),
              i.jsxs("table", {
                "aria-label": text("Aging bucket totals"),
                lang: language("Aging bucket totals"),
                "data-ui-label": "Aging bucket totals",
                children: [
                  i.jsx("thead", {
                    children: i.jsxs("tr", {
                      children: [
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, {
                            children: "Bucket",
                          }),
                        }),
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, {
                            children: "Inclusive days",
                          }),
                        }),
                        i.jsx("th", {
                          children: i.jsx(LocalizedText, { children: "Rows" }),
                        }),
                        i.jsxs("th", {
                          children: [
                            i.jsx(LocalizedText, { children: "Signed " }),
                            ge.totals.currency,
                            " ",
                            i.jsx(LocalizedText, { children: "total" }),
                          ],
                        }),
                      ],
                    }),
                  }),
                  i.jsx("tbody", {
                    children: ge.buckets.map((bucket, ke) =>
                      i.jsxs(
                        "tr",
                        {
                          children: [
                            i.jsx("th", {
                              scope: "row",
                              children: bucket.label,
                            }),
                            i.jsxs("td", {
                              children: [
                                bucket.from,
                                "–",
                                bucket.to ?? "unbounded",
                              ],
                            }),
                            i.jsx("td", { children: bucket.count }),
                            i.jsx("td", { children: bucket.exactAmount }),
                          ],
                        },
                        ke,
                      ),
                    ),
                  }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Exact signed " }),
                  ge.totals.currency,
                  " ",
                  i.jsx(LocalizedText, {
                    children: "control totals: known source ",
                  }),
                  ge.totals.exactKnownSourceAmount,
                  " ",
                  i.jsx(LocalizedText, { children: "= aged" }),
                  " ",
                  ge.totals.exactAgedAmount,
                  " ",
                  i.jsx(LocalizedText, { children: "+ excluded known amount" }),
                  " ",
                  ge.totals.exactExcludedKnownAmount,
                  ".",
                  " ",
                  ge.totals.unknownOrOtherCurrencyRows,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "rows have unknown amounts or another/unknown currency and are outside these totals.",
                  }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Source: " }),
                  ge.sheetName,
                  "!",
                  ge.request.range,
                  " ",
                  i.jsx(LocalizedText, { children: "· workbook " }),
                  ge.workbookId,
                  " ",
                  i.jsx(LocalizedText, { children: "· revision" }),
                  " ",
                  ge.revision,
                  " · ",
                  ge.algorithm,
                ],
              }),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Rows to review" }),
                  i.jsxs("select", {
                    "aria-label": text("Aging rows to review"),
                    value: q,
                    onChange: (event) => {
                      (W(event.target.value), J(0));
                    },
                    lang: language("Aging rows to review"),
                    "data-ui-label": "Aging rows to review",
                    children: [
                      i.jsx("option", {
                        value: "all",
                        children: "All source rows",
                      }),
                      i.jsx("option", { value: "aged", children: "Aged" }),
                      i.jsx("option", {
                        value: "excluded",
                        children: "Excluded",
                      }),
                      i.jsx("option", {
                        value: "missing-date",
                        children: "Missing date",
                      }),
                      i.jsx("option", {
                        value: "invalid-date",
                        children: "Invalid date",
                      }),
                      i.jsx("option", {
                        value: "future-date",
                        children: "Future date",
                      }),
                      i.jsx("option", {
                        value: "unsupported-currency",
                        children: "Unsupported currency",
                      }),
                      i.jsx("option", {
                        value: "invalid-amount",
                        children: "Invalid amount",
                      }),
                    ],
                  }),
                ],
              }),
              i.jsx("div", {
                className: "aging-scroll",
                children: i.jsxs("table", {
                  "aria-label": text("Aging source rows"),
                  lang: language("Aging source rows"),
                  "data-ui-label": "Aging source rows",
                  children: [
                    i.jsx("thead", {
                      children: i.jsxs("tr", {
                        children: [
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Source row / range",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Date / age",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Amount / currency",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Bucket / disposition",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Exceptions",
                            }),
                          }),
                          ge.headers.map((header, ke) =>
                            i.jsxs(
                              "th",
                              {
                                children: [
                                  header,
                                  " (",
                                  columnIndexToLetters(
                                    parseRangeAddress(ge.request.range)
                                      .firstCol + ke,
                                  ),
                                  ")",
                                ],
                              },
                              ke,
                            ),
                          ),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: se.slice(Ee * 20, Ee * 20 + 20).map((Me) =>
                        i.jsxs(
                          "tr",
                          {
                            children: [
                              i.jsxs("th", {
                                scope: "row",
                                children: [Me.sourceRow, " · ", Me.sourceRange],
                              }),
                              i.jsxs("td", {
                                children: [
                                  Me.date ?? "Invalid or missing",
                                  " /",
                                  " ",
                                  Me.ageDays ?? "Unavailable",
                                ],
                              }),
                              i.jsxs("td", {
                                children: [
                                  Me.amount ?? "Invalid",
                                  " /",
                                  " ",
                                  Me.currency ?? "Invalid",
                                ],
                              }),
                              i.jsx("td", {
                                children:
                                  Me.bucket === null
                                    ? "Excluded"
                                    : ge.buckets[Me.bucket].label,
                              }),
                              i.jsx("td", {
                                children:
                                  Me.reasons.join(" ") ||
                                  "Aged once in the configured bucket.",
                              }),
                              Me.cells.map((cell, je) =>
                                i.jsx(
                                  "td",
                                  {
                                    children: i.jsxs("details", {
                                      children: [
                                        i.jsx("summary", {
                                          children:
                                            JSON.stringify(Me.values[je]) ??
                                            "null",
                                        }),
                                        i.jsx("p", {
                                          children: i.jsx(LocalizedText, {
                                            children: "Stored cell evidence",
                                          }),
                                        }),
                                        i.jsx("pre", {
                                          children: JSON.stringify(
                                            cell,
                                            null,
                                            2,
                                          ),
                                        }),
                                      ],
                                    }),
                                  },
                                  je,
                                ),
                              ),
                            ],
                          },
                          Me.sourceRow,
                        ),
                      ),
                    }),
                  ],
                }),
              }),
              i.jsxs("nav", {
                "aria-label": text("Aging result pages"),
                lang: language("Aging result pages"),
                "data-ui-label": "Aging result pages",
                children: [
                  i.jsx("button", {
                    disabled: Ee === 0,
                    onClick: () => J(Ee - 1),
                    children: i.jsx(LocalizedText, {
                      children: "Previous rows",
                    }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      Ee + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      we,
                      " · ",
                      se.length,
                      " ",
                      i.jsx(LocalizedText, { children: "rows in this view" }),
                    ],
                  }),
                  i.jsx("button", {
                    disabled: Ee + 1 >= we,
                    onClick: () => J(Ee + 1),
                    children: i.jsx(LocalizedText, { children: "Next rows" }),
                  }),
                ],
              }),
              i.jsxs("details", {
                children: [
                  i.jsx("summary", {
                    children: i.jsx(LocalizedText, {
                      children: "Aging rules and limitations",
                    }),
                  }),
                  ge.notes.map((note) => i.jsx("p", { children: note }, note)),
                ],
              }),
            ],
          }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: e.onClose,
            children: i.jsx(LocalizedText, {
              children: "Close aging schedule",
            }),
          }),
        }),
      ],
    }),
  });
}

export {
  AgingScheduleScreen,
  materialityScreenLimits,
  loadMaterialityConfiguration,
  checkMaterialityReportFreshness,
  requestAssistantAnswer,
  buildMaterialitySourceEvidence,
  exportMaterialityReportJson,
  runMaterialityScreen,
};
export { listRangeColumnIndexes };
