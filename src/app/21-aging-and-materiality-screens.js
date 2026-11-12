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

async function requestAssistantAnswer(e, r, t, n = {}) {
  if (n.reliability) {
    const { reliability, ...c } = n;
    return reliability.run("provider", "request", () =>
      requestAssistantAnswer(e, r, t, c),
    );
  }
  if (
    typeof r != "string" ||
    !r.trim() ||
    r.length > 1e4 ||
    !Array.isArray(t?.cells) ||
    t.cells.length > workbookLimits.maxLLMCells
  )
    throw new LLMRequestError("invalid_request");
  validateCalculationClock(t);
  const o = structuredClone(t),
    a = n.conversation ? parseAssistantChatRequest(n.conversation) : undefined;
  if (
    (a?.attachments.length || a?.messages.some((l) => l.attachments?.length)) &&
    !e.chat?.attachments
  )
    throw new Error(
      "This provider does not support file attachments. Remove the files or use a chat-enabled host adapter.",
    );
  if (a?.messages.length && !e.chat?.history)
    throw new Error(
      "This provider does not support conversation history. Start a new chat or use a chat-enabled host adapter.",
    );
  const s = await raceWithAbortSignal(
    (l) =>
      a
        ? e.propose(r, structuredClone(o), l, structuredClone(a))
        : e.propose(r, structuredClone(o), l),
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
function normalizeScreenHostScope(e, r) {
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
          (t) => typeof t != "string" || !e.sheets.some((n) => n.id === t),
        )) &&
      failWithMessage("The materiality screen host context is invalid."),
    {
      includeHidden: !!r.includeHidden,
      allowedSheetIds: [
        ...new Set(r.allowedSheetIds ?? e.sheets.map((t) => t.id)),
      ].sort(),
    }
  );
}
function normalizeMaterialityConfig(e, r) {
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
      !e.sheets.some((a) => a.id === r.sheetId)) &&
      failWithMessage("The source worksheet no longer exists."),
    (typeof r.range != "string" ||
      !/^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/.test(
        r.range,
      )) &&
      failWithMessage("Choose one rectangular A1 source range."));
  const n = parseRangeAddress(r.range);
  ((n.lastRow - n.firstRow + 1) * (n.lastCol - n.firstCol + 1) >
    materialityScreenLimits.sourceCells &&
    failWithMessage(
      "The materiality screen supports at most 500 selected source cells.",
    ),
    (typeof r.headers != "boolean" ||
      n.lastRow - n.firstRow + 1 <= Number(r.headers)) &&
      failWithMessage(
        "Choose a source range with at least one data row and valid header handling.",
      ));
  const o = [r.amountColumn, r.currencyColumn, r.dateColumn];
  return (
    (o.some((a) => !Number.isInteger(a) || a < n.firstCol || a > n.lastCol) ||
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
      range: `${formatCellAddress(n.firstRow, n.firstCol)}:${formatCellAddress(n.lastRow, n.lastCol)}`,
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
function resolveMaterialitySourceRange(e, r, t) {
  const n = e.sheets.find((a) => a.id === r.sheetId);
  (!n || !t.allowedSheetIds.includes(n.id)) &&
    failWithMessage("The source is outside the authorized host context.");
  const o = parseRangeAddress(r.range);
  return (
    !t.includeHidden &&
      (n.hidden ||
        n.veryHidden ||
        cn(n, "row").some((a) => a.start <= o.lastRow && a.end >= o.firstRow) ||
        cn(n, "column").some(
          (a) => a.start <= o.lastCol && a.end >= o.firstCol,
        )) &&
      failWithMessage(
        "The source range includes hidden content. Explicitly include hidden content only within authorized host scope.",
      ),
    (n.merges ?? []).some((a) => {
      const s = parseRangeAddress(a);
      return (
        s.firstRow <= o.lastRow &&
        s.lastRow >= o.firstRow &&
        s.firstCol <= o.lastCol &&
        s.lastCol >= o.firstCol
      );
    }) &&
      failWithMessage(
        "Unmerge the source range before screening so every record has unambiguous fields.",
      ),
    n
  );
}
function inspectScreenFormulaCells(e, r, t, n) {
  const o = t.filter((g) => isFormulaInputValue(r.cells[g])),
    a = new Map(t.map((g) => [g, r.cells[g]?.raw ?? null]));
  if (!o.length) return a;
  const s = JSON.stringify(e);
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
        raw: `=SUM(${o.map((g) => `'${r.name.replace(/'/g, "''")}'!${g}`).join(",")})`,
      },
    },
  };
  String(u.cells.A1.raw).length > 8192 &&
    failWithMessage(
      "The selected formulas exceed the bounded calculation inspection length.",
    );
  const l = inspectFormulaDependencies(
      {
        ...e,
        sheets: [
          u,
          ...e.sheets.map((g) => (g.veryHidden ? { ...g, hidden: true } : g)),
        ],
      },
      {
        sheetId: u.id,
        address: "A1",
        direction: "precedents",
        depth: 8,
        includeHidden: n.includeHidden,
        allowedSheetIds: [...n.allowedSheetIds, u.id],
        limits: { maxNodes: 1e3, maxFormulaCells: 5e3, maxReferences: 1e4 },
      },
    ),
    c = new Set(l.nodes.map((g) => g.key));
  (l.coverage.truncated ||
    l.coverage.depthLimited ||
    l.cycles.length ||
    l.issues.some((g) => g.from === l.target.key || c.has(g.from))) &&
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
  for (const g of l.nodes) {
    const b = e.sheets.find((y) => y.id === g.sheetId)?.cells[g.address];
    if (
      ((!b || b.raw === null) &&
        failWithMessage(
          "Formula dependencies must be populated scalar cells; ambiguous blank or spill-child dependencies cannot be screened.",
        ),
      isFormulaInputValue(b))
    ) {
      const y = resolveDefinedNameValue(String(b.raw), e, g.sheetId).replace(
        /"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g,
        "",
      );
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
      ...e,
      sheets: e.sheets
        .filter((g) => n.allowedSheetIds.includes(g.id))
        .map((g) => ({
          ...g,
          cells: Object.fromEntries(
            Object.entries(g.cells).filter(([b]) =>
              c.has(buildCompositeKey(g.id, b)),
            ),
          ),
        })),
    },
    p = calculateWorkbookWithSettings(f);
  for (const g of l.nodes) {
    const b = p(g.sheetId, g.address);
    (p.getOperationalErrors().length &&
      failWithMessage(
        "An operational calculation limit or cycle prevents screening, even when a formula handles the visible error.",
      ),
      p.getSpillRange(g.sheetId, g.address) &&
        failWithMessage(
          "The screen requires scalar fields; a formula produced an array/spill result.",
        ),
      isFormulaInputValue(
        e.sheets.find((y) => y.id === g.sheetId)?.cells[g.address],
      ) &&
        typeof b == "string" &&
        /^#(?:[A-Z0-9_/.]+[!?]?)/.test(b) &&
        failWithMessage(
          "A selected formula or dependency has a calculation error. Correct it before screening.",
        ),
      g.sheetId === r.id && o.includes(g.address) && a.set(g.address, b));
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
function normalizeMaterialityRequest(e, r, t = {}) {
  const n = normalizeMaterialityConfig(e, r);
  return {
    ...buildMaterialitySourceEvidence(
      e,
      {
        sheetId: n.sheetId,
        range: n.range,
        headers: n.headers,
        columns: [n.amountColumn, n.currencyColumn, n.dateColumn],
      },
      t,
    ),
    request: n,
  };
}
function buildMaterialitySourceEvidence(e, r, t = {}) {
  (!r ||
    typeof r.sheetId != "string" ||
    typeof r.range != "string" ||
    typeof r.headers != "boolean" ||
    !Array.isArray(r.columns) ||
    r.columns.length > 500) &&
    failWithMessage(
      "Provide a bounded source range and scalar field selection.",
    );
  const n = parseRangeAddress(r.range),
    o = e.sheets.find((m) => m.id === r.sheetId);
  (!o ||
    n.lastRow >= o.rowCount ||
    n.lastCol >= o.colCount ||
    (n.lastRow - n.firstRow + 1) * (n.lastCol - n.firstCol + 1) >
      materialityScreenLimits.sourceCells ||
    r.columns.some(
      (m) => !Number.isInteger(m) || m < n.firstCol || m > n.lastCol,
    )) &&
    failWithMessage(
      "Choose scalar fields inside one source rectangle of at most 500 cells and within worksheet dimensions.",
    );
  const a = {
      ...r,
      range: `${formatCellAddress(n.firstRow, n.firstCol)}:${formatCellAddress(n.lastRow, n.lastCol)}`,
    },
    s = normalizeScreenHostScope(e, t),
    d = resolveMaterialitySourceRange(e, a, s),
    u = parseRangeAddress(a.range);
  for (let m = u.firstRow; m <= u.lastRow; m++)
    for (let v = u.firstCol; v <= u.lastCol; v++) {
      const g = d.cells[formatCellAddress(m, v)]?.raw;
      typeof g == "number" &&
        !Number.isFinite(g) &&
        failWithMessage(
          "Selected raw evidence contains a nonfinite number and cannot be preserved in JSON. Correct it before screening.",
        );
    }
  const l = [];
  for (let m = u.firstRow + Number(a.headers); m <= u.lastRow; m++)
    for (const v of new Set(a.columns)) l.push(formatCellAddress(m, v));
  const c = inspectScreenFormulaCells(e, d, l, s),
    h = [];
  for (let m = u.firstRow + Number(a.headers); m <= u.lastRow; m++)
    h.push({
      sourceRow: m + 1,
      sourceRange: `${formatCellAddress(m, u.firstCol)}:${formatCellAddress(m, u.lastCol)}`,
      cells: Array.from({ length: u.lastCol - u.firstCol + 1 }, (v, g) =>
        structuredClone(
          d.cells[formatCellAddress(m, u.firstCol + g)] ?? { raw: null },
        ),
      ),
      values: Array.from({ length: u.lastCol - u.firstCol + 1 }, (v, g) => {
        const b = formatCellAddress(m, u.firstCol + g);
        return c.has(b) ? c.get(b) : (d.cells[b]?.raw ?? null);
      }),
    });
  const f = {
    context: s,
    sheetId: d.id,
    sheetName: d.name,
    rows: h,
    headers: Array.from({ length: u.lastCol - u.firstCol + 1 }, (m, v) =>
      a.headers
        ? String(
            d.cells[formatCellAddress(u.firstRow, u.firstCol + v)]?.raw ??
              `Column ${columnIndexToLetters(u.firstCol + v)}`,
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
function runMaterialityScreen(e, r, t = {}) {
  const n = normalizeMaterialityRequest(e, r, t),
    { request, context } = n,
    s = parseRangeAddress(request.range),
    d = [],
    u = {
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
  let l = 0,
    c = 0,
    h = 0;
  for (const m of n.rows) {
    const v = (I) => m.values[I - s.firstCol],
      g = v(request.amountColumn),
      b = v(request.currencyColumn),
      y = v(request.dateColumn),
      w = typeof g == "number" && Number.isFinite(g) ? g : null,
      E =
        typeof b == "string" && /^[A-Za-z]{3}$/.test(b)
          ? b.toUpperCase()
          : null,
      A = normalizeDateToIso(y),
      k = [];
    (w === null &&
      (k.push(
        "Amount is missing or is not a finite numeric value (numeric text is not coerced).",
      ),
      u.invalidAmounts++),
      (!E || E !== request.currency) &&
        (k.push(
          E
            ? `Currency ${E} differs from the declared ${request.currency}; conversion is unsupported.`
            : "Currency is missing or is not a three-letter identifier.",
        ),
        u.currencyExceptions++),
      A
        ? (A < request.effectiveFrom || A > request.effectiveTo) &&
          (k.push("Date is outside the inclusive effective period."),
          u.outsidePeriod++)
        : (k.push(
            "Date is missing or invalid; use ISO YYYY-MM-DD or an integer Excel 1900-system date serial (except 60).",
          ),
          u.invalidDates++));
    const F =
      w === null ? null : request.amountMode === "absolute" ? Math.abs(w) : w;
    (w !== null && E === request.currency ? (l += w) : h++,
      k.length || (u.eligible++, (c += w)),
      d.push({
        ...m,
        amount: w,
        currency: E,
        date: A,
        comparedAmount: F,
        disposition: k.length
          ? "excluded"
          : compareAmountToThreshold(F, request.threshold, request.comparator)
            ? "flagged"
            : "not-flagged",
        reasons: k,
      }));
  }
  if (u.currencyExceptions)
    for (const m of d)
      m.disposition !== "excluded" &&
        ((m.disposition = "not-screened"),
        m.reasons.push(
          "Screen blocked by currency exceptions; no pass or exception conclusion is available.",
        ));
  u.population = d.length;
  for (const m of d)
    m.disposition === "flagged"
      ? u.flagged++
      : m.disposition === "not-flagged"
        ? u.notFlagged++
        : m.disposition === "excluded"
          ? u.excluded++
          : u.notScreened++;
  u.screened = u.flagged + u.notFlagged;
  const f = u.currencyExceptions
    ? null
    : d
        .filter((m) => m.disposition === "flagged")
        .reduce((m, v) => m + v.amount, 0);
  [l, c, l - c, f].some((m) => m !== null && !Number.isFinite(m)) &&
    failWithMessage(
      "The control total exceeds finite numeric capacity. Reduce the population or magnitude.",
    );
  const p = {
    schemaVersion: 1,
    kind: "materiality-screen",
    algorithm: "ledgerglass-materiality-1",
    workbookId: e.id,
    revision: e.revision,
    sheetId: n.sheetId,
    sheetName: n.sheetName,
    generatedAt: new Date().toISOString(),
    request: structuredClone(request),
    context,
    status: u.currencyExceptions
      ? "blocked"
      : u.excluded
        ? "complete-with-exclusions"
        : "complete",
    headers: n.headers,
    rows: d,
    counts: u,
    totals: {
      currency: request.currency,
      knownSourceAmount: l,
      eligibleAmount: c,
      excludedKnownAmount: l - c,
      flaggedAmount: f,
      unknownOrOtherCurrencyRows: h,
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
      snapshot: JSON.stringify(e),
      request: JSON.stringify(request),
      context: JSON.stringify(context),
      content: JSON.stringify(p),
    }),
    p
  );
}
function checkMaterialityReportFreshness(e, r, t, n = {}) {
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
    const a = normalizeScreenHostScope(e, n);
    if (
      (resolveMaterialitySourceRange(e, r.request, a),
      JSON.stringify(a) !== o.context)
    )
      return {
        current: false,
        disclose: false,
        reason:
          "The authorized or hidden-content scope changed. Run the screen again.",
      };
    if (JSON.stringify(e) !== o.snapshot)
      return {
        current: false,
        disclose: false,
        reason: "The source workbook changed. Run the screen again.",
      };
    const s = normalizeMaterialityConfig(e, t);
    return JSON.stringify(s) !== o.request
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
function exportMaterialityReportJson(e, r, t, n = {}) {
  const o = checkMaterialityReportFreshness(e, r, t, n);
  return (o.current || failWithMessage(o.reason), JSON.stringify(r, null, 2));
}
function parseMaterialityReportText(e, r, t = {}) {
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
  return loadMaterialityConfiguration(e, n, t);
}
function loadMaterialityConfiguration(e, r, t = {}) {
  (!r || typeof r != "object" || Array.isArray(r)) &&
    failWithMessage("Choose a supported screen configuration or report.");
  const n = r;
  (n.schemaVersion !== 1 || n.kind !== "materiality-screen") &&
    failWithMessage(
      "The materiality screen configuration version is unsupported.",
    );
  const o = normalizeMaterialityConfig(e, n.request);
  return (
    resolveMaterialitySourceRange(e, o, normalizeScreenHostScope(e, t)),
    o
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
function normalizeAgingDateInput(e, r) {
  if (typeof e == "string") return parseAgingIsoDate(e) === null ? null : e;
  if (
    typeof e != "number" ||
    !Number.isInteger(e) ||
    (r === "1900" && (e < 1 || e > 2958465 || e === 60)) ||
    (r === "1904" && (e < 0 || e > 2957003))
  )
    return null;
  const t = r === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  return new Date(t + (r === "1900" && e > 60 ? e - 1 : e) * dayInMilliseconds)
    .toISOString()
    .slice(0, 10);
}
function defaultMaterialityConfig(e) {
  return {
    sheetId: e.sheetId,
    range: e.range,
    headers: e.headers,
    amountColumn: e.amountColumn,
    currencyColumn: e.currencyColumn,
    dateColumn: e.dateColumn,
    currency: e.currency,
    threshold: 0,
    amountMode: "absolute",
    comparator: "gte",
    effectiveFrom: "0000-01-01",
    effectiveTo: "9999-12-31",
  };
}
function normalizeAgingConfiguration(e, r, t) {
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
  ((!r ||
    typeof r != "object" ||
    Array.isArray(r) ||
    Object.keys(r).length !== n.length ||
    n.some((d) => !Object.hasOwn(r, d))) &&
    throwErrorMessage("Provide a complete, supported aging configuration."),
    (!["due", "document"].includes(r.dateBasis) ||
      !["1900", "1904"].includes(r.dateSystem)) &&
      throwErrorMessage(
        "Choose due or document date and an explicit 1900 or 1904 serial date system.",
      ),
    parseAgingIsoDate(r.asOf) === null &&
      throwErrorMessage(
        "Choose a valid fixed as-of date in ISO YYYY-MM-DD format.",
      ),
    (!Array.isArray(r.buckets) ||
      r.buckets.length < 1 ||
      r.buckets.length > agingScheduleLimits.buckets) &&
      throwErrorMessage("Configure between 1 and 24 aging buckets."));
  let o = 0;
  const a = new Set();
  for (const [d, u] of r.buckets.entries())
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
        (d === r.buckets.length - 1
          ? u.to !== null
          : !Number.isSafeInteger(u.to) || u.to < u.from || u.to > 3652424)) &&
        throwErrorMessage(
          "Buckets must cover every nonnegative day exactly once: start at 0, use contiguous inclusive integer bounds, and leave only the final upper bound open.",
        ),
      (o = u.to === null ? 0 : u.to + 1));
  const s = parseMaterialityReportText(
    e,
    JSON.stringify({
      schemaVersion: 1,
      kind: "materiality-screen",
      request: defaultMaterialityConfig(r),
    }),
    t,
  );
  return {
    sheetId: s.sheetId,
    range: s.range,
    headers: s.headers,
    amountColumn: s.amountColumn,
    currencyColumn: s.currencyColumn,
    dateColumn: s.dateColumn,
    currency: s.currency,
    dateBasis: r.dateBasis,
    dateSystem: r.dateSystem,
    asOf: r.asOf,
    buckets: r.buckets.map((d) => ({ ...d, label: d.label.trim() })),
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
function buildAgingSchedule(e, r, t = {}) {
  const n = normalizeAgingConfiguration(e, r, t),
    o = normalizeMaterialityRequest(e, defaultMaterialityConfig(n), t),
    a = parseRangeAddress(n.range).firstCol,
    s = parseAgingIsoDate(n.asOf),
    d = o.rows.map((y) => {
      const w = y.values[n.amountColumn - a],
        E = y.values[n.currencyColumn - a],
        A = typeof w == "number" && Number.isFinite(w) ? w : null,
        k =
          typeof E == "string" && /^[A-Za-z]{3}$/.test(E)
            ? E.toUpperCase()
            : null,
        F = y.values[n.dateColumn - a],
        I = normalizeAgingDateInput(F, n.dateSystem),
        O = I === null ? null : s - parseAgingIsoDate(I),
        j = [],
        z = [];
      (F === null || F === ""
        ? (j.push("missing-date"),
          z.push("Date is missing; no date is inferred."))
        : I === null
          ? (j.push("invalid-date"),
            z.push(
              `Date must be strict ISO text or an integer ${n.dateSystem}-system calendar serial; time fractions and fictitious 1900-02-29 are unsupported.`,
            ))
          : O < 0 &&
            (j.push("future-date"),
            z.push(
              "Date is after the fixed as-of date; future items are excluded and separately counted.",
            )),
        A === null &&
          (j.push("invalid-amount"),
          z.push(
            "Amount is missing or is not a finite number; numeric text is not coerced.",
          )),
        k !== n.currency &&
          (j.push("unsupported-currency"),
          z.push(
            `Currency is missing, invalid or differs from declared ${n.currency}; no conversion or cross-currency aggregation is performed.`,
          )));
      const V = j.length
        ? null
        : n.buckets.findIndex(
            (D) => O >= D.from && (D.to === null || O <= D.to),
          );
      return (
        V === -1 &&
          throwErrorMessage(
            "The configured aging buckets do not cover a valid source row.",
          ),
        {
          sourceRow: y.sourceRow,
          sourceRange: y.sourceRange,
          values: structuredClone(y.values),
          cells: structuredClone(y.cells),
          amount: A,
          currency: k,
          date: I,
          ageDays: O,
          bucket: V,
          disposition: j.length ? "excluded" : "aged",
          exceptions: j,
          reasons: z,
        }
      );
    }),
    u = d.filter((y) => y.amount !== null && y.currency === n.currency),
    l = d.filter((y) => y.disposition === "aged"),
    c = u.filter((y) => y.disposition === "excluded"),
    h = (y) => addDecimalAmounts(y.map((w) => w.amount)),
    f = h(u),
    p = h(l),
    m = h(c),
    v = n.buckets.map((y, w) => {
      const E = l.filter((k) => k.bucket === w),
        A = h(E);
      return {
        ...y,
        count: E.length,
        amount: assertAgingControlTotal(A),
        exactAmount: A,
      };
    }),
    g = (y) => d.filter((w) => w.exceptions.includes(y)).length,
    b = {
      schemaVersion: 1,
      kind: "aging-schedule",
      algorithm: "ledgerglass-aging-1",
      workbookId: e.id,
      revision: e.revision,
      sheetId: o.sheetId,
      sheetName: o.sheetName,
      generatedAt: new Date().toISOString(),
      request: structuredClone(n),
      context: structuredClone(o.context),
      status: l.length === d.length ? "complete" : "complete-with-exclusions",
      headers: structuredClone(o.headers),
      rows: d,
      buckets: v,
      counts: {
        population: d.length,
        aged: l.length,
        excluded: d.length - l.length,
        missingDates: g("missing-date"),
        invalidDates: g("invalid-date"),
        futureDates: g("future-date"),
        unsupportedCurrencies: g("unsupported-currency"),
        missingCurrencies: d.filter((y) => {
          const w = y.values[n.currencyColumn - a];
          return w === null || w === "";
        }).length,
        invalidAmounts: g("invalid-amount"),
      },
      totals: {
        currency: n.currency,
        knownSourceAmount: assertAgingControlTotal(f),
        agedAmount: assertAgingControlTotal(p),
        excludedKnownAmount: assertAgingControlTotal(m),
        unknownOrOtherCurrencyRows: d.length - u.length,
        exactKnownSourceAmount: f,
        exactAgedAmount: p,
        exactExcludedKnownAmount: m,
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
      snapshot: JSON.stringify(e),
      context: JSON.stringify(o.context),
      request: JSON.stringify(n),
      content: JSON.stringify(b),
    }),
    b
  );
}
function checkAgingReportFreshness(e, r, t, n = {}) {
  const o = _z.get(r);
  if (!o || o.content !== JSON.stringify(r))
    return {
      current: false,
      disclose: false,
      reason:
        "This aging report is untrusted or changed. Run the schedule again.",
    };
  try {
    const a = normalizeAgingConfiguration(e, t, n),
      s = {
        includeHidden: !!n.includeHidden,
        allowedSheetIds: [
          ...new Set(n.allowedSheetIds ?? e.sheets.map((d) => d.id)),
        ].sort(),
      };
    return o.snapshot !== JSON.stringify(e) || o.context !== JSON.stringify(s)
      ? {
          current: false,
          disclose: false,
          reason:
            "The source or authorized scope changed. Run the aging schedule again.",
        }
      : JSON.stringify(a) !== o.request
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
function exportAgingReportJson(e, r, t, n = {}) {
  const o = checkAgingReportFreshness(e, r, t, n);
  return (o.current || throwErrorMessage(o.reason), JSON.stringify(r, null, 2));
}
function loadAgingConfiguration(e, r, t = {}) {
  return (
    (!r ||
      typeof r != "object" ||
      Array.isArray(r) ||
      r.schemaVersion !== 1 ||
      r.kind !== "aging-schedule") &&
      throwErrorMessage("Choose a supported aging configuration or report."),
    normalizeAgingConfiguration(e, r.request, t)
  );
}
const listRangeColumnIndexes = (e) => {
    try {
      const r = parseRangeAddress(e);
      return r.lastCol - r.firstCol < 500
        ? Array.from(
            { length: r.lastCol - r.firstCol + 1 },
            (t, n) => r.firstCol + n,
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
    n = useCancellableOperation(),
    o = ne.useId(),
    a = ne.useRef(null),
    s = ne.useRef(true),
    d = ne.useRef(0),
    u = ne.useRef(""),
    l = listRangeColumnIndexes(e.range),
    [c, h] = ne.useState(e.range),
    [f, p] = ne.useState(true),
    [m, v] = ne.useState(l[0] ?? 0),
    [g, b] = ne.useState(l[1] ?? 1),
    [y, w] = ne.useState(l[2] ?? 2),
    [E, A] = ne.useState("USD"),
    [k, F] = ne.useState("due"),
    [I, O] = ne.useState(e.workbook.datePolicy?.dateSystem ?? "1900"),
    [j, z] = ne.useState(new Date().toISOString().slice(0, 10)),
    [V, D] = ne.useState(structuredClone(defaultAgingBuckets)),
    [B, _] = ne.useState(false),
    [C, T] = ne.useState(null),
    [N, R] = ne.useState(""),
    [L, $] = ne.useState(""),
    [U, G] = ne.useState(""),
    [H, Q] = ne.useState(""),
    [X$1, J] = ne.useState(0),
    [q, W] = ne.useState("all"),
    K = {
      sheetId: e.sheetId,
      range: c,
      headers: f,
      amountColumn: m,
      currencyColumn: g,
      dateColumn: y,
      currency: E,
      dateBasis: k,
      dateSystem: I,
      asOf: j,
      buckets: V,
    },
    te = {
      readOnly: e.readOnly,
      includeHidden: !!e.allowHidden && B,
      allowedSheetIds: e.allowedSheetIds,
    },
    ae = e.workbook.sheets.find((Me) => Me.id === e.sheetId),
    oe = [
      ...new Set(e.allowedSheetIds ?? e.workbook.sheets.map((Me) => Me.id)),
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
        (Me) =>
          q === "all" || q === Me.disposition || Me.exceptions.includes(q),
      ) ?? [],
    we = Math.max(1, Math.ceil(se.length / agingScheduleLimits.pageSize)),
    Ee = Math.min(X$1, we - 1),
    Ae = listRangeColumnIndexes(c),
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
    const Me = document.activeElement;
    return (
      a.current?.querySelector('[data-ui-label="Aging source range"]')?.focus(),
      () => {
        ((s.current = false), d.current++, Me?.isConnected && Me.focus());
      }
    );
  }, []);
  const Te = () => {
      d.current++;
      try {
        (T(buildAgingSchedule(e.workbook, K, te)),
          R(he),
          $(""),
          Q(""),
          W("all"),
          J(0));
      } catch (Me) {
        (T(null), ye(Me));
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
      } catch (Me) {
        ye(Me);
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
        const Je = await importNativeJsonFile(Me, n.begin());
        if (!$e()) return;
        const ze = loadAgingConfiguration(e.workbook, Je, te);
        if (ze.sheetId !== e.sheetId)
          throw Error(
            "Open the saved configuration's source worksheet before loading it.",
          );
        (h(ze.range),
          p(ze.headers),
          v(ze.amountColumn),
          b(ze.currencyColumn),
          w(ze.dateColumn),
          A(ze.currency),
          F(ze.dateBasis),
          O(ze.dateSystem),
          z(ze.asOf),
          D(ze.buckets),
          T(null),
          $(""),
          Q(
            "Aging configuration loaded. Run it against the current source; saved conclusions are not reused.",
          ));
      } catch (Je) {
        $e() && ye(toFileProcessingError(Je));
      }
    },
    ht = (Me, ke) =>
      D((je) => je.map(($e, Je) => (Je === Me ? { ...$e, ...ke } : $e)));
  return i.jsx("div", {
    className: "aging-backdrop",
    onMouseDown: (Me) => {
      Me.target === Me.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      className: "aging-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${o}-title`,
      ref: a,
      onKeyDown: (Me) => {
        var je, $e;
        if (
          (Me.stopPropagation(),
          Me.key === "Escape" && (Me.preventDefault(), e.onClose()),
          Me.key === "Tab")
        ) {
          const Je = Array.from(
            a.current?.querySelectorAll(
              'button:not(:disabled), input:not(:disabled), select:not(:disabled), summary, [tabindex="0"]',
            ) ?? [],
          ).filter(
            (ze) =>
              !ze.closest("details:not([open])") || ze.tagName === "SUMMARY",
          );
          Me.shiftKey && document.activeElement === Je[0]
            ? (Me.preventDefault(), (je = Je.at(-1)) == null || je.focus())
            : !Me.shiftKey &&
              document.activeElement === Je.at(-1) &&
              (Me.preventDefault(), ($e = Je[0]) == null || $e.focus());
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
                  value: de ? c : "",
                  onChange: (Me) => h(Me.target.value),
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
                  onChange: (Me) => p(Me.target.checked),
                }),
                i.jsx(LocalizedText, {
                  children: "First row contains headers",
                }),
              ],
            }),
            [
              ["Amount column", m, v],
              ["Currency column", g, b],
              ["Date column", y, w],
            ].map(([Me, ke, je]) =>
              i.jsxs(
                "label",
                {
                  children: [
                    Me,
                    i.jsx("select", {
                      "aria-label": Me,
                      value: ke,
                      onChange: ($e) => je(Number($e.target.value)),
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
                  value: E,
                  onChange: (Me) => A(Me.target.value),
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
                  value: k,
                  onChange: (Me) => F(Me.target.value),
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
                  value: j,
                  onChange: (Me) => z(Me.target.value),
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
                  value: I,
                  onChange: (Me) => O(Me.target.value),
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
                    children: V.map((Me, ke) =>
                      i.jsxs(
                        "tr",
                        {
                          children: [
                            i.jsx("td", {
                              children: i.jsx("input", {
                                "aria-label": `Bucket ${ke + 1} label`,
                                value: Me.label,
                                maxLength: 80,
                                onChange: (je) =>
                                  ht(ke, { label: je.target.value }),
                              }),
                            }),
                            i.jsx("td", {
                              children: i.jsx(LocalizedNumberField, {
                                "aria-label": `Bucket ${ke + 1} from`,
                                min: "0",
                                step: "1",
                                value: Number.isFinite(Me.from) ? Me.from : "",
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
                                  Me.to === null
                                    ? ""
                                    : Number.isFinite(Me.to)
                                      ? Me.to
                                      : "",
                                onValueChange: (je) =>
                                  ht(ke, { to: je === "" ? null : Number(je) }),
                                lang: language("Open-ended"),
                              }),
                            }),
                            i.jsx("td", {
                              children: i.jsx("button", {
                                "aria-label": `Remove bucket ${ke + 1}`,
                                disabled: V.length === 1,
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
              disabled: V.length >= agingScheduleLimits.buckets,
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
                checked: B,
                onChange: (Me) => _(Me.target.checked),
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
              onClick: Te,
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
                  onChange: (Me) => {
                    const ke = Me.target.files?.[0];
                    ((Me.target.value = ""), qe(ke));
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
                    children: ge.buckets.map((Me, ke) =>
                      i.jsxs(
                        "tr",
                        {
                          children: [
                            i.jsx("th", { scope: "row", children: Me.label }),
                            i.jsxs("td", {
                              children: [Me.from, "–", Me.to ?? "unbounded"],
                            }),
                            i.jsx("td", { children: Me.count }),
                            i.jsx("td", { children: Me.exactAmount }),
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
                    onChange: (Me) => {
                      (W(Me.target.value), J(0));
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
                          ge.headers.map((Me, ke) =>
                            i.jsxs(
                              "th",
                              {
                                children: [
                                  Me,
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
                              Me.cells.map((ke, je) =>
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
                                          children: JSON.stringify(ke, null, 2),
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
                  ge.notes.map((Me) => i.jsx("p", { children: Me }, Me)),
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
