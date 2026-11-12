import {
  assertWorkbookWritable,
  eI,
  formatCellAddress,
} from "./04-cell-input-paste-and-fill.js";
import { isFormulaInputValue } from "./11-local-workbook-store-and-versions.js";
import { formatSortRangeAddress } from "./44-clean-and-reshape.js";
import {
  rangesIntersect,
  isCellInsideRange,
  quoteSheetName,
  throwWorkbookError,
} from "./03-workbook-model.js";
import { X, ArrowRight } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  summarizeNumberFormat,
  formatNumberWithPattern,
  builtInDateTimeFormats,
  columnWidthAt,
  findSheetOrThrow,
  supportedNumberLocales,
  buildNumberFormatCode,
  getImportedNumberFormatWarning,
  supportedCurrencyCodes,
  Uv,
  calculateWorkbookWithSettings,
  isCellAddressInRanges,
  _t,
  cn,
  parseRangeAddress,
  parseSelectionAreas,
  rowHeightAt,
  normalizeCellRecordInput,
  assertSheetUnprotected,
  validateNumberFormatSettings,
  applyNumberFormatTransaction,
  pQ,
  normalizeWorkbookInput,
  createAxisVisibilityIndex,
  columnIndexToLetters,
  checkDataValidationBatch,
  formatCellValueWithNumberFormat,
  countSelectionCells,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  pasteOperationLimits,
  hasCellContent,
  parseWholeColumnReference,
  assertCutSourceWithinLimit,
  createCutSourceSnapshot,
  rangeContainsRange,
  parseMovableA1Rectangle,
  assertFormulaReferencesMovable,
  verifyCutSourceSnapshot,
  isUnsafeStructuredReference,
  assertRangeWithinBoundary,
} from "./47-fill-and-paste-special.js";

function remapReferenceForMove(e, sheetId, reference) {
  if (sheetId !== e.source.id)
    return { sheetId: sheetId, reference: reference, changed: false };
  const wholeColumnReference = parseWholeColumnReference(reference);
  if (!rangesIntersect(wholeColumnReference, e.from))
    return { sheetId: sheetId, reference: reference, changed: false };
  if (rangeContainsRange(e.from, wholeColumnReference))
    return {
      sheetId: e.target.id,
      reference: parseMovableA1Rectangle(reference, e.dr, e.dc),
      changed: true,
    };
  const o = {
      firstRow: Math.max(wholeColumnReference.firstRow, e.from.firstRow),
      lastRow: Math.min(wholeColumnReference.lastRow, e.from.lastRow),
      firstCol: Math.max(wholeColumnReference.firstCol, e.from.firstCol),
      lastCol: Math.min(wholeColumnReference.lastCol, e.from.lastCol),
    },
    a = {
      firstRow: o.firstRow + e.dr,
      lastRow: o.lastRow + e.dr,
      firstCol: o.firstCol + e.dc,
      lastCol: o.lastCol + e.dc,
    };
  if (
    e.source.id === e.target.id &&
    rangeContainsRange(wholeColumnReference, a)
  )
    return { sheetId: sheetId, reference: reference, changed: false };
  throwWorkbookError(
    `Reference ${e.source.name}!${reference} includes cells outside the moved block. This move would require a split or cross-sheet range; use explicit references first.`,
  );
}
function qW(book, r, t = new Set()) {
  const n = `${r.scopeSheetId ?? ""}:${r.name.toLowerCase()}`;
  return t.has(n) || t.size > 32
    ? false
    : (t.add(n),
      assertFormulaReferencesMovable(r.refersTo).every((o) => {
        if (o.kind === "dynamic" || isUnsafeStructuredReference(o))
          return false;
        if (o.kind === "reference") return !!o.sheet;
        if (o.kind === "name" && !o.sheet) {
          const a = (book.definedNames ?? []).find(
            (s) =>
              !s.scopeSheetId && s.name.toLowerCase() === o.name.toLowerCase(),
          );
          return !a || qW(book, a, new Set(t));
        }
        return true;
      }));
}
function XW(book, r, t = new Set()) {
  const n = `${r.scopeSheetId ?? ""}:${r.name.toLowerCase()}`;
  return t.has(n) || t.size > 32
    ? true
    : (t.add(n),
      assertFormulaReferencesMovable(r.refersTo).some((o) => {
        if (isUnsafeStructuredReference(o)) return true;
        if (o.kind !== "name" || o.sheet) return false;
        const a =
          (book.definedNames ?? []).find(
            (s) =>
              s.scopeSheetId === r.scopeSheetId &&
              s.name.toLowerCase() === o.name.toLowerCase(),
          ) ??
          (book.definedNames ?? []).find(
            (s) =>
              !s.scopeSheetId && s.name.toLowerCase() === o.name.toLowerCase(),
          );
        return !!a && XW(book, a, new Set(t));
      }));
}
function repairReferencesForMove(e, refersTo, t, n, o = false) {
  ((e.characters += refersTo.length),
    e.characters > pasteOperationLimits.formulaCharacters &&
      throwWorkbookError(
        "Reference repair exceeds 10 million formula characters. Simplify the workbook before moving cells.",
      ));
  const a = assertFormulaReferencesMovable(refersTo);
  let s = "",
    d = 0;
  for (const l of a) {
    let text = l.text;
    if (
      (l.kind === "dynamic" &&
        throwWorkbookError(
          `${l.name} can refer to cells dynamically. Replace dynamic references with explicit A1 references before moving cells.`,
        ),
      isUnsafeStructuredReference(l) &&
        o &&
        throwWorkbookError(
          "A current-row structured reference cannot be safely moved out of its table context.",
        ),
      l.kind === "name" && o && !l.sheet)
    ) {
      const h = e.book.definedNames ?? [],
        f = h.find(
          (g) =>
            g.scopeSheetId === t &&
            g.name.toLowerCase() === l.name.toLowerCase(),
        ),
        p = h.find(
          (g) =>
            !g.scopeSheetId && g.name.toLowerCase() === l.name.toLowerCase(),
        ),
        m = h.find(
          (g) =>
            g.scopeSheetId === n &&
            g.name.toLowerCase() === l.name.toLowerCase(),
        ),
        v = f ?? p;
      (v &&
        XW(e.book, v) &&
        throwWorkbookError(
          `Name ${l.name} contains a current-row structured reference. Use an explicit qualified cell reference before moving this formula.`,
        ),
        t !== n &&
          (f || m || (p && !qW(e.book, p))) &&
          throwWorkbookError(
            `Name ${l.name} depends on worksheet scope. Keep the formula on its sheet or replace that name with explicit qualified references before moving it.`,
          ));
    }
    if (l.kind === "reference") {
      const sheetId = l.sheet
        ? e.book.sheets.find(
            (sheet) => sheet.name.toLowerCase() === l.sheet.toLowerCase(),
          )?.id
        : t;
      if (!sheetId)
        !l.sheet &&
          rangesIntersect(parseWholeColumnReference(l.reference), e.from) &&
          throwWorkbookError(
            "A workbook-level name has a context-dependent unqualified reference. Qualify its worksheet before moving cells.",
          );
      else {
        const f = remapReferenceForMove(e, sheetId, l.reference);
        f.changed
          ? (text = `${l.sheet || f.sheetId !== n ? quoteSheetName(findSheetOrThrow(e.book, f.sheetId).name) + "!" : ""}${f.reference}`)
          : !l.sheet &&
            t !== n &&
            t &&
            (text = `${quoteSheetName(findSheetOrThrow(e.book, t).name)}!${l.reference}`);
      }
    }
    text !== l.text && ((s += refersTo.slice(d, l.start) + text), (d = l.end));
  }
  return s + refersTo.slice(d);
}
function computeDrawingFootprint(sheet, r) {
  const t = _t(r.anchor),
    n = (s, d, u, l) => {
      const axisVisibilityIndex = createAxisVisibilityIndex(
          s === "row" ? sheet.rowCount : sheet.colCount,
          cn(sheet, s),
        ),
        h = (p) =>
          s === "row" ? rowHeightAt(sheet, p) : columnWidthAt(sheet, p);
      if (axisVisibilityIndex.isHidden(d)) return [d, d];
      const f = (p) => {
        let m = d,
          v = 0;
        if (p >= 0)
          for (; p >= h(m);) {
            p -= h(m);
            const g = axisVisibilityIndex.next(m, 1);
            if (g === null) return m;
            ((m = g),
              ++v > 1e3 &&
                throwWorkbookError(
                  "Drawing footprint exceeds the move inspection budget.",
                ));
          }
        else
          for (; p < 0;) {
            const g = axisVisibilityIndex.next(m, -1);
            if (g === null) return 0;
            ((m = g),
              (p += h(m)),
              ++v > 1e3 &&
                throwWorkbookError(
                  "Drawing footprint exceeds the move inspection budget.",
                ));
          }
        return m;
      };
      return [f(u), f(u + l - 1e-4)];
    },
    o = n("row", t.row, r.offsetY, r.height),
    a = n("column", t.col, r.offsetX, r.width);
  return { firstRow: o[0], lastRow: o[1], firstCol: a[0], lastCol: a[1] };
}
function findMoveBlockers(e) {
  const formulaEntries = collectFormulaEntries(e.book);
  for (const [sheet, n] of [
    [e.source, e.from],
    [e.target, e.to],
  ]) {
    for (const o of sheet.merges ?? [])
      rangesIntersect(n, parseRangeAddress(o)) &&
        throwWorkbookError(
          "Unmerge the source and destination before moving cells.",
        );
    for (const o of sheet.tables ?? [])
      rangesIntersect(n, parseRangeAddress(o.range)) &&
        throwWorkbookError(
          `The move intersects table ${o.name}. Table-cell moves are not supported; use ordinary cells outside tables.`,
        );
    for (const o of [...(sheet.pivots ?? []), ...(sheet.queries ?? [])])
      rangesIntersect(n, parseRangeAddress(o.outputRange)) &&
        throwWorkbookError(
          "The move intersects a generated PivotTable or query output. Move ordinary cells instead.",
        );
    for (const o of [
      ...(sheet.validations ?? []),
      ...(sheet.conditionalFormats ?? []),
    ])
      rangesIntersect(n, parseRangeAddress(o.range)) &&
        throwWorkbookError(
          "The move intersects a validation or conditional-formatting rule. Remove or relocate the rule first.",
        );
    for (const o of sheet.objects ?? [])
      (isCellInsideRange(n, _t(o.anchor)) ||
        rangesIntersect(n, computeDrawingFootprint(sheet, o))) &&
        throwWorkbookError(
          `The move intersects drawing ${o.label || o.id}. Drawing footprints cannot be moved or partially overwritten by this command.`,
        );
    for (const formulaEntry of formulaEntries)
      formulaEntry.sheetId === sheet.id &&
        rangesIntersect(n, parseRangeAddress(formulaEntry.range)) &&
        throwWorkbookError(
          "The move intersects a dynamic-array footprint. Convert it to ordinary values or move cells outside the array.",
        );
  }
}
function collectFormulaEntries(book) {
  const workbookWithSettings = calculateWorkbookWithSettings(book),
    t = [];
  let n = 0,
    o = 0;
  for (const sheet of book.sheets)
    for (const [anchor, cell] of Object.entries(sheet.cells))
      if (isFormulaInputValue(cell)) {
        (++n > pasteOperationLimits.formulaCells ||
          (o += cell.raw.length) > pasteOperationLimits.formulaCharacters) &&
          throwWorkbookError(
            "Array inspection supports at most 10,000 formulas and 10 million formula characters. Simplify the workbook before moving cells.",
          );
        const range = workbookWithSettings.getSpillRange(sheet.id, anchor);
        if (range) {
          const rangeAddress = parseRangeAddress(range);
          (rangeAddress.firstRow !== rangeAddress.lastRow ||
            rangeAddress.firstCol !== rangeAddress.lastCol) &&
            t.push({ sheetId: sheet.id, anchor: anchor, range: range });
        }
      }
  return t;
}
function formulaEntriesSignature(book) {
  return collectFormulaEntries(book)
    .map((r) => `${r.sheetId}:${r.anchor}:${r.range}`)
    .sort().join(`
`);
}
function computeRangeMoveResult(
  book,
  sourceSheetId,
  sourceRange,
  targetSheetId,
  targetRange,
  a = {},
) {
  assertWorkbookWritable(a.readOnly);
  const sheetOrThrow = findSheetOrThrow(book, sourceSheetId),
    d = findSheetOrThrow(book, targetSheetId);
  (assertSheetUnprotected(sheetOrThrow), assertSheetUnprotected(d));
  const from = assertCutSourceWithinLimit(sheetOrThrow, sourceRange),
    l = assertRangeWithinBoundary(d, targetRange),
    c = from.lastRow - from.firstRow + 1,
    h = from.lastCol - from.firstCol + 1;
  (l.lastRow !== l.firstRow || l.lastCol !== l.firstCol) &&
    (l.lastRow - l.firstRow + 1 !== c || l.lastCol - l.firstCol + 1 !== h) &&
    throwWorkbookError(
      "Choose one destination cell or a rectangle with exactly the source dimensions.",
    );
  const f = {
    firstRow: l.firstRow,
    lastRow: l.firstRow + c - 1,
    firstCol: l.firstCol,
    lastCol: l.firstCol + h - 1,
  };
  (f.lastRow >= d.rowCount || f.lastCol >= d.colCount) &&
    throwWorkbookError("The moved cells would exceed the worksheet boundary.");
  const p =
    a.sourceSnapshot ??
    createCutSourceSnapshot(book, sourceSheetId, sourceRange);
  verifyCutSourceSnapshot(book, p, sourceSheetId, sourceRange);
  const m = {
      book: book,
      source: sheetOrThrow,
      target: d,
      from: from,
      to: f,
      dr: f.firstRow - from.firstRow,
      dc: f.firstCol - from.firstCol,
      characters: 0,
    },
    sortRangeAddress = formatSortRangeAddress(sourceRange),
    g = formatSortRangeAddress(
      `${formatCellAddress(f.firstRow, f.firstCol)}:${formatCellAddress(f.lastRow, f.lastCol)}`,
    ),
    metadataChanges = [],
    y = sourceSheetId === targetSheetId && m.dr === 0 && m.dc === 0;
  y || findMoveBlockers(m);
  const workbook = y ? book : structuredClone(book),
    E = new Map();
  let overwrittenCells = 0,
    referenceChanges = 0;
  const F = (B) => {
    (metadataChanges.push(B),
      metadataChanges.length > pasteOperationLimits.metadataChanges &&
        throwWorkbookError(
          "The move would change more than 500 metadata references. Narrow the source.",
        ));
  };
  if (
    (!y &&
      book.sheets.some((sheet) =>
        sheet.conditionalFormats?.some(
          (conditionalFormat) => conditionalFormat.kind === "formula",
        ),
      ) &&
      throwWorkbookError(
        "Range moves cannot repair formula-based conditional rules yet. Remove those rules before moving cells.",
      ),
    !y)
  ) {
    const B = pQ(workbook, sourceSheetId, sortRangeAddress, targetSheetId, g);
    B.watches && (workbook.formulaWatches = B.watches);
    for (const change of B.changes) F(change);
    for (const N of book.reviewThreads ?? [])
      !N.orphaned &&
        N.anchor.sheetId === targetSheetId &&
        isCellInsideRange(f, _t(N.anchor.address)) &&
        !(
          N.anchor.sheetId === sourceSheetId &&
          isCellInsideRange(from, _t(N.anchor.address))
        ) &&
        throwWorkbookError(
          "The destination contains an existing review discussion. Choose a destination without anchored discussions to preserve its evidence context.",
        );
    const _ = findSheetOrThrow(workbook, sourceSheetId),
      C = findSheetOrThrow(workbook, targetSheetId);
    for (let firstRow = from.firstRow; firstRow <= from.lastRow; firstRow++)
      for (let firstCol = from.firstCol; firstCol <= from.lastCol; firstCol++) {
        const cellAddress = formatCellAddress(firstRow, firstCol);
        hasCellContent(sheetOrThrow.cells[cellAddress]) &&
          delete _.cells[cellAddress];
      }
    for (let firstRow = from.firstRow; firstRow <= from.lastRow; firstRow++)
      for (let firstCol = from.firstCol; firstCol <= from.lastCol; firstCol++) {
        const cellAddress = formatCellAddress(firstRow, firstCol),
          $ = formatCellAddress(firstRow + m.dr, firstCol + m.dc),
          cell = sheetOrThrow.cells[cellAddress],
          G = d.cells[$];
        (!(sourceSheetId === targetSheetId && isCellInsideRange(from, _t($))) &&
          hasCellContent(G) &&
          JSON.stringify(G) !== JSON.stringify(cell) &&
          overwrittenCells++,
          hasCellContent(cell)
            ? (C.cells[$] = normalizeCellRecordInput(cell))
            : hasCellContent(C.cells[$]) && delete C.cells[$],
          E.set(`${targetSheetId}:${$}`, cellAddress));
      }
    let T = 0;
    for (const sheet of workbook.sheets)
      for (const [R, cell] of Object.entries(sheet.cells)) {
        ++T > workbookLimits.maxCells &&
          throwWorkbookError(
            "Range moves inspect at most 100,000 populated cells.",
          );
        const $ = E.has(`${sheet.id}:${R}`);
        if (isFormulaInputValue(cell)) {
          const G = repairReferencesForMove(
            m,
            cell.raw,
            $ ? sourceSheetId : sheet.id,
            sheet.id,
            $,
          );
          G !== cell.raw && ((cell.raw = G), referenceChanges++);
        }
        const hyperlink = cell.hyperlink;
        if (
          hyperlink?.kind === "internal" &&
          hyperlink.sheetId === sourceSheetId &&
          isCellInsideRange(from, _t(hyperlink.address))
        ) {
          const G = _t(hyperlink.address),
            before = `${sheetOrThrow.name}!${hyperlink.address}`;
          ((cell.hyperlink = {
            ...hyperlink,
            sheetId: targetSheetId,
            address: formatCellAddress(G.row + m.dr, G.col + m.dc),
          }),
            JSON.stringify(hyperlink) !== JSON.stringify(cell.hyperlink) &&
              (referenceChanges++,
              F({
                kind: "internal-link",
                label: `${sheet.name}!${R}`,
                before: before,
                after: `${d.name}!${cell.hyperlink.address}`,
              })));
        }
      }
    for (const sheet of workbook.sheets)
      for (const R of sheet.tables ?? [])
        for (const column of R.columns)
          if (column.calculatedFormula) {
            const after = repairReferencesForMove(
              m,
              column.calculatedFormula,
              sheet.id,
              sheet.id,
            );
            after !== column.calculatedFormula &&
              (assertSheetUnprotected(findSheetOrThrow(book, sheet.id)),
              F({
                kind: "table-column",
                label: `${sheet.name}: ${R.name}[${column.name}] calculated rule`,
                before: column.calculatedFormula,
                after: after,
              }),
              (column.calculatedFormula = after),
              referenceChanges++);
          }
    for (const N of workbook.definedNames ?? []) {
      const after = repairReferencesForMove(
        m,
        N.refersTo,
        N.scopeSheetId,
        N.scopeSheetId,
      );
      after !== N.refersTo &&
        (N.scopeSheetId &&
          assertSheetUnprotected(findSheetOrThrow(book, N.scopeSheetId)),
        F({
          kind: "defined-name",
          label: N.name,
          before: N.refersTo,
          after: after,
        }),
        (N.refersTo = after),
        referenceChanges++);
    }
    for (const N of workbook.reviewThreads ?? [])
      if (
        !N.orphaned &&
        N.anchor.sheetId === sourceSheetId &&
        isCellInsideRange(from, _t(N.anchor.address))
      ) {
        const R = _t(N.anchor.address),
          before = `${N.anchor.sheetName}!${N.anchor.address}`;
        ((N.anchor = {
          sheetId: targetSheetId,
          sheetName: d.name,
          address: formatCellAddress(R.row + m.dr, R.col + m.dc),
        }),
          F({
            kind: "review-anchor",
            label: N.id,
            before: before,
            after: `${d.name}!${N.anchor.address}`,
          }));
      }
    for (const sheet of workbook.sheets) {
      for (const R of [...(sheet.pivots ?? []), ...(sheet.queries ?? [])]) {
        const L = remapReferenceForMove(m, R.source.sheetId, R.source.range);
        L.changed &&
          (assertSheetUnprotected(findSheetOrThrow(book, sheet.id)),
          F({
            kind: "analysis-source",
            label: `${sheet.name}: ${R.name}`,
            before: `${findSheetOrThrow(book, R.source.sheetId).name}!${R.source.range}`,
            after: `${findSheetOrThrow(book, L.sheetId).name}!${L.reference}`,
          }),
          (R.source = { ...R.source, sheetId: L.sheetId, range: L.reference }),
          referenceChanges++);
      }
      for (const R of sheet.objects ?? [])
        if (R.kind === "chart" && R.range) {
          const L = remapReferenceForMove(m, sheet.id, R.range);
          L.changed &&
            (L.sheetId !== sheet.id &&
              throwWorkbookError(
                "A chart source cannot follow data to another worksheet in this implementation. Move the chart data on its current sheet.",
              ),
            assertSheetUnprotected(findSheetOrThrow(book, sheet.id)),
            F({
              kind: "chart-source",
              label: R.label || R.id,
              before: R.range,
              after: L.reference,
            }),
            (R.range = L.reference),
            referenceChanges++);
        }
    }
  }
  const previewChanges = [];
  let dependentCount = 0;
  if (!y)
    for (const sheet of workbook.sheets) {
      const sheetOrThrow = findSheetOrThrow(book, sheet.id),
        C = new Set([
          ...Object.keys(sheetOrThrow.cells),
          ...Object.keys(sheet.cells),
        ]);
      for (const address of C) {
        const cell = sheetOrThrow.cells[address],
          R = sheet.cells[address];
        JSON.stringify(cell) !== JSON.stringify(R) &&
          (assertSheetUnprotected(sheetOrThrow),
          previewChanges.push({
            sheetId: sheet.id,
            sheetName: sheet.name,
            address: address,
            ...(cell ? { before: normalizeCellRecordInput(cell) } : {}),
            ...(R ? { after: normalizeCellRecordInput(R) } : {}),
          }),
          previewChanges.length > pasteOperationLimits.changedCells &&
            throwWorkbookError(
              "A move may change at most 500 cells in total, including source clears, destination cells, and dependent references. Move a smaller block.",
            ),
          !(
            sheet.id === sourceSheetId && isCellInsideRange(from, _t(address))
          ) &&
            !(
              sheet.id === targetSheetId && isCellInsideRange(f, _t(address))
            ) &&
            dependentCount++);
      }
    }
  const j = previewChanges.length > 0 || metadataChanges.length > 0;
  j &&
    (formulaEntriesSignature(book) !== formulaEntriesSignature(workbook) &&
      throwWorkbookError(
        "The move would create, remove, or resize a dynamic-array footprint. Convert the affected array to ordinary values before moving cells.",
      ),
    (book.events.length >= 2e3 ||
      book.events.reduce((B, _) => B + _.changes.length, 0) +
        previewChanges.length >
        workbookLimits.maxCells) &&
      throwWorkbookError(
        "The move exceeds workbook history limits. Export the workbook before starting a new workpaper.",
      ),
    book.revision >= Number.MAX_SAFE_INTEGER - 1 &&
      throwWorkbookError("Workbook revision limit reached."),
    checkDataValidationBatch(
      workbook,
      previewChanges.filter(
        (previewChange) =>
          previewChange.before?.raw !== previewChange.after?.raw ||
          previewChange.before?.literal !== previewChange.after?.literal,
      ),
    ));
  const z = j ? normalizeWorkbookInput(workbook) : book,
    V = previewChanges.map((previewChange) => ({
      sheetId: previewChange.sheetId,
      address: previewChange.address,
      cell: normalizeCellRecordInput(previewChange.after ?? { raw: null }),
    })),
    D = {
      workbookId: book.id,
      revision: book.revision,
      snapshotGuard: p.snapshotGuard,
      sourceSheetId: sourceSheetId,
      sourceRange: sortRangeAddress,
      sourceSnapshot: { ...p },
      targetSheetId: targetSheetId,
      targetRange: g,
      targetCount: c * h,
      patches: V,
      previewPatches: V,
      count: previewChanges.length,
      dependentCount: dependentCount,
      overwrittenCells: overwrittenCells,
      referenceChanges: referenceChanges,
      metadataChanges: metadataChanges,
      previewChanges: previewChanges,
      summary: `Move ${sheetOrThrow.name}!${sortRangeAddress} to ${d.name}!${g}: ${previewChanges.length} cell changes, ${overwrittenCells} destination cells overwritten, ${metadataChanges.length} metadata changes.`,
    };
  return {
    plan: eI(D),
    next: j ? { ...z, revision: book.revision, events: book.events } : book,
  };
}
function buildRangeMovePlan(workbook, sheetId, range, n, o, a = {}) {
  return computeRangeMoveResult(workbook, sheetId, range, n, o, a).plan;
}
function applyRangeMovePreview(workbook, plan, t = {}) {
  (assertWorkbookWritable(t.readOnly),
    (!plan ||
      plan.workbookId !== workbook.id ||
      plan.revision !== workbook.revision ||
      plan.snapshotGuard !== JSON.stringify(workbook)) &&
      throwWorkbookError(
        "The workbook changed after this move preview. Cut the source again before moving it.",
      ));
  const rangeMoveResult = computeRangeMoveResult(
    workbook,
    plan.sourceSheetId,
    plan.sourceRange,
    plan.targetSheetId,
    plan.targetRange,
    { sourceSnapshot: plan.sourceSnapshot, readOnly: t.readOnly },
  );
  return (
    (JSON.stringify(rangeMoveResult.plan.previewChanges) !==
      JSON.stringify(plan.previewChanges) ||
      JSON.stringify(rangeMoveResult.plan.metadataChanges) !==
        JSON.stringify(plan.metadataChanges)) &&
      throwWorkbookError(
        "The move no longer matches its preview. Cut the source again.",
      ),
    rangeMoveResult.next
  );
}
function describeCellContents(cell) {
  if (!cell) return "(empty)";
  const r = [String(cell.raw ?? "(empty)")];
  return (
    cell.literal && r.push("Literal text"),
    cell.note !== undefined && r.push("Note: " + cell.note),
    cell.tickmark !== undefined && r.push("Tickmark: " + cell.tickmark),
    cell.hyperlink &&
      r.push(
        "Link: " +
          (cell.hyperlink.kind === "external"
            ? cell.hyperlink.url
            : cell.hyperlink.sheetId + "!" + cell.hyperlink.address) +
          (cell.hyperlink.label ? " (" + cell.hyperlink.label + ")" : ""),
      ),
    cell.numberFormat
      ? r.push(
          "Number display: " +
            Object.entries(cell.numberFormat)
              .map(([t, n]) => t + " = " + n)
              .join(", "),
        )
      : (cell.numFmt || cell.format) &&
        r.push("Number display: " + (cell.numFmt ?? cell.format)),
    cell.style &&
      Object.keys(cell.style).length &&
      r.push(
        "Cell style: " +
          Object.entries(cell.style)
            .map(([t, n]) => t + " = " + n)
            .join(", "),
      ),
    r.join(" · ")
  );
}
const staleMovePreviewMessage =
  "The workbook or destination changed. Cancel and cut the current cells again.";
function MoveCellsDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e);
  a.current = e;
  const s = ne.useRef({
      sheetId: e.sheetId,
      range: e.range,
      source: e.source,
      sourceName:
        e.workbook.sheets.find((sheet) => sheet.id === e.source.sheetId)
          ?.name ?? "Deleted worksheet",
      targetName:
        e.workbook.sheets.find((sheet) => sheet.id === e.sheetId)?.name ??
        "Deleted worksheet",
      book: JSON.stringify(e.workbook),
    }),
    [d, u] = ne.useState(""),
    [l, c] = ne.useState(0),
    h =
      s.current.book !== JSON.stringify(e.workbook) ||
      s.current.sheetId !== e.sheetId ||
      s.current.range !== e.range ||
      s.current.source !== e.source,
    [f] = ne.useState(() => {
      try {
        return {
          plan: buildRangeMovePlan(
            e.workbook,
            e.source.sheetId,
            e.source.range,
            e.sheetId,
            e.range,
            { sourceSnapshot: e.source, readOnly: e.readOnly },
          ),
          error: "",
        };
      } catch (error) {
        return {
          plan: null,
          error:
            error instanceof Error
              ? error.message
              : "The cells cannot be moved.",
        };
      }
    }),
    p = [
      ...(f.plan?.previewChanges ?? []).map((k) => ({
        label: `${k.sheetName}!${k.address}`,
        before: describeCellContents(k.before),
        after: describeCellContents(k.after),
      })),
      ...(f.plan?.metadataChanges ?? []).map((k) => ({
        label: k.label,
        before: k.before,
        after: k.after,
      })),
    ],
    m = Math.max(1, Math.ceil(p.length / 15)),
    v = Math.min(l, m - 1);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      o.current?.querySelector("button:not(:disabled)")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  function onClick() {
    try {
      const k = a.current;
      if (k.readOnly) throw Error("This workbook is read-only.");
      if (
        s.current.book !== JSON.stringify(k.workbook) ||
        s.current.sheetId !== k.sheetId ||
        s.current.range !== k.range ||
        s.current.source !== k.source
      )
        throw Error(staleMovePreviewMessage);
      if (!f.plan) throw Error(f.error);
      const F = applyRangeMovePreview(k.workbook, f.plan, {
        readOnly: k.readOnly,
      });
      (k.onWorkbookChange(F, "Move cells and repair references"),
        k.onSelectionChange(f.plan.targetRange),
        k.onApplied(),
        k.onClose());
    } catch (error) {
      u(
        error instanceof Error
          ? error.message
          : "The cells could not be moved.",
      );
    }
  }
  return i.jsx("div", {
    className: "et-backdrop",
    children: i.jsxs("section", {
      ref: o,
      className: "et-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": n,
      tabIndex: -1,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.preventDefault(),
            event.stopPropagation(),
            a.current.onClose()),
          event.key === "Tab")
        ) {
          const I = Array.from(
              o.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ),
            O = I[0],
            j = I.at(-1);
          event.shiftKey && document.activeElement === O
            ? (event.preventDefault(), j == null || j.focus())
            : !event.shiftKey &&
              document.activeElement === j &&
              (event.preventDefault(), O == null || O.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "et-header",
          children: [
            i.jsxs("h2", {
              id: n,
              children: [
                i.jsx(ArrowRight, { size: 19 }),
                i.jsx(LocalizedText, { children: "Move cells" }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": text("Close move preview"),
              onClick: e.onClose,
              lang: language("Close move preview"),
              "data-ui-label": "Close move preview",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "et-body",
          children: [
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "The source stays in place until you apply this move. The preview includes overwritten cells, formula repairs and review metadata.",
              }),
            }),
            i.jsxs("p", {
              children: [
                i.jsxs("strong", {
                  children: [s.current.sourceName, "!", s.current.source.range],
                }),
                " ",
                "→",
                " ",
                i.jsxs("strong", {
                  children: [
                    s.current.targetName,
                    "!",
                    f.plan?.targetRange ?? s.current.range,
                  ],
                }),
              ],
            }),
            f.plan &&
              i.jsxs("section", {
                className: "et-preview",
                "aria-label": text("Move preview"),
                lang: language("Move preview"),
                "data-ui-label": "Move preview",
                children: [
                  i.jsxs("p", {
                    children: [
                      f.plan.count,
                      " ",
                      i.jsx(LocalizedText, { children: "cell changes ·" }),
                      " ",
                      f.plan.overwrittenCells,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "occupied destinations overwritten · ",
                      }),
                      f.plan.referenceChanges,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "dependent references updated · ",
                      }),
                      f.plan.metadataChanges.length,
                      " ",
                      i.jsx(LocalizedText, { children: "metadata changes" }),
                    ],
                  }),
                  i.jsxs("table", {
                    className: "et-preview-table",
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
                                children: "Before",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "After",
                              }),
                            }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: p.slice(v * 15, v * 15 + 15).map((k, F) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("th", { children: k.label }),
                                i.jsx("td", { children: String(k.before) }),
                                i.jsx("td", { children: String(k.after) }),
                              ],
                            },
                            F,
                          ),
                        ),
                      }),
                    ],
                  }),
                  m > 1 &&
                    i.jsxs("div", {
                      className: "et-pagination",
                      children: [
                        i.jsx("button", {
                          type: "button",
                          disabled: !v,
                          onClick: () => c(v - 1),
                          children: i.jsx(LocalizedText, {
                            children: "Previous changes",
                          }),
                        }),
                        i.jsxs("span", {
                          children: [
                            i.jsx(LocalizedText, { children: "Page " }),
                            v + 1,
                            " ",
                            i.jsx(LocalizedText, { children: "of " }),
                            m,
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled: v + 1 >= m,
                          onClick: () => c(v + 1),
                          children: i.jsx(LocalizedText, {
                            children: "Next changes",
                          }),
                        }),
                      ],
                    }),
                  !p.length &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children: "Source and destination already match.",
                      }),
                    }),
                ],
              }),
            (d || h || f.error) &&
              i.jsx("p", {
                className: "et-error",
                role: "alert",
                children: d || (h ? staleMovePreviewMessage : f.error),
              }),
            e.readOnly &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children: "This workbook is read-only.",
                }),
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "et-footer",
          children: [
            i.jsx("button", {
              type: "button",
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              type: "button",
              className: "et-primary",
              disabled: e.readOnly || h || !f.plan || !p.length,
              onClick: onClick,
              children: i.jsx(LocalizedText, { children: "Apply move" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const numberFormatCategoryLabels = {
    general: "General",
    number: "Number",
    currency: "Currency",
    accounting: "Accounting",
    percent: "Percentage",
    date: "Date",
    time: "Time",
    datetime: "Date and time",
    duration: "Elapsed time",
    custom: "Custom",
  },
  localeDisplayNames = {
    "en-US": "English (United States)",
    "en-GB": "English (United Kingdom)",
    "de-DE": "German (Germany)",
    "fr-FR": "French (France)",
    "ja-JP": "Japanese (Japan)",
    "en-IN": "English (India)",
  };
function resolveSelectionNumberFormat(e) {
  const rangeAddress = parseRangeAddress(e.range),
    t = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    n =
      e.active && isCellAddressInRanges(e.ranges ?? [e.range], e.active)
        ? e.active
        : t,
    o = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId)?.cells[n];
  if (o?.numberFormat)
    try {
      return validateNumberFormatSettings(o.numberFormat);
    } catch {}
  const category = o?.format ?? "general";
  return validateNumberFormatSettings({
    category: category,
    locale: "en-US",
    ...(["number", "currency", "percent"].includes(category)
      ? {
          decimals: Math.min(12, summarizeNumberFormat(o?.numFmt) ?? 2),
          useGrouping: o?.numFmt
            ? o.numFmt.includes(",")
            : category !== "percent",
        }
      : {}),
  });
}
function M4e(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const s = ne.useRef({
      book: e.workbook.id,
      revision: e.workbook.revision,
      content: JSON.stringify(e.workbook),
      sheet: e.sheetId,
      selection: JSON.stringify(e.ranges ?? [e.range]),
      active: e.active,
    }),
    [d] = ne.useState(() => resolveSelectionNumberFormat(e)),
    [category, l] = ne.useState(d.category),
    [locale, h] = ne.useState(d.locale),
    [currency, p] = ne.useState(d.currency ?? "USD"),
    [m, v] = ne.useState(String(d.decimals ?? 2)),
    [g, b] = ne.useState(d.useGrouping ?? true),
    [y, w] = ne.useState(String(d.scale ?? 1)),
    [E, A] = ne.useState(d.negative ?? "minus"),
    [k, F] = ne.useState(d.pattern ?? "#,##0.00"),
    [I, O] = ne.useState(""),
    sheet = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    selectionAreas = parseSelectionAreas(e.ranges ?? [e.range]),
    rangeAddress = parseRangeAddress(e.range),
    D =
      e.active && isCellAddressInRanges(selectionAreas, e.active)
        ? e.active
        : `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    cell = sheet?.cells[D],
    _ = countSelectionCells(selectionAreas),
    importedNumberFormatWarning = getImportedNumberFormatWarning(cell),
    T =
      s.current.book !== e.workbook.id ||
      s.current.revision !== e.workbook.revision ||
      s.current.content !== JSON.stringify(e.workbook) ||
      s.current.sheet !== e.sheetId ||
      s.current.selection !== JSON.stringify(e.ranges ?? [e.range]) ||
      s.current.active !== e.active,
    N = e.readOnly || T || !!sheet?.protection?.enabled || !sheet || _ > 500,
    R = ne.useMemo(
      () => calculateWorkbookWithSettings(e.workbook),
      [e.workbook],
    ),
    L = ["number", "currency", "accounting", "percent"].includes(category),
    $ = ["currency", "accounting"].includes(category),
    U = ["date", "time", "datetime", "duration"].includes(category);
  let G,
    H = "",
    Q = null;
  try {
    ((Q = R(e.sheetId, D)),
      (G = validateNumberFormatSettings({
        category: category,
        locale: locale,
        ...(L
          ? {
              decimals: m.trim() === "" ? NaN : Number(m),
              useGrouping: g,
              scale: category === "percent" ? 1 : Number(y),
              negative: category === "accounting" ? "parentheses" : E,
              ...($ ? { currency: currency } : {}),
            }
          : U || category === "custom"
            ? { pattern: k }
            : {}),
      })));
  } catch (error) {
    H =
      error instanceof Error
        ? error.message
        : "The format could not be previewed.";
  }
  const X$1 = [
      {
        name: `${sheet?.name ?? "Worksheet"}!${D}`,
        value: Q,
      },
      ...(U ||
      (category === "custom" && /[ymdhs]/i.test(k.replace(/"[^"]*"/g, "")))
        ? [
            { name: "Date / time example", value: 46273.625 },
            { name: "Elapsed-time example", value: 1.5 },
          ]
        : [
            { name: "Positive example", value: 1234.567 },
            { name: "Negative example", value: -1234.567 },
            { name: "Zero example", value: 0 },
          ]),
    ],
    J = (te) => {
      (l(te),
        O(""),
        te in builtInDateTimeFormats && F(builtInDateTimeFormats[te][0]),
        te === "custom" &&
          F(
            cell?.numberFormat?.category === "custom"
              ? cell.numberFormat.pattern
              : '#,##0.00;(#,##0.00);"-"',
          ));
    };
  return (
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        o.current
          ?.querySelector('[role="option"][aria-selected="true"]')
          ?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []),
    i.jsx("div", {
      className: "nf-backdrop",
      children: i.jsxs("div", {
        className: "nf-dialog",
        role: "dialog",
        "aria-modal": "true",
        "aria-labelledby": `${n}-title`,
        ref: o,
        onKeyDown: (event) => {
          if (
            (event.stopPropagation(),
            event.key === "Escape" && (event.preventDefault(), a.current()),
            event.key === "Tab")
          ) {
            const oe = [
                ...(o.current?.querySelectorAll(
                  'button:not(:disabled):not([tabindex="-1"]),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
                ) ?? []),
              ],
              de = oe[0],
              he = oe.at(-1);
            event.shiftKey && document.activeElement === de
              ? (event.preventDefault(), he == null || he.focus())
              : !event.shiftKey &&
                document.activeElement === he &&
                (event.preventDefault(), de == null || de.focus());
          }
        },
        children: [
          i.jsxs("header", {
            className: "nf-header",
            children: [
              i.jsx("h2", {
                id: `${n}-title`,
                children: i.jsx(LocalizedText, { children: "Format Cells" }),
              }),
              i.jsx("button", {
                "aria-label": text("Close number formats"),
                onClick: e.onClose,
                lang: language("Close number formats"),
                "data-ui-label": "Close number formats",
                children: i.jsx(X, { size: 18 }),
              }),
            ],
          }),
          i.jsxs("div", {
            className: "nf-subtitle",
            children: [
              i.jsx(LocalizedText, { children: "Number · " }),
              _.toLocaleString(),
              " ",
              i.jsx(LocalizedText, { children: "selected cells · " }),
              selectionAreas.length,
              " ",
              selectionAreas.length === 1 ? "area" : "areas",
            ],
          }),
          i.jsxs("div", {
            className: "nf-content",
            children: [
              i.jsx("div", {
                className: "nf-category",
                role: "listbox",
                "aria-label": text("Number format category"),
                lang: language("Number format category"),
                "data-ui-label": "Number format category",
                children: Uv.map((te, ae) =>
                  i.jsx(
                    "button",
                    {
                      role: "option",
                      "aria-selected": category === te,
                      tabIndex: category === te ? 0 : -1,
                      onClick: () => J(te),
                      onKeyDown: (event) => {
                        if (
                          ["ArrowUp", "ArrowDown", "Home", "End"].includes(
                            event.key,
                          )
                        ) {
                          event.preventDefault();
                          const Se =
                            event.key === "Home"
                              ? 0
                              : event.key === "End"
                                ? Uv.length - 1
                                : (ae +
                                    (event.key === "ArrowDown"
                                      ? 1
                                      : Uv.length - 1)) %
                                  Uv.length;
                          (J(Uv[Se]),
                            event.currentTarget.parentElement?.children[
                              Se
                            ]?.focus());
                        }
                      },
                      children: numberFormatCategoryLabels[te],
                    },
                    te,
                  ),
                ),
              }),
              i.jsxs("form", {
                className: "nf-settings",
                onSubmit: (event) => {
                  (event.preventDefault(), O(""));
                  try {
                    if (e.readOnly)
                      throw new Error("This workbook is read-only.");
                    if (T)
                      throw new Error(
                        "The workbook or selection changed. Close and reopen Format Cells.",
                      );
                    if (!G)
                      throw new Error(H || "Choose a valid number format.");
                    const ae = applyNumberFormatTransaction(
                      e.workbook,
                      e.sheetId,
                      selectionAreas,
                      G,
                      {
                        readOnly: e.readOnly,
                      },
                    );
                    (ae !== e.workbook &&
                      e.onWorkbookChange(ae, "Format numbers"),
                      e.onClose());
                  } catch (error) {
                    O(
                      error instanceof Error
                        ? error.message
                        : "The format could not be applied.",
                    );
                  }
                },
                children: [
                  T &&
                    i.jsx("p", {
                      className: "nf-error",
                      role: "alert",
                      children: i.jsx(LocalizedText, {
                        children:
                          "The workbook or selection changed. Close and reopen Format Cells.",
                      }),
                    }),
                  e.readOnly &&
                    i.jsx("p", {
                      className: "nf-notice",
                      children: i.jsx(LocalizedText, {
                        children:
                          "This workbook is read-only. You can inspect format previews.",
                      }),
                    }),
                  sheet?.protection?.enabled &&
                    i.jsx("p", {
                      className: "nf-notice",
                      children: i.jsx(LocalizedText, {
                        children:
                          "This worksheet is protected. Unprotect it before changing cell formats.",
                      }),
                    }),
                  _ > 500 &&
                    i.jsx("p", {
                      className: "nf-error",
                      role: "alert",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Select at most 500 unique cells to apply number formats.",
                      }),
                    }),
                  importedNumberFormatWarning &&
                    i.jsx("p", {
                      className: "nf-notice",
                      children: importedNumberFormatWarning,
                    }),
                  i.jsxs("section", {
                    className: "nf-preview",
                    "aria-label": text("Number format preview"),
                    lang: language("Number format preview"),
                    "data-ui-label": "Number format preview",
                    children: [
                      i.jsx("h3", {
                        children: i.jsx(LocalizedText, { children: "Sample" }),
                      }),
                      i.jsxs("table", {
                        children: [
                          i.jsx("thead", {
                            children: i.jsxs("tr", {
                              children: [
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Example",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Stored value",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Display",
                                  }),
                                }),
                              ],
                            }),
                          }),
                          i.jsx("tbody", {
                            children: X$1.map((te) => {
                              const ae = G
                                ? formatCellValueWithNumberFormat(
                                    te.value,
                                    G,
                                    e.workbook.datePolicy?.dateSystem,
                                  )
                                : undefined;
                              return i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("th", {
                                      scope: "row",
                                      children: te.name,
                                    }),
                                    i.jsx("td", {
                                      children:
                                        te.value === null
                                          ? "Blank"
                                          : String(te.value),
                                    }),
                                    i.jsx("td", {
                                      children: ae?.accounting
                                        ? i.jsxs("span", {
                                            className: `nf-accounting nf-symbol-${ae.accounting.symbolPosition}`,
                                            children: [
                                              i.jsx("span", {
                                                children: ae.accounting.symbol,
                                              }),
                                              i.jsx("span", {
                                                children: ae.accounting.amount,
                                              }),
                                            ],
                                          })
                                        : (ae?.text ?? "Unsupported pattern"),
                                    }),
                                  ],
                                },
                                te.name,
                              );
                            }),
                          }),
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Current display:",
                          }),
                          " ",
                          i.jsx("strong", {
                            children:
                              formatNumberWithPattern(
                                Q,
                                cell,
                                e.workbook.datePolicy?.dateSystem,
                              ) || "Blank",
                          }),
                          i.jsx(LocalizedText, {
                            children:
                              ". Formatting changes appearance only; stored values and formulas stay unchanged.",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Locale" }),
                      i.jsx("select", {
                        "aria-label": text("Number format locale"),
                        value: locale,
                        onChange: (event) => h(event.target.value),
                        lang: language("Number format locale"),
                        "data-ui-label": "Number format locale",
                        children: supportedNumberLocales.map(
                          (supportedNumberLocale) =>
                            i.jsx(
                              "option",
                              {
                                value: supportedNumberLocale,
                                children:
                                  localeDisplayNames[supportedNumberLocale],
                              },
                              supportedNumberLocale,
                            ),
                        ),
                      }),
                    ],
                  }),
                  L &&
                    i.jsxs("div", {
                      className: "nf-fields",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Decimal places",
                            }),
                            i.jsx(LocalizedNumberField, {
                              "aria-label": text("Decimal places"),
                              min: 0,
                              max: 12,
                              step: 1,
                              value: m,
                              onValueChange: (te) => v(te),
                              lang: language("Decimal places"),
                              "data-ui-label": "Decimal places",
                            }),
                          ],
                        }),
                        $ &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, { children: "Currency" }),
                              i.jsx("select", {
                                "aria-label": text("Currency"),
                                value: currency,
                                onChange: (event) => {
                                  p(event.target.value);
                                },
                                lang: language("Currency"),
                                "data-ui-label": "Currency",
                                children: supportedCurrencyCodes.map(
                                  (supportedCurrencyCode) =>
                                    i.jsx(
                                      "option",
                                      {
                                        value: supportedCurrencyCode,
                                        children: supportedCurrencyCode,
                                      },
                                      supportedCurrencyCode,
                                    ),
                                ),
                              }),
                            ],
                          }),
                        category !== "percent" &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Display scale",
                              }),
                              i.jsxs("select", {
                                "aria-label": text("Display scale"),
                                value: y,
                                onChange: (event) => w(event.target.value),
                                lang: language("Display scale"),
                                "data-ui-label": "Display scale",
                                children: [
                                  i.jsx("option", {
                                    value: "1",
                                    children: "Units",
                                  }),
                                  i.jsx("option", {
                                    value: "1000",
                                    children: "Thousands (K)",
                                  }),
                                  i.jsx("option", {
                                    value: "1000000",
                                    children: "Millions (M)",
                                  }),
                                ],
                              }),
                            ],
                          }),
                        category !== "accounting" &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Negative numbers",
                              }),
                              i.jsxs("select", {
                                "aria-label": text("Negative numbers"),
                                value: E,
                                onChange: (event) => A(event.target.value),
                                lang: language("Negative numbers"),
                                "data-ui-label": "Negative numbers",
                                children: [
                                  i.jsx("option", {
                                    value: "minus",
                                    children: "Minus sign",
                                  }),
                                  i.jsx("option", {
                                    value: "parentheses",
                                    children: "Parentheses",
                                  }),
                                ],
                              }),
                            ],
                          }),
                        i.jsxs("label", {
                          className: "nf-check",
                          children: [
                            i.jsx("input", {
                              type: "checkbox",
                              "aria-label": text("Use thousands separator"),
                              checked: g,
                              onChange: (event) => b(event.target.checked),
                              lang: language("Use thousands separator"),
                              "data-ui-label": "Use thousands separator",
                            }),
                            i.jsx(LocalizedText, {
                              children: "Use thousands separator",
                            }),
                          ],
                        }),
                      ],
                    }),
                  U &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Type" }),
                        i.jsxs("select", {
                          "aria-label": text("Date or time format"),
                          value: k,
                          onChange: (event) => F(event.target.value),
                          lang: language("Date or time format"),
                          "data-ui-label": "Date or time format",
                          children: [
                            !builtInDateTimeFormats[category].some(
                              (te) => te === k,
                            ) && i.jsx("option", { value: k, children: k }),
                            builtInDateTimeFormats[category].map((te) =>
                              i.jsx("option", { value: te, children: te }, te),
                            ),
                          ],
                        }),
                      ],
                    }),
                  category === "custom" &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Type" }),
                            i.jsx("input", {
                              "aria-label": text("Custom number format"),
                              value: k,
                              maxLength: 200,
                              onChange: (event) => F(event.target.value),
                              lang: language("Custom number format"),
                              "data-ui-label": "Custom number format",
                            }),
                          ],
                        }),
                        i.jsx("p", {
                          className: "nf-help",
                          children: i.jsx(LocalizedText, {
                            children:
                              "Use 0/# digits, decimals, grouping, percent, quoted labels and positive;negative;zero;text sections. Date/time fields and [h], [m], [s] elapsed time are supported. Colors, conditions, fractions, exponents and spacing-fill codes require other tools or remain unsupported.",
                          }),
                        }),
                      ],
                    }),
                  category === "accounting" &&
                    i.jsx("p", {
                      className: "nf-help",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Accounting separates the currency symbol and amount, uses a dash for zero and parentheses for negative values.",
                      }),
                    }),
                  U &&
                    i.jsx("p", {
                      className: "nf-help",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Dates use Excel's 1900 serial system. Pattern order is explicit; the locale changes month names and fractional separators. Elapsed time can exceed 24 hours. Negative calendar dates/times are outside this subset.",
                      }),
                    }),
                  L &&
                    y !== "1" &&
                    category !== "percent" &&
                    i.jsxs("p", {
                      className: "nf-notice",
                      children: [
                        i.jsx(LocalizedText, {
                          children: "Displayed amounts are divided by",
                        }),
                        " ",
                        Number(y).toLocaleString(),
                        " ",
                        i.jsx(LocalizedText, { children: "and labelled" }),
                        " ",
                        y === "1000" ? "K" : "M",
                        i.jsx(LocalizedText, {
                          children:
                            ". Calculations continue to use the full stored amounts.",
                        }),
                      ],
                    }),
                  G &&
                    i.jsxs("details", {
                      className: "nf-exchange",
                      children: [
                        i.jsx("summary", {
                          children: i.jsx(LocalizedText, {
                            children: "Saved format and Excel exchange",
                          }),
                        }),
                        i.jsx("code", { children: buildNumberFormatCode(G) }),
                        i.jsx("p", {
                          children: i.jsx(LocalizedText, {
                            children:
                              "Native JSON retains these settings. Excel exchange retains the generated number mask; locale, accounting alignment and reopening in this editor can differ. Keep native JSON for the complete format settings.",
                          }),
                        }),
                      ],
                    }),
                  (I || H) &&
                    i.jsx("p", {
                      className: "nf-error",
                      role: "alert",
                      children: I || H,
                    }),
                  i.jsxs("footer", {
                    className: "nf-actions",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        onClick: e.onClose,
                        children: i.jsx(LocalizedText, { children: "Cancel" }),
                      }),
                      i.jsx("button", {
                        className: "nf-primary",
                        disabled: N || !G,
                        children: i.jsx(LocalizedText, {
                          children: "Apply format",
                        }),
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    })
  );
}
const wS = 20,
  staleOutlineMessage =
    "The workbook or selection changed. Close this dialog and reopen it to review the current outline.",
  formatRowColumnSpan = (e, r, t) =>
    e === "row"
      ? "Rows " + (r + 1) + (r === t ? "" : "–" + (t + 1))
      : "Columns " +
        columnIndexToLetters(r) +
        (r === t ? "" : "–" + columnIndexToLetters(t));

export { staleOutlineMessage, MoveCellsDialog, M4e, formatRowColumnSpan, wS };
