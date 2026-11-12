import { isPlainObjectValue } from "./01-cell-styles-and-themes.js";
import { isFiniteNumber } from "./04-pivot-tables-and-pictures.js";
import {
  workbookLimits,
  maxSelectionAreas,
} from "../02-text-and-financial-functions.js";
import {
  hasLinkedOrBoundCells,
  isoDateTimeSchemaFactory,
} from "../01-schemas-date-functions-and-validation.js";
import {
  string,
  number,
  boolean,
  array,
  strictObject,
  literal,
  enum as _enum,
  unknown,
} from "zod";
import {
  validateSplitPaneState,
  normalizeHiddenIntervals,
  conditionalRuleKinds,
  readHiddenAxisIntervals,
  isNonArrayObject,
  hiddenKeyForAxis,
  columnWidthAt,
  sortConditionalRulesByPriority,
  mapIndexAfterMove,
  reviewThreadLimits,
  normalizeReviewTimestamp,
  axisCountOf,
  addOutlineGroup,
  validateOutlineGroups,
  remapOutlineGroupsForEdit,
  isErrorValueText,
  assertBoundedString,
  assertPlainReviewRecord,
  unsafeControlCharPattern,
  assertWorksheetAxis,
  validationLimits,
  validateCellBorder,
  cn,
  resolveReviewTimestamp,
  rowHeightAt,
  throwWithMessage,
  assertSheetUnprotected,
  normalizeBoundedText,
  assertSixDigitHexColor,
  validateSheetDefaultSizes,
  isCellInsideRange,
  cellBorderEdges,
  createAxisVisibilityIndex,
  failWithMessage,
  vx,
  normalizeReviewAuthor,
  worksheetCapacityLimits,
  assertBoundedValidationText,
} from "./01-cell-styles-and-themes.js";
import {
  sumDecimalPartsToNumber,
  exactAverageOfNumbers,
  exactProductOfNumbers,
} from "./02-formula-function-evaluators.js";
import {
  throwXlsxDefinedNameError,
  spreadsheetMlNamespace,
  stripFormulaStringLiterals,
  validateXmlSafeName,
  parseDefinedNameXmlAttributes,
  validateDefinedName,
  parseDefinedNamesXml,
  nP,
  validateNameFormulaText,
  reservedDefinedNames,
} from "./03-defined-names.js";
import {
  containsSubtotalCall,
  inspectPictureDataUrl,
  containsSubtotalOrAggregateCall,
  referenceSensitiveFunctionNames,
  notReferenceFunctionMarker,
  isFormulaCellValue,
} from "./04-pivot-tables-and-pictures.js";
import {
  B2,
  requireWhatIfName,
  HO,
  transformCellValueSchema,
  whatIfLimits,
  expandWhatIfScenarios,
  referenceImpactKinds,
  cellAddressKey,
  validateWhatIfFields,
  isBoundedPlainText,
  validateWhatIfValues,
  requireWhatIfList,
  columnIndexToLetters,
  printRenderLimits,
  shiftFormulaReferences,
} from "./06-data-model-and-analysis.js";

function resolveFilteredRowContext(workbook, r = {}) {
  if (
    !isPlainObjectValue(r) ||
    Object.keys(r).some((u) => u !== "filteredRowsBySheet")
  )
    throw Error("Calculation options support only filteredRowsBySheet.");
  const filteredRowsBySheet = r.filteredRowsBySheet;
  if (
    filteredRowsBySheet !== undefined &&
    !isPlainObjectValue(filteredRowsBySheet)
  )
    throw Error(
      "Filtered rows must be a worksheet-ID map of inclusive row intervals.",
    );
  const n = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet])),
    o = new Map(
      workbook.sheets.map((sheet) => [sheet.name.toLowerCase(), sheet]),
    ),
    a = new Map();
  let s = 0;
  for (const [u, l] of Object.entries(filteredRowsBySheet ?? {})) {
    const c = n.get(u);
    if (!c) throw Error("Filter context refers to an unknown worksheet ID.");
    if (!Array.isArray(l) || (s += l.length) > 1e4)
      throw Error(
        "Filter context supports at most 10,000 intervals across the workbook.",
      );
    a.set(
      u,
      createAxisVisibilityIndex(
        c.rowCount,
        normalizeHiddenIntervals(l, c.rowCount),
      ),
    );
  }
  const d = new Map();
  for (const sheet of workbook.sheets)
    d.set(
      sheet.id,
      createAxisVisibilityIndex(sheet.rowCount, cn(sheet, "row")),
    );
  return (u, l, c) => {
    if (!referenceSensitiveFunctionNames.includes(u))
      return notReferenceFunctionMarker;
    const { issue, isIssue } = c,
      p = (J) => {
        if (isIssue(J)) throw J;
        if (!isPlainObjectValue(J) || !isPlainObjectValue(J.ref))
          throw issue("#VALUE!");
        const ref = J.ref;
        if (ref.sheet !== undefined && typeof ref.sheet != "string")
          throw issue("#REF!");
        const sheet = o.get(
          String(ref.sheet ?? c.currentSheetName ?? "").toLowerCase(),
        );
        if (!sheet) throw issue("#REF!");
        const K = isPlainObjectValue(ref.from) ? ref.from : ref,
          te = isPlainObjectValue(ref.to) ? ref.to : ref,
          ae = [K.row, K.col, te.row, te.col];
        if (ae.some((le) => !Number.isInteger(le))) throw issue("#REF!");
        const [firstRow, firstCol, lastRow, lastCol] = ae;
        if (
          firstRow < 1 ||
          firstCol < 1 ||
          lastRow < firstRow ||
          lastCol < firstCol ||
          lastRow > sheet.rowCount ||
          lastCol > sheet.colCount
        )
          throw issue("#REF!");
        return {
          sheet: sheet,
          firstRow: firstRow,
          lastRow: lastRow,
          firstCol: firstCol,
          lastCol: lastCol,
        };
      };
    if (u === "FORMULATEXT" || u === "ISFORMULA") {
      if (l.length !== 1) return issue("#VALUE!");
      const J = p(l[0]),
        q = (K, te) => {
          const ae = `${columnIndexToLetters(te - 1)}${K}`;
          return (c.readMetadata?.(J.sheet.id, ae), J.sheet.cells[ae]);
        };
      if (u === "FORMULATEXT") {
        const cell = q(J.firstRow, J.firstCol);
        return isFormulaCellValue(cell) && cell.raw.length <= 8192
          ? cell.raw
          : issue("#N/A");
      }
      const W = (J.lastRow - J.firstRow + 1) * (J.lastCol - J.firstCol + 1);
      return W > 1e4
        ? issue("#LIMIT!")
        : W === 1
          ? isFormulaCellValue(q(J.firstRow, J.firstCol))
          : Array.from({ length: J.lastRow - J.firstRow + 1 }, (K, te) =>
              Array.from({ length: J.lastCol - J.firstCol + 1 }, (ae, oe) =>
                isFormulaCellValue(q(J.firstRow + te, J.firstCol + oe)),
              ),
            );
    }
    const m = u === "AGGREGATE";
    if (l.length < (m ? 3 : 2) || l.length > 255) return issue("#VALUE!");
    const v = (J) => {
        const q = c.retrieve(J);
        if (isIssue(q)) throw q;
        if (
          Array.isArray(q) ||
          !["number", "string", "boolean"].includes(typeof q) ||
          (typeof q == "string" && q.trim() === "") ||
          !Number.isFinite(Number(q))
        )
          throw issue("#VALUE!");
        return Number(q);
      },
      g = v(l[0]);
    if (m && !Number.isInteger(g)) return issue("#VALUE!");
    const b = Math.trunc(g),
      y = !m && b > 100 ? b - 100 : b;
    if (y < 1 || y > (m ? 19 : 11)) return issue("#VALUE!");
    const w = m ? (l[1] == null ? 0 : v(l[1])) : 0;
    if (!Number.isInteger(w) || w < 0 || w > 7) return issue("#VALUE!");
    const E = m ? w % 2 === 1 : b > 100,
      A = m && [2, 3, 6, 7].includes(w),
      k = !m || w < 4,
      F = m && y >= 14;
    if (F && (l.length !== 4 || l[3] == null)) return issue("#VALUE!");
    const I = F ? v(l[3]) : 0,
      O = F ? [l[2]] : l.slice(m ? 2 : 1),
      z =
        F && !((J) => isPlainObjectValue(J) && isPlainObjectValue(J.ref))(O[0]),
      V = z ? [] : O.map(p);
    let D = 0;
    for (const J of V) {
      const q = (J.lastRow - J.firstRow + 1) * (J.lastCol - J.firstCol + 1);
      if (((D += q), q > 2e4 || D > 1e5)) return issue("#LIMIT!");
    }
    for (const J of V)
      for (let firstRow = J.firstRow; firstRow <= J.lastRow; firstRow++)
        for (let firstCol = J.firstCol; firstCol <= J.lastCol; firstCol++)
          if (
            c.isActive(
              J.sheet.id,
              `${columnIndexToLetters(firstCol - 1)}${firstRow}`,
            )
          )
            return issue("#CYCLE!");
    let B = 0,
      _ = 0,
      C = 1 / 0,
      T = -1 / 0,
      N = 0,
      R = 0;
    const L = [],
      $ = (J) => {
        if (isIssue(J)) {
          if (["#LIMIT!", "#CYCLE!"].includes(J.error)) return J;
          if (A) return;
          if (y !== 2 && y !== 3) return J;
        }
        if ((J != null && _++, typeof J != "number")) return;
        if (!Number.isFinite(J)) return issue("#NUM!");
        (B++, (C = Math.min(C, J)), (T = Math.max(T, J)));
        const q = J - N;
        ((N += q / B),
          (R += q * (J - N)),
          (y === 1 || y === 6 || y === 9 || y >= 12) && L.push(J));
      };
    if (z) {
      const J = c.retrieve(O[0]),
        q = Array.isArray(J) ? J : [[J]];
      if (!q.length || q.length > 1e4) return issue("#LIMIT!");
      const W = q.every(Array.isArray) ? q : [q],
        K = W[0]?.length ?? 0;
      if (!K || W.some((te) => te.length !== K || te.some(Array.isArray)))
        return issue("#VALUE!");
      if (W.length * K > 1e4) return issue("#LIMIT!");
      for (const te of W)
        for (const ae of te) {
          if (
            ae != null &&
            !isIssue(ae) &&
            !["number", "string", "boolean"].includes(typeof ae)
          )
            return issue("#VALUE!");
          const oe = $(ae);
          if (oe) return oe;
        }
    }
    for (const J of V)
      for (let firstRow = J.firstRow; firstRow <= J.lastRow; firstRow++)
        if (!(
          a.get(J.sheet.id)?.isHidden(firstRow - 1) ||
          (E && d.get(J.sheet.id).isHidden(firstRow - 1))
        ))
          for (let firstCol = J.firstCol; firstCol <= J.lastCol; firstCol++) {
            const K = `${columnIndexToLetters(firstCol - 1)}${firstRow}`;
            c.readMetadata?.(J.sheet.id, K);
            const te = J.sheet.cells[K];
            if (
              k &&
              isFormulaCellValue(te) &&
              (m ? containsSubtotalOrAggregateCall : containsSubtotalCall)(
                c.expandedFormula(te, J.sheet.id),
              )
            )
              continue;
            const ae = $(c.readCell(J.sheet.id, K));
            if (ae) return ae;
          }
    if (y >= 12) {
      if ((L.sort((oe, de) => oe - de), y === 13)) {
        let oe = 1,
          de = 0,
          he = 0,
          Se;
        for (const le of L)
          ((he = le === Se ? he + 1 : 1),
            he > oe && ((oe = he), (de = le)),
            (Se = le));
        return oe > 1 ? de : issue("#N/A");
      }
      if (!L.length) return issue("#NUM!");
      if (y === 12) {
        const oe = Math.floor(L.length / 2);
        return L.length % 2 ? L[oe] : L[oe - 1] / 2 + L[oe] / 2;
      }
      if (y === 14 || y === 15) {
        const oe = I;
        return Number.isInteger(oe) && oe >= 1 && oe <= L.length
          ? L[y === 14 ? L.length - oe : oe - 1]
          : issue("#NUM!");
      }
      const J = y >= 18,
        q = y === 17 || y === 19 ? Math.trunc(I) / 4 : I;
      if (q < 0 || q > 1 || (J && (q === 0 || q === 1))) return issue("#NUM!");
      const W = J ? q * (L.length + 1) - 1 : q * (L.length - 1);
      if (W < 0 || W > L.length - 1) return issue("#NUM!");
      const K = Math.floor(W),
        te = Math.ceil(W),
        ae = W - K;
      return L[K] * (1 - ae) + L[te] * ae;
    }
    if (y === 2) return B;
    if (y === 3) return _;
    if (y === 9) return sumDecimalPartsToNumber(L);
    if (y === 1) return B ? exactAverageOfNumbers(L) : issue("#DIV/0!");
    if (y === 4) return B ? T : 0;
    if (y === 5) return B ? C : 0;
    if (y === 6) return B ? exactProductOfNumbers(L) : 0;
    const U = y === 7 || y === 10;
    if (B < (U ? 2 : 1)) return issue("#DIV/0!");
    const G = Math.max(0, R / (B - (U ? 1 : 0)));
    return y === 7 || y === 8 ? Math.sqrt(G) : G;
  };
}
function formatAxisIntervalRef(e, r) {
  return r === "row"
    ? `${e.start + 1}:${e.end + 1}`
    : `${columnIndexToLetters(e.start)}:${columnIndexToLetters(e.end)}`;
}
function _t(e) {
  const r = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(e);
  if (!r) return throwWithMessage(`Invalid cell address: ${e}`);
  let t = 0;
  for (const o of r[1].toUpperCase()) t = t * 26 + o.charCodeAt(0) - 64;
  const row = Number(r[2]) - 1;
  return row >= workbookLimits.maxRows || t > workbookLimits.maxCols
    ? throwWithMessage(`Cell ${e} exceeds the supported workbook bounds.`)
    : { row: row, col: t - 1 };
}
function buildBorderPresetStyles(
  address,
  sampleGridAddresses,
  preset,
  line = { style: "thin", color: "#000000" },
) {
  const o = _t(address),
    a = validateCellBorder(line),
    s = {};
  ["all", "outline", "inside", "clear"].includes(preset) ||
    failWithMessage("Choose a supported border preset.");
  for (const cellBorderEdge of cellBorderEdges) {
    const u =
        o.row +
        (cellBorderEdge === "top" ? -1 : cellBorderEdge === "bottom" ? 1 : 0),
      l =
        o.col +
        (cellBorderEdge === "left" ? -1 : cellBorderEdge === "right" ? 1 : 0),
      c =
        u >= 0 &&
        u < 1048576 &&
        l >= 0 &&
        l < 16384 &&
        sampleGridAddresses.has(`${columnIndexToLetters(l)}${u + 1}`);
    preset === "clear"
      ? (s[cellBorderEdge] = { style: "none" })
      : (preset === "all" || (preset === "outline" ? !c : c)) &&
        (s[cellBorderEdge] = { ...a });
  }
  return s;
}
function normalizeCellAddress(address) {
  if (typeof address != "string")
    return throwWithMessage("Choose a valid worksheet cell address.");
  const r = _t(address.trim());
  return `${columnIndexToLetters(r.col)}${r.row + 1}`;
}
function validateHyperlinkRecord(hyperlink) {
  const r = assertPlainReviewRecord(hyperlink, "Hyperlink"),
    t =
      r.label === undefined || r.label === ""
        ? undefined
        : normalizeBoundedText(
            r.label,
            "Link display text",
            reviewThreadLimits.labelLength,
          );
  if (r.kind === "internal")
    return {
      kind: "internal",
      sheetId: normalizeBoundedText(r.sheetId, "Link worksheet ID", 100),
      address: normalizeCellAddress(r.address),
      ...(t ? { label: t } : {}),
    };
  if (r.kind !== "external")
    return throwWithMessage("Choose a web, email or worksheet link.");
  const boundedText = normalizeBoundedText(
    r.url,
    "Link address",
    reviewThreadLimits.urlLength,
  );
  let o = boundedText;
  for (let s = 0; s < 4; s++) {
    (unsafeControlCharPattern.test(o) || /\\/.test(o)) &&
      throwWithMessage(
        "Link addresses cannot contain control characters or backslashes.",
      );
    let d = o;
    try {
      d = decodeURIComponent(o);
    } catch {
      throwWithMessage("The link contains invalid encoded characters.");
    }
    if (d === o) break;
    o = d;
  }
  (unsafeControlCharPattern.test(o) ||
    /%(?:0[\da-f]|1[\da-f]|7f|8[\da-f]|9[\da-f])/i.test(o)) &&
    throwWithMessage(
      "Link addresses cannot contain encoded control characters.",
    );
  let a;
  try {
    a = new URL(boundedText);
  } catch {
    return throwWithMessage(
      "Enter a complete https://, http:// or mailto: address.",
    );
  }
  return (
    ["https:", "http:", "mailto:"].includes(a.protocol) ||
      throwWithMessage(
        "Only https://, http:// and mailto: links are supported.",
      ),
    (a.username || a.password) &&
      throwWithMessage("Links cannot contain embedded credentials."),
    a.protocol === "mailto:"
      ? (a.search ||
          a.hash ||
          a.host ||
          !/^[^\s@/?#:,]+@[^\s@/?#:,]+\.[^\s@/?#:,]+$/.test(
            decodeURIComponent(a.pathname),
          )) &&
        throwWithMessage(
          "Use mailto: with one email address and no additional headers.",
        )
      : a.hostname || throwWithMessage("Web links require a hostname."),
    a.href.length > reviewThreadLimits.urlLength &&
      throwWithMessage("The normalized link is too long."),
    { kind: "external", url: a.href, ...(t ? { label: t } : {}) }
  );
}
function normalizeReviewMetadata(e, workbook) {
  const t = assertPlainReviewRecord(e, "Workbook review metadata");
  if (t.reviewThreads !== undefined) {
    (!Array.isArray(t.reviewThreads) ||
      t.reviewThreads.length > reviewThreadLimits.threads) &&
      throwWithMessage(
        `A workbook can contain at most ${reviewThreadLimits.threads} review threads.`,
      );
    const n = new Set(),
      o = new Set();
    let a = 0;
    const s = t.reviewThreads.map((reviewThread) => {
      const u = assertPlainReviewRecord(reviewThread, "Review thread"),
        l = assertPlainReviewRecord(u.anchor, "Review target"),
        boundedText = normalizeBoundedText(u.id, "Review thread ID", 100);
      (n.has(boundedText) &&
        throwWithMessage("Review thread IDs must be unique."),
        n.add(boundedText));
      const reviewTimestamp = normalizeReviewTimestamp(u.createdAt),
        f = normalizeReviewTimestamp(u.updatedAt);
      f < reviewTimestamp &&
        throwWithMessage("Review update time cannot precede creation.");
      const anchor = {
        sheetId: normalizeBoundedText(l.sheetId, "Review worksheet ID", 100),
        sheetName: normalizeBoundedText(
          l.sheetName,
          "Review worksheet name",
          31,
        ),
        address: normalizeCellAddress(l.address),
      };
      (!Array.isArray(u.messages) ||
        !u.messages.length ||
        u.messages.length > reviewThreadLimits.messagesPerThread ||
        (a += u.messages.length) > reviewThreadLimits.totalMessages) &&
        throwWithMessage(
          "Review discussions exceed the per-thread or workbook message limit.",
        );
      let m = reviewTimestamp;
      const messages = u.messages.map((message, y) => {
          const w = assertPlainReviewRecord(message, "Review message"),
            boundedText = normalizeBoundedText(w.id, "Review message ID", 100);
          (o.has(boundedText) &&
            throwWithMessage(
              "Review message IDs must be unique within the workbook.",
            ),
            o.add(boundedText));
          const A = normalizeReviewTimestamp(w.createdAt);
          return (
            (A < m || A > f || (y === 0 && A !== reviewTimestamp)) &&
              throwWithMessage(
                "Review message times must follow their discussion in order.",
              ),
            (m = A),
            {
              id: boundedText,
              author: normalizeReviewAuthor(
                normalizeBoundedText(w.author, "Reviewer name", 100),
              ),
              text: normalizeBoundedText(
                w.text,
                "Comment",
                reviewThreadLimits.textLength,
                true,
              ),
              createdAt: A,
            }
          );
        }),
        g = {
          id: boundedText,
          anchor: anchor,
          messages: messages,
          createdAt: reviewTimestamp,
          updatedAt: f,
        };
      if (u.resolved !== undefined) {
        const b = assertPlainReviewRecord(u.resolved, "Review resolution"),
          reviewTimestamp = normalizeReviewTimestamp(b.createdAt);
        ((reviewTimestamp < messages.at(-1).createdAt || reviewTimestamp > f) &&
          throwWithMessage("Invalid review resolution time."),
          (g.resolved = {
            author: normalizeReviewAuthor(
              normalizeBoundedText(b.author, "Reviewer name", 100),
            ),
            createdAt: reviewTimestamp,
          }));
      }
      if (u.orphaned !== undefined) {
        const b = assertPlainReviewRecord(u.orphaned, "Orphaned review target"),
          y = normalizeReviewTimestamp(b.createdAt);
        ((!["worksheet-deleted", "cell-deleted"].includes(String(b.reason)) ||
          y < reviewTimestamp ||
          y > f) &&
          throwWithMessage("Invalid orphaned review target."),
          (g.orphaned = { reason: b.reason, createdAt: y }));
      }
      return (
        !g.orphaned &&
          !workbook.sheets.some((sheet) => sheet.id === anchor.sheetId) &&
          throwWithMessage(
            "A review thread with a missing worksheet must be retained as orphaned.",
          ),
        g
      );
    });
    s.length ? (workbook.reviewThreads = s) : delete workbook.reviewThreads;
  }
  for (const sheet of workbook.sheets)
    for (const o of hasLinkedOrBoundCells(sheet.cells)
      ? Object.values(sheet.cells)
      : []) {
      if (o.hyperlink === undefined) continue;
      const a = validateHyperlinkRecord(o.hyperlink);
      (a.kind === "internal" &&
        !workbook.sheets.some((sheet) => sheet.id === a.sheetId) &&
        throwWithMessage(
          "An internal hyperlink must target an existing worksheet.",
        ),
        (o.hyperlink = a));
    }
}
function attachReviewThreadsToWorkbook(e, reviewThreads) {
  const workbook = { ...e, reviewThreads: reviewThreads },
    n = {
      ...workbook,
      sheets: workbook.sheets.map((sheet) => ({
        ...sheet,
        cells: Object.fromEntries(
          Object.entries(sheet.cells).map(([a, s]) => [a, { ...s }]),
        ),
      })),
    };
  return (normalizeReviewMetadata(workbook, n), n);
}
function shiftReviewTargetsOnAxisEdit(workbook, sheetId, t, timestamp) {
  const o = t.axis === "row" ? workbookLimits.maxRows : workbookLimits.maxCols;
  (!["row", "column"].includes(t.axis) ||
    !Number.isInteger(t.index) ||
    t.index < 0 ||
    !Number.isInteger(t.count) ||
    t.count < 1 ||
    t.index + t.count > o) &&
    throwWithMessage("Choose a valid row or column change for review targets.");
  const a = (c) => {
    const h = _t(c),
      f = t.axis === "row" ? h.row : h.col;
    if (t.moveBefore === undefined && f < t.index) return c;
    if (t.moveBefore === undefined && t.remove && f < t.index + t.count)
      return null;
    const p =
      t.moveBefore !== undefined
        ? mapIndexAfterMove(f, t)
        : f + (t.remove ? -t.count : t.count);
    return (
      p >= o &&
        throwWithMessage(
          "Insertion would discard a review target or hyperlink at the worksheet boundary.",
        ),
      t.axis === "row"
        ? `${columnIndexToLetters(h.col)}${p + 1}`
        : `${columnIndexToLetters(p)}${h.row + 1}`
    );
  };
  let s = false;
  const reviewThreads = workbook.reviewThreads?.map((reviewThread) => {
      if (reviewThread.orphaned || reviewThread.anchor.sheetId !== sheetId)
        return reviewThread;
      const address = a(reviewThread.anchor.address);
      if (address === reviewThread.anchor.address) return reviewThread;
      if (((s = true), address))
        return {
          ...reviewThread,
          anchor: { ...reviewThread.anchor, address: address },
        };
      const reviewTimestamp = resolveReviewTimestamp(
        { timestamp: timestamp },
        reviewThread.updatedAt,
      );
      return {
        ...reviewThread,
        updatedAt: reviewTimestamp,
        orphaned: { reason: "cell-deleted", createdAt: reviewTimestamp },
      };
    }),
    sheets = workbook.sheets.map((sheet) => {
      let cells;
      for (const [f, p] of Object.entries(sheet.cells)) {
        const hyperlink = p.hyperlink;
        if (hyperlink?.kind !== "internal" || hyperlink.sheetId !== sheetId)
          continue;
        const address = a(hyperlink.address);
        address !== hyperlink.address &&
          ((cells ??= { ...sheet.cells }),
          (cells[f] = { ...p }),
          (s = true),
          address
            ? (cells[f].hyperlink = { ...hyperlink, address: address })
            : delete cells[f].hyperlink);
      }
      return cells ? { ...sheet, cells: cells } : sheet;
    });
  return s
    ? {
        ...workbook,
        sheets: sheets,
        ...(reviewThreads ? { reviewThreads: reviewThreads } : {}),
      }
    : workbook;
}
function shiftFormulaWatchesOnAxisEdit(e, sheetId, t) {
  return e.formulaWatches?.length
    ? {
        ...e,
        formulaWatches: e.formulaWatches.map((formulaWatch) => {
          if (formulaWatch.orphaned || formulaWatch.anchor.sheetId !== sheetId)
            return formulaWatch;
          const a = _t(formulaWatch.anchor.address),
            s = t.axis === "row" ? a.row : a.col;
          if (
            t.moveBefore === undefined &&
            t.remove &&
            s >= t.index &&
            s < t.index + t.count
          )
            return { ...formulaWatch, orphaned: "deleted-cells" };
          const d =
            t.moveBefore !== undefined
              ? mapIndexAfterMove(s, t)
              : s >= t.index + (t.remove ? t.count : 0)
                ? s + (t.remove ? -t.count : t.count)
                : s;
          if (
            d < 0 ||
            d >=
              (t.axis === "row"
                ? workbookLimits.maxRows
                : workbookLimits.maxCols)
          )
            return { ...formulaWatch, orphaned: "out-of-bounds" };
          const address =
            t.axis === "row"
              ? `${columnIndexToLetters(a.col)}${d + 1}`
              : `${columnIndexToLetters(d)}${a.row + 1}`;
          return address === formulaWatch.anchor.address
            ? formulaWatch
            : {
                ...formulaWatch,
                anchor: { ...formulaWatch.anchor, address: address },
              };
        }),
      }
    : e;
}
function createAnchoredDrawingObject(e, range) {
  const t = _t(range.split(":")[0]);
  return {
    id: crypto.randomUUID(),
    kind: e,
    anchor: `${columnIndexToLetters(t.col)}${t.row + 1}`,
    offsetX: 12,
    offsetY: 12,
    width: e === "chart" ? 480 : 240,
    height: e === "chart" ? 300 : 140,
    label: e === "picture" ? "Picture" : e === "chart" ? "Chart" : "Text",
    fill: "#E2EFDA",
    stroke: "#548235",
    ...(e === "shape" ? { shape: "rectangle" } : {}),
  };
}
function buildDrawingExportNotes(e, r, sheet) {
  const n = [];
  for (const s of sheet.objects ?? []) {
    if (s.kind !== "picture") {
      n.push(
        `${s.label || s.kind}: native chart/shape export is not supported; retain the JSON workbook.`,
      );
      continue;
    }
    const { extension } = inspectPictureDataUrl(s.dataURL);
    s.evidence &&
      n.push(
        `${s.label || "Picture"}: XLSX retains the image bytes and placement, but original filename, source reference and evidence provenance require native JSON; a later XLSX import records only observed media.`,
      );
    const u = _t(s.anchor),
      l = e.addImage({ base64: s.dataURL, extension });
    r.addImage(l, {
      tl: {
        col: u.col + s.offsetX / (sheet.columnWidths?.[u.col] ?? 84),
        row: u.row + s.offsetY / (sheet.rowHeights?.[u.row] ?? 21),
      },
      ext: { width: s.width, height: s.height },
      editAs: "oneCell",
    });
  }
  return n;
}
function normalizeFormulaWatchAddress(address) {
  if (
    typeof address != "string" ||
    !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}$/.test(address)
  )
    throw Error("A formula watch must identify one valid Excel cell address.");
  const r = _t(address.replace(/\$/g, "").toUpperCase());
  if (r.row >= workbookLimits.maxRows || r.col >= workbookLimits.maxCols)
    throw Error("A formula watch address is outside Excel's cell bounds.");
  return `${columnIndexToLetters(r.col)}${r.row + 1}`;
}
function validateFormulaWatches(formulaWatches, workbook) {
  if (formulaWatches === undefined) return [];
  if (!Array.isArray(formulaWatches) || formulaWatches.length > B2)
    throw Error("A workbook may contain at most 100 formula watches.");
  const t = new Set(),
    n = new Set();
  return formulaWatches.map((formulaWatch) => {
    if (
      !isPlainObjectValue(formulaWatch) ||
      !HO(formulaWatch, ["id", "anchor", "orphaned"]) ||
      !isBoundedPlainText(formulaWatch.id, 100) ||
      !isPlainObjectValue(formulaWatch.anchor) ||
      !HO(formulaWatch.anchor, ["sheetId", "sheetName", "address"]) ||
      !isBoundedPlainText(formulaWatch.anchor.sheetId, 100) ||
      !isBoundedPlainText(formulaWatch.anchor.sheetName, 31)
    )
      throw Error("Invalid formula watch metadata.");
    if (t.has(formulaWatch.id))
      throw Error("Formula watch IDs must be unique.");
    if (
      (t.add(formulaWatch.id),
      formulaWatch.orphaned !== undefined &&
        !referenceImpactKinds.includes(formulaWatch.orphaned))
    )
      throw Error("Invalid formula watch orphan reason.");
    const anchor = formulaWatch.anchor,
      formulaWatchAddress = normalizeFormulaWatchAddress(anchor.address),
      sheet = workbook.sheets.find((sheet) => sheet.id === anchor.sheetId),
      orphaned = formulaWatch.orphaned ?? (sheet ? undefined : "missing-sheet"),
      l = JSON.stringify([formulaWatch.anchor.sheetId, formulaWatchAddress]);
    if (!orphaned && n.has(l))
      throw Error("Active formula watches must not repeat the same cell.");
    return (
      orphaned || n.add(l),
      {
        id: formulaWatch.id,
        anchor: {
          sheetId: formulaWatch.anchor.sheetId,
          sheetName:
            !orphaned && sheet ? sheet.name : formulaWatch.anchor.sheetName,
          address: formulaWatchAddress,
        },
        ...(orphaned ? { orphaned: orphaned } : {}),
      }
    );
  });
}
function normalizeWhatIfCell(e) {
  const r = validateWhatIfFields(e, ["sheetId", "address"]),
    sheetId = requireWhatIfName(r.sheetId);
  if (typeof r.address != "string")
    throw Error("A what-if cell needs an Excel address.");
  const n = _t(r.address);
  return {
    sheetId: sheetId,
    address: `${columnIndexToLetters(n.col)}${n.row + 1}`,
  };
}
function normalizeWhatIfCellReferences(e, r) {
  const t = requireWhatIfList(e, r).map(normalizeWhatIfCell);
  if (new Set(t.map(cellAddressKey)).size !== t.length)
    throw Error("What-if cell references must not repeat.");
  return t;
}
function validateWhatIfDefinition(request) {
  const r = validateWhatIfFields(request, [
      "name",
      "kind",
      "inputs",
      "outputs",
      "scenarios",
      "orientation",
      "values",
      "rowValues",
      "columnValues",
    ]),
    t = {
      name: requireWhatIfName(r.name),
      inputs: normalizeWhatIfCellReferences(r.inputs, whatIfLimits.inputs),
      outputs: normalizeWhatIfCellReferences(r.outputs, whatIfLimits.outputs),
    };
  if (
    t.inputs.some((input) =>
      t.outputs.some(
        (output) => cellAddressKey(input) === cellAddressKey(output),
      ),
    )
  )
    throw Error("Changing inputs and formula results must be different cells.");
  const n = ["name", "kind", "inputs", "outputs"];
  if (r.kind === "scenarios") {
    validateWhatIfFields(r, [...n, "scenarios"]);
    const scenarios = requireWhatIfList(
      r.scenarios,
      whatIfLimits.scenarios,
    ).map((a) => {
      const s = validateWhatIfFields(a, ["name", "values"]),
        values = validateWhatIfValues(s.values, whatIfLimits.inputs);
      if (values.length !== t.inputs.length)
        throw Error(
          "Each scenario needs exactly one value per changing input.",
        );
      return { name: requireWhatIfName(s.name), values: values };
    });
    if (
      new Set(scenarios.map((scenario) => scenario.name.toLowerCase())).size !==
      scenarios.length
    )
      throw Error("Scenario names must be unique, ignoring case.");
    return { ...t, kind: r.kind, scenarios: scenarios };
  }
  if (r.kind === "one-variable") {
    if (
      (validateWhatIfFields(r, [...n, "orientation", "values"]),
      t.inputs.length !== 1 || !["row", "column"].includes(r.orientation))
    )
      throw Error(
        "A one-variable table needs one input and a row or column orientation.",
      );
    return {
      ...t,
      kind: r.kind,
      orientation: r.orientation,
      values: validateWhatIfValues(r.values, whatIfLimits.axisValues),
    };
  }
  if (r.kind === "two-variable") {
    if (
      (validateWhatIfFields(r, [...n, "rowValues", "columnValues"]),
      t.inputs.length !== 2 || t.outputs.length !== 1)
    )
      throw Error(
        "A two-variable table needs two different inputs and exactly one formula result.",
      );
    const rowValues = validateWhatIfValues(
        r.rowValues,
        whatIfLimits.axisValues,
      ),
      columnValues = validateWhatIfValues(
        r.columnValues,
        whatIfLimits.axisValues,
      );
    if (rowValues.length * columnValues.length > whatIfLimits.trials)
      throw Error("A sensitivity table supports at most 400 combinations.");
    return {
      ...t,
      kind: r.kind,
      rowValues: rowValues,
      columnValues: columnValues,
    };
  }
  throw Error(
    "Choose scenarios, a one-variable table or a two-variable table.",
  );
}
function validateSavedWhatIfAnalyses(whatIfAnalyses) {
  const r = new Set();
  return requireWhatIfList(whatIfAnalyses, whatIfLimits.savedAnalyses, 0).map(
    (t) => {
      const n = validateWhatIfFields(t, [
          "id",
          "reportSheetId",
          "request",
          "baseline",
          "trials",
          "assumptions",
          "sourceRevision",
          "dependencyCells",
          "workUnits",
        ]),
        o = requireWhatIfName(n.id),
        reportSheetId = requireWhatIfName(n.reportSheetId),
        request = validateWhatIfDefinition(n.request);
      if (r.has(o)) throw Error("Saved what-if IDs must be unique.");
      r.add(o);
      const d = (p) => {
          const m = validateWhatIfFields(p, [
              "name",
              "inputs",
              "outputs",
              "error",
            ]),
            v = requireWhatIfName(m.name, 200),
            inputs = validateWhatIfValues(m.inputs, whatIfLimits.inputs),
            outputs =
              m.outputs === null
                ? null
                : validateWhatIfValues(m.outputs, whatIfLimits.outputs),
            y = m.error === null ? null : requireWhatIfName(m.error, 2e3);
          if (
            inputs.length !== request.inputs.length ||
            (outputs && outputs.length !== request.outputs.length) ||
            !!outputs == !!y
          )
            throw Error("Invalid saved what-if result dimensions or outcome.");
          return { name: v, inputs: inputs, outputs: outputs, error: y };
        },
        baseline = d(n.baseline),
        trials = requireWhatIfList(n.trials, whatIfLimits.trials).map(d),
        c = expandWhatIfScenarios(request);
      if (
        trials.length !== c.length ||
        trials.some(
          (trial, m) =>
            trial.name !== c[m].name ||
            JSON.stringify(trial.inputs) !== JSON.stringify(c[m].inputs),
        )
      )
        throw Error(
          "Saved what-if results must match their named assumptions.",
        );
      if (baseline.name !== "Baseline")
        throw Error("Saved what-if baseline is missing.");
      const h = new Set(),
        assumptions = requireWhatIfList(
          n.assumptions,
          whatIfLimits.dependencyCells,
        ).map((p) => {
          const cell = validateWhatIfFields(p, [
              "sheetId",
              "sheetName",
              "address",
              "raw",
              "literal",
            ]),
            whatIfCell = normalizeWhatIfCell({
              sheetId: cell.sheetId,
              address: cell.address,
            });
          if (h.has(cellAddressKey(whatIfCell)))
            throw Error("Saved assumptions must not repeat a source cell.");
          if (
            (h.add(cellAddressKey(whatIfCell)),
            !(
              cell.raw === null ||
              typeof cell.raw == "boolean" ||
              isFiniteNumber(cell.raw) ||
              (typeof cell.raw == "string" && cell.raw.length <= 32767)
            ) ||
              (cell.literal !== undefined && typeof cell.literal != "boolean"))
          )
            throw Error("Invalid saved what-if assumption.");
          return {
            ...whatIfCell,
            sheetName: requireWhatIfName(cell.sheetName, 31),
            raw: cell.raw,
            ...(cell.literal === undefined ? {} : { literal: cell.literal }),
          };
        });
      if (
        [...request.inputs, ...request.outputs].some(
          (p) => !h.has(cellAddressKey(p)),
        ) ||
        request.inputs.some(
          (input, m) =>
            assumptions.find(
              (assumption) =>
                cellAddressKey(assumption) === cellAddressKey(input),
            )?.raw !== baseline.inputs[m],
        )
      )
        throw Error(
          "Saved baseline inputs and dependency assumptions do not reconcile.",
        );
      for (const [p, m] of [
        ["sourceRevision", Number.MAX_SAFE_INTEGER - 1],
        ["dependencyCells", whatIfLimits.dependencyCells],
        ["workUnits", whatIfLimits.workUnits],
      ])
        if (!Number.isSafeInteger(n[p]) || n[p] < 0 || n[p] > m)
          throw Error("Invalid what-if evidence counts.");
      if (
        n.dependencyCells !== assumptions.length ||
        n.workUnits !== assumptions.length * (trials.length + 1)
      )
        throw Error("What-if work counts do not reconcile.");
      return {
        id: o,
        reportSheetId: reportSheetId,
        request: request,
        baseline: baseline,
        trials: trials,
        assumptions: assumptions,
        sourceRevision: n.sourceRevision,
        dependencyCells: assumptions.length,
        workUnits: n.workUnits,
      };
    },
  );
}
function applyChartDrawingLayout(sheet, object) {
  if (object.kind === "chart" && object.chart === "sparkline") {
    const t = _t(object.anchor);
    return {
      offsetX: 2,
      offsetY: 2,
      width: columnWidthAt(sheet, t.col) - 4,
      height: rowHeightAt(sheet, t.row) - 4,
    };
  }
  return {
    offsetX: object.offsetX,
    offsetY: object.offsetY,
    width: object.width,
    height: object.height,
  };
}
function parsePrintTitles(e, r, sheet) {
  if (e === undefined || e === "") return;
  if (typeof e != "string")
    throw Error("Print titles must be a single row or column interval.");
  const n = (
    r === "row"
      ? /^\$?([1-9]\d*):\$?([1-9]\d*)$/
      : /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/
  ).exec(e.trim());
  if (!n)
    throw Error(
      r === "row"
        ? "Repeat rows must use a range such as 1:2."
        : "Repeat columns must use a range such as A:B.",
    );
  const start = r === "row" ? Number(n[1]) - 1 : _t(n[1] + "1").col,
    end = r === "row" ? Number(n[2]) - 1 : _t(n[2] + "1").col,
    s = r === "row" ? sheet.rowCount : sheet.colCount,
    d =
      r === "row"
        ? printRenderLimits.titleRows
        : printRenderLimits.titleColumns;
  if (
    !Number.isSafeInteger(start) ||
    !Number.isSafeInteger(end) ||
    start < 0 ||
    end < start ||
    end >= s ||
    end - start + 1 > d
  )
    throw Error(
      `Print titles support at most ${d} ${r === "row" ? "rows" : "columns"} within the worksheet.`,
    );
  return { start: start, end: end };
}
function shiftPrintSettingsOnAxisEdit(
  pageLayout,
  sheet,
  axis,
  n,
  count,
  remove,
) {
  const s = axis === "row" ? "repeatRows" : "repeatColumns",
    d = axis === "row" ? "rowBreaks" : "columnBreaks",
    u = { ...pageLayout },
    printTitles = parsePrintTitles(pageLayout[s], axis, sheet);
  if (printTitles) {
    let { start, end } = printTitles;
    if (!remove) (n <= start && (start += count), n <= end && (end += count));
    else {
      const m = n + count - 1,
        v = [
          start < n ? { start, end: Math.min(end, n - 1) } : undefined,
          end > m
            ? { start: Math.max(start, m + 1) - count, end: end - count }
            : undefined,
        ].filter(Boolean);
      if (!v.length)
        throw Error(
          "This edit removes every repeated print title. Clear or change print titles first.",
        );
      ((start = v[0].start), (end = v.at(-1).end));
    }
    ((u[s] = formatAxisIntervalRef({ start, end }, axis)),
      parsePrintTitles(u[s], axis, sheet));
  }
  const c = axis === "row" ? sheet.rowCount : sheet.colCount,
    h = (pageLayout[d] ?? [])
      .map((f) =>
        remove
          ? f < n
            ? f
            : f >= n + count
              ? f - count
              : n
          : f >= n
            ? f + count
            : f,
      )
      .filter((f) => f > 0);
  if (h.some((f) => f >= c))
    throw Error(
      "This edit moves a manual print break beyond the worksheet. Remove the break first.",
    );
  return (
    h.length ? (u[d] = [...new Set(h)].sort((f, p) => f - p)) : delete u[d],
    u
  );
}
function isCanonicalCellAddress(e) {
  if (typeof e != "string") return false;
  try {
    const r = _t(e);
    return `${columnIndexToLetters(r.col)}${r.row + 1}` === e;
  } catch {
    return false;
  }
}
const evidenceAddressSchema = string().refine((e) => {
    try {
      const r = _t(e);
      return `${columnIndexToLetters(r.col)}${r.row + 1}` === e;
    } catch {
      return false;
    }
  }, "Invalid evidence address."),
  evidenceSourceOriginSchema = strictObject({
    id: string().min(1).max(100),
    sourceId: string().min(1).max(100),
    sourceName: string().min(1).max(255),
    sourceSheet: string().min(1).max(31),
    originalRow: number().int().min(1).max(1048576),
    kind: _enum(["file", "snapshot"]),
  }),
  evidenceCellRecordSchema = strictObject({
    address: evidenceAddressSchema,
    original: unknown().optional(),
    origin: evidenceSourceOriginSchema.optional(),
  }),
  evidenceFieldValueSchema = strictObject({
    address: evidenceAddressSchema,
    row: number().int().min(1).max(1048576),
    field: string().min(1).max(150),
    value: transformCellValueSchema,
    type: _enum(["text", "number", "boolean", "date", "blank"]),
    sources: array(evidenceAddressSchema).min(1).max(500),
    exception: boolean(),
  }),
  evidenceTransformStepSchema = strictObject({
    step: number().int().min(1).max(12),
    row: number().int().min(1).max(1048576),
    field: string().min(1).max(150),
    raw: transformCellValueSchema,
    sources: array(evidenceAddressSchema).min(1).max(500),
    reason: string().min(1).max(1e3),
  }),
  analysisOutputEvidenceSchema = strictObject({
    id: string().min(1).max(100),
    workbookId: string().min(1).max(100),
    sourceRevision: number().int().nonnegative(),
    sourceSheetId: string().min(1).max(100),
    sourceSheetName: string().min(1).max(31),
    outputSheetId: string().min(1).max(100),
    outputSheetName: string().min(1).max(31),
    createdAt: isoDateTimeSchemaFactory(),
    semantics: literal("reviewed-fields-v1"),
    request: unknown(),
    sources: array(evidenceCellRecordSchema).min(1).max(500),
    outputs: array(evidenceFieldValueSchema).min(1).max(5e3),
    exceptions: array(evidenceTransformStepSchema).max(5e3),
  });
function parseRangeAddress(e) {
  const r = e.trim().split(":");
  if (r.length > 2) return throwWithMessage(`Invalid range: ${e}`);
  const t = _t(r[0]),
    n = _t(r[1] ?? r[0]);
  return {
    firstRow: Math.min(t.row, n.row),
    lastRow: Math.max(t.row, n.row),
    firstCol: Math.min(t.col, n.col),
    lastCol: Math.max(t.col, n.col),
  };
}
function parseSelectionAreas(e) {
  const r = typeof e == "string" ? e.split(",") : [...e];
  if (!r.length || r.length > maxSelectionAreas)
    throw Error(`Select between 1 and ${maxSelectionAreas} rectangular areas.`);
  const t = r.map((n) => {
    const rangeAddress = parseRangeAddress(
        n.trim().toUpperCase().replace(/\$/g, ""),
      ),
      a = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
      s = `${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
    return a === s ? a : `${a}:${s}`;
  });
  return [...new Set(t)];
}
function isCellAddressInRanges(selectionAreas, active) {
  const t = _t(active);
  return selectionAreas.some((selectionArea) => {
    const rangeAddress = parseRangeAddress(selectionArea);
    return (
      t.row >= rangeAddress.firstRow &&
      t.row <= rangeAddress.lastRow &&
      t.col >= rangeAddress.firstCol &&
      t.col <= rangeAddress.lastCol
    );
  });
}
function countSelectionCells(selectionAreas) {
  const r = parseSelectionAreas(selectionAreas).map(parseRangeAddress),
    t = [...new Set(r.flatMap((o) => [o.firstRow, o.lastRow + 1]))].sort(
      (o, a) => o - a,
    );
  let n = 0;
  for (let o = 0; o < t.length - 1; o++) {
    const a = r
      .filter((u) => u.firstRow <= t[o] && u.lastRow >= t[o])
      .map((u) => [u.firstCol, u.lastCol + 1])
      .sort((u, l) => u[0] - l[0]);
    let s = 0,
      d = -1;
    for (const [u, l] of a)
      ((s += Math.max(0, l - Math.max(u, d))), (d = Math.max(d, l)));
    n += s * (t[o + 1] - t[o]);
  }
  return n;
}
function enumerateSelectionCells(e, maxBlankScan = workbookLimits.maxPatches) {
  if (
    !Number.isSafeInteger(maxBlankScan) ||
    maxBlankScan < 1 ||
    maxBlankScan > workbookLimits.maxCells
  )
    throw Error("Invalid selection enumeration limit.");
  const selectionAreas = parseSelectionAreas(e),
    n = countSelectionCells(selectionAreas);
  if (n > maxBlankScan)
    throw Error(
      `Select at most ${maxBlankScan.toLocaleString()} cells for this operation. The current selection contains ${n.toLocaleString()} cells.`,
    );
  const o = new Set();
  for (const selectionArea of selectionAreas) {
    const rangeAddress = parseRangeAddress(selectionArea);
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
        o.add(`${columnIndexToLetters(firstCol)}${firstRow + 1}`);
  }
  return [...o];
}
function pQ(workbook, sourceSheetId, sortRangeAddress, sheetId, o) {
  const rangeAddress = parseRangeAddress(sortRangeAddress),
    s = parseRangeAddress(o),
    d = (f, p) => {
      const m = _t(p);
      return (
        m.row >= f.firstRow &&
        m.row <= f.lastRow &&
        m.col >= f.firstCol &&
        m.col <= f.lastCol
      );
    },
    changes = [],
    sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  return {
    watches: workbook.formulaWatches?.map((formulaWatch) => {
      if (formulaWatch.orphaned) return formulaWatch;
      const before = `${formulaWatch.anchor.sheetName}!${formulaWatch.anchor.address}`;
      if (
        formulaWatch.anchor.sheetId === sourceSheetId &&
        d(rangeAddress, formulaWatch.anchor.address)
      ) {
        const m = _t(formulaWatch.anchor.address),
          address = `${columnIndexToLetters(m.col + s.firstCol - rangeAddress.firstCol)}${m.row + s.firstRow - rangeAddress.firstRow + 1}`,
          after = `${sheet.name}!${address}`;
        return before === after && sourceSheetId === sheetId
          ? formulaWatch
          : (changes.push({
              kind: "formula-watch",
              label: formulaWatch.id,
              before: before,
              after: after,
            }),
            {
              ...formulaWatch,
              anchor: {
                sheetId: sheetId,
                sheetName: sheet.name,
                address: address,
              },
            });
      }
      return formulaWatch.anchor.sheetId === sheetId &&
        d(s, formulaWatch.anchor.address)
        ? (changes.push({
            kind: "formula-watch",
            label: formulaWatch.id,
            before: before,
            after: `${before} — watched source overwritten`,
          }),
          { ...formulaWatch, orphaned: "overwritten" })
        : formulaWatch;
    }),
    changes: changes,
  };
}
function validateFreezePanes(e, r, t = { rows: r.rowCount, cols: r.colCount }) {
  validateSheetDefaultSizes(e, r);
  const n = validateSplitPaneState(e.splitPanes, t);
  if (n && (e.frozenRows || e.frozenCols))
    throw Error("Split panes and frozen panes cannot be enabled together.");
  n && (r.splitPanes = n);
  const hiddenIntervals = normalizeHiddenIntervals(e.hiddenRows, t.rows),
    a = normalizeHiddenIntervals(e.hiddenCols, t.cols),
    s = { ...r, hiddenRows: hiddenIntervals, hiddenCols: a },
    axisVisibilityIndex = createAxisVisibilityIndex(r.rowCount, cn(s, "row")),
    u = createAxisVisibilityIndex(r.colCount, cn(s, "column"));
  if (!axisVisibilityIndex.count || !u.count)
    throw Error("At least one row and one column must remain visible.");
  const l = e.frozenRows === undefined ? 0 : e.frozenRows,
    c = e.frozenCols === undefined ? 0 : e.frozenCols;
  if (
    !Number.isInteger(l) ||
    !Number.isInteger(c) ||
    l < 0 ||
    c < 0 ||
    l > t.rows ||
    c > t.cols
  )
    throw Error(
      "Frozen boundaries must be valid leading row and column counts.",
    );
  const h = axisVisibilityIndex.countBefore(l),
    f = u.countBefore(c);
  if (
    h > worksheetCapacityLimits.maxFrozenVisibleRows ||
    f > worksheetCapacityLimits.maxFrozenVisibleCols
  )
    throw Error(
      "Freeze panes supports at most 100 visible rows and 50 visible columns.",
    );
  if (h === axisVisibilityIndex.count || f === u.count)
    throw Error("Freeze panes must leave an unfrozen visible row and column.");
  for (const p of r.merges ?? []) {
    const rangeAddress = parseRangeAddress(p);
    if (
      (rangeAddress.firstRow < l && rangeAddress.lastRow >= l) ||
      (rangeAddress.firstCol < c && rangeAddress.lastCol >= c)
    )
      throw Error(
        "Freeze panes cannot cross a merged range. Unmerge those cells first.",
      );
  }
  (e.hiddenRows !== undefined && (r.hiddenRows = hiddenIntervals),
    e.hiddenCols !== undefined && (r.hiddenCols = a),
    e.frozenRows !== undefined && (r.frozenRows = l),
    e.frozenCols !== undefined && (r.frozenCols = c));
}
function updateWorksheetById(workbook, sheetId, t) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Worksheet not found.");
  assertSheetUnprotected(sheet);
  const o = t(structuredClone(sheet));
  return (
    validateFreezePanes(o, o),
    {
      ...workbook,
      sheets: workbook.sheets.map((sheet) =>
        sheet.id === sheetId ? o : sheet,
      ),
    }
  );
}
function setHiddenAxisRange(workbook, sheetId, t, start, end, a) {
  if ((assertWorksheetAxis(t), typeof a != "boolean"))
    throw Error("Hidden state must be boolean.");
  return updateWorksheetById(workbook, sheetId, (s) => {
    const d = axisCountOf(s, t);
    normalizeHiddenIntervals([{ start: start, end: end }], d);
    const hiddenAxisIntervals = readHiddenAxisIntervals(s, t),
      l = { start: start, end: end },
      c = a
        ? hiddenAxisIntervals.filter((hiddenAxisInterval) =>
            hiddenAxisInterval.end < l.start - 1 ||
            hiddenAxisInterval.start > l.end + 1
              ? true
              : ((l.start = Math.min(l.start, hiddenAxisInterval.start)),
                (l.end = Math.max(l.end, hiddenAxisInterval.end)),
                false),
          )
        : hiddenAxisIntervals,
      h = a
        ? [...c, l]
        : hiddenAxisIntervals.flatMap((hiddenAxisInterval) =>
            hiddenAxisInterval.end < start || hiddenAxisInterval.start > end
              ? [hiddenAxisInterval]
              : [
                  ...(hiddenAxisInterval.start < start
                    ? [{ start: hiddenAxisInterval.start, end: start - 1 }]
                    : []),
                  ...(hiddenAxisInterval.end > end
                    ? [{ start: end + 1, end: hiddenAxisInterval.end }]
                    : []),
                ],
          );
    return ((s[hiddenKeyForAxis(t)] = normalizeHiddenIntervals(h, d)), s);
  });
}
function yQ(workbook, sheetId, t) {
  return (
    assertWorksheetAxis(t),
    updateWorksheetById(
      workbook,
      sheetId,
      (n) => (delete n[hiddenKeyForAxis(t)], n),
    )
  );
}
function setFreezePanes(workbook, sheetId, frozenRows, frozenCols) {
  return updateWorksheetById(
    workbook,
    sheetId,
    (o) => (
      delete o.splitPanes,
      { ...o, frozenRows: frozenRows, frozenCols: frozenCols }
    ),
  );
}
function setSplitPanes(workbook, sheetId, splitPanes) {
  return updateWorksheetById(
    workbook,
    sheetId,
    (n) => (
      splitPanes === undefined
        ? delete n.splitPanes
        : ((n.splitPanes = validateSplitPaneState(splitPanes, {
            rows: n.rowCount,
            cols: n.colCount,
          })),
          (n.frozenRows = 0),
          (n.frozenCols = 0)),
      n
    ),
  );
}
function setFreezePanesAtCell(workbook, sheetId, t) {
  const n = _t(t);
  return setFreezePanes(workbook, sheetId, n.row, n.col);
}
function applyAxisStructureEdit(e, axis, start, count, remove) {
  assertWorksheetAxis(axis);
  const a = axisCountOf(e, axis);
  if (
    !Number.isInteger(start) ||
    start < 0 ||
    !Number.isInteger(count) ||
    count < 1 ||
    count > 1e4 ||
    start + count > a ||
    typeof remove != "boolean"
  )
    throw Error("Choose a valid structural edit of 1–10,000 rows or columns.");
  const s = readHiddenAxisIntervals(e, axis).flatMap((f) => {
      if (!remove) {
        if (f.end < start) return [f];
        if (f.end + count >= a)
          throw Error(
            "Insertion would discard hidden-axis metadata at the worksheet boundary.",
          );
        return [
          {
            start: f.start >= start ? f.start + count : f.start,
            end: f.end + count,
          },
        ];
      }
      const p = start + count - 1;
      return f.end < start
        ? [f]
        : f.start > p
          ? [{ start: f.start - count, end: f.end - count }]
          : [
              ...(f.start < start ? [{ start: f.start, end: start - 1 }] : []),
              ...(f.end > p ? [{ start: start, end: f.end - count }] : []),
            ];
    }),
    d = hiddenKeyForAxis(axis),
    u = axis === "row" ? "frozenRows" : "frozenCols",
    l = e[u] ?? 0,
    c = remove
      ? l - Math.max(0, Math.min(count, l - start))
      : l + (start < l ? count : 0),
    h = {
      ...e,
      [d]: normalizeHiddenIntervals(s, a),
      [u]: c,
      ...(e.outlineGroups
        ? {
            outlineGroups: remapOutlineGroupsForEdit(
              e,
              axis,
              start,
              count,
              remove,
            ),
          }
        : {}),
    };
  return (
    validateFreezePanes(h, h),
    {
      [d]: h[d],
      [u]: c,
      ...(e.outlineGroups ? { outlineGroups: h.outlineGroups } : {}),
    }
  );
}
function setWorksheetReadOnly(workbook, sheetId, t, n = {}) {
  if (n.readOnly !== undefined && typeof n.readOnly != "boolean")
    throw Error("Read-only state must be boolean.");
  if (n.readOnly) throw Error("This workbook is read-only.");
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Worksheet not found.");
  assertSheetUnprotected(sheet);
  const a = structuredClone(sheet);
  return (
    t(a),
    validateOutlineGroups(a, a),
    validateFreezePanes(a, a),
    JSON.stringify(a) === JSON.stringify(sheet)
      ? workbook
      : {
          ...workbook,
          sheets: workbook.sheets.map((sheet) =>
            sheet.id === sheetId ? a : sheet,
          ),
        }
  );
}
function EQ(workbook, sheetId, axis, start, end, a) {
  return (
    assertWorksheetAxis(axis),
    setWorksheetReadOnly(
      workbook,
      sheetId,
      (s) => {
        s.outlineGroups = [
          ...(s.outlineGroups ?? []),
          {
            id: crypto.randomUUID(),
            axis: axis,
            start: start,
            end: end,
            collapsed: false,
          },
        ];
      },
      a,
    )
  );
}
function removeOutlineGroup(workbook, sheetId, t, n) {
  return setWorksheetReadOnly(
    workbook,
    sheetId,
    (o) => {
      if (!o.outlineGroups?.some((outlineGroup) => outlineGroup.id === t))
        throw Error("Outline group not found.");
      o.outlineGroups = o.outlineGroups.filter(
        (outlineGroup) => outlineGroup.id !== t,
      );
    },
    n,
  );
}
function setOutlineGroupCollapsed(workbook, sheetId, t, n, o) {
  if (typeof n != "boolean") throw Error("Collapsed state must be boolean.");
  return setWorksheetReadOnly(
    workbook,
    sheetId,
    (a) => {
      const s = a.outlineGroups?.find((outlineGroup) => outlineGroup.id === t);
      if (!s) throw Error("Outline group not found.");
      s.collapsed = n;
    },
    o,
  );
}
function setOutlineGroupLevel(workbook, sheetId, t, n, o) {
  return (
    assertWorksheetAxis(t),
    setWorksheetReadOnly(
      workbook,
      sheetId,
      (sheet) => {
        const s = addOutlineGroup(sheet, t),
          d = Math.max(0, ...s.map((c) => c.depth)) + 1;
        if (!Number.isInteger(n) || n < 1 || n > d)
          throw Error("Choose an available outline level.");
        const u = new Map(s.map((c) => [c.id, c.depth]));
        sheet.outlineGroups = sheet.outlineGroups?.map((outlineGroup) =>
          outlineGroup.axis === t
            ? { ...outlineGroup, collapsed: u.get(outlineGroup.id) >= n }
            : outlineGroup,
        );
      },
      o,
    )
  );
}
function clearAxisOutlineGroups(workbook, sheetId, t, n) {
  return (
    assertWorksheetAxis(t),
    setWorksheetReadOnly(
      workbook,
      sheetId,
      (o) => {
        o.outlineGroups = o.outlineGroups?.filter(
          (outlineGroup) => outlineGroup.axis !== t,
        );
      },
      n,
    )
  );
}
function normalizeDataValidationRules(validations) {
  if (
    !Array.isArray(validations) ||
    validations.length > validationLimits.rulesPerSheet
  )
    throw Error("At most 100 validation rules are supported.");
  const r = new Set();
  return validations.map((validation) => {
    if (
      !isNonArrayObject(validation) ||
      ![
        "list",
        "number",
        "integer",
        "textLength",
        "date",
        "time",
        "custom",
      ].includes(String(validation.kind)) ||
      ![
        "between",
        "notBetween",
        "greater",
        "greaterEqual",
        "less",
        "lessEqual",
        "equal",
        "notEqual",
      ].includes(String(validation.operator)) ||
      typeof validation.allowBlank != "boolean" ||
      !Array.isArray(validation.list) ||
      validation.list.length > validationLimits.listItems ||
      typeof validation.min != "number" ||
      !Number.isFinite(validation.min) ||
      typeof validation.max != "number" ||
      !Number.isFinite(validation.max)
    )
      throw Error("Invalid data validation rule.");
    const n = assertBoundedValidationText(validation.id, 100, "ID");
    if (!n.trim() || r.has(n))
      throw Error("Validation rule IDs must be nonempty and unique.");
    r.add(n);
    const rangeAddress = parseRangeAddress(
        assertBoundedValidationText(validation.range, 40, "range"),
      ),
      a = {
        id: n,
        range: `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`,
        kind: validation.kind,
        operator: validation.operator,
        min: validation.min,
        max: validation.max,
        list: validation.list.map((s) =>
          assertBoundedValidationText(s, 1e3, "list item"),
        ),
        allowBlank: validation.allowBlank,
        message: assertBoundedValidationText(
          validation.message,
          500,
          "error message",
        ),
      };
    if (a.kind === "list" && (!a.list.length || a.list.some((s) => !s.length)))
      throw Error("A list rule needs nonempty values.");
    if (
      !["list", "custom"].includes(a.kind) &&
      ["between", "notBetween"].includes(a.operator) &&
      a.min > a.max
    )
      throw Error("Minimum cannot exceed maximum.");
    if (
      a.kind === "time" &&
      (a.min < 0 || a.min >= 1 || a.max < 0 || a.max >= 1)
    )
      throw Error(
        "Time bounds must be fractions from 0 up to, but excluding, 1.",
      );
    if (
      validation.formula !== undefined &&
      ((a.formula = assertBoundedValidationText(
        validation.formula,
        8192,
        "formula",
      )),
      !a.formula.startsWith("=") || a.formula.length < 2)
    )
      throw Error("Start the custom validation formula with =.");
    if (a.kind === "custom" && !a.formula)
      throw Error(
        "A custom rule requires a formula relative to the range's top-left cell.",
      );
    if (validation.errorStyle !== undefined) {
      if (
        !["stop", "warning", "information"].includes(
          String(validation.errorStyle),
        )
      )
        throw Error("Invalid validation error style.");
      a.errorStyle = validation.errorStyle;
    }
    for (const s of ["showInputMessage", "showDropdown"])
      if (validation[s] !== undefined) {
        if (typeof validation[s] != "boolean")
          throw Error("Invalid validation display setting.");
        a[s] = validation[s];
      }
    for (const [s, d] of [
      ["inputTitle", 100],
      ["inputMessage", 500],
      ["errorTitle", 100],
    ])
      validation[s] !== undefined &&
        (a[s] = assertBoundedValidationText(validation[s], d, s));
    return a;
  });
}
function findDataValidationsInRange(activeWorkbook, sheetId, t) {
  const n = _t(t);
  return (
    activeWorkbook.sheets.find((sheet) => sheet.id === sheetId)?.validations ??
    []
  ).filter((a) => {
    const rangeAddress = parseRangeAddress(a.range);
    return (
      n.row >= rangeAddress.firstRow &&
      n.row <= rangeAddress.lastRow &&
      n.col >= rangeAddress.firstCol &&
      n.col <= rangeAddress.lastCol
    );
  });
}
function normalizeConditionalRules(e) {
  if (!Array.isArray(e) || e.length > 100)
    throw Error("At most 100 conditional formatting rules are supported.");
  const r = new Set(),
    t = new Set();
  return e.map((n) => {
    if (!isNonArrayObject(n) || !conditionalRuleKinds.includes(n.kind))
      throw Error("Invalid conditional rule.");
    const o = assertBoundedString(n.id, 100, "conditional rule ID");
    if (!o.trim() || r.has(o))
      throw Error("Conditional rule IDs must be nonempty and unique.");
    r.add(o);
    const rangeAddress = parseRangeAddress(
        assertBoundedString(n.range, 40, "conditional range"),
      ),
      s = {
        id: o,
        range: `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`,
        kind: n.kind,
        value: assertBoundedString(n.value, 1e3, "conditional value"),
        second: assertBoundedString(n.second, 1e3, "conditional second value"),
        fill: assertSixDigitHexColor(n.fill),
        color: assertSixDigitHexColor(n.color),
      };
    if (n.priority !== undefined) {
      if (
        !Number.isInteger(n.priority) ||
        Number(n.priority) < 1 ||
        Number(n.priority) > 100 ||
        t.has(Number(n.priority))
      )
        throw Error(
          "Conditional priorities must be unique integers from 1 to 100.",
        );
      ((s.priority = Number(n.priority)), t.add(s.priority));
    }
    if (n.stopIfTrue !== undefined) {
      if (typeof n.stopIfTrue != "boolean")
        throw Error("Stop if true must be boolean.");
      if (n.stopIfTrue && ["colorScale", "dataBar", "iconSet"].includes(s.kind))
        throw Error(
          "Stop if true is not supported for color scales, data bars, or icon sets.",
        );
      s.stopIfTrue = n.stopIfTrue;
    }
    if (
      (n.formula !== undefined &&
        (s.formula = assertBoundedString(
          n.formula,
          4096,
          "conditional formula",
        )),
      s.kind === "formula" &&
        (!s.formula?.startsWith("=") || !s.formula.slice(1).trim()))
    )
      throw Error("A formula rule requires a formula beginning with =.");
    for (const u of ["minColor", "maxColor"])
      n[u] !== undefined && (s[u] = assertSixDigitHexColor(n[u]));
    if (n.thresholdType !== undefined) {
      if (n.thresholdType !== "percent" && n.thresholdType !== "number")
        throw Error("Invalid icon threshold type.");
      s.thresholdType = n.thresholdType;
    }
    if (n.iconThresholds !== undefined) {
      if (
        !Array.isArray(n.iconThresholds) ||
        n.iconThresholds.length !== 2 ||
        n.iconThresholds.some(
          (iconThreshold) =>
            typeof iconThreshold != "number" || !Number.isFinite(iconThreshold),
        ) ||
        n.iconThresholds[0] >= n.iconThresholds[1] ||
        ((s.thresholdType ?? "percent") === "percent" &&
          (n.iconThresholds[0] < 0 || n.iconThresholds[1] > 100))
      )
        throw Error(
          "Choose two increasing icon thresholds; percentages must be between 0 and 100.",
        );
      s.iconThresholds = [...n.iconThresholds];
    }
    return s;
  });
}
function evaluateConditionalFormatRules(e, r, cells) {
  const n = new Map(),
    o = new Map();
  let a = 0,
    s = 0;
  const d = (ruleId, address, m) => {
    const v = { ruleId: ruleId, address: address, message: m };
    return (n.set(`${ruleId}:${m}`, v), v);
  };
  let u = "";
  if (
    !cells &&
    (e.conditionalFormats ?? []).some((f) =>
      ["duplicate", "unique", "colorScale", "dataBar", "iconSet"].includes(
        f.kind,
      ),
    )
  )
    try {
      cells = vx(e, r);
    } catch (error) {
      u =
        error instanceof Error
          ? error.message
          : "Array values could not be discovered for conditional statistics.";
    }
  const l = Object.keys(cells ?? e.cells).map((address) => ({
      address: address,
      point: _t(address),
    })),
    c = sortConditionalRulesByPriority(e.conditionalFormats ?? []).map(
      (rule) => {
        const rangeAddress = parseRangeAddress(rule.range),
          counts = new Map();
        let min = 1 / 0,
          max = -1 / 0,
          failure = "";
        if (
          (["greater", "less", "between"].includes(rule.kind) &&
            (!rule.value.trim() ||
              !Number.isFinite(Number(rule.value)) ||
              (rule.kind === "between" &&
                (!rule.second.trim() ||
                  !Number.isFinite(Number(rule.second)) ||
                  Number(rule.value) > Number(rule.second)))) &&
            (failure =
              "Numeric rule operands must be finite and minimum cannot exceed maximum."),
          ["duplicate", "unique", "colorScale", "dataBar", "iconSet"].includes(
            rule.kind,
          ))
        ) {
          u && (failure = u);
          for (const { address, point } of l) {
            if (++a > 1e6) {
              failure =
                "Conditional statistics exceed 1,000,000 inspected cells; narrow the rule ranges or remove rules.";
              break;
            }
            if (!isCellInsideRange(rangeAddress, point)) continue;
            const E = r(e.id, address);
            if (
              (typeof E == "number" &&
                Number.isFinite(E) &&
                ((min = Math.min(min, E)), (max = Math.max(max, E))),
              E !== null && E !== "")
            ) {
              const A = String(E).toLowerCase();
              counts.set(A, (counts.get(A) ?? 0) + 1);
            }
          }
        }
        return (
          failure && d(rule.id, "", failure),
          {
            rule: rule,
            bounds: rangeAddress,
            counts: counts,
            min: min,
            max: max,
            failure: failure,
          }
        );
      },
    ),
    h = (f) => {
      if (o.has(f)) return structuredClone(o.get(f));
      const p = _t(f),
        m = { style: {}, errors: [] },
        v = r(e.id, f);
      for (const { rule, bounds, counts, min, max, failure } of c) {
        if (!isCellInsideRange(bounds, p)) continue;
        if (failure) {
          m.errors.push(d(rule.id, f, failure));
          continue;
        }
        let k = false,
          F = {};
        if (rule.kind === "formula") {
          let I;
          try {
            if (++s > 5e3)
              throw Error(
                "Conditional formula evaluation exceeds 5,000 cells; narrow the visible or preview range.",
              );
            if (!r.evaluateFormula)
              throw Error(
                "This calculator cannot evaluate conditional formulas.",
              );
            const O = shiftFormulaReferences(
              rule.formula,
              p.row - bounds.firstRow,
              p.col - bounds.firstCol,
            );
            if (((I = r.evaluateFormula(e.id, f, O)), isErrorValueText(I)))
              throw Error(`Formula returned ${I}.`);
            if (typeof I != "number" && typeof I != "boolean")
              throw Error("Formula must return a scalar boolean or number.");
            k = !!I;
          } catch (error) {
            m.errors.push(
              d(
                rule.id,
                f,
                error instanceof Error
                  ? error.message
                  : "Formula evaluation failed.",
              ),
            );
            continue;
          }
        } else if (rule.kind === "blank" || rule.kind === "notBlank")
          k = (v === null || v === "") == (rule.kind === "blank");
        else if (rule.kind === "error" || rule.kind === "notError")
          try {
            if (++s > 5e3)
              throw Error(
                "Conditional formula evaluation exceeds 5,000 cells; narrow the visible or preview range.",
              );
            if (!r.evaluateFormula)
              throw Error(
                "This calculator cannot distinguish error values from literal text.",
              );
            const I = r.evaluateFormula(e.id, f, "=ISERROR(" + f + ")");
            if (typeof I != "boolean")
              throw Error("Error rule could not evaluate: " + String(I) + ".");
            k = I === (rule.kind === "error");
          } catch (error) {
            m.errors.push(
              d(
                rule.id,
                f,
                error instanceof Error
                  ? error.message
                  : "Error rule evaluation failed.",
              ),
            );
            continue;
          }
        else if (rule.kind === "duplicate" || rule.kind === "unique")
          k =
            v !== null &&
            v !== "" &&
            (counts.get(String(v).toLowerCase()) ?? 0) > 1 ==
              (rule.kind === "duplicate");
        else if (rule.kind === "contains")
          k = String(v ?? "")
            .toLowerCase()
            .includes(rule.value.toLowerCase());
        else if (rule.kind === "equal")
          k = String(v ?? "").toLowerCase() === rule.value.toLowerCase();
        else if (typeof v == "number") {
          if (rule.kind === "greater") k = v > Number(rule.value);
          else if (rule.kind === "less") k = v < Number(rule.value);
          else if (rule.kind === "between")
            k = v >= Number(rule.value) && v <= Number(rule.second);
          else if (Number.isFinite(min) && Number.isFinite(max)) {
            const I =
              max === min
                ? 0.5
                : Math.max(0, Math.min(1, (v - min) / (max - min)));
            if (rule.kind === "colorScale") {
              const O = rule.minColor ?? "#f8e99b",
                j = rule.maxColor ?? "#63be7b";
              ((F.fill =
                "#" +
                [1, 3, 5]
                  .map((z) =>
                    Math.round(
                      parseInt(O.slice(z, z + 2), 16) +
                        (parseInt(j.slice(z, z + 2), 16) -
                          parseInt(O.slice(z, z + 2), 16)) *
                          I,
                    )
                      .toString(16)
                      .padStart(2, "0"),
                  )
                  .join("")),
                (k = true));
            } else if (rule.kind === "dataBar" && !m.dataBar) {
              const O = Math.min(0, min),
                j = Math.max(0, max),
                z = j - O || 1,
                axis = (-O / z) * 100,
                D = ((v - O) / z) * 100;
              ((m.dataBar = {
                start: Math.min(axis, D),
                width: Math.abs(D - axis),
                axis: axis,
                color: v < 0 ? "#d13438" : rule.fill,
                negative: v < 0,
              }),
                (k = true));
            } else if (rule.kind === "iconSet" && !m.icon) {
              const O = rule.iconThresholds ?? [33, 67],
                j =
                  (rule.thresholdType ?? "percent") === "percent" ? I * 100 : v;
              ((m.icon =
                j < O[0]
                  ? { symbol: "▼", color: "#c42b1c", label: "Low" }
                  : j < O[1]
                    ? { symbol: "●", color: "#8a6100", label: "Middle" }
                    : { symbol: "▲", color: "#107c41", label: "High" }),
                (k = true));
            }
          }
        }
        if (
          k &&
          (["colorScale", "dataBar", "iconSet"].includes(rule.kind) ||
            (F = { fill: rule.fill, color: rule.color }),
          (m.style = { ...F, ...m.style }),
          rule.stopIfTrue)
        )
          break;
      }
      return (o.size < 1e4 && o.set(f, structuredClone(m)), m);
    };
  return ((h.errors = () => [...n.values()].map((f) => ({ ...f }))), h);
}
const normalizeWorksheetToolText = (e, r = 200) =>
    typeof e == "string" && e.length <= r
      ? e
      : failWithMessage("Invalid worksheet tool text."),
  CE = (e) => {
    const rangeAddress = parseRangeAddress(e);
    return `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
  },
  doesRangeOverlap = (e, r) => {
    const rangeAddress = parseRangeAddress(e),
      n = parseRangeAddress(r);
    return (
      rangeAddress.firstRow <= n.lastRow &&
      rangeAddress.lastRow >= n.firstRow &&
      rangeAddress.firstCol <= n.lastCol &&
      rangeAddress.lastCol >= n.firstCol
    );
  },
  lD = (e, r) => {
    const rangeAddress = parseRangeAddress(e),
      n = _t(r);
    return (
      n.row >= rangeAddress.firstRow &&
      n.row <= rangeAddress.lastRow &&
      n.col >= rangeAddress.firstCol &&
      n.col <= rangeAddress.lastCol
    );
  };
function validateWorksheetFeatures(sheet, r) {
  if (isNonArrayObject(sheet)) {
    if (sheet.merges !== undefined) {
      (!Array.isArray(sheet.merges) || sheet.merges.length > 1e3) &&
        failWithMessage("At most1,000 merged ranges are supported.");
      const t = [];
      for (const merge of sheet.merges) {
        const o = CE(normalizeWorksheetToolText(merge, 40)),
          rangeAddress = parseRangeAddress(o);
        ((rangeAddress.lastRow - rangeAddress.firstRow + 1) *
          (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
          500 && failWithMessage("A merged area supports at most 500 cells."),
          t.some((s) => doesRangeOverlap(s, o)) &&
            failWithMessage("Merged ranges cannot overlap."),
          t.push(o));
      }
      r.merges = t;
    }
    if (
      (sheet.validations !== undefined &&
        (r.validations = normalizeDataValidationRules(sheet.validations)),
      sheet.conditionalFormats !== undefined &&
        (r.conditionalFormats = normalizeConditionalRules(
          sheet.conditionalFormats,
        )),
      sheet.protection !== undefined)
    ) {
      const protection = sheet.protection;
      ((!isNonArrayObject(protection) ||
        typeof protection.enabled != "boolean" ||
        (protection.hash !== undefined &&
          (!/^[a-f\d]{64}$/i.test(String(protection.hash)) ||
            typeof protection.salt != "string" ||
            protection.salt.length > 100))) &&
        failWithMessage("Invalid sheet protection."),
        (r.protection = {
          enabled: protection.enabled,
          ...(protection.hash
            ? { hash: String(protection.hash), salt: String(protection.salt) }
            : {}),
        }));
    }
  }
}
function mergeWorksheetCells(workbook, sheetId, t, unmerge = false) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  assertSheetUnprotected(sheet);
  const a = CE(t),
    rangeAddress = parseRangeAddress(a);
  (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
    (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    500 && failWithMessage("Merge at most 500 cells.");
  const d = (sheet.merges ?? []).filter((h) => doesRangeOverlap(h, a));
  !unmerge && d.length && failWithMessage("Unmerge overlapping areas first.");
  const cells = { ...sheet.cells };
  if (!unmerge) {
    const h = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`;
    cells[h] = {
      ...cells[h],
      raw: cells[h]?.raw ?? null,
      style: { ...cells[h]?.style, align: "center" },
    };
  }
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet.id === sheetId
        ? {
            ...sheet,
            cells: cells,
            merges: unmerge
              ? (sheet.merges ?? []).filter((f) => !d.includes(f))
              : [...(sheet.merges ?? []), a],
          }
        : sheet,
    ),
  };
}
function oP(names) {
  let r = 0;
  for (const { refersTo } of names) {
    const n = stripFormulaStringLiterals(
      refersTo.replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')+'!/g, ""),
    );
    for (const o of n.matchAll(
      /(?:^|[^A-Za-z0-9_.])((?:\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?)(?![A-Za-z0-9_.])/g,
    )) {
      const rangeAddress = parseRangeAddress(o[1]);
      if (
        ((r +=
          (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
          (rangeAddress.lastCol - rangeAddress.firstCol + 1)),
        r > workbookLimits.maxCells)
      )
        throw Error(
          "Named ranges exceed the 100,000-cell XLSX import/export processing budget. Use native JSON for larger definitions.",
        );
    }
  }
}
function aP(e) {
  const definedNamesXml = parseDefinedNamesXml(e),
    t = { names: [], warnings: [], builtinXml: "" };
  if (!definedNamesXml) return t;
  let body = definedNamesXml.body,
    o = 0;
  const a = new Set(),
    s = [...e.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b[^>]*>/g)].length;
  for (; body.trim();) {
    const d =
      /^\s*<((?:[A-Za-z_][\w.-]*:)?definedName)\b((?:[^"'<>]|"[^"]*"|'[^']*')*?)(?:\/>|>([^<]*)<\/\1\s*>)/.exec(
        body,
      );
    if (!d)
      return throwXlsxDefinedNameError(
        "Malformed definedName content; only text definitions are supported.",
      );
    const definedNameXmlAttributes = parseDefinedNameXmlAttributes(d[2]),
      l = definedNameXmlAttributes.name ?? "",
      localSheetId = definedNameXmlAttributes.localSheetId,
      h = d[1].slice(0, -11);
    if (
      { ...definedNamesXml.namespaces, ...definedNameXmlAttributes }[
        h ? `xmlns:${h.slice(0, -1)}` : "xmlns"
      ] !== spreadsheetMlNamespace
    )
      return throwXlsxDefinedNameError(
        "Unsupported or undeclared definedName namespace.",
      );
    for (const definedNameXmlAttribute of Object.keys(definedNameXmlAttributes))
      if (
        definedNameXmlAttribute.startsWith("xmlns") &&
        definedNameXmlAttributes[definedNameXmlAttribute] !==
          spreadsheetMlNamespace
      )
        return throwXlsxDefinedNameError("Unsupported definedName namespace.");
    if (
      localSheetId !== undefined &&
      (!/^\d+$/.test(localSheetId) ||
        !Number.isSafeInteger(Number(localSheetId)))
    )
      return throwXlsxDefinedNameError(
        `Name "${l}" has an invalid localSheetId.`,
      );
    const refersTo = `=${nP(
      (d[3] ?? "").replace(
        /\r\n?/g,
        `
`,
      ),
    ).replace(/^=/, "")}`;
    if (
      (validateNameFormulaText(refersTo),
      localSheetId !== undefined && Number(localSheetId) >= s)
    )
      return throwXlsxDefinedNameError(
        `Name "${l}" has an out-of-range localSheetId.`,
      );
    const m = `${localSheetId === undefined ? "workbook" : Number(localSheetId)}:${l.toLowerCase()}`;
    if (a.has(m))
      return throwXlsxDefinedNameError(
        `Duplicate name "${l}" in the same scope.`,
      );
    if ((a.add(m), reservedDefinedNames.has(l.toLowerCase())))
      t.builtinXml += `<definedName${Object.entries(definedNameXmlAttributes)
        .filter(([v]) => !v.startsWith("xmlns"))
        .map(([v, g]) => ` ${v}="${validateXmlSafeName(g, true)}"`)
        .join("")}>${validateXmlSafeName(refersTo.slice(1))}</definedName>`;
    else if (l.toLowerCase().startsWith("_xlnm.")) {
      if (++o > 1e3)
        return throwXlsxDefinedNameError(
          "The workbook exceeds the 1,000-defined-name limit.",
        );
      t.warnings.push(
        `Built-in defined name "${l}" is not supported and was omitted. Review its source feature before relying on imported formulas.`,
      );
    } else {
      if (++o > 1e3)
        return throwXlsxDefinedNameError(
          "The workbook exceeds the 1,000-defined-name limit.",
        );
      for (const g of ["function", "vbProcedure", "xlm"])
        if (
          definedNameXmlAttributes[g] !== undefined &&
          !["0", "false"].includes(definedNameXmlAttributes[g])
        )
          return throwXlsxDefinedNameError(
            `Name "${l}" has an executable ${g} attribute.`,
          );
      const v = Object.keys(definedNameXmlAttributes).filter(
        (definedNameXmlAttribute) =>
          !["name", "localSheetId", "comment"].includes(
            definedNameXmlAttribute,
          ) && !definedNameXmlAttribute.startsWith("xmlns"),
      );
      (v.length &&
        t.warnings.push(
          `Defined name "${l}" retained its definition and scope; unsupported attributes (${v.join(", ")}) were omitted. Hidden names are visible in Name Manager.`,
        ),
        t.names.push({
          name: l,
          refersTo: refersTo,
          ...(localSheetId !== undefined
            ? { localSheetId: Number(localSheetId) }
            : {}),
          ...(definedNameXmlAttributes.comment !== undefined
            ? { comment: definedNameXmlAttributes.comment }
            : {}),
        }));
    }
    body = body.slice(d[0].length);
  }
  oP(t.names);
  try {
    validateDefinedName(
      t.names.map((d) => ({
        name: d.name,
        refersTo: d.refersTo,
        ...(d.localSheetId !== undefined
          ? { scopeSheetId: String(d.localSheetId) }
          : {}),
        ...(d.comment !== undefined ? { comment: d.comment } : {}),
      })),
      {
        sheets: [
          ...new Set(
            t.names.flatMap((d) =>
              d.localSheetId === undefined ? [] : [d.localSheetId],
            ),
          ),
        ].map((d) => ({ id: String(d) })),
      },
    );
  } catch (error) {
    return throwXlsxDefinedNameError(error.message);
  }
  return t;
}

export {
  doesRangeOverlap,
  setOutlineGroupCollapsed,
  CE,
  removeOutlineGroup,
  findDataValidationsInRange,
  resolveFilteredRowContext,
  normalizeReviewMetadata,
  EQ,
  normalizeConditionalRules,
  parsePrintTitles,
  evaluateConditionalFormatRules,
  formatAxisIntervalRef,
  isCanonicalCellAddress,
  setFreezePanes,
  setSplitPanes,
  validateFreezePanes,
  validateWhatIfDefinition,
  enumerateSelectionCells,
  attachReviewThreadsToWorkbook,
  createAnchoredDrawingObject,
  clearAxisOutlineGroups,
  validateFormulaWatches,
  validateWorksheetFeatures,
  isCellAddressInRanges,
  validateHyperlinkRecord,
  applyChartDrawingLayout,
  normalizeDataValidationRules,
  _t,
  aP,
  parseRangeAddress,
  setHiddenAxisRange,
  shiftFormulaWatchesOnAxisEdit,
  buildDrawingExportNotes,
  parseSelectionAreas,
  shiftPrintSettingsOnAxisEdit,
  mergeWorksheetCells,
  setOutlineGroupLevel,
  lD,
  oP,
  buildBorderPresetStyles,
  pQ,
  normalizeCellAddress,
  shiftReviewTargetsOnAxisEdit,
  validateSavedWhatIfAnalyses,
  analysisOutputEvidenceSchema,
  setFreezePanesAtCell,
  applyAxisStructureEdit,
  countSelectionCells,
  yQ,
};

/* Also used by the workers. */
export {
  evidenceAddressSchema,
  evidenceCellRecordSchema,
  evidenceFieldValueSchema,
  evidenceSourceOriginSchema,
  evidenceTransformStepSchema,
};

/* Also used by the workers. */
export {
  normalizeFormulaWatchAddress,
  normalizeWhatIfCell,
  normalizeWhatIfCellReferences,
  normalizeWorksheetToolText,
};
