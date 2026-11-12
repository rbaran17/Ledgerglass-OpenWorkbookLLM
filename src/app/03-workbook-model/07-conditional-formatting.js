import { isPlainObjectValue } from "./01-cell-styles-and-themes.js";
import { isFiniteNumber } from "./04-pivot-tables-and-pictures.js";
import {
  workbookLimits,
  maxSelectionAreas,
} from "../02-financial-functions.js";
import {
  hasLinkedOrBoundCells,
  isoDateTimeSchemaFactory,
} from "../01-core-schemas-and-hashing.js";
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
} from "./02-number-formats.js";
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

function resolveFilteredRowContext(e, r = {}) {
  if (
    !isPlainObjectValue(r) ||
    Object.keys(r).some((u) => u !== "filteredRowsBySheet")
  )
    throw Error("Calculation options support only filteredRowsBySheet.");
  const t = r.filteredRowsBySheet;
  if (t !== undefined && !isPlainObjectValue(t))
    throw Error(
      "Filtered rows must be a worksheet-ID map of inclusive row intervals.",
    );
  const n = new Map(e.sheets.map((u) => [u.id, u])),
    o = new Map(e.sheets.map((u) => [u.name.toLowerCase(), u])),
    a = new Map();
  let s = 0;
  for (const [u, l] of Object.entries(t ?? {})) {
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
  for (const u of e.sheets)
    d.set(u.id, createAxisVisibilityIndex(u.rowCount, cn(u, "row")));
  return (u, l, c) => {
    if (!referenceSensitiveFunctionNames.includes(u))
      return notReferenceFunctionMarker;
    const { issue, isIssue } = c,
      p = (J) => {
        if (isIssue(J)) throw J;
        if (!isPlainObjectValue(J) || !isPlainObjectValue(J.ref))
          throw issue("#VALUE!");
        const q = J.ref;
        if (q.sheet !== undefined && typeof q.sheet != "string")
          throw issue("#REF!");
        const W = o.get(
          String(q.sheet ?? c.currentSheetName ?? "").toLowerCase(),
        );
        if (!W) throw issue("#REF!");
        const K = isPlainObjectValue(q.from) ? q.from : q,
          te = isPlainObjectValue(q.to) ? q.to : q,
          ae = [K.row, K.col, te.row, te.col];
        if (ae.some((le) => !Number.isInteger(le))) throw issue("#REF!");
        const [oe, de, he, Se] = ae;
        if (
          oe < 1 ||
          de < 1 ||
          he < oe ||
          Se < de ||
          he > W.rowCount ||
          Se > W.colCount
        )
          throw issue("#REF!");
        return {
          sheet: W,
          firstRow: oe,
          lastRow: he,
          firstCol: de,
          lastCol: Se,
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
        const K = q(J.firstRow, J.firstCol);
        return isFormulaCellValue(K) && K.raw.length <= 8192
          ? K.raw
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
      for (let q = J.firstRow; q <= J.lastRow; q++)
        for (let W = J.firstCol; W <= J.lastCol; W++)
          if (c.isActive(J.sheet.id, `${columnIndexToLetters(W - 1)}${q}`))
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
      for (let q = J.firstRow; q <= J.lastRow; q++)
        if (!(
          a.get(J.sheet.id)?.isHidden(q - 1) ||
          (E && d.get(J.sheet.id).isHidden(q - 1))
        ))
          for (let W = J.firstCol; W <= J.lastCol; W++) {
            const K = `${columnIndexToLetters(W - 1)}${q}`;
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
  const n = Number(r[2]) - 1;
  return n >= workbookLimits.maxRows || t > workbookLimits.maxCols
    ? throwWithMessage(`Cell ${e} exceeds the supported workbook bounds.`)
    : { row: n, col: t - 1 };
}
function buildBorderPresetStyles(
  e,
  r,
  t,
  n = { style: "thin", color: "#000000" },
) {
  const o = _t(e),
    a = validateCellBorder(n),
    s = {};
  ["all", "outline", "inside", "clear"].includes(t) ||
    failWithMessage("Choose a supported border preset.");
  for (const d of cellBorderEdges) {
    const u = o.row + (d === "top" ? -1 : d === "bottom" ? 1 : 0),
      l = o.col + (d === "left" ? -1 : d === "right" ? 1 : 0),
      c =
        u >= 0 &&
        u < 1048576 &&
        l >= 0 &&
        l < 16384 &&
        r.has(`${columnIndexToLetters(l)}${u + 1}`);
    t === "clear"
      ? (s[d] = { style: "none" })
      : (t === "all" || (t === "outline" ? !c : c)) && (s[d] = { ...a });
  }
  return s;
}
function normalizeCellAddress(e) {
  if (typeof e != "string")
    return throwWithMessage("Choose a valid worksheet cell address.");
  const r = _t(e.trim());
  return `${columnIndexToLetters(r.col)}${r.row + 1}`;
}
function validateHyperlinkRecord(e) {
  const r = assertPlainReviewRecord(e, "Hyperlink"),
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
  const n = normalizeBoundedText(
    r.url,
    "Link address",
    reviewThreadLimits.urlLength,
  );
  let o = n;
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
    a = new URL(n);
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
function normalizeReviewMetadata(e, r) {
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
    const s = t.reviewThreads.map((d) => {
      const u = assertPlainReviewRecord(d, "Review thread"),
        l = assertPlainReviewRecord(u.anchor, "Review target"),
        c = normalizeBoundedText(u.id, "Review thread ID", 100);
      (n.has(c) && throwWithMessage("Review thread IDs must be unique."),
        n.add(c));
      const h = normalizeReviewTimestamp(u.createdAt),
        f = normalizeReviewTimestamp(u.updatedAt);
      f < h && throwWithMessage("Review update time cannot precede creation.");
      const p = {
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
      let m = h;
      const v = u.messages.map((b, y) => {
          const w = assertPlainReviewRecord(b, "Review message"),
            E = normalizeBoundedText(w.id, "Review message ID", 100);
          (o.has(E) &&
            throwWithMessage(
              "Review message IDs must be unique within the workbook.",
            ),
            o.add(E));
          const A = normalizeReviewTimestamp(w.createdAt);
          return (
            (A < m || A > f || (y === 0 && A !== h)) &&
              throwWithMessage(
                "Review message times must follow their discussion in order.",
              ),
            (m = A),
            {
              id: E,
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
        g = { id: c, anchor: p, messages: v, createdAt: h, updatedAt: f };
      if (u.resolved !== undefined) {
        const b = assertPlainReviewRecord(u.resolved, "Review resolution"),
          y = normalizeReviewTimestamp(b.createdAt);
        ((y < v.at(-1).createdAt || y > f) &&
          throwWithMessage("Invalid review resolution time."),
          (g.resolved = {
            author: normalizeReviewAuthor(
              normalizeBoundedText(b.author, "Reviewer name", 100),
            ),
            createdAt: y,
          }));
      }
      if (u.orphaned !== undefined) {
        const b = assertPlainReviewRecord(u.orphaned, "Orphaned review target"),
          y = normalizeReviewTimestamp(b.createdAt);
        ((!["worksheet-deleted", "cell-deleted"].includes(String(b.reason)) ||
          y < h ||
          y > f) &&
          throwWithMessage("Invalid orphaned review target."),
          (g.orphaned = { reason: b.reason, createdAt: y }));
      }
      return (
        !g.orphaned &&
          !r.sheets.some((b) => b.id === p.sheetId) &&
          throwWithMessage(
            "A review thread with a missing worksheet must be retained as orphaned.",
          ),
        g
      );
    });
    s.length ? (r.reviewThreads = s) : delete r.reviewThreads;
  }
  for (const n of r.sheets)
    for (const o of hasLinkedOrBoundCells(n.cells)
      ? Object.values(n.cells)
      : []) {
      if (o.hyperlink === undefined) continue;
      const a = validateHyperlinkRecord(o.hyperlink);
      (a.kind === "internal" &&
        !r.sheets.some((s) => s.id === a.sheetId) &&
        throwWithMessage(
          "An internal hyperlink must target an existing worksheet.",
        ),
        (o.hyperlink = a));
    }
}
function attachReviewThreadsToWorkbook(e, r) {
  const t = { ...e, reviewThreads: r },
    n = {
      ...t,
      sheets: t.sheets.map((o) => ({
        ...o,
        cells: Object.fromEntries(
          Object.entries(o.cells).map(([a, s]) => [a, { ...s }]),
        ),
      })),
    };
  return (normalizeReviewMetadata(t, n), n);
}
function shiftReviewTargetsOnAxisEdit(e, r, t, n) {
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
  const d = e.reviewThreads?.map((c) => {
      if (c.orphaned || c.anchor.sheetId !== r) return c;
      const h = a(c.anchor.address);
      if (h === c.anchor.address) return c;
      if (((s = true), h)) return { ...c, anchor: { ...c.anchor, address: h } };
      const f = resolveReviewTimestamp({ timestamp: n }, c.updatedAt);
      return {
        ...c,
        updatedAt: f,
        orphaned: { reason: "cell-deleted", createdAt: f },
      };
    }),
    u = e.sheets.map((c) => {
      let h;
      for (const [f, p] of Object.entries(c.cells)) {
        const m = p.hyperlink;
        if (m?.kind !== "internal" || m.sheetId !== r) continue;
        const v = a(m.address);
        v !== m.address &&
          ((h ??= { ...c.cells }),
          (h[f] = { ...p }),
          (s = true),
          v ? (h[f].hyperlink = { ...m, address: v }) : delete h[f].hyperlink);
      }
      return h ? { ...c, cells: h } : c;
    });
  return s ? { ...e, sheets: u, ...(d ? { reviewThreads: d } : {}) } : e;
}
function shiftFormulaWatchesOnAxisEdit(e, r, t) {
  return e.formulaWatches?.length
    ? {
        ...e,
        formulaWatches: e.formulaWatches.map((o) => {
          if (o.orphaned || o.anchor.sheetId !== r) return o;
          const a = _t(o.anchor.address),
            s = t.axis === "row" ? a.row : a.col;
          if (
            t.moveBefore === undefined &&
            t.remove &&
            s >= t.index &&
            s < t.index + t.count
          )
            return { ...o, orphaned: "deleted-cells" };
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
            return { ...o, orphaned: "out-of-bounds" };
          const u =
            t.axis === "row"
              ? `${columnIndexToLetters(a.col)}${d + 1}`
              : `${columnIndexToLetters(d)}${a.row + 1}`;
          return u === o.anchor.address
            ? o
            : { ...o, anchor: { ...o.anchor, address: u } };
        }),
      }
    : e;
}
function createAnchoredDrawingObject(e, r) {
  const t = _t(r.split(":")[0]);
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
function buildDrawingExportNotes(e, r, t) {
  const n = [];
  for (const s of t.objects ?? []) {
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
        col: u.col + s.offsetX / (t.columnWidths?.[u.col] ?? 84),
        row: u.row + s.offsetY / (t.rowHeights?.[u.row] ?? 21),
      },
      ext: { width: s.width, height: s.height },
      editAs: "oneCell",
    });
  }
  return n;
}
function normalizeFormulaWatchAddress(e) {
  if (typeof e != "string" || !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}$/.test(e))
    throw Error("A formula watch must identify one valid Excel cell address.");
  const r = _t(e.replace(/\$/g, "").toUpperCase());
  if (r.row >= workbookLimits.maxRows || r.col >= workbookLimits.maxCols)
    throw Error("A formula watch address is outside Excel's cell bounds.");
  return `${columnIndexToLetters(r.col)}${r.row + 1}`;
}
function validateFormulaWatches(e, r) {
  if (e === undefined) return [];
  if (!Array.isArray(e) || e.length > B2)
    throw Error("A workbook may contain at most 100 formula watches.");
  const t = new Set(),
    n = new Set();
  return e.map((o) => {
    if (
      !isPlainObjectValue(o) ||
      !HO(o, ["id", "anchor", "orphaned"]) ||
      !isBoundedPlainText(o.id, 100) ||
      !isPlainObjectValue(o.anchor) ||
      !HO(o.anchor, ["sheetId", "sheetName", "address"]) ||
      !isBoundedPlainText(o.anchor.sheetId, 100) ||
      !isBoundedPlainText(o.anchor.sheetName, 31)
    )
      throw Error("Invalid formula watch metadata.");
    if (t.has(o.id)) throw Error("Formula watch IDs must be unique.");
    if (
      (t.add(o.id),
      o.orphaned !== undefined && !referenceImpactKinds.includes(o.orphaned))
    )
      throw Error("Invalid formula watch orphan reason.");
    const a = o.anchor,
      s = normalizeFormulaWatchAddress(a.address),
      d = r.sheets.find((c) => c.id === a.sheetId),
      u = o.orphaned ?? (d ? undefined : "missing-sheet"),
      l = JSON.stringify([o.anchor.sheetId, s]);
    if (!u && n.has(l))
      throw Error("Active formula watches must not repeat the same cell.");
    return (
      u || n.add(l),
      {
        id: o.id,
        anchor: {
          sheetId: o.anchor.sheetId,
          sheetName: !u && d ? d.name : o.anchor.sheetName,
          address: s,
        },
        ...(u ? { orphaned: u } : {}),
      }
    );
  });
}
function normalizeWhatIfCell(e) {
  const r = validateWhatIfFields(e, ["sheetId", "address"]),
    t = requireWhatIfName(r.sheetId);
  if (typeof r.address != "string")
    throw Error("A what-if cell needs an Excel address.");
  const n = _t(r.address);
  return { sheetId: t, address: `${columnIndexToLetters(n.col)}${n.row + 1}` };
}
function normalizeWhatIfCellReferences(e, r) {
  const t = requireWhatIfList(e, r).map(normalizeWhatIfCell);
  if (new Set(t.map(cellAddressKey)).size !== t.length)
    throw Error("What-if cell references must not repeat.");
  return t;
}
function validateWhatIfDefinition(e) {
  const r = validateWhatIfFields(e, [
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
    t.inputs.some((o) =>
      t.outputs.some((a) => cellAddressKey(o) === cellAddressKey(a)),
    )
  )
    throw Error("Changing inputs and formula results must be different cells.");
  const n = ["name", "kind", "inputs", "outputs"];
  if (r.kind === "scenarios") {
    validateWhatIfFields(r, [...n, "scenarios"]);
    const o = requireWhatIfList(r.scenarios, whatIfLimits.scenarios).map(
      (a) => {
        const s = validateWhatIfFields(a, ["name", "values"]),
          d = validateWhatIfValues(s.values, whatIfLimits.inputs);
        if (d.length !== t.inputs.length)
          throw Error(
            "Each scenario needs exactly one value per changing input.",
          );
        return { name: requireWhatIfName(s.name), values: d };
      },
    );
    if (new Set(o.map((a) => a.name.toLowerCase())).size !== o.length)
      throw Error("Scenario names must be unique, ignoring case.");
    return { ...t, kind: r.kind, scenarios: o };
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
    const o = validateWhatIfValues(r.rowValues, whatIfLimits.axisValues),
      a = validateWhatIfValues(r.columnValues, whatIfLimits.axisValues);
    if (o.length * a.length > whatIfLimits.trials)
      throw Error("A sensitivity table supports at most 400 combinations.");
    return { ...t, kind: r.kind, rowValues: o, columnValues: a };
  }
  throw Error(
    "Choose scenarios, a one-variable table or a two-variable table.",
  );
}
function validateSavedWhatIfAnalyses(e) {
  const r = new Set();
  return requireWhatIfList(e, whatIfLimits.savedAnalyses, 0).map((t) => {
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
      a = requireWhatIfName(n.reportSheetId),
      s = validateWhatIfDefinition(n.request);
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
          g = validateWhatIfValues(m.inputs, whatIfLimits.inputs),
          b =
            m.outputs === null
              ? null
              : validateWhatIfValues(m.outputs, whatIfLimits.outputs),
          y = m.error === null ? null : requireWhatIfName(m.error, 2e3);
        if (
          g.length !== s.inputs.length ||
          (b && b.length !== s.outputs.length) ||
          !!b == !!y
        )
          throw Error("Invalid saved what-if result dimensions or outcome.");
        return { name: v, inputs: g, outputs: b, error: y };
      },
      u = d(n.baseline),
      l = requireWhatIfList(n.trials, whatIfLimits.trials).map(d),
      c = expandWhatIfScenarios(s);
    if (
      l.length !== c.length ||
      l.some(
        (p, m) =>
          p.name !== c[m].name ||
          JSON.stringify(p.inputs) !== JSON.stringify(c[m].inputs),
      )
    )
      throw Error("Saved what-if results must match their named assumptions.");
    if (u.name !== "Baseline")
      throw Error("Saved what-if baseline is missing.");
    const h = new Set(),
      f = requireWhatIfList(n.assumptions, whatIfLimits.dependencyCells).map(
        (p) => {
          const m = validateWhatIfFields(p, [
              "sheetId",
              "sheetName",
              "address",
              "raw",
              "literal",
            ]),
            v = normalizeWhatIfCell({ sheetId: m.sheetId, address: m.address });
          if (h.has(cellAddressKey(v)))
            throw Error("Saved assumptions must not repeat a source cell.");
          if (
            (h.add(cellAddressKey(v)),
            !(
              m.raw === null ||
              typeof m.raw == "boolean" ||
              isFiniteNumber(m.raw) ||
              (typeof m.raw == "string" && m.raw.length <= 32767)
            ) ||
              (m.literal !== undefined && typeof m.literal != "boolean"))
          )
            throw Error("Invalid saved what-if assumption.");
          return {
            ...v,
            sheetName: requireWhatIfName(m.sheetName, 31),
            raw: m.raw,
            ...(m.literal === undefined ? {} : { literal: m.literal }),
          };
        },
      );
    if (
      [...s.inputs, ...s.outputs].some((p) => !h.has(cellAddressKey(p))) ||
      s.inputs.some(
        (p, m) =>
          f.find((g) => cellAddressKey(g) === cellAddressKey(p))?.raw !==
          u.inputs[m],
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
      n.dependencyCells !== f.length ||
      n.workUnits !== f.length * (l.length + 1)
    )
      throw Error("What-if work counts do not reconcile.");
    return {
      id: o,
      reportSheetId: a,
      request: s,
      baseline: u,
      trials: l,
      assumptions: f,
      sourceRevision: n.sourceRevision,
      dependencyCells: f.length,
      workUnits: n.workUnits,
    };
  });
}
function applyChartDrawingLayout(e, r) {
  if (r.kind === "chart" && r.chart === "sparkline") {
    const t = _t(r.anchor);
    return {
      offsetX: 2,
      offsetY: 2,
      width: columnWidthAt(e, t.col) - 4,
      height: rowHeightAt(e, t.row) - 4,
    };
  }
  return {
    offsetX: r.offsetX,
    offsetY: r.offsetY,
    width: r.width,
    height: r.height,
  };
}
function parsePrintTitles(e, r, t) {
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
  const o = r === "row" ? Number(n[1]) - 1 : _t(n[1] + "1").col,
    a = r === "row" ? Number(n[2]) - 1 : _t(n[2] + "1").col,
    s = r === "row" ? t.rowCount : t.colCount,
    d =
      r === "row"
        ? printRenderLimits.titleRows
        : printRenderLimits.titleColumns;
  if (
    !Number.isSafeInteger(o) ||
    !Number.isSafeInteger(a) ||
    o < 0 ||
    a < o ||
    a >= s ||
    a - o + 1 > d
  )
    throw Error(
      `Print titles support at most ${d} ${r === "row" ? "rows" : "columns"} within the worksheet.`,
    );
  return { start: o, end: a };
}
function shiftPrintSettingsOnAxisEdit(e, r, t, n, o, a) {
  const s = t === "row" ? "repeatRows" : "repeatColumns",
    d = t === "row" ? "rowBreaks" : "columnBreaks",
    u = { ...e },
    l = parsePrintTitles(e[s], t, r);
  if (l) {
    let { start, end } = l;
    if (!a) (n <= start && (start += o), n <= end && (end += o));
    else {
      const m = n + o - 1,
        v = [
          start < n ? { start, end: Math.min(end, n - 1) } : undefined,
          end > m
            ? { start: Math.max(start, m + 1) - o, end: end - o }
            : undefined,
        ].filter(Boolean);
      if (!v.length)
        throw Error(
          "This edit removes every repeated print title. Clear or change print titles first.",
        );
      ((start = v[0].start), (end = v.at(-1).end));
    }
    ((u[s] = formatAxisIntervalRef({ start, end }, t)),
      parsePrintTitles(u[s], t, r));
  }
  const c = t === "row" ? r.rowCount : r.colCount,
    h = (e[d] ?? [])
      .map((f) =>
        a ? (f < n ? f : f >= n + o ? f - o : n) : f >= n ? f + o : f,
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
    const o = parseRangeAddress(n.trim().toUpperCase().replace(/\$/g, "")),
      a = `${columnIndexToLetters(o.firstCol)}${o.firstRow + 1}`,
      s = `${columnIndexToLetters(o.lastCol)}${o.lastRow + 1}`;
    return a === s ? a : `${a}:${s}`;
  });
  return [...new Set(t)];
}
function isCellAddressInRanges(e, r) {
  const t = _t(r);
  return e.some((n) => {
    const o = parseRangeAddress(n);
    return (
      t.row >= o.firstRow &&
      t.row <= o.lastRow &&
      t.col >= o.firstCol &&
      t.col <= o.lastCol
    );
  });
}
function countSelectionCells(e) {
  const r = parseSelectionAreas(e).map(parseRangeAddress),
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
function enumerateSelectionCells(e, r = workbookLimits.maxPatches) {
  if (!Number.isSafeInteger(r) || r < 1 || r > workbookLimits.maxCells)
    throw Error("Invalid selection enumeration limit.");
  const t = parseSelectionAreas(e),
    n = countSelectionCells(t);
  if (n > r)
    throw Error(
      `Select at most ${r.toLocaleString()} cells for this operation. The current selection contains ${n.toLocaleString()} cells.`,
    );
  const o = new Set();
  for (const a of t) {
    const s = parseRangeAddress(a);
    for (let d = s.firstRow; d <= s.lastRow; d++)
      for (let u = s.firstCol; u <= s.lastCol; u++)
        o.add(`${columnIndexToLetters(u)}${d + 1}`);
  }
  return [...o];
}
function pQ(e, r, t, n, o) {
  const a = parseRangeAddress(t),
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
    u = [],
    l = e.sheets.find((f) => f.id === n);
  return {
    watches: e.formulaWatches?.map((f) => {
      if (f.orphaned) return f;
      const p = `${f.anchor.sheetName}!${f.anchor.address}`;
      if (f.anchor.sheetId === r && d(a, f.anchor.address)) {
        const m = _t(f.anchor.address),
          v = `${columnIndexToLetters(m.col + s.firstCol - a.firstCol)}${m.row + s.firstRow - a.firstRow + 1}`,
          g = `${l.name}!${v}`;
        return p === g && r === n
          ? f
          : (u.push({
              kind: "formula-watch",
              label: f.id,
              before: p,
              after: g,
            }),
            {
              ...f,
              anchor: { sheetId: n, sheetName: l.name, address: v },
            });
      }
      return f.anchor.sheetId === n && d(s, f.anchor.address)
        ? (u.push({
            kind: "formula-watch",
            label: f.id,
            before: p,
            after: `${p} — watched source overwritten`,
          }),
          { ...f, orphaned: "overwritten" })
        : f;
    }),
    changes: u,
  };
}
function validateFreezePanes(e, r, t = { rows: r.rowCount, cols: r.colCount }) {
  validateSheetDefaultSizes(e, r);
  const n = validateSplitPaneState(e.splitPanes, t);
  if (n && (e.frozenRows || e.frozenCols))
    throw Error("Split panes and frozen panes cannot be enabled together.");
  n && (r.splitPanes = n);
  const o = normalizeHiddenIntervals(e.hiddenRows, t.rows),
    a = normalizeHiddenIntervals(e.hiddenCols, t.cols),
    s = { ...r, hiddenRows: o, hiddenCols: a },
    d = createAxisVisibilityIndex(r.rowCount, cn(s, "row")),
    u = createAxisVisibilityIndex(r.colCount, cn(s, "column"));
  if (!d.count || !u.count)
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
  const h = d.countBefore(l),
    f = u.countBefore(c);
  if (
    h > worksheetCapacityLimits.maxFrozenVisibleRows ||
    f > worksheetCapacityLimits.maxFrozenVisibleCols
  )
    throw Error(
      "Freeze panes supports at most 100 visible rows and 50 visible columns.",
    );
  if (h === d.count || f === u.count)
    throw Error("Freeze panes must leave an unfrozen visible row and column.");
  for (const p of r.merges ?? []) {
    const m = parseRangeAddress(p);
    if (
      (m.firstRow < l && m.lastRow >= l) ||
      (m.firstCol < c && m.lastCol >= c)
    )
      throw Error(
        "Freeze panes cannot cross a merged range. Unmerge those cells first.",
      );
  }
  (e.hiddenRows !== undefined && (r.hiddenRows = o),
    e.hiddenCols !== undefined && (r.hiddenCols = a),
    e.frozenRows !== undefined && (r.frozenRows = l),
    e.frozenCols !== undefined && (r.frozenCols = c));
}
function updateWorksheetById(e, r, t) {
  const n = e.sheets.find((a) => a.id === r);
  if (!n) throw Error("Worksheet not found.");
  assertSheetUnprotected(n);
  const o = t(structuredClone(n));
  return (
    validateFreezePanes(o, o),
    { ...e, sheets: e.sheets.map((a) => (a.id === r ? o : a)) }
  );
}
function setHiddenAxisRange(e, r, t, n, o, a) {
  if ((assertWorksheetAxis(t), typeof a != "boolean"))
    throw Error("Hidden state must be boolean.");
  return updateWorksheetById(e, r, (s) => {
    const d = axisCountOf(s, t);
    normalizeHiddenIntervals([{ start: n, end: o }], d);
    const u = readHiddenAxisIntervals(s, t),
      l = { start: n, end: o },
      c = a
        ? u.filter((f) =>
            f.end < l.start - 1 || f.start > l.end + 1
              ? true
              : ((l.start = Math.min(l.start, f.start)),
                (l.end = Math.max(l.end, f.end)),
                false),
          )
        : u,
      h = a
        ? [...c, l]
        : u.flatMap((f) =>
            f.end < n || f.start > o
              ? [f]
              : [
                  ...(f.start < n ? [{ start: f.start, end: n - 1 }] : []),
                  ...(f.end > o ? [{ start: o + 1, end: f.end }] : []),
                ],
          );
    return ((s[hiddenKeyForAxis(t)] = normalizeHiddenIntervals(h, d)), s);
  });
}
function yQ(e, r, t) {
  return (
    assertWorksheetAxis(t),
    updateWorksheetById(e, r, (n) => (delete n[hiddenKeyForAxis(t)], n))
  );
}
function setFreezePanes(e, r, t, n) {
  return updateWorksheetById(
    e,
    r,
    (o) => (delete o.splitPanes, { ...o, frozenRows: t, frozenCols: n }),
  );
}
function setSplitPanes(e, r, t) {
  return updateWorksheetById(
    e,
    r,
    (n) => (
      t === undefined
        ? delete n.splitPanes
        : ((n.splitPanes = validateSplitPaneState(t, {
            rows: n.rowCount,
            cols: n.colCount,
          })),
          (n.frozenRows = 0),
          (n.frozenCols = 0)),
      n
    ),
  );
}
function setFreezePanesAtCell(e, r, t) {
  const n = _t(t);
  return setFreezePanes(e, r, n.row, n.col);
}
function applyAxisStructureEdit(e, r, t, n, o) {
  assertWorksheetAxis(r);
  const a = axisCountOf(e, r);
  if (
    !Number.isInteger(t) ||
    t < 0 ||
    !Number.isInteger(n) ||
    n < 1 ||
    n > 1e4 ||
    t + n > a ||
    typeof o != "boolean"
  )
    throw Error("Choose a valid structural edit of 1–10,000 rows or columns.");
  const s = readHiddenAxisIntervals(e, r).flatMap((f) => {
      if (!o) {
        if (f.end < t) return [f];
        if (f.end + n >= a)
          throw Error(
            "Insertion would discard hidden-axis metadata at the worksheet boundary.",
          );
        return [
          { start: f.start >= t ? f.start + n : f.start, end: f.end + n },
        ];
      }
      const p = t + n - 1;
      return f.end < t
        ? [f]
        : f.start > p
          ? [{ start: f.start - n, end: f.end - n }]
          : [
              ...(f.start < t ? [{ start: f.start, end: t - 1 }] : []),
              ...(f.end > p ? [{ start: t, end: f.end - n }] : []),
            ];
    }),
    d = hiddenKeyForAxis(r),
    u = r === "row" ? "frozenRows" : "frozenCols",
    l = e[u] ?? 0,
    c = o ? l - Math.max(0, Math.min(n, l - t)) : l + (t < l ? n : 0),
    h = {
      ...e,
      [d]: normalizeHiddenIntervals(s, a),
      [u]: c,
      ...(e.outlineGroups
        ? { outlineGroups: remapOutlineGroupsForEdit(e, r, t, n, o) }
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
function setWorksheetReadOnly(e, r, t, n = {}) {
  if (n.readOnly !== undefined && typeof n.readOnly != "boolean")
    throw Error("Read-only state must be boolean.");
  if (n.readOnly) throw Error("This workbook is read-only.");
  const o = e.sheets.find((s) => s.id === r);
  if (!o) throw Error("Worksheet not found.");
  assertSheetUnprotected(o);
  const a = structuredClone(o);
  return (
    t(a),
    validateOutlineGroups(a, a),
    validateFreezePanes(a, a),
    JSON.stringify(a) === JSON.stringify(o)
      ? e
      : { ...e, sheets: e.sheets.map((s) => (s.id === r ? a : s)) }
  );
}
function EQ(e, r, t, n, o, a) {
  return (
    assertWorksheetAxis(t),
    setWorksheetReadOnly(
      e,
      r,
      (s) => {
        s.outlineGroups = [
          ...(s.outlineGroups ?? []),
          {
            id: crypto.randomUUID(),
            axis: t,
            start: n,
            end: o,
            collapsed: false,
          },
        ];
      },
      a,
    )
  );
}
function removeOutlineGroup(e, r, t, n) {
  return setWorksheetReadOnly(
    e,
    r,
    (o) => {
      if (!o.outlineGroups?.some((s) => s.id === t))
        throw Error("Outline group not found.");
      o.outlineGroups = o.outlineGroups.filter((s) => s.id !== t);
    },
    n,
  );
}
function setOutlineGroupCollapsed(e, r, t, n, o) {
  if (typeof n != "boolean") throw Error("Collapsed state must be boolean.");
  return setWorksheetReadOnly(
    e,
    r,
    (a) => {
      const s = a.outlineGroups?.find((u) => u.id === t);
      if (!s) throw Error("Outline group not found.");
      s.collapsed = n;
    },
    o,
  );
}
function setOutlineGroupLevel(e, r, t, n, o) {
  return (
    assertWorksheetAxis(t),
    setWorksheetReadOnly(
      e,
      r,
      (a) => {
        const s = addOutlineGroup(a, t),
          d = Math.max(0, ...s.map((c) => c.depth)) + 1;
        if (!Number.isInteger(n) || n < 1 || n > d)
          throw Error("Choose an available outline level.");
        const u = new Map(s.map((c) => [c.id, c.depth]));
        a.outlineGroups = a.outlineGroups?.map((c) =>
          c.axis === t ? { ...c, collapsed: u.get(c.id) >= n } : c,
        );
      },
      o,
    )
  );
}
function clearAxisOutlineGroups(e, r, t, n) {
  return (
    assertWorksheetAxis(t),
    setWorksheetReadOnly(
      e,
      r,
      (o) => {
        o.outlineGroups = o.outlineGroups?.filter((s) => s.axis !== t);
      },
      n,
    )
  );
}
function normalizeDataValidationRules(e) {
  if (!Array.isArray(e) || e.length > validationLimits.rulesPerSheet)
    throw Error("At most 100 validation rules are supported.");
  const r = new Set();
  return e.map((t) => {
    if (
      !isNonArrayObject(t) ||
      ![
        "list",
        "number",
        "integer",
        "textLength",
        "date",
        "time",
        "custom",
      ].includes(String(t.kind)) ||
      ![
        "between",
        "notBetween",
        "greater",
        "greaterEqual",
        "less",
        "lessEqual",
        "equal",
        "notEqual",
      ].includes(String(t.operator)) ||
      typeof t.allowBlank != "boolean" ||
      !Array.isArray(t.list) ||
      t.list.length > validationLimits.listItems ||
      typeof t.min != "number" ||
      !Number.isFinite(t.min) ||
      typeof t.max != "number" ||
      !Number.isFinite(t.max)
    )
      throw Error("Invalid data validation rule.");
    const n = assertBoundedValidationText(t.id, 100, "ID");
    if (!n.trim() || r.has(n))
      throw Error("Validation rule IDs must be nonempty and unique.");
    r.add(n);
    const o = parseRangeAddress(
        assertBoundedValidationText(t.range, 40, "range"),
      ),
      a = {
        id: n,
        range: `${columnIndexToLetters(o.firstCol)}${o.firstRow + 1}:${columnIndexToLetters(o.lastCol)}${o.lastRow + 1}`,
        kind: t.kind,
        operator: t.operator,
        min: t.min,
        max: t.max,
        list: t.list.map((s) =>
          assertBoundedValidationText(s, 1e3, "list item"),
        ),
        allowBlank: t.allowBlank,
        message: assertBoundedValidationText(t.message, 500, "error message"),
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
      t.formula !== undefined &&
      ((a.formula = assertBoundedValidationText(t.formula, 8192, "formula")),
      !a.formula.startsWith("=") || a.formula.length < 2)
    )
      throw Error("Start the custom validation formula with =.");
    if (a.kind === "custom" && !a.formula)
      throw Error(
        "A custom rule requires a formula relative to the range's top-left cell.",
      );
    if (t.errorStyle !== undefined) {
      if (!["stop", "warning", "information"].includes(String(t.errorStyle)))
        throw Error("Invalid validation error style.");
      a.errorStyle = t.errorStyle;
    }
    for (const s of ["showInputMessage", "showDropdown"])
      if (t[s] !== undefined) {
        if (typeof t[s] != "boolean")
          throw Error("Invalid validation display setting.");
        a[s] = t[s];
      }
    for (const [s, d] of [
      ["inputTitle", 100],
      ["inputMessage", 500],
      ["errorTitle", 100],
    ])
      t[s] !== undefined && (a[s] = assertBoundedValidationText(t[s], d, s));
    return a;
  });
}
function findDataValidationsInRange(e, r, t) {
  const n = _t(t);
  return (e.sheets.find((a) => a.id === r)?.validations ?? []).filter((a) => {
    const s = parseRangeAddress(a.range);
    return (
      n.row >= s.firstRow &&
      n.row <= s.lastRow &&
      n.col >= s.firstCol &&
      n.col <= s.lastCol
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
    const a = parseRangeAddress(
        assertBoundedString(n.range, 40, "conditional range"),
      ),
      s = {
        id: o,
        range: `${columnIndexToLetters(a.firstCol)}${a.firstRow + 1}:${columnIndexToLetters(a.lastCol)}${a.lastRow + 1}`,
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
          (u) => typeof u != "number" || !Number.isFinite(u),
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
function evaluateConditionalFormatRules(e, r, t) {
  const n = new Map(),
    o = new Map();
  let a = 0,
    s = 0;
  const d = (f, p, m) => {
    const v = { ruleId: f, address: p, message: m };
    return (n.set(`${f}:${m}`, v), v);
  };
  let u = "";
  if (
    !t &&
    (e.conditionalFormats ?? []).some((f) =>
      ["duplicate", "unique", "colorScale", "dataBar", "iconSet"].includes(
        f.kind,
      ),
    )
  )
    try {
      t = vx(e, r);
    } catch (f) {
      u =
        f instanceof Error
          ? f.message
          : "Array values could not be discovered for conditional statistics.";
    }
  const l = Object.keys(t ?? e.cells).map((f) => ({
      address: f,
      point: _t(f),
    })),
    c = sortConditionalRulesByPriority(e.conditionalFormats ?? []).map((f) => {
      const p = parseRangeAddress(f.range),
        m = new Map();
      let v = 1 / 0,
        g = -1 / 0,
        b = "";
      if (
        (["greater", "less", "between"].includes(f.kind) &&
          (!f.value.trim() ||
            !Number.isFinite(Number(f.value)) ||
            (f.kind === "between" &&
              (!f.second.trim() ||
                !Number.isFinite(Number(f.second)) ||
                Number(f.value) > Number(f.second)))) &&
          (b =
            "Numeric rule operands must be finite and minimum cannot exceed maximum."),
        ["duplicate", "unique", "colorScale", "dataBar", "iconSet"].includes(
          f.kind,
        ))
      ) {
        u && (b = u);
        for (const { address, point } of l) {
          if (++a > 1e6) {
            b =
              "Conditional statistics exceed 1,000,000 inspected cells; narrow the rule ranges or remove rules.";
            break;
          }
          if (!isCellInsideRange(p, point)) continue;
          const E = r(e.id, address);
          if (
            (typeof E == "number" &&
              Number.isFinite(E) &&
              ((v = Math.min(v, E)), (g = Math.max(g, E))),
            E !== null && E !== "")
          ) {
            const A = String(E).toLowerCase();
            m.set(A, (m.get(A) ?? 0) + 1);
          }
        }
      }
      return (
        b && d(f.id, "", b),
        { rule: f, bounds: p, counts: m, min: v, max: g, failure: b }
      );
    }),
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
          } catch (O) {
            m.errors.push(
              d(
                rule.id,
                f,
                O instanceof Error ? O.message : "Formula evaluation failed.",
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
          } catch (I) {
            m.errors.push(
              d(
                rule.id,
                f,
                I instanceof Error
                  ? I.message
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
                V = (-O / z) * 100,
                D = ((v - O) / z) * 100;
              ((m.dataBar = {
                start: Math.min(V, D),
                width: Math.abs(D - V),
                axis: V,
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
    const r = parseRangeAddress(e);
    return `${columnIndexToLetters(r.firstCol)}${r.firstRow + 1}:${columnIndexToLetters(r.lastCol)}${r.lastRow + 1}`;
  },
  doesRangeOverlap = (e, r) => {
    const t = parseRangeAddress(e),
      n = parseRangeAddress(r);
    return (
      t.firstRow <= n.lastRow &&
      t.lastRow >= n.firstRow &&
      t.firstCol <= n.lastCol &&
      t.lastCol >= n.firstCol
    );
  },
  lD = (e, r) => {
    const t = parseRangeAddress(e),
      n = _t(r);
    return (
      n.row >= t.firstRow &&
      n.row <= t.lastRow &&
      n.col >= t.firstCol &&
      n.col <= t.lastCol
    );
  };
function validateWorksheetFeatures(e, r) {
  if (isNonArrayObject(e)) {
    if (e.merges !== undefined) {
      (!Array.isArray(e.merges) || e.merges.length > 1e3) &&
        failWithMessage("At most1,000 merged ranges are supported.");
      const t = [];
      for (const n of e.merges) {
        const o = CE(normalizeWorksheetToolText(n, 40)),
          a = parseRangeAddress(o);
        ((a.lastRow - a.firstRow + 1) * (a.lastCol - a.firstCol + 1) > 500 &&
          failWithMessage("A merged area supports at most 500 cells."),
          t.some((s) => doesRangeOverlap(s, o)) &&
            failWithMessage("Merged ranges cannot overlap."),
          t.push(o));
      }
      r.merges = t;
    }
    if (
      (e.validations !== undefined &&
        (r.validations = normalizeDataValidationRules(e.validations)),
      e.conditionalFormats !== undefined &&
        (r.conditionalFormats = normalizeConditionalRules(
          e.conditionalFormats,
        )),
      e.protection !== undefined)
    ) {
      const t = e.protection;
      ((!isNonArrayObject(t) ||
        typeof t.enabled != "boolean" ||
        (t.hash !== undefined &&
          (!/^[a-f\d]{64}$/i.test(String(t.hash)) ||
            typeof t.salt != "string" ||
            t.salt.length > 100))) &&
        failWithMessage("Invalid sheet protection."),
        (r.protection = {
          enabled: t.enabled,
          ...(t.hash ? { hash: String(t.hash), salt: String(t.salt) } : {}),
        }));
    }
  }
}
function mergeWorksheetCells(e, r, t, n = false) {
  const o = e.sheets.find((h) => h.id === r);
  assertSheetUnprotected(o);
  const a = CE(t),
    s = parseRangeAddress(a);
  (s.lastRow - s.firstRow + 1) * (s.lastCol - s.firstCol + 1) > 500 &&
    failWithMessage("Merge at most 500 cells.");
  const d = (o.merges ?? []).filter((h) => doesRangeOverlap(h, a));
  !n && d.length && failWithMessage("Unmerge overlapping areas first.");
  const u = { ...o.cells };
  if (!n) {
    const h = `${columnIndexToLetters(s.firstCol)}${s.firstRow + 1}`;
    u[h] = {
      ...u[h],
      raw: u[h]?.raw ?? null,
      style: { ...u[h]?.style, align: "center" },
    };
  }
  return {
    ...e,
    sheets: e.sheets.map((h) =>
      h.id === r
        ? {
            ...h,
            cells: u,
            merges: n
              ? (h.merges ?? []).filter((f) => !d.includes(f))
              : [...(h.merges ?? []), a],
          }
        : h,
    ),
  };
}
function oP(e) {
  let r = 0;
  for (const { refersTo } of e) {
    const n = stripFormulaStringLiterals(
      refersTo.replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')+'!/g, ""),
    );
    for (const o of n.matchAll(
      /(?:^|[^A-Za-z0-9_.])((?:\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?)(?![A-Za-z0-9_.])/g,
    )) {
      const a = parseRangeAddress(o[1]);
      if (
        ((r += (a.lastRow - a.firstRow + 1) * (a.lastCol - a.firstCol + 1)),
        r > workbookLimits.maxCells)
      )
        throw Error(
          "Named ranges exceed the 100,000-cell XLSX import/export processing budget. Use native JSON for larger definitions.",
        );
    }
  }
}
function aP(e) {
  const r = parseDefinedNamesXml(e),
    t = { names: [], warnings: [], builtinXml: "" };
  if (!r) return t;
  let n = r.body,
    o = 0;
  const a = new Set(),
    s = [...e.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b[^>]*>/g)].length;
  for (; n.trim();) {
    const d =
      /^\s*<((?:[A-Za-z_][\w.-]*:)?definedName)\b((?:[^"'<>]|"[^"]*"|'[^']*')*?)(?:\/>|>([^<]*)<\/\1\s*>)/.exec(
        n,
      );
    if (!d)
      return throwXlsxDefinedNameError(
        "Malformed definedName content; only text definitions are supported.",
      );
    const u = parseDefinedNameXmlAttributes(d[2]),
      l = u.name ?? "",
      c = u.localSheetId,
      h = d[1].slice(0, -11);
    if (
      { ...r.namespaces, ...u }[h ? `xmlns:${h.slice(0, -1)}` : "xmlns"] !==
      spreadsheetMlNamespace
    )
      return throwXlsxDefinedNameError(
        "Unsupported or undeclared definedName namespace.",
      );
    for (const v of Object.keys(u))
      if (v.startsWith("xmlns") && u[v] !== spreadsheetMlNamespace)
        return throwXlsxDefinedNameError("Unsupported definedName namespace.");
    if (
      c !== undefined &&
      (!/^\d+$/.test(c) || !Number.isSafeInteger(Number(c)))
    )
      return throwXlsxDefinedNameError(
        `Name "${l}" has an invalid localSheetId.`,
      );
    const p = `=${nP(
      (d[3] ?? "").replace(
        /\r\n?/g,
        `
`,
      ),
    ).replace(/^=/, "")}`;
    if ((validateNameFormulaText(p), c !== undefined && Number(c) >= s))
      return throwXlsxDefinedNameError(
        `Name "${l}" has an out-of-range localSheetId.`,
      );
    const m = `${c === undefined ? "workbook" : Number(c)}:${l.toLowerCase()}`;
    if (a.has(m))
      return throwXlsxDefinedNameError(
        `Duplicate name "${l}" in the same scope.`,
      );
    if ((a.add(m), reservedDefinedNames.has(l.toLowerCase())))
      t.builtinXml += `<definedName${Object.entries(u)
        .filter(([v]) => !v.startsWith("xmlns"))
        .map(([v, g]) => ` ${v}="${validateXmlSafeName(g, true)}"`)
        .join("")}>${validateXmlSafeName(p.slice(1))}</definedName>`;
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
        if (u[g] !== undefined && !["0", "false"].includes(u[g]))
          return throwXlsxDefinedNameError(
            `Name "${l}" has an executable ${g} attribute.`,
          );
      const v = Object.keys(u).filter(
        (g) =>
          !["name", "localSheetId", "comment"].includes(g) &&
          !g.startsWith("xmlns"),
      );
      (v.length &&
        t.warnings.push(
          `Defined name "${l}" retained its definition and scope; unsupported attributes (${v.join(", ")}) were omitted. Hidden names are visible in Name Manager.`,
        ),
        t.names.push({
          name: l,
          refersTo: p,
          ...(c !== undefined ? { localSheetId: Number(c) } : {}),
          ...(u.comment !== undefined ? { comment: u.comment } : {}),
        }));
    }
    n = n.slice(d[0].length);
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
  } catch (d) {
    return throwXlsxDefinedNameError(d.message);
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
