import { X, ArrowRight } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { Mt } from "./part-02.js";
import {
  $B,
  Af,
  Ay,
  Ip,
  Nr,
  PB,
  Q3,
  SE,
  UB,
  Uv,
  Vn,
  W3,
  _t,
  cn,
  ct,
  ff,
  gg,
  hi,
  jn,
  l0,
  nee,
  pQ,
  rn,
  ul,
  ut,
  v0,
  vg,
  y1,
} from "./part-03.js";
import { Dr, M, ci } from "./part-04.js";
import {
  $W,
  D1,
  E2,
  HW,
  Hh,
  PW,
  R3,
  UW,
  VW,
  WW,
  bp,
  gj,
  k4e,
  mj,
  pp,
  w_,
  x4e,
  y_,
  zW,
  zn,
} from "./part-47.js";

function tN(e, r, t) {
  if (r !== e.source.id) return { sheetId: r, reference: t, changed: false };
  const n = HW(t);
  if (!bp(n, e.from)) return { sheetId: r, reference: t, changed: false };
  if (gj(e.from, n))
    return {
      sheetId: e.target.id,
      reference: k4e(t, e.dr, e.dc),
      changed: true,
    };
  const o = {
      firstRow: Math.max(n.firstRow, e.from.firstRow),
      lastRow: Math.min(n.lastRow, e.from.lastRow),
      firstCol: Math.max(n.firstCol, e.from.firstCol),
      lastCol: Math.min(n.lastCol, e.from.lastCol),
    },
    a = {
      firstRow: o.firstRow + e.dr,
      lastRow: o.lastRow + e.dr,
      firstCol: o.firstCol + e.dc,
      lastCol: o.lastCol + e.dc,
    };
  if (e.source.id === e.target.id && gj(n, a))
    return { sheetId: r, reference: t, changed: false };
  zn(
    `Reference ${e.source.name}!${t} includes cells outside the moved block. This move would require a split or cross-sheet range; use explicit references first.`,
  );
}
function qW(e, r, t = new Set()) {
  const n = `${r.scopeSheetId ?? ""}:${r.name.toLowerCase()}`;
  return t.has(n) || t.size > 32
    ? false
    : (t.add(n),
      w_(r.refersTo).every((o) => {
        if (o.kind === "dynamic" || y_(o)) return false;
        if (o.kind === "reference") return !!o.sheet;
        if (o.kind === "name" && !o.sheet) {
          const a = (e.definedNames ?? []).find(
            (s) =>
              !s.scopeSheetId && s.name.toLowerCase() === o.name.toLowerCase(),
          );
          return !a || qW(e, a, new Set(t));
        }
        return true;
      }));
}
function XW(e, r, t = new Set()) {
  const n = `${r.scopeSheetId ?? ""}:${r.name.toLowerCase()}`;
  return t.has(n) || t.size > 32
    ? true
    : (t.add(n),
      w_(r.refersTo).some((o) => {
        if (y_(o)) return true;
        if (o.kind !== "name" || o.sheet) return false;
        const a =
          (e.definedNames ?? []).find(
            (s) =>
              s.scopeSheetId === r.scopeSheetId &&
              s.name.toLowerCase() === o.name.toLowerCase(),
          ) ??
          (e.definedNames ?? []).find(
            (s) =>
              !s.scopeSheetId && s.name.toLowerCase() === o.name.toLowerCase(),
          );
        return !!a && XW(e, a, new Set(t));
      }));
}
function yS(e, r, t, n, o = false) {
  var u;
  ((e.characters += r.length),
    e.characters > D1.formulaCharacters &&
      zn(
        "Reference repair exceeds 10 million formula characters. Simplify the workbook before moving cells.",
      ));
  const a = w_(r);
  let s = "",
    d = 0;
  for (const l of a) {
    let c = l.text;
    if (
      (l.kind === "dynamic" &&
        zn(
          `${l.name} can refer to cells dynamically. Replace dynamic references with explicit A1 references before moving cells.`,
        ),
      y_(l) &&
        o &&
        zn(
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
        zn(
          `Name ${l.name} contains a current-row structured reference. Use an explicit qualified cell reference before moving this formula.`,
        ),
        t !== n &&
          (f || m || (p && !qW(e.book, p))) &&
          zn(
            `Name ${l.name} depends on worksheet scope. Keep the formula on its sheet or replace that name with explicit qualified references before moving it.`,
          ));
    }
    if (l.kind === "reference") {
      const h = l.sheet
        ? (u = e.book.sheets.find(
            (f) => f.name.toLowerCase() === l.sheet.toLowerCase(),
          )) == null
          ? void 0
          : u.id
        : t;
      if (!h)
        !l.sheet &&
          bp(HW(l.reference), e.from) &&
          zn(
            "A workbook-level name has a context-dependent unqualified reference. Qualify its worksheet before moving cells.",
          );
      else {
        const f = tN(e, h, l.reference);
        f.changed
          ? (c = `${l.sheet || f.sheetId !== n ? mj(Nr(e.book, f.sheetId).name) + "!" : ""}${f.reference}`)
          : !l.sheet &&
            t !== n &&
            t &&
            (c = `${mj(Nr(e.book, t).name)}!${l.reference}`);
      }
    }
    c !== l.text && ((s += r.slice(d, l.start) + c), (d = l.end));
  }
  return s + r.slice(d);
}
function S4e(e, r) {
  const t = _t(r.anchor),
    n = (s, d, u, l) => {
      const c = ul(s === "row" ? e.rowCount : e.colCount, cn(e, s)),
        h = (p) => (s === "row" ? gg(e, p) : Ip(e, p));
      if (c.isHidden(d)) return [d, d];
      const f = (p) => {
        let m = d,
          v = 0;
        if (p >= 0)
          for (; p >= h(m);) {
            p -= h(m);
            const g = c.next(m, 1);
            if (g === null) return m;
            ((m = g),
              ++v > 1e3 &&
                zn("Drawing footprint exceeds the move inspection budget."));
          }
        else
          for (; p < 0;) {
            const g = c.next(m, -1);
            if (g === null) return 0;
            ((m = g),
              (p += h(m)),
              ++v > 1e3 &&
                zn("Drawing footprint exceeds the move inspection budget."));
          }
        return m;
      };
      return [f(u), f(u + l - 1e-4)];
    },
    o = n("row", t.row, r.offsetY, r.height),
    a = n("column", t.col, r.offsetX, r.width);
  return { firstRow: o[0], lastRow: o[1], firstCol: a[0], lastCol: a[1] };
}
function T4e(e) {
  const r = GW(e.book);
  for (const [t, n] of [
    [e.source, e.from],
    [e.target, e.to],
  ]) {
    for (const o of t.merges ?? [])
      bp(n, ct(o)) &&
        zn("Unmerge the source and destination before moving cells.");
    for (const o of t.tables ?? [])
      bp(n, ct(o.range)) &&
        zn(
          `The move intersects table ${o.name}. Table-cell moves are not supported; use ordinary cells outside tables.`,
        );
    for (const o of [...(t.pivots ?? []), ...(t.queries ?? [])])
      bp(n, ct(o.outputRange)) &&
        zn(
          "The move intersects a generated PivotTable or query output. Move ordinary cells instead.",
        );
    for (const o of [...(t.validations ?? []), ...(t.conditionalFormats ?? [])])
      bp(n, ct(o.range)) &&
        zn(
          "The move intersects a validation or conditional-formatting rule. Remove or relocate the rule first.",
        );
    for (const o of t.objects ?? [])
      (pp(n, _t(o.anchor)) || bp(n, S4e(t, o))) &&
        zn(
          `The move intersects drawing ${o.label || o.id}. Drawing footprints cannot be moved or partially overwritten by this command.`,
        );
    for (const o of r)
      o.sheetId === t.id &&
        bp(n, ct(o.range)) &&
        zn(
          "The move intersects a dynamic-array footprint. Convert it to ordinary values or move cells outside the array.",
        );
  }
}
function GW(e) {
  const r = Vn(e),
    t = [];
  let n = 0,
    o = 0;
  for (const a of e.sheets)
    for (const [s, d] of Object.entries(a.cells))
      if (PW(d)) {
        (++n > D1.formulaCells || (o += d.raw.length) > D1.formulaCharacters) &&
          zn(
            "Array inspection supports at most 10,000 formulas and 10 million formula characters. Simplify the workbook before moving cells.",
          );
        const u = r.getSpillRange(a.id, s);
        if (u) {
          const l = ct(u);
          (l.firstRow !== l.lastRow || l.firstCol !== l.lastCol) &&
            t.push({ sheetId: a.id, anchor: s, range: u });
        }
      }
  return t;
}
function yj(e) {
  return GW(e)
    .map((r) => `${r.sheetId}:${r.anchor}:${r.range}`)
    .sort().join(`
`);
}
function ZW(e, r, t, n, o, a = {}) {
  $W(a.readOnly);
  const s = Nr(e, r),
    d = Nr(e, n);
  (jn(s), jn(d));
  const u = VW(s, t),
    l = zW(d, o),
    c = u.lastRow - u.firstRow + 1,
    h = u.lastCol - u.firstCol + 1;
  (l.lastRow !== l.firstRow || l.lastCol !== l.firstCol) &&
    (l.lastRow - l.firstRow + 1 !== c || l.lastCol - l.firstCol + 1 !== h) &&
    zn(
      "Choose one destination cell or a rectangle with exactly the source dimensions.",
    );
  const f = {
    firstRow: l.firstRow,
    lastRow: l.firstRow + c - 1,
    firstCol: l.firstCol,
    lastCol: l.firstCol + h - 1,
  };
  (f.lastRow >= d.rowCount || f.lastCol >= d.colCount) &&
    zn("The moved cells would exceed the worksheet boundary.");
  const p = a.sourceSnapshot ?? WW(e, r, t);
  x4e(e, p, r, t);
  const m = {
      book: e,
      source: s,
      target: d,
      from: u,
      to: f,
      dr: f.firstRow - u.firstRow,
      dc: f.firstCol - u.firstCol,
      characters: 0,
    },
    v = R3(t),
    g = R3(`${Hh(f.firstRow, f.firstCol)}:${Hh(f.lastRow, f.lastCol)}`),
    b = [],
    y = r === n && m.dr === 0 && m.dc === 0;
  y || T4e(m);
  const w = y ? e : structuredClone(e),
    E = new Map();
  let A = 0,
    k = 0;
  const F = (B) => {
    (b.push(B),
      b.length > D1.metadataChanges &&
        zn(
          "The move would change more than 500 metadata references. Narrow the source.",
        ));
  };
  if (
    (!y &&
      e.sheets.some((B) => {
        var _;
        return (_ = B.conditionalFormats) == null
          ? void 0
          : _.some((C) => C.kind === "formula");
      }) &&
      zn(
        "Range moves cannot repair formula-based conditional rules yet. Remove those rules before moving cells.",
      ),
    !y)
  ) {
    const B = pQ(w, r, v, n, g);
    B.watches && (w.formulaWatches = B.watches);
    for (const N of B.changes) F(N);
    for (const N of e.reviewThreads ?? [])
      !N.orphaned &&
        N.anchor.sheetId === n &&
        pp(f, _t(N.anchor.address)) &&
        !(N.anchor.sheetId === r && pp(u, _t(N.anchor.address))) &&
        zn(
          "The destination contains an existing review discussion. Choose a destination without anchored discussions to preserve its evidence context.",
        );
    const _ = Nr(w, r),
      C = Nr(w, n);
    for (let N = u.firstRow; N <= u.lastRow; N++)
      for (let R = u.firstCol; R <= u.lastCol; R++) {
        const L = Hh(N, R);
        E2(s.cells[L]) && delete _.cells[L];
      }
    for (let N = u.firstRow; N <= u.lastRow; N++)
      for (let R = u.firstCol; R <= u.lastCol; R++) {
        const L = Hh(N, R),
          $ = Hh(N + m.dr, R + m.dc),
          U = s.cells[L],
          G = d.cells[$];
        (!(r === n && pp(u, _t($))) &&
          E2(G) &&
          JSON.stringify(G) !== JSON.stringify(U) &&
          A++,
          E2(U) ? (C.cells[$] = hi(U)) : E2(C.cells[$]) && delete C.cells[$],
          E.set(`${n}:${$}`, L));
      }
    let T = 0;
    for (const N of w.sheets)
      for (const [R, L] of Object.entries(N.cells)) {
        ++T > Mt.maxCells &&
          zn("Range moves inspect at most 100,000 populated cells.");
        const $ = E.has(`${N.id}:${R}`);
        if (PW(L)) {
          const G = yS(m, L.raw, $ ? r : N.id, N.id, $);
          G !== L.raw && ((L.raw = G), k++);
        }
        const U = L.hyperlink;
        if (
          (U == null ? void 0 : U.kind) === "internal" &&
          U.sheetId === r &&
          pp(u, _t(U.address))
        ) {
          const G = _t(U.address),
            H = `${s.name}!${U.address}`;
          ((L.hyperlink = {
            ...U,
            sheetId: n,
            address: Hh(G.row + m.dr, G.col + m.dc),
          }),
            JSON.stringify(U) !== JSON.stringify(L.hyperlink) &&
              (k++,
              F({
                kind: "internal-link",
                label: `${N.name}!${R}`,
                before: H,
                after: `${d.name}!${L.hyperlink.address}`,
              })));
        }
      }
    for (const N of w.sheets)
      for (const R of N.tables ?? [])
        for (const L of R.columns)
          if (L.calculatedFormula) {
            const $ = yS(m, L.calculatedFormula, N.id, N.id);
            $ !== L.calculatedFormula &&
              (jn(Nr(e, N.id)),
              F({
                kind: "table-column",
                label: `${N.name}: ${R.name}[${L.name}] calculated rule`,
                before: L.calculatedFormula,
                after: $,
              }),
              (L.calculatedFormula = $),
              k++);
          }
    for (const N of w.definedNames ?? []) {
      const R = yS(m, N.refersTo, N.scopeSheetId, N.scopeSheetId);
      R !== N.refersTo &&
        (N.scopeSheetId && jn(Nr(e, N.scopeSheetId)),
        F({
          kind: "defined-name",
          label: N.name,
          before: N.refersTo,
          after: R,
        }),
        (N.refersTo = R),
        k++);
    }
    for (const N of w.reviewThreads ?? [])
      if (
        !N.orphaned &&
        N.anchor.sheetId === r &&
        pp(u, _t(N.anchor.address))
      ) {
        const R = _t(N.anchor.address),
          L = `${N.anchor.sheetName}!${N.anchor.address}`;
        ((N.anchor = {
          sheetId: n,
          sheetName: d.name,
          address: Hh(R.row + m.dr, R.col + m.dc),
        }),
          F({
            kind: "review-anchor",
            label: N.id,
            before: L,
            after: `${d.name}!${N.anchor.address}`,
          }));
      }
    for (const N of w.sheets) {
      for (const R of [...(N.pivots ?? []), ...(N.queries ?? [])]) {
        const L = tN(m, R.source.sheetId, R.source.range);
        L.changed &&
          (jn(Nr(e, N.id)),
          F({
            kind: "analysis-source",
            label: `${N.name}: ${R.name}`,
            before: `${Nr(e, R.source.sheetId).name}!${R.source.range}`,
            after: `${Nr(e, L.sheetId).name}!${L.reference}`,
          }),
          (R.source = { ...R.source, sheetId: L.sheetId, range: L.reference }),
          k++);
      }
      for (const R of N.objects ?? [])
        if (R.kind === "chart" && R.range) {
          const L = tN(m, N.id, R.range);
          L.changed &&
            (L.sheetId !== N.id &&
              zn(
                "A chart source cannot follow data to another worksheet in this implementation. Move the chart data on its current sheet.",
              ),
            jn(Nr(e, N.id)),
            F({
              kind: "chart-source",
              label: R.label || R.id,
              before: R.range,
              after: L.reference,
            }),
            (R.range = L.reference),
            k++);
        }
    }
  }
  const I = [];
  let O = 0;
  if (!y)
    for (const B of w.sheets) {
      const _ = Nr(e, B.id),
        C = new Set([...Object.keys(_.cells), ...Object.keys(B.cells)]);
      for (const T of C) {
        const N = _.cells[T],
          R = B.cells[T];
        JSON.stringify(N) !== JSON.stringify(R) &&
          (jn(_),
          I.push({
            sheetId: B.id,
            sheetName: B.name,
            address: T,
            ...(N ? { before: hi(N) } : {}),
            ...(R ? { after: hi(R) } : {}),
          }),
          I.length > D1.changedCells &&
            zn(
              "A move may change at most 500 cells in total, including source clears, destination cells, and dependent references. Move a smaller block.",
            ),
          !(B.id === r && pp(u, _t(T))) &&
            !(B.id === n && pp(f, _t(T))) &&
            O++);
      }
    }
  const j = I.length > 0 || b.length > 0;
  j &&
    (yj(e) !== yj(w) &&
      zn(
        "The move would create, remove, or resize a dynamic-array footprint. Convert the affected array to ordinary values before moving cells.",
      ),
    (e.events.length >= 2e3 ||
      e.events.reduce((B, _) => B + _.changes.length, 0) + I.length >
        Mt.maxCells) &&
      zn(
        "The move exceeds workbook history limits. Export the workbook before starting a new workpaper.",
      ),
    e.revision >= Number.MAX_SAFE_INTEGER - 1 &&
      zn("Workbook revision limit reached."),
    v0(
      w,
      I.filter((B) => {
        var _, C, T, N;
        return (
          ((_ = B.before) == null ? void 0 : _.raw) !==
            ((C = B.after) == null ? void 0 : C.raw) ||
          ((T = B.before) == null ? void 0 : T.literal) !==
            ((N = B.after) == null ? void 0 : N.literal)
        );
      }),
    ));
  const z = j ? rn(w) : e,
    V = I.map((B) => ({
      sheetId: B.sheetId,
      address: B.address,
      cell: hi(B.after ?? { raw: null }),
    })),
    D = {
      workbookId: e.id,
      revision: e.revision,
      snapshotGuard: p.snapshotGuard,
      sourceSheetId: r,
      sourceRange: v,
      sourceSnapshot: { ...p },
      targetSheetId: n,
      targetRange: g,
      targetCount: c * h,
      patches: V,
      previewPatches: V,
      count: I.length,
      dependentCount: O,
      overwrittenCells: A,
      referenceChanges: k,
      metadataChanges: b,
      previewChanges: I,
      summary: `Move ${s.name}!${v} to ${d.name}!${g}: ${I.length} cell changes, ${A} destination cells overwritten, ${b.length} metadata changes.`,
    };
  return {
    plan: UW(D),
    next: j ? { ...z, revision: e.revision, events: e.events } : e,
  };
}
function N4e(e, r, t, n, o, a = {}) {
  return ZW(e, r, t, n, o, a).plan;
}
function R4e(e, r, t = {}) {
  ($W(t.readOnly),
    (!r ||
      r.workbookId !== e.id ||
      r.revision !== e.revision ||
      r.snapshotGuard !== JSON.stringify(e)) &&
      zn(
        "The workbook changed after this move preview. Cut the source again before moving it.",
      ));
  const n = ZW(
    e,
    r.sourceSheetId,
    r.sourceRange,
    r.targetSheetId,
    r.targetRange,
    { sourceSnapshot: r.sourceSnapshot, readOnly: t.readOnly },
  );
  return (
    (JSON.stringify(n.plan.previewChanges) !==
      JSON.stringify(r.previewChanges) ||
      JSON.stringify(n.plan.metadataChanges) !==
        JSON.stringify(r.metadataChanges)) &&
      zn("The move no longer matches its preview. Cut the source again."),
    n.next
  );
}
function wj(e) {
  if (!e) return "(empty)";
  const r = [String(e.raw ?? "(empty)")];
  return (
    e.literal && r.push("Literal text"),
    e.note !== void 0 && r.push("Note: " + e.note),
    e.tickmark !== void 0 && r.push("Tickmark: " + e.tickmark),
    e.hyperlink &&
      r.push(
        "Link: " +
          (e.hyperlink.kind === "external"
            ? e.hyperlink.url
            : e.hyperlink.sheetId + "!" + e.hyperlink.address) +
          (e.hyperlink.label ? " (" + e.hyperlink.label + ")" : ""),
      ),
    e.numberFormat
      ? r.push(
          "Number display: " +
            Object.entries(e.numberFormat)
              .map(([t, n]) => t + " = " + n)
              .join(", "),
        )
      : (e.numFmt || e.format) &&
        r.push("Number display: " + (e.numFmt ?? e.format)),
    e.style &&
      Object.keys(e.style).length &&
      r.push(
        "Cell style: " +
          Object.entries(e.style)
            .map(([t, n]) => t + " = " + n)
            .join(", "),
      ),
    r.join(" · ")
  );
}
const xj =
  "The workbook or destination changed. Cancel and cut the current cells again.";
function I4e(e) {
  var b, y, w, E, A;
  const { text: r, language: t } = Dr(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e);
  a.current = e;
  const s = ne.useRef({
      sheetId: e.sheetId,
      range: e.range,
      source: e.source,
      sourceName:
        ((b = e.workbook.sheets.find((k) => k.id === e.source.sheetId)) == null
          ? void 0
          : b.name) ?? "Deleted worksheet",
      targetName:
        ((y = e.workbook.sheets.find((k) => k.id === e.sheetId)) == null
          ? void 0
          : y.name) ?? "Deleted worksheet",
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
          plan: N4e(
            e.workbook,
            e.source.sheetId,
            e.source.range,
            e.sheetId,
            e.range,
            { sourceSnapshot: e.source, readOnly: e.readOnly },
          ),
          error: "",
        };
      } catch (k) {
        return {
          plan: null,
          error: k instanceof Error ? k.message : "The cells cannot be moved.",
        };
      }
    }),
    p = [
      ...(((w = f.plan) == null ? void 0 : w.previewChanges) ?? []).map(
        (k) => ({
          label: `${k.sheetName}!${k.address}`,
          before: wj(k.before),
          after: wj(k.after),
        }),
      ),
      ...(((E = f.plan) == null ? void 0 : E.metadataChanges) ?? []).map(
        (k) => ({ label: k.label, before: k.before, after: k.after }),
      ),
    ],
    m = Math.max(1, Math.ceil(p.length / 15)),
    v = Math.min(l, m - 1);
  ne.useEffect(() => {
    var F, I;
    const k = document.activeElement;
    return (
      (I =
        (F = o.current) == null
          ? void 0
          : F.querySelector("button:not(:disabled)")) == null || I.focus(),
      () => {
        k != null && k.isConnected && k.focus();
      }
    );
  }, []);
  function g() {
    try {
      const k = a.current;
      if (k.readOnly) throw Error("This workbook is read-only.");
      if (
        s.current.book !== JSON.stringify(k.workbook) ||
        s.current.sheetId !== k.sheetId ||
        s.current.range !== k.range ||
        s.current.source !== k.source
      )
        throw Error(xj);
      if (!f.plan) throw Error(f.error);
      const F = R4e(k.workbook, f.plan, { readOnly: k.readOnly });
      (k.onWorkbookChange(F, "Move cells and repair references"),
        k.onSelectionChange(f.plan.targetRange),
        k.onApplied(),
        k.onClose());
    } catch (k) {
      u(k instanceof Error ? k.message : "The cells could not be moved.");
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
      onKeyDown: (k) => {
        var F;
        if (
          (k.key === "Escape" &&
            (k.preventDefault(), k.stopPropagation(), a.current.onClose()),
          k.key === "Tab")
        ) {
          const I = Array.from(
              ((F = o.current) == null
                ? void 0
                : F.querySelectorAll(
                    'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
                  )) ?? [],
            ),
            O = I[0],
            j = I.at(-1);
          k.shiftKey && document.activeElement === O
            ? (k.preventDefault(), j == null || j.focus())
            : !k.shiftKey &&
              document.activeElement === j &&
              (k.preventDefault(), O == null || O.focus());
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
                i.jsx(M, { children: "Move cells" }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": r("Close move preview"),
              onClick: e.onClose,
              lang: t("Close move preview"),
              "data-ui-label": "Close move preview",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "et-body",
          children: [
            i.jsx("p", {
              children: i.jsx(M, {
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
                    ((A = f.plan) == null ? void 0 : A.targetRange) ??
                      s.current.range,
                  ],
                }),
              ],
            }),
            f.plan &&
              i.jsxs("section", {
                className: "et-preview",
                "aria-label": r("Move preview"),
                lang: t("Move preview"),
                "data-ui-label": "Move preview",
                children: [
                  i.jsxs("p", {
                    children: [
                      f.plan.count,
                      " ",
                      i.jsx(M, { children: "cell changes ·" }),
                      " ",
                      f.plan.overwrittenCells,
                      " ",
                      i.jsx(M, {
                        children: "occupied destinations overwritten · ",
                      }),
                      f.plan.referenceChanges,
                      " ",
                      i.jsx(M, { children: "dependent references updated · " }),
                      f.plan.metadataChanges.length,
                      " ",
                      i.jsx(M, { children: "metadata changes" }),
                    ],
                  }),
                  i.jsxs("table", {
                    className: "et-preview-table",
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Location" }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Before" }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(M, { children: "After" }),
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
                          children: i.jsx(M, { children: "Previous changes" }),
                        }),
                        i.jsxs("span", {
                          children: [
                            i.jsx(M, { children: "Page " }),
                            v + 1,
                            " ",
                            i.jsx(M, { children: "of " }),
                            m,
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled: v + 1 >= m,
                          onClick: () => c(v + 1),
                          children: i.jsx(M, { children: "Next changes" }),
                        }),
                      ],
                    }),
                  !p.length &&
                    i.jsx("p", {
                      children: i.jsx(M, {
                        children: "Source and destination already match.",
                      }),
                    }),
                ],
              }),
            (d || h || f.error) &&
              i.jsx("p", {
                className: "et-error",
                role: "alert",
                children: d || (h ? xj : f.error),
              }),
            e.readOnly &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(M, { children: "This workbook is read-only." }),
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "et-footer",
          children: [
            i.jsx("button", {
              type: "button",
              onClick: e.onClose,
              children: i.jsx(M, { children: "Cancel" }),
            }),
            i.jsx("button", {
              type: "button",
              className: "et-primary",
              disabled: e.readOnly || h || !f.plan || !p.length,
              onClick: g,
              children: i.jsx(M, { children: "Apply move" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const _4e = {
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
  D4e = {
    "en-US": "English (United States)",
    "en-GB": "English (United Kingdom)",
    "de-DE": "German (Germany)",
    "fr-FR": "French (France)",
    "ja-JP": "Japanese (Japan)",
    "en-IN": "English (India)",
  };
function O4e(e) {
  var s;
  const r = ct(e.range),
    t = `${ut(r.firstCol)}${r.firstRow + 1}`,
    n = e.active && W3(e.ranges ?? [e.range], e.active) ? e.active : t,
    o =
      (s = e.workbook.sheets.find((d) => d.id === e.sheetId)) == null
        ? void 0
        : s.cells[n];
  if (o != null && o.numberFormat)
    try {
      return l0(o.numberFormat);
    } catch {}
  const a = (o == null ? void 0 : o.format) ?? "general";
  return l0({
    category: a,
    locale: "en-US",
    ...(["number", "currency", "percent"].includes(a)
      ? {
          decimals: Math.min(12, $B(o == null ? void 0 : o.numFmt) ?? 2),
          useGrouping:
            o != null && o.numFmt ? o.numFmt.includes(",") : a !== "percent",
        }
      : {}),
  });
}
function M4e(e) {
  var q, W, K;
  const { text: r, language: t } = Dr(),
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
    [d] = ne.useState(() => O4e(e)),
    [u, l] = ne.useState(d.category),
    [c, h] = ne.useState(d.locale),
    [f, p] = ne.useState(d.currency ?? "USD"),
    [m, v] = ne.useState(String(d.decimals ?? 2)),
    [g, b] = ne.useState(d.useGrouping ?? true),
    [y, w] = ne.useState(String(d.scale ?? 1)),
    [E, A] = ne.useState(d.negative ?? "minus"),
    [k, F] = ne.useState(d.pattern ?? "#,##0.00"),
    [I, O] = ne.useState(""),
    j = e.workbook.sheets.find((te) => te.id === e.sheetId),
    z = ff(e.ranges ?? [e.range]),
    V = ct(e.range),
    D =
      e.active && W3(z, e.active)
        ? e.active
        : `${ut(V.firstCol)}${V.firstRow + 1}`,
    B = j == null ? void 0 : j.cells[D],
    _ = y1(z),
    C = SE(B),
    T =
      s.current.book !== e.workbook.id ||
      s.current.revision !== e.workbook.revision ||
      s.current.content !== JSON.stringify(e.workbook) ||
      s.current.sheet !== e.sheetId ||
      s.current.selection !== JSON.stringify(e.ranges ?? [e.range]) ||
      s.current.active !== e.active,
    N =
      e.readOnly ||
      T ||
      !!((q = j == null ? void 0 : j.protection) != null && q.enabled) ||
      !j ||
      _ > 500,
    R = ne.useMemo(() => Vn(e.workbook), [e.workbook]),
    L = ["number", "currency", "accounting", "percent"].includes(u),
    $ = ["currency", "accounting"].includes(u),
    U = ["date", "time", "datetime", "duration"].includes(u);
  let G,
    H = "",
    Q = null;
  try {
    ((Q = R(e.sheetId, D)),
      (G = l0({
        category: u,
        locale: c,
        ...(L
          ? {
              decimals: m.trim() === "" ? NaN : Number(m),
              useGrouping: g,
              scale: u === "percent" ? 1 : Number(y),
              negative: u === "accounting" ? "parentheses" : E,
              ...($ ? { currency: f } : {}),
            }
          : U || u === "custom"
            ? { pattern: k }
            : {}),
      })));
  } catch (te) {
    H = te instanceof Error ? te.message : "The format could not be previewed.";
  }
  const X$1 = [
      {
        name: `${(j == null ? void 0 : j.name) ?? "Worksheet"}!${D}`,
        value: Q,
      },
      ...(U || (u === "custom" && /[ymdhs]/i.test(k.replace(/"[^"]*"/g, "")))
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
      var ae;
      (l(te),
        O(""),
        te in Ay && F(Ay[te][0]),
        te === "custom" &&
          F(
            ((ae = B == null ? void 0 : B.numberFormat) == null
              ? void 0
              : ae.category) === "custom"
              ? B.numberFormat.pattern
              : '#,##0.00;(#,##0.00);"-"',
          ));
    };
  return (
    ne.useEffect(() => {
      var ae, oe;
      const te = document.activeElement;
      return (
        (oe =
          (ae = o.current) == null
            ? void 0
            : ae.querySelector('[role="option"][aria-selected="true"]')) ==
          null || oe.focus(),
        () => {
          te != null && te.isConnected && te.focus();
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
        onKeyDown: (te) => {
          var ae;
          if (
            (te.stopPropagation(),
            te.key === "Escape" && (te.preventDefault(), a.current()),
            te.key === "Tab")
          ) {
            const oe = [
                ...(((ae = o.current) == null
                  ? void 0
                  : ae.querySelectorAll(
                      'button:not(:disabled):not([tabindex="-1"]),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
                    )) ?? []),
              ],
              de = oe[0],
              he = oe.at(-1);
            te.shiftKey && document.activeElement === de
              ? (te.preventDefault(), he == null || he.focus())
              : !te.shiftKey &&
                document.activeElement === he &&
                (te.preventDefault(), de == null || de.focus());
          }
        },
        children: [
          i.jsxs("header", {
            className: "nf-header",
            children: [
              i.jsx("h2", {
                id: `${n}-title`,
                children: i.jsx(M, { children: "Format Cells" }),
              }),
              i.jsx("button", {
                "aria-label": r("Close number formats"),
                onClick: e.onClose,
                lang: t("Close number formats"),
                "data-ui-label": "Close number formats",
                children: i.jsx(X, { size: 18 }),
              }),
            ],
          }),
          i.jsxs("div", {
            className: "nf-subtitle",
            children: [
              i.jsx(M, { children: "Number · " }),
              _.toLocaleString(),
              " ",
              i.jsx(M, { children: "selected cells · " }),
              z.length,
              " ",
              z.length === 1 ? "area" : "areas",
            ],
          }),
          i.jsxs("div", {
            className: "nf-content",
            children: [
              i.jsx("div", {
                className: "nf-category",
                role: "listbox",
                "aria-label": r("Number format category"),
                lang: t("Number format category"),
                "data-ui-label": "Number format category",
                children: Uv.map((te, ae) =>
                  i.jsx(
                    "button",
                    {
                      role: "option",
                      "aria-selected": u === te,
                      tabIndex: u === te ? 0 : -1,
                      onClick: () => J(te),
                      onKeyDown: (oe) => {
                        var de, he;
                        if (
                          ["ArrowUp", "ArrowDown", "Home", "End"].includes(
                            oe.key,
                          )
                        ) {
                          oe.preventDefault();
                          const Se =
                            oe.key === "Home"
                              ? 0
                              : oe.key === "End"
                                ? Uv.length - 1
                                : (ae +
                                    (oe.key === "ArrowDown"
                                      ? 1
                                      : Uv.length - 1)) %
                                  Uv.length;
                          (J(Uv[Se]),
                            (he =
                              (de = oe.currentTarget.parentElement) == null
                                ? void 0
                                : de.children[Se]) == null || he.focus());
                        }
                      },
                      children: _4e[te],
                    },
                    te,
                  ),
                ),
              }),
              i.jsxs("form", {
                className: "nf-settings",
                onSubmit: (te) => {
                  (te.preventDefault(), O(""));
                  try {
                    if (e.readOnly)
                      throw new Error("This workbook is read-only.");
                    if (T)
                      throw new Error(
                        "The workbook or selection changed. Close and reopen Format Cells.",
                      );
                    if (!G)
                      throw new Error(H || "Choose a valid number format.");
                    const ae = nee(e.workbook, e.sheetId, z, G, {
                      readOnly: e.readOnly,
                    });
                    (ae !== e.workbook &&
                      e.onWorkbookChange(ae, "Format numbers"),
                      e.onClose());
                  } catch (ae) {
                    O(
                      ae instanceof Error
                        ? ae.message
                        : "The format could not be applied.",
                    );
                  }
                },
                children: [
                  T &&
                    i.jsx("p", {
                      className: "nf-error",
                      role: "alert",
                      children: i.jsx(M, {
                        children:
                          "The workbook or selection changed. Close and reopen Format Cells.",
                      }),
                    }),
                  e.readOnly &&
                    i.jsx("p", {
                      className: "nf-notice",
                      children: i.jsx(M, {
                        children:
                          "This workbook is read-only. You can inspect format previews.",
                      }),
                    }),
                  ((W = j == null ? void 0 : j.protection) == null
                    ? void 0
                    : W.enabled) &&
                    i.jsx("p", {
                      className: "nf-notice",
                      children: i.jsx(M, {
                        children:
                          "This worksheet is protected. Unprotect it before changing cell formats.",
                      }),
                    }),
                  _ > 500 &&
                    i.jsx("p", {
                      className: "nf-error",
                      role: "alert",
                      children: i.jsx(M, {
                        children:
                          "Select at most 500 unique cells to apply number formats.",
                      }),
                    }),
                  C && i.jsx("p", { className: "nf-notice", children: C }),
                  i.jsxs("section", {
                    className: "nf-preview",
                    "aria-label": r("Number format preview"),
                    lang: t("Number format preview"),
                    "data-ui-label": "Number format preview",
                    children: [
                      i.jsx("h3", {
                        children: i.jsx(M, { children: "Sample" }),
                      }),
                      i.jsxs("table", {
                        children: [
                          i.jsx("thead", {
                            children: i.jsxs("tr", {
                              children: [
                                i.jsx("th", {
                                  children: i.jsx(M, { children: "Example" }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(M, {
                                    children: "Stored value",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(M, { children: "Display" }),
                                }),
                              ],
                            }),
                          }),
                          i.jsx("tbody", {
                            children: X$1.map((te) => {
                              var oe;
                              const ae = G
                                ? vg(
                                    te.value,
                                    G,
                                    (oe = e.workbook.datePolicy) == null
                                      ? void 0
                                      : oe.dateSystem,
                                  )
                                : void 0;
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
                                      children:
                                        ae != null && ae.accounting
                                          ? i.jsxs("span", {
                                              className: `nf-accounting nf-symbol-${ae.accounting.symbolPosition}`,
                                              children: [
                                                i.jsx("span", {
                                                  children:
                                                    ae.accounting.symbol,
                                                }),
                                                i.jsx("span", {
                                                  children:
                                                    ae.accounting.amount,
                                                }),
                                              ],
                                            })
                                          : ((ae == null ? void 0 : ae.text) ??
                                            "Unsupported pattern"),
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
                          i.jsx(M, { children: "Current display:" }),
                          " ",
                          i.jsx("strong", {
                            children:
                              Af(
                                Q,
                                B,
                                (K = e.workbook.datePolicy) == null
                                  ? void 0
                                  : K.dateSystem,
                              ) || "Blank",
                          }),
                          i.jsx(M, {
                            children:
                              ". Formatting changes appearance only; stored values and formulas stay unchanged.",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Locale" }),
                      i.jsx("select", {
                        "aria-label": r("Number format locale"),
                        value: c,
                        onChange: (te) => h(te.target.value),
                        lang: t("Number format locale"),
                        "data-ui-label": "Number format locale",
                        children: PB.map((te) =>
                          i.jsx("option", { value: te, children: D4e[te] }, te),
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
                            i.jsx(M, { children: "Decimal places" }),
                            i.jsx(ci, {
                              "aria-label": r("Decimal places"),
                              min: 0,
                              max: 12,
                              step: 1,
                              value: m,
                              onValueChange: (te) => v(te),
                              lang: t("Decimal places"),
                              "data-ui-label": "Decimal places",
                            }),
                          ],
                        }),
                        $ &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "Currency" }),
                              i.jsx("select", {
                                "aria-label": r("Currency"),
                                value: f,
                                onChange: (te) => {
                                  p(te.target.value);
                                },
                                lang: t("Currency"),
                                "data-ui-label": "Currency",
                                children: UB.map((te) =>
                                  i.jsx(
                                    "option",
                                    { value: te, children: te },
                                    te,
                                  ),
                                ),
                              }),
                            ],
                          }),
                        u !== "percent" &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "Display scale" }),
                              i.jsxs("select", {
                                "aria-label": r("Display scale"),
                                value: y,
                                onChange: (te) => w(te.target.value),
                                lang: t("Display scale"),
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
                        u !== "accounting" &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "Negative numbers" }),
                              i.jsxs("select", {
                                "aria-label": r("Negative numbers"),
                                value: E,
                                onChange: (te) => A(te.target.value),
                                lang: t("Negative numbers"),
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
                              "aria-label": r("Use thousands separator"),
                              checked: g,
                              onChange: (te) => b(te.target.checked),
                              lang: t("Use thousands separator"),
                              "data-ui-label": "Use thousands separator",
                            }),
                            i.jsx(M, { children: "Use thousands separator" }),
                          ],
                        }),
                      ],
                    }),
                  U &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Type" }),
                        i.jsxs("select", {
                          "aria-label": r("Date or time format"),
                          value: k,
                          onChange: (te) => F(te.target.value),
                          lang: t("Date or time format"),
                          "data-ui-label": "Date or time format",
                          children: [
                            !Ay[u].some((te) => te === k) &&
                              i.jsx("option", { value: k, children: k }),
                            Ay[u].map((te) =>
                              i.jsx("option", { value: te, children: te }, te),
                            ),
                          ],
                        }),
                      ],
                    }),
                  u === "custom" &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Type" }),
                            i.jsx("input", {
                              "aria-label": r("Custom number format"),
                              value: k,
                              maxLength: 200,
                              onChange: (te) => F(te.target.value),
                              lang: t("Custom number format"),
                              "data-ui-label": "Custom number format",
                            }),
                          ],
                        }),
                        i.jsx("p", {
                          className: "nf-help",
                          children: i.jsx(M, {
                            children:
                              "Use 0/# digits, decimals, grouping, percent, quoted labels and positive;negative;zero;text sections. Date/time fields and [h], [m], [s] elapsed time are supported. Colors, conditions, fractions, exponents and spacing-fill codes require other tools or remain unsupported.",
                          }),
                        }),
                      ],
                    }),
                  u === "accounting" &&
                    i.jsx("p", {
                      className: "nf-help",
                      children: i.jsx(M, {
                        children:
                          "Accounting separates the currency symbol and amount, uses a dash for zero and parentheses for negative values.",
                      }),
                    }),
                  U &&
                    i.jsx("p", {
                      className: "nf-help",
                      children: i.jsx(M, {
                        children:
                          "Dates use Excel's 1900 serial system. Pattern order is explicit; the locale changes month names and fractional separators. Elapsed time can exceed 24 hours. Negative calendar dates/times are outside this subset.",
                      }),
                    }),
                  L &&
                    y !== "1" &&
                    u !== "percent" &&
                    i.jsxs("p", {
                      className: "nf-notice",
                      children: [
                        i.jsx(M, {
                          children: "Displayed amounts are divided by",
                        }),
                        " ",
                        Number(y).toLocaleString(),
                        " ",
                        i.jsx(M, { children: "and labelled" }),
                        " ",
                        y === "1000" ? "K" : "M",
                        i.jsx(M, {
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
                          children: i.jsx(M, {
                            children: "Saved format and Excel exchange",
                          }),
                        }),
                        i.jsx("code", { children: Q3(G) }),
                        i.jsx("p", {
                          children: i.jsx(M, {
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
                        children: i.jsx(M, { children: "Cancel" }),
                      }),
                      i.jsx("button", {
                        className: "nf-primary",
                        disabled: N || !G,
                        children: i.jsx(M, { children: "Apply format" }),
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
  Ej =
    "The workbook or selection changed. Close this dialog and reopen it to review the current outline.",
  Wm = (e, r, t) =>
    e === "row"
      ? "Rows " + (r + 1) + (r === t ? "" : "–" + (t + 1))
      : "Columns " + ut(r) + (r === t ? "" : "–" + ut(t));

export { Ej, I4e, M4e, Wm, wS };
