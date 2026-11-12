import { es } from "../01-core-schemas-and-hashing.js";
import { Mt } from "../02-financial-functions.js";
import { Si, _N, gx, jn, oh, vs } from "./01-cell-styles-and-themes.js";
import { Q3, l0 } from "./02-number-formats.js";
import { SR } from "./05-consolidation-and-sources.js";
import { Bu, Hn, NC, QU, Vh, ut } from "./06-data-model-and-analysis.js";
import { Np, _t, ct, pB } from "./07-conditional-formatting.js";
import { Vn, v0 } from "./09-calculation-engine.js";
import { Nr, hR } from "./10-print-layout.js";
import { e$, hi, t$ } from "./11-table-model.js";

function p0(e, r, t = "Preparer", n = "Edit cells", o = {}) {
  if (!Array.isArray(r) || r.length > Mt.maxPatches)
    return Hn(
      `A transaction may contain at most ${Mt.maxPatches} cell changes.`,
    );
  if (r.length === 0) return e;
  (Bu(t, "Actor", 100), Bu(n, "Action", 200));
  const a = new Set();
  let s;
  const d = r.map((p) => {
      var E, A;
      if (!Vh(p)) return Hn("Invalid cell patch.");
      const m = Nr(e, p.sheetId);
      jn(m);
      const { row: v, col: g } = _t(p.address);
      if (v >= m.rowCount || g >= m.colCount)
        return Hn(`Cell ${p.address} is outside sheet dimensions.`);
      const b = `${ut(g)}${v + 1}`;
      _N(e, hi(p.cell).style);
      const y = m.cells[b];
      if (
        ((y == null ? void 0 : y.raw) ?? null) !==
          ((E = p.cell) == null ? void 0 : E.raw) ||
        !!(y != null && y.literal) != !!((A = p.cell) != null && A.literal)
      ) {
        s ?? (s = Vn(e));
        const k = s.getSpillAnchor(m.id, b);
        if (k && k !== b) {
          const F = r.find((I) => I.sheetId === m.id && I.address === k);
          if (
            !F ||
            (!F.cell.literal &&
              typeof F.cell.raw == "string" &&
              F.cell.raw.startsWith("="))
          )
            return Hn(
              `Cell ${b} belongs to the array formula in ${k}. Edit the array formula instead.`,
            );
        }
      }
      const w = `${m.id}:${b}`;
      return a.has(w)
        ? Hn("A transaction cannot change the same cell twice.")
        : (a.add(w),
          {
            sheetId: m.id,
            address: b,
            before: m.cells[b] ? hi(m.cells[b]) : void 0,
            after: hi(p.cell),
          });
    }),
    u = d.filter((p) => !p.before).length;
  if (
    e.sheets.reduce(
      (p, m) =>
        p +
        (es(m.cells)
          ? (QU.get(m.cells) ?? Object.keys(m.cells).length)
          : Object.keys(m.cells).length),
      0,
    ) +
      u >
    Mt.maxCells
  )
    return Hn("Workbook exceeds the 100,000 populated cell limit.");
  const l = e.sheets.map((p) => {
      const m = d.filter((g) => g.sheetId === p.id);
      if (!m.length) return p;
      const v = { ...p.cells };
      return (
        m.forEach((g) => {
          v[g.address] = g.after;
        }),
        { ...p, cells: v }
      );
    }),
    c = hR(e, { ...e, sheets: l });
  if (c.sheets !== l)
    for (const p of c.sheets) {
      const m = Nr(e, p.id);
      for (const [v, g] of Object.entries(p.cells)) {
        const b = d.find((y) => y.sheetId === p.id && y.address === v);
        if (b) b.after = hi(g);
        else if (
          JSON.stringify(m.cells[v]) !== JSON.stringify(g) &&
          (d.push({
            sheetId: p.id,
            address: v,
            before: m.cells[v] ? hi(m.cells[v]) : void 0,
            after: hi(g),
          }),
          d.length > Mt.maxPatches)
        )
          return Hn(
            "Renaming this table field affects more than 500 cells. Reduce the dependent formulas before renaming it.",
          );
      }
    }
  const h = v0(
    c,
    d.filter((p) => {
      var m, v;
      return (
        ((m = p.before) == null ? void 0 : m.raw) !== p.after.raw ||
        ((v = p.before) == null ? void 0 : v.literal) !== p.after.literal
      );
    }),
    o.validationApproval,
  );
  if (
    (h && (n = `${n.slice(0, 140)} (accepted ${h} validation warnings)`),
    e.events.length >= 2e3)
  )
    return Hn(
      "Change history has reached its 2,000 transaction limit. Export the workbook before starting a new workpaper.",
    );
  if (
    e.events.reduce((p, m) => p + m.changes.length, 0) + d.length >
    Mt.maxCells
  )
    return Hn(
      "Change history has reached its 100,000 cell-change limit. Export the workbook before starting a new workpaper.",
    );
  if (e.revision >= Number.MAX_SAFE_INTEGER - 1)
    return Hn("Workbook revision limit reached.");
  const f = {
    ...gx(c),
    revision: e.revision + 1,
    sheets: c.sheets,
    events: [
      ...e.events,
      {
        id: crypto.randomUUID(),
        timestamp: new Date().toISOString(),
        actor: t,
        action: n,
        changes: d,
      },
    ],
  };
  return (
    es(e) &&
      NC.has(e) &&
      c.sheets === l &&
      d.every((p) => {
        var m;
        return (
          !p.after.hyperlink && !((m = p.after.style) != null && m.binding)
        );
      }) &&
      t$(f),
    SR(e, f),
    f
  );
}
function KJ(e, r, t, n, o = {}) {
  o.readOnly && vs("This workbook is read-only.");
  const a = Nr(e, r);
  jn(a);
  const s = oh(n),
    d = Np(typeof t == "string" ? [t] : t),
    u = new Set(d);
  if (o.borders || Object.hasOwn(n, "borders") || Object.hasOwn(n, "border")) {
    const h = d.map(_t);
    (a.merges ?? []).some((f) => {
      const p = ct(f);
      return h.some(
        (m) =>
          m.row >= p.firstRow &&
          m.row <= p.lastRow &&
          m.col >= p.firstCol &&
          m.col <= p.lastCol,
      );
    }) &&
      vs(
        "Unmerge the intersecting cells before applying border edges or presets.",
      );
  }
  const l = d.flatMap((h) => {
    var g, b;
    const f = a.cells[h],
      p = { ...(f == null ? void 0 : f.style), ...n, ...s };
    (Object.hasOwn(n, "border") && delete p.borders,
      o.borders &&
        (p.borders = {
          ...((g = f == null ? void 0 : f.style) == null ? void 0 : g.borders),
          ...pB(h, u, o.borders.preset, o.borders.line),
        }),
      n.borders !== void 0 &&
        (p.borders = {
          ...(o.borders
            ? p.borders
            : (b = f == null ? void 0 : f.style) == null
              ? void 0
              : b.borders),
          ...s.borders,
        }));
    const m = oh(p),
      v = {
        ...(f ?? { raw: null }),
        ...(Object.keys(m).length ? { style: m } : {}),
      };
    return (
      Object.keys(m).length || delete v.style,
      !f && !Object.keys(m).length
        ? []
        : JSON.stringify(oh((f == null ? void 0 : f.style) ?? {})) ===
            JSON.stringify(m)
          ? []
          : JSON.stringify(f) === JSON.stringify(v)
            ? []
            : [{ sheetId: r, address: h, cell: v }]
    );
  });
  if (!l.length) return e;
  const c = p0(e, l, "Preparer", "Format cell styles");
  return { ...e, sheets: c.sheets };
}
function nee(e, r, t, n, o = {}) {
  o.readOnly && Si("This workbook is read-only.");
  const a = Nr(e, r);
  jn(a);
  const s = l0(n),
    d = Np(typeof t == "string" ? [t] : t),
    u = Q3(s),
    l = d.flatMap((h) => {
      const f = a.cells[h],
        p = {
          ...(f ?? { raw: null }),
          numberFormat: s,
          numFmt: u,
          format:
            s.category === "percent"
              ? "percent"
              : s.category === "currency" || s.category === "accounting"
                ? "currency"
                : s.category === "date" || s.category === "datetime"
                  ? "date"
                  : s.category === "number"
                    ? "number"
                    : "general",
        };
      return JSON.stringify(f) === JSON.stringify(p)
        ? []
        : [{ sheetId: r, address: h, cell: p }];
    });
  if (!l.length) return e;
  const c = p0(e, l, "Preparer", "Format numbers");
  return { ...e, sheets: c.sheets };
}
function rh(e) {
  const r = e$(e, true);
  return NC.has(r) && es(r) ? r : t$(r);
}

export { KJ, nee, p0, rh };
