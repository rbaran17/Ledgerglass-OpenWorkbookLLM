import { Mt, Ja, Ig, b1, cx, ux } from "../02-financial-functions.js";
import { es, FJ, GF, jJ, ts, hr, Lg } from "../01-core-schemas-and-hashing.js";
import {
  Ac,
  JJ,
  QS,
  Ql,
  RB,
  TB,
  aQ,
  al,
  cQ,
  fQ,
  hQ,
  ho,
  jn,
  mx,
  oh,
  px,
  tQ,
  uf,
} from "./01-cell-styles-and-themes.js";
import { l0 } from "./02-number-formats.js";
import {
  $a,
  $ne,
  $v,
  AP,
  EP,
  N5,
  SP,
  Vw,
  Xne,
  e0,
  ea,
  eg,
  fp,
  iR,
  kP,
  nR,
  qp,
  rR,
  vC,
  wP,
  yP,
  zw,
} from "./03-defined-names.js";
import {
  Jv,
  Nie,
  OP,
  Sx,
  UP,
  Xt,
  Yv,
  aie,
  jP,
  pie,
  qw,
  wR,
  xC,
} from "./04-pivot-tables-and-pictures.js";
import { Lie, boe, uoe, zie } from "./05-consolidation-and-sources.js";
import {
  Aoe,
  Boe,
  Bu,
  Hn,
  Hoe,
  JU,
  Joe,
  KU,
  L2,
  NC,
  P2,
  QU,
  Vh,
  Voe,
  WU,
  YE,
  YR,
  Yoe,
  Z5,
  gp,
  jm,
  qoe,
  ut,
  wae,
  xd,
  yae,
} from "./06-data-model-and-analysis.js";
import {
  EB,
  GO,
  K3,
  TC,
  VQ,
  X3,
  _t,
  ct,
  tae,
  vae,
} from "./07-conditional-formatting.js";
import {
  $1,
  Ax,
  Cn,
  Dp,
  Ef,
  Gp,
  PP,
  Qf,
  VP,
  Xp,
  bR,
  dR,
  dl,
  kx,
  mR,
  tg,
  wC,
  yC,
} from "./08-chart-model.js";
import { LB, Vn, v0 } from "./09-calculation-engine.js";
import {
  Ix,
  JP,
  Nr,
  O5,
  Roe,
  Rx,
  V1,
  YP,
  fU,
  hR,
  lQ,
  vR,
  yB,
  zP,
} from "./10-print-layout.js";

function Ku(e) {
  return !e.trim() || e.length > 31 || /[\\/\[\]*?:]/.test(e)
    ? Hn(
        "Sheet names must contain 1–31 characters and cannot contain \\ / [ ] * ? : .",
      )
    : {
        id: crypto.randomUUID(),
        name: e.trim(),
        cells: {},
        rowCount: Mt.maxRows,
        colCount: Mt.maxCols,
        frozenRows: 0,
      };
}
function gR(e, r, t, n) {
  var f, p, m, v;
  (n != null && n.sheetId) || $a(e);
  const o = t[0]?.length ?? 0;
  ((!o || t.length * o > Mt.maxCells) &&
    Xt("The analysis output exceeds 100,000 cells."),
    t.some((g) => g.some((b) => typeof b == "string" && b.length > 32767)) &&
      Xt(
        "An analysis label exceeds the 32,767-character cell limit. Shorten source labels or value-field names.",
      ));
  const a = structuredClone(e),
    s =
      n != null && n.sheetId ? Nr(a, n.sheetId) : Ku(pie(a, n?.sheetName ?? r));
  ((n != null && n.sheetId) || a.sheets.push(s), jn(s));
  const d = _t(n?.address ?? "A1"),
    u = d.row + t.length - 1,
    l = d.col + o - 1;
  (u >= s.rowCount || l >= s.colCount) &&
    Xt("The output does not fit the destination worksheet.");
  const c = { firstRow: d.row, lastRow: u, firstCol: d.col, lastCol: l };
  for (const g of Object.keys(s.cells)) {
    const b = _t(g);
    Dp(c, b.row, b.col) &&
      Xt(
        "The output destination contains existing cells. Choose an empty destination or a new sheet.",
      );
  }
  return (
    (f = s.tables) != null &&
      f.some((g) => Qf(c, ct(g.range))) &&
      Xt("The output destination overlaps a table."),
    (((p = s.merges) != null && p.some((g) => Qf(c, ct(g)))) ||
      ((m = s.pivots) != null && m.some((g) => Qf(c, ct(g.outputRange)))) ||
      ((v = s.queries) != null && v.some((g) => Qf(c, ct(g.outputRange))))) &&
      Xt(
        "The output destination overlaps an existing merged area or analysis result. Choose another empty destination.",
      ),
    t.forEach((g, b) =>
      g.forEach((y, w) => {
        s.cells[Cn(d.row + b, d.col + w)] = {
          raw: y,
          ...(typeof y == "string" && y.startsWith("=")
            ? { literal: true }
            : {}),
          ...(b === 0
            ? { style: { bold: true, fill: "#174a7e", color: "#ffffff" } }
            : tg(y)
              ? { format: "number" }
              : {}),
        };
      }),
    ),
    xC(a),
    { workbook: a, sheet: s, range: `${Cn(d.row, d.col)}:${Cn(u, l)}` }
  );
}
function yR(e, r) {
  var h;
  const t = wC(r),
    n = V1(e, t.source),
    o = jP(t, n.fields, mR(n, t.filters ?? [])),
    a = OP(t, n.fields, o),
    s = gR(e, t.name, a.matrix, t.destination),
    { destination: d, ...u } = t,
    l = _t(t.destination?.address ?? "A1");
  for (let f = 1; f < a.matrix.length; f++)
    for (let p = Math.max(1, t.rowFields.length); p < a.matrix[0].length; p++) {
      const m =
        t.values[(p - Math.max(1, t.rowFields.length)) % t.values.length];
      m.showAs &&
        m.showAs !== "value" &&
        (s.sheet.cells[Cn(l.row + f, l.col + p)].format = "percent");
    }
  if (!((h = t.destination) != null && h.sheetId)) {
    let f = 26;
    (a.matrix[0].forEach((p, m) => {
      const v = String(p ?? ""),
        g = Math.min(320, Math.max(90, v.length * 7 + 24));
      s.sheet.columnWidths = { ...s.sheet.columnWidths, [l.col + m]: g };
      const b = Cn(l.row, l.col + m),
        y = s.sheet.cells[b];
      ((y.style = { ...y.style, wrap: true }),
        (f = Math.max(
          f,
          Math.min(120, Math.ceil((v.length * 7) / (g - 16)) * 18 + 8),
        )));
    }),
      (s.sheet.rowHeights = { ...s.sheet.rowHeights, [l.row]: f }));
  }
  for (const f of a.subtotalRows)
    for (let p = 0; p < a.matrix[0].length; p++) {
      const m = Cn(l.row + f, l.col + p),
        v = s.sheet.cells[m];
      v.style = { ...v.style, bold: true, fill: "#e4edf7" };
    }
  return (
    (s.sheet.pivots = [
      ...(s.sheet.pivots ?? []),
      {
        ...u,
        source: n.source,
        id: qw(),
        outputRange: s.range,
        sourceRevision: e.revision,
        refresh: {
          at: new Date().toISOString(),
          sourceSheetName: Nr(e, n.source.sheetId).name,
          sourceRange: n.source.range,
          sourceFingerprint: vR(e, { ...t, source: n.source }, n),
          outputFingerprint: bR(s.sheet, s.range),
          sourceRows: n.rows.length,
          includedRows: o.length,
          excludedRows: n.rows.length - o.length,
          measures: a.measures,
        },
      },
    ]),
    s.workbook
  );
}
function mie(e, r, t) {
  const n =
    e.sheets.flatMap((o) => o.pivots ?? []).find((o) => o.id === r) ??
    Xt("Pivot definition not found.");
  return yR(e, {
    ...n,
    destination: t,
    name: `${n.name.slice(0, 90)} refreshed`,
  });
}
function NO(e, r, t, n) {
  (n.readOnly && Xt("This workbook is read-only."),
    n.snapshot !== JSON.stringify(e) &&
      Xt("The workbook changed. Preview pivot filters again."),
    (!dl(t) ||
      Object.keys(t).some(
        (s) =>
          !["slicers", "timelines", "sorts", "values", "chart"].includes(s),
      )) &&
      Xt(
        "Pivot view changes can only set filters, sorting, value display and chart settings.",
      ));
  const { owner: o, definition: a } = Sx(e, r);
  return (
    jn(o),
    yR(e, {
      ...a,
      ...t,
      name: `${a.name.slice(0, 88)} filtered`,
      destination: undefined,
    })
  );
}
function bie(e, r, t, n) {
  (n.readOnly && Xt("This workbook is read-only."),
    n.snapshot !== JSON.stringify(e) &&
      Xt("The workbook changed. Preview drill-through again."));
  const o = O5(e, r, t),
    { owner: a } = Sx(e, r);
  jn(a);
  const s = JSON.stringify({ ...o, records: undefined });
  s.length > 32700 &&
    Xt(
      "Drill context exceeds the cell-note budget. Reduce filter selections before extracting.",
    );
  const d = [
      [
        "Source locator",
        "Source revision",
        "Original source",
        "Original sheet",
        "Original row",
        ...o.fields,
      ],
      ...o.records.map((l) => {
        return [
          `${l.locator.sheetName}!${l.locator.range}`,
          o.context.sourceRevision,
          l.locator.original?.sourceName ?? null,
          l.locator.original?.sourceSheet ?? null,
          l.locator.original?.originalRow ?? null,
          ...l.values,
        ];
      }),
    ],
    u = gR(e, `${o.context.name.slice(0, 18)} details`, d);
  return (
    (u.sheet.cells.A1.note = `Pivot drill-through context (snapshot, not signed): ${s}`),
    (u.sheet.columnWidths = { 0: 240, 1: 110, 2: 160, 3: 130, 4: 110 }),
    (u.sheet.rowOriginIds = {}),
    (u.workbook.rowOrigins = [...(e.rowOrigins ?? [])]),
    o.records.forEach((l, c) => {
      let h = l.locator.original;
      (h ||
        ((h = {
          id: qw(),
          sourceId: e.id,
          sourceName: e.name,
          sourceSheet: l.locator.sheetName,
          originalRow: l.locator.row,
          kind: "snapshot",
        }),
        u.workbook.rowOrigins.length >= Mt.maxCells &&
          Xt("Row provenance exceeds its 100,000 identity budget."),
        u.workbook.rowOrigins.push(h)),
        (u.sheet.rowOriginIds[c + 1] = h.id),
        l.cells.forEach(({ cell: f }, p) => {
          const m = [
            f.note,
            typeof f.raw == "string" && f.raw.startsWith("=") && !f.literal
              ? `Source formula at extraction: ${f.raw}`
              : undefined,
          ].filter(Boolean).join(`
`);
          (m.length > 32767 &&
            Xt("A source note/formula exceeds the drill-output note budget."),
            m && (u.sheet.cells[`${ut(p + 5)}${c + 2}`].note = m));
        }));
    }),
    u.workbook.sheets.reduce(
      (l, c) => l + Object.keys(c.rowOriginIds ?? {}).length,
      0,
    ) > Mt.maxCells &&
      Xt("Drill row associations exceed the 100,000 source-row budget."),
    u.workbook
  );
}
function ZP(e, r) {
  const t = PP(r),
    n = V1(e, t.source),
    o = mR(n, t.filters ?? []),
    a = t.fields.map((f) => Yv(n.fields, f)),
    s = (t.derived ?? []).map((f) => ({
      ...f,
      index: Yv(n.fields, f.field),
      other: f.otherField ? Yv(n.fields, f.otherField) : undefined,
    }));
  if (t.sort) {
    const f = Yv(n.fields, t.sort.field),
      p = t.sort.direction === "ascending" ? 1 : -1;
    o.sort((m, v) => {
      const g = m.values[f],
        b = v.values[f];
      return (
        (tg(g) && tg(b)
          ? g - b
          : String(g ?? "").localeCompare(String(b ?? ""))) * p ||
        m.index - v.index
      );
    });
  }
  const d = [
      ...t.fields.map((f) =>
        Object.hasOwn(t.rename ?? {}, f) ? t.rename[f] : f,
      ),
      ...s.map((f) => f.name),
    ],
    u = [d];
  (o.length + 1) * d.length > Mt.maxCells &&
    Xt("The query output exceeds 100,000 cells.");
  for (const f of o) {
    const p = a.map((m) => f.values[m]);
    for (const m of s) {
      const v = f.values[m.index],
        g = m.other !== undefined ? f.values[m.other] : (m.constant ?? null);
      if (m.operation === "upper") p.push(String(v ?? "").toUpperCase());
      else if (m.operation === "lower") p.push(String(v ?? "").toLowerCase());
      else if (m.operation === "trim") p.push(String(v ?? "").trim());
      else if (m.operation === "concat")
        p.push(String(v ?? "") + String(g ?? ""));
      else {
        ((!tg(v) || !tg(g)) &&
          Xt(
            `Derived field “${m.name}” requires numbers on source row ${f.index + 1}.`,
          ),
          m.operation === "divide" &&
            g === 0 &&
            Xt(
              `Derived field “${m.name}” divides by zero on source row ${f.index + 1}.`,
            ));
        const b =
          m.operation === "add"
            ? v + g
            : m.operation === "subtract"
              ? v - g
              : m.operation === "multiply"
                ? v * g
                : v / g;
        (Number.isFinite(b) ||
          Xt("Derived calculation exceeds finite numeric bounds."),
          p.push(b));
      }
    }
    (p.some((m) => typeof m == "string" && m.length > 32767) &&
      Xt("Derived text exceeds the cell text limit."),
      u.push(p));
  }
  const l = gR(e, t.name, u, t.destination),
    { destination: c, ...h } = t;
  return (
    (l.sheet.queries = [
      ...(l.sheet.queries ?? []),
      {
        ...h,
        source: n.source,
        id: qw(),
        sourceRevision: e.revision,
        outputRange: l.range,
      },
    ]),
    l.workbook
  );
}
function yie(e, r, t) {
  const n =
    e.sheets.flatMap((o) => o.queries ?? []).find((o) => o.id === r) ??
    Xt("Query definition not found.");
  return ZP(e, {
    ...n,
    destination: t,
    name: `${n.name.slice(0, 90)} refreshed`,
  });
}
function U2(e, r) {
  var o;
  if (
    r &&
    es(e) &&
    JU.has(e) &&
    !e.hyperlink &&
    !((o = e.style) != null && o.binding)
  )
    return e;
  if (!Vh(e) || !wae(e, "raw")) return Hn("A cell must have a raw value.");
  const t = e.raw;
  if (
    t !== null &&
    typeof t != "string" &&
    typeof t != "boolean" &&
    (typeof t != "number" || !Number.isFinite(t))
  )
    return Hn("Cell values must be text, finite numbers, booleans, or null.");
  typeof t == "string" &&
    Bu(
      t,
      "Cell value",
      t.startsWith("=") && e.literal !== true
        ? Ja.maxFormulaLength
        : Ja.maxCellTextLength,
    );
  const n = { raw: t };
  if (e.literal !== undefined) {
    if (typeof e.literal != "boolean")
      return Hn("Cell literal flag must be boolean.");
    n.literal = e.literal;
  }
  if (
    (e.numberFormat !== undefined && (n.numberFormat = l0(e.numberFormat)),
    e.numFmt !== undefined &&
      (n.numFmt = Bu(e.numFmt, "Excel number format", 200)),
    e.format !== undefined)
  ) {
    if (typeof e.format != "string" || !yae.has(e.format))
      return Hn("Unsupported cell format.");
    n.format = e.format;
  }
  return (
    e.hyperlink !== undefined && (n.hyperlink = X3(e.hyperlink)),
    e.note !== undefined && (n.note = Bu(e.note, "Cell note", 1e4)),
    e.tickmark !== undefined && (n.tickmark = Bu(e.tickmark, "Tickmark", 40)),
    e.style !== undefined && (n.style = oh(e.style)),
    n
  );
}
function hi(e) {
  return U2(e, false);
}
function qQ(e, r, t, n) {
  const o = e.sheets.find((l) => l.id === r);
  jn(o);
  const a = Ix(n);
  a.length > 500 && al("Paint at most 500 cells.");
  const s = o.cells[t] ?? {},
    d = { ...o.cells };
  for (const l of a)
    d[l] = hi({
      ...d[l],
      raw: d[l]?.raw ?? null,
      style: structuredClone(s.style ?? {}),
      format: s.format ?? "general",
      numFmt: s.numFmt,
      numberFormat: s.numberFormat
        ? structuredClone(s.numberFormat)
        : undefined,
    });
  return {
    ...e,
    sheets: e.sheets.map((l) => (l.id === r ? { ...l, cells: d } : l)),
  };
}
function bC(e, r) {
  if (e === undefined) return;
  const t = () => {
      throw Error(
        "Invalid or over-capacity table source evidence. Keep the native backup.",
      );
    },
    n = (s) => !!s && typeof s == "object" && !Array.isArray(s);
  (!Array.isArray(e) || e.length > 100 || JSON.stringify(e).length > 5e6) &&
    t();
  const o = new Set();
  let a = 0;
  r.tableEvidence = e.map((s) => {
    (!n(s) ||
      Object.keys(s).some(
        (c) =>
          ![
            "id",
            "kind",
            "sourceRevision",
            "sheetId",
            "sheetName",
            "tableId",
            "tableName",
            "columnId",
            "dataType",
            "changes",
          ].includes(c),
      )) &&
      t();
    const d = s;
    for (const c of ["id", "sheetId", "sheetName", "tableId", "tableName"])
      (typeof d[c] != "string" || !d[c] || d[c].length > 512) && t();
    ((o.has(d.id) ||
      !Number.isSafeInteger(d.sourceRevision) ||
      Number(d.sourceRevision) < 0 ||
      !["headers", "columnType"].includes(String(d.kind)) ||
      !Array.isArray(d.changes) ||
      d.changes.length > 500 ||
      (a += d.changes.length) > 1e4) &&
      t(),
      d.kind === "columnType"
        ? ((typeof d.columnId != "string" ||
            !d.columnId ||
            d.columnId.length > 100) &&
            t(),
          iR(d.dataType))
        : (d.columnId !== undefined || d.dataType !== undefined) && t(),
      o.add(d.id));
    const u = new Set(),
      l = d.changes.map((c) => {
        (!n(c) ||
          Object.keys(c).some(
            (m) => !["address", "before", "after"].includes(m),
          ) ||
          typeof c.address != "string") &&
          t();
        const h = c,
          f = h.address,
          p = _t(f);
        return (
          (`${ut(p.col)}${p.row + 1}` !== f || u.has(f)) && t(),
          u.add(f),
          {
            address: f,
            ...(h.before !== undefined ? { before: hi(h.before) } : {}),
            after: hi(h.after),
          }
        );
      });
    return { ...d, changes: l };
  });
}
function RP(e, r, t, n, o = {}) {
  if (o.readOnly) throw Error("This workbook is read-only.");
  if (
    !n ||
    typeof n != "object" ||
    Object.keys(n).some(
      (E) =>
        ![
          "columnId",
          "dataType",
          "mode",
          "invalid",
          "decimal",
          "group",
        ].includes(E),
    )
  )
    throw Error("Invalid column type request.");
  if (
    (iR(n.dataType),
    !["convert", "tag"].includes(n.mode) ||
      !["reject", "keep"].includes(n.invalid) ||
      ![undefined, ".", ","].includes(n.decimal) ||
      ![undefined, "", ".", ",", " "].includes(n.group) ||
      (n.decimal ?? ".") === (n.group ?? ","))
  )
    throw Error(
      "Choose a conversion mode, exception resolution and different decimal/group separators.",
    );
  const a = structuredClone(n),
    s = Nr(e, r),
    d = s.tables?.find((E) => E.id === t);
  if ((jn(s), !d)) throw Error("Table not found.");
  const u = d.columns.findIndex((E) => E.id === a.columnId);
  if (u < 0) throw Error("Select an existing table column.");
  const l = ct(d.range),
    c = l.lastRow - l.firstRow - Number(d.showTotals);
  if (c > 500)
    throw Error(
      "A reviewed column type operation supports at most 500 data rows. Preserve the native source and use bounded extracts for larger conversions.",
    );
  const h = structuredClone(e),
    f = Nr(h, r),
    p = f.tables.find((E) => E.id === t).columns[u];
  a.dataType === "any" ? delete p.dataType : (p.dataType = a.dataType);
  const m = [],
    v = [],
    g = Vn(e);
  for (let E = 1; E <= c; E++) {
    const A = `${ut(l.firstCol + u)}${l.firstRow + E + 1}`,
      k = s.cells[A],
      F = k?.raw ?? null,
      I =
        typeof F == "string" && F.startsWith("=") && !(k != null && k.literal);
    try {
      if (I || a.mode === "tag") {
        if (!Vw(g(r, A), a.dataType))
          throw Error(
            I
              ? "Formula result does not match the selected type; the formula is preserved."
              : "Value does not match the selected type; source values are unchanged.",
          );
        continue;
      }
      const O = g(r, A);
      if (N5(F) && !N5(O)) {
        if (!Vw(O, a.dataType))
          throw Error(
            "Spill result does not match the selected type; the spill is preserved.",
          );
        continue;
      }
      const j = Xne(F, a);
      if (j !== F) {
        const z = hi({
          ...k,
          raw: j,
          ...(typeof j == "string" && j.startsWith("=")
            ? { literal: true }
            : { literal: undefined }),
        });
        ((f.cells[A] = z),
          m.push({
            address: A,
            ...(k ? { before: structuredClone(k) } : {}),
            after: structuredClone(z),
          }));
      }
    } catch (O) {
      v.push({ address: A, value: F, reason: O.message });
    }
  }
  const b = d.columns[u].dataType ?? "any";
  return (
    (m.length || b !== a.dataType) &&
      ((h.tableEvidence = [
        ...(h.tableEvidence ?? []),
        {
          id: `${t}:${p.id}:${e.revision}:${e.tableEvidence?.length ?? 0}`,
          kind: "columnType",
          sourceRevision: e.revision,
          sheetId: r,
          sheetName: s.name,
          tableId: t,
          tableName: d.name,
          columnId: p.id,
          dataType: a.dataType,
          changes: structuredClone(m),
        },
      ]),
      bC(h.tableEvidence, h)),
    v0(
      h,
      m.map((E) => ({ ...E, sheetId: r })),
    ),
    {
      snapshot: JSON.stringify(e),
      sheetId: r,
      tableId: t,
      request: a,
      beforeType: b,
      canApply: !v.length || a.invalid === "keep",
      changes: m,
      exceptions: v,
      next: h,
    }
  );
}
function Zne(e, r, t = {}) {
  if (JSON.stringify(e) !== r.snapshot)
    throw Error("The workbook changed. Preview column types again.");
  const n = RP(e, r.sheetId, r.tableId, r.request, t);
  if (JSON.stringify(n) !== JSON.stringify(r))
    throw Error("The column type preview changed. Preview again.");
  if (!n.canApply)
    throw Error(
      "Resolve the listed exceptions or explicitly choose to retain them unchanged.",
    );
  return JSON.stringify(e) === JSON.stringify(n.next) ? e : n.next;
}
function fie(e, r, t) {
  var c, h;
  const n = Nr(e, r),
    o = dR(t.name),
    a = Ef(t.range, n);
  if (
    (jn(n),
    (c = n.merges) != null &&
      c.some((f) => Qf(a, ct(f))) &&
      Xt("A table cannot overlap merged cells. Unmerge the source first."),
    e.sheets.some((f) => {
      return f.tables?.some((m) => m.name.toLowerCase() === o.toLowerCase());
    }) && Xt("A table with that name already exists."),
    t.showTotals)
  ) {
    a.lastRow + 1 >= n.rowCount && Xt("There is no room for a totals row.");
    for (let f = a.firstCol; f <= a.lastCol; f++)
      n.cells[Cn(a.lastRow + 1, f)] &&
        Xt("The totals row would overwrite existing cells.");
    a.lastRow++;
  }
  ($1(a) > Mt.maxCells ||
    ((h = n.tables) != null && h.some((f) => Qf(a, ct(f.range))))) &&
    Xt("The table overlaps another table or exceeds the cell budget.");
  const s = structuredClone(e),
    d = Nr(s, r),
    u = {
      id: qw(),
      name: o,
      range: `${Cn(a.firstRow, a.firstCol)}:${Cn(a.lastRow, a.lastCol)}`,
      columns: zP(e, r, t.range, t.headers).map((f) => ({ id: qw(), name: f })),
      style: t.style ?? "blue",
      showTotals: !!t.showTotals,
      totals: t.totals ?? {},
    };
  ((d.tables = [...(d.tables ?? []), u]),
    Ax(d, d),
    VP(d, d.tables[d.tables.length - 1]));
  const l = u.columns.flatMap((f, p) => {
    const m = Cn(a.firstRow, a.firstCol + p),
      v = n.cells[m],
      g = d.cells[m];
    return v?.raw === f.name &&
      !(f.name.startsWith("=") && !(v != null && v.literal))
      ? []
      : [
          {
            address: m,
            ...(v ? { before: structuredClone(v) } : {}),
            after: structuredClone(g),
          },
        ];
  });
  return (
    l.length &&
      ((s.tableEvidence = [
        ...(s.tableEvidence ?? []),
        {
          id: `${u.id}:headers`,
          kind: "headers",
          sourceRevision: e.revision,
          sheetId: r,
          sheetName: n.name,
          tableId: u.id,
          tableName: o,
          changes: l,
        },
      ]),
      bC(s.tableEvidence, s)),
    xC(s)
  );
}
function KP(e, r, t, n) {
  const o = Gp(r);
  for (const [a, s] of r.columns.entries())
    if (s.calculatedFormula)
      for (let d = t; d <= n; d++) {
        const u = Cn(d, o.firstCol + a),
          l = xd(s.calculatedFormula, d - o.firstRow, 0);
        wR(l);
        const c = { ...(e.cells[u] ?? { raw: null }), raw: l };
        (delete c.literal, (e.cells[u] = hi(c)));
      }
}
function hie(e, r, t, n) {
  var w;
  const o = Nr(e, r),
    a = o.tables?.find((A) => A.id === t) ?? Xt("Table not found.");
  jn(o);
  const s = dR(n.name),
    d = Ef(n.range, o),
    u = ct(a.range);
  ((d.firstRow !== u.firstRow || d.firstCol !== u.firstCol) &&
    Xt(
      "Resize a table from its existing top-left header. Moving table source cells is a separate operation.",
    ),
    e.sheets.some((A) => {
      return A.tables?.some(
        (F) => F.id !== t && F.name.toLowerCase() === s.toLowerCase(),
      );
    }) && Xt("A table with that name already exists."));
  const l = n.showTotals ?? a.showTotals;
  if (
    (l !== a.showTotals && Xp(n.range) === a.range && (d.lastRow += l ? 1 : -1),
    (d.lastRow >= o.rowCount ||
      d.lastRow < d.firstRow ||
      $1(d) > Mt.maxCells) &&
      Xt("The resized table does not fit the worksheet or processing budget."),
    (w = o.tables) != null &&
      w.some((A) => A.id !== t && Qf(d, ct(A.range))) &&
      Xt("Tables cannot overlap."),
    l && (!a.showTotals || d.lastRow !== u.lastRow))
  )
    for (let A = d.firstCol; A <= d.lastCol; A++)
      o.cells[Cn(d.lastRow, A)] &&
        Xt(
          "The new totals row would overwrite existing cells. Include an empty final row in the resized range.",
        );
  const c = a.columns.some((A) => A.calculatedFormula);
  (c &&
    d.lastCol !== u.lastCol &&
    Xt("Clear calculated column rules before changing table columns."),
    c && (d.lastRow !== u.lastRow || l !== a.showTotals) && YP(e, r, a, d, l));
  const h = structuredClone(e),
    f = Nr(h, r);
  if (a.showTotals && (!l || d.lastRow !== u.lastRow))
    for (let A = u.firstCol; A <= u.lastCol; A++) {
      const k = Cn(u.lastRow, A),
        F = f.cells[k],
        I = kx(a, A - u.firstCol);
      (F &&
        F.raw !== I.raw &&
        Xt(
          "The old totals row was edited. Preserve or move those values before resizing it.",
        ),
        delete f.cells[k]);
    }
  const p = zP(h, r, n.range, n.headers),
    m = p.map((A, k) => {
      return {
        ...a.columns[k],
        id: a.columns[k]?.id ?? `${a.id}:column:${k}:${e.revision}`,
        name: A,
      };
    }),
    v =
      n.totals ??
      Object.fromEntries(
        Object.entries(a.totals).filter(([A]) => m.some((k) => k.name === A)),
      ),
    g = {
      ...a,
      name: s,
      range: `${Cn(d.firstRow, d.firstCol)}:${Cn(d.lastRow, d.lastCol)}`,
      columns: m,
      style: n.style ?? a.style,
      showTotals: l,
      totals: v,
    };
  ((f.tables = f.tables.map((A) => (A.id === t ? g : A))), Ax(f, f), VP(f, g));
  const b = m.flatMap((A, k) => {
    const F = Cn(d.firstRow, d.firstCol + k),
      I = o.cells[F],
      O = f.cells[F];
    return I?.raw === O.raw && !!(I != null && I.literal) == !!O.literal
      ? []
      : [
          {
            address: F,
            ...(I ? { before: structuredClone(I) } : {}),
            after: structuredClone(O),
          },
        ];
  });
  if (
    (b.length &&
      ((h.tableEvidence = [
        ...(h.tableEvidence ?? []),
        {
          id: `${t}:headers:${e.revision}:${h.tableEvidence?.length ?? 0}`,
          kind: "headers",
          sourceRevision: e.revision,
          sheetId: r,
          sheetName: o.name,
          tableId: t,
          tableName: s,
          changes: b,
        },
      ]),
      bC(h.tableEvidence, h)),
    c)
  ) {
    const A = Gp(a),
      k = Gp(g);
    for (
      let F = u.firstRow;
      d.lastRow !== u.lastRow && F <= Math.min(A.lastRow, k.lastRow);
      F++
    )
      for (let I = A.firstCol; I <= A.lastCol; I++) {
        const O = Cn(F, I);
        o.cells[O] && (f.cells[O] = structuredClone(o.cells[O]));
      }
    (a.showTotals && l && d.lastRow !== u.lastRow && JP(e, h, r, a, g),
      KP(f, g, Math.max(A.lastRow + 1, k.firstRow), k.lastRow));
  }
  return xC(hR(e, h));
}
function QP(e, r, t, n, o = {}) {
  var m;
  o.readOnly && Xt("This workbook is read-only.");
  const a = Nr(e, r),
    s = a.tables?.find((v) => v.id === t) ?? Xt("Select an existing table.");
  (jn(a), dl(n) || Xt("Choose a table operation."));
  const d = {
    setColumnFormula: ["kind", "columnId", "formula"],
    clearColumnFormula: ["kind", "columnId"],
    appendRows: ["kind", "count", "rows"],
    resize: ["kind", "range"],
    updateTable: ["kind", "options"],
  };
  (!d[n.kind] || Object.keys(n).some((v) => !d[n.kind].includes(v))) &&
    Xt("Unknown calculated-table operation field.");
  let u = structuredClone(e),
    l = Nr(u, r),
    c = l.tables.find((v) => v.id === t);
  if (n.kind === "setColumnFormula" || n.kind === "clearColumnFormula") {
    const v =
      c.columns.find((g) => g.id === n.columnId) ??
      Xt("Select an existing table column.");
    if (n.kind === "setColumnFormula") {
      v.calculatedFormula = wR(n.formula);
      const g = Gp(c);
      (g.lastRow - g.firstRow + 1 > Jv.changedCells &&
        Xt("A calculated column fill supports at most 500 data-body rows."),
        KP(
          l,
          {
            ...c,
            columns: c.columns.map((b) =>
              b.id === v.id ? b : { ...b, calculatedFormula: undefined },
            ),
          },
          g.firstRow,
          g.lastRow,
        ));
    } else delete v.calculatedFormula;
  } else if (
    n.kind === "appendRows" ||
    n.kind === "resize" ||
    n.kind === "updateTable"
  ) {
    const v = ct(s.range),
      g = Gp(s);
    let b,
      y = 0,
      w;
    n.kind === "appendRows"
      ? (n.rows !== undefined
          ? (n.count !== undefined &&
              Xt("Specify raw rows or a row count, not both."),
            (!Array.isArray(n.rows) ||
              n.rows.length > Jv.appendedRows ||
              n.rows.some(
                (k) => !Array.isArray(k) || k.length !== s.columns.length,
              )) &&
              Xt(
                "Each appended row must contain one raw value for every table column.",
              ),
            (b = n.rows),
            (y = b.length))
          : (y = n.count ?? 0),
        (!Number.isInteger(y) || y < 1 || y > Jv.appendedRows) &&
          Xt("Append 1–100 rows at a time."),
        v.lastRow + y >= a.rowCount &&
          Xt("The appended table exceeds the worksheet boundary."),
        (w = {
          ...s,
          range: `${Cn(v.firstRow, v.firstCol)}:${Cn(v.lastRow + y, v.lastCol)}`,
        }))
      : (w = n.kind === "resize" ? { ...s, range: n.range } : n.options);
    const E = Ef(w.range, a),
      A = w.showTotals ?? s.showTotals;
    if (
      ((n.kind !== "updateTable" ||
        s.columns.some((k) => k.calculatedFormula)) &&
        YP(e, r, s, E, A),
      (u = hie(e, r, t, w)),
      (l = Nr(u, r)),
      (c = l.tables.find((k) => k.id === t)),
      n.kind === "appendRows" && !s.columns.some((k) => k.calculatedFormula))
    ) {
      for (let k = v.firstRow; k <= g.lastRow; k++)
        for (let F = g.firstCol; F <= g.lastCol; F++) {
          const I = Cn(k, F);
          a.cells[I] && (l.cells[I] = structuredClone(a.cells[I]));
        }
      s.showTotals && JP(e, u, r, s, c);
    }
    if (b)
      for (let k = 0; k < b.length; k++)
        for (let F = 0; F < s.columns.length; F++) {
          const I = b[k][F];
          if (s.columns[F].calculatedFormula) {
            I !== null &&
              I !== "" &&
              Xt(
                "Leave calculated-column input slots blank; their saved formulas supply appended values.",
              );
            continue;
          }
          Vw(I, s.columns[F].dataType) ||
            Xt(
              `Appended value in ${s.columns[F].name} does not match its ${s.columns[F].dataType} type. Supply a typed value or explicitly revise the column type; no rows were appended.`,
            );
          const O = Cn(g.lastRow + 1 + k, g.firstCol + F),
            j = hi({
              ...l.cells[O],
              raw: I,
              ...(typeof I == "string" && I.startsWith("=")
                ? { literal: true }
                : { literal: undefined }),
            });
          l.cells[O] = j;
        }
  } else Xt("Unknown calculated-table operation.");
  (Ax(l, l), xC(u));
  const h = [];
  for (const v of u.sheets) {
    const g = Nr(e, v.id);
    (m = g.protection) != null &&
      m.enabled &&
      JSON.stringify(g) !== JSON.stringify(v) &&
      Xt(
        "An affected worksheet is protected. Unprotect it before changing dependent table formulas.",
      );
    for (const b of new Set([
      ...Object.keys(g.cells),
      ...Object.keys(v.cells),
    ])) {
      const y = g.cells[b],
        w = v.cells[b];
      (JSON.stringify(y) !== JSON.stringify(w) &&
        h.push({
          sheetId: v.id,
          address: b,
          ...(y ? { before: structuredClone(y) } : {}),
          ...(w ? { after: structuredClone(w) } : {}),
        }),
        h.length > Jv.changedCells &&
          Xt("A reviewed table operation may change at most 500 cells."));
    }
  }
  v0(u, h);
  const f = [];
  for (const [v, g, b] of [
    ["watch", e.formulaWatches, u.formulaWatches],
    ["review", e.reviewThreads, u.reviewThreads],
  ])
    for (const y of b ?? []) {
      const w = g?.find((E) => E.id === y.id);
      w &&
        w.anchor.address !== y.anchor.address &&
        f.push({
          kind: v,
          id: y.id,
          before: w.anchor.address,
          after: y.anchor.address,
        });
    }
  return {
    snapshot: JSON.stringify(e),
    sheetId: r,
    tableId: t,
    request: structuredClone(n),
    beforeRange: s.range,
    afterRange: c.range,
    beforeColumns: structuredClone(s.columns),
    afterColumns: structuredClone(c.columns),
    changes: h,
    anchorChanges: f,
    next: u,
  };
}
function Eie(e, r, t = {}) {
  r.snapshot !== JSON.stringify(e) &&
    Xt("The workbook changed after the table preview. Preview again.");
  const n = QP(e, r.sheetId, r.tableId, r.request, t);
  return (
    JSON.stringify(n) !== JSON.stringify(r) &&
      Xt("The table preview was changed. Preview again."),
    JSON.stringify(n.next) === JSON.stringify(e) ? e : n.next
  );
}
function HU(e, r) {
  if (e === undefined) return;
  (!Array.isArray(e) || e.length > 100 || JSON.stringify(e).length > 5e6) &&
    gp();
  const t = new Set();
  let n = 0;
  r.textCleanups = e.map((o) => {
    L2(o, [
      "id",
      "workbookId",
      "sourceRevision",
      "sourceSheetId",
      "sourceSheetName",
      "outputSheetId",
      "outputSheetName",
      "createdAt",
      "steps",
      "semantics",
      "rows",
    ]);
    for (const l of ["id", "workbookId", "sourceSheetId", "outputSheetId"])
      jm(o[l], 100) || gp();
    ((!jm(o.sourceSheetName, 31) ||
      !jm(o.outputSheetName, 31) ||
      t.has(o.id) ||
      !Number.isSafeInteger(o.sourceRevision) ||
      Number(o.sourceRevision) < 0 ||
      o.semantics !== "unicode-v2-en-US-explicit-separators" ||
      !Array.isArray(o.rows) ||
      !o.rows.length ||
      o.rows.length > 500 ||
      (n += o.rows.length) > 1e4 ||
      typeof o.createdAt != "string" ||
      !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(o.createdAt) ||
      !Number.isFinite(Date.parse(o.createdAt)) ||
      new Date(o.createdAt).toISOString() !== o.createdAt) &&
      gp(),
      t.add(o.id));
    const a = YR(o.steps),
      s = new Set(),
      d = new Set(),
      u = o.rows.map((l) => {
        (L2(l, [
          "sourceAddress",
          "original",
          "result",
          "outputAddress",
          "origin",
        ]),
          (!GO(l.sourceAddress) ||
            !GO(l.outputAddress) ||
            s.has(l.sourceAddress) ||
            d.has(l.outputAddress) ||
            !WU(l.result)) &&
            gp(),
          s.add(l.sourceAddress),
          d.add(l.outputAddress));
        const c = l.original === undefined ? undefined : hi(l.original);
        c &&
          typeof c.raw == "string" &&
          c.raw.startsWith("=") &&
          !c.literal &&
          gp();
        let h;
        if (l.origin !== undefined) {
          L2(l.origin, [
            "id",
            "sourceId",
            "sourceName",
            "sourceSheet",
            "originalRow",
            "kind",
          ]);
          const f = l.origin;
          ((!jm(f.id, 100) ||
            !jm(f.sourceId, 100) ||
            !jm(f.sourceName, 255) ||
            !jm(f.sourceSheet, 31) ||
            !Number.isInteger(f.originalRow) ||
            Number(f.originalRow) < 1 ||
            Number(f.originalRow) > 1048576 ||
            !["snapshot", "file"].includes(String(f.kind))) &&
            gp(),
            (h = { ...f }));
        }
        return {
          sourceAddress: l.sourceAddress,
          outputAddress: l.outputAddress,
          result: l.result,
          ...(c ? { original: c } : {}),
          ...(h ? { origin: h } : {}),
        };
      });
    return { ...o, steps: a, rows: u };
  });
}
function YU(e, r) {
  if (e === undefined) return;
  if (
    !Array.isArray(e) ||
    e.length > 100 ||
    JSON.stringify(e).length > Z5.evidenceCharacters
  )
    throw Error(
      "Field transformation evidence exceeds its native history budget. Keep the original backup.",
    );
  const t = new Set();
  let n = 0;
  r.fieldTransforms = e.map((o) => {
    const a = vae.safeParse(o);
    if (!a.success)
      throw Error(
        `Invalid field transformation evidence: ${a.error.issues[0]?.message}`,
      );
    const s = a.data,
      d = KU(s.request),
      u = ct(d.sourceRange);
    if (
      new Date(s.createdAt).toISOString() !== s.createdAt ||
      s.sources.length !==
        (u.lastRow - u.firstRow + 1) * (u.lastCol - u.firstCol + 1)
    )
      throw Error(
        "Invalid field evidence timestamp or incomplete original source population.",
      );
    if (t.has(s.id) || (n += s.outputs.length) > 2e4)
      throw Error("Duplicate or over-capacity field transformation evidence.");
    if (
      (t.add(s.id),
      s.sourceSheetId !== d.sourceSheetId || s.outputSheetName !== d.outputName)
    )
      throw Error("Field transformation evidence does not match its recipe.");
    const l = new Set(),
      c = s.sources.map((p) => {
        const m = _t(p.address);
        if (
          m.row < u.firstRow ||
          m.row > u.lastRow ||
          m.col < u.firstCol ||
          m.col > u.lastCol
        )
          throw Error(
            "An original field source is outside the recorded range.",
          );
        if (l.has(p.address)) throw Error("Duplicate original field source.");
        l.add(p.address);
        const v = p.original === undefined ? undefined : hi(p.original);
        if (
          v &&
          typeof v.raw == "string" &&
          v.raw.startsWith("=") &&
          !v.literal
        )
          throw Error("Field source evidence must contain reviewed values.");
        return { ...p, ...(v ? { original: v } : {}) };
      }),
      h = new Set();
    for (const p of s.outputs) {
      if (h.has(p.address))
        throw Error("Duplicate transformed output address.");
      if (
        (h.add(p.address),
        ((p.type === "number" || p.type === "date") &&
          typeof p.value != "number") ||
          (p.type === "boolean" && typeof p.value != "boolean") ||
          (p.type === "text" && typeof p.value != "string") ||
          (p.type === "blank" && p.value !== null))
      )
        throw Error("Transformed output type does not match its value.");
    }
    for (const p of [...s.outputs, ...s.exceptions])
      if (
        p.row < u.firstRow + 1 + Number(d.header) ||
        p.row > u.lastRow + 1 ||
        p.sources.some((m) => !l.has(m)) ||
        new Set(p.sources).size !== p.sources.length
      )
        throw Error("Unknown or repeated original field source locator.");
    return { ...s, request: d, sources: c };
  });
}
function e$(e, r) {
  if (r && es(e) && NC.has(e)) return e;
  if (!Vh(e) || e.schemaVersion !== 1)
    return Hn("Unsupported workbook schema. Expected schemaVersion 1.");
  const t = Bu(e.id, "Workbook ID", 100),
    n = Bu(e.name, "Workbook name", 200);
  if (!t || !n.trim()) return Hn("Workbook ID and name are required.");
  const o = P2(e.revision, 0, Number.MAX_SAFE_INTEGER - 1, "Revision");
  if (!Array.isArray(e.sheets) || !e.sheets.length)
    return Hn("A workbook must contain at least one sheet.");
  let a = 0;
  const s = new Set(),
    d = new Set(),
    u = e.sheets.map((p) => {
      if (!Vh(p)) return Hn("Invalid sheet.");
      const m = Bu(p.id, "Sheet ID", 100),
        v = Bu(p.name, "Sheet name", 31);
      if (
        !m ||
        !v.trim() ||
        /[\\/\[\]*?:]/.test(v) ||
        s.has(m) ||
        d.has(v.toLowerCase())
      )
        return Hn(
          "Sheet IDs and names must be nonempty and unique; sheet names cannot contain Excel-reserved characters.",
        );
      (s.add(m), d.add(v.toLowerCase()));
      const g = P2(p.rowCount, 1, Mt.maxRows, "Row count"),
        b = P2(p.colCount, 1, Mt.maxCols, "Column count");
      if (!Vh(p.cells)) return Hn("Sheet cells must be an address map.");
      const y = r && es(p) ? YE.get(p) : undefined;
      if (y !== undefined) {
        if (((a += y), a > Mt.maxCells))
          return Hn("Workbook exceeds the 100,000 populated cell limit.");
        const A = { ...p };
        return (YE.set(A, y), A);
      }
      const w = {};
      for (const [A, k] of Object.entries(p.cells)) {
        const F = _t(A);
        if (A !== `${ut(F.col)}${F.row + 1}` || F.row >= g || F.col >= b)
          return Hn(`Invalid or out-of-bounds cell: ${A}`);
        if (++a > Mt.maxCells)
          return Hn("Workbook exceeds the 100,000 populated cell limit.");
        w[A] = U2(k, r);
      }
      const E = {
        id: m,
        name: v,
        rowCount: Mt.maxRows,
        colCount: Mt.maxCols,
        cells: w,
      };
      if (p.columnWidths !== undefined) {
        if (!Vh(p.columnWidths))
          return Hn("Column widths must be an index map.");
        E.columnWidths = {};
        for (const [A, k] of Object.entries(p.columnWidths)) {
          if (
            !/^(0|[1-9]\d*)$/.test(A) ||
            Number(A) >= b ||
            typeof k != "number" ||
            !Number.isFinite(k) ||
            k < ux ||
            k > cx
          )
            return Hn(
              "Column widths must be 40–600 pixels for an existing column.",
            );
          E.columnWidths[A] = k;
        }
      }
      if (p.rowHeights !== undefined) {
        if (
          !Vh(p.rowHeights) ||
          (Object.getPrototypeOf(p.rowHeights) !== Object.prototype &&
            Object.getPrototypeOf(p.rowHeights) !== null)
        )
          return Hn("Row heights must be a plain index map.");
        E.rowHeights = {};
        for (const [A, k] of Object.entries(p.rowHeights)) {
          if (
            !/^(0|[1-9]\d*)$/.test(A) ||
            Number(A) >= g ||
            typeof k != "number" ||
            !Number.isFinite(k) ||
            k < b1 ||
            k > Ig
          )
            return Hn(
              `Row heights must be ${b1}–${Ig} pixels for an existing row.`,
            );
          E.rowHeights[A] = k;
        }
      }
      if (p.hidden !== undefined) {
        if (typeof p.hidden != "boolean")
          return Hn("Sheet hidden flag must be boolean.");
        E.hidden = p.hidden;
      }
      return (
        VQ(p, E),
        TB(p, E, { rows: g, cols: b }),
        K3(p, E, { rows: g, cols: b }),
        wP(p, E),
        Joe(p.delimitedSource, E),
        Ax(p, E),
        p.objects !== undefined && yC(p.objects, E),
        p.pageLayout !== undefined && Rx(p.pageLayout, E),
        E
      );
    });
  if (u.every((p) => p.hidden))
    return Hn("At least one sheet must be visible.");
  if (!Array.isArray(e.events) || e.events.length > 2e3)
    return Hn("Workbook history must be an array of at most 2,000 events.");
  const l = new Set();
  let c = 0;
  const h = e.events.map((p) => {
      if (!Vh(p)) return Hn("Invalid history event.");
      const m = Bu(p.id, "Event ID", 100),
        v = Bu(p.timestamp, "Event timestamp", 40);
      return !m || l.has(m) || !Number.isFinite(Date.parse(v))
        ? Hn("History events require unique IDs and valid timestamps.")
        : (l.add(m),
          !Array.isArray(p.changes) ||
          p.changes.length > Mt.maxPatches ||
          (c += p.changes.length) > Mt.maxCells
            ? Hn("Invalid history change count.")
            : {
                id: m,
                timestamp: v,
                actor: Bu(p.actor, "Actor", 100),
                action: Bu(p.action, "Action", 200),
                changes: p.changes.map((g) => {
                  if (
                    !Vh(g) ||
                    typeof g.sheetId != "string" ||
                    !g.sheetId.length ||
                    g.sheetId.length > 100 ||
                    typeof g.address != "string"
                  )
                    return Hn("Invalid history cell reference.");
                  const b = _t(g.address);
                  return g.address !== `${ut(b.col)}${b.row + 1}`
                    ? Hn("History cell address is not canonical.")
                    : g.before === undefined && g.after === undefined
                      ? Hn("History change must have a before or after value.")
                      : {
                          sheetId: g.sheetId,
                          address: g.address,
                          ...(g.before !== undefined
                            ? { before: U2(g.before, r) }
                            : {}),
                          ...(g.after !== undefined
                            ? { after: U2(g.after, r) }
                            : {}),
                        };
                }),
              });
    }),
    f = { schemaVersion: 1, id: t, name: n, revision: o, sheets: u, events: h };
  return (
    Lie(e, f),
    e.datePolicy !== undefined && (f.datePolicy = ts(e.datePolicy)),
    fU(e.savedFilterViews, f),
    GF(e, f),
    uoe(e, f),
    boe(e, f),
    Boe(e, f),
    Aoe(e, f),
    Roe(e, f),
    aie(e, f),
    Yoe(e, f),
    Hoe(e, f),
    qoe(e, f),
    Voe(e, f),
    Nie(e, f),
    HU(e.textCleanups, f),
    YU(e.fieldTransforms, f),
    bC(e.tableEvidence, f),
    e.whatIfAnalyses !== undefined &&
      (f.whatIfAnalyses = tae(e.whatIfAnalyses)),
    e.definedNames !== undefined && (f.definedNames = eg(e.definedNames, f)),
    jJ(e, f),
    nR(e.structureProtection, f),
    px(e.metadata, f),
    UP(f),
    EB(e, f),
    JJ(e.styleCatalog, f),
    e.formulaWatches !== undefined &&
      (f.formulaWatches = TC(e.formulaWatches, f)),
    zie(e, f),
    FJ(e, f),
    f
  );
}
function rn(e) {
  return e$(e, false);
}
function eQ(e, r, t = {}) {
  (t.readOnly && ho("This workbook is read-only."),
    (r.version !== 1 || r.snapshot !== JSON.stringify(e)) &&
      ho(
        "The workbook changed after this preview. Close and reopen Themes & Cell Styles.",
      ));
  const n = yB(e, r.request, t);
  if (
    (Ql(n) !== Ql(r) &&
      ho(
        "This theme preview changed. Create a fresh preview before applying it.",
      ),
    !r.count && Ql(r.catalogBefore) === Ql(r.catalogAfter))
  )
    return e;
  const o = Ac(e);
  r.catalogAfter
    ? (o.styleCatalog = Ac(r.catalogAfter))
    : delete o.styleCatalog;
  for (const a of r.changes) Nr(o, a.sheetId).cells[a.address] = Ac(a.after);
  return rn(o);
}
function Mw(e) {
  const r = rn(e);
  return (px(e.metadata, r), r);
}
function IB(e, r) {
  if (
    !QS(r) ||
    Object.keys(r).some((n) => !["name", "engagement", "period"].includes(n))
  )
    throw new Error("Invalid workbook property fields.");
  const t = Mw(e);
  if (
    (r.name !== undefined && (t.name = uf(r.name, "Workbook name", 200)),
    r.engagement !== undefined || r.period !== undefined)
  ) {
    if (typeof r.engagement == "string" && r.engagement.length > 200)
      throw new Error("Engagement must contain at most 200 characters.");
    if (typeof r.period == "string" && r.period.length > 100)
      throw new Error("Period must contain at most 100 characters.");
    const n = { ...(t.metadata ?? mx()) };
    (r.engagement !== undefined &&
      (typeof r.engagement == "string" && !r.engagement.trim()
        ? delete n.engagement
        : (n.engagement = uf(r.engagement, "Engagement", 200))),
      r.period !== undefined &&
        (typeof r.period == "string" && !r.period.trim()
          ? delete n.period
          : (n.period = uf(r.period, "Period", 100))),
      px(n, t));
  }
  return t;
}
function NQ(e, r) {
  return IB(e, { name: r });
}
function RQ(e) {
  const r = structuredClone(Mw(e)),
    t = new Map(r.sheets.map((a) => [a.id, crypto.randomUUID()])),
    n = new Map(
      r.sheets.flatMap((a) =>
        (a.tables ?? []).map((s) => [s.id, crypto.randomUUID()]),
      ),
    );
  ((r.id = crypto.randomUUID()), (r.revision = 0), (r.events = []));
  for (const a of r.sheets)
    if (a.analysisReport) {
      const s = a.analysisReport.request;
      if (s.kind === "statistics") s.sheetId = t.get(s.sheetId) ?? s.sheetId;
      else
        for (const d of [
          s.objective,
          ...s.variables.map((u) => u.cell),
          ...s.constraints.flatMap((u) => [
            u.left,
            ...(typeof u.right == "number" ? [] : [u.right]),
          ]),
        ])
          d.sheetId = t.get(d.sheetId) ?? d.sheetId;
      a.analysisReport.definitionHash = hr(s);
    }
  const o = [
    r.dataModel,
    ...r.sheets.map((a) => {
      return a.dataModelReport?.definition;
    }),
  ].filter(Boolean);
  for (const a of o)
    for (const s of a.tables)
      ((s.sheetId = t.get(s.sheetId) ?? s.sheetId),
        s.tableId && (s.tableId = n.get(s.tableId) ?? s.tableId));
  for (const a of r.sheets)
    a.dataModelReport &&
      (a.dataModelReport.definitionHash = hr(a.dataModelReport.definition));
  for (const a of r.savedFilterViews ?? [])
    ((a.id = crypto.randomUUID()), (a.sheetId = t.get(a.sheetId) ?? a.sheetId));
  for (const a of r.transformationRecipes ?? []) {
    a.id = crypto.randomUUID();
    for (const s of a.versions) {
      for (const d of [
        s.definition.source,
        ...s.definition.steps.flatMap((u) =>
          u.kind === "append" || u.kind === "merge" ? [u.source] : [],
        ),
      ])
        ((d.sheetId = t.get(d.sheetId) ?? d.sheetId),
          d.tableId && (d.tableId = n.get(d.tableId) ?? d.tableId));
      s.definitionHash = hr(s.definition);
    }
  }
  for (const a of r.sheets) {
    ((a.id = t.get(a.id)), a.outlineGroups && (a.outlineGroups = RB(a)));
    for (const s of a.tables ?? []) {
      s.id = n.get(s.id);
      for (const d of s.columns) d.id = crypto.randomUUID();
    }
    for (const s of [...(a.pivots ?? []), ...(a.queries ?? [])])
      ((s.id = crypto.randomUUID()),
        (s.source.sheetId = t.get(s.source.sheetId) ?? s.source.sheetId),
        s.source.tableId &&
          (s.source.tableId = n.get(s.source.tableId) ?? s.source.tableId),
        (s.sourceRevision = 0));
    for (const s of [
      ...(a.objects ?? []),
      ...(a.validations ?? []),
      ...(a.conditionalFormats ?? []),
    ])
      s.id = crypto.randomUUID();
  }
  for (const a of r.definedNames ?? [])
    a.scopeSheetId &&
      (a.scopeSheetId = t.get(a.scopeSheetId) ?? a.scopeSheetId);
  return Mw(cQ(hQ(tQ(r, crypto.randomUUID()), t), t));
}
function IQ(e, r) {
  const t = uf(r.name, "Workbook name", 200),
    n = uf(r.engagement, "Engagement", 200),
    o = uf(r.period, "Period", 100),
    a = Mw(e),
    s = RQ(a);
  return (
    (s.name = t),
    (s.metadata = mx({
      engagement: n,
      period: o,
      template: {
        id: a.id,
        name: a.name,
        version: a.metadata?.template?.version ?? 1,
        sourceWorkbookId: a.id,
        sourceRevision: a.revision,
      },
    })),
    Mw(s)
  );
}
function t5(e = "Book1") {
  return rn({
    schemaVersion: 1,
    id: crypto.randomUUID(),
    name: e.trim(),
    revision: 0,
    metadata: mx(),
    sheets: [
      {
        id: crypto.randomUUID(),
        name: "Sheet1",
        cells: {},
        rowCount: Mt.maxRows,
        colCount: Mt.maxCols,
        frozenRows: 0,
      },
    ],
    events: [],
  });
}
function $Q(e, r, t = {}) {
  if (
    r.workbookId !== e.id ||
    r.revision !== e.revision ||
    r.snapshotGuard !== JSON.stringify(e)
  )
    throw Error(
      "Conditional format preview is stale. Reopen the manager to refresh it.",
    );
  const n = LB(e, r.sheetId, r.rules, t);
  if (JSON.stringify(n.changes) !== JSON.stringify(r.changes))
    throw Error(
      "Conditional format preview was modified. Preview the rules again.",
    );
  return n.changes.length
    ? rn({
        ...e,
        sheets: e.sheets.map((o) =>
          o.id === r.sheetId ? { ...o, conditionalFormats: n.rules } : o,
        ),
      })
    : e;
}
function w0(e) {
  const r = rn(aQ(e));
  return (
    r.sheets.forEach((t, n) => wP(e.sheets[n], t)),
    nR(e.structureProtection, r),
    r
  );
}
function CP(e, r) {
  return ($a(e), w0({ ...e, sheets: [...e.sheets, Ku(rR(e, r))] }));
}
function Pne(e, r, t) {
  $a(e);
  const n = qp(e, r),
    o = rR(e, t, r);
  if (n.name === o) return e;
  const a = structuredClone(e);
  for (const s of a.sheets)
    for (const d of Object.values(s.cells))
      !d.literal &&
        typeof d.raw == "string" &&
        d.raw.startsWith("=") &&
        (d.raw = e0(d.raw, e, n, o));
  for (const s of a.sheets)
    for (const d of [...(s.conditionalFormats ?? []), ...(s.validations ?? [])])
      if ((d.kind === "formula" || d.kind === "custom") && d.formula) {
        const u = e0(d.formula, e, n, o);
        (u !== d.formula && jn(s), (d.formula = u));
      }
  for (const s of a.definedNames ?? []) s.refersTo = e0(s.refersTo, e, n, o);
  for (const s of a.sheets)
    for (const d of s.tables ?? [])
      for (const u of d.columns)
        if (u.calculatedFormula) {
          const l = e0(u.calculatedFormula, e, n, o);
          (l !== u.calculatedFormula && jn(s), (u.calculatedFormula = l));
        }
  return ((qp(a, r).name = o), w0(a));
}
function xO(e, r, t) {
  $a(e);
  const n = qp(e, r),
    o = e.sheets.indexOf(n);
  if (
    ((!Number.isInteger(t) || t < 0 || t >= e.sheets.length) &&
      ea("Choose a worksheet position inside the workbook."),
    o === t)
  )
    return e;
  EP(e, (s) => {
    vC(s).some((d) => d.sheets) &&
      ea(
        "Reordering worksheets with three-dimensional references requires explicit references first.",
      );
  });
  const a = [...e.sheets];
  return (a.splice(o, 1), a.splice(t, 0, n), w0({ ...e, sheets: a }));
}
function Une(e, r, t) {
  $a(e);
  const n = qp(e, r);
  (["visible", "hidden", "veryHidden"].includes(t) ||
    ea("Choose visible, hidden or very hidden."),
    t !== "visible" &&
      zw(n) &&
      e.sheets.filter(zw).length === 1 &&
      ea("At least one worksheet must remain visible."));
  const o = n.veryHidden ? "veryHidden" : n.hidden ? "hidden" : "visible";
  return t === o
    ? e
    : w0({
        ...e,
        sheets: e.sheets.map((a) =>
          a.id === r
            ? { ...a, hidden: t !== "visible", veryHidden: t === "veryHidden" }
            : a,
        ),
      });
}
function EO(e, r, t) {
  $a(e);
  const n = qp(e, r),
    o = t === undefined ? undefined : yP(t);
  return n.tabColor === o
    ? e
    : w0({
        ...e,
        sheets: e.sheets.map((a) => {
          if (a.id !== r) return a;
          const s = { ...a };
          return (o ? (s.tabColor = o) : delete s.tabColor, s);
        }),
      });
}
function zne(e, r, t) {
  var p, m, v;
  $a(e);
  const n = qp(e, r),
    o = rR(e, $ne(e, n.name));
  (EP(e, (g) => {
    vC(g).some((b) => b.sheets) &&
      ea(
        "Duplicating worksheets with three-dimensional references requires explicit references first.",
      );
  }),
    e.sheets.reduce((g, b) => g + Object.keys(b.cells).length, 0) +
      Object.keys(n.cells).length >
      Mt.maxCells &&
      ea(
        "Duplicating this worksheet would exceed the 100,000 populated-cell budget.",
      ));
  const s = structuredClone(n);
  ((s.id = fp()), s.outlineGroups && (s.outlineGroups = RB(s)), (s.name = o));
  const d = new Set([
      ...e.sheets.flatMap((g) =>
        (g.tables ?? []).map((b) => b.name.toLowerCase()),
      ),
      ...(e.definedNames ?? []).map((g) => g.name.toLowerCase()),
    ]),
    u = new Map(),
    l = new Map();
  for (const g of s.tables ?? []) {
    const b = g.id,
      y = g.name;
    ((g.id = fp()), (g.columns = g.columns.map((w) => ({ ...w, id: fp() }))));
    for (let w = 2; ; w++) {
      const E = `_Copy${w}`,
        A = y.slice(0, 255 - E.length) + E;
      if (!d.has(A.toLowerCase())) {
        ((g.name = A), d.add(A.toLowerCase()));
        break;
      }
    }
    (u.set(b, g.id), l.set(y, g.name));
  }
  const c = (g) => {
    let b = e0(g, e, n, o);
    for (const [y, w] of l) b = $v(b, y, w);
    return b;
  };
  for (const g of s.tables ?? [])
    for (const b of g.columns)
      b.calculatedFormula && (b.calculatedFormula = c(b.calculatedFormula));
  for (const g of Object.values(s.cells))
    !g.literal &&
      typeof g.raw == "string" &&
      g.raw.startsWith("=") &&
      (g.raw = c(g.raw));
  ((p = s.objects) == null ||
    p.forEach((g) => {
      g.id = fp();
    }),
    (m = s.validations) == null ||
      m.forEach((g) => {
        ((g.id = fp()),
          g.kind === "custom" && g.formula && (g.formula = c(g.formula)));
      }),
    (v = s.conditionalFormats) == null ||
      v.forEach((g) => {
        ((g.id = fp()),
          g.kind === "formula" && g.formula && (g.formula = c(g.formula)));
      }));
  for (const g of [...(s.pivots ?? []), ...(s.queries ?? [])])
    ((g.id = fp()),
      (g.sourceRevision = e.revision),
      g.source.sheetId === n.id &&
        ((g.source.sheetId = s.id),
        g.source.tableId &&
          (g.source.tableId =
            u.get(g.source.tableId) ??
            ea("A copied analysis refers to a missing source table."))));
  const h = (e.definedNames ?? [])
      .filter((g) => g.scopeSheetId === n.id)
      .map((g) => ({
        ...structuredClone(g),
        scopeSheetId: s.id,
        refersTo: c(g.refersTo),
      })),
    f = [...e.sheets];
  return (
    f.splice(e.sheets.indexOf(n) + 1, 0, s),
    w0(
      lQ(
        {
          ...e,
          sheets: f,
          ...(h.length
            ? { definedNames: [...(e.definedNames ?? []), ...h] }
            : {}),
        },
        n.id,
        s.id,
      ),
    )
  );
}
function Vne(e, r, t = false) {
  $a(e);
  const n = qp(e, r);
  (e.sheets.length === 1 && ea("The final worksheet cannot be deleted."),
    zw(n) &&
      e.sheets.filter(zw).length === 1 &&
      ea("At least one worksheet must remain visible."));
  const o = AP(e, r);
  (o.some((s) => s.kind === "unsupported-reference") &&
    ea(
      "Deletion is blocked by ambiguous or unsupported references. Replace dynamic or three-dimensional references with explicit references first.",
    ),
    o.length &&
      !t &&
      ea(
        "This worksheet has dependencies. Review them and explicitly confirm deletion.",
      ));
  const a = structuredClone(e);
  a.sheets = a.sheets.filter((s) => s.id !== r);
  for (const s of a.sheets) {
    for (const d of s.tables ?? [])
      for (const u of d.columns)
        if (u.calculatedFormula) {
          const l = e0(u.calculatedFormula, e, n, undefined, true);
          (l !== u.calculatedFormula && jn(s), (u.calculatedFormula = l));
        }
    for (const d of Object.values(s.cells))
      !d.literal &&
        typeof d.raw == "string" &&
        d.raw.startsWith("=") &&
        (d.raw = e0(d.raw, e, n, undefined, true));
    for (const d of [...(s.conditionalFormats ?? []), ...(s.validations ?? [])])
      if ((d.kind === "formula" || d.kind === "custom") && d.formula) {
        const u = e0(d.formula, e, n, undefined, true);
        (u !== d.formula && jn(s), (d.formula = u));
      }
    (s.pivots && (s.pivots = s.pivots.filter((d) => d.source.sheetId !== r)),
      s.queries &&
        (s.queries = s.queries.filter((d) => d.source.sheetId !== r)));
  }
  return (
    a.definedNames &&
      (a.definedNames = a.definedNames
        .filter((s) => s.scopeSheetId !== r)
        .map((s) => ({
          ...s,
          refersTo: e0(s.refersTo, e, n, undefined, true),
        }))),
    w0(fQ(a, r))
  );
}
async function Wne(e, r) {
  ($a(e), kP(r));
  const t = fp(),
    n = r
      ? { enabled: true, salt: t, passwordHash: await SP(r, t) }
      : { enabled: true };
  return w0({ ...e, structureProtection: n });
}
async function Hne(e, r) {
  kP(r);
  const t = e.structureProtection;
  if (!t) return e;
  (nR(t, { ...e }),
    t.passwordHash &&
      (await SP(r, t.salt)) !== t.passwordHash.toLowerCase() &&
      ea("The workbook structure password does not match."));
  const n = { ...e };
  return (delete n.structureProtection, w0(n));
}
function t$(e) {
  Lg(e);
  for (const r of e.sheets) {
    if (YE.has(r)) continue;
    const t = Object.values(r.cells);
    for (const n of t) JU.add(n);
    (QU.set(r.cells, t.length),
      t.some((n) => {
        return n.hyperlink || n.style?.binding;
      }) || YE.set(r, t.length));
  }
  return (NC.add(e), e);
}

export {
  $Q,
  CP,
  EO,
  Eie,
  HU,
  Hne,
  IB,
  IQ,
  Ku,
  NO,
  NQ,
  Pne,
  QP,
  RP,
  Une,
  Vne,
  Wne,
  YU,
  ZP,
  Zne,
  bie,
  e$,
  eQ,
  fie,
  hi,
  mie,
  qQ,
  rn,
  t$,
  t5,
  xO,
  yR,
  yie,
  zne,
};
