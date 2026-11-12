import { Mt, Xy } from "../02-financial-functions.js";
import { qj, L3 } from "../01-core-schemas-and-hashing.js";
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
  AB,
  Ad,
  BB,
  E1,
  F1,
  FQ,
  G3,
  Ip,
  J3,
  Jf,
  Jl,
  Lv,
  MN,
  NB,
  PQ,
  TB,
  TQ,
  UQ,
  Ub,
  Xm,
  YS,
  Z3,
  a1,
  al,
  cB,
  cn,
  fx,
  gg,
  ji,
  jn,
  kc,
  m6,
  mQ,
  n5,
  s0,
  sD,
  ul,
  vs,
  vx,
  w1,
  x1,
  yv,
} from "./01-cell-styles-and-themes.js";
import { Rc, Us, WL } from "./02-number-formats.js";
import { Pi, Uw, Vg, Yy, b5, eg, iP, nP, pC, une } from "./03-defined-names.js";
import { Iie, P1, _ie, iU, ip, oU, z4 } from "./04-pivot-tables-and-pictures.js";
import {
  B2,
  Ep,
  HO,
  Hn,
  JR,
  Ls,
  PU,
  Qoe,
  Vu,
  WO,
  Yh,
  dk,
  e1,
  l1,
  qE,
  ut,
  wg,
  xd,
} from "./06-data-model-and-analysis.js";

function Die(e, r = {}) {
  if (!ip(r) || Object.keys(r).some((u) => u !== "filteredRowsBySheet"))
    throw Error("Calculation options support only filteredRowsBySheet.");
  const t = r.filteredRowsBySheet;
  if (t !== void 0 && !ip(t))
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
    a.set(u, ul(c.rowCount, Ad(l, c.rowCount)));
  }
  const d = new Map();
  for (const u of e.sheets) d.set(u.id, ul(u.rowCount, cn(u, "row")));
  return (u, l, c) => {
    var H, Q, X;
    if (!iU.includes(u)) return oU;
    const { issue: h, isIssue: f } = c,
      p = (J) => {
        if (f(J)) throw J;
        if (!ip(J) || !ip(J.ref)) throw h("#VALUE!");
        const q = J.ref;
        if (q.sheet !== void 0 && typeof q.sheet != "string") throw h("#REF!");
        const W = o.get(
          String(q.sheet ?? c.currentSheetName ?? "").toLowerCase(),
        );
        if (!W) throw h("#REF!");
        const K = ip(q.from) ? q.from : q,
          te = ip(q.to) ? q.to : q,
          ae = [K.row, K.col, te.row, te.col];
        if (ae.some((le) => !Number.isInteger(le))) throw h("#REF!");
        const [oe, de, he, Se] = ae;
        if (
          oe < 1 ||
          de < 1 ||
          he < oe ||
          Se < de ||
          he > W.rowCount ||
          Se > W.colCount
        )
          throw h("#REF!");
        return {
          sheet: W,
          firstRow: oe,
          lastRow: he,
          firstCol: de,
          lastCol: Se,
        };
      };
    if (u === "FORMULATEXT" || u === "ISFORMULA") {
      if (l.length !== 1) return h("#VALUE!");
      const J = p(l[0]),
        q = (K, te) => {
          var oe;
          const ae = `${ut(te - 1)}${K}`;
          return (
            (oe = c.readMetadata) == null || oe.call(c, J.sheet.id, ae),
            J.sheet.cells[ae]
          );
        };
      if (u === "FORMULATEXT") {
        const K = q(J.firstRow, J.firstCol);
        return z4(K) && K.raw.length <= 8192 ? K.raw : h("#N/A");
      }
      const W = (J.lastRow - J.firstRow + 1) * (J.lastCol - J.firstCol + 1);
      return W > 1e4
        ? h("#LIMIT!")
        : W === 1
          ? z4(q(J.firstRow, J.firstCol))
          : Array.from({ length: J.lastRow - J.firstRow + 1 }, (K, te) =>
              Array.from({ length: J.lastCol - J.firstCol + 1 }, (ae, oe) =>
                z4(q(J.firstRow + te, J.firstCol + oe)),
              ),
            );
    }
    const m = u === "AGGREGATE";
    if (l.length < (m ? 3 : 2) || l.length > 255) return h("#VALUE!");
    const v = (J) => {
        const q = c.retrieve(J);
        if (f(q)) throw q;
        if (
          Array.isArray(q) ||
          !["number", "string", "boolean"].includes(typeof q) ||
          (typeof q == "string" && q.trim() === "") ||
          !Number.isFinite(Number(q))
        )
          throw h("#VALUE!");
        return Number(q);
      },
      g = v(l[0]);
    if (m && !Number.isInteger(g)) return h("#VALUE!");
    const b = Math.trunc(g),
      y = !m && b > 100 ? b - 100 : b;
    if (y < 1 || y > (m ? 19 : 11)) return h("#VALUE!");
    const w = m ? (l[1] == null ? 0 : v(l[1])) : 0;
    if (!Number.isInteger(w) || w < 0 || w > 7) return h("#VALUE!");
    const E = m ? w % 2 === 1 : b > 100,
      A = m && [2, 3, 6, 7].includes(w),
      k = !m || w < 4,
      F = m && y >= 14;
    if (F && (l.length !== 4 || l[3] == null)) return h("#VALUE!");
    const I = F ? v(l[3]) : 0,
      O = F ? [l[2]] : l.slice(m ? 2 : 1),
      z = F && !((J) => ip(J) && ip(J.ref))(O[0]),
      V = z ? [] : O.map(p);
    let D = 0;
    for (const J of V) {
      const q = (J.lastRow - J.firstRow + 1) * (J.lastCol - J.firstCol + 1);
      if (((D += q), q > 2e4 || D > 1e5)) return h("#LIMIT!");
    }
    for (const J of V)
      for (let q = J.firstRow; q <= J.lastRow; q++)
        for (let W = J.firstCol; W <= J.lastCol; W++)
          if (c.isActive(J.sheet.id, `${ut(W - 1)}${q}`)) return h("#CYCLE!");
    let B = 0,
      _ = 0,
      C = 1 / 0,
      T = -1 / 0,
      N = 0,
      R = 0;
    const L = [],
      $ = (J) => {
        if (f(J)) {
          if (["#LIMIT!", "#CYCLE!"].includes(J.error)) return J;
          if (A) return;
          if (y !== 2 && y !== 3) return J;
        }
        if ((J != null && _++, typeof J != "number")) return;
        if (!Number.isFinite(J)) return h("#NUM!");
        (B++, (C = Math.min(C, J)), (T = Math.max(T, J)));
        const q = J - N;
        ((N += q / B),
          (R += q * (J - N)),
          (y === 1 || y === 6 || y === 9 || y >= 12) && L.push(J));
      };
    if (z) {
      const J = c.retrieve(O[0]),
        q = Array.isArray(J) ? J : [[J]];
      if (!q.length || q.length > 1e4) return h("#LIMIT!");
      const W = q.every(Array.isArray) ? q : [q],
        K = ((H = W[0]) == null ? void 0 : H.length) ?? 0;
      if (!K || W.some((te) => te.length !== K || te.some(Array.isArray)))
        return h("#VALUE!");
      if (W.length * K > 1e4) return h("#LIMIT!");
      for (const te of W)
        for (const ae of te) {
          if (
            ae != null &&
            !f(ae) &&
            !["number", "string", "boolean"].includes(typeof ae)
          )
            return h("#VALUE!");
          const oe = $(ae);
          if (oe) return oe;
        }
    }
    for (const J of V)
      for (let q = J.firstRow; q <= J.lastRow; q++)
        if (!(
          ((Q = a.get(J.sheet.id)) != null && Q.isHidden(q - 1)) ||
          (E && d.get(J.sheet.id).isHidden(q - 1))
        ))
          for (let W = J.firstCol; W <= J.lastCol; W++) {
            const K = `${ut(W - 1)}${q}`;
            (X = c.readMetadata) == null || X.call(c, J.sheet.id, K);
            const te = J.sheet.cells[K];
            if (
              k &&
              z4(te) &&
              (m ? _ie : Iie)(c.expandedFormula(te, J.sheet.id))
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
        return oe > 1 ? de : h("#N/A");
      }
      if (!L.length) return h("#NUM!");
      if (y === 12) {
        const oe = Math.floor(L.length / 2);
        return L.length % 2 ? L[oe] : L[oe - 1] / 2 + L[oe] / 2;
      }
      if (y === 14 || y === 15) {
        const oe = I;
        return Number.isInteger(oe) && oe >= 1 && oe <= L.length
          ? L[y === 14 ? L.length - oe : oe - 1]
          : h("#NUM!");
      }
      const J = y >= 18,
        q = y === 17 || y === 19 ? Math.trunc(I) / 4 : I;
      if (q < 0 || q > 1 || (J && (q === 0 || q === 1))) return h("#NUM!");
      const W = J ? q * (L.length + 1) - 1 : q * (L.length - 1);
      if (W < 0 || W > L.length - 1) return h("#NUM!");
      const K = Math.floor(W),
        te = Math.ceil(W),
        ae = W - K;
      return L[K] * (1 - ae) + L[te] * ae;
    }
    if (y === 2) return B;
    if (y === 3) return _;
    if (y === 9) return Rc(L);
    if (y === 1) return B ? Us(L) : h("#DIV/0!");
    if (y === 4) return B ? T : 0;
    if (y === 5) return B ? C : 0;
    if (y === 6) return B ? WL(L) : 0;
    const U = y === 7 || y === 10;
    if (B < (U ? 2 : 1)) return h("#DIV/0!");
    const G = Math.max(0, R / (B - (U ? 1 : 0)));
    return y === 7 || y === 8 ? Math.sqrt(G) : G;
  };
}
function G5(e, r) {
  return r === "row"
    ? `${e.start + 1}:${e.end + 1}`
    : `${ut(e.start)}:${ut(e.end)}`;
}
function _t(e) {
  const r = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(e);
  if (!r) return Hn(`Invalid cell address: ${e}`);
  let t = 0;
  for (const o of r[1].toUpperCase()) t = t * 26 + o.charCodeAt(0) - 64;
  const n = Number(r[2]) - 1;
  return n >= Mt.maxRows || t > Mt.maxCols
    ? Hn(`Cell ${e} exceeds the supported workbook bounds.`)
    : { row: n, col: t - 1 };
}
function pB(e, r, t, n = { style: "thin", color: "#000000" }) {
  const o = _t(e),
    a = cB(n),
    s = {};
  ["all", "outline", "inside", "clear"].includes(t) ||
    vs("Choose a supported border preset.");
  for (const d of s0) {
    const u = o.row + (d === "top" ? -1 : d === "bottom" ? 1 : 0),
      l = o.col + (d === "left" ? -1 : d === "right" ? 1 : 0),
      c =
        u >= 0 &&
        u < 1048576 &&
        l >= 0 &&
        l < 16384 &&
        r.has(`${ut(l)}${u + 1}`);
    t === "clear"
      ? (s[d] = { style: "none" })
      : (t === "all" || (t === "outline" ? !c : c)) && (s[d] = { ...a });
  }
  return s;
}
function q3(e) {
  if (typeof e != "string") return ji("Choose a valid worksheet cell address.");
  const r = _t(e.trim());
  return `${ut(r.col)}${r.row + 1}`;
}
function X3(e) {
  const r = Xm(e, "Hyperlink"),
    t =
      r.label === void 0 || r.label === ""
        ? void 0
        : kc(r.label, "Link display text", Jl.labelLength);
  if (r.kind === "internal")
    return {
      kind: "internal",
      sheetId: kc(r.sheetId, "Link worksheet ID", 100),
      address: q3(r.address),
      ...(t ? { label: t } : {}),
    };
  if (r.kind !== "external")
    return ji("Choose a web, email or worksheet link.");
  const n = kc(r.url, "Link address", Jl.urlLength);
  let o = n;
  for (let s = 0; s < 4; s++) {
    (YS.test(o) || /\\/.test(o)) &&
      ji("Link addresses cannot contain control characters or backslashes.");
    let d = o;
    try {
      d = decodeURIComponent(o);
    } catch {
      ji("The link contains invalid encoded characters.");
    }
    if (d === o) break;
    o = d;
  }
  (YS.test(o) || /%(?:0[\da-f]|1[\da-f]|7f|8[\da-f]|9[\da-f])/i.test(o)) &&
    ji("Link addresses cannot contain encoded control characters.");
  let a;
  try {
    a = new URL(n);
  } catch {
    return ji("Enter a complete https://, http:// or mailto: address.");
  }
  return (
    ["https:", "http:", "mailto:"].includes(a.protocol) ||
      ji("Only https://, http:// and mailto: links are supported."),
    (a.username || a.password) &&
      ji("Links cannot contain embedded credentials."),
    a.protocol === "mailto:"
      ? (a.search ||
          a.hash ||
          a.host ||
          !/^[^\s@/?#:,]+@[^\s@/?#:,]+\.[^\s@/?#:,]+$/.test(
            decodeURIComponent(a.pathname),
          )) &&
        ji("Use mailto: with one email address and no additional headers.")
      : a.hostname || ji("Web links require a hostname."),
    a.href.length > Jl.urlLength && ji("The normalized link is too long."),
    { kind: "external", url: a.href, ...(t ? { label: t } : {}) }
  );
}
function EB(e, r) {
  const t = Xm(e, "Workbook review metadata");
  if (t.reviewThreads !== void 0) {
    (!Array.isArray(t.reviewThreads) || t.reviewThreads.length > Jl.threads) &&
      ji(`A workbook can contain at most ${Jl.threads} review threads.`);
    const n = new Set(),
      o = new Set();
    let a = 0;
    const s = t.reviewThreads.map((d) => {
      const u = Xm(d, "Review thread"),
        l = Xm(u.anchor, "Review target"),
        c = kc(u.id, "Review thread ID", 100);
      (n.has(c) && ji("Review thread IDs must be unique."), n.add(c));
      const h = Lv(u.createdAt),
        f = Lv(u.updatedAt);
      f < h && ji("Review update time cannot precede creation.");
      const p = {
        sheetId: kc(l.sheetId, "Review worksheet ID", 100),
        sheetName: kc(l.sheetName, "Review worksheet name", 31),
        address: q3(l.address),
      };
      (!Array.isArray(u.messages) ||
        !u.messages.length ||
        u.messages.length > Jl.messagesPerThread ||
        (a += u.messages.length) > Jl.totalMessages) &&
        ji(
          "Review discussions exceed the per-thread or workbook message limit.",
        );
      let m = h;
      const v = u.messages.map((b, y) => {
          const w = Xm(b, "Review message"),
            E = kc(w.id, "Review message ID", 100);
          (o.has(E) &&
            ji("Review message IDs must be unique within the workbook."),
            o.add(E));
          const A = Lv(w.createdAt);
          return (
            (A < m || A > f || (y === 0 && A !== h)) &&
              ji("Review message times must follow their discussion in order."),
            (m = A),
            {
              id: E,
              author: w1(kc(w.author, "Reviewer name", 100)),
              text: kc(w.text, "Comment", Jl.textLength, true),
              createdAt: A,
            }
          );
        }),
        g = { id: c, anchor: p, messages: v, createdAt: h, updatedAt: f };
      if (u.resolved !== void 0) {
        const b = Xm(u.resolved, "Review resolution"),
          y = Lv(b.createdAt);
        ((y < v.at(-1).createdAt || y > f) &&
          ji("Invalid review resolution time."),
          (g.resolved = {
            author: w1(kc(b.author, "Reviewer name", 100)),
            createdAt: y,
          }));
      }
      if (u.orphaned !== void 0) {
        const b = Xm(u.orphaned, "Orphaned review target"),
          y = Lv(b.createdAt);
        ((!["worksheet-deleted", "cell-deleted"].includes(String(b.reason)) ||
          y < h ||
          y > f) &&
          ji("Invalid orphaned review target."),
          (g.orphaned = { reason: b.reason, createdAt: y }));
      }
      return (
        !g.orphaned &&
          !r.sheets.some((b) => b.id === p.sheetId) &&
          ji(
            "A review thread with a missing worksheet must be retained as orphaned.",
          ),
        g
      );
    });
    s.length ? (r.reviewThreads = s) : delete r.reviewThreads;
  }
  for (const n of r.sheets)
    for (const o of qj(n.cells) ? Object.values(n.cells) : []) {
      if (o.hyperlink === void 0) continue;
      const a = X3(o.hyperlink);
      (a.kind === "internal" &&
        !r.sheets.some((s) => s.id === a.sheetId) &&
        ji("An internal hyperlink must target an existing worksheet."),
        (o.hyperlink = a));
    }
}
function ON(e, r) {
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
  return (EB(t, n), n);
}
function sQ(e, r, t, n) {
  var l;
  const o = t.axis === "row" ? Mt.maxRows : Mt.maxCols;
  (!["row", "column"].includes(t.axis) ||
    !Number.isInteger(t.index) ||
    t.index < 0 ||
    !Number.isInteger(t.count) ||
    t.count < 1 ||
    t.index + t.count > o) &&
    ji("Choose a valid row or column change for review targets.");
  const a = (c) => {
    const h = _t(c),
      f = t.axis === "row" ? h.row : h.col;
    if (t.moveBefore === void 0 && f < t.index) return c;
    if (t.moveBefore === void 0 && t.remove && f < t.index + t.count)
      return null;
    const p =
      t.moveBefore !== void 0 ? Jf(f, t) : f + (t.remove ? -t.count : t.count);
    return (
      p >= o &&
        ji(
          "Insertion would discard a review target or hyperlink at the worksheet boundary.",
        ),
      t.axis === "row" ? `${ut(h.col)}${p + 1}` : `${ut(p)}${h.row + 1}`
    );
  };
  let s = false;
  const d =
      (l = e.reviewThreads) == null
        ? void 0
        : l.map((c) => {
            if (c.orphaned || c.anchor.sheetId !== r) return c;
            const h = a(c.anchor.address);
            if (h === c.anchor.address) return c;
            if (((s = true), h))
              return { ...c, anchor: { ...c.anchor, address: h } };
            const f = fx({ timestamp: n }, c.updatedAt);
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
        if ((m == null ? void 0 : m.kind) !== "internal" || m.sheetId !== r)
          continue;
        const v = a(m.address);
        v !== m.address &&
          (h ?? (h = { ...c.cells }),
          (h[f] = { ...p }),
          (s = true),
          v ? (h[f].hyperlink = { ...m, address: v }) : delete h[f].hyperlink);
      }
      return h ? { ...c, cells: h } : c;
    });
  return s ? { ...e, sheets: u, ...(d ? { reviewThreads: d } : {}) } : e;
}
function dQ(e, r, t) {
  var n;
  return (n = e.formulaWatches) != null && n.length
    ? {
        ...e,
        formulaWatches: e.formulaWatches.map((o) => {
          if (o.orphaned || o.anchor.sheetId !== r) return o;
          const a = _t(o.anchor.address),
            s = t.axis === "row" ? a.row : a.col;
          if (
            t.moveBefore === void 0 &&
            t.remove &&
            s >= t.index &&
            s < t.index + t.count
          )
            return { ...o, orphaned: "deleted-cells" };
          const d =
            t.moveBefore !== void 0
              ? Jf(s, t)
              : s >= t.index + (t.remove ? t.count : 0)
                ? s + (t.remove ? -t.count : t.count)
                : s;
          if (d < 0 || d >= (t.axis === "row" ? Mt.maxRows : Mt.maxCols))
            return { ...o, orphaned: "out-of-bounds" };
          const u =
            t.axis === "row" ? `${ut(a.col)}${d + 1}` : `${ut(d)}${a.row + 1}`;
          return u === o.anchor.address
            ? o
            : { ...o, anchor: { ...o.anchor, address: u } };
        }),
      }
    : e;
}
function Qne(e, r) {
  const t = _t(r.split(":")[0]);
  return {
    id: crypto.randomUUID(),
    kind: e,
    anchor: `${ut(t.col)}${t.row + 1}`,
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
function eie(e, r, t) {
  var o, a;
  const n = [];
  for (const s of t.objects ?? []) {
    if (s.kind !== "picture") {
      n.push(
        `${s.label || s.kind}: native chart/shape export is not supported; retain the JSON workbook.`,
      );
      continue;
    }
    const { extension: d } = P1(s.dataURL);
    s.evidence &&
      n.push(
        `${s.label || "Picture"}: XLSX retains the image bytes and placement, but original filename, source reference and evidence provenance require native JSON; a later XLSX import records only observed media.`,
      );
    const u = _t(s.anchor),
      l = e.addImage({ base64: s.dataURL, extension: d });
    r.addImage(l, {
      tl: {
        col:
          u.col +
          s.offsetX /
            (((o = t.columnWidths) == null ? void 0 : o[u.col]) ?? 84),
        row:
          u.row +
          s.offsetY / (((a = t.rowHeights) == null ? void 0 : a[u.row]) ?? 21),
      },
      ext: { width: s.width, height: s.height },
      editAs: "oneCell",
    });
  }
  return n;
}
function eae(e) {
  if (typeof e != "string" || !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}$/.test(e))
    throw Error("A formula watch must identify one valid Excel cell address.");
  const r = _t(e.replace(/\$/g, "").toUpperCase());
  if (r.row >= Mt.maxRows || r.col >= Mt.maxCols)
    throw Error("A formula watch address is outside Excel's cell bounds.");
  return `${ut(r.col)}${r.row + 1}`;
}
function TC(e, r) {
  if (e === void 0) return [];
  if (!Array.isArray(e) || e.length > B2)
    throw Error("A workbook may contain at most 100 formula watches.");
  const t = new Set(),
    n = new Set();
  return e.map((o) => {
    if (
      !WO(o) ||
      !HO(o, ["id", "anchor", "orphaned"]) ||
      !dk(o.id, 100) ||
      !WO(o.anchor) ||
      !HO(o.anchor, ["sheetId", "sheetName", "address"]) ||
      !dk(o.anchor.sheetId, 100) ||
      !dk(o.anchor.sheetName, 31)
    )
      throw Error("Invalid formula watch metadata.");
    if (t.has(o.id)) throw Error("Formula watch IDs must be unique.");
    if ((t.add(o.id), o.orphaned !== void 0 && !Qoe.includes(o.orphaned)))
      throw Error("Invalid formula watch orphan reason.");
    const a = o.anchor,
      s = eae(a.address),
      d = r.sheets.find((c) => c.id === a.sheetId),
      u = o.orphaned ?? (d ? void 0 : "missing-sheet"),
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
function BU(e) {
  const r = Yh(e, ["sheetId", "address"]),
    t = Ep(r.sheetId);
  if (typeof r.address != "string")
    throw Error("A what-if cell needs an Excel address.");
  const n = _t(r.address);
  return { sheetId: t, address: `${ut(n.col)}${n.row + 1}` };
}
function qO(e, r) {
  const t = l1(e, r).map(BU);
  if (new Set(t.map(Vu)).size !== t.length)
    throw Error("What-if cell references must not repeat.");
  return t;
}
function LU(e) {
  const r = Yh(e, [
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
      name: Ep(r.name),
      inputs: qO(r.inputs, Ls.inputs),
      outputs: qO(r.outputs, Ls.outputs),
    };
  if (t.inputs.some((o) => t.outputs.some((a) => Vu(o) === Vu(a))))
    throw Error("Changing inputs and formula results must be different cells.");
  const n = ["name", "kind", "inputs", "outputs"];
  if (r.kind === "scenarios") {
    Yh(r, [...n, "scenarios"]);
    const o = l1(r.scenarios, Ls.scenarios).map((a) => {
      const s = Yh(a, ["name", "values"]),
        d = e1(s.values, Ls.inputs);
      if (d.length !== t.inputs.length)
        throw Error(
          "Each scenario needs exactly one value per changing input.",
        );
      return { name: Ep(s.name), values: d };
    });
    if (new Set(o.map((a) => a.name.toLowerCase())).size !== o.length)
      throw Error("Scenario names must be unique, ignoring case.");
    return { ...t, kind: r.kind, scenarios: o };
  }
  if (r.kind === "one-variable") {
    if (
      (Yh(r, [...n, "orientation", "values"]),
      t.inputs.length !== 1 || !["row", "column"].includes(r.orientation))
    )
      throw Error(
        "A one-variable table needs one input and a row or column orientation.",
      );
    return {
      ...t,
      kind: r.kind,
      orientation: r.orientation,
      values: e1(r.values, Ls.axisValues),
    };
  }
  if (r.kind === "two-variable") {
    if (
      (Yh(r, [...n, "rowValues", "columnValues"]),
      t.inputs.length !== 2 || t.outputs.length !== 1)
    )
      throw Error(
        "A two-variable table needs two different inputs and exactly one formula result.",
      );
    const o = e1(r.rowValues, Ls.axisValues),
      a = e1(r.columnValues, Ls.axisValues);
    if (o.length * a.length > Ls.trials)
      throw Error("A sensitivity table supports at most 400 combinations.");
    return { ...t, kind: r.kind, rowValues: o, columnValues: a };
  }
  throw Error(
    "Choose scenarios, a one-variable table or a two-variable table.",
  );
}
function tae(e) {
  const r = new Set();
  return l1(e, Ls.savedAnalyses, 0).map((t) => {
    const n = Yh(t, [
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
      o = Ep(n.id),
      a = Ep(n.reportSheetId),
      s = LU(n.request);
    if (r.has(o)) throw Error("Saved what-if IDs must be unique.");
    r.add(o);
    const d = (p) => {
        const m = Yh(p, ["name", "inputs", "outputs", "error"]),
          v = Ep(m.name, 200),
          g = e1(m.inputs, Ls.inputs),
          b = m.outputs === null ? null : e1(m.outputs, Ls.outputs),
          y = m.error === null ? null : Ep(m.error, 2e3);
        if (
          g.length !== s.inputs.length ||
          (b && b.length !== s.outputs.length) ||
          !!b == !!y
        )
          throw Error("Invalid saved what-if result dimensions or outcome.");
        return { name: v, inputs: g, outputs: b, error: y };
      },
      u = d(n.baseline),
      l = l1(n.trials, Ls.trials).map(d),
      c = PU(s);
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
      f = l1(n.assumptions, Ls.dependencyCells).map((p) => {
        const m = Yh(p, ["sheetId", "sheetName", "address", "raw", "literal"]),
          v = BU({ sheetId: m.sheetId, address: m.address });
        if (h.has(Vu(v)))
          throw Error("Saved assumptions must not repeat a source cell.");
        if (
          (h.add(Vu(v)),
          !(
            m.raw === null ||
            typeof m.raw == "boolean" ||
            qE(m.raw) ||
            (typeof m.raw == "string" && m.raw.length <= 32767)
          ) ||
            (m.literal !== void 0 && typeof m.literal != "boolean"))
        )
          throw Error("Invalid saved what-if assumption.");
        return {
          ...v,
          sheetName: Ep(m.sheetName, 31),
          raw: m.raw,
          ...(m.literal === void 0 ? {} : { literal: m.literal }),
        };
      });
    if (
      [...s.inputs, ...s.outputs].some((p) => !h.has(Vu(p))) ||
      s.inputs.some((p, m) => {
        var v;
        return (
          ((v = f.find((g) => Vu(g) === Vu(p))) == null ? void 0 : v.raw) !==
          u.inputs[m]
        );
      })
    )
      throw Error(
        "Saved baseline inputs and dependency assumptions do not reconcile.",
      );
    for (const [p, m] of [
      ["sourceRevision", Number.MAX_SAFE_INTEGER - 1],
      ["dependencyCells", Ls.dependencyCells],
      ["workUnits", Ls.workUnits],
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
function XR(e, r) {
  if (r.kind === "chart" && r.chart === "sparkline") {
    const t = _t(r.anchor);
    return {
      offsetX: 2,
      offsetY: 2,
      width: Ip(e, t.col) - 4,
      height: gg(e, t.row) - 4,
    };
  }
  return {
    offsetX: r.offsetX,
    offsetY: r.offsetY,
    width: r.width,
    height: r.height,
  };
}
function Fg(e, r, t) {
  if (e === void 0 || e === "") return;
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
    d = r === "row" ? wg.titleRows : wg.titleColumns;
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
function iae(e, r, t, n, o, a) {
  const s = t === "row" ? "repeatRows" : "repeatColumns",
    d = t === "row" ? "rowBreaks" : "columnBreaks",
    u = { ...e },
    l = Fg(e[s], t, r);
  if (l) {
    let { start: f, end: p } = l;
    if (!a) (n <= f && (f += o), n <= p && (p += o));
    else {
      const m = n + o - 1,
        v = [
          f < n ? { start: f, end: Math.min(p, n - 1) } : void 0,
          p > m ? { start: Math.max(f, m + 1) - o, end: p - o } : void 0,
        ].filter(Boolean);
      if (!v.length)
        throw Error(
          "This edit removes every repeated print title. Clear or change print titles first.",
        );
      ((f = v[0].start), (p = v.at(-1).end));
    }
    ((u[s] = G5({ start: f, end: p }, t)), Fg(u[s], t, r));
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
function GO(e) {
  if (typeof e != "string") return false;
  try {
    const r = _t(e);
    return `${ut(r.col)}${r.row + 1}` === e;
  } catch {
    return false;
  }
}
const KE = string().refine((e) => {
    try {
      const r = _t(e);
      return `${ut(r.col)}${r.row + 1}` === e;
    } catch {
      return false;
    }
  }, "Invalid evidence address."),
  hae = strictObject({
    id: string().min(1).max(100),
    sourceId: string().min(1).max(100),
    sourceName: string().min(1).max(255),
    sourceSheet: string().min(1).max(31),
    originalRow: number().int().min(1).max(1048576),
    kind: _enum(["file", "snapshot"]),
  }),
  pae = strictObject({
    address: KE,
    original: unknown().optional(),
    origin: hae.optional(),
  }),
  mae = strictObject({
    address: KE,
    row: number().int().min(1).max(1048576),
    field: string().min(1).max(150),
    value: JR,
    type: _enum(["text", "number", "boolean", "date", "blank"]),
    sources: array(KE).min(1).max(500),
    exception: boolean(),
  }),
  gae = strictObject({
    step: number().int().min(1).max(12),
    row: number().int().min(1).max(1048576),
    field: string().min(1).max(150),
    raw: JR,
    sources: array(KE).min(1).max(500),
    reason: string().min(1).max(1e3),
  }),
  vae = strictObject({
    id: string().min(1).max(100),
    workbookId: string().min(1).max(100),
    sourceRevision: number().int().nonnegative(),
    sourceSheetId: string().min(1).max(100),
    sourceSheetName: string().min(1).max(31),
    outputSheetId: string().min(1).max(100),
    outputSheetName: string().min(1).max(31),
    createdAt: L3(),
    semantics: literal("reviewed-fields-v1"),
    request: unknown(),
    sources: array(pae).min(1).max(500),
    outputs: array(mae).min(1).max(5e3),
    exceptions: array(gae).max(5e3),
  });
function ct(e) {
  const r = e.trim().split(":");
  if (r.length > 2) return Hn(`Invalid range: ${e}`);
  const t = _t(r[0]),
    n = _t(r[1] ?? r[0]);
  return {
    firstRow: Math.min(t.row, n.row),
    lastRow: Math.max(t.row, n.row),
    firstCol: Math.min(t.col, n.col),
    lastCol: Math.max(t.col, n.col),
  };
}
function ff(e) {
  const r = typeof e == "string" ? e.split(",") : [...e];
  if (!r.length || r.length > Xy)
    throw Error(`Select between 1 and ${Xy} rectangular areas.`);
  const t = r.map((n) => {
    const o = ct(n.trim().toUpperCase().replace(/\$/g, "")),
      a = `${ut(o.firstCol)}${o.firstRow + 1}`,
      s = `${ut(o.lastCol)}${o.lastRow + 1}`;
    return a === s ? a : `${a}:${s}`;
  });
  return [...new Set(t)];
}
function W3(e, r) {
  const t = _t(r);
  return e.some((n) => {
    const o = ct(n);
    return (
      t.row >= o.firstRow &&
      t.row <= o.lastRow &&
      t.col >= o.firstCol &&
      t.col <= o.lastCol
    );
  });
}
function y1(e) {
  const r = ff(e).map(ct),
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
function Np(e, r = Mt.maxPatches) {
  if (!Number.isSafeInteger(r) || r < 1 || r > Mt.maxCells)
    throw Error("Invalid selection enumeration limit.");
  const t = ff(e),
    n = y1(t);
  if (n > r)
    throw Error(
      `Select at most ${r.toLocaleString()} cells for this operation. The current selection contains ${n.toLocaleString()} cells.`,
    );
  const o = new Set();
  for (const a of t) {
    const s = ct(a);
    for (let d = s.firstRow; d <= s.lastRow; d++)
      for (let u = s.firstCol; u <= s.lastCol; u++) o.add(`${ut(u)}${d + 1}`);
  }
  return [...o];
}
function pQ(e, r, t, n, o) {
  var h;
  const a = ct(t),
    s = ct(o),
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
    watches:
      (h = e.formulaWatches) == null
        ? void 0
        : h.map((f) => {
            if (f.orphaned) return f;
            const p = `${f.anchor.sheetName}!${f.anchor.address}`;
            if (f.anchor.sheetId === r && d(a, f.anchor.address)) {
              const m = _t(f.anchor.address),
                v = `${ut(m.col + s.firstCol - a.firstCol)}${m.row + s.firstRow - a.firstRow + 1}`,
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
function K3(e, r, t = { rows: r.rowCount, cols: r.colCount }) {
  mQ(e, r);
  const n = AB(e.splitPanes, t);
  if (n && (e.frozenRows || e.frozenCols))
    throw Error("Split panes and frozen panes cannot be enabled together.");
  n && (r.splitPanes = n);
  const o = Ad(e.hiddenRows, t.rows),
    a = Ad(e.hiddenCols, t.cols),
    s = { ...r, hiddenRows: o, hiddenCols: a },
    d = ul(r.rowCount, cn(s, "row")),
    u = ul(r.colCount, cn(s, "column"));
  if (!d.count || !u.count)
    throw Error("At least one row and one column must remain visible.");
  const l = e.frozenRows === void 0 ? 0 : e.frozenRows,
    c = e.frozenCols === void 0 ? 0 : e.frozenCols;
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
  if (h > x1.maxFrozenVisibleRows || f > x1.maxFrozenVisibleCols)
    throw Error(
      "Freeze panes supports at most 100 visible rows and 50 visible columns.",
    );
  if (h === d.count || f === u.count)
    throw Error("Freeze panes must leave an unfrozen visible row and column.");
  for (const p of r.merges ?? []) {
    const m = ct(p);
    if (
      (m.firstRow < l && m.lastRow >= l) ||
      (m.firstCol < c && m.lastCol >= c)
    )
      throw Error(
        "Freeze panes cannot cross a merged range. Unmerge those cells first.",
      );
  }
  (e.hiddenRows !== void 0 && (r.hiddenRows = o),
    e.hiddenCols !== void 0 && (r.hiddenCols = a),
    e.frozenRows !== void 0 && (r.frozenRows = l),
    e.frozenCols !== void 0 && (r.frozenCols = c));
}
function Y3(e, r, t) {
  const n = e.sheets.find((a) => a.id === r);
  if (!n) throw Error("Worksheet not found.");
  jn(n);
  const o = t(structuredClone(n));
  return (
    K3(o, o),
    { ...e, sheets: e.sheets.map((a) => (a.id === r ? o : a)) }
  );
}
function d6(e, r, t, n, o, a) {
  if ((Z3(t), typeof a != "boolean"))
    throw Error("Hidden state must be boolean.");
  return Y3(e, r, (s) => {
    const d = MN(s, t);
    Ad([{ start: n, end: o }], d);
    const u = E1(s, t),
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
    return ((s[G3(t)] = Ad(h, d)), s);
  });
}
function yQ(e, r, t) {
  return (Z3(t), Y3(e, r, (n) => (delete n[G3(t)], n)));
}
function Gy(e, r, t, n) {
  return Y3(
    e,
    r,
    (o) => (delete o.splitPanes, { ...o, frozenRows: t, frozenCols: n }),
  );
}
function JS(e, r, t) {
  return Y3(
    e,
    r,
    (n) => (
      t === void 0
        ? delete n.splitPanes
        : ((n.splitPanes = AB(t, { rows: n.rowCount, cols: n.colCount })),
          (n.frozenRows = 0),
          (n.frozenCols = 0)),
      n
    ),
  );
}
function wQ(e, r, t) {
  const n = _t(t);
  return Gy(e, r, n.row, n.col);
}
function xQ(e, r, t, n, o) {
  Z3(r);
  const a = MN(e, r);
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
  const s = E1(e, r).flatMap((f) => {
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
    d = G3(r),
    u = r === "row" ? "frozenRows" : "frozenCols",
    l = e[u] ?? 0,
    c = o ? l - Math.max(0, Math.min(n, l - t)) : l + (t < l ? n : 0),
    h = {
      ...e,
      [d]: Ad(s, a),
      [u]: c,
      ...(e.outlineGroups ? { outlineGroups: TQ(e, r, t, n, o) } : {}),
    };
  return (
    K3(h, h),
    {
      [d]: h[d],
      [u]: c,
      ...(e.outlineGroups ? { outlineGroups: h.outlineGroups } : {}),
    }
  );
}
function hx(e, r, t, n = {}) {
  if (n.readOnly !== void 0 && typeof n.readOnly != "boolean")
    throw Error("Read-only state must be boolean.");
  if (n.readOnly) throw Error("This workbook is read-only.");
  const o = e.sheets.find((s) => s.id === r);
  if (!o) throw Error("Worksheet not found.");
  jn(o);
  const a = structuredClone(o);
  return (
    t(a),
    TB(a, a),
    K3(a, a),
    JSON.stringify(a) === JSON.stringify(o)
      ? e
      : { ...e, sheets: e.sheets.map((s) => (s.id === r ? a : s)) }
  );
}
function EQ(e, r, t, n, o, a) {
  return (
    F1(t),
    hx(
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
function CQ(e, r, t, n) {
  return hx(
    e,
    r,
    (o) => {
      var a;
      if (!((a = o.outlineGroups) != null && a.some((s) => s.id === t)))
        throw Error("Outline group not found.");
      o.outlineGroups = o.outlineGroups.filter((s) => s.id !== t);
    },
    n,
  );
}
function AQ(e, r, t, n, o) {
  if (typeof n != "boolean") throw Error("Collapsed state must be boolean.");
  return hx(
    e,
    r,
    (a) => {
      var d;
      const s =
        (d = a.outlineGroups) == null ? void 0 : d.find((u) => u.id === t);
      if (!s) throw Error("Outline group not found.");
      s.collapsed = n;
    },
    o,
  );
}
function kQ(e, r, t, n, o) {
  return (
    F1(t),
    hx(
      e,
      r,
      (a) => {
        var l;
        const s = NB(a, t),
          d = Math.max(0, ...s.map((c) => c.depth)) + 1;
        if (!Number.isInteger(n) || n < 1 || n > d)
          throw Error("Choose an available outline level.");
        const u = new Map(s.map((c) => [c.id, c.depth]));
        a.outlineGroups =
          (l = a.outlineGroups) == null
            ? void 0
            : l.map((c) =>
                c.axis === t ? { ...c, collapsed: u.get(c.id) >= n } : c,
              );
      },
      o,
    )
  );
}
function SQ(e, r, t, n) {
  return (
    F1(t),
    hx(
      e,
      r,
      (o) => {
        var a;
        o.outlineGroups =
          (a = o.outlineGroups) == null
            ? void 0
            : a.filter((s) => s.axis !== t);
      },
      n,
    )
  );
}
function _B(e) {
  if (!Array.isArray(e) || e.length > a1.rulesPerSheet)
    throw Error("At most 100 validation rules are supported.");
  const r = new Set();
  return e.map((t) => {
    if (
      !FQ(t) ||
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
      t.list.length > a1.listItems ||
      typeof t.min != "number" ||
      !Number.isFinite(t.min) ||
      typeof t.max != "number" ||
      !Number.isFinite(t.max)
    )
      throw Error("Invalid data validation rule.");
    const n = yv(t.id, 100, "ID");
    if (!n.trim() || r.has(n))
      throw Error("Validation rule IDs must be nonempty and unique.");
    r.add(n);
    const o = ct(yv(t.range, 40, "range")),
      a = {
        id: n,
        range: `${ut(o.firstCol)}${o.firstRow + 1}:${ut(o.lastCol)}${o.lastRow + 1}`,
        kind: t.kind,
        operator: t.operator,
        min: t.min,
        max: t.max,
        list: t.list.map((s) => yv(s, 1e3, "list item")),
        allowBlank: t.allowBlank,
        message: yv(t.message, 500, "error message"),
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
      t.formula !== void 0 &&
      ((a.formula = yv(t.formula, 8192, "formula")),
      !a.formula.startsWith("=") || a.formula.length < 2)
    )
      throw Error("Start the custom validation formula with =.");
    if (a.kind === "custom" && !a.formula)
      throw Error(
        "A custom rule requires a formula relative to the range's top-left cell.",
      );
    if (t.errorStyle !== void 0) {
      if (!["stop", "warning", "information"].includes(String(t.errorStyle)))
        throw Error("Invalid validation error style.");
      a.errorStyle = t.errorStyle;
    }
    for (const s of ["showInputMessage", "showDropdown"])
      if (t[s] !== void 0) {
        if (typeof t[s] != "boolean")
          throw Error("Invalid validation display setting.");
        a[s] = t[s];
      }
    for (const [s, d] of [
      ["inputTitle", 100],
      ["inputMessage", 500],
      ["errorTitle", 100],
    ])
      t[s] !== void 0 && (a[s] = yv(t[s], d, s));
    return a;
  });
}
function DB(e, r, t) {
  var o;
  const n = _t(t);
  return (
    ((o = e.sheets.find((a) => a.id === r)) == null ? void 0 : o.validations) ??
    []
  ).filter((a) => {
    const s = ct(a.range);
    return (
      n.row >= s.firstRow &&
      n.row <= s.lastRow &&
      n.col >= s.firstCol &&
      n.col <= s.lastCol
    );
  });
}
function FN(e) {
  if (!Array.isArray(e) || e.length > 100)
    throw Error("At most 100 conditional formatting rules are supported.");
  const r = new Set(),
    t = new Set();
  return e.map((n) => {
    var d;
    if (!PQ(n) || !BB.includes(n.kind))
      throw Error("Invalid conditional rule.");
    const o = Ub(n.id, 100, "conditional rule ID");
    if (!o.trim() || r.has(o))
      throw Error("Conditional rule IDs must be nonempty and unique.");
    r.add(o);
    const a = ct(Ub(n.range, 40, "conditional range")),
      s = {
        id: o,
        range: `${ut(a.firstCol)}${a.firstRow + 1}:${ut(a.lastCol)}${a.lastRow + 1}`,
        kind: n.kind,
        value: Ub(n.value, 1e3, "conditional value"),
        second: Ub(n.second, 1e3, "conditional second value"),
        fill: m6(n.fill),
        color: m6(n.color),
      };
    if (n.priority !== void 0) {
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
    if (n.stopIfTrue !== void 0) {
      if (typeof n.stopIfTrue != "boolean")
        throw Error("Stop if true must be boolean.");
      if (n.stopIfTrue && ["colorScale", "dataBar", "iconSet"].includes(s.kind))
        throw Error(
          "Stop if true is not supported for color scales, data bars, or icon sets.",
        );
      s.stopIfTrue = n.stopIfTrue;
    }
    if (
      (n.formula !== void 0 &&
        (s.formula = Ub(n.formula, 4096, "conditional formula")),
      s.kind === "formula" &&
        (!((d = s.formula) != null && d.startsWith("=")) ||
          !s.formula.slice(1).trim()))
    )
      throw Error("A formula rule requires a formula beginning with =.");
    for (const u of ["minColor", "maxColor"])
      n[u] !== void 0 && (s[u] = m6(n[u]));
    if (n.thresholdType !== void 0) {
      if (n.thresholdType !== "percent" && n.thresholdType !== "number")
        throw Error("Invalid icon threshold type.");
      s.thresholdType = n.thresholdType;
    }
    if (n.iconThresholds !== void 0) {
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
function Fw(e, r, t) {
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
    c = J3(e.conditionalFormats ?? []).map((f) => {
      const p = ct(f.range),
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
        for (const { address: y, point: w } of l) {
          if (++a > 1e6) {
            b =
              "Conditional statistics exceed 1,000,000 inspected cells; narrow the rule ranges or remove rules.";
            break;
          }
          if (!n5(p, w)) continue;
          const E = r(e.id, y);
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
      for (const {
        rule: g,
        bounds: b,
        counts: y,
        min: w,
        max: E,
        failure: A,
      } of c) {
        if (!n5(b, p)) continue;
        if (A) {
          m.errors.push(d(g.id, f, A));
          continue;
        }
        let k = false,
          F = {};
        if (g.kind === "formula") {
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
            const O = xd(g.formula, p.row - b.firstRow, p.col - b.firstCol);
            if (((I = r.evaluateFormula(e.id, f, O)), UQ(I)))
              throw Error(`Formula returned ${I}.`);
            if (typeof I != "number" && typeof I != "boolean")
              throw Error("Formula must return a scalar boolean or number.");
            k = !!I;
          } catch (O) {
            m.errors.push(
              d(
                g.id,
                f,
                O instanceof Error ? O.message : "Formula evaluation failed.",
              ),
            );
            continue;
          }
        } else if (g.kind === "blank" || g.kind === "notBlank")
          k = (v === null || v === "") == (g.kind === "blank");
        else if (g.kind === "error" || g.kind === "notError")
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
            k = I === (g.kind === "error");
          } catch (I) {
            m.errors.push(
              d(
                g.id,
                f,
                I instanceof Error
                  ? I.message
                  : "Error rule evaluation failed.",
              ),
            );
            continue;
          }
        else if (g.kind === "duplicate" || g.kind === "unique")
          k =
            v !== null &&
            v !== "" &&
            (y.get(String(v).toLowerCase()) ?? 0) > 1 ==
              (g.kind === "duplicate");
        else if (g.kind === "contains")
          k = String(v ?? "")
            .toLowerCase()
            .includes(g.value.toLowerCase());
        else if (g.kind === "equal")
          k = String(v ?? "").toLowerCase() === g.value.toLowerCase();
        else if (typeof v == "number") {
          if (g.kind === "greater") k = v > Number(g.value);
          else if (g.kind === "less") k = v < Number(g.value);
          else if (g.kind === "between")
            k = v >= Number(g.value) && v <= Number(g.second);
          else if (Number.isFinite(w) && Number.isFinite(E)) {
            const I =
              E === w ? 0.5 : Math.max(0, Math.min(1, (v - w) / (E - w)));
            if (g.kind === "colorScale") {
              const O = g.minColor ?? "#f8e99b",
                j = g.maxColor ?? "#63be7b";
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
            } else if (g.kind === "dataBar" && !m.dataBar) {
              const O = Math.min(0, w),
                j = Math.max(0, E),
                z = j - O || 1,
                V = (-O / z) * 100,
                D = ((v - O) / z) * 100;
              ((m.dataBar = {
                start: Math.min(V, D),
                width: Math.abs(D - V),
                axis: V,
                color: v < 0 ? "#d13438" : g.fill,
                negative: v < 0,
              }),
                (k = true));
            } else if (g.kind === "iconSet" && !m.icon) {
              const O = g.iconThresholds ?? [33, 67],
                j = (g.thresholdType ?? "percent") === "percent" ? I * 100 : v;
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
          (["colorScale", "dataBar", "iconSet"].includes(g.kind) ||
            (F = { fill: g.fill, color: g.color }),
          (m.style = { ...F, ...m.style }),
          g.stopIfTrue)
        )
          break;
      }
      return (o.size < 1e4 && o.set(f, structuredClone(m)), m);
    };
  return ((h.errors = () => [...n.values()].map((f) => ({ ...f }))), h);
}
const zQ = (e, r = 200) =>
    typeof e == "string" && e.length <= r
      ? e
      : al("Invalid worksheet tool text."),
  CE = (e) => {
    const r = ct(e);
    return `${ut(r.firstCol)}${r.firstRow + 1}:${ut(r.lastCol)}${r.lastRow + 1}`;
  },
  AE = (e, r) => {
    const t = ct(e),
      n = ct(r);
    return (
      t.firstRow <= n.lastRow &&
      t.lastRow >= n.firstRow &&
      t.firstCol <= n.lastCol &&
      t.lastCol >= n.firstCol
    );
  },
  lD = (e, r) => {
    const t = ct(e),
      n = _t(r);
    return (
      n.row >= t.firstRow &&
      n.row <= t.lastRow &&
      n.col >= t.firstCol &&
      n.col <= t.lastCol
    );
  };
function VQ(e, r) {
  if (sD(e)) {
    if (e.merges !== void 0) {
      (!Array.isArray(e.merges) || e.merges.length > 1e3) &&
        al("At most1,000 merged ranges are supported.");
      const t = [];
      for (const n of e.merges) {
        const o = CE(zQ(n, 40)),
          a = ct(o);
        ((a.lastRow - a.firstRow + 1) * (a.lastCol - a.firstCol + 1) > 500 &&
          al("A merged area supports at most 500 cells."),
          t.some((s) => AE(s, o)) && al("Merged ranges cannot overlap."),
          t.push(o));
      }
      r.merges = t;
    }
    if (
      (e.validations !== void 0 && (r.validations = _B(e.validations)),
      e.conditionalFormats !== void 0 &&
        (r.conditionalFormats = FN(e.conditionalFormats)),
      e.protection !== void 0)
    ) {
      const t = e.protection;
      ((!sD(t) ||
        typeof t.enabled != "boolean" ||
        (t.hash !== void 0 &&
          (!/^[a-f\d]{64}$/i.test(String(t.hash)) ||
            typeof t.salt != "string" ||
            t.salt.length > 100))) &&
        al("Invalid sheet protection."),
        (r.protection = {
          enabled: t.enabled,
          ...(t.hash ? { hash: String(t.hash), salt: String(t.salt) } : {}),
        }));
    }
  }
}
function kE(e, r, t, n = false) {
  var l, c;
  const o = e.sheets.find((h) => h.id === r);
  jn(o);
  const a = CE(t),
    s = ct(a);
  (s.lastRow - s.firstRow + 1) * (s.lastCol - s.firstCol + 1) > 500 &&
    al("Merge at most 500 cells.");
  const d = (o.merges ?? []).filter((h) => AE(h, a));
  !n && d.length && al("Unmerge overlapping areas first.");
  const u = { ...o.cells };
  if (!n) {
    const h = `${ut(s.firstCol)}${s.firstRow + 1}`;
    u[h] = {
      ...u[h],
      raw: ((l = u[h]) == null ? void 0 : l.raw) ?? null,
      style: { ...((c = u[h]) == null ? void 0 : c.style), align: "center" },
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
  for (const { refersTo: t } of e) {
    const n = Vg(
      t.replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')+'!/g, ""),
    );
    for (const o of n.matchAll(
      /(?:^|[^A-Za-z0-9_.])((?:\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?)(?![A-Za-z0-9_.])/g,
    )) {
      const a = ct(o[1]);
      if (
        ((r += (a.lastRow - a.firstRow + 1) * (a.lastCol - a.firstCol + 1)),
        r > Mt.maxCells)
      )
        throw Error(
          "Named ranges exceed the 100,000-cell XLSX import/export processing budget. Use native JSON for larger definitions.",
        );
    }
  }
}
function aP(e) {
  const r = iP(e),
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
      return Pi(
        "Malformed definedName content; only text definitions are supported.",
      );
    const u = b5(d[2]),
      l = u.name ?? "",
      c = u.localSheetId,
      h = d[1].slice(0, -11);
    if (
      { ...r.namespaces, ...u }[h ? `xmlns:${h.slice(0, -1)}` : "xmlns"] !== Uw
    )
      return Pi("Unsupported or undeclared definedName namespace.");
    for (const v of Object.keys(u))
      if (v.startsWith("xmlns") && u[v] !== Uw)
        return Pi("Unsupported definedName namespace.");
    if (c !== void 0 && (!/^\d+$/.test(c) || !Number.isSafeInteger(Number(c))))
      return Pi(`Name "${l}" has an invalid localSheetId.`);
    const p = `=${nP(
      (d[3] ?? "").replace(
        /\r\n?/g,
        `
`,
      ),
    ).replace(/^=/, "")}`;
    if ((pC(p), c !== void 0 && Number(c) >= s))
      return Pi(`Name "${l}" has an out-of-range localSheetId.`);
    const m = `${c === void 0 ? "workbook" : Number(c)}:${l.toLowerCase()}`;
    if (a.has(m)) return Pi(`Duplicate name "${l}" in the same scope.`);
    if ((a.add(m), une.has(l.toLowerCase())))
      t.builtinXml += `<definedName${Object.entries(u)
        .filter(([v]) => !v.startsWith("xmlns"))
        .map(([v, g]) => ` ${v}="${Yy(g, true)}"`)
        .join("")}>${Yy(p.slice(1))}</definedName>`;
    else if (l.toLowerCase().startsWith("_xlnm.")) {
      if (++o > 1e3)
        return Pi("The workbook exceeds the 1,000-defined-name limit.");
      t.warnings.push(
        `Built-in defined name "${l}" is not supported and was omitted. Review its source feature before relying on imported formulas.`,
      );
    } else {
      if (++o > 1e3)
        return Pi("The workbook exceeds the 1,000-defined-name limit.");
      for (const g of ["function", "vbProcedure", "xlm"])
        if (u[g] !== void 0 && !["0", "false"].includes(u[g]))
          return Pi(`Name "${l}" has an executable ${g} attribute.`);
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
          ...(c !== void 0 ? { localSheetId: Number(c) } : {}),
          ...(u.comment !== void 0 ? { comment: u.comment } : {}),
        }));
    }
    n = n.slice(d[0].length);
  }
  oP(t.names);
  try {
    eg(
      t.names.map((d) => ({
        name: d.name,
        refersTo: d.refersTo,
        ...(d.localSheetId !== void 0
          ? { scopeSheetId: String(d.localSheetId) }
          : {}),
        ...(d.comment !== void 0 ? { comment: d.comment } : {}),
      })),
      {
        sheets: [
          ...new Set(
            t.names.flatMap((d) =>
              d.localSheetId === void 0 ? [] : [d.localSheetId],
            ),
          ),
        ].map((d) => ({ id: String(d) })),
      },
    );
  } catch (d) {
    return Pi(d.message);
  }
  return t;
}

export {
  AE,
  AQ,
  CE,
  CQ,
  DB,
  Die,
  EB,
  EQ,
  FN,
  Fg,
  Fw,
  G5,
  GO,
  Gy,
  JS,
  K3,
  LU,
  Np,
  ON,
  Qne,
  SQ,
  TC,
  VQ,
  W3,
  X3,
  XR,
  _B,
  _t,
  aP,
  ct,
  d6,
  dQ,
  eie,
  ff,
  iae,
  kE,
  kQ,
  lD,
  oP,
  pB,
  pQ,
  q3,
  sQ,
  tae,
  vae,
  wQ,
  xQ,
  y1,
  yQ,
};
