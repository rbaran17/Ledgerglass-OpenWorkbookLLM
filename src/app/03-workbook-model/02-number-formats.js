import qre from "fast-formula-parser";
import { ax } from "../01-core-schemas-and-hashing.js";
import { Ay, Bw, GQ, PB, Si, UB, Uv, XQ, ZQ, zB } from "./01-cell-styles-and-themes.js";

function l0(e) {
  (!ZQ(e) ||
    Object.keys(e).some(
      (a) =>
        ![
          "category",
          "locale",
          "currency",
          "decimals",
          "useGrouping",
          "scale",
          "negative",
          "pattern",
        ].includes(a),
    )) &&
    Si("Number format must be a supported plain settings object.");
  const r = e;
  Uv.includes(r.category) || Si("Choose a supported number format category.");
  const t = r.category,
    n = r.locale ?? "en-US";
  PB.includes(n) || Si("Choose one of the supported display locales.");
  const o = { category: t, locale: n };
  if (XQ.has(t)) {
    if (t === "currency" || t === "accounting") {
      const u = r.currency ?? "USD";
      (UB.includes(u) || Si("Choose a supported currency."), (o.currency = u));
    } else
      r.currency !== void 0 &&
        Si("Currency applies only to Currency or Accounting formats.");
    const s = r.decimals ?? (o.currency === "JPY" ? 0 : 2);
    ((typeof s != "number" || !Number.isInteger(s) || s < 0 || s > 12) &&
      Si("Choose 0–12 decimal places."),
      (o.decimals = s),
      r.useGrouping !== void 0 &&
        typeof r.useGrouping != "boolean" &&
        Si("The grouping option must be true or false."),
      (o.useGrouping = r.useGrouping ?? t !== "percent"));
    const d = r.scale ?? 1;
    ((![1, 1e3, 1e6].includes(d) || (t === "percent" && d !== 1)) &&
      Si("Choose units, thousands or millions; percentages use units."),
      (o.scale = d),
      r.negative !== void 0 &&
        (typeof r.negative != "string" ||
          !["minus", "parentheses"].includes(r.negative)) &&
        Si("Choose minus signs or parentheses for negatives."),
      t === "accounting" &&
        r.negative !== void 0 &&
        r.negative !== "parentheses" &&
        Si("Accounting negatives use parentheses."),
      (o.negative =
        t === "accounting" ? "parentheses" : (r.negative ?? "minus")),
      r.pattern !== void 0 &&
        Si("Use the Custom category to enter a format pattern."));
  } else if (
    ([r.currency, r.decimals, r.useGrouping, r.scale, r.negative].some(
      (a) => a !== void 0,
    ) && Si("These numeric settings do not apply to the chosen category."),
    GQ.has(t) || t === "custom")
  ) {
    const a = t === "custom" ? "0.00" : Ay[t][0],
      s = r.pattern ?? a;
    ((typeof s != "string" ||
      !s ||
      s.length > 200 ||
      /[\u0000-\u001f\u007f]/.test(s)) &&
      Si(
        "Enter a format pattern of 1–200 characters without control characters.",
      ),
      (o.pattern = s));
    const d = Bw(o.pattern);
    if (t !== "custom") {
      d.kind !== "temporal" &&
        Si("Choose a date, time or elapsed-time pattern for this category.");
      const u = d.category;
      u !== t &&
        Si(
          `This pattern is a ${u} format. Choose the matching category or Custom.`,
        );
    }
  } else
    r.pattern !== void 0 && Si("General does not accept a custom pattern.");
  return o;
}
function QQ(e, r) {
  return e.replace(
    /"([^"]*)"|\\(.)|@/g,
    (t, n, o) => n ?? o ?? (t === "@" ? r : t),
  );
}
const VB = (e) => {
  var r;
  return (
    ((r = new Intl.NumberFormat(e)
      .formatToParts(1.1)
      .find((t) => t.type === "decimal")) == null
      ? void 0
      : r.value) ?? "."
  );
};
function eee(e, r, t) {
  const n = e < 0 && r.sections.length > 1,
    { parsed: o, divisor: a } =
      r.sections[e === 0 && r.sections.length > 2 ? 2 : n ? 1 : 0];
  if (!o.hasNumber) return o.prefix + o.suffix;
  const s = (Math.abs(e) / a) * (o.percent ? 100 : 1);
  if (!Number.isFinite(s)) return "#NUM!";
  let d = new Intl.NumberFormat(t, {
    minimumFractionDigits: o.minDecimals,
    maximumFractionDigits: o.maxDecimals,
    minimumIntegerDigits: Math.max(1, o.minIntegerDigits),
    useGrouping: o.grouping,
  }).format(s);
  return (
    o.minIntegerDigits ||
      (d === "0" && !o.minDecimals
        ? (d = "")
        : d.startsWith(`0${VB(t)}`) && (d = d.slice(1))),
    (e < 0 && !n ? "-" : "") + o.prefix + d + o.suffix
  );
}
function tee(e, r, t, n) {
  const o = r.category === "duration",
    a = 2958465 - (n === "1904" && !o ? 1462 : 0);
  if (Math.abs(e) >= a + 1 || (!o && e < 0)) return "#NUM!";
  const s = r.category !== "date",
    d = 10 ** r.fraction,
    u = r.tokens.some(
      (j) =>
        j.kind === "second" || (j.kind === "elapsed" && j.text.startsWith("s")),
    ),
    l = s
      ? u
        ? Math.round(Math.abs(e) * 86400 * d) / d
        : Math.floor(Math.abs(e) * 86400)
      : Math.floor(e) * 86400;
  let c = Math.floor(l / 86400),
    h = l - c * 86400;
  if (c > a) return "#NUM!";
  const f = Math.round(h * 1e3),
    p = Math.floor(f / 36e5),
    m = Math.floor(f / 6e4) % 60,
    v = Math.floor(f / 1e3) % 60,
    g = Math.floor((f % 1e3) / 10 ** (3 - r.fraction)),
    b = n === "1900" && c === 60,
    y =
      n === "1904"
        ? Date.UTC(1904, 0, 1) + c * 864e5
        : Date.UTC(1899, 11, 31) + (c > 60 ? c - 1 : c) * 864e5,
    w = new Date(y),
    E = b ? 1900 : w.getUTCFullYear(),
    A = b ? 1 : w.getUTCMonth(),
    k = b ? 29 : w.getUTCDate(),
    F = r.tokens.some((j) => j.kind === "ampm"),
    I = (j, z) => String(j).padStart(z, "0"),
    O = r.tokens
      .map((j) => {
        switch (j.kind) {
          case "literal":
            return j.text;
          case "year":
            return j.text.length === 2 ? I(E % 100, 2) : I(E, 4);
          case "month":
            return j.text.length < 3
              ? I(A + 1, j.text.length)
              : new Intl.DateTimeFormat(t, {
                  month: j.text.length === 3 ? "short" : "long",
                  timeZone: "UTC",
                }).format(new Date(Date.UTC(2e3, A, 1)));
          case "day":
            return I(k, j.text.length);
          case "hour":
            return I(F ? p % 12 || 12 : p, j.text.length);
          case "minute":
            return I(m, j.text.length);
          case "second":
            return I(v, j.text.length);
          case "fraction":
            return VB(t) + I(g, r.fraction);
          case "ampm": {
            const z = p < 12 ? "AM" : "PM",
              V = j.text.length === 3 ? z[0] : z;
            return j.text === j.text.toLowerCase() ? V.toLowerCase() : V;
          }
          case "elapsed":
            return I(
              Math.floor(
                l / (j.text[0] === "h" ? 3600 : j.text[0] === "m" ? 60 : 1),
              ),
              j.text.length,
            );
        }
      })
      .join("");
  return (o && e < 0 ? "-" : "") + O;
}
function vg(e, r, t = "1900") {
  var l;
  const n = l0(r);
  if (e === null) return { text: "" };
  if (typeof e == "boolean") return { text: e ? "TRUE" : "FALSE" };
  if (typeof e == "string") {
    if (n.category === "custom" && !/^#[A-Z0-9/?!]+$/.test(e)) {
      const c = Bw(n.pattern);
      if (c.kind === "numeric" && c.textSection !== void 0)
        return { text: QQ(c.textSection, e) };
    }
    return { text: e };
  }
  if (!Number.isFinite(e)) return { text: "#NUM!" };
  if (n.category === "general")
    return {
      text: new Intl.NumberFormat(n.locale, {
        maximumSignificantDigits: 15,
        useGrouping: false,
      }).format(e === 0 ? 0 : e),
    };
  if (n.pattern) {
    const c = Bw(n.pattern);
    return {
      text:
        c.kind === "temporal" ? tee(e, c, n.locale, t) : eee(e, c, n.locale),
    };
  }
  const o = e / (n.scale ?? 1),
    a = n.scale === 1e3 ? " K" : n.scale === 1e6 ? " M" : "";
  if (n.category === "percent" && !Number.isFinite(o * 100))
    return { text: "#NUM!" };
  const s = {
    minimumFractionDigits: n.decimals,
    maximumFractionDigits: n.decimals,
    useGrouping: n.useGrouping,
  };
  if (n.category === "accounting") {
    const h = new Intl.NumberFormat(n.locale, {
        ...s,
        style: "currency",
        currency: n.currency,
      }).formatToParts(Math.abs(o)),
      f =
        ((l = h.find((g) => g.type === "currency")) == null
          ? void 0
          : l.value) ?? n.currency,
      p =
        h.findIndex((g) => g.type === "currency") <
        h.findIndex((g) => g.type === "integer")
          ? "start"
          : "end",
      m = new Intl.NumberFormat(n.locale, s).format(Math.abs(o)),
      v = (o === 0 ? "-" : o < 0 ? `(${m})` : m) + a;
    return {
      text: p === "start" ? `${f} ${v}` : `${v} ${f}`,
      accounting: { symbol: f, amount: v, symbolPosition: p },
    };
  }
  const d = o < 0 && n.negative === "parentheses",
    u = new Intl.NumberFormat(n.locale, {
      ...s,
      ...(n.category === "currency"
        ? { style: "currency", currency: n.currency }
        : n.category === "percent"
          ? { style: "percent" }
          : {}),
    }).format(d ? Math.abs(o) : o);
  return { text: (d ? `(${u})` : u) + a };
}
function SE(e) {
  var r;
  if (e) {
    if (e.numberFormat)
      try {
        l0(e.numberFormat);
        return;
      } catch (t) {
        return t instanceof Error
          ? t.message
          : "Unsupported number format settings.";
      }
    if (
      !(e.numFmt === "@" && typeof e.raw == "string") &&
      !(
        e.format === "date" &&
        ((r = e.numFmt) == null ? void 0 : r.toLowerCase()) === "yyyy-mm-dd"
      )
    ) {
      if (
        e.numFmt &&
        /\[(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\d+)\]/i.test(
          e.numFmt,
        )
      )
        return "The legacy format's color directive is retained but does not color the text. Use conditional formatting to apply color explicitly.";
      if (
        e.numFmt &&
        e.numFmt.toLowerCase() !== "general" &&
        zB(1.25, e.numFmt) === void 0
      )
        return `The imported pattern “${e.numFmt}” is outside the legacy display subset. The basic ${e.format ?? "general"} display is used; the original pattern and raw values are retained. Choose a supported format to change its display explicitly.`;
    }
  }
}
const ree = {
  "en-US": "409",
  "en-GB": "809",
  "de-DE": "407",
  "fr-FR": "40C",
  "ja-JP": "411",
  "en-IN": "4009",
};
function Q3(e) {
  const r = l0(e),
    t = r.locale === "en-US" ? "" : `[$-${ree[r.locale]}]`;
  if (r.category === "general") return "General";
  if (r.pattern) return t + r.pattern;
  const n = r.decimals ? "." + "0".repeat(r.decimals) : "",
    o = r.scale === 1e3 ? "," : r.scale === 1e6 ? ",," : "",
    a = r.scale === 1e3 ? '" K"' : r.scale === 1e6 ? '" M"' : "",
    s = (r.useGrouping ? "#,##0" : "0") + n + o;
  if (r.category === "percent") return t + s + "%";
  if (r.category === "currency" || r.category === "accounting") {
    const d = vg(1, {
        ...r,
        category: "accounting",
        negative: "parentheses",
      }).accounting,
      u = `"${d.symbol.replace(/"/g, '""')}"`;
    if (r.category === "accounting") {
      const c = "?".repeat(r.decimals);
      return d.symbolPosition === "start"
        ? `${t}_(${u}* ${s}${a}_);_(${u}* (${s}${a});_(${u}* "-"${c}_);_(@_)`
        : `${t}_(${s}${a}* ${u}_);_((${s}${a})* ${u});_("-"${c}* ${u}_);_(@_)`;
    }
    const l = d.symbolPosition === "start" ? u + s + a : s + a + '" "' + u;
    return t + l + (r.negative === "parentheses" ? `;(${l})` : "");
  }
  return t + s + a + (r.negative === "parentheses" ? `;(${s}${a})` : "");
}
const JN = ax(qre);
function zL(e) {
  const [r, t = "0"] = Math.abs(e).toString().split("e"),
    [n, o = ""] = r.split(".");
  return {
    coefficient: BigInt(n + o) * (e < 0 ? -1n : 1n),
    exponent: Number(t) - o.length,
  };
}
function VL(e) {
  const r = new Map();
  for (const a of e) {
    if (!Number.isFinite(a)) return;
    const { coefficient: s, exponent: d } = zL(a);
    r.set(d, (r.get(d) ?? 0n) + s);
  }
  const t = [...r].filter(([, a]) => a !== 0n);
  if (!t.length) return { coefficient: 0n, exponent: 0 };
  const n = t.reduce((a, [s]) => Math.min(a, s), 1 / 0);
  return {
    coefficient: t.reduce((a, [s, d]) => a + d * 10n ** BigInt(s - n), 0n),
    exponent: n,
  };
}
function Rc(e) {
  const r = VL(e);
  return r ? +`${r.coefficient}e${r.exponent}` : NaN;
}
function Us(e) {
  if (!e.length) return NaN;
  const r = VL(e);
  if (!r) return NaN;
  const t = r.coefficient < 0n ? "-" : "",
    n = r.coefficient < 0n ? -r.coefficient : r.coefficient,
    o = BigInt(e.length),
    a = n / o,
    s = (((n % o) * 10n ** 40n) / o).toString().padStart(40, "0");
  return +`${t}${a}.${s}e${r.exponent}`;
}
function WL(e) {
  if (e.some((n) => !Number.isFinite(n))) return NaN;
  if (e.some((n) => n === 0)) return 0;
  let r = 1n,
    t = 0;
  for (const n of e) {
    const o = zL(n);
    if (((r *= o.coefficient), (t += o.exponent), e.length > 255)) {
      const a = r < 0n,
        s = (a ? -r : r).toString();
      s.length > 80 &&
        ((t += s.length - 80), (r = BigInt(s.slice(0, 80)) * (a ? -1n : 1n)));
    }
  }
  return +`${r}e${t}`;
}
const HL = Symbol("not-arithmetic"),
  Xre = new Set([
    "ROUND",
    "ROUNDUP",
    "ROUNDDOWN",
    "TRUNC",
    "MROUND",
    "CEILING",
    "FLOOR",
    "SUMPRODUCT",
  ]);
function h5(e) {
  const [r, t = "0"] = Math.abs(e).toString().split("e"),
    [n, o = ""] = r.split(".");
  return { coefficient: BigInt(n + o), exponent: Number(t) - o.length };
}
function qL(e, r, t) {
  const n = e / r,
    o = e % r;
  return (
    n +
    (o !== 0n && (t === "away" || (t === "nearest" && o * 2n >= r)) ? 1n : 0n)
  );
}
function Gre(e, r, t) {
  if (((r = Math.trunc(r)), e === 0 || r >= 324)) return e;
  if (r <= -309) return t === "away" ? 1 / 0 : 0;
  const n = h5(e),
    o = n.exponent + r;
  if (o >= 0) return e;
  const a = qL(n.coefficient, 10n ** BigInt(-o), t);
  return Math.sign(e) * +`${a}e${-r}`;
}
function Zre(e, r, t) {
  const n = h5(e),
    o = h5(r),
    a = Math.min(n.exponent, o.exponent),
    s = n.coefficient * 10n ** BigInt(n.exponent - a),
    d = o.coefficient * 10n ** BigInt(o.exponent - a),
    u = qL(s, d, t);
  return Math.sign(e) * +`${u * o.coefficient}e${o.exponent}`;
}
function Kre(e, r, t) {
  var f;
  if (!Xre.has(e)) return HL;
  const { retrieve: n, issue: o, isIssue: a } = t,
    s = r.map(n),
    d = (p) => (Number.isFinite(p) ? (p === 0 ? 0 : p) : o("#NUM!"));
  if (e === "SUMPRODUCT") {
    if (!r.length || r.length > 255) return o("#VALUE!");
    const p = s.map((b) =>
        Array.isArray(b) ? (b.every(Array.isArray) ? b : [b]) : [[b]],
      ),
      m = p[0].length,
      v = ((f = p[0][0]) == null ? void 0 : f.length) ?? 0;
    if (
      !m ||
      !v ||
      p.some((b) => b.length !== m || b.some((y) => y.length !== v))
    )
      return o("#VALUE!");
    if (m * v * p.length > 1e5) return o("#LIMIT!");
    const g = [];
    for (let b = 0; b < m; b++)
      for (let y = 0; y < v; y++) {
        const w = [];
        for (const A of p) {
          const k = A[b][y];
          if (a(k)) return k;
          w.push(typeof k == "number" ? k : 0);
        }
        const E = WL(w);
        if (!Number.isFinite(E)) return o("#NUM!");
        g.push(E);
      }
    return d(Rc(g));
  }
  if (r.length < (e === "TRUNC" ? 1 : 2) || r.length > 2) return o("#VALUE!");
  const u = [];
  for (const p of s) {
    if (a(p)) return p;
    if (
      Array.isArray(p) ||
      (typeof p == "string" &&
        !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(p.trim()))
    )
      return o("#VALUE!");
    const m = p == null ? 0 : Number(p);
    if (!Number.isFinite(m)) return o("#NUM!");
    u.push(m);
  }
  const [l, c = 0] = u;
  if (["ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC"].includes(e))
    return d(
      Gre(
        l,
        c,
        e === "ROUND" ? "nearest" : e === "ROUNDUP" ? "away" : "toward",
      ),
    );
  if (l === 0 || c === 0) return 0;
  if (
    (e === "MROUND" && Math.sign(l) !== Math.sign(c)) ||
    (e === "FLOOR" && l > 0 && c < 0)
  )
    return o("#NUM!");
  const h =
    e === "MROUND"
      ? "nearest"
      : e === "CEILING"
        ? l < 0 && c > 0
          ? "toward"
          : "away"
        : l < 0 && c > 0
          ? "away"
          : "toward";
  return d(Zre(l, c, h));
}
const XL = ["LINEST", "LOGEST", "TREND", "GROWTH"],
  GL = Symbol("not-regression"),
  Wb = {
    values: 4e4,
    observations: 1e4,
    predictors: 32,
    work: 4e6,
    output: 1e4,
  },
  ad = (e, r) => {
    let t = 0,
      n = 0;
    for (let o = 0; o < e.length; o++) {
      const a = e[o] * r[o],
        s = t + a;
      ((n += Math.abs(t) >= Math.abs(a) ? t - s + a : a - s + t), (t = s));
    }
    return t + n;
  };
function jE(e, r, t) {
  if (!XL.includes(e)) return GL;
  const n = (se = "#VALUE!") => {
      throw t.issue(se);
    },
    o = (se) => (Number.isFinite(se) ? se : n("#NUM!"));
  (!r.length || r.length > 4 || r[0] == null) && n();
  let a = 0;
  const s = (se) => {
      var Ae;
      if (t.isIssue(se)) throw se;
      const we = Array.isArray(se)
          ? se.every(Array.isArray)
            ? se
            : [se]
          : [[se]],
        Ee = ((Ae = we[0]) == null ? void 0 : Ae.length) ?? 0;
      return (
        (!we.length || !Ee || we.some((ye) => ye.length !== Ee)) && n(),
        (a += we.length * Ee),
        a > Wb.values && n("#LIMIT!"),
        we.map((ye) =>
          ye.map((Te) => {
            if (t.isIssue(Te)) throw Te;
            return (typeof Te != "number" && n(), o(Te));
          }),
        )
      );
    },
    d = (se, we) => {
      if (se == null) return we;
      const Ee = t.retrieve(se);
      if (t.isIssue(Ee)) throw Ee;
      return Ee == null
        ? false
        : typeof Ee == "boolean"
          ? Ee
          : typeof Ee == "number"
            ? o(Ee) !== 0
            : typeof Ee == "string" && /^(true|false)$/i.test(Ee)
              ? /^true$/i.test(Ee)
              : n();
    },
    u = e === "TREND" || e === "GROWTH",
    l = e === "LOGEST" || e === "GROWTH",
    c = d(r[u ? 3 : 2], true),
    h = !u && d(r[3], false),
    f = s(t.retrieve(r[0])),
    p = f.length * f[0].length;
  p > Wb.observations && n("#LIMIT!");
  const m = f.flat();
  l && m.some((se) => se <= 0) && n("#NUM!");
  const v = l ? m.map(Math.log) : m,
    g =
      r[1] == null
        ? f.map((se, we) => se.map((Ee, Ae) => we * se.length + Ae + 1))
        : s(t.retrieve(r[1])),
    b = g.length === f.length && g[0].length === f[0].length,
    y = f.length === 1;
  let w;
  if (b) w = g.flat().map((se) => [se]);
  else if (f[0].length === 1 && g.length === p) w = g;
  else if (y && g[0].length === p)
    w = Array.from({ length: p }, (se, we) => g.map((Ee) => Ee[we]));
  else return n("#REF!");
  const E = w[0].length;
  (E > Wb.predictors || p * E * E > Wb.work) && n("#LIMIT!");
  const A = Array.from({ length: E }, (se, we) =>
      c ? Us(w.map((Ee) => Ee[we])) : 0,
    ),
    k = c ? Us(v) : 0,
    F = v.map((se) => o(se - k)),
    I = Math.max(...F.map(Math.abs)) || 1,
    O = F.map((se) => se / I),
    j = [],
    z = Array.from({ length: E }, (se, we) => {
      const Ee = w.map((ye) => o(ye[we] - A[we])),
        Ae = Math.max(...Ee.map(Math.abs)) || 1;
      return (j.push(Ae), Ee.map((ye) => ye / Ae));
    }),
    V = z.map((se) => [...se]),
    D = Array.from({ length: E }, (se, we) => we),
    B = Array.from({ length: E }, () => Array(E).fill(0)),
    _ = [],
    C = [],
    T = 1e-12 * Math.max(1, ...z.map((se) => Math.sqrt(ad(se, se))));
  let N = 0;
  for (let se = 0; se < E; se++) {
    let we = se;
    for (let Ae = se + 1; Ae < E; Ae++)
      ad(z[Ae], z[Ae]) > ad(z[we], z[we]) && (we = Ae);
    (([z[se], z[we]] = [z[we], z[se]]), ([D[se], D[we]] = [D[we], D[se]]));
    for (let Ae = 0; Ae < se; Ae++)
      [B[Ae][se], B[Ae][we]] = [B[Ae][we], B[Ae][se]];
    const Ee = Math.sqrt(ad(z[se], z[se]));
    if (Ee <= T) break;
    ((B[se][se] = Ee),
      _.push(z[se].map((Ae) => Ae / Ee)),
      C.push(ad(_[se], O)),
      N++);
    for (let Ae = se + 1; Ae < E; Ae++)
      for (let ye = 0; ye < 2; ye++) {
        const Te = ad(_[se], z[Ae]);
        ((B[se][Ae] += Te),
          (z[Ae] = z[Ae].map((Pe, qe) => Pe - Te * _[se][qe])));
      }
  }
  const R = (se) => {
      const we = Array(N).fill(0);
      for (let Ee = N - 1; Ee >= 0; Ee--) {
        let Ae = se[Ee];
        for (let ye = Ee + 1; ye < N; ye++) Ae -= B[Ee][ye] * we[ye];
        we[Ee] = o(Ae / B[Ee][Ee]);
      }
      return we;
    },
    L = Array(E).fill(0);
  R(C).forEach((se, we) => (L[D[we]] = se));
  const $ = L.map((se, we) => o(se * (I / j[we]))),
    U = c ? o(k - ad($, A)) : 0,
    G = (se) => {
      const we = o(Math.exp(se));
      return we > 0 ? we : n("#NUM!");
    };
  if (u) {
    const se = r[2] == null ? g : s(t.retrieve(r[2]));
    let we, Ee, Ae;
    if (E === 1)
      ((we = se.flat().map((Te) => [Te])),
        (Ee = se.length),
        (Ae = se[0].length));
    else if (!y && se[0].length === E) ((we = se), (Ee = we.length), (Ae = 1));
    else if (y && se.length === E)
      ((we = Array.from({ length: se[0].length }, (Te, Pe) =>
        se.map((qe) => qe[Pe]),
      )),
        (Ee = 1),
        (Ae = we.length));
    else return n("#REF!");
    Ee * Ae > Wb.output && n("#LIMIT!");
    const ye = we.map((Te) => {
      const Pe = o(
        k +
          ad(
            L,
            Te.map((qe, ht) => o(qe - A[ht]) / j[ht]),
          ) *
            I,
      );
      return l ? G(Pe) : Pe;
    });
    return Array.from({ length: Ee }, (Te, Pe) =>
      ye.slice(Pe * Ae, (Pe + 1) * Ae),
    );
  }
  const H = [...$]
    .reverse()
    .concat(U)
    .map((se) => (l ? G(se) : se));
  if (!h) return [H];
  const Q = O.map(
    (se, we) =>
      se -
      ad(
        L,
        V.map((Ee) => Ee[we]),
      ),
  );
  let X = ad(Q, Q);
  const J = ad(O, O);
  X <= 1e-26 * Math.max(J, Number.MIN_VALUE) && (X = 0);
  const q = p - N - Number(c),
    W = q > 0 ? Math.sqrt(X / q) : null,
    K = () => t.issue("#N/A"),
    te = () => t.issue("#DIV/0!"),
    ae = Array.from({ length: N }, (se, we) =>
      R(Array.from({ length: N }, (Ee, Ae) => (Ae === we ? 1 : 0))),
    ),
    oe = Array(E).fill(0);
  for (let se = 0; se < N; se++)
    oe[D[se]] =
      W === null
        ? te()
        : o(Math.hypot(...ae.map((we) => we[se])) * W * (I / j[D[se]]));
  const de = D.slice(0, N).map((se) => A[se] / j[se]),
    he = 1 / p + ae.reduce((se, we) => se + ad(de, we) ** 2, 0),
    Se = [...oe]
      .reverse()
      .concat(c ? (W === null ? te() : o(Math.sqrt(he) * W * I)) : K()),
    le = Math.max(0, J - X),
    ce = (se) => o(se * I * I),
    ge = (se, we) => [se, we, ...Array.from({ length: E - 1 }, K)];
  return [
    H,
    Se,
    ge(
      J ? Math.max(0, Math.min(1, le / J)) : te(),
      W === null ? te() : o(W * I),
    ),
    ge(q > 0 && N && X ? o(le / N / (X / q)) : te(), q),
    ge(ce(le), ce(X)),
  ];
}
const ZL = ["T.TEST", "Z.TEST"],
  KL = Symbol("not-inference");
function Yre(e, r, t) {
  if (!ZL.includes(e)) return KL;
  const n = (I = "#VALUE!") => {
      throw t.issue(I);
    },
    o = (I) => (Number.isFinite(I) ? I : n("#NUM!"));
  ((e === "T.TEST" && r.length !== 4) ||
    (e === "Z.TEST" && (r.length < 2 || r.length > 3))) &&
    n();
  const a = r.map(t.retrieve);
  let s = 0;
  const d = (I) => {
      if (Array.isArray(I)) return I.flatMap(d);
      if ((++s > 4e4 && n("#LIMIT!"), t.isIssue(I))) throw I;
      return (typeof I == "number" && o(I), [I]);
    },
    u = (I) => {
      if (t.isIssue(I)) throw I;
      if (Array.isArray(I) || I == null || (typeof I == "string" && !I.trim()))
        return n();
      const O = Number(I);
      return Number.isFinite(O) ? O : n();
    },
    l = (I) => I.filter((O) => typeof O == "number"),
    c = (I) => {
      if (I.length < 2) return n("#DIV/0!");
      const O = Us(I);
      return {
        mean: O,
        variance: o(Rc(I.map((j) => (j - O) ** 2)) / (I.length - 1)),
      };
    },
    h = (I) => {
      if (t.isIssue(I)) throw I;
      return typeof I != "number" ||
        !Number.isFinite(I) ||
        I < -1e-12 ||
        I > 1 + 1e-12
        ? n("#NUM!")
        : Math.max(0, Math.min(1, I));
    },
    f = d(a[0]);
  if (e === "Z.TEST") {
    const I = l(f),
      O = I.length;
    if (!O) return n("#N/A");
    const j = u(a[1]),
      z = Math.max(Math.abs(j), ...I.map(Math.abs)) || 1,
      V = I.map((C) => C / z),
      D = r[2] == null ? Math.sqrt(c(V).variance) : u(a[2]) / z;
    if (D < 0) return n("#NUM!");
    if (!D) return n("#DIV/0!");
    const B = o(((Us(V) - j / z) / D) * Math.sqrt(O)),
      _ = t.distribution("ERFC", [B / Math.SQRT2]);
    if (t.isIssue(_)) throw _;
    return typeof _ == "number" ? h(_ / 2) : n("#NUM!");
  }
  const p = d(a[1]),
    m = Math.trunc(u(a[2])),
    v = Math.trunc(u(a[3]));
  ((![1, 2].includes(m) || ![1, 2, 3].includes(v)) && n("#NUM!"),
    v === 1 && f.length !== p.length && n("#N/A"));
  let g, b;
  if (v === 1) {
    const I = f
      .map((O, j) => [O, p[j]])
      .filter((O) => O.every((j) => typeof j == "number"));
    ((g = I.map((O) => O[0])), (b = I.map((O) => O[1])));
  } else ((g = l(f)), (b = l(p)));
  (g.length < 2 || b.length < 2) && n("#DIV/0!");
  const y = Math.max(...g.map(Math.abs), ...b.map(Math.abs)) || 1;
  ((g = g.map((I) => I / y)), (b = b.map((I) => I / y)));
  let w, E, A;
  if (v === 1) {
    const I = c(g.map((O, j) => O - b[j]));
    ((w = I.mean), (E = I.variance / g.length), (A = g.length - 1));
  } else {
    const I = c(g),
      O = c(b);
    if (((w = I.mean - O.mean), v === 2))
      ((A = g.length + b.length - 2),
        (E =
          (((g.length - 1) * I.variance + (b.length - 1) * O.variance) / A) *
          (1 / g.length + 1 / b.length)));
    else {
      const j = I.variance / g.length,
        z = O.variance / b.length;
      ((E = j + z),
        (A = E ** 2 / (j ** 2 / (g.length - 1) + z ** 2 / (b.length - 1))));
    }
  }
  E || n("#DIV/0!");
  const k = o(Math.abs(w) / Math.sqrt(E));
  o(A);
  const F = t.distribution("BETA.DIST", [A / (A + k * k), A / 2, 0.5, true]);
  return (h(F) * m) / 2;
}
const YL = ["CONVERT", "SLOPE", "RSQ", "STEYX", "PEARSON"],
  JL = Symbol("not-specialist"),
  Ky = [],
  Fr = (e, r, t, n = false, o = 0, a = 1) => {
    for (const s of e.split("|"))
      Ky.push({
        symbol: s,
        dimension: r,
        factor: t,
        metric: n,
        offset: o,
        power: a,
      });
  };
Fr("g", "mass", 0.001, true);
Fr("lbm", "mass", 0.45359237);
Fr("ozm", "mass", 0.45359237 / 16);
Fr("grain", "mass", 0.45359237 / 7e3);
Fr("stone", "mass", 0.45359237 * 14);
Fr("ton", "mass", 0.45359237 * 2e3);
Fr("uk_ton|LTON|brton", "mass", 0.45359237 * 2240);
Fr("cwt|shweight", "mass", 0.45359237 * 100);
Fr("uk_cwt|lcwt|hweight", "mass", 0.45359237 * 112);
Fr("m", "length", 1, true);
Fr("in", "length", 0.0254);
Fr("ft", "length", 0.3048);
Fr("yd", "length", 0.9144);
Fr("mi", "length", 1609.344);
Fr("Nmi", "length", 1852);
Fr("ang", "length", 1e-10, true);
Fr("Picapt|Pica", "length", 0.0254 / 72);
Fr("pica", "length", 0.0254 / 6);
Fr("sec|s", "time", 1, true);
Fr("mn|min", "time", 60);
Fr("hr", "time", 3600);
Fr("day|d", "time", 86400);
Fr("yr", "time", 365.25 * 86400);
Fr("Pa|p", "pressure", 1, true);
Fr("atm|at", "pressure", 101325);
Fr("psi", "pressure", 6894.757293168361);
Fr("N", "force", 1, true);
Fr("dyn|dy", "force", 1e-5, true);
Fr("lbf", "force", 4.4482216152605);
Fr("J", "energy", 1, true);
Fr("e", "energy", 1e-7, true);
Fr("c", "energy", 4.184, true);
Fr("cal", "energy", 4.1868, true);
Fr("Wh|wh", "energy", 3600, true);
Fr("BTU|btu", "energy", 1055.05585262);
Fr("W|w", "power", 1, true);
Fr("HP|h", "power", 745.6998715822702);
Fr("PS", "power", 735.49875);
Fr("T", "magnetism", 1, true);
Fr("ga", "magnetism", 1e-4, true);
Fr("K|kel", "temperature", 1, true);
Fr("C|cel", "temperature", 1, false, 273.15);
Fr("F|fah", "temperature", 5 / 9, false, 273.15 - 160 / 9);
Fr("Rank", "temperature", 5 / 9);
Fr("Reau", "temperature", 1.25, false, 273.15);
Fr("l|L|lt", "volume", 0.001, true);
Fr("gal", "volume", 0.003785411784);
Fr("qt", "volume", 0.003785411784 / 4);
Fr("pt|us_pt", "volume", 0.003785411784 / 8);
Fr("cup", "volume", 0.003785411784 / 16);
Fr("oz", "volume", 0.003785411784 / 128);
Fr("tbs", "volume", 0.003785411784 / 256);
Fr("tsp", "volume", 0.003785411784 / 768);
Fr("tspm", "volume", 5e-6);
Fr("uk_gal", "volume", 0.00454609);
Fr("uk_qt", "volume", 0.00454609 / 4);
Fr("uk_pt", "volume", 0.00454609 / 8);
Fr("barrel", "volume", 0.003785411784 * 42);
for (const [e, r] of [
  ["m", 1],
  ["in", 0.0254],
  ["ft", 0.3048],
  ["yd", 0.9144],
  ["mi", 1609.344],
  ["Nmi", 1852],
  ["ang", 1e-10],
])
  (Fr(`${e}2|${e}^2`, "area", r ** 2, e === "m" || e === "ang", 0, 2),
    Fr(`${e}3|${e}^3`, "volume", r ** 3, e === "m" || e === "ang", 0, 3));
Fr("ar", "area", 100, true);
Fr("ha", "area", 1e4);
Fr("uk_acre", "area", 4046.8564224);
Fr("us_acre", "area", (1200 / 3937) ** 2 * 43560);
Fr("bit", "information", 1, true);
Fr("byte", "information", 8, true);
Fr("m/s|m/sec", "speed", 1, true);
Fr("m/h|m/hr", "speed", 1 / 3600, true);
Fr("mph", "speed", 1609.344 / 3600);
Fr("kn", "speed", 1852 / 3600);
const Jre = Object.freeze(Ky.map((e) => Object.freeze(e))),
  cO = {
    Y: 24,
    Z: 21,
    E: 18,
    P: 15,
    T: 12,
    G: 9,
    M: 6,
    k: 3,
    h: 2,
    da: 1,
    e: 1,
    d: -1,
    c: -2,
    m: -3,
    u: -6,
    n: -9,
    p: -12,
    f: -15,
    a: -18,
    z: -21,
    y: -24,
  },
  Qre = { ki: 10, Mi: 20, Gi: 30, Ti: 40, Pi: 50, Ei: 60, Zi: 70, Yi: 80 };
function dO(e) {
  const r = Ky.find((t) => t.symbol === e);
  if (r) return { ...r };
  for (const [t, n] of Object.entries(Qre)) {
    if (!e.startsWith(t)) continue;
    const o = Ky.find(
      (a) => a.symbol === e.slice(t.length) && a.dimension === "information",
    );
    if (o) return { ...o, symbol: e, factor: o.factor * 2 ** n };
  }
  for (const t of Object.keys(cO).sort((n, o) => o.length - n.length)) {
    if (!e.startsWith(t)) continue;
    const n = Ky.find((o) => o.symbol === e.slice(t.length) && o.metric);
    if (n)
      return { ...n, symbol: e, factor: n.factor * 10 ** (cO[t] * n.power) };
  }
}
function ene(e, r, t) {
  if (!YL.includes(e)) return JL;
  const n = (j = "#VALUE!") => {
    throw t.issue(j);
  };
  (r.length !== (e === "CONVERT" ? 3 : 2) || r.some((j) => j == null)) && n();
  const o = r.map(t.retrieve);
  let a = 0;
  const s = (j) => {
    if (Array.isArray(j)) {
      for (const z of j) s(z);
      return;
    }
    if ((++a > 4e4 && n("#LIMIT!"), t.isIssue(j))) throw j;
    typeof j == "number" && !Number.isFinite(j) && n("#NUM!");
  };
  o.forEach(s);
  const d = (j) => (Number.isFinite(j) ? j : n("#NUM!"));
  if (e === "CONVERT") {
    const [j, z, V] = o;
    (typeof z != "string" ||
      typeof V != "string" ||
      Array.isArray(j) ||
      j == null ||
      (typeof j == "string" && !j.trim())) &&
      n();
    const D = Number(j);
    Number.isFinite(D) || n();
    const B = dO(z),
      _ = dO(V);
    return (
      (!B || !_ || B.dimension !== _.dimension) && n("#N/A"),
      B.offset === _.offset
        ? d(D * (B.factor / _.factor))
        : d((D * B.factor + B.offset - _.offset) / _.factor)
    );
  }
  const u = (j) => (Array.isArray(j) ? j.flat(1 / 0) : [j]),
    l = u(o[0]),
    c = u(o[1]);
  l.length !== c.length && n("#N/A");
  const h = l
      .map((j, z) => [j, c[z]])
      .filter((j) => j.every((z) => typeof z == "number")),
    f = h.length;
  (!f && e === "SLOPE" && n("#N/A"),
    f < (e === "STEYX" ? 3 : 2) && n("#DIV/0!"));
  const p = h.map((j) => j[0]),
    m = h.map((j) => j[1]),
    v = Us(p),
    g = Us(m),
    b = m.map((j) => d(j - g)),
    y = p.map((j) => d(j - v)),
    w = Math.max(...b.map(Math.abs)),
    E = Math.max(...y.map(Math.abs));
  if (((!w || (!E && (e === "RSQ" || e === "PEARSON"))) && n("#DIV/0!"), !E))
    return 0;
  const A = b.map((j) => j / w),
    k = y.map((j) => j / E),
    F = Rc(A.map((j) => j * j)),
    I = Rc(k.map((j) => j * j)),
    O = Rc(A.map((j, z) => j * k[z]));
  return e === "SLOPE"
    ? d((O / F) * (E / w))
    : e === "PEARSON"
      ? Math.max(-1, Math.min(1, O / Math.sqrt(F * I)))
      : e === "RSQ"
        ? Math.min(1, Math.max(0, (O / F) * (O / I)))
        : d(
            Math.sqrt(
              Rc(k.map((j, z) => (j - (O / F) * A[z]) ** 2)) / (f - 2),
            ) * E,
          );
}
const QN = [
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "AVERAGE",
    "MIN",
    "MAX",
    "MEDIAN",
    "MODE",
    "MODE.SNGL",
    "STDEV.S",
    "STDEV.P",
    "VAR.S",
    "VAR.P",
    "PERCENTILE",
    "PERCENTILE.INC",
    "PERCENTILE.EXC",
    "QUARTILE",
    "QUARTILE.INC",
    "QUARTILE.EXC",
    "RANK",
    "RANK.EQ",
    "RANK.AVG",
    "AVERAGEA",
    "MAXA",
    "MINA",
    "STDEVA",
    "STDEVPA",
    "VARA",
    "VARPA",
    "SKEW",
    "SKEW.P",
    "TRIMMEAN",
    "MODE.MULT",
    "PERMUT",
    "PERMUTATIONA",
  ],
  QL = Symbol("not-statistical-function"),
  M2 = (e) => (Array.isArray(e) ? e.flatMap(M2) : [e]),
  tne = (e) => !!e && typeof e == "object" && "ref" in e;
function H6(e) {
  let r = 0,
    t = 0;
  for (const n of e) {
    const o = r + n;
    ((t += Math.abs(r) >= Math.abs(n) ? r - o + n : n - o + r), (r = o));
  }
  return r + t;
}
function rne(e, r, t) {
  if (!QN.includes(e)) return QL;
  const n = (O = "#VALUE!") => {
      throw t.issue(O);
    },
    o = e.startsWith("RANK"),
    a = e.startsWith("PERCENTILE"),
    s = e.startsWith("QUARTILE"),
    d = e === "TRIMMEAN",
    u = e === "PERMUT" || e === "PERMUTATIONA",
    l = [
      "AVERAGEA",
      "MAXA",
      "MINA",
      "STDEVA",
      "STDEVPA",
      "VARA",
      "VARPA",
    ].includes(e);
  if (
    !r.length ||
    r.length > 255 ||
    (e === "COUNTBLANK" && r.length !== 1) ||
    ((a || s || d || u) && r.length !== 2) ||
    (o && (r.length < 2 || r.length > 3))
  )
    return n();
  const c = r.map(t.retrieve);
  let h = 0;
  const f = (O) => {
    if (Array.isArray(O)) {
      for (const j of O) f(j);
      return;
    }
    if (
      (++h > 1e5 && n("#LIMIT!"),
      t.isIssue(O) &&
        ["#LIMIT!", "#CYCLE!", "#BLOCKED!", "#ERROR!"].includes(
          String(O.error),
        ))
    )
      throw O;
  };
  for (const O of c) f(O);
  const p = (O) => (Number.isFinite(O) ? O : n("#NUM!")),
    m = (O, j) => {
      if (O === void 0 && j !== void 0) return j;
      if (t.isIssue(O)) throw O;
      if (O == null) return 0;
      if (Array.isArray(O) || (typeof O == "string" && !O.trim())) return n();
      const z = Number(O);
      return Number.isFinite(z) ? z : n();
    };
  if (e === "COUNTA") return c.flatMap(M2).filter((O) => O != null).length;
  if (e === "COUNTBLANK")
    return c.flatMap(M2).filter((O) => O == null || O === "").length;
  const v = o ? 1 : 0;
  if (u) {
    const O = Math.trunc(m(c[0])),
      j = Math.trunc(m(c[1]));
    if (O < 0 || j < 0 || (e === "PERMUT" && j > O) || (!O && j > 0))
      return n("#NUM!");
    if (!j) return 1;
    if (e === "PERMUTATIONA") return p(O ** j);
    if (j > 170) return n("#NUM!");
    let z = 1;
    for (let V = 0; V < j; V++) z = p(z * (O - V));
    return z;
  }
  const g = o || a || s || d ? [c[v]] : c,
    b = [];
  for (const [O, j] of g.entries()) {
    const z = Array.isArray(j) || tne(r[o || a || s || d ? v : O]);
    for (const V of M2(j)) {
      if (t.isIssue(V)) {
        if (e === "COUNT") continue;
        throw V;
      }
      if (typeof V == "number") b.push(p(V));
      else if (l && z && V != null)
        if (typeof V == "boolean") b.push(Number(V));
        else if (typeof V == "string") b.push(0);
        else return n();
      else if (!z && V != null) {
        if (typeof V == "boolean") b.push(Number(V));
        else if (typeof V == "string" && V.trim() && Number.isFinite(Number(V)))
          b.push(Number(V));
        else if (e !== "COUNT") return n();
      }
    }
  }
  if (e === "COUNT") return b.length;
  if (e === "AVERAGE" || e === "AVERAGEA")
    return b.length ? p(Us(b)) : n("#DIV/0!");
  if (e === "MIN" || e === "MINA")
    return b.length ? b.reduce((O, j) => Math.min(O, j)) : 0;
  if (e === "MAX" || e === "MAXA")
    return b.length ? b.reduce((O, j) => Math.max(O, j)) : 0;
  if (e.startsWith("STDEV") || e.startsWith("VAR")) {
    const O = e.endsWith(".S") || e === "STDEVA" || e === "VARA";
    if (b.length < (O ? 2 : 1)) return n("#DIV/0!");
    const j = Us(b);
    let z = b.map((_) => _ - j),
      V = z.reduce((_, C) => Math.max(_, Math.abs(C)), 0);
    if (
      (Number.isFinite(V)
        ? V && (z = z.map((_) => _ / V))
        : ((V = b.reduce((_, C) => Math.max(_, Math.abs(C)), 0)),
          (z = b.map((_) => _ / V - j / V))),
      !V)
    )
      return 0;
    const D = H6(z.map((_) => _ * _)) / (b.length - (O ? 1 : 0)),
      B = Math.sqrt(D) * V;
    return p(e.startsWith("STDEV") ? B : B * B);
  }
  if (e === "SKEW" || e === "SKEW.P") {
    const O = b.length;
    if (O < 3) return n("#DIV/0!");
    const j = Us(b),
      z = b.reduce((C, T) => Math.max(C, Math.abs(T)), 0);
    if (!z) return n("#DIV/0!");
    const V = b.map((C) => C / z - j / z),
      D = H6(V.map((C) => C * C));
    if (!D) return n("#DIV/0!");
    const _ = H6(V.map((C) => C ** 3)) / D ** 1.5;
    return p(
      e === "SKEW.P" ? _ * Math.sqrt(O) : (_ * O * Math.sqrt(O - 1)) / (O - 2),
    );
  }
  if ((b.sort((O, j) => O - j), d)) {
    const O = m(c[1]);
    if (O < 0 || O >= 1 || !b.length) return n("#NUM!");
    const j = Math.floor((b.length * O) / 2);
    return p(Us(b.slice(j, b.length - j)));
  }
  if (o) {
    const O = m(c[0]),
      j = m(c[2], 0),
      z = b.indexOf(O),
      V = b.lastIndexOf(O);
    return z < 0
      ? n("#N/A")
      : j
        ? 1 + (e === "RANK.AVG" ? (z + V) / 2 : z)
        : b.length - (e === "RANK.AVG" ? (z + V) / 2 : V);
  }
  if (e === "MODE.MULT") {
    const O = new Map();
    let j = 1;
    for (const V of b) {
      const D = (O.get(V) ?? 0) + 1;
      (O.set(V, D), (j = Math.max(j, D)));
    }
    if (j === 1) return n("#N/A");
    const z = [...O].filter(([, V]) => V === j).map(([V]) => [V]);
    return z.length > 1e4 ? n("#LIMIT!") : z;
  }
  if (e === "MODE" || e === "MODE.SNGL") {
    let O = 0,
      j = 1,
      z = 0,
      V;
    for (const D of b)
      ((z = D === V ? z + 1 : 1), z > j && ((j = z), (O = D)), (V = D));
    return j > 1 ? O : n("#N/A");
  }
  if (!b.length) return n("#NUM!");
  if (e === "MEDIAN") {
    const O = Math.floor(b.length / 2);
    return b.length % 2 || b[O - 1] === b[O] ? b[O] : b[O - 1] / 2 + b[O] / 2;
  }
  const y = m(c[1]),
    w = s ? Math.trunc(y) / 4 : y,
    E = e.endsWith(".EXC");
  if (w < 0 || w > 1 || (E && (w === 0 || w === 1))) return n("#NUM!");
  const A = E ? w * (b.length + 1) - 1 : w * (b.length - 1);
  if (A < 0 || A > b.length - 1) return n("#NUM!");
  const k = Math.floor(A),
    F = Math.ceil(A),
    I = A - k;
  return b[k] === b[F] ? b[k] : p(b[k] * (1 - I) + b[F] * I);
}
const nne = ["LOOKUP", "CHOOSE"],
  eP = Symbol("not an owned lookup function");
function ine(e, r, t) {
  if (!["VLOOKUP", "HLOOKUP", "LOOKUP", "CHOOSE"].includes(e)) return eP;
  const n = (v = "#VALUE!") => {
      throw t.issue(v);
    },
    o = (v) => {
      if (t.isIssue(v)) throw v;
      return (Array.isArray(v) && n(), v);
    },
    a = (v) => {
      if (t.isIssue(v)) throw v;
      v === void 0 && n();
      const g = Array.isArray(v) ? (v.every(Array.isArray) ? v : [v]) : [[v]];
      return (
        (!g.length ||
          !g[0].length ||
          g.some((b) => b.length !== g[0].length)) &&
          n(),
        g.length * g[0].length > 2e4 && n("#LIMIT!"),
        g
      );
    },
    s = (v) => (v.length !== 1 && v[0].length !== 1 && n(), v.flat());
  if (e === "CHOOSE") {
    (r.length < 2 || r.length > 255) && n();
    const v = o(t.retrieve(r[0]));
    typeof v == "string" && !v.trim() && n();
    const g = Math.trunc(Number(v));
    return ((!Number.isFinite(g) || g < 1 || g >= r.length) && n(), r[g] ?? 0);
  }
  (r.length < (e === "LOOKUP" ? 2 : 3) ||
    r.length > (e === "LOOKUP" ? 3 : 4)) &&
    n();
  const d = o(t.retrieve(r[0]));
  d === void 0 && n();
  const u = a(t.retrieve(r[1]));
  let l,
    c,
    h = true;
  if (e === "LOOKUP")
    r.length === 3
      ? ((l = s(u)), (c = s(a(t.retrieve(r[2])))), l.length !== c.length && n())
      : u[0].length > u.length
        ? ((l = u[0]), (c = u[u.length - 1]))
        : ((l = u.map((v) => v[0])), (c = u.map((v) => v[v.length - 1])));
  else {
    const v = o(t.retrieve(r[2]));
    typeof v == "string" && !v.trim() && n();
    const g = Math.trunc(Number(v));
    ((!Number.isFinite(g) || g < 1) && n(),
      g > (e === "VLOOKUP" ? u[0].length : u.length) && n("#REF!"));
    const b = r[3] == null ? true : o(t.retrieve(r[3]));
    (typeof b == "string" && !/^(true|false)$/i.test(b) && n(),
      (h = typeof b == "string" ? b.toLowerCase() === "true" : !!b),
      (l = e === "VLOOKUP" ? u.map((y) => y[0]) : u[0]),
      (c = e === "VLOOKUP" ? u.map((y) => y[g - 1]) : u[g - 1]));
  }
  const f = l.find(t.isIssue);
  if (f) throw f;
  const p = (v) => (typeof v == "string" ? v.toLowerCase() : v);
  let m = -1;
  if (h) {
    (!["number", "string", "boolean"].includes(typeof d) ||
      l.some((b) => typeof b != typeof d)) &&
      n();
    const v = l.map(p),
      g = p(d);
    (v.some((b, y) => y > 0 && b < v[y - 1]) && n(),
      v.forEach((b, y) => {
        b <= g && (m = y);
      }));
  } else {
    let v = 0;
    const g = typeof d == "string" && /[?*~]/.test(d);
    m = l.findIndex((b) =>
      g
        ? typeof b != "string"
          ? false
          : ((v += b.length * d.length),
            v > 1e5 && n("#LIMIT!"),
            t.matches(b, `=${d}`))
        : p(b) === p(d),
    );
  }
  return m < 0 ? n("#N/A") : (c[m] ?? 0);
}
const tP = Symbol("not an owned logical function");
function one(e, r, t) {
  if (!["AND", "OR", "XOR", "NOT"].includes(e)) return tP;
  const n = () => {
    throw t.issue("#VALUE!");
  };
  (!r.length || r.length > 255 || (e === "NOT" && r.length !== 1)) && n();
  const o = [];
  let a = 0;
  const s = (d, u) => {
    if (++a > 1e5) throw t.issue("#LIMIT!");
    if (t.isIssue(d)) throw d;
    if (Array.isArray(d)) (e === "NOT" && n(), d.forEach((l) => s(l, true)));
    else if (typeof d == "boolean" || typeof d == "number") o.push(!!d);
    else if (d == null) u || o.push(false);
    else if (typeof d == "string") {
      if (u) return;
      (/^(true|false)$/i.test(d) || n(), o.push(d.toLowerCase() === "true"));
    } else n();
  };
  return (
    r.forEach((d) =>
      s(
        t.retrieve(d),
        e !== "NOT" && !!d && typeof d == "object" && "ref" in d,
      ),
    ),
    o.length || n(),
    e === "NOT"
      ? !o[0]
      : e === "AND"
        ? o.every(Boolean)
        : e === "OR"
          ? o.some(Boolean)
          : o.filter(Boolean).length % 2 === 1
  );
}
const rP = [
    "XLOOKUP",
    "XMATCH",
    "IFS",
    "SWITCH",
    "TEXTJOIN",
    "CONCAT",
    "MAXIFS",
    "MINIFS",
    "FILTER",
    "SORT",
    "SORTBY",
    "UNIQUE",
    "SEQUENCE",
    "TRANSPOSE",
    "TAKE",
    "DROP",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "HSTACK",
    "VSTACK",
    "TOCOL",
    "TOROW",
    "EXPAND",
    "TEXTBEFORE",
    "TEXTAFTER",
    "TEXTSPLIT",
    "STDEV.P",
    "VAR.P",
    "PERCENTILE.INC",
    "PERCENTILE.EXC",
    "QUARTILE.INC",
    "QUARTILE.EXC",
    "RANK.EQ",
    "RANK.AVG",
    "MODE.SNGL",
    "MDETERM",
    "MINVERSE",
  ],
  y0 = [
    "CHOOSE",
    "FILTER",
    "SORT",
    "SORTBY",
    "UNIQUE",
    "SEQUENCE",
    "TRANSPOSE",
    "TAKE",
    "DROP",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "HSTACK",
    "VSTACK",
    "TOCOL",
    "TOROW",
    "EXPAND",
    "TEXTSPLIT",
    "MUNIT",
    "MMULT",
    "MINVERSE",
    "FREQUENCY",
    "GROWTH",
    "LINEST",
    "LOGEST",
    "TREND",
    "MODE.MULT",
    "XLOOKUP",
  ],
  p5 = Symbol("not a modern function"),
  Ec = (e) => (Array.isArray(e) ? e.flatMap(Ec) : [e]),
  Ya = (e) => (Array.isArray(e) ? (e.every(Array.isArray) ? e : [e]) : [[e]]),
  zh = (e) =>
    e == null ? "" : typeof e == "boolean" ? (e ? "TRUE" : "FALSE") : String(e),
  j2 = (e, r) =>
    typeof e == "number" && typeof r == "number"
      ? e - r
      : zh(e).toLowerCase().localeCompare(zh(r).toLowerCase()),
  fO = (e, r) =>
    typeof e == "string" && typeof r == "string"
      ? e.toLowerCase() === r.toLowerCase()
      : e === r,
  hO = (e, r, t) => {
    if (e == null || r == null) return e == null ? (r == null ? 0 : 1) : -1;
    const n = (o) => (typeof o == "number" ? 0 : typeof o == "string" ? 1 : 2);
    return (n(e) - n(r) || j2(e, r)) * t;
  };

export {
  Ec,
  GL,
  HL,
  JL,
  JN,
  Jre,
  KL,
  Kre,
  Q3,
  QL,
  QN,
  Rc,
  SE,
  Us,
  WL,
  XL,
  YL,
  Ya,
  Yre,
  ZL,
  eP,
  ene,
  fO,
  hO,
  ine,
  j2,
  jE,
  l0,
  nne,
  one,
  p5,
  rP,
  rne,
  tP,
  vg,
  y0,
  zh,
};
