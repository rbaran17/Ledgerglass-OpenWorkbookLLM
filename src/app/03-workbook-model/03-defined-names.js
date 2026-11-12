import {
  Gf,
  Pg,
  lf,
  AN,
  Nw,
  P3,
  Tw,
  iu,
} from "../01-core-schemas-and-hashing.js";
import { Ec, Ya, fO, hO, j2, p5, rP, y0, zh } from "./02-number-formats.js";

function ane(e, r, t) {
  if (!rP.includes(e)) return p5;
  const n = (h = "#VALUE!") => {
      throw t.issue(h);
    },
    o = (h, f) => {
      if (t.isIssue(h)) throw h;
      if (h === void 0 && f !== void 0) return f;
      if (Array.isArray(h) || (typeof h == "string" && !h.trim())) return n();
      const p = Number(h);
      return Number.isFinite(p) ? p : n();
    },
    a = (h, f) => Math.trunc(o(h, f)),
    s = (h, f) => {
      ((h < 1 || f < 1) && n("#CALC!"), h * f > 1e4 && n("#LIMIT!"));
    },
    d = (h) => {
      if (t.isIssue(h)) throw h;
      return (
        Array.isArray(h) && n(),
        typeof h == "string" && !/^(true|false)$/i.test(h) && n(),
        typeof h == "string" ? h.toLowerCase() === "true" : !!h
      );
    };
  if (e === "IFS") {
    (r.length < 2 || r.length % 2) && n();
    for (let h = 0; h < r.length; h += 2) {
      const f = t.retrieve(r[h]);
      if ((Array.isArray(f) && n(), d(f))) return t.retrieve(r[h + 1]) ?? 0;
    }
    return n("#N/A");
  }
  if (e === "SWITCH") {
    r.length < 3 && n();
    const h = t.retrieve(r[0]);
    if (t.isIssue(h)) return h;
    Array.isArray(h) && n();
    for (let f = 1; f + 1 < r.length; f += 2) {
      const p = t.retrieve(r[f]);
      if (t.isIssue(p)) return p;
      if ((Array.isArray(p) && n(), fO(h, p))) return t.retrieve(r[f + 1]) ?? 0;
    }
    return r.length % 2 === 0 ? (t.retrieve(r.at(-1)) ?? 0) : n("#N/A");
  }
  const u = r.map(t.retrieve),
    c = {
      SEQUENCE: [1, 4],
      FILTER: [2, 3],
      TRANSPOSE: [1, 1],
      SORT: [1, 4],
      SORTBY: [2, 255],
      UNIQUE: [1, 3],
      TAKE: [2, 3],
      DROP: [2, 3],
      HSTACK: [1, 254],
      VSTACK: [1, 254],
    }[e];
  if (
    (c && (r.length < c[0] || r.length > c[1]) && n(),
    c && e !== "SEQUENCE" && u[0] === void 0 && n(),
    c && !["HSTACK", "VSTACK"].includes(e) && t.isIssue(u[0]))
  )
    throw u[0];
  if (!y0.includes(e) && e !== "MAXIFS" && e !== "MINIFS") {
    const h = u.flatMap(Ec).find(t.isIssue);
    if (h) return h;
  }
  if (
    [
      "STDEV.P",
      "VAR.P",
      "PERCENTILE.INC",
      "PERCENTILE.EXC",
      "QUARTILE.INC",
      "QUARTILE.EXC",
      "RANK.EQ",
      "RANK.AVG",
      "MODE.SNGL",
    ].includes(e)
  ) {
    const h = e.startsWith("RANK"),
      f = (
        e === "STDEV.P" || e === "VAR.P" || e === "MODE.SNGL"
          ? u.flatMap(Ec)
          : Ec(u[h ? 1 : 0])
      )
        .filter((b) => typeof b == "number")
        .sort((b, y) => b - y);
    if ((f.length || n("#NUM!"), h)) {
      const b = o(u[0]),
        y = o(u[2], 0),
        w = f.indexOf(b),
        E = f.lastIndexOf(b);
      return (
        w < 0 && n("#N/A"),
        y
          ? 1 + (e === "RANK.AVG" ? (w + E) / 2 : w)
          : f.length - (e === "RANK.AVG" ? (w + E) / 2 : E)
      );
    }
    if (e === "MODE.SNGL") {
      const b = new Map();
      f.forEach((w) => b.set(w, (b.get(w) ?? 0) + 1));
      const y = [...b].sort((w, E) => E[1] - w[1] || w[0] - E[0]);
      return y[0][1] > 1 ? y[0][0] : n("#N/A");
    }
    if (e === "STDEV.P" || e === "VAR.P") {
      const b = f.reduce((w, E) => w + E, 0) / f.length,
        y = f.reduce((w, E) => w + (E - b) ** 2, 0) / f.length;
      return e === "STDEV.P" ? Math.sqrt(y) : y;
    }
    const p = o(u[1]) / (e.startsWith("QUARTILE") ? 4 : 1),
      m = e.endsWith("EXC"),
      v = m ? p * (f.length + 1) - 1 : p * (f.length - 1);
    (p < 0 || p > 1 || v < 0 || v > f.length - 1) && n("#NUM!");
    const g = Math.floor(v);
    return f[g] + (v - g) * ((f[g + 1] ?? f[g]) - f[g]);
  }
  if (e === "MDETERM" || e === "MINVERSE") {
    r.length !== 1 && n();
    const h = Ya(u[0]),
      f = h.length;
    (f > 32 && n("#LIMIT!"), (!f || h.some((g) => g.length !== f)) && n());
    for (const g of h.flat()) {
      if (t.isIssue(g)) return g;
      (typeof g != "number" && n(), Number.isFinite(g) || n("#NUM!"));
    }
    const p = h.map((g, b) => [
      ...g.map((y) => y),
      ...Array.from({ length: f }, (y, w) => (b === w ? 1 : 0)),
    ]);
    let m = 1;
    for (let g = 0; g < f; g++) {
      let b = g;
      for (let w = g + 1; w < f; w++)
        Math.abs(p[w][g]) > Math.abs(p[b][g]) && (b = w);
      if (p[b][g] === 0) return e === "MDETERM" ? 0 : n("#NUM!");
      b !== g && (([p[g], p[b]] = [p[b], p[g]]), (m *= -1));
      const y = p[g][g];
      ((m *= y), (p[g] = p[g].map((w) => w / y)));
      for (let w = 0; w < f; w++)
        if (w !== g) {
          const E = p[w][g];
          p[w] = p[w].map((A, k) => A - E * p[g][k]);
        }
    }
    const v = p.map((g) => g.slice(f));
    return e === "MDETERM"
      ? Number.isFinite(m)
        ? m
        : n("#NUM!")
      : (v.some((g) => g.some((b) => !Number.isFinite(b))) && n("#NUM!"), v);
  }
  if (e === "CONCAT" || e === "TEXTJOIN") {
    e === "TEXTJOIN" && u.length < 3 && n();
    const h = e === "TEXTJOIN" ? zh(u[0]) : "",
      f = e === "TEXTJOIN" && d(u[1]),
      p = u
        .slice(e === "TEXTJOIN" ? 2 : 0)
        .flatMap(Ec)
        .filter((m) => !f || (m !== null && m !== ""));
    return (
      p.reduce((m, v) => m + zh(v).length, 0) +
        Math.max(0, p.length - 1) * h.length >
        32767 && n("#LIMIT!"),
      p.map(zh).join(h)
    );
  }
  if (e === "MAXIFS" || e === "MINIFS") {
    (u.length < 3 || u.length % 2 === 0) && n();
    const h = Ec(u[0]),
      f = [];
    for (let g = 1; g < u.length; g += 2) {
      const b = Ec(u[g]);
      if (
        ((b.length !== h.length ||
          Ya(u[g]).length !== Ya(u[0]).length ||
          Ya(u[g])[0].length !== Ya(u[0])[0].length ||
          Array.isArray(u[g + 1])) &&
          n(),
        t.isIssue(u[g + 1]))
      )
        throw u[g + 1];
      f.push([b, u[g + 1]]);
    }
    h.length * f.length > 1e5 && n("#LIMIT!");
    const p = h.filter((g, b) =>
        f.every(([y, w]) => !t.isIssue(y[b]) && t.matches(y[b], w)),
      ),
      m = p.find(t.isIssue);
    if (m) return m;
    const v = p.filter((g) => typeof g == "number");
    return v.length
      ? v.reduce((g, b) => (e === "MAXIFS" ? Math.max(g, b) : Math.min(g, b)))
      : 0;
  }
  if (e === "XMATCH" || e === "XLOOKUP") {
    ((r.length < (e === "XMATCH" ? 2 : 3) ||
      r.length > (e === "XMATCH" ? 4 : 6)) &&
      n(),
      (u[0] === void 0 ||
        u[1] === void 0 ||
        (e === "XLOOKUP" && u[2] === void 0)) &&
        n());
    const h = u[0],
      f = Ya(u[1]);
    if (t.isIssue(h)) return h;
    if (t.isIssue(u[1])) return u[1];
    (Array.isArray(h) ||
      !f.length ||
      !f[0].length ||
      f.some((A) => A.length !== f[0].length) ||
      (f.length > 1 && f[0].length > 1)) &&
      n();
    const p = Ec(f),
      m = o(u[e === "XMATCH" ? 2 : 4], 0),
      v = o(u[e === "XMATCH" ? 3 : 5], 1);
    (![0, -1, 1, 2].includes(m) || ![1, -1, 2, -2].includes(v)) && n();
    const g = p.find(t.isIssue);
    if (g) return g;
    let b = [];
    if (e === "XLOOKUP") {
      if (t.isIssue(u[2])) return u[2];
      ((b = Ya(u[2])),
        (!b.length ||
          !b[0].length ||
          b.some((A) => A.length !== b[0].length)) &&
          n(),
        (f.length > 1 ? b.length !== p.length : b[0].length !== p.length) &&
          n());
    }
    (Math.abs(v) === 2 || m === -1 || m === 1) &&
      (!["number", "string", "boolean"].includes(typeof h) ||
        p.some((A) => typeof A != typeof h)) &&
      n();
    const y = p.map((A, k) => k);
    (v < 0 && y.reverse(),
      Math.abs(v) === 2 &&
        p.some((A, k) => k > 0 && j2(A, p[k - 1]) * Math.sign(v) < 0) &&
        n());
    let w = 0,
      E = y.find((A) =>
        m !== 2 || typeof h != "string"
          ? fO(p[A], h)
          : typeof p[A] != "string"
            ? false
            : ((w += p[A].length * h.length),
              w > 1e5 && n("#LIMIT!"),
              t.matches(p[A], `=${h}`)),
      );
    if (E === void 0 && (m === -1 || m === 1)) {
      const A = y.filter((k) => Math.sign(j2(p[k], h)) === m);
      (A.sort((k, F) => j2(p[k], p[F]) * m), (E = A[0]));
    }
    return E === void 0
      ? e === "XLOOKUP" && u[3] !== void 0
        ? (u[3] ?? 0)
        : n("#N/A")
      : e === "XMATCH"
        ? E + 1
        : f.length > 1
          ? b[E].length === 1
            ? (b[E][0] ?? 0)
            : [b[E]]
          : b.length === 1
            ? (b[0][E] ?? 0)
            : b.map((A) => [A[E]]);
  }
  if (e === "SEQUENCE") {
    u[0] === void 0 && u[1] === void 0 && n();
    const h = a(u[0], 1),
      f = a(u[1], 1),
      p = o(u[2], 1),
      m = o(u[3], 1);
    return (
      s(h, f),
      Array.from({ length: h }, (v, g) =>
        Array.from({ length: f }, (b, y) => p + (g * f + y) * m),
      )
    );
  }
  if (e === "FILTER") {
    const h = Ya(u[0]),
      f = Ya(u[1]);
    u[1] === void 0 && n();
    const p = Ec(f).find(t.isIssue);
    if (p) throw p;
    let m;
    if (f.length === h.length && f[0].length === 1)
      m = h.filter((v, g) => d(f[g][0]));
    else if (f.length === 1 && f[0].length === h[0].length)
      m = h.map((v) => v.filter((g, b) => d(f[0][b])));
    else return n();
    return !m.length || !m[0].length
      ? u[2] !== void 0
        ? Ya(u[2])
        : n("#CALC!")
      : m;
  }
  if (e === "TRANSPOSE")
    return Ya(u[0])[0].map((h, f) => Ya(u[0]).map((p) => p[f]));
  if (e === "SORT" || e === "SORTBY") {
    const h = Ya(u[0]).map((b) => [...b]);
    if (e === "SORTBY") {
      const b = Ya(u[1]);
      u[1] === void 0 && n();
      const y =
          b.length === 1 &&
          b[0].length === h[0].length &&
          (h.length === 1 || b[0].length > 1),
        w = y ? h[0].length : h.length,
        E = [];
      for (let k = 1; k < u.length; k += 2) {
        u[k] === void 0 && n();
        const F = Ya(u[k]),
          I = o(u[k + 1], 1);
        ((y
          ? F.length !== 1 || F[0].length !== w
          : F.length !== w || F.some((z) => z.length !== 1)) ||
          ![-1, 1].includes(I)) &&
          n();
        const O = Ec(F),
          j = O.find(t.isIssue);
        if (j) throw j;
        E.push({ data: O, order: I });
      }
      const A = Array.from({ length: w }, (k, F) => F).sort((k, F) => {
        for (const I of E) {
          const O = hO(I.data[k], I.data[F], I.order);
          if (O) return O;
        }
        return k - F;
      });
      return y ? h.map((k) => A.map((F) => k[F])) : A.map((k) => h[k]);
    }
    const f = a(u[1], 1) - 1,
      p = o(u[2], 1),
      m = u[3] === void 0 ? false : d(u[3]);
    [-1, 1].includes(p) || n();
    const v = m ? h[0].map((b, y) => h.map((w) => w[y])) : h;
    (f < 0 || f >= v[0].length) && n();
    const g = v.map((b) => b[f]).find(t.isIssue);
    if (g) throw g;
    return (
      v.sort((b, y) => hO(b[f], y[f], p)),
      m ? v[0].map((b, y) => v.map((w) => w[y])) : v
    );
  }
  if (e === "UNIQUE") {
    const h = Ya(u[0]),
      f = u[1] === void 0 ? false : d(u[1]),
      p = u[2] === void 0 ? false : d(u[2]),
      m = f ? h[0].map((w, E) => h.map((A) => A[E])) : h,
      v = (w) =>
        JSON.stringify(
          w.map((E) => (typeof E == "string" ? E.toLowerCase() : (E ?? 0))),
        ),
      g = new Map();
    m.forEach((w) => g.set(v(w), (g.get(v(w)) ?? 0) + 1));
    const b = new Set(),
      y = m.filter((w) => {
        const E = v(w);
        return b.has(E) || (p && g.get(E) !== 1) ? false : (b.add(E), true);
      });
    return (
      y.length || n("#CALC!"),
      f ? y[0].map((w, E) => y.map((A) => A[E])) : y
    );
  }
  if (e === "TAKE" || e === "DROP") {
    const h = Ya(u[0]),
      f = u[1] === void 0 ? void 0 : a(u[1]),
      p = u[2] === void 0 ? void 0 : a(u[2]),
      m = (g, b) =>
        b === void 0
          ? g
          : e === "TAKE"
            ? b > 0
              ? g.slice(0, b)
              : g.slice(b)
            : b >= 0
              ? g.slice(b)
              : g.slice(0, b);
    (f === void 0 && p === void 0 && n(), (f === 0 || p === 0) && n("#CALC!"));
    const v = m(h, f).map((g) => m(g, p));
    return ((!v.length || !v[0].length) && n("#CALC!"), v);
  }
  if (e === "CHOOSECOLS" || e === "CHOOSEROWS") {
    const h = Ya(u[0]),
      f = e === "CHOOSECOLS" ? h[0].length : h.length,
      p = u
        .slice(1)
        .flatMap(Ec)
        .map((m) => {
          const v = a(m),
            g = v > 0 ? v - 1 : f + v;
          return ((!v || g < 0 || g >= f) && n(), g);
        });
    return (
      p.length || n(),
      e === "CHOOSECOLS" ? h.map((m) => p.map((v) => m[v])) : p.map((m) => h[m])
    );
  }
  if (e === "HSTACK" || e === "VSTACK") {
    const h = u.map(Ya),
      f =
        e === "VSTACK"
          ? h.reduce((m, v) => m + v.length, 0)
          : Math.max(...h.map((m) => m.length)),
      p =
        e === "HSTACK"
          ? h.reduce((m, v) => m + v[0].length, 0)
          : Math.max(...h.map((m) => m[0].length));
    return (
      s(f, p),
      e === "VSTACK"
        ? h.flatMap((m) =>
            m.map((v) =>
              Array.from({ length: p }, (g, b) =>
                b < v.length ? v[b] : t.issue("#N/A"),
              ),
            ),
          )
        : Array.from({ length: f }, (m, v) =>
            h.flatMap((g) =>
              Array.from({ length: g[0].length }, (b, y) =>
                v < g.length ? g[v][y] : t.issue("#N/A"),
              ),
            ),
          )
    );
  }
  if (e === "TOCOL" || e === "TOROW") {
    const h = Ya(u[0]),
      f = a(u[1], 0),
      p = u[2] === void 0 ? false : d(u[2]);
    [0, 1, 2, 3].includes(f) || n();
    const m = (
      p ? h[0].flatMap((v, g) => h.map((b) => b[g])) : h.flat()
    ).filter(
      (v) =>
        !((f === 1 || f === 3) && v === null) &&
        !((f === 2 || f === 3) && t.isIssue(v)),
    );
    return (
      m.length || n("#CALC!"),
      e === "TOCOL" ? m.map((v) => [v ?? 0]) : [m.map((v) => v ?? 0)]
    );
  }
  if (e === "EXPAND") {
    const h = Ya(u[0]),
      f = a(u[1], h.length),
      p = a(u[2], h[0].length);
    return (
      s(f, p),
      (f < h.length || p < h[0].length) && n(),
      Array.from({ length: f }, (m, v) =>
        Array.from({ length: p }, (g, b) =>
          v < h.length && b < h[v].length ? h[v][b] : (u[3] ?? t.issue("#N/A")),
        ),
      )
    );
  }
  if (e === "TEXTBEFORE" || e === "TEXTAFTER") {
    const h = zh(u[0]),
      f = zh(u[1]),
      p = a(u[2], 1),
      m = a(u[3], 0) === 1,
      v = a(u[4], 0) === 1;
    (!p || !f) && n();
    const g = m ? h.toLowerCase() : h,
      b = m ? f.toLowerCase() : f,
      y = [];
    for (let E = 0; E <= g.length - b.length;) {
      const A = g.indexOf(b, E);
      if (A < 0) break;
      (y.push(A), (E = A + b.length));
    }
    const w = y[p > 0 ? p - 1 : y.length + p];
    return w === void 0
      ? v && Math.abs(p) === y.length + 1
        ? e === "TEXTBEFORE"
          ? p > 0
            ? h
            : ""
          : p > 0
            ? ""
            : h
        : u[5] !== void 0
          ? u[5]
          : n("#N/A")
      : e === "TEXTBEFORE"
        ? h.slice(0, w)
        : h.slice(w + f.length);
  }
  if (e === "TEXTSPLIT") {
    (u.length < 2 || u.length > 6) && n();
    const h = u.slice(0, 5).flatMap(Ec).find(t.isIssue);
    if (h) return h;
    const f = zh(u[0]),
      p = u[3] === void 0 ? false : d(u[3]),
      m = o(u[4], 0);
    (Array.isArray(u[0]) || ![0, 1].includes(m)) && n();
    const v = (A) =>
        Ec(A)
          .filter((k) => k != null)
          .map((k) => {
            if (t.isIssue(k)) throw k;
            return zh(k);
          })
          .filter(Boolean),
      g = v(u[1]),
      b = v(u[2]);
    (!g.length && !b.length && n(),
      (g.length + b.length > 100 || [...g, ...b].join("").length > 1e4) &&
        n("#LIMIT!"));
    const y = (A, k) =>
        k.length
          ? A.split(
              new RegExp(
                k
                  .map((F) => F.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
                  .join("|"),
                m ? "i" : "",
              ),
            ).filter((F) => !p || F !== "")
          : [A],
      w = y(f, b).map((A) => y(A, g)),
      E = Math.max(...w.map((A) => A.length));
    return (
      s(w.length, E),
      w.map((A) =>
        Array.from({ length: E }, (k, F) => A[F] ?? u[5] ?? t.issue("#N/A")),
      )
    );
  }
  return p5;
}
function eg(e, r) {
  if (e === void 0) return [];
  if (!Array.isArray(e) || e.length > 1e3)
    throw new Error("A workbook may contain at most 1,000 defined names.");
  const t = new Set();
  return e.map((n) => {
    if (!n || typeof n != "object" || Array.isArray(n))
      throw new Error("Invalid defined name.");
    const o = n;
    if (
      Object.keys(o).some(
        (s) => !["name", "refersTo", "scopeSheetId", "comment"].includes(s),
      )
    )
      throw new Error("Unknown defined-name field.");
    if (
      typeof o.name != "string" ||
      !/^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/.test(o.name) ||
      /^(?:[A-Za-z]{1,3}[1-9]\d*|R\d+C\d+|R|C)$/i.test(o.name)
    )
      throw new Error(
        "Names must begin with a letter or underscore and cannot be cell addresses.",
      );
    if (
      typeof o.refersTo != "string" ||
      !o.refersTo.startsWith("=") ||
      o.refersTo.length > 4096 ||
      o.refersTo.length < 2
    )
      throw new Error(
        "A name must refer to a formula beginning with = (up to 4,096 characters).",
      );
    if (
      o.scopeSheetId !== void 0 &&
      (typeof o.scopeSheetId != "string" ||
        !r.sheets.some((s) => s.id === o.scopeSheetId))
    )
      throw new Error("Defined-name worksheet scope was not found.");
    if (
      o.comment !== void 0 &&
      (typeof o.comment != "string" || o.comment.length > 1e3)
    )
      throw new Error(
        "Defined-name comments may contain at most 1,000 characters.",
      );
    const a = `${o.scopeSheetId ?? ""}:${o.name.toLowerCase()}`;
    if (t.has(a)) throw new Error("A name already exists in this scope.");
    return (
      t.add(a),
      {
        name: o.name,
        refersTo: o.refersTo,
        ...(o.scopeSheetId ? { scopeSheetId: o.scopeSheetId } : {}),
        ...(o.comment ? { comment: o.comment } : {}),
      }
    );
  });
}
const Ba = (e) => `${e.scopeSheetId ?? ""}:${e.name.toLowerCase()}`,
  sne = (e) => `'${e.replace(/'/g, "''")}'!`;
function lne(e, r = false) {
  const t = [];
  let n = 0;
  const o = (s) => {
      if (r) throw Error(s);
    },
    a = () => {
      let s = 0;
      do (e[n] === "[" && s++, e[n] === "]" && s--, n++);
      while (n < e.length && s);
      s && o("Unbalanced table references cannot be safely migrated.");
    };
  for (; n < e.length;) {
    const s = n,
      d = e[n];
    if (d === '"') {
      n++;
      let m = false;
      for (; n < e.length;)
        if (e[n++] === '"')
          if (e[n] === '"') n++;
          else {
            m = true;
            break;
          }
      m || o("An unterminated formula string cannot be safely migrated.");
      continue;
    }
    if (d === "#") {
      const m =
        /^#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|CYCLE!|LIMIT!|BLOCKED!|ERROR!)/i.exec(
          e.slice(n),
        );
      if (m) {
        n += m[0].length;
        continue;
      }
      n++;
      continue;
    }
    const u =
      /^(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+):(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!/.exec(
        e.slice(n),
      );
    if (u) {
      (o("Three-dimensional references cannot be safely migrated."),
        (n += u[0].length));
      continue;
    }
    let l;
    if (d === "'") {
      const m = /^'((?:[^']|'')+)'!/.exec(e.slice(n));
      if (!m) {
        for (
          o("An unsupported quoted reference cannot be safely migrated."), n++;
          n < e.length && e[n++] !== "'";
        );
        continue;
      }
      ((l = m[1].replace(/''/g, "'")),
        (n += m[0].length),
        /[\[\]:|]/.test(l) &&
          o(
            "External or three-dimensional references cannot be safely migrated.",
          ));
    } else {
      const m = /^([A-Za-z0-9_\\][A-Za-z0-9_.\\]*)!/.exec(e.slice(n));
      m && ((l = m[1]), (n += m[0].length));
    }
    if (l === void 0 && d === "[") {
      (e.startsWith("[@", n) ||
        o(
          "External or unqualified bracket references cannot be safely migrated.",
        ),
        a());
      continue;
    }
    const c =
      /^(?:\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)(?![A-Za-z0-9_.\\])/.exec(
        e.slice(n),
      );
    if (c) {
      ((n += c[0].length),
        t.push({
          start: s,
          end: n,
          text: e.slice(s, n),
          name: c[0],
          ...(l ? { qualifier: l } : {}),
          cell: true,
        }));
      continue;
    }
    if (l === void 0 && /[0-9.]/.test(d)) {
      const m = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(e.slice(n));
      if (m) {
        n += m[0].length;
        continue;
      }
    }
    const h =
      /^(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?![A-Za-z0-9_.\\])|[A-Za-z_\\][A-Za-z0-9_.\\]*)/.exec(
        e.slice(n),
      );
    if (!h) {
      (l && o("The qualified reference cannot be safely migrated."),
        d === "!" && o("The reference qualifier cannot be safely migrated."),
        d === "@" && e[n + 1] === "[" ? (n++, a()) : n++);
      continue;
    }
    const f = h[0];
    if (((n += f.length), e[n] === "[")) {
      a();
      continue;
    }
    const p = /^\s*\(/.test(e.slice(n));
    (p &&
      /^(?:_xlfn\.)?(?:INDIRECT|EVALUATE|GET\.CELL)$/i.test(f) &&
      o(
        `${f} uses dynamic or locally bound names and cannot be safely migrated.`,
      ),
      !(s > 0 && /[A-Za-z0-9_.$\\]/.test(e[s - 1])) &&
        t.push({
          start: s,
          end: n,
          text: e.slice(s, n),
          name: f,
          ...(p ? { call: true } : {}),
          ...(l ? { qualifier: l } : {}),
          ...(/^\$?[A-Za-z]{1,3}\$?[1-9]\d*$/.test(f) ? { cell: true } : {}),
        }));
  }
  return t;
}
function m5(e, r = false) {
  const t = lne(e, r);
  if (!/\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(e)) return t;
  const n = new Map();
  try {
    Gf(Pg(e), (o, a, s, d) => {
      n.set(o.start, {
        bound: new Set(a),
        excluded: d || (!o.qualifier && a.has(lf(o.name))),
      });
    });
  } catch (o) {
    if (r) throw o;
    return t;
  }
  return t
    .filter((o) => {
      var a;
      return !((a = n.get(o.start)) != null && a.excluded);
    })
    .map((o) => {
      var a;
      return { ...o, bound: (a = n.get(o.start)) == null ? void 0 : a.bound };
    });
}
function g5(e, r, t) {
  let n = "",
    o = 0;
  for (const a of r) ((n += e.slice(o, a.start) + t(a)), (o = a.end));
  return n + e.slice(o);
}
const pO = 500;
function Vg(e) {
  let r = "",
    t = 0;
  for (; t < e.length;) {
    const n = e[t];
    if (n === '"' || n === "'") {
      const a = t++,
        s = n;
      for (; t < e.length;)
        if (e[t++] === s)
          if (e[t] === s) t++;
          else break;
      r += e.slice(a, t);
      continue;
    }
    const o = /^(?:[A-Za-z_][A-Za-z0-9_.]*\[|\[@)/.exec(e.slice(t));
    if (o && (t === 0 || !/[A-Za-z0-9_.$]/.test(e[t - 1]))) {
      const a = t;
      t += o[0].length;
      let s = 1;
      for (; t < e.length && s;)
        (e[t] === "[" && s++, e[t] === "]" && s--, t++);
      const d = e.slice(a, t);
      !s && !/[!|]/.test(d) && !/^[A-Za-z0-9_!]/.test(e.slice(t))
        ? (r += "0")
        : (r += d);
      continue;
    }
    r += e[t++];
  }
  return r;
}
function $v(e, r, t, n = {}, o = false) {
  const a = new Map(Object.entries(n).map(([u, l]) => [u.toLowerCase(), l]));
  let s = "",
    d = 0;
  for (; d < e.length;) {
    const u = e[d];
    if (u === '"' || u === "'") {
      const c = d++,
        h = u;
      for (; d < e.length;)
        if (e[d++] === h)
          if (e[d] === h) d++;
          else break;
      s += e.slice(c, d);
      continue;
    }
    const l = /^([A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(e.slice(d));
    if (l && (d === 0 || !/[A-Za-z0-9_.$]/.test(e[d - 1]))) {
      const c = d;
      d += l[0].length;
      let h = 1;
      for (; d < e.length && h;)
        (e[d] === "[" && h++, e[d] === "]" && h--, d++);
      const f = e.slice(c, d),
        p = l[1] ?? "",
        m = p ? p.toLowerCase() === r.toLowerCase() : o && f.startsWith("[@");
      if (!h && m && !/[!|]/.test(f) && !/^[A-Za-z0-9_!]/.test(e.slice(d))) {
        const v = f
          .slice(p.length)
          .replace(/\[(@?)([^\[\]]+)\]/g, (g, b, y) => {
            const w = a.get(y.toLowerCase());
            return w ? `[${b}${w}]` : g;
          });
        s += `${p ? t : ""}${v}`;
      } else s += f;
      continue;
    }
    s += e[d++];
  }
  return s;
}
const une = new Set([
    "_xlnm.print_area",
    "_xlnm.print_titles",
    "_xlnm._filterdatabase",
  ]),
  Uw = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  Pi = (e) => {
    throw Error(
      `XLSX defined names: ${e} Import/export stopped to prevent changed formula bindings.`,
    );
  },
  v5 = new RegExp(
    "[\\u0000-\\u0008\\u000B\\u000C\\u000E-\\u001F\\uFFFE\\uFFFF]|[\\uD800-\\uDBFF](?![\\uDC00-\\uDFFF])|(?<![\\uD800-\\uDBFF])[\\uDC00-\\uDFFF]",
  );
function nP(e) {
  return e.includes("<") || e.includes("]]>") || v5.test(e)
    ? Pi("Malformed XML text.")
    : e.replace(/&[^;\s<]*;|&/g, (r) => {
        const t = {
          "&amp;": "&",
          "&lt;": "<",
          "&gt;": ">",
          "&quot;": '"',
          "&apos;": "'",
        };
        if (t[r]) return t[r];
        if (!/^&#(?:[0-9]+|x[0-9a-fA-F]+);$/.test(r))
          return Pi("Unknown or malformed XML entity.");
        const n = r.startsWith("&#x"),
          o = Number.parseInt(r.slice(n ? 3 : 2, -1), n ? 16 : 10);
        if (
          !Number.isSafeInteger(o) ||
          o > 1114111 ||
          o < 0 ||
          (o >= 55296 && o <= 57343)
        )
          return Pi("Invalid XML character entity.");
        const a = String.fromCodePoint(o);
        return v5.test(a) ? Pi("Invalid XML character entity.") : a;
      });
}
function Yy(e, r = false) {
  return v5.test(e)
    ? Pi("A name contains an invalid XML character.")
    : e.replace(
        /[&<>"'\r\n\t]/g,
        (t) =>
          ({
            "&": "&amp;",
            "<": "&lt;",
            ">": "&gt;",
            '"': "&quot;",
            "'": "&apos;",
          })[t] ?? (r ? `&#${t.charCodeAt(0)};` : t === "\r" ? "&#13;" : t),
      );
}
function b5(e) {
  const r = Object.create(null);
  let t = e;
  for (; t.trim();) {
    const n = /^\s+([A-Za-z_][A-Za-z0-9_.:-]*)\s*=\s*(["'])([\s\S]*?)\2/.exec(
      t,
    );
    if (!n || Object.hasOwn(r, n[1]))
      return Pi("Malformed or duplicate defined-name XML attribute.");
    ((r[n[1]] = nP(n[3].replace(/\r\n?|\n|\t/g, " "))),
      (t = t.slice(n[0].length)));
  }
  return r;
}
function iP(e) {
  if (/<!DOCTYPE|<!ENTITY/i.test(e))
    return Pi("DTD and entity declarations are not supported.");
  const r = [
    ...e.matchAll(
      /<((?:[A-Za-z_][\w.-]*:)?definedNames)\b((?:[^"'<>]|"[^"]*"|'[^']*')*?)(?:\/>|>([\s\S]*?)<\/\1\s*>)/g,
    ),
  ];
  if (r.length > 1) return Pi("Duplicate definedNames sections.");
  const t = r[0],
    n = t ? e.slice(0, t.index) + e.slice(t.index + t[0].length) : e;
  if (/<\/?(?:[A-Za-z_][\w.-]*:)?definedNames?\b/.test(n))
    return Pi("Malformed or misplaced defined-name elements.");
  if (!t) return;
  const o = b5(t[2]),
    a =
      /<(?:[A-Za-z_][\w.-]*:)?workbook\b((?:[^"'<>]|"[^"]*"|'[^']*')*?)>/.exec(
        e,
      ),
    s = { ...(a ? b5(a[1]) : {}), ...o },
    d = t[1].slice(0, -12);
  if (s[d ? `xmlns:${d.slice(0, -1)}` : "xmlns"] !== Uw)
    return Pi("Unsupported or undeclared definedNames namespace.");
  for (const u of Object.keys(o))
    if (!u.startsWith("xmlns") || o[u] !== Uw)
      return Pi("Unsupported definedNames namespace or attribute.");
  return {
    start: t.index,
    end: t.index + t[0].length,
    prefix: t[1].slice(0, -12),
    body: t[3] ?? "",
    namespaces: s,
  };
}
function pC(e) {
  if (!e.startsWith("=") || e.length < 2 || e.length > 4096)
    return Pi(
      "A definition must contain a formula of 1–4,095 characters after =.",
    );
  let r = 0,
    t = 0,
    n = 0,
    o = 1,
    a = "";
  for (; o < e.length;) {
    const d = e[o];
    if (!t && (d === '"' || d === "'")) {
      const u = d,
        l = o++;
      let c = false;
      for (; o < e.length;)
        if (e[o++] === u)
          if (e[o] === u) o++;
          else {
            c = true;
            break;
          }
      if (!c) return Pi("Unterminated formula string or worksheet qualifier.");
      if (u === "'") {
        if (e[o] !== "!") return Pi("Malformed quoted worksheet qualifier.");
        a += e.slice(l, o);
      } else a += '""';
      continue;
    }
    if ((d === "[" && t++, d === "]" && --t < 0))
      return Pi("Unbalanced bracket reference.");
    if (!t) {
      if ((d === "{" && n++, d === "}" && --n < 0))
        return Pi("Unbalanced array braces.");
      if ((d === "(" && r++, d === ")" && --r < 0))
        return Pi("Unbalanced formula parentheses.");
    }
    ((a += d), o++);
  }
  if (r || t || n)
    return Pi(
      "Unbalanced formula parentheses, array braces or bracket reference.",
    );
  const s = Vg(a);
  if (
    /[\[\]|]/.test(s) ||
    /(?:https?:|file:|javascript:|data:|\\\\)/i.test(s) ||
    /\b(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|EXECUTE|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY|EVALUATE|GET\.CELL)\s*\(/i.test(
      s,
    )
  )
    return Pi("External or executable name formulas are not supported.");
  if (
    /(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+):(?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!/.test(
      s,
    )
  )
    return Pi("Three-dimensional name references are not supported.");
  if (!a.trim()) return Pi("An empty formula is not supported.");
}
function cne(e, r) {
  var n;
  const t = Vg(e.replace(/"(?:[^"]|"")*"/g, '""'));
  for (const o of t.matchAll(
    /(?:'((?:[^']|'')+)'|([A-Za-z_\\][A-Za-z0-9_.\\]*))!/g,
  )) {
    const a = ((n = o[1]) == null ? void 0 : n.replace(/''/g, "'")) ?? o[2];
    if (!r.sheets.some((s) => s.name.toLowerCase() === a.toLowerCase()))
      return Pi(`A name refers to an unknown worksheet "${a}".`);
  }
}
function fne(e, r) {
  const t = iP(e),
    n = r ? `<definedNames xmlns="${Uw}">${r}</definedNames>` : "";
  if (t) return e.slice(0, t.start) + n + e.slice(t.end);
  if (!r) return e;
  const o =
    /<(?:[A-Za-z_][\w.-]*:)?(?:calcPr|oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/(?:[A-Za-z_][\w.-]*:)?workbook\s*>/.exec(
      e,
    );
  return o
    ? e.slice(0, o.index) + n + e.slice(o.index)
    : Pi("The workbook XML has no supported insertion point.");
}
const hne = ["LET", "LAMBDA"],
  wv = { depth: 64, calls: 1024, work: 1e6, retainedValues: 1e5 },
  X6 = (e) =>
    !!e && typeof e == "object" && "kind" in e && e.kind === "formula-closure",
  y5 = "__LGVALUE";
function pne(e, r, t, n) {
  let o = 0,
    a = 0,
    s = 0;
  const d = new Map(),
    u = new Set(),
    l = (E, A) => {
      throw (n.diagnostic(E, A), n.issue(E));
    },
    c = (E, A) => {
      ((o += E),
        (A > wv.depth || o > wv.work) &&
          l(
            "#LIMIT!",
            `Reusable formula evaluation exceeds ${wv.depth} levels or ${wv.work.toLocaleString()} work units.`,
          ));
    },
    h = (E) => {
      const A = Array.isArray(E)
        ? E.reduce((k, F) => k + (Array.isArray(F) ? F.length : 1), 0)
        : 1;
      return (
        (s += A),
        s > wv.retainedValues &&
          l(
            "#LIMIT!",
            "Reusable formula intermediate values exceed 100,000 cells.",
          ),
        E
      );
    },
    f = (E) => {
      let A = d.get(E);
      if (!A) {
        c(E.length, 0);
        try {
          try {
            pC(E.startsWith("=") ? E : `=${E}`);
          } catch (k) {
            const F = k instanceof Error ? k.message : "Unsafe formula.";
            l(
              /External|executable|Three-dimensional/i.test(F)
                ? "#BLOCKED!"
                : "#VALUE!",
              F,
            );
          }
          ((A = Pg(E)),
            Gf(A, (k, F, I, O) => {
              (!O &&
                /^__LG/i.test(k.name) &&
                l(
                  "#NAME?",
                  "Internal calculation names cannot be called from a workbook.",
                ),
                /^(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|EXECUTE|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY|EVALUATE|GET\.CELL)$/i.test(
                  k.name,
                ) &&
                  l(
                    "#BLOCKED!",
                    "External or executable custom functions are not supported.",
                  ));
            }));
        } catch (k) {
          if (n.isIssue(k)) throw k;
          l(
            k instanceof Error && k.message === "#LIMIT!"
              ? "#LIMIT!"
              : "#VALUE!",
            k instanceof Error ? k.message : "Invalid reusable formula.",
          );
        }
        d.set(E, A);
      }
      return A;
    },
    p = (E, A) => {
      var O, j, z;
      const k = E.qualifier
        ? (O = r.sheets.find(
            (V) => V.name.toLowerCase() === E.qualifier.toLowerCase(),
          )) == null
          ? void 0
          : O.id
        : A;
      if (!k) return;
      const F = lf(E.name),
        I =
          ((j = r.definedNames) == null
            ? void 0
            : j.find(
                (V) => V.scopeSheetId === k && V.name.toLowerCase() === F,
              )) ??
          ((z = r.definedNames) == null
            ? void 0
            : z.find((V) => !V.scopeSheetId && V.name.toLowerCase() === F));
      return I ? { found: I, sheetId: I.scopeSheetId ?? k } : void 0;
    },
    m = (E, A, k) => {
      try {
        return n.parse(E, A, k);
      } catch (F) {
        const I =
          F && typeof F == "object" && "details" in F ? F.details : void 0;
        return n.isIssue(I) ? I : n.isIssue(F) ? F : n.issue("#ERROR!");
      }
    },
    v = (E, A, k, F, I) => {
      const O = lf(E.name);
      if (!E.qualifier && k.has(O)) return k.get(O);
      if (!E.qualifier && /^(?:TRUE|FALSE)$/i.test(E.name)) return O === "true";
      const j = p(E, F);
      if (!j)
        return l(
          "#NAME?",
          `Unknown name or unsupported custom function: ${A.slice(E.start, E.end)}.`,
        );
      const z = `${j.sheetId}:${j.found.name.toLowerCase()}`;
      if (n.allowName && !n.allowName(j.sheetId))
        return l(
          "#BLOCKED!",
          "The defined name is outside the authorized inspection context.",
        );
      if (u.has(z))
        return l("#CYCLE!", `Circular defined name: ${j.found.name}.`);
      u.add(z);
      try {
        return w(
          f(j.found.refersTo),
          j.found.refersTo,
          new Map(),
          j.sheetId,
          I + 1,
        );
      } finally {
        u.delete(z);
      }
    },
    g = (E, A, k) => {
      if (!X6(E))
        return l(
          "#VALUE!",
          "Only a LAMBDA formula can be called with arguments.",
        );
      if (A.length !== E.parameters.length)
        return l(
          "#VALUE!",
          `LAMBDA expects ${E.parameters.length} argument(s); received ${A.length}.`,
        );
      if (++a > wv.calls)
        return l("#LIMIT!", "Reusable formula evaluation exceeds 1,024 calls.");
      const F = new Map(E.environment);
      return (
        E.parameters.forEach((I, O) => F.set(I, h(A[O]))),
        w(E.body, E.source, F, E.sheetId, k + 1)
      );
    },
    b = (E, A, k, F, I) => {
      var _, C, T, N;
      if ((c(1, I), E.kind === "symbol")) return v(E, A, k, F, I);
      if (E.kind === "group") return w(E.body, A, k, F, I + 1);
      if (E.kind === "text") return m(E.text, F, []);
      const O = AN(E);
      if (O === "LET") {
        const R = new Map(k);
        for (let L = 0; L < E.args.length - 1; L += 2) {
          const $ = w(E.args[L + 1], A, R, F, I + 1),
            U = lf(Nw(E.args[L]).name);
          (R.set(U, h($)),
            (_ = n.step) == null || _.call(n, "binding", `LET ${U}`, $));
        }
        return w(E.args.at(-1), A, R, F, I + 1);
      }
      if (O === "LAMBDA")
        return {
          kind: "formula-closure",
          parameters: E.args.slice(0, -1).map((R) => lf(Nw(R, true).name)),
          body: E.args.at(-1),
          source: A,
          environment: new Map(k),
          sheetId: F,
        };
      const j =
          E.callee.kind === "symbol" && !E.callee.qualifier
            ? P3(E.callee.name)
            : "",
        z =
          E.callee.kind === "symbol" &&
          !E.callee.qualifier &&
          k.has(lf(E.callee.name));
      if (!z && j === "IF" && E.args.length >= 2 && E.args.length <= 3) {
        const R = w(E.args[0], A, k, F, I + 1);
        if (n.isIssue(R)) return R;
        if (!Array.isArray(R)) {
          if (typeof R == "string" && !/^(?:TRUE|FALSE)$/i.test(R))
            return n.issue("#VALUE!");
          const L = typeof R == "string" ? R.toUpperCase() === "TRUE" : !!R;
          return (
            (C = n.step) == null ||
              C.call(
                n,
                "branch",
                `IF: ${L ? "true" : "false"} branch selected`,
                R,
              ),
            E.args[L ? 1 : 2] ? w(E.args[L ? 1 : 2], A, k, F, I + 1) : false
          );
        }
      }
      if (!z && /^(?:IFERROR|IFNA)$/.test(j) && E.args.length === 2) {
        const R = w(E.args[0], A, k, F, I + 1);
        if (!Array.isArray(R)) {
          if (
            n.isIssue(R) &&
            ["#LIMIT!", "#BLOCKED!", "#CYCLE!", "#ERROR!"].includes(R.error)
          )
            return R;
          const L = n.isIssue(R) && (j === "IFERROR" || R.error === "#N/A");
          return (
            (T = n.step) == null ||
              T.call(
                n,
                "branch",
                `${j}: ${L ? "fallback" : "value"} selected`,
                R,
              ),
            L ? w(E.args[1], A, k, F, I + 1) : R
          );
        }
      }
      if (!z && j && n.supported.has(j)) {
        const R = [],
          L = E.args.map(($) => y($, A, k, F, I + 1, R));
        return m(`${j}(${L.join(",")})`, F, R);
      }
      const V = b(E.callee, A, k, F, I + 1),
        D = E.args.map((R) => (Tw(R).length ? w(R, A, k, F, I + 1) : null)),
        B = g(V, D, I);
      return (
        (N = n.step) == null ||
          N.call(n, "function", A.slice(E.callee.start, E.callee.end), B),
        B
      );
    },
    y = (E, A, k, F, I, O) =>
      E.map((j) => {
        if (j.kind === "text") {
          const V = /^(\d+(?:\.\d*)?|\.\d+)[eE]([+-]?\d+)$/.exec(j.text);
          return V ? `(${V[1]}*10^(${Number(V[2])}))` : j.text;
        }
        const z = b(j, A, k, F, I);
        return X6(z)
          ? l(
              "#CALC!",
              "A LAMBDA must be called with arguments or saved as a reusable name.",
            )
          : (h(z),
            O.push(
              Array.isArray(z)
                ? z.map((V) => (Array.isArray(V) ? [...V] : V))
                : z,
            ),
            `${y5}(${O.length - 1})`);
      }).join(""),
    w = (E, A, k, F, I) => {
      try {
        c(E.length, I);
        const O = Tw(E);
        if (O.length === 1 && O[0].kind !== "text") return b(O[0], A, k, F, I);
        const j = [];
        return m(y(E, A, k, F, I, j), F, j);
      } catch (O) {
        if (
          n.isIssue(O) &&
          !["#LIMIT!", "#BLOCKED!", "#CYCLE!", "#ERROR!"].includes(O.error)
        )
          return O;
        throw O;
      }
    };
  try {
    const E = w(f(e), e, new Map(), t, 0);
    return X6(E)
      ? l(
          "#CALC!",
          "A LAMBDA must be called with arguments or saved as a reusable name.",
        )
      : E;
  } catch (E) {
    if (n.isIssue(E)) return E;
    throw E;
  }
}
function jne(e, r = {}) {
  return e.sheets
    .filter((t) => !r.allowedSheetIds || r.allowedSheetIds.includes(t.id))
    .map((t) => {
      var n;
      return {
        sheetId: t.id,
        name: t.name,
        visibility: t.veryHidden
          ? "veryHidden"
          : t.hidden
            ? "hidden"
            : "visible",
        tabColor: t.tabColor ?? null,
        storedCellCount: Object.keys(t.cells).length,
        formulaCellCount: Object.values(t.cells).filter(
          (o) =>
            !o.literal && typeof o.raw == "string" && o.raw.startsWith("="),
        ).length,
        calculationParticipation: "included-when-referenced",
        structureLocked: !!((n = e.structureProtection) != null && n.enabled),
      };
    });
}
const fp = () => crypto.randomUUID();
function ea(e) {
  throw new Error(e);
}
const Fne = (e) =>
  !!e &&
  typeof e == "object" &&
  !Array.isArray(e) &&
  [null, Object.prototype].includes(Object.getPrototypeOf(e));
function bP(e) {
  return (
    (typeof e != "string" ||
      !e.trim() ||
      e.length > 31 ||
      /[\\/\[\]*?:\u0000-\u001f]/.test(e) ||
      /^'|'$/.test(e)) &&
      ea(
        "Sheet names must contain 1–31 characters and cannot contain control characters, Excel-reserved characters, or begin or end with an apostrophe.",
      ),
    e
  );
}
function rR(e, r, t) {
  const n = bP(r);
  return (
    e.sheets.some(
      (o) => o.id !== t && o.name.toLowerCase() === n.toLowerCase(),
    ) && ea("A worksheet with that name already exists."),
    n
  );
}
const qp = (e, r) =>
    e.sheets.find((t) => t.id === r) ?? ea("The worksheet was not found."),
  zw = (e) => !e.hidden && !e.veryHidden;
function yP(e) {
  return (
    (typeof e != "string" || !/^#[a-f\d]{6}$/i.test(e)) &&
      ea("Tab colors must use a six-digit hexadecimal color."),
    e.toLowerCase()
  );
}
function wP(e, r) {
  (bP(r.name),
    e.veryHidden !== void 0 &&
      (typeof e.veryHidden != "boolean" &&
        ea("The very-hidden flag must be boolean."),
      e.veryHidden &&
        e.hidden === false &&
        ea("A very-hidden worksheet must also be hidden."),
      (r.veryHidden = e.veryHidden),
      e.veryHidden && (r.hidden = true)),
    e.tabColor !== void 0 && (r.tabColor = yP(e.tabColor)));
}
function nR(e, r) {
  const t = e;
  if (t !== void 0) {
    (!Fne(t) ||
      t.enabled !== true ||
      Object.keys(t).some(
        (a) => !["enabled", "passwordHash", "salt"].includes(a),
      )) &&
      ea("Invalid workbook structure protection.");
    const n = t.passwordHash !== void 0,
      o = t.salt !== void 0;
    ((n !== o ||
      (n &&
        (typeof t.passwordHash != "string" ||
          !/^[a-f\d]{64}$/i.test(t.passwordHash) ||
          typeof t.salt != "string" ||
          t.salt.length < 16 ||
          t.salt.length > 100))) &&
      ea(
        "Workbook structure protection requires a valid salted password hash.",
      ),
      (r.structureProtection = {
        enabled: true,
        ...(n
          ? {
              passwordHash: String(t.passwordHash).toLowerCase(),
              salt: String(t.salt),
            }
          : {}),
      }));
  }
  r.sheets.some(zw) || ea("At least one worksheet must remain visible.");
}
function $a(e) {
  var r;
  (r = e.structureProtection) != null &&
    r.enabled &&
    ea(
      "Workbook structure is protected. Unprotect the structure before changing worksheets.",
    );
}
const S5 = String.raw`(?:'(?:[^']|'')+'|[^\s+\-*/^&=<>(),;:'"\[\]!]+)`,
  wO = new RegExp(`^(${S5})!`),
  Bne = new RegExp(`^(${S5}):(${S5})!`),
  J6 = (e) => (e.startsWith("'") ? e.slice(1, -1).replace(/''/g, "'") : e),
  Lne = (e) => `'${e.replace(/'/g, "''")}'`;
function T5(e, r) {
  let t = r,
    n = 0;
  do (e[t] === "[" && n++, e[t] === "]" && n--, t++);
  while (t < e.length && n > 0);
  return n === 0 ? t : -1;
}
function Q6(e, r) {
  const t = Math.max(
    0,
    ...[
      /^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?/,
      /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}/,
      /^\$?[1-9]\d*:\$?[1-9]\d*/,
      /^[A-Za-z_\\][A-Za-z0-9_.\\]*/,
    ].map((o) => {
      var a;
      return ((a = o.exec(e.slice(r))) == null ? void 0 : a[0].length) ?? 0;
    }),
  );
  if (!t) return r;
  let n = r + t;
  return (
    e[n] === "[" && (n = T5(e, n)),
    n < r || /[\w.$]/.test(e[n] ?? "") || /^\s*\(/.test(e.slice(n))
      ? r
      : (e[n] === "#" && n++, n)
  );
}
function vC(e) {
  const r = [];
  let t = 0;
  for (; t < e.length;) {
    const n = t,
      o = e[t];
    if (o === '"') {
      for (t++; t < e.length;)
        if (e[t++] === '"')
          if (e[t] === '"') t++;
          else break;
      continue;
    }
    const a = t === 0 || !/[\w.$]/.test(e[t - 1]);
    a &&
      /^(?:_xlfn\.)?INDIRECT\s*\(/i.test(e.slice(t)) &&
      r.push({
        start: n,
        end: n,
        text: "INDIRECT",
        unsupported: "Dynamic INDIRECT references cannot be safely repaired.",
      });
    const s = a ? Bne.exec(e.slice(t)) : null;
    if (s) {
      t += s[0].length;
      const l = Q6(e, t);
      ((t = Math.max(t, l)),
        r.push({
          start: n,
          end: t,
          text: e.slice(n, t),
          sheets: [J6(s[1]), J6(s[2])],
          unsupported:
            "Three-dimensional worksheet references require explicit references before this operation.",
        }));
      continue;
    }
    const d = a ? wO.exec(e.slice(t)) : null;
    if (d) {
      const l = J6(d[1]);
      t += d[0].length;
      const c = Q6(e, t),
        h = e.slice(t, c);
      if (((t = Math.max(t, c)), l.includes("[") || l.includes("]"))) continue;
      r.push({
        start: n,
        end: t,
        text: e.slice(n, t),
        ...(l.includes(":")
          ? {
              sheets: l.split(":"),
              unsupported:
                "Three-dimensional worksheet references require explicit references before this operation.",
            }
          : {
              sheet: l,
              reference: h,
              ...(h
                ? {}
                : {
                    unsupported:
                      "This worksheet-qualified expression is not a supported reference.",
                  }),
            }),
      });
      continue;
    }
    const u = a ? /^([A-Za-z_][A-Za-z0-9_.]*)\[/.exec(e.slice(t)) : null;
    if (u) {
      const l = T5(e, t + u[1].length);
      if (l > 0) {
        ((t = l),
          r.push({ start: n, end: l, text: e.slice(n, l), table: u[1] }));
        continue;
      }
    }
    if (o === "[") {
      const l = T5(e, t);
      if (l > 0) {
        t = l;
        const c = wO.exec(e.slice(t));
        c && ((t += c[0].length), (t = Math.max(t, Q6(e, t))));
        continue;
      }
    }
    if (o === "'") {
      for (t++; t < e.length;)
        if (e[t++] === "'")
          if (e[t] === "'") t++;
          else break;
      continue;
    }
    t++;
  }
  return r;
}
function xP(e, r, t) {
  if (e.sheet) return e.sheet.toLowerCase() === t.name.toLowerCase();
  if (e.sheets) {
    const n = e.sheets.map((a) =>
        r.sheets.findIndex((s) => s.name.toLowerCase() === a.toLowerCase()),
      ),
      o = r.sheets.findIndex((a) => a.id === t.id);
    return n.some((a) => a < 0) || (o >= Math.min(...n) && o <= Math.max(...n));
  }
  return !!e.unsupported;
}
function e0(e, r, t, n, o = false) {
  const a = new Set((t.tables ?? []).map((u) => u.name.toLowerCase()));
  let s = "",
    d = 0;
  for (const u of vC(e)) {
    const l = xP(u, r, t),
      c = o && u.table && a.has(u.table.toLowerCase());
    (u.unsupported && l && ea(u.unsupported),
      !(!l && !c) &&
        ((s += e.slice(d, u.start)),
        (s += n !== void 0 ? `${Lne(n)}!${u.reference}` : "#REF!"),
        (d = u.end)));
  }
  return s + e.slice(d);
}
function EP(e, r, t) {
  for (const n of e.sheets)
    if (n.id !== t) {
      for (const o of n.tables ?? [])
        for (const a of o.columns)
          a.calculatedFormula && r(a.calculatedFormula, n.id);
      for (const o of [
        ...(n.conditionalFormats ?? []),
        ...(n.validations ?? []),
      ])
        (o.kind === "formula" || o.kind === "custom") &&
          o.formula &&
          r(o.formula, n.id, o.range.split(":")[0]);
      for (const [o, a] of Object.entries(n.cells))
        !a.literal &&
          typeof a.raw == "string" &&
          a.raw.startsWith("=") &&
          r(a.raw, n.id, o);
    }
  for (const n of e.definedNames ?? [])
    r(n.refersTo, n.scopeSheetId ?? "", void 0);
}
function $ne(e, r) {
  for (let t = 2; t < Number.MAX_SAFE_INTEGER; t++) {
    const n = ` (${t})`,
      o = r.slice(0, 31 - n.length) + n;
    if (!e.sheets.some((a) => a.name.toLowerCase() === o.toLowerCase()))
      return o;
  }
  return ea("No unique worksheet name is available.");
}
function AP(e, r) {
  const t = qp(e, r),
    n = [],
    o = new Set((t.tables ?? []).map((s) => s.name.toLowerCase())),
    a = (s, d) => {
      const u = vC(s),
        l = u.filter(
          (c) => xP(c, e, t) || (c.table && o.has(c.table.toLowerCase())),
        );
      l.some((c) => c.unsupported)
        ? n.push({
            ...d,
            kind: "unsupported-reference",
            label: `${d.label}: ${l.find((c) => c.unsupported).unsupported}`,
          })
        : l.length && n.push(d);
    };
  for (const s of e.sheets)
    if (s.id !== r) {
      for (const d of s.tables ?? [])
        for (const u of d.columns)
          u.calculatedFormula &&
            a(u.calculatedFormula, {
              kind: "table-column",
              label: `${s.name}: ${d.name}[${u.name}] calculated rule will contain #REF!`,
              sheetId: s.id,
            });
      for (const [d, u] of Object.entries(s.cells))
        !u.literal &&
          typeof u.raw == "string" &&
          u.raw.startsWith("=") &&
          a(u.raw, {
            kind: "formula",
            label: `${s.name}!${d} will contain #REF!`,
            sheetId: s.id,
            address: d,
          });
      for (const d of [
        ...(s.conditionalFormats ?? []),
        ...(s.validations ?? []),
      ])
        (d.kind === "formula" || d.kind === "custom") &&
          d.formula &&
          a(d.formula, {
            kind: "formula",
            label: `${s.name}: ${d.kind === "custom" ? "validation" : "conditional"} rule ${d.id} will contain #REF!`,
            sheetId: s.id,
            address: d.range.split(":")[0],
          });
      for (const d of s.pivots ?? [])
        d.source.sheetId === r &&
          n.push({
            kind: "pivot-refresh",
            label: `${s.name}: ${d.name} refresh settings will be removed; output cells retained`,
            sheetId: s.id,
          });
      for (const d of s.queries ?? [])
        d.source.sheetId === r &&
          n.push({
            kind: "query-refresh",
            label: `${s.name}: ${d.name} refresh settings will be removed; output cells retained`,
            sheetId: s.id,
          });
    }
  for (const s of e.definedNames ?? [])
    s.scopeSheetId === r
      ? n.push({
          kind: "local-name",
          label: `${t.name}!${s.name} local name will be removed`,
          sheetId: r,
        })
      : a(s.refersTo, {
          kind: "defined-name",
          label: `${s.name} will contain #REF!`,
          ...(s.scopeSheetId ? { sheetId: s.scopeSheetId } : {}),
        });
  for (const s of e.formulaWatches ?? [])
    !s.orphaned &&
      s.anchor.sheetId === r &&
      n.push({
        kind: "formula-watch",
        label: `${t.name}!${s.anchor.address} watch will be retained as a deleted-sheet orphan`,
        sheetId: r,
        address: s.anchor.address,
      });
  return n;
}
function kP(e) {
  return (
    (typeof e != "string" || e.length > 200) &&
      ea("Passwords must be text of at most 200 characters."),
    e
  );
}
async function SP(e, r) {
  const t = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(`workbook-structure\0${r}\0${e}`),
  );
  return [...new Uint8Array(t)]
    .map((n) => n.toString(16).padStart(2, "0"))
    .join("");
}
const TP = ["any", "text", "number", "boolean", "date"];
function iR(e) {
  if (!TP.includes(e)) throw Error("Choose a supported table column type.");
  return e;
}
function qne(e, r) {
  var n;
  if (!((n = e.tableEvidence) != null && n.length)) return r;
  const t = new Map(e.tableEvidence.map((o) => [o.id, o]));
  for (const o of r.tableEvidence ?? []) {
    const a = t.get(o.id);
    if (a && JSON.stringify(a) !== JSON.stringify(o))
      throw Error("Saved table source evidence cannot be changed.");
  }
  return {
    ...r,
    tableEvidence: [
      ...e.tableEvidence,
      ...(r.tableEvidence ?? []).filter((o) => !t.has(o.id)),
    ],
  };
}
const N5 = (e) => e === null || e === "";
function NP(e) {
  if (
    typeof e != "string" ||
    !/^\d{4}-\d{2}-\d{2}$/.test(e) ||
    e.startsWith("0000")
  )
    return false;
  const r = Date.parse(e + "T00:00:00.000Z");
  return Number.isFinite(r) && new Date(r).toISOString().slice(0, 10) === e;
}
function Vw(e, r = "any") {
  return (
    r === "any" ||
    N5(e) ||
    (r === "date"
      ? NP(e)
      : typeof e === (r === "text" ? "string" : r) &&
        (r !== "number" || Number.isFinite(e)))
  );
}
function Xne(e, r) {
  if (Vw(e, r.dataType)) return e;
  if (r.dataType === "text") return String(e);
  if (typeof e != "string")
    throw Error("This value cannot be converted without guessing.");
  const t = e.trim();
  if (r.dataType === "boolean" && /^(true|false)$/i.test(t))
    return /^true$/i.test(t);
  if (r.dataType === "date" && NP(t)) return t;
  if (r.dataType === "number") {
    const n = r.decimal ?? ".",
      o = r.group ?? ",",
      a = (c) => c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
      s = o ? `(?:\\d+|\\d{1,3}(?:${a(o)}\\d{3})+)` : "\\d+";
    if (!new RegExp(`^[+-]?(?:${s}(?:${a(n)}\\d+)?|${a(n)}\\d+)$`).test(t))
      throw Error(
        "Use a number with the selected decimal/group separators; currency, percentages and scientific text require a separate cleanup.",
      );
    const d = (o ? t.split(o).join("") : t).replace(n, "."),
      u = Number(d),
      l = d
        .replace(/^[+-]?0*/, "")
        .replace(".", "")
        .replace(/0+$/, "");
    if (
      !Number.isFinite(u) ||
      l.length > 15 ||
      (Number.isInteger(u) && !Number.isSafeInteger(u))
    )
      throw Error(
        "Conversion would exceed safe numeric precision. Keep the original text.",
      );
    return u;
  }
  throw Error(
    r.dataType === "date"
      ? "Dates must be valid YYYY-MM-DD text; no locale or date-system inference."
      : "Boolean text must be TRUE or FALSE.",
  );
}
const oR = [
  "column",
  "bar",
  "line",
  "area",
  "pie",
  "donut",
  "scatter",
  "combo",
  "waterfall",
  "histogram",
  "sparkline",
];
function Rl(e) {
  throw Error(e);
}
const CO = (e) => !!e && typeof e == "object" && !Array.isArray(e);
function IP(e) {
  if (e === void 0) return {};
  (!CO(e) ||
    Object.keys(e).some(
      (t) =>
        ![
          "altText",
          "xTitle",
          "yTitle",
          "units",
          "yScale",
          "xScale",
          "showLegend",
          "series",
          "includedRows",
          "histogramBins",
          "sparklineKind",
        ].includes(t),
    )) &&
    Rl("Invalid or unknown chart settings.");
  const r = {};
  for (const t of ["altText", "xTitle", "yTitle", "units"])
    e[t] !== void 0 &&
      ((typeof e[t] != "string" ||
        e[t].length > (t === "altText" ? 2e3 : 200)) &&
        Rl(`Invalid chart ${t}.`),
      (r[t] = e[t]));
  for (const t of ["xScale", "yScale"])
    e[t] !== void 0 &&
      (["linear", "log"].includes(String(e[t])) ||
        Rl("Chart scales must be linear or log."),
      (r[t] = e[t]));
  if (
    (e.showLegend !== void 0 &&
      (typeof e.showLegend != "boolean" && Rl("Invalid chart legend setting."),
      (r.showLegend = e.showLegend)),
    e.histogramBins !== void 0 &&
      ((!Number.isInteger(e.histogramBins) ||
        Number(e.histogramBins) < 1 ||
        Number(e.histogramBins) > 50) &&
        Rl("Choose 1 to 50 histogram bins."),
      (r.histogramBins = Number(e.histogramBins))),
    e.sparklineKind !== void 0 &&
      (["line", "column", "winLoss"].includes(String(e.sparklineKind)) ||
        Rl("Invalid sparkline style."),
      (r.sparklineKind = e.sparklineKind)),
    e.series !== void 0)
  ) {
    (!Array.isArray(e.series) || e.series.length > 20) &&
      Rl("Select at most 20 chart series.");
    const t = new Set();
    r.series = e.series.map((n) => {
      ((!CO(n) ||
        Object.keys(n).some((a) => !["column", "name", "kind"].includes(a)) ||
        !Number.isInteger(n.column) ||
        Number(n.column) < 1 ||
        Number(n.column) > 16384 ||
        t.has(Number(n.column))) &&
        Rl("Chart series require unique one-based worksheet columns."),
        t.add(Number(n.column)));
      const o = { column: Number(n.column) };
      return (
        n.name !== void 0 &&
          ((typeof n.name != "string" ||
            !n.name.trim() ||
            n.name.length > 200) &&
            Rl("Chart series names require 1 to 200 characters."),
          (o.name = n.name)),
        n.kind !== void 0 &&
          (["column", "line"].includes(String(n.kind)) ||
            Rl("Combo series must be column or line."),
          (o.kind = n.kind)),
        o
      );
    });
  }
  return (
    e.includedRows !== void 0 &&
      ((!Array.isArray(e.includedRows) ||
        e.includedRows.length > 500 ||
        e.includedRows.some(
          (t) => !Number.isInteger(t) || t < 1 || t > 1048576,
        ) ||
        new Set(e.includedRows).size !== e.includedRows.length) &&
        Rl(
          "Chart filters require at most 500 unique one-based worksheet rows.",
        ),
      (r.includedRows = [...e.includedRows].sort((t, n) => t - n))),
    r
  );
}
function aR(e) {
  const r = e.diagnostics.filter((t) => t.severity === "error");
  r.length &&
    Rl(
      r
        .slice(0, 5)
        .map((t) => t.message)
        .join(" "),
    );
}
const Kne = 100,
  Hw = 2 * 1024 * 1024,
  Yne = 500,
  R5 = (e) => !!e && typeof e == "object" && !Array.isArray(e),
  L4 = (e, r, t, n) => {
    if (typeof e != "number" || !Number.isFinite(e) || e < r || e > t)
      throw new Error(`Invalid drawing ${n}.`);
    return e;
  },
  zv = (e, r, t) => {
    if (typeof e != "string" || e.length > r)
      throw new Error(`Invalid drawing ${t}.`);
    return e;
  },
  AO = (e) => {
    if (typeof e != "string" || !/^#[\da-f]{6}$/i.test(e))
      throw new Error("Drawing colors must be six-digit hex colors.");
    return e;
  };
function FE(e) {
  if (!(e instanceof Uint8Array) || e.length > Hw)
    throw Error("Picture hashing requires at most 2 MB of bytes.");
  return iu(e);
}

export {
  $a,
  $ne,
  $v,
  AO,
  AP,
  Ba,
  EP,
  FE,
  Hw,
  IP,
  Kne,
  L4,
  N5,
  Pi,
  R5,
  Rl,
  SP,
  TP,
  Uw,
  Vg,
  Vw,
  Xne,
  Yne,
  Yy,
  aR,
  ane,
  b5,
  cne,
  e0,
  ea,
  eg,
  fne,
  fp,
  g5,
  hne,
  iP,
  iR,
  jne,
  kP,
  m5,
  nP,
  nR,
  oR,
  pC,
  pO,
  pne,
  qne,
  qp,
  rR,
  sne,
  une,
  vC,
  wP,
  y5,
  yP,
  zv,
  zw,
};
