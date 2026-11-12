import {
  WS,
  iB,
  lx,
  v1,
  zp,
  Rg,
  es,
  g1,
  ts,
  wd,
  hr,
} from "../01-core-schemas-and-hashing.js";
import { Ja, WJ, aB, iD, qS, u6 } from "../02-financial-functions.js";
import {
  string,
  number,
  boolean,
  array,
  strictObject,
  discriminatedUnion,
  literal,
  union,
  enum as _enum,
  null as _null,
} from "zod";
import { Bw, zB } from "./01-cell-styles-and-themes.js";
import {
  GL,
  HL,
  JL,
  KL,
  Kre,
  QL,
  Rc,
  Us,
  Yre,
  eP,
  ene,
  ine,
  jE,
  one,
  p5,
  rne,
  tP,
  vg,
  y0,
} from "./02-number-formats.js";
import { Ba, ane, eg, g5, m5, pO } from "./03-defined-names.js";
import {
  Gl,
  M5,
  Mie,
  Rie,
  dp,
  jie,
  kR,
  nU,
  oU,
  pf,
  wn,
} from "./04-pivot-tables-and-pictures.js";

function Ty(e, r, t = {}) {
  var b;
  if (t.readOnly) throw Error("This workbook is read-only.");
  if ((b = e.structureProtection) != null && b.enabled)
    throw Error(
      "Workbook structure is protected. Unprotect it before changing names.",
    );
  const n = eg(e.definedNames, e),
    o =
      r.kind === "create"
        ? undefined
        : n.find(
            (y) =>
              y.name.toLowerCase() === r.name.toLowerCase() &&
              y.scopeSheetId === r.scopeSheetId,
          );
  if (r.kind !== "create" && !o)
    throw Error("The defined name was not found. Reopen Name Manager.");
  let a;
  if (r.kind === "create" || r.kind === "update") a = eg([r.definition], e)[0];
  else if (r.kind === "rename") a = eg([{ ...o, name: r.newName }], e)[0];
  else if (r.kind !== "delete") throw Error("Unknown name operation.");
  if (
    r.kind === "update" &&
    (o.name !== a.name || o.scopeSheetId !== a.scopeSheetId)
  )
    throw Error(
      "Use Rename to change a name. Existing name scope cannot be changed; create a separate name instead.",
    );
  if (
    a &&
    e.sheets.some((y) => {
      return y.tables?.some(
        (E) => E.name.toLowerCase() === a.name.toLowerCase(),
      );
    })
  )
    throw Error("A table already uses this name.");
  const s = structuredClone(e),
    d = o ? Ba(o) : undefined;
  ((s.definedNames = eg(
    [...n.filter((y) => Ba(y) !== d), ...(a ? [a] : [])],
    e,
  )),
    o && a && (s.definedNames = n.map((y) => (Ba(y) === d ? a : y))));
  const u = a?.scopeSheetId ?? o?.scopeSheetId,
    l = (y) => {
      var E;
      if (
        (E = e.sheets.find((A) => A.id === y)?.protection) != null &&
        E.enabled
      )
        throw Error(
          "An affected worksheet is protected. Unprotect it before changing names.",
        );
    };
  u && l(u);
  const c = [],
    h = [];
  let f = 0;
  const p = true,
    m = (y) => {
      const w = m5(y, p);
      if (((f += w.length + 1), f > 1e6))
        throw Error(
          "The name dependency scan exceeds the supported budget. Split the workbook before changing names.",
        );
      if (
        w.some(
          (E) =>
            E.qualifier &&
            !e.sheets.some(
              (A) => A.name.toLowerCase() === E.qualifier.toLowerCase(),
            ),
        )
      )
        throw Error(
          "An unknown or external worksheet qualifier cannot be safely migrated.",
        );
      return w;
    },
    v = (y, w, E = []) => {
      if (E.length > 32)
        throw Error("The name dependency chain exceeds 32 levels.");
      for (const A of m(y)) {
        const k = dp(A, e, w),
          F = r.kind === "create" ? dp(A, s, w) : undefined;
        if (F && Ba(F) === Ba(a) && (!k || Ba(k) !== Ba(F))) return true;
        if (!k) {
          if (r.kind === "create" && dp(A, s, w) && Ba(dp(A, s, w)) === Ba(a))
            return true;
          continue;
        }
        if (Ba(k) === d) return true;
        const I = Ba(k);
        if (
          !E.includes(I) &&
          v(
            k.refersTo,
            k.scopeSheetId ??
              (A.qualifier
                ? e.sheets.find(
                    (O) => O.name.toLowerCase() === A.qualifier.toLowerCase(),
                  ).id
                : w),
            [...E, I],
          )
        )
          return true;
      }
      return false;
    },
    g = (y, w, E) => {
      const A = m(y);
      if (w.some((I) => v(y, I))) {
        if (
          (h.push({
            label: E.label,
            ...(E.sheetId ? { sheetId: E.sheetId } : {}),
            ...(E.address ? { address: E.address } : {}),
          }),
          h.length > pO)
        )
          throw Error(
            "A name operation may affect at most 500 dependent locations. Split the operation into smaller workbooks.",
          );
        if ((E.sheetId && l(E.sheetId), r.kind === "delete"))
          throw Error(
            `This name is referenced by ${E.label}. Remove or replace its dependencies before deletion.`,
          );
      }
      const F = g5(y, A, (I) => {
        var z;
        if (r.kind !== "rename") return I.text;
        const O = w.map((V) => dp(I, e, V)),
          j = O.map((V) => !!V && Ba(V) === d);
        if (j.some(Boolean) && !j.every(Boolean))
          throw Error(
            `The workbook-scoped definition ${E.label} resolves differently across worksheets. Qualify its references before renaming.`,
          );
        if (j.every(Boolean)) {
          if (
            !I.qualifier &&
            (z = I.bound) != null &&
            z.has(a.name.toLowerCase())
          )
            throw Error(
              `Renaming would be shadowed by a local LET variable or LAMBDA parameter in ${E.label}. Choose a distinct name.`,
            );
          const V = { ...I, name: a.name };
          if (
            w.some((D) => {
              const B = dp(V, s, D);
              return !B || Ba(B) !== Ba(a);
            })
          )
            throw Error(
              `Renaming would be shadowed by another name in ${E.label}. Choose a distinct name.`,
            );
          return I.qualifier
            ? I.text.slice(0, I.text.length - I.name.length) + a.name
            : a.name;
        }
        if (
          w.some((V, D) => {
            const B = dp(I, s, V);
            return (O[D] ? Ba(O[D]) : undefined) !== (B ? Ba(B) : undefined);
          })
        )
          throw Error(
            `Renaming would change an existing reference in ${E.label}. Choose a distinct name.`,
          );
        return I.text;
      });
      if (
        F !== y &&
        (E.sheetId && l(E.sheetId),
        c.push({ ...E, before: y, after: F }),
        c.length > pO)
      )
        throw Error(
          "A name operation may rewrite at most 500 formulas or definitions. Split the migration into smaller workbooks.",
        );
      return F;
    };
  for (const y of s.sheets) {
    for (const w of y.tables ?? [])
      for (const E of w.columns)
        E.calculatedFormula &&
          (E.calculatedFormula = g(E.calculatedFormula, [y.id], {
            kind: "table-column",
            label: `${y.name}: ${w.name}[${E.name}] calculated rule`,
            sheetId: y.id,
          }));
    for (const [w, E] of Object.entries(y.cells))
      typeof E.raw == "string" &&
        E.raw.startsWith("=") &&
        !E.literal &&
        (E.raw = g(E.raw, [y.id], {
          kind: "cell",
          label: `${y.name}!${w}`,
          sheetId: y.id,
          address: w,
        }));
    for (const w of [...(y.conditionalFormats ?? []), ...(y.validations ?? [])])
      (w.kind === "formula" || w.kind === "custom") &&
        w.formula &&
        (w.formula = g(w.formula, [y.id], {
          kind: w.kind === "custom" ? "validation" : "conditional",
          label: `${y.name}: ${w.kind === "custom" ? "validation" : "conditional"} rule ${w.id}`,
          sheetId: y.id,
        }));
  }
  for (const y of s.definedNames) {
    const w = y.scopeSheetId ? [y.scopeSheetId] : e.sheets.map((E) => E.id);
    y.refersTo = g(y.refersTo, w, {
      kind: "name",
      label: `Name ${y.name} (${y.scopeSheetId ? e.sheets.find((E) => E.id === y.scopeSheetId).name : "Workbook"})`,
      ...(y.scopeSheetId ? { sheetId: y.scopeSheetId } : {}),
    });
  }
  if (a)
    for (const y of a.scopeSheetId
      ? [a.scopeSheetId]
      : e.sheets.map((w) => w.id))
      try {
        pf(
          s.definedNames.find((w) => Ba(w) === Ba(a)).refersTo.slice(1),
          s,
          y,
          [Ba(a)],
        );
      } catch (w) {
        throw Error(
          `The resulting name has a cycle or exceeds the expansion limit: ${w instanceof Error ? w.message : String(w)}`,
        );
      }
  return {
    snapshot: JSON.stringify(e),
    request: structuredClone(r),
    ...(o ? { before: structuredClone(o) } : {}),
    ...(a
      ? { after: structuredClone(s.definedNames.find((y) => Ba(y) === Ba(a))) }
      : {}),
    changes: c,
    dependencies: h,
    next: s,
  };
}
function q6(e, r, t = {}) {
  if (JSON.stringify(e) !== r.snapshot)
    throw Error(
      "The workbook changed after the name preview. Review the operation again.",
    );
  const n = Ty(e, r.request, t);
  if (JSON.stringify(n) !== JSON.stringify(r))
    throw Error("The name preview changed. Review the operation again.");
  return JSON.stringify(n.next) === JSON.stringify(e) ? e : n.next;
}
function Xb(e, r) {
  if (typeof r != "string") return e === r;
  const t = /^(<=|>=|<>|=|<|>)?(.*)$/.exec(r),
    n = t[1] ?? "=",
    o = t[2];
  if ((n === "=" || n === "<>") && /[?*~]/.test(o)) {
    if (typeof e != "string") return n === "<>";
    const u = [];
    for (let h = 0; h < o.length; h++)
      o[h] === "~" && h + 1 < o.length
        ? u.push({ char: o[++h].toLowerCase(), wildcard: false })
        : u.push({
            char: o[h].toLowerCase(),
            wildcard: o[h] === "*" || o[h] === "?",
          });
    const l = String(e ?? "").toLowerCase();
    if (u.length * l.length > 1e5) throw jie("#LIMIT!");
    let c = Array(l.length + 1).fill(false);
    c[0] = true;
    for (const h of u) {
      const f = Array(l.length + 1).fill(false);
      f[0] = h.wildcard && h.char === "*" && c[0];
      for (let p = 1; p <= l.length; p++)
        f[p] =
          h.wildcard && h.char === "*"
            ? c[p] || f[p - 1]
            : c[p - 1] &&
              ((h.wildcard && h.char === "?") || l[p - 1] === h.char);
      c = f;
    }
    return n === "<>" ? !c[l.length] : c[l.length];
  }
  const a = o.trim() !== "" && Number.isFinite(Number(o)),
    s = a
      ? e === null || e === ""
        ? 0
        : typeof e == "number"
          ? e
          : NaN
      : String(e ?? "").toLowerCase(),
    d = a ? Number(o) : o.toLowerCase();
  return n === "<>"
    ? s !== d
    : n === "<"
      ? s < d
      : n === ">"
        ? s > d
        : n === "<="
          ? s <= d
          : n === ">="
            ? s >= d
            : s === d;
}
function Fie(e, r, t = kR, n = lx(), o, a) {
  const s = e.utils._applyInfix.bind(e.utils);
  ((e.utils._applyInfix = (u, l, c) => {
    const h = (w, E) => (
      ["=", "<>", "<", ">", "<=", ">="].includes(l) &&
        typeof w == "string" &&
        typeof E == "string" &&
        ((w = w.toLowerCase()), (E = E.toLowerCase())),
      s({ val: w, isArray: false }, l, { val: E, isArray: false })
    );
    if (!Array.isArray(u.val) && !Array.isArray(c.val)) return h(u.val, c.val);
    const f = (w) =>
        Array.isArray(w) ? (w.every(Array.isArray) ? w : [w]) : [[w]],
      p = f(u.val),
      m = f(c.val),
      v = Math.max(p.length, m.length),
      g = Math.max(p[0]?.length ?? 0, m[0]?.length ?? 0);
    return !v ||
      !g ||
      (p.length !== v && p.length !== 1) ||
      (m.length !== v && m.length !== 1) ||
      (p[0].length !== g && p[0].length !== 1) ||
      (m[0].length !== g && m[0].length !== 1)
      ? t("#VALUE!")
      : v * g > Ja.maxSpillCells
        ? t("#LIMIT!")
        : Array.from({ length: v }, (w, E) =>
            Array.from({ length: g }, (A, k) =>
              h(
                p[p.length === 1 ? 0 : E][p[0].length === 1 ? 0 : k],
                m[m.length === 1 ? 0 : E][m[0].length === 1 ? 0 : k],
              ),
            ),
          );
  }),
    (e.checkFormulaResult = (u, l) => {
      const c = e.retrieveRef(u);
      return wn(c)
        ? c
        : Array.isArray(c)
          ? l
            ? c
            : t("#VALUE!")
          : c === null
            ? 0
            : typeof c == "number"
              ? Number.isFinite(c)
                ? c
                : t("#NUM!")
              : typeof c == "string" || typeof c == "boolean"
                ? c
                : t("#VALUE!");
    }));
  const d = e._callFunction.bind(e);
  return (
    (e._callFunction = (u, l) => {
      var h;
      u = u
        .replace(/^_xlfn\./i, "")
        .replace(/^_xlws\./i, "")
        .toUpperCase();
      const c = (f) => (f == null ? undefined : e.retrieveRef(f));
      try {
        if (!Mie.has(u)) return t("#NAME?");
        if (u === "RAND") return l.length ? t("#VALUE!") : Math.random();
        if (u === "RANDBETWEEN") {
          if (l.length !== 2) return t("#VALUE!");
          const D = l.map(c),
            B = M5(D);
          if (B) return B;
          if (
            D.some(
              (T) =>
                Array.isArray(T) ||
                T === undefined ||
                (typeof T == "string" && T.trim() === "") ||
                !Number.isFinite(Number(T)),
            )
          )
            return t("#VALUE!");
          const _ = Math.ceil(Number(D[0])),
            C = Math.floor(Number(D[1]));
          return !Number.isSafeInteger(_) ||
            !Number.isSafeInteger(C) ||
            _ > C ||
            !Number.isSafeInteger(C - _ + 1)
            ? t("#NUM!")
            : _ + Math.floor(Math.random() * (C - _ + 1));
        }
        const f = WS(u, l, { retrieve: c, issue: t, isIssue: wn, clock: n });
        if (f !== iB) return f;
        const p = WJ(u, l, {
          retrieve: c,
          issue: t,
          isIssue: wn,
          clock: n,
          report: o,
        });
        if (p !== aB) return p;
        const m = r(u, l, e);
        if (m !== oU) return m;
        const v = Rie(u, l, {
          retrieve: c,
          issue: t,
          isIssue: wn,
          criterionFormula: a,
        });
        if (v !== nU) return v;
        const g = ene(u, l, { retrieve: c, issue: t, isIssue: wn });
        if (g !== JL) return g;
        const b = jE(u, l, { retrieve: c, issue: t, isIssue: wn });
        if (b !== GL) return b;
        const y = Yre(u, l, {
          retrieve: c,
          issue: t,
          isIssue: wn,
          distribution: d,
        });
        if (y !== KL) return y;
        const w = Kre(u, l, { retrieve: c, issue: t, isIssue: wn });
        if (w !== HL) return w;
        const E = rne(u, l, { retrieve: c, issue: t, isIssue: wn });
        if (E !== QL) return E;
        const A = one(u, l, { retrieve: c, issue: t, isIssue: wn });
        if (A !== tP) return A;
        const k = ine(u, l, {
          retrieve: c,
          issue: t,
          isIssue: wn,
          matches: Xb,
        });
        if (k !== eP) return k;
        if (u === "TEXTBEFORE" || u === "TEXTAFTER")
          return u6(
            l.map(c),
            (D) => {
              try {
                return iD(u, D, {
                  retrieve: (B) => B,
                  issue: t,
                  isIssue: wn,
                  format: () => t("#VALUE!"),
                });
              } catch (B) {
                return wn(B) ? B : t("#VALUE!");
              }
            },
            t,
          );
        const F = iD(u, l, {
          retrieve: c,
          issue: t,
          isIssue: wn,
          temporal: (D) => zp(D, n),
          format: () => {
            const D = Number(c(l[0]) ?? 0),
              B = c(l[1]);
            if (typeof D == "number" && typeof B == "string") {
              try {
                if (Bw(B).kind === "temporal")
                  return vg(
                    D,
                    { category: "custom", locale: "en-US", pattern: B },
                    n.dateSystem,
                  ).text;
              } catch {}
              const _ = B.replace(/"[^"]*"|\\./g, "");
              if (
                n.dateSystem === "1904" &&
                /[ydm]/i.test(_) &&
                !/\[[hms]+\]/i.test(_)
              )
                return d(u, [D + v1(n.dateSystem), B]);
            }
            return d(u, l);
          },
        });
        if (F !== qS) return F;
        const I = ane(u, l, {
          retrieve: c,
          issue: t,
          isIssue: wn,
          matches: Xb,
        });
        if (I !== p5) return I;
        if (u === "INDEX") {
          if (l.length < 2 || l.length > 4) return t("#VALUE!");
          if (wn(l[0])) return l[0];
          const D = [...l];
          for (let _ = 1; _ < l.length; _++) {
            const C = l[_] == null ? (_ === 3 ? 1 : 0) : c(l[_]);
            if (wn(C)) return C;
            if (
              Array.isArray(C) ||
              (typeof C == "string" && C.trim() === "") ||
              !Number.isFinite(Number(C)) ||
              Number(C) < 0
            )
              return t("#VALUE!");
            D[_] = Number(C);
          }
          if (D.length > 3 && Number(D[3]) < 1) return t("#VALUE!");
          const B = d(u, D);
          if (B && typeof B == "object" && "ref" in B) {
            const _ = l[0],
              C = B;
            (h = _?.ref) != null &&
              h.sheet &&
              !C.ref.sheet &&
              (C.ref.sheet = _.ref.sheet);
          }
          return B;
        }
        if (u === "IF") {
          if (l.length < 2 || l.length > 3) return t("#VALUE!");
          const D = c(l[0]);
          if (wn(D)) return D;
          if (Array.isArray(D))
            return u6(
              [D, c(l[1]), l.length === 3 ? c(l[2]) : false],
              ([_, C, T]) =>
                wn(_)
                  ? _
                  : typeof _ == "string" && !/^(TRUE|FALSE)$/i.test(_)
                    ? t("#VALUE!")
                    : (((
                        typeof _ == "string" ? _.toUpperCase() === "TRUE" : !!_
                      )
                        ? C
                        : T) ?? 0),
              t,
            );
          if (typeof D == "string" && !/^(TRUE|FALSE)$/i.test(D))
            return t("#VALUE!");
          const B = typeof D == "string" ? D.toUpperCase() === "TRUE" : !!D;
          return c(B ? l[1] : l.length === 3 ? l[2] : false) ?? 0;
        }
        if (u === "IFERROR" || u === "IFNA") {
          if (l.length !== 2) return t("#VALUE!");
          const D = c(l[0]);
          return Array.isArray(D)
            ? u6(
                [D, c(l[1])],
                ([B, _]) =>
                  wn(B) && (u === "IFERROR" || B.error === "#N/A")
                    ? (_ ?? 0)
                    : (B ?? 0),
                t,
              )
            : wn(D) && (u === "IFERROR" || D.error === "#N/A")
              ? (c(l[1]) ?? 0)
              : (D ?? 0);
        }
        const O = l.map(c),
          j = O.flatMap(Gl).filter((D) => typeof D == "number");
        if (
          O.flatMap(Gl).length > 1e5 ||
          ([
            "REPT",
            "BASE",
            "DEC2BIN",
            "DEC2HEX",
            "DEC2OCT",
            "BIN2HEX",
            "BIN2OCT",
            "HEX2BIN",
            "HEX2OCT",
            "OCT2BIN",
            "OCT2HEX",
            "FIXED",
            "DOLLAR",
          ].includes(u) &&
            ((Math.abs(Number(O[u === "BASE" ? 2 : 1] ?? 0)) > 32767 &&
              !["FIXED", "DOLLAR"].includes(u)) ||
              (["FIXED", "DOLLAR"].includes(u) &&
                Math.abs(Number(O[1] ?? 2)) > 100) ||
              (u === "REPT" &&
                String(O[0] ?? "").length * Number(O[1]) > 32767))) ||
          (["MUNIT", "MDETERM", "MMULT", "GROWTH"].includes(u) &&
            (u === "MUNIT"
              ? Number(O[0]) > 100
              : O.flatMap(Gl).length > 1e3)) ||
          ([
            "BINOM.DIST.RANGE",
            "BINOM.INV",
            "HYPGEOM.DIST",
            "POISSON.DIST",
            "FACT",
            "FACTDOUBLE",
            "COMBIN",
            "COMBINA",
            "MULTINOMIAL",
          ].includes(u) &&
            j.some((D) => Math.abs(D) > 1e4)) ||
          (["BESSELI", "BESSELJ", "BESSELK", "BESSELY"].includes(u) &&
            Math.abs(Number(O[1])) > 1e3)
        )
          return t("#LIMIT!");
        if (u === "ROMAN" && Math.abs(Number(O[0])) > 3999) return t("#VALUE!");
        if (
          [
            "NETWORKDAYS",
            "NETWORKDAYS.INTL",
            "WORKDAY",
            "WORKDAY.INTL",
          ].includes(u) &&
          (Math.abs(
            Number(O[1]) - (u.startsWith("NETWORK") ? Number(O[0]) : 0),
          ) > 366e3 ||
            O.flatMap(Gl).length > 1e3)
        )
          return t("#LIMIT!");
        if (u === "ISERR") return wn(O[0]) && O[0].error !== "#N/A";
        if (u === "ISLOGICAL") return typeof O[0] == "boolean";
        if (u === "ISNONTEXT") return typeof O[0] != "string";
        if (u === "N")
          return wn(O[0]) || typeof O[0] == "number"
            ? O[0]
            : O[0] === true
              ? 1
              : 0;
        if (u === "T") return wn(O[0]) || typeof O[0] == "string" ? O[0] : "";
        if (u === "TYPE")
          return wn(O[0])
            ? 16
            : Array.isArray(O[0])
              ? 64
              : typeof O[0] == "string"
                ? 2
                : typeof O[0] == "boolean"
                  ? 4
                  : 1;
        if (u === "ERROR.TYPE") {
          const D = [
            "#NULL!",
            "#DIV/0!",
            "#VALUE!",
            "#REF!",
            "#NAME?",
            "#NUM!",
            "#N/A",
            "#GETTING_DATA",
          ].indexOf(wn(O[0]) ? O[0].error : "");
          return D < 0 ? t("#N/A") : D + 1;
        }
        if (u === "ISERROR") return wn(O[0]);
        if (u === "ISNA") return wn(O[0]) && O[0].error === "#N/A";
        if (u === "ISNUMBER") return typeof O[0] == "number";
        if (u === "ISTEXT") return typeof O[0] == "string";
        if (u === "ISBLANK") return O[0] === null;
        if (u === "COUNTA")
          return O.flatMap(Gl).filter((D) => D !== null).length;
        if (u === "COUNTBLANK")
          return O.flatMap(Gl).filter((D) => D === null || D === "").length;
        if (
          [
            "SUMIF",
            "AVERAGEIF",
            "COUNTIF",
            "SUMIFS",
            "AVERAGEIFS",
            "COUNTIFS",
          ].includes(u)
        ) {
          const D = u.endsWith("IFS"),
            B = u.startsWith("COUNT"),
            _ = u.startsWith("AVERAGE");
          let C, T;
          if (D) {
            if (((T = B ? O : O.slice(1)), T.length < 2 || T.length % 2 !== 0))
              return t("#VALUE!");
            C = B ? Gl(T[0]).map(() => 1) : Gl(O[0]);
          } else {
            if (O.length < 2 || O.length > (B ? 2 : 3)) return t("#VALUE!");
            ((T = O.slice(0, 2)),
              (C = B
                ? Gl(O[0]).map(() => 1)
                : Gl(l.length < 3 || l[2] == null ? O[0] : O[2])));
          }
          const N = [],
            R = [],
            L = (H) => {
              if (!Array.isArray(H)) return "1:1";
              const Q = H.every(Array.isArray) ? H : [H],
                X = Q[0]?.length ?? 0;
              return Q.length && X && Q.every((q) => q.length === X)
                ? `${Q.length}:${X}`
                : "invalid";
            },
            $ = L(B ? T[0] : D || l.length < 3 || l[2] == null ? O[0] : O[2]);
          for (let H = 0; H < T.length; H += 2) {
            const Q = Gl(T[H]);
            if (
              Q.length !== C.length ||
              $ === "invalid" ||
              L(T[H]) !== $ ||
              Array.isArray(T[H + 1])
            )
              return t("#VALUE!");
            if (wn(T[H + 1])) return T[H + 1];
            N.push(Q);
            const X = T[H + 1];
            R.push(X == null ? "=0" : typeof X == "number" ? `=${X}` : X);
          }
          const U = [];
          let G = 0;
          for (let H = 0; H < C.length; H++) {
            if (
              ((G += N.reduce(
                (Q, X, J) =>
                  Q +
                  Math.max(1, String(X[H] ?? "").length) *
                    Math.max(1, String(R[J] ?? "").length),
                0,
              )),
              G > 2e6)
            )
              return t("#LIMIT!");
            if (N.every((Q, X) => !wn(Q[H]) && Xb(Q[H], R[X]))) {
              if (wn(C[H])) return C[H];
              typeof C[H] == "number" && U.push(C[H]);
            }
          }
          return _ ? (U.length ? Us(U) : t("#DIV/0!")) : Rc(U);
        }
        const z = M5(O);
        if (z && u !== "COUNT") return z;
        if (u === "TRIM")
          return String(O[0] ?? "")
            .replace(/^ +| +$/g, "")
            .replace(/ +/g, " ");
        if (u === "UPPER") {
          const D = String(O[0] ?? "").toUpperCase();
          return D.length > 32767 ? t("#LIMIT!") : D;
        }
        if (u === "VALUE") {
          if (O.length !== 1 || Array.isArray(O[0])) return t("#VALUE!");
          let D = String(O[0] ?? "").trim(),
            B = false;
          D.startsWith("(") &&
            D.endsWith(")") &&
            ((B = true), (D = D.slice(1, -1).trim()));
          const _ = D.endsWith("%");
          if (
            (_ && (D = D.slice(0, -1).trim()),
            (D = D.replace(/^\$/, "")),
            !/^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+|\d*\.\d+)(?:\.\d*)?(?:[eE][+-]?\d+)?$/.test(
              D,
            ))
          )
            return t("#VALUE!");
          const C =
            (Number(D.replace(/,/g, "")) * (B ? -1 : 1)) / (_ ? 100 : 1);
          return Number.isFinite(C) ? C : t("#NUM!");
        }
        if (u === "MATCH") {
          if (O.length < 2 || O.length > 3) return t("#VALUE!");
          const D = O[1];
          if (
            Array.isArray(D) &&
            D.length > 1 &&
            Array.isArray(D[0]) &&
            D[0].length > 1
          )
            return t("#N/A");
          const B = Gl(D),
            _ = O[0],
            C = O[2] === undefined ? 1 : Number(O[2]);
          if (![0, -1, 1].includes(C)) return t("#N/A");
          if (C === 0) {
            const $ = B.findIndex((U) =>
              typeof _ == "string"
                ? typeof U == "string" &&
                  (/[?*~]/.test(_)
                    ? Xb(U, `=${_}`)
                    : U.toLowerCase() === _.toLowerCase())
                : U === _,
            );
            return $ < 0 ? t("#N/A") : $ + 1;
          }
          const T = B.map(($) => (typeof $ == "string" ? $.toLowerCase() : $)),
            N = typeof _ == "string" ? _.toLowerCase() : _;
          if (
            (typeof N != "number" && typeof N != "string") ||
            T.some(($) => typeof $ != typeof N)
          )
            return t("#VALUE!");
          const R = T;
          if (
            R.some(($, U) => U > 0 && (C === 1 ? $ < R[U - 1] : $ > R[U - 1]))
          )
            return t("#VALUE!");
          let L = -1;
          return (
            R.forEach(($, U) => {
              (C === 1 ? $ <= N : $ >= N) && (L = U);
            }),
            L < 0 ? t("#N/A") : L + 1
          );
        }
        if (
          [
            "SUM",
            "AVERAGE",
            "MIN",
            "MAX",
            "MEDIAN",
            "LARGE",
            "SMALL",
            "STDEV.S",
            "VAR.S",
          ].includes(u)
        ) {
          const D = u === "LARGE" || u === "SMALL" ? O.slice(0, 1) : O,
            B = [];
          D.forEach((T, N) => {
            const R = !!l[N] && typeof l[N] == "object" && "ref" in l[N];
            for (const L of Gl(T))
              typeof L == "number"
                ? B.push(L)
                : !R &&
                  !Array.isArray(T) &&
                  (typeof L == "boolean" ||
                    (typeof L == "string" &&
                      L.trim() !== "" &&
                      Number.isFinite(Number(L)))) &&
                  B.push(Number(L));
          });
          const _ = Rc(B);
          if (u === "SUM") return _;
          if (u === "AVERAGE") return B.length ? _ / B.length : t("#DIV/0!");
          if (u === "MIN")
            return B.length ? B.reduce((T, N) => Math.min(T, N)) : 0;
          if (u === "MAX")
            return B.length ? B.reduce((T, N) => Math.max(T, N)) : 0;
          if (u === "STDEV.S" || u === "VAR.S") {
            if (B.length < 2) return t("#DIV/0!");
            const T =
              B.reduce((N, R) => N + (R - _ / B.length) ** 2, 0) /
              (B.length - 1);
            return u === "VAR.S" ? T : Math.sqrt(T);
          }
          if ((B.sort((T, N) => T - N), u === "MEDIAN"))
            return B.length
              ? (B[Math.floor((B.length - 1) / 2)] +
                  B[Math.floor(B.length / 2)]) /
                  2
              : t("#NUM!");
          const C = Number(O[1]);
          return !Number.isInteger(C) || C < 1 || C > B.length
            ? t("#NUM!")
            : u === "LARGE"
              ? B[B.length - C]
              : B[C - 1];
        }
        if (u === "SUBSTITUTE") {
          if (O.length < 3 || O.length > 4 || O.some(Array.isArray))
            return t("#VALUE!");
          const D = String(O[0] ?? ""),
            B = String(O[1] ?? ""),
            _ = String(O[2] ?? ""),
            C = B.length ? D.split(B).length - 1 : 0;
          if (
            O[3] !== undefined &&
            (!Number.isInteger(Number(O[3])) || Number(O[3]) < 1)
          )
            return t("#VALUE!");
          if (
            D.length +
              (O[3] === undefined ? C : Math.min(C, 1)) *
                (_.length - B.length) >
            32767
          )
            return t("#LIMIT!");
          if (!B.length) return D;
          if (O[3] === undefined) return D.split(B).join(_);
          let T = 0;
          return D.split(B).reduce(
            (N, R, L) => (L === 0 ? R : N + (++T === Number(O[3]) ? _ : B) + R),
            "",
          );
        }
        if (u === "SEARCH" || u === "SEARCHB") {
          if (O.length < 2 || O.length > 3 || O.some(Array.isArray))
            return t("#VALUE!");
          const D = String(O[0] ?? "").toLowerCase(),
            B = String(O[1] ?? "").toLowerCase(),
            _ = O[2] === undefined ? 1 : Number(O[2]);
          if (!Number.isInteger(_) || _ < 1 || _ > B.length)
            return t("#VALUE!");
          if (!/[?*~]/.test(D)) {
            const C = B.indexOf(D, _ - 1);
            return C < 0 ? t("#VALUE!") : C + 1;
          }
          if (B.length * B.length * D.length > 2e6) return t("#LIMIT!");
          for (let C = _ - 1; C < B.length; C++)
            if (Xb(B.slice(C), `${D}*`)) return C + 1;
          return t("#VALUE!");
        }
        if (
          u === "CONCATENATE" &&
          O.flatMap(Gl).reduce((D, B) => D + String(B ?? "").length, 0) > 32767
        )
          return t("#LIMIT!");
        const V = d(u, l);
        return /^IM/.test(u) &&
          !["IMABS", "IMREAL", "IMAGINARY", "IMARGUMENT"].includes(u) &&
          typeof V == "number"
          ? String(V)
          : typeof V == "string" && V.length > 32767
            ? t("#LIMIT!")
            : V;
      } catch (f) {
        return wn(f) ? t(f.error) : t("#ERROR!");
      }
    }),
    e
  );
}
function Af(e, r, t = "1900") {
  if (r != null && r.numberFormat)
    try {
      return vg(e, r.numberFormat, t).text;
    } catch {
      return "#FORMAT!";
    }
  if (e === null) return "";
  if (typeof e == "boolean") return e ? "TRUE" : "FALSE";
  if (typeof e != "number") return e;
  if (!Number.isFinite(e)) return "#NUM!";
  if (r?.format === "date") {
    if (r.numFmt)
      try {
        if (Bw(r.numFmt).kind === "temporal")
          return vg(
            e,
            { category: "custom", locale: "en-US", pattern: r.numFmt },
            t,
          ).text;
      } catch {}
    return vg(
      e,
      { category: "date", locale: "en-US", pattern: "yyyy-mm-dd" },
      t,
    ).text;
  }
  if (r != null && r.numFmt) {
    if (r.numFmt.toLowerCase() === "general")
      return String(Number(e.toPrecision(15)));
    const n = zB(e, r.numFmt);
    if (n !== undefined) return n;
  }
  return r?.format === "percent"
    ? new Intl.NumberFormat("en-US", {
        style: "percent",
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }).format(e)
    : r?.format === "currency"
      ? new Intl.NumberFormat("en-US", {
          style: "currency",
          currency: "USD",
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        }).format(e)
      : r?.format === "number"
        ? new Intl.NumberFormat("en-US", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          }).format(e)
        : String(Number(e.toPrecision(15)));
}
const EC = Object.freeze({
    mode: "automatic",
    iterative: false,
    maxIterations: 100,
    maxChange: 0.001,
    timeBudgetMs: 2500,
  }),
  Ny = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  Bie = () => {
    throw Error(
      "Invalid calculation settings. Choose Automatic or Manual, explicit iteration, 1–1000 iterations, convergence from 0 to 1, a 1–10000 ms time budget, 1–128 MiB working data and 1000–10000000 work units.",
    );
  };
function Nx(e) {
  return !Ny(e) ||
    Object.keys(e).some(
      (r) => ![...Object.keys(EC), "memoryBudgetMiB", "workBudget"].includes(r),
    ) ||
    typeof e.mode != "string" ||
    !["automatic", "manual"].includes(e.mode) ||
    typeof e.iterative != "boolean" ||
    !Number.isInteger(e.maxIterations) ||
    Number(e.maxIterations) < 1 ||
    Number(e.maxIterations) > 1e3 ||
    typeof e.maxChange != "number" ||
    !Number.isFinite(e.maxChange) ||
    e.maxChange < 0 ||
    e.maxChange > 1 ||
    !Number.isInteger(e.timeBudgetMs) ||
    Number(e.timeBudgetMs) < 1 ||
    Number(e.timeBudgetMs) > 1e4 ||
    (e.memoryBudgetMiB !== undefined &&
      (!Number.isInteger(e.memoryBudgetMiB) ||
        Number(e.memoryBudgetMiB) < 1 ||
        Number(e.memoryBudgetMiB) > 128)) ||
    (e.workBudget !== undefined &&
      (!Number.isInteger(e.workBudget) ||
        Number(e.workBudget) < 1e3 ||
        Number(e.workBudget) > 1e7))
    ? Bie()
    : {
        mode: e.mode,
        iterative: e.iterative,
        maxIterations: Number(e.maxIterations),
        maxChange: e.maxChange,
        timeBudgetMs: Number(e.timeBudgetMs),
        ...(e.memoryBudgetMiB !== undefined
          ? { memoryBudgetMiB: Number(e.memoryBudgetMiB) }
          : {}),
        ...(e.workBudget !== undefined
          ? { workBudget: Number(e.workBudget) }
          : {}),
      };
}
function Lie(e, r) {
  e.calculation !== undefined && (r.calculation = Nx(e.calculation));
  const t = e.importedCalculationCaches;
  if (t === undefined) return;
  const n = () => {
    throw Error(
      "Invalid imported formula-cache evidence. Caches must remain separately labeled source evidence.",
    );
  };
  if (
    !Ny(t) ||
    t.schemaVersion !== 1 ||
    typeof t.fileName != "string" ||
    !t.fileName ||
    t.fileName.length > 255 ||
    typeof t.sourceSha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(t.sourceSha256) ||
    typeof t.importedAt != "string" ||
    !Number.isFinite(Date.parse(t.importedAt)) ||
    !Ny(t.sourceSettings) ||
    !Array.isArray(t.entries) ||
    t.entries.length > 1e5
  )
    return n();
  const o = t.sourceSettings;
  if (
    typeof o.mode != "string" ||
    o.mode.length > 40 ||
    typeof o.iterative != "boolean" ||
    (o.maxIterations !== null &&
      (!Number.isSafeInteger(o.maxIterations) ||
        Number(o.maxIterations) < 1)) ||
    (o.fullPrecision !== undefined && typeof o.fullPrecision != "boolean") ||
    (o.maxChange !== null &&
      (typeof o.maxChange != "number" ||
        !Number.isFinite(o.maxChange) ||
        o.maxChange < 0))
  )
    return n();
  const a = (u) => {
      const l = /^([A-Z]+)([1-9]\d*)$/.exec(u);
      return (
        !!l &&
        Number(l[2]) <= 1048576 &&
        [...l[1]].reduce((c, h) => c * 26 + h.charCodeAt(0) - 64, 0) <= 16384
      );
    },
    s = new Set(),
    d = t.entries.map((u) => {
      if (
        !Ny(u) ||
        typeof u.sheetId != "string" ||
        !u.sheetId ||
        u.sheetId.length > 100 ||
        typeof u.address != "string" ||
        !/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(u.address) ||
        !a(u.address) ||
        typeof u.formula != "string" ||
        !u.formula.startsWith("=") ||
        u.formula.length > 8192
      )
        return n();
      const l = JSON.stringify([u.sheetId, u.address]);
      if (s.has(l)) return n();
      s.add(l);
      let c;
      if (u.cache !== undefined) {
        const h = u.cache;
        if (
          !Ny(h) ||
          !["error", "blank", "number", "text", "boolean"].includes(
            String(h.type),
          ) ||
          (h.type === "blank" && h.value !== null) ||
          (h.type === "number" &&
            (typeof h.value != "number" || !Number.isFinite(h.value))) ||
          (h.type === "boolean" && typeof h.value != "boolean") ||
          (["text", "error"].includes(String(h.type)) &&
            (typeof h.value != "string" || h.value.length > 32767)) ||
          (h.type === "error" && !/^#[A-Z0-9_/?!]+$/.test(String(h.value)))
        )
          return n();
        c = { type: h.type, value: h.value };
      }
      return {
        sheetId: u.sheetId,
        address: u.address,
        formula: u.formula,
        ...(c ? { cache: c } : {}),
      };
    });
  r.importedCalculationCaches = {
    schemaVersion: 1,
    fileName: t.fileName,
    sourceSha256: t.sourceSha256,
    importedAt: t.importedAt,
    sourceSettings: {
      mode: o.mode,
      iterative: o.iterative,
      maxIterations: o.maxIterations,
      maxChange: o.maxChange,
      ...(o.fullPrecision !== undefined
        ? { fullPrecision: o.fullPrecision }
        : {}),
    },
    entries: d,
  };
}
const tu = new WeakMap(),
  ch = new WeakMap(),
  j5 = 2e5,
  ng = { type: "error", value: "#PENDING!" },
  F5 = (e) => ({ type: "error", value: e }),
  Pie = new RegExp(
    `\\b(?:${[...y0, "LET", "LAMBDA", "INDIRECT", "OFFSET", "ISFORMULA", "TEXTBEFORE", "TEXTAFTER"].join("|")})\\s*\\(`,
    "i",
  ),
  _O = { test: (e) => Rg(e).length > 0 },
  ik = (e, r) => {
    var t;
    return (
      !!((t = e.definedNames) != null && t.length) ||
      [...r.values()].some((n) => {
        const o = n.formula.replace(/"(?:[^"]|"")*"/g, '""');
        return (
          Pie.test(o) ||
          /[{}#]/.test(o) ||
          /^=\s*[^()+*/]*:/.test(o) ||
          (/\b(?:IF|IFERROR|IFNA)\s*\(/i.test(o) && o.includes(":"))
        );
      })
    );
  },
  DO = new WeakMap();
const ok = new WeakMap();
function Uie(e) {
  const r = es(e.cells) ? ok.get(e.cells)?.get(e.id) : undefined;
  if (r) return r;
  const t = [];
  let n, o;
  const a = () =>
      n ??
      (n = Object.entries(e.cells).map(([l, c]) => [
        `${e.id}:${l}`,
        JSON.stringify([c.raw, !!c.literal]),
      ])),
    s = () => o ?? (o = JSON.stringify(a()));
  es(e.cells) || s();
  for (const l in e.cells) {
    const c = e.cells[l],
      h = `${e.id}:${l}`;
    !c.literal &&
      typeof c.raw == "string" &&
      c.raw.startsWith("=") &&
      t.push([h, { sheetId: e.id, address: l, formula: c.raw }]);
  }
  const d = {
    inputs: a,
    formulas: t,
    signature: s,
    volatile: [...new Set(t.flatMap(([, l]) => Rg(l.formula)))],
  };
  if (es(e.cells)) {
    const l = ok.get(e.cells) ?? new Map();
    (l.set(e.id, d), ok.set(e.cells, l));
  }
  return d;
}
function B5(e) {
  const r = Nx(e.calculation ?? EC);
  return JSON.stringify([
    e.id,
    r,
    ts(e.datePolicy),
    e.definedNames,
    e.workbookLinks,
    e.extensions,
    g1(e),
    e.sheets.map((n) => [
      n.id,
      n.name,
      n.rowCount,
      n.colCount,
      n.hiddenRows ?? [],
      n.hiddenCols ?? [],
      !!n.hidden,
      !!n.veryHidden,
      n.tables ?? [],
      n.merges ?? [],
      n.outlineGroups ?? [],
    ]),
  ]);
}
function OO(e) {
  const r = e.sheets.map(Uie);
  let t;
  const n = () => t ?? (t = new Map(r.flatMap((h) => h.inputs()))),
    o = new Map(r.flatMap((h) => h.formulas)),
    a = Nx(e.calculation ?? EC),
    s = B5(e);
  let d, u;
  const l = () =>
      d ??
      (d = `[${JSON.stringify(e.revision)},${JSON.stringify(s)},[${r.map((h) => h.signature()).join(",")}]]`),
    c = () =>
      u ??
      (u = `[${l()},${JSON.stringify(e.importedCalculationCaches ?? null)}]`);
  return (
    es(e) || c(),
    {
      inputs: n,
      formulas: o,
      settings: a,
      metadata: s,
      volatile: [
        ...new Set(
          r
            .flatMap((h) => h.volatile)
            .concat((e.definedNames ?? []).flatMap((h) => Rg(h.refersTo))),
        ),
      ].sort(),
      inputSignature: l,
      signature: c,
    }
  );
}
function W1(e) {
  if (!es(e)) return OO(e);
  let r = DO.get(e);
  return (r || ((r = OO(e)), DO.set(e, r)), r);
}
function nm(e, r, t = W1(r)) {
  return !!e && (es(r) || e.signature() === t.signature());
}
function MO(e, r) {
  if (
    e.revision !== r.revision ||
    B5(e) !== B5(r) ||
    (e.importedCalculationCaches !== r.importedCalculationCaches &&
      JSON.stringify(e.importedCalculationCaches) !==
        JSON.stringify(r.importedCalculationCaches))
  )
    return false;
  for (let t = 0; t < e.sheets.length; t++) {
    const n = e.sheets[t].cells,
      o = r.sheets[t].cells;
    if (n === o) continue;
    const a = Object.keys(n);
    if (a.length !== Object.keys(o).length) return false;
    for (const s of a) {
      const d = n[s],
        u = o[s];
      if (!u || d.raw !== u.raw || !!d.literal != !!u.literal) return false;
    }
  }
  return true;
}
function $ie(e) {
  return tu.has(e) || ch.has(e);
}
function SR(e, r) {
  if (e === r || e.id !== r.id || r.revision <= e.revision) return;
  const t = tu.get(e);
  t && nm(t, e) && ch.set(r, t);
}
function zie(e, r) {
  if (e === r) return;
  const t = tu.get(e);
  t && nm(t, e) && MO(e, r)
    ? tu.set(r, t)
    : ch.has(e) && MO(e, r) && ch.set(r, ch.get(e));
}
function Vie(e, r) {
  return { ...e, calculation: Nx(r) };
}
function A1(e) {
  !e ||
    typeof e != "object" ||
    Object.isFrozen(e) ||
    (Object.values(e).forEach(A1), Object.freeze(e));
}
function Hie(e) {
  const r = tu.get(e);
  return nm(r, e) ? { calculator: r.calculator, report: r.report } : undefined;
}
const PE = new WeakMap();
function qie(e) {
  if (!e) return new Map();
  const r = PE.get(e.report);
  if (r) return r;
  const t = new Map();
  for (const n of e.report.cells) {
    const o = `${n.sheetId}:${n.address}`,
      a = e.values.get(o);
    a && a.value !== "#PENDING!"
      ? t.set(o, { revision: e.report.revision, result: { ...a } })
      : n.lastKnown && t.set(o, n.lastKnown);
  }
  return t;
}
function lU(e) {
  return W1(e).inputSignature();
}
const Cc = Object.freeze({
    rows: 1e4,
    cells: 1e5,
    columns: 64,
    conditions: 128,
    values: 1e3,
    text: 1024,
    comparisonWork: 5e6,
  }),
  Gie =
    "View only. Column rules intersect (AND); advanced groups use OR of AND rules. Counts include blank records in the declared range. Manual/outline hiding is separate. Full scope includes every data row; visible scope excludes filtered and manually hidden rows. Applying a view leaves source addresses, formulas, identities and history unchanged. Native JSON, XLSX, print and AI selections use full source data. Active masks are session state; named criteria can be saved in native JSON and require a fresh preview before reuse.",
  P5 = (e) => e === null || e === "",
  Zw = (e, r = false) =>
    r ? e.normalize("NFC") : e.normalize("NFC").toLocaleLowerCase("en-US"),
  eh = (e, r = false) =>
    P5(e) ? "blank" : `${typeof e}:${typeof e == "string" ? Zw(e, r) : e}`,
  Zie = [
    "equals",
    "notEquals",
    "contains",
    "notContains",
    "startsWith",
    "endsWith",
    "wildcard",
    "gt",
    "gte",
    "lt",
    "lte",
    "blank",
    "notBlank",
  ],
  jO = (e) =>
    e === null ||
    typeof e == "boolean" ||
    (typeof e == "number" && Number.isFinite(e)) ||
    (typeof e == "string" && e.length <= 32767);
function U5(e, r, t) {
  if (r === "number") return typeof e == "number" ? e : null;
  if (r === "date") {
    const n = ts(t.datePolicy),
      o =
        typeof e == "number" ? e : typeof e == "string" ? zp(e, n) : undefined;
    return o !== undefined && o >= 0 && o < wd(n.dateSystem) + 1 ? o : null;
  }
  return typeof e == "string" ? e : null;
}
function Yie(e, r) {
  const t = [];
  for (let d = 0; d < r.length; d++) {
    const u = r[d];
    u === "~" && d + 1 < r.length
      ? t.push({ text: r[++d] })
      : t.push(
          u === "*" ? { star: true } : u === "?" ? { any: true } : { text: u },
        );
  }
  let n = 0,
    o = 0,
    a = -1,
    s = 0;
  for (; n < e.length;)
    if (o < t.length && (t[o].any || t[o].text === e[n])) (n++, o++);
    else if (o < t.length && t[o].star) ((a = o++), (s = n));
    else if (a >= 0) ((o = a + 1), (n = ++s));
    else return false;
  for (; o < t.length && t[o].star;) o++;
  return o === t.length;
}
function Jie(e, r, t, n) {
  if (r.operator === "blank") return P5(e);
  if (r.operator === "notBlank") return !P5(e);
  if (r.type === "value") {
    const s = eh(e, n) === eh(r.value, n);
    return r.operator === "equals" ? s : !s;
  }
  let o = U5(e, r.type, t),
    a = U5(r.value, r.type, t);
  if (o === null || a === null) return false;
  switch (
    (typeof o == "string" &&
      typeof a == "string" &&
      ((o = Zw(o, n)), (a = Zw(a, n))),
    r.operator)
  ) {
    case "equals":
      return o === a;
    case "notEquals":
      return o !== a;
    case "gt":
      return o > a;
    case "gte":
      return o >= a;
    case "lt":
      return o < a;
    case "lte":
      return o <= a;
    case "contains":
      return String(o).includes(String(a));
    case "notContains":
      return !String(o).includes(String(a));
    case "startsWith":
      return String(o).startsWith(String(a));
    case "endsWith":
      return String(o).endsWith(String(a));
    case "wildcard":
      return Yie(String(o), String(a));
  }
  return false;
}
function $5(e, r, t, n) {
  const o = e(r, t),
    a = n?.sheets.find((d) => d.id === r)?.cells[t];
  if (
    a &&
    a.raw === o &&
    (a.literal || typeof a.raw != "string" || !a.raw.startsWith("="))
  )
    return o;
  if (
    typeof o == "string" &&
    /^#(?:PENDING|LIMIT|TIMEOUT|CANCELLED|BUSY|BLOCKED|CALC_LIMIT)!?$/i.test(o)
  )
    throw Error(
      `Calculation is incomplete at ${t}. Recalculate before filtering; unchecked rows are never reported as an empty population.`,
    );
  return o;
}
const z5 = Object.freeze({ views: 20, name: 80, characters: 512e3 });
function xv(e, r) {
  if (
    !e ||
    typeof e != "object" ||
    Array.isArray(e) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(e)) ||
    Object.keys(e).some((t) => !r.includes(t))
  )
    throw Error("Invalid saved filter view metadata.");
}
function Ry(e, r) {
  return (
    typeof e == "string" &&
    !!e.trim() &&
    e.length <= r &&
    !/[\u0000-\u001f\u007f]/.test(e)
  );
}
function BO(e) {
  return e.normalize("NFC").trim().toLocaleLowerCase("en-US");
}
function NR(e) {
  if (!Ry(e, z5.name))
    throw Error("Use a nonempty filter view name of up to 80 characters.");
  return e.normalize("NFC").trim();
}
function RR(e, r) {
  var t, o;
  if (
    ((t = e.structureProtection) != null && t.enabled) ||
    ((o = e.sheets.find((a) => a.id === r)?.protection) != null && o.enabled)
  )
    throw Error(
      "Unprotect the workbook and source worksheet before changing saved filter views.",
    );
}
function _R(e, r) {
  const t = e.savedFilterViews?.find((o) => o.id === r);
  if (!t) throw Error("This saved filter view no longer exists.");
  return t;
}
const CC = "ledgerglass-consolidation-v1",
  k1 = Object.freeze({
    sources: 8,
    categories: 3,
    contributions: 2e3,
    outputRows: 4e3,
    outputCells: 4e4,
    groups: 1e3,
    metadataCharacters: 8e6,
  }),
  Hu = string()
    .min(1)
    .max(128)
    .refine((e) => !!e.trim() && !/[\u0000-\u001f\u007f]/.test(e)),
  gd = string().min(1).max(100),
  V5 = string().regex(/^[a-f0-9]{64}$/),
  hU = union([string().min(1).max(32767), number().finite(), boolean()]),
  ak = discriminatedUnion("kind", [
    strictObject({ kind: literal("constant"), value: Hu }),
    strictObject({ kind: literal("field"), field: Hu }),
  ]),
  roe = strictObject({
    id: gd,
    workbookId: gd,
    sheetId: gd,
    range: string().min(1).max(50),
    tableId: gd.optional(),
    categories: array(Hu).min(1).max(3),
    amountField: Hu,
    period: ak,
    currency: ak,
    unit: ak,
    duplicatePolicy: _enum(["reject", "sum"]),
  }),
  pU = strictObject({
    engine: literal(CC),
    sources: array(roe).min(1).max(8),
    categoryLabels: array(Hu).min(1).max(3),
    period: Hu,
    currency: Hu,
    unit: Hu,
    linkedAmounts: boolean(),
    collapseDetails: boolean(),
  }).superRefine((e, r) => {
    const t = [
      "Row type",
      ...e.categoryLabels,
      "Period",
      "Currency",
      "Unit",
      "Amount",
      "Source",
      "Link state",
    ];
    (new Set(t.map((n) => n.toLowerCase())).size !== t.length ||
      new Set(e.sources.map((n) => n.id)).size !== e.sources.length ||
      e.sources.some(
        (n) =>
          n.categories.length !== e.categoryLabels.length ||
          new Set(n.categories).size !== n.categories.length,
      )) &&
      r.addIssue({
        code: "custom",
        message:
          "Use unique source IDs, distinct category fields, and distinct output headers with matching category counts.",
      });
  }),
  noe = strictObject({
    sourceId: gd,
    workbookId: gd,
    workbookName: string().min(1).max(200),
    sheetId: gd,
    sheetName: string().min(1).max(31),
    row: number().int().min(2).max(1048576),
    address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
    locator: string().max(300),
    categories: array(hU).min(1).max(3),
    amount: number().finite(),
    period: Hu,
    currency: Hu,
    unit: Hu,
    link: _enum(["live", "snapshot", "approved snapshot"]),
  }),
  ioe = strictObject({
    sourceId: gd,
    row: number().int().positive().optional(),
    code: string().max(40),
    severity: _enum(["error", "warning"]),
    message: string().max(4e3),
  }),
  ooe = strictObject({
    sourceId: gd,
    workbookId: gd,
    workbookName: string().max(200),
    sheetId: gd,
    sheetName: string().max(31),
    range: string().max(50),
    fingerprint: V5,
    records: number().int().min(0).max(500),
    acceptedRecords: number().int().min(0).max(500),
    amount: number().finite(),
    periods: array(Hu).max(500),
    currencies: array(Hu).max(500),
    units: array(Hu).max(500),
  }),
  aoe = strictObject({
    categories: array(hU).min(1).max(3),
    level: number().int().min(1).max(3),
    count: number().int().min(1).max(2e3),
    amount: number().finite(),
  }),
  soe = strictObject({
    sources: array(ooe).max(8),
    contributions: array(noe).max(2e3),
    subtotals: array(aoe).max(1e3),
    exceptions: array(ioe).max(8e3),
    grandTotal: number().finite().nullable(),
    sourceTotal: number().finite().nullable(),
    difference: number().finite().nullable(),
    complete: boolean(),
  }),
  loe = strictObject({
    engine: literal(CC),
    definition: pU,
    review: soe,
    origin: strictObject({
      workbookId: gd,
      revision: number().int().nonnegative(),
      at: string().datetime(),
    }),
    evidenceHash: V5,
    outputRange: string().max(50),
    outputHash: V5,
  });
function mU(e) {
  const r = pU.safeParse(e);
  if (!r.success)
    throw Error(`Invalid consolidation: ${r.error.issues[0]?.message}`);
  return r.data;
}
const DR = (e) => JSON.stringify(e.map((r) => [typeof r, r]));
function Op(e) {
  let r = 0,
    t = 0;
  for (const n of e) {
    const o = n - t,
      a = r + o;
    if (((t = a - r - o), (r = a), !Number.isFinite(r) || !Number.isFinite(t)))
      throw Error("Consolidation total exceeds finite numeric limits.");
  }
  return Object.is(r, -0) ? 0 : r;
}
function gU(e) {
  const r = new Map();
  for (const t of e)
    for (let n = 1; n <= t.categories.length; n++) {
      const o = DR(t.categories.slice(0, n)),
        a = r.get(o) ?? [];
      (a.push(t), r.set(o, a));
    }
  if (r.size > k1.groups)
    throw Error(
      "Consolidation exceeds 1,000 subtotal groups. Narrow the sources; no rows were truncated.",
    );
  return [...r].map(([t, n]) => {
    const o = JSON.parse(t).length;
    return {
      categories: n[0].categories.slice(0, o),
      level: o,
      count: n.length,
      amount: Op(n.map((a) => a.amount)),
    };
  });
}
function uoe(e, r) {
  let t = 0;
  for (const s of r.sheets) {
    const d = e.sheets?.find((f) => f.id === s.id)?.consolidation;
    if (d === undefined) continue;
    if (((t += JSON.stringify(d).length), t > k1.metadataCharacters))
      throw Error(
        "Consolidation evidence exceeds 8,000,000 characters. Retain native backups before removing obsolete reports.",
      );
    const u = loe.safeParse(d);
    if (!u.success)
      throw Error(
        `Invalid consolidation evidence: ${u.error.issues[0]?.message}`,
      );
    const l = u.data,
      c = l.review,
      h = l.definition.sources.map((f) => f.id);
    if (
      !c.complete ||
      c.exceptions.some((f) => f.severity === "error") ||
      !c.contributions.length ||
      c.sources.length !== h.length ||
      new Set(c.sources.map((f) => f.sourceId)).size !== h.length ||
      c.sources.some((f) => !h.includes(f.sourceId)) ||
      c.contributions.some(
        (f) =>
          !h.includes(f.sourceId) ||
          f.categories.length !== l.definition.categoryLabels.length ||
          f.period !== l.definition.period ||
          f.currency !== l.definition.currency ||
          f.unit !== l.definition.unit,
      ) ||
      c.grandTotal !== Op(c.contributions.map((f) => f.amount)) ||
      c.sourceTotal !== Op(c.sources.map((f) => f.amount)) ||
      c.difference !== c.grandTotal - c.sourceTotal ||
      hr(c.subtotals) !== hr(gU(c.contributions)) ||
      l.evidenceHash !==
        hr({ definition: l.definition, review: c, origin: l.origin })
    )
      throw Error(
        "Consolidation evidence does not reconcile to its sources, categories or retained fingerprint.",
      );
    for (const f of c.sources) {
      const p = c.contributions.filter((m) => m.sourceId === f.sourceId);
      if (
        f.records !== p.length ||
        f.acceptedRecords !== p.length ||
        f.amount !== Op(p.map((m) => m.amount)) ||
        p.some((m) => m.workbookId !== f.workbookId || m.sheetId !== f.sheetId)
      )
        throw Error("Consolidation source evidence does not reconcile.");
    }
    s.consolidation = l;
  }
}
const OR = "ledgerglass-sources-v1",
  coe = ["database", "odbc", "odata", "api", "erp", "file-folder"],
  fl = Object.freeze({
    sources: 20,
    columns: 50,
    rows: 5e3,
    cells: 1e5,
    responseBytes: 8 * 1024 * 1024,
    timeoutMs: 3e4,
    metadataCharacters: 2e5,
  }),
  UE = string().regex(/^[A-Za-z0-9][A-Za-z0-9._-]{0,99}$/),
  $E = string()
    .min(1)
    .max(128)
    .refine((e) => e === e.trim() && !/[\u0000-\u001f\u007f]/.test(e)),
  Jy = string().datetime({ offset: true }),
  LO = string().regex(/^[a-f0-9]{64}$/),
  doe = strictObject({
    name: $E,
    type: _enum(["text", "number", "boolean", "date"]),
    nullable: boolean(),
  }),
  MR = strictObject({
    id: UE,
    label: $E,
    kind: _enum(coe),
    schemaVersion: UE,
    columns: array(doe).min(1).max(fl.columns),
    controlFields: array($E).max(fl.columns),
  }).superRefine((e, r) => {
    (new Set(e.columns.map((t) => t.name.toLowerCase())).size !==
      e.columns.length ||
      new Set(e.controlFields).size !== e.controlFields.length ||
      e.controlFields.some(
        (t) => !e.columns.some((n) => n.name === t && n.type === "number"),
      )) &&
      r.addIssue({
        code: "custom",
        message: "Use unique columns and declared numeric control fields.",
      });
  }),
  foe = strictObject({ source: MR, expiresAt: Jy }),
  hoe = union([string().max(32767), number().finite(), boolean(), _null()]),
  poe = strictObject({
    protocol: literal(OR),
    source: MR,
    sourceTimestamp: Jy,
    version: UE,
    rows: array(array(hoe).max(fl.columns)).max(fl.rows),
    complete: literal(true),
  }),
  vU = [
    "unavailable",
    "refused",
    "expired",
    "schema",
    "limit",
    "timeout",
    "cancelled",
    "invalid",
    "stale",
  ],
  moe = {
    unavailable:
      "The source is unavailable. The last successful data is retained. Check the host connection, then retry.",
    refused:
      "This source is not authorized by the host. No source data was replaced.",
    expired:
      "Source access expired. Renew access in the host application and preview again.",
    schema:
      "The source schema changed. The working population is retained. Have the host approve the new schema and import it into a new worksheet.",
    limit:
      "The source exceeds the published row, column, cell or response limit. No rows were truncated or published.",
    timeout:
      "Source refresh timed out. The last successful data is retained; retry when the host is available.",
    cancelled:
      "Source refresh cancelled. The last successful data is retained.",
    invalid:
      "The source returned an invalid or incomplete result. The last successful data is retained.",
    stale:
      "The workbook, preview or source approval changed. Preview again before applying a refresh.",
  };
class Rn extends Error {
  constructor(r) {
    (super(moe[r]), (this.code = r), (this.name = "ExternalSourceError"));
  }
}
function Mp(e) {
  return e instanceof Rn
    ? e
    : e instanceof Error &&
        e.name === "ExternalSourceError" &&
        vU.includes(e.code)
      ? new Rn(e.code)
      : new Rn("unavailable");
}
function h0(e, r = Date.now()) {
  const t = foe.safeParse(e);
  if (!t.success) throw new Rn("invalid");
  if (Date.parse(t.data.expiresAt) <= r) throw new Rn("expired");
  return t.data;
}
function AC(e, r) {
  return hr(e) === hr(r);
}
function bU(e, r) {
  if (e && typeof e == "object") {
    const o = e.rows;
    if (
      Array.isArray(o) &&
      (o.length > fl.rows ||
        o.some((a) => Array.isArray(a) && a.length > fl.columns))
    )
      throw new Rn("limit");
  }
  const t = poe.safeParse(e);
  if (!t.success) throw new Rn("invalid");
  const n = t.data;
  if (r && !AC(n.source, r)) throw new Rn("schema");
  if (
    (n.rows.length + 1) * n.source.columns.length > fl.cells ||
    new TextEncoder().encode(JSON.stringify(n)).length > fl.responseBytes
  )
    throw new Rn("limit");
  for (const o of n.rows) {
    if (o.length !== n.source.columns.length) throw new Rn("schema");
    o.forEach((a, s) => {
      const d = n.source.columns[s];
      if (
        a === null
          ? !d.nullable
          : d.type === "date"
            ? typeof a != "string" ||
              !/^\d{4}-\d{2}-\d{2}$/.test(a) ||
              !Number.isFinite(Date.parse(a)) ||
              new Date(a).toISOString().slice(0, 10) !== a
            : typeof a !==
              { text: "string", number: "number", boolean: "boolean" }[d.type]
      )
        throw new Rn("schema");
    });
  }
  return n;
}
function yU(e) {
  return e.source.controlFields.map((r) => {
    const t = e.source.columns.findIndex((d) => d.name === r);
    let n = 0,
      o = 0,
      a = 0,
      s = 0;
    for (const d of e.rows) {
      if (d[t] === null) {
        s++;
        continue;
      }
      const u = d[t] - o,
        l = n + u;
      if (
        ((o = l - n - u),
        (n = l),
        a++,
        !Number.isFinite(n) || !Number.isFinite(o))
      )
        throw new Rn("limit");
    }
    return { field: r, sum: Object.is(n, -0) ? 0 : n, numeric: a, blanks: s };
  });
}
const goe = strictObject({
    field: $E,
    sum: number().finite(),
    numeric: number().int().min(0).max(5e3),
    blanks: number().int().min(0).max(5e3),
  }),
  voe = strictObject({
    protocol: literal(OR),
    source: MR,
    lastSuccess: strictObject({
      at: Jy,
      sourceTimestamp: Jy,
      version: UE,
      rows: number().int().min(0).max(5e3),
      controls: array(goe).max(50),
      dataHash: LO,
    }),
    lastAttempt: strictObject({
      at: Jy,
      outcome: _enum(["succeeded", "failed", "cancelled"]),
      code: _enum(vU).optional(),
    }),
    outputHash: LO,
  });
function boe(e, r) {
  let t = 0,
    n = 0;
  for (const a of r.sheets) {
    const s = e.sheets.find((l) => l.id === a.id)?.externalConnection;
    if (s === undefined) continue;
    if (
      ++n > fl.sources ||
      (t += JSON.stringify(s).length) > fl.metadataCharacters
    )
      throw new Rn("limit");
    const d = voe.safeParse(s);
    if (!d.success) throw new Rn("invalid");
    const u = d.data;
    if (
      u.lastSuccess.controls.length !== u.source.controlFields.length ||
      u.lastSuccess.controls.some(
        (l, c) =>
          l.field !== u.source.controlFields[c] ||
          l.numeric + l.blanks !== u.lastSuccess.rows,
      ) ||
      (u.lastSuccess.rows + 1) * u.source.columns.length > fl.cells ||
      Date.parse(u.lastAttempt.at) < Date.parse(u.lastSuccess.at) ||
      (u.lastAttempt.outcome === "succeeded") == !!u.lastAttempt.code
    )
      throw new Rn("invalid");
    a.externalConnection = u;
  }
}

export {
  $5,
  $ie,
  A1,
  AC,
  Af,
  BO,
  CC,
  Cc,
  DR,
  EC,
  F5,
  Fie,
  Gie,
  Hie,
  Jie,
  Lie,
  Mp,
  NR,
  Nx,
  OR,
  Op,
  PE,
  RR,
  Rn,
  Ry,
  SR,
  Ty,
  U5,
  Vie,
  W1,
  Zie,
  Zw,
  _O,
  _R,
  bU,
  boe,
  ch,
  eh,
  fl,
  gU,
  h0,
  ik,
  j5,
  jO,
  k1,
  lU,
  mU,
  ng,
  nm,
  q6,
  qie,
  tu,
  uoe,
  xv,
  yU,
  z5,
  zie,
};
