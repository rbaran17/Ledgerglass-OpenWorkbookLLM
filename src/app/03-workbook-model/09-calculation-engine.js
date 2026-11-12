import {
  LJ,
  O4,
  SN,
  lx,
  z3,
  pd,
  yE,
  Rg,
  eB,
  nB,
} from "../01-core-schemas-and-hashing.js";
import { Ja, Mt } from "../02-financial-functions.js";
import {
  string,
  number,
  boolean,
  null as _null,
  array,
  literal,
  union,
  enum as _enum,
  unknown,
  object,
} from "zod";
import { FB, a1, jN, jw, n5, r5 } from "./01-cell-styles-and-themes.js";
import { JN, y0 } from "./02-number-formats.js";
import { pne, y5 } from "./03-defined-names.js";
import { Fc, M5, kR, pf, rU, wn } from "./04-pivot-tables-and-pictures.js";
import {
  $ie,
  A1,
  F5,
  Fie,
  PE,
  W1,
  _O,
  ch,
  ik,
  j5,
  lU,
  ng,
  nm,
  qie,
  tu,
} from "./05-consolidation-and-sources.js";
import { AU, Ap, CU, Tc, ut, xd, zE } from "./06-data-model-and-analysis.js";
import { DB, Die, FN, Fw, _t, ct } from "./07-conditional-formatting.js";
import { CR, z1 } from "./08-chart-model.js";

function Cf(e, r = {}, t, n) {
  var K;
  const o = z3(e),
    a = new Set([...Fc, ...o.map((te) => te.name)]),
    s = n?.budget ?? SN(e),
    d = n?.dateContext ?? lx(e.datePolicy),
    u = new Set(),
    l = new Map(),
    c = new Map(),
    h = new Map(),
    f = (te) => (
      [
        "#LIMIT!",
        "#BLOCKED!",
        "#CYCLE!",
        "#ERROR!",
        "#TIMEOUT!",
        "#ITER!",
      ].includes(te) && u.add(te),
      kR(te)
    ),
    p = Die(e, r),
    m = new Map(),
    v = new Map(),
    g = new Map(),
    b = new Map(),
    y = new Set(),
    w = new Set(),
    E = new Map(),
    A = [],
    k = [];
  let F = 0;
  const I = [];
  let O = 0;
  const j = e.definedNames?.some((te) =>
      /\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(te.refersTo),
    ),
    z =
      (K = e.definedNames) != null && K.length
        ? new RegExp(
            `(?:^|[^A-Za-z0-9_.\\\\])(?:${e.definedNames.map((te) => te.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")})(?![A-Za-z0-9_.\\\\])`,
            "i",
          )
        : undefined,
    V = (te) =>
      /\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(te) ||
      !!(j && z != null && z.test(te)),
    D = rU(e),
    B = [...e.sheets, ...D.sheets],
    _ = new Map(B.map((te) => [te.id, te])),
    C = new Map(B.map((te) => [te.name.toLowerCase(), te]));
  let T = 0,
    N = 0;
  const R = new Map(),
    L = new RegExp(
      `\\b(?:${[...y0, "ISFORMULA", "TEXTBEFORE", "TEXTAFTER", ...o.filter((te) => te.fn.result.type === "matrix").map((te) => te.name.replaceAll(".", "\\."))].join("|")})\\s*\\(`,
      "i",
    );
  for (const te of e.sheets)
    R.set(
      te.id,
      Object.entries(te.cells)
        .filter(([, ae]) => {
          return (
            typeof ae.raw == "string" &&
            ae.raw.startsWith("=") &&
            !ae.literal &&
            (V(ae.raw) ||
              L.test(ae.raw) ||
              (/\b(?:IF|IFERROR|IFNA)\s*\(/i.test(ae.raw) &&
                ae.raw.replace(/"(?:[^"]|"")*"/g, '""').includes(":")) ||
              /[{}#]/.test(ae.raw.replace(/"(?:[^"]|"")*"/g, '""')) ||
              /^=\s*[^()+*/]*:/.test(ae.raw) ||
              (e.definedNames?.length &&
                /^=\s*[A-Za-z_][A-Za-z0-9_.]*\s*$/.test(ae.raw)))
          );
        })
        .map(([ae]) => ae)
        .sort((ae, oe) => {
          const de = _t(ae),
            he = _t(oe);
          return de.row - he.row || de.col - he.col;
        }),
    );
  const $ = (te, ae, oe, de = true) => {
      const he = (se, we = true) => {
        const Ee = CR(se, e, ae, oe, we);
        if (Ee) throw (l.set(`${ae}:${oe}`, Ee), f(Ee.error));
      };
      ((te = D.rewrite(te)), he(te, false));
      let Se = de ? pf(te, e, ae, [], t?.allowName) : te;
      ((Se = D.rewrite(Se)), he(Se));
      let le = "",
        ce = 0;
      for (; ce < Se.length;) {
        const se =
          /^(?:'((?:[^']|'')*)'|([A-Za-z_][A-Za-z0-9_.]*))!\s*(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#/i.exec(
            Se.slice(ce),
          );
        if (se) {
          const ye = C.get((se[1]?.replace(/''/g, "'") ?? se[2]).toLowerCase());
          if (!ye) throw f("#REF!");
          const Te = _t(se[3]),
            Pe = `${ut(Te.col)}${Te.row + 1}`,
            qe = Q(ye.id, Pe),
            ht = g.get(`${ye.id}:${Pe}`);
          if (!ht) throw wn(qe) ? qe : f("#REF!");
          ((le += `'${ye.name.replace(/'/g, "''")}'!${ht}`),
            (ce += se[0].length));
          continue;
        }
        if (Se[ce] === '"' || Se[ce] === "'") {
          const ye = Se[ce],
            Te = ce++;
          for (; ce < Se.length;)
            if (Se[ce++] === ye)
              if (Se[ce] === ye) ce++;
              else break;
          le += Se.slice(Te, ce);
          continue;
        }
        const we =
          /^(SEQUENCE|FILTER|SORT|SORTBY|UNIQUE|TRANSPOSE|TAKE|DROP|TEXTSPLIT)\s*\(\s*,/i.exec(
            Se.slice(ce),
          );
        if (we) {
          ((le += `${we[1]}(${we[1].toUpperCase() === "SEQUENCE" ? "1" : "#VALUE!"},`),
            (ce += we[0].length));
          continue;
        }
        const Ee = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(Se.slice(ce));
        if (Ee) {
          const ye = ce;
          ce += Ee[0].length;
          let Te = 1;
          for (; ce < Se.length && Te;)
            (Se[ce] === "[" && Te++, Se[ce] === "]" && Te--, ce++);
          const Pe = !Te && z1(e, ae, Se.slice(ye, ce), oe);
          if (!Pe) throw f("#REF!");
          const qe = _.get(Pe.sheetId);
          le += `'${qe.name.replace(/'/g, "''")}'!${Pe.range}`;
          continue;
        }
        const Ae = /^(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#/.exec(Se.slice(ce));
        if (Ae) {
          const ye = _t(Ae[1]),
            Te = `${ut(ye.col)}${ye.row + 1}`,
            Pe = Q(ae, Te),
            qe = g.get(`${ae}:${Te}`);
          if (!qe) throw wn(Pe) ? Pe : f("#REF!");
          ((le += qe), (ce += Ae[0].length));
          continue;
        }
        le += Se[ce++];
      }
      return le.replace(/\b_xlfn\.(?:_xlws\.)?/gi, "");
    },
    U = (te) => {
      const ae = C.get((te ?? "").toLowerCase());
      if (!ae) throw f("#REF!");
      return ae;
    },
    G = () => {
      const te = Fie(
        new JN({
          onVariable: () => null,
          onCell: ({ sheet: le, row: ce, col: ge }) => {
            const se = U(le);
            if (
              !Number.isInteger(ce) ||
              !Number.isInteger(ge) ||
              ce < 1 ||
              ge < 1 ||
              ce > se.rowCount ||
              ge > se.colCount
            )
              throw f("#REF!");
            return Q(se.id, `${ut(ge - 1)}${ce}`);
          },
          onRange: ({ sheet: le, from: ce, to: ge }) => {
            const se = U(le),
              we = ge.row,
              Ee = ge.col;
            if (
              ce.row < 1 ||
              ce.col < 1 ||
              we > se.rowCount ||
              Ee > se.colCount ||
              we < ce.row ||
              Ee < ce.col
            )
              throw f("#REF!");
            if ((we - ce.row + 1) * (Ee - ce.col + 1) > 2e4) throw f("#LIMIT!");
            const Ae = s.reserve(
              (we - ce.row + 1) * (32 + (Ee - ce.col + 1) * 8),
            );
            if (Ae) throw f(Ae);
            const ye = [];
            for (let Te = ce.row; Te <= we; Te++) {
              const Pe = [];
              for (let qe = ce.col; qe <= Ee; qe++) {
                const ht = Q(se.id, `${ut(qe - 1)}${Te}`);
                if (
                  wn(ht) &&
                  ["#LIMIT!", "#TIMEOUT!", "#BLOCKED!"].includes(ht.error)
                )
                  throw ht;
                Pe.push(ht);
              }
              ye.push(Pe);
            }
            return ye;
          },
        }),
        (le, ce, ge) => {
          return p(le, ce, {
            currentSheetName: ge.position?.sheet,
            retrieve: (we) => ge.retrieveRef(we),
            issue: f,
            isIssue: wn,
            readCell: Q,
            readMetadata:
              t || n
                ? (we, Ee) => {
                    if (
                      (n == null || n.read(`${we}:${Ee}`),
                      t && !t.allowRead(we, Ee))
                    )
                      throw f("#BLOCKED!");
                    t == null ||
                      t.step({
                        kind: "metadata",
                        sheetId: we,
                        address: Ee,
                        operation: "Inspect cell metadata",
                        result: null,
                      });
                  }
                : undefined,
            isActive: (we, Ee) => y.has(`${we}:${Ee}`),
            expandedFormula: (we, Ee) => {
              try {
                return pf(String(we.raw), e, Ee);
              } catch {
                return String(we.raw);
              }
            },
          });
        },
        f,
        d,
        (le) => {
          const ce = te.position;
          if (!ce) return;
          const ge = U(ce.sheet).id,
            se = `${ut(ce.col - 1)}${ce.row}`,
            we = `${ge}:${se}:${le.functionName}`;
          if (h.size >= 1e4 && !h.has(we)) throw f("#LIMIT!");
          h.set(we, { ...le, sheetId: ge, address: se });
        },
        (le, ce, ge) => {
          const se =
            le && typeof le == "object" && "ref" in le ? le.ref : undefined;
          if (!(se != null && se.from) || !se.to) return;
          const we = U(se.sheet ?? te.position?.sheet),
            Ee = `${ut(se.from.col + ge - 1)}${se.from.row + ce}`;
          if (t && !t.allowRead(we.id, Ee)) throw f("#BLOCKED!");
          n == null || n.read(`${we.id}:${Ee}`);
          const Ae = we.cells[Ee];
          if (
            (Ae != null && Ae.literal) ||
            typeof Ae?.raw != "string" ||
            !Ae.raw.startsWith("=")
          )
            return;
          const ye = Ae.raw;
          return (Pe) => {
            if (O >= 32) throw f("#LIMIT!");
            const qe = O++;
            try {
              if (
                t != null &&
                t.allowDatabaseCriterion &&
                !t.allowDatabaseCriterion(we.id, xd(ye, Pe, 0))
              )
                throw f("#BLOCKED!");
              const ht = pf(ye, e, we.id, [], t?.allowName),
                Re = xd(ht, Pe, 0);
              if (Re.length > Ja.maxFormulaLength) throw f("#LIMIT!");
              return V(Re)
                ? H(Re.slice(1), we.id, Ee)
                : (I[qe] ?? (I[qe] = G())).parse(
                    $(Re.slice(1), we.id, Ee),
                    {
                      sheet: we.name,
                      row: se.from.row + ce,
                      col: se.from.col + ge,
                    },
                    true,
                  );
            } finally {
              O--;
            }
          };
        },
      );
      if (o.length) {
        const le = te._callFunction.bind(te);
        te._callFunction = (ce, ge) => {
          const se = ce.replace(/^_xlfn\./i, "").toUpperCase();
          if (!o.some((ht) => ht.name === se)) return le(ce, ge);
          const we = te.position;
          if (!we) return f("#NAME?");
          const Ee = U(we.sheet).id,
            Ae = `${ut(we.col - 1)}${we.row}`,
            ye = ge.map((ht) => te.retrieveRef(ht)),
            Te = M5(ye);
          if (Te) return Te;
          const Pe = LJ(e, se, ye, Ee, Ae);
          if (Pe.error) return f(Pe.error);
          const qe = Pe.value;
          return qe && !Array.isArray(qe) && typeof qe == "object"
            ? qe.label
            : qe;
        };
      }
      const ae = (le) => {
          const ce = s.work() ?? n?.check();
          if (ce) return f(ce);
          const ge = le(),
            se = s.value(ge) ?? n?.check();
          return se ? f(se) : ge;
        },
        oe = te._callFunction.bind(te),
        de = te.utils._applyInfix.bind(te.utils),
        he = te.utils._applyPrefix.bind(te.utils),
        Se = te.utils._applyPostfix.bind(te.utils);
      if (
        ((te._callFunction = (le, ce) => ae(() => oe(le, ce))),
        (te.utils._applyInfix = (le, ce, ge) => ae(() => de(le, ce, ge))),
        (te.utils._applyPrefix = (le, ce, ge) => ae(() => he(le, ce, ge))),
        (te.utils._applyPostfix = (le, ce, ge) => ae(() => Se(le, ce, ge))),
        t)
      ) {
        const le = (Ee, Ae, ye, Te) => {
            const Pe = te.position,
              qe = U(Pe.sheet).id,
              ht = `${ut(Pe.col - 1)}${Pe.row}`;
            try {
              const Re = Te();
              return (
                t.step({
                  kind: Ee,
                  sheetId: qe,
                  address: ht,
                  operation: Ae,
                  inputs: ye,
                  result: Re,
                }),
                Re
              );
            } catch (Re) {
              throw (
                t.step({
                  kind: Ee,
                  sheetId: qe,
                  address: ht,
                  operation: Ae,
                  inputs: ye,
                  result: wn(Re) ? Re : f("#ERROR!"),
                }),
                Re
              );
            }
          },
          ce = te._callFunction.bind(te),
          ge = te.utils._applyInfix.bind(te.utils),
          se = te.utils._applyPrefix.bind(te.utils),
          we = te.utils._applyPostfix.bind(te.utils);
        ((te._callFunction = (Ee, Ae) =>
          le("function", Ee, Ae, () => ce(Ee, Ae))),
          (te.utils._applyInfix = (Ee, Ae, ye) =>
            le("operator", Ae, [Ee.val, ye.val], () => ge(Ee, Ae, ye))),
          (te.utils._applyPrefix = (Ee, Ae, ye) =>
            le("operator", Ee.join(""), [Ae], () => se(Ee, Ae, ye))),
          (te.utils._applyPostfix = (Ee, Ae, ye) =>
            le("operator", ye, [Ee], () => we(Ee, Ae, ye))));
      }
      return te;
    },
    H = (te, ae, oe) =>
      pne(te, e, ae, {
        supported: a,
        issue: f,
        isIssue: wn,
        allowName: t?.allowName,
        step: t
          ? (de, he, Se) =>
              t.step({
                kind: de,
                sheetId: ae,
                address: oe,
                operation: he,
                result: Se,
              })
          : undefined,
        diagnostic: (de, he) =>
          c.set(`${ae}:${oe}`, {
            sheetId: ae,
            address: oe,
            error: de,
            message: he,
          }),
        parse: (de, he, Se) => {
          const le = $(de, he, oe, false),
            ce = le.replace(/"(?:[^"]|"")*"/g, '""');
          if (/[\[\]|]/.test(ce) || /(?:https?:|file:|\\\\)/i.test(ce))
            throw f("#BLOCKED!");
          if (
            Array.from(
              ce
                .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
                .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
            ).some((Ee) => Ee[1] !== y5 && !a.has(Ee[1].toUpperCase()))
          )
            throw f("#NAME?");
          const ge = F++,
            se = k[ge] ?? (k[ge] = G()),
            we = se._callFunction;
          se._callFunction = (Ee, Ae) =>
            Ee === y5
              ? Ae.length === 1 &&
                Number.isInteger(Ae[0]) &&
                Number(Ae[0]) >= 0 &&
                Number(Ae[0]) < Se.length
                ? Se[Number(Ae[0])]
                : f("#NAME?")
              : we(Ee, Ae);
          try {
            const Ee = _t(oe);
            return se.parse(
              le,
              { sheet: _.get(he).name, row: Ee.row + 1, col: Ee.col + 1 },
              true,
            );
          } finally {
            ((se._callFunction = we), F--);
          }
        },
      }),
    Q = (te, ae) => {
      const oe = s.work() ?? n?.check();
      if (oe) return f(oe);
      if (!t) return X(te, ae);
      if (!t.allowRead(te, ae)) return f("#BLOCKED!");
      const de = X(te, ae);
      return (
        t.step({
          kind: "read",
          sheetId: te,
          address: ae,
          operation: "Read cell",
          result: de,
        }),
        de
      );
    },
    X = (te, ae) => {
      var ye;
      if (++T > Mt.maxCells) return f("#LIMIT!");
      let oe, de;
      try {
        ({ row: oe, col: de } = _t(ae));
      } catch {
        return f("#REF!");
      }
      const he = _.get(te);
      if (!he || oe >= he.rowCount || de >= he.colCount) return f("#REF!");
      ae = `${ut(de)}${oe + 1}`;
      const Se = `${te}:${ae}`,
        le = D.values.get(Se);
      if (le) return le.type === "error" ? f(String(le.value)) : le.value;
      n == null || n.read(Se);
      const ce = n?.value(Se, y.size > 0 || w.size > 0);
      if (ce) return ce.type === "error" ? f(String(ce.value)) : ce.value;
      for (const Te of y) (ye = E.get(Te)) == null || ye.add(Se);
      if (y.has(Se) || w.has(Se))
        return (n == null || n.cycle([...y, ...w, Se]), f("#CYCLE!"));
      if (m.has(Se)) return m.get(Se);
      if (v.has(Se)) return v.get(Se).value;
      const ge = he.cells[ae];
      if (!ge || ge.raw === null) {
        const Te = R.get(te) ?? [];
        if (Te.length > 1e3) return f("#LIMIT!");
        for (const Pe of Te) {
          if (Pe === ae) continue;
          const qe = _t(Pe);
          if (
            !(
              qe.row > oe ||
              qe.col > de ||
              (oe - qe.row + 1) * (de - qe.col + 1) > Ja.maxSpillCells
            ) &&
            !(y.has(`${te}:${Pe}`) || w.has(`${te}:${Pe}`)) &&
            (Q(te, Pe), v.has(Se))
          )
            return v.get(Se).value;
        }
        return ge?.raw ?? null;
      }
      if (typeof ge.raw != "string" || !ge.raw.startsWith("=") || ge.literal)
        return ge.raw;
      const se = ge.raw;
      n == null || n.enter(Se);
      let Ee = (() => {
        var $e;
        if (
          y.size + w.size >= Ja.maxFormulaDepth ||
          se.length > Ja.maxFormulaLength ||
          !O4(se)
        )
          return f("#LIMIT!");
        let Te;
        const Pe = V(se);
        w.add(Se);
        try {
          if (((Te = Pe ? se.slice(1) : $(se.slice(1), te, ae)), !O4(Te)))
            return f("#LIMIT!");
        } catch (Je) {
          return wn(Je)
            ? f(Je.error)
            : f(
                Je instanceof Error && /^#/.test(Je.message)
                  ? Je.message
                  : "#NAME?",
              );
        } finally {
          w.delete(Se);
        }
        const qe = Te.replace(/"(?:[^"]|"")*"/g, '""');
        if (!Pe && (/[\[\]|]/.test(qe) || /(?:https?:|file:|\\\\)/i.test(qe)))
          return f("#BLOCKED!");
        if (
          !Pe &&
          Array.from(
            qe
              .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
              .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
          ).some((Je) => !a.has(Je[1].toUpperCase()))
        )
          return (
            c.set(Se, {
              sheetId: te,
              address: ae,
              error: "#NAME?",
              message:
                "The formula calls an unknown or unsupported custom function. Imported VBA, add-in and external custom functions cannot execute.",
            }),
            f("#NAME?")
          );
        const ht = y.size;
        (y.add(Se), E.set(Se, new Set()));
        let Re;
        try {
          const Je = A[ht] ?? (A[ht] = G()),
            ze = Pe
              ? H(Te, te, ae)
              : Je.parse(
                  Te,
                  { sheet: he.name, row: oe + 1, col: de + 1 },
                  true,
                );
          if (wn(ze)) Re = f(ze.error);
          else if (typeof ze == "number")
            Re = Number.isFinite(ze) ? ze : f("#NUM!");
          else if (ze === null) Re = 0;
          else if (typeof ze == "boolean") Re = ze;
          else if (typeof ze == "string")
            Re = ze.length <= 32767 ? ze : f("#LIMIT!");
          else if (Array.isArray(ze)) {
            const at = ze.every(Array.isArray) ? ze : [ze],
              _e = at[0]?.length ?? 0;
            if (!at.length || !_e || at.some((Ze) => Ze.length !== _e))
              Re = f("#CALC!");
            else if (
              at.length * _e > Ja.maxSpillCells ||
              N + at.length * _e > Mt.maxCells
            )
              Re = f("#LIMIT!");
            else if (oe + at.length > he.rowCount || de + _e > he.colCount)
              (b.set(Se, {
                sheetId: te,
                anchor: ae,
                reason: "bounds",
                message:
                  "The array extends beyond the worksheet edge. Move the formula or reduce its output.",
              }),
                (Re = f("#SPILL!")));
            else {
              let Ze = false,
                lt = false;
              const be = {
                  firstRow: oe,
                  lastRow: oe + at.length - 1,
                  firstCol: de,
                  lastCol: de + _e - 1,
                },
                ft = (nt) => {
                  const mt = ct(nt);
                  return (
                    mt.firstRow <= be.lastRow &&
                    mt.lastRow >= be.firstRow &&
                    mt.firstCol <= be.lastCol &&
                    mt.lastCol >= be.firstCol
                  );
                },
                pt = he.merges?.find(ft),
                Fe = he.tables?.find((nt) => ft(nt.range));
              if (pt || Fe) {
                Ze = true;
                const nt = pt ?? Fe.range;
                b.set(Se, {
                  sheetId: te,
                  anchor: ae,
                  reason: pt ? "merged" : "table",
                  obstruction: nt,
                  message: pt
                    ? `The array overlaps merged cells ${nt}. Unmerge those cells or move the formula.`
                    : `The array overlaps table ${Fe.name} (${nt}). Place the formula outside the table.`,
                });
              }
              for (let nt = 0; nt < at.length && !Ze; nt++)
                for (let mt = 0; mt < _e; mt++) {
                  const Wt = `${ut(de + mt)}${oe + nt + 1}`,
                    Zt = he.cells[Wt];
                  if (($e = E.get(Se)) != null && $e.has(`${te}:${Wt}`)) {
                    lt = true;
                    break;
                  }
                  if (
                    Wt !== ae &&
                    ((Zt && Zt.raw !== null) || v.has(`${te}:${Wt}`))
                  ) {
                    ((Ze = true),
                      b.set(Se, {
                        sheetId: te,
                        anchor: ae,
                        reason: "occupied",
                        obstruction: Wt,
                        message: `Cell ${Wt} blocks the array. Move the formula or review and clear the obstruction; existing evidence was preserved.`,
                      }));
                    break;
                  }
                }
              if (lt) Re = f("#CYCLE!");
              else if (Ze) Re = f("#SPILL!");
              else {
                const nt = (mt) =>
                  wn(mt)
                    ? mt
                    : mt === null
                      ? 0
                      : typeof mt == "number"
                        ? Number.isFinite(mt)
                          ? mt
                          : f("#NUM!")
                        : typeof mt == "boolean"
                          ? mt
                          : typeof mt == "string"
                            ? mt.length <= 32767
                              ? mt
                              : f("#LIMIT!")
                            : f("#VALUE!");
                for (let mt = 0; mt < at.length; mt++)
                  for (let Wt = 0; Wt < _e; Wt++) {
                    const Zt = `${ut(de + Wt)}${oe + mt + 1}`;
                    Zt !== ae &&
                      v.set(`${te}:${Zt}`, {
                        value: nt(at[mt][Wt]),
                        anchor: ae,
                      });
                  }
                (g.set(Se, `${ae}:${ut(de + _e - 1)}${oe + at.length}`),
                  (N += at.length * _e),
                  (Re = nt(at[0][0])));
              }
            }
          } else Re = f("#VALUE!");
        } catch (Je) {
          const ze =
            Je && typeof Je == "object" && "details" in Je
              ? Je.details
              : undefined;
          Re = wn(ze) ? f(ze.error) : wn(Je) ? f(Je.error) : f("#ERROR!");
        } finally {
          (y.delete(Se), E.delete(Se));
        }
        return Re;
      })();
      const Ae = s.value(Ee) ?? n?.check();
      return (
        Ae && (Ee = f(Ae)),
        m.set(Se, Ee),
        n == null ||
          n.leave(
            Se,
            wn(Ee)
              ? { type: "error", value: Ee.error }
              : {
                  type:
                    Ee === null
                      ? "blank"
                      : typeof Ee == "string"
                        ? "text"
                        : typeof Ee == "boolean"
                          ? "boolean"
                          : "number",
                  value: Ee,
                },
          ),
        t == null ||
          t.step({
            kind: "formula",
            sheetId: te,
            address: ae,
            operation: ge.raw,
            result: Ee,
          }),
        Ee
      );
    },
    J = (te, ae) => J.getTypedValue(te, ae).value;
  J.getTypedValue = (te, ae) => {
    if (n?.running?.call(n) === false)
      try {
        const Se = _t(ae),
          le = _.get(te);
        if (!le || Se.row >= le.rowCount || Se.col >= le.colCount)
          return { type: "error", value: "#REF!" };
        const ce = `${ut(Se.col)}${Se.row + 1}`,
          ge = `${te}:${ce}`,
          se = n.value(ge, false);
        if (se) return { ...se };
        const we = m.get(ge) ?? v.get(ge)?.value ?? le.cells[ce]?.raw ?? null;
        return wn(we)
          ? { type: "error", value: we.error }
          : {
              type:
                we === null
                  ? "blank"
                  : typeof we == "string"
                    ? "text"
                    : typeof we == "boolean"
                      ? "boolean"
                      : "number",
              value: we,
            };
      } catch {
        return { type: "error", value: "#REF!" };
      }
    s.begin();
    try {
      T = 0;
      const Se = Q(te, ae);
      return wn(Se)
        ? { type: "error", value: Se.error }
        : {
            type:
              Se === null
                ? "blank"
                : typeof Se == "string"
                  ? "text"
                  : typeof Se == "boolean"
                    ? "boolean"
                    : "number",
            value: Se,
          };
    } finally {
      s.end();
    }
  };
  let q;
  return (
    (J.evaluateFormula = (te, ae, oe) => {
      (s.begin(), (T = 0), l.delete(`${te}:${ae}`), c.delete(`${te}:${ae}`));
      const de = `expression:${te}:${ae}`;
      try {
        const he = _.get(te),
          Se = _t(ae);
        if (!he || Se.row >= he.rowCount || Se.col >= he.colCount)
          return "#REF!";
        if (typeof oe != "string" || !oe.startsWith("=")) return "#VALUE!";
        if (oe.length > Ja.maxFormulaLength || !O4(oe))
          return f("#LIMIT!").error;
        const le = V(oe),
          ce = le ? oe.slice(1) : $(oe.slice(1), te, ae);
        if (!O4(ce)) return f("#LIMIT!").error;
        const ge = ce.replace(/"(?:[^"]|"")*"/g, '""');
        if (!le && (/[\[\]|]/.test(ge) || /(?:https?:|file:|\\\\)/i.test(ge)))
          return f("#BLOCKED!").error;
        if (
          !le &&
          Array.from(
            ge
              .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
              .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
          ).some((Ee) => !a.has(Ee[1].toUpperCase()))
        )
          return "#NAME?";
        (y.add(de), E.set(de, new Set()));
        const se = le
            ? H(ce, te, ae)
            : (q ?? (q = G())).parse(
                ce,
                { sheet: he.name, row: Se.row + 1, col: Se.col + 1 },
                true,
              ),
          we = s.value(se) ?? n?.check();
        return we
          ? f(we).error
          : wn(se)
            ? f(se.error).error
            : Array.isArray(se)
              ? "#VALUE!"
              : se === null
                ? 0
                : typeof se == "number"
                  ? Number.isFinite(se)
                    ? se
                    : "#NUM!"
                  : typeof se == "boolean"
                    ? se
                    : typeof se == "string"
                      ? se.length <= 32767
                        ? se
                        : f("#LIMIT!").error
                      : "#VALUE!";
      } catch (he) {
        const Se =
          he && typeof he == "object" && "details" in he
            ? he.details
            : undefined;
        return wn(Se)
          ? f(Se.error).error
          : wn(he)
            ? f(he.error).error
            : he instanceof Error && /^#/.test(he.message)
              ? f(he.message).error
              : f("#ERROR!").error;
      } finally {
        (s.end(), y.delete(de), E.delete(de));
      }
    }),
    (J.getSpillRange = (te, ae) => (J(te, ae), g.get(`${te}:${ae}`))),
    (J.getSpillDiagnostic = (te, ae) => {
      J(te, ae);
      const oe = b.get(`${te}:${ae}`);
      return oe ? { ...oe } : undefined;
    }),
    (J.getSpillAnchor = (te, ae) => {
      return (J(te, ae), v.get(`${te}:${ae}`)?.anchor);
    }),
    (J.getSpillEntries = (te) => {
      s.begin();
      try {
        const oe = R.get(te) ?? [];
        if (oe.length > 1e3)
          throw (
            f("#LIMIT!"),
            new Error("Spill discovery exceeds 1,000 possible array formulas.")
          );
        T = 0;
        for (const he of n?.running?.call(n) === false ? [] : oe)
          if ((Q(te, he), T > Mt.maxCells))
            throw new Error(
              "Spill discovery exceeds 100,000 calculation reads.",
            );
        const de = `${te}:`;
        return [...v.entries()]
          .filter(([he]) => he.startsWith(de))
          .map(([he, { anchor: Se }]) => ({
            address: he.slice(de.length),
            anchor: Se,
          }));
      } finally {
        s.end();
      }
    }),
    (J.getResourceReport = () => s.report()),
    (J.getOperationalErrors = () => [...u]),
    (J.getDateContext = () => ({ ...d })),
    (J.getFinancialDiagnostics = () => structuredClone([...h.values()])),
    (J.peekSpill = (te, ae) => {
      const oe = `${te}:${ae}`,
        de = b.get(oe);
      return { range: g.get(oe), ...(de ? { diagnostic: { ...de } } : {}) };
    }),
    (J.getNamedFormulaDiagnostics = () =>
      [...c.values()].map((te) => ({ ...te }))),
    (J.getReferenceDiagnostics = () =>
      [...l.values()].map((te) => ({ ...te }))),
    J
  );
}
function OB(e) {
  const r = Cf(e);
  return (t, n, o) => {
    const a = r(t, n);
    let s = false,
      d = "Value does not satisfy the rule.",
      u,
      l = false;
    try {
      if (r.getTypedValue(t, n).type === "error")
        throw Error(`Cell calculation returned ${a}.`);
      if (a === null || a === "") s = o.allowBlank;
      else if (o.kind === "custom") {
        const c = _t(n),
          h = ct(o.range);
        u = xd(o.formula, c.row - h.firstRow, c.col - h.firstCol);
        const f = r.evaluateFormula(t, n, u);
        if (
          typeof f != "boolean" &&
          (typeof f != "number" || !Number.isFinite(f))
        )
          throw Error(
            `Custom formula must return TRUE/FALSE or a finite number${typeof f == "string" && /^#[A-Z0-9/?!]+$/.test(f) ? `; returned ${f}` : ""}.`,
          );
        ((s = !!f), (d = "Custom formula returned FALSE or zero."));
      } else s = jw(o, a, e.datePolicy);
    } catch (c) {
      ((l = true),
        (d =
          c instanceof Error
            ? c.message
            : "Validation could not be evaluated."));
    }
    if (!(s && !l))
      return {
        sheetId: t,
        sheetName: e.sheets.find((c) => c.id === t).name,
        address: n,
        ruleId: o.id,
        style: l ? "stop" : (o.errorStyle ?? "stop"),
        value: a,
        title: o.errorTitle || "Invalid value",
        message:
          o.message || "The value does not satisfy the data validation rule.",
        reason: d,
        ...(u ? { formula: u } : {}),
        evaluationError: l,
      };
  };
}
function v0(e, r, t) {
  const n = r.flatMap((d) =>
    DB(e, d.sheetId, d.address).map((u) => ({ ...d, rule: u })),
  );
  if (!n.length) return 0;
  if (n.length > a1.checks)
    throw Error(
      "Data validation exceeds the 100,000-check budget; split the batch.",
    );
  const o = OB(e),
    a = n.map((d) => o(d.sheetId, d.address, d.rule)).filter((d) => !!d);
  if (!a.length) return 0;
  const s = JSON.stringify({ book: e, issues: a });
  if (
    t &&
    r5.get(t) === s &&
    a.every((d) => d.style !== "stop" && !d.evaluationError)
  )
    return (r5.delete(t), a.length);
  throw t
    ? Error(
        "Validation approval is stale or invalid. Review the complete batch again.",
      )
    : new jN(e, a);
}
function LQ(e, r) {
  const t = {
      workbookId: e.id,
      revision: e.revision,
      checks: 0,
      invalidCells: 0,
      violations: 0,
      issues: [],
      complete: true,
      detailsTruncated: false,
      message: "",
    },
    n = OB(e),
    o = new Set();
  let a = 0;
  e: for (const s of e.sheets.filter((d) => !r || d.id === r))
    for (const d of s.validations ?? []) {
      const u = ct(d.range);
      for (let l = u.firstRow; l <= u.lastRow; l++)
        for (let c = u.firstCol; c <= u.lastCol; c++) {
          if (
            t.checks >= a1.checks ||
            (d.kind === "custom" && ++a > a1.customChecks)
          ) {
            t.complete = false;
            break e;
          }
          const h = `${ut(c)}${l + 1}`;
          t.checks++;
          const f = n(s.id, h, d);
          f &&
            (o.add(`${s.id}:${h}`),
            t.violations++,
            t.issues.length < a1.displayedIssues && t.issues.push(f));
        }
    }
  return (
    (t.invalidCells = o.size),
    (t.detailsTruncated = t.violations > t.issues.length),
    (t.message = `${t.complete ? "" : "Incomplete inspection: at least "}${t.invalidCells} invalid cells; ${t.violations} rule violations across ${t.checks} checks.${t.complete ? "" : " The 100,000-check or 5,000-custom-formula limit was reached. Narrow validation ranges and inspect again."}${t.detailsTruncated ? " Showing the first 500 violations." : ""}`),
    FB.set(t, { source: JSON.stringify(e), content: JSON.stringify(t) }),
    t
  );
}
function cU(e, r, t, n) {
  const o = new Map(n.spills.map((c) => [`${c.sheetId}:${c.address}`, c])),
    a = new Map(n.spillMetadata.map((c) => [`${c.sheetId}:${c.address}`, c])),
    s = (c, h) => `${c}:${h.replace(/\$/g, "").toUpperCase()}`,
    d = new Map(e.sheets.map((c) => [c.id, c])),
    u = (c, h) => {
      const f = s(c, h),
        p = t.get(f) ?? o.get(f)?.result;
      if (p) return { ...p };
      const m = d.get(c);
      if (!m) return { type: "error", value: "#REF!" };
      try {
        const y = _t(h);
        if (y.row >= m.rowCount || y.col >= m.colCount)
          return { type: "error", value: "#REF!" };
      } catch {
        return { type: "error", value: "#REF!" };
      }
      const v = m.cells[h.replace(/\$/g, "").toUpperCase()]?.raw ?? null;
      return {
        type:
          v === null
            ? "blank"
            : typeof v == "string"
              ? "text"
              : typeof v == "boolean"
                ? "boolean"
                : "number",
        value: v,
      };
    },
    l = (c, h) => u(c, h).value;
  return (
    Object.assign(l, {
      getTypedValue: u,
      getDateContext: () => ({ ...r.dateContext }),
      getResourceReport: () => ({ ...r.resources }),
      getOperationalErrors: () => [...n.operationalErrors],
      getReferenceDiagnostics: () => structuredClone(n.references),
      getNamedFormulaDiagnostics: () => structuredClone(n.names),
      getFinancialDiagnostics: () => structuredClone(n.financial),
      getSpillEntries: (c) =>
        n.spills
          .filter((h) => h.sheetId === c)
          .map(({ address: h, anchor: f }) => ({ address: h, anchor: f })),
      getSpillAnchor: (c, h) => {
        return o.get(s(c, h))?.anchor;
      },
      getSpillRange: (c, h) => {
        return a.get(s(c, h))?.range;
      },
      getSpillDiagnostic: (c, h) => {
        return structuredClone(a.get(s(c, h))?.diagnostic);
      },
      peekSpill: (c, h) => {
        const f = a.get(s(c, h));
        return f
          ? { range: f.range, diagnostic: structuredClone(f.diagnostic) }
          : {};
      },
      evaluateFormula: (c, h, f) => Cf(e).evaluateFormula(c, h, f),
    }),
    l
  );
}
function L5(e) {
  const r = tu.get(e);
  if (r?.report.execution?.state === "pending" && nm(r, e))
    return { calculator: r.calculator, report: r.report };
  const t = pd(),
    n = W1(e),
    o = tu.get(e) ?? ch.get(e),
    a = new Map(),
    s = new Map(),
    d = qie(o),
    u = new Map();
  for (const v of n.formulas.keys()) {
    a.set(v, ng);
    const g = d.get(v);
    g && u.set(v, g);
  }
  let l;
  const c = {
    workbookId: e.id,
    revision: e.revision,
    run: (o?.report.run ?? 0) + 1,
    settings: n.settings,
    resources: SN(e).report(),
    volatile: { functions: n.volatile, policy: yE },
    dateContext: lx(e.datePolicy),
    reason: "Waiting for isolated calculation",
    formulaCount: n.formulas.size,
    evaluated: 0,
    reused: 0,
    pending: n.formulas.size,
    elapsedMs: pd() - t,
    dependencyEdges: 0,
    cycles: [],
    iteration: {
      status: n.settings.iterative ? "not-needed" : "off",
      passes: 0,
      maxDelta: null,
      message: "Waiting for isolated calculation",
    },
    execution: {
      mode: "worker",
      state: "pending",
      message:
        "Calculation is running separately; source inputs remain editable.",
    },
    limits: {
      dependencyEdges: j5,
      cooperativeTimeBudgetMs: n.settings.timeBudgetMs,
      hardInterrupt: true,
    },
  };
  A1(c);
  const h = Object.freeze({
    ...c,
    get cells() {
      return (
        l ||
          ((l = [...n.formulas].map(([v, g]) => ({
            ...g,
            result: { ...ng },
            status: "pending",
            evaluations: 0,
            dependencies: [],
            ...(u.has(v) ? { lastKnown: u.get(v) } : {}),
          }))),
          A1(l)),
        l
      );
    },
  });
  PE.set(h, u);
  const f = [];
  if (o)
    for (const v of e.sheets)
      for (const g of o.calculator.getSpillEntries(v.id))
        n.formulas.has(`${v.id}:${g.anchor}`) &&
          f.push({ ...g, sheetId: v.id, result: { ...ng } });
  const p = cU(e, h, a, {
    spills: f,
    spillMetadata: [],
    references: [],
    names: [],
    financial: [],
    operationalErrors: n.formulas.size ? ["#PENDING!"] : [],
  });
  return (
    tu.set(e, { ...n, calculator: p, report: h, values: a, dependencies: s }),
    ch.delete(e),
    { calculator: p, report: h }
  );
}
function uU(e, r, t) {
  const n = tu.get(e),
    o = W1(e);
  if (!n || n.report !== r || !nm(n, e, o)) return;
  if (
    t.inputSignature !== lU(e) ||
    t.report.workbookId !== e.id ||
    t.report.revision !== e.revision ||
    t.report.cells.length !== o.formulas.size
  )
    throw Error("The worker result does not match the requested workbook.");
  const a = new Map(),
    s = new Map();
  for (const c of t.report.cells) {
    const h = `${c.sheetId}:${c.address}`;
    if (a.has(h) || o.formulas.get(h)?.formula !== c.formula)
      throw Error(
        "The worker result contains stale or duplicate formula inputs.",
      );
    (a.set(h, c.result), s.set(h, new Set(c.dependencies)));
  }
  const d = { ...t.report, run: r.run };
  A1(d);
  const u = cU(e, d, a, t);
  return (
    tu.set(e, { ...o, calculator: u, report: d, values: a, dependencies: s }),
    { calculator: u, report: d }
  );
}
function Xie(e, r, t, n, o) {
  if (t === "#PENDING!") {
    const s = tu.get(e);
    if (!s || s.report !== r || !nm(s, e)) return;
    const d = Object.freeze({
        mode: "worker",
        state: "cancelled",
        message: n,
        wallElapsedMs: o,
      }),
      u = Object.freeze(
        Object.defineProperties(
          {},
          {
            ...Object.getOwnPropertyDescriptors(r),
            reason: { value: n, enumerable: true },
            execution: { value: d, enumerable: true },
            iteration: {
              value: Object.freeze({
                ...r.iteration,
                status: r.settings.iterative ? "failed" : "off",
                message: n,
                passes: 0,
                maxDelta: null,
              }),
              enumerable: true,
            },
          },
        ),
      ),
      l = PE.get(r);
    return (
      l && PE.set(u, l),
      tu.set(e, { ...s, report: u }),
      { calculator: s.calculator, report: u }
    );
  }
  const a = r.cells.map((s) => ({
    ...s,
    result: F5(t),
    status: "error",
    evaluations: 0,
    dependencies: [],
  }));
  return uU(e, r, {
    inputSignature: lU(e),
    report: {
      ...r,
      cells: a,
      pending: 0,
      evaluated: 0,
      reused: 0,
      dependencyEdges: 0,
      cycles: [],
      reason: n,
      execution: {
        mode: "worker",
        state: t === "#TIMEOUT!" ? "timeout" : "failed",
        message: n,
        wallElapsedMs: o,
      },
      limits: { ...r.limits, hardInterrupt: true },
      iteration: {
        ...r.iteration,
        status: r.settings.iterative ? "failed" : "off",
        message: n,
        passes: 0,
        maxDelta: null,
      },
    },
    spills: [],
    spillMetadata: [],
    references: [],
    names: [],
    financial: [],
    operationalErrors: [t],
  });
}
function dU(e, r, t, n) {
  const o = pd(),
    { formulas: a, settings: s } = r,
    d = lx(e.datePolicy),
    u = SN(e);
  u.begin();
  const l = new Set(),
    c = new Map(),
    h = new Map(),
    f = new Map();
  let p = n
    ? "Explicit full recalculation"
    : t
      ? "Changed inputs and transitive dependents"
      : "First calculation of this snapshot";
  if (n || !t || r.metadata !== t.metadata || ik(e, a) || ik(e, t.formulas))
    (a.forEach((U, G) => l.add(G)),
      t &&
        !n &&
        (p =
          r.metadata !== t.metadata
            ? "Calculation settings or structural metadata changed"
            : "Conservative recalculation for names, dynamic references or arrays"));
  else {
    const U = new Set(),
      G = r.inputs(),
      H = t.inputs();
    for (const J of new Set([...G.keys(), ...H.keys()]))
      G.get(J) !== H.get(J) && U.add(J);
    (a.forEach((J, q) => {
      (_O.test(J.formula) ||
        !t.values.has(q) ||
        ["#TIMEOUT!", "#LIMIT!"].includes(String(t.values.get(q)?.value))) &&
        U.add(q);
    }),
      t.report.cycles.flat().forEach((J) => U.add(J)));
    const Q = new Map();
    t.dependencies.forEach((J, q) =>
      J.forEach((W) => {
        const K = Q.get(W) ?? new Set();
        (K.add(q), Q.set(W, K));
      }),
    );
    const X = [...U];
    for (let J = 0; J < X.length; J++)
      for (const q of Q.get(X[J]) ?? []) U.has(q) || (U.add(q), X.push(q));
    U.forEach((J) => {
      a.has(J) && l.add(J);
    });
  }
  a.forEach((U, G) => {
    !l.has(G) &&
      t != null &&
      t.values.has(G) &&
      (h.set(G, t.values.get(G)), c.set(G, new Set(t.dependencies.get(G))));
  });
  const v = new Map(h),
    g = [],
    b = new Set();
  let y = true,
    w,
    E = [...c.values()].reduce((U, G) => U + G.size, 0);
  const A = () => (
      y && !w && (w = u.check()),
      y && !w && pd() - o > s.timeBudgetMs && (w = "#TIMEOUT!"),
      y ? w : undefined
    ),
    k = (U, G) => {
      const H = [];
      return Cf(e, {}, undefined, {
        budget: u,
        running: () => y,
        dateContext: d,
        check: A,
        read(X) {
          const J = H.at(-1);
          if (!y || !J) return;
          const q = c.get(J);
          q.has(X) || (E >= j5 ? (w = "#LIMIT!") : (q.add(X), E++));
        },
        value(X, J) {
          return J && G != null && G.has(X) ? G.get(X) : U.get(X);
        },
        enter(X) {
          (H.push(X),
            (E -= c.get(X)?.size ?? 0),
            c.set(X, new Set()),
            f.set(X, (f.get(X) ?? 0) + 1));
        },
        leave(X, J) {
          (h.set(X, J), H.pop());
        },
        cycle(X) {
          const J = X.at(-1),
            q = [...new Set(X.slice(X.indexOf(J), -1))];
          (q.forEach((W) => b.add(W)),
            q.length &&
              !g.some(
                (W) => W.length === q.length && W.every((K) => q.includes(K)),
              ) &&
              g.push(q));
        },
      });
    },
    F = s.mode === "manual" && !n;
  F &&
    l.forEach((U) => {
      (h.set(U, ng), v.set(U, ng), c.set(U, new Set(t?.dependencies.get(U))));
    });
  let I = k(v);
  const O = () => {
    for (const [U, G] of a)
      h.set(U, v.get(U) ?? I.getTypedValue(G.sheetId, G.address));
  };
  F || O();
  const j = {
    status: s.iterative ? "not-needed" : "off",
    passes: 0,
    maxDelta: null,
    message: s.iterative
      ? "No circular calculation was required."
      : "Iteration is disabled. Circular references return #CYCLE!.",
  };
  if (!F && b.size && s.iterative) {
    const U = new Set(b),
      G = new Map();
    c.forEach((Q, X) =>
      Q.forEach((J) => {
        const q = G.get(J) ?? new Set();
        (q.add(X), G.set(J, q));
      }),
    );
    const H = [...b];
    for (let Q = 0; Q < H.length; Q++)
      for (const X of G.get(H[Q]) ?? []) U.has(X) || (U.add(X), H.push(X));
    if (
      (h.forEach((Q, X) => {
        U.has(X) || v.set(X, Q);
      }),
      (j.status = "failed"),
      ik(e, a) ||
        [...a.values()].some((Q) => _O.test(Q.formula)) ||
        b.size > 1e3)
    )
      j.message =
        "Iteration supports at most 1,000 deterministic scalar cycle cells. Names, arrays, dynamic references and volatile formulas must be removed from this calculation.";
    else {
      let Q = new Map([...b].map((X) => [X, { type: "number", value: 0 }]));
      j.message = `No convergence within ${s.maxIterations} simultaneous passes from zero.`;
      for (let X = 0; X < s.maxIterations && !A(); X++) {
        I = k(v, Q);
        const J = new Map();
        let q = 0,
          W = true;
        for (const K of b) {
          const te = a.get(K),
            ae = I.getTypedValue(te.sheetId, te.address);
          if (
            (J.set(K, ae), ae.type !== "number" || !Number.isFinite(ae.value))
          ) {
            W = false;
            break;
          }
          q = Math.max(q, Math.abs(Number(ae.value) - Number(Q.get(K).value)));
        }
        if (
          ((j.passes = X + 1), (j.maxDelta = Number.isFinite(q) ? q : null), !W)
        ) {
          j.message =
            "The circular calculation produced a nonnumeric value or calculation error.";
          break;
        }
        if (((Q = J), q <= s.maxChange)) {
          ((j.status = "converged"),
            (j.message = `Converged from zero in ${X + 1} simultaneous passes; maximum change ${q}.`));
          break;
        }
      }
      j.status === "converged" && !w && b.forEach((X) => v.set(X, Q.get(X)));
    }
    ((j.status !== "converged" || w) &&
      b.forEach((Q) => v.set(Q, F5(w ?? "#ITER!"))),
      (I = k(v)),
      O());
  }
  (A(),
    w &&
      (l.forEach((U) => h.set(U, F5(w))),
      j.status === "converged" && (j.status = "failed"),
      b.size &&
        s.iterative &&
        (j.message = `Calculation stopped with ${w}; no intermediate result is current.`)),
    (y = false),
    u.end());
  const z = (U, G) => z.getTypedValue(U, G).value;
  Object.assign(z, I);
  const V = (U, G) => `${U}:${G.replace(/\$/g, "").toUpperCase()}`,
    D = new Map();
  if (F && t)
    for (const U of e.sheets)
      for (const G of t.calculator.getSpillEntries(U.id))
        l.has(`${U.id}:${G.anchor}`) && D.set(`${U.id}:${G.address}`, G.anchor);
  z.getTypedValue = (U, G) => {
    const H = V(U, G);
    if (h.has(H)) return { ...h.get(H) };
    const Q = e.sheets.find((W) => W.id === U)?.cells[
      G.replace(/\$/g, "").toUpperCase()
    ];
    if (
      Q &&
      Q.raw !== null &&
      (Q.literal || typeof Q.raw != "string" || !Q.raw.startsWith("="))
    )
      return {
        type:
          typeof Q.raw == "string"
            ? "text"
            : typeof Q.raw == "boolean"
              ? "boolean"
              : "number",
        value: Q.raw,
      };
    if (D.has(H)) return { ...ng };
    const X = I.getSpillAnchor(U, G);
    return X && h.get(`${U}:${X}`)?.type === "error"
      ? { ...h.get(`${U}:${X}`) }
      : I.getTypedValue(U, G);
  };
  const B = I.getSpillEntries;
  ((z.getSpillEntries = (U) => [
    ...B(U),
    ...[...D]
      .filter(([G]) => G.startsWith(`${U}:`))
      .map(([G, H]) => ({ address: G.slice(U.length + 1), anchor: H })),
  ]),
    (z.getOperationalErrors = () => [
      ...new Set([
        ...I.getOperationalErrors(),
        ...[...h.values()]
          .filter(
            (U) =>
              U.type === "error" &&
              [
                "#PENDING!",
                "#TIMEOUT!",
                "#ITER!",
                "#LIMIT!",
                "#BLOCKED!",
                "#CYCLE!",
                "#ERROR!",
              ].includes(String(U.value)),
          )
          .map((U) => U.value),
      ]),
    ]));
  const _ = (U, G) => [
      ...U.filter(
        (H) =>
          !l.has(`${H.sheetId}:${H.address}`) &&
          a.has(`${H.sheetId}:${H.address}`),
      ),
      ...G,
    ],
    C = t?.calculator.getReferenceDiagnostics() ?? [],
    T = t?.calculator.getNamedFormulaDiagnostics() ?? [],
    N = t?.calculator.getFinancialDiagnostics() ?? [];
  ((z.getFinancialDiagnostics = () =>
    structuredClone(_(N, I.getFinancialDiagnostics()))),
    (z.getReferenceDiagnostics = () => _(C, I.getReferenceDiagnostics())),
    (z.getNamedFormulaDiagnostics = () =>
      _(T, I.getNamedFormulaDiagnostics())));
  const R = new Map(
      t?.report.cells.map((U) => [`${U.sheetId}:${U.address}`, U]),
    ),
    L = new Map();
  for (const { sheetId: U, address: G, ...H } of z.getFinancialDiagnostics()) {
    const Q = `${U}:${G}`;
    L.set(Q, [...(L.get(Q) ?? []), H]);
  }
  const $ = [...a].map(([U, G]) => {
    var X;
    const H = h.get(U) ?? ng,
      Q = t?.values.get(U);
    return {
      ...G,
      ...(H.value !== "#PENDING!" && L.has(U) ? { financial: L.get(U) } : {}),
      result: { ...H },
      status:
        H.type === "error"
          ? H.value === "#PENDING!"
            ? "pending"
            : "error"
          : "current",
      ...(H.value === "#PENDING!" && Q && Q.value !== "#PENDING!"
        ? { lastKnown: { revision: t.report.revision, result: { ...Q } } }
        : {}),
      ...(H.value === "#PENDING!" &&
      Q?.value === "#PENDING!" &&
      (X = R.get(U)) != null &&
      X.lastKnown
        ? { lastKnown: R.get(U).lastKnown }
        : {}),
      dependencies: [...(c.get(U) ?? [])].sort(),
      evaluations: f.get(U) ?? 0,
    };
  });
  return {
    ...r,
    calculator: z,
    values: h,
    dependencies: c,
    report: {
      resources: u.report(),
      volatile: {
        functions: [
          ...new Set(
            [...a.values()]
              .flatMap((U) => Rg(U.formula))
              .concat((e.definedNames ?? []).flatMap((U) => Rg(U.refersTo))),
          ),
        ].sort(),
        policy: yE,
      },
      execution: {
        mode: "cooperative",
        state:
          w === "#TIMEOUT!"
            ? "timeout"
            : w
              ? "failed"
              : F
                ? "pending"
                : "complete",
        message: w
          ? (u.report().reason ?? `Calculation stopped with ${w}.`)
          : "Synchronous calculation uses cooperative limits.",
      },
      dateContext: d,
      workbookId: e.id,
      revision: e.revision,
      run: (t?.report.run ?? 0) + 1,
      settings: s,
      reason: p,
      formulaCount: a.size,
      evaluated: $.filter((U) => U.evaluations > 0).length,
      reused: a.size - l.size,
      pending: $.filter((U) => U.status === "pending").length,
      elapsedMs: Math.max(0, pd() - o),
      dependencyEdges: [...c.values()].reduce((U, G) => U + G.size, 0),
      cycles: g,
      iteration: j,
      cells: $,
      limits: {
        dependencyEdges: j5,
        cooperativeTimeBudgetMs: s.timeBudgetMs,
        hardInterrupt: false,
      },
    },
  };
}
function Gw(e) {
  const r = W1(e),
    t = tu.get(e);
  if (nm(t, e, r)) return { calculator: t.calculator, report: t.report };
  const n = dU(e, r, t ? undefined : ch.get(e), false);
  return (
    A1(n.report),
    tu.set(e, n),
    ch.delete(e),
    { calculator: n.calculator, report: n.report }
  );
}
function Vn(e, r = {}, t) {
  return !t &&
    !Object.keys(r).length &&
    (e.calculation || e.datePolicy || $ie(e) || eB(e) || nB(e))
    ? Gw(e).calculator
    : Cf(e, r, t);
}
function LB(e, r, t, n = {}) {
  var p;
  if (n.readOnly) throw Error("This workbook is read-only.");
  const o = e.sheets.find((m) => m.id === r);
  if (!o) throw Error("Worksheet no longer exists.");
  if ((p = o.protection) != null && p.enabled)
    throw Error("Unprotect the worksheet before changing conditional formats.");
  const a = FN(t),
    s = o.conditionalFormats ?? [],
    d = [];
  for (const m of new Set([...s.map((v) => v.id), ...a.map((v) => v.id)])) {
    const v = s.findIndex((w) => w.id === m),
      g = a.findIndex((w) => w.id === m),
      b = s[v],
      y = a[g];
    (JSON.stringify(b) !== JSON.stringify(y) || v !== g) &&
      d.push({
        id: m,
        ...(b ? { before: structuredClone(b), beforeIndex: v } : {}),
        ...(y ? { after: structuredClone(y), afterIndex: g } : {}),
      });
  }
  const u = Vn(e),
    l = Fw(o, u),
    c = Fw({ ...o, conditionalFormats: a }, u),
    h = new Set();
  for (const m of [...a, ...s]) {
    const v = ct(m.range);
    if ((h.add(`${ut(v.firstCol)}${v.firstRow + 1}`), h.size >= 50)) break;
  }
  if (h.size < 50) {
    for (const m of Object.keys(o.cells))
      if (
        ([...a, ...s].some((v) => n5(ct(v.range), _t(m))) && h.add(m),
        h.size >= 50)
      )
        break;
  }
  const f = [...h].map((m) => ({
    address: m,
    value: u(r, m),
    before: l(m),
    after: c(m),
  }));
  return {
    workbookId: e.id,
    revision: e.revision,
    snapshotGuard: JSON.stringify(e),
    sheetId: r,
    rules: a,
    changes: d,
    samples: f,
    errors: c.errors(),
    previewLimit: 50,
  };
}
function sU(e) {
  const r = W1(e),
    t = tu.get(e),
    n = dU(e, r, nm(t, e, r) ? t : undefined, true);
  return (
    A1(n.report),
    tu.set(e, n),
    ch.delete(e),
    { calculator: n.calculator, report: n.report }
  );
}
function Wie(e, r) {
  const t = Gw(e);
  if (t.report !== r)
    throw Error(
      "The workbook or calculation run changed. Review current calculation results before exporting evidence.",
    );
  const n = e.importedCalculationCaches,
    o = (n?.entries ?? []).map((a) => {
      const s = e.sheets.find((h) => h.id === a.sheetId),
        d = s?.cells[a.address],
        u = s ? t.calculator.getTypedValue(a.sheetId, a.address) : undefined,
        l =
          a.cache &&
          u &&
          a.cache.type === u.type &&
          (typeof a.cache.value == "number" && typeof u.value == "number"
            ? Math.abs(a.cache.value - u.value) <=
              1e-10 * Math.max(1, Math.abs(a.cache.value), Math.abs(u.value))
            : a.cache.value === u.value),
        c = s
          ? d != null && d.literal
            ? "neutralized-or-literal"
            : d?.raw !== a.formula
              ? "formula-changed"
              : u?.type === "error" && u.value === "#PENDING!"
                ? "pending"
                : a.cache
                  ? l
                    ? "matched"
                    : "different"
                  : "missing"
          : "source-sheet-removed";
      return {
        ...a,
        ...(u ? { current: u } : {}),
        status: c,
        verifiedByExcel: false,
      };
    });
  return JSON.stringify(
    {
      schemaVersion: 1,
      report: r,
      imported: n ? { ...n, entries: o } : null,
      boundaries: [
        "Imported caches are unverified source evidence, never current-value fallbacks.",
        "Matching caches do not prove Excel parity or source-file authenticity.",
        "Calculation timings and iteration counts are local runtime evidence.",
        r.limits.hardInterrupt
          ? "This calculation ran in a terminable worker. Timer delivery is subject to runtime scheduling."
          : "This synchronous calculation used cooperative time checks, not a hard worker interrupt.",
        "Working-data accounting is a bounded allocation allowance, not an exact heap/RSS measurement; runtime overhead is outside it.",
        "Diagnostic formula evaluation is separate from the committed revision calculation.",
      ],
    },
    null,
    2,
  );
}
function FR(e) {
  if (
    !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?$/.test(
      e,
    )
  )
    throw Error("Choose one explicit report range.");
  const r = e.replaceAll("$", "").toUpperCase().split(":"),
    t = ct(e);
  if (
    r[0] !== `${ut(t.firstCol)}${t.firstRow + 1}` ||
    (r[1] && r[1] !== `${ut(t.lastCol)}${t.lastRow + 1}`)
  )
    throw Error("Report ranges cannot be reversed.");
  if (
    t.lastCol - t.firstCol + 1 > 20 ||
    t.lastRow - t.firstRow + 1 > 500 ||
    (t.lastCol - t.firstCol + 1) * (t.lastRow - t.firstRow + 1) > 2e3
  )
    throw Error(
      "Reports support at most 500 rows, 20 columns and 2,000 cells.",
    );
  return `${ut(t.firstCol)}${t.firstRow + 1}:${ut(t.lastCol)}${t.lastRow + 1}`;
}
const PO = object({
    column: number().int().min(0).max(16383),
    type: _enum(["value", "text", "number", "date"]),
    operator: _enum([
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
    ]),
    value: union([
      string().max(1024),
      number().finite(),
      boolean(),
      _null(),
    ]).optional(),
  }).strict(),
  koe = object({
    range: Tc(40),
    columns: array(
      object({
        column: number().int().min(0).max(16383),
        values: array(union([Tc(32767), number().finite(), boolean(), _null()]))
          .max(1e3)
          .optional(),
        conditions: array(PO).max(128).optional(),
        join: _enum(["and", "or"]).optional(),
      }).strict(),
    ).max(20),
    advanced: array(array(PO).max(128)).max(128).optional(),
    caseSensitive: boolean().optional(),
    unique: boolean().optional(),
  }).strict(),
  kU = object({
    name: string().trim().min(1).max(100),
    kind: _enum(["range", "object"]),
    sheetId: Ap,
    range: Tc(40).transform(FR),
    objectId: Ap.optional(),
    mode: _enum(["snapshot", "linked"]),
    population: _enum(["full", "visible"]),
    filter: koe.optional(),
  }).strict(),
  Soe = object({
    address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
    value: union([Tc(32767), number().finite(), boolean(), _null()]),
    text: Tc(32767),
    style: unknown().optional(),
  }).strict(),
  Toe = object({
    engine: literal(CU),
    id: Ap,
    request: kU,
    workbookId: Ap,
    workbookName: Tc(200),
    revision: number().int().nonnegative(),
    sheetName: Tc(31),
    calculatedAt: AU,
    sourceHash: zE,
    includeHidden: boolean(),
    authorizedSheetIds: array(Ap).max(1e3),
    sourceCells: array(
      object({
        sheetId: Ap,
        address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
      }).strict(),
    ).max(4e3),
    rows: array(
      object({
        row: number().int().min(1).max(1048576),
        cells: array(Soe).max(20),
      }).strict(),
    ).max(500),
    columns: array(number().int().min(0).max(16383)).max(20),
    widths: array(number().finite().min(40).max(600)).max(20),
    excludedRows: array(number().int().min(1).max(1048576)).max(500),
    object: unknown().optional(),
    warnings: array(Tc(2e3)).max(50),
    hash: zE,
  }).strict(),
  Noe = object({
    id: Ap,
    fileName: Tc(255),
    fileBytes: number()
      .int()
      .min(1)
      .max(10 * 1024 * 1024),
    sourceSha256: zE,
    inspectedAt: AU,
    execution: literal("never"),
    items: array(
      object({
        path: Tc(4096),
        kind: Tc(100),
        bytes: number()
          .int()
          .nonnegative()
          .max(50 * 1024 * 1024),
        sha256: zE.optional(),
        execution: literal("never"),
        note: Tc(2e3),
      }).strict(),
    )
      .min(1)
      .max(1e3),
    warnings: array(Tc(2e3)).max(100),
  }).strict();
function BR(e) {
  const r = kU.parse(e);
  if ((r.kind === "object") != !!r.objectId)
    throw Error("Select an existing drawing for an object report.");
  if (r.filter && FR(r.filter.range) !== r.range)
    throw Error("The report range must match its filter range.");
  return r;
}

export { BR, Cf, FR, Gw, L5, LB, LQ, Noe, Toe, Vn, Wie, Xie, sU, uU, v0 };
