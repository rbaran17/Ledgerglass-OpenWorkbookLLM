import { $p, Gf, Pg, hr, lf, ts } from "./01-core-schemas-and-hashing.js";
import { Mt } from "./02-financial-functions.js";
import {
  $a,
  DR,
  Ku,
  Op,
  Vn,
  _t,
  cn,
  ct,
  gU,
  k1,
  mU,
  pf,
  rn,
  ut,
  y0,
  z1,
} from "./03-workbook-model.js";
import {
  G4,
  Hue,
  Qb,
  T7,
  Wue,
  Xue,
  ey,
  o3,
  que,
} from "./11-workbook-formats-dialog.js";

function m0(e, r) {
  if (!["precedents", "dependents"].includes(r.direction))
    throw Error("Choose precedents or dependents.");
  const t = G4(r.depth, 1, 8, "Inspection depth"),
    n = G4(r.limits?.maxFormulaCells, 5e3, 1e4, "Formula scan limit"),
    o = G4(r.limits?.maxNodes, 500, 1e3, "Result cell limit"),
    a = G4(r.limits?.maxReferences, 1e4, 5e4, "Reference scan limit");
  if (r.includeHidden !== undefined && typeof r.includeHidden != "boolean")
    throw Error("Include hidden must be boolean.");
  if (
    r.allowedSheetIds &&
    (!Array.isArray(r.allowedSheetIds) ||
      r.allowedSheetIds.some(
        (Q) => typeof Q != "string" || !e.sheets.some((X) => X.id === Q),
      ))
  )
    throw Error("The allowed worksheet context is invalid.");
  const s = new Map(e.sheets.map((Q) => [Q.id, Q])),
    d = new Set(r.allowedSheetIds ?? e.sheets.map((Q) => Q.id)),
    u = new Map(
      e.sheets.map((Q) => [
        Q.id,
        { rows: cn(Q, "row"), cols: cn(Q, "column") },
      ]),
    ),
    l = (Q, X) => {
      if (!d.has(Q.id)) return false;
      if (r.includeHidden) return true;
      const J = _t(X);
      return (
        !Q.hidden &&
        !u.get(Q.id).rows.some((q) => J.row >= q.start && J.row <= q.end) &&
        !u.get(Q.id).cols.some((q) => J.col >= q.start && J.col <= q.end)
      );
    },
    c = T7(r.address);
  if (c.includes(":")) throw Error("Choose one cell to inspect.");
  const h = s.get(r.sheetId);
  if (!h || !l(h, c))
    throw Error(
      "The selected cell is outside the visible, authorized inspection context.",
    );
  const f = Qb(h.id, c),
    p = [],
    m = [],
    v = new Map(),
    g = new Set();
  let b = 0,
    y = false,
    w = false,
    E = 0;
  const A = (Q, X, J, q) => {
      (g.add(Q),
        p.length < 200 &&
        !p.some(
          (W) =>
            W.from === Q && W.kind === X && W.token === q && W.message === J,
        )
          ? p.push({ from: Q, kind: X, message: J, ...(q ? { token: q } : {}) })
          : p.length >= 200 && (y = true));
    },
    k = [];
  for (const Q of e.sheets)
    for (const X of Object.entries(Q.cells).sort(([J], [q]) => {
      const W = _t(J),
        K = _t(q);
      return W.row - K.row || W.col - K.col;
    })) {
      if (!l(Q, X[0])) {
        E++;
        continue;
      }
      ey(X[1]) &&
        k.push({ key: Qb(Q.id, X[0]), sheet: Q, address: X[0], raw: X[1].raw });
    }
  k.sort((Q, X) => (Q.key === f ? -1 : X.key === f ? 1 : 0));
  const F = k.slice(0, n);
  k.length > F.length &&
    ((y = true),
    A(
      f,
      "limit",
      "The formula-cell scan limit was reached; additional dependents may exist.",
    ));
  for (const Q of F) {
    const X = [];
    v.set(Q.key, X);
    let J = 0;
    const q = (K, te, ae, oe) => {
        if (b >= a) {
          ((y = true),
            A(
              Q.key,
              "limit",
              "The workbook reference scan limit was reached.",
            ));
          return;
        }
        b++;
        const de = s.get(K);
        if (!de) {
          A(Q.key, "missing", "The referenced worksheet does not exist.", ae);
          return;
        }
        if (!d.has(K) || (de.hidden && !r.includeHidden)) {
          A(
            Q.key,
            "excluded",
            "A referenced worksheet is outside the selected inspection context; its cells and values were excluded.",
          );
          return;
        }
        try {
          const he = T7(te),
            Se = ct(he);
          (X.push({
            from: Q.key,
            token: ae,
            sheetId: K,
            range: he,
            kind: oe,
            cellCount:
              (Se.lastRow - Se.firstRow + 1) * (Se.lastCol - Se.firstCol + 1),
          }),
            !r.includeHidden &&
              (u
                .get(de.id)
                .rows.some(
                  (le) => le.start <= Se.lastRow && le.end >= Se.firstRow,
                ) ||
                u
                  .get(de.id)
                  .cols.some(
                    (le) => le.start <= Se.lastCol && le.end >= Se.firstCol,
                  )) &&
              A(
                Q.key,
                "excluded",
                "A referenced range intersects hidden rows or columns; hidden cell details are excluded.",
                ae,
              ));
        } catch {
          A(
            Q.key,
            "missing",
            "The reference is outside Excel's address bounds or is invalid.",
            ae,
          );
        }
      },
      W = (K, te, ae = [], oe) => {
        if (++J > 64 || ae.length > 16) {
          ((y = true),
            A(
              Q.key,
              "limit",
              "The defined-name expansion limit was reached.",
              oe,
            ));
          return;
        }
        let de = K.startsWith("=") ? 1 : 0;
        const he = new Map();
        if (/\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(K))
          try {
            Gf(Pg(K), (Ae, ye, Te, Pe) => {
              (Pe || (!Ae.qualifier && ye.has(lf(Ae.name)))) &&
                he.set(Ae.start, Ae.end);
            });
          } catch {
            A(
              Q.key,
              "unsupported",
              "Invalid LET/LAMBDA bindings prevent complete dependency inspection.",
            );
            return;
          }
        for (; de < K.length;) {
          if (he.has(de)) {
            de = he.get(de);
            continue;
          }
          if (b >= a) {
            ((y = true),
              A(
                Q.key,
                "limit",
                "The workbook reference scan limit was reached.",
              ));
            return;
          }
          const Ae = K.slice(de),
            ye = K[de];
          if (ye === '"') {
            de++;
            let Je = false;
            for (; de < K.length;)
              if (K[de++] === '"')
                if (K[de] === '"') de++;
                else {
                  Je = true;
                  break;
                }
            Je ||
              A(
                Q.key,
                "unsupported",
                "Unterminated formula text cannot be inspected completely.",
              );
            continue;
          }
          const Te =
            /^(?:'(?:[^']|'')*\[[^\]]+\](?:[^']|'')*'\s*!\s*[^+*/(),;\s]+|\[[^\]]+\][^!+*/(),;]*![^+*/(),;\s]+)/.exec(
              Ae,
            );
          if (Te) {
            (A(
              Q.key,
              "external",
              "External workbook references are not opened or fetched.",
              Te[0],
            ),
              (de += Te[0].length));
            continue;
          }
          const Pe = que.exec(Ae);
          if (Pe || /^'(?:[^']|'')*:(?:[^']|'')*'\s*!/.test(Ae)) {
            const Je = Pe?.[0] ?? /^[^+*/(),;]+/.exec(Ae)[0];
            (A(
              Q.key,
              "unsupported",
              "Three-dimensional worksheet references are not expanded.",
              Je,
            ),
              (de += Je.length));
            continue;
          }
          const qe = /^#[A-Z0-9/?!]+/i.exec(Ae);
          if (qe) {
            (A(
              Q.key,
              qe[0].toUpperCase() === "#REF!" ? "missing" : "unsupported",
              "The formula contains an error token.",
              qe[0],
            ),
              (de += qe[0].length));
            continue;
          }
          const ht = /^([A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(Ae);
          if (ht) {
            const Je = de;
            de += ht[0].length;
            let ze = 1;
            for (; de < K.length && ze;)
              (K[de] === "[" && ze++, K[de] === "]" && ze--, de++);
            const at = K.slice(Je, de),
              _e = !ze && z1(e, te, at, Q.address);
            _e
              ? q(_e.sheetId, _e.range, oe ?? at, oe ? "name" : "table")
              : A(
                  Q.key,
                  "unsupported",
                  "The table, field, or structured-reference form could not be resolved.",
                  at,
                );
            continue;
          }
          const Re =
            /^(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))\s*!\s*([A-Za-z_\\][A-Za-z0-9_.\\]*)/.exec(
              Ae,
            );
          if (Re && !/^[A-Za-z]{1,3}[1-9]\d*$/i.test(Re[3])) {
            const Je = (Re[1] ?? Re[2]).replace(/''/g, "'"),
              ze = e.sheets.find(
                (Ze) => Ze.name.toLowerCase() === Je.toLowerCase(),
              ),
              at = Re[0],
              _e = Re[3];
            if (!ze)
              A(
                Q.key,
                "missing",
                "The referenced worksheet does not exist.",
                at,
              );
            else if (!d.has(ze.id) || (ze.hidden && !r.includeHidden))
              A(
                Q.key,
                "excluded",
                "The defined name belongs to a worksheet outside the selected inspection context.",
              );
            else {
              const Ze =
                e.definedNames?.find(
                  (lt) =>
                    lt.scopeSheetId === ze.id &&
                    lt.name.toLowerCase() === _e.toLowerCase(),
                ) ??
                e.definedNames?.find(
                  (lt) =>
                    !lt.scopeSheetId &&
                    lt.name.toLowerCase() === _e.toLowerCase(),
                );
              if (!Ze)
                A(
                  Q.key,
                  "missing",
                  "The qualified defined name does not exist.",
                  at,
                );
              else {
                const lt = `${Ze.scopeSheetId ?? ""}:${Ze.name.toLowerCase()}`;
                ae.includes(lt)
                  ? A(
                      Q.key,
                      $p(Ze.refersTo) ? "dynamic" : "cycle",
                      "The reusable definition recurs; references are shown once.",
                      at,
                    )
                  : W(
                      Ze.refersTo,
                      Ze.scopeSheetId ?? ze.id,
                      [...ae, lt],
                      oe ?? at,
                    );
              }
            }
            de += at.length;
            continue;
          }
          const Me = /^([A-Za-z_\\][A-Za-z0-9_.\\]*)\s*\(/.exec(Ae);
          if (Me) {
            const Je = Me[1]
              .toUpperCase()
              .replace(/^_XLFN\./, "")
              .replace(/^_XLWS\./, "");
            if (["INDIRECT", "OFFSET"].includes(Je))
              A(
                Q.key,
                "dynamic",
                "Runtime reference selection is not resolved; static argument references are only potential dependencies.",
                Me[1],
              );
            else if (!Xue.has(Je)) {
              const ze =
                e.definedNames?.find(
                  (at) =>
                    at.scopeSheetId === te && at.name.toUpperCase() === Je,
                ) ??
                e.definedNames?.find(
                  (at) => !at.scopeSheetId && at.name.toUpperCase() === Je,
                );
              if (ze) {
                const at = `${ze.scopeSheetId ?? ""}:${ze.name.toLowerCase()}`;
                ae.includes(at)
                  ? A(
                      Q.key,
                      $p(ze.refersTo) ? "dynamic" : "cycle",
                      "Recursive reusable formula: static references are shown once; execution depth is bounded by calculation.",
                      Me[1],
                    )
                  : W(
                      ze.refersTo,
                      ze.scopeSheetId ?? te,
                      [...ae, at],
                      oe ?? Me[1],
                    );
              } else
                A(
                  Q.key,
                  "unsupported",
                  "The function is outside the local supported catalog; its static references may be incomplete.",
                  Me[1],
                );
            }
            de += Me[0].length;
            continue;
          }
          const ke = Hue.exec(Ae);
          if (ke && !/[A-Za-z0-9_.!:]/.test(Ae[ke[0].length] ?? "")) {
            const Je = ke[1]?.trim().replace(/!\s*$/, "").trim(),
              ze = Je?.replace(/^'|'$/g, "").replace(/''/g, "'"),
              at =
                ze === undefined
                  ? s.get(te)
                  : e.sheets.find(
                      (_e) => _e.name.toLowerCase() === ze.toLowerCase(),
                    );
            (at
              ? q(
                  at.id,
                  ke[2],
                  oe ?? ke[0],
                  oe ? "name" : ke[3] ? "spill-anchor" : "a1",
                )
              : A(
                  Q.key,
                  "missing",
                  "The referenced worksheet does not exist.",
                  ke[0],
                ),
              ke[3] &&
                A(
                  Q.key,
                  "dynamic",
                  "Only the spill anchor is included; runtime spill output dependencies are not expanded.",
                  ke[0],
                ),
              (de += ke[0].length));
            continue;
          }
          const je = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(Ae);
          if (je) {
            de += je[0].length;
            continue;
          }
          const $e = Wue.exec(Ae);
          if ($e) {
            const Je = $e[0],
              ze = Ae.slice(Je.length);
            if (/^\s*!/.test(ze)) {
              const _e = /^[^+*/(),;\s]+/.exec(Ae)[0];
              (A(
                Q.key,
                "unsupported",
                "Worksheet-qualified defined names are not resolved by this inspector.",
                _e,
              ),
                (de += _e.length));
              continue;
            }
            const at =
              e.definedNames?.find(
                (_e) =>
                  _e.scopeSheetId === te &&
                  _e.name.toLowerCase() === Je.toLowerCase(),
              ) ??
              e.definedNames?.find(
                (_e) =>
                  !_e.scopeSheetId &&
                  _e.name.toLowerCase() === Je.toLowerCase(),
              );
            if (at) {
              const _e = `${at.scopeSheetId ?? ""}:${at.name.toLowerCase()}`;
              ae.includes(_e)
                ? A(
                    Q.key,
                    "cycle",
                    "Defined names contain a circular reference.",
                    Je,
                  )
                : W(at.refersTo, at.scopeSheetId ?? te, [...ae, _e], oe ?? Je);
            } else
              ["TRUE", "FALSE"].includes(Je.toUpperCase()) ||
                A(
                  Q.key,
                  "missing",
                  "The identifier is not a defined name or supported reference.",
                  Je,
                );
            de += Je.length;
            continue;
          }
          ((ye === "'" ||
            ye === "!" ||
            ye === "[" ||
            ye === "]" ||
            ye === "@" ||
            ye === "|") &&
            A(
              Q.key,
              "unsupported",
              "The reference syntax is outside the inspected subset.",
              ye,
            ),
            de++);
        }
      };
    (W(Q.raw, Q.sheet.id), m.push(...X));
  }
  const I = new Map([[f, { sheet: h, address: c, depth: 0 }]]),
    O = new Map();
  let j = [f];
  const z = (Q, X, J, q, W) => {
    if (!l(Q, X)) return;
    const K = Qb(Q.id, X);
    if (!I.has(K)) {
      if (I.size - 1 >= o) {
        y = true;
        return;
      }
      I.set(K, { sheet: Q, address: X, depth: J });
    }
    O.set(JSON.stringify([q, W]), { from: q, to: W });
  };
  for (let Q = 1; Q <= t && j.length; Q++) {
    const X = new Set(I.keys());
    for (const J of j) {
      const q = I.get(J);
      if (r.direction === "precedents") {
        ey(q.sheet.cells[q.address]) &&
          !v.has(J) &&
          ((y = true),
          A(J, "limit", "This formula was outside the formula scan budget."));
        for (const W of v.get(J) ?? []) {
          const K = ct(W.range),
            te = s.get(W.sheetId);
          let ae = 0;
          e: for (let oe = K.firstRow; oe <= K.lastRow; oe++)
            for (let de = K.firstCol; de <= K.lastCol; de++) {
              if (++ae > o + 1 || I.size - 1 >= o) {
                y = true;
                break e;
              }
              const he = o3(oe, de);
              z(te, he, Q, J, Qb(te.id, he));
            }
        }
      } else {
        const W = _t(q.address);
        for (const K of F)
          (v.get(K.key) ?? []).some((te) => {
            const ae = ct(te.range);
            return (
              te.sheetId === q.sheet.id &&
              W.row >= ae.firstRow &&
              W.row <= ae.lastRow &&
              W.col >= ae.firstCol &&
              W.col <= ae.lastCol
            );
          }) && z(K.sheet, K.address, Q, K.key, J);
      }
    }
    j = [...I.keys()].filter((J) => !X.has(J));
  }
  for (const Q of j) {
    const X = I.get(Q);
    (r.direction === "precedents"
      ? ey(X.sheet.cells[X.address]) &&
        ((v.get(Q)?.length ?? 0) > 0 || !v.has(Q))
      : m.some((J) => {
          const q = _t(X.address),
            W = ct(J.range);
          return (
            J.sheetId === X.sheet.id &&
            q.row >= W.firstRow &&
            q.row <= W.lastRow &&
            q.col >= W.firstCol &&
            q.col <= W.lastCol
          );
        })) && (w = true);
  }
  const V = [...O.values()],
    D = [],
    B = new Map(),
    _ = (Q, X) => {
      if (B.get(Q) === 1) {
        const J = [...X.slice(X.indexOf(Q)), Q];
        D.length < 50 && D.push(J);
        return;
      }
      if (!(B.get(Q) === 2 || X.length > 1e3)) {
        B.set(Q, 1);
        for (const J of V) J.from === Q && _(J.to, [...X, Q]);
        B.set(Q, 2);
      }
    };
  for (const Q of I.keys()) _(Q, []);
  for (const Q of D) for (const X of Q) g.add(X);
  const C = (Q, X = new Set(), J = { cells: 0 }) => {
      if (X.has(Q) || X.size >= 32 || ++J.cells > 2e3 || g.has(Q)) return false;
      const [q, W] = JSON.parse(Q),
        K = s.get(q);
      if (!l(K, W)) return false;
      if (!ey(K.cells[W])) return true;
      const te = v.get(Q);
      if (!te) return false;
      const ae = new Set(X).add(Q);
      for (const oe of te) {
        if (oe.cellCount > 2e3) return false;
        const de = ct(oe.range);
        for (let he = de.firstRow; he <= de.lastRow; he++)
          for (let Se = de.firstCol; Se <= de.lastCol; Se++)
            if (!C(Qb(oe.sheetId, o3(he, Se)), ae, J)) return false;
      }
      return true;
    },
    T = {
      ...e,
      sheets: e.sheets
        .filter((Q) => d.has(Q.id) && (r.includeHidden || !Q.hidden))
        .map((Q) => ({
          ...Q,
          cells: Object.fromEntries(
            Object.entries(Q.cells).filter(([X]) => l(Q, X)),
          ),
        })),
    },
    N = Vn(T),
    R = (Q) => {
      const X = I.get(Q),
        J = X.sheet.cells[X.address],
        q = {
          key: Q,
          sheetId: X.sheet.id,
          sheetName: X.sheet.name,
          address: X.address,
          raw: J?.raw ?? null,
          depth: X.depth,
        };
      if (!ey(J))
        return {
          ...q,
          result: J?.raw ?? null,
          resultStatus: "literal",
        };
      if (!C(Q))
        return {
          ...q,
          result: null,
          resultStatus: "unavailable",
          resultNote:
            "Result withheld: the static dependency closure is cyclic, unresolved, excluded, or beyond the inspection budget.",
        };
      try {
        return {
          ...q,
          result: N(X.sheet.id, X.address),
          resultStatus: "calculated",
        };
      } catch {
        return {
          ...q,
          result: null,
          resultStatus: "unavailable",
          resultNote: "The local calculator could not evaluate this cell.",
        };
      }
    },
    L = [
      "Static potential dependencies include all formula branches; they do not describe the executed branch or evaluation steps.",
      "Blank referenced cells are valid local precedents. Runtime spill outputs are not expanded.",
      "Visibility preferences are local UI controls. The embedding host must supply authorized workbook data and enforce access independently.",
    ];
  return (
    E &&
      L.push(
        "Hidden or host-excluded populated cells were omitted from scanning and result details.",
      ),
    y &&
      L.push(
        "One or more budgets were reached. Additional references, cells, issues, or dependents may exist.",
      ),
    w &&
      L.push(
        "The requested traversal depth ended at cells with further potential relationships.",
      ),
    D.length &&
      L.push(
        "Cycles were found in the displayed static graph; other cycles may exist beyond the selected depth or scan limits.",
      ),
    p.length >= 200 && L.push("Only the first 200 distinct issues are listed."),
    {
      workbookId: e.id,
      revision: e.revision,
      direction: r.direction,
      depth: t,
      target: R(f),
      nodes: [...I.keys()].filter((Q) => Q !== f).map(R),
      edges: V,
      references: m.filter((Q) => I.has(Q.from)),
      issues: p,
      cycles: D,
      coverage: {
        formulaCellsAvailable: k.length,
        formulaCellsScanned: F.length,
        referencesScanned: b,
        excludedCells: E,
        truncated: y,
        depthLimited: w,
        complete: !y && !w && !p.length && !D.length && !E,
        notes: L,
      },
    }
  );
}
const Ik = Object.freeze({
    fuzzyCharacters: 512,
    comparisonWork: 1e7,
    matchPairs: 1e3,
  }),
  N7 = (e, r) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    Object.keys(e).every((t) => r.includes(t));
function K$(e, r, t) {
  if (
    e !== undefined &&
    (!N7(e, ["trim", "collapseWhitespace", "ignoreCase", "unicode"]) ||
      [e.trim, e.collapseWhitespace, e.ignoreCase].some(
        (n) => n !== undefined && typeof n != "boolean",
      ) ||
      (e.unicode !== undefined && !["NFC", "NFKC"].includes(e.unicode)))
  )
    throw Error("Invalid explicit key normalization options.");
  if (
    r !== undefined &&
    (!N7(r, ["key", "threshold"]) ||
      !Number.isInteger(r.key) ||
      r.key < 0 ||
      r.key >= t ||
      typeof r.threshold != "number" ||
      !Number.isFinite(r.threshold) ||
      r.threshold < 0 ||
      r.threshold > 1)
  )
    throw Error(
      "Choose a fuzzy key pair and a similarity threshold from 0 to 1.",
    );
  return {
    ...(e ? { normalization: { ...e } } : {}),
    ...(r ? { fuzzy: { ...r } } : {}),
  };
}
function Gue(e, r = {}) {
  if (typeof e != "string") return e;
  let t = r.unicode ? e.normalize(r.unicode) : e;
  return (
    r.collapseWhitespace && (t = t.replace(/\s+/gu, " ")),
    r.trim && (t = t.trim()),
    r.ignoreCase && (t = t.toLocaleLowerCase("en-US")),
    t
  );
}
const ty = (e) => JSON.stringify(e.map((r) => [typeof r, r])),
  Zue = (e) => e === null || e === "";
function Y$(e, r, t, n, o, a) {
  K$(o, a, t.length);
  const s = {
    policy: a
      ? `Normalized Unicode code-point Levenshtein similarity: 1 - edit distance / maximum text length. Fuzzy key ${a.key + 1}; inclusive threshold ${a.threshold}. Other keys compare typed values exactly after selected text normalization. All qualifying matches are retained; scores are not probabilities or Power Query equivalents.`
      : "Exact typed composite keys after explicitly selected text normalization. Blank keys never match; numbers and booleans are not coerced.",
    normalizedKeys: [],
    exceptions: [],
    matches: [],
    ambiguousLeftRows: 0,
    ambiguousRightRows: 0,
    normalizationCollisionGroups: { left: 0, right: 0 },
    comparisonWork: 0,
  };
  let d = 0;
  const u = (C, T, N) =>
      C.map((R) => {
        const L = T.map((X) => R.values[X]),
          $ = L.map((X) => Gue(X, o));
        if (
          ((d += $.reduce(
            (X, J) => X + (typeof J == "string" ? J.length : 0),
            0,
          )),
          d > 2e6)
        )
          throw Error("Normalized key text exceeds 2,000,000 characters.");
        ty(L) !== ty($) &&
          s.normalizedKeys.push({
            side: N,
            locator: R.locator,
            original: L,
            normalized: $,
          });
        const U = $.some(Zue),
          G = a ? $[a.key] : undefined,
          H = !!a && !U && typeof G != "string";
        H &&
          s.exceptions.push({
            side: N,
            locator: R.locator,
            reason:
              "Fuzzy matching requires text in the selected key. This record is unmatched; no numeric or Boolean coercion was applied.",
          });
        const Q = typeof G == "string" ? Array.from(G) : [];
        if (a && Q.length > Ik.fuzzyCharacters)
          throw Error(
            "Fuzzy keys are limited to 512 Unicode characters. Shorten the reviewed key or use exact matching.",
          );
        return {
          record: R,
          original: L,
          values: $,
          valid: !U && !H,
          isBlank: U,
          key: ty($),
          block: ty($.filter((X, J) => J !== a?.key)),
          chars: Q,
        };
      }),
    l = u(e, t, "left"),
    c = u(r, n, "right"),
    h = (C, T) => {
      const N = new Map();
      for (const R of C)
        R.isBlank || N.set(R.key, [...(N.get(R.key) ?? []), R]);
      return (
        (s.normalizationCollisionGroups[T] = [...N.values()].filter(
          (R) => new Set(R.map((L) => ty(L.original))).size > 1,
        ).length),
        N
      );
    },
    f = h(l, "left"),
    p = h(c, "right"),
    m = new Map(),
    v = new Map(),
    g = new Map(),
    b = new Map(),
    y = (C) => {
      if (((s.comparisonWork += C), s.comparisonWork > Ik.comparisonWork))
        throw Error(
          "Fuzzy comparison exceeds 10,000,000 work units. Narrow the populations or add an exact key before previewing.",
        );
    };
  function w(C, T) {
    if ((y(C.length + T.length + 1), C.join("") === T.join(""))) return 1;
    y(C.length * T.length);
    let N = Array.from({ length: T.length + 1 }, (R, L) => L);
    for (let R = 1; R <= C.length; R++) {
      const L = [R];
      for (let $ = 1; $ <= T.length; $++)
        L[$] = Math.min(
          N[$] + 1,
          L[$ - 1] + 1,
          N[$ - 1] + (C[R - 1] === T[$ - 1] ? 0 : 1),
        );
      N = L;
    }
    return 1 - N[T.length] / Math.max(C.length, T.length);
  }
  const E = new Map();
  for (const C of c)
    if (C.valid) {
      const T = a ? C.block : C.key;
      E.set(T, [...(E.get(T) ?? []), C]);
    }
  for (const C of l) {
    if (!C.valid) continue;
    const T = E.get(a ? C.block : C.key) ?? [];
    for (const N of T) {
      const R = a ? w(C.chars, N.chars) : 1;
      if (
        (g.set(C.record.row, Math.max(g.get(C.record.row) ?? -1, R)),
        b.set(N.record.row, Math.max(b.get(N.record.row) ?? -1, R)),
        a && R < a.threshold)
      )
        continue;
      const L = a
          ? `Text similarity ${R.toFixed(6)} meets threshold ${a.threshold}; other key pairs match exactly after selected normalization.`
          : "Exact typed keys match after selected normalization.",
        $ = m.get(C.record.row) ?? [];
      if (
        ($.push({ record: N.record, score: R, explanation: L }),
        m.set(C.record.row, $),
        v.set(N.record.row, (v.get(N.record.row) ?? 0) + 1),
        s.matches.push({
          left: C.record.locator,
          right: N.record.locator,
          score: R,
          explanation: L,
        }),
        s.matches.length > Ik.matchPairs)
      )
        throw Error(
          "Matching duplicate keys expand beyond 1,000 result rows or reviewable match pairs. Narrow or reconcile the source keys first; no pairs were truncated.",
        );
    }
  }
  ((s.ambiguousLeftRows = [...m.values()].filter((C) => C.length > 1).length),
    (s.ambiguousRightRows = [...v.values()].filter((C) => C > 1).length));
  const A = (C) => JSON.stringify(C.map((T) => String(T))),
    k = new Set(l.filter((C) => C.valid).map((C) => A(C.values))),
    F = new Set(c.filter((C) => C.valid).map((C) => A(C.values))),
    I = (C, T, N) =>
      C.isBlank
        ? "Blank join key; blanks do not match."
        : C.valid
          ? a
            ? T === undefined
              ? "No eligible text candidate with matching exact keys."
              : `Best eligible similarity ${T.toFixed(6)} is below threshold ${a.threshold}.`
            : N.has(A(C.values))
              ? "No exact typed composite-key match; equivalent displayed keys have mismatched types. No coercion was applied."
              : "No exact typed composite-key match."
          : "Non-text fuzzy key; no implicit type conversion.",
    O = l
      .filter((C) => !m.has(C.record.row))
      .map((C) => ({
        ...C.record,
        unmatchedReason: I(C, g.get(C.record.row), F),
      })),
    j = c
      .filter((C) => !v.has(C.record.row))
      .map((C) => ({
        ...C.record,
        unmatchedReason: I(C, b.get(C.record.row), k),
      }));
  for (const [C, T] of [
    ["left", O],
    ["right", j],
  ])
    for (const N of T)
      s.exceptions.some((R) => R.side === C && R.locator === N.locator) ||
        s.exceptions.push({
          side: C,
          locator: N.locator,
          reason: N.unmatchedReason,
        });
  const z = new Map(),
    V = new Map();
  for (const C of s.matches)
    (V.set(C.left, [...(V.get(C.left) ?? []), C.right]),
      z.set(C.right, [...(z.get(C.right) ?? []), C.left]));
  const D = new Set();
  let B = 0,
    _ = 0;
  for (const C of V.keys())
    if (!D.has(C)) {
      const T = new Set(),
        N = new Set(),
        R = [C];
      let L = 0;
      for (; R.length;) {
        const $ = R.pop();
        if (D.has($)) continue;
        (D.add($), T.add($));
        const U = V.get($);
        L += U.length;
        for (const G of U) {
          N.add(G);
          for (const H of z.get(G)) D.has(H) || R.push(H);
        }
      }
      T.size > 1 &&
        N.size > 1 &&
        (B++, (_ += Math.max(0, L - Math.max(T.size, N.size))));
    }
  return {
    byLeft: m,
    unmatchedLeft: O,
    unmatchedRight: j,
    diagnostics: s,
    duplicateLeftKeys: [...f.values()].filter((C) => C.length > 1).length,
    duplicateRightKeys: [...p.values()].filter((C) => C.length > 1).length,
    matchedLeftRows: m.size,
    matchedRightRows: v.size,
    matchPairs: s.matches.length,
    manyToManyKeys: B,
    expandedMatchRows: _,
  };
}
const Wu = {
  sourceCells: 2e3,
  sourceRows: 500,
  sourceColumns: 50,
  outputRows: 1e3,
  outputCells: 2e4,
  keyColumns: 8,
  sourceFormulas: 50,
  scannedFormulas: 1e3,
  dependencyCells: 2e3,
  headerCharacters: 128,
  textCharacters: 32767,
  totalTextCharacters: 1e6,
};
function Kn(e) {
  throw Error(e);
}
const dh = (e, r) => `${ut(r)}${e + 1}`,
  _k = (e, r) => JSON.stringify([e, r]),
  Wv = (e) =>
    !!e && !e.literal && typeof e.raw == "string" && e.raw.startsWith("="),
  pT = (e, r) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    Object.keys(e).every((t) => r.includes(t));
function mT(e) {
  (!pT(e, ["sheetId", "range"]) ||
    typeof e.sheetId != "string" ||
    typeof e.range != "string" ||
    !/^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/.test(
      e.range,
    )) &&
    Kn(
      "Choose a worksheet and one explicit headed cell rectangle for each source.",
    );
  const r = e.range.replace(/\$/g, "").toUpperCase(),
    t = r.split(":"),
    n = _t(t[0]),
    o = _t(t[1] ?? t[0]);
  (n.row > o.row ||
    n.col > o.col ||
    o.row >= Mt.maxRows ||
    o.col >= Mt.maxCols) &&
    Kn("A source rectangle is reversed or outside Excel's address bounds.");
  const a = o.row - n.row + 1,
    s = o.col - n.col + 1;
  return (
    (a > Wu.sourceRows + 1 || s > Wu.sourceColumns || a * s > Wu.sourceCells) &&
      Kn(
        "Each source supports at most 500 data rows, 50 columns and 2,000 cells including headers.",
      ),
    { sheetId: e.sheetId, range: `${dh(n.row, n.col)}:${dh(o.row, o.col)}` }
  );
}
function Kue(e, r, t) {
  return (
    !e.hidden &&
    !e.veryHidden &&
    !cn(e, "row").some((n) => r >= n.start && r <= n.end) &&
    !cn(e, "column").some((n) => t >= n.start && t <= n.end)
  );
}
function q1(e, r) {
  let t = 0;
  const n = (o, a, s = "A1") => {
    if (!o) return;
    ((t += o.length),
      t > Wu.totalTextCharacters &&
        Kn(
          "Output-name safety inspection exceeds 1,000,000 formula characters. Use a smaller reviewed workbook.",
        ));
    const d = o.replace(/"(?:[^"]|"")*"/g, "");
    let u = "",
      l = 0;
    for (; l < d.length;) {
      if (d[l] === "'") {
        const f = l++;
        for (; l < d.length;)
          if (d[l++] === "'")
            if (d[l] === "'") l++;
            else break;
        u += d.slice(f, l);
        continue;
      }
      const c = l,
        h = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(d.slice(l));
      if (h) {
        l += h[0].length;
        let f = 1;
        for (; l < d.length && f;)
          (d[l] === "[" && f++, d[l] === "]" && f--, l++);
        const p = d.slice(c, l);
        u +=
          !f && !/^[A-Za-z0-9_!]/.test(d.slice(l)) && z1(e, a, p, s) ? "0" : p;
      } else u += d[l++];
    }
    for (const c of u.matchAll(
      /('(?:[^']|'')*'|[A-Za-z_.\d\u007F-\uFFFF]+)\s*!/g,
    )) {
      const h = c[1];
      (h.startsWith("'")
        ? h.slice(1, -1).replace(/''/g, "'")
        : h
      ).toLowerCase() === r.toLowerCase() &&
        Kn(
          "The output name would reconnect an existing unresolved worksheet reference. Choose a different name.",
        );
    }
  };
  for (const o of e.sheets) {
    for (const [a, s] of Object.entries(o.cells))
      Wv(s) && n(String(s.raw), o.id, a);
    for (const a of o.conditionalFormats ?? [])
      a.kind === "formula" && n(a.formula, o.id, a.range.split(":")[0]);
    for (const a of o.tables ?? [])
      for (const [s, d] of a.columns.entries()) {
        const u = ct(a.range);
        n(d.calculatedFormula, o.id, dh(u.firstRow + 1, u.firstCol + s));
      }
  }
  for (const o of e.definedNames ?? [])
    n(o.refersTo, o.scopeSheetId ?? e.sheets[0].id);
}
function J$(e, r, t = {}) {
  (!Array.isArray(r) || !r.length || r.length > 8) &&
    Kn("Read one to eight explicit source rectangles.");
  const n = r.map(mT);
  (t.includeHidden !== undefined &&
    typeof t.includeHidden != "boolean" &&
    Kn("Include hidden must be boolean."),
    t.allowedSheetIds !== undefined &&
      (!Array.isArray(t.allowedSheetIds) ||
        t.allowedSheetIds.some(
          (y) => typeof y != "string" || !e.sheets.some((w) => w.id === y),
        )) &&
      Kn("The host scope contains an unknown worksheet."));
  const o = {
      includeHidden: !!t.includeHidden,
      allowedSheetIds: [
        ...new Set(t.allowedSheetIds ?? e.sheets.map((y) => y.id)),
      ].sort(),
    },
    a = n.map((y) => {
      var A;
      const w = e.sheets.find((k) => k.id === y.sheetId);
      (!w || !o.allowedSheetIds.includes(w.id)) &&
        Kn("A source is outside the authorized host context.");
      const E = ct(y.range);
      if (!o.includeHidden)
        for (let k = E.firstRow; k <= E.lastRow; k++)
          for (let F = E.firstCol; F <= E.lastCol; F++)
            Kue(w, k, F) ||
              Kn(
                "A source includes hidden cells. Explicitly include hidden records before previewing.",
              );
      return (
        (A = w.merges) != null &&
          A.some((k) => {
            const F = ct(k);
            return (
              F.firstRow <= E.lastRow &&
              F.lastRow >= E.firstRow &&
              F.firstCol <= E.lastCol &&
              F.lastCol >= E.firstCol
            );
          }) &&
          Kn(
            "Merged source cells are ambiguous; use an unmerged headed population.",
          ),
        w
      );
    }),
    s = new Set(),
    d = new Set(),
    u = [];
  for (const [y, w] of n.entries()) {
    const E = a[y],
      A = ct(w.range);
    for (let k = A.firstRow; k <= A.lastRow; k++)
      for (let F = A.firstCol; F <= A.lastCol; F++) {
        const I = dh(k, F),
          O = _k(E.id, I);
        (s.add(O),
          Wv(E.cells[I]) &&
            !d.has(O) &&
            (d.add(O), u.push({ sheet: E, address: I })));
      }
  }
  u.length > Wu.sourceFormulas &&
    Kn(
      "Source capture supports at most 50 formula source cells. Use a reviewed values snapshot for a larger calculated population.",
    );
  for (const y of new Set(a)) {
    const w = [];
    for (const A of s) {
      const [k, F] = JSON.parse(A);
      k === y.id && (!y.cells[F] || y.cells[F].raw === null) && w.push(_t(F));
    }
    if (!w.length) continue;
    const E = Object.entries(y.cells).filter(
      ([A, k]) => Wv(k) && !d.has(_k(y.id, A)),
    );
    E.length > Wu.scannedFormulas &&
      Kn(
        "Blank source cells require a bounded spill-ambiguity check; this sheet exceeds 1,000 possible formula anchors.",
      );
    for (const [A] of E) {
      const k = _t(A);
      w.some(
        (F) =>
          F.row >= k.row &&
          F.col >= k.col &&
          (F.row - k.row + 1) * (F.col - k.col + 1) <= 1e4,
      ) &&
        Kn(
          "A blank source cell could be a spill child of a formula outside the source. Use a reviewed values snapshot or include and validate the scalar source formula.",
        );
    }
  }
  const l = new Set([
      ...y0,
      "ISFORMULA",
      "INDEX",
      "CHOOSE",
      "INDIRECT",
      "OFFSET",
      "LET",
      "LAMBDA",
      "RAND",
      "RANDBETWEEN",
      "NOW",
      "TODAY",
      "CELL",
      "INFO",
      "RTD",
      "WEBSERVICE",
      "STOCKHISTORY",
    ]),
    c = {
      ...e,
      sheets: e.sheets.map((y) => (y.veryHidden ? { ...y, hidden: true } : y)),
    },
    h = new Set();
  for (const y of u) {
    const w = m0(c, {
        sheetId: y.sheet.id,
        address: y.address,
        direction: "precedents",
        depth: 8,
        includeHidden: o.includeHidden,
        allowedSheetIds: o.allowedSheetIds,
        limits: { maxFormulaCells: 1e3, maxNodes: 500, maxReferences: 2e3 },
      }),
      E = [w.target, ...w.nodes],
      A = new Set(E.map((k) => k.key));
    (w.coverage.truncated ||
      w.coverage.depthLimited ||
      w.cycles.length ||
      w.issues.some((k) => A.has(k.from))) &&
      Kn(
        "A source formula has incomplete, cyclic, hidden, unauthorized or unsupported dependencies.",
      );
    for (const k of E) {
      (s.add(k.key),
        h.add(k.key),
        h.size > Wu.dependencyCells &&
          Kn("Combined source formula dependencies exceed 2,000 cells."));
      const F = e.sheets.find((O) => O.id === k.sheetId),
        I = F.cells[k.address];
      if (
        ((!I || I.raw === null) &&
          Kn(
            "Formula dependencies must be populated static cells; blank or possible spill-child dependencies are outside this subset.",
          ),
        Wv(I))
      ) {
        const O = pf(String(I.raw), e, F.id).replace(
          /"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g,
          "",
        );
        ([...O.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)]
          .map((z) =>
            z[1]
              .toUpperCase()
              .replace(/^_XLFN\./, "")
              .replace(/^_XLWS\./, ""),
          )
          .some((z) => l.has(z)) ||
          /[{}#]/.test(O)) &&
          Kn(
            "Source formulas must have deterministic scalar dependencies; volatile, dynamic and array formulas are excluded.",
          );
      }
    }
  }
  const f = {
      ...e,
      sheets: e.sheets
        .filter((y) => o.allowedSheetIds.includes(y.id))
        .map((y) => ({
          ...y,
          cells: Object.fromEntries(
            Object.entries(y.cells).filter(([w]) => s.has(_k(y.id, w))),
          ),
        })),
    },
    p = Vn(f),
    m =
      /^#(?:DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|CYCLE!|LIMIT!|BLOCKED!|ERROR!)$/;
  let v = 0;
  const g = n.map((y, w) => {
    const E = a[w],
      A = ct(y.range),
      k = [];
    for (let F = A.firstRow; F <= A.lastRow; F++) {
      const I = [];
      for (let O = A.firstCol; O <= A.lastCol; O++) {
        const j = dh(F, O),
          z = E.cells[j],
          V = Wv(z) ? p(E.id, j) : (z?.raw ?? null);
        (((Wv(z) && typeof V == "string" && m.test(V)) ||
          (typeof V == "number" && !Number.isFinite(V)) ||
          (typeof V == "string" && V.length > Wu.textCharacters)) &&
          Kn(
            "A source contains a calculation error, nonfinite number or oversized text. Repair it before capturing source values.",
          ),
          (v += typeof V == "string" ? V.length : 0),
          v > Wu.totalTextCharacters &&
            Kn("Combined source text exceeds 1,000,000 characters."),
          I.push(V));
      }
      k.push(I);
    }
    return { source: y, sheetName: E.name, values: k };
  });
  for (const y of h) {
    const [w, E] = JSON.parse(y);
    p.getSpillRange(w, E) &&
      Kn(
        "Source formula dependencies produced an array; use a reviewed scalar population.",
      );
  }
  p.getOperationalErrors().length &&
    Kn(
      "A source encountered an operational calculation failure, including an error hidden by a formula handler.",
    );
  const b = [...s].sort().map((y) => {
    const [w, E] = JSON.parse(y);
    return {
      sheetId: w,
      address: E,
      cell: e.sheets.find((k) => k.id === w)?.cells[E] ?? null,
    };
  });
  return { grids: g, dependencyCells: b };
}
function X1(e, r, t = {}) {
  const { grids, dependencyCells } = J$(e, r, t);
  return {
    snapshots: grids.map(({ source, sheetName, values }) => {
      const l = ct(source.range);
      values[0].some(
        (h) =>
          typeof h != "string" || !h.trim() || h.length > Wu.headerCharacters,
      ) &&
        Kn(
          "Every source column needs a nonblank text header of at most 128 characters.",
        );
      const c = values[0];
      return (
        new Set(c.map((h) => h.toLowerCase())).size !== c.length &&
          Kn("Source headers must be unique, including case variants."),
        {
          source,
          sheetName,
          headers: c,
          records: values.slice(1).map((h, f) => ({
            row: l.firstRow + f + 2,
            locator: `'${sheetName.replace(/'/g, "''")}'!${dh(l.firstRow + f + 1, l.firstCol)}:${dh(l.firstRow + f + 1, l.lastCol)}`,
            values: h,
          })),
        }
      );
    }),
    dependencyCells,
  };
}
function Q$(e, r, t = {}) {
  (t.readOnly && Kn("This workbook is read-only."),
    $a(e),
    (!pT(r, [
      "left",
      "right",
      "mode",
      "keys",
      "outputSheetName",
      "normalization",
      "fuzzy",
    ]) ||
      ![
        "append",
        "inner",
        "left",
        "right",
        "full",
        "leftAnti",
        "rightAnti",
      ].includes(r.mode)) &&
      Kn("Choose append or a supported join operation."),
    (typeof r.outputSheetName != "string" ||
      r.outputSheetName !== r.outputSheetName.trim() ||
      !r.outputSheetName ||
      r.outputSheetName.length > 31 ||
      /[\\/\[\]*?:\x00-\x1f]/.test(r.outputSheetName) ||
      /^'|'$/.test(r.outputSheetName)) &&
      Kn(
        "Output sheet name must be a valid, unpadded name of 1–31 characters.",
      ),
    e.sheets.some(
      (z) => z.name.toLowerCase() === r.outputSheetName.toLowerCase(),
    ) && Kn("The output sheet name already exists. Choose a new sheet name."),
    q1(e, r.outputSheetName),
    t.includeHidden !== undefined &&
      typeof t.includeHidden != "boolean" &&
      Kn("Include hidden must be boolean."),
    t.allowedSheetIds !== undefined &&
      (!Array.isArray(t.allowedSheetIds) ||
        t.allowedSheetIds.some(
          (z) => typeof z != "string" || !e.sheets.some((V) => V.id === z),
        )) &&
      Kn("The host scope contains an unknown worksheet."));
  const n = {
      includeHidden: !!t.includeHidden,
      allowedSheetIds: [
        ...new Set(t.allowedSheetIds ?? e.sheets.map((z) => z.id)),
      ].sort(),
    },
    o = {
      left: mT(r.left),
      right: mT(r.right),
      mode: r.mode,
      outputSheetName: r.outputSheetName,
    };
  o.mode === "append"
    ? ((r.normalization !== undefined || r.fuzzy !== undefined) &&
        Kn(
          "Append aligns headers; key normalization and fuzzy matching apply only to joins.",
        ),
      r.keys !== undefined &&
        (!Array.isArray(r.keys) || r.keys.length) &&
        Kn("Append matches headers; it does not accept join keys."))
    : ((!Array.isArray(r.keys) ||
        !r.keys.length ||
        r.keys.length > Wu.keyColumns ||
        r.keys.some(
          (z) =>
            !pT(z, ["left", "right"]) ||
            typeof z.left != "string" ||
            typeof z.right != "string",
        )) &&
        Kn("Choose from one to eight pairs of join headers."),
      (o.keys = r.keys.map((z) => ({ ...z }))),
      Object.assign(o, K$(r.normalization, r.fuzzy, o.keys.length)),
      (new Set(o.keys.map((z) => z.left)).size !== o.keys.length ||
        new Set(o.keys.map((z) => z.right)).size !== o.keys.length) &&
        Kn("A join header cannot be repeated in the key pairs."));
  const { snapshots } = X1(e, [o.left, o.right], t),
    [s, d] = snapshots,
    u = [],
    l = [],
    c = [],
    h = [];
  let f,
    p,
    m = 0,
    v = 0,
    g = 0,
    b = 0,
    y = 0,
    w = 0,
    E = 0;
  if (o.mode === "append") {
    ((s.headers.length !== d.headers.length ||
      s.headers.some((V) => !d.headers.includes(V))) &&
      Kn(
        "Append requires exactly matching header names on both sources; column order may differ.",
      ),
      (p = [...s.headers, "Source left row", "Source right row"]));
    const z = s.headers.map((V) => d.headers.indexOf(V));
    (s.records.forEach((V) =>
      u.push({ values: [...V.values], left: V.locator, right: null }),
    ),
      d.records.forEach((V) =>
        u.push({
          values: z.map((D) => V.values[D]),
          left: null,
          right: V.locator,
        }),
      ));
  } else {
    p = [
      ...s.headers.map((_) => `Left: ${_}`),
      ...d.headers.map((_) => `Right: ${_}`),
      "Source left row",
      "Source right row",
    ];
    const z = o.keys.map((_) => s.headers.indexOf(_.left)),
      V = o.keys.map((_) => d.headers.indexOf(_.right));
    (z.includes(-1) || V.includes(-1)) &&
      Kn("Every join key must name an existing source header exactly.");
    const D = (_, C) =>
      [
        ...new Set(
          _.map((T) =>
            T.values[C] === null || T.values[C] === ""
              ? "blank"
              : typeof T.values[C],
          ),
        ),
      ].sort();
    o.keys.forEach((_, C) =>
      h.push({
        leftHeader: _.left,
        rightHeader: _.right,
        leftTypes: D(s.records, z[C]),
        rightTypes: D(d.records, V[C]),
      }),
    );
    const B = Y$(s.records, d.records, z, V, o.normalization, o.fuzzy);
    ((f = B.diagnostics),
      ({
        matchedLeftRows: m,
        matchedRightRows: v,
        matchPairs: g,
        duplicateLeftKeys: b,
        duplicateRightKeys: y,
        manyToManyKeys: w,
        expandedMatchRows: E,
      } = B),
      l.push(...B.unmatchedLeft),
      c.push(...B.unmatchedRight));
    for (const _ of s.records) {
      const C = B.byLeft.get(_.row);
      if (C) {
        if (o.mode !== "leftAnti" && o.mode !== "rightAnti")
          for (const T of C)
            u.push({
              values: [..._.values, ...T.record.values],
              left: _.locator,
              right: T.record.locator,
              ...(o.fuzzy || o.normalization
                ? { match: { score: T.score, explanation: T.explanation } }
                : {}),
            });
      } else
        ["left", "full", "leftAnti"].includes(o.mode) &&
          u.push({
            values: [..._.values, ...d.headers.map(() => null)],
            left: _.locator,
            right: null,
          });
    }
    if (["right", "full", "rightAnti"].includes(o.mode))
      for (const _ of c)
        u.push({
          values: [...s.headers.map(() => null), ..._.values],
          left: null,
          right: _.locator,
        });
  }
  (new Set(p.map((z) => z.toLowerCase())).size !== p.length &&
    Kn("A source header conflicts with a generated provenance column name."),
    (u.length > Wu.outputRows || (u.length + 1) * p.length > Wu.outputCells) &&
      Kn(
        "The output exceeds 1,000 data rows or 20,000 cells including headers and provenance.",
      ),
    u.reduce(
      (z, V) =>
        z +
        V.values.reduce((D, B) => D + (typeof B == "string" ? B.length : 0), 0),
      p.join("").length,
    ) > Wu.totalTextCharacters &&
      Kn("Expanded output text exceeds 1,000,000 characters."),
    e.sheets.reduce((z, V) => z + Object.keys(V.cells).length, 0) +
      (u.length + 1) * p.length >
      Mt.maxCells &&
      Kn("The new output would exceed the workbook's stored-cell budget."));
  const k = Ku(o.outputSheetName),
    F = JSON.stringify(e),
    I = new Set(F.match(/"population-combine-[1-9]\d*"/g) ?? []);
  let O = 1;
  for (; I.has(JSON.stringify(`population-combine-${O}`));) O++;
  ((k.id = `population-combine-${O}`),
    (k.frozenRows = 1),
    (k.columnWidths = Object.fromEntries(
      p.map((z, V) => [V, Math.min(240, Math.max(110, z.length * 7 + 20))]),
    )),
    p.forEach((z, V) => {
      k.cells[dh(0, V)] = {
        raw: z,
        literal: z.startsWith("="),
        style: { bold: true, fill: "#E2F0D9" },
      };
    }),
    (k.cells.A1.note = `Reviewed ${o.mode} values snapshot; source workbook ${e.id}, revision ${e.revision}. Left: '${s.sheetName}'!${s.source.range} (${s.records.length} records); right: '${d.sheetName}'!${d.source.range} (${d.records.length} records). Keys: ${o.keys ? JSON.stringify(o.keys) : "exact matching headers"}. ${f?.policy ?? "Append by exact headers."} Normalization: ${JSON.stringify(o.normalization ?? {})}; blank keys do not match. Matched left/right: ${m}/${v}; match pairs: ${g}; unmatched left/right: ${l.length}/${c.length}; output records: ${u.length}. Duplicate key groups left/right: ${b}/${y}; many-to-many groups: ${w}; extra many-to-many pairs: ${E}. Hidden sources explicitly included: ${n.includeHidden}. Filter views do not exclude records. This snapshot does not automatically refresh.`),
    u.forEach((z, V) =>
      [...z.values, z.left, z.right].forEach((D, B) => {
        k.cells[dh(V + 1, B)] = {
          raw: D,
          ...(typeof D == "string" && D.startsWith("=")
            ? { literal: true }
            : {}),
          note: `Combined values snapshot. Left source: ${z.left ?? "unmatched / not applicable"}. Right source: ${z.right ?? "unmatched / not applicable"}. Workbook ${e.id}; revision ${e.revision}. ${z.match?.explanation ?? "No matched pair / append source record."}`,
        };
      }),
    ));
  const j = structuredClone(e);
  return (
    j.sheets.push(k),
    rn(j),
    {
      snapshot: F,
      request: o,
      context: n,
      left: s,
      right: d,
      headers: p,
      rows: u,
      unmatchedLeft: l,
      unmatchedRight: c,
      keyProfiles: h,
      ...(f ? { matching: f } : {}),
      counts: {
        leftRows: s.records.length,
        rightRows: d.records.length,
        matchedLeftRows: m,
        matchedRightRows: v,
        matchPairs: g,
        duplicateLeftKeys: b,
        duplicateRightKeys: y,
        manyToManyKeys: w,
        expandedMatchRows: E,
        outputRows: u.length,
      },
      outputSheetId: k.id,
      outputRange: `A1:${dh(u.length, p.length - 1)}`,
      next: j,
    }
  );
}
function Yue(e, r, t = {}) {
  (t.readOnly && Kn("This workbook is read-only."),
    r.snapshot !== JSON.stringify(e) &&
      Kn(
        "The workbook changed after preview. Preview and review the populations again.",
      ));
  const n = Q$(e, r.request, t);
  return (
    JSON.stringify(n) !== JSON.stringify(r) &&
      Kn("The preview or host context changed. Preview and review again."),
    n.next
  );
}
function vI(e, r) {
  if (
    (r.readOnly !== undefined && typeof r.readOnly != "boolean") ||
    (r.includeHidden !== undefined && typeof r.includeHidden != "boolean")
  )
    throw Error("Host read-only and hidden-source policies must be boolean.");
  const t = r.allowedSheetIds ?? e.sheets.map((o) => o.id);
  if (
    !Array.isArray(t) ||
    t.some((o) => typeof o != "string" || !e.sheets.some((a) => a.id === o))
  )
    throw Error("Invalid authorized worksheet context.");
  const n = r.approvedWorkbooks ?? [];
  if (
    !Array.isArray(n) ||
    n.length > 8 ||
    new Set(n.map((o) => o.workbook.id)).size !== n.length ||
    n.some(
      (o) =>
        o.workbook.id === e.id ||
        !Array.isArray(o.allowedSheetIds) ||
        o.allowedSheetIds.some(
          (a) => !o.workbook.sheets.some((s) => s.id === a),
        ) ||
        (o.includeHidden !== undefined && typeof o.includeHidden != "boolean"),
    )
  )
    throw Error(
      "Approved workbook identities and worksheet permissions must be explicit, distinct and valid (at most 8 workbooks).",
    );
  return { allowed: t, external: n };
}
function ez(e, r, t) {
  const n = vI(e, t),
    o = n.external.find((p) => p.workbook.id === r.workbookId),
    a = r.workbookId === e.id,
    s = a ? e : o?.workbook;
  if (!s)
    throw Error(
      "Source workbook is missing or has not been approved. Attach and approve the exact workbook identity before refreshing.",
    );
  const d = a ? n.allowed : o.allowedSheetIds,
    u = s.sheets.find((p) => p.id === r.sheetId);
  if (!u || !d.includes(r.sheetId))
    throw Error(
      "Source worksheet is missing or outside the approved worksheet scope.",
    );
  let l = r.range;
  if (r.tableId) {
    const p = u.tables?.find((v) => v.id === r.tableId);
    if (!p) throw Error("Source table is missing. Review the source mapping.");
    const m = ct(p.range);
    l = `${ut(m.firstCol)}${m.firstRow + 1}:${ut(m.lastCol)}${m.lastRow + 1 - (p.showTotals ? 1 : 0)}`;
  }
  const c = X1(s, [{ sheetId: u.id, range: l }], {
      allowedSheetIds: d,
      includeHidden: a
        ? t.includeHidden
        : !!t.includeHidden && !!o.includeHidden,
    }),
    h = c.snapshots[0];
  return {
    book: s,
    sheet: u,
    snapshot: h,
    fingerprint: hr({
      workbookId: s.id,
      workbookName: s.name,
      source: r,
      snapshot: h,
      dependencies: c.dependencyCells,
      datePolicy: ts(s.datePolicy),
    }),
  };
}
function Jue(e, r, t = {}) {
  const n = mU(r);
  vI(e, t);
  const o = {
      sources: [],
      contributions: [],
      subtotals: [],
      exceptions: [],
      grandTotal: null,
      sourceTotal: null,
      difference: null,
      complete: false,
    },
    a = new Set(),
    s = (d, u, l, c, h = "error") =>
      o.exceptions.push({
        sourceId: d,
        code: u,
        message: l,
        ...(c ? { row: c } : {}),
        severity: h,
      });
  for (const d of n.sources)
    try {
      const u = ez(e, d, t),
        { snapshot } = u,
        c = ct(snapshot.source.range),
        h = (y) => {
          const w = snapshot.headers.indexOf(y);
          if (w < 0)
            throw Error(
              `Required field “${y}” is missing. Review the source schema.`,
            );
          return w;
        },
        f = d.categories.map(h),
        p = h(d.amountField),
        m = (y, w) =>
          d[y].kind === "constant" ? d[y].value : w[h(d[y].field)];
      for (const y of ["period", "currency", "unit"])
        d[y].kind === "field" && h(d[y].field);
      const v = {
        sourceId: d.id,
        workbookId: u.book.id,
        workbookName: u.book.name,
        sheetId: u.sheet.id,
        sheetName: u.sheet.name,
        range: snapshot.source.range,
        fingerprint: u.fingerprint,
        records: snapshot.records.length,
        acceptedRecords: 0,
        amount: 0,
        periods: [],
        currencies: [],
        units: [],
      };
      o.sources.push(v);
      const g = new Map(),
        b = [];
      for (const y of snapshot.records) {
        let w = false;
        const E = (V, D) => {
            (s(d.id, V, D, y.row), (w = true));
          },
          A = f.map((V) => y.values[V]),
          k = y.values[p],
          F = `${ut(c.firstCol + p)}${y.row}`,
          I = JSON.stringify([u.book.id, u.sheet.id, F]);
        (a.has(I) &&
          E(
            "overlap",
            `Amount ${u.sheet.name}!${F} is included more than once. Remove overlapping sources to prevent double counting.`,
          ),
          a.add(I),
          A.some((V) => V === null || (typeof V == "string" && !V.trim())) &&
            E(
              "blank-category",
              "Every contributing row needs nonblank categories.",
            ),
          (typeof k != "number" || !Number.isFinite(k)) &&
            E(
              "amount",
              `Amount ${u.sheet.name}!${F} must be a finite number; text is not silently converted.`,
            ));
        const O = {
          period: m("period", y.values),
          currency: m("currency", y.values),
          unit: m("unit", y.values),
        };
        for (const V of ["period", "currency", "unit"]) {
          const D = O[V];
          if (
            typeof D == "string" &&
            D.trim() &&
            D.length <= 128 &&
            !/[\u0000-\u001f\u007f]/.test(D)
          ) {
            const B =
              v[
                V === "currency"
                  ? "currencies"
                  : V === "period"
                    ? "periods"
                    : "units"
              ];
            B.includes(D) || B.push(D);
          }
          D !== n[V] &&
            E(
              V,
              `Source ${V} ${JSON.stringify(D)} does not match ${JSON.stringify(n[V])}. No currency, unit or period conversion is implicit.`,
            );
        }
        const j = DR(A),
          z = g.get(j);
        (z !== undefined
          ? (s(
              d.id,
              "duplicate-category",
              `Repeated category also occurs at row ${z}. ${d.duplicatePolicy === "sum" ? "Explicit sum policy retains both contributions." : "Choose Sum repeated categories only after reviewing these records."}`,
              y.row,
              d.duplicatePolicy === "sum" ? "warning" : "error",
            ),
            d.duplicatePolicy === "reject" && (w = true))
          : g.set(j, y.row),
          w ||
            b.push({
              sourceId: d.id,
              workbookId: u.book.id,
              workbookName: u.book.name,
              sheetId: u.sheet.id,
              sheetName: u.sheet.name,
              row: y.row,
              address: F,
              locator: `${u.book.name} / ${snapshot.sheetName}!${F}`,
              categories: A,
              amount: k,
              period: O.period,
              currency: O.currency,
              unit: O.unit,
              link:
                u.book.id !== e.id
                  ? "approved snapshot"
                  : n.linkedAmounts
                    ? "live"
                    : "snapshot",
            }));
      }
      ((v.acceptedRecords = b.length),
        (v.amount = Op(b.map((y) => y.amount))),
        o.contributions.push(...b));
    } catch (u) {
      s(d.id, "source", u.message.slice(0, 4e3));
    }
  if (o.contributions.length > k1.contributions)
    throw Error(
      "Consolidation exceeds 2,000 contributions. Narrow the sources; no output was truncated.",
    );
  if (
    (o.contributions.length ||
      s(
        n.sources[0].id,
        "empty",
        "No valid contributing amounts. Review the sources and exceptions.",
      ),
    !o.exceptions.some((d) => d.severity === "error"))
  ) {
    ((o.subtotals = gU(o.contributions)),
      (o.grandTotal = Op(o.contributions.map((u) => u.amount))),
      (o.sourceTotal = Op(o.sources.map((u) => u.amount))),
      (o.difference = o.grandTotal - o.sourceTotal));
    const d =
      Number.EPSILON *
      Math.max(1, Math.abs(o.grandTotal), Math.abs(o.sourceTotal)) *
      o.contributions.length *
      4;
    (Math.abs(o.difference) > d &&
      s(
        n.sources[0].id,
        "reconciliation",
        "Source control totals do not reconcile to the contributing amounts.",
      ),
      (o.complete = !o.exceptions.some((u) => u.severity === "error")));
  }
  return o;
}
function tz(e, r) {
  const t = vI(e, r);
  return hr({
    allowed: [...t.allowed].sort(),
    hidden: !!r.includeHidden,
    readOnly: !!r.readOnly,
    approved: t.external
      .map((n) => ({
        id: n.workbook.id,
        sheets: [...n.allowedSheetIds].sort(),
        hidden: !!n.includeHidden,
      }))
      .sort((n, o) => n.id.localeCompare(o.id)),
  });
}
function Que(e, r) {
  if (
    typeof r != "string" ||
    !r.trim() ||
    r !== r.trim() ||
    r.length > 31 ||
    /[\\/?*\[\]:\u0000-\u001f\u007f]/.test(r) ||
    r.startsWith("'") ||
    r.endsWith("'") ||
    e.sheets.some((t) => t.name.toLowerCase() === r.toLowerCase())
  )
    throw Error(
      "Choose a distinct worksheet name of 1–31 characters without invalid characters.",
    );
  q1(e, r);
}
function rz(e, r, t = {}) {
  Que(e, r.outputName);
  const n = mU(r.definition),
    o = Jue(e, n, t);
  return {
    request: { definition: n, outputName: r.outputName },
    workbookHash: hr(e),
    contextHash: tz(e, t),
    review: o,
  };
}
function nz(e, r) {
  const t = ct(r);
  if (
    (t.lastRow - t.firstRow + 1) * (t.lastCol - t.firstCol + 1) >
    k1.outputCells
  )
    throw Error("Consolidation output range exceeds its cell budget.");
  const n = [];
  for (let o = t.firstRow; o <= t.lastRow; o++) {
    const a = [];
    for (let s = t.firstCol; s <= t.lastCol; s++) {
      const d = e.cells[`${ut(s)}${o + 1}`];
      a.push({
        raw: d?.raw ?? null,
        literal: !!(d != null && d.literal),
        note: d?.note ?? "",
      });
    }
    n.push(a);
  }
  return hr(n);
}

export { J$, Q$, X1, Y$, Yue, ez, m0, nz, q1, rz, tz };
