import i from "react/jsx-runtime";
import ne from "react";
import {
  Check,
  FileSpreadsheet,
  Search,
  ArrowUpRight,
  X,
  Filter,
  Plus,
  Layers,
  ClipboardCheck,
  Sparkles,
  History,
  CircleHelp,
  PanelRightClose,
  ListFilter,
  Copy,
  FlaskConical,
  ChartColumn,
  Clock3,
  Download,
  ShieldCheck,
} from "lucide-react";
import { es } from "./01-core-schemas-and-hashing.js";
import { Ig, Mt, b1, cx, ux } from "./02-financial-functions.js";
import {
  Af,
  CP,
  DB,
  Dg,
  Dw,
  Fw,
  Gy,
  JS,
  KS,
  NQ,
  Np,
  Og,
  Pv,
  RC,
  RN,
  SE,
  SR,
  Ta,
  Vn,
  WR,
  _t,
  aD,
  bB,
  cn,
  ct,
  f0,
  ff,
  gx,
  jB,
  jN,
  jQ,
  lD,
  p0,
  rh,
  t5,
  ul,
  ut,
  vx,
  y1,
  ys,
} from "./03-workbook-model.js";
import {
  $2,
  $ae,
  Aae,
  Bae,
  Dr,
  Gm,
  IC,
  J5,
  Lae,
  M,
  Mae,
  Qy,
  Uae,
  aI,
  ci,
  i$,
  lI,
  zae,
} from "./04-cell-input-paste-and-fill.js";
import { Hse, qse } from "./05-add-ins-and-custom-functions.js";
import { H1, La, OC, _$, ow } from "./07-xlsx-archive-reader.js";
import { Sk, V$, gI, x7 } from "./08-xlsx-export-and-preservation.js";
import { FC, Nue, Sue, Tue, W$ } from "./09-workbook-links.js";
import { Fue, X$ } from "./10-source-import-dialog.js";
import { $ue, Oo, zue } from "./11-workbook-formats-dialog.js";
import { ice, nce } from "./13-consolidation-dialog.js";
import { bce, vce, wce } from "./15-data-model-dialog.js";
import { Hce } from "./18-transformation-recipes-dialog.js";
import { Gce, Xce } from "./19-filter-views-dialog.js";
import { Kce, X2, Yce, Zce, s3 } from "./20-calculation-options-and-sources.js";
import { Ede, Sz } from "./21-aging-and-materiality-screens.js";
import { Ade, Cde, Sde, jI } from "./22-themes-and-cell-styles.js";
import { Fz, Nde, Rde } from "./23-format-cells-dialog.js";
import { Bde, Dde, FI, Pde, Ude, sw } from "./24-merge-and-audit-analysis.js";
import { Hde } from "./26-workbook-ribbon.js";
import { $z } from "./28-duplicate-review-and-structure.js";
import { dfe, sfe } from "./29-conditional-formatting-dialog.js";
import { pfe } from "./30-table-and-pivot-builder.js";
import { vfe } from "./31-pivot-explore-and-slicers.js";
import { Cfe, Wz } from "./32-drawing-and-chart-export.js";
import { Bme } from "./33-page-setup-and-print.js";
import { rxe } from "./34-function-browser-and-names.js";
import { ixe, lxe, sxe } from "./35-formula-compatibility-report.js";
import { fxe, mxe } from "./36-watch-window-and-goal-seek.js";
import { Sxe, kxe } from "./38-data-analysis-dialog.js";
import { Fxe } from "./39-report-outputs.js";
import { $xe, zxe } from "./40-what-if-analysis.js";
import { Jxe, Kxe, Yxe, Zxe } from "./41-automation-and-scripts.js";
import { e4e, t4e } from "./42-combine-populations.js";
import { s4e } from "./43-text-cleanup.js";
import { l4e } from "./44-clean-and-reshape.js";
import { m4e, p4e } from "./45-sorting-and-worksheets.js";
import { v4e } from "./46-workbook-manager.js";
import { WW, w4e, y4e } from "./47-fill-and-paste-special.js";
import { I4e, M4e } from "./48-format-cells-and-move.js";
import { B4e, j4e } from "./49-outlines-and-comments.js";
import { P4e } from "./50-materiality-screen.js";
import { Mj, Q4e, Z4e, fg } from "./51-pattern-checks-and-citations.js";
import { u2e } from "./52-workbook-comparison.js";
import {
  C2e,
  CS,
  E2e,
  R2e,
  T2e,
  Tp,
  Uj,
  _2e,
  dH,
  fH,
  up,
  w2e,
  y2e,
} from "./53-find-replace-go-to.js";
import {
  $2e,
  A2,
  P2e,
  V2e,
  xH,
  yH,
  z2e,
} from "./54-workbook-save-conflicts.js";
import {
  Gv,
  Jd,
  K2e,
  NS,
  Q2e,
  RH,
  TS,
  Vj,
  Y2e,
  eEe,
  tEe,
  wy,
} from "./55-external-source-refresh.js";

function IH(e) {
  const r = ne.useMemo(() => xH(), []),
    t = e.reliability ?? r,
    n = ne.useRef({
      propId: e.initialWorkbook?.id,
      emittedId: undefined,
      generation: 0,
    });
  n.current.propId !== e.initialWorkbook?.id &&
    (n.current.emittedId !== e.initialWorkbook?.id && n.current.generation++,
    (n.current.propId = e.initialWorkbook?.id));
  const o = ne.useCallback(
      (c) => {
        var h;
        ((n.current.emittedId = c.id),
          (h = e.onChange) == null || h.call(e, c));
      },
      [e.onChange],
    ),
    a = ne.useCallback(
      (c) => {
        var h;
        ((n.current.emittedId = c.id),
          (h = e.onWorkbookLoaded) == null || h.call(e, c));
      },
      [e.onWorkbookLoaded],
    );
  return i.jsx(Bae, {
    options: e.localization,
    children: i.jsx(W$.Provider, {
      value: { ...e.fileProcessingOptions, reliability: t },
      children: i.jsx(
        rEe,
        { ...e, reliability: t, onChange: o, onWorkbookLoaded: a },
        (e.persistenceKey ?? "session") +
          ":" +
          n.current.generation +
          ":" +
          RH(e.workbookStore),
      ),
    }),
  });
}
function rEe({
  extensions: e,
  linkedWorkbooks: r,
  onLinkedWorkbooksChange: t,
  allowLinkedWorkbookFiles: n = true,
  documentExtractor: o,
  externalSourceAdapter: a,
  reliability: s,
  saveWarningMs: d = 15e3,
  consolidationWorkbooks: u,
  consolidationDefinition: l,
  onConsolidationWorkbooksChange: c,
  allowConsolidationFiles: h = true,
  calculationWorkerFactory: f,
  initialWorkbook: p,
  onChange: m,
  onBeforeChange: v,
  onCellEditPendingChange: g,
  onWorkbookLoaded: b,
  onSelectionChange: y,
  adapter: w = Vj,
  actor: E = "Preparer",
  readOnly: A = false,
  persistenceKey: k,
  workbookStore: F,
  workbookStorePersistent: I = false,
  onPersistenceStateChange: O,
  externalProposal: j,
  onExternalProposalStaged: z,
}) {
  var dc, Y, re, Et, Xr;
  const V = ne.useMemo(() => xH(), []),
    D = s ?? V,
    [B, _] = ne.useState(false),
    C = Number.isFinite(d) && d >= 1 && d <= 12e4 ? d : 15e3,
    { text: T, language: N } = Dr(),
    R = Dr(),
    L = (fe) => aI(fe, R.entryLocale),
    $ = ne.useId(),
    [U, G] = ne.useState(() => rh(p ?? t5())),
    H = ne.useMemo(
      () => (U.extensions || (e != null && e.length) ? IC(U, e) : U),
      [U, e],
    ),
    [Q, X$1] = ne.useState([]),
    [J, q] = ne.useState([]),
    W = r ?? J,
    K = r !== undefined ? t : q,
    te = u ?? Q,
    ae = u ? c : X$1,
    [oe, de] = ne.useState(H.sheets.find((fe) => !fe.hidden).id),
    he =
      H.sheets.find((fe) => fe.id === oe && !fe.hidden) ??
      H.sheets.find((fe) => !fe.hidden),
    Se = ne.useMemo(
      () =>
        $2e(
          F ??
            (k
              ? yH(
                  {
                    getItem: (fe) => localStorage.getItem(fe),
                    setItem: (fe, Oe) => localStorage.setItem(fe, Oe),
                  },
                  k,
                )
              : zue()),
          D,
        ),
      [F, k, D],
    ),
    le = ne.useRef(null),
    [ce, ge] = ne.useState(null),
    se = ne.useRef(false),
    we = ne.useRef(!A);
  we.current = !A;
  const Ee = ne.useRef({ active: false, controllers: new Set() }),
    Ae = ne.useRef(null),
    ye = ne.useRef(O);
  ye.current = O;
  const [Te, Pe] = ne.useState(null),
    qe = ne.useRef(null),
    [ht, Re] = ne.useState(false),
    [Me, ke] = ne.useState(0),
    je = A || !!((dc = he.protection) != null && dc.enabled),
    [$e, Je] = ne.useState(null),
    [ze, at] = ne.useState(),
    [_e, Ze] = ne.useState(null),
    lt = ne.useRef(null);
  wce(lt);
  const [be, ft] = ne.useState(null),
    pt = ne.useRef(null),
    Fe = ne.useRef(null),
    [nt, mt] = ne.useState(0),
    [Wt, Zt] = ne.useState(false),
    [xt, or] = ne.useState("A1"),
    [Wr, Rr] = ne.useState("A1"),
    qt = ne.useRef("A1"),
    [Ye, bt] = ne.useState("A1"),
    [St, cr] = ne.useState([]);
  function dr(fe) {
    (bt(fe), cr([]));
  }
  function Ut(fe) {
    const Oe = ff(fe);
    (cr(Oe.slice(0, -1)), bt(Oe.at(-1)));
  }
  const Jt = ne.useMemo(() => [...St, Ye], [St, Ye]),
    fr = ne.useMemo(() => Jt.map(ct), [Jt]),
    br = Jt.join(", ");
  function rt() {
    if (Jt.length > 1)
      throw Error(
        "This command requires one rectangular range. Click a cell or enter a single range to continue.",
      );
  }
  const [Le, et] = ne.useState("A1"),
    [st, kt] = ne.useState(""),
    [gr, Ir] = ne.useState(null),
    kr = ne.useRef(null),
    Ke = ne.useRef({
      input: null,
      start: 0,
      end: 0,
      dragging: false,
      mode: true,
    }),
    [Tt, zt] = ne.useState(null),
    [Qe, Dt] = ne.useState(true),
    It = ne.useRef(null),
    ar = ne.useRef(() => {});
  (ne.useEffect(() => {
    let fe = 0,
      Oe;
    const Be = () => {
        if (!Ke.current.dragging || !Oe) return;
        const gt = to.current;
        if (gt) {
          const jt = gt.getBoundingClientRect();
          (Oe.y < jt.top + 28
            ? (gt.scrollTop -= 24)
            : Oe.y > jt.bottom - 28 && (gt.scrollTop += 24),
            Oe.x < jt.left + 28
              ? (gt.scrollLeft -= 36)
              : Oe.x > jt.right - 28 && (gt.scrollLeft += 36));
          const rr = document.elementFromPoint
            ?.call(
              document,
              Math.max(jt.left + 48, Math.min(jt.right - 8, Oe.x)),
              Math.max(jt.top + 28, Math.min(jt.bottom - 8, Oe.y)),
            )
            ?.closest("[data-address]");
          rr && gt.contains(rr) && ar.current(rr.dataset.address);
        }
        fe = requestAnimationFrame(Be);
      },
      Ue = (gt) => {
        Ke.current.dragging &&
          ((Oe = { x: gt.clientX, y: gt.clientY }),
          fe || (fe = requestAnimationFrame(Be)));
      },
      Xe = () => {
        const gt = Ke.current.dragging;
        ((Ke.current.dragging = false),
          cancelAnimationFrame(fe),
          (fe = 0),
          (Oe = undefined),
          gt && ls());
      };
    return (
      window.addEventListener("pointermove", Ue),
      window.addEventListener("pointerup", Xe),
      window.addEventListener("pointercancel", Xe),
      () => {
        (cancelAnimationFrame(fe),
          window.removeEventListener("pointermove", Ue),
          window.removeEventListener("pointerup", Xe),
          window.removeEventListener("pointercancel", Xe));
      }
    );
  }, []),
    ne.useEffect(() => {
      (kr.current || $l(), g == null || g(!!kr.current));
    }, [st, gr, H, g]));
  const [Lr, Pr] = ne.useState(null),
    [zr, jr] = ne.useState(""),
    [ii, fn] = ne.useState({}),
    [An, Bn] = ne.useState(),
    [ei, jo] = ne.useState("full"),
    Xo = JSON.stringify([H.id, he.id]),
    Ui = ii[Xo],
    [kn, po] = ne.useState(null),
    [Gi, ti] = ne.useState(false),
    Or = he.splitPanes,
    za = [
      H.id,
      he.id,
      Or?.horizontal !== undefined,
      Or?.vertical !== undefined,
      Or?.row,
      Or?.col,
    ].join(":"),
    an = E2e(za),
    Ra = an.views[an.activeId],
    na = Ra.top,
    va = Ra.left;
  function Co(fe) {
    Ca(an.active.current, fe, an.current.current[an.active.current].left);
  }
  function rs(fe) {
    Ca(an.active.current, an.current.current[an.active.current].top, fe);
  }
  const [ou, Go] = ne.useState(null),
    [ns, mo] = ne.useState([]),
    [Qi, ia] = ne.useState([]),
    [Hs, vn] = ne.useState(""),
    [qs, hn] = ne.useState(""),
    [go, Cu] = ne.useState(false),
    Yu = FC(),
    au = ne.useRef(0),
    ba = ne.useRef(null);
  ne.useEffect(
    () => () => {
      var fe;
      (au.current++, (fe = ba.current) == null || fe.abort());
    },
    [],
  );
  const [oi, xs] = ne.useState(null),
    [pl, su] = ne.useState(),
    [vo, Ml] = ne.useState(null),
    [is, lu] = ne.useState("duplicates"),
    [os, eo] = ne.useState(null),
    [oa, Dd] = ne.useState(25e3),
    [jl, Uc] = ne.useState(5),
    [Fl, Nf] = ne.useState("FY25-01"),
    [$c, zc] = ne.useState(null),
    [Od, um] = ne.useState("2025-12-31"),
    [Ju, C0] = ne.useState(""),
    [Md, A0] = ne.useState(Lae),
    [Xs, Bl] = ne.useState(null),
    [ml, Rf] = ne.useState(null),
    [gl, Ll] = ne.useState(null),
    [Qu, Vc] = ne.useState(false),
    [aa, Ia] = ne.useState(null),
    [ec, cm] = ne.useState(true),
    [vl, k0] = ne.useState(true),
    [Ao, Wc] = ne.useState(100),
    [wh, Hc] = ne.useState(false),
    [S0, Pl] = ne.useState(false),
    [as, jd] = ne.useState(""),
    [uu, sa] = ne.useState(""),
    ya = ne.useRef(null),
    $i = ne.useRef([]),
    [Fo, qc] = ne.useState(null),
    Gs = ne.useRef(Fo);
  Gs.current = Fo;
  const [Xc, Fd] = ne.useState(null),
    [ri, Au] = ne.useState(null),
    ku = ne.useRef(null),
    [Bo, Gc] = ne.useState(null),
    [Zc, Zs] = ne.useState(null);
  function Ul() {
    (Gc(null), Zs(null), (ku.current = null));
  }
  const [, If] = ne.useState(0),
    bl = Bde(H, f, _f, D),
    _a = bl.calculation,
    Qr = _a.calculator,
    Kc = ne.useMemo(() => {
      if (!Ui || Ui.snapshot !== H) return { result: null, error: "" };
      try {
        return { result: f0(H, he.id, Ui.definition, Qr), error: "" };
      } catch (fe) {
        return {
          result: null,
          error: fe instanceof Error ? fe.message : "Filtering is unavailable.",
        };
      }
    }, [H, he.id, Ui, Qr]),
    Hr = Kc.result,
    Bd = !!Ui && Ui.snapshot !== H,
    Yc = !!he.frozenRows,
    tc = ne.useMemo(() => Hr?.hiddenRows ?? cn(he, "row"), [he, Hr]),
    rc = ne.useMemo(() => cn(he, "column"), [he]),
    ni = ne.useMemo(() => ul(he.rowCount, tc), [he.rowCount, tc]),
    Sn = ne.useMemo(() => ul(he.colCount, rc), [he.colCount, rc]),
    Su = Math.max(
      he.frozenRows ?? 0,
      zr || kn ? 1 : 0,
      Hr && ct(Hr.context.range).firstRow === 0
        ? ct(Hr.context.range).firstRow + 1
        : 0,
    ),
    Tu = he.frozenCols ?? 0;
  ne.useEffect(() => {
    const fe = _t(xt);
    if (!ni.isHidden(fe.row) && !Sn.isHidden(fe.col)) return;
    const Oe = ni.next(fe.row, 1, true) ?? ni.next(fe.row, -1, true),
      Be = Sn.next(fe.col, 1, true) ?? Sn.next(fe.col, -1, true);
    if (Oe === null || Be === null || Oe === undefined || Be === undefined)
      return;
    const Ue = `${ut(Be)}${Oe + 1}`;
    (or(Ue), Rr(Ue), (qt.current = Ue), dr(Ue), Go(Ue));
  }, [xt, ni, Sn]);
  const co = Dg(he),
    Ks = Og(he),
    vi = Ao / 100,
    Es = vl ? 58 : 0,
    yl = ne.useRef(null),
    to = ne.useRef(null),
    Ld = ne.useRef(null),
    Nu = ne.useMemo(() => T2e(), []),
    ss = ne.useRef(null),
    Ys = ne.useRef(0),
    sr = ne.useRef({
      book: H,
      readOnly: je,
      sheetId: he.id,
      range: Ye,
      selectionKey: br,
    });
  sr.current = {
    book: H,
    readOnly: je,
    sheetId: he.id,
    range: Ye,
    selectionKey: br,
  };
  const nc = ne.useRef(undefined);
  function _f() {
    (If((fe) => fe + 1), Bl(null), Ll(null), eo(null));
  }
  function wa() {
    (bl.run(),
      vn(
        bl.isolated
          ? "Isolated calculation requested. Source cells and audit history are unchanged."
          : "Calculation refreshed. Source cells and audit history are unchanged.",
      ));
  }
  const cu = ne.useMemo(
      () =>
        new Set(
          (H.reviewThreads ?? [])
            .filter((fe) => fe.anchor.sheetId === he.id && !fe.orphaned)
            .map((fe) => fe.anchor.address),
        ),
      [H.reviewThreads, he.id],
    ),
    ic = ne.useMemo(() => (he.merges ?? []).map(ct), [he.merges]),
    ee = ne.useMemo(() => ct(Ye), [Ye]),
    ie = ne.useMemo(() => y1(Jt), [Jt]),
    ve = { sheetId: he.id, range: Ye },
    Ce = ne.useMemo(() => {
      try {
        return { cells: vx(he, Qr), warning: "" };
      } catch (fe) {
        return {
          cells: he.cells,
          warning:
            fe instanceof Error
              ? fe.message
              : "Array display exceeds this workbook's calculation budget.",
        };
      }
    }, [he, Qr]);
  ne.useEffect(() => {
    Ce.warning && hn(Ce.warning);
  }, [Ce.warning]);
  const Ie = ne.useMemo(() => Fw(he, Qr, Ce.cells), [he, Qr, Ce]),
    He = ne.useMemo(() => (fe) => Ie(fe).style, [Ie]),
    ot = Ie(xt).errors,
    $t = ne.useMemo(
      () => (
        Qr(he.id, xt),
        Qr.getReferenceDiagnostics().filter(
          (fe) => fe.sheetId === he.id && fe.address === xt,
        )
      ),
      [Qr, he.id, xt],
    ),
    yt = ne.useMemo(
      () => (
        Qr(he.id, xt),
        Qr.getNamedFormulaDiagnostics().filter(
          (fe) => fe.sheetId === he.id && fe.address === xt,
        )
      ),
      [Qr, he.id, xt],
    ),
    Pt = ne.useMemo(() => {
      let fe;
      return () =>
        fe ??
        (fe = Object.entries(Ce.cells).map(([Oe, Be]) => ({
          address: Oe,
          cell: Be,
          ..._t(Oe),
        })));
    }, [Ce.cells]),
    er = ne.useMemo(
      () =>
        (ie <= 500
          ? Np(Jt).map((fe) => ({ address: fe }))
          : Pt().filter(({ row: fe, col: Oe }) =>
              fr.some((Be) => Tp(Be, fe, Oe)),
            )
        )
          .map(({ address: fe }) => Qr(he.id, fe))
          .filter(NS),
      [Pt, fr, Jt, ie, Qr, he.id],
    ),
    Gt = er.reduce((fe, Oe) => fe + Oe, 0),
    ur = he.cells[xt],
    _r = ne.useMemo(() => {
      if (!zr && !kn) return null;
      const fe = new Map();
      for (const { address: Ue, row: Xe, cell: gt } of Pt()) {
        if (Xe < Su || ni.isHidden(Xe) || gt.raw === null) continue;
        const Nt = fe.get(Xe) ?? [];
        (Nt.push(Ue), fe.set(Xe, Nt));
      }
      const Oe = zr.toLowerCase(),
        Be = [...fe.keys()]
          .filter(
            (Ue) =>
              !Oe ||
              fe.get(Ue).some((Xe) =>
                String(Qr(he.id, Xe) ?? "")
                  .toLowerCase()
                  .includes(Oe),
              ),
          )
          .sort((Ue, Xe) => Ue - Xe);
      return (
        kn &&
          Be.sort((Ue, Xe) => {
            const gt = Qr(he.id, `${ut(kn.col)}${Ue + 1}`),
              Nt = Qr(he.id, `${ut(kn.col)}${Xe + 1}`);
            return (
              (NS(gt) && NS(Nt)
                ? gt - Nt
                : String(gt ?? "").localeCompare(String(Nt ?? ""))) *
              kn.direction
            );
          }),
        Be
      );
    }, [Pt, Qr, he.id, zr, kn, Su, ni]),
    sn = ne.useMemo(
      () => (_r ? new Map(_r.map((fe, Oe) => [fe, Oe])) : null),
      [_r],
    ),
    [bn, pn] = ne.useState({});
  ne.useEffect(
    () => pn({}),
    [
      he.id,
      he.cells,
      he.columnWidths,
      he.rowHeights,
      he.defaultRowHeight,
      he.defaultColumnWidth,
      Gi,
    ],
  );
  const en = ne.useMemo(
      () => ({
        ...bn,
        ...he.rowHeights,
        ...(ri?.axis === "row" ? { [ri.index]: ri.size } : {}),
      }),
      [bn, he.rowHeights, ri],
    ),
    ln = ne.useMemo(
      () => CS(he.rowCount, co, en, tc),
      [he.rowCount, co, en, tc],
    ),
    Xn = ne.useMemo(
      () =>
        CS(
          he.colCount,
          Ks,
          {
            ...he.columnWidths,
            ...(ri?.axis === "column" ? { [ri.index]: ri.size } : {}),
          },
          rc,
        ),
      [he.colCount, Ks, he.columnWidths, rc, ri],
    ),
    Ni = ni.countBefore(Su),
    yn = Sn.countBefore(Tu),
    Dn = Array.from({ length: Ni }, (fe, Oe) => ni.logicalIndex(Oe)),
    Ln = Array.from({ length: yn }, (fe, Oe) => Sn.logicalIndex(Oe)),
    Gn = ne.useMemo(() => {
      if (!_r) return Uj(ln, Ni);
      const fe = {};
      for (const [Oe, Be] of Object.entries(en)) {
        const Ue = sn?.get(Number(Oe));
        Ue !== undefined && Ue >= 0 && (fe[Ue] = Be);
      }
      return CS(_r.length, co, fe);
    }, [_r, sn, ln, Ni, en, co]),
    mn = ne.useMemo(() => Uj(Xn, yn), [Xn, yn]),
    ko = Gn.count,
    bi = (fe) => (_r ? _r[fe] : ni.logicalIndex(fe + Ni)),
    ro = (fe) => {
      if (!(fe < Su)) return sn ? sn.get(fe) : (ni.viewIndex(fe) ?? -1) - Ni;
    },
    zi = ln.offset(Ni),
    Ri = Xn.offset(yn);
  function Lo(fe) {
    const Oe = Math.max(1, fe.height - ((vl ? 22 : 0) + zi) * vi),
      Be = Math.max(1, fe.width - (Es + Ri) * vi),
      Ue = A2(Gn, fe.top, Oe, vi, 6),
      Xe = A2(mn, fe.left, Be, vi, 3),
      gt = Array.from({ length: Ue.last - Ue.first }, (Mr, Br) =>
        bi(Ue.first + Br),
      ),
      Nt = Array.from({ length: Xe.last - Xe.first }, (Mr, Br) =>
        Sn.logicalIndex(Xe.first + Br + yn),
      ),
      wt = Xe.leading,
      jt = Xe.trailing,
      rr = Nt.length + Ln.length + Number(vl) + +(wt > 0) + +(jt > 0);
    return {
      rowViewport: Oe,
      columnViewport: Be,
      rowWindow: Ue,
      columnWindow: Xe,
      shownRows: gt,
      shownColumns: Nt,
      leftGap: wt,
      rightGap: jt,
      visibleColumnCount: rr,
    };
  }
  const {
    rowViewport: Cs,
    columnViewport: xa,
    rowWindow: Ea,
    columnWindow: Js,
    shownRows: Po,
    shownColumns: Va,
  } = Lo(Ra);
  function Ca(fe, Oe, Be) {
    const Ue = Lo(an.current.current[fe]),
      Xe = A2(Gn, Oe, Ue.rowViewport, vi, 0).logicalOffset,
      gt = A2(mn, Be, Ue.columnViewport, vi, 0).logicalOffset,
      Nt = {};
    for (const wt of Pv(Or)) {
      const jt = Lo(an.current.current[wt]),
        rr = an.current.current[wt],
        Mr = aD(fe, wt, "row") ? up(Xe, Gn.total, jt.rowViewport, vi) : rr.top,
        Br = aD(fe, wt, "column")
          ? up(gt, mn.total, jt.columnViewport, vi)
          : rr.left;
      Nt[wt] = { top: Mr, left: Br };
      const xr = an.refs.current[wt];
      xr &&
        (Math.abs(xr.scrollTop - Mr) > 0.5 && (xr.scrollTop = Mr),
        Math.abs(xr.scrollLeft - Br) > 0.5 && (xr.scrollLeft = Br));
    }
    an.patch(Nt);
  }
  (ne.useEffect(() => {
    const fe = {};
    for (const Be of Pv(Or)) {
      const Ue = Lo(an.current.current[Be]),
        Xe = Or
          ? (ni.next(Or.row, 1, true) ?? ni.next(Or.row, -1, true) ?? 0)
          : 0,
        gt = Or
          ? (Sn.next(Or.col, 1, true) ?? Sn.next(Or.col, -1, true) ?? 0)
          : 0,
        Nt = Be.startsWith("bottom")
          ? up(
              Gn.offset(Math.max(0, ro(Xe) ?? 0)),
              Gn.total,
              Ue.rowViewport,
              vi,
            )
          : 0,
        wt = Be.endsWith("right")
          ? up(
              mn.offset(Math.max(0, (Sn.viewIndex(gt) ?? 0) - yn)),
              mn.total,
              Ue.columnViewport,
              vi,
            )
          : 0;
      fe[Be] = { top: Nt, left: wt };
      const jt = an.refs.current[Be];
      jt && ((jt.scrollTop = Nt), (jt.scrollLeft = wt));
    }
    an.patch(fe);
    const Oe = Pv(Or).at(-1);
    (an.activate(Oe), (to.current = an.refs.current[Oe] ?? null));
  }, [za]),
    ne.useEffect(() => {
      if (!to.current || typeof ResizeObserver > "u") return;
      const Oe = new ResizeObserver((Be) => {
        const Ue = {};
        for (const gt of Be) {
          const Nt = Number(gt.target.dataset.row),
            wt = gt.target.offsetHeight;
          Number.isInteger(Nt) &&
            he.rowHeights?.[Nt] === undefined &&
            ri?.axis !== "row" &&
            wt > co &&
            (Ue[Nt] = wt);
        }
        Object.keys(Ue).length &&
          pn((gt) =>
            Object.entries(Ue).every(([Nt, wt]) => gt[Number(Nt)] === wt)
              ? gt
              : { ...gt, ...Ue },
          );
      });
      return (
        Object.values(an.refs.current).forEach((Be) =>
          Be.querySelectorAll("tr[data-row]").forEach((Ue) => Oe.observe(Ue)),
        ),
        () => Oe.disconnect()
      );
    }),
    ne.useEffect(() => {
      if (!ou || !to.current) return;
      const fe = _t(ou),
        Oe = to.current;
      if (ni.isHidden(fe.row) || Sn.isHidden(fe.col)) {
        Go(null);
        return;
      }
      const Be = ro(fe.row);
      if (Be === undefined && fe.row >= Su) {
        (jr(""), po(null));
        return;
      }
      let Ue = na,
        Xe = va;
      if (Be !== undefined && Be >= 0) {
        const Nt = Gn.offset(Be),
          wt = Gn.offset(Be + 1);
        let jt = Ea.logicalOffset;
        (Nt < jt ? (jt = Nt) : wt > jt + Cs / vi && (jt = wt - Cs / vi),
          (Ue = up(jt, Gn.total, Cs, vi)));
      }
      const gt = (Sn.viewIndex(fe.col) ?? 0) - yn;
      if (gt >= 0) {
        const Nt = mn.offset(gt),
          wt = mn.offset(gt + 1);
        let jt = Js.logicalOffset;
        (Nt < jt ? (jt = Nt) : wt > jt + xa / vi && (jt = wt - xa / vi),
          (Xe = up(jt, mn.total, xa, vi)));
      }
      ((Oe.scrollTop = Ue), (Oe.scrollLeft = Xe), Co(Ue), rs(Xe), Go(null));
    }, [
      ou,
      sn,
      Gn,
      mn,
      Su,
      yn,
      ni,
      Sn,
      Cs,
      xa,
      vi,
      Ea.logicalOffset,
      Js.logicalOffset,
      na,
      va,
    ]));
  function Ii() {
    return Np(Jt);
  }
  function Ho() {
    To(
      Ii().map((fe) => ({
        sheetId: he.id,
        address: fe,
        cell: { ...he.cells[fe], raw: null, literal: false },
      })),
      "Clear contents",
    );
  }
  const Wa = `${he.name}!${br}`;
  (ne.useEffect(() => {
    kr.current || kt(L(he.cells[xt]));
  }, [he, xt, R.entryLocale]),
    ne.useEffect(() => {
      (et(br),
        du("selection", () =>
          y?.({
            sheetId: he.id,
            range: Ye,
            ...(Jt.length > 1 ? { ranges: Jt } : {}),
          }),
        ));
    }, [br, Ye, he.id, y]),
    ne.useEffect(() => {
      se.current = false;
      const fe = { active: true, controllers: new Set() };
      Ee.current = fe;
      try {
        const Oe = Se.load();
        if (((le.current = Oe), ge(Oe), Oe)) {
          (G(rh(Oe.workbook)),
            (sr.current.book = Oe.workbook),
            Zo(Oe.view),
            vn(Oe.warning ?? "Restored your locally saved workbook."),
            la("saved", Oe.workbook));
          try {
            du("loaded", () => b?.(Oe.workbook));
          } catch {
            hn("The host could not acknowledge the loaded workbook.");
          }
        }
        se.current = true;
      } catch (Oe) {
        (As(Oe),
          (qe.current =
            "The saved workbook could not be restored. Import a backup to recover it. Existing saved data will be preserved."),
          hn(qe.current));
      }
      return () => {
        var Be;
        fe.active = false;
        for (const Ue of fe.controllers) Ue.abort();
        const Oe = Ee.current;
        Oe.active = false;
        for (const Ue of Oe.controllers) Ue.abort();
        (Be = Ae.current) == null || Be.abort();
      };
    }, [Se]),
    ne.useEffect(() => {
      var fe;
      if (A) {
        for (const Oe of Ee.current.controllers) Oe.abort();
        (fe = Ae.current) == null || fe.abort();
      }
    }, [A]),
    ne.useEffect(() => {
      if (!k && !("coordination" in Se)) return;
      const fe = (Oe) => {
        const Be = le.current?.guard,
          Ue = k ?? Be?.namespace;
        if (!(!Ue || (Oe.key !== null && Oe.key !== Ue)))
          try {
            const gt = Se.load();
            if (
              JSON.stringify(gt?.guard ?? null) === JSON.stringify(Be ?? null)
            )
              return;
            As(
              new Oo(
                "conflict",
                "Another workbook window changed the saved version. Your local work has been retained.",
                gt?.guard,
              ),
            );
          } catch (gt) {
            As(gt);
          }
      };
      return (
        window.addEventListener("storage", fe),
        () => window.removeEventListener("storage", fe)
      );
    }, [Se, k]),
    ne.useEffect(() => {
      if (!k && !I) return;
      const fe = (Oe) => {
        (kr.current ||
          Ee.current.controllers.size ||
          JSON.stringify(le.current?.workbook) !==
            JSON.stringify(sr.current.book)) &&
          (Oe.preventDefault(), (Oe.returnValue = ""));
      };
      return (
        window.addEventListener("beforeunload", fe),
        () => window.removeEventListener("beforeunload", fe)
      );
    }, [k, I]),
    ne.useEffect(() => {
      if (!se.current || A) return;
      const fe = setTimeout(() => {
        try {
          const Oe = sr.current;
          if (Oe.readOnly || Te || Ee.current.controllers.size) return;
          const Be = yi(Oe.book, Oe.sheetId, Oe.range),
            Ue = le.current;
          if (
            Ue &&
            JSON.stringify(Ue.view) === JSON.stringify(Be) &&
            (Ue.workbook === Oe.book ||
              JSON.stringify(Ue.workbook) === JSON.stringify(Oe.book))
          )
            return;
          const Xe = _i(Oe.book, Be);
          Gv(Xe) && Xe.catch(() => {});
        } catch (Oe) {
          hn(
            `${Oe instanceof Error ? Oe.message : "The current selection could not be saved."} Changes remain in this editor; export JSON to keep a copy.`,
          );
        }
      }, 400);
      return () => clearTimeout(fe);
    }, [he.id, Ye, Se, A, Te]),
    ne.useEffect(() => {
      (eo(null), zc(null));
    }, [Ju, oa, jl, Fl, Od]),
    ne.useEffect(() => {
      var fe;
      return (
        Ys.current++,
        (fe = ss.current) == null || fe.abort(),
        Vc(false),
        Bl(null),
        Ll(null),
        () => {
          return ss.current?.abort();
        }
      );
    }, [H, he.id, br, w]),
    ne.useEffect(() => {
      if (!j || j.requestId === nc.current) return;
      nc.current = j.requestId;
      let fe;
      try {
        if (
          JSON.stringify(j.context) !==
          JSON.stringify(s3(H, { sheetId: he.id, range: Ye, ranges: Jt }))
        )
          throw new Error(
            "The workbook or selected context changed. Request a new external proposal.",
          );
        const Oe = X2(H, { sheetId: he.id, range: Ye }, j.proposal);
        (Ll(Mj(H, { sheetId: he.id, range: Ye }, Oe, j.context)),
          Bl(Oe),
          Pr("assistant"),
          hn(""),
          (fe = { accepted: true }));
      } catch (Oe) {
        const Be =
          Oe instanceof Error ? Oe.message : "Invalid external proposal.";
        (hn(Be), (fe = { accepted: false, error: Be }));
      }
      try {
        du("proposal", () => z?.(j.requestId, fe));
      } catch {
        hn("The host could not acknowledge the staged proposal.");
      }
    }, [j, H, he.id, br, z]));
  function yi(fe, Oe = he.id, Be = Ye) {
    const Ue =
      fe.sheets.find((Xe) => Xe.id === Oe && !Xe.hidden) ??
      fe.sheets.find((Xe) => !Xe.hidden);
    return {
      activeSheetId: Ue.id,
      selection: { sheetId: Ue.id, range: Ue.id === Oe ? Be : "A1" },
    };
  }
  function Zo(fe) {
    const [Oe, Be = Oe] = fe.selection.range.split(":");
    (de(fe.activeSheetId),
      or(Oe),
      Rr(Oe),
      dr(fe.selection.range),
      (qt.current = Be),
      po(null),
      jr(""),
      Co(0),
      rs(0),
      Go(Oe));
  }
  function la(fe, Oe, Be) {
    try {
      du("save-state", () => {
        return ye.current?.call(ye, {
          status: fe,
          workbookId: Oe.id,
          revision: Oe.revision,
          ...(Be ? { message: Be } : {}),
        });
      });
    } catch {
      hn(
        "The host could not acknowledge the save status. Export JSON to retain your work.",
      );
    }
  }
  function du(fe, Oe) {
    const Be = () =>
      hn(
        "The application could not acknowledge this operation. Your local workbook is retained.",
      );
    try {
      const Ue = D.run("host", fe, Oe);
      Ue && typeof Ue.then == "function" && Promise.resolve(Ue).catch(Be);
    } catch {
      Be();
    }
  }
  function As(fe) {
    const Oe =
        fe instanceof Oo
          ? fe
          : new Oo(
              "unavailable",
              "The storage adapter is unavailable. The workbook could not be saved.",
            ),
      Be =
        Oe instanceof Oo && Oe.code === "conflict"
          ? Oe
          : (Ee.current.conflict ?? Oe);
    (Be instanceof Oo && Be.code === "conflict" && (Ee.current.conflict = Be),
      Pe(Be),
      vn(""),
      (qe.current =
        Be.message +
        " Changes remain in this editor. Export JSON to keep a copy."),
      hn(qe.current),
      la(
        Be instanceof Oo && Be.code === "conflict" ? "conflict" : "error",
        sr.current.book,
        Be.message,
      ));
  }
  function Jc(fe) {
    const Oe = Ee.current;
    if (Oe.controllers.size >= 20) {
      const gt = new Oo(
        "limit",
        "Twenty saves are already waiting. Your latest changes remain in this editor; retry saving after the pending saves finish, or export your work.",
      );
      throw (As(gt), gt);
    }
    const Be = new AbortController();
    Oe.controllers.add(Be);
    const Ue = () => {
        if (!Oe.active || Oe !== Ee.current || Be.signal.aborted || !we.current)
          throw new Oo(
            "unavailable",
            "The save was cancelled because this editor or its access changed.",
          );
        if (Oe.conflict) throw Oe.conflict;
        if (Oe.stalled)
          throw new Oo(
            "unavailable",
            "A previous save is still unresolved. Wait for it to finish, then retry or export your local changes.",
          );
        return fe(Be.signal);
      },
      Xe = (gt) => {
        throw (
          Oe.controllers.delete(Be),
          Oe.active &&
            Oe === Ee.current &&
            (ke(Oe.controllers.size),
            Oe.controllers.size || (Oe.stalled = false),
            As(gt)),
          gt
        );
      };
    try {
      const gt = Oe.tail ? Oe.tail.then(Ue) : Ue();
      if (!Gv(gt)) return (Oe.controllers.delete(Be), gt);
      ke(Oe.controllers.size);
      const Nt = setTimeout(() => {
          if (!Oe.active || Oe !== Ee.current) return;
          ((Oe.stalled = true),
            D.start("persistence", "save").finish("failed", "timeout"),
            As(
              new Oo(
                "unavailable",
                "The save has not been acknowledged. Keep this editor open or export your local changes. Retrying is disabled until pending work finishes.",
              ),
            ));
        }, C),
        wt = gt.then(
          (rr) => (
            clearTimeout(Nt),
            Oe.controllers.delete(Be),
            Oe.active &&
              Oe === Ee.current &&
              (ke(Oe.controllers.size),
              Oe.controllers.size || (Oe.stalled = false)),
            rr
          ),
          (rr) => (clearTimeout(Nt), Xe(rr)),
        ),
        jt = wt.then(
          () => {},
          () => {},
        );
      return (
        (Oe.tail = jt),
        jt.then(() => {
          Oe.tail === jt && (Oe.tail = undefined);
        }),
        wt
      );
    } catch (gt) {
      return Xe(gt);
    }
  }
  function Qc(fe, Oe) {
    if (
      !Oe.active ||
      Oe !== Ee.current ||
      ((le.current = fe), ge(fe), !we.current || Oe.conflict)
    )
      return fe;
    if (
      fe.workbook === sr.current.book ||
      JSON.stringify(fe.workbook) === JSON.stringify(sr.current.book)
    ) {
      Pe(null);
      const Be = qe.current;
      (hn((Ue) => (Ue === Be ? "" : Ue)),
        (qe.current = null),
        Oe.stalled && D.start("persistence", "save").finish("recovered"),
        vn(
          k
            ? "Saved to this browser"
            : F
              ? "Saved to the application"
              : "Changes kept in this session",
        ),
        la("saved", fe.workbook));
    }
    return fe;
  }
  function Uo() {
    if (!(Ee.current.controllers.size || !we.current))
      try {
        const fe = sr.current.book,
          Oe = yi(fe),
          Be = z2e(Se, fe);
        if (Be) {
          ((se.current = true),
            Qc(Be, Ee.current),
            D.start("persistence", "save").finish("recovered"));
          return;
        }
        if (!se.current) {
          const Xe = Se.load();
          if (Xe)
            throw new Oo(
              "conflict",
              "Saved data is available again. Compare it before replacing your local changes.",
              Xe.guard,
            );
          se.current = true;
        }
        const Ue = _i(fe, Oe);
        Gv(Ue) && Ue.catch(() => {});
      } catch (fe) {
        As(fe);
      }
  }
  function _i(fe, Oe = yi(fe)) {
    if (!se.current) {
      const gt = new Oo(
        "corrupt",
        "Saved data needs recovery before this session can overwrite it.",
      );
      throw (As(gt), gt);
    }
    const Be = F || k ? structuredClone(fe) : fe,
      Ue = structuredClone(Oe),
      Xe = Ee.current;
    return (
      la("saving", Be),
      Jc((gt) => {
        return TS(
          Se.save(Be, Ue, le.current?.guard ?? null, {
            allowIdentityChange: !!(
              le.current && le.current.workbook.id !== Be.id
            ),
            signal: gt,
          }),
          (wt) => Qc(wt, Xe),
        );
      })
    );
  }
  function Tn() {
    Jd(
      new Blob([Sk(sr.current.book)], { type: "application/json" }),
      sr.current.book.name.replace(/[^a-zA-Z0-9 _.-]/g, "_") +
        " - local changes.json",
    );
  }
  function Aa(fe) {
    var Ue;
    const Oe = Se.load();
    if (!Oe || JSON.stringify(Oe.guard) !== JSON.stringify(fe.guard))
      throw new Error(
        "The saved version changed again. Refresh the comparison before reloading.",
      );
    const Be = Ee.current;
    Be.active = false;
    for (const Xe of Be.controllers) Xe.abort();
    ((Ee.current = { active: true, controllers: new Set() }),
      (Ue = Ae.current) == null || Ue.abort(),
      (le.current = Oe),
      ge(Oe),
      (sr.current.book = Oe.workbook),
      G(rh(Oe.workbook)),
      (kr.current = null),
      Ir(null),
      mo([]),
      ia([]),
      Bl(null),
      Ll(null),
      eo(null),
      Zo(Oe.view),
      Re(false),
      Pe(null),
      hn(""),
      vn(
        "Loaded the saved workbook. The discarded local undo history has been cleared.",
      ));
    try {
      du("change", () => m?.(Oe.workbook));
    } catch {
      hn("The host could not acknowledge the reloaded workbook.");
    }
    la("saved", Oe.workbook);
  }
  function di(fe, Oe = true, Be = false, Ue, Xe = false, gt) {
    var Nt;
    if (
      A ||
      ((je || sr.current.readOnly) &&
        !Be &&
        !(
          Xe &&
          gt &&
          sr.current.book.sheets.some((wt) => {
            var jt;
            return (
              wt.id === gt && !((jt = wt.protection) != null && jt.enabled)
            );
          })
        ))
    )
      throw new Error("This workbook is read-only.");
    if (!(
      !Ue &&
      (fe === sr.current.book ||
        (fe.revision === sr.current.book.revision &&
          JSON.stringify(fe) === JSON.stringify(sr.current.book)))
    )) {
      if (
        (Ue || (Nt = Ae.current) == null || Nt.abort(),
        (fe = rh(Ue || Xe ? fe : gx(fe))),
        v == null || v(fe, sr.current.book),
        Oe)
      ) {
        const wt = sr.current.book;
        (mo((jt) => [...jt.slice(-49), wt]), ia([]));
      }
      if (
        (SR(sr.current.book, fe),
        (sr.current.book = fe),
        G(fe),
        Bl(null),
        Ll(null),
        eo(null),
        hn(""),
        Ue)
      )
        ((le.current = Ue), ge(Ue));
      else
        try {
          const wt = _i(fe);
          Gv(wt) && wt.catch(() => {});
        } catch (wt) {
          As(wt);
        }
      try {
        du("change", () => m?.(fe));
      } catch {
        hn(
          "The host could not acknowledge the change. Export JSON to retain this revision.",
        );
      }
    }
  }
  function Pn(fe) {
    if (fe === "xlsxSource") {
      On(() =>
        Ml({
          snapshot: H,
          report: Sue(H),
          fileName: H.xlsxSource.fileName,
          inspectOnly: true,
        }),
      );
      return;
    }
    if (ui()) {
      if (
        (fe === "populationFilter" && Bn(undefined),
        Jt.length > 1 &&
          ![
            "search",
            "numberFormat",
            "cellStyle",
            "themes",
            "goToSpecial",
            "comments",
            "hyperlink",
            "sheets",
            "workbookProperties",
            "versions",
            "tracePrecedents",
            "traceDependents",
            "formulaWatches",
            "evaluateFormula",
            "functionCompatibility",
            "calculation",
            "calculateNow",
            "automation",
            "externalSources",
            "workbookFormats",
            "importSources",
            "workbookLinks",
            "extensions",
          ].includes(fe))
      ) {
        hn(
          "This command requires one rectangular range. Click a cell or enter a single range to continue.",
        );
        return;
      }
      if (fe === "calculateNow") {
        wa();
        return;
      }
      if (fe === "pasteSpecial" && Bo) {
        Zs(Bo.snapshot);
        return;
      }
      (fe === "sheets" && ft(he.id),
        fe === "automation" && at(undefined),
        Je(fe));
    }
  }
  function So(fe, Oe) {
    var Nt;
    const Be = sr.current.book,
      Ue = sw(Be, fe, E, Oe),
      Xe = fe.sheets.find((wt) => !Be.sheets.some((jt) => jt.id === wt.id));
    (di(
      Ue,
      true,
      Oe === "Unprotect sheet" &&
        !(
          (Nt = fe.sheets.find((wt) => wt.id === he.id)?.protection) != null &&
          Nt.enabled
        ),
    ),
      vn(Oe),
      Xe && (Je(null), wi(Xe.id)));
  }
  function ua(fe, Oe) {
    if (A) throw new Error("This workbook is read-only.");
    const Be = sr.current.book;
    if (JSON.stringify(Be) === JSON.stringify(fe)) return;
    const Ue =
      JSON.stringify(Be.structureProtection) !==
      JSON.stringify(fe.structureProtection);
    (di(sw(Be, fe, E, Oe), true, true),
      Ue && (mo([]), ia([])),
      fe.sheets.some((Xe) => Xe.id === he.id && !Xe.hidden) || Zo(yi(fe)),
      vn(Oe));
  }
  function ca(fe, Oe) {
    const Be = fe.sheets.find(
      (Ue) => !sr.current.book.sheets.some((Xe) => Xe.id === Ue.id),
    );
    (ua(fe, Oe), Be && (Je(null), wi(Be.id)));
  }
  function On(fe) {
    try {
      fe();
    } catch (Oe) {
      hn(
        Oe instanceof Error
          ? Oe.message
          : "The operation could not be completed.",
      );
    }
  }
  function To(fe, Oe = "Edit cells", Be, Ue) {
    try {
      if (Ue && fe.some((Xe) => Xe.sheetId !== Ue))
        throw Error("The formula destination changed.");
      di(p0(sr.current.book, fe, E, Oe), true, false, undefined, true, Ue);
    } catch (Xe) {
      throw (
        Xe instanceof jN &&
          Rf({
            error: Xe,
            patches: structuredClone(fe),
            action: Oe,
            snapshot: JSON.stringify(sr.current.book),
            selectionGuard: Be,
          }),
        Xe
      );
    }
  }
  function $l() {
    ((Ke.current = {
      input: null,
      start: 0,
      end: 0,
      dragging: false,
      mode: true,
    }),
      zt(null),
      Dt(true));
  }
  function ks(fe, Oe = false) {
    const Be = Ke.current,
      Ue = fe.selectionStart ?? fe.value.length,
      Xe = fe.selectionEnd ?? Ue;
    ((Be.input = fe),
      !(Be.restoring && !Oe) &&
        ((Be.restoring = false),
        (Be.start = Ue),
        (Be.end = Xe),
        (Oe || (Be.span && (Ue !== Be.span.end || Xe !== Ue))) &&
          ((Be.span = undefined),
          (Be.anchor = undefined),
          (Be.head = undefined),
          zt(null))));
  }
  function ls() {
    ((Ke.current.restoring = true),
      requestAnimationFrame(() => {
        var Be;
        if (!kr.current) return;
        const fe = Ke.current,
          Oe =
            (Be = fe.input) != null && Be.isConnected ? fe.input : pt.current;
        (Oe &&
          ((fe.input = Oe),
          Oe.focus({ preventScroll: true }),
          Oe.setSelectionRange(fe.start, fe.end)),
          (fe.restoring = false));
      }));
  }
  function Ss() {
    var Oe;
    const fe = sr.current.book.sheets.find((Be) => {
      return Be.id === (kr.current?.sheetId ?? he.id);
    });
    return A || !fe || !!((Oe = fe.protection) != null && Oe.enabled);
  }
  function ai(fe, Oe = false, Be) {
    var rr;
    const Ue = kr.current,
      Xe = Ke.current;
    if (!(Ue != null && Ue.value.startsWith("=")) || Ss() || go) return false;
    const gt = Xe.span ?? J5(Ue.value, Xe.start, Xe.end);
    if (!gt) return false;
    const Nt = Oe && Xe.sheetId === he.id && Xe.anchor ? Xe.anchor : fe,
      wt = Be?.range ?? (Nt === fe ? fe : wy(Nt, fe)),
      jt = Uae(Ue.value, gt, wt, Ue.sheetId === he.id ? undefined : he.name);
    return (
      us(
        jt.value,
        !!((rr = Xe.input) != null && rr.classList.contains("lg-cell-editor")),
      ),
      (Xe.span = jt.span),
      (Xe.start = Xe.end = jt.span.end),
      (Xe.anchor = Nt),
      (Xe.head = fe),
      (Xe.sheetId = he.id),
      mt(jt.span.end),
      zt({
        sheetId: he.id,
        range: wt,
        bounds: Be?.bounds ?? ct(wt),
      }),
      ls(),
      true
    );
  }
  ar.current = (fe) => {
    Ke.current.head !== fe && ai(fe, true);
  };
  function Ha(fe) {
    var Nt;
    if (
      fe.button !== 0 ||
      !((Nt = kr.current) != null && Nt.value.startsWith("="))
    )
      return false;
    const Oe = fe.target;
    if (Oe.closest('input,button,a,[role="separator"]')) return false;
    const Be = Oe.closest("[data-address]"),
      Ue = Oe.closest("[data-column]"),
      Xe = Oe.closest("[data-row-header]");
    let gt = false;
    if (Be) gt = ai(Be.dataset.address, fe.shiftKey);
    else if (Ue) {
      const wt = Number(Ue.dataset.column);
      gt = ai(`${ut(wt)}1`, false, {
        range: `${ut(wt)}:${ut(wt)}`,
        bounds: {
          firstRow: 0,
          lastRow: he.rowCount - 1,
          firstCol: wt,
          lastCol: wt,
        },
      });
    } else if (Xe) {
      const wt = Number(Xe.dataset.rowHeader);
      gt = ai(`A${wt + 1}`, false, {
        range: `${wt + 1}:${wt + 1}`,
        bounds: {
          firstRow: wt,
          lastRow: wt,
          firstCol: 0,
          lastCol: he.colCount - 1,
        },
      });
    }
    return gt
      ? (fe.preventDefault(),
        fe.stopPropagation(),
        (Ke.current.dragging = !!Be),
        (ya.current = null),
        true)
      : false;
  }
  function qa(fe) {
    if (fe.nativeEvent.isComposing || fe.altKey || Ss() || go) return false;
    const Oe = fe.currentTarget;
    if (!Oe.value.startsWith("=")) return false;
    ks(Oe);
    const Be = Ke.current;
    if (fe.key === "F2" && !fe.ctrlKey && !fe.metaKey)
      return (
        fe.preventDefault(),
        kr.current || us(Oe.value),
        (Be.mode = !Be.mode),
        Dt(Be.mode),
        true
      );
    if (fe.key === "F4" && !fe.ctrlKey && !fe.metaKey) {
      const jt = $ae(Oe.value, Be.start, Be.end);
      return jt
        ? (fe.preventDefault(),
          us(jt.value, Oe.classList.contains("lg-cell-editor")),
          (Be.span = jt.span),
          (Be.start = Be.end = jt.span.end),
          mt(jt.span.end),
          ls(),
          true)
        : false;
    }
    if ((fe.ctrlKey || fe.metaKey) && ["PageUp", "PageDown"].includes(fe.key)) {
      fe.preventDefault();
      const jt = H.sheets.filter((Br) => !Br.hidden),
        rr = jt.findIndex((Br) => Br.id === he.id),
        Mr = jt[rr + (fe.key === "PageUp" ? -1 : 1)];
      return (Mr && wi(Mr.id), true);
    }
    if (
      !Be.mode ||
      !fe.key.startsWith("Arrow") ||
      !kr.current ||
      !(Be.span ?? J5(Oe.value, Be.start, Be.end))
    )
      return false;
    fe.preventDefault();
    const Ue = _t(Be.sheetId === he.id && Be.head ? Be.head : xt),
      Xe = ["ArrowUp", "ArrowDown"].includes(fe.key),
      gt = ["ArrowUp", "ArrowLeft"].includes(fe.key) ? -1 : 1,
      Nt =
        fe.ctrlKey || fe.metaKey
          ? Vl(Ue.row, Ue.col, Xe, gt)
          : {
              row: Xe ? cs(Ue.row, gt) : Ue.row,
              col: Xe ? Ue.col : xl(Ue.col, gt),
            },
      wt = `${ut(Nt.col)}${Nt.row + 1}`;
    return (ai(wt, fe.shiftKey), fi(wt), true);
  }
  function wl(fe) {
    fe.sheetId !== he.id &&
      (de(fe.sheetId),
      or(fe.address),
      Rr(fe.address),
      dr(fe.address),
      po(null),
      jr(""),
      Go(fe.address));
  }
  function us(fe, Oe = false) {
    var Ue;
    const Be = kr.current ? undefined : Qr.getSpillAnchor(he.id, xt);
    if (Be && Be !== xt) {
      hn(
        `This cell belongs to the array formula in ${Be}. Edit that cell instead.`,
      );
      return;
    }
    ((Ue = Ae.current) == null || Ue.abort(),
      (kr.current = {
        sheetId: kr.current?.sheetId ?? he.id,
        address: kr.current?.address ?? xt,
        value: fe,
        locale: kr.current ? kr.current.locale : R.entryLocale,
        originalText: kr.current?.originalText ?? L(he.cells[xt]),
      }),
      kt(fe),
      Oe && Ir(fe));
  }
  function pi() {
    requestAnimationFrame(() => {
      return to.current?.focus({ preventScroll: true });
    });
  }
  function Da() {
    const fe = kr.current;
    (fe && wl(fe),
      $l(),
      (kr.current = null),
      Ir(null),
      kt(fe?.originalText ?? L(he.cells[xt])),
      pi());
  }
  function ui(fe) {
    const Oe = kr.current;
    if (!Oe) return true;
    try {
      const Xe = zae(Oe.value),
        gt = sr.current.book.sheets.find((Nt) => Nt.id === Oe.sheetId)?.cells[
          Oe.address
        ];
      return (
        Xe !== Oe.originalText &&
          To(
            [
              {
                sheetId: Oe.sheetId,
                address: Oe.address,
                cell: Mae(
                  Xe,
                  gt,
                  sr.current.book.datePolicy?.dateSystem,
                  Oe.locale,
                ),
              },
            ],
            "Edit cells",
            undefined,
            Oe.sheetId,
          ),
        (It.current =
          Oe.sheetId !== he.id
            ? { sheetId: Oe.sheetId, address: Oe.address }
            : null),
        wl(Oe),
        $l(),
        (kr.current = null),
        Ir(null),
        kt(Xe),
        true
      );
    } catch (Xe) {
      return (
        hn(Xe instanceof Error ? Xe.message : "The cell could not be saved."),
        false
      );
    }
  }
  function fi(fe) {
    Go(fe);
  }
  function bo(fe) {
    if (!Number.isFinite(fe) || fe < 50 || fe > 200) return;
    const Oe = fe / 100,
      Be = {};
    for (const Ue of Pv(Or)) {
      const Xe = an.current.current[Ue],
        gt = Lo(Xe),
        Nt = Math.max(1, Xe.height - ((vl ? 22 : 0) + zi) * Oe),
        wt = Math.max(1, Xe.width - (Es + Ri) * Oe),
        jt = up(gt.rowWindow.logicalOffset, Gn.total, Nt, Oe),
        rr = up(gt.columnWindow.logicalOffset, mn.total, wt, Oe);
      Be[Ue] = { top: jt, left: rr };
      const Mr = an.refs.current[Ue];
      Mr && ((Mr.scrollTop = jt), (Mr.scrollLeft = rr));
    }
    (an.patch(Be), Wc(fe), Go(xt));
  }
  function da(fe, Oe = false, Be = false) {
    if (!ui()) return;
    if (!Oe && !zr && !kn) {
      const gt = (he.merges ?? []).find((Nt) => lD(Nt, fe));
      gt && (fe = gt.split(":")[0]);
    }
    const Ue = Oe ? wy(Wr, fe) : fe;
    let Xe;
    try {
      Xe = ff(Be ? [...Jt, Ue] : [Ue]);
    } catch (gt) {
      hn(gt.message);
      return;
    }
    (fi(fe),
      (qt.current = fe),
      or(Oe ? Wr : fe),
      Ut(Xe),
      Oe || Rr(fe),
      Ir(null));
  }
  function wi(fe) {
    var Oe;
    if ((Oe = kr.current) != null && Oe.value.startsWith("=") && fe !== he.id) {
      (de(fe),
        po(null),
        jr(""),
        Co(0),
        rs(0),
        Ir(null),
        (Ke.current.anchor = undefined),
        (Ke.current.head = undefined),
        zt(null),
        ls());
      return;
    }
    ui() &&
      ((qt.current = "A1"),
      de(fe),
      or("A1"),
      Rr("A1"),
      dr("A1"),
      po(null),
      jr(""),
      Co(0),
      rs(0),
      eo(null),
      C0(""),
      to.current && ((to.current.scrollTop = 0), (to.current.scrollLeft = 0)),
      pi());
  }
  function Ts(fe, Oe) {
    const Be = sr.current.book.sheets.find((Xe) => Xe.id === fe);
    if (!Be || Be.hidden)
      throw Error(
        "Unhide the target worksheet with Manage worksheets before navigating to it.",
      );
    const Ue = _t(Oe);
    if (
      ul(Be.rowCount, cn(Be, "row")).isHidden(Ue.row) ||
      ul(Be.colCount, cn(Be, "column")).isHidden(Ue.col)
    )
      throw Error(
        "The target cell is hidden. Unhide its row/column or expand its outline before navigating.",
      );
    ui() &&
      (fn((Xe) => {
        const gt = { ...Xe };
        return (delete gt[JSON.stringify([H.id, fe])], gt);
      }),
      de(fe),
      or(Oe),
      Rr(Oe),
      (qt.current = Oe),
      dr(Oe),
      jr(""),
      po(null),
      Go(Oe));
  }
  function Pd(fe, Oe) {
    const Be = ff(Oe);
    (Ts(fe, Be.at(-1).split(":")[0]), Ut(Be), Je(null), pi());
  }
  function Vi(fe) {
    On(() => {
      const Oe = Ii();
      if (
        fe.style &&
        (Object.hasOwn(fe.style, "border") ||
          Object.hasOwn(fe.style, "borders")) &&
        Oe.some((Be) => (he.merges ?? []).some((Ue) => lD(Ue, Be)))
      )
        throw Error("Unmerge selected cells before changing border edges.");
      To(
        Oe.map((Be) => {
          return {
            sheetId: he.id,
            address: Be,
            cell: {
              ...he.cells[Be],
              raw: he.cells[Be]?.raw ?? null,
              ...fe,
              ...(fe.style
                ? {
                    style: {
                      ...he.cells[Be]?.style,
                      ...fe.style,
                    },
                  }
                : {}),
            },
          };
        }),
        "Format selection",
      );
    });
  }
  function Ud() {
    On(() => {
      if (!ui()) return;
      const fe = Ade(H, he.id, Le),
        Oe = fe.ranges,
        Be = ct(Oe.at(-1)),
        Ue = `${ut(Be.firstCol)}${Be.firstRow + 1}`,
        Xe = `${ut(Be.lastCol)}${Be.lastRow + 1}`;
      if (
        fe.sheetId === he.id &&
        (ni.isHidden(Be.firstRow) || Sn.isHidden(Be.firstCol))
      )
        throw Error(
          `${Ue} is hidden. Use View → Hide & Unhide or Data → Group & Outline to reveal it.`,
        );
      (fe.sheetId !== he.id && Ts(fe.sheetId, Ue),
        or(Ue),
        Rr(Ue),
        (qt.current = Xe),
        Ut(Oe),
        fe.sheetId === he.id && fi(Ue));
    });
  }
  function fu(fe = false) {
    On(() => {
      if (je || sr.current.readOnly)
        throw new Error("This workbook is read-only.");
      const Oe = fe ? Qi[Qi.length - 1] : ns[ns.length - 1];
      if (!Oe) return;
      if (
        Oe.id === H.id &&
        JSON.stringify(Oe.structureProtection) !==
          JSON.stringify(H.structureProtection)
      )
        throw new Error(
          "Undo cannot change workbook structure protection. Use Manage worksheets to unlock it.",
        );
      const Be =
        Oe.id !== H.id
          ? Ude(H, Oe, E, fe ? "redo" : "undo")
          : sw(
              H,
              {
                ...(es(Oe) ? Oe : structuredClone(Oe)),
                ...(H.rowOrigins
                  ? { rowOrigins: structuredClone(H.rowOrigins) }
                  : {}),
                revision: H.revision,
                events: H.events,
              },
              E,
              fe ? "Redo workbook change" : "Undo workbook change",
              { restoringHistory: true },
            );
      (di(Be, false),
        fe
          ? (ia(Qi.slice(0, -1)), mo([...ns, H]))
          : (mo(ns.slice(0, -1)), ia([...Qi, H])),
        Be.sheets.some((Ue) => Ue.id === oe && !Ue.hidden) ||
          wi(Be.sheets.find((Ue) => !Ue.hidden).id));
    });
  }
  async function Ns(fe) {
    var gt;
    if (!fe) return;
    (gt = ba.current) == null || gt.abort();
    const Oe = new AbortController(),
      Be = ++au.current;
    ba.current = Oe;
    const Ue = {
        ...Yu,
        signal: Yu.signal ? AbortSignal.any([Oe.signal, Yu.signal]) : Oe.signal,
      },
      Xe = () => Be === au.current;
    (Cu(true), hn(""));
    try {
      if (je) throw new Error("This workbook is read-only.");
      if (/\.(xls|xlsb|xltx|ods)$/i.test(fe.name)) {
        (su(fe), Je("workbookFormats"));
        return;
      }
      if (/\.(bas|vba|cls|ts|osts|xlsm|xlam|xltm)$/i.test(fe.name)) {
        if (fe.size > Mt.maxFileBytes)
          throw new Error("Files must be 10 MB or smaller.");
        const wt = await ow(fe, Ue);
        if (!Xe()) return;
        if (sr.current.book !== H || sr.current.readOnly)
          throw new Error(
            "The workbook changed while the file was being read. Open the file again.",
          );
        (at(wt), Je("automation"));
        return;
      }
      if (/\.(csv|tsv)$/i.test(fe.name)) {
        if (fe.size > Mt.maxFileBytes)
          throw new Error("Files must be 10 MB or smaller.");
        const wt = await _$(fe, {}, Ue);
        if (!Xe()) return;
        if (sr.current.book !== H || sr.current.readOnly)
          throw new Error(
            "The workbook changed while the file was being read. Open the file again.",
          );
        xs({
          snapshot: H,
          bytes: new Uint8Array(),
          file: fe,
          initial: wt,
          fileName: fe.name,
        });
        return;
      }
      const Nt = await H1(fe, Ue);
      if (!Xe()) return;
      if (sr.current.book !== H || sr.current.readOnly)
        throw new Error(
          "The workbook changed while the file was being read. Import again to replace the current revision.",
        );
      if (Nt.report) {
        Ml({
          snapshot: H,
          report: Nt.report,
          imported: Nt.workbook,
          fileName: fe.name,
        });
        return;
      }
      (di(Nt.workbook),
        wi(Nt.workbook.sheets.find((wt) => !wt.hidden).id),
        vn(Nt.warnings.length ? Nt.warnings.join(" ") : "Workbook imported."));
    } catch (Nt) {
      if (!Xe()) return;
      if (
        !Oe.signal.aborted &&
        (!(Nt instanceof La) ||
          ["FILE_ACTIVE", "FILE_INVALID"].includes(Nt.code)) &&
        /\.xlsx$/i.test(fe.name) &&
        fe.size <= Mt.maxFileBytes &&
        sr.current.book === H
      )
        try {
          const wt = await ow(fe, Ue);
          if (
            Xe() &&
            !sr.current.readOnly &&
            wt.items.length &&
            sr.current.book === H
          ) {
            (at(wt), Je("automation"));
            return;
          }
        } catch {}
      Xe() && hn(OC(Nt).message);
    } finally {
      Xe() &&
        ((ba.current = null), Cu(false), yl.current && (yl.current.value = ""));
    }
  }
  async function hu(fe) {
    (Cu(true), hn(""));
    try {
      const Oe =
        H.name.replace(/[^a-zA-Z0-9 _-]/g, "").slice(0, 80) || "Workbook";
      if (fe === "xlsx") {
        const Be = await gI(H);
        (Be.report.warnings.push(
          "Full workbook export includes all filtered and hidden rows. Session filter views are not saved in XLSX. Use CSV/TSV with Visible population for a filtered extract.",
        ),
          Ml({
            snapshot: H,
            report: { ...Be.report, fileName: `${Oe}-derived.xlsx` },
            blob: Be.blob,
            fileName: `${Oe}-derived.xlsx`,
          }));
      }
      (fe === "csv" &&
        xs({
          snapshot: H,
          sheetId: he.id,
          ...(Hr ? { population: Hr.context.definition } : {}),
        }),
        fe === "json" &&
          (vn(
            "Native JSON exports the full workbook, including all filtered and hidden rows and any named filter views. Active session masks are not saved; load named criteria, preview and apply after reopening.",
          ),
          Jd(new Blob([Sk(H)], { type: "application/json" }), `${Oe}.json`)));
    } catch (Oe) {
      hn(Oe instanceof Error ? Oe.message : "Export failed.");
    } finally {
      Cu(false);
    }
  }
  function Rs() {
    On(() => {
      if (ei === "visible" && !Hr)
        throw Error(
          "Review and apply a current population filter before visible-only analysis.",
        );
      const fe = Pde(H, he.id, is, {
        range: ei === "visible" && Hr ? Hr.context.range : Ju || Ta(he),
        ...(ei === "visible"
          ? {
              population: {
                definition: Hr?.context.definition ?? Ui?.definition,
                scope: "visible",
              },
            }
          : {}),
        threshold: oa,
        sampleSize: jl,
        seed: Fl,
        asOf: Od,
      });
      (eo(fe),
        zc({
          schemaVersion: 1,
          workbookId: H.id,
          workbookName: H.name,
          revision: H.revision,
          sheetId: he.id,
          sheetName: he.name,
          generatedAt: new Date().toISOString(),
          actor: E,
          populationScope: ei,
          ...(Hr ? { filterContext: Hr.context } : {}),
          kind: is,
          parameters: {
            range: ei === "visible" && Hr ? Hr.context.range : Ju || Ta(he),
            threshold: oa,
            sampleSize: jl,
            seed: Fl,
            asOf: Od,
          },
          result: fe,
        }),
        vn(
          ei === "visible"
            ? "Analysis completed on visible population rows; citations retain original worksheet addresses."
            : "Analysis completed on the full selected source range, including hidden and filtered rows.",
        ));
    });
  }
  async function No(fe, Oe) {
    var Xe;
    if (!fe.trim()) return null;
    if (Jt.length > 1)
      throw (
        hn("Select one rectangular range before asking the assistant."),
        new Error("Select one rectangular range before asking the assistant.")
      );
    (Xe = ss.current) == null || Xe.abort();
    const Be = ++Ys.current;
    (Vc(true), Bl(null), Ll(null), hn(""), vn(""));
    const Ue = new AbortController();
    ss.current = Ue;
    try {
      const gt = s3(H, ve),
        Nt = await Sz(w, fe, gt, {
          conversation: Oe,
          signal: Ue.signal,
          reliability: D,
        });
      if (Ue.signal.aborted || Be !== Ys.current) return null;
      const wt = X2(H, ve, Nt);
      return (Ll(Mj(H, ve, wt, gt)), Bl(wt), wt);
    } catch (gt) {
      if (
        (Be === Ys.current &&
          !Ue.signal.aborted &&
          hn(
            gt instanceof Error
              ? gt.message
              : "The assistant could not complete the request.",
          ),
        Ue.signal.aborted || Be !== Ys.current)
      )
        return null;
      throw gt;
    } finally {
      Be === Ys.current && Vc(false);
    }
  }
  function oc(fe) {
    On(() => {
      (di(fe),
        mo([]),
        ia([]),
        wi(fe.sheets.find((Oe) => !Oe.hidden).id),
        Hc(false),
        Pl(false),
        Pr(null));
    });
  }
  function ac() {
    On(() =>
      So(
        Gy(H, he.id, Yc ? 0 : 1, he.frozenCols ?? 0),
        Yc ? "Unfreeze rows" : "Freeze top row",
      ),
    );
  }
  function ed(fe = false) {
    (rt(), Ii());
    const Oe = { row: ee.firstRow, col: ee.firstCol },
      Be = { row: ee.lastRow, col: ee.lastCol },
      Ue = [];
    for (let Nt = Oe.row; Nt <= Be.row; Nt++) {
      const wt = [];
      for (let jt = Oe.col; jt <= Be.col; jt++) {
        const rr = `${ut(jt)}${Nt + 1}`,
          Mr = Qr(he.id, rr);
        wt.push(
          typeof Mr == "string" && /^[=+@-]/.test(Mr)
            ? `'${Mr}`
            : String(Mr ?? ""),
        );
      }
      Ue.push(wt);
    }
    const Xe = x7.unparse(Ue, {
        delimiter: "	",
        newline: `
`,
      }),
      gt = fe ? WW(H, he.id, Ye) : null;
    return (
      Gc(gt ? { snapshot: gt, text: Xe } : null),
      (ku.current = { text: Xe, snapshot: Aae(H, he.id, Ye) }),
      Xe
    );
  }
  async function Ru(fe) {
    var Oe;
    try {
      if (fe && je) throw new Error("This workbook is read-only.");
      const Be = H,
        Ue = ed(fe);
      if ((await navigator.clipboard.writeText(Ue), fe)) {
        if (
          sr.current.readOnly ||
          sr.current.book !== Be ||
          sr.current.sheetId !== he.id ||
          sr.current.selectionKey !== br
        )
          throw new Error("Workbook changed while copying. Copy again.");
        vn("Cut is ready. Select a destination and paste to preview the move.");
      }
      fe || vn("Selection copied.");
    } catch (Be) {
      (fe && Gc(null),
        hn(
          Be instanceof Error && Be.name !== "NotAllowedError"
            ? Be.message
            : "Clipboard access is unavailable. Use Ctrl/⌘ C or X with the grid focused.",
        ),
        (Oe = to.current) == null || Oe.focus());
    }
  }
  function zl(fe) {
    On(() => {
      if ((rt(), fe.length > Mt.maxFileBytes))
        throw new Error("Clipboard text is too large.");
      const Oe = ku.current?.text === fe ? ku.current : null;
      if (Oe && Bo?.text === fe) {
        Zs(Bo.snapshot);
        return;
      }
      if (Oe) {
        const wt = i$(H, he.id, Ye, Oe.snapshot, {
          mode: "all",
          deferValidationWarnings: true,
          readOnly: sr.current.readOnly,
        });
        (To(wt.patches, "Paste cells"), dr(wt.targetRange));
        return;
      }
      const Be = x7.parse(fe.replace(/\r?\n$/, ""), {
        delimiter: "	",
        skipEmptyLines: false,
      });
      if (Be.errors.length)
        throw new Error("Clipboard text contains malformed quoted cells.");
      const Ue = Be.data,
        Xe = _t(xt),
        gt = [];
      if (!Ue.length || Ue.some((wt) => wt.length !== Ue[0].length))
        throw new Error("Paste requires a rectangular range.");
      if (Ue.length * Ue[0].length > Mt.maxPatches)
        throw new Error(`Paste at most ${Mt.maxPatches} cells at a time.`);
      (Ue.forEach((wt, jt) =>
        wt.forEach((rr, Mr) => {
          if (Xe.row + jt >= he.rowCount || Xe.col + Mr >= he.colCount)
            throw new Error(
              "Paste exceeds the Excel worksheet boundary (XFD1048576).",
            );
          const Br = `${ut(Xe.col + Mr)}${Xe.row + jt + 1}`,
            xr = Gm(rr, he.cells[Br], sr.current.book.datePolicy?.dateSystem);
          gt.push({ sheetId: he.id, address: Br, cell: xr });
        }),
      ),
        To(gt, "Paste cells"),
        dr(wy(xt, `${ut(Xe.col + Ue[0].length - 1)}${Xe.row + Ue.length}`)));
    });
  }
  async function sc() {
    var fe;
    try {
      const Oe = H,
        Be = await navigator.clipboard.readText();
      if (
        sr.current.book !== Oe ||
        sr.current.sheetId !== he.id ||
        sr.current.selectionKey !== br
      )
        throw new Error("Workbook changed while pasting. Try again.");
      zl(Be);
    } catch {
      (hn(
        "Clipboard access is unavailable. Use Ctrl/⌘ V with the grid focused.",
      ),
        (fe = to.current) == null || fe.focus());
    }
  }
  function lc() {
    On(() => {
      if (!as) return;
      const fe = Object.keys(he.cells)
        .filter((Ue) => {
          const Xe = _t(Ue);
          return (
            !ni.isHidden(Xe.row) &&
            !Sn.isHidden(Xe.col) &&
            String(he.cells[Ue]?.raw ?? "")
              .toLowerCase()
              .includes(as.toLowerCase())
          );
        })
        .sort((Ue, Xe) => {
          const gt = _t(Ue),
            Nt = _t(Xe);
          return gt.row - Nt.row || gt.col - Nt.col;
        });
      if (!fe.length) {
        vn(`We couldn't find “${as}”.`);
        return;
      }
      const Oe = _t(xt),
        Be =
          fe.find((Ue) => {
            const Xe = _t(Ue);
            return Xe.row > Oe.row || (Xe.row === Oe.row && Xe.col > Oe.col);
          }) ?? fe[0];
      (da(Be), vn(`${fe.length} matching cell${fe.length === 1 ? "" : "s"}.`));
    });
  }
  function Ro() {
    On(() => {
      if (!as) throw new Error("Enter text to find.");
      const fe = new RegExp(as.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi"),
        Oe = Ii()
          .filter((Be) => {
            return (
              typeof he.cells[Be]?.raw == "string" &&
              String(he.cells[Be].raw).toLowerCase().includes(as.toLowerCase())
            );
          })
          .map((Be) => ({
            sheetId: he.id,
            address: Be,
            cell: {
              ...he.cells[Be],
              raw: String(he.cells[Be].raw).replace(fe, () => uu),
            },
          }));
      (Oe.length && To(Oe, "Replace text in selection"),
        vn(`Replaced text in ${Oe.length} cells.`));
    });
  }
  function cs(fe, Oe) {
    if (!_r) {
      const Xe = ni.viewIndex(fe) ?? ni.countBefore(fe);
      return ni.logicalIndex(Math.max(0, Math.min(ni.count - 1, Xe + Oe)));
    }
    const Be = [...Dn, ..._r],
      Ue = Be.indexOf(fe);
    return Be[Math.max(0, Math.min(Be.length - 1, Ue + Oe))] ?? fe;
  }
  function xl(fe, Oe) {
    const Be = Sn.viewIndex(fe) ?? Sn.countBefore(fe);
    return Sn.logicalIndex(Math.max(0, Math.min(Sn.count - 1, Be + Oe)));
  }
  function Vl(fe, Oe, Be, Ue) {
    const Xe = Be && _r ? [...Dn, ..._r] : null,
      gt = Be ? ni : Sn,
      Nt = Xe ? new Map(Xe.map((xr, Sr) => [xr, Sr])) : null,
      wt = (xr) => (Nt ? (Nt.get(xr) ?? null) : gt.viewIndex(xr)),
      jt = wt(Be ? fe : Oe);
    if (jt === null) return { row: fe, col: Oe };
    const rr = Pt().flatMap((xr) => {
        if (
          xr.cell.raw === null ||
          xr.cell.raw === "" ||
          (Be ? xr.col !== Oe : xr.row !== fe)
        )
          return [];
        const Sr = wt(Be ? xr.row : xr.col);
        return Sr === null ? [] : [Sr];
      }),
      Mr = w2e(jt, Xe?.length ?? gt.count, rr, Ue),
      Br = Xe ? Xe[Mr] : gt.logicalIndex(Mr);
    return Be ? { row: Br, col: Oe } : { row: fe, col: Br };
  }
  function pu(fe, Oe) {
    const Be = It.current;
    if (((It.current = null), Be)) {
      const Nt = H.sheets.find((xr) => xr.id === Be.sheetId),
        wt = _t(Be.address),
        jt = fe === "Enter" ? "row" : "column",
        Mr =
          ul(jt === "row" ? Nt.rowCount : Nt.colCount, cn(Nt, jt)).next(
            jt === "row" ? wt.row : wt.col,
            Oe ? -1 : 1,
          ) ?? (jt === "row" ? wt.row : wt.col),
        Br = `${ut(jt === "column" ? Mr : wt.col)}${(jt === "row" ? Mr : wt.row) + 1}`;
      (de(Be.sheetId), or(Br), Rr(Br), dr(Br), Go(Br), pi());
      return;
    }
    const Ue = _t(xt),
      Xe = fe === "Enter" ? cs(Ue.row, Oe ? -1 : 1) : Ue.row,
      gt = fe === "Tab" ? xl(Ue.col, Oe ? -1 : 1) : Ue.col;
    (da(`${ut(gt)}${Xe + 1}`), pi());
  }
  function td() {
    if (!ui()) return;
    dr(`A1:${ut(he.colCount - 1)}${he.rowCount}`);
    const fe = `${ut(Sn.logicalIndex(0))}${ni.logicalIndex(0) + 1}`;
    (or(fe), Rr(fe), (qt.current = fe), pi());
  }
  function El(fe, Oe, Be) {
    if (!ui()) return;
    const Ue = _t(Wr)[fe === "column" ? "col" : "row"],
      Xe = Be ? Math.min(Ue, Oe) : Oe,
      gt = Be ? Math.max(Ue, Oe) : Oe,
      Nt =
        fe === "column"
          ? `${ut(Oe)}${ni.logicalIndex(0) + 1}`
          : `${ut(Sn.logicalIndex(0))}${Oe + 1}`;
    (dr(
      fe === "column"
        ? `${ut(Xe)}1:${ut(gt)}${he.rowCount}`
        : `A${Xe + 1}:${ut(he.colCount - 1)}${gt + 1}`,
    ),
      (qt.current = Nt),
      or(Be ? Wr : Nt),
      Be || Rr(Nt),
      pi());
  }
  function Cl(fe, Oe) {
    rt();
    const { firstRow: Be, lastRow: Ue, firstCol: Xe, lastCol: gt } = ee,
      Nt =
        fe === "column"
          ? Be === 0 && Ue === he.rowCount - 1
          : Xe === 0 && gt === he.colCount - 1,
      wt = fe === "column" ? Xe : Be,
      jt = fe === "column" ? gt : Ue;
    if (Oe !== undefined && (!Nt || Oe < wt || Oe > jt)) return [Oe];
    const rr = new Set();
    for (const Br of Pt()) {
      const xr = fe === "column" ? Br.col : Br.row;
      xr >= wt && xr <= jt && rr.add(xr);
    }
    const Mr = fe === "column" ? he.columnWidths : he.rowHeights;
    for (const Br of Object.keys(Mr ?? {})) {
      const xr = Number(Br);
      xr >= wt && xr <= jt && rr.add(xr);
    }
    if (jt - wt < 500) for (let Br = wt; Br <= jt; Br++) rr.add(Br);
    return [...rr].sort((Br, xr) => Br - xr);
  }
  function Oa(fe, Oe, Be) {
    const Ue = sr.current.book,
      Xe = Ue.sheets.find((rr) => rr.id === he.id),
      gt = fe === "column" ? "columnWidths" : "rowHeights";
    if (
      !Object.entries(Oe).some(([rr, Mr]) => {
        return (Xe[gt]?.[rr] ?? (fe === "column" ? Og(Xe) : Dg(Xe))) !== Mr;
      })
    )
      return;
    const wt = _t(xt),
      jt =
        (Va.includes(wt.col) || Ln.includes(wt.col)) &&
        (Po.includes(wt.row) || Dn.includes(wt.row));
    (di({
      ...Ue,
      revision: Ue.revision + 1,
      sheets: Ue.sheets.map((rr) =>
        rr.id === he.id ? { ...rr, [gt]: { ...rr[gt], ...Oe } } : rr,
      ),
      events: [
        ...Ue.events,
        {
          id: Q2e(),
          timestamp: new Date().toISOString(),
          actor: E,
          action: Be,
          changes: [],
        },
      ],
    }),
      jt && fi(xt));
  }
  function Ko(fe, Oe) {
    On(() => {
      if (je || sr.current.readOnly || !ui()) return;
      const Be = Cl(fe, Oe),
        Ue = sr.current.book.sheets.find((jt) => jt.id === he.id),
        Xe = Vn(sr.current.book),
        gt =
          fe === "column"
            ? R2e(Ue, Be, Xe, Nu, Gi, sr.current.book.datePolicy?.dateSystem)
            : _2e(Ue, Be, Xe, Nu, Gi, sr.current.book.datePolicy?.dateSystem);
      (Oa(
        fe,
        gt,
        `AutoFit ${fe === "column" ? "column widths" : "row heights"}`,
      ),
        Au(null),
        pi());
    });
  }
  function $d(fe, Oe, Be) {
    if (
      (Be.preventDefault(),
      Be.stopPropagation(),
      Be.button !== 0 || je || !ui())
    )
      return;
    const Ue =
      fe === "column" ? Xn.size(Sn.viewIndex(Oe) ?? 0) : (en[Oe] ?? co);
    Au({
      axis: fe,
      index: Oe,
      start: fe === "column" ? Be.clientX : Be.clientY,
      initial: Ue,
      size: Ue,
    });
  }
  ne.useEffect(() => {
    if (!ri) return;
    const fe = (Xe) =>
        Au(
          (gt) =>
            gt && {
              ...gt,
              size: Math.round(
                Math.max(
                  gt.axis === "column" ? ux : b1,
                  Math.min(
                    gt.axis === "column" ? cx : Ig,
                    gt.initial +
                      ((gt.axis === "column" ? Xe.clientX : Xe.clientY) -
                        gt.start) /
                        vi,
                  ),
                ),
              ),
            },
        ),
      Oe = () => {
        (ri.size !== ri.initial &&
          On(() =>
            Oa(
              ri.axis,
              { [ri.index]: ri.size },
              `Resize ${ri.axis} ${ri.axis === "column" ? ut(ri.index) : ri.index + 1}`,
            ),
          ),
          Au(null));
      },
      Be = () => Au(null),
      Ue = (Xe) => {
        Xe.key === "Escape" &&
          (!Xe.defaultPrevented &&
            Xe.target instanceof Element &&
            Xe.target.closest(".lg-app") === lt.current &&
            Xe.preventDefault(),
          Be());
      };
    return (
      window.addEventListener("pointermove", fe),
      window.addEventListener("pointerup", Oe),
      window.addEventListener("pointercancel", Be),
      window.addEventListener("keydown", Ue),
      () => {
        (window.removeEventListener("pointermove", fe),
          window.removeEventListener("pointerup", Oe),
          window.removeEventListener("pointercancel", Be),
          window.removeEventListener("keydown", Ue));
      }
    );
  }, [ri, H, he.id, vi, je]);
  function Is(fe) {
    (fe.preventDefault(),
      fe.stopPropagation(),
      !(fe.button !== 0 || je || !ui()) &&
        On(() => {
          Ii();
          const Oe = {
            book: sr.current.book,
            sheetId: he.id,
            sourceRange: Ye,
            targetRange: Ye,
          };
          ((ya.current = null), (Gs.current = Oe), qc(Oe));
        }));
  }
  function uc(fe) {
    const Oe = Gs.current;
    if (!Oe) return;
    const Be = ct(Oe.sourceRange),
      Ue = _t(fe),
      Xe = Math.max(Be.firstRow - Ue.row, Ue.row - Be.lastRow, 0),
      gt = Math.max(Be.firstCol - Ue.col, Ue.col - Be.lastCol, 0),
      Nt = Xe >= gt ? Math.min(Be.firstRow, Ue.row) : Be.firstRow,
      wt = Xe >= gt ? Math.max(Be.lastRow, Ue.row) : Be.lastRow,
      jt = gt > Xe ? Math.min(Be.firstCol, Ue.col) : Be.firstCol,
      rr = gt > Xe ? Math.max(Be.lastCol, Ue.col) : Be.lastCol,
      Mr = wy(`${ut(jt)}${Nt + 1}`, `${ut(rr)}${wt + 1}`),
      Br = { ...Oe, targetRange: Mr };
    ((Gs.current = Br), qc(Br));
  }
  ne.useEffect(() => {
    if (!Fo) return;
    const fe = () => {
        ((Gs.current = null), qc(null));
      },
      Oe = () => {
        const Ue = Gs.current;
        if ((fe(), !(!Ue || Ue.targetRange === Ue.sourceRange))) {
          if (
            sr.current.readOnly ||
            sr.current.book !== Ue.book ||
            sr.current.sheetId !== Ue.sheetId
          ) {
            hn(
              "The workbook changed while filling. Select the cells and try again.",
            );
            return;
          }
          Fd(Ue);
        }
      },
      Be = (Ue) => {
        Ue.key === "Escape" &&
          (!Ue.defaultPrevented &&
            Ue.target instanceof Element &&
            Ue.target.closest(".lg-app") === lt.current &&
            Ue.preventDefault(),
          fe());
      };
    return (
      window.addEventListener("pointerup", Oe),
      window.addEventListener("pointercancel", fe),
      window.addEventListener("keydown", Be),
      () => {
        (window.removeEventListener("pointerup", Oe),
          window.removeEventListener("pointercancel", fe),
          window.removeEventListener("keydown", Be));
      }
    );
  }, [!!Fo]);
  function Qs(fe) {
    var wt;
    if (
      fe.target !== fe.currentTarget ||
      fe.defaultPrevented ||
      fe.nativeEvent.isComposing ||
      (fe.altKey &&
        !((fe.ctrlKey || fe.metaKey) && fe.key.toLowerCase() === "v")) ||
      fe.getModifierState("AltGraph")
    )
      return;
    const Oe = fe.ctrlKey || fe.metaKey,
      Be = fe.key.toLowerCase();
    if (
      (Oe &&
        ["tab", "enter", "f2", "f5", "f6", "delete", "backspace"].includes(
          Be,
        )) ||
      (Oe &&
        fe.shiftKey &&
        ![
          "z",
          "f",
          "home",
          "end",
          " ",
          "arrowleft",
          "arrowright",
          "arrowup",
          "arrowdown",
        ].includes(Be))
    )
      return;
    if (
      go &&
      ((Oe && ["z", "y", "b", "i", "u", "d", "r", "v"].includes(Be)) ||
        (!Oe &&
          (fe.key.length === 1 ||
            ["F2", "Delete", "Backspace"].includes(fe.key))))
    ) {
      (fe.preventDefault(),
        vn("Wait for the current workbook operation to finish."));
      return;
    }
    if (
      (fe.key === "ContextMenu" && !fe.shiftKey && !Oe) ||
      (fe.key === "F10" && fe.shiftKey && !fe.ctrlKey && !fe.metaKey)
    ) {
      fe.preventDefault();
      const rr = (
        fe.currentTarget.querySelector(`[data-address="${xt}"]`) ??
        fe.currentTarget
      ).getBoundingClientRect();
      Ze({
        x: rr.left + Math.min(40, rr.width),
        y: rr.top + Math.min(24, rr.height),
        kind:
          ct(Ye).lastRow - ct(Ye).firstRow + 1 === he.rowCount
            ? "column"
            : ct(Ye).lastCol - ct(Ye).firstCol + 1 === he.colCount
              ? "row"
              : "cell",
      });
      return;
    }
    if (fe.key === "Escape" && Bo) {
      (fe.preventDefault(),
        Ul(),
        vn("Cut canceled; source cells remain unchanged."));
      return;
    }
    if (fe.key === "Escape" && St.length) {
      (fe.preventDefault(), dr(xt), Rr(xt), (qt.current = xt));
      return;
    }
    if (fe.key === "Escape" && !fe.ctrlKey && !fe.metaKey && !fe.shiftKey) {
      (fe.preventDefault(),
        (wt = lt.current?.querySelector(
          '[role="tab"][aria-selected="true"]',
        )) == null || wt.focus());
      return;
    }
    if (
      (fe.key === "F5" && !fe.shiftKey && !Oe) ||
      ((fe.ctrlKey || fe.metaKey) && fe.key.toLowerCase() === "g")
    ) {
      (fe.preventDefault(), Pn("goToSpecial"));
      return;
    }
    if (
      (fe.ctrlKey || fe.metaKey) &&
      fe.shiftKey &&
      fe.key.toLowerCase() === "f"
    ) {
      (fe.preventDefault(), Pn("search"));
      return;
    }
    if (St.length && ["Tab", "Enter"].includes(fe.key)) {
      (fe.preventDefault(),
        On(() => {
          const jt = Np(Jt).filter((xr) => {
            const Sr = _t(xr);
            return !ni.isHidden(Sr.row) && !Sn.isHidden(Sr.col);
          });
          if (!jt.length) return;
          const rr = jt.indexOf(xt),
            Mr = fe.shiftKey ? -1 : 1,
            Br = jt[(rr + Mr + jt.length) % jt.length];
          ui() && (or(Br), Rr(Br), (qt.current = Br), fi(Br));
        }));
      return;
    }
    const Ue = _t(fe.shiftKey && fe.key.startsWith("Arrow") ? qt.current : xt);
    let Xe = Ue.row,
      gt = Ue.col;
    if ((fe.ctrlKey || fe.metaKey) && fe.key.toLowerCase() === "z") {
      (fe.preventDefault(), fu(fe.shiftKey));
      return;
    }
    if (fe.ctrlKey || fe.metaKey) {
      const jt = fe.key.toLowerCase();
      if (jt === "1") {
        (fe.preventDefault(), Pn("numberFormat"));
        return;
      }
      if (jt === "v" && fe.altKey) {
        (fe.preventDefault(), Pn("pasteSpecial"));
        return;
      }
      if (jt === " ") {
        (fe.preventDefault(), El("column", Ue.col, fe.shiftKey));
        return;
      }
      if (jt === "pageup" || jt === "pagedown") {
        fe.preventDefault();
        const rr = H.sheets.filter((xr) => !xr.hidden),
          Mr = rr.findIndex((xr) => xr.id === he.id),
          Br = rr[Mr + (jt === "pageup" ? -1 : 1)];
        Br && ui() && wi(Br.id);
        return;
      }
      if (jt.startsWith("arrow")) {
        fe.preventDefault();
        const rr = Vl(
          Ue.row,
          Ue.col,
          jt === "arrowup" || jt === "arrowdown",
          jt === "arrowup" || jt === "arrowleft" ? -1 : 1,
        );
        da(`${ut(rr.col)}${rr.row + 1}`, fe.shiftKey);
        return;
      }
      if (jt === "a") {
        (fe.preventDefault(), td());
        return;
      }
      if (["b", "i", "u"].includes(jt)) {
        if ((fe.preventDefault(), !je)) {
          const rr = jt === "b" ? "bold" : jt === "i" ? "italic" : "underline";
          Vi({ style: { [rr]: !ys(ur?.style)[rr] } });
        }
        return;
      }
      if (jt === "d") {
        (fe.preventDefault(),
          je ||
            On(() => {
              (rt(), So(Qy(H, $2(H, he.id, Ye, "down")), "Fill down"));
            }));
        return;
      }
      if (jt === "r") {
        (fe.preventDefault(),
          je ||
            On(() => {
              (rt(), So(Qy(H, $2(H, he.id, Ye, "right")), "Fill right"));
            }));
        return;
      }
      if (jt === "y") {
        (fe.preventDefault(), fu(true));
        return;
      }
      if (jt === "f" || jt === "h") {
        (fe.preventDefault(),
          Pl(true),
          requestAnimationFrame(() => {
            return Ld.current?.focus();
          }));
        return;
      }
      if (jt === "home") {
        (fe.preventDefault(),
          da(
            `${ut(Sn.logicalIndex(0))}${ni.logicalIndex(0) + 1}`,
            fe.shiftKey,
          ));
        return;
      }
      if (jt === "end") {
        fe.preventDefault();
        const rr = RC(he);
        da(
          `${ut(Sn.next(rr?.lastCol ?? 0, -1, true) ?? Sn.logicalIndex(0))}${(ni.next(rr?.lastRow ?? 0, -1, true) ?? ni.logicalIndex(0)) + 1}`,
          fe.shiftKey,
        );
        return;
      }
      return;
    }
    if (fe.shiftKey && fe.key === " ") {
      (fe.preventDefault(), El("row", Ue.row, false));
      return;
    }
    if (["Home", "PageUp", "PageDown"].includes(fe.key)) {
      (fe.preventDefault(),
        fe.key === "Home"
          ? (gt = Sn.logicalIndex(0))
          : (Xe = cs(
              Xe,
              Math.max(1, Math.floor(Cs / vi / co) - 1) *
                (fe.key === "PageUp" ? -1 : 1),
            )),
        da(`${ut(gt)}${Xe + 1}`, fe.shiftKey));
      return;
    }
    if (
      [
        "ArrowDown",
        "ArrowUp",
        "ArrowLeft",
        "ArrowRight",
        "Tab",
        "Enter",
      ].includes(fe.key)
    ) {
      if (
        (fe.preventDefault(),
        fe.key === "ArrowDown" && (Xe = cs(Xe, 1)),
        fe.key === "Enter" && (Xe = cs(Xe, fe.shiftKey ? -1 : 1)),
        fe.key === "ArrowUp" && (Xe = cs(Xe, -1)),
        fe.key === "ArrowLeft" && (gt = xl(gt, -1)),
        (fe.key === "ArrowRight" || fe.key === "Tab") &&
          (gt = xl(gt, fe.shiftKey && fe.key === "Tab" ? -1 : 1)),
        !fe.shiftKey && !zr && !kn)
      ) {
        const rr = ic.find((Mr) => Tp(Mr, Ue.row, Ue.col));
        rr &&
          (gt > Ue.col && (gt = Sn.next(rr.lastCol, 1) ?? Ue.col),
          gt < Ue.col && (gt = Sn.next(rr.firstCol, -1) ?? Ue.col),
          Xe > Ue.row && (Xe = ni.next(rr.lastRow, 1) ?? Ue.row),
          Xe < Ue.row && (Xe = ni.next(rr.firstRow, -1) ?? Ue.row));
      }
      const jt = `${ut(Math.max(0, Math.min(he.colCount - 1, gt)))}${Math.max(1, Math.min(he.rowCount, Xe + 1))}`;
      da(jt, fe.shiftKey && !["Tab", "Enter"].includes(fe.key));
      return;
    }
    (!je &&
      (fe.key === "Delete" || fe.key === "Backspace") &&
      (fe.preventDefault(),
      On(() =>
        To(
          Ii().map((jt) => ({
            sheetId: he.id,
            address: jt,
            cell: { ...he.cells[jt], raw: null },
          })),
          "Clear selection",
        ),
      )),
      !je &&
        (fe.key === "F2" ||
          (fe.key.length === 1 && !fe.ctrlKey && !fe.metaKey && !fe.altKey)) &&
        (fe.preventDefault(),
        (Ke.current.mode = fe.key !== "F2"),
        Dt(Ke.current.mode),
        us(fe.key === "F2" ? L(ur) : fe.key, true)));
  }
  function el() {
    aa &&
      On(() => {
        if (aa.kind === "note") Vi({ note: aa.value });
        else if (aa.kind === "rename") ua(NQ(H, aa.value), "Rename workbook");
        else {
          const fe = CP(H, aa.value),
            Oe = fe.sheets.at(-1);
          (ua(fe, `Add sheet: ${Oe.name}`), wi(Oe.id));
        }
        Ia(null);
      });
  }
  const _s = Object.fromEntries(Ln.map((fe, Oe) => [fe, Es + Xn.offset(Oe)])),
    Iu = (fe, Oe = false) =>
      i.jsxs(
        "th",
        {
          role: "columnheader",
          scope: "col",
          "data-column": fe,
          "aria-colindex": fe + 2,
          className: `${_t(xt).col === fe ? "lg-selected-header" : ""} ${Oe ? "lg-frozen-column" : ""}`,
          style: Oe
            ? { position: "sticky", left: _s[fe], zIndex: 7 }
            : undefined,
          onClick: (Be) => {
            ai(`${ut(fe)}1`, false, {
              range: `${ut(fe)}:${ut(fe)}`,
              bounds: {
                firstRow: 0,
                lastRow: he.rowCount - 1,
                firstCol: fe,
                lastCol: fe,
              },
            }) || El("column", fe, Be.shiftKey);
          },
          children: [
            ut(fe),
            i.jsx("span", {
              className: "lg-column-resize",
              role: "separator",
              "aria-label": `Resize column ${ut(fe)}`,
              "aria-orientation": "vertical",
              "aria-disabled": je,
              title: T("Drag to resize; double-click to AutoFit column width"),
              onPointerDown: (Be) => $d("column", fe, Be),
              onDoubleClick: (Be) => {
                (Be.stopPropagation(), Ko("column", fe));
              },
              onClick: (Be) => Be.stopPropagation(),
              lang: N("Drag to resize; double-click to AutoFit column width"),
            }),
          ],
        },
        fe,
      ),
    Ma = (fe, Oe, Be, Ue) => {
      var gt;
      return i.jsx(
        nEe,
        {
          idPrefix: `${$}-${Ue}`,
          row: fe,
          pinned: Oe,
          pinnedTop: (vl ? 22 : 0) + ln.offset(ni.viewIndex(fe) ?? 0),
          frozenColumns: Ln,
          frozenColumnOffsets: _s,
          headings: vl,
          columns: Be.shownColumns,
          leftGap: Be.leftGap,
          rightGap: Be.rightGap,
          minimumHeight: en[fe] ?? co,
          fixedHeight:
            he.rowHeights?.[fe] !== undefined ||
            (ri?.axis === "row" && ri.index === fe),
          onRowSelect: (Nt) => {
            ai(`A${fe + 1}`, false, {
              range: `${fe + 1}:${fe + 1}`,
              bounds: {
                firstRow: fe,
                lastRow: fe,
                firstCol: 0,
                lastCol: he.colCount - 1,
              },
            }) || El("row", fe, Nt);
          },
          onResize: (Nt) => $d("row", fe, Nt),
          onAutoFit: () => Ko("row", fe),
          onCancelEdit: Da,
          onFormulaInput: ks,
          onFormulaKey: qa,
          formulaHelpId: `${$}-formula-help`,
          onFormulaPick: ai,
          formulaReference: Tt?.sheetId === he.id ? Tt.bounds : null,
          formulaDraftActive: !!(
            (gt = kr.current) != null && gt.value.startsWith("=")
          ),
          onDragStart: (Nt, wt) => {
            ((ya.current = Nt), ($i.current = wt ? [...Jt] : []));
          },
          onDragEnter: (Nt, wt) => {
            if (Ke.current.dragging && wt === 1) {
              ai(Nt, true);
              return;
            }
            Gs.current && wt === 1
              ? uc(Nt)
              : ya.current &&
                wt === 1 &&
                ((qt.current = Nt),
                On(() => Ut([...$i.current, wy(ya.current, Nt)])),
                or(ya.current));
          },
          onNavigate: pu,
          onPivotDrill: (Nt) => {
            var jt;
            const wt = _t(Nt);
            return (jt = he.pivots) != null &&
              jt.some((rr) => {
                const Mr = ct(rr.outputRange);
                return (
                  wt.row > Mr.firstRow &&
                  wt.row <= Mr.lastRow &&
                  wt.col >= Mr.firstCol + Math.max(1, rr.rowFields.length) &&
                  wt.col <= Mr.lastCol
                );
              })
              ? (da(Nt), Pn("pivotExplore"), true)
              : false;
          },
          commentAddresses: cu,
          onReview: (Nt) => {
            (da(Nt), Pn("comments"));
          },
          onLink: (Nt, wt) => On(() => Ts(Nt, wt)),
          onFillStart: Is,
          onFillOptions: () => Pn("fillSeries"),
          fillDragging: !!Fo || St.length > 0,
          workbook: H,
          mergeBounds: ic,
          displayedRows: Oe ? Dn : Be.shownRows,
          filterHeader:
            Hr && fe === ct(Hr.context.range).firstRow ? Hr.fields : undefined,
          filteredColumns:
            Hr?.context.definition.columns.map((Nt) => Nt.column) ?? [],
          onColumnFilter: (Nt) => {
            (Pn("populationFilter"), Bn(Nt));
          },
          sortedView: !!(zr || kn),
          ruleStyles: He,
          conditionalFormatter: Ie,
          sheetId: he.id,
          cells: he.cells,
          active: xt,
          selectedBounds: Fo ? ct(Fo.targetRange) : ee,
          selectionBounds: Fo ? [ct(Fo.targetRange)] : fr,
          calc: Qr,
          showFormulas: Gi,
          editing:
            an.activeId === Ue && (!kr.current || kr.current.sheetId === he.id)
              ? gr
              : null,
          setEditing: (Nt) => (Nt === null ? Da() : us(Nt, true)),
          select: (Nt, wt, jt) => {
            var rr;
            (da(Nt, wt, jt),
              (rr = to.current) == null || rr.focus({ preventScroll: true }));
          },
          saveCell: ui,
          readOnly: je,
        },
        fe,
      );
    },
    zd = (fe, Oe, Be) =>
      fe > 0
        ? i.jsx(
            "tr",
            {
              "aria-hidden": "true",
              className: "lg-virtual-spacer",
              style: { height: fe },
              children: i.jsx("td", { colSpan: Be, style: { height: fe } }),
            },
            Oe,
          )
        : null;
  function tl(fe) {
    var Oe;
    return an.active.current === fe
      ? true
      : !((Oe = kr.current) != null && Oe.value.startsWith("=")) && !ui()
        ? false
        : (an.activate(fe), (to.current = an.refs.current[fe] ?? null), true);
  }
  function Al(fe) {
    const Oe = _t(xt),
      Be = zr || kn ? undefined : ic.find((Sr) => Tp(Sr, Oe.row, Oe.col)),
      Ue = Be ? `${ut(Be.firstCol)}${Be.firstRow + 1}` : xt,
      Xe = an.views[fe],
      gt = Lo(Xe),
      {
        rowWindow: Nt,
        columnWindow: wt,
        shownRows: jt,
        shownColumns: rr,
        leftGap: Mr,
        rightGap: Br,
        visibleColumnCount: xr,
      } = gt;
    return i.jsxs("div", {
      className: "lg-grid-stage",
      onPointerDownCapture: (Sr) => {
        Ha(Sr) || tl(fe) || (Sr.preventDefault(), Sr.stopPropagation());
      },
      onFocusCapture: () => tl(fe),
      children: [
        i.jsx("div", {
          className: "lg-grid-scroll",
          id: `${$}-${fe}-grid`,
          ref: (Sr) => {
            (an.bindings[fe](Sr),
              an.active.current === fe && (to.current = Sr));
          },
          tabIndex: 0,
          role: "grid",
          "aria-rowcount": he.rowCount + 1,
          "aria-colcount": he.colCount + 1,
          "aria-readonly": je,
          "aria-multiselectable": "true",
          "aria-describedby": `${$}-grid-help`,
          "aria-activedescendant":
            [...Dn, ...jt].includes(_t(xt).row) &&
            [...Ln, ...rr].includes(_t(xt).col)
              ? `${$}-${fe}-${Ue}`
              : undefined,
          "aria-label": Or
            ? `Worksheet pane ${fe.replace("-", " ")}`
            : T("Editable worksheet grid"),
          onScroll: (Sr) => {
            Ca(fe, Sr.currentTarget.scrollTop, Sr.currentTarget.scrollLeft);
          },
          onKeyDown: (Sr) => {
            var Wi;
            if (
              Or &&
              Sr.key === "F6" &&
              Sr.target === Sr.currentTarget &&
              !Sr.defaultPrevented &&
              !Sr.ctrlKey &&
              !Sr.metaKey &&
              !Sr.altKey
            ) {
              Sr.preventDefault();
              const Di = Pv(Or),
                $o = Di.indexOf(fe),
                Ds = Di[($o + (Sr.shiftKey ? Di.length - 1 : 1)) % Di.length];
              tl(Ds) &&
                ((Wi = an.refs.current[Ds]) == null ||
                  Wi.focus({ preventScroll: true }));
            } else Qs(Sr);
          },
          onContextMenu: (Sr) => {
            const Wi = Sr.target;
            if (Wi.closest("input,textarea,button,a,[contenteditable=true]"))
              return;
            const Di = Wi.closest("[data-address]"),
              $o = Wi.closest("[data-column]"),
              Ds = Wi.closest("[data-row-header]");
            if (!(!Di && !$o && !Ds) && ui()) {
              if ((Sr.preventDefault(), $o)) {
                const qo = Number($o.dataset.column);
                fr.some(
                  (yo) =>
                    yo.firstRow === 0 &&
                    yo.lastRow === he.rowCount - 1 &&
                    qo >= yo.firstCol &&
                    qo <= yo.lastCol,
                ) || El("column", qo, false);
              } else if (Ds) {
                const qo = Number(Ds.dataset.rowHeader);
                fr.some(
                  (yo) =>
                    yo.firstCol === 0 &&
                    yo.lastCol === he.colCount - 1 &&
                    qo >= yo.firstRow &&
                    qo <= yo.lastRow,
                ) || El("row", qo, false);
              } else if (Di) {
                const qo = _t(Di.dataset.address);
                fr.some((yo) => Tp(yo, qo.row, qo.col)) ||
                  da(Di.dataset.address);
              }
              Ze({
                x: Sr.clientX,
                y: Sr.clientY,
                kind: $o ? "column" : Ds ? "row" : "cell",
              });
            }
          },
          onCopy: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(),
              On(() => Sr.clipboardData.setData("text/plain", ed())));
          },
          onCut: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              je ||
              go ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(),
              On(() => {
                (Sr.clipboardData.setData("text/plain", ed(true)),
                  vn(
                    "Cut is ready. Select a destination and paste to preview the move.",
                  ));
              }));
          },
          onPaste: (Sr) => {
            Sr.target !== Sr.currentTarget ||
              je ||
              go ||
              Sr.defaultPrevented ||
              (Sr.preventDefault(), zl(Sr.clipboardData.getData("text/plain")));
          },
          onPointerUp: () => {
            ya.current = null;
          },
          onPointerLeave: (Sr) => {
            Sr.buttons === 0 && (ya.current = null);
          },
          children: i.jsxs("table", {
            className: `lg-grid ${ec ? "" : "no-gridlines"} ${vl ? "" : "no-headings"}`,
            style: { zoom: vi, width: wt.total + Ri + Es, minWidth: 0 },
            role: "presentation",
            children: [
              i.jsxs("colgroup", {
                children: [
                  vl && i.jsx("col", { style: { width: Es } }),
                  Ln.map((Sr) =>
                    i.jsx(
                      "col",
                      { style: { width: Xn.size(Sn.viewIndex(Sr)) } },
                      Sr,
                    ),
                  ),
                  Mr > 0 && i.jsx("col", { style: { width: Mr } }),
                  rr.map((Sr) =>
                    i.jsx(
                      "col",
                      { style: { width: Xn.size(Sn.viewIndex(Sr)) } },
                      Sr,
                    ),
                  ),
                  Br > 0 && i.jsx("col", { style: { width: Br } }),
                ],
              }),
              i.jsx("thead", {
                children: i.jsxs("tr", {
                  role: "row",
                  "aria-rowindex": 1,
                  children: [
                    vl &&
                      i.jsx("th", {
                        role: "columnheader",
                        className: "lg-corner",
                        "aria-label": T("Select all cells"),
                        onClick: td,
                        lang: N("Select all cells"),
                        "data-ui-label": "Select all cells",
                        children: i.jsx("span", {
                          className: "lg-sr-only",
                          children: i.jsx(M, { children: "Select all cells" }),
                        }),
                      }),
                    Ln.map((Sr) => Iu(Sr, true)),
                    Mr > 0 &&
                      i.jsx("th", {
                        className: "lg-column-spacer",
                        "aria-hidden": "true",
                      }),
                    rr.map((Sr) => Iu(Sr)),
                    Br > 0 &&
                      i.jsx("th", {
                        className: "lg-column-spacer",
                        "aria-hidden": "true",
                      }),
                  ],
                }),
              }),
              i.jsxs("tbody", {
                children: [
                  Dn.map((Sr) => Ma(Sr, true, gt, fe)),
                  zd(Nt.leading, "before", xr),
                  jt.map((Sr) => Ma(Sr, false, gt, fe)),
                  zd(Nt.trailing, "after", xr),
                ],
              }),
            ],
          }),
        }),
        i.jsx(Cfe, {
          workbook: H,
          sheetId: he.id,
          selection: JSON.stringify(Jt),
          calc: Qr,
          readOnly: je,
          scale: vi,
          onWorkbookChange: So,
          positionForCell: (Sr) => {
            const Wi = _t(Sr);
            if (ni.isHidden(Wi.row) || Sn.isHidden(Wi.col)) return null;
            const Di = ro(Wi.row),
              $o = Wi.row < Su,
              Ds = Wi.col < Tu;
            if (!$o && Di === undefined) return null;
            const qo =
                (Es +
                  (Ds
                    ? Xn.offset(Sn.viewIndex(Wi.col))
                    : Ri +
                      mn.offset(Sn.viewIndex(Wi.col) - yn) -
                      wt.logicalOffset)) *
                vi,
              yo =
                ((vl ? 22 : 0) +
                  ($o
                    ? ln.offset(ni.viewIndex(Wi.row))
                    : zi + Gn.offset(Math.max(0, Di)) - Nt.logicalOffset)) *
                vi;
            return qo < -5e3 || yo < -5e3 || qo > Xe.width || yo > Xe.height
              ? null
              : { left: qo, top: yo };
          },
        }),
      ],
    });
  }
  const cc = ne.useCallback((fe) => {
      var Oe;
      (Ze(null),
        fe && ((Oe = to.current) == null || Oe.focus({ preventScroll: true })));
    }, []),
    no = go
      ? "Wait for the current workbook operation to finish."
      : je
        ? A
          ? "The host has opened this workbook read-only."
          : "Unprotect the worksheet before editing."
        : undefined,
    fo =
      Jt.length > 1
        ? "Select one rectangular range for this command."
        : undefined;
  let Xa;
  try {
    _e && Ii();
  } catch (fe) {
    Xa = fe.message;
  }
  const ds = [
    { name: "Cut", run: () => void Ru(true), reason: no ?? fo ?? Xa },
    { name: "Copy", run: () => void Ru(false), reason: fo ?? Xa },
    { name: "Paste", run: () => void sc(), reason: no ?? fo },
    { name: "Paste Special", run: () => Pn("pasteSpecial"), reason: no ?? fo },
    { name: "Format Cells", run: () => Pn("numberFormat"), reason: no },
    { name: "Cell Styles", run: () => Pn("cellStyle"), reason: no },
    {
      name: "Insert / Delete / Move Rows",
      run: () => Pn("insertRow"),
      reason: no ?? fo,
    },
    {
      name: "Insert / Delete / Move Columns",
      run: () => Pn("insertColumn"),
      reason: no ?? fo,
    },
    { name: "AutoFit Column Width", run: () => Ko("column"), reason: no ?? fo },
    { name: "AutoFit Row Height", run: () => Ko("row"), reason: no ?? fo },
    {
      name: "Hide / Unhide / Freeze",
      run: () => Pn("viewSettings"),
      reason: no ?? fo,
    },
    {
      name: "Clear Contents",
      run: () =>
        On(() =>
          To(
            Ii().map((fe) => ({
              sheetId: he.id,
              address: fe,
              cell: { ...he.cells[fe], raw: null },
            })),
            "Clear selection",
          ),
        ),
      reason: no ?? Xa,
    },
    {
      name: "Add / Edit Note",
      run: () => Ia({ kind: "note", value: ur?.note ?? "" }),
      reason: no,
    },
    { name: "Comments", run: () => Pn("comments") },
    { name: "Link", run: () => Pn("hyperlink") },
  ];
  return i.jsxs("div", {
    className: "lg-app",
    ref: lt,
    lang: R.locale,
    role: "region",
    "aria-label": H.name,
    onKeyDown: (fe) => {
      var wt;
      if (
        fe.key !== "F6" ||
        fe.defaultPrevented ||
        fe.ctrlKey ||
        fe.metaKey ||
        fe.altKey ||
        fe.nativeEvent.isComposing ||
        $e ||
        aa ||
        (kr.current && (kr.current.value !== kr.current.originalText || !ui()))
      )
        return;
      fe.preventDefault();
      const Oe = [
          lt.current?.querySelector('[role="tab"][aria-selected="true"]'),
          pt.current,
          to.current,
          lt.current?.querySelector('[data-sheet-id][aria-current="true"]'),
        ].filter((jt) => !!jt),
        Be = document.activeElement,
        Ue =
          Be != null && Be.closest(".xl-ribbon")
            ? 0
            : Be != null && Be.closest(".lg-sheet-tabs")
              ? 3
              : Oe.findIndex((jt) => jt === Be || jt.contains(Be)),
        Xe = fe.shiftKey ? -1 : 1;
      (wt =
        Oe[
          Ue < 0
            ? fe.shiftKey
              ? Oe.length - 1
              : 0
            : (Ue + Xe + Oe.length) % Oe.length
        ]) == null || wt.focus({ preventScroll: true });
    },
    children: [
      i.jsxs("div", {
        className: "lg-skip-links",
        children: [
          i.jsx("button", {
            onClick: () => {
              return to.current?.focus();
            },
            children: i.jsx(M, { children: "Skip to worksheet" }),
          }),
          i.jsx("button", {
            onClick: () => {
              return pt.current?.focus();
            },
            children: i.jsx(M, { children: "Skip to formula bar" }),
          }),
        ],
      }),
      i.jsx("p", {
        id: `${$}-grid-help`,
        className: "lg-sr-only",
        children: i.jsx(M, {
          children:
            "Arrow keys move between cells. F2 edits. Escape returns to the ribbon. Use the selected range field to jump to a cell. Formula names, references and decimal points inside formulas use invariant syntax; enter dates as YYYY-MM-DD.",
        }),
      }),
      i.jsx("p", {
        className: "lg-sr-only",
        role: "status",
        "aria-live": "polite",
        "aria-atomic": "true",
        "aria-label": R.text("Cell position"),
        children: R.message(
          "{sheet}, cell {address}, row {row}, column {column}. {value}{formula}{access}",
          {
            sheet: he.name,
            address: xt,
            row: _t(xt).row + 1,
            column: _t(xt).col + 1,
            value: (Qr(he.id, xt) == null
              ? R.text("Blank")
              : lI(
                  Qr(he.id, xt),
                  he.cells[xt],
                  H.datePolicy?.dateSystem,
                  R.entryLocale,
                )
            ).slice(0, 250),
            formula:
              typeof he.cells[xt]?.raw == "string" &&
              String(he.cells[xt].raw).startsWith("=") &&
              !he.cells[xt].literal
                ? R.message(" Formula: {formula}.", {
                    formula: String(he.cells[xt].raw).slice(0, 250),
                  })
                : "",
            access: je ? R.text(" Read only.") : "",
          },
        ),
      }),
      _e &&
        i.jsx(
          Kce,
          {
            ..._e,
            label: `${_e.kind === "cell" ? "Cells" : _e.kind === "row" ? "Rows" : "Columns"} ${Ye}`,
            actions: ds,
            workbook: H,
            selection: `${he.id}:${Jt.join(";")}`,
            onClose: cc,
          },
          `${_e.x}:${_e.y}:${_e.kind}`,
        ),
      Xc &&
        i.jsx(w4e, {
          workbook: H,
          sheetId: Xc.sheetId,
          sourceRange: Xc.sourceRange,
          targetRange: Xc.targetRange,
          readOnly: je,
          onWorkbookChange: So,
          onSelectionChange: dr,
          onClose: () => {
            (Fd(null), pi());
          },
        }),
      Zc &&
        i.jsx(I4e, {
          workbook: H,
          source: Zc,
          sheetId: he.id,
          range: Ye,
          readOnly: A,
          onWorkbookChange: ua,
          onSelectionChange: dr,
          onApplied: Ul,
          onClose: () => {
            (Zs(null), pi());
          },
        }),
      $e &&
        (() => {
          const fe = {
            mode: $e,
            workbook: H,
            sheetId: he.id,
            range: Ye,
            readOnly: $e === "protection" ? A : je,
            onClose: () => Je(null),
            onWorkbookChange: So,
          };
          if ($e === "automation")
            return i.jsx(
              Jxe,
              {
                workbook: H,
                sheetId: he.id,
                range: Ye,
                readOnly: je,
                initialInspection: ze,
                onClose: fe.onClose,
                onPreserve: (Ue) => {
                  const Xe = sr.current.book,
                    gt = Zxe(Xe, Ue, {
                      actor: E,
                      readOnly: sr.current.readOnly,
                    });
                  (gt !== Xe && di(gt),
                    vn(
                      "Original automation bytes preserved in native JSON. No source code was executed.",
                    ));
                },
                onRemove: (Ue) =>
                  di(
                    Kxe(sr.current.book, Ue, {
                      actor: E,
                      readOnly: sr.current.readOnly,
                    }),
                  ),
                onApply: (Ue, Xe, gt) => {
                  const Nt = Yxe(sr.current.book, Ue, {
                    actor: E,
                    scope: Xe,
                    readOnly: sr.current.readOnly,
                    retainSource: gt,
                  });
                  (di(Nt),
                    vn(
                      "Reviewed migration applied in one Undo step; source hashes and run provenance recorded.",
                    ));
                },
                onDownload: (Ue, Xe) =>
                  Jd(
                    new Blob([Uint8Array.from(Ue)], {
                      type: "application/octet-stream",
                    }),
                    Xe,
                  ),
              },
              ze?.sha256 ?? "automation",
            );
          if ($e === "search" || $e === "goToSpecial")
            return i.jsx(
              y2e,
              {
                ...fe,
                ranges: Jt,
                active: xt,
                readOnly: A,
                initialTab: $e === "goToSpecial" ? "special" : "find",
                onWorkbookChange: ua,
                onNavigate: Ts,
                onSelectRanges: Pd,
              },
              $e,
            );
          if ($e === "cellStyle")
            return i.jsx(Nde, { ...fe, ranges: Jt, active: xt });
          if ($e === "themes") return i.jsx(Sde, { ...fe, ranges: Jt });
          if ($e === "conditional") return i.jsx(Rde, { ...fe });
          if ($e === "merge") return i.jsx(Dde, { ...fe });
          if ($e === "numberFormat")
            return i.jsx(M4e, { ...fe, ranges: Jt, active: xt });
          if ($e === "outlines") return i.jsx(j4e, { ...fe });
          if ($e === "comments" || $e === "hyperlink")
            return i.jsx(B4e, {
              ...fe,
              range: xt,
              mode: $e,
              active: xt,
              actor: E,
              onNavigate: Ts,
            });
          if (["pasteSpecial", "fillSeries", "viewSettings"].includes($e))
            return i.jsx(
              y4e,
              {
                ...fe,
                mode: $e,
                clipboard: ku.current?.snapshot,
                onSelectionChange: dr,
              },
              $e,
            );
          if ($e === "sheets")
            return i.jsx(m4e, {
              workbook: H,
              sheetId: be ?? he.id,
              readOnly: A,
              onWorkbookChange: ua,
              onSelectSheet: (Ue) => {
                (Je(null), wi(Ue));
              },
              onClose: fe.onClose,
            });
          if (["workbookProperties", "versions", "template"].includes($e))
            return i.jsx(v4e, {
              mode: $e === "workbookProperties" ? "properties" : $e,
              workbook: H,
              readOnly: A,
              persistent: !!(k || (F && I)),
              checkpoints: (ce?.checkpoints ?? []).filter(
                (Ue) => Ue.workbookId === H.id,
              ),
              onClose: fe.onClose,
              onWorkbookChange: ua,
              onCreateWorkbook: (Ue) => {
                if (A) throw new Error("This workbook is read-only.");
                (Jd(
                  new Blob([Sk(Ue)], { type: "application/json" }),
                  `${Ue.name.replace(/[^a-zA-Z0-9 _.-]/g, "_")}.json`,
                ),
                  vn(
                    "Created an independent workbook download. The template remains open.",
                  ));
              },
              onCheckpoint: (Ue) => {
                if (A) throw new Error("This workbook is read-only.");
                const Xe = Ee.current,
                  gt = sr.current.book;
                return TS(_i(gt), () =>
                  Jc((Nt) => {
                    if (sr.current.book !== gt)
                      throw new Error(
                        "The workbook changed. Save the checkpoint again for the current revision.",
                      );
                    return TS(
                      Se.checkpoint(Ue, le.current.guard, { signal: Nt }),
                      (wt) => {
                        Qc(wt, Xe);
                      },
                    );
                  }),
                );
              },
              onCompare: (Ue) => Se.compareCheckpoint(Ue, sr.current.book),
              onRestore: (Ue) => {
                var wt;
                if (A) throw new Error("This workbook is read-only.");
                if (!le.current)
                  throw new Error(
                    "Save this workbook before restoring a checkpoint.",
                  );
                const Xe = sr.current.book,
                  gt = Ee.current;
                (wt = Ae.current) == null || wt.abort();
                const Nt = new AbortController();
                return (
                  (Ae.current = Nt),
                  Jc((jt) => {
                    if (sr.current.book !== Xe || Nt.signal.aborted)
                      throw new Error(
                        "The workbook changed. Compare the checkpoint again before restoring.",
                      );
                    const rr = () => Nt.abort();
                    jt.addEventListener("abort", rr, { once: true });
                    const Mr = (xr) => {
                        if ((jt.removeEventListener("abort", rr), gt.conflict))
                          throw gt.conflict;
                        if (
                          !gt.active ||
                          gt !== Ee.current ||
                          Nt.signal.aborted ||
                          sr.current.book !== Xe ||
                          kr.current ||
                          sr.current.readOnly
                        )
                          throw new Error(
                            "The editor changed before the restore completed. Your current local changes were retained.",
                          );
                        (di(xr.workbook, true, false, xr),
                          Zo(xr.view),
                          vn("Restored checkpoint as a new revision."),
                          la("saved", xr.workbook));
                      },
                      Br = Se.restoreCheckpoint(Ue, Xe, le.current.guard, E, {
                        signal: Nt.signal,
                      });
                    return Gv(Br)
                      ? Br.then(Mr).finally(() =>
                          jt.removeEventListener("abort", rr),
                        )
                      : Mr(Br);
                  })
                );
              },
            });
          if (["table", "pivot", "query"].includes($e))
            return i.jsx(pfe, { ...fe, mode: $e }, $e);
          if ($e === "pivotExplore") return i.jsx(vfe, { ...fe });
          if (["picture", "shape", "chart"].includes($e))
            return i.jsx(Wz, { ...fe }, $e);
          if (
            [
              "margins",
              "orientation",
              "size",
              "printArea",
              "printTitles",
              "pageBreaks",
              "fit",
              "pageLayout",
            ].includes($e)
          )
            return i.jsx(Bme, { ...fe }, $e);
          if ($e === "functionCompatibility")
            return i.jsx(ixe, {
              workbook: H,
              onClose: () => Je(null),
              onNavigate: Ts,
            });
          if (["function", "names"].includes($e))
            return i.jsx(
              rxe,
              {
                ...fe,
                mode: $e,
                onInsertFormula: (Ue) => {
                  (Je(null),
                    us(Ue),
                    mt(Ue.length),
                    Zt(true),
                    requestAnimationFrame(() => {
                      var Xe, gt;
                      ((Xe = pt.current) == null || Xe.focus(),
                        (gt = pt.current) == null ||
                          gt.setSelectionRange(Ue.length, Ue.length));
                    }));
                },
              },
              $e,
            );
          if ($e === "populationFilter")
            return i.jsx(Xce, {
              workbook: H,
              sheetId: he.id,
              range: Ye,
              column: An,
              definition: Ui?.definition,
              readOnly: A,
              onWorkbookChange: (Ue, Xe) => {
                if (sr.current.book !== H)
                  throw Error(
                    "The workbook changed. Reopen the filter tool before saving a view.",
                  );
                ua(Ue, Xe);
              },
              onClose: fe.onClose,
              onApply: (Ue) => {
                try {
                  (f0(H, he.id, Ue, Qr),
                    fn((Xe) => ({
                      ...Xe,
                      [Xo]: { snapshot: H, definition: Ue },
                    })),
                    jr(""),
                    po(null),
                    jo("full"),
                    Co(0),
                    Je(null),
                    vn(
                      "Filter view applied. Source rows, formulas and history are unchanged.",
                    ));
                } catch (Xe) {
                  hn(
                    Xe instanceof Error
                      ? Xe.message
                      : "Recalculate before filtering.",
                  );
                }
              },
            });
          if ($e === "externalSources")
            return i.jsx(Y2e, {
              workbook: H,
              sheetId: he.id,
              adapter: a,
              reliability: D,
              readOnly: A,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (sr.current.book !== H)
                  throw Error(
                    "The workbook changed. Preview the source again.",
                  );
                ca(Ue, Xe);
              },
            });
          if ($e === "solver" || $e === "statistics")
            return i.jsx(kxe, {
              workbook: H,
              sheetId: he.id,
              range: Ye,
              initialKind: $e,
              readOnly: A,
              allowHidden: true,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (sr.current.book !== H)
                  throw Error(
                    "The workbook changed. Preview the analysis again.",
                  );
                ca(Ue, Xe);
              },
            });
          if ($e === "extensions")
            return i.jsx(qse, {
              workbook: H,
              extensions: e,
              readOnly: A,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (sr.current.book !== H || sr.current.readOnly)
                  throw Error(
                    "The workbook or permissions changed. Review extensions again.",
                  );
                ua(Ue, Xe);
              },
            });
          if ($e === "workbookLinks")
            return i.jsx(Nue, {
              workbook: H,
              readOnly: A,
              approvedSources: W,
              onApprovedSourcesChange: K,
              allowFileSources: n,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (sr.current.book !== H || sr.current.readOnly)
                  throw Error(
                    "The workbook or permissions changed. Preview links again.",
                  );
                ua(Ue, Xe);
              },
            });
          if ($e === "importSources")
            return i.jsx(Fue, {
              workbook: H,
              readOnly: A,
              extractor: o,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (sr.current.book !== H || sr.current.readOnly)
                  throw Error(
                    "The workbook or permissions changed. Preview the source again.",
                  );
                ca(Ue, Xe);
              },
            });
          if ($e === "workbookFormats")
            return i.jsx($ue, {
              workbook: H,
              readOnly: je,
              initialFile: pl,
              onClose: () => {
                (su(undefined), fe.onClose());
              },
              onWorkbookChange: (Ue) => {
                if (sr.current.book !== H || sr.current.readOnly)
                  throw Error(
                    "The workbook or permissions changed. Review the conversion again.",
                  );
                (di(Ue), wi(Ue.sheets.find((Xe) => !Xe.hidden).id));
              },
              onAutomation: async (Ue) => {
                const Xe = await ow(Ue, Yu);
                (at(Xe), su(undefined), Je("automation"));
              },
            });
          if ($e === "reportOutputs")
            return i.jsx(Fxe, {
              workbook: H,
              sheetId: he.id,
              range: Ye,
              filter: Hr?.context.definition,
              readOnly: A,
              allowHidden: true,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (sr.current.book !== H)
                  throw Error(
                    "The workbook changed. Preview the report again.",
                  );
                ua(Ue, Xe);
              },
            });
          if ($e === "dataModel")
            return i.jsx(vce, {
              workbook: H,
              sheetId: he.id,
              range: Ye,
              readOnly: A,
              allowHidden: true,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                if (sr.current.book !== H)
                  throw Error("The workbook changed. Preview the model again.");
                ca(Ue, Xe);
              },
            });
          if ($e === "sort")
            return i.jsx(p4e, {
              workbook: H,
              sheetId: he.id,
              range: Ye,
              readOnly: A,
              onClose: fe.onClose,
              onWorkbookChange: (Ue, Xe) => {
                (ua(Ue, Xe), po(null), jr(""));
              },
            });
          if (
            [
              "combinePopulations",
              "transformationRecipes",
              "consolidation",
              "textToColumns",
              "textCleanup",
              "fieldTransform",
              "duplicates",
            ].includes($e)
          ) {
            const Ue = {
              workbook: H,
              sheetId: he.id,
              range: $e === "duplicates" && !Ye.includes(":") ? Ta(he) : Ye,
              readOnly: A,
              allowHidden: true,
              onClose: fe.onClose,
              onWorkbookChange: ca,
            };
            return $e === "consolidation"
              ? i.jsx(ice, {
                  ...Ue,
                  initialDefinition: he.consolidation?.definition ?? l,
                  approvedWorkbooks: te,
                  onApprovedWorkbooksChange: ae,
                  allowFileSources: h && !!ae,
                })
              : $e === "transformationRecipes"
                ? i.jsx(Hce, { ...Ue })
                : $e === "combinePopulations"
                  ? i.jsx(e4e, { ...Ue })
                  : $e === "textToColumns"
                    ? i.jsx(t4e, { ...Ue })
                    : $e === "textCleanup"
                      ? i.jsx(s4e, { ...Ue })
                      : $e === "fieldTransform"
                        ? i.jsx(l4e, { ...Ue })
                        : i.jsx($z, { ...Ue });
          }
          return $e === "patterns"
            ? i.jsx(Z4e, {
                workbook: H,
                sheetId: he.id,
                range: Ju || (Ye.includes(":") ? Ye : Ta(he)),
                readOnly: A,
                allowHidden: true,
                onClose: fe.onClose,
              })
            : $e === "aging"
              ? i.jsx(Ede, {
                  workbook: H,
                  sheetId: he.id,
                  range: Ju || (Ye.includes(":") ? Ye : Ta(he)),
                  readOnly: A,
                  allowHidden: true,
                  onClose: fe.onClose,
                })
              : $e === "materiality"
                ? i.jsx(P4e, {
                    workbook: H,
                    sheetId: he.id,
                    range: Ju || (Ye.includes(":") ? Ye : Ta(he)),
                    readOnly: A,
                    allowHidden: true,
                    onClose: fe.onClose,
                  })
                : $e === "workbookComparison"
                  ? i.jsx(u2e, {
                      workbook: H,
                      sheetId: he.id,
                      range: Ye.includes(":") ? Ye : Ta(he),
                      readOnly: A,
                      allowHidden: true,
                      onClose: fe.onClose,
                    })
                  : $e === "scenarios" || $e === "dataTable"
                    ? i.jsx($xe, {
                        workbook: H,
                        sheetId: he.id,
                        address: xt,
                        initialKind:
                          $e === "scenarios" ? "scenarios" : "one-variable",
                        readOnly: A,
                        allowHidden: true,
                        onClose: fe.onClose,
                        onWorkbookChange: ca,
                      })
                    : $e === "goalSeek"
                      ? i.jsx(mxe, {
                          workbook: H,
                          sheetId: he.id,
                          address: xt,
                          readOnly: A,
                          allowHidden: true,
                          onClose: fe.onClose,
                          onWorkbookChange: ua,
                        })
                      : $e === "calculation"
                        ? i.jsx(Zce, {
                            workbook: H,
                            readOnly: A,
                            onClose: fe.onClose,
                            onWorkbookChange: ua,
                            onCalculated: _f,
                            onCalculate: wa,
                            onCancelCalculation: bl.cancel,
                            onNavigate: Ts,
                          })
                        : $e === "formulaWatches"
                          ? i.jsx(fxe, {
                              workbook: H,
                              sheetId: he.id,
                              range: Ye,
                              ranges: Jt,
                              readOnly: A,
                              allowHidden: true,
                              onClose: fe.onClose,
                              onWorkbookChange: ua,
                              onNavigate: Ts,
                            })
                          : $e === "evaluateFormula"
                            ? i.jsx(sxe, {
                                workbook: H,
                                sheetId: he.id,
                                address: xt,
                                allowHidden: true,
                                onClose: fe.onClose,
                                onNavigate: Ts,
                              })
                            : $e === "tracePrecedents" ||
                                $e === "traceDependents"
                              ? i.jsx(lxe, {
                                  workbook: H,
                                  sheetId: he.id,
                                  address: xt,
                                  initialDirection:
                                    $e === "tracePrecedents"
                                      ? "precedents"
                                      : "dependents",
                                  allowHidden: true,
                                  onClose: fe.onClose,
                                  onNavigate: Ts,
                                })
                              : i.jsx(sfe, { ...fe, onNavigate: Ts }, $e);
        })(),
      oi &&
        i.jsx(zxe, {
          source: oi,
          blocked:
            oi.snapshot !== H
              ? "The workbook changed. Cancel and review again for the current revision."
              : "bytes" in oi && je
                ? "This workbook is read-only. Import is unavailable."
                : undefined,
          onClose: () => xs(null),
          onReport: (fe) =>
            Jd(
              new Blob([JSON.stringify(fe, null, 2)], {
                type: "application/json",
              }),
              `${fe.fileName}.schema.json`,
            ),
          onImport: (fe, Oe) => {
            if (sr.current.book !== oi.snapshot)
              throw new Error("The workbook changed. Review the file again.");
            if (sr.current.readOnly)
              throw new Error("This workbook is read-only.");
            (di(fe),
              wi(fe.sheets[0].id),
              vn(
                `Imported ${Oe.records} records using the reviewed schema. Formula-like source text remains literal.`,
              ),
              xs(null));
          },
          onExport: (fe, Oe) => {
            if (sr.current.book !== oi.snapshot)
              throw new Error("The workbook changed. Review the export again.");
            if (Oe.population && !("bytes" in oi)) {
              const Be = V$(sr.current.book, oi.sheetId, Oe.options);
              if (
                JSON.stringify(Be.report.population) !==
                  JSON.stringify(Oe.population) ||
                Be.bytes.length !== fe.length ||
                Be.bytes.some((Ue, Xe) => Ue !== fe[Xe])
              )
                throw Error(
                  "Calculated values or population membership changed. Preview the export again.",
                );
            }
            (Jd(
              new Blob([Uint8Array.from(fe)], {
                type: "application/octet-stream",
              }),
              Oe.fileName,
            ),
              vn(
                `Delimited data exported; ${Oe.safeTextChanges} formula-like text values neutralized under the reviewed safe-export policy.`,
              ),
              xs(null));
          },
        }),
      vo &&
        i.jsx(X$, {
          report: vo.report,
          inspectOnly: vo.inspectOnly,
          onOriginal: () =>
            On(() => {
              const fe = WR(vo.imported ?? vo.snapshot);
              Jd(
                new Blob([new Uint8Array(fe.bytes)], {
                  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                }),
                fe.fileName,
              );
            }),
          blocked:
            vo.snapshot !== H
              ? "The workbook changed. Cancel and review the file again for the current revision."
              : vo.imported && je
                ? "This workbook is read-only. Import is unavailable."
                : undefined,
          onClose: () => Ml(null),
          onReport: () =>
            Jd(
              new Blob([JSON.stringify(vo.report, null, 2)], {
                type: "application/json",
              }),
              `${vo.fileName}.compatibility.json`,
            ),
          onConfirm: () =>
            On(() => {
              if (sr.current.book !== vo.snapshot)
                throw new Error("The workbook changed. Review the file again.");
              if (vo.imported) {
                if (sr.current.readOnly)
                  throw new Error("This workbook is read-only.");
                (di(vo.imported),
                  wi(vo.imported.sheets.find((fe) => !fe.hidden).id),
                  vn(
                    "Derived XLSX workbook opened. The exact original is available from File → Original XLSX and compatibility. Save native JSON to retain it across sessions.",
                  ));
              } else
                vo.blob &&
                  (Jd(vo.blob, vo.fileName),
                  vn(
                    "Derived XLSX downloaded. Keep a native JSON backup for full application metadata.",
                  ));
              Ml(null);
            }),
        }),
      i.jsxs("div", {
        className: "lg-main",
        children: [
          i.jsx("input", {
            type: "file",
            "aria-label": T("Import workbook file"),
            ref: yl,
            accept:
              ".xlsx,.xltx,.xls,.xlsb,.xlsm,.xltm,.ods,.csv,.tsv,.json,.bas,.vba,.cls,.ts,.osts,.xlam",
            hidden: true,
            onChange: (fe) => {
              return void Ns(fe.target.files?.[0]);
            },
            lang: N("Import workbook file"),
            "data-ui-label": "Import workbook file",
          }),
          i.jsx(Hde, {
            onTool: Pn,
            sheetProtected: !!((Y = he.protection) != null && Y.enabled),
            allowProtectionChange: !A,
            workbookStructureProtected: !!(
              (re = H.structureProtection) != null && re.enabled
            ),
            readOnly: je,
            busy: go,
            canUndo: ns.length > 0,
            canRedo: Qi.length > 0,
            cell: ur,
            showFormulas: Gi,
            gridlines: ec,
            headings: vl,
            freezeRows: Yc,
            zoom: Ao,
            onFormat: Vi,
            onUndo: () => fu(),
            onRedo: () => fu(true),
            onImport: () => {
              return yl.current?.click();
            },
            onExport: (fe) => void hu(fe),
            onNew: () => oc(t5()),
            onDemo: () => oc(FI()),
            onRename: () => Ia({ kind: "rename", value: H.name }),
            onAddSheet: () =>
              Ia({ kind: "sheet", value: `Sheet${H.sheets.length + 1}` }),
            onAddRow: () => Pn("insertRow"),
            onAddColumn: () => Pn("insertColumn"),
            onAutoFitColumns: () => Ko("column"),
            onAutoFitRows: () => Ko("row"),
            onAutosum: () => On(() => To([jQ(H, he.id, xt)], "AutoSum")),
            onFillDown: () =>
              On(() => {
                (rt(), So(Qy(H, $2(H, he.id, Ye, "down")), "Fill down"));
              }),
            onClear: () => On(Ho),
            onCopy: () => void Ru(false),
            onCut: () => void Ru(true),
            onPaste: () => void sc(),
            onFind: () => Pl((fe) => !fe),
            onSort: (fe) => {
              (po(
                fe
                  ? { col: _t(xt).col, direction: fe === "asc" ? 1 : -1 }
                  : null,
              ),
                Co(0),
                to.current && (to.current.scrollTop = 0));
            },
            onToggleFilter: () => {
              (Hc((fe) => !fe), jr(""));
            },
            onShowFormulas: () => ti((fe) => !fe),
            onGridlines: () => cm((fe) => !fe),
            onHeadings: () => k0((fe) => !fe),
            onFreeze: ac,
            onZoom: bo,
            onPanel: Pr,
            onNote: () =>
              Ia({
                kind: "note",
                value: ur?.note ?? "",
              }),
            onReviewed: () =>
              Vi({ tickmark: ur != null && ur.tickmark ? "" : "✓" }),
          }),
          i.jsx(nce, {
            workbook: H,
            sheetId: he.id,
            options: { approvedWorkbooks: te },
            onReview: () => Pn("consolidation"),
          }),
          i.jsx(Hse, {
            workbook: H,
            extensions: e,
            onReview: () => Pn("extensions"),
          }),
          i.jsx(Tue, {
            workbook: H,
            sources: W,
            onReview: () => Pn("workbookLinks"),
          }),
          i.jsx(K2e, {
            workbook: H,
            sheetId: he.id,
            onReview: () => Pn("externalSources"),
          }),
          i.jsx(Sxe, {
            workbook: H,
            sheetId: he.id,
            onReview: () => {
              return Pn(
                he.analysisReport?.request.kind === "solver"
                  ? "solver"
                  : "statistics",
              );
            },
          }),
          i.jsx(bce, {
            workbook: H,
            sheetId: he.id,
            onReview: () => Pn("dataModel"),
          }),
          (_a.report.settings.mode === "manual" ||
            _a.report.pending > 0 ||
            _a.report.execution?.state === "timeout" ||
            _a.report.execution?.state === "failed") &&
            i.jsxs("div", {
              className: "lg-notice",
              role: "status",
              "aria-label": T("Calculation status"),
              lang: N("Calculation status"),
              "data-ui-label": "Calculation status",
              children: [
                i.jsxs("span", {
                  children: [
                    _a.report.settings.mode === "manual"
                      ? "Manual calculation"
                      : "Calculation",
                    " ",
                    "· ",
                    _a.report.pending,
                    " ",
                    i.jsx(M, { children: "pending formulas" }),
                    _a.report.execution?.state === "timeout" ||
                    _a.report.execution?.state === "failed"
                      ? ` · ${_a.report.execution.message}`
                      : "",
                  ],
                }),
                i.jsx("button", {
                  onClick: () => Pn("calculateNow"),
                  children: i.jsx(M, { children: "Calculate Now" }),
                }),
                _a.report.execution?.mode === "worker" &&
                  _a.report.execution.state === "pending" &&
                  i.jsx("button", {
                    onClick: bl.cancel,
                    children: i.jsx(M, { children: "Cancel calculation" }),
                  }),
                i.jsx("button", {
                  onClick: () => Pn("calculation"),
                  children: i.jsx(M, { children: "Calculation Options" }),
                }),
              ],
            }),
          (qs || Hs) &&
            i.jsxs("div", {
              className: `lg-notice ${qs ? "error" : ""}`,
              role: qs ? "alert" : "status",
              children: [
                i.jsx(M, { children: qs || Hs }),
                i.jsx("button", {
                  "aria-label": T("Dismiss notification"),
                  onClick: () => {
                    (hn(""), vn(""));
                  },
                  lang: N("Dismiss notification"),
                  "data-ui-label": "Dismiss notification",
                  children: i.jsx(X, { size: 14 }),
                }),
              ],
            }),
          kr.current &&
            kr.current.locale !== R.entryLocale &&
            i.jsx("p", {
              role: "status",
              className: "lg-notice",
              children: i.jsx(M, {
                children:
                  "The current edit keeps its original numeric locale until you apply or cancel it.",
              }),
            }),
          Te &&
            i.jsxs("div", {
              className: "lg-notice",
              role: "status",
              "aria-label": T("Unsaved workbook changes"),
              lang: N("Unsaved workbook changes"),
              "data-ui-label": "Unsaved workbook changes",
              children: [
                i.jsx("span", {
                  children: i.jsx(M, {
                    children: "Your local changes are retained in this editor.",
                  }),
                }),
                Te instanceof Oo && Te.code === "conflict"
                  ? i.jsx("button", {
                      onClick: () => Re(true),
                      children: i.jsx(M, { children: "Review save conflict" }),
                    })
                  : i.jsx("button", {
                      disabled: A || Me > 0,
                      onClick: Uo,
                      children: i.jsx(M, { children: "Retry saving" }),
                    }),
                Me > 0 &&
                  i.jsx("span", {
                    children: i.jsx(M, {
                      children: "Waiting for save acknowledgement",
                    }),
                  }),
                i.jsx("button", {
                  onClick: Tn,
                  children: i.jsx(M, { children: "Export my changes" }),
                }),
              ],
            }),
          ht &&
            i.jsx(P2e, {
              workbook: H,
              loadSaved: () => Se.load(),
              onExport: Tn,
              onReload: Aa,
              onClose: () => Re(false),
            }),
          i.jsx(Cde, {}),
          i.jsxs("div", {
            className: "lg-workarea",
            children: [
              i.jsxs("section", {
                className: "lg-workbook",
                "aria-label": T("Spreadsheet workbook"),
                lang: N("Spreadsheet workbook"),
                "data-ui-label": "Spreadsheet workbook",
                children: [
                  i.jsxs("div", {
                    className: "lg-formula-bar",
                    children: [
                      i.jsx("input", {
                        className: "lg-name-box",
                        "aria-label": T("Selected range"),
                        value: Le,
                        onChange: (fe) => et(fe.target.value),
                        onBlur: Ud,
                        onKeyDown: (fe) => {
                          var Oe;
                          fe.key === "Enter" &&
                            (Ud(), (Oe = to.current) == null || Oe.focus());
                        },
                        lang: N("Selected range"),
                        "data-ui-label": "Selected range",
                      }),
                      i.jsx("button", {
                        className: "lg-formula-cancel",
                        "aria-label": T("Cancel cell edit"),
                        onPointerDown: (fe) => fe.preventDefault(),
                        onMouseDown: (fe) => fe.preventDefault(),
                        onClick: Da,
                        lang: N("Cancel cell edit"),
                        "data-ui-label": "Cancel cell edit",
                        children: i.jsx(X, { size: 15 }),
                      }),
                      i.jsx("button", {
                        "aria-label": T("Apply cell value"),
                        disabled: Ss(),
                        onPointerDown: (fe) => fe.preventDefault(),
                        onMouseDown: (fe) => fe.preventDefault(),
                        onClick: () => {
                          ui() && pi();
                        },
                        lang: N("Apply cell value"),
                        "data-ui-label": "Apply cell value",
                        children: i.jsx(Check, { size: 15 }),
                      }),
                      i.jsx("span", {
                        className: "lg-fx",
                        children: i.jsx(M, { children: "ƒx" }),
                      }),
                      i.jsx("input", {
                        ref: pt,
                        "aria-label": T("Formula bar"),
                        readOnly: Ss(),
                        value: st,
                        "aria-describedby":
                          (Et = kr.current) != null && Et.value.startsWith("=")
                            ? `${$}-formula-help`
                            : undefined,
                        onFocus: (fe) => {
                          (!kr.current && st.startsWith("=") && !je && us(st),
                            Zt(true),
                            ks(fe.currentTarget));
                        },
                        onSelect: (fe) => {
                          (ks(fe.currentTarget),
                            mt(fe.currentTarget.selectionStart ?? st.length));
                        },
                        onChange: (fe) => {
                          (us(fe.target.value),
                            ks(fe.currentTarget, true),
                            mt(
                              fe.target.selectionStart ??
                                fe.target.value.length,
                            ));
                        },
                        onBlur: (fe) => {
                          dH(fe.currentTarget, fe.relatedTarget) ||
                            (Zt(false), ui());
                        },
                        onKeyDown: (fe) => {
                          var Oe, Be;
                          if (!fe.nativeEvent.isComposing) {
                            if (fe.key === "Escape") {
                              if (
                                !Ke.current.span &&
                                st.length > 1 &&
                                (Oe = Fe.current) != null &&
                                Oe.handleKeyDown(fe)
                              )
                                return;
                              (fe.preventDefault(), Da());
                              return;
                            }
                            qa(fe) ||
                              ((Be = Fe.current) != null &&
                                Be.handleKeyDown(fe)) ||
                              ((fe.key === "Enter" || fe.key === "Tab") &&
                                (fe.preventDefault(),
                                ui() && pu(fe.key, fe.shiftKey)));
                          }
                        },
                        lang: N("Formula bar"),
                        "data-ui-label": "Formula bar",
                      }),
                      i.jsx(fH, {
                        ref: Fe,
                        onExit: () => {
                          (Zt(false), ui());
                        },
                        value: st,
                        caret: nt,
                        workbook: H,
                        sheetId: he.id,
                        inputRef: pt,
                        disabled: Ss() || !Wt,
                        onAccept: (fe, Oe) => {
                          (us(fe),
                            mt(Oe),
                            requestAnimationFrame(() => {
                              var Be, Ue;
                              ((Be = pt.current) == null || Be.focus(),
                                (Ue = pt.current) == null ||
                                  Ue.setSelectionRange(Oe, Oe));
                            }));
                        },
                      }),
                    ],
                  }),
                  kr.current?.value.startsWith("=") &&
                    i.jsxs("div", {
                      className: "lg-formula-entry-help",
                      id: `${$}-formula-help`,
                      lang: "en",
                      children: [
                        i.jsxs("span", {
                          children: [
                            "Editing",
                            " ",
                            H.sheets.find((fe) => {
                              return fe.id === kr.current?.sheetId;
                            })?.name,
                            "!",
                            kr.current?.address,
                            " ·",
                            " ",
                            Qe ? "Point mode" : "Edit mode",
                            " · Select cells to insert a reference. F2 switches arrow keys; F4 locks references. Enter applies; Esc cancels.",
                          ],
                        }),
                        i.jsx("span", {
                          role: "status",
                          "aria-live": "polite",
                          "aria-atomic": "true",
                          "aria-label": "Formula reference",
                          children: Tt
                            ? `Reference: ${H.sheets.find((fe) => fe.id === Tt.sheetId)?.name}!${Tt.range}`
                            : "Formula not yet applied",
                        }),
                      ],
                    }),
                  DB(H, he.id, xt).map((fe) =>
                    i.jsxs(
                      "div",
                      {
                        className: "lg-view-note",
                        "aria-label": T("Validation input guidance"),
                        lang: N("Validation input guidance"),
                        "data-ui-label": "Validation input guidance",
                        children: [
                          fe.showInputMessage !== false &&
                            (fe.inputTitle || fe.inputMessage) &&
                            i.jsxs("span", {
                              children: [
                                i.jsx("strong", { children: fe.inputTitle }),
                                " ",
                                fe.inputMessage,
                                " ",
                              ],
                            }),
                          fe.kind === "list" &&
                            fe.showDropdown !== false &&
                            i.jsxs("label", {
                              children: [
                                i.jsx(M, { children: "Choose a value for " }),
                                xt,
                                " ",
                                i.jsxs("select", {
                                  "aria-label": `Validation dropdown for ${xt}`,
                                  disabled: je,
                                  value: "",
                                  onChange: (Oe) => {
                                    const Be = fe.list[Number(Oe.target.value)];
                                    Be !== undefined &&
                                      On(() => {
                                        (Da(),
                                          To(
                                            [
                                              {
                                                sheetId: he.id,
                                                address: xt,
                                                cell: {
                                                  ...he.cells[xt],
                                                  raw: Be,
                                                  literal: true,
                                                },
                                              },
                                            ],
                                            "Choose validation list value",
                                          ));
                                      });
                                  },
                                  children: [
                                    i.jsx("option", {
                                      value: "",
                                      disabled: true,
                                      children: "Select…",
                                    }),
                                    fe.list.map((Oe, Be) =>
                                      i.jsx(
                                        "option",
                                        { value: Be, children: Oe },
                                        Be,
                                      ),
                                    ),
                                  ],
                                }),
                              ],
                            }),
                        ],
                      },
                      fe.id,
                    ),
                  ),
                  Qr.getSpillDiagnostic(he.id, xt) &&
                    i.jsxs("div", {
                      className: "lg-format-warning",
                      role: "status",
                      "aria-label": T("Array spill diagnostic"),
                      lang: N("Array spill diagnostic"),
                      "data-ui-label": "Array spill diagnostic",
                      children: [
                        xt,
                        ": ",
                        Qr.getSpillDiagnostic(he.id, xt).message,
                      ],
                    }),
                  (Qr.getSpillAnchor(he.id, xt) ||
                    Qr.getSpillRange(he.id, xt)) &&
                    i.jsxs("div", {
                      className: "lg-view-note",
                      role: "status",
                      "aria-label": T("Dynamic array range"),
                      lang: N("Dynamic array range"),
                      "data-ui-label": "Dynamic array range",
                      children: [
                        i.jsx(M, { children: "Dynamic array" }),
                        " ",
                        Qr.getSpillRange(
                          he.id,
                          Qr.getSpillAnchor(he.id, xt) ?? xt,
                        ),
                        ".",
                        Qr.getSpillAnchor(he.id, xt)
                          ? ` Edit the formula in ${Qr.getSpillAnchor(he.id, xt)}; this cell is a calculated result.`
                          : " Results resize when source data changes.",
                      ],
                    }),
                  yt.length > 0 &&
                    i.jsxs("div", {
                      className: "lg-format-warning",
                      role: "status",
                      "aria-label": T("Named formula diagnostic"),
                      lang: N("Named formula diagnostic"),
                      "data-ui-label": "Named formula diagnostic",
                      children: [
                        xt,
                        ":",
                        " ",
                        yt.map((fe) => fe.message).join(" "),
                      ],
                    }),
                  $t.length > 0 &&
                    i.jsxs("div", {
                      className: "lg-format-warning",
                      role: "status",
                      "aria-label": T("Formula reference diagnostic"),
                      lang: N("Formula reference diagnostic"),
                      "data-ui-label": "Formula reference diagnostic",
                      children: [
                        xt,
                        ":",
                        " ",
                        $t.map((fe) => fe.message).join(" "),
                      ],
                    }),
                  ot.length > 0 &&
                    i.jsxs("div", {
                      className: "lg-format-warning",
                      role: "status",
                      children: [
                        xt,
                        ":",
                        " ",
                        ot.map((fe) => fe.message).join(" "),
                        " ",
                        i.jsx("button", {
                          type: "button",
                          onClick: () => Pn("conditional"),
                          children: i.jsx(M, {
                            children: "Review conditional rules",
                          }),
                        }),
                      ],
                    }),
                  (KS(ur?.style) || H.styleCatalog?.missingTheme) &&
                    i.jsxs("div", {
                      className: "lg-format-warning",
                      role: "status",
                      children: [
                        KS(ur?.style) || bB(H).join(" "),
                        i.jsx("button", {
                          onClick: () => Pn("themes"),
                          children: i.jsx(M, {
                            children: "Review themes and styles",
                          }),
                        }),
                      ],
                    }),
                  SE(he.cells[xt]) &&
                    i.jsxs("div", {
                      className: "lg-view-note",
                      role: "status",
                      children: [xt, ": ", SE(he.cells[xt])],
                    }),
                  Ui &&
                    i.jsxs("div", {
                      className: "lg-population-context",
                      role: "status",
                      "aria-label": T("Population filter context"),
                      lang: N("Population filter context"),
                      "data-ui-label": "Population filter context",
                      children: [
                        i.jsx(Filter, { size: 14 }),
                        i.jsx("strong", {
                          children: Bd
                            ? "Filter paused — workbook changed"
                            : Kc.error
                              ? "Filter unavailable — source rows shown"
                              : `${Hr?.context.visibleRows} visible / ${Hr?.context.totalRows} total data rows`,
                        }),
                        Hr &&
                          i.jsxs("span", {
                            children: [
                              Hr.context.range,
                              " ·",
                              " ",
                              Hr.context.filteredRows,
                              " ",
                              i.jsx(M, { children: "filtered out ·" }),
                              " ",
                              Hr.context.hiddenMatchingRows,
                              " ",
                              i.jsx(M, {
                                children: "matching rows manually hidden",
                              }),
                            ],
                          }),
                        i.jsx("button", {
                          onClick: () => Pn("populationFilter"),
                          children: Bd ? "Review and reapply" : "Edit filters",
                        }),
                        i.jsx("button", {
                          onClick: () => {
                            (fn((fe) => {
                              const Oe = { ...fe };
                              return (delete Oe[Xo], Oe);
                            }),
                              jo("full"));
                          },
                          children: i.jsx(M, {
                            children: "Clear population filters",
                          }),
                        }),
                        Hr &&
                          i.jsx("button", {
                            onClick: () =>
                              Jd(
                                new Blob(
                                  [JSON.stringify(Hr.context, null, 2)],
                                  { type: "application/json" },
                                ),
                                "Population-filter-context.json",
                              ),
                            children: i.jsx(M, {
                              children: "Save filter context",
                            }),
                          }),
                        i.jsxs("details", {
                          children: [
                            i.jsx("summary", {
                              children: i.jsx(M, {
                                children: "Population scope",
                              }),
                            }),
                            i.jsx("p", {
                              children: i.jsx(M, {
                                children:
                                  "Column filters intersect; advanced groups use OR. These are session views; named criteria can be saved in the native workbook through Edit filters. Native JSON, XLSX, print, other analysis tools and AI selections include full source data in their chosen ranges. CSV/TSV and Audit Checks offer full or visible scope. View filtering does not change SUM or ordinary formulas; SUBTOTAL/AGGREGATE retain their existing worksheet visibility semantics.",
                              }),
                            }),
                          ],
                        }),
                        Kc.error &&
                          i.jsx("span", { role: "alert", children: Kc.error }),
                      ],
                    }),
                  Hr?.context.visibleRows === 0 &&
                    i.jsxs("div", {
                      className: "lg-population-empty",
                      role: "status",
                      children: [
                        i.jsx(M, {
                          children: "No matching visible records. All",
                        }),
                        " ",
                        Hr.context.totalRows,
                        " ",
                        i.jsx(M, {
                          children:
                            "source data rows are retained. Clear population filters to show them again; manually hidden rows remain hidden.",
                        }),
                      ],
                    }),
                  wh &&
                    i.jsxs("div", {
                      className: "lg-find-bar",
                      children: [
                        i.jsx(Filter, { size: 14 }),
                        i.jsx("input", {
                          "aria-label": T("Filter rows"),
                          placeholder: T("Filter rows (first row is a header)"),
                          value: zr,
                          onChange: (fe) => {
                            (jr(fe.target.value),
                              fn((Oe) => {
                                const Be = { ...Oe };
                                return (delete Be[Xo], Be);
                              }),
                              Co(0),
                              to.current && (to.current.scrollTop = 0));
                          },
                          lang: N("Filter rows"),
                          "data-ui-label": "Filter rows",
                        }),
                        i.jsx("button", {
                          "aria-label": T("Close filter"),
                          onClick: () => {
                            (Hc(false), jr(""), pi());
                          },
                          lang: N("Close filter"),
                          "data-ui-label": "Close filter",
                          children: i.jsx(X, { size: 14 }),
                        }),
                      ],
                    }),
                  i.jsx(dfe, {
                    workbook: H,
                    sheetId: oe,
                    onOpen: () => Pn("pivot"),
                    onExplore: () => Pn("pivotExplore"),
                  }),
                  S0 &&
                    i.jsxs("div", {
                      className: "lg-find-bar",
                      children: [
                        i.jsx(Search, { size: 14 }),
                        i.jsx("input", {
                          ref: Ld,
                          "aria-label": T("Find text"),
                          placeholder: T("Find"),
                          value: as,
                          onChange: (fe) => jd(fe.target.value),
                          onKeyDown: (fe) => {
                            fe.key === "Enter" && lc();
                          },
                          lang: N("Find text"),
                          "data-ui-label": "Find text",
                        }),
                        i.jsx("button", {
                          onClick: lc,
                          children: i.jsx(M, { children: "Find Next" }),
                        }),
                        i.jsx("input", {
                          "aria-label": T("Replace text"),
                          placeholder: T("Replace with"),
                          value: uu,
                          onChange: (fe) => sa(fe.target.value),
                          lang: N("Replace text"),
                          "data-ui-label": "Replace text",
                        }),
                        i.jsx("button", {
                          disabled: je,
                          onClick: Ro,
                          children: i.jsx(M, {
                            children: "Replace in selection",
                          }),
                        }),
                        i.jsx("button", {
                          "aria-label": T("Close find"),
                          onClick: () => {
                            (Pl(false), pi());
                          },
                          lang: N("Close find"),
                          "data-ui-label": "Close find",
                          children: i.jsx(X, { size: 14 }),
                        }),
                      ],
                    }),
                  Fo &&
                    i.jsxs("div", {
                      className: "lg-view-note",
                      role: "status",
                      children: [
                        i.jsx(M, { children: "Fill preview · " }),
                        Fo.targetRange,
                        " ",
                        i.jsx(M, {
                          children:
                            "· Release to review, or press Escape to cancel",
                        }),
                      ],
                    }),
                  zr &&
                    ko === 0 &&
                    i.jsx("div", {
                      className: "lg-population-empty",
                      role: "status",
                      children: i.jsx(M, {
                        children:
                          "No rows match this search. Source data is unchanged; clear the search to restore the view.",
                      }),
                    }),
                  (zr || kn) &&
                    i.jsxs("div", {
                      className: "lg-view-note",
                      children: [
                        i.jsx(Filter, { size: 12 }),
                        " ",
                        i.jsx(M, {
                          children:
                            "View only · source row numbers and formula references stay fixed",
                        }),
                        zr && ` · ${ko} matching rows`,
                      ],
                    }),
                  Bo &&
                    i.jsxs("div", {
                      className: "lg-view-note",
                      role: "status",
                      children: [
                        i.jsx(M, { children: "Cut ready:" }),
                        " ",
                        H.sheets.find((fe) => fe.id === Bo.snapshot.sheetId)
                          ?.name ?? "Deleted worksheet",
                        "!",
                        Bo.snapshot.range,
                        " ",
                        i.jsx(M, { children: "· Paste to review the move." }),
                        " ",
                        i.jsx("button", {
                          type: "button",
                          onClick: Ul,
                          children: i.jsx(M, { children: "Cancel cut" }),
                        }),
                      ],
                    }),
                  i.jsx(C2e, {
                    split: Or,
                    snapshot: H,
                    readOnly: je,
                    active: an.activeId,
                    onResize: (fe) =>
                      On(() =>
                        So(
                          JS(sr.current.book, he.id, fe),
                          "Resize split panes",
                        ),
                      ),
                    renderPane: Al,
                  }),
                  i.jsxs("div", {
                    className: "lg-sheet-tabs",
                    "aria-label": T("Worksheet navigation"),
                    role: "group",
                    lang: N("Worksheet navigation"),
                    "data-ui-label": "Worksheet navigation",
                    children: [
                      i.jsxs("div", {
                        children: [
                          H.sheets
                            .filter((fe) => !fe.hidden)
                            .map((fe) =>
                              i.jsx(
                                "button",
                                {
                                  "aria-current":
                                    fe.id === he.id ? "true" : undefined,
                                  className: fe.id === he.id ? "active" : "",
                                  "data-sheet-id": fe.id,
                                  tabIndex: fe.id === he.id ? 0 : -1,
                                  onPointerDown: (Oe) => {
                                    var Be;
                                    (Be = kr.current) != null &&
                                      Be.value.startsWith("=") &&
                                      Oe.preventDefault();
                                  },
                                  onMouseDown: (Oe) => {
                                    var Be;
                                    (Be = kr.current) != null &&
                                      Be.value.startsWith("=") &&
                                      Oe.preventDefault();
                                  },
                                  onKeyDown: (Oe) => {
                                    if (
                                      ![
                                        "ArrowLeft",
                                        "ArrowRight",
                                        "Home",
                                        "End",
                                      ].includes(Oe.key) ||
                                      Oe.altKey ||
                                      Oe.ctrlKey ||
                                      Oe.metaKey
                                    )
                                      return;
                                    Oe.preventDefault();
                                    const Be = H.sheets.filter(
                                        (gt) => !gt.hidden,
                                      ),
                                      Ue = Be.findIndex(
                                        (gt) => gt.id === fe.id,
                                      ),
                                      Xe =
                                        Oe.key === "Home"
                                          ? Be[0]
                                          : Oe.key === "End"
                                            ? Be.at(-1)
                                            : Be[
                                                (Ue +
                                                  (Oe.key === "ArrowRight"
                                                    ? 1
                                                    : Be.length - 1)) %
                                                  Be.length
                                              ];
                                    (wi(Xe.id),
                                      requestAnimationFrame(() => {
                                        return Array.from(
                                          lt.current?.querySelectorAll(
                                            "[data-sheet-id]",
                                          ) ?? [],
                                        )
                                          .find(
                                            (wt) =>
                                              wt.dataset.sheetId === Xe.id,
                                          )
                                          ?.focus();
                                      }));
                                  },
                                  onClick: () => wi(fe.id),
                                  style: fe.tabColor
                                    ? {
                                        boxShadow: `inset 0 -3px ${fe.tabColor}`,
                                      }
                                    : undefined,
                                  onDoubleClick: () => {
                                    ui() && (ft(fe.id), Je("sheets"));
                                  },
                                  onContextMenu: (Oe) => {
                                    (Oe.preventDefault(),
                                      ui() && (ft(fe.id), Je("sheets")));
                                  },
                                  children: fe.name,
                                },
                                fe.id,
                              ),
                            ),
                          i.jsx("button", {
                            className: "lg-add-sheet",
                            "aria-label": T("Add worksheet"),
                            disabled:
                              A ||
                              !!(
                                (Xr = H.structureProtection) != null &&
                                Xr.enabled
                              ),
                            onClick: () =>
                              Ia({
                                kind: "sheet",
                                value: `Sheet ${H.sheets.length + 1}`,
                              }),
                            lang: N("Add worksheet"),
                            "data-ui-label": "Add worksheet",
                            children: i.jsx(Plus, { size: 16 }),
                          }),
                          i.jsx("button", {
                            "aria-label": T("Manage worksheets"),
                            title: T("Manage worksheets"),
                            onClick: () => Pn("sheets"),
                            lang: N("Manage worksheets"),
                            "data-ui-label": "Manage worksheets",
                            children: i.jsx(Layers, { size: 16 }),
                          }),
                        ],
                      }),
                      i.jsx("span", {
                        className: "lg-book-name",
                        title: T("Workbook name"),
                        lang: N("Workbook name"),
                        children: H.name,
                      }),
                    ],
                  }),
                  i.jsxs("footer", {
                    role: "group",
                    className: "lg-grid-footer",
                    "aria-label": T("Selection summary"),
                    lang: N("Selection summary"),
                    "data-ui-label": "Selection summary",
                    children: [
                      i.jsxs("span", {
                        children: [
                          i.jsx("span", { className: "lg-ready-dot" }),
                          i.jsx(M, { children: go ? "Working…" : "Ready" }),
                          i.jsx("button", {
                            type: "button",
                            onClick: () => _(true),
                            children: i.jsx(M, { children: "Workbook health" }),
                          }),
                          go &&
                            ba.current &&
                            i.jsx("button", {
                              onClick: () => {
                                return ba.current?.abort();
                              },
                              children: i.jsx(M, {
                                children: "Cancel file processing",
                              }),
                            }),
                          i.jsx("span", { className: "lg-footer-separator" }),
                          i.jsxs("span", {
                            title: T(
                              "Selection count, numeric count, sum and average include hidden and filtered cells within the selected range.",
                            ),
                            lang: N(
                              "Selection count, numeric count, sum and average include hidden and filtered cells within the selected range.",
                            ),
                            children: [
                              i.jsx(M, { children: "Count: " }),
                              ie,
                              " ",
                              St.length > 0 ? `· ${Jt.length} areas ` : "",
                              " ",
                              i.jsx(M, { children: "· Numbers: " }),
                              er.length,
                              i.jsxs("span", {
                                className: "lg-selection-inclusion",
                                children: [
                                  " ",
                                  i.jsx(M, {
                                    children:
                                      "· Includes hidden and filtered cells",
                                  }),
                                ],
                              }),
                            ],
                          }),
                          je ? " · Read only" : "",
                        ],
                      }),
                      i.jsxs("div", {
                        children: [
                          i.jsxs("span", {
                            children: [
                              i.jsx(M, { children: "Average" }),
                              " ",
                              i.jsx("b", {
                                children: Af(er.length ? Gt / er.length : 0),
                              }),
                            ],
                          }),
                          i.jsxs("span", {
                            children: [
                              i.jsx(M, { children: "Sum" }),
                              " ",
                              i.jsx("b", {
                                children: Af(Gt, { raw: 0, format: "number" }),
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        className: "lg-zoom",
                        children: [
                          i.jsx("button", {
                            "aria-label": T("Zoom out"),
                            onClick: () => bo(Math.max(50, Ao - 10)),
                            lang: N("Zoom out"),
                            "data-ui-label": "Zoom out",
                            children: "−",
                          }),
                          i.jsx("input", {
                            "aria-label": T("Zoom"),
                            type: "range",
                            min: "50",
                            max: "200",
                            step: "10",
                            value: Ao,
                            onChange: (fe) => bo(Number(fe.target.value)),
                            lang: N("Zoom"),
                            "data-ui-label": "Zoom",
                          }),
                          i.jsx("button", {
                            "aria-label": T("Zoom in"),
                            onClick: () => bo(Math.min(200, Ao + 10)),
                            lang: N("Zoom in"),
                            "data-ui-label": "Zoom in",
                            children: "+",
                          }),
                          i.jsxs("span", { children: [Ao, "%"] }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              Lr &&
                i.jsxs("section", {
                  "aria-label":
                    Lr === "assistant"
                      ? "Workbook assistant panel"
                      : "Workbook side panel",
                  className: `lg-panel ${Lr === "assistant" ? "lg-chat-panel" : ""}`,
                  children: [
                    i.jsxs("div", {
                      className: "lg-panel-heading",
                      children: [
                        i.jsxs("div", {
                          children: [
                            Lr === "checks"
                              ? i.jsx(ClipboardCheck, { size: 18 })
                              : Lr === "assistant"
                                ? i.jsx(Sparkles, { size: 18 })
                                : Lr === "history"
                                  ? i.jsx(History, { size: 18 })
                                  : i.jsx(CircleHelp, { size: 18 }),
                            i.jsx("h2", {
                              children:
                                Lr === "checks"
                                  ? "Audit toolkit"
                                  : Lr === "assistant"
                                    ? "Workbook assistant"
                                    : Lr === "history"
                                      ? "Activity"
                                      : "About this workspace",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          "aria-label": T("Close side panel"),
                          onClick: () => Pr(null),
                          lang: N("Close side panel"),
                          "data-ui-label": "Close side panel",
                          children: i.jsx(PanelRightClose, { size: 16 }),
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "lg-panel-body",
                      children: [
                        Lr === "checks" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("span", {
                                className: "lg-eyebrow",
                                children: i.jsx(M, { children: "AUDIT" }),
                              }),
                              i.jsx("h3", {
                                children: i.jsx(M, {
                                  children: "Audit checks",
                                }),
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx(M, { children: "Analysis population" }),
                                  i.jsxs("select", {
                                    "aria-label": T("Analysis population"),
                                    value: ei,
                                    onChange: (fe) => {
                                      (jo(fe.target.value), eo(null));
                                    },
                                    lang: N("Analysis population"),
                                    "data-ui-label": "Analysis population",
                                    children: [
                                      i.jsx("option", {
                                        value: "full",
                                        children:
                                          "Full selected source range (including hidden/filtered)",
                                      }),
                                      i.jsx("option", {
                                        value: "visible",
                                        disabled: !Hr,
                                        children: "Visible filter population",
                                      }),
                                    ],
                                  }),
                                ],
                              }),
                              i.jsxs("p", {
                                className: "lg-muted",
                                children: [
                                  i.jsx(M, {
                                    children:
                                      "Applies to the five checks below. Pattern checks and other tool dialogs use their own complete source ranges and explicit source filters.",
                                  }),
                                  " ",
                                  ei === "visible" && Hr
                                    ? `Using ${Hr.context.range}, ${Hr.context.visibleRows} visible / ${Hr.context.totalRows} total rows. The custom check range is ignored.`
                                    : "Full scope includes hidden and filtered rows.",
                                ],
                              }),
                              i.jsx("p", {
                                className: "lg-muted",
                                children: i.jsx(M, {
                                  children:
                                    "Run a reproducible check, then trace each finding directly to its source.",
                                }),
                              }),
                              i.jsxs("div", {
                                className: "lg-test-choices",
                                children: [
                                  i.jsxs("button", {
                                    onClick: () => Pn("patterns"),
                                    children: [
                                      i.jsx(ListFilter, { size: 17 }),
                                      i.jsxs("span", {
                                        children: [
                                          i.jsx("strong", {
                                            children: i.jsx(M, {
                                              children: "Pattern checks",
                                            }),
                                          }),
                                          i.jsx("small", {
                                            children: i.jsx(M, {
                                              children:
                                                "Duplicates, sequence gaps and unusual postings",
                                            }),
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                                  [
                                    {
                                      kind: "duplicates",
                                      title: "Duplicate entries",
                                      icon: Copy,
                                      detail: "Find repeated transaction IDs",
                                    },
                                    {
                                      kind: "materiality",
                                      title: "Materiality scan",
                                      icon: ListFilter,
                                      detail:
                                        "Surface amounts above a threshold",
                                    },
                                    {
                                      kind: "sample",
                                      title: "Select a sample",
                                      icon: FlaskConical,
                                      detail:
                                        "Repeatable selection with a seed",
                                    },
                                    {
                                      kind: "summary",
                                      title: "Account summary",
                                      icon: ChartColumn,
                                      detail: "Group and total the population",
                                    },
                                    {
                                      kind: "aging",
                                      title: "Aging analysis",
                                      icon: Clock3,
                                      detail: "Bucket transactions by age",
                                    },
                                  ].map((fe) =>
                                    i.jsxs(
                                      "button",
                                      {
                                        onClick: () => {
                                          if (
                                            fe.kind === "materiality" ||
                                            fe.kind === "aging"
                                          ) {
                                            Pn(fe.kind);
                                            return;
                                          }
                                          (lu(fe.kind), eo(null));
                                        },
                                        className:
                                          is === fe.kind ? "active" : "",
                                        children: [
                                          i.jsx(fe.icon, { size: 17 }),
                                          i.jsxs("span", {
                                            children: [
                                              i.jsx("strong", {
                                                children: fe.title,
                                              }),
                                              i.jsx("small", {
                                                children: fe.detail,
                                              }),
                                            ],
                                          }),
                                          i.jsx("span", {
                                            className: "lg-radio",
                                            children:
                                              is === fe.kind &&
                                              i.jsx("span", {}),
                                          }),
                                        ],
                                      },
                                      fe.kind,
                                    ),
                                  ),
                                ],
                              }),
                              i.jsxs("label", {
                                className: "lg-field",
                                children: [
                                  i.jsx(M, { children: "Source range" }),
                                  " ",
                                  i.jsx("input", {
                                    "aria-label": T("Audit source range"),
                                    placeholder: `A1:${ut(he.colCount - 1)}${he.rowCount} (entire sheet)`,
                                    value: Ju,
                                    onChange: (fe) => C0(fe.target.value),
                                    lang: N("Audit source range"),
                                    "data-ui-label": "Audit source range",
                                  }),
                                ],
                              }),
                              i.jsxs("span", {
                                className: "lg-field-hint",
                                children: [
                                  i.jsx(M, {
                                    children: "First row contains field names.",
                                  }),
                                  " ",
                                  ei === "visible"
                                    ? "Uses the visible rows of the applied filter range."
                                    : "Includes all source rows.",
                                ],
                              }),
                              is === "materiality" &&
                                i.jsxs("label", {
                                  className: "lg-field",
                                  children: [
                                    i.jsx(M, {
                                      children: "Absolute amount threshold",
                                    }),
                                    i.jsx(ci, {
                                      "aria-label": T("Materiality threshold"),
                                      min: "0",
                                      value: oa,
                                      onValueChange: (fe) => Dd(Number(fe)),
                                      lang: N("Materiality threshold"),
                                      "data-ui-label": "Materiality threshold",
                                    }),
                                  ],
                                }),
                              is === "sample" &&
                                i.jsxs("div", {
                                  className: "lg-two-fields",
                                  children: [
                                    i.jsxs("label", {
                                      className: "lg-field",
                                      children: [
                                        i.jsx(M, { children: "Sample size" }),
                                        i.jsx(ci, {
                                          "aria-label": T("Sample size"),
                                          min: "1",
                                          value: jl,
                                          onValueChange: (fe) => Uc(Number(fe)),
                                          lang: N("Sample size"),
                                          "data-ui-label": "Sample size",
                                        }),
                                      ],
                                    }),
                                    i.jsxs("label", {
                                      className: "lg-field",
                                      children: [
                                        i.jsx(M, { children: "Seed" }),
                                        i.jsx("input", {
                                          "aria-label": T("Sampling seed"),
                                          value: Fl,
                                          onChange: (fe) => Nf(fe.target.value),
                                          lang: N("Sampling seed"),
                                          "data-ui-label": "Sampling seed",
                                        }),
                                      ],
                                    }),
                                  ],
                                }),
                              is === "aging" &&
                                i.jsxs("label", {
                                  className: "lg-field",
                                  children: [
                                    i.jsx(M, { children: "As of date" }),
                                    i.jsx("input", {
                                      "aria-label": T("Aging as of date"),
                                      type: "date",
                                      value: Od,
                                      onChange: (fe) => um(fe.target.value),
                                      lang: N("Aging as of date"),
                                      "data-ui-label": "Aging as of date",
                                    }),
                                  ],
                                }),
                              i.jsxs("button", {
                                className: "lg-button dark full",
                                onClick: Rs,
                                children: [
                                  i.jsx(ClipboardCheck, { size: 15 }),
                                  i.jsx(M, { children: "Run check" }),
                                  i.jsx(ArrowUpRight, { size: 15 }),
                                ],
                              }),
                              os
                                ? i.jsxs("div", {
                                    className: "lg-results",
                                    children: [
                                      i.jsxs("div", {
                                        className: "lg-results-heading",
                                        children: [
                                          i.jsx("h4", { children: os.title }),
                                          i.jsx("span", {
                                            children: os.findings.length,
                                          }),
                                        ],
                                      }),
                                      i.jsx("p", { children: os.summary }),
                                      i.jsxs("button", {
                                        className: "lg-button full",
                                        onClick: () =>
                                          Jd(
                                            new Blob(
                                              [JSON.stringify($c, null, 2)],
                                              { type: "application/json" },
                                            ),
                                            `Audit-${is}-revision-${H.revision}.json`,
                                          ),
                                        children: [
                                          i.jsx(Download, { size: 13 }),
                                          i.jsx(M, {
                                            children: "Export evidence report",
                                          }),
                                        ],
                                      }),
                                      os.findings.map((fe) =>
                                        i.jsxs(
                                          "article",
                                          {
                                            className: `lg-finding ${fe.severity}`,
                                            children: [
                                              i.jsx("h4", {
                                                children: fe.title,
                                              }),
                                              i.jsx("p", {
                                                children: fe.detail,
                                              }),
                                              i.jsx("div", {
                                                className: "lg-citations",
                                                children: fe.citations.map(
                                                  (Oe, Be) =>
                                                    i.jsxs(
                                                      "button",
                                                      {
                                                        onClick: () => {
                                                          (fn((Ue) => {
                                                            const Xe = {
                                                              ...Ue,
                                                            };
                                                            return (
                                                              delete Xe[
                                                                JSON.stringify([
                                                                  H.id,
                                                                  Oe.sheetId,
                                                                ])
                                                              ],
                                                              Xe
                                                            );
                                                          }),
                                                            wi(Oe.sheetId),
                                                            da(Oe.address),
                                                            eo(os),
                                                            vn(
                                                              `${Oe.sheetName}!${Oe.address}: ${String(Oe.value ?? "")}`,
                                                            ));
                                                        },
                                                        children: [
                                                          Oe.sheetName,
                                                          "!",
                                                          Oe.address,
                                                          i.jsx(ArrowUpRight, {
                                                            size: 10,
                                                          }),
                                                        ],
                                                      },
                                                      `${Oe.address}-${Be}`,
                                                    ),
                                                ),
                                              }),
                                            ],
                                          },
                                          fe.id,
                                        ),
                                      ),
                                    ],
                                  })
                                : i.jsxs("div", {
                                    className: "lg-info-card",
                                    children: [
                                      i.jsx(ShieldCheck, { size: 18 }),
                                      i.jsxs("div", {
                                        children: [
                                          i.jsx("strong", {
                                            children: i.jsx(M, {
                                              children:
                                                "Evidence stays connected",
                                            }),
                                          }),
                                          i.jsx("p", {
                                            children: i.jsx(M, {
                                              children:
                                                "Every finding includes a worksheet and cell reference. Checks leave source data unchanged.",
                                            }),
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                            ],
                          }),
                        Lr === "assistant" &&
                          i.jsx(ne.Suspense, {
                            fallback: i.jsx("p", {
                              role: "status",
                              className: "lg-muted",
                              children: "Loading chat…",
                            }),
                            children: i.jsx(eEe, {
                              session: Md,
                              setSession: A0,
                              adapter: w,
                              local: w === Vj || w.localOnly === true,
                              scope: JSON.stringify([
                                H.id,
                                H.revision,
                                he.id,
                                br,
                                RH(w),
                              ]),
                              selectionLabel: Wa,
                              selectedCount: ie,
                              busy: Qu,
                              proposalId: Xs?.id,
                              onSend: No,
                              onCancel: () => {
                                var fe;
                                (Ys.current++,
                                  (fe = ss.current) == null || fe.abort(),
                                  Vc(false),
                                  Qu &&
                                    vn(
                                      "Assistant request cancelled. No workbook changes were made.",
                                    ));
                              },
                              onReset: () => {
                                (Bl(null), Ll(null), hn(""));
                              },
                              review:
                                Xs && gl
                                  ? i.jsxs(i.Fragment, {
                                      children: [
                                        Xs &&
                                          gl &&
                                          !fg(H, ve, gl).current &&
                                          i.jsx("p", {
                                            role: "status",
                                            children: i.jsx(M, {
                                              children:
                                                "The answer evidence is stale or unavailable. Request a fresh answer before review.",
                                            }),
                                          }),
                                        Xs &&
                                          gl &&
                                          fg(H, ve, gl).current &&
                                          i.jsxs("div", {
                                            className: "lg-proposal",
                                            children: [
                                              i.jsx("span", {
                                                className: "lg-eyebrow",
                                                children: Xs.patches.length
                                                  ? "REVIEW PROPOSED CHANGES"
                                                  : "RESPONSE",
                                              }),
                                              i.jsx(
                                                Q4e,
                                                {
                                                  workbook: H,
                                                  selection: ve,
                                                  evidence: gl,
                                                  renderText: (fe) =>
                                                    i.jsx(tEe, {
                                                      children: fe,
                                                    }),
                                                  onNavigate: (fe, Oe) => {
                                                    const Be = sr.current;
                                                    if (kr.current)
                                                      throw Error(
                                                        "Finish the current cell edit and request fresh evidence before following a citation.",
                                                      );
                                                    if (
                                                      fe !== Be.sheetId ||
                                                      !fg(
                                                        Be.book,
                                                        {
                                                          sheetId: Be.sheetId,
                                                          range: Be.range,
                                                        },
                                                        gl,
                                                      ).current
                                                    )
                                                      throw Error(
                                                        "The current approved source changed. Request a fresh answer.",
                                                      );
                                                    ((zr || kn || Ui) &&
                                                      (fn((Ue) => {
                                                        const Xe = { ...Ue };
                                                        return (
                                                          delete Xe[Xo],
                                                          Xe
                                                        );
                                                      }),
                                                      jr(""),
                                                      po(null),
                                                      vn(
                                                        "The row view was cleared to reveal the cited cell. The approved shared range is unchanged.",
                                                      )),
                                                      or(Oe),
                                                      Go(Oe),
                                                      pi());
                                                  },
                                                },
                                                Xs.id,
                                              ),
                                              Xs.patches.map((fe) => {
                                                var Be;
                                                return i.jsxs(
                                                  "div",
                                                  {
                                                    className: "lg-diff",
                                                    children: [
                                                      i.jsx("b", {
                                                        children: fe.address,
                                                      }),
                                                      i.jsx("del", {
                                                        children: String(
                                                          he.cells[fe.address]
                                                            ?.raw ?? "(blank)",
                                                        ),
                                                      }),
                                                      i.jsxs("span", {
                                                        children: [
                                                          String(
                                                            fe.cell.raw ??
                                                              "(blank)",
                                                          ),
                                                          (Be =
                                                            fe.cell.style) !=
                                                            null && Be.fill
                                                            ? " · highlight"
                                                            : "",
                                                        ],
                                                      }),
                                                      i.jsx("dl", {
                                                        className:
                                                          "lg-proposal-fields",
                                                        "aria-label":
                                                          fe.address +
                                                          " proposed field changes",
                                                        children: [
                                                          [
                                                            "Value type",
                                                            he.cells[fe.address]
                                                              ?.raw == null
                                                              ? "blank"
                                                              : typeof he.cells[
                                                                  fe.address
                                                                ]?.raw,
                                                            fe.cell.raw === null
                                                              ? "blank"
                                                              : typeof fe.cell
                                                                  .raw,
                                                          ],
                                                          ...Array.from(
                                                            new Set([
                                                              ...Object.keys(
                                                                he.cells[
                                                                  fe.address
                                                                ] ?? {},
                                                              ),
                                                              ...Object.keys(
                                                                fe.cell,
                                                              ),
                                                            ]),
                                                          )
                                                            .filter(
                                                              (gt) =>
                                                                gt !== "raw",
                                                            )
                                                            .map((gt) => {
                                                              return [
                                                                gt,
                                                                he.cells[
                                                                  fe.address
                                                                ]?.[gt],
                                                                fe.cell[gt],
                                                              ];
                                                            }),
                                                        ]
                                                          .filter(
                                                            ([, gt, Nt]) =>
                                                              JSON.stringify(
                                                                gt,
                                                              ) !==
                                                              JSON.stringify(
                                                                Nt,
                                                              ),
                                                          )
                                                          .map(([gt, Nt, wt]) =>
                                                            i.jsxs(
                                                              ne.Fragment,
                                                              {
                                                                children: [
                                                                  i.jsx("dt", {
                                                                    children:
                                                                      String(
                                                                        gt,
                                                                      ),
                                                                  }),
                                                                  i.jsxs("dd", {
                                                                    children: [
                                                                      i.jsx(
                                                                        "del",
                                                                        {
                                                                          children:
                                                                            Nt ===
                                                                            undefined
                                                                              ? "(not set)"
                                                                              : JSON.stringify(
                                                                                  Nt,
                                                                                ),
                                                                        },
                                                                      ),
                                                                      i.jsx(
                                                                        "ins",
                                                                        {
                                                                          children:
                                                                            wt ===
                                                                            undefined
                                                                              ? "(not set)"
                                                                              : JSON.stringify(
                                                                                  wt,
                                                                                ),
                                                                        },
                                                                      ),
                                                                    ],
                                                                  }),
                                                                ],
                                                              },
                                                              String(gt),
                                                            ),
                                                          ),
                                                      }),
                                                    ],
                                                  },
                                                  fe.address,
                                                );
                                              }),
                                              Xs.patches.length > 0 &&
                                                i.jsxs(i.Fragment, {
                                                  children: [
                                                    i.jsxs("button", {
                                                      className:
                                                        "lg-button dark full",
                                                      disabled: je,
                                                      onClick: () =>
                                                        On(() => {
                                                          const fe = sr.current,
                                                            Oe = {
                                                              sheetId:
                                                                fe.sheetId,
                                                              range: fe.range,
                                                            };
                                                          if (
                                                            fe.readOnly ||
                                                            kr.current ||
                                                            !fg(fe.book, Oe, gl)
                                                              .current
                                                          )
                                                            throw Error(
                                                              "The current approved source changed or editing is unavailable. Request a fresh answer before applying changes.",
                                                            );
                                                          const Be = X2(
                                                            fe.book,
                                                            Oe,
                                                            Xs,
                                                          );
                                                          (To(
                                                            Be.patches,
                                                            `Accepted assistant proposal ${Be.id}`,
                                                            fe.selectionKey,
                                                          ),
                                                            vn(
                                                              "Reviewed proposal applied and recorded in activity.",
                                                            ));
                                                        }),
                                                      children: [
                                                        i.jsx(Check, {
                                                          size: 15,
                                                        }),
                                                        i.jsx(M, {
                                                          children: "Apply ",
                                                        }),
                                                        Xs.patches.length,
                                                        " ",
                                                        i.jsx(M, {
                                                          children: "changes",
                                                        }),
                                                      ],
                                                    }),
                                                    i.jsx("button", {
                                                      className:
                                                        "lg-button full",
                                                      onClick: () => {
                                                        (Bl(null), Ll(null));
                                                      },
                                                      children: i.jsx(M, {
                                                        children:
                                                          "Discard proposal",
                                                      }),
                                                    }),
                                                  ],
                                                }),
                                            ],
                                          }),
                                      ],
                                    })
                                  : null,
                            }),
                          }),
                        Lr === "history" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("span", {
                                className: "lg-eyebrow",
                                children: i.jsx(M, {
                                  children: "WORKBOOK ACTIVITY",
                                }),
                              }),
                              i.jsx("h3", {
                                children: i.jsx(M, {
                                  children: "Change history",
                                }),
                              }),
                              i.jsx("p", {
                                className: "lg-muted",
                                children: i.jsx(M, {
                                  children:
                                    "Local activity is included in JSON exports. This log is editable with the file and is not an immutable audit record.",
                                }),
                              }),
                              H.events.length === 0
                                ? i.jsxs("div", {
                                    className: "lg-empty",
                                    children: [
                                      i.jsx(History, { size: 26 }),
                                      i.jsx("p", {
                                        children: i.jsx(M, {
                                          children:
                                            "Your first edit starts the activity trail.",
                                        }),
                                      }),
                                    ],
                                  })
                                : i.jsx("div", {
                                    className: "lg-timeline",
                                    children: [...H.events]
                                      .reverse()
                                      .map((fe) =>
                                        i.jsxs(
                                          "article",
                                          {
                                            children: [
                                              i.jsx("span", {
                                                className: "lg-timeline-dot",
                                              }),
                                              i.jsx("h4", {
                                                children: fe.action,
                                              }),
                                              i.jsxs("p", {
                                                children: [
                                                  fe.actor,
                                                  " ·",
                                                  " ",
                                                  new Date(
                                                    fe.timestamp,
                                                  ).toLocaleString(),
                                                ],
                                              }),
                                              fe.changes
                                                .slice(0, 5)
                                                .map((Oe, Be) => {
                                                  return i.jsxs(
                                                    "small",
                                                    {
                                                      children: [
                                                        H.sheets.find(
                                                          (Nt) =>
                                                            Nt.id ===
                                                            Oe.sheetId,
                                                        )?.name ?? "Sheet",
                                                        "!",
                                                        Oe.address,
                                                        ":",
                                                        " ",
                                                        String(
                                                          Oe.before?.raw ?? "∅",
                                                        ),
                                                        " →",
                                                        " ",
                                                        String(
                                                          Oe.after?.raw ?? "∅",
                                                        ),
                                                      ],
                                                    },
                                                    Be,
                                                  );
                                                }),
                                              fe.changes.length > 5 &&
                                                i.jsxs("small", {
                                                  children: [
                                                    "+ ",
                                                    fe.changes.length - 5,
                                                    " ",
                                                    i.jsx(M, {
                                                      children:
                                                        "more cell changes",
                                                    }),
                                                  ],
                                                }),
                                            ],
                                          },
                                          fe.id,
                                        ),
                                      ),
                                  }),
                            ],
                          }),
                        Lr === "help" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("span", {
                                className: "lg-eyebrow",
                                children: i.jsx(M, {
                                  children: "SPREADSHEET HELP",
                                }),
                              }),
                              i.jsx("h3", {
                                children: i.jsx(M, {
                                  children: "Workbook help",
                                }),
                              }),
                              i.jsx("p", {
                                className: "lg-muted",
                                children: i.jsx(M, {
                                  children:
                                    "An embeddable React workbook for audit preparation and review.",
                                }),
                              }),
                              i.jsx("h4", {
                                children: i.jsx(M, {
                                  children: "Working with cells",
                                }),
                              }),
                              i.jsxs("ul", {
                                className: "lg-help-list",
                                children: [
                                  i.jsx("li", {
                                    children: i.jsx(M, {
                                      children:
                                        "Click a cell, then type or use the formula bar. Enter commits the value.",
                                    }),
                                  }),
                                  i.jsx("li", {
                                    children: i.jsx(M, {
                                      children:
                                        "Enter dates as YYYY-MM-DD using the workbook's selected 1900 or 1904 date system. Other date formats stay text; prefix an ISO date with an apostrophe to keep it as text. The formula bar shows the stored serial and the grid shows the date.",
                                    }),
                                  }),
                                  i.jsx("li", {
                                    children: i.jsx(M, {
                                      children:
                                        "Ctrl/Command-click to add areas; Shift-click extends a range. Enter A1:D10, F1:F10 in the name box.",
                                    }),
                                  }),
                                  i.jsx("li", {
                                    children: i.jsx(M, {
                                      children:
                                        "Copy and paste rectangular tab-separated values with Ctrl/⌘ C and V.",
                                    }),
                                  }),
                                  i.jsx("li", {
                                    children: i.jsx(M, {
                                      children:
                                        "Prefix text with an apostrophe to prevent formula evaluation.",
                                    }),
                                  }),
                                  i.jsx("li", {
                                    children: i.jsx(M, {
                                      children:
                                        "Sort and search change the view; source row numbers remain fixed.",
                                    }),
                                  }),
                                ],
                              }),
                              i.jsx("h4", {
                                children: i.jsx(M, {
                                  children: "Compatibility",
                                }),
                              }),
                              i.jsx("p", {
                                className: "lg-muted",
                                children: i.jsx(M, {
                                  children:
                                    "XLSX retains a documented subset of cells, formulas and formatting. JSON retains the complete native workpaper. Macros, Power Query, DAX, advanced PivotTables, and full Excel layout fidelity are target requirements, not features of this release.",
                                }),
                              }),
                              i.jsx(Yce, {}),
                              i.jsxs("button", {
                                className: "lg-button full",
                                onClick: () => window.print(),
                                children: [
                                  i.jsx(FileSpreadsheet, { size: 15 }),
                                  i.jsx(M, { children: "Print current view" }),
                                ],
                              }),
                            ],
                          }),
                      ],
                    }),
                  ],
                }),
            ],
          }),
        ],
      }),
      B && i.jsx(V2e, { monitor: D, onClose: () => _(false) }),
      ml &&
        i.jsx(Gce, {
          error: ml.error,
          readOnly: je,
          onRetry: () => {
            (Rf(null),
              requestAnimationFrame(() => {
                return pt.current?.focus();
              }));
          },
          onCancel: () => {
            (Rf(null), hn(""), Da());
          },
          onAccept: () =>
            On(() => {
              if (sr.current.readOnly)
                throw Error("This workbook is read-only.");
              if (
                ml.selectionGuard &&
                ml.selectionGuard !== sr.current.selectionKey
              )
                throw Error(
                  "Approved selection changed. Review the proposal again.",
                );
              if (ml.snapshot !== JSON.stringify(sr.current.book))
                throw Error("Workbook changed. Review the batch again.");
              (di(
                p0(sr.current.book, ml.patches, E, ml.action, {
                  validationApproval: jB(ml.error),
                }),
              ),
                Rf(null),
                (kr.current = null),
                Ir(null),
                vn("Validation warnings accepted for the complete batch."),
                pi());
            }),
        }),
      aa &&
        i.jsx("div", {
          className: "lg-modal-backdrop",
          onKeyDown: (fe) => {
            fe.key === "Escape" && Ia(null);
          },
          onMouseDown: (fe) => {
            fe.target === fe.currentTarget && Ia(null);
          },
          children: i.jsxs("section", {
            className: "lg-modal",
            role: "dialog",
            "aria-modal": "true",
            "aria-labelledby": `${$}-dialog-title`,
            children: [
              i.jsx("button", {
                className: "lg-modal-close",
                "aria-label": T("Close dialog"),
                onClick: () => Ia(null),
                lang: N("Close dialog"),
                "data-ui-label": "Close dialog",
                children: i.jsx(X, { size: 18 }),
              }),
              i.jsx("h2", {
                id: `${$}-dialog-title`,
                children:
                  aa.kind === "note"
                    ? "Add a cell note"
                    : aa.kind === "rename"
                      ? "Rename workbook"
                      : "Add worksheet",
              }),
              aa.kind === "note"
                ? i.jsx("textarea", {
                    autoFocus: true,
                    "aria-label": T("Cell note"),
                    rows: 5,
                    value: aa.value,
                    onChange: (fe) => Ia({ ...aa, value: fe.target.value }),
                    lang: N("Cell note"),
                    "data-ui-label": "Cell note",
                  })
                : i.jsx("input", {
                    autoFocus: true,
                    "aria-label": T("Name"),
                    value: aa.value,
                    onChange: (fe) => Ia({ ...aa, value: fe.target.value }),
                    onKeyDown: (fe) => {
                      (fe.key === "Enter" && el(),
                        fe.key === "Escape" && Ia(null));
                    },
                    lang: N("Name"),
                    "data-ui-label": "Name",
                  }),
              i.jsxs("div", {
                className: "lg-modal-actions",
                children: [
                  i.jsx("button", {
                    className: "lg-button",
                    onClick: () => Ia(null),
                    children: i.jsx(M, { children: "Cancel" }),
                  }),
                  i.jsx("button", {
                    className: "lg-button dark",
                    onClick: el,
                    children: i.jsx(M, { children: "Save" }),
                  }),
                ],
              }),
            ],
          }),
        }),
    ],
  });
}
function nEe({
  idPrefix: e,
  pinned: r,
  pinnedTop: t,
  frozenColumns: n,
  frozenColumnOffsets: o,
  headings: a,
  onRowSelect: s,
  onDragStart: d,
  onDragEnter: u,
  onNavigate: l,
  onPivotDrill: c,
  onFillStart: h,
  onFillOptions: f,
  fillDragging: p,
  commentAddresses: m,
  onReview: v,
  onLink: g,
  row: b,
  columns: y,
  leftGap: w,
  rightGap: E,
  minimumHeight: A,
  fixedHeight: k,
  onResize: F,
  onAutoFit: I,
  onCancelEdit: O,
  onFormulaInput: j,
  onFormulaKey: z,
  formulaHelpId: V,
  onFormulaPick: D,
  formulaReference: B,
  formulaDraftActive: _,
  workbook: C,
  mergeBounds: T,
  conditionalFormatter: N,
  displayedRows: R,
  sortedView: L,
  filterHeader: $,
  filteredColumns: U,
  onColumnFilter: G,
  ruleStyles: H,
  sheetId: Q,
  cells: X,
  active: J,
  selectedBounds: q,
  selectionBounds: W,
  calc: K,
  showFormulas: te,
  editing: ae,
  setEditing: oe,
  select: de,
  saveCell: he,
  readOnly: Se,
}) {
  const { text: le, language: ce } = Dr(),
    ge = Dr(),
    se = (Te) => aI(Te, ge.entryLocale),
    we = ne.useRef(null),
    Ee = ne.useRef(null),
    [Ae, ye] = ne.useState(0);
  return i.jsx(i.Fragment, {
    children: i.jsxs("tr", {
      role: "row",
      "data-row": b,
      style: { height: A, ...(r ? { top: t } : {}) },
      className: r ? "lg-pinned-row" : "",
      "aria-rowindex": b + 2,
      children: [
        a &&
          i.jsxs("th", {
            role: "rowheader",
            scope: "row",
            "data-row-header": b,
            onClick: (Te) => s(Te.shiftKey),
            "aria-label": `Select row ${b + 1}`,
            className: _t(J).row === b ? "lg-selected-header" : "",
            children: [
              b + 1,
              i.jsx("span", {
                className: "lg-row-resize",
                role: "separator",
                "aria-label": `Resize row ${b + 1}`,
                "aria-orientation": "horizontal",
                "aria-disabled": Se,
                title: le("Drag to resize; double-click to AutoFit row height"),
                onPointerDown: F,
                onClick: (Te) => Te.stopPropagation(),
                onDoubleClick: (Te) => {
                  (Te.stopPropagation(), I());
                },
                lang: ce("Drag to resize; double-click to AutoFit row height"),
              }),
            ],
          }),
        [...n, ...y].map((Te, Pe) => {
          const qe = T.find((lt) => Tp(lt, b, Te));
          let ht = 1,
            Re = 1,
            Me = `${ut(Te)}${b + 1}`;
          if (qe && !L) {
            const lt = (n.includes(Te) ? n : y).filter(
                (ft) => ft >= qe.firstCol && ft <= qe.lastCol,
              ),
              be = R.filter((ft) => ft >= qe.firstRow && ft <= qe.lastRow);
            if (Te !== lt[0] || b !== be[0]) return null;
            ((ht = lt.length),
              (Re = be.length),
              (Me = `${ut(qe.firstCol)}${qe.firstRow + 1}`));
          }
          const ke = X[Me],
            je = {
              ...ke,
              raw: ke?.raw ?? null,
              style: { ...ke?.style, ...H(Me) },
            },
            $e = K(Q, Me),
            Je =
              typeof je?.raw == "string" &&
              je.raw.startsWith("=") &&
              !je.literal,
            ze = J === Me,
            at = {
              ...Dw(je.style, $e),
              ...(n.includes(Te)
                ? { position: "sticky", left: o[Te], zIndex: 3 }
                : {}),
            };
          return i.jsxs(
            ne.Fragment,
            {
              children: [
                Pe === n.length &&
                  w > 0 &&
                  i.jsx("td", {
                    className: "lg-column-spacer",
                    "aria-hidden": "true",
                  }),
                i.jsxs(
                  "td",
                  {
                    id: `${e}-${Me}`,
                    colSpan: ht,
                    rowSpan: Re,
                    "aria-colspan": ht > 1 ? ht : undefined,
                    "aria-rowspan": Re > 1 ? Re : undefined,
                    role: "gridcell",
                    "aria-label": `${Me} ${ge.entryLocale ? lI($e, je, C.datePolicy?.dateSystem, ge.entryLocale) : String($e ?? "")}`,
                    "aria-colindex": Te + 2,
                    "aria-selected": W.some((lt) => Tp(lt, b, Te)),
                    "data-address": Me,
                    "data-formula-reference":
                      B && Tp(B, b, Te) ? "true" : undefined,
                    className: `${ze ? "lg-cell-active" : ""} ${W.some((lt) => Tp(lt, b, Te)) ? "lg-cell-selected" : ""} ${Je ? "lg-formula-cell" : ""} ${typeof $e == "string" && $e.startsWith("#") ? "lg-cell-error" : ""} ${n.includes(Te) ? "lg-frozen-column" : ""}`,
                    style: at,
                    title: `${Me}${je != null && je.note ? ` · ${je.note}` : ""}${Je ? ` · ${je.raw}` : ""}`,
                    onPointerDown: (lt) => {
                      lt.button !== 0 ||
                        ae !== null ||
                        (lt.preventDefault(),
                        d(Me, lt.ctrlKey || lt.metaKey),
                        de(Me, lt.shiftKey, lt.ctrlKey || lt.metaKey));
                    },
                    onPointerEnter: (lt) => u(Me, lt.buttons),
                    onClick: (lt) => {
                      lt.detail === 0 &&
                        ((_ && D(Me, lt.shiftKey)) ||
                          de(Me, lt.shiftKey, lt.ctrlKey || lt.metaKey));
                    },
                    onDoubleClick: () => {
                      _ || c(Me) || Se || oe(se(je));
                    },
                    children: [
                      ze && ae !== null
                        ? i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("input", {
                                ref: we,
                                className: "lg-cell-editor",
                                onFocus: (lt) => j(lt.currentTarget),
                                autoFocus: true,
                                readOnly: Se,
                                "aria-label": `Edit ${Me}`,
                                "aria-describedby": _ ? V : undefined,
                                value: ae,
                                onSelect: (lt) => {
                                  (j(lt.currentTarget),
                                    ye(
                                      lt.currentTarget.selectionStart ??
                                        ae.length,
                                    ));
                                },
                                onChange: (lt) => {
                                  (oe(lt.target.value),
                                    j(lt.currentTarget, true),
                                    ye(
                                      lt.target.selectionStart ??
                                        lt.target.value.length,
                                    ));
                                },
                                onClick: (lt) => lt.stopPropagation(),
                                onBlur: (lt) => {
                                  dH(lt.currentTarget, lt.relatedTarget) ||
                                    he(ae);
                                },
                                onKeyDown: (lt) => {
                                  var be, ft;
                                  if (
                                    (lt.stopPropagation(),
                                    !lt.nativeEvent.isComposing)
                                  ) {
                                    if (lt.key === "Escape") {
                                      if (
                                        !B &&
                                        ae.length > 1 &&
                                        (be = Ee.current) != null &&
                                        be.handleKeyDown(lt)
                                      )
                                        return;
                                      (lt.preventDefault(), O());
                                      return;
                                    }
                                    z(lt) ||
                                      ((ft = Ee.current) != null &&
                                        ft.handleKeyDown(lt)) ||
                                      ((lt.key === "Enter" ||
                                        lt.key === "Tab") &&
                                        (lt.preventDefault(),
                                        he(ae) && l(lt.key, lt.shiftKey)),
                                      lt.key === "Escape" &&
                                        (lt.preventDefault(), O()));
                                  }
                                },
                              }),
                              i.jsx(fH, {
                                ref: Ee,
                                onExit: () => {
                                  he(ae);
                                },
                                value: ae,
                                caret: Ae,
                                workbook: C,
                                sheetId: Q,
                                inputRef: we,
                                disabled: Se,
                                onAccept: (lt, be) => {
                                  (oe(lt),
                                    ye(be),
                                    requestAnimationFrame(() => {
                                      var ft, pt;
                                      ((ft = we.current) == null || ft.focus(),
                                        (pt = we.current) == null ||
                                          pt.setSelectionRange(be, be));
                                    }));
                                },
                              }),
                            ],
                          })
                        : i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("span", {
                                style: {
                                  ...RN(je.style),
                                  ...(k && Re === 1
                                    ? { maxHeight: A - 1, minHeight: 0 }
                                    : {}),
                                },
                                children:
                                  te && Je
                                    ? String(je.raw)
                                    : i.jsx(Fz, {
                                        decoration: N(Me),
                                        children: i.jsx(jI, {
                                          value: $e,
                                          cell: je,
                                          dateSystem: C.datePolicy?.dateSystem,
                                        }),
                                      }),
                              }),
                              $?.some((lt) => lt.column === Te) &&
                                i.jsx("button", {
                                  className: "lg-header-filter",
                                  "aria-label": `Filter ${$.find((lt) => lt.column === Te).label} column`,
                                  "aria-pressed": U.includes(Te),
                                  onPointerDown: (lt) => lt.stopPropagation(),
                                  onClick: (lt) => {
                                    (lt.stopPropagation(), G(Te));
                                  },
                                  title: le("Filter this column"),
                                  lang: ce("Filter this column"),
                                  children: "▾",
                                }),
                              je?.hyperlink &&
                                (je.hyperlink.kind === "external"
                                  ? i.jsx("a", {
                                      className: "lg-cell-link",
                                      href: je.hyperlink.url,
                                      target: "_blank",
                                      rel: "noopener noreferrer",
                                      "aria-label": `Open link at ${Me}`,
                                      title:
                                        je.hyperlink.label || je.hyperlink.url,
                                      onPointerDown: (lt) =>
                                        lt.stopPropagation(),
                                      onClick: (lt) => lt.stopPropagation(),
                                      children: "↗",
                                    })
                                  : i.jsx("button", {
                                      className: "lg-cell-link",
                                      "aria-label": `Open link at ${Me}`,
                                      title:
                                        je.hyperlink.label ||
                                        `Go to ${je.hyperlink.address}`,
                                      onPointerDown: (lt) =>
                                        lt.stopPropagation(),
                                      onClick: (lt) => {
                                        lt.stopPropagation();
                                        const be = je.hyperlink;
                                        be?.kind === "internal" &&
                                          g(be.sheetId, be.address);
                                      },
                                      children: "↗",
                                    })),
                              m.has(Me) &&
                                i.jsx("button", {
                                  className: "lg-comment-marker",
                                  style: {
                                    right: je != null && je.note ? 8 : 0,
                                  },
                                  "aria-label": `Review comments at ${Me}`,
                                  title: le("Open review discussion"),
                                  onPointerDown: (lt) => lt.stopPropagation(),
                                  onClick: (lt) => {
                                    (lt.stopPropagation(), v(Me));
                                  },
                                  lang: ce("Open review discussion"),
                                  children: "●",
                                }),
                              je?.note &&
                                i.jsx("i", { className: "lg-note-marker" }),
                              je?.tickmark &&
                                i.jsx("span", {
                                  className: "lg-tick",
                                  children: je.tickmark,
                                }),
                            ],
                          }),
                      !Se &&
                        !p &&
                        ae === null &&
                        b === q.lastRow &&
                        Te === q.lastCol &&
                        i.jsx("button", {
                          className: "lg-fill-handle",
                          "aria-label": le("Fill handle"),
                          title: le(
                            "Drag to extend the selection; use Fill Series for keyboard options",
                          ),
                          onPointerDown: h,
                          onClick: (lt) => {
                            (lt.stopPropagation(), lt.detail === 0 && f());
                          },
                          onDoubleClick: (lt) => lt.stopPropagation(),
                          lang: ce("Fill handle"),
                          "data-ui-label": "Fill handle",
                        }),
                    ],
                  },
                  `${Me}:${b}:${Te}`,
                ),
              ],
            },
            `${Me}:${b}:${Te}`,
          );
        }),
        E > 0 &&
          i.jsx("td", { className: "lg-column-spacer", "aria-hidden": "true" }),
      ],
    }),
  });
}

export { IH };
