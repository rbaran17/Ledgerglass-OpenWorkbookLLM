import i from "react/jsx-runtime";
import ne from "react";
import { iu } from "./part-01.js";
import { Vn, X5, ct, ol, p0, rn, ut } from "./part-03.js";
import { Dr, M } from "./part-04.js";
import { ow } from "./part-07.js";
import { L$, jC } from "./part-08.js";
import { FC } from "./part-09.js";
import { NI } from "./part-20.js";
import { sw } from "./part-24.js";
import { Hxe, KT, Vxe, Wxe, ZT, pS } from "./part-40.js";

function qxe(e) {
  const r = [],
    t = [],
    n = new Map(),
    o = new Map();
  let a = false,
    s = false,
    d = false,
    u = "";
  const l = (c) => {
    let h = new RegExp(
      `^(?:ThisWorkbook\\.)?(?:Worksheets|Sheets)\\((${KT})\\)$`,
      "i",
    ).exec(c);
    if (h) return { kind: "sheet", sheet: String(pS(h[1])) };
    if (/^ActiveSheet$/i.test(c)) return { kind: "sheet" };
    if (
      ((h = new RegExp(`^(?:(.+)\\.)?Range\\((${KT})\\)$`, "i").exec(c)), h)
    ) {
      const p = h[1] ? l(h[1]) : { kind: "sheet" };
      if (p.kind !== "sheet")
        throw Error("Range requires a worksheet binding.");
      return { kind: "range", sheet: p.sheet, range: String(pS(h[2])) };
    }
    const f = n.get(c.toLowerCase());
    if (f) return f;
    throw Error(
      "Unsupported VBA object or dynamic reference. Use a literal worksheet/range in the current workbook.",
    );
  };
  for (const [c, h] of e.split(/\r\n|\n|\r/).entries()) {
    const f = c + 1;
    try {
      if (/^\s*Rem(?:\s|$)/i.test(h)) continue;
      const p = Hxe(h);
      if (!p || /^Rem(?:\s|$)/i.test(p)) continue;
      if (!s && !d && /^Option Explicit$/i.test(p)) {
        a = !0;
        continue;
      }
      if (
        !s &&
        !d &&
        (/^Option Explicit$/i.test(p) ||
          /^Attribute VB_Name = "(?:[^"]|"")*"$/i.test(p))
      )
        continue;
      let m =
        /^(?:(?:Public|Private)\s+)?Sub\s+([A-Za-z_][\w]*)\s*\(\s*\)$/i.exec(p);
      if (m) {
        if (s || d)
          throw Error("Only one parameterless VBA Sub may be migrated.");
        ((s = !0), (u = m[1]));
        continue;
      }
      if (/^End Sub$/i.test(p)) {
        if (!s) throw Error("Unexpected End Sub.");
        ((s = !1), (d = !0));
        continue;
      }
      if (!s) throw Error("Code outside the single Sub is unsupported.");
      if (
        ((m = /^Dim\s+([A-Za-z_]\w*)\s+As\s+(Worksheet|Range)$/i.exec(p)), m)
      ) {
        const v = m[1].toLowerCase();
        if (o.has(v) || n.has(v))
          throw Error("Duplicate or late VBA declaration.");
        o.set(v, m[2].toLowerCase() === "worksheet" ? "sheet" : "range");
        continue;
      }
      if (((m = /^Set\s+([A-Za-z_]\w*)\s*=\s*(.+)$/i.exec(p)), m)) {
        const v = m[1].toLowerCase();
        if (n.has(v)) throw Error("Binding reassignment is unsupported.");
        const g = l(m[2]);
        if (a && !o.has(v))
          throw Error(
            "Option Explicit requires a declared worksheet/range binding.",
          );
        if (o.has(v) && o.get(v) !== g.kind)
          throw Error("VBA binding type does not match its declaration.");
        n.set(v, g);
        continue;
      }
      if (((m = /^(.+)\.(Value2?|Formula)\s*=\s*(.+)$/i.exec(p)), m)) {
        const v = l(m[1]);
        if (v.kind !== "range")
          throw Error("Assignment requires a literal range.");
        r.push({
          line: f,
          kind:
            m[2].toLowerCase() === "formula" ? "set-formulas" : "set-values",
          sheet: v.sheet,
          range: v.range,
          values: pS(m[3]),
        });
        continue;
      }
      if (((m = /^(.+)\.ClearContents$/i.exec(p)), m)) {
        const v = l(m[1]);
        if (v.kind !== "range") throw Error("ClearContents requires a range.");
        r.push({
          line: f,
          kind: "clear-contents",
          sheet: v.sheet,
          range: v.range,
        });
        continue;
      }
      throw Error(
        "Unsupported VBA statement/API. Loops, expressions, calls, events, files, COM and external operations are not migrated.",
      );
    } catch (p) {
      t.push({
        line: f,
        message: p instanceof Error ? p.message : "Unsupported VBA source.",
      });
    }
  }
  return (
    (s || !d) &&
      t.push({
        line: 1,
        message: "A complete single parameterless Sub / End Sub is required.",
      }),
    { commands: r, issues: t, entryPoint: u }
  );
}
function SW(e, r) {
  if (
    typeof e != "string" ||
    new TextEncoder().encode(e).length > ol.maxSourceBytes ||
    e.includes("\0")
  )
    throw Error(
      "Migration source must be UTF-8 text of at most 256 KiB without nulls.",
    );
  if (!["vba", "office-script"].includes(r))
    throw Error("Unsupported automation language.");
  let t = [],
    n = [],
    o = r === "office-script" ? "main" : "";
  try {
    if (r === "vba") ({ commands: t, issues: n, entryPoint: o } = qxe(e));
    else {
      const a = new Wxe(Vxe(e));
      (a.parse(), (t = a.commands), (n = a.issues));
    }
  } catch (a) {
    n.push(ZT(a));
  }
  return (
    t.length > ol.maxCommands &&
      n.push({ line: 1, message: "Migration is limited to 100 commands." }),
    t.length ||
      n.push({ line: 1, message: "No supported range commands were found." }),
    {
      adapter: "literal-ranges-v1",
      language: r,
      codeSha256: iu(new TextEncoder().encode(e)),
      supported: !n.length,
      entryPoint: o,
      commands: t,
      issues: n,
      warnings: [
        "This is a strict migration adapter, not a VBA/JavaScript runtime. Unsupported statements block the entire script; no partial execution occurs.",
        "Only literal range values, explicit formulas and ClearContents are supported. One approved worksheet/range is the read/write scope, up to 100 commands and 500 touched cells. Source events never become automatic triggers.",
        "Value writes that would be interpreted as formulas in Excel are refused. Use an explicit formula command and review its selected-range dependencies.",
      ],
    }
  );
}
const YT = new WeakMap(),
  TW = (e) => {
    const r = ct(e.range);
    if (
      (r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1) >
      ol.maxCells
    )
      throw Error("Approve one rectangular scope of at most 500 cells.");
    return {
      sheetId: e.sheetId,
      range: `${ut(r.firstCol)}${r.firstRow + 1}:${ut(r.lastCol)}${r.lastRow + 1}`,
    };
  };
function NW(e, r, t = false) {
  var o;
  if (t) throw Error("This workbook is read-only.");
  const n = e.sheets.find((a) => a.id === r.sheetId);
  if (!n) throw Error("The approved worksheet does not exist.");
  if (n.hidden || n.veryHidden)
    throw Error("Hidden worksheets are outside the automation scope.");
  if ((o = n.protection) != null && o.enabled)
    throw Error("The approved worksheet is protected.");
  return n;
}
const Xxe = (e) => {
  const r = ct(e.range),
    t = new Set();
  for (let n = r.firstRow; n <= r.lastRow; n++)
    for (let o = r.firstCol; o <= r.lastCol; o++) t.add(`${ut(o)}${n + 1}`);
  return t;
};
function Gxe(e, r, t, n = {}) {
  var f, p;
  if ((jC(r), !r.script))
    throw Error(
      "This artifact supports inventory only. Export/import UTF-8 VBA or Office Scripts source for migration.",
    );
  const o = rn(e),
    a = TW(t),
    s = NW(o, a, n.readOnly),
    d = Xxe(a),
    u = SW(r.script.text, r.script.language),
    l = {
      schemaVersion: 1,
      workbookId: o.id,
      revision: o.revision,
      sourceName: r.fileName,
      sourceSha256: r.sha256,
      compatibility: u,
      scope: a,
      valid: false,
      issues: u.issues.map((m) => `Line ${m.line}: ${m.message}`),
      touchedCells: 0,
      changes: [],
      capabilities: [
        "Write only the approved worksheet rectangle; existing dependent formulas recalculate normally",
        "Set literal values",
        "Set explicit scalar formulas with direct references inside the approved rectangle",
        "Clear cell contents; retain formatting",
        "No network, files, external workbook, COM, events or arbitrary code",
      ],
    },
    c = new Map();
  if (u.supported)
    for (const m of u.commands)
      try {
        if (
          m.sheet !== void 0 &&
          m.sheet.toLowerCase() !== s.name.toLowerCase()
        )
          throw Error(
            `Worksheet ${m.sheet} is outside the approved sheet ${s.name}.`,
          );
        const v = ct(m.range),
          g = v.lastRow - v.firstRow + 1,
          b = v.lastCol - v.firstCol + 1,
          y = g * b;
        if (((l.touchedCells += y), l.touchedCells > ol.maxCells))
          throw Error(
            "Migration exceeds 500 touched cells across all commands.",
          );
        const w = Array.isArray(m.values) ? m.values : void 0;
        if (
          w &&
          (w.length !== g || w.some((E) => !Array.isArray(E) || E.length !== b))
        )
          throw Error(
            "The literal matrix dimensions do not match the target range.",
          );
        if (m.kind === "set-formulas" && !w && y !== 1)
          throw Error(
            "Scalar formula migration requires one cell; use explicit setFormulas matrices for multiple cells.",
          );
        for (let E = v.firstRow; E <= v.lastRow; E++)
          for (let A = v.firstCol; A <= v.lastCol; A++) {
            const k = `${ut(A)}${E + 1}`;
            if (!d.has(k))
              throw Error(
                `Target ${k} is outside the approved range ${a.range}.`,
              );
            const F =
              m.kind === "clear-contents"
                ? null
                : w
                  ? w[E - v.firstRow][A - v.firstCol]
                  : m.values;
            if (m.kind === "set-values" && F === null)
              throw Error(
                "Null setValue semantics are not migrated. Use an explicit ClearContents command.",
              );
            if (
              m.kind === "set-values" &&
              typeof F == "string" &&
              /^[\s\u0000-\u001f]*[=+\-@＝＋－＠]/u.test(F)
            )
              throw Error(
                "Formula-like value assignment has different Excel semantics. Use an explicit setFormula / .Formula statement or change the reviewed source.",
              );
            const I = ((f = c.get(k)) == null ? void 0 : f.cell) ?? s.cells[k],
              O = { ...I, raw: F };
            if ((delete O.literal, m.kind === "set-formulas")) {
              if (typeof F != "string" || !F.startsWith("="))
                throw Error(
                  "Formula commands require explicit A1 formulas beginning with =.",
                );
              NI(F, d);
            }
            c.set(k, { sheetId: s.id, address: k, cell: O });
          }
      } catch (v) {
        l.issues.push(
          `Line ${m.line}: ${v instanceof Error ? v.message : "Migration validation failed."}`,
        );
      }
  const h = [...c.values()];
  if (!l.issues.length)
    try {
      const m = p0(o, h, "Migration preview", "Preview migrated automation"),
        v = ((p = m.events.at(-1)) == null ? void 0 : p.changes) ?? [],
        g = Vn(m);
      for (const b of h) {
        if (typeof b.cell.raw != "string" || !b.cell.raw.startsWith("="))
          continue;
        if (
          (g(s.id, b.address),
          g.getOperationalErrors().length || g.getReferenceDiagnostics().length)
        )
          throw Error(
            "A migrated formula has invalid syntax, a circular reference, an unsupported reference or a calculation-limit failure.",
          );
        const y = g.getSpillRange(s.id, b.address);
        if (
          g.getSpillDiagnostic(s.id, b.address) ||
          (y && y !== `${b.address}:${b.address}`)
        )
          throw Error(
            "Array-spilling formula migration is unsupported. Use explicit scalar formulas so every write is reviewable.",
          );
      }
      if (v.some((b) => b.sheetId !== a.sheetId || !d.has(b.address)))
        throw Error(
          "A dependent table/header rewrite would leave the approved scope. Migrate a separate range instead.",
        );
      ((l.changes = v.map((b) => ({
        sheetId: b.sheetId,
        address: b.address,
        before: b.before ?? null,
        after: b.after,
      }))),
        (l.valid = !0));
    } catch (m) {
      l.issues.push(
        m instanceof Error ? m.message : "Workbook validation failed.",
      );
    }
  return (
    YT.set(l, {
      signature: JSON.stringify(l),
      snapshot: JSON.stringify(e),
      inspection: r,
      patches: h,
    }),
    l
  );
}
function RW(e, r) {
  var t;
  return (t = e.automationArtifacts) != null &&
    t.some((n) => n.sha256 === r.sha256 && n.fileName === r.fileName)
    ? e
    : rn({ ...e, automationArtifacts: [...(e.automationArtifacts ?? []), r] });
}
function Zxe(e, r, t) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  const n = RW(e, L$(r, t.actor));
  return n === e
    ? e
    : sw(
        e,
        n,
        t.actor,
        `Preserve automation source: ${r.fileName.slice(0, 100)}`,
      );
}
function Kxe(e, r, t) {
  var n;
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (!((n = e.automationArtifacts) != null && n.some((o) => o.id === r)))
    throw Error("Automation artifact was not found.");
  return sw(
    e,
    {
      ...e,
      automationArtifacts: e.automationArtifacts.filter((o) => o.id !== r),
    },
    t.actor,
    "Remove preserved automation artifact",
  );
}
function Yxe(e, r, t) {
  var d;
  const n = YT.get(r);
  if (!n || n.signature !== JSON.stringify(r) || !r.valid)
    throw Error(
      "The migration plan is invalid, modified or untrusted. Create and review a fresh plan.",
    );
  if (JSON.stringify(e) !== n.snapshot)
    throw Error("The workbook changed. Review a fresh migration plan.");
  if (JSON.stringify(TW(t.scope)) !== JSON.stringify(r.scope))
    throw Error("The approved automation scope changed. Review again.");
  if (
    (NW(e, r.scope, t.readOnly),
    jC(n.inspection),
    (((d = e.automationRuns) == null ? void 0 : d.length) ?? 0) >= ol.maxRuns)
  )
    throw Error(
      "Automation run evidence is full (100 runs). Export native JSON before continuing.",
    );
  let o = p0(
    e,
    n.patches,
    t.actor,
    `Run migrated ${r.compatibility.language}: ${r.sourceName.slice(0, 90)} · ${r.sourceSha256.slice(0, 12)}`,
  );
  t.retainSource !== false && (o = RW(o, L$(n.inspection, t.actor)));
  const a = o.events.at(-1),
    s = {
      id: crypto.randomUUID(),
      eventId: a.id,
      sourceName: r.sourceName,
      sourceSha256: r.sourceSha256,
      codeSha256: r.compatibility.codeSha256,
      language: r.compatibility.language,
      adapter: "literal-ranges-v1",
      actor: a.actor,
      timestamp: a.timestamp,
      beforeRevision: e.revision,
      afterRevision: o.revision,
      sheetId: r.scope.sheetId,
      range: r.scope.range,
      commands: r.compatibility.commands.length,
      changedCells: r.changes.length,
      sourceLines: r.compatibility.commands.map((u) => u.line),
    };
  return (
    (o = rn({ ...o, automationRuns: [...(e.automationRuns ?? []), s] })),
    YT.delete(r),
    o
  );
}
function Jxe(e) {
  var Q, X, J;
  const { text: r, language: t } = Dr(),
    n = FC(),
    o = ne.useRef(null),
    a = ne.useId(),
    s = ne.useRef(null),
    d = ne.useRef(null),
    u = ne.useRef(0),
    l = ne.useRef(e);
  l.current = e;
  const [c, h] = ne.useState(e.initialInspection ?? null),
    [f, p] = ne.useState(null),
    [m, v] = ne.useState(`function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange("B2").setValue(125);
  sheet.getRange("B3").setFormula("=B2*2");
}`),
    [g, b] = ne.useState("office-script"),
    [y, w] = ne.useState("Reviewed-script.ts"),
    [E, A] = ne.useState(e.sheetId),
    [k, F] = ne.useState(e.range),
    [I, O] = ne.useState(""),
    [j, z] = ne.useState(false),
    [V, D] = ne.useState(false),
    [B, _] = ne.useState(true),
    [C, T] = ne.useState(""),
    N = ne.useMemo(
      () =>
        c != null && c.script ? SW(c.script.text, c.script.language) : null,
      [c],
    ),
    R = { sheetId: E, range: k },
    L = () => {
      (p(null), D(false), O(""));
    },
    $ = () => {
      var q;
      ((q = o.current) == null || q.abort(), u.current++, z(false));
    };
  (ne.useEffect(() => {
    var W;
    const q = document.activeElement;
    return (
      (W = s.current) == null || W.focus(),
      () => {
        var K;
        ((K = o.current) == null || K.abort(),
          u.current++,
          q != null && q.isConnected && q.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      L();
    }, [e.workbook, e.readOnly]));
  const U = (q) => {
    try {
      (O(""), q());
    } catch (W) {
      O(W instanceof Error ? W.message : "Automation review failed.");
    }
  };
  async function G(q) {
    var te;
    if (!q) return;
    (te = o.current) == null || te.abort();
    const W = new AbortController();
    o.current = W;
    const K = ++u.current;
    (L(), z(true), O(""));
    try {
      if (q.size > 10 * 1024 * 1024)
        throw Error("Automation files must be 10 MB or smaller.");
      const ae = await ow(q, {
        ...n,
        signal: n.signal ? AbortSignal.any([W.signal, n.signal]) : W.signal,
      });
      if (K !== u.current) return;
      (h(ae), T(""), p(null), D(!1));
    } catch (ae) {
      if (K !== u.current) return;
      O(ae instanceof Error ? ae.message : "Unable to inspect automation.");
    } finally {
      K === u.current && (z(false), d.current && (d.current.value = ""));
    }
  }
  const H = () =>
    new TextEncoder().encode(
      JSON.stringify({ inspection: c, compatibility: N, plan: f }, null, 2),
    );
  return i.jsx("div", {
    className: "lg-automation-overlay",
    children: i.jsxs("section", {
      className: "lg-automation",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": a,
      tabIndex: -1,
      ref: s,
      onKeyDown: (q) => {
        var W;
        if (
          (q.stopPropagation(),
          q.key === "Escape" && (q.preventDefault(), e.onClose()),
          q.key === "Tab")
        ) {
          const K = [
              ...(((W = s.current) == null
                ? void 0
                : W.querySelectorAll(
                    "button:not(:disabled),input:not(:disabled):not([type=file]),select:not(:disabled),textarea:not(:disabled),summary",
                  )) ?? []),
            ],
            te = K[0],
            ae = K.at(-1);
          q.shiftKey &&
          (document.activeElement === te ||
            document.activeElement === s.current)
            ? (q.preventDefault(), ae == null || ae.focus())
            : !q.shiftKey &&
              document.activeElement === ae &&
              (q.preventDefault(), te == null || te.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: a,
              children: i.jsx(M, { children: "Automation & Scripts" }),
            }),
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Inspect, preserve and migrate approved workbook operations.",
              }),
            }),
            i.jsx("button", {
              "aria-label": r("Close automation review"),
              onClick: e.onClose,
              lang: t("Close automation review"),
              "data-ui-label": "Close automation review",
              children: "×",
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-automation-scroll",
          children: [
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Imported VBA, XLM and Office Scripts never run automatically. This adapter accepts a strict subset of literal range operations; unsupported APIs block the whole migration.",
              }),
            }),
            i.jsxs("div", {
              className: "lg-automation-actions",
              children: [
                i.jsx("button", {
                  disabled: j,
                  onClick: () => {
                    var q;
                    return (q = d.current) == null ? void 0 : q.click();
                  },
                  children: i.jsx(M, { children: "Inspect automation file" }),
                }),
                i.jsx("input", {
                  ref: d,
                  type: "file",
                  hidden: true,
                  "aria-label": r("Automation source file"),
                  accept:
                    ".bas,.vba,.cls,.ts,.osts,.xlsm,.xlam,.xltm,.xlsx,.xltx,.xls,.xlsb",
                  onChange: (q) => {
                    var W;
                    return void G((W = q.target.files) == null ? void 0 : W[0]);
                  },
                  lang: t("Automation source file"),
                  "data-ui-label": "Automation source file",
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Preserved source" }),
                    i.jsxs("select", {
                      value: C,
                      onChange: (q) =>
                        U(() => {
                          var K;
                          ($(), T(q.target.value));
                          const W =
                            (K = e.workbook.automationArtifacts) == null
                              ? void 0
                              : K.find((te) => te.id === q.target.value);
                          W && (G(new File([X5(W.base64)], W.fileName)), L());
                        }),
                      children: [
                        i.jsx("option", {
                          value: "",
                          children: "Select an artifact",
                        }),
                        (Q = e.workbook.automationArtifacts) == null
                          ? void 0
                          : Q.map((q) =>
                              i.jsxs(
                                "option",
                                {
                                  value: q.id,
                                  children: [
                                    q.fileName,
                                    " · ",
                                    q.sha256.slice(0, 12),
                                  ],
                                },
                                q.id,
                              ),
                            ),
                      ],
                    }),
                  ],
                }),
                i.jsx("button", {
                  disabled: !C || e.readOnly,
                  onClick: () =>
                    U(() => {
                      (e.onRemove(C), T(""), L());
                    }),
                  children: i.jsx(M, { children: "Remove preserved source" }),
                }),
              ],
            }),
            i.jsxs("details", {
              children: [
                i.jsx("summary", {
                  children: i.jsx(M, {
                    children: "Paste source for compatibility review",
                  }),
                }),
                i.jsxs("div", {
                  className: "lg-automation-fields",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Source language" }),
                        i.jsxs("select", {
                          value: g,
                          onChange: (q) => {
                            ($(),
                              b(q.target.value),
                              w(
                                q.target.value === "vba"
                                  ? "Reviewed-module.bas"
                                  : "Reviewed-script.ts",
                              ),
                              L(),
                              h(null));
                          },
                          children: [
                            i.jsx("option", {
                              value: "office-script",
                              children: "Office Scripts",
                            }),
                            i.jsx("option", {
                              value: "vba",
                              children: "VBA module",
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Source file name" }),
                        i.jsx("input", {
                          value: y,
                          onChange: (q) => {
                            ($(), w(q.target.value), L(), h(null));
                          },
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Source code" }),
                    i.jsx("textarea", {
                      spellCheck: false,
                      rows: 8,
                      value: m,
                      onChange: (q) => {
                        ($(), v(q.target.value), L(), h(null));
                      },
                    }),
                  ],
                }),
                i.jsx("button", {
                  onClick: () =>
                    U(() => {
                      if (
                        ($(),
                        (g === "vba" && !/\.(bas|vba|cls)$/i.test(y)) ||
                          (g === "office-script" && !/\.ts$/i.test(y)))
                      )
                        throw Error(
                          "Match the source file extension to the chosen language.",
                        );
                      (G(new File([m], y)), L(), T(""));
                    }),
                  children: i.jsx(M, { children: "Inspect pasted source" }),
                }),
              ],
            }),
            j &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(M, { children: "Inspecting source bytes…" }),
              }),
            c &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(M, { children: "Automation inventory" }),
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx("strong", { children: c.fileName }),
                      " ·",
                      " ",
                      (c.fileBytes / 1024).toFixed(1),
                      " ",
                      i.jsx(M, { children: "KiB ·" }),
                      " ",
                      c.execution === "inventory-only"
                        ? "Inventory only — execution unsupported"
                        : "Migration candidate — compatibility and approval required",
                    ],
                  }),
                  i.jsxs("p", {
                    className: "lg-automation-hash",
                    children: [i.jsx(M, { children: "SHA-256: " }), c.sha256],
                  }),
                  i.jsxs("table", {
                    "aria-label": r("Automation inventory"),
                    lang: t("Automation inventory"),
                    "data-ui-label": "Automation inventory",
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Artifact" }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Kind" }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Support" }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Notes" }),
                            }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: c.items.map((q, W) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: q.path }),
                                i.jsx("td", { children: q.kind }),
                                i.jsx("td", { children: q.execution }),
                                i.jsx("td", { children: q.note }),
                              ],
                            },
                            W,
                          ),
                        ),
                      }),
                    ],
                  }),
                  !c.items.length &&
                    i.jsx("p", {
                      children: i.jsx(M, {
                        children:
                          "No known package automation parts detected. Separate or external scripts remain unknown.",
                      }),
                    }),
                  i.jsx("ul", {
                    children: c.warnings.map((q, W) =>
                      i.jsx("li", { children: q }, W),
                    ),
                  }),
                  i.jsxs("div", {
                    className: "lg-automation-actions",
                    children: [
                      i.jsx("button", {
                        disabled: e.readOnly || !c.retainedEligible,
                        onClick: () =>
                          U(() => {
                            (e.onPreserve(c), p(null), D(!1));
                          }),
                        children: i.jsx(M, {
                          children: "Preserve original source",
                        }),
                      }),
                      i.jsx("button", {
                        onClick: () =>
                          U(() => {
                            var W;
                            const q =
                              (W = l.current.workbook.automationArtifacts) ==
                              null
                                ? void 0
                                : W.find(
                                    (K) =>
                                      K.sha256 === c.sha256 &&
                                      K.fileName === c.fileName,
                                  );
                            if (!q)
                              throw Error(
                                "Preserve the original source before downloading it from the workbook.",
                              );
                            e.onDownload(X5(q.base64), q.fileName);
                          }),
                        children: i.jsx(M, {
                          children: "Download preserved original",
                        }),
                      }),
                    ],
                  }),
                  N &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("h3", {
                          children: i.jsx(M, {
                            children: "Migration compatibility",
                          }),
                        }),
                        i.jsxs("p", {
                          children: [
                            N.supported
                              ? "Supported source grammar; workbook scope checks still required."
                              : "Unsupported source — migration blocked.",
                            " ",
                            i.jsx(M, { children: "Entry point:" }),
                            " ",
                            N.entryPoint || "(not identified)",
                            ".",
                            " ",
                            N.commands.length,
                            " ",
                            i.jsx(M, { children: "recognized commands." }),
                          ],
                        }),
                        i.jsx("ul", {
                          "aria-label": r("Source compatibility issues"),
                          lang: t("Source compatibility issues"),
                          "data-ui-label": "Source compatibility issues",
                          children: N.issues.map((q, W) =>
                            i.jsxs(
                              "li",
                              {
                                children: [
                                  i.jsx(M, { children: "Line " }),
                                  q.line,
                                  ": ",
                                  q.message,
                                ],
                              },
                              W,
                            ),
                          ),
                        }),
                        i.jsx("ul", {
                          children: N.warnings.map((q, W) =>
                            i.jsx("li", { children: q }, W),
                          ),
                        }),
                        i.jsxs("details", {
                          children: [
                            i.jsx("summary", {
                              children: i.jsx(M, {
                                children: "Reviewed source text",
                              }),
                            }),
                            i.jsx("pre", {
                              children:
                                (X = c.script) == null ? void 0 : X.text,
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "lg-automation-fields",
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(M, { children: "Approved worksheet" }),
                                i.jsx("select", {
                                  value: E,
                                  onChange: (q) => {
                                    (A(q.target.value), L());
                                  },
                                  children: e.workbook.sheets
                                    .filter((q) => !q.hidden && !q.veryHidden)
                                    .map((q) =>
                                      i.jsx(
                                        "option",
                                        { value: q.id, children: q.name },
                                        q.id,
                                      ),
                                    ),
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(M, {
                                  children:
                                    "Approved range (maximum 500 cells)",
                                }),
                                i.jsx("input", {
                                  value: k,
                                  onChange: (q) => {
                                    (F(q.target.value), L());
                                  },
                                }),
                              ],
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          disabled: e.readOnly || !N.supported,
                          onClick: () =>
                            U(() => {
                              (D(!1),
                                p(
                                  Gxe(e.workbook, c, R, {
                                    readOnly: e.readOnly,
                                  }),
                                ));
                            }),
                          children: i.jsx(M, {
                            children: "Preview migrated commands",
                          }),
                        }),
                      ],
                    }),
                ],
              }),
            f &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(M, {
                      children: "Command and change review",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      f.compatibility.commands.length,
                      " ",
                      i.jsx(M, { children: "commands ·" }),
                      " ",
                      f.touchedCells,
                      " ",
                      i.jsx(M, { children: "touched cells · " }),
                      f.changes.length,
                      " ",
                      i.jsx(M, { children: "final cell writes · revision " }),
                      f.revision,
                      ".",
                    ],
                  }),
                  i.jsx("ul", {
                    children: f.capabilities.map((q) =>
                      i.jsx("li", { children: q }, q),
                    ),
                  }),
                  !!f.issues.length &&
                    i.jsx("div", {
                      role: "alert",
                      children: i.jsx("ul", {
                        children: f.issues.map((q, W) =>
                          i.jsx("li", { children: q }, W),
                        ),
                      }),
                    }),
                  i.jsxs("table", {
                    "aria-label": r("Migrated workbook commands"),
                    lang: t("Migrated workbook commands"),
                    "data-ui-label": "Migrated workbook commands",
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Source line" }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Command" }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Target" }),
                            }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: f.compatibility.commands.map((q, W) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: q.line }),
                                i.jsx("td", { children: q.kind }),
                                i.jsxs("td", {
                                  children: [
                                    q.sheet ?? "(approved worksheet)",
                                    "!",
                                    q.range,
                                  ],
                                }),
                              ],
                            },
                            W,
                          ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("table", {
                    "aria-label": r("Automation change preview"),
                    lang: t("Automation change preview"),
                    "data-ui-label": "Automation change preview",
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Cell" }),
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
                        children: f.changes.map((q) => {
                          var W;
                          return i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: q.address }),
                                i.jsx("td", {
                                  children: String(
                                    ((W = q.before) == null ? void 0 : W.raw) ??
                                      "(blank)",
                                  ),
                                }),
                                i.jsx("td", {
                                  children: String(q.after.raw ?? "(blank)"),
                                }),
                              ],
                            },
                            q.address,
                          );
                        }),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    className: "lg-automation-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: B,
                        onChange: (q) => {
                          (_(q.target.checked), D(false));
                        },
                      }),
                      i.jsx(M, {
                        children:
                          "Retain original source with the run (native JSON)",
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    className: "lg-automation-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: V,
                        disabled: !f.valid || e.readOnly,
                        onChange: (q) => D(q.target.checked),
                      }),
                      i.jsx(M, {
                        children:
                          "I approve these commands and their changes within this scope",
                      }),
                    ],
                  }),
                ],
              }),
            !!((J = e.workbook.automationRuns) != null && J.length) &&
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(M, { children: "Recorded automation runs (" }),
                      e.workbook.automationRuns.length,
                      ")",
                    ],
                  }),
                  i.jsx("ul", {
                    children: e.workbook.automationRuns.slice(-10).map((q) =>
                      i.jsxs(
                        "li",
                        {
                          children: [
                            q.sourceName,
                            " · ",
                            q.actor,
                            " · ",
                            q.timestamp,
                            " ",
                            i.jsx(M, { children: "· revision" }),
                            " ",
                            q.beforeRevision,
                            " → ",
                            q.afterRevision,
                            " · ",
                            q.changedCells,
                            " ",
                            i.jsx(M, { children: "writes · source " }),
                            q.sourceSha256,
                          ],
                        },
                        q.id,
                      ),
                    ),
                  }),
                  i.jsx("p", {
                    children: i.jsx(M, {
                      children:
                        "Native JSON contains the full retained run record. Local history is editable and is not an immutable audit service.",
                    }),
                  }),
                ],
              }),
          ],
        }),
        e.readOnly &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(M, {
              children:
                "Read-only: inspection and report download remain available; preservation and execution are disabled.",
            }),
          }),
        I && i.jsx("p", { role: "alert", children: I }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(M, { children: "Cancel" }),
            }),
            i.jsx("button", {
              disabled: !c,
              onClick: () =>
                e.onDownload(
                  H(),
                  `${(c == null ? void 0 : c.fileName) ?? "automation"}.compatibility.json`,
                ),
              children: i.jsx(M, { children: "Save automation report" }),
            }),
            i.jsx("button", {
              disabled: !V || !(f != null && f.valid) || e.readOnly,
              onClick: () =>
                U(() => {
                  !f || !V || (e.onApply(f, R, B), p(null), D(!1));
                }),
              children: i.jsx(M, { children: "Run approved migration" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const w2 = 25,
  Qxe = (e) =>
    e === null
      ? "(blank)"
      : typeof e == "boolean"
        ? String(e).toUpperCase()
        : String(e);

export { Jxe, Kxe, Qxe, Yxe, Zxe, w2 };
