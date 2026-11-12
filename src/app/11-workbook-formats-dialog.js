import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { Lg, lE, yc } from "./01-core-schemas-and-hashing.js";
import { Mt } from "./02-financial-functions.js";
import { Fc, Xoe, ct, gx, rh, rn, ut } from "./03-workbook-model.js";
import { Dr, M } from "./04-cell-input-paste-and-fill.js";
import { H1, M$, kle, ule } from "./07-xlsx-archive-reader.js";
import { Rle, gI } from "./08-xlsx-export-and-preservation.js";
import { yh } from "./09-workbook-links.js";
import { Bue, Lue, Uue, X$ } from "./10-source-import-dialog.js";

function $ue(e) {
  const { text } = Dr(),
    t = ne.useId(),
    n = ne.useRef(null),
    o = yh(),
    a = ne.useRef(e),
    s = ne.useRef(0),
    d = ne.useRef(true);
  a.current = e;
  const [u, l] = ne.useState(null),
    [c, h] = ne.useState(null),
    [f, p] = ne.useState(false),
    [m, v] = ne.useState(false),
    [g, b] = ne.useState(""),
    [y, w] = ne.useState(null),
    [E, A] = ne.useState("");
  (ne.useEffect(() => {
    var D;
    d.current = true;
    const V = document.activeElement;
    return (
      (D = n.current) == null || D.focus(),
      () => {
        ((d.current = false),
          s.current++,
          V != null && V.isConnected && V.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      var V;
      y || (V = n.current) == null || V.focus();
    }, [y]));
  const k = async (V) => {
      (b(""), A(""), v(true));
      try {
        await V();
      } catch (D) {
        d.current && b(D.message);
      } finally {
        d.current && v(false);
      }
    },
    F = (V) => {
      const D = ++s.current;
      return (
        l(null),
        h(null),
        p(false),
        w(null),
        k(async () => {
          const B = await ule(V, o.begin());
          d.current && D === s.current && (l(B), h(V));
        })
      );
    };
  ne.useEffect(() => {
    e.initialFile && F(e.initialFile);
  }, [e.initialFile]);
  const I = (V, D) =>
      k(async () => {
        if (a.current.readOnly) throw Error("This workbook is read-only.");
        const B = a.current.workbook,
          _ = ++s.current,
          C = D ? await H1(V, o.begin()) : await Bue(u, V, f, o.begin());
        if (!(!d.current || _ !== s.current)) {
          if (a.current.workbook !== B || a.current.readOnly)
            throw Error(
              "The workbook or permissions changed. Review the conversion again.",
            );
          if (!C.report)
            throw Error(
              "A compatibility report is required for format conversion.",
            );
          w({ snapshot: B, imported: C, report: C.report });
        }
      }),
    O = (V) =>
      k(async () => {
        const D = a.current.workbook,
          B = ++s.current,
          _ = V ? await Rle(D) : await gI(D);
        if (!(!d.current || B !== s.current)) {
          if (a.current.workbook !== D)
            throw Error("The workbook changed. Review the export again.");
          w({
            snapshot: D,
            blob: _.blob,
            report: { ..._.report, fileName: M$(D, V ? "xltx" : "xlsx") },
          });
        }
      }),
    j = (V, D) => (a.current.onDownload ?? Uue)({ filename: V, blob: D }),
    z = (V) =>
      k(async () => {
        const D = V ? Xoe(a.current.workbook) : Lue(u);
        await j(
          D.fileName,
          new Blob([new Uint8Array(D.bytes)], {
            type: "application/octet-stream",
          }),
        );
      });
  return y
    ? i.jsx(X$, {
        report: y.report,
        originalLabel: u
          ? `Download original ${u.format.toUpperCase()}`
          : undefined,
        blocked:
          g ||
          (y.snapshot !== e.workbook
            ? "The workbook changed. Cancel and review again."
            : y.imported && e.readOnly
              ? "This workbook is read-only."
              : m
                ? "Completing the reviewed operation."
                : undefined),
        onClose: () => {
          (w(null), b(""));
        },
        onOriginal: u && y.imported ? () => void z(false) : undefined,
        onReport: () =>
          void k(async () => {
            await j(
              y.report.fileName + ".compatibility.json",
              new Blob([JSON.stringify(y.report, null, 2)], {
                type: "application/json",
              }),
            );
          }),
        onConfirm: () =>
          void k(async () => {
            if (a.current.workbook !== y.snapshot)
              throw Error("The workbook changed. Review the conversion again.");
            if (y.imported) {
              if (a.current.readOnly)
                throw Error("This workbook is read-only.");
              (a.current.onWorkbookChange(
                y.imported.workbook,
                "Open reviewed workbook conversion",
              ),
                A(
                  "Converted workbook opened. Keep a native JSON backup to retain both source files and conversion provenance.",
                ));
            } else
              y.blob &&
                (await j(y.report.fileName, y.blob),
                A(
                  "Derived workbook downloaded. Verify fidelity in the source application.",
                ));
            w(null);
          }),
      })
    : i.jsx("div", {
        className: "lg-reports-overlay",
        children: i.jsxs("section", {
          role: "dialog",
          "aria-modal": "true",
          "aria-labelledby": t,
          lang: "en",
          tabIndex: -1,
          ref: n,
          className: "lg-reports-dialog lg-formats-dialog",
          onKeyDown: (V) => {
            var D, B;
            if (
              (V.key === "Escape" && (V.stopPropagation(), e.onClose()),
              V.key === "Tab")
            ) {
              const _ = [
                ...n.current.querySelectorAll(
                  'button:not(:disabled),input:not(:disabled),summary,[tabindex="0"]',
                ),
              ];
              V.shiftKey &&
              (document.activeElement === _[0] ||
                document.activeElement === n.current)
                ? (V.preventDefault(), (D = _.at(-1)) == null || D.focus())
                : !V.shiftKey &&
                  document.activeElement === _.at(-1) &&
                  (V.preventDefault(), (B = _[0]) == null || B.focus());
            }
          },
          children: [
            i.jsxs("header", {
              children: [
                i.jsxs("div", {
                  children: [
                    i.jsx("p", { children: "WORKBOOK EXCHANGE" }),
                    i.jsx("h2", {
                      id: t,
                      children: i.jsx(M, { children: "Workbook formats" }),
                    }),
                  ],
                }),
                i.jsx("button", {
                  "aria-label": text("Close workbook formats"),
                  onClick: e.onClose,
                  children: i.jsx(X, { size: 20 }),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "lg-reports-scroll",
              children: [
                i.jsx("p", {
                  className: "lg-reports-intro",
                  children:
                    "Open a template or bring an older workbook into this editor. Review conversion losses and keep the untouched original with your work.",
                }),
                g &&
                  i.jsx("p", {
                    role: "alert",
                    className: "lg-formats-error",
                    children: g,
                  }),
                E && i.jsx("p", { role: "status", children: E }),
                m &&
                  i.jsx("p", {
                    role: "status",
                    children: "Processing workbook…",
                  }),
                i.jsxs("div", {
                  className: "lg-formats-actions",
                  children: [
                    i.jsxs("label", {
                      children: [
                        "Inspect source file",
                        i.jsx("input", {
                          type: "file",
                          accept: ".xls,.xlsb,.xlsm,.xltx,.xltm,.ods",
                          "aria-label": text("Inspect alternate workbook"),
                          disabled: m,
                          onChange: (V) => {
                            const D = V.target.files?.[0];
                            ((V.target.value = ""), D && F(D));
                          },
                        }),
                      ],
                    }),
                    i.jsx("button", {
                      disabled: m,
                      onClick: () => void O(false),
                      children: "Review XLSX export",
                    }),
                    i.jsx("button", {
                      disabled: m,
                      onClick: () => void O(true),
                      children: "Review XLTX export",
                    }),
                  ],
                }),
                e.workbook.alternateSource &&
                  i.jsxs("section", {
                    className: "lg-formats-source",
                    children: [
                      i.jsx("h3", { children: "Retained original" }),
                      i.jsxs("p", {
                        children: [
                          e.workbook.alternateSource.fileName,
                          " ·",
                          " ",
                          e.workbook.alternateSource.conversion.method,
                          " · equivalence not verified",
                        ],
                      }),
                      i.jsxs("p", {
                        className: "lg-formats-hash",
                        children: [
                          "SHA-256: ",
                          e.workbook.alternateSource.sha256,
                        ],
                      }),
                      i.jsx("button", {
                        disabled: m,
                        onClick: () => void z(true),
                        children: "Download retained original",
                      }),
                    ],
                  }),
                u &&
                  i.jsxs("section", {
                    className: "lg-formats-source",
                    children: [
                      i.jsx("h3", { children: u.fileName }),
                      i.jsxs("p", {
                        children: [
                          u.format.toUpperCase(),
                          " identified by content ·",
                          " ",
                          (u.fileBytes / 1024).toFixed(1),
                          " KiB ·",
                          " ",
                          u.parts.length,
                          " inventory entries",
                        ],
                      }),
                      i.jsx("p", { children: u.identity }),
                      i.jsxs("p", {
                        className: "lg-formats-hash",
                        children: ["SHA-256: ", u.sha256],
                      }),
                      i.jsxs("div", {
                        className: "lg-formats-actions",
                        children: [
                          i.jsx("button", {
                            disabled: m,
                            onClick: () => void z(false),
                            children: "Download inspected original",
                          }),
                          i.jsx("button", {
                            disabled: m,
                            onClick: () =>
                              void k(async () => {
                                await j(
                                  u.fileName + ".format-review.json",
                                  new Blob([JSON.stringify(u, null, 2)], {
                                    type: "application/json",
                                  }),
                                );
                              }),
                            children: "Download format report",
                          }),
                          e.onAutomation &&
                            ["xls", "xlsb", "xlsm", "xltx", "xltm"].includes(
                              u.format,
                            ) &&
                            i.jsx("button", {
                              disabled: m,
                              onClick: () =>
                                void k(async () => {
                                  await e.onAutomation(c);
                                }),
                              children: "Inspect automation",
                            }),
                        ],
                      }),
                      i.jsx("ul", {
                        children: u.warnings.map((V) =>
                          i.jsx("li", { children: V }, V),
                        ),
                      }),
                      u.format === "xltx" &&
                        i.jsx("p", {
                          children: i.jsx("button", {
                            className: "lg-reports-primary",
                            disabled: m || e.readOnly,
                            onClick: () => void I(c, true),
                            children: "Preview template workbook",
                          }),
                        }),
                      i.jsxs("details", {
                        open: u.format !== "xltx",
                        children: [
                          i.jsx("summary", {
                            children: "Convert with the source application",
                          }),
                          i.jsx("ol", {
                            children: u.steps.map((V) =>
                              i.jsx("li", { children: V }, V),
                            ),
                          }),
                          i.jsxs("label", {
                            className: "lg-formats-ack",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: f,
                                disabled: m || e.readOnly,
                                onChange: (V) => p(V.target.checked),
                              }),
                              "I acknowledge possible conversion losses and will compare formulas, dates, styles and special parts with the original.",
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              "Choose converted XLSX",
                              i.jsx("input", {
                                type: "file",
                                accept: ".xlsx",
                                "aria-label": text("Choose converted XLSX"),
                                disabled: m || e.readOnly || !f,
                                onChange: (V) => {
                                  const D = V.target.files?.[0];
                                  ((V.target.value = ""), D && I(D, false));
                                },
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("details", {
                        children: [
                          i.jsxs("summary", {
                            children: [
                              "Source inventory (",
                              u.parts.length,
                              ")",
                            ],
                          }),
                          i.jsx("p", {
                            children:
                              "Original-only inventory; no content is executed. This is not a lossless conversion claim.",
                          }),
                          i.jsx("ul", {
                            className: "lg-formats-inventory",
                            children: u.parts.map((V, D) =>
                              i.jsxs(
                                "li",
                                {
                                  children: [
                                    V.path,
                                    " · ",
                                    V.bytes,
                                    " known bytes ·",
                                    " ",
                                    V.disposition,
                                  ],
                                },
                                D,
                              ),
                            ),
                          }),
                        ],
                      }),
                    ],
                  }),
                i.jsx("h3", {
                  children: "Read, write and preservation support",
                }),
                i.jsx("div", {
                  className: "lg-formats-table",
                  children: i.jsxs("table", {
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", { scope: "col", children: "Format" }),
                            i.jsx("th", { scope: "col", children: "Read" }),
                            i.jsx("th", { scope: "col", children: "Write" }),
                            i.jsx("th", { scope: "col", children: "Preserve" }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: kle.map((V) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsxs("th", {
                                  scope: "row",
                                  children: [".", V.format],
                                }),
                                i.jsx("td", { children: V.read }),
                                i.jsx("td", { children: V.write }),
                                i.jsx("td", { children: V.preserve }),
                              ],
                            },
                            V.format,
                          ),
                        ),
                      }),
                    ],
                  }),
                }),
                i.jsx("p", {
                  children:
                    "Files: up to 10 MiB. ZIP packages: up to 2,000 entries and 50 MiB expanded. Password-protected containers require an unencrypted copy. External conversions and native Office fidelity must be validated in the source application.",
                }),
              ],
            }),
            i.jsxs("footer", {
              children: [
                i.jsx("span", {
                  children: e.readOnly
                    ? "Read-only: inspection and export available; import disabled."
                    : "Imports replace the current workbook after review. Undo restores the previous workbook.",
                }),
                i.jsx("button", { onClick: e.onClose, children: "Close" }),
              ],
            }),
          ],
        }),
      });
}
class Oo extends Error {
  constructor(t, n, o) {
    super(
      o
        ? `${n} Current saved workbook ${o.workbookId}, revision ${o.revision}.`
        : n,
    );
    yc(this, "currentRevision");
    yc(this, "currentWorkbookId");
    ((this.code = t),
      (this.name = "PersistenceError"),
      (this.currentRevision = o?.revision),
      (this.currentWorkbookId = o?.workbookId));
  }
}
const X4 = "ledgerglass-local-workbook",
  Nk = 20,
  N1 = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    Object.getPrototypeOf(e) === Object.prototype,
  dd = (e) => structuredClone(e),
  k7 = (e) => new TextEncoder().encode(e).byteLength;
function Mi(e, r, t) {
  throw new Oo(e, r, t);
}
function Uu(e) {
  return e === undefined
    ? "undefined"
    : e === null || typeof e != "object"
      ? JSON.stringify(e)
      : Array.isArray(e)
        ? `[${e.map(Uu).join(",")}]`
        : `{${Object.keys(e)
            .sort()
            .filter((r) => e[r] !== undefined)
            .map((r) => `${JSON.stringify(r)}:${Uu(e[r])}`)
            .join(",")}}`;
}
function i3(e) {
  let r = 2654435769,
    t = 2246822507,
    n = 3266489909,
    o = 668265263;
  for (let a = 0; a < e.length; a++) {
    const s = e.charCodeAt(a);
    ((r = Math.imul(r ^ s, 2246822507)),
      (t = Math.imul(t ^ s, 3266489909)),
      (n = Math.imul(n ^ s, 668265263)),
      (o = Math.imul(o ^ s, 374761393)));
  }
  return [e.length, r, t, n, o]
    .map((a) => (a >>> 0).toString(16).padStart(8, "0"))
    .join("");
}
function Ag(e) {
  return rn(e);
}
function G$(e, r) {
  const t = e.sheets.find((d) => !d.hidden);
  if (!t) return Mi("invalid", "A saved workbook needs a visible worksheet.");
  if (r === undefined)
    return { activeSheetId: t.id, selection: { sheetId: t.id, range: "A1" } };
  if (!N1(r) || typeof r.activeSheetId != "string" || !N1(r.selection))
    return Mi("invalid", "Invalid saved worksheet selection.");
  const n = e.sheets.find((d) => d.id === r.activeSheetId && !d.hidden);
  if (
    !n ||
    r.selection.sheetId !== n.id ||
    typeof r.selection.range != "string"
  )
    return Mi(
      "invalid",
      "Saved selection must belong to the active visible worksheet.",
    );
  let o;
  try {
    o = ct(r.selection.range);
  } catch {
    return Mi("invalid", "Invalid saved cell range.");
  }
  if (o.lastRow >= n.rowCount || o.lastCol >= n.colCount)
    return Mi("invalid", "Saved selection exceeds worksheet dimensions.");
  const a = `${ut(o.firstCol)}${o.firstRow + 1}`,
    s = `${ut(o.lastCol)}${o.lastRow + 1}`;
  return {
    activeSheetId: n.id,
    selection: { sheetId: n.id, range: a === s ? a : `${a}:${s}` },
  };
}
function H2(e, r, t = new Date().toISOString()) {
  const n = Ag(e),
    o = G$(n, r),
    a = { workbook: n, view: o, savedAt: t };
  return { ...a, digest: i3(Uu(a)) };
}
function Rk(e) {
  if (
    !N1(e) ||
    !N1(e.workbook) ||
    typeof e.savedAt != "string" ||
    !Number.isFinite(Date.parse(e.savedAt)) ||
    typeof e.digest != "string"
  )
    return Mi("corrupt", "Stored workbook snapshot is incomplete.");
  if (
    i3(Uu({ workbook: e.workbook, view: e.view, savedAt: e.savedAt })) !==
    e.digest
  )
    return Mi(
      "corrupt",
      "Stored workbook snapshot failed its integrity check.",
    );
  try {
    return H2(Ag(e.workbook), e.view, e.savedAt);
  } catch {
    return Mi(
      "corrupt",
      "Stored workbook snapshot does not satisfy workbook validation.",
    );
  }
}
function S7(e) {
  return {
    id: e.id,
    label: e.label,
    createdAt: e.createdAt,
    workbookId: e.snapshot.workbook.id,
    revision: e.snapshot.workbook.revision,
  };
}
function fT(e, r) {
  const t = Ag(e),
    n = Ag(r);
  if (t.id !== n.id)
    return Mi(
      "invalid",
      "Version comparison requires the same workbook identity.",
    );
  const o = {
      fromRevision: t.revision,
      toRevision: n.revision,
      cellChanges: [],
      metadataChanges: [],
      structuralChanges: [],
    },
    a = (m, v) => Uu(m) !== Uu(v),
    s = new Set(t.events.map((m) => m.id)),
    d = n.events.filter((m) => !s.has(m.id)),
    u = new Map();
  for (const m of d)
    for (const v of m.changes)
      u.set(`${v.sheetId}:${v.address}`, { event: m, change: v });
  const l = (m, v) =>
      m
        ? { basis: v, actor: m.actor, timestamp: m.timestamp, eventId: m.id }
        : { basis: "unavailable" },
    c = (m) =>
      l(
        [...d].reverse().find((v) => m.test(v.action)),
        "operation-history",
      ),
    h = () =>
      c(
        /sheet|structure|tab color|visibility|restore|row|column|merge|drawing|picture|shape|chart|validation|conditional|page|table|pivot|query/i,
      );
  for (const m of new Set([...Object.keys(t), ...Object.keys(n)])) {
    if (["sheets", "events", "revision", "id", "schemaVersion"].includes(m))
      continue;
    const v = t[m],
      g = n[m];
    a(v, g) &&
      o.metadataChanges.push({
        path: m,
        before: dd(v),
        after: dd(g),
        attribution: c(
          m === "definedNames"
            ? /name|formula|restore/i
            : m === "formulaWatches"
              ? /watch|row|column|sheet|move|cut|restore/i
              : m === "styleCatalog"
                ? /theme|style|restore/i
                : /workbook|properties|engagement|period|template|restore/i,
        ),
      });
  }
  a(
    t.sheets.map((m) => m.id),
    n.sheets.map((m) => m.id),
  ) &&
    o.structuralChanges.push({
      path: "sheets.order",
      before: t.sheets.map((m) => m.id),
      after: n.sheets.map((m) => m.id),
      attribution: h(),
    });
  for (const m of new Set([
    ...t.sheets.map((v) => v.id),
    ...n.sheets.map((v) => v.id),
  ])) {
    const v = t.sheets.find((E) => E.id === m),
      g = n.sheets.find((E) => E.id === m),
      b = (E) =>
        E
          ? Object.fromEntries(Object.entries(E).filter(([A]) => A !== "cells"))
          : undefined,
      y = b(v),
      w = b(g);
    if (!v || !g)
      o.structuralChanges.push({
        path: `sheets.${m}`,
        before: y,
        after: w,
        attribution: h(),
      });
    else
      for (const E of new Set([...Object.keys(y), ...Object.keys(w)]))
        a(y[E], w[E]) &&
          o.structuralChanges.push({
            path: `sheets.${m}.${E}`,
            before: dd(y[E]),
            after: dd(w[E]),
            attribution: h(),
          });
    for (const E of new Set([
      ...Object.keys(v?.cells ?? {}),
      ...Object.keys(g?.cells ?? {}),
    ])) {
      const A = v?.cells[E],
        k = g?.cells[E];
      a(A, k) &&
        o.cellChanges.push({
          sheetId: m,
          address: E,
          ...(A === undefined ? {} : { before: dd(A) }),
          ...(k === undefined ? {} : { after: dd(k) }),
          attribution: l(
            Uu(u.get(`${m}:${E}`)?.change.after) === Uu(k)
              ? u.get(`${m}:${E}`)?.event
              : undefined,
            "cell-history",
          ),
        });
    }
  }
  return o;
}
function BC(e, r, t = {}) {
  if (typeof r != "string" || !r.trim() || r.length > 300)
    return Mi(
      "invalid",
      "A local workbook namespace must contain 1–300 characters.",
    );
  const n = t.maxCheckpoints ?? Nk,
    o = t.maxBytes ?? 10 * 1024 * 1024;
  if (
    !Number.isInteger(n) ||
    n < 1 ||
    n > Nk ||
    !Number.isInteger(o) ||
    o < 1024 ||
    o > 50 * 1024 * 1024
  )
    return Mi(
      "invalid",
      "Use 1–20 checkpoints and a local envelope size between 1 KiB and 50 MiB.",
    );
  function a() {
    try {
      return e.getItem(r);
    } catch {
      return Mi(
        "unavailable",
        "Browser storage is unavailable. Keep the workbook open and export your unsaved changes.",
      );
    }
  }
  function s() {
    const h = a();
    if (h === null) return null;
    if (h.length > o || k7(h) > o)
      return Mi(
        "limit",
        "The saved local workbook exceeds this store's size limit. Preserve it for export or increase the configured limit.",
      );
    let f;
    try {
      f = JSON.parse(h);
    } catch {
      return Mi(
        "corrupt",
        "The saved workbook cannot be read. Its original storage has been preserved; export in-memory work before choosing another namespace.",
      );
    }
    if (!N1(f)) return Mi("corrupt", "Invalid local workbook envelope.");
    let p,
      m = false,
      v;
    if (f.schemaVersion === 1) {
      let y;
      try {
        y = Ag(f);
      } catch {
        return Mi(
          "corrupt",
          "The legacy saved workbook is invalid; original storage has been preserved.",
        );
      }
      const w = H2(y, G$(y));
      ((p = {
        format: X4,
        version: 1,
        generation: 0,
        current: w,
        checkpoints: [],
      }),
        (v =
          "Legacy workbook loaded. The next successful save will add local recovery and version metadata."));
    } else {
      if (
        f.format !== X4 ||
        f.version !== 1 ||
        !Number.isSafeInteger(f.generation) ||
        Number(f.generation) < 1 ||
        !Array.isArray(f.checkpoints) ||
        f.checkpoints.length > Nk
      )
        return Mi(
          "corrupt",
          "Unsupported or corrupt local workbook envelope; stored data was preserved.",
        );
      let y, w;
      if (f.previous !== undefined)
        try {
          w = Rk(f.previous);
        } catch {
          v =
            "The previous recovery snapshot is invalid; the current workbook remains available.";
        }
      try {
        y = Rk(f.current);
      } catch {
        if (!w)
          return Mi(
            "corrupt",
            "Both saved workbook snapshots are invalid. Original storage has been preserved for recovery.",
          );
        ((y = w),
          (m = true),
          (v =
            "Recovered the previous complete save because the latest snapshot is invalid. Review it before saving; the original storage is unchanged."));
      }
      const E = new Set(),
        A = [];
      for (const k of f.checkpoints) {
        if (
          !N1(k) ||
          typeof k.id != "string" ||
          !k.id ||
          k.id.length > 100 ||
          E.has(k.id) ||
          typeof k.label != "string" ||
          !k.label.trim() ||
          k.label.length > 120 ||
          typeof k.createdAt != "string" ||
          !Number.isFinite(Date.parse(k.createdAt))
        )
          return Mi(
            "corrupt",
            "Saved checkpoint metadata is invalid; original storage was preserved.",
          );
        let F;
        try {
          F = Rk(k.snapshot);
        } catch {
          return Mi(
            "corrupt",
            "A saved checkpoint is corrupt; original storage was preserved.",
          );
        }
        if (F.workbook.id !== y.workbook.id)
          return Mi(
            "corrupt",
            "A checkpoint belongs to another workbook; original storage was preserved.",
          );
        (E.add(k.id),
          A.push({
            id: k.id,
            label: k.label,
            createdAt: k.createdAt,
            snapshot: F,
          }));
      }
      p = {
        format: X4,
        version: 1,
        generation: Number(f.generation),
        current: y,
        ...(w ? { previous: w } : {}),
        checkpoints: A,
      };
    }
    const { workbook, view } = p.current;
    return {
      envelope: p,
      result: {
        workbook: dd(workbook),
        view: dd(view),
        guard: {
          namespace: r,
          generation: p.generation,
          workbookId: workbook.id,
          revision: workbook.revision,
          digest: i3(h),
        },
        recovered: m,
        ...(v ? { warning: v } : {}),
        checkpoints: p.checkpoints.map(S7),
      },
    };
  }
  function d(h) {
    const f = s();
    return h === null
      ? f
        ? Mi(
            "conflict",
            "This local namespace already contains a workbook. Reload it before saving; unsaved work remains in memory.",
            f.result.guard,
          )
        : null
      : !f || h.namespace !== r || Uu(h) !== Uu(f.result.guard)
        ? Mi(
            "conflict",
            "Another workbook window saved a different version. Reload or export your unsaved changes before resolving the conflict.",
            f?.result.guard,
          )
        : f;
  }
  function u(h, f) {
    const p = JSON.stringify(h);
    if (p.length > o || k7(p) > o)
      return Mi(
        "limit",
        "Local version storage exceeds its configured size limit. Export the workbook or use a larger host store; the previous save is unchanged.",
      );
    d(f);
    try {
      e.setItem(r, p);
    } catch (m) {
      const v =
        m instanceof Error &&
        /quota|space|full/i.test(`${m.name} ${m.message}`);
      return Mi(
        v ? "quota" : "unavailable",
        "Local save failed. The previous complete save is unchanged; keep this workbook open and export your unsaved changes.",
      );
    }
    return {
      workbook: dd(h.current.workbook),
      view: dd(h.current.view),
      guard: {
        namespace: r,
        generation: h.generation,
        workbookId: h.current.workbook.id,
        revision: h.current.workbook.revision,
        digest: i3(p),
      },
      recovered: false,
      checkpoints: h.checkpoints.map(S7),
    };
  }
  function l(h) {
    const f = (h?.envelope.generation ?? 0) + 1;
    return Number.isSafeInteger(f)
      ? f
      : Mi("limit", "Local version generation limit reached.");
  }
  function c(h, f) {
    const p = h?.envelope.checkpoints.find((m) => m.id === f);
    return (
      p ||
      Mi(
        "invalid",
        "This checkpoint is unavailable or has expired from the bounded local history.",
      )
    );
  }
  return {
    load: () => {
      return s()?.result ?? null;
    },
    save(h, f, p, m = {}) {
      const v = d(p),
        g = H2(h, f),
        b = !v || v.result.workbook.id === g.workbook.id;
      return !b && !m.allowIdentityChange
        ? Mi(
            "conflict",
            "Opening another workbook requires an explicit namespace replacement or a separate namespace.",
            v?.result.guard,
          )
        : b && v && g.workbook.revision < v.result.workbook.revision
          ? Mi(
              "conflict",
              "An older workbook revision cannot overwrite a newer saved revision. Restore a checkpoint as a new revision instead.",
              v.result.guard,
            )
          : b &&
              v &&
              g.workbook.revision === v.result.workbook.revision &&
              Uu(g.workbook) !== Uu(v.result.workbook)
            ? Mi(
                "conflict",
                "Workbook content changed without a new revision. Commit the edit before saving.",
                v.result.guard,
              )
            : u(
                {
                  format: X4,
                  version: 1,
                  generation: l(v),
                  current: g,
                  ...(v ? { previous: v.envelope.current } : {}),
                  checkpoints: b ? (v?.envelope.checkpoints ?? []) : [],
                },
                p,
              );
    },
    checkpoint(h, f) {
      if (typeof h != "string" || !h.trim() || h.length > 120)
        return Mi("invalid", "Checkpoint label must contain 1–120 characters.");
      const p = d(f),
        m = {
          id: crypto.randomUUID(),
          label: h.trim(),
          createdAt: new Date().toISOString(),
          snapshot: dd(p.envelope.current),
        },
        v = [...p.envelope.checkpoints, m].slice(-n),
        g = u({ ...p.envelope, generation: l(p), checkpoints: v }, f);
      return (
        p.envelope.checkpoints.length >= n &&
          (g.warning = `The oldest checkpoint expired. This local store retains the newest ${n} checkpoints.`),
        g
      );
    },
    listCheckpoints: () => {
      return s()?.result.checkpoints ?? [];
    },
    compareCheckpoint: (h, f) => fT(c(s(), h).snapshot.workbook, f),
    restoreCheckpoint(h, f, p, m = "Auditor") {
      var k;
      const v = d(p),
        g = Ag(f);
      if (
        g.id !== v.result.workbook.id ||
        g.revision !== v.result.workbook.revision ||
        Uu(g) !== Uu(v.result.workbook)
      )
        return Mi(
          "conflict",
          "Save or export your current unsaved changes before restoring a checkpoint.",
          v.result.guard,
        );
      if (
        ((k = g.structureProtection) != null && k.enabled) ||
        g.sheets.some((I) => {
          return I.protection?.enabled;
        })
      )
        return Mi(
          "invalid",
          "Unlock the workbook structure and protected worksheets before restoring a version.",
        );
      if (typeof m != "string" || !m.trim() || m.length > 100)
        return Mi("invalid", "A restore actor must contain 1–100 characters.");
      const b = c(v, h),
        y = fT(g, b.snapshot.workbook),
        w = new Date().toISOString();
      let E = dd(b.snapshot.workbook);
      if (
        ((E.revision = g.revision + 1),
        !Number.isSafeInteger(E.revision) ||
          E.revision >= Number.MAX_SAFE_INTEGER)
      )
        return Mi("limit", "Workbook revision limit reached.");
      (g.metadata &&
        (E.metadata = {
          ...(E.metadata ?? g.metadata),
          createdAt: g.metadata.createdAt,
          modifiedAt: [g.metadata.modifiedAt, E.metadata?.modifiedAt ?? w, w]
            .sort()
            .at(-1),
        }),
        (E = gx(E, w)));
      const A = Math.max(1, Math.ceil(y.cellChanges.length / 500));
      return (
        (E.events = [
          ...g.events,
          ...Array.from({ length: A }, (I, O) => ({
            id: crypto.randomUUID(),
            timestamp: w,
            actor: m,
            action: `Restored checkpoint “${b.label}” from revision ${b.snapshot.workbook.revision}${A > 1 ? ` (${O + 1}/${A})` : ""}`,
            changes: y.cellChanges
              .slice(O * 500, (O + 1) * 500)
              .map(({ attribution, ...z }) => z),
          })),
        ]),
        (E = Ag(E)),
        u(
          {
            ...v.envelope,
            generation: l(v),
            previous: v.envelope.current,
            current: H2(E, b.snapshot.view, w),
          },
          p,
        )
      );
    },
  };
}
function zue(e = {}) {
  let r = null,
    t = null;
  const n = `session-${crypto.randomUUID()}`,
    o = BC(
      {
        getItem: () => r,
        setItem: (f, p) => {
          r = p;
        },
      },
      n,
      e,
    );
  let a = null;
  const s = new Map(),
    d = e.maxBytes ?? 10 * 1024 * 1024,
    u = (f, p) => f === p || JSON.stringify(f) === JSON.stringify(p);
  function l(f) {
    if (!u(f, t?.guard ?? null))
      throw new Oo(
        "conflict",
        "The saved session changed. Reload before replacing it.",
        t?.guard,
      );
  }
  function c(f) {
    const p = (t?.guard.generation ?? 0) + 1;
    if (!Number.isSafeInteger(p))
      throw new Oo("limit", "Session generation limit reached.");
    return (
      (t = Lg({
        ...f,
        workbook: rh(f.workbook),
        view: structuredClone(f.view),
        guard: {
          namespace: n,
          generation: p,
          workbookId: f.workbook.id,
          revision: f.workbook.revision,
          digest: crypto.randomUUID(),
        },
      })),
      t
    );
  }
  function h() {
    if (!t)
      throw new Oo("invalid", "Save this session before creating a version.");
    const f = o.save(t.workbook, t.view, a, { allowIdentityChange: true });
    return ((a = f.guard), f);
  }
  return {
    load: () => t,
    save(f, p, m, v = {}) {
      l(m);
      const g = rh(f),
        b = g.sheets.find((I) => I.id === p.activeSheetId && !I.hidden);
      if (!b || p.selection.sheetId !== b.id)
        throw new Oo(
          "invalid",
          "Saved selection must belong to the active visible worksheet.",
        );
      const y = !t || t.workbook.id === g.id;
      if (!y && !v.allowIdentityChange)
        throw new Oo(
          "conflict",
          "Opening another workbook requires an explicit session replacement.",
          t?.guard,
        );
      if (
        y &&
        t &&
        (g.revision < t.workbook.revision ||
          (g.revision === t.workbook.revision && !u(g, t.workbook)))
      )
        throw new Oo(
          "conflict",
          "Commit a new revision before replacing saved content.",
          t.guard,
        );
      const w = Vue(g, p),
        E = (I, O) => ({
          workbook: I,
          view: O,
          savedAt: "1970-01-01T00:00:00.000Z",
          digest: "0".repeat(40),
        }),
        A = y ? (t?.checkpoints ?? []) : [],
        k = A.reduce((I, O) => I + (s.get(O.id) ?? 0), 0);
      if (
        lE({
          format: "ledgerglass-local-workbook",
          version: 1,
          generation: (t?.guard.generation ?? 0) + 1,
          current: E(g, w),
          ...(t ? { previous: E(t.workbook, t.view) } : {}),
          checkpoints: [],
        }) +
          k +
          Math.max(0, A.length - 1) >
        d
      )
        throw new Oo(
          "limit",
          "Session version storage exceeds its configured size limit. Export the workbook or use a larger host store; the previous save is unchanged.",
        );
      return (
        y || s.clear(),
        c({
          workbook: g,
          view: w,
          guard: m,
          recovered: false,
          checkpoints: y ? (t?.checkpoints ?? []) : [],
        })
      );
    },
    checkpoint(f, p) {
      l(p);
      const m = h(),
        v = o.checkpoint(f, m.guard);
      a = v.guard;
      const g = new Set(v.checkpoints.map((b) => b.id));
      for (const b of s.keys()) g.has(b) || s.delete(b);
      for (const b of g)
        s.has(b) ||
          s.set(
            b,
            lE({
              id: b,
              label: v.checkpoints.find((y) => y.id === b).label,
              createdAt: "1970-01-01T00:00:00.000Z",
              snapshot: {
                workbook: v.workbook,
                view: v.view,
                savedAt: "1970-01-01T00:00:00.000Z",
                digest: "0".repeat(40),
              },
            }),
          );
      return c(v);
    },
    listCheckpoints: () => structuredClone(t?.checkpoints ?? []),
    compareCheckpoint: (f, p) => o.compareCheckpoint(f, p),
    restoreCheckpoint(f, p, m, v) {
      if ((l(m), !t || !u(p, t.workbook)))
        throw new Oo(
          "conflict",
          "Save the current revision before restoring a checkpoint.",
          t?.guard,
        );
      const g = h(),
        b = o.restoreCheckpoint(f, p, g.guard, v);
      return ((a = b.guard), c(b));
    },
  };
}
function Vue(e, r) {
  const t = e.sheets.find((s) => s.id === r.activeSheetId),
    n = ct(r.selection.range);
  if (n.lastRow >= t.rowCount || n.lastCol >= t.colCount)
    throw new Oo("invalid", "Saved selection exceeds worksheet dimensions.");
  const o = `${ut(n.firstCol)}${n.firstRow + 1}`,
    a = `${ut(n.lastCol)}${n.lastRow + 1}`;
  return {
    activeSheetId: t.id,
    selection: { sheetId: t.id, range: o === a ? o : `${o}:${a}` },
  };
}
const Qb = (e, r) => JSON.stringify([e, r]),
  o3 = (e, r) => `${ut(r)}${e + 1}`,
  ey = (e) =>
    !!e && !e.literal && typeof e.raw == "string" && e.raw.startsWith("="),
  Wue = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/,
  hT = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_\\][A-Za-z0-9_.\\]*)`,
  Z$ = String.raw`(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?:\s*:\s*\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}\s*:\s*\$?[A-Za-z]{1,3}|\$?[1-9]\d*\s*:\s*\$?[1-9]\d*)`,
  Hue = new RegExp(`^(${hT}\\s*!\\s*)?(${Z$})(#)?`),
  que = new RegExp(`^${hT}\\s*:\\s*${hT}\\s*!\\s*${Z$}`),
  Xue = new Set(Fc.map((e) => e.toUpperCase())),
  G4 = (e, r, t, n) => {
    const o = e ?? r;
    if (!Number.isInteger(o) || o < 1 || o > t)
      throw Error(`${n} must be an integer from 1 to ${t}.`);
    return o;
  };
function T7(e) {
  const r = e.replace(/[\s$]/g, "").toUpperCase(),
    t = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(r),
    n = /^([1-9]\d*):([1-9]\d*)$/.exec(r),
    o = ct(t ? `${t[1]}1:${t[2]}${Mt.maxRows}` : n ? `A${n[1]}:XFD${n[2]}` : r),
    a = o3(o.firstRow, o.firstCol),
    s = o3(o.lastRow, o.lastCol);
  return a === s ? a : `${a}:${s}`;
}

export { $ue, BC, G4, Hue, Oo, Qb, T7, Wue, Xue, ey, fT, o3, que, zue };
