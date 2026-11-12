import Ox from "saxes";
import {
  strFromU8,
  strToU8,
  zipSync,
  Unzip,
  UnzipInflate,
  unzipSync,
} from "fflate";
import iue from "exceljs";
import lue from "papaparse";
import { ax, iu, ts, wd } from "./01-core-schemas-and-hashing.js";
import { Ja, Mt } from "./02-financial-functions.js";
import {
  $a,
  Dg,
  E1,
  EC,
  GE,
  J3,
  Nx,
  Og,
  Q3,
  RC,
  VR,
  Vg,
  Vn,
  WR,
  Zoe,
  _t,
  aP,
  aae,
  cn,
  ct,
  dne,
  eie,
  f0,
  md,
  rn,
  s0,
  ut,
  wie,
  x1,
  ys,
} from "./03-workbook-model.js";
import { T1 } from "./04-cell-input-paste-and-fill.js";
import {
  Ale,
  Cle,
  D$,
  Ele,
  I$,
  M$,
  Nle,
  Tle,
  ble,
  mle,
  r3,
} from "./07-xlsx-archive-reader.js";

async function Rle(e) {
  const r = await gI(e),
    t = Tle(new Uint8Array(await Nle(r.blob)), "xlsx", "xltx"),
    o = [
      ...r.warnings,
      "Derived XLTX template: supported worksheet content only. Macro projects, native automation and unsupported special parts are not transferred. Reopening creates a new workbook; keep native JSON and original source files.",
    ];
  return {
    blob: new Blob([new Uint8Array(t)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
    }),
    warnings: o,
    report: {
      ...r.report,
      format: "xltx",
      fileName: M$(e, "xltx"),
      fileBytes: t.length,
      warnings: o,
    },
  };
}
const g7 = "http://schemas.openxmlformats.org/package/2006/relationships",
  z2 = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
  Ile = "http://schemas.openxmlformats.org/package/2006/content-types",
  _le = "application/vnd.openxmlformats-officedocument.custom-properties+xml",
  xk = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
  Dle = { scope: "workbook" },
  H4 = (e) =>
    e.replace(
      /[&<>"']/g,
      (r) =>
        ({
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
          "'": "&apos;",
        })[r],
    );
function Cg(e, r, t, n) {
  const o = new Ox.SaxesParser({ xmlns: true });
  (o.on("opentag", r),
    t && o.on("closetag", t),
    n && o.on("text", n),
    o.on("doctype", () => {
      throw Error("DTD is not permitted in preservation metadata.");
    }),
    o.write(strFromU8(e)).close());
}
function j$(e) {
  return e === "_rels/.rels"
    ? ""
    : e.replace(/(^|\/)\_rels\/([^/]+)\.rels$/, "$1$2");
}
function sT(e) {
  if (!e) return "_rels/.rels";
  const r = e.lastIndexOf("/");
  return `${e.slice(0, r + 1)}_rels/${e.slice(r + 1)}.rels`;
}
function lT(e) {
  const r = new Map(),
    t = new Map(),
    n = [],
    o = new Map();
  Cg(e["[Content_Types].xml"], (a) => {
    var s, d, u, l;
    (a.local === "Override" &&
      r.set(
        (((s = a.attributes.PartName) == null ? void 0 : s.value) ?? "").slice(
          1,
        ),
        ((d = a.attributes.ContentType) == null ? void 0 : d.value) ?? "",
      ),
      a.local === "Default" &&
        t.set(
          ((u = a.attributes.Extension) == null ? void 0 : u.value) ?? "",
          ((l = a.attributes.ContentType) == null ? void 0 : l.value) ?? "",
        ));
  });
  for (const [a, s] of Object.entries(e))
    a.endsWith(".rels") &&
      Cg(s, (d) => {
        var c, h, f, p;
        if (d.local !== "Relationship") return;
        const u =
            ((c = d.attributes.TargetMode) == null ? void 0 : c.value) ===
            "External",
          l = ((h = d.attributes.Target) == null ? void 0 : h.value) ?? "";
        n.push({
          owner: j$(a),
          rels: a,
          id: ((f = d.attributes.Id) == null ? void 0 : f.value) ?? "",
          type: ((p = d.attributes.Type) == null ? void 0 : p.value) ?? "",
          external: u,
          target: u ? l : r3(a, l),
        });
      });
  return (
    Cg(e["xl/workbook.xml"], (a) => {
      var s;
      if (a.local === "sheet") {
        const d = n.find((u) => {
          var l;
          return (
            u.owner === "xl/workbook.xml" &&
            u.id ===
              ((l = Object.values(a.attributes).find(
                (c) => c.local === "id" && c.uri === z2.slice(0, -1),
              )) == null
                ? void 0
                : l.value)
          );
        });
        d &&
          o.set(
            d.target,
            ((s = a.attributes.name) == null ? void 0 : s.value) ??
              "Unknown worksheet",
          );
      }
    }),
    {
      relationships: n,
      sheets: o,
      contentType: (a) =>
        r.get(a) ?? t.get(a.split(".").at(-1)) ?? "undeclared",
      types: r,
    }
  );
}
function Ole(e, r) {
  return e.endsWith("/")
    ? "directory"
    : e === "[Content_Types].xml" || e.endsWith(".rels")
      ? "package relationships"
      : /vba|macrosheet|activeX|embedding/i.test(e + r)
        ? "active content"
        : /^customXml\//.test(e)
          ? "custom XML"
          : e === "docProps/custom.xml"
            ? "custom properties"
            : /^xl\/charts\//.test(e)
              ? "charts"
              : /^xl\/(?:model|dataModel)\//i.test(e) ||
                  /model\+data|dataModel/i.test(r)
                ? "Excel Data Model"
                : /^xl\/(pivotTables|pivotCache)\//.test(e)
                  ? "pivots"
                  : /connections|queryTables|externalLinks/.test(e)
                    ? "connections and external links"
                    : /slicers|slicerCaches|timelines/.test(e)
                      ? "slicers and timelines"
                      : /threadedComments|persons/.test(e)
                        ? "threaded comments"
                        : /^_xmlsignatures\//.test(e)
                          ? "signatures"
                          : /^xl\/worksheets\/[^/]+\.xml$/.test(e)
                            ? "worksheet"
                            : e === "xl/workbook.xml"
                              ? "workbook and names"
                              : /^xl\/tables\//.test(e)
                                ? "tables"
                                : /^xl\/drawings\//.test(e)
                                  ? "drawings"
                                  : /^xl\/media\//.test(e)
                                    ? "media"
                                    : /^xl\/comments/.test(e)
                                      ? "notes"
                                      : e === "xl/sharedStrings.xml"
                                        ? "shared strings"
                                        : e === "xl/styles.xml" ||
                                            /^xl\/theme\//.test(e)
                                          ? "formatting"
                                          : /^docProps\//.test(e)
                                            ? "document properties"
                                            : "unknown part";
}
function F$(e, r) {
  var p;
  const t = lT(e),
    n = new Map(),
    o = new Map();
  for (const m of t.relationships)
    (n.set(m.owner, [...(n.get(m.owner) ?? []), m]),
      m.external || o.set(m.target, [...(o.get(m.target) ?? []), m]));
  const a = new Map();
  for (const [m, v] of t.sheets) {
    const g = [m],
      b = new Set();
    for (; g.length;) {
      const y = g.pop();
      if (b.has(y)) continue;
      b.add(y);
      const w = a.get(y) ?? new Set();
      (w.add(v), a.set(y, w));
      for (const E of n.get(y) ?? []) E.external || g.push(E.target);
    }
  }
  const s = new Set(),
    d = new Map(),
    u = Object.keys(e).filter(
      (m) => m.startsWith("customXml/") && !m.endsWith("/"),
    );
  if (u.length) {
    const m = new Set(u),
      v = t.relationships.filter(
        (E) => m.has(E.owner) || (!E.external && m.has(E.target)),
      );
    let g = !Object.keys(e).some((E) => /xmlMaps|singleXmlCells/i.test(E));
    (g &&
      (g = u.every(
        (E) =>
          E.endsWith(".rels") ||
          (/\.xml$/.test(E) &&
            ["application/xml", "text/xml", xk].includes(t.contentType(E))),
      )),
      g &&
        (g = v.every(
          (E) =>
            !E.external &&
            ((E.owner === "xl/workbook.xml" &&
              E.type === z2 + "customXml" &&
              m.has(E.target)) ||
              (m.has(E.owner) &&
                m.has(E.target) &&
                E.type === z2 + "customXmlProps")),
        )),
      g &&
        (g = v.every((E) =>
          E.owner === "xl/workbook.xml"
            ? ["application/xml", "text/xml"].includes(t.contentType(E.target))
            : t.contentType(E.owner) !== xk && t.contentType(E.target) === xk,
        )));
    const b = v.filter((E) => E.owner === "xl/workbook.xml"),
      y = new Set(),
      w = b.map((E) => E.target);
    for (; w.length;) {
      const E = w.pop();
      if (!y.has(E)) {
        (y.add(E), e[sT(E)] && y.add(sT(E)));
        for (const A of n.get(E) ?? []) A.external || w.push(A.target);
      }
    }
    g && (g = b.length > 0 && u.every((E) => y.has(E)));
    for (const E of u)
      g
        ? s.add(E)
        : d.set(
            E,
            "Custom XML has bindings, unrecognized types, orphaned parts or relationships outside the passive metadata island. Retained only in the original.",
          );
  }
  const l = "docProps/custom.xml";
  if (e[l]) {
    const m = o.get(l) ?? [];
    let v =
      t.contentType(l) === _le &&
      m.length === 1 &&
      m[0].owner === "" &&
      m[0].type === z2 + "custom-properties" &&
      !((p = n.get(l)) != null && p.length);
    (Cg(e[l], (g) => {
      (g.local === "property" && g.attributes.linkTarget && (v = false),
        (g.uri ===
          "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" &&
          ["Properties", "property"].includes(g.local)) ||
          (g.uri ===
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" &&
            [
              "lpwstr",
              "lpstr",
              "bstr",
              "i1",
              "i2",
              "i4",
              "i8",
              "int",
              "ui1",
              "ui2",
              "ui4",
              "ui8",
              "uint",
              "r4",
              "r8",
              "decimal",
              "bool",
              "date",
              "filetime",
              "empty",
              "null",
            ].includes(g.local)) ||
          (v = false));
    }),
      v
        ? s.add(l)
        : d.set(
            l,
            "Custom properties contain links, unsupported value types or relationships; retained only in the original.",
          ));
  }
  const c = new Set();
  if (e["xl/sharedStrings.xml"]) {
    let m = -1;
    Cg(e["xl/sharedStrings.xml"], (v) => {
      (v.local === "si" && m++, v.local === "r" && c.add(m));
    });
  }
  let h = 0;
  const f = Object.entries(e)
    .sort(([m], [v]) => m.localeCompare(v))
    .map(([m, v]) => {
      const g = Ole(m, t.contentType(m)),
        b = [...(a.get(m.endsWith(".rels") ? j$(m) : m) ?? [])].map((k) => ({
          sheet: k,
          scope: "sheet",
        }));
      b.length ||
        b.push(
          [
            "charts",
            "pivots",
            "drawings",
            "notes",
            "unknown part",
            "media",
          ].includes(g)
            ? {
                scope: "unresolved",
                detail:
                  "No worksheet relationship identifies the affected cells.",
              }
            : Dle,
        );
      let y = "dropped",
        w =
          "No safe editable mapping. Exact source bytes remain in the original file.";
      (g === "Excel Data Model" &&
        (w =
          "Imported Excel Data Model, relationships and DAX are not executed or translated. Worksheet results are historical snapshots. Exact model bytes remain only in the original source; rebuild reviewed local typed tables and relationships to calculate new results."),
        ["worksheet", "workbook and names", "tables", "notes"].includes(g) &&
          ((y = "editable"),
          (w =
            "Supported content is editable and regenerated; inspect the feature details and conversion notices for losses. XML bytes are not preserved.")),
        [
          "package relationships",
          "shared strings",
          "formatting",
          "document properties",
          "drawings",
        ].includes(g) &&
          ((y = "flattened"),
          (w =
            "Supported content is regenerated or flattened. Unsupported settings and original XML are retained only in the source.")),
        g === "media" &&
          /^image\/(png|jpeg)$/.test(t.contentType(m)) &&
          ((y = "flattened"),
          (w =
            "Supported worksheet pictures are re-encoded with their supported placement; orphaned images and chart assets are not carried over.")),
        s.has(m) &&
          ((y = "preserved-opaque"),
          (w =
            "Passive metadata snapshot copied byte-for-byte, without interpreting or refreshing its contents. Its original meaning may be stale after edits.")),
        d.has(m) && (w = d.get(m)),
        g === "signatures" &&
          (w =
            "Signatures cannot authenticate an edited workbook; retained only in the original file."));
      const E = [],
        A = (k, F, I, O, j) => {
          if (++h > VR.maxDetails)
            throw Error(
              "XLSX compatibility details exceed the 100,000-entry budget. Split the source workbook before importing.",
            );
          E.push({
            feature: k,
            disposition: F,
            location: {
              ...b[0],
              ...(I ? { range: I, scope: "cell" } : {}),
              ...(O ? { detail: O } : {}),
            },
            reason:
              j ??
              (F === "editable"
                ? "Supported subset is editable; review conversion notices for unsupported attributes."
                : "This source feature is not preserved as an editable feature in the derived workbook."),
          });
        };
      if (
        g === "worksheet" ||
        g === "workbook and names" ||
        g === "tables" ||
        g === "drawings" ||
        g === "pivots" ||
        g === "charts"
      ) {
        let k = "",
          F = "",
          I,
          O = "",
          j = false,
          z = false,
          V = "";
        Cg(
          v,
          (D) => {
            const B = (_) => {
              var C;
              return (C = D.attributes[_]) == null ? void 0 : C.value;
            };
            (D.local === "c" &&
              ((k = B("r") ?? ""),
              (F = B("t") ?? ""),
              F === "e" && A("error cell", "flattened", k)),
              D.local === "v" && ((z = true), (V = "")),
              D.local === "f" && ((j = true), (O = "")),
              D.local === "definedName" &&
                ((I = B("name")), A("defined name", "editable", void 0, I)),
              D.local === "dataValidation" &&
                A("data validation", "editable", B("sqref"), B("type")),
              D.local === "conditionalFormatting" &&
                A("conditional formatting", "editable", B("sqref")),
              D.local === "mergeCell" &&
                A("merged cells", "editable", B("ref")),
              D.local === "table" &&
                A("structured table", "editable", B("ref"), B("name")),
              D.local === "location" && A("pivot output", "dropped", B("ref")),
              D.local === "worksheetSource" &&
                A("pivot source", "dropped", B("ref"), B("sheet")),
              D.local === "r" &&
                k &&
                F === "inlineStr" &&
                A("rich text", "flattened", k),
              [
                "sheetProtection",
                "workbookProtection",
                "autoFilter",
                "extLst",
                "hyperlink",
                "legacyDrawingHF",
                "webPublishItems",
                "controls",
                "oleObjects",
                "pivotTable",
                "pivotTablePart",
                "slicer",
                "timeline",
                "sp",
                "graphicFrame",
              ].includes(D.local) &&
                A(D.local, "dropped", B("ref"), B("name")));
          },
          (D) => {
            (D.local === "v" &&
              ((z = false),
              k &&
                F === "s" &&
                c.has(Number(V)) &&
                A("rich text", "flattened", k)),
              D.local === "c" && (k = ""),
              D.local === "f" &&
                ((j = false),
                g === "charts"
                  ? A(
                      "chart source reference",
                      "dropped",
                      void 0,
                      O.slice(0, 1024),
                    )
                  : /(?:\[|https?:|file:|WEBSERVICE|HYPERLINK|RTD|DDE|CALL\s*\(|EXEC\s*\()/i.test(
                      O,
                    ) &&
                    A(
                      "external/action formula",
                      "flattened",
                      k,
                      void 0,
                      "External/action references may be neutralized to literal text; the formula cache report provides the actual conversion outcome.",
                    )));
          },
          (D) => {
            (j && (O += D), z && (V += D));
          },
        );
      }
      for (const k of n.get(m) ?? [])
        k.external &&
          A(
            "external relationship",
            "dropped",
            void 0,
            k.type.split("/").at(-1),
            "External target is never fetched or carried into the derived copy; its URI remains only in the original.",
          );
      return {
        path: m,
        contentType: m.endsWith("/") ? "directory" : t.contentType(m),
        bytes: v.length,
        sha256: iu(v),
        family: g,
        disposition: y,
        reason: w,
        locations: b,
        features: E,
        retainedInOriginal: true,
      };
    });
  return mI({
    schemaVersion: 1,
    policy: "passive-metadata-v1",
    complete: true,
    phase: "eligible-for-export",
    sourceSha256: r,
    parts: f,
    counts: { editable: 0, "preserved-opaque": 0, flattened: 0, dropped: 0 },
    opaqueBytes: 0,
    note: "All package entries are inventoried. Locations come from source relationships and XML; unresolved locations are explicit. Dispositions describe the derived copy, while all parts remain byte-exact in the retained original. Editable parts can also contain flattened or dropped features. No source part is executed or fetched.",
  });
}
function mI(e) {
  ((e.counts = {
    editable: 0,
    "preserved-opaque": 0,
    flattened: 0,
    dropped: 0,
  }),
    (e.opaqueBytes = 0));
  for (const r of e.parts)
    (e.counts[r.disposition]++,
      r.disposition === "preserved-opaque" && (e.opaqueBytes += r.bytes));
  return e;
}
function Mle(e, r, t) {
  const n = structuredClone(t),
    o = lT(e),
    a = lT(r),
    s = { ...r };
  n.phase = "export-result";
  for (const d of ["custom XML", "custom properties"]) {
    const l = n.parts
      .filter(
        (m) =>
          m.family === d ||
          (d === "custom XML" && m.path.startsWith("customXml/")),
      )
      .filter((m) => m.disposition === "preserved-opaque");
    if (!l.length) continue;
    const c = new Set(l.map((m) => m.path)),
      h = o.relationships.filter((m) => !c.has(m.owner) && c.has(m.target));
    if (
      l.some((m) => s[m.path] && iu(s[m.path]) !== m.sha256) ||
      h.some((m) =>
        a.relationships.some((v) => v.owner === m.owner && v.type === m.type),
      )
    ) {
      for (const m of l)
        ((m.disposition = "dropped"),
          (m.reason =
            "Generated package has a conflicting part or relationship. The entire metadata island remains only in the original."));
      continue;
    }
    for (const m of l) s[m.path] = e[m.path];
    for (const m of new Set(h.map((v) => v.owner))) {
      const v = sT(m),
        g = s[v] ? strFromU8(s[v]) : `<Relationships xmlns="${g7}"/>`,
        b = new Set();
      s[v] &&
        Cg(s[v], (E) => {
          var A;
          E.local === "Relationship" &&
            b.add(((A = E.attributes.Id) == null ? void 0 : A.value) ?? "");
        });
      let y = 1;
      const w = h
        .filter((E) => E.owner === m)
        .map((E) => {
          for (; b.has(`lgPreserve${y}`);) y++;
          const A = `lgPreserve${y++}`;
          return (
            b.add(A),
            `<Relationship xmlns="${g7}" Id="${A}" Type="${H4(E.type)}" Target="/${H4(E.target)}"/>`
          );
        })
        .join("");
      s[v] = strToU8(v7(g, "Relationships", w));
    }
    const p = l
      .filter((m) => !m.path.endsWith(".rels") && !a.types.has(m.path))
      .map(
        (m) =>
          `<Override xmlns="${Ile}" PartName="/${H4(m.path)}" ContentType="${H4(m.contentType)}"/>`,
      )
      .join("");
    s["[Content_Types].xml"] = strToU8(
      v7(strFromU8(s["[Content_Types].xml"]), "Types", p),
    );
  }
  return { bytes: zipSync(s), inventory: mI(n) };
}
function v7(e, r, t) {
  if (!t) return e;
  const n = new RegExp(`</((?:[\\w.-]+:)?${r})\\s*>\\s*$`);
  if (n.test(e)) return e.replace(n, `${t}$&`);
  const o = new RegExp(`<((?:[\\w.-]+:)?${r})([^>]*)/>(\\s*)$`);
  if (o.test(e)) return e.replace(o, `<$1$2>${t}</$1>$3`);
  throw Error("Unable to preserve XLSX package metadata safely.");
}
function jle(e, r) {
  const t = structuredClone(e);
  t.phase = "export-result";
  for (const n of t.parts)
    n.disposition === "preserved-opaque" &&
      ((n.disposition = "dropped"), (n.reason = r));
  return mI(t);
}
const Ek = Object.freeze({
  storage: "IEEE-754 binary64",
  excelSignificantDigits: 15,
  rounding:
    "ROUND uses decimal ties away from zero; monetary rounding is explicit, never automatic.",
  display:
    "General displays 15 significant digits. Display precision does not round stored values or calculations.",
});
function Fle(e) {
  const r = /^[+-]?(\d+\.?\d*|\.\d+)(?:e([+-]?\d+))?$/i.exec(e.trim());
  if (!r) return;
  const t = r[1].replace(".", "").replace(/^0+/, "").replace(/0+$/, ""),
    n = Number(e);
  if (!Number.isFinite(n)) return "outside the finite numeric range";
  if (t && (n === 0 || Math.abs(n) < 22250738585072014e-324))
    return "below Excel's normal floating-point range";
  if (t.length > 15) return "more than 15 significant digits";
  if (Number.isInteger(n) && !Number.isSafeInteger(n))
    return "outside the exact integer range";
}
function Ble(e) {
  const r = new Map(),
    t = new Ox.SaxesParser({ xmlns: true });
  let n = "",
    o = "",
    a = "",
    s = false;
  t.on("opentag", (u) => {
    var l, c;
    u.local === "c"
      ? ((n = String(((l = u.attributes.r) == null ? void 0 : l.value) ?? "")),
        (o = String(((c = u.attributes.t) == null ? void 0 : c.value) ?? "n")))
      : u.local === "v" && n && ((s = true), (a = ""));
  });
  const d = (u) => {
    if (s && ((a += u), a.length > 1024))
      throw Error("An XLSX numeric token exceeds 1,024 characters.");
  };
  return (
    t.on("text", d),
    t.on("cdata", d),
    t.on("closetag", (u) => {
      if (u.local === "v") {
        if (o === "n" && a.trim()) {
          const l = a.trim();
          if (
            !/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(l) ||
            !Number.isFinite(Number(l))
          )
            throw Error(`Invalid finite numeric value at ${n}.`);
          r.set(n, l);
        }
        s = false;
      } else u.local === "c" && (n = "");
    }),
    t.write(e).close(),
    r
  );
}
const n3 = (e, r) => `${ut(r)}${e + 1}`,
  Jw = (e) => {
    const r = ct(e);
    return `${n3(r.firstRow, r.firstCol)}:${n3(r.lastRow, r.lastCol)}`;
  },
  b7 = (e) => {
    const r = ct(e);
    return (r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1);
  },
  Jb = (e) => `FF${e.slice(1).toUpperCase()}`;
function Lle(e) {
  const r = [];
  for (const t of e.matchAll(
    /<(?:\w+:)?dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?dataValidation>)/g,
  )) {
    const n = (d) => {
        var u;
        return (u = new RegExp("\\b" + d + `=["\\']([^"\\']*)["\\']`).exec(
          t[1],
        )) == null
          ? void 0
          : u[1];
      },
      o = (n("sqref") ?? "").trim().split(/\s+/).filter(Boolean);
    if (!o.length) continue;
    const a = Jw(o[0]).split(":")[0];
    if (r.length + o.length > 1e3)
      throw Error(
        "XLSX has more than 1,000 validation areas; narrow the source rules before import.",
      );
    const s = [
      ...(t[2] ?? "").matchAll(
        /<(?:\w+:)?formula[12]\b[^>]*>([\s\S]*?)<\/(?:\w+:)?formula[12]>/g,
      ),
    ].map((d) => d[1].trim());
    for (const d of o)
      r.push({
        range: Jw(d),
        anchor: a,
        operands: s,
        dropdown: !["1", "true"].includes(n("showDropDown") ?? "0"),
        alert: ["1", "true"].includes(n("showErrorMessage") ?? "0"),
      });
  }
  return r;
}
function Ple(e, r) {
  return e.replace(/<dataValidation\b[^>]*>/g, (t) => {
    var a, s;
    const n = (a = /\bsqref="([^"]+)"/.exec(t)) == null ? void 0 : a[1],
      o =
        (s = r.validations) == null
          ? void 0
          : s.find((d) => n && Jw(d.range) === Jw(n));
    return (o == null ? void 0 : o.kind) === "list" && o.showDropdown === false
      ? t
          .replace(/\s+showDropDown="[^"]*"/, "")
          .replace(/>$/, ' showDropDown="1">')
      : t;
  });
}
const Ule = {
  between: "between",
  notBetween: "notBetween",
  greater: "greaterThan",
  greaterEqual: "greaterThanOrEqual",
  less: "lessThan",
  lessEqual: "lessThanOrEqual",
  equal: "equal",
  notEqual: "notEqual",
};
function $le(e) {
  const r = {
    list: "list",
    number: "decimal",
    integer: "whole",
    textLength: "textLength",
    date: "date",
    time: "time",
    custom: "custom",
  }[e.kind];
  if (
    e.kind === "list" &&
    (e.list.some((t) => /[",\r\n]/.test(t)) || e.list.join(",").length > 253)
  )
    throw Error(
      "This list validation cannot be represented by an inline XLSX list. Retain the native workbook or use shorter values without commas, quotes or line breaks.",
    );
  if (e.kind === "custom" && B$(e.formula))
    throw Error(
      "An unsafe custom validation formula cannot be exported. Keep the native workbook and review the rule.",
    );
  return {
    type: r,
    ...(["list", "custom"].includes(e.kind)
      ? {}
      : { operator: Ule[e.operator] }),
    allowBlank: e.allowBlank,
    showErrorMessage: true,
    errorStyle: e.errorStyle ?? "stop",
    errorTitle: e.errorTitle ?? "Invalid value",
    error: e.message,
    showInputMessage: e.showInputMessage !== false,
    promptTitle: e.inputTitle ?? "",
    prompt: e.inputMessage ?? "",
    formulae:
      e.kind === "list"
        ? [JSON.stringify(e.list.join(","))]
        : e.kind === "custom"
          ? [e.formula.slice(1)]
          : [
              e.min,
              ...(["between", "notBetween"].includes(e.operator)
                ? [e.max]
                : []),
            ].map((t) =>
              e.kind === "date"
                ? new Date(Date.UTC(1899, 11, 30) + t * 864e5)
                : t,
            ),
  };
}
const B$ = (e) => {
    const r = Vg(e).replace(/"(?:[^"]|"")*"/g, '""');
    return (
      /[\[\]|]/.test(r) ||
      /(?:https?:|file:|javascript:|data:|\\\\)/i.test(r) ||
      /\b(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY)\s*\(/i.test(
        r.replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet"),
      )
    );
  },
  zle = (e) => '"' + e.replace(/"/g, '""') + '"';
function Vle(e, r) {
  return "ISNUMBER(FIND(LOWER(" + zle(e) + "),LOWER(" + r + '&"")))';
}
function Wle(e, r) {
  const t = ct(r),
    n = n3(t.firstRow, t.firstCol),
    o = Jw(r).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2"),
    a =
      "SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" +
      n +
      ',"~","~~"),"*","~*"),"?","~?")';
  return (
    "AND(LEN(" +
    n +
    ")>0,COUNTIF(" +
    o +
    ',"="&' +
    a +
    ")" +
    (e === "duplicate" ? ">1" : "=1") +
    ")"
  );
}
function Hle(e) {
  if (e === null || e === "") return true;
  if (typeof e == "number") return Number.isFinite(e);
  if (typeof e != "string") return false;
  const r = e.trim() !== "" && Number.isFinite(Number(e));
  return !(
    /[^\x20-\x7e]/.test(e) ||
    /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|ERROR!|CYCLE!|SPILL!|CALC!|LIMIT!|BLOCKED!)/.test(
      e,
    ) ||
    e.replace(/[~*?]/g, "~~").length > 254 ||
    (r && String(Number(e)) !== e) ||
    (!r && /^[\s+-]*\d/.test(e))
  );
}
function qle(e) {
  var n, o, a;
  const r = {},
    t = new Set();
  for (const s of e.matchAll(/<(?:\w+:)?cfRule\b([^>]*)>/g)) {
    const d =
        (n = /\bpriority=["']([^"']*)["']/.exec(s[1])) == null ? void 0 : n[1],
      u =
        (o = /\bstopIfTrue=["']([^"']+)["']/.exec(s[1])) == null
          ? void 0
          : o[1];
    if (d === void 0 && /\bid=["']/.test(s[1])) continue;
    const l = Number(d);
    if (!d || !/^\d+$/.test(d) || !Number.isSafeInteger(l) || l < 1 || t.has(l))
      throw Error("Invalid or duplicate conditional-format priority.");
    if ((t.add(l), d && u !== void 0)) {
      if (!["1", "0", "true", "false"].includes(u))
        throw Error("Invalid conditional-format stop flag.");
      r[String(l)] = u === "1" || u === "true";
    }
  }
  for (const s of e.matchAll(/<(?:\w+:)?cfvo\b([^>]*)>/g)) {
    const d = (a = /\bgte=["']([^"']*)["']/.exec(s[1])) == null ? void 0 : a[1];
    if (d !== void 0 && d !== "1" && d !== "true")
      throw Error(
        "Strict or malformed conditional icon thresholds are not supported. Use inclusive thresholds or native JSON.",
      );
  }
  return r;
}
function Xle(e, r) {
  const t = J3(r.conditionalFormats ?? []);
  return e.replace(/<cfRule\b[^>]*>/g, (n) => {
    var s;
    const o = Number(
        (s = /\bpriority="(\d+)"/.exec(n)) == null ? void 0 : s[1],
      ),
      a = t[o - 1];
    return a != null && a.stopIfTrue
      ? n.replace(/\s+stopIfTrue="[^"]*"/, "").replace(/>$/, ' stopIfTrue="1">')
      : n;
  });
}
function Gle(e, r) {
  const t = {
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: Jb(e.fill) },
      },
      font: { color: { argb: Jb(e.color) } },
    },
    n = ct(e.range),
    o = n3(n.firstRow, n.firstCol),
    a = (u) => ({
      type: "expression",
      priority: r,
      formulae: [u.replace(/^=/, "")],
      style: t,
    });
  if (e.kind === "formula") {
    if (B$(e.formula ?? ""))
      throw Error(
        "External-action conditional formulas cannot be exported. Remove that rule or use native JSON.",
      );
    return a(e.formula);
  }
  if (["blank", "notBlank", "error", "notError"].includes(e.kind))
    return a(
      e.kind === "blank"
        ? `LEN(${o})=0`
        : e.kind === "notBlank"
          ? `LEN(${o})>0`
          : e.kind === "error"
            ? `ISERROR(${o})`
            : `NOT(ISERROR(${o}))`,
    );
  if (e.kind === "unique" || e.kind === "duplicate")
    return a(Wle(e.kind, e.range));
  if (e.kind === "dataBar")
    return {
      type: "dataBar",
      priority: r,
      cfvo: [{ type: "min" }, { type: "max" }],
      color: { argb: Jb(e.fill) },
      gradient: false,
      showValue: true,
      negativeBarColorSameAsPositive: false,
      negativeFillColor: { argb: "FFD13438" },
      axisColor: { argb: "FF777777" },
      axisPosition: "auto",
      direction: "leftToRight",
      minLength: 0,
      maxLength: 100,
    };
  if (e.kind === "iconSet")
    return {
      type: "iconSet",
      priority: r,
      iconSet: "3Arrows",
      showValue: true,
      reverse: false,
      cfvo: [
        { type: "min" },
        ...(e.iconThresholds ?? [33, 67]).map((u) => ({
          type: e.thresholdType === "number" ? "num" : "percent",
          value: u,
        })),
      ],
    };
  if (e.kind === "contains") return a(Vle(e.value, o));
  if (e.kind === "colorScale")
    return {
      type: "colorScale",
      priority: r,
      cfvo: [{ type: "min" }, { type: "max" }],
      color: [
        { argb: Jb(e.minColor ?? "#f8e99b") },
        { argb: Jb(e.maxColor ?? "#63be7b") },
      ],
    };
  const s =
      e.kind === "greater"
        ? "greaterThan"
        : e.kind === "less"
          ? "lessThan"
          : e.kind === "equal"
            ? "equal"
            : "between",
    d = (u) =>
      u.trim() !== "" && Number.isFinite(Number(u))
        ? Number(u)
        : `"${u.replace(/"/g, '""')}"`;
  return {
    type: "cellIs",
    operator: s,
    formulae: [d(e.value), ...(e.kind === "between" ? [d(e.second)] : [])],
    priority: r,
    style: t,
  };
}
function Zle(e, r) {
  let t = 0,
    n = 0;
  for (const d of r.merges ?? []) {
    if ((t += b7(d)) > Mt.maxCells)
      throw Error("Merged ranges exceed the 100,000-cell interchange budget.");
    const u = ct(d);
    for (const [l, c] of Object.entries(r.cells)) {
      const h = _t(l);
      if (
        h.row >= u.firstRow &&
        h.row <= u.lastRow &&
        h.col >= u.firstCol &&
        h.col <= u.lastCol &&
        (h.row !== u.firstRow || h.col !== u.firstCol) &&
        ((c.raw !== null && c.raw !== "") ||
          c.note ||
          c.tickmark ||
          c.hyperlink)
      )
        throw Error(
          "Exporting this merged area would discard populated cells or annotations. Unmerge it or preserve those values before XLSX export.",
        );
    }
    e.mergeCells(d);
  }
  const o = [];
  for (const d of r.validations ?? []) {
    const u = ct(d.range);
    if (
      o.some(
        (l) =>
          l.firstRow <= u.lastRow &&
          l.lastRow >= u.firstRow &&
          l.firstCol <= u.lastCol &&
          l.lastCol >= u.firstCol,
      )
    )
      throw Error(
        "Overlapping validation rules cannot be represented by one Excel rule per cell. Use native JSON or resolve overlapping rules before export.",
      );
    if ((o.push(u), (n += b7(d.range)) > Mt.maxCells))
      throw Error(
        "XLSX validation ranges exceed the 100,000-cell interchange budget. Use native JSON to preserve larger rules.",
      );
    e.dataValidations.add(d.range, $le(d));
  }
  const a = new Map(
    J3(r.conditionalFormats ?? []).map((d, u) => [d.id, u + 1]),
  );
  let s = 0;
  for (const d of r.conditionalFormats ?? []) {
    if (d.kind !== "duplicate" && d.kind !== "unique") continue;
    const u = ct(d.range);
    for (const [l, c] of Object.entries(r.cells)) {
      if (++s > 1e6)
        throw Error(
          "Duplicate/unique conditional export exceeds 1,000,000 population checks. Use native JSON or fewer rules.",
        );
      const h = _t(l);
      if (
        h.row < u.firstRow ||
        h.row > u.lastRow ||
        h.col < u.firstCol ||
        h.col > u.lastCol
      )
        continue;
      let f = c.raw;
      if (
        (typeof c.raw == "string" &&
          c.raw.startsWith("=") &&
          !c.literal &&
          (f = e.getCell(l).result),
        !Hle(f))
      )
        throw Error(
          `Duplicate/unique conditional rule at ${l} cannot preserve Excel comparison semantics for this value. Use native JSON; XLSX supports finite numbers, canonical numeric text and short printable ASCII text without date-like numeric prefixes.`,
        );
    }
  }
  (r.conditionalFormats ?? []).forEach((d) =>
    e.addConditionalFormatting({ ref: d.range, rules: [Gle(d, a.get(d.id))] }),
  );
}
function Kle(e) {
  let r = {
    mode: "auto",
    iterative: false,
    maxIterations: null,
    maxChange: null,
  };
  const t = new Ox.SaxesParser({ xmlns: true });
  let n = false;
  return (
    t.on("opentag", (o) => {
      if (o.local !== "calcPr" || !/spreadsheetml/.test(o.uri)) return;
      if (n)
        throw Error(
          "Duplicate XLSX calculation properties. Repair the source workbook.",
        );
      n = true;
      const a = (h) => {
          const f = o.attributes[h];
          return f && typeof f != "string" ? f.value : f;
        },
        s = a("calcMode") ?? "auto",
        d = a("iterate"),
        u = a("iterateCount"),
        l = a("iterateDelta"),
        c = a("fullPrecision");
      if (
        !["auto", "autoNoTable", "manual"].includes(s) ||
        (c !== void 0 && !["true", "false", "1", "0"].includes(c)) ||
        (d !== void 0 && !["true", "false", "1", "0"].includes(d)) ||
        (u !== void 0 &&
          (!/^\d+$/.test(u) ||
            !Number.isSafeInteger(Number(u)) ||
            Number(u) < 1)) ||
        (l !== void 0 &&
          (!l.trim() || !Number.isFinite(Number(l)) || Number(l) < 0))
      )
        throw Error(
          "Invalid XLSX calculation properties. Repair the source workbook.",
        );
      r = {
        mode: s,
        iterative: d === "true" || d === "1",
        maxIterations: u === void 0 ? null : Number(u),
        maxChange: l === void 0 ? null : Number(l),
        ...(c === void 0 ? {} : { fullPrecision: c === "true" || c === "1" }),
      };
    }),
    t.on("doctype", () => {
      throw Error("XLSX calculation properties cannot contain a DTD.");
    }),
    t.write(e).close(),
    r
  );
}
function Yle(e, r) {
  const t = Nx(r.calculation ?? EC),
    n = `<calcPr fullPrecision="1" calcId="0" calcMode="${t.mode === "manual" ? "manual" : "auto"}" fullCalcOnLoad="1" forceFullCalc="1" calcOnSave="${t.mode === "manual" ? "0" : "1"}" iterate="${t.iterative ? "1" : "0"}" iterateCount="${t.maxIterations}" iterateDelta="${t.maxChange}"/>`;
  return /<(?:\w+:)?calcPr\b[^>]*\/>/.test(e)
    ? e.replace(/<(?:\w+:)?calcPr\b[^>]*\/>/, n)
    : /<(?:\w+:)?calcPr\b[^>]*>[\s\S]*?<\/(?:\w+:)?calcPr>/.test(e)
      ? e.replace(/<(?:\w+:)?calcPr\b[^>]*>[\s\S]*?<\/(?:\w+:)?calcPr>/, n)
      : e.replace(/<\/(?:\w+:)?workbook>/, `${n}$&`);
}
function jC(e) {
  const r = I$.get(e);
  if (!r || r.signature !== JSON.stringify(e))
    throw Error(
      "Automation inspection was modified or is untrusted. Inspect the original bytes again.",
    );
  return r.bytes.slice();
}
function L$(e, r) {
  const t = jC(e);
  if (!e.retainedEligible)
    throw Error(
      "This source exceeds the 2 MB artifact-retention limit. Keep the original file separately.",
    );
  if (!r.trim() || r.length > 100)
    throw Error("A bounded reviewer name is required.");
  return {
    id: crypto.randomUUID(),
    fileName: e.fileName,
    kind: e.kind,
    sha256: e.sha256,
    byteLength: t.length,
    base64: Zoe(t),
    retainedAt: new Date().toISOString(),
    retainedBy: r,
  };
}
function uT(e) {
  jC(e);
  const r = e.items.filter((t) =>
    /embedded|activeX|OLE|External automation/i.test(t.kind),
  );
  if (!r.length)
    throw Error(
      "No known OLE, ActiveX or embedded-object references were detected. Separately stored objects remain unknown.",
    );
  return {
    id: crypto.randomUUID(),
    fileName: e.fileName,
    fileBytes: e.fileBytes,
    sourceSha256: e.sha256,
    inspectedAt: new Date().toISOString(),
    execution: "never",
    items: r.map((t) => ({ ...t, execution: "never" })),
    warnings: [
      ...e.warnings,
      "Inventory only: no payload is opened, fetched or executed. Missing parts and external references remain unresolved. Keep the original source separately.",
    ],
  };
}
function Jle(e, r, t = {}) {
  var o;
  if (t.readOnly) throw Error("This workbook is read-only.");
  $a(e);
  const n = uT(r);
  if (
    (o = e.embeddedObjectInventories) != null &&
    o.some((a) => a.sourceSha256 === n.sourceSha256)
  )
    throw Error("This source inventory has already been retained.");
  return rn({
    ...e,
    embeddedObjectInventories: [...(e.embeddedObjectInventories ?? []), n],
  });
}
function Qle(e) {
  return (e.embeddedObjectInventories ?? []).flatMap((r) =>
    r.items.map((t) => ({
      sourceFile: r.fileName,
      sourceSha256: r.sourceSha256,
      path: t.path,
      kind: t.kind,
      bytes: t.bytes,
      ...(t.sha256 ? { sha256: t.sha256 } : {}),
      disposition: "dropped",
      execution: "never",
      reason:
        "Inventory retained in native JSON only; unsupported OLE/ActiveX payload was not imported and is not emitted in derived XLSX.",
    })),
  );
}
function P$(e, r, t, n, o) {
  var a;
  return {
    schemaVersion: 1,
    direction: e,
    fileName: r,
    fileBytes: t,
    workbookName: n.name,
    revision: n.revision,
    compatibility: "derived-subset",
    lossless: false,
    originalFileRetained: !!n.xlsxSource,
    ...(n.alternateSource
      ? {
          alternateOriginal: {
            fileName: n.alternateSource.fileName,
            byteLength: n.alternateSource.byteLength,
            sha256: n.alternateSource.sha256,
            format: n.alternateSource.format,
          },
        }
      : {}),
    ...(n.xlsxSource
      ? {
          originalSource: {
            fileName: n.xlsxSource.fileName,
            byteLength: n.xlsxSource.byteLength,
            sha256: n.xlsxSource.sha256,
          },
        }
      : {}),
    sheets: n.sheets.map((s) => ({
      name: s.name,
      storedCells: Object.keys(s.cells).length,
      hidden: !!(s.hidden || s.veryHidden),
    })),
    warnings: [...new Set(o)],
    ...(e === "export" && (a = n.embeddedObjectInventories) != null && a.length
      ? { embeddedObjectLosses: Qle(n) }
      : {}),
    caches: {
      total: 0,
      pending: 0,
      matched: 0,
      different: 0,
      missing: 0,
      neutralized: 0,
      written: 0,
      omitted: 0,
      issues: [],
      matches: [],
      detailsOmitted: 0,
    },
  };
}
const eue = new Set([
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
]);
function tue(e) {
  if (e.type !== "blank")
    return e.type === "error"
      ? eue.has(String(e.value))
        ? { error: e.value }
        : void 0
      : e.value;
}
function rue(e, r) {
  const t = (o, a) => {
    var d;
    const s =
      (d = new RegExp(`(?:^|\\s)${a}\\s*=\\s*(["'])(.*?)\\1`, "i").exec(o)) ==
      null
        ? void 0
        : d[2];
    return s == null
      ? void 0
      : s.replace(/&#(x[0-9a-f]+|[0-9]+);/gi, (u, l) =>
          String.fromCodePoint(
            l[0].toLowerCase() === "x" ? parseInt(l.slice(1), 16) : Number(l),
          ),
        );
  };
  let n;
  for (const o of e.matchAll(
    /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<(\/?)(?:\w+:)?(c|f)\b([^>]*)>/gi,
  )) {
    if (!o[2]) continue;
    if (o[2].toLowerCase() === "c") {
      n = o[1] || /\/\s*$/.test(o[3]) ? void 0 : o[3];
      continue;
    }
    if (n === void 0 || o[1] || t(o[3], "t") !== "array") continue;
    const a = t(n, "r") ?? "unknown address",
      s = t(o[3], "ref") ?? a,
      d =
        t(n, "cm") !== void 0
          ? "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)"
          : "Legacy CSE array formula";
    throw new Error(
      `${d} at ${r}!${a}, range ${s}. Import stopped before changing the workbook. Open the original in Excel and convert the array to ordinary supported formulas or an explicitly reviewed values snapshot, then import that copy. Fixed CSE ranges and Excel array metadata are not silently treated as dynamic spills.`,
    );
  }
}

const oue = ax(iue);

const x7 = ax(lue),
  U$ = {
    delimiter: "auto",
    encoding: "auto",
    quote: '"',
    escape: "double",
    newline: "auto",
    hasHeader: true,
    decimal: ".",
    grouping: "none",
    dateOrder: "ymd",
    empty: "blank",
    trim: false,
    columns: {},
  },
  uue = ["auto", "text", "number", "integer", "boolean", "date"],
  $$ = Array.from({ length: 256 }, (e, r) => String.fromCharCode(r));
[
  8364, 129, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 352, 8249, 338, 141,
  381, 143, 144, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732, 8482, 353, 8250,
  339, 157, 382, 376,
].forEach((e, r) => ($$[128 + r] = String.fromCharCode(e)));
const E7 = (e, r, t, n, o) => {
    (e.issueCount++,
      e.issues.length < 100 &&
        e.issues.push({
          record: r,
          column: t,
          message: n,
          ...(o !== void 0 ? { value: o.slice(0, 200) } : {}),
        }));
  },
  C7 = (e, r, t, n, o, a, s) => {
    (e.changeCount++,
      e.changes.push({
        record: r,
        column: t,
        address: n,
        reason: o,
        before: a,
        after: s,
      }));
  };
function cue(e = {}) {
  const r = { ...U$, ...e, columns: { ...(e.columns ?? {}) } };
  for (const [t, n] of Object.entries({
    delimiter: ["auto", ",", ";", "	", "|"],
    encoding: ["auto", "utf-8", "utf-16le", "utf-16be", "windows-1252"],
    quote: ['"', "'", "none"],
    escape: ["double", "backslash"],
    newline: [
      "auto",
      `\r
`,
      `
`,
      "\r",
    ],
    decimal: [".", ","],
    grouping: ["none", ".", ",", " "],
    dateOrder: ["ymd", "dmy", "mdy"],
    empty: ["blank", "text"],
  }))
    if (!n.includes(String(r[t])))
      throw Error(`Unsupported delimited ${t} setting.`);
  if (typeof r.hasHeader != "boolean" || typeof r.trim != "boolean")
    throw Error("Header and whitespace controls must be boolean.");
  if (r.decimal === r.grouping)
    throw Error("Decimal and grouping separators must differ.");
  for (const [t, n] of Object.entries(r.columns))
    if (
      !/^(0|[1-9]\d*)$/.test(t) ||
      Number(t) >= Mt.maxCols ||
      !uue.includes(n)
    )
      throw Error("Invalid column type selection.");
  return r;
}
function due(e, r, t, n) {
  return {
    schemaVersion: 1,
    direction: e,
    fileName: r,
    fileBytes: t,
    options: n,
    records: 0,
    columns: 0,
    storedCells: 0,
    blankRecords: 0,
    schema: [],
    preview: [],
    issueCount: 0,
    issues: [],
    changeCount: 0,
    changes: [],
    safeTextChanges: 0,
    complete: true,
    valid: false,
    warnings: [],
    safeExportPolicy: "apostrophe-prefix-formula-like-text",
  };
}
const Ak = (e) => e.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
function A7(e, r) {
  const t = r.grouping === "none" ? "" : r.grouping,
    n = t ? `(?:\\d+|\\d{1,3}(?:${Ak(t)}\\d{3})+)` : "\\d+";
  if (
    !new RegExp(
      `^[+-]?(?:${n}(?:${Ak(r.decimal)}\\d*)?|${Ak(r.decimal)}\\d+)(?:[eE][+-]?\\d+)?$`,
    ).test(e)
  )
    throw Error("Invalid number for the selected decimal/grouping separators.");
  const a = (t ? e.split(t).join("") : e).replace(r.decimal, "."),
    s = a
      .split(/[eE]/)[0]
      .replace(/[^0-9]/g, "")
      .replace(/^0+/, "")
      .replace(/0+$/, "").length,
    d = Number(a);
  if (
    !Number.isFinite(d) ||
    s > 15 ||
    (d === 0 && /[1-9]/.test(a.split(/[eE]/)[0]))
  )
    throw Error(
      "Number is outside supported precision/range; select Text to retain every digit.",
    );
  return d;
}
function fue(e, r, t = "1900") {
  const n =
    /^(\d{1,4})([-/])(\d{1,2})\2(\d{1,4})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(
      e,
    );
  if (!n)
    throw Error(
      "Date must match the selected year/month/day order, optionally HH:mm:ss.",
    );
  const o = [n[1], n[3], n[4]],
    [a, s, d] =
      r === "ymd" ? o : r === "dmy" ? [o[2], o[1], o[0]] : [o[2], o[0], o[1]];
  if (a.length !== 4) throw Error("Dates require a four-digit year.");
  const u = Number(n[5] ?? 0),
    l = Number(n[6] ?? 0),
    c = Number(n[7] ?? 0);
  if (u > 23 || l > 59 || c > 59) throw Error("Invalid date/time component.");
  return (
    T1(`${a}-${s.padStart(2, "0")}-${d.padStart(2, "0")}`, t) +
    (u * 3600 + l * 60 + c) / 86400
  );
}
function hue(e, r, t = "1900") {
  if (!Number.isFinite(e) || e < 0 || e >= wd(t) + 1)
    throw Error(`Date serial is outside 0–${wd(t)} for the ${t} system.`);
  let n = Math.floor(e),
    o = Math.round((e - n) * 86400);
  if ((o === 86400 && (n++, (o = 0)), n > wd(t)))
    throw Error("Rounded date/time exceeds year 9999.");
  const a =
      t === "1900" && n === 60
        ? "1900-02-29"
        : new Date(
            t === "1904"
              ? Date.UTC(1904, 0, 1) + n * 864e5
              : Date.UTC(1899, 11, 31) + (n > 60 ? n - 1 : n) * 864e5,
          )
            .toISOString()
            .slice(0, 10),
    [s, d, u] = a.split("-");
  return (
    (r === "ymd" ? a : r === "dmy" ? `${u}/${d}/${s}` : `${d}/${u}/${s}`) +
    (o
      ? ` ${String(Math.floor(o / 3600)).padStart(2, "0")}:${String(Math.floor(o / 60) % 60).padStart(2, "0")}:${String(o % 60).padStart(2, "0")}`
      : "")
  );
}
function pue(e, r, t, n = "1900") {
  if (e === "") return { raw: t.empty === "text" ? "" : null };
  const o = t.trim ? e.trim() : e;
  if (r === "auto") {
    if (
      /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(o) &&
      o.split(/[eE]/)[0].replace(/[^0-9]/g, "").length <= 15 &&
      t.decimal === "." &&
      t.grouping === "none"
    )
      try {
        return { raw: A7(o, t) };
      } catch {}
    return { raw: o, ...(o.startsWith("=") ? { literal: true } : {}) };
  }
  if (r === "text")
    return { raw: o, ...(o.startsWith("=") ? { literal: true } : {}) };
  if (r === "date")
    return {
      raw: fue(o, t.dateOrder, n),
      format: "date",
      numFmt: /[ T]\d\d:/.test(o) ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd",
    };
  if (r === "boolean") {
    if (/^(true|false)$/i.test(o)) return { raw: o.toLowerCase() === "true" };
    throw Error("Boolean values must be TRUE or FALSE (case-insensitive).");
  }
  const a = A7(o, t);
  if (r === "integer" && !Number.isSafeInteger(a))
    throw Error("Integer column requires a safe whole number.");
  return { raw: a };
}
function z$(e, r) {
  var o, a;
  const t = e.sheets.find((s) => s.id === r);
  if (!t) throw Error("Worksheet not found.");
  const n = RC(t, false);
  return `A1:${ut(Math.max((n == null ? void 0 : n.lastCol) ?? 0, (((o = t.delimitedSource) == null ? void 0 : o.columns) ?? 1) - 1))}${Math.max((n == null ? void 0 : n.lastRow) ?? 0, (((a = t.delimitedSource) == null ? void 0 : a.records) ?? 1) - 1) + 1}`;
}
function mue(e, r, t) {
  if (
    !["utf-8", "utf-16le", "utf-16be", "windows-1252"].includes(r) ||
    typeof t != "boolean"
  )
    throw Error("Invalid output encoding or BOM setting.");
  for (const s of e)
    if (s.length === 1 && s.charCodeAt(0) >= 55296 && s.charCodeAt(0) <= 57343)
      throw Error(
        "Unpaired Unicode surrogate cannot be exported; correct the text first.",
      );
  if (r === "utf-8") return new TextEncoder().encode((t ? "\uFEFF" : "") + e);
  if (r === "windows-1252") {
    if (t)
      throw Error(
        "Windows-1252 does not support a BOM; disable BOM or choose UTF-8.",
      );
    const s = new Map($$.map((u, l) => [u, l])),
      d = [];
    for (const u of e) {
      const l = s.get(u);
      if (l === void 0)
        throw Error(
          "A character cannot be represented in Windows-1252. Choose UTF-8; text is never replaced.",
        );
      d.push(l);
    }
    return Uint8Array.from(d);
  }
  const n = (t ? "\uFEFF" : "") + e,
    o = new Uint8Array(n.length * 2),
    a = new DataView(o.buffer);
  for (let s = 0; s < n.length; s++)
    a.setUint16(s * 2, n.charCodeAt(s), r === "utf-16le");
  return o;
}
const gue = (e) =>
  /^[\s\u0000-\u001f]*[=+\-@＝＋－＠]/u.test(e) || /^[\t\r\n]/.test(e);
function vue(e, r) {
  const t =
    e.includes(r.delimiter) ||
    /[\r\n\uFEFF]/.test(e) ||
    (r.quote !== "none" && e.includes(r.quote)) ||
    /^[ ]|[ ]$/.test(e) ||
    r.quoteAll;
  if (r.quote === "none") {
    if (e.includes(r.delimiter) || /[\r\n]/.test(e))
      throw Error(
        "A field needs quoting; choose a quote character to preserve delimiters/newlines.",
      );
    return e;
  }
  if (!t) return e;
  const n =
    r.escape === "double"
      ? e.split(r.quote).join(r.quote + r.quote)
      : e
          .replaceAll("\\", "\\\\")
          .split(r.quote)
          .join("\\" + r.quote);
  return r.quote + n + r.quote;
}
function V$(e, r, t = {}) {
  var g, b, y;
  const n = rn(e),
    o = n.sheets.find((w) => w.id === r);
  if (!o) throw Error("Worksheet not found.");
  const a = {
    ...cue({
      ...t,
      delimiter: t.delimiter ?? ",",
      encoding: t.encoding ?? "utf-8",
      newline:
        t.newline ??
        `\r
`,
    }),
    includeHeader: t.includeHeader ?? true,
    bom: t.bom ?? false,
    quoteAll: t.quoteAll ?? false,
    dateOutput: t.dateOutput ?? "calendar",
    range:
      t.range ??
      ((g = t.population) == null ? void 0 : g.definition.range) ??
      z$(n, r),
    ...(t.population ? { population: t.population } : {}),
  };
  for (const w of ["includeHeader", "bom", "quoteAll"])
    if (typeof a[w] != "boolean") throw Error(`Invalid export ${w} control.`);
  if (!["calendar", "serial"].includes(a.dateOutput))
    throw Error("Invalid date export mode.");
  (a.delimiter === "auto" && (a.delimiter = ","),
    a.encoding === "auto" && (a.encoding = "utf-8"),
    a.newline === "auto" &&
      (a.newline = `\r
`));
  const s = ct(a.range),
    d = s.lastRow - s.firstRow + 1,
    u = s.lastCol - s.firstCol + 1;
  if (d * u > Mt.maxCells)
    throw Error(
      "Delimited/CSV export would expand more than 100,000 cells including blank gaps. Choose a smaller range or use native JSON.",
    );
  const l = due(
      "export",
      `${o.name}.${a.delimiter === "	" ? "tsv" : "csv"}`,
      0,
      a,
    ),
    c = Vn(n),
    h = [];
  if (a.population) {
    if (!["full", "visible"].includes(a.population.scope) || !a.hasHeader)
      throw Error(
        "Filtered exports require a header and an explicit full or visible scope.",
      );
    const w = f0(n, r, a.population.definition, c).context;
    if (a.range !== w.range && a.range !== a.population.definition.range)
      throw Error(
        "The export range must match the declared filter population.",
      );
    ((l.population = {
      ...w,
      scope: a.population.scope,
      exportedSourceRows:
        a.population.scope === "visible" ? w.visibleSourceRows : w.sourceRows,
    }),
      l.warnings.push(
        `Population scope: ${a.population.scope}; exporting ${l.population.exportedSourceRows.length} of ${w.totalRows} source data rows. Save the schema report for filter criteria, row locators and source IDs. Session filter criteria are not encoded in CSV/TSV.`,
      ));
  } else
    l.warnings.push(
      "Population scope: full selected rectangle, including hidden and filtered rows.",
    );
  const f = l.population ? new Set(l.population.exportedSourceRows) : null;
  ((l.columns = u),
    l.warnings.push(
      "Values-only export: formulas are evaluated; formatting, notes, history and workbook metadata are not carried by CSV/TSV. Reimport with the visible schema; the file itself has no type schema.",
      "Safe export always prefixes an apostrophe to formula-like text, including whitespace/control-prefixed and full-width triggers. Numeric negatives stay numeric. Quotes alone do not neutralize spreadsheet formulas. Changed text is listed below; it is not automatically unescaped on reimport.",
    ),
    o.delimitedSource &&
      l.warnings.push(
        "The default range includes the historical imported footprint so trailing blank records survive. Review the range after structural edits; this is not a live table binding.",
      ));
  let p = 0;
  for (let w = s.firstRow; w <= s.lastRow; w++) {
    if (w !== s.firstRow && f && !f.has(w + 1)) continue;
    const E = w - s.firstRow + 1,
      A = a.hasHeader && w === s.firstRow,
      k = [],
      F = [],
      I = [];
    let O = true;
    for (let V = 0; V < u; V++) {
      const D = `${ut(s.firstCol + V)}${w + 1}`,
        B = o.cells[D],
        _ = c(r, D),
        C = a.columns[V] ?? "auto";
      if (
        (w === s.firstRow &&
          l.schema.push({
            index: V,
            label: a.hasHeader ? String(_ ?? "") : ut(s.firstCol + V),
            type: C,
            observed: [],
          }),
        A && !a.includeHeader)
      )
        continue;
      let T = String(_ ?? ""),
        N = typeof _ == "string";
      try {
        const R =
          A || C === "auto"
            ? { raw: _ }
            : typeof _ == "number" && ["number", "integer", "date"].includes(C)
              ? { raw: _ }
              : pue(
                  T,
                  C,
                  a,
                  (b = n.datePolicy) == null ? void 0 : b.dateSystem,
                );
        if (
          !A &&
          C === "integer" &&
          typeof R.raw == "number" &&
          !Number.isSafeInteger(R.raw)
        )
          throw Error("Integer column requires a safe whole number.");
        if (
          ((N = typeof R.raw == "string"),
          (T = String(R.raw ?? "")),
          typeof R.raw == "number" &&
            (!A &&
            (C === "date" ||
              (C === "auto" && (B == null ? void 0 : B.format) === "date")) &&
            a.dateOutput === "calendar"
              ? (T = hue(
                  R.raw,
                  a.dateOrder,
                  (y = n.datePolicy) == null ? void 0 : y.dateSystem,
                ))
              : (T = T.replace(".", a.decimal))),
          N && gue(T))
        ) {
          const $ = T;
          ((T = "'" + T),
            l.safeTextChanges++,
            C7(
              l,
              E,
              V + 1,
              D,
              "Formula-like text prefixed with apostrophe",
              $,
              T,
            ));
        } else
          T !== String(_ ?? "") &&
            C7(
              l,
              E,
              V + 1,
              D,
              "Selected type, number or date representation",
              String(_ ?? ""),
              T,
            );
        const L = R.raw === null ? "blank" : typeof R.raw;
        if (
          (l.schema[V].observed.includes(L) || l.schema[V].observed.push(L),
          T.length > Ja.maxCellTextLength)
        )
          throw Error(
            "Safe-export prefix would exceed the 32,767-character field limit. Shorten the text before exporting.",
          );
        (T !== "" && (O = !1), k.push(vue(T, a)));
      } catch (R) {
        (E7(
          l,
          E,
          V + 1,
          R instanceof Error ? R.message : "Export conversion failed.",
          T,
        ),
          k.push(""));
      }
      (_ !== null && l.storedCells++,
        V < 200 && (F.push(String(_ ?? "")), I.push(T)));
    }
    if (A && !a.includeHeader) continue;
    const j = k.join(a.delimiter);
    O && l.blankRecords++;
    const z = j === "" && a.quote !== "none" ? a.quote + a.quote : j;
    if (
      ((p +=
        new TextEncoder().encode(z).length + (h.length ? a.newline.length : 0)),
      p > Mt.maxFileBytes * 3)
    )
      throw Error(
        "Delimited export exceeds the 10 MB output limit. Choose a smaller range.",
      );
    (h.push(z),
      l.records++,
      l.preview.length < 25 &&
        l.preview.push({ record: E, raw: F, values: I }));
  }
  let m = h.join(a.newline);
  h.at(-1) === "" && a.quote === "none" && (m += a.newline);
  let v = new Uint8Array();
  try {
    v = mue(m, a.encoding, a.bom);
  } catch (w) {
    E7(l, 0, void 0, w instanceof Error ? w.message : "Encoding failure.");
  }
  if (v.length > Mt.maxFileBytes)
    throw Error(
      "CSV/TSV export exceeds the 10 MB output limit. Choose a smaller range or native JSON.",
    );
  return (
    (l.fileBytes = v.length),
    (l.valid = l.issueCount === 0),
    (l.complete = true),
    { text: m, bytes: l.valid ? v : new Uint8Array(), report: l }
  );
}
const bue = Ja.maxXlsxExpandedBytes,
  yue =
    "XLSX exchange supports cell values, formulas, common formats, supported fonts/strike/alignment/indent/per-edge borders, notes, hidden sheets, column widths, explicit row heights, worksheet default dimensions, bounded hidden rows/columns, frozen rows and columns, basic tables, safe merges, supported validation/conditional-format rules, PNG/JPEG pictures, repeated print titles, full-axis manual print breaks and page setup. XLSX processes at most 100,000 hidden rows per workbook; native JSON keeps larger visibility intervals compactly. Workbook and worksheet-scoped defined names, constant/formula definitions and supported name comments are preserved. Unsupported name attributes are reported; unsafe or malformed definitions stop import/export. Enhanced number formats export supported masks; locale, scaling labels, duration/date-system and accounting alignment fidelity may differ, and enhanced format metadata remains native JSON only. Reusable calculated-column rules are native JSON only; XLSX imports retain supported cell formulas but do not infer rules for future appends. Editable outline groups, threaded discussions and interactive hyperlink metadata remain native-JSON features; XLSX retains only collapsed outline visibility. Native pivot/query refresh definitions, charts/shapes, filters, split panes, default-hidden-row settings, password protection and other unsupported Excel metadata are not preserved. External or executable formulas are converted to literal text on import and export. Keep the original workbook; use native JSON for a lossless application backup.",
  wue = (e) =>
    /[\[\]|]/.test(Vg(e).replace(/"(?:[^"]|"")*"/g, '""')) ||
    /(?:https?:|file:|javascript:|data:|\\\\)/i.test(e) ||
    /\b(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY)\s*\(/i.test(
      e,
    );
function il(e, r) {
  var t;
  return (t = new RegExp(
    "(?:^|\\s)" + r + `\\s*=\\s*(["'])([\\s\\S]*?)\\1`,
  ).exec(e)) == null
    ? void 0
    : t[2];
}
function xue(e) {
  const r = [];
  for (const t of (e.startsWith("/") ? e.slice(1) : "xl/" + e).split("/"))
    t === ".." ? r.pop() : t !== "." && r.push(t);
  return r.join("/");
}
const Ev = (e) =>
  e.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[a-f\d]+);/gi, (r) => {
    var t, n;
    return r === "&amp;"
      ? "&"
      : r === "&lt;"
        ? "<"
        : r === "&gt;"
          ? ">"
          : r === "&quot;"
            ? '"'
            : r === "&apos;"
              ? "'"
              : String.fromCodePoint(
                  Number.parseInt(
                    r.slice(
                      ((t = r[2]) == null ? void 0 : t.toLowerCase()) === "x"
                        ? 3
                        : 2,
                      -1,
                    ),
                    ((n = r[2]) == null ? void 0 : n.toLowerCase()) === "x"
                      ? 16
                      : 10,
                  ),
                );
  });
function W2(e, r) {
  const t = ble(e),
    n = new Set(t.keys()),
    o = t.size,
    a = Ele();
  if (
    [...n].some((w) =>
      /(?:vbaProject\.bin|xl\/macrosheets\/|xl\/activeX\/|xl\/embeddings\/)/i.test(
        w,
      ),
    )
  )
    throw new Error(
      "Workbooks containing macros, ActiveX or embedded objects are not accepted. Save a clean XLSX copy first.",
    );
  let s = 0,
    d,
    u = 0,
    l = 0,
    c = 0,
    h = 0,
    f = 0,
    p = 0,
    m = 0,
    v = false;
  const g = new Set(),
    b = new Unzip((w) => {
      if ((u++, !n.has(w.name) || g.has(w.name))) {
        d = new Error("Archive entries do not match its directory.");
        return;
      }
      g.add(w.name);
      let E = 0;
      const A = /\.(?:xml|rels|vml)$/i.test(w.name),
        k = A ? new TextDecoder("utf-8", { fatal: true }) : void 0;
      let F = -1,
        I = "";
      ((w.ondata = (O, j, z) => {
        var V;
        if (O) {
          d = new Error(
            `Corrupt XLSX compressed content in ${w.name}. Repair the original workbook and save a clean XLSX copy.`,
          );
          return;
        }
        if (((s += j.length), (E += j.length), s > bue)) {
          (w.terminate(),
            (d = new Error("The XLSX archive expands beyond 50 MB.")));
          return;
        }
        if (E > t.get(w.name).size) {
          (w.terminate(),
            (d = new Error(
              "The XLSX archive expansion does not match its directory.",
            )));
          return;
        }
        F = D$(F, j);
        try {
          k && (I += k.decode(j, { stream: !z }));
        } catch {
          d = new Error(
            `Unsupported XML encoding in ${w.name}. Save a standard UTF-8 XLSX copy.`,
          );
          return;
        }
        if (z) {
          if (E !== t.get(w.name).size) {
            d = new Error(
              "The XLSX archive expansion does not match its directory.",
            );
            return;
          }
          if ((F ^ -1) >>> 0 !== t.get(w.name).crc) {
            d = new Error(
              `XLSX checksum mismatch in ${w.name}. The archive is corrupt; repair the original and save a clean XLSX copy.`,
            );
            return;
          }
          try {
            if ((A && Cle(I, w.name, t, a), /^xl\/styles\.xml$/i.test(w.name)))
              for (const D of I.matchAll(/<(?:\w+:)?sz\b([^>]*)>/g)) {
                const B = Number(il(D[1], "val"));
                (!Number.isInteger(B) || B < 8 || B > 72) && (v = !0);
              }
            if (/^xl\/worksheets\//i.test(w.name)) {
              (rue(I, w.name),
                (m = mle(I, w.name, m)),
                /^xl\/worksheets\/[^/]+\.xml$/i.test(w.name) &&
                  r.numericTokens.set(w.name, Ble(I)));
              for (const C of I.matchAll(/<(?:\w+:)?pageSetup\b([^>]*)>/gi)) {
                const T = il(C[1], "pageOrder");
                if (
                  T !== void 0 &&
                  !["overThenDown", "downThenOver"].includes(Ev(T))
                )
                  throw new Error(
                    "Unsupported XLSX page order. Import stopped to preserve the source layout.",
                  );
              }
              const D = {};
              for (const [C, T, N, R] of [
                ["rowBreaks", "rowBreaks", Mt.maxRows, Mt.maxCols],
                ["colBreaks", "columnBreaks", Mt.maxCols, Mt.maxRows],
              ]) {
                const L = [
                  ...I.matchAll(
                    new RegExp(
                      "<(?:\\w+:)?" +
                        C +
                        "\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?" +
                        C +
                        ">",
                      "gi",
                    ),
                  ),
                ];
                if (L.length > 1)
                  throw Error(
                    "Duplicate XLSX page-break sections are not supported.",
                  );
                const $ = [];
                let U = 0;
                for (const H of (
                  ((V = L[0]) == null ? void 0 : V[1]) ?? ""
                ).matchAll(/<(?:\w+:)?brk\b([^>]*)\/?\s*>/gi)) {
                  if (++U > 100)
                    throw Error(
                      "XLSX page breaks exceed 100 boundaries per axis.",
                    );
                  const Q = il(H[1], "id"),
                    X = il(H[1], "man");
                  if (
                    !/^[1-9]\d*$/.test(Q ?? "") ||
                    !Number.isSafeInteger(Number(Q)) ||
                    Number(Q) >= N
                  )
                    throw Error("Invalid XLSX page-break boundary.");
                  if (X !== "1" && X !== "true") continue;
                  const J = Number(il(H[1], "min") ?? 0),
                    q = Number(il(H[1], "max") ?? R - 1);
                  if (J !== 0 || !Number.isSafeInteger(q) || q < R - 1)
                    throw Error(
                      "Partial-span XLSX page breaks are not supported. Use full-row or full-column breaks.",
                    );
                  $.push(Number(Q));
                }
                $.sort((H, Q) => H - Q);
                const G = GE($, N);
                G && (D[T] = G);
              }
              (r.printBreaks.set(w.name, D),
                r.conditionalStops.set(w.name, qle(I)),
                r.validationAreas.set(w.name, Lle(I)));
              const B = [],
                _ = [];
              for (const C of I.matchAll(/<(?:\w+:)?col\b([^>]*)>/gi)) {
                const T = il(C[1], "min"),
                  N = il(C[1], "max"),
                  R = il(C[1], "width"),
                  L = Number(T) - 1,
                  $ = Number(N) - 1;
                if (
                  !/^[1-9]\d*$/.test(T ?? "") ||
                  !/^[1-9]\d*$/.test(N ?? "") ||
                  !Number.isSafeInteger(L) ||
                  !Number.isSafeInteger($) ||
                  L < 0 ||
                  $ < L ||
                  $ >= Mt.maxCols ||
                  _.length >= Mt.maxCols
                )
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
                const U = R === void 0 ? void 0 : Number(R);
                if (
                  U !== void 0 &&
                  (!(R != null && R.trim()) || !Number.isFinite(U) || U < 0)
                )
                  throw Error(
                    "XLSX column widths must be finite nonnegative values.",
                  );
                _.push({
                  start: L,
                  end: $,
                  ...(U === void 0 ? {} : { width: U }),
                });
              }
              _.sort((C, T) => C.start - T.start);
              for (let C = 1; C < _.length; C++)
                if (_[C].start <= _[C - 1].end)
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
              if (
                (r.columns.set(w.name, _),
                /<(?:\w+:)?sheetFormatPr\b[^>]*\bzeroHeight\s*=\s*["'](?:1|true)["']/i.test(
                  I,
                ))
              )
                throw new Error(
                  "Default-hidden worksheet rows cannot be imported safely. Replace the default with explicit hidden rows before importing.",
                );
              for (const C of I.matchAll(/<(?:\w+:)?row\b([^>]*)>/gi)) {
                if (!/\bhidden\s*=\s*["'](?:1|true)["']/i.test(C[1])) continue;
                const T = Number(il(C[1], "r")) - 1;
                if (!Number.isInteger(T) || T < 0 || T >= Mt.maxRows)
                  throw new Error(
                    "Hidden XLSX rows must have valid row numbers.",
                  );
                const N = B.at(-1);
                if (N && T <= N.end)
                  throw new Error(
                    "Hidden XLSX row numbers must be distinct and ascending.",
                  );
                if (
                  (N && T === N.end + 1
                    ? (N.end = T)
                    : B.push({ start: T, end: T }),
                  B.length > x1.maxIntervalsPerAxis)
                )
                  throw new Error(
                    "Hidden metadata exceeds 10,000 intervals per worksheet axis.",
                  );
                if (++p > x1.maxXlsxHiddenRows)
                  throw new Error(
                    "XLSX hidden rows exceed the 100,000-row processing budget. Reduce hidden rows in the source; native JSON preserves larger visibility intervals.",
                  );
              }
              B.length && r.hiddenRows.set(w.name, B);
            }
            if (w.name === "xl/workbook.xml")
              for (const D of I.matchAll(/<(?:\w+:)?sheet\b([^>]*)>/g)) {
                const B = il(D[1], "name"),
                  _ = il(D[1], "r:id");
                B && _ && r.sheetNames.set(Ev(B), Ev(_));
              }
            if (w.name === "xl/_rels/workbook.xml.rels")
              for (const D of I.matchAll(/<(?:\w+:)?Relationship\b([^>]*)>/g)) {
                const B = il(D[1], "Id"),
                  _ = il(D[1], "Target");
                B &&
                  _ &&
                  il(D[1], "TargetMode") !== "External" &&
                  r.sheetRelationships.set(Ev(B), xue(Ev(_)));
              }
            for (const D of I.matchAll(
              /<(?:\w+:)?(mergeCell|dataValidation)\s[^>]*?\b(ref|sqref)\s*=\s*["']([^"']+)["'][^>]*>/g,
            ))
              for (const B of Ev(D[3]).trim().split(/\s+/)) {
                const _ = ct(B),
                  C =
                    (_.lastRow - _.firstRow + 1) * (_.lastCol - _.firstCol + 1);
                if (D[1] === "mergeCell") {
                  if (C > 500 || ++f > 1e3 || (c += C) > Mt.maxCells)
                    throw new Error(
                      "Merged ranges exceed the supported import budget (500 cells per merge and 100,000 cells overall).",
                    );
                } else if ((h += C) > Mt.maxCells)
                  throw new Error(
                    "XLSX validation ranges exceed the 100,000-cell import budget. Narrow the validation area in the source workbook.",
                  );
              }
            if (w.name === "xl/workbook.xml") {
              r.calculation = Kle(I);
              const D = aP(I);
              ((r.names = D.names),
                (r.nameWarnings = D.warnings),
                (r.builtinNameXml = D.builtinXml));
            }
          } catch (D) {
            d =
              D instanceof Error
                ? D
                : new Error("Invalid XLSX worksheet metadata.");
            return;
          }
          l++;
        }
      }),
        w.start());
    });
  b.register(UnzipInflate);
  for (let w = 0; w < e.length; w += 2048) {
    if (d) throw d;
    try {
      b.push(e.subarray(w, Math.min(w + 2048, e.length)), w + 2048 >= e.length);
    } catch {
      throw new Error(
        "Corrupt XLSX compressed content. Repair the original workbook and save a clean XLSX copy.",
      );
    }
  }
  if (d) throw d;
  if (u !== o || l !== o)
    throw new Error("The XLSX archive contains incomplete entries.");
  const y = [yue, ...r.nameWarnings, ...Ale(a, n)];
  return (
    (r.emptyStringCaches = a.emptyStringCaches),
    v &&
      y.push(
        "Unsupported fractional or out-of-range font sizes were found in the source. The XLSX reader may normalize them; the editor supports integer font sizes from 8 to 72 points.",
      ),
    [...n].some((w) => /externalLinks|connections\.xml/i.test(w)) &&
      y.push(
        "External workbook links or connections were detected. They will not be refreshed or included in the derived copy; the original retains their source parts.",
      ),
    y
  );
}
const kk = (e) => {
    const r = e.slice(1);
    return `FF${r.length === 3 ? [...r].map((t) => t + t).join("") : r}`;
  },
  Eue = {
    general: "General",
    number: "#,##0.00",
    currency: "$#,##0.00;[Red]($#,##0.00)",
    percent: "0.00%",
    date: "yyyy-mm-dd",
  };
function Cue(e, r) {
  if (
    !r.some((n) => {
      var o, a;
      return (
        ((o = n.hiddenRows) == null ? void 0 : o.length) ||
        ((a = n.hiddenCols) == null ? void 0 : a.length)
      );
    })
  )
    return e;
  const t = unzipSync(e);
  return (
    r.forEach((n, o) => {
      var h, f;
      const a = E1(n, "row");
      if (!a.length && !((h = n.hiddenCols) != null && h.length)) return;
      const s = "xl/worksheets/sheet" + (o + 1) + ".xml";
      let d = strFromU8(t[s]);
      if (
        ((d = d.replace(/<col\b([^>]*)\/>/g, (p, m) => {
          var w, E;
          const v = Number(il(m, "min")) - 1,
            g = Number(il(m, "max")) - 1;
          if (
            !Number.isInteger(v) ||
            !Number.isInteger(g) ||
            v < 0 ||
            g < v ||
            g >= Mt.maxCols
          )
            throw Error("The generated worksheet has invalid column metadata.");
          const b = m
              .replace(
                /\s+(?:min|max|width|customWidth)\s*=\s*(["'])[\s\S]*?\1/g,
                "",
              )
              .replace(/\/\s*$/, ""),
            y = [];
          for (let A = v; A <= g;) {
            const k = (w = n.columnWidths) == null ? void 0 : w[A];
            let F = A;
            for (
              ;
              F < g && ((E = n.columnWidths) == null ? void 0 : E[F + 1]) === k;
            )
              F++;
            (y.push(
              "<col" +
                b +
                ' min="' +
                (A + 1) +
                '" max="' +
                (F + 1) +
                '"' +
                (k === void 0
                  ? ' customWidth="0"'
                  : ' width="' + k / 7 + '" customWidth="1"') +
                "/>",
            ),
              (A = F + 1));
          }
          return y.join("");
        })),
        !a.length)
      ) {
        t[s] = strToU8(d);
        return;
      }
      const u =
        /<sheetData\b[^>]*\/>|<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(
          d,
        );
      if (!u)
        throw new Error("The generated worksheet is missing its data section.");
      const l = new Map();
      for (const p of (u[1] ?? "").matchAll(
        /<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/g,
      )) {
        const m = Number(
          il(
            ((f = /^<row\b([^>]*)>/.exec(p[0])) == null ? void 0 : f[1]) ?? "",
            "r",
          ),
        );
        if (!Number.isInteger(m) || m < 1 || m > Mt.maxRows)
          throw new Error("The generated worksheet has an invalid row number.");
        l.set(m, p[0]);
      }
      for (const p of a)
        for (let m = p.start; m <= p.end; m++) {
          const v = m + 1,
            g = l.get(v);
          l.set(
            v,
            g
              ? g.replace(
                  /^<row\b([^>]*)>/,
                  (b, y) =>
                    "<row" +
                    y
                      .replace(/\s+hidden\s*=\s*(["'])[\s\S]*?\1/g, "")
                      .replace(/\/\s*$/, "") +
                    ' hidden="1"' +
                    (/\/\s*$/.test(y) ? "/>" : ">"),
                )
              : '<row r="' + v + '" hidden="1"/>',
          );
        }
      const c = [...l]
        .sort(([p], [m]) => p - m)
        .map(([, p]) => p)
        .join("");
      t[s] = strToU8(d.replace(u[0], "<sheetData>" + c + "</sheetData>"));
    }),
    zipSync(t)
  );
}
function Aue(e, r) {
  if (
    !r.some((n) => {
      var o, a, s, d;
      return (
        ((a = (o = n.pageLayout) == null ? void 0 : o.rowBreaks) == null
          ? void 0
          : a.length) ||
        ((d = (s = n.pageLayout) == null ? void 0 : s.columnBreaks) == null
          ? void 0
          : d.length)
      );
    })
  )
    return e;
  const t = unzipSync(e);
  return (
    r.forEach((n, o) => {
      var c, h;
      const a = n.pageLayout;
      if (
        !((c = a == null ? void 0 : a.rowBreaks) != null && c.length) &&
        !((h = a == null ? void 0 : a.columnBreaks) != null && h.length)
      )
        return;
      const s = `xl/worksheets/sheet${o + 1}.xml`;
      let d = strFromU8(t[s]);
      d = d.replace(
        /<(?:rowBreaks|colBreaks)\b[^>]*(?:\/>|>[\s\S]*?<\/(?:rowBreaks|colBreaks)>)/g,
        "",
      );
      const l = [
        ["rowBreaks", a.rowBreaks, 16383],
        ["colBreaks", a.columnBreaks, 1048575],
      ]
        .map(([f, p, m]) =>
          p != null && p.length
            ? `<${f} count="${p.length}" manualBreakCount="${p.length}">${p.map((v) => `<brk id="${v}" min="0" max="${m}" man="1"/>`).join("")}</${f}>`
            : "",
        )
        .join("");
      ((d = d.replace(
        /<(?:customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|<\/worksheet>/,
        (f) => l + f,
      )),
        (t[s] = strToU8(d)));
    }),
    zipSync(t)
  );
}
function kue(e, r) {
  if (
    !r.some((n) => {
      var o, a;
      return (
        ((o = n.conditionalFormats) == null
          ? void 0
          : o.some((s) => s.stopIfTrue)) ||
        ((a = n.validations) == null
          ? void 0
          : a.some((s) => s.showDropdown === false))
      );
    })
  )
    return e;
  const t = unzipSync(e);
  return (
    r.forEach((n, o) => {
      const a = `xl/worksheets/sheet${o + 1}.xml`;
      t[a] = strToU8(Ple(Xle(strFromU8(t[a]), n), n));
    }),
    zipSync(t)
  );
}
async function gI(e) {
  var p, m, v, g, b, y, w, E, A, k, F, I, O, j, z, V;
  const r = [
    "Dynamic-array formulas are exported as ordinary anchor formulas without Excel dynamic-array metadata. Native JSON retains local spill behavior. Legacy CSE and metadata-qualified Excel arrays are identified and refused on import; convert a reviewed copy in Excel before importing. Independent Excel round-trip acceptance is required.",
    `Date calculations use the Gregorian ${ts(e.datePolicy).dateSystem} serial system, preserved in XLSX. Native JSON retains the explicit text-date order and calculation time zone; XLSX does not retain that policy or the calculation clock evidence. DATEVALUE/VALUE text interpretation and TODAY/NOW results may change in Excel. Export calculation evidence separately and review calendar-sensitive formulas.`,
    "XLSX exported with supported values, formulas, styles and page layout. Keep native JSON for editable review history, outline groups, links, pivots, charts, shapes and application metadata. Number formats may differ in Excel. External-action formulas are saved as text.",
  ];
  (e.styleCatalog ||
    e.sheets.some((D) =>
      Object.values(D.cells).some((B) => {
        var _;
        return (_ = B.style) == null ? void 0 : _.binding;
      }),
    )) &&
    r.push(
      "Named styles and theme bindings are flattened to their current effective formatting in XLSX. Keep native JSON to preserve reusable styles, inherited values and explicit overrides.",
    );
  const t = rn(e),
    n = Vn(t);
  (t.workbookLinks &&
    r.push(
      "Workbook link targets, reviewed source caches and last-resolved revision evidence are retained in native JSON only. External formulas export as inert text; use Data > Workbook Links to preview breaking them into values before XLSX export. Retained originals are unchanged.",
    ),
    t.extensions &&
      r.push(
        "Custom-function manifests, reviewed result caches and linked entity fields are retained in native JSON only. XLSX does not install or authorize extensions. Custom function formulas require a compatible, separately approved implementation in the receiving application; cached results are historical and not proof of live calculation.",
      ),
    (p = t.sourceImports) != null &&
      p.length &&
      r.push(
        "Structured/document source originals, field paths, native-type evidence, page regions and review decisions are native JSON only and omitted from derived XLSX. Keep a native backup and use File > Import sources to retrieve the originals.",
      ),
    t.alternateSource &&
      r.push(
        "The retained alternate-format original and conversion provenance are native JSON only and are omitted from derived XLSX. Keep the original from File > Workbook formats. Conversion equivalence has not been verified.",
      ),
    (m = t.reportOutputs) != null &&
      m.length &&
      r.push(
        "Reusable report outputs, fixed snapshots and host-reviewed links are native JSON only and are omitted from XLSX. Export each report as SVG/PNG/HTML/JSON for downstream reuse; XLSX does not create Office OLE links or refresh authority.",
      ),
    (v = t.embeddedObjectInventories) != null &&
      v.length &&
      r.push(
        "Unsupported OLE/ActiveX/embedded-object metadata and payloads are omitted from derived XLSX. Every retained inventory item is listed as dropped in embeddedObjectLosses. No object was executed or fetched; preserve the original files separately and native JSON for the inventory.",
      ),
    t.sheets.some((D) => D.analysisReport) &&
      r.push(
        "Solver and statistical analysis definitions, methods, assumptions, exclusions and freshness evidence require native JSON. XLSX preserves historical report cells only; native Solver models and Analysis ToolPak settings are not generated. Reopen native JSON to review and rerun.",
      ),
    (t.dataModel || t.sheets.some((D) => D.dataModelReport)) &&
      r.push(
        "Relational model tables, relationships, typed contracts, measures, queries and source-freshness evidence require native JSON. XLSX exports supported worksheet values only; no Excel Data Model or DAX is generated. Published model results are snapshots and are not recalculated by Excel from these definitions.",
      ),
    (g = t.importedCalculationCaches) != null &&
      g.entries.length &&
      r.push(
        "Original imported formula caches, source file hash and cache comparisons are native JSON evidence only. XLSX writes current local results; it never writes imported historical caches as current values.",
      ),
    ((b = t.calculation) == null ? void 0 : b.mode) === "manual" &&
      r.push(
        "Manual calculation mode is exported. Pending formula caches are omitted; choose Calculate Now before relying on formula results in another application.",
      ),
    (((y = t.automationArtifacts) != null && y.length) ||
      ((w = t.automationRuns) != null && w.length)) &&
      r.push(
        "Preserved VBA/Office Scripts source bytes and migration run provenance are native JSON only. XLSX contains the resulting cells; it does not embed executable macros or Office Scripts. Download original artifacts separately from Automation & Scripts.",
      ),
    (E = t.whatIfAnalyses) != null &&
      E.length &&
      r.push(
        "What-if scenario and sensitivity definitions, retained assumptions and baseline evidence are native JSON metadata. XLSX preserves report cells only; native Excel Scenario Manager and data-table formulas are not created. Keep native JSON to reopen and refresh the analysis.",
      ),
    t.sheets.some((D) => {
      var B;
      return (B = D.pivots) == null ? void 0 : B.length;
    }) &&
      r.push(
        "Native PivotTable definitions, slicers, timelines, Show values as settings, field sorting, PivotCharts, calculated fields, grouping, refresh fingerprints and freshness labels are not preserved in XLSX. Only the current output cells are exported; keep native JSON and export pivot context to verify and refresh the result.",
      ),
    (((A = t.tableEvidence) != null && A.length) ||
      t.sheets.some((D) => {
        var B;
        return (B = D.tables) == null
          ? void 0
          : B.some((_) => _.columns.some((C) => C.dataType));
      })) &&
      r.push(
        "Table column types and original header/conversion evidence are native JSON only. XLSX retains supported tables and current cell values; save native JSON to retain the original source values and type contracts.",
      ),
    (k = t.textCleanups) != null &&
      k.length &&
      r.push(
        "Text cleanup history, original evidence snapshots and cleanup recipes are retained in native JSON only. XLSX keeps the output values and cell notes but cannot restore the cleanup history.",
      ),
    (F = t.savedFilterViews) != null &&
      F.length &&
      r.push(
        "Named filter views and their historical preview counts are retained in native JSON only. XLSX exports the full source population and does not write Excel AutoFilter criteria. Keep native JSON to reopen saved views and review fresh results.",
      ),
    (((I = t.transformationRecipes) != null && I.length) ||
      t.sheets.some((D) => D.transformationRun)) &&
      r.push(
        "Versioned transformation recipes, source/output hashes and complete step execution evidence are native JSON only. XLSX retains current output values and source-row notes; keep the native workbook for repeatable recipe refresh and lineage.",
      ),
    t.sheets.some((D) => D.consolidation) &&
      r.push(
        "Consolidation definitions, approved external source identities, review evidence and live source-freshness checks require native JSON. XLSX retains supported cells, same-workbook formulas and source notes, but no consolidation refresh or external workbook links. External amounts remain reviewed snapshots; keep native JSON and reapprove source workbooks to refresh.",
      ),
    t.sheets.some((D) => D.externalConnection) &&
      r.push(
        "External source connections, refresh attempts, source versions, stale state and control-total evidence are native JSON only. XLSX contains a values snapshot and source notes, with no refresh or credential authority. Keep the native workbook and reauthorize the host source before refreshing.",
      ),
    (O = t.fieldTransforms) != null &&
      O.length &&
      r.push(
        "Clean & Reshape recipes, reviewed inference examples and original source snapshots are native JSON only. XLSX keeps current output values, derivation notes and the exceptions appendix; save native JSON to retain reversible transformation evidence.",
      ),
    t.sheets.some((D) => {
      var B;
      return (B = D.tables) == null
        ? void 0
        : B.some((_) => _.columns.some((C) => C.calculatedFormula));
    }) &&
      r.push(
        "Calculated table column rules are retained in native JSON only. XLSX retains the individual cell formulas and values, but does not preserve the reusable calculated-column rule for future appends.",
      ),
    (j = t.formulaWatches) != null &&
      j.length &&
      r.push(
        "The formula Watch Window list is retained in native JSON only; XLSX cellWatches are not exported.",
      ),
    t.sheets.some((D) => D.splitPanes) &&
      r.push(
        "Split panes are retained in native JSON only and are not preserved in XLSX. Save workbook data to keep divider layouts and initial scroll anchors.",
      ),
    (z = t.rowOrigins) != null &&
      z.length &&
      r.push(
        "Original source row identities and their current associations are retained in native JSON only. XLSX export does not preserve the source registry; reopening that XLSX creates a new file-source identity.",
      ),
    (V = t.definedNames) != null &&
      V.length &&
      r.push(
        "Defined-name formula text, workbook/worksheet scope and supported comments are preserved. Recalculation uses the supported local engine; relative, volatile and dynamic name formulas require independent Excel review.",
      ),
    t.sheets.some((D) => D.delimitedSource) &&
      r.push(
        "Delimited source footprints and schemas are not carried by XLSX. Keep native JSON for the historical source footprint and the separate delimited schema report for type choices.",
      ));
  for (const D of t.sheets)
    ((D.hiddenRows = cn(D, "row")), (D.hiddenCols = cn(D, "column")));
  if (
    t.sheets.reduce(
      (D, B) => D + E1(B, "row").reduce((_, C) => _ + C.end - C.start + 1, 0),
      0,
    ) > x1.maxXlsxHiddenRows
  )
    throw new Error(
      "XLSX export exceeds the 100,000-hidden-row processing budget. Use native JSON to preserve larger hidden ranges without expanding them.",
    );
  const a = new oue.Workbook();
  ((a.creator = "LedgerGlass"),
    (a.properties.date1904 = ts(t.datePolicy).dateSystem === "1904"));
  const s = t.sheets
    .flatMap((D) => Object.values(D.cells))
    .filter((D) => typeof D.raw == "number" && Fle(String(D.raw)));
  (s.length &&
    r.push(
      `Numeric precision risk: ${s.length} stored numeric value(s) exceed Excel's 15-digit precision or numeric range. Export preserves stored numbers, not precision as displayed. Use explicit ROUND for monetary amounts and Text for identifiers; retain native JSON.`,
    ),
    t.metadata &&
      ((a.created = new Date(t.metadata.createdAt)),
      (a.modified = new Date(t.metadata.modifiedAt))));
  const d = P$("export", `${t.name}.xlsx`, 0, t, r);
  for (const D of t.sheets) {
    const B = a.addWorksheet(D.name, {
      state: D.veryHidden ? "veryHidden" : D.hidden ? "hidden" : "visible",
      properties: {
        defaultRowHeight: Dg(D) * 0.75,
        defaultColWidth: Og(D) / 7,
        ...(D.tabColor
          ? { tabColor: { argb: `FF${D.tabColor.slice(1)}` } }
          : {}),
      },
      views:
        D.frozenRows || D.frozenCols
          ? [
              {
                state: "frozen",
                ySplit: D.frozenRows ?? 0,
                xSplit: D.frozenCols ?? 0,
              },
            ]
          : [],
    });
    for (const [_, C] of Object.entries(D.columnWidths ?? {}))
      B.getColumn(Number(_) + 1).width = C / 7;
    for (const [_, C] of Object.entries(D.rowHeights ?? {}))
      B.getRow(Number(_) + 1).height = C * 0.75;
    for (const _ of E1(D, "column"))
      for (let C = _.start; C <= _.end; C++) B.getColumn(C + 1).hidden = true;
    for (const [_, C] of Object.entries(D.cells)) {
      const T = B.getCell(_);
      if (
        typeof C.raw == "string" &&
        C.raw.startsWith("=") &&
        !C.literal &&
        !wue(C.raw)
      ) {
        const N = tue(n.getTypedValue(D.id, _));
        (d.caches.total++,
          N === void 0 ? d.caches.omitted++ : d.caches.written++,
          (T.value = {
            formula: C.raw.slice(1),
            ...(N !== void 0 ? { result: N } : {}),
          }));
      } else
        ((T.value = C.raw),
          typeof C.raw == "string" &&
            C.raw.startsWith("=") &&
            !C.literal &&
            (d.caches.total++, d.caches.neutralized++));
      if (
        ((T.numFmt = C.numberFormat
          ? (Q3(C.numberFormat) ?? "General")
          : (C.numFmt ?? Eue[C.format ?? "general"])),
        C.style)
      ) {
        const N = ys(C.style);
        ((T.font = {
          bold: N.bold,
          italic: N.italic,
          strike: N.strikethrough,
          underline: N.underline,
          name: N.fontFamily,
          size: N.fontSize,
          ...(N.color ? { color: { argb: kk(N.color) } } : {}),
        }),
          N.fill &&
            (T.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: kk(N.fill) },
            }),
          (N.align ||
            N.wrap !== void 0 ||
            N.verticalAlign ||
            N.indent !== void 0) &&
            (T.alignment = {
              horizontal: N.align,
              wrapText: N.wrap,
              vertical: N.verticalAlign,
              indent: N.indent,
            }));
        const R = {};
        for (const L of s0) {
          const $ = md(N, L);
          $ &&
            $.style !== "none" &&
            (R[L] = {
              style: $.style,
              color: { argb: kk($.color ?? "#000000") },
            });
        }
        (N.border || N.borders) && (T.border = R);
      }
      C.note && (T.note = C.note);
    }
    (wie(B, D), Zle(B, D), r.push(...eie(a, B, D)), r.push(...aae(B, D, t)));
  }
  a.calcProperties.fullCalcOnLoad = true;
  const u = await a.xlsx.writeBuffer(),
    l = unzipSync(new Uint8Array(u));
  l["xl/workbook.xml"] = strToU8(
    Yle(dne(strFromU8(l["xl/workbook.xml"]), t), t),
  );
  const c = zipSync(l);
  let h = new Uint8Array(kue(Aue(Cue(c, t.sheets), t.sheets), t.sheets));
  if (t.xlsxSource) {
    const D = WR(t);
    W2(D.bytes, cT());
    const B = unzipSync(D.bytes),
      _ = F$(B, D.sha256),
      C = Mle(B, unzipSync(h), _);
    try {
      if (C.bytes.length > Mt.maxFileBytes)
        throw Error("Preserved metadata exceeds the XLSX reopen budget.");
      (W2(C.bytes, cT()),
        (h = new Uint8Array(C.bytes)),
        (d.partInventory = C.inventory));
    } catch {
      ((d.partInventory = jle(
        _,
        "Opaque metadata could not pass the final output integrity or capacity checks. Retained only in the original; the derived copy excludes it.",
      )),
        r.push(
          "Passive metadata could not be safely included within the export budgets. Inspect the part inventory; download the original for all source content.",
        ));
    }
    r.push(
      `${d.partInventory.counts["preserved-opaque"]} source part(s) preserved byte-for-byte as passive metadata. All source bytes remain available through File → Original XLSX and compatibility and native JSON. Derived XLSX does not embed the original source attachment.`,
    );
  }
  const f = new Blob([h], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  if (f.size > Mt.maxFileBytes)
    throw new Error(
      "XLSX export exceeds the 10 MB reopen limit. Reduce the workbook or use native JSON.",
    );
  return (
    W2(h, {
      names: [],
      nameWarnings: [],
      numericTokens: new Map(),
      builtinNameXml: "",
      hiddenRows: new Map(),
      conditionalStops: new Map(),
      validationAreas: new Map(),
      printBreaks: new Map(),
      columns: new Map(),
      sheetRelationships: new Map(),
      sheetNames: new Map(),
    }),
    d.caches.omitted &&
      r.push(
        `${d.caches.omitted} formula cache(s) omitted because the local result is blank or an application-specific error. Excel must recalculate them; unsupported results are never written as successful text caches.`,
      ),
    d.caches.neutralized &&
      r.push(
        `${d.caches.neutralized} external/action formula(s) exported as literal text.`,
      ),
    (d.fileBytes = f.size),
    (d.warnings = [...new Set(r)]),
    { blob: f, warnings: d.warnings, report: d }
  );
}
function Sk(e) {
  const r = JSON.stringify(rn(e), null, 2);
  if (new TextEncoder().encode(r).length > Mt.maxNativeBytes)
    throw new Error(
      "Native workbook exceeds the 50 MB save limit. Reduce the workbook size before saving.",
    );
  return r;
}
function cT() {
  return {
    names: [],
    nameWarnings: [],
    numericTokens: new Map(),
    builtinNameXml: "",
    hiddenRows: new Map(),
    conditionalStops: new Map(),
    validationAreas: new Map(),
    printBreaks: new Map(),
    columns: new Map(),
    sheetRelationships: new Map(),
    sheetNames: new Map(),
  };
}

export { Ek, F$, Jle, L$, P$, Rle, Sk, U$, V$, W2, cT, gI, jC, uT, x7, z$ };
