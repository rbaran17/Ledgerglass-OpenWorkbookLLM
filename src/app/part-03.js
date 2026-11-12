import qre from "fast-formula-parser";
import { strToU8 } from "fflate";
import {
  string,
  number,
  boolean,
  null as _null,
  array,
  strictObject,
  discriminatedUnion,
  literal,
  record,
  union,
  enum as _enum,
  unknown,
  object,
} from "zod";
import {
  $p,
  AN,
  Cy,
  FJ,
  GF,
  Gf,
  L3,
  LJ,
  Lg,
  Nw,
  O4,
  P3,
  Pg,
  Rg,
  SN,
  TN,
  Tw,
  WS,
  ZF,
  ax,
  eB,
  es,
  g1,
  hr,
  iB,
  iu,
  jJ,
  lf,
  lx,
  nB,
  pd,
  qj,
  ts,
  v1,
  wd,
  yE,
  yc,
  z3,
  zp,
} from "./part-01.js";
import {
  Ig,
  Ja,
  Mt,
  WJ,
  Xy,
  aB,
  b1,
  cx,
  dx,
  iD,
  lB,
  oB,
  qS,
  sB,
  u6,
  ux,
} from "./part-02.js";

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
const s0 = ["top", "right", "bottom", "left"],
  uB = ["none", "thin", "medium", "thick", "dotted", "dashed", "double"],
  GJ = { indentEm: 0.75 },
  ZJ = new Set([
    "binding",
    "bold",
    "italic",
    "underline",
    "strikethrough",
    "fontFamily",
    "fontSize",
    "wrap",
    "border",
    "borders",
    "fill",
    "color",
    "align",
    "verticalAlign",
    "indent",
  ]);
function vs(e) {
  throw Error(e);
}
function _w(e) {
  return !!(
    e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e))
  );
}
function GS(e) {
  return (
    (typeof e != "string" || !/^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(e)) &&
      vs("Style colors must use hexadecimal colors."),
    e
  );
}
function cB(e) {
  return (
    (!_w(e) ||
      Object.keys(e).some((r) => !["style", "color"].includes(r)) ||
      !uB.includes(e.style)) &&
      vs("Choose a supported border line style and hexadecimal color."),
    e.color !== void 0 && GS(e.color),
    e.style === "none"
      ? { style: "none" }
      : { style: e.style, color: e.color === void 0 ? "#000000" : GS(e.color) }
  );
}
function oh(e) {
  _w(e) || vs("Cell style must be an object.");
  for (const t of Object.keys(e))
    ZJ.has(t) || vs(`Unsupported cell style setting: ${t}.`);
  const r = {};
  e.binding !== void 0 && (r.binding = NN(e.binding));
  for (const t of ["bold", "italic", "underline", "strikethrough", "wrap"])
    e[t] !== void 0 &&
      (typeof e[t] != "boolean" && vs(`Style ${t} must be boolean.`),
      (r[t] = e[t]));
  (e.fontFamily !== void 0 &&
    (dx.includes(e.fontFamily) || vs("Unsupported font family."),
    (r.fontFamily = e.fontFamily)),
    e.fontSize !== void 0 &&
      ((typeof e.fontSize != "number" ||
        !Number.isInteger(e.fontSize) ||
        e.fontSize < 8 ||
        e.fontSize > 72) &&
        vs("Font size must be an integer between 8 and 72."),
      (r.fontSize = e.fontSize)));
  for (const t of ["fill", "color"]) e[t] !== void 0 && (r[t] = GS(e[t]));
  if (
    (e.align !== void 0 &&
      (["left", "center", "right"].includes(e.align) ||
        vs("Invalid text alignment."),
      (r.align = e.align)),
    e.verticalAlign !== void 0 &&
      (["top", "middle", "bottom"].includes(e.verticalAlign) ||
        vs("Choose Top, Center or Bottom vertical alignment."),
      (r.verticalAlign = e.verticalAlign)),
    e.indent !== void 0 &&
      ((typeof e.indent != "number" ||
        !Number.isInteger(e.indent) ||
        e.indent < 0 ||
        e.indent > 15) &&
        vs("Indent must be an integer between 0 and 15."),
      (r.indent = e.indent)),
    e.border !== void 0 &&
      (["all", "bottom", "none"].includes(e.border) ||
        vs("Unsupported cell border."),
      (r.border = e.border)),
    e.borders !== void 0)
  ) {
    (!_w(e.borders) || Object.keys(e.borders).some((n) => !s0.includes(n))) &&
      vs("Cell borders support top, right, bottom and left edges only.");
    const t = {};
    for (const n of s0) e.borders[n] !== void 0 && (t[n] = cB(e.borders[n]));
    Object.keys(t).length && (r.borders = t);
  }
  return r;
}
function NN(e) {
  (!_w(e) ||
    Object.keys(e).some(
      (t) =>
        !["namespace", "styleId", "base", "missing", "themeFallback"].includes(
          t,
        ),
    )) &&
    vs(
      "Named style binding must contain a namespace, style ID and supported base.",
    );
  for (const t of ["namespace", "styleId"])
    (typeof e[t] != "string" ||
      !e[t].trim() ||
      e[t].length > 100 ||
      /[\x00-\x1f\x7f]/.test(e[t])) &&
      vs(
        "Named style identifiers must contain 1–100 characters without control characters.",
      );
  (!_w(e.base) || Object.hasOwn(e.base, "binding")) &&
    vs("A named style base cannot contain another binding.");
  for (const t of ["missing", "themeFallback"])
    e[t] !== void 0 &&
      e[t] !== true &&
      vs("Invalid named style fallback flag.");
  const r = oh(e.base);
  return (
    e.missing &&
      Object.keys(r).length &&
      vs("A missing named style must use an empty fallback base."),
    {
      namespace: e.namespace,
      styleId: e.styleId,
      base: r,
      ...(e.missing ? { missing: true } : {}),
      ...(e.themeFallback ? { themeFallback: true } : {}),
    }
  );
}
function ys(e) {
  if (!e) return {};
  const { binding: r, ...t } = e;
  if (!r) return t;
  const n = r.missing ? {} : r.base,
    o = { ...n, ...t };
  return (
    Object.hasOwn(t, "border")
      ? delete o.borders
      : n.borders && (o.borders = { ...n.borders }),
    t.borders && (o.borders = { ...o.borders, ...t.borders }),
    o
  );
}
function md(e, r) {
  var t;
  if (((e = ys(e)), (t = e == null ? void 0 : e.borders) != null && t[r]))
    return e.borders[r];
  if ((e == null ? void 0 : e.border) === "none") return { style: "none" };
  if (
    (e == null ? void 0 : e.border) === "all" ||
    ((e == null ? void 0 : e.border) === "bottom" && r === "bottom")
  )
    return { style: "thin", color: "#000000" };
}
function dB(e) {
  return !e || e.style === "none"
    ? 0
    : e.style === "medium"
      ? 2
      : e.style === "thick" || e.style === "double"
        ? 3
        : 1;
}
function M4(e) {
  if (e)
    return e.style === "none"
      ? "none"
      : `${dB(e)}px ${["dotted", "dashed", "double"].includes(e.style) ? e.style : "solid"} ${e.color ?? "#000000"}`;
}
function fB(e) {
  return (
    (e = ys(e)),
    ((e == null ? void 0 : e.indent) ?? 0) *
      GJ.indentEm *
      (e != null && e.fontSize ? (e.fontSize * 4) / 3 : 14)
  );
}
function Dw(e, r, t = {}) {
  e = ys(e);
  const n =
      (e == null ? void 0 : e.align) ??
      (typeof r == "number" ? "right" : "left"),
    o = fB(e),
    a = [
      (e == null ? void 0 : e.underline) && "underline",
      (e == null ? void 0 : e.strikethrough) && "line-through",
    ]
      .filter(Boolean)
      .join(" ");
  return {
    background: e == null ? void 0 : e.fill,
    color: e == null ? void 0 : e.color,
    fontWeight: e != null && e.bold ? 700 : 400,
    fontStyle: e != null && e.italic ? "italic" : "normal",
    fontFamily: e == null ? void 0 : e.fontFamily,
    fontSize: e != null && e.fontSize ? `${e.fontSize}pt` : void 0,
    textDecoration: a || "none",
    textAlign: n,
    verticalAlign: e == null ? void 0 : e.verticalAlign,
    paddingLeft: 3 + (n === "left" ? o : 0),
    paddingRight: 3 + (n === "right" ? o : 0),
    borderTop: M4(md(e, "top")) ?? t.defaultBorder,
    borderRight: M4(md(e, "right")) ?? t.defaultBorder,
    borderBottom: M4(md(e, "bottom")) ?? t.defaultBorder,
    borderLeft: M4(md(e, "left")) ?? t.defaultBorder,
  };
}
function RN(e) {
  return (
    (e = ys(e)),
    {
      whiteSpace: e != null && e.wrap ? "pre-wrap" : "nowrap",
      overflowWrap: e != null && e.wrap ? "anywhere" : void 0,
    }
  );
}
function hB(e) {
  if (((e = ys(e)), !(e != null && e.borders))) {
    const t =
      (e == null ? void 0 : e.border) === "all"
        ? 2
        : (e == null ? void 0 : e.border) === "none"
          ? 0
          : 1;
    return { horizontal: t, vertical: t };
  }
  const r = (t) => {
    const n = md(e, t);
    return n ? dB(n) : 0.5;
  };
  return {
    horizontal: r("left") + r("right"),
    vertical: r("top") + r("bottom"),
  };
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
function KJ(e, r, t, n, o = {}) {
  o.readOnly && vs("This workbook is read-only.");
  const a = Nr(e, r);
  jn(a);
  const s = oh(n),
    d = Np(typeof t == "string" ? [t] : t),
    u = new Set(d);
  if (o.borders || Object.hasOwn(n, "borders") || Object.hasOwn(n, "border")) {
    const h = d.map(_t);
    (a.merges ?? []).some((f) => {
      const p = ct(f);
      return h.some(
        (m) =>
          m.row >= p.firstRow &&
          m.row <= p.lastRow &&
          m.col >= p.firstCol &&
          m.col <= p.lastCol,
      );
    }) &&
      vs(
        "Unmerge the intersecting cells before applying border edges or presets.",
      );
  }
  const l = d.flatMap((h) => {
    var g, b;
    const f = a.cells[h],
      p = { ...(f == null ? void 0 : f.style), ...n, ...s };
    (Object.hasOwn(n, "border") && delete p.borders,
      o.borders &&
        (p.borders = {
          ...((g = f == null ? void 0 : f.style) == null ? void 0 : g.borders),
          ...pB(h, u, o.borders.preset, o.borders.line),
        }),
      n.borders !== void 0 &&
        (p.borders = {
          ...(o.borders
            ? p.borders
            : (b = f == null ? void 0 : f.style) == null
              ? void 0
              : b.borders),
          ...s.borders,
        }));
    const m = oh(p),
      v = {
        ...(f ?? { raw: null }),
        ...(Object.keys(m).length ? { style: m } : {}),
      };
    return (
      Object.keys(m).length || delete v.style,
      !f && !Object.keys(m).length
        ? []
        : JSON.stringify(oh((f == null ? void 0 : f.style) ?? {})) ===
            JSON.stringify(m)
          ? []
          : JSON.stringify(f) === JSON.stringify(v)
            ? []
            : [{ sheetId: r, address: h, cell: v }]
    );
  });
  if (!l.length) return e;
  const c = p0(e, l, "Preparer", "Format cell styles");
  return { ...e, sheets: c.sheets };
}
const _g = [
    "background",
    "text",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
  ],
  YJ = { maxStyles: 50 },
  wE = {
    name: "Office",
    palette: {
      background: "#FFFFFF",
      text: "#000000",
      accent1: "#4472C4",
      accent2: "#ED7D31",
      accent3: "#A5A5A5",
      accent4: "#FFC000",
      accent5: "#5B9BD5",
      accent6: "#70AD47",
    },
    fonts: { body: "Calibri", heading: "Calibri" },
  },
  c6 = [
    wE,
    {
      name: "Mint",
      palette: {
        background: "#FFFFFF",
        text: "#183B32",
        accent1: "#217346",
        accent2: "#5A9E80",
        accent3: "#A8CFC0",
        accent4: "#E4B458",
        accent5: "#538FA0",
        accent6: "#8B739B",
      },
      fonts: { body: "Aptos", heading: "Aptos" },
    },
    {
      name: "Slate",
      palette: {
        background: "#FFFFFF",
        text: "#243447",
        accent1: "#355C7D",
        accent2: "#6C8EA4",
        accent3: "#B2BBC7",
        accent4: "#D4A95C",
        accent5: "#527A72",
        accent6: "#916B80",
      },
      fonts: { body: "Arial", heading: "Arial" },
    },
    {
      name: "Amber",
      palette: {
        background: "#FFFFFF",
        text: "#3D3226",
        accent1: "#A66A15",
        accent2: "#DAA94E",
        accent3: "#C7B59E",
        accent4: "#8E674E",
        accent5: "#6C8A75",
        accent6: "#7C739B",
      },
      fonts: { body: "Calibri", heading: "Times New Roman" },
    },
  ];
function ho(e) {
  throw Error(e);
}
function ah(e) {
  return !!(
    e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e))
  );
}
function mg(e, r, t) {
  Object.keys(e).some((n) => !r.includes(n)) && ho(`Unsupported ${t} setting.`);
}
function Rp(e, r, t = 80) {
  return (
    (typeof e != "string" ||
      !e.trim() ||
      e.length > t ||
      /[\x00-\x1f\x7f]/.test(e)) &&
      ho(`${r} must contain 1–${t} characters without control characters.`),
    e.trim()
  );
}
function Ac(e) {
  return structuredClone(e);
}
function Ql(e) {
  return Array.isArray(e)
    ? `[${e.map(Ql).join(",")}]`
    : ah(e)
      ? `{${Object.keys(e)
          .filter((r) => e[r] !== void 0)
          .sort()
          .map((r) => `${JSON.stringify(r)}:${Ql(e[r])}`)
          .join(",")}}`
      : JSON.stringify(e);
}
function mB(e) {
  (ah(e) || ho("Choose a workbook theme."),
    mg(e, ["name", "palette", "fonts"], "theme"),
    (!ah(e.palette) || !ah(e.fonts)) &&
      ho("A theme needs eight colors and body and heading fonts."),
    mg(e.palette, _g, "theme palette"),
    mg(e.fonts, ["body", "heading"], "theme font"));
  const r = {};
  for (const t of _g) {
    const n = e.palette[t];
    ((typeof n != "string" || !/^#[0-9a-f]{6}$/i.test(n)) &&
      ho("Theme colors must use six-digit hexadecimal colors."),
      (r[t] = n.toUpperCase()));
  }
  for (const t of ["body", "heading"])
    dx.includes(e.fonts[t]) || ho("Choose a supported theme font.");
  return {
    name: Rp(e.name, "Theme name"),
    palette: r,
    fonts: { body: e.fonts.body, heading: e.fonts.heading },
  };
}
function ZS(e) {
  var t;
  (ah(e) || ho("Named style definition must be an object."),
    mg(
      e,
      ["base", "fontRole", "fillRole", "colorRole", "borderRoles"],
      "named style",
    ),
    (!ah(e.base) || Object.hasOwn(e.base, "binding")) &&
      ho(
        "A named style definition needs supported formatting without another binding.",
      ));
  const r = { base: oh(e.base) };
  e.fontRole !== void 0 &&
    (["body", "heading"].includes(e.fontRole) ||
      ho("Choose Body or Heading theme font."),
    (r.fontRole = e.fontRole));
  for (const n of ["fillRole", "colorRole"])
    e[n] !== void 0 &&
      (_g.includes(e[n]) || ho("Choose a supported theme color role."),
      (r[n] = e[n]));
  if (e.borderRoles !== void 0) {
    (ah(e.borderRoles) || ho("Border color roles must be an object."),
      mg(e.borderRoles, s0, "border role"));
    const n = {};
    for (const o of s0)
      e.borderRoles[o] !== void 0 &&
        (_g.includes(e.borderRoles[o]) ||
          ho("Choose a supported border color role."),
        (!md(r.base, o) ||
          ((t = md(r.base, o)) == null ? void 0 : t.style) === "none") &&
          ho("Add a border line before assigning its theme color."),
        (n[o] = e.borderRoles[o]));
    Object.keys(n).length && (r.borderRoles = n);
  }
  return r;
}
function gB(e) {
  return IN({
    namespace: e,
    theme: wE,
    styles: [
      {
        id: "normal",
        name: "Normal",
        definition: { base: {}, fontRole: "body", colorRole: "text" },
      },
      {
        id: "heading",
        name: "Heading",
        definition: {
          base: { bold: true, fontSize: 12 },
          fontRole: "heading",
          fillRole: "accent1",
          colorRole: "background",
        },
      },
      {
        id: "input",
        name: "Input",
        definition: {
          base: { italic: true },
          fontRole: "body",
          colorRole: "accent1",
        },
      },
      {
        id: "total",
        name: "Total",
        definition: {
          base: {
            bold: true,
            borders: { top: { style: "double", color: "#000000" } },
          },
          fontRole: "body",
          colorRole: "text",
          borderRoles: { top: "accent1" },
        },
      },
      {
        id: "warning",
        name: "Warning",
        definition: {
          base: { bold: true },
          fontRole: "body",
          fillRole: "accent4",
          colorRole: "text",
        },
      },
    ],
  });
}
function IN(e) {
  (ah(e) || ho("Workbook styles must be an object."),
    mg(
      e,
      ["namespace", "theme", "styles", "missingTheme"],
      "workbook style catalog",
    ),
    e.missingTheme !== void 0 &&
      e.missingTheme !== true &&
      ho("Invalid missing theme flag."),
    (!Array.isArray(e.styles) || e.styles.length > YJ.maxStyles) &&
      ho("A workbook supports at most 50 named styles."));
  const r = new Set(),
    t = new Set(),
    n = e.styles.map((s) => {
      (ah(s) || ho("Invalid named style."),
        mg(s, ["id", "name", "definition"], "named style"));
      const d = Rp(s.id, "Style ID", 100),
        u = Rp(s.name, "Style name");
      return (
        (r.has(d) || t.has(u.toLocaleLowerCase("en-US"))) &&
          ho("Named style IDs and names must be unique within this workbook."),
        r.add(d),
        t.add(u.toLocaleLowerCase("en-US")),
        { id: d, name: u, definition: ZS(s.definition) }
      );
    }),
    o = e.theme === void 0 || e.missingTheme === true,
    a = e.theme === void 0 ? Ac(wE) : mB(e.theme);
  return (
    o &&
      Ql(a) !== Ql(wE) &&
      ho("A missing theme must use the Office fallback."),
    {
      namespace: Rp(e.namespace, "Style namespace", 100),
      theme: a,
      styles: n,
      ...(o ? { missingTheme: true } : {}),
    }
  );
}
function vB(e, r) {
  var n;
  const t = Ac(e.base);
  (e.fontRole && (t.fontFamily = r.fonts[e.fontRole]),
    e.fillRole && (t.fill = r.palette[e.fillRole]),
    e.colorRole && (t.color = r.palette[e.colorRole]));
  for (const o of s0)
    (n = e.borderRoles) != null &&
      n[o] &&
      (t.borders = {
        ...t.borders,
        [o]: { ...md(t, o), color: r.palette[e.borderRoles[o]] },
      });
  return oh(t);
}
function xE(e, r) {
  e &&
    r.namespace !== e.namespace &&
    ho(
      "This style belongs to another workbook. Import or remap its style catalog before pasting the binding.",
    );
  const t = e == null ? void 0 : e.styles.find((n) => n.id === r.styleId);
  return {
    namespace: r.namespace,
    styleId: r.styleId,
    base: t && e ? vB(t.definition, e.theme) : {},
    ...(t ? {} : { missing: true }),
    ...(e != null && e.missingTheme ? { themeFallback: true } : {}),
  };
}
function JJ(e, r) {
  var o;
  const t = e === void 0 ? void 0 : IN(e);
  t ? (r.styleCatalog = t) : delete r.styleCatalog;
  const n = ah(e) && e.theme === void 0;
  for (const a of r.sheets)
    for (const s of qj(a.cells) ? Object.values(a.cells) : []) {
      if (!((o = s.style) != null && o.binding)) continue;
      const d = NN(s.style.binding),
        u = xE(t, d);
      (!u.missing &&
        !n &&
        Ql(d) !== Ql(u) &&
        ho(
          "Named style appearance does not match the workbook catalog. Reapply the style from a trusted catalog.",
        ),
        (s.style.binding = u));
    }
}
function _N(e, r) {
  if (!(r != null && r.binding)) return;
  const t = NN(r.binding),
    n = xE(e.styleCatalog, t);
  Ql(t) !== Ql(n) &&
    ho(
      "Named style binding is stale or inconsistent with this workbook. Reapply its named style.",
    );
}
function KS(e) {
  var r, t;
  if ((r = e == null ? void 0 : e.binding) != null && r.missing)
    return "Named style is missing. Default formatting and direct cell formatting are shown.";
  if ((t = e == null ? void 0 : e.binding) != null && t.themeFallback)
    return "Workbook theme is missing. The Office theme is shown.";
}
function bB(e) {
  var t;
  const r = new Set();
  (t = e.styleCatalog) != null &&
    t.missingTheme &&
    r.add("Workbook theme is missing. The Office theme is shown.");
  for (const n of e.sheets)
    for (const o of Object.values(n.cells)) {
      const a = KS(o.style);
      a && r.add(a);
    }
  return [...r];
}
function QJ(e, r) {
  return s0.some((t) => Ql(md(e, t)) !== Ql(md(r, t)));
}
function yB(e, r, t = {}) {
  var m;
  t.readOnly && ho("This workbook is read-only.");
  for (const v of e.sheets)
    for (const g of Object.values(v.cells)) _N(e, g.style);
  const n = e.styleCatalog,
    o = Ac(n ?? gB(e.id));
  let a = Ac(r);
  const s = [],
    d = [],
    u = (v) =>
      o.styles.find((g) => g.id === v) ??
      ho("This named style no longer exists.");
  function l(v, g, b) {
    var E;
    const y = Nr(e, v),
      w = y.cells[g];
    if (Ql(w) !== Ql(b)) {
      if ((jn(y), QJ(w == null ? void 0 : w.style, b.style))) {
        const A = _t(g);
        (E = y.merges) != null &&
          E.some((k) => {
            const F = ct(k);
            return (
              A.row >= F.firstRow &&
              A.row <= F.lastRow &&
              A.col >= F.firstCol &&
              A.col <= F.lastCol
            );
          }) &&
          ho(
            "Unmerge the intersecting cells before changing named-style border edges.",
          );
      }
      (s.push({
        sheetId: v,
        address: g,
        ...(w ? { before: Ac(w) } : {}),
        after: Ac(b),
      }),
        s.length > 500 &&
          ho(
            "This action changes more than 500 cells. Use smaller selections or fewer bound cells.",
          ));
    }
  }
  function c(v) {
    var g;
    for (const b of e.sheets)
      for (const [y, w] of Object.entries(b.cells))
        (g = w.style) != null &&
          g.binding &&
          (!v || w.style.binding.styleId === v) &&
          (jn(b),
          l(b.id, y, {
            ...w,
            style: { ...w.style, binding: xE(o, w.style.binding) },
          }));
  }
  switch (a.kind) {
    case "setTheme":
      ((o.theme = mB(a.theme)),
        delete o.missingTheme,
        (a.theme = Ac(o.theme)),
        c());
      break;
    case "createStyle": {
      const g = Rp(a.name, "Style name"),
        b =
          a.styleId === void 0
            ? `style-${crypto.randomUUID()}`
            : Rp(a.styleId, "Style ID", 100),
        y = ZS(a.definition);
      (o.styles.push({ id: b, name: g, definition: y }),
        (a = { ...a, name: g, styleId: b, definition: y }),
        c(b));
      break;
    }
    case "updateStyle": {
      const g = u(a.styleId);
      ((g.definition = ZS(a.definition)),
        (a.definition = Ac(g.definition)),
        c(g.id));
      break;
    }
    case "renameStyle": {
      const g = u(a.styleId);
      ((g.name = Rp(a.name, "Style name")), (a.name = g.name), c(g.id));
      break;
    }
    case "deleteStyle":
      u(a.styleId);
      const v = a.styleId;
      ((o.styles = o.styles.filter((g) => g.id !== v)),
        c(a.styleId),
        d.push(
          "Cells using this style will show default formatting plus their direct overrides. Detach the style first to retain its current appearance.",
        ));
      break;
    case "applyStyle":
    case "reapplyStyle":
    case "detachStyle": {
      const g = Nr(e, a.sheetId);
      (jn(g), a.kind !== "detachStyle" && u(a.styleId));
      for (const b of Np(a.ranges)) {
        const y = g.cells[b] ?? { raw: null };
        if (a.kind === "detachStyle") {
          if (!((m = y.style) != null && m.binding)) continue;
          const w = ys(y.style),
            E = { ...y, style: w };
          (Object.keys(w).length || delete E.style, l(g.id, b, E));
        } else {
          const w = xE(o, { namespace: o.namespace, styleId: a.styleId });
          l(g.id, b, {
            ...y,
            style: { ...(a.kind === "applyStyle" ? y.style : {}), binding: w },
          });
        }
      }
      break;
    }
    default:
      ho("Choose a supported theme or named style action.");
  }
  const h = IN(o),
    f = a.kind === "detachStyle" && !n ? void 0 : h,
    p = {
      setTheme: "Change workbook theme",
      createStyle: "Create named style",
      updateStyle: "Update named style",
      renameStyle: "Rename named style",
      deleteStyle: "Delete named style",
      applyStyle: "Apply named style",
      reapplyStyle: "Reapply named style and clear direct style overrides",
      detachStyle: "Detach named style and keep appearance",
    };
  return {
    version: 1,
    snapshot: JSON.stringify(e),
    request: a,
    ...(n ? { catalogBefore: Ac(n) } : {}),
    ...(f ? { catalogAfter: f } : {}),
    changes: s,
    count: s.length,
    summary: p[a.kind],
    warnings: d,
  };
}
function eQ(e, r, t = {}) {
  (t.readOnly && ho("This workbook is read-only."),
    (r.version !== 1 || r.snapshot !== JSON.stringify(e)) &&
      ho(
        "The workbook changed after this preview. Close and reopen Themes & Cell Styles.",
      ));
  const n = yB(e, r.request, t);
  if (
    (Ql(n) !== Ql(r) &&
      ho(
        "This theme preview changed. Create a fresh preview before applying it.",
      ),
    !r.count && Ql(r.catalogBefore) === Ql(r.catalogAfter))
  )
    return e;
  const o = Ac(e);
  r.catalogAfter
    ? (o.styleCatalog = Ac(r.catalogAfter))
    : delete o.styleCatalog;
  for (const a of r.changes) Nr(o, a.sheetId).cells[a.address] = Ac(a.after);
  return rn(o);
}
function tQ(e, r) {
  var s;
  const t = Rp(r, "Style namespace", 100),
    n = Ac(e),
    o = (s = n.styleCatalog) == null ? void 0 : s.namespace;
  if (!o) return n;
  n.styleCatalog.namespace = t;
  const a = (d) => {
    var u, l;
    ((l = (u = d == null ? void 0 : d.style) == null ? void 0 : u.binding) ==
    null
      ? void 0
      : l.namespace) === o && (d.style.binding.namespace = t);
  };
  for (const d of n.sheets) for (const u of Object.values(d.cells)) a(u);
  for (const d of n.events)
    for (const u of d.changes) (a(u.before), a(u.after));
  return n;
}
function rQ(e, r) {
  if (
    !Number.isInteger(e.index) ||
    e.index < 0 ||
    !Number.isInteger(e.count) ||
    e.count < 1 ||
    e.count > 1e4 ||
    e.index + e.count > r ||
    !Number.isInteger(e.moveBefore) ||
    e.moveBefore < 0 ||
    e.moveBefore > r
  )
    throw new Error(
      "Choose a valid source, 1–10,000 rows or columns, and a destination within the worksheet.",
    );
  if (e.permutation !== void 0) {
    if (
      !Array.isArray(e.permutation) ||
      e.permutation.length !== e.count ||
      new Set(e.permutation).size !== e.count ||
      e.permutation.some(
        (t) => !Number.isInteger(t) || t < e.index || t >= e.index + e.count,
      )
    )
      throw Error(
        "A row permutation must contain every destination exactly once within its source block.",
      );
    return;
  }
  if (e.moveBefore >= e.index && e.moveBefore <= e.index + e.count)
    throw new Error(
      "The destination is inside or immediately after the selected block; this would not move anything.",
    );
}
function Jf(e, r) {
  if (r.permutation)
    return e >= r.index && e < r.index + r.count
      ? r.permutation[e - r.index]
      : e;
  const { index: t, count: n, moveBefore: o } = r,
    a = o < t ? o : o - n;
  return e >= t && e < t + n
    ? a + e - t
    : o < t && e >= o && e < t
      ? e + n
      : o > t && e >= t + n && e < o
        ? e - n
        : e;
}
function wB(e, r, t) {
  if (t.permutation) {
    const o = [],
      a = (s, d, u) => {
        if (s > d) return;
        const l = o.at(-1);
        l && l.end + 1 === s && l.offset === u
          ? (l.end = d)
          : o.push({ start: s, end: d, offset: u });
      };
    a(e, Math.min(r, t.index - 1), 0);
    for (
      let s = Math.max(e, t.index);
      s <= Math.min(r, t.index + t.count - 1);
      s++
    )
      a(s, s, Jf(s, t) - s);
    return (a(Math.max(e, t.index + t.count), r, 0), o);
  }
  const n = [
    ...new Set(
      [e, r + 1, t.index, t.index + t.count, t.moveBefore].filter(
        (o) => o >= e && o <= r + 1,
      ),
    ),
  ].sort((o, a) => o - a);
  return n
    .slice(0, -1)
    .map((o, a) => ({ start: o, end: n[a + 1] - 1, offset: Jf(o, t) - o }));
}
function xB(e, r, t) {
  if (t.permutation && e <= t.index && r >= t.index + t.count - 1)
    return [{ start: e, end: r }];
  const n = wB(e, r, t)
      .map((a) => ({ start: a.start + a.offset, end: a.end + a.offset }))
      .sort((a, s) => a.start - s.start),
    o = [];
  for (const a of n) {
    const s = o.at(-1);
    s && a.start <= s.end + 1 ? (s.end = Math.max(s.end, a.end)) : o.push(a);
  }
  return o;
}
function H3(e, r, t, n = "range", o = false) {
  const a = xB(e, r, t);
  if (a.length !== 1)
    throw new Error(
      `This move would split ${n} into separate ranges. Move the complete range or use explicit cell references first.`,
    );
  if (o && new Set(wB(e, r, t).map((s) => s.offset)).size !== 1)
    throw new Error(
      `This move would reorder part of ${n}. Move the complete structure together.`,
    );
  return a[0];
}
const Jl = {
    threads: 1e3,
    messagesPerThread: 100,
    totalMessages: 5e3,
    textLength: 4e3,
    authorLength: 100,
    urlLength: 2048,
    labelLength: 256,
  },
  YS =
    /[\u0000-\u001f\u007f-\u009f\u200b-\u200f\u202a-\u202e\u2060-\u206f\ufeff]/,
  ji = (e) => {
    throw new Error(e);
  };
function Xm(e, r) {
  return (
    (!e ||
      typeof e != "object" ||
      Array.isArray(e) ||
      ![Object.prototype, null].includes(Object.getPrototypeOf(e))) &&
      ji(`${r} must be a plain object.`),
    e
  );
}
function kc(e, r, t, n = false) {
  return (
    (typeof e != "string" ||
      !e.trim() ||
      e.length > t ||
      YS.test(n ? e.replace(/[\n\r\t]/g, "") : e)) &&
      ji(
        `${r} must contain 1–${t} characters without unsupported control characters.`,
      ),
    e.trim()
  );
}
function Lv(e, r = "Review timestamp") {
  return (
    (typeof e != "string" ||
      !Number.isFinite(Date.parse(e)) ||
      new Date(e).toISOString() !== e) &&
      ji(`${r} must be an ISO UTC timestamp.`),
    e
  );
}
function q3(e) {
  if (typeof e != "string") return ji("Choose a valid worksheet cell address.");
  const r = _t(e.trim());
  return `${ut(r.col)}${r.row + 1}`;
}
function w1(e = "Reviewer") {
  return kc(e, "Reviewer name", Jl.authorLength);
}
function fx(e = {}, r) {
  const t =
    e.timestamp === void 0
      ? new Date(Math.max(Date.now(), r ? Date.parse(r) : 0)).toISOString()
      : Lv(e.timestamp);
  return (
    r && t < r && ji("The review time cannot precede the existing discussion."),
    t
  );
}
function DN(e, r, t) {
  (t.readOnly && ji("This workbook is read-only."), jn(Nr(e, r)));
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
function nQ(e, r, t, n, o = "Reviewer", a = {}) {
  DN(e, r, a);
  const s = Nr(e, r),
    d = fx(a),
    u = w1(o),
    l = {
      id: crypto.randomUUID(),
      anchor: { sheetId: r, sheetName: s.name, address: q3(t) },
      createdAt: d,
      updatedAt: d,
      messages: [
        {
          id: crypto.randomUUID(),
          author: u,
          text: kc(n, "Comment", Jl.textLength, true),
          createdAt: d,
        },
      ],
    };
  return ON(e, [...(e.reviewThreads ?? []), l]);
}
function CB(e, r, t) {
  var o;
  t.readOnly && ji("This workbook is read-only.");
  const n = (o = e.reviewThreads) == null ? void 0 : o.find((a) => a.id === r);
  return n
    ? (n.orphaned &&
        ji(
          "This discussion is retained for review because its target was deleted.",
        ),
      DN(e, n.anchor.sheetId, t),
      n)
    : ji("The discussion could not be found. Reopen the review tools.");
}
function iQ(e, r, t, n = "Reviewer", o = {}) {
  const a = CB(e, r, o);
  a.resolved && ji("Reopen the discussion before adding a reply.");
  const s = fx(o, a.updatedAt),
    d = {
      ...a,
      updatedAt: s,
      messages: [
        ...a.messages,
        {
          id: crypto.randomUUID(),
          author: w1(n),
          text: kc(t, "Reply", Jl.textLength, true),
          createdAt: s,
        },
      ],
    };
  return ON(
    e,
    e.reviewThreads.map((u) => (u.id === r ? d : u)),
  );
}
function oQ(e, r, t, n = "Reviewer", o = {}) {
  typeof t != "boolean" &&
    ji("Choose whether to resolve or reopen the discussion.");
  const a = CB(e, r, o);
  if (!!a.resolved === t) return e;
  const s = fx(o, a.updatedAt),
    d = { ...a, updatedAt: s };
  return (
    t
      ? (d.resolved = { author: w1(n), createdAt: s })
      : (w1(n), delete d.resolved),
    ON(
      e,
      e.reviewThreads.map((u) => (u.id === r ? d : u)),
    )
  );
}
function oD(e, r, t, n, o = {}) {
  var f;
  DN(e, r, o);
  const a = Nr(e, r),
    s = q3(t),
    d = a.cells[s],
    u = _t(s);
  (f = a.merges) != null &&
    f.some((p) => {
      const m = ct(p);
      return (
        u.row >= m.firstRow &&
        u.row <= m.lastRow &&
        u.col >= m.firstCol &&
        u.col <= m.lastCol &&
        (u.row !== m.firstRow || u.col !== m.firstCol)
      );
    }) &&
    ji(
      "Choose the merged area's first cell or unmerge it before changing its link.",
    );
  const l = n === null ? void 0 : X3(n);
  if (
    ((l == null ? void 0 : l.kind) === "internal" && Nr(e, l.sheetId),
    JSON.stringify(d == null ? void 0 : d.hyperlink) === JSON.stringify(l))
  )
    return e;
  !d &&
    e.sheets.reduce((p, m) => p + Object.keys(m.cells).length, 0) >=
      Mt.maxCells &&
    ji("Adding this link would exceed the workbook populated-cell limit.");
  const c = { ...(d ?? { raw: null }) };
  l ? (c.hyperlink = l) : delete c.hyperlink;
  const h = { ...a.cells, [s]: c };
  return (
    c.raw === null && Object.keys(c).length === 1 && delete h[s],
    { ...e, sheets: e.sheets.map((p) => (p.id === r ? { ...p, cells: h } : p)) }
  );
}
function aQ(e, r) {
  var s;
  const t = new Map(e.sheets.map((d) => [d.id, d]));
  let n = false;
  const o =
      (s = e.reviewThreads) == null
        ? void 0
        : s.map((d) => {
            if (d.orphaned) return d;
            const u = t.get(d.anchor.sheetId);
            if (!u) {
              n = true;
              const l = fx({ timestamp: r }, d.updatedAt);
              return {
                ...d,
                updatedAt: l,
                orphaned: { reason: "worksheet-deleted", createdAt: l },
              };
            }
            return u.name !== d.anchor.sheetName
              ? ((n = true),
                { ...d, anchor: { ...d.anchor, sheetName: u.name } })
              : d;
          }),
    a = e.sheets.map((d) => {
      var l;
      let u;
      for (const [c, h] of Object.entries(d.cells))
        ((l = h.hyperlink) == null ? void 0 : l.kind) !== "internal" ||
          t.has(h.hyperlink.sheetId) ||
          (u ?? (u = { ...d.cells }),
          (u[c] = { ...h }),
          delete u[c].hyperlink,
          (n = true));
      return u ? { ...d, cells: u } : d;
    });
  return n ? { ...e, sheets: a, ...(o ? { reviewThreads: o } : {}) } : e;
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
function lQ(e, r, t) {
  const n = Nr(e, t);
  let o = false;
  const a = Object.fromEntries(
    Object.entries(n.cells).map(([s, d]) => {
      var u;
      return ((u = d.hyperlink) == null ? void 0 : u.kind) !== "internal" ||
        d.hyperlink.sheetId !== r
        ? [s, d]
        : ((o = true),
          [s, { ...d, hyperlink: { ...d.hyperlink, sheetId: t } }]);
    }),
  );
  return o
    ? {
        ...e,
        sheets: e.sheets.map((s) => (s.id === t ? { ...s, cells: a } : s)),
      }
    : e;
}
function uQ(e, r) {
  return {
    ...e,
    ...(e.reviewThreads
      ? {
          reviewThreads: e.reviewThreads.map((t) => ({
            ...t,
            anchor: {
              ...t.anchor,
              sheetId: r.get(t.anchor.sheetId) ?? t.anchor.sheetId,
            },
          })),
        }
      : {}),
    sheets: e.sheets.map((t) => ({
      ...t,
      cells: Object.fromEntries(
        Object.entries(t.cells).map(([n, o]) => {
          var a;
          return [
            n,
            ((a = o.hyperlink) == null ? void 0 : a.kind) === "internal"
              ? {
                  ...o,
                  hyperlink: {
                    ...o.hyperlink,
                    sheetId: r.get(o.hyperlink.sheetId) ?? o.hyperlink.sheetId,
                  },
                }
              : o,
          ];
        }),
      ),
    })),
  };
}
function cQ(e, r) {
  const t = uQ(e, r);
  return t.reviewThreads
    ? {
        ...t,
        reviewThreads: t.reviewThreads.map((n) => ({
          ...n,
          id: crypto.randomUUID(),
          messages: n.messages.map((o) => ({ ...o, id: crypto.randomUUID() })),
        })),
      }
    : t;
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
function fQ(e, r) {
  var t;
  return (t = e.formulaWatches) != null && t.length
    ? {
        ...e,
        formulaWatches: e.formulaWatches.map((n) =>
          !n.orphaned && n.anchor.sheetId === r
            ? { ...n, orphaned: "deleted-sheet" }
            : n,
        ),
      }
    : e;
}
function hQ(e, r) {
  var t;
  return (t = e.formulaWatches) != null && t.length
    ? {
        ...e,
        formulaWatches: e.formulaWatches.map((n) => ({
          ...n,
          id: crypto.randomUUID(),
          anchor: {
            ...n.anchor,
            sheetId: r.get(n.anchor.sheetId) ?? n.anchor.sheetId,
          },
        })),
      }
    : e;
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
function Pv(e) {
  return [
    "top-left",
    ...((e == null ? void 0 : e.vertical) !== void 0 ? ["top-right"] : []),
    ...((e == null ? void 0 : e.horizontal) !== void 0
      ? [
          "bottom-left",
          ...((e == null ? void 0 : e.vertical) !== void 0
            ? ["bottom-right"]
            : []),
        ]
      : []),
  ];
}
function aD(e, r, t) {
  return (
    e.split("-")[t === "row" ? 0 : 1] === r.split("-")[t === "row" ? 0 : 1]
  );
}
function AB(e, r) {
  if (e === void 0) return;
  if (
    !e ||
    typeof e != "object" ||
    Array.isArray(e) ||
    (Object.getPrototypeOf(e) !== Object.prototype &&
      Object.getPrototypeOf(e) !== null) ||
    Object.keys(e).some(
      (n) => !["horizontal", "vertical", "row", "col"].includes(n),
    )
  )
    throw Error(
      "Split panes require divider percentages and valid row/column anchors.",
    );
  const t = e;
  if (t.horizontal === void 0 && t.vertical === void 0)
    throw Error("Choose a horizontal split, vertical split, or both.");
  for (const n of [t.horizontal, t.vertical])
    if (n !== void 0 && (!Number.isFinite(n) || n < 20 || n > 80))
      throw Error("Split dividers must be between 20% and 80%.");
  if (
    !Number.isInteger(t.row) ||
    t.row < 0 ||
    t.row >= r.rows ||
    !Number.isInteger(t.col) ||
    t.col < 0 ||
    t.col >= r.cols
  )
    throw Error(
      "Split anchors must be valid zero-based worksheet coordinates.",
    );
  return {
    row: t.row,
    col: t.col,
    ...(t.horizontal !== void 0 ? { horizontal: t.horizontal } : {}),
    ...(t.vertical !== void 0 ? { vertical: t.vertical } : {}),
  };
}
function EE(e, r) {
  const t = r === "row" ? b1 : ux,
    n = r === "row" ? Ig : cx;
  if (typeof e != "number" || !Number.isFinite(e) || e < t || e > n)
    throw Error(
      (r === "row" ? "Default row height" : "Default column width") +
        " must be " +
        t +
        "–" +
        n +
        " pixels.",
    );
  return e;
}
function mQ(e, r) {
  const t =
      e.defaultRowHeight === void 0 ? void 0 : EE(e.defaultRowHeight, "row"),
    n =
      e.defaultColumnWidth === void 0
        ? void 0
        : EE(e.defaultColumnWidth, "column");
  (t !== void 0 && (r.defaultRowHeight = t),
    n !== void 0 && (r.defaultColumnWidth = n));
}
function Dg(e) {
  return e.defaultRowHeight ?? lB;
}
function Og(e) {
  return e.defaultColumnWidth ?? sB;
}
function gg(e, r) {
  var t;
  if (!Number.isInteger(r) || r < 0 || r >= e.rowCount)
    throw Error("Row index is outside the worksheet.");
  return ((t = e.rowHeights) == null ? void 0 : t[r]) ?? Dg(e);
}
function Ip(e, r) {
  var t;
  if (!Number.isInteger(r) || r < 0 || r >= e.colCount)
    throw Error("Column index is outside the worksheet.");
  return ((t = e.columnWidths) == null ? void 0 : t[r]) ?? Og(e);
}
function gQ(e) {
  return Object.freeze({
    workbookId: e.id,
    revision: e.revision,
    snapshotGuard: JSON.stringify(e),
  });
}
function vQ(e, r, t, n = {}) {
  if (n.readOnly !== void 0 && typeof n.readOnly != "boolean")
    throw Error("Read-only state must be boolean.");
  if (n.readOnly) throw Error("This workbook is read-only.");
  const o = n.expectedSnapshot;
  if (
    o !== void 0 &&
    (!o ||
      o.workbookId !== e.id ||
      o.revision !== e.revision ||
      o.snapshotGuard !== JSON.stringify(e))
  )
    throw Error(
      "The workbook changed after these sizes were opened. Close the dialog and reopen it before applying sizes.",
    );
  const a = e.sheets.find((l) => l.id === r);
  if (!a) throw Error("Worksheet not found.");
  if (
    (jn(a),
    !t ||
      typeof t != "object" ||
      Array.isArray(t) ||
      (Object.getPrototypeOf(t) !== Object.prototype &&
        Object.getPrototypeOf(t) !== null) ||
      Object.keys(t).some(
        (l) =>
          ![
            "rowHeight",
            "columnWidth",
            "resetRowOverrides",
            "resetColumnOverrides",
          ].includes(l),
      ))
  )
    throw Error("Choose valid worksheet size settings.");
  for (const l of [t.resetRowOverrides, t.resetColumnOverrides])
    if (l !== void 0 && typeof l != "boolean")
      throw Error("Reset choices must be boolean.");
  const s = t.rowHeight === void 0 ? void 0 : EE(t.rowHeight, "row"),
    d = t.columnWidth === void 0 ? void 0 : EE(t.columnWidth, "column"),
    u = { ...a };
  return (
    s !== void 0 &&
      s !== Dg(a) &&
      (s === lB ? delete u.defaultRowHeight : (u.defaultRowHeight = s)),
    d !== void 0 &&
      d !== Og(a) &&
      (d === sB ? delete u.defaultColumnWidth : (u.defaultColumnWidth = d)),
    t.resetRowOverrides &&
      Object.keys(a.rowHeights ?? {}).length &&
      delete u.rowHeights,
    t.resetColumnOverrides &&
      Object.keys(a.columnWidths ?? {}).length &&
      delete u.columnWidths,
    u.defaultRowHeight === a.defaultRowHeight &&
    u.defaultColumnWidth === a.defaultColumnWidth &&
    u.rowHeights === a.rowHeights &&
    u.columnWidths === a.columnWidths
      ? e
      : { ...e, sheets: e.sheets.map((l) => (l.id === r ? u : l)) }
  );
}
const x1 = {
    maxIntervalsPerAxis: 1e4,
    maxFrozenVisibleRows: 100,
    maxFrozenVisibleCols: 50,
    maxXlsxHiddenRows: 1e5,
  },
  MN = (e, r) => (r === "row" ? e.rowCount : e.colCount),
  G3 = (e) => (e === "row" ? "hiddenRows" : "hiddenCols");
function Z3(e) {
  if (e !== "row" && e !== "column") throw Error("Choose rows or columns.");
}
function bQ(e) {
  if (!Number.isSafeInteger(e) || e < 0)
    throw Error("Axis capacity must be a nonnegative integer.");
}
function Ad(e, r) {
  if ((bQ(r), e === void 0)) return [];
  if (!Array.isArray(e) || e.length > x1.maxIntervalsPerAxis)
    throw Error("Hidden indices require at most 10,000 intervals per axis.");
  const t = e
      .map((o) => {
        if (
          !o ||
          typeof o != "object" ||
          Array.isArray(o) ||
          (Object.getPrototypeOf(o) !== Object.prototype &&
            Object.getPrototypeOf(o) !== null) ||
          Object.keys(o).some((a) => a !== "start" && a !== "end") ||
          !Number.isInteger(o.start) ||
          !Number.isInteger(o.end) ||
          o.start < 0 ||
          o.end < o.start ||
          o.end >= r
        )
          throw Error(
            "Hidden intervals require valid zero-based inclusive axis bounds.",
          );
        return { start: o.start, end: o.end };
      })
      .sort((o, a) => o.start - a.start || o.end - a.end),
    n = [];
  for (const o of t) {
    const a = n.at(-1);
    a && o.start <= a.end + 1 ? (a.end = Math.max(a.end, o.end)) : n.push(o);
  }
  return n;
}
function ul(e, r = []) {
  const t = Ad(r, e),
    n = [0];
  for (const l of t) n.push(n.at(-1) + l.end - l.start + 1);
  const o = e - n.at(-1),
    a = (l) => {
      let c = 0,
        h = t.length;
      for (; c < h;) {
        const f = (c + h) >>> 1;
        t[f].start <= l ? (c = f + 1) : (h = f);
      }
      return c - 1;
    },
    s = (l) => {
      const c = Math.max(0, Math.min(e, Math.floor(l))),
        h = a(c - 1),
        f =
          h < 0
            ? 0
            : n[h] + Math.min(c - t[h].start, t[h].end - t[h].start + 1);
      return c - f;
    },
    d = (l) => {
      const c = a(l);
      return c >= 0 && l <= t[c].end;
    },
    u = (l) => {
      if (!Number.isInteger(l) || l < 0 || l >= o) return -1;
      let c = 0,
        h = e - 1;
      for (; c < h;) {
        const f = Math.floor((c + h) / 2);
        s(f + 1) <= l ? (c = f + 1) : (h = f);
      }
      return c;
    };
  return {
    count: o,
    logicalCount: e,
    logicalIndex: u,
    countBefore: s,
    isHidden: d,
    viewIndex: (l) =>
      Number.isInteger(l) && l >= 0 && l < e && !d(l) ? s(l) : null,
    next: (l, c, h = false) => {
      const f = Math.trunc(l) + (h ? 0 : c);
      if (!o || (c === 1 && f >= e) || (c === -1 && f < 0)) return null;
      const p = c === 1 ? s(Math.max(0, f)) : s(Math.min(e, f + 1)) - 1,
        m = u(p);
      return m < 0 ? null : m;
    },
  };
}
function E1(e, r) {
  return (Z3(r), Ad(e[G3(r)], MN(e, r)));
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
const kB = { maxGroups: 1e3, maxDepth: 8 },
  SB = (e, r) => (r === "row" ? e.rowCount : e.colCount);
function F1(e) {
  if (e !== "row" && e !== "column") throw Error("Choose rows or columns.");
}
function Ow(e, r) {
  const t = e
      .filter((o) => o.axis === r)
      .sort((o, a) => o.start - a.start || a.end - o.end),
    n = [];
  return t.map((o) => {
    for (; n.length && n.at(-1).end < o.start;) n.pop();
    const a = n.at(-1);
    if (a && a.start === o.start && a.end === o.end)
      throw Error("An outline group already exists for these rows or columns.");
    if (a && a.end < o.end)
      throw Error(
        "Outline groups must be nested or separate; partially overlapping groups are not supported.",
      );
    if ((n.push(o), n.length > kB.maxDepth))
      throw Error("Outline groups support at most 8 nested groups.");
    return { ...o, depth: n.length };
  });
}
function TB(e, r, t = { rows: r.rowCount, cols: r.colCount }) {
  if (e.outlineGroups === void 0) return;
  if (!Array.isArray(e.outlineGroups) || e.outlineGroups.length > kB.maxGroups)
    throw Error("A worksheet supports at most 1,000 outline groups.");
  const n = new Set(),
    o = e.outlineGroups.map((a) => {
      if (
        !a ||
        typeof a != "object" ||
        Array.isArray(a) ||
        (Object.getPrototypeOf(a) !== Object.prototype &&
          Object.getPrototypeOf(a) !== null) ||
        Object.keys(a).some(
          (s) => !["id", "axis", "start", "end", "collapsed"].includes(s),
        ) ||
        typeof a.id != "string" ||
        !a.id.trim() ||
        a.id.length > 100 ||
        n.has(a.id) ||
        !["row", "column"].includes(a.axis) ||
        !Number.isSafeInteger(a.start) ||
        !Number.isSafeInteger(a.end) ||
        a.start < 0 ||
        a.end < a.start ||
        a.end >= (a.axis === "row" ? t.rows : t.cols) ||
        typeof a.collapsed != "boolean"
      )
        throw Error(
          "Outline groups require unique IDs, valid row or column bounds, and a boolean collapsed state.",
        );
      return (
        n.add(a.id),
        {
          id: a.id,
          axis: a.axis,
          start: a.start,
          end: a.end,
          collapsed: a.collapsed,
        }
      );
    });
  (Ow(o, "row"), Ow(o, "column"), (r.outlineGroups = o));
}
function NB(e, r) {
  return (F1(r), Ow(e.outlineGroups ?? [], r));
}
function cn(e, r) {
  F1(r);
  const t = SB(e, r),
    n = [
      ...Ad(r === "row" ? e.hiddenRows : e.hiddenCols, t),
      ...(e.outlineGroups ?? [])
        .filter((a) => a.axis === r && a.collapsed)
        .map(({ start: a, end: s }) => ({ start: a, end: s })),
    ].sort((a, s) => a.start - s.start || a.end - s.end),
    o = [];
  for (const a of n) {
    const s = o.at(-1);
    s && a.start <= s.end + 1
      ? (s.end = Math.max(s.end, a.end))
      : o.push({ ...a });
  }
  return Ad(o, t);
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
function RB(e) {
  var r;
  return (r = e.outlineGroups) == null
    ? void 0
    : r.map((t) => ({ ...t, id: crypto.randomUUID() }));
}
function TQ(e, r, t, n, o) {
  F1(r);
  const a = SB(e, r);
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
  if (!e.outlineGroups) return;
  const s = t + n - 1,
    d = e.outlineGroups.flatMap((u) => {
      if (u.axis !== r || u.end < t) return [{ ...u }];
      if (!o) {
        if (u.end + n >= a)
          throw Error(
            "Insertion would discard outline groups at the worksheet boundary. Ungroup those rows or columns first.",
          );
        return [
          { ...u, start: u.start >= t ? u.start + n : u.start, end: u.end + n },
        ];
      }
      return u.start > s
        ? [{ ...u, start: u.start - n, end: u.end - n }]
        : u.start >= t && u.end <= s
          ? []
          : [
              {
                ...u,
                start: Math.min(u.start, t),
                end: u.end > s ? u.end - n : t - 1,
              },
            ];
    });
  try {
    (Ow(d, "row"), Ow(d, "column"));
  } catch {
    throw Error(
      "This edit would make outline groups overlap or duplicate. Ungroup the affected rows or columns first.",
    );
  }
  return d;
}
const QS = (e) =>
  !!e &&
  typeof e == "object" &&
  !Array.isArray(e) &&
  [Object.prototype, null].includes(Object.getPrototypeOf(e));
function uf(e, r, t) {
  if (typeof e != "string" || !e.trim() || e.length > t)
    throw new Error(`${r} must contain 1–${t} characters.`);
  return e.trim();
}
function e5(e) {
  if (
    typeof e != "string" ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(e) ||
    !Number.isFinite(Date.parse(e)) ||
    new Date(e).toISOString() !== e
  )
    throw new Error("Workbook metadata requires a valid UTC timestamp.");
  return e;
}
function px(e, r) {
  if (e === void 0) return;
  if (
    !QS(e) ||
    Object.keys(e).some(
      (a) =>
        ![
          "createdAt",
          "modifiedAt",
          "engagement",
          "period",
          "template",
        ].includes(a),
    )
  )
    throw new Error("Invalid workbook metadata fields.");
  const t = e5(e.createdAt),
    n = e5(e.modifiedAt);
  if (n < t)
    throw new Error("Workbook modified time cannot precede its creation time.");
  const o = { createdAt: t, modifiedAt: n };
  if (
    (e.engagement !== void 0 &&
      (o.engagement = uf(e.engagement, "Engagement", 200)),
    e.period !== void 0 && (o.period = uf(e.period, "Period", 100)),
    e.template !== void 0)
  ) {
    const a = e.template;
    if (
      !QS(a) ||
      Object.keys(a).some(
        (s) =>
          ![
            "id",
            "name",
            "version",
            "sourceWorkbookId",
            "sourceRevision",
          ].includes(s),
      ) ||
      !Number.isSafeInteger(a.version) ||
      Number(a.version) < 1 ||
      (a.sourceRevision !== void 0 &&
        (!Number.isSafeInteger(a.sourceRevision) ||
          Number(a.sourceRevision) < 0))
    )
      throw new Error("Invalid template provenance.");
    o.template = {
      id: uf(a.id, "Template ID", 100),
      name: uf(a.name, "Template name", 200),
      version: Number(a.version),
      sourceWorkbookId: uf(
        a.sourceWorkbookId,
        "Template source workbook ID",
        100,
      ),
      ...(a.sourceRevision !== void 0
        ? { sourceRevision: Number(a.sourceRevision) }
        : {}),
    };
  }
  r.metadata = o;
}
function mx(e = {}, r = new Date().toISOString()) {
  const t = {};
  return (px({ createdAt: r, modifiedAt: r, ...e }, t), t.metadata);
}
function gx(e, r = new Date().toISOString()) {
  var o, a;
  const t = { ...e };
  e5(r);
  const n = [
    r,
    ((o = e.metadata) == null ? void 0 : o.createdAt) ?? r,
    ((a = e.metadata) == null ? void 0 : a.modifiedAt) ?? r,
  ]
    .sort()
    .at(-1);
  return (px(e.metadata ? { ...e.metadata, modifiedAt: n } : mx({}, r), t), t);
}
function Mw(e) {
  const r = rn(e);
  return (px(e.metadata, r), r);
}
function IB(e, r) {
  if (
    !QS(r) ||
    Object.keys(r).some((n) => !["name", "engagement", "period"].includes(n))
  )
    throw new Error("Invalid workbook property fields.");
  const t = Mw(e);
  if (
    (r.name !== void 0 && (t.name = uf(r.name, "Workbook name", 200)),
    r.engagement !== void 0 || r.period !== void 0)
  ) {
    if (typeof r.engagement == "string" && r.engagement.length > 200)
      throw new Error("Engagement must contain at most 200 characters.");
    if (typeof r.period == "string" && r.period.length > 100)
      throw new Error("Period must contain at most 100 characters.");
    const n = { ...(t.metadata ?? mx()) };
    (r.engagement !== void 0 &&
      (typeof r.engagement == "string" && !r.engagement.trim()
        ? delete n.engagement
        : (n.engagement = uf(r.engagement, "Engagement", 200))),
      r.period !== void 0 &&
        (typeof r.period == "string" && !r.period.trim()
          ? delete n.period
          : (n.period = uf(r.period, "Period", 100))),
      px(n, t));
  }
  return t;
}
function NQ(e, r) {
  return IB(e, { name: r });
}
function RQ(e) {
  const r = structuredClone(Mw(e)),
    t = new Map(r.sheets.map((a) => [a.id, crypto.randomUUID()])),
    n = new Map(
      r.sheets.flatMap((a) =>
        (a.tables ?? []).map((s) => [s.id, crypto.randomUUID()]),
      ),
    );
  ((r.id = crypto.randomUUID()), (r.revision = 0), (r.events = []));
  for (const a of r.sheets)
    if (a.analysisReport) {
      const s = a.analysisReport.request;
      if (s.kind === "statistics") s.sheetId = t.get(s.sheetId) ?? s.sheetId;
      else
        for (const d of [
          s.objective,
          ...s.variables.map((u) => u.cell),
          ...s.constraints.flatMap((u) => [
            u.left,
            ...(typeof u.right == "number" ? [] : [u.right]),
          ]),
        ])
          d.sheetId = t.get(d.sheetId) ?? d.sheetId;
      a.analysisReport.definitionHash = hr(s);
    }
  const o = [
    r.dataModel,
    ...r.sheets.map((a) => {
      var s;
      return (s = a.dataModelReport) == null ? void 0 : s.definition;
    }),
  ].filter(Boolean);
  for (const a of o)
    for (const s of a.tables)
      ((s.sheetId = t.get(s.sheetId) ?? s.sheetId),
        s.tableId && (s.tableId = n.get(s.tableId) ?? s.tableId));
  for (const a of r.sheets)
    a.dataModelReport &&
      (a.dataModelReport.definitionHash = hr(a.dataModelReport.definition));
  for (const a of r.savedFilterViews ?? [])
    ((a.id = crypto.randomUUID()), (a.sheetId = t.get(a.sheetId) ?? a.sheetId));
  for (const a of r.transformationRecipes ?? []) {
    a.id = crypto.randomUUID();
    for (const s of a.versions) {
      for (const d of [
        s.definition.source,
        ...s.definition.steps.flatMap((u) =>
          u.kind === "append" || u.kind === "merge" ? [u.source] : [],
        ),
      ])
        ((d.sheetId = t.get(d.sheetId) ?? d.sheetId),
          d.tableId && (d.tableId = n.get(d.tableId) ?? d.tableId));
      s.definitionHash = hr(s.definition);
    }
  }
  for (const a of r.sheets) {
    ((a.id = t.get(a.id)), a.outlineGroups && (a.outlineGroups = RB(a)));
    for (const s of a.tables ?? []) {
      s.id = n.get(s.id);
      for (const d of s.columns) d.id = crypto.randomUUID();
    }
    for (const s of [...(a.pivots ?? []), ...(a.queries ?? [])])
      ((s.id = crypto.randomUUID()),
        (s.source.sheetId = t.get(s.source.sheetId) ?? s.source.sheetId),
        s.source.tableId &&
          (s.source.tableId = n.get(s.source.tableId) ?? s.source.tableId),
        (s.sourceRevision = 0));
    for (const s of [
      ...(a.objects ?? []),
      ...(a.validations ?? []),
      ...(a.conditionalFormats ?? []),
    ])
      s.id = crypto.randomUUID();
  }
  for (const a of r.definedNames ?? [])
    a.scopeSheetId &&
      (a.scopeSheetId = t.get(a.scopeSheetId) ?? a.scopeSheetId);
  return Mw(cQ(hQ(tQ(r, crypto.randomUUID()), t), t));
}
function IQ(e, r) {
  var d, u;
  const t = uf(r.name, "Workbook name", 200),
    n = uf(r.engagement, "Engagement", 200),
    o = uf(r.period, "Period", 100),
    a = Mw(e),
    s = RQ(a);
  return (
    (s.name = t),
    (s.metadata = mx({
      engagement: n,
      period: o,
      template: {
        id: a.id,
        name: a.name,
        version:
          ((u = (d = a.metadata) == null ? void 0 : d.template) == null
            ? void 0
            : u.version) ?? 1,
        sourceWorkbookId: a.id,
        sourceRevision: a.revision,
      },
    })),
    Mw(s)
  );
}
function t5(e = "Book1") {
  return rn({
    schemaVersion: 1,
    id: crypto.randomUUID(),
    name: e.trim(),
    revision: 0,
    metadata: mx(),
    sheets: [
      {
        id: crypto.randomUUID(),
        name: "Sheet1",
        cells: {},
        rowCount: Mt.maxRows,
        colCount: Mt.maxCols,
        frozenRows: 0,
      },
    ],
    events: [],
  });
}
const f6 = /[A-Za-z0-9_.\\\u0080-\uffff]/,
  _Q =
    /^(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*)(?::(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*))?!/,
  DQ = (e) =>
    [...e.toUpperCase()].reduce((r, t) => r * 26 + t.charCodeAt(0) - 64, 0) - 1;
function h6(e, r, t) {
  const n = DQ(r) + (e ? 0 : t);
  return n < 0 || n >= Mt.maxCols ? "#REF!" : e + ut(n);
}
function p6(e, r, t) {
  const n = Number(r) + (e ? 0 : t);
  return n < 1 || n > Mt.maxRows ? "#REF!" : e + n;
}
function xd(e, r, t) {
  if (!Number.isSafeInteger(r) || !Number.isSafeInteger(t))
    throw new Error("Formula offsets must be integers.");
  if (!e.startsWith("=")) return e;
  let n = "",
    o = 0;
  for (; o < e.length;) {
    const a = e.slice(o),
      s = e[o],
      d = (o === 0 || !f6.test(e[o - 1])) && a.match(_Q);
    if (d) {
      ((n += d[0]), (o += d[0].length));
      continue;
    }
    if (s === '"' || s === "'") {
      const p = o++;
      for (; o < e.length;)
        if (e[o++] === s)
          if (e[o] === s) o++;
          else break;
      n += e.slice(p, o);
      continue;
    }
    if (s === "[") {
      const p = o++;
      let m = 1;
      for (; o < e.length && m;)
        (e[o] === "[" && m++, e[o] === "]" && m--, o++);
      n += e.slice(p, o);
      continue;
    }
    const u = o === 0 || !f6.test(e[o - 1]),
      l = u ? /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/.exec(a) : null,
      c = u ? /^(\$?)([1-9]\d*):(\$?)([1-9]\d*)/.exec(a) : null,
      h = u ? /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/.exec(a) : null,
      f = l ?? c ?? h;
    if (
      f &&
      !f6.test(a[f[0].length] ?? "") &&
      !/^\s*[(!]/.test(a.slice(f[0].length))
    ) {
      if (l) n += `${h6(l[1], l[2], t)}:${h6(l[3], l[4], t)}`;
      else if (c) n += `${p6(c[1], c[2], r)}:${p6(c[3], c[4], r)}`;
      else if (h) {
        const p = h6(h[1], h[2], t),
          m = p6(h[3], h[4], r);
        n += p === "#REF!" || m === "#REF!" ? "#REF!" : p + m;
      }
      o += f[0].length;
      continue;
    }
    ((n += s), o++);
  }
  return n;
}
function OQ(e, r) {
  const t = Ix(r);
  if (
    t.some((n) => {
      const { row: o, col: a } = _t(n);
      return o >= e.rowCount || a >= e.colCount;
    })
  )
    throw new Error("Selection exceeds worksheet dimensions.");
  return t;
}
const MQ = (e) =>
  e ? { ...e, ...(e.style ? { style: { ...e.style } } : {}) } : { raw: null };
function jQ(e, r, t) {
  const n = Nr(e, r),
    o = OQ(n, t);
  if (o.length !== 1) throw new Error("Select one empty total cell.");
  t = o[0];
  const a = _t(t),
    s = n.cells[t];
  if ((s == null ? void 0 : s.raw) !== void 0 && s.raw !== null && s.raw !== "")
    throw new Error("Select an empty cell for AutoSum.");
  const d = (h, f) => {
    if (h < 0 || f < 0) return false;
    const p = n.cells[`${ut(f)}${h + 1}`];
    return (
      typeof (p == null ? void 0 : p.raw) == "number" ||
      (typeof (p == null ? void 0 : p.raw) == "string" &&
        p.raw.startsWith("=") &&
        !p.literal)
    );
  };
  let u, l;
  if (d(a.row - 1, a.col)) {
    let h = a.row - 1;
    for (; d(h - 1, a.col);) h--;
    ((u = `${ut(a.col)}${h + 1}`), (l = `${ut(a.col)}${a.row}`));
  } else if (d(a.row, a.col - 1)) {
    let h = a.col - 1;
    for (; d(a.row, h - 1);) h--;
    ((u = `${ut(h)}${a.row + 1}`), (l = `${ut(a.col - 1)}${a.row + 1}`));
  } else
    throw new Error(
      "AutoSum needs adjacent numbers or formulas above or to the left.",
    );
  const c = MQ(s);
  return (
    (c.raw = `=SUM(${u}:${l})`),
    delete c.literal,
    { sheetId: r, address: t, cell: c }
  );
}
const a1 = Object.freeze({
    rulesPerSheet: 100,
    listItems: 200,
    checks: 1e5,
    customChecks: 5e3,
    displayedIssues: 500,
  }),
  FQ = (e) => !!e && typeof e == "object" && !Array.isArray(e),
  yv = (e, r, t) => {
    if (typeof e != "string" || e.length > r)
      throw Error(`Invalid validation ${t}.`);
    return e;
  };
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
function jw(e, r, t = TN) {
  if (r === null || r === "") return e.allowBlank;
  if (e.kind === "custom") return false;
  if (e.kind === "list")
    return e.list.some((o) => o.toLowerCase() === String(r).toLowerCase());
  let n =
    e.kind === "textLength" ? String(r).length : typeof r == "number" ? r : NaN;
  if (
    (["date", "time"].includes(e.kind) &&
      typeof r == "string" &&
      (n = zp(r, t) ?? NaN),
    !Number.isFinite(n) ||
      (e.kind === "integer" && !Number.isInteger(n)) ||
      (e.kind === "date" &&
        (n < 1 ||
          n > Cy.maxSerial + 0.999999999 ||
          (typeof r == "string" && !/[/-]/.test(r)))) ||
      (e.kind === "time" && (n < 0 || n >= 1)))
  )
    return false;
  switch (e.operator) {
    case "between":
      return n >= e.min && n <= e.max;
    case "notBetween":
      return n < e.min || n > e.max;
    case "greater":
      return n > e.min;
    case "greaterEqual":
      return n >= e.min;
    case "less":
      return n < e.min;
    case "lessEqual":
      return n <= e.min;
    case "notEqual":
      return n !== e.min;
    default:
      return n === e.min;
  }
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
const MB = new WeakMap(),
  r5 = new WeakMap();
class jN extends Error {
  constructor(t, n) {
    super(
      n
        .slice(0, 5)
        .map((o) => `${o.sheetName}!${o.address}: ${o.message} ${o.reason}`)
        .join(`
`) +
        (n.length > 5
          ? `
${n.length} validation violations in this batch.`
          : ""),
    );
    yc(this, "issues");
    ((this.name = "DataValidationError"),
      (this.issues = structuredClone(n)),
      MB.set(this, JSON.stringify({ book: t, issues: n })));
  }
}
function jB(e) {
  const r = MB.get(e);
  if (
    !r ||
    !e.issues.length ||
    e.issues.some((n) => n.style === "stop" || n.evaluationError) ||
    JSON.stringify(JSON.parse(r).issues) !== JSON.stringify(e.issues)
  )
    throw Error(
      "Stop errors cannot be overridden. Correct the invalid values.",
    );
  const t = Object.freeze({ kind: "validation-warning-approval" });
  return (r5.set(t, r), t);
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
const FB = new WeakMap();
function BQ(e, r) {
  const t = FB.get(r);
  if (!t || t.source !== JSON.stringify(e) || t.content !== JSON.stringify(r))
    throw Error(
      "Validation inspection is stale or unknown. Inspect existing values again.",
    );
  return JSON.stringify(
    { schemaVersion: 1, kind: "local-data-validation-inspection", report: r },
    null,
    2,
  );
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
function vx(e, r) {
  if (!r.getSpillEntries) return e.cells;
  const t = r.getSpillEntries(e.id);
  if (!t.length) return e.cells;
  const n = { ...e.cells };
  for (const { address: o } of t) n[o] = { ...e.cells[o], raw: r(e.id, o) };
  return n;
}
const BB = [
    "greater",
    "less",
    "equal",
    "between",
    "contains",
    "duplicate",
    "unique",
    "blank",
    "notBlank",
    "error",
    "notError",
    "formula",
    "colorScale",
    "dataBar",
    "iconSet",
  ],
  PQ = (e) => !!e && typeof e == "object" && !Array.isArray(e);
function Ub(e, r, t) {
  if (typeof e != "string" || e.length > r) throw Error(`Invalid ${t}.`);
  return e;
}
function m6(e) {
  if (typeof e != "string" || !/^#[\da-f]{6}$/i.test(e))
    throw Error("Choose a six-digit hexadecimal color.");
  return e;
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
const n5 = (e, r) =>
    r.row >= e.firstRow &&
    r.row <= e.lastRow &&
    r.col >= e.firstCol &&
    r.col <= e.lastCol,
  UQ = (e) =>
    typeof e == "string" &&
    /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|ERROR!|CYCLE!|SPILL!|CALC!|LIMIT!|BLOCKED!)/.test(
      e,
    );
function J3(e) {
  return e
    .map((r, t) => ({ rule: r, rank: r.priority ?? 101 + e.length - t }))
    .sort((r, t) => r.rank - t.rank)
    .map((r) => r.rule);
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
function $Q(e, r, t = {}) {
  if (
    r.workbookId !== e.id ||
    r.revision !== e.revision ||
    r.snapshotGuard !== JSON.stringify(e)
  )
    throw Error(
      "Conditional format preview is stale. Reopen the manager to refresh it.",
    );
  const n = LB(e, r.sheetId, r.rules, t);
  if (JSON.stringify(n.changes) !== JSON.stringify(r.changes))
    throw Error(
      "Conditional format preview was modified. Preview the rules again.",
    );
  return n.changes.length
    ? rn({
        ...e,
        sheets: e.sheets.map((o) =>
          o.id === r.sheetId ? { ...o, conditionalFormats: n.rules } : o,
        ),
      })
    : e;
}
const sD = (e) => !!e && typeof e == "object" && !Array.isArray(e);
function al(e) {
  throw Error(e);
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
function jn(e) {
  var r;
  (r = e.protection) != null &&
    r.enabled &&
    al("This sheet is protected. Unprotect it before changing its contents.");
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
function WQ(e, r, t, n, o = true) {
  var p;
  const a = e.sheets.find((m) => m.id === r);
  jn(a);
  const s = Ix(t);
  s.length > 500 && al("Select at most 500 cells to remove duplicates.");
  const d = ct(t);
  (!n.length || n.some((m) => m < d.firstCol || m > d.lastCol)) &&
    al("Choose one or more columns inside the selected range.");
  for (const m of a.tables ?? []) {
    const v = ct(m.range);
    AE(t, m.range) &&
      ((d.firstCol > v.firstCol || d.lastCol < v.lastCol) &&
        al("Include all table columns when removing duplicate table rows."),
      m.showTotals &&
        d.firstRow <= v.lastRow &&
        d.lastRow >= v.lastRow &&
        al("Select data rows above the table totals row."),
      !o &&
        d.firstRow <= v.firstRow &&
        d.lastRow >= v.firstRow &&
        al("Keep the table header row selected as headers."));
  }
  (a.merges ?? []).some((m) => AE(t, m)) &&
    al("Unmerge selected cells before removing duplicates.");
  const u = Vn(e),
    l = new Set(),
    c = [];
  let h = 0;
  for (let m = d.firstRow; m <= d.lastRow; m++) {
    const v = JSON.stringify(n.map((g) => u(r, `${ut(g)}${m + 1}`)));
    if (m !== d.firstRow || !o) {
      if (l.has(v)) {
        h++;
        continue;
      }
      l.add(v);
    }
    c.push(
      Array.from({ length: d.lastCol - d.firstCol + 1 }, (g, b) => {
        const y = structuredClone(
          a.cells[`${ut(d.firstCol + b)}${m + 1}`] ?? { raw: null },
        );
        return (
          typeof y.raw == "string" &&
            y.raw.startsWith("=") &&
            !y.literal &&
            (y.raw = xd(y.raw, d.firstRow + c.length - m, 0)),
          y
        );
      }),
    );
  }
  const f = { ...a.cells };
  for (const m of s) {
    const v = _t(m);
    f[m] = ((p = c[v.row - d.firstRow]) == null
      ? void 0
      : p[v.col - d.firstCol]) ?? { ...f[m], raw: null };
  }
  return {
    workbook: {
      ...e,
      sheets: e.sheets.map((m) => (m.id === r ? { ...m, cells: f } : m)),
    },
    removed: h,
  };
}
async function uD(e, r) {
  const t = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(r + "\0" + e),
  );
  return [...new Uint8Array(t)]
    .map((n) => n.toString(16).padStart(2, "0"))
    .join("");
}
async function HQ(e, r, t, n) {
  var d, u;
  const o = e.sheets.find((l) => l.id === r);
  (n &&
    (d = o.protection) != null &&
    d.enabled &&
    al("Unprotect this sheet before changing its protection settings."),
    t.length > 200 && al("Password is too long."),
    !n &&
      (u = o.protection) != null &&
      u.hash &&
      (await uD(t, o.protection.salt)) !== o.protection.hash &&
      al("The password does not match."));
  const a = crypto.randomUUID(),
    s = n
      ? { enabled: true, ...(t ? { salt: a, hash: await uD(t, a) } : {}) }
      : { enabled: false };
  return {
    ...e,
    sheets: e.sheets.map((l) => (l.id === r ? { ...l, protection: s } : l)),
  };
}
function qQ(e, r, t, n) {
  var u;
  const o = e.sheets.find((l) => l.id === r);
  jn(o);
  const a = Ix(n);
  a.length > 500 && al("Paint at most 500 cells.");
  const s = o.cells[t] ?? {},
    d = { ...o.cells };
  for (const l of a)
    d[l] = hi({
      ...d[l],
      raw: ((u = d[l]) == null ? void 0 : u.raw) ?? null,
      style: structuredClone(s.style ?? {}),
      format: s.format ?? "general",
      numFmt: s.numFmt,
      numberFormat: s.numberFormat ? structuredClone(s.numberFormat) : void 0,
    });
  return {
    ...e,
    sheets: e.sheets.map((l) => (l.id === r ? { ...l, cells: d } : l)),
  };
}
const PB = ["en-US", "en-GB", "de-DE", "fr-FR", "ja-JP", "en-IN"],
  UB = ["USD", "EUR", "GBP", "JPY", "CAD", "AUD", "INR", "CHF"],
  Uv = [
    "general",
    "number",
    "currency",
    "accounting",
    "percent",
    "date",
    "time",
    "datetime",
    "duration",
    "custom",
  ],
  Ay = {
    date: ["yyyy-mm-dd", "dd/mm/yyyy", "m/d/yyyy", "d mmm yyyy", "d mmmm yyyy"],
    time: ["hh:mm", "hh:mm:ss", "h:mm AM/PM", "hh:mm:ss.000"],
    datetime: [
      "yyyy-mm-dd hh:mm:ss",
      "dd/mm/yyyy hh:mm",
      "m/d/yyyy h:mm AM/PM",
    ],
    duration: ["[h]:mm", "[h]:mm:ss", "[h]:mm:ss.000", "[m]:ss", "[s].000"],
  };
function Si(e) {
  throw new Error(e);
}
const XQ = new Set(["number", "currency", "accounting", "percent"]),
  GQ = new Set(["date", "time", "datetime", "duration"]);
function ZQ(e) {
  return (
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e))
  );
}
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
function BN(e) {
  const r = [""];
  let t = false;
  for (let n = 0; n < e.length; n++) {
    const o = e[n];
    if (o === "\\" && n + 1 < e.length) {
      r[r.length - 1] += o + e[++n];
      continue;
    }
    (o === '"' && (t = !t),
      o === ";" && !t ? r.push("") : (r[r.length - 1] += o));
  }
  return t || r.length > 4 ? void 0 : r;
}
function LN(e) {
  const r = {
      prefix: "",
      suffix: "",
      minDecimals: 0,
      maxDecimals: 0,
      minIntegerDigits: 1,
      grouping: false,
      percent: false,
      hasNumber: false,
    },
    t = (n) => {
      r.hasNumber ? (r.suffix += n) : (r.prefix += n);
    };
  for (let n = 0; n < e.length;) {
    const o = e[n];
    if (o === '"') {
      const a = e.indexOf('"', n + 1);
      if (a < 0) return;
      (t(e.slice(n + 1, a)), (n = a + 1));
    } else if (o === "\\") {
      if (n + 1 >= e.length) return;
      (t(e[n + 1]), (n += 2));
    } else if (o === "[") {
      const a =
        /^\[(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\d+)\]/i.exec(
          e.slice(n),
        );
      if (!a) return;
      n += a[0].length;
    } else if (o === "0" || o === "#") {
      if (r.hasNumber) return;
      const a = /^[0#.,]+/.exec(e.slice(n))[0],
        s = /^(0+|#+0*|#,##0+)(?:\.(0*#*))?$/.exec(a);
      if (!s) return;
      const d = s[2] ?? "";
      if (
        ((r.minDecimals = [...d].filter((u) => u === "0").length),
        (r.maxDecimals = d.length),
        (r.minIntegerDigits = [...s[1]].filter((u) => u === "0").length),
        r.maxDecimals > 20 || r.minIntegerDigits > 21)
      )
        return;
      ((r.grouping = s[1].includes(",")),
        (r.hasNumber = true),
        (n += a.length));
    } else if (o === "%") {
      if (r.percent) return;
      ((r.percent = true), t(o), n++);
    } else if (/[$€£¥ ()+\-]/.test(o)) (t(o), n++);
    else return;
  }
  return r;
}
function $B(e) {
  if (!e) return;
  const r = BN(e);
  if (!r) return;
  const t = LN(r[0]);
  return t != null && t.hasNumber ? t.maxDecimals : void 0;
}
function zB(e, r) {
  const t = BN(r);
  if (!t) return;
  const n = e < 0 && t.length > 1,
    o = e === 0 && t.length > 2 ? 2 : n ? 1 : 0,
    a = LN(t[o]);
  if (!a) return;
  if (!a.hasNumber) return a.prefix + a.suffix;
  const s = Math.abs(e) * (a.percent ? 100 : 1);
  if (!Number.isFinite(s)) return "#NUM!";
  let d = new Intl.NumberFormat("en-US", {
    minimumFractionDigits: a.minDecimals,
    maximumFractionDigits: a.maxDecimals,
    minimumIntegerDigits: Math.max(1, a.minIntegerDigits),
    useGrouping: a.grouping,
  }).format(s);
  return (
    a.minIntegerDigits === 0 &&
      (d === "0" && a.minDecimals === 0
        ? (d = "")
        : d.startsWith("0.") && (d = d.slice(1))),
    (e < 0 && !n ? "-" : "") + a.prefix + d + a.suffix
  );
}
const $b = new Map();
function KQ(e) {
  const r = e.replace(/"[^"]*"|\\./g, "");
  if (!/[ymdhs]/i.test(r)) return;
  const t = [];
  let n = 0;
  for (let u = 0; u < e.length;) {
    const l = e.slice(u),
      c = l[0];
    if (c === '"') {
      const v = e.indexOf('"', u + 1);
      (v < 0 && Si("The format contains an unclosed quoted literal."),
        t.push({ kind: "literal", text: e.slice(u + 1, v) }),
        (u = v + 1));
      continue;
    }
    if (c === "\\") {
      (l.length < 2 && Si("A format escape must be followed by a character."),
        t.push({ kind: "literal", text: l[1] }),
        (u += 2));
      continue;
    }
    const h = /^(AM\/PM|A\/P)/i.exec(l);
    if (h) {
      (t.push({ kind: "ampm", text: h[0] }), (u += h[0].length));
      continue;
    }
    const f = /^\[(h{1,2}|m{1,2}|s{1,2})\]/i.exec(l);
    if (f) {
      (t.push({ kind: "elapsed", text: f[1].toLowerCase() }),
        (u += f[0].length));
      continue;
    }
    const p = /^(y+|m+|d+|h+|s+)/i.exec(l);
    if (p) {
      const v = p[0].toLowerCase(),
        g = v[0];
      (((g === "y" && ![2, 4].includes(v.length)) ||
        (g === "m" && v.length > 4) ||
        (["d", "h", "s"].includes(g) && v.length > 2)) &&
        Si(
          "This date pattern uses unsupported fields. Use yy/yyyy, m–mmmm, d/dd, h/hh or s/ss.",
        ),
        t.push({
          kind: { y: "year", m: "month", d: "day", h: "hour", s: "second" }[g],
          text: v,
        }),
        (u += v.length));
      continue;
    }
    const m = /^\.(0{1,3})(?!0)/.exec(l);
    if (m) {
      ((!t.length ||
        !["second", "elapsed"].includes(t.at(-1).kind) ||
        (t.at(-1).kind === "elapsed" && !t.at(-1).text.startsWith("s")) ||
        n) &&
        Si("Fractional seconds must follow seconds and use 1–3 zero digits."),
        (n = m[1].length),
        t.push({ kind: "fraction", text: m[1] }),
        (u += m[0].length));
      continue;
    }
    if (/^[ /:,.-]$/.test(c)) {
      (t.push({ kind: "literal", text: c }), u++);
      continue;
    }
    Si(
      "Unsupported date/time code. Quote literal words and use one supported date/time pattern.",
    );
  }
  const o = t.filter((u) => u.kind !== "literal");
  o.forEach((u, l) => {
    var c, h, f;
    u.kind === "month" &&
      u.text.length <= 2 &&
      (((c = o[l - 1]) == null ? void 0 : c.kind) === "hour" ||
        (((h = o[l - 1]) == null ? void 0 : h.kind) === "elapsed" &&
          o[l - 1].text.startsWith("h")) ||
        ((f = o[l + 1]) == null ? void 0 : f.kind) === "second") &&
      (u.kind = "minute");
  });
  const a = o.some((u) => ["year", "month", "day"].includes(u.kind)),
    s = o.some((u) => ["hour", "minute", "second"].includes(u.kind)),
    d = o.filter((u) => u.kind === "elapsed");
  if (d.length) {
    (d.length !== 1 ||
      a ||
      o.some((l) => ["hour", "ampm"].includes(l.kind)) ||
      o[0].kind !== "elapsed") &&
      Si(
        "Elapsed-time patterns start with one [h], [m] or [s] field and cannot include calendar or AM/PM fields.",
      );
    const u = d[0].text[0];
    ((u === "s" && s) || (u === "m" && o.some((l) => l.kind === "minute"))) &&
      Si("Use [h]:mm:ss, [m]:ss or [s] for elapsed time.");
  }
  return (
    o.some((u) => u.kind === "ampm") &&
      !o.some((u) => u.kind === "hour") &&
      Si("AM/PM requires an hour field."),
    !a &&
      !s &&
      !d.length &&
      Si("The format does not contain a date or time field."),
    {
      kind: "temporal",
      category: d.length
        ? "duration"
        : a && s
          ? "datetime"
          : a
            ? "date"
            : "time",
      tokens: t,
      fraction: n,
    }
  );
}
function YQ(e) {
  var a;
  let r = "",
    t = 1,
    n = false;
  for (let s = 0; s < e.length;) {
    const d = e.slice(s);
    if (d[0] === '"') {
      const u = e.indexOf('"', s + 1);
      (u < 0 && Si("The format contains an unclosed quoted literal."),
        (r += e.slice(s, u + 1)),
        (s = u + 1));
    } else if (d[0] === "\\")
      (d.length < 2 && Si("A format escape must be followed by a character."),
        (r += d.slice(0, 2)),
        (s += 2));
    else if (/^[0#]/.test(d)) {
      n && Si("Only one numeric field per custom section is supported.");
      const u = /^[0#.,]+/.exec(d)[0],
        l = ((a = /,+$/.exec(u)) == null ? void 0 : a[0]) ?? "";
      (l.length > 2 &&
        Si("Custom display scaling supports at most two trailing commas."),
        (t = 1e3 ** l.length),
        (r += u.slice(0, u.length - l.length)),
        (s += u.length),
        (n = true));
    } else ((r += d[0]), s++);
  }
  /\[/.test(r.replace(/"[^"]*"|\\./g, "")) &&
    Si(
      "Custom colors, conditions and locale tags are not supported. Use conditional formatting or the locale selector.",
    );
  const o = LN(r);
  return (
    o ||
      Si(
        "Unsupported custom format. Use 0/# digits, grouping, decimals, percent, quoted literals and positive/negative/zero sections.",
      ),
    o.maxDecimals > 12 &&
      Si("Custom formats support at most 12 decimal places."),
    { parsed: o, divisor: t }
  );
}
function Bw(e) {
  const r = $b.get(e);
  if (r) return r;
  (typeof e != "string" ||
    !e.length ||
    e.length > 200 ||
    /[\u0000-\u001f\u007f]/.test(e)) &&
    Si("Enter a supported format pattern of 1–200 characters.");
  let t;
  const n = BN(e);
  n || Si("A format has at most four sections and must close quoted literals.");
  const o = n.length === 1 ? KQ(e) : void 0;
  return (
    o
      ? (t = o)
      : (n.length === 4 &&
          !JQ(n[3]) &&
          Si(
            "The text section supports @ and quoted or escaped literals only.",
          ),
        (t = {
          kind: "numeric",
          sections: n.slice(0, 3).map(YQ),
          ...(n.length === 4 ? { textSection: n[3] } : {}),
        })),
    $b.size >= 256 && $b.delete($b.keys().next().value),
    $b.set(e, t),
    t
  );
}
function JQ(e) {
  return e.replace(/"[^"]*"|\\./g, "").replace(/[@ ]/g, "") === "";
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
function nee(e, r, t, n, o = {}) {
  o.readOnly && Si("This workbook is read-only.");
  const a = Nr(e, r);
  jn(a);
  const s = l0(n),
    d = Np(typeof t == "string" ? [t] : t),
    u = Q3(s),
    l = d.flatMap((h) => {
      const f = a.cells[h],
        p = {
          ...(f ?? { raw: null }),
          numberFormat: s,
          numFmt: u,
          format:
            s.category === "percent"
              ? "percent"
              : s.category === "currency" || s.category === "accounting"
                ? "currency"
                : s.category === "date" || s.category === "datetime"
                  ? "date"
                  : s.category === "number"
                    ? "number"
                    : "general",
        };
      return JSON.stringify(f) === JSON.stringify(p)
        ? []
        : [{ sheetId: r, address: h, cell: p }];
    });
  if (!l.length) return e;
  const c = p0(e, l, "Preparer", "Format numbers");
  return { ...e, sheets: c.sheets };
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
function dp(e, r, t) {
  var o, a, s;
  if (e.cell || (e.call && !e.qualifier && Fc.includes(P3(e.name)))) return;
  const n = e.qualifier
    ? (o = r.sheets.find(
        (d) => d.name.toLowerCase() === e.qualifier.toLowerCase(),
      )) == null
      ? void 0
      : o.id
    : t;
  if (n)
    return (
      ((a = r.definedNames) == null
        ? void 0
        : a.find(
            (d) =>
              d.scopeSheetId === n &&
              d.name.toLowerCase() === e.name.toLowerCase(),
          )) ??
      ((s = r.definedNames) == null
        ? void 0
        : s.find(
            (d) =>
              !d.scopeSheetId && d.name.toLowerCase() === e.name.toLowerCase(),
          ))
    );
}
function g5(e, r, t) {
  let n = "",
    o = 0;
  for (const a of r) ((n += e.slice(o, a.start) + t(a)), (o = a.end));
  return n + e.slice(o);
}
function pf(e, r, t, n = [], o) {
  if (n.length > 32) throw new Error("#LIMIT!");
  const a = g5(e, m5(e), (s) => {
    if (s.call) return s.text;
    const d = dp(s, r, t);
    if (!d || $p(d.refersTo)) return s.text;
    const u = Ba(d);
    if (n.includes(u)) throw new Error("#CYCLE!");
    const l =
      d.scopeSheetId ??
      (s.qualifier
        ? r.sheets.find(
            (h) => h.name.toLowerCase() === s.qualifier.toLowerCase(),
          ).id
        : t);
    if (o && !o(l)) throw new Error("#BLOCKED!");
    let c = pf(d.refersTo.slice(1), r, l, [...n, u], o);
    if (l !== t) {
      const h = r.sheets.find((p) => p.id === l),
        f = m5(c);
      c = g5(c, f, (p) => {
        const m = f.indexOf(p),
          v = f[m - 1],
          g =
            (v == null ? void 0 : v.cell) &&
            c.slice(v.end, p.start).trim() === ":";
        return p.cell && !p.qualifier && !g ? sne(h.name) + p.text : p.text;
      });
    }
    return `(${c})`;
  });
  if (a.length > 16384) throw new Error("#LIMIT!");
  return a;
}
const pO = 500;
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
        ? void 0
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
      var w;
      return (w = y.tables) == null
        ? void 0
        : w.some((E) => E.name.toLowerCase() === a.name.toLowerCase());
    })
  )
    throw Error("A table already uses this name.");
  const s = structuredClone(e),
    d = o ? Ba(o) : void 0;
  ((s.definedNames = eg(
    [...n.filter((y) => Ba(y) !== d), ...(a ? [a] : [])],
    e,
  )),
    o && a && (s.definedNames = n.map((y) => (Ba(y) === d ? a : y))));
  const u =
      (a == null ? void 0 : a.scopeSheetId) ??
      (o == null ? void 0 : o.scopeSheetId),
    l = (y) => {
      var w, E;
      if (
        (E =
          (w = e.sheets.find((A) => A.id === y)) == null
            ? void 0
            : w.protection) != null &&
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
          F = r.kind === "create" ? dp(A, s, w) : void 0;
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
            return (O[D] ? Ba(O[D]) : void 0) !== (B ? Ba(B) : void 0);
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
function dne(e, r) {
  const t = eg(r.definedNames, r);
  oP(t);
  const { builtinXml: n } = aP(e);
  let o = n;
  for (const a of t) {
    if (a.name.toLowerCase().startsWith("_xlnm."))
      return Pi(
        `Reserved built-in name "${a.name}" cannot be supplied as a user name.`,
      );
    (pC(a.refersTo), cne(a.refersTo, r));
    const s = a.scopeSheetId
      ? r.sheets.findIndex((d) => d.id === a.scopeSheetId)
      : void 0;
    o += `<definedName name="${Yy(a.name, true)}"${s !== void 0 ? ` localSheetId="${s}"` : ""}${a.comment !== void 0 ? ` comment="${Yy(a.comment, true)}"` : ""}>${Yy(a.refersTo.slice(1))}</definedName>`;
  }
  return fne(e, o);
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
function w0(e) {
  const r = rn(aQ(e));
  return (
    r.sheets.forEach((t, n) => wP(e.sheets[n], t)),
    nR(e.structureProtection, r),
    r
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
function CP(e, r) {
  return ($a(e), w0({ ...e, sheets: [...e.sheets, Ku(rR(e, r))] }));
}
function Pne(e, r, t) {
  $a(e);
  const n = qp(e, r),
    o = rR(e, t, r);
  if (n.name === o) return e;
  const a = structuredClone(e);
  for (const s of a.sheets)
    for (const d of Object.values(s.cells))
      !d.literal &&
        typeof d.raw == "string" &&
        d.raw.startsWith("=") &&
        (d.raw = e0(d.raw, e, n, o));
  for (const s of a.sheets)
    for (const d of [...(s.conditionalFormats ?? []), ...(s.validations ?? [])])
      if ((d.kind === "formula" || d.kind === "custom") && d.formula) {
        const u = e0(d.formula, e, n, o);
        (u !== d.formula && jn(s), (d.formula = u));
      }
  for (const s of a.definedNames ?? []) s.refersTo = e0(s.refersTo, e, n, o);
  for (const s of a.sheets)
    for (const d of s.tables ?? [])
      for (const u of d.columns)
        if (u.calculatedFormula) {
          const l = e0(u.calculatedFormula, e, n, o);
          (l !== u.calculatedFormula && jn(s), (u.calculatedFormula = l));
        }
  return ((qp(a, r).name = o), w0(a));
}
function xO(e, r, t) {
  $a(e);
  const n = qp(e, r),
    o = e.sheets.indexOf(n);
  if (
    ((!Number.isInteger(t) || t < 0 || t >= e.sheets.length) &&
      ea("Choose a worksheet position inside the workbook."),
    o === t)
  )
    return e;
  EP(e, (s) => {
    vC(s).some((d) => d.sheets) &&
      ea(
        "Reordering worksheets with three-dimensional references requires explicit references first.",
      );
  });
  const a = [...e.sheets];
  return (a.splice(o, 1), a.splice(t, 0, n), w0({ ...e, sheets: a }));
}
function Une(e, r, t) {
  $a(e);
  const n = qp(e, r);
  (["visible", "hidden", "veryHidden"].includes(t) ||
    ea("Choose visible, hidden or very hidden."),
    t !== "visible" &&
      zw(n) &&
      e.sheets.filter(zw).length === 1 &&
      ea("At least one worksheet must remain visible."));
  const o = n.veryHidden ? "veryHidden" : n.hidden ? "hidden" : "visible";
  return t === o
    ? e
    : w0({
        ...e,
        sheets: e.sheets.map((a) =>
          a.id === r
            ? { ...a, hidden: t !== "visible", veryHidden: t === "veryHidden" }
            : a,
        ),
      });
}
function EO(e, r, t) {
  $a(e);
  const n = qp(e, r),
    o = t === void 0 ? void 0 : yP(t);
  return n.tabColor === o
    ? e
    : w0({
        ...e,
        sheets: e.sheets.map((a) => {
          if (a.id !== r) return a;
          const s = { ...a };
          return (o ? (s.tabColor = o) : delete s.tabColor, s);
        }),
      });
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
function zne(e, r, t) {
  var p, m, v;
  $a(e);
  const n = qp(e, r),
    o = rR(e, $ne(e, n.name));
  (EP(e, (g) => {
    vC(g).some((b) => b.sheets) &&
      ea(
        "Duplicating worksheets with three-dimensional references requires explicit references first.",
      );
  }),
    e.sheets.reduce((g, b) => g + Object.keys(b.cells).length, 0) +
      Object.keys(n.cells).length >
      Mt.maxCells &&
      ea(
        "Duplicating this worksheet would exceed the 100,000 populated-cell budget.",
      ));
  const s = structuredClone(n);
  ((s.id = fp()), s.outlineGroups && (s.outlineGroups = RB(s)), (s.name = o));
  const d = new Set([
      ...e.sheets.flatMap((g) =>
        (g.tables ?? []).map((b) => b.name.toLowerCase()),
      ),
      ...(e.definedNames ?? []).map((g) => g.name.toLowerCase()),
    ]),
    u = new Map(),
    l = new Map();
  for (const g of s.tables ?? []) {
    const b = g.id,
      y = g.name;
    ((g.id = fp()), (g.columns = g.columns.map((w) => ({ ...w, id: fp() }))));
    for (let w = 2; ; w++) {
      const E = `_Copy${w}`,
        A = y.slice(0, 255 - E.length) + E;
      if (!d.has(A.toLowerCase())) {
        ((g.name = A), d.add(A.toLowerCase()));
        break;
      }
    }
    (u.set(b, g.id), l.set(y, g.name));
  }
  const c = (g) => {
    let b = e0(g, e, n, o);
    for (const [y, w] of l) b = $v(b, y, w);
    return b;
  };
  for (const g of s.tables ?? [])
    for (const b of g.columns)
      b.calculatedFormula && (b.calculatedFormula = c(b.calculatedFormula));
  for (const g of Object.values(s.cells))
    !g.literal &&
      typeof g.raw == "string" &&
      g.raw.startsWith("=") &&
      (g.raw = c(g.raw));
  ((p = s.objects) == null ||
    p.forEach((g) => {
      g.id = fp();
    }),
    (m = s.validations) == null ||
      m.forEach((g) => {
        ((g.id = fp()),
          g.kind === "custom" && g.formula && (g.formula = c(g.formula)));
      }),
    (v = s.conditionalFormats) == null ||
      v.forEach((g) => {
        ((g.id = fp()),
          g.kind === "formula" && g.formula && (g.formula = c(g.formula)));
      }));
  for (const g of [...(s.pivots ?? []), ...(s.queries ?? [])])
    ((g.id = fp()),
      (g.sourceRevision = e.revision),
      g.source.sheetId === n.id &&
        ((g.source.sheetId = s.id),
        g.source.tableId &&
          (g.source.tableId =
            u.get(g.source.tableId) ??
            ea("A copied analysis refers to a missing source table."))));
  const h = (e.definedNames ?? [])
      .filter((g) => g.scopeSheetId === n.id)
      .map((g) => ({
        ...structuredClone(g),
        scopeSheetId: s.id,
        refersTo: c(g.refersTo),
      })),
    f = [...e.sheets];
  return (
    f.splice(e.sheets.indexOf(n) + 1, 0, s),
    w0(
      lQ(
        {
          ...e,
          sheets: f,
          ...(h.length
            ? { definedNames: [...(e.definedNames ?? []), ...h] }
            : {}),
        },
        n.id,
        s.id,
      ),
    )
  );
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
function Vne(e, r, t = false) {
  $a(e);
  const n = qp(e, r);
  (e.sheets.length === 1 && ea("The final worksheet cannot be deleted."),
    zw(n) &&
      e.sheets.filter(zw).length === 1 &&
      ea("At least one worksheet must remain visible."));
  const o = AP(e, r);
  (o.some((s) => s.kind === "unsupported-reference") &&
    ea(
      "Deletion is blocked by ambiguous or unsupported references. Replace dynamic or three-dimensional references with explicit references first.",
    ),
    o.length &&
      !t &&
      ea(
        "This worksheet has dependencies. Review them and explicitly confirm deletion.",
      ));
  const a = structuredClone(e);
  a.sheets = a.sheets.filter((s) => s.id !== r);
  for (const s of a.sheets) {
    for (const d of s.tables ?? [])
      for (const u of d.columns)
        if (u.calculatedFormula) {
          const l = e0(u.calculatedFormula, e, n, void 0, true);
          (l !== u.calculatedFormula && jn(s), (u.calculatedFormula = l));
        }
    for (const d of Object.values(s.cells))
      !d.literal &&
        typeof d.raw == "string" &&
        d.raw.startsWith("=") &&
        (d.raw = e0(d.raw, e, n, void 0, true));
    for (const d of [...(s.conditionalFormats ?? []), ...(s.validations ?? [])])
      if ((d.kind === "formula" || d.kind === "custom") && d.formula) {
        const u = e0(d.formula, e, n, void 0, true);
        (u !== d.formula && jn(s), (d.formula = u));
      }
    (s.pivots && (s.pivots = s.pivots.filter((d) => d.source.sheetId !== r)),
      s.queries &&
        (s.queries = s.queries.filter((d) => d.source.sheetId !== r)));
  }
  return (
    a.definedNames &&
      (a.definedNames = a.definedNames
        .filter((s) => s.scopeSheetId !== r)
        .map((s) => ({ ...s, refersTo: e0(s.refersTo, e, n, void 0, true) }))),
    w0(fQ(a, r))
  );
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
async function Wne(e, r) {
  ($a(e), kP(r));
  const t = fp(),
    n = r
      ? { enabled: true, salt: t, passwordHash: await SP(r, t) }
      : { enabled: true };
  return w0({ ...e, structureProtection: n });
}
async function Hne(e, r) {
  kP(r);
  const t = e.structureProtection;
  if (!t) return e;
  (nR(t, { ...e }),
    t.passwordHash &&
      (await SP(r, t.salt)) !== t.passwordHash.toLowerCase() &&
      ea("The workbook structure password does not match."));
  const n = { ...e };
  return (delete n.structureProtection, w0(n));
}
const TP = ["any", "text", "number", "boolean", "date"];
function iR(e) {
  if (!TP.includes(e)) throw Error("Choose a supported table column type.");
  return e;
}
function bC(e, r) {
  if (e === void 0) return;
  const t = () => {
      throw Error(
        "Invalid or over-capacity table source evidence. Keep the native backup.",
      );
    },
    n = (s) => !!s && typeof s == "object" && !Array.isArray(s);
  (!Array.isArray(e) || e.length > 100 || JSON.stringify(e).length > 5e6) &&
    t();
  const o = new Set();
  let a = 0;
  r.tableEvidence = e.map((s) => {
    (!n(s) ||
      Object.keys(s).some(
        (c) =>
          ![
            "id",
            "kind",
            "sourceRevision",
            "sheetId",
            "sheetName",
            "tableId",
            "tableName",
            "columnId",
            "dataType",
            "changes",
          ].includes(c),
      )) &&
      t();
    const d = s;
    for (const c of ["id", "sheetId", "sheetName", "tableId", "tableName"])
      (typeof d[c] != "string" || !d[c] || d[c].length > 512) && t();
    ((o.has(d.id) ||
      !Number.isSafeInteger(d.sourceRevision) ||
      Number(d.sourceRevision) < 0 ||
      !["headers", "columnType"].includes(String(d.kind)) ||
      !Array.isArray(d.changes) ||
      d.changes.length > 500 ||
      (a += d.changes.length) > 1e4) &&
      t(),
      d.kind === "columnType"
        ? ((typeof d.columnId != "string" ||
            !d.columnId ||
            d.columnId.length > 100) &&
            t(),
          iR(d.dataType))
        : (d.columnId !== void 0 || d.dataType !== void 0) && t(),
      o.add(d.id));
    const u = new Set(),
      l = d.changes.map((c) => {
        (!n(c) ||
          Object.keys(c).some(
            (m) => !["address", "before", "after"].includes(m),
          ) ||
          typeof c.address != "string") &&
          t();
        const h = c,
          f = h.address,
          p = _t(f);
        return (
          (`${ut(p.col)}${p.row + 1}` !== f || u.has(f)) && t(),
          u.add(f),
          {
            address: f,
            ...(h.before !== void 0 ? { before: hi(h.before) } : {}),
            after: hi(h.after),
          }
        );
      });
    return { ...d, changes: l };
  });
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
function Gne(e, r, t, n) {
  var h;
  const o = Nr(e, r),
    a = (h = o.tables) == null ? void 0 : h.find((f) => f.id === t);
  if (!a) throw Error("Table not found.");
  const s = a.columns.findIndex((f) => f.id === n);
  if (s < 0) throw Error("Select an existing table column.");
  const d = ct(a.range),
    u = Math.max(0, d.lastRow - d.firstRow - Number(a.showTotals)),
    l = Vn(e),
    c = [];
  for (let f = 1; f <= Math.min(u, 500); f++) {
    const p = `${ut(d.firstCol + s)}${d.firstRow + f + 1}`,
      m = l(r, p);
    Vw(m, a.columns[s].dataType) || c.push({ address: p, value: m });
  }
  return {
    dataType: a.columns[s].dataType ?? "any",
    checked: Math.min(u, 500),
    total: u,
    truncated: u > 500,
    exceptions: c,
  };
}
function RP(e, r, t, n, o = {}) {
  var y, w;
  if (o.readOnly) throw Error("This workbook is read-only.");
  if (
    !n ||
    typeof n != "object" ||
    Object.keys(n).some(
      (E) =>
        ![
          "columnId",
          "dataType",
          "mode",
          "invalid",
          "decimal",
          "group",
        ].includes(E),
    )
  )
    throw Error("Invalid column type request.");
  if (
    (iR(n.dataType),
    !["convert", "tag"].includes(n.mode) ||
      !["reject", "keep"].includes(n.invalid) ||
      ![void 0, ".", ","].includes(n.decimal) ||
      ![void 0, "", ".", ",", " "].includes(n.group) ||
      (n.decimal ?? ".") === (n.group ?? ","))
  )
    throw Error(
      "Choose a conversion mode, exception resolution and different decimal/group separators.",
    );
  const a = structuredClone(n),
    s = Nr(e, r),
    d = (y = s.tables) == null ? void 0 : y.find((E) => E.id === t);
  if ((jn(s), !d)) throw Error("Table not found.");
  const u = d.columns.findIndex((E) => E.id === a.columnId);
  if (u < 0) throw Error("Select an existing table column.");
  const l = ct(d.range),
    c = l.lastRow - l.firstRow - Number(d.showTotals);
  if (c > 500)
    throw Error(
      "A reviewed column type operation supports at most 500 data rows. Preserve the native source and use bounded extracts for larger conversions.",
    );
  const h = structuredClone(e),
    f = Nr(h, r),
    p = f.tables.find((E) => E.id === t).columns[u];
  a.dataType === "any" ? delete p.dataType : (p.dataType = a.dataType);
  const m = [],
    v = [],
    g = Vn(e);
  for (let E = 1; E <= c; E++) {
    const A = `${ut(l.firstCol + u)}${l.firstRow + E + 1}`,
      k = s.cells[A],
      F = (k == null ? void 0 : k.raw) ?? null,
      I =
        typeof F == "string" && F.startsWith("=") && !(k != null && k.literal);
    try {
      if (I || a.mode === "tag") {
        if (!Vw(g(r, A), a.dataType))
          throw Error(
            I
              ? "Formula result does not match the selected type; the formula is preserved."
              : "Value does not match the selected type; source values are unchanged.",
          );
        continue;
      }
      const O = g(r, A);
      if (N5(F) && !N5(O)) {
        if (!Vw(O, a.dataType))
          throw Error(
            "Spill result does not match the selected type; the spill is preserved.",
          );
        continue;
      }
      const j = Xne(F, a);
      if (j !== F) {
        const z = hi({
          ...k,
          raw: j,
          ...(typeof j == "string" && j.startsWith("=")
            ? { literal: !0 }
            : { literal: void 0 }),
        });
        ((f.cells[A] = z),
          m.push({
            address: A,
            ...(k ? { before: structuredClone(k) } : {}),
            after: structuredClone(z),
          }));
      }
    } catch (O) {
      v.push({ address: A, value: F, reason: O.message });
    }
  }
  const b = d.columns[u].dataType ?? "any";
  return (
    (m.length || b !== a.dataType) &&
      ((h.tableEvidence = [
        ...(h.tableEvidence ?? []),
        {
          id: `${t}:${p.id}:${e.revision}:${((w = e.tableEvidence) == null ? void 0 : w.length) ?? 0}`,
          kind: "columnType",
          sourceRevision: e.revision,
          sheetId: r,
          sheetName: s.name,
          tableId: t,
          tableName: d.name,
          columnId: p.id,
          dataType: a.dataType,
          changes: structuredClone(m),
        },
      ]),
      bC(h.tableEvidence, h)),
    v0(
      h,
      m.map((E) => ({ ...E, sheetId: r })),
    ),
    {
      snapshot: JSON.stringify(e),
      sheetId: r,
      tableId: t,
      request: a,
      beforeType: b,
      canApply: !v.length || a.invalid === "keep",
      changes: m,
      exceptions: v,
      next: h,
    }
  );
}
function Zne(e, r, t = {}) {
  if (JSON.stringify(e) !== r.snapshot)
    throw Error("The workbook changed. Preview column types again.");
  const n = RP(e, r.sheetId, r.tableId, r.request, t);
  if (JSON.stringify(n) !== JSON.stringify(r))
    throw Error("The column type preview changed. Preview again.");
  if (!n.canApply)
    throw Error(
      "Resolve the listed exceptions or explicitly choose to retain them unchanged.",
    );
  return JSON.stringify(e) === JSON.stringify(n.next) ? e : n.next;
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
function Ww(e, r, t) {
  var I, O, j;
  (r.kind !== "chart" || !oR.includes(r.chart)) &&
    Rl("Choose a supported chart type.");
  const n = r.chart,
    o = r.range ?? "",
    a = ct(o),
    s = IP(r.chartOptions);
  (a.lastRow >= e.rowCount ||
    a.lastCol >= e.colCount ||
    (a.lastRow - a.firstRow + 1) * (a.lastCol - a.firstCol + 1) > 500) &&
    Rl("Select at most 500 cells inside this worksheet for a chart.");
  const d = a.firstRow + (r.headers ? 1 : 0),
    u = a.firstCol + (a.lastCol > a.firstCol ? 1 : 0);
  a.lastCol - u + 1 > 20 &&
    Rl(
      "Chart inspection supports at most 20 value columns; narrow the source range.",
    );
  const l = (z, V) => {
      const D = `${ut(V)}${z + 1}`,
        B = e.cells[D] ?? { raw: null };
      return {
        address: D,
        raw: B.raw,
        value: t(e.id, D),
        ...(B.literal ? { literal: true } : {}),
        ...(B.note ? { note: B.note } : {}),
      };
    },
    c = Array.from({ length: a.lastCol - u + 1 }, (z, V) => {
      var C;
      const D = u + V,
        B = (C = s.series) == null ? void 0 : C.find((T) => T.column === D + 1),
        _ = r.headers ? l(a.firstRow, D).value : ut(D);
      return {
        column: D + 1,
        name: (B == null ? void 0 : B.name) ?? String(_ ?? ut(D)),
        kind:
          (B == null ? void 0 : B.kind) ??
          (n === "combo" && V > 0 ? "line" : "column"),
        included: s.series === void 0 || !!B,
      };
    });
  ((I = s.series) != null &&
    I.some((z) => z.column < u + 1 || z.column > a.lastCol + 1) &&
    Rl(
      "A selected chart series no longer exists in its source range. Re-select series.",
    ),
    (O = s.includedRows) != null &&
      O.some((z) => z < d + 1 || z > a.lastRow + 1) &&
      Rl(
        "A chart filter row is outside its source range. Reset the row filter.",
      ));
  const h = cn(e, "row"),
    f = (z) => h.some((V) => z >= V.start && z <= V.end),
    p = Array.from({ length: Math.max(0, a.lastRow - d + 1) }, (z, V) => {
      const D = d + V,
        B = l(D, a.firstCol);
      return {
        row: D + 1,
        label: a.firstCol < u ? String(B.value ?? "") : B.address,
        category: B,
        included: s.includedRows === void 0 || s.includedRows.includes(D + 1),
        hidden: f(D),
        values: c.map((_) => l(D, _.column - 1)),
      };
    }),
    m = c.map((z, V) => (z.included ? V : -1)).filter((z) => z >= 0),
    v = p.filter((z) => z.included),
    g = [],
    b = (z, V) =>
      g.push({ severity: "error", message: z, ...(V ? { address: V } : {}) });
  (p.length || b("Select at least one data row below the chart header."),
    v.length || b("The chart filter includes no source rows."),
    m.length || b("Select at least one value series."),
    ["pie", "donut", "waterfall", "histogram", "sparkline"].includes(n) &&
      m.length !== 1 &&
      b(
        `${n} requires exactly one selected series. Other source series remain available in inspection.`,
      ));
  const y = (z) => z === null || z === "";
  for (let z = 0; z < c.length; z++) {
    let V = 0;
    for (const D of p) {
      const B = D.values[z];
      if (y(B.value)) continue;
      const _ = typeof B.value == "number" && Number.isFinite(B.value),
        C = D.included && c[z].included;
      if (!_) {
        g.push({
          severity: C ? "error" : "warning",
          message: `${B.address}: ${c[z].name} contains a nonnumeric or invalid calculated value${C ? "" : " outside the plotted selection"}.`,
          address: B.address,
        });
        continue;
      }
      C &&
        (V++,
        s.yScale === "log" &&
          Number(B.value) <= 0 &&
          b(`${B.address}: log-axis values must be positive.`, B.address),
        ["pie", "donut"].includes(n) &&
          Number(B.value) < 0 &&
          b(`${B.address}: pie/donut values cannot be negative.`, B.address));
    }
    c[z].included &&
      !V &&
      b(`${c[z].name}: the selected series is empty or has no numeric values.`);
  }
  if (n === "scatter") {
    a.firstCol === u &&
      b(
        "Scatter requires a numeric X column followed by at least one Y column.",
      );
    for (const V of v) {
      if (!V.values.some((B, _) => c[_].included && !y(B.value))) continue;
      const D = V.category.value;
      typeof D != "number" || !Number.isFinite(D)
        ? b(
            `${V.category.address}: scatter X values must be numeric.`,
            V.category.address,
          )
        : s.xScale === "log" &&
          D <= 0 &&
          b(
            `${V.category.address}: log X-axis values must be positive.`,
            V.category.address,
          );
    }
    const z = v
      .map((V) => V.category.value)
      .filter((V) => typeof V == "number" && Number.isFinite(V));
    z.length &&
      !Number.isFinite(Math.max(...z) - Math.min(...z)) &&
      b("Scatter X magnitude exceeds the finite plotting range.");
  } else
    s.xScale === "log" &&
      b("A logarithmic X axis is available only for scatter charts.");
  (s.yScale === "log" &&
    ["pie", "donut", "waterfall", "histogram", "sparkline"].includes(n) &&
    b(`Log axes are not supported for ${n}; choose a linear scale.`),
    ["pie", "donut"].includes(n) &&
      m.length === 1 &&
      !v.some((z) => Number(z.values[m[0]].value) > 0) &&
      b("Pie/donut requires at least one positive value."));
  const w = {
      sourceRows: p.length,
      includedRows: v.length,
      excludedRows: p.length - v.length,
      hiddenRows: p.filter((z) => z.hidden).length,
      availableSeries: c.length,
      selectedSeries: m.length,
      policy:
        "All source rows, including worksheet-hidden or filtered rows, are inspected. Only explicit chart row/series selections affect the plotted population. Blank cells are gaps; invalid values are diagnostics, never zero.",
    },
    E = `${e.name}!${o} · ${v.length}/${p.length} rows · ${m.length}/${c.length} series · ${w.hiddenRows} hidden rows in source`,
    A = {
      version: 1,
      objectId: r.id,
      kind: n,
      title: r.label,
      altText:
        ((j = s.altText) == null ? void 0 : j.trim()) ||
        `${r.label || n}. ${E}. ${w.policy}`,
      source: { sheetId: e.id, sheetName: e.name, range: o },
      axes: {
        x: { title: s.xTitle ?? "", scale: s.xScale ?? "linear" },
        y: {
          title: s.yTitle ?? (n === "histogram" ? "Frequency (records)" : ""),
          scale: s.yScale ?? "linear",
        },
        units: s.units ?? "",
      },
      options: s,
      population: w,
      scope: E,
      series: c,
      points: p,
      diagnostics: g,
    },
    k =
      m.length === 1
        ? v
            .filter(
              (z) =>
                typeof z.values[m[0]].value == "number" &&
                Number.isFinite(z.values[m[0]].value),
            )
            .map((z) => ({
              label: z.label,
              row: z.row,
              value: z.values[m[0]].value,
            }))
        : [];
  if (n === "histogram" && k.length) {
    const z = Math.min(...k.map((_) => _.value)),
      V = Math.max(...k.map((_) => _.value)),
      D =
        z === V
          ? 1
          : (s.histogramBins ?? Math.min(50, Math.ceil(Math.sqrt(k.length)))),
      B = (V - z) / D;
    if (!Number.isFinite(B))
      b("Histogram magnitude exceeds the finite calculation range.");
    else {
      A.histogram = Array.from({ length: D }, (_, C) => ({
        lower: z + C * B,
        upper: C === D - 1 ? V : z + (C + 1) * B,
        count: 0,
        rows: [],
        upperInclusive: C === D - 1,
      }));
      for (const _ of k) {
        const C =
          A.histogram[B ? Math.min(D - 1, Math.floor((_.value - z) / B)) : 0];
        (C.count++, C.rows.push(_.row));
      }
    }
  }
  if (n === "waterfall" && k.length) {
    let z = 0;
    ((A.waterfall = k.map((V) => {
      const D = z;
      return ((z += V.value), { ...V, start: D, end: z });
    })),
      A.waterfall.push({
        label: "Total",
        row: null,
        start: 0,
        end: z,
        value: z,
      }),
      (!Number.isFinite(z) ||
        A.waterfall.some((V) => !Number.isFinite(V.end))) &&
        b("Waterfall cumulative values exceed the finite calculation range."));
  }
  ["pie", "donut"].includes(n) &&
    m.length === 1 &&
    !Number.isFinite(k.reduce((z, V) => z + V.value, 0)) &&
    b("Pie/donut total exceeds the finite calculation range.");
  const F = m.flatMap((z) =>
    v
      .map((V) => V.values[z].value)
      .filter((V) => typeof V == "number" && Number.isFinite(V)),
  );
  return (
    F.length &&
      !Number.isFinite(Math.max(...F) - Math.min(0, ...F)) &&
      b("Chart magnitude exceeds the finite plotting range."),
    A
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
function P1(e) {
  var s;
  if (typeof e != "string" || e.length > Math.ceil(Hw / 3) * 4 + 30)
    throw Error("Pictures must be local PNG/JPEG data of at most 2 MB.");
  const r = /^data:image\/(png|jpeg);base64,([A-Za-z0-9+/]+={0,2})$/.exec(e);
  if (!r || r[2].length % 4 !== 0)
    throw Error("Pictures require canonical local PNG/JPEG base64 data.");
  let t;
  try {
    const d = atob(r[2]);
    if (btoa(d) !== r[2]) throw Error();
    t = Uint8Array.from(d, (u) => u.charCodeAt(0));
  } catch {
    throw Error("The picture data is invalid.");
  }
  if (t.length > Hw) throw Error("Pictures must be at most 2 MB.");
  let n = 0,
    o = 0;
  const a = () => {
    if (!n || !o || n > 1e4 || o > 1e4 || n * o > 4e7)
      throw Error(
        "Picture dimensions are invalid or exceed 40 megapixels / 10,000 pixels per side.",
      );
  };
  if (r[1] === "png") {
    if (
      t.length < 45 ||
      ![137, 80, 78, 71, 13, 10, 26, 10].every((p, m) => t[m] === p)
    )
      throw Error("The file contents do not match a PNG picture.");
    const d = new DataView(t.buffer, t.byteOffset, t.byteLength);
    if (
      d.getUint32(8) !== 13 ||
      String.fromCharCode(...t.subarray(12, 16)) !== "IHDR"
    )
      throw Error("PNG requires an intact first IHDR chunk.");
    if (
      ((n = d.getUint32(16)),
      (o = d.getUint32(20)),
      a(),
      !(
        (s = {
          0: [1, 2, 4, 8, 16],
          2: [8, 16],
          3: [1, 2, 4, 8],
          4: [8, 16],
          6: [8, 16],
        }[t[25]]) != null && s.includes(t[24])
      ) ||
        t[26] !== 0 ||
        t[27] !== 0 ||
        t[28] > 1)
    )
      throw Error("Unsupported PNG pixel encoding.");
    let l = 8,
      c = 0,
      h = false,
      f = false;
    for (; l < t.length;) {
      if (++c > 4096 || l + 12 > t.length)
        throw Error("PNG chunk structure is invalid or excessive.");
      const p = d.getUint32(l),
        m = String.fromCharCode(...t.subarray(l + 4, l + 8));
      if (p > t.length - l - 12 || !/^[A-Za-z]{4}$/.test(m))
        throw Error("PNG chunk length or type is invalid.");
      if (m === "IHDR" && l !== 8) throw Error("PNG has duplicate dimensions.");
      if ((m === "IDAT" && p > 0 && (h = true), m === "IEND")) {
        if (p !== 0 || l + 12 !== t.length)
          throw Error("PNG has invalid trailing data.");
        f = true;
      }
      l += p + 12;
    }
    if (!h || !f) throw Error("PNG data or end marker is missing.");
  } else {
    if (t.length < 12 || t[0] !== 255 || t[1] !== 216)
      throw Error("The file contents do not match a JPEG picture.");
    let d = 2,
      u = 0,
      l = false,
      c = false;
    for (; d < t.length;) {
      if (++u > 65536) throw Error("JPEG has too many segments.");
      if (l) for (; d < t.length && t[d] !== 255;) d++;
      if (d >= t.length || t[d++] !== 255)
        throw Error("JPEG marker structure is invalid.");
      for (; t[d] === 255;) d++;
      const h = t[d++];
      if (l && (h === 0 || (h >= 208 && h <= 215))) continue;
      if (h === 217) {
        if (d !== t.length) throw Error("JPEG has trailing data.");
        c = true;
        break;
      }
      if (h === 1) continue;
      if (d + 2 > t.length) throw Error("JPEG segment is truncated.");
      const f = t[d] * 256 + t[d + 1];
      if (f < 2 || d + f > t.length)
        throw Error("JPEG segment length is invalid.");
      if (h >= 192 && h <= 207 && ![196, 200, 204].includes(h)) {
        if (![192, 193, 194].includes(h) || f < 8 || t[d + 2] !== 8)
          throw Error("Unsupported JPEG pixel encoding.");
        const p = t[d + 7];
        if (p < 1 || p > 4 || f !== 8 + p * 3)
          throw Error("JPEG frame components are invalid.");
        const m = t[d + 3] * 256 + t[d + 4],
          v = t[d + 5] * 256 + t[d + 6];
        if (n && (n !== v || o !== m)) throw Error("JPEG dimensions conflict.");
        ((n = v), (o = m), a());
      }
      if (h === 218) {
        if (!n || f < 6 || f !== 6 + t[d + 2] * 2)
          throw Error("JPEG scan is invalid.");
        l = true;
      }
      d += f;
    }
    if (!l || !c) throw Error("JPEG scan or end marker is missing.");
    a();
  }
  return { bytes: t, extension: r[1], width: n, height: o };
}
function Jne(e, r, t) {
  const n = zv(e, r, t);
  if (!n.trim() || /[\u0000-\u001f\u007f]/.test(n))
    throw Error(
      `Invalid picture ${t}: use nonblank plain text without control characters.`,
    );
  return n;
}
function ek(e, r) {
  const { bytes: t } = P1(e),
    n = FE(t);
  return _P(
    {
      ...r,
      version: 1,
      sourceReference: `sha256:${n}`,
      sha256: n,
      byteLength: t.length,
    },
    t,
  );
}
function _P(e, r) {
  if (
    !R5(e) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(e)) ||
    Reflect.ownKeys(e).some(
      (o) =>
        typeof o != "string" ||
        ![
          "version",
          "origin",
          "sourceReference",
          "sha256",
          "byteLength",
          "filename",
          "sourceLocator",
          "containerFilename",
          "mediaReference",
        ].includes(o),
    )
  )
    throw Error("Invalid picture evidence metadata.");
  if (
    e.version !== 1 ||
    typeof e.origin != "string" ||
    !["local-file", "xlsx-media", "legacy-native"].includes(e.origin)
  )
    throw Error("Invalid picture evidence origin.");
  const t = FE(r);
  if (
    typeof e.sha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(e.sha256) ||
    e.sha256 !== t ||
    e.sourceReference !== `sha256:${t}` ||
    e.byteLength !== r.length
  )
    throw Error(
      "Picture evidence digest or byte length does not match the embedded image.",
    );
  const n = {
    version: 1,
    origin: e.origin,
    sourceReference: `sha256:${t}`,
    sha256: t,
    byteLength: r.length,
  };
  for (const o of [
    "filename",
    "sourceLocator",
    "containerFilename",
    "mediaReference",
  ])
    e[o] !== void 0 && (n[o] = Jne(e[o], o === "sourceLocator" ? 2e3 : 512, o));
  if (n.origin === "local-file" && !n.filename)
    throw Error("A local evidence picture requires its original filename.");
  if (n.origin !== "xlsx-media" && (n.containerFilename || n.mediaReference))
    throw Error("Only observed XLSX media may have container metadata.");
  if (n.origin === "legacy-native" && n.filename)
    throw Error(
      "An original filename cannot be inferred for a legacy picture.",
    );
  return n;
}
function F2(e, r, t) {
  if (!e.sheets.some((n) => n.id === r))
    throw Error("The worksheet no longer exists.");
  return {
    workbookId: e.id,
    revision: e.revision,
    sheetId: r,
    ...(t !== void 0 ? { selection: t } : {}),
    guard: JSON.stringify(e),
  };
}
function BE(e, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  const n = e.sheets.find((a) => a.id === r);
  if (!n) throw Error("The worksheet no longer exists.");
  jn(n);
  const o = t.snapshot;
  if (
    o &&
    (o.workbookId !== e.id ||
      o.revision !== e.revision ||
      o.sheetId !== r ||
      o.guard !== JSON.stringify(e) ||
      o.selection !== t.selection)
  )
    throw Error(
      "The drawing preview is stale. Reopen it or select the object again.",
    );
}
function sR(e, r) {
  const t = ct(e);
  if (
    t.lastRow >= r.rowCount ||
    t.lastCol >= r.colCount ||
    (t.lastRow - t.firstRow + 1) * (t.lastCol - t.firstCol + 1) > Yne
  )
    throw new Error(
      "Select at most 500 cells inside this worksheet for a chart.",
    );
  return t;
}
function yC(e, r) {
  if (e === void 0) return;
  if (!Array.isArray(e) || e.length > Kne)
    throw new Error("A worksheet supports at most 100 drawing objects.");
  const t = new Set();
  let n = 0,
    o = 0;
  for (const s of e)
    if (R5(s) && s.kind === "picture" && typeof s.dataURL == "string") {
      if (s.dataURL.length > Math.ceil(Hw / 3) * 4 + 100)
        throw Error("Pictures must be at most 2 MB.");
      const d = s.dataURL.slice(s.dataURL.indexOf(",") + 1);
      if (
        ((o += Math.max(
          0,
          Math.floor((d.length * 3) / 4) -
            (d.endsWith("==") ? 2 : d.endsWith("=") ? 1 : 0),
        )),
        o > 10 * 1024 * 1024)
      )
        throw Error("Pictures in one worksheet may use at most 10 MB.");
    }
  const a = e.map((s) => {
    if (!R5(s) || !["picture", "shape", "chart"].includes(String(s.kind)))
      throw new Error("Invalid worksheet drawing object.");
    const d = zv(s.id, 100, "id"),
      u = zv(s.anchor, 20, "anchor");
    if (!d || t.has(d)) throw new Error("Drawing IDs must be unique.");
    t.add(d);
    const l = _t(u);
    if (l.row >= r.rowCount || l.col >= r.colCount)
      throw new Error("Drawing anchor exceeds the worksheet.");
    const c = {
      id: d,
      kind: s.kind,
      anchor: `${ut(l.col)}${l.row + 1}`,
      offsetX: L4(s.offsetX, -4e3, 4e3, "horizontal offset"),
      offsetY: L4(s.offsetY, -4e3, 4e3, "vertical offset"),
      width: L4(s.width, 24, 4e3, "width"),
      height: L4(s.height, 24, 4e3, "height"),
      label: zv(s.label, 500, "label"),
      fill: AO(s.fill),
      stroke: AO(s.stroke),
    };
    if (c.kind === "picture") {
      c.dataURL = zv(s.dataURL, Math.ceil(Hw / 3) * 4 + 100, "picture");
      const h = P1(c.dataURL);
      ((n += h.bytes.length),
        s.evidence !== void 0 && (c.evidence = _P(s.evidence, h.bytes)));
    } else if (s.evidence !== void 0)
      throw Error("Evidence metadata belongs only to pictures.");
    if (c.kind === "shape") {
      if (!["rectangle", "ellipse", "arrow", "text"].includes(String(s.shape)))
        throw new Error("Unsupported shape.");
      c.shape = s.shape;
    }
    if (c.kind === "chart") {
      if (!oR.includes(s.chart)) throw new Error("Unsupported chart type.");
      if (
        ((c.chart = s.chart),
        (c.range = zv(s.range, 50, "chart range")),
        sR(c.range, r),
        typeof s.headers != "boolean")
      )
        throw new Error("Invalid chart header setting.");
      ((c.headers = s.headers),
        s.chartOptions !== void 0 && (c.chartOptions = IP(s.chartOptions)));
    }
    return c;
  });
  if (n > 10 * 1024 * 1024)
    throw new Error("Pictures in one worksheet may use at most 10 MB.");
  r.objects = a;
}
function I5(e, r, t, n = {}) {
  BE(e, r, n);
  const o = e.sheets.find((s) => s.id === r),
    a = { ...o };
  return (
    yC(t, a),
    JSON.stringify(o.objects ?? []) === JSON.stringify(a.objects ?? [])
      ? e
      : { ...e, sheets: e.sheets.map((s) => (s.id !== r ? s : a)) }
  );
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
const Pa = (e) => {
    throw Error(e);
  },
  lR = (e) => typeof e == "number" && Number.isFinite(e),
  LE = (e) => e === null || e === "",
  kO = (e) => !!e && typeof e == "object" && !Array.isArray(e);
function tie(e) {
  kO(e) || Pa("Invalid pivot refresh evidence.");
  const r = e;
  (typeof r.at != "string" ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(r.at) ||
    !Number.isFinite(Date.parse(r.at)) ||
    new Date(r.at).toISOString() !== r.at ||
    typeof r.sourceSheetName != "string" ||
    !r.sourceSheetName ||
    r.sourceSheetName.length > 31 ||
    typeof r.sourceRange != "string" ||
    !/^[A-Z]+\d+:[A-Z]+\d+$/.test(r.sourceRange)) &&
    Pa("Invalid pivot refresh time or source label.");
  for (const n of ["sourceFingerprint", "outputFingerprint"])
    (typeof r[n] != "string" || !/^[a-f0-9]{64}$/.test(r[n])) &&
      Pa("Invalid pivot content fingerprint.");
  for (const n of ["sourceRows", "includedRows", "excludedRows"])
    (!Number.isSafeInteger(r[n]) || Number(r[n]) < 0 || Number(r[n]) > 1e5) &&
      Pa("Invalid pivot population counts.");
  (Number(r.sourceRows) !== Number(r.includedRows) + Number(r.excludedRows) ||
    !Array.isArray(r.measures) ||
    !r.measures.length ||
    r.measures.length > 100) &&
    Pa("Pivot population evidence does not reconcile.");
  const t = r.measures.map((n) => {
    (!kO(n) ||
      typeof n.name != "string" ||
      !n.name ||
      n.name.length > 512 ||
      typeof n.field != "string" ||
      !n.field ||
      n.field.length > 255 ||
      !["sum", "count", "average", "min", "max", "distinctCount"].includes(
        String(n.operation),
      ) ||
      !(
        n.result === null ||
        typeof n.result == "boolean" ||
        lR(n.result) ||
        (typeof n.result == "string" && n.result.length <= 32767)
      )) &&
      Pa("Invalid pivot measure evidence.");
    const o = n;
    for (const a of ["numericCount", "nonblankCount"])
      (!Number.isSafeInteger(o[a]) ||
        Number(o[a]) < 0 ||
        Number(o[a]) > Number(r.includedRows)) &&
        Pa("Invalid pivot measure count.");
    return (
      Number(o.numericCount) > Number(o.nonblankCount) &&
        Pa("Invalid pivot numeric count."),
      {
        name: o.name,
        field: o.field,
        operation: o.operation,
        result: o.result,
        numericCount: o.numericCount,
        nonblankCount: o.nonblankCount,
      }
    );
  });
  return {
    at: r.at,
    sourceSheetName: r.sourceSheetName,
    sourceRange: r.sourceRange,
    sourceFingerprint: r.sourceFingerprint,
    outputFingerprint: r.outputFingerprint,
    sourceRows: r.sourceRows,
    includedRows: r.includedRows,
    excludedRows: r.excludedRows,
    measures: t,
  };
}
function DP(e, r, t) {
  if (LE(e)) return null;
  let n;
  if (r.dateSystem === "iso") {
    if (typeof e != "string" || !/^\d{4}-\d{2}-\d{2}$/.test(e))
      return Pa(
        `Date grouping for “${r.field}” needs YYYY-MM-DD text on source row ${t + 1}; choose the correct date system or clean the source.`,
      );
    if (
      ((n = new Date(e + "T00:00:00.000Z")),
      !Number.isFinite(n.getTime()) ||
        n.toISOString().slice(0, 10) !== e ||
        e.startsWith("0000"))
    )
      return Pa(`Invalid date in “${r.field}” on source row ${t + 1}.`);
  } else {
    if (!lR(e) || e < (r.dateSystem === "1900" ? 1 : 0) || e > 2958465)
      return Pa(
        `Date grouping for “${r.field}” requires a valid ${r.dateSystem} numeric date serial on source row ${t + 1}.`,
      );
    const s = Math.floor(e),
      d =
        r.dateSystem === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
    if (
      ((n = new Date(
        d + (s - (r.dateSystem === "1900" && s >= 60 ? 1 : 0)) * 864e5,
      )),
      n.getUTCFullYear() > 9999)
    )
      return Pa("Grouped date exceeds year 9999.");
  }
  const o = String(n.getUTCFullYear()).padStart(4, "0"),
    a = n.getUTCMonth() + 1;
  return r.unit === "year"
    ? o
    : r.unit === "quarter"
      ? `${o} Q${Math.ceil(a / 3)}`
      : `${o}-${String(a).padStart(2, "0")}`;
}
const tk = (e = false) => ({
  count: 0,
  numeric: 0,
  sum: 0,
  min: 1 / 0,
  max: -1 / 0,
  ...(e ? { distinct: new Set() } : {}),
});
function SO(e, r) {
  var t;
  (LE(r) ||
    (e.count++,
    (t = e.distinct) == null ||
      t.add(
        JSON.stringify([
          typeof r,
          typeof r == "string" ? r.toLocaleLowerCase("en-US") : r,
        ]),
      )),
    lR(r) &&
      (e.numeric++,
      (e.sum += r),
      (e.min = Math.min(e.min, r)),
      (e.max = Math.max(e.max, r)),
      Number.isFinite(e.sum) ||
        Pa("Aggregation exceeds finite numeric bounds.")));
}
function rie(e, r) {
  var t;
  return r === "count"
    ? e.count
    : r === "distinctCount"
      ? (((t = e.distinct) == null ? void 0 : t.size) ?? 0)
      : r === "sum"
        ? e.sum
        : e.numeric
          ? r === "average"
            ? e.sum / e.numeric
            : r === "min"
              ? e.min
              : e.max
          : null;
}
function OP(e, r, t) {
  const n = (C) => {
      const T = r.indexOf(C);
      return T < 0
        ? Pa(`Source field “${C}” was not found. Refresh the field selection.`)
        : T;
    },
    o = e.rowFields.map(n),
    a = e.columnFields.map(n);
  for (const C of e.dateGroups ?? [])
    [...e.rowFields, ...e.columnFields].includes(C.field) ||
      Pa(`Date group “${C.field}” must be selected as a row or column field.`);
  for (const C of e.calculatedFields ?? [])
    (r.some((T) => T.toLowerCase() === C.name.toLowerCase()) &&
      Pa(`Calculated field “${C.name}” duplicates a source header.`),
      n(C.field),
      C.otherField && n(C.otherField));
  const s = e.values.map((C) => {
      var N;
      const T =
        (N = e.calculatedFields) == null
          ? void 0
          : N.find((R) => R.name === C.field);
      return (
        T &&
          C.operation !== "sum" &&
          Pa(
            `Calculated field “${C.field}” requires sum aggregation; its arithmetic is applied to source-field sums.`,
          ),
        {
          ...C,
          index: n((T == null ? void 0 : T.field) ?? C.field),
          calculated: T,
          other: T != null && T.otherField ? n(T.otherField) : void 0,
        }
      );
    }),
    d = () =>
      s.map((C) => ({
        ...tk(C.operation === "distinctCount"),
        ...(C.other !== void 0 ? { right: tk() } : {}),
      }));
  t.length *
    s.length *
    Math.max(1, e.subtotals ? o.length : 1) *
    Math.max(1, e.subtotals ? a.length : 1) >
    2e6 &&
    Pa(
      "Pivot aggregation exceeds the 2,000,000 contribution budget. Reduce value fields, hierarchy depth or source rows.",
    );
  let u = 0;
  const l = (C, T) =>
      s.forEach((N, R) => {
        var $, U;
        const L = (($ = C[R].distinct) == null ? void 0 : $.size) ?? 0;
        (SO(C[R], T[N.index]),
          (u += (((U = C[R].distinct) == null ? void 0 : U.size) ?? 0) - L),
          u > 25e4 &&
            Pa(
              "Pivot distinct counts exceed 250,000 retained group entries. Reduce groups, value fields or source rows.",
            ),
          N.other !== void 0 && SO(C[R].right, T[N.other]));
      }),
    c = (C, T) => {
      const N = s[T],
        R = C == null ? void 0 : C[T];
      if (!N.calculated)
        return rie(R ?? tk(N.operation === "distinctCount"), N.operation);
      if (!R) return null;
      const L = R.sum,
        $ = N.other !== void 0 ? R.right.sum : N.calculated.constant;
      N.calculated.operation === "divide" &&
        $ === 0 &&
        Pa(
          `Calculated field “${N.field}” divides by zero in a group or total. Filter, clean or change the calculation; no result was written.`,
        );
      const U =
        N.calculated.operation === "add"
          ? L + $
          : N.calculated.operation === "subtract"
            ? L - $
            : N.calculated.operation === "multiply"
              ? L * $
              : L / $;
      return Number.isFinite(U)
        ? U
        : Pa(`Calculated field “${N.field}” exceeds finite numeric bounds.`);
    },
    h = (C, T) =>
      C.map((N) => {
        var $;
        const R =
            ($ = e.dateGroups) == null
              ? void 0
              : $.find((U) => U.field === r[N]),
          L = T.values[N];
        return R ? DP(L, R, T.index) : LE(L) ? null : L;
      }),
    f = (C) =>
      [
        C,
        ...(e.subtotals
          ? Array.from({ length: Math.max(0, C.length - 1) }, (T, N) =>
              C.slice(0, N + 1),
            )
          : []),
      ].map((T) => ({
        key: JSON.stringify(T),
        parts: T,
        subtotal: T.length < C.length,
      })),
    p = new Map(),
    m = new Map(),
    v = new Map(),
    g = new Map(),
    b = new Map(),
    y = d(),
    w = (C, T, N) => {
      (C.has(T) || C.set(T, d()), l(C.get(T), N));
    };
  for (const C of t) {
    const T = f(h(o, C)),
      N = f(h(a, C));
    (T.forEach((L) => p.set(L.key, L)), N.forEach((L) => m.set(L.key, L)));
    const R =
      Math.max(1, o.length) +
      (m.size + +(!!e.grandTotals && !!a.length)) * s.length;
    (p.size + 2) * R > 1e5 &&
      Pa(
        "The pivot grouping would exceed the 100,000-cell output budget. Reduce group fields or filter the source.",
      );
    for (const L of T) {
      w(g, L.key, C.values);
      for (const $ of N) w(v, JSON.stringify([L.key, $.key]), C.values);
    }
    for (const L of N) w(b, L.key, C.values);
    l(y, C.values);
  }
  (m.size || m.set("[]", { key: "[]", parts: [], subtotal: false }),
    !p.size &&
      !o.length &&
      p.set("[]", { key: "[]", parts: [], subtotal: false }));
  function E(C, T) {
    var $;
    if (!e.subtotals && !(($ = e.sorts) != null && $.length))
      return [...C.values()];
    if (!T.length) return [...C.values()];
    const N = new Map();
    for (const U of C.values())
      if (!U.subtotal)
        for (let G = 0; G < U.parts.length; G++) {
          const H = JSON.stringify(U.parts.slice(0, G)),
            Q = U.parts.slice(0, G + 1),
            X = JSON.stringify(Q);
          (N.has(H) || N.set(H, new Map()), N.get(H).set(X, Q));
        }
    const R = [],
      L = (U) => {
        var X, J;
        const G =
            (X = e.sorts) == null
              ? void 0
              : X.find((q) => q.field === T[U.length]),
          H = [
            ...(((J = N.get(JSON.stringify(U))) == null
              ? void 0
              : J.values()) ?? []),
          ];
        G &&
          H.sort((q, W) => {
            const K = q.at(-1),
              te = W.at(-1),
              ae =
                K === te
                  ? 0
                  : K === null
                    ? -1
                    : te === null
                      ? 1
                      : typeof K == "number" && typeof te == "number"
                        ? K - te
                        : String(K).localeCompare(String(te), "en-US", {
                            numeric: true,
                          });
            return G.direction === "ascending" ? ae : -ae;
          });
        for (const q of H) L(q);
        const Q = C.get(JSON.stringify(U));
        Q && R.push(Q);
      };
    return (L([]), R);
  }
  const A = E(m, e.columnFields),
    k = E(p, e.rowFields),
    F = (C) =>
      C.name ??
      (C.calculated
        ? `Calculated ${C.field}`
        : `${C.operation === "distinctCount" ? "Distinct count" : C.operation[0].toUpperCase() + C.operation.slice(1)} of ${C.field}`),
    I = (C) => (LE(C) ? "(blank)" : String(C)),
    O = e.rowFields.length
      ? e.rowFields.map((C) => {
          var N;
          const T =
            (N = e.dateGroups) == null ? void 0 : N.find((R) => R.field === C);
          return T ? `${C} (${T.unit})` : C;
        })
      : ["Row labels"];
  for (const C of A)
    for (const T of s)
      O.push(
        [...C.parts.map(I), ...(C.subtotal ? ["Subtotal"] : []), F(T)].join(
          " / ",
        ),
      );
  if (e.grandTotals && a.length)
    for (const C of s) O.push(`Grand Total / ${F(C)}`);
  s.forEach((C, T) => {
    if (C.showAs && C.showAs !== "value")
      for (let N = O.length - 1; N >= Math.max(1, o.length); N--)
        (N - Math.max(1, o.length)) % s.length === T &&
          (O[N] =
            `${O[N]} (% of ${C.showAs === "percentGrand" ? "grand" : C.showAs === "percentRow" ? "row" : "column"} total)`);
  });
  const j = [],
    z = [...A, ...(e.grandTotals && a.length ? [null] : [])],
    V = [O],
    D = (C, T, N, R) => {
      const L = s[T].showAs;
      if (!L || L === "value" || C === null) return C;
      const $ = c(L === "percentGrand" ? y : L === "percentRow" ? N : R, T);
      return typeof C == "number" && typeof $ == "number" && $ !== 0
        ? C / $
        : "#DIV/0!";
    },
    B = [O],
    _ = [];
  for (const C of k) {
    const T = C.subtotal
        ? [
            ...C.parts,
            "Subtotal",
            ...Array(o.length - C.parts.length - 1).fill(null),
          ]
        : C.parts.length
          ? [...C.parts]
          : ["Total"],
      N = [...T];
    for (const R of z)
      s.forEach((L, $) => {
        const U = c(
          R ? v.get(JSON.stringify([C.key, R.key])) : g.get(C.key),
          $,
        );
        (N.push(U), T.push(D(U, $, g.get(C.key), R ? b.get(R.key) : y)));
      });
    (C.subtotal && _.push(B.length), B.push(T), V.push(N), j.push(C));
  }
  if (e.grandTotals && o.length) {
    const C = ["Grand Total", ...Array(o.length - 1).fill(null)],
      T = [...C];
    for (const N of z)
      s.forEach((R, L) => {
        const $ = c(N ? b.get(N.key) : y, L);
        (T.push($), C.push(D($, L, y, N ? b.get(N.key) : y)));
      });
    (B.push(C), V.push(T), j.push(null), _.push(B.length - 1));
  }
  return {
    matrix: B,
    baseMatrix: V,
    rowAxes: j,
    columnAxes: z,
    chartRows: k
      .filter((C) => !C.subtotal)
      .map((C) => ({
        label: C.parts.map(I).join(" / ") || "Total",
        values: s.map((T, N) => D(c(g.get(C.key), N), N, g.get(C.key), y)),
      })),
    subtotalRows: _,
    measures: s.map((C, T) => ({
      name: F(C),
      field: C.field,
      operation: C.operation,
      result: c(y, T),
      numericCount: y[T].numeric,
      nonblankCount: y[T].count,
    })),
  };
}
const nie = ["value", "percentGrand", "percentRow", "percentColumn"],
  Lu = (e) => {
    throw Error(e);
  },
  P4 = (e) => !!e && typeof e == "object" && !Array.isArray(e),
  rk = (e) =>
    typeof e == "string" && e.trim() && e.length <= 255
      ? e
      : Lu("Choose a valid pivot field."),
  _5 = (e) =>
    typeof e == "string" &&
    /^\d{4}-\d{2}-\d{2}$/.test(e) &&
    !e.startsWith("0000") &&
    Number.isFinite(Date.parse(e)) &&
    new Date(e).toISOString().slice(0, 10) === e,
  U4 = (e, r) =>
    Array.isArray(e) && e.length <= r
      ? e
      : Lu(`Pivot controls require a list of at most ${r} entries.`),
  Cp = (e) =>
    JSON.stringify([
      typeof e,
      typeof e == "string" ? e.toLocaleLowerCase("en-US") : e,
    ]);
function iie(e) {
  var t;
  const r = {};
  (e.slicers !== void 0 &&
    (r.slicers = U4(e.slicers, 20).map((n) => {
      if (!P4(n)) return Lu("Invalid pivot slicer.");
      const o = U4(n.selected, 500).map(
        (a) => (
          a !== null &&
            typeof a != "boolean" &&
            !(typeof a == "number" && Number.isFinite(a)) &&
            !(typeof a == "string" && a.length <= 32767) &&
            Lu("Invalid slicer item."),
          a
        ),
      );
      return (
        new Set(o.map(Cp)).size !== o.length &&
          Lu("Slicer selections must be unique."),
        { field: rk(n.field), selected: o }
      );
    })),
    e.timelines !== void 0 &&
      (r.timelines = U4(e.timelines, 10).map((n) =>
        !P4(n) ||
        !_5(n.from) ||
        !_5(n.to) ||
        n.from > n.to ||
        !["year", "quarter", "month"].includes(String(n.unit)) ||
        !["iso", "1900", "1904"].includes(String(n.dateSystem))
          ? Lu(
              "Timeline requires ordered YYYY-MM-DD bounds, a period and an explicit date system.",
            )
          : {
              field: rk(n.field),
              from: n.from,
              to: n.to,
              unit: n.unit,
              dateSystem: n.dateSystem,
            },
      )),
    e.sorts !== void 0 &&
      (r.sorts = U4(e.sorts, 100).map((n) =>
        !P4(n) || !["ascending", "descending"].includes(String(n.direction))
          ? Lu("Pivot field sorting requires ascending or descending order.")
          : { field: rk(n.field), direction: n.direction },
      )));
  for (const n of [r.slicers, r.timelines, r.sorts])
    n &&
      new Set(n.map((o) => o.field)).size !== n.length &&
      Lu("Choose only one control of each kind per field.");
  if (e.chart !== void 0) {
    const n = e.chart;
    if (
      !P4(n) ||
      !["column", "bar", "line"].includes(String(n.kind)) ||
      !Number.isInteger(n.valueIndex) ||
      Number(n.valueIndex) < 0 ||
      Number(n.valueIndex) >= Number((t = e.values) == null ? void 0 : t.length)
    )
      return Lu(
        "PivotChart requires a supported chart and an existing value field.",
      );
    r.chart = { kind: n.kind, valueIndex: n.valueIndex };
  }
  return r;
}
function uR(e, r) {
  var o;
  const t = new Set(
      (o = e.calculatedFields) == null ? void 0 : o.map((a) => a.name),
    ),
    n = [
      ...e.rowFields,
      ...e.columnFields,
      ...(e.filters ?? []).map((a) => a.field),
      ...(e.slicers ?? []).map((a) => a.field),
      ...(e.timelines ?? []).map((a) => a.field),
      ...(e.dateGroups ?? []).map((a) => a.field),
      ...(e.sorts ?? []).map((a) => a.field),
      ...e.values.filter((a) => !t.has(a.field)).map((a) => a.field),
      ...(e.calculatedFields ?? []).flatMap((a) => [
        a.field,
        ...(a.otherField ? [a.otherField] : []),
      ]),
    ];
  for (const a of n)
    r.includes(a) ||
      Lu(
        `Broken pivot field “${a}”: source field was not found. Restore the header or remove/reselect this field before refreshing. Historical totals are not current.`,
      );
  for (const a of e.sorts ?? [])
    [...e.rowFields, ...e.columnFields].includes(a.field) ||
      Lu(`Sort field “${a.field}” must be a row or column group.`);
}
function MP(e, r, t, n) {
  if (e === null || e === "") return null;
  if (r === "iso")
    return _5(e)
      ? e
      : Lu(
          `Timeline field “${t}” has an invalid ISO date on source row ${n + 1}.`,
        );
  if (
    typeof e != "number" ||
    !Number.isFinite(e) ||
    e < (r === "1900" ? 1 : 0) ||
    e > 2958465
  )
    return Lu(
      `Timeline field “${t}” requires ${r} date serials on source row ${n + 1}.`,
    );
  const o = Math.floor(e);
  if (r === "1900" && o === 60) return "1900-02-29";
  const a = new Date(
    (r === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1)) +
      (o - (r === "1900" && o >= 60 ? 1 : 0)) * 864e5,
  );
  return a.getUTCFullYear() > 9999
    ? Lu("Timeline date exceeds year 9999.")
    : a.toISOString().slice(0, 10);
}
function jP(e, r, t) {
  uR(e, r);
  const n = (e.slicers ?? []).map((a) => ({
      index: r.indexOf(a.field),
      keys: new Set(a.selected.map(Cp)),
    })),
    o = (e.timelines ?? []).map((a) => ({ ...a, index: r.indexOf(a.field) }));
  return t.filter((a) => {
    const s = o.map((d) =>
      MP(a.values[d.index], d.dateSystem, d.field, a.index),
    );
    return (
      n.every((d) => d.keys.has(Cp(a.values[d.index]))) &&
      o.every((d, u) => s[u] !== null && s[u] >= d.from && s[u] <= d.to)
    );
  });
}
function oie(e, r) {
  const t = new Map();
  for (const n of e) {
    const o = MP(n, r.dateSystem, r.field, 0);
    if (!o) continue;
    const a = Number(o.slice(0, 4)),
      s = Number(o.slice(5, 7)),
      d =
        r.unit === "year"
          ? 1
          : r.unit === "quarter"
            ? Math.floor((s - 1) / 3) * 3 + 1
            : s,
      u = r.unit === "year" ? 12 : r.unit === "quarter" ? d + 2 : d,
      l = String(a).padStart(4, "0"),
      c =
        r.unit === "year"
          ? l
          : r.unit === "quarter"
            ? `${l} Q${Math.ceil(s / 3)}`
            : `${l}-${String(s).padStart(2, "0")}`,
      h = new Date(0);
    (h.setUTCFullYear(a, u, 0),
      h.setUTCHours(0, 0, 0, 0),
      t.set(c, {
        label: c,
        from: `${l}-${String(d).padStart(2, "0")}-01`,
        to: h.toISOString().slice(0, 10),
      }));
  }
  return [...t.values()].sort((n, o) => n.from.localeCompare(o.from));
}
const TO = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    (Object.getPrototypeOf(e) === Object.prototype ||
      Object.getPrototypeOf(e) === null),
  $4 = (e, r) => typeof e == "string" && e.length > 0 && e.length <= r;
function aie(e, r) {
  const t = e.sheets;
  if (e.rowOrigins === void 0 && !t.some((a) => a.rowOriginIds !== void 0))
    return;
  if (!Array.isArray(e.rowOrigins) || e.rowOrigins.length > Mt.maxCells)
    throw Error(
      "Row provenance requires a registry of at most 100,000 source rows.",
    );
  const n = new Set();
  r.rowOrigins = e.rowOrigins.map((a) => {
    if (
      !TO(a) ||
      !$4(a.id, 100) ||
      n.has(a.id) ||
      !$4(a.sourceId, 100) ||
      !$4(a.sourceName, 255) ||
      !$4(a.sourceSheet, 31) ||
      !["file", "snapshot"].includes(String(a.kind)) ||
      !Number.isInteger(a.originalRow) ||
      Number(a.originalRow) < 1 ||
      Number(a.originalRow) > Mt.maxRows
    )
      throw Error("Invalid, duplicated or out-of-bounds original source row.");
    return (
      n.add(a.id),
      {
        id: a.id,
        sourceId: a.sourceId,
        sourceName: a.sourceName,
        sourceSheet: a.sourceSheet,
        originalRow: a.originalRow,
        kind: a.kind,
      }
    );
  });
  let o = 0;
  for (let a = 0; a < t.length; a++) {
    const s = t[a].rowOriginIds;
    if (s === void 0) continue;
    if (!TO(s))
      throw Error("Source row associations must be a plain index map.");
    const d = {};
    for (const [u, l] of Object.entries(s)) {
      if (
        !/^(0|[1-9]\d*)$/.test(u) ||
        Number(u) >= Number(t[a].rowCount) ||
        typeof l != "string" ||
        !n.has(l) ||
        ++o > Mt.maxCells
      )
        throw Error("Invalid or over-capacity source row association.");
      d[u] = l;
    }
    r.sheets[a].rowOriginIds = d;
  }
}
function U1(e, r, t) {
  var a, s, d;
  const n =
      (s =
        (a = e.sheets.find((u) => u.id === r)) == null
          ? void 0
          : a.rowOriginIds) == null
        ? void 0
        : s[t],
    o = n && ((d = e.rowOrigins) == null ? void 0 : d.find((u) => u.id === n));
  return o ? { ...o } : null;
}
function sie(e, r) {
  var n;
  if (!((n = e.rowOrigins) != null && n.length)) return;
  const t = new Map((r.rowOrigins ?? []).map((o) => [o.id, o]));
  for (const o of e.rowOrigins)
    if (
      !t.has(o.id) ||
      [
        "id",
        "sourceId",
        "sourceName",
        "sourceSheet",
        "originalRow",
        "kind",
      ].some((a) => t.get(o.id)[a] !== o[a])
    )
      throw Error(
        "Original source row identities cannot be changed or removed. Keep the source registry and change only current row associations.",
      );
}
const cR = ["sum", "count", "average", "min", "max"],
  FP = [...cR, "distinctCount"],
  qw = () => crypto.randomUUID();
function Xt(e) {
  throw new Error(e);
}
const dl = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    (Object.getPrototypeOf(e) === Object.prototype ||
      Object.getPrototypeOf(e) === null),
  xf = (e, r, t = 255) =>
    typeof e == "string" && e.trim() && e.length <= t
      ? e
      : Xt(`${r} must contain 1–${t} characters.`),
  s1 = (e) => xf(e, "Metadata ID", 100),
  Xp = (e) => {
    const r = ct(e);
    return `${ut(r.firstCol)}${r.firstRow + 1}:${ut(r.lastCol)}${r.lastRow + 1}`;
  },
  Cn = (e, r) => `${ut(r)}${e + 1}`,
  Dp = (e, r, t) =>
    r >= e.firstRow && r <= e.lastRow && t >= e.firstCol && t <= e.lastCol,
  Qf = (e, r) =>
    e.firstRow <= r.lastRow &&
    e.lastRow >= r.firstRow &&
    e.firstCol <= r.lastCol &&
    e.lastCol >= r.firstCol,
  $1 = (e) => (e.lastRow - e.firstRow + 1) * (e.lastCol - e.firstCol + 1),
  Mg = (e) => e === null || e === "",
  tg = (e) => typeof e == "number" && Number.isFinite(e),
  lie = {
    blue: ["#174a7e", "#e8f1fa"],
    green: ["#176344", "#e9f3ed"],
    gray: ["#4b5563", "#f0f1f3"],
  },
  uie = {
    blue: "TableStyleMedium2",
    green: "TableStyleMedium4",
    gray: "TableStyleMedium9",
  };
function Ef(e, r) {
  const t = ct(xf(e, "Source range", 50));
  return (
    $1(t) > Mt.maxCells &&
      Xt(
        "Table, query and pivot ranges are limited to 100,000 cells including blank gaps. Select the used data range.",
      ),
    r &&
      (t.lastRow >= r.rowCount || t.lastCol >= r.colCount) &&
      Xt("Range exceeds worksheet bounds."),
    t
  );
}
function sl(e) {
  const r = xf(e, "Column name");
  return (
    /[\[\]\r\n]/.test(r) &&
      Xt("Column names cannot contain brackets or line breaks."),
    r
  );
}
function dR(e) {
  const r = xf(e, "Table name");
  return (
    (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(r) ||
      /^[RC]$/i.test(r) ||
      /^[A-Z]{1,3}[1-9]\d*$/i.test(r) ||
      /^R\d+C\d+$/i.test(r)) &&
      Xt(
        "Use a table name beginning with a letter or underscore, without spaces or a cell-address name.",
      ),
    r
  );
}
function D5(e, r) {
  (!Array.isArray(e) || e.length > Mt.maxCols) &&
    Xt(`${r} must be a bounded list of column names.`);
  const t = e.map(sl);
  return (
    new Set(t.map((n) => n.toLowerCase())).size !== t.length &&
      Xt(`${r} contains duplicate fields.`),
    t
  );
}
function fR(e) {
  dl(e) || Xt("A source worksheet and range are required.");
  const r = {
    sheetId: s1(e.sheetId),
    range: Xp(xf(e.range, "Source range", 50)),
    ...(e.tableId !== void 0 ? { tableId: s1(e.tableId) } : {}),
  };
  return (Ef(r.range), r);
}
function BP(e) {
  return e === void 0
    ? []
    : ((!Array.isArray(e) || e.length > 100) &&
        Xt("Use at most 100 source filters."),
      e.map((r) => {
        (!dl(r) ||
          ![
            "equals",
            "notEquals",
            "contains",
            "gt",
            "gte",
            "lt",
            "lte",
            "blank",
            "notBlank",
          ].includes(String(r.operator))) &&
          Xt("Invalid source filter.");
        const t = r.value;
        return (
          t != null &&
            !(typeof t == "string" && t.length <= 32767) &&
            typeof t != "boolean" &&
            !(typeof t == "number" && Number.isFinite(t)) &&
            Xt(
              "Filter values must be finite numbers, text, booleans or blank.",
            ),
          {
            field: sl(r.field),
            operator: r.operator,
            ...(t !== void 0 ? { value: t } : {}),
          }
        );
      }));
}
function LP(e) {
  if (e === void 0) return;
  dl(e) || Xt("Invalid output destination.");
  const r = e.address === void 0 ? "A1" : xf(e.address, "Output address", 20),
    t = _t(r);
  return {
    address: Cn(t.row, t.col),
    ...(e.sheetId !== void 0 ? { sheetId: s1(e.sheetId) } : {}),
    ...(e.sheetName !== void 0
      ? { sheetName: xf(e.sheetName, "Output sheet name", 31) }
      : {}),
  };
}
function wC(e) {
  (dl(e) || Xt("Invalid pivot definition."),
    (!Array.isArray(e.values) || !e.values.length || e.values.length > 100) &&
      Xt("Choose between 1 and 100 pivot value fields."));
  const r = e.values.map(
    (o) => (
      (!dl(o) || !FP.includes(o.operation)) &&
        Xt(
          "Invalid pivot aggregation. Choose sum, count, average, min, max or distinctCount.",
        ),
      o.showAs !== void 0 &&
        !nie.includes(o.showAs) &&
        Xt("Choose a supported Show values as setting."),
      {
        field: sl(o.field),
        operation: o.operation,
        ...(o.name !== void 0 ? { name: sl(o.name) } : {}),
        ...(o.showAs !== void 0 ? { showAs: o.showAs } : {}),
      }
    ),
  );
  (e.grandTotals !== void 0 &&
    typeof e.grandTotals != "boolean" &&
    Xt("Grand totals must be boolean."),
    e.subtotals !== void 0 &&
      typeof e.subtotals != "boolean" &&
      Xt("Subtotals must be boolean."),
    e.dateGroups !== void 0 &&
      (!Array.isArray(e.dateGroups) || e.dateGroups.length > 100) &&
      Xt("Use at most 100 date groups."));
  const t = (e.dateGroups ?? []).map(
    (o) => (
      (!dl(o) ||
        !["year", "quarter", "month"].includes(String(o.unit)) ||
        !["iso", "1900", "1904"].includes(String(o.dateSystem))) &&
        Xt(
          "Date grouping requires a year, quarter or month and an explicit ISO/1900/1904 date system.",
        ),
      { field: sl(o.field), unit: o.unit, dateSystem: o.dateSystem }
    ),
  );
  (new Set(t.map((o) => o.field)).size !== t.length &&
    Xt("Choose one date grouping per source field."),
    e.calculatedFields !== void 0 &&
      (!Array.isArray(e.calculatedFields) || e.calculatedFields.length > 100) &&
      Xt("Use at most 100 calculated fields."));
  const n = (e.calculatedFields ?? []).map(
    (o) => (
      (!dl(o) ||
        !["add", "subtract", "multiply", "divide"].includes(
          String(o.operation),
        ) ||
        (o.otherField === void 0) == (o.constant === void 0) ||
        (o.constant !== void 0 &&
          (typeof o.constant != "number" || !Number.isFinite(o.constant)))) &&
        Xt(
          "A calculated field needs add/subtract/multiply/divide and exactly one numeric constant or other source field.",
        ),
      {
        name: sl(o.name),
        field: sl(o.field),
        operation: o.operation,
        ...(o.otherField !== void 0
          ? { otherField: sl(o.otherField) }
          : { constant: o.constant }),
      }
    ),
  );
  return (
    new Set(n.map((o) => o.name.toLowerCase())).size !== n.length &&
      Xt("Calculated field names must be unique."),
    {
      name: xf(e.name, "Pivot name", 100),
      source: fR(e.source),
      rowFields: D5(e.rowFields, "Row fields"),
      columnFields: D5(e.columnFields, "Column fields"),
      values: r,
      filters: BP(e.filters),
      grandTotals: e.grandTotals !== false,
      ...(e.subtotals !== void 0 ? { subtotals: e.subtotals } : {}),
      ...(e.dateGroups !== void 0 ? { dateGroups: t } : {}),
      ...(e.calculatedFields !== void 0 ? { calculatedFields: n } : {}),
      ...iie(e),
      destination: LP(e.destination),
    }
  );
}
function PP(e) {
  dl(e) || Xt("Invalid query definition.");
  const r = D5(e.fields, "Selected fields");
  r.length || Xt("Choose at least one output field.");
  const t = {};
  if (e.rename !== void 0) {
    dl(e.rename) || Xt("Rename fields must be a name map.");
    for (const [s, d] of Object.entries(e.rename))
      (r.includes(s) || Xt("Only selected fields can be renamed."),
        Object.defineProperty(t, s, {
          value: sl(d),
          enumerable: true,
          writable: true,
          configurable: true,
        }));
  }
  let n;
  (e.sort !== void 0 &&
    ((!dl(e.sort) ||
      !["ascending", "descending"].includes(String(e.sort.direction))) &&
      Xt("Invalid query sort."),
    (n = { field: sl(e.sort.field), direction: e.sort.direction })),
    e.derived !== void 0 &&
      (!Array.isArray(e.derived) || e.derived.length > 100) &&
      Xt("Use at most 100 derived fields."));
  const o = (e.derived ?? []).map(
      (s) => (
        (!dl(s) ||
          ![
            "add",
            "subtract",
            "multiply",
            "divide",
            "concat",
            "upper",
            "lower",
            "trim",
          ].includes(String(s.operation))) &&
          Xt("Invalid derived operation."),
        s.constant !== void 0 &&
          !(typeof s.constant == "string" && s.constant.length <= 32767) &&
          !(typeof s.constant == "number" && Number.isFinite(s.constant)) &&
          Xt("A derived constant must be bounded text or a finite number."),
        {
          name: sl(s.name),
          field: sl(s.field),
          operation: s.operation,
          ...(s.otherField !== void 0 ? { otherField: sl(s.otherField) } : {}),
          ...(s.constant !== void 0 ? { constant: s.constant } : {}),
        }
      ),
    ),
    a = [
      ...r.map((s) => (Object.hasOwn(t, s) ? t[s] : s)),
      ...o.map((s) => s.name),
    ];
  return (
    new Set(a.map((s) => s.toLowerCase())).size !== a.length &&
      Xt("Output column names must be unique."),
    {
      name: xf(e.name, "Query name", 100),
      source: fR(e.source),
      fields: r,
      rename: t,
      filters: BP(e.filters),
      ...(n ? { sort: n } : {}),
      derived: o,
      destination: LP(e.destination),
    }
  );
}
function Ax(e, r) {
  if (dl(e)) {
    if (e.tables !== void 0) {
      (!Array.isArray(e.tables) || e.tables.length > Mt.maxCells) &&
        Xt("Tables must be a bounded list.");
      const t = new Set(),
        n = new Set();
      r.tables = e.tables.map((s) => {
        var p;
        (!dl(s) ||
          !Array.isArray(s.columns) ||
          typeof s.showTotals != "boolean" ||
          !["blue", "green", "gray"].includes(String(s.style)) ||
          !dl(s.totals)) &&
          Xt("Invalid table metadata.");
        const d = s1(s.id),
          u = dR(s.name),
          l = Xp(xf(s.range, "Table range", 50)),
          c = Ef(l, r);
        ((p = r.merges) != null &&
          p.some((m) => Qf(c, ct(m))) &&
          Xt("A table cannot overlap merged cells. Unmerge the source first."),
          (t.has(d) || n.has(u.toLowerCase())) &&
            Xt("Table IDs and names must be unique."),
          t.add(d),
          n.add(u.toLowerCase()),
          (s.columns.length !== c.lastCol - c.firstCol + 1 ||
            (s.showTotals && c.lastRow <= c.firstRow)) &&
            Xt("Table columns or totals do not fit the table range."));
        const h = s.columns.map(
          (m) => (
            dl(m) || Xt("Invalid table column."),
            {
              id: s1(m.id),
              name: sl(m.name),
              ...(m.dataType !== void 0 ? { dataType: iR(m.dataType) } : {}),
              ...(m.calculatedFormula !== void 0
                ? { calculatedFormula: wR(m.calculatedFormula, true) }
                : {}),
            }
          ),
        );
        (new Set(h.map((m) => m.id)).size !== h.length ||
          new Set(h.map((m) => m.name.toLowerCase())).size !== h.length) &&
          Xt("Table column IDs and names must be unique.");
        const f = {};
        for (const [m, v] of Object.entries(s.totals))
          ((!h.some((g) => g.name === m) || !cR.includes(v)) &&
            Xt("Invalid table total field or operation."),
            Object.defineProperty(f, m, {
              value: v,
              enumerable: true,
              writable: true,
              configurable: true,
            }));
        return {
          id: d,
          name: u,
          range: l,
          columns: h,
          style: s.style,
          showTotals: s.showTotals,
          totals: f,
        };
      });
      const o = new Set();
      let a = 0;
      for (const s of r.tables) {
        const d = ct(s.range);
        (a += $1(d)) > Mt.maxCells &&
          Xt("The total table footprint exceeds 100,000 cells.");
        for (let u = d.firstRow; u <= d.lastRow; u++)
          for (let l = d.firstCol; l <= d.lastCol; l++) {
            const c = Cn(u, l);
            (o.has(c) && Xt("Tables cannot overlap."), o.add(c));
          }
      }
    }
    for (const t of ["pivots", "queries"]) {
      if (e[t] === void 0) continue;
      (!Array.isArray(e[t]) || e[t].length > Mt.maxCells) &&
        Xt("Analysis definitions must be a bounded list.");
      const n = new Set(),
        o = e[t].map((a) => {
          dl(a) || Xt("Invalid analysis metadata.");
          const s = t === "pivots" ? wC(a) : PP(a),
            d = s1(a.id),
            u = Xp(xf(a.outputRange, "Output range", 50));
          (Ef(u, r),
            (n.has(d) ||
              !Number.isSafeInteger(a.sourceRevision) ||
              Number(a.sourceRevision) < 0) &&
              Xt("Invalid analysis identity or source revision."),
            n.add(d));
          const { destination: l, ...c } = s;
          return {
            ...c,
            id: d,
            outputRange: u,
            sourceRevision: Number(a.sourceRevision),
            ...(t === "pivots" && a.refresh !== void 0
              ? { refresh: tie(a.refresh) }
              : {}),
          };
        });
      t === "pivots" ? (r.pivots = o) : (r.queries = o);
    }
  }
}
function UP(e) {
  const r = new Set(),
    t = new Set();
  for (const n of e.sheets)
    for (const o of n.tables ?? [])
      ((r.has(o.name.toLowerCase()) || t.has(o.id)) &&
        Xt("Table names and IDs must be unique across the workbook."),
        r.add(o.name.toLowerCase()),
        t.add(o.id));
}
function cie(e, r, t, n, o, a) {
  const s = t === "row" ? Mt.maxRows : Mt.maxCols;
  ((!Number.isInteger(n) ||
    n < 0 ||
    n >= s ||
    !Number.isInteger(o) ||
    o < 1 ||
    n + o > s) &&
    Xt("Invalid structural edit bounds."),
    e.sheets.some((l) => {
      var c;
      return (c = l.tables) == null
        ? void 0
        : c.some((h) => h.columns.some((f) => f.calculatedFormula));
    }) &&
      Xt(
        "Calculated table formulas require explicit table append/resize. Clear their saved rules (keeping cell formulas) before inserting or deleting worksheet rows/columns, then reapply the rules.",
      ));
  const d = structuredClone(e);
  function u(l, c, h = false) {
    const f = Ef(l),
      p = t === "row" ? f.firstRow : f.firstCol,
      m = t === "row" ? f.lastRow : f.lastCol;
    let v = p,
      g = m;
    if (a) {
      const b = n + o - 1;
      (n <= p &&
        b >= p &&
        Xt(
          `This edit would remove a ${c} header. Move or recreate that structure first.`,
        ),
        b < p
          ? ((v -= o), (g -= o))
          : n <= m &&
            (t === "column" &&
              Xt(
                `Deleting columns inside a ${c} would invalidate its fields. Resize or recreate it first.`,
              ),
            h &&
              b >= m &&
              Xt(
                "This edit would remove a table totals row. Turn off table totals first.",
              ),
            (g -= Math.min(m, b) - n + 1)));
    } else
      n <= p
        ? ((v += o), (g += o))
        : n <= m &&
          (t === "column" &&
            Xt(
              `Inserting columns inside a ${c} requires explicit field reconciliation. Resize or recreate it first.`,
            ),
          (g += o));
    return (
      g >= s &&
        Xt(
          `The structural edit would push a ${c} beyond the worksheet boundary.`,
        ),
      t === "row"
        ? ((f.firstRow = v), (f.lastRow = g))
        : ((f.firstCol = v), (f.lastCol = g)),
      $1(f) > Mt.maxCells &&
        Xt(
          "The shifted table or analysis exceeds the 100,000-cell range budget.",
        ),
      `${Cn(f.firstRow, f.firstCol)}:${Cn(f.lastRow, f.lastCol)}`
    );
  }
  for (const l of d.sheets) {
    if (l.id === r) {
      for (const c of l.tables ?? [])
        c.range = u(c.range, "table", c.showTotals);
      for (const c of [...(l.pivots ?? []), ...(l.queries ?? [])])
        c.outputRange = u(c.outputRange, "analysis result");
    }
    for (const c of [...(l.pivots ?? []), ...(l.queries ?? [])])
      c.source.sheetId === r &&
        (c.source.range = u(c.source.range, "analysis source"));
  }
  return d;
}
function xC(e) {
  return (
    e.sheets.reduce((r, t) => r + Object.keys(t.cells).length, 0) >
      Mt.maxCells &&
      Xt("The result would exceed the 100,000 populated-cell workbook budget."),
    e
  );
}
function Gp(e) {
  const r = ct(e.range);
  return {
    ...r,
    firstRow: r.firstRow + 1,
    lastRow: r.lastRow - (e.showTotals ? 1 : 0),
  };
}
function z1(e, r, t, n) {
  var m, v, g;
  const o = /^([A-Za-z_][A-Za-z0-9_.]*)?(\[.+\])$/.exec(t);
  if (!o) return null;
  let a;
  if (o[1])
    for (const b of e.sheets) {
      const y =
        (m = b.tables) == null
          ? void 0
          : m.find((w) => w.name.toLowerCase() === o[1].toLowerCase());
      if (y) {
        a = { sheet: b, table: y };
        break;
      }
    }
  else if (n) {
    const b = e.sheets.find((E) => E.id === r),
      y = _t(n),
      w =
        (v = b == null ? void 0 : b.tables) == null
          ? void 0
          : v.find((E) => Dp(ct(E.range), y.row, y.col));
    b && w && (a = { sheet: b, table: w });
  }
  if (!a) return null;
  const { sheet: s, table: d } = a,
    u = ct(d.range),
    l = Gp(d),
    c = o[2]
      .replace(/^\[|\]$/g, "")
      .replace(/[\[\]]/g, "")
      .split(/[:,]/)
      .map((b) => b.trim()),
    h =
      (g = c.find((b) => b.startsWith("#"))) == null ? void 0 : g.toLowerCase(),
    f = c.some((b) => b.startsWith("@")) || h === "#this row";
  if (
    c.some(
      (b) =>
        b.startsWith("#") &&
        !["#all", "#data", "#headers", "#totals", "#this row"].includes(
          b.toLowerCase(),
        ),
    )
  )
    return null;
  if (
    (h === "#all" && Object.assign(l, u),
    h === "#headers" && Object.assign(l, u, { lastRow: u.firstRow }),
    h === "#totals")
  ) {
    if (!d.showTotals) return null;
    Object.assign(l, u, { firstRow: u.lastRow });
  }
  if (f) {
    if (!n || s.id !== r) return null;
    const b = _t(n);
    if (b.row < l.firstRow || b.row > l.lastRow) return null;
    ((l.firstRow = b.row), (l.lastRow = b.row));
  }
  const p = c.filter((b) => !b.startsWith("#")).map((b) => b.replace(/^@/, ""));
  if (p.length) {
    if (p.length > 2) return null;
    const b = p.map((y) =>
      d.columns.findIndex((w) => w.name.toLowerCase() === y.toLowerCase()),
    );
    if (b.some((y) => y < 0)) return null;
    ((l.firstCol = u.firstCol + Math.min(...b)),
      (l.lastCol = u.firstCol + Math.max(...b)));
  }
  return l.firstRow > l.lastRow
    ? null
    : {
        sheetId: s.id,
        range: `${Cn(l.firstRow, l.firstCol)}:${Cn(l.lastRow, l.lastCol)}`,
      };
}
function die(e, r, t) {
  var s;
  const n = Nr(e, r),
    o = ct(t);
  if ($1(o) !== 1) return Xp(t);
  const a =
    (s = n.tables) == null
      ? void 0
      : s.find((d) => Dp(ct(d.range), o.firstRow, o.firstCol));
  return a ? a.range : Ta(n, false);
}
function $P(e, r, t) {
  const n = Nr(e, r),
    o = Ef(t, n),
    a = new Set(),
    s = [];
  for (let d = o.firstCol; d <= o.lastCol; d++) {
    const u = Cn(o.firstRow, d),
      l = n.cells[u],
      c = (l == null ? void 0 : l.raw) ?? null;
    let h = `Column${d - o.firstCol + 1}`;
    try {
      Mg(c) || (h = sl(String(c)));
    } catch {}
    let f = h,
      p = 2;
    for (; a.has(f.toLowerCase());) f = `${h.slice(0, 250)}${p++}`;
    (a.add(f.toLowerCase()),
      s.push({
        address: u,
        original: c,
        suggested: f,
        requiresResolution:
          c !== f || (f.startsWith("=") && !(l != null && l.literal)),
      }));
  }
  return s;
}
function zP(e, r, t, n) {
  const o = $P(e, r, t);
  if (n === void 0)
    return (
      o.some((s) => s.requiresResolution) &&
        Xt(
          "Review table headers: resolve blank, duplicate, non-text or formula headers explicitly. Original header values will be retained in native source evidence.",
        ),
      o.map((s) => s.suggested)
    );
  (!Array.isArray(n) || n.length !== o.length) &&
    Xt("Provide one reviewed header for every table column.");
  const a = n.map(sl);
  return (
    new Set(a.map((s) => s.toLowerCase())).size !== a.length &&
      Xt("Reviewed table headers must be unique, nonblank text."),
    a
  );
}
function kx(e, r) {
  const t = Gp(e),
    n = e.columns[r].name,
    o = Object.hasOwn(e.totals, n) ? e.totals[n] : void 0;
  return o
    ? t.lastRow < t.firstRow
      ? { raw: o === "count" || o === "sum" ? 0 : null }
      : {
          raw: `=${o === "count" ? "COUNTA" : o.toUpperCase()}(${Cn(t.firstRow, t.firstCol + r)}:${Cn(t.lastRow, t.firstCol + r)})`,
        }
    : { raw: r === 0 ? "Total" : null };
}
function VP(e, r) {
  const t = Ef(r.range, e),
    [n, o] = lie[r.style];
  for (let a = t.firstRow; a <= t.lastRow; a++)
    for (let s = t.firstCol; s <= t.lastCol; s++) {
      const d = Cn(a, s),
        u = e.cells[d] ?? { raw: null },
        l = a === t.firstRow,
        c = r.showTotals && a === t.lastRow,
        h = l
          ? {
              ...u,
              raw: r.columns[s - t.firstCol].name,
              literal: r.columns[s - t.firstCol].name.startsWith("=")
                ? true
                : void 0,
            }
          : c
            ? { ...u, ...kx(r, s - t.firstCol), literal: void 0 }
            : u;
      e.cells[d] = {
        ...h,
        style: {
          ...h.style,
          bold: l || c,
          color: l ? "#ffffff" : "#222222",
          fill: l ? n : c || (a - t.firstRow) % 2 === 0 ? o : "#ffffff",
          ...(c ? { border: "bottom" } : {}),
        },
      };
    }
}
function fie(e, r, t) {
  var c, h;
  const n = Nr(e, r),
    o = dR(t.name),
    a = Ef(t.range, n);
  if (
    (jn(n),
    (c = n.merges) != null &&
      c.some((f) => Qf(a, ct(f))) &&
      Xt("A table cannot overlap merged cells. Unmerge the source first."),
    e.sheets.some((f) => {
      var p;
      return (p = f.tables) == null
        ? void 0
        : p.some((m) => m.name.toLowerCase() === o.toLowerCase());
    }) && Xt("A table with that name already exists."),
    t.showTotals)
  ) {
    a.lastRow + 1 >= n.rowCount && Xt("There is no room for a totals row.");
    for (let f = a.firstCol; f <= a.lastCol; f++)
      n.cells[Cn(a.lastRow + 1, f)] &&
        Xt("The totals row would overwrite existing cells.");
    a.lastRow++;
  }
  ($1(a) > Mt.maxCells ||
    ((h = n.tables) != null && h.some((f) => Qf(a, ct(f.range))))) &&
    Xt("The table overlaps another table or exceeds the cell budget.");
  const s = structuredClone(e),
    d = Nr(s, r),
    u = {
      id: qw(),
      name: o,
      range: `${Cn(a.firstRow, a.firstCol)}:${Cn(a.lastRow, a.lastCol)}`,
      columns: zP(e, r, t.range, t.headers).map((f) => ({ id: qw(), name: f })),
      style: t.style ?? "blue",
      showTotals: !!t.showTotals,
      totals: t.totals ?? {},
    };
  ((d.tables = [...(d.tables ?? []), u]),
    Ax(d, d),
    VP(d, d.tables[d.tables.length - 1]));
  const l = u.columns.flatMap((f, p) => {
    const m = Cn(a.firstRow, a.firstCol + p),
      v = n.cells[m],
      g = d.cells[m];
    return (v == null ? void 0 : v.raw) === f.name &&
      !(f.name.startsWith("=") && !(v != null && v.literal))
      ? []
      : [
          {
            address: m,
            ...(v ? { before: structuredClone(v) } : {}),
            after: structuredClone(g),
          },
        ];
  });
  return (
    l.length &&
      ((s.tableEvidence = [
        ...(s.tableEvidence ?? []),
        {
          id: `${u.id}:headers`,
          kind: "headers",
          sourceRevision: e.revision,
          sheetId: r,
          sheetName: n.name,
          tableId: u.id,
          tableName: o,
          changes: l,
        },
      ]),
      bC(s.tableEvidence, s)),
    xC(s)
  );
}
function hie(e, r, t, n) {
  var y, w, E;
  const o = Nr(e, r),
    a =
      ((y = o.tables) == null ? void 0 : y.find((A) => A.id === t)) ??
      Xt("Table not found.");
  jn(o);
  const s = dR(n.name),
    d = Ef(n.range, o),
    u = ct(a.range);
  ((d.firstRow !== u.firstRow || d.firstCol !== u.firstCol) &&
    Xt(
      "Resize a table from its existing top-left header. Moving table source cells is a separate operation.",
    ),
    e.sheets.some((A) => {
      var k;
      return (k = A.tables) == null
        ? void 0
        : k.some((F) => F.id !== t && F.name.toLowerCase() === s.toLowerCase());
    }) && Xt("A table with that name already exists."));
  const l = n.showTotals ?? a.showTotals;
  if (
    (l !== a.showTotals && Xp(n.range) === a.range && (d.lastRow += l ? 1 : -1),
    (d.lastRow >= o.rowCount ||
      d.lastRow < d.firstRow ||
      $1(d) > Mt.maxCells) &&
      Xt("The resized table does not fit the worksheet or processing budget."),
    (w = o.tables) != null &&
      w.some((A) => A.id !== t && Qf(d, ct(A.range))) &&
      Xt("Tables cannot overlap."),
    l && (!a.showTotals || d.lastRow !== u.lastRow))
  )
    for (let A = d.firstCol; A <= d.lastCol; A++)
      o.cells[Cn(d.lastRow, A)] &&
        Xt(
          "The new totals row would overwrite existing cells. Include an empty final row in the resized range.",
        );
  const c = a.columns.some((A) => A.calculatedFormula);
  (c &&
    d.lastCol !== u.lastCol &&
    Xt("Clear calculated column rules before changing table columns."),
    c && (d.lastRow !== u.lastRow || l !== a.showTotals) && YP(e, r, a, d, l));
  const h = structuredClone(e),
    f = Nr(h, r);
  if (a.showTotals && (!l || d.lastRow !== u.lastRow))
    for (let A = u.firstCol; A <= u.lastCol; A++) {
      const k = Cn(u.lastRow, A),
        F = f.cells[k],
        I = kx(a, A - u.firstCol);
      (F &&
        F.raw !== I.raw &&
        Xt(
          "The old totals row was edited. Preserve or move those values before resizing it.",
        ),
        delete f.cells[k]);
    }
  const p = zP(h, r, n.range, n.headers),
    m = p.map((A, k) => {
      var F;
      return {
        ...a.columns[k],
        id:
          ((F = a.columns[k]) == null ? void 0 : F.id) ??
          `${a.id}:column:${k}:${e.revision}`,
        name: A,
      };
    }),
    v =
      n.totals ??
      Object.fromEntries(
        Object.entries(a.totals).filter(([A]) => m.some((k) => k.name === A)),
      ),
    g = {
      ...a,
      name: s,
      range: `${Cn(d.firstRow, d.firstCol)}:${Cn(d.lastRow, d.lastCol)}`,
      columns: m,
      style: n.style ?? a.style,
      showTotals: l,
      totals: v,
    };
  ((f.tables = f.tables.map((A) => (A.id === t ? g : A))), Ax(f, f), VP(f, g));
  const b = m.flatMap((A, k) => {
    const F = Cn(d.firstRow, d.firstCol + k),
      I = o.cells[F],
      O = f.cells[F];
    return (I == null ? void 0 : I.raw) === O.raw &&
      !!(I != null && I.literal) == !!O.literal
      ? []
      : [
          {
            address: F,
            ...(I ? { before: structuredClone(I) } : {}),
            after: structuredClone(O),
          },
        ];
  });
  if (
    (b.length &&
      ((h.tableEvidence = [
        ...(h.tableEvidence ?? []),
        {
          id: `${t}:headers:${e.revision}:${((E = h.tableEvidence) == null ? void 0 : E.length) ?? 0}`,
          kind: "headers",
          sourceRevision: e.revision,
          sheetId: r,
          sheetName: o.name,
          tableId: t,
          tableName: s,
          changes: b,
        },
      ]),
      bC(h.tableEvidence, h)),
    c)
  ) {
    const A = Gp(a),
      k = Gp(g);
    for (
      let F = u.firstRow;
      d.lastRow !== u.lastRow && F <= Math.min(A.lastRow, k.lastRow);
      F++
    )
      for (let I = A.firstCol; I <= A.lastCol; I++) {
        const O = Cn(F, I);
        o.cells[O] && (f.cells[O] = structuredClone(o.cells[O]));
      }
    (a.showTotals && l && d.lastRow !== u.lastRow && JP(e, h, r, a, g),
      KP(f, g, Math.max(A.lastRow + 1, k.firstRow), k.lastRow));
  }
  return xC(hR(e, h));
}
function hR(e, r) {
  var o, a;
  const t = [];
  for (const s of e.sheets) {
    const d = r.sheets.find((u) => u.id === s.id);
    if (d)
      for (const u of s.tables ?? []) {
        const l =
          (o = d.tables) == null ? void 0 : o.find((m) => m.id === u.id);
        if (!l) continue;
        const c = ct(u.range),
          h = ct(l.range),
          f = {},
          p = l.columns.map((m, v) => {
            const g = u.columns.findIndex((k) => k.id === m.id);
            if (g < 0) return m;
            const b = u.columns[g],
              y = s.cells[Cn(c.firstRow, c.firstCol + g)],
              w = d.cells[Cn(h.firstRow, h.firstCol + v)],
              A =
                (y == null ? void 0 : y.raw) !== (w == null ? void 0 : w.raw) ||
                (y == null ? void 0 : y.literal) !==
                  (w == null ? void 0 : w.literal)
                  ? w && !Mg(w.raw)
                    ? sl(String(w.raw))
                    : Xt(
                        "Table headers must be nonblank text. The header edit was not applied.",
                      )
                  : m.name;
            return (
              A !== b.name &&
                Object.defineProperty(f, b.name, {
                  value: A,
                  enumerable: true,
                }),
              A === m.name ? m : { ...m, name: A }
            );
          });
        (u.name !== l.name || Object.keys(f).length) &&
          (new Set(p.map((m) => m.name.toLowerCase())).size !== p.length &&
            Xt(
              "Table headers must be unique, nonblank text. The header edit was not applied.",
            ),
          t.push({ sheetId: d.id, old: u, current: l, columns: p, names: f }));
      }
  }
  if (!t.length) return r;
  const n = structuredClone(r);
  for (const s of t) {
    const d = Nr(n, s.sheetId),
      u = d.tables.find((f) => f.id === s.current.id),
      l = (f) => (Object.hasOwn(s.names, f) ? s.names[f] : f);
    ((u.columns = s.columns),
      (u.totals = Object.fromEntries(
        Object.entries(u.totals).map(([f, p]) => [l(f), p]),
      )));
    const c = ct(u.range);
    u.columns.forEach((f, p) => {
      const m = Cn(c.firstRow, c.firstCol + p);
      d.cells[m] = {
        ...(d.cells[m] ?? { raw: f.name }),
        raw: f.name,
        ...(f.name.startsWith("=") ? { literal: true } : { literal: void 0 }),
      };
    });
    const h = (f) =>
      f.tableId === u.id ||
      (!f.tableId && f.sheetId === d.id && Xp(f.range) === s.old.range);
    for (const f of n.sheets) {
      for (const p of f.tables ?? [])
        for (const m of p.columns)
          m.calculatedFormula &&
            (m.calculatedFormula = $v(
              m.calculatedFormula,
              s.old.name,
              u.name,
              s.names,
              p.id === u.id && f.id === d.id,
            ));
      for (const [p, m] of Object.entries(f.cells)) {
        if (typeof m.raw != "string" || !m.raw.startsWith("=") || m.literal)
          continue;
        const v = _t(p),
          g = f.id === d.id && Dp(c, v.row, v.col);
        m.raw = $v(m.raw, s.old.name, u.name, s.names, g);
      }
      for (const p of f.conditionalFormats ?? [])
        if (p.kind === "formula" && p.formula) {
          const m = ct(p.range),
            v = $v(p.formula, s.old.name, u.name, s.names, false),
            g = $v(p.formula, s.old.name, u.name, s.names, true),
            b = f.id === d.id && Qf(c, m),
            y =
              b && Dp(c, m.firstRow, m.firstCol) && Dp(c, m.lastRow, m.lastCol);
          (v !== g &&
            b &&
            !y &&
            Xt(
              "This conditional formula uses unqualified table fields across a table boundary. Split or remove its applied range before renaming the field.",
            ),
            (p.formula = y ? g : v));
        }
      for (const p of [...(f.pivots ?? []), ...(f.queries ?? [])])
        if (h(p.source))
          if (
            ((p.source = { ...p.source, tableId: u.id, range: u.range }),
            p.filters &&
              (p.filters = p.filters.map((m) => ({ ...m, field: l(m.field) }))),
            "rowFields" in p)
          ) {
            ((p.rowFields = p.rowFields.map(l)),
              (p.columnFields = p.columnFields.map(l)),
              (p.values = p.values.map((m) => ({ ...m, field: l(m.field) }))),
              p.dateGroups &&
                (p.dateGroups = p.dateGroups.map((m) => ({
                  ...m,
                  field: l(m.field),
                }))));
            for (const m of ["slicers", "timelines", "sorts"])
              p[m] && (p[m] = p[m].map((v) => ({ ...v, field: l(v.field) })));
            p.calculatedFields &&
              (p.calculatedFields = p.calculatedFields.map((m) => ({
                ...m,
                field: l(m.field),
                ...(m.otherField ? { otherField: l(m.otherField) } : {}),
              })));
          } else
            ((p.fields = p.fields.map(l)),
              p.rename &&
                (p.rename = Object.fromEntries(
                  Object.entries(p.rename).map(([m, v]) => [l(m), v]),
                )),
              p.sort && (p.sort = { ...p.sort, field: l(p.sort.field) }),
              p.derived &&
                (p.derived = p.derived.map((m) => ({
                  ...m,
                  field: l(m.field),
                  ...(m.otherField ? { otherField: l(m.otherField) } : {}),
                }))));
    }
    for (const f of n.definedNames ?? [])
      f.refersTo = $v(f.refersTo, s.old.name, u.name, s.names);
    Ax(d, d);
  }
  UP(n);
  for (const s of e.sheets) {
    if (!((a = s.protection) != null && a.enabled)) continue;
    const d = n.sheets.find((u) => u.id === s.id);
    JSON.stringify(s) !== JSON.stringify(d) &&
      Xt(
        `This change would update dependent cells or analysis on protected worksheet ${s.name}. Unprotect it before renaming the table field.`,
      );
  }
  return n;
}
function V1(e, r) {
  var l, c;
  const t = fR(r),
    n = Nr(e, t.sheetId),
    o = t.tableId
      ? (((l = n.tables) == null
          ? void 0
          : l.find((h) => h.id === t.tableId)) ??
        Xt("The source table no longer exists."))
      : (c = n.tables) == null
        ? void 0
        : c.find((h) => h.range === Xp(t.range)),
    a = Ef((o == null ? void 0 : o.range) ?? t.range, n),
    s = Vn(e);
  o != null && o.showTotals && a.lastRow--;
  const d = [];
  for (let h = a.firstCol; h <= a.lastCol; h++)
    d.push(sl(s(n.id, Cn(a.firstRow, h))));
  new Set(d.map((h) => h.toLowerCase())).size !== d.length &&
    Xt(
      "Source headers must be nonblank, unique text. Create a table to normalize blank or duplicate headers.",
    );
  const u = [];
  for (let h = a.firstRow + 1; h <= a.lastRow; h++) {
    const f = d.map((p, m) => s(n.id, Cn(h, a.firstCol + m)));
    (f.forEach((p, m) => {
      const v = Cn(h, a.firstCol + m);
      typeof p == "string" &&
        p.startsWith("#") &&
        s.evaluateFormula(n.id, v, `=ISERROR(${v})`) === true &&
        Xt(
          `Source calculation error ${p} at ${n.name}!${v}. Repair the source calculation before building or refreshing analysis; no totals were written.`,
        );
    }),
      (f.some((p) => !Mg(p)) ||
        d.some((p, m) => {
          var v;
          return !Mg(
            ((v = n.cells[Cn(h, a.firstCol + m)]) == null ? void 0 : v.raw) ??
              null,
          );
        })) &&
        u.push({ index: h, values: f }));
  }
  return (
    s.getOperationalErrors().length &&
      Xt(
        `Source calculation is incomplete (${s.getOperationalErrors().join(", ")}). Repair blocked or over-limit dependencies before building analysis.`,
      ),
    {
      fields: d,
      rows: u,
      source: {
        ...t,
        range: (o == null ? void 0 : o.range) ?? Xp(t.range),
        ...(o ? { tableId: o.id } : {}),
      },
    }
  );
}
function pR(e, r, t) {
  return V1(e, { sheetId: r, range: t }).fields;
}
function Yv(e, r) {
  const t = e.indexOf(r);
  return t >= 0
    ? t
    : Xt(`Source field “${r}” was not found. Refresh the field selection.`);
}
function mR(e, r) {
  const t = r.map((n) => ({ ...n, index: Yv(e.fields, n.field) }));
  return e.rows.filter((n) =>
    t.every((o) => {
      const a = n.values[o.index],
        s = o.value ?? null;
      switch (o.operator) {
        case "blank":
          return Mg(a);
        case "notBlank":
          return !Mg(a);
        case "contains":
          return String(a ?? "")
            .toLocaleLowerCase()
            .includes(String(s ?? "").toLocaleLowerCase());
        case "equals":
          return (
            String(a ?? "").toLocaleLowerCase() ===
            String(s ?? "").toLocaleLowerCase()
          );
        case "notEquals":
          return (
            String(a ?? "").toLocaleLowerCase() !==
            String(s ?? "").toLocaleLowerCase()
          );
        default:
          return (
            tg(s) || Xt("Numeric comparisons require a numeric filter value."),
            tg(a)
              ? o.operator === "gt"
                ? a > s
                : o.operator === "gte"
                  ? a >= s
                  : o.operator === "lt"
                    ? a < s
                    : a <= s
              : false
          );
      }
    }),
  );
}
function pie(e, r) {
  const t =
    r
      .replace(/[\\/\[\]*?:]/g, " ")
      .trim()
      .slice(0, 31) || "Analysis";
  let n = t,
    o = 2;
  for (; e.sheets.some((a) => a.name.toLowerCase() === n.toLowerCase());)
    n = `${t.slice(0, 25)} (${o++})`;
  return n;
}
function gR(e, r, t, n) {
  var h, f, p, m, v;
  (n != null && n.sheetId) || $a(e);
  const o = ((h = t[0]) == null ? void 0 : h.length) ?? 0;
  ((!o || t.length * o > Mt.maxCells) &&
    Xt("The analysis output exceeds 100,000 cells."),
    t.some((g) => g.some((b) => typeof b == "string" && b.length > 32767)) &&
      Xt(
        "An analysis label exceeds the 32,767-character cell limit. Shorten source labels or value-field names.",
      ));
  const a = structuredClone(e),
    s =
      n != null && n.sheetId
        ? Nr(a, n.sheetId)
        : Ku(pie(a, (n == null ? void 0 : n.sheetName) ?? r));
  ((n != null && n.sheetId) || a.sheets.push(s), jn(s));
  const d = _t((n == null ? void 0 : n.address) ?? "A1"),
    u = d.row + t.length - 1,
    l = d.col + o - 1;
  (u >= s.rowCount || l >= s.colCount) &&
    Xt("The output does not fit the destination worksheet.");
  const c = { firstRow: d.row, lastRow: u, firstCol: d.col, lastCol: l };
  for (const g of Object.keys(s.cells)) {
    const b = _t(g);
    Dp(c, b.row, b.col) &&
      Xt(
        "The output destination contains existing cells. Choose an empty destination or a new sheet.",
      );
  }
  return (
    (f = s.tables) != null &&
      f.some((g) => Qf(c, ct(g.range))) &&
      Xt("The output destination overlaps a table."),
    (((p = s.merges) != null && p.some((g) => Qf(c, ct(g)))) ||
      ((m = s.pivots) != null && m.some((g) => Qf(c, ct(g.outputRange)))) ||
      ((v = s.queries) != null && v.some((g) => Qf(c, ct(g.outputRange))))) &&
      Xt(
        "The output destination overlaps an existing merged area or analysis result. Choose another empty destination.",
      ),
    t.forEach((g, b) =>
      g.forEach((y, w) => {
        s.cells[Cn(d.row + b, d.col + w)] = {
          raw: y,
          ...(typeof y == "string" && y.startsWith("=")
            ? { literal: true }
            : {}),
          ...(b === 0
            ? { style: { bold: true, fill: "#174a7e", color: "#ffffff" } }
            : tg(y)
              ? { format: "number" }
              : {}),
        };
      }),
    ),
    xC(a),
    { workbook: a, sheet: s, range: `${Cn(d.row, d.col)}:${Cn(u, l)}` }
  );
}
const WP = (e) => {
  const r = JSON.stringify(e);
  r.length > 16e6 &&
    Xt(
      "Pivot refresh evidence exceeds the 16-million-character fingerprint budget. Use a smaller source.",
    );
  const t = [];
  for (let n = 0; n < r.length; n += 5e5)
    t.push(FE(new Uint8Array(strToU8(r.slice(n, n + 5e5)))));
  return FE(new Uint8Array(strToU8(JSON.stringify([r.length, ...t]))));
};
function vR(e, r, t) {
  const n = Nr(e, t.source.sheetId),
    o = ct(t.source.range),
    a = Object.entries(n.cells)
      .filter(([u]) => {
        const l = _t(u);
        return Dp(o, l.row, l.col);
      })
      .sort(([u], [l]) => u.localeCompare(l))
      .map(([u, l]) => [u, l.raw, !!l.literal]),
    { destination: s, ...d } = wC(r);
  return WP({
    source: t.source,
    fields: t.fields,
    rows: t.rows,
    cells: a,
    settings: d,
  });
}
function bR(e, r) {
  const t = ct(r);
  return WP(
    Object.entries(e.cells)
      .filter(([n]) => {
        const o = _t(n);
        return Dp(t, o.row, o.col);
      })
      .sort(([n], [o]) => n.localeCompare(o))
      .map(([n, o]) => [n, o.raw, !!o.literal]),
  );
}
function HP(e, r) {
  var a;
  const t = e.sheets.find((s) => {
      var d;
      return (d = s.pivots) == null ? void 0 : d.some((u) => u.id === r);
    }),
    n =
      (a = t == null ? void 0 : t.pivots) == null
        ? void 0
        : a.find((s) => s.id === r);
  if (((!t || !n) && Xt("Pivot definition not found."), !n.refresh))
    return {
      state: "unverified",
      reason:
        "This older result has no refresh evidence. Refresh it before relying on its totals.",
      outputModified: false,
      definition: n,
    };
  const o = bR(t, n.outputRange) !== n.refresh.outputFingerprint;
  try {
    const s = V1(e, n.source);
    uR(n, s.fields);
    const d = vR(e, n, s) !== n.refresh.sourceFingerprint;
    return {
      state: d || o ? "stale" : "current",
      reason: o
        ? "Output cells were edited. Refresh into a new result before relying on these totals."
        : d
          ? "Source content, calculated values, range or settings changed. Refresh into a new result."
          : "Source content and output match the saved refresh snapshot.",
      outputModified: o,
      definition: n,
    };
  } catch (s) {
    return {
      state: "unavailable",
      reason: s.message,
      outputModified: o,
      definition: n,
    };
  }
}
function yR(e, r) {
  var c, h;
  const t = wC(r),
    n = V1(e, t.source),
    o = jP(t, n.fields, mR(n, t.filters ?? [])),
    a = OP(t, n.fields, o),
    s = gR(e, t.name, a.matrix, t.destination),
    { destination: d, ...u } = t,
    l = _t(((c = t.destination) == null ? void 0 : c.address) ?? "A1");
  for (let f = 1; f < a.matrix.length; f++)
    for (let p = Math.max(1, t.rowFields.length); p < a.matrix[0].length; p++) {
      const m =
        t.values[(p - Math.max(1, t.rowFields.length)) % t.values.length];
      m.showAs &&
        m.showAs !== "value" &&
        (s.sheet.cells[Cn(l.row + f, l.col + p)].format = "percent");
    }
  if (!((h = t.destination) != null && h.sheetId)) {
    let f = 26;
    (a.matrix[0].forEach((p, m) => {
      const v = String(p ?? ""),
        g = Math.min(320, Math.max(90, v.length * 7 + 24));
      s.sheet.columnWidths = { ...s.sheet.columnWidths, [l.col + m]: g };
      const b = Cn(l.row, l.col + m),
        y = s.sheet.cells[b];
      ((y.style = { ...y.style, wrap: true }),
        (f = Math.max(
          f,
          Math.min(120, Math.ceil((v.length * 7) / (g - 16)) * 18 + 8),
        )));
    }),
      (s.sheet.rowHeights = { ...s.sheet.rowHeights, [l.row]: f }));
  }
  for (const f of a.subtotalRows)
    for (let p = 0; p < a.matrix[0].length; p++) {
      const m = Cn(l.row + f, l.col + p),
        v = s.sheet.cells[m];
      v.style = { ...v.style, bold: true, fill: "#e4edf7" };
    }
  return (
    (s.sheet.pivots = [
      ...(s.sheet.pivots ?? []),
      {
        ...u,
        source: n.source,
        id: qw(),
        outputRange: s.range,
        sourceRevision: e.revision,
        refresh: {
          at: new Date().toISOString(),
          sourceSheetName: Nr(e, n.source.sheetId).name,
          sourceRange: n.source.range,
          sourceFingerprint: vR(e, { ...t, source: n.source }, n),
          outputFingerprint: bR(s.sheet, s.range),
          sourceRows: n.rows.length,
          includedRows: o.length,
          excludedRows: n.rows.length - o.length,
          measures: a.measures,
        },
      },
    ]),
    s.workbook
  );
}
function mie(e, r, t) {
  const n =
    e.sheets.flatMap((o) => o.pivots ?? []).find((o) => o.id === r) ??
    Xt("Pivot definition not found.");
  return yR(e, {
    ...n,
    destination: t,
    name: `${n.name.slice(0, 90)} refreshed`,
  });
}
function Sx(e, r) {
  var o;
  const t = e.sheets.find((a) => {
      var s;
      return (s = a.pivots) == null ? void 0 : s.some((d) => d.id === r);
    }),
    n =
      (o = t == null ? void 0 : t.pivots) == null
        ? void 0
        : o.find((a) => a.id === r);
  return (
    (!t || !n) && Xt("Pivot definition not found."),
    { owner: t, definition: n }
  );
}
function gie(e, r) {
  const { definition: t } = Sx(e, r),
    n = HP(e, r);
  return qP(e, t, n);
}
function qP(e, r, t) {
  var n, o, a;
  return {
    workbookId: e.id,
    revision: e.revision,
    pivotId: r.id,
    name: r.name,
    source: structuredClone(r.source),
    sourceRevision: r.sourceRevision,
    refreshedAt: ((n = r.refresh) == null ? void 0 : n.at) ?? null,
    sourceFingerprint:
      ((o = r.refresh) == null ? void 0 : o.sourceFingerprint) ?? null,
    status: t.state,
    diagnostic: t.reason,
    outputRange: r.outputRange,
    rowFields: [...r.rowFields],
    columnFields: [...r.columnFields],
    values: structuredClone(r.values),
    filters: structuredClone(r.filters ?? []),
    slicers: structuredClone(r.slicers ?? []),
    timelines: structuredClone(r.timelines ?? []),
    sorts: structuredClone(r.sorts ?? []),
    dateGroups: structuredClone(r.dateGroups ?? []),
    calculatedFields: structuredClone(r.calculatedFields ?? []),
    chart: structuredClone(r.chart ?? null),
    population: r.refresh
      ? {
          source: r.refresh.sourceRows,
          included: r.refresh.includedRows,
          excluded: r.refresh.excludedRows,
        }
      : null,
    measures: structuredClone(
      ((a = r.refresh) == null ? void 0 : a.measures) ?? [],
    ),
    semantics:
      "Controls intersect. Slicer items use type-sensitive, case-insensitive equality; an empty selection includes no records. Hidden source rows are included. Percentages use filtered aggregate totals. Drill-through returns numerator records, including blank/nonnumeric measure inputs; it does not expand denominator populations.",
  };
}
function nk(e, r, t) {
  const n = XP(e, r, [t]).get(t);
  return (n.error && Xt(n.error), n.items);
}
function XP(e, r, t) {
  t.length > 30 && Xt("Inspect at most 30 pivot control fields together.");
  const { definition: n } = Sx(e, r),
    o = V1(e, n.source);
  return new Map(
    [...new Set(t)].map((a) => {
      const s = o.fields.indexOf(a),
        d = new Map();
      if (s < 0)
        return [
          a,
          {
            items: [],
            error: `Broken pivot field “${a}”: source field was not found. Restore or reselect it.`,
          },
        ];
      for (const u of o.rows)
        if ((d.set(Cp(u.values[s]), u.values[s]), d.size > 500))
          return [
            a,
            {
              items: [],
              error:
                "This slicer exceeds 500 distinct items. Choose another field or reduce the source.",
            },
          ];
      return [a, { items: [...d.values()], error: "" }];
    }),
  );
}
function NO(e, r, t, n) {
  (n.readOnly && Xt("This workbook is read-only."),
    n.snapshot !== JSON.stringify(e) &&
      Xt("The workbook changed. Preview pivot filters again."),
    (!dl(t) ||
      Object.keys(t).some(
        (s) =>
          !["slicers", "timelines", "sorts", "values", "chart"].includes(s),
      )) &&
      Xt(
        "Pivot view changes can only set filters, sorting, value display and chart settings.",
      ));
  const { owner: o, definition: a } = Sx(e, r);
  return (
    jn(o),
    yR(e, {
      ...a,
      ...t,
      name: `${a.name.slice(0, 88)} filtered`,
      destination: void 0,
    })
  );
}
function GP(e, r) {
  const { owner: t, definition: n } = Sx(e, r),
    o = wC(n),
    a = V1(e, n.source);
  uR(o, a.fields);
  const s = jP(o, a.fields, mR(a, o.filters ?? []));
  return (
    (!n.refresh ||
      vR(e, o, a) !== n.refresh.sourceFingerprint ||
      bR(t, n.outputRange) !== n.refresh.outputFingerprint) &&
      Xt(
        "This PivotTable is stale or unverified. Refresh it before drilling or charting; historical totals cannot be matched to current source rows.",
      ),
    {
      owner: t,
      definition: n,
      config: o,
      data: a,
      rows: s,
      built: OP(o, a.fields, s),
    }
  );
}
function vie(e, r) {
  var s, d;
  const { definition: t, built: n } = GP(e, r),
    o = ((s = t.chart) == null ? void 0 : s.valueIndex) ?? 0;
  n.chartRows.length > 200 &&
    Xt(
      "PivotCharts are limited to 200 leaf groups. Filter or regroup before charting.",
    );
  const a = n.chartRows.map((u) => ({ label: u.label, value: u.values[o] }));
  return (
    a.some((u) => u.value !== null && typeof u.value != "number") &&
      Xt(
        "The PivotChart contains a calculation error. Resolve zero denominators or invalid source calculations.",
      ),
    {
      kind: ((d = t.chart) == null ? void 0 : d.kind) ?? "column",
      measure: n.measures[o].name,
      showAs: t.values[o].showAs ?? "value",
      points: a,
      sourceRevision: t.sourceRevision,
    }
  );
}
function O5(e, r, t) {
  const { definition: n, config: o, data: a, rows: s, built: d } = GP(e, r),
    u = _t(t),
    l = ct(n.outputRange),
    c = u.row - l.firstRow - 1,
    h = u.col - l.firstCol - Math.max(1, o.rowFields.length);
  (c < 0 ||
    c >= d.rowAxes.length ||
    h < 0 ||
    u.col > l.lastCol ||
    u.row > l.lastRow) &&
    Xt(
      "Choose a PivotTable value cell, subtotal or grand total; headers and group labels cannot be drilled.",
    );
  const f = d.rowAxes[c],
    p = d.columnAxes[Math.floor(h / o.values.length)],
    m = h % o.values.length,
    v = (E, A, k) =>
      !A ||
      A.parts.every((F, I) => {
        var z;
        const O = E.values[Yv(a.fields, k[I])],
          j =
            (z = o.dateGroups) == null
              ? void 0
              : z.find((V) => V.field === k[I]);
        return (
          JSON.stringify(j ? DP(O, j, E.index) : Mg(O) ? null : O) ===
          JSON.stringify(F)
        );
      }),
    g = s.filter((E) => v(E, f, o.rowFields) && v(E, p, o.columnFields)),
    b = Nr(e, a.source.sheetId),
    y = ct(a.source.range),
    w = g.map((E) => ({
      locator: {
        workbookId: e.id,
        sheetId: b.id,
        sheetName: b.name,
        row: E.index + 1,
        range: `${ut(y.firstCol)}${E.index + 1}:${ut(y.lastCol)}${E.index + 1}`,
        original: U1(e, b.id, E.index),
      },
      values: [...E.values],
      cells: a.fields.map((A, k) => ({
        address: Cn(E.index, y.firstCol + k),
        cell: structuredClone(
          b.cells[Cn(E.index, y.firstCol + k)] ?? { raw: null },
        ),
      })),
    }));
  return {
    context: qP(e, n, {
      state: "current",
      reason: "Source and output fingerprints match the saved calculation.",
    }),
    address: `${ut(u.col)}${u.row + 1}`,
    valueIndex: m,
    measure: o.values[m],
    underlyingValue: d.baseMatrix[c + 1][u.col - l.firstCol],
    displayedValue: d.matrix[c + 1][u.col - l.firstCol],
    rowGroup: (f == null ? void 0 : f.parts) ?? null,
    columnGroup: (p == null ? void 0 : p.parts) ?? null,
    fields: [...a.fields],
    records: w,
  };
}
function bie(e, r, t, n) {
  (n.readOnly && Xt("This workbook is read-only."),
    n.snapshot !== JSON.stringify(e) &&
      Xt("The workbook changed. Preview drill-through again."));
  const o = O5(e, r, t),
    { owner: a } = Sx(e, r);
  jn(a);
  const s = JSON.stringify({ ...o, records: void 0 });
  s.length > 32700 &&
    Xt(
      "Drill context exceeds the cell-note budget. Reduce filter selections before extracting.",
    );
  const d = [
      [
        "Source locator",
        "Source revision",
        "Original source",
        "Original sheet",
        "Original row",
        ...o.fields,
      ],
      ...o.records.map((l) => {
        var c, h, f;
        return [
          `${l.locator.sheetName}!${l.locator.range}`,
          o.context.sourceRevision,
          ((c = l.locator.original) == null ? void 0 : c.sourceName) ?? null,
          ((h = l.locator.original) == null ? void 0 : h.sourceSheet) ?? null,
          ((f = l.locator.original) == null ? void 0 : f.originalRow) ?? null,
          ...l.values,
        ];
      }),
    ],
    u = gR(e, `${o.context.name.slice(0, 18)} details`, d);
  return (
    (u.sheet.cells.A1.note = `Pivot drill-through context (snapshot, not signed): ${s}`),
    (u.sheet.columnWidths = { 0: 240, 1: 110, 2: 160, 3: 130, 4: 110 }),
    (u.sheet.rowOriginIds = {}),
    (u.workbook.rowOrigins = [...(e.rowOrigins ?? [])]),
    o.records.forEach((l, c) => {
      let h = l.locator.original;
      (h ||
        ((h = {
          id: qw(),
          sourceId: e.id,
          sourceName: e.name,
          sourceSheet: l.locator.sheetName,
          originalRow: l.locator.row,
          kind: "snapshot",
        }),
        u.workbook.rowOrigins.length >= Mt.maxCells &&
          Xt("Row provenance exceeds its 100,000 identity budget."),
        u.workbook.rowOrigins.push(h)),
        (u.sheet.rowOriginIds[c + 1] = h.id),
        l.cells.forEach(({ cell: f }, p) => {
          const m = [
            f.note,
            typeof f.raw == "string" && f.raw.startsWith("=") && !f.literal
              ? `Source formula at extraction: ${f.raw}`
              : void 0,
          ].filter(Boolean).join(`
`);
          (m.length > 32767 &&
            Xt("A source note/formula exceeds the drill-output note budget."),
            m && (u.sheet.cells[`${ut(p + 5)}${c + 2}`].note = m));
        }));
    }),
    u.workbook.sheets.reduce(
      (l, c) => l + Object.keys(c.rowOriginIds ?? {}).length,
      0,
    ) > Mt.maxCells &&
      Xt("Drill row associations exceed the 100,000 source-row budget."),
    u.workbook
  );
}
function ZP(e, r) {
  const t = PP(r),
    n = V1(e, t.source),
    o = mR(n, t.filters ?? []),
    a = t.fields.map((f) => Yv(n.fields, f)),
    s = (t.derived ?? []).map((f) => ({
      ...f,
      index: Yv(n.fields, f.field),
      other: f.otherField ? Yv(n.fields, f.otherField) : void 0,
    }));
  if (t.sort) {
    const f = Yv(n.fields, t.sort.field),
      p = t.sort.direction === "ascending" ? 1 : -1;
    o.sort((m, v) => {
      const g = m.values[f],
        b = v.values[f];
      return (
        (tg(g) && tg(b)
          ? g - b
          : String(g ?? "").localeCompare(String(b ?? ""))) * p ||
        m.index - v.index
      );
    });
  }
  const d = [
      ...t.fields.map((f) =>
        Object.hasOwn(t.rename ?? {}, f) ? t.rename[f] : f,
      ),
      ...s.map((f) => f.name),
    ],
    u = [d];
  (o.length + 1) * d.length > Mt.maxCells &&
    Xt("The query output exceeds 100,000 cells.");
  for (const f of o) {
    const p = a.map((m) => f.values[m]);
    for (const m of s) {
      const v = f.values[m.index],
        g = m.other !== void 0 ? f.values[m.other] : (m.constant ?? null);
      if (m.operation === "upper") p.push(String(v ?? "").toUpperCase());
      else if (m.operation === "lower") p.push(String(v ?? "").toLowerCase());
      else if (m.operation === "trim") p.push(String(v ?? "").trim());
      else if (m.operation === "concat")
        p.push(String(v ?? "") + String(g ?? ""));
      else {
        ((!tg(v) || !tg(g)) &&
          Xt(
            `Derived field “${m.name}” requires numbers on source row ${f.index + 1}.`,
          ),
          m.operation === "divide" &&
            g === 0 &&
            Xt(
              `Derived field “${m.name}” divides by zero on source row ${f.index + 1}.`,
            ));
        const b =
          m.operation === "add"
            ? v + g
            : m.operation === "subtract"
              ? v - g
              : m.operation === "multiply"
                ? v * g
                : v / g;
        (Number.isFinite(b) ||
          Xt("Derived calculation exceeds finite numeric bounds."),
          p.push(b));
      }
    }
    (p.some((m) => typeof m == "string" && m.length > 32767) &&
      Xt("Derived text exceeds the cell text limit."),
      u.push(p));
  }
  const l = gR(e, t.name, u, t.destination),
    { destination: c, ...h } = t;
  return (
    (l.sheet.queries = [
      ...(l.sheet.queries ?? []),
      {
        ...h,
        source: n.source,
        id: qw(),
        sourceRevision: e.revision,
        outputRange: l.range,
      },
    ]),
    l.workbook
  );
}
function yie(e, r, t) {
  const n =
    e.sheets.flatMap((o) => o.queries ?? []).find((o) => o.id === r) ??
    Xt("Query definition not found.");
  return ZP(e, {
    ...n,
    destination: t,
    name: `${n.name.slice(0, 90)} refreshed`,
  });
}
function wie(e, r) {
  for (const t of r.tables ?? []) {
    const n = Ef(t.range, r),
      o = Gp(t),
      a = [];
    for (let d = o.firstRow; d <= o.lastRow; d++) {
      const u = [];
      for (let l = o.firstCol; l <= o.lastCol; l++)
        u.push(e.getCell(Cn(d, l)).value);
      a.push(u);
    }
    const s = t.showTotals
      ? t.columns.map((d, u) => e.getCell(Cn(n.lastRow, n.firstCol + u)).value)
      : [];
    (e.addTable({
      name: t.name,
      ref: Cn(n.firstRow, n.firstCol),
      headerRow: true,
      totalsRow: t.showTotals,
      style: { theme: uie[t.style], showRowStripes: true },
      columns: t.columns.map((d, u) => ({
        name: d.name,
        filterButton: true,
        ...(t.showTotals
          ? {
              totalsRowFunction: Object.hasOwn(t.totals, d.name)
                ? t.totals[d.name]
                : "none",
              ...(u === 0 && !Object.hasOwn(t.totals, d.name)
                ? { totalsRowLabel: String(kx(t, 0).raw ?? "") }
                : {}),
            }
          : {}),
      })),
      rows: a,
    }),
      s.forEach((d, u) => {
        e.getCell(Cn(n.lastRow, n.firstCol + u)).value = d;
      }));
  }
}
const Jv = { changedCells: 500, appendedRows: 100, formulaLength: 4096 };
function wR(e, r = false) {
  (typeof e != "string" ||
    e.length > Jv.formulaLength ||
    e !== e.trim() ||
    !e.startsWith("=")) &&
    Xt(
      "Calculated column formulas must begin with = and contain at most 4,096 characters, without outer whitespace.",
    );
  try {
    pC(e);
  } catch (t) {
    Xt(
      `Calculated column formula: ${t.message.replace(/^XLSX defined names: /, "").replace(/ Import\/export stopped.*$/, "")}`,
    );
  }
  return (
    !r &&
      /#REF!/.test(Vg(e.replace(/"(?:[^"]|"")*"/g, '""'))) &&
      Xt(
        "Repair broken #REF! references before saving a calculated column rule.",
      ),
    e
  );
}
function KP(e, r, t, n) {
  const o = Gp(r);
  for (const [a, s] of r.columns.entries())
    if (s.calculatedFormula)
      for (let d = t; d <= n; d++) {
        const u = Cn(d, o.firstCol + a),
          l = xd(s.calculatedFormula, d - o.firstRow, 0);
        wR(l);
        const c = { ...(e.cells[u] ?? { raw: null }), raw: l };
        (delete c.literal, (e.cells[u] = hi(c)));
      }
}
function YP(e, r, t, n, o) {
  const a = Nr(e, r),
    s = ct(t.range);
  (o !== t.showTotals &&
    Xt("Clear calculated column rules before changing totals-row visibility."),
    (n.firstRow !== s.firstRow ||
      n.firstCol !== s.firstCol ||
      n.lastCol !== s.lastCol ||
      n.lastRow < s.lastRow) &&
      Xt(
        "Calculated tables support downward resize only; keep the existing header and columns. Clear rules before shrinking or changing columns.",
      ),
    (n.lastRow - s.lastRow) * (n.lastCol - n.firstCol + 1) +
      (t.showTotals ? t.columns.length : 0) >
      Jv.changedCells &&
      Xt("Table append/resize may change at most 500 cells."));
  for (let u = s.lastRow + 1; u <= n.lastRow; u++)
    for (let l = n.firstCol; l <= n.lastCol; l++)
      a.cells[Cn(u, l)] &&
        Xt(
          "Table append/resize would overwrite adjacent cells or their formatting. Move that content first.",
        );
  if (t.showTotals && n.lastRow !== s.lastRow) {
    xie(e, r, t);
    for (const u of [...(e.formulaWatches ?? []), ...(e.reviewThreads ?? [])]) {
      if (u.orphaned || u.anchor.sheetId !== r) continue;
      const l = _t(u.anchor.address);
      l.row === n.lastRow &&
        l.col >= n.firstCol &&
        l.col <= n.lastCol &&
        Xt(
          "The new totals row already has a saved watch or review discussion. Preserve or remove that destination anchor before appending.",
        );
    }
  }
}
function xie(e, r, t) {
  var s;
  const n = Nr(e, r),
    o = ct(t.range),
    a = (d, u) => {
      var c, h;
      const l = Vg(d.replace(/"(?:[^"]|"")*"/g, '""'));
      /\b(?:INDIRECT|OFFSET|EVALUATE|LET|LAMBDA)\s*\(/i.test(l) &&
        Xt(
          "A dynamic formula could refer to the moving totals row. Preserve or remove that ambiguity before appending.",
        );
      for (const f of l.matchAll(
        new RegExp(
          "(?<![A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_\\\\][A-Za-z0-9_.\\\\]*))\\s*!\\s*)?(\\$?[1-9]\\d*)\\s*:\\s*(\\$?[1-9]\\d*)(?![A-Za-z0-9_.])",
          "g",
        ),
      )) {
        const p = ((c = f[1]) == null ? void 0 : c.replace(/''/g, "'")) ?? f[2];
        if (
          p ? p.toLowerCase() !== n.name.toLowerCase() : u !== void 0 && u !== r
        )
          continue;
        const m = Number(f[3].replace("$", "")) - 1,
          v = Number(f[4].replace("$", "")) - 1;
        o.lastRow >= Math.min(m, v) &&
          o.lastRow <= Math.max(m, v) &&
          Xt(
            "A formula refers directly to the moving totals row. Use a structured table totals reference or move those references before appending.",
          );
      }
      for (const f of l.matchAll(
        new RegExp(
          "(?<![A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_\\\\][A-Za-z0-9_.\\\\]*))\\s*!\\s*)?(\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*)(?:\\s*:\\s*(\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*))?(?![A-Za-z0-9_.])",
          "g",
        ),
      )) {
        const p = ((h = f[1]) == null ? void 0 : h.replace(/''/g, "'")) ?? f[2];
        if (
          p ? p.toLowerCase() !== n.name.toLowerCase() : u !== void 0 && u !== r
        )
          continue;
        const m = ct(`${f[3]}:${f[4] ?? f[3]}`);
        Qf(m, { ...o, firstRow: o.lastRow }) &&
          Xt(
            "A formula refers directly to the moving totals row. Use a structured table totals reference or move those references before appending.",
          );
      }
    };
  for (const d of e.sheets) {
    for (const [u, l] of Object.entries(d.cells)) {
      const c = _t(u);
      if (
        !(
          d.id === r &&
          c.row === o.lastRow &&
          c.col >= o.firstCol &&
          c.col <= o.lastCol
        ) &&
        (typeof l.raw == "string" &&
          l.raw.startsWith("=") &&
          !l.literal &&
          a(l.raw, d.id),
        ((s = l.hyperlink) == null ? void 0 : s.kind) === "internal" &&
          l.hyperlink.sheetId === r)
      ) {
        const h = _t(l.hyperlink.address);
        h.row === o.lastRow &&
          h.col >= o.firstCol &&
          h.col <= o.lastCol &&
          Xt(
            "An internal hyperlink targets the moving totals row. Retarget or remove it before appending.",
          );
      }
    }
    for (const u of d.conditionalFormats ?? [])
      u.kind === "formula" && u.formula && a(u.formula, d.id);
    for (const u of d.tables ?? [])
      for (const l of u.columns)
        l.calculatedFormula && a(l.calculatedFormula, d.id);
  }
  for (const d of e.definedNames ?? []) a(d.refersTo, d.scopeSheetId);
  for (let d = o.firstCol; d <= o.lastCol; d++) {
    const u = n.cells[Cn(o.lastRow, d)],
      l = kx(t, d - o.firstCol);
    u &&
      (u.raw !== l.raw || u.literal) &&
      Xt("The totals row was edited. Preserve its values before appending.");
  }
}
function JP(e, r, t, n, o) {
  const a = ct(n.range),
    s = ct(o.range),
    d = Nr(e, t),
    u = Nr(r, t);
  for (let c = a.firstCol; c <= a.lastCol; c++) {
    const h = Cn(a.lastRow, c),
      f = Cn(s.lastRow, c),
      p = d.cells[h];
    if (p) {
      const m = { ...structuredClone(p), ...kx(o, c - a.firstCol) };
      (delete m.literal, (u.cells[f] = m));
    }
  }
  const l = (c) => {
    if (c.sheetId !== t) return;
    const h = _t(c.address);
    h.row === a.lastRow &&
      h.col >= a.firstCol &&
      h.col <= a.lastCol &&
      (c.address = Cn(s.lastRow, h.col));
  };
  for (const c of r.formulaWatches ?? []) c.orphaned || l(c.anchor);
  for (const c of r.reviewThreads ?? []) c.orphaned || l(c.anchor);
}
function QP(e, r, t, n, o = {}) {
  var p, m;
  o.readOnly && Xt("This workbook is read-only.");
  const a = Nr(e, r),
    s =
      ((p = a.tables) == null ? void 0 : p.find((v) => v.id === t)) ??
      Xt("Select an existing table.");
  (jn(a), dl(n) || Xt("Choose a table operation."));
  const d = {
    setColumnFormula: ["kind", "columnId", "formula"],
    clearColumnFormula: ["kind", "columnId"],
    appendRows: ["kind", "count", "rows"],
    resize: ["kind", "range"],
    updateTable: ["kind", "options"],
  };
  (!d[n.kind] || Object.keys(n).some((v) => !d[n.kind].includes(v))) &&
    Xt("Unknown calculated-table operation field.");
  let u = structuredClone(e),
    l = Nr(u, r),
    c = l.tables.find((v) => v.id === t);
  if (n.kind === "setColumnFormula" || n.kind === "clearColumnFormula") {
    const v =
      c.columns.find((g) => g.id === n.columnId) ??
      Xt("Select an existing table column.");
    if (n.kind === "setColumnFormula") {
      v.calculatedFormula = wR(n.formula);
      const g = Gp(c);
      (g.lastRow - g.firstRow + 1 > Jv.changedCells &&
        Xt("A calculated column fill supports at most 500 data-body rows."),
        KP(
          l,
          {
            ...c,
            columns: c.columns.map((b) =>
              b.id === v.id ? b : { ...b, calculatedFormula: void 0 },
            ),
          },
          g.firstRow,
          g.lastRow,
        ));
    } else delete v.calculatedFormula;
  } else if (
    n.kind === "appendRows" ||
    n.kind === "resize" ||
    n.kind === "updateTable"
  ) {
    const v = ct(s.range),
      g = Gp(s);
    let b,
      y = 0,
      w;
    n.kind === "appendRows"
      ? (n.rows !== void 0
          ? (n.count !== void 0 &&
              Xt("Specify raw rows or a row count, not both."),
            (!Array.isArray(n.rows) ||
              n.rows.length > Jv.appendedRows ||
              n.rows.some(
                (k) => !Array.isArray(k) || k.length !== s.columns.length,
              )) &&
              Xt(
                "Each appended row must contain one raw value for every table column.",
              ),
            (b = n.rows),
            (y = b.length))
          : (y = n.count ?? 0),
        (!Number.isInteger(y) || y < 1 || y > Jv.appendedRows) &&
          Xt("Append 1–100 rows at a time."),
        v.lastRow + y >= a.rowCount &&
          Xt("The appended table exceeds the worksheet boundary."),
        (w = {
          ...s,
          range: `${Cn(v.firstRow, v.firstCol)}:${Cn(v.lastRow + y, v.lastCol)}`,
        }))
      : (w = n.kind === "resize" ? { ...s, range: n.range } : n.options);
    const E = Ef(w.range, a),
      A = w.showTotals ?? s.showTotals;
    if (
      ((n.kind !== "updateTable" ||
        s.columns.some((k) => k.calculatedFormula)) &&
        YP(e, r, s, E, A),
      (u = hie(e, r, t, w)),
      (l = Nr(u, r)),
      (c = l.tables.find((k) => k.id === t)),
      n.kind === "appendRows" && !s.columns.some((k) => k.calculatedFormula))
    ) {
      for (let k = v.firstRow; k <= g.lastRow; k++)
        for (let F = g.firstCol; F <= g.lastCol; F++) {
          const I = Cn(k, F);
          a.cells[I] && (l.cells[I] = structuredClone(a.cells[I]));
        }
      s.showTotals && JP(e, u, r, s, c);
    }
    if (b)
      for (let k = 0; k < b.length; k++)
        for (let F = 0; F < s.columns.length; F++) {
          const I = b[k][F];
          if (s.columns[F].calculatedFormula) {
            I !== null &&
              I !== "" &&
              Xt(
                "Leave calculated-column input slots blank; their saved formulas supply appended values.",
              );
            continue;
          }
          Vw(I, s.columns[F].dataType) ||
            Xt(
              `Appended value in ${s.columns[F].name} does not match its ${s.columns[F].dataType} type. Supply a typed value or explicitly revise the column type; no rows were appended.`,
            );
          const O = Cn(g.lastRow + 1 + k, g.firstCol + F),
            j = hi({
              ...l.cells[O],
              raw: I,
              ...(typeof I == "string" && I.startsWith("=")
                ? { literal: true }
                : { literal: void 0 }),
            });
          l.cells[O] = j;
        }
  } else Xt("Unknown calculated-table operation.");
  (Ax(l, l), xC(u));
  const h = [];
  for (const v of u.sheets) {
    const g = Nr(e, v.id);
    (m = g.protection) != null &&
      m.enabled &&
      JSON.stringify(g) !== JSON.stringify(v) &&
      Xt(
        "An affected worksheet is protected. Unprotect it before changing dependent table formulas.",
      );
    for (const b of new Set([
      ...Object.keys(g.cells),
      ...Object.keys(v.cells),
    ])) {
      const y = g.cells[b],
        w = v.cells[b];
      (JSON.stringify(y) !== JSON.stringify(w) &&
        h.push({
          sheetId: v.id,
          address: b,
          ...(y ? { before: structuredClone(y) } : {}),
          ...(w ? { after: structuredClone(w) } : {}),
        }),
        h.length > Jv.changedCells &&
          Xt("A reviewed table operation may change at most 500 cells."));
    }
  }
  v0(u, h);
  const f = [];
  for (const [v, g, b] of [
    ["watch", e.formulaWatches, u.formulaWatches],
    ["review", e.reviewThreads, u.reviewThreads],
  ])
    for (const y of b ?? []) {
      const w = g == null ? void 0 : g.find((E) => E.id === y.id);
      w &&
        w.anchor.address !== y.anchor.address &&
        f.push({
          kind: v,
          id: y.id,
          before: w.anchor.address,
          after: y.anchor.address,
        });
    }
  return {
    snapshot: JSON.stringify(e),
    sheetId: r,
    tableId: t,
    request: structuredClone(n),
    beforeRange: s.range,
    afterRange: c.range,
    beforeColumns: structuredClone(s.columns),
    afterColumns: structuredClone(c.columns),
    changes: h,
    anchorChanges: f,
    next: u,
  };
}
function Eie(e, r, t = {}) {
  r.snapshot !== JSON.stringify(e) &&
    Xt("The workbook changed after the table preview. Preview again.");
  const n = QP(e, r.sheetId, r.tableId, r.request, t);
  return (
    JSON.stringify(n) !== JSON.stringify(r) &&
      Xt("The table preview was changed. Preview again."),
    JSON.stringify(n.next) === JSON.stringify(e) ? e : n.next
  );
}
const RO = "\\$?[A-Za-z]{1,3}\\$?[1-9]\\d{0,6}",
  Cie = new RegExp(`^(${RO}(?::${RO})?)(?![A-Za-z0-9_.$:#\\[])`);
function Tx(e, r = 5e3) {
  const t = (p) => {
      const m = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(p);
      if (!m)
        throw Error("External links require bounded A1 cells or rectangles.");
      let v = 0;
      for (const b of m[1].toUpperCase()) v = v * 26 + b.charCodeAt(0) - 64;
      const g = Number(m[2]);
      if (v > 16384 || g > 1048576)
        throw Error("External reference exceeds worksheet bounds.");
      return { row: g, col: v };
    },
    [n, o = n, a] = e.split(":");
  if (a) throw Error("Invalid external range.");
  const s = t(n),
    d = t(o),
    u = [],
    l = Math.min(s.row, d.row),
    c = Math.max(s.row, d.row),
    h = Math.min(s.col, d.col),
    f = Math.max(s.col, d.col);
  if ((c - l + 1) * (f - h + 1) > r)
    throw Error(
      `External link ranges are limited to ${r.toLocaleString()} cells.`,
    );
  for (let p = l; p <= c; p++)
    for (let m = h; m <= f; m++) {
      let v = m,
        g = "";
      for (; v;)
        ((g = String.fromCharCode(65 + ((v - 1) % 26)) + g),
          (v = Math.floor((v - 1) / 26)));
      u.push(g + p);
    }
  return u;
}
function eU(e) {
  const r = [];
  for (let t = 0; t < e.length;) {
    if (e[t] === '"') {
      for (t++; t < e.length;)
        if (e[t++] === '"')
          if (e[t] === '"') t++;
          else break;
      continue;
    }
    const n = t,
      o = /^'((?:[^']|'')*)'\s*!/.exec(e.slice(t)),
      a = /^\[([^\[\]\r\n]+)\]([A-Za-z_][A-Za-z0-9_.]*)!/.exec(e.slice(t));
    let s,
      d = 0;
    if (
      (o
        ? ((s = o[1].replace(/''/g, "'")), (d = o[0].length))
        : a && ((s = `[${a[1]}]${a[2]}`), (d = a[0].length)),
      s)
    ) {
      const u = /^(.*\[[^\[\]\r\n]+\])([^\[\]:\r\n]+)$/.exec(s),
        l = Cie.exec(e.slice(t + d));
      if (u && l && (!t || !/[A-Za-z0-9_.\]]/.test(e[t - 1])))
        try {
          (Tx(l[1]),
            (t += d + l[0].length),
            r.push({
              start: n,
              end: t,
              token: e.slice(n, t),
              target: u[1],
              sheet: u[2],
              range: l[1].toUpperCase(),
            }));
          continue;
        } catch {}
      t += d;
      continue;
    }
    if (e[t] === "[") {
      let u = 0;
      do (e[t] === "[" && u++, e[t] === "]" && u--, t++);
      while (t < e.length && u > 0);
    } else if (e[t] === "'") {
      for (t++; t < e.length;)
        if (e[t++] === "'")
          if (e[t] === "'") t++;
          else break;
    } else t++;
  }
  return r;
}
function xR(e, r) {
  const t = eU(e);
  for (const n of t.reverse()) {
    const o = r(n);
    o !== void 0 && (e = e.slice(0, n.start) + o + e.slice(n.end));
  }
  return e;
}
const Aie = (e, r, t) => `'${(e + r).replace(/'/g, "''")}'!${t}`,
  uh = Object.freeze({
    sources: 8,
    sheets: 32,
    cells: 5e3,
    formulas: 500,
    metadataCharacters: 2e6,
  }),
  rg = (e) =>
    string()
      .min(1)
      .max(e)
      .refine((r) => !/[\u0000-\u001f\u007f]/.test(r)),
  Xw = rg(2048).refine(
    (e) => /^[^\[\]]*\[[^\[\]]+\]$/.test(e),
    "Use an exact workbook qualifier such as [Source.xlsx].",
  ),
  tU = string().regex(/^[a-f0-9]{64}$/),
  kie = discriminatedUnion("type", [
    strictObject({ type: literal("number"), value: number().finite() }),
    strictObject({ type: literal("text"), value: string().max(32767) }),
    strictObject({ type: literal("boolean"), value: boolean() }),
    strictObject({ type: literal("blank"), value: _null() }),
    strictObject({
      type: literal("error"),
      value: string()
        .regex(/^#[A-Z0-9/?!_.]+$/)
        .max(30),
    }),
  ]),
  Sie = strictObject({
    target: Xw,
    sourceId: rg(100),
    sourceName: rg(200),
    sourceRevision: number().int().nonnegative(),
    sourceFingerprint: tU,
    resolvedAt: L3(),
    sheets: array(
      strictObject({
        id: rg(100),
        name: rg(31),
        cells: record(string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/), kie),
      }),
    )
      .min(1)
      .max(32),
  }),
  Tie = strictObject({
    schemaVersion: literal(1),
    sources: array(Sie).max(8),
    packageTargets: array(
      strictObject({
        key: Xw,
        target: rg(4096),
        part: rg(1024),
        sourceSha256: tU,
      }),
    )
      .max(100)
      .optional(),
  }),
  IO = new WeakSet();
function Nie(e, r) {
  var a, s;
  if (e.workbookLinks === void 0) return;
  if (
    e.workbookLinks &&
    typeof e.workbookLinks == "object" &&
    IO.has(e.workbookLinks)
  ) {
    r.workbookLinks = e.workbookLinks;
    return;
  }
  if (JSON.stringify(e.workbookLinks).length > uh.metadataCharacters)
    throw Error("Workbook link metadata exceeds 2 million characters.");
  const t = Tie.parse(e.workbookLinks);
  if (new Set(t.sources.map((d) => d.target)).size !== t.sources.length)
    throw Error("Workbook link targets must be unique.");
  if (
    new Set((a = t.packageTargets) == null ? void 0 : a.map((d) => d.key))
      .size !== (((s = t.packageTargets) == null ? void 0 : s.length) ?? 0)
  )
    throw Error("Package link identities must be unique.");
  let n = 0,
    o = 0;
  for (const d of t.sources) {
    if (
      new Set(d.sheets.map((u) => u.name.toLowerCase())).size !==
      d.sheets.length
    )
      throw Error("Linked worksheet names must be unique.");
    for (const u of d.sheets) {
      o++;
      for (const l of Object.keys(u.cells)) (Tx(l), n++);
    }
  }
  if (n > uh.cells || o > uh.sheets)
    throw Error("Workbook link cache exceeds its cell or sheet limit.");
  ((r.workbookLinks = Lg(t)), IO.add(r.workbookLinks));
}
function ER(e) {
  const r = new TextEncoder().encode(
    JSON.stringify([
      e.id,
      e.name,
      e.revision,
      e.calculation,
      e.datePolicy,
      e.definedNames,
      e.workbookLinks,
      e.sheets.map((t) => [
        t.id,
        t.name,
        t.rowCount,
        t.colCount,
        t.hidden,
        t.veryHidden,
        t.hiddenRows,
        t.hiddenCols,
        t.tables,
        t.merges,
        t.outlineGroups,
        Object.entries(t.cells).map(([n, o]) => [n, o.raw, !!o.literal]),
      ]),
    ]),
  );
  return iu(r);
}
function rU(e) {
  var s;
  const r = [],
    t = new Map(),
    n = new Map(),
    o = new Set(e.sheets.map((d) => d.name.toLowerCase()));
  for (const [d, u] of (
    ((s = e.workbookLinks) == null ? void 0 : s.sources) ?? []
  ).entries())
    for (const [l, c] of u.sheets.entries()) {
      let h = `_LG_LINK_${d}_${l}`;
      for (; o.has(h.toLowerCase());) h += "_";
      o.add(h.toLowerCase());
      let f = `linked:${d}:${l}`;
      for (; e.sheets.some((m) => m.id === f);) f += "_";
      const p = {};
      for (const [m, v] of Object.entries(c.cells))
        ((p[m] = { raw: v.value, literal: true }), t.set(`${f}:${m}`, v));
      (r.push({ id: f, name: h, rowCount: 1048576, colCount: 16384, cells: p }),
        n.set(JSON.stringify([u.target, c.name.toLowerCase()]), h));
    }
  return {
    sheets: r,
    values: t,
    rewrite: (d) =>
      xR(d, (u) => {
        var h;
        const l =
            (h = e.workbookLinks) == null
              ? void 0
              : h.sources.find((f) => f.target === u.target),
          c =
            l == null
              ? void 0
              : l.sheets.find(
                  (f) => f.name.toLowerCase() === u.sheet.toLowerCase(),
                );
        if (!(!c || !Tx(u.range).every((f) => Object.hasOwn(c.cells, f))))
          return `'${n.get(JSON.stringify([u.target, u.sheet.toLowerCase()]))}'!${u.range}`;
      }),
  };
}
function CR(e, r, t, n, o = true) {
  let a = "";
  for (let c = 0; c < e.length;) {
    const h = c,
      f = e[c];
    if (f === '"' || f === "'") {
      for (c++; c < e.length;)
        if (e[c++] === f)
          if (e[c] === f) c++;
          else break;
      a += f === '"' ? '""' : e.slice(h, c);
      continue;
    }
    const p = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(e.slice(c));
    if (p) {
      c += p[0].length;
      let m = 1;
      for (; c < e.length && m;)
        (e[c] === "[" && m++, e[c] === "]" && m--, c++);
      const v = e.slice(h, c);
      a += !m && z1(r, t, v, n) ? "TABLETOKEN" : v;
      continue;
    }
    a += e[c++];
  }
  const s = (c, h, f) => ({
    sheetId: t,
    address: n,
    kind: c,
    error: h,
    message: f,
  });
  if (
    /\[[^\]]+\][^!+*/^&=<>(),;{}]*!/.test(a) ||
    /\|/.test(a) ||
    /(?:https?:|file:|\\\\)/i.test(a)
  )
    return s(
      "external",
      "#BLOCKED!",
      "External workbook, path, network and DDE references are not supported. Use an authorized local worksheet reference.",
    );
  const d = `(?:'(?:[^']|'')*'|[^\\s!+*/^&=<>(),;:{}\\[\\]\\"'%\\-]+)`;
  if (
    /'(?:[^']|'')*:(?:[^']|'')*'\s*!/.test(a) ||
    new RegExp(`${d}\\s*:\\s*${d}\\s*!`).test(a)
  )
    return s(
      "three-dimensional",
      "#VALUE!",
      "3-D references across a worksheet interval are not supported. List the individual worksheet references explicitly.",
    );
  const l = a
    .replace(/'(?:[^']|'')*'\s*!/g, "Sheet!")
    .replace(
      new RegExp(
        "(?<![\\p{L}\\p{N}_.\\]])\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*(?![\\p{L}\\p{N}_.])",
        "gu",
      ),
      "CELLREF",
    );
  if (
    o &&
    (new RegExp(
      "(?<![\\p{L}\\p{N}_.])R(?:\\d+|\\[[+-]?\\d+\\])?C(?:\\d+|\\[[+-]?\\d+\\])?(?![\\p{L}\\p{N}_.]|\\s*[!(])",
      "iu",
    ).test(l) ||
      new RegExp(
        "(?<![\\p{L}\\p{N}_.])[RC]\\[[+-]?\\d+\\](?![\\p{L}\\p{N}_.])",
        "iu",
      ).test(l))
  )
    return s(
      "r1c1",
      "#REF!",
      "R1C1-style references are not supported. Use A1 references, including $A$1, $A1 or A$1 as needed.",
    );
}
const AR = [
    "DSUM",
    "DAVERAGE",
    "DCOUNT",
    "DCOUNTA",
    "DMAX",
    "DMIN",
    "DPRODUCT",
    "DSTDEV",
    "DSTDEVP",
    "DVAR",
    "DVARP",
    "DGET",
  ],
  nU = Symbol("not a database function"),
  np = {
    databaseCells: 2e4,
    criteriaCells: 1e3,
    predicateComparisons: 1e5,
    textComparisonSteps: 5e6,
    criterionCharacters: 255,
    headerCharacters: 255,
    formulaEvaluations: 1e4,
  };
function Rie(e, r, t) {
  if (!AR.includes(e)) return nU;
  const n = (T = "#VALUE!") => {
    throw t.issue(T);
  };
  if (r.length !== 3 || r[0] == null || r[2] == null) return n();
  const o = (T, N) => {
      const R = t.retrieve(T);
      if (t.isIssue(R)) throw R;
      if (!Array.isArray(R) || R.length < 2) return n();
      if (R.length > N) return n("#LIMIT!");
      if (!Array.isArray(R[0])) return n();
      const L = R,
        $ = L[0].length;
      if (!$) return n();
      if (L.length * $ > N) return n("#LIMIT!");
      for (let U = 0; U < L.length; U++)
        if (!Array.isArray(L[U]) || L[U].length !== $) return n();
      for (const U of L)
        for (const G of U)
          if (t.isIssue(G)) {
            if (
              [
                "#LIMIT!",
                "#CYCLE!",
                "#BLOCKED!",
                "#TIMEOUT!",
                "#ITER!",
                "#ERROR!",
              ].includes(G.error)
            )
              throw G;
          } else {
            if (
              G != null &&
              !["number", "string", "boolean"].includes(typeof G)
            )
              return n();
            if (typeof G == "number" && !Number.isFinite(G)) return n("#NUM!");
          }
      return L;
    },
    a = o(r[0], np.databaseCells),
    s = o(r[2], np.criteriaCells);
  if ((a.length - 1) * (s.length - 1) * s[0].length > np.predicateComparisons)
    return n("#LIMIT!");
  const d = (T) => {
      if (t.isIssue(T)) throw T;
      return typeof T != "string" || !T.trim() || T.length > np.headerCharacters
        ? n()
        : T.toLowerCase();
    },
    u = a[0].map(d);
  if (new Set(u).size !== u.length) return n();
  const l = s[0].map((T) => {
      if (t.isIssue(T)) throw T;
      return T == null || T === "" ? -1 : u.indexOf(d(T));
    }),
    c = e === "DCOUNT" || e === "DCOUNTA";
  let h;
  if (r[1] == null) {
    if (!c) return n();
  } else {
    const T = t.retrieve(r[1]);
    if (t.isIssue(T)) throw T;
    if (typeof T == "number") {
      if (!Number.isInteger(T) || T < 1 || T > u.length) return n();
      h = T - 1;
    } else if (typeof T == "string" && T.length) {
      if (((h = u.indexOf(T.toLowerCase())), h < 0)) return n();
    } else return n();
  }
  let f = 0;
  const p = (T) => {
      ((f += T), f > np.textComparisonSteps && n("#LIMIT!"));
    },
    m = (T, N, R) =>
      R === "<>"
        ? T !== N
        : R === "<"
          ? T < N
          : R === ">"
            ? T > N
            : R === "<="
              ? T <= N
              : R === ">="
                ? T >= N
                : T === N,
    v = (T) => {
      var H;
      if (t.isIssue(T)) throw T;
      if (T == null || T === "") return null;
      if (typeof T == "number" || typeof T == "boolean") return (Q) => Q === T;
      if (typeof T != "string") return n();
      if (T.length > np.criterionCharacters) return n("#LIMIT!");
      const N = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(T),
        R = N[1] ?? "=",
        L = N[2];
      if (N[1] && /^[<>=]/.test(L)) return n();
      if (!L && N[1])
        return R !== "=" && R !== "<>"
          ? n()
          : (Q) => (R === "=" ? Q == null || Q === "" : Q != null && Q !== "");
      if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(L.trim())) {
        const Q = Number(L);
        return Number.isFinite(Q)
          ? (X) => (typeof X == "number" ? m(X, Q, R) : R === "<>")
          : n("#NUM!");
      }
      if (/^(TRUE|FALSE)$/i.test(L)) {
        if (R !== "=" && R !== "<>") return n();
        const Q = L.toUpperCase() === "TRUE";
        return (X) => (R === "=" ? X === Q : X !== Q);
      }
      const $ = L.toLowerCase(),
        U = [];
      let G = false;
      for (let Q = 0; Q < $.length; Q++) {
        const X = $[Q];
        if (X === "~") {
          const J = $[Q + 1];
          J && ["*", "?", "~"].includes(J)
            ? (U.push({ kind: "literal", char: J }), Q++, (G = true))
            : U.push({ kind: "literal", char: X });
        } else
          X === "*" || X === "?"
            ? ((G = true),
              (X === "?" ||
                ((H = U.at(-1)) == null ? void 0 : H.kind) !== "many") &&
                U.push({ kind: X === "*" ? "many" : "one" }))
            : U.push({ kind: "literal", char: X });
      }
      return ["<", ">", "<=", ">="].includes(R)
        ? G
          ? n()
          : (Q) =>
              typeof Q != "string"
                ? false
                : (p(Q.length + $.length), m(Q.toLowerCase(), $, R))
        : (!N[1] && !G && U.push({ kind: "many" }),
          (Q) => {
            if (typeof Q != "string") return R === "<>";
            p((Q.length + 1) * (U.length + 1));
            const X = Q.toLowerCase();
            let J = new Uint8Array(X.length + 1);
            J[0] = 1;
            for (const W of U) {
              const K = new Uint8Array(X.length + 1);
              if (W.kind === "many") {
                K[0] = J[0];
                for (let te = 1; te <= X.length; te++)
                  K[te] = J[te] || K[te - 1];
              } else
                for (let te = 1; te <= X.length; te++)
                  K[te] =
                    J[te - 1] && (W.kind === "one" || W.char === X[te - 1])
                      ? 1
                      : 0;
              J = K;
            }
            const q = !!J[X.length];
            return R === "<>" ? !q : q;
          });
    };
  let g = 0;
  const b = s.slice(1).map((T, N) =>
    T.map((R, L) => {
      var $;
      return l[L] < 0
        ? ($ = t.criterionFormula) == null
          ? void 0
          : $.call(t, r[2], N + 1, L)
        : void 0;
    }),
  );
  l.forEach((T, N) => {
    T < 0 && !b.some((R) => R[N]) && n();
  });
  const y = s.slice(1).map((T, N) =>
    T.map((R, L) => {
      if (l[L] >= 0) return v(R);
      if (!b[N][L] && !(R == null || R === "")) return n();
    }),
  );
  b.flat().filter(Boolean).length * (a.length - 1) > np.formulaEvaluations &&
    n("#LIMIT!");
  let E = 0,
    A = 0,
    k = 0,
    F = 0,
    I = 1,
    O = 1 / 0,
    j = -1 / 0,
    z = 0,
    V = 0,
    D;
  for (const [T, N] of a.slice(1).entries()) {
    let R = false;
    for (const [U, G] of y.entries()) {
      let H = true;
      for (let Q = 0; Q < G.length; Q++) {
        const X = b[U][Q];
        if (X) {
          ++g > np.formulaEvaluations && n("#LIMIT!");
          const W = X(T);
          if (t.isIssue(W)) throw W;
          (typeof W != "boolean" && n(), W || (H = false));
          continue;
        }
        const J = G[Q];
        if (!J) continue;
        const q = N[l[Q]];
        if (t.isIssue(q)) throw q;
        J(q) || (H = false);
      }
      H && (R = true);
    }
    if (!R || (E++, h === void 0)) continue;
    const L = N[h];
    if (((D = L), L != null && k++, t.isIssue(L))) {
      if (!["DCOUNT", "DCOUNTA", "DGET"].includes(e)) throw L;
      continue;
    }
    if (typeof L != "number") continue;
    (A++, (F += L), (I *= L), (O = Math.min(O, L)), (j = Math.max(j, L)));
    const $ = L - z;
    ((z += $ / A), (V += $ * (L - z)));
  }
  if (e === "DGET") {
    if (E !== 1) return n(E ? "#NUM!" : "#VALUE!");
    if (t.isIssue(D)) throw D;
    return D ?? 0;
  }
  if (e === "DCOUNT") return h === void 0 ? E : A;
  if (e === "DCOUNTA") return h === void 0 ? E : k;
  const B = (T) => (Number.isFinite(T) ? T : n("#NUM!"));
  if (e === "DSUM") return B(F);
  if (e === "DMAX") return A ? j : 0;
  if (e === "DMIN") return A ? O : 0;
  if (e === "DPRODUCT") return A ? B(I) : 0;
  if (e === "DAVERAGE") return A ? B(F / A) : n("#DIV/0!");
  const _ = e === "DSTDEV" || e === "DVAR";
  if (A < (_ ? 2 : 1)) return n("#DIV/0!");
  const C = Math.max(0, V / (A - (_ ? 1 : 0)));
  return B(e === "DSTDEV" || e === "DSTDEVP" ? Math.sqrt(C) : C);
}
const iU = ["SUBTOTAL", "AGGREGATE", "ISFORMULA", "FORMULATEXT"],
  oU = Symbol("not-reference-function"),
  ip = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  z4 = (e) =>
    typeof (e == null ? void 0 : e.raw) == "string" &&
    e.raw.startsWith("=") &&
    !e.literal;
function Iie(e) {
  return /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?SUBTOTAL\s*\(/i.test(
    e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, ""),
  );
}
function _ie(e) {
  return /(?:^|[^A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?(?:SUBTOTAL|AGGREGATE)\s*\(/i.test(
    e.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, ""),
  );
}
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
const Oie = [
    "SUM",
    "AVERAGE",
    "MIN",
    "MAX",
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "SUMIF",
    "SUMIFS",
    "COUNTIF",
    "COUNTIFS",
    "AVERAGEIF",
    "AVERAGEIFS",
    "IF",
    "IFERROR",
    "IFNA",
    "AND",
    "OR",
    "NOT",
    "TRUE",
    "FALSE",
    "ABS",
    "ROUND",
    "ROUNDDOWN",
    "ROUNDUP",
    "INT",
    "MOD",
    "SIGN",
    "SQRT",
    "POWER",
    "SUMPRODUCT",
    "VLOOKUP",
    "HLOOKUP",
    "INDEX",
    "MATCH",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "TRIM",
    "CLEAN",
    "UPPER",
    "LOWER",
    "CONCATENATE",
    "TEXT",
    "VALUE",
    "SUBSTITUTE",
    "REPLACE",
    "FIND",
    "SEARCH",
    "DATE",
    "DAY",
    "MONTH",
    "YEAR",
    "TODAY",
    "EDATE",
    "EOMONTH",
    "DAYS",
    "NETWORKDAYS",
    "WORKDAY",
    "YEARFRAC",
    "ISBLANK",
    "ISNUMBER",
    "ISTEXT",
    "ISERROR",
    "ISNA",
    "ROW",
    "COLUMN",
    "ROWS",
    "COLUMNS",
    "LARGE",
    "SMALL",
    "MEDIAN",
    "STDEV.S",
    "VAR.S",
  ],
  Fc = [
    ...new Set([
      ...Oie,
      ...ZF,
      ...rP,
      ...oB,
      ...iU,
      ...AR,
      ...QN,
      ...nne,
      ...hne,
      ...YL,
      ...XL,
      ...ZL,
    ]),
  ].sort(),
  Mie = new Set(Fc),
  aU = JN.FormulaError,
  kR = (e) => new aU(e),
  jie = kR,
  wn = (e) => e instanceof aU,
  Gl = (e) => (Array.isArray(e) ? e.flatMap(Gl) : [e]),
  M5 = (e) => e.flatMap(Gl).find(wn);
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
    var b, y;
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
      g = Math.max(
        ((b = p[0]) == null ? void 0 : b.length) ?? 0,
        ((y = m[0]) == null ? void 0 : y.length) ?? 0,
      );
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
      const c = (f) => (f == null ? void 0 : e.retrieveRef(f));
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
                T === void 0 ||
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
            (h = _ == null ? void 0 : _.ref) != null &&
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
              [D, c(l[1]), l.length === 3 ? c(l[2]) : !1],
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
          return c(B ? l[1] : l.length === 3 ? l[2] : !1) ?? 0;
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
            : O[0] === !0
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
              var J;
              if (!Array.isArray(H)) return "1:1";
              const Q = H.every(Array.isArray) ? H : [H],
                X = ((J = Q[0]) == null ? void 0 : J.length) ?? 0;
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
            B = !1;
          D.startsWith("(") &&
            D.endsWith(")") &&
            ((B = !0), (D = D.slice(1, -1).trim()));
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
            C = O[2] === void 0 ? 1 : Number(O[2]);
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
            O[3] !== void 0 &&
            (!Number.isInteger(Number(O[3])) || Number(O[3]) < 1)
          )
            return t("#VALUE!");
          if (
            D.length +
              (O[3] === void 0 ? C : Math.min(C, 1)) * (_.length - B.length) >
            32767
          )
            return t("#LIMIT!");
          if (!B.length) return D;
          if (O[3] === void 0) return D.split(B).join(_);
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
            _ = O[2] === void 0 ? 1 : Number(O[2]);
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
function Cf(e, r = {}, t, n) {
  var W, K;
  const o = z3(e),
    a = new Set([...Fc, ...o.map((te) => te.name)]),
    s = (n == null ? void 0 : n.budget) ?? SN(e),
    d = (n == null ? void 0 : n.dateContext) ?? lx(e.datePolicy),
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
  const j =
      (W = e.definedNames) == null
        ? void 0
        : W.some((te) =>
            /\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(te.refersTo),
          ),
    z =
      (K = e.definedNames) != null && K.length
        ? new RegExp(
            `(?:^|[^A-Za-z0-9_.\\\\])(?:${e.definedNames.map((te) => te.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")})(?![A-Za-z0-9_.\\\\])`,
            "i",
          )
        : void 0,
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
          var oe;
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
              (((oe = e.definedNames) == null ? void 0 : oe.length) &&
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
      var ge;
      const he = (se, we = true) => {
        const Ee = CR(se, e, ae, oe, we);
        if (Ee) throw (l.set(`${ae}:${oe}`, Ee), f(Ee.error));
      };
      ((te = D.rewrite(te)), he(te, false));
      let Se = de ? pf(te, e, ae, [], t == null ? void 0 : t.allowName) : te;
      ((Se = D.rewrite(Se)), he(Se));
      let le = "",
        ce = 0;
      for (; ce < Se.length;) {
        const se =
          /^(?:'((?:[^']|'')*)'|([A-Za-z_][A-Za-z0-9_.]*))!\s*(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#/i.exec(
            Se.slice(ce),
          );
        if (se) {
          const ye = C.get(
            (
              ((ge = se[1]) == null ? void 0 : ge.replace(/''/g, "'")) ?? se[2]
            ).toLowerCase(),
          );
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
          var se;
          return p(le, ce, {
            currentSheetName: (se = ge.position) == null ? void 0 : se.sheet,
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
                : void 0,
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
          var Te;
          const se =
            le && typeof le == "object" && "ref" in le ? le.ref : void 0;
          if (!(se != null && se.from) || !se.to) return;
          const we = U(
              se.sheet ?? ((Te = te.position) == null ? void 0 : Te.sheet),
            ),
            Ee = `${ut(se.from.col + ge - 1)}${se.from.row + ce}`;
          if (t && !t.allowRead(we.id, Ee)) throw f("#BLOCKED!");
          n == null || n.read(`${we.id}:${Ee}`);
          const Ae = we.cells[Ee];
          if (
            (Ae != null && Ae.literal) ||
            typeof (Ae == null ? void 0 : Ae.raw) != "string" ||
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
              const ht = pf(ye, e, we.id, [], t == null ? void 0 : t.allowName),
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
                    !0,
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
          const ce = s.work() ?? (n == null ? void 0 : n.check());
          if (ce) return f(ce);
          const ge = le(),
            se = s.value(ge) ?? (n == null ? void 0 : n.check());
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
        allowName: t == null ? void 0 : t.allowName,
        step: t
          ? (de, he, Se) =>
              t.step({
                kind: de,
                sheetId: ae,
                address: oe,
                operation: he,
                result: Se,
              })
          : void 0,
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
              !0,
            );
          } finally {
            ((se._callFunction = we), F--);
          }
        },
      }),
    Q = (te, ae) => {
      const oe = s.work() ?? (n == null ? void 0 : n.check());
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
      const ce = n == null ? void 0 : n.value(Se, y.size > 0 || w.size > 0);
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
        return (ge == null ? void 0 : ge.raw) ?? null;
      }
      if (typeof ge.raw != "string" || !ge.raw.startsWith("=") || ge.literal)
        return ge.raw;
      const se = ge.raw;
      n == null || n.enter(Se);
      let Ee = (() => {
        var Me, ke, je, $e;
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
              : Je.parse(Te, { sheet: he.name, row: oe + 1, col: de + 1 }, !0);
          if (wn(ze)) Re = f(ze.error);
          else if (typeof ze == "number")
            Re = Number.isFinite(ze) ? ze : f("#NUM!");
          else if (ze === null) Re = 0;
          else if (typeof ze == "boolean") Re = ze;
          else if (typeof ze == "string")
            Re = ze.length <= 32767 ? ze : f("#LIMIT!");
          else if (Array.isArray(ze)) {
            const at = ze.every(Array.isArray) ? ze : [ze],
              _e = ((Me = at[0]) == null ? void 0 : Me.length) ?? 0;
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
              let Ze = !1,
                lt = !1;
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
                pt = (ke = he.merges) == null ? void 0 : ke.find(ft),
                Fe =
                  (je = he.tables) == null
                    ? void 0
                    : je.find((nt) => ft(nt.range));
              if (pt || Fe) {
                Ze = !0;
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
                    lt = !0;
                    break;
                  }
                  if (
                    Wt !== ae &&
                    ((Zt && Zt.raw !== null) || v.has(`${te}:${Wt}`))
                  ) {
                    ((Ze = !0),
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
              : void 0;
          Re = wn(ze) ? f(ze.error) : wn(Je) ? f(Je.error) : f("#ERROR!");
        } finally {
          (y.delete(Se), E.delete(Se));
        }
        return Re;
      })();
      const Ae = s.value(Ee) ?? (n == null ? void 0 : n.check());
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
    var oe, de, he;
    if (
      ((oe = n == null ? void 0 : n.running) == null ? void 0 : oe.call(n)) ===
      false
    )
      try {
        const Se = _t(ae),
          le = _.get(te);
        if (!le || Se.row >= le.rowCount || Se.col >= le.colCount)
          return { type: "error", value: "#REF!" };
        const ce = `${ut(Se.col)}${Se.row + 1}`,
          ge = `${te}:${ce}`,
          se = n.value(ge, !1);
        if (se) return { ...se };
        const we =
          m.get(ge) ??
          ((de = v.get(ge)) == null ? void 0 : de.value) ??
          ((he = le.cells[ce]) == null ? void 0 : he.raw) ??
          null;
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
                !0,
              ),
          we = s.value(se) ?? (n == null ? void 0 : n.check());
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
          he && typeof he == "object" && "details" in he ? he.details : void 0;
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
      return oe ? { ...oe } : void 0;
    }),
    (J.getSpillAnchor = (te, ae) => {
      var oe;
      return (
        J(te, ae),
        (oe = v.get(`${te}:${ae}`)) == null ? void 0 : oe.anchor
      );
    }),
    (J.getSpillEntries = (te) => {
      var ae;
      s.begin();
      try {
        const oe = R.get(te) ?? [];
        if (oe.length > 1e3)
          throw (
            f("#LIMIT!"),
            new Error("Spill discovery exceeds 1,000 possible array formulas.")
          );
        T = 0;
        for (const he of ((ae = n == null ? void 0 : n.running) == null
          ? void 0
          : ae.call(n)) === !1
          ? []
          : oe)
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
function Vn(e, r = {}, t) {
  return !t &&
    !Object.keys(r).length &&
    (e.calculation || e.datePolicy || $ie(e) || eB(e) || nB(e))
    ? Gw(e).calculator
    : Cf(e, r, t);
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
  if ((r == null ? void 0 : r.format) === "date") {
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
    if (n !== void 0) return n;
  }
  return (r == null ? void 0 : r.format) === "percent"
    ? new Intl.NumberFormat("en-US", {
        style: "percent",
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }).format(e)
    : (r == null ? void 0 : r.format) === "currency"
      ? new Intl.NumberFormat("en-US", {
          style: "currency",
          currency: "USD",
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        }).format(e)
      : (r == null ? void 0 : r.format) === "number"
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
    (e.memoryBudgetMiB !== void 0 &&
      (!Number.isInteger(e.memoryBudgetMiB) ||
        Number(e.memoryBudgetMiB) < 1 ||
        Number(e.memoryBudgetMiB) > 128)) ||
    (e.workBudget !== void 0 &&
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
        ...(e.memoryBudgetMiB !== void 0
          ? { memoryBudgetMiB: Number(e.memoryBudgetMiB) }
          : {}),
        ...(e.workBudget !== void 0
          ? { workBudget: Number(e.workBudget) }
          : {}),
      };
}
function Lie(e, r) {
  e.calculation !== void 0 && (r.calculation = Nx(e.calculation));
  const t = e.importedCalculationCaches;
  if (t === void 0) return;
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
    (o.fullPrecision !== void 0 && typeof o.fullPrecision != "boolean") ||
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
      if (u.cache !== void 0) {
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
      ...(o.fullPrecision !== void 0 ? { fullPrecision: o.fullPrecision } : {}),
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
function W1(e) {
  if (!es(e)) return OO(e);
  let r = DO.get(e);
  return (r || ((r = OO(e)), DO.set(e, r)), r);
}
const ok = new WeakMap();
function Uie(e) {
  var u;
  const r = es(e.cells)
    ? (u = ok.get(e.cells)) == null
      ? void 0
      : u.get(e.id)
    : void 0;
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
function nm(e, r, t = W1(r)) {
  return !!e && (es(r) || e.signature() === t.signature());
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
function Gw(e) {
  const r = W1(e),
    t = tu.get(e);
  if (nm(t, e, r)) return { calculator: t.calculator, report: t.report };
  const n = dU(e, r, t ? void 0 : ch.get(e), false);
  return (
    A1(n.report),
    tu.set(e, n),
    ch.delete(e),
    { calculator: n.calculator, report: n.report }
  );
}
function sU(e) {
  const r = W1(e),
    t = tu.get(e),
    n = dU(e, r, nm(t, e, r) ? t : void 0, true);
  return (
    A1(n.report),
    tu.set(e, n),
    ch.delete(e),
    { calculator: n.calculator, report: n.report }
  );
}
function Vie(e, r) {
  return { ...e, calculation: Nx(r) };
}
function Wie(e, r) {
  const t = Gw(e);
  if (t.report !== r)
    throw Error(
      "The workbook or calculation run changed. Review current calculation results before exporting evidence.",
    );
  const n = e.importedCalculationCaches,
    o = ((n == null ? void 0 : n.entries) ?? []).map((a) => {
      const s = e.sheets.find((h) => h.id === a.sheetId),
        d = s == null ? void 0 : s.cells[a.address],
        u = s ? t.calculator.getTypedValue(a.sheetId, a.address) : void 0,
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
            : (d == null ? void 0 : d.raw) !== a.formula
              ? "formula-changed"
              : (u == null ? void 0 : u.type) === "error" &&
                  u.value === "#PENDING!"
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
function A1(e) {
  !e ||
    typeof e != "object" ||
    Object.isFrozen(e) ||
    (Object.values(e).forEach(A1), Object.freeze(e));
}
function Hie(e) {
  const r = tu.get(e);
  return nm(r, e) ? { calculator: r.calculator, report: r.report } : void 0;
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
function L5(e) {
  var m;
  const r = tu.get(e);
  if (
    ((m = r == null ? void 0 : r.report.execution) == null
      ? void 0
      : m.state) === "pending" &&
    nm(r, e)
  )
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
    run: ((o == null ? void 0 : o.report.run) ?? 0) + 1,
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
function lU(e) {
  return W1(e).inputSignature();
}
function uU(e, r, t) {
  var l;
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
    if (
      a.has(h) ||
      ((l = o.formulas.get(h)) == null ? void 0 : l.formula) !== c.formula
    )
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
function cU(e, r, t, n) {
  const o = new Map(n.spills.map((c) => [`${c.sheetId}:${c.address}`, c])),
    a = new Map(n.spillMetadata.map((c) => [`${c.sheetId}:${c.address}`, c])),
    s = (c, h) => `${c}:${h.replace(/\$/g, "").toUpperCase()}`,
    d = new Map(e.sheets.map((c) => [c.id, c])),
    u = (c, h) => {
      var g, b;
      const f = s(c, h),
        p = t.get(f) ?? ((g = o.get(f)) == null ? void 0 : g.result);
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
      const v =
        ((b = m.cells[h.replace(/\$/g, "").toUpperCase()]) == null
          ? void 0
          : b.raw) ?? null;
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
        var f;
        return (f = o.get(s(c, h))) == null ? void 0 : f.anchor;
      },
      getSpillRange: (c, h) => {
        var f;
        return (f = a.get(s(c, h))) == null ? void 0 : f.range;
      },
      getSpillDiagnostic: (c, h) => {
        var f;
        return structuredClone(
          (f = a.get(s(c, h))) == null ? void 0 : f.diagnostic,
        );
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
      var W;
      (_O.test(J.formula) ||
        !t.values.has(q) ||
        ["#TIMEOUT!", "#LIMIT!"].includes(
          String((W = t.values.get(q)) == null ? void 0 : W.value),
        )) &&
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
      y ? w : void 0
    ),
    k = (U, G) => {
      const H = [];
      return Cf(e, {}, void 0, {
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
          var J;
          (H.push(X),
            (E -= ((J = c.get(X)) == null ? void 0 : J.size) ?? 0),
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
      (h.set(U, ng),
        v.set(U, ng),
        c.set(U, new Set(t == null ? void 0 : t.dependencies.get(U))));
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
    var J, q;
    const H = V(U, G);
    if (h.has(H)) return { ...h.get(H) };
    const Q =
      (J = e.sheets.find((W) => W.id === U)) == null
        ? void 0
        : J.cells[G.replace(/\$/g, "").toUpperCase()];
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
    return X && ((q = h.get(`${U}:${X}`)) == null ? void 0 : q.type) === "error"
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
    C = (t == null ? void 0 : t.calculator.getReferenceDiagnostics()) ?? [],
    T = (t == null ? void 0 : t.calculator.getNamedFormulaDiagnostics()) ?? [],
    N = (t == null ? void 0 : t.calculator.getFinancialDiagnostics()) ?? [];
  ((z.getFinancialDiagnostics = () =>
    structuredClone(_(N, I.getFinancialDiagnostics()))),
    (z.getReferenceDiagnostics = () => _(C, I.getReferenceDiagnostics())),
    (z.getNamedFormulaDiagnostics = () =>
      _(T, I.getNamedFormulaDiagnostics())));
  const R = new Map(
      t == null
        ? void 0
        : t.report.cells.map((U) => [`${U.sheetId}:${U.address}`, U]),
    ),
    L = new Map();
  for (const { sheetId: U, address: G, ...H } of z.getFinancialDiagnostics()) {
    const Q = `${U}:${G}`;
    L.set(Q, [...(L.get(Q) ?? []), H]);
  }
  const $ = [...a].map(([U, G]) => {
    var X;
    const H = h.get(U) ?? ng,
      Q = t == null ? void 0 : t.values.get(U);
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
      (Q == null ? void 0 : Q.value) === "#PENDING!" &&
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
      run: ((t == null ? void 0 : t.report.run) ?? 0) + 1,
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
function Kie(e, r, t = "A1") {
  var d;
  const n = Nr(e, r),
    o = ct(t);
  if (o.firstRow !== o.lastRow || o.firstCol !== o.lastCol) return t;
  const a =
    (d = n.tables) == null
      ? void 0
      : d.find((u) => {
          const l = ct(u.range);
          return (
            o.firstRow >= l.firstRow &&
            o.firstRow <= l.lastRow &&
            o.firstCol >= l.firstCol &&
            o.firstCol <= l.lastCol
          );
        });
  if (!a) return Ta(n, false);
  const s = ct(a.range);
  return `${ut(s.firstCol)}${s.firstRow + 1}:${ut(s.lastCol)}${s.lastRow + 1 - (a.showTotals ? 1 : 0)}`;
}
function TR(e, r, t) {
  var h;
  if (
    !t ||
    typeof t.range != "string" ||
    !Array.isArray(t.columns) ||
    t.columns.length > Cc.columns ||
    [t.caseSensitive, t.unique].some(
      (f) => f !== void 0 && typeof f != "boolean",
    )
  )
    throw Error("Choose a header range and up to 64 column filters.");
  const n = Nr(e, r),
    o = ct(t.range),
    a = o.lastRow - o.firstRow,
    s = o.lastCol - o.firstCol + 1;
  if (o.lastRow >= n.rowCount || o.lastCol >= n.colCount || a < 1)
    throw Error(
      "Filter range needs one header row and at least one data row within the worksheet.",
    );
  if (a > Cc.rows || (a + 1) * s > Cc.cells || s > Cc.columns)
    throw Error(
      "Filtering is limited to 10,000 data rows, 100,000 cells including headers, and 64 columns. Choose a smaller range.",
    );
  if (
    (h = n.merges) != null &&
    h.some((f) => {
      const p = ct(f);
      return (
        p.firstRow <= o.lastRow &&
        p.lastRow >= o.firstRow &&
        p.firstCol <= o.lastCol &&
        p.lastCol >= o.firstCol
      );
    })
  )
    throw Error(
      "Unmerge cells in the filter range first; records cannot be split by filtering.",
    );
  const d = (f) => Number.isInteger(f) && f >= o.firstCol && f <= o.lastCol,
    u = (f) => {
      if (
        !f ||
        !d(f.column) ||
        !["value", "text", "number", "date"].includes(f.type) ||
        !Zie.includes(f.operator) ||
        (!["blank", "notBlank"].includes(f.operator) && !jO(f.value))
      )
        throw Error("Invalid filter condition, column or value.");
      if (
        !["blank", "notBlank"].includes(f.operator) &&
        f.type === "text" &&
        typeof f.value != "string"
      )
        throw Error("Text criteria require a text value.");
      if (typeof f.value == "string" && f.value.length > Cc.text)
        throw Error("Filter criteria text is limited to 1,024 characters.");
      if (
        [
          "contains",
          "notContains",
          "startsWith",
          "endsWith",
          "wildcard",
        ].includes(f.operator) &&
        f.type !== "text"
      )
        throw Error("Text matching requires a Text condition.");
      if (["gt", "gte", "lt", "lte"].includes(f.operator) && f.type === "value")
        throw Error("Choose Number, Date or Text for ordered comparisons.");
      if (
        !["blank", "notBlank"].includes(f.operator) &&
        ["number", "date"].includes(f.type) &&
        U5(f.value, f.type, e) === null
      )
        throw Error(
          `Invalid ${f.type} filter value. Use a finite number or a date in the workbook's date policy.`,
        );
    };
  let l = 0;
  const c = new Set();
  for (const f of t.columns) {
    if (
      !f ||
      !d(f.column) ||
      c.has(f.column) ||
      (f.join !== void 0 && !["and", "or"].includes(f.join))
    )
      throw Error("Filter columns must be valid and distinct.");
    if (
      (c.add(f.column),
      f.values !== void 0 &&
        (!Array.isArray(f.values) ||
          f.values.length > Cc.values ||
          !f.values.every(jO)))
    )
      throw Error("Value filters allow up to 1,000 finite typed values.");
    if (
      f.conditions !== void 0 &&
      (!Array.isArray(f.conditions) || f.conditions.length > 2)
    )
      throw Error("Each column supports up to two custom conditions.");
    for (const p of f.conditions ?? []) {
      if (p.column !== f.column)
        throw Error("A column condition must use its own column.");
      (u(p), l++);
    }
  }
  if (
    t.advanced !== void 0 &&
    (!Array.isArray(t.advanced) || t.advanced.length > Cc.conditions)
  )
    throw Error("Invalid advanced filter groups.");
  for (const f of t.advanced ?? []) {
    if (!Array.isArray(f) || !f.length || f.length > Cc.conditions)
      throw Error(
        "Each advanced OR group must contain at least one AND condition.",
      );
    for (const p of f) (u(p), l++);
  }
  if (l > Cc.conditions)
    throw Error("At most 128 combined filter conditions are allowed.");
  return structuredClone({
    ...t,
    range: `${ut(o.firstCol)}${o.firstRow + 1}:${ut(o.lastCol)}${o.lastRow + 1}`,
  });
}
function U5(e, r, t) {
  if (r === "number") return typeof e == "number" ? e : null;
  if (r === "date") {
    const n = ts(t.datePolicy),
      o = typeof e == "number" ? e : typeof e == "string" ? zp(e, n) : void 0;
    return o !== void 0 && o >= 0 && o < wd(n.dateSystem) + 1 ? o : null;
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
  var s;
  const o = e(r, t),
    a =
      (s = n == null ? void 0 : n.sheets.find((d) => d.id === r)) == null
        ? void 0
        : s.cells[t];
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
function f0(e, r, t, n = Vn(e)) {
  var F, I;
  const o = TR(e, r, t),
    a = Nr(e, r),
    s = ct(o.range),
    d = !!o.caseSensitive,
    u = ul(a.rowCount, cn(a, "row")),
    l = [],
    c = [],
    h = [],
    f = [],
    p = new Set(),
    m = {},
    v = o.columns.map((O) =>
      O.values === void 0 ? null : new Set(O.values.map((j) => eh(j, d))),
    );
  let g = 0,
    b = 0,
    y = 0;
  const w = (O, j) => {
      if (
        ((y +=
          j.operator === "wildcard"
            ? Math.max(1, String(O ?? "").length) *
              Math.max(1, String(j.value ?? "").length)
            : 1 + String(O ?? "").length + String(j.value ?? "").length),
        y > Cc.comparisonWork)
      )
        throw Error(
          "Filter comparison work exceeds its 5,000,000-unit budget. Narrow the range or simplify criteria.",
        );
      return Jie(O, j, e, d);
    },
    E = Array.from({ length: s.lastCol - s.firstCol + 1 }, (O, j) => ({
      column: s.firstCol + j,
      label:
        String($5(n, r, `${ut(s.firstCol + j)}${s.firstRow + 1}`, e) ?? "") ||
        ut(s.firstCol + j),
    }));
  for (let O = s.firstRow + 1; O <= s.lastRow; O++) {
    const j = E.map((B) => $5(n, r, `${ut(B.column)}${O + 1}`, e)),
      z = (B) => j[B - s.firstCol];
    if (
      ((y += j.reduce(
        (B, _) => B + 1 + (typeof _ == "string" ? _.length : 0),
        0,
      )),
      y > Cc.comparisonWork)
    )
      throw Error(
        "Filter comparison work exceeds its 5,000,000-unit budget. Narrow the range or simplify criteria.",
      );
    (l.push(O + 1),
      (F = a.rowOriginIds) != null && F[O] && (m[O + 1] = a.rowOriginIds[O]));
    const V = u.isHidden(O);
    V && g++;
    let D =
      o.columns.every((B, _) => {
        var C;
        return (
          (v[_] === null || v[_].has(eh(z(B.column), d))) &&
          (!((C = B.conditions) != null && C.length) ||
            (B.join === "or"
              ? B.conditions.some((T) => w(z(B.column), T))
              : B.conditions.every((T) => w(z(B.column), T))))
        );
      }) &&
      (!((I = o.advanced) != null && I.length) ||
        o.advanced.some((B) => B.every((_) => w(z(_.column), _))));
    if (D && o.unique) {
      const B = JSON.stringify(j.map((_) => eh(_, d)));
      p.has(B) ? (D = false) : p.add(B);
    }
    if (D) (c.push(O + 1), V ? b++ : h.push(O + 1));
    else {
      const B = f.at(-1);
      (B == null ? void 0 : B.end) === O - 1
        ? (B.end = O)
        : f.push({ start: O, end: O });
    }
  }
  const A = [];
  for (const O of [...cn(a, "row"), ...f].sort((j, z) => j.start - z.start)) {
    const j = A.at(-1);
    j && O.start <= j.end + 1
      ? (j.end = Math.max(j.end, O.end))
      : A.push({ ...O });
  }
  return {
    hiddenRows: Ad(A, a.rowCount),
    fields: E,
    excludedRows: f,
    context: {
      workbookId: e.id,
      revision: e.revision,
      sheetId: r,
      sheetName: a.name,
      range: o.range,
      definition: o,
      totalRows: l.length,
      matchedRows: c.length,
      visibleRows: h.length,
      filteredRows: l.length - c.length,
      manuallyHiddenRows: g,
      hiddenMatchingRows: b,
      sourceRows: l,
      matchingSourceRows: c,
      visibleSourceRows: h,
      originalSourceIds: m,
      policy: Gie,
    },
  };
}
function FO(e, r, t, n, o = "") {
  const a = TR(e, r, t),
    s = ct(a.range);
  if (!Number.isInteger(n) || n < s.firstCol || n > s.lastCol)
    throw Error("Facet column is outside the filter range.");
  const d = Vn(e),
    u = f0(
      e,
      r,
      { ...a, unique: false, columns: a.columns.filter((c) => c.column !== n) },
      d,
    ),
    l = new Map();
  for (const c of u.context.matchingSourceRows) {
    const h = $5(d, r, `${ut(n)}${c}`, e),
      f = eh(h, a.caseSensitive);
    if (!Zw(String(h ?? "(Blanks)")).includes(Zw(o))) continue;
    const p = l.get(f);
    p ? p.count++ : l.set(f, { value: h, count: 1, key: f });
  }
  return {
    values: [...l.values()].slice(0, Cc.values),
    distinct: l.size,
    truncated: l.size > Cc.values,
  };
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
function fU(e, r) {
  if (e === void 0) return;
  if (
    !Array.isArray(e) ||
    e.length > z5.views ||
    JSON.stringify(e).length > z5.characters
  )
    throw Error(
      "Saved filter views are limited to 20 views and 512,000 metadata characters per workbook.",
    );
  const t = new Set(),
    n = new Set();
  r.savedFilterViews = e.map((o) => {
    if (
      (xv(o, [
        "id",
        "name",
        "sheetId",
        "sheetName",
        "sourceWorkbookId",
        "sourceRevision",
        "savedAt",
        "datePolicy",
        "definition",
        "population",
      ]),
      !Ry(o.id, 100) ||
        !Ry(o.sheetId, 100) ||
        !Ry(o.sheetName, 31) ||
        !Ry(o.sourceWorkbookId, 100) ||
        !Number.isSafeInteger(o.sourceRevision) ||
        o.sourceRevision < 0 ||
        o.sourceRevision > Number.MAX_SAFE_INTEGER - 1 ||
        typeof o.savedAt != "string" ||
        !Number.isFinite(Date.parse(o.savedAt)) ||
        new Date(o.savedAt).toISOString() !== o.savedAt)
    )
      throw Error("Invalid saved filter view identity, revision or timestamp.");
    const a = NR(o.name);
    if (t.has(o.id) || n.has(BO(a)))
      throw Error("Saved filter view names and IDs must be unique.");
    (t.add(o.id),
      n.add(BO(a)),
      xv(o.datePolicy, [
        "calendar",
        "dateSystem",
        "textDateOrder",
        "timeZone",
      ]));
    const s = ts(o.datePolicy);
    xv(o.definition, [
      "range",
      "columns",
      "advanced",
      "caseSensitive",
      "unique",
    ]);
    const d = (h) => xv(h, ["column", "type", "operator", "value"]);
    if (!Array.isArray(o.definition.columns))
      throw Error("Invalid saved filter columns.");
    for (const h of o.definition.columns)
      if (
        (xv(h, ["column", "values", "conditions", "join"]),
        h.conditions !== void 0)
      ) {
        if (!Array.isArray(h.conditions))
          throw Error("Invalid saved filter conditions.");
        h.conditions.forEach(d);
      }
    if (o.definition.advanced !== void 0) {
      if (!Array.isArray(o.definition.advanced))
        throw Error("Invalid saved advanced criteria.");
      for (const h of o.definition.advanced) {
        if (!Array.isArray(h)) throw Error("Invalid saved advanced criteria.");
        h.forEach(d);
      }
    }
    const u = {
        ...r,
        datePolicy: s,
        sheets: [
          {
            id: o.sheetId,
            name: o.sheetName,
            rowCount: Mt.maxRows,
            colCount: Mt.maxCols,
            cells: {},
          },
        ],
      },
      l = TR(u, o.sheetId, o.definition);
    xv(o.population, [
      "totalRows",
      "matchedRows",
      "visibleRows",
      "hiddenMatchingRows",
    ]);
    const c = o.population;
    if (
      Object.keys(c).length !== 4 ||
      Object.values(c).some((h) => !Number.isInteger(h) || Number(h) < 0) ||
      c.totalRows !== ct(l.range).lastRow - ct(l.range).firstRow ||
      c.totalRows > 1e4 ||
      c.totalRows < 1 ||
      c.matchedRows > c.totalRows ||
      c.visibleRows > c.matchedRows ||
      c.hiddenMatchingRows !== c.matchedRows - c.visibleRows
    )
      throw Error("Invalid saved population counts.");
    return {
      id: o.id,
      name: a,
      sheetId: o.sheetId,
      sheetName: o.sheetName,
      sourceWorkbookId: o.sourceWorkbookId,
      sourceRevision: o.sourceRevision,
      savedAt: o.savedAt,
      datePolicy: s,
      definition: l,
      population: { ...c },
    };
  });
}
function RR(e, r) {
  var t, n, o;
  if (
    ((t = e.structureProtection) != null && t.enabled) ||
    ((o =
      (n = e.sheets.find((a) => a.id === r)) == null ? void 0 : n.protection) !=
      null &&
      o.enabled)
  )
    throw Error(
      "Unprotect the workbook and source worksheet before changing saved filter views.",
    );
}
function IR(e, r) {
  const t = { ...e };
  return (fU(r, t), t);
}
function _R(e, r) {
  var n;
  const t =
    (n = e.savedFilterViews) == null ? void 0 : n.find((o) => o.id === r);
  if (!t) throw Error("This saved filter view no longer exists.");
  return t;
}
function Qie(e, r, t, n, o) {
  const a = Nr(e, r);
  if ((RR(e, r), o && _R(e, o).sheetId !== r))
    throw Error("A saved view must keep its source worksheet.");
  const s = f0(e, r, t).context,
    d = {
      id: o ?? crypto.randomUUID(),
      name: NR(n),
      sheetId: r,
      sheetName: a.name,
      sourceWorkbookId: e.id,
      sourceRevision: e.revision,
      savedAt: new Date().toISOString(),
      datePolicy: ts(e.datePolicy),
      definition: s.definition,
      population: {
        totalRows: s.totalRows,
        matchedRows: s.matchedRows,
        visibleRows: s.visibleRows,
        hiddenMatchingRows: s.hiddenMatchingRows,
      },
    };
  return IR(e, [...(e.savedFilterViews ?? []).filter((u) => u.id !== o), d]);
}
function eoe(e, r, t) {
  return (
    RR(e, _R(e, r).sheetId),
    IR(
      e,
      e.savedFilterViews.map((n) => (n.id === r ? { ...n, name: NR(t) } : n)),
    )
  );
}
function toe(e, r) {
  return (
    RR(e, _R(e, r).sheetId),
    IR(
      e,
      e.savedFilterViews.filter((t) => t.id !== r),
    )
  );
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
  var t;
  const r = pU.safeParse(e);
  if (!r.success)
    throw Error(
      `Invalid consolidation: ${(t = r.error.issues[0]) == null ? void 0 : t.message}`,
    );
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
  var n, o, a;
  let t = 0;
  for (const s of r.sheets) {
    const d =
      (o = (n = e.sheets) == null ? void 0 : n.find((f) => f.id === s.id)) ==
      null
        ? void 0
        : o.consolidation;
    if (d === void 0) continue;
    if (((t += JSON.stringify(d).length), t > k1.metadataCharacters))
      throw Error(
        "Consolidation evidence exceeds 8,000,000 characters. Retain native backups before removing obsolete reports.",
      );
    const u = loe.safeParse(d);
    if (!u.success)
      throw Error(
        `Invalid consolidation evidence: ${(a = u.error.issues[0]) == null ? void 0 : a.message}`,
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
  var o;
  let t = 0,
    n = 0;
  for (const a of r.sheets) {
    const s =
      (o = e.sheets.find((l) => l.id === a.id)) == null
        ? void 0
        : o.externalConnection;
    if (s === void 0) continue;
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
const jR = "ledgerglass-analysis-v1",
  ig = Object.freeze({
    variables: 6,
    constraints: 16,
    candidates: 5e4,
    dependencyCells: 500,
    depth: 16,
    expressionCharacters: 4096,
    syntaxNodes: 512,
    sourceRows: 500,
    sourceColumns: 6,
    sourceCells: 2e3,
    outputCells: 2e4,
    metadataCharacters: 2e6,
  }),
  th = number().finite(),
  Qv = th.min(-1e9).max(1e9),
  V4 = strictObject({
    sheetId: string().min(1).max(100),
    address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
  }),
  wU = string()
    .trim()
    .min(1)
    .max(31)
    .refine(
      (e) => !/[\\/?*\[\]:\u0000-\u001f]/.test(e),
      "Invalid output worksheet name",
    ),
  yoe = strictObject({
    kind: literal("solver"),
    outputName: wU,
    objective: V4,
    direction: _enum(["maximize", "minimize", "target"]),
    target: Qv.optional(),
    variables: array(
      strictObject({
        cell: V4,
        lower: Qv,
        upper: Qv,
        domain: _enum(["continuous", "integer", "binary"]),
      }),
    )
      .min(1)
      .max(6),
    constraints: array(
      strictObject({
        left: V4,
        operator: _enum(["<=", ">=", "="]),
        right: union([Qv, V4]),
      }),
    ).max(16),
    tolerance: th.min(1e-10).max(1e-5).default(1e-8),
    maxCandidates: number().int().min(1).max(5e4).default(5e4),
  }),
  woe = strictObject({
    kind: literal("statistics"),
    outputName: wU,
    sheetId: string().min(1).max(100),
    range: string().min(1).max(50),
    method: _enum([
      "descriptive",
      "correlation",
      "covariance",
      "regression",
      "histogram",
      "distribution",
      "forecast",
    ]),
    missing: _enum(["reject", "exclude"]).default("reject"),
    bins: array(Qv).min(1).max(50).optional(),
    distribution: _enum(["normal", "binomial", "poisson"]).optional(),
    mean: Qv.optional(),
    standardDeviation: th.positive().max(1e9).optional(),
    trials: number().int().min(0).max(1e3).optional(),
    probability: th.min(0).max(1).optional(),
    lambda: th.positive().max(1e3).optional(),
    forecast: _enum(["linear", "moving-average", "exponential"]).optional(),
    horizon: number().int().min(1).max(24).optional(),
    window: number().int().min(2).max(100).optional(),
    alpha: th.positive().max(1).optional(),
  }),
  xU = discriminatedUnion("kind", [yoe, woe]),
  xoe = union([string().max(32767), th, boolean(), _null()]),
  Eoe = strictObject({
    status: _enum([
      "complete",
      "optimal",
      "infeasible",
      "work-limit",
      "precision-limit",
    ]),
    success: boolean(),
    message: string().max(2e3),
    algorithm: string().max(300),
    assumptions: array(string().max(1e3)).max(30),
    exclusions: array(
      strictObject({
        row: number().int().positive(),
        reason: string().max(200),
        locator: string().max(300),
      }),
    ).max(500),
    included: number().int().nonnegative(),
    inspected: number().int().nonnegative(),
    iterations: number().int().nonnegative(),
    rows: array(array(xoe).max(16)).max(3e3),
    solution: array(th).max(6).optional(),
    objective: th.optional(),
    baseline: array(th).max(6).optional(),
  }),
  sk = string().regex(/^[a-f0-9]{64}$/),
  Coe = strictObject({
    engine: literal(jR),
    request: xU,
    result: Eoe,
    definitionHash: sk,
    sourceHash: sk,
    outputHash: sk,
    outputRange: string().max(50),
    createdAt: string().datetime(),
    sourceRevision: number().int().nonnegative(),
    includeHidden: boolean(),
    applied: boolean(),
  });
function EU(e) {
  var n, o;
  const r = xU.safeParse(e);
  if (!r.success)
    throw Error(
      `Invalid analysis request: ${(n = r.error.issues[0]) == null ? void 0 : n.path.join(".")}: ${(o = r.error.issues[0]) == null ? void 0 : o.message}`,
    );
  const t = r.data;
  if (t.kind === "solver") {
    if (t.direction === "target" && t.target === void 0)
      throw Error("Enter a finite target value.");
    if (
      new Set(t.variables.map((a) => JSON.stringify(a.cell))).size !==
      t.variables.length
    )
      throw Error("Changing cells must be unique.");
    if (
      t.variables.some(
        (a) =>
          a.lower > a.upper ||
          (a.domain === "binary" && (a.lower < 0 || a.upper > 1)),
      )
    )
      throw Error(
        "Variable bounds must be ordered; binary bounds must lie within zero and one.",
      );
    if (
      t.variables.some((a) => a.domain === "continuous") &&
      t.variables.some((a) => a.domain !== "continuous")
    )
      throw Error(
        "Mixed integer and continuous models are outside this selected Solver subset.",
      );
  }
  return t;
}
function Aoe(e, r) {
  var o, a, s, d;
  const t = e;
  let n = 0;
  for (const u of r.sheets) {
    const l =
      (o = t.sheets.find((m) => m.id === u.id)) == null
        ? void 0
        : o.analysisReport;
    if (l === void 0) continue;
    if (((n += JSON.stringify(l).length), n > ig.metadataCharacters))
      throw Error("Analysis metadata exceeds 2,000,000 characters.");
    const c = Coe.safeParse(l);
    if (!c.success)
      throw Error(
        `Invalid analysis report: ${(a = c.error.issues[0]) == null ? void 0 : a.message}`,
      );
    const h = c.data;
    if (
      ((h.request = EU(h.request)),
      h.definitionHash !== hr(h.request) ||
        !h.result.success ||
        !["complete", "optimal"].includes(h.result.status) ||
        h.result.rows.reduce((m, v) => m + v.length, 0) > ig.outputCells)
    )
      throw Error("Analysis report metadata is inconsistent.");
    const f = /^([A-Z]{1,3})([1-9]\d{0,6}):([A-Z]{1,3})([1-9]\d{0,6})$/.exec(
        h.outputRange,
      ),
      p = (m) => [...m].reduce((v, g) => v * 26 + g.charCodeAt(0) - 64, 0);
    if (
      !f ||
      p(f[3]) > 16384 ||
      Number(f[4]) > 1048576 ||
      p(f[3]) < p(f[1]) ||
      Number(f[4]) < Number(f[2]) ||
      (p(f[3]) - p(f[1]) + 1) * (Number(f[4]) - Number(f[2]) + 1) >
        ig.outputCells
    )
      throw Error("Invalid or oversized analysis output range.");
    if (
      (h.request.kind === "solver" &&
        (h.result.status !== "optimal" ||
          ((s = h.result.solution) == null ? void 0 : s.length) !==
            h.request.variables.length ||
          ((d = h.result.baseline) == null ? void 0 : d.length) !==
            h.request.variables.length ||
          h.result.objective === void 0)) ||
      (h.request.kind === "statistics" &&
        (h.result.status !== "complete" || h.applied))
    )
      throw Error("Analysis result does not match its request.");
    u.analysisReport = h;
  }
}
const CU = "ledgerglass-report-output-v1",
  Vv = Object.freeze({
    outputs: 20,
    cells: 2e3,
    rows: 500,
    columns: 20,
    metadataBytes: 16e6,
    imagePixels: 16e6,
    imageWidth: 4096,
    imageHeight: 16384,
    inventories: 20,
    inventoryItems: 1e3,
  }),
  Tc = (e) => string().max(e),
  Ap = string().min(1).max(200),
  zE = string().regex(/^[a-f0-9]{64}$/),
  AU = string().datetime();
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
const W5 = (e) => {
  const { hash: r, ...t } = e;
  return hr(t);
};
function Wg(e) {
  if (new TextEncoder().encode(JSON.stringify(e)).length > 4e6)
    throw Error("A report snapshot exceeds 4 MB.");
  const r = Toe.parse(e),
    t = ct(r.request.range);
  if ((BR(r.request), r.hash !== W5(r)))
    throw Error("Report snapshot integrity check failed.");
  const n = Array.from(
    { length: t.lastCol - t.firstCol + 1 },
    (s, d) => t.firstCol + d,
  );
  if (hr(n) !== hr(r.columns) || r.widths.length !== n.length)
    throw Error("Invalid report column layout.");
  const o = r.rows.map((s) => s.row),
    a = [...o, ...r.excludedRows].sort((s, d) => s - d);
  if (
    r.rows.length * n.length > 2e3 ||
    a.length !== t.lastRow - t.firstRow + 1 ||
    a.some((s, d) => s !== t.firstRow + d + 1) ||
    o.some((s, d) => d > 0 && s <= o[d - 1])
  )
    throw Error("Invalid report row accounting.");
  for (const s of r.rows) {
    if (s.cells.length !== n.length) throw Error("Invalid report cell layout.");
    s.cells.forEach((d, u) => {
      if (d.address !== `${ut(n[u])}${s.row}`)
        throw Error("Invalid report cell address.");
      d.style && (d.style = oh(d.style));
    });
  }
  if (r.object) {
    const s = {
      id: r.request.sheetId,
      name: r.sheetName,
      rowCount: 1048576,
      colCount: 16384,
    };
    (yC([r.object], s), (r.object = s.objects[0]));
  }
  if (
    (r.request.kind === "object") != !!r.object ||
    (r.object && r.object.id !== r.request.objectId)
  )
    throw Error("Invalid report drawing identity.");
  if (r.hash !== W5(r)) throw Error("Report metadata is not canonical.");
  return r;
}
function Roe(e, r) {
  if (e.reportOutputs !== void 0) {
    const t = array(
      object({
        id: Ap,
        version: number().int().min(1).max(1e6),
        snapshot: unknown(),
      }).strict(),
    )
      .max(20)
      .parse(e.reportOutputs);
    if (new Set(t.map((n) => n.id)).size !== t.length)
      throw Error("Duplicate report output IDs.");
    if (new TextEncoder().encode(JSON.stringify(t)).length > Vv.metadataBytes)
      throw Error("Saved reports exceed 16 MB.");
    r.reportOutputs = t.map((n) => ({ ...n, snapshot: Wg(n.snapshot) }));
  }
  if (e.embeddedObjectInventories !== void 0) {
    const t = array(Noe).max(20).parse(e.embeddedObjectInventories);
    if (
      new Set(t.map((n) => n.id)).size !== t.length ||
      JSON.stringify(t).length > 2e6
    )
      throw Error("Invalid embedded-object inventory count or size.");
    r.embeddedObjectInventories = t;
  }
}
const t0 = Object.freeze({
    expressionCharacters: 4096,
    syntaxNodes: 512,
    depth: 48,
    calculatedColumns: 16,
  }),
  H5 = Object.freeze({
    BLANK: [0, 0],
    TRUE: [0, 0],
    FALSE: [0, 0],
    ISBLANK: [1, 1],
    IF: [2, 3],
    COALESCE: [2, 16],
    DIVIDE: [2, 3],
    SUM: [1, 1],
    AVERAGE: [1, 1],
    MIN: [1, 1],
    MAX: [1, 1],
    DISTINCTCOUNT: [1, 1],
    COUNTROWS: [1, 1],
    SUMX: [2, 2],
    CALCULATE: [1, 17],
    FILTER: [2, 2],
    ALL: [1, 1],
    RELATED: [1, 1],
    SELECTEDVALUE: [1, 2],
    TOTALYTD: [2, 2],
  });
class LR extends Error {
  constructor(r, t) {
    (super(t), (this.code = r), (this.name = "DaxError"));
  }
}
const on = (e, r) => {
  throw new LR(e, `DAX: ${r}`);
};
function UO(e) {
  if (typeof e != "string" || !e.trim() || e.length > t0.expressionCharacters)
    return on("DAX_LIMIT", "Enter an expression of 1–4,096 characters.");
  const r = e.trim().replace(/^=/, ""),
    t = [];
  let n = 0;
  const o = (m, v, g) => {
    const b = n++;
    let y = "",
      w = false;
    for (; n < r.length;) {
      const E = r[n++];
      if (E === v)
        if (r[n] === v) ((y += v), n++);
        else {
          w = true;
          break;
        }
      else y += E;
    }
    (w || on("DAX_SYNTAX", `Unclosed ${m} at position ${b + 1}.`),
      t.push({ kind: g, value: y, position: b }));
  };
  for (; n < r.length;) {
    if (/\s/.test(r[n])) {
      n++;
      continue;
    }
    if (r[n] === '"') {
      o('"', '"', "string");
      continue;
    }
    if (r[n] === "'") {
      o("'", "'", "quoted");
      continue;
    }
    if (r[n] === "[") {
      o("[", "]", "bracket");
      continue;
    }
    const m = r.slice(n),
      v = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(m),
      g = /^[A-Za-z_][A-Za-z0-9_]*/.exec(m),
      b = /^(?:==|<>|<=|>=|&&|\|\||[+\-*/^&=<>(),])/.exec(m),
      y = v ?? g ?? b;
    (y ||
      on(
        "DAX_UNSUPPORTED",
        `Unsupported syntax at position ${n + 1}. Use the documented DAX subset; worksheet and query syntax are not evaluated.`,
      ),
      t.push({
        kind: v ? "number" : g ? "identifier" : "op",
        value: y[0],
        position: n,
      }),
      (n += y[0].length),
      t.length > t0.syntaxNodes * 3 &&
        on("DAX_LIMIT", "Expression token budget exceeded."));
  }
  t.push({ kind: "end", value: "", position: n });
  let a = 0,
    s = 0;
  const d = () => t[a],
    u = () => t[a++],
    l = (m) => {
      u().value !== m && on("DAX_SYNTAX", `Expected ${m}.`);
    },
    c = (m) => (
      ++s > t0.syntaxNodes &&
        on("DAX_LIMIT", "Expression exceeds 512 syntax nodes."),
      m
    ),
    h = {
      "||": 1,
      "&&": 2,
      "=": 4,
      "==": 4,
      "<>": 4,
      "<": 4,
      ">": 4,
      "<=": 4,
      ">=": 4,
      "&": 5,
      "+": 6,
      "-": 6,
      "*": 7,
      "/": 7,
      "^": 9,
    },
    f = (m = 0, v = 0) => {
      v > t0.depth && on("DAX_LIMIT", "Expression nesting exceeds 48 levels.");
      const g = u();
      let b;
      if (g.kind === "number") {
        const y = Number(g.value);
        (Number.isFinite(y) ||
          on("DAX_NUMBER", "Numeric literal exceeds finite precision."),
          (b = c({ kind: "literal", value: y })));
      } else if (g.kind === "string")
        b = c({ kind: "literal", value: g.value });
      else if (g.value === "(") ((b = f(0, v + 1)), l(")"));
      else if (
        g.value === "+" ||
        g.value === "-" ||
        (g.kind === "identifier" && g.value.toUpperCase() === "NOT")
      )
        b = c({
          kind: "unary",
          op: g.value.toUpperCase(),
          value: f(g.value.toUpperCase() === "NOT" ? 3 : 8, v + 1),
        });
      else if (g.kind === "bracket") b = c({ kind: "measure", name: g.value });
      else if (g.kind === "identifier" || g.kind === "quoted")
        if (d().value === "(" && g.kind === "identifier") {
          const y = g.value.toUpperCase();
          (Object.prototype.hasOwnProperty.call(H5, y) ||
            on(
              "DAX_UNSUPPORTED",
              `Unsupported function ${g.value}. No worksheet fallback is used.`,
            ),
            u());
          const w = [];
          if (d().value !== ")")
            do {
              if ((w.push(f(0, v + 1)), d().value !== ",")) break;
              u();
            } while (true);
          l(")");
          const E = H5[y];
          ((w.length < E[0] || w.length > E[1]) &&
            on(
              "DAX_ARGUMENT",
              `${y} expects ${E[0] === E[1] ? E[0] : `${E[0]}–${E[1]}`} arguments in this subset.`,
            ),
            (b = c({ kind: "call", name: y, args: w })));
        } else
          d().kind === "bracket"
            ? (b = c({ kind: "field", table: g.value, column: u().value }))
            : g.kind === "identifier" &&
                ["TRUE", "FALSE"].includes(g.value.toUpperCase())
              ? (b = c({
                  kind: "literal",
                  value: g.value.toUpperCase() === "TRUE",
                }))
              : (b = c({ kind: "table", name: g.value }));
      else
        return on(
          "DAX_SYNTAX",
          `Expected an expression at position ${g.position + 1}.`,
        );
      for (
        ;
        d().kind === "op" && h[d().value] !== void 0 && h[d().value] >= m;
      ) {
        const y = u().value,
          w = h[y];
        b = c({ kind: "binary", op: y, left: b, right: f(w + 1, v + 1) });
      }
      return b;
    },
    p = f();
  return (
    d().kind !== "end" &&
      on(
        "DAX_UNSUPPORTED",
        `Unexpected ${d().value}. VAR, RETURN, EVALUATE, worksheet references and unsupported syntax are not evaluated.`,
      ),
    p
  );
}
function Kh(e, r) {
  (r(e),
    e.kind === "call"
      ? e.args.forEach((t) => Kh(t, r))
      : e.kind === "binary"
        ? (Kh(e.left, r), Kh(e.right, r))
        : e.kind === "unary" && Kh(e.value, r));
}
const PR = "ledgerglass-columnar-v1",
  ta = Object.freeze({
    tables: 8,
    columns: 50,
    rowsPerTable: 500,
    cellsPerTable: 2e3,
    modelCells: 16e3,
    relationships: 12,
    measures: 16,
    filters: 16,
    groupFields: 3,
    groups: 1e3,
    outputCells: 2e4,
    operations: 2e6,
    metadataCharacters: 2e6,
  }),
  S1 = string()
    .min(1)
    .max(100)
    .regex(/^[A-Za-z0-9_.-]+$/),
  jg = string()
    .trim()
    .min(1)
    .max(128)
    .refine((e) => !/[\u0000-\u001f\u007f]/.test(e)),
  Ioe = strictObject({
    name: jg,
    type: _enum(["text", "number", "boolean", "date"]),
    nullable: boolean(),
    expression: string().trim().min(1).max(t0.expressionCharacters).optional(),
  }),
  _oe = strictObject({
    id: S1,
    name: jg,
    sheetId: string().min(1).max(100),
    range: string().min(1).max(50),
    tableId: string().min(1).max(100).optional(),
    columns: array(Ioe).min(1).max(ta.columns),
    dateColumn: jg.optional(),
  }),
  VE = strictObject({ tableId: S1, column: jg }),
  Doe = strictObject({
    id: S1,
    one: VE,
    many: VE,
    cardinality: _enum(["one-to-many", "one-to-one"]),
    direction: literal("single"),
  }),
  Ooe = strictObject({
    id: S1,
    name: jg,
    tableId: S1,
    column: jg.optional(),
    operation: _enum([
      "sum",
      "count",
      "countRows",
      "average",
      "min",
      "max",
      "distinctCount",
      "dax",
    ]),
    expression: string().trim().min(1).max(t0.expressionCharacters).optional(),
  }),
  SU = strictObject({
    engine: literal(PR),
    name: jg,
    tables: array(_oe).min(1).max(ta.tables),
    relationships: array(Doe).max(ta.relationships),
    measures: array(Ooe).max(ta.measures),
  }),
  Moe = union([string().max(32767), number().finite(), boolean(), _null()]),
  joe = strictObject({
    ...VE.shape,
    operator: _enum([
      "equals",
      "notEquals",
      "gt",
      "gte",
      "lt",
      "lte",
      "blank",
      "notBlank",
    ]),
    value: Moe.optional(),
  }),
  TU = strictObject({
    groupBy: array(VE).max(ta.groupFields),
    filters: array(joe).max(ta.filters),
    measureIds: array(S1).max(ta.measures),
  }),
  mp = (e, r) => {
    if (new Set(e.map((t) => t.toLowerCase())).size !== e.length)
      throw Error(`${r} must be unique, including case variants.`);
  };
function WE(e) {
  var o, a;
  const r = SU.safeParse(e);
  if (!r.success)
    throw Error(
      `Invalid data model: ${(o = r.error.issues[0]) == null ? void 0 : o.path.join(".")}: ${(a = r.error.issues[0]) == null ? void 0 : a.message}`,
    );
  const t = r.data;
  if (
    (mp(
      t.tables.map((s) => s.id),
      "Model table IDs",
    ),
    mp(
      t.tables.map((s) => s.name),
      "Model table names",
    ),
    mp(
      t.relationships.map((s) => s.id),
      "Relationship IDs",
    ),
    mp(
      t.measures.map((s) => s.id),
      "Measure IDs",
    ),
    mp(
      t.measures.map((s) => s.name),
      "Measure names",
    ),
    t.tables.reduce(
      (s, d) => s + d.columns.filter((u) => u.expression !== void 0).length,
      0,
    ) > t0.calculatedColumns)
  )
    throw Error("The model supports at most 16 DAX calculated columns.");
  for (const s of t.tables) {
    if (
      (mp(
        s.columns.map((d) => d.name),
        `Columns in ${s.name}`,
      ),
      !s.columns.some((d) => d.expression === void 0))
    )
      throw Error(
        "A model table requires at least one physical source column.",
      );
    if (
      s.dateColumn &&
      !s.columns.some(
        (d) =>
          d.name === s.dateColumn &&
          d.type === "date" &&
          !d.nullable &&
          d.expression === void 0,
      )
    )
      throw Error(
        "A marked calendar requires a nonnullable physical date column.",
      );
  }
  const n = (s) => {
    const d = t.tables.find((l) => l.id === s.tableId),
      u = d == null ? void 0 : d.columns.find((l) => l.name === s.column);
    if (!u) throw Error(`Unknown model field ${s.tableId}.${s.column}.`);
    return u;
  };
  for (const s of t.relationships) {
    if (s.one.tableId === s.many.tableId)
      throw Error(
        "Self relationships are unsupported; use distinct role-playing tables.",
      );
    if (n(s.one).type !== n(s.many).type)
      throw Error(`Relationship ${s.id} requires matching key types.`);
    if (n(s.one).expression !== void 0 || n(s.many).expression !== void 0)
      throw Error(
        "Relationship keys must be physical source columns, not DAX calculated columns.",
      );
  }
  for (const s of t.measures) {
    if (!t.tables.some((d) => d.id === s.tableId))
      throw Error(`Unknown measure table ${s.tableId}.`);
    if (s.operation === "dax") {
      if (!s.expression || s.column !== void 0)
        throw Error(
          "A DAX measure requires an expression and no aggregate column.",
        );
    } else {
      if (s.expression !== void 0)
        throw Error(
          "Unrecognized expression for an aggregate measure; only DAX measures may contain it.",
        );
      if (s.operation === "countRows") {
        if (s.column !== void 0)
          throw Error("Count rows does not take a column.");
      } else {
        const d = n({ tableId: s.tableId, column: s.column ?? "" });
        if (
          ["sum", "average", "min", "max"].includes(s.operation) &&
          d.type !== "number"
        )
          throw Error(`Measure ${s.name} requires a number column.`);
      }
    }
  }
  return t;
}
function UR(e, r) {
  var o, a;
  const t = TU.safeParse(e);
  if (!t.success)
    throw Error(
      `Invalid model query: ${(o = t.error.issues[0]) == null ? void 0 : o.message}`,
    );
  const n = t.data;
  (mp(n.measureIds, "Selected measures"),
    mp(
      n.groupBy.map((s) => JSON.stringify(s)),
      "Group fields",
    ));
  for (const s of [...n.groupBy, ...n.filters])
    if (!(
      (a = r.tables.find((d) => d.id === s.tableId)) != null &&
      a.columns.some((d) => d.name === s.column)
    ))
      throw Error(`Unknown query field ${s.tableId}.${s.column}.`);
  if (n.measureIds.some((s) => !r.measures.some((d) => d.id === s)))
    throw Error("A selected model measure is missing.");
  return n;
}
const lk = string().regex(/^[a-f0-9]{64}$/),
  Foe = strictObject({
    definition: SU,
    query: TU,
    sourceHash: lk,
    definitionHash: lk,
    outputHash: lk,
    outputRange: string().min(1).max(50),
    createdAt: string().datetime(),
    sourceRevision: number().int().nonnegative(),
    includeHidden: boolean(),
    totals: array(number().finite().nullable()).max(ta.measures),
  });
function Boe(e, r) {
  var o, a;
  const t = e;
  let n = 0;
  if (t.dataModel !== void 0) {
    if (((n += JSON.stringify(t.dataModel).length), n > ta.metadataCharacters))
      throw Error("Data model metadata exceeds 2,000,000 characters.");
    r.dataModel = WE(t.dataModel);
  }
  for (const s of r.sheets) {
    const d =
      (o = t.sheets.find((c) => c.id === s.id)) == null
        ? void 0
        : o.dataModelReport;
    if (d === void 0) continue;
    if (((n += JSON.stringify(d).length), n > ta.metadataCharacters))
      throw Error("Data model metadata exceeds 2,000,000 characters.");
    const u = Foe.safeParse(d);
    if (!u.success)
      throw Error(
        `Invalid model report: ${(a = u.error.issues[0]) == null ? void 0 : a.message}`,
      );
    const l = u.data;
    if (
      ((l.definition = WE(l.definition)),
      (l.query = UR(l.query, l.definition)),
      l.definitionHash !== hr(l.definition) ||
        l.totals.length !== l.query.measureIds.length)
    )
      throw Error(
        "Model report definition or measure metadata is inconsistent.",
      );
    s.dataModelReport = l;
  }
}
const Cd = Object.freeze({
    bytes: 4 * 1024 * 1024,
    retainedBytes: 12 * 1024 * 1024,
    imports: 10,
    tables: 20,
    rows: 1e3,
    columns: 100,
    cells: 2e4,
    depth: 32,
    nodes: 1e5,
    confidence: 0.9,
  }),
  NU = ["json", "xml", "pdf", "png", "jpeg", "html"],
  Il = (e) => string().max(e),
  RU = string().regex(/^[a-f0-9]{64}$/),
  IU = union([Il(32767), number().finite(), boolean(), _null()]),
  Loe = object({
    page: number().int().min(1).max(1e4),
    x: number().min(0).max(1),
    y: number().min(0).max(1),
    width: number().positive().max(1),
    height: number().positive().max(1),
  })
    .strict()
    .refine(
      (e) => e.x + e.width <= 1.000001 && e.y + e.height <= 1.000001,
      "Source region must fit its page.",
    );
function Poe(e, r) {
  if (typeof e != "string") return false;
  try {
    const t = JSON.parse(e);
    return r === "array"
      ? Array.isArray(t)
      : !!t && typeof t == "object" && !Array.isArray(t);
  } catch {
    return false;
  }
}
const Uoe = object({
    value: IU,
    type: _enum(["string", "number", "boolean", "null", "array", "object"]),
    path: Il(2048),
    lexical: Il(32767).optional(),
    declaredType: Il(200).optional(),
    missing: boolean().optional(),
    confidence: number().min(0).max(1),
    region: Loe.optional(),
  })
    .strict()
    .refine(
      (e) =>
        e.type === "array" || e.type === "object"
          ? Poe(e.value, e.type)
          : e.type === "null"
            ? e.value === null
            : typeof e.value === e.type,
      "Native value and type disagree.",
    )
    .refine(
      (e) => !e.missing || (e.value === null && e.type === "null"),
      "Missing fields must be null placeholders.",
    ),
  _U = object({
    id: Il(200).min(1),
    name: Il(200).min(1),
    path: Il(2048),
    ambiguous: boolean(),
    columns: array(object({ name: Il(200).min(1), path: Il(2048) }).strict())
      .min(1)
      .max(Cd.columns),
    rows: array(array(Uoe).max(Cd.columns)).min(1).max(Cd.rows),
  })
    .strict()
    .superRefine((e, r) => {
      (new Set(e.columns.map((t) => t.path)).size !== e.columns.length ||
        e.rows.some((t) => t.length !== e.columns.length)) &&
        r.addIssue({
          code: "custom",
          message: "Table columns and rows must align with unique paths.",
        });
    }),
  HE = object({
    fileName: Il(255)
      .min(1)
      .refine((e) => !/[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(e)),
    format: _enum(NU),
    byteLength: number().int().min(1).max(Cd.bytes),
    sha256: RU,
    base64: Il(Math.ceil(Cd.bytes / 3) * 4),
  }).strict(),
  DU = object({
    schemaVersion: literal(1),
    fileName: HE.shape.fileName,
    format: _enum(NU),
    byteLength: HE.shape.byteLength,
    sha256: RU,
    tables: array(_U).max(Cd.tables),
    method: _enum(["native-json", "native-xml", "document-adapter"]),
    extractor: Il(200).min(1),
    pageCount: number().int().min(1).max(1e4).optional(),
    warnings: array(Il(2048)).max(30),
  }).strict(),
  OU = object({
    decision: _enum(["accept", "exclude", "correct"]),
    value: IU.optional(),
  })
    .strict()
    .refine(
      (e) => e.decision !== "correct" || e.value !== void 0,
      "A corrected value is required.",
    ),
  $oe = object({
    id: Il(100).min(1),
    sheetId: Il(100).min(1),
    sheetName: Il(31).min(1),
    importedAt: string().datetime(),
    original: HE,
    method: DU.shape.method,
    extractor: Il(200).min(1),
    pageCount: number().int().min(1).max(1e4).optional(),
    table: _U,
    resolutions: record(string().regex(/^\d+:\d+$/), OU),
    ambiguityReviewed: boolean(),
  }).strict();
function kC(e) {
  const r = DU.parse(e);
  let t = 0;
  if (new Set(r.tables.map((n) => n.id)).size !== r.tables.length)
    throw Error("Source table IDs must be unique.");
  for (const n of r.tables)
    for (const o of n.rows)
      for (const a of o) {
        if (++t > Cd.cells) throw Error("Source cell budget exceeded.");
        if (
          r.method === "document-adapter" &&
          (!a.region ||
            !r.pageCount ||
            a.region.page > r.pageCount ||
            (["png", "jpeg"].includes(r.format) && r.pageCount !== 1))
        )
          throw Error(
            "Document cells require a source page and bounding region.",
          );
      }
  if (
    (r.format === "json" && r.method !== "native-json") ||
    (r.format === "xml" && r.method !== "native-xml") ||
    (!["json", "xml"].includes(r.format) && r.method !== "document-adapter")
  )
    throw Error("Source format and extraction method disagree.");
  return r;
}
function zoe(e, r, t) {
  let n = "";
  for (let a = 0; a < e.length; a += 8192)
    n += String.fromCharCode(...e.subarray(a, a + 8192));
  const o = {
    fileName: r,
    format: t,
    byteLength: e.length,
    sha256: iu(e),
    base64: btoa(n),
  };
  return ($R(o), o);
}
function $R(e) {
  var a;
  const r = HE.parse(e),
    t = (a = r.fileName.split(".").at(-1)) == null ? void 0 : a.toLowerCase();
  if ((t === "jpg" ? "jpeg" : t === "htm" ? "html" : t) !== r.format)
    throw Error("Source extension and format disagree.");
  if (
    r.base64.length !== Math.ceil(r.byteLength / 3) * 4 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(r.base64)
  )
    throw Error("Invalid retained source encoding.");
  const n = atob(r.base64);
  if (btoa(n) !== r.base64 || n.length !== r.byteLength)
    throw Error("Retained source byte length changed.");
  const o = Uint8Array.from(n, (s) => s.charCodeAt(0));
  if (iu(o) !== r.sha256) throw Error("Retained source checksum mismatch.");
  return o;
}
function zR(e, r) {
  if (e.confidence < Cd.confidence && !r)
    throw Error("Review every low-confidence cell before importing.");
  if (!r) return e.value;
  const t = OU.parse(r);
  return t.decision === "exclude"
    ? null
    : t.decision === "correct"
      ? t.value
      : e.value;
}
const uk = new WeakSet();
function Voe(e, r) {
  if (e.sourceImports === void 0) return;
  if (!Array.isArray(e.sourceImports) || e.sourceImports.length > Cd.imports)
    throw Error("Source import count exceeded.");
  if (
    e.sourceImports.some((o) => !o || typeof o != "object" || !uk.has(o)) &&
    new TextEncoder().encode(JSON.stringify(e.sourceImports)).length >
      24 * 1024 * 1024
  )
    throw Error("Source provenance metadata budget exceeded.");
  let t = 0;
  const n = new Set();
  if (
    ((r.sourceImports = e.sourceImports.map((o) => {
      if (o && typeof o == "object" && uk.has(o)) {
        const d = o;
        if (((t += d.original.byteLength), n.has(d.id)))
          throw Error("Duplicate source import ID.");
        return (n.add(d.id), d);
      }
      const a = $oe.parse(o);
      if (
        ($R(a.original),
        kC({
          schemaVersion: 1,
          fileName: a.original.fileName,
          format: a.original.format,
          byteLength: a.original.byteLength,
          sha256: a.original.sha256,
          tables: [a.table],
          method: a.method,
          extractor: a.extractor,
          ...(a.pageCount ? { pageCount: a.pageCount } : {}),
          warnings: [],
        }),
        a.table.ambiguous && !a.ambiguityReviewed)
      )
        throw Error("Ambiguous source tables require review.");
      const s = new Set();
      if (
        (a.table.rows.forEach((d, u) =>
          d.forEach((l, c) => {
            const h = `${u}:${c}`;
            (s.add(h), zR(l, a.resolutions[h]));
          }),
        ),
        Object.keys(a.resolutions).some((d) => !s.has(d)))
      )
        throw Error("Unknown source review cell.");
      if (n.has(a.id)) throw Error("Duplicate source import ID.");
      return (n.add(a.id), (t += a.original.byteLength), MU(a), uk.add(a), a);
    })),
    t > Cd.retainedBytes)
  )
    throw Error("Retained source byte budget exceeded.");
}
function MU(e) {
  Object.freeze(e);
  for (const r of Object.values(e)) r && typeof r == "object" && MU(r);
}
const VR = Object.freeze({
    maxBytes: 10 * 1024 * 1024,
    maxParts: 2e3,
    maxDetails: 1e5,
  }),
  q5 = new WeakSet();
function SC(e) {
  if (
    !e ||
    typeof e != "object" ||
    e.schemaVersion !== 1 ||
    typeof e.fileName != "string" ||
    !/\.xlsx$/i.test(e.fileName) ||
    e.fileName.length > 255 ||
    /[\\/\x00-\x1f\x7f]/.test(e.fileName) ||
    !Number.isSafeInteger(e.byteLength) ||
    e.byteLength < 22 ||
    e.byteLength > VR.maxBytes ||
    typeof e.sha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(e.sha256) ||
    typeof e.base64 != "string" ||
    e.base64.length !== Math.ceil(e.byteLength / 3) * 4 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(e.base64)
  )
    throw Error(
      "Invalid retained XLSX source metadata or 10 MiB source limit exceeded.",
    );
  let r;
  try {
    r = atob(e.base64);
  } catch {
    throw Error("Invalid retained XLSX source encoding.");
  }
  if (btoa(r) !== e.base64 || r.length !== e.byteLength)
    throw Error("Retained XLSX source byte length or encoding changed.");
  const t = Uint8Array.from(r, (n) => n.charCodeAt(0));
  if (iu(t) !== e.sha256)
    throw Error("Retained XLSX source checksum does not match its bytes.");
  return t;
}
function Woe(e, r) {
  if (e.byteLength > VR.maxBytes)
    throw Error("Original XLSX exceeds the 10 MiB retention limit.");
  let t = "";
  for (let o = 0; o < e.length; o += 8192)
    t += String.fromCharCode(...e.subarray(o, o + 8192));
  const n = {
    schemaVersion: 1,
    fileName: r,
    byteLength: e.length,
    sha256: iu(e),
    base64: btoa(t),
  };
  return (SC(n), Object.freeze(n), q5.add(n), n);
}
function Hoe(e, r) {
  if (e.xlsxSource === void 0) return;
  const t = e.xlsxSource;
  if (t && typeof t == "object" && q5.has(t)) {
    r.xlsxSource = t;
    return;
  }
  (SC(t),
    (r.xlsxSource = Object.freeze({
      schemaVersion: 1,
      fileName: t.fileName,
      byteLength: t.byteLength,
      sha256: t.sha256,
      base64: t.base64,
    })),
    q5.add(r.xlsxSource));
}
function WR(e) {
  if (!e.xlsxSource)
    throw Error(
      "This workbook has no retained XLSX source. Open an XLSX file to retain its original bytes.",
    );
  return {
    bytes: SC(e.xlsxSource),
    fileName: e.xlsxSource.fileName,
    sha256: e.xlsxSource.sha256,
  };
}
const HR = ["xls", "xlsb", "xlsm", "xltx", "xltm", "ods"],
  $O = (e) =>
    typeof e == "string" &&
    e.length > 0 &&
    e.length <= 255 &&
    !/[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(e),
  jU = new WeakSet();
function qR(e) {
  if (
    !e ||
    typeof e != "object" ||
    !HR.includes(e.format) ||
    !$O(e.fileName) ||
    !e.fileName.toLowerCase().endsWith("." + e.format) ||
    !e.conversion ||
    !["template-adapter", "external-conversion"].includes(
      e.conversion.method,
    ) ||
    (e.conversion.method === "template-adapter" && e.format !== "xltx") ||
    !$O(e.conversion.convertedFileName) ||
    !/\.xlsx$/i.test(e.conversion.convertedFileName) ||
    typeof e.conversion.convertedSha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(e.conversion.convertedSha256) ||
    typeof e.conversion.reviewedAt != "string" ||
    !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(e.conversion.reviewedAt) ||
    !Number.isFinite(Date.parse(e.conversion.reviewedAt)) ||
    e.conversion.equivalence !== "not-verified"
  )
    throw Error("Invalid alternate workbook source or conversion provenance.");
  return SC({ ...e, fileName: "retained.xlsx" });
}
function FU(e, r, t, n) {
  const o = {
    ...Woe(e, "retained.xlsx"),
    fileName: r,
    format: t,
    conversion: { ...n },
  };
  return (qR(o), Object.freeze(o.conversion), Object.freeze(o), jU.add(o), o);
}
function qoe(e, r) {
  if (e.alternateSource === void 0) return;
  const t = e.alternateSource;
  if (jU.has(t)) {
    r.alternateSource = t;
    return;
  }
  const n = qR(t);
  r.alternateSource = FU(n, t.fileName, t.format, t.conversion);
}
function Xoe(e) {
  const r = e.alternateSource;
  if (!r)
    throw Error("This workbook has no retained alternate-format original.");
  return {
    bytes: qR(r),
    fileName: r.fileName,
    sha256: r.sha256,
    format: r.format,
  };
}
const ol = Object.freeze({
    maxFileBytes: 10 * 1024 * 1024,
    maxSourceBytes: 256 * 1024,
    maxArtifactBytes: 2 * 1024 * 1024,
    maxStoredBytes: 10 * 1024 * 1024,
    maxArtifacts: 20,
    maxRuns: 100,
    maxCommands: 100,
    maxCells: 500,
    maxTokens: 2e4,
  }),
  Goe = [
    "vba",
    "office-script",
    "office-script-container",
    "excel-container",
    "legacy-container",
  ];
function Zoe(e) {
  let r = "";
  for (let t = 0; t < e.length; t += 8192)
    r += String.fromCharCode(...e.subarray(t, t + 8192));
  return btoa(r);
}
function X5(e) {
  if (
    typeof e != "string" ||
    e.length > Math.ceil(ol.maxArtifactBytes / 3) * 4 ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(e)
  )
    throw Error("Invalid automation artifact encoding or size.");
  const r = atob(e);
  if (btoa(r) !== e)
    throw Error("Automation artifact is not canonical base64.");
  return Uint8Array.from(r, (t) => t.charCodeAt(0));
}
const qh = (e, r, t) => {
    if (typeof e != "string" || !e.trim() || e.length > r)
      throw Error(`Invalid automation ${t}.`);
    return e;
  },
  ck = (e) => {
    if (typeof e != "string" || !/^[a-f0-9]{64}$/.test(e))
      throw Error("Invalid automation source hash.");
    return e;
  },
  Gb = (e, r, t) => {
    if (!Number.isSafeInteger(e) || Number(e) < r || Number(e) > t)
      throw Error("Invalid automation evidence count.");
    return Number(e);
  },
  zO = (e) => {
    const r = qh(e, 40, "timestamp");
    if (!/^\d{4}-\d\d-\d\dT/.test(r) || !Number.isFinite(Date.parse(r)))
      throw Error("Invalid automation timestamp.");
    return r;
  },
  Koe = (e) => {
    const r = qh(e, 40, "approved range"),
      t = /^([A-Z]{1,3})([1-9]\d*):([A-Z]{1,3})([1-9]\d*)$/.exec(r);
    if (!t) throw Error("Invalid automation approved range.");
    const n = (u) => [...u].reduce((l, c) => l * 26 + c.charCodeAt(0) - 64, 0),
      [o, a, s, d] = [n(t[1]), Number(t[2]), n(t[3]), Number(t[4])];
    if (
      o > s ||
      a > d ||
      s > 16384 ||
      d > 1048576 ||
      (s - o + 1) * (d - a + 1) > ol.maxCells
    )
      throw Error("Invalid automation approved range.");
    return r;
  };
function VO(e) {
  const r = qh(e, 255, "file name");
  if (/[\\/\x00-\x1f\x7f]/.test(r) || r === "." || r === "..")
    throw Error("Invalid automation file name.");
  return r;
}
function Yoe(e, r) {
  if (e.automationArtifacts !== void 0) {
    if (
      !Array.isArray(e.automationArtifacts) ||
      e.automationArtifacts.length > ol.maxArtifacts
    )
      throw Error("Keep at most 20 automation artifacts.");
    let t = 0;
    const n = new Set();
    r.automationArtifacts = e.automationArtifacts.map((o) => {
      if (!o || typeof o != "object" || Array.isArray(o))
        throw Error("Invalid automation artifact.");
      const a = o,
        s = qh(a.id, 100, "artifact ID");
      if (n.has(s)) throw Error("Duplicate automation artifact ID.");
      n.add(s);
      const d = X5(a.base64);
      if (
        ((t += d.length),
        !d.length ||
          d.length !== a.byteLength ||
          d.length > ol.maxArtifactBytes ||
          t > ol.maxStoredBytes)
      )
        throw Error(
          "Automation artifacts exceed their 2 MB per-file or 10 MB retained-byte budget.",
        );
      if (iu(d) !== ck(a.sha256))
        throw Error(
          "Automation artifact checksum does not match its preserved bytes.",
        );
      if (!Goe.includes(a.kind))
        throw Error("Unknown automation artifact kind.");
      return {
        id: s,
        fileName: VO(a.fileName),
        kind: a.kind,
        sha256: a.sha256,
        byteLength: d.length,
        base64: a.base64,
        retainedAt: zO(a.retainedAt),
        retainedBy: qh(a.retainedBy, 100, "retainer"),
      };
    });
  }
  if (e.automationRuns !== void 0) {
    if (
      !Array.isArray(e.automationRuns) ||
      e.automationRuns.length > ol.maxRuns
    )
      throw Error("Automation run evidence is limited to 100 runs.");
    const t = new Set();
    r.automationRuns = e.automationRuns.map((n) => {
      if (!n || typeof n != "object" || Array.isArray(n))
        throw Error("Invalid automation run evidence.");
      const o = n,
        a = qh(o.id, 100, "run ID");
      if (t.has(a)) throw Error("Duplicate automation run ID.");
      if (
        (t.add(a),
        !["vba", "office-script"].includes(o.language) ||
          o.adapter !== "literal-ranges-v1")
      )
        throw Error("Unknown automation migration adapter.");
      if (
        !Array.isArray(o.sourceLines) ||
        o.sourceLines.length > ol.maxCommands
      )
        throw Error("Invalid automation source lines.");
      const s = Gb(o.beforeRevision, 0, Number.MAX_SAFE_INTEGER - 2),
        d = Gb(o.afterRevision, 1, Number.MAX_SAFE_INTEGER - 1);
      if (d !== s + 1)
        throw Error("Automation run revisions are inconsistent.");
      if (o.sourceLines.length !== o.commands)
        throw Error(
          "Automation source-line evidence must match the command count.",
        );
      return {
        id: a,
        eventId: qh(o.eventId, 100, "event ID"),
        sourceName: VO(o.sourceName),
        sourceSha256: ck(o.sourceSha256),
        codeSha256: ck(o.codeSha256),
        language: o.language,
        adapter: o.adapter,
        actor: qh(o.actor, 100, "actor"),
        timestamp: zO(o.timestamp),
        beforeRevision: s,
        afterRevision: d,
        sheetId: qh(o.sheetId, 100, "sheet ID"),
        range: Koe(o.range),
        commands: Gb(o.commands, 1, ol.maxCommands),
        changedCells: Gb(o.changedCells, 0, ol.maxCells),
        sourceLines: o.sourceLines.map((u) => Gb(u, 1, ol.maxSourceBytes + 1)),
      };
    });
  }
}
function Joe(e, r) {
  if (e === void 0) return;
  if (!e || typeof e != "object" || Array.isArray(e))
    throw Error("Invalid delimited source footprint.");
  const t = e;
  if (
    !Number.isInteger(t.records) ||
    Number(t.records) < 1 ||
    Number(t.records) > Mt.maxRows ||
    !Number.isInteger(t.columns) ||
    Number(t.columns) < 1 ||
    Number(t.columns) > Mt.maxCols ||
    typeof t.hasHeader != "boolean"
  )
    throw Error(
      "Delimited source footprint exceeds worksheet bounds or has invalid header metadata.",
    );
  r.delimitedSource = {
    records: Number(t.records),
    columns: Number(t.columns),
    hasHeader: t.hasHeader,
  };
}
const B2 = 100,
  Qoe = [
    "deleted-cells",
    "deleted-sheet",
    "overwritten",
    "missing-sheet",
    "out-of-bounds",
  ],
  WO = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  HO = (e, r) => Object.keys(e).every((t) => r.includes(t)),
  dk = (e, r) =>
    typeof e == "string" &&
    !!e.trim() &&
    e.length <= r &&
    !/[\u0000-\u001f]/.test(e);
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
const Ls = {
    inputs: 32,
    outputs: 16,
    scenarios: 100,
    axisValues: 100,
    trials: 400,
    dependencyCells: 500,
    workUnits: 5e4,
    savedAnalyses: 40,
  },
  Vu = (e) => JSON.stringify([e.sheetId, e.address]),
  qE = (e) => typeof e == "number" && Number.isFinite(e);
function Yh(e, r) {
  if (
    !e ||
    typeof e != "object" ||
    Array.isArray(e) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(e)) ||
    Object.keys(e).some((t) => !r.includes(t))
  )
    throw Error("Invalid what-if analysis fields.");
  return e;
}
function Ep(e, r = 100) {
  if (
    typeof e != "string" ||
    !e.trim() ||
    e.length > r ||
    /[\u0000-\u001f]/.test(e)
  )
    throw Error(
      `What-if names must be nonempty text of at most ${r} characters.`,
    );
  return e.trim();
}
function l1(e, r, t = 1) {
  if (!Array.isArray(e) || e.length < t || e.length > r)
    throw Error(`What-if lists require ${t}–${r} entries.`);
  return Array.from(e);
}
function e1(e, r) {
  return l1(e, r).map((t) => {
    if (!qE(t))
      throw Error("Every what-if input value must be a finite number.");
    return t;
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
function PU(e) {
  return e.kind === "scenarios"
    ? e.scenarios.map((r) => ({ name: r.name, inputs: r.values }))
    : e.kind === "one-variable"
      ? e.values.map((r) => ({ name: String(r), inputs: [r] }))
      : e.columnValues.flatMap((r) =>
          e.rowValues.map((t) => ({ name: `${t} / ${r}`, inputs: [t, r] })),
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
const wg = {
    areaCells: 5e3,
    pages: 100,
    titleRows: 100,
    titleColumns: 50,
    breaksPerAxis: 100,
    renderedCells: 1e5,
  },
  XE = {
    header: `&[Workbook]
Engagement: &[Engagement] · Period: &[Period]`,
    footer: "Revision &[Revision] · &A · Page &P of &N",
  },
  GR = /&&|&\[(Workbook|Revision|Engagement|Period)\]|&[PNA]/g;
function UU(e, r) {
  var t, n;
  return e === "Workbook"
    ? ((r == null ? void 0 : r.name) ?? "[Workbook not available]")
    : e === "Revision"
      ? r
        ? String(r.revision)
        : "[Revision not available]"
      : e === "Engagement"
        ? ((t = r == null ? void 0 : r.metadata) == null
            ? void 0
            : t.engagement) || "[Engagement not set]"
        : ((n = r == null ? void 0 : r.metadata) == null ? void 0 : n.period) ||
          "[Period not set]";
}
function rae(e, r, t, n, o) {
  return e.replace(GR, (a, s) =>
    s
      ? UU(s, r)
      : a === "&&"
        ? "&"
        : a === "&A"
          ? t.name
          : String(a === "&P" ? n : o),
  );
}
function XO(e, r) {
  const t = e.replace(GR, (n, o) => {
    if (!o) return n;
    if (!r)
      throw Error(
        "Exporting print identity fields requires the current workbook snapshot.",
      );
    return UU(o, r).replace(/&/g, "&&");
  });
  if (t.length > 255)
    throw Error(
      "Resolved print identifiers exceed Excel's 255-character header/footer limit. Shorten the header, footer or workbook properties before exporting.",
    );
  return t;
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
function G5(e, r) {
  return r === "row"
    ? `${e.start + 1}:${e.end + 1}`
    : `${ut(e.start)}:${ut(e.end)}`;
}
function GE(e, r) {
  if (e !== void 0) {
    if (
      !Array.isArray(e) ||
      e.length > wg.breaksPerAxis ||
      e.some(
        (t, n) =>
          !Number.isSafeInteger(t) ||
          t < 1 ||
          t >= r ||
          (n > 0 && t <= e[n - 1]),
      )
    )
      throw Error(
        "Manual page breaks require at most 100 distinct ascending boundaries within each worksheet axis (before row/column 2 or later).",
      );
    return e.length ? [...e] : void 0;
  }
}
function nae(e) {
  return Object.freeze({
    workbookId: e.id,
    revision: e.revision,
    snapshotGuard: JSON.stringify(e),
  });
}
function $U(e, r) {
  if (
    r.workbookId !== e.id ||
    r.revision !== e.revision ||
    r.snapshotGuard !== JSON.stringify(e)
  )
    throw Error(
      "The workbook changed after page setup opened. Close and reopen page setup.",
    );
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
function ZR(e) {
  return {
    orientation: "portrait",
    paper: "letter",
    margins: {
      top: 0.75,
      bottom: 0.75,
      left: 0.7,
      right: 0.7,
      header: 0.3,
      footer: 0.3,
    },
    printArea: e ? Ta(e) : "A1",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    scale: 100,
    gridlines: true,
    headings: false,
    pageOrder: "overThenDown",
    header: "",
    footer: "",
  };
}
function Rx(e, r) {
  if (e === void 0) return;
  if (!e || typeof e != "object" || Array.isArray(e))
    throw new Error("Invalid page layout.");
  const t = e,
    n = t.margins,
    o = (b, y, w) => {
      if (typeof b != "number" || !Number.isFinite(b) || b < y || b > w)
        throw new Error("Page setup contains an out-of-range number.");
      return b;
    };
  if (
    !["portrait", "landscape"].includes(String(t.orientation)) ||
    !["letter", "a4", "legal"].includes(String(t.paper)) ||
    !n ||
    typeof n != "object"
  )
    throw new Error("Invalid paper, orientation, or margins.");
  if (
    t.pageOrder !== void 0 &&
    t.pageOrder !== "overThenDown" &&
    t.pageOrder !== "downThenOver"
  )
    throw Error(
      "Unsupported print page order. Choose Over, then down or Down, then over.",
    );
  const a = Object.fromEntries(
      ["top", "bottom", "left", "right", "header", "footer"].map((b) => [
        b,
        o(n[b], 0, 3),
      ]),
    ),
    s =
      t.paper === "a4"
        ? [8.2677, 11.6929]
        : t.paper === "legal"
          ? [8.5, 14]
          : [8.5, 11],
    [d, u] = t.orientation === "landscape" ? s.reverse() : s;
  if (a.left + a.right >= d - 0.25 || a.top + a.bottom >= u - 0.25)
    throw new Error("Margins leave no printable page area.");
  if (typeof t.printArea != "string") throw new Error("Enter a print area.");
  const l = ct(t.printArea);
  if (l.lastRow >= r.rowCount || l.lastCol >= r.colCount)
    throw new Error("Print area exceeds the worksheet.");
  if (
    ["fitToPage", "gridlines", "headings"].some((b) => typeof t[b] != "boolean")
  )
    throw new Error("Invalid print option.");
  if (
    ["header", "footer"].some(
      (b) => typeof t[b] != "string" || String(t[b]).length > 255,
    )
  )
    throw new Error("Headers and footers must be at most 255 characters.");
  const c = o(t.fitToWidth, 0, 10),
    h = o(t.fitToHeight, 0, 10),
    f = o(t.scale, 10, 400);
  if (![c, h, f].every(Number.isInteger))
    throw new Error("Page fitting and scale must be whole numbers.");
  const p = Fg(t.repeatRows, "row", r),
    m = Fg(t.repeatColumns, "column", r),
    v = GE(t.rowBreaks, r.rowCount),
    g = GE(t.columnBreaks, r.colCount);
  r.pageLayout = {
    orientation: t.orientation,
    paper: t.paper,
    margins: a,
    printArea: t.printArea,
    fitToPage: t.fitToPage,
    fitToWidth: c,
    fitToHeight: h,
    scale: f,
    gridlines: t.gridlines,
    headings: t.headings,
    header: t.header,
    footer: t.footer,
    ...(t.pageOrder !== void 0 ? { pageOrder: t.pageOrder } : {}),
    ...(p ? { repeatRows: G5(p, "row") } : {}),
    ...(m ? { repeatColumns: G5(m, "column") } : {}),
    ...(v ? { rowBreaks: v } : {}),
    ...(g ? { columnBreaks: g } : {}),
  };
}
function oae(e, r, t, n = {}) {
  if (n.readOnly) throw Error("This workbook is read-only.");
  n.expectedSnapshot && $U(e, n.expectedSnapshot);
  const o = Nr(e, r);
  jn(o);
  const a = { ...o };
  return (
    Rx(t, a),
    zU(a, a.pageLayout),
    JSON.stringify(a.pageLayout) === JSON.stringify(o.pageLayout)
      ? e
      : { ...e, sheets: e.sheets.map((s) => (s.id === r ? a : s)) }
  );
}
function KR(e) {
  const r =
    e.paper === "a4"
      ? [793.7007874015749, 1122.5196850393702]
      : e.paper === "legal"
        ? [816, 1344]
        : [816, 1056];
  return e.orientation === "landscape"
    ? { width: r[1], height: r[0] }
    : { width: r[0], height: r[1] };
}
function zU(e, r) {
  var I, O;
  const t = { ...e };
  (Rx(r, t), (r = t.pageLayout));
  const n = ct(r.printArea);
  if (
    (n.lastRow - n.firstRow + 1) * (n.lastCol - n.firstCol + 1) >
    wg.areaCells
  )
    throw Error(
      "Print preview supports at most 5,000 cells per print area. Choose a smaller print area.",
    );
  const a = KR(r),
    s = a.width - (r.margins.left + r.margins.right) * 96,
    d = a.height - (r.margins.top + r.margins.bottom) * 96,
    u = Fg(r.repeatRows, "row", e),
    l = Fg(r.repeatColumns, "column", e);
  for (const j of e.merges ?? []) {
    const z = ct(j);
    for (const [V, D, B] of [
      [u, z.firstRow, z.lastRow],
      [l, z.firstCol, z.lastCol],
    ])
      if (V && D <= V.end && B >= V.start && (D < V.start || B > V.end))
        throw Error(
          "A merged area crosses a repeated-title boundary. Unmerge it or adjust the print titles first.",
        );
  }
  for (const j of e.objects ?? []) {
    const z = _t(j.anchor),
      V = XR(e, j);
    for (const [D, B, _, C, T] of [
      [u, "row", z.row, V.offsetY, V.height],
      [l, "column", z.col, V.offsetX, V.width],
    ]) {
      if (!D) continue;
      const N = ul(B === "row" ? e.rowCount : e.colCount, cn(e, B));
      if (N.isHidden(_)) continue;
      const R = (G) => (B === "row" ? gg(e, G) : Ip(e, G)),
        L = (G) => {
          let H = _,
            Q = 0;
          for (; G < 0;) {
            const X = N.next(H, -1);
            if (X === null) return H;
            if (((H = X), (G += R(H)), ++Q > 1e3))
              throw Error(
                "Drawing print inspection exceeds its geometry budget.",
              );
          }
          for (; G >= R(H);) {
            G -= R(H);
            const X = N.next(H, 1);
            if (X === null) return H;
            if (((H = X), ++Q > 1e3))
              throw Error(
                "Drawing print inspection exceeds its geometry budget.",
              );
          }
          return H;
        },
        $ = L(C),
        U = L(C + T - 1e-4);
      if ($ <= D.end && U >= D.start && ($ < D.start || U > D.end))
        throw Error(
          "A drawing crosses a repeated-title boundary. Move or resize the drawing, or adjust the print titles first.",
        );
    }
  }
  const c = (j, z, V) => {
      const D = ul(j === "row" ? e.rowCount : e.colCount, cn(e, j)),
        B = [];
      for (let _ = z; _ <= V; _++) D.isHidden(_) || B.push(_);
      return B;
    },
    h = c("row", n.firstRow, n.lastRow),
    f = c("column", n.firstCol, n.lastCol);
  if (!h.length || !f.length)
    throw Error(
      "The print area has no visible cells. Unhide its rows or columns, or choose another print area.",
    );
  const p = u ? c("row", u.start, u.end) : [],
    m = l ? c("column", l.start, l.end) : [],
    v = (j, z, V, D, B) => {
      const _ = new Set(z),
        C = j.filter((U) => !_.has(U)),
        T = z.reduce((U, G) => U + D(G), 0);
      if (T > V + 1e-7 || (C.length && T >= V - 1e-7))
        throw Error(
          "Repeated titles leave no space for worksheet data. Reduce the titles, margins or scale.",
        );
      if (!C.length) return [z];
      const N = [];
      let R = [],
        L = T,
        $;
      for (const U of C) {
        const G = D(U),
          H = $ !== void 0 && B.some((Q) => Q > $ && Q <= U);
        if (
          (R.length &&
            (L + G > V + 1e-7 || H) &&
            (N.push([...z, ...R]), (R = []), (L = T)),
          z.length && L + G > V + 1e-7 && !R.length)
        )
          throw Error(
            "A data row or column cannot fit beside the repeated titles. Reduce the scale or title dimensions.",
          );
        (R.push(U), (L += G), ($ = U));
      }
      return (R.length && N.push([...z, ...R]), N);
    },
    g = (j) => ({
      rows: v(
        h,
        p,
        d / j - (r.headings ? 22 : 0),
        (z) => gg(e, z),
        r.fitToPage ? [] : (r.rowBreaks ?? []),
      ),
      columns: v(
        f,
        m,
        s / j - (r.headings ? 32 : 0),
        (z) => Ip(e, z),
        r.fitToPage ? [] : (r.columnBreaks ?? []),
      ),
    });
  let b = r.fitToPage ? 1 : r.scale / 100;
  if (r.fitToPage && (r.fitToWidth || r.fitToHeight)) {
    const j = (z) => {
      try {
        const V = g(z);
        return (
          (!r.fitToWidth || V.columns.length <= r.fitToWidth) &&
          (!r.fitToHeight || V.rows.length <= r.fitToHeight)
        );
      } catch {
        return false;
      }
    };
    if (!j(1)) {
      let z = 0,
        V = 1;
      for (let D = 0; D < 40; D++) {
        const B = (z + V) / 2;
        j(B) ? (z = B) : (V = B);
      }
      b = z;
    }
  }
  if (!Number.isFinite(b) || b <= 0)
    throw Error(
      "The print content cannot fit the selected paper. Reduce the print area or titles.",
    );
  const y = g(b);
  if (y.rows.length * y.columns.length > wg.pages)
    throw Error(
      "This layout exceeds 100 pages. Reduce the print area or fit it to fewer pages.",
    );
  if (
    y.rows.reduce((j, z) => j + z.length, 0) *
      y.columns.reduce((j, z) => j + z.length, 0) >
    wg.renderedCells
  )
    throw Error(
      "Repeated print content exceeds 100,000 rendered cells. Reduce titles, pages or the print area.",
    );
  const E = !!(u || l),
    A = (j, z) => ({
      firstRow: Math.min(...j),
      lastRow: Math.max(...j),
      firstCol: Math.min(...z),
      lastCol: Math.max(...z),
      ...(E ? { rows: j, columns: z } : {}),
    }),
    k =
      r.pageOrder === "downThenOver"
        ? y.columns.flatMap((j) => y.rows.map((z) => A(z, j)))
        : y.rows.flatMap((j) => y.columns.map((z) => A(j, z))),
    F =
      r.fitToPage &&
      (((I = r.rowBreaks) == null ? void 0 : I.length) ?? 0) +
        (((O = r.columnBreaks) == null ? void 0 : O.length) ?? 0)
        ? [
            "Fit to pages ignores saved manual page breaks. Turn it off to use the breaks.",
          ]
        : [];
  return { pages: k, scale: b, ...(F.length ? { notices: F } : {}) };
}
function aae(e, r, t) {
  if (!r.pageLayout) return [];
  const n = { ...r };
  Rx(r.pageLayout, n);
  const o = n.pageLayout,
    a = XO(o.header, t),
    s = XO(o.footer, t),
    d = [o.header, o.footer].some((u) =>
      [...u.matchAll(GR)].some((l) => !!l[1]),
    );
  return (
    (e.pageSetup = {
      ...e.pageSetup,
      orientation: o.orientation,
      paperSize: o.paper === "a4" ? 9 : o.paper === "legal" ? 5 : 1,
      margins: o.margins,
      printArea: o.printArea,
      fitToPage: o.fitToPage,
      fitToWidth: o.fitToWidth,
      fitToHeight: o.fitToHeight,
      scale: o.scale,
      showGridLines: o.gridlines,
      showRowColHeaders: o.headings,
      pageOrder: o.pageOrder ?? "overThenDown",
      ...(o.repeatRows ? { printTitlesRow: o.repeatRows } : {}),
      ...(o.repeatColumns ? { printTitlesColumn: o.repeatColumns } : {}),
    }),
    (e.rowBreaks = (o.rowBreaks ?? []).map((u) => ({
      id: u,
      max: 16383,
      man: 1,
    }))),
    (e.headerFooter = { oddHeader: a, oddFooter: s }),
    d
      ? [
          "Print workbook/engagement/period/revision fields are exported as literal text from this workbook snapshot. Native JSON retains the live identity tokens; XLSX does not update them after later edits.",
        ]
      : []
  );
}
const VU = [
    "TRIM",
    "CLEAN",
    "UPPER",
    "LOWER",
    "PROPER",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "FIND",
    "SEARCH",
    "SUBSTITUTE",
    "REPLACE",
    "VALUE",
    "NUMBERVALUE",
    "TEXT",
    "CONCAT",
    "TEXTJOIN",
    "EXACT",
    "TEXTBEFORE",
    "TEXTAFTER",
  ],
  sae = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e)),
  gp = () => {
    throw Error(
      "Invalid or over-capacity text cleanup evidence. Keep the original native backup.",
    );
  };
function L2(e, r) {
  (!sae(e) || Object.keys(e).some((t) => !r.includes(t))) && gp();
}
function WU(e) {
  return (
    e === null ||
    typeof e == "boolean" ||
    (typeof e == "number" && Number.isFinite(e)) ||
    (typeof e == "string" && e.length <= 32767)
  );
}
function jm(e, r) {
  return typeof e == "string" && e.length > 0 && e.length <= r;
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
function YR(e) {
  if (!Array.isArray(e) || !e.length || e.length > 12)
    throw Error("Choose 1–12 text cleanup steps.");
  return e.map(
    (r) => (
      L2(r, ["function", "args"]),
      (!VU.includes(r.function) ||
        !Array.isArray(r.args) ||
        r.args.length > 5 ||
        r.args.some((t) => t === null || !WU(t)) ||
        JSON.stringify(r.args).length > 4096) &&
        gp(),
      { function: r.function, args: [...r.args] }
    ),
  );
}
function HU(e, r) {
  if (e === void 0) return;
  (!Array.isArray(e) || e.length > 100 || JSON.stringify(e).length > 5e6) &&
    gp();
  const t = new Set();
  let n = 0;
  r.textCleanups = e.map((o) => {
    L2(o, [
      "id",
      "workbookId",
      "sourceRevision",
      "sourceSheetId",
      "sourceSheetName",
      "outputSheetId",
      "outputSheetName",
      "createdAt",
      "steps",
      "semantics",
      "rows",
    ]);
    for (const l of ["id", "workbookId", "sourceSheetId", "outputSheetId"])
      jm(o[l], 100) || gp();
    ((!jm(o.sourceSheetName, 31) ||
      !jm(o.outputSheetName, 31) ||
      t.has(o.id) ||
      !Number.isSafeInteger(o.sourceRevision) ||
      Number(o.sourceRevision) < 0 ||
      o.semantics !== "unicode-v2-en-US-explicit-separators" ||
      !Array.isArray(o.rows) ||
      !o.rows.length ||
      o.rows.length > 500 ||
      (n += o.rows.length) > 1e4 ||
      typeof o.createdAt != "string" ||
      !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(o.createdAt) ||
      !Number.isFinite(Date.parse(o.createdAt)) ||
      new Date(o.createdAt).toISOString() !== o.createdAt) &&
      gp(),
      t.add(o.id));
    const a = YR(o.steps),
      s = new Set(),
      d = new Set(),
      u = o.rows.map((l) => {
        (L2(l, [
          "sourceAddress",
          "original",
          "result",
          "outputAddress",
          "origin",
        ]),
          (!GO(l.sourceAddress) ||
            !GO(l.outputAddress) ||
            s.has(l.sourceAddress) ||
            d.has(l.outputAddress) ||
            !WU(l.result)) &&
            gp(),
          s.add(l.sourceAddress),
          d.add(l.outputAddress));
        const c = l.original === void 0 ? void 0 : hi(l.original);
        c &&
          typeof c.raw == "string" &&
          c.raw.startsWith("=") &&
          !c.literal &&
          gp();
        let h;
        if (l.origin !== void 0) {
          L2(l.origin, [
            "id",
            "sourceId",
            "sourceName",
            "sourceSheet",
            "originalRow",
            "kind",
          ]);
          const f = l.origin;
          ((!jm(f.id, 100) ||
            !jm(f.sourceId, 100) ||
            !jm(f.sourceName, 255) ||
            !jm(f.sourceSheet, 31) ||
            !Number.isInteger(f.originalRow) ||
            Number(f.originalRow) < 1 ||
            Number(f.originalRow) > 1048576 ||
            !["snapshot", "file"].includes(String(f.kind))) &&
            gp(),
            (h = { ...f }));
        }
        return {
          sourceAddress: l.sourceAddress,
          outputAddress: l.outputAddress,
          result: l.result,
          ...(c ? { original: c } : {}),
          ...(h ? { origin: h } : {}),
        };
      });
    return { ...o, steps: a, rows: u };
  });
}
function lae(e, r) {
  var n;
  if (!((n = e.textCleanups) != null && n.length)) return r;
  const t = new Map(e.textCleanups.map((o) => [o.id, o]));
  for (const o of r.textCleanups ?? []) {
    const a = t.get(o.id);
    if (a && JSON.stringify(a) !== JSON.stringify(o))
      throw Error("Saved original text cleanup evidence cannot be changed.");
  }
  return {
    ...r,
    textCleanups: [
      ...e.textCleanups,
      ...(r.textCleanups ?? []).filter((o) => !t.has(o.id)),
    ],
  };
}
const Z5 = {
    inferencePatterns: 8e3,
    inferenceCandidates: 30,
    evidenceCharacters: 5e6,
  },
  ZE = number().int().min(0).max(49),
  Bh = array(ZE)
    .min(1)
    .max(50)
    .refine(
      (e) => new Set(e).size === e.length,
      "Choose distinct column positions.",
    ),
  qU = string().max(32767),
  JR = union([qU, number().finite(), boolean(), _null()]),
  XU = _enum(["none", "upper", "lower", "proper"]),
  uae = strictObject({
    column: ZE,
    mode: _enum(["whole", "token", "slice", "before", "after"]),
    delimiter: string().min(1).max(16).default(" "),
    index: number().int().min(-50).max(49).default(0),
    start: number().int().min(0).max(32767).default(0),
    length: number().int().min(1).max(32767).default(1),
  }),
  GU = strictObject({
    parts: array(uae).min(1).max(8),
    separator: string().max(100),
    prefix: string().max(1e3),
    suffix: string().max(1e3),
    casing: XU,
  }),
  ZU = strictObject({ row: number().int().min(1).max(1048576), value: qU }),
  cae = array(ZU)
    .min(2)
    .max(10)
    .refine(
      (e) => new Set(e.map((r) => r.row)).size === e.length,
      "Use examples from distinct source rows.",
    ),
  dae = discriminatedUnion("kind", [
    strictObject({
      kind: literal("trim"),
      columns: Bh,
      mode: _enum(["edges", "collapse"]),
    }),
    strictObject({ kind: literal("clean"), columns: Bh }),
    strictObject({ kind: literal("case"), columns: Bh, mode: XU }),
    strictObject({
      kind: literal("replace"),
      columns: Bh,
      find: string().min(1).max(1e3),
      replacement: string().max(1e3),
    }),
    strictObject({
      kind: literal("split"),
      column: ZE,
      mode: _enum(["delimited", "fixed"]),
      delimiters: string().max(16),
      collapse: boolean(),
      breaks: array(number().int().min(1).max(32766)).max(49),
    }),
    strictObject({
      kind: literal("concat"),
      columns: Bh,
      separator: string().max(100),
      name: string().min(1).max(100),
      skipBlank: boolean(),
    }),
    strictObject({
      kind: literal("convert"),
      columns: Bh,
      to: _enum(["text", "number", "integer", "boolean", "date"]),
      decimal: _enum([".", ","]),
      group: _enum(["", ".", ",", " ", "'"]),
      dateOrder: _enum(["MDY", "DMY"]),
    }),
    strictObject({
      kind: literal("fill"),
      columns: Bh,
      direction: _enum(["down", "up", "value"]),
      value: JR,
      groupColumns: array(ZE).max(50),
    }),
    strictObject({ kind: literal("select"), columns: Bh }),
    strictObject({
      kind: literal("pattern"),
      name: string().min(1).max(100),
      pattern: GU,
      examples: cae,
      inferenceColumns: Bh.max(8),
    }),
  ]),
  fae = strictObject({
    sourceSheetId: string().min(1).max(100),
    sourceRange: string().min(1).max(50),
    header: boolean(),
    outputName: string().min(1).max(31),
    steps: array(dae).max(12),
  });
function KU(e) {
  var n;
  const r = fae.safeParse(e);
  if (!r.success)
    throw Error(
      `Invalid field transformation: ${((n = r.error.issues[0]) == null ? void 0 : n.message) ?? "review the recipe"}`,
    );
  const t = r.data;
  if (
    ((t.outputName = t.outputName.trim()),
    !t.outputName ||
      /[\u0000-\u001f\u007f\\/\[\]*?:]/.test(t.outputName) ||
      /^'|'$/.test(t.outputName))
  )
    throw Error("Choose a valid output worksheet name of 1–31 characters.");
  for (const o of t.steps) {
    if (o.kind === "split") {
      if (
        o.mode === "delimited" &&
        (![...o.delimiters].length ||
          [...o.delimiters].length > 8 ||
          /["\r\n\0]/.test(o.delimiters))
      )
        throw Error(
          "Choose 1–8 delimiter characters without quotes, line breaks or NUL.",
        );
      if (
        o.mode === "fixed" &&
        (!o.breaks.length ||
          o.breaks.some((a, s) => s > 0 && a <= o.breaks[s - 1]))
      )
        throw Error(
          "Fixed-width breaks must be distinct increasing character positions.",
        );
    }
    if (o.kind === "convert" && o.to !== "text" && o.decimal === o.group)
      throw Error("Decimal and grouping separators must differ.");
    if (
      o.kind === "fill" &&
      new Set(o.groupColumns).size !== o.groupColumns.length
    )
      throw Error("Choose distinct fill group columns.");
    if (o.kind === "fill" && o.columns.some((a) => o.groupColumns.includes(a)))
      throw Error(
        "Fill targets must differ from the group columns. Use a separate preceding step to prepare group keys.",
      );
  }
  return t;
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
function YU(e, r) {
  if (e === void 0) return;
  if (
    !Array.isArray(e) ||
    e.length > 100 ||
    JSON.stringify(e).length > Z5.evidenceCharacters
  )
    throw Error(
      "Field transformation evidence exceeds its native history budget. Keep the original backup.",
    );
  const t = new Set();
  let n = 0;
  r.fieldTransforms = e.map((o) => {
    var f;
    const a = vae.safeParse(o);
    if (!a.success)
      throw Error(
        `Invalid field transformation evidence: ${(f = a.error.issues[0]) == null ? void 0 : f.message}`,
      );
    const s = a.data,
      d = KU(s.request),
      u = ct(d.sourceRange);
    if (
      new Date(s.createdAt).toISOString() !== s.createdAt ||
      s.sources.length !==
        (u.lastRow - u.firstRow + 1) * (u.lastCol - u.firstCol + 1)
    )
      throw Error(
        "Invalid field evidence timestamp or incomplete original source population.",
      );
    if (t.has(s.id) || (n += s.outputs.length) > 2e4)
      throw Error("Duplicate or over-capacity field transformation evidence.");
    if (
      (t.add(s.id),
      s.sourceSheetId !== d.sourceSheetId || s.outputSheetName !== d.outputName)
    )
      throw Error("Field transformation evidence does not match its recipe.");
    const l = new Set(),
      c = s.sources.map((p) => {
        const m = _t(p.address);
        if (
          m.row < u.firstRow ||
          m.row > u.lastRow ||
          m.col < u.firstCol ||
          m.col > u.lastCol
        )
          throw Error(
            "An original field source is outside the recorded range.",
          );
        if (l.has(p.address)) throw Error("Duplicate original field source.");
        l.add(p.address);
        const v = p.original === void 0 ? void 0 : hi(p.original);
        if (
          v &&
          typeof v.raw == "string" &&
          v.raw.startsWith("=") &&
          !v.literal
        )
          throw Error("Field source evidence must contain reviewed values.");
        return { ...p, ...(v ? { original: v } : {}) };
      }),
      h = new Set();
    for (const p of s.outputs) {
      if (h.has(p.address))
        throw Error("Duplicate transformed output address.");
      if (
        (h.add(p.address),
        ((p.type === "number" || p.type === "date") &&
          typeof p.value != "number") ||
          (p.type === "boolean" && typeof p.value != "boolean") ||
          (p.type === "text" && typeof p.value != "string") ||
          (p.type === "blank" && p.value !== null))
      )
        throw Error("Transformed output type does not match its value.");
    }
    for (const p of [...s.outputs, ...s.exceptions])
      if (
        p.row < u.firstRow + 1 + Number(d.header) ||
        p.row > u.lastRow + 1 ||
        p.sources.some((m) => !l.has(m)) ||
        new Set(p.sources).size !== p.sources.length
      )
        throw Error("Unknown or repeated original field source locator.");
    return { ...s, request: d, sources: c };
  });
}
function bae(e, r) {
  var n;
  if (!((n = e.fieldTransforms) != null && n.length)) return r;
  const t = new Map(e.fieldTransforms.map((o) => [o.id, o]));
  for (const o of r.fieldTransforms ?? [])
    if (t.has(o.id) && JSON.stringify(t.get(o.id)) !== JSON.stringify(o))
      throw Error(
        "Saved original field transformation evidence cannot be changed.",
      );
  return {
    ...r,
    fieldTransforms: [
      ...e.fieldTransforms,
      ...(r.fieldTransforms ?? []).filter((o) => !t.has(o.id)),
    ],
  };
}
const JU = new WeakSet(),
  NC = new WeakSet(),
  YE = new WeakMap(),
  QU = new WeakMap(),
  yae = new Set(["general", "number", "currency", "percent", "date"]),
  wae = (e, r) => Object.prototype.hasOwnProperty.call(e, r),
  Vh = (e) => !!e && typeof e == "object" && !Array.isArray(e),
  Hn = (e) => {
    throw new Error(e);
  },
  Bu = (e, r, t = 32767) =>
    typeof e == "string" && e.length <= t
      ? e
      : Hn(`${r} must be text of at most ${t} characters.`),
  P2 = (e, r, t, n) =>
    typeof e == "number" && Number.isInteger(e) && e >= r && e <= t
      ? e
      : Hn(`${n} must be an integer between ${r} and ${t}.`);
function ut(e) {
  P2(e, 0, Mt.maxCols - 1, "Column");
  let r = "";
  for (let t = e + 1; t > 0; t = Math.floor((t - 1) / 26))
    r = String.fromCharCode(65 + ((t - 1) % 26)) + r;
  return r;
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
function RC(e, r = true) {
  let t = null;
  for (const [n, o] of Object.entries(e.cells)) {
    const a = o.raw !== null && o.raw !== "",
      s =
        r &&
        (o.numberFormat ||
          o.hyperlink ||
          o.note ||
          o.tickmark ||
          o.numFmt ||
          o.format ||
          Object.keys(o.style ?? {}).length);
    if (!a && !s) continue;
    const { row: d, col: u } = _t(n);
    t
      ? ((t.firstRow = Math.min(t.firstRow, d)),
        (t.lastRow = Math.max(t.lastRow, d)),
        (t.firstCol = Math.min(t.firstCol, u)),
        (t.lastCol = Math.max(t.lastCol, u)))
      : (t = { firstRow: d, lastRow: d, firstCol: u, lastCol: u });
  }
  return t;
}
function Ta(e, r = true) {
  const t = RC(e, r);
  if (!t) return "A1";
  const n = `${ut(t.firstCol)}${t.firstRow + 1}`,
    o = `${ut(t.lastCol)}${t.lastRow + 1}`;
  return n === o ? n : `${n}:${o}`;
}
function Ix(e) {
  const { firstRow: r, lastRow: t, firstCol: n, lastCol: o } = ct(e);
  if ((t - r + 1) * (o - n + 1) > Mt.maxCells)
    return Hn("Range is too large. Select at most 100,000 cells.");
  const a = [];
  for (let s = r; s <= t; s++)
    for (let d = n; d <= o; d++) a.push(`${ut(d)}${s + 1}`);
  return a;
}
function Nr(e, r) {
  return e.sheets.find((t) => t.id === r) ?? Hn(`Sheet not found: ${r}`);
}
function Ku(e) {
  return !e.trim() || e.length > 31 || /[\\/\[\]*?:]/.test(e)
    ? Hn(
        "Sheet names must contain 1–31 characters and cannot contain \\ / [ ] * ? : .",
      )
    : {
        id: crypto.randomUUID(),
        name: e.trim(),
        cells: {},
        rowCount: Mt.maxRows,
        colCount: Mt.maxCols,
        frozenRows: 0,
      };
}
function hi(e) {
  return U2(e, false);
}
function U2(e, r) {
  var o;
  if (
    r &&
    es(e) &&
    JU.has(e) &&
    !e.hyperlink &&
    !((o = e.style) != null && o.binding)
  )
    return e;
  if (!Vh(e) || !wae(e, "raw")) return Hn("A cell must have a raw value.");
  const t = e.raw;
  if (
    t !== null &&
    typeof t != "string" &&
    typeof t != "boolean" &&
    (typeof t != "number" || !Number.isFinite(t))
  )
    return Hn("Cell values must be text, finite numbers, booleans, or null.");
  typeof t == "string" &&
    Bu(
      t,
      "Cell value",
      t.startsWith("=") && e.literal !== true
        ? Ja.maxFormulaLength
        : Ja.maxCellTextLength,
    );
  const n = { raw: t };
  if (e.literal !== void 0) {
    if (typeof e.literal != "boolean")
      return Hn("Cell literal flag must be boolean.");
    n.literal = e.literal;
  }
  if (
    (e.numberFormat !== void 0 && (n.numberFormat = l0(e.numberFormat)),
    e.numFmt !== void 0 &&
      (n.numFmt = Bu(e.numFmt, "Excel number format", 200)),
    e.format !== void 0)
  ) {
    if (typeof e.format != "string" || !yae.has(e.format))
      return Hn("Unsupported cell format.");
    n.format = e.format;
  }
  return (
    e.hyperlink !== void 0 && (n.hyperlink = X3(e.hyperlink)),
    e.note !== void 0 && (n.note = Bu(e.note, "Cell note", 1e4)),
    e.tickmark !== void 0 && (n.tickmark = Bu(e.tickmark, "Tickmark", 40)),
    e.style !== void 0 && (n.style = oh(e.style)),
    n
  );
}
function p0(e, r, t = "Preparer", n = "Edit cells", o = {}) {
  if (!Array.isArray(r) || r.length > Mt.maxPatches)
    return Hn(
      `A transaction may contain at most ${Mt.maxPatches} cell changes.`,
    );
  if (r.length === 0) return e;
  (Bu(t, "Actor", 100), Bu(n, "Action", 200));
  const a = new Set();
  let s;
  const d = r.map((p) => {
      var E, A;
      if (!Vh(p)) return Hn("Invalid cell patch.");
      const m = Nr(e, p.sheetId);
      jn(m);
      const { row: v, col: g } = _t(p.address);
      if (v >= m.rowCount || g >= m.colCount)
        return Hn(`Cell ${p.address} is outside sheet dimensions.`);
      const b = `${ut(g)}${v + 1}`;
      _N(e, hi(p.cell).style);
      const y = m.cells[b];
      if (
        ((y == null ? void 0 : y.raw) ?? null) !==
          ((E = p.cell) == null ? void 0 : E.raw) ||
        !!(y != null && y.literal) != !!((A = p.cell) != null && A.literal)
      ) {
        s ?? (s = Vn(e));
        const k = s.getSpillAnchor(m.id, b);
        if (k && k !== b) {
          const F = r.find((I) => I.sheetId === m.id && I.address === k);
          if (
            !F ||
            (!F.cell.literal &&
              typeof F.cell.raw == "string" &&
              F.cell.raw.startsWith("="))
          )
            return Hn(
              `Cell ${b} belongs to the array formula in ${k}. Edit the array formula instead.`,
            );
        }
      }
      const w = `${m.id}:${b}`;
      return a.has(w)
        ? Hn("A transaction cannot change the same cell twice.")
        : (a.add(w),
          {
            sheetId: m.id,
            address: b,
            before: m.cells[b] ? hi(m.cells[b]) : void 0,
            after: hi(p.cell),
          });
    }),
    u = d.filter((p) => !p.before).length;
  if (
    e.sheets.reduce(
      (p, m) =>
        p +
        (es(m.cells)
          ? (QU.get(m.cells) ?? Object.keys(m.cells).length)
          : Object.keys(m.cells).length),
      0,
    ) +
      u >
    Mt.maxCells
  )
    return Hn("Workbook exceeds the 100,000 populated cell limit.");
  const l = e.sheets.map((p) => {
      const m = d.filter((g) => g.sheetId === p.id);
      if (!m.length) return p;
      const v = { ...p.cells };
      return (
        m.forEach((g) => {
          v[g.address] = g.after;
        }),
        { ...p, cells: v }
      );
    }),
    c = hR(e, { ...e, sheets: l });
  if (c.sheets !== l)
    for (const p of c.sheets) {
      const m = Nr(e, p.id);
      for (const [v, g] of Object.entries(p.cells)) {
        const b = d.find((y) => y.sheetId === p.id && y.address === v);
        if (b) b.after = hi(g);
        else if (
          JSON.stringify(m.cells[v]) !== JSON.stringify(g) &&
          (d.push({
            sheetId: p.id,
            address: v,
            before: m.cells[v] ? hi(m.cells[v]) : void 0,
            after: hi(g),
          }),
          d.length > Mt.maxPatches)
        )
          return Hn(
            "Renaming this table field affects more than 500 cells. Reduce the dependent formulas before renaming it.",
          );
      }
    }
  const h = v0(
    c,
    d.filter((p) => {
      var m, v;
      return (
        ((m = p.before) == null ? void 0 : m.raw) !== p.after.raw ||
        ((v = p.before) == null ? void 0 : v.literal) !== p.after.literal
      );
    }),
    o.validationApproval,
  );
  if (
    (h && (n = `${n.slice(0, 140)} (accepted ${h} validation warnings)`),
    e.events.length >= 2e3)
  )
    return Hn(
      "Change history has reached its 2,000 transaction limit. Export the workbook before starting a new workpaper.",
    );
  if (
    e.events.reduce((p, m) => p + m.changes.length, 0) + d.length >
    Mt.maxCells
  )
    return Hn(
      "Change history has reached its 100,000 cell-change limit. Export the workbook before starting a new workpaper.",
    );
  if (e.revision >= Number.MAX_SAFE_INTEGER - 1)
    return Hn("Workbook revision limit reached.");
  const f = {
    ...gx(c),
    revision: e.revision + 1,
    sheets: c.sheets,
    events: [
      ...e.events,
      {
        id: crypto.randomUUID(),
        timestamp: new Date().toISOString(),
        actor: t,
        action: n,
        changes: d,
      },
    ],
  };
  return (
    es(e) &&
      NC.has(e) &&
      c.sheets === l &&
      d.every((p) => {
        var m;
        return (
          !p.after.hyperlink && !((m = p.after.style) != null && m.binding)
        );
      }) &&
      t$(f),
    SR(e, f),
    f
  );
}
function rn(e) {
  return e$(e, false);
}
function e$(e, r) {
  if (r && es(e) && NC.has(e)) return e;
  if (!Vh(e) || e.schemaVersion !== 1)
    return Hn("Unsupported workbook schema. Expected schemaVersion 1.");
  const t = Bu(e.id, "Workbook ID", 100),
    n = Bu(e.name, "Workbook name", 200);
  if (!t || !n.trim()) return Hn("Workbook ID and name are required.");
  const o = P2(e.revision, 0, Number.MAX_SAFE_INTEGER - 1, "Revision");
  if (!Array.isArray(e.sheets) || !e.sheets.length)
    return Hn("A workbook must contain at least one sheet.");
  let a = 0;
  const s = new Set(),
    d = new Set(),
    u = e.sheets.map((p) => {
      if (!Vh(p)) return Hn("Invalid sheet.");
      const m = Bu(p.id, "Sheet ID", 100),
        v = Bu(p.name, "Sheet name", 31);
      if (
        !m ||
        !v.trim() ||
        /[\\/\[\]*?:]/.test(v) ||
        s.has(m) ||
        d.has(v.toLowerCase())
      )
        return Hn(
          "Sheet IDs and names must be nonempty and unique; sheet names cannot contain Excel-reserved characters.",
        );
      (s.add(m), d.add(v.toLowerCase()));
      const g = P2(p.rowCount, 1, Mt.maxRows, "Row count"),
        b = P2(p.colCount, 1, Mt.maxCols, "Column count");
      if (!Vh(p.cells)) return Hn("Sheet cells must be an address map.");
      const y = r && es(p) ? YE.get(p) : void 0;
      if (y !== void 0) {
        if (((a += y), a > Mt.maxCells))
          return Hn("Workbook exceeds the 100,000 populated cell limit.");
        const A = { ...p };
        return (YE.set(A, y), A);
      }
      const w = {};
      for (const [A, k] of Object.entries(p.cells)) {
        const F = _t(A);
        if (A !== `${ut(F.col)}${F.row + 1}` || F.row >= g || F.col >= b)
          return Hn(`Invalid or out-of-bounds cell: ${A}`);
        if (++a > Mt.maxCells)
          return Hn("Workbook exceeds the 100,000 populated cell limit.");
        w[A] = U2(k, r);
      }
      const E = {
        id: m,
        name: v,
        rowCount: Mt.maxRows,
        colCount: Mt.maxCols,
        cells: w,
      };
      if (p.columnWidths !== void 0) {
        if (!Vh(p.columnWidths))
          return Hn("Column widths must be an index map.");
        E.columnWidths = {};
        for (const [A, k] of Object.entries(p.columnWidths)) {
          if (
            !/^(0|[1-9]\d*)$/.test(A) ||
            Number(A) >= b ||
            typeof k != "number" ||
            !Number.isFinite(k) ||
            k < ux ||
            k > cx
          )
            return Hn(
              "Column widths must be 40–600 pixels for an existing column.",
            );
          E.columnWidths[A] = k;
        }
      }
      if (p.rowHeights !== void 0) {
        if (
          !Vh(p.rowHeights) ||
          (Object.getPrototypeOf(p.rowHeights) !== Object.prototype &&
            Object.getPrototypeOf(p.rowHeights) !== null)
        )
          return Hn("Row heights must be a plain index map.");
        E.rowHeights = {};
        for (const [A, k] of Object.entries(p.rowHeights)) {
          if (
            !/^(0|[1-9]\d*)$/.test(A) ||
            Number(A) >= g ||
            typeof k != "number" ||
            !Number.isFinite(k) ||
            k < b1 ||
            k > Ig
          )
            return Hn(
              `Row heights must be ${b1}–${Ig} pixels for an existing row.`,
            );
          E.rowHeights[A] = k;
        }
      }
      if (p.hidden !== void 0) {
        if (typeof p.hidden != "boolean")
          return Hn("Sheet hidden flag must be boolean.");
        E.hidden = p.hidden;
      }
      return (
        VQ(p, E),
        TB(p, E, { rows: g, cols: b }),
        K3(p, E, { rows: g, cols: b }),
        wP(p, E),
        Joe(p.delimitedSource, E),
        Ax(p, E),
        p.objects !== void 0 && yC(p.objects, E),
        p.pageLayout !== void 0 && Rx(p.pageLayout, E),
        E
      );
    });
  if (u.every((p) => p.hidden))
    return Hn("At least one sheet must be visible.");
  if (!Array.isArray(e.events) || e.events.length > 2e3)
    return Hn("Workbook history must be an array of at most 2,000 events.");
  const l = new Set();
  let c = 0;
  const h = e.events.map((p) => {
      if (!Vh(p)) return Hn("Invalid history event.");
      const m = Bu(p.id, "Event ID", 100),
        v = Bu(p.timestamp, "Event timestamp", 40);
      return !m || l.has(m) || !Number.isFinite(Date.parse(v))
        ? Hn("History events require unique IDs and valid timestamps.")
        : (l.add(m),
          !Array.isArray(p.changes) ||
          p.changes.length > Mt.maxPatches ||
          (c += p.changes.length) > Mt.maxCells
            ? Hn("Invalid history change count.")
            : {
                id: m,
                timestamp: v,
                actor: Bu(p.actor, "Actor", 100),
                action: Bu(p.action, "Action", 200),
                changes: p.changes.map((g) => {
                  if (
                    !Vh(g) ||
                    typeof g.sheetId != "string" ||
                    !g.sheetId.length ||
                    g.sheetId.length > 100 ||
                    typeof g.address != "string"
                  )
                    return Hn("Invalid history cell reference.");
                  const b = _t(g.address);
                  return g.address !== `${ut(b.col)}${b.row + 1}`
                    ? Hn("History cell address is not canonical.")
                    : g.before === void 0 && g.after === void 0
                      ? Hn("History change must have a before or after value.")
                      : {
                          sheetId: g.sheetId,
                          address: g.address,
                          ...(g.before !== void 0
                            ? { before: U2(g.before, r) }
                            : {}),
                          ...(g.after !== void 0
                            ? { after: U2(g.after, r) }
                            : {}),
                        };
                }),
              });
    }),
    f = { schemaVersion: 1, id: t, name: n, revision: o, sheets: u, events: h };
  return (
    Lie(e, f),
    e.datePolicy !== void 0 && (f.datePolicy = ts(e.datePolicy)),
    fU(e.savedFilterViews, f),
    GF(e, f),
    uoe(e, f),
    boe(e, f),
    Boe(e, f),
    Aoe(e, f),
    Roe(e, f),
    aie(e, f),
    Yoe(e, f),
    Hoe(e, f),
    qoe(e, f),
    Voe(e, f),
    Nie(e, f),
    HU(e.textCleanups, f),
    YU(e.fieldTransforms, f),
    bC(e.tableEvidence, f),
    e.whatIfAnalyses !== void 0 && (f.whatIfAnalyses = tae(e.whatIfAnalyses)),
    e.definedNames !== void 0 && (f.definedNames = eg(e.definedNames, f)),
    jJ(e, f),
    nR(e.structureProtection, f),
    px(e.metadata, f),
    UP(f),
    EB(e, f),
    JJ(e.styleCatalog, f),
    e.formulaWatches !== void 0 && (f.formulaWatches = TC(e.formulaWatches, f)),
    zie(e, f),
    FJ(e, f),
    f
  );
}
function rh(e) {
  const r = e$(e, true);
  return NC.has(r) && es(r) ? r : t$(r);
}
function t$(e) {
  Lg(e);
  for (const r of e.sheets) {
    if (YE.has(r)) continue;
    const t = Object.values(r.cells);
    for (const n of t) JU.add(n);
    (QU.set(r.cells, t.length),
      t.some((n) => {
        var o;
        return n.hyperlink || ((o = n.style) == null ? void 0 : o.binding);
      }) || YE.set(r, t.length));
  }
  return (NC.add(e), e);
}

export {
  $B,
  $P,
  $Q,
  $R,
  $U,
  $a,
  AC,
  AP,
  AQ,
  AR,
  Ad,
  Af,
  Aie,
  Ay,
  B2,
  BB,
  BE,
  BQ,
  BR,
  CC,
  CE,
  CP,
  CQ,
  CR,
  CU,
  Cd,
  Cf,
  Cp,
  DB,
  DR,
  Dg,
  Dw,
  E1,
  EC,
  EO,
  EQ,
  ER,
  EU,
  Eie,
  F2,
  FN,
  FO,
  FP,
  FR,
  FU,
  Fc,
  Fg,
  Fw,
  GE,
  GU,
  Gne,
  Gw,
  Gy,
  H3,
  H5,
  HP,
  HQ,
  HR,
  HU,
  Hie,
  Hne,
  I5,
  IB,
  IQ,
  Ip,
  Ix,
  J3,
  JN,
  JS,
  Jf,
  Jl,
  Jre,
  KJ,
  KR,
  KS,
  KU,
  Kh,
  Kie,
  Ku,
  L5,
  LB,
  LQ,
  LR,
  LU,
  Ls,
  Mp,
  NB,
  NO,
  NQ,
  Np,
  Nr,
  Nx,
  O5,
  OR,
  Og,
  Op,
  P1,
  PB,
  PR,
  PU,
  Pne,
  Pv,
  Q3,
  QN,
  QP,
  Qie,
  Qne,
  RC,
  RN,
  RP,
  Rc,
  Rn,
  Rx,
  SE,
  SQ,
  SR,
  TC,
  TP,
  Ta,
  Tx,
  Ty,
  U1,
  UB,
  UO,
  UR,
  Une,
  Us,
  Uv,
  VR,
  VU,
  Vg,
  Vie,
  Vn,
  Vne,
  Vu,
  Vv,
  W3,
  W5,
  WE,
  WQ,
  WR,
  Wg,
  Wie,
  Wne,
  Ww,
  X3,
  X5,
  XE,
  XP,
  XR,
  Xie,
  Xoe,
  Xw,
  YR,
  YU,
  Z5,
  ZP,
  ZR,
  ZU,
  Zne,
  Zoe,
  _B,
  _N,
  _g,
  _t,
  aD,
  aP,
  aR,
  aae,
  bB,
  bU,
  bae,
  bie,
  c6,
  cR,
  cie,
  cn,
  ct,
  d6,
  dQ,
  die,
  dne,
  eQ,
  eU,
  eh,
  eie,
  ek,
  eoe,
  f0,
  fB,
  ff,
  fie,
  fl,
  gB,
  gQ,
  gU,
  gg,
  gie,
  gx,
  h0,
  hB,
  hR,
  hi,
  iQ,
  iae,
  ig,
  jB,
  jE,
  jN,
  jQ,
  jR,
  jn,
  jne,
  jw,
  k1,
  kC,
  kE,
  kQ,
  l0,
  lD,
  lae,
  mU,
  md,
  mie,
  mx,
  nQ,
  nae,
  nee,
  nk,
  oD,
  oQ,
  oR,
  oae,
  oh,
  oie,
  ol,
  on,
  p0,
  pB,
  pQ,
  pR,
  pf,
  q6,
  qE,
  qQ,
  qne,
  rQ,
  rU,
  rae,
  rh,
  rn,
  s0,
  sQ,
  sR,
  sU,
  sie,
  t0,
  t5,
  ta,
  toe,
  uB,
  uU,
  uh,
  ul,
  ut,
  v0,
  vB,
  vQ,
  vg,
  vie,
  vx,
  wQ,
  wg,
  wie,
  x1,
  xB,
  xO,
  xQ,
  xR,
  xd,
  y0,
  y1,
  yB,
  yC,
  yQ,
  yR,
  yU,
  yie,
  ys,
  z1,
  zR,
  zU,
  zne,
  zoe,
};
