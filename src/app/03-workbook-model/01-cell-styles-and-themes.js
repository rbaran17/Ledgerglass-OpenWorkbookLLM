import { dx, Ig, b1, cx, ux, lB, sB, Mt } from "../02-financial-functions.js";
import { qj, Cy, TN, zp, yc } from "../01-core-schemas-and-hashing.js";

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
    e.color !== undefined && GS(e.color),
    e.style === "none"
      ? { style: "none" }
      : {
          style: e.style,
          color: e.color === undefined ? "#000000" : GS(e.color),
        }
  );
}
function oh(e) {
  _w(e) || vs("Cell style must be an object.");
  for (const t of Object.keys(e))
    ZJ.has(t) || vs(`Unsupported cell style setting: ${t}.`);
  const r = {};
  e.binding !== undefined && (r.binding = NN(e.binding));
  for (const t of ["bold", "italic", "underline", "strikethrough", "wrap"])
    e[t] !== undefined &&
      (typeof e[t] != "boolean" && vs(`Style ${t} must be boolean.`),
      (r[t] = e[t]));
  (e.fontFamily !== undefined &&
    (dx.includes(e.fontFamily) || vs("Unsupported font family."),
    (r.fontFamily = e.fontFamily)),
    e.fontSize !== undefined &&
      ((typeof e.fontSize != "number" ||
        !Number.isInteger(e.fontSize) ||
        e.fontSize < 8 ||
        e.fontSize > 72) &&
        vs("Font size must be an integer between 8 and 72."),
      (r.fontSize = e.fontSize)));
  for (const t of ["fill", "color"]) e[t] !== undefined && (r[t] = GS(e[t]));
  if (
    (e.align !== undefined &&
      (["left", "center", "right"].includes(e.align) ||
        vs("Invalid text alignment."),
      (r.align = e.align)),
    e.verticalAlign !== undefined &&
      (["top", "middle", "bottom"].includes(e.verticalAlign) ||
        vs("Choose Top, Center or Bottom vertical alignment."),
      (r.verticalAlign = e.verticalAlign)),
    e.indent !== undefined &&
      ((typeof e.indent != "number" ||
        !Number.isInteger(e.indent) ||
        e.indent < 0 ||
        e.indent > 15) &&
        vs("Indent must be an integer between 0 and 15."),
      (r.indent = e.indent)),
    e.border !== undefined &&
      (["all", "bottom", "none"].includes(e.border) ||
        vs("Unsupported cell border."),
      (r.border = e.border)),
    e.borders !== undefined)
  ) {
    (!_w(e.borders) || Object.keys(e.borders).some((n) => !s0.includes(n))) &&
      vs("Cell borders support top, right, bottom and left edges only.");
    const t = {};
    for (const n of s0) e.borders[n] !== undefined && (t[n] = cB(e.borders[n]));
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
    e[t] !== undefined &&
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
  if (((e = ys(e)), (t = e?.borders) != null && t[r])) return e.borders[r];
  if (e?.border === "none") return { style: "none" };
  if (e?.border === "all" || (e?.border === "bottom" && r === "bottom"))
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
    (e?.indent ?? 0) *
      GJ.indentEm *
      (e != null && e.fontSize ? (e.fontSize * 4) / 3 : 14)
  );
}
function Dw(e, r, t = {}) {
  e = ys(e);
  const n = e?.align ?? (typeof r == "number" ? "right" : "left"),
    o = fB(e),
    a = [e?.underline && "underline", e?.strikethrough && "line-through"]
      .filter(Boolean)
      .join(" ");
  return {
    background: e?.fill,
    color: e?.color,
    fontWeight: e != null && e.bold ? 700 : 400,
    fontStyle: e != null && e.italic ? "italic" : "normal",
    fontFamily: e?.fontFamily,
    fontSize: e != null && e.fontSize ? `${e.fontSize}pt` : undefined,
    textDecoration: a || "none",
    textAlign: n,
    verticalAlign: e?.verticalAlign,
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
      overflowWrap: e != null && e.wrap ? "anywhere" : undefined,
    }
  );
}
function hB(e) {
  if (((e = ys(e)), !(e != null && e.borders))) {
    const t = e?.border === "all" ? 2 : e?.border === "none" ? 0 : 1;
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
          .filter((r) => e[r] !== undefined)
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
  e.fontRole !== undefined &&
    (["body", "heading"].includes(e.fontRole) ||
      ho("Choose Body or Heading theme font."),
    (r.fontRole = e.fontRole));
  for (const n of ["fillRole", "colorRole"])
    e[n] !== undefined &&
      (_g.includes(e[n]) || ho("Choose a supported theme color role."),
      (r[n] = e[n]));
  if (e.borderRoles !== undefined) {
    (ah(e.borderRoles) || ho("Border color roles must be an object."),
      mg(e.borderRoles, s0, "border role"));
    const n = {};
    for (const o of s0)
      e.borderRoles[o] !== undefined &&
        (_g.includes(e.borderRoles[o]) ||
          ho("Choose a supported border color role."),
        (!md(r.base, o) || md(r.base, o)?.style === "none") &&
          ho("Add a border line before assigning its theme color."),
        (n[o] = e.borderRoles[o]));
    Object.keys(n).length && (r.borderRoles = n);
  }
  return r;
}
function IN(e) {
  (ah(e) || ho("Workbook styles must be an object."),
    mg(
      e,
      ["namespace", "theme", "styles", "missingTheme"],
      "workbook style catalog",
    ),
    e.missingTheme !== undefined &&
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
    o = e.theme === undefined || e.missingTheme === true,
    a = e.theme === undefined ? Ac(wE) : mB(e.theme);
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
  const t = e?.styles.find((n) => n.id === r.styleId);
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
  const t = e === undefined ? undefined : IN(e);
  t ? (r.styleCatalog = t) : delete r.styleCatalog;
  const n = ah(e) && e.theme === undefined;
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
  if ((r = e?.binding) != null && r.missing)
    return "Named style is missing. Default formatting and direct cell formatting are shown.";
  if ((t = e?.binding) != null && t.themeFallback)
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
function tQ(e, r) {
  const t = Rp(r, "Style namespace", 100),
    n = Ac(e),
    o = n.styleCatalog?.namespace;
  if (!o) return n;
  n.styleCatalog.namespace = t;
  const a = (d) => {
    d?.style?.binding?.namespace === o && (d.style.binding.namespace = t);
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
  if (e.permutation !== undefined) {
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
function w1(e = "Reviewer") {
  return kc(e, "Reviewer name", Jl.authorLength);
}
function fx(e = {}, r) {
  const t =
    e.timestamp === undefined
      ? new Date(Math.max(Date.now(), r ? Date.parse(r) : 0)).toISOString()
      : Lv(e.timestamp);
  return (
    r && t < r && ji("The review time cannot precede the existing discussion."),
    t
  );
}
function aQ(e, r) {
  const t = new Map(e.sheets.map((d) => [d.id, d]));
  let n = false;
  const o = e.reviewThreads?.map((d) => {
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
        ? ((n = true), { ...d, anchor: { ...d.anchor, sheetName: u.name } })
        : d;
    }),
    a = e.sheets.map((d) => {
      let u;
      for (const [c, h] of Object.entries(d.cells))
        h.hyperlink?.kind !== "internal" ||
          t.has(h.hyperlink.sheetId) ||
          (u ?? (u = { ...d.cells }),
          (u[c] = { ...h }),
          delete u[c].hyperlink,
          (n = true));
      return u ? { ...d, cells: u } : d;
    });
  return n ? { ...e, sheets: a, ...(o ? { reviewThreads: o } : {}) } : e;
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
          return [
            n,
            o.hyperlink?.kind === "internal"
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
function Pv(e) {
  return [
    "top-left",
    ...(e?.vertical !== undefined ? ["top-right"] : []),
    ...(e?.horizontal !== undefined
      ? ["bottom-left", ...(e?.vertical !== undefined ? ["bottom-right"] : [])]
      : []),
  ];
}
function aD(e, r, t) {
  return (
    e.split("-")[t === "row" ? 0 : 1] === r.split("-")[t === "row" ? 0 : 1]
  );
}
function AB(e, r) {
  if (e === undefined) return;
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
  if (t.horizontal === undefined && t.vertical === undefined)
    throw Error("Choose a horizontal split, vertical split, or both.");
  for (const n of [t.horizontal, t.vertical])
    if (n !== undefined && (!Number.isFinite(n) || n < 20 || n > 80))
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
    ...(t.horizontal !== undefined ? { horizontal: t.horizontal } : {}),
    ...(t.vertical !== undefined ? { vertical: t.vertical } : {}),
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
      e.defaultRowHeight === undefined
        ? undefined
        : EE(e.defaultRowHeight, "row"),
    n =
      e.defaultColumnWidth === undefined
        ? undefined
        : EE(e.defaultColumnWidth, "column");
  (t !== undefined && (r.defaultRowHeight = t),
    n !== undefined && (r.defaultColumnWidth = n));
}
function Dg(e) {
  return e.defaultRowHeight ?? lB;
}
function Og(e) {
  return e.defaultColumnWidth ?? sB;
}
function gg(e, r) {
  if (!Number.isInteger(r) || r < 0 || r >= e.rowCount)
    throw Error("Row index is outside the worksheet.");
  return e.rowHeights?.[r] ?? Dg(e);
}
function Ip(e, r) {
  if (!Number.isInteger(r) || r < 0 || r >= e.colCount)
    throw Error("Column index is outside the worksheet.");
  return e.columnWidths?.[r] ?? Og(e);
}
function gQ(e) {
  return Object.freeze({
    workbookId: e.id,
    revision: e.revision,
    snapshotGuard: JSON.stringify(e),
  });
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
  if ((bQ(r), e === undefined)) return [];
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
  if (e.outlineGroups === undefined) return;
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
function RB(e) {
  return e.outlineGroups?.map((t) => ({ ...t, id: crypto.randomUUID() }));
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
  if (e === undefined) return;
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
    (e.engagement !== undefined &&
      (o.engagement = uf(e.engagement, "Engagement", 200)),
    e.period !== undefined && (o.period = uf(e.period, "Period", 100)),
    e.template !== undefined)
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
      (a.sourceRevision !== undefined &&
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
      ...(a.sourceRevision !== undefined
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
  const t = { ...e };
  e5(r);
  const n = [r, e.metadata?.createdAt ?? r, e.metadata?.modifiedAt ?? r]
    .sort()
    .at(-1);
  return (px(e.metadata ? { ...e.metadata, modifiedAt: n } : mx({}, r), t), t);
}
const f6 = /[A-Za-z0-9_.\\\u0080-\uffff]/,
  _Q =
    /^(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*)(?::(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*))?!/,
  DQ = (e) =>
    [...e.toUpperCase()].reduce((r, t) => r * 26 + t.charCodeAt(0) - 64, 0) - 1;
function p6(e, r, t) {
  const n = Number(r) + (e ? 0 : t);
  return n < 1 || n > Mt.maxRows ? "#REF!" : e + n;
}
const MQ = (e) =>
  e ? { ...e, ...(e.style ? { style: { ...e.style } } : {}) } : { raw: null };
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
const sD = (e) => !!e && typeof e == "object" && !Array.isArray(e);
function al(e) {
  throw Error(e);
}
function jn(e) {
  var r;
  (r = e.protection) != null &&
    r.enabled &&
    al("This sheet is protected. Unprotect it before changing its contents.");
}
function vQ(e, r, t, n = {}) {
  if (n.readOnly !== undefined && typeof n.readOnly != "boolean")
    throw Error("Read-only state must be boolean.");
  if (n.readOnly) throw Error("This workbook is read-only.");
  const o = n.expectedSnapshot;
  if (
    o !== undefined &&
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
    if (l !== undefined && typeof l != "boolean")
      throw Error("Reset choices must be boolean.");
  const s = t.rowHeight === undefined ? undefined : EE(t.rowHeight, "row"),
    d = t.columnWidth === undefined ? undefined : EE(t.columnWidth, "column"),
    u = { ...a };
  return (
    s !== undefined &&
      s !== Dg(a) &&
      (s === lB ? delete u.defaultRowHeight : (u.defaultRowHeight = s)),
    d !== undefined &&
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
  return t || r.length > 4 ? undefined : r;
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
  return t != null && t.hasNumber ? t.maxDecimals : undefined;
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
    u.kind === "month" &&
      u.text.length <= 2 &&
      (o[l - 1]?.kind === "hour" ||
        (o[l - 1]?.kind === "elapsed" && o[l - 1].text.startsWith("h")) ||
        o[l + 1]?.kind === "second") &&
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
        l = /,+$/.exec(u)?.[0] ?? "";
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
function JQ(e) {
  return e.replace(/"[^"]*"|\\./g, "").replace(/[@ ]/g, "") === "";
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
  const o = n.length === 1 ? KQ(e) : undefined;
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

export {
  $B,
  AB,
  Ac,
  Ad,
  Ay,
  BB,
  BQ,
  Bw,
  DQ,
  Dg,
  Dw,
  E1,
  F1,
  FB,
  FQ,
  G3,
  GQ,
  H3,
  HQ,
  IN,
  Ip,
  J3,
  JJ,
  Jf,
  Jl,
  KS,
  Lv,
  MN,
  MQ,
  NB,
  Og,
  PB,
  PQ,
  Pv,
  QJ,
  QS,
  Ql,
  RB,
  RN,
  Rp,
  Si,
  TB,
  TQ,
  UB,
  UQ,
  Ub,
  Uv,
  XQ,
  Xm,
  YS,
  Z3,
  ZQ,
  ZS,
  _N,
  _Q,
  _g,
  a1,
  aD,
  aQ,
  al,
  bB,
  c6,
  cB,
  cQ,
  cn,
  f6,
  fB,
  fQ,
  fx,
  gB,
  gQ,
  gg,
  gx,
  hB,
  hQ,
  ho,
  jB,
  jN,
  ji,
  jn,
  jw,
  kc,
  m6,
  mB,
  mQ,
  md,
  mx,
  n5,
  oh,
  p6,
  px,
  r5,
  rQ,
  s0,
  sD,
  tQ,
  uB,
  uf,
  ul,
  vB,
  vQ,
  vs,
  vx,
  w1,
  x1,
  xB,
  xE,
  ys,
  yv,
  zB,
};
