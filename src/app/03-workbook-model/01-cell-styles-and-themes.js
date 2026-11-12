import {
  supportedFontFamilies,
  maxRowHeight,
  minRowHeight,
  maxColumnWidth,
  minColumnWidth,
  defaultRowHeight,
  defaultColumnWidth,
  workbookLimits,
} from "../02-financial-functions.js";
import {
  hasLinkedOrBoundCells,
  dateCalculationLimits,
  defaultDatePolicy,
  parseDateText,
  defineDataPropertyByKey,
} from "../01-core-schemas-and-hashing.js";

const cellBorderEdges = ["top", "right", "bottom", "left"],
  borderLineStyles = [
    "none",
    "thin",
    "medium",
    "thick",
    "dotted",
    "dashed",
    "double",
  ],
  cellIndentMetrics = { indentEm: 0.75 },
  supportedCellStyleKeys = new Set([
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
function failWithMessage(e) {
  throw Error(e);
}
function isPlainObjectLiteral(e) {
  return !!(
    e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e))
  );
}
function validateStyleColor(e) {
  return (
    (typeof e != "string" || !/^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(e)) &&
      failWithMessage("Style colors must use hexadecimal colors."),
    e
  );
}
function validateCellBorder(e) {
  return (
    (!isPlainObjectLiteral(e) ||
      Object.keys(e).some((r) => !["style", "color"].includes(r)) ||
      !borderLineStyles.includes(e.style)) &&
      failWithMessage(
        "Choose a supported border line style and hexadecimal color.",
      ),
    e.color !== undefined && validateStyleColor(e.color),
    e.style === "none"
      ? { style: "none" }
      : {
          style: e.style,
          color:
            e.color === undefined ? "#000000" : validateStyleColor(e.color),
        }
  );
}
function validateCellStyle(e) {
  isPlainObjectLiteral(e) || failWithMessage("Cell style must be an object.");
  for (const t of Object.keys(e))
    supportedCellStyleKeys.has(t) ||
      failWithMessage(`Unsupported cell style setting: ${t}.`);
  const r = {};
  e.binding !== undefined && (r.binding = validateNamedStyleBinding(e.binding));
  for (const t of ["bold", "italic", "underline", "strikethrough", "wrap"])
    e[t] !== undefined &&
      (typeof e[t] != "boolean" &&
        failWithMessage(`Style ${t} must be boolean.`),
      (r[t] = e[t]));
  (e.fontFamily !== undefined &&
    (supportedFontFamilies.includes(e.fontFamily) ||
      failWithMessage("Unsupported font family."),
    (r.fontFamily = e.fontFamily)),
    e.fontSize !== undefined &&
      ((typeof e.fontSize != "number" ||
        !Number.isInteger(e.fontSize) ||
        e.fontSize < 8 ||
        e.fontSize > 72) &&
        failWithMessage("Font size must be an integer between 8 and 72."),
      (r.fontSize = e.fontSize)));
  for (const t of ["fill", "color"])
    e[t] !== undefined && (r[t] = validateStyleColor(e[t]));
  if (
    (e.align !== undefined &&
      (["left", "center", "right"].includes(e.align) ||
        failWithMessage("Invalid text alignment."),
      (r.align = e.align)),
    e.verticalAlign !== undefined &&
      (["top", "middle", "bottom"].includes(e.verticalAlign) ||
        failWithMessage("Choose Top, Center or Bottom vertical alignment."),
      (r.verticalAlign = e.verticalAlign)),
    e.indent !== undefined &&
      ((typeof e.indent != "number" ||
        !Number.isInteger(e.indent) ||
        e.indent < 0 ||
        e.indent > 15) &&
        failWithMessage("Indent must be an integer between 0 and 15."),
      (r.indent = e.indent)),
    e.border !== undefined &&
      (["all", "bottom", "none"].includes(e.border) ||
        failWithMessage("Unsupported cell border."),
      (r.border = e.border)),
    e.borders !== undefined)
  ) {
    (!isPlainObjectLiteral(e.borders) ||
      Object.keys(e.borders).some((n) => !cellBorderEdges.includes(n))) &&
      failWithMessage(
        "Cell borders support top, right, bottom and left edges only.",
      );
    const t = {};
    for (const n of cellBorderEdges)
      e.borders[n] !== undefined && (t[n] = validateCellBorder(e.borders[n]));
    Object.keys(t).length && (r.borders = t);
  }
  return r;
}
function validateNamedStyleBinding(e) {
  (!isPlainObjectLiteral(e) ||
    Object.keys(e).some(
      (t) =>
        !["namespace", "styleId", "base", "missing", "themeFallback"].includes(
          t,
        ),
    )) &&
    failWithMessage(
      "Named style binding must contain a namespace, style ID and supported base.",
    );
  for (const t of ["namespace", "styleId"])
    (typeof e[t] != "string" ||
      !e[t].trim() ||
      e[t].length > 100 ||
      /[\x00-\x1f\x7f]/.test(e[t])) &&
      failWithMessage(
        "Named style identifiers must contain 1–100 characters without control characters.",
      );
  (!isPlainObjectLiteral(e.base) || Object.hasOwn(e.base, "binding")) &&
    failWithMessage("A named style base cannot contain another binding.");
  for (const t of ["missing", "themeFallback"])
    e[t] !== undefined &&
      e[t] !== true &&
      failWithMessage("Invalid named style fallback flag.");
  const r = validateCellStyle(e.base);
  return (
    e.missing &&
      Object.keys(r).length &&
      failWithMessage("A missing named style must use an empty fallback base."),
    {
      namespace: e.namespace,
      styleId: e.styleId,
      base: r,
      ...(e.missing ? { missing: true } : {}),
      ...(e.themeFallback ? { themeFallback: true } : {}),
    }
  );
}
function resolveEffectiveCellStyle(e) {
  if (!e) return {};
  const { binding, ...t } = e;
  if (!binding) return t;
  const n = binding.missing ? {} : binding.base,
    o = { ...n, ...t };
  return (
    Object.hasOwn(t, "border")
      ? delete o.borders
      : n.borders && (o.borders = { ...n.borders }),
    t.borders && (o.borders = { ...o.borders, ...t.borders }),
    o
  );
}
function resolveBorderEdge(e, r) {
  if (((e = resolveEffectiveCellStyle(e)), e?.borders?.[r]))
    return e.borders[r];
  if (e?.border === "none") return { style: "none" };
  if (e?.border === "all" || (e?.border === "bottom" && r === "bottom"))
    return { style: "thin", color: "#000000" };
}
function borderWidthPx(e) {
  return !e || e.style === "none"
    ? 0
    : e.style === "medium"
      ? 2
      : e.style === "thick" || e.style === "double"
        ? 3
        : 1;
}
function borderToCssShorthand(e) {
  if (e)
    return e.style === "none"
      ? "none"
      : `${borderWidthPx(e)}px ${["dotted", "dashed", "double"].includes(e.style) ? e.style : "solid"} ${e.color ?? "#000000"}`;
}
function indentToPx(e) {
  return (
    (e = resolveEffectiveCellStyle(e)),
    (e?.indent ?? 0) *
      cellIndentMetrics.indentEm *
      (e?.fontSize ? (e.fontSize * 4) / 3 : 14)
  );
}
function cellStyleToCss(e, r, t = {}) {
  e = resolveEffectiveCellStyle(e);
  const n = e?.align ?? (typeof r == "number" ? "right" : "left"),
    o = indentToPx(e),
    a = [e?.underline && "underline", e?.strikethrough && "line-through"]
      .filter(Boolean)
      .join(" ");
  return {
    background: e?.fill,
    color: e?.color,
    fontWeight: e?.bold ? 700 : 400,
    fontStyle: e?.italic ? "italic" : "normal",
    fontFamily: e?.fontFamily,
    fontSize: e?.fontSize ? `${e.fontSize}pt` : undefined,
    textDecoration: a || "none",
    textAlign: n,
    verticalAlign: e?.verticalAlign,
    paddingLeft: 3 + (n === "left" ? o : 0),
    paddingRight: 3 + (n === "right" ? o : 0),
    borderTop:
      borderToCssShorthand(resolveBorderEdge(e, "top")) ?? t.defaultBorder,
    borderRight:
      borderToCssShorthand(resolveBorderEdge(e, "right")) ?? t.defaultBorder,
    borderBottom:
      borderToCssShorthand(resolveBorderEdge(e, "bottom")) ?? t.defaultBorder,
    borderLeft:
      borderToCssShorthand(resolveBorderEdge(e, "left")) ?? t.defaultBorder,
  };
}
function textWrapToCss(e) {
  return (
    (e = resolveEffectiveCellStyle(e)),
    {
      whiteSpace: e?.wrap ? "pre-wrap" : "nowrap",
      overflowWrap: e?.wrap ? "anywhere" : undefined,
    }
  );
}
function measureCellBorderInsets(e) {
  if (((e = resolveEffectiveCellStyle(e)), !e?.borders)) {
    const t = e?.border === "all" ? 2 : e?.border === "none" ? 0 : 1;
    return { horizontal: t, vertical: t };
  }
  const r = (t) => {
    const n = resolveBorderEdge(e, t);
    return n ? borderWidthPx(n) : 0.5;
  };
  return {
    horizontal: r("left") + r("right"),
    vertical: r("top") + r("bottom"),
  };
}
const themePaletteRoles = [
    "background",
    "text",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
  ],
  workbookStyleLimits = { maxStyles: 50 },
  officeThemeDefinition = {
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
  builtInWorkbookThemes = [
    officeThemeDefinition,
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

function assertKnownStyleKeys(e, r, t) {
  Object.keys(e).some((n) => !r.includes(n)) &&
    failWithMessage(`Unsupported ${t} setting.`);
}
function validateBoundedStyleText(e, r, t = 80) {
  return (
    (typeof e != "string" ||
      !e.trim() ||
      e.length > t ||
      /[\x00-\x1f\x7f]/.test(e)) &&
      failWithMessage(
        `${r} must contain 1–${t} characters without control characters.`,
      ),
    e.trim()
  );
}
function cloneStyleValue(e) {
  return structuredClone(e);
}
function canonicalStyleKey(e) {
  return Array.isArray(e)
    ? `[${e.map(canonicalStyleKey).join(",")}]`
    : isPlainObjectLiteral(e)
      ? `{${Object.keys(e)
          .filter((r) => e[r] !== undefined)
          .sort()
          .map((r) => `${JSON.stringify(r)}:${canonicalStyleKey(e[r])}`)
          .join(",")}}`
      : JSON.stringify(e);
}
function validateWorkbookTheme(e) {
  (isPlainObjectLiteral(e) || failWithMessage("Choose a workbook theme."),
    assertKnownStyleKeys(e, ["name", "palette", "fonts"], "theme"),
    (!isPlainObjectLiteral(e.palette) || !isPlainObjectLiteral(e.fonts)) &&
      failWithMessage("A theme needs eight colors and body and heading fonts."),
    assertKnownStyleKeys(e.palette, themePaletteRoles, "theme palette"),
    assertKnownStyleKeys(e.fonts, ["body", "heading"], "theme font"));
  const r = {};
  for (const t of themePaletteRoles) {
    const n = e.palette[t];
    ((typeof n != "string" || !/^#[0-9a-f]{6}$/i.test(n)) &&
      failWithMessage("Theme colors must use six-digit hexadecimal colors."),
      (r[t] = n.toUpperCase()));
  }
  for (const t of ["body", "heading"])
    supportedFontFamilies.includes(e.fonts[t]) ||
      failWithMessage("Choose a supported theme font.");
  return {
    name: validateBoundedStyleText(e.name, "Theme name"),
    palette: r,
    fonts: { body: e.fonts.body, heading: e.fonts.heading },
  };
}
function validateNamedStyleDefinition(e) {
  (isPlainObjectLiteral(e) ||
    failWithMessage("Named style definition must be an object."),
    assertKnownStyleKeys(
      e,
      ["base", "fontRole", "fillRole", "colorRole", "borderRoles"],
      "named style",
    ),
    (!isPlainObjectLiteral(e.base) || Object.hasOwn(e.base, "binding")) &&
      failWithMessage(
        "A named style definition needs supported formatting without another binding.",
      ));
  const r = { base: validateCellStyle(e.base) };
  e.fontRole !== undefined &&
    (["body", "heading"].includes(e.fontRole) ||
      failWithMessage("Choose Body or Heading theme font."),
    (r.fontRole = e.fontRole));
  for (const n of ["fillRole", "colorRole"])
    e[n] !== undefined &&
      (themePaletteRoles.includes(e[n]) ||
        failWithMessage("Choose a supported theme color role."),
      (r[n] = e[n]));
  if (e.borderRoles !== undefined) {
    (isPlainObjectLiteral(e.borderRoles) ||
      failWithMessage("Border color roles must be an object."),
      assertKnownStyleKeys(e.borderRoles, cellBorderEdges, "border role"));
    const n = {};
    for (const o of cellBorderEdges)
      e.borderRoles[o] !== undefined &&
        (themePaletteRoles.includes(e.borderRoles[o]) ||
          failWithMessage("Choose a supported border color role."),
        (!resolveBorderEdge(r.base, o) ||
          resolveBorderEdge(r.base, o)?.style === "none") &&
          failWithMessage(
            "Add a border line before assigning its theme color.",
          ),
        (n[o] = e.borderRoles[o]));
    Object.keys(n).length && (r.borderRoles = n);
  }
  return r;
}
function validateWorkbookStyleCatalog(e) {
  (isPlainObjectLiteral(e) ||
    failWithMessage("Workbook styles must be an object."),
    assertKnownStyleKeys(
      e,
      ["namespace", "theme", "styles", "missingTheme"],
      "workbook style catalog",
    ),
    e.missingTheme !== undefined &&
      e.missingTheme !== true &&
      failWithMessage("Invalid missing theme flag."),
    (!Array.isArray(e.styles) ||
      e.styles.length > workbookStyleLimits.maxStyles) &&
      failWithMessage("A workbook supports at most 50 named styles."));
  const r = new Set(),
    t = new Set(),
    n = e.styles.map((s) => {
      (isPlainObjectLiteral(s) || failWithMessage("Invalid named style."),
        assertKnownStyleKeys(s, ["id", "name", "definition"], "named style"));
      const d = validateBoundedStyleText(s.id, "Style ID", 100),
        u = validateBoundedStyleText(s.name, "Style name");
      return (
        (r.has(d) || t.has(u.toLocaleLowerCase("en-US"))) &&
          failWithMessage(
            "Named style IDs and names must be unique within this workbook.",
          ),
        r.add(d),
        t.add(u.toLocaleLowerCase("en-US")),
        {
          id: d,
          name: u,
          definition: validateNamedStyleDefinition(s.definition),
        }
      );
    }),
    o = e.theme === undefined || e.missingTheme === true,
    a =
      e.theme === undefined
        ? cloneStyleValue(officeThemeDefinition)
        : validateWorkbookTheme(e.theme);
  return (
    o &&
      canonicalStyleKey(a) !== canonicalStyleKey(officeThemeDefinition) &&
      failWithMessage("A missing theme must use the Office fallback."),
    {
      namespace: validateBoundedStyleText(e.namespace, "Style namespace", 100),
      theme: a,
      styles: n,
      ...(o ? { missingTheme: true } : {}),
    }
  );
}
function createDefaultStyleCatalog(e) {
  return validateWorkbookStyleCatalog({
    namespace: e,
    theme: officeThemeDefinition,
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
function resolveNamedStyleAppearance(e, r) {
  const t = cloneStyleValue(e.base);
  (e.fontRole && (t.fontFamily = r.fonts[e.fontRole]),
    e.fillRole && (t.fill = r.palette[e.fillRole]),
    e.colorRole && (t.color = r.palette[e.colorRole]));
  for (const o of cellBorderEdges)
    e.borderRoles?.[o] &&
      (t.borders = {
        ...t.borders,
        [o]: { ...resolveBorderEdge(t, o), color: r.palette[e.borderRoles[o]] },
      });
  return validateCellStyle(t);
}
function resolveNamedStyleBinding(e, r) {
  e &&
    r.namespace !== e.namespace &&
    failWithMessage(
      "This style belongs to another workbook. Import or remap its style catalog before pasting the binding.",
    );
  const t = e?.styles.find((n) => n.id === r.styleId);
  return {
    namespace: r.namespace,
    styleId: r.styleId,
    base: t && e ? resolveNamedStyleAppearance(t.definition, e.theme) : {},
    ...(t ? {} : { missing: true }),
    ...(e?.missingTheme ? { themeFallback: true } : {}),
  };
}
function reconcileWorkbookStyleBindings(e, r) {
  const t = e === undefined ? undefined : validateWorkbookStyleCatalog(e);
  t ? (r.styleCatalog = t) : delete r.styleCatalog;
  const n = isPlainObjectLiteral(e) && e.theme === undefined;
  for (const a of r.sheets)
    for (const s of hasLinkedOrBoundCells(a.cells)
      ? Object.values(a.cells)
      : []) {
      if (!s.style?.binding) continue;
      const d = validateNamedStyleBinding(s.style.binding),
        u = resolveNamedStyleBinding(t, d);
      (!u.missing &&
        !n &&
        canonicalStyleKey(d) !== canonicalStyleKey(u) &&
        failWithMessage(
          "Named style appearance does not match the workbook catalog. Reapply the style from a trusted catalog.",
        ),
        (s.style.binding = u));
    }
}
function assertNamedStyleBindingCurrent(e, r) {
  if (!r?.binding) return;
  const t = validateNamedStyleBinding(r.binding),
    n = resolveNamedStyleBinding(e.styleCatalog, t);
  canonicalStyleKey(t) !== canonicalStyleKey(n) &&
    failWithMessage(
      "Named style binding is stale or inconsistent with this workbook. Reapply its named style.",
    );
}
function describeNamedStyleFallback(e) {
  if (e?.binding?.missing)
    return "Named style is missing. Default formatting and direct cell formatting are shown.";
  if (e?.binding?.themeFallback)
    return "Workbook theme is missing. The Office theme is shown.";
}
function collectStyleFallbackWarnings(e) {
  const r = new Set();
  e.styleCatalog?.missingTheme &&
    r.add("Workbook theme is missing. The Office theme is shown.");
  for (const n of e.sheets)
    for (const o of Object.values(n.cells)) {
      const a = describeNamedStyleFallback(o.style);
      a && r.add(a);
    }
  return [...r];
}
function hasBorderEdgeChanges(e, r) {
  return cellBorderEdges.some(
    (t) =>
      canonicalStyleKey(resolveBorderEdge(e, t)) !==
      canonicalStyleKey(resolveBorderEdge(r, t)),
  );
}
function renameStyleNamespace(e, r) {
  const t = validateBoundedStyleText(r, "Style namespace", 100),
    n = cloneStyleValue(e),
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
function validateRowColumnMove(e, r) {
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
function mapIndexAfterMove(e, r) {
  if (r.permutation)
    return e >= r.index && e < r.index + r.count
      ? r.permutation[e - r.index]
      : e;
  const { index, count, moveBefore } = r,
    a = moveBefore < index ? moveBefore : moveBefore - count;
  return e >= index && e < index + count
    ? a + e - index
    : moveBefore < index && e >= moveBefore && e < index
      ? e + count
      : moveBefore > index && e >= index + count && e < moveBefore
        ? e - count
        : e;
}
function planMoveSegments(e, r, t) {
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
      a(s, s, mapIndexAfterMove(s, t) - s);
    return (a(Math.max(e, t.index + t.count), r, 0), o);
  }
  const n = [
    ...new Set(
      [e, r + 1, t.index, t.index + t.count, t.moveBefore].filter(
        (o) => o >= e && o <= r + 1,
      ),
    ),
  ].sort((o, a) => o - a);
  return n.slice(0, -1).map((o, a) => ({
    start: o,
    end: n[a + 1] - 1,
    offset: mapIndexAfterMove(o, t) - o,
  }));
}
function mapRangeAfterMove(e, r, t) {
  if (t.permutation && e <= t.index && r >= t.index + t.count - 1)
    return [{ start: e, end: r }];
  const n = planMoveSegments(e, r, t)
      .map((a) => ({ start: a.start + a.offset, end: a.end + a.offset }))
      .sort((a, s) => a.start - s.start),
    o = [];
  for (const a of n) {
    const s = o.at(-1);
    s && a.start <= s.end + 1 ? (s.end = Math.max(s.end, a.end)) : o.push(a);
  }
  return o;
}
function resolveMovedRange(e, r, t, n = "range", o = false) {
  const a = mapRangeAfterMove(e, r, t);
  if (a.length !== 1)
    throw new Error(
      `This move would split ${n} into separate ranges. Move the complete range or use explicit cell references first.`,
    );
  if (o && new Set(planMoveSegments(e, r, t).map((s) => s.offset)).size !== 1)
    throw new Error(
      `This move would reorder part of ${n}. Move the complete structure together.`,
    );
  return a[0];
}
const reviewThreadLimits = {
    threads: 1e3,
    messagesPerThread: 100,
    totalMessages: 5e3,
    textLength: 4e3,
    authorLength: 100,
    urlLength: 2048,
    labelLength: 256,
  },
  unsafeControlCharPattern =
    /[\u0000-\u001f\u007f-\u009f\u200b-\u200f\u202a-\u202e\u2060-\u206f\ufeff]/,
  throwWithMessage = (e) => {
    throw new Error(e);
  };
function assertPlainReviewRecord(e, r) {
  return (
    (!e ||
      typeof e != "object" ||
      Array.isArray(e) ||
      ![Object.prototype, null].includes(Object.getPrototypeOf(e))) &&
      throwWithMessage(`${r} must be a plain object.`),
    e
  );
}
function normalizeBoundedText(e, r, t, n = false) {
  return (
    (typeof e != "string" ||
      !e.trim() ||
      e.length > t ||
      unsafeControlCharPattern.test(n ? e.replace(/[\n\r\t]/g, "") : e)) &&
      throwWithMessage(
        `${r} must contain 1–${t} characters without unsupported control characters.`,
      ),
    e.trim()
  );
}
function normalizeReviewTimestamp(e, r = "Review timestamp") {
  return (
    (typeof e != "string" ||
      !Number.isFinite(Date.parse(e)) ||
      new Date(e).toISOString() !== e) &&
      throwWithMessage(`${r} must be an ISO UTC timestamp.`),
    e
  );
}
function normalizeReviewAuthor(e = "Reviewer") {
  return normalizeBoundedText(
    e,
    "Reviewer name",
    reviewThreadLimits.authorLength,
  );
}
function resolveReviewTimestamp(e = {}, r) {
  const t =
    e.timestamp === undefined
      ? new Date(Math.max(Date.now(), r ? Date.parse(r) : 0)).toISOString()
      : normalizeReviewTimestamp(e.timestamp);
  return (
    r &&
      t < r &&
      throwWithMessage(
        "The review time cannot precede the existing discussion.",
      ),
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
        const l = resolveReviewTimestamp({ timestamp: r }, d.updatedAt);
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
          ((u ??= { ...d.cells }),
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
        Object.entries(t.cells).map(([n, o]) => [
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
        ]),
      ),
    })),
  };
}
function createReviewThread(e, r) {
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
  return e.formulaWatches?.length
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
function createFormulaWatch(e, r) {
  return e.formulaWatches?.length
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
function anchorsMatchOnAxis(e, r, t) {
  return (
    e.split("-")[t === "row" ? 0 : 1] === r.split("-")[t === "row" ? 0 : 1]
  );
}
function validateSplitPaneState(e, r) {
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
function validateDefaultSizePixels(e, r) {
  const t = r === "row" ? minRowHeight : minColumnWidth,
    n = r === "row" ? maxRowHeight : maxColumnWidth;
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
function validateSheetDefaultSizes(e, r) {
  const t =
      e.defaultRowHeight === undefined
        ? undefined
        : validateDefaultSizePixels(e.defaultRowHeight, "row"),
    n =
      e.defaultColumnWidth === undefined
        ? undefined
        : validateDefaultSizePixels(e.defaultColumnWidth, "column");
  (t !== undefined && (r.defaultRowHeight = t),
    n !== undefined && (r.defaultColumnWidth = n));
}
function defaultRowHeightOf(e) {
  return e.defaultRowHeight ?? defaultRowHeight;
}
function defaultColumnWidthOf(e) {
  return e.defaultColumnWidth ?? defaultColumnWidth;
}
function rowHeightAt(e, r) {
  if (!Number.isInteger(r) || r < 0 || r >= e.rowCount)
    throw Error("Row index is outside the worksheet.");
  return e.rowHeights?.[r] ?? defaultRowHeightOf(e);
}
function columnWidthAt(e, r) {
  if (!Number.isInteger(r) || r < 0 || r >= e.colCount)
    throw Error("Column index is outside the worksheet.");
  return e.columnWidths?.[r] ?? defaultColumnWidthOf(e);
}
function buildSnapshotGuard(e) {
  return Object.freeze({
    workbookId: e.id,
    revision: e.revision,
    snapshotGuard: JSON.stringify(e),
  });
}
const worksheetCapacityLimits = {
    maxIntervalsPerAxis: 1e4,
    maxFrozenVisibleRows: 100,
    maxFrozenVisibleCols: 50,
    maxXlsxHiddenRows: 1e5,
  },
  axisCountOf = (e, r) => (r === "row" ? e.rowCount : e.colCount),
  hiddenKeyForAxis = (e) => (e === "row" ? "hiddenRows" : "hiddenCols");
function assertWorksheetAxis(e) {
  if (e !== "row" && e !== "column") throw Error("Choose rows or columns.");
}
function assertAxisCapacity(e) {
  if (!Number.isSafeInteger(e) || e < 0)
    throw Error("Axis capacity must be a nonnegative integer.");
}
function normalizeHiddenIntervals(e, r) {
  if ((assertAxisCapacity(r), e === undefined)) return [];
  if (
    !Array.isArray(e) ||
    e.length > worksheetCapacityLimits.maxIntervalsPerAxis
  )
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
function createAxisVisibilityIndex(e, r = []) {
  const t = normalizeHiddenIntervals(r, e),
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
function readHiddenAxisIntervals(e, r) {
  return (
    assertWorksheetAxis(r),
    normalizeHiddenIntervals(e[hiddenKeyForAxis(r)], axisCountOf(e, r))
  );
}
const outlineGroupLimits = { maxGroups: 1e3, maxDepth: 8 };

function assertOutlineGroupNesting(e, r) {
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
    if ((n.push(o), n.length > outlineGroupLimits.maxDepth))
      throw Error("Outline groups support at most 8 nested groups.");
    return { ...o, depth: n.length };
  });
}
function validateOutlineGroups(
  e,
  r,
  t = { rows: r.rowCount, cols: r.colCount },
) {
  if (e.outlineGroups === undefined) return;
  if (
    !Array.isArray(e.outlineGroups) ||
    e.outlineGroups.length > outlineGroupLimits.maxGroups
  )
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
  (assertOutlineGroupNesting(o, "row"),
    assertOutlineGroupNesting(o, "column"),
    (r.outlineGroups = o));
}
function addOutlineGroup(e, r) {
  return (
    assertWorksheetAxis(r),
    assertOutlineGroupNesting(e.outlineGroups ?? [], r)
  );
}
function cn(e, r) {
  assertWorksheetAxis(r);
  const t = axisCountOf(e, r),
    n = [
      ...normalizeHiddenIntervals(r === "row" ? e.hiddenRows : e.hiddenCols, t),
      ...(e.outlineGroups ?? [])
        .filter((a) => a.axis === r && a.collapsed)
        .map(({ start, end }) => ({ start, end })),
    ].sort((a, s) => a.start - s.start || a.end - s.end),
    o = [];
  for (const a of n) {
    const s = o.at(-1);
    s && a.start <= s.end + 1
      ? (s.end = Math.max(s.end, a.end))
      : o.push({ ...a });
  }
  return normalizeHiddenIntervals(o, t);
}
function regenerateOutlineGroupIds(e) {
  return e.outlineGroups?.map((t) => ({ ...t, id: crypto.randomUUID() }));
}
function remapOutlineGroupsForEdit(e, r, t, n, o) {
  assertWorksheetAxis(r);
  const a = axisCountOf(e, r);
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
    (assertOutlineGroupNesting(d, "row"),
      assertOutlineGroupNesting(d, "column"));
  } catch {
    throw Error(
      "This edit would make outline groups overlap or duplicate. Ungroup the affected rows or columns first.",
    );
  }
  return d;
}
const isPlainObjectValue = (e) =>
  !!e &&
  typeof e == "object" &&
  !Array.isArray(e) &&
  [Object.prototype, null].includes(Object.getPrototypeOf(e));
function trimRequiredMetadataText(e, r, t) {
  if (typeof e != "string" || !e.trim() || e.length > t)
    throw new Error(`${r} must contain 1–${t} characters.`);
  return e.trim();
}
function isValidUtcTimestamp(e) {
  if (
    typeof e != "string" ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(e) ||
    !Number.isFinite(Date.parse(e)) ||
    new Date(e).toISOString() !== e
  )
    throw new Error("Workbook metadata requires a valid UTC timestamp.");
  return e;
}
function validateWorkbookMetadata(e, r) {
  if (e === undefined) return;
  if (
    !isPlainObjectValue(e) ||
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
  const t = isValidUtcTimestamp(e.createdAt),
    n = isValidUtcTimestamp(e.modifiedAt);
  if (n < t)
    throw new Error("Workbook modified time cannot precede its creation time.");
  const o = { createdAt: t, modifiedAt: n };
  if (
    (e.engagement !== undefined &&
      (o.engagement = trimRequiredMetadataText(
        e.engagement,
        "Engagement",
        200,
      )),
    e.period !== undefined &&
      (o.period = trimRequiredMetadataText(e.period, "Period", 100)),
    e.template !== undefined)
  ) {
    const a = e.template;
    if (
      !isPlainObjectValue(a) ||
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
      id: trimRequiredMetadataText(a.id, "Template ID", 100),
      name: trimRequiredMetadataText(a.name, "Template name", 200),
      version: Number(a.version),
      sourceWorkbookId: trimRequiredMetadataText(
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
function createWorkbookMetadata(e = {}, r = new Date().toISOString()) {
  const t = {};
  return (
    validateWorkbookMetadata({ createdAt: r, modifiedAt: r, ...e }, t),
    t.metadata
  );
}
function touchWorkbookModifiedAt(e, r = new Date().toISOString()) {
  const t = { ...e };
  isValidUtcTimestamp(r);
  const n = [r, e.metadata?.createdAt ?? r, e.metadata?.modifiedAt ?? r]
    .sort()
    .at(-1);
  return (
    validateWorkbookMetadata(
      e.metadata
        ? { ...e.metadata, modifiedAt: n }
        : createWorkbookMetadata({}, r),
      t,
    ),
    t
  );
}
const unquotedNameCharPattern = /[A-Za-z0-9_.\\\u0080-\uffff]/,
  sheetReferencePrefixPattern =
    /^(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*)(?::(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*))?!/,
  columnLettersToIndex = (e) =>
    [...e.toUpperCase()].reduce((r, t) => r * 26 + t.charCodeAt(0) - 64, 0) - 1;
function offsetRowReference(e, r, t) {
  const n = Number(r) + (e ? 0 : t);
  return n < 1 || n > workbookLimits.maxRows ? "#REF!" : e + n;
}
const cloneCellRecord = (e) =>
  e ? { ...e, ...(e.style ? { style: { ...e.style } } : {}) } : { raw: null };
const validationLimits = Object.freeze({
    rulesPerSheet: 100,
    listItems: 200,
    checks: 1e5,
    customChecks: 5e3,
    displayedIssues: 500,
  }),
  isNonArrayObject = (e) => !!e && typeof e == "object" && !Array.isArray(e),
  assertBoundedValidationText = (e, r, t) => {
    if (typeof e != "string" || e.length > r)
      throw Error(`Invalid validation ${t}.`);
    return e;
  };
function validateDataValidationRule(e, r, t = defaultDatePolicy) {
  if (r === null || r === "") return e.allowBlank;
  if (e.kind === "custom") return false;
  if (e.kind === "list")
    return e.list.some((o) => o.toLowerCase() === String(r).toLowerCase());
  let n =
    e.kind === "textLength" ? String(r).length : typeof r == "number" ? r : NaN;
  if (
    (["date", "time"].includes(e.kind) &&
      typeof r == "string" &&
      (n = parseDateText(r, t) ?? NaN),
    !Number.isFinite(n) ||
      (e.kind === "integer" && !Number.isInteger(n)) ||
      (e.kind === "date" &&
        (n < 1 ||
          n > dateCalculationLimits.maxSerial + 0.999999999 ||
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
class DataValidationError extends Error {
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
    defineDataPropertyByKey(this, "issues");
    ((this.name = "DataValidationError"),
      (this.issues = structuredClone(n)),
      MB.set(this, JSON.stringify({ book: t, issues: n })));
  }
}
function assertNoStopValidationIssues(e) {
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
function resolveValidationInspection(e, r) {
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
  for (const { address } of t)
    n[address] = { ...e.cells[address], raw: r(e.id, address) };
  return n;
}
const conditionalRuleKinds = [
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
];
function assertBoundedString(e, r, t) {
  if (typeof e != "string" || e.length > r) throw Error(`Invalid ${t}.`);
  return e;
}
function assertSixDigitHexColor(e) {
  if (typeof e != "string" || !/^#[\da-f]{6}$/i.test(e))
    throw Error("Choose a six-digit hexadecimal color.");
  return e;
}
const isCellInsideRange = (e, r) =>
    r.row >= e.firstRow &&
    r.row <= e.lastRow &&
    r.col >= e.firstCol &&
    r.col <= e.lastCol,
  isErrorValueText = (e) =>
    typeof e == "string" &&
    /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|ERROR!|CYCLE!|SPILL!|CALC!|LIMIT!|BLOCKED!)/.test(
      e,
    );
function sortConditionalRulesByPriority(e) {
  return e
    .map((r, t) => ({ rule: r, rank: r.priority ?? 101 + e.length - t }))
    .sort((r, t) => r.rank - t.rank)
    .map((r) => r.rule);
}

function assertSheetUnprotected(e) {
  e.protection?.enabled &&
    failWithMessage(
      "This sheet is protected. Unprotect it before changing its contents.",
    );
}
function applySheetSizeEdits(e, r, t, n = {}) {
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
    (assertSheetUnprotected(a),
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
  const s =
      t.rowHeight === undefined
        ? undefined
        : validateDefaultSizePixels(t.rowHeight, "row"),
    d =
      t.columnWidth === undefined
        ? undefined
        : validateDefaultSizePixels(t.columnWidth, "column"),
    u = { ...a };
  return (
    s !== undefined &&
      s !== defaultRowHeightOf(a) &&
      (s === defaultRowHeight
        ? delete u.defaultRowHeight
        : (u.defaultRowHeight = s)),
    d !== undefined &&
      d !== defaultColumnWidthOf(a) &&
      (d === defaultColumnWidth
        ? delete u.defaultColumnWidth
        : (u.defaultColumnWidth = d)),
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
async function hashSheetPassword(e, r) {
  const t = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(r + "\0" + e),
  );
  return [...new Uint8Array(t)]
    .map((n) => n.toString(16).padStart(2, "0"))
    .join("");
}
async function updateSheetProtection(e, r, t, n) {
  const o = e.sheets.find((l) => l.id === r);
  (n &&
    o.protection?.enabled &&
    failWithMessage(
      "Unprotect this sheet before changing its protection settings.",
    ),
    t.length > 200 && failWithMessage("Password is too long."),
    !n &&
      o.protection?.hash &&
      (await hashSheetPassword(t, o.protection.salt)) !== o.protection.hash &&
      failWithMessage("The password does not match."));
  const a = crypto.randomUUID(),
    s = n
      ? {
          enabled: true,
          ...(t ? { salt: a, hash: await hashSheetPassword(t, a) } : {}),
        }
      : { enabled: false };
  return {
    ...e,
    sheets: e.sheets.map((l) => (l.id === r ? { ...l, protection: s } : l)),
  };
}
const supportedNumberLocales = [
    "en-US",
    "en-GB",
    "de-DE",
    "fr-FR",
    "ja-JP",
    "en-IN",
  ],
  supportedCurrencyCodes = [
    "USD",
    "EUR",
    "GBP",
    "JPY",
    "CAD",
    "AUD",
    "INR",
    "CHF",
  ],
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
  builtInDateTimeFormats = {
    date: ["yyyy-mm-dd", "dd/mm/yyyy", "m/d/yyyy", "d mmm yyyy", "d mmmm yyyy"],
    time: ["hh:mm", "hh:mm:ss", "h:mm AM/PM", "hh:mm:ss.000"],
    datetime: [
      "yyyy-mm-dd hh:mm:ss",
      "dd/mm/yyyy hh:mm",
      "m/d/yyyy h:mm AM/PM",
    ],
    duration: ["[h]:mm", "[h]:mm:ss", "[h]:mm:ss.000", "[m]:ss", "[s].000"],
  };
function throwWorkbookError(e) {
  throw new Error(e);
}
const numericFormatKinds = new Set([
    "number",
    "currency",
    "accounting",
    "percent",
  ]),
  dateTimeFormatKinds = new Set(["date", "time", "datetime", "duration"]);
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
function parseNumberFormatCode(e) {
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
function summarizeNumberFormat(e) {
  if (!e) return;
  const r = BN(e);
  if (!r) return;
  const t = parseNumberFormatCode(r[0]);
  return t?.hasNumber ? t.maxDecimals : undefined;
}
function numberFormatToIntlOptions(e, r) {
  const t = BN(r);
  if (!t) return;
  const n = e < 0 && t.length > 1,
    o = e === 0 && t.length > 2 ? 2 : n ? 1 : 0,
    a = parseNumberFormatCode(t[o]);
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
function parseDateTimeFormatCode(e) {
  const r = e.replace(/"[^"]*"|\\./g, "");
  if (!/[ymdhs]/i.test(r)) return;
  const t = [];
  let n = 0;
  for (let u = 0; u < e.length;) {
    const l = e.slice(u),
      c = l[0];
    if (c === '"') {
      const v = e.indexOf('"', u + 1);
      (v < 0 &&
        throwWorkbookError("The format contains an unclosed quoted literal."),
        t.push({ kind: "literal", text: e.slice(u + 1, v) }),
        (u = v + 1));
      continue;
    }
    if (c === "\\") {
      (l.length < 2 &&
        throwWorkbookError("A format escape must be followed by a character."),
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
        throwWorkbookError(
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
        throwWorkbookError(
          "Fractional seconds must follow seconds and use 1–3 zero digits.",
        ),
        (n = m[1].length),
        t.push({ kind: "fraction", text: m[1] }),
        (u += m[0].length));
      continue;
    }
    if (/^[ /:,.-]$/.test(c)) {
      (t.push({ kind: "literal", text: c }), u++);
      continue;
    }
    throwWorkbookError(
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
      throwWorkbookError(
        "Elapsed-time patterns start with one [h], [m] or [s] field and cannot include calendar or AM/PM fields.",
      );
    const u = d[0].text[0];
    ((u === "s" && s) || (u === "m" && o.some((l) => l.kind === "minute"))) &&
      throwWorkbookError("Use [h]:mm:ss, [m]:ss or [s] for elapsed time.");
  }
  return (
    o.some((u) => u.kind === "ampm") &&
      !o.some((u) => u.kind === "hour") &&
      throwWorkbookError("AM/PM requires an hour field."),
    !a &&
      !s &&
      !d.length &&
      throwWorkbookError("The format does not contain a date or time field."),
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
function parseCustomNumberFormat(e) {
  let r = "",
    t = 1,
    n = false;
  for (let s = 0; s < e.length;) {
    const d = e.slice(s);
    if (d[0] === '"') {
      const u = e.indexOf('"', s + 1);
      (u < 0 &&
        throwWorkbookError("The format contains an unclosed quoted literal."),
        (r += e.slice(s, u + 1)),
        (s = u + 1));
    } else if (d[0] === "\\")
      (d.length < 2 &&
        throwWorkbookError("A format escape must be followed by a character."),
        (r += d.slice(0, 2)),
        (s += 2));
    else if (/^[0#]/.test(d)) {
      n &&
        throwWorkbookError(
          "Only one numeric field per custom section is supported.",
        );
      const u = /^[0#.,]+/.exec(d)[0],
        l = /,+$/.exec(u)?.[0] ?? "";
      (l.length > 2 &&
        throwWorkbookError(
          "Custom display scaling supports at most two trailing commas.",
        ),
        (t = 1e3 ** l.length),
        (r += u.slice(0, u.length - l.length)),
        (s += u.length),
        (n = true));
    } else ((r += d[0]), s++);
  }
  /\[/.test(r.replace(/"[^"]*"|\\./g, "")) &&
    throwWorkbookError(
      "Custom colors, conditions and locale tags are not supported. Use conditional formatting or the locale selector.",
    );
  const o = parseNumberFormatCode(r);
  return (
    o ||
      throwWorkbookError(
        "Unsupported custom format. Use 0/# digits, grouping, decimals, percent, quoted literals and positive/negative/zero sections.",
      ),
    o.maxDecimals > 12 &&
      throwWorkbookError("Custom formats support at most 12 decimal places."),
    { parsed: o, divisor: t }
  );
}
function isTextOnlyFormatSection(e) {
  return e.replace(/"[^"]*"|\\./g, "").replace(/[@ ]/g, "") === "";
}
function validateNumberFormatCode(e) {
  const r = $b.get(e);
  if (r) return r;
  (typeof e != "string" ||
    !e.length ||
    e.length > 200 ||
    /[\u0000-\u001f\u007f]/.test(e)) &&
    throwWorkbookError("Enter a supported format pattern of 1–200 characters.");
  let t;
  const n = BN(e);
  n ||
    throwWorkbookError(
      "A format has at most four sections and must close quoted literals.",
    );
  const o = n.length === 1 ? parseDateTimeFormatCode(e) : undefined;
  return (
    o
      ? (t = o)
      : (n.length === 4 &&
          !isTextOnlyFormatSection(n[3]) &&
          throwWorkbookError(
            "The text section supports @ and quoted or escaped literals only.",
          ),
        (t = {
          kind: "numeric",
          sections: n.slice(0, 3).map(parseCustomNumberFormat),
          ...(n.length === 4 ? { textSection: n[3] } : {}),
        })),
    $b.size >= 256 && $b.delete($b.keys().next().value),
    $b.set(e, t),
    t
  );
}

export {
  summarizeNumberFormat,
  validateSplitPaneState,
  cloneStyleValue,
  normalizeHiddenIntervals,
  builtInDateTimeFormats,
  conditionalRuleKinds,
  resolveValidationInspection,
  validateNumberFormatCode,
  columnLettersToIndex,
  defaultRowHeightOf,
  cellStyleToCss,
  readHiddenAxisIntervals,
  FB,
  isNonArrayObject,
  hiddenKeyForAxis,
  dateTimeFormatKinds,
  resolveMovedRange,
  updateSheetProtection,
  validateWorkbookStyleCatalog,
  columnWidthAt,
  sortConditionalRulesByPriority,
  reconcileWorkbookStyleBindings,
  mapIndexAfterMove,
  reviewThreadLimits,
  describeNamedStyleFallback,
  normalizeReviewTimestamp,
  axisCountOf,
  cloneCellRecord,
  addOutlineGroup,
  defaultColumnWidthOf,
  supportedNumberLocales,
  Pv,
  hasBorderEdgeChanges,
  isPlainObjectValue,
  canonicalStyleKey,
  regenerateOutlineGroupIds,
  textWrapToCss,
  validateBoundedStyleText,
  throwWorkbookError,
  validateOutlineGroups,
  remapOutlineGroupsForEdit,
  supportedCurrencyCodes,
  isErrorValueText,
  assertBoundedString,
  Uv,
  numericFormatKinds,
  assertPlainReviewRecord,
  unsafeControlCharPattern,
  assertWorksheetAxis,
  ZQ,
  validateNamedStyleDefinition,
  assertNamedStyleBindingCurrent,
  sheetReferencePrefixPattern,
  themePaletteRoles,
  validationLimits,
  anchorsMatchOnAxis,
  aQ,
  collectStyleFallbackWarnings,
  builtInWorkbookThemes,
  validateCellBorder,
  createReviewThread,
  cn,
  unquotedNameCharPattern,
  indentToPx,
  fQ,
  resolveReviewTimestamp,
  createDefaultStyleCatalog,
  buildSnapshotGuard,
  rowHeightAt,
  touchWorkbookModifiedAt,
  measureCellBorderInsets,
  createFormulaWatch,
  assertNoStopValidationIssues,
  DataValidationError,
  throwWithMessage,
  assertSheetUnprotected,
  validateDataValidationRule,
  normalizeBoundedText,
  assertSixDigitHexColor,
  validateWorkbookTheme,
  validateSheetDefaultSizes,
  resolveBorderEdge,
  createWorkbookMetadata,
  isCellInsideRange,
  validateCellStyle,
  offsetRowReference,
  validateWorkbookMetadata,
  r5,
  validateRowColumnMove,
  cellBorderEdges,
  renameStyleNamespace,
  borderLineStyles,
  trimRequiredMetadataText,
  createAxisVisibilityIndex,
  resolveNamedStyleAppearance,
  applySheetSizeEdits,
  failWithMessage,
  vx,
  normalizeReviewAuthor,
  worksheetCapacityLimits,
  mapRangeAfterMove,
  resolveNamedStyleBinding,
  resolveEffectiveCellStyle,
  assertBoundedValidationText,
  numberFormatToIntlOptions,
};
export { isPlainObjectLiteral };

/* Also used by the workers. */
export {
  $b,
  BN,
  assertAxisCapacity,
  isTextOnlyFormatSection,
  parseCustomNumberFormat,
  parseDateTimeFormatCode,
  parseNumberFormatCode,
};

/* Also used by the workers. */
export {
  assertKnownStyleKeys,
  assertOutlineGroupNesting,
  isValidUtcTimestamp,
  officeThemeDefinition,
  outlineGroupLimits,
  supportedCellStyleKeys,
  validateNamedStyleBinding,
  validateStyleColor,
};

/* Also used by the workers. */
export { validateDefaultSizePixels };
