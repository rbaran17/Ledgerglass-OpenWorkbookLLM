import {
  supportedFontFamilies,
  maxRowHeight,
  minRowHeight,
  maxColumnWidth,
  minColumnWidth,
  defaultRowHeight,
  defaultColumnWidth,
  workbookLimits,
} from "../02-text-and-financial-functions.js";
import {
  hasLinkedOrBoundCells,
  dateCalculationLimits,
  defaultDatePolicy,
  parseDateText,
  defineDataPropertyByKey,
} from "../01-schemas-date-functions-and-validation.js";

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
function validateStyleColor(color) {
  return (
    (typeof color != "string" ||
      !/^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(color)) &&
      failWithMessage("Style colors must use hexadecimal colors."),
    color
  );
}
function validateCellBorder(line) {
  return (
    (!isPlainObjectLiteral(line) ||
      Object.keys(line).some((r) => !["style", "color"].includes(r)) ||
      !borderLineStyles.includes(line.style)) &&
      failWithMessage(
        "Choose a supported border line style and hexadecimal color.",
      ),
    line.color !== undefined && validateStyleColor(line.color),
    line.style === "none"
      ? { style: "none" }
      : {
          style: line.style,
          color:
            line.color === undefined
              ? "#000000"
              : validateStyleColor(line.color),
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
      Object.keys(e.borders).some(
        (border) => !cellBorderEdges.includes(border),
      )) &&
      failWithMessage(
        "Cell borders support top, right, bottom and left edges only.",
      );
    const t = {};
    for (const cellBorderEdge of cellBorderEdges)
      e.borders[cellBorderEdge] !== undefined &&
        (t[cellBorderEdge] = validateCellBorder(e.borders[cellBorderEdge]));
    Object.keys(t).length && (r.borders = t);
  }
  return r;
}
function validateNamedStyleBinding(binding) {
  (!isPlainObjectLiteral(binding) ||
    Object.keys(binding).some(
      (t) =>
        !["namespace", "styleId", "base", "missing", "themeFallback"].includes(
          t,
        ),
    )) &&
    failWithMessage(
      "Named style binding must contain a namespace, style ID and supported base.",
    );
  for (const t of ["namespace", "styleId"])
    (typeof binding[t] != "string" ||
      !binding[t].trim() ||
      binding[t].length > 100 ||
      /[\x00-\x1f\x7f]/.test(binding[t])) &&
      failWithMessage(
        "Named style identifiers must contain 1–100 characters without control characters.",
      );
  (!isPlainObjectLiteral(binding.base) ||
    Object.hasOwn(binding.base, "binding")) &&
    failWithMessage("A named style base cannot contain another binding.");
  for (const t of ["missing", "themeFallback"])
    binding[t] !== undefined &&
      binding[t] !== true &&
      failWithMessage("Invalid named style fallback flag.");
  const base = validateCellStyle(binding.base);
  return (
    binding.missing &&
      Object.keys(base).length &&
      failWithMessage("A missing named style must use an empty fallback base."),
    {
      namespace: binding.namespace,
      styleId: binding.styleId,
      base: base,
      ...(binding.missing ? { missing: true } : {}),
      ...(binding.themeFallback ? { themeFallback: true } : {}),
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
function resolveBorderEdge(e, cellBorderEdge) {
  if (((e = resolveEffectiveCellStyle(e)), e?.borders?.[cellBorderEdge]))
    return e.borders[cellBorderEdge];
  if (e?.border === "none") return { style: "none" };
  if (
    e?.border === "all" ||
    (e?.border === "bottom" && cellBorderEdge === "bottom")
  )
    return { style: "thin", color: "#000000" };
}
function borderWidthPx(borderEdge) {
  return !borderEdge || borderEdge.style === "none"
    ? 0
    : borderEdge.style === "medium"
      ? 2
      : borderEdge.style === "thick" || borderEdge.style === "double"
        ? 3
        : 1;
}
function borderToCssShorthand(borderEdge) {
  if (borderEdge)
    return borderEdge.style === "none"
      ? "none"
      : `${borderWidthPx(borderEdge)}px ${["dotted", "dashed", "double"].includes(borderEdge.style) ? borderEdge.style : "solid"} ${borderEdge.color ?? "#000000"}`;
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
  const textAlign = e?.align ?? (typeof r == "number" ? "right" : "left"),
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
    textAlign: textAlign,
    verticalAlign: e?.verticalAlign,
    paddingLeft: 3 + (textAlign === "left" ? o : 0),
    paddingRight: 3 + (textAlign === "right" ? o : 0),
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
  const r = (cellBorderEdge) => {
    const borderEdge = resolveBorderEdge(e, cellBorderEdge);
    return borderEdge ? borderWidthPx(borderEdge) : 0.5;
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
function validateWorkbookTheme(theme) {
  (isPlainObjectLiteral(theme) || failWithMessage("Choose a workbook theme."),
    assertKnownStyleKeys(theme, ["name", "palette", "fonts"], "theme"),
    (!isPlainObjectLiteral(theme.palette) ||
      !isPlainObjectLiteral(theme.fonts)) &&
      failWithMessage("A theme needs eight colors and body and heading fonts."),
    assertKnownStyleKeys(theme.palette, themePaletteRoles, "theme palette"),
    assertKnownStyleKeys(theme.fonts, ["body", "heading"], "theme font"));
  const palette = {};
  for (const themePaletteRole of themePaletteRoles) {
    const n = theme.palette[themePaletteRole];
    ((typeof n != "string" || !/^#[0-9a-f]{6}$/i.test(n)) &&
      failWithMessage("Theme colors must use six-digit hexadecimal colors."),
      (palette[themePaletteRole] = n.toUpperCase()));
  }
  for (const t of ["body", "heading"])
    supportedFontFamilies.includes(theme.fonts[t]) ||
      failWithMessage("Choose a supported theme font.");
  return {
    name: validateBoundedStyleText(theme.name, "Theme name"),
    palette: palette,
    fonts: { body: theme.fonts.body, heading: theme.fonts.heading },
  };
}
function validateNamedStyleDefinition(definition) {
  (isPlainObjectLiteral(definition) ||
    failWithMessage("Named style definition must be an object."),
    assertKnownStyleKeys(
      definition,
      ["base", "fontRole", "fillRole", "colorRole", "borderRoles"],
      "named style",
    ),
    (!isPlainObjectLiteral(definition.base) ||
      Object.hasOwn(definition.base, "binding")) &&
      failWithMessage(
        "A named style definition needs supported formatting without another binding.",
      ));
  const r = { base: validateCellStyle(definition.base) };
  definition.fontRole !== undefined &&
    (["body", "heading"].includes(definition.fontRole) ||
      failWithMessage("Choose Body or Heading theme font."),
    (r.fontRole = definition.fontRole));
  for (const n of ["fillRole", "colorRole"])
    definition[n] !== undefined &&
      (themePaletteRoles.includes(definition[n]) ||
        failWithMessage("Choose a supported theme color role."),
      (r[n] = definition[n]));
  if (definition.borderRoles !== undefined) {
    (isPlainObjectLiteral(definition.borderRoles) ||
      failWithMessage("Border color roles must be an object."),
      assertKnownStyleKeys(
        definition.borderRoles,
        cellBorderEdges,
        "border role",
      ));
    const n = {};
    for (const cellBorderEdge of cellBorderEdges)
      definition.borderRoles[cellBorderEdge] !== undefined &&
        (themePaletteRoles.includes(definition.borderRoles[cellBorderEdge]) ||
          failWithMessage("Choose a supported border color role."),
        (!resolveBorderEdge(r.base, cellBorderEdge) ||
          resolveBorderEdge(r.base, cellBorderEdge)?.style === "none") &&
          failWithMessage(
            "Add a border line before assigning its theme color.",
          ),
        (n[cellBorderEdge] = definition.borderRoles[cellBorderEdge]));
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
    styles = e.styles.map((style) => {
      (isPlainObjectLiteral(style) || failWithMessage("Invalid named style."),
        assertKnownStyleKeys(
          style,
          ["id", "name", "definition"],
          "named style",
        ));
      const d = validateBoundedStyleText(style.id, "Style ID", 100),
        u = validateBoundedStyleText(style.name, "Style name");
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
          definition: validateNamedStyleDefinition(style.definition),
        }
      );
    }),
    o = e.theme === undefined || e.missingTheme === true,
    theme =
      e.theme === undefined
        ? cloneStyleValue(officeThemeDefinition)
        : validateWorkbookTheme(e.theme);
  return (
    o &&
      canonicalStyleKey(theme) !== canonicalStyleKey(officeThemeDefinition) &&
      failWithMessage("A missing theme must use the Office fallback."),
    {
      namespace: validateBoundedStyleText(e.namespace, "Style namespace", 100),
      theme: theme,
      styles: styles,
      ...(o ? { missingTheme: true } : {}),
    }
  );
}
function createDefaultStyleCatalog(namespace) {
  return validateWorkbookStyleCatalog({
    namespace: namespace,
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
function resolveNamedStyleAppearance(definition, theme) {
  const t = cloneStyleValue(definition.base);
  (definition.fontRole && (t.fontFamily = theme.fonts[definition.fontRole]),
    definition.fillRole && (t.fill = theme.palette[definition.fillRole]),
    definition.colorRole && (t.color = theme.palette[definition.colorRole]));
  for (const cellBorderEdge of cellBorderEdges)
    definition.borderRoles?.[cellBorderEdge] &&
      (t.borders = {
        ...t.borders,
        [cellBorderEdge]: {
          ...resolveBorderEdge(t, cellBorderEdge),
          color: theme.palette[definition.borderRoles[cellBorderEdge]],
        },
      });
  return validateCellStyle(t);
}
function resolveNamedStyleBinding(e, r) {
  e &&
    r.namespace !== e.namespace &&
    failWithMessage(
      "This style belongs to another workbook. Import or remap its style catalog before pasting the binding.",
    );
  const t = e?.styles.find((style) => style.id === r.styleId);
  return {
    namespace: r.namespace,
    styleId: r.styleId,
    base: t && e ? resolveNamedStyleAppearance(t.definition, e.theme) : {},
    ...(t ? {} : { missing: true }),
    ...(e?.missingTheme ? { themeFallback: true } : {}),
  };
}
function reconcileWorkbookStyleBindings(styleCatalog, workbook) {
  const t =
    styleCatalog === undefined
      ? undefined
      : validateWorkbookStyleCatalog(styleCatalog);
  t ? (workbook.styleCatalog = t) : delete workbook.styleCatalog;
  const n =
    isPlainObjectLiteral(styleCatalog) && styleCatalog.theme === undefined;
  for (const sheet of workbook.sheets)
    for (const s of hasLinkedOrBoundCells(sheet.cells)
      ? Object.values(sheet.cells)
      : []) {
      if (!s.style?.binding) continue;
      const d = validateNamedStyleBinding(s.style.binding),
        namedStyleBinding = resolveNamedStyleBinding(t, d);
      (!namedStyleBinding.missing &&
        !n &&
        canonicalStyleKey(d) !== canonicalStyleKey(namedStyleBinding) &&
        failWithMessage(
          "Named style appearance does not match the workbook catalog. Reapply the style from a trusted catalog.",
        ),
        (s.style.binding = namedStyleBinding));
    }
}
function assertNamedStyleBindingCurrent(workbook, r) {
  if (!r?.binding) return;
  const t = validateNamedStyleBinding(r.binding),
    namedStyleBinding = resolveNamedStyleBinding(workbook.styleCatalog, t);
  canonicalStyleKey(t) !== canonicalStyleKey(namedStyleBinding) &&
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
function collectStyleFallbackWarnings(workbook) {
  const r = new Set();
  workbook.styleCatalog?.missingTheme &&
    r.add("Workbook theme is missing. The Office theme is shown.");
  for (const sheet of workbook.sheets)
    for (const cell of Object.values(sheet.cells)) {
      const a = describeNamedStyleFallback(cell.style);
      a && r.add(a);
    }
  return [...r];
}
function hasBorderEdgeChanges(e, r) {
  return cellBorderEdges.some(
    (cellBorderEdge) =>
      canonicalStyleKey(resolveBorderEdge(e, cellBorderEdge)) !==
      canonicalStyleKey(resolveBorderEdge(r, cellBorderEdge)),
  );
}
function renameStyleNamespace(e, r) {
  const t = validateBoundedStyleText(r, "Style namespace", 100),
    workbook = cloneStyleValue(e),
    o = workbook.styleCatalog?.namespace;
  if (!o) return workbook;
  workbook.styleCatalog.namespace = t;
  const a = (d) => {
    d?.style?.binding?.namespace === o && (d.style.binding.namespace = t);
  };
  for (const sheet of workbook.sheets)
    for (const cell of Object.values(sheet.cells)) a(cell);
  for (const d of workbook.events)
    for (const change of d.changes) (a(change.before), a(change.after));
  return workbook;
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
function planMoveSegments(start, end, t) {
  if (t.permutation) {
    const o = [],
      a = (start, end, offset) => {
        if (start > end) return;
        const l = o.at(-1);
        l && l.end + 1 === start && l.offset === offset
          ? (l.end = end)
          : o.push({ start: start, end: end, offset: offset });
      };
    a(start, Math.min(end, t.index - 1), 0);
    for (
      let s = Math.max(start, t.index);
      s <= Math.min(end, t.index + t.count - 1);
      s++
    )
      a(s, s, mapIndexAfterMove(s, t) - s);
    return (a(Math.max(start, t.index + t.count), end, 0), o);
  }
  const n = [
    ...new Set(
      [start, end + 1, t.index, t.index + t.count, t.moveBefore].filter(
        (o) => o >= start && o <= end + 1,
      ),
    ),
  ].sort((o, a) => o - a);
  return n.slice(0, -1).map((start, a) => ({
    start: start,
    end: n[a + 1] - 1,
    offset: mapIndexAfterMove(start, t) - start,
  }));
}
function mapRangeAfterMove(start, end, t) {
  if (t.permutation && start <= t.index && end >= t.index + t.count - 1)
    return [{ start: start, end: end }];
  const n = planMoveSegments(start, end, t)
      .map((a) => ({ start: a.start + a.offset, end: a.end + a.offset }))
      .sort((a, s) => a.start - s.start),
    o = [];
  for (const a of n) {
    const s = o.at(-1);
    s && a.start <= s.end + 1 ? (s.end = Math.max(s.end, a.end)) : o.push(a);
  }
  return o;
}
function resolveMovedRange(start, end, t, n = "range", o = false) {
  const a = mapRangeAfterMove(start, end, t);
  if (a.length !== 1)
    throw new Error(
      `This move would split ${n} into separate ranges. Move the complete range or use explicit cell references first.`,
    );
  if (
    o &&
    new Set(planMoveSegments(start, end, t).map((s) => s.offset)).size !== 1
  )
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
function resolveReviewTimestamp(e = {}, updatedAt) {
  const t =
    e.timestamp === undefined
      ? new Date(
          Math.max(Date.now(), updatedAt ? Date.parse(updatedAt) : 0),
        ).toISOString()
      : normalizeReviewTimestamp(e.timestamp);
  return (
    updatedAt &&
      t < updatedAt &&
      throwWithMessage(
        "The review time cannot precede the existing discussion.",
      ),
    t
  );
}
function aQ(workbook, timestamp) {
  const t = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet]));
  let n = false;
  const reviewThreads = workbook.reviewThreads?.map((reviewThread) => {
      if (reviewThread.orphaned) return reviewThread;
      const u = t.get(reviewThread.anchor.sheetId);
      if (!u) {
        n = true;
        const reviewTimestamp = resolveReviewTimestamp(
          { timestamp: timestamp },
          reviewThread.updatedAt,
        );
        return {
          ...reviewThread,
          updatedAt: reviewTimestamp,
          orphaned: { reason: "worksheet-deleted", createdAt: reviewTimestamp },
        };
      }
      return u.name !== reviewThread.anchor.sheetName
        ? ((n = true),
          {
            ...reviewThread,
            anchor: { ...reviewThread.anchor, sheetName: u.name },
          })
        : reviewThread;
    }),
    sheets = workbook.sheets.map((sheet) => {
      let cells;
      for (const [c, h] of Object.entries(sheet.cells))
        h.hyperlink?.kind !== "internal" ||
          t.has(h.hyperlink.sheetId) ||
          ((cells ??= { ...sheet.cells }),
          (cells[c] = { ...h }),
          delete cells[c].hyperlink,
          (n = true));
      return cells ? { ...sheet, cells: cells } : sheet;
    });
  return n
    ? {
        ...workbook,
        sheets: sheets,
        ...(reviewThreads ? { reviewThreads: reviewThreads } : {}),
      }
    : workbook;
}
function uQ(workbook, r) {
  return {
    ...workbook,
    ...(workbook.reviewThreads
      ? {
          reviewThreads: workbook.reviewThreads.map((reviewThread) => ({
            ...reviewThread,
            anchor: {
              ...reviewThread.anchor,
              sheetId:
                r.get(reviewThread.anchor.sheetId) ??
                reviewThread.anchor.sheetId,
            },
          })),
        }
      : {}),
    sheets: workbook.sheets.map((sheet) => ({
      ...sheet,
      cells: Object.fromEntries(
        Object.entries(sheet.cells).map(([n, o]) => [
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
function createReviewThread(workbook, r) {
  const t = uQ(workbook, r);
  return t.reviewThreads
    ? {
        ...t,
        reviewThreads: t.reviewThreads.map((reviewThread) => ({
          ...reviewThread,
          id: crypto.randomUUID(),
          messages: reviewThread.messages.map((message) => ({
            ...message,
            id: crypto.randomUUID(),
          })),
        })),
      }
    : t;
}
function fQ(e, sheetId) {
  return e.formulaWatches?.length
    ? {
        ...e,
        formulaWatches: e.formulaWatches.map((formulaWatch) =>
          !formulaWatch.orphaned && formulaWatch.anchor.sheetId === sheetId
            ? { ...formulaWatch, orphaned: "deleted-sheet" }
            : formulaWatch,
        ),
      }
    : e;
}
function createFormulaWatch(e, r) {
  return e.formulaWatches?.length
    ? {
        ...e,
        formulaWatches: e.formulaWatches.map((formulaWatch) => ({
          ...formulaWatch,
          id: crypto.randomUUID(),
          anchor: {
            ...formulaWatch.anchor,
            sheetId:
              r.get(formulaWatch.anchor.sheetId) ?? formulaWatch.anchor.sheetId,
          },
        })),
      }
    : e;
}
function Pv(sheetSplitPanes) {
  return [
    "top-left",
    ...(sheetSplitPanes?.vertical !== undefined ? ["top-right"] : []),
    ...(sheetSplitPanes?.horizontal !== undefined
      ? [
          "bottom-left",
          ...(sheetSplitPanes?.vertical !== undefined ? ["bottom-right"] : []),
        ]
      : []),
  ];
}
function anchorsMatchOnAxis(e, r, t) {
  return (
    e.split("-")[t === "row" ? 0 : 1] === r.split("-")[t === "row" ? 0 : 1]
  );
}
function validateSplitPaneState(splitPanes, r) {
  if (splitPanes === undefined) return;
  if (
    !splitPanes ||
    typeof splitPanes != "object" ||
    Array.isArray(splitPanes) ||
    (Object.getPrototypeOf(splitPanes) !== Object.prototype &&
      Object.getPrototypeOf(splitPanes) !== null) ||
    Object.keys(splitPanes).some(
      (splitPane) =>
        !["horizontal", "vertical", "row", "col"].includes(splitPane),
    )
  )
    throw Error(
      "Split panes require divider percentages and valid row/column anchors.",
    );
  const t = splitPanes;
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
function rowHeightAt(sheet, r) {
  if (!Number.isInteger(r) || r < 0 || r >= sheet.rowCount)
    throw Error("Row index is outside the worksheet.");
  return sheet.rowHeights?.[r] ?? defaultRowHeightOf(sheet);
}
function columnWidthAt(sheet, r) {
  if (!Number.isInteger(r) || r < 0 || r >= sheet.colCount)
    throw Error("Column index is outside the worksheet.");
  return sheet.columnWidths?.[r] ?? defaultColumnWidthOf(sheet);
}
function buildSnapshotGuard(workbook) {
  return Object.freeze({
    workbookId: workbook.id,
    revision: workbook.revision,
    snapshotGuard: JSON.stringify(workbook),
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
function createAxisVisibilityIndex(logicalCount, r = []) {
  const hiddenIntervals = normalizeHiddenIntervals(r, logicalCount),
    n = [0];
  for (const hiddenInterval of hiddenIntervals)
    n.push(n.at(-1) + hiddenInterval.end - hiddenInterval.start + 1);
  const count = logicalCount - n.at(-1),
    a = (l) => {
      let c = 0,
        h = hiddenIntervals.length;
      for (; c < h;) {
        const f = (c + h) >>> 1;
        hiddenIntervals[f].start <= l ? (c = f + 1) : (h = f);
      }
      return c - 1;
    },
    countBefore = (l) => {
      const c = Math.max(0, Math.min(logicalCount, Math.floor(l))),
        h = a(c - 1),
        f =
          h < 0
            ? 0
            : n[h] +
              Math.min(
                c - hiddenIntervals[h].start,
                hiddenIntervals[h].end - hiddenIntervals[h].start + 1,
              );
      return c - f;
    },
    isHidden = (l) => {
      const c = a(l);
      return c >= 0 && l <= hiddenIntervals[c].end;
    },
    logicalIndex = (l) => {
      if (!Number.isInteger(l) || l < 0 || l >= count) return -1;
      let c = 0,
        h = logicalCount - 1;
      for (; c < h;) {
        const f = Math.floor((c + h) / 2);
        countBefore(f + 1) <= l ? (c = f + 1) : (h = f);
      }
      return c;
    };
  return {
    count: count,
    logicalCount: logicalCount,
    logicalIndex: logicalIndex,
    countBefore: countBefore,
    isHidden: isHidden,
    viewIndex: (l) =>
      Number.isInteger(l) && l >= 0 && l < logicalCount && !isHidden(l)
        ? countBefore(l)
        : null,
    next: (l, c, h = false) => {
      const f = Math.trunc(l) + (h ? 0 : c);
      if (!count || (c === 1 && f >= logicalCount) || (c === -1 && f < 0))
        return null;
      const p =
          c === 1
            ? countBefore(Math.max(0, f))
            : countBefore(Math.min(logicalCount, f + 1)) - 1,
        m = logicalIndex(p);
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
  sheet,
  r,
  t = { rows: r.rowCount, cols: r.colCount },
) {
  if (sheet.outlineGroups === undefined) return;
  if (
    !Array.isArray(sheet.outlineGroups) ||
    sheet.outlineGroups.length > outlineGroupLimits.maxGroups
  )
    throw Error("A worksheet supports at most 1,000 outline groups.");
  const n = new Set(),
    o = sheet.outlineGroups.map((outlineGroup) => {
      if (
        !outlineGroup ||
        typeof outlineGroup != "object" ||
        Array.isArray(outlineGroup) ||
        (Object.getPrototypeOf(outlineGroup) !== Object.prototype &&
          Object.getPrototypeOf(outlineGroup) !== null) ||
        Object.keys(outlineGroup).some(
          (s) => !["id", "axis", "start", "end", "collapsed"].includes(s),
        ) ||
        typeof outlineGroup.id != "string" ||
        !outlineGroup.id.trim() ||
        outlineGroup.id.length > 100 ||
        n.has(outlineGroup.id) ||
        !["row", "column"].includes(outlineGroup.axis) ||
        !Number.isSafeInteger(outlineGroup.start) ||
        !Number.isSafeInteger(outlineGroup.end) ||
        outlineGroup.start < 0 ||
        outlineGroup.end < outlineGroup.start ||
        outlineGroup.end >= (outlineGroup.axis === "row" ? t.rows : t.cols) ||
        typeof outlineGroup.collapsed != "boolean"
      )
        throw Error(
          "Outline groups require unique IDs, valid row or column bounds, and a boolean collapsed state.",
        );
      return (
        n.add(outlineGroup.id),
        {
          id: outlineGroup.id,
          axis: outlineGroup.axis,
          start: outlineGroup.start,
          end: outlineGroup.end,
          collapsed: outlineGroup.collapsed,
        }
      );
    });
  (assertOutlineGroupNesting(o, "row"),
    assertOutlineGroupNesting(o, "column"),
    (r.outlineGroups = o));
}
function addOutlineGroup(sheet, r) {
  return (
    assertWorksheetAxis(r),
    assertOutlineGroupNesting(sheet.outlineGroups ?? [], r)
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
function regenerateOutlineGroupIds(sheet) {
  return sheet.outlineGroups?.map((outlineGroup) => ({
    ...outlineGroup,
    id: crypto.randomUUID(),
  }));
}
function remapOutlineGroupsForEdit(e, axis, start, count, remove) {
  assertWorksheetAxis(axis);
  const a = axisCountOf(e, axis);
  if (
    !Number.isInteger(start) ||
    start < 0 ||
    !Number.isInteger(count) ||
    count < 1 ||
    count > 1e4 ||
    start + count > a ||
    typeof remove != "boolean"
  )
    throw Error("Choose a valid structural edit of 1–10,000 rows or columns.");
  if (!e.outlineGroups) return;
  const s = start + count - 1,
    d = e.outlineGroups.flatMap((outlineGroup) => {
      if (outlineGroup.axis !== axis || outlineGroup.end < start)
        return [{ ...outlineGroup }];
      if (!remove) {
        if (outlineGroup.end + count >= a)
          throw Error(
            "Insertion would discard outline groups at the worksheet boundary. Ungroup those rows or columns first.",
          );
        return [
          {
            ...outlineGroup,
            start:
              outlineGroup.start >= start
                ? outlineGroup.start + count
                : outlineGroup.start,
            end: outlineGroup.end + count,
          },
        ];
      }
      return outlineGroup.start > s
        ? [
            {
              ...outlineGroup,
              start: outlineGroup.start - count,
              end: outlineGroup.end - count,
            },
          ]
        : outlineGroup.start >= start && outlineGroup.end <= s
          ? []
          : [
              {
                ...outlineGroup,
                start: Math.min(outlineGroup.start, start),
                end:
                  outlineGroup.end > s ? outlineGroup.end - count : start - 1,
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
function validateWorkbookMetadata(metadata, r) {
  if (metadata === undefined) return;
  if (
    !isPlainObjectValue(metadata) ||
    Object.keys(metadata).some(
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
  const createdAt = isValidUtcTimestamp(metadata.createdAt),
    modifiedAt = isValidUtcTimestamp(metadata.modifiedAt);
  if (modifiedAt < createdAt)
    throw new Error("Workbook modified time cannot precede its creation time.");
  const o = { createdAt: createdAt, modifiedAt: modifiedAt };
  if (
    (metadata.engagement !== undefined &&
      (o.engagement = trimRequiredMetadataText(
        metadata.engagement,
        "Engagement",
        200,
      )),
    metadata.period !== undefined &&
      (o.period = trimRequiredMetadataText(metadata.period, "Period", 100)),
    metadata.template !== undefined)
  ) {
    const template = metadata.template;
    if (
      !isPlainObjectValue(template) ||
      Object.keys(template).some(
        (s) =>
          ![
            "id",
            "name",
            "version",
            "sourceWorkbookId",
            "sourceRevision",
          ].includes(s),
      ) ||
      !Number.isSafeInteger(template.version) ||
      Number(template.version) < 1 ||
      (template.sourceRevision !== undefined &&
        (!Number.isSafeInteger(template.sourceRevision) ||
          Number(template.sourceRevision) < 0))
    )
      throw new Error("Invalid template provenance.");
    o.template = {
      id: trimRequiredMetadataText(template.id, "Template ID", 100),
      name: trimRequiredMetadataText(template.name, "Template name", 200),
      version: Number(template.version),
      sourceWorkbookId: trimRequiredMetadataText(
        template.sourceWorkbookId,
        "Template source workbook ID",
        100,
      ),
      ...(template.sourceRevision !== undefined
        ? { sourceRevision: Number(template.sourceRevision) }
        : {}),
    };
  }
  r.metadata = o;
}
function createWorkbookMetadata(e = {}, timestamp = new Date().toISOString()) {
  const t = {};
  return (
    validateWorkbookMetadata(
      { createdAt: timestamp, modifiedAt: timestamp, ...e },
      t,
    ),
    t.metadata
  );
}
function touchWorkbookModifiedAt(e, timestamp = new Date().toISOString()) {
  const t = { ...e };
  isValidUtcTimestamp(timestamp);
  const modifiedAt = [
    timestamp,
    e.metadata?.createdAt ?? timestamp,
    e.metadata?.modifiedAt ?? timestamp,
  ]
    .sort()
    .at(-1);
  return (
    validateWorkbookMetadata(
      e.metadata
        ? { ...e.metadata, modifiedAt: modifiedAt }
        : createWorkbookMetadata({}, timestamp),
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
function validateDataValidationRule(e, r, datePolicy = defaultDatePolicy) {
  if (r === null || r === "") return e.allowBlank;
  if (e.kind === "custom") return false;
  if (e.kind === "list")
    return e.list.some((o) => o.toLowerCase() === String(r).toLowerCase());
  let n =
    e.kind === "textLength" ? String(r).length : typeof r == "number" ? r : NaN;
  if (
    (["date", "time"].includes(e.kind) &&
      typeof r == "string" &&
      (n = parseDateText(r, datePolicy) ?? NaN),
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
  constructor(book, issues) {
    super(
      issues
        .slice(0, 5)
        .map((o) => `${o.sheetName}!${o.address}: ${o.message} ${o.reason}`)
        .join(`
`) +
        (issues.length > 5
          ? `
${issues.length} validation violations in this batch.`
          : ""),
    );
    defineDataPropertyByKey(this, "issues");
    ((this.name = "DataValidationError"),
      (this.issues = structuredClone(issues)),
      MB.set(this, JSON.stringify({ book: book, issues: issues })));
  }
}
function assertNoStopValidationIssues(error) {
  const r = MB.get(error);
  if (
    !r ||
    !error.issues.length ||
    error.issues.some(
      (issue) => issue.style === "stop" || issue.evaluationError,
    ) ||
    JSON.stringify(JSON.parse(r).issues) !== JSON.stringify(error.issues)
  )
    throw Error(
      "Stop errors cannot be overridden. Correct the invalid values.",
    );
  const t = Object.freeze({ kind: "validation-warning-approval" });
  return (r5.set(t, r), t);
}
const FB = new WeakMap();
function resolveValidationInspection(workbook, report) {
  const t = FB.get(report);
  if (
    !t ||
    t.source !== JSON.stringify(workbook) ||
    t.content !== JSON.stringify(report)
  )
    throw Error(
      "Validation inspection is stale or unknown. Inspect existing values again.",
    );
  return JSON.stringify(
    {
      schemaVersion: 1,
      kind: "local-data-validation-inspection",
      report: report,
    },
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
    .map((rule, t) => ({
      rule: rule,
      rank: rule.priority ?? 101 + e.length - t,
    }))
    .sort((r, t) => r.rank - t.rank)
    .map((r) => r.rule);
}

function assertSheetUnprotected(e) {
  e.protection?.enabled &&
    failWithMessage(
      "This sheet is protected. Unprotect it before changing its contents.",
    );
}
function applySheetSizeEdits(workbook, sheetId, t, n = {}) {
  if (n.readOnly !== undefined && typeof n.readOnly != "boolean")
    throw Error("Read-only state must be boolean.");
  if (n.readOnly) throw Error("This workbook is read-only.");
  const expectedSnapshot = n.expectedSnapshot;
  if (
    expectedSnapshot !== undefined &&
    (!expectedSnapshot ||
      expectedSnapshot.workbookId !== workbook.id ||
      expectedSnapshot.revision !== workbook.revision ||
      expectedSnapshot.snapshotGuard !== JSON.stringify(workbook))
  )
    throw Error(
      "The workbook changed after these sizes were opened. Close the dialog and reopen it before applying sizes.",
    );
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Worksheet not found.");
  if (
    (assertSheetUnprotected(sheet),
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
    u = { ...sheet };
  return (
    s !== undefined &&
      s !== defaultRowHeightOf(sheet) &&
      (s === defaultRowHeight
        ? delete u.defaultRowHeight
        : (u.defaultRowHeight = s)),
    d !== undefined &&
      d !== defaultColumnWidthOf(sheet) &&
      (d === defaultColumnWidth
        ? delete u.defaultColumnWidth
        : (u.defaultColumnWidth = d)),
    t.resetRowOverrides &&
      Object.keys(sheet.rowHeights ?? {}).length &&
      delete u.rowHeights,
    t.resetColumnOverrides &&
      Object.keys(sheet.columnWidths ?? {}).length &&
      delete u.columnWidths,
    u.defaultRowHeight === sheet.defaultRowHeight &&
    u.defaultColumnWidth === sheet.defaultColumnWidth &&
    u.rowHeights === sheet.rowHeights &&
    u.columnWidths === sheet.columnWidths
      ? workbook
      : {
          ...workbook,
          sheets: workbook.sheets.map((sheet) =>
            sheet.id === sheetId ? u : sheet,
          ),
        }
  );
}
async function hashSheetPassword(e, salt) {
  const t = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(salt + "\0" + e),
  );
  return [...new Uint8Array(t)]
    .map((n) => n.toString(16).padStart(2, "0"))
    .join("");
}
async function updateSheetProtection(workbook, sheetId, t, n) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  (n &&
    sheet.protection?.enabled &&
    failWithMessage(
      "Unprotect this sheet before changing its protection settings.",
    ),
    t.length > 200 && failWithMessage("Password is too long."),
    !n &&
      sheet.protection?.hash &&
      (await hashSheetPassword(t, sheet.protection.salt)) !==
        sheet.protection.hash &&
      failWithMessage("The password does not match."));
  const salt = crypto.randomUUID(),
    protection = n
      ? {
          enabled: true,
          ...(t ? { salt: salt, hash: await hashSheetPassword(t, salt) } : {}),
        }
      : { enabled: false };
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet.id === sheetId ? { ...sheet, protection: protection } : sheet,
    ),
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
  const numberFormatCode = parseNumberFormatCode(r[0]);
  return numberFormatCode?.hasNumber ? numberFormatCode.maxDecimals : undefined;
}
function numberFormatToIntlOptions(e, numFmt) {
  const t = BN(numFmt);
  if (!t) return;
  const n = e < 0 && t.length > 1,
    o = e === 0 && t.length > 2 ? 2 : n ? 1 : 0,
    numberFormatCode = parseNumberFormatCode(t[o]);
  if (!numberFormatCode) return;
  if (!numberFormatCode.hasNumber)
    return numberFormatCode.prefix + numberFormatCode.suffix;
  const s = Math.abs(e) * (numberFormatCode.percent ? 100 : 1);
  if (!Number.isFinite(s)) return "#NUM!";
  let d = new Intl.NumberFormat("en-US", {
    minimumFractionDigits: numberFormatCode.minDecimals,
    maximumFractionDigits: numberFormatCode.maxDecimals,
    minimumIntegerDigits: Math.max(1, numberFormatCode.minIntegerDigits),
    useGrouping: numberFormatCode.grouping,
  }).format(s);
  return (
    numberFormatCode.minIntegerDigits === 0 &&
      (d === "0" && numberFormatCode.minDecimals === 0
        ? (d = "")
        : d.startsWith("0.") && (d = d.slice(1))),
    (e < 0 && !n ? "-" : "") +
      numberFormatCode.prefix +
      d +
      numberFormatCode.suffix
  );
}
const $b = new Map();
function parseDateTimeFormatCode(e) {
  const r = e.replace(/"[^"]*"|\\./g, "");
  if (!/[ymdhs]/i.test(r)) return;
  const tokens = [];
  let fraction = 0;
  for (let u = 0; u < e.length;) {
    const l = e.slice(u),
      c = l[0];
    if (c === '"') {
      const v = e.indexOf('"', u + 1);
      (v < 0 &&
        throwWorkbookError("The format contains an unclosed quoted literal."),
        tokens.push({ kind: "literal", text: e.slice(u + 1, v) }),
        (u = v + 1));
      continue;
    }
    if (c === "\\") {
      (l.length < 2 &&
        throwWorkbookError("A format escape must be followed by a character."),
        tokens.push({ kind: "literal", text: l[1] }),
        (u += 2));
      continue;
    }
    const h = /^(AM\/PM|A\/P)/i.exec(l);
    if (h) {
      (tokens.push({ kind: "ampm", text: h[0] }), (u += h[0].length));
      continue;
    }
    const f = /^\[(h{1,2}|m{1,2}|s{1,2})\]/i.exec(l);
    if (f) {
      (tokens.push({ kind: "elapsed", text: f[1].toLowerCase() }),
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
        tokens.push({
          kind: { y: "year", m: "month", d: "day", h: "hour", s: "second" }[g],
          text: v,
        }),
        (u += v.length));
      continue;
    }
    const m = /^\.(0{1,3})(?!0)/.exec(l);
    if (m) {
      ((!tokens.length ||
        !["second", "elapsed"].includes(tokens.at(-1).kind) ||
        (tokens.at(-1).kind === "elapsed" &&
          !tokens.at(-1).text.startsWith("s")) ||
        fraction) &&
        throwWorkbookError(
          "Fractional seconds must follow seconds and use 1–3 zero digits.",
        ),
        (fraction = m[1].length),
        tokens.push({ kind: "fraction", text: m[1] }),
        (u += m[0].length));
      continue;
    }
    if (/^[ /:,.-]$/.test(c)) {
      (tokens.push({ kind: "literal", text: c }), u++);
      continue;
    }
    throwWorkbookError(
      "Unsupported date/time code. Quote literal words and use one supported date/time pattern.",
    );
  }
  const o = tokens.filter((token) => token.kind !== "literal");
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
      tokens: tokens,
      fraction: fraction,
    }
  );
}
function parseCustomNumberFormat(e) {
  let r = "",
    divisor = 1,
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
        (divisor = 1e3 ** l.length),
        (r += u.slice(0, u.length - l.length)),
        (s += u.length),
        (n = true));
    } else ((r += d[0]), s++);
  }
  /\[/.test(r.replace(/"[^"]*"|\\./g, "")) &&
    throwWorkbookError(
      "Custom colors, conditions and locale tags are not supported. Use conditional formatting or the locale selector.",
    );
  const numberFormatCode = parseNumberFormatCode(r);
  return (
    numberFormatCode ||
      throwWorkbookError(
        "Unsupported custom format. Use 0/# digits, grouping, decimals, percent, quoted literals and positive/negative/zero sections.",
      ),
    numberFormatCode.maxDecimals > 12 &&
      throwWorkbookError("Custom formats support at most 12 decimal places."),
    { parsed: numberFormatCode, divisor: divisor }
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
