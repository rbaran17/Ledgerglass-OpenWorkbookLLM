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
function failWithMessage(message) {
  throw Error(message);
}
function isPlainObjectLiteral(value) {
  return !!(
    value &&
    typeof value == "object" &&
    !Array.isArray(value) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(value))
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
      Object.keys(line).some(
        (propertyName) => !["style", "color"].includes(propertyName),
      ) ||
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
function validateCellStyle(cellStyle) {
  isPlainObjectLiteral(cellStyle) ||
    failWithMessage("Cell style must be an object.");
  for (const settingKey of Object.keys(cellStyle))
    supportedCellStyleKeys.has(settingKey) ||
      failWithMessage(`Unsupported cell style setting: ${settingKey}.`);
  const validatedStyle = {};
  cellStyle.binding !== undefined &&
    (validatedStyle.binding = validateNamedStyleBinding(cellStyle.binding));
  for (const flagName of [
    "bold",
    "italic",
    "underline",
    "strikethrough",
    "wrap",
  ])
    cellStyle[flagName] !== undefined &&
      (typeof cellStyle[flagName] != "boolean" &&
        failWithMessage(`Style ${flagName} must be boolean.`),
      (validatedStyle[flagName] = cellStyle[flagName]));
  (cellStyle.fontFamily !== undefined &&
    (supportedFontFamilies.includes(cellStyle.fontFamily) ||
      failWithMessage("Unsupported font family."),
    (validatedStyle.fontFamily = cellStyle.fontFamily)),
    cellStyle.fontSize !== undefined &&
      ((typeof cellStyle.fontSize != "number" ||
        !Number.isInteger(cellStyle.fontSize) ||
        cellStyle.fontSize < 8 ||
        cellStyle.fontSize > 72) &&
        failWithMessage("Font size must be an integer between 8 and 72."),
      (validatedStyle.fontSize = cellStyle.fontSize)));
  for (const colorName of ["fill", "color"])
    cellStyle[colorName] !== undefined &&
      (validatedStyle[colorName] = validateStyleColor(cellStyle[colorName]));
  if (
    (cellStyle.align !== undefined &&
      (["left", "center", "right"].includes(cellStyle.align) ||
        failWithMessage("Invalid text alignment."),
      (validatedStyle.align = cellStyle.align)),
    cellStyle.verticalAlign !== undefined &&
      (["top", "middle", "bottom"].includes(cellStyle.verticalAlign) ||
        failWithMessage("Choose Top, Center or Bottom vertical alignment."),
      (validatedStyle.verticalAlign = cellStyle.verticalAlign)),
    cellStyle.indent !== undefined &&
      ((typeof cellStyle.indent != "number" ||
        !Number.isInteger(cellStyle.indent) ||
        cellStyle.indent < 0 ||
        cellStyle.indent > 15) &&
        failWithMessage("Indent must be an integer between 0 and 15."),
      (validatedStyle.indent = cellStyle.indent)),
    cellStyle.border !== undefined &&
      (["all", "bottom", "none"].includes(cellStyle.border) ||
        failWithMessage("Unsupported cell border."),
      (validatedStyle.border = cellStyle.border)),
    cellStyle.borders !== undefined)
  ) {
    (!isPlainObjectLiteral(cellStyle.borders) ||
      Object.keys(cellStyle.borders).some(
        (border) => !cellBorderEdges.includes(border),
      )) &&
      failWithMessage(
        "Cell borders support top, right, bottom and left edges only.",
      );
    const validatedBorders = {};
    for (const cellBorderEdge of cellBorderEdges)
      cellStyle.borders[cellBorderEdge] !== undefined &&
        (validatedBorders[cellBorderEdge] = validateCellBorder(
          cellStyle.borders[cellBorderEdge],
        ));
    Object.keys(validatedBorders).length &&
      (validatedStyle.borders = validatedBorders);
  }
  return validatedStyle;
}
function validateNamedStyleBinding(binding) {
  (!isPlainObjectLiteral(binding) ||
    Object.keys(binding).some(
      (bindingKey) =>
        !["namespace", "styleId", "base", "missing", "themeFallback"].includes(
          bindingKey,
        ),
    )) &&
    failWithMessage(
      "Named style binding must contain a namespace, style ID and supported base.",
    );
  for (const identifierField of ["namespace", "styleId"])
    (typeof binding[identifierField] != "string" ||
      !binding[identifierField].trim() ||
      binding[identifierField].length > 100 ||
      /[\x00-\x1f\x7f]/.test(binding[identifierField])) &&
      failWithMessage(
        "Named style identifiers must contain 1–100 characters without control characters.",
      );
  (!isPlainObjectLiteral(binding.base) ||
    Object.hasOwn(binding.base, "binding")) &&
    failWithMessage("A named style base cannot contain another binding.");
  for (const flagName of ["missing", "themeFallback"])
    binding[flagName] !== undefined &&
      binding[flagName] !== true &&
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
function resolveEffectiveCellStyle(cellStyle) {
  if (!cellStyle) return {};
  const { binding, ...directStyle } = cellStyle;
  if (!binding) return directStyle;
  const baseStyle = binding.missing ? {} : binding.base,
    effectiveStyle = { ...baseStyle, ...directStyle };
  return (
    Object.hasOwn(directStyle, "border")
      ? delete effectiveStyle.borders
      : baseStyle.borders &&
        (effectiveStyle.borders = { ...baseStyle.borders }),
    directStyle.borders &&
      (effectiveStyle.borders = {
        ...effectiveStyle.borders,
        ...directStyle.borders,
      }),
    effectiveStyle
  );
}
function resolveBorderEdge(cellStyle, cellBorderEdge) {
  if (
    ((cellStyle = resolveEffectiveCellStyle(cellStyle)),
    cellStyle?.borders?.[cellBorderEdge])
  )
    return cellStyle.borders[cellBorderEdge];
  if (cellStyle?.border === "none") return { style: "none" };
  if (
    cellStyle?.border === "all" ||
    (cellStyle?.border === "bottom" && cellBorderEdge === "bottom")
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
function indentToPx(cellStyle) {
  return (
    (cellStyle = resolveEffectiveCellStyle(cellStyle)),
    (cellStyle?.indent ?? 0) *
      cellIndentMetrics.indentEm *
      (cellStyle?.fontSize ? (cellStyle.fontSize * 4) / 3 : 14)
  );
}
function cellStyleToCss(cellStyle, cellValue, options = {}) {
  cellStyle = resolveEffectiveCellStyle(cellStyle);
  const textAlign =
      cellStyle?.align ?? (typeof cellValue == "number" ? "right" : "left"),
    indentPx = indentToPx(cellStyle),
    textDecorations = [
      cellStyle?.underline && "underline",
      cellStyle?.strikethrough && "line-through",
    ]
      .filter(Boolean)
      .join(" ");
  return {
    background: cellStyle?.fill,
    color: cellStyle?.color,
    fontWeight: cellStyle?.bold ? 700 : 400,
    fontStyle: cellStyle?.italic ? "italic" : "normal",
    fontFamily: cellStyle?.fontFamily,
    fontSize: cellStyle?.fontSize ? `${cellStyle.fontSize}pt` : undefined,
    textDecoration: textDecorations || "none",
    textAlign: textAlign,
    verticalAlign: cellStyle?.verticalAlign,
    paddingLeft: 3 + (textAlign === "left" ? indentPx : 0),
    paddingRight: 3 + (textAlign === "right" ? indentPx : 0),
    borderTop:
      borderToCssShorthand(resolveBorderEdge(cellStyle, "top")) ??
      options.defaultBorder,
    borderRight:
      borderToCssShorthand(resolveBorderEdge(cellStyle, "right")) ??
      options.defaultBorder,
    borderBottom:
      borderToCssShorthand(resolveBorderEdge(cellStyle, "bottom")) ??
      options.defaultBorder,
    borderLeft:
      borderToCssShorthand(resolveBorderEdge(cellStyle, "left")) ??
      options.defaultBorder,
  };
}
function textWrapToCss(cellStyle) {
  return (
    (cellStyle = resolveEffectiveCellStyle(cellStyle)),
    {
      whiteSpace: cellStyle?.wrap ? "pre-wrap" : "nowrap",
      overflowWrap: cellStyle?.wrap ? "anywhere" : undefined,
    }
  );
}
function measureCellBorderInsets(cellStyle) {
  if (
    ((cellStyle = resolveEffectiveCellStyle(cellStyle)), !cellStyle?.borders)
  ) {
    const uniformInset =
      cellStyle?.border === "all" ? 2 : cellStyle?.border === "none" ? 0 : 1;
    return { horizontal: uniformInset, vertical: uniformInset };
  }
  const edgeInset = (cellBorderEdge) => {
    const borderEdge = resolveBorderEdge(cellStyle, cellBorderEdge);
    return borderEdge ? borderWidthPx(borderEdge) : 0.5;
  };
  return {
    horizontal: edgeInset("left") + edgeInset("right"),
    vertical: edgeInset("top") + edgeInset("bottom"),
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

function assertKnownStyleKeys(styleObject, allowedKeys, label) {
  Object.keys(styleObject).some((key) => !allowedKeys.includes(key)) &&
    failWithMessage(`Unsupported ${label} setting.`);
}
function validateBoundedStyleText(text, label, maxLength = 80) {
  return (
    (typeof text != "string" ||
      !text.trim() ||
      text.length > maxLength ||
      /[\x00-\x1f\x7f]/.test(text)) &&
      failWithMessage(
        `${label} must contain 1–${maxLength} characters without control characters.`,
      ),
    text.trim()
  );
}
function cloneStyleValue(value) {
  return structuredClone(value);
}
function canonicalStyleKey(value) {
  return Array.isArray(value)
    ? `[${value.map(canonicalStyleKey).join(",")}]`
    : isPlainObjectLiteral(value)
      ? `{${Object.keys(value)
          .filter((definedKey) => value[definedKey] !== undefined)
          .sort()
          .map(
            (propertyKey) =>
              `${JSON.stringify(propertyKey)}:${canonicalStyleKey(value[propertyKey])}`,
          )
          .join(",")}}`
      : JSON.stringify(value);
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
    const colorValue = theme.palette[themePaletteRole];
    ((typeof colorValue != "string" || !/^#[0-9a-f]{6}$/i.test(colorValue)) &&
      failWithMessage("Theme colors must use six-digit hexadecimal colors."),
      (palette[themePaletteRole] = colorValue.toUpperCase()));
  }
  for (const fontRole of ["body", "heading"])
    supportedFontFamilies.includes(theme.fonts[fontRole]) ||
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
  const validatedDefinition = { base: validateCellStyle(definition.base) };
  definition.fontRole !== undefined &&
    (["body", "heading"].includes(definition.fontRole) ||
      failWithMessage("Choose Body or Heading theme font."),
    (validatedDefinition.fontRole = definition.fontRole));
  for (const roleField of ["fillRole", "colorRole"])
    definition[roleField] !== undefined &&
      (themePaletteRoles.includes(definition[roleField]) ||
        failWithMessage("Choose a supported theme color role."),
      (validatedDefinition[roleField] = definition[roleField]));
  if (definition.borderRoles !== undefined) {
    (isPlainObjectLiteral(definition.borderRoles) ||
      failWithMessage("Border color roles must be an object."),
      assertKnownStyleKeys(
        definition.borderRoles,
        cellBorderEdges,
        "border role",
      ));
    const validatedBorderRoles = {};
    for (const cellBorderEdge of cellBorderEdges)
      definition.borderRoles[cellBorderEdge] !== undefined &&
        (themePaletteRoles.includes(definition.borderRoles[cellBorderEdge]) ||
          failWithMessage("Choose a supported border color role."),
        (!resolveBorderEdge(validatedDefinition.base, cellBorderEdge) ||
          resolveBorderEdge(validatedDefinition.base, cellBorderEdge)?.style ===
            "none") &&
          failWithMessage(
            "Add a border line before assigning its theme color.",
          ),
        (validatedBorderRoles[cellBorderEdge] =
          definition.borderRoles[cellBorderEdge]));
    Object.keys(validatedBorderRoles).length &&
      (validatedDefinition.borderRoles = validatedBorderRoles);
  }
  return validatedDefinition;
}
function validateWorkbookStyleCatalog(catalog) {
  (isPlainObjectLiteral(catalog) ||
    failWithMessage("Workbook styles must be an object."),
    assertKnownStyleKeys(
      catalog,
      ["namespace", "theme", "styles", "missingTheme"],
      "workbook style catalog",
    ),
    catalog.missingTheme !== undefined &&
      catalog.missingTheme !== true &&
      failWithMessage("Invalid missing theme flag."),
    (!Array.isArray(catalog.styles) ||
      catalog.styles.length > workbookStyleLimits.maxStyles) &&
      failWithMessage("A workbook supports at most 50 named styles."));
  const seenIds = new Set(),
    seenNames = new Set(),
    styles = catalog.styles.map((style) => {
      (isPlainObjectLiteral(style) || failWithMessage("Invalid named style."),
        assertKnownStyleKeys(
          style,
          ["id", "name", "definition"],
          "named style",
        ));
      const styleId = validateBoundedStyleText(style.id, "Style ID", 100),
        styleName = validateBoundedStyleText(style.name, "Style name");
      return (
        (seenIds.has(styleId) ||
          seenNames.has(styleName.toLocaleLowerCase("en-US"))) &&
          failWithMessage(
            "Named style IDs and names must be unique within this workbook.",
          ),
        seenIds.add(styleId),
        seenNames.add(styleName.toLocaleLowerCase("en-US")),
        {
          id: styleId,
          name: styleName,
          definition: validateNamedStyleDefinition(style.definition),
        }
      );
    }),
    isMissingTheme =
      catalog.theme === undefined || catalog.missingTheme === true,
    theme =
      catalog.theme === undefined
        ? cloneStyleValue(officeThemeDefinition)
        : validateWorkbookTheme(catalog.theme);
  return (
    isMissingTheme &&
      canonicalStyleKey(theme) !== canonicalStyleKey(officeThemeDefinition) &&
      failWithMessage("A missing theme must use the Office fallback."),
    {
      namespace: validateBoundedStyleText(
        catalog.namespace,
        "Style namespace",
        100,
      ),
      theme: theme,
      styles: styles,
      ...(isMissingTheme ? { missingTheme: true } : {}),
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
  const appearance = cloneStyleValue(definition.base);
  (definition.fontRole &&
    (appearance.fontFamily = theme.fonts[definition.fontRole]),
    definition.fillRole &&
      (appearance.fill = theme.palette[definition.fillRole]),
    definition.colorRole &&
      (appearance.color = theme.palette[definition.colorRole]));
  for (const cellBorderEdge of cellBorderEdges)
    definition.borderRoles?.[cellBorderEdge] &&
      (appearance.borders = {
        ...appearance.borders,
        [cellBorderEdge]: {
          ...resolveBorderEdge(appearance, cellBorderEdge),
          color: theme.palette[definition.borderRoles[cellBorderEdge]],
        },
      });
  return validateCellStyle(appearance);
}
function resolveNamedStyleBinding(styleCatalog, binding) {
  styleCatalog &&
    binding.namespace !== styleCatalog.namespace &&
    failWithMessage(
      "This style belongs to another workbook. Import or remap its style catalog before pasting the binding.",
    );
  const namedStyle = styleCatalog?.styles.find(
    (style) => style.id === binding.styleId,
  );
  return {
    namespace: binding.namespace,
    styleId: binding.styleId,
    base:
      namedStyle && styleCatalog
        ? resolveNamedStyleAppearance(namedStyle.definition, styleCatalog.theme)
        : {},
    ...(namedStyle ? {} : { missing: true }),
    ...(styleCatalog?.missingTheme ? { themeFallback: true } : {}),
  };
}
function reconcileWorkbookStyleBindings(styleCatalog, workbook) {
  const validatedCatalog =
    styleCatalog === undefined
      ? undefined
      : validateWorkbookStyleCatalog(styleCatalog);
  validatedCatalog
    ? (workbook.styleCatalog = validatedCatalog)
    : delete workbook.styleCatalog;
  const isThemeOmitted =
    isPlainObjectLiteral(styleCatalog) && styleCatalog.theme === undefined;
  for (const sheet of workbook.sheets)
    for (const cell of hasLinkedOrBoundCells(sheet.cells)
      ? Object.values(sheet.cells)
      : []) {
      if (!cell.style?.binding) continue;
      const validatedBinding = validateNamedStyleBinding(cell.style.binding),
        namedStyleBinding = resolveNamedStyleBinding(
          validatedCatalog,
          validatedBinding,
        );
      (!namedStyleBinding.missing &&
        !isThemeOmitted &&
        canonicalStyleKey(validatedBinding) !==
          canonicalStyleKey(namedStyleBinding) &&
        failWithMessage(
          "Named style appearance does not match the workbook catalog. Reapply the style from a trusted catalog.",
        ),
        (cell.style.binding = namedStyleBinding));
    }
}
function assertNamedStyleBindingCurrent(workbook, cellStyle) {
  if (!cellStyle?.binding) return;
  const validatedBinding = validateNamedStyleBinding(cellStyle.binding),
    namedStyleBinding = resolveNamedStyleBinding(
      workbook.styleCatalog,
      validatedBinding,
    );
  canonicalStyleKey(validatedBinding) !==
    canonicalStyleKey(namedStyleBinding) &&
    failWithMessage(
      "Named style binding is stale or inconsistent with this workbook. Reapply its named style.",
    );
}
function describeNamedStyleFallback(cellStyle) {
  if (cellStyle?.binding?.missing)
    return "Named style is missing. Default formatting and direct cell formatting are shown.";
  if (cellStyle?.binding?.themeFallback)
    return "Workbook theme is missing. The Office theme is shown.";
}
function collectStyleFallbackWarnings(workbook) {
  const warnings = new Set();
  workbook.styleCatalog?.missingTheme &&
    warnings.add("Workbook theme is missing. The Office theme is shown.");
  for (const sheet of workbook.sheets)
    for (const cell of Object.values(sheet.cells)) {
      const warning = describeNamedStyleFallback(cell.style);
      warning && warnings.add(warning);
    }
  return [...warnings];
}
function hasBorderEdgeChanges(firstStyle, secondStyle) {
  return cellBorderEdges.some(
    (cellBorderEdge) =>
      canonicalStyleKey(resolveBorderEdge(firstStyle, cellBorderEdge)) !==
      canonicalStyleKey(resolveBorderEdge(secondStyle, cellBorderEdge)),
  );
}
function renameStyleNamespace(sourceWorkbook, newNamespace) {
  const validatedNamespace = validateBoundedStyleText(
      newNamespace,
      "Style namespace",
      100,
    ),
    workbook = cloneStyleValue(sourceWorkbook),
    previousNamespace = workbook.styleCatalog?.namespace;
  if (!previousNamespace) return workbook;
  workbook.styleCatalog.namespace = validatedNamespace;
  const renameBinding = (cellRecord) => {
    cellRecord?.style?.binding?.namespace === previousNamespace &&
      (cellRecord.style.binding.namespace = validatedNamespace);
  };
  for (const sheet of workbook.sheets)
    for (const cell of Object.values(sheet.cells)) renameBinding(cell);
  for (const workbookEvent of workbook.events)
    for (const change of workbookEvent.changes)
      (renameBinding(change.before), renameBinding(change.after));
  return workbook;
}
function validateRowColumnMove(moveRequest, totalCount) {
  if (
    !Number.isInteger(moveRequest.index) ||
    moveRequest.index < 0 ||
    !Number.isInteger(moveRequest.count) ||
    moveRequest.count < 1 ||
    moveRequest.count > 1e4 ||
    moveRequest.index + moveRequest.count > totalCount ||
    !Number.isInteger(moveRequest.moveBefore) ||
    moveRequest.moveBefore < 0 ||
    moveRequest.moveBefore > totalCount
  )
    throw new Error(
      "Choose a valid source, 1–10,000 rows or columns, and a destination within the worksheet.",
    );
  if (moveRequest.permutation !== undefined) {
    if (
      !Array.isArray(moveRequest.permutation) ||
      moveRequest.permutation.length !== moveRequest.count ||
      new Set(moveRequest.permutation).size !== moveRequest.count ||
      moveRequest.permutation.some(
        (destinationIndex) =>
          !Number.isInteger(destinationIndex) ||
          destinationIndex < moveRequest.index ||
          destinationIndex >= moveRequest.index + moveRequest.count,
      )
    )
      throw Error(
        "A row permutation must contain every destination exactly once within its source block.",
      );
    return;
  }
  if (
    moveRequest.moveBefore >= moveRequest.index &&
    moveRequest.moveBefore <= moveRequest.index + moveRequest.count
  )
    throw new Error(
      "The destination is inside or immediately after the selected block; this would not move anything.",
    );
}
function mapIndexAfterMove(originalIndex, move) {
  if (move.permutation)
    return originalIndex >= move.index &&
      originalIndex < move.index + move.count
      ? move.permutation[originalIndex - move.index]
      : originalIndex;
  const { index, count, moveBefore } = move,
    movedBlockStart = moveBefore < index ? moveBefore : moveBefore - count;
  return originalIndex >= index && originalIndex < index + count
    ? movedBlockStart + originalIndex - index
    : moveBefore < index && originalIndex >= moveBefore && originalIndex < index
      ? originalIndex + count
      : moveBefore > index &&
          originalIndex >= index + count &&
          originalIndex < moveBefore
        ? originalIndex - count
        : originalIndex;
}
function planMoveSegments(start, end, move) {
  if (move.permutation) {
    const segments = [],
      pushSegment = (start, end, offset) => {
        if (start > end) return;
        const lastSegment = segments.at(-1);
        lastSegment &&
        lastSegment.end + 1 === start &&
        lastSegment.offset === offset
          ? (lastSegment.end = end)
          : segments.push({ start: start, end: end, offset: offset });
      };
    pushSegment(start, Math.min(end, move.index - 1), 0);
    for (
      let blockIndex = Math.max(start, move.index);
      blockIndex <= Math.min(end, move.index + move.count - 1);
      blockIndex++
    )
      pushSegment(
        blockIndex,
        blockIndex,
        mapIndexAfterMove(blockIndex, move) - blockIndex,
      );
    return (
      pushSegment(Math.max(start, move.index + move.count), end, 0),
      segments
    );
  }
  const boundaries = [
    ...new Set(
      [
        start,
        end + 1,
        move.index,
        move.index + move.count,
        move.moveBefore,
      ].filter((boundary) => boundary >= start && boundary <= end + 1),
    ),
  ].sort((leftBoundary, rightBoundary) => leftBoundary - rightBoundary);
  return boundaries.slice(0, -1).map((start, segmentIndex) => ({
    start: start,
    end: boundaries[segmentIndex + 1] - 1,
    offset: mapIndexAfterMove(start, move) - start,
  }));
}
function mapRangeAfterMove(start, end, move) {
  if (
    move.permutation &&
    start <= move.index &&
    end >= move.index + move.count - 1
  )
    return [{ start: start, end: end }];
  const shiftedSegments = planMoveSegments(start, end, move)
      .map((segment) => ({
        start: segment.start + segment.offset,
        end: segment.end + segment.offset,
      }))
      .sort((leftRange, rightRange) => leftRange.start - rightRange.start),
    mergedRanges = [];
  for (const range of shiftedSegments) {
    const lastMergedRange = mergedRanges.at(-1);
    lastMergedRange && range.start <= lastMergedRange.end + 1
      ? (lastMergedRange.end = Math.max(lastMergedRange.end, range.end))
      : mergedRanges.push(range);
  }
  return mergedRanges;
}
function resolveMovedRange(
  start,
  end,
  move,
  rangeLabel = "range",
  isStructureMove = false,
) {
  const movedRanges = mapRangeAfterMove(start, end, move);
  if (movedRanges.length !== 1)
    throw new Error(
      `This move would split ${rangeLabel} into separate ranges. Move the complete range or use explicit cell references first.`,
    );
  if (
    isStructureMove &&
    new Set(planMoveSegments(start, end, move).map((segment) => segment.offset))
      .size !== 1
  )
    throw new Error(
      `This move would reorder part of ${rangeLabel}. Move the complete structure together.`,
    );
  return movedRanges[0];
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
  throwWithMessage = (message) => {
    throw new Error(message);
  };
function assertPlainReviewRecord(value, label) {
  return (
    (!value ||
      typeof value != "object" ||
      Array.isArray(value) ||
      ![Object.prototype, null].includes(Object.getPrototypeOf(value))) &&
      throwWithMessage(`${label} must be a plain object.`),
    value
  );
}
function normalizeBoundedText(text, label, maxLength, isMultiline = false) {
  return (
    (typeof text != "string" ||
      !text.trim() ||
      text.length > maxLength ||
      unsafeControlCharPattern.test(
        isMultiline ? text.replace(/[\n\r\t]/g, "") : text,
      )) &&
      throwWithMessage(
        `${label} must contain 1–${maxLength} characters without unsupported control characters.`,
      ),
    text.trim()
  );
}
function normalizeReviewTimestamp(timestamp, label = "Review timestamp") {
  return (
    (typeof timestamp != "string" ||
      !Number.isFinite(Date.parse(timestamp)) ||
      new Date(timestamp).toISOString() !== timestamp) &&
      throwWithMessage(`${label} must be an ISO UTC timestamp.`),
    timestamp
  );
}
function normalizeReviewAuthor(author = "Reviewer") {
  return normalizeBoundedText(
    author,
    "Reviewer name",
    reviewThreadLimits.authorLength,
  );
}
function resolveReviewTimestamp(options = {}, updatedAt) {
  const timestamp =
    options.timestamp === undefined
      ? new Date(
          Math.max(Date.now(), updatedAt ? Date.parse(updatedAt) : 0),
        ).toISOString()
      : normalizeReviewTimestamp(options.timestamp);
  return (
    updatedAt &&
      timestamp < updatedAt &&
      throwWithMessage(
        "The review time cannot precede the existing discussion.",
      ),
    timestamp
  );
}
function aQ(workbook, timestamp) {
  const sheetsById = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet]));
  let hasChanges = false;
  const reviewThreads = workbook.reviewThreads?.map((reviewThread) => {
      if (reviewThread.orphaned) return reviewThread;
      const anchorSheet = sheetsById.get(reviewThread.anchor.sheetId);
      if (!anchorSheet) {
        hasChanges = true;
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
      return anchorSheet.name !== reviewThread.anchor.sheetName
        ? ((hasChanges = true),
          {
            ...reviewThread,
            anchor: { ...reviewThread.anchor, sheetName: anchorSheet.name },
          })
        : reviewThread;
    }),
    sheets = workbook.sheets.map((sheet) => {
      let cells;
      for (const [cellAddress, cell] of Object.entries(sheet.cells))
        cell.hyperlink?.kind !== "internal" ||
          sheetsById.has(cell.hyperlink.sheetId) ||
          ((cells ??= { ...sheet.cells }),
          (cells[cellAddress] = { ...cell }),
          delete cells[cellAddress].hyperlink,
          (hasChanges = true));
      return cells ? { ...sheet, cells: cells } : sheet;
    });
  return hasChanges
    ? {
        ...workbook,
        sheets: sheets,
        ...(reviewThreads ? { reviewThreads: reviewThreads } : {}),
      }
    : workbook;
}
function uQ(workbook, sheetIdMap) {
  return {
    ...workbook,
    ...(workbook.reviewThreads
      ? {
          reviewThreads: workbook.reviewThreads.map((reviewThread) => ({
            ...reviewThread,
            anchor: {
              ...reviewThread.anchor,
              sheetId:
                sheetIdMap.get(reviewThread.anchor.sheetId) ??
                reviewThread.anchor.sheetId,
            },
          })),
        }
      : {}),
    sheets: workbook.sheets.map((sheet) => ({
      ...sheet,
      cells: Object.fromEntries(
        Object.entries(sheet.cells).map(([cellAddress, cell]) => [
          cellAddress,
          cell.hyperlink?.kind === "internal"
            ? {
                ...cell,
                hyperlink: {
                  ...cell.hyperlink,
                  sheetId:
                    sheetIdMap.get(cell.hyperlink.sheetId) ??
                    cell.hyperlink.sheetId,
                },
              }
            : cell,
        ]),
      ),
    })),
  };
}
function createReviewThread(workbook, sheetIdMap) {
  const remappedWorkbook = uQ(workbook, sheetIdMap);
  return remappedWorkbook.reviewThreads
    ? {
        ...remappedWorkbook,
        reviewThreads: remappedWorkbook.reviewThreads.map((reviewThread) => ({
          ...reviewThread,
          id: crypto.randomUUID(),
          messages: reviewThread.messages.map((message) => ({
            ...message,
            id: crypto.randomUUID(),
          })),
        })),
      }
    : remappedWorkbook;
}
function fQ(workbook, sheetId) {
  return workbook.formulaWatches?.length
    ? {
        ...workbook,
        formulaWatches: workbook.formulaWatches.map((formulaWatch) =>
          !formulaWatch.orphaned && formulaWatch.anchor.sheetId === sheetId
            ? { ...formulaWatch, orphaned: "deleted-sheet" }
            : formulaWatch,
        ),
      }
    : workbook;
}
function createFormulaWatch(workbook, sheetIdMap) {
  return workbook.formulaWatches?.length
    ? {
        ...workbook,
        formulaWatches: workbook.formulaWatches.map((formulaWatch) => ({
          ...formulaWatch,
          id: crypto.randomUUID(),
          anchor: {
            ...formulaWatch.anchor,
            sheetId:
              sheetIdMap.get(formulaWatch.anchor.sheetId) ??
              formulaWatch.anchor.sheetId,
          },
        })),
      }
    : workbook;
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
function anchorsMatchOnAxis(firstAnchor, secondAnchor, axis) {
  return (
    firstAnchor.split("-")[axis === "row" ? 0 : 1] ===
    secondAnchor.split("-")[axis === "row" ? 0 : 1]
  );
}
function validateSplitPaneState(splitPanes, dimensions) {
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
  const splitPaneSpec = splitPanes;
  if (
    splitPaneSpec.horizontal === undefined &&
    splitPaneSpec.vertical === undefined
  )
    throw Error("Choose a horizontal split, vertical split, or both.");
  for (const dividerPercent of [
    splitPaneSpec.horizontal,
    splitPaneSpec.vertical,
  ])
    if (
      dividerPercent !== undefined &&
      (!Number.isFinite(dividerPercent) ||
        dividerPercent < 20 ||
        dividerPercent > 80)
    )
      throw Error("Split dividers must be between 20% and 80%.");
  if (
    !Number.isInteger(splitPaneSpec.row) ||
    splitPaneSpec.row < 0 ||
    splitPaneSpec.row >= dimensions.rows ||
    !Number.isInteger(splitPaneSpec.col) ||
    splitPaneSpec.col < 0 ||
    splitPaneSpec.col >= dimensions.cols
  )
    throw Error(
      "Split anchors must be valid zero-based worksheet coordinates.",
    );
  return {
    row: splitPaneSpec.row,
    col: splitPaneSpec.col,
    ...(splitPaneSpec.horizontal !== undefined
      ? { horizontal: splitPaneSpec.horizontal }
      : {}),
    ...(splitPaneSpec.vertical !== undefined
      ? { vertical: splitPaneSpec.vertical }
      : {}),
  };
}
function validateDefaultSizePixels(sizePixels, axis) {
  const minSize = axis === "row" ? minRowHeight : minColumnWidth,
    maxSize = axis === "row" ? maxRowHeight : maxColumnWidth;
  if (
    typeof sizePixels != "number" ||
    !Number.isFinite(sizePixels) ||
    sizePixels < minSize ||
    sizePixels > maxSize
  )
    throw Error(
      (axis === "row" ? "Default row height" : "Default column width") +
        " must be " +
        minSize +
        "–" +
        maxSize +
        " pixels.",
    );
  return sizePixels;
}
function validateSheetDefaultSizes(sourceSheet, targetSheet) {
  const validatedRowHeight =
      sourceSheet.defaultRowHeight === undefined
        ? undefined
        : validateDefaultSizePixels(sourceSheet.defaultRowHeight, "row"),
    validatedColumnWidth =
      sourceSheet.defaultColumnWidth === undefined
        ? undefined
        : validateDefaultSizePixels(sourceSheet.defaultColumnWidth, "column");
  (validatedRowHeight !== undefined &&
    (targetSheet.defaultRowHeight = validatedRowHeight),
    validatedColumnWidth !== undefined &&
      (targetSheet.defaultColumnWidth = validatedColumnWidth));
}
function defaultRowHeightOf(sheet) {
  return sheet.defaultRowHeight ?? defaultRowHeight;
}
function defaultColumnWidthOf(sheet) {
  return sheet.defaultColumnWidth ?? defaultColumnWidth;
}
function rowHeightAt(sheet, rowIndex) {
  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= sheet.rowCount)
    throw Error("Row index is outside the worksheet.");
  return sheet.rowHeights?.[rowIndex] ?? defaultRowHeightOf(sheet);
}
function columnWidthAt(sheet, columnIndex) {
  if (
    !Number.isInteger(columnIndex) ||
    columnIndex < 0 ||
    columnIndex >= sheet.colCount
  )
    throw Error("Column index is outside the worksheet.");
  return sheet.columnWidths?.[columnIndex] ?? defaultColumnWidthOf(sheet);
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
  axisCountOf = (sheet, axis) =>
    axis === "row" ? sheet.rowCount : sheet.colCount,
  hiddenKeyForAxis = (axis) => (axis === "row" ? "hiddenRows" : "hiddenCols");
function assertWorksheetAxis(axis) {
  if (axis !== "row" && axis !== "column")
    throw Error("Choose rows or columns.");
}
function assertAxisCapacity(capacity) {
  if (!Number.isSafeInteger(capacity) || capacity < 0)
    throw Error("Axis capacity must be a nonnegative integer.");
}
function normalizeHiddenIntervals(rawIntervals, axisCount) {
  if ((assertAxisCapacity(axisCount), rawIntervals === undefined)) return [];
  if (
    !Array.isArray(rawIntervals) ||
    rawIntervals.length > worksheetCapacityLimits.maxIntervalsPerAxis
  )
    throw Error("Hidden indices require at most 10,000 intervals per axis.");
  const sortedIntervals = rawIntervals
      .map((rawInterval) => {
        if (
          !rawInterval ||
          typeof rawInterval != "object" ||
          Array.isArray(rawInterval) ||
          (Object.getPrototypeOf(rawInterval) !== Object.prototype &&
            Object.getPrototypeOf(rawInterval) !== null) ||
          Object.keys(rawInterval).some(
            (key) => key !== "start" && key !== "end",
          ) ||
          !Number.isInteger(rawInterval.start) ||
          !Number.isInteger(rawInterval.end) ||
          rawInterval.start < 0 ||
          rawInterval.end < rawInterval.start ||
          rawInterval.end >= axisCount
        )
          throw Error(
            "Hidden intervals require valid zero-based inclusive axis bounds.",
          );
        return { start: rawInterval.start, end: rawInterval.end };
      })
      .sort((left, right) => left.start - right.start || left.end - right.end),
    merged = [];
  for (const sortedInterval of sortedIntervals) {
    const previousInterval = merged.at(-1);
    previousInterval && sortedInterval.start <= previousInterval.end + 1
      ? (previousInterval.end = Math.max(
          previousInterval.end,
          sortedInterval.end,
        ))
      : merged.push(sortedInterval);
  }
  return merged;
}
function createAxisVisibilityIndex(logicalCount, rawHiddenIntervals = []) {
  const hiddenIntervals = normalizeHiddenIntervals(
      rawHiddenIntervals,
      logicalCount,
    ),
    cumulativeHiddenCounts = [0];
  for (const hiddenInterval of hiddenIntervals)
    cumulativeHiddenCounts.push(
      cumulativeHiddenCounts.at(-1) +
        hiddenInterval.end -
        hiddenInterval.start +
        1,
    );
  const count = logicalCount - cumulativeHiddenCounts.at(-1),
    findIntervalIndex = (logicalPosition) => {
      let low = 0,
        high = hiddenIntervals.length;
      for (; low < high;) {
        const middle = (low + high) >>> 1;
        hiddenIntervals[middle].start <= logicalPosition
          ? (low = middle + 1)
          : (high = middle);
      }
      return low - 1;
    },
    countBefore = (boundary) => {
      const clampedBoundary = Math.max(
          0,
          Math.min(logicalCount, Math.floor(boundary)),
        ),
        containingIndex = findIntervalIndex(clampedBoundary - 1),
        hiddenBefore =
          containingIndex < 0
            ? 0
            : cumulativeHiddenCounts[containingIndex] +
              Math.min(
                clampedBoundary - hiddenIntervals[containingIndex].start,
                hiddenIntervals[containingIndex].end -
                  hiddenIntervals[containingIndex].start +
                  1,
              );
      return clampedBoundary - hiddenBefore;
    },
    isHidden = (candidate) => {
      const hiddenIndex = findIntervalIndex(candidate);
      return hiddenIndex >= 0 && candidate <= hiddenIntervals[hiddenIndex].end;
    },
    logicalIndex = (visibleIndex) => {
      if (
        !Number.isInteger(visibleIndex) ||
        visibleIndex < 0 ||
        visibleIndex >= count
      )
        return -1;
      let lowerBound = 0,
        upperBound = logicalCount - 1;
      for (; lowerBound < upperBound;) {
        const midpoint = Math.floor((lowerBound + upperBound) / 2);
        countBefore(midpoint + 1) <= visibleIndex
          ? (lowerBound = midpoint + 1)
          : (upperBound = midpoint);
      }
      return lowerBound;
    };
  return {
    count: count,
    logicalCount: logicalCount,
    logicalIndex: logicalIndex,
    countBefore: countBefore,
    isHidden: isHidden,
    viewIndex: (sourceIndex) =>
      Number.isInteger(sourceIndex) &&
      sourceIndex >= 0 &&
      sourceIndex < logicalCount &&
      !isHidden(sourceIndex)
        ? countBefore(sourceIndex)
        : null,
    next: (fromIndex, step, shouldIncludeCurrent = false) => {
      const startIndex =
        Math.trunc(fromIndex) + (shouldIncludeCurrent ? 0 : step);
      if (
        !count ||
        (step === 1 && startIndex >= logicalCount) ||
        (step === -1 && startIndex < 0)
      )
        return null;
      const viewPosition =
          step === 1
            ? countBefore(Math.max(0, startIndex))
            : countBefore(Math.min(logicalCount, startIndex + 1)) - 1,
        nextLogicalIndex = logicalIndex(viewPosition);
      return nextLogicalIndex < 0 ? null : nextLogicalIndex;
    },
  };
}
function readHiddenAxisIntervals(sheet, axis) {
  return (
    assertWorksheetAxis(axis),
    normalizeHiddenIntervals(
      sheet[hiddenKeyForAxis(axis)],
      axisCountOf(sheet, axis),
    )
  );
}
const outlineGroupLimits = { maxGroups: 1e3, maxDepth: 8 };

function assertOutlineGroupNesting(outlineGroups, axis) {
  const axisGroups = outlineGroups
      .filter((outlineGroup) => outlineGroup.axis === axis)
      .sort(
        (leftGroup, rightGroup) =>
          leftGroup.start - rightGroup.start || rightGroup.end - leftGroup.end,
      ),
    openGroups = [];
  return axisGroups.map((group) => {
    for (; openGroups.length && openGroups.at(-1).end < group.start;)
      openGroups.pop();
    const parentGroup = openGroups.at(-1);
    if (
      parentGroup &&
      parentGroup.start === group.start &&
      parentGroup.end === group.end
    )
      throw Error("An outline group already exists for these rows or columns.");
    if (parentGroup && parentGroup.end < group.end)
      throw Error(
        "Outline groups must be nested or separate; partially overlapping groups are not supported.",
      );
    if (
      (openGroups.push(group), openGroups.length > outlineGroupLimits.maxDepth)
    )
      throw Error("Outline groups support at most 8 nested groups.");
    return { ...group, depth: openGroups.length };
  });
}
function validateOutlineGroups(
  sheet,
  targetSheet,
  bounds = { rows: targetSheet.rowCount, cols: targetSheet.colCount },
) {
  if (sheet.outlineGroups === undefined) return;
  if (
    !Array.isArray(sheet.outlineGroups) ||
    sheet.outlineGroups.length > outlineGroupLimits.maxGroups
  )
    throw Error("A worksheet supports at most 1,000 outline groups.");
  const seenIds = new Set(),
    validGroups = sheet.outlineGroups.map((outlineGroup) => {
      if (
        !outlineGroup ||
        typeof outlineGroup != "object" ||
        Array.isArray(outlineGroup) ||
        (Object.getPrototypeOf(outlineGroup) !== Object.prototype &&
          Object.getPrototypeOf(outlineGroup) !== null) ||
        Object.keys(outlineGroup).some(
          (key) => !["id", "axis", "start", "end", "collapsed"].includes(key),
        ) ||
        typeof outlineGroup.id != "string" ||
        !outlineGroup.id.trim() ||
        outlineGroup.id.length > 100 ||
        seenIds.has(outlineGroup.id) ||
        !["row", "column"].includes(outlineGroup.axis) ||
        !Number.isSafeInteger(outlineGroup.start) ||
        !Number.isSafeInteger(outlineGroup.end) ||
        outlineGroup.start < 0 ||
        outlineGroup.end < outlineGroup.start ||
        outlineGroup.end >=
          (outlineGroup.axis === "row" ? bounds.rows : bounds.cols) ||
        typeof outlineGroup.collapsed != "boolean"
      )
        throw Error(
          "Outline groups require unique IDs, valid row or column bounds, and a boolean collapsed state.",
        );
      return (
        seenIds.add(outlineGroup.id),
        {
          id: outlineGroup.id,
          axis: outlineGroup.axis,
          start: outlineGroup.start,
          end: outlineGroup.end,
          collapsed: outlineGroup.collapsed,
        }
      );
    });
  (assertOutlineGroupNesting(validGroups, "row"),
    assertOutlineGroupNesting(validGroups, "column"),
    (targetSheet.outlineGroups = validGroups));
}
function addOutlineGroup(sheet, axis) {
  return (
    assertWorksheetAxis(axis),
    assertOutlineGroupNesting(sheet.outlineGroups ?? [], axis)
  );
}
function cn(sheet, axis) {
  assertWorksheetAxis(axis);
  const axisCount = axisCountOf(sheet, axis),
    combinedIntervals = [
      ...normalizeHiddenIntervals(
        axis === "row" ? sheet.hiddenRows : sheet.hiddenCols,
        axisCount,
      ),
      ...(sheet.outlineGroups ?? [])
        .filter((group) => group.axis === axis && group.collapsed)
        .map(({ start, end }) => ({ start, end })),
    ].sort((left, right) => left.start - right.start || left.end - right.end),
    merged = [];
  for (const interval of combinedIntervals) {
    const previousInterval = merged.at(-1);
    previousInterval && interval.start <= previousInterval.end + 1
      ? (previousInterval.end = Math.max(previousInterval.end, interval.end))
      : merged.push({ ...interval });
  }
  return normalizeHiddenIntervals(merged, axisCount);
}
function regenerateOutlineGroupIds(sheet) {
  return sheet.outlineGroups?.map((outlineGroup) => ({
    ...outlineGroup,
    id: crypto.randomUUID(),
  }));
}
function remapOutlineGroupsForEdit(sheet, axis, start, count, remove) {
  assertWorksheetAxis(axis);
  const axisCount = axisCountOf(sheet, axis);
  if (
    !Number.isInteger(start) ||
    start < 0 ||
    !Number.isInteger(count) ||
    count < 1 ||
    count > 1e4 ||
    start + count > axisCount ||
    typeof remove != "boolean"
  )
    throw Error("Choose a valid structural edit of 1–10,000 rows or columns.");
  if (!sheet.outlineGroups) return;
  const editEnd = start + count - 1,
    remappedGroups = sheet.outlineGroups.flatMap((outlineGroup) => {
      if (outlineGroup.axis !== axis || outlineGroup.end < start)
        return [{ ...outlineGroup }];
      if (!remove) {
        if (outlineGroup.end + count >= axisCount)
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
      return outlineGroup.start > editEnd
        ? [
            {
              ...outlineGroup,
              start: outlineGroup.start - count,
              end: outlineGroup.end - count,
            },
          ]
        : outlineGroup.start >= start && outlineGroup.end <= editEnd
          ? []
          : [
              {
                ...outlineGroup,
                start: Math.min(outlineGroup.start, start),
                end:
                  outlineGroup.end > editEnd
                    ? outlineGroup.end - count
                    : start - 1,
              },
            ];
    });
  try {
    (assertOutlineGroupNesting(remappedGroups, "row"),
      assertOutlineGroupNesting(remappedGroups, "column"));
  } catch {
    throw Error(
      "This edit would make outline groups overlap or duplicate. Ungroup the affected rows or columns first.",
    );
  }
  return remappedGroups;
}
const isPlainObjectValue = (value) =>
  !!value &&
  typeof value == "object" &&
  !Array.isArray(value) &&
  [Object.prototype, null].includes(Object.getPrototypeOf(value));
function trimRequiredMetadataText(text, label, maxLength) {
  if (typeof text != "string" || !text.trim() || text.length > maxLength)
    throw new Error(`${label} must contain 1–${maxLength} characters.`);
  return text.trim();
}
function isValidUtcTimestamp(timestamp) {
  if (
    typeof timestamp != "string" ||
    !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(timestamp) ||
    !Number.isFinite(Date.parse(timestamp)) ||
    new Date(timestamp).toISOString() !== timestamp
  )
    throw new Error("Workbook metadata requires a valid UTC timestamp.");
  return timestamp;
}
function validateWorkbookMetadata(metadata, workbook) {
  if (metadata === undefined) return;
  if (
    !isPlainObjectValue(metadata) ||
    Object.keys(metadata).some(
      (key) =>
        ![
          "createdAt",
          "modifiedAt",
          "engagement",
          "period",
          "template",
        ].includes(key),
    )
  )
    throw new Error("Invalid workbook metadata fields.");
  const createdAt = isValidUtcTimestamp(metadata.createdAt),
    modifiedAt = isValidUtcTimestamp(metadata.modifiedAt);
  if (modifiedAt < createdAt)
    throw new Error("Workbook modified time cannot precede its creation time.");
  const normalizedMetadata = { createdAt: createdAt, modifiedAt: modifiedAt };
  if (
    (metadata.engagement !== undefined &&
      (normalizedMetadata.engagement = trimRequiredMetadataText(
        metadata.engagement,
        "Engagement",
        200,
      )),
    metadata.period !== undefined &&
      (normalizedMetadata.period = trimRequiredMetadataText(
        metadata.period,
        "Period",
        100,
      )),
    metadata.template !== undefined)
  ) {
    const template = metadata.template;
    if (
      !isPlainObjectValue(template) ||
      Object.keys(template).some(
        (key) =>
          ![
            "id",
            "name",
            "version",
            "sourceWorkbookId",
            "sourceRevision",
          ].includes(key),
      ) ||
      !Number.isSafeInteger(template.version) ||
      Number(template.version) < 1 ||
      (template.sourceRevision !== undefined &&
        (!Number.isSafeInteger(template.sourceRevision) ||
          Number(template.sourceRevision) < 0))
    )
      throw new Error("Invalid template provenance.");
    normalizedMetadata.template = {
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
  workbook.metadata = normalizedMetadata;
}
function createWorkbookMetadata(
  overrides = {},
  timestamp = new Date().toISOString(),
) {
  const metadataHolder = {};
  return (
    validateWorkbookMetadata(
      { createdAt: timestamp, modifiedAt: timestamp, ...overrides },
      metadataHolder,
    ),
    metadataHolder.metadata
  );
}
function touchWorkbookModifiedAt(
  workbook,
  timestamp = new Date().toISOString(),
) {
  const updatedWorkbook = { ...workbook };
  isValidUtcTimestamp(timestamp);
  const modifiedAt = [
    timestamp,
    workbook.metadata?.createdAt ?? timestamp,
    workbook.metadata?.modifiedAt ?? timestamp,
  ]
    .sort()
    .at(-1);
  return (
    validateWorkbookMetadata(
      workbook.metadata
        ? { ...workbook.metadata, modifiedAt: modifiedAt }
        : createWorkbookMetadata({}, timestamp),
      updatedWorkbook,
    ),
    updatedWorkbook
  );
}
const unquotedNameCharPattern = /[A-Za-z0-9_.\\\u0080-\uffff]/,
  sheetReferencePrefixPattern =
    /^(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*)(?::(?:'(?:[^']|'')*'|[A-Za-z_\\\u0080-\uffff][A-Za-z0-9_.\\\u0080-\uffff]*))?!/,
  columnLettersToIndex = (columnLetters) =>
    [...columnLetters.toUpperCase()].reduce(
      (total, letter) => total * 26 + letter.charCodeAt(0) - 64,
      0,
    ) - 1;
function offsetRowReference(rowPrefix, rowValue, rowDelta) {
  const rowNumber = Number(rowValue) + (rowPrefix ? 0 : rowDelta);
  return rowNumber < 1 || rowNumber > workbookLimits.maxRows
    ? "#REF!"
    : rowPrefix + rowNumber;
}
const cloneCellRecord = (cell) =>
  cell
    ? { ...cell, ...(cell.style ? { style: { ...cell.style } } : {}) }
    : { raw: null };
const validationLimits = Object.freeze({
    rulesPerSheet: 100,
    listItems: 200,
    checks: 1e5,
    customChecks: 5e3,
    displayedIssues: 500,
  }),
  isNonArrayObject = (value) =>
    !!value && typeof value == "object" && !Array.isArray(value),
  assertBoundedValidationText = (text, maxLength, label) => {
    if (typeof text != "string" || text.length > maxLength)
      throw Error(`Invalid validation ${label}.`);
    return text;
  };
function validateDataValidationRule(
  rule,
  value,
  datePolicy = defaultDatePolicy,
) {
  if (value === null || value === "") return rule.allowBlank;
  if (rule.kind === "custom") return false;
  if (rule.kind === "list")
    return rule.list.some(
      (listItem) => listItem.toLowerCase() === String(value).toLowerCase(),
    );
  let numericValue =
    rule.kind === "textLength"
      ? String(value).length
      : typeof value == "number"
        ? value
        : NaN;
  if (
    (["date", "time"].includes(rule.kind) &&
      typeof value == "string" &&
      (numericValue = parseDateText(value, datePolicy) ?? NaN),
    !Number.isFinite(numericValue) ||
      (rule.kind === "integer" && !Number.isInteger(numericValue)) ||
      (rule.kind === "date" &&
        (numericValue < 1 ||
          numericValue > dateCalculationLimits.maxSerial + 0.999999999 ||
          (typeof value == "string" && !/[/-]/.test(value)))) ||
      (rule.kind === "time" && (numericValue < 0 || numericValue >= 1)))
  )
    return false;
  switch (rule.operator) {
    case "between":
      return numericValue >= rule.min && numericValue <= rule.max;
    case "notBetween":
      return numericValue < rule.min || numericValue > rule.max;
    case "greater":
      return numericValue > rule.min;
    case "greaterEqual":
      return numericValue >= rule.min;
    case "less":
      return numericValue < rule.min;
    case "lessEqual":
      return numericValue <= rule.min;
    case "notEqual":
      return numericValue !== rule.min;
    default:
      return numericValue === rule.min;
  }
}
const MB = new WeakMap(),
  r5 = new WeakMap();
class DataValidationError extends Error {
  constructor(book, issues) {
    super(
      issues
        .slice(0, 5)
        .map(
          (issue) =>
            `${issue.sheetName}!${issue.address}: ${issue.message} ${issue.reason}`,
        ).join(`
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
  const serializedError = MB.get(error);
  if (
    !serializedError ||
    !error.issues.length ||
    error.issues.some(
      (issue) => issue.style === "stop" || issue.evaluationError,
    ) ||
    JSON.stringify(JSON.parse(serializedError).issues) !==
      JSON.stringify(error.issues)
  )
    throw Error(
      "Stop errors cannot be overridden. Correct the invalid values.",
    );
  const approval = Object.freeze({ kind: "validation-warning-approval" });
  return (r5.set(approval, serializedError), approval);
}
const FB = new WeakMap();
function resolveValidationInspection(workbook, report) {
  const inspection = FB.get(report);
  if (
    !inspection ||
    inspection.source !== JSON.stringify(workbook) ||
    inspection.content !== JSON.stringify(report)
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
function vx(sheet, evaluator) {
  if (!evaluator.getSpillEntries) return sheet.cells;
  const spillEntries = evaluator.getSpillEntries(sheet.id);
  if (!spillEntries.length) return sheet.cells;
  const cellsWithSpills = { ...sheet.cells };
  for (const { address } of spillEntries)
    cellsWithSpills[address] = {
      ...sheet.cells[address],
      raw: evaluator(sheet.id, address),
    };
  return cellsWithSpills;
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
function assertBoundedString(text, maxLength, label) {
  if (typeof text != "string" || text.length > maxLength)
    throw Error(`Invalid ${label}.`);
  return text;
}
function assertSixDigitHexColor(color) {
  if (typeof color != "string" || !/^#[\da-f]{6}$/i.test(color))
    throw Error("Choose a six-digit hexadecimal color.");
  return color;
}
const isCellInsideRange = (range, cell) =>
    cell.row >= range.firstRow &&
    cell.row <= range.lastRow &&
    cell.col >= range.firstCol &&
    cell.col <= range.lastCol,
  isErrorValueText = (value) =>
    typeof value == "string" &&
    /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|ERROR!|CYCLE!|SPILL!|CALC!|LIMIT!|BLOCKED!)/.test(
      value,
    );
function sortConditionalRulesByPriority(rules) {
  return rules
    .map((rule, ruleIndex) => ({
      rule: rule,
      rank: rule.priority ?? 101 + rules.length - ruleIndex,
    }))
    .sort((leftEntry, rightEntry) => leftEntry.rank - rightEntry.rank)
    .map((entry) => entry.rule);
}

function assertSheetUnprotected(sheet) {
  sheet.protection?.enabled &&
    failWithMessage(
      "This sheet is protected. Unprotect it before changing its contents.",
    );
}
function applySheetSizeEdits(workbook, sheetId, sizeEdits, options = {}) {
  if (options.readOnly !== undefined && typeof options.readOnly != "boolean")
    throw Error("Read-only state must be boolean.");
  if (options.readOnly) throw Error("This workbook is read-only.");
  const expectedSnapshot = options.expectedSnapshot;
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
    !sizeEdits ||
      typeof sizeEdits != "object" ||
      Array.isArray(sizeEdits) ||
      (Object.getPrototypeOf(sizeEdits) !== Object.prototype &&
        Object.getPrototypeOf(sizeEdits) !== null) ||
      Object.keys(sizeEdits).some(
        (key) =>
          ![
            "rowHeight",
            "columnWidth",
            "resetRowOverrides",
            "resetColumnOverrides",
          ].includes(key),
      ))
  )
    throw Error("Choose valid worksheet size settings.");
  for (const resetChoice of [
    sizeEdits.resetRowOverrides,
    sizeEdits.resetColumnOverrides,
  ])
    if (resetChoice !== undefined && typeof resetChoice != "boolean")
      throw Error("Reset choices must be boolean.");
  const rowHeight =
      sizeEdits.rowHeight === undefined
        ? undefined
        : validateDefaultSizePixels(sizeEdits.rowHeight, "row"),
    columnWidth =
      sizeEdits.columnWidth === undefined
        ? undefined
        : validateDefaultSizePixels(sizeEdits.columnWidth, "column"),
    updatedSheet = { ...sheet };
  return (
    rowHeight !== undefined &&
      rowHeight !== defaultRowHeightOf(sheet) &&
      (rowHeight === defaultRowHeight
        ? delete updatedSheet.defaultRowHeight
        : (updatedSheet.defaultRowHeight = rowHeight)),
    columnWidth !== undefined &&
      columnWidth !== defaultColumnWidthOf(sheet) &&
      (columnWidth === defaultColumnWidth
        ? delete updatedSheet.defaultColumnWidth
        : (updatedSheet.defaultColumnWidth = columnWidth)),
    sizeEdits.resetRowOverrides &&
      Object.keys(sheet.rowHeights ?? {}).length &&
      delete updatedSheet.rowHeights,
    sizeEdits.resetColumnOverrides &&
      Object.keys(sheet.columnWidths ?? {}).length &&
      delete updatedSheet.columnWidths,
    updatedSheet.defaultRowHeight === sheet.defaultRowHeight &&
    updatedSheet.defaultColumnWidth === sheet.defaultColumnWidth &&
    updatedSheet.rowHeights === sheet.rowHeights &&
    updatedSheet.columnWidths === sheet.columnWidths
      ? workbook
      : {
          ...workbook,
          sheets: workbook.sheets.map((sheet) =>
            sheet.id === sheetId ? updatedSheet : sheet,
          ),
        }
  );
}
async function hashSheetPassword(password, salt) {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(salt + "\0" + password),
  );
  return [...new Uint8Array(digest)]
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}
async function updateSheetProtection(
  workbook,
  sheetId,
  password,
  isProtecting,
) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  (isProtecting &&
    sheet.protection?.enabled &&
    failWithMessage(
      "Unprotect this sheet before changing its protection settings.",
    ),
    password.length > 200 && failWithMessage("Password is too long."),
    !isProtecting &&
      sheet.protection?.hash &&
      (await hashSheetPassword(password, sheet.protection.salt)) !==
        sheet.protection.hash &&
      failWithMessage("The password does not match."));
  const salt = crypto.randomUUID(),
    protection = isProtecting
      ? {
          enabled: true,
          ...(password
            ? { salt: salt, hash: await hashSheetPassword(password, salt) }
            : {}),
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
function throwWorkbookError(message) {
  throw new Error(message);
}
const numericFormatKinds = new Set([
    "number",
    "currency",
    "accounting",
    "percent",
  ]),
  dateTimeFormatKinds = new Set(["date", "time", "datetime", "duration"]);
function ZQ(value) {
  return (
    !!value &&
    typeof value == "object" &&
    !Array.isArray(value) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(value))
  );
}
function BN(formatCode) {
  const sections = [""];
  let inQuotes = false;
  for (let charIndex = 0; charIndex < formatCode.length; charIndex++) {
    const character = formatCode[charIndex];
    if (character === "\\" && charIndex + 1 < formatCode.length) {
      sections[sections.length - 1] += character + formatCode[++charIndex];
      continue;
    }
    (character === '"' && (inQuotes = !inQuotes),
      character === ";" && !inQuotes
        ? sections.push("")
        : (sections[sections.length - 1] += character));
  }
  return inQuotes || sections.length > 4 ? undefined : sections;
}
function parseNumberFormatCode(sectionCode) {
  const parsedFormat = {
      prefix: "",
      suffix: "",
      minDecimals: 0,
      maxDecimals: 0,
      minIntegerDigits: 1,
      grouping: false,
      percent: false,
      hasNumber: false,
    },
    appendLiteral = (literal) => {
      parsedFormat.hasNumber
        ? (parsedFormat.suffix += literal)
        : (parsedFormat.prefix += literal);
    };
  for (let charIndex = 0; charIndex < sectionCode.length;) {
    const character = sectionCode[charIndex];
    if (character === '"') {
      const closingQuoteIndex = sectionCode.indexOf('"', charIndex + 1);
      if (closingQuoteIndex < 0) return;
      (appendLiteral(sectionCode.slice(charIndex + 1, closingQuoteIndex)),
        (charIndex = closingQuoteIndex + 1));
    } else if (character === "\\") {
      if (charIndex + 1 >= sectionCode.length) return;
      (appendLiteral(sectionCode[charIndex + 1]), (charIndex += 2));
    } else if (character === "[") {
      const colorMatch =
        /^\[(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\d+)\]/i.exec(
          sectionCode.slice(charIndex),
        );
      if (!colorMatch) return;
      charIndex += colorMatch[0].length;
    } else if (character === "0" || character === "#") {
      if (parsedFormat.hasNumber) return;
      const digitRun = /^[0#.,]+/.exec(sectionCode.slice(charIndex))[0],
        digitMatch = /^(0+|#+0*|#,##0+)(?:\.(0*#*))?$/.exec(digitRun);
      if (!digitMatch) return;
      const fractionDigits = digitMatch[2] ?? "";
      if (
        ((parsedFormat.minDecimals = [...fractionDigits].filter(
          (fractionDigit) => fractionDigit === "0",
        ).length),
        (parsedFormat.maxDecimals = fractionDigits.length),
        (parsedFormat.minIntegerDigits = [...digitMatch[1]].filter(
          (integerDigit) => integerDigit === "0",
        ).length),
        parsedFormat.maxDecimals > 20 || parsedFormat.minIntegerDigits > 21)
      )
        return;
      ((parsedFormat.grouping = digitMatch[1].includes(",")),
        (parsedFormat.hasNumber = true),
        (charIndex += digitRun.length));
    } else if (character === "%") {
      if (parsedFormat.percent) return;
      ((parsedFormat.percent = true), appendLiteral(character), charIndex++);
    } else if (/[$€£¥ ()+\-]/.test(character))
      (appendLiteral(character), charIndex++);
    else return;
  }
  return parsedFormat;
}
function summarizeNumberFormat(formatCode) {
  if (!formatCode) return;
  const sections = BN(formatCode);
  if (!sections) return;
  const numberFormatCode = parseNumberFormatCode(sections[0]);
  return numberFormatCode?.hasNumber ? numberFormatCode.maxDecimals : undefined;
}
function numberFormatToIntlOptions(value, numFmt) {
  const sections = BN(numFmt);
  if (!sections) return;
  const usesNegativeSection = value < 0 && sections.length > 1,
    sectionIndex =
      value === 0 && sections.length > 2 ? 2 : usesNegativeSection ? 1 : 0,
    numberFormatCode = parseNumberFormatCode(sections[sectionIndex]);
  if (!numberFormatCode) return;
  if (!numberFormatCode.hasNumber)
    return numberFormatCode.prefix + numberFormatCode.suffix;
  const magnitude = Math.abs(value) * (numberFormatCode.percent ? 100 : 1);
  if (!Number.isFinite(magnitude)) return "#NUM!";
  let formatted = new Intl.NumberFormat("en-US", {
    minimumFractionDigits: numberFormatCode.minDecimals,
    maximumFractionDigits: numberFormatCode.maxDecimals,
    minimumIntegerDigits: Math.max(1, numberFormatCode.minIntegerDigits),
    useGrouping: numberFormatCode.grouping,
  }).format(magnitude);
  return (
    numberFormatCode.minIntegerDigits === 0 &&
      (formatted === "0" && numberFormatCode.minDecimals === 0
        ? (formatted = "")
        : formatted.startsWith("0.") && (formatted = formatted.slice(1))),
    (value < 0 && !usesNegativeSection ? "-" : "") +
      numberFormatCode.prefix +
      formatted +
      numberFormatCode.suffix
  );
}
const $b = new Map();
function parseDateTimeFormatCode(formatCode) {
  const unquotedFormatText = formatCode.replace(/"[^"]*"|\\./g, "");
  if (!/[ymdhs]/i.test(unquotedFormatText)) return;
  const tokens = [];
  let fraction = 0;
  for (let scanIndex = 0; scanIndex < formatCode.length;) {
    const remainingText = formatCode.slice(scanIndex),
      currentChar = remainingText[0];
    if (currentChar === '"') {
      const closingQuoteIndex = formatCode.indexOf('"', scanIndex + 1);
      (closingQuoteIndex < 0 &&
        throwWorkbookError("The format contains an unclosed quoted literal."),
        tokens.push({
          kind: "literal",
          text: formatCode.slice(scanIndex + 1, closingQuoteIndex),
        }),
        (scanIndex = closingQuoteIndex + 1));
      continue;
    }
    if (currentChar === "\\") {
      (remainingText.length < 2 &&
        throwWorkbookError("A format escape must be followed by a character."),
        tokens.push({ kind: "literal", text: remainingText[1] }),
        (scanIndex += 2));
      continue;
    }
    const ampmMatch = /^(AM\/PM|A\/P)/i.exec(remainingText);
    if (ampmMatch) {
      (tokens.push({ kind: "ampm", text: ampmMatch[0] }),
        (scanIndex += ampmMatch[0].length));
      continue;
    }
    const elapsedMatch = /^\[(h{1,2}|m{1,2}|s{1,2})\]/i.exec(remainingText);
    if (elapsedMatch) {
      (tokens.push({ kind: "elapsed", text: elapsedMatch[1].toLowerCase() }),
        (scanIndex += elapsedMatch[0].length));
      continue;
    }
    const dateFieldMatch = /^(y+|m+|d+|h+|s+)/i.exec(remainingText);
    if (dateFieldMatch) {
      const fieldText = dateFieldMatch[0].toLowerCase(),
        fieldLetter = fieldText[0];
      (((fieldLetter === "y" && ![2, 4].includes(fieldText.length)) ||
        (fieldLetter === "m" && fieldText.length > 4) ||
        (["d", "h", "s"].includes(fieldLetter) && fieldText.length > 2)) &&
        throwWorkbookError(
          "This date pattern uses unsupported fields. Use yy/yyyy, m–mmmm, d/dd, h/hh or s/ss.",
        ),
        tokens.push({
          kind: { y: "year", m: "month", d: "day", h: "hour", s: "second" }[
            fieldLetter
          ],
          text: fieldText,
        }),
        (scanIndex += fieldText.length));
      continue;
    }
    const fractionMatch = /^\.(0{1,3})(?!0)/.exec(remainingText);
    if (fractionMatch) {
      ((!tokens.length ||
        !["second", "elapsed"].includes(tokens.at(-1).kind) ||
        (tokens.at(-1).kind === "elapsed" &&
          !tokens.at(-1).text.startsWith("s")) ||
        fraction) &&
        throwWorkbookError(
          "Fractional seconds must follow seconds and use 1–3 zero digits.",
        ),
        (fraction = fractionMatch[1].length),
        tokens.push({ kind: "fraction", text: fractionMatch[1] }),
        (scanIndex += fractionMatch[0].length));
      continue;
    }
    if (/^[ /:,.-]$/.test(currentChar)) {
      (tokens.push({ kind: "literal", text: currentChar }), scanIndex++);
      continue;
    }
    throwWorkbookError(
      "Unsupported date/time code. Quote literal words and use one supported date/time pattern.",
    );
  }
  const nonLiteralTokens = tokens.filter((token) => token.kind !== "literal");
  nonLiteralTokens.forEach((entry, entryIndex) => {
    entry.kind === "month" &&
      entry.text.length <= 2 &&
      (nonLiteralTokens[entryIndex - 1]?.kind === "hour" ||
        (nonLiteralTokens[entryIndex - 1]?.kind === "elapsed" &&
          nonLiteralTokens[entryIndex - 1].text.startsWith("h")) ||
        nonLiteralTokens[entryIndex + 1]?.kind === "second") &&
      (entry.kind = "minute");
  });
  const hasDateField = nonLiteralTokens.some((calendarToken) =>
      ["year", "month", "day"].includes(calendarToken.kind),
    ),
    hasTimeField = nonLiteralTokens.some((timeToken) =>
      ["hour", "minute", "second"].includes(timeToken.kind),
    ),
    elapsedTokens = nonLiteralTokens.filter(
      (elapsedToken) => elapsedToken.kind === "elapsed",
    );
  if (elapsedTokens.length) {
    (elapsedTokens.length !== 1 ||
      hasDateField ||
      nonLiteralTokens.some((hourOrAmPmToken) =>
        ["hour", "ampm"].includes(hourOrAmPmToken.kind),
      ) ||
      nonLiteralTokens[0].kind !== "elapsed") &&
      throwWorkbookError(
        "Elapsed-time patterns start with one [h], [m] or [s] field and cannot include calendar or AM/PM fields.",
      );
    const elapsedUnit = elapsedTokens[0].text[0];
    ((elapsedUnit === "s" && hasTimeField) ||
      (elapsedUnit === "m" &&
        nonLiteralTokens.some(
          (minuteToken) => minuteToken.kind === "minute",
        ))) &&
      throwWorkbookError("Use [h]:mm:ss, [m]:ss or [s] for elapsed time.");
  }
  return (
    nonLiteralTokens.some((ampmToken) => ampmToken.kind === "ampm") &&
      !nonLiteralTokens.some((hourToken) => hourToken.kind === "hour") &&
      throwWorkbookError("AM/PM requires an hour field."),
    !hasDateField &&
      !hasTimeField &&
      !elapsedTokens.length &&
      throwWorkbookError("The format does not contain a date or time field."),
    {
      kind: "temporal",
      category: elapsedTokens.length
        ? "duration"
        : hasDateField && hasTimeField
          ? "datetime"
          : hasDateField
            ? "date"
            : "time",
      tokens: tokens,
      fraction: fraction,
    }
  );
}
function parseCustomNumberFormat(formatCode) {
  let patternText = "",
    divisor = 1,
    hasNumericField = false;
  for (let scanIndex = 0; scanIndex < formatCode.length;) {
    const remainingText = formatCode.slice(scanIndex);
    if (remainingText[0] === '"') {
      const closingQuoteIndex = formatCode.indexOf('"', scanIndex + 1);
      (closingQuoteIndex < 0 &&
        throwWorkbookError("The format contains an unclosed quoted literal."),
        (patternText += formatCode.slice(scanIndex, closingQuoteIndex + 1)),
        (scanIndex = closingQuoteIndex + 1));
    } else if (remainingText[0] === "\\")
      (remainingText.length < 2 &&
        throwWorkbookError("A format escape must be followed by a character."),
        (patternText += remainingText.slice(0, 2)),
        (scanIndex += 2));
    else if (/^[0#]/.test(remainingText)) {
      hasNumericField &&
        throwWorkbookError(
          "Only one numeric field per custom section is supported.",
        );
      const numericFieldText = /^[0#.,]+/.exec(remainingText)[0],
        trailingCommas = /,+$/.exec(numericFieldText)?.[0] ?? "";
      (trailingCommas.length > 2 &&
        throwWorkbookError(
          "Custom display scaling supports at most two trailing commas.",
        ),
        (divisor = 1e3 ** trailingCommas.length),
        (patternText += numericFieldText.slice(
          0,
          numericFieldText.length - trailingCommas.length,
        )),
        (scanIndex += numericFieldText.length),
        (hasNumericField = true));
    } else ((patternText += remainingText[0]), scanIndex++);
  }
  /\[/.test(patternText.replace(/"[^"]*"|\\./g, "")) &&
    throwWorkbookError(
      "Custom colors, conditions and locale tags are not supported. Use conditional formatting or the locale selector.",
    );
  const numberFormatCode = parseNumberFormatCode(patternText);
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
function isTextOnlyFormatSection(formatSection) {
  return formatSection.replace(/"[^"]*"|\\./g, "").replace(/[@ ]/g, "") === "";
}
function validateNumberFormatCode(formatCode) {
  const cachedFormat = $b.get(formatCode);
  if (cachedFormat) return cachedFormat;
  (typeof formatCode != "string" ||
    !formatCode.length ||
    formatCode.length > 200 ||
    /[\u0000-\u001f\u007f]/.test(formatCode)) &&
    throwWorkbookError("Enter a supported format pattern of 1–200 characters.");
  let parsedFormat;
  const formatSections = BN(formatCode);
  formatSections ||
    throwWorkbookError(
      "A format has at most four sections and must close quoted literals.",
    );
  const dateTimeFormat =
    formatSections.length === 1
      ? parseDateTimeFormatCode(formatCode)
      : undefined;
  return (
    dateTimeFormat
      ? (parsedFormat = dateTimeFormat)
      : (formatSections.length === 4 &&
          !isTextOnlyFormatSection(formatSections[3]) &&
          throwWorkbookError(
            "The text section supports @ and quoted or escaped literals only.",
          ),
        (parsedFormat = {
          kind: "numeric",
          sections: formatSections.slice(0, 3).map(parseCustomNumberFormat),
          ...(formatSections.length === 4
            ? { textSection: formatSections[3] }
            : {}),
        })),
    $b.size >= 256 && $b.delete($b.keys().next().value),
    $b.set(formatCode, parsedFormat),
    parsedFormat
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
