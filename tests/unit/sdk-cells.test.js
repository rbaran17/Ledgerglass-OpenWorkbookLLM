import { describe, expect, test } from "vitest";
import {
  applyAutoSum,
  applyBorderEditTransaction,
  applyNumberFormatTransaction,
  borderLineStyles,
  buildNumberFormatCode,
  cellStyleToCss,
  columnIndexToLetters,
  countSelectionCells,
  enumerateSelectionCells,
  findOccupiedRangeBounds,
  formatCellValueWithNumberFormat,
  formatOccupiedRangeAddress,
  isCellAddressInRanges,
  isFormulaCellValue,
  normalizeCellRecordInput,
  parseRangeAddress,
  parseSelectionAreas,
  quoteSheetName,
  rangesIntersect,
  setCellHyperlink,
  supportedCurrencyCodes,
  supportedNumberLocales,
  unquoteSheetName,
  validateCellStyle,
  validateNumberFormatSettings,
} from "../../src/sdk/cells.js";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
} from "../../src/app/03-workbook-model.js";

// Builds a workbook whose first sheet holds the given cell records, without going through a transaction.
function sheetWith(cells, sheetExtras = {}) {
  const workbook = createBlankWorkbook("Cells");
  return {
    ...workbook,
    sheets: [{ ...workbook.sheets[0], cells: { ...cells }, ...sheetExtras }],
  };
}

const idOf = (workbook) => workbook.sheets[0].id;

// Applies an AutoSum patch by hand, so the calculator can evaluate the result.
function withPatch(workbook, patch) {
  const sheet = workbook.sheets[0];
  return {
    ...workbook,
    sheets: [
      { ...sheet, cells: { ...sheet.cells, [patch.address]: patch.cell } },
    ],
  };
}

describe("parseRangeAddress", () => {
  test("a single address gives one zero-based cell", () => {
    expect(parseRangeAddress("A1")).toEqual({
      firstRow: 0,
      lastRow: 0,
      firstCol: 0,
      lastCol: 0,
    });
  });

  test("reversed corners are normalized and dollar signs are ignored", () => {
    expect(parseRangeAddress("$C$3:d1")).toEqual({
      firstRow: 0,
      lastRow: 2,
      firstCol: 2,
      lastCol: 3,
    });
    expect(parseRangeAddress("b5:a1")).toEqual({
      firstRow: 0,
      lastRow: 4,
      firstCol: 0,
      lastCol: 1,
    });
  });

  test("the last cell of the grid is accepted", () => {
    expect(parseRangeAddress("XFD1048576")).toEqual({
      firstRow: 1048575,
      lastRow: 1048575,
      firstCol: 16383,
      lastCol: 16383,
    });
  });

  test("throws for three parts, row 0, and a column past the grid", () => {
    expect(() => parseRangeAddress("A1:B2:C3")).toThrow(
      "Invalid range: A1:B2:C3",
    );
    expect(() => parseRangeAddress("A0")).toThrow("Invalid cell address: A0");
    expect(() => parseRangeAddress("XFE1")).toThrow(
      "exceeds the supported workbook bounds",
    );
  });
});

describe("parseSelectionAreas", () => {
  test("a comma-separated string is normalized, trimmed and de-duplicated", () => {
    expect(parseSelectionAreas("b2:a1, c3 ,a1,$D$4")).toEqual([
      "A1:B2",
      "C3",
      "A1",
      "D4",
    ]);
  });

  test("an array of areas is accepted as well", () => {
    expect(parseSelectionAreas(["A1:A1", "A2"])).toEqual(["A1", "A2"]);
  });

  test("throws for an empty string or an invalid area", () => {
    expect(() => parseSelectionAreas("")).toThrow();
    expect(() => parseSelectionAreas("A1:B2:C3")).toThrow("Invalid range");
  });
});

describe("columnIndexToLetters", () => {
  test("converts zero-based indexes to column letters", () => {
    expect(
      [0, 25, 26, 27, 701, 702, 16383].map((index) =>
        columnIndexToLetters(index),
      ),
    ).toEqual(["A", "Z", "AA", "AB", "ZZ", "AAA", "XFD"]);
  });

  test("throws for an index outside the grid", () => {
    expect(() => columnIndexToLetters(16384)).toThrow(
      "Column must be an integer between 0 and 16383.",
    );
    expect(() => columnIndexToLetters(-1)).toThrow("Column must be");
  });
});

describe("isCellAddressInRanges", () => {
  test("reports whether the active cell lies in one of the areas", () => {
    const areas = ["A1:B2", "D4"];
    expect(isCellAddressInRanges(areas, "B2")).toBe(true);
    expect(isCellAddressInRanges(areas, "D4")).toBe(true);
    expect(isCellAddressInRanges(areas, "C3")).toBe(false);
  });

  test("the active address is case-insensitive", () => {
    expect(isCellAddressInRanges(["d4"], "D4")).toBe(true);
  });

  test("throws when the active address is not an address", () => {
    expect(() => isCellAddressInRanges(["A1"], "1")).toThrow(
      "Invalid cell address: 1",
    );
  });
});

describe("countSelectionCells", () => {
  test("counts cells in a rectangle", () => {
    expect(countSelectionCells("A1:B2")).toBe(4);
  });

  test("counts a cell shared by two areas once", () => {
    expect(countSelectionCells("A1:B2,B2:C3")).toBe(7);
    expect(countSelectionCells(["A1", "a1"])).toBe(1);
  });

  test("counts a whole column of the grid", () => {
    expect(countSelectionCells("A1:A1048576")).toBe(1048576);
  });
});

describe("enumerateSelectionCells", () => {
  test("lists addresses area by area, row by row, without repeats", () => {
    expect(enumerateSelectionCells("B1:C2,A1")).toEqual([
      "B1",
      "C1",
      "B2",
      "C2",
      "A1",
    ]);
  });

  test("throws when the selection is larger than the limit", () => {
    expect(() => enumerateSelectionCells("A1:Z1000", 10)).toThrow(
      "Select at most 10 cells",
    );
  });

  test("the default limit is 500 cells", () => {
    expect(enumerateSelectionCells("A1:A500")).toHaveLength(500);
    expect(() => enumerateSelectionCells("A1:Z30")).toThrow(
      "Select at most 500 cells",
    );
  });

  test("throws for a limit that is not a positive integer", () => {
    expect(() => enumerateSelectionCells("A1", 0)).toThrow(
      "Invalid selection enumeration limit.",
    );
  });
});

describe("rangesIntersect", () => {
  test("ranges that share a cell intersect", () => {
    expect(
      rangesIntersect(parseRangeAddress("A1:B2"), parseRangeAddress("B2:C3")),
    ).toBe(true);
  });

  test("ranges that only touch by side do not intersect", () => {
    expect(
      rangesIntersect(parseRangeAddress("A1:B2"), parseRangeAddress("C1:C3")),
    ).toBe(false);
  });
});

describe("findOccupiedRangeBounds", () => {
  test("returns null for an empty sheet", () => {
    expect(findOccupiedRangeBounds(sheetWith({}).sheets[0])).toBeNull();
  });

  test("includes formatting-only cells unless told not to", () => {
    const sheet = sheetWith({
      A1: { raw: 1 },
      C5: { raw: null, style: { bold: true } },
    }).sheets[0];
    expect(findOccupiedRangeBounds(sheet)).toEqual({
      firstRow: 0,
      lastRow: 4,
      firstCol: 0,
      lastCol: 2,
    });
    expect(findOccupiedRangeBounds(sheet, false)).toEqual({
      firstRow: 0,
      lastRow: 0,
      firstCol: 0,
      lastCol: 0,
    });
  });

  test("an empty string does not count as a value", () => {
    const sheet = sheetWith({ B2: { raw: "" } }).sheets[0];
    expect(findOccupiedRangeBounds(sheet)).toBeNull();
  });
});

describe("formatOccupiedRangeAddress", () => {
  test("returns A1 for an empty sheet", () => {
    expect(formatOccupiedRangeAddress(sheetWith({}).sheets[0])).toBe("A1");
  });

  test("returns the occupied rectangle as text", () => {
    const sheet = sheetWith({
      A1: { raw: 1 },
      C5: { raw: null, style: { bold: true } },
    }).sheets[0];
    expect(formatOccupiedRangeAddress(sheet)).toBe("A1:C5");
    expect(formatOccupiedRangeAddress(sheet, false)).toBe("A1");
  });
});

describe("quoteSheetName and unquoteSheetName", () => {
  test("quoteSheetName always wraps the name in single quotes and doubles inner quotes", () => {
    expect(quoteSheetName("Sheet 1")).toBe("'Sheet 1'");
    expect(quoteSheetName("Plain")).toBe("'Plain'");
    expect(quoteSheetName("O'Brien")).toBe("'O''Brien'");
  });

  test("unquoteSheetName removes the quotes and restores inner quotes", () => {
    expect(unquoteSheetName("'Sheet 1'")).toBe("Sheet 1");
    expect(unquoteSheetName("'O''Brien'")).toBe("O'Brien");
  });

  test("unquoteSheetName leaves unquoted text alone and reverses quoteSheetName", () => {
    expect(unquoteSheetName("Plain")).toBe("Plain");
    expect(unquoteSheetName(quoteSheetName("it's 2"))).toBe("it's 2");
  });
});

describe("isFormulaCellValue", () => {
  test("a cell whose raw text starts with = is a formula", () => {
    expect(isFormulaCellValue({ raw: "=A1+1" })).toBe(true);
  });

  test("a formula stored as literal text is not a formula, and neither are values or missing cells", () => {
    expect(isFormulaCellValue({ raw: "=A1", literal: true })).toBe(false);
    expect(isFormulaCellValue({ raw: "text" })).toBe(false);
    expect(isFormulaCellValue(undefined)).toBe(false);
  });
});

describe("normalizeCellRecordInput", () => {
  test("keeps a valid record and drops fields it does not know", () => {
    expect(normalizeCellRecordInput({ raw: 5 })).toEqual({ raw: 5 });
    expect(normalizeCellRecordInput({ raw: 1, bogus: 2 })).toEqual({ raw: 1 });
    expect(normalizeCellRecordInput({ raw: "=A1", literal: true })).toEqual({
      raw: "=A1",
      literal: true,
    });
  });

  test("fills in the defaults of a number format", () => {
    expect(
      normalizeCellRecordInput({
        raw: 1,
        numberFormat: { category: "percent", decimals: 1 },
      }).numberFormat,
    ).toEqual({
      category: "percent",
      locale: "en-US",
      decimals: 1,
      useGrouping: false,
      scale: 1,
      negative: "minus",
    });
  });

  test("throws when raw is missing or not a finite value", () => {
    expect(() => normalizeCellRecordInput({})).toThrow(
      "A cell must have a raw value.",
    );
    expect(() => normalizeCellRecordInput({ raw: NaN })).toThrow(
      "Cell values must be text, finite numbers, booleans, or null.",
    );
  });
});

describe("validateCellStyle", () => {
  test("keeps supported settings and returns them", () => {
    expect(
      validateCellStyle({ bold: true, fill: "#ffcc00", fontSize: 11 }),
    ).toEqual({ bold: true, fill: "#ffcc00", fontSize: 11 });
  });

  test("throws for an unknown setting", () => {
    expect(() => validateCellStyle({ sparkle: true })).toThrow(
      "Unsupported cell style setting: sparkle.",
    );
  });

  test("throws for a color that is not hexadecimal and a font size out of range", () => {
    expect(() => validateCellStyle({ fill: "red" })).toThrow(
      "Style colors must use hexadecimal colors.",
    );
    expect(() => validateCellStyle({ fontSize: 7 })).toThrow(
      "Font size must be an integer between 8 and 72.",
    );
  });

  test("a border edge gets the default black color when none is given", () => {
    expect(validateCellStyle({ borders: { left: { style: "thin" } } })).toEqual(
      { borders: { left: { style: "thin", color: "#000000" } } },
    );
  });

  test("throws for an unknown border edge or indent out of range", () => {
    expect(() =>
      validateCellStyle({ borders: { diagonal: { style: "thin" } } }),
    ).toThrow("Cell borders support top, right, bottom and left edges only.");
    expect(() => validateCellStyle({ indent: 16 })).toThrow(
      "Indent must be an integer between 0 and 15.",
    );
  });
});

describe("cellStyleToCss", () => {
  test("text with bold, underline and centered alignment", () => {
    const css = cellStyleToCss(
      { bold: true, underline: true, align: "center", fill: "#ffffff" },
      "hi",
    );
    expect(css).toMatchObject({
      background: "#ffffff",
      fontWeight: 700,
      textDecoration: "underline",
      textAlign: "center",
    });
  });

  test("a number is right aligned and indent is converted to padding", () => {
    expect(cellStyleToCss({ indent: 2, fontSize: 12 }, 5)).toMatchObject({
      textAlign: "right",
      fontSize: "12pt",
      paddingLeft: 3,
      paddingRight: 27,
    });
  });

  test("the border shorthand uses the default for every edge", () => {
    expect(
      cellStyleToCss({ border: "all" }, 1, { defaultBorder: "1px solid #ddd" })
        .borderTop,
    ).toBe("1px solid #000000");
    const css = cellStyleToCss(undefined, null, {
      defaultBorder: "1px solid #ddd",
    });
    expect(css.borderLeft).toBe("1px solid #ddd");
    expect(css.textAlign).toBe("left");
  });

  test("a double edge is three pixels wide and keeps its color", () => {
    const css = cellStyleToCss(
      { borders: { top: { style: "double", color: "#ff0000" } } },
      1,
    );
    expect(css.borderTop).toBe("3px double #ff0000");
    expect(css.borderRight).toBeUndefined();
  });
});

describe("applyBorderEditTransaction", () => {
  test("the all preset puts four edges on each cell of the area", () => {
    const workbook = sheetWith({
      A1: { raw: 10 },
      A2: { raw: 20 },
      A3: { raw: 30 },
    });
    const next = applyBorderEditTransaction(
      workbook,
      idOf(workbook),
      "A1:A2",
      {},
      { borders: { preset: "all" } },
    );
    const edges = next.sheets[0].cells.A1.style.borders;
    expect(Object.keys(edges).sort()).toEqual([
      "bottom",
      "left",
      "right",
      "top",
    ]);
    expect(next.sheets[0].cells.A2.style.borders.top).toEqual({
      style: "thin",
      color: "#000000",
    });
    expect(next.sheets[0].cells.A3).toEqual({ raw: 30 });
  });

  test("the outline preset gives the top-left cell its top and left edges only", () => {
    const workbook = sheetWith({});
    const next = applyBorderEditTransaction(
      workbook,
      idOf(workbook),
      "A1:B2",
      {},
      { borders: { preset: "outline" } },
    );
    expect(Object.keys(next.sheets[0].cells.A1.style.borders).sort()).toEqual([
      "left",
      "top",
    ]);
    expect(Object.keys(next.sheets[0].cells.B2.style.borders).sort()).toEqual([
      "bottom",
      "right",
    ]);
  });

  test("a style patch applies to the existing raw value and keeps it", () => {
    const workbook = sheetWith({ A1: { raw: 10 } });
    const next = applyBorderEditTransaction(workbook, idOf(workbook), "A1", {
      bold: true,
    });
    expect(next.sheets[0].cells.A1).toEqual({ raw: 10, style: { bold: true } });
  });

  test("the argument is not changed, and the revision and history stay as they were", () => {
    const workbook = sheetWith({ A1: { raw: 10 } });
    const next = applyBorderEditTransaction(workbook, idOf(workbook), "A1", {
      italic: true,
    });
    expect(workbook.sheets[0].cells.A1).toEqual({ raw: 10 });
    expect(next.revision).toBe(workbook.revision);
    expect(next.events).toHaveLength(workbook.events.length);
  });

  test("a second identical edit returns the same workbook object", () => {
    const workbook = sheetWith({});
    const first = applyBorderEditTransaction(workbook, idOf(workbook), "A1", {
      bold: true,
    });
    expect(
      applyBorderEditTransaction(first, idOf(first), "A1", { bold: true }),
    ).toBe(first);
  });

  test("an array selects several areas, and a comma string is treated as one area and throws", () => {
    const workbook = sheetWith({});
    const next = applyBorderEditTransaction(
      workbook,
      idOf(workbook),
      ["A1", "C3:C4"],
      { italic: true },
    );
    expect(Object.keys(next.sheets[0].cells).sort()).toEqual([
      "A1",
      "C3",
      "C4",
    ]);
    expect(() =>
      applyBorderEditTransaction(workbook, idOf(workbook), "A1:A2,C3", {
        bold: true,
      }),
    ).toThrow("Invalid cell address: A2,C3");
  });

  test("throws for a read-only workbook, a selection over 500 cells, and an intersecting merge", () => {
    const workbook = sheetWith({});
    expect(() =>
      applyBorderEditTransaction(
        workbook,
        idOf(workbook),
        "A1",
        {},
        {
          readOnly: true,
          borders: { preset: "all" },
        },
      ),
    ).toThrow("This workbook is read-only.");
    expect(() =>
      applyBorderEditTransaction(workbook, idOf(workbook), "A1:Z30", {
        bold: true,
      }),
    ).toThrow("Select at most 500 cells");
    const merged = sheetWith({}, { merges: ["A1:B2"] });
    expect(() =>
      applyBorderEditTransaction(merged, idOf(merged), "A1", {
        borders: { top: { style: "thin" } },
      }),
    ).toThrow("Unmerge the intersecting cells");
  });
});

describe("setCellHyperlink", () => {
  test("stores an external link and keeps the cell value", () => {
    const workbook = sheetWith({ A1: { raw: 10 } });
    const next = setCellHyperlink(workbook, idOf(workbook), "A1", {
      kind: "external",
      url: "https://example.com/docs",
    });
    expect(next.sheets[0].cells.A1).toEqual({
      raw: 10,
      hyperlink: { kind: "external", url: "https://example.com/docs" },
    });
  });

  test("normalizes the address, so a bare host gets a trailing slash", () => {
    const workbook = sheetWith({});
    const next = setCellHyperlink(workbook, idOf(workbook), "a1", {
      kind: "external",
      url: "https://example.com",
    });
    expect(next.sheets[0].cells.A1.hyperlink.url).toBe("https://example.com/");
  });

  test("stores an internal link to a cell and keeps its label", () => {
    const workbook = sheetWith({});
    const sheetId = idOf(workbook);
    const next = setCellHyperlink(workbook, sheetId, "C5", {
      kind: "internal",
      sheetId,
      address: "A1",
      label: "Go",
    });
    expect(next.sheets[0].cells.C5.hyperlink).toEqual({
      kind: "internal",
      sheetId,
      address: "A1",
      label: "Go",
    });
  });

  test("passing null removes the link and an empty cell disappears", () => {
    const workbook = sheetWith({});
    const linked = setCellHyperlink(workbook, idOf(workbook), "A1", {
      kind: "external",
      url: "https://example.com",
    });
    const cleared = setCellHyperlink(linked, idOf(linked), "A1", null);
    expect(cleared.sheets[0].cells.A1).toBeUndefined();
  });

  test("throws for a script link, a missing sheet, a range, and a read-only workbook", () => {
    const workbook = sheetWith({});
    const sheetId = idOf(workbook);
    expect(() =>
      setCellHyperlink(workbook, sheetId, "A1", {
        kind: "external",
        url: "javascript:alert(1)",
      }),
    ).toThrow("Only https://, http:// and mailto: links are supported.");
    expect(() =>
      setCellHyperlink(workbook, sheetId, "A1", {
        kind: "internal",
        sheetId: "nope",
        address: "A1",
      }),
    ).toThrow("Sheet not found: nope");
    expect(() =>
      setCellHyperlink(workbook, sheetId, "A1:B2", {
        kind: "external",
        url: "https://example.com",
      }),
    ).toThrow("Invalid cell address: A1:B2");
    expect(() =>
      setCellHyperlink(
        workbook,
        sheetId,
        "A1",
        { kind: "external", url: "https://example.com" },
        { readOnly: true },
      ),
    ).toThrow("This workbook is read-only.");
  });

  test("the returned workbook keeps the revision and history", () => {
    const workbook = sheetWith({});
    const next = setCellHyperlink(workbook, idOf(workbook), "A1", {
      kind: "external",
      url: "https://example.com",
    });
    expect(next.revision).toBe(workbook.revision);
    expect(next.events).toHaveLength(workbook.events.length);
  });
});

describe("validateNumberFormatSettings", () => {
  test("a number gets two decimals, grouping, unit scale and minus signs by default", () => {
    expect(validateNumberFormatSettings({ category: "number" })).toEqual({
      category: "number",
      locale: "en-US",
      decimals: 2,
      useGrouping: true,
      scale: 1,
      negative: "minus",
    });
  });

  test("yen defaults to no decimal places", () => {
    expect(
      validateNumberFormatSettings({ category: "currency", currency: "JPY" })
        .decimals,
    ).toBe(0);
  });

  test("a general format keeps only its category and locale", () => {
    expect(validateNumberFormatSettings({ category: "general" })).toEqual({
      category: "general",
      locale: "en-US",
    });
  });

  test("throws for an unsupported currency, a scale on a percent, and a locale that is not listed", () => {
    expect(() =>
      validateNumberFormatSettings({ category: "currency", currency: "XYZ" }),
    ).toThrow("Choose a supported currency.");
    expect(() =>
      validateNumberFormatSettings({ category: "percent", scale: 1000 }),
    ).toThrow("Choose units, thousands or millions; percentages use units.");
    expect(() =>
      validateNumberFormatSettings({ category: "number", locale: "xx-XX" }),
    ).toThrow("Choose one of the supported display locales.");
  });

  test("throws when a date category gets a number pattern", () => {
    expect(() =>
      validateNumberFormatSettings({ category: "date", pattern: "0.00" }),
    ).toThrow("Choose a date, time or elapsed-time pattern for this category.");
  });
});

describe("buildNumberFormatCode", () => {
  test("builds Excel codes for number, percent and general", () => {
    expect(buildNumberFormatCode({ category: "number", decimals: 2 })).toBe(
      "#,##0.00",
    );
    expect(buildNumberFormatCode({ category: "percent", decimals: 1 })).toBe(
      "0.0%",
    );
    expect(buildNumberFormatCode({ category: "general" })).toBe("General");
  });

  test("currency puts the symbol in quotes, and thousands scale adds a K suffix", () => {
    expect(
      buildNumberFormatCode({
        category: "currency",
        currency: "USD",
        decimals: 2,
      }),
    ).toBe('"$"#,##0.00');
    expect(
      buildNumberFormatCode({ category: "number", decimals: 0, scale: 1000 }),
    ).toBe('#,##0," K"');
  });

  test("a non-US locale adds its LCID prefix", () => {
    expect(
      buildNumberFormatCode({
        category: "number",
        locale: "de-DE",
        decimals: 2,
      }),
    ).toBe("[$-407]#,##0.00");
  });

  test("throws for invalid settings", () => {
    expect(() =>
      buildNumberFormatCode({ category: "number", currency: "USD" }),
    ).toThrow("Currency applies only to Currency or Accounting formats.");
  });
});

describe("formatCellValueWithNumberFormat", () => {
  test("formats numbers with grouping and fixed decimals", () => {
    expect(
      formatCellValueWithNumberFormat(1234.5, {
        category: "number",
        decimals: 2,
      }).text,
    ).toBe("1,234.50");
  });

  test("a negative currency value with the minus style puts the sign before the symbol", () => {
    expect(
      formatCellValueWithNumberFormat(-1234.5, {
        category: "currency",
        currency: "EUR",
        decimals: 2,
      }).text,
    ).toBe("-€1,234.50");
  });

  test("parentheses style wraps negatives", () => {
    expect(
      formatCellValueWithNumberFormat(-5, {
        category: "currency",
        currency: "USD",
        decimals: 2,
        negative: "parentheses",
      }).text,
    ).toBe("($5.00)");
  });

  test("accounting format returns its parts as well as the text", () => {
    const result = formatCellValueWithNumberFormat(-1234.5, {
      category: "accounting",
      currency: "USD",
      decimals: 2,
    });
    expect(result.text).toBe("$ (1,234.50)");
    expect(result.accounting).toMatchObject({
      symbol: "$",
      symbolPosition: "start",
    });
  });

  test("percent multiplies by 100, and a million scale adds an M suffix", () => {
    expect(
      formatCellValueWithNumberFormat(0.256, {
        category: "percent",
        decimals: 1,
      }).text,
    ).toBe("25.6%");
    expect(
      formatCellValueWithNumberFormat(1234567, {
        category: "number",
        decimals: 1,
        scale: 1e6,
      }).text,
    ).toBe("1.2 M");
  });

  test("dates count from 1900 by default and from 1904 when asked", () => {
    const date = { category: "date", pattern: "yyyy-mm-dd" };
    expect(formatCellValueWithNumberFormat(45000, date).text).toBe(
      "2023-03-15",
    );
    expect(formatCellValueWithNumberFormat(0, date, "1904").text).toBe(
      "1904-01-01",
    );
  });

  test("null, booleans and NaN have their own display text", () => {
    const settings = { category: "number", decimals: 2 };
    expect(formatCellValueWithNumberFormat(null, settings).text).toBe("");
    expect(formatCellValueWithNumberFormat(true, settings).text).toBe("TRUE");
    expect(formatCellValueWithNumberFormat(NaN, settings).text).toBe("#NUM!");
  });

  test("a custom pattern picks its section by sign, and text passes through a numeric-only pattern", () => {
    const custom = { category: "custom", pattern: '0;"neg "0;"zero"' };
    expect(formatCellValueWithNumberFormat(-3, custom).text).toBe("neg 3");
    expect(formatCellValueWithNumberFormat(0, custom).text).toBe("zero");
    expect(
      formatCellValueWithNumberFormat("abc", {
        category: "custom",
        pattern: "0.00",
      }).text,
    ).toBe("abc");
  });

  test("throws for invalid settings", () => {
    expect(() =>
      formatCellValueWithNumberFormat(1, { category: "general", pattern: "0" }),
    ).toThrow("General does not accept a custom pattern.");
  });
});

describe("applyNumberFormatTransaction", () => {
  test("applies the format to each cell of the area and keeps the values", () => {
    const workbook = sheetWith({ A1: { raw: 10 }, A2: { raw: 20 } });
    const next = applyNumberFormatTransaction(
      workbook,
      idOf(workbook),
      "A1:A2",
      { category: "currency", currency: "EUR", decimals: 2 },
    );
    expect(next.sheets[0].cells.A1).toMatchObject({
      raw: 10,
      numFmt: '"€"#,##0.00',
      format: "currency",
      numberFormat: { currency: "EUR", decimals: 2 },
    });
    expect(next.sheets[0].cells.A2.raw).toBe(20);
  });

  test("an array of areas is accepted and percent sets the percent format", () => {
    const workbook = sheetWith({});
    const next = applyNumberFormatTransaction(
      workbook,
      idOf(workbook),
      ["A1", "B2"],
      { category: "percent", decimals: 1 },
    );
    expect(next.sheets[0].cells.A1.format).toBe("percent");
    expect(next.sheets[0].cells.B2.numFmt).toBe("0.0%");
  });

  test("applying the same format again returns the same workbook object", () => {
    const workbook = sheetWith({ A1: { raw: 1 } });
    const settings = { category: "number", decimals: 0 };
    const first = applyNumberFormatTransaction(
      workbook,
      idOf(workbook),
      "A1",
      settings,
    );
    expect(
      applyNumberFormatTransaction(first, idOf(first), "A1", settings),
    ).toBe(first);
  });

  test("the argument is not changed, and the revision and history stay as they were", () => {
    const workbook = sheetWith({ A1: { raw: 1 } });
    const next = applyNumberFormatTransaction(workbook, idOf(workbook), "A1", {
      category: "number",
    });
    expect(workbook.sheets[0].cells.A1).toEqual({ raw: 1 });
    expect(next.revision).toBe(workbook.revision);
    expect(next.events).toHaveLength(workbook.events.length);
  });

  test("throws for invalid settings, a read-only workbook, and a selection over 500 cells", () => {
    const workbook = sheetWith({});
    expect(() =>
      applyNumberFormatTransaction(workbook, idOf(workbook), "A1", {
        category: "number",
        currency: "USD",
      }),
    ).toThrow("Currency applies only");
    expect(() =>
      applyNumberFormatTransaction(
        workbook,
        idOf(workbook),
        "A1",
        { category: "number" },
        { readOnly: true },
      ),
    ).toThrow("This workbook is read-only.");
    expect(() =>
      applyNumberFormatTransaction(workbook, idOf(workbook), "A1:Z30", {
        category: "number",
      }),
    ).toThrow("Select at most 500 cells");
  });
});

describe("applyAutoSum", () => {
  test("sums the numbers above an empty cell and the calculator agrees", () => {
    const workbook = sheetWith({
      A1: { raw: 10 },
      A2: { raw: 20 },
      A3: { raw: 30 },
    });
    const patch = applyAutoSum(workbook, idOf(workbook), "A4");
    expect(patch).toEqual({
      sheetId: idOf(workbook),
      address: "A4",
      cell: { raw: "=SUM(A1:A3)" },
    });
    const calculated = withPatch(workbook, patch);
    const { calculator } = getWorkbookCalculation(calculated);
    expect(calculator.getTypedValue(idOf(workbook), "A4")).toEqual({
      type: "number",
      value: 60,
    });
  });

  test("sums the numbers to the left when nothing is above", () => {
    const workbook = sheetWith({ A1: { raw: 10 }, B1: { raw: 1 } });
    expect(applyAutoSum(workbook, idOf(workbook), "C1").cell).toEqual({
      raw: "=SUM(A1:B1)",
    });
  });

  test("a formula above counts as a number", () => {
    const workbook = sheetWith({ A1: { raw: "=2+3" } });
    expect(applyAutoSum(workbook, idOf(workbook), "A2").cell.raw).toBe(
      "=SUM(A1:A1)",
    );
  });

  test("the argument is not changed", () => {
    const workbook = sheetWith({ A1: { raw: 10 } });
    applyAutoSum(workbook, idOf(workbook), "A2");
    expect(workbook.sheets[0].cells).toEqual({ A1: { raw: 10 } });
  });

  test("throws for a filled target, several targets, or no numbers next to it", () => {
    const workbook = sheetWith({ A1: { raw: 10 }, A2: { raw: 20 } });
    expect(() => applyAutoSum(workbook, idOf(workbook), "A1")).toThrow(
      "Select an empty cell for AutoSum.",
    );
    expect(() => applyAutoSum(workbook, idOf(workbook), "A3:A4")).toThrow(
      "Select one empty total cell.",
    );
    expect(() => applyAutoSum(workbook, idOf(workbook), "D9")).toThrow(
      "AutoSum needs adjacent numbers or formulas above or to the left.",
    );
  });
});

describe("constants", () => {
  test("borderLineStyles lists the line styles in order", () => {
    expect([...borderLineStyles]).toEqual([
      "none",
      "thin",
      "medium",
      "thick",
      "dotted",
      "dashed",
      "double",
    ]);
  });

  test("supportedCurrencyCodes and supportedNumberLocales list what the settings accept", () => {
    expect(supportedCurrencyCodes).toContain("USD");
    expect(supportedCurrencyCodes).toHaveLength(8);
    expect([...supportedNumberLocales]).toEqual([
      "en-US",
      "en-GB",
      "de-DE",
      "fr-FR",
      "ja-JP",
      "en-IN",
    ]);
  });
});
