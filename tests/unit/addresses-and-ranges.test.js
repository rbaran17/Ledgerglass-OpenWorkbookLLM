import { describe, expect, test } from "vitest";
import {
  _t as parseCellAddress,
  columnIndexToLetters,
  countSelectionCells,
  enumerateSelectionCells,
  isCellAddressInRanges,
  isCellInsideRange,
  mapIndexAfterMove,
  mapRangeAfterMove,
  parseRangeAddress,
  parseReportRange,
  parseSelectionAreas,
  rangesIntersect,
  resolveMovedRange,
  shiftFormulaReferences,
} from "../../src/app/03-workbook-model.js";

describe("column letters", () => {
  test("converts zero-based indexes to column letters", () => {
    const letters = [0, 1, 25, 26, 27, 51, 52, 701, 702, 16383].map(
      columnIndexToLetters,
    );
    expect(letters).toEqual([
      "A",
      "B",
      "Z",
      "AA",
      "AB",
      "AZ",
      "BA",
      "ZZ",
      "AAA",
      "XFD",
    ]);
  });

  test("rejects indexes outside the sheet and non-integers", () => {
    expect(() => columnIndexToLetters(16384)).toThrow();
    expect(() => columnIndexToLetters(-1)).toThrow();
    expect(() => columnIndexToLetters(1.5)).toThrow();
  });
});

describe("cell addresses", () => {
  test("parses A1 and C10 into zero-based row and column", () => {
    expect(parseCellAddress("A1")).toEqual({ row: 0, col: 0 });
    expect(parseCellAddress("C10")).toEqual({ row: 9, col: 2 });
  });

  test("accepts $ anchors and lowercase letters", () => {
    expect(parseCellAddress("$C$10")).toEqual({ row: 9, col: 2 });
    expect(parseCellAddress("c10")).toEqual({ row: 9, col: 2 });
  });

  test("accepts the last row and last column", () => {
    expect(parseCellAddress("XFD1048576")).toEqual({
      row: 1048575,
      col: 16383,
    });
  });

  test("rejects addresses beyond the sheet bounds", () => {
    expect(() => parseCellAddress("XFE1")).toThrow();
    expect(() => parseCellAddress("A1048577")).toThrow();
  });

  test("rejects malformed addresses", () => {
    for (const bad of ["A0", "1A", "A01", "AAAA1", "A", "", "A1:B2"]) {
      expect(() => parseCellAddress(bad)).toThrow();
    }
  });

  test("column letters and cell parsing round-trip", () => {
    for (const index of [0, 25, 26, 701, 702, 16383]) {
      const address = `${columnIndexToLetters(index)}7`;
      expect(parseCellAddress(address)).toEqual({ row: 6, col: index });
    }
  });
});

describe("range addresses", () => {
  test("parses B2:D5 into inclusive zero-based bounds", () => {
    expect(parseRangeAddress("B2:D5")).toEqual({
      firstRow: 1,
      lastRow: 4,
      firstCol: 1,
      lastCol: 3,
    });
  });

  test("normalises reversed corners so D5:B2 equals B2:D5", () => {
    expect(parseRangeAddress("D5:B2")).toEqual(parseRangeAddress("B2:D5"));
  });

  test("accepts $ anchors on either corner", () => {
    expect(parseRangeAddress("$B$2:D5")).toEqual(parseRangeAddress("B2:D5"));
  });

  test("rejects ranges with more than two corners", () => {
    expect(() => parseRangeAddress("A1:B2:C3")).toThrow();
  });

  test("report ranges are canonicalised without $ and in upper case", () => {
    expect(parseReportRange("$b$2:$d$5")).toBe("B2:D5");
  });

  test("report ranges reject reversed corners and oversized areas", () => {
    expect(() => parseReportRange("D5:B2")).toThrow();
    expect(() => parseReportRange("A1:A1048576")).toThrow();
  });
});

describe("containment and overlap", () => {
  const block = parseRangeAddress("B2:D5");

  test("isCellInsideRange includes the edges and excludes outside cells", () => {
    expect(isCellInsideRange(block, { row: 1, col: 1 })).toBe(true);
    expect(isCellInsideRange(block, { row: 4, col: 3 })).toBe(true);
    expect(isCellInsideRange(block, { row: 5, col: 3 })).toBe(false);
    expect(isCellInsideRange(block, { row: 1, col: 4 })).toBe(false);
  });

  test("rangesIntersect detects shared cells in either argument order", () => {
    const corner = parseRangeAddress("D5:F9");
    expect(rangesIntersect(block, corner)).toBe(true);
    expect(rangesIntersect(corner, block)).toBe(true);
  });

  test("rangesIntersect is false for ranges that only touch edges", () => {
    expect(
      rangesIntersect(parseRangeAddress("A1:A2"), parseRangeAddress("A3:A4")),
    ).toBe(false);
    expect(
      rangesIntersect(parseRangeAddress("B2:D5"), parseRangeAddress("E2:F9")),
    ).toBe(false);
  });

  test("isCellAddressInRanges matches a cell against any listed range", () => {
    const ranges = ["B2:D5", "F1"];
    expect(isCellAddressInRanges(ranges, "F1")).toBe(true);
    expect(isCellAddressInRanges(ranges, "E5")).toBe(false);
    expect(isCellAddressInRanges(["B2:D5"], "A1")).toBe(false);
  });
});

describe("selections", () => {
  test("parseSelectionAreas normalises, de-duplicates and collapses single cells", () => {
    expect(parseSelectionAreas("B2:D5, a1, A1:B2, b2")).toEqual([
      "B2:D5",
      "A1",
      "A1:B2",
      "B2",
    ]);
  });

  test("parseSelectionAreas rejects empty selections and too many areas", () => {
    expect(() => parseSelectionAreas("")).toThrow();
    expect(() => parseSelectionAreas(Array(257).fill("A1"))).toThrow();
  });

  test("countSelectionCells counts overlapping cells once", () => {
    expect(countSelectionCells("A1:B2,B2:C3")).toBe(7);
    expect(countSelectionCells("A1:A3,A2")).toBe(3);
  });

  test("enumerateSelectionCells lists every cell and enforces the limit", () => {
    expect(enumerateSelectionCells("A1:B2")).toEqual(["A1", "B1", "A2", "B2"]);
    expect(() => enumerateSelectionCells("A1:C3", 4)).toThrow();
  });
});

describe("shifting references", () => {
  test("relative references move with rows and absolute ones stay put", () => {
    expect(shiftFormulaReferences("=A1+$B$2+C3:C5", 1, 0)).toBe(
      "=A2+$B$2+C4:C6",
    );
  });

  test("relative references move with columns", () => {
    expect(shiftFormulaReferences("=A1+$B$2+C3:C5", 0, 2)).toBe(
      "=C1+$B$2+E3:E5",
    );
  });

  test("whole-row and whole-column references shift on their own axis", () => {
    expect(shiftFormulaReferences("=1:3", 2, 0)).toBe("=3:5");
    expect(shiftFormulaReferences("=A:B", 0, 1)).toBe("=B:C");
  });

  test("references that move off the sheet become #REF!", () => {
    expect(shiftFormulaReferences("=A1", -2, 0)).toBe("=#REF!");
    expect(shiftFormulaReferences("=A1048575", 1, 0)).toBe("=A1048576");
    expect(shiftFormulaReferences("=XFD1", 0, 1)).toBe("=#REF!");
    expect(shiftFormulaReferences("=A1048576", 1, 0)).toBe("=#REF!");
  });

  test("string literals, function names and plain text are left alone", () => {
    expect(shiftFormulaReferences('="A1"&A1', 1, 0)).toBe('="A1"&A2');
    expect(shiftFormulaReferences("=LOG10(A1)", 1, 0)).toBe("=LOG10(A2)");
    expect(shiftFormulaReferences("=A1<>B1", 1, 0)).toBe("=A2<>B2");
    expect(shiftFormulaReferences("A1", 1, 1)).toBe("A1");
  });
});

describe("move mapping", () => {
  // Moving the two-row block at index 2..3 to sit before row 8.
  const move = { index: 2, count: 2, moveBefore: 8 };

  test("mapIndexAfterMove follows the moved block and shifts its neighbours", () => {
    expect(mapIndexAfterMove(3, move)).toBe(7);
    expect(mapIndexAfterMove(5, move)).toBe(3);
    expect(mapIndexAfterMove(9, move)).toBe(9);
  });

  test("mapRangeAfterMove keeps a range that moves as a whole contiguous", () => {
    expect(mapRangeAfterMove(0, 9, move)).toEqual([{ start: 0, end: 9 }]);
  });

  test("resolveMovedRange refuses a range that the move would split", () => {
    expect(resolveMovedRange(0, 9, move)).toEqual({ start: 0, end: 9 });
    expect(() => resolveMovedRange(0, 3, move)).toThrow();
  });
});
