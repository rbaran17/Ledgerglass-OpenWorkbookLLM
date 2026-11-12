import { describe, expect, test } from "vitest";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
} from "../../src/app/03-workbook-model.js";
import {
  addSavedFilterView,
  applyColumnTypePreview,
  applyFilterConditions,
  applyPivotViewChange,
  applyTableOperation,
  applyTableOperationPreview,
  buildFilterFacetValues,
  buildPivotChartModel,
  buildPivotDrillDown,
  buildPivotTableOutput,
  createSavedFilterView,
  createWorksheetTable,
  extractPivotDrillThrough,
  findTableForRange,
  getPivotRefreshState,
  getPivotSlicerItems,
  getPivotSourceFields,
  pivotAggregateOperations,
  previewColumnTypeChange,
  refreshPivotTable,
  removeSavedFilterView,
  rewriteStructuredReferences,
  setTableColumnDataType,
  suggestTableHeaderName,
  tableColumnDataTypes,
} from "../../src/sdk/tables.js";

/** Region, Units, Price over three data rows: North 10, South 4, North 7. */
function salesBook() {
  const workbook = createBlankWorkbook("Sales");
  const cells = workbook.sheets[0].cells;
  [
    ["Region", "Units", "Price"],
    ["North", 10, 2.5],
    ["South", 4, 3],
    ["North", 7, 1],
  ].forEach((row, rowIndex) =>
    row.forEach((raw, columnIndex) => {
      cells[`${"ABC"[columnIndex]}${rowIndex + 1}`] = { raw };
    }),
  );
  return workbook;
}

/** The sales book with a "Sales" table over A1:C4. */
function tableBook() {
  const book = salesBook();
  const workbook = createWorksheetTable(book, book.sheets[0].id, {
    name: "Sales",
    range: "A1:C4",
  });
  return { workbook, sheetId: workbook.sheets[0].id, table: workbook.sheets[0].tables[0] };
}

/** The sales book with a pivot of Units by Region, written to a new sheet. */
function pivotBook(extra = {}) {
  const book = salesBook();
  const sheetId = book.sheets[0].id;
  const workbook = buildPivotTableOutput(book, {
    name: "By region",
    source: { sheetId, range: "A1:C4" },
    rowFields: ["Region"],
    columnFields: [],
    values: [{ field: "Units", operation: "sum" }],
    grandTotals: true,
    ...extra,
  });
  const pivot = workbook.sheets[1].pivots[0];
  return { workbook, sheetId, pivotId: pivot.id };
}

const valueOf = (workbook, sheetId, address) =>
  getWorkbookCalculation(workbook).calculator.getTypedValue(sheetId, address);

describe("createWorksheetTable", () => {
  test("adds a table over the range, using the header row as column names", () => {
    const book = salesBook();
    const workbook = createWorksheetTable(book, book.sheets[0].id, {
      name: "Sales",
      range: "A1:C4",
    });
    const table = workbook.sheets[0].tables[0];
    expect(table).toMatchObject({
      name: "Sales",
      range: "A1:C4",
      style: "blue",
      showTotals: false,
    });
    expect(table.columns.map((column) => column.name)).toEqual([
      "Region",
      "Units",
      "Price",
    ]);
    expect(book.sheets[0].tables).toBeUndefined();
  });

  test("uses the reviewed headers when they are given", () => {
    const book = salesBook();
    const workbook = createWorksheetTable(book, book.sheets[0].id, {
      name: "Sales",
      range: "A1:C4",
      headers: ["Area", "Count", "Unit price"],
    });
    expect(workbook.sheets[0].tables[0].columns.map((c) => c.name)).toEqual([
      "Area",
      "Count",
      "Unit price",
    ]);
  });

  test("a totals row extends the table by one row", () => {
    const book = salesBook();
    const workbook = createWorksheetTable(book, book.sheets[0].id, {
      name: "Sales",
      range: "A1:C4",
      showTotals: true,
      totals: {},
    });
    expect(workbook.sheets[0].tables[0]).toMatchObject({
      range: "A1:C5",
      showTotals: true,
    });
  });

  test("refuses a name that is already used, ignoring case", () => {
    const { workbook, sheetId } = tableBook();
    const other = createBlankWorkbook("Other");
    other.sheets[0].cells.E1 = { raw: "x" };
    expect(() =>
      createWorksheetTable(workbook, sheetId, { name: "sales", range: "E1:E1", headers: ["x"] }),
    ).toThrow("A table with that name already exists.");
  });

  test("refuses a name that is not a valid table name", () => {
    const book = salesBook();
    expect(() =>
      createWorksheetTable(book, book.sheets[0].id, { name: "bad name!", range: "A1:C4" }),
    ).toThrow("Use a table name beginning with a letter or underscore");
  });

  test("refuses blank headers until they are reviewed", () => {
    const book = salesBook();
    book.sheets[0].cells.A1 = { raw: "" };
    expect(() =>
      createWorksheetTable(book, book.sheets[0].id, { name: "Sales", range: "A1:C4" }),
    ).toThrow("Review table headers");
  });
});

describe("applyTableOperation and applyTableOperationPreview", () => {
  test("appending rows previews a table that is one row longer", () => {
    const { workbook, sheetId, table } = tableBook();
    const plan = applyTableOperation(workbook, sheetId, table.id, {
      kind: "appendRows",
      rows: [["East", 5, 2]],
    });
    expect(plan.beforeRange).toBe("A1:C4");
    expect(plan.afterRange).toBe("A1:C5");
    const next = applyTableOperationPreview(workbook, plan);
    expect(next.sheets[0].tables[0].range).toBe("A1:C5");
    expect(next.sheets[0].cells.A5).toMatchObject({ raw: "East" });
    expect(workbook.sheets[0].tables[0].range).toBe("A1:C4");
  });

  test("a row count appends empty rows", () => {
    const { workbook, sheetId, table } = tableBook();
    const plan = applyTableOperation(workbook, sheetId, table.id, { kind: "appendRows", count: 2 });
    expect(applyTableOperationPreview(workbook, plan).sheets[0].tables[0].range).toBe("A1:C6");
  });

  test("appending both rows and a count is refused", () => {
    const { workbook, sheetId, table } = tableBook();
    expect(() =>
      applyTableOperation(workbook, sheetId, table.id, {
        kind: "appendRows",
        rows: [["East", 5, 2]],
        count: 1,
      }),
    ).toThrow("Specify raw rows or a row count, not both.");
  });

  test("a formula on one column is written to every data row", () => {
    const { workbook, sheetId, table } = tableBook();
    const unitsId = table.columns[1].id;
    const plan = applyTableOperation(workbook, sheetId, table.id, {
      kind: "setColumnFormula",
      columnId: unitsId,
      formula: "=[@Price]*2",
    });
    const next = applyTableOperationPreview(workbook, plan);
    expect(next.sheets[0].tables[0].columns[1].calculatedFormula).toBe("=[@Price]*2");
    expect(valueOf(next, sheetId, "B2")).toEqual({ type: "number", value: 5 });
    expect(valueOf(next, sheetId, "B4")).toEqual({ type: "number", value: 2 });
  });

  test("clearing a formula removes it from the column", () => {
    const { workbook, sheetId, table } = tableBook();
    const withFormula = applyTableOperationPreview(
      workbook,
      applyTableOperation(workbook, sheetId, table.id, {
        kind: "setColumnFormula",
        columnId: table.columns[2].id,
        formula: "=[@Units]+1",
      }),
    );
    const cleared = applyTableOperationPreview(
      withFormula,
      applyTableOperation(withFormula, sheetId, table.id, {
        kind: "clearColumnFormula",
        columnId: table.columns[2].id,
      }),
    );
    expect(cleared.sheets[0].tables[0].columns[2].calculatedFormula).toBeUndefined();
  });

  test("a formula that does not begin with = is refused", () => {
    const { workbook, sheetId, table } = tableBook();
    expect(() =>
      applyTableOperation(workbook, sheetId, table.id, {
        kind: "setColumnFormula",
        columnId: table.columns[2].id,
        formula: "not a formula",
      }),
    ).toThrow("Calculated column formulas must begin with =");
  });

  test("an unknown operation field is refused", () => {
    const { workbook, sheetId, table } = tableBook();
    expect(() =>
      applyTableOperation(workbook, sheetId, table.id, {
        kind: "clearColumnFormula",
        columnId: table.columns[2].id,
        formula: "=1",
      }),
    ).toThrow("Unknown calculated-table operation field.");
  });

  test("an unknown table is refused", () => {
    const { workbook, sheetId } = tableBook();
    expect(() =>
      applyTableOperation(workbook, sheetId, "missing", { kind: "resize", range: "A1:C6" }),
    ).toThrow("Select an existing table.");
  });

  test("a read-only workbook is refused", () => {
    const { workbook, sheetId, table } = tableBook();
    expect(() =>
      applyTableOperation(
        workbook,
        sheetId,
        table.id,
        { kind: "resize", range: "A1:C6" },
        { readOnly: true },
      ),
    ).toThrow("This workbook is read-only.");
  });

  test("a preview is refused when the workbook has changed since it was made", () => {
    const { workbook, sheetId, table } = tableBook();
    const plan = applyTableOperation(workbook, sheetId, table.id, {
      kind: "resize",
      range: "A1:C6",
    });
    const changed = structuredClone(workbook);
    changed.sheets[0].cells.E1 = { raw: "x" };
    expect(() => applyTableOperationPreview(changed, plan)).toThrow(
      "The workbook changed after the table preview.",
    );
  });

  test("a preview that changes nothing returns the same workbook object", () => {
    const { workbook, sheetId, table } = tableBook();
    const plan = applyTableOperation(workbook, sheetId, table.id, {
      kind: "updateTable",
      options: { name: "Sales", range: "A1:C4", style: "blue", showTotals: false, totals: {} },
    });
    expect(applyTableOperationPreview(workbook, plan)).toBe(workbook);
  });
});

describe("findTableForRange", () => {
  test("returns the table's range for a cell inside the table", () => {
    const { workbook, sheetId } = tableBook();
    expect(findTableForRange(workbook, sheetId, "B3")).toBe("A1:C4");
  });

  test("returns the occupied range of the sheet for a cell outside any table", () => {
    const book = createBlankWorkbook("Outside");
    const sheetId = book.sheets[0].id;
    book.sheets[0].cells.B2 = { raw: "x" };
    book.sheets[0].cells.D5 = { raw: "y" };
    const workbook = createWorksheetTable(book, sheetId, { name: "Dat", range: "B2:B3", headers: ["Hdr"] });
    expect(findTableForRange(workbook, sheetId, "B3")).toBe("B2:B3");
    expect(findTableForRange(workbook, sheetId, "D5")).toBe("B2:D5");
  });

  test("throws for a sheet that does not exist", () => {
    const { workbook } = tableBook();
    expect(() => findTableForRange(workbook, "missing", "A1")).toThrow("Sheet not found: missing");
  });
});

describe("suggestTableHeaderName", () => {
  test("accepts clean headers without a review", () => {
    const { workbook, sheetId } = tableBook();
    const suggestions = suggestTableHeaderName(workbook, sheetId, "A1:C1");
    expect(suggestions.map((s) => s.suggested)).toEqual(["Region", "Units", "Price"]);
    expect(suggestions.every((s) => s.requiresResolution === false)).toBe(true);
  });

  test("suggests a numbered name for a repeated header and marks it for review", () => {
    const book = salesBook();
    book.sheets[0].cells.B1 = { raw: "Region" };
    const [, repeated] = suggestTableHeaderName(book, book.sheets[0].id, "A1:C1");
    expect(repeated).toMatchObject({
      address: "B1",
      original: "Region",
      suggested: "Region2",
      requiresResolution: true,
    });
  });

  test("suggests a column number for a blank header", () => {
    const book = salesBook();
    delete book.sheets[0].cells.C1;
    const blank = suggestTableHeaderName(book, book.sheets[0].id, "A1:C1")[2];
    expect(blank).toMatchObject({ suggested: "Column3", original: null, requiresResolution: true });
  });

  test("throws for a sheet that does not exist", () => {
    const { workbook } = tableBook();
    expect(() => suggestTableHeaderName(workbook, "missing", "A1:C1")).toThrow("Sheet not found");
  });
});

describe("setTableColumnDataType", () => {
  test("reports the values that do not match the column's type", () => {
    const book = salesBook();
    book.sheets[0].cells.B3 = { raw: "x" };
    const { workbook, sheetId, table } = tableBook2(book);
    const typed = applyColumnTypePreview(
      workbook,
      previewColumnTypeChange(workbook, sheetId, table.id, {
        columnId: table.columns[1].id,
        dataType: "number",
        mode: "convert",
        invalid: "keep",
      }),
    );
    const check = setTableColumnDataType(typed, sheetId, table.id, table.columns[1].id);
    expect(check).toMatchObject({ dataType: "number", checked: 3, total: 3, truncated: false });
    expect(check.exceptions.map((e) => e.address)).toEqual(["B3"]);
  });

  test("a column with type any has no exceptions", () => {
    const book = salesBook();
    book.sheets[0].cells.B3 = { raw: "x" };
    const { workbook, sheetId, table } = tableBook2(book);
    expect(setTableColumnDataType(workbook, sheetId, table.id, table.columns[1].id)).toMatchObject({
      dataType: "any",
      exceptions: [],
    });
  });

  test("does not change the workbook", () => {
    const { workbook, sheetId, table } = tableBook();
    const before = JSON.stringify(workbook);
    setTableColumnDataType(workbook, sheetId, table.id, table.columns[1].id);
    expect(JSON.stringify(workbook)).toBe(before);
  });

  test("throws for a column that is not in the table", () => {
    const { workbook, sheetId, table } = tableBook();
    expect(() => setTableColumnDataType(workbook, sheetId, table.id, "missing")).toThrow(
      "Select an existing table column.",
    );
  });
});

/** Like tableBook, for a book already changed by the caller. */
function tableBook2(book) {
  const workbook = createWorksheetTable(book, book.sheets[0].id, { name: "Sales", range: "A1:C4" });
  return { workbook, sheetId: workbook.sheets[0].id, table: workbook.sheets[0].tables[0] };
}

describe("previewColumnTypeChange and applyColumnTypePreview", () => {
  test("converts text numbers to numbers, with separators", () => {
    const book = salesBook();
    book.sheets[0].cells.C2 = { raw: "1,234.5" };
    const { workbook, sheetId, table } = tableBook2(book);
    const plan = previewColumnTypeChange(workbook, sheetId, table.id, {
      columnId: table.columns[2].id,
      dataType: "number",
      mode: "convert",
      invalid: "reject",
    });
    expect(plan.canApply).toBe(true);
    expect(plan.changes.map((c) => [c.address, c.after.raw])).toEqual([["C2", 1234.5]]);
    const next = applyColumnTypePreview(workbook, plan);
    expect(next.sheets[0].tables[0].columns[2].dataType).toBe("number");
    expect(next.sheets[0].cells.C2.raw).toBe(1234.5);
  });

  test("with reject, a value that cannot be converted blocks the change", () => {
    const book = salesBook();
    book.sheets[0].cells.B3 = { raw: "x" };
    const { workbook, sheetId, table } = tableBook2(book);
    const plan = previewColumnTypeChange(workbook, sheetId, table.id, {
      columnId: table.columns[1].id,
      dataType: "number",
      mode: "convert",
      invalid: "reject",
    });
    expect(plan.canApply).toBe(false);
    expect(plan.exceptions.map((e) => [e.address, e.value])).toEqual([["B3", "x"]]);
    expect(() => applyColumnTypePreview(workbook, plan)).toThrow(
      "Resolve the listed exceptions or explicitly choose to retain them unchanged.",
    );
  });

  test("with keep, the unconvertible value stays and the change applies", () => {
    const book = salesBook();
    book.sheets[0].cells.B3 = { raw: "x" };
    const { workbook, sheetId, table } = tableBook2(book);
    const plan = previewColumnTypeChange(workbook, sheetId, table.id, {
      columnId: table.columns[1].id,
      dataType: "number",
      mode: "convert",
      invalid: "keep",
    });
    expect(plan.canApply).toBe(true);
    const next = applyColumnTypePreview(workbook, plan);
    expect(next.sheets[0].cells.B3.raw).toBe("x");
    expect(next.sheets[0].tables[0].columns[1].dataType).toBe("number");
  });

  test("tag mode sets no values and reports exceptions", () => {
    const book = salesBook();
    book.sheets[0].cells.B3 = { raw: "x" };
    const { workbook, sheetId, table } = tableBook2(book);
    const plan = previewColumnTypeChange(workbook, sheetId, table.id, {
      columnId: table.columns[1].id,
      dataType: "number",
      mode: "tag",
      invalid: "reject",
    });
    expect(plan.changes).toEqual([]);
    expect(plan.canApply).toBe(false);
  });

  test("setting the type to any removes it from the column", () => {
    const { workbook, sheetId, table } = tableBook();
    const typed = applyColumnTypePreview(
      workbook,
      previewColumnTypeChange(workbook, sheetId, table.id, {
        columnId: table.columns[1].id,
        dataType: "number",
        mode: "tag",
        invalid: "keep",
      }),
    );
    const cleared = applyColumnTypePreview(
      typed,
      previewColumnTypeChange(typed, sheetId, table.id, {
        columnId: table.columns[1].id,
        dataType: "any",
        mode: "tag",
        invalid: "keep",
      }),
    );
    expect(cleared.sheets[0].tables[0].columns[1].dataType).toBeUndefined();
  });

  test("a stale preview is refused", () => {
    const { workbook, sheetId, table } = tableBook();
    const plan = previewColumnTypeChange(workbook, sheetId, table.id, {
      columnId: table.columns[1].id,
      dataType: "number",
      mode: "convert",
      invalid: "keep",
    });
    const changed = structuredClone(workbook);
    changed.sheets[0].cells.E1 = { raw: "x" };
    expect(() => applyColumnTypePreview(changed, plan)).toThrow(
      "The workbook changed. Preview column types again.",
    );
  });

  test("a request with an unknown key is refused", () => {
    const { workbook, sheetId, table } = tableBook();
    expect(() =>
      previewColumnTypeChange(workbook, sheetId, table.id, {
        columnId: table.columns[1].id,
        dataType: "number",
        mode: "convert",
        invalid: "keep",
        locale: "de",
      }),
    ).toThrow("Invalid column type request.");
  });

  test("decimal and group separators must differ", () => {
    const { workbook, sheetId, table } = tableBook();
    expect(() =>
      previewColumnTypeChange(workbook, sheetId, table.id, {
        columnId: table.columns[1].id,
        dataType: "number",
        mode: "convert",
        invalid: "keep",
        decimal: ",",
        group: ",",
      }),
    ).toThrow("different decimal/group separators");
  });
});

describe("tableColumnDataTypes", () => {
  test("lists the five column types", () => {
    expect(tableColumnDataTypes).toEqual(["any", "text", "number", "boolean", "date"]);
  });

  test("every listed type is accepted by a column type preview", () => {
    const { workbook, sheetId, table } = tableBook();
    for (const dataType of tableColumnDataTypes) {
      expect(() =>
        previewColumnTypeChange(workbook, sheetId, table.id, {
          columnId: table.columns[1].id,
          dataType,
          mode: "tag",
          invalid: "keep",
        }),
      ).not.toThrow();
    }
  });

  test("a type that is not listed is refused", () => {
    const { workbook, sheetId, table } = tableBook();
    expect(() =>
      previewColumnTypeChange(workbook, sheetId, table.id, {
        columnId: table.columns[1].id,
        dataType: "money",
        mode: "tag",
        invalid: "keep",
      }),
    ).toThrow("Choose a supported table column type.");
  });
});

const northOnly = { range: "A1:C4", columns: [{ column: 0, values: ["North"] }] };

describe("applyFilterConditions", () => {
  test("a value list keeps the matching rows and reports the others as hidden", () => {
    const { workbook, sheetId } = tableBook();
    const result = applyFilterConditions(workbook, sheetId, northOnly);
    expect(result.context.matchedRows).toBe(2);
    expect(result.context.totalRows).toBe(3);
    expect(result.hiddenRows).toEqual([{ start: 2, end: 2 }]);
    expect(result.excludedRows).toEqual([{ start: 2, end: 2 }]);
    expect(result.fields.map((field) => field.label)).toEqual(["Region", "Units", "Price"]);
  });

  test("a number condition compares numerically", () => {
    const { workbook, sheetId } = tableBook();
    const result = applyFilterConditions(workbook, sheetId, {
      range: "A1:C4",
      columns: [
        {
          column: 1,
          conditions: [{ column: 1, type: "number", operator: "gt", value: 5 }],
        },
      ],
    });
    expect(result.context.matchedRows).toBe(2);
  });

  test("two conditions on one column combine with and by default, or when join is or", () => {
    const { workbook, sheetId } = tableBook();
    const conditions = [
      { column: 1, type: "number", operator: "gt", value: 8 },
      { column: 1, type: "number", operator: "lt", value: 5 },
    ];
    const matched = (join) =>
      applyFilterConditions(workbook, sheetId, {
        range: "A1:C4",
        columns: [{ column: 1, conditions, ...(join ? { join } : {}) }],
      }).context.matchedRows;
    expect(matched()).toBe(0);
    expect(matched("and")).toBe(0);
    expect(matched("or")).toBe(2);
  });

  test("advanced groups combine with or, and their conditions with and", () => {
    const { workbook, sheetId } = tableBook();
    const matched = (advanced) =>
      applyFilterConditions(workbook, sheetId, { range: "A1:C4", columns: [], advanced })
        .context.matchedRows;
    expect(
      matched([[{ column: 1, type: "number", operator: "gt", value: 8 }], [{ column: 1, type: "number", operator: "lt", value: 5 }]]),
    ).toBe(2);
    expect(
      matched([[{ column: 1, type: "number", operator: "gt", value: 3 }, { column: 1, type: "number", operator: "lt", value: 8 }]]),
    ).toBe(2);
    expect(
      matched([[{ column: 1, type: "number", operator: "gt", value: 5 }, { column: 1, type: "number", operator: "lt", value: 8 }]]),
    ).toBe(1);
  });

  test("text criteria match case-insensitively unless caseSensitive is set", () => {
    const { workbook, sheetId } = tableBook();
    const lower = { column: 0, values: ["north"] };
    expect(
      applyFilterConditions(workbook, sheetId, { range: "A1:C4", columns: [lower] }).context.matchedRows,
    ).toBe(2);
    expect(
      applyFilterConditions(workbook, sheetId, { range: "A1:C4", caseSensitive: true, columns: [lower] })
        .context.matchedRows,
    ).toBe(0);
  });

  test("unique keeps one of each set of identical rows", () => {
    const book = salesBook();
    book.sheets[0].cells.A4 = { raw: "North" };
    book.sheets[0].cells.B4 = { raw: 10 };
    book.sheets[0].cells.C4 = { raw: 2.5 };
    const { sheetId } = { sheetId: book.sheets[0].id };
    const result = applyFilterConditions(book, sheetId, { range: "A1:C4", unique: true, columns: [] });
    expect(result.context.visibleRows).toBe(2);
  });

  test("a column that is not in the range is refused", () => {
    const { workbook, sheetId } = tableBook();
    expect(() =>
      applyFilterConditions(workbook, sheetId, { range: "A1:C4", columns: [{ column: 9, values: ["x"] }] }),
    ).toThrow("Filter columns must be valid and distinct.");
  });

  test("a range with no data rows is refused", () => {
    const { workbook, sheetId } = tableBook();
    expect(() => applyFilterConditions(workbook, sheetId, { range: "A1:C1", columns: [] })).toThrow(
      "Filter range needs one header row and at least one data row",
    );
  });

  test("a text condition needs a text value", () => {
    const { workbook, sheetId } = tableBook();
    expect(() =>
      applyFilterConditions(workbook, sheetId, {
        range: "A1:C4",
        columns: [
          {
            column: 0,
            conditions: [{ column: 0, type: "text", operator: "equals", value: 3 }],
          },
        ],
      }),
    ).toThrow("Text criteria require a text value.");
  });

  test("does not change the workbook", () => {
    const { workbook, sheetId } = tableBook();
    const before = JSON.stringify(workbook);
    applyFilterConditions(workbook, sheetId, northOnly);
    expect(JSON.stringify(workbook)).toBe(before);
  });
});

describe("buildFilterFacetValues", () => {
  test("lists each distinct value with its row count", () => {
    const { workbook, sheetId } = tableBook();
    const facet = buildFilterFacetValues(workbook, sheetId, { range: "A1:C4", columns: [] }, 0);
    expect(facet.distinct).toBe(2);
    expect(facet.values.map((v) => [v.value, v.count])).toEqual([
      ["North", 2],
      ["South", 1],
    ]);
  });

  test("the search text narrows the list, ignoring case", () => {
    const { workbook, sheetId } = tableBook();
    const facet = buildFilterFacetValues(workbook, sheetId, { range: "A1:C4", columns: [] }, 0, "NO");
    expect(facet.values.map((v) => v.value)).toEqual(["North"]);
  });

  test("the list only counts rows that the other filters keep", () => {
    const { workbook, sheetId } = tableBook();
    const facet = buildFilterFacetValues(workbook, sheetId, northOnly, 1);
    expect(facet.values.map((v) => v.value)).toEqual([10, 7]);
  });

  test("a facet column outside the range is refused", () => {
    const { workbook, sheetId } = tableBook();
    expect(() => buildFilterFacetValues(workbook, sheetId, { range: "A1:C4", columns: [] }, 5)).toThrow(
      "Facet column is outside the filter range.",
    );
  });
});

describe("saved filter views", () => {
  test("createSavedFilterView stores the definition and the counts when it was saved", () => {
    const { workbook, sheetId } = tableBook();
    const saved = createSavedFilterView(workbook, sheetId, northOnly, "North only");
    expect(saved.savedFilterViews).toHaveLength(1);
    expect(saved.savedFilterViews[0]).toMatchObject({
      name: "North only",
      sheetId,
      sheetName: "Sheet1",
      population: { totalRows: 3, matchedRows: 2, visibleRows: 2, hiddenMatchingRows: 0 },
    });
    expect(workbook.savedFilterViews).toBeUndefined();
  });

  test("passing a view id replaces that view", () => {
    const { workbook, sheetId } = tableBook();
    const first = createSavedFilterView(workbook, sheetId, northOnly, "North only");
    const id = first.savedFilterViews[0].id;
    const replaced = createSavedFilterView(
      first,
      sheetId,
      { range: "A1:C4", columns: [{ column: 0, values: ["South"] }] },
      "South only",
      id,
    );
    expect(replaced.savedFilterViews).toHaveLength(1);
    expect(replaced.savedFilterViews[0]).toMatchObject({
      id,
      name: "South only",
      population: { matchedRows: 1 },
    });
  });

  test("addSavedFilterView renames a view", () => {
    const { workbook, sheetId } = tableBook();
    const saved = createSavedFilterView(workbook, sheetId, northOnly, "North only");
    const renamed = addSavedFilterView(saved, saved.savedFilterViews[0].id, "Northern");
    expect(renamed.savedFilterViews[0].name).toBe("Northern");
    expect(saved.savedFilterViews[0].name).toBe("North only");
  });

  test("addSavedFilterView refuses an empty name and an unknown view", () => {
    const { workbook, sheetId } = tableBook();
    const saved = createSavedFilterView(workbook, sheetId, northOnly, "North only");
    const id = saved.savedFilterViews[0].id;
    expect(() => addSavedFilterView(saved, id, "   ")).toThrow(
      "Use a nonempty filter view name of up to 80 characters.",
    );
    expect(() => addSavedFilterView(saved, "missing", "x")).toThrow(
      "This saved filter view no longer exists.",
    );
  });

  test("removeSavedFilterView deletes the view and refuses an unknown one", () => {
    const { workbook, sheetId } = tableBook();
    const saved = createSavedFilterView(workbook, sheetId, northOnly, "North only");
    expect(removeSavedFilterView(saved, saved.savedFilterViews[0].id).savedFilterViews).toEqual([]);
    expect(() => removeSavedFilterView(saved, "missing")).toThrow(
      "This saved filter view no longer exists.",
    );
  });

  test("createSavedFilterView refuses a view name that is too long", () => {
    const { workbook, sheetId } = tableBook();
    expect(() =>
      createSavedFilterView(workbook, sheetId, northOnly, "x".repeat(81)),
    ).toThrow("Use a nonempty filter view name of up to 80 characters.");
  });
});

describe("rewriteStructuredReferences", () => {
  test("resolves a named column to its data range", () => {
    const { workbook, sheetId } = tableBook();
    expect(rewriteStructuredReferences(workbook, sheetId, "Sales[Units]")).toEqual({
      sheetId,
      range: "B2:B4",
    });
  });

  test("resolves adjacent columns given as a span", () => {
    const { workbook, sheetId } = tableBook();
    expect(
      rewriteStructuredReferences(workbook, sheetId, "Sales[[Units]:[Price]]")?.range,
    ).toBe("B2:C4");
  });

  test("[@Column] resolves to the row the address is in", () => {
    const { workbook, sheetId } = tableBook();
    expect(rewriteStructuredReferences(workbook, sheetId, "[@Units]", "B3")).toEqual({
      sheetId,
      range: "B3:B3",
    });
  });

  test("[@Column] returns null outside the table's data rows", () => {
    const { workbook, sheetId } = tableBook();
    expect(rewriteStructuredReferences(workbook, sheetId, "[@Units]", "E3")).toBeNull();
    expect(rewriteStructuredReferences(workbook, sheetId, "[@Units]", "B1")).toBeNull();
  });

  test("returns null for an unknown column, an unknown table and a totals row that does not exist", () => {
    const { workbook, sheetId } = tableBook();
    expect(rewriteStructuredReferences(workbook, sheetId, "Sales[Nope]")).toBeNull();
    expect(rewriteStructuredReferences(workbook, sheetId, "Missing[Units]")).toBeNull();
    expect(rewriteStructuredReferences(workbook, sheetId, "Sales[#Totals]")).toBeNull();
  });

  test("returns null for text that is not a structured reference", () => {
    const { workbook, sheetId } = tableBook();
    expect(rewriteStructuredReferences(workbook, sheetId, "A1")).toBeNull();
  });
});

describe("getPivotSourceFields", () => {
  test("returns the header texts of the range in column order", () => {
    const book = salesBook();
    expect(getPivotSourceFields(book, book.sheets[0].id, "A1:C4")).toEqual([
      "Region",
      "Units",
      "Price",
    ]);
  });

  test("throws for a sheet that does not exist", () => {
    const book = salesBook();
    expect(() => getPivotSourceFields(book, "missing", "A1:C4")).toThrow("Sheet not found");
  });
});

describe("buildPivotTableOutput", () => {
  test("writes a pivot to a new sheet with a grand total and records its figures", () => {
    const { workbook, sheetId } = pivotBook();
    expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1", "By region"]);
    const output = workbook.sheets[1];
    expect(output.cells.B4).toMatchObject({ raw: 21 });
    expect(output.cells.A4).toMatchObject({ raw: "Grand Total" });
    const pivot = output.pivots[0];
    expect(pivot.source).toEqual({ sheetId, range: "A1:C4" });
    expect(pivot.refresh.measures[0]).toMatchObject({ name: "Sum of Units", result: 21 });
    expect(pivot.refresh.includedRows).toBe(3);
  });

  test("grandTotals false leaves out the grand total row", () => {
    const { workbook } = pivotBook({ grandTotals: false });
    expect(Object.values(workbook.sheets[1].cells).map((c) => c.raw)).not.toContain("Grand Total");
  });

  test("a source filter excludes rows before aggregation", () => {
    const { workbook } = pivotBook({
      filters: [{ field: "Units", operator: "gte", value: 5 }],
    });
    const pivot = workbook.sheets[1].pivots[0];
    expect(pivot.refresh.measures[0].result).toBe(17);
    expect(pivot.refresh.excludedRows).toBe(1);
  });

  test("a row percentage share is written as a fraction of the row total", () => {
    const book = salesBook();
    const workbook = buildPivotTableOutput(book, {
      name: "Share",
      source: { sheetId: book.sheets[0].id, range: "A1:C4" },
      rowFields: ["Region"],
      columnFields: ["Price"],
      values: [{ field: "Units", operation: "sum", showAs: "percentRow" }],
    });
    const output = workbook.sheets[1];
    const northAt2_5 = Object.entries(output.cells).find(
      ([, cell]) => cell.raw === "2.5 / Sum of Units (% of row total)",
    );
    expect(northAt2_5).toBeDefined();
    expect(output.cells.B2.raw).toBeCloseTo(10 / 17, 12);
  });

  test("an unknown aggregation is refused", () => {
    const book = salesBook();
    expect(() =>
      buildPivotTableOutput(book, {
        name: "Bad",
        source: { sheetId: book.sheets[0].id, range: "A1:C4" },
        rowFields: ["Region"],
        columnFields: [],
        values: [{ field: "Units", operation: "median" }],
      }),
    ).toThrow("Invalid pivot aggregation.");
  });

  test("a field that is not in the source is refused", () => {
    const book = salesBook();
    expect(() =>
      buildPivotTableOutput(book, {
        name: "Bad",
        source: { sheetId: book.sheets[0].id, range: "A1:C4" },
        rowFields: ["Nope"],
        columnFields: [],
        values: [{ field: "Units", operation: "sum" }],
      }),
    ).toThrow("source field was not found");
  });

  test("a pivot needs at least one value field", () => {
    const book = salesBook();
    expect(() =>
      buildPivotTableOutput(book, {
        name: "Empty",
        source: { sheetId: book.sheets[0].id, range: "A1:C4" },
        rowFields: ["Region"],
        columnFields: [],
        values: [],
      }),
    ).toThrow("Choose between 1 and 100 pivot value fields.");
  });
});

describe("refreshPivotTable", () => {
  test("without a destination, writes the refreshed result to a new sheet", () => {
    const { workbook, pivotId } = pivotBook();
    const edited = structuredClone(workbook);
    edited.sheets[0].cells.B2 = { raw: 100 };
    const refreshed = refreshPivotTable(edited, pivotId);
    expect(refreshed.sheets.map((sheet) => sheet.name)).toEqual([
      "Sheet1",
      "By region",
      "By region refreshed",
    ]);
    expect(refreshed.sheets[2].cells.B2.raw).toBe(107);
  });

  test("with a destination, writes the output at that address", () => {
    const { workbook, pivotId, sheetId } = pivotBook();
    const edited = structuredClone(workbook);
    edited.sheets[0].cells.B2 = { raw: 100 };
    const refreshed = refreshPivotTable(edited, pivotId, { sheetId, address: "E1" });
    expect(refreshed.sheets[0].cells.F2).toMatchObject({ raw: 107 });
    expect(refreshed.sheets.map((sheet) => sheet.name)).toEqual(["Sheet1", "By region"]);
  });

  test("throws for a pivot that does not exist", () => {
    const { workbook } = pivotBook();
    expect(() => refreshPivotTable(workbook, "missing")).toThrow("Pivot definition not found.");
  });
});

describe("getPivotRefreshState", () => {
  test("a pivot built from unchanged data is current", () => {
    const { workbook, pivotId } = pivotBook();
    expect(getPivotRefreshState(workbook, pivotId)).toMatchObject({
      state: "current",
      outputModified: false,
    });
  });

  test("a change to the source makes the pivot stale", () => {
    const { workbook, pivotId } = pivotBook();
    const edited = structuredClone(workbook);
    edited.sheets[0].cells.B2 = { raw: 100 };
    expect(getPivotRefreshState(edited, pivotId).state).toBe("stale");
  });

  test("an edit to the output makes the pivot stale and reports it", () => {
    const { workbook, pivotId } = pivotBook();
    const edited = structuredClone(workbook);
    edited.sheets[1].cells.B3 = { raw: 999 };
    expect(getPivotRefreshState(edited, pivotId)).toMatchObject({
      state: "stale",
      outputModified: true,
    });
  });

  test("a renamed source header makes the pivot unavailable", () => {
    const { workbook, pivotId } = pivotBook();
    const edited = structuredClone(workbook);
    edited.sheets[0].cells.B1 = { raw: "Quantity" };
    const state = getPivotRefreshState(edited, pivotId);
    expect(state.state).toBe("unavailable");
    expect(state.reason).toContain("source field was not found");
  });

  test("a pivot with no refresh record is unverified", () => {
    const { workbook, pivotId } = pivotBook();
    const stripped = structuredClone(workbook);
    delete stripped.sheets[1].pivots[0].refresh;
    expect(getPivotRefreshState(stripped, pivotId).state).toBe("unverified");
  });

  test("throws for a pivot that does not exist", () => {
    const { workbook } = pivotBook();
    expect(() => getPivotRefreshState(workbook, "missing")).toThrow("Pivot definition not found.");
  });
});

describe("buildPivotDrillDown", () => {
  test("lists the source rows behind a value cell", () => {
    const { workbook, pivotId } = pivotBook();
    const drill = buildPivotDrillDown(workbook, pivotId, "B2");
    expect(drill.underlyingValue).toBe(17);
    expect(drill.fields).toEqual(["Region", "Units", "Price"]);
    expect(drill.records).toHaveLength(2);
    expect(drill.records[0]).toMatchObject({
      locator: { sheetName: "Sheet1", row: 2, range: "A2:C2" },
      values: ["North", 10, 2.5],
    });
  });

  test("the grand total can be drilled", () => {
    const { workbook, pivotId } = pivotBook();
    expect(buildPivotDrillDown(workbook, pivotId, "B4").records).toHaveLength(3);
  });

  test("a header cell cannot be drilled", () => {
    const { workbook, pivotId } = pivotBook();
    expect(() => buildPivotDrillDown(workbook, pivotId, "A1")).toThrow(
      "headers and group labels cannot be drilled",
    );
  });

  test("throws for a pivot that does not exist", () => {
    const { workbook } = pivotBook();
    expect(() => buildPivotDrillDown(workbook, "missing", "B2")).toThrow("Pivot definition not found.");
  });
});

describe("extractPivotDrillThrough", () => {
  test("writes the source rows to a new sheet", () => {
    const { workbook, pivotId } = pivotBook();
    const extracted = extractPivotDrillThrough(workbook, pivotId, "B2", {
      snapshot: JSON.stringify(workbook),
    });
    expect(extracted.sheets).toHaveLength(3);
    const sheet = extracted.sheets[2];
    expect(sheet.name).toBe("By region details");
    expect(sheet.cells.A2).toMatchObject({ raw: "Sheet1!A2:C2" });
    expect(sheet.cells.F2).toMatchObject({ raw: "North" });
    expect(sheet.cells.G3).toMatchObject({ raw: 7 });
  });

  test("a stale snapshot is refused", () => {
    const { workbook, pivotId } = pivotBook();
    expect(() =>
      extractPivotDrillThrough(workbook, pivotId, "B2", { snapshot: "{}" }),
    ).toThrow("The workbook changed. Preview drill-through again.");
  });

  test("a read-only workbook is refused", () => {
    const { workbook, pivotId } = pivotBook();
    expect(() =>
      extractPivotDrillThrough(workbook, pivotId, "B2", {
        snapshot: JSON.stringify(workbook),
        readOnly: true,
      }),
    ).toThrow("This workbook is read-only.");
  });
});

describe("getPivotSlicerItems", () => {
  test("lists the distinct items of a row field", () => {
    const { workbook, pivotId } = pivotBook();
    expect(getPivotSlicerItems(workbook, pivotId, "Region")).toEqual(["North", "South"]);
  });

  test("lists the numbers of a value field in source order", () => {
    const { workbook, pivotId } = pivotBook();
    expect(getPivotSlicerItems(workbook, pivotId, "Units")).toEqual([10, 4, 7]);
  });

  test("a field that is not in the source is refused", () => {
    const { workbook, pivotId } = pivotBook();
    expect(() => getPivotSlicerItems(workbook, pivotId, "Nope")).toThrow("Broken pivot field");
  });
});

describe("applyPivotViewChange", () => {
  test("a slicer writes a filtered copy to a new sheet and leaves the original", () => {
    const { workbook, pivotId } = pivotBook();
    const filtered = applyPivotViewChange(workbook, pivotId, {
      slicers: [{ field: "Region", selected: ["South"] }],
    }, { snapshot: JSON.stringify(workbook) });
    expect(filtered.sheets.map((sheet) => sheet.name)).toEqual([
      "Sheet1",
      "By region",
      "By region filtered",
    ]);
    const copy = filtered.sheets[2].pivots[0];
    expect(copy.refresh).toMatchObject({ includedRows: 1, excludedRows: 2 });
    expect(filtered.sheets[1].pivots[0].id).toBe(pivotId);
  });

  test("a sort is recorded on the new pivot", () => {
    const { workbook, pivotId } = pivotBook();
    const sorted = applyPivotViewChange(workbook, pivotId, {
      sorts: [{ field: "Region", direction: "descending" }],
    }, { snapshot: JSON.stringify(workbook) });
    expect(sorted.sheets[2].pivots[0].sorts).toEqual([{ field: "Region", direction: "descending" }]);
  });

  test("a chart setting changes the kind reported by buildPivotChartModel", () => {
    const { workbook, pivotId } = pivotBook();
    const charted = applyPivotViewChange(workbook, pivotId, {
      chart: { kind: "bar", valueIndex: 0 },
    }, { snapshot: JSON.stringify(workbook) });
    const newId = charted.sheets[2].pivots[0].id;
    expect(buildPivotChartModel(charted, newId).kind).toBe("bar");
  });

  test("a view change that is not a view setting is refused", () => {
    const { workbook, pivotId } = pivotBook();
    expect(() =>
      applyPivotViewChange(workbook, pivotId, { destination: { address: "Z1" } }, {
        snapshot: JSON.stringify(workbook),
      }),
    ).toThrow("Pivot view changes can only set filters, sorting, value display and chart settings.");
  });

  test("a stale snapshot is refused", () => {
    const { workbook, pivotId } = pivotBook();
    expect(() =>
      applyPivotViewChange(workbook, pivotId, { slicers: [] }, { snapshot: "{}" }),
    ).toThrow("The workbook changed. Preview pivot filters again.");
  });
});

describe("buildPivotChartModel", () => {
  test("gives one point per row group of the first value field", () => {
    const { workbook, pivotId } = pivotBook();
    expect(buildPivotChartModel(workbook, pivotId)).toEqual({
      kind: "column",
      measure: "Sum of Units",
      showAs: "value",
      points: [
        { label: "North", value: 17 },
        { label: "South", value: 4 },
      ],
      sourceRevision: 0,
    });
  });

  test("throws for a pivot that does not exist", () => {
    const { workbook } = pivotBook();
    expect(() => buildPivotChartModel(workbook, "missing")).toThrow("Pivot definition not found.");
  });
});

describe("pivotAggregateOperations", () => {
  test("lists the six aggregations a value field may use", () => {
    expect(pivotAggregateOperations).toEqual(["sum", "count", "average", "min", "max", "distinctCount"]);
  });

  test("each listed aggregation builds a pivot", () => {
    const book = salesBook();
    for (const operation of pivotAggregateOperations) {
      expect(() =>
        buildPivotTableOutput(book, {
          name: operation,
          source: { sheetId: book.sheets[0].id, range: "A1:C4" },
          rowFields: ["Region"],
          columnFields: [],
          values: [{ field: "Units", operation }],
        }),
      ).not.toThrow();
    }
  });
});
