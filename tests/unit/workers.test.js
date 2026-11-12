import { describe, expect, test, vi } from "vitest";
import ExcelJS from "exceljs";
import { createBlankWorkbook } from "../../src/app/03-workbook-model.js";

/**
 * The workers are plain scripts that talk through `self`. Loading one with a
 * stand-in `self` runs the same code the browser runs, without a browser.
 */
async function startWorker(path) {
  vi.resetModules();
  let deliver;
  const scope = { postMessage: (message) => deliver?.(message) };
  vi.stubGlobal("self", scope);
  await import(path);
  return (data) =>
    new Promise((resolve) => {
      deliver = resolve;
      scope.onmessage({ data });
    });
}

const bytesOf = (text) => new TextEncoder().encode(text);

function workbookWith(cells) {
  const workbook = createBlankWorkbook("Worker");
  for (const [address, raw] of Object.entries(cells))
    workbook.sheets[0].cells[address] = { raw };
  return workbook;
}

describe("calculation worker", () => {
  const start = () => startWorker("../../src/workers/calculation.worker.js");

  test("answers with the request id and a complete calculation report", async () => {
    const send = await start();
    const reply = await send({
      id: 7,
      force: true,
      workbook: workbookWith({ A1: "2", A2: "=A1*21" }),
    });
    expect(reply.id).toBe(7);
    expect(reply.error).toBeUndefined();
    expect(reply.result.report.execution.state).toBe("complete");
    expect(typeof reply.result.inputSignature).toBe("string");
    expect(reply.result.operationalErrors).toEqual([]);
  });

  test("returns the cells a dynamic array spills into", async () => {
    const send = await start();
    const workbook = workbookWith({ A3: "=SEQUENCE(2)" });
    const { result } = await send({ id: 1, force: true, workbook });
    expect(result.spills).toEqual([
      {
        address: "A4",
        anchor: "A3",
        sheetId: workbook.sheets[0].id,
        result: { type: "number", value: 2 },
      },
    ]);
  });

  test("gives the same input signature for the same workbook", async () => {
    const send = await start();
    const workbook = workbookWith({ A1: "1", B1: "=A1+1" });
    const first = await send({ id: 1, force: true, workbook });
    const second = await send({ id: 2, force: true, workbook });
    expect(second.result.inputSignature).toBe(first.result.inputSignature);
  });

  test("refuses a delta that does not follow the revision it holds", async () => {
    const send = await start();
    const workbook = workbookWith({ A1: "1" });
    await send({ id: 1, force: true, workbook });
    const reply = await send({ id: 2, workbook, baseRevision: 99 });
    expect(reply).toEqual({
      id: 2,
      error: "Calculation worker failed; no result was published.",
    });
  });

  test("recovers after a refused request", async () => {
    const send = await start();
    const workbook = workbookWith({ A1: "1" });
    await send({ id: 1, workbook, baseRevision: 5 });
    const reply = await send({ id: 2, force: true, workbook });
    expect(reply.result.report.execution.state).toBe("complete");
  });
});

describe("file parser worker", () => {
  // Each worker instance handles one file, as in the application.
  const parse = async (message) =>
    (await startWorker("../../src/workers/file-parser.worker.js"))(message);

  test("imports delimited text with quoted fields and typed numbers", async () => {
    const reply = await parse({
      id: "csv",
      operation: "delimited",
      fileName: "ledger.csv",
      bytes: bytesOf('Name,Amount\n"Smith, J",1200.5\nLee,-3\n'),
    });
    expect(reply.ok).toBe(true);
    expect(reply.value.report).toMatchObject({
      records: 3,
      columns: 2,
      issueCount: 0,
    });
    expect(reply.value.report.preview[1].values).toEqual(["Smith, J", 1200.5]);
    expect(reply.value.report.preview[2].values).toEqual(["Lee", -3]);
    const cells = reply.value.workbook.sheets[0].cells;
    expect(cells.A2.raw).toBe("Smith, J");
    expect(cells.B2.raw).toBe(1200.5);
    expect(cells.B3.raw).toBe(-3);
  });

  test("imports an XLSX file with values and formulas", async () => {
    const source = new ExcelJS.Workbook();
    const sheet = source.addWorksheet("Data");
    sheet.getCell("A1").value = "Item";
    sheet.getCell("B1").value = 12.5;
    sheet.getCell("B2").value = { formula: "B1*2", result: 25 };
    const reply = await parse({
      id: "xlsx",
      operation: "workbook",
      fileName: "book.xlsx",
      bytes: new Uint8Array(await source.xlsx.writeBuffer()),
    });
    expect(reply.ok).toBe(true);
    const [imported] = reply.value.workbook.sheets;
    expect(imported.name).toBe("Data");
    expect(imported.cells).toMatchObject({
      A1: { raw: "Item" },
      B1: { raw: 12.5 },
      B2: { raw: "=B1*2" },
    });
  });

  test("rejects an unknown operation", async () => {
    const reply = await parse({
      id: "op",
      operation: "nope",
      fileName: "ledger.csv",
      bytes: bytesOf("x"),
    });
    expect(reply).toMatchObject({
      id: "op",
      ok: false,
      error: { code: "FILE_TYPE" },
    });
  });

  test("rejects a payload that is not bytes", async () => {
    const reply = await parse({
      id: "bytes",
      operation: "delimited",
      fileName: "ledger.csv",
      bytes: "not bytes",
    });
    expect(reply).toMatchObject({ ok: false, error: { code: "FILE_INVALID" } });
  });

  test("rejects a file that is not a real XLSX archive", async () => {
    const reply = await parse({
      id: "fake",
      operation: "workbook",
      fileName: "book.xlsx",
      bytes: bytesOf("this is not a zip archive"),
    });
    expect(reply.ok).toBe(false);
    expect(typeof reply.error.code).toBe("string");
  });

  test("rejects native workbook data that is not valid JSON", async () => {
    const reply = await parse({
      id: "json",
      operation: "json",
      fileName: "book.json",
      bytes: bytesOf("{ not json"),
    });
    expect(reply.ok).toBe(false);
  });
});
