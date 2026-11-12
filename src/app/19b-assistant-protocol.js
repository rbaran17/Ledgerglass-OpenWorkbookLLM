import {
  defineDataPropertyByKey,
  normalizeBindingName,
  normalizeDatePolicy,
  tokenizeFormulaText,
  walkFormulaNodes,
  workbookUsesNowOrToday,
} from "./01-schemas-date-functions-and-validation.js";
import { array, discriminatedUnion, literal, object, string } from "zod";
import {
  builtinFunctionNames,
  calculateWorkbookWithSettings,
  columnIndexToLetters,
  evaluateWorkbookFormulas,
  parseRangeAddress,
} from "./03-workbook-model.js";
import {
  supportedFontFamilies,
  workbookLimits,
} from "./02-text-and-financial-functions.js";

/* From 20-calculation-options-and-assistant.js */
const assistantRequestLimits = Object.freeze({
  filesPerMessage: 4,
  fileBytes: 5 * 1024 * 1024,
  totalFileBytes: 10 * 1024 * 1024,
  historyMessages: 20,
  historyCharacters: 5e4,
  promptCharacters: 1e4,
});
const attachmentMediaTypes = {
  txt: "text/plain",
  md: "text/markdown",
  csv: "text/csv",
  tsv: "text/tab-separated-values",
  json: "application/json",
  pdf: "application/pdf",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  webp: "image/webp",
};
const textAttachmentExtensions = new Set(["txt", "md", "csv", "tsv", "json"]);
const getFileExtension = (e) => e.split(".").pop()?.toLowerCase() ?? "";
function assertCondition(e, r) {
  if (!e) throw new Error(r);
}
function isSafeAttachmentFilename(e) {
  return (
    typeof e == "string" &&
    e.length > 0 &&
    e.length <= 180 &&
    !/[\x00-\x1f\x7f/\\]/.test(e)
  );
}
function matchesFileSignature(attachmentBytes, fileExtension) {
  const t = (...n) => n.every((o, a) => attachmentBytes[a] === o);
  return fileExtension === "pdf"
    ? t(37, 80, 68, 70, 45)
    : fileExtension === "png"
      ? t(137, 80, 78, 71, 13, 10, 26, 10)
      : fileExtension === "jpg" || fileExtension === "jpeg"
        ? t(255, 216, 255)
        : fileExtension === "webp"
          ? t(82, 73, 70, 70) &&
            String.fromCharCode(...attachmentBytes.slice(8, 12)) === "WEBP"
          : fileExtension === "xlsx" || fileExtension === "docx"
            ? t(80, 75, 3, 4)
            : textAttachmentExtensions.has(fileExtension);
}
function parseAssistantChatRequest(e) {
  (assertCondition(
    hasOnlyAllowedKeys(e, ["version", "messages", "attachments"]) &&
      e.version === 1 &&
      Array.isArray(e.messages) &&
      Array.isArray(e.attachments),
    "Invalid chat request.",
  ),
    assertCondition(
      e.messages.length <= assistantRequestLimits.historyMessages,
      "This conversation has reached 20 prior messages. Start a new chat to continue.",
    ));
  let r = 0,
    t = 0;
  const n = (a) => {
    assertCondition(
      Array.isArray(a) && a.length <= assistantRequestLimits.filesPerMessage,
      "Attach up to 4 files per message.",
    );
    const s = new Set();
    return a.map((d) => {
      (assertCondition(
        hasOnlyAllowedKeys(d, [
          "id",
          "name",
          "mediaType",
          "size",
          "encoding",
          "data",
        ]) &&
          typeof d.id == "string" &&
          d.id.length > 0 &&
          d.id.length <= 100 &&
          !s.has(d.id),
        "Invalid or duplicate attachment.",
      ),
        s.add(d.id),
        assertCondition(
          isSafeAttachmentFilename(d.name) &&
            Object.hasOwn(attachmentMediaTypes, getFileExtension(d.name)) &&
            d.mediaType === attachmentMediaTypes[getFileExtension(d.name)],
          "Unsupported attachment type.",
        ),
        assertCondition(
          Number.isSafeInteger(d.size) &&
            Number(d.size) > 0 &&
            Number(d.size) <= assistantRequestLimits.fileBytes &&
            typeof d.data == "string" &&
            d.data.length <=
              Math.ceil(assistantRequestLimits.fileBytes / 3) * 4,
          "Attachment exceeds the 5 MB limit.",
        ));
      const u = textAttachmentExtensions.has(getFileExtension(d.name));
      assertCondition(
        d.encoding === (u ? "utf8" : "base64"),
        "Invalid attachment encoding.",
      );
      let attachmentBytes;
      return (
        u
          ? (assertCondition(
              !d.data.includes("\0"),
              "Invalid text attachment.",
            ),
            (attachmentBytes = new TextEncoder().encode(d.data)))
          : (assertCondition(
              d.data.length % 4 === 0 && /^[A-Za-z0-9+/]*={0,2}$/.test(d.data),
              "Invalid attachment encoding.",
            ),
            (attachmentBytes = Uint8Array.from(atob(d.data), (c) =>
              c.charCodeAt(0),
            ))),
        assertCondition(
          attachmentBytes.length === d.size &&
            matchesFileSignature(attachmentBytes, getFileExtension(d.name)),
          "Attachment content or size is invalid.",
        ),
        (r += Number(d.size)),
        assertCondition(
          r <= assistantRequestLimits.totalFileBytes,
          "The conversation contains more than 10 MB of attachments. Remove files or start a new chat.",
        ),
        {
          id: d.id,
          name: d.name,
          mediaType: d.mediaType,
          size: Number(d.size),
          encoding: d.encoding,
          data: d.data,
        }
      );
    });
  };
  return {
    version: 1,
    messages: e.messages.map(
      (message) => (
        assertCondition(
          hasOnlyAllowedKeys(message, ["role", "text", "attachments"]) &&
            (message.role === "user" || message.role === "assistant") &&
            typeof message.text == "string" &&
            message.text.length <= 1e4,
          "Invalid chat history message.",
        ),
        assertCondition(
          message.role === "user" || message.attachments === undefined,
          "Assistant messages cannot attach files.",
        ),
        (t += message.text.length),
        assertCondition(
          t <= assistantRequestLimits.historyCharacters,
          "This conversation has reached its context limit. Start a new chat to continue.",
        ),
        {
          role: message.role,
          text: message.text,
          ...(message.attachments
            ? { attachments: n(message.attachments) }
            : {}),
        }
      ),
    ),
    attachments: n(e.attachments),
  };
}
const assistantErrorMessages = {
  timeout:
    "The assistant request timed out. Retry or check the host provider connection. No workbook changes were made.",
  refused:
    "The assistant declined this request. Rephrase the task or contact the host administrator. No workbook changes were made.",
  malformed_response:
    "The assistant returned an invalid response. Ask the host to check the response schema, values and selection scope, then retry. No workbook changes were made.",
  unavailable:
    "The assistant provider is unavailable. Check the host connection, credentials or service availability, then retry. No workbook changes were made.",
  cancelled: "Assistant request cancelled. No workbook changes were made.",
  invalid_request:
    "The assistant request is invalid. Use a prompt of 1 to 10000 characters and select at most 500 cells. The host must supply a complete bounded context from buildContext; then retry.",
};
class LLMRequestError extends Error {
  constructor(t) {
    const n =
      typeof t == "string" && Object.hasOwn(assistantErrorMessages, t)
        ? t
        : "unavailable";
    super(assistantErrorMessages[n]);
    defineDataPropertyByKey(this, "code");
    ((this.name = "LLMRequestError"), (this.code = n));
  }
}
const answerGroundingLimits = Object.freeze({
  claims: 20,
  sources: 20,
  rangesPerClaim: 20,
  referencedCells: 1e3,
  receiptBytes: 2e6,
});
const identifierSchema = string().regex(/^[A-Za-z][A-Za-z0-9_-]{0,59}$/);
const Pk = string().trim().min(1).max(1e4);
const rangeAddressSchema = string()
  .regex(/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/)
  .max(32);
const answerGroundingSchema = object({
  claims: array(
    discriminatedUnion("kind", [
      object({
        id: identifierSchema,
        kind: literal("workbook"),
        text: Pk,
        ranges: array(rangeAddressSchema)
          .min(1)
          .max(answerGroundingLimits.rangesPerClaim),
      }).strict(),
      object({
        id: identifierSchema,
        kind: literal("external"),
        text: Pk,
        sourceIds: array(identifierSchema)
          .min(1)
          .max(answerGroundingLimits.sources),
      }).strict(),
      object({
        id: identifierSchema,
        kind: literal("commentary"),
        text: Pk,
      }).strict(),
    ]),
  )
    .min(1)
    .max(answerGroundingLimits.claims),
  externalSources: array(
    object({
      id: identifierSchema,
      title: string().trim().min(1).max(300),
      url: string().max(2048),
    }).strict(),
  ).max(answerGroundingLimits.sources),
}).strict();
function validateCitationRange(range) {
  if (!/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/.test(range))
    throw Error("Citation requires an explicit A1 cell or rectangle.");
  const rangeAddress = parseRangeAddress(range),
    t = range.split(":");
  if (
    t[1] &&
    `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}` !==
      t[0]
  )
    throw Error("Citation range must run top left to bottom right.");
  if (
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    500
  )
    throw Error("Citations support at most 500 cells per range.");
  const n = [];
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  )
    for (
      let firstCol = rangeAddress.firstCol;
      firstCol <= rangeAddress.lastCol;
      firstCol++
    )
      n.push(`${columnIndexToLetters(firstCol)}${firstRow + 1}`);
  return n;
}
function validateExternalSourceUrl(url) {
  if (
    typeof url != "string" ||
    url.length > 2048 ||
    url !== url.trim() ||
    /[\u0000-\u0020\u007f]/.test(url)
  )
    throw Error(
      "External sources require a valid HTTPS URL without credentials.",
    );
  let r;
  try {
    r = new URL(url);
  } catch {
    throw Error(
      "External sources require a valid HTTPS URL without credentials.",
    );
  }
  if (r.protocol !== "https:" || r.username || r.password || !r.hostname)
    throw Error(
      "External sources require a valid HTTPS URL without credentials.",
    );
  return r.href;
}
function validateAnswerGrounding(grounding, r) {
  const t = answerGroundingSchema.safeParse(grounding);
  if (!t.success)
    throw Error(
      "Invalid answer grounding: provide typed claims and separate external sources.",
    );
  const data = t.data,
    o = new Set(r.cells.map((cell) => cell.address)),
    a = new Set(),
    s = new Set(),
    d = new Set();
  let u = 0;
  for (const externalSource of data.externalSources) {
    if (s.has(externalSource.id))
      throw Error("External source IDs must be unique.");
    (s.add(externalSource.id),
      (externalSource.url = validateExternalSourceUrl(externalSource.url)));
  }
  if (data.claims.reduce((l, c) => l + c.text.length, 0) > 2e4)
    throw Error("Answer claims exceed the 20,000 character evidence limit.");
  for (const claim of data.claims) {
    if (a.has(claim.id)) throw Error("Answer claim IDs must be unique.");
    if ((a.add(claim.id), claim.kind === "workbook")) {
      if (new Set(claim.ranges).size !== claim.ranges.length)
        throw Error("Duplicate citation ranges in a claim.");
      for (const range of claim.ranges) {
        const h = validateCitationRange(range);
        if (((u += h.length), u > answerGroundingLimits.referencedCells))
          throw Error(
            "Answer citations exceed the 1,000 cell-reference budget.",
          );
        if (h.some((f) => !o.has(f)))
          throw Error(
            "Answer citations must stay inside the approved selection.",
          );
      }
    } else if (claim.kind === "external") {
      if (
        new Set(claim.sourceIds).size !== claim.sourceIds.length ||
        claim.sourceIds.some((sourceId) => !s.has(sourceId))
      )
        throw Error(
          "External claims require distinct references to declared external sources.",
        );
      claim.sourceIds.forEach((sourceId) => d.add(sourceId));
    }
  }
  if (data.externalSources.some((externalSource) => !d.has(externalSource.id)))
    throw Error("Unused external sources are not accepted as grounding.");
  return data;
}
const nonVolatileFunctionNames = builtinFunctionNames.filter(
  (builtinFunctionName) =>
    ![
      "RAND",
      "RANDBETWEEN",
      "TODAY",
      "NOW",
      "INDIRECT",
      "OFFSET",
      "WEBSERVICE",
      "RTD",
      "HYPERLINK",
      "CALL",
    ].includes(builtinFunctionName),
);
const nonVolatileFunctionNameSet = new Set(nonVolatileFunctionNames);
const assertNoUnexpectedFields = (e, r) =>
  assertCondition(
    Object.keys(e).every((t) => r.includes(t)),
    "Unexpected field in LLM proposal.",
  );
function parseCellAddress(e) {
  const r = /^([A-Z]{1,3})([1-9]\d*)$/.exec(e);
  assertCondition(r, `Invalid cell address: ${e}`);
  const col = [...r[1]].reduce((o, a) => o * 26 + a.charCodeAt(0) - 64, 0),
    row = Number(r[2]);
  return (
    assertCondition(
      row <= workbookLimits.maxRows && col <= workbookLimits.maxCols,
      "Cell address exceeds supported dimensions.",
    ),
    { row: row, col: col }
  );
}
function columnNumberToLetters(col) {
  let r = "";
  for (; col > 0; col = Math.floor((col - 1) / 26))
    r = String.fromCharCode(65 + ((col - 1) % 26)) + r;
  return r;
}
function resolveAssistantSelection(workbook, r) {
  assertCondition(
    !r.ranges || (r.ranges.length === 1 && r.ranges[0] === r.range),
    "Select one rectangular range before using an assistant.",
  );
  const t = workbook.sheets.find((sheet) => sheet.id === r.sheetId);
  assertCondition(
    t && !t.hidden,
    "Select a visible worksheet before using an assistant.",
  );
  const n = r.range.toUpperCase().split(":");
  assertCondition(n.length <= 2, "Select one rectangular range.");
  const cellAddress = parseCellAddress(n[0]),
    a = parseCellAddress(n[1] || n[0]);
  (assertCondition(
    cellAddress.row <= a.row && cellAddress.col <= a.col,
    "Selection range must run from top left to bottom right.",
  ),
    assertCondition(
      a.row <= t.rowCount && a.col <= t.colCount,
      "Selection exceeds worksheet dimensions.",
    ),
    assertCondition(
      (a.row - cellAddress.row + 1) * (a.col - cellAddress.col + 1) <=
        workbookLimits.maxLLMCells,
      `Select at most ${workbookLimits.maxLLMCells} cells for the assistant.`,
    ));
  const s = [];
  for (let row = cellAddress.row; row <= a.row; row++)
    for (let col = cellAddress.col; col <= a.col; col++)
      s.push(`${columnNumberToLetters(col)}${row}`);
  return s;
}
function buildAssistantSelectionContext(workbook, r) {
  const assistantSelection = resolveAssistantSelection(workbook, r),
    n = workbook.sheets.find((sheet) => sheet.id === r.sheetId),
    workbookWithSettings = calculateWorkbookWithSettings(workbook);
  for (const a of assistantSelection) {
    const s = workbookWithSettings.getTypedValue(n.id, a);
    (assertCondition(
      !(s.type === "error" && s.value === "#PENDING!"),
      "Calculation is pending in the approved selection. Choose Calculate Now before sharing current values with an assistant.",
    ),
      assertCondition(
        !(
          s.type === "error" &&
          ["#TIMEOUT!", "#LIMIT!", "#BLOCKED!"].includes(String(s.value))
        ),
        "Calculation failed at a resource or execution boundary in the approved selection. Resolve the limit and choose Calculate Now before sharing current values with an assistant.",
      ));
  }
  return {
    workbookId: workbook.id,
    ...(workbook.datePolicy
      ? { datePolicy: normalizeDatePolicy(workbook.datePolicy) }
      : {}),
    ...(workbookUsesNowOrToday(workbook)
      ? { calculationClock: workbookWithSettings.getDateContext() }
      : {}),
    revision: workbook.revision,
    sheetId: n.id,
    sheetName: n.name,
    range: r.range.toUpperCase(),
    cells: assistantSelection.map((address) => ({
      address: address,
      raw: n.cells[address]?.raw ?? null,
      value: workbookWithSettings(n.id, address),
    })),
  };
}
function validateAssistantFormula(e, r) {
  let t = e.replace(/"(?:[^"]|"")*"/g, '""');
  if (
    (assertCondition(
      !/[!\[\]|\\#]/.test(t),
      "Assistant formulas cannot reference other worksheets, external workbooks, table names, existing spills or executable content.",
    ),
    assertCondition(
      !/(?:https?:|javascript:|data:|file:)/i.test(e),
      "Assistant formulas cannot contain network or executable URLs.",
    ),
    /\b(?:LET|LAMBDA)\s*\(/i.test(t))
  ) {
    const a = [];
    walkFormulaNodes(tokenizeFormulaText(t), (s, d, u, l) => {
      (l || (!s.qualifier && d.has(normalizeBindingName(s.name)))) &&
        a.push({ start: s.start, end: s.end, value: u ? "SUM" : "0" });
    });
    for (const s of a.sort((d, u) => u.start - d.start))
      t = t.slice(0, s.start) + s.value + t.slice(s.end);
  }
  const n = [...t.matchAll(/([A-Z_][A-Z0-9_.]*)\s*\(/gi)];
  for (const a of n)
    assertCondition(
      nonVolatileFunctionNameSet.has(a[1].toUpperCase()),
      `Assistant formula function ${a[1]} is not allowed.`,
    );
  for (const a of t.matchAll(
    new RegExp(
      "(?<![A-Z0-9_.])\\$?([A-Z]{1,3})\\$?([1-9]\\d*)(?![A-Z0-9_]|\\s*\\()",
      "gi",
    ),
  ))
    assertCondition(
      r.has(`${a[1].toUpperCase()}${a[2]}`),
      "Assistant formulas may only reference cells in the approved selection.",
    );
  const o = t
    .replace(/([A-Z_][A-Z0-9_.]*)\s*\(/gi, "(")
    .replace(/\$?[A-Z]{1,3}\$?[1-9]\d*(?![A-Z0-9_])/gi, "0")
    .replace(/""/g, "0")
    .replace(/\b(?:TRUE|FALSE)\b/gi, "0")
    .replace(/\d+(?:\.\d+)?(?:E[+-]?\d+)?/gi, "0");
  assertCondition(
    !/[A-Z_]/i.test(o),
    "Named ranges and unsupported formula tokens are not allowed in assistant changes.",
  );
}
function validateProposedCell(cell, r) {
  (assertCondition(
    isPlainRecord(cell),
    "Each proposed cell must be an object.",
  ),
    assertNoUnexpectedFields(cell, [
      "raw",
      "literal",
      "numFmt",
      "format",
      "style",
      "note",
      "tickmark",
    ]),
    assertCondition(
      Object.hasOwn(cell, "raw") &&
        (cell.raw === null ||
          typeof cell.raw == "string" ||
          typeof cell.raw == "boolean" ||
          (typeof cell.raw == "number" && Number.isFinite(cell.raw))),
      "Unsupported proposed cell value.",
    ),
    typeof cell.raw == "string" &&
      assertCondition(
        cell.raw.length <= (cell.raw.startsWith("=") ? 4096 : 32767),
        "Proposed cell text is too long.",
      ),
    cell.literal !== undefined &&
      assertCondition(
        typeof cell.literal == "boolean",
        "Cell literal must be boolean.",
      ),
    cell.numFmt !== undefined &&
      assertCondition(
        typeof cell.numFmt == "string" && cell.numFmt.length <= 200,
        "Invalid number format.",
      ),
    cell.format !== undefined &&
      assertCondition(
        ["general", "number", "currency", "percent", "date"].includes(
          String(cell.format),
        ),
        "Unsupported cell format.",
      ));
  for (const t of ["note", "tickmark"])
    cell[t] !== undefined &&
      assertCondition(
        typeof cell[t] == "string" &&
          String(cell[t]).length <= (t === "note" ? 1e4 : 40),
        `Invalid cell ${t}.`,
      );
  if (cell.style !== undefined) {
    (assertCondition(
      isPlainRecord(cell.style),
      "Cell style must be an object.",
    ),
      assertNoUnexpectedFields(cell.style, [
        "bold",
        "italic",
        "underline",
        "fontFamily",
        "fontSize",
        "wrap",
        "border",
        "fill",
        "color",
        "align",
      ]));
    for (const t of ["bold", "italic", "underline", "wrap"])
      cell.style[t] !== undefined &&
        assertCondition(
          typeof cell.style[t] == "boolean",
          "Invalid font style.",
        );
    (cell.style.fontFamily !== undefined &&
      assertCondition(
        supportedFontFamilies.includes(cell.style.fontFamily),
        "Unsupported font family.",
      ),
      cell.style.fontSize !== undefined &&
        assertCondition(
          typeof cell.style.fontSize == "number" &&
            Number.isInteger(cell.style.fontSize) &&
            cell.style.fontSize >= 8 &&
            cell.style.fontSize <= 72,
          "Font size must be an integer between 8 and 72.",
        ),
      cell.style.border !== undefined &&
        assertCondition(
          ["all", "bottom", "none"].includes(String(cell.style.border)),
          "Unsupported cell border.",
        ));
    for (const t of ["fill", "color"])
      cell.style[t] !== undefined &&
        assertCondition(
          typeof cell.style[t] == "string" &&
            /^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(String(cell.style[t])),
          "Use hexadecimal colors.",
        );
    cell.style.align !== undefined &&
      assertCondition(
        ["left", "center", "right"].includes(String(cell.style.align)),
        "Invalid alignment.",
      );
  }
  return (
    typeof cell.raw == "string" &&
      cell.raw.startsWith("=") &&
      !cell.literal &&
      validateAssistantFormula(cell.raw, r),
    structuredClone(cell)
  );
}
function validateAssistantProposal(e, r) {
  (assertCondition(
    isPlainRecord(e),
    "The assistant response must be a proposal object.",
  ),
    assertNoUnexpectedFields(e, [
      "id",
      "workbookId",
      "revision",
      "explanation",
      "citations",
      "patches",
      "grounding",
    ]),
    assertCondition(
      typeof e.id == "string" && e.id.length > 0 && e.id.length <= 100,
      "Invalid proposal identifier.",
    ),
    assertCondition(
      e.workbookId === r.workbookId && e.revision === r.revision,
      "This proposal is stale or belongs to another workbook. Generate a new proposal.",
    ),
    assertCondition(
      typeof e.explanation == "string" &&
        e.explanation.length > 0 &&
        e.explanation.length <= 1e4,
      "Invalid proposal explanation.",
    ));
  const t = new Set(r.cells.map((cell) => cell.address));
  (assertCondition(
    Array.isArray(e.citations) &&
      e.citations.length <= workbookLimits.maxLLMCells,
    "Invalid proposal citations.",
  ),
    assertCondition(
      e.citations.every(
        (citation) => typeof citation == "string" && t.has(citation),
      ),
      "Citations must reference the approved selection.",
    ),
    assertCondition(
      Array.isArray(e.patches) && e.patches.length <= workbookLimits.maxPatches,
      "Too many proposed edits.",
    ));
  const n = new Set(),
    patches = e.patches.map(
      (patch) => (
        assertCondition(isPlainRecord(patch), "Each patch must be an object."),
        assertNoUnexpectedFields(patch, ["sheetId", "address", "cell"]),
        assertCondition(
          patch.sheetId === r.sheetId &&
            typeof patch.address == "string" &&
            t.has(patch.address),
          "Assistant changes must stay inside the approved selection.",
        ),
        assertCondition(
          !n.has(patch.address),
          "The proposal contains duplicate cell edits.",
        ),
        n.add(patch.address),
        {
          sheetId: r.sheetId,
          address: patch.address,
          cell: validateProposedCell(patch.cell, t),
        }
      ),
    );
  return {
    id: e.id,
    workbookId: r.workbookId,
    revision: r.revision,
    explanation: e.explanation,
    citations: [...e.citations],
    patches: patches,
    ...(e.grounding !== undefined
      ? { grounding: validateAnswerGrounding(e.grounding, r) }
      : {}),
  };
}
function validateProposedCalculation(workbook, assistantSelection, t) {
  const cells = resolveAssistantSelection(workbook, assistantSelection).map(
      (address) => ({
        address: address,
        raw: null,
        value: null,
      }),
    ),
    o = validateAssistantProposal(t, {
      workbookId: workbook.id,
      revision: workbook.revision,
      sheetId: assistantSelection.sheetId,
      cells: cells,
    }),
    a = workbook.sheets.find(
      (sheet) => sheet.id === assistantSelection.sheetId,
    ),
    s = {
      ...o,
      patches: o.patches.map((patch) => {
        const cell = a.cells[patch.address],
          f = {
            ...cell,
            ...patch.cell,
            ...(cell?.style || patch.cell.style
              ? {
                  style: { ...cell?.style, ...patch.cell.style },
                }
              : {}),
          };
        (cell?.raw !== patch.cell.raw &&
          patch.cell.literal === undefined &&
          delete f.literal,
          patch.cell.format &&
            patch.cell.format !== cell?.format &&
            patch.cell.numFmt === undefined &&
            delete f.numFmt);
        const cellAddress = parseCellAddress(patch.address),
          m = (a.tables ?? []).some((v) => {
            const [g, b] = v.range.split(":"),
              y = parseCellAddress(g),
              w = parseCellAddress(b ?? g);
            return (
              cellAddress.row === y.row &&
              cellAddress.col >= y.col &&
              cellAddress.col <= w.col
            );
          });
        return (
          assertCondition(
            !m || (cell?.raw === f.raw && !!cell?.literal == !!f.literal),
            "Assistant changes cannot rename table headers because dependent references may change outside the approved selection.",
          ),
          { ...patch, cell: f }
        );
      }),
    },
    d = new Set(cells.map((cell) => cell.address)),
    u = {
      ...workbook,
      sheets: workbook.sheets.map((sheet) =>
        sheet.id === a.id
          ? {
              ...sheet,
              cells: {
                ...sheet.cells,
                ...Object.fromEntries(
                  s.patches.map((patch) => [patch.address, patch.cell]),
                ),
              },
            }
          : sheet,
      ),
    },
    l = evaluateWorkbookFormulas(
      u,
      {},
      {
        allowRead: () => true,
        allowDatabaseCriterion: (c, h) => {
          if (c !== a.id) return false;
          try {
            return (validateAssistantFormula(h, d), true);
          } catch {
            return false;
          }
        },
        step: () => {},
      },
    );
  for (const patch of s.patches) {
    if (
      typeof patch.cell.raw != "string" ||
      !patch.cell.raw.startsWith("=") ||
      patch.cell.literal
    )
      continue;
    const h = l(a.id, patch.address);
    (assertCondition(
      h !== "#LIMIT!",
      "The proposed formula exceeds calculation limits.",
    ),
      assertCondition(
        l.getOperationalErrors().length === 0,
        "The proposed formula has invalid syntax, a circular reference or an operational calculation failure. Correct the formula before review; error handlers cannot suppress this rejection.",
      ));
    const f = l.getSpillRange(a.id, patch.address);
    if (
      (assertCondition(
        !l.getSpillDiagnostic(a.id, patch.address),
        "The proposed array is blocked. Review the spill obstruction before requesting approval.",
      ),
      f)
    ) {
      const [p, m] = f.split(":"),
        cellAddress = parseCellAddress(p),
        g = parseCellAddress(m ?? p);
      assertCondition(
        (g.row - cellAddress.row + 1) * (g.col - cellAddress.col + 1) <=
          workbookLimits.maxLLMCells,
        "Assistant spill outputs must stay inside the approved selection.",
      );
      for (let row = cellAddress.row; row <= g.row; row++)
        for (let col = cellAddress.col; col <= g.col; col++)
          assertCondition(
            d.has(`${columnNumberToLetters(col)}${row}`),
            "Assistant spill outputs must stay inside the approved selection.",
          );
    }
  }
  return (
    assertCondition(
      l.getOperationalErrors().length === 0,
      "The proposed formula cannot be validated because calculation encountered an operational failure.",
    ),
    s
  );
}
function buildLocalDemoProposal() {
  return {
    name: "Local demo · no model or network",
    localOnly: true,
    chat: { history: true, attachments: true },
    async propose(e, r, t, n) {
      if (t?.aborted)
        throw new DOMException("Request cancelled.", "AbortError");
      const o = {
          id: `local-${r.revision}-${Date.now()}`,
          workbookId: r.workbookId,
          revision: r.revision,
          explanation: "",
          citations: [],
          patches: [],
        },
        a = r.cells.filter(
          (cell) =>
            typeof cell.value == "number" && Number.isFinite(cell.value),
        ),
        s = a.reduce((d, u) => d + Number(u.value), 0);
      if (/highlight|high values|outlier/i.test(e)) {
        const d =
            /(?:above|over|greater than|exceed(?:ing|s)?)\s*[$€£]?\s*(-?\d[\d,]*(?:\.\d+)?)/i.exec(
              e,
            ),
          u = d
            ? Number(d[1].replaceAll(",", ""))
            : a.length
              ? Math.abs(s / a.length) * 1.5
              : 0,
          l = a.filter(
            (c) => (d ? Number(c.value) : Math.abs(Number(c.value))) > u,
          );
        ((o.explanation = `Local deterministic demo: ${l.length} values exceed ${d ? `the requested threshold of ${u.toLocaleString()}` : `1.5 times the absolute arithmetic mean (${u.toLocaleString()})`}. This is a simple illustration, not a statistical audit conclusion. The listed cells receive an amber fill; existing annotations are retained.`),
          (o.citations = l.map((c) => c.address)),
          (o.patches = l.map((cell) => ({
            sheetId: r.sheetId,
            address: cell.address,
            cell: { raw: cell.raw, style: { fill: "#fff0b3" } },
          }))));
      } else if (/\b(?:total|sum)\b/i.test(e)) {
        const d = [...r.cells]
          .reverse()
          .find((cell) => cell.raw === null || cell.raw === "");
        if (d && a.length) {
          const raw = `=SUM(${a.map((l) => l.address).join(",")})`;
          ((o.patches = [
            {
              sheetId: r.sheetId,
              address: d.address,
              cell: { raw: raw, format: "number", style: { bold: true } },
            },
          ]),
            (o.citations = a.map((l) => l.address)),
            (o.explanation = `Local deterministic demo: put a SUM formula for ${a.length} numeric selected cells in empty cell ${d.address}. Current result: ${s.toLocaleString()}. Review the included addresses for subtotals or duplicates before accepting.`));
        } else
          o.explanation =
            "Local deterministic demo: select numeric cells and at least one empty destination cell to propose a total.";
      } else if (/explain/i.test(e)) {
        const d = r.cells.filter(
          (cell) => typeof cell.raw == "string" && cell.raw.startsWith("="),
        );
        ((o.explanation = d.length
          ? `Local deterministic demo: ${d.map((cell) => `${cell.address} contains ${cell.raw} and currently returns ${String(cell.value)}`).join("; ")}. This lists formulas and results; no language model interpretation is used.`
          : "Local deterministic demo: the selection contains no formula expressions. Values and formulas are worksheet data, never instructions."),
          (o.citations = d.map((u) => u.address)));
      } else
        ((o.explanation = `Local deterministic demo: ${r.cells.length} selected cells; ${a.length} numeric values; sum ${s.toLocaleString()}; ${r.cells.filter((cell) => cell.raw === null || cell.raw === "").length} empty cells. Workbook content is treated only as data.`),
          (o.citations = a.slice(0, 20).map((d) => d.address)));
      if (
        ((o.grounding = {
          claims: [
            {
              id: "local-evidence",
              kind: "workbook",
              text: o.explanation,
              ranges: [r.range],
            },
          ],
          externalSources: [],
        }),
        n?.attachments.length)
      ) {
        const attachments = parseAssistantChatRequest(n).attachments,
          u = `Received ${attachments.length} attachment${attachments.length === 1 ? "" : "s"}: ${attachments.map((attachment) => `${attachment.name} (${attachment.size.toLocaleString()} bytes)`).join(", ")}. Files remain in this browser in local demo mode. Preview text and images in the attachment cards; document interpretation requires a connected model.`;
        ((o.explanation += `

${u}`),
          o.grounding.claims.push({
            id: "local-attachments",
            kind: "commentary",
            text: u,
          }));
      }
      return validateAssistantProposal(o, r);
    },
  };
}

/* From 11-local-workbook-store-and-versions.js */
const isPlainRecord = (e) =>
  !!e &&
  typeof e == "object" &&
  !Array.isArray(e) &&
  Object.getPrototypeOf(e) === Object.prototype;

/* From 12-consolidation-and-dependency-inspection.js */
const hasOnlyAllowedKeys = (e, r) =>
  !!e &&
  typeof e == "object" &&
  !Array.isArray(e) &&
  Object.keys(e).every((t) => r.includes(t));

export {
  LLMRequestError,
  Pk,
  answerGroundingLimits,
  answerGroundingSchema,
  assertCondition,
  assertNoUnexpectedFields,
  assistantErrorMessages,
  assistantRequestLimits,
  attachmentMediaTypes,
  buildAssistantSelectionContext,
  buildLocalDemoProposal,
  columnNumberToLetters,
  getFileExtension,
  hasOnlyAllowedKeys,
  identifierSchema,
  isPlainRecord,
  isSafeAttachmentFilename,
  matchesFileSignature,
  nonVolatileFunctionNameSet,
  nonVolatileFunctionNames,
  parseAssistantChatRequest,
  parseCellAddress,
  rangeAddressSchema,
  resolveAssistantSelection,
  textAttachmentExtensions,
  validateAnswerGrounding,
  validateAssistantFormula,
  validateAssistantProposal,
  validateCitationRange,
  validateExternalSourceUrl,
  validateProposedCalculation,
  validateProposedCell,
};
