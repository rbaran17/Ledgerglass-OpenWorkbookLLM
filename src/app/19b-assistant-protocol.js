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
function matchesFileSignature(e, r) {
  const t = (...n) => n.every((o, a) => e[a] === o);
  return r === "pdf"
    ? t(37, 80, 68, 70, 45)
    : r === "png"
      ? t(137, 80, 78, 71, 13, 10, 26, 10)
      : r === "jpg" || r === "jpeg"
        ? t(255, 216, 255)
        : r === "webp"
          ? t(82, 73, 70, 70) &&
            String.fromCharCode(...e.slice(8, 12)) === "WEBP"
          : r === "xlsx" || r === "docx"
            ? t(80, 75, 3, 4)
            : textAttachmentExtensions.has(r);
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
      let l;
      return (
        u
          ? (assertCondition(
              !d.data.includes("\0"),
              "Invalid text attachment.",
            ),
            (l = new TextEncoder().encode(d.data)))
          : (assertCondition(
              d.data.length % 4 === 0 && /^[A-Za-z0-9+/]*={0,2}$/.test(d.data),
              "Invalid attachment encoding.",
            ),
            (l = Uint8Array.from(atob(d.data), (c) => c.charCodeAt(0)))),
        assertCondition(
          l.length === d.size &&
            matchesFileSignature(l, getFileExtension(d.name)),
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
      (a) => (
        assertCondition(
          hasOnlyAllowedKeys(a, ["role", "text", "attachments"]) &&
            (a.role === "user" || a.role === "assistant") &&
            typeof a.text == "string" &&
            a.text.length <= 1e4,
          "Invalid chat history message.",
        ),
        assertCondition(
          a.role === "user" || a.attachments === undefined,
          "Assistant messages cannot attach files.",
        ),
        (t += a.text.length),
        assertCondition(
          t <= assistantRequestLimits.historyCharacters,
          "This conversation has reached its context limit. Start a new chat to continue.",
        ),
        {
          role: a.role,
          text: a.text,
          ...(a.attachments ? { attachments: n(a.attachments) } : {}),
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
function validateCitationRange(e) {
  if (!/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/.test(e))
    throw Error("Citation requires an explicit A1 cell or rectangle.");
  const r = parseRangeAddress(e),
    t = e.split(":");
  if (t[1] && `${columnIndexToLetters(r.firstCol)}${r.firstRow + 1}` !== t[0])
    throw Error("Citation range must run top left to bottom right.");
  if ((r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1) > 500)
    throw Error("Citations support at most 500 cells per range.");
  const n = [];
  for (let o = r.firstRow; o <= r.lastRow; o++)
    for (let a = r.firstCol; a <= r.lastCol; a++)
      n.push(`${columnIndexToLetters(a)}${o + 1}`);
  return n;
}
function validateExternalSourceUrl(e) {
  if (
    typeof e != "string" ||
    e.length > 2048 ||
    e !== e.trim() ||
    /[\u0000-\u0020\u007f]/.test(e)
  )
    throw Error(
      "External sources require a valid HTTPS URL without credentials.",
    );
  let r;
  try {
    r = new URL(e);
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
function validateAnswerGrounding(e, r) {
  const t = answerGroundingSchema.safeParse(e);
  if (!t.success)
    throw Error(
      "Invalid answer grounding: provide typed claims and separate external sources.",
    );
  const n = t.data,
    o = new Set(r.cells.map((l) => l.address)),
    a = new Set(),
    s = new Set(),
    d = new Set();
  let u = 0;
  for (const l of n.externalSources) {
    if (s.has(l.id)) throw Error("External source IDs must be unique.");
    (s.add(l.id), (l.url = validateExternalSourceUrl(l.url)));
  }
  if (n.claims.reduce((l, c) => l + c.text.length, 0) > 2e4)
    throw Error("Answer claims exceed the 20,000 character evidence limit.");
  for (const l of n.claims) {
    if (a.has(l.id)) throw Error("Answer claim IDs must be unique.");
    if ((a.add(l.id), l.kind === "workbook")) {
      if (new Set(l.ranges).size !== l.ranges.length)
        throw Error("Duplicate citation ranges in a claim.");
      for (const c of l.ranges) {
        const h = validateCitationRange(c);
        if (((u += h.length), u > answerGroundingLimits.referencedCells))
          throw Error(
            "Answer citations exceed the 1,000 cell-reference budget.",
          );
        if (h.some((f) => !o.has(f)))
          throw Error(
            "Answer citations must stay inside the approved selection.",
          );
      }
    } else if (l.kind === "external") {
      if (
        new Set(l.sourceIds).size !== l.sourceIds.length ||
        l.sourceIds.some((c) => !s.has(c))
      )
        throw Error(
          "External claims require distinct references to declared external sources.",
        );
      l.sourceIds.forEach((c) => d.add(c));
    }
  }
  if (n.externalSources.some((l) => !d.has(l.id)))
    throw Error("Unused external sources are not accepted as grounding.");
  return n;
}
const nonVolatileFunctionNames = builtinFunctionNames.filter(
  (e) =>
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
    ].includes(e),
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
  const t = [...r[1]].reduce((o, a) => o * 26 + a.charCodeAt(0) - 64, 0),
    n = Number(r[2]);
  return (
    assertCondition(
      n <= workbookLimits.maxRows && t <= workbookLimits.maxCols,
      "Cell address exceeds supported dimensions.",
    ),
    { row: n, col: t }
  );
}
function columnNumberToLetters(e) {
  let r = "";
  for (; e > 0; e = Math.floor((e - 1) / 26))
    r = String.fromCharCode(65 + ((e - 1) % 26)) + r;
  return r;
}
function resolveAssistantSelection(e, r) {
  assertCondition(
    !r.ranges || (r.ranges.length === 1 && r.ranges[0] === r.range),
    "Select one rectangular range before using an assistant.",
  );
  const t = e.sheets.find((d) => d.id === r.sheetId);
  assertCondition(
    t && !t.hidden,
    "Select a visible worksheet before using an assistant.",
  );
  const n = r.range.toUpperCase().split(":");
  assertCondition(n.length <= 2, "Select one rectangular range.");
  const o = parseCellAddress(n[0]),
    a = parseCellAddress(n[1] || n[0]);
  (assertCondition(
    o.row <= a.row && o.col <= a.col,
    "Selection range must run from top left to bottom right.",
  ),
    assertCondition(
      a.row <= t.rowCount && a.col <= t.colCount,
      "Selection exceeds worksheet dimensions.",
    ),
    assertCondition(
      (a.row - o.row + 1) * (a.col - o.col + 1) <= workbookLimits.maxLLMCells,
      `Select at most ${workbookLimits.maxLLMCells} cells for the assistant.`,
    ));
  const s = [];
  for (let d = o.row; d <= a.row; d++)
    for (let u = o.col; u <= a.col; u++)
      s.push(`${columnNumberToLetters(u)}${d}`);
  return s;
}
function buildAssistantSelectionContext(e, r) {
  const t = resolveAssistantSelection(e, r),
    n = e.sheets.find((a) => a.id === r.sheetId),
    o = calculateWorkbookWithSettings(e);
  for (const a of t) {
    const s = o.getTypedValue(n.id, a);
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
    workbookId: e.id,
    ...(e.datePolicy ? { datePolicy: normalizeDatePolicy(e.datePolicy) } : {}),
    ...(workbookUsesNowOrToday(e)
      ? { calculationClock: o.getDateContext() }
      : {}),
    revision: e.revision,
    sheetId: n.id,
    sheetName: n.name,
    range: r.range.toUpperCase(),
    cells: t.map((a) => ({
      address: a,
      raw: n.cells[a]?.raw ?? null,
      value: o(n.id, a),
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
function validateProposedCell(e, r) {
  (assertCondition(isPlainRecord(e), "Each proposed cell must be an object."),
    assertNoUnexpectedFields(e, [
      "raw",
      "literal",
      "numFmt",
      "format",
      "style",
      "note",
      "tickmark",
    ]),
    assertCondition(
      Object.hasOwn(e, "raw") &&
        (e.raw === null ||
          typeof e.raw == "string" ||
          typeof e.raw == "boolean" ||
          (typeof e.raw == "number" && Number.isFinite(e.raw))),
      "Unsupported proposed cell value.",
    ),
    typeof e.raw == "string" &&
      assertCondition(
        e.raw.length <= (e.raw.startsWith("=") ? 4096 : 32767),
        "Proposed cell text is too long.",
      ),
    e.literal !== undefined &&
      assertCondition(
        typeof e.literal == "boolean",
        "Cell literal must be boolean.",
      ),
    e.numFmt !== undefined &&
      assertCondition(
        typeof e.numFmt == "string" && e.numFmt.length <= 200,
        "Invalid number format.",
      ),
    e.format !== undefined &&
      assertCondition(
        ["general", "number", "currency", "percent", "date"].includes(
          String(e.format),
        ),
        "Unsupported cell format.",
      ));
  for (const t of ["note", "tickmark"])
    e[t] !== undefined &&
      assertCondition(
        typeof e[t] == "string" &&
          String(e[t]).length <= (t === "note" ? 1e4 : 40),
        `Invalid cell ${t}.`,
      );
  if (e.style !== undefined) {
    (assertCondition(isPlainRecord(e.style), "Cell style must be an object."),
      assertNoUnexpectedFields(e.style, [
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
      e.style[t] !== undefined &&
        assertCondition(typeof e.style[t] == "boolean", "Invalid font style.");
    (e.style.fontFamily !== undefined &&
      assertCondition(
        supportedFontFamilies.includes(e.style.fontFamily),
        "Unsupported font family.",
      ),
      e.style.fontSize !== undefined &&
        assertCondition(
          typeof e.style.fontSize == "number" &&
            Number.isInteger(e.style.fontSize) &&
            e.style.fontSize >= 8 &&
            e.style.fontSize <= 72,
          "Font size must be an integer between 8 and 72.",
        ),
      e.style.border !== undefined &&
        assertCondition(
          ["all", "bottom", "none"].includes(String(e.style.border)),
          "Unsupported cell border.",
        ));
    for (const t of ["fill", "color"])
      e.style[t] !== undefined &&
        assertCondition(
          typeof e.style[t] == "string" &&
            /^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(String(e.style[t])),
          "Use hexadecimal colors.",
        );
    e.style.align !== undefined &&
      assertCondition(
        ["left", "center", "right"].includes(String(e.style.align)),
        "Invalid alignment.",
      );
  }
  return (
    typeof e.raw == "string" &&
      e.raw.startsWith("=") &&
      !e.literal &&
      validateAssistantFormula(e.raw, r),
    structuredClone(e)
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
  const t = new Set(r.cells.map((a) => a.address));
  (assertCondition(
    Array.isArray(e.citations) &&
      e.citations.length <= workbookLimits.maxLLMCells,
    "Invalid proposal citations.",
  ),
    assertCondition(
      e.citations.every((a) => typeof a == "string" && t.has(a)),
      "Citations must reference the approved selection.",
    ),
    assertCondition(
      Array.isArray(e.patches) && e.patches.length <= workbookLimits.maxPatches,
      "Too many proposed edits.",
    ));
  const n = new Set(),
    o = e.patches.map(
      (a) => (
        assertCondition(isPlainRecord(a), "Each patch must be an object."),
        assertNoUnexpectedFields(a, ["sheetId", "address", "cell"]),
        assertCondition(
          a.sheetId === r.sheetId &&
            typeof a.address == "string" &&
            t.has(a.address),
          "Assistant changes must stay inside the approved selection.",
        ),
        assertCondition(
          !n.has(a.address),
          "The proposal contains duplicate cell edits.",
        ),
        n.add(a.address),
        {
          sheetId: r.sheetId,
          address: a.address,
          cell: validateProposedCell(a.cell, t),
        }
      ),
    );
  return {
    id: e.id,
    workbookId: r.workbookId,
    revision: r.revision,
    explanation: e.explanation,
    citations: [...e.citations],
    patches: o,
    ...(e.grounding !== undefined
      ? { grounding: validateAnswerGrounding(e.grounding, r) }
      : {}),
  };
}
function validateProposedCalculation(e, r, t) {
  const n = resolveAssistantSelection(e, r).map((c) => ({
      address: c,
      raw: null,
      value: null,
    })),
    o = validateAssistantProposal(t, {
      workbookId: e.id,
      revision: e.revision,
      sheetId: r.sheetId,
      cells: n,
    }),
    a = e.sheets.find((c) => c.id === r.sheetId),
    s = {
      ...o,
      patches: o.patches.map((c) => {
        const h = a.cells[c.address],
          f = {
            ...h,
            ...c.cell,
            ...(h?.style || c.cell.style
              ? {
                  style: { ...h?.style, ...c.cell.style },
                }
              : {}),
          };
        (h?.raw !== c.cell.raw &&
          c.cell.literal === undefined &&
          delete f.literal,
          c.cell.format &&
            c.cell.format !== h?.format &&
            c.cell.numFmt === undefined &&
            delete f.numFmt);
        const p = parseCellAddress(c.address),
          m = (a.tables ?? []).some((v) => {
            const [g, b] = v.range.split(":"),
              y = parseCellAddress(g),
              w = parseCellAddress(b ?? g);
            return p.row === y.row && p.col >= y.col && p.col <= w.col;
          });
        return (
          assertCondition(
            !m || (h?.raw === f.raw && !!h?.literal == !!f.literal),
            "Assistant changes cannot rename table headers because dependent references may change outside the approved selection.",
          ),
          { ...c, cell: f }
        );
      }),
    },
    d = new Set(n.map((c) => c.address)),
    u = {
      ...e,
      sheets: e.sheets.map((c) =>
        c.id === a.id
          ? {
              ...c,
              cells: {
                ...c.cells,
                ...Object.fromEntries(
                  s.patches.map((h) => [h.address, h.cell]),
                ),
              },
            }
          : c,
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
  for (const c of s.patches) {
    if (
      typeof c.cell.raw != "string" ||
      !c.cell.raw.startsWith("=") ||
      c.cell.literal
    )
      continue;
    const h = l(a.id, c.address);
    (assertCondition(
      h !== "#LIMIT!",
      "The proposed formula exceeds calculation limits.",
    ),
      assertCondition(
        l.getOperationalErrors().length === 0,
        "The proposed formula has invalid syntax, a circular reference or an operational calculation failure. Correct the formula before review; error handlers cannot suppress this rejection.",
      ));
    const f = l.getSpillRange(a.id, c.address);
    if (
      (assertCondition(
        !l.getSpillDiagnostic(a.id, c.address),
        "The proposed array is blocked. Review the spill obstruction before requesting approval.",
      ),
      f)
    ) {
      const [p, m] = f.split(":"),
        v = parseCellAddress(p),
        g = parseCellAddress(m ?? p);
      assertCondition(
        (g.row - v.row + 1) * (g.col - v.col + 1) <= workbookLimits.maxLLMCells,
        "Assistant spill outputs must stay inside the approved selection.",
      );
      for (let b = v.row; b <= g.row; b++)
        for (let y = v.col; y <= g.col; y++)
          assertCondition(
            d.has(`${columnNumberToLetters(y)}${b}`),
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
          (d) => typeof d.value == "number" && Number.isFinite(d.value),
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
          (o.patches = l.map((c) => ({
            sheetId: r.sheetId,
            address: c.address,
            cell: { raw: c.raw, style: { fill: "#fff0b3" } },
          }))));
      } else if (/\b(?:total|sum)\b/i.test(e)) {
        const d = [...r.cells]
          .reverse()
          .find((u) => u.raw === null || u.raw === "");
        if (d && a.length) {
          const u = `=SUM(${a.map((l) => l.address).join(",")})`;
          ((o.patches = [
            {
              sheetId: r.sheetId,
              address: d.address,
              cell: { raw: u, format: "number", style: { bold: true } },
            },
          ]),
            (o.citations = a.map((l) => l.address)),
            (o.explanation = `Local deterministic demo: put a SUM formula for ${a.length} numeric selected cells in empty cell ${d.address}. Current result: ${s.toLocaleString()}. Review the included addresses for subtotals or duplicates before accepting.`));
        } else
          o.explanation =
            "Local deterministic demo: select numeric cells and at least one empty destination cell to propose a total.";
      } else if (/explain/i.test(e)) {
        const d = r.cells.filter(
          (u) => typeof u.raw == "string" && u.raw.startsWith("="),
        );
        ((o.explanation = d.length
          ? `Local deterministic demo: ${d.map((u) => `${u.address} contains ${u.raw} and currently returns ${String(u.value)}`).join("; ")}. This lists formulas and results; no language model interpretation is used.`
          : "Local deterministic demo: the selection contains no formula expressions. Values and formulas are worksheet data, never instructions."),
          (o.citations = d.map((u) => u.address)));
      } else
        ((o.explanation = `Local deterministic demo: ${r.cells.length} selected cells; ${a.length} numeric values; sum ${s.toLocaleString()}; ${r.cells.filter((d) => d.raw === null || d.raw === "").length} empty cells. Workbook content is treated only as data.`),
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
        const d = parseAssistantChatRequest(n).attachments,
          u = `Received ${d.length} attachment${d.length === 1 ? "" : "s"}: ${d.map((l) => `${l.name} (${l.size.toLocaleString()} bytes)`).join(", ")}. Files remain in this browser in local demo mode. Preview text and images in the attachment cards; document interpretation requires a connected model.`;
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
