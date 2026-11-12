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
const getFileExtension = (fileName) =>
  fileName.split(".").pop()?.toLowerCase() ?? "";
function assertCondition(condition, message) {
  if (!condition) throw new Error(message);
}
function isSafeAttachmentFilename(fileName) {
  return (
    typeof fileName == "string" &&
    fileName.length > 0 &&
    fileName.length <= 180 &&
    !/[\x00-\x1f\x7f/\\]/.test(fileName)
  );
}
function matchesFileSignature(attachmentBytes, fileExtension) {
  const startsWithBytes = (...expectedBytes) =>
    expectedBytes.every(
      (expectedByte, byteIndex) => attachmentBytes[byteIndex] === expectedByte,
    );
  return fileExtension === "pdf"
    ? startsWithBytes(37, 80, 68, 70, 45)
    : fileExtension === "png"
      ? startsWithBytes(137, 80, 78, 71, 13, 10, 26, 10)
      : fileExtension === "jpg" || fileExtension === "jpeg"
        ? startsWithBytes(255, 216, 255)
        : fileExtension === "webp"
          ? startsWithBytes(82, 73, 70, 70) &&
            String.fromCharCode(...attachmentBytes.slice(8, 12)) === "WEBP"
          : fileExtension === "xlsx" || fileExtension === "docx"
            ? startsWithBytes(80, 75, 3, 4)
            : textAttachmentExtensions.has(fileExtension);
}
function parseAssistantChatRequest(request) {
  (assertCondition(
    hasOnlyAllowedKeys(request, ["version", "messages", "attachments"]) &&
      request.version === 1 &&
      Array.isArray(request.messages) &&
      Array.isArray(request.attachments),
    "Invalid chat request.",
  ),
    assertCondition(
      request.messages.length <= assistantRequestLimits.historyMessages,
      "This conversation has reached 20 prior messages. Start a new chat to continue.",
    ));
  let totalAttachmentBytes = 0,
    historyCharacterCount = 0;
  const parseAttachments = (attachments) => {
    assertCondition(
      Array.isArray(attachments) &&
        attachments.length <= assistantRequestLimits.filesPerMessage,
      "Attach up to 4 files per message.",
    );
    const seenIds = new Set();
    return attachments.map((attachment) => {
      (assertCondition(
        hasOnlyAllowedKeys(attachment, [
          "id",
          "name",
          "mediaType",
          "size",
          "encoding",
          "data",
        ]) &&
          typeof attachment.id == "string" &&
          attachment.id.length > 0 &&
          attachment.id.length <= 100 &&
          !seenIds.has(attachment.id),
        "Invalid or duplicate attachment.",
      ),
        seenIds.add(attachment.id),
        assertCondition(
          isSafeAttachmentFilename(attachment.name) &&
            Object.hasOwn(
              attachmentMediaTypes,
              getFileExtension(attachment.name),
            ) &&
            attachment.mediaType ===
              attachmentMediaTypes[getFileExtension(attachment.name)],
          "Unsupported attachment type.",
        ),
        assertCondition(
          Number.isSafeInteger(attachment.size) &&
            Number(attachment.size) > 0 &&
            Number(attachment.size) <= assistantRequestLimits.fileBytes &&
            typeof attachment.data == "string" &&
            attachment.data.length <=
              Math.ceil(assistantRequestLimits.fileBytes / 3) * 4,
          "Attachment exceeds the 5 MB limit.",
        ));
      const isTextAttachment = textAttachmentExtensions.has(
        getFileExtension(attachment.name),
      );
      assertCondition(
        attachment.encoding === (isTextAttachment ? "utf8" : "base64"),
        "Invalid attachment encoding.",
      );
      let attachmentBytes;
      return (
        isTextAttachment
          ? (assertCondition(
              !attachment.data.includes("\0"),
              "Invalid text attachment.",
            ),
            (attachmentBytes = new TextEncoder().encode(attachment.data)))
          : (assertCondition(
              attachment.data.length % 4 === 0 &&
                /^[A-Za-z0-9+/]*={0,2}$/.test(attachment.data),
              "Invalid attachment encoding.",
            ),
            (attachmentBytes = Uint8Array.from(
              atob(attachment.data),
              (character) => character.charCodeAt(0),
            ))),
        assertCondition(
          attachmentBytes.length === attachment.size &&
            matchesFileSignature(
              attachmentBytes,
              getFileExtension(attachment.name),
            ),
          "Attachment content or size is invalid.",
        ),
        (totalAttachmentBytes += Number(attachment.size)),
        assertCondition(
          totalAttachmentBytes <= assistantRequestLimits.totalFileBytes,
          "The conversation contains more than 10 MB of attachments. Remove files or start a new chat.",
        ),
        {
          id: attachment.id,
          name: attachment.name,
          mediaType: attachment.mediaType,
          size: Number(attachment.size),
          encoding: attachment.encoding,
          data: attachment.data,
        }
      );
    });
  };
  return {
    version: 1,
    messages: request.messages.map(
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
        (historyCharacterCount += message.text.length),
        assertCondition(
          historyCharacterCount <= assistantRequestLimits.historyCharacters,
          "This conversation has reached its context limit. Start a new chat to continue.",
        ),
        {
          role: message.role,
          text: message.text,
          ...(message.attachments
            ? { attachments: parseAttachments(message.attachments) }
            : {}),
        }
      ),
    ),
    attachments: parseAttachments(request.attachments),
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
  constructor(requestedCode) {
    const resolvedCode =
      typeof requestedCode == "string" &&
      Object.hasOwn(assistantErrorMessages, requestedCode)
        ? requestedCode
        : "unavailable";
    super(assistantErrorMessages[resolvedCode]);
    defineDataPropertyByKey(this, "code");
    ((this.name = "LLMRequestError"), (this.code = resolvedCode));
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
    rangeParts = range.split(":");
  if (
    rangeParts[1] &&
    `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}` !==
      rangeParts[0]
  )
    throw Error("Citation range must run top left to bottom right.");
  if (
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    500
  )
    throw Error("Citations support at most 500 cells per range.");
  const cellAddresses = [];
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
      cellAddresses.push(`${columnIndexToLetters(firstCol)}${firstRow + 1}`);
  return cellAddresses;
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
  let parsedUrl;
  try {
    parsedUrl = new URL(url);
  } catch {
    throw Error(
      "External sources require a valid HTTPS URL without credentials.",
    );
  }
  if (
    parsedUrl.protocol !== "https:" ||
    parsedUrl.username ||
    parsedUrl.password ||
    !parsedUrl.hostname
  )
    throw Error(
      "External sources require a valid HTTPS URL without credentials.",
    );
  return parsedUrl.href;
}
function validateAnswerGrounding(grounding, selectionContext) {
  const parseResult = answerGroundingSchema.safeParse(grounding);
  if (!parseResult.success)
    throw Error(
      "Invalid answer grounding: provide typed claims and separate external sources.",
    );
  const data = parseResult.data,
    selectionAddresses = new Set(
      selectionContext.cells.map((cell) => cell.address),
    ),
    claimIds = new Set(),
    externalSourceIds = new Set(),
    referencedSourceIds = new Set();
  let referencedCellCount = 0;
  for (const externalSource of data.externalSources) {
    if (externalSourceIds.has(externalSource.id))
      throw Error("External source IDs must be unique.");
    (externalSourceIds.add(externalSource.id),
      (externalSource.url = validateExternalSourceUrl(externalSource.url)));
  }
  if (data.claims.reduce((total, claim) => total + claim.text.length, 0) > 2e4)
    throw Error("Answer claims exceed the 20,000 character evidence limit.");
  for (const claim of data.claims) {
    if (claimIds.has(claim.id)) throw Error("Answer claim IDs must be unique.");
    if ((claimIds.add(claim.id), claim.kind === "workbook")) {
      if (new Set(claim.ranges).size !== claim.ranges.length)
        throw Error("Duplicate citation ranges in a claim.");
      for (const range of claim.ranges) {
        const rangeCellAddresses = validateCitationRange(range);
        if (
          ((referencedCellCount += rangeCellAddresses.length),
          referencedCellCount > answerGroundingLimits.referencedCells)
        )
          throw Error(
            "Answer citations exceed the 1,000 cell-reference budget.",
          );
        if (
          rangeCellAddresses.some((address) => !selectionAddresses.has(address))
        )
          throw Error(
            "Answer citations must stay inside the approved selection.",
          );
      }
    } else if (claim.kind === "external") {
      if (
        new Set(claim.sourceIds).size !== claim.sourceIds.length ||
        claim.sourceIds.some((sourceId) => !externalSourceIds.has(sourceId))
      )
        throw Error(
          "External claims require distinct references to declared external sources.",
        );
      claim.sourceIds.forEach((sourceId) => referencedSourceIds.add(sourceId));
    }
  }
  if (
    data.externalSources.some(
      (externalSource) => !referencedSourceIds.has(externalSource.id),
    )
  )
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
const assertNoUnexpectedFields = (record, allowedFields) =>
  assertCondition(
    Object.keys(record).every((fieldName) => allowedFields.includes(fieldName)),
    "Unexpected field in LLM proposal.",
  );
function parseCellAddress(address) {
  const match = /^([A-Z]{1,3})([1-9]\d*)$/.exec(address);
  assertCondition(match, `Invalid cell address: ${address}`);
  const col = [...match[1]].reduce(
      (total, character) => total * 26 + character.charCodeAt(0) - 64,
      0,
    ),
    row = Number(match[2]);
  return (
    assertCondition(
      row <= workbookLimits.maxRows && col <= workbookLimits.maxCols,
      "Cell address exceeds supported dimensions.",
    ),
    { row: row, col: col }
  );
}
function columnNumberToLetters(col) {
  let letters = "";
  for (; col > 0; col = Math.floor((col - 1) / 26))
    letters = String.fromCharCode(65 + ((col - 1) % 26)) + letters;
  return letters;
}
function resolveAssistantSelection(workbook, selection) {
  assertCondition(
    !selection.ranges ||
      (selection.ranges.length === 1 &&
        selection.ranges[0] === selection.range),
    "Select one rectangular range before using an assistant.",
  );
  const sheet = workbook.sheets.find((sheet) => sheet.id === selection.sheetId);
  assertCondition(
    sheet && !sheet.hidden,
    "Select a visible worksheet before using an assistant.",
  );
  const rangeParts = selection.range.toUpperCase().split(":");
  assertCondition(rangeParts.length <= 2, "Select one rectangular range.");
  const cellAddress = parseCellAddress(rangeParts[0]),
    endAddress = parseCellAddress(rangeParts[1] || rangeParts[0]);
  (assertCondition(
    cellAddress.row <= endAddress.row && cellAddress.col <= endAddress.col,
    "Selection range must run from top left to bottom right.",
  ),
    assertCondition(
      endAddress.row <= sheet.rowCount && endAddress.col <= sheet.colCount,
      "Selection exceeds worksheet dimensions.",
    ),
    assertCondition(
      (endAddress.row - cellAddress.row + 1) *
        (endAddress.col - cellAddress.col + 1) <=
        workbookLimits.maxLLMCells,
      `Select at most ${workbookLimits.maxLLMCells} cells for the assistant.`,
    ));
  const addresses = [];
  for (let row = cellAddress.row; row <= endAddress.row; row++)
    for (let col = cellAddress.col; col <= endAddress.col; col++)
      addresses.push(`${columnNumberToLetters(col)}${row}`);
  return addresses;
}
function buildAssistantSelectionContext(workbook, selection) {
  const assistantSelection = resolveAssistantSelection(workbook, selection),
    sheet = workbook.sheets.find((sheet) => sheet.id === selection.sheetId),
    workbookWithSettings = calculateWorkbookWithSettings(workbook);
  for (const address of assistantSelection) {
    const typedValue = workbookWithSettings.getTypedValue(sheet.id, address);
    (assertCondition(
      !(typedValue.type === "error" && typedValue.value === "#PENDING!"),
      "Calculation is pending in the approved selection. Choose Calculate Now before sharing current values with an assistant.",
    ),
      assertCondition(
        !(
          typedValue.type === "error" &&
          ["#TIMEOUT!", "#LIMIT!", "#BLOCKED!"].includes(
            String(typedValue.value),
          )
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
    sheetId: sheet.id,
    sheetName: sheet.name,
    range: selection.range.toUpperCase(),
    cells: assistantSelection.map((address) => ({
      address: address,
      raw: sheet.cells[address]?.raw ?? null,
      value: workbookWithSettings(sheet.id, address),
    })),
  };
}
function validateAssistantFormula(formulaText, selectedAddresses) {
  let formulaWithoutStrings = formulaText.replace(/"(?:[^"]|"")*"/g, '""');
  if (
    (assertCondition(
      !/[!\[\]|\\#]/.test(formulaWithoutStrings),
      "Assistant formulas cannot reference other worksheets, external workbooks, table names, existing spills or executable content.",
    ),
    assertCondition(
      !/(?:https?:|javascript:|data:|file:)/i.test(formulaText),
      "Assistant formulas cannot contain network or executable URLs.",
    ),
    /\b(?:LET|LAMBDA)\s*\(/i.test(formulaWithoutStrings))
  ) {
    const replacements = [];
    walkFormulaNodes(
      tokenizeFormulaText(formulaWithoutStrings),
      (node, boundNames, isCallee, isBindingDeclaration) => {
        (isBindingDeclaration ||
          (!node.qualifier &&
            boundNames.has(normalizeBindingName(node.name)))) &&
          replacements.push({
            start: node.start,
            end: node.end,
            value: isCallee ? "SUM" : "0",
          });
      },
    );
    for (const replacement of replacements.sort(
      (leftReplacement, rightReplacement) =>
        rightReplacement.start - leftReplacement.start,
    ))
      formulaWithoutStrings =
        formulaWithoutStrings.slice(0, replacement.start) +
        replacement.value +
        formulaWithoutStrings.slice(replacement.end);
  }
  const functionCallMatches = [
    ...formulaWithoutStrings.matchAll(/([A-Z_][A-Z0-9_.]*)\s*\(/gi),
  ];
  for (const functionCallMatch of functionCallMatches)
    assertCondition(
      nonVolatileFunctionNameSet.has(functionCallMatch[1].toUpperCase()),
      `Assistant formula function ${functionCallMatch[1]} is not allowed.`,
    );
  for (const cellReferenceMatch of formulaWithoutStrings.matchAll(
    new RegExp(
      "(?<![A-Z0-9_.])\\$?([A-Z]{1,3})\\$?([1-9]\\d*)(?![A-Z0-9_]|\\s*\\()",
      "gi",
    ),
  ))
    assertCondition(
      selectedAddresses.has(
        `${cellReferenceMatch[1].toUpperCase()}${cellReferenceMatch[2]}`,
      ),
      "Assistant formulas may only reference cells in the approved selection.",
    );
  const residualFormula = formulaWithoutStrings
    .replace(/([A-Z_][A-Z0-9_.]*)\s*\(/gi, "(")
    .replace(/\$?[A-Z]{1,3}\$?[1-9]\d*(?![A-Z0-9_])/gi, "0")
    .replace(/""/g, "0")
    .replace(/\b(?:TRUE|FALSE)\b/gi, "0")
    .replace(/\d+(?:\.\d+)?(?:E[+-]?\d+)?/gi, "0");
  assertCondition(
    !/[A-Z_]/i.test(residualFormula),
    "Named ranges and unsupported formula tokens are not allowed in assistant changes.",
  );
}
function validateProposedCell(cell, selectedAddresses) {
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
  for (const textField of ["note", "tickmark"])
    cell[textField] !== undefined &&
      assertCondition(
        typeof cell[textField] == "string" &&
          String(cell[textField]).length <= (textField === "note" ? 1e4 : 40),
        `Invalid cell ${textField}.`,
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
    for (const styleFlag of ["bold", "italic", "underline", "wrap"])
      cell.style[styleFlag] !== undefined &&
        assertCondition(
          typeof cell.style[styleFlag] == "boolean",
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
    for (const colorField of ["fill", "color"])
      cell.style[colorField] !== undefined &&
        assertCondition(
          typeof cell.style[colorField] == "string" &&
            /^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(
              String(cell.style[colorField]),
            ),
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
      validateAssistantFormula(cell.raw, selectedAddresses),
    structuredClone(cell)
  );
}
function validateAssistantProposal(proposal, selectionContext) {
  (assertCondition(
    isPlainRecord(proposal),
    "The assistant response must be a proposal object.",
  ),
    assertNoUnexpectedFields(proposal, [
      "id",
      "workbookId",
      "revision",
      "explanation",
      "citations",
      "patches",
      "grounding",
    ]),
    assertCondition(
      typeof proposal.id == "string" &&
        proposal.id.length > 0 &&
        proposal.id.length <= 100,
      "Invalid proposal identifier.",
    ),
    assertCondition(
      proposal.workbookId === selectionContext.workbookId &&
        proposal.revision === selectionContext.revision,
      "This proposal is stale or belongs to another workbook. Generate a new proposal.",
    ),
    assertCondition(
      typeof proposal.explanation == "string" &&
        proposal.explanation.length > 0 &&
        proposal.explanation.length <= 1e4,
      "Invalid proposal explanation.",
    ));
  const selectedAddresses = new Set(
    selectionContext.cells.map((cell) => cell.address),
  );
  (assertCondition(
    Array.isArray(proposal.citations) &&
      proposal.citations.length <= workbookLimits.maxLLMCells,
    "Invalid proposal citations.",
  ),
    assertCondition(
      proposal.citations.every(
        (citation) =>
          typeof citation == "string" && selectedAddresses.has(citation),
      ),
      "Citations must reference the approved selection.",
    ),
    assertCondition(
      Array.isArray(proposal.patches) &&
        proposal.patches.length <= workbookLimits.maxPatches,
      "Too many proposed edits.",
    ));
  const patchAddresses = new Set(),
    patches = proposal.patches.map(
      (patch) => (
        assertCondition(isPlainRecord(patch), "Each patch must be an object."),
        assertNoUnexpectedFields(patch, ["sheetId", "address", "cell"]),
        assertCondition(
          patch.sheetId === selectionContext.sheetId &&
            typeof patch.address == "string" &&
            selectedAddresses.has(patch.address),
          "Assistant changes must stay inside the approved selection.",
        ),
        assertCondition(
          !patchAddresses.has(patch.address),
          "The proposal contains duplicate cell edits.",
        ),
        patchAddresses.add(patch.address),
        {
          sheetId: selectionContext.sheetId,
          address: patch.address,
          cell: validateProposedCell(patch.cell, selectedAddresses),
        }
      ),
    );
  return {
    id: proposal.id,
    workbookId: selectionContext.workbookId,
    revision: selectionContext.revision,
    explanation: proposal.explanation,
    citations: [...proposal.citations],
    patches: patches,
    ...(proposal.grounding !== undefined
      ? {
          grounding: validateAnswerGrounding(
            proposal.grounding,
            selectionContext,
          ),
        }
      : {}),
  };
}
function validateProposedCalculation(workbook, assistantSelection, proposal) {
  const cells = resolveAssistantSelection(workbook, assistantSelection).map(
      (address) => ({
        address: address,
        raw: null,
        value: null,
      }),
    ),
    validatedProposal = validateAssistantProposal(proposal, {
      workbookId: workbook.id,
      revision: workbook.revision,
      sheetId: assistantSelection.sheetId,
      cells: cells,
    }),
    a = workbook.sheets.find(
      (sheet) => sheet.id === assistantSelection.sheetId,
    ),
    mergedProposal = {
      ...validatedProposal,
      patches: validatedProposal.patches.map((patch) => {
        const cell = a.cells[patch.address],
          mergedCell = {
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
          delete mergedCell.literal,
          patch.cell.format &&
            patch.cell.format !== cell?.format &&
            patch.cell.numFmt === undefined &&
            delete mergedCell.numFmt);
        const cellAddress = parseCellAddress(patch.address),
          isTableHeaderCell = (a.tables ?? []).some((table) => {
            const [tableStartRef, tableEndRef] = table.range.split(":"),
              tableStart = parseCellAddress(tableStartRef),
              tableEnd = parseCellAddress(tableEndRef ?? tableStartRef);
            return (
              cellAddress.row === tableStart.row &&
              cellAddress.col >= tableStart.col &&
              cellAddress.col <= tableEnd.col
            );
          });
        return (
          assertCondition(
            !isTableHeaderCell ||
              (cell?.raw === mergedCell.raw &&
                !!cell?.literal == !!mergedCell.literal),
            "Assistant changes cannot rename table headers because dependent references may change outside the approved selection.",
          ),
          { ...patch, cell: mergedCell }
        );
      }),
    },
    selectedAddresses = new Set(cells.map((cell) => cell.address)),
    workbookWithPatches = {
      ...workbook,
      sheets: workbook.sheets.map((sheet) =>
        sheet.id === a.id
          ? {
              ...sheet,
              cells: {
                ...sheet.cells,
                ...Object.fromEntries(
                  mergedProposal.patches.map((patch) => [
                    patch.address,
                    patch.cell,
                  ]),
                ),
              },
            }
          : sheet,
      ),
    },
    calculator = evaluateWorkbookFormulas(
      workbookWithPatches,
      {},
      {
        allowRead: () => true,
        allowDatabaseCriterion: (sheetId, criterionFormula) => {
          if (sheetId !== a.id) return false;
          try {
            return (
              validateAssistantFormula(criterionFormula, selectedAddresses),
              true
            );
          } catch {
            return false;
          }
        },
        step: () => {},
      },
    );
  for (const patch of mergedProposal.patches) {
    if (
      typeof patch.cell.raw != "string" ||
      !patch.cell.raw.startsWith("=") ||
      patch.cell.literal
    )
      continue;
    const calculatedValue = calculator(a.id, patch.address);
    (assertCondition(
      calculatedValue !== "#LIMIT!",
      "The proposed formula exceeds calculation limits.",
    ),
      assertCondition(
        calculator.getOperationalErrors().length === 0,
        "The proposed formula has invalid syntax, a circular reference or an operational calculation failure. Correct the formula before review; error handlers cannot suppress this rejection.",
      ));
    const spillRange = calculator.getSpillRange(a.id, patch.address);
    if (
      (assertCondition(
        !calculator.getSpillDiagnostic(a.id, patch.address),
        "The proposed array is blocked. Review the spill obstruction before requesting approval.",
      ),
      spillRange)
    ) {
      const [spillStartRef, spillEndRef] = spillRange.split(":"),
        cellAddress = parseCellAddress(spillStartRef),
        spillEnd = parseCellAddress(spillEndRef ?? spillStartRef);
      assertCondition(
        (spillEnd.row - cellAddress.row + 1) *
          (spillEnd.col - cellAddress.col + 1) <=
          workbookLimits.maxLLMCells,
        "Assistant spill outputs must stay inside the approved selection.",
      );
      for (let row = cellAddress.row; row <= spillEnd.row; row++)
        for (let col = cellAddress.col; col <= spillEnd.col; col++)
          assertCondition(
            selectedAddresses.has(`${columnNumberToLetters(col)}${row}`),
            "Assistant spill outputs must stay inside the approved selection.",
          );
    }
  }
  return (
    assertCondition(
      calculator.getOperationalErrors().length === 0,
      "The proposed formula cannot be validated because calculation encountered an operational failure.",
    ),
    mergedProposal
  );
}
function buildLocalDemoProposal() {
  return {
    name: "Local demo · no model or network",
    localOnly: true,
    chat: { history: true, attachments: true },
    async propose(prompt, selectionContext, abortSignal, chatRequest) {
      if (abortSignal?.aborted)
        throw new DOMException("Request cancelled.", "AbortError");
      const proposal = {
          id: `local-${selectionContext.revision}-${Date.now()}`,
          workbookId: selectionContext.workbookId,
          revision: selectionContext.revision,
          explanation: "",
          citations: [],
          patches: [],
        },
        numericCells = selectionContext.cells.filter(
          (cell) =>
            typeof cell.value == "number" && Number.isFinite(cell.value),
        ),
        numericSum = numericCells.reduce(
          (total, cell) => total + Number(cell.value),
          0,
        );
      if (/highlight|high values|outlier/i.test(prompt)) {
        const thresholdMatch =
            /(?:above|over|greater than|exceed(?:ing|s)?)\s*[$€£]?\s*(-?\d[\d,]*(?:\.\d+)?)/i.exec(
              prompt,
            ),
          threshold = thresholdMatch
            ? Number(thresholdMatch[1].replaceAll(",", ""))
            : numericCells.length
              ? Math.abs(numericSum / numericCells.length) * 1.5
              : 0,
          flaggedCells = numericCells.filter(
            (cell) =>
              (thresholdMatch
                ? Number(cell.value)
                : Math.abs(Number(cell.value))) > threshold,
          );
        ((proposal.explanation = `Local deterministic demo: ${flaggedCells.length} values exceed ${thresholdMatch ? `the requested threshold of ${threshold.toLocaleString()}` : `1.5 times the absolute arithmetic mean (${threshold.toLocaleString()})`}. This is a simple illustration, not a statistical audit conclusion. The listed cells receive an amber fill; existing annotations are retained.`),
          (proposal.citations = flaggedCells.map((cell) => cell.address)),
          (proposal.patches = flaggedCells.map((cell) => ({
            sheetId: selectionContext.sheetId,
            address: cell.address,
            cell: { raw: cell.raw, style: { fill: "#fff0b3" } },
          }))));
      } else if (/\b(?:total|sum)\b/i.test(prompt)) {
        const emptyCell = [...selectionContext.cells]
          .reverse()
          .find((cell) => cell.raw === null || cell.raw === "");
        if (emptyCell && numericCells.length) {
          const raw = `=SUM(${numericCells.map((cell) => cell.address).join(",")})`;
          ((proposal.patches = [
            {
              sheetId: selectionContext.sheetId,
              address: emptyCell.address,
              cell: { raw: raw, format: "number", style: { bold: true } },
            },
          ]),
            (proposal.citations = numericCells.map((cell) => cell.address)),
            (proposal.explanation = `Local deterministic demo: put a SUM formula for ${numericCells.length} numeric selected cells in empty cell ${emptyCell.address}. Current result: ${numericSum.toLocaleString()}. Review the included addresses for subtotals or duplicates before accepting.`));
        } else
          proposal.explanation =
            "Local deterministic demo: select numeric cells and at least one empty destination cell to propose a total.";
      } else if (/explain/i.test(prompt)) {
        const formulaCells = selectionContext.cells.filter(
          (cell) => typeof cell.raw == "string" && cell.raw.startsWith("="),
        );
        ((proposal.explanation = formulaCells.length
          ? `Local deterministic demo: ${formulaCells.map((cell) => `${cell.address} contains ${cell.raw} and currently returns ${String(cell.value)}`).join("; ")}. This lists formulas and results; no language model interpretation is used.`
          : "Local deterministic demo: the selection contains no formula expressions. Values and formulas are worksheet data, never instructions."),
          (proposal.citations = formulaCells.map((cell) => cell.address)));
      } else
        ((proposal.explanation = `Local deterministic demo: ${selectionContext.cells.length} selected cells; ${numericCells.length} numeric values; sum ${numericSum.toLocaleString()}; ${selectionContext.cells.filter((cell) => cell.raw === null || cell.raw === "").length} empty cells. Workbook content is treated only as data.`),
          (proposal.citations = numericCells
            .slice(0, 20)
            .map((cell) => cell.address)));
      if (
        ((proposal.grounding = {
          claims: [
            {
              id: "local-evidence",
              kind: "workbook",
              text: proposal.explanation,
              ranges: [selectionContext.range],
            },
          ],
          externalSources: [],
        }),
        chatRequest?.attachments.length)
      ) {
        const attachments = parseAssistantChatRequest(chatRequest).attachments,
          attachmentSummary = `Received ${attachments.length} attachment${attachments.length === 1 ? "" : "s"}: ${attachments.map((attachment) => `${attachment.name} (${attachment.size.toLocaleString()} bytes)`).join(", ")}. Files remain in this browser in local demo mode. Preview text and images in the attachment cards; document interpretation requires a connected model.`;
        ((proposal.explanation += `

${attachmentSummary}`),
          proposal.grounding.claims.push({
            id: "local-attachments",
            kind: "commentary",
            text: attachmentSummary,
          }));
      }
      return validateAssistantProposal(proposal, selectionContext);
    },
  };
}

/* From 11-local-workbook-store-and-versions.js */
const isPlainRecord = (value) =>
  !!value &&
  typeof value == "object" &&
  !Array.isArray(value) &&
  Object.getPrototypeOf(value) === Object.prototype;

/* From 12-consolidation-and-dependency-inspection.js */
const hasOnlyAllowedKeys = (value, allowedKeys) =>
  !!value &&
  typeof value == "object" &&
  !Array.isArray(value) &&
  Object.keys(value).every((key) => allowedKeys.includes(key));

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
