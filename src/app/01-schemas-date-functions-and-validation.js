import {
  iso,
  string,
  union,
  number,
  boolean,
  null as _null,
  array,
  strictObject,
  enum as _enum,
  discriminatedUnion,
  literal,
  record,
  unknown,
} from "zod";

var objectDefineProperty = Object.defineProperty;
var defineDataProperty = (target, key, value) =>
  key in target
    ? objectDefineProperty(target, key, {
        enumerable: true,
        configurable: true,
        writable: true,
        value: value,
      })
    : (target[key] = value);
var defineDataPropertyByKey = (target, key, value) =>
  defineDataProperty(target, typeof key != "symbol" ? key + "" : key, value);
function copyModuleProperties(target, sources) {
  for (var index = 0; index < sources.length; index++) {
    const source = sources[index];
    if (typeof source != "string" && !Array.isArray(source)) {
      for (const key in source)
        if (key !== "default" && !(key in target)) {
          const descriptor = Object.getOwnPropertyDescriptor(source, key);
          descriptor &&
            Object.defineProperty(
              target,
              key,
              descriptor.get
                ? descriptor
                : { enumerable: true, get: () => source[key] },
            );
        }
    }
  }
  return Object.freeze(
    Object.defineProperty(target, Symbol.toStringTag, { value: "Module" }),
  );
}
function unwrapDefaultExport(moduleExports) {
  return moduleExports &&
    moduleExports.__esModule &&
    Object.prototype.hasOwnProperty.call(moduleExports, "default")
    ? moduleExports.default
    : moduleExports;
}

const deepFrozenValueSet = new WeakSet();
function isDeepFrozenValue(value) {
  return !!value && typeof value == "object" && deepFrozenValueSet.has(value);
}
function deepFreezeValue(value) {
  if (!value || typeof value != "object" || deepFrozenValueSet.has(value))
    return value;
  for (const childValue of Object.values(value)) deepFreezeValue(childValue);
  return (Object.freeze(value), deepFrozenValueSet.add(value), value);
}
const linkedOrBoundCellCache = new WeakMap();
function hasLinkedOrBoundCells(cells) {
  if (!isDeepFrozenValue(cells)) return true;
  let cachedResult = linkedOrBoundCellCache.get(cells);
  if (cachedResult === undefined) {
    cachedResult = false;
    for (const cellKey in cells)
      if (cells[cellKey].hyperlink || cells[cellKey].style?.binding) {
        cachedResult = true;
        break;
      }
    linkedOrBoundCellCache.set(cells, cachedResult);
  }
  return cachedResult;
}
const serializedByteCache = new WeakMap(),
  utf8Encoder = new TextEncoder();
function measureJsonStringBytes(text) {
  return /^[\x20-\x21\x23-\x5b\x5d-\x7e]*$/.test(text)
    ? text.length + 2
    : utf8Encoder.encode(JSON.stringify(text)).byteLength;
}
function measureSerializedBytes(value) {
  if (typeof value == "string") return measureJsonStringBytes(value);
  if (!value || typeof value != "object")
    return JSON.stringify(value)?.length ?? 0;
  const cachedBytes = isDeepFrozenValue(value)
    ? serializedByteCache.get(value)
    : undefined;
  if (cachedBytes !== undefined) return cachedBytes;
  let totalBytes = 2,
    entryCount = 0;
  if (Array.isArray(value))
    for (const element of value)
      totalBytes +=
        (entryCount++ ? 1 : 0) +
        measureSerializedBytes(element === undefined ? null : element);
  else
    for (const key in value) {
      if (!Object.prototype.hasOwnProperty.call(value, key)) continue;
      const propertyValue = value[key];
      propertyValue !== undefined &&
        (totalBytes +=
          (entryCount++ ? 1 : 0) +
          measureJsonStringBytes(key) +
          1 +
          measureSerializedBytes(propertyValue));
    }
  return (
    isDeepFrozenValue(value) && serializedByteCache.set(value, totalBytes),
    totalBytes
  );
}
const isoDateTimeSchemaFactory = iso.datetime;
function sha256HexDigest(bytes) {
  if (
    !ArrayBuffer.isView(bytes) ||
    Object.prototype.toString.call(bytes) !== "[object Uint8Array]" ||
    bytes.length > 50 * 1024 * 1024
  )
    throw Error("Hashing requires at most 50 MB of bytes.");
  const roundConstants = [
      1116352408, 1899447441, 3049323471, 3921009573, 961987163, 1508970993,
      2453635748, 2870763221, 3624381080, 310598401, 607225278, 1426881987,
      1925078388, 2162078206, 2614888103, 3248222580, 3835390401, 4022224774,
      264347078, 604807628, 770255983, 1249150122, 1555081692, 1996064986,
      2554220882, 2821834349, 2952996808, 3210313671, 3336571891, 3584528711,
      113926993, 338241895, 666307205, 773529912, 1294757372, 1396182291,
      1695183700, 1986661051, 2177026350, 2456956037, 2730485921, 2820302411,
      3259730800, 3345764771, 3516065817, 3600352804, 4094571909, 275423344,
      430227734, 506948616, 659060556, 883997877, 958139571, 1322822218,
      1537002063, 1747873779, 1955562222, 2024104815, 2227730452, 2361852424,
      2428436474, 2756734187, 3204031479, 3329325298,
    ],
    hashState = [
      1779033703, 3144134277, 1013904242, 2773480762, 1359893119, 2600822924,
      528734635, 1541459225,
    ],
    paddedBytes = new Uint8Array(Math.ceil((bytes.length + 9) / 64) * 64);
  (paddedBytes.set(bytes), (paddedBytes[bytes.length] = 128));
  const paddedView = new DataView(paddedBytes.buffer);
  paddedView.setUint32(paddedBytes.length - 4, bytes.length * 8);
  const messageSchedule = new Uint32Array(64),
    rotateRight = (word, shift) => (word >>> shift) | (word << (32 - shift));
  for (
    let blockOffset = 0;
    blockOffset < paddedBytes.length;
    blockOffset += 64
  ) {
    for (let wordIndex = 0; wordIndex < 16; wordIndex++)
      messageSchedule[wordIndex] = paddedView.getUint32(
        blockOffset + wordIndex * 4,
      );
    for (let scheduleIndex = 16; scheduleIndex < 64; scheduleIndex++) {
      const earlierWord = messageSchedule[scheduleIndex - 15],
        recentWord = messageSchedule[scheduleIndex - 2];
      messageSchedule[scheduleIndex] =
        (messageSchedule[scheduleIndex - 16] +
          (rotateRight(earlierWord, 7) ^
            rotateRight(earlierWord, 18) ^
            (earlierWord >>> 3)) +
          messageSchedule[scheduleIndex - 7] +
          (rotateRight(recentWord, 17) ^
            rotateRight(recentWord, 19) ^
            (recentWord >>> 10))) >>>
        0;
    }
    let [varA, varB, varC, varD, varE, varF, varG, varH] = hashState;
    for (let roundIndex = 0; roundIndex < 64; roundIndex++) {
      const temp1 =
          (varH +
            (rotateRight(varE, 6) ^
              rotateRight(varE, 11) ^
              rotateRight(varE, 25)) +
            ((varE & varF) ^ (~varE & varG)) +
            roundConstants[roundIndex] +
            messageSchedule[roundIndex]) >>>
          0,
        temp2 =
          ((rotateRight(varA, 2) ^
            rotateRight(varA, 13) ^
            rotateRight(varA, 22)) +
            ((varA & varB) ^ (varA & varC) ^ (varB & varC))) >>>
          0;
      ((varH = varG),
        (varG = varF),
        (varF = varE),
        (varE = (varD + temp1) >>> 0),
        (varD = varC),
        (varC = varB),
        (varB = varA),
        (varA = (temp1 + temp2) >>> 0));
    }
    [varA, varB, varC, varD, varE, varF, varG, varH].forEach(
      (workingValue, index) =>
        (hashState[index] = (hashState[index] + workingValue) >>> 0),
    );
  }
  return hashState
    .map((hashWord) => hashWord.toString(16).padStart(8, "0"))
    .join("");
}
const transformationRecipeLimits = Object.freeze({
    sources: 8,
    steps: 20,
    rows: 1e3,
    columns: 50,
    cells: 2e4,
    characters: 1e6,
    lineageLinks: 2e4,
    recipes: 20,
    versions: 20,
    metadataCharacters: 2e6,
    runMetadataCharacters: 8e6,
  }),
  recipeStoreFormatId = "ledgerglass-recipes-v1",
  singleLineTextSchema = string()
    .trim()
    .min(1)
    .max(128)
    .refine((text) => !/[\u0000-\u001f\u007f]/.test(text)),
  opaqueIdSchema = string().min(1).max(100),
  cellScalarSchema = union([
    string().max(32767),
    number().finite(),
    boolean(),
    _null(),
  ]),
  distinctFieldListSchema = array(singleLineTextSchema)
    .max(50)
    .refine(
      (fieldNames) => new Set(fieldNames).size === fieldNames.length,
      "Fields must be distinct.",
    ),
  sheetRangeSourceSchema = strictObject({
    sheetId: opaqueIdSchema,
    range: string().min(1).max(50),
    tableId: opaqueIdSchema.optional(),
  }),
  aggregateMeasureSchema = strictObject({
    field: singleLineTextSchema,
    operation: _enum([
      "sum",
      "count",
      "average",
      "min",
      "max",
      "distinctCount",
    ]),
    name: singleLineTextSchema,
  }),
  filterConditionSchema = strictObject({
    field: singleLineTextSchema,
    operator: _enum([
      "equals",
      "notEquals",
      "contains",
      "startsWith",
      "endsWith",
      "gt",
      "gte",
      "lt",
      "lte",
      "blank",
      "notBlank",
    ]),
    value: cellScalarSchema.optional(),
  }),
  od = { id: opaqueIdSchema },
  transformationStepSchema = discriminatedUnion("kind", [
    strictObject({
      ...od,
      kind: literal("select"),
      fields: distinctFieldListSchema.min(1),
    }),
    strictObject({
      ...od,
      kind: literal("rename"),
      mappings: array(
        strictObject({ from: singleLineTextSchema, to: singleLineTextSchema }),
      )
        .min(1)
        .max(50),
    }),
    strictObject({
      ...od,
      kind: literal("filter"),
      conditions: array(filterConditionSchema).min(1).max(20),
      join: _enum(["and", "or"]),
    }),
    strictObject({
      ...od,
      kind: literal("types"),
      fields: distinctFieldListSchema.min(1),
      to: _enum(["text", "number", "integer", "boolean", "date"]),
      decimal: _enum([".", ","]),
      group: _enum(["", ".", ",", " ", "'"]),
      dateOrder: _enum(["MDY", "DMY"]),
    }),
    strictObject({
      ...od,
      kind: literal("replace"),
      fields: distinctFieldListSchema.min(1),
      mode: _enum(["value", "text"]),
      find: cellScalarSchema,
      replacement: cellScalarSchema,
    }),
    strictObject({
      ...od,
      kind: literal("derive"),
      name: singleLineTextSchema,
      expression: string().min(1).max(1e3),
    }),
    strictObject({
      ...od,
      kind: literal("sort"),
      fields: array(
        strictObject({
          field: singleLineTextSchema,
          direction: _enum(["ascending", "descending"]),
        }),
      )
        .min(1)
        .max(8),
    }),
    strictObject({
      ...od,
      kind: literal("unpivot"),
      fields: distinctFieldListSchema.min(1),
      attribute: singleLineTextSchema,
      value: singleLineTextSchema,
      keepNulls: boolean(),
    }),
    strictObject({
      ...od,
      kind: literal("group"),
      keys: distinctFieldListSchema,
      values: array(aggregateMeasureSchema).min(1).max(20),
    }),
    strictObject({
      ...od,
      kind: literal("pivot"),
      keys: distinctFieldListSchema,
      column: singleLineTextSchema,
      value: singleLineTextSchema,
      operation: aggregateMeasureSchema.shape.operation,
    }),
    strictObject({
      ...od,
      kind: literal("append"),
      source: sheetRangeSourceSchema,
    }),
    strictObject({
      ...od,
      kind: literal("merge"),
      source: sheetRangeSourceSchema,
      mode: _enum(["inner", "left", "right", "full", "leftAnti", "rightAnti"]),
      keys: array(
        strictObject({
          left: singleLineTextSchema,
          right: singleLineTextSchema,
        }),
      )
        .min(1)
        .max(8),
      prefix: singleLineTextSchema,
      normalization: strictObject({
        trim: boolean().optional(),
        collapseWhitespace: boolean().optional(),
        ignoreCase: boolean().optional(),
        unicode: _enum(["NFC", "NFKC"]).optional(),
      }).optional(),
      fuzzy: strictObject({
        key: number().int().min(0).max(7),
        threshold: number().finite().min(0).max(1),
      }).optional(),
    }),
  ]),
  transformationRecipeSchema = strictObject({
    engine: literal(recipeStoreFormatId),
    source: sheetRangeSourceSchema,
    steps: array(transformationStepSchema).max(
      transformationRecipeLimits.steps,
    ),
  }).superRefine((recipe, context) => {
    (new Set(recipe.steps.map((step) => step.id)).size !==
      recipe.steps.length &&
      context.addIssue({ code: "custom", message: "Step IDs must be unique." }),
      new Set(
        [
          recipe.source,
          ...recipe.steps.flatMap((step) =>
            step.kind === "append" || step.kind === "merge"
              ? [step.source]
              : [],
          ),
        ].map((source) => JSON.stringify(source)),
      ).size > transformationRecipeLimits.sources &&
        context.addIssue({
          code: "custom",
          message: "Use no more than eight recipe sources.",
        }));
  });
function canonicalJsonStringify(value) {
  return JSON.stringify(value, (key, nodeValue) =>
    nodeValue && typeof nodeValue == "object" && !Array.isArray(nodeValue)
      ? Object.fromEntries(
          Object.keys(nodeValue)
            .sort()
            .map((propertyName) => [propertyName, nodeValue[propertyName]]),
        )
      : nodeValue,
  );
}
function fingerprintCanonicalValue(value) {
  const canonicalJson = canonicalJsonStringify(value);
  if (canonicalJson.length > 16e6)
    throw Error(
      "Recipe fingerprint exceeds the 16-million-character evidence budget.",
    );
  return sha256HexDigest(
    new Uint8Array(new TextEncoder().encode(canonicalJson)),
  );
}
function parseTransformationRecipe(definition) {
  const parseResult = transformationRecipeSchema.safeParse(definition);
  if (!parseResult.success)
    throw Error(`Invalid recipe: ${parseResult.error.issues[0]?.message}`);
  return parseResult.data;
}
const sha256HexSchema = string().regex(/^[0-9a-f]{64}$/),
  nonnegativeSafeIntegerSchema = number()
    .int()
    .nonnegative()
    .max(Number.MAX_SAFE_INTEGER - 1),
  isoTimestampSchema = string().refine(
    (timestamp) =>
      Number.isFinite(Date.parse(timestamp)) &&
      new Date(timestamp).toISOString() === timestamp,
  ),
  transformationRecipeVersionSchema = strictObject({
    version: number().int().min(1).max(transformationRecipeLimits.versions),
    definition: transformationRecipeSchema,
    definitionHash: sha256HexSchema,
    createdAt: isoTimestampSchema,
  }),
  SJ = array(
    strictObject({
      id: opaqueIdSchema,
      name: singleLineTextSchema,
      versions: array(transformationRecipeVersionSchema)
        .min(1)
        .max(transformationRecipeLimits.versions),
    }),
  ).max(transformationRecipeLimits.recipes),
  columnTypeProfileSchema = strictObject({
    name: singleLineTextSchema,
    types: array(_enum(["blank", "number", "string", "boolean"])).max(4),
    declaredType: string().max(32).optional(),
  }),
  numericColumnSummarySchema = strictObject({
    field: singleLineTextSchema,
    numericCount: nonnegativeSafeIntegerSchema,
    blankCount: nonnegativeSafeIntegerSchema,
    sum: number().finite(),
  }),
  transformationStepRunSchema = strictObject({
    index: number().int().min(1).max(20),
    id: opaqueIdSchema,
    kind: string().max(32),
    status: _enum(["complete", "failed", "skipped"]),
    inputSchema: array(columnTypeProfileSchema).max(50),
    inputRows: nonnegativeSafeIntegerSchema,
    outputSchema: array(columnTypeProfileSchema).max(50).optional(),
    outputRows: nonnegativeSafeIntegerSchema.optional(),
    inputHash: sha256HexSchema.optional(),
    outputHash: sha256HexSchema.optional(),
    controls: array(numericColumnSummarySchema).max(50).optional(),
    error: string().max(4e3).optional(),
    details: record(string(), unknown()).optional(),
  }),
  transformationRunEvidenceSchema = strictObject({
    engine: literal(recipeStoreFormatId),
    recipeId: opaqueIdSchema,
    recipeName: singleLineTextSchema,
    version: number().int().min(1).max(20),
    definition: transformationRecipeSchema,
    definitionHash: sha256HexSchema,
    sourceWorkbookId: opaqueIdSchema,
    sourceRevision: nonnegativeSafeIntegerSchema,
    sourceHash: sha256HexSchema,
    outputHash: sha256HexSchema,
    outputRows: number().int().min(0).max(1e3),
    outputRange: string().min(1).max(50),
    at: isoTimestampSchema,
    steps: array(transformationStepRunSchema).max(20),
    controls: array(numericColumnSummarySchema).max(50),
    lineage: record(
      string().regex(/^[1-9]\d*$/),
      array(string().min(1).max(200)).max(4e3),
    ),
  });
function validateSavedRecipeMetadata(metadataInput, workbook) {
  const metadata = metadataInput;
  if (metadata.transformationRecipes !== undefined) {
    if (
      JSON.stringify(metadata.transformationRecipes).length >
      transformationRecipeLimits.metadataCharacters
    )
      throw Error("Saved recipe metadata exceeds 2,000,000 characters.");
    const recipesResult = SJ.safeParse(metadata.transformationRecipes);
    if (!recipesResult.success)
      throw Error(
        `Invalid saved recipes: ${recipesResult.error.issues[0]?.message}`,
      );
    if (
      new Set(recipesResult.data.map((savedRecipe) => savedRecipe.id)).size !==
        recipesResult.data.length ||
      new Set(
        recipesResult.data.map((namedRecipe) => namedRecipe.name.toLowerCase()),
      ).size !== recipesResult.data.length
    )
      throw Error("Saved recipe names and IDs must be unique.");
    for (const recipe of recipesResult.data)
      for (const [versionIndex, versionEntry] of recipe.versions.entries())
        if (
          versionEntry.version !== versionIndex + 1 ||
          versionEntry.definitionHash !==
            fingerprintCanonicalValue(versionEntry.definition)
        )
          throw Error(
            "Saved recipe versions or definition hashes are inconsistent.",
          );
    workbook.transformationRecipes = recipesResult.data;
  }
  let runMetadataCharacters = 0;
  for (const sheet of workbook.sheets) {
    const sheetMetadata = metadata.sheets?.find(
      (metadataSheet) => metadataSheet.id === sheet.id,
    );
    if (sheetMetadata?.transformationRun === undefined) continue;
    if (
      ((runMetadataCharacters += JSON.stringify(
        sheetMetadata.transformationRun,
      ).length),
      runMetadataCharacters > transformationRecipeLimits.runMetadataCharacters)
    )
      throw Error(
        "Recipe execution evidence exceeds 8,000,000 characters. Retain a native backup and remove obsolete output sheets.",
      );
    const runResult = transformationRunEvidenceSchema.safeParse(
      sheetMetadata.transformationRun,
    );
    if (!runResult.success)
      throw Error(
        `Invalid recipe execution evidence: ${runResult.error.issues[0]?.message}`,
      );
    const data = runResult.data;
    if (
      data.definitionHash !== fingerprintCanonicalValue(data.definition) ||
      data.steps.length !== data.definition.steps.length ||
      data.steps.some(
        (step, stepIndex) =>
          step.status !== "complete" ||
          step.index !== stepIndex + 1 ||
          step.id !== data.definition.steps[stepIndex].id ||
          step.kind !== data.definition.steps[stepIndex].kind,
      ) ||
      Object.values(data.lineage).reduce(
        (total, links) => total + links.length,
        0,
      ) > transformationRecipeLimits.lineageLinks ||
      Object.keys(data.lineage).some(
        (rowKey) => Number(rowKey) < 2 || Number(rowKey) > data.outputRows + 1,
      ) ||
      Object.keys(data.lineage).length !== data.outputRows
    )
      throw Error("Recipe execution metadata is inconsistent.");
    sheet.transformationRun = data;
  }
}
const normalizeFunctionName = (rawFunctionName) =>
    rawFunctionName.replace(/^_xlfn\.(?:_xlws\.)?/i, "").toUpperCase(),
  normalizeBindingName = (rawBindingName) =>
    rawBindingName.replace(/^_xlpm\./i, "").toLowerCase();
function isValidFormulaName(formulaName, isLambdaParameter = false) {
  return (
    (formulaName = formulaName.replace(/^_xlpm\./i, "")),
    /^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/.test(formulaName) &&
      !/^(?:[A-Za-z]{1,3}[1-9]\d*|R\d+C\d+|R|C)$/i.test(formulaName) &&
      !/^__LG/i.test(formulaName) &&
      (!isLambdaParameter || !formulaName.includes("."))
  );
}
function tokenizeFormulaText(formulaText) {
  let end = formulaText.startsWith("=") ? 1 : 0;
  const throwSyntaxError = () => {
      throw Error("Unbalanced or malformed formula syntax.");
    },
    skipQuoted = (quoteChar) => {
      for (end++; end < formulaText.length;)
        if (formulaText[end++] === quoteChar)
          if (formulaText[end] === quoteChar) end++;
          else return;
      throwSyntaxError();
    },
    skipBalanced = (openChar, closeChar) => {
      let depth = 0;
      do {
        const currentChar = formulaText[end];
        if (currentChar === '"' || (currentChar === "'" && openChar !== "[")) {
          skipQuoted(currentChar);
          continue;
        }
        (currentChar === openChar && depth++,
          currentChar === closeChar && depth--,
          end++);
      } while (end < formulaText.length && depth);
      depth && throwSyntaxError();
    },
    parseTokens = (nestingLevel) => {
      if (nestingLevel > 128) throw Error("#LIMIT!");
      const tokenList = [];
      for (
        ;
        end < formulaText.length && ![",", ")"].includes(formulaText[end]);
      ) {
        const start = end;
        let callee;
        const firstChar = formulaText[end];
        if (firstChar === '"') skipQuoted(firstChar);
        else if (firstChar === "{") skipBalanced("{", "}");
        else if (firstChar === "[") skipBalanced("[", "]");
        else if (firstChar === "(") {
          end++;
          const body = parseTokens(nestingLevel + 1);
          (formulaText[end++] !== ")" && throwSyntaxError(),
            (callee = { kind: "group", body: body, start: start, end: end }));
        } else {
          let qualifier;
          const qualifierMatch =
            /^(?:'((?:[^']|'')+)'|([A-Za-z0-9_\\][A-Za-z0-9_.\\]*))!/.exec(
              formulaText.slice(end),
            );
          qualifierMatch &&
            ((qualifier =
              qualifierMatch[1]?.replace(/''/g, "'") ?? qualifierMatch[2]),
            (end += qualifierMatch[0].length));
          const literalMatch =
            /^(?:#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|CYCLE!|LIMIT!|BLOCKED!|ERROR!)|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w.])|\$?[1-9]\d*:\$?[1-9]\d*(?![\w.])|\$?[A-Za-z]{1,3}\$?[1-9]\d*(?![\w.\\])|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)/i.exec(
              formulaText.slice(end),
            );
          if (literalMatch) end += literalMatch[0].length;
          else {
            const nameMatch = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(
              formulaText.slice(end),
            );
            nameMatch
              ? ((end += nameMatch[0].length),
                formulaText[end] === "["
                  ? skipBalanced("[", "]")
                  : (callee = {
                      kind: "symbol",
                      name: nameMatch[0],
                      qualifier: qualifier,
                      start: start,
                      end: end,
                    }))
              : ((qualifier || firstChar === "'" || firstChar === "}") &&
                  throwSyntaxError(),
                end++);
          }
        }
        if (
          ((callee ??= {
            kind: "text",
            text: formulaText.slice(start, end),
            start: start,
            end: end,
          }),
          callee.kind !== "text")
        )
          for (;;) {
            const openParenMatch = /^\s*\(/.exec(formulaText.slice(end));
            if (!openParenMatch) break;
            end += openParenMatch[0].length;
            const args = [];
            if (formulaText[end] !== ")")
              for (
                ;
                args.push(parseTokens(nestingLevel + 1)),
                  formulaText[end] === ",";
              )
                end++;
            (formulaText[end++] !== ")" && throwSyntaxError(),
              (callee = {
                kind: "call",
                callee: callee,
                args: args,
                start: start,
                end: end,
              }));
          }
        tokenList.push(callee);
      }
      return tokenList;
    },
    topLevelTokens = parseTokens(0);
  return (end !== formulaText.length && throwSyntaxError(), topLevelTokens);
}
const dropBlankTextTokens = (tokens) =>
  tokens.filter((token) => token.kind !== "text" || token.text.trim());
function validateBindingName(tokens, isLambdaParameter = false) {
  const nonBlankTokens = dropBlankTextTokens(tokens),
    nameToken = nonBlankTokens[0];
  if (
    nonBlankTokens.length !== 1 ||
    nameToken?.kind !== "symbol" ||
    nameToken.qualifier ||
    !isValidFormulaName(nameToken.name, isLambdaParameter)
  )
    throw Error(
      `Invalid ${isLambdaParameter ? "LAMBDA parameter" : "LET variable"}: use a name, not a cell address${isLambdaParameter ? ", without a period" : ""}.`,
    );
  return nameToken;
}
function validateLetLambdaCall(callNode) {
  if (callNode.callee.kind !== "symbol" || callNode.callee.qualifier) return;
  const functionName = normalizeFunctionName(callNode.callee.name);
  if (functionName !== "LET" && functionName !== "LAMBDA") return;
  if (
    functionName === "LET" &&
    (callNode.args.length < 3 ||
      callNode.args.length > 253 ||
      callNode.args.length % 2 !== 1)
  )
    throw Error("LET requires 1–126 name/value pairs and a final calculation.");
  if (
    functionName === "LAMBDA" &&
    (callNode.args.length < 1 || callNode.args.length > 254)
  )
    throw Error("LAMBDA requires a calculation and at most 253 parameters.");
  if (callNode.args.some((arg) => !dropBlankTextTokens(arg).length))
    throw Error(`${functionName} arguments cannot be empty.`);
  const bindingNames = new Set();
  for (
    let argIndex = 0;
    argIndex < callNode.args.length - 1;
    argIndex += functionName === "LET" ? 2 : 1
  ) {
    const bindingName = normalizeBindingName(
      validateBindingName(callNode.args[argIndex], functionName === "LAMBDA")
        .name,
    );
    if (bindingNames.has(bindingName))
      throw Error(`Duplicate ${functionName} binding: ${bindingName}.`);
    bindingNames.add(bindingName);
  }
  return functionName;
}
function walkFormulaNodes(nodes, visitor, boundNames = new Set()) {
  for (const node of nodes) {
    if (
      (node.kind === "symbol" && visitor(node, boundNames, false, false),
      node.kind === "group" && walkFormulaNodes(node.body, visitor, boundNames),
      node.kind !== "call")
    )
      continue;
    const letLambdaName = validateLetLambdaCall(node);
    if (
      (node.callee.kind === "symbol"
        ? visitor(node.callee, boundNames, true, false)
        : walkFormulaNodes([node.callee], visitor, boundNames),
      !letLambdaName)
    ) {
      node.args.forEach((arg) => walkFormulaNodes(arg, visitor, boundNames));
      continue;
    }
    const innerNames = new Set(boundNames);
    for (
      let bindingIndex = 0;
      bindingIndex < node.args.length - 1;
      bindingIndex += letLambdaName === "LET" ? 2 : 1
    ) {
      const bindingNode = validateBindingName(
        node.args[bindingIndex],
        letLambdaName === "LAMBDA",
      );
      (visitor(bindingNode, innerNames, false, true),
        letLambdaName === "LET" &&
          walkFormulaNodes(node.args[bindingIndex + 1], visitor, innerNames),
        innerNames.add(normalizeBindingName(bindingNode.name)));
    }
    walkFormulaNodes(node.args.at(-1), visitor, innerNames);
  }
}
function parseFormulaTree(refersTo) {
  try {
    const tokens = dropBlankTextTokens(tokenizeFormulaText(refersTo)),
      firstToken = tokens[0];
    return tokens.length !== 1 ||
      firstToken.kind !== "call" ||
      validateLetLambdaCall(firstToken) !== "LAMBDA"
      ? undefined
      : firstToken.args
          .slice(0, -1)
          .map((argNode) =>
            validateBindingName(argNode, true).name.replace(/^_xlpm\./i, ""),
          );
  } catch {
    return;
  }
}
const supportedFunctionNames = [
    "ABS",
    "ACOS",
    "ACOSH",
    "ACOT",
    "ACOTH",
    "ADDRESS",
    "AND",
    "ARABIC",
    "AREAS",
    "ASC",
    "ASIN",
    "ASINH",
    "ATAN",
    "ATAN2",
    "ATANH",
    "AVEDEV",
    "AVERAGE",
    "AVERAGEA",
    "AVERAGEIF",
    "BAHTTEXT",
    "BASE",
    "BESSELI",
    "BESSELJ",
    "BESSELK",
    "BESSELY",
    "BETA.DIST",
    "BETA.INV",
    "BIN2DEC",
    "BIN2HEX",
    "BIN2OCT",
    "BINOM.DIST",
    "BINOM.DIST.RANGE",
    "BINOM.INV",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "CEILING",
    "CEILING.MATH",
    "CEILING.PRECISE",
    "CHAR",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CLEAN",
    "CODE",
    "COLUMN",
    "COLUMNS",
    "COMBIN",
    "COMBINA",
    "COMPLEX",
    "CONCAT",
    "CONCATENATE",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "CORREL",
    "COS",
    "COSH",
    "COT",
    "COTH",
    "COUNT",
    "COUNTIF",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "CSC",
    "CSCH",
    "DATE",
    "DATEDIF",
    "DATEVALUE",
    "DAY",
    "DAYS",
    "DAYS360",
    "DBCS",
    "DEC2BIN",
    "DEC2HEX",
    "DEC2OCT",
    "DECIMAL",
    "DEGREES",
    "DELTA",
    "DEVSQ",
    "DOLLAR",
    "EDATE",
    "ENCODEURL",
    "EOMONTH",
    "ERF",
    "ERFC",
    "ERROR.TYPE",
    "EVEN",
    "EXACT",
    "EXP",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FACT",
    "FACTDOUBLE",
    "FALSE",
    "FIND",
    "FINDB",
    "FISHER",
    "FISHERINV",
    "FIXED",
    "FLOOR",
    "FLOOR.MATH",
    "FLOOR.PRECISE",
    "FORECAST",
    "FORECAST.LINEAR",
    "FREQUENCY",
    "GAMMA",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN",
    "GAMMALN.PRECISE",
    "GAUSS",
    "GCD",
    "GEOMEAN",
    "GESTEP",
    "GROWTH",
    "HARMEAN",
    "HEX2BIN",
    "HEX2DEC",
    "HEX2OCT",
    "HLOOKUP",
    "HOUR",
    "HYPGEOM.DIST",
    "IF",
    "IFERROR",
    "IFNA",
    "IFS",
    "IMABS",
    "IMAGINARY",
    "IMARGUMENT",
    "IMCONJUGATE",
    "IMCOS",
    "IMCOSH",
    "IMCOT",
    "IMCSC",
    "IMCSCH",
    "IMDIV",
    "IMEXP",
    "IMLN",
    "IMLOG10",
    "IMLOG2",
    "IMPOWER",
    "IMPRODUCT",
    "IMREAL",
    "IMSEC",
    "IMSECH",
    "IMSIN",
    "IMSINH",
    "IMSQRT",
    "IMSUB",
    "IMSUM",
    "IMTAN",
    "INDEX",
    "INT",
    "INTERCEPT",
    "ISBLANK",
    "ISERR",
    "ISERROR",
    "ISEVEN",
    "ISLOGICAL",
    "ISNA",
    "ISNONTEXT",
    "ISNUMBER",
    "ISO.CEILING",
    "ISOWEEKNUM",
    "ISREF",
    "ISTEXT",
    "KURT",
    "LCM",
    "LEFT",
    "LEFTB",
    "LN",
    "LOG",
    "LOG10",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "LOWER",
    "MDETERM",
    "MID",
    "MIDB",
    "MINUTE",
    "MMULT",
    "MOD",
    "MONTH",
    "MROUND",
    "MULTINOMIAL",
    "MUNIT",
    "N",
    "NA",
    "NEGBINOM.DIST",
    "NETWORKDAYS",
    "NETWORKDAYS.INTL",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "NOT",
    "NOW",
    "NUMBERVALUE",
    "OCT2BIN",
    "OCT2DEC",
    "OCT2HEX",
    "ODD",
    "OR",
    "PHI",
    "PI",
    "POISSON.DIST",
    "POWER",
    "PRODUCT",
    "PROPER",
    "QUOTIENT",
    "RADIANS",
    "RAND",
    "RANDBETWEEN",
    "REPLACE",
    "REPLACEB",
    "REPT",
    "RIGHT",
    "RIGHTB",
    "ROMAN",
    "ROUND",
    "ROUNDDOWN",
    "ROUNDUP",
    "ROW",
    "ROWS",
    "SEARCH",
    "SEARCHB",
    "SEC",
    "SECH",
    "SECOND",
    "SERIESSUM",
    "SIGN",
    "SIN",
    "SINH",
    "SQRT",
    "SQRTPI",
    "STANDARDIZE",
    "SUM",
    "SUMIF",
    "SUMPRODUCT",
    "SUMSQ",
    "SUMX2MY2",
    "SUMX2PY2",
    "SUMXMY2",
    "T",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "TAN",
    "TANH",
    "TEXT",
    "TIME",
    "TIMEVALUE",
    "TODAY",
    "TRANSPOSE",
    "TRIM",
    "TRUE",
    "TRUNC",
    "TYPE",
    "UNICHAR",
    "UNICODE",
    "VLOOKUP",
    "WEEKDAY",
    "WEEKNUM",
    "WEIBULL.DIST",
    "WORKDAY",
    "WORKDAY.INTL",
    "XOR",
    "YEAR",
    "YEARFRAC",
  ],
  extensionRuntimeLimits = Object.freeze({
    extensions: 8,
    functions: 40,
    calls: 100,
    inputCells: 2e3,
    outputCells: 2e3,
    metadataCharacters: 2e6,
    timeoutMs: 15e3,
    maxTimeoutMs: 6e4,
  }),
  boundedTextSchema = (maxLength) =>
    string()
      .min(1)
      .max(maxLength)
      .refine((text) => !/[\u0000-\u001f\u007f]/.test(text)),
  safeIdentifierSchema = string()
    .regex(/^[A-Za-z][A-Za-z0-9_-]{0,63}$/)
    .refine(
      (identifier) =>
        !["constructor", "prototype", "__proto__"].includes(identifier),
    ),
  extensionScalarKindSchema = _enum(["number", "string", "boolean", "scalar"]),
  extensionScalarTypeSchema = strictObject({
    type: extensionScalarKindSchema,
    nullable: boolean().optional(),
  }),
  extensionMatrixTypeSchema = strictObject({
    type: literal("matrix"),
    items: extensionScalarTypeSchema,
    maxRows: number().int().min(1).max(100),
    maxColumns: number().int().min(1).max(100),
  }),
  extensionEntityTypeSchema = strictObject({
    type: literal("entity"),
    fields: record(safeIdentifierSchema, extensionScalarTypeSchema).refine(
      (fields) =>
        Object.keys(fields).length > 0 && Object.keys(fields).length <= 20,
    ),
  }),
  extensionValueTypeSchema = union([
    extensionScalarTypeSchema,
    extensionMatrixTypeSchema,
    extensionEntityTypeSchema,
  ]),
  extensionPermissionSchema = _enum([
    "read-arguments",
    "return-results",
    "linked-data",
    "remote-compute",
  ]),
  extensionManifestSchema = strictObject({
    schemaVersion: literal(1),
    id: safeIdentifierSchema,
    label: boundedTextSchema(120),
    version: string()
      .regex(/^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$/)
      .max(64),
    namespace: string()
      .regex(/^[A-Z][A-Z0-9_]{1,23}$/)
      .refine(
        (namespace) =>
          !["SQL", "REGISTER", "__LG", "_XLFN", "_XLUDF"].includes(namespace),
      ),
    runtime: _enum(["host", "office-addin", "python", "r", "linked-data"]),
    permissions: array(extensionPermissionSchema).min(2).max(4),
    recalculation: strictObject({
      mode: literal("manual"),
      invalidateOn: literal("argument-values"),
      volatile: literal(false),
    }),
    functions: array(
      strictObject({
        id: string().regex(/^[A-Z][A-Z0-9_]{0,39}$/),
        description: boundedTextSchema(300),
        parameters: array(
          strictObject({
            name: safeIdentifierSchema,
            schema: union([
              extensionScalarTypeSchema,
              extensionMatrixTypeSchema,
            ]),
          }),
        ).max(16),
        result: extensionValueTypeSchema,
      }),
    )
      .min(1)
      .max(20),
  });
function validateExtensionManifest(manifestInput) {
  const manifest = extensionManifestSchema.parse(manifestInput);
  if (
    new Set(manifest.permissions).size !== manifest.permissions.length ||
    !manifest.permissions.includes("read-arguments") ||
    !manifest.permissions.includes("return-results")
  )
    throw Error(
      "Extensions must declare distinct read-arguments and return-results permissions.",
    );
  if (
    manifest.runtime !== "host" &&
    !manifest.permissions.includes("remote-compute")
  )
    throw Error("External runtimes must declare remote-compute.");
  if (
    new Set(
      manifest.functions.map((functionDefinition) => functionDefinition.id),
    ).size !== manifest.functions.length
  )
    throw Error("Duplicate custom function names.");
  for (const functionDef of manifest.functions) {
    if (
      supportedFunctionNames.includes(`${manifest.namespace}.${functionDef.id}`)
    )
      throw Error("An extension cannot shadow a built-in function.");
    if (
      new Set(functionDef.parameters.map((parameter) => parameter.name))
        .size !== functionDef.parameters.length
    )
      throw Error("Duplicate extension parameter names.");
    if (
      functionDef.result.type === "entity" &&
      !manifest.permissions.includes("linked-data")
    )
      throw Error("Entity results require the linked-data permission.");
  }
  return manifest;
}
function validateExtensionValue(value, schema) {
  const checkScalar = (scalarValue, scalarSchema) => {
    if (scalarValue === null && scalarSchema.nullable) return null;
    if (
      (typeof scalarValue == "number" &&
        Number.isFinite(scalarValue) &&
        (scalarSchema.type === "number" || scalarSchema.type === "scalar")) ||
      (typeof scalarValue == "string" &&
        scalarValue.length <= 32767 &&
        (scalarSchema.type === "string" || scalarSchema.type === "scalar")) ||
      (typeof scalarValue == "boolean" &&
        (scalarSchema.type === "boolean" || scalarSchema.type === "scalar"))
    )
      return scalarValue;
    throw Error("Extension value does not match its declared schema.");
  };
  if (schema.type === "matrix") {
    if (
      !Array.isArray(value) ||
      !value.length ||
      value.length > schema.maxRows ||
      !Array.isArray(value[0]) ||
      !value[0].length ||
      value[0].length > schema.maxColumns ||
      value.length * value[0].length > extensionRuntimeLimits.outputCells
    )
      throw Error("Extension matrix exceeds its declared shape or cell limit.");
    const columnCount = value[0].length;
    return value.map((row) => {
      if (!Array.isArray(row) || row.length !== columnCount)
        throw Error("Extension matrix must be rectangular.");
      return row.map((cell) => checkScalar(cell, schema.items));
    });
  }
  if (schema.type === "entity") {
    const entity = strictObject({
      kind: literal("entity"),
      id: boundedTextSchema(200),
      label: boundedTextSchema(300),
      fields: record(safeIdentifierSchema, unknown()),
      source: boundedTextSchema(1e3),
      asOf: isoDateTimeSchemaFactory(),
    }).parse(value);
    if (
      Object.keys(entity.fields).sort().join("|") !==
      Object.keys(schema.fields).sort().join("|")
    )
      throw Error(
        "Linked entity fields must match the declared schema exactly.",
      );
    return {
      ...entity,
      fields: Object.fromEntries(
        Object.entries(schema.fields).map(([fieldName, fieldSchema]) => [
          fieldName,
          checkScalar(entity.fields[fieldName], fieldSchema),
        ]),
      ),
    };
  }
  return checkScalar(value, schema);
}
const tD = new WeakSet();
function validateExtensionState(workbook, targetWorkbook) {
  if (workbook.extensions === undefined) return;
  const definedNameSet = new Set(
      (targetWorkbook.definedNames ?? []).map((definedName) =>
        definedName.name.toUpperCase(),
      ),
    ),
    extensions = workbook.extensions;
  if (
    Array.isArray(extensions?.installed) &&
    extensions.installed.some((installedEntry) =>
      installedEntry?.manifest?.functions?.some((customFunction) =>
        definedNameSet.has(
          `${installedEntry.manifest.namespace}.${customFunction.id}`,
        ),
      ),
    )
  )
    throw Error("Extension functions cannot shadow defined names.");
  if (tD.has(workbook.extensions)) {
    targetWorkbook.extensions = workbook.extensions;
    return;
  }
  if (
    JSON.stringify(workbook.extensions).length >
    extensionRuntimeLimits.metadataCharacters
  )
    throw Error("Extension metadata exceeds its limit.");
  const parsedExtensions = strictObject({
      schemaVersion: literal(1),
      installed: array(
        strictObject({ manifest: unknown(), enabled: boolean() }),
      ).max(extensionRuntimeLimits.extensions),
      caches: array(
        strictObject({
          key: string().regex(/^[a-f0-9]{64}$/),
          sheetId: boundedTextSchema(100),
          address: string()
            .regex(/^[A-Z]{1,3}[1-9]\d*$/)
            .refine((address) => {
              const [, columnLetters, rowNumber] = /^([A-Z]+)(\d+)$/.exec(
                address,
              );
              return (
                Number(rowNumber) <= 1048576 &&
                [...columnLetters].reduce(
                  (columnIndex, letter) =>
                    columnIndex * 26 + letter.charCodeAt(0) - 64,
                  0,
                ) <= 16384
              );
            }),
          formula: boundedTextSchema(8192),
          functionName: boundedTextSchema(80),
          arguments: array(unknown()).max(16),
          manifestHash: string().regex(/^[a-f0-9]{64}$/),
          scopeHash: string().regex(/^[a-f0-9]{64}$/),
          result: unknown(),
          calculatedAt: isoDateTimeSchemaFactory(),
          providerRevision: boundedTextSchema(200),
        }),
      ).max(extensionRuntimeLimits.calls),
    }).parse(workbook.extensions),
    installed = parsedExtensions.installed.map((rawEntry) => ({
      ...rawEntry,
      manifest: validateExtensionManifest(rawEntry.manifest),
    }));
  if (
    new Set(installed.map((identityEntry) => identityEntry.manifest.id))
      .size !== installed.length ||
    new Set(
      installed.map((namespaceEntry) => namespaceEntry.manifest.namespace),
    ).size !== installed.length ||
    installed.reduce(
      (total, entry) => total + entry.manifest.functions.length,
      0,
    ) > extensionRuntimeLimits.functions
  )
    throw Error(
      "Extension identities/namespaces must be unique within the function limit.",
    );
  const caches = parsedExtensions.caches.map((cach) => {
    const matchingEntry = installed.find((candidateEntry) =>
        cach.functionName.startsWith(candidateEntry.manifest.namespace + "."),
      ),
      matchingFunction = matchingEntry?.manifest.functions.find(
        (candidateFunction) =>
          `${matchingEntry.manifest.namespace}.${candidateFunction.id}` ===
          cach.functionName,
      );
    if (
      !matchingFunction ||
      !matchingEntry ||
      cach.arguments.length !== matchingFunction.parameters.length
    )
      throw Error("Extension cache does not match an installed function.");
    if (
      cach.key !==
      fingerprintExtensionCall(
        cach.sheetId,
        cach.address,
        cach.formula,
        cach.functionName,
        cach.arguments,
        cach.manifestHash,
      )
    )
      throw Error("Extension cache identity does not match its arguments.");
    return {
      ...cach,
      arguments: cach.arguments.map((argument, argumentIndex) =>
        validateExtensionValue(
          argument,
          matchingFunction.parameters[argumentIndex].schema,
        ),
      ),
      result: validateExtensionValue(cach.result, matchingFunction.result),
    };
  });
  if (new Set(caches.map((event) => event.key)).size !== caches.length)
    throw Error("Duplicate extension cache entries.");
  const validatedState = {
    schemaVersion: 1,
    installed: installed,
    caches: caches,
  };
  (deepFreezeValue(validatedState),
    tD.add(validatedState),
    (targetWorkbook.extensions = validatedState));
}
const hashExtensionManifest = (manifest) =>
    fingerprintCanonicalValue(validateExtensionManifest(manifest)),
  fingerprintExtensionCall = (
    sheetId,
    address,
    formula,
    functionName,
    callArguments,
    manifestHash,
  ) =>
    fingerprintCanonicalValue([
      sheetId,
      address,
      formula,
      functionName,
      callArguments,
      manifestHash,
    ]);
function listInstalledExtensionFunctions(workbook) {
  return (workbook.extensions?.installed ?? []).flatMap((entry) =>
    entry.manifest.functions.map((functionDef) => ({
      entry: entry,
      fn: functionDef,
      name: `${entry.manifest.namespace}.${functionDef.id}`,
    })),
  );
}
const extensionCacheEntriesByWorkbook = new WeakMap();
function setExtensionCacheEntries(workbook, entries) {
  extensionCacheEntriesByWorkbook.set(workbook, structuredClone(entries));
}
function getExtensionCacheEntries(workbook) {
  return extensionCacheEntriesByWorkbook.get(workbook) ?? [];
}
function copyExtensionCacheEntries(workbook, targetWorkbook) {
  const sourceEntries = extensionCacheEntriesByWorkbook.get(workbook);
  sourceEntries &&
    extensionCacheEntriesByWorkbook.set(targetWorkbook, sourceEntries);
}
function hasExtensionCacheForSheet(
  workbook,
  extensionId,
  manifestHash,
  sheetId,
) {
  return getExtensionCacheEntries(workbook).some(
    (cacheEntry) =>
      cacheEntry.id === extensionId &&
      cacheEntry.manifestHash === manifestHash &&
      cacheEntry.sheetIds.includes(sheetId),
  );
}
function resolveExtensionFunctionCall(
  workbook,
  functionName,
  args,
  sheetId,
  address,
) {
  const installedFunction = listInstalledExtensionFunctions(workbook).find(
    (functionEntry) => functionEntry.name === functionName,
  );
  if (!installedFunction) return { error: "#NAME?" };
  const manifestHash = hashExtensionManifest(installedFunction.entry.manifest);
  if (
    !installedFunction.entry.enabled ||
    !hasExtensionCacheForSheet(
      workbook,
      installedFunction.entry.manifest.id,
      manifestHash,
      sheetId,
    )
  )
    return { error: "#NAME?" };
  if (args.length !== installedFunction.fn.parameters.length)
    return { error: "#VALUE!" };
  let validatedArgs;
  try {
    validatedArgs = args.map((arg, argumentIndex) => {
      const schema = installedFunction.fn.parameters[argumentIndex].schema;
      return validateExtensionValue(
        schema.type === "matrix" && !Array.isArray(arg) ? [[arg]] : arg,
        schema,
      );
    });
  } catch {
    return { error: "#VALUE!" };
  }
  const formulaText = String(
      workbook.sheets.find((sheet) => sheet.id === sheetId)?.cells[address]
        ?.raw ?? "",
    ),
    callKey = fingerprintExtensionCall(
      sheetId,
      address,
      formulaText,
      functionName,
      validatedArgs,
      manifestHash,
    ),
    cachedCall = workbook.extensions?.caches.find(
      (cach) => cach.key === callKey,
    );
  return cachedCall &&
    !getExtensionCacheEntries(workbook).some(
      (cacheEntry) =>
        cacheEntry.id === installedFunction.entry.manifest.id &&
        cacheEntry.scopeHash === cachedCall.scopeHash,
    )
    ? { error: "#NAME?" }
    : cachedCall
      ? { value: cachedCall.result }
      : { error: "#PENDING!" };
}
const calculationBudgetDefaults = Object.freeze({
  defaultMemoryBudgetMiB: 32,
  maxMemoryBudgetMiB: 128,
  defaultWorkBudget: 1e6,
  maxWorkBudget: 1e7,
  maxNesting: 64,
  memoryUnit:
    "Cumulative accounted working data: UTF-16 text, 8-byte scalars and container allowances, including repeated intermediate allocations. This is not a browser heap or process RSS limit.",
  timing:
    "Isolated calculations terminate their worker on deadline or cancellation. Delivery of timer events is subject to host scheduling; synchronous APIs use cooperative checks.",
  scope:
    "Calculation inputs, intermediate values and results. The existing committed workbook, UI, parser code and runtime overhead are outside the working-data budget.",
});
function readMonotonicClockMs() {
  return typeof performance < "u" ? performance.now() : Date.now();
}
function resolveCalculationBudget(calculation) {
  return {
    memoryBudgetBytes:
      (calculation?.memoryBudgetMiB ??
        calculationBudgetDefaults.defaultMemoryBudgetMiB) *
      1024 *
      1024,
    workBudget:
      calculation?.workBudget ?? calculationBudgetDefaults.defaultWorkBudget,
  };
}
function estimateValueFootprintBytes(value, memoryBudgetBytes) {
  let footprintBytes = 0;
  const visitValue = (currentValue, depth) => {
    if (!(footprintBytes > memoryBudgetBytes)) {
      if (depth > 80) {
        footprintBytes = memoryBudgetBytes + 1;
        return;
      }
      if (typeof currentValue == "string")
        footprintBytes += 24 + currentValue.length * 2;
      else if (!currentValue || typeof currentValue != "object")
        footprintBytes += 8;
      else if (Array.isArray(currentValue)) {
        footprintBytes += 32 + currentValue.length * 8;
        for (const arrayItem of currentValue)
          if (
            (visitValue(arrayItem, depth + 1),
            footprintBytes > memoryBudgetBytes)
          )
            break;
      } else {
        footprintBytes += 64;
        for (const propertyName in currentValue) {
          if (!Object.prototype.hasOwnProperty.call(currentValue, propertyName))
            continue;
          footprintBytes += 32 + propertyName.length * 2;
          const descriptor = Object.getOwnPropertyDescriptor(
            currentValue,
            propertyName,
          );
          if (!("value" in descriptor)) {
            footprintBytes = memoryBudgetBytes + 1;
            break;
          }
          if (
            (visitValue(descriptor.value, depth + 1),
            footprintBytes > memoryBudgetBytes)
          )
            break;
        }
      }
    }
  };
  return (visitValue(value, 0), footprintBytes);
}
const rD = new WeakMap();
function estimateWorkbookMemoryBytes(workbook, memoryBudgetBytes) {
  let totalBytes = 1024;
  const addValue = (value) => {
    totalBytes += estimateValueFootprintBytes(
      value,
      Math.max(0, memoryBudgetBytes - totalBytes),
    );
  };
  (addValue(workbook.definedNames),
    addValue(workbook.workbookLinks),
    addValue(workbook.extensions),
    addValue(workbook.calculation),
    addValue(workbook.datePolicy));
  for (const sheet of workbook.sheets) {
    addValue([
      sheet.id,
      sheet.name,
      sheet.tables,
      sheet.merges,
      sheet.hiddenRows,
      sheet.hiddenCols,
      sheet.outlineGroups,
    ]);
    const cellsAreFrozen = isDeepFrozenValue(sheet.cells),
      cachedCellBytes = cellsAreFrozen ? rD.get(sheet.cells) : undefined;
    if (cachedCellBytes !== undefined) totalBytes += cachedCellBytes;
    else {
      const bytesBeforeCells = totalBytes;
      for (const cellAddress in sheet.cells)
        if (
          ((totalBytes += 128 + 2 * cellAddress.length),
          addValue(sheet.cells[cellAddress].raw),
          totalBytes > memoryBudgetBytes)
        )
          return totalBytes;
      cellsAreFrozen && rD.set(sheet.cells, totalBytes - bytesBeforeCells);
    }
    if (totalBytes > memoryBudgetBytes) return totalBytes;
  }
  return totalBytes;
}
function createCalculationBudgetTracker(workbook) {
  const calculationBudget = resolveCalculationBudget(workbook.calculation);
  let accountedBytes = estimateWorkbookMemoryBytes(
      workbook,
      calculationBudget.memoryBudgetBytes,
    ),
    workUnits = 0,
    stopped,
    reason,
    startedAt = 0,
    elapsedMs = 0,
    nestingDepth = 0;
  const check = () => (
    !stopped &&
      accountedBytes > calculationBudget.memoryBudgetBytes &&
      ((stopped = "#LIMIT!"), (reason = "Working-data memory budget exceeded")),
    !stopped &&
      workUnits > calculationBudget.workBudget &&
      ((stopped = "#LIMIT!"), (reason = "Calculation work budget exceeded")),
    !stopped &&
      elapsedMs + (nestingDepth ? readMonotonicClockMs() - startedAt : 0) >
        (workbook.calculation?.timeBudgetMs ?? 2500) &&
      ((stopped = "#TIMEOUT!"),
      (reason = "Cooperative calculation time budget exceeded")),
    stopped
  );
  return {
    begin() {
      nestingDepth++ === 0 && (startedAt = readMonotonicClockMs());
    },
    end() {
      nestingDepth > 0 &&
        --nestingDepth === 0 &&
        (elapsedMs += readMonotonicClockMs() - startedAt);
    },
    check: check,
    work(units = 1) {
      return (
        stopped ||
        (workUnits + Math.max(0, units) > calculationBudget.workBudget
          ? ((stopped = "#LIMIT!"),
            (reason = "Calculation work budget exceeded"),
            stopped)
          : ((workUnits += Math.max(0, units)), check()))
      );
    },
    reserve(reservedBytes) {
      return (
        stopped || ((accountedBytes += Math.max(0, reservedBytes)), check())
      );
    },
    value(value) {
      return (
        stopped ||
        ((accountedBytes += estimateValueFootprintBytes(
          value,
          Math.max(0, calculationBudget.memoryBudgetBytes - accountedBytes),
        )),
        check())
      );
    },
    report: () => ({
      ...calculationBudget,
      accountedBytes: accountedBytes,
      workUnits: workUnits,
      ...(stopped ? { stopped: stopped, reason: reason } : {}),
    }),
  };
}
function isWithinNestingLimit(formulaText) {
  const strippedText = formulaText.replace(
    /"(?:[^"]|"")*"|'(?:[^']|'')*'/g,
    "",
  );
  let nestingDepth = 0;
  for (const character of strippedText) {
    if (
      ((character === "(" || character === "{") && nestingDepth++,
      nestingDepth > calculationBudgetDefaults.maxNesting)
    )
      return false;
    (character === ")" || character === "}") && nestingDepth--;
  }
  return true;
}
const dateSystemSerialOffset = (dateSystem = "1900") =>
    dateSystem === "1904" ? 1462 : 0,
  dateSystemMaxSerial = (dateSystem = "1900") =>
    2958465 - dateSystemSerialOffset(dateSystem),
  defaultDatePolicy = Object.freeze({
    calendar: "gregorian",
    dateSystem: "1900",
    textDateOrder: "MDY",
    timeZone: "UTC",
  }),
  dateCalculationLimits = Object.freeze({
    maxSerial: 2958465,
    holidays: 1e3,
    businessDays: 366e3,
  });
function buildCalculationClock(assistantSelectionContext) {
  const datePolicy = assistantSelectionContext.datePolicy,
    calculationClock = assistantSelectionContext.calculationClock;
  return {
    ...(datePolicy
      ? {
          datePolicy: {
            calendar: datePolicy.calendar,
            dateSystem: datePolicy.dateSystem,
            textDateOrder: datePolicy.textDateOrder,
            timeZone: datePolicy.timeZone,
          },
        }
      : {}),
    ...(calculationClock
      ? {
          calculationClock: {
            calendar: calculationClock.calendar,
            dateSystem: calculationClock.dateSystem,
            textDateOrder: calculationClock.textDateOrder,
            timeZone: calculationClock.timeZone,
            calculatedAt: calculationClock.calculatedAt,
            localDateTime: calculationClock.localDateTime,
            utcOffsetMinutes: calculationClock.utcOffsetMinutes,
            nowSerial: calculationClock.nowSerial,
            todaySerial: calculationClock.todaySerial,
          },
        }
      : {}),
  };
}
function workbookUsesNowOrToday(workbook) {
  const isNowOrTodayFormula = (formulaText) =>
    /\b(?:NOW|TODAY)\s*\(/i.test(formulaText.replace(/"(?:[^"]|"")*"/g, '""'));
  return (
    workbook.sheets.some((sheet) =>
      Object.values(sheet.cells).some(
        (cell) =>
          !cell.literal &&
          typeof cell.raw == "string" &&
          cell.raw.startsWith("=") &&
          isNowOrTodayFormula(cell.raw),
      ),
    ) ||
    !!workbook.definedNames?.some((definedName) =>
      isNowOrTodayFormula(definedName.refersTo),
    )
  );
}
const millisecondsPerDay = 864e5,
  serialZeroUtcMs = Date.UTC(1899, 11, 31);
function normalizeDatePolicy(datePolicy = defaultDatePolicy) {
  const policy = datePolicy;
  if (
    !policy ||
    typeof policy != "object" ||
    Array.isArray(policy) ||
    Object.keys(policy).some(
      (key) =>
        !["calendar", "dateSystem", "textDateOrder", "timeZone"].includes(key),
    ) ||
    policy.calendar !== "gregorian" ||
    !["1900", "1904"].includes(policy.dateSystem) ||
    !["MDY", "DMY"].includes(policy.textDateOrder) ||
    typeof policy.timeZone != "string" ||
    policy.timeZone.length > 100 ||
    !/^[A-Za-z_]+(?:\/[A-Za-z0-9_+.-]+)*$/.test(policy.timeZone)
  )
    throw Error(
      "Unsupported date policy. Use the Gregorian calendar, 1900 or 1904 serial system, MDY or DMY text dates, and a valid IANA time zone (or UTC).",
    );
  try {
    const timeZone = new Intl.DateTimeFormat("en-US", {
      timeZone: policy.timeZone,
    }).resolvedOptions().timeZone;
    return {
      calendar: "gregorian",
      dateSystem: policy.dateSystem,
      textDateOrder: policy.textDateOrder,
      timeZone: timeZone,
    };
  } catch {
    throw Error(
      `Unsupported time zone: ${policy.timeZone}. Use an IANA time zone or UTC.`,
    );
  }
}
const isExcelLeapYear = (year) =>
    year === 1900 || (year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0)),
  daysInExcelMonth = (year, month) =>
    month === 2
      ? isExcelLeapYear(year)
        ? 29
        : 28
      : [4, 6, 9, 11].includes(month)
        ? 30
        : 31;
function dateToSerial(year, month, day, dateSystem = "1900") {
  if (dateSystem === "1904")
    return (
      (Date.UTC(year, month - 1, day) - Date.UTC(1904, 0, 1)) /
      millisecondsPerDay
    );
  if (year === 1900 && month === 2 && day === 29) return 60;
  const utcMilliseconds = Date.UTC(year, month - 1, day);
  return (
    (utcMilliseconds - serialZeroUtcMs) / millisecondsPerDay +
    (utcMilliseconds >= Date.UTC(1900, 2, 1) ? 1 : 0)
  );
}
function serialToDateParts(serialNumber, dateSystem = "1900") {
  const wholeSerial = Math.floor(serialNumber);
  if (dateSystem === "1900" && wholeSerial === 0)
    return { year: 1900, month: 1, day: 0 };
  if (dateSystem === "1900" && wholeSerial === 60)
    return { year: 1900, month: 2, day: 29 };
  const serialDate = new Date(
    dateSystem === "1904"
      ? Date.UTC(1904, 0, 1) + wholeSerial * millisecondsPerDay
      : serialZeroUtcMs +
          (wholeSerial - (wholeSerial > 60 ? 1 : 0)) * millisecondsPerDay,
  );
  return {
    year: serialDate.getUTCFullYear(),
    month: serialDate.getUTCMonth() + 1,
    day: serialDate.getUTCDate(),
  };
}
function nD(timeText) {
  const match =
    /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d{1,9})?))?\s*(AM|PM)?$/i.exec(timeText);
  if (!match) return;
  let hours = Number(match[1]);
  const minutes = Number(match[2]),
    seconds = Number(match[3] ?? 0),
    meridiem = match[4]?.toUpperCase();
  if (!(
    minutes > 59 ||
    seconds >= 60 ||
    (meridiem ? hours < 1 || hours > 12 : hours > 23)
  ))
    return (
      meridiem && (hours = (hours % 12) + (meridiem === "PM" ? 12 : 0)),
      (hours * 3600 + minutes * 60 + seconds) / 86400
    );
}
function parseDateText(text, datePolicy = defaultDatePolicy) {
  const trimmedText = text.trim(),
    match =
      /^(?:(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})\/(\d{1,2})\/(\d{4}))(?:[ T](.+))?$/.exec(
        trimmedText,
      );
  if (!match) return nD(trimmedText);
  const year = Number(match[1] ?? match[6]),
    month = Number(
      match[2] ?? match[datePolicy.textDateOrder === "MDY" ? 4 : 5],
    ),
    day = Number(match[3] ?? match[datePolicy.textDateOrder === "MDY" ? 5 : 4]);
  if (
    year < (datePolicy.dateSystem === "1904" ? 1904 : 1900) ||
    year > 9999 ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > daysInExcelMonth(year, month)
  )
    return;
  const timeFraction = match[7] ? nD(match[7]) : 0;
  return timeFraction === undefined
    ? undefined
    : dateToSerial(year, month, day, datePolicy.dateSystem) + timeFraction;
}
function resolveLocalDateTime(policyInput, instant = new Date()) {
  const datePolicy = normalizeDatePolicy(policyInput),
    calculatedAt = instant.toISOString(),
    zonedParts = new Intl.DateTimeFormat("en-US-u-ca-gregory-nu-latn", {
      timeZone: datePolicy.timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23",
    }).formatToParts(instant),
    readPart = (partType) =>
      Number(zonedParts.find((part) => part.type === partType)?.value),
    year = readPart("year"),
    month = readPart("month"),
    day = readPart("day"),
    hour = readPart("hour"),
    minute = readPart("minute"),
    second = readPart("second"),
    milliseconds = instant.getUTCMilliseconds(),
    todaySerial = dateToSerial(year, month, day, datePolicy.dateSystem);
  if (
    todaySerial < 0 ||
    todaySerial > dateSystemMaxSerial(datePolicy.dateSystem)
  )
    throw Error(
      "Calculation clock is outside the supported 1900–9999 date range.",
    );
  const localDateTime = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}T${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}:${String(second).padStart(2, "0")}.${String(milliseconds).padStart(3, "0")}`;
  return Object.freeze({
    ...datePolicy,
    calculatedAt: calculatedAt,
    localDateTime: localDateTime,
    utcOffsetMinutes:
      (Date.UTC(year, month - 1, day, hour, minute, second, milliseconds) -
        instant.getTime()) /
      6e4,
    todaySerial: todaySerial,
    nowSerial:
      todaySerial +
      (hour * 3600 + minute * 60 + second + milliseconds / 1e3) / 86400,
  });
}
const calculationFeatureSupport = Object.freeze({
  supported: ["RAND", "RANDBETWEEN", "NOW", "TODAY"],
  unsupported: ["RANDARRAY", "OFFSET", "INDIRECT", "CELL", "INFO"],
  refresh:
    "Automatic committed revisions and explicit Calculate Now refresh volatile formulas and their dependents. Repeated reads, selection and rendering reuse the same snapshot. Manual edits leave affected results pending until Calculate Now.",
  random:
    "RAND and RANDBETWEEN use the runtime pseudorandom generator without a workbook seed. They are not cryptographic and are not reproducible audit samples. Separate seeded audit sampling never consumes spreadsheet random state.",
  clock:
    "NOW and TODAY share one captured instant and configured time zone per calculation; no background timer refreshes the workbook.",
  inspection:
    "Diagnostic formula evaluation uses a separate snapshot; replay does not redraw random values. Native reopening starts a new calculation; XLSX formula caches are historical only.",
});
function findVolatileFunctionNames(formulaText) {
  const strippedText = formulaText.replace(
    /"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g,
    "",
  );
  return [
    ...new Set(
      [
        ...strippedText.matchAll(
          /\b(?:_xlfn\.)?(RAND|RANDBETWEEN|NOW|TODAY)\s*\(/gi,
        ),
      ].map((match) => match[1].toUpperCase()),
    ),
  ];
}
function workbookHasVolatileFunctions(workbook) {
  return (
    !!workbook.definedNames?.some(
      (definedName) => findVolatileFunctionNames(definedName.refersTo).length,
    ) ||
    workbook.sheets.some((sheet) =>
      Object.values(sheet.cells).some(
        (cell) =>
          !cell.literal &&
          typeof cell.raw == "string" &&
          cell.raw.startsWith("=") &&
          findVolatileFunctionNames(cell.raw).length,
      ),
    )
  );
}
const dateFunctionNames = [
    "DATE",
    "DATEVALUE",
    "DAY",
    "MONTH",
    "YEAR",
    "EDATE",
    "EOMONTH",
    "DATEDIF",
    "DAYS",
    "DAYS360",
    "YEARFRAC",
    "WORKDAY",
    "WORKDAY.INTL",
    "NETWORKDAYS",
    "NETWORKDAYS.INTL",
    "WEEKDAY",
    "WEEKNUM",
    "ISOWEEKNUM",
    "TODAY",
    "NOW",
    "TIME",
    "TIMEVALUE",
    "HOUR",
    "MINUTE",
    "SECOND",
  ],
  notDateFunctionSentinel = Symbol("not-date-function"),
  UJ = {
    DATE: [3, 3],
    DATEVALUE: [1, 1],
    DAY: [1, 1],
    MONTH: [1, 1],
    YEAR: [1, 1],
    EDATE: [2, 2],
    EOMONTH: [2, 2],
    DATEDIF: [3, 3],
    DAYS: [2, 2],
    DAYS360: [2, 3],
    YEARFRAC: [2, 3],
    WORKDAY: [2, 3],
    "WORKDAY.INTL": [2, 4],
    NETWORKDAYS: [2, 3],
    "NETWORKDAYS.INTL": [2, 4],
    WEEKDAY: [1, 2],
    WEEKNUM: [1, 2],
    ISOWEEKNUM: [1, 1],
    TODAY: [0, 0],
    NOW: [0, 0],
    TIME: [3, 3],
    TIMEVALUE: [1, 1],
    HOUR: [1, 1],
    MINUTE: [1, 1],
    SECOND: [1, 1],
  };
function evaluateDateFunction(functionName, argNodes, context) {
  if (!dateFunctionNames.includes(functionName)) return notDateFunctionSentinel;
  const dateSystem = context.clock.dateSystem,
    maxSerial = dateSystemMaxSerial(dateSystem),
    serialFromParts = (year, month, day) =>
      dateToSerial(year, month, day, dateSystem),
    datePartsOf = (serial) => serialToDateParts(serial, dateSystem),
    weekdayIndexOf = (dateSerial) =>
      (((dateSerial + dateSystemSerialOffset(dateSystem) + 5) % 7) + 7) % 7,
    throwIssue = (errorCode = "#VALUE!") => {
      throw context.issue(errorCode);
    },
    [minArgs, maxArgs] = UJ[functionName];
  (argNodes.length < minArgs || argNodes.length > maxArgs) && throwIssue();
  const argValues = argNodes.map(context.retrieve),
    unwrapScalar = (value) => {
      if (context.isIssue(value)) throw value;
      return Array.isArray(value)
        ? value.length === 1
          ? unwrapScalar(value[0])
          : throwIssue()
        : value;
    },
    toNumber = (rawValue) => {
      if (((rawValue = unwrapScalar(rawValue)), rawValue == null)) return 0;
      if (typeof rawValue == "boolean") return Number(rawValue);
      typeof rawValue != "number" &&
        (typeof rawValue != "string" ||
          !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(rawValue)) &&
        throwIssue();
      const numberValue = Number(rawValue);
      return Number.isFinite(numberValue) ? numberValue : throwIssue("#NUM!");
    },
    checkSerial = (serialNumber) =>
      Number.isFinite(serialNumber) &&
      serialNumber >= 0 &&
      serialNumber < maxSerial + 1
        ? serialNumber
        : throwIssue("#NUM!"),
    coerceSerial = (dateInput) => {
      dateInput = unwrapScalar(dateInput);
      const parsedSerial =
        typeof dateInput == "string" &&
        !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(dateInput)
          ? parseDateText(dateInput, context.clock)
          : toNumber(dateInput);
      return parsedSerial === undefined
        ? throwIssue()
        : checkSerial(parsedSerial);
    },
    toDaySerial = (dateValue) => Math.floor(coerceSerial(dateValue)),
    shiftMonths = (baseSerial, monthOffset, isEndOfMonth = false) => {
      const baseParts = datePartsOf(baseSerial),
        monthIndex = baseParts.year * 12 + baseParts.month - 1 + monthOffset,
        targetYear = Math.floor(monthIndex / 12),
        targetMonth = (((monthIndex % 12) + 12) % 12) + 1;
      return (
        (targetYear < (dateSystem === "1904" ? 1904 : 1900) ||
          targetYear > 9999) &&
          throwIssue("#NUM!"),
        checkSerial(
          serialFromParts(
            targetYear,
            targetMonth,
            isEndOfMonth
              ? daysInExcelMonth(targetYear, targetMonth)
              : Math.min(
                  baseParts.day,
                  daysInExcelMonth(targetYear, targetMonth),
                ),
          ),
        )
      );
    };
  if (functionName === "TODAY") return context.clock.todaySerial;
  if (functionName === "NOW") return context.clock.nowSerial;
  if (functionName === "DATE") {
    let dateYear = Math.trunc(toNumber(argValues[0]));
    const dateMonth = Math.trunc(toNumber(argValues[1])),
      dateDay = Math.trunc(toNumber(argValues[2]));
    ((dateYear < 0 || dateYear >= 1e4) && throwIssue("#NUM!"),
      dateYear < 1900 && (dateYear += 1900));
    const totalMonths = dateYear * 12 + dateMonth - 1;
    dateYear = Math.floor(totalMonths / 12);
    const normalizedMonth = (((totalMonths % 12) + 12) % 12) + 1;
    return (
      (dateYear < 1899 ||
        dateYear > 1e4 ||
        Math.abs(dateDay) > dateCalculationLimits.maxSerial) &&
        throwIssue("#NUM!"),
      checkSerial(serialFromParts(dateYear, normalizedMonth, 1) + dateDay - 1)
    );
  }
  if (functionName === "DATEVALUE" || functionName === "TIMEVALUE") {
    const textValue = unwrapScalar(argValues[0]);
    typeof textValue != "string" && throwIssue();
    const dateText = parseDateText(textValue, context.clock);
    return (
      dateText === undefined && throwIssue(),
      functionName === "DATEVALUE" &&
        !/^(?:\d{4}-\d{1,2}-\d{1,2}|\d{1,2}\/\d{1,2}\/\d{4})(?:[ T]|$)/.test(
          textValue.trim(),
        ) &&
        throwIssue(),
      functionName === "DATEVALUE"
        ? Math.floor(dateText)
        : dateText - Math.floor(dateText)
    );
  }
  if (functionName === "TIME") {
    const [hours, minutes, seconds] = argValues.map((argValue) =>
      Math.trunc(toNumber(argValue)),
    );
    return (
      [hours, minutes, seconds].some(
        (component) => component < 0 || component > 32767,
      ) && throwIssue("#NUM!"),
      ((hours * 3600 + minutes * 60 + seconds) % 86400) / 86400
    );
  }
  if (["HOUR", "MINUTE", "SECOND"].includes(functionName)) {
    const timeSerial = coerceSerial(argValues[0]),
      tolerance = Math.max(1e-9, Number.EPSILON * timeSerial * 86400),
      secondsOfDay =
        Math.floor((timeSerial - Math.floor(timeSerial)) * 86400 + tolerance) %
        86400;
    return functionName === "HOUR"
      ? Math.floor(secondsOfDay / 3600)
      : functionName === "MINUTE"
        ? Math.floor(secondsOfDay / 60) % 60
        : secondsOfDay % 60;
  }
  if (["DAY", "MONTH", "YEAR"].includes(functionName))
    return datePartsOf(toDaySerial(argValues[0]))[functionName.toLowerCase()];
  if (functionName === "EDATE" || functionName === "EOMONTH")
    return shiftMonths(
      toDaySerial(argValues[0]),
      Math.trunc(toNumber(argValues[1])),
      functionName === "EOMONTH",
    );
  if (functionName === "DAYS")
    return toDaySerial(argValues[0]) - toDaySerial(argValues[1]);
  if (functionName === "WEEKDAY") {
    const weekdaySerial = toDaySerial(argValues[0]),
      returnType =
        argValues[1] == null ? 1 : Math.trunc(toNumber(argValues[1])),
      weekdayIndex = weekdayIndexOf(weekdaySerial);
    if (returnType === 3) return weekdayIndex;
    const weekStartIndex =
      returnType === 1 || returnType === 17
        ? 6
        : returnType === 2 || returnType === 11
          ? 0
          : returnType >= 12 && returnType <= 16
            ? returnType - 11
            : throwIssue("#NUM!");
    return ((weekdayIndex - weekStartIndex + 7) % 7) + 1;
  }
  if (functionName === "WEEKNUM" || functionName === "ISOWEEKNUM") {
    const weekSerial = toDaySerial(argValues[0]),
      weekMode =
        functionName === "ISOWEEKNUM"
          ? 21
          : argValues[1] == null
            ? 1
            : Math.trunc(toNumber(argValues[1]));
    if (weekMode === 21) {
      const thursdaySerial = weekSerial + 3 - weekdayIndexOf(weekSerial),
        year = datePartsOf(thursdaySerial).year;
      return Math.floor((thursdaySerial - serialFromParts(year, 1, 1)) / 7) + 1;
    }
    const startDayIndex =
        weekMode === 1 || weekMode === 17
          ? 6
          : weekMode === 2 || weekMode === 11
            ? 0
            : weekMode >= 12 && weekMode <= 16
              ? weekMode - 11
              : throwIssue("#NUM!"),
      yearStartSerial = serialFromParts(datePartsOf(weekSerial).year, 1, 1);
    return (
      Math.floor(
        (weekSerial -
          yearStartSerial +
          ((weekdayIndexOf(yearStartSerial) - startDayIndex + 7) % 7)) /
          7,
      ) + 1
    );
  }
  if (functionName === "DATEDIF") {
    const startSerial = toDaySerial(argValues[0]),
      endSerial = toDaySerial(argValues[1]),
      unit = String(unwrapScalar(argValues[2])).toUpperCase();
    startSerial > endSerial && throwIssue("#NUM!");
    const startParts = datePartsOf(startSerial),
      endParts = datePartsOf(endSerial),
      monthsBetween =
        (endParts.year - startParts.year) * 12 +
        endParts.month -
        startParts.month -
        (endParts.day < startParts.day ? 1 : 0);
    if (unit === "D") return endSerial - startSerial;
    if (unit === "Y") return Math.floor(monthsBetween / 12);
    if (unit === "M") return monthsBetween;
    if (unit === "YM") return monthsBetween % 12;
    if (unit === "MD")
      return endSerial - shiftMonths(startSerial, monthsBetween);
    if (unit === "YD") {
      const anniversaryYear =
        endParts.year -
        (endParts.month < startParts.month ||
        (endParts.month === startParts.month && endParts.day < startParts.day)
          ? 1
          : 0);
      return (
        endSerial -
        serialFromParts(
          anniversaryYear,
          startParts.month,
          Math.min(
            startParts.day,
            daysInExcelMonth(anniversaryYear, startParts.month),
          ),
        )
      );
    }
    throwIssue("#NUM!");
  }
  const days360 = (fromSerial, toSerial, isEuropeanMethod) => {
    const fromParts = datePartsOf(fromSerial),
      toParts = datePartsOf(toSerial);
    let day = fromParts.day,
      endDay = toParts.day;
    return (
      isEuropeanMethod
        ? ((day = Math.min(day, 30)), (endDay = Math.min(endDay, 30)))
        : (day === daysInExcelMonth(fromParts.year, fromParts.month) &&
            (day = 30),
          endDay === 31 && day >= 30 && (endDay = 30),
          fromParts.month === 2 &&
            fromParts.day === daysInExcelMonth(fromParts.year, 2) &&
            toParts.month === 2 &&
            endDay === daysInExcelMonth(toParts.year, 2) &&
            (endDay = 30)),
      (toParts.year - fromParts.year) * 360 +
        (toParts.month - fromParts.month) * 30 +
        endDay -
        day
    );
  };
  if (functionName === "DAYS360") {
    const methodValue = unwrapScalar(argValues[2] ?? false);
    return (
      typeof methodValue != "boolean" &&
        typeof methodValue != "number" &&
        throwIssue(),
      days360(
        toDaySerial(argValues[0]),
        toDaySerial(argValues[1]),
        !!methodValue,
      )
    );
  }
  if (functionName === "YEARFRAC") {
    let yearfracStart = toDaySerial(argValues[0]),
      yearfracEnd = toDaySerial(argValues[1]);
    const basis = argValues[2] == null ? 0 : Math.trunc(toNumber(argValues[2]));
    if (
      ((basis < 0 || basis > 4) && throwIssue("#NUM!"),
      yearfracStart > yearfracEnd &&
        ([yearfracStart, yearfracEnd] = [yearfracEnd, yearfracStart]),
      yearfracStart === yearfracEnd)
    )
      return 0;
    if (basis === 0 || basis === 4)
      return days360(yearfracStart, yearfracEnd, basis === 4) / 360;
    if (basis === 2 || basis === 3)
      return (yearfracEnd - yearfracStart) / (basis === 2 ? 360 : 365);
    const startDateParts = datePartsOf(yearfracStart),
      endDateParts = datePartsOf(yearfracEnd);
    let yearLength = 365;
    const oneYearLater = serialFromParts(
      startDateParts.year + 1,
      startDateParts.month,
      Math.min(
        startDateParts.day,
        daysInExcelMonth(startDateParts.year + 1, startDateParts.month),
      ),
    );
    if (startDateParts.year === endDateParts.year)
      yearLength = isExcelLeapYear(startDateParts.year) ? 366 : 365;
    else if (yearfracEnd <= oneYearLater)
      for (let year = startDateParts.year; year <= endDateParts.year; year++)
        isExcelLeapYear(year) &&
          serialFromParts(year, 2, 29) >= yearfracStart &&
          serialFromParts(year, 2, 29) <= yearfracEnd &&
          (yearLength = 366);
    else
      yearLength =
        (serialFromParts(endDateParts.year + 1, 1, 1) -
          serialFromParts(startDateParts.year, 1, 1)) /
        (endDateParts.year - startDateParts.year + 1);
    return (yearfracEnd - yearfracStart) / yearLength;
  }
  if (
    functionName.startsWith("WORKDAY") ||
    functionName.startsWith("NETWORKDAYS")
  ) {
    const isNetworkDays = functionName.startsWith("NETWORK"),
      isIntl = functionName.endsWith(".INTL"),
      originSerial = toDaySerial(argValues[0]),
      weekendSpec = unwrapScalar(isIntl ? (argValues[2] ?? 1) : 1);
    let weekendMask;
    if (typeof weekendSpec == "string")
      (/^[01]{7}$/.test(weekendSpec) || throwIssue(),
        (weekendMask = weekendSpec));
    else {
      const weekendCode = toNumber(weekendSpec);
      let weekendDays;
      if (weekendCode >= 1 && weekendCode <= 7 && Number.isInteger(weekendCode))
        weekendDays = [(weekendCode + 4) % 7, (weekendCode + 5) % 7];
      else if (
        weekendCode >= 11 &&
        weekendCode <= 17 &&
        Number.isInteger(weekendCode)
      )
        weekendDays = [(weekendCode - 5) % 7];
      else return throwIssue("#NUM!");
      weekendMask = Array.from({ length: 7 }, (G, dayIndex) =>
        weekendDays.includes(dayIndex) ? "1" : "0",
      ).join("");
    }
    const flattenValues = (entry) =>
        Array.isArray(entry) ? entry.flatMap(flattenValues) : [entry],
      holidayValues =
        argValues[isIntl ? 3 : 2] === undefined
          ? []
          : flattenValues(argValues[isIntl ? 3 : 2]);
    holidayValues.length > dateCalculationLimits.holidays &&
      throwIssue("#LIMIT!");
    const holidaySerials = new Set();
    for (const holidayValue of holidayValues)
      holidayValue == null ||
        holidayValue === "" ||
        (typeof holidayValue == "boolean" && throwIssue(),
        holidaySerials.add(toDaySerial(holidayValue)));
    const workdaysPerWeek = [...weekendMask].filter(
        (maskChar) => maskChar === "0",
      ).length,
      isWorkday = (workdaySerial) =>
        weekendMask[weekdayIndexOf(workdaySerial)] === "0",
      countWorkdays = (rangeStart, rangeEnd) => {
        if (rangeStart > rangeEnd) return 0;
        const spanDays = rangeEnd - rangeStart + 1;
        let workdayTotal = Math.floor(spanDays / 7) * workdaysPerWeek;
        for (let offset = 0; offset < spanDays % 7; offset++)
          isWorkday(rangeStart + offset) && workdayTotal++;
        for (const holidaySerial of holidaySerials)
          holidaySerial >= rangeStart &&
            holidaySerial <= rangeEnd &&
            isWorkday(holidaySerial) &&
            workdayTotal--;
        return workdayTotal;
      };
    if (isNetworkDays) {
      let networkEndSerial = toDaySerial(argValues[1]);
      return (
        Math.abs(networkEndSerial - originSerial) >
          dateCalculationLimits.businessDays && throwIssue("#LIMIT!"),
        networkEndSerial < originSerial
          ? -countWorkdays(networkEndSerial, originSerial)
          : countWorkdays(originSerial, networkEndSerial)
      );
    }
    const workdayOffset = Math.trunc(toNumber(argValues[1]));
    if (
      (workdaysPerWeek || throwIssue(),
      Math.abs(workdayOffset) > dateCalculationLimits.businessDays &&
        throwIssue("#LIMIT!"),
      workdayOffset === 0)
    )
      return originSerial;
    const isForward = workdayOffset > 0,
      stepCount = Math.abs(workdayOffset),
      searchLimit = isForward ? maxSerial - originSerial : originSerial,
      countWorkdaysUpTo = (steps) =>
        isForward
          ? countWorkdays(originSerial + 1, originSerial + steps)
          : countWorkdays(originSerial - steps, originSerial - 1);
    countWorkdaysUpTo(searchLimit) < stepCount && throwIssue("#NUM!");
    let lowBound = 1,
      highBound = searchLimit;
    for (; lowBound < highBound;) {
      const midStep = Math.floor((lowBound + highBound) / 2);
      countWorkdaysUpTo(midStep) >= stepCount
        ? (highBound = midStep)
        : (lowBound = midStep + 1);
    }
    return originSerial + (isForward ? lowBound : -lowBound);
  }
  return throwIssue();
}
const financialSolverLimits = Object.freeze({
  cashFlows: 1e4,
  periods: 1e4,
  couponPeriods: 1e3,
  evaluations: 2048,
  minRate: -0.999999999,
  maxRate: 1e6,
  relativeResidual: 1e-11,
});

export {
  fingerprintExtensionCall,
  parseFormulaTree,
  validateLetLambdaCall,
  parseTransformationRecipe,
  dateCalculationLimits,
  copyExtensionCacheEntries,
  validateSavedRecipeMetadata,
  walkFormulaNodes,
  calculationBudgetDefaults,
  resolveCalculationBudget,
  isoDateTimeSchemaFactory,
  resolveExtensionFunctionCall,
  deepFreezeValue,
  extensionRuntimeLimits,
  validateBindingName,
  isWithinNestingLimit,
  normalizeFunctionName,
  buildCalculationClock,
  tokenizeFormulaText,
  estimateWorkbookMemoryBytes,
  findVolatileFunctionNames,
  validateExtensionManifest,
  createCalculationBudgetTracker,
  defaultDatePolicy,
  dropBlankTextTokens,
  dateToSerial,
  dateFunctionNames,
  evaluateDateFunction,
  setExtensionCacheEntries,
  transformationRecipeLimits,
  supportedFunctionNames,
  unwrapDefaultExport,
  hashExtensionManifest,
  workbookUsesNowOrToday,
  isDeepFrozenValue,
  getExtensionCacheEntries,
  fingerprintCanonicalValue,
  notDateFunctionSentinel,
  sha256HexDigest,
  validateExtensionState,
  estimateValueFootprintBytes,
  measureSerializedBytes,
  normalizeBindingName,
  resolveLocalDateTime,
  validateExtensionValue,
  canonicalJsonStringify,
  workbookHasVolatileFunctions,
  daysInExcelMonth,
  financialSolverLimits,
  readMonotonicClockMs,
  recipeStoreFormatId,
  hasLinkedOrBoundCells,
  serialToDateParts,
  normalizeDatePolicy,
  dateSystemSerialOffset,
  dateSystemMaxSerial,
  calculationFeatureSupport,
  defineDataPropertyByKey,
  listInstalledExtensionFunctions,
  parseDateText,
};
export { copyModuleProperties };
export { boundedTextSchema };

/* Also used by the workers. */
export {
  SJ,
  UJ,
  aggregateMeasureSchema,
  cellScalarSchema,
  columnTypeProfileSchema,
  deepFrozenValueSet,
  distinctFieldListSchema,
  extensionEntityTypeSchema,
  extensionManifestSchema,
  extensionMatrixTypeSchema,
  extensionPermissionSchema,
  extensionScalarKindSchema,
  extensionScalarTypeSchema,
  extensionValueTypeSchema,
  filterConditionSchema,
  hasExtensionCacheForSheet,
  isExcelLeapYear,
  isValidFormulaName,
  isoTimestampSchema,
  linkedOrBoundCellCache,
  millisecondsPerDay,
  nD,
  nonnegativeSafeIntegerSchema,
  numericColumnSummarySchema,
  objectDefineProperty,
  od,
  opaqueIdSchema,
  safeIdentifierSchema,
  serialZeroUtcMs,
  sha256HexSchema,
  sheetRangeSourceSchema,
  singleLineTextSchema,
  transformationRecipeSchema,
  transformationRecipeVersionSchema,
  transformationRunEvidenceSchema,
  transformationStepRunSchema,
  transformationStepSchema,
  utf8Encoder,
};
