import i from "react/jsx-runtime";
import ne from "react";
import { sha256HexDigest } from "./01-schemas-date-functions-and-validation.js";
import {
  calculateWorkbookWithSettings,
  decodeAutomationArtifact,
  parseRangeAddress,
  automationLimits,
  applyCellPatchTransaction,
  normalizeWorkbookInput,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { importAutomationFile } from "./07-file-import-and-worker-dispatch.js";
import {
  retainSourceArtifact,
  verifyAutomationInspection,
} from "./08-xlsx-import-export-and-preservation.js";
import { useOperationScope } from "./09-workbook-links.js";
import { validateAssistantFormula } from "./20-calculation-options-and-assistant.js";
import { applyWorkbookChange } from "./24-merge-audit-and-ribbon-helpers.js";
import {
  splitVbaStatements,
  vbaStringLiteralPattern,
  tokenizeVbaSource,
  VbaAssignmentParser,
  createVbaSyntaxError,
  parseVbaAssignmentValue,
} from "./40-what-if-and-delimited-dialogs.js";

function parseVbaCommands(source) {
  const commands = [],
    issues = [],
    bindings = new Map(),
    declaredKinds = new Map();
  let hasOptionExplicit = false,
    isInSub = false,
    hasSubEnded = false,
    entryPoint = "";
  const resolveVbaObject = (reference) => {
    let objectMatch = new RegExp(
      `^(?:ThisWorkbook\\.)?(?:Worksheets|Sheets)\\((${vbaStringLiteralPattern})\\)$`,
      "i",
    ).exec(reference);
    if (objectMatch)
      return {
        kind: "sheet",
        sheet: String(parseVbaAssignmentValue(objectMatch[1])),
      };
    if (/^ActiveSheet$/i.test(reference)) return { kind: "sheet" };
    if (
      ((objectMatch = new RegExp(
        `^(?:(.+)\\.)?Range\\((${vbaStringLiteralPattern})\\)$`,
        "i",
      ).exec(reference)),
      objectMatch)
    ) {
      const sheetReference = objectMatch[1]
        ? resolveVbaObject(objectMatch[1])
        : { kind: "sheet" };
      if (sheetReference.kind !== "sheet")
        throw Error("Range requires a worksheet binding.");
      return {
        kind: "range",
        sheet: sheetReference.sheet,
        range: String(parseVbaAssignmentValue(objectMatch[2])),
      };
    }
    const namedBinding = bindings.get(reference.toLowerCase());
    if (namedBinding) return namedBinding;
    throw Error(
      "Unsupported VBA object or dynamic reference. Use a literal worksheet/range in the current workbook.",
    );
  };
  for (const [lineIndex, lineText] of source.split(/\r\n|\n|\r/).entries()) {
    const line = lineIndex + 1;
    try {
      if (/^\s*Rem(?:\s|$)/i.test(lineText)) continue;
      const statement = splitVbaStatements(lineText);
      if (!statement || /^Rem(?:\s|$)/i.test(statement)) continue;
      if (!isInSub && !hasSubEnded && /^Option Explicit$/i.test(statement)) {
        hasOptionExplicit = true;
        continue;
      }
      if (
        !isInSub &&
        !hasSubEnded &&
        (/^Option Explicit$/i.test(statement) ||
          /^Attribute VB_Name = "(?:[^"]|"")*"$/i.test(statement))
      )
        continue;
      let statementMatch =
        /^(?:(?:Public|Private)\s+)?Sub\s+([A-Za-z_][\w]*)\s*\(\s*\)$/i.exec(
          statement,
        );
      if (statementMatch) {
        if (isInSub || hasSubEnded)
          throw Error("Only one parameterless VBA Sub may be migrated.");
        ((isInSub = true), (entryPoint = statementMatch[1]));
        continue;
      }
      if (/^End Sub$/i.test(statement)) {
        if (!isInSub) throw Error("Unexpected End Sub.");
        ((isInSub = false), (hasSubEnded = true));
        continue;
      }
      if (!isInSub) throw Error("Code outside the single Sub is unsupported.");
      if (
        ((statementMatch =
          /^Dim\s+([A-Za-z_]\w*)\s+As\s+(Worksheet|Range)$/i.exec(statement)),
        statementMatch)
      ) {
        const declaredName = statementMatch[1].toLowerCase();
        if (declaredKinds.has(declaredName) || bindings.has(declaredName))
          throw Error("Duplicate or late VBA declaration.");
        declaredKinds.set(
          declaredName,
          statementMatch[2].toLowerCase() === "worksheet" ? "sheet" : "range",
        );
        continue;
      }
      if (
        ((statementMatch = /^Set\s+([A-Za-z_]\w*)\s*=\s*(.+)$/i.exec(
          statement,
        )),
        statementMatch)
      ) {
        const boundName = statementMatch[1].toLowerCase();
        if (bindings.has(boundName))
          throw Error("Binding reassignment is unsupported.");
        const boundObject = resolveVbaObject(statementMatch[2]);
        if (hasOptionExplicit && !declaredKinds.has(boundName))
          throw Error(
            "Option Explicit requires a declared worksheet/range binding.",
          );
        if (
          declaredKinds.has(boundName) &&
          declaredKinds.get(boundName) !== boundObject.kind
        )
          throw Error("VBA binding type does not match its declaration.");
        bindings.set(boundName, boundObject);
        continue;
      }
      if (
        ((statementMatch = /^(.+)\.(Value2?|Formula)\s*=\s*(.+)$/i.exec(
          statement,
        )),
        statementMatch)
      ) {
        const assignTarget = resolveVbaObject(statementMatch[1]);
        if (assignTarget.kind !== "range")
          throw Error("Assignment requires a literal range.");
        commands.push({
          line: line,
          kind:
            statementMatch[2].toLowerCase() === "formula"
              ? "set-formulas"
              : "set-values",
          sheet: assignTarget.sheet,
          range: assignTarget.range,
          values: parseVbaAssignmentValue(statementMatch[3]),
        });
        continue;
      }
      if (
        ((statementMatch = /^(.+)\.ClearContents$/i.exec(statement)),
        statementMatch)
      ) {
        const clearTarget = resolveVbaObject(statementMatch[1]);
        if (clearTarget.kind !== "range")
          throw Error("ClearContents requires a range.");
        commands.push({
          line: line,
          kind: "clear-contents",
          sheet: clearTarget.sheet,
          range: clearTarget.range,
        });
        continue;
      }
      throw Error(
        "Unsupported VBA statement/API. Loops, expressions, calls, events, files, COM and external operations are not migrated.",
      );
    } catch (error) {
      issues.push({
        line: line,
        message:
          error instanceof Error ? error.message : "Unsupported VBA source.",
      });
    }
  }
  return (
    (isInSub || !hasSubEnded) &&
      issues.push({
        line: 1,
        message: "A complete single parameterless Sub / End Sub is required.",
      }),
    { commands: commands, issues: issues, entryPoint: entryPoint }
  );
}
function migrateAutomationSource(source, language) {
  if (
    typeof source != "string" ||
    new TextEncoder().encode(source).length > automationLimits.maxSourceBytes ||
    source.includes("\0")
  )
    throw Error(
      "Migration source must be UTF-8 text of at most 256 KiB without nulls.",
    );
  if (!["vba", "office-script"].includes(language))
    throw Error("Unsupported automation language.");
  let commands = [],
    issues = [],
    entryPoint = language === "office-script" ? "main" : "";
  try {
    if (language === "vba")
      ({
        commands: commands,
        issues: issues,
        entryPoint: entryPoint,
      } = parseVbaCommands(source));
    else {
      const parser = new VbaAssignmentParser(tokenizeVbaSource(source));
      (parser.parse(), (commands = parser.commands), (issues = parser.issues));
    }
  } catch (error) {
    issues.push(createVbaSyntaxError(error));
  }
  return (
    commands.length > automationLimits.maxCommands &&
      issues.push({
        line: 1,
        message: "Migration is limited to 100 commands.",
      }),
    commands.length ||
      issues.push({
        line: 1,
        message: "No supported range commands were found.",
      }),
    {
      adapter: "literal-ranges-v1",
      language: language,
      codeSha256: sha256HexDigest(new TextEncoder().encode(source)),
      supported: !issues.length,
      entryPoint: entryPoint,
      commands: commands,
      issues: issues,
      warnings: [
        "This is a strict migration adapter, not a VBA/JavaScript runtime. Unsupported statements block the entire script; no partial execution occurs.",
        "Only literal range values, explicit formulas and ClearContents are supported. One approved worksheet/range is the read/write scope, up to 100 commands and 500 touched cells. Source events never become automatic triggers.",
        "Value writes that would be interpreted as formulas in Excel are refused. Use an explicit formula command and review its selected-range dependencies.",
      ],
    }
  );
}
const compiledMigrationCache = new WeakMap(),
  validateAutomationScope = (scope) => {
    const rangeAddress = parseRangeAddress(scope.range);
    if (
      (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
        (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      automationLimits.maxCells
    )
      throw Error("Approve one rectangular scope of at most 500 cells.");
    return {
      sheetId: scope.sheetId,
      range: `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`,
    };
  };
function assertAutomationSheetAllowed(workbookInput, scope, readOnly = false) {
  if (readOnly) throw Error("This workbook is read-only.");
  const sheet = workbookInput.sheets.find(
    (sheet) => sheet.id === scope.sheetId,
  );
  if (!sheet) throw Error("The approved worksheet does not exist.");
  if (sheet.hidden || sheet.veryHidden)
    throw Error("Hidden worksheets are outside the automation scope.");
  if (sheet.protection?.enabled)
    throw Error("The approved worksheet is protected.");
  return sheet;
}
const toApprovedRectangle = (scope) => {
  const rangeAddress = parseRangeAddress(scope.range),
    approvedAddresses = new Set();
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
      approvedAddresses.add(`${columnIndexToLetters(firstCol)}${firstRow + 1}`);
  return approvedAddresses;
};
function buildAutomationChanges(
  workbook,
  inspection,
  scopeRequest,
  options = {},
) {
  if ((verifyAutomationInspection(inspection), !inspection.script))
    throw Error(
      "This artifact supports inventory only. Export/import UTF-8 VBA or Office Scripts source for migration.",
    );
  const workbookInput = normalizeWorkbookInput(workbook),
    scope = validateAutomationScope(scopeRequest),
    sheet = assertAutomationSheetAllowed(
      workbookInput,
      scope,
      options.readOnly,
    ),
    approvedRectangle = toApprovedRectangle(scope),
    compatibility = migrateAutomationSource(
      inspection.script.text,
      inspection.script.language,
    ),
    plan = {
      schemaVersion: 1,
      workbookId: workbookInput.id,
      revision: workbookInput.revision,
      sourceName: inspection.fileName,
      sourceSha256: inspection.sha256,
      compatibility: compatibility,
      scope: scope,
      valid: false,
      issues: compatibility.issues.map(
        (issue) => `Line ${issue.line}: ${issue.message}`,
      ),
      touchedCells: 0,
      changes: [],
      capabilities: [
        "Write only the approved worksheet rectangle; existing dependent formulas recalculate normally",
        "Set literal values",
        "Set explicit scalar formulas with direct references inside the approved rectangle",
        "Clear cell contents; retain formatting",
        "No network, files, external workbook, COM, events or arbitrary code",
      ],
    },
    pendingPatches = new Map();
  if (compatibility.supported)
    for (const command of compatibility.commands)
      try {
        if (
          command.sheet !== undefined &&
          command.sheet.toLowerCase() !== sheet.name.toLowerCase()
        )
          throw Error(
            `Worksheet ${command.sheet} is outside the approved sheet ${sheet.name}.`,
          );
        const rangeAddress = parseRangeAddress(command.range),
          rowCount = rangeAddress.lastRow - rangeAddress.firstRow + 1,
          columnCount = rangeAddress.lastCol - rangeAddress.firstCol + 1,
          cellCount = rowCount * columnCount;
        if (
          ((plan.touchedCells += cellCount),
          plan.touchedCells > automationLimits.maxCells)
        )
          throw Error(
            "Migration exceeds 500 touched cells across all commands.",
          );
        const valueMatrix = Array.isArray(command.values)
          ? command.values
          : undefined;
        if (
          valueMatrix &&
          (valueMatrix.length !== rowCount ||
            valueMatrix.some(
              (matrixRow) =>
                !Array.isArray(matrixRow) || matrixRow.length !== columnCount,
            ))
        )
          throw Error(
            "The literal matrix dimensions do not match the target range.",
          );
        if (command.kind === "set-formulas" && !valueMatrix && cellCount !== 1)
          throw Error(
            "Scalar formula migration requires one cell; use explicit setFormulas matrices for multiple cells.",
          );
        for (
          let firstRow = rangeAddress.firstRow;
          firstRow <= rangeAddress.lastRow;
          firstRow++
        )
          for (
            let firstCol = rangeAddress.firstCol;
            firstCol <= rangeAddress.lastCol;
            firstCol++
          ) {
            const address = `${columnIndexToLetters(firstCol)}${firstRow + 1}`;
            if (!approvedRectangle.has(address))
              throw Error(
                `Target ${address} is outside the approved range ${scope.range}.`,
              );
            const raw =
              command.kind === "clear-contents"
                ? null
                : valueMatrix
                  ? valueMatrix[firstRow - rangeAddress.firstRow][
                      firstCol - rangeAddress.firstCol
                    ]
                  : command.values;
            if (command.kind === "set-values" && raw === null)
              throw Error(
                "Null setValue semantics are not migrated. Use an explicit ClearContents command.",
              );
            if (
              command.kind === "set-values" &&
              typeof raw == "string" &&
              /^[\s\u0000-\u001f]*[=+\-@＝＋－＠]/u.test(raw)
            )
              throw Error(
                "Formula-like value assignment has different Excel semantics. Use an explicit setFormula / .Formula statement or change the reviewed source.",
              );
            const existingCell =
                pendingPatches.get(address)?.cell ?? sheet.cells[address],
              cell = { ...existingCell, raw: raw };
            if ((delete cell.literal, command.kind === "set-formulas")) {
              if (typeof raw != "string" || !raw.startsWith("="))
                throw Error(
                  "Formula commands require explicit A1 formulas beginning with =.",
                );
              validateAssistantFormula(raw, approvedRectangle);
            }
            pendingPatches.set(address, {
              sheetId: sheet.id,
              address: address,
              cell: cell,
            });
          }
      } catch (error) {
        plan.issues.push(
          `Line ${command.line}: ${error instanceof Error ? error.message : "Migration validation failed."}`,
        );
      }
  const patches = [...pendingPatches.values()];
  if (!plan.issues.length)
    try {
      const workbook = applyCellPatchTransaction(
          workbookInput,
          patches,
          "Migration preview",
          "Preview migrated automation",
        ),
        eventChanges = workbook.events.at(-1)?.changes ?? [],
        workbookWithSettings = calculateWorkbookWithSettings(workbook);
      for (const patch of patches) {
        if (
          typeof patch.cell.raw != "string" ||
          !patch.cell.raw.startsWith("=")
        )
          continue;
        if (
          (workbookWithSettings(sheet.id, patch.address),
          workbookWithSettings.getOperationalErrors().length ||
            workbookWithSettings.getReferenceDiagnostics().length)
        )
          throw Error(
            "A migrated formula has invalid syntax, a circular reference, an unsupported reference or a calculation-limit failure.",
          );
        const spillRange = workbookWithSettings.getSpillRange(
          sheet.id,
          patch.address,
        );
        if (
          workbookWithSettings.getSpillDiagnostic(sheet.id, patch.address) ||
          (spillRange && spillRange !== `${patch.address}:${patch.address}`)
        )
          throw Error(
            "Array-spilling formula migration is unsupported. Use explicit scalar formulas so every write is reviewable.",
          );
      }
      if (
        eventChanges.some(
          (change) =>
            change.sheetId !== scope.sheetId ||
            !approvedRectangle.has(change.address),
        )
      )
        throw Error(
          "A dependent table/header rewrite would leave the approved scope. Migrate a separate range instead.",
        );
      ((plan.changes = eventChanges.map((eventChange) => ({
        sheetId: eventChange.sheetId,
        address: eventChange.address,
        before: eventChange.before ?? null,
        after: eventChange.after,
      }))),
        (plan.valid = true));
    } catch (error) {
      plan.issues.push(
        error instanceof Error ? error.message : "Workbook validation failed.",
      );
    }
  return (
    compiledMigrationCache.set(plan, {
      signature: JSON.stringify(plan),
      snapshot: JSON.stringify(workbook),
      inspection: inspection,
      patches: patches,
    }),
    plan
  );
}
function findAutomationArtifact(book, artifact) {
  return book.automationArtifacts?.some(
    (automationArtifact) =>
      automationArtifact.sha256 === artifact.sha256 &&
      automationArtifact.fileName === artifact.fileName,
  )
    ? book
    : normalizeWorkbookInput({
        ...book,
        automationArtifacts: [...(book.automationArtifacts ?? []), artifact],
      });
}
function preserveAutomationArtifact(book, inspection, options) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  const automationArtifact = findAutomationArtifact(
    book,
    retainSourceArtifact(inspection, options.actor),
  );
  return automationArtifact === book
    ? book
    : applyWorkbookChange(
        book,
        automationArtifact,
        options.actor,
        `Preserve automation source: ${inspection.fileName.slice(0, 100)}`,
      );
}
function removeAutomationArtifact(book, artifactId, options) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  if (
    !book.automationArtifacts?.some(
      (automationArtifact) => automationArtifact.id === artifactId,
    )
  )
    throw Error("Automation artifact was not found.");
  return applyWorkbookChange(
    book,
    {
      ...book,
      automationArtifacts: book.automationArtifacts.filter(
        (automationArtifact) => automationArtifact.id !== artifactId,
      ),
    },
    options.actor,
    "Remove preserved automation artifact",
  );
}
function runApprovedMigration(book, plan, options) {
  const compiledEntry = compiledMigrationCache.get(plan);
  if (
    !compiledEntry ||
    compiledEntry.signature !== JSON.stringify(plan) ||
    !plan.valid
  )
    throw Error(
      "The migration plan is invalid, modified or untrusted. Create and review a fresh plan.",
    );
  if (JSON.stringify(book) !== compiledEntry.snapshot)
    throw Error("The workbook changed. Review a fresh migration plan.");
  if (
    JSON.stringify(validateAutomationScope(options.scope)) !==
    JSON.stringify(plan.scope)
  )
    throw Error("The approved automation scope changed. Review again.");
  if (
    (assertAutomationSheetAllowed(book, plan.scope, options.readOnly),
    verifyAutomationInspection(compiledEntry.inspection),
    (book.automationRuns?.length ?? 0) >= automationLimits.maxRuns)
  )
    throw Error(
      "Automation run evidence is full (100 runs). Export native JSON before continuing.",
    );
  let updatedBook = applyCellPatchTransaction(
    book,
    compiledEntry.patches,
    options.actor,
    `Run migrated ${plan.compatibility.language}: ${plan.sourceName.slice(0, 90)} · ${plan.sourceSha256.slice(0, 12)}`,
  );
  options.retainSource !== false &&
    (updatedBook = findAutomationArtifact(
      updatedBook,
      retainSourceArtifact(compiledEntry.inspection, options.actor),
    ));
  const runEvent = updatedBook.events.at(-1),
    automationRun = {
      id: crypto.randomUUID(),
      eventId: runEvent.id,
      sourceName: plan.sourceName,
      sourceSha256: plan.sourceSha256,
      codeSha256: plan.compatibility.codeSha256,
      language: plan.compatibility.language,
      adapter: "literal-ranges-v1",
      actor: runEvent.actor,
      timestamp: runEvent.timestamp,
      beforeRevision: book.revision,
      afterRevision: updatedBook.revision,
      sheetId: plan.scope.sheetId,
      range: plan.scope.range,
      commands: plan.compatibility.commands.length,
      changedCells: plan.changes.length,
      sourceLines: plan.compatibility.commands.map((command) => command.line),
    };
  return (
    (updatedBook = normalizeWorkbookInput({
      ...updatedBook,
      automationRuns: [...(book.automationRuns ?? []), automationRun],
    })),
    compiledMigrationCache.delete(plan),
    updatedBook
  );
}
function AutomationScriptsDialog(props) {
  const { text, language } = useWorkbookMessages(),
    operationScope = useOperationScope(),
    abortControllerRef = ne.useRef(null),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    fileInputRef = ne.useRef(null),
    requestIdRef = ne.useRef(0),
    latestPropsRef = ne.useRef(props);
  latestPropsRef.current = props;
  const [inspection, setInspection] = ne.useState(
      props.initialInspection ?? null,
    ),
    [plan, setPlan] = ne.useState(null),
    [sourceText, setSourceText] =
      ne.useState(`function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange("B2").setValue(125);
  sheet.getRange("B3").setFormula("=B2*2");
}`),
    [sourceLanguage, setSourceLanguage] = ne.useState("office-script"),
    [sourceFileName, setSourceFileName] = ne.useState("Reviewed-script.ts"),
    [sheetId, setSheetId] = ne.useState(props.sheetId),
    [range, setRange] = ne.useState(props.range),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [disabled, setDisabled] = ne.useState(false),
    [checked, setChecked] = ne.useState(false),
    [shouldRetainSource, setShouldRetainSource] = ne.useState(true),
    [selectedArtifactId, setSelectedArtifactId] = ne.useState(""),
    compatibility = ne.useMemo(
      () =>
        inspection?.script
          ? migrateAutomationSource(
              inspection.script.text,
              inspection.script.language,
            )
          : null,
      [inspection],
    ),
    scopeRequest = { sheetId: sheetId, range: range },
    resetReviewState = () => {
      (setPlan(null), setChecked(false), setErrorMessage(""));
    },
    cancelInspection = () => {
      (abortControllerRef.current?.abort(),
        requestIdRef.current++,
        setDisabled(false));
    };
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.focus(),
      () => {
        (abortControllerRef.current?.abort(),
          requestIdRef.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      resetReviewState();
    }, [props.workbook, props.readOnly]));
  const runReviewAction = (action) => {
    try {
      (setErrorMessage(""), action());
    } catch (error) {
      setErrorMessage(
        error instanceof Error ? error.message : "Automation review failed.",
      );
    }
  };
  async function inspectFile(file) {
    if (!file) return;
    abortControllerRef.current?.abort();
    const abortController = new AbortController();
    abortControllerRef.current = abortController;
    const requestId = ++requestIdRef.current;
    (resetReviewState(), setDisabled(true), setErrorMessage(""));
    try {
      if (file.size > 10 * 1024 * 1024)
        throw Error("Automation files must be 10 MB or smaller.");
      const inspectionResult = await importAutomationFile(file, {
        ...operationScope,
        signal: operationScope.signal
          ? AbortSignal.any([abortController.signal, operationScope.signal])
          : abortController.signal,
      });
      if (requestId !== requestIdRef.current) return;
      (setInspection(inspectionResult),
        setSelectedArtifactId(""),
        setPlan(null),
        setChecked(false));
    } catch (error) {
      if (requestId !== requestIdRef.current) return;
      setErrorMessage(
        error instanceof Error
          ? error.message
          : "Unable to inspect automation.",
      );
    } finally {
      requestId === requestIdRef.current &&
        (setDisabled(false),
        fileInputRef.current && (fileInputRef.current.value = ""));
    }
  }
  const buildReportBytes = () =>
    new TextEncoder().encode(
      JSON.stringify(
        { inspection: inspection, compatibility: compatibility, plan: plan },
        null,
        2,
      ),
    );
  return i.jsx("div", {
    className: "lg-automation-overlay",
    children: i.jsxs("section", {
      className: "lg-automation",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      tabIndex: -1,
      ref: dialogRef,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
              ...(dialogRef.current?.querySelectorAll(
                "button:not(:disabled),input:not(:disabled):not([type=file]),select:not(:disabled),textarea:not(:disabled),summary",
              ) ?? []),
            ],
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1);
          event.shiftKey &&
          (document.activeElement === firstFocusable ||
            document.activeElement === dialogRef.current)
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === lastFocusable &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: titleId,
              children: i.jsx(LocalizedText, {
                children: "Automation & Scripts",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Inspect, preserve and migrate approved workbook operations.",
              }),
            }),
            i.jsx("button", {
              "aria-label": text("Close automation review"),
              onClick: props.onClose,
              lang: language("Close automation review"),
              "data-ui-label": "Close automation review",
              children: "×",
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-automation-scroll",
          children: [
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Imported VBA, XLM and Office Scripts never run automatically. This adapter accepts a strict subset of literal range operations; unsupported APIs block the whole migration.",
              }),
            }),
            i.jsxs("div", {
              className: "lg-automation-actions",
              children: [
                i.jsx("button", {
                  disabled: disabled,
                  onClick: () => fileInputRef.current?.click(),
                  children: i.jsx(LocalizedText, {
                    children: "Inspect automation file",
                  }),
                }),
                i.jsx("input", {
                  ref: fileInputRef,
                  type: "file",
                  hidden: true,
                  "aria-label": text("Automation source file"),
                  accept:
                    ".bas,.vba,.cls,.ts,.osts,.xlsm,.xlam,.xltm,.xlsx,.xltx,.xls,.xlsb",
                  onChange: (event) =>
                    void inspectFile(event.target.files?.[0]),
                  lang: language("Automation source file"),
                  "data-ui-label": "Automation source file",
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Preserved source" }),
                    i.jsxs("select", {
                      value: selectedArtifactId,
                      onChange: (event) =>
                        runReviewAction(() => {
                          (cancelInspection(),
                            setSelectedArtifactId(event.target.value));
                          const selectedArtifact =
                            props.workbook.automationArtifacts?.find(
                              (automationArtifact) =>
                                automationArtifact.id === event.target.value,
                            );
                          selectedArtifact &&
                            (inspectFile(
                              new File(
                                [
                                  decodeAutomationArtifact(
                                    selectedArtifact.base64,
                                  ),
                                ],
                                selectedArtifact.fileName,
                              ),
                            ),
                            resetReviewState());
                        }),
                      children: [
                        i.jsx("option", {
                          value: "",
                          children: "Select an artifact",
                        }),
                        props.workbook.automationArtifacts?.map(
                          (automationArtifact) =>
                            i.jsxs(
                              "option",
                              {
                                value: automationArtifact.id,
                                children: [
                                  automationArtifact.fileName,
                                  " · ",
                                  automationArtifact.sha256.slice(0, 12),
                                ],
                              },
                              automationArtifact.id,
                            ),
                        ),
                      ],
                    }),
                  ],
                }),
                i.jsx("button", {
                  disabled: !selectedArtifactId || props.readOnly,
                  onClick: () =>
                    runReviewAction(() => {
                      (props.onRemove(selectedArtifactId),
                        setSelectedArtifactId(""),
                        resetReviewState());
                    }),
                  children: i.jsx(LocalizedText, {
                    children: "Remove preserved source",
                  }),
                }),
              ],
            }),
            i.jsxs("details", {
              children: [
                i.jsx("summary", {
                  children: i.jsx(LocalizedText, {
                    children: "Paste source for compatibility review",
                  }),
                }),
                i.jsxs("div", {
                  className: "lg-automation-fields",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Source language" }),
                        i.jsxs("select", {
                          value: sourceLanguage,
                          onChange: (event) => {
                            (cancelInspection(),
                              setSourceLanguage(event.target.value),
                              setSourceFileName(
                                event.target.value === "vba"
                                  ? "Reviewed-module.bas"
                                  : "Reviewed-script.ts",
                              ),
                              resetReviewState(),
                              setInspection(null));
                          },
                          children: [
                            i.jsx("option", {
                              value: "office-script",
                              children: "Office Scripts",
                            }),
                            i.jsx("option", {
                              value: "vba",
                              children: "VBA module",
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Source file name" }),
                        i.jsx("input", {
                          value: sourceFileName,
                          onChange: (event) => {
                            (cancelInspection(),
                              setSourceFileName(event.target.value),
                              resetReviewState(),
                              setInspection(null));
                          },
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Source code" }),
                    i.jsx("textarea", {
                      spellCheck: false,
                      rows: 8,
                      value: sourceText,
                      onChange: (event) => {
                        (cancelInspection(),
                          setSourceText(event.target.value),
                          resetReviewState(),
                          setInspection(null));
                      },
                    }),
                  ],
                }),
                i.jsx("button", {
                  onClick: () =>
                    runReviewAction(() => {
                      if (
                        (cancelInspection(),
                        (sourceLanguage === "vba" &&
                          !/\.(bas|vba|cls)$/i.test(sourceFileName)) ||
                          (sourceLanguage === "office-script" &&
                            !/\.ts$/i.test(sourceFileName)))
                      )
                        throw Error(
                          "Match the source file extension to the chosen language.",
                        );
                      (inspectFile(new File([sourceText], sourceFileName)),
                        resetReviewState(),
                        setSelectedArtifactId(""));
                    }),
                  children: i.jsx(LocalizedText, {
                    children: "Inspect pasted source",
                  }),
                }),
              ],
            }),
            disabled &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children: "Inspecting source bytes…",
                }),
              }),
            inspection &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Automation inventory",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx("strong", { children: inspection.fileName }),
                      " ·",
                      " ",
                      (inspection.fileBytes / 1024).toFixed(1),
                      " ",
                      i.jsx(LocalizedText, { children: "KiB ·" }),
                      " ",
                      inspection.execution === "inventory-only"
                        ? "Inventory only — execution unsupported"
                        : "Migration candidate — compatibility and approval required",
                    ],
                  }),
                  i.jsxs("p", {
                    className: "lg-automation-hash",
                    children: [
                      i.jsx(LocalizedText, { children: "SHA-256: " }),
                      inspection.sha256,
                    ],
                  }),
                  i.jsxs("table", {
                    "aria-label": text("Automation inventory"),
                    lang: language("Automation inventory"),
                    "data-ui-label": "Automation inventory",
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Artifact",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Kind",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Support",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Notes",
                              }),
                            }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: inspection.items.map((item, itemIndex) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: item.path }),
                                i.jsx("td", { children: item.kind }),
                                i.jsx("td", { children: item.execution }),
                                i.jsx("td", { children: item.note }),
                              ],
                            },
                            itemIndex,
                          ),
                        ),
                      }),
                    ],
                  }),
                  !inspection.items.length &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "No known package automation parts detected. Separate or external scripts remain unknown.",
                      }),
                    }),
                  i.jsx("ul", {
                    children: inspection.warnings.map((warning, warningIndex) =>
                      i.jsx("li", { children: warning }, warningIndex),
                    ),
                  }),
                  i.jsxs("div", {
                    className: "lg-automation-actions",
                    children: [
                      i.jsx("button", {
                        disabled:
                          props.readOnly || !inspection.retainedEligible,
                        onClick: () =>
                          runReviewAction(() => {
                            (props.onPreserve(inspection),
                              setPlan(null),
                              setChecked(false));
                          }),
                        children: i.jsx(LocalizedText, {
                          children: "Preserve original source",
                        }),
                      }),
                      i.jsx("button", {
                        onClick: () =>
                          runReviewAction(() => {
                            const preservedArtifact =
                              latestPropsRef.current.workbook.automationArtifacts?.find(
                                (automationArtifact) =>
                                  automationArtifact.sha256 ===
                                    inspection.sha256 &&
                                  automationArtifact.fileName ===
                                    inspection.fileName,
                              );
                            if (!preservedArtifact)
                              throw Error(
                                "Preserve the original source before downloading it from the workbook.",
                              );
                            props.onDownload(
                              decodeAutomationArtifact(
                                preservedArtifact.base64,
                              ),
                              preservedArtifact.fileName,
                            );
                          }),
                        children: i.jsx(LocalizedText, {
                          children: "Download preserved original",
                        }),
                      }),
                    ],
                  }),
                  compatibility &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("h3", {
                          children: i.jsx(LocalizedText, {
                            children: "Migration compatibility",
                          }),
                        }),
                        i.jsxs("p", {
                          children: [
                            compatibility.supported
                              ? "Supported source grammar; workbook scope checks still required."
                              : "Unsupported source — migration blocked.",
                            " ",
                            i.jsx(LocalizedText, { children: "Entry point:" }),
                            " ",
                            compatibility.entryPoint || "(not identified)",
                            ".",
                            " ",
                            compatibility.commands.length,
                            " ",
                            i.jsx(LocalizedText, {
                              children: "recognized commands.",
                            }),
                          ],
                        }),
                        i.jsx("ul", {
                          "aria-label": text("Source compatibility issues"),
                          lang: language("Source compatibility issues"),
                          "data-ui-label": "Source compatibility issues",
                          children: compatibility.issues.map(
                            (issue, issueIndex) =>
                              i.jsxs(
                                "li",
                                {
                                  children: [
                                    i.jsx(LocalizedText, { children: "Line " }),
                                    issue.line,
                                    ": ",
                                    issue.message,
                                  ],
                                },
                                issueIndex,
                              ),
                          ),
                        }),
                        i.jsx("ul", {
                          children: compatibility.warnings.map(
                            (warning, compatibilityWarningIndex) =>
                              i.jsx(
                                "li",
                                { children: warning },
                                compatibilityWarningIndex,
                              ),
                          ),
                        }),
                        i.jsxs("details", {
                          children: [
                            i.jsx("summary", {
                              children: i.jsx(LocalizedText, {
                                children: "Reviewed source text",
                              }),
                            }),
                            i.jsx("pre", {
                              children: inspection.script?.text,
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "lg-automation-fields",
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Approved worksheet",
                                }),
                                i.jsx("select", {
                                  value: sheetId,
                                  onChange: (event) => {
                                    (setSheetId(event.target.value),
                                      resetReviewState());
                                  },
                                  children: props.workbook.sheets
                                    .filter(
                                      (sheet) =>
                                        !sheet.hidden && !sheet.veryHidden,
                                    )
                                    .map((worksheet) =>
                                      i.jsx(
                                        "option",
                                        {
                                          value: worksheet.id,
                                          children: worksheet.name,
                                        },
                                        worksheet.id,
                                      ),
                                    ),
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children:
                                    "Approved range (maximum 500 cells)",
                                }),
                                i.jsx("input", {
                                  value: range,
                                  onChange: (event) => {
                                    (setRange(event.target.value),
                                      resetReviewState());
                                  },
                                }),
                              ],
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          disabled: props.readOnly || !compatibility.supported,
                          onClick: () =>
                            runReviewAction(() => {
                              (setChecked(false),
                                setPlan(
                                  buildAutomationChanges(
                                    props.workbook,
                                    inspection,
                                    scopeRequest,
                                    {
                                      readOnly: props.readOnly,
                                    },
                                  ),
                                ));
                            }),
                          children: i.jsx(LocalizedText, {
                            children: "Preview migrated commands",
                          }),
                        }),
                      ],
                    }),
                ],
              }),
            plan &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Command and change review",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      plan.compatibility.commands.length,
                      " ",
                      i.jsx(LocalizedText, { children: "commands ·" }),
                      " ",
                      plan.touchedCells,
                      " ",
                      i.jsx(LocalizedText, { children: "touched cells · " }),
                      plan.changes.length,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "final cell writes · revision ",
                      }),
                      plan.revision,
                      ".",
                    ],
                  }),
                  i.jsx("ul", {
                    children: plan.capabilities.map((capability) =>
                      i.jsx("li", { children: capability }, capability),
                    ),
                  }),
                  !!plan.issues.length &&
                    i.jsx("div", {
                      role: "alert",
                      children: i.jsx("ul", {
                        children: plan.issues.map((issue, planIssueIndex) =>
                          i.jsx("li", { children: issue }, planIssueIndex),
                        ),
                      }),
                    }),
                  i.jsxs("table", {
                    "aria-label": text("Migrated workbook commands"),
                    lang: language("Migrated workbook commands"),
                    "data-ui-label": "Migrated workbook commands",
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Source line",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Command",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Target",
                              }),
                            }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: plan.compatibility.commands.map(
                          (command, commandIndex) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsx("td", { children: command.line }),
                                  i.jsx("td", { children: command.kind }),
                                  i.jsxs("td", {
                                    children: [
                                      command.sheet ?? "(approved worksheet)",
                                      "!",
                                      command.range,
                                    ],
                                  }),
                                ],
                              },
                              commandIndex,
                            ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("table", {
                    "aria-label": text("Automation change preview"),
                    lang: language("Automation change preview"),
                    "data-ui-label": "Automation change preview",
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Cell",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Before",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "After",
                              }),
                            }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: plan.changes.map((change) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: change.address }),
                                i.jsx("td", {
                                  children: String(
                                    change.before?.raw ?? "(blank)",
                                  ),
                                }),
                                i.jsx("td", {
                                  children: String(
                                    change.after.raw ?? "(blank)",
                                  ),
                                }),
                              ],
                            },
                            change.address,
                          ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    className: "lg-automation-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: shouldRetainSource,
                        onChange: (event) => {
                          (setShouldRetainSource(event.target.checked),
                            setChecked(false));
                        },
                      }),
                      i.jsx(LocalizedText, {
                        children:
                          "Retain original source with the run (native JSON)",
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    className: "lg-automation-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: checked,
                        disabled: !plan.valid || props.readOnly,
                        onChange: (event) => setChecked(event.target.checked),
                      }),
                      i.jsx(LocalizedText, {
                        children:
                          "I approve these commands and their changes within this scope",
                      }),
                    ],
                  }),
                ],
              }),
            !!props.workbook.automationRuns?.length &&
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Recorded automation runs (",
                      }),
                      props.workbook.automationRuns.length,
                      ")",
                    ],
                  }),
                  i.jsx("ul", {
                    children: props.workbook.automationRuns
                      .slice(-10)
                      .map((automationRun) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              automationRun.sourceName,
                              " · ",
                              automationRun.actor,
                              " · ",
                              automationRun.timestamp,
                              " ",
                              i.jsx(LocalizedText, { children: "· revision" }),
                              " ",
                              automationRun.beforeRevision,
                              " → ",
                              automationRun.afterRevision,
                              " · ",
                              automationRun.changedCells,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "writes · source ",
                              }),
                              automationRun.sourceSha256,
                            ],
                          },
                          automationRun.id,
                        ),
                      ),
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Native JSON contains the full retained run record. Local history is editable and is not an immutable audit service.",
                    }),
                  }),
                ],
              }),
          ],
        }),
        props.readOnly &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children:
                "Read-only: inspection and report download remain available; preservation and execution are disabled.",
            }),
          }),
        errorMessage && i.jsx("p", { role: "alert", children: errorMessage }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              disabled: !inspection,
              onClick: () =>
                props.onDownload(
                  buildReportBytes(),
                  `${inspection?.fileName ?? "automation"}.compatibility.json`,
                ),
              children: i.jsx(LocalizedText, {
                children: "Save automation report",
              }),
            }),
            i.jsx("button", {
              disabled: !checked || !plan?.valid || props.readOnly,
              onClick: () =>
                runReviewAction(() => {
                  !plan ||
                    !checked ||
                    (props.onApply(plan, scopeRequest, shouldRetainSource),
                    setPlan(null),
                    setChecked(false));
                }),
              children: i.jsx(LocalizedText, {
                children: "Run approved migration",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
const w2 = 25;

export {
  AutomationScriptsDialog,
  removeAutomationArtifact,
  runApprovedMigration,
  preserveAutomationArtifact,
  w2,
};
