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

function parseVbaCommands(e) {
  const commands = [],
    issues = [],
    n = new Map(),
    o = new Map();
  let a = false,
    s = false,
    d = false,
    entryPoint = "";
  const l = (c) => {
    let h = new RegExp(
      `^(?:ThisWorkbook\\.)?(?:Worksheets|Sheets)\\((${vbaStringLiteralPattern})\\)$`,
      "i",
    ).exec(c);
    if (h)
      return { kind: "sheet", sheet: String(parseVbaAssignmentValue(h[1])) };
    if (/^ActiveSheet$/i.test(c)) return { kind: "sheet" };
    if (
      ((h = new RegExp(
        `^(?:(.+)\\.)?Range\\((${vbaStringLiteralPattern})\\)$`,
        "i",
      ).exec(c)),
      h)
    ) {
      const p = h[1] ? l(h[1]) : { kind: "sheet" };
      if (p.kind !== "sheet")
        throw Error("Range requires a worksheet binding.");
      return {
        kind: "range",
        sheet: p.sheet,
        range: String(parseVbaAssignmentValue(h[2])),
      };
    }
    const f = n.get(c.toLowerCase());
    if (f) return f;
    throw Error(
      "Unsupported VBA object or dynamic reference. Use a literal worksheet/range in the current workbook.",
    );
  };
  for (const [c, h] of e.split(/\r\n|\n|\r/).entries()) {
    const line = c + 1;
    try {
      if (/^\s*Rem(?:\s|$)/i.test(h)) continue;
      const p = splitVbaStatements(h);
      if (!p || /^Rem(?:\s|$)/i.test(p)) continue;
      if (!s && !d && /^Option Explicit$/i.test(p)) {
        a = true;
        continue;
      }
      if (
        !s &&
        !d &&
        (/^Option Explicit$/i.test(p) ||
          /^Attribute VB_Name = "(?:[^"]|"")*"$/i.test(p))
      )
        continue;
      let m =
        /^(?:(?:Public|Private)\s+)?Sub\s+([A-Za-z_][\w]*)\s*\(\s*\)$/i.exec(p);
      if (m) {
        if (s || d)
          throw Error("Only one parameterless VBA Sub may be migrated.");
        ((s = true), (entryPoint = m[1]));
        continue;
      }
      if (/^End Sub$/i.test(p)) {
        if (!s) throw Error("Unexpected End Sub.");
        ((s = false), (d = true));
        continue;
      }
      if (!s) throw Error("Code outside the single Sub is unsupported.");
      if (
        ((m = /^Dim\s+([A-Za-z_]\w*)\s+As\s+(Worksheet|Range)$/i.exec(p)), m)
      ) {
        const v = m[1].toLowerCase();
        if (o.has(v) || n.has(v))
          throw Error("Duplicate or late VBA declaration.");
        o.set(v, m[2].toLowerCase() === "worksheet" ? "sheet" : "range");
        continue;
      }
      if (((m = /^Set\s+([A-Za-z_]\w*)\s*=\s*(.+)$/i.exec(p)), m)) {
        const v = m[1].toLowerCase();
        if (n.has(v)) throw Error("Binding reassignment is unsupported.");
        const g = l(m[2]);
        if (a && !o.has(v))
          throw Error(
            "Option Explicit requires a declared worksheet/range binding.",
          );
        if (o.has(v) && o.get(v) !== g.kind)
          throw Error("VBA binding type does not match its declaration.");
        n.set(v, g);
        continue;
      }
      if (((m = /^(.+)\.(Value2?|Formula)\s*=\s*(.+)$/i.exec(p)), m)) {
        const v = l(m[1]);
        if (v.kind !== "range")
          throw Error("Assignment requires a literal range.");
        commands.push({
          line: line,
          kind:
            m[2].toLowerCase() === "formula" ? "set-formulas" : "set-values",
          sheet: v.sheet,
          range: v.range,
          values: parseVbaAssignmentValue(m[3]),
        });
        continue;
      }
      if (((m = /^(.+)\.ClearContents$/i.exec(p)), m)) {
        const v = l(m[1]);
        if (v.kind !== "range") throw Error("ClearContents requires a range.");
        commands.push({
          line: line,
          kind: "clear-contents",
          sheet: v.sheet,
          range: v.range,
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
    (s || !d) &&
      issues.push({
        line: 1,
        message: "A complete single parameterless Sub / End Sub is required.",
      }),
    { commands: commands, issues: issues, entryPoint: entryPoint }
  );
}
function migrateAutomationSource(e, language) {
  if (
    typeof e != "string" ||
    new TextEncoder().encode(e).length > automationLimits.maxSourceBytes ||
    e.includes("\0")
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
      } = parseVbaCommands(e));
    else {
      const a = new VbaAssignmentParser(tokenizeVbaSource(e));
      (a.parse(), (commands = a.commands), (issues = a.issues));
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
      codeSha256: sha256HexDigest(new TextEncoder().encode(e)),
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
  validateAutomationScope = (e) => {
    const rangeAddress = parseRangeAddress(e.range);
    if (
      (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
        (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      automationLimits.maxCells
    )
      throw Error("Approve one rectangular scope of at most 500 cells.");
    return {
      sheetId: e.sheetId,
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
const toApprovedRectangle = (e) => {
  const rangeAddress = parseRangeAddress(e.range),
    t = new Set();
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
      t.add(`${columnIndexToLetters(firstCol)}${firstRow + 1}`);
  return t;
};
function buildAutomationChanges(workbook, inspection, t, n = {}) {
  if ((verifyAutomationInspection(inspection), !inspection.script))
    throw Error(
      "This artifact supports inventory only. Export/import UTF-8 VBA or Office Scripts source for migration.",
    );
  const workbookInput = normalizeWorkbookInput(workbook),
    scope = validateAutomationScope(t),
    s = assertAutomationSheetAllowed(workbookInput, scope, n.readOnly),
    approvedRectangle = toApprovedRectangle(scope),
    compatibility = migrateAutomationSource(
      inspection.script.text,
      inspection.script.language,
    ),
    l = {
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
    c = new Map();
  if (compatibility.supported)
    for (const command of compatibility.commands)
      try {
        if (
          command.sheet !== undefined &&
          command.sheet.toLowerCase() !== s.name.toLowerCase()
        )
          throw Error(
            `Worksheet ${command.sheet} is outside the approved sheet ${s.name}.`,
          );
        const rangeAddress = parseRangeAddress(command.range),
          g = rangeAddress.lastRow - rangeAddress.firstRow + 1,
          b = rangeAddress.lastCol - rangeAddress.firstCol + 1,
          y = g * b;
        if (((l.touchedCells += y), l.touchedCells > automationLimits.maxCells))
          throw Error(
            "Migration exceeds 500 touched cells across all commands.",
          );
        const w = Array.isArray(command.values) ? command.values : undefined;
        if (
          w &&
          (w.length !== g || w.some((E) => !Array.isArray(E) || E.length !== b))
        )
          throw Error(
            "The literal matrix dimensions do not match the target range.",
          );
        if (command.kind === "set-formulas" && !w && y !== 1)
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
                : w
                  ? w[firstRow - rangeAddress.firstRow][
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
            const I = c.get(address)?.cell ?? s.cells[address],
              cell = { ...I, raw: raw };
            if ((delete cell.literal, command.kind === "set-formulas")) {
              if (typeof raw != "string" || !raw.startsWith("="))
                throw Error(
                  "Formula commands require explicit A1 formulas beginning with =.",
                );
              validateAssistantFormula(raw, approvedRectangle);
            }
            c.set(address, { sheetId: s.id, address: address, cell: cell });
          }
      } catch (error) {
        l.issues.push(
          `Line ${command.line}: ${error instanceof Error ? error.message : "Migration validation failed."}`,
        );
      }
  const patches = [...c.values()];
  if (!l.issues.length)
    try {
      const workbook = applyCellPatchTransaction(
          workbookInput,
          patches,
          "Migration preview",
          "Preview migrated automation",
        ),
        v = workbook.events.at(-1)?.changes ?? [],
        workbookWithSettings = calculateWorkbookWithSettings(workbook);
      for (const patch of patches) {
        if (
          typeof patch.cell.raw != "string" ||
          !patch.cell.raw.startsWith("=")
        )
          continue;
        if (
          (workbookWithSettings(s.id, patch.address),
          workbookWithSettings.getOperationalErrors().length ||
            workbookWithSettings.getReferenceDiagnostics().length)
        )
          throw Error(
            "A migrated formula has invalid syntax, a circular reference, an unsupported reference or a calculation-limit failure.",
          );
        const y = workbookWithSettings.getSpillRange(s.id, patch.address);
        if (
          workbookWithSettings.getSpillDiagnostic(s.id, patch.address) ||
          (y && y !== `${patch.address}:${patch.address}`)
        )
          throw Error(
            "Array-spilling formula migration is unsupported. Use explicit scalar formulas so every write is reviewable.",
          );
      }
      if (
        v.some(
          (b) =>
            b.sheetId !== scope.sheetId || !approvedRectangle.has(b.address),
        )
      )
        throw Error(
          "A dependent table/header rewrite would leave the approved scope. Migrate a separate range instead.",
        );
      ((l.changes = v.map((b) => ({
        sheetId: b.sheetId,
        address: b.address,
        before: b.before ?? null,
        after: b.after,
      }))),
        (l.valid = true));
    } catch (error) {
      l.issues.push(
        error instanceof Error ? error.message : "Workbook validation failed.",
      );
    }
  return (
    compiledMigrationCache.set(l, {
      signature: JSON.stringify(l),
      snapshot: JSON.stringify(workbook),
      inspection: inspection,
      patches: patches,
    }),
    l
  );
}
function findAutomationArtifact(book, r) {
  return book.automationArtifacts?.some(
    (automationArtifact) =>
      automationArtifact.sha256 === r.sha256 &&
      automationArtifact.fileName === r.fileName,
  )
    ? book
    : normalizeWorkbookInput({
        ...book,
        automationArtifacts: [...(book.automationArtifacts ?? []), r],
      });
}
function preserveAutomationArtifact(book, inspection, t) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  const automationArtifact = findAutomationArtifact(
    book,
    retainSourceArtifact(inspection, t.actor),
  );
  return automationArtifact === book
    ? book
    : applyWorkbookChange(
        book,
        automationArtifact,
        t.actor,
        `Preserve automation source: ${inspection.fileName.slice(0, 100)}`,
      );
}
function removeAutomationArtifact(book, r, t) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (
    !book.automationArtifacts?.some(
      (automationArtifact) => automationArtifact.id === r,
    )
  )
    throw Error("Automation artifact was not found.");
  return applyWorkbookChange(
    book,
    {
      ...book,
      automationArtifacts: book.automationArtifacts.filter(
        (automationArtifact) => automationArtifact.id !== r,
      ),
    },
    t.actor,
    "Remove preserved automation artifact",
  );
}
function runApprovedMigration(book, r, t) {
  const n = compiledMigrationCache.get(r);
  if (!n || n.signature !== JSON.stringify(r) || !r.valid)
    throw Error(
      "The migration plan is invalid, modified or untrusted. Create and review a fresh plan.",
    );
  if (JSON.stringify(book) !== n.snapshot)
    throw Error("The workbook changed. Review a fresh migration plan.");
  if (
    JSON.stringify(validateAutomationScope(t.scope)) !== JSON.stringify(r.scope)
  )
    throw Error("The approved automation scope changed. Review again.");
  if (
    (assertAutomationSheetAllowed(book, r.scope, t.readOnly),
    verifyAutomationInspection(n.inspection),
    (book.automationRuns?.length ?? 0) >= automationLimits.maxRuns)
  )
    throw Error(
      "Automation run evidence is full (100 runs). Export native JSON before continuing.",
    );
  let o = applyCellPatchTransaction(
    book,
    n.patches,
    t.actor,
    `Run migrated ${r.compatibility.language}: ${r.sourceName.slice(0, 90)} · ${r.sourceSha256.slice(0, 12)}`,
  );
  t.retainSource !== false &&
    (o = findAutomationArtifact(
      o,
      retainSourceArtifact(n.inspection, t.actor),
    ));
  const a = o.events.at(-1),
    s = {
      id: crypto.randomUUID(),
      eventId: a.id,
      sourceName: r.sourceName,
      sourceSha256: r.sourceSha256,
      codeSha256: r.compatibility.codeSha256,
      language: r.compatibility.language,
      adapter: "literal-ranges-v1",
      actor: a.actor,
      timestamp: a.timestamp,
      beforeRevision: book.revision,
      afterRevision: o.revision,
      sheetId: r.scope.sheetId,
      range: r.scope.range,
      commands: r.compatibility.commands.length,
      changedCells: r.changes.length,
      sourceLines: r.compatibility.commands.map((command) => command.line),
    };
  return (
    (o = normalizeWorkbookInput({
      ...o,
      automationRuns: [...(book.automationRuns ?? []), s],
    })),
    compiledMigrationCache.delete(r),
    o
  );
}
function AutomationScriptsDialog(e) {
  const { text, language } = useWorkbookMessages(),
    operationScope = useOperationScope(),
    o = ne.useRef(null),
    a = ne.useId(),
    s = ne.useRef(null),
    d = ne.useRef(null),
    u = ne.useRef(0),
    l = ne.useRef(e);
  l.current = e;
  const [inspection, h] = ne.useState(e.initialInspection ?? null),
    [plan, p] = ne.useState(null),
    [m, v] = ne.useState(`function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();
  sheet.getRange("B2").setValue(125);
  sheet.getRange("B3").setFormula("=B2*2");
}`),
    [g, b] = ne.useState("office-script"),
    [y, w] = ne.useState("Reviewed-script.ts"),
    [sheetId, A] = ne.useState(e.sheetId),
    [range, F] = ne.useState(e.range),
    [I, O] = ne.useState(""),
    [disabled, z] = ne.useState(false),
    [checked, D] = ne.useState(false),
    [B, _] = ne.useState(true),
    [C, T] = ne.useState(""),
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
    R = { sheetId: sheetId, range: range },
    L = () => {
      (p(null), D(false), O(""));
    },
    $ = () => {
      (o.current?.abort(), u.current++, z(false));
    };
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      s.current?.focus(),
      () => {
        (o.current?.abort(),
          u.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      L();
    }, [e.workbook, e.readOnly]));
  const U = (q) => {
    try {
      (O(""), q());
    } catch (error) {
      O(error instanceof Error ? error.message : "Automation review failed.");
    }
  };
  async function G(q) {
    if (!q) return;
    o.current?.abort();
    const W = new AbortController();
    o.current = W;
    const K = ++u.current;
    (L(), z(true), O(""));
    try {
      if (q.size > 10 * 1024 * 1024)
        throw Error("Automation files must be 10 MB or smaller.");
      const ae = await importAutomationFile(q, {
        ...operationScope,
        signal: operationScope.signal
          ? AbortSignal.any([W.signal, operationScope.signal])
          : W.signal,
      });
      if (K !== u.current) return;
      (h(ae), T(""), p(null), D(false));
    } catch (error) {
      if (K !== u.current) return;
      O(
        error instanceof Error
          ? error.message
          : "Unable to inspect automation.",
      );
    } finally {
      K === u.current && (z(false), d.current && (d.current.value = ""));
    }
  }
  const H = () =>
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
      "aria-labelledby": a,
      tabIndex: -1,
      ref: s,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), e.onClose()),
          event.key === "Tab")
        ) {
          const K = [
              ...(s.current?.querySelectorAll(
                "button:not(:disabled),input:not(:disabled):not([type=file]),select:not(:disabled),textarea:not(:disabled),summary",
              ) ?? []),
            ],
            te = K[0],
            ae = K.at(-1);
          event.shiftKey &&
          (document.activeElement === te ||
            document.activeElement === s.current)
            ? (event.preventDefault(), ae == null || ae.focus())
            : !event.shiftKey &&
              document.activeElement === ae &&
              (event.preventDefault(), te == null || te.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: a,
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
              onClick: e.onClose,
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
                  onClick: () => d.current?.click(),
                  children: i.jsx(LocalizedText, {
                    children: "Inspect automation file",
                  }),
                }),
                i.jsx("input", {
                  ref: d,
                  type: "file",
                  hidden: true,
                  "aria-label": text("Automation source file"),
                  accept:
                    ".bas,.vba,.cls,.ts,.osts,.xlsm,.xlam,.xltm,.xlsx,.xltx,.xls,.xlsb",
                  onChange: (event) => void G(event.target.files?.[0]),
                  lang: language("Automation source file"),
                  "data-ui-label": "Automation source file",
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Preserved source" }),
                    i.jsxs("select", {
                      value: C,
                      onChange: (event) =>
                        U(() => {
                          ($(), T(event.target.value));
                          const W = e.workbook.automationArtifacts?.find(
                            (automationArtifact) =>
                              automationArtifact.id === event.target.value,
                          );
                          W &&
                            (G(
                              new File(
                                [decodeAutomationArtifact(W.base64)],
                                W.fileName,
                              ),
                            ),
                            L());
                        }),
                      children: [
                        i.jsx("option", {
                          value: "",
                          children: "Select an artifact",
                        }),
                        e.workbook.automationArtifacts?.map(
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
                  disabled: !C || e.readOnly,
                  onClick: () =>
                    U(() => {
                      (e.onRemove(C), T(""), L());
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
                          value: g,
                          onChange: (event) => {
                            ($(),
                              b(event.target.value),
                              w(
                                event.target.value === "vba"
                                  ? "Reviewed-module.bas"
                                  : "Reviewed-script.ts",
                              ),
                              L(),
                              h(null));
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
                          value: y,
                          onChange: (event) => {
                            ($(), w(event.target.value), L(), h(null));
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
                      value: m,
                      onChange: (event) => {
                        ($(), v(event.target.value), L(), h(null));
                      },
                    }),
                  ],
                }),
                i.jsx("button", {
                  onClick: () =>
                    U(() => {
                      if (
                        ($(),
                        (g === "vba" && !/\.(bas|vba|cls)$/i.test(y)) ||
                          (g === "office-script" && !/\.ts$/i.test(y)))
                      )
                        throw Error(
                          "Match the source file extension to the chosen language.",
                        );
                      (G(new File([m], y)), L(), T(""));
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
                        children: inspection.items.map((item, W) =>
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
                            W,
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
                    children: inspection.warnings.map((warning, W) =>
                      i.jsx("li", { children: warning }, W),
                    ),
                  }),
                  i.jsxs("div", {
                    className: "lg-automation-actions",
                    children: [
                      i.jsx("button", {
                        disabled: e.readOnly || !inspection.retainedEligible,
                        onClick: () =>
                          U(() => {
                            (e.onPreserve(inspection), p(null), D(false));
                          }),
                        children: i.jsx(LocalizedText, {
                          children: "Preserve original source",
                        }),
                      }),
                      i.jsx("button", {
                        onClick: () =>
                          U(() => {
                            const q =
                              l.current.workbook.automationArtifacts?.find(
                                (automationArtifact) =>
                                  automationArtifact.sha256 ===
                                    inspection.sha256 &&
                                  automationArtifact.fileName ===
                                    inspection.fileName,
                              );
                            if (!q)
                              throw Error(
                                "Preserve the original source before downloading it from the workbook.",
                              );
                            e.onDownload(
                              decodeAutomationArtifact(q.base64),
                              q.fileName,
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
                          children: compatibility.issues.map((issue, W) =>
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
                              W,
                            ),
                          ),
                        }),
                        i.jsx("ul", {
                          children: compatibility.warnings.map((warning, W) =>
                            i.jsx("li", { children: warning }, W),
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
                                    (A(event.target.value), L());
                                  },
                                  children: e.workbook.sheets
                                    .filter(
                                      (sheet) =>
                                        !sheet.hidden && !sheet.veryHidden,
                                    )
                                    .map((q) =>
                                      i.jsx(
                                        "option",
                                        { value: q.id, children: q.name },
                                        q.id,
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
                                    (F(event.target.value), L());
                                  },
                                }),
                              ],
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          disabled: e.readOnly || !compatibility.supported,
                          onClick: () =>
                            U(() => {
                              (D(false),
                                p(
                                  buildAutomationChanges(
                                    e.workbook,
                                    inspection,
                                    R,
                                    {
                                      readOnly: e.readOnly,
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
                        children: plan.issues.map((issue, W) =>
                          i.jsx("li", { children: issue }, W),
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
                          (command, W) =>
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
                              W,
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
                        checked: B,
                        onChange: (event) => {
                          (_(event.target.checked), D(false));
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
                        disabled: !plan.valid || e.readOnly,
                        onChange: (event) => D(event.target.checked),
                      }),
                      i.jsx(LocalizedText, {
                        children:
                          "I approve these commands and their changes within this scope",
                      }),
                    ],
                  }),
                ],
              }),
            !!e.workbook.automationRuns?.length &&
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Recorded automation runs (",
                      }),
                      e.workbook.automationRuns.length,
                      ")",
                    ],
                  }),
                  i.jsx("ul", {
                    children: e.workbook.automationRuns.slice(-10).map((q) =>
                      i.jsxs(
                        "li",
                        {
                          children: [
                            q.sourceName,
                            " · ",
                            q.actor,
                            " · ",
                            q.timestamp,
                            " ",
                            i.jsx(LocalizedText, { children: "· revision" }),
                            " ",
                            q.beforeRevision,
                            " → ",
                            q.afterRevision,
                            " · ",
                            q.changedCells,
                            " ",
                            i.jsx(LocalizedText, {
                              children: "writes · source ",
                            }),
                            q.sourceSha256,
                          ],
                        },
                        q.id,
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
        e.readOnly &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children:
                "Read-only: inspection and report download remain available; preservation and execution are disabled.",
            }),
          }),
        I && i.jsx("p", { role: "alert", children: I }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              disabled: !inspection,
              onClick: () =>
                e.onDownload(
                  H(),
                  `${inspection?.fileName ?? "automation"}.compatibility.json`,
                ),
              children: i.jsx(LocalizedText, {
                children: "Save automation report",
              }),
            }),
            i.jsx("button", {
              disabled: !checked || !plan?.valid || e.readOnly,
              onClick: () =>
                U(() => {
                  !plan ||
                    !checked ||
                    (e.onApply(plan, R, B), p(null), D(false));
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
