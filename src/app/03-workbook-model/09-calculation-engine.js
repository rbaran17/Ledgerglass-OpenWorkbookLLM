import {
  resolveExtensionFunctionCall,
  isWithinNestingLimit,
  createCalculationBudgetTracker,
  resolveLocalDateTime,
  listInstalledExtensionFunctions,
  readMonotonicClockMs,
  calculationFeatureSupport,
  findVolatileFunctionNames,
  workbookUsesNowOrToday,
  workbookHasVolatileFunctions,
} from "../01-schemas-date-functions-and-validation.js";
import {
  fileImportLimits,
  workbookLimits,
} from "../02-text-and-financial-functions.js";
import {
  string,
  number,
  boolean,
  null as _null,
  array,
  literal,
  union,
  enum as _enum,
  unknown,
  object,
} from "zod";
import {
  FB,
  validationLimits,
  DataValidationError,
  validateDataValidationRule,
  isCellInsideRange,
  r5,
} from "./01-cell-styles-and-themes.js";
import {
  FormulaParser,
  matrixResultFunctionNames,
} from "./02-formula-function-evaluators.js";
import { evaluateReusableFormula, y5 } from "./03-defined-names.js";
import {
  builtinFunctionNames,
  findFormulaErrorArg,
  createFormulaError,
  resolveDefinedNameValue,
  buildExternalLinkAliases,
  isFormulaError,
} from "./04-pivot-tables-and-pictures.js";
import {
  hasCalculatorRecord,
  deepFreeze,
  makeFormulaError,
  createWorkbookFormulaEngine,
  reportCellValueMaps,
  getCachedCalculationSnapshot,
  _O,
  carriedCalculatorRecords,
  ik,
  j5,
  getCalculationInputSignature,
  pendingFormulaValue,
  isCalculationRecordCurrent,
  buildReportCellValueMap,
  calculatorRecords,
} from "./05-formula-engine-and-name-operations.js";
import {
  AU,
  nonEmptyLabelSchema,
  reportOutputFormatVersion,
  createBoundedTextSchema,
  columnIndexToLetters,
  shiftFormulaReferences,
  reportHashSchema,
} from "./06-data-model-and-analysis.js";
import {
  findDataValidationsInRange,
  resolveFilteredRowContext,
  normalizeConditionalRules,
  evaluateConditionalFormatRules,
  _t,
  parseRangeAddress,
} from "./07-conditional-rules-worksheet-features.js";
import {
  rewriteFormulaReferences,
  rewriteStructuredReferences,
} from "./08-table-pivot-query-definitions.js";

function evaluateWorkbookFormulas(workbook, r = {}, t, n) {
  const installedExtensionFunctions = listInstalledExtensionFunctions(workbook),
    supported = new Set([
      ...builtinFunctionNames,
      ...installedExtensionFunctions.map(
        (installedExtensionFunction) => installedExtensionFunction.name,
      ),
    ]),
    s = n?.budget ?? createCalculationBudgetTracker(workbook),
    clock = n?.dateContext ?? resolveLocalDateTime(workbook.datePolicy),
    u = new Set(),
    l = new Map(),
    c = new Map(),
    h = new Map(),
    issue = (te) => (
      [
        "#LIMIT!",
        "#BLOCKED!",
        "#CYCLE!",
        "#ERROR!",
        "#TIMEOUT!",
        "#ITER!",
      ].includes(te) && u.add(te),
      createFormulaError(te)
    ),
    filteredRowContext = resolveFilteredRowContext(workbook, r),
    m = new Map(),
    v = new Map(),
    g = new Map(),
    b = new Map(),
    y = new Set(),
    w = new Set(),
    E = new Map(),
    A = [],
    k = [];
  let F = 0;
  const I = [];
  let O = 0;
  const j = workbook.definedNames?.some((definedName) =>
      /\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(definedName.refersTo),
    ),
    z = workbook.definedNames?.length
      ? new RegExp(
          `(?:^|[^A-Za-z0-9_.\\\\])(?:${workbook.definedNames.map((definedName) => definedName.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|")})(?![A-Za-z0-9_.\\\\])`,
          "i",
        )
      : undefined,
    V = (te) =>
      /\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(te) || !!(j && z?.test(te)),
    externalLinkAliases = buildExternalLinkAliases(workbook),
    B = [...workbook.sheets, ...externalLinkAliases.sheets],
    _ = new Map(B.map((te) => [te.id, te])),
    C = new Map(B.map((te) => [te.name.toLowerCase(), te]));
  let T = 0,
    N = 0;
  const R = new Map(),
    L = new RegExp(
      `\\b(?:${[...matrixResultFunctionNames, "ISFORMULA", "TEXTBEFORE", "TEXTAFTER", ...installedExtensionFunctions.filter((installedExtensionFunction) => installedExtensionFunction.fn.result.type === "matrix").map((te) => te.name.replaceAll(".", "\\."))].join("|")})\\s*\\(`,
      "i",
    );
  for (const sheet of workbook.sheets)
    R.set(
      sheet.id,
      Object.entries(sheet.cells)
        .filter(
          ([, cell]) =>
            typeof cell.raw == "string" &&
            cell.raw.startsWith("=") &&
            !cell.literal &&
            (V(cell.raw) ||
              L.test(cell.raw) ||
              (/\b(?:IF|IFERROR|IFNA)\s*\(/i.test(cell.raw) &&
                cell.raw.replace(/"(?:[^"]|"")*"/g, '""').includes(":")) ||
              /[{}#]/.test(cell.raw.replace(/"(?:[^"]|"")*"/g, '""')) ||
              /^=\s*[^()+*/]*:/.test(cell.raw) ||
              (workbook.definedNames?.length &&
                /^=\s*[A-Za-z_][A-Za-z0-9_.]*\s*$/.test(cell.raw))),
        )
        .map(([ae]) => ae)
        .sort((ae, oe) => {
          const de = _t(ae),
            he = _t(oe);
          return de.row - he.row || de.col - he.col;
        }),
    );
  const $ = (te, ae, address, de = true) => {
      const he = (se, we = true) => {
        const Ee = rewriteFormulaReferences(se, workbook, ae, address, we);
        if (Ee) throw (l.set(`${ae}:${address}`, Ee), issue(Ee.error));
      };
      ((te = externalLinkAliases.rewrite(te)), he(te, false));
      let Se = de
        ? resolveDefinedNameValue(te, workbook, ae, [], t?.allowName)
        : te;
      ((Se = externalLinkAliases.rewrite(Se)), he(Se));
      let le = "",
        ce = 0;
      for (; ce < Se.length;) {
        const se =
          /^(?:'((?:[^']|'')*)'|([A-Za-z_][A-Za-z0-9_.]*))!\s*(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#/i.exec(
            Se.slice(ce),
          );
        if (se) {
          const ye = C.get((se[1]?.replace(/''/g, "'") ?? se[2]).toLowerCase());
          if (!ye) throw issue("#REF!");
          const Te = _t(se[3]),
            Pe = `${columnIndexToLetters(Te.col)}${Te.row + 1}`,
            cell = readCell(ye.id, Pe),
            ht = g.get(`${ye.id}:${Pe}`);
          if (!ht) throw isFormulaError(cell) ? cell : issue("#REF!");
          ((le += `'${ye.name.replace(/'/g, "''")}'!${ht}`),
            (ce += se[0].length));
          continue;
        }
        if (Se[ce] === '"' || Se[ce] === "'") {
          const ye = Se[ce],
            Te = ce++;
          for (; ce < Se.length;)
            if (Se[ce++] === ye)
              if (Se[ce] === ye) ce++;
              else break;
          le += Se.slice(Te, ce);
          continue;
        }
        const we =
          /^(SEQUENCE|FILTER|SORT|SORTBY|UNIQUE|TRANSPOSE|TAKE|DROP|TEXTSPLIT)\s*\(\s*,/i.exec(
            Se.slice(ce),
          );
        if (we) {
          ((le += `${we[1]}(${we[1].toUpperCase() === "SEQUENCE" ? "1" : "#VALUE!"},`),
            (ce += we[0].length));
          continue;
        }
        const Ee = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(Se.slice(ce));
        if (Ee) {
          const ye = ce;
          ce += Ee[0].length;
          let Te = 1;
          for (; ce < Se.length && Te;)
            (Se[ce] === "[" && Te++, Se[ce] === "]" && Te--, ce++);
          const Pe =
            !Te &&
            rewriteStructuredReferences(
              workbook,
              ae,
              Se.slice(ye, ce),
              address,
            );
          if (!Pe) throw issue("#REF!");
          const qe = _.get(Pe.sheetId);
          le += `'${qe.name.replace(/'/g, "''")}'!${Pe.range}`;
          continue;
        }
        const Ae = /^(\$?[A-Za-z]{1,3}\$?[1-9]\d*)#/.exec(Se.slice(ce));
        if (Ae) {
          const ye = _t(Ae[1]),
            Te = `${columnIndexToLetters(ye.col)}${ye.row + 1}`,
            cell = readCell(ae, Te),
            qe = g.get(`${ae}:${Te}`);
          if (!qe) throw isFormulaError(cell) ? cell : issue("#REF!");
          ((le += qe), (ce += Ae[0].length));
          continue;
        }
        le += Se[ce++];
      }
      return le.replace(/\b_xlfn\.(?:_xlws\.)?/gi, "");
    },
    U = (te) => {
      const ae = C.get((te ?? "").toLowerCase());
      if (!ae) throw issue("#REF!");
      return ae;
    },
    G = () => {
      const workbookFormulaEngine = createWorkbookFormulaEngine(
        new FormulaParser({
          onVariable: () => null,
          onCell: ({ sheet, row, col }) => {
            const se = U(sheet);
            if (
              !Number.isInteger(row) ||
              !Number.isInteger(col) ||
              row < 1 ||
              col < 1 ||
              row > se.rowCount ||
              col > se.colCount
            )
              throw issue("#REF!");
            return readCell(se.id, `${columnIndexToLetters(col - 1)}${row}`);
          },
          onRange: ({ sheet, from, to }) => {
            const se = U(sheet),
              row = to.row,
              col = to.col;
            if (
              from.row < 1 ||
              from.col < 1 ||
              row > se.rowCount ||
              col > se.colCount ||
              row < from.row ||
              col < from.col
            )
              throw issue("#REF!");
            if ((row - from.row + 1) * (col - from.col + 1) > 2e4)
              throw issue("#LIMIT!");
            const Ae = s.reserve(
              (row - from.row + 1) * (32 + (col - from.col + 1) * 8),
            );
            if (Ae) throw issue(Ae);
            const ye = [];
            for (let Te = from.row; Te <= row; Te++) {
              const Pe = [];
              for (let qe = from.col; qe <= col; qe++) {
                const cell = readCell(
                  se.id,
                  `${columnIndexToLetters(qe - 1)}${Te}`,
                );
                if (
                  isFormulaError(cell) &&
                  ["#LIMIT!", "#TIMEOUT!", "#BLOCKED!"].includes(cell.error)
                )
                  throw cell;
                Pe.push(cell);
              }
              ye.push(Pe);
            }
            return ye;
          },
        }),
        (le, ce, ge) =>
          filteredRowContext(le, ce, {
            currentSheetName: ge.position?.sheet,
            retrieve: (we) => ge.retrieveRef(we),
            issue: issue,
            isIssue: isFormulaError,
            readCell: readCell,
            readMetadata:
              t || n
                ? (sheetId, address) => {
                    if (
                      (n?.read(`${sheetId}:${address}`),
                      t && !t.allowRead(sheetId, address))
                    )
                      throw issue("#BLOCKED!");
                    t?.step({
                      kind: "metadata",
                      sheetId: sheetId,
                      address: address,
                      operation: "Inspect cell metadata",
                      result: null,
                    });
                  }
                : undefined,
            isActive: (we, Ee) => y.has(`${we}:${Ee}`),
            expandedFormula: (cell, Ee) => {
              try {
                return resolveDefinedNameValue(String(cell.raw), workbook, Ee);
              } catch {
                return String(cell.raw);
              }
            },
          }),
        issue,
        clock,
        (le) => {
          const position = workbookFormulaEngine.position;
          if (!position) return;
          const sheetId = U(position.sheet).id,
            address = `${columnIndexToLetters(position.col - 1)}${position.row}`,
            we = `${sheetId}:${address}:${le.functionName}`;
          if (h.size >= 1e4 && !h.has(we)) throw issue("#LIMIT!");
          h.set(we, { ...le, sheetId: sheetId, address: address });
        },
        (le, ce, ge) => {
          const se =
            le && typeof le == "object" && "ref" in le ? le.ref : undefined;
          if (!se?.from || !se.to) return;
          const we = U(se.sheet ?? workbookFormulaEngine.position?.sheet),
            Ee = `${columnIndexToLetters(se.from.col + ge - 1)}${se.from.row + ce}`;
          if (t && !t.allowRead(we.id, Ee)) throw issue("#BLOCKED!");
          n?.read(`${we.id}:${Ee}`);
          const cell = we.cells[Ee];
          if (
            cell?.literal ||
            typeof cell?.raw != "string" ||
            !cell.raw.startsWith("=")
          )
            return;
          const raw = cell.raw;
          return (Pe) => {
            if (O >= 32) throw issue("#LIMIT!");
            const qe = O++;
            try {
              if (
                t?.allowDatabaseCriterion &&
                !t.allowDatabaseCriterion(
                  we.id,
                  shiftFormulaReferences(raw, Pe, 0),
                )
              )
                throw issue("#BLOCKED!");
              const definedNameValue = resolveDefinedNameValue(
                  raw,
                  workbook,
                  we.id,
                  [],
                  t?.allowName,
                ),
                Re = shiftFormulaReferences(definedNameValue, Pe, 0);
              if (Re.length > fileImportLimits.maxFormulaLength)
                throw issue("#LIMIT!");
              return V(Re)
                ? H(Re.slice(1), we.id, Ee)
                : (I[qe] ??= G()).parse(
                    $(Re.slice(1), we.id, Ee),
                    {
                      sheet: we.name,
                      row: se.from.row + ce,
                      col: se.from.col + ge,
                    },
                    true,
                  );
            } finally {
              O--;
            }
          };
        },
      );
      if (installedExtensionFunctions.length) {
        const le = workbookFormulaEngine._callFunction.bind(
          workbookFormulaEngine,
        );
        workbookFormulaEngine._callFunction = (ce, ge) => {
          const se = ce.replace(/^_xlfn\./i, "").toUpperCase();
          if (
            !installedExtensionFunctions.some(
              (installedExtensionFunction) =>
                installedExtensionFunction.name === se,
            )
          )
            return le(ce, ge);
          const position = workbookFormulaEngine.position;
          if (!position) return issue("#NAME?");
          const Ee = U(position.sheet).id,
            Ae = `${columnIndexToLetters(position.col - 1)}${position.row}`,
            ye = ge.map((ht) => workbookFormulaEngine.retrieveRef(ht)),
            formulaErrorArg = findFormulaErrorArg(ye);
          if (formulaErrorArg) return formulaErrorArg;
          const extensionFunctionCall = resolveExtensionFunctionCall(
            workbook,
            se,
            ye,
            Ee,
            Ae,
          );
          if (extensionFunctionCall.error)
            return issue(extensionFunctionCall.error);
          const value = extensionFunctionCall.value;
          return value && !Array.isArray(value) && typeof value == "object"
            ? value.label
            : value;
        };
      }
      const ae = (le) => {
          const ce = s.work() ?? n?.check();
          if (ce) return issue(ce);
          const ge = le(),
            se = s.value(ge) ?? n?.check();
          return se ? issue(se) : ge;
        },
        oe = workbookFormulaEngine._callFunction.bind(workbookFormulaEngine),
        de = workbookFormulaEngine.utils._applyInfix.bind(
          workbookFormulaEngine.utils,
        ),
        he = workbookFormulaEngine.utils._applyPrefix.bind(
          workbookFormulaEngine.utils,
        ),
        Se = workbookFormulaEngine.utils._applyPostfix.bind(
          workbookFormulaEngine.utils,
        );
      if (
        ((workbookFormulaEngine._callFunction = (le, ce) =>
          ae(() => oe(le, ce))),
        (workbookFormulaEngine.utils._applyInfix = (le, ce, ge) =>
          ae(() => de(le, ce, ge))),
        (workbookFormulaEngine.utils._applyPrefix = (le, ce, ge) =>
          ae(() => he(le, ce, ge))),
        (workbookFormulaEngine.utils._applyPostfix = (le, ce, ge) =>
          ae(() => Se(le, ce, ge))),
        t)
      ) {
        const le = (Ee, operation, inputs, Te) => {
            const position = workbookFormulaEngine.position,
              sheetId = U(position.sheet).id,
              address = `${columnIndexToLetters(position.col - 1)}${position.row}`;
            try {
              const Re = Te();
              return (
                t.step({
                  kind: Ee,
                  sheetId: sheetId,
                  address: address,
                  operation: operation,
                  inputs: inputs,
                  result: Re,
                }),
                Re
              );
            } catch (error) {
              throw (
                t.step({
                  kind: Ee,
                  sheetId: sheetId,
                  address: address,
                  operation: operation,
                  inputs: inputs,
                  result: isFormulaError(error) ? error : issue("#ERROR!"),
                }),
                error
              );
            }
          },
          ce = workbookFormulaEngine._callFunction.bind(workbookFormulaEngine),
          ge = workbookFormulaEngine.utils._applyInfix.bind(
            workbookFormulaEngine.utils,
          ),
          se = workbookFormulaEngine.utils._applyPrefix.bind(
            workbookFormulaEngine.utils,
          ),
          we = workbookFormulaEngine.utils._applyPostfix.bind(
            workbookFormulaEngine.utils,
          );
        ((workbookFormulaEngine._callFunction = (Ee, Ae) =>
          le("function", Ee, Ae, () => ce(Ee, Ae))),
          (workbookFormulaEngine.utils._applyInfix = (Ee, Ae, ye) =>
            le("operator", Ae, [Ee.val, ye.val], () => ge(Ee, Ae, ye))),
          (workbookFormulaEngine.utils._applyPrefix = (Ee, Ae, ye) =>
            le("operator", Ee.join(""), [Ae], () => se(Ee, Ae, ye))),
          (workbookFormulaEngine.utils._applyPostfix = (Ee, Ae, ye) =>
            le("operator", ye, [Ee], () => we(Ee, Ae, ye))));
      }
      return workbookFormulaEngine;
    },
    H = (te, sheetId, address) =>
      evaluateReusableFormula(te, workbook, sheetId, {
        supported: supported,
        issue: issue,
        isIssue: isFormulaError,
        allowName: t?.allowName,
        step: t
          ? (de, operation, Se) =>
              t.step({
                kind: de,
                sheetId: sheetId,
                address: address,
                operation: operation,
                result: Se,
              })
          : undefined,
        diagnostic: (de, he) =>
          c.set(`${sheetId}:${address}`, {
            sheetId: sheetId,
            address: address,
            error: de,
            message: he,
          }),
        parse: (de, he, Se) => {
          const le = $(de, he, address, false),
            ce = le.replace(/"(?:[^"]|"")*"/g, '""');
          if (/[\[\]|]/.test(ce) || /(?:https?:|file:|\\\\)/i.test(ce))
            throw issue("#BLOCKED!");
          if (
            Array.from(
              ce
                .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
                .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
            ).some((Ee) => Ee[1] !== y5 && !supported.has(Ee[1].toUpperCase()))
          )
            throw issue("#NAME?");
          const ge = F++,
            se = (k[ge] ??= G()),
            we = se._callFunction;
          se._callFunction = (Ee, Ae) =>
            Ee === y5
              ? Ae.length === 1 &&
                Number.isInteger(Ae[0]) &&
                Number(Ae[0]) >= 0 &&
                Number(Ae[0]) < Se.length
                ? Se[Number(Ae[0])]
                : issue("#NAME?")
              : we(Ee, Ae);
          try {
            const Ee = _t(address);
            return se.parse(
              le,
              { sheet: _.get(he).name, row: Ee.row + 1, col: Ee.col + 1 },
              true,
            );
          } finally {
            ((se._callFunction = we), F--);
          }
        },
      }),
    readCell = (sheetId, address) => {
      const oe = s.work() ?? n?.check();
      if (oe) return issue(oe);
      if (!t) return X(sheetId, address);
      if (!t.allowRead(sheetId, address)) return issue("#BLOCKED!");
      const de = X(sheetId, address);
      return (
        t.step({
          kind: "read",
          sheetId: sheetId,
          address: address,
          operation: "Read cell",
          result: de,
        }),
        de
      );
    },
    X = (sheetId, ae) => {
      if (++T > workbookLimits.maxCells) return issue("#LIMIT!");
      let firstRow, firstCol;
      try {
        ({ row: firstRow, col: firstCol } = _t(ae));
      } catch {
        return issue("#REF!");
      }
      const sheet = _.get(sheetId);
      if (!sheet || firstRow >= sheet.rowCount || firstCol >= sheet.colCount)
        return issue("#REF!");
      ae = `${columnIndexToLetters(firstCol)}${firstRow + 1}`;
      const Se = `${sheetId}:${ae}`,
        le = externalLinkAliases.values.get(Se);
      if (le) return le.type === "error" ? issue(String(le.value)) : le.value;
      n?.read(Se);
      const ce = n?.value(Se, y.size > 0 || w.size > 0);
      if (ce) return ce.type === "error" ? issue(String(ce.value)) : ce.value;
      for (const Te of y) E.get(Te)?.add(Se);
      if (y.has(Se) || w.has(Se))
        return (n?.cycle([...y, ...w, Se]), issue("#CYCLE!"));
      if (m.has(Se)) return m.get(Se);
      if (v.has(Se)) return v.get(Se).value;
      const cell = sheet.cells[ae];
      if (!cell || cell.raw === null) {
        const Te = R.get(sheetId) ?? [];
        if (Te.length > 1e3) return issue("#LIMIT!");
        for (const Pe of Te) {
          if (Pe === ae) continue;
          const qe = _t(Pe);
          if (
            !(
              qe.row > firstRow ||
              qe.col > firstCol ||
              (firstRow - qe.row + 1) * (firstCol - qe.col + 1) >
                fileImportLimits.maxSpillCells
            ) &&
            !(y.has(`${sheetId}:${Pe}`) || w.has(`${sheetId}:${Pe}`)) &&
            (readCell(sheetId, Pe), v.has(Se))
          )
            return v.get(Se).value;
        }
        return cell?.raw ?? null;
      }
      if (
        typeof cell.raw != "string" ||
        !cell.raw.startsWith("=") ||
        cell.literal
      )
        return cell.raw;
      const raw = cell.raw;
      n?.enter(Se);
      let Ee = (() => {
        if (
          y.size + w.size >= fileImportLimits.maxFormulaDepth ||
          raw.length > fileImportLimits.maxFormulaLength ||
          !isWithinNestingLimit(raw)
        )
          return issue("#LIMIT!");
        let Te;
        const Pe = V(raw);
        w.add(Se);
        try {
          if (
            ((Te = Pe ? raw.slice(1) : $(raw.slice(1), sheetId, ae)),
            !isWithinNestingLimit(Te))
          )
            return issue("#LIMIT!");
        } catch (error) {
          return isFormulaError(error)
            ? issue(error.error)
            : issue(
                error instanceof Error && /^#/.test(error.message)
                  ? error.message
                  : "#NAME?",
              );
        } finally {
          w.delete(Se);
        }
        const qe = Te.replace(/"(?:[^"]|"")*"/g, '""');
        if (!Pe && (/[\[\]|]/.test(qe) || /(?:https?:|file:|\\\\)/i.test(qe)))
          return issue("#BLOCKED!");
        if (
          !Pe &&
          Array.from(
            qe
              .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
              .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
          ).some((Je) => !supported.has(Je[1].toUpperCase()))
        )
          return (
            c.set(Se, {
              sheetId: sheetId,
              address: ae,
              error: "#NAME?",
              message:
                "The formula calls an unknown or unsupported custom function. Imported VBA, add-in and external custom functions cannot execute.",
            }),
            issue("#NAME?")
          );
        const size = y.size;
        (y.add(Se), E.set(Se, new Set()));
        let Re;
        try {
          const Je = (A[size] ??= G()),
            ze = Pe
              ? H(Te, sheetId, ae)
              : Je.parse(
                  Te,
                  { sheet: sheet.name, row: firstRow + 1, col: firstCol + 1 },
                  true,
                );
          if (isFormulaError(ze)) Re = issue(ze.error);
          else if (typeof ze == "number")
            Re = Number.isFinite(ze) ? ze : issue("#NUM!");
          else if (ze === null) Re = 0;
          else if (typeof ze == "boolean") Re = ze;
          else if (typeof ze == "string")
            Re = ze.length <= 32767 ? ze : issue("#LIMIT!");
          else if (Array.isArray(ze)) {
            const at = ze.every(Array.isArray) ? ze : [ze],
              _e = at[0]?.length ?? 0;
            if (!at.length || !_e || at.some((Ze) => Ze.length !== _e))
              Re = issue("#CALC!");
            else if (
              at.length * _e > fileImportLimits.maxSpillCells ||
              N + at.length * _e > workbookLimits.maxCells
            )
              Re = issue("#LIMIT!");
            else if (
              firstRow + at.length > sheet.rowCount ||
              firstCol + _e > sheet.colCount
            )
              (b.set(Se, {
                sheetId: sheetId,
                anchor: ae,
                reason: "bounds",
                message:
                  "The array extends beyond the worksheet edge. Move the formula or reduce its output.",
              }),
                (Re = issue("#SPILL!")));
            else {
              let Ze = false,
                lt = false;
              const be = {
                  firstRow: firstRow,
                  lastRow: firstRow + at.length - 1,
                  firstCol: firstCol,
                  lastCol: firstCol + _e - 1,
                },
                ft = (nt) => {
                  const rangeAddress = parseRangeAddress(nt);
                  return (
                    rangeAddress.firstRow <= be.lastRow &&
                    rangeAddress.lastRow >= be.firstRow &&
                    rangeAddress.firstCol <= be.lastCol &&
                    rangeAddress.lastCol >= be.firstCol
                  );
                },
                pt = sheet.merges?.find(ft),
                Fe = sheet.tables?.find((table) => ft(table.range));
              if (pt || Fe) {
                Ze = true;
                const obstruction = pt ?? Fe.range;
                b.set(Se, {
                  sheetId: sheetId,
                  anchor: ae,
                  reason: pt ? "merged" : "table",
                  obstruction: obstruction,
                  message: pt
                    ? `The array overlaps merged cells ${obstruction}. Unmerge those cells or move the formula.`
                    : `The array overlaps table ${Fe.name} (${obstruction}). Place the formula outside the table.`,
                });
              }
              for (let nt = 0; nt < at.length && !Ze; nt++)
                for (let mt = 0; mt < _e; mt++) {
                  const obstruction = `${columnIndexToLetters(firstCol + mt)}${firstRow + nt + 1}`,
                    cell = sheet.cells[obstruction];
                  if (E.get(Se)?.has(`${sheetId}:${obstruction}`)) {
                    lt = true;
                    break;
                  }
                  if (
                    obstruction !== ae &&
                    ((cell && cell.raw !== null) ||
                      v.has(`${sheetId}:${obstruction}`))
                  ) {
                    ((Ze = true),
                      b.set(Se, {
                        sheetId: sheetId,
                        anchor: ae,
                        reason: "occupied",
                        obstruction: obstruction,
                        message: `Cell ${obstruction} blocks the array. Move the formula or review and clear the obstruction; existing evidence was preserved.`,
                      }));
                    break;
                  }
                }
              if (lt) Re = issue("#CYCLE!");
              else if (Ze) Re = issue("#SPILL!");
              else {
                const nt = (mt) =>
                  isFormulaError(mt)
                    ? mt
                    : mt === null
                      ? 0
                      : typeof mt == "number"
                        ? Number.isFinite(mt)
                          ? mt
                          : issue("#NUM!")
                        : typeof mt == "boolean"
                          ? mt
                          : typeof mt == "string"
                            ? mt.length <= 32767
                              ? mt
                              : issue("#LIMIT!")
                            : issue("#VALUE!");
                for (let mt = 0; mt < at.length; mt++)
                  for (let Wt = 0; Wt < _e; Wt++) {
                    const Zt = `${columnIndexToLetters(firstCol + Wt)}${firstRow + mt + 1}`;
                    Zt !== ae &&
                      v.set(`${sheetId}:${Zt}`, {
                        value: nt(at[mt][Wt]),
                        anchor: ae,
                      });
                  }
                (g.set(
                  Se,
                  `${ae}:${columnIndexToLetters(firstCol + _e - 1)}${firstRow + at.length}`,
                ),
                  (N += at.length * _e),
                  (Re = nt(at[0][0])));
              }
            }
          } else Re = issue("#VALUE!");
        } catch (error) {
          const ze =
            error && typeof error == "object" && "details" in error
              ? error.details
              : undefined;
          Re = isFormulaError(ze)
            ? issue(ze.error)
            : isFormulaError(error)
              ? issue(error.error)
              : issue("#ERROR!");
        } finally {
          (y.delete(Se), E.delete(Se));
        }
        return Re;
      })();
      const Ae = s.value(Ee) ?? n?.check();
      return (
        Ae && (Ee = issue(Ae)),
        m.set(Se, Ee),
        n?.leave(
          Se,
          isFormulaError(Ee)
            ? { type: "error", value: Ee.error }
            : {
                type:
                  Ee === null
                    ? "blank"
                    : typeof Ee == "string"
                      ? "text"
                      : typeof Ee == "boolean"
                        ? "boolean"
                        : "number",
                value: Ee,
              },
        ),
        t?.step({
          kind: "formula",
          sheetId: sheetId,
          address: ae,
          operation: cell.raw,
          result: Ee,
        }),
        Ee
      );
    },
    J = (te, ae) => J.getTypedValue(te, ae).value;
  J.getTypedValue = (te, ae) => {
    if (n?.running?.() === false)
      try {
        const Se = _t(ae),
          sheet = _.get(te);
        if (!sheet || Se.row >= sheet.rowCount || Se.col >= sheet.colCount)
          return { type: "error", value: "#REF!" };
        const ce = `${columnIndexToLetters(Se.col)}${Se.row + 1}`,
          ge = `${te}:${ce}`,
          se = n.value(ge, false);
        if (se) return { ...se };
        const we =
          m.get(ge) ?? v.get(ge)?.value ?? sheet.cells[ce]?.raw ?? null;
        return isFormulaError(we)
          ? { type: "error", value: we.error }
          : {
              type:
                we === null
                  ? "blank"
                  : typeof we == "string"
                    ? "text"
                    : typeof we == "boolean"
                      ? "boolean"
                      : "number",
              value: we,
            };
      } catch {
        return { type: "error", value: "#REF!" };
      }
    s.begin();
    try {
      T = 0;
      const cell = readCell(te, ae);
      return isFormulaError(cell)
        ? { type: "error", value: cell.error }
        : {
            type:
              cell === null
                ? "blank"
                : typeof cell == "string"
                  ? "text"
                  : typeof cell == "boolean"
                    ? "boolean"
                    : "number",
            value: cell,
          };
    } finally {
      s.end();
    }
  };
  let q;
  return (
    (J.evaluateFormula = (te, ae, oe) => {
      (s.begin(), (T = 0), l.delete(`${te}:${ae}`), c.delete(`${te}:${ae}`));
      const de = `expression:${te}:${ae}`;
      try {
        const he = _.get(te),
          Se = _t(ae);
        if (!he || Se.row >= he.rowCount || Se.col >= he.colCount)
          return "#REF!";
        if (typeof oe != "string" || !oe.startsWith("=")) return "#VALUE!";
        if (
          oe.length > fileImportLimits.maxFormulaLength ||
          !isWithinNestingLimit(oe)
        )
          return issue("#LIMIT!").error;
        const le = V(oe),
          ce = le ? oe.slice(1) : $(oe.slice(1), te, ae);
        if (!isWithinNestingLimit(ce)) return issue("#LIMIT!").error;
        const ge = ce.replace(/"(?:[^"]|"")*"/g, '""');
        if (!le && (/[\[\]|]/.test(ge) || /(?:https?:|file:|\\\\)/i.test(ge)))
          return issue("#BLOCKED!").error;
        if (
          !le &&
          Array.from(
            ge
              .replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet")
              .matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g),
          ).some((Ee) => !supported.has(Ee[1].toUpperCase()))
        )
          return "#NAME?";
        (y.add(de), E.set(de, new Set()));
        const se = le
            ? H(ce, te, ae)
            : (q ??= G()).parse(
                ce,
                { sheet: he.name, row: Se.row + 1, col: Se.col + 1 },
                true,
              ),
          we = s.value(se) ?? n?.check();
        return we
          ? issue(we).error
          : isFormulaError(se)
            ? issue(se.error).error
            : Array.isArray(se)
              ? "#VALUE!"
              : se === null
                ? 0
                : typeof se == "number"
                  ? Number.isFinite(se)
                    ? se
                    : "#NUM!"
                  : typeof se == "boolean"
                    ? se
                    : typeof se == "string"
                      ? se.length <= 32767
                        ? se
                        : issue("#LIMIT!").error
                      : "#VALUE!";
      } catch (error) {
        const Se =
          error && typeof error == "object" && "details" in error
            ? error.details
            : undefined;
        return isFormulaError(Se)
          ? issue(Se.error).error
          : isFormulaError(error)
            ? issue(error.error).error
            : error instanceof Error && /^#/.test(error.message)
              ? issue(error.message).error
              : issue("#ERROR!").error;
      } finally {
        (s.end(), y.delete(de), E.delete(de));
      }
    }),
    (J.getSpillRange = (te, ae) => (J(te, ae), g.get(`${te}:${ae}`))),
    (J.getSpillDiagnostic = (te, ae) => {
      J(te, ae);
      const oe = b.get(`${te}:${ae}`);
      return oe ? { ...oe } : undefined;
    }),
    (J.getSpillAnchor = (te, ae) => (J(te, ae), v.get(`${te}:${ae}`)?.anchor)),
    (J.getSpillEntries = (te) => {
      s.begin();
      try {
        const oe = R.get(te) ?? [];
        if (oe.length > 1e3)
          throw (
            issue("#LIMIT!"),
            new Error("Spill discovery exceeds 1,000 possible array formulas.")
          );
        T = 0;
        for (const he of n?.running?.() === false ? [] : oe)
          if ((readCell(te, he), T > workbookLimits.maxCells))
            throw new Error(
              "Spill discovery exceeds 100,000 calculation reads.",
            );
        const de = `${te}:`;
        return [...v.entries()]
          .filter(([he]) => he.startsWith(de))
          .map(([he, { anchor }]) => ({
            address: he.slice(de.length),
            anchor,
          }));
      } finally {
        s.end();
      }
    }),
    (J.getResourceReport = () => s.report()),
    (J.getOperationalErrors = () => [...u]),
    (J.getDateContext = () => ({ ...clock })),
    (J.getFinancialDiagnostics = () => structuredClone([...h.values()])),
    (J.peekSpill = (te, ae) => {
      const oe = `${te}:${ae}`,
        de = b.get(oe);
      return { range: g.get(oe), ...(de ? { diagnostic: { ...de } } : {}) };
    }),
    (J.getNamedFormulaDiagnostics = () =>
      [...c.values()].map((te) => ({ ...te }))),
    (J.getReferenceDiagnostics = () =>
      [...l.values()].map((te) => ({ ...te }))),
    J
  );
}
function evaluateValidationRule(workbook) {
  const r = evaluateWorkbookFormulas(workbook);
  return (sheetId, address, o) => {
    const a = r(sheetId, address);
    let s = false,
      reason = "Value does not satisfy the rule.",
      formula,
      evaluationError = false;
    try {
      if (r.getTypedValue(sheetId, address).type === "error")
        throw Error(`Cell calculation returned ${a}.`);
      if (a === null || a === "") s = o.allowBlank;
      else if (o.kind === "custom") {
        const c = _t(address),
          rangeAddress = parseRangeAddress(o.range);
        formula = shiftFormulaReferences(
          o.formula,
          c.row - rangeAddress.firstRow,
          c.col - rangeAddress.firstCol,
        );
        const f = r.evaluateFormula(sheetId, address, formula);
        if (
          typeof f != "boolean" &&
          (typeof f != "number" || !Number.isFinite(f))
        )
          throw Error(
            `Custom formula must return TRUE/FALSE or a finite number${typeof f == "string" && /^#[A-Z0-9/?!]+$/.test(f) ? `; returned ${f}` : ""}.`,
          );
        ((s = !!f), (reason = "Custom formula returned FALSE or zero."));
      } else s = validateDataValidationRule(o, a, workbook.datePolicy);
    } catch (error) {
      ((evaluationError = true),
        (reason =
          error instanceof Error
            ? error.message
            : "Validation could not be evaluated."));
    }
    if (!(s && !evaluationError))
      return {
        sheetId: sheetId,
        sheetName: workbook.sheets.find((sheet) => sheet.id === sheetId).name,
        address: address,
        ruleId: o.id,
        style: evaluationError ? "stop" : (o.errorStyle ?? "stop"),
        value: a,
        title: o.errorTitle || "Invalid value",
        message:
          o.message || "The value does not satisfy the data validation rule.",
        reason: reason,
        ...(formula ? { formula: formula } : {}),
        evaluationError: evaluationError,
      };
  };
}
function checkDataValidationBatch(book, r, validationApproval) {
  const n = r.flatMap((d) =>
    findDataValidationsInRange(book, d.sheetId, d.address).map((rule) => ({
      ...d,
      rule: rule,
    })),
  );
  if (!n.length) return 0;
  if (n.length > validationLimits.checks)
    throw Error(
      "Data validation exceeds the 100,000-check budget; split the batch.",
    );
  const o = evaluateValidationRule(book),
    issues = n.map((d) => o(d.sheetId, d.address, d.rule)).filter((d) => !!d);
  if (!issues.length) return 0;
  const s = JSON.stringify({ book: book, issues: issues });
  if (
    validationApproval &&
    r5.get(validationApproval) === s &&
    issues.every((issue) => issue.style !== "stop" && !issue.evaluationError)
  )
    return (r5.delete(validationApproval), issues.length);
  throw validationApproval
    ? Error(
        "Validation approval is stale or invalid. Review the complete batch again.",
      )
    : new DataValidationError(book, issues);
}
function summarizeValidationInspection(workbook, sheetId) {
  const t = {
      workbookId: workbook.id,
      revision: workbook.revision,
      checks: 0,
      invalidCells: 0,
      violations: 0,
      issues: [],
      complete: true,
      detailsTruncated: false,
      message: "",
    },
    n = evaluateValidationRule(workbook),
    o = new Set();
  let a = 0;
  e: for (const s of workbook.sheets.filter(
    (sheet) => !sheetId || sheet.id === sheetId,
  ))
    for (const d of s.validations ?? []) {
      const rangeAddress = parseRangeAddress(d.range);
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
          if (
            t.checks >= validationLimits.checks ||
            (d.kind === "custom" && ++a > validationLimits.customChecks)
          ) {
            t.complete = false;
            break e;
          }
          const h = `${columnIndexToLetters(firstCol)}${firstRow + 1}`;
          t.checks++;
          const f = n(s.id, h, d);
          f &&
            (o.add(`${s.id}:${h}`),
            t.violations++,
            t.issues.length < validationLimits.displayedIssues &&
              t.issues.push(f));
        }
    }
  return (
    (t.invalidCells = o.size),
    (t.detailsTruncated = t.violations > t.issues.length),
    (t.message = `${t.complete ? "" : "Incomplete inspection: at least "}${t.invalidCells} invalid cells; ${t.violations} rule violations across ${t.checks} checks.${t.complete ? "" : " The 100,000-check or 5,000-custom-formula limit was reached. Narrow validation ranges and inspect again."}${t.detailsTruncated ? " Showing the first 500 violations." : ""}`),
    FB.set(t, { source: JSON.stringify(workbook), content: JSON.stringify(t) }),
    t
  );
}
function cU(workbook, report, values, n) {
  const o = new Map(
      n.spills.map((spill) => [`${spill.sheetId}:${spill.address}`, spill]),
    ),
    a = new Map(n.spillMetadata.map((c) => [`${c.sheetId}:${c.address}`, c])),
    s = (c, h) => `${c}:${h.replace(/\$/g, "").toUpperCase()}`,
    d = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet])),
    getTypedValue = (c, h) => {
      const f = s(c, h),
        p = values.get(f) ?? o.get(f)?.result;
      if (p) return { ...p };
      const sheet = d.get(c);
      if (!sheet) return { type: "error", value: "#REF!" };
      try {
        const y = _t(h);
        if (y.row >= sheet.rowCount || y.col >= sheet.colCount)
          return { type: "error", value: "#REF!" };
      } catch {
        return { type: "error", value: "#REF!" };
      }
      const v = sheet.cells[h.replace(/\$/g, "").toUpperCase()]?.raw ?? null;
      return {
        type:
          v === null
            ? "blank"
            : typeof v == "string"
              ? "text"
              : typeof v == "boolean"
                ? "boolean"
                : "number",
        value: v,
      };
    },
    l = (c, h) => getTypedValue(c, h).value;
  return (
    Object.assign(l, {
      getTypedValue: getTypedValue,
      getDateContext: () => ({ ...report.dateContext }),
      getResourceReport: () => ({ ...report.resources }),
      getOperationalErrors: () => [...n.operationalErrors],
      getReferenceDiagnostics: () => structuredClone(n.references),
      getNamedFormulaDiagnostics: () => structuredClone(n.names),
      getFinancialDiagnostics: () => structuredClone(n.financial),
      getSpillEntries: (c) =>
        n.spills
          .filter((spill) => spill.sheetId === c)
          .map(({ address, anchor }) => ({ address, anchor })),
      getSpillAnchor: (c, h) => o.get(s(c, h))?.anchor,
      getSpillRange: (c, h) => a.get(s(c, h))?.range,
      getSpillDiagnostic: (c, h) => structuredClone(a.get(s(c, h))?.diagnostic),
      peekSpill: (c, h) => {
        const f = a.get(s(c, h));
        return f
          ? { range: f.range, diagnostic: structuredClone(f.diagnostic) }
          : {};
      },
      evaluateFormula: (c, h, f) =>
        evaluateWorkbookFormulas(workbook).evaluateFormula(c, h, f),
    }),
    l
  );
}
function getIsolatedCalculation(workbook) {
  const r = calculatorRecords.get(workbook);
  if (
    r?.report.execution?.state === "pending" &&
    isCalculationRecordCurrent(r, workbook)
  )
    return { calculator: r.calculator, report: r.report };
  const monotonicClockMs = readMonotonicClockMs(),
    cachedCalculationSnapshot = getCachedCalculationSnapshot(workbook),
    o =
      calculatorRecords.get(workbook) ?? carriedCalculatorRecords.get(workbook),
    values = new Map(),
    dependencies = new Map(),
    reportCellValueMap = buildReportCellValueMap(o),
    u = new Map();
  for (const v of cachedCalculationSnapshot.formulas.keys()) {
    values.set(v, pendingFormulaValue);
    const g = reportCellValueMap.get(v);
    g && u.set(v, g);
  }
  let l;
  const c = {
    workbookId: workbook.id,
    revision: workbook.revision,
    run: (o?.report.run ?? 0) + 1,
    settings: cachedCalculationSnapshot.settings,
    resources: createCalculationBudgetTracker(workbook).report(),
    volatile: {
      functions: cachedCalculationSnapshot.volatile,
      policy: calculationFeatureSupport,
    },
    dateContext: resolveLocalDateTime(workbook.datePolicy),
    reason: "Waiting for isolated calculation",
    formulaCount: cachedCalculationSnapshot.formulas.size,
    evaluated: 0,
    reused: 0,
    pending: cachedCalculationSnapshot.formulas.size,
    elapsedMs: readMonotonicClockMs() - monotonicClockMs,
    dependencyEdges: 0,
    cycles: [],
    iteration: {
      status: cachedCalculationSnapshot.settings.iterative
        ? "not-needed"
        : "off",
      passes: 0,
      maxDelta: null,
      message: "Waiting for isolated calculation",
    },
    execution: {
      mode: "worker",
      state: "pending",
      message:
        "Calculation is running separately; source inputs remain editable.",
    },
    limits: {
      dependencyEdges: j5,
      cooperativeTimeBudgetMs: cachedCalculationSnapshot.settings.timeBudgetMs,
      hardInterrupt: true,
    },
  };
  deepFreeze(c);
  const report = Object.freeze({
    ...c,
    get cells() {
      return (
        l ||
          ((l = [...cachedCalculationSnapshot.formulas].map(([v, g]) => ({
            ...g,
            result: { ...pendingFormulaValue },
            status: "pending",
            evaluations: 0,
            dependencies: [],
            ...(u.has(v) ? { lastKnown: u.get(v) } : {}),
          }))),
          deepFreeze(l)),
        l
      );
    },
  });
  reportCellValueMaps.set(report, u);
  const spills = [];
  if (o)
    for (const sheet of workbook.sheets)
      for (const g of o.calculator.getSpillEntries(sheet.id))
        cachedCalculationSnapshot.formulas.has(`${sheet.id}:${g.anchor}`) &&
          spills.push({
            ...g,
            sheetId: sheet.id,
            result: { ...pendingFormulaValue },
          });
  const calculator = cU(workbook, report, values, {
    spills: spills,
    spillMetadata: [],
    references: [],
    names: [],
    financial: [],
    operationalErrors: cachedCalculationSnapshot.formulas.size
      ? ["#PENDING!"]
      : [],
  });
  return (
    calculatorRecords.set(workbook, {
      ...cachedCalculationSnapshot,
      calculator: calculator,
      report: report,
      values: values,
      dependencies: dependencies,
    }),
    carriedCalculatorRecords.delete(workbook),
    { calculator: calculator, report: report }
  );
}
function acceptWorkerCalculationResult(workbook, report, t) {
  const n = calculatorRecords.get(workbook),
    cachedCalculationSnapshot = getCachedCalculationSnapshot(workbook);
  if (
    !n ||
    n.report !== report ||
    !isCalculationRecordCurrent(n, workbook, cachedCalculationSnapshot)
  )
    return;
  if (
    t.inputSignature !== getCalculationInputSignature(workbook) ||
    t.report.workbookId !== workbook.id ||
    t.report.revision !== workbook.revision ||
    t.report.cells.length !== cachedCalculationSnapshot.formulas.size
  )
    throw Error("The worker result does not match the requested workbook.");
  const values = new Map(),
    dependencies = new Map();
  for (const cell of t.report.cells) {
    const h = `${cell.sheetId}:${cell.address}`;
    if (
      values.has(h) ||
      cachedCalculationSnapshot.formulas.get(h)?.formula !== cell.formula
    )
      throw Error(
        "The worker result contains stale or duplicate formula inputs.",
      );
    (values.set(h, cell.result),
      dependencies.set(h, new Set(cell.dependencies)));
  }
  const d = { ...t.report, run: report.run };
  deepFreeze(d);
  const calculator = cU(workbook, d, values, t);
  return (
    calculatorRecords.set(workbook, {
      ...cachedCalculationSnapshot,
      calculator: calculator,
      report: d,
      values: values,
      dependencies: dependencies,
    }),
    { calculator: calculator, report: d }
  );
}
function runIterativeCalculation(workbook, report, t, n, wallElapsedMs) {
  if (t === "#PENDING!") {
    const s = calculatorRecords.get(workbook);
    if (!s || s.report !== report || !isCalculationRecordCurrent(s, workbook))
      return;
    const d = Object.freeze({
        mode: "worker",
        state: "cancelled",
        message: n,
        wallElapsedMs: wallElapsedMs,
      }),
      u = Object.freeze(
        Object.defineProperties(
          {},
          {
            ...Object.getOwnPropertyDescriptors(report),
            reason: { value: n, enumerable: true },
            execution: { value: d, enumerable: true },
            iteration: {
              value: Object.freeze({
                ...report.iteration,
                status: report.settings.iterative ? "failed" : "off",
                message: n,
                passes: 0,
                maxDelta: null,
              }),
              enumerable: true,
            },
          },
        ),
      ),
      l = reportCellValueMaps.get(report);
    return (
      l && reportCellValueMaps.set(u, l),
      calculatorRecords.set(workbook, { ...s, report: u }),
      { calculator: s.calculator, report: u }
    );
  }
  const cells = report.cells.map((cell) => ({
    ...cell,
    result: makeFormulaError(t),
    status: "error",
    evaluations: 0,
    dependencies: [],
  }));
  return acceptWorkerCalculationResult(workbook, report, {
    inputSignature: getCalculationInputSignature(workbook),
    report: {
      ...report,
      cells: cells,
      pending: 0,
      evaluated: 0,
      reused: 0,
      dependencyEdges: 0,
      cycles: [],
      reason: n,
      execution: {
        mode: "worker",
        state: t === "#TIMEOUT!" ? "timeout" : "failed",
        message: n,
        wallElapsedMs: wallElapsedMs,
      },
      limits: { ...report.limits, hardInterrupt: true },
      iteration: {
        ...report.iteration,
        status: report.settings.iterative ? "failed" : "off",
        message: n,
        passes: 0,
        maxDelta: null,
      },
    },
    spills: [],
    spillMetadata: [],
    references: [],
    names: [],
    financial: [],
    operationalErrors: [t],
  });
}
function recalculateWorkbookFormulas(
  workbook,
  cachedCalculationSnapshot,
  t,
  n,
) {
  const monotonicClockMs = readMonotonicClockMs(),
    { formulas, settings } = cachedCalculationSnapshot,
    localDateTime = resolveLocalDateTime(workbook.datePolicy),
    calculationBudgetTracker = createCalculationBudgetTracker(workbook);
  calculationBudgetTracker.begin();
  const l = new Set(),
    dependencies = new Map(),
    values = new Map(),
    f = new Map();
  let reason = n
    ? "Explicit full recalculation"
    : t
      ? "Changed inputs and transitive dependents"
      : "First calculation of this snapshot";
  if (
    n ||
    !t ||
    cachedCalculationSnapshot.metadata !== t.metadata ||
    ik(workbook, formulas) ||
    ik(workbook, t.formulas)
  )
    (formulas.forEach((formula, G) => l.add(G)),
      t &&
        !n &&
        (reason =
          cachedCalculationSnapshot.metadata !== t.metadata
            ? "Calculation settings or structural metadata changed"
            : "Conservative recalculation for names, dynamic references or arrays"));
  else {
    const U = new Set(),
      G = cachedCalculationSnapshot.inputs(),
      H = t.inputs();
    for (const J of new Set([...G.keys(), ...H.keys()]))
      G.get(J) !== H.get(J) && U.add(J);
    (formulas.forEach((formula, q) => {
      (_O.test(formula.formula) ||
        !t.values.has(q) ||
        ["#TIMEOUT!", "#LIMIT!"].includes(String(t.values.get(q)?.value))) &&
        U.add(q);
    }),
      t.report.cycles.flat().forEach((J) => U.add(J)));
    const Q = new Map();
    t.dependencies.forEach((dependency, q) =>
      dependency.forEach((W) => {
        const K = Q.get(W) ?? new Set();
        (K.add(q), Q.set(W, K));
      }),
    );
    const X = [...U];
    for (let J = 0; J < X.length; J++)
      for (const q of Q.get(X[J]) ?? []) U.has(q) || (U.add(q), X.push(q));
    U.forEach((J) => {
      formulas.has(J) && l.add(J);
    });
  }
  formulas.forEach((formula, G) => {
    !l.has(G) &&
      t?.values.has(G) &&
      (values.set(G, t.values.get(G)),
      dependencies.set(G, new Set(t.dependencies.get(G))));
  });
  const v = new Map(values),
    cycles = [],
    b = new Set();
  let y = true,
    w,
    E = [...dependencies.values()].reduce((U, G) => U + G.size, 0);
  const check = () => (
      y && !w && (w = calculationBudgetTracker.check()),
      y &&
        !w &&
        readMonotonicClockMs() - monotonicClockMs > settings.timeBudgetMs &&
        (w = "#TIMEOUT!"),
      y ? w : undefined
    ),
    k = (U, G) => {
      const H = [];
      return evaluateWorkbookFormulas(workbook, {}, undefined, {
        budget: calculationBudgetTracker,
        running: () => y,
        dateContext: localDateTime,
        check: check,
        read(X) {
          const J = H.at(-1);
          if (!y || !J) return;
          const q = dependencies.get(J);
          q.has(X) || (E >= j5 ? (w = "#LIMIT!") : (q.add(X), E++));
        },
        value(X, J) {
          return J && G?.has(X) ? G.get(X) : U.get(X);
        },
        enter(X) {
          (H.push(X),
            (E -= dependencies.get(X)?.size ?? 0),
            dependencies.set(X, new Set()),
            f.set(X, (f.get(X) ?? 0) + 1));
        },
        leave(X, J) {
          (values.set(X, J), H.pop());
        },
        cycle(X) {
          const J = X.at(-1),
            q = [...new Set(X.slice(X.indexOf(J), -1))];
          (q.forEach((W) => b.add(W)),
            q.length &&
              !cycles.some(
                (cycle) =>
                  cycle.length === q.length &&
                  cycle.every((K) => q.includes(K)),
              ) &&
              cycles.push(q));
        },
      });
    },
    F = settings.mode === "manual" && !n;
  F &&
    l.forEach((U) => {
      (values.set(U, pendingFormulaValue),
        v.set(U, pendingFormulaValue),
        dependencies.set(U, new Set(t?.dependencies.get(U))));
    });
  let I = k(v);
  const O = () => {
    for (const [U, G] of formulas)
      values.set(U, v.get(U) ?? I.getTypedValue(G.sheetId, G.address));
  };
  F || O();
  const iteration = {
    status: settings.iterative ? "not-needed" : "off",
    passes: 0,
    maxDelta: null,
    message: settings.iterative
      ? "No circular calculation was required."
      : "Iteration is disabled. Circular references return #CYCLE!.",
  };
  if (!F && b.size && settings.iterative) {
    const U = new Set(b),
      G = new Map();
    dependencies.forEach((dependency, X) =>
      dependency.forEach((J) => {
        const q = G.get(J) ?? new Set();
        (q.add(X), G.set(J, q));
      }),
    );
    const H = [...b];
    for (let Q = 0; Q < H.length; Q++)
      for (const X of G.get(H[Q]) ?? []) U.has(X) || (U.add(X), H.push(X));
    if (
      (values.forEach((value, X) => {
        U.has(X) || v.set(X, value);
      }),
      (iteration.status = "failed"),
      ik(workbook, formulas) ||
        [...formulas.values()].some((Q) => _O.test(Q.formula)) ||
        b.size > 1e3)
    )
      iteration.message =
        "Iteration supports at most 1,000 deterministic scalar cycle cells. Names, arrays, dynamic references and volatile formulas must be removed from this calculation.";
    else {
      let Q = new Map([...b].map((X) => [X, { type: "number", value: 0 }]));
      iteration.message = `No convergence within ${settings.maxIterations} simultaneous passes from zero.`;
      for (let X = 0; X < settings.maxIterations && !check(); X++) {
        I = k(v, Q);
        const J = new Map();
        let q = 0,
          W = true;
        for (const K of b) {
          const te = formulas.get(K),
            ae = I.getTypedValue(te.sheetId, te.address);
          if (
            (J.set(K, ae), ae.type !== "number" || !Number.isFinite(ae.value))
          ) {
            W = false;
            break;
          }
          q = Math.max(q, Math.abs(Number(ae.value) - Number(Q.get(K).value)));
        }
        if (
          ((iteration.passes = X + 1),
          (iteration.maxDelta = Number.isFinite(q) ? q : null),
          !W)
        ) {
          iteration.message =
            "The circular calculation produced a nonnumeric value or calculation error.";
          break;
        }
        if (((Q = J), q <= settings.maxChange)) {
          ((iteration.status = "converged"),
            (iteration.message = `Converged from zero in ${X + 1} simultaneous passes; maximum change ${q}.`));
          break;
        }
      }
      iteration.status === "converged" &&
        !w &&
        b.forEach((X) => v.set(X, Q.get(X)));
    }
    ((iteration.status !== "converged" || w) &&
      b.forEach((Q) => v.set(Q, makeFormulaError(w ?? "#ITER!"))),
      (I = k(v)),
      O());
  }
  (check(),
    w &&
      (l.forEach((U) => values.set(U, makeFormulaError(w))),
      iteration.status === "converged" && (iteration.status = "failed"),
      b.size &&
        settings.iterative &&
        (iteration.message = `Calculation stopped with ${w}; no intermediate result is current.`)),
    (y = false),
    calculationBudgetTracker.end());
  const calculator = (U, G) => calculator.getTypedValue(U, G).value;
  Object.assign(calculator, I);
  const V = (U, G) => `${U}:${G.replace(/\$/g, "").toUpperCase()}`,
    D = new Map();
  if (F && t)
    for (const sheet of workbook.sheets)
      for (const G of t.calculator.getSpillEntries(sheet.id))
        l.has(`${sheet.id}:${G.anchor}`) &&
          D.set(`${sheet.id}:${G.address}`, G.anchor);
  calculator.getTypedValue = (U, G) => {
    const H = V(U, G);
    if (values.has(H)) return { ...values.get(H) };
    const cell = workbook.sheets.find((sheet) => sheet.id === U)?.cells[
      G.replace(/\$/g, "").toUpperCase()
    ];
    if (
      cell &&
      cell.raw !== null &&
      (cell.literal || typeof cell.raw != "string" || !cell.raw.startsWith("="))
    )
      return {
        type:
          typeof cell.raw == "string"
            ? "text"
            : typeof cell.raw == "boolean"
              ? "boolean"
              : "number",
        value: cell.raw,
      };
    if (D.has(H)) return { ...pendingFormulaValue };
    const X = I.getSpillAnchor(U, G);
    return X && values.get(`${U}:${X}`)?.type === "error"
      ? { ...values.get(`${U}:${X}`) }
      : I.getTypedValue(U, G);
  };
  const getSpillEntries = I.getSpillEntries;
  ((calculator.getSpillEntries = (U) => [
    ...getSpillEntries(U),
    ...[...D]
      .filter(([G]) => G.startsWith(`${U}:`))
      .map(([G, anchor]) => ({
        address: G.slice(U.length + 1),
        anchor: anchor,
      })),
  ]),
    (calculator.getOperationalErrors = () => [
      ...new Set([
        ...I.getOperationalErrors(),
        ...[...values.values()]
          .filter(
            (U) =>
              U.type === "error" &&
              [
                "#PENDING!",
                "#TIMEOUT!",
                "#ITER!",
                "#LIMIT!",
                "#BLOCKED!",
                "#CYCLE!",
                "#ERROR!",
              ].includes(String(U.value)),
          )
          .map((U) => U.value),
      ]),
    ]));
  const _ = (U, G) => [
      ...U.filter(
        (H) =>
          !l.has(`${H.sheetId}:${H.address}`) &&
          formulas.has(`${H.sheetId}:${H.address}`),
      ),
      ...G,
    ],
    C = t?.calculator.getReferenceDiagnostics() ?? [],
    T = t?.calculator.getNamedFormulaDiagnostics() ?? [],
    N = t?.calculator.getFinancialDiagnostics() ?? [];
  ((calculator.getFinancialDiagnostics = () =>
    structuredClone(_(N, I.getFinancialDiagnostics()))),
    (calculator.getReferenceDiagnostics = () =>
      _(C, I.getReferenceDiagnostics())),
    (calculator.getNamedFormulaDiagnostics = () =>
      _(T, I.getNamedFormulaDiagnostics())));
  const R = new Map(
      t?.report.cells.map((cell) => [`${cell.sheetId}:${cell.address}`, cell]),
    ),
    L = new Map();
  for (const {
    sheetId,
    address,
    ...H
  } of calculator.getFinancialDiagnostics()) {
    const Q = `${sheetId}:${address}`;
    L.set(Q, [...(L.get(Q) ?? []), H]);
  }
  const cells = [...formulas].map(([U, G]) => {
    const H = values.get(U) ?? pendingFormulaValue,
      Q = t?.values.get(U);
    return {
      ...G,
      ...(H.value !== "#PENDING!" && L.has(U) ? { financial: L.get(U) } : {}),
      result: { ...H },
      status:
        H.type === "error"
          ? H.value === "#PENDING!"
            ? "pending"
            : "error"
          : "current",
      ...(H.value === "#PENDING!" && Q && Q.value !== "#PENDING!"
        ? { lastKnown: { revision: t.report.revision, result: { ...Q } } }
        : {}),
      ...(H.value === "#PENDING!" &&
      Q?.value === "#PENDING!" &&
      R.get(U)?.lastKnown
        ? { lastKnown: R.get(U).lastKnown }
        : {}),
      dependencies: [...(dependencies.get(U) ?? [])].sort(),
      evaluations: f.get(U) ?? 0,
    };
  });
  return {
    ...cachedCalculationSnapshot,
    calculator: calculator,
    values: values,
    dependencies: dependencies,
    report: {
      resources: calculationBudgetTracker.report(),
      volatile: {
        functions: [
          ...new Set(
            [...formulas.values()]
              .flatMap((U) => findVolatileFunctionNames(U.formula))
              .concat(
                (workbook.definedNames ?? []).flatMap((U) =>
                  findVolatileFunctionNames(U.refersTo),
                ),
              ),
          ),
        ].sort(),
        policy: calculationFeatureSupport,
      },
      execution: {
        mode: "cooperative",
        state:
          w === "#TIMEOUT!"
            ? "timeout"
            : w
              ? "failed"
              : F
                ? "pending"
                : "complete",
        message: w
          ? (calculationBudgetTracker.report().reason ??
            `Calculation stopped with ${w}.`)
          : "Synchronous calculation uses cooperative limits.",
      },
      dateContext: localDateTime,
      workbookId: workbook.id,
      revision: workbook.revision,
      run: (t?.report.run ?? 0) + 1,
      settings,
      reason: reason,
      formulaCount: formulas.size,
      evaluated: cells.filter((cell) => cell.evaluations > 0).length,
      reused: formulas.size - l.size,
      pending: cells.filter((cell) => cell.status === "pending").length,
      elapsedMs: Math.max(0, readMonotonicClockMs() - monotonicClockMs),
      dependencyEdges: [...dependencies.values()].reduce(
        (U, G) => U + G.size,
        0,
      ),
      cycles: cycles,
      iteration: iteration,
      cells: cells,
      limits: {
        dependencyEdges: j5,
        cooperativeTimeBudgetMs: settings.timeBudgetMs,
        hardInterrupt: false,
      },
    },
  };
}
function getWorkbookCalculation(workbook) {
  const cachedCalculationSnapshot = getCachedCalculationSnapshot(workbook),
    t = calculatorRecords.get(workbook);
  if (isCalculationRecordCurrent(t, workbook, cachedCalculationSnapshot))
    return { calculator: t.calculator, report: t.report };
  const n = recalculateWorkbookFormulas(
    workbook,
    cachedCalculationSnapshot,
    t ? undefined : carriedCalculatorRecords.get(workbook),
    false,
  );
  return (
    deepFreeze(n.report),
    calculatorRecords.set(workbook, n),
    carriedCalculatorRecords.delete(workbook),
    { calculator: n.calculator, report: n.report }
  );
}
function calculateWorkbookWithSettings(workbook, r = {}, t) {
  return !t &&
    !Object.keys(r).length &&
    (workbook.calculation ||
      workbook.datePolicy ||
      hasCalculatorRecord(workbook) ||
      workbookUsesNowOrToday(workbook) ||
      workbookHasVolatileFunctions(workbook))
    ? getWorkbookCalculation(workbook).calculator
    : evaluateWorkbookFormulas(workbook, r, t);
}
function updateConditionalFormatRules(workbook, sheetId, rules, n = {}) {
  if (n.readOnly) throw Error("This workbook is read-only.");
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw Error("Worksheet no longer exists.");
  if (sheet.protection?.enabled)
    throw Error("Unprotect the worksheet before changing conditional formats.");
  const conditionalRules = normalizeConditionalRules(rules),
    s = sheet.conditionalFormats ?? [],
    changes = [];
  for (const m of new Set([
    ...s.map((v) => v.id),
    ...conditionalRules.map((conditionalRule) => conditionalRule.id),
  ])) {
    const beforeIndex = s.findIndex((w) => w.id === m),
      afterIndex = conditionalRules.findIndex(
        (conditionalRule) => conditionalRule.id === m,
      ),
      b = s[beforeIndex],
      y = conditionalRules[afterIndex];
    (JSON.stringify(b) !== JSON.stringify(y) || beforeIndex !== afterIndex) &&
      changes.push({
        id: m,
        ...(b ? { before: structuredClone(b), beforeIndex: beforeIndex } : {}),
        ...(y ? { after: structuredClone(y), afterIndex: afterIndex } : {}),
      });
  }
  const workbookWithSettings = calculateWorkbookWithSettings(workbook),
    l = evaluateConditionalFormatRules(sheet, workbookWithSettings),
    c = evaluateConditionalFormatRules(
      { ...sheet, conditionalFormats: conditionalRules },
      workbookWithSettings,
    ),
    h = new Set();
  for (const m of [...conditionalRules, ...s]) {
    const rangeAddress = parseRangeAddress(m.range);
    if (
      (h.add(
        `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
      ),
      h.size >= 50)
    )
      break;
  }
  if (h.size < 50) {
    for (const cell of Object.keys(sheet.cells))
      if (
        ([...conditionalRules, ...s].some((v) =>
          isCellInsideRange(parseRangeAddress(v.range), _t(cell)),
        ) && h.add(cell),
        h.size >= 50)
      )
        break;
  }
  const samples = [...h].map((address) => ({
    address: address,
    value: workbookWithSettings(sheetId, address),
    before: l(address),
    after: c(address),
  }));
  return {
    workbookId: workbook.id,
    revision: workbook.revision,
    snapshotGuard: JSON.stringify(workbook),
    sheetId: sheetId,
    rules: conditionalRules,
    changes: changes,
    samples: samples,
    errors: c.errors(),
    previewLimit: 50,
  };
}
function refreshCalculationReport(workbook) {
  const cachedCalculationSnapshot = getCachedCalculationSnapshot(workbook),
    t = calculatorRecords.get(workbook),
    n = recalculateWorkbookFormulas(
      workbook,
      cachedCalculationSnapshot,
      isCalculationRecordCurrent(t, workbook, cachedCalculationSnapshot)
        ? t
        : undefined,
      true,
    );
  return (
    deepFreeze(n.report),
    calculatorRecords.set(workbook, n),
    carriedCalculatorRecords.delete(workbook),
    { calculator: n.calculator, report: n.report }
  );
}
function buildCalculationEvidence(workbook, report) {
  const workbookCalculation = getWorkbookCalculation(workbook);
  if (workbookCalculation.report !== report)
    throw Error(
      "The workbook or calculation run changed. Review current calculation results before exporting evidence.",
    );
  const importedCalculationCaches = workbook.importedCalculationCaches,
    entries = (importedCalculationCaches?.entries ?? []).map((a) => {
      const sheet = workbook.sheets.find((sheet) => sheet.id === a.sheetId),
        cell = sheet?.cells[a.address],
        u = sheet
          ? workbookCalculation.calculator.getTypedValue(a.sheetId, a.address)
          : undefined,
        l =
          a.cache &&
          u &&
          a.cache.type === u.type &&
          (typeof a.cache.value == "number" && typeof u.value == "number"
            ? Math.abs(a.cache.value - u.value) <=
              1e-10 * Math.max(1, Math.abs(a.cache.value), Math.abs(u.value))
            : a.cache.value === u.value),
        c = sheet
          ? cell?.literal
            ? "neutralized-or-literal"
            : cell?.raw !== a.formula
              ? "formula-changed"
              : u?.type === "error" && u.value === "#PENDING!"
                ? "pending"
                : a.cache
                  ? l
                    ? "matched"
                    : "different"
                  : "missing"
          : "source-sheet-removed";
      return {
        ...a,
        ...(u ? { current: u } : {}),
        status: c,
        verifiedByExcel: false,
      };
    });
  return JSON.stringify(
    {
      schemaVersion: 1,
      report: report,
      imported: importedCalculationCaches
        ? { ...importedCalculationCaches, entries: entries }
        : null,
      boundaries: [
        "Imported caches are unverified source evidence, never current-value fallbacks.",
        "Matching caches do not prove Excel parity or source-file authenticity.",
        "Calculation timings and iteration counts are local runtime evidence.",
        report.limits.hardInterrupt
          ? "This calculation ran in a terminable worker. Timer delivery is subject to runtime scheduling."
          : "This synchronous calculation used cooperative time checks, not a hard worker interrupt.",
        "Working-data accounting is a bounded allocation allowance, not an exact heap/RSS measurement; runtime overhead is outside it.",
        "Diagnostic formula evaluation is separate from the committed revision calculation.",
      ],
    },
    null,
    2,
  );
}
function parseReportRange(range) {
  if (
    !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?$/.test(
      range,
    )
  )
    throw Error("Choose one explicit report range.");
  const r = range.replaceAll("$", "").toUpperCase().split(":"),
    rangeAddress = parseRangeAddress(range);
  if (
    r[0] !==
      `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}` ||
    (r[1] &&
      r[1] !==
        `${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`)
  )
    throw Error("Report ranges cannot be reversed.");
  if (
    rangeAddress.lastCol - rangeAddress.firstCol + 1 > 20 ||
    rangeAddress.lastRow - rangeAddress.firstRow + 1 > 500 ||
    (rangeAddress.lastCol - rangeAddress.firstCol + 1) *
      (rangeAddress.lastRow - rangeAddress.firstRow + 1) >
      2e3
  )
    throw Error(
      "Reports support at most 500 rows, 20 columns and 2,000 cells.",
    );
  return `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
}
const PO = object({
    column: number().int().min(0).max(16383),
    type: _enum(["value", "text", "number", "date"]),
    operator: _enum([
      "equals",
      "notEquals",
      "contains",
      "notContains",
      "startsWith",
      "endsWith",
      "wildcard",
      "gt",
      "gte",
      "lt",
      "lte",
      "blank",
      "notBlank",
    ]),
    value: union([
      string().max(1024),
      number().finite(),
      boolean(),
      _null(),
    ]).optional(),
  }).strict(),
  filterRequestSchema = object({
    range: createBoundedTextSchema(40),
    columns: array(
      object({
        column: number().int().min(0).max(16383),
        values: array(
          union([
            createBoundedTextSchema(32767),
            number().finite(),
            boolean(),
            _null(),
          ]),
        )
          .max(1e3)
          .optional(),
        conditions: array(PO).max(128).optional(),
        join: _enum(["and", "or"]).optional(),
      }).strict(),
    ).max(20),
    advanced: array(array(PO).max(128)).max(128).optional(),
    caseSensitive: boolean().optional(),
    unique: boolean().optional(),
  }).strict(),
  kU = object({
    name: string().trim().min(1).max(100),
    kind: _enum(["range", "object"]),
    sheetId: nonEmptyLabelSchema,
    range: createBoundedTextSchema(40).transform(parseReportRange),
    objectId: nonEmptyLabelSchema.optional(),
    mode: _enum(["snapshot", "linked"]),
    population: _enum(["full", "visible"]),
    filter: filterRequestSchema.optional(),
  }).strict(),
  cellSnapshotSchema = object({
    address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
    value: union([
      createBoundedTextSchema(32767),
      number().finite(),
      boolean(),
      _null(),
    ]),
    text: createBoundedTextSchema(32767),
    style: unknown().optional(),
  }).strict(),
  calculationReportSchema = object({
    engine: literal(reportOutputFormatVersion),
    id: nonEmptyLabelSchema,
    request: kU,
    workbookId: nonEmptyLabelSchema,
    workbookName: createBoundedTextSchema(200),
    revision: number().int().nonnegative(),
    sheetName: createBoundedTextSchema(31),
    calculatedAt: AU,
    sourceHash: reportHashSchema,
    includeHidden: boolean(),
    authorizedSheetIds: array(nonEmptyLabelSchema).max(1e3),
    sourceCells: array(
      object({
        sheetId: nonEmptyLabelSchema,
        address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
      }).strict(),
    ).max(4e3),
    rows: array(
      object({
        row: number().int().min(1).max(1048576),
        cells: array(cellSnapshotSchema).max(20),
      }).strict(),
    ).max(500),
    columns: array(number().int().min(0).max(16383)).max(20),
    widths: array(number().finite().min(40).max(600)).max(20),
    excludedRows: array(number().int().min(1).max(1048576)).max(500),
    object: unknown().optional(),
    warnings: array(createBoundedTextSchema(2e3)).max(50),
    hash: reportHashSchema,
  }).strict(),
  importedCacheEvidenceSchema = object({
    id: nonEmptyLabelSchema,
    fileName: createBoundedTextSchema(255),
    fileBytes: number()
      .int()
      .min(1)
      .max(10 * 1024 * 1024),
    sourceSha256: reportHashSchema,
    inspectedAt: AU,
    execution: literal("never"),
    items: array(
      object({
        path: createBoundedTextSchema(4096),
        kind: createBoundedTextSchema(100),
        bytes: number()
          .int()
          .nonnegative()
          .max(50 * 1024 * 1024),
        sha256: reportHashSchema.optional(),
        execution: literal("never"),
        note: createBoundedTextSchema(2e3),
      }).strict(),
    )
      .min(1)
      .max(1e3),
    warnings: array(createBoundedTextSchema(2e3)).max(100),
  }).strict();
function validateObjectReportRequest(request) {
  const r = kU.parse(request);
  if ((r.kind === "object") != !!r.objectId)
    throw Error("Select an existing drawing for an object report.");
  if (r.filter && parseReportRange(r.filter.range) !== r.range)
    throw Error("The report range must match its filter range.");
  return r;
}

export {
  validateObjectReportRequest,
  evaluateWorkbookFormulas,
  parseReportRange,
  getWorkbookCalculation,
  getIsolatedCalculation,
  updateConditionalFormatRules,
  summarizeValidationInspection,
  importedCacheEvidenceSchema,
  calculationReportSchema,
  calculateWorkbookWithSettings,
  buildCalculationEvidence,
  runIterativeCalculation,
  refreshCalculationReport,
  acceptWorkerCalculationResult,
  checkDataValidationBatch,
};

/* Also used by the workers. */
export {
  PO,
  cellSnapshotSchema,
  filterRequestSchema,
  kU,
  recalculateWorkbookFormulas,
};

/* Also used by the workers. */
export { evaluateValidationRule };
