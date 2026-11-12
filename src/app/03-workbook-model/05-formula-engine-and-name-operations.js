import { flattenNestedValues } from "../02-text-and-financial-functions.js";
import { isPlainObjectValue } from "./01-cell-styles-and-themes.js";
import { isNullOrEmptyValue } from "./03-defined-names.js";
import {
  evaluateDateFunction,
  notDateFunctionSentinel,
  resolveLocalDateTime,
  dateSystemSerialOffset,
  parseDateText,
  findVolatileFunctionNames,
  isDeepFrozenValue,
  getExtensionCacheEntries,
  normalizeDatePolicy,
  dateSystemMaxSerial,
  fingerprintCanonicalValue,
} from "../01-schemas-date-functions-and-validation.js";
import {
  fileImportLimits,
  evaluateFinancialFunction,
  notFinancialSentinel,
  evaluateTextFunction,
  notTextFunctionSentinel,
  broadcastOverArrays,
} from "../02-text-and-financial-functions.js";
import {
  string,
  number,
  boolean,
  array,
  strictObject,
  discriminatedUnion,
  literal,
  union,
  enum as _enum,
  null as _null,
} from "zod";
import {
  validateNumberFormatCode,
  numberFormatToIntlOptions,
} from "./01-cell-styles-and-themes.js";
import {
  notRegressionSentinel,
  notArithmeticSentinel,
  notSpecialistSentinel,
  notInferenceSentinel,
  evaluateArithmeticFunction,
  notStatisticalSentinel,
  sumDecimalPartsToNumber,
  exactAverageOfNumbers,
  evaluateSignificanceTestFunction,
  notOwnedLookupSentinel,
  evaluateSpecialistFunction,
  evaluateLookupFunction,
  evaluateRegressionFunction,
  evaluateLogicalFunction,
  notModernFunctionSentinel,
  evaluateStatisticalFunction,
  notOwnedLogicalSentinel,
  formatCellValueWithNumberFormat,
  matrixResultFunctionNames,
} from "./02-formula-function-evaluators.js";
import {
  definedNameScopeKey,
  evaluateDefinedNameFunction,
  validateDefinedName,
  replaceTextSpans,
  m5,
  pO,
} from "./03-defined-names.js";
import {
  findFormulaErrorArg,
  builtinFunctionNameSet,
  evaluateDatabaseFunction,
  findScopedDefinedName,
  formulaErrorFactory,
  createFormulaError,
  notDatabaseFunctionMarker,
  notReferenceFunctionMarker,
  resolveDefinedNameValue,
  isFormulaError,
} from "./04-pivot-tables-and-pictures.js";

function applyDefinedNameOperation(workbook, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (workbook.structureProtection?.enabled)
    throw Error(
      "Workbook structure is protected. Unprotect it before changing names.",
    );
  const n = validateDefinedName(workbook.definedNames, workbook),
    o =
      r.kind === "create"
        ? undefined
        : n.find(
            (y) =>
              y.name.toLowerCase() === r.name.toLowerCase() &&
              y.scopeSheetId === r.scopeSheetId,
          );
  if (r.kind !== "create" && !o)
    throw Error("The defined name was not found. Reopen Name Manager.");
  let a;
  if (r.kind === "create" || r.kind === "update")
    a = validateDefinedName([r.definition], workbook)[0];
  else if (r.kind === "rename")
    a = validateDefinedName([{ ...o, name: r.newName }], workbook)[0];
  else if (r.kind !== "delete") throw Error("Unknown name operation.");
  if (
    r.kind === "update" &&
    (o.name !== a.name || o.scopeSheetId !== a.scopeSheetId)
  )
    throw Error(
      "Use Rename to change a name. Existing name scope cannot be changed; create a separate name instead.",
    );
  if (
    a &&
    workbook.sheets.some((sheet) =>
      sheet.tables?.some(
        (table) => table.name.toLowerCase() === a.name.toLowerCase(),
      ),
    )
  )
    throw Error("A table already uses this name.");
  const next = structuredClone(workbook),
    d = o ? definedNameScopeKey(o) : undefined;
  ((next.definedNames = validateDefinedName(
    [...n.filter((y) => definedNameScopeKey(y) !== d), ...(a ? [a] : [])],
    workbook,
  )),
    o &&
      a &&
      (next.definedNames = n.map((y) =>
        definedNameScopeKey(y) === d ? a : y,
      )));
  const u = a?.scopeSheetId ?? o?.scopeSheetId,
    l = (y) => {
      if (workbook.sheets.find((sheet) => sheet.id === y)?.protection?.enabled)
        throw Error(
          "An affected worksheet is protected. Unprotect it before changing names.",
        );
    };
  u && l(u);
  const changes = [],
    dependencies = [];
  let f = 0;
  const p = true,
    m = (y) => {
      const w = m5(y, p);
      if (((f += w.length + 1), f > 1e6))
        throw Error(
          "The name dependency scan exceeds the supported budget. Split the workbook before changing names.",
        );
      if (
        w.some(
          (E) =>
            E.qualifier &&
            !workbook.sheets.some(
              (sheet) => sheet.name.toLowerCase() === E.qualifier.toLowerCase(),
            ),
        )
      )
        throw Error(
          "An unknown or external worksheet qualifier cannot be safely migrated.",
        );
      return w;
    },
    v = (y, w, E = []) => {
      if (E.length > 32)
        throw Error("The name dependency chain exceeds 32 levels.");
      for (const A of m(y)) {
        const scopedDefinedName = findScopedDefinedName(A, workbook, w),
          F =
            r.kind === "create" ? findScopedDefinedName(A, next, w) : undefined;
        if (
          F &&
          definedNameScopeKey(F) === definedNameScopeKey(a) &&
          (!scopedDefinedName ||
            definedNameScopeKey(scopedDefinedName) !== definedNameScopeKey(F))
        )
          return true;
        if (!scopedDefinedName) {
          if (
            r.kind === "create" &&
            findScopedDefinedName(A, next, w) &&
            definedNameScopeKey(findScopedDefinedName(A, next, w)) ===
              definedNameScopeKey(a)
          )
            return true;
          continue;
        }
        if (definedNameScopeKey(scopedDefinedName) === d) return true;
        const I = definedNameScopeKey(scopedDefinedName);
        if (
          !E.includes(I) &&
          v(
            scopedDefinedName.refersTo,
            scopedDefinedName.scopeSheetId ??
              (A.qualifier
                ? workbook.sheets.find(
                    (sheet) =>
                      sheet.name.toLowerCase() === A.qualifier.toLowerCase(),
                  ).id
                : w),
            [...E, I],
          )
        )
          return true;
      }
      return false;
    },
    g = (before, w, E) => {
      const A = m(before);
      if (w.some((I) => v(before, I))) {
        if (
          (dependencies.push({
            label: E.label,
            ...(E.sheetId ? { sheetId: E.sheetId } : {}),
            ...(E.address ? { address: E.address } : {}),
          }),
          dependencies.length > pO)
        )
          throw Error(
            "A name operation may affect at most 500 dependent locations. Split the operation into smaller workbooks.",
          );
        if ((E.sheetId && l(E.sheetId), r.kind === "delete"))
          throw Error(
            `This name is referenced by ${E.label}. Remove or replace its dependencies before deletion.`,
          );
      }
      const after = replaceTextSpans(before, A, (I) => {
        if (r.kind !== "rename") return I.text;
        const O = w.map((V) => findScopedDefinedName(I, workbook, V)),
          j = O.map((V) => !!V && definedNameScopeKey(V) === d);
        if (j.some(Boolean) && !j.every(Boolean))
          throw Error(
            `The workbook-scoped definition ${E.label} resolves differently across worksheets. Qualify its references before renaming.`,
          );
        if (j.every(Boolean)) {
          if (!I.qualifier && I.bound?.has(a.name.toLowerCase()))
            throw Error(
              `Renaming would be shadowed by a local LET variable or LAMBDA parameter in ${E.label}. Choose a distinct name.`,
            );
          const V = { ...I, name: a.name };
          if (
            w.some((D) => {
              const scopedDefinedName = findScopedDefinedName(V, next, D);
              return (
                !scopedDefinedName ||
                definedNameScopeKey(scopedDefinedName) !==
                  definedNameScopeKey(a)
              );
            })
          )
            throw Error(
              `Renaming would be shadowed by another name in ${E.label}. Choose a distinct name.`,
            );
          return I.qualifier
            ? I.text.slice(0, I.text.length - I.name.length) + a.name
            : a.name;
        }
        if (
          w.some((V, D) => {
            const scopedDefinedName = findScopedDefinedName(I, next, V);
            return (
              (O[D] ? definedNameScopeKey(O[D]) : undefined) !==
              (scopedDefinedName
                ? definedNameScopeKey(scopedDefinedName)
                : undefined)
            );
          })
        )
          throw Error(
            `Renaming would change an existing reference in ${E.label}. Choose a distinct name.`,
          );
        return I.text;
      });
      if (
        after !== before &&
        (E.sheetId && l(E.sheetId),
        changes.push({ ...E, before: before, after: after }),
        changes.length > pO)
      )
        throw Error(
          "A name operation may rewrite at most 500 formulas or definitions. Split the migration into smaller workbooks.",
        );
      return after;
    };
  for (const sheet of next.sheets) {
    for (const w of sheet.tables ?? [])
      for (const column of w.columns)
        column.calculatedFormula &&= g(column.calculatedFormula, [sheet.id], {
          kind: "table-column",
          label: `${sheet.name}: ${w.name}[${column.name}] calculated rule`,
          sheetId: sheet.id,
        });
    for (const [address, cell] of Object.entries(sheet.cells))
      typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal &&
        (cell.raw = g(cell.raw, [sheet.id], {
          kind: "cell",
          label: `${sheet.name}!${address}`,
          sheetId: sheet.id,
          address: address,
        }));
    for (const w of [
      ...(sheet.conditionalFormats ?? []),
      ...(sheet.validations ?? []),
    ])
      (w.kind === "formula" || w.kind === "custom") &&
        w.formula &&
        (w.formula = g(w.formula, [sheet.id], {
          kind: w.kind === "custom" ? "validation" : "conditional",
          label: `${sheet.name}: ${w.kind === "custom" ? "validation" : "conditional"} rule ${w.id}`,
          sheetId: sheet.id,
        }));
  }
  for (const definedName of next.definedNames) {
    const w = definedName.scopeSheetId
      ? [definedName.scopeSheetId]
      : workbook.sheets.map((sheet) => sheet.id);
    definedName.refersTo = g(definedName.refersTo, w, {
      kind: "name",
      label: `Name ${definedName.name} (${definedName.scopeSheetId ? workbook.sheets.find((sheet) => sheet.id === definedName.scopeSheetId).name : "Workbook"})`,
      ...(definedName.scopeSheetId
        ? { sheetId: definedName.scopeSheetId }
        : {}),
    });
  }
  if (a)
    for (const y of a.scopeSheetId
      ? [a.scopeSheetId]
      : workbook.sheets.map((sheet) => sheet.id))
      try {
        resolveDefinedNameValue(
          next.definedNames
            .find(
              (definedName) =>
                definedNameScopeKey(definedName) === definedNameScopeKey(a),
            )
            .refersTo.slice(1),
          next,
          y,
          [definedNameScopeKey(a)],
        );
      } catch (error) {
        throw Error(
          `The resulting name has a cycle or exceeds the expansion limit: ${error instanceof Error ? error.message : String(error)}`,
        );
      }
  return {
    snapshot: JSON.stringify(workbook),
    request: structuredClone(r),
    ...(o ? { before: structuredClone(o) } : {}),
    ...(a
      ? {
          after: structuredClone(
            next.definedNames.find(
              (definedName) =>
                definedNameScopeKey(definedName) === definedNameScopeKey(a),
            ),
          ),
        }
      : {}),
    changes: changes,
    dependencies: dependencies,
    next: next,
  };
}
function assertNamePreviewCurrent(workbook, r, t = {}) {
  if (JSON.stringify(workbook) !== r.snapshot)
    throw Error(
      "The workbook changed after the name preview. Review the operation again.",
    );
  const n = applyDefinedNameOperation(workbook, r.request, t);
  if (JSON.stringify(n) !== JSON.stringify(r))
    throw Error("The name preview changed. Review the operation again.");
  return JSON.stringify(n.next) === JSON.stringify(workbook)
    ? workbook
    : n.next;
}
function matchCriteriaValue(e, r) {
  if (typeof r != "string") return e === r;
  const t = /^(<=|>=|<>|=|<|>)?(.*)$/.exec(r),
    n = t[1] ?? "=",
    o = t[2];
  if ((n === "=" || n === "<>") && /[?*~]/.test(o)) {
    if (typeof e != "string") return n === "<>";
    const u = [];
    for (let h = 0; h < o.length; h++)
      o[h] === "~" && h + 1 < o.length
        ? u.push({ char: o[++h].toLowerCase(), wildcard: false })
        : u.push({
            char: o[h].toLowerCase(),
            wildcard: o[h] === "*" || o[h] === "?",
          });
    const l = String(e ?? "").toLowerCase();
    if (u.length * l.length > 1e5) throw formulaErrorFactory("#LIMIT!");
    let c = Array(l.length + 1).fill(false);
    c[0] = true;
    for (const h of u) {
      const f = Array(l.length + 1).fill(false);
      f[0] = h.wildcard && h.char === "*" && c[0];
      for (let p = 1; p <= l.length; p++)
        f[p] =
          h.wildcard && h.char === "*"
            ? c[p] || f[p - 1]
            : c[p - 1] &&
              ((h.wildcard && h.char === "?") || l[p - 1] === h.char);
      c = f;
    }
    return n === "<>" ? !c[l.length] : c[l.length];
  }
  const a = o.trim() !== "" && Number.isFinite(Number(o)),
    s = a
      ? e === null || e === ""
        ? 0
        : typeof e == "number"
          ? e
          : NaN
      : String(e ?? "").toLowerCase(),
    d = a ? Number(o) : o.toLowerCase();
  return n === "<>"
    ? s !== d
    : n === "<"
      ? s < d
      : n === ">"
        ? s > d
        : n === "<="
          ? s <= d
          : n === ">="
            ? s >= d
            : s === d;
}
function createWorkbookFormulaEngine(
  e,
  r,
  issue = createFormulaError,
  clock = resolveLocalDateTime(),
  report,
  criterionFormula,
) {
  const s = e.utils._applyInfix.bind(e.utils);
  ((e.utils._applyInfix = (u, l, c) => {
    const h = (val, E) => (
      ["=", "<>", "<", ">", "<=", ">="].includes(l) &&
        typeof val == "string" &&
        typeof E == "string" &&
        ((val = val.toLowerCase()), (E = E.toLowerCase())),
      s({ val: val, isArray: false }, l, { val: E, isArray: false })
    );
    if (!Array.isArray(u.val) && !Array.isArray(c.val)) return h(u.val, c.val);
    const f = (w) =>
        Array.isArray(w) ? (w.every(Array.isArray) ? w : [w]) : [[w]],
      p = f(u.val),
      m = f(c.val),
      v = Math.max(p.length, m.length),
      g = Math.max(p[0]?.length ?? 0, m[0]?.length ?? 0);
    return !v ||
      !g ||
      (p.length !== v && p.length !== 1) ||
      (m.length !== v && m.length !== 1) ||
      (p[0].length !== g && p[0].length !== 1) ||
      (m[0].length !== g && m[0].length !== 1)
      ? issue("#VALUE!")
      : v * g > fileImportLimits.maxSpillCells
        ? issue("#LIMIT!")
        : Array.from({ length: v }, (w, E) =>
            Array.from({ length: g }, (A, k) =>
              h(
                p[p.length === 1 ? 0 : E][p[0].length === 1 ? 0 : k],
                m[m.length === 1 ? 0 : E][m[0].length === 1 ? 0 : k],
              ),
            ),
          );
  }),
    (e.checkFormulaResult = (u, l) => {
      const c = e.retrieveRef(u);
      return isFormulaError(c)
        ? c
        : Array.isArray(c)
          ? l
            ? c
            : issue("#VALUE!")
          : c === null
            ? 0
            : typeof c == "number"
              ? Number.isFinite(c)
                ? c
                : issue("#NUM!")
              : typeof c == "string" || typeof c == "boolean"
                ? c
                : issue("#VALUE!");
    }));
  const distribution = e._callFunction.bind(e);
  return (
    (e._callFunction = (u, l) => {
      u = u
        .replace(/^_xlfn\./i, "")
        .replace(/^_xlws\./i, "")
        .toUpperCase();
      const retrieve = (f) => (f == null ? undefined : e.retrieveRef(f));
      try {
        if (!builtinFunctionNameSet.has(u)) return issue("#NAME?");
        if (u === "RAND") return l.length ? issue("#VALUE!") : Math.random();
        if (u === "RANDBETWEEN") {
          if (l.length !== 2) return issue("#VALUE!");
          const D = l.map(retrieve),
            formulaErrorArg = findFormulaErrorArg(D);
          if (formulaErrorArg) return formulaErrorArg;
          if (
            D.some(
              (T) =>
                Array.isArray(T) ||
                T === undefined ||
                (typeof T == "string" && T.trim() === "") ||
                !Number.isFinite(Number(T)),
            )
          )
            return issue("#VALUE!");
          const _ = Math.ceil(Number(D[0])),
            C = Math.floor(Number(D[1]));
          return !Number.isSafeInteger(_) ||
            !Number.isSafeInteger(C) ||
            _ > C ||
            !Number.isSafeInteger(C - _ + 1)
            ? issue("#NUM!")
            : _ + Math.floor(Math.random() * (C - _ + 1));
        }
        const f = evaluateDateFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          clock: clock,
        });
        if (f !== notDateFunctionSentinel) return f;
        const p = evaluateFinancialFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          clock: clock,
          report: report,
        });
        if (p !== notFinancialSentinel) return p;
        const m = r(u, l, e);
        if (m !== notReferenceFunctionMarker) return m;
        const v = evaluateDatabaseFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          criterionFormula: criterionFormula,
        });
        if (v !== notDatabaseFunctionMarker) return v;
        const g = evaluateSpecialistFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
        });
        if (g !== notSpecialistSentinel) return g;
        const b = evaluateRegressionFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
        });
        if (b !== notRegressionSentinel) return b;
        const y = evaluateSignificanceTestFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          distribution: distribution,
        });
        if (y !== notInferenceSentinel) return y;
        const w = evaluateArithmeticFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
        });
        if (w !== notArithmeticSentinel) return w;
        const E = evaluateStatisticalFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
        });
        if (E !== notStatisticalSentinel) return E;
        const A = evaluateLogicalFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
        });
        if (A !== notOwnedLogicalSentinel) return A;
        const k = evaluateLookupFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          matches: matchCriteriaValue,
        });
        if (k !== notOwnedLookupSentinel) return k;
        if (u === "TEXTBEFORE" || u === "TEXTAFTER")
          return broadcastOverArrays(
            l.map(retrieve),
            (D) => {
              try {
                return evaluateTextFunction(u, D, {
                  retrieve: (B) => B,
                  issue: issue,
                  isIssue: isFormulaError,
                  format: () => issue("#VALUE!"),
                });
              } catch (error) {
                return isFormulaError(error) ? error : issue("#VALUE!");
              }
            },
            issue,
          );
        const F = evaluateTextFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          temporal: (D) => parseDateText(D, clock),
          format: () => {
            const D = Number(retrieve(l[0]) ?? 0),
              pattern = retrieve(l[1]);
            if (typeof D == "number" && typeof pattern == "string") {
              try {
                if (validateNumberFormatCode(pattern).kind === "temporal")
                  return formatCellValueWithNumberFormat(
                    D,
                    { category: "custom", locale: "en-US", pattern: pattern },
                    clock.dateSystem,
                  ).text;
              } catch {}
              const _ = pattern.replace(/"[^"]*"|\\./g, "");
              if (
                clock.dateSystem === "1904" &&
                /[ydm]/i.test(_) &&
                !/\[[hms]+\]/i.test(_)
              )
                return distribution(u, [
                  D + dateSystemSerialOffset(clock.dateSystem),
                  pattern,
                ]);
            }
            return distribution(u, l);
          },
        });
        if (F !== notTextFunctionSentinel) return F;
        const I = evaluateDefinedNameFunction(u, l, {
          retrieve: retrieve,
          issue: issue,
          isIssue: isFormulaError,
          matches: matchCriteriaValue,
        });
        if (I !== notModernFunctionSentinel) return I;
        if (u === "INDEX") {
          if (l.length < 2 || l.length > 4) return issue("#VALUE!");
          if (isFormulaError(l[0])) return l[0];
          const D = [...l];
          for (let _ = 1; _ < l.length; _++) {
            const C = l[_] == null ? (_ === 3 ? 1 : 0) : retrieve(l[_]);
            if (isFormulaError(C)) return C;
            if (
              Array.isArray(C) ||
              (typeof C == "string" && C.trim() === "") ||
              !Number.isFinite(Number(C)) ||
              Number(C) < 0
            )
              return issue("#VALUE!");
            D[_] = Number(C);
          }
          if (D.length > 3 && Number(D[3]) < 1) return issue("#VALUE!");
          const B = distribution(u, D);
          if (B && typeof B == "object" && "ref" in B) {
            const _ = l[0],
              C = B;
            _?.ref?.sheet && !C.ref.sheet && (C.ref.sheet = _.ref.sheet);
          }
          return B;
        }
        if (u === "IF") {
          if (l.length < 2 || l.length > 3) return issue("#VALUE!");
          const D = retrieve(l[0]);
          if (isFormulaError(D)) return D;
          if (Array.isArray(D))
            return broadcastOverArrays(
              [D, retrieve(l[1]), l.length === 3 ? retrieve(l[2]) : false],
              ([_, C, T]) =>
                isFormulaError(_)
                  ? _
                  : typeof _ == "string" && !/^(TRUE|FALSE)$/i.test(_)
                    ? issue("#VALUE!")
                    : (((
                        typeof _ == "string" ? _.toUpperCase() === "TRUE" : !!_
                      )
                        ? C
                        : T) ?? 0),
              issue,
            );
          if (typeof D == "string" && !/^(TRUE|FALSE)$/i.test(D))
            return issue("#VALUE!");
          const B = typeof D == "string" ? D.toUpperCase() === "TRUE" : !!D;
          return retrieve(B ? l[1] : l.length === 3 ? l[2] : false) ?? 0;
        }
        if (u === "IFERROR" || u === "IFNA") {
          if (l.length !== 2) return issue("#VALUE!");
          const D = retrieve(l[0]);
          return Array.isArray(D)
            ? broadcastOverArrays(
                [D, retrieve(l[1])],
                ([B, _]) =>
                  isFormulaError(B) && (u === "IFERROR" || B.error === "#N/A")
                    ? (_ ?? 0)
                    : (B ?? 0),
                issue,
              )
            : isFormulaError(D) && (u === "IFERROR" || D.error === "#N/A")
              ? (retrieve(l[1]) ?? 0)
              : (D ?? 0);
        }
        const O = l.map(retrieve),
          j = O.flatMap(flattenNestedValues).filter(
            (D) => typeof D == "number",
          );
        if (
          O.flatMap(flattenNestedValues).length > 1e5 ||
          ([
            "REPT",
            "BASE",
            "DEC2BIN",
            "DEC2HEX",
            "DEC2OCT",
            "BIN2HEX",
            "BIN2OCT",
            "HEX2BIN",
            "HEX2OCT",
            "OCT2BIN",
            "OCT2HEX",
            "FIXED",
            "DOLLAR",
          ].includes(u) &&
            ((Math.abs(Number(O[u === "BASE" ? 2 : 1] ?? 0)) > 32767 &&
              !["FIXED", "DOLLAR"].includes(u)) ||
              (["FIXED", "DOLLAR"].includes(u) &&
                Math.abs(Number(O[1] ?? 2)) > 100) ||
              (u === "REPT" &&
                String(O[0] ?? "").length * Number(O[1]) > 32767))) ||
          (["MUNIT", "MDETERM", "MMULT", "GROWTH"].includes(u) &&
            (u === "MUNIT"
              ? Number(O[0]) > 100
              : O.flatMap(flattenNestedValues).length > 1e3)) ||
          ([
            "BINOM.DIST.RANGE",
            "BINOM.INV",
            "HYPGEOM.DIST",
            "POISSON.DIST",
            "FACT",
            "FACTDOUBLE",
            "COMBIN",
            "COMBINA",
            "MULTINOMIAL",
          ].includes(u) &&
            j.some((D) => Math.abs(D) > 1e4)) ||
          (["BESSELI", "BESSELJ", "BESSELK", "BESSELY"].includes(u) &&
            Math.abs(Number(O[1])) > 1e3)
        )
          return issue("#LIMIT!");
        if (u === "ROMAN" && Math.abs(Number(O[0])) > 3999)
          return issue("#VALUE!");
        if (
          [
            "NETWORKDAYS",
            "NETWORKDAYS.INTL",
            "WORKDAY",
            "WORKDAY.INTL",
          ].includes(u) &&
          (Math.abs(
            Number(O[1]) - (u.startsWith("NETWORK") ? Number(O[0]) : 0),
          ) > 366e3 ||
            O.flatMap(flattenNestedValues).length > 1e3)
        )
          return issue("#LIMIT!");
        if (u === "ISERR") return isFormulaError(O[0]) && O[0].error !== "#N/A";
        if (u === "ISLOGICAL") return typeof O[0] == "boolean";
        if (u === "ISNONTEXT") return typeof O[0] != "string";
        if (u === "N")
          return isFormulaError(O[0]) || typeof O[0] == "number"
            ? O[0]
            : O[0] === true
              ? 1
              : 0;
        if (u === "T")
          return isFormulaError(O[0]) || typeof O[0] == "string" ? O[0] : "";
        if (u === "TYPE")
          return isFormulaError(O[0])
            ? 16
            : Array.isArray(O[0])
              ? 64
              : typeof O[0] == "string"
                ? 2
                : typeof O[0] == "boolean"
                  ? 4
                  : 1;
        if (u === "ERROR.TYPE") {
          const D = [
            "#NULL!",
            "#DIV/0!",
            "#VALUE!",
            "#REF!",
            "#NAME?",
            "#NUM!",
            "#N/A",
            "#GETTING_DATA",
          ].indexOf(isFormulaError(O[0]) ? O[0].error : "");
          return D < 0 ? issue("#N/A") : D + 1;
        }
        if (u === "ISERROR") return isFormulaError(O[0]);
        if (u === "ISNA") return isFormulaError(O[0]) && O[0].error === "#N/A";
        if (u === "ISNUMBER") return typeof O[0] == "number";
        if (u === "ISTEXT") return typeof O[0] == "string";
        if (u === "ISBLANK") return O[0] === null;
        if (u === "COUNTA")
          return O.flatMap(flattenNestedValues).filter((D) => D !== null)
            .length;
        if (u === "COUNTBLANK")
          return O.flatMap(flattenNestedValues).filter(
            (D) => D === null || D === "",
          ).length;
        if (
          [
            "SUMIF",
            "AVERAGEIF",
            "COUNTIF",
            "SUMIFS",
            "AVERAGEIFS",
            "COUNTIFS",
          ].includes(u)
        ) {
          const D = u.endsWith("IFS"),
            B = u.startsWith("COUNT"),
            _ = u.startsWith("AVERAGE");
          let C, T;
          if (D) {
            if (((T = B ? O : O.slice(1)), T.length < 2 || T.length % 2 !== 0))
              return issue("#VALUE!");
            C = B
              ? flattenNestedValues(T[0]).map(() => 1)
              : flattenNestedValues(O[0]);
          } else {
            if (O.length < 2 || O.length > (B ? 2 : 3)) return issue("#VALUE!");
            ((T = O.slice(0, 2)),
              (C = B
                ? flattenNestedValues(O[0]).map(() => 1)
                : flattenNestedValues(
                    l.length < 3 || l[2] == null ? O[0] : O[2],
                  )));
          }
          const N = [],
            R = [],
            L = (H) => {
              if (!Array.isArray(H)) return "1:1";
              const Q = H.every(Array.isArray) ? H : [H],
                X = Q[0]?.length ?? 0;
              return Q.length && X && Q.every((q) => q.length === X)
                ? `${Q.length}:${X}`
                : "invalid";
            },
            $ = L(B ? T[0] : D || l.length < 3 || l[2] == null ? O[0] : O[2]);
          for (let H = 0; H < T.length; H += 2) {
            const Q = flattenNestedValues(T[H]);
            if (
              Q.length !== C.length ||
              $ === "invalid" ||
              L(T[H]) !== $ ||
              Array.isArray(T[H + 1])
            )
              return issue("#VALUE!");
            if (isFormulaError(T[H + 1])) return T[H + 1];
            N.push(Q);
            const X = T[H + 1];
            R.push(X == null ? "=0" : typeof X == "number" ? `=${X}` : X);
          }
          const U = [];
          let G = 0;
          for (let H = 0; H < C.length; H++) {
            if (
              ((G += N.reduce(
                (Q, X, J) =>
                  Q +
                  Math.max(1, String(X[H] ?? "").length) *
                    Math.max(1, String(R[J] ?? "").length),
                0,
              )),
              G > 2e6)
            )
              return issue("#LIMIT!");
            if (
              N.every(
                (Q, X) =>
                  !isFormulaError(Q[H]) && matchCriteriaValue(Q[H], R[X]),
              )
            ) {
              if (isFormulaError(C[H])) return C[H];
              typeof C[H] == "number" && U.push(C[H]);
            }
          }
          return _
            ? U.length
              ? exactAverageOfNumbers(U)
              : issue("#DIV/0!")
            : sumDecimalPartsToNumber(U);
        }
        const formulaErrorArg = findFormulaErrorArg(O);
        if (formulaErrorArg && u !== "COUNT") return formulaErrorArg;
        if (u === "TRIM")
          return String(O[0] ?? "")
            .replace(/^ +| +$/g, "")
            .replace(/ +/g, " ");
        if (u === "UPPER") {
          const D = String(O[0] ?? "").toUpperCase();
          return D.length > 32767 ? issue("#LIMIT!") : D;
        }
        if (u === "VALUE") {
          if (O.length !== 1 || Array.isArray(O[0])) return issue("#VALUE!");
          let D = String(O[0] ?? "").trim(),
            B = false;
          D.startsWith("(") &&
            D.endsWith(")") &&
            ((B = true), (D = D.slice(1, -1).trim()));
          const _ = D.endsWith("%");
          if (
            (_ && (D = D.slice(0, -1).trim()),
            (D = D.replace(/^\$/, "")),
            !/^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+|\d*\.\d+)(?:\.\d*)?(?:[eE][+-]?\d+)?$/.test(
              D,
            ))
          )
            return issue("#VALUE!");
          const C =
            (Number(D.replace(/,/g, "")) * (B ? -1 : 1)) / (_ ? 100 : 1);
          return Number.isFinite(C) ? C : issue("#NUM!");
        }
        if (u === "MATCH") {
          if (O.length < 2 || O.length > 3) return issue("#VALUE!");
          const D = O[1];
          if (
            Array.isArray(D) &&
            D.length > 1 &&
            Array.isArray(D[0]) &&
            D[0].length > 1
          )
            return issue("#N/A");
          const B = flattenNestedValues(D),
            _ = O[0],
            C = O[2] === undefined ? 1 : Number(O[2]);
          if (![0, -1, 1].includes(C)) return issue("#N/A");
          if (C === 0) {
            const $ = B.findIndex((U) =>
              typeof _ == "string"
                ? typeof U == "string" &&
                  (/[?*~]/.test(_)
                    ? matchCriteriaValue(U, `=${_}`)
                    : U.toLowerCase() === _.toLowerCase())
                : U === _,
            );
            return $ < 0 ? issue("#N/A") : $ + 1;
          }
          const T = B.map(($) => (typeof $ == "string" ? $.toLowerCase() : $)),
            N = typeof _ == "string" ? _.toLowerCase() : _;
          if (
            (typeof N != "number" && typeof N != "string") ||
            T.some(($) => typeof $ != typeof N)
          )
            return issue("#VALUE!");
          const R = T;
          if (
            R.some(($, U) => U > 0 && (C === 1 ? $ < R[U - 1] : $ > R[U - 1]))
          )
            return issue("#VALUE!");
          let L = -1;
          return (
            R.forEach(($, U) => {
              (C === 1 ? $ <= N : $ >= N) && (L = U);
            }),
            L < 0 ? issue("#N/A") : L + 1
          );
        }
        if (
          [
            "SUM",
            "AVERAGE",
            "MIN",
            "MAX",
            "MEDIAN",
            "LARGE",
            "SMALL",
            "STDEV.S",
            "VAR.S",
          ].includes(u)
        ) {
          const D = u === "LARGE" || u === "SMALL" ? O.slice(0, 1) : O,
            B = [];
          D.forEach((T, N) => {
            const R = !!l[N] && typeof l[N] == "object" && "ref" in l[N];
            for (const L of flattenNestedValues(T))
              typeof L == "number"
                ? B.push(L)
                : !R &&
                  !Array.isArray(T) &&
                  (typeof L == "boolean" ||
                    (typeof L == "string" &&
                      L.trim() !== "" &&
                      Number.isFinite(Number(L)))) &&
                  B.push(Number(L));
          });
          const _ = sumDecimalPartsToNumber(B);
          if (u === "SUM") return _;
          if (u === "AVERAGE")
            return B.length ? _ / B.length : issue("#DIV/0!");
          if (u === "MIN")
            return B.length ? B.reduce((T, N) => Math.min(T, N)) : 0;
          if (u === "MAX")
            return B.length ? B.reduce((T, N) => Math.max(T, N)) : 0;
          if (u === "STDEV.S" || u === "VAR.S") {
            if (B.length < 2) return issue("#DIV/0!");
            const T =
              B.reduce((N, R) => N + (R - _ / B.length) ** 2, 0) /
              (B.length - 1);
            return u === "VAR.S" ? T : Math.sqrt(T);
          }
          if ((B.sort((T, N) => T - N), u === "MEDIAN"))
            return B.length
              ? (B[Math.floor((B.length - 1) / 2)] +
                  B[Math.floor(B.length / 2)]) /
                  2
              : issue("#NUM!");
          const C = Number(O[1]);
          return !Number.isInteger(C) || C < 1 || C > B.length
            ? issue("#NUM!")
            : u === "LARGE"
              ? B[B.length - C]
              : B[C - 1];
        }
        if (u === "SUBSTITUTE") {
          if (O.length < 3 || O.length > 4 || O.some(Array.isArray))
            return issue("#VALUE!");
          const D = String(O[0] ?? ""),
            B = String(O[1] ?? ""),
            _ = String(O[2] ?? ""),
            C = B.length ? D.split(B).length - 1 : 0;
          if (
            O[3] !== undefined &&
            (!Number.isInteger(Number(O[3])) || Number(O[3]) < 1)
          )
            return issue("#VALUE!");
          if (
            D.length +
              (O[3] === undefined ? C : Math.min(C, 1)) *
                (_.length - B.length) >
            32767
          )
            return issue("#LIMIT!");
          if (!B.length) return D;
          if (O[3] === undefined) return D.split(B).join(_);
          let T = 0;
          return D.split(B).reduce(
            (N, R, L) => (L === 0 ? R : N + (++T === Number(O[3]) ? _ : B) + R),
            "",
          );
        }
        if (u === "SEARCH" || u === "SEARCHB") {
          if (O.length < 2 || O.length > 3 || O.some(Array.isArray))
            return issue("#VALUE!");
          const D = String(O[0] ?? "").toLowerCase(),
            B = String(O[1] ?? "").toLowerCase(),
            _ = O[2] === undefined ? 1 : Number(O[2]);
          if (!Number.isInteger(_) || _ < 1 || _ > B.length)
            return issue("#VALUE!");
          if (!/[?*~]/.test(D)) {
            const C = B.indexOf(D, _ - 1);
            return C < 0 ? issue("#VALUE!") : C + 1;
          }
          if (B.length * B.length * D.length > 2e6) return issue("#LIMIT!");
          for (let C = _ - 1; C < B.length; C++)
            if (matchCriteriaValue(B.slice(C), `${D}*`)) return C + 1;
          return issue("#VALUE!");
        }
        if (
          u === "CONCATENATE" &&
          O.flatMap(flattenNestedValues).reduce(
            (D, B) => D + String(B ?? "").length,
            0,
          ) > 32767
        )
          return issue("#LIMIT!");
        const V = distribution(u, l);
        return /^IM/.test(u) &&
          !["IMABS", "IMREAL", "IMAGINARY", "IMARGUMENT"].includes(u) &&
          typeof V == "number"
          ? String(V)
          : typeof V == "string" && V.length > 32767
            ? issue("#LIMIT!")
            : V;
      } catch (error) {
        return isFormulaError(error) ? issue(error.error) : issue("#ERROR!");
      }
    }),
    e
  );
}
function formatNumberWithPattern(e, cell, t = "1900") {
  if (cell?.numberFormat)
    try {
      return formatCellValueWithNumberFormat(e, cell.numberFormat, t).text;
    } catch {
      return "#FORMAT!";
    }
  if (e === null) return "";
  if (typeof e == "boolean") return e ? "TRUE" : "FALSE";
  if (typeof e != "number") return e;
  if (!Number.isFinite(e)) return "#NUM!";
  if (cell?.format === "date") {
    if (cell.numFmt)
      try {
        if (validateNumberFormatCode(cell.numFmt).kind === "temporal")
          return formatCellValueWithNumberFormat(
            e,
            { category: "custom", locale: "en-US", pattern: cell.numFmt },
            t,
          ).text;
      } catch {}
    return formatCellValueWithNumberFormat(
      e,
      { category: "date", locale: "en-US", pattern: "yyyy-mm-dd" },
      t,
    ).text;
  }
  if (cell?.numFmt) {
    if (cell.numFmt.toLowerCase() === "general")
      return String(Number(e.toPrecision(15)));
    const n = numberFormatToIntlOptions(e, cell.numFmt);
    if (n !== undefined) return n;
  }
  return cell?.format === "percent"
    ? new Intl.NumberFormat("en-US", {
        style: "percent",
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }).format(e)
    : cell?.format === "currency"
      ? new Intl.NumberFormat("en-US", {
          style: "currency",
          currency: "USD",
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        }).format(e)
      : cell?.format === "number"
        ? new Intl.NumberFormat("en-US", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          }).format(e)
        : String(Number(e.toPrecision(15)));
}
const defaultCalculationSettings = Object.freeze({
    mode: "automatic",
    iterative: false,
    maxIterations: 100,
    maxChange: 0.001,
    timeBudgetMs: 2500,
  }),
  throwInvalidCalculationSettings = () => {
    throw Error(
      "Invalid calculation settings. Choose Automatic or Manual, explicit iteration, 1–1000 iterations, convergence from 0 to 1, a 1–10000 ms time budget, 1–128 MiB working data and 1000–10000000 work units.",
    );
  };
function normalizeCalculationSettings(e) {
  return !isPlainObjectValue(e) ||
    Object.keys(e).some(
      (r) =>
        ![
          ...Object.keys(defaultCalculationSettings),
          "memoryBudgetMiB",
          "workBudget",
        ].includes(r),
    ) ||
    typeof e.mode != "string" ||
    !["automatic", "manual"].includes(e.mode) ||
    typeof e.iterative != "boolean" ||
    !Number.isInteger(e.maxIterations) ||
    Number(e.maxIterations) < 1 ||
    Number(e.maxIterations) > 1e3 ||
    typeof e.maxChange != "number" ||
    !Number.isFinite(e.maxChange) ||
    e.maxChange < 0 ||
    e.maxChange > 1 ||
    !Number.isInteger(e.timeBudgetMs) ||
    Number(e.timeBudgetMs) < 1 ||
    Number(e.timeBudgetMs) > 1e4 ||
    (e.memoryBudgetMiB !== undefined &&
      (!Number.isInteger(e.memoryBudgetMiB) ||
        Number(e.memoryBudgetMiB) < 1 ||
        Number(e.memoryBudgetMiB) > 128)) ||
    (e.workBudget !== undefined &&
      (!Number.isInteger(e.workBudget) ||
        Number(e.workBudget) < 1e3 ||
        Number(e.workBudget) > 1e7))
    ? throwInvalidCalculationSettings()
    : {
        mode: e.mode,
        iterative: e.iterative,
        maxIterations: Number(e.maxIterations),
        maxChange: e.maxChange,
        timeBudgetMs: Number(e.timeBudgetMs),
        ...(e.memoryBudgetMiB !== undefined
          ? { memoryBudgetMiB: Number(e.memoryBudgetMiB) }
          : {}),
        ...(e.workBudget !== undefined
          ? { workBudget: Number(e.workBudget) }
          : {}),
      };
}
function importFormulaCacheEvidence(workbook, r) {
  workbook.calculation !== undefined &&
    (r.calculation = normalizeCalculationSettings(workbook.calculation));
  const importedCalculationCaches = workbook.importedCalculationCaches;
  if (importedCalculationCaches === undefined) return;
  const n = () => {
    throw Error(
      "Invalid imported formula-cache evidence. Caches must remain separately labeled source evidence.",
    );
  };
  if (
    !isPlainObjectValue(importedCalculationCaches) ||
    importedCalculationCaches.schemaVersion !== 1 ||
    typeof importedCalculationCaches.fileName != "string" ||
    !importedCalculationCaches.fileName ||
    importedCalculationCaches.fileName.length > 255 ||
    typeof importedCalculationCaches.sourceSha256 != "string" ||
    !/^[a-f0-9]{64}$/.test(importedCalculationCaches.sourceSha256) ||
    typeof importedCalculationCaches.importedAt != "string" ||
    !Number.isFinite(Date.parse(importedCalculationCaches.importedAt)) ||
    !isPlainObjectValue(importedCalculationCaches.sourceSettings) ||
    !Array.isArray(importedCalculationCaches.entries) ||
    importedCalculationCaches.entries.length > 1e5
  )
    return n();
  const sourceSettings = importedCalculationCaches.sourceSettings;
  if (
    typeof sourceSettings.mode != "string" ||
    sourceSettings.mode.length > 40 ||
    typeof sourceSettings.iterative != "boolean" ||
    (sourceSettings.maxIterations !== null &&
      (!Number.isSafeInteger(sourceSettings.maxIterations) ||
        Number(sourceSettings.maxIterations) < 1)) ||
    (sourceSettings.fullPrecision !== undefined &&
      typeof sourceSettings.fullPrecision != "boolean") ||
    (sourceSettings.maxChange !== null &&
      (typeof sourceSettings.maxChange != "number" ||
        !Number.isFinite(sourceSettings.maxChange) ||
        sourceSettings.maxChange < 0))
  )
    return n();
  const a = (u) => {
      const l = /^([A-Z]+)([1-9]\d*)$/.exec(u);
      return (
        !!l &&
        Number(l[2]) <= 1048576 &&
        [...l[1]].reduce((c, h) => c * 26 + h.charCodeAt(0) - 64, 0) <= 16384
      );
    },
    s = new Set(),
    entries = importedCalculationCaches.entries.map((entry) => {
      if (
        !isPlainObjectValue(entry) ||
        typeof entry.sheetId != "string" ||
        !entry.sheetId ||
        entry.sheetId.length > 100 ||
        typeof entry.address != "string" ||
        !/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(entry.address) ||
        !a(entry.address) ||
        typeof entry.formula != "string" ||
        !entry.formula.startsWith("=") ||
        entry.formula.length > 8192
      )
        return n();
      const l = JSON.stringify([entry.sheetId, entry.address]);
      if (s.has(l)) return n();
      s.add(l);
      let c;
      if (entry.cache !== undefined) {
        const cache = entry.cache;
        if (
          !isPlainObjectValue(cache) ||
          !["error", "blank", "number", "text", "boolean"].includes(
            String(cache.type),
          ) ||
          (cache.type === "blank" && cache.value !== null) ||
          (cache.type === "number" &&
            (typeof cache.value != "number" ||
              !Number.isFinite(cache.value))) ||
          (cache.type === "boolean" && typeof cache.value != "boolean") ||
          (["text", "error"].includes(String(cache.type)) &&
            (typeof cache.value != "string" || cache.value.length > 32767)) ||
          (cache.type === "error" &&
            !/^#[A-Z0-9_/?!]+$/.test(String(cache.value)))
        )
          return n();
        c = { type: cache.type, value: cache.value };
      }
      return {
        sheetId: entry.sheetId,
        address: entry.address,
        formula: entry.formula,
        ...(c ? { cache: c } : {}),
      };
    });
  r.importedCalculationCaches = {
    schemaVersion: 1,
    fileName: importedCalculationCaches.fileName,
    sourceSha256: importedCalculationCaches.sourceSha256,
    importedAt: importedCalculationCaches.importedAt,
    sourceSettings: {
      mode: sourceSettings.mode,
      iterative: sourceSettings.iterative,
      maxIterations: sourceSettings.maxIterations,
      maxChange: sourceSettings.maxChange,
      ...(sourceSettings.fullPrecision !== undefined
        ? { fullPrecision: sourceSettings.fullPrecision }
        : {}),
    },
    entries: entries,
  };
}
const calculatorRecords = new WeakMap(),
  carriedCalculatorRecords = new WeakMap(),
  j5 = 2e5,
  pendingFormulaValue = { type: "error", value: "#PENDING!" },
  makeFormulaError = (e) => ({ type: "error", value: e }),
  dynamicFunctionPattern = new RegExp(
    `\\b(?:${[...matrixResultFunctionNames, "LET", "LAMBDA", "INDIRECT", "OFFSET", "ISFORMULA", "TEXTBEFORE", "TEXTAFTER"].join("|")})\\s*\\(`,
    "i",
  ),
  _O = { test: (e) => findVolatileFunctionNames(e).length > 0 },
  ik = (e, r) =>
    !!e.definedNames?.length ||
    [...r.values()].some((n) => {
      const o = n.formula.replace(/"(?:[^"]|"")*"/g, '""');
      return (
        dynamicFunctionPattern.test(o) ||
        /[{}#]/.test(o) ||
        /^=\s*[^()+*/]*:/.test(o) ||
        (/\b(?:IF|IFERROR|IFNA)\s*\(/i.test(o) && o.includes(":"))
      );
    }),
  calculationSnapshotCache = new WeakMap();
const sheetInputCache = new WeakMap();
function buildSheetCalculationSnapshot(e) {
  const r = isDeepFrozenValue(e.cells)
    ? sheetInputCache.get(e.cells)?.get(e.id)
    : undefined;
  if (r) return r;
  const formulas = [];
  let n, o;
  const inputs = () =>
      (n ??= Object.entries(e.cells).map(([l, cell]) => [
        `${e.id}:${l}`,
        JSON.stringify([cell.raw, !!cell.literal]),
      ])),
    signature = () => (o ??= JSON.stringify(inputs()));
  isDeepFrozenValue(e.cells) || signature();
  for (const address in e.cells) {
    const cell = e.cells[address],
      h = `${e.id}:${address}`;
    !cell.literal &&
      typeof cell.raw == "string" &&
      cell.raw.startsWith("=") &&
      formulas.push([
        h,
        { sheetId: e.id, address: address, formula: cell.raw },
      ]);
  }
  const d = {
    inputs: inputs,
    formulas: formulas,
    signature: signature,
    volatile: [
      ...new Set(
        formulas.flatMap(([, l]) => findVolatileFunctionNames(l.formula)),
      ),
    ],
  };
  if (isDeepFrozenValue(e.cells)) {
    const l = sheetInputCache.get(e.cells) ?? new Map();
    (l.set(e.id, d), sheetInputCache.set(e.cells, l));
  }
  return d;
}
function buildWorkbookCalculationDescriptor(workbook) {
  const calculationSettings = normalizeCalculationSettings(
    workbook.calculation ?? defaultCalculationSettings,
  );
  return JSON.stringify([
    workbook.id,
    calculationSettings,
    normalizeDatePolicy(workbook.datePolicy),
    workbook.definedNames,
    workbook.workbookLinks,
    workbook.extensions,
    getExtensionCacheEntries(workbook),
    workbook.sheets.map((sheet) => [
      sheet.id,
      sheet.name,
      sheet.rowCount,
      sheet.colCount,
      sheet.hiddenRows ?? [],
      sheet.hiddenCols ?? [],
      !!sheet.hidden,
      !!sheet.veryHidden,
      sheet.tables ?? [],
      sheet.merges ?? [],
      sheet.outlineGroups ?? [],
    ]),
  ]);
}
function buildWorkbookCalculationSignature(workbook) {
  const r = workbook.sheets.map(buildSheetCalculationSnapshot);
  let t;
  const inputs = () => (t ??= new Map(r.flatMap((h) => h.inputs()))),
    formulas = new Map(r.flatMap((h) => h.formulas)),
    calculationSettings = normalizeCalculationSettings(
      workbook.calculation ?? defaultCalculationSettings,
    ),
    workbookCalculationDescriptor =
      buildWorkbookCalculationDescriptor(workbook);
  let d, u;
  const inputSignature = () =>
      (d ??= `[${JSON.stringify(workbook.revision)},${JSON.stringify(workbookCalculationDescriptor)},[${r.map((h) => h.signature()).join(",")}]]`),
    signature = () =>
      (u ??= `[${inputSignature()},${JSON.stringify(workbook.importedCalculationCaches ?? null)}]`);
  return (
    isDeepFrozenValue(workbook) || signature(),
    {
      inputs: inputs,
      formulas: formulas,
      settings: calculationSettings,
      metadata: workbookCalculationDescriptor,
      volatile: [
        ...new Set(
          r
            .flatMap((h) => h.volatile)
            .concat(
              (workbook.definedNames ?? []).flatMap((h) =>
                findVolatileFunctionNames(h.refersTo),
              ),
            ),
        ),
      ].sort(),
      inputSignature: inputSignature,
      signature: signature,
    }
  );
}
function getCachedCalculationSnapshot(workbook) {
  if (!isDeepFrozenValue(workbook))
    return buildWorkbookCalculationSignature(workbook);
  let r = calculationSnapshotCache.get(workbook);
  return (
    r ||
      ((r = buildWorkbookCalculationSignature(workbook)),
      calculationSnapshotCache.set(workbook, r)),
    r
  );
}
function isCalculationRecordCurrent(
  e,
  workbook,
  cachedCalculationSnapshot = getCachedCalculationSnapshot(workbook),
) {
  return (
    !!e &&
    (isDeepFrozenValue(workbook) ||
      e.signature() === cachedCalculationSnapshot.signature())
  );
}
function MO(workbook, r) {
  if (
    workbook.revision !== r.revision ||
    buildWorkbookCalculationDescriptor(workbook) !==
      buildWorkbookCalculationDescriptor(r) ||
    (workbook.importedCalculationCaches !== r.importedCalculationCaches &&
      JSON.stringify(workbook.importedCalculationCaches) !==
        JSON.stringify(r.importedCalculationCaches))
  )
    return false;
  for (let t = 0; t < workbook.sheets.length; t++) {
    const cells = workbook.sheets[t].cells,
      o = r.sheets[t].cells;
    if (cells === o) continue;
    const a = Object.keys(cells);
    if (a.length !== Object.keys(o).length) return false;
    for (const s of a) {
      const cell = cells[s],
        u = o[s];
      if (!u || cell.raw !== u.raw || !!cell.literal != !!u.literal)
        return false;
    }
  }
  return true;
}
function hasCalculatorRecord(workbook) {
  return (
    calculatorRecords.has(workbook) || carriedCalculatorRecords.has(workbook)
  );
}
function carryCalculatorRecordForward(workbook, r) {
  if (workbook === r || workbook.id !== r.id || r.revision <= workbook.revision)
    return;
  const t = calculatorRecords.get(workbook);
  t &&
    isCalculationRecordCurrent(t, workbook) &&
    carriedCalculatorRecords.set(r, t);
}
function zie(e, workbook) {
  if (e === workbook) return;
  const t = calculatorRecords.get(e);
  t && isCalculationRecordCurrent(t, e) && MO(e, workbook)
    ? calculatorRecords.set(workbook, t)
    : carriedCalculatorRecords.has(e) &&
      MO(e, workbook) &&
      carriedCalculatorRecords.set(workbook, carriedCalculatorRecords.get(e));
}
function withCalculationSettings(workbook, r) {
  return { ...workbook, calculation: normalizeCalculationSettings(r) };
}
function deepFreeze(e) {
  !e ||
    typeof e != "object" ||
    Object.isFrozen(e) ||
    (Object.values(e).forEach(deepFreeze), Object.freeze(e));
}
function getCurrentCalculatorRecord(activeWorkbook) {
  const r = calculatorRecords.get(activeWorkbook);
  return isCalculationRecordCurrent(r, activeWorkbook)
    ? { calculator: r.calculator, report: r.report }
    : undefined;
}
const reportCellValueMaps = new WeakMap();
function buildReportCellValueMap(e) {
  if (!e) return new Map();
  const r = reportCellValueMaps.get(e.report);
  if (r) return r;
  const t = new Map();
  for (const cell of e.report.cells) {
    const o = `${cell.sheetId}:${cell.address}`,
      a = e.values.get(o);
    a && a.value !== "#PENDING!"
      ? t.set(o, { revision: e.report.revision, result: { ...a } })
      : cell.lastKnown && t.set(o, cell.lastKnown);
  }
  return t;
}
function getCalculationInputSignature(workbook) {
  return getCachedCalculationSnapshot(workbook).inputSignature();
}
const filterEvaluationLimits = Object.freeze({
    rows: 1e4,
    cells: 1e5,
    columns: 64,
    conditions: 128,
    values: 1e3,
    text: 1024,
    comparisonWork: 5e6,
  }),
  savedFilterViewHelpText =
    "View only. Column rules intersect (AND); advanced groups use OR of AND rules. Counts include blank records in the declared range. Manual/outline hiding is separate. Full scope includes every data row; visible scope excludes filtered and manually hidden rows. Applying a view leaves source addresses, formulas, identities and history unchanged. Native JSON, XLSX, print and AI selections use full source data. Active masks are session state; named criteria can be saved in native JSON and require a fresh preview before reuse.",
  normalizeComparableText = (e, r = false) =>
    r ? e.normalize("NFC") : e.normalize("NFC").toLocaleLowerCase("en-US"),
  comparableCellValueKey = (e, r = false) =>
    isNullOrEmptyValue(e)
      ? "blank"
      : `${typeof e}:${typeof e == "string" ? normalizeComparableText(e, r) : e}`,
  textFilterOperators = [
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
  ],
  isStorableCellValue = (e) =>
    e === null ||
    typeof e == "boolean" ||
    (typeof e == "number" && Number.isFinite(e)) ||
    (typeof e == "string" && e.length <= 32767);
function resolveDatePolicyContext(e, r, workbook) {
  if (r === "number") return typeof e == "number" ? e : null;
  if (r === "date") {
    const datePolicy = normalizeDatePolicy(workbook.datePolicy),
      o =
        typeof e == "number"
          ? e
          : typeof e == "string"
            ? parseDateText(e, datePolicy)
            : undefined;
    return o !== undefined &&
      o >= 0 &&
      o < dateSystemMaxSerial(datePolicy.dateSystem) + 1
      ? o
      : null;
  }
  return typeof e == "string" ? e : null;
}
function parseWildcardPattern(e, r) {
  const t = [];
  for (let d = 0; d < r.length; d++) {
    const u = r[d];
    u === "~" && d + 1 < r.length
      ? t.push({ text: r[++d] })
      : t.push(
          u === "*" ? { star: true } : u === "?" ? { any: true } : { text: u },
        );
  }
  let n = 0,
    o = 0,
    a = -1,
    s = 0;
  for (; n < e.length;)
    if (o < t.length && (t[o].any || t[o].text === e[n])) (n++, o++);
    else if (o < t.length && t[o].star) ((a = o++), (s = n));
    else if (a >= 0) ((o = a + 1), (n = ++s));
    else return false;
  for (; o < t.length && t[o].star;) o++;
  return o === t.length;
}
function matchesTextFilter(e, r, workbook, n) {
  if (r.operator === "blank") return isNullOrEmptyValue(e);
  if (r.operator === "notBlank") return !isNullOrEmptyValue(e);
  if (r.type === "value") {
    const s =
      comparableCellValueKey(e, n) === comparableCellValueKey(r.value, n);
    return r.operator === "equals" ? s : !s;
  }
  let datePolicyContext = resolveDatePolicyContext(e, r.type, workbook),
    a = resolveDatePolicyContext(r.value, r.type, workbook);
  if (datePolicyContext === null || a === null) return false;
  switch (
    (typeof datePolicyContext == "string" &&
      typeof a == "string" &&
      ((datePolicyContext = normalizeComparableText(datePolicyContext, n)),
      (a = normalizeComparableText(a, n))),
    r.operator)
  ) {
    case "equals":
      return datePolicyContext === a;
    case "notEquals":
      return datePolicyContext !== a;
    case "gt":
      return datePolicyContext > a;
    case "gte":
      return datePolicyContext >= a;
    case "lt":
      return datePolicyContext < a;
    case "lte":
      return datePolicyContext <= a;
    case "contains":
      return String(datePolicyContext).includes(String(a));
    case "notContains":
      return !String(datePolicyContext).includes(String(a));
    case "startsWith":
      return String(datePolicyContext).startsWith(String(a));
    case "endsWith":
      return String(datePolicyContext).endsWith(String(a));
    case "wildcard":
      return parseWildcardPattern(String(datePolicyContext), String(a));
  }
  return false;
}
function assertCalculatedBeforeFilter(e, sheetId, t, workbook) {
  const o = e(sheetId, t),
    cell = workbook?.sheets.find((sheet) => sheet.id === sheetId)?.cells[t];
  if (
    cell &&
    cell.raw === o &&
    (cell.literal || typeof cell.raw != "string" || !cell.raw.startsWith("="))
  )
    return o;
  if (
    typeof o == "string" &&
    /^#(?:PENDING|LIMIT|TIMEOUT|CANCELLED|BUSY|BLOCKED|CALC_LIMIT)!?$/i.test(o)
  )
    throw Error(
      `Calculation is incomplete at ${t}. Recalculate before filtering; unchecked rows are never reported as an empty population.`,
    );
  return o;
}
const savedFilterViewLimits = Object.freeze({
  views: 20,
  name: 80,
  characters: 512e3,
});
function validateSavedFilterViewMetadata(e, r) {
  if (
    !e ||
    typeof e != "object" ||
    Array.isArray(e) ||
    ![Object.prototype, null].includes(Object.getPrototypeOf(e)) ||
    Object.keys(e).some((t) => !r.includes(t))
  )
    throw Error("Invalid saved filter view metadata.");
}
function isValidFilterViewText(e, r) {
  return (
    typeof e == "string" &&
    !!e.trim() &&
    e.length <= r &&
    !/[\u0000-\u001f\u007f]/.test(e)
  );
}
function normalizeNameForComparison(savedFilterViewName) {
  return savedFilterViewName.normalize("NFC").trim().toLocaleLowerCase("en-US");
}
function normalizeSavedFilterViewName(e) {
  if (!isValidFilterViewText(e, savedFilterViewLimits.name))
    throw Error("Use a nonempty filter view name of up to 80 characters.");
  return e.normalize("NFC").trim();
}
function assertSavedFilterViewsEditable(workbook, sheetId) {
  if (
    workbook.structureProtection?.enabled ||
    workbook.sheets.find((sheet) => sheet.id === sheetId)?.protection?.enabled
  )
    throw Error(
      "Unprotect the workbook and source worksheet before changing saved filter views.",
    );
}
function findSavedFilterView(workbook, r) {
  const t = workbook.savedFilterViews?.find(
    (savedFilterView) => savedFilterView.id === r,
  );
  if (!t) throw Error("This saved filter view no longer exists.");
  return t;
}
const consolidationFormatVersion = "ledgerglass-consolidation-v1",
  consolidationLimits = Object.freeze({
    sources: 8,
    categories: 3,
    contributions: 2e3,
    outputRows: 4e3,
    outputCells: 4e4,
    groups: 1e3,
    metadataCharacters: 8e6,
  }),
  boundedLabelSchema = string()
    .min(1)
    .max(128)
    .refine((e) => !!e.trim() && !/[\u0000-\u001f\u007f]/.test(e)),
  shortLabelSchema = string().min(1).max(100),
  V5 = string().regex(/^[a-f0-9]{64}$/),
  scalarValueSchema = union([
    string().min(1).max(32767),
    number().finite(),
    boolean(),
  ]),
  categoryValueSchema = discriminatedUnion("kind", [
    strictObject({ kind: literal("constant"), value: boundedLabelSchema }),
    strictObject({ kind: literal("field"), field: boundedLabelSchema }),
  ]),
  consolidationSourceSchema = strictObject({
    id: shortLabelSchema,
    workbookId: shortLabelSchema,
    sheetId: shortLabelSchema,
    range: string().min(1).max(50),
    tableId: shortLabelSchema.optional(),
    categories: array(boundedLabelSchema).min(1).max(3),
    amountField: boundedLabelSchema,
    period: categoryValueSchema,
    currency: categoryValueSchema,
    unit: categoryValueSchema,
    duplicatePolicy: _enum(["reject", "sum"]),
  }),
  consolidationDefinitionSchema = strictObject({
    engine: literal(consolidationFormatVersion),
    sources: array(consolidationSourceSchema).min(1).max(8),
    categoryLabels: array(boundedLabelSchema).min(1).max(3),
    period: boundedLabelSchema,
    currency: boundedLabelSchema,
    unit: boundedLabelSchema,
    linkedAmounts: boolean(),
    collapseDetails: boolean(),
  }).superRefine((e, r) => {
    const t = [
      "Row type",
      ...e.categoryLabels,
      "Period",
      "Currency",
      "Unit",
      "Amount",
      "Source",
      "Link state",
    ];
    (new Set(t.map((n) => n.toLowerCase())).size !== t.length ||
      new Set(e.sources.map((source) => source.id)).size !== e.sources.length ||
      e.sources.some(
        (source) =>
          source.categories.length !== e.categoryLabels.length ||
          new Set(source.categories).size !== source.categories.length,
      )) &&
      r.addIssue({
        code: "custom",
        message:
          "Use unique source IDs, distinct category fields, and distinct output headers with matching category counts.",
      });
  }),
  consolidationSourceSnapshotSchema = strictObject({
    sourceId: shortLabelSchema,
    workbookId: shortLabelSchema,
    workbookName: string().min(1).max(200),
    sheetId: shortLabelSchema,
    sheetName: string().min(1).max(31),
    row: number().int().min(2).max(1048576),
    address: string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/),
    locator: string().max(300),
    categories: array(scalarValueSchema).min(1).max(3),
    amount: number().finite(),
    period: boundedLabelSchema,
    currency: boundedLabelSchema,
    unit: boundedLabelSchema,
    link: _enum(["live", "snapshot", "approved snapshot"]),
  }),
  consolidationIssueSchema = strictObject({
    sourceId: shortLabelSchema,
    row: number().int().positive().optional(),
    code: string().max(40),
    severity: _enum(["error", "warning"]),
    message: string().max(4e3),
  }),
  consolidationSourceResultSchema = strictObject({
    sourceId: shortLabelSchema,
    workbookId: shortLabelSchema,
    workbookName: string().max(200),
    sheetId: shortLabelSchema,
    sheetName: string().max(31),
    range: string().max(50),
    fingerprint: V5,
    records: number().int().min(0).max(500),
    acceptedRecords: number().int().min(0).max(500),
    amount: number().finite(),
    periods: array(boundedLabelSchema).max(500),
    currencies: array(boundedLabelSchema).max(500),
    units: array(boundedLabelSchema).max(500),
  }),
  consolidationSubtotalSchema = strictObject({
    categories: array(scalarValueSchema).min(1).max(3),
    level: number().int().min(1).max(3),
    count: number().int().min(1).max(2e3),
    amount: number().finite(),
  }),
  consolidationTotalsSchema = strictObject({
    sources: array(consolidationSourceResultSchema).max(8),
    contributions: array(consolidationSourceSnapshotSchema).max(2e3),
    subtotals: array(consolidationSubtotalSchema).max(1e3),
    exceptions: array(consolidationIssueSchema).max(8e3),
    grandTotal: number().finite().nullable(),
    sourceTotal: number().finite().nullable(),
    difference: number().finite().nullable(),
    complete: boolean(),
  }),
  consolidationEvidenceSchema = strictObject({
    engine: literal(consolidationFormatVersion),
    definition: consolidationDefinitionSchema,
    review: consolidationTotalsSchema,
    origin: strictObject({
      workbookId: shortLabelSchema,
      revision: number().int().nonnegative(),
      at: string().datetime(),
    }),
    evidenceHash: V5,
    outputRange: string().max(50),
    outputHash: V5,
  });
function parseConsolidationDefinition(definition) {
  const r = consolidationDefinitionSchema.safeParse(definition);
  if (!r.success)
    throw Error(`Invalid consolidation: ${r.error.issues[0]?.message}`);
  return r.data;
}
const serializeTypedKeyValues = (e) =>
  JSON.stringify(e.map((r) => [typeof r, r]));
function sumConsolidationAmounts(e) {
  let r = 0,
    t = 0;
  for (const n of e) {
    const o = n - t,
      a = r + o;
    if (((t = a - r - o), (r = a), !Number.isFinite(r) || !Number.isFinite(t)))
      throw Error("Consolidation total exceeds finite numeric limits.");
  }
  return Object.is(r, -0) ? 0 : r;
}
function buildConsolidationSubtotals(contributions) {
  const r = new Map();
  for (const contribution of contributions)
    for (let n = 1; n <= contribution.categories.length; n++) {
      const o = serializeTypedKeyValues(contribution.categories.slice(0, n)),
        a = r.get(o) ?? [];
      (a.push(contribution), r.set(o, a));
    }
  if (r.size > consolidationLimits.groups)
    throw Error(
      "Consolidation exceeds 1,000 subtotal groups. Narrow the sources; no rows were truncated.",
    );
  return [...r].map(([t, n]) => {
    const level = JSON.parse(t).length;
    return {
      categories: n[0].categories.slice(0, level),
      level: level,
      count: n.length,
      amount: sumConsolidationAmounts(n.map((a) => a.amount)),
    };
  });
}
function verifyConsolidationEvidence(workbook, r) {
  let t = 0;
  for (const sheet of r.sheets) {
    const d = workbook.sheets?.find((f) => f.id === sheet.id)?.consolidation;
    if (d === undefined) continue;
    if (
      ((t += JSON.stringify(d).length),
      t > consolidationLimits.metadataCharacters)
    )
      throw Error(
        "Consolidation evidence exceeds 8,000,000 characters. Retain native backups before removing obsolete reports.",
      );
    const u = consolidationEvidenceSchema.safeParse(d);
    if (!u.success)
      throw Error(
        `Invalid consolidation evidence: ${u.error.issues[0]?.message}`,
      );
    const data = u.data,
      review = data.review,
      h = data.definition.sources.map((source) => source.id);
    if (
      !review.complete ||
      review.exceptions.some((exception) => exception.severity === "error") ||
      !review.contributions.length ||
      review.sources.length !== h.length ||
      new Set(review.sources.map((source) => source.sourceId)).size !==
        h.length ||
      review.sources.some((source) => !h.includes(source.sourceId)) ||
      review.contributions.some(
        (contribution) =>
          !h.includes(contribution.sourceId) ||
          contribution.categories.length !==
            data.definition.categoryLabels.length ||
          contribution.period !== data.definition.period ||
          contribution.currency !== data.definition.currency ||
          contribution.unit !== data.definition.unit,
      ) ||
      review.grandTotal !==
        sumConsolidationAmounts(
          review.contributions.map((contribution) => contribution.amount),
        ) ||
      review.sourceTotal !==
        sumConsolidationAmounts(
          review.sources.map((source) => source.amount),
        ) ||
      review.difference !== review.grandTotal - review.sourceTotal ||
      fingerprintCanonicalValue(review.subtotals) !==
        fingerprintCanonicalValue(
          buildConsolidationSubtotals(review.contributions),
        ) ||
      data.evidenceHash !==
        fingerprintCanonicalValue({
          definition: data.definition,
          review: review,
          origin: data.origin,
        })
    )
      throw Error(
        "Consolidation evidence does not reconcile to its sources, categories or retained fingerprint.",
      );
    for (const source of review.sources) {
      const p = review.contributions.filter(
        (contribution) => contribution.sourceId === source.sourceId,
      );
      if (
        source.records !== p.length ||
        source.acceptedRecords !== p.length ||
        source.amount !== sumConsolidationAmounts(p.map((m) => m.amount)) ||
        p.some(
          (m) =>
            m.workbookId !== source.workbookId || m.sheetId !== source.sheetId,
        )
      )
        throw Error("Consolidation source evidence does not reconcile.");
    }
    sheet.consolidation = data;
  }
}
const externalSourceFormatVersion = "ledgerglass-sources-v1",
  externalSourceKinds = [
    "database",
    "odbc",
    "odata",
    "api",
    "erp",
    "file-folder",
  ],
  externalSourceLimits = Object.freeze({
    sources: 20,
    columns: 50,
    rows: 5e3,
    cells: 1e5,
    responseBytes: 8 * 1024 * 1024,
    timeoutMs: 3e4,
    metadataCharacters: 2e5,
  }),
  sourceIdSchema = string().regex(/^[A-Za-z0-9][A-Za-z0-9._-]{0,99}$/),
  columnNameSchema = string()
    .min(1)
    .max(128)
    .refine((e) => e === e.trim() && !/[\u0000-\u001f\u007f]/.test(e)),
  isoDateTimeSchema = string().datetime({ offset: true }),
  sourceHashSchema = string().regex(/^[a-f0-9]{64}$/),
  sourceColumnSchema = strictObject({
    name: columnNameSchema,
    type: _enum(["text", "number", "boolean", "date"]),
    nullable: boolean(),
  }),
  sourceTableSchema = strictObject({
    id: sourceIdSchema,
    label: columnNameSchema,
    kind: _enum(externalSourceKinds),
    schemaVersion: sourceIdSchema,
    columns: array(sourceColumnSchema).min(1).max(externalSourceLimits.columns),
    controlFields: array(columnNameSchema).max(externalSourceLimits.columns),
  }).superRefine((e, r) => {
    (new Set(e.columns.map((column) => column.name.toLowerCase())).size !==
      e.columns.length ||
      new Set(e.controlFields).size !== e.controlFields.length ||
      e.controlFields.some(
        (controlField) =>
          !e.columns.some(
            (column) =>
              column.name === controlField && column.type === "number",
          ),
      )) &&
      r.addIssue({
        code: "custom",
        message: "Use unique columns and declared numeric control fields.",
      });
  }),
  sourceApprovalSchema = strictObject({
    source: sourceTableSchema,
    expiresAt: isoDateTimeSchema,
  }),
  sourceCellValueSchema = union([
    string().max(32767),
    number().finite(),
    boolean(),
    _null(),
  ]),
  sourceResultPayloadSchema = strictObject({
    protocol: literal(externalSourceFormatVersion),
    source: sourceTableSchema,
    sourceTimestamp: isoDateTimeSchema,
    version: sourceIdSchema,
    rows: array(
      array(sourceCellValueSchema).max(externalSourceLimits.columns),
    ).max(externalSourceLimits.rows),
    complete: literal(true),
  }),
  externalSourceErrorCodes = [
    "unavailable",
    "refused",
    "expired",
    "schema",
    "limit",
    "timeout",
    "cancelled",
    "invalid",
    "stale",
  ],
  externalSourceErrorMessages = {
    unavailable:
      "The source is unavailable. The last successful data is retained. Check the host connection, then retry.",
    refused:
      "This source is not authorized by the host. No source data was replaced.",
    expired:
      "Source access expired. Renew access in the host application and preview again.",
    schema:
      "The source schema changed. The working population is retained. Have the host approve the new schema and import it into a new worksheet.",
    limit:
      "The source exceeds the published row, column, cell or response limit. No rows were truncated or published.",
    timeout:
      "Source refresh timed out. The last successful data is retained; retry when the host is available.",
    cancelled:
      "Source refresh cancelled. The last successful data is retained.",
    invalid:
      "The source returned an invalid or incomplete result. The last successful data is retained.",
    stale:
      "The workbook, preview or source approval changed. Preview again before applying a refresh.",
  };
class ExternalSourceError extends Error {
  constructor(r) {
    (super(externalSourceErrorMessages[r]),
      (this.code = r),
      (this.name = "ExternalSourceError"));
  }
}
function isExternalSourceError(error) {
  return error instanceof ExternalSourceError
    ? error
    : error instanceof Error &&
        error.name === "ExternalSourceError" &&
        externalSourceErrorCodes.includes(error.code)
      ? new ExternalSourceError(error.code)
      : new ExternalSourceError("unavailable");
}
function parseActiveSourceApproval(e, r = Date.now()) {
  const t = sourceApprovalSchema.safeParse(e);
  if (!t.success) throw new ExternalSourceError("invalid");
  if (Date.parse(t.data.expiresAt) <= r)
    throw new ExternalSourceError("expired");
  return t.data;
}
function isSameContentHash(source, r) {
  return fingerprintCanonicalValue(source) === fingerprintCanonicalValue(r);
}
function validateExternalSourcePayload(snapshot, source) {
  if (snapshot && typeof snapshot == "object") {
    const rows = snapshot.rows;
    if (
      Array.isArray(rows) &&
      (rows.length > externalSourceLimits.rows ||
        rows.some(
          (row) =>
            Array.isArray(row) && row.length > externalSourceLimits.columns,
        ))
    )
      throw new ExternalSourceError("limit");
  }
  const t = sourceResultPayloadSchema.safeParse(snapshot);
  if (!t.success) throw new ExternalSourceError("invalid");
  const data = t.data;
  if (source && !isSameContentHash(data.source, source))
    throw new ExternalSourceError("schema");
  if (
    (data.rows.length + 1) * data.source.columns.length >
      externalSourceLimits.cells ||
    new TextEncoder().encode(JSON.stringify(data)).length >
      externalSourceLimits.responseBytes
  )
    throw new ExternalSourceError("limit");
  for (const row of data.rows) {
    if (row.length !== data.source.columns.length)
      throw new ExternalSourceError("schema");
    row.forEach((a, s) => {
      const d = data.source.columns[s];
      if (
        a === null
          ? !d.nullable
          : d.type === "date"
            ? typeof a != "string" ||
              !/^\d{4}-\d{2}-\d{2}$/.test(a) ||
              !Number.isFinite(Date.parse(a)) ||
              new Date(a).toISOString().slice(0, 10) !== a
            : typeof a !==
              { text: "string", number: "number", boolean: "boolean" }[d.type]
      )
        throw new ExternalSourceError("schema");
    });
  }
  return data;
}
function computeSourceControlTotals(snapshot) {
  return snapshot.source.controlFields.map((controlField) => {
    const t = snapshot.source.columns.findIndex(
      (column) => column.name === controlField,
    );
    let n = 0,
      o = 0,
      numeric = 0,
      blanks = 0;
    for (const row of snapshot.rows) {
      if (row[t] === null) {
        blanks++;
        continue;
      }
      const u = row[t] - o,
        l = n + u;
      if (
        ((o = l - n - u),
        (n = l),
        numeric++,
        !Number.isFinite(n) || !Number.isFinite(o))
      )
        throw new ExternalSourceError("limit");
    }
    return {
      field: controlField,
      sum: Object.is(n, -0) ? 0 : n,
      numeric: numeric,
      blanks: blanks,
    };
  });
}
const sourceControlFieldSchema = strictObject({
    field: columnNameSchema,
    sum: number().finite(),
    numeric: number().int().min(0).max(5e3),
    blanks: number().int().min(0).max(5e3),
  }),
  sourceRefreshStateSchema = strictObject({
    protocol: literal(externalSourceFormatVersion),
    source: sourceTableSchema,
    lastSuccess: strictObject({
      at: isoDateTimeSchema,
      sourceTimestamp: isoDateTimeSchema,
      version: sourceIdSchema,
      rows: number().int().min(0).max(5e3),
      controls: array(sourceControlFieldSchema).max(50),
      dataHash: sourceHashSchema,
    }),
    lastAttempt: strictObject({
      at: isoDateTimeSchema,
      outcome: _enum(["succeeded", "failed", "cancelled"]),
      code: _enum(externalSourceErrorCodes).optional(),
    }),
    outputHash: sourceHashSchema,
  });
function buildExternalSourceRefreshState(workbook, r) {
  let t = 0,
    n = 0;
  for (const sheet of r.sheets) {
    const s = workbook.sheets.find(
      (l) => l.id === sheet.id,
    )?.externalConnection;
    if (s === undefined) continue;
    if (
      ++n > externalSourceLimits.sources ||
      (t += JSON.stringify(s).length) > externalSourceLimits.metadataCharacters
    )
      throw new ExternalSourceError("limit");
    const d = sourceRefreshStateSchema.safeParse(s);
    if (!d.success) throw new ExternalSourceError("invalid");
    const data = d.data;
    if (
      data.lastSuccess.controls.length !== data.source.controlFields.length ||
      data.lastSuccess.controls.some(
        (control, c) =>
          control.field !== data.source.controlFields[c] ||
          control.numeric + control.blanks !== data.lastSuccess.rows,
      ) ||
      (data.lastSuccess.rows + 1) * data.source.columns.length >
        externalSourceLimits.cells ||
      Date.parse(data.lastAttempt.at) < Date.parse(data.lastSuccess.at) ||
      (data.lastAttempt.outcome === "succeeded") == !!data.lastAttempt.code
    )
      throw new ExternalSourceError("invalid");
    sheet.externalConnection = data;
  }
}

export {
  assertCalculatedBeforeFilter,
  hasCalculatorRecord,
  deepFreeze,
  isSameContentHash,
  formatNumberWithPattern,
  normalizeNameForComparison,
  consolidationFormatVersion,
  filterEvaluationLimits,
  serializeTypedKeyValues,
  defaultCalculationSettings,
  makeFormulaError,
  createWorkbookFormulaEngine,
  savedFilterViewHelpText,
  getCurrentCalculatorRecord,
  matchesTextFilter,
  importFormulaCacheEvidence,
  isExternalSourceError,
  normalizeSavedFilterViewName,
  normalizeCalculationSettings,
  externalSourceFormatVersion,
  sumConsolidationAmounts,
  reportCellValueMaps,
  assertSavedFilterViewsEditable,
  ExternalSourceError,
  isValidFilterViewText,
  carryCalculatorRecordForward,
  applyDefinedNameOperation,
  resolveDatePolicyContext,
  withCalculationSettings,
  getCachedCalculationSnapshot,
  textFilterOperators,
  normalizeComparableText,
  _O,
  findSavedFilterView,
  validateExternalSourcePayload,
  buildExternalSourceRefreshState,
  carriedCalculatorRecords,
  comparableCellValueKey,
  externalSourceLimits,
  buildConsolidationSubtotals,
  parseActiveSourceApproval,
  ik,
  j5,
  isStorableCellValue,
  consolidationLimits,
  getCalculationInputSignature,
  parseConsolidationDefinition,
  pendingFormulaValue,
  isCalculationRecordCurrent,
  assertNamePreviewCurrent,
  buildReportCellValueMap,
  calculatorRecords,
  verifyConsolidationEvidence,
  validateSavedFilterViewMetadata,
  computeSourceControlTotals,
  savedFilterViewLimits,
  zie,
};

/* Also used by the workers. */
export {
  MO,
  boundedLabelSchema,
  buildSheetCalculationSnapshot,
  buildWorkbookCalculationDescriptor,
  buildWorkbookCalculationSignature,
  categoryValueSchema,
  columnNameSchema,
  consolidationDefinitionSchema,
  consolidationEvidenceSchema,
  consolidationIssueSchema,
  consolidationSourceResultSchema,
  consolidationSourceSchema,
  consolidationSourceSnapshotSchema,
  consolidationSubtotalSchema,
  consolidationTotalsSchema,
  dynamicFunctionPattern,
  externalSourceErrorCodes,
  externalSourceKinds,
  isoDateTimeSchema,
  matchCriteriaValue,
  scalarValueSchema,
  sourceApprovalSchema,
  sourceColumnSchema,
  sourceControlFieldSchema,
  sourceIdSchema,
  sourceRefreshStateSchema,
  sourceResultPayloadSchema,
  sourceTableSchema,
  throwInvalidCalculationSettings,
};

/* Also used by the workers. */
export { externalSourceErrorMessages };
