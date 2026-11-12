import { hasOnlyAllowedKeys } from "./19b-assistant-protocol.js";

import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { isNullOrEmptyValue, failWithMessage } from "./03-workbook-model.js";
import {
  parseFormulaTree,
  walkFormulaNodes,
  tokenizeFormulaText,
  fingerprintCanonicalValue,
  normalizeBindingName,
  normalizeDatePolicy,
} from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  assertWorkbookStructureEditable,
  serializeTypedKeyValues,
  createWorksheetRecord,
  sumConsolidationAmounts,
  calculateWorkbookWithSettings,
  _t,
  cn,
  parseRangeAddress,
  buildConsolidationSubtotals,
  consolidationLimits,
  parseConsolidationDefinition,
  resolveDefinedNameValue,
  normalizeWorkbookInput,
  columnIndexToLetters,
  matrixResultFunctionNames,
  rewriteStructuredReferences,
} from "./03-workbook-model.js";
import {
  requireIntegerInRange,
  singleRangeReferencePattern,
  buildCompositeKey,
  parseCellReference,
  bareSheetNamePattern,
  knownFunctionNameSet,
  isFormulaInputValue,
  sheetRangeReferencePattern,
} from "./11-local-workbook-store-and-versions.js";

function inspectFormulaDependencies(workbook, r) {
  if (!["precedents", "dependents"].includes(r.direction))
    throw Error("Choose precedents or dependents.");
  const depth = requireIntegerInRange(r.depth, 1, 8, "Inspection depth"),
    n = requireIntegerInRange(
      r.limits?.maxFormulaCells,
      5e3,
      1e4,
      "Formula scan limit",
    ),
    o = requireIntegerInRange(
      r.limits?.maxNodes,
      500,
      1e3,
      "Result cell limit",
    ),
    a = requireIntegerInRange(
      r.limits?.maxReferences,
      1e4,
      5e4,
      "Reference scan limit",
    );
  if (r.includeHidden !== undefined && typeof r.includeHidden != "boolean")
    throw Error("Include hidden must be boolean.");
  if (
    r.allowedSheetIds &&
    (!Array.isArray(r.allowedSheetIds) ||
      r.allowedSheetIds.some(
        (allowedSheetId) =>
          typeof allowedSheetId != "string" ||
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ))
  )
    throw Error("The allowed worksheet context is invalid.");
  const s = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet])),
    d = new Set(r.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id)),
    u = new Map(
      workbook.sheets.map((sheet) => [
        sheet.id,
        { rows: cn(sheet, "row"), cols: cn(sheet, "column") },
      ]),
    ),
    l = (Q, X) => {
      if (!d.has(Q.id)) return false;
      if (r.includeHidden) return true;
      const J = _t(X);
      return (
        !Q.hidden &&
        !u
          .get(Q.id)
          .rows.some((row) => J.row >= row.start && J.row <= row.end) &&
        !u.get(Q.id).cols.some((col) => J.col >= col.start && J.col <= col.end)
      );
    },
    cellReference = parseCellReference(r.address);
  if (cellReference.includes(":")) throw Error("Choose one cell to inspect.");
  const h = s.get(r.sheetId);
  if (!h || !l(h, cellReference))
    throw Error(
      "The selected cell is outside the visible, authorized inspection context.",
    );
  const compositeKey = buildCompositeKey(h.id, cellReference),
    issues = [],
    m = [],
    v = new Map(),
    g = new Set();
  let referencesScanned = 0,
    truncated = false,
    depthLimited = false,
    excludedCells = 0;
  const A = (from, X, J, token) => {
      (g.add(from),
        issues.length < 200 &&
        !issues.some(
          (issue) =>
            issue.from === from &&
            issue.kind === X &&
            issue.token === token &&
            issue.message === J,
        )
          ? issues.push({
              from: from,
              kind: X,
              message: J,
              ...(token ? { token: token } : {}),
            })
          : issues.length >= 200 && (truncated = true));
    },
    k = [];
  for (const sheet of workbook.sheets)
    for (const X of Object.entries(sheet.cells).sort(([J], [q]) => {
      const W = _t(J),
        K = _t(q);
      return W.row - K.row || W.col - K.col;
    })) {
      if (!l(sheet, X[0])) {
        excludedCells++;
        continue;
      }
      isFormulaInputValue(X[1]) &&
        k.push({
          key: buildCompositeKey(sheet.id, X[0]),
          sheet: sheet,
          address: X[0],
          raw: X[1].raw,
        });
    }
  k.sort((event, X) =>
    event.key === compositeKey ? -1 : X.key === compositeKey ? 1 : 0,
  );
  const F = k.slice(0, n);
  k.length > F.length &&
    ((truncated = true),
    A(
      compositeKey,
      "limit",
      "The formula-cell scan limit was reached; additional dependents may exist.",
    ));
  for (const cell of F) {
    const X = [];
    v.set(cell.key, X);
    let J = 0;
    const q = (sheetId, address, token, oe) => {
        if (referencesScanned >= a) {
          ((truncated = true),
            A(
              cell.key,
              "limit",
              "The workbook reference scan limit was reached.",
            ));
          return;
        }
        referencesScanned++;
        const de = s.get(sheetId);
        if (!de) {
          A(
            cell.key,
            "missing",
            "The referenced worksheet does not exist.",
            token,
          );
          return;
        }
        if (!d.has(sheetId) || (de.hidden && !r.includeHidden)) {
          A(
            cell.key,
            "excluded",
            "A referenced worksheet is outside the selected inspection context; its cells and values were excluded.",
          );
          return;
        }
        try {
          const he = parseCellReference(address),
            rangeAddress = parseRangeAddress(he);
          (X.push({
            from: cell.key,
            token: token,
            sheetId: sheetId,
            range: he,
            kind: oe,
            cellCount:
              (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
              (rangeAddress.lastCol - rangeAddress.firstCol + 1),
          }),
            !r.includeHidden &&
              (u
                .get(de.id)
                .rows.some(
                  (row) =>
                    row.start <= rangeAddress.lastRow &&
                    row.end >= rangeAddress.firstRow,
                ) ||
                u
                  .get(de.id)
                  .cols.some(
                    (col) =>
                      col.start <= rangeAddress.lastCol &&
                      col.end >= rangeAddress.firstCol,
                  )) &&
              A(
                cell.key,
                "excluded",
                "A referenced range intersects hidden rows or columns; hidden cell details are excluded.",
                token,
              ));
        } catch {
          A(
            cell.key,
            "missing",
            "The reference is outside Excel's address bounds or is invalid.",
            token,
          );
        }
      },
      W = (K, te, ae = [], oe) => {
        if (++J > 64 || ae.length > 16) {
          ((truncated = true),
            A(
              cell.key,
              "limit",
              "The defined-name expansion limit was reached.",
              oe,
            ));
          return;
        }
        let de = K.startsWith("=") ? 1 : 0;
        const he = new Map();
        if (/\b(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(K))
          try {
            walkFormulaNodes(tokenizeFormulaText(K), (Ae, ye, Te, Pe) => {
              (Pe ||
                (!Ae.qualifier && ye.has(normalizeBindingName(Ae.name)))) &&
                he.set(Ae.start, Ae.end);
            });
          } catch {
            A(
              cell.key,
              "unsupported",
              "Invalid LET/LAMBDA bindings prevent complete dependency inspection.",
            );
            return;
          }
        for (; de < K.length;) {
          if (he.has(de)) {
            de = he.get(de);
            continue;
          }
          if (referencesScanned >= a) {
            ((truncated = true),
              A(
                cell.key,
                "limit",
                "The workbook reference scan limit was reached.",
              ));
            return;
          }
          const Ae = K.slice(de),
            ye = K[de];
          if (ye === '"') {
            de++;
            let Je = false;
            for (; de < K.length;)
              if (K[de++] === '"')
                if (K[de] === '"') de++;
                else {
                  Je = true;
                  break;
                }
            Je ||
              A(
                cell.key,
                "unsupported",
                "Unterminated formula text cannot be inspected completely.",
              );
            continue;
          }
          const Te =
            /^(?:'(?:[^']|'')*\[[^\]]+\](?:[^']|'')*'\s*!\s*[^+*/(),;\s]+|\[[^\]]+\][^!+*/(),;]*![^+*/(),;\s]+)/.exec(
              Ae,
            );
          if (Te) {
            (A(
              cell.key,
              "external",
              "External workbook references are not opened or fetched.",
              Te[0],
            ),
              (de += Te[0].length));
            continue;
          }
          const Pe = sheetRangeReferencePattern.exec(Ae);
          if (Pe || /^'(?:[^']|'')*:(?:[^']|'')*'\s*!/.test(Ae)) {
            const Je = Pe?.[0] ?? /^[^+*/(),;]+/.exec(Ae)[0];
            (A(
              cell.key,
              "unsupported",
              "Three-dimensional worksheet references are not expanded.",
              Je,
            ),
              (de += Je.length));
            continue;
          }
          const qe = /^#[A-Z0-9/?!]+/i.exec(Ae);
          if (qe) {
            (A(
              cell.key,
              qe[0].toUpperCase() === "#REF!" ? "missing" : "unsupported",
              "The formula contains an error token.",
              qe[0],
            ),
              (de += qe[0].length));
            continue;
          }
          const ht = /^([A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(Ae);
          if (ht) {
            const Je = de;
            de += ht[0].length;
            let ze = 1;
            for (; de < K.length && ze;)
              (K[de] === "[" && ze++, K[de] === "]" && ze--, de++);
            const at = K.slice(Je, de),
              _e =
                !ze &&
                rewriteStructuredReferences(workbook, te, at, cell.address);
            _e
              ? q(_e.sheetId, _e.range, oe ?? at, oe ? "name" : "table")
              : A(
                  cell.key,
                  "unsupported",
                  "The table, field, or structured-reference form could not be resolved.",
                  at,
                );
            continue;
          }
          const Re =
            /^(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))\s*!\s*([A-Za-z_\\][A-Za-z0-9_.\\]*)/.exec(
              Ae,
            );
          if (Re && !/^[A-Za-z]{1,3}[1-9]\d*$/i.test(Re[3])) {
            const Je = (Re[1] ?? Re[2]).replace(/''/g, "'"),
              ze = workbook.sheets.find(
                (sheet) => sheet.name.toLowerCase() === Je.toLowerCase(),
              ),
              at = Re[0],
              _e = Re[3];
            if (!ze)
              A(
                cell.key,
                "missing",
                "The referenced worksheet does not exist.",
                at,
              );
            else if (!d.has(ze.id) || (ze.hidden && !r.includeHidden))
              A(
                cell.key,
                "excluded",
                "The defined name belongs to a worksheet outside the selected inspection context.",
              );
            else {
              const Ze =
                workbook.definedNames?.find(
                  (definedName) =>
                    definedName.scopeSheetId === ze.id &&
                    definedName.name.toLowerCase() === _e.toLowerCase(),
                ) ??
                workbook.definedNames?.find(
                  (definedName) =>
                    !definedName.scopeSheetId &&
                    definedName.name.toLowerCase() === _e.toLowerCase(),
                );
              if (!Ze)
                A(
                  cell.key,
                  "missing",
                  "The qualified defined name does not exist.",
                  at,
                );
              else {
                const lt = `${Ze.scopeSheetId ?? ""}:${Ze.name.toLowerCase()}`;
                ae.includes(lt)
                  ? A(
                      cell.key,
                      parseFormulaTree(Ze.refersTo) ? "dynamic" : "cycle",
                      "The reusable definition recurs; references are shown once.",
                      at,
                    )
                  : W(
                      Ze.refersTo,
                      Ze.scopeSheetId ?? ze.id,
                      [...ae, lt],
                      oe ?? at,
                    );
              }
            }
            de += at.length;
            continue;
          }
          const Me = /^([A-Za-z_\\][A-Za-z0-9_.\\]*)\s*\(/.exec(Ae);
          if (Me) {
            const Je = Me[1]
              .toUpperCase()
              .replace(/^_XLFN\./, "")
              .replace(/^_XLWS\./, "");
            if (["INDIRECT", "OFFSET"].includes(Je))
              A(
                cell.key,
                "dynamic",
                "Runtime reference selection is not resolved; static argument references are only potential dependencies.",
                Me[1],
              );
            else if (!knownFunctionNameSet.has(Je)) {
              const ze =
                workbook.definedNames?.find(
                  (definedName) =>
                    definedName.scopeSheetId === te &&
                    definedName.name.toUpperCase() === Je,
                ) ??
                workbook.definedNames?.find(
                  (definedName) =>
                    !definedName.scopeSheetId &&
                    definedName.name.toUpperCase() === Je,
                );
              if (ze) {
                const at = `${ze.scopeSheetId ?? ""}:${ze.name.toLowerCase()}`;
                ae.includes(at)
                  ? A(
                      cell.key,
                      parseFormulaTree(ze.refersTo) ? "dynamic" : "cycle",
                      "Recursive reusable formula: static references are shown once; execution depth is bounded by calculation.",
                      Me[1],
                    )
                  : W(
                      ze.refersTo,
                      ze.scopeSheetId ?? te,
                      [...ae, at],
                      oe ?? Me[1],
                    );
              } else
                A(
                  cell.key,
                  "unsupported",
                  "The function is outside the local supported catalog; its static references may be incomplete.",
                  Me[1],
                );
            }
            de += Me[0].length;
            continue;
          }
          const ke = singleRangeReferencePattern.exec(Ae);
          if (ke && !/[A-Za-z0-9_.!:]/.test(Ae[ke[0].length] ?? "")) {
            const Je = ke[1]?.trim().replace(/!\s*$/, "").trim(),
              ze = Je?.replace(/^'|'$/g, "").replace(/''/g, "'"),
              at =
                ze === undefined
                  ? s.get(te)
                  : workbook.sheets.find(
                      (sheet) => sheet.name.toLowerCase() === ze.toLowerCase(),
                    );
            (at
              ? q(
                  at.id,
                  ke[2],
                  oe ?? ke[0],
                  oe ? "name" : ke[3] ? "spill-anchor" : "a1",
                )
              : A(
                  cell.key,
                  "missing",
                  "The referenced worksheet does not exist.",
                  ke[0],
                ),
              ke[3] &&
                A(
                  cell.key,
                  "dynamic",
                  "Only the spill anchor is included; runtime spill output dependencies are not expanded.",
                  ke[0],
                ),
              (de += ke[0].length));
            continue;
          }
          const je = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(Ae);
          if (je) {
            de += je[0].length;
            continue;
          }
          const $e = bareSheetNamePattern.exec(Ae);
          if ($e) {
            const Je = $e[0],
              ze = Ae.slice(Je.length);
            if (/^\s*!/.test(ze)) {
              const _e = /^[^+*/(),;\s]+/.exec(Ae)[0];
              (A(
                cell.key,
                "unsupported",
                "Worksheet-qualified defined names are not resolved by this inspector.",
                _e,
              ),
                (de += _e.length));
              continue;
            }
            const at =
              workbook.definedNames?.find(
                (definedName) =>
                  definedName.scopeSheetId === te &&
                  definedName.name.toLowerCase() === Je.toLowerCase(),
              ) ??
              workbook.definedNames?.find(
                (definedName) =>
                  !definedName.scopeSheetId &&
                  definedName.name.toLowerCase() === Je.toLowerCase(),
              );
            if (at) {
              const _e = `${at.scopeSheetId ?? ""}:${at.name.toLowerCase()}`;
              ae.includes(_e)
                ? A(
                    cell.key,
                    "cycle",
                    "Defined names contain a circular reference.",
                    Je,
                  )
                : W(at.refersTo, at.scopeSheetId ?? te, [...ae, _e], oe ?? Je);
            } else
              ["TRUE", "FALSE"].includes(Je.toUpperCase()) ||
                A(
                  cell.key,
                  "missing",
                  "The identifier is not a defined name or supported reference.",
                  Je,
                );
            de += Je.length;
            continue;
          }
          ((ye === "'" ||
            ye === "!" ||
            ye === "[" ||
            ye === "]" ||
            ye === "@" ||
            ye === "|") &&
            A(
              cell.key,
              "unsupported",
              "The reference syntax is outside the inspected subset.",
              ye,
            ),
            de++);
        }
      };
    (W(cell.raw, cell.sheet.id), m.push(...X));
  }
  const I = new Map([
      [compositeKey, { sheet: h, address: cellReference, depth: 0 }],
    ]),
    O = new Map();
  let j = [compositeKey];
  const z = (sheet, address, J, from, W) => {
    if (!l(sheet, address)) return;
    const K = buildCompositeKey(sheet.id, address);
    if (!I.has(K)) {
      if (I.size - 1 >= o) {
        truncated = true;
        return;
      }
      I.set(K, { sheet: sheet, address: address, depth: J });
    }
    O.set(JSON.stringify([from, W]), { from: from, to: W });
  };
  for (let Q = 1; Q <= depth && j.length; Q++) {
    const X = new Set(I.keys());
    for (const J of j) {
      const q = I.get(J);
      if (r.direction === "precedents") {
        isFormulaInputValue(q.sheet.cells[q.address]) &&
          !v.has(J) &&
          ((truncated = true),
          A(J, "limit", "This formula was outside the formula scan budget."));
        for (const W of v.get(J) ?? []) {
          const rangeAddress = parseRangeAddress(W.range),
            te = s.get(W.sheetId);
          let ae = 0;
          e: for (
            let firstRow = rangeAddress.firstRow;
            firstRow <= rangeAddress.lastRow;
            firstRow++
          )
            for (
              let firstCol = rangeAddress.firstCol;
              firstCol <= rangeAddress.lastCol;
              firstCol++
            ) {
              if (++ae > o + 1 || I.size - 1 >= o) {
                truncated = true;
                break e;
              }
              const cellAddress = formatCellAddress(firstRow, firstCol);
              z(te, cellAddress, Q, J, buildCompositeKey(te.id, cellAddress));
            }
        }
      } else {
        const W = _t(q.address);
        for (const K of F)
          (v.get(K.key) ?? []).some((te) => {
            const rangeAddress = parseRangeAddress(te.range);
            return (
              te.sheetId === q.sheet.id &&
              W.row >= rangeAddress.firstRow &&
              W.row <= rangeAddress.lastRow &&
              W.col >= rangeAddress.firstCol &&
              W.col <= rangeAddress.lastCol
            );
          }) && z(K.sheet, K.address, Q, K.key, J);
      }
    }
    j = [...I.keys()].filter((J) => !X.has(J));
  }
  for (const Q of j) {
    const X = I.get(Q);
    (r.direction === "precedents"
      ? isFormulaInputValue(X.sheet.cells[X.address]) &&
        ((v.get(Q)?.length ?? 0) > 0 || !v.has(Q))
      : m.some((J) => {
          const q = _t(X.address),
            rangeAddress = parseRangeAddress(J.range);
          return (
            J.sheetId === X.sheet.id &&
            q.row >= rangeAddress.firstRow &&
            q.row <= rangeAddress.lastRow &&
            q.col >= rangeAddress.firstCol &&
            q.col <= rangeAddress.lastCol
          );
        })) && (depthLimited = true);
  }
  const edges = [...O.values()],
    cycles = [],
    B = new Map(),
    _ = (Q, X) => {
      if (B.get(Q) === 1) {
        const J = [...X.slice(X.indexOf(Q)), Q];
        cycles.length < 50 && cycles.push(J);
        return;
      }
      if (!(B.get(Q) === 2 || X.length > 1e3)) {
        B.set(Q, 1);
        for (const edge of edges) edge.from === Q && _(edge.to, [...X, Q]);
        B.set(Q, 2);
      }
    };
  for (const Q of I.keys()) _(Q, []);
  for (const cycle of cycles) for (const X of cycle) g.add(X);
  const C = (Q, X = new Set(), J = { cells: 0 }) => {
      if (X.has(Q) || X.size >= 32 || ++J.cells > 2e3 || g.has(Q)) return false;
      const [q, W] = JSON.parse(Q),
        K = s.get(q);
      if (!l(K, W)) return false;
      if (!isFormulaInputValue(K.cells[W])) return true;
      const te = v.get(Q);
      if (!te) return false;
      const ae = new Set(X).add(Q);
      for (const oe of te) {
        if (oe.cellCount > 2e3) return false;
        const rangeAddress = parseRangeAddress(oe.range);
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
            if (
              !C(
                buildCompositeKey(
                  oe.sheetId,
                  formatCellAddress(firstRow, firstCol),
                ),
                ae,
                J,
              )
            )
              return false;
      }
      return true;
    },
    T = {
      ...workbook,
      sheets: workbook.sheets
        .filter(
          (sheet) => d.has(sheet.id) && (r.includeHidden || !sheet.hidden),
        )
        .map((Q) => ({
          ...Q,
          cells: Object.fromEntries(
            Object.entries(Q.cells).filter(([X]) => l(Q, X)),
          ),
        })),
    },
    workbookWithSettings = calculateWorkbookWithSettings(T),
    R = (Q) => {
      const X = I.get(Q),
        cell = X.sheet.cells[X.address],
        q = {
          key: Q,
          sheetId: X.sheet.id,
          sheetName: X.sheet.name,
          address: X.address,
          raw: cell?.raw ?? null,
          depth: X.depth,
        };
      if (!isFormulaInputValue(cell))
        return {
          ...q,
          result: cell?.raw ?? null,
          resultStatus: "literal",
        };
      if (!C(Q))
        return {
          ...q,
          result: null,
          resultStatus: "unavailable",
          resultNote:
            "Result withheld: the static dependency closure is cyclic, unresolved, excluded, or beyond the inspection budget.",
        };
      try {
        return {
          ...q,
          result: workbookWithSettings(X.sheet.id, X.address),
          resultStatus: "calculated",
        };
      } catch {
        return {
          ...q,
          result: null,
          resultStatus: "unavailable",
          resultNote: "The local calculator could not evaluate this cell.",
        };
      }
    },
    notes = [
      "Static potential dependencies include all formula branches; they do not describe the executed branch or evaluation steps.",
      "Blank referenced cells are valid local precedents. Runtime spill outputs are not expanded.",
      "Visibility preferences are local UI controls. The embedding host must supply authorized workbook data and enforce access independently.",
    ];
  return (
    excludedCells &&
      notes.push(
        "Hidden or host-excluded populated cells were omitted from scanning and result details.",
      ),
    truncated &&
      notes.push(
        "One or more budgets were reached. Additional references, cells, issues, or dependents may exist.",
      ),
    depthLimited &&
      notes.push(
        "The requested traversal depth ended at cells with further potential relationships.",
      ),
    cycles.length &&
      notes.push(
        "Cycles were found in the displayed static graph; other cycles may exist beyond the selected depth or scan limits.",
      ),
    issues.length >= 200 &&
      notes.push("Only the first 200 distinct issues are listed."),
    {
      workbookId: workbook.id,
      revision: workbook.revision,
      direction: r.direction,
      depth: depth,
      target: R(compositeKey),
      nodes: [...I.keys()].filter((Q) => Q !== compositeKey).map(R),
      edges: edges,
      references: m.filter((Q) => I.has(Q.from)),
      issues: issues,
      cycles: cycles,
      coverage: {
        formulaCellsAvailable: k.length,
        formulaCellsScanned: F.length,
        referencesScanned: referencesScanned,
        excludedCells: excludedCells,
        truncated: truncated,
        depthLimited: depthLimited,
        complete:
          !truncated &&
          !depthLimited &&
          !issues.length &&
          !cycles.length &&
          !excludedCells,
        notes: notes,
      },
    }
  );
}
const fuzzyMatchLimits = Object.freeze({
  fuzzyCharacters: 512,
  comparisonWork: 1e7,
  matchPairs: 1e3,
});
function parseKeyNormalization(normalization, fuzzy, t) {
  if (
    normalization !== undefined &&
    (!hasOnlyAllowedKeys(normalization, [
      "trim",
      "collapseWhitespace",
      "ignoreCase",
      "unicode",
    ]) ||
      [
        normalization.trim,
        normalization.collapseWhitespace,
        normalization.ignoreCase,
      ].some((n) => n !== undefined && typeof n != "boolean") ||
      (normalization.unicode !== undefined &&
        !["NFC", "NFKC"].includes(normalization.unicode)))
  )
    throw Error("Invalid explicit key normalization options.");
  if (
    fuzzy !== undefined &&
    (!hasOnlyAllowedKeys(fuzzy, ["key", "threshold"]) ||
      !Number.isInteger(fuzzy.key) ||
      fuzzy.key < 0 ||
      fuzzy.key >= t ||
      typeof fuzzy.threshold != "number" ||
      !Number.isFinite(fuzzy.threshold) ||
      fuzzy.threshold < 0 ||
      fuzzy.threshold > 1)
  )
    throw Error(
      "Choose a fuzzy key pair and a similarity threshold from 0 to 1.",
    );
  return {
    ...(normalization ? { normalization: { ...normalization } } : {}),
    ...(fuzzy ? { fuzzy: { ...fuzzy } } : {}),
  };
}
function normalizeKeyText(e, normalization = {}) {
  if (typeof e != "string") return e;
  let t = normalization.unicode ? e.normalize(normalization.unicode) : e;
  return (
    normalization.collapseWhitespace && (t = t.replace(/\s+/gu, " ")),
    normalization.trim && (t = t.trim()),
    normalization.ignoreCase && (t = t.toLocaleLowerCase("en-US")),
    t
  );
}

function matchKeyPopulations(records, r, t, n, normalization, fuzzy) {
  parseKeyNormalization(normalization, fuzzy, t.length);
  const diagnostics = {
    policy: fuzzy
      ? `Normalized Unicode code-point Levenshtein similarity: 1 - edit distance / maximum text length. Fuzzy key ${fuzzy.key + 1}; inclusive threshold ${fuzzy.threshold}. Other keys compare typed values exactly after selected text normalization. All qualifying matches are retained; scores are not probabilities or Power Query equivalents.`
      : "Exact typed composite keys after explicitly selected text normalization. Blank keys never match; numbers and booleans are not coerced.",
    normalizedKeys: [],
    exceptions: [],
    matches: [],
    ambiguousLeftRows: 0,
    ambiguousRightRows: 0,
    normalizationCollisionGroups: { left: 0, right: 0 },
    comparisonWork: 0,
  };
  let d = 0;
  const u = (C, T, side) =>
      C.map((record) => {
        const original = T.map((X) => record.values[X]),
          $ = original.map((X) => normalizeKeyText(X, normalization));
        if (
          ((d += $.reduce(
            (X, J) => X + (typeof J == "string" ? J.length : 0),
            0,
          )),
          d > 2e6)
        )
          throw Error("Normalized key text exceeds 2,000,000 characters.");
        serializeTypedKeyValues(original) !== serializeTypedKeyValues($) &&
          diagnostics.normalizedKeys.push({
            side: side,
            locator: record.locator,
            original: original,
            normalized: $,
          });
        const isBlank = $.some(isNullOrEmptyValue),
          G = fuzzy ? $[fuzzy.key] : undefined,
          H = !!fuzzy && !isBlank && typeof G != "string";
        H &&
          diagnostics.exceptions.push({
            side: side,
            locator: record.locator,
            reason:
              "Fuzzy matching requires text in the selected key. This record is unmatched; no numeric or Boolean coercion was applied.",
          });
        const chars = typeof G == "string" ? Array.from(G) : [];
        if (fuzzy && chars.length > fuzzyMatchLimits.fuzzyCharacters)
          throw Error(
            "Fuzzy keys are limited to 512 Unicode characters. Shorten the reviewed key or use exact matching.",
          );
        return {
          record: record,
          original: original,
          values: $,
          valid: !isBlank && !H,
          isBlank: isBlank,
          key: serializeTypedKeyValues($),
          block: serializeTypedKeyValues($.filter((X, J) => J !== fuzzy?.key)),
          chars: chars,
        };
      }),
    l = u(records, t, "left"),
    c = u(r, n, "right"),
    h = (C, T) => {
      const N = new Map();
      for (const R of C)
        R.isBlank || N.set(R.key, [...(N.get(R.key) ?? []), R]);
      return (
        (diagnostics.normalizationCollisionGroups[T] = [...N.values()].filter(
          (R) =>
            new Set(R.map((L) => serializeTypedKeyValues(L.original))).size > 1,
        ).length),
        N
      );
    },
    f = h(l, "left"),
    p = h(c, "right"),
    byLeft = new Map(),
    v = new Map(),
    g = new Map(),
    b = new Map(),
    y = (C) => {
      if (
        ((diagnostics.comparisonWork += C),
        diagnostics.comparisonWork > fuzzyMatchLimits.comparisonWork)
      )
        throw Error(
          "Fuzzy comparison exceeds 10,000,000 work units. Narrow the populations or add an exact key before previewing.",
        );
    };
  function w(C, T) {
    if ((y(C.length + T.length + 1), C.join("") === T.join(""))) return 1;
    y(C.length * T.length);
    let N = Array.from({ length: T.length + 1 }, (R, L) => L);
    for (let R = 1; R <= C.length; R++) {
      const L = [R];
      for (let $ = 1; $ <= T.length; $++)
        L[$] = Math.min(
          N[$] + 1,
          L[$ - 1] + 1,
          N[$ - 1] + (C[R - 1] === T[$ - 1] ? 0 : 1),
        );
      N = L;
    }
    return 1 - N[T.length] / Math.max(C.length, T.length);
  }
  const E = new Map();
  for (const C of c)
    if (C.valid) {
      const T = fuzzy ? C.block : C.key;
      E.set(T, [...(E.get(T) ?? []), C]);
    }
  for (const C of l) {
    if (!C.valid) continue;
    const T = E.get(fuzzy ? C.block : C.key) ?? [];
    for (const N of T) {
      const score = fuzzy ? w(C.chars, N.chars) : 1;
      if (
        (g.set(C.record.row, Math.max(g.get(C.record.row) ?? -1, score)),
        b.set(N.record.row, Math.max(b.get(N.record.row) ?? -1, score)),
        fuzzy && score < fuzzy.threshold)
      )
        continue;
      const explanation = fuzzy
          ? `Text similarity ${score.toFixed(6)} meets threshold ${fuzzy.threshold}; other key pairs match exactly after selected normalization.`
          : "Exact typed keys match after selected normalization.",
        $ = byLeft.get(C.record.row) ?? [];
      if (
        ($.push({ record: N.record, score: score, explanation: explanation }),
        byLeft.set(C.record.row, $),
        v.set(N.record.row, (v.get(N.record.row) ?? 0) + 1),
        diagnostics.matches.push({
          left: C.record.locator,
          right: N.record.locator,
          score: score,
          explanation: explanation,
        }),
        diagnostics.matches.length > fuzzyMatchLimits.matchPairs)
      )
        throw Error(
          "Matching duplicate keys expand beyond 1,000 result rows or reviewable match pairs. Narrow or reconcile the source keys first; no pairs were truncated.",
        );
    }
  }
  ((diagnostics.ambiguousLeftRows = [...byLeft.values()].filter(
    (C) => C.length > 1,
  ).length),
    (diagnostics.ambiguousRightRows = [...v.values()].filter(
      (C) => C > 1,
    ).length));
  const A = (C) => JSON.stringify(C.map((T) => String(T))),
    k = new Set(l.filter((C) => C.valid).map((C) => A(C.values))),
    F = new Set(c.filter((C) => C.valid).map((C) => A(C.values))),
    I = (C, T, N) =>
      C.isBlank
        ? "Blank join key; blanks do not match."
        : C.valid
          ? fuzzy
            ? T === undefined
              ? "No eligible text candidate with matching exact keys."
              : `Best eligible similarity ${T.toFixed(6)} is below threshold ${fuzzy.threshold}.`
            : N.has(A(C.values))
              ? "No exact typed composite-key match; equivalent displayed keys have mismatched types. No coercion was applied."
              : "No exact typed composite-key match."
          : "Non-text fuzzy key; no implicit type conversion.",
    unmatchedLeft = l
      .filter((C) => !byLeft.has(C.record.row))
      .map((C) => ({
        ...C.record,
        unmatchedReason: I(C, g.get(C.record.row), F),
      })),
    unmatchedRight = c
      .filter((C) => !v.has(C.record.row))
      .map((C) => ({
        ...C.record,
        unmatchedReason: I(C, b.get(C.record.row), k),
      }));
  for (const [side, T] of [
    ["left", unmatchedLeft],
    ["right", unmatchedRight],
  ])
    for (const N of T)
      diagnostics.exceptions.some(
        (exception) =>
          exception.side === side && exception.locator === N.locator,
      ) ||
        diagnostics.exceptions.push({
          side: side,
          locator: N.locator,
          reason: N.unmatchedReason,
        });
  const z = new Map(),
    V = new Map();
  for (const match of diagnostics.matches)
    (V.set(match.left, [...(V.get(match.left) ?? []), match.right]),
      z.set(match.right, [...(z.get(match.right) ?? []), match.left]));
  const D = new Set();
  let manyToManyKeys = 0,
    expandedMatchRows = 0;
  for (const C of V.keys())
    if (!D.has(C)) {
      const T = new Set(),
        N = new Set(),
        R = [C];
      let L = 0;
      for (; R.length;) {
        const $ = R.pop();
        if (D.has($)) continue;
        (D.add($), T.add($));
        const U = V.get($);
        L += U.length;
        for (const G of U) {
          N.add(G);
          for (const H of z.get(G)) D.has(H) || R.push(H);
        }
      }
      T.size > 1 &&
        N.size > 1 &&
        (manyToManyKeys++,
        (expandedMatchRows += Math.max(0, L - Math.max(T.size, N.size))));
    }
  return {
    byLeft: byLeft,
    unmatchedLeft: unmatchedLeft,
    unmatchedRight: unmatchedRight,
    diagnostics: diagnostics,
    duplicateLeftKeys: [...f.values()].filter((C) => C.length > 1).length,
    duplicateRightKeys: [...p.values()].filter((C) => C.length > 1).length,
    matchedLeftRows: byLeft.size,
    matchedRightRows: v.size,
    matchPairs: diagnostics.matches.length,
    manyToManyKeys: manyToManyKeys,
    expandedMatchRows: expandedMatchRows,
  };
}
const consolidationBudgets = {
  sourceCells: 2e3,
  sourceRows: 500,
  sourceColumns: 50,
  outputRows: 1e3,
  outputCells: 2e4,
  keyColumns: 8,
  sourceFormulas: 50,
  scannedFormulas: 1e3,
  dependencyCells: 2e3,
  headerCharacters: 128,
  textCharacters: 32767,
  totalTextCharacters: 1e6,
};

function parseSourceRectangle(e) {
  (!hasOnlyAllowedKeys(e, ["sheetId", "range"]) ||
    typeof e.sheetId != "string" ||
    typeof e.range != "string" ||
    !/^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/.test(
      e.range,
    )) &&
    failWithMessage(
      "Choose a worksheet and one explicit headed cell rectangle for each source.",
    );
  const r = e.range.replace(/\$/g, "").toUpperCase(),
    t = r.split(":"),
    n = _t(t[0]),
    o = _t(t[1] ?? t[0]);
  (n.row > o.row ||
    n.col > o.col ||
    o.row >= workbookLimits.maxRows ||
    o.col >= workbookLimits.maxCols) &&
    failWithMessage(
      "A source rectangle is reversed or outside Excel's address bounds.",
    );
  const a = o.row - n.row + 1,
    s = o.col - n.col + 1;
  return (
    (a > consolidationBudgets.sourceRows + 1 ||
      s > consolidationBudgets.sourceColumns ||
      a * s > consolidationBudgets.sourceCells) &&
      failWithMessage(
        "Each source supports at most 500 data rows, 50 columns and 2,000 cells including headers.",
      ),
    {
      sheetId: e.sheetId,
      range: `${formatCellAddress(n.row, n.col)}:${formatCellAddress(o.row, o.col)}`,
    }
  );
}
function isCellVisible(e, firstRow, firstCol) {
  return (
    !e.hidden &&
    !e.veryHidden &&
    !cn(e, "row").some((n) => firstRow >= n.start && firstRow <= n.end) &&
    !cn(e, "column").some((n) => firstCol >= n.start && firstCol <= n.end)
  );
}
function inspectOutputNameSafety(workbook, r) {
  let t = 0;
  const n = (o, a, address = "A1") => {
    if (!o) return;
    ((t += o.length),
      t > consolidationBudgets.totalTextCharacters &&
        failWithMessage(
          "Output-name safety inspection exceeds 1,000,000 formula characters. Use a smaller reviewed workbook.",
        ));
    const d = o.replace(/"(?:[^"]|"")*"/g, "");
    let u = "",
      l = 0;
    for (; l < d.length;) {
      if (d[l] === "'") {
        const f = l++;
        for (; l < d.length;)
          if (d[l++] === "'")
            if (d[l] === "'") l++;
            else break;
        u += d.slice(f, l);
        continue;
      }
      const c = l,
        h = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(d.slice(l));
      if (h) {
        l += h[0].length;
        let f = 1;
        for (; l < d.length && f;)
          (d[l] === "[" && f++, d[l] === "]" && f--, l++);
        const p = d.slice(c, l);
        u +=
          !f &&
          !/^[A-Za-z0-9_!]/.test(d.slice(l)) &&
          rewriteStructuredReferences(workbook, a, p, address)
            ? "0"
            : p;
      } else u += d[l++];
    }
    for (const c of u.matchAll(
      /('(?:[^']|'')*'|[A-Za-z_.\d\u007F-\uFFFF]+)\s*!/g,
    )) {
      const h = c[1];
      (h.startsWith("'")
        ? h.slice(1, -1).replace(/''/g, "'")
        : h
      ).toLowerCase() === r.toLowerCase() &&
        failWithMessage(
          "The output name would reconnect an existing unresolved worksheet reference. Choose a different name.",
        );
    }
  };
  for (const sheet of workbook.sheets) {
    for (const [a, cell] of Object.entries(sheet.cells))
      isFormulaInputValue(cell) && n(String(cell.raw), sheet.id, a);
    for (const a of sheet.conditionalFormats ?? [])
      a.kind === "formula" && n(a.formula, sheet.id, a.range.split(":")[0]);
    for (const a of sheet.tables ?? [])
      for (const [s, d] of a.columns.entries()) {
        const rangeAddress = parseRangeAddress(a.range);
        n(
          d.calculatedFormula,
          sheet.id,
          formatCellAddress(
            rangeAddress.firstRow + 1,
            rangeAddress.firstCol + s,
          ),
        );
      }
  }
  for (const o of workbook.definedNames ?? [])
    n(o.refersTo, o.scopeSheetId ?? workbook.sheets[0].id);
}
function readSourceRectangles(workbookInput, r, t = {}) {
  (!Array.isArray(r) || !r.length || r.length > 8) &&
    failWithMessage("Read one to eight explicit source rectangles.");
  const n = r.map(parseSourceRectangle);
  (t.includeHidden !== undefined &&
    typeof t.includeHidden != "boolean" &&
    failWithMessage("Include hidden must be boolean."),
    t.allowedSheetIds !== undefined &&
      (!Array.isArray(t.allowedSheetIds) ||
        t.allowedSheetIds.some(
          (allowedSheetId) =>
            typeof allowedSheetId != "string" ||
            !workbookInput.sheets.some((sheet) => sheet.id === allowedSheetId),
        )) &&
      failWithMessage("The host scope contains an unknown worksheet."));
  const o = {
      includeHidden: !!t.includeHidden,
      allowedSheetIds: [
        ...new Set(
          t.allowedSheetIds ?? workbookInput.sheets.map((sheet) => sheet.id),
        ),
      ].sort(),
    },
    a = n.map((y) => {
      const w = workbookInput.sheets.find((sheet) => sheet.id === y.sheetId);
      (!w || !o.allowedSheetIds.includes(w.id)) &&
        failWithMessage("A source is outside the authorized host context.");
      const rangeAddress = parseRangeAddress(y.range);
      if (!o.includeHidden)
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
            isCellVisible(w, firstRow, firstCol) ||
              failWithMessage(
                "A source includes hidden cells. Explicitly include hidden records before previewing.",
              );
      return (
        w.merges?.some((merge) => {
          const F = parseRangeAddress(merge);
          return (
            F.firstRow <= rangeAddress.lastRow &&
            F.lastRow >= rangeAddress.firstRow &&
            F.firstCol <= rangeAddress.lastCol &&
            F.lastCol >= rangeAddress.firstCol
          );
        }) &&
          failWithMessage(
            "Merged source cells are ambiguous; use an unmerged headed population.",
          ),
        w
      );
    }),
    s = new Set(),
    d = new Set(),
    u = [];
  for (const [y, w] of n.entries()) {
    const sheet = a[y],
      rangeAddress = parseRangeAddress(w.range);
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
        const cellAddress = formatCellAddress(firstRow, firstCol),
          compositeKey = buildCompositeKey(sheet.id, cellAddress);
        (s.add(compositeKey),
          isFormulaInputValue(sheet.cells[cellAddress]) &&
            !d.has(compositeKey) &&
            (d.add(compositeKey),
            u.push({ sheet: sheet, address: cellAddress })));
      }
  }
  u.length > consolidationBudgets.sourceFormulas &&
    failWithMessage(
      "Source capture supports at most 50 formula source cells. Use a reviewed values snapshot for a larger calculated population.",
    );
  for (const y of new Set(a)) {
    const w = [];
    for (const A of s) {
      const [k, F] = JSON.parse(A);
      k === y.id && (!y.cells[F] || y.cells[F].raw === null) && w.push(_t(F));
    }
    if (!w.length) continue;
    const E = Object.entries(y.cells).filter(
      ([A, k]) => isFormulaInputValue(k) && !d.has(buildCompositeKey(y.id, A)),
    );
    E.length > consolidationBudgets.scannedFormulas &&
      failWithMessage(
        "Blank source cells require a bounded spill-ambiguity check; this sheet exceeds 1,000 possible formula anchors.",
      );
    for (const [A] of E) {
      const k = _t(A);
      w.some(
        (F) =>
          F.row >= k.row &&
          F.col >= k.col &&
          (F.row - k.row + 1) * (F.col - k.col + 1) <= 1e4,
      ) &&
        failWithMessage(
          "A blank source cell could be a spill child of a formula outside the source. Use a reviewed values snapshot or include and validate the scalar source formula.",
        );
    }
  }
  const l = new Set([
      ...matrixResultFunctionNames,
      "ISFORMULA",
      "INDEX",
      "CHOOSE",
      "INDIRECT",
      "OFFSET",
      "LET",
      "LAMBDA",
      "RAND",
      "RANDBETWEEN",
      "NOW",
      "TODAY",
      "CELL",
      "INFO",
      "RTD",
      "WEBSERVICE",
      "STOCKHISTORY",
    ]),
    workbook = {
      ...workbookInput,
      sheets: workbookInput.sheets.map((sheet) =>
        sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
      ),
    },
    h = new Set();
  for (const y of u) {
    const w = inspectFormulaDependencies(workbook, {
        sheetId: y.sheet.id,
        address: y.address,
        direction: "precedents",
        depth: 8,
        includeHidden: o.includeHidden,
        allowedSheetIds: o.allowedSheetIds,
        limits: { maxFormulaCells: 1e3, maxNodes: 500, maxReferences: 2e3 },
      }),
      E = [w.target, ...w.nodes],
      A = new Set(E.map((event) => event.key));
    (w.coverage.truncated ||
      w.coverage.depthLimited ||
      w.cycles.length ||
      w.issues.some((issue) => A.has(issue.from))) &&
      failWithMessage(
        "A source formula has incomplete, cyclic, hidden, unauthorized or unsupported dependencies.",
      );
    for (const k of E) {
      (s.add(k.key),
        h.add(k.key),
        h.size > consolidationBudgets.dependencyCells &&
          failWithMessage(
            "Combined source formula dependencies exceed 2,000 cells.",
          ));
      const F = workbookInput.sheets.find((sheet) => sheet.id === k.sheetId),
        cell = F.cells[k.address];
      if (
        ((!cell || cell.raw === null) &&
          failWithMessage(
            "Formula dependencies must be populated static cells; blank or possible spill-child dependencies are outside this subset.",
          ),
        isFormulaInputValue(cell))
      ) {
        const O = resolveDefinedNameValue(
          String(cell.raw),
          workbookInput,
          F.id,
        ).replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, "");
        ([...O.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)]
          .map((z) =>
            z[1]
              .toUpperCase()
              .replace(/^_XLFN\./, "")
              .replace(/^_XLWS\./, ""),
          )
          .some((z) => l.has(z)) ||
          /[{}#]/.test(O)) &&
          failWithMessage(
            "Source formulas must have deterministic scalar dependencies; volatile, dynamic and array formulas are excluded.",
          );
      }
    }
  }
  const f = {
      ...workbookInput,
      sheets: workbookInput.sheets
        .filter((sheet) => o.allowedSheetIds.includes(sheet.id))
        .map((y) => ({
          ...y,
          cells: Object.fromEntries(
            Object.entries(y.cells).filter(([w]) =>
              s.has(buildCompositeKey(y.id, w)),
            ),
          ),
        })),
    },
    workbookWithSettings = calculateWorkbookWithSettings(f),
    m =
      /^#(?:DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|CYCLE!|LIMIT!|BLOCKED!|ERROR!)$/;
  let v = 0;
  const grids = n.map((source, w) => {
    const E = a[w],
      rangeAddress = parseRangeAddress(source.range),
      values = [];
    for (
      let firstRow = rangeAddress.firstRow;
      firstRow <= rangeAddress.lastRow;
      firstRow++
    ) {
      const I = [];
      for (
        let firstCol = rangeAddress.firstCol;
        firstCol <= rangeAddress.lastCol;
        firstCol++
      ) {
        const cellAddress = formatCellAddress(firstRow, firstCol),
          cell = E.cells[cellAddress],
          V = isFormulaInputValue(cell)
            ? workbookWithSettings(E.id, cellAddress)
            : (cell?.raw ?? null);
        (((isFormulaInputValue(cell) && typeof V == "string" && m.test(V)) ||
          (typeof V == "number" && !Number.isFinite(V)) ||
          (typeof V == "string" &&
            V.length > consolidationBudgets.textCharacters)) &&
          failWithMessage(
            "A source contains a calculation error, nonfinite number or oversized text. Repair it before capturing source values.",
          ),
          (v += typeof V == "string" ? V.length : 0),
          v > consolidationBudgets.totalTextCharacters &&
            failWithMessage(
              "Combined source text exceeds 1,000,000 characters.",
            ),
          I.push(V));
      }
      values.push(I);
    }
    return { source: source, sheetName: E.name, values: values };
  });
  for (const y of h) {
    const [w, E] = JSON.parse(y);
    workbookWithSettings.getSpillRange(w, E) &&
      failWithMessage(
        "Source formula dependencies produced an array; use a reviewed scalar population.",
      );
  }
  workbookWithSettings.getOperationalErrors().length &&
    failWithMessage(
      "A source encountered an operational calculation failure, including an error hidden by a formula handler.",
    );
  const dependencyCells = [...s].sort().map((y) => {
    const [sheetId, address] = JSON.parse(y);
    return {
      sheetId: sheetId,
      address: address,
      cell:
        workbookInput.sheets.find((sheet) => sheet.id === sheetId)?.cells[
          address
        ] ?? null,
    };
  });
  return { grids: grids, dependencyCells: dependencyCells };
}
function readSourceHeaders(workbookInput, r, t = {}) {
  const { grids, dependencyCells } = readSourceRectangles(workbookInput, r, t);
  return {
    snapshots: grids.map(({ source, sheetName, values }) => {
      const rangeAddress = parseRangeAddress(source.range);
      values[0].some(
        (h) =>
          typeof h != "string" ||
          !h.trim() ||
          h.length > consolidationBudgets.headerCharacters,
      ) &&
        failWithMessage(
          "Every source column needs a nonblank text header of at most 128 characters.",
        );
      const headers = values[0];
      return (
        new Set(headers.map((header) => header.toLowerCase())).size !==
          headers.length &&
          failWithMessage(
            "Source headers must be unique, including case variants.",
          ),
        {
          source,
          sheetName,
          headers: headers,
          records: values.slice(1).map((h, f) => ({
            row: rangeAddress.firstRow + f + 2,
            locator: `'${sheetName.replace(/'/g, "''")}'!${formatCellAddress(rangeAddress.firstRow + f + 1, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.firstRow + f + 1, rangeAddress.lastCol)}`,
            values: h,
          })),
        }
      );
    }),
    dependencyCells,
  };
}
function buildConsolidationOutput(workbook, request, t = {}) {
  (t.readOnly && failWithMessage("This workbook is read-only."),
    assertWorkbookStructureEditable(workbook),
    (!hasOnlyAllowedKeys(request, [
      "left",
      "right",
      "mode",
      "keys",
      "outputSheetName",
      "normalization",
      "fuzzy",
    ]) ||
      ![
        "append",
        "inner",
        "left",
        "right",
        "full",
        "leftAnti",
        "rightAnti",
      ].includes(request.mode)) &&
      failWithMessage("Choose append or a supported join operation."),
    (typeof request.outputSheetName != "string" ||
      request.outputSheetName !== request.outputSheetName.trim() ||
      !request.outputSheetName ||
      request.outputSheetName.length > 31 ||
      /[\\/\[\]*?:\x00-\x1f]/.test(request.outputSheetName) ||
      /^'|'$/.test(request.outputSheetName)) &&
      failWithMessage(
        "Output sheet name must be a valid, unpadded name of 1–31 characters.",
      ),
    workbook.sheets.some(
      (sheet) =>
        sheet.name.toLowerCase() === request.outputSheetName.toLowerCase(),
    ) &&
      failWithMessage(
        "The output sheet name already exists. Choose a new sheet name.",
      ),
    inspectOutputNameSafety(workbook, request.outputSheetName),
    t.includeHidden !== undefined &&
      typeof t.includeHidden != "boolean" &&
      failWithMessage("Include hidden must be boolean."),
    t.allowedSheetIds !== undefined &&
      (!Array.isArray(t.allowedSheetIds) ||
        t.allowedSheetIds.some(
          (allowedSheetId) =>
            typeof allowedSheetId != "string" ||
            !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
        )) &&
      failWithMessage("The host scope contains an unknown worksheet."));
  const context = {
      includeHidden: !!t.includeHidden,
      allowedSheetIds: [
        ...new Set(
          t.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
        ),
      ].sort(),
    },
    o = {
      left: parseSourceRectangle(request.left),
      right: parseSourceRectangle(request.right),
      mode: request.mode,
      outputSheetName: request.outputSheetName,
    };
  o.mode === "append"
    ? ((request.normalization !== undefined || request.fuzzy !== undefined) &&
        failWithMessage(
          "Append aligns headers; key normalization and fuzzy matching apply only to joins.",
        ),
      request.keys !== undefined &&
        (!Array.isArray(request.keys) || request.keys.length) &&
        failWithMessage(
          "Append matches headers; it does not accept join keys.",
        ))
    : ((!Array.isArray(request.keys) ||
        !request.keys.length ||
        request.keys.length > consolidationBudgets.keyColumns ||
        request.keys.some(
          (key) =>
            !hasOnlyAllowedKeys(key, ["left", "right"]) ||
            typeof key.left != "string" ||
            typeof key.right != "string",
        )) &&
        failWithMessage("Choose from one to eight pairs of join headers."),
      (o.keys = request.keys.map((key) => ({ ...key }))),
      Object.assign(
        o,
        parseKeyNormalization(
          request.normalization,
          request.fuzzy,
          o.keys.length,
        ),
      ),
      (new Set(o.keys.map((key) => key.left)).size !== o.keys.length ||
        new Set(o.keys.map((key) => key.right)).size !== o.keys.length) &&
        failWithMessage("A join header cannot be repeated in the key pairs."));
  const { snapshots } = readSourceHeaders(workbook, [o.left, o.right], t),
    [left, right] = snapshots,
    rows = [],
    unmatchedLeft = [],
    unmatchedRight = [],
    keyProfiles = [];
  let matching,
    headers,
    matchedLeftRows = 0,
    matchedRightRows = 0,
    matchPairs = 0,
    duplicateLeftKeys = 0,
    duplicateRightKeys = 0,
    manyToManyKeys = 0,
    expandedMatchRows = 0;
  if (o.mode === "append") {
    ((left.headers.length !== right.headers.length ||
      left.headers.some((header) => !right.headers.includes(header))) &&
      failWithMessage(
        "Append requires exactly matching header names on both sources; column order may differ.",
      ),
      (headers = [...left.headers, "Source left row", "Source right row"]));
    const z = left.headers.map((header) => right.headers.indexOf(header));
    (left.records.forEach((record) =>
      rows.push({
        values: [...record.values],
        left: record.locator,
        right: null,
      }),
    ),
      right.records.forEach((record) =>
        rows.push({
          values: z.map((D) => record.values[D]),
          left: null,
          right: record.locator,
        }),
      ));
  } else {
    headers = [
      ...left.headers.map((header) => `Left: ${header}`),
      ...right.headers.map((header) => `Right: ${header}`),
      "Source left row",
      "Source right row",
    ];
    const z = o.keys.map((key) => left.headers.indexOf(key.left)),
      V = o.keys.map((key) => right.headers.indexOf(key.right));
    (z.includes(-1) || V.includes(-1)) &&
      failWithMessage(
        "Every join key must name an existing source header exactly.",
      );
    const D = (_, C) =>
      [
        ...new Set(
          _.map((T) =>
            T.values[C] === null || T.values[C] === ""
              ? "blank"
              : typeof T.values[C],
          ),
        ),
      ].sort();
    o.keys.forEach((key, C) =>
      keyProfiles.push({
        leftHeader: key.left,
        rightHeader: key.right,
        leftTypes: D(left.records, z[C]),
        rightTypes: D(right.records, V[C]),
      }),
    );
    const B = matchKeyPopulations(
      left.records,
      right.records,
      z,
      V,
      o.normalization,
      o.fuzzy,
    );
    ((matching = B.diagnostics),
      ({
        matchedLeftRows: matchedLeftRows,
        matchedRightRows: matchedRightRows,
        matchPairs: matchPairs,
        duplicateLeftKeys: duplicateLeftKeys,
        duplicateRightKeys: duplicateRightKeys,
        manyToManyKeys: manyToManyKeys,
        expandedMatchRows: expandedMatchRows,
      } = B),
      unmatchedLeft.push(...B.unmatchedLeft),
      unmatchedRight.push(...B.unmatchedRight));
    for (const record of left.records) {
      const C = B.byLeft.get(record.row);
      if (C) {
        if (o.mode !== "leftAnti" && o.mode !== "rightAnti")
          for (const T of C)
            rows.push({
              values: [...record.values, ...T.record.values],
              left: record.locator,
              right: T.record.locator,
              ...(o.fuzzy || o.normalization
                ? { match: { score: T.score, explanation: T.explanation } }
                : {}),
            });
      } else
        ["left", "full", "leftAnti"].includes(o.mode) &&
          rows.push({
            values: [...record.values, ...right.headers.map(() => null)],
            left: record.locator,
            right: null,
          });
    }
    if (["right", "full", "rightAnti"].includes(o.mode))
      for (const _ of unmatchedRight)
        rows.push({
          values: [...left.headers.map(() => null), ..._.values],
          left: null,
          right: _.locator,
        });
  }
  (new Set(headers.map((header) => header.toLowerCase())).size !==
    headers.length &&
    failWithMessage(
      "A source header conflicts with a generated provenance column name.",
    ),
    (rows.length > consolidationBudgets.outputRows ||
      (rows.length + 1) * headers.length > consolidationBudgets.outputCells) &&
      failWithMessage(
        "The output exceeds 1,000 data rows or 20,000 cells including headers and provenance.",
      ),
    rows.reduce(
      (z, V) =>
        z +
        V.values.reduce((D, B) => D + (typeof B == "string" ? B.length : 0), 0),
      headers.join("").length,
    ) > consolidationBudgets.totalTextCharacters &&
      failWithMessage("Expanded output text exceeds 1,000,000 characters."),
    workbook.sheets.reduce((z, V) => z + Object.keys(V.cells).length, 0) +
      (rows.length + 1) * headers.length >
      workbookLimits.maxCells &&
      failWithMessage(
        "The new output would exceed the workbook's stored-cell budget.",
      ));
  const worksheetRecord = createWorksheetRecord(o.outputSheetName),
    snapshot = JSON.stringify(workbook),
    I = new Set(snapshot.match(/"population-combine-[1-9]\d*"/g) ?? []);
  let O = 1;
  for (; I.has(JSON.stringify(`population-combine-${O}`));) O++;
  ((worksheetRecord.id = `population-combine-${O}`),
    (worksheetRecord.frozenRows = 1),
    (worksheetRecord.columnWidths = Object.fromEntries(
      headers.map((header, V) => [
        V,
        Math.min(240, Math.max(110, header.length * 7 + 20)),
      ]),
    )),
    headers.forEach((raw, V) => {
      worksheetRecord.cells[formatCellAddress(0, V)] = {
        raw: raw,
        literal: raw.startsWith("="),
        style: { bold: true, fill: "#E2F0D9" },
      };
    }),
    (worksheetRecord.cells.A1.note = `Reviewed ${o.mode} values snapshot; source workbook ${workbook.id}, revision ${workbook.revision}. Left: '${left.sheetName}'!${left.source.range} (${left.records.length} records); right: '${right.sheetName}'!${right.source.range} (${right.records.length} records). Keys: ${o.keys ? JSON.stringify(o.keys) : "exact matching headers"}. ${matching?.policy ?? "Append by exact headers."} Normalization: ${JSON.stringify(o.normalization ?? {})}; blank keys do not match. Matched left/right: ${matchedLeftRows}/${matchedRightRows}; match pairs: ${matchPairs}; unmatched left/right: ${unmatchedLeft.length}/${unmatchedRight.length}; output records: ${rows.length}. Duplicate key groups left/right: ${duplicateLeftKeys}/${duplicateRightKeys}; many-to-many groups: ${manyToManyKeys}; extra many-to-many pairs: ${expandedMatchRows}. Hidden sources explicitly included: ${context.includeHidden}. Filter views do not exclude records. This snapshot does not automatically refresh.`),
    rows.forEach((row, V) =>
      [...row.values, row.left, row.right].forEach((raw, B) => {
        worksheetRecord.cells[formatCellAddress(V + 1, B)] = {
          raw: raw,
          ...(typeof raw == "string" && raw.startsWith("=")
            ? { literal: true }
            : {}),
          note: `Combined values snapshot. Left source: ${row.left ?? "unmatched / not applicable"}. Right source: ${row.right ?? "unmatched / not applicable"}. Workbook ${workbook.id}; revision ${workbook.revision}. ${row.match?.explanation ?? "No matched pair / append source record."}`,
        };
      }),
    ));
  const next = structuredClone(workbook);
  return (
    next.sheets.push(worksheetRecord),
    normalizeWorkbookInput(next),
    {
      snapshot: snapshot,
      request: o,
      context: context,
      left: left,
      right: right,
      headers: headers,
      rows: rows,
      unmatchedLeft: unmatchedLeft,
      unmatchedRight: unmatchedRight,
      keyProfiles: keyProfiles,
      ...(matching ? { matching: matching } : {}),
      counts: {
        leftRows: left.records.length,
        rightRows: right.records.length,
        matchedLeftRows: matchedLeftRows,
        matchedRightRows: matchedRightRows,
        matchPairs: matchPairs,
        duplicateLeftKeys: duplicateLeftKeys,
        duplicateRightKeys: duplicateRightKeys,
        manyToManyKeys: manyToManyKeys,
        expandedMatchRows: expandedMatchRows,
        outputRows: rows.length,
      },
      outputSheetId: worksheetRecord.id,
      outputRange: `A1:${formatCellAddress(rows.length, headers.length - 1)}`,
      next: next,
    }
  );
}
function assertConsolidationPreviewCurrent(workbook, r, t = {}) {
  (t.readOnly && failWithMessage("This workbook is read-only."),
    r.snapshot !== JSON.stringify(workbook) &&
      failWithMessage(
        "The workbook changed after preview. Preview and review the populations again.",
      ));
  const consolidationOutput = buildConsolidationOutput(workbook, r.request, t);
  return (
    JSON.stringify(consolidationOutput) !== JSON.stringify(r) &&
      failWithMessage(
        "The preview or host context changed. Preview and review again.",
      ),
    consolidationOutput.next
  );
}
function normalizeAuthorizedContext(workbook, options) {
  if (
    (options.readOnly !== undefined && typeof options.readOnly != "boolean") ||
    (options.includeHidden !== undefined &&
      typeof options.includeHidden != "boolean")
  )
    throw Error("Host read-only and hidden-source policies must be boolean.");
  const allowed =
    options.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id);
  if (
    !Array.isArray(allowed) ||
    allowed.some(
      (o) =>
        typeof o != "string" ||
        !workbook.sheets.some((sheet) => sheet.id === o),
    )
  )
    throw Error("Invalid authorized worksheet context.");
  const external = options.approvedWorkbooks ?? [];
  if (
    !Array.isArray(external) ||
    external.length > 8 ||
    new Set(external.map((o) => o.workbook.id)).size !== external.length ||
    external.some(
      (o) =>
        o.workbook.id === workbook.id ||
        !Array.isArray(o.allowedSheetIds) ||
        o.allowedSheetIds.some(
          (allowedSheetId) =>
            !o.workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
        ) ||
        (o.includeHidden !== undefined && typeof o.includeHidden != "boolean"),
    )
  )
    throw Error(
      "Approved workbook identities and worksheet permissions must be explicit, distinct and valid (at most 8 workbooks).",
    );
  return { allowed: allowed, external: external };
}
function readSourceTableRecords(workbook, source, options) {
  const authorizedContext = normalizeAuthorizedContext(workbook, options),
    o = authorizedContext.external.find(
      (p) => p.workbook.id === source.workbookId,
    ),
    a = source.workbookId === workbook.id,
    book = a ? workbook : o?.workbook;
  if (!book)
    throw Error(
      "Source workbook is missing or has not been approved. Attach and approve the exact workbook identity before refreshing.",
    );
  const allowedSheetIds = a ? authorizedContext.allowed : o.allowedSheetIds,
    u = book.sheets.find((sheet) => sheet.id === source.sheetId);
  if (!u || !allowedSheetIds.includes(source.sheetId))
    throw Error(
      "Source worksheet is missing or outside the approved worksheet scope.",
    );
  let range = source.range;
  if (source.tableId) {
    const p = u.tables?.find((table) => table.id === source.tableId);
    if (!p) throw Error("Source table is missing. Review the source mapping.");
    const rangeAddress = parseRangeAddress(p.range);
    range = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1 - (p.showTotals ? 1 : 0)}`;
  }
  const sourceHeaders = readSourceHeaders(
      book,
      [{ sheetId: u.id, range: range }],
      {
        allowedSheetIds: allowedSheetIds,
        includeHidden: a
          ? options.includeHidden
          : !!options.includeHidden && !!o.includeHidden,
      },
    ),
    snapshot = sourceHeaders.snapshots[0];
  return {
    book: book,
    sheet: u,
    snapshot: snapshot,
    fingerprint: fingerprintCanonicalValue({
      workbookId: book.id,
      workbookName: book.name,
      source: source,
      snapshot: snapshot,
      dependencies: sourceHeaders.dependencyCells,
      datePolicy: normalizeDatePolicy(book.datePolicy),
    }),
  };
}
function reconcileConsolidationSources(workbook, r, options = {}) {
  const consolidationDefinition = parseConsolidationDefinition(r);
  normalizeAuthorizedContext(workbook, options);
  const o = {
      sources: [],
      contributions: [],
      subtotals: [],
      exceptions: [],
      grandTotal: null,
      sourceTotal: null,
      difference: null,
      complete: false,
    },
    a = new Set(),
    s = (sourceId, code, l, row, severity = "error") =>
      o.exceptions.push({
        sourceId: sourceId,
        code: code,
        message: l,
        ...(row ? { row: row } : {}),
        severity: severity,
      });
  for (const source of consolidationDefinition.sources)
    try {
      const sourceTableRecords = readSourceTableRecords(
          workbook,
          source,
          options,
        ),
        { snapshot } = sourceTableRecords,
        rangeAddress = parseRangeAddress(snapshot.source.range),
        h = (y) => {
          const w = snapshot.headers.indexOf(y);
          if (w < 0)
            throw Error(
              `Required field “${y}” is missing. Review the source schema.`,
            );
          return w;
        },
        f = source.categories.map(h),
        p = h(source.amountField),
        m = (y, w) =>
          source[y].kind === "constant"
            ? source[y].value
            : w[h(source[y].field)];
      for (const y of ["period", "currency", "unit"])
        source[y].kind === "field" && h(source[y].field);
      const v = {
        sourceId: source.id,
        workbookId: sourceTableRecords.book.id,
        workbookName: sourceTableRecords.book.name,
        sheetId: sourceTableRecords.sheet.id,
        sheetName: sourceTableRecords.sheet.name,
        range: snapshot.source.range,
        fingerprint: sourceTableRecords.fingerprint,
        records: snapshot.records.length,
        acceptedRecords: 0,
        amount: 0,
        periods: [],
        currencies: [],
        units: [],
      };
      o.sources.push(v);
      const g = new Map(),
        b = [];
      for (const record of snapshot.records) {
        let w = false;
        const E = (V, D) => {
            (s(source.id, V, D, record.row), (w = true));
          },
          categories = f.map((V) => record.values[V]),
          amount = record.values[p],
          address = `${columnIndexToLetters(rangeAddress.firstCol + p)}${record.row}`,
          I = JSON.stringify([
            sourceTableRecords.book.id,
            sourceTableRecords.sheet.id,
            address,
          ]);
        (a.has(I) &&
          E(
            "overlap",
            `Amount ${sourceTableRecords.sheet.name}!${address} is included more than once. Remove overlapping sources to prevent double counting.`,
          ),
          a.add(I),
          categories.some(
            (category) =>
              category === null ||
              (typeof category == "string" && !category.trim()),
          ) &&
            E(
              "blank-category",
              "Every contributing row needs nonblank categories.",
            ),
          (typeof amount != "number" || !Number.isFinite(amount)) &&
            E(
              "amount",
              `Amount ${sourceTableRecords.sheet.name}!${address} must be a finite number; text is not silently converted.`,
            ));
        const O = {
          period: m("period", record.values),
          currency: m("currency", record.values),
          unit: m("unit", record.values),
        };
        for (const V of ["period", "currency", "unit"]) {
          const D = O[V];
          if (
            typeof D == "string" &&
            D.trim() &&
            D.length <= 128 &&
            !/[\u0000-\u001f\u007f]/.test(D)
          ) {
            const B =
              v[
                V === "currency"
                  ? "currencies"
                  : V === "period"
                    ? "periods"
                    : "units"
              ];
            B.includes(D) || B.push(D);
          }
          D !== consolidationDefinition[V] &&
            E(
              V,
              `Source ${V} ${JSON.stringify(D)} does not match ${JSON.stringify(consolidationDefinition[V])}. No currency, unit or period conversion is implicit.`,
            );
        }
        const j = serializeTypedKeyValues(categories),
          z = g.get(j);
        (z !== undefined
          ? (s(
              source.id,
              "duplicate-category",
              `Repeated category also occurs at row ${z}. ${source.duplicatePolicy === "sum" ? "Explicit sum policy retains both contributions." : "Choose Sum repeated categories only after reviewing these records."}`,
              record.row,
              source.duplicatePolicy === "sum" ? "warning" : "error",
            ),
            source.duplicatePolicy === "reject" && (w = true))
          : g.set(j, record.row),
          w ||
            b.push({
              sourceId: source.id,
              workbookId: sourceTableRecords.book.id,
              workbookName: sourceTableRecords.book.name,
              sheetId: sourceTableRecords.sheet.id,
              sheetName: sourceTableRecords.sheet.name,
              row: record.row,
              address: address,
              locator: `${sourceTableRecords.book.name} / ${snapshot.sheetName}!${address}`,
              categories: categories,
              amount: amount,
              period: O.period,
              currency: O.currency,
              unit: O.unit,
              link:
                sourceTableRecords.book.id !== workbook.id
                  ? "approved snapshot"
                  : consolidationDefinition.linkedAmounts
                    ? "live"
                    : "snapshot",
            }));
      }
      ((v.acceptedRecords = b.length),
        (v.amount = sumConsolidationAmounts(b.map((y) => y.amount))),
        o.contributions.push(...b));
    } catch (error) {
      s(source.id, "source", error.message.slice(0, 4e3));
    }
  if (o.contributions.length > consolidationLimits.contributions)
    throw Error(
      "Consolidation exceeds 2,000 contributions. Narrow the sources; no output was truncated.",
    );
  if (
    (o.contributions.length ||
      s(
        consolidationDefinition.sources[0].id,
        "empty",
        "No valid contributing amounts. Review the sources and exceptions.",
      ),
    !o.exceptions.some((exception) => exception.severity === "error"))
  ) {
    ((o.subtotals = buildConsolidationSubtotals(o.contributions)),
      (o.grandTotal = sumConsolidationAmounts(
        o.contributions.map((contribution) => contribution.amount),
      )),
      (o.sourceTotal = sumConsolidationAmounts(
        o.sources.map((source) => source.amount),
      )),
      (o.difference = o.grandTotal - o.sourceTotal));
    const d =
      Number.EPSILON *
      Math.max(1, Math.abs(o.grandTotal), Math.abs(o.sourceTotal)) *
      o.contributions.length *
      4;
    (Math.abs(o.difference) > d &&
      s(
        consolidationDefinition.sources[0].id,
        "reconciliation",
        "Source control totals do not reconcile to the contributing amounts.",
      ),
      (o.complete = !o.exceptions.some(
        (exception) => exception.severity === "error",
      )));
  }
  return o;
}
function buildAuthorizedContextSnapshot(workbook, options) {
  const authorizedContext = normalizeAuthorizedContext(workbook, options);
  return fingerprintCanonicalValue({
    allowed: [...authorizedContext.allowed].sort(),
    hidden: !!options.includeHidden,
    readOnly: !!options.readOnly,
    approved: authorizedContext.external
      .map((n) => ({
        id: n.workbook.id,
        sheets: [...n.allowedSheetIds].sort(),
        hidden: !!n.includeHidden,
      }))
      .sort((n, o) => n.id.localeCompare(o.id)),
  });
}
function validateOutputSheetName(workbook, outputName) {
  if (
    typeof outputName != "string" ||
    !outputName.trim() ||
    outputName !== outputName.trim() ||
    outputName.length > 31 ||
    /[\\/?*\[\]:\u0000-\u001f\u007f]/.test(outputName) ||
    outputName.startsWith("'") ||
    outputName.endsWith("'") ||
    workbook.sheets.some(
      (sheet) => sheet.name.toLowerCase() === outputName.toLowerCase(),
    )
  )
    throw Error(
      "Choose a distinct worksheet name of 1–31 characters without invalid characters.",
    );
  inspectOutputNameSafety(workbook, outputName);
}
function createConsolidationPreview(workbook, request, options = {}) {
  validateOutputSheetName(workbook, request.outputName);
  const consolidationDefinition = parseConsolidationDefinition(
      request.definition,
    ),
    review = reconcileConsolidationSources(
      workbook,
      consolidationDefinition,
      options,
    );
  return {
    request: {
      definition: consolidationDefinition,
      outputName: request.outputName,
    },
    workbookHash: fingerprintCanonicalValue(workbook),
    contextHash: buildAuthorizedContextSnapshot(workbook, options),
    review: review,
  };
}
function buildConsolidationOutputCells(worksheetRecord, outputRange) {
  const rangeAddress = parseRangeAddress(outputRange);
  if (
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    consolidationLimits.outputCells
  )
    throw Error("Consolidation output range exceeds its cell budget.");
  const n = [];
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  ) {
    const a = [];
    for (
      let firstCol = rangeAddress.firstCol;
      firstCol <= rangeAddress.lastCol;
      firstCol++
    ) {
      const cell =
        worksheetRecord.cells[
          `${columnIndexToLetters(firstCol)}${firstRow + 1}`
        ];
      a.push({
        raw: cell?.raw ?? null,
        literal: !!cell?.literal,
        note: cell?.note ?? "",
      });
    }
    n.push(a);
  }
  return fingerprintCanonicalValue(n);
}

export {
  readSourceRectangles,
  buildConsolidationOutput,
  readSourceHeaders,
  matchKeyPopulations,
  assertConsolidationPreviewCurrent,
  readSourceTableRecords,
  inspectFormulaDependencies,
  buildConsolidationOutputCells,
  inspectOutputNameSafety,
  createConsolidationPreview,
  buildAuthorizedContextSnapshot,
};
export { hasOnlyAllowedKeys };
