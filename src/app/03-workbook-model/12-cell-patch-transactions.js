import {
  isNonArrayObject,
  throwWithMessage,
} from "./01-cell-styles-and-themes.js";
import { isDeepFrozenValue } from "../01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "../02-text-and-financial-functions.js";
import {
  throwWorkbookError,
  assertNamedStyleBindingCurrent,
  touchWorkbookModifiedAt,
  assertSheetUnprotected,
  validateCellStyle,
  failWithMessage,
} from "./01-cell-styles-and-themes.js";
import {
  buildNumberFormatCode,
  validateNumberFormatSettings,
} from "./02-formula-function-evaluators.js";
import { carryCalculatorRecordForward } from "./05-formula-engine-and-name-operations.js";
import {
  requireBoundedText,
  NC,
  QU,
  columnIndexToLetters,
} from "./06-data-model-and-analysis.js";
import {
  enumerateSelectionCells,
  _t,
  parseRangeAddress,
  buildBorderPresetStyles,
} from "./07-conditional-rules-worksheet-features.js";
import {
  calculateWorkbookWithSettings,
  checkDataValidationBatch,
} from "./09-calculation-engine.js";
import {
  findSheetOrThrow,
  renameTableHeaderField,
} from "./10-table-filters-and-pivot-views.js";
import {
  normalizeWorkbookModel,
  normalizeCellRecordInput,
  t$,
} from "./11-table-model.js";

function applyCellPatchTransaction(
  e,
  r,
  t = "Preparer",
  n = "Edit cells",
  o = {},
) {
  if (!Array.isArray(r) || r.length > workbookLimits.maxPatches)
    return throwWithMessage(
      `A transaction may contain at most ${workbookLimits.maxPatches} cell changes.`,
    );
  if (r.length === 0) return e;
  (requireBoundedText(t, "Actor", 100), requireBoundedText(n, "Action", 200));
  const a = new Set();
  let s;
  const d = r.map((p) => {
      if (!isNonArrayObject(p)) return throwWithMessage("Invalid cell patch.");
      const m = findSheetOrThrow(e, p.sheetId);
      assertSheetUnprotected(m);
      const { row, col } = _t(p.address);
      if (row >= m.rowCount || col >= m.colCount)
        return throwWithMessage(
          `Cell ${p.address} is outside sheet dimensions.`,
        );
      const b = `${columnIndexToLetters(col)}${row + 1}`;
      assertNamedStyleBindingCurrent(e, normalizeCellRecordInput(p.cell).style);
      const y = m.cells[b];
      if (
        (y?.raw ?? null) !== p.cell?.raw ||
        !!y?.literal != !!p.cell?.literal
      ) {
        s ??= calculateWorkbookWithSettings(e);
        const k = s.getSpillAnchor(m.id, b);
        if (k && k !== b) {
          const F = r.find((I) => I.sheetId === m.id && I.address === k);
          if (
            !F ||
            (!F.cell.literal &&
              typeof F.cell.raw == "string" &&
              F.cell.raw.startsWith("="))
          )
            return throwWithMessage(
              `Cell ${b} belongs to the array formula in ${k}. Edit the array formula instead.`,
            );
        }
      }
      const w = `${m.id}:${b}`;
      return a.has(w)
        ? throwWithMessage("A transaction cannot change the same cell twice.")
        : (a.add(w),
          {
            sheetId: m.id,
            address: b,
            before: m.cells[b]
              ? normalizeCellRecordInput(m.cells[b])
              : undefined,
            after: normalizeCellRecordInput(p.cell),
          });
    }),
    u = d.filter((p) => !p.before).length;
  if (
    e.sheets.reduce(
      (p, m) =>
        p +
        (isDeepFrozenValue(m.cells)
          ? (QU.get(m.cells) ?? Object.keys(m.cells).length)
          : Object.keys(m.cells).length),
      0,
    ) +
      u >
    workbookLimits.maxCells
  )
    return throwWithMessage(
      "Workbook exceeds the 100,000 populated cell limit.",
    );
  const l = e.sheets.map((p) => {
      const m = d.filter((g) => g.sheetId === p.id);
      if (!m.length) return p;
      const v = { ...p.cells };
      return (
        m.forEach((g) => {
          v[g.address] = g.after;
        }),
        { ...p, cells: v }
      );
    }),
    c = renameTableHeaderField(e, { ...e, sheets: l });
  if (c.sheets !== l)
    for (const p of c.sheets) {
      const m = findSheetOrThrow(e, p.id);
      for (const [v, g] of Object.entries(p.cells)) {
        const b = d.find((y) => y.sheetId === p.id && y.address === v);
        if (b) b.after = normalizeCellRecordInput(g);
        else if (
          JSON.stringify(m.cells[v]) !== JSON.stringify(g) &&
          (d.push({
            sheetId: p.id,
            address: v,
            before: m.cells[v]
              ? normalizeCellRecordInput(m.cells[v])
              : undefined,
            after: normalizeCellRecordInput(g),
          }),
          d.length > workbookLimits.maxPatches)
        )
          return throwWithMessage(
            "Renaming this table field affects more than 500 cells. Reduce the dependent formulas before renaming it.",
          );
      }
    }
  const h = checkDataValidationBatch(
    c,
    d.filter(
      (p) =>
        p.before?.raw !== p.after.raw || p.before?.literal !== p.after.literal,
    ),
    o.validationApproval,
  );
  if (
    (h && (n = `${n.slice(0, 140)} (accepted ${h} validation warnings)`),
    e.events.length >= 2e3)
  )
    return throwWithMessage(
      "Change history has reached its 2,000 transaction limit. Export the workbook before starting a new workpaper.",
    );
  if (
    e.events.reduce((p, m) => p + m.changes.length, 0) + d.length >
    workbookLimits.maxCells
  )
    return throwWithMessage(
      "Change history has reached its 100,000 cell-change limit. Export the workbook before starting a new workpaper.",
    );
  if (e.revision >= Number.MAX_SAFE_INTEGER - 1)
    return throwWithMessage("Workbook revision limit reached.");
  const f = {
    ...touchWorkbookModifiedAt(c),
    revision: e.revision + 1,
    sheets: c.sheets,
    events: [
      ...e.events,
      {
        id: crypto.randomUUID(),
        timestamp: new Date().toISOString(),
        actor: t,
        action: n,
        changes: d,
      },
    ],
  };
  return (
    isDeepFrozenValue(e) &&
      NC.has(e) &&
      c.sheets === l &&
      d.every((p) => !p.after.hyperlink && !p.after.style?.binding) &&
      t$(f),
    carryCalculatorRecordForward(e, f),
    f
  );
}
function applyBorderEditTransaction(e, r, t, n, o = {}) {
  o.readOnly && failWithMessage("This workbook is read-only.");
  const a = findSheetOrThrow(e, r);
  assertSheetUnprotected(a);
  const s = validateCellStyle(n),
    d = enumerateSelectionCells(typeof t == "string" ? [t] : t),
    u = new Set(d);
  if (o.borders || Object.hasOwn(n, "borders") || Object.hasOwn(n, "border")) {
    const h = d.map(_t);
    (a.merges ?? []).some((f) => {
      const p = parseRangeAddress(f);
      return h.some(
        (m) =>
          m.row >= p.firstRow &&
          m.row <= p.lastRow &&
          m.col >= p.firstCol &&
          m.col <= p.lastCol,
      );
    }) &&
      failWithMessage(
        "Unmerge the intersecting cells before applying border edges or presets.",
      );
  }
  const l = d.flatMap((h) => {
    const f = a.cells[h],
      p = { ...f?.style, ...n, ...s };
    (Object.hasOwn(n, "border") && delete p.borders,
      o.borders &&
        (p.borders = {
          ...f?.style?.borders,
          ...buildBorderPresetStyles(h, u, o.borders.preset, o.borders.line),
        }),
      n.borders !== undefined &&
        (p.borders = {
          ...(o.borders ? p.borders : f?.style?.borders),
          ...s.borders,
        }));
    const m = validateCellStyle(p),
      v = {
        ...(f ?? { raw: null }),
        ...(Object.keys(m).length ? { style: m } : {}),
      };
    return (
      Object.keys(m).length || delete v.style,
      !f && !Object.keys(m).length
        ? []
        : JSON.stringify(validateCellStyle(f?.style ?? {})) ===
            JSON.stringify(m)
          ? []
          : JSON.stringify(f) === JSON.stringify(v)
            ? []
            : [{ sheetId: r, address: h, cell: v }]
    );
  });
  if (!l.length) return e;
  const c = applyCellPatchTransaction(e, l, "Preparer", "Format cell styles");
  return { ...e, sheets: c.sheets };
}
function applyNumberFormatTransaction(e, r, t, n, o = {}) {
  o.readOnly && throwWorkbookError("This workbook is read-only.");
  const a = findSheetOrThrow(e, r);
  assertSheetUnprotected(a);
  const s = validateNumberFormatSettings(n),
    d = enumerateSelectionCells(typeof t == "string" ? [t] : t),
    u = buildNumberFormatCode(s),
    l = d.flatMap((h) => {
      const f = a.cells[h],
        p = {
          ...(f ?? { raw: null }),
          numberFormat: s,
          numFmt: u,
          format:
            s.category === "percent"
              ? "percent"
              : s.category === "currency" || s.category === "accounting"
                ? "currency"
                : s.category === "date" || s.category === "datetime"
                  ? "date"
                  : s.category === "number"
                    ? "number"
                    : "general",
        };
      return JSON.stringify(f) === JSON.stringify(p)
        ? []
        : [{ sheetId: r, address: h, cell: p }];
    });
  if (!l.length) return e;
  const c = applyCellPatchTransaction(e, l, "Preparer", "Format numbers");
  return { ...e, sheets: c.sheets };
}
function buildCanonicalWorkbookSnapshot(e) {
  const r = normalizeWorkbookModel(e, true);
  return NC.has(r) && isDeepFrozenValue(r) ? r : t$(r);
}

export {
  applyBorderEditTransaction,
  applyNumberFormatTransaction,
  applyCellPatchTransaction,
  buildCanonicalWorkbookSnapshot,
};
