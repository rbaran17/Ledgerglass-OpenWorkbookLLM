import { isPlainRecord } from "./19b-assistant-protocol.js";

import { downloadBlobAsFile } from "./09-workbook-links.js";
import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  deepFreezeValue,
  measureSerializedBytes,
  defineDataPropertyByKey,
} from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  builtinFunctionNames,
  requireAlternateFormatOriginal,
  parseRangeAddress,
  touchWorkbookModifiedAt,
  buildCanonicalWorkbookSnapshot,
  normalizeWorkbookInput,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  importWorkbookFile,
  describeImportSource,
  templateConversionRoutes,
  detectSupportedFileFormats,
} from "./07-file-import-and-worker-dispatch.js";
import {
  exportXltxTemplate,
  exportWorkbookToXlsx,
} from "./08-xlsx-import-export-and-preservation.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import {
  validateConvertedSource,
  describeSourceFile,
  XlsxConversionReview,
} from "./10-source-import-dialog.js";

function FormatConversionDialog(e) {
  const { text } = useWorkbookMessages(),
    t = ne.useId(),
    n = ne.useRef(null),
    cancellableOperation = useCancellableOperation(),
    a = ne.useRef(e),
    s = ne.useRef(0),
    d = ne.useRef(true);
  a.current = e;
  const [u, l] = ne.useState(null),
    [c, h] = ne.useState(null),
    [checked, p] = ne.useState(false),
    [disabled, v] = ne.useState(false),
    [g, b] = ne.useState(""),
    [y, w] = ne.useState(null),
    [E, A] = ne.useState("");
  (ne.useEffect(() => {
    d.current = true;
    const activeElement = document.activeElement;
    return (
      n.current?.focus(),
      () => {
        ((d.current = false),
          s.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      y || n.current?.focus();
    }, [y]));
  const k = async (V) => {
      (b(""), A(""), v(true));
      try {
        await V();
      } catch (error) {
        d.current && b(error.message);
      } finally {
        d.current && v(false);
      }
    },
    F = (V) => {
      const D = ++s.current;
      return (
        l(null),
        h(null),
        p(false),
        w(null),
        k(async () => {
          const B = await detectSupportedFileFormats(
            V,
            cancellableOperation.begin(),
          );
          d.current && D === s.current && (l(B), h(V));
        })
      );
    };
  ne.useEffect(() => {
    e.initialFile && F(e.initialFile);
  }, [e.initialFile]);
  const I = (V, D) =>
      k(async () => {
        if (a.current.readOnly) throw Error("This workbook is read-only.");
        const workbook = a.current.workbook,
          _ = ++s.current,
          imported = D
            ? await importWorkbookFile(V, cancellableOperation.begin())
            : await validateConvertedSource(
                u,
                V,
                checked,
                cancellableOperation.begin(),
              );
        if (!(!d.current || _ !== s.current)) {
          if (a.current.workbook !== workbook || a.current.readOnly)
            throw Error(
              "The workbook or permissions changed. Review the conversion again.",
            );
          if (!imported.report)
            throw Error(
              "A compatibility report is required for format conversion.",
            );
          w({
            snapshot: workbook,
            imported: imported,
            report: imported.report,
          });
        }
      }),
    O = (V) =>
      k(async () => {
        const workbook = a.current.workbook,
          B = ++s.current,
          _ = V
            ? await exportXltxTemplate(workbook)
            : await exportWorkbookToXlsx(workbook);
        if (!(!d.current || B !== s.current)) {
          if (a.current.workbook !== workbook)
            throw Error("The workbook changed. Review the export again.");
          w({
            snapshot: workbook,
            blob: _.blob,
            report: {
              ..._.report,
              fileName: describeImportSource(workbook, V ? "xltx" : "xlsx"),
            },
          });
        }
      }),
    j = (filename, blob) =>
      (a.current.onDownload ?? downloadBlobAsFile)({
        filename: filename,
        blob: blob,
      }),
    z = (V) =>
      k(async () => {
        const D = V
          ? requireAlternateFormatOriginal(a.current.workbook)
          : describeSourceFile(u);
        await j(
          D.fileName,
          new Blob([new Uint8Array(D.bytes)], {
            type: "application/octet-stream",
          }),
        );
      });
  return y
    ? i.jsx(XlsxConversionReview, {
        report: y.report,
        originalLabel: u
          ? `Download original ${u.format.toUpperCase()}`
          : undefined,
        blocked:
          g ||
          (y.snapshot !== e.workbook
            ? "The workbook changed. Cancel and review again."
            : y.imported && e.readOnly
              ? "This workbook is read-only."
              : disabled
                ? "Completing the reviewed operation."
                : undefined),
        onClose: () => {
          (w(null), b(""));
        },
        onOriginal: u && y.imported ? () => void z(false) : undefined,
        onReport: () =>
          void k(async () => {
            await j(
              y.report.fileName + ".compatibility.json",
              new Blob([JSON.stringify(y.report, null, 2)], {
                type: "application/json",
              }),
            );
          }),
        onConfirm: () =>
          void k(async () => {
            if (a.current.workbook !== y.snapshot)
              throw Error("The workbook changed. Review the conversion again.");
            if (y.imported) {
              if (a.current.readOnly)
                throw Error("This workbook is read-only.");
              (a.current.onWorkbookChange(
                y.imported.workbook,
                "Open reviewed workbook conversion",
              ),
                A(
                  "Converted workbook opened. Keep a native JSON backup to retain both source files and conversion provenance.",
                ));
            } else
              y.blob &&
                (await j(y.report.fileName, y.blob),
                A(
                  "Derived workbook downloaded. Verify fidelity in the source application.",
                ));
            w(null);
          }),
      })
    : i.jsx("div", {
        className: "lg-reports-overlay",
        children: i.jsxs("section", {
          role: "dialog",
          "aria-modal": "true",
          "aria-labelledby": t,
          lang: "en",
          tabIndex: -1,
          ref: n,
          className: "lg-reports-dialog lg-formats-dialog",
          onKeyDown: (event) => {
            var D, B;
            if (
              (event.key === "Escape" && (event.stopPropagation(), e.onClose()),
              event.key === "Tab")
            ) {
              const _ = [
                ...n.current.querySelectorAll(
                  'button:not(:disabled),input:not(:disabled),summary,[tabindex="0"]',
                ),
              ];
              event.shiftKey &&
              (document.activeElement === _[0] ||
                document.activeElement === n.current)
                ? (event.preventDefault(), (D = _.at(-1)) == null || D.focus())
                : !event.shiftKey &&
                  document.activeElement === _.at(-1) &&
                  (event.preventDefault(), (B = _[0]) == null || B.focus());
            }
          },
          children: [
            i.jsxs("header", {
              children: [
                i.jsxs("div", {
                  children: [
                    i.jsx("p", { children: "WORKBOOK EXCHANGE" }),
                    i.jsx("h2", {
                      id: t,
                      children: i.jsx(LocalizedText, {
                        children: "Workbook formats",
                      }),
                    }),
                  ],
                }),
                i.jsx("button", {
                  "aria-label": text("Close workbook formats"),
                  onClick: e.onClose,
                  children: i.jsx(X, { size: 20 }),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "lg-reports-scroll",
              children: [
                i.jsx("p", {
                  className: "lg-reports-intro",
                  children:
                    "Open a template or bring an older workbook into this editor. Review conversion losses and keep the untouched original with your work.",
                }),
                g &&
                  i.jsx("p", {
                    role: "alert",
                    className: "lg-formats-error",
                    children: g,
                  }),
                E && i.jsx("p", { role: "status", children: E }),
                disabled &&
                  i.jsx("p", {
                    role: "status",
                    children: "Processing workbook…",
                  }),
                i.jsxs("div", {
                  className: "lg-formats-actions",
                  children: [
                    i.jsxs("label", {
                      children: [
                        "Inspect source file",
                        i.jsx("input", {
                          type: "file",
                          accept: ".xls,.xlsb,.xlsm,.xltx,.xltm,.ods",
                          "aria-label": text("Inspect alternate workbook"),
                          disabled: disabled,
                          onChange: (event) => {
                            const D = event.target.files?.[0];
                            ((event.target.value = ""), D && F(D));
                          },
                        }),
                      ],
                    }),
                    i.jsx("button", {
                      disabled: disabled,
                      onClick: () => void O(false),
                      children: "Review XLSX export",
                    }),
                    i.jsx("button", {
                      disabled: disabled,
                      onClick: () => void O(true),
                      children: "Review XLTX export",
                    }),
                  ],
                }),
                e.workbook.alternateSource &&
                  i.jsxs("section", {
                    className: "lg-formats-source",
                    children: [
                      i.jsx("h3", { children: "Retained original" }),
                      i.jsxs("p", {
                        children: [
                          e.workbook.alternateSource.fileName,
                          " ·",
                          " ",
                          e.workbook.alternateSource.conversion.method,
                          " · equivalence not verified",
                        ],
                      }),
                      i.jsxs("p", {
                        className: "lg-formats-hash",
                        children: [
                          "SHA-256: ",
                          e.workbook.alternateSource.sha256,
                        ],
                      }),
                      i.jsx("button", {
                        disabled: disabled,
                        onClick: () => void z(true),
                        children: "Download retained original",
                      }),
                    ],
                  }),
                u &&
                  i.jsxs("section", {
                    className: "lg-formats-source",
                    children: [
                      i.jsx("h3", { children: u.fileName }),
                      i.jsxs("p", {
                        children: [
                          u.format.toUpperCase(),
                          " identified by content ·",
                          " ",
                          (u.fileBytes / 1024).toFixed(1),
                          " KiB ·",
                          " ",
                          u.parts.length,
                          " inventory entries",
                        ],
                      }),
                      i.jsx("p", { children: u.identity }),
                      i.jsxs("p", {
                        className: "lg-formats-hash",
                        children: ["SHA-256: ", u.sha256],
                      }),
                      i.jsxs("div", {
                        className: "lg-formats-actions",
                        children: [
                          i.jsx("button", {
                            disabled: disabled,
                            onClick: () => void z(false),
                            children: "Download inspected original",
                          }),
                          i.jsx("button", {
                            disabled: disabled,
                            onClick: () =>
                              void k(async () => {
                                await j(
                                  u.fileName + ".format-review.json",
                                  new Blob([JSON.stringify(u, null, 2)], {
                                    type: "application/json",
                                  }),
                                );
                              }),
                            children: "Download format report",
                          }),
                          e.onAutomation &&
                            ["xls", "xlsb", "xlsm", "xltx", "xltm"].includes(
                              u.format,
                            ) &&
                            i.jsx("button", {
                              disabled: disabled,
                              onClick: () =>
                                void k(async () => {
                                  await e.onAutomation(c);
                                }),
                              children: "Inspect automation",
                            }),
                        ],
                      }),
                      i.jsx("ul", {
                        children: u.warnings.map((warning) =>
                          i.jsx("li", { children: warning }, warning),
                        ),
                      }),
                      u.format === "xltx" &&
                        i.jsx("p", {
                          children: i.jsx("button", {
                            className: "lg-reports-primary",
                            disabled: disabled || e.readOnly,
                            onClick: () => void I(c, true),
                            children: "Preview template workbook",
                          }),
                        }),
                      i.jsxs("details", {
                        open: u.format !== "xltx",
                        children: [
                          i.jsx("summary", {
                            children: "Convert with the source application",
                          }),
                          i.jsx("ol", {
                            children: u.steps.map((step) =>
                              i.jsx("li", { children: step }, step),
                            ),
                          }),
                          i.jsxs("label", {
                            className: "lg-formats-ack",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: checked,
                                disabled: disabled || e.readOnly,
                                onChange: (event) => p(event.target.checked),
                              }),
                              "I acknowledge possible conversion losses and will compare formulas, dates, styles and special parts with the original.",
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              "Choose converted XLSX",
                              i.jsx("input", {
                                type: "file",
                                accept: ".xlsx",
                                "aria-label": text("Choose converted XLSX"),
                                disabled: disabled || e.readOnly || !checked,
                                onChange: (event) => {
                                  const D = event.target.files?.[0];
                                  ((event.target.value = ""), D && I(D, false));
                                },
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("details", {
                        children: [
                          i.jsxs("summary", {
                            children: [
                              "Source inventory (",
                              u.parts.length,
                              ")",
                            ],
                          }),
                          i.jsx("p", {
                            children:
                              "Original-only inventory; no content is executed. This is not a lossless conversion claim.",
                          }),
                          i.jsx("ul", {
                            className: "lg-formats-inventory",
                            children: u.parts.map((part, D) =>
                              i.jsxs(
                                "li",
                                {
                                  children: [
                                    part.path,
                                    " · ",
                                    part.bytes,
                                    " known bytes ·",
                                    " ",
                                    part.disposition,
                                  ],
                                },
                                D,
                              ),
                            ),
                          }),
                        ],
                      }),
                    ],
                  }),
                i.jsx("h3", {
                  children: "Read, write and preservation support",
                }),
                i.jsx("div", {
                  className: "lg-formats-table",
                  children: i.jsxs("table", {
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", { scope: "col", children: "Format" }),
                            i.jsx("th", { scope: "col", children: "Read" }),
                            i.jsx("th", { scope: "col", children: "Write" }),
                            i.jsx("th", { scope: "col", children: "Preserve" }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: templateConversionRoutes.map(
                          (templateConversionRoute) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("th", {
                                    scope: "row",
                                    children: [
                                      ".",
                                      templateConversionRoute.format,
                                    ],
                                  }),
                                  i.jsx("td", {
                                    children: templateConversionRoute.read,
                                  }),
                                  i.jsx("td", {
                                    children: templateConversionRoute.write,
                                  }),
                                  i.jsx("td", {
                                    children: templateConversionRoute.preserve,
                                  }),
                                ],
                              },
                              templateConversionRoute.format,
                            ),
                        ),
                      }),
                    ],
                  }),
                }),
                i.jsx("p", {
                  children:
                    "Files: up to 10 MiB. ZIP packages: up to 2,000 entries and 50 MiB expanded. Password-protected containers require an unencrypted copy. External conversions and native Office fidelity must be validated in the source application.",
                }),
              ],
            }),
            i.jsxs("footer", {
              children: [
                i.jsx("span", {
                  children: e.readOnly
                    ? "Read-only: inspection and export available; import disabled."
                    : "Imports replace the current workbook after review. Undo restores the previous workbook.",
                }),
                i.jsx("button", { onClick: e.onClose, children: "Close" }),
              ],
            }),
          ],
        }),
      });
}
class WorkbookRevisionConflictError extends Error {
  constructor(t, n, o) {
    super(
      o
        ? `${n} Current saved workbook ${o.workbookId}, revision ${o.revision}.`
        : n,
    );
    defineDataPropertyByKey(this, "currentRevision");
    defineDataPropertyByKey(this, "currentWorkbookId");
    ((this.code = t),
      (this.name = "PersistenceError"),
      (this.currentRevision = o?.revision),
      (this.currentWorkbookId = o?.workbookId));
  }
}
const localWorkbookNamespace = "ledgerglass-local-workbook",
  maxLocalCheckpoints = 20,
  cloneStructuredValue = (e) => structuredClone(e),
  measureUtf8Bytes = (e) => new TextEncoder().encode(e).byteLength;
function throwPersistenceError(e, r, guard) {
  throw new WorkbookRevisionConflictError(e, r, guard);
}
function stringifyCanonicalJson(e) {
  return e === undefined
    ? "undefined"
    : e === null || typeof e != "object"
      ? JSON.stringify(e)
      : Array.isArray(e)
        ? `[${e.map(stringifyCanonicalJson).join(",")}]`
        : `{${Object.keys(e)
            .sort()
            .filter((r) => e[r] !== undefined)
            .map((r) => `${JSON.stringify(r)}:${stringifyCanonicalJson(e[r])}`)
            .join(",")}}`;
}
function computeTextDigest(e) {
  let r = 2654435769,
    t = 2246822507,
    n = 3266489909,
    o = 668265263;
  for (let a = 0; a < e.length; a++) {
    const s = e.charCodeAt(a);
    ((r = Math.imul(r ^ s, 2246822507)),
      (t = Math.imul(t ^ s, 3266489909)),
      (n = Math.imul(n ^ s, 668265263)),
      (o = Math.imul(o ^ s, 374761393)));
  }
  return [e.length, r, t, n, o]
    .map((a) => (a >>> 0).toString(16).padStart(8, "0"))
    .join("");
}
function normalizeWorkbookSnapshot(workbook) {
  return normalizeWorkbookInput(workbook);
}
function normalizeSavedSelection(workbookSnapshot, view) {
  const t = workbookSnapshot.sheets.find((sheet) => !sheet.hidden);
  if (!t)
    return throwPersistenceError(
      "invalid",
      "A saved workbook needs a visible worksheet.",
    );
  if (view === undefined)
    return { activeSheetId: t.id, selection: { sheetId: t.id, range: "A1" } };
  if (
    !isPlainRecord(view) ||
    typeof view.activeSheetId != "string" ||
    !isPlainRecord(view.selection)
  )
    return throwPersistenceError(
      "invalid",
      "Invalid saved worksheet selection.",
    );
  const n = workbookSnapshot.sheets.find(
    (sheet) => sheet.id === view.activeSheetId && !sheet.hidden,
  );
  if (
    !n ||
    view.selection.sheetId !== n.id ||
    typeof view.selection.range != "string"
  )
    return throwPersistenceError(
      "invalid",
      "Saved selection must belong to the active visible worksheet.",
    );
  let o;
  try {
    o = parseRangeAddress(view.selection.range);
  } catch {
    return throwPersistenceError("invalid", "Invalid saved cell range.");
  }
  if (o.lastRow >= n.rowCount || o.lastCol >= n.colCount)
    return throwPersistenceError(
      "invalid",
      "Saved selection exceeds worksheet dimensions.",
    );
  const a = `${columnIndexToLetters(o.firstCol)}${o.firstRow + 1}`,
    s = `${columnIndexToLetters(o.lastCol)}${o.lastRow + 1}`;
  return {
    activeSheetId: n.id,
    selection: { sheetId: n.id, range: a === s ? a : `${a}:${s}` },
  };
}
function createWorkbookSnapshot(
  workbook,
  view,
  savedAt = new Date().toISOString(),
) {
  const workbookSnapshot = normalizeWorkbookSnapshot(workbook),
    savedSelection = normalizeSavedSelection(workbookSnapshot, view),
    a = { workbook: workbookSnapshot, view: savedSelection, savedAt: savedAt };
  return { ...a, digest: computeTextDigest(stringifyCanonicalJson(a)) };
}
function parseWorkbookSnapshot(e) {
  if (
    !isPlainRecord(e) ||
    !isPlainRecord(e.workbook) ||
    typeof e.savedAt != "string" ||
    !Number.isFinite(Date.parse(e.savedAt)) ||
    typeof e.digest != "string"
  )
    return throwPersistenceError(
      "corrupt",
      "Stored workbook snapshot is incomplete.",
    );
  if (
    computeTextDigest(
      stringifyCanonicalJson({
        workbook: e.workbook,
        view: e.view,
        savedAt: e.savedAt,
      }),
    ) !== e.digest
  )
    return throwPersistenceError(
      "corrupt",
      "Stored workbook snapshot failed its integrity check.",
    );
  try {
    return createWorkbookSnapshot(
      normalizeWorkbookSnapshot(e.workbook),
      e.view,
      e.savedAt,
    );
  } catch {
    return throwPersistenceError(
      "corrupt",
      "Stored workbook snapshot does not satisfy workbook validation.",
    );
  }
}
function createCheckpointRecord(e) {
  return {
    id: e.id,
    label: e.label,
    createdAt: e.createdAt,
    workbookId: e.snapshot.workbook.id,
    revision: e.snapshot.workbook.revision,
  };
}
function compareWorkbookVersions(e, workbook) {
  const workbookSnapshot = normalizeWorkbookSnapshot(e),
    n = normalizeWorkbookSnapshot(workbook);
  if (workbookSnapshot.id !== n.id)
    return throwPersistenceError(
      "invalid",
      "Version comparison requires the same workbook identity.",
    );
  const o = {
      fromRevision: workbookSnapshot.revision,
      toRevision: n.revision,
      cellChanges: [],
      metadataChanges: [],
      structuralChanges: [],
    },
    a = (m, v) => stringifyCanonicalJson(m) !== stringifyCanonicalJson(v),
    s = new Set(workbookSnapshot.events.map((m) => m.id)),
    d = n.events.filter((m) => !s.has(m.id)),
    u = new Map();
  for (const m of d)
    for (const change of m.changes)
      u.set(`${change.sheetId}:${change.address}`, {
        event: m,
        change: change,
      });
  const l = (m, basis) =>
      m
        ? {
            basis: basis,
            actor: m.actor,
            timestamp: m.timestamp,
            eventId: m.id,
          }
        : { basis: "unavailable" },
    c = (m) =>
      l(
        [...d].reverse().find((v) => m.test(v.action)),
        "operation-history",
      ),
    h = () =>
      c(
        /sheet|structure|tab color|visibility|restore|row|column|merge|drawing|picture|shape|chart|validation|conditional|page|table|pivot|query/i,
      );
  for (const path of new Set([
    ...Object.keys(workbookSnapshot),
    ...Object.keys(n),
  ])) {
    if (["sheets", "events", "revision", "id", "schemaVersion"].includes(path))
      continue;
    const v = workbookSnapshot[path],
      g = n[path];
    a(v, g) &&
      o.metadataChanges.push({
        path: path,
        before: cloneStructuredValue(v),
        after: cloneStructuredValue(g),
        attribution: c(
          path === "definedNames"
            ? /name|formula|restore/i
            : path === "formulaWatches"
              ? /watch|row|column|sheet|move|cut|restore/i
              : path === "styleCatalog"
                ? /theme|style|restore/i
                : /workbook|properties|engagement|period|template|restore/i,
        ),
      });
  }
  a(
    workbookSnapshot.sheets.map((sheet) => sheet.id),
    n.sheets.map((sheet) => sheet.id),
  ) &&
    o.structuralChanges.push({
      path: "sheets.order",
      before: workbookSnapshot.sheets.map((sheet) => sheet.id),
      after: n.sheets.map((sheet) => sheet.id),
      attribution: h(),
    });
  for (const sheetId of new Set([
    ...workbookSnapshot.sheets.map((sheet) => sheet.id),
    ...n.sheets.map((sheet) => sheet.id),
  ])) {
    const v = workbookSnapshot.sheets.find((sheet) => sheet.id === sheetId),
      g = n.sheets.find((sheet) => sheet.id === sheetId),
      b = (E) =>
        E
          ? Object.fromEntries(Object.entries(E).filter(([A]) => A !== "cells"))
          : undefined,
      before = b(v),
      after = b(g);
    if (!v || !g)
      o.structuralChanges.push({
        path: `sheets.${sheetId}`,
        before: before,
        after: after,
        attribution: h(),
      });
    else
      for (const E of new Set([...Object.keys(before), ...Object.keys(after)]))
        a(before[E], after[E]) &&
          o.structuralChanges.push({
            path: `sheets.${sheetId}.${E}`,
            before: cloneStructuredValue(before[E]),
            after: cloneStructuredValue(after[E]),
            attribution: h(),
          });
    for (const address of new Set([
      ...Object.keys(v?.cells ?? {}),
      ...Object.keys(g?.cells ?? {}),
    ])) {
      const A = v?.cells[address],
        k = g?.cells[address];
      a(A, k) &&
        o.cellChanges.push({
          sheetId: sheetId,
          address: address,
          ...(A === undefined ? {} : { before: cloneStructuredValue(A) }),
          ...(k === undefined ? {} : { after: cloneStructuredValue(k) }),
          attribution: l(
            stringifyCanonicalJson(
              u.get(`${sheetId}:${address}`)?.change.after,
            ) === stringifyCanonicalJson(k)
              ? u.get(`${sheetId}:${address}`)?.event
              : undefined,
            "cell-history",
          ),
        });
    }
  }
  return o;
}
function createLocalWorkbookStore(e, namespace, t = {}) {
  if (
    typeof namespace != "string" ||
    !namespace.trim() ||
    namespace.length > 300
  )
    return throwPersistenceError(
      "invalid",
      "A local workbook namespace must contain 1–300 characters.",
    );
  const n = t.maxCheckpoints ?? maxLocalCheckpoints,
    o = t.maxBytes ?? 10 * 1024 * 1024;
  if (
    !Number.isInteger(n) ||
    n < 1 ||
    n > maxLocalCheckpoints ||
    !Number.isInteger(o) ||
    o < 1024 ||
    o > 50 * 1024 * 1024
  )
    return throwPersistenceError(
      "invalid",
      "Use 1–20 checkpoints and a local envelope size between 1 KiB and 50 MiB.",
    );
  function a() {
    try {
      return e.getItem(namespace);
    } catch {
      return throwPersistenceError(
        "unavailable",
        "Browser storage is unavailable. Keep the workbook open and export your unsaved changes.",
      );
    }
  }
  function s() {
    const h = a();
    if (h === null) return null;
    if (h.length > o || measureUtf8Bytes(h) > o)
      return throwPersistenceError(
        "limit",
        "The saved local workbook exceeds this store's size limit. Preserve it for export or increase the configured limit.",
      );
    let f;
    try {
      f = JSON.parse(h);
    } catch {
      return throwPersistenceError(
        "corrupt",
        "The saved workbook cannot be read. Its original storage has been preserved; export in-memory work before choosing another namespace.",
      );
    }
    if (!isPlainRecord(f))
      return throwPersistenceError(
        "corrupt",
        "Invalid local workbook envelope.",
      );
    let envelope,
      recovered = false,
      warning;
    if (f.schemaVersion === 1) {
      let y;
      try {
        y = normalizeWorkbookSnapshot(f);
      } catch {
        return throwPersistenceError(
          "corrupt",
          "The legacy saved workbook is invalid; original storage has been preserved.",
        );
      }
      const workbookSnapshot = createWorkbookSnapshot(
        y,
        normalizeSavedSelection(y),
      );
      ((envelope = {
        format: localWorkbookNamespace,
        version: 1,
        generation: 0,
        current: workbookSnapshot,
        checkpoints: [],
      }),
        (warning =
          "Legacy workbook loaded. The next successful save will add local recovery and version metadata."));
    } else {
      if (
        f.format !== localWorkbookNamespace ||
        f.version !== 1 ||
        !Number.isSafeInteger(f.generation) ||
        Number(f.generation) < 1 ||
        !Array.isArray(f.checkpoints) ||
        f.checkpoints.length > maxLocalCheckpoints
      )
        return throwPersistenceError(
          "corrupt",
          "Unsupported or corrupt local workbook envelope; stored data was preserved.",
        );
      let y, previous;
      if (f.previous !== undefined)
        try {
          previous = parseWorkbookSnapshot(f.previous);
        } catch {
          warning =
            "The previous recovery snapshot is invalid; the current workbook remains available.";
        }
      try {
        y = parseWorkbookSnapshot(f.current);
      } catch {
        if (!previous)
          return throwPersistenceError(
            "corrupt",
            "Both saved workbook snapshots are invalid. Original storage has been preserved for recovery.",
          );
        ((y = previous),
          (recovered = true),
          (warning =
            "Recovered the previous complete save because the latest snapshot is invalid. Review it before saving; the original storage is unchanged."));
      }
      const E = new Set(),
        checkpoints = [];
      for (const checkpoint of f.checkpoints) {
        if (
          !isPlainRecord(checkpoint) ||
          typeof checkpoint.id != "string" ||
          !checkpoint.id ||
          checkpoint.id.length > 100 ||
          E.has(checkpoint.id) ||
          typeof checkpoint.label != "string" ||
          !checkpoint.label.trim() ||
          checkpoint.label.length > 120 ||
          typeof checkpoint.createdAt != "string" ||
          !Number.isFinite(Date.parse(checkpoint.createdAt))
        )
          return throwPersistenceError(
            "corrupt",
            "Saved checkpoint metadata is invalid; original storage was preserved.",
          );
        let snapshot;
        try {
          snapshot = parseWorkbookSnapshot(checkpoint.snapshot);
        } catch {
          return throwPersistenceError(
            "corrupt",
            "A saved checkpoint is corrupt; original storage was preserved.",
          );
        }
        if (snapshot.workbook.id !== y.workbook.id)
          return throwPersistenceError(
            "corrupt",
            "A checkpoint belongs to another workbook; original storage was preserved.",
          );
        (E.add(checkpoint.id),
          checkpoints.push({
            id: checkpoint.id,
            label: checkpoint.label,
            createdAt: checkpoint.createdAt,
            snapshot: snapshot,
          }));
      }
      envelope = {
        format: localWorkbookNamespace,
        version: 1,
        generation: Number(f.generation),
        current: y,
        ...(previous ? { previous: previous } : {}),
        checkpoints: checkpoints,
      };
    }
    const { workbook, view } = envelope.current;
    return {
      envelope: envelope,
      result: {
        workbook: cloneStructuredValue(workbook),
        view: cloneStructuredValue(view),
        guard: {
          namespace: namespace,
          generation: envelope.generation,
          workbookId: workbook.id,
          revision: workbook.revision,
          digest: computeTextDigest(h),
        },
        recovered: recovered,
        ...(warning ? { warning: warning } : {}),
        checkpoints: envelope.checkpoints.map(createCheckpointRecord),
      },
    };
  }
  function d(h) {
    const f = s();
    return h === null
      ? f
        ? throwPersistenceError(
            "conflict",
            "This local namespace already contains a workbook. Reload it before saving; unsaved work remains in memory.",
            f.result.guard,
          )
        : null
      : !f ||
          h.namespace !== namespace ||
          stringifyCanonicalJson(h) !== stringifyCanonicalJson(f.result.guard)
        ? throwPersistenceError(
            "conflict",
            "Another workbook window saved a different version. Reload or export your unsaved changes before resolving the conflict.",
            f?.result.guard,
          )
        : f;
  }
  function u(h, f) {
    const p = JSON.stringify(h);
    if (p.length > o || measureUtf8Bytes(p) > o)
      return throwPersistenceError(
        "limit",
        "Local version storage exceeds its configured size limit. Export the workbook or use a larger host store; the previous save is unchanged.",
      );
    d(f);
    try {
      e.setItem(namespace, p);
    } catch (error) {
      const v =
        error instanceof Error &&
        /quota|space|full/i.test(`${error.name} ${error.message}`);
      return throwPersistenceError(
        v ? "quota" : "unavailable",
        "Local save failed. The previous complete save is unchanged; keep this workbook open and export your unsaved changes.",
      );
    }
    return {
      workbook: cloneStructuredValue(h.current.workbook),
      view: cloneStructuredValue(h.current.view),
      guard: {
        namespace: namespace,
        generation: h.generation,
        workbookId: h.current.workbook.id,
        revision: h.current.workbook.revision,
        digest: computeTextDigest(p),
      },
      recovered: false,
      checkpoints: h.checkpoints.map(createCheckpointRecord),
    };
  }
  function l(h) {
    const f = (h?.envelope.generation ?? 0) + 1;
    return Number.isSafeInteger(f)
      ? f
      : throwPersistenceError(
          "limit",
          "Local version generation limit reached.",
        );
  }
  function c(h, f) {
    const p = h?.envelope.checkpoints.find((checkpoint) => checkpoint.id === f);
    return (
      p ||
      throwPersistenceError(
        "invalid",
        "This checkpoint is unavailable or has expired from the bounded local history.",
      )
    );
  }
  return {
    load: () => s()?.result ?? null,
    save(workbook, view, p, m = {}) {
      const v = d(p),
        workbookSnapshot = createWorkbookSnapshot(workbook, view),
        b = !v || v.result.workbook.id === workbookSnapshot.workbook.id;
      return !b && !m.allowIdentityChange
        ? throwPersistenceError(
            "conflict",
            "Opening another workbook requires an explicit namespace replacement or a separate namespace.",
            v?.result.guard,
          )
        : b &&
            v &&
            workbookSnapshot.workbook.revision < v.result.workbook.revision
          ? throwPersistenceError(
              "conflict",
              "An older workbook revision cannot overwrite a newer saved revision. Restore a checkpoint as a new revision instead.",
              v.result.guard,
            )
          : b &&
              v &&
              workbookSnapshot.workbook.revision ===
                v.result.workbook.revision &&
              stringifyCanonicalJson(workbookSnapshot.workbook) !==
                stringifyCanonicalJson(v.result.workbook)
            ? throwPersistenceError(
                "conflict",
                "Workbook content changed without a new revision. Commit the edit before saving.",
                v.result.guard,
              )
            : u(
                {
                  format: localWorkbookNamespace,
                  version: 1,
                  generation: l(v),
                  current: workbookSnapshot,
                  ...(v ? { previous: v.envelope.current } : {}),
                  checkpoints: b ? (v?.envelope.checkpoints ?? []) : [],
                },
                p,
              );
    },
    checkpoint(h, f) {
      if (typeof h != "string" || !h.trim() || h.length > 120)
        return throwPersistenceError(
          "invalid",
          "Checkpoint label must contain 1–120 characters.",
        );
      const p = d(f),
        m = {
          id: crypto.randomUUID(),
          label: h.trim(),
          createdAt: new Date().toISOString(),
          snapshot: cloneStructuredValue(p.envelope.current),
        },
        checkpoints = [...p.envelope.checkpoints, m].slice(-n),
        g = u({ ...p.envelope, generation: l(p), checkpoints: checkpoints }, f);
      return (
        p.envelope.checkpoints.length >= n &&
          (g.warning = `The oldest checkpoint expired. This local store retains the newest ${n} checkpoints.`),
        g
      );
    },
    listCheckpoints: () => s()?.result.checkpoints ?? [],
    compareCheckpoint: (h, workbook) =>
      compareWorkbookVersions(c(s(), h).snapshot.workbook, workbook),
    restoreCheckpoint(h, workbook, p, actor = "Auditor") {
      const v = d(p),
        workbookSnapshot = normalizeWorkbookSnapshot(workbook);
      if (
        workbookSnapshot.id !== v.result.workbook.id ||
        workbookSnapshot.revision !== v.result.workbook.revision ||
        stringifyCanonicalJson(workbookSnapshot) !==
          stringifyCanonicalJson(v.result.workbook)
      )
        return throwPersistenceError(
          "conflict",
          "Save or export your current unsaved changes before restoring a checkpoint.",
          v.result.guard,
        );
      if (
        workbookSnapshot.structureProtection?.enabled ||
        workbookSnapshot.sheets.some((sheet) => sheet.protection?.enabled)
      )
        return throwPersistenceError(
          "invalid",
          "Unlock the workbook structure and protected worksheets before restoring a version.",
        );
      if (typeof actor != "string" || !actor.trim() || actor.length > 100)
        return throwPersistenceError(
          "invalid",
          "A restore actor must contain 1–100 characters.",
        );
      const b = c(v, h),
        y = compareWorkbookVersions(workbookSnapshot, b.snapshot.workbook),
        timestamp = new Date().toISOString();
      let E = cloneStructuredValue(b.snapshot.workbook);
      if (
        ((E.revision = workbookSnapshot.revision + 1),
        !Number.isSafeInteger(E.revision) ||
          E.revision >= Number.MAX_SAFE_INTEGER)
      )
        return throwPersistenceError(
          "limit",
          "Workbook revision limit reached.",
        );
      (workbookSnapshot.metadata &&
        (E.metadata = {
          ...(E.metadata ?? workbookSnapshot.metadata),
          createdAt: workbookSnapshot.metadata.createdAt,
          modifiedAt: [
            workbookSnapshot.metadata.modifiedAt,
            E.metadata?.modifiedAt ?? timestamp,
            timestamp,
          ]
            .sort()
            .at(-1),
        }),
        (E = touchWorkbookModifiedAt(E, timestamp)));
      const A = Math.max(1, Math.ceil(y.cellChanges.length / 500));
      return (
        (E.events = [
          ...workbookSnapshot.events,
          ...Array.from({ length: A }, (I, O) => ({
            id: crypto.randomUUID(),
            timestamp: timestamp,
            actor: actor,
            action: `Restored checkpoint “${b.label}” from revision ${b.snapshot.workbook.revision}${A > 1 ? ` (${O + 1}/${A})` : ""}`,
            changes: y.cellChanges
              .slice(O * 500, (O + 1) * 500)
              .map(({ attribution, ...z }) => z),
          })),
        ]),
        (E = normalizeWorkbookSnapshot(E)),
        u(
          {
            ...v.envelope,
            generation: l(v),
            previous: v.envelope.current,
            current: createWorkbookSnapshot(E, b.snapshot.view, timestamp),
          },
          p,
        )
      );
    },
  };
}
function createWorkbookSession(e = {}) {
  let r = null,
    t = null;
  const namespace = `session-${crypto.randomUUID()}`,
    localWorkbookStore = createLocalWorkbookStore(
      {
        getItem: () => r,
        setItem: (f, p) => {
          r = p;
        },
      },
      namespace,
      e,
    );
  let a = null;
  const s = new Map(),
    d = e.maxBytes ?? 10 * 1024 * 1024,
    u = (f, p) => f === p || JSON.stringify(f) === JSON.stringify(p);
  function l(f) {
    if (!u(f, t?.guard ?? null))
      throw new WorkbookRevisionConflictError(
        "conflict",
        "The saved session changed. Reload before replacing it.",
        t?.guard,
      );
  }
  function c(f) {
    const generation = (t?.guard.generation ?? 0) + 1;
    if (!Number.isSafeInteger(generation))
      throw new WorkbookRevisionConflictError(
        "limit",
        "Session generation limit reached.",
      );
    return (
      (t = deepFreezeValue({
        ...f,
        workbook: buildCanonicalWorkbookSnapshot(f.workbook),
        view: structuredClone(f.view),
        guard: {
          namespace: namespace,
          generation: generation,
          workbookId: f.workbook.id,
          revision: f.workbook.revision,
          digest: crypto.randomUUID(),
        },
      })),
      t
    );
  }
  function h() {
    if (!t)
      throw new WorkbookRevisionConflictError(
        "invalid",
        "Save this session before creating a version.",
      );
    const f = localWorkbookStore.save(t.workbook, t.view, a, {
      allowIdentityChange: true,
    });
    return ((a = f.guard), f);
  }
  return {
    load: () => t,
    save(f, p, guard, v = {}) {
      l(guard);
      const canonicalWorkbookSnapshot = buildCanonicalWorkbookSnapshot(f),
        b = canonicalWorkbookSnapshot.sheets.find(
          (sheet) => sheet.id === p.activeSheetId && !sheet.hidden,
        );
      if (!b || p.selection.sheetId !== b.id)
        throw new WorkbookRevisionConflictError(
          "invalid",
          "Saved selection must belong to the active visible worksheet.",
        );
      const y = !t || t.workbook.id === canonicalWorkbookSnapshot.id;
      if (!y && !v.allowIdentityChange)
        throw new WorkbookRevisionConflictError(
          "conflict",
          "Opening another workbook requires an explicit session replacement.",
          t?.guard,
        );
      if (
        y &&
        t &&
        (canonicalWorkbookSnapshot.revision < t.workbook.revision ||
          (canonicalWorkbookSnapshot.revision === t.workbook.revision &&
            !u(canonicalWorkbookSnapshot, t.workbook)))
      )
        throw new WorkbookRevisionConflictError(
          "conflict",
          "Commit a new revision before replacing saved content.",
          t.guard,
        );
      const view = validateSavedSelection(canonicalWorkbookSnapshot, p),
        E = (workbook, O) => ({
          workbook: workbook,
          view: O,
          savedAt: "1970-01-01T00:00:00.000Z",
          digest: "0".repeat(40),
        }),
        A = y ? (t?.checkpoints ?? []) : [],
        k = A.reduce((I, O) => I + (s.get(O.id) ?? 0), 0);
      if (
        measureSerializedBytes({
          format: "ledgerglass-local-workbook",
          version: 1,
          generation: (t?.guard.generation ?? 0) + 1,
          current: E(canonicalWorkbookSnapshot, view),
          ...(t ? { previous: E(t.workbook, t.view) } : {}),
          checkpoints: [],
        }) +
          k +
          Math.max(0, A.length - 1) >
        d
      )
        throw new WorkbookRevisionConflictError(
          "limit",
          "Session version storage exceeds its configured size limit. Export the workbook or use a larger host store; the previous save is unchanged.",
        );
      return (
        y || s.clear(),
        c({
          workbook: canonicalWorkbookSnapshot,
          view: view,
          guard: guard,
          recovered: false,
          checkpoints: y ? (t?.checkpoints ?? []) : [],
        })
      );
    },
    checkpoint(f, p) {
      l(p);
      const m = h(),
        v = localWorkbookStore.checkpoint(f, m.guard);
      a = v.guard;
      const g = new Set(v.checkpoints.map((checkpoint) => checkpoint.id));
      for (const b of s.keys()) g.has(b) || s.delete(b);
      for (const b of g)
        s.has(b) ||
          s.set(
            b,
            measureSerializedBytes({
              id: b,
              label: v.checkpoints.find((checkpoint) => checkpoint.id === b)
                .label,
              createdAt: "1970-01-01T00:00:00.000Z",
              snapshot: {
                workbook: v.workbook,
                view: v.view,
                savedAt: "1970-01-01T00:00:00.000Z",
                digest: "0".repeat(40),
              },
            }),
          );
      return c(v);
    },
    listCheckpoints: () => structuredClone(t?.checkpoints ?? []),
    compareCheckpoint: (f, p) => localWorkbookStore.compareCheckpoint(f, p),
    restoreCheckpoint(f, p, m, v) {
      if ((l(m), !t || !u(p, t.workbook)))
        throw new WorkbookRevisionConflictError(
          "conflict",
          "Save the current revision before restoring a checkpoint.",
          t?.guard,
        );
      const g = h(),
        b = localWorkbookStore.restoreCheckpoint(f, p, g.guard, v);
      return ((a = b.guard), c(b));
    },
  };
}
function validateSavedSelection(canonicalWorkbookSnapshot, r) {
  const t = canonicalWorkbookSnapshot.sheets.find(
      (sheet) => sheet.id === r.activeSheetId,
    ),
    rangeAddress = parseRangeAddress(r.selection.range);
  if (rangeAddress.lastRow >= t.rowCount || rangeAddress.lastCol >= t.colCount)
    throw new WorkbookRevisionConflictError(
      "invalid",
      "Saved selection exceeds worksheet dimensions.",
    );
  const o = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    a = `${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
  return {
    activeSheetId: t.id,
    selection: { sheetId: t.id, range: o === a ? o : `${o}:${a}` },
  };
}
const buildCompositeKey = (e, r) => JSON.stringify([e, r]),
  isFormulaInputValue = (cell) =>
    !!cell &&
    !cell.literal &&
    typeof cell.raw == "string" &&
    cell.raw.startsWith("="),
  bareSheetNamePattern = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/,
  sheetNamePattern = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_\\][A-Za-z0-9_.\\]*)`,
  cellRangePattern = String.raw`(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?:\s*:\s*\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}\s*:\s*\$?[A-Za-z]{1,3}|\$?[1-9]\d*\s*:\s*\$?[1-9]\d*)`,
  singleRangeReferencePattern = new RegExp(
    `^(${sheetNamePattern}\\s*!\\s*)?(${cellRangePattern})(#)?`,
  ),
  sheetRangeReferencePattern = new RegExp(
    `^${sheetNamePattern}\\s*:\\s*${sheetNamePattern}\\s*!\\s*${cellRangePattern}`,
  ),
  knownFunctionNameSet = new Set(
    builtinFunctionNames.map((builtinFunctionName) =>
      builtinFunctionName.toUpperCase(),
    ),
  ),
  requireIntegerInRange = (e, r, t, n) => {
    const o = e ?? r;
    if (!Number.isInteger(o) || o < 1 || o > t)
      throw Error(`${n} must be an integer from 1 to ${t}.`);
    return o;
  };
function parseCellReference(address) {
  const r = address.replace(/[\s$]/g, "").toUpperCase(),
    t = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(r),
    n = /^([1-9]\d*):([1-9]\d*)$/.exec(r),
    rangeAddress = parseRangeAddress(
      t
        ? `${t[1]}1:${t[2]}${workbookLimits.maxRows}`
        : n
          ? `A${n[1]}:XFD${n[2]}`
          : r,
    ),
    cellAddress = formatCellAddress(
      rangeAddress.firstRow,
      rangeAddress.firstCol,
    ),
    s = formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol);
  return cellAddress === s ? cellAddress : `${cellAddress}:${s}`;
}

export {
  FormatConversionDialog,
  createLocalWorkbookStore,
  requireIntegerInRange,
  singleRangeReferencePattern,
  WorkbookRevisionConflictError,
  buildCompositeKey,
  parseCellReference,
  bareSheetNamePattern,
  knownFunctionNameSet,
  isFormulaInputValue,
  compareWorkbookVersions,
  sheetRangeReferencePattern,
  createWorkbookSession,
};
export { cloneStructuredValue, isPlainRecord };
