import { unzipSync } from "fflate";
import ne from "react";
import i from "react/jsx-runtime";
import { X } from "lucide-react";
import { string, number, array, literal, unknown, object } from "zod";
import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  validateRetainedSourceFile,
  assertWorkbookStructureEditable,
  sourceImportLimits,
  requireRetainedXlsxSource,
  validateExtractedSourceTables,
  normalizeWorkbookInput,
  workbookLinkLimits,
  columnIndexToLetters,
  resolveExtractedCellValue,
  buildRetainedSourceRecord,
} from "./03-workbook-model.js";
import {
  assertSourceInspectionCurrent,
  applyWorkbookLinkPreview,
  recordSourceInspection,
  previewWorkbookLinkOperation,
  buildWorkbookLinkReview,
} from "./05-add-ins-and-custom-functions.js";
import {
  importWorkbookFile,
  runFileImport,
} from "./07-file-import-and-worker-dispatch.js";
import {
  F$,
  buildExchangeSchemaReport,
  importXlsxWorkbook,
  createEmptyExportLedger,
} from "./08-xlsx-import-export-and-preservation.js";

function buildSourceImportRecord(activeWorkbook) {
  const r = requireRetainedXlsxSource(activeWorkbook),
    xlsxWorkbookPackage = importXlsxWorkbook(
      r.bytes,
      createEmptyExportLedger(),
    ),
    exchangeSchemaReport = buildExchangeSchemaReport(
      "import",
      r.fileName,
      r.bytes.length,
      activeWorkbook,
      xlsxWorkbookPackage,
    );
  return (
    (exchangeSchemaReport.partInventory = F$(unzipSync(r.bytes), r.sha256)),
    exchangeSchemaReport
  );
}
const operationScopeContext = ne.createContext({}),
  useOperationScope = () => ne.useContext(operationScopeContext);
function useCancellableOperation() {
  const operationScope = useOperationScope(),
    r = ne.useRef(null),
    cancel = ne.useCallback(() => r.current?.abort(), []);
  return (
    ne.useEffect(() => () => cancel(), [cancel]),
    {
      cancel: cancel,
      begin: () => {
        cancel();
        const n = new AbortController();
        return (
          (r.current = n),
          {
            ...operationScope,
            signal: operationScope.signal
              ? AbortSignal.any([n.signal, operationScope.signal])
              : n.signal,
          }
        );
      },
    }
  );
}
function WorkbookLinkStatusBadge({ workbook, sources, onReview }) {
  const n = ne.useMemo(() => {
    try {
      const workbookLinkReview = buildWorkbookLinkReview(workbook, { sources });
      return !workbookLinkReview.links.length &&
        !workbookLinkReview.unsupported.length
        ? null
        : `${workbookLinkReview.links.length} workbook link targets · ${workbookLinkReview.links.filter((link) => link.status !== "matches-approved-snapshot").length} need source review · ${workbookLinkReview.unsupported.length} unsupported. Displayed linked values are cached snapshots.`;
    } catch {
      return "Workbook link inventory needs review; links exceed the supported review limit.";
    }
  }, [workbook, sources]);
  return n
    ? i.jsxs("div", {
        className: "lg-workbook-link-status",
        role: "status",
        children: [
          i.jsx("span", { children: n }),
          i.jsx("button", {
            onClick: onReview,
            children: "Review workbook links",
          }),
        ],
      })
    : null;
}
function WorkbookLinksDialog(e) {
  const r = ne.useRef(e);
  r.current = e;
  const t = ne.useRef(null),
    cancellableOperation = useCancellableOperation(),
    o = ne.useRef(0),
    a = ne.useRef(true),
    [s, d] = ne.useState([]),
    [u, l] = ne.useState(""),
    [replacementTarget, h] = ne.useState(""),
    [action, p] = ne.useState("refresh"),
    [m, v] = ne.useState(false),
    [g, b] = ne.useState(false),
    [y, w] = ne.useState(null),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(0),
    [j, z] = ne.useState(0),
    [disabled, D] = ne.useState(false),
    [workbook, _] = ne.useState(null),
    [target, T] = ne.useState(""),
    [allowedSheetIds, R] = ne.useState([]),
    [L, $] = ne.useState([]),
    [checked, G] = ne.useState(false),
    sources = e.approvedSources ?? s,
    Q = ne.useMemo(() => {
      try {
        return {
          inventory: buildWorkbookLinkReview(e.workbook, { sources: sources }),
          error: "",
        };
      } catch (error) {
        return { inventory: null, error: error.message };
      }
    }, [e.workbook, sources]),
    X$1 = Q.inventory,
    J = X$1?.links.find((link) => link.target === (u || X$1.links[0]?.target)),
    q = J?.target ?? "",
    W =
      X$1?.cells.filter((cell) =>
        cell.references.some((reference) => reference.target === q),
      ) ?? [],
    K = (se) => {
      if (e.approvedSources !== undefined && !e.onApprovedSourcesChange)
        throw Error("Source approvals are controlled by your host.");
      (e.onApprovedSourcesChange?.(se),
        e.approvedSources === undefined && d(se),
        w(null));
    },
    te = () => ({
      sources: r.current.approvedSources ?? s,
      readOnly: r.current.readOnly,
      allowedSheetIds: r.current.allowedSheetIds,
    });
  (ne.useEffect(() => {
    a.current = true;
    const activeElement = document.activeElement;
    return (
      t.current?.focus(),
      () => {
        ((a.current = false),
          o.current++,
          cancellableOperation.cancel(),
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      w(null);
    }, [e.workbook, e.approvedSources]));
  const ae = (se) => {
      (A(""), F(""));
      try {
        se();
      } catch (error) {
        A(
          error instanceof Error
            ? error.message
            : "Workbook link operation failed.",
        );
      }
    },
    onClick = () =>
      ae(() => {
        (w(
          previewWorkbookLinkOperation(
            e.workbook,
            {
              action: action,
              target: q,
              replacementTarget: replacementTarget,
              activateLiteral: m,
              acknowledgeCached: g,
            },
            te(),
          ),
        ),
          O(0),
          z(0));
      }),
    de = () =>
      ae(() => {
        if (!y) return;
        const se = applyWorkbookLinkPreview(r.current.workbook, y, te());
        (r.current.onWorkbookChange(se, `${y.request.action} workbook links`),
          w(null),
          F(
            "Reviewed workbook link changes applied. Undo restores the prior workbook.",
          ));
      }),
    he = async (se) => {
      const we = ++o.current;
      (D(true), A(""), _(null), w(null));
      try {
        const Ee = await importWorkbookFile(se, cancellableOperation.begin());
        if (!a.current || we !== o.current) return;
        (_(Ee.workbook),
          T(q || `[${se.name}]`),
          R(
            Ee.workbook.sheets
              .filter((sheet) => !sheet.hidden && !sheet.veryHidden)
              .map((Ae) => Ae.id),
          ),
          $(Ee.warnings),
          G(false));
      } catch (error) {
        a.current && we === o.current && A(error.message);
      } finally {
        a.current && we === o.current && D(false);
      }
    },
    Se =
      e.allowFileSources !== false &&
      (e.approvedSources === undefined || !!e.onApprovedSourcesChange),
    le = () => {
      (w(null), A(""), F(""), O(0));
    };
  return i.jsx("div", {
    className: "lg-reports-overlay",
    children: i.jsxs("section", {
      className: "lg-reports-dialog lg-workbook-links",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": "Workbook links",
      tabIndex: -1,
      ref: t,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.preventDefault(), e.onClose()),
          event.key === "Tab")
        ) {
          const Ee = Array.from(
              t.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ).filter((Te) => !Te.hidden),
            Ae = Ee[0],
            ye = Ee.at(-1);
          event.shiftKey &&
          (document.activeElement === Ae ||
            document.activeElement === t.current)
            ? (event.preventDefault(), ye == null || ye.focus())
            : !event.shiftKey &&
              document.activeElement === ye &&
              (event.preventDefault(), Ae == null || Ae.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("p", { children: "DATA & WORKBOOK CONNECTIONS" }),
                i.jsx("h2", { children: "Workbook links" }),
              ],
            }),
            i.jsx("button", {
              "aria-label": "Close workbook links",
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
                "Inspect and review linked values before refreshing, changing a source or breaking links. Workbook paths are labels; this editor never fetches them.",
            }),
            (E || Q.error) &&
              i.jsx("p", {
                role: "alert",
                className: "lg-links-error",
                children: E || Q.error,
              }),
            k && i.jsx("p", { role: "status", children: k }),
            i.jsxs("div", {
              className: "lg-links-summary",
              children: [
                i.jsxs("strong", {
                  children: [X$1?.links.length ?? 0, " linked sources"],
                }),
                i.jsx("span", { children: "Manual refresh only" }),
                i.jsx("span", {
                  children:
                    "Native JSON retains cached values and revision evidence",
                }),
              ],
            }),
            X$1?.links.length
              ? i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("label", {
                      children: [
                        "Linked workbook",
                        i.jsx("select", {
                          "aria-label": "Linked workbook",
                          value: q,
                          onChange: (event) => {
                            (l(event.target.value), v(false), b(false), le());
                          },
                          children: X$1.links.map((link) =>
                            i.jsxs(
                              "option",
                              {
                                value: link.target,
                                children: [link.target, " · ", link.status],
                              },
                              link.target,
                            ),
                          ),
                        }),
                      ],
                    }),
                    i.jsxs("section", {
                      className: "lg-links-card",
                      children: [
                        i.jsx("h3", { children: "Source and cache" }),
                        i.jsxs("dl", {
                          children: [
                            i.jsx("dt", { children: "Target workbook" }),
                            i.jsx("dd", { children: q }),
                            J?.packageTarget &&
                              i.jsxs(i.Fragment, {
                                children: [
                                  i.jsx("dt", {
                                    children: "Original XLSX target",
                                  }),
                                  i.jsx("dd", { children: J.packageTarget }),
                                ],
                              }),
                            i.jsx("dt", { children: "Status" }),
                            i.jsx("dd", {
                              children: J?.status,
                            }),
                            i.jsx("dt", { children: "Last resolved revision" }),
                            i.jsx("dd", {
                              children:
                                J?.lastResolvedRevision ?? "Never resolved",
                            }),
                            i.jsx("dt", { children: "Resolved at" }),
                            i.jsx("dd", {
                              children: J?.lastResolvedAt ?? "Not available",
                            }),
                            i.jsx("dt", { children: "Source snapshot" }),
                            i.jsx("dd", {
                              children: J?.sourceName ?? "No approved result",
                            }),
                          ],
                        }),
                        i.jsx("p", { children: X$1.cachePolicy }),
                        J?.cached &&
                          J.status !== "matches-approved-snapshot" &&
                          i.jsx("p", {
                            className: "lg-links-warning",
                            children:
                              "The previous cached values remain in this workbook. They are not verified against a currently authorized source.",
                          }),
                      ],
                    }),
                    i.jsxs("section", {
                      children: [
                        i.jsx("h3", { children: "Linked cells and targets" }),
                        i.jsx("div", {
                          className: "lg-links-table",
                          children: i.jsxs("table", {
                            children: [
                              i.jsx("thead", {
                                children: i.jsxs("tr", {
                                  children: [
                                    i.jsx("th", { children: "Destination" }),
                                    i.jsx("th", {
                                      children: "Source worksheet / range",
                                    }),
                                    i.jsx("th", { children: "Formula" }),
                                    i.jsx("th", { children: "State" }),
                                  ],
                                }),
                              }),
                              i.jsx("tbody", {
                                children: W.slice(I * 20, I * 20 + 20).map(
                                  (se) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [
                                              se.sheetName,
                                              "!",
                                              se.address,
                                            ],
                                          }),
                                          i.jsx("td", {
                                            children: se.references
                                              .filter(
                                                (reference) =>
                                                  reference.target === q,
                                              )
                                              .map(
                                                (we) =>
                                                  `${we.sheet}!${we.range}`,
                                              )
                                              .join(", "),
                                          }),
                                          i.jsx("td", { children: se.formula }),
                                          i.jsx("td", {
                                            children: se.literal
                                              ? "Inert text — activation required"
                                              : "Formula",
                                          }),
                                        ],
                                      },
                                      se.sheetId + se.address,
                                    ),
                                ),
                              }),
                            ],
                          }),
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "lg-links-actions",
                      children: [
                        i.jsx("button", {
                          disabled: I === 0,
                          onClick: () => O(I - 1),
                          children: "Previous linked cells",
                        }),
                        i.jsxs("span", {
                          children: [
                            "Page ",
                            I + 1,
                            " of ",
                            Math.max(1, Math.ceil(W.length / 20)),
                          ],
                        }),
                        i.jsx("button", {
                          disabled: (I + 1) * 20 >= W.length,
                          onClick: () => O(I + 1),
                          children: "Next linked cells",
                        }),
                      ],
                    }),
                    i.jsxs("section", {
                      className: "lg-links-card",
                      children: [
                        i.jsx("h3", { children: "Review a change" }),
                        i.jsxs("label", {
                          children: [
                            "Operation",
                            i.jsxs("select", {
                              "aria-label": "Link operation",
                              value: action,
                              onChange: (event) => {
                                (p(event.target.value), le());
                              },
                              children: [
                                i.jsx("option", {
                                  value: "refresh",
                                  children: "Resolve / refresh approved source",
                                }),
                                i.jsx("option", {
                                  value: "replace",
                                  children: "Replace source workbook",
                                }),
                                i.jsx("option", {
                                  value: "break",
                                  children: "Break links and keep values",
                                }),
                              ],
                            }),
                          ],
                        }),
                        action === "replace" &&
                          i.jsxs("label", {
                            children: [
                              "Replacement source",
                              i.jsxs("select", {
                                "aria-label": "Replacement source",
                                value: replacementTarget,
                                onChange: (event) => {
                                  (h(event.target.value), le());
                                },
                                children: [
                                  i.jsx("option", {
                                    value: "",
                                    children: "Choose an approved source",
                                  }),
                                  sources
                                    .filter((event) => event.target !== q)
                                    .map((se) =>
                                      i.jsxs(
                                        "option",
                                        {
                                          value: se.target,
                                          children: [
                                            se.target,
                                            " · ",
                                            se.workbook.name,
                                          ],
                                        },
                                        se.target,
                                      ),
                                    ),
                                ],
                              }),
                            ],
                          }),
                        W.some((se) => se.literal) &&
                          i.jsxs("label", {
                            className: "lg-links-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: m,
                                onChange: (event) => {
                                  (v(event.target.checked), le());
                                },
                              }),
                              "Activate these inert formulas using only reviewed cached source values",
                            ],
                          }),
                        action === "break" &&
                          i.jsxs("label", {
                            className: "lg-links-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: g,
                                onChange: (event) => {
                                  (b(event.target.checked), le());
                                },
                              }),
                              "I understand that complete formulas will be replaced with the previewed cached values",
                            ],
                          }),
                        i.jsx("button", {
                          className: "lg-reports-primary",
                          disabled: e.readOnly || disabled,
                          onClick: onClick,
                          children:
                            action === "refresh"
                              ? "Preview refresh"
                              : action === "replace"
                                ? "Preview replacement"
                                : "Preview break links",
                        }),
                      ],
                    }),
                  ],
                })
              : i.jsx("p", {
                  children:
                    "No supported external workbook references were found. Static links such as =SUM('[Source.xlsx]Sheet1'!A1:A3) can be reviewed here.",
                }),
            !!X$1?.unsupported.length &&
              i.jsxs("section", {
                className: "lg-links-warning",
                children: [
                  i.jsx("h3", { children: "Unsupported external syntax" }),
                  X$1.unsupported.slice(0, 20).map((se) =>
                    i.jsxs(
                      "p",
                      {
                        children: [
                          se.sheetName,
                          "!",
                          se.address,
                          ": ",
                          se.unsupported,
                        ],
                      },
                      se.sheetId + se.address,
                    ),
                  ),
                ],
              }),
            !!e.workbook.workbookLinks?.packageTargets?.length &&
              i.jsxs("details", {
                children: [
                  i.jsx("summary", {
                    children: "Retained XLSX target inventory",
                  }),
                  e.workbook.workbookLinks.packageTargets.map((packageTarget) =>
                    i.jsxs(
                      "p",
                      {
                        children: [
                          packageTarget.key,
                          " → ",
                          packageTarget.target,
                          " (",
                          packageTarget.part,
                          "). Retained original metadata; never fetched.",
                        ],
                      },
                      packageTarget.key,
                    ),
                  ),
                ],
              }),
            i.jsxs("section", {
              className: "lg-links-card",
              children: [
                i.jsx("h3", { children: "Approved source snapshots" }),
                sources.length
                  ? i.jsx("ul", {
                      children: sources.map((source) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              source.target,
                              " · ",
                              source.workbook.name,
                              " · revision",
                              " ",
                              source.workbook.revision,
                              " · ",
                              source.allowedSheetIds.length,
                              " approved worksheets",
                              " ",
                              (!e.approvedSources ||
                                e.onApprovedSourcesChange) &&
                                i.jsxs("button", {
                                  onClick: () =>
                                    ae(() =>
                                      K(sources.filter((we) => we !== source)),
                                    ),
                                  children: [
                                    "Remove approval for ",
                                    source.target,
                                  ],
                                }),
                            ],
                          },
                          source.target,
                        ),
                      ),
                    })
                  : i.jsx("p", {
                      children:
                        "No source workbook is approved in this session.",
                    }),
                Se
                  ? i.jsxs("label", {
                      children: [
                        "Attach source workbook",
                        i.jsx("input", {
                          "aria-label": "Attach source workbook",
                          type: "file",
                          accept: ".xlsx,.xltx,.json,.csv,.tsv",
                          disabled: disabled,
                          onChange: (event) => {
                            const we = event.target.files?.[0];
                            ((event.target.value = ""), we && he(we));
                          },
                        }),
                      ],
                    })
                  : i.jsx("p", {
                      children:
                        "Your host supplies source approvals. Ask the host application to attach an authorized snapshot.",
                    }),
                disabled &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsx("p", {
                        role: "status",
                        children: "Inspecting selected source…",
                      }),
                      i.jsx("button", {
                        onClick: () => {
                          (o.current++,
                            cancellableOperation.cancel(),
                            D(false),
                            _(null));
                        },
                        children: "Cancel source inspection",
                      }),
                    ],
                  }),
                workbook &&
                  i.jsxs("div", {
                    className: "lg-links-pending",
                    children: [
                      i.jsxs("h4", { children: ["Approve ", workbook.name] }),
                      i.jsxs("label", {
                        children: [
                          "Exact formula qualifier",
                          i.jsx("input", {
                            "aria-label": "Source qualifier",
                            value: target,
                            onChange: (event) => T(event.target.value),
                          }),
                        ],
                      }),
                      i.jsx("p", {
                        children:
                          "Match the selected reference exactly. A matching filename alone does not grant access.",
                      }),
                      workbook.sheets.map((sheet) =>
                        i.jsxs(
                          "label",
                          {
                            className: "lg-links-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: allowedSheetIds.includes(sheet.id),
                                disabled: !!(sheet.hidden || sheet.veryHidden),
                                onChange: (event) =>
                                  R(
                                    event.target.checked
                                      ? [...allowedSheetIds, sheet.id]
                                      : allowedSheetIds.filter(
                                          (allowedSheetId) =>
                                            allowedSheetId !== sheet.id,
                                        ),
                                  ),
                              }),
                              sheet.name,
                              sheet.hidden || sheet.veryHidden
                                ? " (hidden; host approval required)"
                                : "",
                            ],
                          },
                          sheet.id,
                        ),
                      ),
                      L.length > 0 &&
                        i.jsxs("details", {
                          children: [
                            i.jsxs("summary", {
                              children: [L.length, " import notices"],
                            }),
                            L.map((se, we) => i.jsx("p", { children: se }, we)),
                          ],
                        }),
                      i.jsxs("label", {
                        className: "lg-links-check",
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: checked,
                            onChange: (event) => G(event.target.checked),
                          }),
                          "I approve this source and the selected worksheets for explicit refresh",
                        ],
                      }),
                      i.jsx("button", {
                        disabled: !checked || !allowedSheetIds.length,
                        onClick: () =>
                          ae(() => {
                            const se = [
                              ...sources.filter(
                                (event) => event.target !== target,
                              ),
                              {
                                target: target,
                                workbook: workbook,
                                allowedSheetIds: allowedSheetIds,
                              },
                            ];
                            (buildWorkbookLinkReview(e.workbook, {
                              sources: se,
                            }),
                              K(se),
                              _(null));
                          }),
                        children: "Approve selected source",
                      }),
                      i.jsx("button", {
                        onClick: () => _(null),
                        children: "Discard source",
                      }),
                    ],
                  }),
              ],
            }),
            y &&
              i.jsxs("section", {
                className: "lg-links-confirm",
                role: "region",
                "aria-label": "Workbook link preview",
                children: [
                  i.jsx("h3", {
                    children:
                      y.request.action === "break"
                        ? "Values that will replace formulas"
                        : "Reviewed results",
                  }),
                  i.jsx("p", { children: y.warning }),
                  i.jsx("div", {
                    className: "lg-links-table",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", { children: "Cell" }),
                              i.jsx("th", { children: "Before" }),
                              i.jsx("th", { children: "After" }),
                              i.jsx("th", { children: "Calculated value" }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: y.cells
                            .slice(j * 20, j * 20 + 20)
                            .map((se) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsxs("td", {
                                      children: [se.sheetName, "!", se.address],
                                    }),
                                    i.jsx("td", { children: se.before }),
                                    i.jsx("td", {
                                      children: String(se.after ?? "(blank)"),
                                    }),
                                    i.jsxs("td", {
                                      children: [
                                        se.result.type,
                                        " ·",
                                        " ",
                                        String(se.result.value ?? "(blank)"),
                                      ],
                                    }),
                                  ],
                                },
                                se.sheetId + se.address,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "lg-links-actions",
                    children: [
                      i.jsx("button", {
                        disabled: j === 0,
                        onClick: () => z(j - 1),
                        children: "Previous preview cells",
                      }),
                      i.jsxs("span", {
                        children: [
                          y.cells.length,
                          " affected cells · preview page",
                          " ",
                          j + 1,
                          " of ",
                          Math.ceil(y.cells.length / 20),
                        ],
                      }),
                      i.jsx("button", {
                        disabled: (j + 1) * 20 >= y.cells.length,
                        onClick: () => z(j + 1),
                        children: "Next preview cells",
                      }),
                    ],
                  }),
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: e.readOnly || disabled,
                    onClick: de,
                    children: "Apply reviewed link changes",
                  }),
                  i.jsx("button", {
                    onClick: () => w(null),
                    children: "Cancel link preview",
                  }),
                ],
              }),
            i.jsxs("p", {
              children: [
                "Limits: ",
                workbookLinkLimits.sources,
                " sources,",
                " ",
                workbookLinkLimits.cells.toLocaleString(),
                " cached source cells and ",
                workbookLinkLimits.formulas,
                " destination formulas. Static A1 cells/ranges only. Linked arrays, external names, DDE and automatic/network refresh are unsupported. XLSX exports link formulas as inert text; native JSON retains the reviewed cache. Breaking links does not alter an attached original file.",
              ],
            }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children: e.readOnly
                ? "Read-only: inspection remains available."
                : "Every mutation requires a preview and confirmation.",
            }),
            i.jsx("button", { onClick: e.onClose, children: "Close" }),
          ],
        }),
      ],
    }),
  });
}
function Rue(e, r) {
  return runFileImport("source", e, r);
}
function toSourceFileDescriptor(inspection) {
  return {
    fileName: inspection.fileName,
    bytes: assertSourceInspectionCurrent(inspection),
    sha256: inspection.sha256,
  };
}
function parseExtractionResult(inspection, r) {
  const t = assertSourceInspectionCurrent(inspection);
  if (inspection.method !== "document-adapter")
    throw Error("JSON/XML use the native structured parser.");
  const n = object({
    schemaVersion: literal(1),
    sourceSha256: string(),
    extractor: string().min(1).max(200),
    pageCount: number().int().min(1).max(1e4),
    tables: array(unknown()).min(1).max(sourceImportLimits.tables),
  })
    .strict()
    .parse(r);
  if (n.sourceSha256 !== inspection.sha256)
    throw Error("Extraction does not match the selected source hash.");
  if (["png", "jpeg"].includes(inspection.format) && n.pageCount !== 1)
    throw Error("An image has one source page.");
  const o = validateExtractedSourceTables({
    ...inspection,
    extractor: n.extractor,
    pageCount: n.pageCount,
    tables: n.tables,
    warnings: [
      ...inspection.warnings,
      "Extraction values, confidence and regions are adapter assertions. Review the original; hashes do not prove extraction accuracy.",
    ],
  });
  for (const table of o.tables)
    for (const row of table.rows)
      for (const d of row)
        if (d.region.page > n.pageCount)
          throw Error("A source region exceeds the declared page count.");
  return (
    o.tables.length > 1 &&
      o.tables.forEach((table) => (table.ambiguous = true)),
    recordSourceInspection(o, t, inspection.fileName),
    o
  );
}
async function runDocumentExtraction(inspection, extractor, checked, n = {}) {
  if (checked !== true)
    throw Error(
      "Approve the host extractor before sharing the selected source.",
    );
  if (inspection.method !== "document-adapter")
    throw Error("Choose a document source for extraction.");
  const sourceFileDescriptor = toSourceFileDescriptor(inspection);
  if (n.signal?.aborted) throw Error("Document extraction cancelled.");
  const a = n.timeoutMs ?? 3e4;
  if (!Number.isInteger(a) || a < 1 || a > 6e4)
    throw Error("Invalid document extraction time budget.");
  const s = new AbortController(),
    d = () => s.abort();
  n.signal?.addEventListener("abort", d, { once: true });
  let u;
  try {
    return await new Promise((f, p) => {
      let m = false;
      const v = (b, y) => {
          m ||
            ((m = true),
            s.signal.removeEventListener("abort", g),
            s.abort(),
            b ? p(b) : f(y));
        },
        g = () => v(Error("Document extraction cancelled."));
      (s.signal.addEventListener("abort", g, { once: true }),
        (u = setTimeout(
          () => v(Error("Document extraction exceeded its time budget.")),
          a,
        )),
        Promise.resolve()
          .then(() =>
            extractor.extract({
              file: new File(
                [new Uint8Array(sourceFileDescriptor.bytes)],
                sourceFileDescriptor.fileName,
              ),
              sha256: inspection.sha256,
              signal: s.signal,
            }),
          )
          .then(
            (b) => {
              if (!m)
                try {
                  v(undefined, parseExtractionResult(inspection, b));
                } catch {
                  v(
                    Error(
                      "Extraction has invalid types, confidence, regions or source identity.",
                    ),
                  );
                }
            },
            () =>
              v(
                Error("Document extraction failed. The workbook is unchanged."),
              ),
          ));
    });
  } finally {
    (clearTimeout(u), n.signal?.removeEventListener("abort", d), s.abort());
  }
}
const importPreviewRegistry = new WeakMap();
function buildSourceReview(workbook, inspection, t) {
  const n = assertSourceInspectionCurrent(inspection);
  if (
    (validateExtractedSourceTables(inspection),
    assertWorkbookStructureEditable(workbook),
    typeof t.sheetName != "string" ||
      !t.sheetName.trim() ||
      t.sheetName.length > 31 ||
      /[\\/?*\[\]:\x00-\x1f]/.test(t.sheetName) ||
      workbook.sheets.some(
        (sheet) =>
          sheet.name.toLowerCase() === t.sheetName.trim().toLowerCase(),
      ))
  )
    throw Error("Choose a unique worksheet name of 1–31 characters.");
  const o = inspection.tables.find((table) => table.id === t.tableId);
  if (!o) throw Error("Select a source table.");
  if (
    (o.ambiguous || inspection.tables.length > 1) &&
    t.ambiguityReviewed !== true
  )
    throw Error("Review and confirm the ambiguous table selection.");
  if (!t.resolutions || typeof t.resolutions != "object")
    throw Error("Invalid source review.");
  const cells = {},
    s = new Set();
  let accepted = 0,
    excluded = 0,
    corrected = 0;
  if (
    (o.columns.forEach((column, m) => {
      cells[columnIndexToLetters(m) + "1"] = {
        raw: column.name,
        literal: true,
        style: { bold: true },
      };
    }),
    o.rows.forEach((row, m) =>
      row.forEach((v, g) => {
        const b = `${m}:${g}`;
        s.add(b);
        const y = t.resolutions[b],
          extractedCellValue = resolveExtractedCellValue(v, y);
        (y?.decision === "exclude"
          ? excluded++
          : y?.decision === "correct"
            ? corrected++
            : accepted++,
          (cells[columnIndexToLetters(g) + (m + 2)] = {
            raw: extractedCellValue,
            ...(typeof extractedCellValue == "string" ? { literal: true } : {}),
          }));
      }),
    ),
    Object.keys(t.resolutions).some((resolution) => !s.has(resolution)))
  )
    throw Error("A review decision refers to an unknown cell.");
  const c = {
      id: crypto.randomUUID(),
      name: t.sheetName.trim(),
      cells: cells,
      rowCount: workbookLimits.maxRows,
      colCount: workbookLimits.maxCols,
    },
    record = {
      id: crypto.randomUUID(),
      sheetId: c.id,
      sheetName: c.name,
      importedAt: new Date().toISOString(),
      original: buildRetainedSourceRecord(
        n,
        inspection.fileName,
        inspection.format,
      ),
      method: inspection.method,
      extractor: inspection.extractor,
      ...(inspection.pageCount ? { pageCount: inspection.pageCount } : {}),
      table: structuredClone(o),
      resolutions: structuredClone(t.resolutions),
      ambiguityReviewed: t.ambiguityReviewed,
    };
  normalizeWorkbookInput({
    ...workbook,
    sheets: [...workbook.sheets, c],
    sourceImports: [...(workbook.sourceImports ?? []), record],
  });
  const f = {
    workbookHash: fingerprintCanonicalValue(workbook),
    inspection: inspection,
    review: structuredClone(t),
    record: record,
    sheet: c,
    accepted: accepted,
    excluded: excluded,
    corrected: corrected,
  };
  return (importPreviewRegistry.set(f, fingerprintCanonicalValue(f)), f);
}
function commitSourceImport(workbook, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (
    (assertWorkbookStructureEditable(workbook),
    importPreviewRegistry.get(r) !== fingerprintCanonicalValue(r) ||
      r.workbookHash !== fingerprintCanonicalValue(workbook))
  )
    throw Error("The workbook or review changed. Preview the import again.");
  assertSourceInspectionCurrent(r.inspection);
  const workbookInput = normalizeWorkbookInput({
    ...workbook,
    sheets: [...workbook.sheets, r.sheet],
    sourceImports: [...(workbook.sourceImports ?? []), r.record],
  });
  return (importPreviewRegistry.delete(r), workbookInput);
}
function findSourceImport(workbook, r) {
  const t = workbook.sourceImports?.find(
    (sourceImport) => sourceImport.id === r,
  );
  if (!t) throw Error("Source import is unavailable.");
  return {
    bytes: validateRetainedSourceFile(t.original),
    fileName: t.original.fileName,
    sha256: t.original.sha256,
  };
}
function resolveImportedSourceCell(workbook, r, t, n) {
  const o = workbook.sourceImports?.find(
      (sourceImport) => sourceImport.id === r,
    ),
    source = o?.table.rows[t]?.[n];
  if (!o || !source) throw Error("Imported source cell is unavailable.");
  const address = columnIndexToLetters(n) + (t + 2),
    d = workbook.sheets.find((sheet) => sheet.id === o.sheetId),
    cell = d?.cells[address],
    l = o.resolutions[`${t}:${n}`],
    extractedCellValue = resolveExtractedCellValue(source, l);
  return {
    address: address,
    sheetId: o.sheetId,
    sourceHash: o.original.sha256,
    source: source,
    decision: l?.decision ?? "accepted",
    importedValue: extractedCellValue,
    currentValue: cell?.raw ?? null,
    status: d
      ? cell?.raw === extractedCellValue &&
        (typeof extractedCellValue != "string" || cell.literal === true)
        ? "matches-import"
        : "changed"
      : "missing-sheet",
  };
}
const downloadBlobAsFile = ({ filename, blob }) => {
  const t = URL.createObjectURL(blob),
    n = document.createElement("a");
  ((n.href = t),
    (n.download = filename),
    n.click(),
    setTimeout(() => URL.revokeObjectURL(t), 1e3));
};

export {
  commitSourceImport,
  useOperationScope,
  toSourceFileDescriptor,
  runDocumentExtraction,
  resolveImportedSourceCell,
  WorkbookLinksDialog,
  findSourceImport,
  Rue,
  buildSourceImportRecord,
  WorkbookLinkStatusBadge,
  operationScopeContext,
  buildSourceReview,
  downloadBlobAsFile,
  parseExtractionResult,
  useCancellableOperation,
};
