import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { string, array, discriminatedUnion, literal, object } from "zod";
import {
  Gf,
  Iw,
  Pg,
  eB,
  lf,
  lx,
  ts,
  yE,
  yc,
} from "./01-core-schemas-and-hashing.js";
import { HJ, Mt, dx } from "./02-financial-functions.js";
import { Cf, Fc, Gw, Vie, Vn, Wie, ct, sU, ut } from "./03-workbook-model.js";
import { Dr, M, ci } from "./04-cell-input-paste-and-fill.js";
import { Ek } from "./08-xlsx-export-and-preservation.js";

function Zce(e) {
  const { text: r, language: t } = Dr(),
    n = ne.useId(),
    o = ne.useRef(null),
    [, a] = ne.useState(0),
    { report: s } = Gw(e.workbook),
    [d, u] = ne.useState(s.settings),
    [l, c] = ne.useState(""),
    [h, f] = ne.useState(""),
    [p, m] = ne.useState(0),
    [v, g] = ne.useState(() => ts(e.workbook.datePolicy));
  (ne.useEffect(() => {
    (u(Gw(e.workbook).report.settings), g(ts(e.workbook.datePolicy)), m(0));
  }, [e.workbook]),
    ne.useEffect(() => {
      var V;
      const j = document.activeElement;
      return (
        (V = o.current?.querySelector(
          '[data-ui-label="Close Calculation Options"]',
        )) == null || V.focus(),
        () => {
          j != null && j.isConnected && j.focus();
        }
      );
    }, []));
  const b = (j) => {
      try {
        (j(), c(""));
      } catch (z) {
        c(z instanceof Error ? z.message : "Calculation action failed.");
      }
    },
    y = () =>
      b(() => {
        var j;
        if (e.onCalculate) {
          e.onCalculate();
          return;
        }
        (sU(e.workbook),
          a((z) => z + 1),
          (j = e.onCalculated) == null || j.call(e));
      }),
    w = () =>
      b(() => {
        if (e.readOnly)
          throw Error(
            "This host is read-only. Calculation settings cannot be changed.",
          );
        e.onWorkbookChange(
          { ...Vie(e.workbook, d), datePolicy: ts(v) },
          `Set calculation to ${d.mode}; iteration ${d.iterative ? "enabled" : "disabled"}; dates ${v.dateSystem} ${v.textDateOrder}, time zone ${v.timeZone}`,
        );
      }),
    E = () =>
      b(() => {
        const j = Wie(e.workbook, s);
        if (e.onExport) {
          e.onExport(j);
          return;
        }
        const z = URL.createObjectURL(
            new Blob([j], { type: "application/json" }),
          ),
          V = document.createElement("a");
        ((V.href = z),
          (V.download = "Calculation-evidence.json"),
          V.click(),
          setTimeout(() => URL.revokeObjectURL(z), 1e3));
      }),
    A = s.cells.filter((j) => {
      return `${e.workbook.sheets.find((V) => V.id === j.sheetId)?.name}!${j.address} ${j.formula} ${j.status} ${j.result.value}`
        .toLowerCase()
        .includes(h.toLowerCase());
    }),
    k = A.slice(p * 100, (p + 1) * 100),
    F = (j) => {
      var V, D;
      if (
        (j.key === "Escape" && (j.stopPropagation(), e.onClose()),
        j.key !== "Tab")
      )
        return;
      const z = [
        ...o.current.querySelectorAll(
          'button:not(:disabled),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
        ),
      ];
      j.shiftKey && document.activeElement === z[0]
        ? (j.preventDefault(), (V = z.at(-1)) == null || V.focus())
        : !j.shiftKey &&
          document.activeElement === z.at(-1) &&
          (j.preventDefault(), (D = z[0]) == null || D.focus());
    };
  return i.jsx("div", {
    className: "fw-backdrop",
    children: i.jsxs("section", {
      ref: o,
      className: "fw-dialog rc-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: F,
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
                  children: i.jsx(M, { children: "Calculation Options" }),
                }),
                i.jsx("p", {
                  children: i.jsx(M, {
                    children:
                      "Review current results, dependency changes and separately retained imported caches.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": r("Close Calculation Options"),
              onClick: e.onClose,
              lang: t("Close Calculation Options"),
              "data-ui-label": "Close Calculation Options",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("fieldset", {
          disabled: e.readOnly,
          className: "rc-settings",
          children: [
            i.jsx("legend", {
              children: i.jsx(M, { children: "Workbook settings" }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Calculation mode" }),
                i.jsxs("select", {
                  "aria-label": r("Calculation mode"),
                  value: d.mode,
                  onChange: (j) => u({ ...d, mode: j.target.value }),
                  lang: t("Calculation mode"),
                  "data-ui-label": "Calculation mode",
                  children: [
                    i.jsx("option", {
                      value: "automatic",
                      children: "Automatic",
                    }),
                    i.jsx("option", { value: "manual", children: "Manual" }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: d.iterative,
                  onChange: (j) => u({ ...d, iterative: j.target.checked }),
                }),
                i.jsx(M, { children: "Enable iterative calculation" }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Maximum iterations" }),
                i.jsx(ci, {
                  "aria-label": r("Maximum iterations"),
                  min: 1,
                  max: 1e3,
                  value: d.maxIterations,
                  onValueChange: (j) => u({ ...d, maxIterations: Number(j) }),
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Maximum change" }),
                i.jsx(ci, {
                  "aria-label": r("Maximum change"),
                  min: 0,
                  max: 1,
                  step: "0.000001",
                  value: d.maxChange,
                  onValueChange: (j) => u({ ...d, maxChange: Number(j) }),
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Time budget (ms)" }),
                i.jsx(ci, {
                  "aria-label": r("Time budget (ms)"),
                  min: 1,
                  max: 1e4,
                  value: d.timeBudgetMs,
                  onValueChange: (j) => u({ ...d, timeBudgetMs: Number(j) }),
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Working-data budget (MiB)" }),
                i.jsx(ci, {
                  "aria-label": r("Working-data budget (MiB)"),
                  min: 1,
                  max: 128,
                  value: d.memoryBudgetMiB ?? Iw.defaultMemoryBudgetMiB,
                  onValueChange: (j) => u({ ...d, memoryBudgetMiB: Number(j) }),
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Calculation work budget" }),
                i.jsx(ci, {
                  "aria-label": r("Calculation work budget"),
                  min: 1e3,
                  max: 1e7,
                  step: 1e3,
                  value: d.workBudget ?? Iw.defaultWorkBudget,
                  onValueChange: (j) => u({ ...d, workBudget: Number(j) }),
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Workbook date system" }),
                i.jsxs("select", {
                  "aria-label": r("Workbook date system"),
                  value: v.dateSystem,
                  onChange: (j) => g({ ...v, dateSystem: j.target.value }),
                  lang: t("Workbook date system"),
                  "data-ui-label": "Workbook date system",
                  children: [
                    i.jsx("option", {
                      value: "1900",
                      children: "1900 (Windows default)",
                    }),
                    i.jsx("option", {
                      value: "1904",
                      children: "1904 (legacy Mac)",
                    }),
                  ],
                }),
              ],
            }),
            v.dateSystem !== s.dateContext.dateSystem &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(M, {
                  children:
                    "Changing the date system keeps stored numbers and formulas. Existing date serials will be interpreted using the new system; ordinary corresponding dates differ by 1,462 days. The 1900 system includes the fictitious serial 60 as a special boundary. Review this change before applying. Undo restores the previous setting.",
                }),
              }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Text date order" }),
                i.jsxs("select", {
                  "aria-label": r("Text date order"),
                  value: v.textDateOrder,
                  onChange: (j) => g({ ...v, textDateOrder: j.target.value }),
                  lang: t("Text date order"),
                  "data-ui-label": "Text date order",
                  children: [
                    i.jsx("option", {
                      value: "MDY",
                      children: "Month/day/year",
                    }),
                    i.jsx("option", {
                      value: "DMY",
                      children: "Day/month/year",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Calculation time zone" }),
                i.jsx("input", {
                  "aria-label": r("Calculation time zone"),
                  value: v.timeZone,
                  onChange: (j) => g({ ...v, timeZone: j.target.value }),
                  placeholder: r("UTC or America/New_York"),
                  lang: t("Calculation time zone"),
                  "data-ui-label": "Calculation time zone",
                }),
              ],
            }),
            i.jsx("button", {
              onClick: w,
              children: i.jsx(M, { children: "Apply calculation settings" }),
            }),
          ],
        }),
        e.readOnly &&
          i.jsx("p", {
            children: i.jsx(M, {
              children:
                "Read-only host: settings are locked. Calculate Now refreshes local results without editing source cells or history.",
            }),
          }),
        i.jsxs("div", {
          className: "fw-toolbar",
          children: [
            i.jsx("button", {
              onClick: y,
              children: i.jsx(M, { children: "Calculate Now" }),
            }),
            s.execution?.mode === "worker" &&
              s.execution.state === "pending" &&
              e.onCancelCalculation &&
              i.jsx("button", {
                onClick: e.onCancelCalculation,
                children: i.jsx(M, { children: "Cancel calculation" }),
              }),
            i.jsx("button", {
              onClick: E,
              children: i.jsx(M, { children: "Download calculation evidence" }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Find formula" }),
                i.jsx("input", {
                  "aria-label": r("Find formula results"),
                  value: h,
                  onChange: (j) => {
                    (f(j.target.value), m(0));
                  },
                  lang: t("Find formula results"),
                  "data-ui-label": "Find formula results",
                }),
              ],
            }),
          ],
        }),
        l && i.jsx("p", { role: "alert", children: l }),
        i.jsxs("p", {
          role: "status",
          children: [
            i.jsx(M, { children: "Revision " }),
            s.revision,
            " · ",
            s.settings.mode,
            " · ",
            s.formulaCount,
            " ",
            i.jsx(M, { children: "formulas · " }),
            s.evaluated,
            " ",
            i.jsx(M, { children: "evaluated ·" }),
            " ",
            s.reused,
            " ",
            i.jsx(M, { children: "reused · " }),
            s.pending,
            " ",
            i.jsx(M, { children: "pending ·" }),
            " ",
            s.dependencyEdges,
            " ",
            i.jsx(M, { children: "dependency edges" }),
          ],
        }),
        i.jsxs("p", { children: [s.reason, ". ", s.iteration.message] }),
        i.jsxs("details", {
          className: "rc-date-evidence",
          open: true,
          children: [
            i.jsx("summary", {
              children: i.jsx(M, {
                children: "Calculation limits and volatile formulas",
              }),
            }),
            i.jsxs("p", {
              role: "status",
              children: [
                s.execution?.message,
                " ",
                s.limits.hardInterrupt
                  ? "Worker deadline and cancellation are enabled."
                  : "Cooperative execution: this host does not interrupt an individual running function.",
              ],
            }),
            i.jsxs("p", {
              children: [
                i.jsx(M, { children: "Working data:" }),
                " ",
                (s.resources.accountedBytes / 1048576).toFixed(2),
                " /",
                " ",
                (s.resources.memoryBudgetBytes / 1048576).toFixed(0),
                " ",
                i.jsx(M, { children: "MiB accounted. Work: " }),
                s.resources.workUnits.toLocaleString(),
                " /",
                " ",
                s.resources.workBudget.toLocaleString(),
                " ",
                i.jsx(M, { children: "units. Time budget:" }),
                " ",
                s.settings.timeBudgetMs,
                " ",
                i.jsx(M, { children: "ms." }),
              ],
            }),
            s.resources.reason &&
              i.jsxs("p", {
                role: "alert",
                children: [
                  s.resources.reason,
                  i.jsx(M, {
                    children:
                      ". Committed inputs remain intact. Adjust the data or budget and choose Calculate Now to retry.",
                  }),
                ],
              }),
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "The memory setting bounds accounted input, intermediate and result data. It does not measure the browser's total memory; parser and runtime overhead are separate. Worker deadlines depend on the browser delivering timer events.",
              }),
            }),
            i.jsx("p", { children: yE.refresh }),
            i.jsx("p", { children: yE.random }),
            i.jsxs("p", {
              children: [
                i.jsx(M, { children: "Volatile functions in this workbook:" }),
                " ",
                s.volatile.functions.join(", ") || "None",
                i.jsx(M, {
                  children:
                    ". Unsupported volatile functions require a compatibility review.",
                }),
              ],
            }),
          ],
        }),
        i.jsxs("details", {
          className: "rc-date-evidence",
          open: true,
          children: [
            i.jsx("summary", {
              children: i.jsx(M, {
                children: "Precision and monetary rounding",
              }),
            }),
            i.jsxs("p", {
              children: [
                i.jsx(M, { children: "Numbers use " }),
                Ek.storage,
                ". ",
                Ek.display,
                " ",
                Ek.rounding,
              ],
            }),
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Binary arithmetic can produce 0.30000000000000004 for 0.1 + 0.2. General displays 0.3; ROUND(0.1+0.2,2) explicitly produces 0.3. Store identifiers longer than 15 digits as Text. Native JSON retains stored numbers; XLSX may expose Excel precision limits. SUM uses decimal accumulation of stored numbers, so it can differ from a chain of binary additions.",
              }),
            }),
          ],
        }),
        i.jsxs("details", {
          className: "rc-date-evidence",
          open: true,
          children: [
            i.jsx("summary", {
              children: i.jsx(M, {
                children: "Date and business-calendar policy",
              }),
            }),
            i.jsxs("p", {
              children: [
                i.jsx(M, { children: "Gregorian calendar · " }),
                s.dateContext.dateSystem,
                " ",
                i.jsx(M, {
                  children:
                    "serial system · ISO YYYY-MM-DD within the supported date range accepted · slash dates use ",
                }),
                s.dateContext.textDateOrder,
                i.jsx(M, {
                  children:
                    ". Other calendars are unsupported. XLSX retains the selected serial system; native JSON also retains text order and time zone. Holiday ranges and weekend masks are supplied in WORKDAY / NETWORKDAYS formulas.",
                }),
              ],
            }),
            i.jsxs("p", {
              children: [
                i.jsx(M, { children: "Calculation snapshot:" }),
                " ",
                i.jsx("time", {
                  dateTime: s.dateContext.calculatedAt,
                  children: s.dateContext.calculatedAt,
                }),
                " ",
                i.jsx(M, { children: "· time zone " }),
                s.dateContext.timeZone,
                " ",
                i.jsx(M, { children: "· UTC offset" }),
                " ",
                s.dateContext.utcOffsetMinutes,
                " ",
                i.jsx(M, { children: "minutes · local clock" }),
                " ",
                s.dateContext.localDateTime,
                ".",
              ],
            }),
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "TODAY and NOW share this captured instant. Calculate Now refreshes it; there is no background clock. Pending formulas have no current time result. Calendar policy persists in native JSON; XLSX does not retain this time zone or text-date policy.",
              }),
            }),
          ],
        }),
        !!s.cycles.length &&
          i.jsxs("p", {
            role: "alert",
            children: [
              i.jsx(M, { children: "Circular references:" }),
              " ",
              s.cycles
                .slice(0, 10)
                .map((j) =>
                  j
                    .map((z) => {
                      const V = z.lastIndexOf(":");
                      return `${e.workbook.sheets.find((B) => B.id === z.slice(0, V))?.name ?? "(removed sheet)"}!${z.slice(V + 1)}`;
                    })
                    .join(" → "),
                )
                .join("; "),
              s.cycles.length > 10 ? " (more in evidence)" : "",
            ],
          }),
        !!s.pending &&
          i.jsx("p", {
            children: i.jsx(M, {
              children:
                "Pending formulas display #PENDING!. Historical and imported values are never substituted as current results. Use Calculate Now before relying on or sharing them.",
            }),
          }),
        i.jsx("div", {
          className: "fw-table-wrap",
          children: i.jsxs("table", {
            children: [
              i.jsx("thead", {
                children: i.jsxs("tr", {
                  children: [
                    i.jsx("th", { children: i.jsx(M, { children: "Source" }) }),
                    i.jsx("th", {
                      children: i.jsx(M, { children: "Formula" }),
                    }),
                    i.jsx("th", {
                      children: i.jsx(M, { children: "Current result" }),
                    }),
                    i.jsx("th", {
                      children: i.jsx(M, { children: "Status / evaluations" }),
                    }),
                    i.jsx("th", {
                      children: i.jsx(M, { children: "Historical result" }),
                    }),
                  ],
                }),
              }),
              i.jsx("tbody", {
                children: k.map((j) => {
                  return i.jsxs(
                    "tr",
                    {
                      children: [
                        i.jsx("td", {
                          children: i.jsxs("button", {
                            onClick: () => {
                              var D;
                              ((D = e.onNavigate) == null ||
                                D.call(e, j.sheetId, j.address),
                                e.onClose());
                            },
                            children: [
                              e.workbook.sheets.find((D) => D.id === j.sheetId)
                                ?.name,
                              "!",
                              j.address,
                            ],
                          }),
                        }),
                        i.jsx("td", { children: j.formula }),
                        i.jsxs("td", {
                          children: [
                            String(j.result.value ?? "(blank)"),
                            " (",
                            j.result.type,
                            ")",
                          ],
                        }),
                        i.jsxs("td", {
                          children: [
                            j.status,
                            " / ",
                            j.evaluations,
                            j.financial?.map((D, B) =>
                              i.jsxs(
                                "details",
                                {
                                  children: [
                                    i.jsxs("summary", {
                                      children: [
                                        D.functionName,
                                        " ",
                                        i.jsx(M, {
                                          children: "financial details — ",
                                        }),
                                        D.status,
                                      ],
                                    }),
                                    i.jsx("p", { children: D.message }),
                                    i.jsx("p", { children: D.convention }),
                                    D.solver &&
                                      i.jsxs(i.Fragment, {
                                        children: [
                                          i.jsxs("p", {
                                            children: [
                                              i.jsx(M, {
                                                children: "Guess ",
                                              }),
                                              D.guess,
                                              "; ",
                                              D.solver.evaluations,
                                              " ",
                                              i.jsx(M, {
                                                children:
                                                  "equation evaluations. Bounded search: ",
                                              }),
                                              D.solver.status,
                                              i.jsx(M, {
                                                children: ". Candidate roots:",
                                              }),
                                              " ",
                                              D.solver.candidates
                                                .map((_) =>
                                                  _.rate.toPrecision(12),
                                                )
                                                .join(", ") || "None",
                                              ".",
                                            ],
                                          }),
                                          D.solver.selected &&
                                            i.jsxs("p", {
                                              children: [
                                                i.jsx(M, {
                                                  children: "Selected rate",
                                                }),
                                                " ",
                                                D.solver.selected.rate.toPrecision(
                                                  12,
                                                ),
                                                i.jsx(M, {
                                                  children:
                                                    "; equation residual",
                                                }),
                                                " ",
                                                D.solver.selected.residual.toPrecision(
                                                  6,
                                                ),
                                                i.jsx(M, {
                                                  children:
                                                    "; relative residual",
                                                }),
                                                " ",
                                                D.solver.selected.relativeResidual.toExponential(
                                                  3,
                                                ),
                                                ".",
                                              ],
                                            }),
                                          i.jsx("p", {
                                            children: i.jsx(M, {
                                              children:
                                                "Root discovery is not exhaustive. Repeated calls of the same function in this cell retain the last diagnostic; use Evaluate Formula for the expression trace.",
                                            }),
                                          }),
                                        ],
                                      }),
                                  ],
                                },
                                `${D.functionName}-${B}`,
                              ),
                            ),
                          ],
                        }),
                        i.jsx("td", {
                          children: j.lastKnown
                            ? `${String(j.lastKnown.result.value)} at revision ${j.lastKnown.revision} — historical only`
                            : "None retained",
                        }),
                      ],
                    },
                    `${j.sheetId}:${j.address}`,
                  );
                }),
              }),
            ],
          }),
        }),
        i.jsxs("div", {
          className: "fw-toolbar",
          children: [
            i.jsx("button", {
              disabled: p === 0,
              onClick: () => m((j) => j - 1),
              children: i.jsx(M, { children: "Previous results" }),
            }),
            i.jsxs("span", {
              children: [
                A.length ? p * 100 + 1 : 0,
                "–",
                Math.min((p + 1) * 100, A.length),
                " ",
                i.jsx(M, { children: "of " }),
                A.length,
              ],
            }),
            i.jsx("button", {
              disabled: (p + 1) * 100 >= A.length,
              onClick: () => m((j) => j + 1),
              children: i.jsx(M, { children: "Next results" }),
            }),
          ],
        }),
        e.workbook.importedCalculationCaches &&
          i.jsxs("details", {
            children: [
              i.jsxs("summary", {
                children: [
                  i.jsx(M, { children: "Imported cache evidence —" }),
                  " ",
                  e.workbook.importedCalculationCaches.entries.length,
                  " ",
                  i.jsx(M, { children: "source formulas" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  e.workbook.importedCalculationCaches.fileName,
                  " ",
                  i.jsx(M, { children: "· SHA-256" }),
                  " ",
                  e.workbook.importedCalculationCaches.sourceSha256,
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(M, { children: "Imported " }),
                  e.workbook.importedCalculationCaches.importedAt,
                  i.jsx(M, { children: "; source mode " }),
                  e.workbook.importedCalculationCaches.sourceSettings.mode,
                  i.jsx(M, {
                    children:
                      ". Source caches and settings remain historical, unverified evidence. Download calculation evidence for every typed cache comparison, including missing caches and changed formulas. Original XLSX bytes are not retained by this field.",
                  }),
                ],
              }),
            ],
          }),
        i.jsx("p", {
          children: i.jsx(M, {
            children:
              "Scalar dependency changes reuse unaffected results. Names, structure and arrays trigger a conservative recalculation. Iteration starts from zero and requires numeric convergence; it does not promise Excel-identical iteration order. Worker execution can be terminated; synchronous SDK and diagnostic calculations use cooperative checks. Formula inspection is a separate diagnostic calculation.",
          }),
        }),
      ],
    }),
  });
}
function Kce({
  x: e,
  y: r,
  label: t,
  actions: n,
  workbook: o,
  selection: a,
  onClose: s,
}) {
  const d = ne.useRef(null),
    u = ne.useRef({ workbook: o, snapshot: JSON.stringify(o), selection: a }),
    [l, c] = ne.useState({ left: e, top: r }),
    f =
      u.current.workbook !== o ||
      u.current.snapshot !== JSON.stringify(o) ||
      u.current.selection !== a
        ? "The workbook or selection changed. Close and reopen this menu."
        : undefined;
  return (
    ne.useLayoutEffect(() => {
      const p = d.current?.getBoundingClientRect();
      p &&
        c({
          left: Math.max(8, Math.min(e, window.innerWidth - p.width - 8)),
          top: Math.max(8, Math.min(r, window.innerHeight - p.height - 8)),
        });
    }, [e, r]),
    ne.useEffect(() => {
      let p = false;
      const m = requestAnimationFrame(() => {
          return d.current?.querySelector("button")?.focus();
        }),
        v = (g) => {
          var b;
          (b = d.current) != null && b.contains(g.target)
            ? (p = true)
            : (g.type !== "focusin" || p) && s(false);
        };
      return (
        document.addEventListener("pointerdown", v),
        document.addEventListener("focusin", v),
        () => {
          (cancelAnimationFrame(m),
            document.removeEventListener("pointerdown", v),
            document.removeEventListener("focusin", v));
        }
      );
    }, [s]),
    i.jsxs("div", {
      ref: d,
      className: "lg-grid-context",
      role: "menu",
      "aria-label": `${t} commands`,
      style: l,
      onContextMenu: (p) => p.preventDefault(),
      onKeyDown: (p) => {
        var g;
        if (p.defaultPrevented || p.ctrlKey || p.metaKey || p.altKey) return;
        const m = Array.from(
            p.currentTarget.querySelectorAll('[role="menuitem"]'),
          ),
          v = m.indexOf(document.activeElement);
        if (["ArrowDown", "ArrowUp", "Home", "End"].includes(p.key)) {
          (p.preventDefault(), p.stopPropagation());
          const b =
            p.key === "Home"
              ? 0
              : p.key === "End"
                ? m.length - 1
                : (v + (p.key === "ArrowDown" ? 1 : m.length - 1)) % m.length;
          (g = m[b]) == null || g.focus();
        } else
          (p.key === "Escape" || p.key === "Tab") &&
            (p.preventDefault(), p.stopPropagation(), s(true));
      },
      children: [
        i.jsx("strong", { children: t }),
        n.map((p, m) => {
          const v = f ?? p.reason;
          return i.jsxs(
            "button",
            {
              type: "button",
              role: "menuitem",
              tabIndex: m === 0 ? 0 : -1,
              "aria-label": p.name,
              "aria-disabled": !!v,
              "aria-description": v,
              title: v ? `${p.name}: ${v}` : p.name,
              onClick: () => {
                v || (s(true), p.run());
              },
              children: [p.name, v && i.jsx("small", { children: v })],
            },
            p.name,
          );
        }),
      ],
    })
  );
}
function Yce() {
  const { text: e, language: r } = Dr();
  return i.jsxs("section", {
    "aria-label": e("Workbook capacity limits"),
    lang: r("Workbook capacity limits"),
    "data-ui-label": "Workbook capacity limits",
    children: [
      i.jsx("h4", { children: i.jsx(M, { children: "Capacity" }) }),
      i.jsx("p", {
        className: "lg-muted",
        children: i.jsx(M, {
          children:
            "The worksheet address space matches Excel. Processing and file limits below are this application's limits.",
        }),
      }),
      i.jsx("dl", {
        className: "lg-capacity-list",
        children: HJ.map(([t, n]) =>
          i.jsxs(
            "div",
            {
              children: [
                i.jsx("dt", { children: t }),
                i.jsx("dd", { children: n }),
              ],
            },
            t,
          ),
        ),
      }),
      i.jsx("p", {
        className: "lg-muted",
        children: i.jsx(M, {
          children:
            "Text units use UTF-16: some symbols, including many emoji, count as two. File limits count bytes. Excel allows 8,192 formula characters; this editor allows 4,096 text units.",
        }),
      }),
      i.jsx("p", {
        className: "lg-muted",
        children: i.jsx(M, {
          children:
            "Imports that exceed a limit are refused before replacing your workbook. Empty rows between distant cells are not filled in. Available memory and browser storage quota can impose lower limits.",
        }),
      }),
      i.jsx("p", {
        className: "lg-muted",
        children: i.jsx(M, {
          children:
            "Uploaded files are parsed in isolated workers with a 30-second default deadline and at most two simultaneous requests. Cancel file processing stops a pending import. A blocked worker never falls back to parsing in the editor. Error messages omit source details; keep the original file for repairs. Hosts can provide a trusted worker and a deadline up to 60 seconds.",
        }),
      }),
    ],
  });
}
const Zm = Object.freeze({
    filesPerMessage: 4,
    fileBytes: 5 * 1024 * 1024,
    totalFileBytes: 10 * 1024 * 1024,
    historyMessages: 20,
    historyCharacters: 5e4,
    promptCharacters: 1e4,
  }),
  Qw = {
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
  },
  cEe = Object.keys(Qw)
    .map((e) => `.${e}`)
    .join(","),
  SI = new Set(["txt", "md", "csv", "tsv", "json"]),
  jy = (e) => {
    return e.split(".").pop()?.toLowerCase() ?? "";
  };
function js(e, r) {
  if (!e) throw new Error(r);
}
function wz(e) {
  return (
    typeof e == "string" &&
    e.length > 0 &&
    e.length <= 180 &&
    !/[\x00-\x1f\x7f/\\]/.test(e)
  );
}
function xz(e, r) {
  const t = (...n) => n.every((o, a) => e[a] === o);
  return r === "pdf"
    ? t(37, 80, 68, 70, 45)
    : r === "png"
      ? t(137, 80, 78, 71, 13, 10, 26, 10)
      : r === "jpg" || r === "jpeg"
        ? t(255, 216, 255)
        : r === "webp"
          ? t(82, 73, 70, 70) &&
            String.fromCharCode(...e.slice(8, 12)) === "WEBP"
          : r === "xlsx" || r === "docx"
            ? t(80, 75, 3, 4)
            : SI.has(r);
}
function Jce(e, r) {
  return new Promise((t, n) => {
    const o = new FileReader();
    let a = false;
    const s = (l, c) => {
        a ||
          ((a = true),
          clearTimeout(u),
          r == null || r.removeEventListener("abort", d),
          c ? n(c) : t(new Uint8Array(l)));
      },
      d = () => {
        (s(
          undefined,
          new DOMException("File preparation cancelled.", "AbortError"),
        ),
          o.abort());
      },
      u = setTimeout(() => {
        (s(
          undefined,
          new Error("Reading the attachment timed out. Try a smaller file."),
        ),
          o.abort());
      }, 15e3);
    ((o.onload = () => s(o.result)),
      (o.onerror = () =>
        s(
          undefined,
          new Error("The attachment could not be read. Select it again."),
        )),
      (o.onabort = () =>
        s(
          undefined,
          new DOMException("File preparation cancelled.", "AbortError"),
        )),
      r == null || r.addEventListener("abort", d, { once: true }),
      r != null && r.aborted ? d() : o.readAsArrayBuffer(e));
  });
}
async function dEe(e, r) {
  js(
    wz(e.name),
    "Use a filename of 1–180 characters without path or control characters.",
  );
  const t = jy(e.name);
  (js(
    Object.hasOwn(Qw, t),
    "Supported attachments: TXT, Markdown, CSV, TSV, JSON, PDF, DOCX, XLSX, PNG, JPEG and WebP.",
  ),
    js(
      e.size > 0 && e.size <= Zm.fileBytes,
      "Each attachment must be nonempty and no larger than 5 MB.",
    ));
  const n = await Jce(e, r);
  js(
    n.length === e.size && xz(n, t),
    "The file content does not match its extension. Choose the original supported file.",
  );
  let o;
  const a = SI.has(t) ? "utf8" : "base64";
  if (a === "utf8") {
    try {
      o = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(n);
    } catch {
      throw new Error(
        "Text attachments must use UTF-8 encoding. Save the file as UTF-8 and try again.",
      );
    }
    js(
      !o.includes("\0"),
      "This appears to be a binary file. Choose a supported document or UTF-8 text file.",
    );
  } else {
    let s = "";
    for (let d = 0; d < n.length; d += 8192)
      s += String.fromCharCode(...n.subarray(d, d + 8192));
    o = btoa(s);
  }
  return {
    id: crypto.randomUUID(),
    name: e.name,
    mediaType: Qw[t],
    size: n.length,
    encoding: a,
    data: o,
  };
}
const Lk = (e, r) =>
  !!e &&
  typeof e == "object" &&
  !Array.isArray(e) &&
  Object.keys(e).every((t) => r.includes(t));
function Ez(e) {
  (js(
    Lk(e, ["version", "messages", "attachments"]) &&
      e.version === 1 &&
      Array.isArray(e.messages) &&
      Array.isArray(e.attachments),
    "Invalid chat request.",
  ),
    js(
      e.messages.length <= Zm.historyMessages,
      "This conversation has reached 20 prior messages. Start a new chat to continue.",
    ));
  let r = 0,
    t = 0;
  const n = (a) => {
    js(
      Array.isArray(a) && a.length <= Zm.filesPerMessage,
      "Attach up to 4 files per message.",
    );
    const s = new Set();
    return a.map((d) => {
      (js(
        Lk(d, ["id", "name", "mediaType", "size", "encoding", "data"]) &&
          typeof d.id == "string" &&
          d.id.length > 0 &&
          d.id.length <= 100 &&
          !s.has(d.id),
        "Invalid or duplicate attachment.",
      ),
        s.add(d.id),
        js(
          wz(d.name) &&
            Object.hasOwn(Qw, jy(d.name)) &&
            d.mediaType === Qw[jy(d.name)],
          "Unsupported attachment type.",
        ),
        js(
          Number.isSafeInteger(d.size) &&
            Number(d.size) > 0 &&
            Number(d.size) <= Zm.fileBytes &&
            typeof d.data == "string" &&
            d.data.length <= Math.ceil(Zm.fileBytes / 3) * 4,
          "Attachment exceeds the 5 MB limit.",
        ));
      const u = SI.has(jy(d.name));
      js(
        d.encoding === (u ? "utf8" : "base64"),
        "Invalid attachment encoding.",
      );
      let l;
      return (
        u
          ? (js(!d.data.includes("\0"), "Invalid text attachment."),
            (l = new TextEncoder().encode(d.data)))
          : (js(
              d.data.length % 4 === 0 && /^[A-Za-z0-9+/]*={0,2}$/.test(d.data),
              "Invalid attachment encoding.",
            ),
            (l = Uint8Array.from(atob(d.data), (c) => c.charCodeAt(0)))),
        js(
          l.length === d.size && xz(l, jy(d.name)),
          "Attachment content or size is invalid.",
        ),
        (r += Number(d.size)),
        js(
          r <= Zm.totalFileBytes,
          "The conversation contains more than 10 MB of attachments. Remove files or start a new chat.",
        ),
        {
          id: d.id,
          name: d.name,
          mediaType: d.mediaType,
          size: Number(d.size),
          encoding: d.encoding,
          data: d.data,
        }
      );
    });
  };
  return {
    version: 1,
    messages: e.messages.map(
      (a) => (
        js(
          Lk(a, ["role", "text", "attachments"]) &&
            (a.role === "user" || a.role === "assistant") &&
            typeof a.text == "string" &&
            a.text.length <= 1e4,
          "Invalid chat history message.",
        ),
        js(
          a.role === "user" || a.attachments === undefined,
          "Assistant messages cannot attach files.",
        ),
        (t += a.text.length),
        js(
          t <= Zm.historyCharacters,
          "This conversation has reached its context limit. Start a new chat to continue.",
        ),
        {
          role: a.role,
          text: a.text,
          ...(a.attachments ? { attachments: n(a.attachments) } : {}),
        }
      ),
    ),
    attachments: n(e.attachments),
  };
}
const L7 = {
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
class cf extends Error {
  constructor(t) {
    const n = typeof t == "string" && Object.hasOwn(L7, t) ? t : "unavailable";
    super(L7[n]);
    yc(this, "code");
    ((this.name = "LLMRequestError"), (this.code = n));
  }
}
function Qce(e) {
  return e instanceof cf
    ? new cf(e.code)
    : e instanceof Error && e.name === "TimeoutError"
      ? new cf("timeout")
      : e instanceof Error && e.name === "AbortError"
        ? new cf("cancelled")
        : new cf("unavailable");
}
async function ede(e, r, t = 3e4) {
  if (!Number.isFinite(t) || t < 1 || t > 12e4) throw new cf("invalid_request");
  if (r != null && r.aborted) throw new cf("cancelled");
  const n = new AbortController();
  let o;
  const a = new Promise((l, c) => {
      o = c;
    }),
    s = (l) => {
      const c = new cf(l);
      (o(c), n.abort(c));
    },
    d = () => s("cancelled");
  r == null || r.addEventListener("abort", d, { once: true });
  const u = setTimeout(() => s("timeout"), t);
  try {
    return await Promise.race([a, e(n.signal)]);
  } catch (l) {
    throw Qce(l);
  } finally {
    (clearTimeout(u), r == null || r.removeEventListener("abort", d));
  }
}
const r1 = Object.freeze({
    claims: 20,
    sources: 20,
    rangesPerClaim: 20,
    referencedCells: 1e3,
    receiptBytes: 2e6,
  }),
  iy = string().regex(/^[A-Za-z][A-Za-z0-9_-]{0,59}$/),
  Pk = string().trim().min(1).max(1e4),
  tde = string()
    .regex(/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/)
    .max(32),
  rde = object({
    claims: array(
      discriminatedUnion("kind", [
        object({
          id: iy,
          kind: literal("workbook"),
          text: Pk,
          ranges: array(tde).min(1).max(r1.rangesPerClaim),
        }).strict(),
        object({
          id: iy,
          kind: literal("external"),
          text: Pk,
          sourceIds: array(iy).min(1).max(r1.sources),
        }).strict(),
        object({ id: iy, kind: literal("commentary"), text: Pk }).strict(),
      ]),
    )
      .min(1)
      .max(r1.claims),
    externalSources: array(
      object({
        id: iy,
        title: string().trim().min(1).max(300),
        url: string().max(2048),
      }).strict(),
    ).max(r1.sources),
  }).strict();
function TI(e) {
  if (!/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/.test(e))
    throw Error("Citation requires an explicit A1 cell or rectangle.");
  const r = ct(e),
    t = e.split(":");
  if (t[1] && `${ut(r.firstCol)}${r.firstRow + 1}` !== t[0])
    throw Error("Citation range must run top left to bottom right.");
  if ((r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1) > 500)
    throw Error("Citations support at most 500 cells per range.");
  const n = [];
  for (let o = r.firstRow; o <= r.lastRow; o++)
    for (let a = r.firstCol; a <= r.lastCol; a++) n.push(`${ut(a)}${o + 1}`);
  return n;
}
function nde(e) {
  if (
    typeof e != "string" ||
    e.length > 2048 ||
    e !== e.trim() ||
    /[\u0000-\u0020\u007f]/.test(e)
  )
    throw Error(
      "External sources require a valid HTTPS URL without credentials.",
    );
  let r;
  try {
    r = new URL(e);
  } catch {
    throw Error(
      "External sources require a valid HTTPS URL without credentials.",
    );
  }
  if (r.protocol !== "https:" || r.username || r.password || !r.hostname)
    throw Error(
      "External sources require a valid HTTPS URL without credentials.",
    );
  return r.href;
}
function ide(e, r) {
  const t = rde.safeParse(e);
  if (!t.success)
    throw Error(
      "Invalid answer grounding: provide typed claims and separate external sources.",
    );
  const n = t.data,
    o = new Set(r.cells.map((l) => l.address)),
    a = new Set(),
    s = new Set(),
    d = new Set();
  let u = 0;
  for (const l of n.externalSources) {
    if (s.has(l.id)) throw Error("External source IDs must be unique.");
    (s.add(l.id), (l.url = nde(l.url)));
  }
  if (n.claims.reduce((l, c) => l + c.text.length, 0) > 2e4)
    throw Error("Answer claims exceed the 20,000 character evidence limit.");
  for (const l of n.claims) {
    if (a.has(l.id)) throw Error("Answer claim IDs must be unique.");
    if ((a.add(l.id), l.kind === "workbook")) {
      if (new Set(l.ranges).size !== l.ranges.length)
        throw Error("Duplicate citation ranges in a claim.");
      for (const c of l.ranges) {
        const h = TI(c);
        if (((u += h.length), u > r1.referencedCells))
          throw Error(
            "Answer citations exceed the 1,000 cell-reference budget.",
          );
        if (h.some((f) => !o.has(f)))
          throw Error(
            "Answer citations must stay inside the approved selection.",
          );
      }
    } else if (l.kind === "external") {
      if (
        new Set(l.sourceIds).size !== l.sourceIds.length ||
        l.sourceIds.some((c) => !s.has(c))
      )
        throw Error(
          "External claims require distinct references to declared external sources.",
        );
      l.sourceIds.forEach((c) => d.add(c));
    }
  }
  if (n.externalSources.some((l) => !d.has(l.id)))
    throw Error("Unused external sources are not accepted as grounding.");
  return n;
}
const ode = Fc.filter(
    (e) =>
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
      ].includes(e),
  ),
  ade = new Set(ode),
  kg = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    Object.getPrototypeOf(e) === Object.prototype;
function Zr(e, r) {
  if (!e) throw new Error(r);
}
const Sg = (e, r) =>
  Zr(
    Object.keys(e).every((t) => r.includes(t)),
    "Unexpected field in LLM proposal.",
  );
function Zf(e) {
  const r = /^([A-Z]{1,3})([1-9]\d*)$/.exec(e);
  Zr(r, `Invalid cell address: ${e}`);
  const t = [...r[1]].reduce((o, a) => o * 26 + a.charCodeAt(0) - 64, 0),
    n = Number(r[2]);
  return (
    Zr(
      n <= Mt.maxRows && t <= Mt.maxCols,
      "Cell address exceeds supported dimensions.",
    ),
    { row: n, col: t }
  );
}
function Cz(e) {
  let r = "";
  for (; e > 0; e = Math.floor((e - 1) / 26))
    r = String.fromCharCode(65 + ((e - 1) % 26)) + r;
  return r;
}
function Az(e, r) {
  Zr(
    !r.ranges || (r.ranges.length === 1 && r.ranges[0] === r.range),
    "Select one rectangular range before using an assistant.",
  );
  const t = e.sheets.find((d) => d.id === r.sheetId);
  Zr(t && !t.hidden, "Select a visible worksheet before using an assistant.");
  const n = r.range.toUpperCase().split(":");
  Zr(n.length <= 2, "Select one rectangular range.");
  const o = Zf(n[0]),
    a = Zf(n[1] || n[0]);
  (Zr(
    o.row <= a.row && o.col <= a.col,
    "Selection range must run from top left to bottom right.",
  ),
    Zr(
      a.row <= t.rowCount && a.col <= t.colCount,
      "Selection exceeds worksheet dimensions.",
    ),
    Zr(
      (a.row - o.row + 1) * (a.col - o.col + 1) <= Mt.maxLLMCells,
      `Select at most ${Mt.maxLLMCells} cells for the assistant.`,
    ));
  const s = [];
  for (let d = o.row; d <= a.row; d++)
    for (let u = o.col; u <= a.col; u++) s.push(`${Cz(u)}${d}`);
  return s;
}
function s3(e, r) {
  const t = Az(e, r),
    n = e.sheets.find((a) => a.id === r.sheetId),
    o = Vn(e);
  for (const a of t) {
    const s = o.getTypedValue(n.id, a);
    (Zr(
      !(s.type === "error" && s.value === "#PENDING!"),
      "Calculation is pending in the approved selection. Choose Calculate Now before sharing current values with an assistant.",
    ),
      Zr(
        !(
          s.type === "error" &&
          ["#TIMEOUT!", "#LIMIT!", "#BLOCKED!"].includes(String(s.value))
        ),
        "Calculation failed at a resource or execution boundary in the approved selection. Resolve the limit and choose Calculate Now before sharing current values with an assistant.",
      ));
  }
  return {
    workbookId: e.id,
    ...(e.datePolicy ? { datePolicy: ts(e.datePolicy) } : {}),
    ...(eB(e) ? { calculationClock: o.getDateContext() } : {}),
    revision: e.revision,
    sheetId: n.id,
    sheetName: n.name,
    range: r.range.toUpperCase(),
    cells: t.map((a) => {
      return {
        address: a,
        raw: n.cells[a]?.raw ?? null,
        value: o(n.id, a),
      };
    }),
  };
}
function NI(e, r) {
  let t = e.replace(/"(?:[^"]|"")*"/g, '""');
  if (
    (Zr(
      !/[!\[\]|\\#]/.test(t),
      "Assistant formulas cannot reference other worksheets, external workbooks, table names, existing spills or executable content.",
    ),
    Zr(
      !/(?:https?:|javascript:|data:|file:)/i.test(e),
      "Assistant formulas cannot contain network or executable URLs.",
    ),
    /\b(?:LET|LAMBDA)\s*\(/i.test(t))
  ) {
    const a = [];
    Gf(Pg(t), (s, d, u, l) => {
      (l || (!s.qualifier && d.has(lf(s.name)))) &&
        a.push({ start: s.start, end: s.end, value: u ? "SUM" : "0" });
    });
    for (const s of a.sort((d, u) => u.start - d.start))
      t = t.slice(0, s.start) + s.value + t.slice(s.end);
  }
  const n = [...t.matchAll(/([A-Z_][A-Z0-9_.]*)\s*\(/gi)];
  for (const a of n)
    Zr(
      ade.has(a[1].toUpperCase()),
      `Assistant formula function ${a[1]} is not allowed.`,
    );
  for (const a of t.matchAll(
    new RegExp(
      "(?<![A-Z0-9_.])\\$?([A-Z]{1,3})\\$?([1-9]\\d*)(?![A-Z0-9_]|\\s*\\()",
      "gi",
    ),
  ))
    Zr(
      r.has(`${a[1].toUpperCase()}${a[2]}`),
      "Assistant formulas may only reference cells in the approved selection.",
    );
  const o = t
    .replace(/([A-Z_][A-Z0-9_.]*)\s*\(/gi, "(")
    .replace(/\$?[A-Z]{1,3}\$?[1-9]\d*(?![A-Z0-9_])/gi, "0")
    .replace(/""/g, "0")
    .replace(/\b(?:TRUE|FALSE)\b/gi, "0")
    .replace(/\d+(?:\.\d+)?(?:E[+-]?\d+)?/gi, "0");
  Zr(
    !/[A-Z_]/i.test(o),
    "Named ranges and unsupported formula tokens are not allowed in assistant changes.",
  );
}
function sde(e, r) {
  (Zr(kg(e), "Each proposed cell must be an object."),
    Sg(e, ["raw", "literal", "numFmt", "format", "style", "note", "tickmark"]),
    Zr(
      Object.hasOwn(e, "raw") &&
        (e.raw === null ||
          typeof e.raw == "string" ||
          typeof e.raw == "boolean" ||
          (typeof e.raw == "number" && Number.isFinite(e.raw))),
      "Unsupported proposed cell value.",
    ),
    typeof e.raw == "string" &&
      Zr(
        e.raw.length <= (e.raw.startsWith("=") ? 4096 : 32767),
        "Proposed cell text is too long.",
      ),
    e.literal !== undefined &&
      Zr(typeof e.literal == "boolean", "Cell literal must be boolean."),
    e.numFmt !== undefined &&
      Zr(
        typeof e.numFmt == "string" && e.numFmt.length <= 200,
        "Invalid number format.",
      ),
    e.format !== undefined &&
      Zr(
        ["general", "number", "currency", "percent", "date"].includes(
          String(e.format),
        ),
        "Unsupported cell format.",
      ));
  for (const t of ["note", "tickmark"])
    e[t] !== undefined &&
      Zr(
        typeof e[t] == "string" &&
          String(e[t]).length <= (t === "note" ? 1e4 : 40),
        `Invalid cell ${t}.`,
      );
  if (e.style !== undefined) {
    (Zr(kg(e.style), "Cell style must be an object."),
      Sg(e.style, [
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
    for (const t of ["bold", "italic", "underline", "wrap"])
      e.style[t] !== undefined &&
        Zr(typeof e.style[t] == "boolean", "Invalid font style.");
    (e.style.fontFamily !== undefined &&
      Zr(dx.includes(e.style.fontFamily), "Unsupported font family."),
      e.style.fontSize !== undefined &&
        Zr(
          typeof e.style.fontSize == "number" &&
            Number.isInteger(e.style.fontSize) &&
            e.style.fontSize >= 8 &&
            e.style.fontSize <= 72,
          "Font size must be an integer between 8 and 72.",
        ),
      e.style.border !== undefined &&
        Zr(
          ["all", "bottom", "none"].includes(String(e.style.border)),
          "Unsupported cell border.",
        ));
    for (const t of ["fill", "color"])
      e.style[t] !== undefined &&
        Zr(
          typeof e.style[t] == "string" &&
            /^#[0-9a-f]{3}([0-9a-f]{3})?$/i.test(String(e.style[t])),
          "Use hexadecimal colors.",
        );
    e.style.align !== undefined &&
      Zr(
        ["left", "center", "right"].includes(String(e.style.align)),
        "Invalid alignment.",
      );
  }
  return (
    typeof e.raw == "string" &&
      e.raw.startsWith("=") &&
      !e.literal &&
      NI(e.raw, r),
    structuredClone(e)
  );
}
function RI(e, r) {
  (Zr(kg(e), "The assistant response must be a proposal object."),
    Sg(e, [
      "id",
      "workbookId",
      "revision",
      "explanation",
      "citations",
      "patches",
      "grounding",
    ]),
    Zr(
      typeof e.id == "string" && e.id.length > 0 && e.id.length <= 100,
      "Invalid proposal identifier.",
    ),
    Zr(
      e.workbookId === r.workbookId && e.revision === r.revision,
      "This proposal is stale or belongs to another workbook. Generate a new proposal.",
    ),
    Zr(
      typeof e.explanation == "string" &&
        e.explanation.length > 0 &&
        e.explanation.length <= 1e4,
      "Invalid proposal explanation.",
    ));
  const t = new Set(r.cells.map((a) => a.address));
  (Zr(
    Array.isArray(e.citations) && e.citations.length <= Mt.maxLLMCells,
    "Invalid proposal citations.",
  ),
    Zr(
      e.citations.every((a) => typeof a == "string" && t.has(a)),
      "Citations must reference the approved selection.",
    ),
    Zr(
      Array.isArray(e.patches) && e.patches.length <= Mt.maxPatches,
      "Too many proposed edits.",
    ));
  const n = new Set(),
    o = e.patches.map(
      (a) => (
        Zr(kg(a), "Each patch must be an object."),
        Sg(a, ["sheetId", "address", "cell"]),
        Zr(
          a.sheetId === r.sheetId &&
            typeof a.address == "string" &&
            t.has(a.address),
          "Assistant changes must stay inside the approved selection.",
        ),
        Zr(!n.has(a.address), "The proposal contains duplicate cell edits."),
        n.add(a.address),
        { sheetId: r.sheetId, address: a.address, cell: sde(a.cell, t) }
      ),
    );
  return {
    id: e.id,
    workbookId: r.workbookId,
    revision: r.revision,
    explanation: e.explanation,
    citations: [...e.citations],
    patches: o,
    ...(e.grounding !== undefined ? { grounding: ide(e.grounding, r) } : {}),
  };
}
function X2(e, r, t) {
  const n = Az(e, r).map((c) => ({ address: c, raw: null, value: null })),
    o = RI(t, {
      workbookId: e.id,
      revision: e.revision,
      sheetId: r.sheetId,
      cells: n,
    }),
    a = e.sheets.find((c) => c.id === r.sheetId),
    s = {
      ...o,
      patches: o.patches.map((c) => {
        const h = a.cells[c.address],
          f = {
            ...h,
            ...c.cell,
            ...((h != null && h.style) || c.cell.style
              ? {
                  style: { ...h?.style, ...c.cell.style },
                }
              : {}),
          };
        (h?.raw !== c.cell.raw &&
          c.cell.literal === undefined &&
          delete f.literal,
          c.cell.format &&
            c.cell.format !== h?.format &&
            c.cell.numFmt === undefined &&
            delete f.numFmt);
        const p = Zf(c.address),
          m = (a.tables ?? []).some((v) => {
            const [g, b] = v.range.split(":"),
              y = Zf(g),
              w = Zf(b ?? g);
            return p.row === y.row && p.col >= y.col && p.col <= w.col;
          });
        return (
          Zr(
            !m ||
              (h?.raw === f.raw && !!(h != null && h.literal) == !!f.literal),
            "Assistant changes cannot rename table headers because dependent references may change outside the approved selection.",
          ),
          { ...c, cell: f }
        );
      }),
    },
    d = new Set(n.map((c) => c.address)),
    u = {
      ...e,
      sheets: e.sheets.map((c) =>
        c.id === a.id
          ? {
              ...c,
              cells: {
                ...c.cells,
                ...Object.fromEntries(
                  s.patches.map((h) => [h.address, h.cell]),
                ),
              },
            }
          : c,
      ),
    },
    l = Cf(
      u,
      {},
      {
        allowRead: () => true,
        allowDatabaseCriterion: (c, h) => {
          if (c !== a.id) return false;
          try {
            return (NI(h, d), true);
          } catch {
            return false;
          }
        },
        step: () => {},
      },
    );
  for (const c of s.patches) {
    if (
      typeof c.cell.raw != "string" ||
      !c.cell.raw.startsWith("=") ||
      c.cell.literal
    )
      continue;
    const h = l(a.id, c.address);
    (Zr(h !== "#LIMIT!", "The proposed formula exceeds calculation limits."),
      Zr(
        l.getOperationalErrors().length === 0,
        "The proposed formula has invalid syntax, a circular reference or an operational calculation failure. Correct the formula before review; error handlers cannot suppress this rejection.",
      ));
    const f = l.getSpillRange(a.id, c.address);
    if (
      (Zr(
        !l.getSpillDiagnostic(a.id, c.address),
        "The proposed array is blocked. Review the spill obstruction before requesting approval.",
      ),
      f)
    ) {
      const [p, m] = f.split(":"),
        v = Zf(p),
        g = Zf(m ?? p);
      Zr(
        (g.row - v.row + 1) * (g.col - v.col + 1) <= Mt.maxLLMCells,
        "Assistant spill outputs must stay inside the approved selection.",
      );
      for (let b = v.row; b <= g.row; b++)
        for (let y = v.col; y <= g.col; y++)
          Zr(
            d.has(`${Cz(y)}${b}`),
            "Assistant spill outputs must stay inside the approved selection.",
          );
    }
  }
  return (
    Zr(
      l.getOperationalErrors().length === 0,
      "The proposed formula cannot be validated because calculation encountered an operational failure.",
    ),
    s
  );
}
function kz() {
  return {
    name: "Local demo · no model or network",
    localOnly: true,
    chat: { history: true, attachments: true },
    async propose(e, r, t, n) {
      if (t != null && t.aborted)
        throw new DOMException("Request cancelled.", "AbortError");
      const o = {
          id: `local-${r.revision}-${Date.now()}`,
          workbookId: r.workbookId,
          revision: r.revision,
          explanation: "",
          citations: [],
          patches: [],
        },
        a = r.cells.filter(
          (d) => typeof d.value == "number" && Number.isFinite(d.value),
        ),
        s = a.reduce((d, u) => d + Number(u.value), 0);
      if (/highlight|high values|outlier/i.test(e)) {
        const d =
            /(?:above|over|greater than|exceed(?:ing|s)?)\s*[$€£]?\s*(-?\d[\d,]*(?:\.\d+)?)/i.exec(
              e,
            ),
          u = d
            ? Number(d[1].replaceAll(",", ""))
            : a.length
              ? Math.abs(s / a.length) * 1.5
              : 0,
          l = a.filter(
            (c) => (d ? Number(c.value) : Math.abs(Number(c.value))) > u,
          );
        ((o.explanation = `Local deterministic demo: ${l.length} values exceed ${d ? `the requested threshold of ${u.toLocaleString()}` : `1.5 times the absolute arithmetic mean (${u.toLocaleString()})`}. This is a simple illustration, not a statistical audit conclusion. The listed cells receive an amber fill; existing annotations are retained.`),
          (o.citations = l.map((c) => c.address)),
          (o.patches = l.map((c) => ({
            sheetId: r.sheetId,
            address: c.address,
            cell: { raw: c.raw, style: { fill: "#fff0b3" } },
          }))));
      } else if (/\b(?:total|sum)\b/i.test(e)) {
        const d = [...r.cells]
          .reverse()
          .find((u) => u.raw === null || u.raw === "");
        if (d && a.length) {
          const u = `=SUM(${a.map((l) => l.address).join(",")})`;
          ((o.patches = [
            {
              sheetId: r.sheetId,
              address: d.address,
              cell: { raw: u, format: "number", style: { bold: true } },
            },
          ]),
            (o.citations = a.map((l) => l.address)),
            (o.explanation = `Local deterministic demo: put a SUM formula for ${a.length} numeric selected cells in empty cell ${d.address}. Current result: ${s.toLocaleString()}. Review the included addresses for subtotals or duplicates before accepting.`));
        } else
          o.explanation =
            "Local deterministic demo: select numeric cells and at least one empty destination cell to propose a total.";
      } else if (/explain/i.test(e)) {
        const d = r.cells.filter(
          (u) => typeof u.raw == "string" && u.raw.startsWith("="),
        );
        ((o.explanation = d.length
          ? `Local deterministic demo: ${d.map((u) => `${u.address} contains ${u.raw} and currently returns ${String(u.value)}`).join("; ")}. This lists formulas and results; no language model interpretation is used.`
          : "Local deterministic demo: the selection contains no formula expressions. Values and formulas are worksheet data, never instructions."),
          (o.citations = d.map((u) => u.address)));
      } else
        ((o.explanation = `Local deterministic demo: ${r.cells.length} selected cells; ${a.length} numeric values; sum ${s.toLocaleString()}; ${r.cells.filter((d) => d.raw === null || d.raw === "").length} empty cells. Workbook content is treated only as data.`),
          (o.citations = a.slice(0, 20).map((d) => d.address)));
      if (
        ((o.grounding = {
          claims: [
            {
              id: "local-evidence",
              kind: "workbook",
              text: o.explanation,
              ranges: [r.range],
            },
          ],
          externalSources: [],
        }),
        n != null && n.attachments.length)
      ) {
        const d = Ez(n).attachments,
          u = `Received ${d.length} attachment${d.length === 1 ? "" : "s"}: ${d.map((l) => `${l.name} (${l.size.toLocaleString()} bytes)`).join(", ")}. Files remain in this browser in local demo mode. Preview text and images in the attachment cards; document interpretation requires a connected model.`;
        ((o.explanation += `

${u}`),
          o.grounding.claims.push({
            id: "local-attachments",
            kind: "commentary",
            text: u,
          }));
      }
      return RI(o, r);
    },
  };
}
function lde(e) {
  try {
    (Zr(kg(e), "context"),
      Sg(e, [
        "workbookId",
        "revision",
        "sheetId",
        "sheetName",
        "range",
        "cells",
        "datePolicy",
        "calculationClock",
      ]));
    const r = ts(e.datePolicy);
    if (e.calculationClock !== undefined) {
      const d = e.calculationClock;
      (Zr(kg(d), "clock"),
        Zr(
          typeof d.calculatedAt == "string" &&
            /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(d.calculatedAt),
          "clock timestamp",
        ));
      const u = lx(r, new Date(d.calculatedAt));
      (Sg(d, Object.keys(u)),
        Zr(
          Object.entries(u).every(([l, c]) => d[l] === c),
          "clock policy/value mismatch",
        ));
    }
    for (const d of ["workbookId", "sheetId", "sheetName"])
      Zr(
        typeof e[d] == "string" && e[d].length > 0 && e[d].length <= 200,
        "identity",
      );
    (Zr(Number.isSafeInteger(e.revision) && e.revision >= 0, "revision"),
      Zr(typeof e.range == "string" && e.range.length <= 64, "range"));
    const t = e.range.split(":");
    Zr(t.length <= 2, "range");
    const n = Zf(t[0]),
      o = Zf(t[1] ?? t[0]),
      a = (o.row - n.row + 1) * (o.col - n.col + 1);
    (Zr(o.row >= n.row && o.col >= n.col && a <= Mt.maxLLMCells, "bounds"),
      Zr(Array.isArray(e.cells) && e.cells.length === a, "population"));
    const s = new Set();
    for (const d of e.cells) {
      (Zr(kg(d), "cell"),
        Sg(d, ["address", "raw", "value"]),
        Zr(typeof d.address == "string" && !s.has(d.address), "address"));
      const u = Zf(d.address);
      (Zr(
        u.row >= n.row && u.row <= o.row && u.col >= n.col && u.col <= o.col,
        "scope",
      ),
        s.add(d.address));
      for (const l of [d.raw, d.value])
        Zr(
          l === null ||
            typeof l == "boolean" ||
            (typeof l == "number" && Number.isFinite(l)) ||
            (typeof l == "string" && l.length <= 32767),
          "value",
        );
    }
  } catch {
    throw new cf("invalid_request");
  }
}

export {
  Ez,
  Kce,
  NI,
  RI,
  TI,
  X2,
  Yce,
  Zce,
  Zm,
  cEe,
  cf,
  dEe,
  ede,
  kz,
  lde,
  r1,
  s3,
};
