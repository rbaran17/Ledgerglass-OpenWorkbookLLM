import { formatCellValueForDisplay } from "./17-recipe-step-editors.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  fileImportLimits,
  workbookLimits,
} from "./02-text-and-financial-functions.js";
import {
  normalizeTextCleanups,
  createWorksheetRecord,
  findRowOrigin,
  calculateWorkbookWithSettings,
  validateTextCleanupSteps,
  parseRangeAddress,
  normalizeWorkbookInput,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  buildConsolidationOutput,
  assertConsolidationPreviewCurrent,
} from "./12-consolidation-and-dependency-inspection.js";
import {
  assertSelectionInAuthorizedSheet,
  validateTextToColumnsPermissions,
  assertOutputNameUnreferenced,
} from "./15-data-model-dialog.js";
import {
  discoverArrayFormulaAnchors,
  applyTextSplitPreview,
  buildTextSplitResult,
} from "./16-transformation-recipe-engine.js";
import { w2 } from "./41-automation-and-scripts.js";

function CombinePopulationsDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(null),
    s = ne.useRef(e.onClose);
  s.current = e.onClose;
  const [sheetId, u] = ne.useState(e.sheetId),
    [l, c] = ne.useState(e.sheetId),
    [range, f] = ne.useState(e.range),
    [p, m] = ne.useState(""),
    [mode, g] = ne.useState("append"),
    [outputSheetName, y] = ne.useState("Combined population"),
    [keys, E] = ne.useState([{ left: "", right: "" }]),
    [normalization, k] = ne.useState({}),
    [checked, I] = ne.useState(false),
    [O, j] = ne.useState(0),
    [z, V] = ne.useState("0.8"),
    [D, B] = ne.useState(false),
    [_, C] = ne.useState(null),
    [T, N] = ne.useState(""),
    [R, L] = ne.useState(""),
    [$, U] = ne.useState(""),
    [G, H] = ne.useState("output"),
    [Q, X$1] = ne.useState(0),
    J = {
      includeHidden: !!e.allowHidden && D,
      allowedSheetIds: e.allowedSheetIds,
      readOnly: e.readOnly,
    },
    q = [
      ...new Set(
        e.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    W = e.workbook.sheets.filter(
      (sheet) =>
        q.includes(sheet.id) &&
        (J.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    K = JSON.stringify([e.sheetId, e.range]),
    te = JSON.stringify(e.workbook),
    ae = JSON.stringify([te, K, q, J.includeHidden, e.readOnly]),
    oe =
      !!_ &&
      !e.readOnly &&
      T === K &&
      _.snapshot === te &&
      _.context.includeHidden === J.includeHidden &&
      JSON.stringify(_.context.allowedSheetIds) === JSON.stringify(q);
  (ne.useEffect(() => {
    var Re;
    _ &&
      oe &&
      ((Re = a.current) == null ? undefined : Re.scrollIntoView)?.call(Re, {
        block: "start",
      });
  }, [_, oe]),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        o.current
          ?.querySelector('[data-ui-label="Combine operation"]')
          ?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  const onClick = () => {
      (C(null), L(""), X$1(0));
    },
    he = (Re, Me) => {
      (Re(Me), onClick());
    },
    Se = (Re) => {
      (U(ae),
        L(
          Re instanceof Error
            ? Re.message
            : "The populations could not be combined.",
        ));
    },
    le = () => {
      try {
        (C(
          buildConsolidationOutput(
            e.workbook,
            {
              left: { sheetId: sheetId, range: range },
              right: { sheetId: l, range: p },
              mode: mode,
              ...(mode === "append"
                ? {}
                : {
                    keys: keys,
                    ...(Object.values(normalization).some(Boolean)
                      ? { normalization: normalization }
                      : {}),
                    ...(checked
                      ? {
                          fuzzy: {
                            key: O,
                            threshold: z.trim() ? Number(z) : NaN,
                          },
                        }
                      : {}),
                  }),
              outputSheetName: outputSheetName,
            },
            J,
          ),
        ),
          N(K),
          L(""),
          X$1(0),
          H("output"));
      } catch (error) {
        (C(null), Se(error));
      }
    },
    ce = () => {
      if (!(!_ || !oe))
        try {
          const Re = assertConsolidationPreviewCurrent(e.workbook, _, J);
          (e.onWorkbookChange(
            Re,
            `Combine populations: ${_.request.mode} → ${_.request.outputSheetName}`,
            _.outputSheetId,
          ),
            C(null),
            e.onClose());
        } catch (error) {
          Se(error);
        }
    },
    ge = oe
      ? G === "left"
        ? ["Source row", "Match outcome", ..._.left.headers]
        : G === "right"
          ? ["Source row", "Match outcome", ..._.right.headers]
          : G === "matches"
            ? ["Left source", "Right source", "Similarity", "Explanation"]
            : G === "normalization"
              ? ["Side", "Source", "Original keys", "Normalized keys"]
              : G === "exceptions"
                ? ["Side", "Source", "Exception"]
                : _.headers
      : [],
    se = oe
      ? G === "left"
        ? _.unmatchedLeft.map((Re) => [
            Re.locator,
            Re.unmatchedReason ?? "",
            ...Re.values,
          ])
        : G === "right"
          ? _.unmatchedRight.map((Re) => [
              Re.locator,
              Re.unmatchedReason ?? "",
              ...Re.values,
            ])
          : G === "matches"
            ? (_.matching?.matches ?? []).map((Re) => [
                Re.left,
                Re.right,
                Re.score,
                Re.explanation,
              ])
            : G === "normalization"
              ? (_.matching?.normalizedKeys ?? []).map((Re) => [
                  Re.side,
                  Re.locator,
                  JSON.stringify(Re.original),
                  JSON.stringify(Re.normalized),
                ])
              : G === "exceptions"
                ? (_.matching?.exceptions ?? []).map((Re) => [
                    Re.side,
                    Re.locator,
                    Re.reason,
                  ])
                : _.rows.map((row) => [...row.values, row.left, row.right])
      : [],
    we = Math.max(1, Math.ceil(se.length / w2)),
    Ee = Math.min(Q, we - 1);
  return i.jsx("div", {
    className: "pc-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      ref: o,
      className: "pc-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (event) => {
        var ke, je;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), s.current()),
          event.key === "Tab")
        ) {
          const $e = [
            ...(o.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled)",
            ) ?? []),
          ].filter((Je) => !Je.hidden);
          event.shiftKey && document.activeElement === $e[0]
            ? (event.preventDefault(), (ke = $e.at(-1)) == null || ke.focus())
            : !event.shiftKey &&
              document.activeElement === $e.at(-1) &&
              (event.preventDefault(), (je = $e[0]) == null || je.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Combine populations",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Review appended records and explained matches before creating a new worksheet.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close Combine populations"),
              onClick: e.onClose,
              lang: language("Close Combine populations"),
              "data-ui-label": "Close Combine populations",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "pc-body",
          children: [
            i.jsxs("div", {
              className: "pc-fields",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Operation" }),
                    i.jsxs("select", {
                      "aria-label": text("Combine operation"),
                      disabled: e.readOnly,
                      value: mode,
                      onChange: (event) => {
                        (g(event.target.value), onClick());
                      },
                      lang: language("Combine operation"),
                      "data-ui-label": "Combine operation",
                      children: [
                        i.jsx("option", {
                          value: "append",
                          children: "Append rows by matching headers",
                        }),
                        i.jsx("option", {
                          value: "inner",
                          children: "Inner join — matched rows",
                        }),
                        i.jsx("option", {
                          value: "left",
                          children: "Left join — all left rows",
                        }),
                        i.jsx("option", {
                          value: "right",
                          children: "Right join — all right rows",
                        }),
                        i.jsx("option", {
                          value: "full",
                          children: "Full join — all rows",
                        }),
                        i.jsx("option", {
                          value: "leftAnti",
                          children: "Left anti join — unmatched left rows",
                        }),
                        i.jsx("option", {
                          value: "rightAnti",
                          children: "Right anti join — unmatched right rows",
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "New worksheet name" }),
                    i.jsx("input", {
                      "aria-label": text("Combined worksheet name"),
                      value: outputSheetName,
                      disabled: e.readOnly,
                      onChange: (event) => he(y, event.target.value),
                      lang: language("Combined worksheet name"),
                      "data-ui-label": "Combined worksheet name",
                    }),
                  ],
                }),
              ],
            }),
            i.jsx("div", {
              className: "pc-sources",
              children: ["left", "right"].map((Re) => {
                const Me = Re === "left" ? sheetId : l,
                  ke = Re === "left" ? range : p;
                return i.jsxs(
                  "fieldset",
                  {
                    disabled: e.readOnly,
                    children: [
                      i.jsxs("legend", {
                        children: [
                          Re === "left" ? "Left" : "Right",
                          " ",
                          i.jsx(LocalizedText, { children: "source" }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Worksheet" }),
                          i.jsxs("select", {
                            "aria-label": `${Re === "left" ? "Left" : "Right"} source worksheet`,
                            value: W.some((je) => je.id === Me) ? Me : "",
                            onChange: (event) =>
                              he(Re === "left" ? u : c, event.target.value),
                            children: [
                              i.jsx("option", {
                                value: "",
                                children: "Choose an allowed worksheet",
                              }),
                              W.map((je) =>
                                i.jsx(
                                  "option",
                                  { value: je.id, children: je.name },
                                  je.id,
                                ),
                              ),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Headed range" }),
                          i.jsx("input", {
                            "aria-label": `${Re === "left" ? "Left" : "Right"} source range`,
                            value: ke,
                            placeholder: text("A1:D20"),
                            onChange: (event) =>
                              he(Re === "left" ? f : m, event.target.value),
                            lang: language("A1:D20"),
                          }),
                        ],
                      }),
                    ],
                  },
                  Re,
                );
              }),
            }),
            mode !== "append" &&
              i.jsxs("fieldset", {
                className: "pc-keys",
                disabled: e.readOnly,
                children: [
                  i.jsx("legend", {
                    children: i.jsx(LocalizedText, {
                      children: "Matching headers in priority order",
                    }),
                  }),
                  keys.map((key, Me) =>
                    i.jsxs(
                      "div",
                      {
                        className: "pc-key",
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Left header ",
                              }),
                              Me + 1,
                              i.jsx("input", {
                                "aria-label": `Left key ${Me + 1}`,
                                value: key.left,
                                onChange: (event) => {
                                  (E(
                                    keys.map((key, $e) =>
                                      Me === $e
                                        ? { ...key, left: event.target.value }
                                        : key,
                                    ),
                                  ),
                                    onClick());
                                },
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Right header ",
                              }),
                              Me + 1,
                              i.jsx("input", {
                                "aria-label": `Right key ${Me + 1}`,
                                value: key.right,
                                onChange: (event) => {
                                  (E(
                                    keys.map((key, $e) =>
                                      Me === $e
                                        ? { ...key, right: event.target.value }
                                        : key,
                                    ),
                                  ),
                                    onClick());
                                },
                              }),
                            ],
                          }),
                          i.jsx("button", {
                            disabled: keys.length === 1,
                            "aria-label": `Remove key ${Me + 1}`,
                            onClick: () => {
                              (E(keys.filter((key, je) => je !== Me)),
                                j(0),
                                onClick());
                            },
                            children: i.jsx(LocalizedText, {
                              children: "Remove",
                            }),
                          }),
                        ],
                      },
                      Me,
                    ),
                  ),
                  i.jsx("button", {
                    disabled: keys.length >= 8,
                    onClick: () => {
                      (E([...keys, { left: "", right: "" }]), onClick());
                    },
                    children: i.jsx(LocalizedText, {
                      children: "Add matching key",
                    }),
                  }),
                ],
              }),
            mode !== "append" &&
              i.jsxs("fieldset", {
                className: "pc-keys",
                disabled: e.readOnly,
                children: [
                  i.jsx("legend", {
                    children: i.jsx(LocalizedText, {
                      children: "Key normalization and fuzzy matching",
                    }),
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "These choices affect comparison only. Original key values stay in the source and output.",
                    }),
                  }),
                  i.jsxs("div", {
                    className: "pc-normalization",
                    children: [
                      [
                        ["trim", "Trim surrounding whitespace"],
                        ["collapseWhitespace", "Collapse repeated whitespace"],
                        ["ignoreCase", "Ignore text case"],
                      ].map(([Re, Me]) =>
                        i.jsxs(
                          "label",
                          {
                            className: "pc-checkbox",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: !!normalization[Re],
                                onChange: (event) => {
                                  (k({
                                    ...normalization,
                                    [Re]: event.target.checked,
                                  }),
                                    onClick());
                                },
                              }),
                              Me,
                            ],
                          },
                          Re,
                        ),
                      ),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Unicode normalization",
                          }),
                          i.jsxs("select", {
                            "aria-label": text("Key Unicode normalization"),
                            value: normalization.unicode ?? "",
                            onChange: (event) => {
                              (k({
                                ...normalization,
                                unicode: event.target.value
                                  ? event.target.value
                                  : undefined,
                              }),
                                onClick());
                            },
                            lang: language("Key Unicode normalization"),
                            "data-ui-label": "Key Unicode normalization",
                            children: [
                              i.jsx("option", { value: "", children: "None" }),
                              i.jsx("option", {
                                value: "NFC",
                                children: "NFC · equivalent characters",
                              }),
                              i.jsx("option", {
                                value: "NFKC",
                                children: "NFKC · compatibility characters",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    className: "pc-checkbox",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: checked,
                        onChange: (event) => {
                          (I(event.target.checked), onClick());
                        },
                      }),
                      i.jsx(LocalizedText, {
                        children: "Use fuzzy text matching",
                      }),
                    ],
                  }),
                  checked &&
                    i.jsxs("div", {
                      className: "pc-fields",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Fuzzy key pair",
                            }),
                            i.jsx("select", {
                              "aria-label": text("Fuzzy key pair"),
                              value: O,
                              onChange: (event) => {
                                (j(Number(event.target.value)), onClick());
                              },
                              lang: language("Fuzzy key pair"),
                              "data-ui-label": "Fuzzy key pair",
                              children: keys.map((key, Me) =>
                                i.jsxs(
                                  "option",
                                  {
                                    value: Me,
                                    children: [
                                      "Pair ",
                                      Me + 1,
                                      ": ",
                                      key.left || "Left",
                                      " ↔",
                                      " ",
                                      key.right || "Right",
                                    ],
                                  },
                                  Me,
                                ),
                              ),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Minimum similarity (0–1)",
                            }),
                            i.jsx(LocalizedNumberField, {
                              "aria-label": text("Fuzzy similarity threshold"),
                              min: "0",
                              max: "1",
                              step: "0.01",
                              value: z,
                              onValueChange: (Re) => {
                                (V(Re), onClick());
                              },
                              lang: language("Fuzzy similarity threshold"),
                              "data-ui-label": "Fuzzy similarity threshold",
                            }),
                          ],
                        }),
                      ],
                    }),
                  checked &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Similarity uses normalized edit distance on Unicode characters. All pairs at or above the threshold are retained, including ties. Other key pairs must match exactly after the selected normalization. Scores are not probabilities or Power Query scores. Text-only fuzzy keys; 512 characters per key and a bounded comparison workload.",
                      }),
                    }),
                ],
              }),
            i.jsx("p", {
              className: "pc-policy",
              children: i.jsx(LocalizedText, {
                children:
                  "Headers match exactly; append may reorder columns. Join keys preserve types; selected text normalization is explicit. Blank keys do not match. Every record in each rectangle is included regardless of the current filter view. Output is a values snapshot with source row columns and notes.",
              }),
            }),
            e.allowHidden &&
              i.jsxs("label", {
                className: "pc-checkbox",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    "aria-label": text("Include hidden source records"),
                    checked: D,
                    disabled: e.readOnly,
                    onChange: (event) => {
                      (B(event.target.checked), onClick());
                    },
                    lang: language("Include hidden source records"),
                    "data-ui-label": "Include hidden source records",
                  }),
                  i.jsx(LocalizedText, {
                    children:
                      "Include hidden source records and formula dependencies",
                  }),
                ],
              }),
            e.readOnly &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "The host is read-only. Preview and Apply are unavailable.",
                }),
              }),
            R &&
              $ === ae &&
              i.jsx("p", { className: "pc-error", role: "alert", children: R }),
            i.jsxs("div", {
              className: "pc-actions",
              children: [
                i.jsx("button", {
                  disabled: e.readOnly,
                  onClick: le,
                  children: i.jsx(LocalizedText, {
                    children: "Preview combination",
                  }),
                }),
                _ &&
                  i.jsx("button", {
                    onClick: onClick,
                    children: i.jsx(LocalizedText, {
                      children: "Discard preview",
                    }),
                  }),
              ],
            }),
            _ &&
              !oe &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "The workbook or host context changed. Preview details are withheld. Review a fresh preview.",
                }),
              }),
            oe &&
              _ &&
              i.jsxs("section", {
                className: "pc-review",
                ref: a,
                "aria-label": text("Population combination preview"),
                lang: language("Population combination preview"),
                "data-ui-label": "Population combination preview",
                children: [
                  i.jsxs("h3", {
                    children: [_.request.outputSheetName, " · ", _.outputRange],
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, { children: "Left: " }),
                      _.counts.leftRows,
                      " ",
                      i.jsx(LocalizedText, { children: "records. Right:" }),
                      " ",
                      _.counts.rightRows,
                      " ",
                      i.jsx(LocalizedText, { children: "records. Output:" }),
                      " ",
                      _.counts.outputRows,
                      " ",
                      i.jsx(LocalizedText, { children: "records." }),
                    ],
                  }),
                  mode !== "append" &&
                    i.jsxs(i.Fragment, {
                      children: [
                        _.keyProfiles.map((keyProfile, Me) =>
                          i.jsxs(
                            "p",
                            {
                              children: [
                                i.jsx(LocalizedText, { children: "Key " }),
                                Me + 1,
                                ": ",
                                keyProfile.leftHeader,
                                " (",
                                keyProfile.leftTypes.join(", ") || "no records",
                                ") ↔",
                                " ",
                                keyProfile.rightHeader,
                                " (",
                                keyProfile.rightTypes.join(", ") ||
                                  "no records",
                                ").",
                              ],
                            },
                            Me,
                          ),
                        ),
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Matched source rows: ",
                            }),
                            _.counts.matchedLeftRows,
                            " ",
                            i.jsx(LocalizedText, { children: "left /" }),
                            " ",
                            _.counts.matchedRightRows,
                            " ",
                            i.jsx(LocalizedText, {
                              children: "right. Match pairs:",
                            }),
                            " ",
                            _.counts.matchPairs,
                            i.jsx(LocalizedText, { children: ". Unmatched:" }),
                            " ",
                            _.unmatchedLeft.length,
                            " ",
                            i.jsx(LocalizedText, { children: "left /" }),
                            " ",
                            _.unmatchedRight.length,
                            " ",
                            i.jsx(LocalizedText, { children: "right." }),
                          ],
                        }),
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Duplicate key groups: ",
                            }),
                            _.counts.duplicateLeftKeys,
                            " ",
                            i.jsx(LocalizedText, { children: "left /" }),
                            " ",
                            _.counts.duplicateRightKeys,
                            " ",
                            i.jsx(LocalizedText, {
                              children: "right. Many-to-many key groups: ",
                            }),
                            _.counts.manyToManyKeys,
                            i.jsx(LocalizedText, {
                              children: ". Extra many-to-many pairs: ",
                            }),
                            _.counts.expandedMatchRows,
                            ".",
                          ],
                        }),
                        _.matching &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("p", { children: _.matching.policy }),
                              i.jsxs("p", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Rows with multiple matches:",
                                  }),
                                  " ",
                                  _.matching.ambiguousLeftRows,
                                  " ",
                                  i.jsx(LocalizedText, { children: "left /" }),
                                  " ",
                                  _.matching.ambiguousRightRows,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children:
                                      "right. Normalization collision groups:",
                                  }),
                                  " ",
                                  _.matching.normalizationCollisionGroups.left,
                                  " ",
                                  i.jsx(LocalizedText, { children: "left /" }),
                                  " ",
                                  _.matching.normalizationCollisionGroups.right,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children:
                                      "right. Review every match; no automatic best-match choice is made.",
                                  }),
                                ],
                              }),
                              (mode === "leftAnti" || mode === "rightAnti") &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children:
                                        "Anti output contains only unmatched",
                                    }),
                                    " ",
                                    mode === "leftAnti" ? "left" : "right",
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children:
                                        "records. Matched pairs and both unmatched lists remain available below.",
                                    }),
                                  ],
                                }),
                            ],
                          }),
                      ],
                    }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Review records" }),
                      i.jsxs("select", {
                        "aria-label": text("Population review view"),
                        value: G,
                        onChange: (event) => {
                          (H(event.target.value), X$1(0));
                        },
                        lang: language("Population review view"),
                        "data-ui-label": "Population review view",
                        children: [
                          i.jsx("option", {
                            value: "output",
                            children: "Complete output",
                          }),
                          mode !== "append" &&
                            i.jsxs(i.Fragment, {
                              children: [
                                i.jsxs("option", {
                                  value: "left",
                                  children: [
                                    "All unmatched left records (",
                                    _.unmatchedLeft.length,
                                    ")",
                                  ],
                                }),
                                i.jsxs("option", {
                                  value: "right",
                                  children: [
                                    "All unmatched right records (",
                                    _.unmatchedRight.length,
                                    ")",
                                  ],
                                }),
                                i.jsxs("option", {
                                  value: "matches",
                                  children: [
                                    "All match scores and explanations (",
                                    _.matching?.matches.length ?? 0,
                                    ")",
                                  ],
                                }),
                                i.jsxs("option", {
                                  value: "normalization",
                                  children: [
                                    "Changed keys (",
                                    _.matching?.normalizedKeys.length ?? 0,
                                    ")",
                                  ],
                                }),
                                i.jsxs("option", {
                                  value: "exceptions",
                                  children: [
                                    "Match exceptions (",
                                    _.matching?.exceptions.length ?? 0,
                                    ")",
                                  ],
                                }),
                              ],
                            }),
                        ],
                      }),
                    ],
                  }),
                  i.jsx("div", {
                    className: "pc-table",
                    tabIndex: 0,
                    "aria-label": text("Scrollable population records"),
                    lang: language("Scrollable population records"),
                    "data-ui-label": "Scrollable population records",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsx("tr", {
                            children: ge.map((Re, Me) =>
                              i.jsx("th", { children: Re }, Me),
                            ),
                          }),
                        }),
                        i.jsx("tbody", {
                          children: se
                            .slice(Ee * w2, (Ee + 1) * w2)
                            .map((Re, Me) =>
                              i.jsx(
                                "tr",
                                {
                                  children: Re.map((ke, je) =>
                                    i.jsx(
                                      "td",
                                      {
                                        title:
                                          ke === null ? "Blank" : typeof ke,
                                        children: formatCellValueForDisplay(ke),
                                      },
                                      je,
                                    ),
                                  ),
                                },
                                Ee * w2 + Me,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  !se.length &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "No records in this view. A zero-match output contains its headers only.",
                      }),
                    }),
                  i.jsxs("div", {
                    className: "pc-pages",
                    children: [
                      i.jsx("button", {
                        disabled: !Ee,
                        onClick: () => X$1(Ee - 1),
                        children: i.jsx(LocalizedText, {
                          children: "Previous records",
                        }),
                      }),
                      i.jsxs("span", {
                        "aria-live": "polite",
                        children: [
                          i.jsx(LocalizedText, { children: "Page " }),
                          Ee + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          we,
                          " · ",
                          se.length,
                          " ",
                          i.jsx(LocalizedText, { children: "records" }),
                        ],
                      }),
                      i.jsx("button", {
                        disabled: Ee >= we - 1,
                        onClick: () => X$1(Ee + 1),
                        children: i.jsx(LocalizedText, {
                          children: "Next records",
                        }),
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Source cells remain unchanged. Applying creates one worksheet; the host records one change for Undo.",
                    }),
                  }),
                  i.jsx("button", {
                    onClick: () => {
                      const Re = {
                          schemaVersion: 1,
                          workbookId: _.next.id,
                          revision: _.next.revision,
                          request: _.request,
                          context: _.context,
                          left: _.left,
                          right: _.right,
                          counts: _.counts,
                          headers: _.headers,
                          rows: _.rows,
                          unmatchedLeft: _.unmatchedLeft,
                          unmatchedRight: _.unmatchedRight,
                          keyProfiles: _.keyProfiles,
                          matching: _.matching,
                        },
                        Me = URL.createObjectURL(
                          new Blob([JSON.stringify(Re, null, 2)], {
                            type: "application/json",
                          }),
                        ),
                        ke = document.createElement("a");
                      ((ke.href = Me),
                        (ke.download = "Population-combine-review.json"),
                        ke.click(),
                        setTimeout(() => URL.revokeObjectURL(Me), 5e3));
                    },
                    children: i.jsx(LocalizedText, {
                      children: "Save combination review",
                    }),
                  }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              disabled: !oe,
              onClick: ce,
              children: i.jsx(LocalizedText, {
                children: "Create combined worksheet",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
const Fv = 20,
  stringifyJsonValue = (e) => JSON.stringify(e);
function TextToColumnsDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(false),
    [sourceSheetId, d] = ne.useState(e.sheetId),
    [sourceRange, l] = ne.useState(e.range),
    [separator, h] = ne.useState(","),
    [custom, p] = ne.useState("|"),
    [m, v] = ne.useState(false),
    [destination, b] = ne.useState(
      e.allowNewSheet === false ? "range" : "new-sheet",
    ),
    [y, w] = ne.useState(e.sheetId),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState("Split text"),
    [checked, O] = ne.useState(false),
    [formats, z] = ne.useState([]),
    [V, D] = ne.useState(0),
    [B, _] = ne.useState(null),
    [C, T] = ne.useState(0),
    [N, R] = ne.useState(null),
    allowedSheetIds = [
      ...new Set(
        e.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    $ = {
      readOnly: e.readOnly,
      allowedSheetIds: allowedSheetIds,
      includeHidden: !!e.allowHidden && checked,
      allowNewSheet: e.allowNewSheet ?? true,
    },
    U = e.workbook.sheets.filter(
      (sheet) =>
        allowedSheetIds.includes(sheet.id) &&
        ($.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    G = JSON.stringify({
      sourceSheetId: sourceSheetId,
      sourceRange: sourceRange,
      separator: separator,
      custom: custom,
      header: m,
      destination: destination,
      targetSheetId: y,
      targetAddress: E,
      name: k,
      formats: formats,
      options: $,
      selection: [e.sheetId, e.range],
    }),
    H = B?.key === G && B.plan.snapshot === JSON.stringify(e.workbook),
    Q = H ? B.plan : null;
  (ne.useEffect(() => {
    const activeElement = document.activeElement,
      oe = o.current?.querySelector('[data-ui-label="Source range"]');
    return (
      (oe && !oe.matches(":disabled")
        ? oe
        : o.current?.querySelector("button")
      )?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      (d(e.sheetId), l(e.range), z([]), D(0), _(null), R(null));
    }, [e.sheetId, e.range]));
  const X$1 = () => {
      (_(null), R(null), (a.current = false));
    },
    J = (ae, oe) => {
      (ae(oe), X$1());
    },
    q = (ae) =>
      R({
        key: G,
        snapshot: JSON.stringify(e.workbook),
        message:
          ae instanceof Error
            ? ae.message
            : "Text could not be split. Review the source and settings.",
      }),
    onClick = () => {
      X$1();
      try {
        const request = {
            sourceSheetId: sourceSheetId,
            sourceRange: sourceRange,
            delimiter: separator === "custom" ? custom : separator,
            header: m ? "first-row" : "none",
            formats: formats,
            destination:
              destination === "new-sheet"
                ? { kind: "new-sheet", name: k }
                : { kind: "range", sheetId: y, address: E },
          },
          textSplitResult = buildTextSplitResult(e.workbook, request, $);
        (D(textSplitResult.width), T(0), _({ plan: textSplitResult, key: G }));
      } catch (error) {
        q(error);
      }
    },
    K = () => {
      if (!(!Q || a.current || e.readOnly))
        try {
          const ae = applyTextSplitPreview(e.workbook, Q, $);
          ((a.current = true),
            e.onWorkbookChange(
              ae,
              `Text to Columns: ${Q.sourceSheetName}!${Q.request.sourceRange} → ${Q.targetSheetName}!${Q.targetRange}`,
            ),
            _(null),
            R(null));
        } catch (error) {
          q(error);
        }
    },
    te = (ae) => U.some((oe) => oe.id === ae);
  return i.jsx("div", {
    className: "ttc-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      ref: o,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      className: "ttc-dialog",
      onKeyDown: (event) => {
        var de, he;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), e.onClose()),
          event.key === "Tab")
        ) {
          const Se = [
            ...(o.current?.querySelectorAll("button,input,select,summary") ??
              []),
          ].filter(
            (le) =>
              !le.matches(":disabled") &&
              !le.hidden &&
              !le.closest("details:not([open]) > :not(summary)"),
          );
          event.shiftKey && document.activeElement === Se[0]
            ? (event.preventDefault(), (de = Se.at(-1)) == null || de.focus())
            : !event.shiftKey &&
              document.activeElement === Se.at(-1) &&
              (event.preventDefault(), (he = Se[0]) == null || he.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Text to Columns",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Split one source column into a separate range. Review every output before Apply.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close Text to Columns"),
              onClick: e.onClose,
              lang: language("Close Text to Columns"),
              "data-ui-label": "Close Text to Columns",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Original cells and evidence remain in place. Text keeps leading zeros and long identifiers. General converts plain numbers and TRUE/FALSE; dates, currency and locale formats stay text.",
          }),
        }),
        i.jsxs("fieldset", {
          disabled: e.readOnly,
          className: "ttc-fields",
          children: [
            i.jsx("legend", {
              children: i.jsx(LocalizedText, { children: "Split settings" }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source worksheet" }),
                i.jsxs("select", {
                  "aria-label": text("Source worksheet"),
                  value: te(sourceSheetId) ? sourceSheetId : "",
                  onChange: (event) => {
                    (J(d, event.target.value), z([]), D(0));
                  },
                  lang: language("Source worksheet"),
                  "data-ui-label": "Source worksheet",
                  children: [
                    i.jsx("option", {
                      value: "",
                      disabled: true,
                      children: "Choose worksheet",
                    }),
                    U.map((ae) =>
                      i.jsx(
                        "option",
                        { value: ae.id, children: ae.name },
                        ae.id,
                      ),
                    ),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Source range"),
                  value: sourceRange,
                  onChange: (event) => {
                    (J(l, event.target.value), z([]), D(0));
                  },
                  lang: language("Source range"),
                  "data-ui-label": "Source range",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Delimiter" }),
                i.jsxs("select", {
                  "aria-label": text("Delimiter"),
                  value: separator,
                  onChange: (event) => {
                    (J(h, event.target.value), z([]), D(0));
                  },
                  lang: language("Delimiter"),
                  "data-ui-label": "Delimiter",
                  children: [
                    i.jsx("option", { value: "	", children: "Tab" }),
                    i.jsx("option", { value: ",", children: "Comma" }),
                    i.jsx("option", { value: ";", children: "Semicolon" }),
                    i.jsx("option", { value: " ", children: "Space" }),
                    i.jsx("option", {
                      value: "custom",
                      children: "Custom character",
                    }),
                  ],
                }),
              ],
            }),
            separator === "custom" &&
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Custom delimiter" }),
                  i.jsx("input", {
                    "aria-label": text("Custom delimiter"),
                    value: custom,
                    onChange: (event) => {
                      (J(p, event.target.value), z([]), D(0));
                    },
                    lang: language("Custom delimiter"),
                    "data-ui-label": "Custom delimiter",
                  }),
                ],
              }),
            i.jsxs("label", {
              className: "ttc-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: m,
                  onChange: (event) => {
                    (v(event.target.checked), X$1());
                  },
                }),
                i.jsx(LocalizedText, {
                  children: "First source row contains headings",
                }),
              ],
            }),
            e.allowHidden &&
              i.jsxs("label", {
                className: "ttc-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => {
                      (O(event.target.checked), X$1());
                    },
                  }),
                  i.jsx(LocalizedText, { children: "Include hidden cells" }),
                ],
              }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Output destination" }),
                i.jsxs("select", {
                  "aria-label": text("Output destination"),
                  value: destination,
                  onChange: (event) => J(b, event.target.value),
                  lang: language("Output destination"),
                  "data-ui-label": "Output destination",
                  children: [
                    i.jsx("option", {
                      value: "new-sheet",
                      disabled: e.allowNewSheet === false,
                      children: "New worksheet",
                    }),
                    i.jsx("option", {
                      value: "range",
                      children: "Empty range",
                    }),
                  ],
                }),
              ],
            }),
            destination === "new-sheet"
              ? i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Output worksheet name" }),
                    i.jsx("input", {
                      "aria-label": text("Output worksheet name"),
                      value: k,
                      onChange: (event) => J(F, event.target.value),
                      lang: language("Output worksheet name"),
                      "data-ui-label": "Output worksheet name",
                    }),
                  ],
                })
              : i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Output worksheet" }),
                        i.jsxs("select", {
                          "aria-label": text("Output worksheet"),
                          value: te(y) ? y : "",
                          onChange: (event) => J(w, event.target.value),
                          lang: language("Output worksheet"),
                          "data-ui-label": "Output worksheet",
                          children: [
                            i.jsx("option", {
                              value: "",
                              disabled: true,
                              children: "Choose worksheet",
                            }),
                            U.map((ae) =>
                              i.jsx(
                                "option",
                                { value: ae.id, children: ae.name },
                                ae.id,
                              ),
                            ),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, {
                          children: "Output starting cell",
                        }),
                        i.jsx("input", {
                          "aria-label": text("Output starting cell"),
                          value: E,
                          placeholder: text("For example D1"),
                          onChange: (event) => J(A, event.target.value),
                          lang: language("Output starting cell"),
                          "data-ui-label": "Output starting cell",
                        }),
                      ],
                    }),
                  ],
                }),
          ],
        }),
        !!V &&
          i.jsxs("fieldset", {
            disabled: e.readOnly,
            className: "ttc-formats",
            children: [
              i.jsx("legend", {
                children: i.jsx(LocalizedText, {
                  children: "Column conversion",
                }),
              }),
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "Choose a type for each field, then preview again. Heading fields remain text.",
                }),
              }),
              Array.from({ length: V }, (ae, oe) =>
                i.jsxs(
                  "label",
                  {
                    children: [
                      i.jsx(LocalizedText, { children: "Field " }),
                      oe + 1,
                      i.jsxs("select", {
                        "aria-label": `Field ${oe + 1} conversion`,
                        value: formats[oe] ?? "text",
                        onChange: (event) => {
                          const he = Array.from(
                            { length: V },
                            (Se, le) => formats[le] ?? "text",
                          );
                          ((he[oe] = event.target.value), z(he), X$1());
                        },
                        children: [
                          i.jsx("option", { value: "text", children: "Text" }),
                          i.jsx("option", {
                            value: "general",
                            children: "General",
                          }),
                        ],
                      }),
                    ],
                  },
                  oe,
                ),
              ),
            ],
          }),
        i.jsx("button", {
          className: "ttc-primary",
          disabled: e.readOnly,
          onClick: onClick,
          children: i.jsx(LocalizedText, { children: "Preview split" }),
        }),
        e.readOnly &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children: "This workbook is read-only.",
            }),
          }),
        N?.key === G &&
          N.snapshot === JSON.stringify(e.workbook) &&
          i.jsx("p", {
            className: "ttc-error",
            role: "alert",
            children: N.message,
          }),
        B &&
          !H &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children:
                "The workbook, selection or permissions changed. Preview again.",
            }),
          }),
        Q &&
          i.jsxs("section", {
            className: "ttc-preview",
            "aria-label": text("Text to Columns preview"),
            lang: language("Text to Columns preview"),
            "data-ui-label": "Text to Columns preview",
            children: [
              i.jsxs("h3", {
                children: [
                  Q.height,
                  " ",
                  i.jsx(LocalizedText, { children: "source rows → " }),
                  Q.width,
                  " ",
                  i.jsx(LocalizedText, { children: "output columns" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  Q.sourceSheetName,
                  "!",
                  Q.request.sourceRange,
                  " →",
                  " ",
                  Q.targetSheetName,
                  "!",
                  Q.targetRange,
                  ". ",
                  Q.outputs.length,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "new cells. Source cells and annotations remain in place.",
                  }),
                ],
              }),
              !!Q.diagnostics.length &&
                i.jsxs("details", {
                  children: [
                    i.jsxs("summary", {
                      children: [
                        Q.diagnostics.length,
                        " ",
                        i.jsx(LocalizedText, { children: "split notices" }),
                      ],
                    }),
                    i.jsx("ul", {
                      children: Q.diagnostics.map((diagnostic, oe) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              diagnostic.sourceAddress,
                              ": ",
                              diagnostic.message,
                            ],
                          },
                          oe,
                        ),
                      ),
                    }),
                  ],
                }),
              i.jsx("div", {
                className: "ttc-scroll",
                children: i.jsxs("table", {
                  children: [
                    i.jsxs("caption", {
                      children: [
                        i.jsx(LocalizedText, { children: "Output cells " }),
                        C * Fv + 1,
                        "–",
                        Math.min((C + 1) * Fv, Q.outputs.length),
                        " ",
                        i.jsx(LocalizedText, { children: "of" }),
                        " ",
                        Q.outputs.length,
                      ],
                    }),
                    i.jsx("thead", {
                      children: i.jsxs("tr", {
                        children: [
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Source",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Field text",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Destination",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Result and type",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Evidence and metadata",
                            }),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: Q.outputs
                        .slice(C * Fv, (C + 1) * Fv)
                        .map((ae) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: ae.sourceAddress }),
                                i.jsxs("td", {
                                  children: [
                                    i.jsx("code", {
                                      children: stringifyJsonValue(ae.field),
                                    }),
                                    ae.padded &&
                                      i.jsxs("span", {
                                        children: [
                                          " ",
                                          i.jsx(LocalizedText, {
                                            children: "(padded)",
                                          }),
                                        ],
                                      }),
                                  ],
                                }),
                                i.jsx("td", { children: ae.address }),
                                i.jsxs("td", {
                                  children: [
                                    i.jsx("code", {
                                      children: stringifyJsonValue(ae.cell.raw),
                                    }),
                                    i.jsx("br", {}),
                                    ae.type,
                                    " · ",
                                    ae.format,
                                  ],
                                }),
                                i.jsx("td", {
                                  children: i.jsxs("details", {
                                    children: [
                                      i.jsx("summary", {
                                        children: i.jsx(LocalizedText, {
                                          children:
                                            "Raw source and output metadata",
                                        }),
                                      }),
                                      i.jsx("pre", {
                                        children: JSON.stringify(
                                          {
                                            source: Q.sources.find(
                                              (source) =>
                                                source.address ===
                                                ae.sourceAddress,
                                            ),
                                            output: ae.cell,
                                          },
                                          null,
                                          2,
                                        ),
                                      }),
                                    ],
                                  }),
                                }),
                              ],
                            },
                            ae.address,
                          ),
                        ),
                    }),
                  ],
                }),
              }),
              i.jsxs("div", {
                className: "ttc-pages",
                children: [
                  i.jsx("button", {
                    onClick: () => T(C - 1),
                    disabled: C === 0,
                    children: i.jsx(LocalizedText, {
                      children: "Previous output cells",
                    }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      C + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      Math.ceil(Q.outputs.length / Fv),
                    ],
                  }),
                  i.jsx("button", {
                    onClick: () => T(C + 1),
                    disabled: (C + 1) * Fv >= Q.outputs.length,
                    children: i.jsx(LocalizedText, {
                      children: "Next output cells",
                    }),
                  }),
                ],
              }),
              i.jsxs("footer", {
                children: [
                  i.jsx("button", {
                    onClick: () => _(null),
                    children: i.jsx(LocalizedText, {
                      children: "Cancel split preview",
                    }),
                  }),
                  i.jsx("button", {
                    className: "ttc-primary",
                    onClick: K,
                    disabled: e.readOnly,
                    children: i.jsx(LocalizedText, { children: "Apply split" }),
                  }),
                ],
              }),
            ],
          }),
        i.jsx("p", {
          className: "ttc-boundary",
          children: i.jsx(LocalizedText, {
            children:
              "One delimiter character; double quotes group fields and doubled quotes escape a quote. Empty fields are retained. Up to 500 output cells and 50 fields; fixed-width splitting and automatic date conversion are not included.",
          }),
        }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: e.onClose,
            children: i.jsx(LocalizedText, { children: "Close" }),
          }),
        }),
      ],
    }),
  });
}
const formatFormulaLiteral = (e) =>
  typeof e == "string"
    ? `"${e.replaceAll('"', '""')}"`
    : typeof e == "boolean"
      ? e
        ? "TRUE"
        : "FALSE"
      : String(e);
function buildTextCleanupFormula(steps, r) {
  let t = r;
  for (const n of validateTextCleanupSteps(steps)) {
    const o = n.args.map(formatFormulaLiteral);
    (["FIND", "SEARCH"].includes(n.function)
      ? o.splice(1, 0, t)
      : n.function === "TEXTJOIN"
        ? o.splice(2, 0, t)
        : o.unshift(t),
      (t = `${n.function}(${o.join(",")})`));
  }
  return `=${t}`;
}
function validateTextCleanupRequest(request) {
  if (
    !request ||
    typeof request != "object" ||
    Object.keys(request).some(
      (n) =>
        !["sourceSheetId", "sourceRange", "outputName", "steps"].includes(n),
    ) ||
    typeof request.sourceSheetId != "string" ||
    typeof request.sourceRange != "string" ||
    typeof request.outputName != "string"
  )
    throw Error(
      "Choose a source range, cleanup steps and output worksheet name.",
    );
  const rangeAddress = parseRangeAddress(request.sourceRange),
    outputName = request.outputName.trim();
  if (
    !outputName ||
    outputName.length > 31 ||
    /[\u0000-\u001f\u007f\\/\[\]*?:]/.test(outputName) ||
    outputName.startsWith("'") ||
    outputName.endsWith("'")
  )
    throw Error("Choose a valid output worksheet name of 1–31 characters.");
  return {
    sourceSheetId: request.sourceSheetId,
    sourceRange: `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`,
    outputName: outputName,
    steps: validateTextCleanupSteps(request.steps),
  };
}
function previewTextCleanup(workbook, request, t = {}) {
  const context = validateTextToColumnsPermissions(workbook, t),
    o = validateTextCleanupRequest(request);
  if (!context.allowNewSheet || workbook.structureProtection?.enabled)
    throw Error("Creating an output worksheet is not permitted.");
  const sheet = workbook.sheets.find((sheet) => sheet.id === o.sourceSheetId);
  if (!sheet) throw Error("Choose an existing source worksheet.");
  assertSelectionInAuthorizedSheet(sheet, o.sourceRange, context);
  const rangeAddress = parseRangeAddress(o.sourceRange),
    d = rangeAddress.lastRow - rangeAddress.firstRow + 1;
  if (rangeAddress.firstCol !== rangeAddress.lastCol || d > 500)
    throw Error("Select one source column containing at most 500 cells.");
  if (
    sheet.merges?.some((merge) => {
      const y = parseRangeAddress(merge);
      return (
        y.firstRow <= rangeAddress.lastRow &&
        y.lastRow >= rangeAddress.firstRow &&
        y.firstCol <= rangeAddress.lastCol &&
        y.lastCol >= rangeAddress.firstCol
      );
    })
  )
    throw Error("Unmerge the source before cleaning text.");
  if (
    workbook.sheets.some(
      (sheet) => sheet.name.toLowerCase() === o.outputName.toLowerCase(),
    )
  )
    throw Error("The output worksheet name is already used.");
  if (
    (assertOutputNameUnreferenced(workbook, o.outputName),
    discoverArrayFormulaAnchors(
      workbook,
      [{ sheet: sheet, range: o.sourceRange }],
      context,
    ),
    workbook.sheets.reduce((b, y) => b + Object.keys(y.cells).length, 0) + d >
      workbookLimits.maxCells)
  )
    throw Error("Cleanup output exceeds the 100,000 stored-cell budget.");
  const rows = [],
    formulas = [],
    errors = [],
    worksheetRecord = createWorksheetRecord("Text calculation");
  worksheetRecord.id = "text-calculation";
  let f = 0;
  for (let b = 0; b < d; b++) {
    const sourceAddress = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + b + 1}`,
      cell = sheet.cells[sourceAddress];
    if (
      cell &&
      typeof cell.raw == "string" &&
      cell.raw.startsWith("=") &&
      !cell.literal
    )
      throw Error(
        `Source ${sourceAddress} is a formula. Create a reviewed values snapshot before cleanup.`,
      );
    if ((f += String(cell?.raw ?? "").length) > 2e5)
      throw Error(
        "Source text exceeds the 200,000-character processing budget.",
      );
    worksheetRecord.cells[`A${b + 1}`] = {
      raw: cell?.raw ?? null,
      literal: true,
    };
    const textCleanupFormula = buildTextCleanupFormula(o.steps, `A${b + 1}`);
    if (textCleanupFormula.length > fileImportLimits.maxFormulaLength)
      throw Error(
        "The cleanup recipe exceeds the 4,096-character formula budget.",
      );
    (formulas.push(textCleanupFormula),
      (worksheetRecord.cells[`B${b + 1}`] = { raw: textCleanupFormula }));
    const rowOrigin = findRowOrigin(
      workbook,
      sheet.id,
      rangeAddress.firstRow + b,
    );
    rows.push({
      sourceAddress: sourceAddress,
      ...(cell ? { original: structuredClone(cell) } : {}),
      outputAddress: `A${b + 1}`,
      result: null,
      ...(rowOrigin ? { origin: rowOrigin } : {}),
    });
  }
  const workbookWithSettings = calculateWorkbookWithSettings({
    schemaVersion: 1,
    id: "text-calculation",
    name: "Text calculation",
    revision: 0,
    sheets: [worksheetRecord],
    events: [],
  });
  rows.forEach((row, y) => {
    ((row.result = workbookWithSettings(worksheetRecord.id, `B${y + 1}`)),
      workbookWithSettings.evaluateFormula(
        worksheetRecord.id,
        `C${y + 1}`,
        `=ISERROR(B${y + 1})`,
      ) === true &&
        errors.push({
          sourceAddress: row.sourceAddress,
          error: String(row.result),
        }));
  });
  const m = {
    snapshot: JSON.stringify(workbook),
    context: context,
    request: o,
    sourceSheetName: sheet.name,
    rows: rows,
    formulas: formulas,
    errors: errors,
  };
  return (
    normalizeTextCleanups(
      [
        ...(workbook.textCleanups ?? []),
        buildTextCleanupEvidence(
          m,
          workbook.id,
          workbook.revision,
          "preview-record",
          "preview-output",
          "2026-01-01T00:00:00.000Z",
        ),
      ],
      {},
    ),
    m
  );
}
function buildTextCleanupEvidence(
  e,
  workbookId,
  revision,
  n,
  outputSheetId,
  createdAt,
) {
  return {
    id: n,
    workbookId: workbookId,
    sourceRevision: revision,
    sourceSheetId: e.request.sourceSheetId,
    sourceSheetName: e.sourceSheetName,
    outputSheetId: outputSheetId,
    outputSheetName: e.request.outputName,
    createdAt: createdAt,
    steps: structuredClone(e.request.steps),
    semantics: "unicode-v2-en-US-explicit-separators",
    rows: structuredClone(e.rows),
  };
}
function createTextCleanupOutput(workbook, r, t = {}) {
  if (r.snapshot !== JSON.stringify(workbook))
    throw Error(
      "The workbook changed after the cleanup preview. Preview again.",
    );
  const n = previewTextCleanup(workbook, r.request, t);
  if (JSON.stringify(n) !== JSON.stringify(r))
    throw Error(
      "The cleanup preview or authorized context changed. Preview again.",
    );
  if (r.errors.length)
    throw Error(
      "Resolve every cleanup error before creating output. No cells were changed.",
    );
  const worksheetRecord = createWorksheetRecord(r.request.outputName),
    a = crypto.randomUUID();
  for (const row of r.rows) {
    const cell = {
      raw: row.result,
      note: `Text cleanup evidence: ${a}
Source at revision ${workbook.revision}: ${JSON.stringify(r.sourceSheetName)}!${row.sourceAddress}
Original value and recipe retained in native text cleanup history.`,
    };
    (typeof cell.raw == "string" &&
      ((cell.numFmt = "@"), cell.raw.startsWith("=") && (cell.literal = true)),
      (worksheetRecord.cells[row.outputAddress] = cell));
  }
  const textCleanupEvidence = buildTextCleanupEvidence(
    r,
    workbook.id,
    workbook.revision,
    a,
    worksheetRecord.id,
    new Date().toISOString(),
  );
  return normalizeWorkbookInput({
    ...workbook,
    sheets: [...workbook.sheets, worksheetRecord],
    textCleanups: [...(workbook.textCleanups ?? []), textCleanupEvidence],
  });
}
const createTextFunctionField = (e, r = "", t, optional = false) => ({
    label: e,
    value: r,
    kind: t,
    optional: optional,
  }),
  textFunctionFieldSpecs = {
    LEFT: [createTextFunctionField("Characters", "1", "number")],
    RIGHT: [createTextFunctionField("Characters", "1", "number")],
    MID: [
      createTextFunctionField("Start position", "1", "number"),
      createTextFunctionField("Characters", "1", "number"),
    ],
    FIND: [
      createTextFunctionField("Find text"),
      createTextFunctionField("Start position", "1", "number"),
    ],
    SEARCH: [
      createTextFunctionField("Find text"),
      createTextFunctionField("Start position", "1", "number"),
    ],
    SUBSTITUTE: [
      createTextFunctionField("Old text"),
      createTextFunctionField("New text"),
      createTextFunctionField("Occurrence (blank = all)", "", "number", true),
    ],
    REPLACE: [
      createTextFunctionField("Start position", "1", "number"),
      createTextFunctionField("Characters", "1", "number"),
      createTextFunctionField("New text"),
    ],
    NUMBERVALUE: [
      createTextFunctionField("Decimal separator", "."),
      createTextFunctionField("Group separator", ","),
    ],
    TEXT: [createTextFunctionField("Format", "00000")],
    CONCAT: [createTextFunctionField("Append text")],
    TEXTJOIN: [
      createTextFunctionField("Delimiter", "-"),
      createTextFunctionField("Ignore empty", "TRUE", "boolean"),
      createTextFunctionField("Append text"),
    ],
    EXACT: [createTextFunctionField("Compare with")],
    TEXTBEFORE: [
      createTextFunctionField("Delimiter", "-"),
      createTextFunctionField("Occurrence", "1", "number"),
      createTextFunctionField("Ignore case", "0", "number"),
      createTextFunctionField("Match end", "0", "number"),
    ],
    TEXTAFTER: [
      createTextFunctionField("Delimiter", "-"),
      createTextFunctionField("Occurrence", "1", "number"),
      createTextFunctionField("Ignore case", "0", "number"),
      createTextFunctionField("Match end", "0", "number"),
    ],
  },
  getTextFunctionSpec = (e) => ({
    name: e,
    fields: (textFunctionFieldSpecs[e] ?? []).map((r) => ({ ...r })),
  }),
  stringifyEscapedJson = (e) =>
    JSON.stringify(e).replaceAll(" ", "\\u00a0").replaceAll(" ", "\\u202f");

export {
  createTextFunctionField,
  previewTextCleanup,
  CombinePopulationsDialog,
  getTextFunctionSpec,
  createTextCleanupOutput,
  TextToColumnsDialog,
  stringifyEscapedJson,
};
export { stringifyJsonValue };
