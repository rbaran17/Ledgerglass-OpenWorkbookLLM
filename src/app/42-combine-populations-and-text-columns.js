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

function CombinePopulationsDialog(props) {
  const { text, language } = useWorkbookMessages(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    reviewRef = ne.useRef(null),
    onCloseRef = ne.useRef(props.onClose);
  onCloseRef.current = props.onClose;
  const [sheetId, setSheetId] = ne.useState(props.sheetId),
    [rightSheetId, setRightSheetId] = ne.useState(props.sheetId),
    [range, setRange] = ne.useState(props.range),
    [rightRange, setRightRange] = ne.useState(""),
    [mode, setMode] = ne.useState("append"),
    [outputSheetName, setOutputSheetName] = ne.useState("Combined population"),
    [keys, setKeys] = ne.useState([{ left: "", right: "" }]),
    [normalization, setNormalization] = ne.useState({}),
    [checked, setChecked] = ne.useState(false),
    [fuzzyKeyIndex, setFuzzyKeyIndex] = ne.useState(0),
    [fuzzyThreshold, setFuzzyThreshold] = ne.useState("0.8"),
    [includeHidden, setIncludeHidden] = ne.useState(false),
    [preview, setPreview] = ne.useState(null),
    [previewSelectionKey, setPreviewSelectionKey] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [errorContextKey, setErrorContextKey] = ne.useState(""),
    [reviewView, setReviewView] = ne.useState("output"),
    [pageIndex, X$1] = ne.useState(0),
    sourceContext = {
      includeHidden: !!props.allowHidden && includeHidden,
      allowedSheetIds: props.allowedSheetIds,
      readOnly: props.readOnly,
    },
    allowedSheetIds = [
      ...new Set(
        props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    selectableSheets = props.workbook.sheets.filter(
      (sheet) =>
        allowedSheetIds.includes(sheet.id) &&
        (sourceContext.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    selectionKey = JSON.stringify([props.sheetId, props.range]),
    workbookSnapshot = JSON.stringify(props.workbook),
    inputsKey = JSON.stringify([
      workbookSnapshot,
      selectionKey,
      allowedSheetIds,
      sourceContext.includeHidden,
      props.readOnly,
    ]),
    isPreviewCurrent =
      !!preview &&
      !props.readOnly &&
      previewSelectionKey === selectionKey &&
      preview.snapshot === workbookSnapshot &&
      preview.context.includeHidden === sourceContext.includeHidden &&
      JSON.stringify(preview.context.allowedSheetIds) ===
        JSON.stringify(allowedSheetIds);
  (ne.useEffect(() => {
    var reviewElement;
    preview &&
      isPreviewCurrent &&
      ((reviewElement = reviewRef.current) == null
        ? undefined
        : reviewElement.scrollIntoView
      )?.call(reviewElement, {
        block: "start",
      });
  }, [preview, isPreviewCurrent]),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        dialogRef.current
          ?.querySelector('[data-ui-label="Combine operation"]')
          ?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  const onClick = () => {
      (setPreview(null), setErrorMessage(""), X$1(0));
    },
    updateInput = (setValue, nextValue) => {
      (setValue(nextValue), onClick());
    },
    handlePreviewError = (previewError) => {
      (setErrorContextKey(inputsKey),
        setErrorMessage(
          previewError instanceof Error
            ? previewError.message
            : "The populations could not be combined.",
        ));
    },
    buildPreview = () => {
      try {
        (setPreview(
          buildConsolidationOutput(
            props.workbook,
            {
              left: { sheetId: sheetId, range: range },
              right: { sheetId: rightSheetId, range: rightRange },
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
                            key: fuzzyKeyIndex,
                            threshold: fuzzyThreshold.trim()
                              ? Number(fuzzyThreshold)
                              : NaN,
                          },
                        }
                      : {}),
                  }),
              outputSheetName: outputSheetName,
            },
            sourceContext,
          ),
        ),
          setPreviewSelectionKey(selectionKey),
          setErrorMessage(""),
          X$1(0),
          setReviewView("output"));
      } catch (error) {
        (setPreview(null), handlePreviewError(error));
      }
    },
    applyCombination = () => {
      if (!(!preview || !isPreviewCurrent))
        try {
          const nextWorkbook = assertConsolidationPreviewCurrent(
            props.workbook,
            preview,
            sourceContext,
          );
          (props.onWorkbookChange(
            nextWorkbook,
            `Combine populations: ${preview.request.mode} → ${preview.request.outputSheetName}`,
            preview.outputSheetId,
          ),
            setPreview(null),
            props.onClose());
        } catch (error) {
          handlePreviewError(error);
        }
    },
    tableHeaders = isPreviewCurrent
      ? reviewView === "left"
        ? ["Source row", "Match outcome", ...preview.left.headers]
        : reviewView === "right"
          ? ["Source row", "Match outcome", ...preview.right.headers]
          : reviewView === "matches"
            ? ["Left source", "Right source", "Similarity", "Explanation"]
            : reviewView === "normalization"
              ? ["Side", "Source", "Original keys", "Normalized keys"]
              : reviewView === "exceptions"
                ? ["Side", "Source", "Exception"]
                : preview.headers
      : [],
    tableRows = isPreviewCurrent
      ? reviewView === "left"
        ? preview.unmatchedLeft.map((leftRecord) => [
            leftRecord.locator,
            leftRecord.unmatchedReason ?? "",
            ...leftRecord.values,
          ])
        : reviewView === "right"
          ? preview.unmatchedRight.map((rightRecord) => [
              rightRecord.locator,
              rightRecord.unmatchedReason ?? "",
              ...rightRecord.values,
            ])
          : reviewView === "matches"
            ? (preview.matching?.matches ?? []).map((match) => [
                match.left,
                match.right,
                match.score,
                match.explanation,
              ])
            : reviewView === "normalization"
              ? (preview.matching?.normalizedKeys ?? []).map(
                  (normalizedKey) => [
                    normalizedKey.side,
                    normalizedKey.locator,
                    JSON.stringify(normalizedKey.original),
                    JSON.stringify(normalizedKey.normalized),
                  ],
                )
              : reviewView === "exceptions"
                ? (preview.matching?.exceptions ?? []).map((exception) => [
                    exception.side,
                    exception.locator,
                    exception.reason,
                  ])
                : preview.rows.map((row) => [
                    ...row.values,
                    row.left,
                    row.right,
                  ])
      : [],
    pageCount = Math.max(1, Math.ceil(tableRows.length / w2)),
    currentPage = Math.min(pageIndex, pageCount - 1);
  return i.jsx("div", {
    className: "pc-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "pc-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(), onCloseRef.current()),
          event.key === "Tab")
        ) {
          const focusableElements = [
            ...(dialogRef.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled)",
            ) ?? []),
          ].filter((element) => !element.hidden);
          event.shiftKey && document.activeElement === focusableElements[0]
            ? (event.preventDefault(),
              (lastFocusable = focusableElements.at(-1)) == null ||
                lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === focusableElements.at(-1) &&
              (event.preventDefault(),
              (firstFocusable = focusableElements[0]) == null ||
                firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${dialogId}-title`,
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
              onClick: props.onClose,
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
                      disabled: props.readOnly,
                      value: mode,
                      onChange: (event) => {
                        (setMode(event.target.value), onClick());
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
                      disabled: props.readOnly,
                      onChange: (event) =>
                        updateInput(setOutputSheetName, event.target.value),
                      lang: language("Combined worksheet name"),
                      "data-ui-label": "Combined worksheet name",
                    }),
                  ],
                }),
              ],
            }),
            i.jsx("div", {
              className: "pc-sources",
              children: ["left", "right"].map((side) => {
                const sourceSheetId = side === "left" ? sheetId : rightSheetId,
                  sourceRange = side === "left" ? range : rightRange;
                return i.jsxs(
                  "fieldset",
                  {
                    disabled: props.readOnly,
                    children: [
                      i.jsxs("legend", {
                        children: [
                          side === "left" ? "Left" : "Right",
                          " ",
                          i.jsx(LocalizedText, { children: "source" }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Worksheet" }),
                          i.jsxs("select", {
                            "aria-label": `${side === "left" ? "Left" : "Right"} source worksheet`,
                            value: selectableSheets.some(
                              (entry) => entry.id === sourceSheetId,
                            )
                              ? sourceSheetId
                              : "",
                            onChange: (event) =>
                              updateInput(
                                side === "left" ? setSheetId : setRightSheetId,
                                event.target.value,
                              ),
                            children: [
                              i.jsx("option", {
                                value: "",
                                children: "Choose an allowed worksheet",
                              }),
                              selectableSheets.map((worksheet) =>
                                i.jsx(
                                  "option",
                                  {
                                    value: worksheet.id,
                                    children: worksheet.name,
                                  },
                                  worksheet.id,
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
                            "aria-label": `${side === "left" ? "Left" : "Right"} source range`,
                            value: sourceRange,
                            placeholder: text("A1:D20"),
                            onChange: (event) =>
                              updateInput(
                                side === "left" ? setRange : setRightRange,
                                event.target.value,
                              ),
                            lang: language("A1:D20"),
                          }),
                        ],
                      }),
                    ],
                  },
                  side,
                );
              }),
            }),
            mode !== "append" &&
              i.jsxs("fieldset", {
                className: "pc-keys",
                disabled: props.readOnly,
                children: [
                  i.jsx("legend", {
                    children: i.jsx(LocalizedText, {
                      children: "Matching headers in priority order",
                    }),
                  }),
                  keys.map((key, keyIndex) =>
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
                              keyIndex + 1,
                              i.jsx("input", {
                                "aria-label": `Left key ${keyIndex + 1}`,
                                value: key.left,
                                onChange: (event) => {
                                  (setKeys(
                                    keys.map((key, pairIndex) =>
                                      keyIndex === pairIndex
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
                              keyIndex + 1,
                              i.jsx("input", {
                                "aria-label": `Right key ${keyIndex + 1}`,
                                value: key.right,
                                onChange: (event) => {
                                  (setKeys(
                                    keys.map((key, keyPairIndex) =>
                                      keyIndex === keyPairIndex
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
                            "aria-label": `Remove key ${keyIndex + 1}`,
                            onClick: () => {
                              (setKeys(
                                keys.filter(
                                  (key, position) => position !== keyIndex,
                                ),
                              ),
                                setFuzzyKeyIndex(0),
                                onClick());
                            },
                            children: i.jsx(LocalizedText, {
                              children: "Remove",
                            }),
                          }),
                        ],
                      },
                      keyIndex,
                    ),
                  ),
                  i.jsx("button", {
                    disabled: keys.length >= 8,
                    onClick: () => {
                      (setKeys([...keys, { left: "", right: "" }]), onClick());
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
                disabled: props.readOnly,
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
                      ].map(([normalizationKey, optionLabel]) =>
                        i.jsxs(
                          "label",
                          {
                            className: "pc-checkbox",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: !!normalization[normalizationKey],
                                onChange: (event) => {
                                  (setNormalization({
                                    ...normalization,
                                    [normalizationKey]: event.target.checked,
                                  }),
                                    onClick());
                                },
                              }),
                              optionLabel,
                            ],
                          },
                          normalizationKey,
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
                              (setNormalization({
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
                          (setChecked(event.target.checked), onClick());
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
                              value: fuzzyKeyIndex,
                              onChange: (event) => {
                                (setFuzzyKeyIndex(Number(event.target.value)),
                                  onClick());
                              },
                              lang: language("Fuzzy key pair"),
                              "data-ui-label": "Fuzzy key pair",
                              children: keys.map((key, fuzzyPairIndex) =>
                                i.jsxs(
                                  "option",
                                  {
                                    value: fuzzyPairIndex,
                                    children: [
                                      "Pair ",
                                      fuzzyPairIndex + 1,
                                      ": ",
                                      key.left || "Left",
                                      " ↔",
                                      " ",
                                      key.right || "Right",
                                    ],
                                  },
                                  fuzzyPairIndex,
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
                              value: fuzzyThreshold,
                              onValueChange: (nextThreshold) => {
                                (setFuzzyThreshold(nextThreshold), onClick());
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
            props.allowHidden &&
              i.jsxs("label", {
                className: "pc-checkbox",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    "aria-label": text("Include hidden source records"),
                    checked: includeHidden,
                    disabled: props.readOnly,
                    onChange: (event) => {
                      (setIncludeHidden(event.target.checked), onClick());
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
            props.readOnly &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "The host is read-only. Preview and Apply are unavailable.",
                }),
              }),
            errorMessage &&
              errorContextKey === inputsKey &&
              i.jsx("p", {
                className: "pc-error",
                role: "alert",
                children: errorMessage,
              }),
            i.jsxs("div", {
              className: "pc-actions",
              children: [
                i.jsx("button", {
                  disabled: props.readOnly,
                  onClick: buildPreview,
                  children: i.jsx(LocalizedText, {
                    children: "Preview combination",
                  }),
                }),
                preview &&
                  i.jsx("button", {
                    onClick: onClick,
                    children: i.jsx(LocalizedText, {
                      children: "Discard preview",
                    }),
                  }),
              ],
            }),
            preview &&
              !isPreviewCurrent &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "The workbook or host context changed. Preview details are withheld. Review a fresh preview.",
                }),
              }),
            isPreviewCurrent &&
              preview &&
              i.jsxs("section", {
                className: "pc-review",
                ref: reviewRef,
                "aria-label": text("Population combination preview"),
                lang: language("Population combination preview"),
                "data-ui-label": "Population combination preview",
                children: [
                  i.jsxs("h3", {
                    children: [
                      preview.request.outputSheetName,
                      " · ",
                      preview.outputRange,
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, { children: "Left: " }),
                      preview.counts.leftRows,
                      " ",
                      i.jsx(LocalizedText, { children: "records. Right:" }),
                      " ",
                      preview.counts.rightRows,
                      " ",
                      i.jsx(LocalizedText, { children: "records. Output:" }),
                      " ",
                      preview.counts.outputRows,
                      " ",
                      i.jsx(LocalizedText, { children: "records." }),
                    ],
                  }),
                  mode !== "append" &&
                    i.jsxs(i.Fragment, {
                      children: [
                        preview.keyProfiles.map((keyProfile, keyProfileIndex) =>
                          i.jsxs(
                            "p",
                            {
                              children: [
                                i.jsx(LocalizedText, { children: "Key " }),
                                keyProfileIndex + 1,
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
                            keyProfileIndex,
                          ),
                        ),
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Matched source rows: ",
                            }),
                            preview.counts.matchedLeftRows,
                            " ",
                            i.jsx(LocalizedText, { children: "left /" }),
                            " ",
                            preview.counts.matchedRightRows,
                            " ",
                            i.jsx(LocalizedText, {
                              children: "right. Match pairs:",
                            }),
                            " ",
                            preview.counts.matchPairs,
                            i.jsx(LocalizedText, { children: ". Unmatched:" }),
                            " ",
                            preview.unmatchedLeft.length,
                            " ",
                            i.jsx(LocalizedText, { children: "left /" }),
                            " ",
                            preview.unmatchedRight.length,
                            " ",
                            i.jsx(LocalizedText, { children: "right." }),
                          ],
                        }),
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Duplicate key groups: ",
                            }),
                            preview.counts.duplicateLeftKeys,
                            " ",
                            i.jsx(LocalizedText, { children: "left /" }),
                            " ",
                            preview.counts.duplicateRightKeys,
                            " ",
                            i.jsx(LocalizedText, {
                              children: "right. Many-to-many key groups: ",
                            }),
                            preview.counts.manyToManyKeys,
                            i.jsx(LocalizedText, {
                              children: ". Extra many-to-many pairs: ",
                            }),
                            preview.counts.expandedMatchRows,
                            ".",
                          ],
                        }),
                        preview.matching &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("p", { children: preview.matching.policy }),
                              i.jsxs("p", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Rows with multiple matches:",
                                  }),
                                  " ",
                                  preview.matching.ambiguousLeftRows,
                                  " ",
                                  i.jsx(LocalizedText, { children: "left /" }),
                                  " ",
                                  preview.matching.ambiguousRightRows,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children:
                                      "right. Normalization collision groups:",
                                  }),
                                  " ",
                                  preview.matching.normalizationCollisionGroups
                                    .left,
                                  " ",
                                  i.jsx(LocalizedText, { children: "left /" }),
                                  " ",
                                  preview.matching.normalizationCollisionGroups
                                    .right,
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
                        value: reviewView,
                        onChange: (event) => {
                          (setReviewView(event.target.value), X$1(0));
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
                                    preview.unmatchedLeft.length,
                                    ")",
                                  ],
                                }),
                                i.jsxs("option", {
                                  value: "right",
                                  children: [
                                    "All unmatched right records (",
                                    preview.unmatchedRight.length,
                                    ")",
                                  ],
                                }),
                                i.jsxs("option", {
                                  value: "matches",
                                  children: [
                                    "All match scores and explanations (",
                                    preview.matching?.matches.length ?? 0,
                                    ")",
                                  ],
                                }),
                                i.jsxs("option", {
                                  value: "normalization",
                                  children: [
                                    "Changed keys (",
                                    preview.matching?.normalizedKeys.length ??
                                      0,
                                    ")",
                                  ],
                                }),
                                i.jsxs("option", {
                                  value: "exceptions",
                                  children: [
                                    "Match exceptions (",
                                    preview.matching?.exceptions.length ?? 0,
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
                            children: tableHeaders.map((header, columnIndex) =>
                              i.jsx("th", { children: header }, columnIndex),
                            ),
                          }),
                        }),
                        i.jsx("tbody", {
                          children: tableRows
                            .slice(currentPage * w2, (currentPage + 1) * w2)
                            .map((tableRow, pageRowIndex) =>
                              i.jsx(
                                "tr",
                                {
                                  children: tableRow.map(
                                    (cellValue, cellIndex) =>
                                      i.jsx(
                                        "td",
                                        {
                                          title:
                                            cellValue === null
                                              ? "Blank"
                                              : typeof cellValue,
                                          children:
                                            formatCellValueForDisplay(
                                              cellValue,
                                            ),
                                        },
                                        cellIndex,
                                      ),
                                  ),
                                },
                                currentPage * w2 + pageRowIndex,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  !tableRows.length &&
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
                        disabled: !currentPage,
                        onClick: () => X$1(currentPage - 1),
                        children: i.jsx(LocalizedText, {
                          children: "Previous records",
                        }),
                      }),
                      i.jsxs("span", {
                        "aria-live": "polite",
                        children: [
                          i.jsx(LocalizedText, { children: "Page " }),
                          currentPage + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          pageCount,
                          " · ",
                          tableRows.length,
                          " ",
                          i.jsx(LocalizedText, { children: "records" }),
                        ],
                      }),
                      i.jsx("button", {
                        disabled: currentPage >= pageCount - 1,
                        onClick: () => X$1(currentPage + 1),
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
                      const reviewExport = {
                          schemaVersion: 1,
                          workbookId: preview.next.id,
                          revision: preview.next.revision,
                          request: preview.request,
                          context: preview.context,
                          left: preview.left,
                          right: preview.right,
                          counts: preview.counts,
                          headers: preview.headers,
                          rows: preview.rows,
                          unmatchedLeft: preview.unmatchedLeft,
                          unmatchedRight: preview.unmatchedRight,
                          keyProfiles: preview.keyProfiles,
                          matching: preview.matching,
                        },
                        objectUrl = URL.createObjectURL(
                          new Blob([JSON.stringify(reviewExport, null, 2)], {
                            type: "application/json",
                          }),
                        ),
                        downloadLink = document.createElement("a");
                      ((downloadLink.href = objectUrl),
                        (downloadLink.download =
                          "Population-combine-review.json"),
                        downloadLink.click(),
                        setTimeout(() => URL.revokeObjectURL(objectUrl), 5e3));
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
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              disabled: !isPreviewCurrent,
              onClick: applyCombination,
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
  stringifyJsonValue = (value) => JSON.stringify(value);
function TextToColumnsDialog(props) {
  const { text, language } = useWorkbookMessages(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    isAppliedRef = ne.useRef(false),
    [sourceSheetId, setSourceSheetId] = ne.useState(props.sheetId),
    [sourceRange, setSourceRange] = ne.useState(props.range),
    [separator, setSeparator] = ne.useState(","),
    [custom, setCustom] = ne.useState("|"),
    [hasHeader, setHasHeader] = ne.useState(false),
    [destination, setDestination] = ne.useState(
      props.allowNewSheet === false ? "range" : "new-sheet",
    ),
    [targetSheetId, setTargetSheetId] = ne.useState(props.sheetId),
    [targetAddress, setTargetAddress] = ne.useState(""),
    [outputSheetName, setOutputSheetName] = ne.useState("Split text"),
    [checked, setIncludeHidden] = ne.useState(false),
    [formats, setFormats] = ne.useState([]),
    [fieldCount, setFieldCount] = ne.useState(0),
    [pendingPreview, setPendingPreview] = ne.useState(null),
    [previewPage, setPreviewPage] = ne.useState(0),
    [previewError, setPreviewError] = ne.useState(null),
    allowedSheetIds = [
      ...new Set(
        props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    splitOptions = {
      readOnly: props.readOnly,
      allowedSheetIds: allowedSheetIds,
      includeHidden: !!props.allowHidden && checked,
      allowNewSheet: props.allowNewSheet ?? true,
    },
    visibleSheets = props.workbook.sheets.filter(
      (sheet) =>
        allowedSheetIds.includes(sheet.id) &&
        (splitOptions.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    formKey = JSON.stringify({
      sourceSheetId: sourceSheetId,
      sourceRange: sourceRange,
      separator: separator,
      custom: custom,
      header: hasHeader,
      destination: destination,
      targetSheetId: targetSheetId,
      targetAddress: targetAddress,
      name: outputSheetName,
      formats: formats,
      options: splitOptions,
      selection: [props.sheetId, props.range],
    }),
    isPreviewCurrent =
      pendingPreview?.key === formKey &&
      pendingPreview.plan.snapshot === JSON.stringify(props.workbook),
    currentPlan = isPreviewCurrent ? pendingPreview.plan : null;
  (ne.useEffect(() => {
    const activeElement = document.activeElement,
      sourceRangeInput = dialogRef.current?.querySelector(
        '[data-ui-label="Source range"]',
      );
    return (
      (sourceRangeInput && !sourceRangeInput.matches(":disabled")
        ? sourceRangeInput
        : dialogRef.current?.querySelector("button")
      )?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      (setSourceSheetId(props.sheetId),
        setSourceRange(props.range),
        setFormats([]),
        setFieldCount(0),
        setPendingPreview(null),
        setPreviewError(null));
    }, [props.sheetId, props.range]));
  const X$1 = () => {
      (setPendingPreview(null),
        setPreviewError(null),
        (isAppliedRef.current = false));
    },
    updateSetting = (setState, nextValue) => {
      (setState(nextValue), X$1());
    },
    recordPreviewError = (caughtError) =>
      setPreviewError({
        key: formKey,
        snapshot: JSON.stringify(props.workbook),
        message:
          caughtError instanceof Error
            ? caughtError.message
            : "Text could not be split. Review the source and settings.",
      }),
    onClick = () => {
      X$1();
      try {
        const request = {
            sourceSheetId: sourceSheetId,
            sourceRange: sourceRange,
            delimiter: separator === "custom" ? custom : separator,
            header: hasHeader ? "first-row" : "none",
            formats: formats,
            destination:
              destination === "new-sheet"
                ? { kind: "new-sheet", name: outputSheetName }
                : {
                    kind: "range",
                    sheetId: targetSheetId,
                    address: targetAddress,
                  },
          },
          textSplitResult = buildTextSplitResult(
            props.workbook,
            request,
            splitOptions,
          );
        (setFieldCount(textSplitResult.width),
          setPreviewPage(0),
          setPendingPreview({ plan: textSplitResult, key: formKey }));
      } catch (error) {
        recordPreviewError(error);
      }
    },
    applyPreview = () => {
      if (!(!currentPlan || isAppliedRef.current || props.readOnly))
        try {
          const nextWorkbook = applyTextSplitPreview(
            props.workbook,
            currentPlan,
            splitOptions,
          );
          ((isAppliedRef.current = true),
            props.onWorkbookChange(
              nextWorkbook,
              `Text to Columns: ${currentPlan.sourceSheetName}!${currentPlan.request.sourceRange} → ${currentPlan.targetSheetName}!${currentPlan.targetRange}`,
            ),
            setPendingPreview(null),
            setPreviewError(null));
        } catch (error) {
          recordPreviewError(error);
        }
    },
    isVisibleSheet = (sheetId) =>
      visibleSheets.some((visibleSheet) => visibleSheet.id === sheetId);
  return i.jsx("div", {
    className: "ttc-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      ref: dialogRef,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      className: "ttc-dialog",
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
            ...(dialogRef.current?.querySelectorAll(
              "button,input,select,summary",
            ) ?? []),
          ].filter(
            (element) =>
              !element.matches(":disabled") &&
              !element.hidden &&
              !element.closest("details:not([open]) > :not(summary)"),
          );
          event.shiftKey && document.activeElement === focusableElements[0]
            ? (event.preventDefault(),
              (lastFocusable = focusableElements.at(-1)) == null ||
                lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === focusableElements.at(-1) &&
              (event.preventDefault(),
              (firstFocusable = focusableElements[0]) == null ||
                firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${dialogId}-title`,
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
              onClick: props.onClose,
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
          disabled: props.readOnly,
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
                  value: isVisibleSheet(sourceSheetId) ? sourceSheetId : "",
                  onChange: (event) => {
                    (updateSetting(setSourceSheetId, event.target.value),
                      setFormats([]),
                      setFieldCount(0));
                  },
                  lang: language("Source worksheet"),
                  "data-ui-label": "Source worksheet",
                  children: [
                    i.jsx("option", {
                      value: "",
                      disabled: true,
                      children: "Choose worksheet",
                    }),
                    visibleSheets.map((sourceSheet) =>
                      i.jsx(
                        "option",
                        { value: sourceSheet.id, children: sourceSheet.name },
                        sourceSheet.id,
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
                    (updateSetting(setSourceRange, event.target.value),
                      setFormats([]),
                      setFieldCount(0));
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
                    (updateSetting(setSeparator, event.target.value),
                      setFormats([]),
                      setFieldCount(0));
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
                      (updateSetting(setCustom, event.target.value),
                        setFormats([]),
                        setFieldCount(0));
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
                  checked: hasHeader,
                  onChange: (event) => {
                    (setHasHeader(event.target.checked), X$1());
                  },
                }),
                i.jsx(LocalizedText, {
                  children: "First source row contains headings",
                }),
              ],
            }),
            props.allowHidden &&
              i.jsxs("label", {
                className: "ttc-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => {
                      (setIncludeHidden(event.target.checked), X$1());
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
                  onChange: (event) =>
                    updateSetting(setDestination, event.target.value),
                  lang: language("Output destination"),
                  "data-ui-label": "Output destination",
                  children: [
                    i.jsx("option", {
                      value: "new-sheet",
                      disabled: props.allowNewSheet === false,
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
                      value: outputSheetName,
                      onChange: (event) =>
                        updateSetting(setOutputSheetName, event.target.value),
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
                          value: isVisibleSheet(targetSheetId)
                            ? targetSheetId
                            : "",
                          onChange: (event) =>
                            updateSetting(setTargetSheetId, event.target.value),
                          lang: language("Output worksheet"),
                          "data-ui-label": "Output worksheet",
                          children: [
                            i.jsx("option", {
                              value: "",
                              disabled: true,
                              children: "Choose worksheet",
                            }),
                            visibleSheets.map((outputSheet) =>
                              i.jsx(
                                "option",
                                {
                                  value: outputSheet.id,
                                  children: outputSheet.name,
                                },
                                outputSheet.id,
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
                          value: targetAddress,
                          placeholder: text("For example D1"),
                          onChange: (event) =>
                            updateSetting(setTargetAddress, event.target.value),
                          lang: language("Output starting cell"),
                          "data-ui-label": "Output starting cell",
                        }),
                      ],
                    }),
                  ],
                }),
          ],
        }),
        !!fieldCount &&
          i.jsxs("fieldset", {
            disabled: props.readOnly,
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
              Array.from({ length: fieldCount }, (ae, fieldIndex) =>
                i.jsxs(
                  "label",
                  {
                    children: [
                      i.jsx(LocalizedText, { children: "Field " }),
                      fieldIndex + 1,
                      i.jsxs("select", {
                        "aria-label": `Field ${fieldIndex + 1} conversion`,
                        value: formats[fieldIndex] ?? "text",
                        onChange: (event) => {
                          const nextFormats = Array.from(
                            { length: fieldCount },
                            (Se, formatIndex) => formats[formatIndex] ?? "text",
                          );
                          ((nextFormats[fieldIndex] = event.target.value),
                            setFormats(nextFormats),
                            X$1());
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
                  fieldIndex,
                ),
              ),
            ],
          }),
        i.jsx("button", {
          className: "ttc-primary",
          disabled: props.readOnly,
          onClick: onClick,
          children: i.jsx(LocalizedText, { children: "Preview split" }),
        }),
        props.readOnly &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children: "This workbook is read-only.",
            }),
          }),
        previewError?.key === formKey &&
          previewError.snapshot === JSON.stringify(props.workbook) &&
          i.jsx("p", {
            className: "ttc-error",
            role: "alert",
            children: previewError.message,
          }),
        pendingPreview &&
          !isPreviewCurrent &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children:
                "The workbook, selection or permissions changed. Preview again.",
            }),
          }),
        currentPlan &&
          i.jsxs("section", {
            className: "ttc-preview",
            "aria-label": text("Text to Columns preview"),
            lang: language("Text to Columns preview"),
            "data-ui-label": "Text to Columns preview",
            children: [
              i.jsxs("h3", {
                children: [
                  currentPlan.height,
                  " ",
                  i.jsx(LocalizedText, { children: "source rows → " }),
                  currentPlan.width,
                  " ",
                  i.jsx(LocalizedText, { children: "output columns" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  currentPlan.sourceSheetName,
                  "!",
                  currentPlan.request.sourceRange,
                  " →",
                  " ",
                  currentPlan.targetSheetName,
                  "!",
                  currentPlan.targetRange,
                  ". ",
                  currentPlan.outputs.length,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "new cells. Source cells and annotations remain in place.",
                  }),
                ],
              }),
              !!currentPlan.diagnostics.length &&
                i.jsxs("details", {
                  children: [
                    i.jsxs("summary", {
                      children: [
                        currentPlan.diagnostics.length,
                        " ",
                        i.jsx(LocalizedText, { children: "split notices" }),
                      ],
                    }),
                    i.jsx("ul", {
                      children: currentPlan.diagnostics.map(
                        (diagnostic, diagnosticIndex) =>
                          i.jsxs(
                            "li",
                            {
                              children: [
                                diagnostic.sourceAddress,
                                ": ",
                                diagnostic.message,
                              ],
                            },
                            diagnosticIndex,
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
                        previewPage * Fv + 1,
                        "–",
                        Math.min(
                          (previewPage + 1) * Fv,
                          currentPlan.outputs.length,
                        ),
                        " ",
                        i.jsx(LocalizedText, { children: "of" }),
                        " ",
                        currentPlan.outputs.length,
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
                      children: currentPlan.outputs
                        .slice(previewPage * Fv, (previewPage + 1) * Fv)
                        .map((output) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: output.sourceAddress }),
                                i.jsxs("td", {
                                  children: [
                                    i.jsx("code", {
                                      children: stringifyJsonValue(
                                        output.field,
                                      ),
                                    }),
                                    output.padded &&
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
                                i.jsx("td", { children: output.address }),
                                i.jsxs("td", {
                                  children: [
                                    i.jsx("code", {
                                      children: stringifyJsonValue(
                                        output.cell.raw,
                                      ),
                                    }),
                                    i.jsx("br", {}),
                                    output.type,
                                    " · ",
                                    output.format,
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
                                            source: currentPlan.sources.find(
                                              (source) =>
                                                source.address ===
                                                output.sourceAddress,
                                            ),
                                            output: output.cell,
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
                            output.address,
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
                    onClick: () => setPreviewPage(previewPage - 1),
                    disabled: previewPage === 0,
                    children: i.jsx(LocalizedText, {
                      children: "Previous output cells",
                    }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      previewPage + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      Math.ceil(currentPlan.outputs.length / Fv),
                    ],
                  }),
                  i.jsx("button", {
                    onClick: () => setPreviewPage(previewPage + 1),
                    disabled:
                      (previewPage + 1) * Fv >= currentPlan.outputs.length,
                    children: i.jsx(LocalizedText, {
                      children: "Next output cells",
                    }),
                  }),
                ],
              }),
              i.jsxs("footer", {
                children: [
                  i.jsx("button", {
                    onClick: () => setPendingPreview(null),
                    children: i.jsx(LocalizedText, {
                      children: "Cancel split preview",
                    }),
                  }),
                  i.jsx("button", {
                    className: "ttc-primary",
                    onClick: applyPreview,
                    disabled: props.readOnly,
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
            onClick: props.onClose,
            children: i.jsx(LocalizedText, { children: "Close" }),
          }),
        }),
      ],
    }),
  });
}
const formatFormulaLiteral = (value) =>
  typeof value == "string"
    ? `"${value.replaceAll('"', '""')}"`
    : typeof value == "boolean"
      ? value
        ? "TRUE"
        : "FALSE"
      : String(value);
function buildTextCleanupFormula(steps, sourceReference) {
  let expression = sourceReference;
  for (const step of validateTextCleanupSteps(steps)) {
    const formattedArgs = step.args.map(formatFormulaLiteral);
    (["FIND", "SEARCH"].includes(step.function)
      ? formattedArgs.splice(1, 0, expression)
      : step.function === "TEXTJOIN"
        ? formattedArgs.splice(2, 0, expression)
        : formattedArgs.unshift(expression),
      (expression = `${step.function}(${formattedArgs.join(",")})`));
  }
  return `=${expression}`;
}
function validateTextCleanupRequest(request) {
  if (
    !request ||
    typeof request != "object" ||
    Object.keys(request).some(
      (key) =>
        !["sourceSheetId", "sourceRange", "outputName", "steps"].includes(key),
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
function previewTextCleanup(workbook, request, options = {}) {
  const context = validateTextToColumnsPermissions(workbook, options),
    cleanupRequest = validateTextCleanupRequest(request);
  if (!context.allowNewSheet || workbook.structureProtection?.enabled)
    throw Error("Creating an output worksheet is not permitted.");
  const sheet = workbook.sheets.find(
    (sheet) => sheet.id === cleanupRequest.sourceSheetId,
  );
  if (!sheet) throw Error("Choose an existing source worksheet.");
  assertSelectionInAuthorizedSheet(sheet, cleanupRequest.sourceRange, context);
  const rangeAddress = parseRangeAddress(cleanupRequest.sourceRange),
    rowCount = rangeAddress.lastRow - rangeAddress.firstRow + 1;
  if (rangeAddress.firstCol !== rangeAddress.lastCol || rowCount > 500)
    throw Error("Select one source column containing at most 500 cells.");
  if (
    sheet.merges?.some((merge) => {
      const mergeRange = parseRangeAddress(merge);
      return (
        mergeRange.firstRow <= rangeAddress.lastRow &&
        mergeRange.lastRow >= rangeAddress.firstRow &&
        mergeRange.firstCol <= rangeAddress.lastCol &&
        mergeRange.lastCol >= rangeAddress.firstCol
      );
    })
  )
    throw Error("Unmerge the source before cleaning text.");
  if (
    workbook.sheets.some(
      (sheet) =>
        sheet.name.toLowerCase() === cleanupRequest.outputName.toLowerCase(),
    )
  )
    throw Error("The output worksheet name is already used.");
  if (
    (assertOutputNameUnreferenced(workbook, cleanupRequest.outputName),
    discoverArrayFormulaAnchors(
      workbook,
      [{ sheet: sheet, range: cleanupRequest.sourceRange }],
      context,
    ),
    workbook.sheets.reduce(
      (total, worksheet) => total + Object.keys(worksheet.cells).length,
      0,
    ) +
      rowCount >
      workbookLimits.maxCells)
  )
    throw Error("Cleanup output exceeds the 100,000 stored-cell budget.");
  const rows = [],
    formulas = [],
    errors = [],
    worksheetRecord = createWorksheetRecord("Text calculation");
  worksheetRecord.id = "text-calculation";
  let sourceCharacterCount = 0;
  for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
    const sourceAddress = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + rowOffset + 1}`,
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
    if ((sourceCharacterCount += String(cell?.raw ?? "").length) > 2e5)
      throw Error(
        "Source text exceeds the 200,000-character processing budget.",
      );
    worksheetRecord.cells[`A${rowOffset + 1}`] = {
      raw: cell?.raw ?? null,
      literal: true,
    };
    const textCleanupFormula = buildTextCleanupFormula(
      cleanupRequest.steps,
      `A${rowOffset + 1}`,
    );
    if (textCleanupFormula.length > fileImportLimits.maxFormulaLength)
      throw Error(
        "The cleanup recipe exceeds the 4,096-character formula budget.",
      );
    (formulas.push(textCleanupFormula),
      (worksheetRecord.cells[`B${rowOffset + 1}`] = {
        raw: textCleanupFormula,
      }));
    const rowOrigin = findRowOrigin(
      workbook,
      sheet.id,
      rangeAddress.firstRow + rowOffset,
    );
    rows.push({
      sourceAddress: sourceAddress,
      ...(cell ? { original: structuredClone(cell) } : {}),
      outputAddress: `A${rowOffset + 1}`,
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
  rows.forEach((row, rowIndex) => {
    ((row.result = workbookWithSettings(
      worksheetRecord.id,
      `B${rowIndex + 1}`,
    )),
      workbookWithSettings.evaluateFormula(
        worksheetRecord.id,
        `C${rowIndex + 1}`,
        `=ISERROR(B${rowIndex + 1})`,
      ) === true &&
        errors.push({
          sourceAddress: row.sourceAddress,
          error: String(row.result),
        }));
  });
  const preview = {
    snapshot: JSON.stringify(workbook),
    context: context,
    request: cleanupRequest,
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
          preview,
          workbook.id,
          workbook.revision,
          "preview-record",
          "preview-output",
          "2026-01-01T00:00:00.000Z",
        ),
      ],
      {},
    ),
    preview
  );
}
function buildTextCleanupEvidence(
  preview,
  workbookId,
  revision,
  recordId,
  outputSheetId,
  createdAt,
) {
  return {
    id: recordId,
    workbookId: workbookId,
    sourceRevision: revision,
    sourceSheetId: preview.request.sourceSheetId,
    sourceSheetName: preview.sourceSheetName,
    outputSheetId: outputSheetId,
    outputSheetName: preview.request.outputName,
    createdAt: createdAt,
    steps: structuredClone(preview.request.steps),
    semantics: "unicode-v2-en-US-explicit-separators",
    rows: structuredClone(preview.rows),
  };
}
function createTextCleanupOutput(workbook, preview, options = {}) {
  if (preview.snapshot !== JSON.stringify(workbook))
    throw Error(
      "The workbook changed after the cleanup preview. Preview again.",
    );
  const replayedPreview = previewTextCleanup(
    workbook,
    preview.request,
    options,
  );
  if (JSON.stringify(replayedPreview) !== JSON.stringify(preview))
    throw Error(
      "The cleanup preview or authorized context changed. Preview again.",
    );
  if (preview.errors.length)
    throw Error(
      "Resolve every cleanup error before creating output. No cells were changed.",
    );
  const worksheetRecord = createWorksheetRecord(preview.request.outputName),
    evidenceId = crypto.randomUUID();
  for (const row of preview.rows) {
    const cell = {
      raw: row.result,
      note: `Text cleanup evidence: ${evidenceId}
Source at revision ${workbook.revision}: ${JSON.stringify(preview.sourceSheetName)}!${row.sourceAddress}
Original value and recipe retained in native text cleanup history.`,
    };
    (typeof cell.raw == "string" &&
      ((cell.numFmt = "@"), cell.raw.startsWith("=") && (cell.literal = true)),
      (worksheetRecord.cells[row.outputAddress] = cell));
  }
  const textCleanupEvidence = buildTextCleanupEvidence(
    preview,
    workbook.id,
    workbook.revision,
    evidenceId,
    worksheetRecord.id,
    new Date().toISOString(),
  );
  return normalizeWorkbookInput({
    ...workbook,
    sheets: [...workbook.sheets, worksheetRecord],
    textCleanups: [...(workbook.textCleanups ?? []), textCleanupEvidence],
  });
}
const createTextFunctionField = (
    label,
    defaultValue = "",
    kind,
    optional = false,
  ) => ({
    label: label,
    value: defaultValue,
    kind: kind,
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
  getTextFunctionSpec = (functionName) => ({
    name: functionName,
    fields: (textFunctionFieldSpecs[functionName] ?? []).map((field) => ({
      ...field,
    })),
  }),
  stringifyEscapedJson = (value) =>
    JSON.stringify(value).replaceAll(" ", "\\u00a0").replaceAll(" ", "\\u202f");

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
