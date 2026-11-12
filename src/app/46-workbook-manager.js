import { X, ArrowLeft, ArrowRight, ArrowDown, ArrowUp } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  updateWorkbookProperties,
  createWorkbookFromTemplate,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  worksheetPropertyLabels,
  formatChangeAttribution,
  formatAuditCellValue,
  formatAuditDetailValue,
} from "./45-sorting-and-worksheets.js";

function WorkbookManagerDialog(props) {
  const { text, language } = useWorkbookMessages(),
    dialogRef = ne.useRef(null),
    onCloseRef = ne.useRef(props.onClose);
  onCloseRef.current = props.onClose;
  const [workbookName, setWorkbookName] = ne.useState(
      props.mode === "template"
        ? `${props.workbook.name.slice(0, 185)} — copy`
        : props.workbook.name,
    ),
    [engagement, setEngagement] = ne.useState(
      props.workbook.metadata?.engagement ?? "",
    ),
    [period, setPeriod] = ne.useState(props.workbook.metadata?.period ?? ""),
    [checkpointName, setCheckpointName] = ne.useState(""),
    [selectedCheckpointId, setSelectedCheckpointId] = ne.useState(""),
    [comparison, setComparison] = ne.useState(null),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [noticeMessage, setNoticeMessage] = ne.useState(""),
    [isSaving, setIsSaving] = ne.useState(false),
    isMountedRef = ne.useRef(true);
  ne.useEffect(
    () => (
      (isMountedRef.current = true),
      () => {
        isMountedRef.current = false;
      }
    ),
    [],
  );
  const [checked, setChecked] = ne.useState(false),
    [pageIndex, setPageIndex] = ne.useState(0),
    isProtected = !!(
      props.workbook.structureProtection?.enabled ||
      props.workbook.sheets.some((sheet) => sheet.protection?.enabled)
    );
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("input, button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      (setComparison(null), setChecked(false), setPageIndex(0));
    }, [props.workbook.revision, props.workbook.id]));
  const sheetNameOf = (sheetId) =>
      props.workbook.sheets.find((sheet) => sheet.id === sheetId)?.name ??
      comparison?.structuralChanges.find(
        (structuralChange) => structuralChange.path === `sheets.${sheetId}`,
      )?.before?.name ??
      "Removed worksheet",
    changeRows = comparison
      ? [
          ...comparison.cellChanges.map((cellChange) => ({
            label: `${sheetNameOf(cellChange.sheetId)}!${cellChange.address}`,
            before: formatAuditCellValue(cellChange.before),
            after: formatAuditCellValue(cellChange.after),
            attribution: cellChange.attribution,
          })),
          ...[
            ...comparison.structuralChanges,
            ...comparison.metadataChanges,
          ].map((change) => {
            const pathParts = change.path.split(".");
            return {
              label:
                change.path === "sheets.order"
                  ? "Worksheet order"
                  : pathParts[0] === "sheets"
                    ? `${sheetNameOf(pathParts[1])}${pathParts[2] ? ` · ${worksheetPropertyLabels[pathParts[2]] ?? pathParts[2]}` : ""}`
                    : (worksheetPropertyLabels[change.path] ?? change.path),
              before:
                change.path === "sheets.order" && Array.isArray(change.before)
                  ? change.before.map(sheetNameOf).join(" → ")
                  : formatAuditDetailValue(change.before),
              after:
                change.path === "sheets.order" && Array.isArray(change.after)
                  ? change.after.map(sheetNameOf).join(" → ")
                  : formatAuditDetailValue(change.after),
              attribution: change.attribution,
            };
          }),
        ]
      : [],
    runOperation = (operation) => {
      try {
        (setErrorMessage(""), setNoticeMessage(""));
        const result = operation();
        result &&
          typeof result.then == "function" &&
          (setIsSaving(true),
          result
            .catch((failure) => {
              isMountedRef.current &&
                setErrorMessage(
                  failure instanceof Error
                    ? failure.message
                    : "The workbook operation could not be completed.",
                );
            })
            .finally(() => {
              isMountedRef.current && setIsSaving(false);
            }));
      } catch (error) {
        setErrorMessage(
          error instanceof Error
            ? error.message
            : "The workbook operation could not be completed.",
        );
      }
    };
  return i.jsx("div", {
    className: "wm-backdrop",
    children: i.jsxs("div", {
      ref: dialogRef,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": "wm-title",
      className: "wm-dialog",
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.preventDefault(),
            event.stopPropagation(),
            onCloseRef.current()),
          event.key === "Tab")
        ) {
          const focusableElements = Array.from(
              dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ),
            firstElement = focusableElements[0],
            lastElement = focusableElements.at(-1);
          event.shiftKey && document.activeElement === firstElement
            ? (event.preventDefault(),
              lastElement == null || lastElement.focus())
            : !event.shiftKey &&
              document.activeElement === lastElement &&
              (event.preventDefault(),
              firstElement == null || firstElement.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: "wm-title",
              children:
                props.mode === "versions"
                  ? "Version history"
                  : props.mode === "template"
                    ? "Create from this template"
                    : "Workbook properties",
            }),
            i.jsx("button", {
              "aria-label": text("Close workbook manager"),
              onClick: props.onClose,
              lang: language("Close workbook manager"),
              "data-ui-label": "Close workbook manager",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "wm-body",
          children: [
            props.mode !== "versions"
              ? i.jsxs("form", {
                  onSubmit: (event) => {
                    (event.preventDefault(),
                      runOperation(() => {
                        if (props.readOnly)
                          throw new Error("This workbook is read-only.");
                        props.mode === "template"
                          ? (props.onCreateWorkbook(
                              createWorkbookFromTemplate(props.workbook, {
                                name: workbookName,
                                engagement: engagement,
                                period: period,
                              }),
                            ),
                            props.onClose())
                          : (props.onWorkbookChange(
                              updateWorkbookProperties(props.workbook, {
                                name: workbookName,
                                engagement: engagement,
                                period: period,
                              }),
                              "Update workbook properties",
                            ),
                            setNoticeMessage("Workbook properties saved."));
                      }));
                  },
                  children: [
                    props.mode === "template" &&
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Download a new workbook with the current sheets, formulas and supported layout. Its identity and history will be independent, with a record of this template and its revision. The current workbook stays open.",
                        }),
                      }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Workbook name" }),
                        i.jsx("input", {
                          "aria-label": text("Workbook name"),
                          required: true,
                          maxLength: 200,
                          value: workbookName,
                          disabled: props.readOnly || isSaving,
                          onChange: (event) =>
                            setWorkbookName(event.target.value),
                          lang: language("Workbook name"),
                          "data-ui-label": "Workbook name",
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "wm-fields",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Engagement" }),
                            props.mode === "template" ? " *" : "",
                            i.jsx("input", {
                              "aria-label": text("Engagement"),
                              required: props.mode === "template",
                              value: engagement,
                              maxLength: 200,
                              disabled: props.readOnly || isSaving,
                              onChange: (event) =>
                                setEngagement(event.target.value),
                              lang: language("Engagement"),
                              "data-ui-label": "Engagement",
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Period" }),
                            props.mode === "template" ? " *" : "",
                            i.jsx("input", {
                              "aria-label": text("Period"),
                              required: props.mode === "template",
                              value: period,
                              maxLength: 100,
                              placeholder: text("For example, FY2026"),
                              disabled: props.readOnly || isSaving,
                              onChange: (event) =>
                                setPeriod(event.target.value),
                              lang: language("Period"),
                              "data-ui-label": "Period",
                            }),
                          ],
                        }),
                      ],
                    }),
                    props.mode === "properties" &&
                      i.jsxs("dl", {
                        children: [
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Created",
                            }),
                          }),
                          i.jsx("dd", {
                            children: props.workbook.metadata?.createdAt
                              ? new Date(
                                  props.workbook.metadata.createdAt,
                                ).toLocaleString()
                              : "Not recorded in this file",
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Modified",
                            }),
                          }),
                          i.jsx("dd", {
                            children: props.workbook.metadata?.modifiedAt
                              ? new Date(
                                  props.workbook.metadata.modifiedAt,
                                ).toLocaleString()
                              : "Not recorded in this file",
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Revision",
                            }),
                          }),
                          i.jsx("dd", { children: props.workbook.revision }),
                          props.workbook.metadata?.template &&
                            i.jsxs(i.Fragment, {
                              children: [
                                i.jsx("dt", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Template",
                                  }),
                                }),
                                i.jsxs("dd", {
                                  children: [
                                    props.workbook.metadata.template.name,
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "· version",
                                    }),
                                    " ",
                                    props.workbook.metadata.template.version,
                                    props.workbook.metadata.template
                                      .sourceRevision !== undefined
                                      ? ` · source revision ${props.workbook.metadata.template.sourceRevision}`
                                      : "",
                                  ],
                                }),
                              ],
                            }),
                        ],
                      }),
                    i.jsx("button", {
                      type: "submit",
                      className: "wm-primary",
                      disabled: props.readOnly || isSaving,
                      children:
                        props.mode === "template"
                          ? "Create independent workbook"
                          : "Save properties",
                    }),
                  ],
                })
              : i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("p", {
                      children: [
                        props.persistent
                          ? "Named checkpoints are saved in this browser with this workbook."
                          : "Named checkpoints are kept for this editor session. The embedding application can provide a persistent version store.",
                        " ",
                        i.jsx(LocalizedText, {
                          children:
                            "Restoring records a new revision and keeps intervening change history.",
                        }),
                      ],
                    }),
                    i.jsxs("form", {
                      className: "wm-checkpoint",
                      onSubmit: (event) => {
                        (event.preventDefault(),
                          runOperation(() => {
                            if (props.readOnly)
                              throw new Error("This workbook is read-only.");
                            const saveResult =
                                props.onCheckpoint(checkpointName),
                              onSaveComplete = () => {
                                isMountedRef.current &&
                                  (setCheckpointName(""),
                                  setNoticeMessage("Checkpoint saved."));
                              };
                            return saveResult
                              ? saveResult.then(onSaveComplete)
                              : onSaveComplete();
                          }));
                      },
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Checkpoint name",
                            }),
                            i.jsx("input", {
                              "aria-label": text("Checkpoint name"),
                              value: checkpointName,
                              required: true,
                              maxLength: 100,
                              disabled: props.readOnly || isSaving,
                              onChange: (event) =>
                                setCheckpointName(event.target.value),
                              lang: language("Checkpoint name"),
                              "data-ui-label": "Checkpoint name",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          disabled: props.readOnly || isSaving,
                          type: "submit",
                          className: "wm-primary",
                          children: i.jsx(LocalizedText, {
                            children: "Save checkpoint",
                          }),
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      className: "wm-muted",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Up to 20 checkpoints are retained. Saving another checkpoint replaces the oldest checkpoint when the limit is reached.",
                      }),
                    }),
                    props.checkpoints.length
                      ? i.jsx("div", {
                          className: "wm-versions",
                          "aria-label": text("Saved checkpoints"),
                          lang: language("Saved checkpoints"),
                          "data-ui-label": "Saved checkpoints",
                          children: [...props.checkpoints]
                            .reverse()
                            .map((checkpoint) =>
                              i.jsxs(
                                "button",
                                {
                                  "aria-pressed":
                                    selectedCheckpointId === checkpoint.id,
                                  onClick: () =>
                                    runOperation(() => {
                                      (setSelectedCheckpointId(checkpoint.id),
                                        setChecked(false),
                                        setPageIndex(0),
                                        setComparison(
                                          props.onCompare(checkpoint.id),
                                        ));
                                    }),
                                  children: [
                                    i.jsx("strong", {
                                      children: checkpoint.label,
                                    }),
                                    i.jsxs("span", {
                                      children: [
                                        i.jsx(LocalizedText, {
                                          children: "Revision ",
                                        }),
                                        checkpoint.revision,
                                        " ·",
                                        " ",
                                        new Date(
                                          checkpoint.createdAt,
                                        ).toLocaleString(),
                                      ],
                                    }),
                                  ],
                                },
                                checkpoint.id,
                              ),
                            ),
                        })
                      : i.jsx("p", {
                          children: i.jsx(LocalizedText, {
                            children: "No checkpoints have been saved.",
                          }),
                        }),
                    comparison &&
                      i.jsxs("section", {
                        className: "wm-comparison",
                        "aria-label": text("Version comparison"),
                        lang: language("Version comparison"),
                        "data-ui-label": "Version comparison",
                        children: [
                          i.jsx("h3", {
                            children: i.jsx(LocalizedText, {
                              children: "Changes since the selected checkpoint",
                            }),
                          }),
                          i.jsxs("p", {
                            children: [
                              comparison.cellChanges.length,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "cell changes ·",
                              }),
                              " ",
                              comparison.structuralChanges.length,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "structure changes ·",
                              }),
                              " ",
                              comparison.metadataChanges.length,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "property changes",
                              }),
                            ],
                          }),
                          i.jsx("div", {
                            className: "wm-diffs",
                            children: changeRows
                              .slice(pageIndex * 100, (pageIndex + 1) * 100)
                              .map((row, rowIndex) =>
                                i.jsxs(
                                  "div",
                                  {
                                    children: [
                                      i.jsx("strong", { children: row.label }),
                                      i.jsxs("span", {
                                        children: [
                                          row.before,
                                          " → ",
                                          row.after,
                                        ],
                                      }),
                                      i.jsx("small", {
                                        children: formatChangeAttribution(
                                          row.attribution,
                                        ),
                                      }),
                                    ],
                                  },
                                  `${pageIndex}-${rowIndex}`,
                                ),
                              ),
                          }),
                          changeRows.length > 100 &&
                            i.jsxs("div", {
                              className: "wm-pagination",
                              children: [
                                i.jsx("button", {
                                  disabled: pageIndex === 0,
                                  onClick: () =>
                                    setPageIndex(
                                      (currentPageIndex) =>
                                        currentPageIndex - 1,
                                    ),
                                  children: i.jsx(LocalizedText, {
                                    children: "Previous changes",
                                  }),
                                }),
                                i.jsxs("span", {
                                  children: [
                                    i.jsx(LocalizedText, { children: "Page " }),
                                    pageIndex + 1,
                                    " ",
                                    i.jsx(LocalizedText, { children: "of " }),
                                    Math.ceil(changeRows.length / 100),
                                  ],
                                }),
                                i.jsx("button", {
                                  disabled:
                                    (pageIndex + 1) * 100 >= changeRows.length,
                                  onClick: () =>
                                    setPageIndex(
                                      (currentPageIndex) =>
                                        currentPageIndex + 1,
                                    ),
                                  children: i.jsx(LocalizedText, {
                                    children: "Next changes",
                                  }),
                                }),
                              ],
                            }),
                          isProtected &&
                            i.jsx("p", {
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Unprotect the workbook and its worksheets before restoring a checkpoint.",
                              }),
                            }),
                          i.jsxs("label", {
                            className: "wm-confirm",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: checked,
                                disabled:
                                  props.readOnly || isProtected || isSaving,
                                onChange: (event) =>
                                  setChecked(event.target.checked),
                              }),
                              i.jsx(LocalizedText, {
                                children:
                                  "Restore this checkpoint as a new revision.",
                              }),
                            ],
                          }),
                          i.jsx("button", {
                            className: "wm-primary",
                            disabled:
                              props.readOnly ||
                              isProtected ||
                              !checked ||
                              isSaving,
                            onClick: () =>
                              runOperation(() => {
                                const restoreResult =
                                    props.onRestore(selectedCheckpointId),
                                  onRestoreComplete = () => {
                                    isMountedRef.current &&
                                      (setComparison(null),
                                      setChecked(false),
                                      setNoticeMessage(
                                        "Checkpoint restored as a new revision.",
                                      ));
                                  };
                                return restoreResult
                                  ? restoreResult.then(onRestoreComplete)
                                  : onRestoreComplete();
                              }),
                            children: i.jsx(LocalizedText, {
                              children: "Restore checkpoint",
                            }),
                          }),
                        ],
                      }),
                  ],
                }),
            isSaving &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children: "Saving workbook changes…",
                }),
              }),
            errorMessage &&
              i.jsx("p", {
                role: "alert",
                className: "wm-error",
                children: errorMessage,
              }),
            noticeMessage &&
              i.jsx("p", {
                role: "status",
                className: "wm-notice",
                children: noticeMessage,
              }),
          ],
        }),
      ],
    }),
  });
}
const fillDirectionOptions = [
    { value: "down", label: "Fill down", icon: ArrowDown },
    { value: "right", label: "Fill right", icon: ArrowRight },
    { value: "up", label: "Fill up", icon: ArrowUp },
    { value: "left", label: "Fill left", icon: ArrowLeft },
  ],
  describeWorkbookError = (error) =>
    error instanceof Error
      ? error.message
      : "This change could not be completed.";

export { fillDirectionOptions, describeWorkbookError, WorkbookManagerDialog };
