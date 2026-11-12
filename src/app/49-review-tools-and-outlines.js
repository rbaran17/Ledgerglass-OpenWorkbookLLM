import {
  X,
  ListTree,
  Check,
  MessageSquare,
  Link2,
  RotateCcw,
  ExternalLink,
} from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  setOutlineGroupCollapsed,
  removeOutlineGroup,
  EQ,
  reviewThreadLimits,
  addOutlineGroup,
  clearAxisOutlineGroups,
  validateHyperlinkRecord,
  _t,
  parseRangeAddress,
  addReviewThreadReply,
  setOutlineGroupLevel,
  nQ,
  setCellHyperlink,
  setReviewThreadResolved,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  staleOutlineMessage,
  formatRowColumnSpan,
  wS,
} from "./48-move-cells-and-reference-repair.js";

function OutlineToolsPanel(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    latestPropsRef = ne.useRef(props);
  latestPropsRef.current = props;
  const baselineRef = ne.useRef(null);
  baselineRef.current ||= {
    workbookId: props.workbook.id,
    revision: props.workbook.revision,
    snapshot: JSON.stringify(props.workbook),
    sheetId: props.sheetId,
    range: props.range,
  };
  const [axis, setAxis] = ne.useState(
      /column/i.test(props.mode) ? "column" : "row",
    ),
    [pageIndex, setPageIndex] = ne.useState(0),
    [errorMessage, setErrorMessage] = ne.useState(""),
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    groups = ne.useMemo(
      () => (sheet ? addOutlineGroup(sheet, axis) : []),
      [sheet, axis],
    ),
    rangeAddress = parseRangeAddress(props.range),
    rangeStart = axis === "row" ? rangeAddress.firstRow : rangeAddress.firstCol,
    rangeEnd = axis === "row" ? rangeAddress.lastRow : rangeAddress.lastCol,
    isReadOnly = props.readOnly || !!sheet?.protection?.enabled,
    baseline = baselineRef.current,
    hasWorkbookChanged = ne.useMemo(
      () => JSON.stringify(props.workbook) !== baseline.snapshot,
      [props.workbook, baseline.snapshot],
    ),
    isStale =
      baseline.workbookId !== props.workbook.id ||
      baseline.revision !== props.workbook.revision ||
      baseline.sheetId !== props.sheetId ||
      baseline.range !== props.range ||
      hasWorkbookChanged ||
      !sheet,
    disabled = isReadOnly || isStale,
    levelCount = Math.max(0, ...groups.map((group) => group.depth)) + 1,
    pageCount = Math.max(1, Math.ceil(groups.length / wS)),
    currentPage = Math.min(pageIndex, pageCount - 1);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button:not(:disabled)")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  function applyOutlineChange(updateWorkbook, changeLabel) {
    try {
      const latestProps = latestPropsRef.current;
      if (
        latestProps.readOnly ||
        latestProps.workbook.sheets.find(
          (sheet) => sheet.id === latestProps.sheetId,
        )?.protection?.enabled
      )
        throw Error(
          "This worksheet is read-only. Unlock it before changing outlines.",
        );
      if (
        latestProps.workbook.id !== baseline.workbookId ||
        latestProps.workbook.revision !== baseline.revision ||
        latestProps.sheetId !== baseline.sheetId ||
        latestProps.range !== baseline.range ||
        JSON.stringify(latestProps.workbook) !== baseline.snapshot
      )
        throw Error(staleOutlineMessage);
      const nextWorkbook = updateWorkbook(latestProps.workbook);
      (nextWorkbook !== latestProps.workbook &&
        latestProps.onWorkbookChange(nextWorkbook, changeLabel),
        latestProps.onClose());
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : "The outline could not be changed.",
      );
    }
  }
  return i.jsx("div", {
    className: "ot-backdrop",
    children: i.jsxs("section", {
      className: "ot-dialog",
      ref: dialogRef,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      tabIndex: -1,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.preventDefault(),
            event.stopPropagation(),
            latestPropsRef.current.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = Array.from(
              dialogRef.current?.querySelectorAll(
                "button:not(:disabled),select:not(:disabled),input:not(:disabled),[tabindex='0']",
              ) ?? [],
            ),
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1);
          event.shiftKey &&
          (document.activeElement === firstFocusable ||
            document.activeElement === dialogRef.current)
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === lastFocusable &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "ot-header",
          children: [
            i.jsx(ListTree, { size: 20 }),
            i.jsx("h2", {
              id: titleId,
              children: i.jsx(LocalizedText, {
                children: "Row and column outlines",
              }),
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": text("Close outline tools"),
              onClick: props.onClose,
              lang: language("Close outline tools"),
              "data-ui-label": "Close outline tools",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "ot-body",
          children: [
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Group detail rows or columns, leaving summary cells outside the group. Collapsing changes visibility only; formulas and totals remain unchanged.",
              }),
            }),
            isReadOnly &&
              i.jsx("p", {
                className: "ot-notice",
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "This worksheet is read-only. Outline details can be inspected.",
                }),
              }),
            (errorMessage || isStale) &&
              i.jsx("p", {
                role: "alert",
                className: "ot-error",
                children: errorMessage || staleOutlineMessage,
              }),
            i.jsxs("div", {
              className: "ot-create",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Group by" }),
                    i.jsxs("select", {
                      "aria-label": text("Outline axis"),
                      value: axis,
                      onChange: (event) => {
                        (setAxis(event.target.value),
                          setPageIndex(0),
                          setErrorMessage(""));
                      },
                      lang: language("Outline axis"),
                      "data-ui-label": "Outline axis",
                      children: [
                        i.jsx("option", { value: "row", children: "Rows" }),
                        i.jsx("option", {
                          value: "column",
                          children: "Columns",
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("div", {
                  children: [
                    i.jsx("span", {
                      className: "ot-small",
                      children: i.jsx(LocalizedText, {
                        children: "Selected range",
                      }),
                    }),
                    i.jsx("strong", {
                      children: formatRowColumnSpan(axis, rangeStart, rangeEnd),
                    }),
                  ],
                }),
                i.jsx("button", {
                  type: "button",
                  className: "ot-primary",
                  disabled: disabled,
                  onClick: () =>
                    applyOutlineChange(
                      (workbook) =>
                        EQ(workbook, props.sheetId, axis, rangeStart, rangeEnd),
                      "Group " +
                        formatRowColumnSpan(
                          axis,
                          rangeStart,
                          rangeEnd,
                        ).toLowerCase(),
                    ),
                  children: i.jsx(LocalizedText, {
                    children: "Group selection",
                  }),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "ot-levels",
              children: [
                i.jsx("strong", {
                  children: i.jsx(LocalizedText, { children: "Show level" }),
                }),
                Array.from({ length: levelCount }, (V, levelIndex) =>
                  i.jsx(
                    "button",
                    {
                      type: "button",
                      disabled: disabled || !groups.length,
                      "aria-label": "Show outline level " + (levelIndex + 1),
                      onClick: () =>
                        applyOutlineChange(
                          (workbook) =>
                            setOutlineGroupLevel(
                              workbook,
                              props.sheetId,
                              axis,
                              levelIndex + 1,
                            ),
                          "Show " + axis + " outline level " + (levelIndex + 1),
                        ),
                      children: levelIndex + 1,
                    },
                    levelIndex,
                  ),
                ),
              ],
            }),
            i.jsxs("p", {
              className: "ot-small",
              children: [
                i.jsx(LocalizedText, {
                  children:
                    "Level 1 hides all grouped details. Higher levels reveal nested detail; level ",
                }),
                levelCount,
                " ",
                i.jsx(LocalizedText, {
                  children:
                    "expands every group on this axis. Manually hidden rows and columns remain hidden.",
                }),
              ],
            }),
            i.jsxs("div", {
              className: "ot-groups-heading",
              children: [
                i.jsxs("h3", {
                  children: [
                    groups.length,
                    " ",
                    axis === "row" ? "row" : "column",
                    " ",
                    groups.length === 1 ? "group" : "groups",
                  ],
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: disabled || !groups.length,
                  onClick: () =>
                    applyOutlineChange(
                      (workbook) =>
                        clearAxisOutlineGroups(workbook, props.sheetId, axis),
                      "Clear " + axis + " outline groups",
                    ),
                  children: i.jsx(LocalizedText, { children: "Clear groups" }),
                }),
              ],
            }),
            groups.length
              ? i.jsx("ul", {
                  className: "ot-list",
                  children: groups
                    .slice(currentPage * wS, (currentPage + 1) * wS)
                    .map((group) =>
                      i.jsxs(
                        "li",
                        {
                          style: { paddingLeft: 12 + (group.depth - 1) * 12 },
                          children: [
                            i.jsxs("div", {
                              children: [
                                i.jsx("strong", {
                                  children: formatRowColumnSpan(
                                    axis,
                                    group.start,
                                    group.end,
                                  ),
                                }),
                                i.jsxs("span", {
                                  className: "ot-small",
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Depth ",
                                    }),
                                    group.depth,
                                    " ·",
                                    " ",
                                    group.collapsed ? "Collapsed" : "Expanded",
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("div", {
                              className: "ot-actions",
                              children: [
                                i.jsx("button", {
                                  type: "button",
                                  disabled: disabled,
                                  "aria-label":
                                    (group.collapsed
                                      ? "Expand "
                                      : "Collapse ") +
                                    formatRowColumnSpan(
                                      axis,
                                      group.start,
                                      group.end,
                                    ).toLowerCase(),
                                  onClick: () =>
                                    applyOutlineChange(
                                      (workbook) =>
                                        setOutlineGroupCollapsed(
                                          workbook,
                                          props.sheetId,
                                          group.id,
                                          !group.collapsed,
                                        ),
                                      (group.collapsed
                                        ? "Expand "
                                        : "Collapse ") +
                                        formatRowColumnSpan(
                                          axis,
                                          group.start,
                                          group.end,
                                        ).toLowerCase(),
                                    ),
                                  children: group.collapsed
                                    ? "Expand"
                                    : "Collapse",
                                }),
                                i.jsx("button", {
                                  type: "button",
                                  disabled: disabled,
                                  "aria-label":
                                    "Ungroup " +
                                    formatRowColumnSpan(
                                      axis,
                                      group.start,
                                      group.end,
                                    ).toLowerCase(),
                                  onClick: () =>
                                    applyOutlineChange(
                                      (workbook) =>
                                        removeOutlineGroup(
                                          workbook,
                                          props.sheetId,
                                          group.id,
                                        ),
                                      "Ungroup " +
                                        formatRowColumnSpan(
                                          axis,
                                          group.start,
                                          group.end,
                                        ).toLowerCase(),
                                    ),
                                  children: i.jsx(LocalizedText, {
                                    children: "Ungroup",
                                  }),
                                }),
                              ],
                            }),
                          ],
                        },
                        group.id,
                      ),
                    ),
                })
              : i.jsx("p", {
                  className: "ot-empty",
                  children: i.jsx(LocalizedText, {
                    children:
                      "Select detail cells in the worksheet, then choose Group selection.",
                  }),
                }),
            pageCount > 1 &&
              i.jsxs("div", {
                className: "ot-pager",
                children: [
                  i.jsx("button", {
                    type: "button",
                    disabled: currentPage === 0,
                    onClick: () => setPageIndex(currentPage - 1),
                    children: i.jsx(LocalizedText, {
                      children: "Previous groups",
                    }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      currentPage + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      pageCount,
                    ],
                  }),
                  i.jsx("button", {
                    type: "button",
                    disabled: currentPage + 1 === pageCount,
                    onClick: () => setPageIndex(currentPage + 1),
                    children: i.jsx(LocalizedText, { children: "Next groups" }),
                  }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "ot-footer",
          children: [
            i.jsx("span", {
              children: i.jsx(LocalizedText, {
                children: "Up to 8 nested groups · 1,000 groups per worksheet",
              }),
            }),
            i.jsx("button", {
              type: "button",
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Done" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const reviewPageSize = 20;
function resolveSelectionAnchorAddress(selection) {
  const rangeAddress = parseRangeAddress(selection.range);
  if (selection.active) {
    const activeCell = _t(selection.active);
    if (
      activeCell.row >= rangeAddress.firstRow &&
      activeCell.row <= rangeAddress.lastRow &&
      activeCell.col >= rangeAddress.firstCol &&
      activeCell.col <= rangeAddress.lastCol
    )
      return `${columnIndexToLetters(activeCell.col)}${activeCell.row + 1}`;
  }
  return `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`;
}
const formatLocalDateTime = (timestamp) => new Date(timestamp).toLocaleString();
function ReviewToolsPanel(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    onCloseRef = ne.useRef(props.onClose);
  onCloseRef.current = props.onClose;
  const baselineRef = ne.useRef({
      id: props.workbook.id,
      revision: props.workbook.revision,
      sheetId: props.sheetId,
      range: props.range,
      active: props.active,
      content: JSON.stringify(props.workbook),
    }),
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    selectionAnchorAddress = resolveSelectionAnchorAddress(props),
    hyperlink = sheet?.cells[selectionAnchorAddress]?.hyperlink,
    [activeTab, setActiveTab] = ne.useState(props.mode ?? "comments"),
    [threadFilter, setThreadFilter] = ne.useState("cell"),
    [pageIndex, setPageIndex] = ne.useState(0),
    [selectedThreadId, setSelectedThreadId] = ne.useState(
      props.workbook.reviewThreads?.find(
        (reviewThread) =>
          !reviewThread.orphaned &&
          reviewThread.anchor.sheetId === props.sheetId &&
          reviewThread.anchor.address === selectionAnchorAddress,
      )?.id ?? "",
    ),
    [messagePageIndex, setMessagePageIndex] = ne.useState(0),
    [newCommentText, setNewCommentText] = ne.useState(""),
    [replyText, setReplyText] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [linkKind, setLinkKind] = ne.useState(hyperlink?.kind ?? "external"),
    [url, setUrl] = ne.useState(
      hyperlink?.kind === "external" ? hyperlink.url : "",
    ),
    [sheetId, setLinkSheetId] = ne.useState(
      hyperlink?.kind === "internal" ? hyperlink.sheetId : props.sheetId,
    ),
    [address, setLinkAddress] = ne.useState(
      hyperlink?.kind === "internal" ? hyperlink.address : "A1",
    ),
    [linkLabel, setLinkLabel] = ne.useState(hyperlink?.label ?? ""),
    isStale =
      baselineRef.current.id !== props.workbook.id ||
      baselineRef.current.revision !== props.workbook.revision ||
      baselineRef.current.sheetId !== props.sheetId ||
      baselineRef.current.range !== props.range ||
      baselineRef.current.active !== props.active ||
      baselineRef.current.content !== JSON.stringify(props.workbook),
    disabled =
      props.readOnly || !!sheet?.protection?.enabled || isStale || !sheet,
    visibleThreads = [...(props.workbook.reviewThreads ?? [])]
      .sort((thread, otherThread) =>
        otherThread.updatedAt.localeCompare(thread.updatedAt),
      )
      .filter(
        (thread) =>
          threadFilter === "all" ||
          (threadFilter === "orphaned"
            ? !!thread.orphaned
            : threadFilter === "unresolved"
              ? !thread.resolved && !thread.orphaned
              : !thread.orphaned &&
                thread.anchor.sheetId === props.sheetId &&
                thread.anchor.address === selectionAnchorAddress),
      ),
    selectedThread = visibleThreads.find(
      (thread) => thread.id === selectedThreadId,
    ),
    anchorSheet = props.workbook.sheets.find(
      (sheet) => sheet.id === selectedThread?.anchor.sheetId,
    ),
    X$1 =
      props.readOnly ||
      isStale ||
      !!selectedThread?.orphaned ||
      !anchorSheet ||
      !!anchorSheet.protection?.enabled,
    actorName = props.actor?.trim() || "Reviewer";
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button:not(:disabled)")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const applyReviewChange = (buildNextWorkbook, changeLabel) => {
      try {
        if ((setErrorMessage(""), props.readOnly))
          throw new Error("This workbook is read-only.");
        if (isStale)
          throw new Error(
            "The workbook or selection changed. Close and reopen the review tools.",
          );
        const nextWorkbook = buildNextWorkbook();
        (nextWorkbook !== props.workbook &&
          JSON.stringify(nextWorkbook) !== JSON.stringify(props.workbook) &&
          props.onWorkbookChange(nextWorkbook, changeLabel),
          props.onClose());
      } catch (error) {
        setErrorMessage(
          error instanceof Error
            ? error.message
            : "The review change could not be saved.",
        );
      }
    },
    validatedLink = (() => {
      try {
        return validateHyperlinkRecord(
          linkKind === "external"
            ? { kind: "external", url: url, label: linkLabel }
            : {
                kind: "internal",
                sheetId: sheetId,
                address: address,
                label: linkLabel,
              },
        );
      } catch {
        return null;
      }
    })(),
    navigateToLink = (targetSheetId, targetAddress) => {
      try {
        (setErrorMessage(""),
          props.onNavigate?.(targetSheetId, targetAddress),
          props.onClose());
      } catch (error) {
        setErrorMessage(
          error instanceof Error
            ? error.message
            : "The link target could not be opened.",
        );
      }
    };
  return i.jsx("div", {
    className: "rv-backdrop",
    children: i.jsxs("div", {
      className: "rv-dialog",
      ref: dialogRef,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${titleId}-title`,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(), onCloseRef.current()),
          event.key === "Tab")
        ) {
          const focusableElements = [
              ...(dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),a[href],[tabindex="0"]',
              ) ?? []),
            ],
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1);
          event.shiftKey && document.activeElement === firstFocusable
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === lastFocusable &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "rv-header",
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${titleId}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Comments & links",
                  }),
                }),
                i.jsxs("p", {
                  children: [
                    sheet?.name ?? "Worksheet unavailable",
                    "!",
                    selectionAnchorAddress,
                  ],
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close review tools"),
              onClick: props.onClose,
              lang: language("Close review tools"),
              "data-ui-label": "Close review tools",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("div", {
          className: "rv-tabs",
          role: "tablist",
          "aria-label": text("Review tools"),
          lang: language("Review tools"),
          "data-ui-label": "Review tools",
          children: ["comments", "hyperlink"].map((tab, tabIndex) =>
            i.jsxs(
              "button",
              {
                role: "tab",
                id: `${titleId}-${tab}-tab`,
                "aria-controls": `${titleId}-${tab}-panel`,
                "aria-selected": activeTab === tab,
                tabIndex: activeTab === tab ? 0 : -1,
                onClick: () => {
                  (setActiveTab(tab), setErrorMessage(""));
                },
                onKeyDown: (event) => {
                  if (
                    ["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                      event.key,
                    )
                  ) {
                    event.preventDefault();
                    const targetTab =
                      event.key === "Home"
                        ? "comments"
                        : event.key === "End"
                          ? "hyperlink"
                          : tabIndex
                            ? "comments"
                            : "hyperlink";
                    (setActiveTab(targetTab),
                      document
                        .getElementById(`${titleId}-${targetTab}-tab`)
                        ?.focus());
                  }
                },
                children: [
                  tab === "comments"
                    ? i.jsx(MessageSquare, { size: 15 })
                    : i.jsx(Link2, { size: 15 }),
                  tab === "comments" ? "Comments" : "Hyperlink",
                ],
              },
              tab,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "rv-body",
          children: [
            isStale &&
              i.jsx("p", {
                className: "rv-error",
                role: "alert",
                children: i.jsx(LocalizedText, {
                  children:
                    "The workbook or selection changed. Close and reopen the review tools.",
                }),
              }),
            props.readOnly &&
              i.jsx("p", {
                className: "rv-muted",
                children: i.jsx(LocalizedText, {
                  children:
                    "This workbook is read-only. Discussions and links remain available for inspection.",
                }),
              }),
            !props.readOnly &&
              sheet?.protection?.enabled &&
              i.jsx("p", {
                className: "rv-muted",
                children: i.jsx(LocalizedText, {
                  children:
                    "This worksheet is protected. Unprotect it before adding comments or changing links.",
                }),
              }),
            activeTab === "comments"
              ? i.jsxs("section", {
                  role: "tabpanel",
                  id: `${titleId}-comments-panel`,
                  "aria-labelledby": `${titleId}-comments-tab`,
                  children: [
                    i.jsxs("p", {
                      className: "rv-muted",
                      children: [
                        i.jsx(LocalizedText, { children: "Posting as " }),
                        i.jsx("strong", { children: actorName }),
                        i.jsx(LocalizedText, {
                          children:
                            ". Names and times come from this application; they are not verified identities. Comments stay separate from cell notes.",
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      className: "rv-filter",
                      children: [
                        i.jsx(LocalizedText, { children: "Show discussions" }),
                        i.jsxs("select", {
                          "aria-label": text("Show discussions"),
                          value: threadFilter,
                          onChange: (event) => {
                            (setThreadFilter(event.target.value),
                              setPageIndex(0),
                              setSelectedThreadId(""),
                              setReplyText(""));
                          },
                          lang: language("Show discussions"),
                          "data-ui-label": "Show discussions",
                          children: [
                            i.jsx("option", {
                              value: "cell",
                              children: "This cell",
                            }),
                            i.jsx("option", {
                              value: "unresolved",
                              children: "All unresolved",
                            }),
                            i.jsx("option", {
                              value: "all",
                              children: "All discussions",
                            }),
                            i.jsx("option", {
                              value: "orphaned",
                              children: "Deleted targets",
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "rv-review-grid",
                      children: [
                        i.jsxs("div", {
                          className: "rv-thread-list",
                          "aria-label": text("Discussions"),
                          lang: language("Discussions"),
                          "data-ui-label": "Discussions",
                          children: [
                            !visibleThreads.length &&
                              i.jsx("p", {
                                className: "rv-empty",
                                children: i.jsx(LocalizedText, {
                                  children: "No discussions in this view.",
                                }),
                              }),
                            visibleThreads
                              .slice(
                                pageIndex * reviewPageSize,
                                (pageIndex + 1) * reviewPageSize,
                              )
                              .map((thread) =>
                                i.jsxs(
                                  "button",
                                  {
                                    "aria-pressed":
                                      thread.id === selectedThreadId,
                                    className: "rv-thread-button",
                                    onClick: () => {
                                      (setSelectedThreadId(thread.id),
                                        setReplyText(""),
                                        setMessagePageIndex(0),
                                        setErrorMessage(""));
                                    },
                                    children: [
                                      i.jsxs("span", {
                                        children: [
                                          i.jsxs("strong", {
                                            children: [
                                              thread.anchor.sheetName,
                                              "!",
                                              thread.anchor.address,
                                            ],
                                          }),
                                          i.jsx("small", {
                                            children: thread.orphaned
                                              ? "Deleted target"
                                              : thread.resolved
                                                ? "Resolved"
                                                : "Open",
                                          }),
                                        ],
                                      }),
                                      i.jsx("span", {
                                        className: "rv-thread-summary",
                                        children: thread.messages[0].text,
                                      }),
                                      i.jsxs("small", {
                                        children: [
                                          thread.messages.length,
                                          " ",
                                          thread.messages.length === 1
                                            ? "message"
                                            : "messages",
                                          " ",
                                          "· ",
                                          thread.messages.at(-1).author,
                                        ],
                                      }),
                                    ],
                                  },
                                  thread.id,
                                ),
                              ),
                            visibleThreads.length > reviewPageSize &&
                              i.jsxs("div", {
                                className: "rv-pages",
                                children: [
                                  i.jsx("button", {
                                    disabled: !pageIndex,
                                    onClick: () => setPageIndex(pageIndex - 1),
                                    children: i.jsx(LocalizedText, {
                                      children: "Previous discussions",
                                    }),
                                  }),
                                  i.jsxs("span", {
                                    children: [
                                      pageIndex + 1,
                                      " / ",
                                      Math.ceil(
                                        visibleThreads.length / reviewPageSize,
                                      ),
                                    ],
                                  }),
                                  i.jsx("button", {
                                    disabled:
                                      (pageIndex + 1) * reviewPageSize >=
                                      visibleThreads.length,
                                    onClick: () => setPageIndex(pageIndex + 1),
                                    children: i.jsx(LocalizedText, {
                                      children: "Next discussions",
                                    }),
                                  }),
                                ],
                              }),
                          ],
                        }),
                        i.jsx("div", {
                          className: "rv-thread-detail",
                          children: selectedThread
                            ? i.jsxs("section", {
                                "aria-label": text("Selected discussion"),
                                lang: language("Selected discussion"),
                                "data-ui-label": "Selected discussion",
                                children: [
                                  i.jsxs("div", {
                                    className: "rv-thread-title",
                                    children: [
                                      i.jsxs("h3", {
                                        children: [
                                          selectedThread.anchor.sheetName,
                                          "!",
                                          selectedThread.anchor.address,
                                        ],
                                      }),
                                      !selectedThread.orphaned &&
                                        props.onNavigate &&
                                        i.jsx("button", {
                                          onClick: () =>
                                            navigateToLink(
                                              selectedThread.anchor.sheetId,
                                              selectedThread.anchor.address,
                                            ),
                                          children: i.jsx(LocalizedText, {
                                            children: "Go to cell",
                                          }),
                                        }),
                                    ],
                                  }),
                                  selectedThread.orphaned &&
                                    i.jsxs("p", {
                                      className: "rv-orphan",
                                      children: [
                                        selectedThread.orphaned.reason ===
                                        "worksheet-deleted"
                                          ? "The original worksheet was deleted."
                                          : "The original row or column was deleted.",
                                        " ",
                                        i.jsx(LocalizedText, {
                                          children:
                                            "This discussion is retained for review and cannot receive new replies.",
                                        }),
                                      ],
                                    }),
                                  selectedThread.resolved &&
                                    i.jsxs("p", {
                                      className: "rv-resolution",
                                      children: [
                                        i.jsx(Check, { size: 14 }),
                                        " ",
                                        i.jsx(LocalizedText, {
                                          children: "Resolved by",
                                        }),
                                        " ",
                                        selectedThread.resolved.author,
                                        " ·",
                                        " ",
                                        i.jsx("time", {
                                          dateTime:
                                            selectedThread.resolved.createdAt,
                                          title:
                                            selectedThread.resolved.createdAt,
                                          children: formatLocalDateTime(
                                            selectedThread.resolved.createdAt,
                                          ),
                                        }),
                                      ],
                                    }),
                                  selectedThread.messages
                                    .slice(
                                      messagePageIndex * 10,
                                      (messagePageIndex + 1) * 10,
                                    )
                                    .map((message) =>
                                      i.jsxs(
                                        "article",
                                        {
                                          className: "rv-message",
                                          children: [
                                            i.jsxs("div", {
                                              children: [
                                                i.jsx("strong", {
                                                  children: message.author,
                                                }),
                                                i.jsx("time", {
                                                  dateTime: message.createdAt,
                                                  title: message.createdAt,
                                                  children: formatLocalDateTime(
                                                    message.createdAt,
                                                  ),
                                                }),
                                              ],
                                            }),
                                            i.jsx("p", {
                                              children: message.text,
                                            }),
                                          ],
                                        },
                                        message.id,
                                      ),
                                    ),
                                  selectedThread.messages.length > 10 &&
                                    i.jsxs("div", {
                                      className: "rv-pages",
                                      children: [
                                        i.jsx("button", {
                                          disabled: !messagePageIndex,
                                          onClick: () =>
                                            setMessagePageIndex(
                                              messagePageIndex - 1,
                                            ),
                                          children: i.jsx(LocalizedText, {
                                            children: "Previous messages",
                                          }),
                                        }),
                                        i.jsxs("span", {
                                          children: [
                                            messagePageIndex + 1,
                                            " /",
                                            " ",
                                            Math.ceil(
                                              selectedThread.messages.length /
                                                10,
                                            ),
                                          ],
                                        }),
                                        i.jsx("button", {
                                          disabled:
                                            (messagePageIndex + 1) * 10 >=
                                            selectedThread.messages.length,
                                          onClick: () =>
                                            setMessagePageIndex(
                                              messagePageIndex + 1,
                                            ),
                                          children: i.jsx(LocalizedText, {
                                            children: "Next messages",
                                          }),
                                        }),
                                      ],
                                    }),
                                  !selectedThread.orphaned &&
                                    i.jsx("div", {
                                      className: "rv-thread-actions",
                                      children: i.jsxs("button", {
                                        disabled: X$1,
                                        onClick: () =>
                                          applyReviewChange(
                                            () =>
                                              setReviewThreadResolved(
                                                props.workbook,
                                                selectedThread.id,
                                                !selectedThread.resolved,
                                                actorName,
                                                { readOnly: props.readOnly },
                                              ),
                                            selectedThread.resolved
                                              ? "Reopen review discussion"
                                              : "Resolve review discussion",
                                          ),
                                        children: [
                                          selectedThread.resolved
                                            ? i.jsx(RotateCcw, { size: 14 })
                                            : i.jsx(Check, { size: 14 }),
                                          selectedThread.resolved
                                            ? "Reopen discussion"
                                            : "Resolve discussion",
                                        ],
                                      }),
                                    }),
                                  !selectedThread.resolved &&
                                    !selectedThread.orphaned &&
                                    i.jsxs("form", {
                                      className: "rv-compose",
                                      onSubmit: (event) => {
                                        (event.preventDefault(),
                                          applyReviewChange(
                                            () =>
                                              addReviewThreadReply(
                                                props.workbook,
                                                selectedThread.id,
                                                replyText,
                                                actorName,
                                                {
                                                  readOnly: props.readOnly,
                                                },
                                              ),
                                            "Reply to review discussion",
                                          ));
                                      },
                                      children: [
                                        i.jsxs("label", {
                                          children: [
                                            i.jsx(LocalizedText, {
                                              children: "Reply",
                                            }),
                                            i.jsx("textarea", {
                                              "aria-label": text("Reply"),
                                              rows: 3,
                                              maxLength:
                                                reviewThreadLimits.textLength,
                                              value: replyText,
                                              disabled: X$1,
                                              onChange: (event) =>
                                                setReplyText(
                                                  event.target.value,
                                                ),
                                              lang: language("Reply"),
                                              "data-ui-label": "Reply",
                                            }),
                                          ],
                                        }),
                                        i.jsx("button", {
                                          className: "rv-primary",
                                          disabled: X$1 || !replyText.trim(),
                                          children: i.jsx(LocalizedText, {
                                            children: "Post reply",
                                          }),
                                        }),
                                      ],
                                    }),
                                ],
                              })
                            : i.jsx("p", {
                                className: "rv-empty",
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Select a discussion to read its messages.",
                                }),
                              }),
                        }),
                      ],
                    }),
                    i.jsxs("form", {
                      className: "rv-compose rv-new",
                      onSubmit: (event) => {
                        (event.preventDefault(),
                          applyReviewChange(
                            () =>
                              nQ(
                                props.workbook,
                                props.sheetId,
                                selectionAnchorAddress,
                                newCommentText,
                                actorName,
                                {
                                  readOnly: props.readOnly,
                                },
                              ),
                            "Add review discussion",
                          ));
                      },
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "New comment on ",
                            }),
                            sheet?.name,
                            "!",
                            selectionAnchorAddress,
                            i.jsx("textarea", {
                              "aria-label": text("New comment"),
                              rows: 3,
                              maxLength: reviewThreadLimits.textLength,
                              value: newCommentText,
                              disabled: disabled,
                              onChange: (event) =>
                                setNewCommentText(event.target.value),
                              lang: language("New comment"),
                              "data-ui-label": "New comment",
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "rv-compose-footer",
                          children: [
                            i.jsxs("small", {
                              children: [
                                newCommentText.length,
                                " / ",
                                reviewThreadLimits.textLength,
                              ],
                            }),
                            i.jsx("button", {
                              className: "rv-primary",
                              disabled: disabled || !newCommentText.trim(),
                              children: i.jsx(LocalizedText, {
                                children: "Post comment",
                              }),
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      className: "rv-muted",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Comments are stored in this workbook. Typing an @name does not notify or send a message to anyone.",
                      }),
                    }),
                  ],
                })
              : i.jsxs("section", {
                  role: "tabpanel",
                  id: `${titleId}-hyperlink-panel`,
                  "aria-labelledby": `${titleId}-hyperlink-tab`,
                  children: [
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Add a link to this cell without replacing its value, formula, formatting or note.",
                      }),
                    }),
                    i.jsxs("form", {
                      className: "rv-link-form",
                      onSubmit: (event) => {
                        (event.preventDefault(),
                          applyReviewChange(
                            () =>
                              setCellHyperlink(
                                props.workbook,
                                props.sheetId,
                                selectionAnchorAddress,
                                linkKind === "external"
                                  ? {
                                      kind: "external",
                                      url: url,
                                      label: linkLabel,
                                    }
                                  : {
                                      kind: "internal",
                                      sheetId: sheetId,
                                      address: address,
                                      label: linkLabel,
                                    },
                                { readOnly: props.readOnly },
                              ),
                            "Update cell hyperlink",
                          ));
                      },
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Link type" }),
                            i.jsxs("select", {
                              "aria-label": text("Link type"),
                              value: linkKind,
                              disabled: disabled,
                              onChange: (event) =>
                                setLinkKind(event.target.value),
                              lang: language("Link type"),
                              "data-ui-label": "Link type",
                              children: [
                                i.jsx("option", {
                                  value: "external",
                                  children: "Web or email address",
                                }),
                                i.jsx("option", {
                                  value: "internal",
                                  children: "Place in this workbook",
                                }),
                              ],
                            }),
                          ],
                        }),
                        linkKind === "external"
                          ? i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, { children: "Address" }),
                                i.jsx("input", {
                                  "aria-label": text("Link address"),
                                  type: "text",
                                  value: url,
                                  disabled: disabled,
                                  placeholder: text(
                                    "https://example.com/evidence",
                                  ),
                                  maxLength: reviewThreadLimits.urlLength,
                                  onChange: (event) =>
                                    setUrl(event.target.value),
                                  lang: language("Link address"),
                                  "data-ui-label": "Link address",
                                }),
                              ],
                            })
                          : i.jsxs("div", {
                              className: "rv-link-target",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Worksheet",
                                    }),
                                    i.jsx("select", {
                                      "aria-label": text("Link worksheet"),
                                      value: sheetId,
                                      disabled: disabled,
                                      onChange: (event) =>
                                        setLinkSheetId(event.target.value),
                                      lang: language("Link worksheet"),
                                      "data-ui-label": "Link worksheet",
                                      children: props.workbook.sheets.map(
                                        (sheet) =>
                                          i.jsxs(
                                            "option",
                                            {
                                              value: sheet.id,
                                              children: [
                                                sheet.name,
                                                sheet.hidden ? " (hidden)" : "",
                                              ],
                                            },
                                            sheet.id,
                                          ),
                                      ),
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Cell address",
                                    }),
                                    i.jsx("input", {
                                      "aria-label": text("Link cell address"),
                                      value: address,
                                      maxLength: 12,
                                      disabled: disabled,
                                      onChange: (event) =>
                                        setLinkAddress(event.target.value),
                                      lang: language("Link cell address"),
                                      "data-ui-label": "Link cell address",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Display text (optional)",
                            }),
                            i.jsx("input", {
                              "aria-label": text("Link display text"),
                              value: linkLabel,
                              disabled: disabled,
                              maxLength: reviewThreadLimits.labelLength,
                              onChange: (event) =>
                                setLinkLabel(event.target.value),
                              lang: language("Link display text"),
                              "data-ui-label": "Link display text",
                            }),
                          ],
                        }),
                        validatedLink &&
                          i.jsxs("div", {
                            className: "rv-link-preview",
                            "aria-label": text("Link preview"),
                            lang: language("Link preview"),
                            "data-ui-label": "Link preview",
                            children: [
                              i.jsx(Link2, { size: 15 }),
                              i.jsx("span", {
                                children:
                                  validatedLink.label ||
                                  (validatedLink.kind === "external"
                                    ? validatedLink.url
                                    : `${props.workbook.sheets.find((sheet) => sheet.id === validatedLink.sheetId)?.name ?? "Missing worksheet"}!${validatedLink.address}`),
                              }),
                              validatedLink.kind === "external"
                                ? i.jsxs("a", {
                                    href: validatedLink.url,
                                    target: "_blank",
                                    rel: "noopener noreferrer",
                                    "aria-label": text("Open link preview"),
                                    lang: language("Open link preview"),
                                    "data-ui-label": "Open link preview",
                                    children: [
                                      i.jsx(ExternalLink, { size: 15 }),
                                      i.jsx(LocalizedText, {
                                        children: "Open link",
                                      }),
                                    ],
                                  })
                                : props.onNavigate &&
                                  i.jsx("button", {
                                    type: "button",
                                    onClick: () =>
                                      navigateToLink(
                                        validatedLink.sheetId,
                                        validatedLink.address,
                                      ),
                                    children: i.jsx(LocalizedText, {
                                      children: "Go to target",
                                    }),
                                  }),
                            ],
                          }),
                        i.jsx("p", {
                          className: "rv-muted",
                          children: i.jsx(LocalizedText, {
                            children:
                              "Web links use https or http. Email links use one mailto recipient. Local files, embedded credentials and executable links are not accepted. Opening a link is a separate action.",
                          }),
                        }),
                        i.jsxs("div", {
                          className: "rv-link-actions",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              disabled: disabled || !hyperlink,
                              onClick: () =>
                                applyReviewChange(
                                  () =>
                                    setCellHyperlink(
                                      props.workbook,
                                      props.sheetId,
                                      selectionAnchorAddress,
                                      null,
                                      {
                                        readOnly: props.readOnly,
                                      },
                                    ),
                                  "Remove cell hyperlink",
                                ),
                              children: i.jsx(LocalizedText, {
                                children: "Remove link",
                              }),
                            }),
                            i.jsx("button", {
                              className: "rv-primary",
                              disabled:
                                disabled ||
                                !(linkKind === "external"
                                  ? url.trim()
                                  : address.trim()),
                              children: i.jsx(LocalizedText, {
                                children: "Save link",
                              }),
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
            errorMessage &&
              i.jsx("p", {
                role: "alert",
                className: "rv-error",
                children: errorMessage,
              }),
          ],
        }),
      ],
    }),
  });
}
const formatCellValueLabel = (value) =>
    value === null
      ? "Blank"
      : typeof value == "string"
        ? JSON.stringify(value)
        : String(value),
  comparatorSymbols = { gt: ">", gte: "≥", lt: "<", lte: "≤", eq: "=" };

export {
  ReviewToolsPanel,
  comparatorSymbols,
  OutlineToolsPanel,
  formatCellValueLabel,
};
