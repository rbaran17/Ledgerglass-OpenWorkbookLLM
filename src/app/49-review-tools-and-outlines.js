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

function OutlineToolsPanel(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e);
  a.current = e;
  const s = ne.useRef(null);
  s.current ||= {
    workbookId: e.workbook.id,
    revision: e.workbook.revision,
    snapshot: JSON.stringify(e.workbook),
    sheetId: e.sheetId,
    range: e.range,
  };
  const [d, u] = ne.useState(/column/i.test(e.mode) ? "column" : "row"),
    [l, c] = ne.useState(0),
    [h, f] = ne.useState(""),
    sheet = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    m = ne.useMemo(() => (sheet ? addOutlineGroup(sheet, d) : []), [sheet, d]),
    rangeAddress = parseRangeAddress(e.range),
    g = d === "row" ? rangeAddress.firstRow : rangeAddress.firstCol,
    b = d === "row" ? rangeAddress.lastRow : rangeAddress.lastCol,
    y = e.readOnly || !!sheet?.protection?.enabled,
    w = s.current,
    E = ne.useMemo(
      () => JSON.stringify(e.workbook) !== w.snapshot,
      [e.workbook, w.snapshot],
    ),
    A =
      w.workbookId !== e.workbook.id ||
      w.revision !== e.workbook.revision ||
      w.sheetId !== e.sheetId ||
      w.range !== e.range ||
      E ||
      !sheet,
    disabled = y || A,
    F = Math.max(0, ...m.map((V) => V.depth)) + 1,
    I = Math.max(1, Math.ceil(m.length / wS)),
    O = Math.min(l, I - 1);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      o.current?.querySelector("button:not(:disabled)")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  function j(V, D) {
    try {
      const C = a.current;
      if (
        C.readOnly ||
        C.workbook.sheets.find((sheet) => sheet.id === C.sheetId)?.protection
          ?.enabled
      )
        throw Error(
          "This worksheet is read-only. Unlock it before changing outlines.",
        );
      if (
        C.workbook.id !== w.workbookId ||
        C.workbook.revision !== w.revision ||
        C.sheetId !== w.sheetId ||
        C.range !== w.range ||
        JSON.stringify(C.workbook) !== w.snapshot
      )
        throw Error(staleOutlineMessage);
      const T = V(C.workbook);
      (T !== C.workbook && C.onWorkbookChange(T, D), C.onClose());
    } catch (error) {
      f(
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
      ref: o,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": n,
      tabIndex: -1,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.preventDefault(),
            event.stopPropagation(),
            a.current.onClose()),
          event.key === "Tab")
        ) {
          const B = Array.from(
              o.current?.querySelectorAll(
                "button:not(:disabled),select:not(:disabled),input:not(:disabled),[tabindex='0']",
              ) ?? [],
            ),
            _ = B[0],
            C = B.at(-1);
          event.shiftKey &&
          (document.activeElement === _ || document.activeElement === o.current)
            ? (event.preventDefault(), C == null || C.focus())
            : !event.shiftKey &&
              document.activeElement === C &&
              (event.preventDefault(), _ == null || _.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "ot-header",
          children: [
            i.jsx(ListTree, { size: 20 }),
            i.jsx("h2", {
              id: n,
              children: i.jsx(LocalizedText, {
                children: "Row and column outlines",
              }),
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": text("Close outline tools"),
              onClick: e.onClose,
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
            y &&
              i.jsx("p", {
                className: "ot-notice",
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "This worksheet is read-only. Outline details can be inspected.",
                }),
              }),
            (h || A) &&
              i.jsx("p", {
                role: "alert",
                className: "ot-error",
                children: h || staleOutlineMessage,
              }),
            i.jsxs("div", {
              className: "ot-create",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Group by" }),
                    i.jsxs("select", {
                      "aria-label": text("Outline axis"),
                      value: d,
                      onChange: (event) => {
                        (u(event.target.value), c(0), f(""));
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
                    i.jsx("strong", { children: formatRowColumnSpan(d, g, b) }),
                  ],
                }),
                i.jsx("button", {
                  type: "button",
                  className: "ot-primary",
                  disabled: disabled,
                  onClick: () =>
                    j(
                      (workbook) => EQ(workbook, e.sheetId, d, g, b),
                      "Group " + formatRowColumnSpan(d, g, b).toLowerCase(),
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
                Array.from({ length: F }, (V, D) =>
                  i.jsx(
                    "button",
                    {
                      type: "button",
                      disabled: disabled || !m.length,
                      "aria-label": "Show outline level " + (D + 1),
                      onClick: () =>
                        j(
                          (workbook) =>
                            setOutlineGroupLevel(workbook, e.sheetId, d, D + 1),
                          "Show " + d + " outline level " + (D + 1),
                        ),
                      children: D + 1,
                    },
                    D,
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
                F,
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
                    m.length,
                    " ",
                    d === "row" ? "row" : "column",
                    " ",
                    m.length === 1 ? "group" : "groups",
                  ],
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: disabled || !m.length,
                  onClick: () =>
                    j(
                      (workbook) =>
                        clearAxisOutlineGroups(workbook, e.sheetId, d),
                      "Clear " + d + " outline groups",
                    ),
                  children: i.jsx(LocalizedText, { children: "Clear groups" }),
                }),
              ],
            }),
            m.length
              ? i.jsx("ul", {
                  className: "ot-list",
                  children: m.slice(O * wS, (O + 1) * wS).map((V) =>
                    i.jsxs(
                      "li",
                      {
                        style: { paddingLeft: 12 + (V.depth - 1) * 12 },
                        children: [
                          i.jsxs("div", {
                            children: [
                              i.jsx("strong", {
                                children: formatRowColumnSpan(
                                  d,
                                  V.start,
                                  V.end,
                                ),
                              }),
                              i.jsxs("span", {
                                className: "ot-small",
                                children: [
                                  i.jsx(LocalizedText, { children: "Depth " }),
                                  V.depth,
                                  " ·",
                                  " ",
                                  V.collapsed ? "Collapsed" : "Expanded",
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
                                  (V.collapsed ? "Expand " : "Collapse ") +
                                  formatRowColumnSpan(
                                    d,
                                    V.start,
                                    V.end,
                                  ).toLowerCase(),
                                onClick: () =>
                                  j(
                                    (workbook) =>
                                      setOutlineGroupCollapsed(
                                        workbook,
                                        e.sheetId,
                                        V.id,
                                        !V.collapsed,
                                      ),
                                    (V.collapsed ? "Expand " : "Collapse ") +
                                      formatRowColumnSpan(
                                        d,
                                        V.start,
                                        V.end,
                                      ).toLowerCase(),
                                  ),
                                children: V.collapsed ? "Expand" : "Collapse",
                              }),
                              i.jsx("button", {
                                type: "button",
                                disabled: disabled,
                                "aria-label":
                                  "Ungroup " +
                                  formatRowColumnSpan(
                                    d,
                                    V.start,
                                    V.end,
                                  ).toLowerCase(),
                                onClick: () =>
                                  j(
                                    (workbook) =>
                                      removeOutlineGroup(
                                        workbook,
                                        e.sheetId,
                                        V.id,
                                      ),
                                    "Ungroup " +
                                      formatRowColumnSpan(
                                        d,
                                        V.start,
                                        V.end,
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
                      V.id,
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
            I > 1 &&
              i.jsxs("div", {
                className: "ot-pager",
                children: [
                  i.jsx("button", {
                    type: "button",
                    disabled: O === 0,
                    onClick: () => c(O - 1),
                    children: i.jsx(LocalizedText, {
                      children: "Previous groups",
                    }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      O + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      I,
                    ],
                  }),
                  i.jsx("button", {
                    type: "button",
                    disabled: O + 1 === I,
                    onClick: () => c(O + 1),
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
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Done" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const reviewPageSize = 20;
function resolveSelectionAnchorAddress(e) {
  const rangeAddress = parseRangeAddress(e.range);
  if (e.active) {
    const t = _t(e.active);
    if (
      t.row >= rangeAddress.firstRow &&
      t.row <= rangeAddress.lastRow &&
      t.col >= rangeAddress.firstCol &&
      t.col <= rangeAddress.lastCol
    )
      return `${columnIndexToLetters(t.col)}${t.row + 1}`;
  }
  return `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`;
}
const formatLocalDateTime = (e) => new Date(e).toLocaleString();
function ReviewToolsPanel(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const s = ne.useRef({
      id: e.workbook.id,
      revision: e.workbook.revision,
      sheetId: e.sheetId,
      range: e.range,
      active: e.active,
      content: JSON.stringify(e.workbook),
    }),
    sheet = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    selectionAnchorAddress = resolveSelectionAnchorAddress(e),
    l = sheet?.cells[selectionAnchorAddress]?.hyperlink,
    [c, h] = ne.useState(e.mode ?? "comments"),
    [f, p] = ne.useState("cell"),
    [m, v] = ne.useState(0),
    [g, b] = ne.useState(
      e.workbook.reviewThreads?.find(
        (reviewThread) =>
          !reviewThread.orphaned &&
          reviewThread.anchor.sheetId === e.sheetId &&
          reviewThread.anchor.address === selectionAnchorAddress,
      )?.id ?? "",
    ),
    [y, w] = ne.useState(0),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(""),
    [j, z] = ne.useState(l?.kind ?? "external"),
    [url, D] = ne.useState(l?.kind === "external" ? l.url : ""),
    [sheetId, _] = ne.useState(l?.kind === "internal" ? l.sheetId : e.sheetId),
    [address, T] = ne.useState(l?.kind === "internal" ? l.address : "A1"),
    [N, R] = ne.useState(l?.label ?? ""),
    L =
      s.current.id !== e.workbook.id ||
      s.current.revision !== e.workbook.revision ||
      s.current.sheetId !== e.sheetId ||
      s.current.range !== e.range ||
      s.current.active !== e.active ||
      s.current.content !== JSON.stringify(e.workbook),
    disabled = e.readOnly || !!sheet?.protection?.enabled || L || !sheet,
    G = [...(e.workbook.reviewThreads ?? [])]
      .sort((se, we) => we.updatedAt.localeCompare(se.updatedAt))
      .filter(
        (se) =>
          f === "all" ||
          (f === "orphaned"
            ? !!se.orphaned
            : f === "unresolved"
              ? !se.resolved && !se.orphaned
              : !se.orphaned &&
                se.anchor.sheetId === e.sheetId &&
                se.anchor.address === selectionAnchorAddress),
      ),
    H = G.find((se) => se.id === g),
    Q = e.workbook.sheets.find((sheet) => sheet.id === H?.anchor.sheetId),
    X$1 = e.readOnly || L || !!H?.orphaned || !Q || !!Q.protection?.enabled,
    J = e.actor?.trim() || "Reviewer";
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      o.current?.querySelector("button:not(:disabled)")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const q = (se, we) => {
      try {
        if ((O(""), e.readOnly)) throw new Error("This workbook is read-only.");
        if (L)
          throw new Error(
            "The workbook or selection changed. Close and reopen the review tools.",
          );
        const Ee = se();
        (Ee !== e.workbook &&
          JSON.stringify(Ee) !== JSON.stringify(e.workbook) &&
          e.onWorkbookChange(Ee, we),
          e.onClose());
      } catch (error) {
        O(
          error instanceof Error
            ? error.message
            : "The review change could not be saved.",
        );
      }
    },
    K = (() => {
      try {
        return validateHyperlinkRecord(
          j === "external"
            ? { kind: "external", url: url, label: N }
            : {
                kind: "internal",
                sheetId: sheetId,
                address: address,
                label: N,
              },
        );
      } catch {
        return null;
      }
    })(),
    te = (se, we) => {
      try {
        (O(""), e.onNavigate?.(se, we), e.onClose());
      } catch (error) {
        O(
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
      ref: o,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), a.current()),
          event.key === "Tab")
        ) {
          const Ee = [
              ...(o.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),a[href],[tabindex="0"]',
              ) ?? []),
            ],
            Ae = Ee[0],
            ye = Ee.at(-1);
          event.shiftKey && document.activeElement === Ae
            ? (event.preventDefault(), ye == null || ye.focus())
            : !event.shiftKey &&
              document.activeElement === ye &&
              (event.preventDefault(), Ae == null || Ae.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "rv-header",
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
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
              onClick: e.onClose,
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
          children: ["comments", "hyperlink"].map((se, we) =>
            i.jsxs(
              "button",
              {
                role: "tab",
                id: `${n}-${se}-tab`,
                "aria-controls": `${n}-${se}-panel`,
                "aria-selected": c === se,
                tabIndex: c === se ? 0 : -1,
                onClick: () => {
                  (h(se), O(""));
                },
                onKeyDown: (event) => {
                  if (
                    ["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                      event.key,
                    )
                  ) {
                    event.preventDefault();
                    const ye =
                      event.key === "Home"
                        ? "comments"
                        : event.key === "End"
                          ? "hyperlink"
                          : we
                            ? "comments"
                            : "hyperlink";
                    (h(ye), document.getElementById(`${n}-${ye}-tab`)?.focus());
                  }
                },
                children: [
                  se === "comments"
                    ? i.jsx(MessageSquare, { size: 15 })
                    : i.jsx(Link2, { size: 15 }),
                  se === "comments" ? "Comments" : "Hyperlink",
                ],
              },
              se,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "rv-body",
          children: [
            L &&
              i.jsx("p", {
                className: "rv-error",
                role: "alert",
                children: i.jsx(LocalizedText, {
                  children:
                    "The workbook or selection changed. Close and reopen the review tools.",
                }),
              }),
            e.readOnly &&
              i.jsx("p", {
                className: "rv-muted",
                children: i.jsx(LocalizedText, {
                  children:
                    "This workbook is read-only. Discussions and links remain available for inspection.",
                }),
              }),
            !e.readOnly &&
              sheet?.protection?.enabled &&
              i.jsx("p", {
                className: "rv-muted",
                children: i.jsx(LocalizedText, {
                  children:
                    "This worksheet is protected. Unprotect it before adding comments or changing links.",
                }),
              }),
            c === "comments"
              ? i.jsxs("section", {
                  role: "tabpanel",
                  id: `${n}-comments-panel`,
                  "aria-labelledby": `${n}-comments-tab`,
                  children: [
                    i.jsxs("p", {
                      className: "rv-muted",
                      children: [
                        i.jsx(LocalizedText, { children: "Posting as " }),
                        i.jsx("strong", { children: J }),
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
                          value: f,
                          onChange: (event) => {
                            (p(event.target.value), v(0), b(""), F(""));
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
                            !G.length &&
                              i.jsx("p", {
                                className: "rv-empty",
                                children: i.jsx(LocalizedText, {
                                  children: "No discussions in this view.",
                                }),
                              }),
                            G.slice(
                              m * reviewPageSize,
                              (m + 1) * reviewPageSize,
                            ).map((se) =>
                              i.jsxs(
                                "button",
                                {
                                  "aria-pressed": se.id === g,
                                  className: "rv-thread-button",
                                  onClick: () => {
                                    (b(se.id), F(""), w(0), O(""));
                                  },
                                  children: [
                                    i.jsxs("span", {
                                      children: [
                                        i.jsxs("strong", {
                                          children: [
                                            se.anchor.sheetName,
                                            "!",
                                            se.anchor.address,
                                          ],
                                        }),
                                        i.jsx("small", {
                                          children: se.orphaned
                                            ? "Deleted target"
                                            : se.resolved
                                              ? "Resolved"
                                              : "Open",
                                        }),
                                      ],
                                    }),
                                    i.jsx("span", {
                                      className: "rv-thread-summary",
                                      children: se.messages[0].text,
                                    }),
                                    i.jsxs("small", {
                                      children: [
                                        se.messages.length,
                                        " ",
                                        se.messages.length === 1
                                          ? "message"
                                          : "messages",
                                        " ",
                                        "· ",
                                        se.messages.at(-1).author,
                                      ],
                                    }),
                                  ],
                                },
                                se.id,
                              ),
                            ),
                            G.length > reviewPageSize &&
                              i.jsxs("div", {
                                className: "rv-pages",
                                children: [
                                  i.jsx("button", {
                                    disabled: !m,
                                    onClick: () => v(m - 1),
                                    children: i.jsx(LocalizedText, {
                                      children: "Previous discussions",
                                    }),
                                  }),
                                  i.jsxs("span", {
                                    children: [
                                      m + 1,
                                      " / ",
                                      Math.ceil(G.length / reviewPageSize),
                                    ],
                                  }),
                                  i.jsx("button", {
                                    disabled:
                                      (m + 1) * reviewPageSize >= G.length,
                                    onClick: () => v(m + 1),
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
                          children: H
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
                                          H.anchor.sheetName,
                                          "!",
                                          H.anchor.address,
                                        ],
                                      }),
                                      !H.orphaned &&
                                        e.onNavigate &&
                                        i.jsx("button", {
                                          onClick: () =>
                                            te(
                                              H.anchor.sheetId,
                                              H.anchor.address,
                                            ),
                                          children: i.jsx(LocalizedText, {
                                            children: "Go to cell",
                                          }),
                                        }),
                                    ],
                                  }),
                                  H.orphaned &&
                                    i.jsxs("p", {
                                      className: "rv-orphan",
                                      children: [
                                        H.orphaned.reason ===
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
                                  H.resolved &&
                                    i.jsxs("p", {
                                      className: "rv-resolution",
                                      children: [
                                        i.jsx(Check, { size: 14 }),
                                        " ",
                                        i.jsx(LocalizedText, {
                                          children: "Resolved by",
                                        }),
                                        " ",
                                        H.resolved.author,
                                        " ·",
                                        " ",
                                        i.jsx("time", {
                                          dateTime: H.resolved.createdAt,
                                          title: H.resolved.createdAt,
                                          children: formatLocalDateTime(
                                            H.resolved.createdAt,
                                          ),
                                        }),
                                      ],
                                    }),
                                  H.messages
                                    .slice(y * 10, (y + 1) * 10)
                                    .map((se) =>
                                      i.jsxs(
                                        "article",
                                        {
                                          className: "rv-message",
                                          children: [
                                            i.jsxs("div", {
                                              children: [
                                                i.jsx("strong", {
                                                  children: se.author,
                                                }),
                                                i.jsx("time", {
                                                  dateTime: se.createdAt,
                                                  title: se.createdAt,
                                                  children: formatLocalDateTime(
                                                    se.createdAt,
                                                  ),
                                                }),
                                              ],
                                            }),
                                            i.jsx("p", { children: se.text }),
                                          ],
                                        },
                                        se.id,
                                      ),
                                    ),
                                  H.messages.length > 10 &&
                                    i.jsxs("div", {
                                      className: "rv-pages",
                                      children: [
                                        i.jsx("button", {
                                          disabled: !y,
                                          onClick: () => w(y - 1),
                                          children: i.jsx(LocalizedText, {
                                            children: "Previous messages",
                                          }),
                                        }),
                                        i.jsxs("span", {
                                          children: [
                                            y + 1,
                                            " /",
                                            " ",
                                            Math.ceil(H.messages.length / 10),
                                          ],
                                        }),
                                        i.jsx("button", {
                                          disabled:
                                            (y + 1) * 10 >= H.messages.length,
                                          onClick: () => w(y + 1),
                                          children: i.jsx(LocalizedText, {
                                            children: "Next messages",
                                          }),
                                        }),
                                      ],
                                    }),
                                  !H.orphaned &&
                                    i.jsx("div", {
                                      className: "rv-thread-actions",
                                      children: i.jsxs("button", {
                                        disabled: X$1,
                                        onClick: () =>
                                          q(
                                            () =>
                                              setReviewThreadResolved(
                                                e.workbook,
                                                H.id,
                                                !H.resolved,
                                                J,
                                                { readOnly: e.readOnly },
                                              ),
                                            H.resolved
                                              ? "Reopen review discussion"
                                              : "Resolve review discussion",
                                          ),
                                        children: [
                                          H.resolved
                                            ? i.jsx(RotateCcw, { size: 14 })
                                            : i.jsx(Check, { size: 14 }),
                                          H.resolved
                                            ? "Reopen discussion"
                                            : "Resolve discussion",
                                        ],
                                      }),
                                    }),
                                  !H.resolved &&
                                    !H.orphaned &&
                                    i.jsxs("form", {
                                      className: "rv-compose",
                                      onSubmit: (event) => {
                                        (event.preventDefault(),
                                          q(
                                            () =>
                                              addReviewThreadReply(
                                                e.workbook,
                                                H.id,
                                                k,
                                                J,
                                                {
                                                  readOnly: e.readOnly,
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
                                              value: k,
                                              disabled: X$1,
                                              onChange: (event) =>
                                                F(event.target.value),
                                              lang: language("Reply"),
                                              "data-ui-label": "Reply",
                                            }),
                                          ],
                                        }),
                                        i.jsx("button", {
                                          className: "rv-primary",
                                          disabled: X$1 || !k.trim(),
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
                          q(
                            () =>
                              nQ(
                                e.workbook,
                                e.sheetId,
                                selectionAnchorAddress,
                                E,
                                J,
                                {
                                  readOnly: e.readOnly,
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
                              value: E,
                              disabled: disabled,
                              onChange: (event) => A(event.target.value),
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
                                E.length,
                                " / ",
                                reviewThreadLimits.textLength,
                              ],
                            }),
                            i.jsx("button", {
                              className: "rv-primary",
                              disabled: disabled || !E.trim(),
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
                  id: `${n}-hyperlink-panel`,
                  "aria-labelledby": `${n}-hyperlink-tab`,
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
                          q(
                            () =>
                              setCellHyperlink(
                                e.workbook,
                                e.sheetId,
                                selectionAnchorAddress,
                                j === "external"
                                  ? { kind: "external", url: url, label: N }
                                  : {
                                      kind: "internal",
                                      sheetId: sheetId,
                                      address: address,
                                      label: N,
                                    },
                                { readOnly: e.readOnly },
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
                              value: j,
                              disabled: disabled,
                              onChange: (event) => z(event.target.value),
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
                        j === "external"
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
                                  onChange: (event) => D(event.target.value),
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
                                        _(event.target.value),
                                      lang: language("Link worksheet"),
                                      "data-ui-label": "Link worksheet",
                                      children: e.workbook.sheets.map((sheet) =>
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
                                        T(event.target.value),
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
                              value: N,
                              disabled: disabled,
                              maxLength: reviewThreadLimits.labelLength,
                              onChange: (event) => R(event.target.value),
                              lang: language("Link display text"),
                              "data-ui-label": "Link display text",
                            }),
                          ],
                        }),
                        K &&
                          i.jsxs("div", {
                            className: "rv-link-preview",
                            "aria-label": text("Link preview"),
                            lang: language("Link preview"),
                            "data-ui-label": "Link preview",
                            children: [
                              i.jsx(Link2, { size: 15 }),
                              i.jsx("span", {
                                children:
                                  K.label ||
                                  (K.kind === "external"
                                    ? K.url
                                    : `${e.workbook.sheets.find((sheet) => sheet.id === K.sheetId)?.name ?? "Missing worksheet"}!${K.address}`),
                              }),
                              K.kind === "external"
                                ? i.jsxs("a", {
                                    href: K.url,
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
                                : e.onNavigate &&
                                  i.jsx("button", {
                                    type: "button",
                                    onClick: () => te(K.sheetId, K.address),
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
                              disabled: disabled || !l,
                              onClick: () =>
                                q(
                                  () =>
                                    setCellHyperlink(
                                      e.workbook,
                                      e.sheetId,
                                      selectionAnchorAddress,
                                      null,
                                      {
                                        readOnly: e.readOnly,
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
                                !(j === "external"
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
            I &&
              i.jsx("p", { role: "alert", className: "rv-error", children: I }),
          ],
        }),
      ],
    }),
  });
}
const formatCellValueLabel = (e) =>
    e === null ? "Blank" : typeof e == "string" ? JSON.stringify(e) : String(e),
  comparatorSymbols = { gt: ">", gte: "≥", lt: "<", lte: "≤", eq: "=" };

export {
  ReviewToolsPanel,
  comparatorSymbols,
  OutlineToolsPanel,
  formatCellValueLabel,
};
