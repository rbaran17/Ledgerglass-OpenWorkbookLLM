import i from "react/jsx-runtime";
import ne from "react";
import {
  Info,
  Check,
  Bell,
  UserRound,
  FileSpreadsheet,
  FolderClosed,
  LayoutGrid,
  ChevronRight,
  FolderOpen,
  PanelLeftClose,
  Search,
  ChevronDown,
  ArrowUpRight,
  PanelLeftOpen,
  Minimize2,
  Maximize2,
  X,
  BookOpen,
  ArrowLeft,
} from "lucide-react";
import Hq from "react-dom/client";
import { rn, t5 } from "./part-03.js";
import { BC } from "./part-11.js";
import { kz } from "./part-20.js";
import { FI } from "./part-24.js";
import { oN, yH } from "./part-54.js";
import { IH } from "./part-56.js";

const iEe = "embedded-excel-demo-v1",
  Wj = "embedded-excel-demo-v2:active-document",
  T_ = (e) => `embedded-excel-demo-v2:${e}`,
  xp = ["ledger", "draft"],
  Wf = { ledger: "General ledger", draft: "Draft workbook" };
function RS({ label: e, trigger: r, children: t, className: n = "" }) {
  const o = ne.useRef(null);
  return (
    ne.useEffect(() => {
      const a = (s) => {
        var d;
        (d = o.current) != null &&
          d.open &&
          s.target instanceof Node &&
          !o.current.contains(s.target) &&
          (o.current.open = false);
      };
      return (
        document.addEventListener("pointerdown", a),
        () => document.removeEventListener("pointerdown", a)
      );
    }, []),
    i.jsxs("details", {
      ref: o,
      className: `emb-header-menu ${n}`,
      onKeyDown: (a) => {
        var s;
        a.key === "Escape" &&
          a.currentTarget.open &&
          (a.preventDefault(),
          a.stopPropagation(),
          (a.currentTarget.open = false),
          (s = a.currentTarget.querySelector("summary")) == null || s.focus());
      },
      children: [
        i.jsx("summary", { "aria-label": e, title: e, children: r }),
        i.jsx("div", {
          className: "emb-menu-panel",
          role: "region",
          "aria-label": e,
          onClick: (a) => {
            var s, d;
            a.target.closest("button") &&
              (o.current && (o.current.open = false),
              (d =
                (s = o.current) == null
                  ? void 0
                  : s.querySelector("summary")) == null || d.focus());
          },
          children: t,
        }),
      ],
    })
  );
}
function _H() {
  const e = {
    version: 1,
    activeDocument: "ledger",
    documents: { ledger: FI(), draft: t5("Draft workbook") },
  };
  try {
    const r = localStorage.getItem(iEe);
    if (!r) return e;
    const t = JSON.parse(r);
    if (
      t.version !== 1 ||
      !t.documents ||
      !xp.includes(t.activeDocument) ||
      (t.versionStores !== void 0 &&
        (!t.versionStores ||
          typeof t.versionStores != "object" ||
          Array.isArray(t.versionStores) ||
          !Object.values(t.versionStores).every((n) => typeof n == "string")))
    )
      throw new Error("Invalid workspace");
    return {
      version: 1,
      legacy: !0,
      activeDocument: t.activeDocument,
      documents: {
        ledger: rn(t.documents.ledger),
        draft: rn(t.documents.draft),
      },
      versionStores: t.versionStores ?? {},
    };
  } catch {
    return {
      ...e,
      restoreProblem:
        "Saved workbooks could not be restored. Changes are kept in this tab. Use File → Save as Excel to keep a copy.",
    };
  }
}
function DH(e, r) {
  var s;
  const t = e.documents[r];
  let n = ((s = e.versionStores) == null ? void 0 : s[r]) ?? JSON.stringify(t);
  const o = BC(
      {
        getItem: () => n,
        setItem: (d, u) => {
          n = u;
        },
      },
      T_(r),
    ),
    a = o.load();
  return (
    JSON.stringify(a.workbook) !== JSON.stringify(t) &&
      o.save(t, a.view, a.guard),
    n
  );
}
function oEe(e, r) {
  let t = null;
  try {
    t = DH(e, r);
  } catch {}
  return BC(
    {
      getItem: () => t,
      setItem: (n, o) => {
        t = o;
      },
    },
    T_(r),
  );
}
async function aEe(e) {
  const r = await Promise.all(
    xp.map(async (t) => {
      const n = T_(t),
        o = yH(
          {
            getItem: (a) => localStorage.getItem(a),
            setItem: (a, s) => localStorage.setItem(a, s),
          },
          n,
        );
      try {
        if (e.legacy && localStorage.getItem(n) === null) {
          if (!navigator.locks)
            throw new Error(
              "This browser cannot coordinate safe workbook saves between tabs.",
            );
          const s = new AbortController(),
            d = setTimeout(() => s.abort(), 5e3);
          try {
            await navigator.locks.request(
              oN(n),
              { mode: "exclusive", signal: s.signal },
              (u) => {
                if (
                  s.signal.aborted ||
                  !u ||
                  u.mode !== "exclusive" ||
                  u.name !== oN(n)
                )
                  throw new Error(
                    "Safe workbook recovery could not acquire its document lock.",
                  );
                if (localStorage.getItem(n) !== null) return;
                const l = _H();
                if (l.restoreProblem || !l.legacy)
                  throw new Error(
                    "The original saved workspace changed during recovery.",
                  );
                localStorage.setItem(n, DH(l, t));
              },
            );
          } finally {
            clearTimeout(d);
          }
        }
        const a = o.load();
        if (!a && e.restoreProblem) throw new Error(e.restoreProblem);
        return {
          id: t,
          store: o,
          workbook: (a == null ? void 0 : a.workbook) ?? e.documents[t],
          problem: "",
        };
      } catch (a) {
        const s = `${a instanceof Error ? a.message : "Saved workbooks could not be restored."} Original saved data is preserved. Changes are kept in this tab; export a copy to keep them.`;
        return {
          id: t,
          store: oEe(e, t),
          workbook: e.documents[t],
          problem: s,
        };
      }
    }),
  );
  return {
    documents: Object.fromEntries(r.map(({ id: t, workbook: n }) => [t, n])),
    stores: Object.fromEntries(r.map(({ id: t, store: n }) => [t, n])),
    problems: Object.fromEntries(r.map(({ id: t, problem: n }) => [t, n])),
  };
}
function sEe() {
  var N;
  const [e] = ne.useState(_H),
    [r, t] = ne.useState(e.documents),
    [n, o] = ne.useState(null),
    [a, s] = ne.useState(() => {
      try {
        const R = sessionStorage.getItem(Wj);
        if (xp.includes(R)) return R;
      } catch {}
      return e.activeDocument;
    }),
    [d, u] = ne.useState([a]),
    [l, c] = ne.useState({}),
    h = ne.useMemo(
      () =>
        Object.fromEntries(
          xp.map((R) => [
            R,
            {
              onChange: (L) => t(($) => ({ ...$, [R]: L })),
              onWorkbookLoaded: (L) => t(($) => ({ ...$, [R]: L })),
              onSelectionChange: (L) => F(($) => ({ ...$, [R]: L })),
              onPersistenceStateChange: (L) => c(($) => ({ ...$, [R]: L })),
            },
          ]),
        ),
      [],
    ),
    [f, p] = ne.useState(true),
    [m, v] = ne.useState(false),
    [g, b] = ne.useState(true),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(false),
    [k, F] = ne.useState({}),
    I = l[a],
    O =
      (n == null ? void 0 : n.problems[a]) ||
      ((I == null ? void 0 : I.status) === "error" ||
      (I == null ? void 0 : I.status) === "conflict"
        ? (I.message ??
          "The workbook could not be saved. Changes remain in this tab.")
        : ""),
    j = !O && (I == null ? void 0 : I.status) === "saved",
    z = ne.useMemo(() => kz(), []),
    V = r[a],
    D = k[a]
      ? V.sheets.find((R) => R.id === k[a].sheetId)
      : V.sheets.find((R) => !R.hidden),
    B = ((N = k[a]) == null ? void 0 : N.range) ?? "A1",
    _ = `${(D == null ? void 0 : D.name) ?? "Sheet1"}!${B}`,
    C = xp.filter((R) =>
      `${Wf[R]} ${r[R].name}`.toLowerCase().includes(y.toLowerCase()),
    );
  (ne.useEffect(() => {
    let R = true;
    return (
      aEe(e).then((L) => {
        R && (t(L.documents), o(L));
      }),
      () => {
        R = false;
      }
    );
  }, [e]),
    ne.useEffect(() => {
      try {
        sessionStorage.setItem(Wj, a);
      } catch {}
    }, [a]),
    ne.useEffect(() => {
      if (!E) return;
      const R = (L) => {
        L.key === "Escape" && A(false);
      };
      return (
        document.addEventListener("keydown", R),
        () => document.removeEventListener("keydown", R)
      );
    }, [E]));
  const T = (R) => {
    (u((L) => (L.includes(R) ? L : [...L, R])), s(R), p(true));
  };
  return i.jsxs("div", {
    className: `embedded-demo ${g ? "" : "emb-documents-hidden"} ${m && f ? "emb-is-expanded" : ""}`,
    children: [
      i.jsxs("header", {
        className: "emb-topbar",
        children: [
          i.jsxs("div", {
            className: "emb-brand",
            "aria-label": "Ledgerglass | Workbooks",
            children: [
              i.jsxs("svg", {
                className: "emb-brand-mark",
                viewBox: "0 0 36 36",
                fill: "none",
                "aria-hidden": "true",
                children: [
                  i.jsx("path", {
                    d: "M8 4h12L8 16H2V10a6 6 0 0 1 6-6Z",
                    fill: "currentColor",
                  }),
                  i.jsx("path", {
                    d: "M22 4h6a6 6 0 0 1 6 6v6H10L22 4Z",
                    fill: "#B8ACFF",
                  }),
                  i.jsx("path", {
                    d: "M2 20h24L14 32H8a6 6 0 0 1-6-6v-6Z",
                    fill: "#B8ACFF",
                  }),
                  i.jsx("path", {
                    d: "M28 20h6v6a6 6 0 0 1-6 6H16l12-12Z",
                    fill: "currentColor",
                  }),
                ],
              }),
              i.jsx("span", {
                className: "emb-platform-name",
                children: "Ledgerglass",
              }),
              i.jsx("span", {
                className: "emb-brand-divider",
                "aria-hidden": "true",
              }),
              i.jsx("span", {
                className: "emb-application-name",
                children: "Workbooks",
              }),
            ],
          }),
          i.jsxs("div", {
            className: "emb-topbar-right",
            children: [
              i.jsx("span", {
                className: "emb-example-badge",
                children: "Local workspace",
              }),
              i.jsx("button", {
                type: "button",
                "aria-label": "About embedded workbook",
                title: "About this example",
                onClick: () => A(true),
                children: i.jsx(Info, { size: 17 }),
              }),
              i.jsxs(RS, {
                label: "Notifications",
                trigger: i.jsx(Bell, { size: 20, "aria-hidden": "true" }),
                children: [
                  i.jsx("h2", { children: "Workspace notifications" }),
                  i.jsxs("div", {
                    className: `emb-notification ${O ? "emb-notification-warning" : ""}`,
                    children: [
                      O
                        ? i.jsx(Info, { size: 19 })
                        : i.jsx(Check, { size: 19 }),
                      i.jsxs("div", {
                        children: [
                          i.jsx("strong", {
                            children: O
                              ? "Workbook needs attention"
                              : j
                                ? "Your workbook is saved"
                                : "Saving your workbook",
                          }),
                          i.jsx("p", {
                            children:
                              O ||
                              (j
                                ? `${Wf[a]} is saved in this browser.`
                                : "Changes are being saved in this browser."),
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "emb-menu-note",
                    children: "Status for this local workspace.",
                  }),
                ],
              }),
              i.jsxs(RS, {
                label: "Profile — Example preparer",
                className: "emb-profile-menu",
                trigger: i.jsx("span", {
                  className: "emb-avatar",
                  children: "EP",
                }),
                children: [
                  i.jsxs("div", {
                    className: "emb-profile-heading",
                    children: [
                      i.jsx(UserRound, { size: 22 }),
                      i.jsxs("div", {
                        children: [
                          i.jsx("strong", { children: "Example preparer" }),
                          i.jsx("p", {
                            children: "Local demonstration profile",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "emb-menu-note",
                    children:
                      "Workbooks stay in this browser. The assistant uses the local demo.",
                  }),
                  i.jsxs("button", {
                    type: "button",
                    onClick: () => A(true),
                    children: [
                      i.jsx(Info, { size: 16 }),
                      "About this workspace",
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      i.jsxs("nav", {
        className: "emb-secondary-nav",
        "aria-label": "Workspace navigation",
        children: [
          i.jsxs(RS, {
            label: "Workspace switcher",
            className: "emb-workspace-switcher",
            trigger: i.jsx(LayoutGrid, { size: 21, "aria-hidden": "true" }),
            children: [
              i.jsx("h2", { children: "Your workbooks" }),
              xp.map((R) =>
                i.jsxs(
                  "button",
                  {
                    type: "button",
                    onClick: () => T(R),
                    children: [i.jsx(FileSpreadsheet, { size: 18 }), Wf[R]],
                  },
                  R,
                ),
              ),
              i.jsxs("button", {
                type: "button",
                "aria-label": "Toggle document browser",
                "aria-pressed": g,
                onClick: () => b((R) => !R),
                children: [
                  i.jsx(FolderClosed, { size: 18 }),
                  g ? "Hide" : "Show",
                  " document browser",
                ],
              }),
            ],
          }),
          i.jsx("span", { className: "emb-navigation-divider" }),
          i.jsxs("div", {
            className: "emb-breadcrumb",
            children: [
              i.jsx("span", { children: "Workspace" }),
              i.jsx(ChevronRight, { size: 14, "aria-hidden": "true" }),
              i.jsx("span", { children: "Example engagement" }),
              i.jsx(ChevronRight, { size: 14, "aria-hidden": "true" }),
              i.jsx("strong", {
                "aria-current": "page",
                children: "Workpapers",
              }),
            ],
          }),
          i.jsxs("span", {
            className: "emb-nav-status",
            children: [
              i.jsx("span", { className: "emb-local-dot" }),
              "Local AI demo",
            ],
          }),
        ],
      }),
      i.jsxs("aside", {
        className: "emb-documents",
        "aria-label": "Engagement documents",
        hidden: !g,
        children: [
          i.jsxs("div", {
            className: "emb-engagement-heading",
            children: [
              i.jsx("div", {
                className: "emb-engagement-icon",
                children: i.jsx(FolderOpen, { size: 21 }),
              }),
              i.jsxs("div", {
                children: [
                  i.jsx("strong", { children: "Example engagement" }),
                  i.jsx("span", { children: "Financial statement audit" }),
                ],
              }),
            ],
          }),
          i.jsxs("div", {
            className: "emb-document-section",
            children: [
              i.jsx("span", { children: "WORKPAPERS" }),
              i.jsx("button", {
                type: "button",
                "aria-label": "Hide document browser",
                title: "Hide document browser",
                onClick: () => b(false),
                children: i.jsx(PanelLeftClose, { size: 15 }),
              }),
            ],
          }),
          i.jsxs("label", {
            className: "emb-search",
            children: [
              i.jsx(Search, { size: 14 }),
              i.jsx("input", {
                type: "search",
                "aria-label": "Search workpapers",
                placeholder: "Find a workpaper",
                value: y,
                onChange: (R) => w(R.target.value),
              }),
            ],
          }),
          i.jsxs("div", {
            className: "emb-folder-label",
            children: [
              i.jsx(ChevronDown, { size: 12 }),
              i.jsx(FolderOpen, { size: 14 }),
              i.jsx("span", { children: "Audit files" }),
              i.jsx("span", { children: xp.length }),
            ],
          }),
          i.jsxs("div", {
            className: "emb-document-list",
            children: [
              C.map((R) =>
                i.jsxs(
                  "button",
                  {
                    type: "button",
                    className: `emb-document ${a === R && f ? "emb-document-active" : ""}`,
                    "aria-label": `Open ${Wf[R]}`,
                    "aria-current": a === R && f ? "page" : void 0,
                    onClick: () => T(R),
                    children: [
                      i.jsx("div", {
                        className: "emb-document-icon",
                        children: i.jsx(FileSpreadsheet, { size: 17 }),
                      }),
                      i.jsxs("div", {
                        children: [
                          i.jsx("strong", { children: Wf[R] }),
                          i.jsx("span", {
                            children:
                              R === "ledger"
                                ? "Synthetic audit example"
                                : "Your working file",
                          }),
                        ],
                      }),
                      a === R &&
                        f &&
                        i.jsx("span", { className: "emb-document-current" }),
                    ],
                  },
                  R,
                ),
              ),
              C.length === 0 &&
                i.jsxs("p", {
                  className: "emb-no-documents",
                  children: ["No workpapers match “", y, "”."],
                }),
            ],
          }),
          i.jsxs("div", {
            className: "emb-document-bottom",
            children: [
              i.jsxs("div", {
                children: [
                  i.jsx("span", { className: "emb-local-dot" }),
                  "Local workspace",
                ],
              }),
              i.jsx("p", {
                children: O
                  ? "Changes are kept in this tab. Export a copy to keep them."
                  : "Workbook changes are saved in this browser.",
              }),
              i.jsxs("button", {
                type: "button",
                onClick: () => A(true),
                children: [
                  "About this example",
                  i.jsx(ArrowUpRight, { size: 12 }),
                ],
              }),
            ],
          }),
        ],
      }),
      i.jsxs("main", {
        className: "emb-workspace",
        children: [
          i.jsxs("div", {
            className: "emb-workspace-heading",
            children: [
              i.jsxs("div", {
                children: [
                  i.jsx("div", {
                    className: "emb-section-kicker",
                    children: "EXAMPLE ENGAGEMENT",
                  }),
                  i.jsx("h1", { children: "Workpapers" }),
                  i.jsx("p", {
                    children: "Review and prepare your audit evidence.",
                  }),
                ],
              }),
              i.jsxs("div", {
                className: "emb-heading-actions",
                children: [
                  !g &&
                    i.jsxs("button", {
                      type: "button",
                      className: "emb-show-documents",
                      onClick: () => b(true),
                      children: [
                        i.jsx(PanelLeftOpen, { size: 16 }),
                        i.jsx("span", { children: "Documents" }),
                      ],
                    }),
                  i.jsxs("span", {
                    className: "emb-workpaper-count",
                    children: [
                      i.jsx(FileSpreadsheet, { size: 14 }),
                      "2 workbooks",
                    ],
                  }),
                ],
              }),
            ],
          }),
          i.jsx("div", {
            className: "emb-mobile-documents",
            "aria-label": "Choose workpaper",
            children: xp.map((R) =>
              i.jsxs(
                "button",
                {
                  type: "button",
                  className: a === R && f ? "emb-mobile-active" : "",
                  "aria-label": `Open ${Wf[R]}`,
                  onClick: () => T(R),
                  children: [i.jsx(FileSpreadsheet, { size: 14 }), Wf[R]],
                },
                R,
              ),
            ),
          }),
          O &&
            i.jsxs("div", {
              className: "emb-save-warning",
              role: "status",
              "aria-label": "Workspace save status",
              children: [
                i.jsx(Info, { size: 15 }),
                i.jsx("span", { children: O }),
              ],
            }),
          i.jsxs("section", {
            className: `emb-workbook-window ${m ? "emb-window-expanded" : ""}`,
            "aria-label": "Embedded workbook window",
            hidden: !f,
            style: f ? void 0 : { display: "none" },
            children: [
              i.jsxs("div", {
                className: "emb-window-titlebar",
                children: [
                  i.jsxs("div", {
                    className: "emb-window-document",
                    children: [
                      i.jsx("div", {
                        className: "emb-window-file-icon",
                        children: i.jsx(FileSpreadsheet, { size: 17 }),
                      }),
                      i.jsxs("div", {
                        children: [
                          i.jsx("strong", { children: Wf[a] }),
                          i.jsx("span", { children: V.name }),
                        ],
                      }),
                      i.jsx("span", {
                        className: "emb-window-type",
                        children: "Workbook",
                      }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "emb-window-actions",
                    children: [
                      i.jsxs("span", {
                        className: `emb-save-state ${O ? "emb-save-state-warning" : ""}`,
                        role: "status",
                        children: [
                          j
                            ? i.jsx(Check, { size: 13 })
                            : i.jsx("span", { className: "emb-save-dot" }),
                          i.jsx("span", {
                            children: O
                              ? "Kept in this tab"
                              : j
                                ? "Saved in this browser"
                                : "Saving…",
                          }),
                        ],
                      }),
                      i.jsx("span", { className: "emb-window-action-divider" }),
                      i.jsx("button", {
                        type: "button",
                        "aria-label": m
                          ? "Restore workbook window"
                          : "Expand workbook window",
                        title: m
                          ? "Restore workbook window"
                          : "Expand workbook window",
                        onClick: () => v((R) => !R),
                        children: m
                          ? i.jsx(Minimize2, { size: 16 })
                          : i.jsx(Maximize2, { size: 16 }),
                      }),
                      i.jsx("button", {
                        type: "button",
                        "aria-label": "Close workbook window",
                        title: "Close workbook window",
                        onClick: () => {
                          (p(false), v(false));
                        },
                        children: i.jsx(X, { size: 17 }),
                      }),
                    ],
                  }),
                ],
              }),
              !n &&
                i.jsx("div", {
                  className: "emb-editor",
                  role: "status",
                  children: "Opening saved workbooks…",
                }),
              n &&
                d.map((R) =>
                  i.jsx(
                    "div",
                    {
                      className: "emb-editor",
                      hidden: R !== a,
                      style: R !== a ? { display: "none" } : void 0,
                      children: i.jsx(IH, {
                        initialWorkbook: n.documents[R],
                        ...h[R],
                        actor: "Example preparer",
                        adapter: z,
                        workbookStore: n.stores[R],
                        workbookStorePersistent: !n.problems[R],
                      }),
                    },
                    R,
                  ),
                ),
              i.jsxs("div", {
                className: "emb-host-footer",
                children: [
                  i.jsxs("div", {
                    children: [
                      i.jsx("span", { className: "emb-host-context-dot" }),
                      "Connected to workpaper",
                      i.jsx("span", {
                        className: "emb-footer-divider",
                        children: "/",
                      }),
                      i.jsx("span", {
                        className: "emb-host-citation",
                        title: `${Wf[a]} · ${_}`,
                        children: _,
                      }),
                    ],
                  }),
                  i.jsxs("span", {
                    children: [
                      V.sheets.filter((R) => !R.hidden).length,
                      " sheets",
                      i.jsx("span", {
                        className: "emb-footer-divider",
                        children: "·",
                      }),
                      "Revision",
                      " ",
                      V.revision,
                    ],
                  }),
                ],
              }),
            ],
          }),
          !f &&
            i.jsxs("section", {
              className: "emb-empty-window",
              "aria-label": "Closed workbook",
              children: [
                i.jsx("div", {
                  className: "emb-empty-icon",
                  children: i.jsx(FileSpreadsheet, { size: 31 }),
                }),
                i.jsx("h2", { children: "Workbook closed" }),
                i.jsxs("p", {
                  children: [
                    "Your changes remain in this workspace.",
                    i.jsx("br", {}),
                    "Open a workpaper to continue.",
                  ],
                }),
                i.jsxs("button", {
                  type: "button",
                  onClick: () => p(true),
                  children: [i.jsx(FolderOpen, { size: 16 }), "Reopen ", Wf[a]],
                }),
              ],
            }),
          i.jsxs("footer", {
            className: "emb-workspace-footer",
            children: [
              i.jsxs("span", {
                children: [
                  i.jsx(LayoutGrid, { size: 12 }),
                  "Embedded workbook editor",
                ],
              }),
              i.jsxs("span", {
                children: [
                  "Example data",
                  i.jsx("span", { children: "·" }),
                  "Local AI demo",
                ],
              }),
            ],
          }),
        ],
      }),
      E &&
        i.jsx("div", {
          className: "emb-modal-backdrop",
          onClick: () => A(false),
          children: i.jsxs("section", {
            className: "emb-info-dialog",
            role: "dialog",
            "aria-modal": "true",
            "aria-labelledby": "emb-info-title",
            onClick: (R) => R.stopPropagation(),
            onKeyDown: (R) => {
              if (R.key !== "Tab") return;
              const L = Array.from(
                  R.currentTarget.querySelectorAll("button:not(:disabled)"),
                ),
                $ = L[0],
                U = L.at(-1);
              (R.shiftKey &&
                document.activeElement === $ &&
                (R.preventDefault(), U == null || U.focus()),
                !R.shiftKey &&
                  document.activeElement === U &&
                  (R.preventDefault(), $ == null || $.focus()));
            },
            children: [
              i.jsx("button", {
                type: "button",
                className: "emb-info-close",
                "aria-label": "Close example details",
                autoFocus: true,
                onClick: () => A(false),
                children: i.jsx(X, { size: 18 }),
              }),
              i.jsx("div", {
                className: "emb-info-icon",
                children: i.jsx(BookOpen, { size: 25 }),
              }),
              i.jsx("p", {
                className: "emb-section-kicker",
                children: "INTEGRATION EXAMPLE",
              }),
              i.jsxs("h2", {
                id: "emb-info-title",
                children: [
                  "Your application.",
                  i.jsx("br", {}),
                  "The same workbook editor.",
                ],
              }),
              i.jsx("p", {
                children:
                  "This example hosts the spreadsheet in a document window. Switch workpapers, expand the editor, or close and reopen it. The host application retains each workbook and receives the selected range.",
              }),
              i.jsxs("div", {
                className: "emb-info-detail",
                children: [
                  i.jsx(Check, { size: 16 }),
                  i.jsx("span", {
                    children: O
                      ? "Changes are kept in this tab while browser storage is unavailable."
                      : "Changes are saved in this browser by the host application.",
                  }),
                ],
              }),
              i.jsxs("div", {
                className: "emb-info-detail",
                children: [
                  i.jsx(FileSpreadsheet, { size: 16 }),
                  i.jsx("span", {
                    children:
                      "The general ledger contains synthetic example data. The draft workbook starts blank.",
                  }),
                ],
              }),
              i.jsxs("div", {
                className: "emb-info-detail",
                children: [
                  i.jsx(Info, { size: 16 }),
                  i.jsx("span", {
                    children:
                      "AI Assistant uses the local demonstration adapter. No external AI provider is connected.",
                  }),
                ],
              }),
              i.jsxs("button", {
                type: "button",
                className: "emb-info-return",
                onClick: () => A(false),
                children: [
                  i.jsx(ArrowLeft, { size: 15 }),
                  "Back to workpapers",
                ],
              }),
            ],
          }),
        }),
    ],
  });
}
const k2 = new URLSearchParams(location.search);
Hq.createRoot(document.getElementById("root")).render(
  k2.has("embedded") || document.documentElement.dataset.mode === "embedded"
    ? i.jsx(sEe, {})
    : i.jsx(IH, {
        actor: "Preparer",
        initialWorkbook: k2.has("demo") ? FI() : void 0,
        persistenceKey:
          k2.has("temporary") || k2.has("demo") ? void 0 : "excel-workbook-v2",
      }),
);
