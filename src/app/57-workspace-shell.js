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
import {
  normalizeWorkbookInput,
  createBlankWorkbook,
} from "./03-workbook-model.js";
import { createLocalWorkbookStore } from "./11-local-workbook-store-and-versions.js";
import { buildLocalDemoProposal } from "./20-calculation-options-and-assistant.js";
import { buildSampleLedgerWorkbook } from "./24-merge-audit-and-ribbon-helpers.js";
import {
  workbookLockName,
  saveWorkbookWithLock,
} from "./54-workbook-save-conflicts.js";
import { WorkbookSyncBoundary } from "./56-main-worksheet-editor.js";

const legacyEmbeddedDemoKey = "embedded-excel-demo-v1",
  embeddedDemoActiveDocumentKey = "embedded-excel-demo-v2:active-document",
  buildEmbeddedDemoKey = (documentKind) =>
    `embedded-excel-demo-v2:${documentKind}`,
  workspaceDocumentKinds = ["ledger", "draft"],
  workspaceDocumentLabels = {
    ledger: "General ledger",
    draft: "Draft workbook",
  };
function EmbeddedHeaderMenu({ label, trigger, children, className = "" }) {
  const menuRef = ne.useRef(null);
  return (
    ne.useEffect(() => {
      const handlePointerDown = (event) => {
        menuRef.current?.open &&
          event.target instanceof Node &&
          !menuRef.current.contains(event.target) &&
          (menuRef.current.open = false);
      };
      return (
        document.addEventListener("pointerdown", handlePointerDown),
        () => document.removeEventListener("pointerdown", handlePointerDown)
      );
    }, []),
    i.jsxs("details", {
      ref: menuRef,
      className: `emb-header-menu ${className}`,
      onKeyDown: (event) => {
        event.key === "Escape" &&
          event.currentTarget.open &&
          (event.preventDefault(),
          event.stopPropagation(),
          (event.currentTarget.open = false),
          event.currentTarget.querySelector("summary")?.focus());
      },
      children: [
        i.jsx("summary", {
          "aria-label": label,
          title: label,
          children: trigger,
        }),
        i.jsx("div", {
          className: "emb-menu-panel",
          role: "region",
          "aria-label": label,
          onClick: (event) => {
            event.target.closest("button") &&
              (menuRef.current && (menuRef.current.open = false),
              menuRef.current?.querySelector("summary")?.focus());
          },
          children,
        }),
      ],
    })
  );
}
function restoreSavedWorkspaceDocuments() {
  const defaultWorkspace = {
    version: 1,
    activeDocument: "ledger",
    documents: {
      ledger: buildSampleLedgerWorkbook(),
      draft: createBlankWorkbook("Draft workbook"),
    },
  };
  try {
    const storedValue = localStorage.getItem(legacyEmbeddedDemoKey);
    if (!storedValue) return defaultWorkspace;
    const storedWorkspace = JSON.parse(storedValue);
    if (
      storedWorkspace.version !== 1 ||
      !storedWorkspace.documents ||
      !workspaceDocumentKinds.includes(storedWorkspace.activeDocument) ||
      (storedWorkspace.versionStores !== undefined &&
        (!storedWorkspace.versionStores ||
          typeof storedWorkspace.versionStores != "object" ||
          Array.isArray(storedWorkspace.versionStores) ||
          !Object.values(storedWorkspace.versionStores).every(
            (versionStore) => typeof versionStore == "string",
          )))
    )
      throw new Error("Invalid workspace");
    return {
      version: 1,
      legacy: true,
      activeDocument: storedWorkspace.activeDocument,
      documents: {
        ledger: normalizeWorkbookInput(storedWorkspace.documents.ledger),
        draft: normalizeWorkbookInput(storedWorkspace.documents.draft),
      },
      versionStores: storedWorkspace.versionStores ?? {},
    };
  } catch {
    return {
      ...defaultWorkspace,
      restoreProblem:
        "Saved workbooks could not be restored. Changes are kept in this tab. Use File → Save as Excel to keep a copy.",
    };
  }
}
function persistWorkspaceDocumentVersion(workspace, workspaceDocumentKind) {
  const workbook = workspace.documents[workspaceDocumentKind];
  let versionStore =
    workspace.versionStores?.[workspaceDocumentKind] ??
    JSON.stringify(workbook);
  const localWorkbookStore = createLocalWorkbookStore(
      {
        getItem: () => versionStore,
        setItem: (storageKey, serializedValue) => {
          versionStore = serializedValue;
        },
      },
      buildEmbeddedDemoKey(workspaceDocumentKind),
    ),
    loadedState = localWorkbookStore.load();
  return (
    JSON.stringify(loadedState.workbook) !== JSON.stringify(workbook) &&
      localWorkbookStore.save(workbook, loadedState.view, loadedState.guard),
    versionStore
  );
}
function createWorkspaceDocumentStore(workspace, workspaceDocumentKind) {
  let storedVersion = null;
  try {
    storedVersion = persistWorkspaceDocumentVersion(
      workspace,
      workspaceDocumentKind,
    );
  } catch {}
  return createLocalWorkbookStore(
    {
      getItem: () => storedVersion,
      setItem: (storageKey, serializedValue) => {
        storedVersion = serializedValue;
      },
    },
    buildEmbeddedDemoKey(workspaceDocumentKind),
  );
}
async function recoverSavedWorkspaceDocuments(restoredWorkspace) {
  const results = await Promise.all(
    workspaceDocumentKinds.map(async (workspaceDocumentKind) => {
      const embeddedDemoKey = buildEmbeddedDemoKey(workspaceDocumentKind),
        store = saveWorkbookWithLock(
          {
            getItem: (storageKey) => localStorage.getItem(storageKey),
            setItem: (writeKey, serializedValue) =>
              localStorage.setItem(writeKey, serializedValue),
          },
          embeddedDemoKey,
        );
      try {
        if (
          restoredWorkspace.legacy &&
          localStorage.getItem(embeddedDemoKey) === null
        ) {
          if (!navigator.locks)
            throw new Error(
              "This browser cannot coordinate safe workbook saves between tabs.",
            );
          const abortController = new AbortController(),
            abortTimer = setTimeout(() => abortController.abort(), 5e3);
          try {
            await navigator.locks.request(
              workbookLockName(embeddedDemoKey),
              { mode: "exclusive", signal: abortController.signal },
              (lock) => {
                if (
                  abortController.signal.aborted ||
                  !lock ||
                  lock.mode !== "exclusive" ||
                  lock.name !== workbookLockName(embeddedDemoKey)
                )
                  throw new Error(
                    "Safe workbook recovery could not acquire its document lock.",
                  );
                if (localStorage.getItem(embeddedDemoKey) !== null) return;
                const freshWorkspace = restoreSavedWorkspaceDocuments();
                if (freshWorkspace.restoreProblem || !freshWorkspace.legacy)
                  throw new Error(
                    "The original saved workspace changed during recovery.",
                  );
                localStorage.setItem(
                  embeddedDemoKey,
                  persistWorkspaceDocumentVersion(
                    freshWorkspace,
                    workspaceDocumentKind,
                  ),
                );
              },
            );
          } finally {
            clearTimeout(abortTimer);
          }
        }
        const loadResult = store.load();
        if (!loadResult && restoredWorkspace.restoreProblem)
          throw new Error(restoredWorkspace.restoreProblem);
        return {
          id: workspaceDocumentKind,
          store: store,
          workbook:
            loadResult?.workbook ??
            restoredWorkspace.documents[workspaceDocumentKind],
          problem: "",
        };
      } catch (error) {
        const problem = `${error instanceof Error ? error.message : "Saved workbooks could not be restored."} Original saved data is preserved. Changes are kept in this tab; export a copy to keep them.`;
        return {
          id: workspaceDocumentKind,
          store: createWorkspaceDocumentStore(
            restoredWorkspace,
            workspaceDocumentKind,
          ),
          workbook: restoredWorkspace.documents[workspaceDocumentKind],
          problem: problem,
        };
      }
    }),
  );
  return {
    documents: Object.fromEntries(
      results.map(({ id: documentKind, workbook }) => [documentKind, workbook]),
    ),
    stores: Object.fromEntries(
      results.map(({ id: storeDocumentKind, store }) => [
        storeDocumentKind,
        store,
      ]),
    ),
    problems: Object.fromEntries(
      results.map(({ id: problemDocumentKind, problem }) => [
        problemDocumentKind,
        problem,
      ]),
    ),
  };
}
function EmbeddedWorkbookWorkspace() {
  const [savedWorkspace] = ne.useState(restoreSavedWorkspaceDocuments),
    [documents, setDocuments] = ne.useState(savedWorkspace.documents),
    [recoveryResult, setRecoveryResult] = ne.useState(null),
    [activeDocument, setActiveDocument] = ne.useState(() => {
      try {
        const storedActiveDocument = sessionStorage.getItem(
          embeddedDemoActiveDocumentKey,
        );
        if (workspaceDocumentKinds.includes(storedActiveDocument))
          return storedActiveDocument;
      } catch {}
      return savedWorkspace.activeDocument;
    }),
    [openedDocuments, setOpenedDocuments] = ne.useState([activeDocument]),
    [persistenceStates, setPersistenceStates] = ne.useState({}),
    documentHandlers = ne.useMemo(
      () =>
        Object.fromEntries(
          workspaceDocumentKinds.map((workspaceDocumentKind) => [
            workspaceDocumentKind,
            {
              onChange: (nextWorkbook) =>
                setDocuments((previousDocuments) => ({
                  ...previousDocuments,
                  [workspaceDocumentKind]: nextWorkbook,
                })),
              onWorkbookLoaded: (loadedWorkbook) =>
                setDocuments((documentsBeforeLoad) => ({
                  ...documentsBeforeLoad,
                  [workspaceDocumentKind]: loadedWorkbook,
                })),
              onSelectionChange: (nextSelection) =>
                setSelections((previousSelections) => ({
                  ...previousSelections,
                  [workspaceDocumentKind]: nextSelection,
                })),
              onPersistenceStateChange: (nextPersistenceState) =>
                setPersistenceStates((previousPersistenceStates) => ({
                  ...previousPersistenceStates,
                  [workspaceDocumentKind]: nextPersistenceState,
                })),
            },
          ]),
        ),
      [],
    ),
    [isWorkbookWindowOpen, setIsWorkbookWindowOpen] = ne.useState(true),
    [isWorkbookWindowExpanded, setIsWorkbookWindowExpanded] =
      ne.useState(false),
    [isDocumentBrowserOpen, setIsDocumentBrowserOpen] = ne.useState(true),
    [documentFilter, setDocumentFilter] = ne.useState(""),
    [isAboutOpen, setIsAboutOpen] = ne.useState(false),
    [selections, setSelections] = ne.useState({}),
    activePersistence = persistenceStates[activeDocument],
    activeProblem =
      recoveryResult?.problems[activeDocument] ||
      (activePersistence?.status === "error" ||
      activePersistence?.status === "conflict"
        ? (activePersistence.message ??
          "The workbook could not be saved. Changes remain in this tab.")
        : ""),
    isSaved = !activeProblem && activePersistence?.status === "saved",
    adapter = ne.useMemo(() => buildLocalDemoProposal(), []),
    workbook = documents[activeDocument],
    activeSheet = selections[activeDocument]
      ? workbook.sheets.find(
          (sheet) => sheet.id === selections[activeDocument].sheetId,
        )
      : workbook.sheets.find((sheet) => !sheet.hidden),
    activeRange = selections[activeDocument]?.range ?? "A1",
    activeAddress = `${activeSheet?.name ?? "Sheet1"}!${activeRange}`,
    filteredDocumentKinds = workspaceDocumentKinds.filter(
      (workspaceDocumentKind) =>
        `${workspaceDocumentLabels[workspaceDocumentKind]} ${documents[workspaceDocumentKind].name}`
          .toLowerCase()
          .includes(documentFilter.toLowerCase()),
    );
  (ne.useEffect(() => {
    let isMounted = true;
    return (
      recoverSavedWorkspaceDocuments(savedWorkspace).then((recovered) => {
        isMounted &&
          (setDocuments(recovered.documents), setRecoveryResult(recovered));
      }),
      () => {
        isMounted = false;
      }
    );
  }, [savedWorkspace]),
    ne.useEffect(() => {
      try {
        sessionStorage.setItem(embeddedDemoActiveDocumentKey, activeDocument);
      } catch {}
    }, [activeDocument]),
    ne.useEffect(() => {
      if (!isAboutOpen) return;
      const handleEscapeKey = (event) => {
        event.key === "Escape" && setIsAboutOpen(false);
      };
      return (
        document.addEventListener("keydown", handleEscapeKey),
        () => document.removeEventListener("keydown", handleEscapeKey)
      );
    }, [isAboutOpen]));
  const openDocument = (nextDocumentKind) => {
    (setOpenedDocuments((previousOpened) =>
      previousOpened.includes(nextDocumentKind)
        ? previousOpened
        : [...previousOpened, nextDocumentKind],
    ),
      setActiveDocument(nextDocumentKind),
      setIsWorkbookWindowOpen(true));
  };
  return i.jsxs("div", {
    className: `embedded-demo ${isDocumentBrowserOpen ? "" : "emb-documents-hidden"} ${isWorkbookWindowExpanded && isWorkbookWindowOpen ? "emb-is-expanded" : ""}`,
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
                onClick: () => setIsAboutOpen(true),
                children: i.jsx(Info, { size: 17 }),
              }),
              i.jsxs(EmbeddedHeaderMenu, {
                label: "Notifications",
                trigger: i.jsx(Bell, { size: 20, "aria-hidden": "true" }),
                children: [
                  i.jsx("h2", { children: "Workspace notifications" }),
                  i.jsxs("div", {
                    className: `emb-notification ${activeProblem ? "emb-notification-warning" : ""}`,
                    children: [
                      activeProblem
                        ? i.jsx(Info, { size: 19 })
                        : i.jsx(Check, { size: 19 }),
                      i.jsxs("div", {
                        children: [
                          i.jsx("strong", {
                            children: activeProblem
                              ? "Workbook needs attention"
                              : isSaved
                                ? "Your workbook is saved"
                                : "Saving your workbook",
                          }),
                          i.jsx("p", {
                            children:
                              activeProblem ||
                              (isSaved
                                ? `${workspaceDocumentLabels[activeDocument]} is saved in this browser.`
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
              i.jsxs(EmbeddedHeaderMenu, {
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
                    onClick: () => setIsAboutOpen(true),
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
          i.jsxs(EmbeddedHeaderMenu, {
            label: "Workspace switcher",
            className: "emb-workspace-switcher",
            trigger: i.jsx(LayoutGrid, { size: 21, "aria-hidden": "true" }),
            children: [
              i.jsx("h2", { children: "Your workbooks" }),
              workspaceDocumentKinds.map((workspaceDocumentKind) =>
                i.jsxs(
                  "button",
                  {
                    type: "button",
                    onClick: () => openDocument(workspaceDocumentKind),
                    children: [
                      i.jsx(FileSpreadsheet, { size: 18 }),
                      workspaceDocumentLabels[workspaceDocumentKind],
                    ],
                  },
                  workspaceDocumentKind,
                ),
              ),
              i.jsxs("button", {
                type: "button",
                "aria-label": "Toggle document browser",
                "aria-pressed": isDocumentBrowserOpen,
                onClick: () =>
                  setIsDocumentBrowserOpen(
                    (isPreviouslyOpen) => !isPreviouslyOpen,
                  ),
                children: [
                  i.jsx(FolderClosed, { size: 18 }),
                  isDocumentBrowserOpen ? "Hide" : "Show",
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
        hidden: !isDocumentBrowserOpen,
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
                onClick: () => setIsDocumentBrowserOpen(false),
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
                value: documentFilter,
                onChange: (event) => setDocumentFilter(event.target.value),
              }),
            ],
          }),
          i.jsxs("div", {
            className: "emb-folder-label",
            children: [
              i.jsx(ChevronDown, { size: 12 }),
              i.jsx(FolderOpen, { size: 14 }),
              i.jsx("span", { children: "Audit files" }),
              i.jsx("span", { children: workspaceDocumentKinds.length }),
            ],
          }),
          i.jsxs("div", {
            className: "emb-document-list",
            children: [
              filteredDocumentKinds.map((documentKind) =>
                i.jsxs(
                  "button",
                  {
                    type: "button",
                    className: `emb-document ${activeDocument === documentKind && isWorkbookWindowOpen ? "emb-document-active" : ""}`,
                    "aria-label": `Open ${workspaceDocumentLabels[documentKind]}`,
                    "aria-current":
                      activeDocument === documentKind && isWorkbookWindowOpen
                        ? "page"
                        : undefined,
                    onClick: () => openDocument(documentKind),
                    children: [
                      i.jsx("div", {
                        className: "emb-document-icon",
                        children: i.jsx(FileSpreadsheet, { size: 17 }),
                      }),
                      i.jsxs("div", {
                        children: [
                          i.jsx("strong", {
                            children: workspaceDocumentLabels[documentKind],
                          }),
                          i.jsx("span", {
                            children:
                              documentKind === "ledger"
                                ? "Synthetic audit example"
                                : "Your working file",
                          }),
                        ],
                      }),
                      activeDocument === documentKind &&
                        isWorkbookWindowOpen &&
                        i.jsx("span", { className: "emb-document-current" }),
                    ],
                  },
                  documentKind,
                ),
              ),
              filteredDocumentKinds.length === 0 &&
                i.jsxs("p", {
                  className: "emb-no-documents",
                  children: ["No workpapers match “", documentFilter, "”."],
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
                children: activeProblem
                  ? "Changes are kept in this tab. Export a copy to keep them."
                  : "Workbook changes are saved in this browser.",
              }),
              i.jsxs("button", {
                type: "button",
                onClick: () => setIsAboutOpen(true),
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
                  !isDocumentBrowserOpen &&
                    i.jsxs("button", {
                      type: "button",
                      className: "emb-show-documents",
                      onClick: () => setIsDocumentBrowserOpen(true),
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
            children: workspaceDocumentKinds.map((workspaceDocumentKind) =>
              i.jsxs(
                "button",
                {
                  type: "button",
                  className:
                    activeDocument === workspaceDocumentKind &&
                    isWorkbookWindowOpen
                      ? "emb-mobile-active"
                      : "",
                  "aria-label": `Open ${workspaceDocumentLabels[workspaceDocumentKind]}`,
                  onClick: () => openDocument(workspaceDocumentKind),
                  children: [
                    i.jsx(FileSpreadsheet, { size: 14 }),
                    workspaceDocumentLabels[workspaceDocumentKind],
                  ],
                },
                workspaceDocumentKind,
              ),
            ),
          }),
          activeProblem &&
            i.jsxs("div", {
              className: "emb-save-warning",
              role: "status",
              "aria-label": "Workspace save status",
              children: [
                i.jsx(Info, { size: 15 }),
                i.jsx("span", { children: activeProblem }),
              ],
            }),
          i.jsxs("section", {
            className: `emb-workbook-window ${isWorkbookWindowExpanded ? "emb-window-expanded" : ""}`,
            "aria-label": "Embedded workbook window",
            hidden: !isWorkbookWindowOpen,
            style: isWorkbookWindowOpen ? undefined : { display: "none" },
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
                          i.jsx("strong", {
                            children: workspaceDocumentLabels[activeDocument],
                          }),
                          i.jsx("span", { children: workbook.name }),
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
                        className: `emb-save-state ${activeProblem ? "emb-save-state-warning" : ""}`,
                        role: "status",
                        children: [
                          isSaved
                            ? i.jsx(Check, { size: 13 })
                            : i.jsx("span", { className: "emb-save-dot" }),
                          i.jsx("span", {
                            children: activeProblem
                              ? "Kept in this tab"
                              : isSaved
                                ? "Saved in this browser"
                                : "Saving…",
                          }),
                        ],
                      }),
                      i.jsx("span", { className: "emb-window-action-divider" }),
                      i.jsx("button", {
                        type: "button",
                        "aria-label": isWorkbookWindowExpanded
                          ? "Restore workbook window"
                          : "Expand workbook window",
                        title: isWorkbookWindowExpanded
                          ? "Restore workbook window"
                          : "Expand workbook window",
                        onClick: () =>
                          setIsWorkbookWindowExpanded(
                            (isPreviouslyExpanded) => !isPreviouslyExpanded,
                          ),
                        children: isWorkbookWindowExpanded
                          ? i.jsx(Minimize2, { size: 16 })
                          : i.jsx(Maximize2, { size: 16 }),
                      }),
                      i.jsx("button", {
                        type: "button",
                        "aria-label": "Close workbook window",
                        title: "Close workbook window",
                        onClick: () => {
                          (setIsWorkbookWindowOpen(false),
                            setIsWorkbookWindowExpanded(false));
                        },
                        children: i.jsx(X, { size: 17 }),
                      }),
                    ],
                  }),
                ],
              }),
              !recoveryResult &&
                i.jsx("div", {
                  className: "emb-editor",
                  role: "status",
                  children: "Opening saved workbooks…",
                }),
              recoveryResult &&
                openedDocuments.map((openedDocumentKind) =>
                  i.jsx(
                    "div",
                    {
                      className: "emb-editor",
                      hidden: openedDocumentKind !== activeDocument,
                      style:
                        openedDocumentKind !== activeDocument
                          ? { display: "none" }
                          : undefined,
                      children: i.jsx(WorkbookSyncBoundary, {
                        initialWorkbook:
                          recoveryResult.documents[openedDocumentKind],
                        ...documentHandlers[openedDocumentKind],
                        actor: "Example preparer",
                        adapter: adapter,
                        workbookStore:
                          recoveryResult.stores[openedDocumentKind],
                        workbookStorePersistent:
                          !recoveryResult.problems[openedDocumentKind],
                      }),
                    },
                    openedDocumentKind,
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
                        title: `${workspaceDocumentLabels[activeDocument]} · ${activeAddress}`,
                        children: activeAddress,
                      }),
                    ],
                  }),
                  i.jsxs("span", {
                    children: [
                      workbook.sheets.filter((sheet) => !sheet.hidden).length,
                      " sheets",
                      i.jsx("span", {
                        className: "emb-footer-divider",
                        children: "·",
                      }),
                      "Revision",
                      " ",
                      workbook.revision,
                    ],
                  }),
                ],
              }),
            ],
          }),
          !isWorkbookWindowOpen &&
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
                  onClick: () => setIsWorkbookWindowOpen(true),
                  children: [
                    i.jsx(FolderOpen, { size: 16 }),
                    "Reopen ",
                    workspaceDocumentLabels[activeDocument],
                  ],
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
      isAboutOpen &&
        i.jsx("div", {
          className: "emb-modal-backdrop",
          onClick: () => setIsAboutOpen(false),
          children: i.jsxs("section", {
            className: "emb-info-dialog",
            role: "dialog",
            "aria-modal": "true",
            "aria-labelledby": "emb-info-title",
            onClick: (event) => event.stopPropagation(),
            onKeyDown: (event) => {
              if (event.key !== "Tab") return;
              const focusableButtons = Array.from(
                  event.currentTarget.querySelectorAll("button:not(:disabled)"),
                ),
                firstButton = focusableButtons[0],
                lastButton = focusableButtons.at(-1);
              (event.shiftKey &&
                document.activeElement === firstButton &&
                (event.preventDefault(), lastButton?.focus()),
                !event.shiftKey &&
                  document.activeElement === lastButton &&
                  (event.preventDefault(), firstButton?.focus()));
            },
            children: [
              i.jsx("button", {
                type: "button",
                className: "emb-info-close",
                "aria-label": "Close example details",
                autoFocus: true,
                onClick: () => setIsAboutOpen(false),
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
                    children: activeProblem
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
                onClick: () => setIsAboutOpen(false),
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
const pageUrlParams = new URLSearchParams(location.search);
Hq.createRoot(document.getElementById("root")).render(
  pageUrlParams.has("embedded") ||
    document.documentElement.dataset.mode === "embedded"
    ? i.jsx(EmbeddedWorkbookWorkspace, {})
    : i.jsx(WorkbookSyncBoundary, {
        actor: "Preparer",
        initialWorkbook: pageUrlParams.has("demo")
          ? buildSampleLedgerWorkbook()
          : undefined,
        persistenceKey:
          pageUrlParams.has("temporary") || pageUrlParams.has("demo")
            ? undefined
            : "excel-workbook-v2",
      }),
);
