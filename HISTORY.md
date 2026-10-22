# Ledgerglass

## Docker and scalable deployment

The current handoff includes Docker web and stdio MCP targets, a Compose load balancer, Kubernetes scaling templates, runtime checks and CI verification. Start with [the Docker handoff](docs/DOCKER-AND-SCALING.md). Existing workbook features and coverage statuses are unchanged. Docker/cluster execution is explicitly unverified on the preparation machine; see docs/DOCKER-PACKAGING-VERIFICATION.json.


## AW-052 Refresh All — current checkpoint (2026-09-19)

Select connected worksheets, compare their source identities, timestamps, row counts and controls, inspect every record, and apply the reviewed batch as one Undo/Redo action. Every source is reauthorized; one failure retains all prior populations. Cancellation, stale host contexts, batch limits and slow backend request phases are handled explicitly. Single-source connections and existing connector families remain available.

**8,734 tests passed across 326 files**, including 52 added checks, with zero failures/skips. Both portable applications, the compiled SDK/workers and all 49 validation hosts build. Eight browser observations cover review, rows/focus, atomic Apply, Undo/Redo, expired access, cancellation and successful retry. Actual local HTTP/SQLite/file and worker checks passed; the browser practice host uses synthetic in-memory sources.

AW-052 remains **Partial** for native ODBC, real vendor/production identity/secret/tenant configurations and supported-host acceptance. Scheduled/incremental refresh, dependency refresh graphs and XLSX connection exchange remain outside scope. Base coverage remains **70 accepted / 5 Partial / 1 Planned**. Sixteen new manual templates bring the register to **3,167**: 3,025 Not run, 51 blocked unimplemented, 91 blocked external and zero manually Passed.

[Changes, integration and acceptance](docs/AW-052-REFRESH-ALL.md) · [Executed evidence](docs/AW-052-REFRESH-ALL-VERIFICATION.json)

## AW-101 accessibility and localization — previous checkpoint (2026-09-19)

Shared localized numeric entry now covers **40 control sites across 23 components**. Required blank/invalid values do not become zero; number edits preserve precision and the in-progress locale. Scenario lists support decimal commas. Modal focus handles hidden/disabled controls and nested dialogs. German chat/numeric labels and placeholder-safe translations are included; long ribbon labels wrap without overlap. The German starter catalog has 301 entries; complete advanced-tool translation remains open.

**8,682 tests pass across 323 files**, including 42 new checks, with zero failures/skips. The final numeric validation guard also passed 68 focused checks across three files before the final rebuild. Both portable applications and all **49 validation hosts** build. Nine browser checks and four axe scans report zero violations; incomplete contrast checks remain in the enlarged/whole-workbook views (50, 32 and 34 nodes), so this is not independent accessibility certification.

AW-101 remains **Partial** for actual screen-reader journeys, native zoom/reflow and device modes, complete localization and independent WCAG review. Original criteria and prior features are preserved. Base coverage stays **70 accepted / 5 Partial / 1 Planned**. Sixteen new manual cases bring the register to **3,151** (3,009 Not run, 51 blocked unimplemented, 91 blocked external, zero manually Passed).

[Changes and remaining acceptance](docs/AW-101-FOLLOW-UP.md) · [Executed evidence](docs/AW-101-FOLLOW-UP-VERIFICATION.json)

## Assistant chat — previous checkpoint (2026-09-19)

The panel uses **assistant-ui 0.15.21 (MIT)** for its React thread, messages, viewport and composer. It adds conversation history, file selection/drop/paste, text and image previews, removal, Enter/Shift+Enter, Stop, Retry, Regenerate, Copy and New chat. Chat modules load only when opened; standalone exports include them in the single HTML file. Existing selected-cell disclosure, cited answers, reviewed proposals, undo and MCP integration remain.

**8,640 tests across 321 files pass**, with no failures or skips, including 32 new attachment/chat checks. All **49 validation hosts** build. 10 observed browser journeys and 2 full-page chat scans are recorded; the scans have zero violations and 43 unresolved contrast checks each, mostly existing clipped workbook controls. Those unresolved checks are not accessibility certification.

The local demo acknowledges files but does not interpret documents with a model. The host can enable history and attachments with `createHttpAdapter({ endpoint, chat: true })`; legacy adapters remain compatible. Actual provider document interpretation, production authorization/retention, token streaming and independent assistive-technology acceptance remain outside this local verification.

CHAT-001–012 add manual validation templates: **3,135 total** (2,997 Not run, 51 Blocked unimplemented, 87 Blocked external, zero Passed). All 3,123 prior cases are preserved. Base coverage remains **70 accepted / 5 Partial / 1 Planned**; no requirement was promoted by this UI change. Original coverage workbook and historical proofs are unchanged.

[Chat integration guide](docs/CHAT-UI-INTEGRATION.md) · [Executed evidence](docs/CHAT-UI-VERIFICATION.json)

## Usability and accessibility — previous checkpoint (2026-09-18)

**8,608 tests across 319 files pass**, zero failures or skips, including 70 new formula-entry and accessibility checks. All **49 validation hosts** build. The compiled browser review confirms mouse range insertion from the formula bar and in-cell editor, keyboard reference selection, reference locking, protected cross-sheet sources, commit/cancel/undo, completion scrolling and dialog focus return. Recorded browser scans have no serious/critical violations; incomplete contrast items remain explicit review work.

**AW-101 remains Partial.** Actual VoiceOver/NVDA journeys, native browser 200% zoom and device/browser matrices, complete locale coverage and independent WCAG 2.2 AA acceptance remain open. This is a usability improvement with bounded local evidence, not identical Excel accessibility. The 200% CSS layout check is not native browser zoom. External model/production acceptance is not claimed.

All prior functionality, original requirements and criteria, local closure records and historical proofs remain. UX-001–012 add independent manual templates: **3,123 total**, with 2,987 Not run, 51 Blocked unimplemented, 85 Blocked external and zero Passed. Base coverage remains **70 accepted / 5 Partial / 1 Planned — 6 open of 76**. Full coverage remains **78 accepted / 19 Partial / 11 Planned — 30 open of 108**.

[Review and remaining gaps](docs/USABILITY-ACCESSIBILITY-REVIEW.md) · [Executed evidence](docs/USABILITY-VERIFICATION.json).

**Clara interface update:** coordinated navy/indigo styling, lavender-gray borders, clearer focus states and a responsive Clara example host with workbook switcher, save notifications and profile details. All implemented features and workbook formatting are retained. See [theme choices and verification](docs/CLARA-THEME.md).

A portable Excel-style workbook and embeddable React component with a provider-neutral LLM proposal interface. The worksheet fills the component: ribbon, formula bar, grid, sheet tabs and status bar, with no branded application navigation or dashboard.

**This is a working initial application, not full Microsoft Excel parity.** The current coverage sheet tracks 76 base Excel capabilities and 152 acceptance criteria. The full 108-requirement product register is retained separately, including audit and AI extensions; all implemented features remain available. The editor now includes 395 formula functions, tables, grouped PivotTable summaries, charts, pictures, shapes, page setup, multi-area selection, advanced Find/Replace, outlines and threaded review comments. Complete Excel conformance, Power Query/M, DAX, VBA, coauthoring and production authorization remain outside the delivered subset.

**Base Excel coverage: 70 locally accepted, 5 Partial and 1 Planned requirements — 6 open out of 76.** AW-063 adds reusable report snapshots, reviewed linked refresh and inert embedded-object inventories with explicit export losses. [Report acceptance](docs/REPORT-OUTPUT-ACCEPTANCE.md) records scope and limits. AW-061 provides selected bounded Solver and statistical tools with reviewed reports and trusted fresh results. [Analysis acceptance](docs/ANALYSIS-ACCEPTANCE.md) records the scope and limits. AW-055 provides a separately evaluated 20-function DAX subset with calculated columns, model measures, context transition, calendar-year YTD and trusted current-result reads. [DAX acceptance](docs/DAX-ACCEPTANCE.md) records the local closure and limits. AW-054 provides reviewed typed relational models, single-direction relationships and independent multi-fact measures under Data → Data Model. [Model acceptance](docs/DATA-MODEL-ACCEPTANCE.md) records the local closure and limits. AW-052 includes concrete server-side JSON/OData/ERP HTTP, SQLite and CSV/TSV folder loaders plus a node-odbc-compatible cursor adapter, with reviewed refresh and last-success recovery. It remains Partial pending native ODBC, production authorization and actual vendor/host acceptance. AW-103 reliability and recovery still awaits deployed monitoring, alerts and durable restore acceptance. Earlier AW-102 performance and AW-101 accessibility gates remain open. The [full product register](docs/ALL-FEATURE-REGISTER.csv) retains all 108 requirements (78 accepted, 19 Partial, 11 Planned), including implemented audit and AI features. Read the [developer handoff](docs/DEVELOPER-HANDOFF.md), [current coverage](docs/COVERAGE-SUMMARY.md) and [accessibility and localization acceptance story](docs/ACCESSIBILITY-LOCALIZATION-ACCEPTANCE.md). Local acceptance does not imply universal Excel or production acceptance.

## AW-102 performance remediation

**8,538 tests across 316 files pass**, zero failures or skips, including 25 new performance and integrity checks. Twenty-five compiled SDK checks pass; all 48 validation hosts build. The final five-sheet, 100,000-cell/10,000-formula benchmark reduces source-edit p95 from **166.3 ms to 36.1 ms** on the local machine. Undo/redo p95 is **21.7 ms** and warm mount p95 **228.8 ms**.

**AW-102 remains Partial.** Final scroll p95 is **50.4 ms**, maximum **51.9 ms**; two responses exceed 50 ms. A separate 120-edit run records edit p95 **45.5 ms**, scroll p95 **50.6 ms** and maximum **65.2 ms**. No outliers were removed. The longest steady interaction task across these runs was **70 ms**; mount tasks reached **232 ms** and are separate from steady interaction. The original physical four-core/16-GB current/previous Chrome, Edge and Safari matrix, durable host saves, larger analysis tools and long-duration heap/GC/worker-memory traces remain open.

The supported 100,000-cell CSV import is **506.2 ms p95 / 2,600.6 ms maximum**, the 99,999-value aggregate **199.7 ms p95 / 200 ms maximum**, and active cancellation **0.2 ms p95 / 0.3 ms maximum**. The aggregate is exactly 10000499996. The two-million-cell population is explicitly refused by the unchanged 100,000-cell release capacity. These are local paint-opportunity and worker measurements, not a production SLA or proof of reference-machine performance.

Lazy population and formula-candidate inventories, deferred exact calculation signatures and diagnostic rows, and a cancellation path retaining historical evidence and spill protection remove repeated work. Mutable inputs, permission checks, late-response rejection, source values, selection totals, navigation and AutoFit remain covered. The benchmark now captures explicit bounded heap observations and downloads its report. Session memory is not durable storage.

PERF-R-001–012 preserve all 3,099 earlier manual templates: **3,111 total** (2,977 Not run, 51 Blocked unimplemented, 83 Blocked external, zero Passed). Only AW-102 delivery/gap/evidence fields change in the coverage sheet; status stays Partial. Base coverage remains **70 accepted / 5 Partial / 1 Planned — 6 open of 76**. Full coverage remains **78 accepted / 19 Partial / 11 Planned — 30 open of 108**. All original descriptions and 216 criteria, prior closure records and historical proofs remain intact.

[Measured evidence and remaining gates](docs/PERFORMANCE-REMEDIATION-VERIFICATION.json) · [Reproduce the benchmark](docs/INTERACTIVE-PERFORMANCE-ACCEPTANCE.md).

## Try the application

- Open **Ledgerglass.html** in a modern browser. All application code and styling are inside this file; the default assistant is local and makes no network requests. Coordinated browser saves require Web Locks; use localhost or HTTPS for predictable storage support. If a local file cannot save, export native JSON to keep a portable backup.
- Open **Embedded-Workbook.html** for the same workbook inside an example audit application. Its document list, window controls and persistence belong to the host; the Excel interface remains the reusable component. When serving the site, use **?embedded**.
- Or, with Node.js 22+, run `npm start` from this directory and open **http://127.0.0.1:4173**. This serves the prebuilt application without installing packages.
- A fresh session opens a blank **Book1 / Sheet1** with **A1** selected and the task pane closed. The standalone page restores its own locally saved workbook when available. Add `?temporary` to bypass browser persistence.
- Choose **File → Open** and select `examples/Advanced-features-demo.json` for a synthetic table/PivotTable/chart/shape/formula/page-layout example. Its control total is **54,400**.
- Open `examples/Review-workflow-demo.json` to try a collapsed outline, a two-message review discussion and an internal total link. The synthetic schedule totals **52,000**; Data → Group & Outline expands its detail rows.
- Open `examples/Audit-presentation-demo.json` to try reviewed cell moves, accounting formats and worksheet defaults. Its synthetic balances total **1,000**; B2 has an evidence note, E2 displays **36:00**, and the worksheet uses 28-pixel rows and 140-pixel columns. See the [integrated walkthrough](docs/CUT-SIZING-FORMATTING.md).
- Open `examples/Themes-evidence-demo.json` for bound styles with an explicit override, an independently hashed synthetic picture, a **400** total and **four** identified print pages. The original picture bytes are in `examples/Synthetic-evidence.png`; see [the integrated validation story](docs/THEMES-EVIDENCE-PRINT.md).
- Open `examples/Schedule-layout-demo.json` for cell styles, safe merging, conditional rules and two printed pages with repeated titles. Its synthetic control total is **400**; [the integrated validation story](docs/PRESENTATION-REMEDIATION.md) lists expected results.
- Choose **File → Open audit example**, or use `?demo`, for deterministic **synthetic demonstration data**: 40 journal lines, a connected trial balance, and an audit plan. Debit and credit totals are both **1,046,304.08**, with one deliberate repeated journal-line ID. The demo URL does not persist changes.

Use `examples/Closure-after.json` and `examples/Closure-before.json` for Review → Compare Workbooks and Risk Screening. The [end-to-end validation journey](docs/REQUIREMENT-CLOSURE-REVIEW.md) gives exact mappings, source controls, expected differences and signed/absolute results.

To develop: `npm ci`, then `npm run dev`. To build: `npm run build:portable` and `npm run build:lib`. The lockfile pins the resolved dependency versions. For the separate two-tab browser fixture, run `node scripts/build-concurrency-host.mjs` after both builds, then serve `/concurrency.html?run=review-01&document=shared`. The [concurrent-save procedure](docs/CONCURRENT-WORKBOOKS.md) explains lock controls, isolated run storage and conflict resolution.

## What works

- **Governed collaboration:** `CollaborativeWorkbook` adds host-authenticated editing, explicit synchronization, retained conflict review, offline recovery/idempotency, presence and shared discussions. `/collaboration.html` is the two-user synthetic validation host. [Acceptance](docs/COLLABORATION-ACCEPTANCE.md), [integration](docs/COLLABORATION-INTEGRATION.md), [verification](docs/COLLABORATION-VERIFICATION.json).

- **Custom functions and sanctioned add-in adapters:** Formulas → Add-ins & Functions. Versioned schemas, host permissions, input processing approval, result review, typed caches and linked entities; disable preserves raw formulas and labels historical values. `/workbook-extensions.html` is the synthetic validation host. [Acceptance](docs/WORKBOOK-EXTENSIONS-ACCEPTANCE.md), [integration](docs/WORKBOOK-EXTENSIONS-INTEGRATION.md), [verification](docs/WORKBOOK-EXTENSIONS-VERIFICATION.json).

- **External workbook links:** Data → Workbook Links. Inspect exact sources/ranges, approve snapshots and worksheet scopes, preview refresh or source replacement, and break complete formulas into typed values. Stale/missing source status and cached revisions remain visible. `/workbook-links.html` supplies synthetic validation fixtures. [Acceptance](docs/WORKBOOK-LINKS-ACCEPTANCE.md), [integration](docs/WORKBOOK-LINKS-INTEGRATION.md), [verification](docs/WORKBOOK-LINKS-VERIFICATION.json).

- **Structured and document sources:** File → Import sources. Native JSON/XML tables preserve types and field paths. Optional PDF/image/HTML extraction uses an explicitly approved host adapter or matching extraction JSON, with confidence review and retained original/cell evidence. `/source-imports.html` supplies synthetic validation fixtures. [Acceptance](docs/SOURCE-IMPORT-ACCEPTANCE.md), [integration](docs/SOURCE-IMPORT-INTEGRATION.md), [verification](docs/SOURCE-IMPORT-VERIFICATION.json).

- **Workbook formats:** File → Workbook formats. Native XLTX template import/derived export and explicit external conversion routes for XLS/XLSB/XLSM/XLTM/ODS, with isolated identity checks and exact original retention. `/workbook-formats.html` supplies synthetic validation fixtures. [Acceptance](docs/ALTERNATE-FORMATS-ACCEPTANCE.md), [integration](docs/ALTERNATE-FORMATS-INTEGRATION.md), [verification](docs/ALTERNATE-FORMATS-VERIFICATION.json).

- **Reusable report outputs:** Data → Report Outputs. Publish fixed or linked schedules and drawings; export PNG/SVG/HTML/JSON or embed a snapshot in a host report. `/report-outputs.html` provides the local acceptance fixture. [Integration](docs/REPORT-OUTPUT-INTEGRATION.md) and [execution evidence](docs/REPORT-OUTPUT-VERIFICATION.json).

- **Solver and statistics:** Data → Solver / Data Analysis. Preview bounded linear optimization, summaries, regression, correlation/covariance, distributions and forecasts. `/analysis.html` contains the local acceptance fixture. [Integration](docs/ANALYSIS-INTEGRATION.md) and [execution evidence](docs/ANALYSIS-VERIFICATION.json).

- **DAX calculations:** Data → Data Model → Measures / Calculated columns supports the documented subset. Open `/dax.html` for a synthetic sales/calendar example. [Integration](docs/DAX-INTEGRATION.md) and [execution record](docs/DAX-VERIFICATION.json) explain supported functions and stale-result safeguards.

- **Relational data model:** Data → Data Model creates typed sources, relationships, filters and aggregate measures. Open `/data-model.html` for the synthetic shared-dimension example, then preview and publish a result. Native JSON retains the model; XLSX exports warn about model limitations. [Integration](docs/DATA-MODEL-INTEGRATION.md) and [execution record](docs/DATA-MODEL-VERIFICATION.json).

- **Concrete source connectors:** run `npm ci` then `npm run demo:sources` on Node 22.16+ and open `http://127.0.0.1:4182/external-connectors.html` for actual local HTTP, SQLite and CSV folder reads. [Connector integration](docs/SOURCE-CONNECTOR-INTEGRATION.md), [acceptance story](docs/EXTERNAL-CONNECTOR-ACCEPTANCE.md), [current execution record](docs/EXTERNAL-CONNECTOR-VERIFICATION.json). Native ODBC and production vendor services remain untested.

- **External source connections:** Data → Connections & Refresh previews complete typed source data, rows and control-total changes, then rechecks host access before an undoable refresh. Failed requests retain a labeled stale snapshot. Try `/external-sources.html`. [Acceptance story](docs/EXTERNAL-SOURCES-ACCEPTANCE.md), [host integration](docs/EXTERNAL-SOURCES-INTEGRATION.md), [execution record](docs/EXTERNAL-SOURCES-VERIFICATION.json).

- **Reliability and recovery:** use **Workbook health** for classified outcomes and timings. Save failures remain visible; pending writes block duplicate retry, and readback recovers a lost acknowledgement. Run `/reliability.html` for the isolated fault demo. [Acceptance story](docs/RELIABILITY-RECOVERY-ACCEPTANCE.md), [execution record](docs/RELIABILITY-VERIFICATION.json) and [host runbook](docs/RELIABILITY-RUNBOOK.md).

- **Interactive performance:** immutable internal snapshots, session saves without repeated population serialization, batched cancellable calculations and safe changed-cell worker transport. Run `/performance.html` to reproduce 30-sample checks. [Measurements](docs/PERFORMANCE-VERIFICATION.json) and [acceptance story](docs/INTERACTIVE-PERFORMANCE-ACCEPTANCE.md).

- **Accessibility and localization:** focusable grid and cell announcements, modal focus isolation/return, skip links, responsive controls, reduced-motion/forced-color styling and a host message/number-locale API. Try `/accessibility.html`. [Story and open gates](docs/ACCESSIBILITY-LOCALIZATION-ACCEPTANCE.md).

- **Secure file processing:** workbook, delimited, configuration, picture and automation uploads use disposable parser workers, bounded input/results, cancellation, deadlines and safe error messages. Failed attempts preserve the open workbook; CSV/XLSX formula-injection and unsafe-link safeguards remain. Try `/file-security.html`. [Acceptance and integration](docs/FILE-SECURITY-ACCEPTANCE.md).

- **XLSX source preservation:** complete source-part inventory, exact original recovery from File → Original XLSX and compatibility, and byte-exact export of eligible passive custom metadata. Reviewed derived copies report dependent/unsafe part losses. Native JSON retains the original. Try `/xlsx-preservation.html`. [Acceptance story](docs/XLSX-PRESERVATION-ACCEPTANCE.md).

- **Consolidation and subtotals:** Data → Consolidate / Subtotals maps categories across sheets or approved workbooks, exposes source periods/currencies/units and contribution controls, publishes grouped linked schedules and marks stale/missing sources visibly. Try `/consolidation.html`. [Complete acceptance story](docs/CONSOLIDATION-ACCEPTANCE.md).

- **Repeatable transformation recipes:** Data → Transformation Recipes provides ordered steps for all required AW-051 families, saved immutable versions, source/definition/output fingerprints, per-step schemas and controls, complete output review and failure-safe separate publication. Try `/transformation-recipes.html`. Native JSON retains definitions and historical run evidence. [Acceptance story](docs/TRANSFORMATION-RECIPES-ACCEPTANCE.md).


- **Append and join populations:** Data → Append / Join supports append, inner/left/right/full/left-anti/right-anti modes, explicit text-key normalization, bounded fuzzy matching, complete match/exception review and cardinality reconciliation. Output is a guarded separate values snapshot; source data remains unchanged. Try `/advanced-combine.html`. [Contract and acceptance story](docs/POPULATION-COMBINE-ACCEPTANCE.md).
- **Saved filter views:** save, replace, rename and delete named criteria; native reopening retains historical source context and requires fresh preview before use. Try `/saved-filter-views.html`. [Policy and acceptance cases](docs/SAVED-FILTER-VIEWS.md).

- **Names and reusable formulas:** scoped names/constants/ranges, LET, named and inline LAMBDA, guarded rename, autocomplete and a non-mutating Name Manager test preview. Try `examples/Reusable-formulas.json` or `/reusable-formulas.html`. [Story and boundaries](docs/REUSABLE-FORMULAS-ACCEPTANCE.md).

- **Dynamic array schedules:** full spill recalculation, horizontal/multikey sorting, nested arrays, typed text extraction, spill diagnostics and derived-cell guidance. Legacy CSE imports are identified and stopped for review. Try `examples/Dynamic-arrays.json` or `/dynamic-arrays.html`. [Story and boundaries](docs/DYNAMIC-ARRAY-ACCEPTANCE.md).

- **Data → Clean up text** previews ordered text/numeric/extraction recipes, preserves leading-zero identifiers, and records original evidence and row identities in native cleanup history. Errors, hidden/unauthorized sources and stale previews block output. Try the before/after examples or `/text-cleanup.html`. [Story and boundaries](docs/TEXT-CLEANUP-ACCEPTANCE.md).

- Page Layout now checks numeric clipping, empty schedules, identifiers, pictures and drawing boundaries before export. **Create PDF** offers a direct local PDF download from the embedded editor. PDFs preserve rendered pages at 192 dpi; text is not selectable/searchable/tagged. [Print and export story](docs/PRINT-EXPORT-SCHEDULES.md).

- F1 or the quick-action book icon opens a searchable reference to current ribbon/File commands and shortcuts. Right-click cells or row/column headers, or use Shift+F10, for editing, sizing and review commands. Disabled commands explain their conditions. Escape from the grid returns to the ribbon so Shift+Tab through File reaches host controls. [Commands and shortcuts](docs/COMMANDS-AND-SHORTCUTS.md).

- View → Split compares worksheet areas in two or four panes with independent scroll groups, shared editing/formulas/Undo, F6 focus and adjustable dividers. Native workbook data preserves the layout; XLSX discloses its omission. Existing AutoFit, resizing, hidden views, freeze and zoom are verified together. [Worksheet view story](docs/WORKSHEET-VIEWS.md).

- Home or Insert → Insert Row / Insert Column now opens reviewed Insert, Delete and Move before operations. Preview rewritten formulas, new reference errors and original source-row associations; Apply once and Undo/Redo the complete change. CSV/TSV/XLSX imports retain stable original row identities in native JSON. [Row and column operations](docs/ROW-COLUMN-STRUCTURE.md).

- Page Layout → Themes manages workbook palettes/fonts and reusable named styles with inherited values and explicit cell overrides. Preview create/update/rename/delete/apply/reapply/detach changes before Apply. Native JSON keeps live bindings; XLSX freezes current supported formatting. [Themes and styles](docs/THEMES-NAMED-STYLES.md).
- Picture Properties retains and displays source references, exact-byte SHA-256, observed filename and plain-text locator. Image uploads and object edits guard current selection/workbook/access. Native JSON retains provenance; XLSX reports the loss and reopened images record observed media identity. [Evidence pictures](docs/EVIDENCE-PICTURES.md).
- Page Setup supports both page orders and live workbook/revision/engagement/period identifiers. Native retains tokens; XLSX freezes custom identity text with a warning. [Print ordering and identity](docs/PRINT-ORDER-IDENTITY.md).

- Excel-style ribbon tabs that adapt to the component width and scroll without overlapping Undo/Redo/AI controls, with contextual command groups, a name box and formula bar, neutral row 1, rectangular and drag selection, keyboard navigation, bottom sheet tabs and selection aggregates. Audit tools and the assistant open only when requested.
- Multiple-area selection with Ctrl/Command-click or comma-separated ranges, overlap-safe counts, union formatting/clear and bounded Tab/Enter traversal. The host receives every selected area; rectangle-only tools and LLM context reject multi-area requests explicitly. See [multiple-area selection](docs/MULTI-AREA-SELECTION.md).
- Editable sheets, quoted multiline TSV, internal copy with relative-formula translation, and Paste Special for values/formulas/formats, transpose, skip blanks and arithmetic. Cut captures a pending source; Paste previews source clears, destination changes and supported reference/metadata repairs before one atomic move. Unsupported object or reference intersections are rejected. See [reviewed cut and range moves](docs/RANGE-MOVE.md).
- Fill Down/Right/Up/Left, bounded number and date series with explicit month-end policies, and a fill handle that previews pattern copying or a two-number series before Apply. Each applied operation has one Undo step. Complete pattern inference and drag-edge automatic scrolling remain open. See [grid editing](docs/GRID-EDITING.md).
- Formula suggestions appear after `=` in both the formula bar and in-cell editor. Type a prefix, use the arrow keys, Tab or a mouse choice; function signatures and argument hints stay beside the draft. 395 local functions, modern lookups, bounded dynamic arrays, named ranges, structured table references, cross-sheet and absolute/relative references are supported. Refer to [formula compatibility](docs/FORMULA-COMPATIBILITY.md) for exact coverage and errors.
- Table creation with styles/totals and structured references; PivotTable summaries with row/column grouping, sum/count/average/min/max and grand totals; refreshable range queries with selected fields, filters, sorting and derived columns. Outputs preserve source data.
- Local PNG/JPEG pictures, movable/resizable shapes and eleven linked chart views, in-cell sparklines and complete bounded source inspection; page margins, portrait/landscape, Letter/A4/Legal paper, print areas, scaling, headers/footers and paginated print preview. Repeated title rows/columns and full-axis manual page breaks have dedicated controls and supported native/XLSX exchange. [Print title and break details](docs/PRINT-TITLES-BREAKS.md).
- Conditional formatting now has draft/edit/delete/reorder/preview/Apply management, anchored formula rules, priority/Stop if true, blanks/errors/duplicate/unique conditions, color scales, bars/icons and visible rule errors in grid/print. Supported XLSX conversions preserve meaning or reject unsupported cases; full Excel conditional parity remains open. [Rule details](docs/CONDITIONAL-FORMATTING.md).
- Data validation, reviewed duplicate-derived populations, format painting, merge/unmerge, local worksheet protection and browser-supported spelling corrections. Insert/delete rows and columns shift supported references and metadata with explicit guards against unsupported destructive edits.
- Format Cells provides number, currency, accounting, percentage, date/time, signed elapsed-time and bounded custom formats in six configured locales. Currency, precision, display scale and negative presentation have immediate previews; numeric raw values and formulas stay unchanged. Accounting separates symbol and amount in grid and print. Unsupported new patterns are rejected and legacy fallback patterns expose warnings. See [number formats](docs/NUMBER-FORMATS.md).
- Home → Cell Styles provides font family/size, bold/italic/underline/strikethrough, colors, horizontal/vertical alignment, indentation and per-edge borders, including union-aware presets. Border changes require Unmerge first. Wrapped text, notes and review tickmarks remain supported. [Cell style details](docs/CELL-STYLES.md). Wrapped and enlarged cells update visible row measurements; explicit AutoFit stores a content-based size. Full Excel font/layout and arbitrary custom-format fidelity remain open.
- Drag row/column boundaries to resize, or double-click to AutoFit. Shift-click headers to size multiple rows/columns; Home also exposes both AutoFit commands. View → Default Sizes sets worksheet-wide defaults and previews any explicit override reset. Sizing respects formatted values, fonts and wrapping, supports undo, and survives supported JSON/XLSX exchange. Escape cancels an unfinished drag. Heights support 21–546 CSS pixels and column widths 40–600 pixels. See [worksheet dimensions](docs/SHEET-DIMENSIONS.md).
- Freeze multiple rows and columns together, hide/unhide rows or columns using compact intervals, and retain their native/XLSX metadata. Print layout omits manually hidden rows/columns and hidden drawing anchors. Gridlines, headings and zoom remain independent controls. Freeze panes supports up to 100 visible rows and 50 visible columns; split-pane behavior is documented in the worksheet view story.
- Nested row/column groups, collapse/expand/ungroup and level controls preserve manual hiding, calculations and totals. Native JSON retains editable groups; XLSX retains collapsed visibility. See [outlines](docs/OUTLINES.md).
- Local review discussions support replies, resolution/reopening, author/time labels, and retained deleted targets. Safe web/email and internal workbook links have interactive cell controls. Native JSON preserves threads and links; XLSX does not. See [comments and links](docs/REVIEW-COMMENTS.md).
- Formula drafts commit to their original cells when selection changes; Escape/Cancel discards edits and restores grid focus. Unchanged edits do not create history. Shift-selection keeps the original active cell; Ctrl/⌘A selects the full sheet without allocating its cells.
- Source-preserving sort/filter views, keyboard navigation through visible cells, Ctrl/Command+Arrow data-edge jumps, Page Up/Down, Home and row/column/sheet shortcuts. Find skips manually hidden cells and can clear a filter to reveal a match. Selection totals include hidden and filtered cells. Advanced Find adds selection/sheet/workbook scope, case/whole-cell/wildcard matching, formula/calculated/formatted search, paginated results and exact replacement previews. Go To Special selects formulas, errors, constants, notes or blanks. Reviewed physical sorts are available through Custom Sort; complete Excel search/selection parity remains open. See [search](docs/SEARCH.md).
- Duplicate-ID checks, absolute row-amount screens, deterministic row sampling, account summaries and dated aging buckets, each with source-cell references. Export an evidence report containing the revision, parameters, findings and citations.
- XLSX import/export of the documented subset, CSV/TSV import, CSV export and native JSON import/export. Native JSON retains application metadata and activity; XLSX is a lossy interchange format.
- File controls for workbook properties, required template engagement/period, independent template copies with exact source-revision provenance, and worksheet add/rename/copy/move/delete/visibility/tab colors. Supported references follow structural edits; dependency previews and local structure locks prevent unsupported destructive changes.
- Named checkpoints with cell/structure/property differences, recorded actor/time where available, and restoration as a new revision retaining intervening history. Complete current/previous snapshots preserve selection and recover from a corrupt current copy. Storage errors and detected stale saves retain unsaved work for export.
- Session-memory storage by default, opt-in coordinated browser storage or a host-owned version store. Same-origin browser saves use an exclusive per-document lock and retain conflicting local proposals for review/export. Activity is editable with the native file; it is not a secure or immutable audit trail.
- Provider-neutral assistant context, proposal validation, preview/discard/apply, cancellation, stale-response rejection and a deterministic local demo. Omitted note/style fields survive accepted assistant edits.
- React component, typed workbook API and compiled ES module library.
- Official-SDK MCP server with nine tools for selected context, range reading, formula inspection, calculation replay, audit checks, capabilities, function contracts, scoped compatibility reports and queued proposals. The server cannot apply cell edits. A host-delivered proposal is validated and shown for human approval.

Collapse the ribbon using the chevron beside Undo, double-clicking any tab, or Ctrl+F1 while focused in the workbook. Tabs temporarily reveal commands; Expand keeps them open. The setting is local to the mounted component.

## Quick walkthrough

1. Choose **File → Open audit example**. On **General ledger**, select E2 and change the value using the formula bar. G2 and the Trial balance recalculate. Undo restores it.
2. Open **Review → Audit Checks** (also accessible from Data), then select **Duplicate entries → Run check**. The finding points to A16 and A18. Repeated IDs are review candidates, not automatic evidence of an error.
3. Select **Select a sample**, enter a seed and sample size, and run the check. The same revision, row order, range and seed reproduce the sample.
4. Enter `E2:E5` in the selection box and select **AI Assistant** at the right of the ribbon. Choose **Highlight values above 25000**, inspect the proposal, then Apply or Discard.
5. Choose **File → Workbook properties** to set the workpaper name, engagement and period. **Manage worksheets** handles sheet structure, visibility, color and local protection.
6. Choose **File → Version history**, save a checkpoint, make a change, and compare or restore the checkpoint. A restore creates a new revision and keeps intervening activity.
7. Choose **File → Save workbook data** for the native JSON workbook, or **File → Save as Excel** for the supported XLSX subset. Named checkpoints live in the separate version store; retain that envelope as well if checkpoint history must travel with the file. **New from current template** downloads an independent native workbook and keeps the source unchanged.

Every formerly placeholder ribbon command now opens a working tool. Commands can still be disabled by read-only access, worksheet protection, unavailable history or a pending operation. Each tool implements the documented subset; the ribbon does not imply full Excel parity. [The interface specification](docs/EXCEL-INTERFACE.md) records the requested layout and behavior independently from the wider Excel parity backlog.

Audit checks treat the selected range's first row as field names and disclose amount-column inference and omitted rows. Labelled total rows are excluded. They include rows hidden by the search view. For aging, provide a suitable open-item population and an explicit as-of date. Sampling does not establish a statistical sample size or a population conclusion.

## Embed in a React application

Build the library, then install this folder or its npm package archive in your host project. The package is named `@ledgerglass/workbook`; it is not published to a registry.

```tsx
import { useMemo, useState } from "react";
import {
  AuditWorkbook,
  createBrowserWorkbookStore,
  createHttpAdapter,
} from "@ledgerglass/workbook";
import "@ledgerglass/workbook/style.css";
import type { PersistenceState, Workbook } from "@ledgerglass/workbook";

const adapter = createHttpAdapter({ endpoint: "/api/workbook-assistant" });

export function EngagementWorkpaper({ documentId, workbook }: {
  documentId: string;
  workbook: Workbook;
}) {
  const [localWorkbook, setLocalWorkbook] = useState(workbook);
  const [saveState, setSaveState] = useState<PersistenceState | null>(null);
  const store = useMemo(
    () => createBrowserWorkbookStore(localStorage, `workpaper:${documentId}`),
    [documentId],
  );
  const saved = saveState?.status === "saved" &&
    saveState.workbookId === localWorkbook.id &&
    saveState.revision === localWorkbook.revision;
  return (
    <>
      <p>{saved ? "Saved in this browser" : saveState?.message ?? "Not saved yet"}</p>
      <AuditWorkbook
        key={documentId}
        initialWorkbook={workbook}
        actor="Joey"
        adapter={adapter}
        workbookStore={store}
        workbookStorePersistent
        onChange={setLocalWorkbook}
        onWorkbookLoaded={setLocalWorkbook}
        onPersistenceStateChange={setSaveState}
      />
    </>
  );
}
```

`onChange` reports local workbook changes immediately; it is not a durable-save acknowledgment. `onWorkbookLoaded` reports the saved workbook loaded when a session opens, so host metadata reflects its actual revision. `onPersistenceStateChange` reports `saving`, `saved`, `error` or `conflict`, together with workbook ID, revision and any message. Base the host's saved indicator on a confirmed event for the current workbook. Omit `adapter` for the local demo. Omit both `persistenceKey` and `workbookStore` for an in-memory session. **No provider key belongs in browser code.** See [the LLM protocol](docs/LLM-PROTOCOL.md).

`initialWorkbook` seeds a mounted instance; this is an uncontrolled editor. To explicitly replace a local session with another same-ID revision, use a different React `key` after retaining any unsaved work. Changing the host document, persistence namespace or store object starts a fresh session with isolated undo/AI state. Echoing the component's own loaded/imported identity through a callback preserves that session and its pending save. `readOnly` suppresses UI mutations; it is not server authorization. Keep adapter and store objects stable during a session.

Use a unique stable document namespace with `createBrowserWorkbookStore` for browser persistence, or the same namespace for editors intentionally sharing a saved document. The component's `persistenceKey` uses this guarded browser store as well. Saves, checkpoints and restores acquire the same per-namespace Web Lock and complete asynchronously; missing or denied locks refuse writes and retain local work. The synchronous `createWorkbookStore` supports isolated memory stores or equivalent host-owned coordination. It cannot coordinate tabs by itself. `workbookStorePersistent` controls the displayed storage description; it does not make an adapter durable.

The embedded example owns separate document envelopes, retains hidden editor sessions across workpaper switching/closing, and reports confirmed saves through the callback. Its legacy workspace migrates under document locks only when a destination is absent; the original workspace bytes remain preserved. The active workpaper preference is tab-local. Remounts restore confirmed saves/checkpoints; unsaved proposals require the retained live session or an export. See [workbook lifecycle and persistence](docs/WORKBOOK-LIFECYCLE.md) and [concurrent workbooks](docs/CONCURRENT-WORKBOOKS.md).

For a simple iframe embed, host `dist/` and use `<iframe src="/workbook/" title="Audit workbook" width="100%" height="900"></iframe>`. There is no cross-origin messaging protocol in this release; use the React component for data integration. The workbook and ribbon use scoped component styles; the standalone root removes the browser's default body margin. A host can render `<AuditWorkbook />` for a blank editor or pass `initialWorkbook` for its own data. `createBlankWorkbook()` is also exported for host use.

## Supported limits

| Resource                                         |                                                  Current application cap |
| ------------------------------------------------ | -----------------------------------------------------------------------: |
| Worksheets                                       | No fixed sheet-count cap; available memory and workbook resource budgets |
| Rows / columns per worksheet                     |                                                 1,048,576 / 16,384 (XFD) |
| Stored cells per workbook (including formatting/notes)                     |                                                                  100,000 |
| Cell changes per patch batch / assistant context |                                                                500 / 500 |
| XLSX, CSV, TSV source file                       |                                                                    10 MiB |
| XLSX expanded content                            |                                                                    50 MiB |
| Native JSON save and reopen                      |                                                              50 MiB UTF-8 |
| Activity events / retained UI undo snapshots     |                                                               2,000 / 50 |
| Named local checkpoints per store                |                                            20 default; configurable 1–20 |
| Complete local version envelope                  |    10 MiB default; configurable 1 KiB–50 MiB; browser quota may be lower |
| Cell text / formula text | 32,767 / 4,096 UTF-16 units (formula includes =) |
| Formula rectangle / reads / dependent levels / one spill | 20,000 / 100,000 / 64 / 10,000 |
| CSV output | 10 MiB UTF-8; 100,000 rectangle cells including gaps |
| XLSX entries / explicit source cells | 2,000 / 100,000 (blank source cells count) |
| XLSX expansion ratio | 200:1 per entry larger than 1 MiB expanded |

Dedicated structural/table/query/pivot/restore commands can produce larger bounded operations within the 100,000-populated-cell budget, recording cell-change events in batches of at most 500. They do not expand the assistant’s 500-cell authority.

These are enforced safety limits, not performance promises. Work happens in the main browser thread in this release. The address space matches Excel; populated-data capacity, calculation limits and complete feature parity remain separate. Both row and column rendering are virtualized, and empty grid cells are not allocated. References to hidden sheets can affect selected formula results; only selected raw values/results are sent to the assistant. Selection is a data-minimization mechanism, not an authorization boundary.

XLSX uses ExcelJS. Unsupported formatting/layout/features can be lost. Rich text is flattened. Basic tables, safe merges, supported validation/conditional rules, PNG/JPEG pictures and page setup are exchanged; supported global/local range, constant and computed names with comments are exchanged. Native pivot/query refresh definitions, charts/shapes, tickmarks/history, passwords, engagement/period/template provenance, unsupported name semantics/attributes, native watch lists and other advanced Excel metadata require native JSON. Supported created/modified dates, hidden/very-hidden states and tab colors have XLSX mappings. Checkpoint envelopes require separate host-store retention. Macro-enabled/legacy formats, executable objects and suspicious ZIP archives are rejected. External-action formulas are exported as literal text. Always keep the original source workbook and read the import notices.

## MCP and external AI

Import the Node-only server factory from `@ledgerglass/workbook/mcp`, separately from the browser component. The included stdio reference reads an explicitly configured native JSON workbook and queues validated proposals to a separate outbox. `docs/MCP.md` contains the CLI configuration and host bridge contract. The browser accepts `externalProposal` plus its original selected context and stages it for review; it never applies external changes automatically.

The file bridge does not synchronize an open browser automatically. A deployed host supplies that connection, authenticated access and durable storage. Local SDK protocol tests and a real stdio client handshake were exercised; no external LLM provider was connected.

## Files and validation

- [Requirements and architecture](docs/REQUIREMENTS.md)
- [Excel-style interface specification](docs/EXCEL-INTERFACE.md)
- [Traceable feature matrix](docs/FEATURE-MATRIX.csv)
- [Workbook lifecycle, templates and local recovery](docs/WORKBOOK-LIFECYCLE.md)
- [Gap closure progress in requirement order](docs/GAP-CLOSURE-PROGRESS.md)
- [Reviewed moves, sizing and formatting milestone](docs/CUT-SIZING-FORMATTING.md)
- [Number formats, locales and custom-pattern boundaries](docs/NUMBER-FORMATS.md)
- [Formula compatibility](docs/FORMULA-COMPATIBILITY.md)
- [Advanced feature user stories and acceptance criteria](docs/ADVANCED-FEATURE-USER-STORIES.md)
- [Tables, PivotTables and range queries](docs/TABLES-PIVOTS.md)
- [Drawings and page layout](docs/DRAWINGS-PAGE-LAYOUT.md)
- [LLM adapter protocol and host responsibilities](docs/LLM-PROTOCOL.md)
- [MCP setup and host integration](docs/MCP.md)
- [Validation checklist](docs/VALIDATION-CHECKLIST.md) and [complete test-case register](docs/VALIDATION-TEST-CASES.csv)
- [Verification results and remaining release gates](docs/VERIFICATION.md)
- `src/`: complete TypeScript/React source; `lib/`: compiled component plus declarations; `dist/`: prebuilt standalone website; `Ledgerglass.html`: portable standalone workbook; `Embedded-Workbook.html`: portable host application example.
- `examples/`: embedding source, audit workbooks and a 25-sheet boundary workbook containing XFD1048576. Regenerate these after building the library with `npm run generate:examples`.

The validation register is an execution plan, not a claim that every target feature passed. It includes tests for unimplemented features and external AI environments, with explicit starting statuses.

Run `npm test` for the calculation, file, proposal, MCP and React-state suites. Browser tests are provided in `tests/browser`; run `npx playwright install chromium` and `npm run test:e2e`, or specify an installed browser with `CHROME_PATH`. The development environment used direct in-app browser checks because launching a separate automated browser was blocked; see verification for what was actually exercised.

Dependencies are under their respective licenses; see `THIRD-PARTY-NOTICES.txt`. The code uses open-source dependencies, with no mandatory cloud account or commercial spreadsheet license. This does not imply that full Excel-compatible commercial deployment has been legally or operationally certified.

The presentation remediation adds reviewed merge previews with retained covered values and annotations. Native JSON preserves them; XLSX refuses an export that would discard covered populated content. See [the presentation story](docs/PRESENTATION-REMEDIATION.md) and [verification record](docs/VERIFICATION.md) for executed scope and remaining acceptance gates.

`exportXlsxWithWarnings(workbook)` returns `{ blob, warnings }` for hosts that display the conversion inventory. The existing `exportXlsx(workbook)` retains its Blob-only signature. Export warnings describe preserved and flattened metadata; native JSON remains the complete editable application format.

## Calculation review tools

Formulas → **Evaluate Formula** records actual local evaluation steps, errors, lexical bindings/branches and spill reads. Replay stays on one captured revision without re-running volatile inputs. Export JSON for review or use the public `evaluateFormulaSteps` / `FormulaEvaluationTools` interfaces. See [AW-041 acceptance and limits](docs/CALCULATION-INSPECTION-ACCEPTANCE.md) and open `dist/calculation-inspection.html` for the isolated 625/99/21 control example.

Formulas → Trace Precedents/Trace Dependents shows bounded static relationships and explicit missing/hidden/unsupported diagnostics without changing the workbook. Name Manager previews scoped renames before Apply. SUBTOTAL, ISFORMULA and FORMULATEXT are available in completion and Insert Function. Quick text filtering remains view-only; SDK filter exclusions are explicit calculator context.

Open `examples/Calculation-review-demo.json` for 100/80/120/7 control totals, a hidden source row, a nested subtotal and workbook/local names. Follow [the integrated story](docs/CALCULATION-REVIEW.md) and the appended validation cases.

## Watched calculation review

Formulas → Watch Window saves up to 100 watched cells with reviewed list changes, current raw/results and explicit unavailable/orphan states. Watches follow supported structural moves and persist in native JSON; XLSX conversion discloses their omission. AGGREGATE supports 19 selectors and options0–7 within documented array/resource bounds. Supported name constants, formulas, comments and local/global scopes now survive XLSX.

Open `examples/Watched-aggregate-demo.json` and follow [the end-to-end story](docs/AGGREGATE-WATCHES-NAME-EXCHANGE.md) for 35/55/10 and 7/100/110 controls. Full Excel/provider acceptance remains open.

## Data analysis additions

The twelve database functions are available through formula completion, Insert Function, the public SDK and approved assistant/MCP proposals. Insert → Table provides reviewed calculated-column fills/clears and explicit row appends. Native JSON retains reusable rules; XLSX retains supported cell formulas with a rule-loss notice. Data → Goal Seek previews a bounded single-variable candidate and residual before one Apply/Undo.

Open `examples/Data-analysis-demo.json` and follow [the integrated story](docs/DATA-ANALYSIS-REMEDIATION.md). Controls start at 400/200/2/200; the appended record changes them to 500/(500/3)/3/200. Goal Seek proposes input 20 for revenue 600. Per-record database formula criteria, automatic adjacent-typing table expansion, Scenario Manager and complete Excel/provider acceptance remain open.

## Reviewed population preparation

Data → Text to Columns previews quote-aware field splitting and explicit Text/General conversion into a separate empty range or new sheet. Remove Duplicates shows every member and creates a unique population plus excluded-row appendix, preserving original evidence. Append / Join reconciles exact typed matches, unmatched records and many-to-many expansion before creating a fixed output snapshot. All three use guarded previews and one host Apply/Undo.

Open `examples/Population-preparation-demo.json` and follow [the end-to-end story](docs/POPULATION-PREPARATION.md). Controls are five source rows totaling 850, four unique rows totaling 650, three vendor matches plus one unmatched record on each side, and five rows totaling 725 after appending the additional record to the unique population. Full recipes/fuzzy matching/Flash Fill and independent Excel/provider acceptance remain open. The old direct SDK `removeDuplicateRows` helper is deprecated; use `planDuplicateReview` / `applyDuplicateReviewPlan`.

## Release evidence

From the full source checkout, `npm run verify:release` builds both deliveries, runs the complete suite and checks its actual JSON report against the reviewed release scope. After implementation review, update the matrix and per-criterion evidence, record the actual test counts in `docs/RELEASE-SCOPE.json`, and run `npm run update:coverage`. `npm run check:claims` detects stale coverage without rewriting it; supply `-- --report <path>` to validate actual execution evidence. The npm SDK archive intentionally contains the runtime, types and documentation; the full source archive supplies the tests and release scripts. [Gate contract and limits](docs/RELEASE-CLAIM-GATE.md).

## Previous checkpoint: typed entry, aging and adapter acceptance

Five original requirements close in this pass: **AW-003, AW-011, AW-012, AW-076 and AW-086**. Coverage changes from **25 to 30 locally accepted**, **62 to 57 Partial**, and **21 to 21 Planned**; open requirements fall **83 to 78**. All 108 original descriptions and 216 acceptance criteria remain unchanged.

Added hidden-sheet calculation inspection, defined-name navigation and visible selection policy, deterministic typed ISO dates, configurable signed aging with exact controls/exceptions, and provider-neutral typed errors/deadlines/cancellation. Independent review fixed aging row-order dependence, rejected non-string error codes and incomplete caller contexts, and retained immediate adapter dispatch while bounding requests. The delivered-fixture SDK/UI journey verifies 11 source rows, 7 aged, 4 excluded; USD 2,225 = USD 2,050 + USD 175; EUR 1,000 is outside declared-currency totals.

Open `examples/Aging-review-demo.json` and follow [the supplied user story](docs/PARTIAL-COVERAGE-CLOSURE-STORY.md). Review → Aging Schedule exposes editable date/bucket settings, exception rows and a downloadable current report. The Name Box resolves static local/global range names. Both editors and external clipboard paste recognize exact ISO dates; ambiguous dates remain text. Manage Worksheets exposes calculation participation, and the footer states that selection aggregates include hidden/filtered cells. `requestLLMProposal` provides bounded typed adapter errors and cancellation; real-provider validation remains external.

Final automated regression: **3,671 passed across 145 files**. The manual register has **1,550 scenarios**, including 46 new Not run templates; original requirement targets remain unchanged.

## Formula families and local embedding acceptance

Five original requirements close in this pass: **AW-031, AW-032, AW-033, AW-036 and AW-096**. Coverage changes from **30 to 35 locally accepted**, **57 to 52 Partial**, and **21 to 21 Planned**; open requirements fall **78 to 73**. All 108 original descriptions and 216 acceptance criteria remain unchanged.

Arithmetic and rounding, logical/information, lookup/reference and descriptive-statistical families now meet their complete original local fixture contracts. LOOKUP, CHOOSE, MODE, PERCENTILE, QUARTILE and RANK bring the catalog to 363 functions, with 359 eligible for reviewed assistant proposals. Corrections include decimal ties, exact rounding multiples, signed cancellation, representable means, bounded products, typed wildcard lookups, blank-result preservation, invalid shape/mode refusal and operational-error propagation. Static local runtime has a visible offline notice, safe host failure/recovery and a verified separate React consumer; mixed-module dependency packaging was corrected.

Final regression: **4,191 tests across 151 files passed**. The validation register contains **1,607 scenarios**, including 57 new Not run templates. Follow the [integrated story](docs/FORMULA-LOCAL-RUNTIME-CLOSURE-STORY.md) and [independent host procedure](docs/LOCAL-RUNTIME-ACCEPTANCE.md).

Local acceptance uses declared semantics and independently specified fixtures. OFFSET/INDIRECT still return #NAME? under the original explicit-refusal criterion; strict conditional range shapes, scalar conditional tests, deterministic decimal policy and bounded calculations are documented. It is not universal Excel parity. Licensed Excel comparison, other browser/assistive-technology profiles, real external LLMs and production-host controls remain separate unexecuted gates.

## Concurrent workbook acceptance — AW-009

AW-009 now meets its original description and both criteria: separate editors retain independent cells, selection, storage and Undo, while shared-document browser saves use an exclusive lock around the full stale-guard check and write. A conflicting editor keeps its complete local proposal for comparison, export and explicitly confirmed reload. Embedded document switches and window close/reopen retain those live sessions.

Final regression: **4,251 tests across 155 files passed**. The manual register contains **1,623 scenarios**, including 16 new Not run templates; all 1,607 prior records are unchanged. The real two-tab public-SDK check observed one saved winner, the retained losing proposal, explicit reload, and independent-document Undo/Redo. Follow [the story and integration contract](docs/CONCURRENT-WORKBOOKS.md) and [observed browser evidence](docs/CONCURRENT-WORKBOOKS-BROWSER-EVIDENCE.md).

All writers to one browser namespace must use the coordinated store or its exact exclusive lock. Automatic merge, cross-device coauthoring, production-host authorization and crash-durable unsaved drafts remain separate work.

## Explicit capacity acceptance — AW-010

AW-010 now meets its unchanged original description and both criteria. Workbook help, the public `WORKBOOK_CAPACITY` export and MCP `workbook_get_capabilities` publish the same capacity and distinguish Excel address limits from local processing budgets. XLSX dimensions and source coordinates are checked before workbook parsing; CSV exports count exact UTF-8 bytes including escaping and safety prefixes. Invalid imports retain the current workbook, saved bytes, selection and Undo.

Final regression: **4,305 tests across 159 files passed**, including 54 new capacity checks. The manual register contains **1,647 scenarios**; 24 new SC templates append all 1,623 prior records unchanged. See [the capacity story and boundary corpus](docs/WORKBOOK-CAPACITY.md), [recorded measurements](docs/CAPACITY-BENCHMARK.json) and [browser evidence](docs/CAPACITY-BROWSER-EVIDENCE.md).

Earlier capacity-checkpoint edits took about 1.2–1.8 seconds. The AW-102 final local 30-edit run measures 105.8 ms p95 and 41.7 ms scroll p95; its 57.3 ms maximum scroll and unexecuted reference profile keep **AW-102 Partial**. The 2,000,000-cell population is explicitly outside release capacity. Same-address-space support does not imply full Excel compatibility, unrestricted storage, a performance SLA or validated external LLM integration.

## Worksheet structure acceptance — AW-015

AW-015 now meets its unchanged original description and both criteria within the published local boundaries: insertion, deletion and whole-row/column moves update supported dependencies and metadata, preview resulting reference errors, and preserve original file-row identities. The embedded dialog supports paged review, destructive-change acknowledgment, stale/read-only/protected refusal and one-step Undo/Redo. Native save/reopen retains the source registry; XLSX export warns of its omission.

At the AW-015 release: **4,355 tests across 162 files passed**, including 50 new structure/provenance/editor checks. That release contained **1,671 manual scenarios**; 24 ST templates append all 1,647 prior records unchanged. See [the end-to-end story](docs/ROW-COLUMN-STRUCTURE.md) and [recorded browser execution](docs/STRUCTURE-BROWSER-EVIDENCE.md). Build the isolated public-SDK fixture with `node scripts/build-structure-host.mjs` after the library and portable builds, then open `/structure.html`.

Three-dimensional/external references, split or destructive structures and other explicitly documented rewrites refuse before mutation. Broader cell-transformation lineage and tamper-proof evidence remain open under AW-073. Historical AW-010 performance measurements identify their original build; they are not measurements of this AW-015 build. No live external LLM or full Excel compatibility claim is added.

## Worksheet view acceptance — AW-016

AW-016 meets its unchanged original description and both criteria within the published local limits. View → Split supports top/bottom, left/right and four panes with shared editing, formulas, selection and Undo. Drag or keyboard-resize dividers, switch panes with F6, and remove or replace the split with frozen panes. Wrapped AutoFit, invalid-size rejection, hidden-row/column restoration and zoom retain the underlying values and source identities.

At the AW-016 release: **4,382 tests across 164 files passed**, including 27 new split-view, source-preservation, shared-recovery and print-layout checks. That release contained **1,695 manual scenarios**; 24 VW templates append all 1,671 prior records unchanged. See [the end-to-end acceptance story](docs/WORKSHEET-VIEWS.md) and [final-build browser evidence](docs/VIEW-BROWSER-EVIDENCE.md). Build the isolated public-SDK fixture with `node scripts/build-view-host.mjs` after the library and portable builds, then open `/views.html`.

Split layout is native-only. Scroll positions and 50–200% zoom are local to the editor session; row/column and divider limits are documented. The current local closure does not assert full Excel geometry/interchange parity, the complete browser/PDF matrix or external model acceptance. Historical AW-010/AW-015 browser and benchmark records retain their original artifact scope.

## Command acceptance — AW-020

At the AW-020 release: **4,404 tests across 165 files passed**. AW-020 closes against its unchanged command-discovery and host-focus criteria. That release listed **68 remaining gaps**. The [current coverage document](docs/COVERAGE-SUMMARY.md) contains the updated concise list and detailed gaps. AW-073 wording credits the source-row IDs already delivered with AW-015; complete lineage and evidence hash validation remain Partial.

The reference derives **129 command locations and 29 shortcut entries** from the actual ribbon/File definitions. Context menus provide 15 routes into existing operations and Undo. Host-reserved shortcuts, native text editing, unlisted modifier keys, independent workbook focus and stale/read-only actions are guarded. The reference overlay was visually checked above two embedded windows. Build the public-SDK fixture with `node scripts/build-command-host.mjs` after the library and portable builds, then open `/commands.html`.

That release contained **1,719 manual scenarios**; KC-001–024 append all 1,695 prior records unchanged. See [the complete user story](docs/COMMANDS-AND-SHORTCUTS.md) and [AW-020 browser evidence](docs/COMMAND-BROWSER-EVIDENCE.md). Accessibility/browser matrices and external LLM provider acceptance remain separate open gates.

## Print/PDF acceptance — AW-028

At the AW-038 release, regression was **4,803 tests across 177 files passed**, with no failures or skips. Coverage is **43 locally accepted, 44 Partial, 21 Planned — 65 open**. The manual register contains **1,803 scenarios**, including 24 new DY cases; all 1,779 preceding records are unchanged.

AW-028 closes against its unchanged page-layout and fixed-output criteria. [Browser-created sample PDFs and execution evidence](docs/PRINT-PDF-EVIDENCE.md) demonstrate repeated headings, FY2025/revision 7/page identifiers and unclipped monetary totals. The direct PDF is a bounded image snapshot; native printer/Excel/browser/font/accessibility matrices remain separate gates. [Current remaining gaps](docs/COVERAGE-SUMMARY.md).

AW-040 now includes 18 further regression/statistical functions and per-record database formula criteria, with 395 local functions and 391 assistant-safe names. See [specialist remediation](docs/SPECIALIST-FUNCTIONS-REMEDIATION.md). Its verified-Excel-version criterion remains open.

## Structured tables and PivotTables — AW-044 / AW-056

Insert / Table provides explicit header resolution, typed-column inspection and conversion previews with preserved original source cells. Existing totals, structured references, calculated-column rules and reviewed append/resize remain available. Insert / PivotTable adds distinct counts, row/column subtotals, date groups and aggregate calculated fields. The result strip shows source range/version/time and visibly stale output; refresh creates a new result while retaining the prior result. New output headings fit and wrap.

Open `dist/structured-pivot.html` for the isolated public-SDK host or load `examples/Structured-tables.json` and `examples/Pivot-tables.json`. The reconciliation control is **4 included + 1 excluded = 5 records**, with **sum 500 / average 125 / distinct count 3**. `Structured-source-evidence.json` retains source values and `Pivot-reconciliation.json` records saved lineage. [The story and all feature criteria](docs/STRUCTURED-TABLES-PIVOTS-ACCEPTANCE.md) document public interfaces and limits. Native JSON preserves types/evidence/pivot definitions; XLSX exports supported current values and tables without native pivot caches or these review records.

At the AW-044 / AW-056 release, regression was **5,159 tests across 193 files, no failures or skips**. Coverage: **47 locally accepted / 40 Partial / 21 Planned — 61 open**. The manual register has **1,927 scenarios**, preserving all previous 1,883 records byte-for-byte. Actual Excel/provider and production host acceptance remain separate gates.

## Pivot exploration, charts and sparklines — AW-057 / AW-058

Filter / explore PivotTable adds slicers, timelines, field sorting, three percentage modes, PivotCharts and exact value-cell drill-through. Double-click a pivot value to inspect source rows and their original locators, then export or extract them. Stale totals refuse drill. Chart properties offers every original chart-view family, axes/units/alt text, explicit row/series filters, full source inspection and SVG/JSON exports. Sparkline views fit one cell in the worksheet and print.

Open `dist/pivot-charts.html` (or `Pivot-charts-demo.html` in the local ZIP). Native examples: `Chart-views.json` and `Pivot-exploration.json`; their inspection, filter-context and drill reports accompany them. [All features, end-to-end story and acceptance criteria](docs/PIVOT-CHARTS-ACCEPTANCE.md) include precise limits and external AI/Excel gates.

At the AW-057 / AW-058 release, regression was **5,270 tests across 197 files, no failures or skips**. Coverage: **49 locally accepted / 38 Partial / 21 Planned — 59 open**. The manual register has **1,971 cases**, preserving all previous 1,927 records byte-for-byte. Native XLSX chart/pivot/slicer/sparkline XML and actual provider/production acceptance remain outside this release.

## Scenarios, sensitivity tables and Goal Seek — AW-060

Data → Scenario Manager compares named assumptions, retains baseline/source evidence, saves reusable definitions and summary reports, and applies a selected successful scenario only after review. Data Table adds one-variable row/column layouts with multiple results and two-variable matrices. Explicit refresh keeps earlier reports; saved baseline restoration, read-only inspection, native round-trip and Undo/Redo are verified. Goal Seek retains its bounded search, residual review and nonconvergence safeguards.

Open `dist/what-if.html` or `What-if-demo.html` in the local ZIP. Saved examples: `What-if-analyses.json`, `What-if-evidence.json` and `Goal-seek-review.json`. [The end-to-end story and acceptance criteria](docs/WHAT-IF-ACCEPTANCE.md) cover every local feature, precise budgets, SDK integration and external AI/Excel gates.

At the AW-060 release, regression was **5,356 tests across 199 files, zero failures or skips**. Coverage: **50 locally accepted / 37 Partial / 21 Planned — 58 open**. Manual register: **2,003 cases**, preserving all 1,971 previous rows byte-for-byte. Native Excel scenario/data-table definitions, automatic live table recalculation, multivariable Solver and actual provider/production acceptance remain unclaimed.

## Safe XLSX exchange — AW-064

File → Open and Save as Excel now stage a compatibility review before replacing the workbook or downloading a derived copy. Review and export a report of retained content, limitations and source-cache differences. Archive checks validate CRC, directories, XML and internal parts; supported values, formula text and typed caches survive fresh reopen. Source caches remain historical evidence; local recalculation controls the editor.

Open `dist/xlsx-exchange.html` or `XLSX-exchange-demo.html` in the local ZIP. Try `XLSX-exchange-fixture.xlsx`, `XLSX-stale-cache.xlsx` and deliberately invalid `XLSX-corrupt.xlsx`; native baseline and exact SDK evidence are also included. [End-to-end story and criteria](docs/XLSX-EXCHANGE-ACCEPTANCE.md), [browser evidence](docs/XLSX-EXCHANGE-BROWSER-EVIDENCE.md).

At the AW-064 release, regression was **5,418 tests across 201 files, zero failures or skips**; **62 new tests across two files**. Coverage: **51 locally accepted / 36 Partial / 21 Planned — 57 open**. Manual register: **2,035 cases**, preserving all 2,003 previous rows byte-for-byte; XF-001–032 remain Not run. Original-file retention and full unsupported-part preservation remain AW-065; alternate codecs remain AW-066. Actual licensed Excel and external provider acceptance remain unclaimed.

## Reviewed CSV/TSV exchange — AW-067

File → Open and Save as CSV now review delimiters, encodings, quotes/newlines, headers, column types and number/date conventions. Select Tab in export for TSV. Malformed records and type failures block import; deliberate conversions and protective text prefixes have a complete downloadable change ledger. Blank positions, identifiers and trailing records are retained under the visible schema and bounded processing limits.

Open `dist/delimited-exchange.html` or `Delimited-exchange-demo.html` in the local ZIP. Synthetic CSV/TSV, malformed and localized fixtures, native JSON and schema reports are included. [End-to-end story and feature criteria](docs/DELIMITED-EXCHANGE-ACCEPTANCE.md), [browser observations](docs/DELIMITED-EXCHANGE-BROWSER-EVIDENCE.md).

At the AW-067 release, regression was **5,491 tests across 203 files, zero failures or skips**; **73 new tests across two files**. Coverage: **52 locally accepted / 35 Partial / 21 Planned — 56 open**. Manual register: **2,067 cases**, with all previous 2,035 rows unchanged; DL-001–032 remain Not run. Actual Excel/provider acceptance remains unclaimed.

## Reviewed automation interoperability — AW-071

Review or File → Automation & Scripts inventories VBA/XLM/Office Script artifacts and separate source, preserves eligible original bytes, and translates declared literal range operations into reviewed workbook commands. Preview the approved worksheet/range and before/after values, then approve a single transaction with source-linked provenance and Undo/Redo. Unsupported code blocks the whole migration.

This is the documented **literal-ranges-v1 migration adapter**; it does not run arbitrary VBA/JavaScript, binary macros, COM, events, network or file calls. Original artifacts are inert native JSON attachments. XLSX includes resulting cells and an explicit native-only metadata warning. Actual licensed Office/runtime equivalence and external provider acceptance remain separate gates.

Open `dist/automation.html` or `Automation-demo.html` in the refreshed local ZIP. [Story and all feature criteria](docs/AUTOMATION-INTEROPERABILITY-ACCEPTANCE.md), [browser evidence](docs/AUTOMATION-INTEROPERABILITY-BROWSER-EVIDENCE.md), and Automation source/native/report fixtures are included.

At the AW-071 release, regression was **5,590 tests across 205 files, zero failures or skips**, including **99 new tests** (87 core and 12 UI). Coverage: **53 locally accepted / 34 Partial / 21 Planned — 55 open**. Manual register: **2,103 cases**, with all previous 2,067 rows preserved byte-for-byte; AM-001–036 remain Not run.

## Pattern checks — AW-079

Review → Pattern Checks (also in Audit toolkit) tests explicit invoice/payment keys, grouped sequence gaps, round amounts, configured weekends, supplied holidays and All/Any row rules. Review every source position, rule settings, matching groups, exact gap intervals, citations and counts; download current complete or failed evidence. Missing fields and invalid types fail explicitly. These are review candidates, not proof of error or fraud.

The reusable PatternAuditTools component and runPatternAudit SDK function are read-only. Existing MCP workbook_run_audit accepts kind patterns and explicit pattern settings, confined to the trusted host selection. It excludes notes/origins and requires formula dependencies to stay within the selected cells. No model-provider connection or editing authority is added.

Open `dist/pattern-checks.html` or `Pattern-checks-demo.html` in the refreshed local ZIP. [Story and all feature criteria](docs/PATTERN-CHECKS-ACCEPTANCE.md), [browser evidence](docs/PATTERN-CHECKS-BROWSER-EVIDENCE.md), and six Pattern source/configuration/report/evidence fixtures are included. Limits: 500 selected cells, 12 rules, eight key fields, 1000 supplied holidays, 2 MB complete evidence. Actual Excel reference, external LLM and production authorization/accessibility gates remain unverified.

At the AW-079 release, regression was **5,726 tests across 208 files, zero failures or skips**, including **136 new tests** (108 core, 15 UI, 13 official-SDK MCP). Coverage: **54 locally accepted / 33 Partial / 21 Planned — 54 open**. Manual register: **2,139 cases**; all previous 2,103 rows are preserved byte-for-byte and PT-001–036 remain Not run.

## Answer citations — AW-088

The AI assistant now links typed workbook claims to clickable selected cells/ranges. Inspect formulas, values, workbook revision and recorded original-row lineage; focus individual cells while retaining approved scope. Changed content/scope/calculation invalidates evidence. External claims have separate supplied HTTPS sources, explicitly unverified; legacy narrative and commentary are labeled unverified. Download a complete current JSON receipt. Local source binding is not proof of semantic entailment or original-file authenticity.

The provider-neutral proposal contract and existing MCP queue tool accept optional grounding. The browser captures source facts locally; model-visible MCP evidence excludes private notes and origins. Public SDK helpers and AssistantEvidence support custom hosts. No live provider or automatic editing authority is added.

Open `dist/answer-citations.html` or `Answer-citations-demo.html` from the local ZIP. [Story and all acceptance criteria](docs/CITATION-GROUNDING-ACCEPTANCE.md), [browser evidence](docs/CITATION-GROUNDING-BROWSER-EVIDENCE.md), and five compiled-SDK Citation fixtures are included.

At the AW-088 release, regression was **5,837 tests across 211 files, zero failures or skips**, including **111 new tests** (84 core/adapter, 14 UI, 13 official-SDK MCP). Coverage: **55 locally accepted / 32 Partial / 21 Planned — 53 open**. Manual register: **2,175 cases**; all previous 2,139 rows remain byte-for-byte intact and CE-001–036 are Not run. Real external model/entailment, production access/retention and full browser/accessibility gates remain unverified.

## Dependency recalculation — AW-030

Formulas → Calculation Options now exposes Automatic/Manual calculation, explicit Calculate Now, numeric iteration controls and current/pending/error results. Committed scalar source edits refresh dependent formulas once while reusing unrelated results; metadata and arrays use conservative invalidation. Manual outputs are visibly pending, historical values are labeled separately, and AI/MCP refuse pending selected values. Calculation alone does not edit source inputs or history.

Imported formula caches remain typed native evidence with original settings and file hash. Review them separately against current results; stale cached numbers never substitute for errors. Imported iteration requires explicit local opt-in. XLSX writes supported calculation settings and current caches, omitting pending/app-specific failures. Bounded iteration starts from zero and reports convergence; the synchronous calculation path uses cooperative limits. AW-043 adds worker deadlines and cancellation to the browser editor and an opt-in worker service for other hosts. Excel iteration equivalence remains unverified.

Open `dist/recalculation.html` or `Calculation-recalculation-demo.html` in the local ZIP. Seven Calculation fixtures, [the complete user story/criteria](docs/RECALCULATION-ACCEPTANCE.md) and [browser observations](docs/RECALCULATION-BROWSER-EVIDENCE.md) are included. Public SDK exports revision tracking, calculation settings, current reports, explicit recalculation and CalculationTools for custom hosts.

At the AW-030 release, regression was **5,926 tests across 215 files, zero failures or skips**, including **89 new tests** (46 dependency/runtime, 25 exchange, 11 UI, 7 official-SDK MCP). Coverage: **56 accepted / 31 Partial / 21 Planned — 52 open**. Manual register: **2,211 cases**; all prior 2,175 rows remain byte-for-byte intact and RC-001–036 are Not run. Independent Excel/provider, production authorization and complete browser/accessibility acceptance remain unverified.

## Dates and business calendars — AW-035

Formulas → Calculation Options now controls explicit MDY/DMY date text and the IANA time zone used for TODAY/NOW. Every calculation records one UTC instant, local clock and offset; Calculate Now refreshes results without a source edit. Strict date parsing, Gregorian/1900 leap/month-end rules, completed/elapsed periods, day-count bases, holiday/weekend deadlines and time calculations have bounded local implementations and typed error handling. Holidays remain user-supplied. Unsupported calendars refuse explicitly. AW-042 adds the declared 1904 runtime policy. Native JSON preserves settings; XLSX warns that timezone/text order/clock evidence are not retained.

Open `dist/date-calendar.html` or `Date-calendar-demo.html` in the local ZIP. Six calendar fixtures, [the complete user story/criteria](docs/DATE-CALENDAR-ACCEPTANCE.md) and [browser observations](docs/DATE-CALENDAR-BROWSER-EVIDENCE.md) are included. Explicit policy and current-time evidence bind approved AI/MCP context; stale or malformed clock data refuse. Models still cannot write NOW/TODAY, change settings or bypass human review.

At the AW-035 release, regression was **6,124 tests across 219 files, zero failures or skips**, including **198 new tests** (162 date/runtime, 24 exchange/aging/provider integration, 6 UI, 6 official-SDK MCP). Coverage: **57 accepted / 30 Partial / 21 Planned — 51 open**. Manual register: **2,247 cases**; all previous 2,211 rows remain byte-for-byte intact and DT-001–036 are Not run. Gregorian/1900 local conventions, strict text parsing and documented DATEDIF/day-count rules are not full Excel parity; independent Excel/provider and production acceptance remain unverified.

## Financial schedules — AW-037

Twenty financial functions cover debt/lease payments and interest/principal splits, periodic/dated valuations, return rates, four depreciation methods, regular-coupon bond price/yield and accrued interest. Eight additions raise local functions to **377**, with **373 assistant-safe names**. Formulas → Calculation Options shows each financial convention, failure reason and residual-checked rate candidates; finite search never proves uniqueness. Native assumptions/formulas persist and ordinary XLSX recalculates; export solver evidence separately.

Open `dist/financial-schedules.html` or `Financial-schedules-demo.html` in the local ZIP. The six-sheet synthetic example contains 220 formulas, four fixture/evidence files and [the full user story/acceptance checklist](docs/FINANCIAL-SCHEDULES-ACCEPTANCE.md). [Browser observations](docs/FINANCIAL-SCHEDULES-BROWSER-EVIDENCE.md) cover source edits, Undo/Redo, roots/errors, read-only Manual review, securities and local cited answers. No real provider or independent Excel was executed.

At the AW-037 release, regression was **6,308 tests across 223 files, zero failures or skips**, including **184 new tests** (141 financial function/solver, 24 whole-schedule/exchange/scope, 6 UI, 5 MCP and 8 added baseline formula cases). Coverage: **58 accepted / 29 Partial / 21 Planned — 50 open**. Manual register: **2,291 cases**; prior 2,247 records are preserved and FS-001–036/FN-370–377 are Not run. Financial/date/root-search limits and independent acceptance remain explicit.

## Data entry validation — AW-047

Data → Data Validation now edits lists, decimals/whole numbers, dates/times, text length and relative custom formulas with input guidance, dropdowns and Stop/Warning/Information messages. The same evaluator checks manual entry, paste, SDK writes and approved assistant changes. A failing batch changes nothing. Warning approval applies the whole batch once and records the decision; Undo/Redo restores it. Read-only users can inspect rules and existing values.

Open `dist/validation-entry.html` or `Validation-entry-demo.html` in the local ZIP. The synthetic host provides standard, invalid, warning and Manual fixtures. Existing-value inspection identifies seven seeded invalid cells across 21 checks, including a required blank, with navigation and evidence export. Inspection does not publish Manual formula results or modify source. [Full user story and 36 feature criteria](docs/DATA-VALIDATION-ACCEPTANCE.md); [browser observations](docs/DATA-VALIDATION-BROWSER-EVIDENCE.md). Range-backed lists, formula-valued numeric bounds, specialized preview warning confirmation and full Excel equivalence remain bounded as documented.

At the AW-047 release, regression was **6,411 tests across 227 files, zero failures or skips**, including **103 new tests** (62 core, 24 integration/exchange, 13 UI and four MCP). Coverage: **59 accepted / 28 Partial / 21 Planned — 49 open**. The manual register has **2,327 cases**; prior 2,291 records are preserved and DV-001–036 remain Not run. Actual external model/client, independent Excel and production acceptance are unverified.

## Clean and reshape fields — AW-048

Data → Clean & Reshape creates a reviewed recipe for trim/control cleanup, case/replacement, quoted multi-delimiter or fixed-width splitting, concatenation, explicit type conversion, gap filling, field order and example-based patterns. Preview shows original raw values, typed results, source locators and parse failures. Explicit Apply creates a separate worksheet and exceptions appendix; source remains intact and Undo/Redo is atomic. Native JSON retains original snapshots and reusable recipes. XLSX preserves supported current values/notes and warns about native-only evidence.

Open `dist/field-reshape.html` or `Field-reshape-demo.html` in the local ZIP. Review the prepared ten-step recipe: four records, nine fields and five retained exceptions. The exported `FieldTransformTools` component accepts an inert host-proposed `initialRequest`, so provider-neutral integrations can reuse the same review. Nine existing MCP tools retain their current approval boundaries; historical source evidence is excluded from model context. [Story and all 36 criteria](docs/FIELD-TRANSFORM-ACCEPTANCE.md); [observed browser checks](docs/FIELD-TRANSFORM-BROWSER-EVIDENCE.md).

At the AW-048 checkpoint, regression was **6,524 tests across 231 files, zero failures or skips**, including **113 new tests** (80 core, 16 integration/exchange, 14 UI and three MCP). Coverage: **60 accepted / 27 Partial / 21 Planned — 48 open**. The manual register has **2,363 cases**; prior 2,327 records remain unchanged and FT-001–036 remain Not run. Limits include 500 source cells, 12 steps, 50 fields and 5000 output cells. Bounded inference does not claim full Excel Flash Fill/Power Query equivalence. Actual provider/client, independent Excel and production acceptance remain unverified.

## AW-040 expanded specialist calculations

18 additions bring the catalog to **395 local / 391 reviewed-assistant functions**. LINEST/LOGEST, multivariate TREND/GROWTH, t/z tests, skewness, trimmed means, multiple modes, logical/text-aware statistics, permutations and per-record database formula criteria are delivered. Open **Specialist-demo.html** in the local package or `/specialist-functions.html` on the local server. [Story, full contracts, limits and acceptance criteria](docs/SPECIALIST-FUNCTIONS-REMEDIATION.md).

At the AW-040 checkpoint, regression was **6,661 tests / 235 files**, zero failures/skips; public SDK/MCP, portable builds and 24 SDK examples rebuilt. 54 new validation templates bring the register to 2,417. **AW-040 remains Partial** until reference results are verified in a named Microsoft Excel version. No actual external-provider or production connector execution is claimed.

## AW-042 precision and date systems

Calculation Options now declares binary precision and monetary ROUND behavior, and supports both 1900 and 1904 date systems throughout calculations, formatting, preparation and file exchange. XLSX preserves the source date system, raw serials/fractions and text identifiers. Precision-risk and precision-as-displayed warnings make conversion limits explicit. Undo/Redo, native reopening, read-only review and approved MCP contexts are covered.

Open **Precision-date-demo.html** in the local package, or `/precision-date.html` on the local server. [End-to-end story and 30 validation cases](docs/PRECISION-DATE-ACCEPTANCE.md), [browser observations](docs/PRECISION-DATE-BROWSER-EVIDENCE.md), and [verification record](docs/PRECISION-DATE-VERIFICATION.json). The full regression passed **6,760 tests in 239 files**; the final warning clarification also passed **10 UI follow-up tests**. All 25 SDK example hosts and portable pages are rebuilt. **AW-042 is locally accepted**; independent Excel and actual external LLM execution remain unverified.

## AW-043 calculation resource controls

The editor calculates in an isolated worker, with deadlines, cancellation and recovery. Calculation Options exposes shared work and cumulative working-data budgets; failed runs preserve inputs and do not publish partial formula results as current. RAND/RANDBETWEEN and NOW/TODAY follow a declared snapshot policy, and seeded audit sampling remains separate. Hosts can use browser/Node worker sessions and pass an exact private calculation snapshot to the MCP bridge.

Open **Calculation-resources-demo.html** in the local package, or `/calculation-resources.html` on the local server. [Story, limits and 32 manual validation cases](docs/CALCULATION-RESOURCES-ACCEPTANCE.md), [browser observations](docs/CALCULATION-RESOURCES-BROWSER-EVIDENCE.md), and [verification record](docs/CALCULATION-RESOURCES-VERIFICATION.json). The full suite passed **6,812 tests across 243 files**. All 26 SDK example hosts and portable pages are rebuilt. Memory accounting is not an exact browser-heap cap; direct synchronous consumers remain cooperative. Independent Excel and actual external LLM execution remain unverified. Earlier sections record historical checkpoints.

## AW-045 reviewed record sorting

Data → Custom Sort adds a reviewed physical reorder of complete worksheet records. Use up to 64 ordered levels, numeric/date/text modes, case sensitivity and custom text lists. Preview the row mapping, original source IDs, formula rewrites and metadata changes; partial selections require expansion and hidden rows require explicit inclusion. Unsafe/protected/stale operations refuse without changing the source. Apply uses one history revision with Undo/Redo; native JSON retains source identities. Existing temporary view sorting remains available.

Open **Population-sort-demo.html** in the local ZIP, or `/population-sort.html` on the server. The compiled SDK exports SortTools and plan/apply helpers; MCP exposes the policy while preserving human review and stale-context rejection. [Story, bounds and 32 validation templates](docs/POPULATION-SORT-ACCEPTANCE.md), [browser observations](docs/POPULATION-SORT-BROWSER-EVIDENCE.md), [verification](docs/POPULATION-SORT-VERIFICATION.json). The full suite passed **6,869 tests across 246 files**, including 57 new sort tests; all 27 SDK hosts are rebuilt. Base coverage is **55 accepted / 10 Partial / 11 Planned — 21 open**. Row sorting is bounded to 10,000 rows and 100,000 record cells. Color/icon and left-to-right sorts, independently verified Excel behavior and actual external model/client execution remain outside this local acceptance. Earlier checkpoints are historical.

## AW-046 population filtering

Data → Column Filters adds header controls, typed value selections and integrated value slicers, text/wildcard/number/date rules, custom AND/OR and an advanced OR-of-AND builder. The grid displays visible/total and hidden-match counts. Empty results preserve all source records. Filtering does not mutate cells, IDs, physical order or history; edits pause session views until reviewed. CSV/TSV review and Audit Checks offer explicit full/visible scope and original source citations. Native JSON/XLSX/print/other tools/AI selections remain full-source with disclosure.

Open **Population-filter-demo.html** in the local ZIP or `/population-filter.html` on the server. [Original contract, end-to-end story and 32 feature cases](docs/POPULATION-FILTER-ACCEPTANCE.md), [browser observations](docs/POPULATION-FILTER-BROWSER-EVIDENCE.md), [verification](docs/POPULATION-FILTER-VERIFICATION.json). Full regression: **6,942 tests across 249 files**, including 73 new filter tests; all 28 SDK hosts rebuilt. Base coverage: **56 accepted / 9 Partial / 11 Planned — 20 open**. Filtering is bounded to 10,000 data rows, 100,000 cells including headers, 64 columns and explicit comparison budgets. Session filter views, Excel AutoFilter persistence, view-aware SUBTOTAL/AGGREGATE, actual Excel and external model/client boundaries are documented. Earlier checkpoints are historical.

## AW-050 and saved filter views — 18 September 2026

**7,026 tests across 255 files passed**, with no failures or skips, including 43 new join tests and 41 saved-view tests. TypeScript, the public React/MCP library, the Node worker, portable pages and all 30 validation hosts build. The compiled SDK checks all six join modes, threshold boundaries, unchanged sources, native output and supported XLSX values/notes. Browser checks confirm explained matching, anti output, fresh preview, Apply/Undo/Redo and native saved-view reopening. [Current verification](docs/POPULATION-COMBINE-VERIFICATION.json) and [browser observations](docs/POPULATION-COMBINE-BROWSER-EVIDENCE.md). Earlier counts and hashes above are historical checkpoints. Actual external models, Excel reference comparisons and production acceptance remain unverified.

## AW-051 — 18 September 2026

**7,174 tests across 259 files passed**, zero failures/skips, including 148 new recipe/expression/UI/MCP tests. Public SDK/MCP modules, Node calculation worker, portable pages and all 31 validation hosts build. Browser checks verify the full 13-step recipe, unchanged sources, Publish/Undo/Redo, native reopening, identical repeat results, immutable version editing and a failed step with downstream publication blocked. [Current verification](docs/TRANSFORMATION-RECIPES-VERIFICATION.json) and [browser evidence](docs/TRANSFORMATION-RECIPES-BROWSER-EVIDENCE.md). Earlier counts and artifact hashes are historical checkpoints.

## AW-059 — 18 September 2026

**7,274 tests across 262 files passed**, zero failures/skips, including 100 new consolidation/core/UI/MCP tests. Public SDK, MCP, Node calculation worker, portable application and all 32 validation hosts build. Browser checks verify eight contributions totaling 315, nested subtotals, Publish/Undo/Redo, native reopening without approval, reapproval, a live change to 335 with visible stale context, and failed currency refresh preserving prior output. [Current verification](docs/CONSOLIDATION-VERIFICATION.json) and [browser evidence](docs/CONSOLIDATION-BROWSER-EVIDENCE.md). Earlier counts and hashes are historical checkpoints.

## AW-065 — 18 September 2026

**7,320 tests across 265 files passed**, zero failures/skips, including 46 new preservation/core/UI/MCP tests. All 33 compiled SDK validation hosts build. The browser journey verifies inventory filtering, import/cancel, Undo/Redo, editing, native reopening, read-only source review and final export outcomes. The original download action was exercised; the in-app browser did not expose a completion event, so OS-level persistence remains unverified. Exact original Blob bytes and passive part bytes passed automated checks. [Current execution](docs/XLSX-PRESERVATION-VERIFICATION.json) and [browser evidence](docs/XLSX-PRESERVATION-BROWSER-EVIDENCE.md). Earlier counts and hashes are historical checkpoints.

## AW-098 secure file processing — historical checkpoint

**7,397 tests across 269 files passed**, zero failures/skips, including 77 new security, worker, UI and MCP checks. All 34 compiled SDK validation hosts build. Twelve compiled SDK checks cover isolated imports, safe literal export, worker denial/deadline/recovery, original-buffer retention and unchanged MCP authority. Browser checks verify hostile-file rejection, responsive stalled-worker termination, Cancel, reviewed CSV import, Undo/Redo and read-only refusal. [Current execution](docs/FILE-SECURITY-VERIFICATION.json) and [browser observations](docs/FILE-SECURITY-BROWSER-EVIDENCE.md). Independent security review, strict-CSP target deployment and actual external model/client testing remain external gates. Earlier counts and hashes are historical checkpoints.

## AW-101 accessibility and localization — historical checkpoint

**7,457 tests across 273 files pass**, including 60 new accessibility/localization checks. Twelve compiled SDK checks pass; all 35 validation hosts build. Twenty-four English/German browser scans report zero violations, with incomplete contrast checks retained for review. AW-101 remains **Partial** pending actual assistive-technology, native zoom, full locale and independent WCAG acceptance. Base coverage stays **61 accepted / 4 Partial / 11 Planned — 15 open**. Full coverage stays **69 accepted / 18 Partial / 21 Planned — 39 open**. [Execution](docs/ACCESSIBILITY-VERIFICATION.json), [browser observations](docs/ACCESSIBILITY-BROWSER-EVIDENCE.md). The 2,733 manual templates remain Not run or Blocked. Earlier counts/hashes are historical checkpoints.

## AW-102 interactive performance — historical checkpoint

**7,491 tests across 275 files pass**, with zero failures or skips, including 34 new performance/correctness checks. Nineteen compiled public-SDK checks pass. The final local five-sheet, 100,000-cell/10,000-formula fixture records 30 observations per operation: edit p95 **105.8 ms**, scroll p95 **41.7 ms**, undo/redo p95 **67.7 ms**, and warmed mount p95 **235.7 ms**. The longest observed steady interaction task is **84 ms**. One scroll took **57.3 ms**, above the original 50 ms individual-response threshold. Edit p95 also remains above its 100 ms target. The faster local machine and single Chromium browser do not establish the prescribed reference-profile acceptance.

AW-102 remains **Partial**. The original four-core/16-GB current/previous Chrome, Edge and Safari matrix, remaining edit and scroll response targets, durable host-save and larger synchronous analysis measurements, and complete heap/GC/sustained-session traces remain open. The two-million-cell import is explicitly excluded by the visible 100,000-cell capacity. Base coverage remains **61 accepted / 4 Partial / 11 Planned — 15 open**; full coverage remains **69 accepted / 18 Partial / 21 Planned — 39 open**.

[Reproduce the benchmark](docs/INTERACTIVE-PERFORMANCE-ACCEPTANCE.md). Importing the supported 100,000-cell CSV takes 467.8 ms p95; the 99,999-value aggregate takes 218.2 ms p95 and returns 10000499996. Cancellation settles within 0.8 ms in the local run. All 2,733 prior manual templates remain, with PERF-001–024 bringing the total to 2,757. Earlier counts and artifact hashes are historical.
