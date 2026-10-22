# Ledgerglass

**An embeddable, AI-connected workbook for the browser.**

Ledgerglass is a TypeScript/React spreadsheet component with an Excel-style ribbon, formula bar, worksheet grid and sheet tabs. Run it as a standalone application, embed it in a larger application, or connect its workbook model to your own LLM through a provider-neutral adapter and Model Context Protocol (MCP).

![Ledgerglass feature overview: workbook editing, formulas, data analysis, file interoperability, AI, embedding and deployment](docs/assets/ledgerglass-feature-overview.png)

The illustration summarizes the feature families, including capabilities that require host integration. See the [complete feature index](docs/FEATURE-INDEX.md) for all 108 requirement IDs and their status.

> **Scope:** this is a substantial working implementation with documented limits, not full Microsoft Excel parity. The base Excel coverage register records **70 locally accepted, 5 Partial and 1 Planned** requirements. The full product register, including audit and AI extensions, records **78 locally accepted, 19 Partial and 11 Planned**. Acceptance applies to the documented local scope; production, external-provider and independent accessibility acceptance remain separate.

## Contents

- [Quick start](#quick-start)
- [Capabilities](#capabilities)
- [Embed in React](#embed-in-react)
- [AI assistant and MCP](#ai-assistant-and-mcp)
- [Data, persistence and file exchange](#data-persistence-and-file-exchange)
- [Docker and scaling](#docker-and-scaling)
- [Limits and compatibility](#limits-and-compatibility)
- [Testing and validation](#testing-and-validation)
- [Repository layout](#repository-layout)
- [Documentation](#documentation)
- [Troubleshooting](#troubleshooting)
- [Contributing and licensing](#contributing-and-licensing)

## Quick start

### Requirements

- Node.js **22.16 or later in the Node 22 family** and npm. The SQLite connector demonstration requires Node 22.16+.
- A modern desktop browser. Browser APIs, clipboard, worker, storage and accessibility support must be validated in your target host.
- Docker Engine with Compose v2 only if you want to use the container setup.

From a checkout of this repository:

```sh
npm ci
npm run dev -- --port 4177
```

Open **http://127.0.0.1:4177/**. No model API key or backend service is required for the local demo.

| URL | Experience |
| --- | --- |
| `/` | Standalone workbook; opens blank or restores locally saved work |
| `/?embedded` | Workbook inside the Test example application |
| `/?temporary` | Temporary workbook without browser persistence |
| `/?demo` | Synthetic audit demonstration data |

Select **File → Open** and load `examples/Advanced-features-demo.json` for tables, a PivotTable, charts, shapes, formulas and page layout. Its control total is **54,400**. The examples are synthetic validation fixtures, not client workbooks.

### Build and run locally

```sh
npm run build:portable
npm start
```

Open **http://127.0.0.1:4173/**. `npm start` serves the previously built `dist/` directory; it does not compile the application.

The portable build also creates `Ledgerglass.html` and `Embedded-Workbook.html`. These files include the application code and styles and can be opened directly. Local file storage and browser coordination vary by browser; localhost or HTTPS provides the more predictable experience. Export native JSON as a portable backup.

This source repository omits installed dependencies, generated bundles, portable HTML and large handoff archives. The commands above regenerate the application. Specialist validation pages have individual builders under `scripts/`; for example, after `npm run build:lib` and `npm run build`, run `node scripts/build-collaboration-host.mjs`, then `npm start` and open `/collaboration.html`.

## Capabilities

The links below document supported behavior and its boundaries; a feature family does not imply every feature of the equivalent Microsoft product.

| Area | Delivered capability |
| --- | --- |
| Workbook lifecycle | Create/open/save, sheet management, protection, revisions, undo/redo, recovery and concurrent-instance conflict handling |
| Familiar editing | Keyboard navigation, multi-area selections, formula range selection, autocomplete after `=`, copy/cut/paste, fill, find/replace, freeze/split panes, outlines and row/column AutoFit |
| Formulas | 395 catalogued local function names; arithmetic, logic, lookup, text, date/calendar, financial, statistical/database families, dynamic arrays, names and reusable formulas within declared limits |
| Calculation review | Recalculation controls, dependency/error inspection, watches, evaluation steps, precision/date-system policies and bounded worker execution |
| Formatting and layout | Cell/number styles, themes, conditional formatting, merges, pictures/shapes, page setup, print areas/titles, print preview and schedule PDF export |
| Tables and preparation | Structured tables, calculated columns, safe sorting/filtering, saved views, validation, text cleanup/extraction, duplicate review, append/join and repeatable transformation recipes |
| Connections | Registered external sources, reviewed single-source refresh and atomic Refresh All; server-side JSON/OData, SQLite, CSV/TSV folder and ODBC-compatible adapter contracts |
| Models and analysis | Relational models, a documented 20-function DAX subset, PivotTables, pivot filtering/drill-through, charts/sparklines, consolidation/subtotals, scenarios, Goal Seek and bounded Solver/statistical tools |
| Interoperability | Native workbook JSON, XLSX, delimited files, supported structured sources, original-source retention, explicit export-loss reports and host adapters for alternate formats/document extraction |
| Links and extensions | Reviewed external-workbook snapshots, versioned custom-function/add-in adapters, reusable report outputs and bounded automation translation with inert unsupported source retention |
| Audit extensions | Notes/tickmarks, evidence links, comparisons, aging/materiality screens, sampling and duplicate/gap/unusual-pattern checks; retained alongside base spreadsheet functionality |
| Assistant | Chat built with assistant-ui, file staging/previews, history, Stop/Retry/Regenerate/Copy, selected-cell citations and explicit before/after proposal review |
| LLM interoperability | Local deterministic demo, provider-neutral HTTP/custom adapters and nine scoped MCP tools; proposals require human approval before cell changes |
| Collaboration | Host-authorized shared editing contracts, conflict review, presence, discussions and offline recovery; production authority/storage/transport are host responsibilities |
| Embedding and operations | React SDK, browser/Node calculation integrations, parser workers, localization interfaces, diagnostics, stateless web container, Compose and Kubernetes templates |

See [all features and acceptance criteria](docs/ALL-FEATURE-REGISTER.csv), the [base Excel coverage sheet](docs/BASE-EXCEL-COVERAGE.xlsx), and the [validation checklist](docs/VALIDATION-CHECKLIST.md).

## Embed in React

### Build the SDK

```sh
npm run build:lib
npm pack
```

Install the resulting package in your React 19 host:

```sh
npm install /absolute/path/to/ledgerglass-workbook-0.1.0.tgz
```

The package is named `@ledgerglass/workbook`. It is marked `private` and is not assumed to be available from the public npm registry. `npm pack` creates a local tarball. Preserve every generated file in `lib/`, including lazy chunks and workers.

### Minimal editor

```tsx
import { AuditWorkbook } from '@ledgerglass/workbook';
import '@ledgerglass/workbook/style.css';

export function WorkbookEditor() {
  return (
    <div style={{ height: '80vh', minHeight: 420 }}>
      <AuditWorkbook />
    </div>
  );
}
```

The default assistant is a local deterministic demonstration. For an application integration, use the [complete React host example](deploy/ReactHost.example.tsx): it accepts an authorized initial workbook, a host store, an actor, persistence notifications and HTTP assistant/source adapters.

`onChange` signals an in-memory change; it is not a durable-save acknowledgement. Implement the documented `WorkbookStoreLike` contract and wait for successful persistence before showing “Saved.” Load and authorize the workbook before mounting, and use a workbook identity key when switching documents. The host owns user identity, permissions, tenant boundaries and durable storage.

| Package entry | Purpose |
| --- | --- |
| `@ledgerglass/workbook` | React editor, workbook model, operations, types and browser integration APIs |
| `@ledgerglass/workbook/style.css` | Required editor styles |
| `@ledgerglass/workbook/mcp` | Server-side MCP factory and proposal guards |
| `@ledgerglass/workbook/calculation-node` | Node calculation integration |
| `@ledgerglass/workbook/file-worker.js` | Packaged file-processing worker |
| `@ledgerglass/workbook/source-gateway` | Server-side source gateway |
| `@ledgerglass/workbook/source-connectors` | Server-side connector implementations |

Keep server-only entries out of browser bundles. Review the [developer handoff](docs/DEVELOPER-HANDOFF.md), [calculation worker contract](docs/CALCULATION-RESOURCES-ACCEPTANCE.md), and [collaboration integration](docs/COLLABORATION-INTEGRATION.md) before adapting production services.

An iframe can embed the served editor if its CSP permits the parent origin. It is a visual embed; no general-purpose `postMessage` workbook API is supplied. Use the React SDK for programmatic integration.

## AI assistant and MCP

### Connect your own model

```tsx
import { AuditWorkbook, createHttpAdapter } from '@ledgerglass/workbook';
import '@ledgerglass/workbook/style.css';

const adapter = createHttpAdapter({
  endpoint: '/api/workbook-assistant',
  chat: true,
});

export function ConnectedWorkbook() {
  return <div style={{ height: '80vh' }}><AuditWorkbook adapter={adapter} /></div>;
}
```

Your backend implements the endpoint. Authenticate and authorize the workbook and selected context, translate the request into your provider's API, and return the validated `LLMProposal` contract. Keep provider keys on the server. The component checks revision/scope, displays citations and before/after patches, and revalidates when the user chooses Apply.

The chat supports text, CSV/TSV, JSON, PDF, DOCX, XLSX and common image attachments within explicit limits. The local demo acknowledges staged files but does not use an LLM to interpret them. The host must opt into chat, implement document handling and enforce authorization, request limits, scanning and retention. Responses currently complete as a single proposal; token streaming is not implemented.

Read the [LLM protocol](docs/LLM-PROTOCOL.md), [chat integration guide](docs/CHAT-UI-INTEGRATION.md), and [citation acceptance](docs/CITATION-GROUNDING-ACCEPTANCE.md).

### MCP server

The server provides scoped context/range reads, formula inspection/evaluation/contracts/compatibility, audit checks, capability discovery and proposal queuing. It cannot apply changes or widen the host-approved selection.

After `npm run build:lib`:

```sh
npm run mcp -- \
  --workbook /absolute/path/workbook.json \
  --sheet worksheet-id-from-native-json \
  --range A1:D20 \
  --outbox /absolute/path/proposals.json
```

The sheet argument is the stable JSON sheet ID, not its display name. Export a native workbook through File → Save workbook data. The input file and proposal outbox must be separate.

This reference uses stdio and a file bridge. It does not synchronize an open browser window automatically or expose a public HTTP MCP endpoint. The embedding host connects its authorized workbook, proposal queue and review UI through the server factory. See [MCP tools, configuration and integration](docs/MCP.md).

## Data, persistence and file exchange

Native JSON preserves the application's workbook state, retained evidence and supported metadata more fully than exchange formats. XLSX import/export implements a documented subset, inventories unsupported original parts and reports losses when producing a derived file. Recovering an exact original is distinct from preserving every feature in an edited export.

CSV/TSV exchange includes delimiter/encoding/type review and formula-injection controls. Alternate formats and document extraction use explicit supported paths or host converters. Macro/source retention does not execute native VBA, Office Scripts, ActiveX or OLE code.

For real local connector demonstrations:

```sh
npm run build:lib
npm run build
node scripts/build-source-connectors-host.mjs
npm run demo:sources
```

Open **http://127.0.0.1:4182/external-connectors.html**. The demo uses local HTTP, SQLite and CSV folder sources. Real ERP services, ODBC drivers and production identities require your environment. See [source connector integration](docs/SOURCE-CONNECTOR-INTEGRATION.md) and [Refresh All](docs/AW-052-REFRESH-ALL.md).

Browser local/session storage is useful for demonstrations and recovery, but is not a shared database or backup service. Production integrations need a revision-aware durable store, authorization, concurrency handling, backups and a verified restore process.

## Docker and scaling

### Local container deployment

```sh
docker compose up --build -d --wait --scale web=2
```

Open **http://127.0.0.1:8080/?embedded**. The Compose gateway balances requests across stateless web replicas.

```sh
docker compose down
```

Or build and run just the web image:

```sh
docker build --target web -t ledgerglass-workbook:0.1.0 .
docker run --rm --read-only --cap-drop ALL \
  --security-opt no-new-privileges:true \
  -p 127.0.0.1:8080:8080 ledgerglass-workbook:0.1.0
```

The `mcp` Docker target packages the scoped stdio server separately. The web image serves immutable assets with a non-root Node runtime, readiness/health endpoints, gzip support and graceful shutdown. It does not provide authentication, workbook database storage or a model proxy.

### Configuration

| Variable | Default | Purpose |
| --- | --- | --- |
| `LEDGERGLASS_PORT` | `8080` | Compose localhost gateway port |
| `HOST` / `PORT` | `0.0.0.0` / `8080` | Container HTTP listener |
| `FRAME_ANCESTORS` | `'self'` | Explicit allowed iframe parent origins |
| `CONNECT_ORIGINS` | empty | Additional CSP connection origins; does not configure adapters/CORS |
| `SHUTDOWN_TIMEOUT_MS` | `10000` | HTTP connection drain limit |
| `NODE_IMAGE` / `NGINX_IMAGE` | See Docker/Compose files | Build/gateway image overrides; pin approved digests for releases |

Example:

```sh
FRAME_ANCESTORS="'self' https://test.example.com" \
  docker compose up --build -d --wait --scale web=2
```

Kubernetes manifests include a two-replica Deployment, Service, disruption budget and HPA template for 2–10 replicas. Supply your registry image, namespace, TLS/ingress, Metrics API and platform policies. Scaling the asset server does not raise workbook limits or scale your database/model services. Shared backend state and retained assets across rolling releases require host implementation.

The [Docker and scaling guide](docs/DOCKER-AND-SCALING.md) covers architecture, configuration, React integration, MCP isolation, Kubernetes, rollout/rollback and developer acceptance. Container/cluster execution was **not verified** in the recorded local handoff. The included GitHub Actions workflow runs container validation on pull requests or manual dispatch; it does not publish images or deploy the application.

## Limits and compatibility

| Limit | Current local contract |
| --- | --- |
| Worksheet address space | 1,048,576 rows × 16,384 columns, through XFD1048576 |
| Sheets | No fixed tab count; memory and workbook/file budgets apply |
| Stored cells | **100,000 per workbook**, including formatting-only or note-only cells |
| Cell/formula text | 32,767 / 4,096 UTF-16 units |
| Referenced formula rectangle | 20,000 cells; additional work/depth budgets apply |
| Dynamic array output | 10,000 cells per array |
| Assistant context/proposal | 500 selected cells / 500 proposed patches |
| XLSX/CSV/TSV import | 10 MiB; archive expansion and parser budgets also apply |
| Native JSON save/reopen | 50 MiB |
| Chat attachments | Up to four per message, 5 MB per file, 10 MB across request attachments |

Excel-sized addresses do not mean a fully populated Excel-sized worksheet can be processed. The [capacity source](src/core/capacity.ts), per-feature contracts and file-processing policies define the supported workload.

### Remaining base Excel requirements

| ID | Remaining gap |
| --- | --- |
| AW-040 | Named Excel-version reference acceptance and broader specialist-function fidelity |
| AW-052 | Native ODBC, real vendor sources, production authorization and host acceptance; scheduled/incremental refresh remains outside this delivery |
| AW-101 | Independent assistive-technology/zoom/device acceptance, complete localization and unresolved contrast checks |
| AW-102 | Required hardware/browser performance matrix, scroll budget and production-host measurements |
| AW-103 | Deployed diagnostics/alerts, durable backup and restore acceptance |
| AW-053 | Power Query M compatibility requirements remain Planned |

The [coverage summary](docs/COVERAGE-SUMMARY.md) details these six gaps. The [full register](docs/ALL-FEATURE-REGISTER.csv) separately tracks 30 open requirements across the complete product. Neither table is a claim of universal Microsoft compatibility, production readiness or accessibility certification.

## Testing and validation

```sh
npm run build:lib
npm run build:portable
npm test
npm run test:deployment
```

Additional checks:

```sh
# Browser journeys: install a Playwright browser once if needed.
npx playwright install chromium
npm run test:e2e

# Requires a working Docker daemon and image-registry access.
npm run test:containers

# Build, test and reconcile recorded release claims.
npm run verify:release
```

The Playwright config starts a development server on port 4173 and supports `CHROME_PATH`. Close any unrelated server on that port before testing. Container checks create and remove their own disposable Compose project; Kubernetes and actual provider acceptance require separate target environments.

**Recorded product evidence:** the 19 September 2026 checkpoint reports **8,734 passing tests across 326 files** and **14 deployment-server tests**. These are dated local records, not a claim that GitHub CI, Docker, Kubernetes or external LLM integration has run successfully. The manual checklist contains **3,167 templates**, with none automatically marked manually Passed. Consult the individual verification records for exact environments and exclusions.

| Validation resource | Use |
| --- | --- |
| [Validation checklist](docs/VALIDATION-CHECKLIST.md) | End-to-end user journeys and execution instructions |
| [Test-case register](docs/VALIDATION-TEST-CASES.csv) | Individual cases, expected outcomes and recorded status |
| [Base Excel feature matrix](docs/FEATURE-MATRIX.csv) | Requirement-level scope and acceptance |
| [Full product register](docs/ALL-FEATURE-REGISTER.csv) | All requirements, including audit/AI/embedding |
| [Refresh All evidence](docs/AW-052-REFRESH-ALL-VERIFICATION.json) | Latest recorded full-suite product checkpoint |
| [Docker packaging evidence](docs/DOCKER-PACKAGING-VERIFICATION.json) | Build/runtime/consumer checks and explicit unexecuted deployment gates |
| [Historical release notes](HISTORY.md) | Earlier remediation checkpoints; later checkpoints supersede earlier counts |

## Repository layout

```text
src/
  components/        Workbook UI and feature dialogs
  core/              Workbook model, formulas and deterministic operations
  integrations/      Files, LLM, persistence, sources and collaboration
  mcp/               Scoped MCP server
scripts/              Worker/SDK builders, demos and validation-host generators
tests/                Unit, component, integration and browser tests
examples/             Synthetic workbooks, source fixtures and expected outputs
docs/                 Requirements, acceptance stories, guides and evidence
deploy/               Static runtime, React example, tests and Kubernetes templates
.github/workflows/    Container verification workflow
Dockerfile            Separate web and MCP build targets
compose.yaml          Local gateway and scalable web service
```

`src/generated/`, `lib/`, `dist/`, `dist-container/`, portable HTML, dependencies and scratch reports are build outputs. The repository retains the generators and source fixtures. Historical verification files may refer to the original preparation machine or packaged artifacts; those paths are evidence, not paths required by a fresh checkout.

## Documentation

- **Start/integrate:** [developer handoff](docs/DEVELOPER-HANDOFF.md), [React host](deploy/ReactHost.example.tsx), [Docker and scaling](docs/DOCKER-AND-SCALING.md).
- **Requirements:** [full specification](docs/REQUIREMENTS.md), [feature index](docs/FEATURE-INDEX.md), [current gaps](docs/COVERAGE-SUMMARY.md), [base coverage workbook](docs/BASE-EXCEL-COVERAGE.xlsx).
- **AI:** [LLM protocol](docs/LLM-PROTOCOL.md), [chat UI](docs/CHAT-UI-INTEGRATION.md), [MCP](docs/MCP.md), [citation grounding](docs/CITATION-GROUNDING-ACCEPTANCE.md).
- **Data:** [external sources](docs/EXTERNAL-SOURCES-INTEGRATION.md), [connectors](docs/SOURCE-CONNECTOR-INTEGRATION.md), [data models](docs/DATA-MODEL-INTEGRATION.md), [DAX subset](docs/DAX-INTEGRATION.md), [file preservation](docs/XLSX-PRESERVATION-ACCEPTANCE.md).
- **Host acceptance:** [collaboration](docs/COLLABORATION-INTEGRATION.md), [accessibility/localization](docs/ACCESSIBILITY-LOCALIZATION-ACCEPTANCE.md), [performance](docs/INTERACTIVE-PERFORMANCE-ACCEPTANCE.md), [recovery runbook](docs/RELIABILITY-RUNBOOK.md), [file security](docs/FILE-SECURITY-ACCEPTANCE.md).

## Troubleshooting

| Symptom | Check |
| --- | --- |
| `npm start` returns 404 | Run `npm run build:portable` first; generated distributions are not committed |
| SDK import cannot find `lib/` | Run `npm run build:lib`, then pack/install the complete package |
| Worker or lazy chat load fails | Preserve generated chunks/workers and review host asset paths/CSP |
| Blank or clipped embedded editor | Give the parent a definite height and import the package CSS |
| Iframe is blocked | Set `FRAME_ANCESTORS` to the exact authorized parent origin and serve over HTTPS |
| Assistant/source endpoint returns 404 | The static server has no backend proxy; implement and route the host API |
| Local demo does not interpret an uploaded PDF | Connect a host model/document adapter; the demo has no external model |
| Source refresh fails | Check source registration, authorization, supported formats and row/field/batch limits |
| Local file saves are unreliable | Use localhost/HTTPS and export native JSON; verify browser storage support |
| Large workbook is refused | Check stored-cell, file, archive and calculation budgets; split or reduce the input |

## Contributing and licensing

Keep changes scoped and include relevant automated checks and synthetic reproduction fixtures. Preserve requirement IDs and historical acceptance evidence. Update the feature register and validation checklist when behavior changes; promote coverage only with evidence for the stated contract. Do not commit client data, model credentials, local workbooks or generated dependency folders.

Use your repository's approved private reporting channel for sensitive security findings. Production hosts must provide authentication, authorization, secret management, data retention and operational controls; client-side checks are not the security boundary.

The project is provided under the [MIT License](LICENSE). Dependency licenses are listed in [THIRD-PARTY-NOTICES.txt](THIRD-PARTY-NOTICES.txt). Microsoft Excel is referenced for workflow/compatibility context; Ledgerglass is an independent implementation and does not claim Microsoft affiliation.
