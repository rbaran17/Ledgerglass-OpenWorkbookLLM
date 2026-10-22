# Ledgerglass — Docker and developer integration

This package contains the complete current workbook source, compiled React SDK, browser application, portable HTML editions, tests, requirements, validation evidence and examples. No implemented workbook feature was removed. Deployment scaffolding does not change the published feature limits or coverage statuses.

## Run with Docker

From this `auditor-workbook` directory, with Docker Engine/Desktop and Compose v2 available:

```sh
docker compose up --build -d --scale web=2 --wait
```

Open **http://localhost:8080/?embedded**. A gateway balances requests across two stateless web replicas. Only the gateway publishes a localhost port. To change capacity:

```sh
docker compose up -d --scale web=3 --wait
docker compose down
```

Use `LEDGERGLASS_PORT=8081 docker compose up --build -d --wait` if port 8080 is occupied. The first build downloads official images and locked npm dependencies. No image registry account, LLM key, database or production service is needed for the local demo. The packaged browser demo uses local browser storage and a deterministic assistant; Docker does not turn it into shared server storage or a connected LLM.

## Integrate into your application

From your React 19 project, install the tarball included in the kit's sibling `packages/` directory:

```sh
npm install /absolute/path/to/packages/ledgerglass-workbook-0.1.0.tgz
```

Use [deploy/ReactHost.example.tsx](deploy/ReactHost.example.tsx) as the component boundary. The parent application provides the authorized workbook, storage adapter, user identity and same-origin backend routes. React integration uses your own application deployment; the standalone web container is optional.

The example expects `/api/workbook-assistant` and `/api/workbook-sources` in **your** backend. The static container does not implement or proxy these routes. Use the existing [assistant protocol](docs/CHAT-UI-INTEGRATION.md), [source gateway](docs/EXTERNAL-SOURCES-INTEGRATION.md), [source connectors](docs/SOURCE-CONNECTOR-INTEGRATION.md), [collaboration contract](docs/COLLABORATION-INTEGRATION.md) and [MCP integration](docs/MCP.md). Those guides describe the actual supported contracts and acceptance boundaries.

## Files to give your developer

| Entry | Purpose |
| --- | --- |
| `Dockerfile`, `.dockerignore`, `compose.yaml` | Rebuildable web and isolated stdio MCP images, local load balancing |
| `deploy/kubernetes/` | Deployment, Service, HPA, disruption budget and optional ingress template |
| `deploy/server.mjs`, `deploy/tests/` | Dependency-free HTTP runtime and executed runtime checks |
| `deploy/verify-containers.mjs` | Docker build, three-replica, scale-down and MCP startup verification |
| `.github/workflows/containers.yml` | Optional CI example; builds/tests, does not publish or deploy |
| `src/`, `lib/`, `examples/`, `tests/`, `scripts/` | Current implementation, SDK, workers, integration examples and regression tests |
| `docs/BASE-EXCEL-COVERAGE.xlsx` | Current base Excel coverage, including remaining gaps |
| `docs/DOCKER-AND-SCALING.md` | Deployment design, configuration, operations and developer acceptance |
| `docs/DOCKER-PACKAGING-VERIFICATION.json` | Executed checks and explicit unexecuted Docker/cluster checks |

The kit root also contains `HANDOFF-MANIFEST.json`, `SHA256SUMS.txt`, the npm package and the packaging validation report. Verify these hashes after transfer.

**Verification boundary:** Docker, kubectl and a cluster were unavailable on the preparation machine. Runtime HTTP tests, builds and package checks are performed locally; actual image builds, Compose behavior, Kubernetes scaling, production identity/storage and external-provider calls require the receiving environment. Run `npm ci`, `npm run test:deployment` and `npm run test:containers` there. See the executed record for exact results.
