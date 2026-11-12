# Ledgerglass — Docker and developer integration

This repository contains the browser application (`src/`), the committed production build (`dist/`), portable HTML editions, the MCP server, deployment files and tests. Deployment scaffolding does not change the feature limits described in README.md.

## Run with Docker

From the repository root, with Docker Engine/Desktop and Compose v2 available:

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

There is no published or buildable React SDK in this repository. The application can be embedded as an iframe of the served app (`/?embedded`) if the host's Content-Security-Policy and the `FRAME_ANCESTORS` setting allow it; see [Embed in React](README.md#embed-in-react). The static container serves the application only. It does not implement backend routes for the assistant or the source gateway, so a host integration needs its own backend. For the MCP server, see [docs/MCP.md](docs/MCP.md).

## Files to give your developer

| Entry | Purpose |
| --- | --- |
| `Dockerfile`, `.dockerignore`, `compose.yaml` | Rebuildable web and isolated stdio MCP images, local load balancing |
| `deploy/kubernetes/` | Deployment, Service, HPA, disruption budget and optional ingress template |
| `deploy/server.mjs`, `deploy/tests/` | Dependency-free HTTP runtime and executed runtime checks |
| `deploy/verify-containers.mjs` | Docker build, three-replica, scale-down and MCP startup verification |
| `.github/workflows/ci.yml` | Runs the runtime tests, a dependency audit, the unit tests and browser checks, and builds the web and MCP images; does not publish or deploy |
| `src/`, `tests/`, `scripts/` | Current implementation, workers, regression tests and build/run scripts |
| `docs/MCP.md` | MCP tools, configuration and integration |

**Verification boundary:** Docker, kubectl and a cluster were unavailable when these checks were last run, so container and cluster behavior was not verified here. Runtime HTTP tests and builds run locally; actual image builds, Compose behavior, Kubernetes scaling, production identity/storage and external-provider calls require the receiving environment. Run `npm ci`, `npm run test:deployment` and `npm run test:containers` there.
