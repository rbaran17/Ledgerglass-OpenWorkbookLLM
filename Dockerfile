# syntax=docker/dockerfile:1
# Pin approved image digests through --build-arg NODE_IMAGE=... in release CI.
ARG NODE_IMAGE=node:22-bookworm-slim
FROM ${NODE_IMAGE} AS dependencies
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --no-audit --no-fund
COPY . .

FROM dependencies AS web-build
RUN npm run build:container

FROM dependencies AS sdk-build
RUN npm run build:lib

FROM ${NODE_IMAGE} AS production-dependencies
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev --no-audit --no-fund && npm cache clean --force

# Local, per-workbook stdio MCP. Not an unauthenticated network MCP endpoint.
FROM ${NODE_IMAGE} AS mcp
ENV NODE_ENV=production
WORKDIR /app
COPY --from=production-dependencies --chown=node:node /app/node_modules ./node_modules
COPY --from=sdk-build --chown=node:node /app/lib ./lib
COPY --chown=node:node package.json LICENSE THIRD-PARTY-NOTICES.txt ./
COPY --chown=node:node scripts/mcp-stdio.mjs ./scripts/mcp-stdio.mjs
USER node
ENTRYPOINT ["node", "scripts/mcp-stdio.mjs"]
CMD ["--help"]

# Default image: no npm dependencies, source code, credentials or workbook files.
FROM ${NODE_IMAGE} AS web
ENV NODE_ENV=production HOST=0.0.0.0 PORT=8080 STATIC_ROOT=/app/public
WORKDIR /app
COPY --from=web-build --chown=node:node /app/dist-container ./public
COPY --chown=node:node deploy/server.mjs deploy/healthcheck.mjs ./deploy/
COPY --chown=node:node LICENSE THIRD-PARTY-NOTICES.txt ./
USER node
EXPOSE 8080
HEALTHCHECK --interval=15s --timeout=3s --start-period=10s --retries=3 CMD ["node", "deploy/healthcheck.mjs"]
STOPSIGNAL SIGTERM
CMD ["node", "deploy/server.mjs"]
