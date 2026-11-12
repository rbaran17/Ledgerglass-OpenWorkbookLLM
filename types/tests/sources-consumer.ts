/**
 * Uses the source gateway, the connectors and the browser adapter the way a host would. Checked by
 * `npm run test:types`; nothing here runs.
 */
import {
  createServer,
  type IncomingMessage,
  type ServerResponse,
} from "node:http";
import {
  createSourceGateway,
  createSourceGatewayHttpHandler,
  type SourceApproval,
  type SourceConnector,
  type SourceGateway,
} from "@ledgerglass/workbook/source-gateway";
import {
  createDelimitedFileConnector,
  createJsonApiConnector,
  createODataConnector,
  createQueryConnector,
  createSqliteConnector,
} from "@ledgerglass/workbook/source-connectors";
import {
  ExternalSourceError,
  createHttpSourceAdapter,
  type AuditWorkbookProps,
  type ExternalSourceAdapter,
} from "@ledgerglass/workbook";

interface Caller {
  name: string;
  roles: string[];
}

const invoices = createDelimitedFileConnector({
  id: "invoices",
  label: "Invoices",
  directory: "/srv/exports",
  file: "invoices.csv",
  columns: [
    { name: "Invoice", type: "text" },
    { name: "Issued", type: "date" },
    { name: "Amount", type: "number", nullable: true },
  ],
  controlFields: ["Amount"],
});

const orders = createJsonApiConnector({
  id: "orders",
  label: "Orders",
  url: "https://api.example.com/orders",
  headers: { "X-Api-Key": "placeholder" },
  rowsPath: "data.items",
  fields: ["order", "qty"],
  columns: [
    { name: "Order", type: "text" },
    { name: "Qty", type: "number" },
  ],
});

const openInvoices = createODataConnector({
  id: "open-invoices",
  label: "Open invoices",
  url: "https://erp.example.com/odata/Invoices",
  filter: "Status eq 'open'",
  columns: [
    { name: "Number", type: "text" },
    { name: "Total", type: "number" },
  ],
});

const ledger = createSqliteConnector({
  id: "ledger",
  label: "Ledger balances",
  database: "/srv/ledger.db",
  query: "SELECT account, balance FROM balances WHERE balance <> ?",
  parameters: [0],
  columns: [
    { name: "Account", type: "text" },
    { name: "Balance", type: "number" },
  ],
  controlFields: ["Balance"],
});

const stock = createQueryConnector({
  id: "erp-stock",
  label: "Stock on hand",
  kind: "erp",
  columns: [
    { name: "Item", type: "text" },
    { name: "OnHand", type: "number" },
  ],
  async run({ signal }) {
    if (signal?.aborted) return [];
    return [["bolt", 12]];
  },
});

const connectors: SourceConnector[] = [invoices, orders, openInvoices, stock];

const gateway: SourceGateway<Caller> = createSourceGateway<Caller>({
  connectors,
  isAllowed: (sourceId, caller) =>
    caller.roles.includes("finance") && sourceId !== "erp-stock",
  approvalTtlMs: 10 * 60 * 1000,
});

function callerForToken(token: string | undefined): Caller | null {
  return token === undefined ? null : { name: "alice", roles: ["finance"] };
}

const handle = createSourceGatewayHttpHandler(gateway, {
  authenticate: (request: IncomingMessage) =>
    callerForToken(request.headers.authorization),
  path: "/api/sources",
  timeoutMs: 30_000,
});

/** Serves the gateway on one port and answers everything else with 404. */
export function startSourceServer(port: number) {
  const server = createServer(
    async (request: IncomingMessage, response: ServerResponse) => {
      if (await handle(request, response)) return;
      response.statusCode = 404;
      response.end();
    },
  );
  server.listen(port);
  return server;
}

/** Reads a source the way a host would, and reports what it found. */
export async function describeSource(caller: Caller): Promise<string> {
  const approvals: SourceApproval[] = await gateway.catalog({
    principal: caller,
  });
  try {
    const payload = await gateway.refresh("invoices", {
      principal: caller,
      signal: AbortSignal.timeout(30_000),
    });
    return `${payload.source.label}: ${payload.rows.length} rows, ${approvals.length} sources`;
  } catch (error) {
    if (error instanceof ExternalSourceError) return `refused: ${error.code}`;
    throw error;
  }
}

/** The browser side: the adapter is passed as the editor's prop. */
export function editorProps(): AuditWorkbookProps {
  const adapter: ExternalSourceAdapter = createHttpSourceAdapter({
    endpoint: "/api/sources",
    credentials: "same-origin",
  });
  return { externalSourceAdapter: adapter };
}

// @ts-expect-error isAllowed is required: there is no default that allows access
createSourceGateway({ connectors });

createDelimitedFileConnector({
  id: "prices",
  label: "Prices",
  directory: "/srv/exports",
  file: "prices.csv",
  // @ts-expect-error "money" is not a column type
  columns: [{ name: "Price", type: "money" }],
});
