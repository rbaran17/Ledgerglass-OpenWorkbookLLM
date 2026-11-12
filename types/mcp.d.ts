/** Type declarations for `@ledgerglass/workbook/mcp`. */
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type {
  AssistantProposal,
  AssistantSelection,
  Workbook,
} from "./index.js";

export interface WorkbookMcpHost {
  /** The workbook this user is authorized to read. */
  getWorkbook(): Workbook;
  /** The selection they approved sharing. */
  getSelection(): AssistantSelection;
  /** Stores a validated proposal for a person to review. It must not apply it. */
  queueProposal(proposal: AssistantProposal): unknown;
  /** Runs before each tool call, for example to refresh the workbook. */
  beforeRequest?(): unknown;
}

/** Creates an MCP server with nine read-only tools and proposal queuing over one selection. */
export function createWorkbookMcpServer(host: WorkbookMcpHost): McpServer;
