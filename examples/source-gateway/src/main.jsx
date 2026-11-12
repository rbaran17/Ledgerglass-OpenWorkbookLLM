import { createRoot } from "react-dom/client";
import {
  AuditWorkbook,
  createBlankWorkbook,
  createHttpSourceAdapter,
} from "@ledgerglass/workbook";
import "@ledgerglass/workbook/style.css";

// The adapter only names a source; the server decides who may see it and how it is read.
const sources = createHttpSourceAdapter({ endpoint: "/api/sources" });
const workbook = createBlankWorkbook("Receivables review");

createRoot(document.getElementById("root")).render(
  <div style={{ height: "100vh" }}>
    <AuditWorkbook initialWorkbook={workbook} externalSourceAdapter={sources} />
  </div>,
);
