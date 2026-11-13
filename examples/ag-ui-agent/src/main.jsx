import { createRoot } from "react-dom/client";
import {
  AuditWorkbook,
  createAgUiAdapter,
  createBlankWorkbook,
} from "@ledgerglass/workbook";
import "@ledgerglass/workbook/style.css";

// The assistant panel talks to an AG-UI agent. The agent sees only the selection the user
// approves, and what it proposes is shown for review before anything changes.
const assistant = createAgUiAdapter({ url: "/agent", name: "AG-UI agent" });

const workbook = createBlankWorkbook("Budget");
Object.assign(workbook.sheets[0].cells, {
  A1: { raw: "Item" },
  B1: { raw: "Cost" },
  A2: { raw: "Rent" },
  B2: { raw: 1200 },
  A3: { raw: "Food" },
  B3: { raw: 450 },
  A4: { raw: "Travel" },
  B4: { raw: 120 },
  A5: { raw: "Total" },
});

createRoot(document.getElementById("root")).render(
  <div style={{ height: "100vh" }}>
    <AuditWorkbook initialWorkbook={workbook} adapter={assistant} />
  </div>,
);
