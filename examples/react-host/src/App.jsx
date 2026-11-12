// A product page that embeds the Ledgerglass editor: a toolbar for the document (download,
// open, reset) above a sized container that holds <AuditWorkbook />.
import { useCallback, useRef, useState } from "react";
import {
  AuditWorkbook,
  applyCellPatchTransaction,
  createBlankWorkbook,
  createHttpAdapter,
  exportWorkbookToXlsx,
  importDelimitedFile,
  importWorkbookFile,
  renameWorksheet,
} from "@ledgerglass/workbook";

// The assistant asks this page's own server. The server builds the proposal and the editor
// validates it against the approved selection before showing it. No model key lives here.
const assistant = createHttpAdapter({ endpoint: "/api/propose" });

// The starting document. Each call returns a workbook with a new id, so passing a new one to the
// editor switches documents (that is how Reset and Open file work).
function createStartingWorkbook() {
  const blank = createBlankWorkbook("Household budget");
  const sheetId = blank.sheets[0].id;
  const workbook = renameWorksheet(blank, sheetId, "Budget");
  const cells = {
    A1: "Item",
    B1: "Amount",
    A2: "Rent",
    B2: 1200,
    A3: "Food",
    B3: 450,
    A4: "Transport",
    B4: 120,
    D1: "Summary",
    D2: "Total of amounts",
    E2: "=SUM(B2:B4)",
  };
  const patches = Object.entries(cells).map(([address, raw]) => ({
    sheetId,
    address,
    cell: { raw },
  }));
  return applyCellPatchTransaction(
    workbook,
    patches,
    "Host",
    "Load starting budget",
  );
}

export default function App() {
  // `opened` is the document the editor starts with. The editor keeps its own copy after that.
  const [opened, setOpened] = useState(() => createStartingWorkbook());
  const [current, setCurrent] = useState({
    name: opened.name,
    revision: opened.revision,
  });
  const [status, setStatus] = useState("");
  // Download reads the latest workbook from here, so it never waits for a re-render.
  const latest = useRef(opened);

  const remember = useCallback((workbook) => {
    latest.current = workbook;
    setCurrent({ name: workbook.name, revision: workbook.revision });
  }, []);

  const showDocument = useCallback(
    (workbook) => {
      remember(workbook);
      setOpened(workbook);
    },
    [remember],
  );

  async function download() {
    try {
      const { blob, warnings } = await exportWorkbookToXlsx(latest.current);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `${latest.current.name}.xlsx`;
      link.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      setStatus(`Downloaded XLSX with ${warnings.length} export notes.`);
    } catch (error) {
      setStatus(error.message);
    }
  }

  async function openFile(event) {
    const file = event.target.files?.[0];
    event.target.value = ""; // so picking the same file again still fires onChange
    if (!file) return;
    try {
      // CSV goes through the delimited importer; anything else is read as XLSX.
      const isCsv = /\.csv$/i.test(file.name);
      const result = isCsv
        ? await importDelimitedFile(file)
        : await importWorkbookFile(file);
      const warnings = result.warnings ?? [];
      showDocument(result.workbook);
      setStatus(`Opened ${file.name} with ${warnings.length} warnings.`);
    } catch (error) {
      // File errors are written for the person who chose the file, so show them as they are.
      setStatus(error.message);
    }
  }

  function reset() {
    showDocument(createStartingWorkbook());
    setStatus("Reset to the starting budget.");
  }

  return (
    <div className="page">
      <header className="toolbar">
        <div className="toolbar-title">
          <strong data-testid="workbook-name">{current.name}</strong>
          <span data-testid="revision">Revision {current.revision}</span>
        </div>
        <div className="toolbar-actions">
          <button type="button" onClick={download}>
            Download XLSX
          </button>
          <label className="file-button">
            Open file
            <input
              type="file"
              accept=".xlsx,.csv"
              onChange={openFile}
            />
          </label>
          <button type="button" onClick={reset}>
            Reset
          </button>
        </div>
        <p role="status" className="status">
          {status}
        </p>
      </header>
      <main className="editor">
        <AuditWorkbook
          initialWorkbook={opened}
          adapter={assistant}
          onChange={remember}
          onWorkbookLoaded={remember}
        />
      </main>
    </div>
  );
}
