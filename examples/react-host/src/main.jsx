// Entry point of the host page. The library's stylesheet comes from the package
// (`@ledgerglass/workbook/style.css`); the host's own layout lives in app.css.
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@ledgerglass/workbook/style.css";
import "./app.css";
import App from "./App.jsx";

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
