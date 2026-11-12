import i from "react/jsx-runtime";
import { LocalizedText } from "../04-cell-input-paste-and-fill.js";
import { formatNumberWithPattern } from "../03-workbook-model.js";

function renderStatusBar({
  dialogState,
  selectionState,
  operationStatusState,
  viewDisplayState,
  text,
  language,
  isEditingLocked,
  selectionAreas,
  selectedCellCount,
  selectedNumericValues,
  selectedNumericSum,
  applyZoomPercent,
}) {
  return i.jsxs("footer", {
    role: "group",
    className: "lg-grid-footer",
    "aria-label": text("Selection summary"),
    lang: language("Selection summary"),
    "data-ui-label": "Selection summary",
    children: [
      i.jsxs("span", {
        children: [
          i.jsx("span", { className: "lg-ready-dot" }),
          i.jsx(LocalizedText, {
            children: operationStatusState.isOperationBusy
              ? "Working…"
              : "Ready",
          }),
          i.jsx("button", {
            type: "button",
            onClick: () => dialogState.setIsWorkbookHealthDialogOpen(true),
            children: i.jsx(LocalizedText, {
              children: "Workbook health",
            }),
          }),
          operationStatusState.isOperationBusy &&
            operationStatusState.activeOperationControllerRef.current &&
            i.jsx("button", {
              onClick: () =>
                operationStatusState.activeOperationControllerRef.current?.abort(),
              children: i.jsx(LocalizedText, {
                children: "Cancel file processing",
              }),
            }),
          i.jsx("span", { className: "lg-footer-separator" }),
          i.jsxs("span", {
            title: text(
              "Selection count, numeric count, sum and average include hidden and filtered cells within the selected range.",
            ),
            lang: language(
              "Selection count, numeric count, sum and average include hidden and filtered cells within the selected range.",
            ),
            children: [
              i.jsx(LocalizedText, { children: "Count: " }),
              selectedCellCount,
              " ",
              selectionState.completedSelectionAreas.length > 0
                ? `· ${selectionAreas.length} areas `
                : "",
              " ",
              i.jsx(LocalizedText, { children: "· Numbers: " }),
              selectedNumericValues.length,
              i.jsxs("span", {
                className: "lg-selection-inclusion",
                children: [
                  " ",
                  i.jsx(LocalizedText, {
                    children: "· Includes hidden and filtered cells",
                  }),
                ],
              }),
            ],
          }),
          isEditingLocked ? " · Read only" : "",
        ],
      }),
      i.jsxs("div", {
        children: [
          i.jsxs("span", {
            children: [
              i.jsx(LocalizedText, { children: "Average" }),
              " ",
              i.jsx("b", {
                children: formatNumberWithPattern(
                  selectedNumericValues.length
                    ? selectedNumericSum / selectedNumericValues.length
                    : 0,
                ),
              }),
            ],
          }),
          i.jsxs("span", {
            children: [
              i.jsx(LocalizedText, { children: "Sum" }),
              " ",
              i.jsx("b", {
                children: formatNumberWithPattern(selectedNumericSum, {
                  raw: 0,
                  format: "number",
                }),
              }),
            ],
          }),
        ],
      }),
      i.jsxs("div", {
        className: "lg-zoom",
        children: [
          i.jsx("button", {
            "aria-label": text("Zoom out"),
            onClick: () =>
              applyZoomPercent(Math.max(50, viewDisplayState.zoomPercent - 10)),
            lang: language("Zoom out"),
            "data-ui-label": "Zoom out",
            children: "−",
          }),
          i.jsx("input", {
            "aria-label": text("Zoom"),
            type: "range",
            min: "50",
            max: "200",
            step: "10",
            value: viewDisplayState.zoomPercent,
            onChange: (fe) => applyZoomPercent(Number(fe.target.value)),
            lang: language("Zoom"),
            "data-ui-label": "Zoom",
          }),
          i.jsx("button", {
            "aria-label": text("Zoom in"),
            onClick: () =>
              applyZoomPercent(
                Math.min(200, viewDisplayState.zoomPercent + 10),
              ),
            lang: language("Zoom in"),
            "data-ui-label": "Zoom in",
            children: "+",
          }),
          i.jsxs("span", { children: [viewDisplayState.zoomPercent, "%"] }),
        ],
      }),
    ],
  });
}

export { renderStatusBar };
