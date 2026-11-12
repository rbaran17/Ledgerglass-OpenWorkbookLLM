import { ChevronDown, PaintBucket } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { LocalizedText } from "./04-cell-input-paste-and-fill.js";
import { Wde } from "./24-merge-audit-and-ribbon-helpers.js";

function ColorPickerButton({
  commandId,
  reason,
  kind,
  value,
  onChange,
  disabled,
}) {
  const [isPaletteOpen, setPaletteOpen] = ne.useState(false),
    controlRef = ne.useRef(null),
    colorLabel = kind === "fill" ? "Fill Color" : "Font Color";
  return (
    ne.useEffect(() => {
      if (!isPaletteOpen) return;
      const handleDocumentPointerDown = (event) => {
        controlRef.current?.contains(event.target) || setPaletteOpen(false);
      };
      return (
        document.addEventListener("pointerdown", handleDocumentPointerDown),
        () =>
          document.removeEventListener("pointerdown", handleDocumentPointerDown)
      );
    }, [isPaletteOpen]),
    i.jsxs("div", {
      className: "xl-color-control",
      ref: controlRef,
      children: [
        i.jsxs("button", {
          "data-workbook-command": commandId,
          "aria-description": reason,
          type: "button",
          className: "xl-icon-button xl-color-button",
          disabled,
          "aria-label": colorLabel,
          title: reason ? `${colorLabel}: ${reason}` : colorLabel,
          "aria-expanded": isPaletteOpen,
          onClick: () => setPaletteOpen(!isPaletteOpen),
          children: [
            i.jsx("span", {
              className: "xl-color-glyph",
              style: { "--swatch-color": value },
              children:
                kind === "fill"
                  ? i.jsx(PaintBucket, { size: 15 })
                  : i.jsx("span", {
                      className: "xl-font-color-glyph",
                      children: i.jsx(LocalizedText, { children: "A" }),
                    }),
            }),
            i.jsx(ChevronDown, { size: 9 }),
          ],
        }),
        isPaletteOpen &&
          !disabled &&
          i.jsxs("div", {
            className: "xl-color-palette",
            role: "dialog",
            "aria-label": colorLabel,
            onKeyDown: (event) => {
              event.key === "Escape" &&
                (event.preventDefault(),
                event.stopPropagation(),
                setPaletteOpen(false),
                controlRef.current?.querySelector("button")?.focus());
            },
            children: [
              i.jsx("strong", {
                children: i.jsx(LocalizedText, { children: "Theme Colors" }),
              }),
              i.jsx("div", {
                className: "xl-color-grid",
                children: Wde.map((background, swatchIndex) =>
                  i.jsx(
                    "button",
                    {
                      type: "button",
                      "aria-label": `${colorLabel} ${background}`,
                      title: background,
                      style: { background: background },
                      onClick: () => {
                        (onChange(background), setPaletteOpen(false));
                      },
                    },
                    `${background}-${swatchIndex}`,
                  ),
                ),
              }),
              i.jsxs("label", {
                className: "xl-custom-color",
                children: [
                  i.jsx(LocalizedText, { children: "More Colors…" }),
                  i.jsx("input", {
                    type: "color",
                    "aria-label": `Custom ${colorLabel}`,
                    value,
                    onChange: (event) => onChange(event.target.value),
                  }),
                ],
              }),
            ],
          }),
      ],
    })
  );
}

export { ColorPickerButton };
