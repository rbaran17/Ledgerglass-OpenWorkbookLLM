import { ChevronDown, PaintBucket } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { LocalizedText } from "./04-cell-input-paste-and-fill.js";
import { Wde } from "./24-merge-and-audit-analysis.js";

function ColorPickerButton({
  commandId,
  reason,
  kind,
  value,
  onChange,
  disabled,
}) {
  const [s, d] = ne.useState(false),
    u = ne.useRef(null),
    l = kind === "fill" ? "Fill Color" : "Font Color";
  return (
    ne.useEffect(() => {
      if (!s) return;
      const c = (h) => {
        var f;
        ((f = u.current) != null && f.contains(h.target)) || d(false);
      };
      return (
        document.addEventListener("pointerdown", c),
        () => document.removeEventListener("pointerdown", c)
      );
    }, [s]),
    i.jsxs("div", {
      className: "xl-color-control",
      ref: u,
      children: [
        i.jsxs("button", {
          "data-workbook-command": commandId,
          "aria-description": reason,
          type: "button",
          className: "xl-icon-button xl-color-button",
          disabled,
          "aria-label": l,
          title: reason ? `${l}: ${reason}` : l,
          "aria-expanded": s,
          onClick: () => d(!s),
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
        s &&
          !disabled &&
          i.jsxs("div", {
            className: "xl-color-palette",
            role: "dialog",
            "aria-label": l,
            onKeyDown: (c) => {
              var f;
              c.key === "Escape" &&
                (c.preventDefault(),
                c.stopPropagation(),
                d(false),
                (f = u.current?.querySelector("button")) == null || f.focus());
            },
            children: [
              i.jsx("strong", {
                children: i.jsx(LocalizedText, { children: "Theme Colors" }),
              }),
              i.jsx("div", {
                className: "xl-color-grid",
                children: Wde.map((c, h) =>
                  i.jsx(
                    "button",
                    {
                      type: "button",
                      "aria-label": `${l} ${c}`,
                      title: c,
                      style: { background: c },
                      onClick: () => {
                        (onChange(c), d(false));
                      },
                    },
                    `${c}-${h}`,
                  ),
                ),
              }),
              i.jsxs("label", {
                className: "xl-custom-color",
                children: [
                  i.jsx(LocalizedText, { children: "More Colors…" }),
                  i.jsx("input", {
                    type: "color",
                    "aria-label": `Custom ${l}`,
                    value,
                    onChange: (c) => onChange(c.target.value),
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
