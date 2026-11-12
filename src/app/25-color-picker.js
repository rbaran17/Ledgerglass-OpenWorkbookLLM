import { ChevronDown, PaintBucket } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { M } from "./04-cell-input-paste-and-fill.js";
import { Wde } from "./24-merge-and-audit-analysis.js";

function Hk({
  commandId: e,
  reason: r,
  kind: t,
  value: n,
  onChange: o,
  disabled: a,
}) {
  const [s, d] = ne.useState(false),
    u = ne.useRef(null),
    l = t === "fill" ? "Fill Color" : "Font Color";
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
          "data-workbook-command": e,
          "aria-description": r,
          type: "button",
          className: "xl-icon-button xl-color-button",
          disabled: a,
          "aria-label": l,
          title: r ? `${l}: ${r}` : l,
          "aria-expanded": s,
          onClick: () => d(!s),
          children: [
            i.jsx("span", {
              className: "xl-color-glyph",
              style: { "--swatch-color": n },
              children:
                t === "fill"
                  ? i.jsx(PaintBucket, { size: 15 })
                  : i.jsx("span", {
                      className: "xl-font-color-glyph",
                      children: i.jsx(M, { children: "A" }),
                    }),
            }),
            i.jsx(ChevronDown, { size: 9 }),
          ],
        }),
        s &&
          !a &&
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
                children: i.jsx(M, { children: "Theme Colors" }),
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
                        (o(c), d(false));
                      },
                    },
                    `${c}-${h}`,
                  ),
                ),
              }),
              i.jsxs("label", {
                className: "xl-custom-color",
                children: [
                  i.jsx(M, { children: "More Colors…" }),
                  i.jsx("input", {
                    type: "color",
                    "aria-label": `Custom ${l}`,
                    value: n,
                    onChange: (c) => o(c.target.value),
                  }),
                ],
              }),
            ],
          }),
      ],
    })
  );
}

export { Hk };
