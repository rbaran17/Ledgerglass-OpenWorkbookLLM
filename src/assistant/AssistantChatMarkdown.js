import Markdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { jsx, jsxs } from "react/jsx-runtime";

/**
 * Renders an assistant reply. Model output is untrusted: raw HTML is dropped,
 * images are replaced by a text placeholder, and only http(s) links are kept.
 */
export function AssistantChatMarkdown({ children }) {
  return jsx(Markdown, {
    remarkPlugins: [remarkGfm],
    skipHtml: true,
    components: {
      img: ({ alt }) =>
        jsxs("span", {
          children: ["[Image: ", alt || "remote image omitted", "]"],
        }),
      a: ({ href, children: label }) =>
        jsx("a", {
          href: /^https?:\/\//i.test(href ?? "") ? href : undefined,
          target: "_blank",
          rel: "noopener noreferrer",
          children: label,
        }),
    },
    children,
  });
}
