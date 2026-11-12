import {
  useExternalStoreRuntime,
  AssistantRuntimeProvider,
  ThreadPrimitive,
  MessagePrimitive,
  ComposerPrimitive,
} from "@assistant-ui/react";
import { LoaderCircle, Paperclip } from "lucide-react";
import "../styles/assistant.css";
import {
  u,
  r,
  j,
  P,
  W,
  T,
  S,
  A,
  C,
  d,
  e,
  f,
  h,
  i,
  l,
  p,
  v,
  k,
  F,
  X,
} from "../main.js";
import { AssistantChatMarkdown } from "./AssistantChatMarkdown.js";

const db = (t) =>
  t < 1024 * 1024
    ? `${Math.max(1, Math.ceil(t / 1024))} KB`
    : `${(t / 1024 / 1024).toFixed(1)} MB`;
function ai({ file, remove }) {
  const { text, language, message } = u();
  return j.jsxs("div", {
    className: "lg-chat-file",
    children: [
      j.jsxs("details", {
        children: [
          j.jsxs("summary", {
            children: [
              j.jsx(F, { size: 17, "aria-hidden": "true" }),
              j.jsxs("span", {
                children: [
                  j.jsx("strong", { children: file.name }),
                  j.jsxs("small", {
                    children: [
                      db(file.size),
                      " ",
                      j.jsx(W, { children: "·" }),
                      " ",
                      file.encoding === "utf8" ||
                      file.mediaType.startsWith("image/")
                        ? text("Preview")
                        : text("File ready"),
                    ],
                  }),
                ],
              }),
            ],
          }),
          file.encoding === "utf8"
            ? j.jsxs("pre", {
                className: "lg-chat-file-preview",
                children: [
                  file.data.slice(0, 4e3),
                  file.data.length > 4e3
                    ? `
[Preview limited to 4,000 characters; the full file is attached.]`
                    : "",
                ],
              })
            : file.mediaType.startsWith("image/")
              ? j.jsx("img", {
                  src: `data:${file.mediaType};base64,${file.data}`,
                  alt: message("Preview of {name}", { name: file.name }),
                })
              : j.jsx("p", {
                  className: "lg-chat-file-description",
                  children: j.jsx(W, {
                    children:
                      "The original file will accompany your message. Preview is unavailable for this format.",
                  }),
                }),
        ],
      }),
      remove &&
        j.jsx("button", {
          type: "button",
          "aria-label": message("Remove attachment {name}", {
            name: file.name,
          }),
          title: text("Remove attachment"),
          onClick: remove,
          children: j.jsx(X, { size: 14 }),
        }),
    ],
  });
}
function mb(t) {
  const { text, language, message } = u(),
    r$1 = r.useRef(t);
  r$1.current = t;
  const i$1 = r.useRef(false),
    o = r.useRef(null),
    a = r.useRef(null),
    c = r.useRef(null),
    [l$1, u$1] = r.useState(false),
    [h$1, d$1] = r.useState(""),
    [f$1, p$1] = r.useState(false),
    [m, _] = r.useState(null),
    [x, y] = r.useState("");
  (r.useEffect(() => () => c.current?.abort(), []),
    r.useEffect(() => {
      (c.current?.abort(), u$1(false), d$1(""));
    }, [t.session.id, t.adapter]));
  async function w(v) {
    if (!v.length || !t.adapter.chat?.attachments || c.current || l$1 || t.busy)
      return;
    const D = t.session.attachments;
    if ((d$1(""), D.length + v.length > l.filesPerMessage)) {
      d$1(
        "Attach up to 4 files per message. Remove a file before adding more.",
      );
      return;
    }
    if (
      D.reduce((F, H) => F + H.size, 0) + v.reduce((F, H) => F + H.size, 0) >
      l.totalFileBytes
    ) {
      d$1("Attachments must total 10 MB or less.");
      return;
    }
    if (
      new Set([
        ...D.map((F) => `${F.name}:${F.size}`),
        ...v.map((F) => `${F.name}:${F.size}`),
      ]).size !==
      D.length + v.length
    ) {
      d$1("This file is already attached. Remove it before adding it again.");
      return;
    }
    const $ = new AbortController(),
      U = t.session.id;
    ((c.current = $), u$1(true));
    try {
      const F = await Promise.all(v.map((H) => p(H, $.signal)));
      $.signal.aborted ||
        t.setSession((H) =>
          H.id === U ? { ...H, attachments: [...H.attachments, ...F] } : H,
        );
    } catch (F) {
      $.signal.aborted ||
        d$1(
          F instanceof Error ? F.message : "The attachment could not be read.",
        );
    } finally {
      c.current === $ && (u$1(false), (c.current = null));
    }
  }
  async function S$1(v$1, D) {
    const $ = r$1.current;
    if (i$1.current || $.busy || l$1 || !v$1.trim()) return;
    const U = D ?? $.session.attachments,
      K = $.scope,
      F = $.session.id,
      H = [];
    for (let N = 1; N < $.session.messages.length; N++) {
      const V = $.session.messages[N],
        L = $.session.messages[N - 1];
      V.role === "assistant" &&
        V.outcome === "complete" &&
        V.scope === K &&
        L.role === "user" &&
        L.scope === K &&
        H.push(
          {
            role: "user",
            text: L.text,
            ...(L.attachments?.length ? { attachments: L.attachments } : {}),
          },
          { role: "assistant", text: V.text },
        );
    }
    let J;
    try {
      if ($.session.messages.length >= 60)
        throw Error(
          "Start a new chat to continue. This chat has reached its message limit.",
        );
      if (v$1.length > l.promptCharacters)
        throw Error("Keep your message under 10,000 characters.");
      if (U.length && !$.adapter.chat?.attachments)
        throw Error(
          "This provider does not support file attachments. Remove the files or connect a chat-enabled provider.",
        );
      $.adapter.chat &&
        (J = v({
          version: 1,
          messages: $.adapter.chat.history ? H : [],
          attachments: U,
        }));
    } catch (N) {
      d$1(N.message);
      return;
    }
    const X = {
      id: crypto.randomUUID(),
      role: "user",
      text: v$1.trim(),
      scope: K,
      label: $.selectionLabel,
      attachments: U,
    };
    ((i$1.current = true),
      d$1(""),
      y(""),
      $.setSession((N) => ({
        ...N,
        draft: "",
        attachments: [],
        messages: [...N.messages, X],
      })));
    try {
      const N = await $.onSend(X.text, J);
      $.setSession((V) =>
        V.id !== F
          ? V
          : {
              ...V,
              messages: [
                ...V.messages,
                {
                  id: crypto.randomUUID(),
                  role: "assistant",
                  text:
                    N?.explanation ?? "Request stopped. No changes were made.",
                  scope: K,
                  label: X.label,
                  proposalId: N?.id,
                  outcome: N ? "complete" : "stopped",
                },
              ],
            },
      );
    } catch {
      $.setSession((N) =>
        N.id !== F
          ? N
          : {
              ...N,
              draft: N.draft || X.text,
              attachments: N.attachments.length ? N.attachments : U,
              messages: [
                ...N.messages,
                {
                  id: crypto.randomUUID(),
                  role: "assistant",
                  text: "The request did not complete. Your message is ready to retry.",
                  scope: K,
                  label: X.label,
                  outcome: "failed",
                },
              ],
            },
      );
    } finally {
      i$1.current = false;
    }
  }
  const I = r.useMemo(
      () =>
        t.session.messages.map((v) => ({
          id: v.id,
          role: v.role,
          content:
            v.scope === t.scope
              ? v.text
              : text(
                  "This message belongs to an earlier selection or workbook revision and is excluded from the current context.",
                ),
          metadata: { custom: { source: v } },
        })),
      [t.session.messages, t.scope],
    ),
    R = useExternalStoreRuntime({
      messages: I,
      convertMessage: (v) => v,
      isRunning: t.busy,
      onNew: async (v) => {
        await S$1(
          v.content.filter((D) => D.type === "text").map((D) => D.text).join(`
`),
        );
      },
      onCancel: async () => t.onCancel(),
    }),
    k$1 = t.session.messages.some(
      (v) =>
        v.proposalId && v.proposalId === t.proposalId && v.scope === t.scope,
    ),
    M = () => {
      const v = [...t.session.messages]
        .reverse()
        .find((D) => D.role === "user" && D.scope === t.scope);
      v && S$1(v.text, v.attachments);
    },
    P$1 = () => {
      (t.onCancel(),
        t.onReset(),
        c.current?.abort(),
        t.setSession(k()),
        y("New chat started."),
        a.current?.focus());
    },
    E = async (v) => {
      try {
        (await navigator.clipboard.writeText(v.text),
          _(v.id),
          y("Message copied."));
      } catch {
        y(
          "Copy is unavailable. Select the message text and use your browser’s Copy command.",
        );
      }
    };
  return j.jsx(AssistantRuntimeProvider, {
    runtime: R,
    children: j.jsxs(ThreadPrimitive.Root, {
      className: "lg-chat",
      role: "group",
      "aria-label": text("Workbook chat"),
      children: [
        j.jsxs("div", {
          className: "lg-chat-toolbar",
          children: [
            j.jsxs("span", {
              className: "lg-chat-provider",
              title: t.adapter.name,
              children: [
                j.jsx("i", {}),
                t.local ? text("Local demo") : t.adapter.name,
              ],
            }),
            j.jsxs("button", {
              type: "button",
              onClick: P$1,
              "aria-label": text("New chat"),
              title: text("New chat"),
              children: [
                j.jsx(P, { size: 16 }),
                j.jsx("span", {
                  children: j.jsx(W, { children: "New chat" }),
                }),
              ],
            }),
          ],
        }),
        j.jsxs("div", {
          className: "lg-chat-context",
          role: "group",
          "aria-label": text("Shared workbook context"),
          children: [
            j.jsx(T, { size: 16 }),
            j.jsxs("span", {
              children: [
                j.jsx("strong", { children: t.selectionLabel }),
                j.jsx("small", {
                  children: message(
                    "{count} selected cells · shared when you send",
                    {
                      count: t.selectedCount,
                    },
                  ),
                }),
              ],
            }),
          ],
        }),
        j.jsxs(ThreadPrimitive.Viewport, {
          className: "lg-chat-messages",
          role: "log",
          "aria-label": text("Conversation"),
          "aria-live": "polite",
          "aria-relevant": "additions text",
          tabIndex: 0,
          autoScroll: t.session.messages.length > 0,
          scrollToBottomOnInitialize: t.session.messages.length > 0,
          scrollToBottomOnThreadSwitch: false,
          children: [
            !t.session.messages.length &&
              !t.proposalId &&
              j.jsxs("div", {
                className: "lg-chat-welcome",
                children: [
                  j.jsx("div", {
                    className: "lg-chat-welcome-icon",
                    children: j.jsx(S, { size: 24 }),
                  }),
                  j.jsx("h3", {
                    children: j.jsx(W, { children: "How can I help?" }),
                  }),
                  j.jsx("p", {
                    children: j.jsx(W, {
                      children:
                        "Ask about your selection, attach a file, or work through a formula together.",
                    }),
                  }),
                  j.jsx("div", {
                    className: "lg-chat-suggestions",
                    children: [
                      "Explain this selection",
                      "Highlight values above 25000",
                      "Add a total formula",
                    ].map((v) =>
                      j.jsxs(
                        "button",
                        {
                          type: "button",
                          disabled: t.busy || l$1,
                          onClick: () => {
                            (t.setSession((D) => ({ ...D, draft: v })), S$1(v));
                          },
                          children: [
                            j.jsx(W, { children: v }),
                            j.jsx(A, { size: 13 }),
                          ],
                        },
                        v,
                      ),
                    ),
                  }),
                ],
              }),
            j.jsx(ThreadPrimitive.Messages, {
              children: ({ message }) => {
                const D = message.metadata.custom.source;
                if (!D) return null;
                const $ = D.scope === t.scope,
                  U = $ && D.proposalId === t.proposalId && !!t.review;
                return j.jsxs(MessagePrimitive.Root, {
                  role: "article",
                  className: `lg-chat-message ${D.role}`,
                  "aria-label":
                    D.role === "user"
                      ? text("Your message")
                      : text("Assistant message"),
                  children: [
                    j.jsxs("div", {
                      className: "lg-chat-message-label",
                      children: [
                        D.role === "user"
                          ? "You"
                          : j.jsxs(j.Fragment, {
                              children: [
                                j.jsx(S, { size: 13 }),
                                j.jsx(W, { children: "Assistant" }),
                              ],
                            }),
                        j.jsx("small", { children: D.label }),
                      ],
                    }),
                    j.jsxs("div", {
                      className: `lg-chat-bubble ${$ ? "" : "archived"}`,
                      children: [
                        $ &&
                          D.attachments?.map((F) =>
                            j.jsx(ai, { file: F }, F.id),
                          ),
                        U
                          ? t.review
                          : j.jsx(AssistantChatMarkdown, {
                              children: $
                                ? D.text
                                : text(
                                    "This message belongs to an earlier selection or workbook revision and is excluded from the current context.",
                                  ),
                            }),
                      ],
                    }),
                    $ &&
                      D.outcome === "complete" &&
                      j.jsxs("div", {
                        className: "lg-chat-message-actions",
                        children: [
                          j.jsx("button", {
                            type: "button",
                            "aria-label": text("Copy assistant response"),
                            title: text("Copy response"),
                            onClick: () => void E(D),
                            children:
                              m === D.id
                                ? j.jsx(C, { size: 14 })
                                : j.jsx(d, { size: 14 }),
                          }),
                          D === t.session.messages.at(-1) &&
                            j.jsx("button", {
                              type: "button",
                              "aria-label": text("Regenerate response"),
                              title: text("Regenerate response"),
                              disabled: t.busy || l$1,
                              onClick: M,
                              children: j.jsx(e, { size: 14 }),
                            }),
                        ],
                      }),
                    $ &&
                      (D.outcome === "failed" || D.outcome === "stopped") &&
                      D === t.session.messages.at(-1) &&
                      j.jsxs("button", {
                        type: "button",
                        className: "lg-chat-retry",
                        disabled: t.busy || l$1,
                        onClick: M,
                        children: [
                          j.jsx(e, { size: 13 }),
                          j.jsx(W, { children: "Retry message" }),
                        ],
                      }),
                  ],
                });
              },
            }),
            !k$1 &&
              t.review &&
              j.jsx("div", {
                className: "lg-chat-external-review",
                children: t.review,
              }),
            t.busy &&
              j.jsxs("div", {
                role: "status",
                className: "lg-chat-thinking",
                children: [
                  j.jsx(LoaderCircle, { size: 15, className: "lg-spin" }),
                  j.jsx(W, { children: "Preparing a response…" }),
                ],
              }),
          ],
        }),
        j.jsxs("div", {
          className: "lg-chat-composer-area",
          children: [
            j.jsx(ThreadPrimitive.ScrollToBottom, {
              className: "lg-chat-scroll",
              "aria-label": text("Scroll to latest message"),
              title: text("Scroll to latest message"),
              children: j.jsx(f, { size: 14 }),
            }),
            h$1 &&
              j.jsx("p", {
                role: "alert",
                className: "lg-chat-error",
                children: j.jsx("span", {
                  lang: language(h$1),
                  children: text(h$1),
                }),
              }),
            j.jsxs(ComposerPrimitive.Root, {
              className: `lg-chat-composer ${f$1 ? "dragging" : ""}`,
              onSubmit: (v) => {
                (v.preventDefault(), S$1(t.session.draft));
              },
              onDragOver: (v) => {
                v.dataTransfer.types.includes("Files") &&
                  (v.preventDefault(), p$1(true));
              },
              onDragLeave: (v) => {
                v.currentTarget.contains(v.relatedTarget) || p$1(false);
              },
              onDrop: (v) => {
                (v.preventDefault(),
                  p$1(false),
                  t.adapter.chat?.attachments
                    ? w(Array.from(v.dataTransfer.files))
                    : d$1("This provider does not support attachments."));
              },
              children: [
                t.session.attachments.length > 0 &&
                  j.jsx("div", {
                    className: "lg-chat-attachments",
                    role: "group",
                    "aria-label": text("Attached files"),
                    children: t.session.attachments.map((v) =>
                      j.jsx(
                        ai,
                        {
                          file: v,
                          remove: () =>
                            t.setSession((D) => ({
                              ...D,
                              attachments: D.attachments.filter(
                                ($) => $.id !== v.id,
                              ),
                            })),
                        },
                        v.id,
                      ),
                    ),
                  }),
                l$1 &&
                  j.jsxs("p", {
                    role: "status",
                    className: "lg-chat-preparing",
                    children: [
                      j.jsx(LoaderCircle, { size: 13, className: "lg-spin" }),
                      j.jsx(W, { children: "Preparing attachments…" }),
                    ],
                  }),
                j.jsx(ComposerPrimitive.Input, {
                  ref: a,
                  "aria-label": text("Assistant prompt"),
                  lang: language("Assistant prompt"),
                  "data-ui-label": "Assistant prompt",
                  placeholder: text("Message the workbook assistant…"),
                  value: t.session.draft,
                  onChange: (v) =>
                    t.setSession((D) => ({ ...D, draft: v.target.value })),
                  minRows: 2,
                  maxRows: 6,
                  maxLength: 1e4,
                  autoFocus: false,
                  cancelOnEscape: false,
                  addAttachmentOnPaste: false,
                  onKeyDown: (v) => v.stopPropagation(),
                  onPaste: (v) => {
                    v.clipboardData.files.length &&
                      (v.preventDefault(),
                      t.adapter.chat?.attachments
                        ? w(Array.from(v.clipboardData.files))
                        : d$1("This provider does not support attachments."));
                  },
                }),
                j.jsxs("div", {
                  className: "lg-chat-composer-actions",
                  children: [
                    j.jsxs("div", {
                      children: [
                        j.jsx("input", {
                          ref: o,
                          type: "file",
                          hidden: true,
                          multiple: true,
                          accept: h,
                          "aria-label": text("Upload chat attachments"),
                          onChange: (v) => {
                            const D = Array.from(v.currentTarget.files ?? []);
                            ((v.currentTarget.value = ""), w(D));
                          },
                        }),
                        j.jsx("button", {
                          type: "button",
                          "aria-label": text("Attach files"),
                          title: t.adapter.chat?.attachments
                            ? text("Attach files · up to 4 files, 5 MB each")
                            : text(
                                "File support is not enabled by this host adapter",
                              ),
                          disabled:
                            !t.adapter.chat?.attachments || t.busy || l$1,
                          onClick: () => o.current?.click(),
                          children: j.jsx(Paperclip, { size: 18 }),
                        }),
                        j.jsx("span", {
                          children: text(
                            f$1 ? "Drop files here" : "Enter to send",
                          ),
                        }),
                      ],
                    }),
                    t.busy
                      ? j.jsx("button", {
                          type: "button",
                          className: "lg-chat-send",
                          "aria-label": text("Cancel request"),
                          title: text("Stop response"),
                          onClick: t.onCancel,
                          children: j.jsx(i, {
                            size: 15,
                            fill: "currentColor",
                          }),
                        })
                      : j.jsx("button", {
                          type: "submit",
                          className: "lg-chat-send",
                          "aria-label": text("Send selection"),
                          title: text("Send message and selected context"),
                          disabled: !t.session.draft.trim() || l$1,
                          children: j.jsx(A, { size: 19 }),
                        }),
                  ],
                }),
              ],
            }),
            j.jsxs("p", {
              className: "lg-chat-footnote",
              children: [
                t.local
                  ? text("Local demo · files stay in this browser")
                  : t.adapter.chat
                    ? text(
                        "Selection, same-context chat and files go to your host",
                      )
                    : text(
                        "Selection only · file and chat-history support unavailable",
                      ),
                j.jsx("br", {}),
                j.jsx(W, {
                  children:
                    "Review proposed changes before applying. Shift+Enter for a new line.",
                }),
              ],
            }),
            j.jsx("span", {
              className: "lg-sr-only",
              role: "status",
              children: j.jsx("span", { lang: language(x), children: text(x) }),
            }),
          ],
        }),
      ],
    }),
  });
}

export { mb as AssistantChat };
