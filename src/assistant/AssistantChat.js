import {
  useExternalStoreRuntime,
  AssistantRuntimeProvider,
  ThreadPrimitive,
  MessagePrimitive,
  ComposerPrimitive,
} from "@assistant-ui/react";
import { LoaderCircle, Paperclip } from "lucide-react";
import "../styles/assistant.css";
import jsxRuntime from "react/jsx-runtime";
import React from "react";
import {
  ArrowDown,
  ArrowUp,
  Check,
  Copy,
  FileText,
  Plus,
  RotateCcw,
  Sparkles,
  Square,
  Table2,
  X,
} from "lucide-react";
import {
  LocalizedText,
  createEmptyAssistantConversation,
  useWorkbookMessages,
} from "../app/04-cell-input-paste-and-fill.js";
import {
  assistantRequestLimits,
  parseAssistantChatRequest,
} from "../app/19b-assistant-protocol.js";
import {
  attachmentAcceptAttribute,
  prepareChatAttachment,
} from "../app/20-calculation-options-and-assistant.js";
import { AssistantChatMarkdown } from "./AssistantChatMarkdown.js";

const db = (t) =>
  t < 1024 * 1024
    ? `${Math.max(1, Math.ceil(t / 1024))} KB`
    : `${(t / 1024 / 1024).toFixed(1)} MB`;
function ai({ file, remove }) {
  const { text, language, message } = useWorkbookMessages();
  return jsxRuntime.jsxs("div", {
    className: "lg-chat-file",
    children: [
      jsxRuntime.jsxs("details", {
        children: [
          jsxRuntime.jsxs("summary", {
            children: [
              jsxRuntime.jsx(FileText, { size: 17, "aria-hidden": "true" }),
              jsxRuntime.jsxs("span", {
                children: [
                  jsxRuntime.jsx("strong", { children: file.name }),
                  jsxRuntime.jsxs("small", {
                    children: [
                      db(file.size),
                      " ",
                      jsxRuntime.jsx(LocalizedText, { children: "·" }),
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
            ? jsxRuntime.jsxs("pre", {
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
              ? jsxRuntime.jsx("img", {
                  src: `data:${file.mediaType};base64,${file.data}`,
                  alt: message("Preview of {name}", { name: file.name }),
                })
              : jsxRuntime.jsx("p", {
                  className: "lg-chat-file-description",
                  children: jsxRuntime.jsx(LocalizedText, {
                    children:
                      "The original file will accompany your message. Preview is unavailable for this format.",
                  }),
                }),
        ],
      }),
      remove &&
        jsxRuntime.jsx("button", {
          type: "button",
          "aria-label": message("Remove attachment {name}", {
            name: file.name,
          }),
          title: text("Remove attachment"),
          onClick: remove,
          children: jsxRuntime.jsx(X, { size: 14 }),
        }),
    ],
  });
}
function mb(t) {
  const { text, language, message } = useWorkbookMessages(),
    r$1 = React.useRef(t);
  r$1.current = t;
  const i$1 = React.useRef(false),
    o = React.useRef(null),
    a = React.useRef(null),
    c = React.useRef(null),
    [l$1, u$1] = React.useState(false),
    [h$1, d$1] = React.useState(""),
    [f$1, p$1] = React.useState(false),
    [m, _] = React.useState(null),
    [x, y] = React.useState("");
  (React.useEffect(() => () => c.current?.abort(), []),
    React.useEffect(() => {
      (c.current?.abort(), u$1(false), d$1(""));
    }, [t.session.id, t.adapter]));
  async function w(v) {
    if (!v.length || !t.adapter.chat?.attachments || c.current || l$1 || t.busy)
      return;
    const D = t.session.attachments;
    if (
      (d$1(""), D.length + v.length > assistantRequestLimits.filesPerMessage)
    ) {
      d$1(
        "Attach up to 4 files per message. Remove a file before adding more.",
      );
      return;
    }
    if (
      D.reduce((F, H) => F + H.size, 0) + v.reduce((F, H) => F + H.size, 0) >
      assistantRequestLimits.totalFileBytes
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
      const F = await Promise.all(
        v.map((H) => prepareChatAttachment(H, $.signal)),
      );
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
      if (v$1.length > assistantRequestLimits.promptCharacters)
        throw Error("Keep your message under 10,000 characters.");
      if (U.length && !$.adapter.chat?.attachments)
        throw Error(
          "This provider does not support file attachments. Remove the files or connect a chat-enabled provider.",
        );
      $.adapter.chat &&
        (J = parseAssistantChatRequest({
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
  const I = React.useMemo(
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
        t.setSession(createEmptyAssistantConversation()),
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
  return jsxRuntime.jsx(AssistantRuntimeProvider, {
    runtime: R,
    children: jsxRuntime.jsxs(ThreadPrimitive.Root, {
      className: "lg-chat",
      role: "group",
      "aria-label": text("Workbook chat"),
      children: [
        jsxRuntime.jsxs("div", {
          className: "lg-chat-toolbar",
          children: [
            jsxRuntime.jsxs("span", {
              className: "lg-chat-provider",
              title: t.adapter.name,
              children: [
                jsxRuntime.jsx("i", {}),
                t.local ? text("Local demo") : t.adapter.name,
              ],
            }),
            jsxRuntime.jsxs("button", {
              type: "button",
              onClick: P$1,
              "aria-label": text("New chat"),
              title: text("New chat"),
              children: [
                jsxRuntime.jsx(Plus, { size: 16 }),
                jsxRuntime.jsx("span", {
                  children: jsxRuntime.jsx(LocalizedText, {
                    children: "New chat",
                  }),
                }),
              ],
            }),
          ],
        }),
        jsxRuntime.jsxs("div", {
          className: "lg-chat-context",
          role: "group",
          "aria-label": text("Shared workbook context"),
          children: [
            jsxRuntime.jsx(Table2, { size: 16 }),
            jsxRuntime.jsxs("span", {
              children: [
                jsxRuntime.jsx("strong", { children: t.selectionLabel }),
                jsxRuntime.jsx("small", {
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
        jsxRuntime.jsxs(ThreadPrimitive.Viewport, {
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
              jsxRuntime.jsxs("div", {
                className: "lg-chat-welcome",
                children: [
                  jsxRuntime.jsx("div", {
                    className: "lg-chat-welcome-icon",
                    children: jsxRuntime.jsx(Sparkles, { size: 24 }),
                  }),
                  jsxRuntime.jsx("h3", {
                    children: jsxRuntime.jsx(LocalizedText, {
                      children: "How can I help?",
                    }),
                  }),
                  jsxRuntime.jsx("p", {
                    children: jsxRuntime.jsx(LocalizedText, {
                      children:
                        "Ask about your selection, attach a file, or work through a formula together.",
                    }),
                  }),
                  jsxRuntime.jsx("div", {
                    className: "lg-chat-suggestions",
                    children: [
                      "Explain this selection",
                      "Highlight values above 25000",
                      "Add a total formula",
                    ].map((v) =>
                      jsxRuntime.jsxs(
                        "button",
                        {
                          type: "button",
                          disabled: t.busy || l$1,
                          onClick: () => {
                            (t.setSession((D) => ({ ...D, draft: v })), S$1(v));
                          },
                          children: [
                            jsxRuntime.jsx(LocalizedText, { children: v }),
                            jsxRuntime.jsx(ArrowUp, { size: 13 }),
                          ],
                        },
                        v,
                      ),
                    ),
                  }),
                ],
              }),
            jsxRuntime.jsx(ThreadPrimitive.Messages, {
              children: ({ message }) => {
                const D = message.metadata.custom.source;
                if (!D) return null;
                const $ = D.scope === t.scope,
                  U = $ && D.proposalId === t.proposalId && !!t.review;
                return jsxRuntime.jsxs(MessagePrimitive.Root, {
                  role: "article",
                  className: `lg-chat-message ${D.role}`,
                  "aria-label":
                    D.role === "user"
                      ? text("Your message")
                      : text("Assistant message"),
                  children: [
                    jsxRuntime.jsxs("div", {
                      className: "lg-chat-message-label",
                      children: [
                        D.role === "user"
                          ? "You"
                          : jsxRuntime.jsxs(jsxRuntime.Fragment, {
                              children: [
                                jsxRuntime.jsx(Sparkles, { size: 13 }),
                                jsxRuntime.jsx(LocalizedText, {
                                  children: "Assistant",
                                }),
                              ],
                            }),
                        jsxRuntime.jsx("small", { children: D.label }),
                      ],
                    }),
                    jsxRuntime.jsxs("div", {
                      className: `lg-chat-bubble ${$ ? "" : "archived"}`,
                      children: [
                        $ &&
                          D.attachments?.map((F) =>
                            jsxRuntime.jsx(ai, { file: F }, F.id),
                          ),
                        U
                          ? t.review
                          : jsxRuntime.jsx(AssistantChatMarkdown, {
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
                      jsxRuntime.jsxs("div", {
                        className: "lg-chat-message-actions",
                        children: [
                          jsxRuntime.jsx("button", {
                            type: "button",
                            "aria-label": text("Copy assistant response"),
                            title: text("Copy response"),
                            onClick: () => void E(D),
                            children:
                              m === D.id
                                ? jsxRuntime.jsx(Check, { size: 14 })
                                : jsxRuntime.jsx(Copy, { size: 14 }),
                          }),
                          D === t.session.messages.at(-1) &&
                            jsxRuntime.jsx("button", {
                              type: "button",
                              "aria-label": text("Regenerate response"),
                              title: text("Regenerate response"),
                              disabled: t.busy || l$1,
                              onClick: M,
                              children: jsxRuntime.jsx(RotateCcw, { size: 14 }),
                            }),
                        ],
                      }),
                    $ &&
                      (D.outcome === "failed" || D.outcome === "stopped") &&
                      D === t.session.messages.at(-1) &&
                      jsxRuntime.jsxs("button", {
                        type: "button",
                        className: "lg-chat-retry",
                        disabled: t.busy || l$1,
                        onClick: M,
                        children: [
                          jsxRuntime.jsx(RotateCcw, { size: 13 }),
                          jsxRuntime.jsx(LocalizedText, {
                            children: "Retry message",
                          }),
                        ],
                      }),
                  ],
                });
              },
            }),
            !k$1 &&
              t.review &&
              jsxRuntime.jsx("div", {
                className: "lg-chat-external-review",
                children: t.review,
              }),
            t.busy &&
              jsxRuntime.jsxs("div", {
                role: "status",
                className: "lg-chat-thinking",
                children: [
                  jsxRuntime.jsx(LoaderCircle, {
                    size: 15,
                    className: "lg-spin",
                  }),
                  jsxRuntime.jsx(LocalizedText, {
                    children: "Preparing a response…",
                  }),
                ],
              }),
          ],
        }),
        jsxRuntime.jsxs("div", {
          className: "lg-chat-composer-area",
          children: [
            jsxRuntime.jsx(ThreadPrimitive.ScrollToBottom, {
              className: "lg-chat-scroll",
              "aria-label": text("Scroll to latest message"),
              title: text("Scroll to latest message"),
              children: jsxRuntime.jsx(ArrowDown, { size: 14 }),
            }),
            h$1 &&
              jsxRuntime.jsx("p", {
                role: "alert",
                className: "lg-chat-error",
                children: jsxRuntime.jsx("span", {
                  lang: language(h$1),
                  children: text(h$1),
                }),
              }),
            jsxRuntime.jsxs(ComposerPrimitive.Root, {
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
                  jsxRuntime.jsx("div", {
                    className: "lg-chat-attachments",
                    role: "group",
                    "aria-label": text("Attached files"),
                    children: t.session.attachments.map((v) =>
                      jsxRuntime.jsx(
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
                  jsxRuntime.jsxs("p", {
                    role: "status",
                    className: "lg-chat-preparing",
                    children: [
                      jsxRuntime.jsx(LoaderCircle, {
                        size: 13,
                        className: "lg-spin",
                      }),
                      jsxRuntime.jsx(LocalizedText, {
                        children: "Preparing attachments…",
                      }),
                    ],
                  }),
                jsxRuntime.jsx(ComposerPrimitive.Input, {
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
                jsxRuntime.jsxs("div", {
                  className: "lg-chat-composer-actions",
                  children: [
                    jsxRuntime.jsxs("div", {
                      children: [
                        jsxRuntime.jsx("input", {
                          ref: o,
                          type: "file",
                          hidden: true,
                          multiple: true,
                          accept: attachmentAcceptAttribute,
                          "aria-label": text("Upload chat attachments"),
                          onChange: (v) => {
                            const D = Array.from(v.currentTarget.files ?? []);
                            ((v.currentTarget.value = ""), w(D));
                          },
                        }),
                        jsxRuntime.jsx("button", {
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
                          children: jsxRuntime.jsx(Paperclip, { size: 18 }),
                        }),
                        jsxRuntime.jsx("span", {
                          children: text(
                            f$1 ? "Drop files here" : "Enter to send",
                          ),
                        }),
                      ],
                    }),
                    t.busy
                      ? jsxRuntime.jsx("button", {
                          type: "button",
                          className: "lg-chat-send",
                          "aria-label": text("Cancel request"),
                          title: text("Stop response"),
                          onClick: t.onCancel,
                          children: jsxRuntime.jsx(Square, {
                            size: 15,
                            fill: "currentColor",
                          }),
                        })
                      : jsxRuntime.jsx("button", {
                          type: "submit",
                          className: "lg-chat-send",
                          "aria-label": text("Send selection"),
                          title: text("Send message and selected context"),
                          disabled: !t.session.draft.trim() || l$1,
                          children: jsxRuntime.jsx(ArrowUp, { size: 19 }),
                        }),
                  ],
                }),
              ],
            }),
            jsxRuntime.jsxs("p", {
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
                jsxRuntime.jsx("br", {}),
                jsxRuntime.jsx(LocalizedText, {
                  children:
                    "Review proposed changes before applying. Shift+Enter for a new line.",
                }),
              ],
            }),
            jsxRuntime.jsx("span", {
              className: "lg-sr-only",
              role: "status",
              children: jsxRuntime.jsx("span", {
                lang: language(x),
                children: text(x),
              }),
            }),
          ],
        }),
      ],
    }),
  });
}

export { mb as AssistantChat };
