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

const db = (byteCount) =>
  byteCount < 1024 * 1024
    ? `${Math.max(1, Math.ceil(byteCount / 1024))} KB`
    : `${(byteCount / 1024 / 1024).toFixed(1)} MB`;
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
function mb(props) {
  const { text, language, message } = useWorkbookMessages(),
    r$1 = React.useRef(props);
  r$1.current = props;
  const i$1 = React.useRef(false),
    fileInputRef = React.useRef(null),
    promptInputRef = React.useRef(null),
    attachmentAbortRef = React.useRef(null),
    [l$1, u$1] = React.useState(false),
    [h$1, d$1] = React.useState(""),
    [f$1, p$1] = React.useState(false),
    [copiedMessageId, setCopiedMessageId] = React.useState(null),
    [statusMessage, setStatusMessage] = React.useState("");
  (React.useEffect(() => () => attachmentAbortRef.current?.abort(), []),
    React.useEffect(() => {
      (attachmentAbortRef.current?.abort(), u$1(false), d$1(""));
    }, [props.session.id, props.adapter]));
  async function attachFiles(files) {
    if (
      !files.length ||
      !props.adapter.chat?.attachments ||
      attachmentAbortRef.current ||
      l$1 ||
      props.busy
    )
      return;
    const attachments = props.session.attachments;
    if (
      (d$1(""),
      attachments.length + files.length >
        assistantRequestLimits.filesPerMessage)
    ) {
      d$1(
        "Attach up to 4 files per message. Remove a file before adding more.",
      );
      return;
    }
    if (
      attachments.reduce((total, attachment) => total + attachment.size, 0) +
        files.reduce((total, file) => total + file.size, 0) >
      assistantRequestLimits.totalFileBytes
    ) {
      d$1("Attachments must total 10 MB or less.");
      return;
    }
    if (
      new Set([
        ...attachments.map(
          (attachment) => `${attachment.name}:${attachment.size}`,
        ),
        ...files.map((file) => `${file.name}:${file.size}`),
      ]).size !==
      attachments.length + files.length
    ) {
      d$1("This file is already attached. Remove it before adding it again.");
      return;
    }
    const controller = new AbortController(),
      sessionId = props.session.id;
    ((attachmentAbortRef.current = controller), u$1(true));
    try {
      const preparedAttachments = await Promise.all(
        files.map((file) => prepareChatAttachment(file, controller.signal)),
      );
      controller.signal.aborted ||
        props.setSession((session) =>
          session.id === sessionId
            ? {
                ...session,
                attachments: [...session.attachments, ...preparedAttachments],
              }
            : session,
        );
    } catch (error) {
      controller.signal.aborted ||
        d$1(
          error instanceof Error
            ? error.message
            : "The attachment could not be read.",
        );
    } finally {
      attachmentAbortRef.current === controller &&
        (u$1(false), (attachmentAbortRef.current = null));
    }
  }
  async function S$1(v$1, attachmentsOverride) {
    const latestProps = r$1.current;
    if (i$1.current || latestProps.busy || l$1 || !v$1.trim()) return;
    const attachments = attachmentsOverride ?? latestProps.session.attachments,
      scope = latestProps.scope,
      sessionId = latestProps.session.id,
      chatHistory = [];
    for (
      let messageIndex = 1;
      messageIndex < latestProps.session.messages.length;
      messageIndex++
    ) {
      const assistantMessage = latestProps.session.messages[messageIndex],
        precedingUserMessage = latestProps.session.messages[messageIndex - 1];
      assistantMessage.role === "assistant" &&
        assistantMessage.outcome === "complete" &&
        assistantMessage.scope === scope &&
        precedingUserMessage.role === "user" &&
        precedingUserMessage.scope === scope &&
        chatHistory.push(
          {
            role: "user",
            text: precedingUserMessage.text,
            ...(precedingUserMessage.attachments?.length
              ? { attachments: precedingUserMessage.attachments }
              : {}),
          },
          { role: "assistant", text: assistantMessage.text },
        );
    }
    let chatRequest;
    try {
      if (latestProps.session.messages.length >= 60)
        throw Error(
          "Start a new chat to continue. This chat has reached its message limit.",
        );
      if (v$1.length > assistantRequestLimits.promptCharacters)
        throw Error("Keep your message under 10,000 characters.");
      if (attachments.length && !latestProps.adapter.chat?.attachments)
        throw Error(
          "This provider does not support file attachments. Remove the files or connect a chat-enabled provider.",
        );
      latestProps.adapter.chat &&
        (chatRequest = parseAssistantChatRequest({
          version: 1,
          messages: latestProps.adapter.chat.history ? chatHistory : [],
          attachments: attachments,
        }));
    } catch (error) {
      d$1(error.message);
      return;
    }
    const userMessage = {
      id: crypto.randomUUID(),
      role: "user",
      text: v$1.trim(),
      scope: scope,
      label: latestProps.selectionLabel,
      attachments: attachments,
    };
    ((i$1.current = true),
      d$1(""),
      setStatusMessage(""),
      latestProps.setSession((session) => ({
        ...session,
        draft: "",
        attachments: [],
        messages: [...session.messages, userMessage],
      })));
    try {
      const proposal = await latestProps.onSend(userMessage.text, chatRequest);
      latestProps.setSession((session) =>
        session.id !== sessionId
          ? session
          : {
              ...session,
              messages: [
                ...session.messages,
                {
                  id: crypto.randomUUID(),
                  role: "assistant",
                  text:
                    proposal?.explanation ??
                    "Request stopped. No changes were made.",
                  scope: scope,
                  label: userMessage.label,
                  proposalId: proposal?.id,
                  outcome: proposal ? "complete" : "stopped",
                },
              ],
            },
      );
    } catch {
      latestProps.setSession((session) =>
        session.id !== sessionId
          ? session
          : {
              ...session,
              draft: session.draft || userMessage.text,
              attachments: session.attachments.length
                ? session.attachments
                : attachments,
              messages: [
                ...session.messages,
                {
                  id: crypto.randomUUID(),
                  role: "assistant",
                  text: "The request did not complete. Your message is ready to retry.",
                  scope: scope,
                  label: userMessage.label,
                  outcome: "failed",
                },
              ],
            },
      );
    } finally {
      i$1.current = false;
    }
  }
  const messages = React.useMemo(
      () =>
        props.session.messages.map((message) => ({
          id: message.id,
          role: message.role,
          content:
            message.scope === props.scope
              ? message.text
              : text(
                  "This message belongs to an earlier selection or workbook revision and is excluded from the current context.",
                ),
          metadata: { custom: { source: message } },
        })),
      [props.session.messages, props.scope],
    ),
    externalStoreRuntime = useExternalStoreRuntime({
      messages: messages,
      convertMessage: (runtimeMessage) => runtimeMessage,
      isRunning: props.busy,
      onNew: async (newMessage) => {
        await S$1(
          newMessage.content
            .filter((part) => part.type === "text")
            .map((part) => part.text).join(`
`),
        );
      },
      onCancel: async () => props.onCancel(),
    }),
    k$1 = props.session.messages.some(
      (message) =>
        message.proposalId &&
        message.proposalId === props.proposalId &&
        message.scope === props.scope,
    ),
    onClick = () => {
      const lastUserMessage = [...props.session.messages]
        .reverse()
        .find((entry) => entry.role === "user" && entry.scope === props.scope);
      lastUserMessage && S$1(lastUserMessage.text, lastUserMessage.attachments);
    },
    P$1 = () => {
      (props.onCancel(),
        props.onReset(),
        attachmentAbortRef.current?.abort(),
        props.setSession(createEmptyAssistantConversation()),
        setStatusMessage("New chat started."),
        promptInputRef.current?.focus());
    },
    copyMessage = async (sourceMessage) => {
      try {
        (await navigator.clipboard.writeText(sourceMessage.text),
          setCopiedMessageId(sourceMessage.id),
          setStatusMessage("Message copied."));
      } catch {
        setStatusMessage(
          "Copy is unavailable. Select the message text and use your browser’s Copy command.",
        );
      }
    };
  return jsxRuntime.jsx(AssistantRuntimeProvider, {
    runtime: externalStoreRuntime,
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
              title: props.adapter.name,
              children: [
                jsxRuntime.jsx("i", {}),
                props.local ? text("Local demo") : props.adapter.name,
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
                jsxRuntime.jsx("strong", { children: props.selectionLabel }),
                jsxRuntime.jsx("small", {
                  children: message(
                    "{count} selected cells · shared when you send",
                    {
                      count: props.selectedCount,
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
          autoScroll: props.session.messages.length > 0,
          scrollToBottomOnInitialize: props.session.messages.length > 0,
          scrollToBottomOnThreadSwitch: false,
          children: [
            !props.session.messages.length &&
              !props.proposalId &&
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
                    ].map((draft) =>
                      jsxRuntime.jsxs(
                        "button",
                        {
                          type: "button",
                          disabled: props.busy || l$1,
                          onClick: () => {
                            (props.setSession((session) => ({
                              ...session,
                              draft: draft,
                            })),
                              S$1(draft));
                          },
                          children: [
                            jsxRuntime.jsx(LocalizedText, { children: draft }),
                            jsxRuntime.jsx(ArrowUp, { size: 13 }),
                          ],
                        },
                        draft,
                      ),
                    ),
                  }),
                ],
              }),
            jsxRuntime.jsx(ThreadPrimitive.Messages, {
              children: ({ message }) => {
                const source = message.metadata.custom.source;
                if (!source) return null;
                const isCurrentScope = source.scope === props.scope,
                  shouldShowReview =
                    isCurrentScope &&
                    source.proposalId === props.proposalId &&
                    !!props.review;
                return jsxRuntime.jsxs(MessagePrimitive.Root, {
                  role: "article",
                  className: `lg-chat-message ${source.role}`,
                  "aria-label":
                    source.role === "user"
                      ? text("Your message")
                      : text("Assistant message"),
                  children: [
                    jsxRuntime.jsxs("div", {
                      className: "lg-chat-message-label",
                      children: [
                        source.role === "user"
                          ? "You"
                          : jsxRuntime.jsxs(jsxRuntime.Fragment, {
                              children: [
                                jsxRuntime.jsx(Sparkles, { size: 13 }),
                                jsxRuntime.jsx(LocalizedText, {
                                  children: "Assistant",
                                }),
                              ],
                            }),
                        jsxRuntime.jsx("small", { children: source.label }),
                      ],
                    }),
                    jsxRuntime.jsxs("div", {
                      className: `lg-chat-bubble ${isCurrentScope ? "" : "archived"}`,
                      children: [
                        isCurrentScope &&
                          source.attachments?.map((attachment) =>
                            jsxRuntime.jsx(
                              ai,
                              { file: attachment },
                              attachment.id,
                            ),
                          ),
                        shouldShowReview
                          ? props.review
                          : jsxRuntime.jsx(AssistantChatMarkdown, {
                              children: isCurrentScope
                                ? source.text
                                : text(
                                    "This message belongs to an earlier selection or workbook revision and is excluded from the current context.",
                                  ),
                            }),
                      ],
                    }),
                    isCurrentScope &&
                      source.outcome === "complete" &&
                      jsxRuntime.jsxs("div", {
                        className: "lg-chat-message-actions",
                        children: [
                          jsxRuntime.jsx("button", {
                            type: "button",
                            "aria-label": text("Copy assistant response"),
                            title: text("Copy response"),
                            onClick: () => void copyMessage(source),
                            children:
                              copiedMessageId === source.id
                                ? jsxRuntime.jsx(Check, { size: 14 })
                                : jsxRuntime.jsx(Copy, { size: 14 }),
                          }),
                          source === props.session.messages.at(-1) &&
                            jsxRuntime.jsx("button", {
                              type: "button",
                              "aria-label": text("Regenerate response"),
                              title: text("Regenerate response"),
                              disabled: props.busy || l$1,
                              onClick: onClick,
                              children: jsxRuntime.jsx(RotateCcw, { size: 14 }),
                            }),
                        ],
                      }),
                    isCurrentScope &&
                      (source.outcome === "failed" ||
                        source.outcome === "stopped") &&
                      source === props.session.messages.at(-1) &&
                      jsxRuntime.jsxs("button", {
                        type: "button",
                        className: "lg-chat-retry",
                        disabled: props.busy || l$1,
                        onClick: onClick,
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
              props.review &&
              jsxRuntime.jsx("div", {
                className: "lg-chat-external-review",
                children: props.review,
              }),
            props.busy &&
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
              onSubmit: (event) => {
                (event.preventDefault(), S$1(props.session.draft));
              },
              onDragOver: (event) => {
                event.dataTransfer.types.includes("Files") &&
                  (event.preventDefault(), p$1(true));
              },
              onDragLeave: (event) => {
                event.currentTarget.contains(event.relatedTarget) || p$1(false);
              },
              onDrop: (event) => {
                (event.preventDefault(),
                  p$1(false),
                  props.adapter.chat?.attachments
                    ? attachFiles(Array.from(event.dataTransfer.files))
                    : d$1("This provider does not support attachments."));
              },
              children: [
                props.session.attachments.length > 0 &&
                  jsxRuntime.jsx("div", {
                    className: "lg-chat-attachments",
                    role: "group",
                    "aria-label": text("Attached files"),
                    children: props.session.attachments.map((attachment) =>
                      jsxRuntime.jsx(
                        ai,
                        {
                          file: attachment,
                          remove: () =>
                            props.setSession((session) => ({
                              ...session,
                              attachments: session.attachments.filter(
                                (entry) => entry.id !== attachment.id,
                              ),
                            })),
                        },
                        attachment.id,
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
                  ref: promptInputRef,
                  "aria-label": text("Assistant prompt"),
                  lang: language("Assistant prompt"),
                  "data-ui-label": "Assistant prompt",
                  placeholder: text("Message the workbook assistant…"),
                  value: props.session.draft,
                  onChange: (event) =>
                    props.setSession((session) => ({
                      ...session,
                      draft: event.target.value,
                    })),
                  minRows: 2,
                  maxRows: 6,
                  maxLength: 1e4,
                  autoFocus: false,
                  cancelOnEscape: false,
                  addAttachmentOnPaste: false,
                  onKeyDown: (event) => event.stopPropagation(),
                  onPaste: (event) => {
                    event.clipboardData.files.length &&
                      (event.preventDefault(),
                      props.adapter.chat?.attachments
                        ? attachFiles(Array.from(event.clipboardData.files))
                        : d$1("This provider does not support attachments."));
                  },
                }),
                jsxRuntime.jsxs("div", {
                  className: "lg-chat-composer-actions",
                  children: [
                    jsxRuntime.jsxs("div", {
                      children: [
                        jsxRuntime.jsx("input", {
                          ref: fileInputRef,
                          type: "file",
                          hidden: true,
                          multiple: true,
                          accept: attachmentAcceptAttribute,
                          "aria-label": text("Upload chat attachments"),
                          onChange: (event) => {
                            const selectedFiles = Array.from(
                              event.currentTarget.files ?? [],
                            );
                            ((event.currentTarget.value = ""),
                              attachFiles(selectedFiles));
                          },
                        }),
                        jsxRuntime.jsx("button", {
                          type: "button",
                          "aria-label": text("Attach files"),
                          title: props.adapter.chat?.attachments
                            ? text("Attach files · up to 4 files, 5 MB each")
                            : text(
                                "File support is not enabled by this host adapter",
                              ),
                          disabled:
                            !props.adapter.chat?.attachments ||
                            props.busy ||
                            l$1,
                          onClick: () => fileInputRef.current?.click(),
                          children: jsxRuntime.jsx(Paperclip, { size: 18 }),
                        }),
                        jsxRuntime.jsx("span", {
                          children: text(
                            f$1 ? "Drop files here" : "Enter to send",
                          ),
                        }),
                      ],
                    }),
                    props.busy
                      ? jsxRuntime.jsx("button", {
                          type: "button",
                          className: "lg-chat-send",
                          "aria-label": text("Cancel request"),
                          title: text("Stop response"),
                          onClick: props.onCancel,
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
                          disabled: !props.session.draft.trim() || l$1,
                          children: jsxRuntime.jsx(ArrowUp, { size: 19 }),
                        }),
                  ],
                }),
              ],
            }),
            jsxRuntime.jsxs("p", {
              className: "lg-chat-footnote",
              children: [
                props.local
                  ? text("Local demo · files stay in this browser")
                  : props.adapter.chat
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
                lang: language(statusMessage),
                children: text(statusMessage),
              }),
            }),
          ],
        }),
      ],
    }),
  });
}

export { mb as AssistantChat };
