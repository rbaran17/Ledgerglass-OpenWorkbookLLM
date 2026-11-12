/**
 * The editor as an agent-rendered component: give it the description an agent produced and it
 * shows the workbook, or a plain explanation when the description cannot be shown.
 */
import { useCallback, useRef } from "react";
import { jsx, jsxs } from "react/jsx-runtime";
import { WorkbookSyncBoundary } from "../app/56-main-worksheet-editor.js";
import {
  WorkbookSpecError,
  buildWorkbookFromSpec,
  describeWorkbookAsSpec,
} from "./workbook-spec.js";

function hashText(text) {
  let hash = 2166136261;
  for (let index = 0; index < text.length; index++)
    hash = Math.imul(hash ^ text.charCodeAt(index), 16777619);
  return (hash >>> 0).toString(36);
}

function buildFromSpec(spec, workbookId, specKey) {
  try {
    return buildWorkbookFromSpec(spec, {
      id: `${workbookId}:${hashText(specKey)}`,
    });
  } catch (error) {
    return {
      problems:
        error instanceof WorkbookSpecError
          ? error.problems
          : ["The workbook could not be shown."],
    };
  }
}

/**
 * @param props.spec        The agent's description (see `workbookSpecSchema`).
 * @param props.workbookId  Stable identifier, such as the tool call or surface identifier.
 * @param props.height      Height of the component; a number is pixels. Defaults to 480.
 * @param props.onSpecChange  Called after the user edits, with the edited workbook as a
 *   description and as a workbook. Passing that description back as `spec` keeps the editor as
 *   it is, so the component can be bound two ways.
 * @param props.onInvalid   Called with the list of problems when the description is rejected.
 * Other props (adapter, actor, onSelectionChange, ...) go to the editor.
 */
export function AgentWorkbook({
  spec,
  workbookId = "agent-workbook",
  height = 480,
  onSpecChange,
  onInvalid,
  onChange,
  readOnly,
  ...editorProps
}) {
  const stateRef = useRef({ specKey: null, built: null, emittedKey: null });
  const state = stateRef.current;
  const specKey = JSON.stringify(spec ?? null);
  if (specKey !== state.specKey) {
    if (specKey !== state.emittedKey || !state.built) {
      state.built = buildFromSpec(spec, workbookId, specKey);
      state.emittedKey = null;
      if (state.built.problems) onInvalid?.(state.built.problems);
    }
    state.specKey = specKey;
  }
  const built = state.built;
  const isReadOnly = readOnly ?? built.readOnly ?? false;

  const handleChange = useCallback(
    (workbook) => {
      onChange?.(workbook);
      if (!onSpecChange) return;
      const editedSpec = describeWorkbookAsSpec(workbook, {
        readOnly: stateRef.current.built?.readOnly,
      });
      stateRef.current.emittedKey = JSON.stringify(editedSpec);
      onSpecChange(editedSpec, workbook);
    },
    [onChange, onSpecChange],
  );

  const frameStyle = {
    height: typeof height === "number" ? `${height}px` : height,
    display: "flex",
    flexDirection: "column",
    minWidth: 0,
  };
  if (built.problems)
    return jsxs("div", {
      role: "alert",
      className:
        "ledgerglass-agent-workbook ledgerglass-agent-workbook-invalid",
      style: {
        padding: "12px 16px",
        border: "1px solid #c9ccd1",
        borderRadius: "6px",
        font: "14px/1.4 system-ui, sans-serif",
      },
      children: [
        jsx("strong", { children: "This workbook could not be shown." }),
        jsx("ul", {
          style: { margin: "8px 0 0", paddingLeft: "20px" },
          children: built.problems.map((problem, index) =>
            jsx("li", { children: problem }, index),
          ),
        }),
      ],
    });
  return jsx("div", {
    className: "ledgerglass-agent-workbook",
    "data-workbook-id": built.workbook.id,
    style: frameStyle,
    children: jsx("div", {
      style: { flex: "1 1 auto", minHeight: 0, position: "relative" },
      children: jsx(WorkbookSyncBoundary, {
        ...editorProps,
        initialWorkbook: built.workbook,
        readOnly: isReadOnly,
        onChange: handleChange,
      }),
    }),
  });
}
