/**
 * A2UI: the workbook as a catalog component. An agent that is offered this catalog can place a
 * `Workbook` component on a surface; the host's A2UI renderer then shows the editor.
 *
 * Written against A2UI 1.0 and the official React renderer. The A2UI packages are not
 * dependencies of this library: the host passes in the pieces it already has, which also keeps
 * the host in charge of their versions.
 */
import { useId } from "react";
import { jsx } from "react/jsx-runtime";
import { AgentWorkbook } from "./AgentWorkbook.js";
import { workbookSpecLimits } from "./workbook-spec.js";

/** Identifies the catalog to agents and renderers. It is a name; nothing is fetched from it. */
export const WORKBOOK_CATALOG_ID =
  "https://github.com/rbaran17/Ledgerglass/blob/main/catalogs/workbook/v1/catalog.json";

export const WORKBOOK_COMPONENT_NAME = "Workbook";

/** Appended to the agent's instructions by A2UI tooling that reads catalog instructions. */
export const workbookCatalogInstructions = [
  "Use the Workbook component to show an interactive spreadsheet when tabular figures, calculations or a schedule are easier to read or adjust in a grid than in text.",
  '`sheets` is a list of worksheets: `{ "name": "Sales", "rows": [["Item", "Amount"], ["A", 10]], "cells": { "B3": "=SUM(B2:B2)" } }`.',
  '`rows` fills cells row by row from `startCell` (default "A1"); `cells` sets single cells by address. A cell is a number, text, true/false or null. Text starting with "=" is a formula (Excel syntax, A1 references).',
  `Keep a workbook to at most ${workbookSpecLimits.maxSheets} sheets and ${workbookSpecLimits.maxCells} filled cells.`,
  "Bind `sheets` to the data model to receive the user's edits: the component writes the edited sheets back to that path, each sheet as `{ name, cells }`.",
].join("\n");

/**
 * Builds the A2UI catalog holding the Workbook component.
 *
 * @param a2ui.Catalog                        from `@a2ui/web_core/v1_0`
 * @param a2ui.CommonSchemas                  from `@a2ui/web_core/v1_0`
 * @param a2ui.createComponentImplementation  from `@a2ui/react`
 * @param a2ui.z                              the `zod` the A2UI packages use
 * @param options.catalogId        Defaults to `WORKBOOK_CATALOG_ID`.
 * @param options.protocolVersion  Defaults to "1.0".
 * @param options.height           Default height in pixels. Defaults to 480.
 * @param options.editorProps      Props for every editor the catalog renders (adapter, actor, ...).
 */
export function createWorkbookCatalog(a2ui, options = {}) {
  const { Catalog, CommonSchemas, createComponentImplementation, z } = a2ui;
  if (!Catalog || !CommonSchemas || !createComponentImplementation || !z)
    throw new TypeError(
      "createWorkbookCatalog needs { Catalog, CommonSchemas, createComponentImplementation, z }.",
    );
  const defaultHeight = options.height ?? 480;
  const workbookApi = {
    name: WORKBOOK_COMPONENT_NAME,
    schema: z.object({
      title: CommonSchemas.DynamicString.optional().describe(
        "Short name of the workbook.",
      ),
      sheets: CommonSchemas.DynamicValue.describe(
        "The worksheets: a list of { name, rows?, startCell?, cells? }, or a binding to one in the data model.",
      ),
      readOnly: CommonSchemas.DynamicBoolean.optional().describe(
        "True to show the workbook without allowing edits.",
      ),
      height: z
        .number()
        .optional()
        .describe("Height in pixels. Defaults to 480."),
      onChange: CommonSchemas.Action.optional().describe(
        "Action sent after the user edits the workbook.",
      ),
    }),
  };
  const Workbook = createComponentImplementation(workbookApi, ({ props }) => {
    const instanceId = useId();
    return jsx(AgentWorkbook, {
      ...options.editorProps,
      workbookId: `a2ui${instanceId}`,
      height: props.height ?? defaultHeight,
      spec: {
        title: props.title ?? undefined,
        sheets: props.sheets,
        ...(props.readOnly ? { readOnly: true } : {}),
      },
      onSpecChange: (editedSpec) => {
        props.setSheets?.(editedSpec.sheets);
        props.onChange?.();
      },
    });
  });
  return new Catalog(
    options.catalogId ?? WORKBOOK_CATALOG_ID,
    options.protocolVersion ?? "1.0",
    [Workbook],
    [],
    undefined,
    workbookCatalogInstructions,
  );
}
