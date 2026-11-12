/**
 * Recording of edits. In the editor, every change to a workbook advances `revision` and adds an
 * entry to `events`: cell edits through `applyCellPatchTransaction`, which lists each changed
 * cell, and every other edit (sheets, panes, formats, tables, rules) by the editor itself, with an
 * empty change list. The model functions for those other edits return a new workbook and leave
 * the recording to their caller.
 *
 * The SDK is such a caller: the functions it exports from the worksheets, cells, tables, rules
 * and calculation areas are wrapped here so that an edit made through them is recorded the way
 * the editor records it.
 */
import { touchWorkbookModifiedAt } from "../app/03-workbook-model.js";

const defaultActor = "Preparer";

function isWorkbook(value) {
  return (
    value !== null &&
    typeof value === "object" &&
    Array.isArray(value.sheets) &&
    Array.isArray(value.events) &&
    Number.isInteger(value.revision) &&
    typeof value.id === "string"
  );
}

function requireLabel(value, name, maximum) {
  if (typeof value !== "string" || !value.trim() || value.length > maximum)
    throw new TypeError(`${name} must be 1 to ${maximum} characters of text.`);
  return value;
}

/**
 * Records one edit: returns `after` with `revision` one higher than `before`, an event added to
 * the log and the modified time updated. Use it for an edit you made yourself, for example by
 * building a workbook object by hand. The SDK's own edit functions already call it.
 *
 * If `after` is the same object as `before`, or has already been recorded (its revision is
 * higher), it is returned unchanged.
 *
 * @param details.action  What was done, for the log, such as "Add worksheet". Required.
 * @param details.actor   Who did it. Default "Preparer".
 */
export function recordWorkbookChange(before, after, details = {}) {
  if (!isWorkbook(before) || !isWorkbook(after))
    throw new TypeError("recordWorkbookChange needs two workbooks.");
  const action = requireLabel(details.action, "action", 200);
  const actor = requireLabel(details.actor ?? defaultActor, "actor", 100);
  if (after === before || after.revision > before.revision) return after;
  if (before.revision >= Number.MAX_SAFE_INTEGER - 1)
    throw new Error("Workbook revision limit reached.");
  return {
    ...touchWorkbookModifiedAt(after),
    revision: before.revision + 1,
    events: [
      ...before.events,
      {
        id: crypto.randomUUID(),
        timestamp: new Date().toISOString(),
        actor,
        action,
        // Only cell patches list changes; the editor records other edits the same way.
        changes: [],
      },
    ],
  };
}

/** "renameWorksheet" -> "Rename worksheet" */
function actionLabelFor(name) {
  const words = name.replace(/([a-z0-9])([A-Z])/g, "$1 $2").toLowerCase();
  return words.charAt(0).toUpperCase() + words.slice(1);
}

function recordResult(before, result, action, actor) {
  if (isWorkbook(result) && result.id === before.id)
    return recordWorkbookChange(before, result, { action, actor });
  // Some edits return the workbook together with a report.
  if (
    result !== null &&
    typeof result === "object" &&
    !Array.isArray(result) &&
    isWorkbook(result.workbook) &&
    result.workbook.id === before.id
  ) {
    const workbook = recordWorkbookChange(before, result.workbook, {
      action,
      actor,
    });
    return workbook === result.workbook ? result : { ...result, workbook };
  }
  return result;
}

/**
 * Wraps a model function so that, when it is given a workbook and returns a changed one, the
 * change is recorded. Anything else it returns passes through untouched, as do errors.
 */
export function recorded(name, modelFunction, actor = defaultActor) {
  const action = actionLabelFor(name);
  const wrapper = function (workbook, ...rest) {
    const result = modelFunction(workbook, ...rest);
    return isWorkbook(workbook)
      ? recordResult(workbook, result, action, actor)
      : result;
  };
  Object.defineProperty(wrapper, "name", { value: name });
  return wrapper;
}

/**
 * Wraps every lower-case-named function of a module. Classes and constants are passed through.
 */
export function recordEdits(model, actor = defaultActor) {
  const wrapped = {};
  for (const [name, value] of Object.entries(model))
    wrapped[name] =
      typeof value === "function" && /^[a-z]/.test(name)
        ? recorded(name, value, actor)
        : value;
  return wrapped;
}

import * as calculationModel from "./model/calculation.js";
import * as cellsModel from "./model/cells.js";
import * as rulesModel from "./model/rules.js";
import * as tablesModel from "./model/tables.js";
import * as worksheetsModel from "./model/worksheets.js";

/**
 * The edit functions of the worksheets, cells, tables, rules and calculation areas, recording
 * their edits under the given name instead of "Preparer".
 *
 *   const edits = createActorEdits("A. Patel");
 *   const next = edits.addWorksheet(workbook, "Notes");
 *   next.events.at(-1).actor; // "A. Patel"
 *
 * For cell edits, pass the actor to `applyCellPatchTransaction` as before.
 */
export function createActorEdits(actor) {
  requireLabel(actor, "actor", 100);
  return Object.freeze({
    ...recordEdits(calculationModel, actor),
    ...recordEdits(cellsModel, actor),
    ...recordEdits(rulesModel, actor),
    ...recordEdits(tablesModel, actor),
    ...recordEdits(worksheetsModel, actor),
  });
}
