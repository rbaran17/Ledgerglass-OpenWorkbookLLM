/**
 * Type declarations for the worksheet SDK surface (`src/sdk/worksheets.js`).
 *
 * Every function takes a workbook and returns a new workbook unless its doc comment says otherwise.
 * Functions that return a new workbook never modify the argument. Each of them throws an Error when
 * the change is not allowed, such as a duplicate name or a protected structure.
 */
import type { Workbook, Sheet } from "../index";

/** Whether a worksheet tab is shown, hidden, or hidden so the user cannot unhide it. */
export type WorksheetVisibility = "visible" | "hidden" | "veryHidden";

/** The two directions of a worksheet grid. */
export type WorksheetAxis = "row" | "column";

/** An inclusive, zero-based run of hidden rows or columns. */
export interface HiddenAxisInterval {
  start: number;
  end: number;
}

/** A split at a zero-based cell, with divider positions as percentages between 20 and 80. */
export interface SplitPaneState {
  row: number;
  col: number;
  horizontal?: number;
  vertical?: number;
}

/** A group of rows or columns that can be collapsed. */
export interface OutlineGroup {
  id: string;
  axis: WorksheetAxis;
  start: number;
  end: number;
  collapsed: boolean;
  [extra: string]: unknown;
}

/** One row of the list returned by `summarizeWorksheets`. */
export interface WorksheetSummary {
  sheetId: string;
  name: string;
  visibility: WorksheetVisibility;
  /** A "#rrggbb" colour, or null when the tab has no colour. */
  tabColor: string | null;
  storedCellCount: number;
  formulaCellCount: number;
  calculationParticipation: string;
  /** True when the workbook structure is protected. */
  structureLocked: boolean;
  [extra: string]: unknown;
}

/** One place that a worksheet deletion would change. `label` is a sentence for people to read. */
export interface DeletionImpact {
  kind: string;
  label: string;
  sheetId?: string;
  address?: string;
  [extra: string]: unknown;
}

/** Fields that `updateWorkbookProperties` may change. An empty engagement or period removes it. */
export interface WorkbookPropertyUpdates {
  name?: string;
  engagement?: string;
  period?: string;
}

/** Name, engagement and period for a workbook created from a template. All three are required. */
export interface WorkbookTemplateDetails {
  name: string;
  engagement: string;
  period: string;
}

/** Protection state of one worksheet. `salt` and `hash` are present only when a password was set. */
export interface SheetProtectionState {
  enabled: boolean;
  salt?: string;
  hash?: string;
  [extra: string]: unknown;
}

/** Row height or column width changes for one worksheet, in pixels. */
export interface SheetSizeEdits {
  rowHeight?: number;
  columnWidth?: number;
  resetRowOverrides?: boolean;
  resetColumnOverrides?: boolean;
}

/** Options for `applySheetSizeEdits`. */
export interface SheetSizeEditOptions {
  readOnly?: boolean;
  /** Fails if the workbook has changed since the values were read. */
  expectedSnapshot?: {
    workbookId: string;
    revision: number;
    snapshotGuard: string;
  };
}

/** Options for outline group changes. */
export interface OutlineEditOptions {
  readOnly?: boolean;
}

/** Restricts `summarizeWorksheets` to some worksheets. */
export interface WorksheetSummaryOptions {
  allowedSheetIds?: string[];
}

/** Returns a new workbook with a worksheet added at the end. Throws if the name is invalid or already used, or if the structure is protected. */
export function addWorksheet(
  workbook: Workbook,
  worksheetName: string,
): Workbook;

/** Returns a new workbook with the worksheet renamed, and formulas that use its name rewritten to match. Returns the same object when the name is unchanged. Throws if the name is invalid or already used. */
export function renameWorksheet(
  workbook: Workbook,
  sheetId: string,
  newName: string,
): Workbook;

/** Returns a new workbook with the worksheet moved to `targetIndex`, counted from zero. Returns the same object when the index is unchanged. Throws if the index is outside the workbook or a formula uses a three-dimensional reference. */
export function moveWorksheet(
  workbook: Workbook,
  sheetId: string,
  targetIndex: number,
): Workbook;

/** Returns a new workbook with a copy of the worksheet placed after it, named with the next free " (n)" suffix. Formulas on the copy that point at the original sheet are rewritten to point at the copy. Throws if a formula uses a three-dimensional reference. */
export function duplicateWorksheet(
  workbook: Workbook,
  sheetId: string,
): Workbook;

/** Returns a new workbook without the worksheet; formulas that referred to it become #REF!. If anything depends on the worksheet, pass `true` as `isDeletionConfirmed`, otherwise this throws. Throws for the only worksheet and when no visible worksheet would remain. */
export function deleteWorksheet(
  workbook: Workbook,
  sheetId: string,
  isDeletionConfirmed?: boolean,
): Workbook;

/** Returns a new workbook with the worksheet shown, hidden, or very hidden. Returns the same object when the visibility is unchanged. Throws if the value is not one of the three names or no visible worksheet would remain. */
export function setWorksheetVisibility(
  workbook: Workbook,
  sheetId: string,
  visibility: WorksheetVisibility,
): Workbook;

/** Returns a new workbook with the tab colour set to a "#rrggbb" value, or cleared when `tabColor` is undefined. Throws if the colour is not six hexadecimal digits with a leading "#". */
export function setWorksheetTabColor(
  workbook: Workbook,
  sheetId: string,
  tabColor: string | undefined,
): Workbook;

/** Returns a new workbook with the given name, trimmed. Throws if the name is empty or longer than 200 characters. */
export function renameWorkbook(workbook: Workbook, newName: string): Workbook;

/** Returns a new workbook with the given name, engagement or period changed. Fields that are not listed are kept; an empty engagement or period removes it. Throws for unknown fields or values that are too long. */
export function updateWorkbookProperties(
  workbook: Workbook,
  propertyUpdates: WorkbookPropertyUpdates,
): Workbook;

/** Returns a new workbook with a copy of the template's cells, sheets and settings under a new id and the given name, engagement and period. The template is not modified. Throws if any of the three details is empty or too long. */
export function createWorkbookFromTemplate(
  workbook: Workbook,
  details: WorkbookTemplateDetails,
): Workbook;

/** Returns a new workbook with the structure protected. An empty password protects the structure without one. Throws if the password is not text or is longer than 200 characters. */
export function protectWorkbookStructure(
  workbook: Workbook,
  password: string,
): Promise<Workbook>;

/** Returns a promise of a new workbook with the structure protection removed. Throws if a password was set and `password` does not match it. Resolves to the same object when the structure is not protected. */
export function unprotectWorkbookStructure(
  workbook: Workbook,
  password: string,
): Promise<Workbook>;

/** Returns a new workbook with the worksheet protected (`isProtecting` true) or unprotected (false). An empty password means no password. Throws when protecting an already protected sheet, when unprotecting with the wrong password, or when `password` is not text. */
export function updateSheetProtection(
  workbook: Workbook,
  sheetId: string,
  password: string,
  isProtecting: boolean,
): Promise<Workbook>;

/** Returns one summary per worksheet, in tab order, with its visibility, tab colour, cell counts and whether the structure is locked. Pass `allowedSheetIds` to list only some sheets. Does not change the workbook. */
export function summarizeWorksheets(
  workbook: Workbook,
  options?: WorksheetSummaryOptions,
): WorksheetSummary[];

/** Returns one entry for each formula, rule or refresh setting elsewhere in the workbook that deleting the worksheet would change, without changing the workbook. An empty list means nothing depends on it. Throws if the worksheet is not found. */
export function describeSheetDeletionImpacts(
  workbook: Workbook,
  sheetId: string,
): DeletionImpact[];

/** Returns the worksheet with the given id, as stored in the workbook. Throws if no worksheet has that id. */
export function findSheetOrThrow(workbook: Workbook, sheetId: string): Sheet;

/** Returns a new workbook with the given number of rows and columns frozen at the top and left. Removes any split panes. Throws if the counts are negative or larger than the grid. */
export function setFreezePanes(
  workbook: Workbook,
  sheetId: string,
  frozenRows: number,
  frozenCols?: number,
): Workbook;

/** Returns a new workbook with the rows above and the columns left of the given cell frozen, for example "B3" freezes two rows and one column. Throws if the address is not a valid cell. */
export function setFreezePanesAtCell(
  workbook: Workbook,
  sheetId: string,
  cellAddress: string,
): Workbook;

/** Returns a new workbook with split panes set, or removed when `splitPanes` is undefined. A split clears any frozen panes. Throws if neither divider is given, a divider is outside 20 to 80, or the anchor cell is outside the grid. */
export function setSplitPanes(
  workbook: Workbook,
  sheetId: string,
  splitPanes: SplitPaneState | undefined,
): Workbook;

/** Returns a new workbook with the range merged into one cell, or with overlapping merges removed when `unmerge` is true. The anchor cell keeps its value and is centred. Throws for a range over 500 cells, a merge that overlaps an existing one, or a protected sheet. */
export function mergeWorksheetCells(
  workbook: Workbook,
  sheetId: string,
  rangeText: string,
  unmerge?: boolean,
): Workbook;

/** Returns a new workbook with the inclusive, zero-based rows or columns from `start` to `end` hidden or shown. Throws if the axis is not "row" or "column", the range is outside the grid, or the sheet is protected. */
export function setHiddenAxisRange(
  workbook: Workbook,
  sheetId: string,
  axis: WorksheetAxis,
  start: number,
  end: number,
  isHidden: boolean,
): Workbook;

/** Returns the hidden row or column runs of a worksheet as sorted, merged intervals. Returns an empty list when nothing is hidden. Does not change the sheet. Throws if the axis is not "row" or "column". */
export function readHiddenAxisIntervals(
  sheet: Sheet,
  axis: WorksheetAxis,
): HiddenAxisInterval[];

/** Returns a new workbook with the default row height or column width changed, or with custom row heights or column widths reset. Sizes are in pixels. Returns the same object when nothing changes. Throws for sizes outside the allowed range, unknown keys, a read-only workbook, a stale `expectedSnapshot`, or a protected sheet. */
export function applySheetSizeEdits(
  workbook: Workbook,
  sheetId: string,
  sizeEdits: SheetSizeEdits,
  options?: SheetSizeEditOptions,
): Workbook;

/** Returns a new workbook without the outline group with the given id. Throws if no such group exists on the worksheet or the sheet is protected. */
export function removeOutlineGroup(
  workbook: Workbook,
  sheetId: string,
  outlineGroupId: string,
  options?: OutlineEditOptions,
): Workbook;

/** Returns a new workbook with the outline group collapsed or expanded. Throws if the group does not exist, `isCollapsed` is not a boolean, or the sheet is protected. */
export function setOutlineGroupCollapsed(
  workbook: Workbook,
  sheetId: string,
  outlineGroupId: string,
  isCollapsed: boolean,
  options?: OutlineEditOptions,
): Workbook;
