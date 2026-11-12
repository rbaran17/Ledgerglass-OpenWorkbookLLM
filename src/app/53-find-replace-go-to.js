import { throwWithMessage, isFormulaCellValue } from "./03-workbook-model.js";
import { eI, formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { Search, X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import A2e from "react-dom";
import {
  maxRowHeight,
  workbookLimits,
  minRowHeight,
  maxColumnWidth,
  minColumnWidth,
} from "./02-text-and-financial-functions.js";
import {
  formatNumberWithPattern,
  defaultRowHeightOf,
  columnWidthAt,
  enumerateSelectionCells,
  findSheetOrThrow,
  defaultColumnWidthOf,
  Pv,
  findOccupiedRangeBounds,
  calculateWorkbookWithSettings,
  _t,
  parseRangeAddress,
  indentToPx,
  measureCellBorderInsets,
  normalizeCellRecordInput,
  createAxisVisibilityIndex,
  columnIndexToLetters,
  vx,
  resolveEffectiveCellStyle,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  buildCellChangePreview,
  LocalizedText,
  applyCellChangePreview,
} from "./04-cell-input-paste-and-fill.js";
import {
  getFormulaCompletions,
  applyFormulaCompletion,
} from "./33-page-setup-and-print.js";
import {
  compareCellPositions,
  compileSearchPattern,
  searchWorkbookCells,
  normalizeSearchPage,
  findAdjacentSearchResult,
  normalizeSearchRequest,
  describeHiddenState,
  createSearchBudget,
  toComparableCellValue,
  comparisonLimits,
  matchTextWithBudget,
  rangeContainsCell,
} from "./52-workbook-comparison.js";

function replaceSearchMatches(
  workbook,
  searchRequest,
  replacementText,
  replaceOptions = {},
) {
  (replaceOptions.readOnly && throwWithMessage("This workbook is read-only."),
    replaceOptions.readOnly !== undefined &&
      typeof replaceOptions.readOnly != "boolean" &&
      throwWithMessage("Read-only must be boolean."),
    (typeof replacementText != "string" || replacementText.length > 32767) &&
      throwWithMessage(
        "Replacement text must contain at most 32,767 characters.",
      ),
    searchRequest.lookIn &&
      !["raw", "formulas"].includes(searchRequest.lookIn) &&
      throwWithMessage(
        "Replace searches Raw contents or Formulas only. Calculated and formatted results are read-only search modes.",
      ));
  const searchResult = searchWorkbookCells(workbook, searchRequest),
    options = searchResult.options,
    searchPattern = compileSearchPattern(options),
    searchBudget = createSearchBudget(),
    matchedResults = replaceOptions.match
      ? searchResult.results.filter(
          (result) =>
            result.sheetId === replaceOptions.match.sheetId &&
            result.address === replaceOptions.match.address.toUpperCase(),
        )
      : searchResult.results;
  replaceOptions.match &&
    !matchedResults.length &&
    throwWithMessage(
      "The chosen cell does not match this search. Find it again before replacing.",
    );
  const cellChanges = [];
  let excludedProtected = 0,
    replacedOccurrences = 0;
  for (const searchMatch of matchedResults) {
    if (searchMatch.protected) {
      excludedProtected++;
      continue;
    }
    const matchRanges = matchTextWithBudget(
      searchMatch.text,
      options,
      searchPattern,
      searchBudget,
    );
    let replacedText = "",
      copyStart = 0;
    searchMatch.text.length -
      matchRanges.reduce(
        (total, [matchStart, matchEnd]) => total + matchEnd - matchStart,
        0,
      ) +
      matchRanges.length * replacementText.length >
      32767 &&
      throwWithMessage(
        `Replacement would exceed 32,767 characters at ${searchMatch.sheetName}!${searchMatch.address}.`,
      );
    for (const [matchStart, matchEnd] of matchRanges)
      ((replacedText +=
        searchMatch.text.slice(copyStart, matchStart) + replacementText),
        (copyStart = matchEnd));
    replacedText += searchMatch.text.slice(copyStart);
    const cell = findSheetOrThrow(workbook, searchMatch.sheetId).cells[
        searchMatch.address
      ],
      comparableCellValue = toComparableCellValue(cell, replacedText);
    JSON.stringify(comparableCellValue) !==
      JSON.stringify(normalizeCellRecordInput(cell)) &&
      (cellChanges.push({
        sheetId: searchMatch.sheetId,
        address: searchMatch.address,
        cell: comparableCellValue,
      }),
      (replacedOccurrences += matchRanges.length),
      cellChanges.length > workbookLimits.maxPatches &&
        throwWithMessage(
          "Replace may change at most 500 cells, including dependent formula changes. Narrow the scope or replace one result.",
        ));
  }
  const replacePlan = {
    targetSheetId:
      cellChanges[0]?.sheetId ??
      workbook.sheets.find((sheet) => !sheet.protection?.enabled)?.id ??
      options.sheetId,
    targetRange: cellChanges[0]?.address ?? options.range,
    targetCount: cellChanges.length,
    matchedCells: matchedResults.length,
    replacedCells: cellChanges.length,
    replacedOccurrences: replacedOccurrences,
    excludedProtected: excludedProtected,
    summary: `Replace in ${options.scope}: ${cellChanges.length} cells; ${excludedProtected} protected matches excluded.`,
  };
  return cellChanges.length
    ? buildCellChangePreview(
        workbook,
        cellChanges,
        replacePlan,
        replaceOptions.readOnly,
      )
    : eI({
        ...replacePlan,
        workbookId: workbook.id,
        revision: workbook.revision,
        patches: [],
        previewPatches: [],
        count: 0,
        dependentCount: 0,
        snapshotGuard: JSON.stringify(workbook),
      });
}
function assertReplacePreviewCurrent(workbook, plan, replaceOptions = {}) {
  return plan.count
    ? applyCellChangePreview(workbook, plan, replaceOptions)
    : (replaceOptions.readOnly !== undefined &&
        typeof replaceOptions.readOnly != "boolean" &&
        throwWithMessage("Read-only must be boolean."),
      replaceOptions.readOnly &&
        throwWithMessage("This workbook is read-only."),
      (plan.workbookId !== workbook.id ||
        plan.revision !== workbook.revision ||
        plan.snapshotGuard !== JSON.stringify(workbook)) &&
        throwWithMessage(
          "The workbook changed after this preview. Create a new preview before applying it.",
        ),
      workbook);
}
function buildSelectionRanges(cellPositions) {
  const rowRanges = [];
  for (const cellPosition of [...cellPositions].sort(
    (leftPosition, rightPosition) =>
      compareCellPositions("rows", leftPosition, rightPosition),
  )) {
    const lastRange = rowRanges.at(-1);
    lastRange &&
    lastRange.firstRow === cellPosition.row &&
    lastRange.lastCol + 1 === cellPosition.col
      ? (lastRange.lastCol = cellPosition.col)
      : rowRanges.push({
          firstRow: cellPosition.row,
          lastRow: cellPosition.row,
          firstCol: cellPosition.col,
          lastCol: cellPosition.col,
        });
  }
  const mergedRanges = [],
    rangeByColumns = new Map();
  for (const rowRange of rowRanges) {
    const columnKey = `${rowRange.firstCol}:${rowRange.lastCol}`,
      previousRange = rangeByColumns.get(columnKey);
    previousRange && previousRange.lastRow + 1 === rowRange.firstRow
      ? (previousRange.lastRow = rowRange.lastRow)
      : (mergedRanges.push(rowRange), rangeByColumns.set(columnKey, rowRange));
  }
  return mergedRanges.map((rowRange) => {
    const cellAddress = formatCellAddress(rowRange.firstRow, rowRange.firstCol),
      endAddress = formatCellAddress(rowRange.lastRow, rowRange.lastCol);
    return cellAddress === endAddress
      ? cellAddress
      : `${cellAddress}:${endAddress}`;
  });
}
function selectSpecialCells(workbook, rawRequest, category) {
  const searchRequest = normalizeSearchRequest(workbook, rawRequest),
    sheetOrThrow = findSheetOrThrow(workbook, searchRequest.sheetId),
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    isHiddenCell = describeHiddenState(sheetOrThrow);
  (searchRequest.scope === "workbook" &&
    throwWithMessage(
      "Go To Special selects cells on one sheet. Choose Selection or Sheet.",
    ),
    ["formulas", "errors", "constants", "blanks", "notes"].includes(category) ||
      throwWithMessage("Choose a supported Go To Special category."));
  const occupiedRangeBounds = findOccupiedRangeBounds(sheetOrThrow),
    rangeAddresses =
      searchRequest.scope === "selection"
        ? searchRequest.ranges
        : occupiedRangeBounds
          ? [
              formatCellAddress(
                occupiedRangeBounds.firstRow,
                occupiedRangeBounds.firstCol,
              ) +
                ":" +
                formatCellAddress(
                  occupiedRangeBounds.lastRow,
                  occupiedRangeBounds.lastCol,
                ),
            ]
          : [],
    parsedRanges = rangeAddresses.map(parseRangeAddress),
    candidateAddresses = [],
    matchedPositions = [],
    finishSelection = () => {
      matchedPositions.sort((leftPosition, rightPosition) =>
        compareCellPositions(searchRequest.order, leftPosition, rightPosition),
      );
      const selectionRanges = buildSelectionRanges(matchedPositions);
      return (
        selectionRanges.length > comparisonLimits.maxSpecialAreas &&
          throwWithMessage(
            "This selection contains more than 256 separate areas. Narrow the selection.",
          ),
        eI({
          sheetId: sheetOrThrow.id,
          kind: category,
          count: matchedPositions.length,
          addresses: matchedPositions.map((position) =>
            formatCellAddress(position.row, position.col),
          ),
          ranges: selectionRanges,
          workbookId: workbook.id,
          revision: workbook.revision,
          snapshotGuard: JSON.stringify(workbook),
        })
      );
    };
  if (
    !parsedRanges.length ||
    (sheetOrThrow.hidden && !searchRequest.includeHidden)
  )
    return finishSelection();
  if (category === "blanks")
    try {
      candidateAddresses.push(
        ...enumerateSelectionCells(
          rangeAddresses,
          comparisonLimits.maxBlankScan,
        ),
      );
    } catch {
      throwWithMessage(
        "Blank selection scans at most 100,000 cells. Select a smaller rectangle.",
      );
    }
  else if (
    (candidateAddresses.push(...Object.keys(sheetOrThrow.cells)),
    category === "errors")
  )
    for (const spillEntry of workbookWithSettings.getSpillEntries(
      sheetOrThrow.id,
    ))
      Object.hasOwn(sheetOrThrow.cells, spillEntry.address) ||
        candidateAddresses.push(spillEntry.address);
  candidateAddresses.length > workbookLimits.maxCells &&
    throwWithMessage("Go To Special scans at most 100,000 cells.");
  let calculatedCellCount = 0;
  for (const candidateAddress of candidateAddresses) {
    const candidatePosition = _t(candidateAddress);
    if (
      !parsedRanges.some((parsedRange) =>
        rangeContainsCell(
          candidatePosition.row,
          candidatePosition.col,
          parsedRange,
        ),
      ) ||
      (!searchRequest.includeHidden &&
        isHiddenCell(candidatePosition.row, candidatePosition.col))
    )
      continue;
    const cell = sheetOrThrow.cells[candidateAddress],
      isFormulaCell = isFormulaCellValue(cell);
    let isSelected = false;
    if (category === "formulas") isSelected = isFormulaCell;
    else if (category === "notes") isSelected = !!cell?.note;
    else if (category === "constants")
      isSelected =
        !!cell && cell.raw !== null && cell.raw !== "" && !isFormulaCell;
    else {
      if (!isFormulaCell && cell && cell.raw !== null) continue;
      ++calculatedCellCount > comparisonLimits.maxCalculatedCells &&
        category === "errors" &&
        throwWithMessage(
          "Error selection calculates at most 10,000 cells. Narrow the selection.",
        );
      const calculatedValue = workbookWithSettings(
        sheetOrThrow.id,
        candidateAddress,
      );
      category === "blanks"
        ? (isSelected =
            (cell?.raw === null || !cell) && calculatedValue === null)
        : (isSelected =
            (isFormulaCell ||
              !!workbookWithSettings.getSpillAnchor(
                sheetOrThrow.id,
                candidateAddress,
              )) &&
            typeof calculatedValue == "string" &&
            /^#(?:BLOCKED!|CALC!|CYCLE!|DIV\/0!|ERROR!|LIMIT!|N\/A|NAME\?|NULL!|NUM!|REF!|SPILL!|VALUE!)$/.test(
              calculatedValue,
            ));
    }
    isSelected &&
      (matchedPositions.push(candidatePosition),
      matchedPositions.length > comparisonLimits.maxSpecialCells &&
        throwWithMessage(
          "Go To Special may select at most 500 cells. Narrow the selection.",
        ));
  }
  return finishSelection();
}
const describeOperationError = (error) =>
    error instanceof Error
      ? error.message
      : "This operation could not be completed.",
  formatSearchResultValue = (resultValue) =>
    resultValue == null
      ? "(blank)"
      : typeof resultValue == "boolean"
        ? resultValue
          ? "TRUE"
          : "FALSE"
        : String(resultValue),
  Bv = 20;
function FindReplaceDialog(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    firstFieldRef = ne.useRef(null),
    propsRef = ne.useRef(props);
  propsRef.current = props;
  const selectionRef = ne.useRef({
      sheetId: props.sheetId,
      range: props.range,
      ranges: props.ranges ? [...props.ranges] : [props.range],
      sheetName:
        props.workbook.sheets.find((sheet) => sheet.id === props.sheetId)
          ?.name ?? "Worksheet",
    }),
    [activeTab, setActiveTab] = ne.useState(props.initialTab ?? "find"),
    [scope, setScope] = ne.useState("sheet"),
    [lookIn, setLookIn] = ne.useState("raw"),
    [order, setOrder] = ne.useState("rows"),
    [query, setQuery] = ne.useState(""),
    [replaceText, setReplaceText] = ne.useState(""),
    [matchCase, setMatchCase] = ne.useState(false),
    [wholeCell, setWholeCell] = ne.useState(false),
    [wildcards, setWildcards] = ne.useState(true),
    [includeHidden, setIncludeHidden] = ne.useState(false),
    [specialCellKind, setSpecialCellKind] = ne.useState("formulas"),
    [searchResults, setSearchResults] = ne.useState(null),
    [match, setMatch] = ne.useState(null),
    [plan, setPlan] = ne.useState(null),
    [specialSelection, setSpecialSelection] = ne.useState(null),
    [resultPage, setResultPage] = ne.useState(0),
    [previewPage, setPreviewPage] = ne.useState(0),
    [X$1, setErrorMessage] = ne.useState(""),
    [statusMessage, setStatusMessage] = ne.useState(""),
    [originalCells, setOriginalCells] = ne.useState({}),
    getSearchOptions = () => ({
      sheetId: selectionRef.current.sheetId,
      range: selectionRef.current.range,
      ranges: [...selectionRef.current.ranges],
      scope: scope,
      query: query,
      lookIn: lookIn,
      order: order,
      matchCase: matchCase,
      wholeCell: wholeCell,
      wildcards: wildcards,
      includeHidden: includeHidden,
    });
  function resetResults() {
    (setSearchResults(null),
      setMatch(null),
      setPlan(null),
      setSpecialSelection(null),
      setResultPage(0),
      setPreviewPage(0),
      setErrorMessage(""),
      setStatusMessage(""));
  }
  function changeTab(tabId) {
    (setActiveTab(tabId),
      tabId === "replace" &&
        !["raw", "formulas"].includes(lookIn ?? "raw") &&
        setLookIn("raw"),
      tabId === "special" && scope === "workbook" && setScope("sheet"),
      resetResults());
  }
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      (
        firstFieldRef.current ??
        dialogRef.current?.querySelector("select,input") ??
        dialogRef.current?.querySelector("button")
      )?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const assertResultCurrent = (resultSet) => {
    const workbook = propsRef.current.workbook;
    if (
      resultSet.workbookId !== workbook.id ||
      resultSet.revision !== workbook.revision ||
      resultSet.snapshotGuard !== JSON.stringify(workbook)
    )
      throw new Error(
        "The workbook changed. Find or preview again before using these results.",
      );
  };
  function onClick() {
    try {
      (setErrorMessage(""), setPlan(null), setSpecialSelection(null));
      const searchResult = searchWorkbookCells(
        propsRef.current.workbook,
        getSearchOptions(),
      );
      return (
        setSearchResults(searchResult),
        setMatch(null),
        setResultPage(0),
        setStatusMessage(
          searchResult.total
            ? `${searchResult.total.toLocaleString()} matching cells.`
            : "No matches found.",
        ),
        searchResult
      );
    } catch (error) {
      return (
        setErrorMessage(describeOperationError(error)),
        setSearchResults(null),
        null
      );
    }
  }
  function goToSearchResult(searchHit, hitResults = searchResults) {
    try {
      if (!hitResults) return;
      if (
        (assertResultCurrent(hitResults),
        setErrorMessage(""),
        setMatch(searchHit),
        setResultPage(
          Math.floor(
            hitResults.results.findIndex(
              (result) =>
                result.sheetId === searchHit.sheetId &&
                result.address === searchHit.address,
            ) / Bv,
          ),
        ),
        searchHit.hidden)
      ) {
        setStatusMessage(
          `${searchHit.sheetName}!${searchHit.address} is in a ${searchHit.hiddenReason}. Unhide it before navigating; its visibility has not changed.`,
        );
        return;
      }
      (propsRef.current.onNavigate(searchHit.sheetId, searchHit.address),
        setStatusMessage(
          `${searchHit.sheetName}!${searchHit.address}${searchHit.protected ? " · protected" : ""}`,
        ));
    } catch (error) {
      setErrorMessage(describeOperationError(error));
    }
  }
  function moveToAdjacentResult(direction) {
    const currentResults = searchResults ?? onClick();
    if (currentResults)
      try {
        assertResultCurrent(currentResults);
        const rangeAddress = parseRangeAddress(selectionRef.current.range),
          searchOrigin = match ?? {
            sheetId: selectionRef.current.sheetId,
            address:
              props.active ??
              `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
          },
          adjacentSearchResult = findAdjacentSearchResult(
            currentResults,
            searchOrigin,
            direction,
          );
        adjacentSearchResult
          ? goToSearchResult(adjacentSearchResult, currentResults)
          : setStatusMessage("No matches found.");
      } catch (error) {
        setErrorMessage(describeOperationError(error));
      }
  }
  function previewReplacement(shouldReplaceSelected) {
    try {
      if ((setErrorMessage(""), shouldReplaceSelected && !match))
        throw new Error("Select a matching result before replacing one cell.");
      shouldReplaceSelected &&
        searchResults &&
        assertResultCurrent(searchResults);
      const replacePlan = replaceSearchMatches(
        propsRef.current.workbook,
        getSearchOptions(),
        replaceText,
        {
          readOnly: propsRef.current.readOnly,
          ...(shouldReplaceSelected ? { match: match } : {}),
        },
      );
      (setPlan(replacePlan),
        setOriginalCells(
          Object.fromEntries(
            replacePlan.previewPatches.map((previewPatch) => {
              const sheetOrThrow = findSheetOrThrow(
                propsRef.current.workbook,
                previewPatch.sheetId,
              );
              return [
                `${previewPatch.sheetId}:${previewPatch.address}`,
                {
                  sheetName: sheetOrThrow.name,
                  raw: sheetOrThrow.cells[previewPatch.address]?.raw ?? null,
                },
              ];
            }),
          ),
        ),
        setSpecialSelection(null),
        setPreviewPage(0),
        setStatusMessage(
          replacePlan.matchedCells
            ? `${replacePlan.matchedCells} matched; ${replacePlan.replacedCells} cells to replace; ${replacePlan.excludedProtected} protected matches excluded.`
            : "No matches found.",
        ));
    } catch (error) {
      (setPlan(null), setErrorMessage(describeOperationError(error)));
    }
  }
  function applyReplacement() {
    if (plan)
      try {
        setErrorMessage("");
        const workbook = propsRef.current.workbook,
          nextWorkbook = assertReplacePreviewCurrent(workbook, plan, {
            readOnly: propsRef.current.readOnly,
          });
        (nextWorkbook !== workbook &&
          propsRef.current.onWorkbookChange(
            nextWorkbook,
            `Replace ${plan.replacedCells} cells`,
          ),
          propsRef.current.onClose());
      } catch (error) {
        setErrorMessage(describeOperationError(error));
      }
  }
  function previewSpecialSelection() {
    try {
      setErrorMessage("");
      const specialCells = selectSpecialCells(
        propsRef.current.workbook,
        getSearchOptions(),
        specialCellKind,
      );
      (setSpecialSelection(specialCells),
        setPlan(null),
        setResultPage(0),
        setStatusMessage(
          specialCells.count
            ? `${specialCells.count} cells in ${specialCells.ranges.length} areas.`
            : "No matching cells found.",
        ));
    } catch (error) {
      (setSpecialSelection(null),
        setErrorMessage(describeOperationError(error)));
    }
  }
  function applySpecialSelection() {
    if (specialSelection?.count)
      try {
        if (
          (setErrorMessage(""),
          assertResultCurrent(specialSelection),
          !propsRef.current.onSelectRanges)
        )
          throw new Error(
            "This embedding host has not connected multiple-area selection.",
          );
        if (
          findSheetOrThrow(propsRef.current.workbook, specialSelection.sheetId)
            .hidden
        )
          throw new Error("Unhide this sheet before selecting its cells.");
        (propsRef.current.onSelectRanges(specialSelection.sheetId, [
          ...specialSelection.ranges,
        ]),
          propsRef.current.onClose());
      } catch (error) {
        setErrorMessage(describeOperationError(error));
      }
  }
  const resultPageCount = Math.max(
      1,
      Math.ceil((searchResults?.total ?? 0) / Bv),
    ),
    currentResultPage = Math.min(resultPage, resultPageCount - 1),
    previewPageCount = Math.max(
      1,
      Math.ceil((plan?.previewPatches.length ?? 0) / Bv),
    ),
    currentPreviewPage = Math.min(previewPage, previewPageCount - 1),
    isWorkbookStale = !!(
      (searchResults &&
        (searchResults.workbookId !== props.workbook.id ||
          searchResults.revision !== props.workbook.revision)) ||
      (plan &&
        (plan.workbookId !== props.workbook.id ||
          plan.revision !== props.workbook.revision)) ||
      (specialSelection &&
        (specialSelection.workbookId !== props.workbook.id ||
          specialSelection.revision !== props.workbook.revision))
    );
  return i.jsx("div", {
    className: "st-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      className: "st-dialog",
      ref: dialogRef,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      onKeyDown: (event) => {
        if (event.key === "Escape")
          (event.preventDefault(),
            event.stopPropagation(),
            propsRef.current.onClose());
        else if (event.key === "Tab") {
          const focusables = Array.from(
              dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ).filter((focusable) => !focusable.hidden),
            firstFocusable = focusables[0],
            lastFocusable = focusables.at(-1);
          ((event.shiftKey && document.activeElement === firstFocusable) ||
            (!event.shiftKey && document.activeElement === lastFocusable)) &&
            (event.preventDefault(),
            (event.shiftKey ? lastFocusable : firstFocusable)?.focus());
        } else
          event.key === "Enter" &&
            event.target instanceof HTMLInputElement &&
            event.target.type !== "checkbox" &&
            (event.preventDefault(),
            moveToAdjacentResult(event.shiftKey ? "previous" : "next"));
      },
      children: [
        i.jsxs("header", {
          className: "st-header",
          children: [
            i.jsxs("h2", {
              id: titleId,
              children: [
                i.jsx(Search, { size: 18 }),
                i.jsx(LocalizedText, { children: "Find and Replace" }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": text("Close Find and Replace"),
              onClick: props.onClose,
              lang: language("Close Find and Replace"),
              "data-ui-label": "Close Find and Replace",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("div", {
          className: "st-tabs",
          role: "tablist",
          "aria-label": text("Search tools"),
          lang: language("Search tools"),
          "data-ui-label": "Search tools",
          children: [
            ["find", "Find"],
            ["replace", "Replace"],
            ["special", "Go To Special"],
          ].map(([tabKey, tabLabel], tabIndex) =>
            i.jsx(
              "button",
              {
                type: "button",
                role: "tab",
                "aria-selected": activeTab === tabKey,
                tabIndex: activeTab === tabKey ? 0 : -1,
                onClick: () => changeTab(tabKey),
                onKeyDown: (event) => {
                  if (
                    ["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                      event.key,
                    )
                  ) {
                    event.preventDefault();
                    const tabIds = ["find", "replace", "special"],
                      targetIndex =
                        event.key === "Home"
                          ? 0
                          : event.key === "End"
                            ? 2
                            : (tabIndex +
                                (event.key === "ArrowRight" ? 1 : 2)) %
                              3;
                    (changeTab(tabIds[targetIndex]),
                      event.currentTarget.parentElement?.children[
                        targetIndex
                      ]?.focus());
                  }
                },
                children: tabLabel,
              },
              tabKey,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "st-body",
          children: [
            i.jsxs("p", {
              className: "st-context",
              children: [
                i.jsx(LocalizedText, { children: "Selection:" }),
                " ",
                i.jsxs("strong", {
                  children: [
                    selectionRef.current.sheetName,
                    "!",
                    selectionRef.current.ranges.join(", "),
                  ],
                }),
              ],
            }),
            i.jsxs("div", {
              className: "st-options",
              children: [
                activeTab !== "special"
                  ? i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("label", {
                          className: "st-wide",
                          children: [
                            i.jsx(LocalizedText, { children: "Find what" }),
                            i.jsx("input", {
                              ref: firstFieldRef,
                              value: query,
                              maxLength: 256,
                              onChange: (event) => {
                                (setQuery(event.target.value), resetResults());
                              },
                            }),
                          ],
                        }),
                        activeTab === "replace" &&
                          i.jsxs("label", {
                            className: "st-wide",
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Replace with",
                              }),
                              i.jsx("input", {
                                value: replaceText,
                                maxLength: 32767,
                                onChange: (event) => {
                                  (setReplaceText(event.target.value),
                                    setPlan(null),
                                    setErrorMessage(""));
                                },
                              }),
                            ],
                          }),
                      ],
                    })
                  : i.jsxs("label", {
                      className: "st-wide",
                      children: [
                        i.jsx(LocalizedText, {
                          children: "Select cells containing",
                        }),
                        i.jsxs("select", {
                          value: specialCellKind,
                          onChange: (event) => {
                            (setSpecialCellKind(event.target.value),
                              resetResults());
                          },
                          children: [
                            i.jsx("option", {
                              value: "formulas",
                              children: "Formulas",
                            }),
                            i.jsx("option", {
                              value: "errors",
                              children: "Formula errors",
                            }),
                            i.jsx("option", {
                              value: "constants",
                              children: "Constants",
                            }),
                            i.jsx("option", {
                              value: "blanks",
                              children: "Blanks",
                            }),
                            i.jsx("option", {
                              value: "notes",
                              children: "Notes",
                            }),
                          ],
                        }),
                      ],
                    }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Within" }),
                    i.jsxs("select", {
                      value: scope,
                      onChange: (event) => {
                        (setScope(event.target.value), resetResults());
                      },
                      children: [
                        i.jsx("option", {
                          value: "selection",
                          children: "Selection",
                        }),
                        i.jsx("option", { value: "sheet", children: "Sheet" }),
                        activeTab !== "special" &&
                          i.jsx("option", {
                            value: "workbook",
                            children: "Workbook",
                          }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Search order" }),
                    i.jsxs("select", {
                      value: order,
                      onChange: (event) => {
                        (setOrder(event.target.value), resetResults());
                      },
                      children: [
                        i.jsx("option", { value: "rows", children: "By rows" }),
                        i.jsx("option", {
                          value: "columns",
                          children: "By columns",
                        }),
                      ],
                    }),
                  ],
                }),
                activeTab !== "special" &&
                  i.jsxs("label", {
                    className: "st-wide",
                    children: [
                      i.jsx(LocalizedText, { children: "Look in" }),
                      i.jsxs("select", {
                        value: lookIn,
                        onChange: (event) => {
                          (setLookIn(event.target.value), resetResults());
                        },
                        children: [
                          i.jsx("option", {
                            value: "raw",
                            children:
                              "Raw contents (text, numbers and formulas)",
                          }),
                          i.jsx("option", {
                            value: "formulas",
                            children: "Formulas only",
                          }),
                          activeTab === "find" &&
                            i.jsxs(i.Fragment, {
                              children: [
                                i.jsx("option", {
                                  value: "calculated",
                                  children: "Calculated values",
                                }),
                                i.jsx("option", {
                                  value: "formatted",
                                  children: "Formatted display",
                                }),
                              ],
                            }),
                        ],
                      }),
                    ],
                  }),
              ],
            }),
            i.jsxs("div", {
              className: "st-switches",
              children: [
                activeTab !== "special" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: matchCase,
                            onChange: (event) => {
                              (setMatchCase(event.target.checked),
                                resetResults());
                            },
                          }),
                          i.jsx(LocalizedText, { children: "Match case" }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: wholeCell,
                            onChange: (event) => {
                              (setWholeCell(event.target.checked),
                                resetResults());
                            },
                          }),
                          i.jsx(LocalizedText, {
                            children: "Match entire cell contents",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: wildcards,
                            onChange: (event) => {
                              (setWildcards(event.target.checked),
                                resetResults());
                            },
                          }),
                          i.jsx(LocalizedText, {
                            children: "Use wildcards (* ? ~)",
                          }),
                        ],
                      }),
                    ],
                  }),
                i.jsxs("label", {
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: includeHidden,
                      onChange: (event) => {
                        (setIncludeHidden(event.target.checked),
                          resetResults());
                      },
                    }),
                    i.jsx(LocalizedText, {
                      children: "Include hidden cells and sheets",
                    }),
                  ],
                }),
              ],
            }),
            i.jsx("p", {
              className: "st-help",
              children:
                activeTab === "special"
                  ? "Select up to 500 cells in 256 separate areas. Sheet scope uses the occupied rectangle for blanks. Formula results of empty text are not blank cells."
                  : "? matches one character; * matches any sequence. Prefix *, ? or ~ with ~ to find it literally. Search follows workbook sheet order.",
            }),
            props.readOnly &&
              i.jsx("p", {
                className: "st-notice",
                children: i.jsx(LocalizedText, {
                  children:
                    "This workbook is read-only. Finding and selecting cells remain available.",
                }),
              }),
            isWorkbookStale &&
              i.jsx("p", {
                className: "st-notice",
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "The workbook changed. Find or preview again before continuing.",
                }),
              }),
            i.jsx("div", {
              className: "st-actions",
              children:
                activeTab === "special"
                  ? i.jsx("button", {
                      type: "button",
                      onClick: previewSpecialSelection,
                      children: i.jsx(LocalizedText, {
                        children: "Preview selection",
                      }),
                    })
                  : i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("button", {
                          type: "button",
                          onClick: onClick,
                          children: i.jsx(LocalizedText, {
                            children: "Find all",
                          }),
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => moveToAdjacentResult("previous"),
                          children: i.jsx(LocalizedText, {
                            children: "Find previous",
                          }),
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => moveToAdjacentResult("next"),
                          children: i.jsx(LocalizedText, {
                            children: "Find next",
                          }),
                        }),
                        activeTab === "replace" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("button", {
                                type: "button",
                                disabled: props.readOnly || !match,
                                onClick: () => previewReplacement(true),
                                children: i.jsx(LocalizedText, {
                                  children: "Preview replace selected",
                                }),
                              }),
                              i.jsx("button", {
                                type: "button",
                                disabled: props.readOnly,
                                onClick: () => previewReplacement(false),
                                children: i.jsx(LocalizedText, {
                                  children: "Preview replace all",
                                }),
                              }),
                            ],
                          }),
                      ],
                    }),
            }),
            X$1 &&
              i.jsx("p", {
                className: "st-error",
                role: "alert",
                children: X$1,
              }),
            statusMessage &&
              i.jsx("p", {
                className: "st-status",
                role: "status",
                children: statusMessage,
              }),
            searchResults &&
              i.jsxs("div", {
                className: "st-results",
                "aria-label": text("Find results"),
                lang: language("Find results"),
                "data-ui-label": "Find results",
                children: [
                  i.jsx("div", {
                    className: "st-table-scroll",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Sheet",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Cell",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Matching contents",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "State",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: normalizeSearchPage(
                            searchResults,
                            currentResultPage,
                            Bv,
                          ).map((resultEntry) =>
                            i.jsxs(
                              "tr",
                              {
                                className:
                                  match?.sheetId === resultEntry.sheetId &&
                                  match.address === resultEntry.address
                                    ? "st-selected"
                                    : "",
                                children: [
                                  i.jsx("td", {
                                    children: resultEntry.sheetName,
                                  }),
                                  i.jsx("td", {
                                    children: i.jsx("button", {
                                      type: "button",
                                      "aria-label": `Go to ${resultEntry.sheetName}!${resultEntry.address}`,
                                      onClick: () =>
                                        goToSearchResult(resultEntry),
                                      children: resultEntry.address,
                                    }),
                                  }),
                                  i.jsx("td", {
                                    className: "st-value",
                                    title: resultEntry.text,
                                    children: resultEntry.text,
                                  }),
                                  i.jsx("td", {
                                    children:
                                      [
                                        resultEntry.protected
                                          ? "Protected"
                                          : "",
                                        resultEntry.hidden
                                          ? resultEntry.hiddenReason
                                          : "",
                                      ]
                                        .filter(Boolean)
                                        .join(", ") || "Visible",
                                  }),
                                ],
                              },
                              `${resultEntry.sheetId}:${resultEntry.address}`,
                            ),
                          ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "st-pagination",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        "aria-label": text("Previous results page"),
                        disabled: !currentResultPage,
                        onClick: () => setResultPage(currentResultPage - 1),
                        lang: language("Previous results page"),
                        "data-ui-label": "Previous results page",
                        children: i.jsx(LocalizedText, {
                          children: "Previous",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Results page " }),
                          currentResultPage + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          resultPageCount,
                        ],
                      }),
                      i.jsx("button", {
                        type: "button",
                        "aria-label": text("Next results page"),
                        disabled: currentResultPage + 1 >= resultPageCount,
                        onClick: () => setResultPage(currentResultPage + 1),
                        lang: language("Next results page"),
                        "data-ui-label": "Next results page",
                        children: i.jsx(LocalizedText, { children: "Next" }),
                      }),
                    ],
                  }),
                ],
              }),
            plan &&
              i.jsxs("div", {
                className: "st-preview",
                "aria-label": text("Replacement preview"),
                lang: language("Replacement preview"),
                "data-ui-label": "Replacement preview",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Replacement preview",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      plan.replacedCells,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "directly replaced cells · ",
                      }),
                      plan.count,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "total changed cells · ",
                      }),
                      plan.excludedProtected,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "protected matches excluded.",
                      }),
                    ],
                  }),
                  plan.dependentCount > 0 &&
                    i.jsxs("p", {
                      className: "st-notice",
                      children: [
                        i.jsx(LocalizedText, { children: "Includes " }),
                        plan.dependentCount,
                        " ",
                        i.jsx(LocalizedText, {
                          children:
                            "dependent cell changes outside the direct replacement targets. Every change is listed below.",
                        }),
                      ],
                    }),
                  i.jsx("div", {
                    className: "st-table-scroll",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Sheet / cell",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Before",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "After",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: plan.previewPatches
                            .slice(
                              currentPreviewPage * Bv,
                              (currentPreviewPage + 1) * Bv,
                            )
                            .map((patch) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsxs("td", {
                                      children: [
                                        originalCells[
                                          `${patch.sheetId}:${patch.address}`
                                        ]?.sheetName ?? patch.sheetId,
                                        "!",
                                        patch.address,
                                      ],
                                    }),
                                    i.jsx("td", {
                                      className: "st-value",
                                      children: formatSearchResultValue(
                                        originalCells[
                                          `${patch.sheetId}:${patch.address}`
                                        ]?.raw,
                                      ),
                                    }),
                                    i.jsx("td", {
                                      className: "st-value",
                                      children: formatSearchResultValue(
                                        patch.cell.raw,
                                      ),
                                    }),
                                  ],
                                },
                                `${patch.sheetId}:${patch.address}`,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "st-pagination",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        "aria-label": text("Previous preview page"),
                        disabled: !currentPreviewPage,
                        onClick: () => setPreviewPage(currentPreviewPage - 1),
                        lang: language("Previous preview page"),
                        "data-ui-label": "Previous preview page",
                        children: i.jsx(LocalizedText, {
                          children: "Previous",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Preview page " }),
                          currentPreviewPage + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          previewPageCount,
                        ],
                      }),
                      i.jsx("button", {
                        type: "button",
                        "aria-label": text("Next preview page"),
                        disabled: currentPreviewPage + 1 >= previewPageCount,
                        onClick: () => setPreviewPage(currentPreviewPage + 1),
                        lang: language("Next preview page"),
                        "data-ui-label": "Next preview page",
                        children: i.jsx(LocalizedText, { children: "Next" }),
                      }),
                    ],
                  }),
                ],
              }),
            specialSelection &&
              i.jsxs("div", {
                className: "st-preview",
                "aria-label": text("Selection preview"),
                lang: language("Selection preview"),
                "data-ui-label": "Selection preview",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Selection preview",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      specialSelection.count,
                      " ",
                      i.jsx(LocalizedText, { children: "cells in " }),
                      specialSelection.ranges.length,
                      " ",
                      i.jsx(LocalizedText, { children: "areas" }),
                    ],
                  }),
                  i.jsx("div", {
                    className: "st-ranges",
                    children: specialSelection.ranges.map((range) =>
                      i.jsx("code", { children: range }, range),
                    ),
                  }),
                  !props.onSelectRanges &&
                    i.jsx("p", {
                      className: "st-notice",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Multiple-area selection is not connected by this embedding host.",
                      }),
                    }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "st-footer",
          children: [
            i.jsx("button", {
              type: "button",
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Close" }),
            }),
            activeTab === "replace" &&
              i.jsx("button", {
                type: "button",
                className: "st-primary",
                disabled: props.readOnly || !plan?.count || isWorkbookStale,
                onClick: applyReplacement,
                children: i.jsx(LocalizedText, {
                  children: "Apply replacement",
                }),
              }),
            activeTab === "special" &&
              i.jsx("button", {
                type: "button",
                className: "st-primary",
                disabled:
                  !specialSelection?.count ||
                  !props.onSelectRanges ||
                  isWorkbookStale,
                onClick: applySpecialSelection,
                children: i.jsx(LocalizedText, { children: "Select cells" }),
              }),
          ],
        }),
      ],
    }),
  });
}
function resolveCellNavigation(
  currentIndex,
  cellCount,
  visibleIndexes,
  direction,
) {
  if (
    !Number.isInteger(currentIndex) ||
    !Number.isInteger(cellCount) ||
    cellCount < 1 ||
    currentIndex < 0 ||
    currentIndex >= cellCount
  )
    throw new Error("Navigation requires a valid visible cell.");
  if (direction !== 1 && direction !== -1)
    throw new Error("Choose a navigation direction.");
  const visibleSet = new Set(
      visibleIndexes.filter(
        (visibleIndex) =>
          Number.isInteger(visibleIndex) &&
          visibleIndex >= 0 &&
          visibleIndex < cellCount,
      ),
    ),
    nextIndex = currentIndex + direction;
  if (nextIndex < 0 || nextIndex >= cellCount) return currentIndex;
  if (visibleSet.has(currentIndex) && visibleSet.has(nextIndex)) {
    let runEnd = nextIndex;
    for (; visibleSet.has(runEnd + direction);) runEnd += direction;
    return runEnd;
  }
  let nearestIndex = direction === 1 ? cellCount - 1 : 0;
  for (const candidateIndex of visibleSet)
    (direction === 1
      ? candidateIndex > currentIndex && candidateIndex < nearestIndex
      : candidateIndex < currentIndex && candidateIndex > nearestIndex) &&
      (nearestIndex = candidateIndex);
  return nearestIndex;
}
const paneQuadrantNames = [
    "top-left",
    "top-right",
    "bottom-left",
    "bottom-right",
  ],
  createDefaultPaneRects = () =>
    Object.fromEntries(
      paneQuadrantNames.map((paneQuadrantName) => [
        paneQuadrantName,
        { top: 0, left: 0, width: 1e3, height: 600 },
      ]),
    );
function useSplitPaneSizes(splitPaneSizesKey) {
  const [views, setViews] = ne.useState(createDefaultPaneRects),
    viewsRef = ne.useRef(views),
    refs = ne.useRef({}),
    [activeId, setActiveId] = ne.useState("top-left"),
    active = ne.useRef("top-left");
  function patch(rectUpdates) {
    let nextViews = viewsRef.current;
    for (const paneQuadrantName of paneQuadrantNames)
      if (rectUpdates[paneQuadrantName]) {
        const mergedRect = {
          ...nextViews[paneQuadrantName],
          ...rectUpdates[paneQuadrantName],
        };
        Object.entries(mergedRect).some(
          ([rectKey, rectValue]) =>
            nextViews[paneQuadrantName][rectKey] !== rectValue,
        ) && (nextViews = { ...nextViews, [paneQuadrantName]: mergedRect });
      }
    nextViews !== viewsRef.current &&
      ((viewsRef.current = nextViews), setViews(nextViews));
  }
  const bindings = ne.useMemo(
    () =>
      Object.fromEntries(
        paneQuadrantNames.map((paneQuadrantName) => [
          paneQuadrantName,
          (paneElement) => {
            paneElement
              ? (refs.current[paneQuadrantName] = paneElement)
              : delete refs.current[paneQuadrantName];
          },
        ]),
      ),
    [],
  );
  ne.useEffect(() => {
    const syncPaneSizes = () => {
      const measuredSizes = {};
      for (const paneQuadrantName of paneQuadrantNames) {
        const paneNode = refs.current[paneQuadrantName];
        paneNode?.clientWidth &&
          paneNode.clientHeight &&
          (measuredSizes[paneQuadrantName] = {
            width: paneNode.clientWidth,
            height: paneNode.clientHeight,
          });
      }
      patch(measuredSizes);
    };
    if ((syncPaneSizes(), typeof ResizeObserver > "u")) return;
    const resizeObserver = new ResizeObserver(syncPaneSizes);
    return (
      Object.values(refs.current).forEach((observedElement) =>
        resizeObserver.observe(observedElement),
      ),
      () => resizeObserver.disconnect()
    );
  }, [splitPaneSizesKey]);
  function activate(paneId) {
    ((active.current = paneId), setActiveId(paneId));
  }
  return {
    views: views,
    current: viewsRef,
    refs: refs,
    bindings: bindings,
    activeId: activeId,
    active: active,
    activate: activate,
    patch: patch,
  };
}
function SplitPaneLayout(props) {
  const { text, language } = useWorkbookMessages(),
    layoutRef = ne.useRef(null),
    latestProps = ne.useRef(props);
  latestProps.current = props;
  const [dragState, setDragState] = ne.useState(null),
    dragStateRef = ne.useRef(dragState);
  dragStateRef.current = dragState;
  const [errorMessage, setErrorMessage] = ne.useState(""),
    sheetSplitPanes =
      dragState && dragState.snapshot === props.snapshot && !props.readOnly
        ? { ...props.split, [dragState.axis]: dragState.value }
        : props.split,
    resizeSplit = (axis, percent) => {
      latestProps.current.readOnly ||
        !latestProps.current.split ||
        props.onResize({
          ...latestProps.current.split,
          [axis]: Math.max(20, Math.min(80, percent)),
        });
    };
  return i.jsxs("div", {
    ref: layoutRef,
    className: `lg-pane-layout ${sheetSplitPanes ? "is-split" : ""}`,
    style: {
      gridTemplateColumns:
        sheetSplitPanes?.vertical !== undefined
          ? `minmax(0,${sheetSplitPanes.vertical}fr) 7px minmax(0,${100 - sheetSplitPanes.vertical}fr)`
          : "minmax(0,1fr)",
      gridTemplateRows:
        sheetSplitPanes?.horizontal !== undefined
          ? `minmax(0,${sheetSplitPanes.horizontal}fr) 7px minmax(0,${100 - sheetSplitPanes.horizontal}fr)`
          : "minmax(0,1fr)",
    },
    children: [
      Pv(sheetSplitPanes).map((pane) =>
        i.jsx(
          "div",
          {
            className: `lg-pane ${props.active === pane ? "is-active" : ""}`,
            "data-pane": pane,
            style: {
              gridRow: pane.startsWith("bottom") ? 3 : 1,
              gridColumn: pane.endsWith("right") ? 3 : 1,
            },
            children: props.renderPane(pane),
          },
          pane,
        ),
      ),
      ["horizontal", "vertical"].map((axis) =>
        sheetSplitPanes?.[axis] === undefined
          ? null
          : i.jsx(
              "div",
              {
                className: `lg-split-divider lg-split-${axis}`,
                role: "separator",
                tabIndex: props.readOnly ? -1 : 0,
                "aria-label": `Resize ${axis} split`,
                "aria-orientation": axis,
                "aria-valuenow": sheetSplitPanes[axis],
                "aria-valuemin": 20,
                "aria-valuemax": 80,
                "aria-disabled": props.readOnly,
                title: text(
                  "Drag or use arrow keys to resize; double-click to center",
                ),
                style:
                  axis === "horizontal"
                    ? { gridRow: 2, gridColumn: "1 / -1" }
                    : { gridColumn: 2, gridRow: "1 / -1" },
                onDoubleClick: () => resizeSplit(axis, 50),
                onKeyDown: (event) => {
                  if (event.key === "Escape") {
                    ((dragStateRef.current = null), setDragState(null));
                    return;
                  }
                  const directionIndex = (
                    axis === "horizontal"
                      ? ["ArrowUp", "ArrowDown"]
                      : ["ArrowLeft", "ArrowRight"]
                  ).indexOf(event.key);
                  directionIndex >= 0
                    ? (event.preventDefault(),
                      resizeSplit(
                        axis,
                        sheetSplitPanes[axis] +
                          (directionIndex ? 1 : -1) * (event.shiftKey ? 10 : 1),
                      ))
                    : (event.key === "Home" || event.key === "End") &&
                      (event.preventDefault(),
                      resizeSplit(axis, event.key === "Home" ? 20 : 80));
                },
                onPointerDown: (event) => {
                  var pointerTarget;
                  if (props.readOnly || event.button !== 0) return;
                  (event.preventDefault(),
                    event.currentTarget.focus(),
                    setErrorMessage(""));
                  const dragSession = {
                    axis: axis,
                    value: sheetSplitPanes[axis],
                    snapshot: props.snapshot,
                    snapshotGuard: JSON.stringify(props.snapshot),
                    split: props.split,
                  };
                  ((dragStateRef.current = dragSession),
                    setDragState(dragSession),
                    (pointerTarget =
                      event.currentTarget).setPointerCapture?.call(
                      pointerTarget,
                      event.pointerId,
                    ));
                },
                onPointerMove: (event) => {
                  const activeDrag = dragStateRef.current,
                    layoutRect = layoutRef.current?.getBoundingClientRect();
                  if (!activeDrag || !layoutRect) return;
                  const axisLength =
                    axis === "horizontal"
                      ? layoutRect.height
                      : layoutRect.width;
                  if (axisLength <= 0) return;
                  const pointerOffset =
                      axis === "horizontal"
                        ? event.clientY - layoutRect.top
                        : event.clientX - layoutRect.left,
                    nextDrag = {
                      ...activeDrag,
                      value: Math.max(
                        20,
                        Math.min(
                          80,
                          Math.round((pointerOffset / axisLength) * 100),
                        ),
                      ),
                    };
                  ((dragStateRef.current = nextDrag), setDragState(nextDrag));
                },
                onPointerUp: (event) => {
                  var pointerTarget;
                  const activeDrag = dragStateRef.current;
                  if (
                    ((dragStateRef.current = null),
                    setDragState(null),
                    !!activeDrag)
                  ) {
                    if (
                      activeDrag.snapshot !== latestProps.current.snapshot ||
                      activeDrag.snapshotGuard !==
                        JSON.stringify(latestProps.current.snapshot) ||
                      latestProps.current.readOnly
                    ) {
                      setErrorMessage(
                        "The workbook changed. Resize the split again.",
                      );
                      return;
                    }
                    (activeDrag.value !== activeDrag.split[axis] &&
                      resizeSplit(axis, activeDrag.value),
                      (pointerTarget =
                        event.currentTarget).hasPointerCapture?.call(
                        pointerTarget,
                        event.pointerId,
                      ) &&
                        event.currentTarget.releasePointerCapture(
                          event.pointerId,
                        ));
                  }
                },
                onPointerCancel: () => {
                  ((dragStateRef.current = null), setDragState(null));
                },
                onLostPointerCapture: () => {
                  ((dragStateRef.current = null), setDragState(null));
                },
                lang: language(
                  "Drag or use arrow keys to resize; double-click to center",
                ),
              },
              axis,
            ),
      ),
      errorMessage &&
        i.jsx("span", {
          role: "alert",
          className: "lg-split-error",
          children: errorMessage,
        }),
    ],
  });
}

function isWithinFormulaSuggestions(currentTarget, relatedTarget) {
  if (!(relatedTarget instanceof Node)) return false;
  const controlledElement = document.getElementById(
    currentTarget.getAttribute("aria-controls") ?? "",
  );
  return !!controlledElement
    ?.closest(".formula-suggestions")
    ?.contains(relatedTarget);
}
const FormulaSuggestionList = ne.forwardRef(function (
    {
      value,
      caret = value.length,
      workbook,
      sheetId,
      inputRef,
      onAccept,
      disabled,
      onExit,
    },
    forwardedRef,
  ) {
    const listId = ne.useId(),
      popupRef = ne.useRef(null),
      completions = ne.useMemo(
        () => getFormulaCompletions(value, caret, workbook, sheetId),
        [value, caret, workbook, sheetId],
      ),
      [activeIndex, setActiveIndex] = ne.useState(-1),
      [dismissedKey, setDismissedKey] = ne.useState(null),
      [isFocusWithin, setIsFocusWithin] = ne.useState(false),
      [popupPosition, setPopupPosition] = ne.useState({
        left: 0,
        top: 0,
        width: 280,
      }),
      valueKey = `${value}:${caret}`,
      isSuggestionsVisible =
        !disabled &&
        isFocusWithin &&
        dismissedKey !== valueKey &&
        (completions.items.length > 0 || completions.hint);
    (ne.useEffect(() => {
      const inputElement = inputRef.current;
      if (!inputElement) return;
      const ariaAttributeNames = [
          "aria-autocomplete",
          "aria-controls",
          "aria-activedescendant",
        ],
        previousAriaValues = ariaAttributeNames.map((attributeName) =>
          inputElement.getAttribute(attributeName),
        );
      return (
        isSuggestionsVisible && completions.items.length
          ? (inputElement.setAttribute("aria-autocomplete", "list"),
            inputElement.setAttribute("aria-controls", `${listId}-list`),
            activeIndex >= 0 && completions.items[activeIndex]
              ? inputElement.setAttribute(
                  "aria-activedescendant",
                  `${listId}-${activeIndex}`,
                )
              : inputElement.removeAttribute("aria-activedescendant"))
          : (inputElement.removeAttribute("aria-controls"),
            inputElement.removeAttribute("aria-activedescendant")),
        () =>
          ariaAttributeNames.forEach((attributeName, attributeIndex) =>
            previousAriaValues[attributeIndex] === null
              ? inputElement.removeAttribute(attributeName)
              : inputElement.setAttribute(
                  attributeName,
                  previousAriaValues[attributeIndex],
                ),
          )
      );
    }, [
      isSuggestionsVisible,
      activeIndex,
      completions.items.length,
      inputRef,
      listId,
    ]),
      ne.useEffect(() => setActiveIndex(-1), [valueKey]),
      ne.useEffect(() => {
        const inputElement = inputRef.current;
        if (!inputElement) return;
        const updatePopup = () => {
            setIsFocusWithin(
              document.activeElement === inputElement ||
                !!popupRef.current?.contains(document.activeElement),
            );
            const inputRect = inputElement.getBoundingClientRect();
            setPopupPosition({
              left: Math.max(
                8,
                Math.min(inputRect.left, window.innerWidth - 350),
              ),
              top: inputRect.bottom + 3,
              width: Math.max(280, Math.min(400, inputRect.width)),
            });
          },
          handleBlur = (event) => {
            popupRef.current?.contains(event.relatedTarget) ||
              setIsFocusWithin(false);
          };
        return (
          updatePopup(),
          inputElement.addEventListener("focus", updatePopup),
          inputElement.addEventListener("blur", handleBlur),
          window.addEventListener("resize", updatePopup),
          window.addEventListener("scroll", updatePopup, true),
          () => {
            (inputElement.removeEventListener("focus", updatePopup),
              inputElement.removeEventListener("blur", handleBlur),
              window.removeEventListener("resize", updatePopup),
              window.removeEventListener("scroll", updatePopup, true));
          }
        );
      }, [inputRef]),
      ne.useEffect(() => {
        if (!isSuggestionsVisible || activeIndex < 0) return;
        const activeOption = document.getElementById(
            `${listId}-${activeIndex}`,
          ),
          popupElement = popupRef.current;
        if (!activeOption || !popupElement) return;
        const optionRect = activeOption.getBoundingClientRect(),
          popupRect = popupElement.getBoundingClientRect();
        optionRect.bottom > popupRect.bottom
          ? (popupElement.scrollTop += optionRect.bottom - popupRect.bottom)
          : optionRect.top < popupRect.top &&
            (popupElement.scrollTop -= popupRect.top - optionRect.top);
      }, [isSuggestionsVisible, activeIndex, listId]));
    const acceptCompletionAt = (itemIndex) => {
      const completionItem = completions.items[itemIndex];
      if (!completionItem) return;
      inputRef.current?.focus({ preventScroll: true });
      const completionResult = applyFormulaCompletion(value, completionItem);
      (onAccept(completionResult.value, completionResult.caret),
        setActiveIndex(-1),
        setDismissedKey(null),
        requestAnimationFrame(() => {
          (inputRef.current?.focus(),
            inputRef.current?.setSelectionRange(
              completionResult.caret,
              completionResult.caret,
            ));
        }));
    };
    return (
      ne.useImperativeHandle(forwardedRef, () => ({
        handleKeyDown(event) {
          return event.nativeEvent.isComposing ||
            event.defaultPrevented ||
            !isSuggestionsVisible
            ? false
            : (event.key === "ArrowDown" || event.key === "ArrowUp") &&
                completions.items.length
              ? (event.preventDefault(),
                setActiveIndex((previousIndex) =>
                  event.key === "ArrowDown"
                    ? (previousIndex + 1) % completions.items.length
                    : previousIndex <= 0
                      ? completions.items.length - 1
                      : previousIndex - 1,
                ),
                true)
              : event.key === "Tab" &&
                  !event.shiftKey &&
                  activeIndex < 0 &&
                  completions.items.length &&
                  /[A-Za-z_][A-Za-z0-9_.]*(?:\[[^\]]*)?$/.test(
                    value.slice(0, caret),
                  )
                ? (event.preventDefault(), acceptCompletionAt(0), true)
                : (event.key === "Tab" || event.key === "Enter") &&
                    activeIndex >= 0 &&
                    completions.items[activeIndex]
                  ? (event.preventDefault(),
                    acceptCompletionAt(activeIndex),
                    true)
                  : event.key === "Escape"
                    ? (event.preventDefault(), setDismissedKey(valueKey), true)
                    : false;
        },
      })),
      !isSuggestionsVisible || typeof document > "u"
        ? null
        : A2e.createPortal(
            i.jsxs("div", {
              className: "formula-suggestions",
              ref: popupRef,
              role: "region",
              "aria-label": "Formula assistance",
              onBlur: (event) => {
                event.currentTarget.contains(event.relatedTarget) ||
                  event.relatedTarget === inputRef.current ||
                  (setIsFocusWithin(false), onExit?.());
              },
              style: { position: "fixed", ...popupPosition },
              children: [
                completions.items.length > 0 &&
                  i.jsx("ul", {
                    id: `${listId}-list`,
                    role: "listbox",
                    "aria-label": "Formula suggestions",
                    tabIndex: 0,
                    "aria-activedescendant":
                      activeIndex >= 0 && completions.items[activeIndex]
                        ? `${listId}-${activeIndex}`
                        : undefined,
                    onFocus: () => {
                      (setIsFocusWithin(true),
                        activeIndex < 0 && setActiveIndex(0));
                    },
                    onKeyDown: (event) => {
                      var inputElement;
                      event.nativeEvent.isComposing ||
                        (event.key === "ArrowDown" || event.key === "ArrowUp"
                          ? (event.preventDefault(),
                            setActiveIndex((previousIndex) =>
                              event.key === "ArrowDown"
                                ? (previousIndex + 1) % completions.items.length
                                : (previousIndex <= 0
                                    ? completions.items.length
                                    : previousIndex) - 1,
                            ))
                          : event.key === "Enter" ||
                              (event.key === "Tab" && !event.shiftKey)
                            ? (event.preventDefault(),
                              acceptCompletionAt(Math.max(0, activeIndex)))
                            : (event.key === "Escape" || event.key === "Tab") &&
                              (event.preventDefault(),
                              event.key === "Escape" &&
                                setDismissedKey(valueKey),
                              (inputElement = inputRef.current) == null ||
                                inputElement.focus()));
                    },
                    children: completions.items.map((item, optionIndex) =>
                      i.jsxs(
                        "li",
                        {
                          id: `${listId}-${optionIndex}`,
                          role: "option",
                          "aria-selected": activeIndex === optionIndex,
                          onClick: (event) => {
                            event.detail === 0 &&
                              acceptCompletionAt(optionIndex);
                          },
                          onMouseDown: (event) => {
                            (event.preventDefault(),
                              acceptCompletionAt(optionIndex));
                          },
                          onMouseEnter: () => setActiveIndex(optionIndex),
                          children: [
                            i.jsx("span", {
                              className: "formula-suggestion-kind",
                              children:
                                item.kind === "function"
                                  ? "ƒx"
                                  : item.kind === "table"
                                    ? "▦"
                                    : "N",
                            }),
                            i.jsxs("div", {
                              children: [
                                i.jsx("strong", { children: item.label }),
                                i.jsx("small", { children: item.detail }),
                              ],
                            }),
                          ],
                        },
                        `${item.kind}:${item.label}`,
                      ),
                    ),
                  }),
                completions.hint &&
                  i.jsxs("div", {
                    className: "formula-argument-hint",
                    role: "status",
                    "aria-live": "polite",
                    "aria-atomic": "true",
                    children: [
                      i.jsx("strong", { children: completions.hint.name }),
                      "(",
                      completions.hint.signature,
                      ")",
                      i.jsxs("small", {
                        children: [
                          i.jsx(LocalizedText, { children: "Argument " }),
                          completions.hint.argumentIndex + 1,
                        ],
                      }),
                    ],
                  }),
                completions.items.length > 0 &&
                  i.jsx("div", {
                    className: "formula-suggestion-help",
                    children: i.jsx(LocalizedText, {
                      children: "↑ ↓ choose · Tab / Enter insert · Esc close",
                    }),
                  }),
              ],
            }),
            document.body,
          )
    );
  }),
  k2e = 6,
  S2e = 1.35,
  clampNumber = (value, min, max) => Math.max(min, Math.min(max, value));
function buildCellFontString(cell) {
  const effectiveCellStyle = resolveEffectiveCellStyle(cell?.style),
    fontSizeCss = effectiveCellStyle?.fontSize
      ? `${effectiveCellStyle.fontSize}pt`
      : "14px",
    fontFamilyCss = effectiveCellStyle?.fontFamily
      ? JSON.stringify(effectiveCellStyle.fontFamily)
      : "Calibri, Aptos, Arial, sans-serif";
  return `${effectiveCellStyle?.italic ? "italic" : "normal"} ${effectiveCellStyle?.bold ? "700" : "400"} ${fontSizeCss} ${fontFamilyCss}`;
}
function parseFontSizePx(text, cellFontString) {
  const sizeMatch = cellFontString.match(/([\d.]+)(px|pt)/),
    fontSizePx = sizeMatch
      ? Number(sizeMatch[1]) * (sizeMatch[2] === "pt" ? 4 / 3 : 1)
      : 14,
    boldScale = cellFontString.includes("700") ? 1.06 : 1;
  return (
    Array.from(text).reduce(
      (total, character) =>
        total +
        (/\s/u.test(character)
          ? 0.33
          : /[ilI.,'`!:;]/.test(character)
            ? 0.28
            : /[MW@#]/.test(character)
              ? 0.85
              : /[^\u0000-\u024f]/u.test(character)
                ? 1
                : 0.56),
      0,
    ) *
    fontSizePx *
    boldScale
  );
}
function createTextMeasurer() {
  if (typeof document > "u" || typeof CanvasRenderingContext2D > "u")
    return parseFontSizePx;
  try {
    const context = document.createElement("canvas").getContext("2d");
    if (!context) return parseFontSizePx;
    let lastFont = "";
    return (text, cellFontString) => {
      lastFont !== cellFontString &&
        ((context.font = cellFontString), (lastFont = cellFontString));
      const width = context.measureText(text).width;
      return Number.isFinite(width) && width >= 0
        ? width
        : parseFontSizePx(text, cellFontString);
    };
  } catch {
    return parseFontSizePx;
  }
}
function measureTextWidth(textMeasurer, text, cellFontString) {
  try {
    const measuredWidth = textMeasurer(text, cellFontString);
    if (Number.isFinite(measuredWidth) && measuredWidth >= 0)
      return measuredWidth;
  } catch {}
  return parseFontSizePx(text, cellFontString);
}
function measureTextWithTabs(text, cellFontString, textMeasurer) {
  if (!text.includes("	"))
    return measureTextWidth(textMeasurer, text, cellFontString);
  const tabStopWidth = Math.max(
    1,
    measureTextWidth(textMeasurer, " ", cellFontString) * 8,
  );
  let cursorX = 0;
  return (
    text.split("	").forEach((segment, segmentIndex) => {
      (segmentIndex &&
        (cursorX = (Math.floor(cursorX / tabStopWidth) + 1) * tabStopWidth),
        (cursorX += measureTextWidth(textMeasurer, segment, cellFontString)));
    }),
    cursorX
  );
}
function resolveDisplayedCellText(
  sheet,
  cellAddress,
  cell,
  workbookWithSettings,
  isShowingFormulas,
  dateSystem,
) {
  return isShowingFormulas &&
    typeof cell.raw == "string" &&
    cell.raw.startsWith("=") &&
    !cell.literal
    ? cell.raw
    : formatNumberWithPattern(
        workbookWithSettings(sheet.id, cellAddress),
        cell,
        dateSystem,
      );
}
function measureNoteIndicatorWidth(cell, textMeasurer) {
  return Math.max(
    cell.note ? 5 : 0,
    cell.tickmark
      ? Math.min(
          40,
          measureTextWidth(
            textMeasurer,
            cell.tickmark,
            "normal 400 8px Arial",
          ) + 4,
        )
      : 0,
  );
}
function computeCellTextAlignment(cell, textMeasurer) {
  return (
    k2e +
    measureCellBorderInsets(cell.style).horizontal +
    (cell.style?.align === "center" ? 0 : indentToPx(cell.style)) +
    measureNoteIndicatorWidth(cell, textMeasurer)
  );
}
function validateAutoFitIndices(selectedAxisIndexes, axisCount, axisName) {
  for (const selectedAxisIndex of selectedAxisIndexes)
    if (
      !Number.isInteger(selectedAxisIndex) ||
      selectedAxisIndex < 0 ||
      selectedAxisIndex >= axisCount
    )
      throw new Error(`AutoFit ${axisName} index is outside the worksheet.`);
  return new Set(selectedAxisIndexes);
}
function autoFitColumnWidths(
  sheet,
  selectedAxisIndexes,
  workbookWithSettings,
  textMeasurer,
  isShowingFormulas = false,
  dateSystem = "1900",
) {
  const selectedColumns = validateAutoFitIndices(
      selectedAxisIndexes,
      sheet.colCount,
      "column",
    ),
    columnWidths = Object.fromEntries(
      [...selectedColumns].map((columnIndex) => [
        String(columnIndex),
        defaultColumnWidthOf(sheet),
      ]),
    ),
    measuredColumns = new Set();
  for (const [cellAddress, cell] of Object.entries(
    vx(sheet, workbookWithSettings),
  )) {
    const { col } = _t(cellAddress);
    if (!selectedColumns.has(col)) continue;
    const displayedCellText = resolveDisplayedCellText(
      sheet,
      cellAddress,
      cell,
      workbookWithSettings,
      isShowingFormulas,
      dateSystem,
    );
    if (!displayedCellText && !cell.note && !cell.tickmark) continue;
    const cellFontString = buildCellFontString(cell),
      textLines = resolveEffectiveCellStyle(cell.style).wrap
        ? displayedCellText.replace(
            /\r\n?/g,
            `
`,
          ).split(`
`)
        : [displayedCellText.replace(/[\t\n\r\f ]+/g, " ")];
    let maxTextWidth = 0;
    for (const line of textLines)
      maxTextWidth = Math.max(
        maxTextWidth,
        measureTextWithTabs(line, cellFontString, textMeasurer),
      );
    const columnWidth = clampNumber(
      Math.ceil(maxTextWidth + computeCellTextAlignment(cell, textMeasurer)),
      minColumnWidth,
      maxColumnWidth,
    );
    ((columnWidths[String(col)] = measuredColumns.has(col)
      ? Math.max(columnWidths[String(col)], columnWidth)
      : columnWidth),
      measuredColumns.add(col));
  }
  return columnWidths;
}
function countWrappedTextLines(
  displayedCellText,
  availableWidth,
  cellFontString,
  textMeasurer,
  maxLines,
) {
  let lineCount = 0;
  for (const paragraph of displayedCellText.replace(
    /\r\n?/g,
    `
`,
  ).split(`
`)) {
    const chars = Array.from(paragraph);
    if (!chars.length) {
      if (++lineCount >= maxLines) return maxLines;
      continue;
    }
    let lineStart = 0;
    for (; lineStart < chars.length;) {
      if (++lineCount >= maxLines) return maxLines;
      let fitEnd = lineStart + 1,
        searchEnd = chars.length;
      for (; fitEnd < searchEnd;) {
        const midIndex = Math.ceil((fitEnd + searchEnd) / 2);
        measureTextWithTabs(
          chars.slice(lineStart, midIndex).join(""),
          cellFontString,
          textMeasurer,
        ) <= availableWidth
          ? (fitEnd = midIndex)
          : (searchEnd = midIndex - 1);
      }
      let nextLineStart = fitEnd;
      if (nextLineStart >= chars.length) break;
      for (
        let breakIndex = nextLineStart - 1;
        breakIndex > lineStart;
        breakIndex--
      )
        if (/[\t \-\u200b]/u.test(chars[breakIndex])) {
          nextLineStart = breakIndex + 1;
          break;
        }
      for (
        ;
        nextLineStart < chars.length && /[\t ]/u.test(chars[nextLineStart]);
      )
        nextLineStart++;
      lineStart = nextLineStart;
    }
  }
  return Math.max(1, lineCount);
}
function autoFitRowHeights(
  sheet,
  selectedAxisIndexes,
  workbookWithSettings,
  textMeasurer,
  isShowingFormulas = false,
  dateSystem = "1900",
) {
  const validRowIndexes = validateAutoFitIndices(
      selectedAxisIndexes,
      sheet.rowCount,
      "row",
    ),
    rowHeights = Object.fromEntries(
      [...validRowIndexes].map((rowIndex) => [
        String(rowIndex),
        defaultRowHeightOf(sheet),
      ]),
    ),
    measuredRows = new Set();
  for (const [cellAddress, cell] of Object.entries(
    vx(sheet, workbookWithSettings),
  )) {
    const { row, col } = _t(cellAddress);
    if (!validRowIndexes.has(row)) continue;
    const displayedCellText = resolveDisplayedCellText(
      sheet,
      cellAddress,
      cell,
      workbookWithSettings,
      isShowingFormulas,
      dateSystem,
    );
    if (!displayedCellText && !resolveEffectiveCellStyle(cell.style).fontSize)
      continue;
    const cellFontString = buildCellFontString(cell),
      lineHeight =
        (resolveEffectiveCellStyle(cell.style).fontSize
          ? (resolveEffectiveCellStyle(cell.style).fontSize * 4) / 3
          : 14) * S2e,
      vertical = measureCellBorderInsets(cell.style).vertical,
      maxLineCount = Math.ceil((maxRowHeight - vertical) / lineHeight),
      textWidth = Math.max(
        1,
        columnWidthAt(sheet, col) -
          computeCellTextAlignment(cell, textMeasurer),
      ),
      lineCount = resolveEffectiveCellStyle(cell.style).wrap
        ? countWrappedTextLines(
            displayedCellText,
            textWidth,
            cellFontString,
            textMeasurer,
            maxLineCount,
          )
        : 1,
      rowHeight = clampNumber(
        Math.ceil(Math.max(20, lineCount * lineHeight) + vertical),
        minRowHeight,
        maxRowHeight,
      );
    ((rowHeights[String(row)] = measuredRows.has(row)
      ? Math.max(rowHeights[String(row)], rowHeight)
      : rowHeight),
      measuredRows.add(row));
  }
  return rowHeights;
}
const D2e = 16e6;
function Tp(range, row, col) {
  return (
    row >= range.firstRow &&
    row <= range.lastRow &&
    col >= range.firstCol &&
    col <= range.lastCol
  );
}

function lowerBoundIndex(sortedArray, target) {
  let low = 0,
    high = sortedArray.length;
  for (; low < high;) {
    const mid = (low + high) >>> 1;
    sortedArray[mid] < target ? (low = mid + 1) : (high = mid);
  }
  return low;
}
function buildVisibleAxisLayout(
  e,
  defaultSize,
  sizeOverrides = {},
  hiddenIndexes = [],
) {
  const axisVisibilityIndex = createAxisVisibilityIndex(e, hiddenIndexes),
    viewSizes = Object.fromEntries(
      Object.entries(sizeOverrides).flatMap(([logicalKey, overrideSize]) => {
        const viewIndex = axisVisibilityIndex.viewIndex(Number(logicalKey));
        return viewIndex === null ? [] : [[viewIndex, overrideSize]];
      }),
    );
  e = axisVisibilityIndex.count;
  const overriddenIndexes = Object.keys(viewSizes)
      .map(Number)
      .filter(
        (candidateIndex) =>
          candidateIndex >= 0 &&
          candidateIndex < e &&
          viewSizes[candidateIndex] > 0 &&
          viewSizes[candidateIndex] !== defaultSize,
      )
      .sort((leftIndex, rightIndex) => leftIndex - rightIndex),
    cumulativeOffsets = [0];
  for (const overriddenIndex of overriddenIndexes)
    cumulativeOffsets.push(
      cumulativeOffsets.at(-1) + viewSizes[overriddenIndex] - defaultSize,
    );
  const offset = (logicalIndex) => {
    const clampedIndex = clampNumber(logicalIndex, 0, e);
    return (
      clampedIndex * defaultSize +
      cumulativeOffsets[lowerBoundIndex(overriddenIndexes, clampedIndex)]
    );
  };
  return {
    ...axisVisibilityIndex,
    total: offset(e),
    offset: offset,
    size: (axisIndex) => viewSizes[axisIndex] ?? defaultSize,
    indexAt: (targetOffset) => {
      let low = 0,
        high = e;
      for (; low < high;) {
        const mid = Math.floor((low + high) / 2);
        offset(mid + 1) <= targetOffset ? (low = mid + 1) : (high = mid);
      }
      return Math.min(Math.max(0, e - 1), low);
    },
  };
}
function sliceVisibleAxis(visibleAxis, frozenRowCount) {
  if (
    !Number.isInteger(frozenRowCount) ||
    frozenRowCount < 0 ||
    frozenRowCount > visibleAxis.count
  )
    throw Error("Choose a valid visible-axis slice.");
  const frozenOffset = visibleAxis.offset(frozenRowCount),
    count = visibleAxis.count - frozenRowCount,
    logicalIndex = (sliceIndex) =>
      Number.isInteger(sliceIndex) && sliceIndex >= 0 && sliceIndex < count
        ? visibleAxis.logicalIndex(sliceIndex + frozenRowCount)
        : -1,
    viewIndex = (sourceIndex) => {
      const fullViewIndex = visibleAxis.viewIndex(sourceIndex);
      return fullViewIndex === null || fullViewIndex < frozenRowCount
        ? null
        : fullViewIndex - frozenRowCount;
    },
    countBefore = (beforeIndex) =>
      Math.max(0, visibleAxis.countBefore(beforeIndex) - frozenRowCount);
  return {
    count: count,
    logicalCount: visibleAxis.logicalCount,
    logicalIndex: logicalIndex,
    viewIndex: viewIndex,
    countBefore: countBefore,
    isHidden: (axisIndex) => viewIndex(axisIndex) === null,
    next: (fromIndex, step, shouldIncludeCurrent = false) => {
      const startIndex = fromIndex + (shouldIncludeCurrent ? 0 : step),
        visibleIndex =
          step === 1
            ? countBefore(Math.max(0, startIndex))
            : countBefore(Math.min(visibleAxis.logicalCount, startIndex + 1)) -
              1,
        nextIndex = logicalIndex(visibleIndex);
      return nextIndex < 0 ? null : nextIndex;
    },
    total: visibleAxis.total - frozenOffset,
    size: (sizeIndex) => visibleAxis.size(sizeIndex + frozenRowCount),
    offset: (offsetIndex) =>
      visibleAxis.offset(
        Math.max(0, Math.min(count, offsetIndex)) + frozenRowCount,
      ) - frozenOffset,
    indexAt: (targetOffset) =>
      count
        ? Math.max(
            0,
            Math.min(
              count - 1,
              visibleAxis.indexAt(targetOffset + frozenOffset) - frozenRowCount,
            ),
          )
        : 0,
  };
}
function resolveAxisScale(
  total,
  viewportSize,
  zoomScale,
  maxPhysicalPixels = D2e,
) {
  const viewport = Math.max(1, viewportSize),
    logicalPixels = total * zoomScale,
    physicalPixels = Math.min(logicalPixels, maxPhysicalPixels);
  return {
    viewport: viewport,
    logicalPixels: logicalPixels,
    physicalPixels: physicalPixels,
    logicalMax: Math.max(0, logicalPixels - viewport),
    physicalMax: Math.max(0, physicalPixels - viewport),
  };
}
function physicalToLogicalPixels(
  physicalOffset,
  total,
  viewportSize,
  zoomScale,
) {
  const axisScale = resolveAxisScale(total, viewportSize, zoomScale);
  return axisScale.physicalMax
    ? ((clampNumber(physicalOffset, 0, axisScale.physicalMax) /
        axisScale.physicalMax) *
        axisScale.logicalMax) /
        zoomScale
    : 0;
}
function logicalToPhysicalPixels(
  logicalOffset,
  total,
  viewportSize,
  zoomScale,
) {
  const axisScale = resolveAxisScale(total, viewportSize, zoomScale);
  return axisScale.logicalMax
    ? (clampNumber(logicalOffset * zoomScale, 0, axisScale.logicalMax) /
        axisScale.logicalMax) *
        axisScale.physicalMax
    : 0;
}

export {
  SplitPaneLayout,
  buildVisibleAxisLayout,
  resolveAxisScale,
  useSplitPaneSizes,
  physicalToLogicalPixels,
  autoFitColumnWidths,
  createTextMeasurer,
  Tp,
  sliceVisibleAxis,
  autoFitRowHeights,
  isWithinFormulaSuggestions,
  FormulaSuggestionList,
  logicalToPhysicalPixels,
  resolveCellNavigation,
  FindReplaceDialog,
};
export { clampNumber };
