import {
  coerceValueToText,
  flattenNestedValues,
} from "../02-text-and-financial-functions.js";
import qre from "fast-formula-parser";
import formulaOperators from "fast-formula-parser/formulas/operators.js";
import { unwrapDefaultExport } from "../01-schemas-date-functions-and-validation.js";
import {
  builtInDateTimeFormats,
  validateNumberFormatCode,
  dateTimeFormatKinds,
  supportedNumberLocales,
  throwWorkbookError,
  supportedCurrencyCodes,
  Uv,
  numericFormatKinds,
  ZQ,
  numberFormatToIntlOptions,
} from "./01-cell-styles-and-themes.js";

function validateNumberFormatSettings(settings) {
  (!ZQ(settings) ||
    Object.keys(settings).some(
      (key) =>
        ![
          "category",
          "locale",
          "currency",
          "decimals",
          "useGrouping",
          "scale",
          "negative",
          "pattern",
        ].includes(key),
    )) &&
    throwWorkbookError(
      "Number format must be a supported plain settings object.",
    );
  const rawSettings = settings;
  Uv.includes(rawSettings.category) ||
    throwWorkbookError("Choose a supported number format category.");
  const category = rawSettings.category,
    locale = rawSettings.locale ?? "en-US";
  supportedNumberLocales.includes(locale) ||
    throwWorkbookError("Choose one of the supported display locales.");
  const normalizedSettings = { category: category, locale: locale };
  if (numericFormatKinds.has(category)) {
    if (category === "currency" || category === "accounting") {
      const currencyCode = rawSettings.currency ?? "USD";
      (supportedCurrencyCodes.includes(currencyCode) ||
        throwWorkbookError("Choose a supported currency."),
        (normalizedSettings.currency = currencyCode));
    } else
      rawSettings.currency !== undefined &&
        throwWorkbookError(
          "Currency applies only to Currency or Accounting formats.",
        );
    const decimalPlaces =
      rawSettings.decimals ?? (normalizedSettings.currency === "JPY" ? 0 : 2);
    ((typeof decimalPlaces != "number" ||
      !Number.isInteger(decimalPlaces) ||
      decimalPlaces < 0 ||
      decimalPlaces > 12) &&
      throwWorkbookError("Choose 0–12 decimal places."),
      (normalizedSettings.decimals = decimalPlaces),
      rawSettings.useGrouping !== undefined &&
        typeof rawSettings.useGrouping != "boolean" &&
        throwWorkbookError("The grouping option must be true or false."),
      (normalizedSettings.useGrouping =
        rawSettings.useGrouping ?? category !== "percent"));
    const scale = rawSettings.scale ?? 1;
    ((![1, 1e3, 1e6].includes(scale) ||
      (category === "percent" && scale !== 1)) &&
      throwWorkbookError(
        "Choose units, thousands or millions; percentages use units.",
      ),
      (normalizedSettings.scale = scale),
      rawSettings.negative !== undefined &&
        (typeof rawSettings.negative != "string" ||
          !["minus", "parentheses"].includes(rawSettings.negative)) &&
        throwWorkbookError("Choose minus signs or parentheses for negatives."),
      category === "accounting" &&
        rawSettings.negative !== undefined &&
        rawSettings.negative !== "parentheses" &&
        throwWorkbookError("Accounting negatives use parentheses."),
      (normalizedSettings.negative =
        category === "accounting"
          ? "parentheses"
          : (rawSettings.negative ?? "minus")),
      rawSettings.pattern !== undefined &&
        throwWorkbookError(
          "Use the Custom category to enter a format pattern.",
        ));
  } else if (
    ([
      rawSettings.currency,
      rawSettings.decimals,
      rawSettings.useGrouping,
      rawSettings.scale,
      rawSettings.negative,
    ].some((settingValue) => settingValue !== undefined) &&
      throwWorkbookError(
        "These numeric settings do not apply to the chosen category.",
      ),
    dateTimeFormatKinds.has(category) || category === "custom")
  ) {
    const defaultPattern =
        category === "custom" ? "0.00" : builtInDateTimeFormats[category][0],
      pattern = rawSettings.pattern ?? defaultPattern;
    ((typeof pattern != "string" ||
      !pattern ||
      pattern.length > 200 ||
      /[\u0000-\u001f\u007f]/.test(pattern)) &&
      throwWorkbookError(
        "Enter a format pattern of 1–200 characters without control characters.",
      ),
      (normalizedSettings.pattern = pattern));
    const patternInfo = validateNumberFormatCode(normalizedSettings.pattern);
    if (category !== "custom") {
      patternInfo.kind !== "temporal" &&
        throwWorkbookError(
          "Choose a date, time or elapsed-time pattern for this category.",
        );
      const patternCategory = patternInfo.category;
      patternCategory !== category &&
        throwWorkbookError(
          `This pattern is a ${patternCategory} format. Choose the matching category or Custom.`,
        );
    }
  } else
    rawSettings.pattern !== undefined &&
      throwWorkbookError("General does not accept a custom pattern.");
  return normalizedSettings;
}
function renderTextSectionPattern(textSection, replacementText) {
  return textSection.replace(
    /"([^"]*)"|\\(.)|@/g,
    (matchedText, quotedText, escapedChar) =>
      quotedText ??
      escapedChar ??
      (matchedText === "@" ? replacementText : matchedText),
  );
}
const getLocaleDecimalSeparator = (locale) =>
  new Intl.NumberFormat(locale)
    .formatToParts(1.1)
    .find((part) => part.type === "decimal")?.value ?? ".";
function formatNumericSection(value, compiledFormat, locale) {
  const isNegativeSection = value < 0 && compiledFormat.sections.length > 1,
    { parsed, divisor } =
      compiledFormat.sections[
        value === 0 && compiledFormat.sections.length > 2
          ? 2
          : isNegativeSection
            ? 1
            : 0
      ];
  if (!parsed.hasNumber) return parsed.prefix + parsed.suffix;
  const scaledValue = (Math.abs(value) / divisor) * (parsed.percent ? 100 : 1);
  if (!Number.isFinite(scaledValue)) return "#NUM!";
  let formattedDigits = new Intl.NumberFormat(locale, {
    minimumFractionDigits: parsed.minDecimals,
    maximumFractionDigits: parsed.maxDecimals,
    minimumIntegerDigits: Math.max(1, parsed.minIntegerDigits),
    useGrouping: parsed.grouping,
  }).format(scaledValue);
  return (
    parsed.minIntegerDigits ||
      (formattedDigits === "0" && !parsed.minDecimals
        ? (formattedDigits = "")
        : formattedDigits.startsWith(`0${getLocaleDecimalSeparator(locale)}`) &&
          (formattedDigits = formattedDigits.slice(1))),
    (value < 0 && !isNegativeSection ? "-" : "") +
      parsed.prefix +
      formattedDigits +
      parsed.suffix
  );
}
function formatTemporalSerialValue(
  serialValue,
  compiledFormat,
  locale,
  dateSystem,
) {
  const isDuration = compiledFormat.category === "duration",
    maxSerial = 2958465 - (dateSystem === "1904" && !isDuration ? 1462 : 0);
  if (
    Math.abs(serialValue) >= maxSerial + 1 ||
    (!isDuration && serialValue < 0)
  )
    return "#NUM!";
  const isNotDateCategory = compiledFormat.category !== "date",
    fractionScale = 10 ** compiledFormat.fraction,
    hasSecondToken = compiledFormat.tokens.some(
      (token) =>
        token.kind === "second" ||
        (token.kind === "elapsed" && token.text.startsWith("s")),
    ),
    totalSeconds = isNotDateCategory
      ? hasSecondToken
        ? Math.round(Math.abs(serialValue) * 86400 * fractionScale) /
          fractionScale
        : Math.floor(Math.abs(serialValue) * 86400)
      : Math.floor(serialValue) * 86400;
  let dayNumber = Math.floor(totalSeconds / 86400),
    secondsOfDay = totalSeconds - dayNumber * 86400;
  if (dayNumber > maxSerial) return "#NUM!";
  const millisecondsOfDay = Math.round(secondsOfDay * 1e3),
    hours = Math.floor(millisecondsOfDay / 36e5),
    minutes = Math.floor(millisecondsOfDay / 6e4) % 60,
    seconds = Math.floor(millisecondsOfDay / 1e3) % 60,
    fractionDigits = Math.floor(
      (millisecondsOfDay % 1e3) / 10 ** (3 - compiledFormat.fraction),
    ),
    isPhantomLeapDay = dateSystem === "1900" && dayNumber === 60,
    utcTimestamp =
      dateSystem === "1904"
        ? Date.UTC(1904, 0, 1) + dayNumber * 864e5
        : Date.UTC(1899, 11, 31) +
          (dayNumber > 60 ? dayNumber - 1 : dayNumber) * 864e5,
    dateValue = new Date(utcTimestamp),
    year = isPhantomLeapDay ? 1900 : dateValue.getUTCFullYear(),
    monthIndex = isPhantomLeapDay ? 1 : dateValue.getUTCMonth(),
    dayOfMonth = isPhantomLeapDay ? 29 : dateValue.getUTCDate(),
    hasAmPmToken = compiledFormat.tokens.some((token) => token.kind === "ampm"),
    padNumber = (number, width) => String(number).padStart(width, "0"),
    formattedText = compiledFormat.tokens
      .map((token) => {
        switch (token.kind) {
          case "literal":
            return token.text;
          case "year":
            return token.text.length === 2
              ? padNumber(year % 100, 2)
              : padNumber(year, 4);
          case "month":
            return token.text.length < 3
              ? padNumber(monthIndex + 1, token.text.length)
              : new Intl.DateTimeFormat(locale, {
                  month: token.text.length === 3 ? "short" : "long",
                  timeZone: "UTC",
                }).format(new Date(Date.UTC(2e3, monthIndex, 1)));
          case "day":
            return padNumber(dayOfMonth, token.text.length);
          case "hour":
            return padNumber(
              hasAmPmToken ? hours % 12 || 12 : hours,
              token.text.length,
            );
          case "minute":
            return padNumber(minutes, token.text.length);
          case "second":
            return padNumber(seconds, token.text.length);
          case "fraction":
            return (
              getLocaleDecimalSeparator(locale) +
              padNumber(fractionDigits, compiledFormat.fraction)
            );
          case "ampm": {
            const meridiem = hours < 12 ? "AM" : "PM",
              meridiemText = token.text.length === 3 ? meridiem[0] : meridiem;
            return token.text === token.text.toLowerCase()
              ? meridiemText.toLowerCase()
              : meridiemText;
          }
          case "elapsed":
            return padNumber(
              Math.floor(
                totalSeconds /
                  (token.text[0] === "h"
                    ? 3600
                    : token.text[0] === "m"
                      ? 60
                      : 1),
              ),
              token.text.length,
            );
        }
      })
      .join("");
  return (isDuration && serialValue < 0 ? "-" : "") + formattedText;
}
function formatCellValueWithNumberFormat(
  value,
  numberFormatSettings,
  dateSystem = "1900",
) {
  const settings = validateNumberFormatSettings(numberFormatSettings);
  if (value === null) return { text: "" };
  if (typeof value == "boolean") return { text: value ? "TRUE" : "FALSE" };
  if (typeof value == "string") {
    if (settings.category === "custom" && !/^#[A-Z0-9/?!]+$/.test(value)) {
      const textPatternInfo = validateNumberFormatCode(settings.pattern);
      if (
        textPatternInfo.kind === "numeric" &&
        textPatternInfo.textSection !== undefined
      )
        return {
          text: renderTextSectionPattern(textPatternInfo.textSection, value),
        };
    }
    return { text: value };
  }
  if (!Number.isFinite(value)) return { text: "#NUM!" };
  if (settings.category === "general")
    return {
      text: new Intl.NumberFormat(settings.locale, {
        maximumSignificantDigits: 15,
        useGrouping: false,
      }).format(value === 0 ? 0 : value),
    };
  if (settings.pattern) {
    const patternInfo = validateNumberFormatCode(settings.pattern);
    return {
      text:
        patternInfo.kind === "temporal"
          ? formatTemporalSerialValue(
              value,
              patternInfo,
              settings.locale,
              dateSystem,
            )
          : formatNumericSection(value, patternInfo, settings.locale),
    };
  }
  const scaledValue = value / (settings.scale ?? 1),
    scaleSuffix =
      settings.scale === 1e3 ? " K" : settings.scale === 1e6 ? " M" : "";
  if (settings.category === "percent" && !Number.isFinite(scaledValue * 100))
    return { text: "#NUM!" };
  const formatOptions = {
    minimumFractionDigits: settings.decimals,
    maximumFractionDigits: settings.decimals,
    useGrouping: settings.useGrouping,
  };
  if (settings.category === "accounting") {
    const currencyParts = new Intl.NumberFormat(settings.locale, {
        ...formatOptions,
        style: "currency",
        currency: settings.currency,
      }).formatToParts(Math.abs(scaledValue)),
      symbol =
        currencyParts.find((part) => part.type === "currency")?.value ??
        settings.currency,
      symbolPosition =
        currencyParts.findIndex(
          (currencyPart) => currencyPart.type === "currency",
        ) <
        currencyParts.findIndex((integerPart) => integerPart.type === "integer")
          ? "start"
          : "end",
      absoluteAmountText = new Intl.NumberFormat(
        settings.locale,
        formatOptions,
      ).format(Math.abs(scaledValue)),
      amount =
        (scaledValue === 0
          ? "-"
          : scaledValue < 0
            ? `(${absoluteAmountText})`
            : absoluteAmountText) + scaleSuffix;
    return {
      text:
        symbolPosition === "start"
          ? `${symbol} ${amount}`
          : `${amount} ${symbol}`,
      accounting: {
        symbol: symbol,
        amount: amount,
        symbolPosition: symbolPosition,
      },
    };
  }
  const isParenthesizedNegative =
      scaledValue < 0 && settings.negative === "parentheses",
    formattedNumber = new Intl.NumberFormat(settings.locale, {
      ...formatOptions,
      ...(settings.category === "currency"
        ? { style: "currency", currency: settings.currency }
        : settings.category === "percent"
          ? { style: "percent" }
          : {}),
    }).format(isParenthesizedNegative ? Math.abs(scaledValue) : scaledValue);
  return {
    text:
      (isParenthesizedNegative ? `(${formattedNumber})` : formattedNumber) +
      scaleSuffix,
  };
}
function getImportedNumberFormatWarning(cell) {
  if (cell) {
    if (cell.numberFormat)
      try {
        validateNumberFormatSettings(cell.numberFormat);
        return;
      } catch (error) {
        return error instanceof Error
          ? error.message
          : "Unsupported number format settings.";
      }
    if (
      !(cell.numFmt === "@" && typeof cell.raw == "string") &&
      !(cell.format === "date" && cell.numFmt?.toLowerCase() === "yyyy-mm-dd")
    ) {
      if (
        cell.numFmt &&
        /\[(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\d+)\]/i.test(
          cell.numFmt,
        )
      )
        return "The legacy format's color directive is retained but does not color the text. Use conditional formatting to apply color explicitly.";
      if (
        cell.numFmt &&
        cell.numFmt.toLowerCase() !== "general" &&
        numberFormatToIntlOptions(1.25, cell.numFmt) === undefined
      )
        return `The imported pattern “${cell.numFmt}” is outside the legacy display subset. The basic ${cell.format ?? "general"} display is used; the original pattern and raw values are retained. Choose a supported format to change its display explicitly.`;
    }
  }
}
const lcidByLocale = {
  "en-US": "409",
  "en-GB": "809",
  "de-DE": "407",
  "fr-FR": "40C",
  "ja-JP": "411",
  "en-IN": "4009",
};
function buildNumberFormatCode(numberFormat) {
  const settings = validateNumberFormatSettings(numberFormat),
    localePrefix =
      settings.locale === "en-US" ? "" : `[$-${lcidByLocale[settings.locale]}]`;
  if (settings.category === "general") return "General";
  if (settings.pattern) return localePrefix + settings.pattern;
  const decimalPattern = settings.decimals
      ? "." + "0".repeat(settings.decimals)
      : "",
    scalingCommas =
      settings.scale === 1e3 ? "," : settings.scale === 1e6 ? ",," : "",
    scaleSuffixLiteral =
      settings.scale === 1e3 ? '" K"' : settings.scale === 1e6 ? '" M"' : "",
    numberPattern =
      (settings.useGrouping ? "#,##0" : "0") + decimalPattern + scalingCommas;
  if (settings.category === "percent")
    return localePrefix + numberPattern + "%";
  if (settings.category === "currency" || settings.category === "accounting") {
    const accounting = formatCellValueWithNumberFormat(1, {
        ...settings,
        category: "accounting",
        negative: "parentheses",
      }).accounting,
      quotedSymbol = `"${accounting.symbol.replace(/"/g, '""')}"`;
    if (settings.category === "accounting") {
      const placeholderDigits = "?".repeat(settings.decimals);
      return accounting.symbolPosition === "start"
        ? `${localePrefix}_(${quotedSymbol}* ${numberPattern}${scaleSuffixLiteral}_);_(${quotedSymbol}* (${numberPattern}${scaleSuffixLiteral});_(${quotedSymbol}* "-"${placeholderDigits}_);_(@_)`
        : `${localePrefix}_(${numberPattern}${scaleSuffixLiteral}* ${quotedSymbol}_);_((${numberPattern}${scaleSuffixLiteral})* ${quotedSymbol});_("-"${placeholderDigits}* ${quotedSymbol}_);_(@_)`;
    }
    const currencyPattern =
      accounting.symbolPosition === "start"
        ? quotedSymbol + numberPattern + scaleSuffixLiteral
        : numberPattern + scaleSuffixLiteral + '" "' + quotedSymbol;
    return (
      localePrefix +
      currencyPattern +
      (settings.negative === "parentheses" ? `;(${currencyPattern})` : "")
    );
  }
  return (
    localePrefix +
    numberPattern +
    scaleSuffixLiteral +
    (settings.negative === "parentheses"
      ? `;(${numberPattern}${scaleSuffixLiteral})`
      : "")
  );
}
const FormulaParser = unwrapDefaultExport(qre);
// The parser library joins operands of "&" with JavaScript's own number
// formatting. Replace that one operator so "&" converts numbers the same way
// as CONCAT and the other text functions do.
unwrapDefaultExport(formulaOperators).Infix.concatOp = (
  left,
  operator,
  right,
  leftIsArray,
  rightIsArray,
) =>
  coerceValueToText(leftIsArray ? left[0][0] : left) +
  coerceValueToText(rightIsArray ? right[0][0] : right);
function toSignedDecimalParts(value) {
  const [mantissaText, exponentText = "0"] = Math.abs(value)
      .toString()
      .split("e"),
    [integerDigits, fractionDigits = ""] = mantissaText.split(".");
  return {
    coefficient:
      BigInt(integerDigits + fractionDigits) * (value < 0 ? -1n : 1n),
    exponent: Number(exponentText) - fractionDigits.length,
  };
}
function sumDecimalParts(values) {
  const coefficientsByExponent = new Map();
  for (const value of values) {
    if (!Number.isFinite(value)) return;
    const { coefficient, exponent } = toSignedDecimalParts(value);
    coefficientsByExponent.set(
      exponent,
      (coefficientsByExponent.get(exponent) ?? 0n) + coefficient,
    );
  }
  const nonZeroEntries = [...coefficientsByExponent].filter(
    ([, summedCoefficient]) => summedCoefficient !== 0n,
  );
  if (!nonZeroEntries.length) return { coefficient: 0n, exponent: 0 };
  const exponent = nonZeroEntries.reduce(
    (minExponent, [exponentKey]) => Math.min(minExponent, exponentKey),
    1 / 0,
  );
  return {
    coefficient: nonZeroEntries.reduce(
      (total, [entryExponent, entryCoefficient]) =>
        total + entryCoefficient * 10n ** BigInt(entryExponent - exponent),
      0n,
    ),
    exponent: exponent,
  };
}
function sumDecimalPartsToNumber(values) {
  const decimalSum = sumDecimalParts(values);
  return decimalSum ? +`${decimalSum.coefficient}e${decimalSum.exponent}` : NaN;
}
function exactAverageOfNumbers(values) {
  if (!values.length) return NaN;
  const decimalSum = sumDecimalParts(values);
  if (!decimalSum) return NaN;
  const signText = decimalSum.coefficient < 0n ? "-" : "",
    absoluteCoefficient =
      decimalSum.coefficient < 0n
        ? -decimalSum.coefficient
        : decimalSum.coefficient,
    valueCount = BigInt(values.length),
    wholePart = absoluteCoefficient / valueCount,
    fractionText = (
      ((absoluteCoefficient % valueCount) * 10n ** 40n) /
      valueCount
    )
      .toString()
      .padStart(40, "0");
  return +`${signText}${wholePart}.${fractionText}e${decimalSum.exponent}`;
}
function exactProductOfNumbers(values) {
  if (values.some((candidate) => !Number.isFinite(candidate))) return NaN;
  if (values.some((operand) => operand === 0)) return 0;
  let productCoefficient = 1n,
    productExponent = 0;
  for (const factor of values) {
    const signedDecimalParts = toSignedDecimalParts(factor);
    if (
      ((productCoefficient *= signedDecimalParts.coefficient),
      (productExponent += signedDecimalParts.exponent),
      values.length > 255)
    ) {
      const isNegativeProduct = productCoefficient < 0n,
        digitsText = (
          isNegativeProduct ? -productCoefficient : productCoefficient
        ).toString();
      digitsText.length > 80 &&
        ((productExponent += digitsText.length - 80),
        (productCoefficient =
          BigInt(digitsText.slice(0, 80)) * (isNegativeProduct ? -1n : 1n)));
    }
  }
  return +`${productCoefficient}e${productExponent}`;
}
const notArithmeticSentinel = Symbol("not-arithmetic"),
  arithmeticFunctionNames = new Set([
    "ROUND",
    "ROUNDUP",
    "ROUNDDOWN",
    "TRUNC",
    "MROUND",
    "CEILING",
    "FLOOR",
    "SUMPRODUCT",
  ]);
function toUnsignedDecimalParts(value) {
  const [mantissaText, exponentText = "0"] = Math.abs(value)
      .toString()
      .split("e"),
    [integerDigits, fractionDigits = ""] = mantissaText.split(".");
  return {
    coefficient: BigInt(integerDigits + fractionDigits),
    exponent: Number(exponentText) - fractionDigits.length,
  };
}
function roundQuotientWithMode(coefficient, divisor, roundingMode) {
  const quotient = coefficient / divisor,
    remainder = coefficient % divisor;
  return (
    quotient +
    (remainder !== 0n &&
    (roundingMode === "away" ||
      (roundingMode === "nearest" && remainder * 2n >= divisor))
      ? 1n
      : 0n)
  );
}
function roundDecimalToDigits(value, digits, roundingMode) {
  if (((digits = Math.trunc(digits)), value === 0 || digits >= 324))
    return value;
  if (digits <= -309) return roundingMode === "away" ? 1 / 0 : 0;
  const unsignedDecimalParts = toUnsignedDecimalParts(value),
    roundingExponent = unsignedDecimalParts.exponent + digits;
  if (roundingExponent >= 0) return value;
  const roundedCoefficient = roundQuotientWithMode(
    unsignedDecimalParts.coefficient,
    10n ** BigInt(-roundingExponent),
    roundingMode,
  );
  return Math.sign(value) * +`${roundedCoefficient}e${-digits}`;
}
function roundToNearestMultiple(value, multiple, roundingMode) {
  const unsignedDecimalParts = toUnsignedDecimalParts(value),
    multipleParts = toUnsignedDecimalParts(multiple),
    commonExponent = Math.min(
      unsignedDecimalParts.exponent,
      multipleParts.exponent,
    ),
    coefficient =
      unsignedDecimalParts.coefficient *
      10n ** BigInt(unsignedDecimalParts.exponent - commonExponent),
    multipleCoefficient =
      multipleParts.coefficient *
      10n ** BigInt(multipleParts.exponent - commonExponent),
    multipleCount = roundQuotientWithMode(
      coefficient,
      multipleCoefficient,
      roundingMode,
    );
  return (
    Math.sign(value) *
    +`${multipleCount * multipleParts.coefficient}e${multipleParts.exponent}`
  );
}
function evaluateArithmeticFunction(functionName, args, context) {
  if (!arithmeticFunctionNames.has(functionName)) return notArithmeticSentinel;
  const { retrieve, issue, isIssue } = context,
    retrievedArgs = args.map(retrieve),
    toFiniteResult = (value) =>
      Number.isFinite(value) ? (value === 0 ? 0 : value) : issue("#NUM!");
  if (functionName === "SUMPRODUCT") {
    if (!args.length || args.length > 255) return issue("#VALUE!");
    const normalizedArgs = retrievedArgs.map((retrievedArg) =>
        Array.isArray(retrievedArg)
          ? retrievedArg.every(Array.isArray)
            ? retrievedArg
            : [retrievedArg]
          : [[retrievedArg]],
      ),
      rowCount = normalizedArgs[0].length,
      columnCount = normalizedArgs[0][0]?.length ?? 0;
    if (
      !rowCount ||
      !columnCount ||
      normalizedArgs.some(
        (matrix) =>
          matrix.length !== rowCount ||
          matrix.some((row) => row.length !== columnCount),
      )
    )
      return issue("#VALUE!");
    if (rowCount * columnCount * normalizedArgs.length > 1e5)
      return issue("#LIMIT!");
    const productTerms = [];
    for (let rowIndex = 0; rowIndex < rowCount; rowIndex++)
      for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
        const cellFactors = [];
        for (const normalizedArg of normalizedArgs) {
          const cellValue = normalizedArg[rowIndex][columnIndex];
          if (isIssue(cellValue)) return cellValue;
          cellFactors.push(typeof cellValue == "number" ? cellValue : 0);
        }
        const cellProduct = exactProductOfNumbers(cellFactors);
        if (!Number.isFinite(cellProduct)) return issue("#NUM!");
        productTerms.push(cellProduct);
      }
    return toFiniteResult(sumDecimalPartsToNumber(productTerms));
  }
  if (args.length < (functionName === "TRUNC" ? 1 : 2) || args.length > 2)
    return issue("#VALUE!");
  const numericArgs = [];
  for (const argValue of retrievedArgs) {
    if (isIssue(argValue)) return argValue;
    if (
      Array.isArray(argValue) ||
      (typeof argValue == "string" &&
        !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(argValue.trim()))
    )
      return issue("#VALUE!");
    const numericValue = argValue == null ? 0 : Number(argValue);
    if (!Number.isFinite(numericValue)) return issue("#NUM!");
    numericArgs.push(numericValue);
  }
  const [sourceNumber, secondNumber = 0] = numericArgs;
  if (["ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC"].includes(functionName))
    return toFiniteResult(
      roundDecimalToDigits(
        sourceNumber,
        secondNumber,
        functionName === "ROUND"
          ? "nearest"
          : functionName === "ROUNDUP"
            ? "away"
            : "toward",
      ),
    );
  if (sourceNumber === 0 || secondNumber === 0) return 0;
  if (
    (functionName === "MROUND" &&
      Math.sign(sourceNumber) !== Math.sign(secondNumber)) ||
    (functionName === "FLOOR" && sourceNumber > 0 && secondNumber < 0)
  )
    return issue("#NUM!");
  const roundingMode =
    functionName === "MROUND"
      ? "nearest"
      : functionName === "CEILING"
        ? sourceNumber < 0 && secondNumber > 0
          ? "toward"
          : "away"
        : sourceNumber < 0 && secondNumber > 0
          ? "away"
          : "toward";
  return toFiniteResult(
    roundToNearestMultiple(sourceNumber, secondNumber, roundingMode),
  );
}
const XL = ["LINEST", "LOGEST", "TREND", "GROWTH"],
  notRegressionSentinel = Symbol("not-regression"),
  regressionSizeLimits = {
    values: 4e4,
    observations: 1e4,
    predictors: 32,
    work: 4e6,
    output: 1e4,
  },
  dotProductWithCompensation = (leftValues, rightValues) => {
    let sum = 0,
      compensation = 0;
    for (let index = 0; index < leftValues.length; index++) {
      const product = leftValues[index] * rightValues[index],
        nextSum = sum + product;
      ((compensation +=
        Math.abs(sum) >= Math.abs(product)
          ? sum - nextSum + product
          : product - nextSum + sum),
        (sum = nextSum));
    }
    return sum + compensation;
  };
function evaluateRegressionFunction(functionName, args, context) {
  if (!XL.includes(functionName)) return notRegressionSentinel;
  const fail = (errorCode = "#VALUE!") => {
      throw context.issue(errorCode);
    },
    toFiniteNumber = (value) =>
      Number.isFinite(value) ? value : fail("#NUM!");
  (!args.length || args.length > 4 || args[0] == null) && fail();
  let totalCells = 0;
  const readNumericMatrix = (rawValue) => {
      if (context.isIssue(rawValue)) throw rawValue;
      const rows = Array.isArray(rawValue)
          ? rawValue.every(Array.isArray)
            ? rawValue
            : [rawValue]
          : [[rawValue]],
        columnCount = rows[0]?.length ?? 0;
      return (
        (!rows.length ||
          !columnCount ||
          rows.some((row) => row.length !== columnCount)) &&
          fail(),
        (totalCells += rows.length * columnCount),
        totalCells > regressionSizeLimits.values && fail("#LIMIT!"),
        rows.map((row) =>
          row.map((cell) => {
            if (context.isIssue(cell)) throw cell;
            return (typeof cell != "number" && fail(), toFiniteNumber(cell));
          }),
        )
      );
    },
    readFlag = (reference, defaultValue) => {
      if (reference == null) return defaultValue;
      const retrievedValue = context.retrieve(reference);
      if (context.isIssue(retrievedValue)) throw retrievedValue;
      return retrievedValue == null
        ? false
        : typeof retrievedValue == "boolean"
          ? retrievedValue
          : typeof retrievedValue == "number"
            ? toFiniteNumber(retrievedValue) !== 0
            : typeof retrievedValue == "string" &&
                /^(true|false)$/i.test(retrievedValue)
              ? /^true$/i.test(retrievedValue)
              : fail();
    },
    isTrendOrGrowth = functionName === "TREND" || functionName === "GROWTH",
    isLogarithmic = functionName === "LOGEST" || functionName === "GROWTH",
    hasIntercept = readFlag(args[isTrendOrGrowth ? 3 : 2], true),
    hasStatistics = !isTrendOrGrowth && readFlag(args[3], false),
    knownYs = readNumericMatrix(context.retrieve(args[0])),
    observationCount = knownYs.length * knownYs[0].length;
  observationCount > regressionSizeLimits.observations && fail("#LIMIT!");
  const flatKnownYs = knownYs.flat();
  isLogarithmic && flatKnownYs.some((value) => value <= 0) && fail("#NUM!");
  const responseValues = isLogarithmic
      ? flatKnownYs.map(Math.log)
      : flatKnownYs,
    knownXs =
      args[1] == null
        ? knownYs.map((row, rowIndex) =>
            row.map(
              (Ee, columnIndex) => rowIndex * row.length + columnIndex + 1,
            ),
          )
        : readNumericMatrix(context.retrieve(args[1])),
    hasMatchingXShape =
      knownXs.length === knownYs.length &&
      knownXs[0].length === knownYs[0].length,
    isSingleRowY = knownYs.length === 1;
  let predictorRows;
  if (hasMatchingXShape) predictorRows = knownXs.flat().map((value) => [value]);
  else if (knownYs[0].length === 1 && knownXs.length === observationCount)
    predictorRows = knownXs;
  else if (isSingleRowY && knownXs[0].length === observationCount)
    predictorRows = Array.from(
      { length: observationCount },
      (se, predictorIndex) => knownXs.map((row) => row[predictorIndex]),
    );
  else return fail("#REF!");
  const predictorCount = predictorRows[0].length;
  (predictorCount > regressionSizeLimits.predictors ||
    observationCount * predictorCount * predictorCount >
      regressionSizeLimits.work) &&
    fail("#LIMIT!");
  const predictorMeans = Array.from(
      { length: predictorCount },
      (se, predictorIndex) =>
        hasIntercept
          ? exactAverageOfNumbers(
              predictorRows.map((row) => row[predictorIndex]),
            )
          : 0,
    ),
    responseMean = hasIntercept ? exactAverageOfNumbers(responseValues) : 0,
    centeredResponses = responseValues.map((value) =>
      toFiniteNumber(value - responseMean),
    ),
    responseScale = Math.max(...centeredResponses.map(Math.abs)) || 1,
    scaledResponses = centeredResponses.map((value) => value / responseScale),
    predictorScales = [],
    scaledColumns = Array.from(
      { length: predictorCount },
      (se, predictorIndex) => {
        const centeredColumn = predictorRows.map((row) =>
            toFiniteNumber(
              row[predictorIndex] - predictorMeans[predictorIndex],
            ),
          ),
          columnScale = Math.max(...centeredColumn.map(Math.abs)) || 1;
        return (
          predictorScales.push(columnScale),
          centeredColumn.map((value) => value / columnScale)
        );
      },
    ),
    initialScaledColumns = scaledColumns.map((column) => [...column]),
    columnOrder = Array.from({ length: predictorCount }, (se, index) => index),
    triangularFactor = Array.from({ length: predictorCount }, () =>
      Array(predictorCount).fill(0),
    ),
    orthonormalBasis = [],
    projections = [],
    rankTolerance =
      1e-12 *
      Math.max(
        1,
        ...scaledColumns.map((column) =>
          Math.sqrt(dotProductWithCompensation(column, column)),
        ),
      );
  let rank = 0;
  for (let pivotIndex = 0; pivotIndex < predictorCount; pivotIndex++) {
    let pivotColumn = pivotIndex;
    for (
      let scanIndex = pivotIndex + 1;
      scanIndex < predictorCount;
      scanIndex++
    )
      dotProductWithCompensation(
        scaledColumns[scanIndex],
        scaledColumns[scanIndex],
      ) >
        dotProductWithCompensation(
          scaledColumns[pivotColumn],
          scaledColumns[pivotColumn],
        ) && (pivotColumn = scanIndex);
    (([scaledColumns[pivotIndex], scaledColumns[pivotColumn]] = [
      scaledColumns[pivotColumn],
      scaledColumns[pivotIndex],
    ]),
      ([columnOrder[pivotIndex], columnOrder[pivotColumn]] = [
        columnOrder[pivotColumn],
        columnOrder[pivotIndex],
      ]));
    for (let priorRowIndex = 0; priorRowIndex < pivotIndex; priorRowIndex++)
      [
        triangularFactor[priorRowIndex][pivotIndex],
        triangularFactor[priorRowIndex][pivotColumn],
      ] = [
        triangularFactor[priorRowIndex][pivotColumn],
        triangularFactor[priorRowIndex][pivotIndex],
      ];
    const pivotNorm = Math.sqrt(
      dotProductWithCompensation(
        scaledColumns[pivotIndex],
        scaledColumns[pivotIndex],
      ),
    );
    if (pivotNorm <= rankTolerance) break;
    ((triangularFactor[pivotIndex][pivotIndex] = pivotNorm),
      orthonormalBasis.push(
        scaledColumns[pivotIndex].map((value) => value / pivotNorm),
      ),
      projections.push(
        dotProductWithCompensation(
          orthonormalBasis[pivotIndex],
          scaledResponses,
        ),
      ),
      rank++);
    for (
      let laterColumnIndex = pivotIndex + 1;
      laterColumnIndex < predictorCount;
      laterColumnIndex++
    )
      for (let passIndex = 0; passIndex < 2; passIndex++) {
        const projectionCoefficient = dotProductWithCompensation(
          orthonormalBasis[pivotIndex],
          scaledColumns[laterColumnIndex],
        );
        ((triangularFactor[pivotIndex][laterColumnIndex] +=
          projectionCoefficient),
          (scaledColumns[laterColumnIndex] = scaledColumns[
            laterColumnIndex
          ].map(
            (value, elementIndex) =>
              value -
              projectionCoefficient *
                orthonormalBasis[pivotIndex][elementIndex],
          )));
      }
  }
  const solveUpperTriangular = (rhsVector) => {
      const solution = Array(rank).fill(0);
      for (let rowIndex = rank - 1; rowIndex >= 0; rowIndex--) {
        let residual = rhsVector[rowIndex];
        for (let columnIndex = rowIndex + 1; columnIndex < rank; columnIndex++)
          residual -=
            triangularFactor[rowIndex][columnIndex] * solution[columnIndex];
        solution[rowIndex] = toFiniteNumber(
          residual / triangularFactor[rowIndex][rowIndex],
        );
      }
      return solution;
    },
    scaledCoefficients = Array(predictorCount).fill(0);
  solveUpperTriangular(projections).forEach(
    (scaledCoefficient, pivotPosition) =>
      (scaledCoefficients[columnOrder[pivotPosition]] = scaledCoefficient),
  );
  const coefficients = scaledCoefficients.map(
      (scaledCoefficient, predictorIndex) =>
        toFiniteNumber(
          scaledCoefficient * (responseScale / predictorScales[predictorIndex]),
        ),
    ),
    interceptValue = hasIntercept
      ? toFiniteNumber(
          responseMean -
            dotProductWithCompensation(coefficients, predictorMeans),
        )
      : 0,
    exponentiate = (exponent) => {
      const expValue = toFiniteNumber(Math.exp(exponent));
      return expValue > 0 ? expValue : fail("#NUM!");
    };
  if (isTrendOrGrowth) {
    const newXMatrix =
      args[2] == null ? knownXs : readNumericMatrix(context.retrieve(args[2]));
    let predictionRows, outputRowCount, outputColumnCount;
    if (predictorCount === 1)
      ((predictionRows = newXMatrix.flat().map((value) => [value])),
        (outputRowCount = newXMatrix.length),
        (outputColumnCount = newXMatrix[0].length));
    else if (!isSingleRowY && newXMatrix[0].length === predictorCount)
      ((predictionRows = newXMatrix),
        (outputRowCount = predictionRows.length),
        (outputColumnCount = 1));
    else if (isSingleRowY && newXMatrix.length === predictorCount)
      ((predictionRows = Array.from(
        { length: newXMatrix[0].length },
        (Te, columnIndex) => newXMatrix.map((row) => row[columnIndex]),
      )),
        (outputRowCount = 1),
        (outputColumnCount = predictionRows.length));
    else return fail("#REF!");
    outputRowCount * outputColumnCount > regressionSizeLimits.output &&
      fail("#LIMIT!");
    const predictions = predictionRows.map((predictionRow) => {
      const predictedValue = toFiniteNumber(
        responseMean +
          dotProductWithCompensation(
            scaledCoefficients,
            predictionRow.map(
              (newXValue, predictorIndex) =>
                toFiniteNumber(newXValue - predictorMeans[predictorIndex]) /
                predictorScales[predictorIndex],
            ),
          ) *
            responseScale,
      );
      return isLogarithmic ? exponentiate(predictedValue) : predictedValue;
    });
    return Array.from({ length: outputRowCount }, (Te, outputRowIndex) =>
      predictions.slice(
        outputRowIndex * outputColumnCount,
        (outputRowIndex + 1) * outputColumnCount,
      ),
    );
  }
  const coefficientRow = [...coefficients]
    .reverse()
    .concat(interceptValue)
    .map((value) => (isLogarithmic ? exponentiate(value) : value));
  if (!hasStatistics) return [coefficientRow];
  const scaledResiduals = scaledResponses.map(
    (scaledResponse, observationIndex) =>
      scaledResponse -
      dotProductWithCompensation(
        scaledCoefficients,
        initialScaledColumns.map(
          (scaledColumn) => scaledColumn[observationIndex],
        ),
      ),
  );
  let residualSumSquares = dotProductWithCompensation(
    scaledResiduals,
    scaledResiduals,
  );
  const totalSumSquares = dotProductWithCompensation(
    scaledResponses,
    scaledResponses,
  );
  residualSumSquares <= 1e-26 * Math.max(totalSumSquares, Number.MIN_VALUE) &&
    (residualSumSquares = 0);
  const residualDegrees = observationCount - rank - Number(hasIntercept),
    standardError =
      residualDegrees > 0
        ? Math.sqrt(residualSumSquares / residualDegrees)
        : null,
    notAvailableError = () => context.issue("#N/A"),
    divisionByZeroError = () => context.issue("#DIV/0!"),
    inverseColumns = Array.from({ length: rank }, (se, unitIndex) =>
      solveUpperTriangular(
        Array.from({ length: rank }, (Ee, elementIndex) =>
          elementIndex === unitIndex ? 1 : 0,
        ),
      ),
    ),
    coefficientStandardErrors = Array(predictorCount).fill(0);
  for (let rankIndex = 0; rankIndex < rank; rankIndex++)
    coefficientStandardErrors[columnOrder[rankIndex]] =
      standardError === null
        ? divisionByZeroError()
        : toFiniteNumber(
            Math.hypot(
              ...inverseColumns.map(
                (inverseColumn) => inverseColumn[rankIndex],
              ),
            ) *
              standardError *
              (responseScale / predictorScales[columnOrder[rankIndex]]),
          );
  const scaledMeans = columnOrder
      .slice(0, rank)
      .map(
        (predictorIndex) =>
          predictorMeans[predictorIndex] / predictorScales[predictorIndex],
      ),
    interceptVariance =
      1 / observationCount +
      inverseColumns.reduce(
        (total, inverseVector) =>
          total + dotProductWithCompensation(scaledMeans, inverseVector) ** 2,
        0,
      ),
    standardErrorRow = [...coefficientStandardErrors]
      .reverse()
      .concat(
        hasIntercept
          ? standardError === null
            ? divisionByZeroError()
            : toFiniteNumber(
                Math.sqrt(interceptVariance) * standardError * responseScale,
              )
          : notAvailableError(),
      ),
    regressionSumSquares = Math.max(0, totalSumSquares - residualSumSquares),
    unscaleSumSquares = (scaledSumSquares) =>
      toFiniteNumber(scaledSumSquares * responseScale * responseScale),
    buildStatsRow = (firstValue, secondValue) => [
      firstValue,
      secondValue,
      ...Array.from({ length: predictorCount - 1 }, notAvailableError),
    ];
  return [
    coefficientRow,
    standardErrorRow,
    buildStatsRow(
      totalSumSquares
        ? Math.max(0, Math.min(1, regressionSumSquares / totalSumSquares))
        : divisionByZeroError(),
      standardError === null
        ? divisionByZeroError()
        : toFiniteNumber(standardError * responseScale),
    ),
    buildStatsRow(
      residualDegrees > 0 && rank && residualSumSquares
        ? toFiniteNumber(
            regressionSumSquares /
              rank /
              (residualSumSquares / residualDegrees),
          )
        : divisionByZeroError(),
      residualDegrees,
    ),
    buildStatsRow(
      unscaleSumSquares(regressionSumSquares),
      unscaleSumSquares(residualSumSquares),
    ),
  ];
}
const significanceTestFunctionNames = ["T.TEST", "Z.TEST"],
  notInferenceSentinel = Symbol("not-inference");
function evaluateSignificanceTestFunction(functionName, args, context) {
  if (!significanceTestFunctionNames.includes(functionName))
    return notInferenceSentinel;
  const fail = (errorCode = "#VALUE!") => {
      throw context.issue(errorCode);
    },
    requireFinite = (candidate) =>
      Number.isFinite(candidate) ? candidate : fail("#NUM!");
  ((functionName === "T.TEST" && args.length !== 4) ||
    (functionName === "Z.TEST" && (args.length < 2 || args.length > 3))) &&
    fail();
  const retrievedArgs = args.map(context.retrieve);
  let valueCount = 0;
  const flattenValues = (element) => {
      if (Array.isArray(element)) return element.flatMap(flattenValues);
      if ((++valueCount > 4e4 && fail("#LIMIT!"), context.isIssue(element)))
        throw element;
      return (typeof element == "number" && requireFinite(element), [element]);
    },
    toNumberArgument = (argument) => {
      if (context.isIssue(argument)) throw argument;
      if (
        Array.isArray(argument) ||
        argument == null ||
        (typeof argument == "string" && !argument.trim())
      )
        return fail();
      const numericArgument = Number(argument);
      return Number.isFinite(numericArgument) ? numericArgument : fail();
    },
    onlyNumbers = (rawValues) =>
      rawValues.filter((entry) => typeof entry == "number"),
    meanAndVariance = (sample) => {
      if (sample.length < 2) return fail("#DIV/0!");
      const mean = exactAverageOfNumbers(sample);
      return {
        mean: mean,
        variance: requireFinite(
          sumDecimalPartsToNumber(
            sample.map((sampleValue) => (sampleValue - mean) ** 2),
          ) /
            (sample.length - 1),
        ),
      };
    },
    toProbability = (probability) => {
      if (context.isIssue(probability)) throw probability;
      return typeof probability != "number" ||
        !Number.isFinite(probability) ||
        probability < -1e-12 ||
        probability > 1 + 1e-12
        ? fail("#NUM!")
        : Math.max(0, Math.min(1, probability));
    },
    firstSample = flattenValues(retrievedArgs[0]);
  if (functionName === "Z.TEST") {
    const numericSample = onlyNumbers(firstSample),
      sampleCount = numericSample.length;
    if (!sampleCount) return fail("#N/A");
    const testValue = toNumberArgument(retrievedArgs[1]),
      scale =
        Math.max(Math.abs(testValue), ...numericSample.map(Math.abs)) || 1,
      scaledSample = numericSample.map((numericItem) => numericItem / scale),
      scaledSigma =
        args[2] == null
          ? Math.sqrt(meanAndVariance(scaledSample).variance)
          : toNumberArgument(retrievedArgs[2]) / scale;
    if (scaledSigma < 0) return fail("#NUM!");
    if (!scaledSigma) return fail("#DIV/0!");
    const zStatistic = requireFinite(
        ((exactAverageOfNumbers(scaledSample) - testValue / scale) /
          scaledSigma) *
          Math.sqrt(sampleCount),
      ),
      erfcResult = context.distribution("ERFC", [zStatistic / Math.SQRT2]);
    if (context.isIssue(erfcResult)) throw erfcResult;
    return typeof erfcResult == "number"
      ? toProbability(erfcResult / 2)
      : fail("#NUM!");
  }
  const secondSample = flattenValues(retrievedArgs[1]),
    tails = Math.trunc(toNumberArgument(retrievedArgs[2])),
    testType = Math.trunc(toNumberArgument(retrievedArgs[3]));
  ((![1, 2].includes(tails) || ![1, 2, 3].includes(testType)) && fail("#NUM!"),
    testType === 1 &&
      firstSample.length !== secondSample.length &&
      fail("#N/A"));
  let firstValues, secondValues;
  if (testType === 1) {
    const pairs = firstSample
      .map((firstItem, firstIndex) => [firstItem, secondSample[firstIndex]])
      .filter((candidatePair) =>
        candidatePair.every((pairValue) => typeof pairValue == "number"),
      );
    ((firstValues = pairs.map((pair) => pair[0])),
      (secondValues = pairs.map((pair) => pair[1])));
  } else
    ((firstValues = onlyNumbers(firstSample)),
      (secondValues = onlyNumbers(secondSample)));
  (firstValues.length < 2 || secondValues.length < 2) && fail("#DIV/0!");
  const normalizer =
    Math.max(...firstValues.map(Math.abs), ...secondValues.map(Math.abs)) || 1;
  ((firstValues = firstValues.map((sampleNumber) => sampleNumber / normalizer)),
    (secondValues = secondValues.map(
      (sampleNumber) => sampleNumber / normalizer,
    )));
  let meanDifference, squaredStandardError, degreesOfFreedom;
  if (testType === 1) {
    const differenceStats = meanAndVariance(
      firstValues.map(
        (firstNumber, valueIndex) => firstNumber - secondValues[valueIndex],
      ),
    );
    ((meanDifference = differenceStats.mean),
      (squaredStandardError = differenceStats.variance / firstValues.length),
      (degreesOfFreedom = firstValues.length - 1));
  } else {
    const firstStats = meanAndVariance(firstValues),
      secondStats = meanAndVariance(secondValues);
    if (((meanDifference = firstStats.mean - secondStats.mean), testType === 2))
      ((degreesOfFreedom = firstValues.length + secondValues.length - 2),
        (squaredStandardError =
          (((firstValues.length - 1) * firstStats.variance +
            (secondValues.length - 1) * secondStats.variance) /
            degreesOfFreedom) *
          (1 / firstValues.length + 1 / secondValues.length)));
    else {
      const firstMeanVariance = firstStats.variance / firstValues.length,
        secondMeanVariance = secondStats.variance / secondValues.length;
      ((squaredStandardError = firstMeanVariance + secondMeanVariance),
        (degreesOfFreedom =
          squaredStandardError ** 2 /
          (firstMeanVariance ** 2 / (firstValues.length - 1) +
            secondMeanVariance ** 2 / (secondValues.length - 1))));
    }
  }
  squaredStandardError || fail("#DIV/0!");
  const tStatistic = requireFinite(
    Math.abs(meanDifference) / Math.sqrt(squaredStandardError),
  );
  requireFinite(degreesOfFreedom);
  const betaProbability = context.distribution("BETA.DIST", [
    degreesOfFreedom / (degreesOfFreedom + tStatistic * tStatistic),
    degreesOfFreedom / 2,
    0.5,
    true,
  ]);
  return (toProbability(betaProbability) * tails) / 2;
}
const specialistFunctionNames = ["CONVERT", "SLOPE", "RSQ", "STEYX", "PEARSON"],
  notSpecialistSentinel = Symbol("not-specialist"),
  conversionUnits = [],
  addConversionUnitSymbols = (
    symbolList,
    dimension,
    factor,
    metric = false,
    offset = 0,
    power = 1,
  ) => {
    for (const symbol of symbolList.split("|"))
      conversionUnits.push({
        symbol: symbol,
        dimension: dimension,
        factor: factor,
        metric: metric,
        offset: offset,
        power: power,
      });
  };
addConversionUnitSymbols("g", "mass", 0.001, true);
addConversionUnitSymbols("lbm", "mass", 0.45359237);
addConversionUnitSymbols("ozm", "mass", 0.45359237 / 16);
addConversionUnitSymbols("grain", "mass", 0.45359237 / 7e3);
addConversionUnitSymbols("stone", "mass", 0.45359237 * 14);
addConversionUnitSymbols("ton", "mass", 0.45359237 * 2e3);
addConversionUnitSymbols("uk_ton|LTON|brton", "mass", 0.45359237 * 2240);
addConversionUnitSymbols("cwt|shweight", "mass", 0.45359237 * 100);
addConversionUnitSymbols("uk_cwt|lcwt|hweight", "mass", 0.45359237 * 112);
addConversionUnitSymbols("m", "length", 1, true);
addConversionUnitSymbols("in", "length", 0.0254);
addConversionUnitSymbols("ft", "length", 0.3048);
addConversionUnitSymbols("yd", "length", 0.9144);
addConversionUnitSymbols("mi", "length", 1609.344);
addConversionUnitSymbols("Nmi", "length", 1852);
addConversionUnitSymbols("ang", "length", 1e-10, true);
addConversionUnitSymbols("Picapt|Pica", "length", 0.0254 / 72);
addConversionUnitSymbols("pica", "length", 0.0254 / 6);
addConversionUnitSymbols("sec|s", "time", 1, true);
addConversionUnitSymbols("mn|min", "time", 60);
addConversionUnitSymbols("hr", "time", 3600);
addConversionUnitSymbols("day|d", "time", 86400);
addConversionUnitSymbols("yr", "time", 365.25 * 86400);
addConversionUnitSymbols("Pa|p", "pressure", 1, true);
addConversionUnitSymbols("atm|at", "pressure", 101325);
addConversionUnitSymbols("psi", "pressure", 6894.757293168361);
addConversionUnitSymbols("N", "force", 1, true);
addConversionUnitSymbols("dyn|dy", "force", 1e-5, true);
addConversionUnitSymbols("lbf", "force", 4.4482216152605);
addConversionUnitSymbols("J", "energy", 1, true);
addConversionUnitSymbols("e", "energy", 1e-7, true);
addConversionUnitSymbols("c", "energy", 4.184, true);
addConversionUnitSymbols("cal", "energy", 4.1868, true);
addConversionUnitSymbols("Wh|wh", "energy", 3600, true);
addConversionUnitSymbols("BTU|btu", "energy", 1055.05585262);
addConversionUnitSymbols("W|w", "power", 1, true);
addConversionUnitSymbols("HP|h", "power", 745.6998715822702);
addConversionUnitSymbols("PS", "power", 735.49875);
addConversionUnitSymbols("T", "magnetism", 1, true);
addConversionUnitSymbols("ga", "magnetism", 1e-4, true);
addConversionUnitSymbols("K|kel", "temperature", 1, true);
addConversionUnitSymbols("C|cel", "temperature", 1, false, 273.15);
addConversionUnitSymbols(
  "F|fah",
  "temperature",
  5 / 9,
  false,
  273.15 - 160 / 9,
);
addConversionUnitSymbols("Rank", "temperature", 5 / 9);
addConversionUnitSymbols("Reau", "temperature", 1.25, false, 273.15);
addConversionUnitSymbols("l|L|lt", "volume", 0.001, true);
addConversionUnitSymbols("gal", "volume", 0.003785411784);
addConversionUnitSymbols("qt", "volume", 0.003785411784 / 4);
addConversionUnitSymbols("pt|us_pt", "volume", 0.003785411784 / 8);
addConversionUnitSymbols("cup", "volume", 0.003785411784 / 16);
addConversionUnitSymbols("oz", "volume", 0.003785411784 / 128);
addConversionUnitSymbols("tbs", "volume", 0.003785411784 / 256);
addConversionUnitSymbols("tsp", "volume", 0.003785411784 / 768);
addConversionUnitSymbols("tspm", "volume", 5e-6);
addConversionUnitSymbols("uk_gal", "volume", 0.00454609);
addConversionUnitSymbols("uk_qt", "volume", 0.00454609 / 4);
addConversionUnitSymbols("uk_pt", "volume", 0.00454609 / 8);
addConversionUnitSymbols("barrel", "volume", 0.003785411784 * 42);
for (const [unitSymbol, unitFactor] of [
  ["m", 1],
  ["in", 0.0254],
  ["ft", 0.3048],
  ["yd", 0.9144],
  ["mi", 1609.344],
  ["Nmi", 1852],
  ["ang", 1e-10],
])
  (addConversionUnitSymbols(
    `${unitSymbol}2|${unitSymbol}^2`,
    "area",
    unitFactor ** 2,
    unitSymbol === "m" || unitSymbol === "ang",
    0,
    2,
  ),
    addConversionUnitSymbols(
      `${unitSymbol}3|${unitSymbol}^3`,
      "volume",
      unitFactor ** 3,
      unitSymbol === "m" || unitSymbol === "ang",
      0,
      3,
    ));
addConversionUnitSymbols("ar", "area", 100, true);
addConversionUnitSymbols("ha", "area", 1e4);
addConversionUnitSymbols("uk_acre", "area", 4046.8564224);
addConversionUnitSymbols("us_acre", "area", (1200 / 3937) ** 2 * 43560);
addConversionUnitSymbols("bit", "information", 1, true);
addConversionUnitSymbols("byte", "information", 8, true);
addConversionUnitSymbols("m/s|m/sec", "speed", 1, true);
addConversionUnitSymbols("m/h|m/hr", "speed", 1 / 3600, true);
addConversionUnitSymbols("mph", "speed", 1609.344 / 3600);
addConversionUnitSymbols("kn", "speed", 1852 / 3600);
const frozenConversionUnits = Object.freeze(
    conversionUnits.map((conversionUnit) => Object.freeze(conversionUnit)),
  ),
  siPrefixPowersOfTen = {
    Y: 24,
    Z: 21,
    E: 18,
    P: 15,
    T: 12,
    G: 9,
    M: 6,
    k: 3,
    h: 2,
    da: 1,
    e: 1,
    d: -1,
    c: -2,
    m: -3,
    u: -6,
    n: -9,
    p: -12,
    f: -15,
    a: -18,
    z: -21,
    y: -24,
  },
  binaryPrefixExponents = {
    ki: 10,
    Mi: 20,
    Gi: 30,
    Ti: 40,
    Pi: 50,
    Ei: 60,
    Zi: 70,
    Yi: 80,
  };
function findConversionUnit(symbol) {
  const conversionUnit = conversionUnits.find(
    (conversionUnit) => conversionUnit.symbol === symbol,
  );
  if (conversionUnit) return { ...conversionUnit };
  for (const [binaryPrefix, binaryExponent] of Object.entries(
    binaryPrefixExponents,
  )) {
    if (!symbol.startsWith(binaryPrefix)) continue;
    const conversionUnit = conversionUnits.find(
      (conversionUnit) =>
        conversionUnit.symbol === symbol.slice(binaryPrefix.length) &&
        conversionUnit.dimension === "information",
    );
    if (conversionUnit)
      return {
        ...conversionUnit,
        symbol: symbol,
        factor: conversionUnit.factor * 2 ** binaryExponent,
      };
  }
  for (const siPrefix of Object.keys(siPrefixPowersOfTen).sort(
    (leftPrefix, rightPrefix) => rightPrefix.length - leftPrefix.length,
  )) {
    if (!symbol.startsWith(siPrefix)) continue;
    const conversionUnit = conversionUnits.find(
      (conversionUnit) =>
        conversionUnit.symbol === symbol.slice(siPrefix.length) &&
        conversionUnit.metric,
    );
    if (conversionUnit)
      return {
        ...conversionUnit,
        symbol: symbol,
        factor:
          conversionUnit.factor *
          10 ** (siPrefixPowersOfTen[siPrefix] * conversionUnit.power),
      };
  }
}
function evaluateSpecialistFunction(functionName, args, context) {
  if (!specialistFunctionNames.includes(functionName))
    return notSpecialistSentinel;
  const fail = (errorCode = "#VALUE!") => {
    throw context.issue(errorCode);
  };
  (args.length !== (functionName === "CONVERT" ? 3 : 2) ||
    args.some((argument) => argument == null)) &&
    fail();
  const resolvedArgs = args.map(context.retrieve);
  let visitedCount = 0;
  const checkValue = (value) => {
    if (Array.isArray(value)) {
      for (const element of value) checkValue(element);
      return;
    }
    if ((++visitedCount > 4e4 && fail("#LIMIT!"), context.isIssue(value)))
      throw value;
    typeof value == "number" && !Number.isFinite(value) && fail("#NUM!");
  };
  resolvedArgs.forEach(checkValue);
  const requireFinite = (numericValue) =>
    Number.isFinite(numericValue) ? numericValue : fail("#NUM!");
  if (functionName === "CONVERT") {
    const [amount, symbol, targetSymbol] = resolvedArgs;
    (typeof symbol != "string" ||
      typeof targetSymbol != "string" ||
      Array.isArray(amount) ||
      amount == null ||
      (typeof amount == "string" && !amount.trim())) &&
      fail();
    const numericAmount = Number(amount);
    Number.isFinite(numericAmount) || fail();
    const conversionUnit = findConversionUnit(symbol),
      targetUnit = findConversionUnit(targetSymbol);
    return (
      (!conversionUnit ||
        !targetUnit ||
        conversionUnit.dimension !== targetUnit.dimension) &&
        fail("#N/A"),
      conversionUnit.offset === targetUnit.offset
        ? requireFinite(
            numericAmount * (conversionUnit.factor / targetUnit.factor),
          )
        : requireFinite(
            (numericAmount * conversionUnit.factor +
              conversionUnit.offset -
              targetUnit.offset) /
              targetUnit.factor,
          )
    );
  }
  const flattenToArray = (input) =>
      Array.isArray(input) ? input.flat(1 / 0) : [input],
    knownYValues = flattenToArray(resolvedArgs[0]),
    knownXValues = flattenToArray(resolvedArgs[1]);
  knownYValues.length !== knownXValues.length && fail("#N/A");
  const numericPairs = knownYValues
      .map((knownY, pairIndex) => [knownY, knownXValues[pairIndex]])
      .filter((pair) =>
        pair.every((pairValue) => typeof pairValue == "number"),
      ),
    pairCount = numericPairs.length;
  (!pairCount && functionName === "SLOPE" && fail("#N/A"),
    pairCount < (functionName === "STEYX" ? 3 : 2) && fail("#DIV/0!"));
  const pairedYs = numericPairs.map((pair) => pair[0]),
    pairedXs = numericPairs.map((pair) => pair[1]),
    meanY = exactAverageOfNumbers(pairedYs),
    meanX = exactAverageOfNumbers(pairedXs),
    xDeviations = pairedXs.map((xValue) => requireFinite(xValue - meanX)),
    yDeviations = pairedYs.map((yValue) => requireFinite(yValue - meanY)),
    maxXDeviation = Math.max(...xDeviations.map(Math.abs)),
    maxYDeviation = Math.max(...yDeviations.map(Math.abs));
  if (
    ((!maxXDeviation ||
      (!maxYDeviation &&
        (functionName === "RSQ" || functionName === "PEARSON"))) &&
      fail("#DIV/0!"),
    !maxYDeviation)
  )
    return 0;
  const scaledXDeviations = xDeviations.map(
      (xDeviation) => xDeviation / maxXDeviation,
    ),
    scaledYDeviations = yDeviations.map(
      (yDeviation) => yDeviation / maxYDeviation,
    ),
    sumSquaredX = sumDecimalPartsToNumber(
      scaledXDeviations.map(
        (scaledXDeviation) => scaledXDeviation * scaledXDeviation,
      ),
    ),
    sumSquaredY = sumDecimalPartsToNumber(
      scaledYDeviations.map(
        (scaledYDeviation) => scaledYDeviation * scaledYDeviation,
      ),
    ),
    crossProductSum = sumDecimalPartsToNumber(
      scaledXDeviations.map(
        (scaledXDeviation, valueIndex) =>
          scaledXDeviation * scaledYDeviations[valueIndex],
      ),
    );
  return functionName === "SLOPE"
    ? requireFinite(
        (crossProductSum / sumSquaredX) * (maxYDeviation / maxXDeviation),
      )
    : functionName === "PEARSON"
      ? Math.max(
          -1,
          Math.min(1, crossProductSum / Math.sqrt(sumSquaredX * sumSquaredY)),
        )
      : functionName === "RSQ"
        ? Math.min(
            1,
            Math.max(
              0,
              (crossProductSum / sumSquaredX) * (crossProductSum / sumSquaredY),
            ),
          )
        : requireFinite(
            Math.sqrt(
              sumDecimalPartsToNumber(
                scaledYDeviations.map(
                  (scaledYDeviation, valueIndex) =>
                    (scaledYDeviation -
                      (crossProductSum / sumSquaredX) *
                        scaledXDeviations[valueIndex]) **
                    2,
                ),
              ) /
                (pairCount - 2),
            ) * maxYDeviation,
          );
}
const statisticalFunctionNames = [
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "AVERAGE",
    "MIN",
    "MAX",
    "MEDIAN",
    "MODE",
    "MODE.SNGL",
    "STDEV.S",
    "STDEV.P",
    "VAR.S",
    "VAR.P",
    "PERCENTILE",
    "PERCENTILE.INC",
    "PERCENTILE.EXC",
    "QUARTILE",
    "QUARTILE.INC",
    "QUARTILE.EXC",
    "RANK",
    "RANK.EQ",
    "RANK.AVG",
    "AVERAGEA",
    "MAXA",
    "MINA",
    "STDEVA",
    "STDEVPA",
    "VARA",
    "VARPA",
    "SKEW",
    "SKEW.P",
    "TRIMMEAN",
    "MODE.MULT",
    "PERMUT",
    "PERMUTATIONA",
  ],
  notStatisticalSentinel = Symbol("not-statistical-function"),
  isReferenceArgument = (value) =>
    !!value && typeof value == "object" && "ref" in value;
function sumWithKahanCompensation(values) {
  let sum = 0,
    compensation = 0;
  for (const value of values) {
    const nextSum = sum + value;
    ((compensation +=
      Math.abs(sum) >= Math.abs(value)
        ? sum - nextSum + value
        : value - nextSum + sum),
      (sum = nextSum));
  }
  return sum + compensation;
}
function evaluateStatisticalFunction(
  functionName,
  argumentList,
  evaluationContext,
) {
  if (!statisticalFunctionNames.includes(functionName))
    return notStatisticalSentinel;
  const raiseError = (errorCode = "#VALUE!") => {
      throw evaluationContext.issue(errorCode);
    },
    isRankFunction = functionName.startsWith("RANK"),
    isPercentileFunction = functionName.startsWith("PERCENTILE"),
    isQuartileFunction = functionName.startsWith("QUARTILE"),
    isTrimmeanFunction = functionName === "TRIMMEAN",
    isPermutFunction =
      functionName === "PERMUT" || functionName === "PERMUTATIONA",
    isAVariantFunction = [
      "AVERAGEA",
      "MAXA",
      "MINA",
      "STDEVA",
      "STDEVPA",
      "VARA",
      "VARPA",
    ].includes(functionName);
  if (
    !argumentList.length ||
    argumentList.length > 255 ||
    (functionName === "COUNTBLANK" && argumentList.length !== 1) ||
    ((isPercentileFunction ||
      isQuartileFunction ||
      isTrimmeanFunction ||
      isPermutFunction) &&
      argumentList.length !== 2) ||
    (isRankFunction && (argumentList.length < 2 || argumentList.length > 3))
  )
    return raiseError();
  const argumentValues = argumentList.map(evaluationContext.retrieve);
  let visitedValueCount = 0;
  const visitValue = (visitedValue) => {
    if (Array.isArray(visitedValue)) {
      for (const nestedItem of visitedValue) visitValue(nestedItem);
      return;
    }
    if (
      (++visitedValueCount > 1e5 && raiseError("#LIMIT!"),
      evaluationContext.isIssue(visitedValue) &&
        ["#LIMIT!", "#CYCLE!", "#BLOCKED!", "#ERROR!"].includes(
          String(visitedValue.error),
        ))
    )
      throw visitedValue;
  };
  for (const argumentValue of argumentValues) visitValue(argumentValue);
  const requireFiniteNumber = (numericInput) =>
      Number.isFinite(numericInput) ? numericInput : raiseError("#NUM!"),
    coerceNumericArgument = (rawValue, fallbackValue) => {
      if (rawValue === undefined && fallbackValue !== undefined)
        return fallbackValue;
      if (evaluationContext.isIssue(rawValue)) throw rawValue;
      if (rawValue == null) return 0;
      if (
        Array.isArray(rawValue) ||
        (typeof rawValue == "string" && !rawValue.trim())
      )
        return raiseError();
      const parsedNumber = Number(rawValue);
      return Number.isFinite(parsedNumber) ? parsedNumber : raiseError();
    };
  if (functionName === "COUNTA")
    return argumentValues
      .flatMap(flattenNestedValues)
      .filter((flatValue) => flatValue != null).length;
  if (functionName === "COUNTBLANK")
    return argumentValues
      .flatMap(flattenNestedValues)
      .filter((flatCell) => flatCell == null || flatCell === "").length;
  const primaryIndex = isRankFunction ? 1 : 0;
  if (isPermutFunction) {
    const totalCount = Math.trunc(coerceNumericArgument(argumentValues[0])),
      chosenCount = Math.trunc(coerceNumericArgument(argumentValues[1]));
    if (
      totalCount < 0 ||
      chosenCount < 0 ||
      (functionName === "PERMUT" && chosenCount > totalCount) ||
      (!totalCount && chosenCount > 0)
    )
      return raiseError("#NUM!");
    if (!chosenCount) return 1;
    if (functionName === "PERMUTATIONA")
      return requireFiniteNumber(totalCount ** chosenCount);
    if (chosenCount > 170) return raiseError("#NUM!");
    let permutationCount = 1;
    for (let termIndex = 0; termIndex < chosenCount; termIndex++)
      permutationCount = requireFiniteNumber(
        permutationCount * (totalCount - termIndex),
      );
    return permutationCount;
  }
  const selectedValues =
      isRankFunction ||
      isPercentileFunction ||
      isQuartileFunction ||
      isTrimmeanFunction
        ? [argumentValues[primaryIndex]]
        : argumentValues,
    numericValues = [];
  for (const [valueIndex, selectedValue] of selectedValues.entries()) {
    const isRangeArgument =
      Array.isArray(selectedValue) ||
      isReferenceArgument(
        argumentList[
          isRankFunction ||
          isPercentileFunction ||
          isQuartileFunction ||
          isTrimmeanFunction
            ? primaryIndex
            : valueIndex
        ],
      );
    for (const cellValue of flattenNestedValues(selectedValue)) {
      if (evaluationContext.isIssue(cellValue)) {
        if (functionName === "COUNT") continue;
        throw cellValue;
      }
      if (typeof cellValue == "number")
        numericValues.push(requireFiniteNumber(cellValue));
      else if (isAVariantFunction && isRangeArgument && cellValue != null)
        if (typeof cellValue == "boolean")
          numericValues.push(Number(cellValue));
        else if (typeof cellValue == "string") numericValues.push(0);
        else return raiseError();
      else if (!isRangeArgument && cellValue != null) {
        if (typeof cellValue == "boolean")
          numericValues.push(Number(cellValue));
        else if (
          typeof cellValue == "string" &&
          cellValue.trim() &&
          Number.isFinite(Number(cellValue))
        )
          numericValues.push(Number(cellValue));
        else if (functionName !== "COUNT") return raiseError();
      }
    }
  }
  if (functionName === "COUNT") return numericValues.length;
  if (functionName === "AVERAGE" || functionName === "AVERAGEA")
    return numericValues.length
      ? requireFiniteNumber(exactAverageOfNumbers(numericValues))
      : raiseError("#DIV/0!");
  if (functionName === "MIN" || functionName === "MINA")
    return numericValues.length
      ? numericValues.reduce((smallest, candidate) =>
          Math.min(smallest, candidate),
        )
      : 0;
  if (functionName === "MAX" || functionName === "MAXA")
    return numericValues.length
      ? numericValues.reduce((largest, nextValue) =>
          Math.max(largest, nextValue),
        )
      : 0;
  if (functionName.startsWith("STDEV") || functionName.startsWith("VAR")) {
    const isSampleStatistic =
      functionName.endsWith(".S") ||
      functionName === "STDEVA" ||
      functionName === "VARA";
    if (numericValues.length < (isSampleStatistic ? 2 : 1))
      return raiseError("#DIV/0!");
    const averageValue = exactAverageOfNumbers(numericValues);
    let deviations = numericValues.map(
        (numericItem) => numericItem - averageValue,
      ),
      scaleFactor = deviations.reduce(
        (largestAbs, deviation) => Math.max(largestAbs, Math.abs(deviation)),
        0,
      );
    if (
      (Number.isFinite(scaleFactor)
        ? scaleFactor &&
          (deviations = deviations.map(
            (deviationItem) => deviationItem / scaleFactor,
          ))
        : ((scaleFactor = numericValues.reduce(
            (largestAbsSoFar, valueItem) =>
              Math.max(largestAbsSoFar, Math.abs(valueItem)),
            0,
          )),
          (deviations = numericValues.map(
            (numericEntry) =>
              numericEntry / scaleFactor - averageValue / scaleFactor,
          ))),
      !scaleFactor)
    )
      return 0;
    const variance =
        sumWithKahanCompensation(
          deviations.map(
            (scaledDeviation) => scaledDeviation * scaledDeviation,
          ),
        ) /
        (numericValues.length - (isSampleStatistic ? 1 : 0)),
      standardDeviation = Math.sqrt(variance) * scaleFactor;
    return requireFiniteNumber(
      functionName.startsWith("STDEV")
        ? standardDeviation
        : standardDeviation * standardDeviation,
    );
  }
  if (functionName === "SKEW" || functionName === "SKEW.P") {
    const numericCount = numericValues.length;
    if (numericCount < 3) return raiseError("#DIV/0!");
    const meanValue = exactAverageOfNumbers(numericValues),
      maxAbsValue = numericValues.reduce(
        (accumulatedMax, absItem) =>
          Math.max(accumulatedMax, Math.abs(absItem)),
        0,
      );
    if (!maxAbsValue) return raiseError("#DIV/0!");
    const centeredValues = numericValues.map(
        (skewItem) => skewItem / maxAbsValue - meanValue / maxAbsValue,
      ),
      sumOfSquares = sumWithKahanCompensation(
        centeredValues.map((centeredItem) => centeredItem * centeredItem),
      );
    if (!sumOfSquares) return raiseError("#DIV/0!");
    const rawSkew =
      sumWithKahanCompensation(
        centeredValues.map((centeredEntry) => centeredEntry ** 3),
      ) /
      sumOfSquares ** 1.5;
    return requireFiniteNumber(
      functionName === "SKEW.P"
        ? rawSkew * Math.sqrt(numericCount)
        : (rawSkew * numericCount * Math.sqrt(numericCount - 1)) /
            (numericCount - 2),
    );
  }
  if (
    (numericValues.sort((leftValue, rightValue) => leftValue - rightValue),
    isTrimmeanFunction)
  ) {
    const trimFraction = coerceNumericArgument(argumentValues[1]);
    if (trimFraction < 0 || trimFraction >= 1 || !numericValues.length)
      return raiseError("#NUM!");
    const trimCount = Math.floor((numericValues.length * trimFraction) / 2);
    return requireFiniteNumber(
      exactAverageOfNumbers(
        numericValues.slice(trimCount, numericValues.length - trimCount),
      ),
    );
  }
  if (isRankFunction) {
    const rankedValue = coerceNumericArgument(argumentValues[0]),
      orderFlag = coerceNumericArgument(argumentValues[2], 0),
      firstIndex = numericValues.indexOf(rankedValue),
      lastIndex = numericValues.lastIndexOf(rankedValue);
    return firstIndex < 0
      ? raiseError("#N/A")
      : orderFlag
        ? 1 +
          (functionName === "RANK.AVG"
            ? (firstIndex + lastIndex) / 2
            : firstIndex)
        : numericValues.length -
          (functionName === "RANK.AVG"
            ? (firstIndex + lastIndex) / 2
            : lastIndex);
  }
  if (functionName === "MODE.MULT") {
    const frequencyByValue = new Map();
    let maxFrequency = 1;
    for (const numericValue of numericValues) {
      const frequency = (frequencyByValue.get(numericValue) ?? 0) + 1;
      (frequencyByValue.set(numericValue, frequency),
        (maxFrequency = Math.max(maxFrequency, frequency)));
    }
    if (maxFrequency === 1) return raiseError("#N/A");
    const modeValues = [...frequencyByValue]
      .filter(([, entryCount]) => entryCount === maxFrequency)
      .map(([modeValue]) => [modeValue]);
    return modeValues.length > 1e4 ? raiseError("#LIMIT!") : modeValues;
  }
  if (functionName === "MODE" || functionName === "MODE.SNGL") {
    let bestValue = 0,
      bestCount = 1,
      runLength = 0,
      previousValue;
    for (const currentValue of numericValues)
      ((runLength = currentValue === previousValue ? runLength + 1 : 1),
        runLength > bestCount &&
          ((bestCount = runLength), (bestValue = currentValue)),
        (previousValue = currentValue));
    return bestCount > 1 ? bestValue : raiseError("#N/A");
  }
  if (!numericValues.length) return raiseError("#NUM!");
  if (functionName === "MEDIAN") {
    const middleIndex = Math.floor(numericValues.length / 2);
    return numericValues.length % 2 ||
      numericValues[middleIndex - 1] === numericValues[middleIndex]
      ? numericValues[middleIndex]
      : numericValues[middleIndex - 1] / 2 + numericValues[middleIndex] / 2;
  }
  const quantileArgument = coerceNumericArgument(argumentValues[1]),
    quantile = isQuartileFunction
      ? Math.trunc(quantileArgument) / 4
      : quantileArgument,
    isExclusive = functionName.endsWith(".EXC");
  if (
    quantile < 0 ||
    quantile > 1 ||
    (isExclusive && (quantile === 0 || quantile === 1))
  )
    return raiseError("#NUM!");
  const fractionalIndex = isExclusive
    ? quantile * (numericValues.length + 1) - 1
    : quantile * (numericValues.length - 1);
  if (fractionalIndex < 0 || fractionalIndex > numericValues.length - 1)
    return raiseError("#NUM!");
  const lowerIndex = Math.floor(fractionalIndex),
    upperIndex = Math.ceil(fractionalIndex),
    interpolationWeight = fractionalIndex - lowerIndex;
  return numericValues[lowerIndex] === numericValues[upperIndex]
    ? numericValues[lowerIndex]
    : requireFiniteNumber(
        numericValues[lowerIndex] * (1 - interpolationWeight) +
          numericValues[upperIndex] * interpolationWeight,
      );
}
const ownedLookupFunctionNames = ["LOOKUP", "CHOOSE"],
  notOwnedLookupSentinel = Symbol("not an owned lookup function");
function evaluateLookupFunction(functionName, args, context) {
  if (!["VLOOKUP", "HLOOKUP", "LOOKUP", "CHOOSE"].includes(functionName))
    return notOwnedLookupSentinel;
  const fail = (errorCode = "#VALUE!") => {
      throw context.issue(errorCode);
    },
    scalarArgument = (value) => {
      if (context.isIssue(value)) throw value;
      return (Array.isArray(value) && fail(), value);
    },
    matrixArgument = (value) => {
      if (context.isIssue(value)) throw value;
      value === undefined && fail();
      const matrix = Array.isArray(value)
        ? value.every(Array.isArray)
          ? value
          : [value]
        : [[value]];
      return (
        (!matrix.length ||
          !matrix[0].length ||
          matrix.some((row) => row.length !== matrix[0].length)) &&
          fail(),
        matrix.length * matrix[0].length > 2e4 && fail("#LIMIT!"),
        matrix
      );
    },
    vectorArgument = (matrix) => (
      matrix.length !== 1 && matrix[0].length !== 1 && fail(),
      matrix.flat()
    );
  if (functionName === "CHOOSE") {
    (args.length < 2 || args.length > 255) && fail();
    const chooseIndexValue = scalarArgument(context.retrieve(args[0]));
    typeof chooseIndexValue == "string" && !chooseIndexValue.trim() && fail();
    const chooseIndex = Math.trunc(Number(chooseIndexValue));
    return (
      (!Number.isFinite(chooseIndex) ||
        chooseIndex < 1 ||
        chooseIndex >= args.length) &&
        fail(),
      args[chooseIndex] ?? 0
    );
  }
  (args.length < (functionName === "LOOKUP" ? 2 : 3) ||
    args.length > (functionName === "LOOKUP" ? 3 : 4)) &&
    fail();
  const lookupValue = scalarArgument(context.retrieve(args[0]));
  lookupValue === undefined && fail();
  const table = matrixArgument(context.retrieve(args[1]));
  let lookupVector,
    resultVector,
    isApproximate = true;
  if (functionName === "LOOKUP")
    args.length === 3
      ? ((lookupVector = vectorArgument(table)),
        (resultVector = vectorArgument(
          matrixArgument(context.retrieve(args[2])),
        )),
        lookupVector.length !== resultVector.length && fail())
      : table[0].length > table.length
        ? ((lookupVector = table[0]), (resultVector = table[table.length - 1]))
        : ((lookupVector = table.map((row) => row[0])),
          (resultVector = table.map((row) => row[row.length - 1])));
  else {
    const returnIndexValue = scalarArgument(context.retrieve(args[2]));
    typeof returnIndexValue == "string" && !returnIndexValue.trim() && fail();
    const returnIndex = Math.trunc(Number(returnIndexValue));
    ((!Number.isFinite(returnIndex) || returnIndex < 1) && fail(),
      returnIndex >
        (functionName === "VLOOKUP" ? table[0].length : table.length) &&
        fail("#REF!"));
    const rangeLookupValue =
      args[3] == null ? true : scalarArgument(context.retrieve(args[3]));
    (typeof rangeLookupValue == "string" &&
      !/^(true|false)$/i.test(rangeLookupValue) &&
      fail(),
      (isApproximate =
        typeof rangeLookupValue == "string"
          ? rangeLookupValue.toLowerCase() === "true"
          : !!rangeLookupValue),
      (lookupVector =
        functionName === "VLOOKUP" ? table.map((row) => row[0]) : table[0]),
      (resultVector =
        functionName === "VLOOKUP"
          ? table.map((row) => row[returnIndex - 1])
          : table[returnIndex - 1]));
  }
  const lookupIssue = lookupVector.find(context.isIssue);
  if (lookupIssue) throw lookupIssue;
  const normalizeCase = (value) =>
    typeof value == "string" ? value.toLowerCase() : value;
  let matchIndex = -1;
  if (isApproximate) {
    (!["number", "string", "boolean"].includes(typeof lookupValue) ||
      lookupVector.some((entry) => typeof entry != typeof lookupValue)) &&
      fail();
    const normalizedVector = lookupVector.map(normalizeCase),
      normalizedLookupValue = normalizeCase(lookupValue);
    (normalizedVector.some(
      (normalizedEntry, entryIndex) =>
        entryIndex > 0 && normalizedEntry < normalizedVector[entryIndex - 1],
    ) && fail(),
      normalizedVector.forEach((normalizedEntry, entryIndex) => {
        normalizedEntry <= normalizedLookupValue && (matchIndex = entryIndex);
      }));
  } else {
    let comparisonCost = 0;
    const hasWildcard =
      typeof lookupValue == "string" && /[?*~]/.test(lookupValue);
    matchIndex = lookupVector.findIndex((entry) =>
      hasWildcard
        ? typeof entry != "string"
          ? false
          : ((comparisonCost += entry.length * lookupValue.length),
            comparisonCost > 1e5 && fail("#LIMIT!"),
            context.matches(entry, `=${lookupValue}`))
        : normalizeCase(entry) === normalizeCase(lookupValue),
    );
  }
  return matchIndex < 0 ? fail("#N/A") : (resultVector[matchIndex] ?? 0);
}
const notOwnedLogicalSentinel = Symbol("not an owned logical function");
function evaluateLogicalFunction(functionName, args, context) {
  if (!["AND", "OR", "XOR", "NOT"].includes(functionName))
    return notOwnedLogicalSentinel;
  const fail = () => {
    throw context.issue("#VALUE!");
  };
  (!args.length ||
    args.length > 255 ||
    (functionName === "NOT" && args.length !== 1)) &&
    fail();
  const booleans = [];
  let visitCount = 0;
  const collectBooleans = (value, isFromReference) => {
    if (++visitCount > 1e5) throw context.issue("#LIMIT!");
    if (context.isIssue(value)) throw value;
    if (Array.isArray(value))
      (functionName === "NOT" && fail(),
        value.forEach((element) => collectBooleans(element, true)));
    else if (typeof value == "boolean" || typeof value == "number")
      booleans.push(!!value);
    else if (value == null) isFromReference || booleans.push(false);
    else if (typeof value == "string") {
      if (isFromReference) return;
      (/^(true|false)$/i.test(value) || fail(),
        booleans.push(value.toLowerCase() === "true"));
    } else fail();
  };
  return (
    args.forEach((argument) =>
      collectBooleans(
        context.retrieve(argument),
        functionName !== "NOT" &&
          !!argument &&
          typeof argument == "object" &&
          "ref" in argument,
      ),
    ),
    booleans.length || fail(),
    functionName === "NOT"
      ? !booleans[0]
      : functionName === "AND"
        ? booleans.every(Boolean)
        : functionName === "OR"
          ? booleans.some(Boolean)
          : booleans.filter(Boolean).length % 2 === 1
  );
}
const modernFunctionNames = [
    "XLOOKUP",
    "XMATCH",
    "IFS",
    "SWITCH",
    "TEXTJOIN",
    "CONCAT",
    "MAXIFS",
    "MINIFS",
    "FILTER",
    "SORT",
    "SORTBY",
    "UNIQUE",
    "SEQUENCE",
    "TRANSPOSE",
    "TAKE",
    "DROP",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "HSTACK",
    "VSTACK",
    "TOCOL",
    "TOROW",
    "EXPAND",
    "TEXTBEFORE",
    "TEXTAFTER",
    "TEXTSPLIT",
    "STDEV.P",
    "VAR.P",
    "PERCENTILE.INC",
    "PERCENTILE.EXC",
    "QUARTILE.INC",
    "QUARTILE.EXC",
    "RANK.EQ",
    "RANK.AVG",
    "MODE.SNGL",
    "MDETERM",
    "MINVERSE",
  ],
  matrixResultFunctionNames = [
    "CHOOSE",
    "FILTER",
    "SORT",
    "SORTBY",
    "UNIQUE",
    "SEQUENCE",
    "TRANSPOSE",
    "TAKE",
    "DROP",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "HSTACK",
    "VSTACK",
    "TOCOL",
    "TOROW",
    "EXPAND",
    "TEXTSPLIT",
    "MUNIT",
    "MMULT",
    "MINVERSE",
    "FREQUENCY",
    "GROWTH",
    "LINEST",
    "LOGEST",
    "TREND",
    "MODE.MULT",
    "XLOOKUP",
  ],
  notModernFunctionSentinel = Symbol("not a modern function"),
  toMatrixArgument = (value) =>
    Array.isArray(value)
      ? value.every(Array.isArray)
        ? value
        : [value]
      : [[value]],
  compareSortValues = (leftValue, rightValue) =>
    typeof leftValue == "number" && typeof rightValue == "number"
      ? leftValue - rightValue
      : coerceValueToText(leftValue)
          .toLowerCase()
          .localeCompare(coerceValueToText(rightValue).toLowerCase()),
  valuesEqualIgnoringCase = (leftValue, rightValue) =>
    typeof leftValue == "string" && typeof rightValue == "string"
      ? leftValue.toLowerCase() === rightValue.toLowerCase()
      : leftValue === rightValue,
  compareSortKeyWithDirection = (leftKey, rightKey, sortDirection) => {
    if (leftKey == null || rightKey == null)
      return leftKey == null ? (rightKey == null ? 0 : 1) : -1;
    const typeRank = (value) =>
      typeof value == "number" ? 0 : typeof value == "string" ? 1 : 2;
    return (
      (typeRank(leftKey) - typeRank(rightKey) ||
        compareSortValues(leftKey, rightKey)) * sortDirection
    );
  };

export {
  notRegressionSentinel,
  notArithmeticSentinel,
  notSpecialistSentinel,
  FormulaParser,
  frozenConversionUnits,
  notInferenceSentinel,
  evaluateArithmeticFunction,
  buildNumberFormatCode,
  notStatisticalSentinel,
  statisticalFunctionNames,
  sumDecimalPartsToNumber,
  getImportedNumberFormatWarning,
  exactAverageOfNumbers,
  exactProductOfNumbers,
  XL,
  specialistFunctionNames,
  toMatrixArgument,
  evaluateSignificanceTestFunction,
  significanceTestFunctionNames,
  notOwnedLookupSentinel,
  evaluateSpecialistFunction,
  valuesEqualIgnoringCase,
  compareSortKeyWithDirection,
  evaluateLookupFunction,
  compareSortValues,
  evaluateRegressionFunction,
  validateNumberFormatSettings,
  ownedLookupFunctionNames,
  evaluateLogicalFunction,
  notModernFunctionSentinel,
  modernFunctionNames,
  evaluateStatisticalFunction,
  notOwnedLogicalSentinel,
  formatCellValueWithNumberFormat,
  matrixResultFunctionNames,
};

/* Also used by the workers. */
export {
  addConversionUnitSymbols,
  arithmeticFunctionNames,
  binaryPrefixExponents,
  conversionUnits,
  dotProductWithCompensation,
  findConversionUnit,
  formatNumericSection,
  formatTemporalSerialValue,
  getLocaleDecimalSeparator,
  isReferenceArgument,
  regressionSizeLimits,
  renderTextSectionPattern,
  roundDecimalToDigits,
  roundQuotientWithMode,
  roundToNearestMultiple,
  siPrefixPowersOfTen,
  sumDecimalParts,
  sumWithKahanCompensation,
  toSignedDecimalParts,
  toUnsignedDecimalParts,
};
