import {
  dateToSerial,
  evaluateDateFunction,
  daysInExcelMonth,
  financialSolverLimits,
  serialToDateParts,
  dateSystemMaxSerial,
  parseDateText,
} from "./01-schemas-date-functions-and-validation.js";

function solveRateCandidates(
  evaluateResidual,
  guess,
  minRate = financialSolverLimits.minRate,
) {
  const minLogRate = Math.log1p(minRate),
    maxLogRate = Math.log1p(financialSolverLimits.maxRate),
    guessLogRate = Math.log1p(guess);
  let evaluations = 0;
  const candidates = [],
    evaluateLogRate = (probeLogRate) => {
      if (
        probeLogRate < minLogRate ||
        probeLogRate > maxLogRate ||
        !Number.isFinite(probeLogRate) ||
        evaluations >= financialSolverLimits.evaluations
      )
        return;
      evaluations++;
      const residualPoint = evaluateResidual(Math.expm1(probeLogRate));
      if (!(
        !Number.isFinite(residualPoint.value) ||
        !Number.isFinite(residualPoint.scale) ||
        residualPoint.scale <= 0
      ))
        return {
          ...residualPoint,
          normalized: residualPoint.value / residualPoint.scale,
        };
    },
    recordCandidate = (candidateLogRate, candidatePoint) => {
      if (
        !candidatePoint ||
        Math.abs(candidatePoint.normalized) >
          financialSolverLimits.relativeResidual
      )
        return false;
      const rate = Math.expm1(candidateLogRate);
      return (
        candidates.some(
          (candidate) =>
            Math.abs(candidate.rate - rate) <=
            1e-7 * Math.max(1, Math.abs(rate)),
        ) ||
          candidates.push({
            rate: rate,
            residual: candidatePoint.value,
            relativeResidual: Math.abs(candidatePoint.normalized),
          }),
        true
      );
    };
  if (
    [0, Math.log1p(0.1), Math.log1p(0.5)]
      .map(evaluateLogRate)
      .every((evaluatedPoint) => evaluatedPoint?.value === 0)
  )
    return {
      status: "indeterminate",
      candidates: candidates,
      evaluations: evaluations,
      exhaustive: false,
    };
  let currentLogRate = guessLogRate;
  for (let iteration = 0; iteration < 60; iteration++) {
    const currentPoint = evaluateLogRate(currentLogRate);
    if (!currentPoint || recordCandidate(currentLogRate, currentPoint)) break;
    const derivativeStep = 1e-5,
      lowerPoint = evaluateLogRate(
        Math.max(minLogRate, currentLogRate - derivativeStep),
      ),
      upperPoint = evaluateLogRate(
        Math.min(maxLogRate, currentLogRate + derivativeStep),
      );
    if (!lowerPoint || !upperPoint) break;
    const derivative =
      (upperPoint.normalized - lowerPoint.normalized) /
      (Math.min(maxLogRate, currentLogRate + derivativeStep) -
        Math.max(minLogRate, currentLogRate - derivativeStep));
    if (!Number.isFinite(derivative) || Math.abs(derivative) < 1e-15) break;
    const nextLogRate =
      currentLogRate -
      Math.max(-1, Math.min(1, currentPoint.normalized / derivative));
    if (
      nextLogRate < minLogRate ||
      nextLogRate > maxLogRate ||
      nextLogRate === currentLogRate
    )
      break;
    currentLogRate = nextLogRate;
  }
  const gridLogRates = [
    ...new Set([
      ...Array.from(
        { length: 193 },
        (g, gridIndex) =>
          minLogRate + ((maxLogRate - minLogRate) * gridIndex) / 192,
      ),
      ...[
        -0.9,
        -0.5,
        -0.1,
        0,
        0.01,
        0.025,
        0.05,
        0.075,
        0.1,
        0.15,
        0.2,
        0.3,
        0.5,
        1,
        2,
        10,
        guess,
      ]
        .filter((candidateRate) => candidateRate >= minRate)
        .map(Math.log1p),
    ]),
  ].sort((leftLogRate, rightLogRate) => leftLogRate - rightLogRate);
  let previousSample;
  for (const scanLogRate of gridLogRates) {
    const scanPoint = evaluateLogRate(scanLogRate);
    if (!scanPoint) {
      previousSample = undefined;
      continue;
    }
    if (
      (recordCandidate(scanLogRate, scanPoint),
      previousSample && previousSample.y.normalized * scanPoint.normalized < 0)
    ) {
      let bracketLowLogRate = previousSample.x,
        bracketHighLogRate = scanLogRate,
        normalized = previousSample.y.normalized;
      for (let bisectionStep = 0; bisectionStep < 100; bisectionStep++) {
        const midLogRate = (bracketLowLogRate + bracketHighLogRate) / 2,
          midPoint = evaluateLogRate(midLogRate);
        if (
          !midPoint ||
          recordCandidate(midLogRate, midPoint) ||
          midLogRate === bracketLowLogRate ||
          midLogRate === bracketHighLogRate
        )
          break;
        normalized * midPoint.normalized < 0
          ? (bracketHighLogRate = midLogRate)
          : ((bracketLowLogRate = midLogRate),
            (normalized = midPoint.normalized));
      }
    }
    previousSample = { x: scanLogRate, y: scanPoint };
  }
  candidates.sort(
    (leftCandidate, rightCandidate) => leftCandidate.rate - rightCandidate.rate,
  );
  const selected = [...candidates].sort(
      (firstCandidate, secondCandidate) =>
        Math.abs(Math.log1p(firstCandidate.rate) - guessLogRate) -
        Math.abs(Math.log1p(secondCandidate.rate) - guessLogRate),
    )[0],
    isEvaluationLimitReached = evaluations >= financialSolverLimits.evaluations;
  return {
    ...(selected && !isEvaluationLimitReached ? { selected: selected } : {}),
    candidates: candidates,
    evaluations: evaluations,
    status: isEvaluationLimitReached
      ? "limit"
      : selected
        ? "converged"
        : "no-convergence",
    exhaustive: false,
  };
}
const supportedFinancialFunctionNames = [
    "PV",
    "FV",
    "PMT",
    "IPMT",
    "PPMT",
    "RATE",
    "NPER",
    "NPV",
    "XNPV",
    "IRR",
    "XIRR",
    "MIRR",
    "SLN",
    "DB",
    "DDB",
    "SYD",
    "PRICE",
    "YIELD",
    "ACCRINT",
    "ACCRINTM",
  ],
  notFinancialSentinel = Symbol("not-financial"),
  financialArgumentRanges = {
    PV: [3, 5],
    FV: [3, 5],
    PMT: [3, 5],
    IPMT: [4, 6],
    PPMT: [4, 6],
    RATE: [3, 6],
    NPER: [3, 5],
    NPV: [2, 255],
    XNPV: [3, 3],
    IRR: [1, 2],
    XIRR: [2, 3],
    MIRR: [3, 3],
    SLN: [3, 3],
    DB: [4, 5],
    DDB: [4, 5],
    SYD: [4, 4],
    PRICE: [6, 7],
    YIELD: [6, 7],
    ACCRINT: [6, 8],
    ACCRINTM: [4, 5],
  },
  flattenNestedValues = (nestedValue) =>
    Array.isArray(nestedValue)
      ? nestedValue.flatMap(flattenNestedValues)
      : [nestedValue],
  compensatedSum = (values) => {
    let sum = 0,
      compensation = 0;
    for (const value of values) {
      const correctedValue = value - compensation,
        nextSum = sum + correctedValue;
      ((compensation = nextSum - sum - correctedValue), (sum = nextSum));
    }
    return sum;
  },
  countSignChanges = (values) => {
    const signs = values.filter((value) => value !== 0).map(Math.sign);
    return signs.slice(1).filter((sign, signIndex) => sign !== signs[signIndex])
      .length;
  };
function evaluateFinancialFunction(functionName, argumentNodes, context) {
  if (!supportedFinancialFunctionNames.includes(functionName))
    return notFinancialSentinel;
  let convention =
      "Cash paid out is negative; cash received is positive. Rates and periods must use the same interval. No automatic cent rounding.",
    hasReported = false;
  const dateSystem = context.clock.dateSystem,
    toDateSerial = (year, month, day) =>
      dateToSerial(year, month, day, dateSystem),
    toDateParts = (serial) => serialToDateParts(serial, dateSystem),
    reportCalculation = (details) => {
      ((hasReported = true),
        context.report?.({
          functionName: functionName,
          convention: convention,
          ...details,
        }));
    },
    throwFinancialError = (errorCode, errorMessage) => {
      throw (
        reportCalculation({
          status: "error",
          error: errorCode,
          message: errorMessage,
        }),
        context.issue(errorCode)
      );
    },
    ensureFinite = (numericValue) =>
      Number.isFinite(numericValue)
        ? numericValue
        : throwFinancialError(
            "#NUM!",
            "The financial equation overflowed or has no finite result. Review rates, signs and periods.",
          ),
    toScalar = (argument) => {
      if (context.isIssue(argument)) throw argument;
      return Array.isArray(argument)
        ? argument.length === 1
          ? toScalar(argument[0])
          : throwFinancialError(
              "#VALUE!",
              "A scalar argument must contain exactly one value.",
            )
        : argument;
    },
    toFinancialNumber = (inputValue, fallbackValue) => (
      (inputValue = toScalar(inputValue)),
      inputValue == null
        ? (fallbackValue ?? 0)
        : typeof inputValue == "boolean"
          ? Number(inputValue)
          : typeof inputValue != "number" &&
              (typeof inputValue != "string" ||
                !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?\s*$/i.test(
                  inputValue,
                ))
            ? throwFinancialError(
                "#VALUE!",
                "Expected a finite number; nonnumeric text cannot be used as a financial input.",
              )
            : ensureFinite(Number(inputValue))
    ),
    toCashFlowDate = (dateInput) => {
      if (
        ((dateInput = toScalar(dateInput)),
        dateInput == null || typeof dateInput == "boolean")
      )
        return throwFinancialError(
          "#VALUE!",
          "A cash-flow or security date is missing or is not a date.",
        );
      const serialValue =
        typeof dateInput == "string"
          ? parseDateText(dateInput, context.clock)
          : toFinancialNumber(dateInput);
      return serialValue === undefined ||
        serialValue < (dateSystem === "1904" ? 0 : 1) ||
        serialValue >= dateSystemMaxSerial(dateSystem) + 1
        ? throwFinancialError(
            "#VALUE!",
            "Invalid cash-flow or security date. Use a valid Gregorian/workbook-system serial or explicit full-year date text.",
          )
        : Math.floor(serialValue);
    },
    validatePeriods = (periodCount) => (
      periodCount <= 0 &&
        throwFinancialError("#NUM!", "Periods must be greater than zero."),
      periodCount > financialSolverLimits.periods &&
        throwFinancialError(
          "#LIMIT!",
          "Financial periods exceed the 10000-period budget.",
        ),
      periodCount
    ),
    validateRate = (rateValue) =>
      rateValue > -1
        ? rateValue
        : throwFinancialError("#NUM!", "The rate must be greater than -100%."),
    validateTiming = (timingValue) =>
      [0, 1].includes(timingValue)
        ? timingValue
        : throwFinancialError(
            "#NUM!",
            "Payment timing must be 0 (end) or 1 (beginning).",
          ),
    growthFactor = (compoundRate, elapsedPeriods) =>
      Math.exp(elapsedPeriods * Math.log1p(compoundRate)),
    annuityFactor = (annuityRate, annuityPeriods) =>
      annuityRate === 0
        ? annuityPeriods
        : Math.expm1(annuityPeriods * Math.log1p(annuityRate)) / annuityRate,
    paymentAmount = (
      paymentRate,
      paymentPeriods,
      presentValue,
      futureValue,
      paymentTiming,
    ) =>
      -(
        presentValue * growthFactor(paymentRate, paymentPeriods) +
        futureValue
      ) /
      ((1 + paymentRate * paymentTiming) *
        annuityFactor(paymentRate, paymentPeriods)),
    collectCashFlows = (cashFlowInput, shouldRequireDates = false) => {
      const flatValues = flattenNestedValues(cashFlowInput);
      flatValues.length > financialSolverLimits.cashFlows &&
        throwFinancialError(
          "#LIMIT!",
          "Cash-flow input exceeds 10000 positions.",
        );
      const cashFlows = [];
      for (const flatValue of flatValues) {
        if (context.isIssue(flatValue)) throw flatValue;
        typeof flatValue == "number"
          ? cashFlows.push(ensureFinite(flatValue))
          : shouldRequireDates &&
            throwFinancialError(
              "#VALUE!",
              "Dated cash flows require one numeric amount for every date; blanks, booleans and text cannot be silently removed.",
            );
      }
      return (
        cashFlows.length ||
          throwFinancialError("#NUM!", "No numeric cash flows were supplied."),
        cashFlows
      );
    },
    solveRate = (residualFunction, guess, signChanges, lowerRateBound) => {
      (guess < (lowerRateBound ?? financialSolverLimits.minRate) ||
        guess > financialSolverLimits.maxRate) &&
        throwFinancialError(
          "#NUM!",
          "The rate guess is outside the supported search interval.",
        );
      const solver = solveRateCandidates(
          residualFunction,
          guess,
          lowerRateBound,
        ),
        failureCode = solver.status === "limit" ? "#LIMIT!" : "#NUM!";
      if (
        (reportCalculation({
          status: solver.selected ? "converged" : "error",
          ...(solver.selected ? {} : { error: failureCode }),
          guess: guess,
          signChanges: signChanges,
          solver: solver,
          message: solver.selected
            ? `Residual-checked rate. ${solver.candidates.length} candidate root(s) found; selected nearest the guess in log-rate space. ${signChanges !== undefined && signChanges > 1 ? "Cash flows change sign more than once; multiple returns may exist. " : ""}The bounded search does not prove uniqueness.`
            : solver.status === "indeterminate"
              ? "The equation does not identify a unique rate; no rate has been invented."
              : "No residual-checked rate converged within the supported search interval and evaluation budget. Review signs, timing and guess.",
        }),
        !solver.selected)
      )
        throw context.issue(failureCode);
      return solver.selected.rate;
    };
  try {
    const [minArgs, maxArgs] = financialArgumentRanges[functionName];
    (argumentNodes.length < minArgs || argumentNodes.length > maxArgs) &&
      throwFinancialError(
        "#VALUE!",
        `${functionName} requires ${minArgs}–${maxArgs} arguments.`,
      );
    const argumentValues = argumentNodes.map(context.retrieve);
    let financialResult;
    if (
      ["PV", "FV", "PMT", "IPMT", "PPMT", "RATE", "NPER"].includes(functionName)
    )
      if (functionName === "RATE") {
        const ratePeriods = validatePeriods(
            toFinancialNumber(argumentValues[0]),
          ),
          ratePayment = toFinancialNumber(argumentValues[1]),
          pvAmount = toFinancialNumber(argumentValues[2]),
          fvAmount = toFinancialNumber(argumentValues[3]),
          rateTiming = validateTiming(toFinancialNumber(argumentValues[4]));
        ([pvAmount, ratePayment, fvAmount].every((amount) => amount === 0) &&
          throwFinancialError(
            "#NUM!",
            "All amounts are zero; no rate can be determined.",
          ),
          (financialResult = solveRate(
            (candidateRate) => {
              const residualTerms = [
                pvAmount * growthFactor(candidateRate, ratePeriods),
                ratePayment *
                  (1 + candidateRate * rateTiming) *
                  annuityFactor(candidateRate, ratePeriods),
                fvAmount,
              ];
              return {
                value: compensatedSum(residualTerms),
                scale: compensatedSum(residualTerms.map(Math.abs)),
              };
            },
            toFinancialNumber(argumentValues[5], 0.1),
          )));
      } else if (functionName === "NPER") {
        const periodicRate = validateRate(toFinancialNumber(argumentValues[0])),
          nperPayment = toFinancialNumber(argumentValues[1]),
          nperPresentValue = toFinancialNumber(argumentValues[2]),
          nperFutureValue = toFinancialNumber(argumentValues[3]),
          nperTiming = validateTiming(toFinancialNumber(argumentValues[4]));
        if (periodicRate === 0)
          (nperPayment === 0 &&
            throwFinancialError(
              "#DIV/0!",
              "Zero rate and zero payment cannot determine a period count.",
            ),
            (financialResult =
              -(nperPresentValue + nperFutureValue) / nperPayment));
        else {
          const paymentTerm = nperPayment * (1 + periodicRate * nperTiming),
            logRatio =
              (paymentTerm - nperFutureValue * periodicRate) /
              (nperPresentValue * periodicRate + paymentTerm);
          (logRatio > 0 ||
            throwFinancialError(
              "#NUM!",
              "Cash-flow signs or payment size do not produce a finite real period count.",
            ),
            (financialResult = Math.log(logRatio) / Math.log1p(periodicRate)));
        }
        validatePeriods(financialResult);
      } else if (functionName === "IPMT" || functionName === "PPMT") {
        const interestRate = validateRate(toFinancialNumber(argumentValues[0])),
          period = toFinancialNumber(argumentValues[1]),
          totalPeriods = validatePeriods(toFinancialNumber(argumentValues[2])),
          loanPresentValue = toFinancialNumber(argumentValues[3]),
          loanFutureValue = toFinancialNumber(argumentValues[4]),
          ipmtTiming = validateTiming(toFinancialNumber(argumentValues[5]));
        (period < 1 || period > totalPeriods || !Number.isInteger(period)) &&
          throwFinancialError(
            "#NUM!",
            "Interest/principal period must be an integer from 1 through nper.",
          );
        const periodicPayment = paymentAmount(
            interestRate,
            totalPeriods,
            loanPresentValue,
            loanFutureValue,
            ipmtTiming,
          ),
          interestPart =
            ipmtTiming && period === 1
              ? 0
              : (-(
                  loanPresentValue * growthFactor(interestRate, period - 1) +
                  periodicPayment *
                    (1 + interestRate * ipmtTiming) *
                    annuityFactor(interestRate, period - 1)
                ) *
                  interestRate) /
                (1 + interestRate * ipmtTiming);
        financialResult =
          functionName === "IPMT"
            ? interestPart
            : periodicPayment - interestPart;
      } else {
        const pmtRate = validateRate(toFinancialNumber(argumentValues[0])),
          pmtPeriods = validatePeriods(toFinancialNumber(argumentValues[1])),
          pmtPresentValue = toFinancialNumber(argumentValues[2]),
          pmtFutureValue = toFinancialNumber(argumentValues[3]),
          pmtTiming = validateTiming(toFinancialNumber(argumentValues[4]));
        financialResult =
          functionName === "PMT"
            ? paymentAmount(
                pmtRate,
                pmtPeriods,
                pmtPresentValue,
                pmtFutureValue,
                pmtTiming,
              )
            : functionName === "FV"
              ? -(
                  pmtFutureValue * growthFactor(pmtRate, pmtPeriods) +
                  pmtPresentValue *
                    (1 + pmtRate * pmtTiming) *
                    annuityFactor(pmtRate, pmtPeriods)
                )
              : -(
                  pmtFutureValue +
                  pmtPresentValue *
                    (1 + pmtRate * pmtTiming) *
                    annuityFactor(pmtRate, pmtPeriods)
                ) / growthFactor(pmtRate, pmtPeriods);
      }
    else if (["NPV", "XNPV", "IRR", "XIRR", "MIRR"].includes(functionName)) {
      const isDated = functionName.startsWith("X"),
        isIrrFunction = functionName === "IRR" || functionName === "XIRR",
        numericCashFlows = collectCashFlows(
          functionName === "NPV"
            ? argumentValues.slice(1)
            : argumentValues[isIrrFunction || functionName === "MIRR" ? 0 : 1],
          isDated,
        );
      let timePositions = numericCashFlows.map(
        (_, periodIndex) => periodIndex + (functionName === "NPV" ? 1 : 0),
      );
      if (isDated) {
        const dateValues = flattenNestedValues(
          argumentValues[isIrrFunction ? 1 : 2],
        );
        dateValues.length !== numericCashFlows.length &&
          throwFinancialError(
            "#NUM!",
            "Every cash-flow amount must have exactly one date.",
          );
        const dateSerials = dateValues.map(toCashFlowDate);
        (dateSerials.some((serialEntry) => serialEntry < dateSerials[0]) &&
          throwFinancialError(
            "#NUM!",
            "A cash-flow date precedes the first cash-flow date.",
          ),
          (timePositions = dateSerials.map(
            (dateEntry) => (dateEntry - dateSerials[0]) / 365,
          )),
          (convention =
            "Dated cash flows use truncated Gregorian/workbook-system dates and actual elapsed days / 365, including leap years. The first date is time zero; later dates may be unsorted or repeated. One numeric amount per date."));
      } else
        convention =
          "Periodic cash-flow positions are equally spaced. NPV starts at period 1; IRR/MIRR start at period 0. Numeric zeros retain a period; blanks, text and booleans in cash-flow lists are ignored. No automatic cent rounding.";
      const npvResidual = (trialRate) => {
        const discountedTerms = numericCashFlows.map(
          (flowValue, flowIndex) =>
            flowValue /
            Math.exp(timePositions[flowIndex] * Math.log1p(trialRate)),
        );
        return {
          value: compensatedSum(discountedTerms),
          scale: compensatedSum(discountedTerms.map(Math.abs)),
        };
      };
      if (isIrrFunction || functionName === "MIRR")
        if (
          ((!numericCashFlows.some((cashFlowAmount) => cashFlowAmount < 0) ||
            !numericCashFlows.some((cashFlowToCheck) => cashFlowToCheck > 0)) &&
            throwFinancialError(
              "#NUM!",
              "At least one positive and one negative cash flow are required.",
            ),
          functionName === "MIRR")
        ) {
          const financeRate = validateRate(
              toFinancialNumber(argumentValues[1]),
            ),
            reinvestRate = validateRate(toFinancialNumber(argumentValues[2])),
            lastPeriod = numericCashFlows.length - 1,
            negativePresentValue = compensatedSum(
              numericCashFlows.map((cashFlow, flowPeriod) =>
                cashFlow < 0
                  ? cashFlow / Math.exp(flowPeriod * Math.log1p(financeRate))
                  : 0,
              ),
            ),
            positiveFutureValue = compensatedSum(
              numericCashFlows.map((laterFlow, laterPeriod) =>
                laterFlow > 0
                  ? laterFlow *
                    Math.exp(
                      (lastPeriod - laterPeriod) * Math.log1p(reinvestRate),
                    )
                  : 0,
              ),
            );
          financialResult = Math.expm1(
            Math.log(-positiveFutureValue / negativePresentValue) / lastPeriod,
          );
        } else
          financialResult = solveRate(
            npvResidual,
            toFinancialNumber(argumentValues[isDated ? 2 : 1], 0.1),
            countSignChanges(numericCashFlows),
          );
      else
        financialResult = npvResidual(
          validateRate(toFinancialNumber(argumentValues[0])),
        ).value;
    } else if (["SLN", "SYD", "DB", "DDB"].includes(functionName)) {
      const assetCost = toFinancialNumber(argumentValues[0]),
        salvageValue = toFinancialNumber(argumentValues[1]),
        usefulLife = validatePeriods(toFinancialNumber(argumentValues[2]));
      if (
        ((convention =
          "Depreciation uses matching life/period units and no automatic cent rounding. DB rounds its fixed rate to three decimals and prorates first/last years; DDB does not switch to straight line."),
        (assetCost < 0 || salvageValue < 0) &&
          throwFinancialError("#NUM!", "Cost and salvage must be nonnegative."),
        functionName === "SLN")
      )
        financialResult = (assetCost - salvageValue) / usefulLife;
      else {
        const depreciationPeriod = toFinancialNumber(argumentValues[3]);
        if (
          ((depreciationPeriod < 1 ||
            depreciationPeriod > usefulLife + (functionName === "DB" ? 1 : 0) ||
            !Number.isInteger(depreciationPeriod)) &&
            throwFinancialError(
              "#NUM!",
              "Depreciation period must be a supported whole period within asset life (plus DB partial final year).",
            ),
          functionName === "SYD")
        )
          financialResult =
            ((assetCost - salvageValue) *
              (usefulLife - depreciationPeriod + 1) *
              2) /
            (usefulLife * (usefulLife + 1));
        else if (functionName === "DB") {
          const firstYearMonths = toFinancialNumber(argumentValues[4], 12);
          if (
            ((!Number.isInteger(usefulLife) ||
              firstYearMonths < 1 ||
              firstYearMonths > 12 ||
              !Number.isInteger(firstYearMonths) ||
              depreciationPeriod >
                usefulLife + (firstYearMonths < 12 ? 1 : 0)) &&
              throwFinancialError(
                "#NUM!",
                "DB requires whole life/period and 1–12 first-year months; a partial first year permits one partial final year.",
              ),
            assetCost === 0 || salvageValue >= assetCost)
          )
            financialResult = 0;
          else {
            const dbRate =
              Math.round(
                (1 - (salvageValue / assetCost) ** (1 / usefulLife)) * 1e3,
              ) / 1e3;
            let remainingCost = assetCost;
            financialResult = 0;
            for (
              let depreciationYear = 1;
              depreciationYear <= depreciationPeriod;
              depreciationYear++
            )
              ((financialResult =
                remainingCost *
                dbRate *
                (depreciationYear === 1
                  ? firstYearMonths / 12
                  : depreciationYear > usefulLife
                    ? (12 - firstYearMonths) / 12
                    : 1)),
                (remainingCost -= financialResult));
          }
        } else {
          const decliningFactor = toFinancialNumber(argumentValues[4], 2);
          decliningFactor <= 0 &&
            throwFinancialError(
              "#NUM!",
              "DDB factor must be greater than zero.",
            );
          let bookValue = assetCost;
          financialResult = 0;
          for (let yearIndex = 1; yearIndex <= depreciationPeriod; yearIndex++)
            ((financialResult = Math.max(
              0,
              Math.min(
                (bookValue * decliningFactor) / usefulLife,
                bookValue - salvageValue,
              ),
            )),
              (bookValue -= financialResult));
        }
      }
    } else {
      const basisIndex =
          functionName === "ACCRINT" ? 6 : functionName === "ACCRINTM" ? 4 : 6,
        dayCountBasis = Math.trunc(
          toFinancialNumber(argumentValues[basisIndex]),
        );
      ([0, 1, 2, 3, 4].includes(dayCountBasis) ||
        throwFinancialError(
          "#NUM!",
          "Day-count basis must be 0 (US 30/360), 1 (actual/actual), 2 (actual/360), 3 (actual/365), or 4 (European 30/360).",
        ),
        (convention = `Security basis ${dayCountBasis}; annual coupon/yield rates, regular maturity-anchored coupons at frequency 1/2/4, end-of-month anchoring. PRICE/YIELD are per 100 face value. Gregorian/workbook-system dates; no business-day adjustment or automatic rounding.`));
      const daysBetween = (startDate, endDate) =>
          dayCountBasis === 0 || dayCountBasis === 4
            ? evaluateDateFunction(
                "DAYS360",
                [startDate, endDate, dayCountBasis === 4],
                {
                  ...context,
                  retrieve: (rawValue) => rawValue,
                },
              )
            : endDate - startDate,
        couponPeriodLength = (periodStart, periodEnd, periodsPerYear) =>
          dayCountBasis === 1
            ? periodEnd - periodStart
            : (dayCountBasis === 3 ? 365 : 360) / periodsPerYear,
        readFrequency = (frequencyInput) => {
          const frequency = Math.trunc(toFinancialNumber(frequencyInput));
          return [1, 2, 4].includes(frequency)
            ? frequency
            : throwFinancialError(
                "#NUM!",
                "Coupon frequency must be 1, 2 or 4.",
              );
        },
        addMonths = (anchorSerial, monthOffset) => {
          const dateParts = toDateParts(anchorSerial),
            totalMonths =
              dateParts.year * 12 + dateParts.month - 1 + monthOffset,
            targetYear = Math.floor(totalMonths / 12),
            targetMonth = (((totalMonths % 12) + 12) % 12) + 1;
          return toDateSerial(
            targetYear,
            targetMonth,
            dateParts.day === daysInExcelMonth(dateParts.year, dateParts.month)
              ? daysInExcelMonth(targetYear, targetMonth)
              : Math.min(
                  dateParts.day,
                  daysInExcelMonth(targetYear, targetMonth),
                ),
          );
        };
      if (functionName === "PRICE" || functionName === "YIELD") {
        const priceSettlementSerial = toCashFlowDate(argumentValues[0]),
          priceMaturitySerial = toCashFlowDate(argumentValues[1]),
          couponRate = toFinancialNumber(argumentValues[2]),
          yieldRate = toFinancialNumber(argumentValues[3]),
          redemptionValue = toFinancialNumber(argumentValues[4]),
          couponFrequency = readFrequency(argumentValues[5]);
        (priceSettlementSerial >= priceMaturitySerial ||
          couponRate < 0 ||
          redemptionValue <= 0 ||
          yieldRate < 0 ||
          (functionName === "YIELD" && yieldRate === 0)) &&
          throwFinancialError(
            "#NUM!",
            "Require settlement before maturity, nonnegative coupon/yield, and positive redemption/price.",
          );
        let couponCount = 1,
          previousCouponSerial = addMonths(
            priceMaturitySerial,
            -12 / couponFrequency,
          );
        for (; previousCouponSerial > priceSettlementSerial;)
          (couponCount >= financialSolverLimits.couponPeriods &&
            throwFinancialError(
              "#LIMIT!",
              "Security exceeds 1000 coupon periods.",
            ),
            couponCount++,
            (previousCouponSerial = addMonths(
              priceMaturitySerial,
              (-couponCount * 12) / couponFrequency,
            )));
        const nextCouponSerial = addMonths(
            priceMaturitySerial,
            (-(couponCount - 1) * 12) / couponFrequency,
          ),
          couponPeriodDays = couponPeriodLength(
            previousCouponSerial,
            nextCouponSerial,
            couponFrequency,
          ),
          accruedDays = daysBetween(
            previousCouponSerial,
            priceSettlementSerial,
          ),
          daysRemaining =
            dayCountBasis === 0 || dayCountBasis === 4
              ? couponPeriodDays - accruedDays
              : daysBetween(priceSettlementSerial, nextCouponSerial),
          fractionRemaining = daysRemaining / couponPeriodDays,
          couponPayment = (100 * couponRate) / couponFrequency,
          priceForYield = (candidateYield) =>
            couponCount === 1
              ? (redemptionValue + couponPayment) /
                  (1 + (candidateYield / couponFrequency) * fractionRemaining) -
                (couponPayment * accruedDays) / couponPeriodDays
              : compensatedSum(
                  Array.from(
                    { length: couponCount },
                    (ae, couponIndex) =>
                      couponPayment /
                      (1 + candidateYield / couponFrequency) **
                        (couponIndex + fractionRemaining),
                  ),
                ) +
                redemptionValue /
                  (1 + candidateYield / couponFrequency) **
                    (couponCount - 1 + fractionRemaining) -
                (couponPayment * accruedDays) / couponPeriodDays;
        financialResult =
          functionName === "PRICE"
            ? priceForYield(yieldRate)
            : solveRate(
                (trialYield) => {
                  const priceValue = priceForYield(trialYield);
                  return {
                    value: priceValue - yieldRate,
                    scale: Math.abs(priceValue) + yieldRate,
                  };
                },
                0.1,
                undefined,
                0,
              );
      } else if (functionName === "ACCRINTM") {
        const issueSerial = toCashFlowDate(argumentValues[0]),
          interestSettlementSerial = toCashFlowDate(argumentValues[1]),
          annualRate = toFinancialNumber(argumentValues[2]),
          parValue = toFinancialNumber(argumentValues[3], 1e3);
        (issueSerial >= interestSettlementSerial ||
          annualRate <= 0 ||
          parValue <= 0) &&
          throwFinancialError(
            "#NUM!",
            "Require issue before settlement, positive coupon rate and positive par.",
          );
        const yearFraction = evaluateDateFunction(
          "YEARFRAC",
          [issueSerial, interestSettlementSerial, dayCountBasis],
          {
            ...context,
            retrieve: (rawArgument) => rawArgument,
          },
        );
        ((financialResult = parValue * annualRate * yearFraction),
          (convention = `Accrued interest at maturity: par × annual rate × YEARFRAC basis ${dayCountBasis}, using the published local day-count convention. Gregorian/workbook-system; no automatic rounding.`));
      } else {
        const issueDateSerial = toCashFlowDate(argumentValues[0]),
          firstInterestSerial = toCashFlowDate(argumentValues[1]),
          accruedSettlementSerial = toCashFlowDate(argumentValues[2]),
          accrualRate = toFinancialNumber(argumentValues[3]),
          accrualPar = toFinancialNumber(argumentValues[4], 1e3),
          accrualFrequency = readFrequency(argumentValues[5]),
          calcMethod = validateTiming(toFinancialNumber(argumentValues[7], 1));
        (issueDateSerial >= accruedSettlementSerial ||
          firstInterestSerial <= issueDateSerial ||
          accrualRate <= 0 ||
          accrualPar <= 0) &&
          throwFinancialError(
            "#NUM!",
            "Require issue before first interest and settlement, positive rate and positive par.",
          );
        const accrualStartSerial =
          !calcMethod && accruedSettlementSerial > firstInterestSerial
            ? firstInterestSerial
            : issueDateSerial;
        let couponStepIndex = Math.floor(
          (toDateParts(accrualStartSerial).year * 12 +
            toDateParts(accrualStartSerial).month -
            toDateParts(firstInterestSerial).year * 12 -
            toDateParts(firstInterestSerial).month) /
            (12 / accrualFrequency),
        );
        for (
          ;
          addMonths(
            firstInterestSerial,
            (couponStepIndex * 12) / accrualFrequency,
          ) > accrualStartSerial;
        )
          couponStepIndex--;
        let periodBeginSerial = addMonths(
            firstInterestSerial,
            (couponStepIndex * 12) / accrualFrequency,
          ),
          periodEndSerial = addMonths(
            firstInterestSerial,
            ((couponStepIndex + 1) * 12) / accrualFrequency,
          ),
          accruedFractionSum = 0,
          quasiCouponCount = 0;
        for (; periodBeginSerial < accruedSettlementSerial;) {
          ++quasiCouponCount > financialSolverLimits.couponPeriods &&
            throwFinancialError(
              "#LIMIT!",
              "Accrual exceeds 1000 quasi-coupon periods.",
            );
          const segmentStart = Math.max(accrualStartSerial, periodBeginSerial),
            segmentEnd = Math.min(accruedSettlementSerial, periodEndSerial);
          (segmentEnd > segmentStart &&
            (accruedFractionSum +=
              daysBetween(segmentStart, segmentEnd) /
              couponPeriodLength(
                periodBeginSerial,
                periodEndSerial,
                accrualFrequency,
              )),
            (periodBeginSerial = periodEndSerial),
            (periodEndSerial = addMonths(
              firstInterestSerial,
              ((++couponStepIndex + 1) * 12) / accrualFrequency,
            )));
        }
        ((financialResult =
          ((accrualPar * accrualRate) / accrualFrequency) * accruedFractionSum),
          (convention = `Accrued coupon interest: quasi-coupon periods anchored to first_interest, basis ${dayCountBasis}, frequency ${accrualFrequency}; calc_method ${calcMethod} starts at ${!calcMethod && accruedSettlementSerial > firstInterestSerial ? "first_interest" : "issue"}. No business-day adjustment or automatic rounding.`));
      }
    }
    return (
      (financialResult = ensureFinite(financialResult)),
      hasReported ||
        reportCalculation({
          status: "calculated",
          message:
            "Calculated under the displayed sign, timing and day-count convention.",
        }),
      financialResult
    );
  } catch (error) {
    throw (
      hasReported ||
        reportCalculation({
          status: "error",
          error: context.isIssue(error)
            ? (error.error ?? "#VALUE!")
            : "#VALUE!",
          message:
            "An input or dependent calculation failed; inspect the source values and dates.",
        }),
      error
    );
  }
}
const defaultColumnWidth = 84,
  minColumnWidth = 40,
  maxColumnWidth = 600,
  defaultRowHeight = 21,
  minRowHeight = 21,
  maxRowHeight = 546,
  supportedFontFamilies = [
    "Calibri",
    "Aptos",
    "Arial",
    "Times New Roman",
    "Courier New",
  ],
  workbookLimits = {
    maxRows: 1048576,
    maxCols: 16384,
    maxCells: 1e5,
    maxFileBytes: 10 * 1024 * 1024,
    maxNativeBytes: 50 * 1024 * 1024,
    maxLLMCells: 500,
    maxPatches: 500,
  },
  fileImportLimits = {
    ...workbookLimits,
    maxCellTextLength: 32767,
    maxFormulaLength: 4096,
    maxFormulaRangeCells: 2e4,
    maxFormulaDepth: 64,
    maxSpillCells: 1e4,
    maxXlsxExpandedBytes: 50 * 1024 * 1024,
    maxXlsxEntries: 2e3,
    maxXlsxExpansionRatio: 200,
    xlsxRatioThresholdBytes: 1024 * 1024,
    maxXlsxCellRecords: workbookLimits.maxCells,
  },
  workbookLimitTable = [
    ["Rows per worksheet", "1,048,576"],
    ["Columns per worksheet", "16,384 (A–XFD)"],
    ["Worksheets", "No fixed count; available memory and file budgets apply"],
    [
      "Stored cells per workbook",
      "100,000, including cells with only formatting or notes",
    ],
    [
      "Excluded population size",
      "100,000 rows × 20 columns (2,000,000 cells) exceeds this release; split the source before importing",
    ],
    [
      "Interactive performance",
      "100 ms edit / 50 ms scroll targets; local benchmark available, reference browser and host validation required",
    ],
    ["Cell text", "32,767 text units"],
    ["Formula text", "4,096 text units, including the leading ="],
    ["Formula range", "20,000 cells per referenced rectangle"],
    ["Calculation work", "100,000 cell reads; 64 dependent formula levels"],
    ["Dynamic array output", "10,000 cells per array"],
    [
      "Cell changes / assistant selection",
      "500 per patch batch / 500 selected cells",
    ],
    ["XLSX, CSV and TSV imports", "10 MiB (10,485,760 bytes)"],
    [
      "XLSX exports",
      "10 MiB; generated files pass the same archive and expansion limits as import",
    ],
    ["CSV exports", "10 MiB; 100,000 cells including blank gaps"],
    ["Native JSON save and reopen", "50 MiB (52,428,800 UTF-8 bytes)"],
    [
      "XLSX processing",
      "50 MiB expanded; 2,000 archive entries; 100,000 cell records",
    ],
    ["XLSX compression", "200:1 maximum for any entry larger than 1 MiB"],
  ];
function broadcastOverArrays(argumentValues, scalarFunction, issue) {
  if (!argumentValues.some(Array.isArray))
    return scalarFunction(argumentValues);
  const grids = argumentValues.map((argument) =>
      Array.isArray(argument)
        ? argument.every(Array.isArray)
          ? argument
          : [argument]
        : [[argument]],
    ),
    rowCount = Math.max(...grids.map((gridForRows) => gridForRows.length)),
    columnCount = Math.max(
      ...grids.map((gridForColumns) => gridForColumns[0]?.length ?? 0),
    );
  return !rowCount ||
    !columnCount ||
    grids.some(
      (grid) =>
        !grid.length ||
        !grid[0].length ||
        grid.some((row) => row.length !== grid[0].length) ||
        (grid.length !== rowCount && grid.length !== 1) ||
        (grid[0].length !== columnCount && grid[0].length !== 1),
    )
    ? issue("#VALUE!")
    : rowCount * columnCount > fileImportLimits.maxSpillCells
      ? issue("#LIMIT!")
      : Array.from({ length: rowCount }, (s, rowIndex) =>
          Array.from({ length: columnCount }, (u, columnIndex) =>
            scalarFunction(
              grids.map(
                (argumentGrid) =>
                  argumentGrid[argumentGrid.length === 1 ? 0 : rowIndex][
                    argumentGrid[0].length === 1 ? 0 : columnIndex
                  ],
              ),
            ),
          ),
        );
}
const supportedTextFunctionNames = [
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "TRIM",
    "CLEAN",
    "UPPER",
    "LOWER",
    "PROPER",
    "FIND",
    "SEARCH",
    "SUBSTITUTE",
    "REPLACE",
    "VALUE",
    "NUMBERVALUE",
    "TEXT",
    "CONCAT",
    "CONCATENATE",
    "TEXTJOIN",
    "EXACT",
    "TEXTBEFORE",
    "TEXTAFTER",
  ],
  notTextFunctionSentinel = Symbol("not-text-function"),
  // Numbers become text with 15 significant digits, as they do in Excel, so
  // =0.1+0.2&"" gives "0.3" and not "0.30000000000000004".
  coerceValueToText = (value) =>
    value == null
      ? ""
      : typeof value == "boolean"
        ? value
          ? "TRUE"
          : "FALSE"
        : typeof value == "number" && Number.isFinite(value)
          ? String(Number(value.toPrecision(15)))
          : String(value);
function parseDateTimeText(text) {
  const dateMatch =
    /^(?:(\d{4})-(\d{1,2})-(\d{1,2})|(\d{1,2})\/(\d{1,2})\/(\d{4}))(?:[ T](.+))?$/.exec(
      text,
    );
  let dateSerial = 0,
    timeText = text;
  if (dateMatch) {
    const year = Number(dateMatch[1] ?? dateMatch[6]),
      month = Number(dateMatch[2] ?? dateMatch[4]),
      day = Number(dateMatch[3] ?? dateMatch[5]);
    if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1)
      return;
    const isFakeLeapDay = year === 1900 && month === 2 && day === 29,
      utcMillis = Date.UTC(year, month - 1, day),
      utcDate = new Date(utcMillis);
    if (
      !isFakeLeapDay &&
      (utcDate.getUTCMonth() !== month - 1 || utcDate.getUTCDate() !== day)
    )
      return;
    if (
      ((dateSerial = isFakeLeapDay
        ? 60
        : (utcMillis - Date.UTC(1899, 11, 31)) / 864e5 +
          (utcMillis >= Date.UTC(1900, 2, 1) ? 1 : 0)),
      !dateMatch[7])
    )
      return dateSerial;
    timeText = dateMatch[7];
  }
  const timeMatch =
    /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*(AM|PM)?$/i.exec(timeText);
  if (!timeMatch) return;
  let hours = Number(timeMatch[1]);
  const minutes = Number(timeMatch[2]),
    seconds = Number(timeMatch[3] ?? 0),
    meridiem = timeMatch[4]?.toUpperCase();
  if (!(
    minutes > 59 ||
    seconds >= 60 ||
    (meridiem ? hours < 1 || hours > 12 : hours > 23)
  ))
    return (
      meridiem && (hours = (hours % 12) + (meridiem === "PM" ? 12 : 0)),
      dateSerial + (hours * 3600 + minutes * 60 + seconds) / 86400
    );
}
function evaluateTextFunction(functionName, args, context) {
  if (!supportedTextFunctionNames.includes(functionName))
    return notTextFunctionSentinel;
  const fail = (errorCode = "#VALUE!") => {
      throw context.issue(errorCode);
    },
    argumentRanges = {
      LEFT: [1, 2],
      RIGHT: [1, 2],
      MID: [3, 3],
      LEN: [1, 1],
      TRIM: [1, 1],
      CLEAN: [1, 1],
      UPPER: [1, 1],
      LOWER: [1, 1],
      PROPER: [1, 1],
      FIND: [2, 3],
      SEARCH: [2, 3],
      SUBSTITUTE: [3, 4],
      REPLACE: [4, 4],
      VALUE: [1, 1],
      NUMBERVALUE: [1, 3],
      TEXT: [2, 2],
      CONCAT: [1, 253],
      CONCATENATE: [1, 255],
      TEXTJOIN: [3, 254],
      EXACT: [2, 2],
      TEXTBEFORE: [2, 6],
      TEXTAFTER: [2, 6],
    },
    [minArgs, maxArgs] = argumentRanges[functionName];
  (args.length < minArgs || args.length > maxArgs) && fail();
  const argValues = args.map(context.retrieve),
    issueValue = argValues.flatMap(flattenNestedValues).find(context.isIssue);
  if (issueValue) return issueValue;
  ((!(functionName === "CONCAT" || functionName === "TEXTJOIN") &&
    argValues.some(Array.isArray)) ||
    (functionName === "TEXTJOIN" &&
      argValues.slice(0, 2).some(Array.isArray))) &&
    fail();
  const limitText = (text) => (text.length <= 32767 ? text : fail("#LIMIT!")),
    coerceNumberArg = (rawArg, fallbackValue) => {
      if (rawArg === undefined && fallbackValue !== undefined)
        return fallbackValue;
      if (
        Array.isArray(rawArg) ||
        (typeof rawArg == "string" && !rawArg.trim())
      )
        return fail();
      const numericValue = Number(rawArg ?? 0);
      return Number.isFinite(numericValue) ? numericValue : fail();
    },
    coerceIntegerArg = (integerArg, integerFallback) =>
      Math.trunc(coerceNumberArg(integerArg, integerFallback)),
    coerceBooleanArg = (booleanArg) =>
      typeof booleanArg == "string"
        ? (/^(true|false)$/i.test(booleanArg) || fail(),
          booleanArg.toLowerCase() === "true")
        : !!booleanArg,
    sourceText = coerceValueToText(argValues[0]);
  if (functionName === "LEN") return [...sourceText].length;
  if (functionName === "TRIM")
    return limitText(sourceText.replace(/^ +| +$/g, "").replace(/ +/g, " "));
  if (functionName === "CLEAN")
    return limitText(sourceText.replace(/[\u0000-\u001f]/g, ""));
  if (functionName === "UPPER") return limitText(sourceText.toUpperCase());
  if (functionName === "LOWER") return limitText(sourceText.toLowerCase());
  if (functionName === "PROPER")
    return limitText(
      sourceText
        .toLowerCase()
        .replace(
          new RegExp("(^|[^\\p{L}\\p{M}])(\\p{L})", "gu"),
          (fullMatch, prefix, letter) => prefix + letter.toUpperCase(),
        ),
    );
  if (functionName === "EXACT")
    return sourceText === coerceValueToText(argValues[1]);
  if (functionName === "LEFT" || functionName === "RIGHT") {
    const charCount = coerceIntegerArg(argValues[1], 1),
      textChars = [...sourceText];
    return (
      charCount < 0 && fail(),
      (functionName === "LEFT"
        ? textChars.slice(0, charCount)
        : charCount
          ? textChars.slice(-charCount)
          : []
      ).join("")
    );
  }
  if (functionName === "MID" || functionName === "REPLACE") {
    const startPosition = coerceIntegerArg(argValues[1]),
      numChars = coerceIntegerArg(argValues[2]),
      sourceChars = [...sourceText];
    return (
      (startPosition < 1 || numChars < 0) && fail(),
      functionName === "MID"
        ? sourceChars
            .slice(startPosition - 1, startPosition - 1 + numChars)
            .join("")
        : limitText(
            sourceChars.slice(0, startPosition - 1).join("") +
              coerceValueToText(argValues[3]) +
              sourceChars.slice(startPosition - 1 + numChars).join(""),
          )
    );
  }
  if (functionName === "FIND" || functionName === "SEARCH") {
    const searchChars = [...coerceValueToText(argValues[1])],
      searchStart = coerceIntegerArg(argValues[2], 1);
    (searchStart < 1 || searchStart > Math.max(1, searchChars.length)) &&
      fail();
    const withinChars = [...sourceText],
      patternTokens = [];
    for (let charIndex = 0; charIndex < withinChars.length; charIndex++) {
      const char = withinChars[charIndex];
      functionName === "SEARCH" &&
      char === "~" &&
      ["*", "?", "~"].includes(withinChars[charIndex + 1])
        ? patternTokens.push({
            char: withinChars[++charIndex],
            kind: "literal",
          })
        : patternTokens.push({
            char: char,
            kind:
              functionName === "SEARCH" && char === "*"
                ? "many"
                : functionName === "SEARCH" && char === "?"
                  ? "one"
                  : "literal",
          });
    }
    patternTokens.length * searchChars.length > 2e6 && fail("#LIMIT!");
    let laterMatches = Array(searchChars.length + 1).fill(true);
    for (
      let tokenIndex = patternTokens.length - 1;
      tokenIndex >= 0;
      tokenIndex--
    ) {
      const token = patternTokens[tokenIndex],
        currentMatches = Array(searchChars.length + 1).fill(false);
      currentMatches[searchChars.length] =
        token.kind === "many" && laterMatches[searchChars.length];
      for (
        let textIndex = searchChars.length - 1;
        textIndex >= searchStart - 1;
        textIndex--
      )
        currentMatches[textIndex] =
          token.kind === "many"
            ? laterMatches[textIndex] || currentMatches[textIndex + 1]
            : laterMatches[textIndex + 1] &&
              (token.kind === "one" ||
                (functionName === "SEARCH"
                  ? searchChars[textIndex].toLowerCase() ===
                    token.char.toLowerCase()
                  : searchChars[textIndex] === token.char));
      laterMatches = currentMatches;
    }
    const matchIndex = laterMatches.findIndex(
      (isMatch, candidateIndex) => candidateIndex >= searchStart - 1 && isMatch,
    );
    return matchIndex < 0 ? fail() : matchIndex + 1;
  }
  if (functionName === "SUBSTITUTE") {
    const oldText = coerceValueToText(argValues[1]),
      newText = coerceValueToText(argValues[2]),
      targetOccurrence =
        argValues[3] === undefined ? undefined : coerceIntegerArg(argValues[3]);
    if (
      (targetOccurrence !== undefined && targetOccurrence < 1 && fail(),
      !oldText)
    )
      return sourceText;
    const textParts = sourceText.split(oldText),
      occurrenceCount = textParts.length - 1;
    return (
      sourceText.length +
        (targetOccurrence === undefined
          ? occurrenceCount
          : targetOccurrence <= occurrenceCount
            ? 1
            : 0) *
          (newText.length - oldText.length) >
        32767 && fail("#LIMIT!"),
      textParts.reduce(
        (accumulated, part, partIndex) =>
          partIndex === 0
            ? part
            : accumulated +
              (targetOccurrence === undefined || targetOccurrence === partIndex
                ? newText
                : oldText) +
              part,
        "",
      )
    );
  }
  if (
    functionName === "CONCAT" ||
    functionName === "CONCATENATE" ||
    functionName === "TEXTJOIN"
  ) {
    const delimiter =
        functionName === "TEXTJOIN" ? coerceValueToText(argValues[0]) : "",
      ignoreEmpty =
        functionName === "TEXTJOIN" && coerceBooleanArg(argValues[1]),
      textPieces = argValues
        .slice(functionName === "TEXTJOIN" ? 2 : 0)
        .flatMap(flattenNestedValues)
        .filter((value) => !ignoreEmpty || (value != null && value !== ""))
        .map(coerceValueToText);
    return (
      textPieces.reduce((total, piece) => total + piece.length, 0) +
        Math.max(0, textPieces.length - 1) * delimiter.length >
        32767 && fail("#LIMIT!"),
      textPieces.join(delimiter)
    );
  }
  if (functionName === "NUMBERVALUE") {
    const decimalSeparator = [
        ...coerceValueToText(argValues[1] === undefined ? "." : argValues[1]),
      ][0],
      groupSeparator = [
        ...coerceValueToText(argValues[2] === undefined ? "," : argValues[2]),
      ][0];
    (!decimalSeparator ||
      !groupSeparator ||
      decimalSeparator === groupSeparator ||
      /[0-9+\-eE%]/.test(decimalSeparator) ||
      /[0-9+\-eE%]/.test(groupSeparator)) &&
      fail();
    let numericText = sourceText.replace(/ /g, "");
    if (!numericText) return 0;
    const percentCount = /%+$/.exec(numericText)?.[0].length ?? 0;
    numericText = numericText.slice(0, numericText.length - percentCount);
    const decimalParts = numericText.split(decimalSeparator);
    ((decimalParts.length > 2 || decimalParts[1]?.includes(groupSeparator)) &&
      fail(),
      (numericText =
        decimalParts[0].split(groupSeparator).join("") +
        (decimalParts.length === 2 ? "." + decimalParts[1] : "")),
      /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(numericText) ||
        fail());
    let parsedNumber = Number(numericText);
    Number.isFinite(parsedNumber) || fail("#NUM!");
    for (let percentIndex = 0; percentIndex < percentCount; percentIndex++)
      parsedNumber /= 100;
    return parsedNumber;
  }
  if (functionName === "VALUE") {
    const temporalValue = context.temporal
      ? context.temporal(sourceText.trim())
      : parseDateTimeText(sourceText.trim());
    if (temporalValue !== undefined) return temporalValue;
    let numberText = sourceText.trim(),
      isNegative = false;
    numberText.startsWith("(") &&
      numberText.endsWith(")") &&
      ((isNegative = true), (numberText = numberText.slice(1, -1).trim()));
    const isPercent = numberText.endsWith("%");
    (isPercent && (numberText = numberText.slice(0, -1).trim()),
      (numberText = numberText.replace(/^\$\s*/, "")),
      isNegative && /^[+-]/.test(numberText) && fail(),
      /^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(
        numberText,
      ) || fail());
    const parsedValue =
      (Number(numberText.replace(/,/g, "")) * (isNegative ? -1 : 1)) /
      (isPercent ? 100 : 1);
    return Number.isFinite(parsedValue) ? parsedValue : fail("#NUM!");
  }
  if (functionName === "TEXT") {
    ((typeof argValues[1] != "string" || argValues[1].length > 256) && fail(),
      coerceNumberArg(argValues[0]));
    const formattedValue = context.format();
    return typeof formattedValue == "string"
      ? limitText(formattedValue)
      : formattedValue;
  }
  if (functionName === "TEXTBEFORE" || functionName === "TEXTAFTER") {
    const delimiterText = coerceValueToText(argValues[1]),
      instanceNum = coerceIntegerArg(argValues[2], 1),
      matchMode = coerceNumberArg(argValues[3], 0),
      matchEndMode = coerceNumberArg(argValues[4], 0);
    if (
      ((!instanceNum ||
        ![0, 1].includes(matchMode) ||
        ![0, 1].includes(matchEndMode)) &&
        fail(),
      !sourceText)
    )
      return "";
    if (
      (Math.abs(instanceNum) > [...sourceText].length && fail(), !delimiterText)
    )
      return (functionName === "TEXTBEFORE") == instanceNum > 0
        ? ""
        : sourceText;
    const escapedDelimiter = delimiterText.replace(
        /[.*+?^${}()|[\]\\]/g,
        "\\$&",
      ),
      delimiterRegex = new RegExp(
        instanceNum < 0 ? `(?=${escapedDelimiter})` : escapedDelimiter,
        matchMode ? "giu" : "gu",
      );
    let delimiterIndexes = [...sourceText.matchAll(delimiterRegex)].map(
      (match) => match.index,
    );
    if (instanceNum < 0) {
      const pickedIndexes = [];
      let searchEnd = sourceText.length;
      for (
        let reverseIndex = delimiterIndexes.length - 1;
        reverseIndex >= 0;
        reverseIndex--
      )
        delimiterIndexes[reverseIndex] + delimiterText.length <= searchEnd &&
          (pickedIndexes.push(delimiterIndexes[reverseIndex]),
          (searchEnd = delimiterIndexes[reverseIndex]));
      delimiterIndexes = pickedIndexes.reverse();
    }
    const delimiterIndex =
      delimiterIndexes[
        instanceNum > 0
          ? instanceNum - 1
          : delimiterIndexes.length + instanceNum
      ];
    return delimiterIndex === undefined
      ? matchEndMode && Math.abs(instanceNum) === delimiterIndexes.length + 1
        ? (functionName === "TEXTBEFORE") == instanceNum > 0
          ? sourceText
          : ""
        : argValues[5] !== undefined
          ? argValues[5]
          : fail("#N/A")
      : functionName === "TEXTBEFORE"
        ? sourceText.slice(0, delimiterIndex)
        : sourceText.slice(delimiterIndex + delimiterText.length);
  }
  return notTextFunctionSentinel;
}
const maxSelectionAreas = 256;

export {
  workbookLimitTable,
  maxRowHeight,
  fileImportLimits,
  workbookLimits,
  evaluateFinancialFunction,
  maxSelectionAreas,
  notFinancialSentinel,
  minRowHeight,
  maxColumnWidth,
  supportedFontFamilies,
  evaluateTextFunction,
  defaultRowHeight,
  supportedFinancialFunctionNames,
  notTextFunctionSentinel,
  defaultColumnWidth,
  broadcastOverArrays,
  minColumnWidth,
};
export { coerceValueToText, flattenNestedValues };

/* Also used by the workers. */
export {
  compensatedSum,
  countSignChanges,
  financialArgumentRanges,
  parseDateTimeText,
  solveRateCandidates,
  supportedTextFunctionNames,
};
