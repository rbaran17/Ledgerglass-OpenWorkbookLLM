import { describe, expect, test } from "vitest";
import {
  createBlankWorkbook,
  getWorkbookCalculation,
} from "../../src/app/03-workbook-model.js";

// Evaluates a formula in cell Z1. Extra cells (for range arguments) are
// written as formulas such as "=-1000" so the calculator treats them as numbers.
function evaluate(formula, cells = {}) {
  const workbook = createBlankWorkbook("Unit");
  const sheet = workbook.sheets[0];
  for (const [address, raw] of Object.entries(cells))
    sheet.cells[address] = { raw };
  sheet.cells.Z1 = { raw: formula };
  const { calculator } = getWorkbookCalculation(workbook);
  return calculator.getTypedValue(sheet.id, "Z1");
}

describe("PMT", () => {
  test("monthly payment on a 5% loan of 10000 over 60 months", () => {
    const result = evaluate("=PMT(0.05/12,60,10000)");
    expect(result.type).toBe("number");
    expect(result.value).toBeCloseTo(-188.71, 2);
  });

  test("a positive loan produces a negative payment, a negative loan a positive one", () => {
    expect(evaluate("=PMT(0.1,2,-1000)").value).toBeCloseTo(576.19, 2);
    expect(evaluate("=PMT(0.1,2,1000)").value).toBeCloseTo(-576.19, 2);
  });

  test("annuity due (type 1) payment is smaller than ordinary annuity payment in magnitude", () => {
    const due = evaluate("=PMT(0.1,2,1000,0,1)").value;
    expect(due).toBeCloseTo(-523.81, 2);
    expect(Math.abs(due)).toBeLessThan(
      Math.abs(evaluate("=PMT(0.1,2,1000)").value),
    );
  });

  test("zero interest rate divides the principal evenly", () => {
    expect(evaluate("=PMT(0,10,1000)").value).toBeCloseTo(-100, 10);
  });

  test("zero periods returns #NUM!", () => {
    expect(evaluate("=PMT(0.1,0,1000)")).toMatchObject({
      type: "error",
      value: "#NUM!",
    });
  });
});

describe("PV and FV", () => {
  test("future value of monthly deposits of 100 at 6% annual over 120 months", () => {
    const result = evaluate("=FV(0.06/12,120,-100)");
    expect(result.value).toBeCloseTo(16387.93, 2);
  });

  test("deposits are shown as positive future value when paid as negatives", () => {
    expect(evaluate("=FV(0.1,2,-100,0,1)").value).toBeCloseTo(231, 6);
  });

  test("zero interest rate makes FV and PV the plain sum of payments", () => {
    expect(evaluate("=FV(0,10,-100)").value).toBeCloseTo(1000, 10);
    expect(evaluate("=PV(0,10,-100)").value).toBeCloseTo(1000, 10);
  });

  test("present value of an annuity due of 500 per quarter for 20 quarters", () => {
    expect(evaluate("=PV(0.08/4,20,-500,0,1)").value).toBeCloseTo(8339.23, 2);
  });

  test("present value of a 60-month payment stream roughly returns the loan principal", () => {
    expect(evaluate("=PV(0.05/12,60,-188.71)").value).toBeCloseTo(10000, 0);
  });

  test("a text argument returns #VALUE!", () => {
    expect(evaluate('=FV("x",1,1)')).toMatchObject({
      type: "error",
      value: "#VALUE!",
    });
  });
});

describe("NPER and RATE", () => {
  test("periods needed to repay 1000 at 1% a period with payments of 100", () => {
    expect(evaluate("=NPER(0.01,-100,1000)").value).toBeCloseTo(10.5886, 3);
  });

  test("zero interest rate gives periods as principal divided by payment", () => {
    expect(evaluate("=NPER(0,-100,1000)").value).toBeCloseTo(10, 10);
  });

  test("rate implied by 48 payments of 200 on a principal of 8000", () => {
    expect(evaluate("=RATE(48,-200,8000)").value).toBeCloseTo(0.0077015, 6);
  });

  test("NPER with no payment and no principal returns #NUM!", () => {
    expect(evaluate("=NPER(0.1,0,0)")).toMatchObject({
      type: "error",
      value: "#NUM!",
    });
  });
});

describe("NPV and IRR", () => {
  test("NPV of three cash flows of 100 at 10%", () => {
    expect(evaluate("=NPV(0.1,100,100,100)").value).toBeCloseTo(248.69, 2);
  });

  test("NPV discounts the first cash flow by one period", () => {
    expect(evaluate("=NPV(0.1,-100,200)").value).toBeCloseTo(74.38, 2);
  });

  test("IRR of an outlay of 1000 returning 400 for three periods", () => {
    const cells = { A1: "=-1000", A2: "=400", A3: "=400", A4: "=400" };
    expect(evaluate("=IRR(A1:A4)", cells).value).toBeCloseTo(0.09701, 4);
  });

  test("IRR with no sign change in cash flows returns #NUM!", () => {
    const cells = { A1: "=100", A2: "=200" };
    expect(evaluate("=IRR(A1:A2)", cells)).toMatchObject({
      type: "error",
      value: "#NUM!",
    });
  });
});

describe("IPMT and PPMT", () => {
  test("interest in period 1 of a 3-period loan of 8000 at 10%", () => {
    expect(evaluate("=IPMT(0.1,1,3,8000)").value).toBeCloseTo(-800, 6);
  });

  test("principal in period 1 of the same loan", () => {
    expect(evaluate("=PPMT(0.1,1,3,8000)").value).toBeCloseTo(-2416.92, 2);
  });

  test("interest and principal for a period add up to the payment", () => {
    const interest = evaluate("=IPMT(0.1,2,3,8000)").value;
    const principal = evaluate("=PPMT(0.1,2,3,8000)").value;
    const payment = evaluate("=PMT(0.1,3,8000)").value;
    expect(interest + principal).toBeCloseTo(payment, 6);
  });

  test("a period outside 1 to nper returns #NUM!", () => {
    expect(evaluate("=IPMT(0.1,0,3,8000)")).toMatchObject({
      type: "error",
      value: "#NUM!",
    });
    expect(evaluate("=IPMT(0.1,5,3,8000)")).toMatchObject({
      type: "error",
      value: "#NUM!",
    });
  });
});

describe("SLN, SYD, DB and DDB", () => {
  test("straight-line depreciation spreads cost less salvage evenly", () => {
    expect(evaluate("=SLN(30000,7500,10)").value).toBeCloseTo(2250, 10);
  });

  test("zero life returns #NUM!", () => {
    expect(evaluate("=SLN(1000,100,0)")).toMatchObject({
      type: "error",
      value: "#NUM!",
    });
  });

  test("sum-of-years-digits depreciation is larger in early periods", () => {
    expect(evaluate("=SYD(30000,7500,10,1)").value).toBeCloseTo(4090.91, 2);
    expect(evaluate("=SYD(30000,7500,10,10)").value).toBeCloseTo(409.09, 2);
  });

  test("fixed-declining depreciation first period matches the documented example", () => {
    expect(evaluate("=DB(1000000,100000,6,1)").value).toBeCloseTo(319000, 0);
    expect(evaluate("=DB(1000000,100000,6,7,6)").value).toBeCloseTo(
      19635.14,
      2,
    );
  });

  test("double-declining balance uses a factor of 2 by default", () => {
    expect(evaluate("=DDB(2400,300,10,1)").value).toBeCloseTo(480, 10);
  });

  test("double-declining balance honours an explicit factor", () => {
    expect(evaluate("=DDB(2400,300,10,1,1.5)").value).toBeCloseTo(360, 10);
  });

  test("negative cost returns #NUM!", () => {
    expect(evaluate("=DDB(-1,300,10,1)")).toMatchObject({
      type: "error",
      value: "#NUM!",
    });
  });
});
