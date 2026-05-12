import { describe, expect, it } from "vitest";
import {
  collectMappedItems,
  employeeNumberFromHireDate,
  formatHireDate,
  getPreviousPayrollPeriod
} from "./payroll";

describe("payroll date helpers", () => {
  it("calculates the previous month payroll period from the payment date", () => {
    expect(getPreviousPayrollPeriod("2026-04-10")).toEqual({
      titleMonth: "2026년 03월",
      fileMonth: "2026년_03월",
      periodText: "2026.03.01 ~ 2026.03.31",
      paymentDateText: "2026/04/10"
    });
  });

  it("handles January payment dates by using the previous December", () => {
    expect(getPreviousPayrollPeriod("2026-01-10").periodText).toBe(
      "2025.12.01 ~ 2025.12.31"
    );
  });

  it("creates an employee number from the hire date", () => {
    expect(employeeNumberFromHireDate("2025-04-10")).toBe("250410");
    expect(formatHireDate("2025-04-10")).toBe("2025.04.10");
  });
});

describe("collectMappedItems", () => {
  it("filters empty and zero values, applies labels, and preserves Excel column order", () => {
    const items = collectMappedItems(
      new Map([
        ["주민세", { value: "₩3,970", columnIndex: 14 }],
        ["국민연금", { value: 117130, columnIndex: 9 }],
        ["건강보험", { value: 0, columnIndex: 10 }],
        ["고용보험", { value: null, columnIndex: 11 }]
      ]),
      [
        { source: "주민세", label: "지방소득세" },
        { source: "건강보험", label: "건강보험" },
        { source: "국민연금", label: "국민연금" },
        { source: "고용보험", label: "고용보험" }
      ]
    );

    expect(items).toEqual([
      { label: "국민연금", amount: 117130, columnIndex: 9 },
      { label: "지방소득세", amount: 3970, columnIndex: 14 }
    ]);
  });
});
