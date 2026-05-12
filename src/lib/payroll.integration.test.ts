import ExcelJS from "exceljs";
import { readFile } from "node:fs/promises";
import { describe, expect, it } from "vitest";
import hireDatesJson from "../data/hireDates.json";
import mappingsJson from "../data/payItemMappings.json";
import {
  HireDateMap,
  PayItemMappings,
  generatePayrollStatements
} from "./payroll";

describe("generatePayrollStatements", () => {
  it("generates one statement sheet from the sample payroll workbook", async () => {
    const source = await readFile("주식회사브로넥스-202604 (4).xlsx");
    const template = await readFile("public/payroll-template.xlsx");
    const sourceBuffer = new ArrayBuffer(source.byteLength);
    const templateBuffer = new ArrayBuffer(template.byteLength);
    new Uint8Array(sourceBuffer).set(source);
    new Uint8Array(templateBuffer).set(template);
    const result = await generatePayrollStatements({
      sourceBuffer,
      templateBuffer,
      companyName: "주식회사브로넥스",
      paymentDate: "2026-04-10",
      hireDates: hireDatesJson as HireDateMap,
      mappings: mappingsJson as PayItemMappings
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(result.buffer);
    const worksheet = workbook.getWorksheet("최태호");

    expect(result.fileName).toBe("2026년_03월_급여명세서_주식회사브로넥스.xlsx");
    expect(result.employeeCount).toBe(1);
    expect(worksheet).toBeDefined();
    expect(worksheet?.getCell("B2").value).toBe("2026년 03월");
    expect(worksheet?.getCell("D2").value).toBe("급여명세서");
    expect(worksheet?.getCell("B3").value).toBe("주식회사브로넥스");
    expect(worksheet?.getCell("D3").value).toBe("급여근속기간 : ");
    expect(worksheet?.getCell("E3").value).toBe("2026.03.01 ~ 2026.03.31");
    expect(worksheet?.getCell("C4").value).toBe("최 태 호");
    expect(worksheet?.getCell("E4").value).toBe(250421);
    expect(worksheet?.getCell("C22").value).toBe(0);
    expect(worksheet?.getCell("E23").value).toBe(0);
    expect(worksheet?.getColumn(2).width).toBe(35.75);
    expect(worksheet?.getRow(2).height).toBe(48);
    expect(worksheet?.getRow(30).height).toBe(20);
    expect(JSON.stringify(worksheet?.getCell("B7").fill)).toContain("FFD9D9D9");
  });
});
