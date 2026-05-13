import ExcelJS from "exceljs";
import JSZip from "jszip";
import { PDFDocument, PDFFont, PDFPage, rgb } from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";

export const MAX_ITEM_ROWS = 15;

const ITEM_START_ROW = 7;
const SUMMARY_ROW = ITEM_START_ROW + MAX_ITEM_ROWS;
const NET_PAY_ROW = SUMMARY_ROW + 1;
const NOTES_START_ROW = NET_PAY_ROW + 1;
const SHEET_LAST_ROW = NOTES_START_ROW + 6;

export type HireDateMap = Record<string, string>;

export type PayItemMapping = {
  source: string;
  label: string;
};

export type PayItemMappings = {
  allowances: PayItemMapping[];
  deductions: PayItemMapping[];
};

export type PayItem = {
  label: string;
  amount: number;
  columnIndex: number;
};

export type PayrollEmployee = {
  name: string;
  allowances: PayItem[];
  deductions: PayItem[];
  allowanceTotal: number;
  deductionTotal: number;
};

export type PayrollPeriod = {
  titleMonth: string;
  fileMonth: string;
  periodText: string;
  paymentDateText: string;
};

export type GeneratePayrollInput = {
  sourceBuffer: ArrayBuffer;
  templateBuffer?: ArrayBuffer;
  fontBuffer?: ArrayBuffer;
  companyName: string;
  paymentDate: string;
  hireDates: HireDateMap;
  mappings: PayItemMappings;
};

export type GeneratePayrollResult = {
  buffer: ArrayBuffer;
  fileName: string;
  employeeCount: number;
  period: PayrollPeriod;
};

export type GeneratePayrollPdfArchiveResult = GeneratePayrollResult;

type PreparedPayrollStatements = {
  company: string;
  period: PayrollPeriod;
  employees: PayrollEmployee[];
};

type SourceValue = {
  value: unknown;
  columnIndex: number;
};

type CellFontStyle = Omit<Partial<ExcelJS.Font>, "color"> & {
  color?: string | Partial<ExcelJS.Color>;
};

type CellStyle = {
  fill?: string;
  font?: CellFontStyle;
  alignment?: Partial<ExcelJS.Alignment>;
  border?: true;
};

const COLORS = {
  brown: "FF9A3A08",
  border: "FFEBD7A4",
  beige: "FFF2DDB2",
  yellow: "FFFFFFC8",
  gray: "FFD9D9D9",
  white: "FFFFFFFF"
} as const;

const ACCOUNTING_KRW_FORMAT = '_-"₩"* #,##0_-;\\-"₩"* #,##0_-;_-"₩"* "-"_-;_-@_-';
const ACCOUNTING_TEXT_FORMAT = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-';

const BORDER_SIDE: Partial<ExcelJS.Border> = {
  style: "thin",
  color: { argb: COLORS.border }
};

const DOUBLE_BORDER_SIDE: Partial<ExcelJS.Border> = {
  style: "double",
  color: { argb: COLORS.border }
};

const CELL_BORDER: Partial<ExcelJS.Borders> = {
  top: BORDER_SIDE,
  left: BORDER_SIDE,
  bottom: BORDER_SIDE,
  right: BORDER_SIDE
};

export function getPreviousPayrollPeriod(paymentDate: string): PayrollPeriod {
  const date = parseIsoDate(paymentDate, "급여지급일");
  const previousMonth = new Date(date.getFullYear(), date.getMonth() - 1, 1);
  const start = new Date(previousMonth.getFullYear(), previousMonth.getMonth(), 1);
  const end = new Date(previousMonth.getFullYear(), previousMonth.getMonth() + 1, 0);
  const year = previousMonth.getFullYear();
  const month = previousMonth.getMonth() + 1;

  return {
    titleMonth: `${year}년 ${pad2(month)}월`,
    fileMonth: `${year}년_${pad2(month)}월`,
    periodText: `${formatDateDot(start)} ~ ${formatDateDot(end)}`,
    paymentDateText: formatDateSlash(date)
  };
}

export function employeeNumberFromHireDate(hireDate: string): string {
  const date = parseIsoDate(hireDate, "입사일");
  return `${String(date.getFullYear()).slice(-2)}${pad2(date.getMonth() + 1)}${pad2(
    date.getDate()
  )}`;
}

export function formatHireDate(hireDate: string): string {
  return formatDateDot(parseIsoDate(hireDate, "입사일"));
}

export function formatEmployeeName(name: string): string {
  const trimmed = name.trim();
  return trimmed.includes(" ") ? trimmed : [...trimmed].join(" ");
}

export function collectMappedItems(
  sourceValues: Map<string, SourceValue>,
  mappings: PayItemMapping[]
): PayItem[] {
  return mappings
    .map((mapping) => {
      const source = sourceValues.get(normalizeHeader(mapping.source));
      const amount = toNumber(source?.value);
      if (!source || amount === null || amount === 0) {
        return null;
      }

      return {
        label: mapping.label,
        amount,
        columnIndex: source.columnIndex
      };
    })
    .filter((item): item is PayItem => item !== null)
    .sort((left, right) => left.columnIndex - right.columnIndex);
}

export async function parsePayrollWorkbook(
  sourceBuffer: ArrayBuffer,
  mappings: PayItemMappings
): Promise<PayrollEmployee[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(sourceBuffer);
  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    throw new Error("급여대장 파일에 시트가 없습니다.");
  }

  const identityHeaders = buildHeaderIndex(worksheet, [2, 3]);
  const itemHeaders = buildHeaderIndex(worksheet, [3]);
  const nameColumn = identityHeaders.get(normalizeHeader("사원명"));

  if (!nameColumn) {
    throw new Error("급여대장에서 '사원명' 컬럼을 찾을 수 없습니다.");
  }

  const employees: PayrollEmployee[] = [];

  for (let rowIndex = 4; rowIndex <= worksheet.rowCount; rowIndex += 1) {
    const row = worksheet.getRow(rowIndex);
    const name = getCellText(row.getCell(nameColumn).value).trim();

    if (!name) {
      continue;
    }

    const sourceValues = buildSourceValueMap(row, itemHeaders);
    const allowances = collectMappedItems(sourceValues, mappings.allowances);
    const deductions = collectMappedItems(sourceValues, mappings.deductions);
    const allowanceTotal =
      toNumber(sourceValues.get(normalizeHeader("지급액계"))?.value) ?? sumItems(allowances);
    const deductionTotal =
      toNumber(sourceValues.get(normalizeHeader("공제액계"))?.value) ?? sumItems(deductions);

    employees.push({
      name,
      allowances,
      deductions,
      allowanceTotal,
      deductionTotal
    });
  }

  if (employees.length === 0) {
    throw new Error("급여대장에서 직원 데이터를 찾을 수 없습니다.");
  }

  return employees;
}

export async function generatePayrollStatements({
  sourceBuffer,
  templateBuffer,
  companyName,
  paymentDate,
  hireDates,
  mappings
}: GeneratePayrollInput): Promise<GeneratePayrollResult> {
  const { company, period, employees } = await preparePayrollStatements({
    sourceBuffer,
    companyName,
    paymentDate,
    hireDates,
    mappings
  });

  const workbook = templateBuffer
    ? await buildTemplateWorkbook(templateBuffer, employees, hireDates, company, period)
    : buildManualWorkbook(employees, hireDates, company, period);

  const output = await patchWorkbookDefaultFont(toArrayBuffer(await workbook.xlsx.writeBuffer()));

  return {
    buffer: output,
    fileName: `${period.fileMonth}_급여명세서_${sanitizeFileName(company)}.xlsx`,
    employeeCount: employees.length,
    period
  };
}

export async function generatePayrollPdfArchive({
  sourceBuffer,
  fontBuffer,
  companyName,
  paymentDate,
  hireDates,
  mappings
}: GeneratePayrollInput): Promise<GeneratePayrollPdfArchiveResult> {
  const { company, period, employees } = await preparePayrollStatements({
    sourceBuffer,
    companyName,
    paymentDate,
    hireDates,
    mappings
  });
  const resolvedFontBuffer = fontBuffer ?? (await loadDefaultPdfFontBuffer());
  const zip = new JSZip();

  await Promise.all(
    employees.map(async (employee) => {
      const pdf = await buildEmployeePdf(employee, {
        company,
        hireDate: hireDates[employee.name],
        period,
        fontBuffer: resolvedFontBuffer
      });

      zip.file(
        `${period.titleMonth}_급여명세서_${sanitizeFileName(company)}(${sanitizeFileName(
          employee.name
        )}).pdf`,
        pdf
      );
    })
  );

  const output = await zip.generateAsync({ type: "arraybuffer" });

  return {
    buffer: output,
    fileName: `${period.fileMonth}_급여명세서_${sanitizeFileName(company)}_PDF.zip`,
    employeeCount: employees.length,
    period
  };
}

async function preparePayrollStatements({
  sourceBuffer,
  companyName,
  paymentDate,
  hireDates,
  mappings
}: GeneratePayrollInput): Promise<PreparedPayrollStatements> {
  const company = companyName.trim();
  if (!company) {
    throw new Error("회사명을 입력해야 합니다.");
  }

  const period = getPreviousPayrollPeriod(paymentDate);
  const employees = await parsePayrollWorkbook(sourceBuffer, mappings);
  const missingHireDates = employees
    .map((employee) => employee.name)
    .filter((name) => !hireDates[name]);

  if (missingHireDates.length > 0) {
    throw new Error(`입사일 JSON에 없는 직원: ${unique(missingHireDates).join(", ")}`);
  }

  const overflow = employees.find(
    (employee) =>
      employee.allowances.length > MAX_ITEM_ROWS || employee.deductions.length > MAX_ITEM_ROWS
  );

  if (overflow) {
    throw new Error(
      `${overflow.name} 직원의 수당 또는 공제 항목이 ${MAX_ITEM_ROWS}개를 초과했습니다. 매핑 JSON을 조정해 주세요.`
    );
  }

  return {
    company,
    period,
    employees
  };
}

const PDF_FONT_PATH = "/fonts/NanumGothic.ttc";
const PDF_FONT_REGULAR = "NanumGothic";
const PDF_FONT_BOLD = "NanumGothicBold";
const PDF_PAGE_WIDTH = 595.28;
const PDF_PAGE_HEIGHT = 841.89;
const PDF_TABLE_WIDTH = 560;
const PDF_TABLE_LEFT = (PDF_PAGE_WIDTH - PDF_TABLE_WIDTH) / 2;
const PDF_START_Y = 800;
const PDF_COLUMN_WIDTH_UNITS = [2.75, 35.75, 27.75, 35.75, 27.75, 2.75];
const PDF_COLUMN_WIDTHS = scalePdfColumnWidths(PDF_COLUMN_WIDTH_UNITS, PDF_TABLE_WIDTH);

type PdfStatementContext = {
  company: string;
  hireDate: string;
  period: PayrollPeriod;
  fontBuffer: ArrayBuffer;
};

type PdfFonts = {
  regular: PDFFont;
  bold: PDFFont;
};

type PdfRect = {
  x: number;
  y: number;
  width: number;
  height: number;
};

type PdfCellOptions = {
  fill?: ReturnType<typeof rgb>;
  border?: boolean;
};

type PdfTextOptions = {
  align?: "left" | "center" | "right";
  color?: ReturnType<typeof rgb>;
  font?: PDFFont;
  minSize?: number;
  padding?: number;
  size?: number;
};

async function buildEmployeePdf(
  employee: PayrollEmployee,
  context: PdfStatementContext
): Promise<Uint8Array> {
  const pdfDoc = await PDFDocument.create();
  const regular = await embedKoreanFont(pdfDoc, context.fontBuffer, PDF_FONT_REGULAR);
  const bold = await embedKoreanFont(pdfDoc, context.fontBuffer, PDF_FONT_BOLD);
  const fonts = { regular, bold };
  const page = pdfDoc.addPage([PDF_PAGE_WIDTH, PDF_PAGE_HEIGHT]);
  const rows = buildPdfRowRects();
  const columns = buildPdfColumnRects();
  const color = colorFromArgb(COLORS.brown);
  const borderColor = colorFromArgb(COLORS.border);

  page.drawRectangle({
    x: 0,
    y: 0,
    width: PDF_PAGE_WIDTH,
    height: PDF_PAGE_HEIGHT,
    color: rgb(1, 1, 1)
  });
  drawOuterPdfFrame(page, rows, columns, borderColor);

  drawPdfText(page, `${context.period.titleMonth}`, mergePdfCells(rows, columns, 2, 2, 3), fonts, {
    align: "right",
    color,
    font: fonts.regular,
    size: 20,
    padding: 8
  });
  drawPdfText(page, "급여명세서", mergePdfCells(rows, columns, 2, 4, 5), fonts, {
    align: "left",
    color,
    font: fonts.regular,
    size: 20,
    padding: 8
  });

  drawPdfText(page, context.company, pdfCell(rows, columns, 3, 2), fonts, {
    align: "left",
    color,
    size: 9.5
  });
  drawPdfText(page, "급여근속기간 : ", pdfCell(rows, columns, 3, 4), fonts, {
    align: "right",
    color,
    size: 9.5
  });
  drawPdfText(page, context.period.periodText, pdfCell(rows, columns, 3, 5), fonts, {
    align: "center",
    color,
    minSize: 7.2,
    size: 9.2
  });
  drawPdfLine(
    page,
    pdfCell(rows, columns, 3, 2).x,
    pdfCell(rows, columns, 3, 2).y,
    pdfCell(rows, columns, 3, 3).x + pdfCell(rows, columns, 3, 3).width,
    pdfCell(rows, columns, 3, 2).y,
    borderColor
  );

  drawLabeledPdfRow(
    page,
    rows,
    columns,
    fonts,
    4,
    ["직원 이름", formatEmployeeName(employee.name), "직원 번호", employeeNumberFromHireDate(context.hireDate)],
    color,
    colorFromArgb(COLORS.beige)
  );

  drawPdfMergedCell(page, rows, columns, 6, 2, 3, colorFromArgb(COLORS.yellow), borderColor);
  drawPdfMergedCell(page, rows, columns, 6, 4, 5, colorFromArgb(COLORS.yellow), borderColor);
  drawPdfText(page, "급여내역", mergePdfCells(rows, columns, 6, 2, 3), fonts, {
    align: "center",
    color,
    font: fonts.bold,
    size: 9.5
  });
  drawPdfText(page, "원천징수공제내역", mergePdfCells(rows, columns, 6, 4, 5), fonts, {
    align: "center",
    color,
    font: fonts.bold,
    size: 9.5
  });

  for (let offset = 0; offset < MAX_ITEM_ROWS; offset += 1) {
    const row = ITEM_START_ROW + offset;
    const allowance = employee.allowances[offset];
    const deduction = employee.deductions[offset];

    drawPdfItemPair(page, rows, columns, fonts, row, 2, allowance, color, borderColor);
    drawPdfItemPair(page, rows, columns, fonts, row, 4, deduction, color, borderColor);
  }

  drawLabeledPdfRow(
    page,
    rows,
    columns,
    fonts,
    SUMMARY_ROW,
    ["급여 합계", formatPdfCurrency(employee.allowanceTotal), "원천징수공제내역", formatPdfCurrency(employee.deductionTotal)],
    color,
    colorFromArgb(COLORS.beige),
    true
  );

  drawPdfCell(page, mergePdfCells(rows, columns, NET_PAY_ROW, 2, 3), { border: false });
  drawPdfCell(page, pdfCell(rows, columns, NET_PAY_ROW, 4), {
    border: true,
    fill: colorFromArgb(COLORS.yellow)
  });
  drawPdfCell(page, pdfCell(rows, columns, NET_PAY_ROW, 5), {
    border: true,
    fill: colorFromArgb(COLORS.yellow)
  });
  drawPdfText(
    page,
    `실지급 급여 [ ${context.period.paymentDateText} ]`,
    pdfCell(rows, columns, NET_PAY_ROW, 4),
    fonts,
    {
      align: "center",
      color,
      font: fonts.bold,
      minSize: 7.5,
      size: 9.5
    }
  );
  drawPdfText(
    page,
    formatPdfCurrency(employee.allowanceTotal - employee.deductionTotal),
    pdfCell(rows, columns, NET_PAY_ROW, 5),
    fonts,
    {
      align: "right",
      color,
      font: fonts.bold,
      size: 9.5
    }
  );

  drawPdfNotes(page, rows, columns, fonts, context.hireDate, color);

  return pdfDoc.save();
}

function drawLabeledPdfRow(
  page: PDFPage,
  rows: Map<number, PdfRect>,
  columns: Map<number, PdfRect>,
  fonts: PdfFonts,
  row: number,
  values: [string, string, string, string],
  color: ReturnType<typeof rgb>,
  fill: ReturnType<typeof rgb>,
  bold = false
) {
  [2, 3, 4, 5].forEach((column) => {
    drawPdfCell(page, pdfCell(rows, columns, row, column), { border: true, fill });
  });

  const valueFont = bold ? fonts.bold : fonts.regular;
  drawPdfText(page, values[0], pdfCell(rows, columns, row, 2), fonts, {
    align: "center",
    color,
    font: valueFont,
    size: 9.3
  });
  drawPdfText(page, values[1], pdfCell(rows, columns, row, 3), fonts, {
    align: row === SUMMARY_ROW ? "right" : "center",
    color,
    font: valueFont,
    size: 9.3
  });
  drawPdfText(page, values[2], pdfCell(rows, columns, row, 4), fonts, {
    align: "center",
    color,
    font: valueFont,
    minSize: 7.5,
    size: 9.3
  });
  drawPdfText(page, values[3], pdfCell(rows, columns, row, 5), fonts, {
    align: "right",
    color,
    font: valueFont,
    size: 9.3
  });
}

function drawPdfItemPair(
  page: PDFPage,
  rows: Map<number, PdfRect>,
  columns: Map<number, PdfRect>,
  fonts: PdfFonts,
  row: number,
  startColumn: number,
  item: PayItem | undefined,
  color: ReturnType<typeof rgb>,
  borderColor: ReturnType<typeof rgb>
) {
  const labelCell = pdfCell(rows, columns, row, startColumn);
  const amountCell = pdfCell(rows, columns, row, startColumn + 1);
  const fill = item ? undefined : colorFromArgb(COLORS.gray);

  drawPdfCell(page, labelCell, { border: true, fill });
  drawPdfCell(page, amountCell, { border: true, fill });

  if (!item) {
    return;
  }

  drawPdfText(page, item.label, labelCell, fonts, {
    align: "left",
    color,
    size: 9.2
  });
  drawPdfText(page, formatPdfCurrency(item.amount), amountCell, fonts, {
    align: "right",
    color,
    size: 9.2
  });

  drawPdfLine(page, amountCell.x, amountCell.y, amountCell.x, amountCell.y + amountCell.height, borderColor);
}

function drawPdfNotes(
  page: PDFPage,
  rows: Map<number, PdfRect>,
  columns: Map<number, PdfRect>,
  fonts: PdfFonts,
  hireDate: string,
  color: ReturnType<typeof rgb>
) {
  const notes = [
    `1. 입사일: ${formatHireDate(hireDate)}`,
    "     : 비과세 항목에 8세이하 자녀 육아수당 반영",
    "2. 2017년 건강보험료 확정분(3.06%) 반영_2016년 보수총액 신고액 기준",
    "3. 2016년 국민연금보험료 확정 적용"
  ];

  notes.forEach((note, index) => {
    drawPdfText(page, note, mergePdfCells(rows, columns, NOTES_START_ROW + index, 2, 5), fonts, {
      align: "left",
      color,
      minSize: 7.2,
      size: 8.8
    });
  });
}

function drawOuterPdfFrame(
  page: PDFPage,
  rows: Map<number, PdfRect>,
  columns: Map<number, PdfRect>,
  borderColor: ReturnType<typeof rgb>
) {
  const left = pdfCell(rows, columns, 2, 1).x;
  const right = pdfCell(rows, columns, 2, 6).x + pdfCell(rows, columns, 2, 6).width;
  const top = pdfCell(rows, columns, 2, 1).y + pdfCell(rows, columns, 2, 1).height;
  const bottom = pdfCell(rows, columns, SHEET_LAST_ROW, 1).y;

  drawPdfLine(page, left, top, right, top, borderColor, 1.2);
  drawPdfLine(page, left, top - 3, right, top - 3, borderColor, 0.8);
  drawPdfLine(page, left, bottom, right, bottom, borderColor);
  drawPdfLine(page, left, bottom, left, top, borderColor);
  drawPdfLine(page, right, bottom, right, top, borderColor);
}

function drawPdfMergedCell(
  page: PDFPage,
  rows: Map<number, PdfRect>,
  columns: Map<number, PdfRect>,
  row: number,
  startColumn: number,
  endColumn: number,
  fill: ReturnType<typeof rgb>,
  borderColor: ReturnType<typeof rgb>
) {
  const rect = mergePdfCells(rows, columns, row, startColumn, endColumn);
  page.drawRectangle({
    ...rect,
    color: fill,
    borderColor,
    borderWidth: 0.8
  });
}

function drawPdfCell(page: PDFPage, rect: PdfRect, options: PdfCellOptions = {}) {
  page.drawRectangle({
    ...rect,
    color: options.fill,
    borderColor: options.border ? colorFromArgb(COLORS.border) : undefined,
    borderWidth: options.border ? 0.8 : undefined
  });
}

function drawPdfText(
  page: PDFPage,
  text: string,
  rect: PdfRect,
  fonts: PdfFonts,
  options: PdfTextOptions = {}
) {
  const font = options.font ?? fonts.regular;
  const padding = options.padding ?? 6;
  const size = fitPdfFontSize(
    font,
    text,
    rect.width - padding * 2,
    options.size ?? 9.5,
    options.minSize ?? 7
  );
  const textWidth = font.widthOfTextAtSize(text, size);
  const align = options.align ?? "left";
  const x =
    align === "right"
      ? rect.x + rect.width - padding - textWidth
      : align === "center"
        ? rect.x + (rect.width - textWidth) / 2
        : rect.x + padding;
  const y = rect.y + (rect.height - size) / 2 + 1.4;

  page.drawText(text, {
    x,
    y,
    size,
    font,
    color: options.color ?? colorFromArgb(COLORS.brown)
  });
}

function drawPdfLine(
  page: PDFPage,
  startX: number,
  startY: number,
  endX: number,
  endY: number,
  color: ReturnType<typeof rgb>,
  thickness = 0.8
) {
  page.drawLine({
    start: { x: startX, y: startY },
    end: { x: endX, y: endY },
    color,
    thickness
  });
}

function buildPdfRowRects() {
  const rowRects = new Map<number, PdfRect>();
  const heights = buildPdfRowHeights();
  let top = PDF_START_Y;

  for (let row = 1; row <= SHEET_LAST_ROW; row += 1) {
    const height = heights.get(row) ?? 20;
    rowRects.set(row, {
      x: PDF_TABLE_LEFT,
      y: top - height,
      width: PDF_TABLE_WIDTH,
      height
    });
    top -= height;
  }

  return rowRects;
}

function buildPdfColumnRects() {
  const columnRects = new Map<number, PdfRect>();
  let left = PDF_TABLE_LEFT;

  PDF_COLUMN_WIDTHS.forEach((width, index) => {
    columnRects.set(index + 1, {
      x: left,
      y: 0,
      width,
      height: PDF_PAGE_HEIGHT
    });
    left += width;
  });

  return columnRects;
}

function buildPdfRowHeights() {
  const rowHeights = new Map<number, number>([
    [1, 45],
    [2, 48],
    [3, 20],
    [4, 21.75],
    [5, 12],
    [6, 21.75],
    [NET_PAY_ROW + 2, 18.75]
  ]);

  for (let row = ITEM_START_ROW; row < SUMMARY_ROW; row += 1) {
    rowHeights.set(row, row <= 13 ? 20 : 18.75);
  }

  for (let row = SUMMARY_ROW; row <= SHEET_LAST_ROW; row += 1) {
    if (!rowHeights.has(row)) {
      rowHeights.set(row, 20);
    }
  }

  return rowHeights;
}

function pdfCell(
  rows: Map<number, PdfRect>,
  columns: Map<number, PdfRect>,
  row: number,
  column: number
): PdfRect {
  const rowRect = rows.get(row);
  const columnRect = columns.get(column);

  if (!rowRect || !columnRect) {
    throw new Error("PDF 양식 좌표를 계산하지 못했습니다.");
  }

  return {
    x: columnRect.x,
    y: rowRect.y,
    width: columnRect.width,
    height: rowRect.height
  };
}

function mergePdfCells(
  rows: Map<number, PdfRect>,
  columns: Map<number, PdfRect>,
  row: number,
  startColumn: number,
  endColumn: number
): PdfRect {
  const start = pdfCell(rows, columns, row, startColumn);
  const end = pdfCell(rows, columns, row, endColumn);

  return {
    x: start.x,
    y: start.y,
    width: end.x + end.width - start.x,
    height: start.height
  };
}

function fitPdfFontSize(
  font: PDFFont,
  text: string,
  maxWidth: number,
  preferredSize: number,
  minSize: number
) {
  let size = preferredSize;

  while (size > minSize && font.widthOfTextAtSize(text, size) > maxWidth) {
    size -= 0.2;
  }

  return Math.max(size, minSize);
}

async function embedKoreanFont(
  pdfDoc: PDFDocument,
  fontBuffer: ArrayBuffer,
  postscriptName: string
) {
  pdfDoc.registerFontkit({
    ...fontkit,
    create: (data: Uint8Array, requestedPostscriptName?: string) =>
      fontkit.create(data, requestedPostscriptName ?? postscriptName)
  });

  return pdfDoc.embedFont(fontBuffer, { subset: true });
}

async function loadDefaultPdfFontBuffer() {
  const response = await fetch(PDF_FONT_PATH);

  if (!response.ok) {
    throw new Error("PDF 한글 폰트 파일을 불러오지 못했습니다.");
  }

  return response.arrayBuffer();
}

function formatPdfCurrency(amount: number) {
  const rounded = Math.round(amount);
  const absolute = Math.abs(rounded).toLocaleString("ko-KR");
  return rounded < 0 ? `-₩ ${absolute}` : `₩ ${absolute}`;
}

function colorFromArgb(argb: string) {
  const hex = argb.slice(2);
  const red = Number.parseInt(hex.slice(0, 2), 16) / 255;
  const green = Number.parseInt(hex.slice(2, 4), 16) / 255;
  const blue = Number.parseInt(hex.slice(4, 6), 16) / 255;
  return rgb(red, green, blue);
}

function scalePdfColumnWidths(widths: number[], totalWidth: number) {
  const totalUnits = widths.reduce((sum, width) => sum + width, 0);
  return widths.map((width) => (width / totalUnits) * totalWidth);
}

async function buildTemplateWorkbook(
  templateBuffer: ArrayBuffer,
  employees: PayrollEmployee[],
  hireDates: HireDateMap,
  company: string,
  period: PayrollPeriod
): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  const templateWorksheet =
    workbook.getWorksheet("3월") ?? workbook.worksheets[workbook.worksheets.length - 1];

  if (!templateWorksheet) {
    throw new Error("급여명세서 템플릿 시트를 찾을 수 없습니다.");
  }

  const originalWorksheetIds = workbook.worksheets.map((worksheet) => worksheet.id);
  const usedSheetNames = new Set<string>();

  employees.forEach((employee) => {
    const hireDate = hireDates[employee.name];
    const sheetName = createUniqueSheetName(employee.name, usedSheetNames);
    const worksheet = cloneWorksheet(workbook, templateWorksheet, sheetName);

    fillTemplateWorksheet(worksheet, employee, {
      company,
      hireDate,
      period
    });
  });

  originalWorksheetIds.forEach((id) => {
    workbook.removeWorksheet(id);
  });
  workbook.views = [
    {
      x: 0,
      y: 0,
      width: 10000,
      height: 20000,
      activeTab: 0,
      firstSheet: 0,
      visibility: "visible"
    }
  ];

  workbook.creator = "Hangyeol Pay";
  workbook.created = new Date();
  workbook.modified = new Date();

  return workbook;
}

function cloneWorksheet(
  workbook: ExcelJS.Workbook,
  templateWorksheet: ExcelJS.Worksheet,
  sheetName: string
): ExcelJS.Worksheet {
  const worksheet = workbook.addWorksheet(sheetName);
  const model = JSON.parse(JSON.stringify(templateWorksheet.model)) as ExcelJS.WorksheetModel;
  const merges = [...(model.merges ?? [])];

  model.id = worksheet.id;
  model.name = sheetName;
  model.merges = [];
  worksheet.model = model;
  merges.forEach((range) => worksheet.mergeCells(range));

  return worksheet;
}

function fillTemplateWorksheet(
  worksheet: ExcelJS.Worksheet,
  employee: PayrollEmployee,
  context: {
    company: string;
    hireDate: string;
    period: PayrollPeriod;
  }
) {
  worksheet.getCell("B2").value = context.period.titleMonth;
  worksheet.getCell("D2").value = "급여명세서";
  worksheet.getCell("B3").value = context.company;
  worksheet.getCell("D3").value = "급여근속기간 : ";
  worksheet.getCell("E3").value = context.period.periodText;

  setRowValues(worksheet, 4, {
    B: "직원 이름",
    C: formatEmployeeName(employee.name),
    D: "직원 번호",
    E: Number(employeeNumberFromHireDate(context.hireDate))
  });

  worksheet.getCell("B6").value = "급여내역";
  worksheet.getCell("D6").value = "원천징수공제내역";

  const allowanceLabelStyle = worksheet.getCell("B7").style;
  const allowanceAmountStyle = worksheet.getCell("C7").style;
  const deductionLabelStyle = worksheet.getCell("D7").style;
  const deductionAmountStyle = worksheet.getCell("E7").style;
  const emptyDeductionLabelStyle = worksheet.getCell("D13").style;
  const emptyDeductionAmountStyle = worksheet.getCell("E13").style;

  for (let offset = 0; offset < MAX_ITEM_ROWS; offset += 1) {
    const rowNumber = ITEM_START_ROW + offset;
    const allowance = employee.allowances[offset];
    const deduction = employee.deductions[offset];

    setCellStyle(worksheet.getCell(`B${rowNumber}`), allowanceLabelStyle);
    setCellStyle(worksheet.getCell(`C${rowNumber}`), allowanceAmountStyle);
    if (!allowance) {
      setCellFill(worksheet.getCell(`B${rowNumber}`), COLORS.gray);
      setCellFill(worksheet.getCell(`C${rowNumber}`), COLORS.gray);
    }
    worksheet.getCell(`B${rowNumber}`).value = allowance?.label ?? null;
    worksheet.getCell(`C${rowNumber}`).value = allowance?.amount ?? null;

    setCellStyle(
      worksheet.getCell(`D${rowNumber}`),
      deduction ? deductionLabelStyle : emptyDeductionLabelStyle
    );
    setCellStyle(
      worksheet.getCell(`E${rowNumber}`),
      deduction ? deductionAmountStyle : emptyDeductionAmountStyle
    );
    worksheet.getCell(`D${rowNumber}`).value = deduction?.label ?? null;
    worksheet.getCell(`E${rowNumber}`).value = deduction?.amount ?? null;
  }

  setRowValues(worksheet, SUMMARY_ROW, {
    B: "급여 합계",
    C: employee.allowanceTotal,
    D: "원천징수공제내역",
    E: employee.deductionTotal
  });

  setRowValues(worksheet, NET_PAY_ROW, {
    D: `     실지급 급여 [ ${context.period.paymentDateText} ]`,
    E: employee.allowanceTotal - employee.deductionTotal
  });

  worksheet.getCell(`B${NOTES_START_ROW}`).value = `1. 입사일: ${formatHireDate(context.hireDate)}`;
  worksheet.getCell(`B${NOTES_START_ROW + 1}`).value =
    "     : 비과세 항목에 8세이하 자녀 육아수당 반영";
  worksheet.getCell(`B${NOTES_START_ROW + 2}`).value =
    "2. 2017년 건강보험료 확정분(3.06%) 반영_2016년 보수총액 신고액 기준";
  worksheet.getCell(`B${NOTES_START_ROW + 3}`).value = "3. 2016년 국민연금보험료 확정 적용";
  worksheet.getCell(`B${NOTES_START_ROW + 4}`).value = null;
  worksheet.getCell(`B${NOTES_START_ROW + 5}`).value = null;
}

function buildManualWorkbook(
  employees: PayrollEmployee[],
  hireDates: HireDateMap,
  company: string,
  period: PayrollPeriod
): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Hangyeol Pay";
  workbook.created = new Date();
  workbook.modified = new Date();

  const usedSheetNames = new Set<string>();
  employees.forEach((employee) => {
    const hireDate = hireDates[employee.name];
    const sheetName = createUniqueSheetName(employee.name, usedSheetNames);
    const worksheet = workbook.addWorksheet(sheetName, {
      pageSetup: {
        paperSize: 9,
        orientation: "portrait",
        fitToPage: true,
        fitToHeight: 0,
        scale: 86,
        horizontalCentered: true,
        printArea: `A1:F${SHEET_LAST_ROW}`,
        margins: {
          left: 0.5905511811023623,
          right: 0.5905511811023623,
          top: 0.3937007874015748,
          bottom: 0.3937007874015748,
          header: 0.1,
          footer: 0.1
        }
      },
      views: [{ showGridLines: false }]
    });

    buildEmployeeSheet(worksheet, employee, {
      company,
      hireDate,
      period
    });
  });

  return workbook;
}

async function patchWorkbookDefaultFont(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer);
  const stylesFile = zip.file("xl/styles.xml");

  if (!stylesFile) {
    return buffer;
  }

  const stylesXml = await stylesFile.async("string");
  const patchedStylesXml = stylesXml.replace(
    '<font><color theme="1"/><family val="2"/><scheme val="minor"/><sz val="11"/><name val="Calibri"/></font>',
    '<font><color theme="1"/><family val="2"/><sz val="10"/><name val="Arial"/></font>'
  );

  if (patchedStylesXml === stylesXml) {
    return patchWorksheetPageSetup(zip, buffer);
  }

  zip.file("xl/styles.xml", patchedStylesXml);
  return patchWorksheetPageSetup(zip, buffer);
}

async function patchWorksheetPageSetup(
  zip: JSZip,
  fallbackBuffer: ArrayBuffer
): Promise<ArrayBuffer> {
  let changed = false;

  await Promise.all(
    Object.keys(zip.files)
      .filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/.test(path))
      .map(async (path) => {
        const file = zip.file(path);
        if (!file) {
          return;
        }

        const xml = await file.async("string");
        const patchedXml = xml
          .replace(/ horizontalDpi="4294967295" verticalDpi="4294967295"/g, "")
          .replace(/ fitToWidth="1"/g, "")
          .replace(/ header="0.1" footer="0.1"/g, ' header="0.19685039370078741" footer="0.19685039370078741"');

        if (patchedXml !== xml) {
          changed = true;
          zip.file(path, patchedXml);
        }
      })
  );

  if (!changed) {
    return fallbackBuffer;
  }

  return zip.generateAsync({ type: "arraybuffer" });
}

function buildEmployeeSheet(
  worksheet: ExcelJS.Worksheet,
  employee: PayrollEmployee,
  context: {
    company: string;
    hireDate: string;
    period: PayrollPeriod;
  }
) {
  worksheet.properties.defaultRowHeight = 24;
  worksheet.columns = [
    { width: 2.75 },
    { width: 35.75 },
    { width: 27.75 },
    { width: 35.75 },
    { width: 27.75 },
    { width: 2.75 }
  ];

  applyManualRowHeights(worksheet);

  mergeAndSet(worksheet, "B2:C2", context.period.titleMonth);
  mergeAndSet(worksheet, "D2:E2", "급여명세서");

  worksheet.getCell("B3").value = context.company;
  worksheet.getCell("D3").value = "급여근속기간 : ";
  worksheet.getCell("E3").value = context.period.periodText;

  setRowValues(worksheet, 4, {
    B: "직원 이름",
    C: formatEmployeeName(employee.name),
    D: "직원 번호",
    E: Number(employeeNumberFromHireDate(context.hireDate))
  });

  mergeAndSet(worksheet, "B6:C6", "급여내역");
  mergeAndSet(worksheet, "D6:E6", "원천징수공제내역");

  for (let offset = 0; offset < MAX_ITEM_ROWS; offset += 1) {
    const rowNumber = ITEM_START_ROW + offset;
    const allowance = employee.allowances[offset];
    const deduction = employee.deductions[offset];

    worksheet.getCell(`B${rowNumber}`).value = allowance?.label ?? null;
    worksheet.getCell(`C${rowNumber}`).value = allowance?.amount ?? null;
    worksheet.getCell(`D${rowNumber}`).value = deduction?.label ?? null;
    worksheet.getCell(`E${rowNumber}`).value = deduction?.amount ?? null;
  }

  setRowValues(worksheet, SUMMARY_ROW, {
    B: "급여 합계",
    C: employee.allowanceTotal,
    D: "원천징수공제내역",
    E: employee.deductionTotal
  });

  mergeAndSet(worksheet, `B${NET_PAY_ROW}:C${NET_PAY_ROW}`, null);
  setRowValues(worksheet, NET_PAY_ROW, {
    D: `     실지급 급여 [ ${context.period.paymentDateText} ]`,
    E: employee.allowanceTotal - employee.deductionTotal
  });

  mergeAndSet(
    worksheet,
    `B${NOTES_START_ROW}:E${NOTES_START_ROW}`,
    `1. 입사일: ${formatHireDate(context.hireDate)}`
  );
  mergeAndSet(
    worksheet,
    `B${NOTES_START_ROW + 1}:E${NOTES_START_ROW + 1}`,
    "     : 비과세 항목에 8세이하 자녀 육아수당 반영"
  );
  mergeAndSet(
    worksheet,
    `B${NOTES_START_ROW + 2}:E${NOTES_START_ROW + 2}`,
    "2. 2017년 건강보험료 확정분(3.06%) 반영_2016년 보수총액 신고액 기준"
  );
  mergeAndSet(
    worksheet,
    `B${NOTES_START_ROW + 3}:E${NOTES_START_ROW + 3}`,
    "3. 2016년 국민연금보험료 확정 적용"
  );
  mergeAndSet(worksheet, `B${NOTES_START_ROW + 4}:E${NOTES_START_ROW + 4}`, null);
  mergeAndSet(worksheet, `B${NOTES_START_ROW + 5}:E${NOTES_START_ROW + 5}`, null);
  mergeAndSet(worksheet, `B${SHEET_LAST_ROW}:C${SHEET_LAST_ROW}`, null);

  applyStatementTemplateStyles(worksheet, employee);
}

function applyManualRowHeights(worksheet: ExcelJS.Worksheet) {
  const rowHeights = new Map<number, number>([
    [1, 45],
    [2, 48],
    [3, 20],
    [4, 21.75],
    [5, 12],
    [6, 21.75],
    [NET_PAY_ROW + 2, 18.75]
  ]);

  for (let row = ITEM_START_ROW; row < SUMMARY_ROW; row += 1) {
    rowHeights.set(row, row <= 13 ? 20 : 18.75);
  }

  for (let row = SUMMARY_ROW; row <= SHEET_LAST_ROW; row += 1) {
    if (!rowHeights.has(row)) {
      rowHeights.set(row, 20);
    }
  }

  rowHeights.forEach((height, row) => {
    worksheet.getRow(row).height = height;
  });
}

function applyStatementTemplateStyles(worksheet: ExcelJS.Worksheet, employee: PayrollEmployee) {
  styleRange(worksheet, 1, SHEET_LAST_ROW, 1, 6, {
    font: { color: COLORS.brown, name: "Arial", size: 10 },
    alignment: { vertical: "middle" }
  });

  styleRange(worksheet, 2, 2, 2, 5, {
    font: { size: 20 },
    alignment: { vertical: "middle" }
  });
  worksheet.getCell("B2").alignment = { horizontal: "right", vertical: "middle" };
  worksheet.getCell("D2").alignment = { horizontal: "left", vertical: "middle" };

  worksheet.getCell("B3").alignment = { horizontal: "left", vertical: "middle" };
  worksheet.getCell("D3").alignment = { horizontal: "right", vertical: "middle" };
  worksheet.getCell("E3").alignment = { horizontal: "center", vertical: "middle" };

  styleRange(worksheet, 4, 4, 2, 5, {
    fill: COLORS.beige
  });
  worksheet.getCell("C4").font = {
    ...worksheet.getCell("C4").font,
    name: "Malgun Gothic",
    color: { argb: COLORS.brown }
  };
  worksheet.getCell("E4").alignment = { horizontal: "right", vertical: "middle" };

  styleRange(worksheet, 6, 6, 2, 5, {
    fill: COLORS.yellow,
    alignment: { horizontal: "center", vertical: "middle" }
  });

  for (let row = ITEM_START_ROW; row <= NET_PAY_ROW; row += 1) {
    worksheet.getCell(`C${row}`).numFmt = ACCOUNTING_KRW_FORMAT;
    worksheet.getCell(`E${row}`).numFmt = ACCOUNTING_KRW_FORMAT;
  }

  for (let offset = 0; offset < MAX_ITEM_ROWS; offset += 1) {
    const row = ITEM_START_ROW + offset;
    if (!employee.allowances[offset]) {
      styleRange(worksheet, row, row, 2, 3, {
        fill: COLORS.gray
      });
    }
    if (!employee.deductions[offset]) {
      styleRange(worksheet, row, row, 4, 5, {
        fill: COLORS.gray
      });
    }
  }

  styleRange(worksheet, SUMMARY_ROW, SUMMARY_ROW, 2, 5, {
    fill: COLORS.beige,
    font: { bold: true }
  });
  styleRange(worksheet, NET_PAY_ROW, NET_PAY_ROW, 4, 5, {
    fill: COLORS.yellow,
    font: { bold: true }
  });

  for (let row = 1; row <= SHEET_LAST_ROW; row += 1) {
    for (let column = 1; column <= 6; column += 1) {
      worksheet.getCell(row, column).numFmt =
        column === 3 || column === 5 ? worksheet.getCell(row, column).numFmt : ACCOUNTING_TEXT_FORMAT;
    }
  }

  worksheet.getCell("E4").numFmt = "General";

  applyTemplateBorders(worksheet);
}

function applyTemplateBorders(worksheet: ExcelJS.Worksheet) {
  for (let row = 2; row <= SHEET_LAST_ROW; row += 1) {
    setCellBorder(worksheet.getCell(`A${row}`), { left: BORDER_SIDE });
    setCellBorder(worksheet.getCell(`F${row}`), { right: BORDER_SIDE });
  }

  for (let column = 1; column <= 6; column += 1) {
    setCellBorder(worksheet.getCell(2, column), { top: DOUBLE_BORDER_SIDE });
    setCellBorder(worksheet.getCell(SHEET_LAST_ROW, column), { bottom: BORDER_SIDE });
  }
  setCellBorder(worksheet.getCell(`A${SHEET_LAST_ROW}`), {
    left: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
  setCellBorder(worksheet.getCell(`F${SHEET_LAST_ROW}`), {
    right: BORDER_SIDE,
    bottom: BORDER_SIDE
  });

  setCellBorder(worksheet.getCell("B3"), { bottom: BORDER_SIDE });
  setCellBorder(worksheet.getCell("C3"), { bottom: BORDER_SIDE });

  setCellBorder(worksheet.getCell("B4"), {
    left: BORDER_SIDE,
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
  setCellBorder(worksheet.getCell("C4"), {
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
  setCellBorder(worksheet.getCell("D4"), {
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
  setCellBorder(worksheet.getCell("E4"), {
    right: BORDER_SIDE,
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });

  setSectionBorder(worksheet, 6);

  for (let row = ITEM_START_ROW; row <= SUMMARY_ROW; row += 1) {
    setTableRowBorder(worksheet, row);
  }

  setCellBorder(worksheet.getCell(`B${NET_PAY_ROW}`), { top: BORDER_SIDE });
  setCellBorder(worksheet.getCell(`C${NET_PAY_ROW}`), { top: BORDER_SIDE });
  setCellBorder(worksheet.getCell(`D${NET_PAY_ROW}`), {
    left: BORDER_SIDE,
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
  setCellBorder(worksheet.getCell(`E${NET_PAY_ROW}`), {
    right: BORDER_SIDE,
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
}

function setSectionBorder(worksheet: ExcelJS.Worksheet, row: number) {
  setCellBorder(worksheet.getCell(`B${row}`), {
    left: BORDER_SIDE,
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
  setCellBorder(worksheet.getCell(`C${row}`), {
    right: BORDER_SIDE,
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
  setCellBorder(worksheet.getCell(`D${row}`), {
    left: BORDER_SIDE,
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
  setCellBorder(worksheet.getCell(`E${row}`), {
    right: BORDER_SIDE,
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
}

function setTableRowBorder(worksheet: ExcelJS.Worksheet, row: number) {
  setCellBorder(worksheet.getCell(`B${row}`), {
    left: BORDER_SIDE,
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
  setCellBorder(worksheet.getCell(`C${row}`), {
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
  setCellBorder(worksheet.getCell(`D${row}`), {
    left: BORDER_SIDE,
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
  setCellBorder(worksheet.getCell(`E${row}`), {
    right: BORDER_SIDE,
    top: BORDER_SIDE,
    bottom: BORDER_SIDE
  });
}

function setCellBorder(cell: ExcelJS.Cell, border: Partial<ExcelJS.Borders>) {
  cell.border = {
    ...cell.border,
    ...border
  };
}

function setCellStyle(cell: ExcelJS.Cell, style: Partial<ExcelJS.Style>) {
  cell.style = JSON.parse(JSON.stringify(style));
}

function setCellFill(cell: ExcelJS.Cell, color: string) {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: color }
  };
}

function buildHeaderIndex(worksheet: ExcelJS.Worksheet, rowNumbers: number[]) {
  const headers = new Map<string, number>();

  rowNumbers.forEach((rowNumber) => {
    const row = worksheet.getRow(rowNumber);
    for (let column = 1; column <= worksheet.columnCount; column += 1) {
      const label = normalizeHeader(getCellText(row.getCell(column).value));
      if (label && !headers.has(label)) {
        headers.set(label, column);
      }
    }
  });

  return headers;
}

function buildSourceValueMap(row: ExcelJS.Row, headers: Map<string, number>) {
  const sourceValues = new Map<string, SourceValue>();
  headers.forEach((columnIndex, header) => {
    sourceValues.set(header, {
      value: row.getCell(columnIndex).value,
      columnIndex
    });
  });
  return sourceValues;
}

function getCellText(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value instanceof Date) {
    return formatDateDot(value);
  }
  if (typeof value === "object") {
    const richText = (value as { richText?: { text: string }[] }).richText;
    if (richText) {
      return richText.map((part) => part.text).join("");
    }

    const text = (value as { text?: string }).text;
    if (text) {
      return text;
    }

    const result = (value as { result?: unknown }).result;
    if (result !== undefined) {
      return getCellText(result);
    }
  }
  return "";
}

function toNumber(value: unknown): number | null {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value === "object" && !(value instanceof Date)) {
    const result = (value as { result?: unknown }).result;
    if (result !== undefined) {
      return toNumber(result);
    }
  }
  if (typeof value === "string") {
    const normalized = value.replace(/[₩,\s]/g, "").replace(/^\((.*)\)$/, "-$1");
    if (!normalized) {
      return null;
    }
    const number = Number(normalized);
    return Number.isFinite(number) ? number : null;
  }
  return null;
}

function sumItems(items: PayItem[]) {
  return items.reduce((sum, item) => sum + item.amount, 0);
}

function normalizeHeader(value: string): string {
  return value.trim().replace(/\s+/g, "");
}

function parseIsoDate(value: string, label: string): Date {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) {
    throw new Error(`${label}은 YYYY-MM-DD 형식이어야 합니다.`);
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    throw new Error(`${label}이 올바른 날짜가 아닙니다.`);
  }

  return date;
}

function formatDateDot(date: Date): string {
  return `${date.getFullYear()}.${pad2(date.getMonth() + 1)}.${pad2(date.getDate())}`;
}

function formatDateSlash(date: Date): string {
  return `${date.getFullYear()}/${pad2(date.getMonth() + 1)}/${pad2(date.getDate())}`;
}

function pad2(value: number): string {
  return String(value).padStart(2, "0");
}

function unique(values: string[]): string[] {
  return [...new Set(values)];
}

function createUniqueSheetName(name: string, usedSheetNames: Set<string>): string {
  const baseName = sanitizeSheetName(name) || "직원";
  let candidate = baseName.slice(0, 31);
  let suffix = 2;

  while (usedSheetNames.has(candidate)) {
    const suffixText = ` (${suffix})`;
    candidate = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`;
    suffix += 1;
  }

  usedSheetNames.add(candidate);
  return candidate;
}

function sanitizeSheetName(value: string): string {
  return value.replace(/[\\/?*[\]:]/g, "").trim();
}

function sanitizeFileName(value: string): string {
  return value.replace(/[\\/:*?"<>|]/g, "_").trim() || "회사";
}

function toArrayBuffer(buffer: ExcelJS.Buffer): ArrayBuffer {
  if (buffer instanceof ArrayBuffer) {
    return buffer;
  }

  const view = buffer as Uint8Array;
  const arrayBuffer = new ArrayBuffer(view.byteLength);
  new Uint8Array(arrayBuffer).set(view);
  return arrayBuffer;
}

function mergeAndSet(worksheet: ExcelJS.Worksheet, range: string, value: unknown) {
  worksheet.mergeCells(range);
  worksheet.getCell(range.split(":")[0]).value = value as ExcelJS.CellValue;
}

function setRowValues(
  worksheet: ExcelJS.Worksheet,
  rowNumber: number,
  values: Partial<Record<"B" | "C" | "D" | "E", ExcelJS.CellValue>>
) {
  Object.entries(values).forEach(([column, value]) => {
    worksheet.getCell(`${column}${rowNumber}`).value = value ?? null;
  });
}

function styleRange(
  worksheet: ExcelJS.Worksheet,
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
  style: CellStyle
) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let column = startColumn; column <= endColumn; column += 1) {
      styleCell(worksheet.getCell(row, column), style);
    }
  }
}

function styleCell(
  cell: ExcelJS.Cell,
  style: CellStyle
) {
  if (style.fill) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: style.fill }
    };
  }
  if (style.font) {
    const { color, ...font } = style.font;
    cell.font = {
      ...cell.font,
      ...font,
      ...(color ? { color: typeof color === "string" ? { argb: color } : color } : {})
    };
  }
  if (style.alignment) {
    cell.alignment = {
      ...cell.alignment,
      ...style.alignment
    };
  }
  if (style.border) {
    cell.border = CELL_BORDER;
  }
}
