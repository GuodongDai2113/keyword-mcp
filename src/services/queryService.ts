import ExcelJS from "exceljs";
import type { SortRule, WorkbookRow } from "../types/keywordProject.js";
import { readHeaders, requireWorksheet } from "./workbookService.js";

/** ExcelJS 是 CommonJS 包，运行时必须从默认导入对象中取 Workbook。 */
const { Workbook } = ExcelJS;

/** 工作表窗口读取参数。 */
export interface ReadSheetInput {
  /** 目标工作表名称。 */
  sheet: string;
  /** 从第几个数据行开始。 */
  start?: number;
  /** 返回行数限制。 */
  limit?: number;
  /** 需要返回的列。 */
  columns?: string[];
  /** 排序规则。 */
  sort?: SortRule[];
  /** 是否包含原始关键词中已过滤的行。 */
  includeFiltered?: boolean;
}

/** 工作表搜索参数。 */
export interface SearchSheetInput extends ReadSheetInput {
  /** 查询文本。 */
  query: string;
  /** 参与搜索的列。 */
  searchColumns?: string[];
}

/** 把单元格值转换为可排序值。 */
function sortableValue(value: unknown): string | number {
  if (typeof value === "number") return value;
  const text = String(value ?? "").trim();
  const number = Number(text.replace(/,/g, ""));
  return Number.isFinite(number) && text !== "" ? number : text.toLowerCase();
}

/** 按指定表头从 Excel 行中取值。 */
function rowToValues(headers: string[], rowValues: unknown[], selectedHeaders: string[]): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  for (const header of selectedHeaders) {
    const index = headers.indexOf(header);
    result[header] = index >= 0 ? rowValues[index + 1] : undefined;
  }
  return result;
}

/** 按多列排序规则稳定排序输出行。 */
function sortRows(rows: WorkbookRow[], sortRules: SortRule[]): WorkbookRow[] {
  return [...rows].sort((left, right) => {
    for (const rule of sortRules) {
      const leftValue = sortableValue(left.values[rule.column]);
      const rightValue = sortableValue(right.values[rule.column]);
      if (leftValue < rightValue) return rule.direction === "desc" ? 1 : -1;
      if (leftValue > rightValue) return rule.direction === "desc" ? -1 : 1;
    }
    return left.excelRow - right.excelRow;
  });
}

/** 根据当前表头返回实际存在的列名。 */
function selectHeaders(headers: string[], requested?: string[]): string[] {
  return requested?.filter((column) => headers.includes(column)) ?? headers;
}

/** 读取指定工作表所有非空数据行。 */
async function readAllRows(workbookPath: string, sheetName: string): Promise<{ headers: string[]; rows: WorkbookRow[] }> {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(workbookPath);
  const sheet = requireWorksheet(workbook, sheetName);
  const headers = readHeaders(sheet);
  const rows: WorkbookRow[] = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (!row.hasValues) continue;
    const rowValues = row.values as unknown[];
    const values = rowToValues(headers, rowValues, headers);
    if (Object.values(values).some((value) => value != null && value !== "")) {
      rows.push({ excelRow: rowNumber, values });
    }
  }
  return { headers, rows };
}

/** 判断原始关键词行是否应在默认查询中跳过。 */
function shouldSkipFilteredRawRow(sheetName: string, row: WorkbookRow, includeFiltered?: boolean): boolean {
  return sheetName === "原始关键词" && !includeFiltered && String(row.values["过滤"] ?? "否") === "已过滤";
}

/** 读取指定工作表的一段数据。 */
export async function readSheetWindow(workbookPath: string, input: ReadSheetInput): Promise<{ workbook: string; sheet: string; headers: string[]; selectedHeaders: string[]; start: number; limit: number; rows: WorkbookRow[]; returnedRows: number }> {
  const { headers, rows } = await readAllRows(workbookPath, input.sheet);
  const selectedHeaders = selectHeaders(headers, input.columns);
  const sortRules = input.sort?.filter((rule) => headers.includes(rule.column)) ?? [];
  const start = Math.max(1, input.start ?? 1);
  const limit = Math.max(0, input.limit ?? 20);
  const visibleRows = rows.filter((row) => !shouldSkipFilteredRawRow(input.sheet, row, input.includeFiltered));
  const orderedRows = sortRules.length > 0 ? sortRows(visibleRows, sortRules) : visibleRows;
  const slicedRows = limit === 0 ? [] : orderedRows.slice(start - 1, start - 1 + limit);
  const outputRows = slicedRows.map((row) => ({
    excelRow: row.excelRow,
    values: Object.fromEntries(selectedHeaders.map((header) => [header, row.values[header]])),
  }));
  return {
    workbook: workbookPath,
    sheet: input.sheet,
    headers,
    selectedHeaders,
    start,
    limit,
    rows: outputRows,
    returnedRows: outputRows.length,
  };
}

/** 搜索指定工作表并返回匹配行。 */
export async function searchSheetRows(workbookPath: string, input: SearchSheetInput): Promise<{ workbook: string; sheet: string; query: string; headers: string[]; selectedHeaders: string[]; searchedHeaders: string[]; rows: WorkbookRow[]; matchedRows: number }> {
  const { headers, rows } = await readAllRows(workbookPath, input.sheet);
  const selectedHeaders = selectHeaders(headers, input.columns);
  const searchedHeaders = selectHeaders(headers, input.searchColumns);
  const query = input.query.toLowerCase();
  const visibleRows = rows.filter((row) => !shouldSkipFilteredRawRow(input.sheet, row, input.includeFiltered));
  const matched = visibleRows.filter((row) => searchedHeaders.some((header) => String(row.values[header] ?? "").toLowerCase().includes(query)));
  const sortRules = input.sort?.filter((rule) => headers.includes(rule.column)) ?? [];
  const orderedRows = sortRules.length > 0 ? sortRows(matched, sortRules) : matched;
  const limit = Math.max(0, input.limit ?? 20);
  const outputRows = orderedRows.slice(0, limit).map((row) => ({
    excelRow: row.excelRow,
    values: Object.fromEntries(selectedHeaders.map((header) => [header, row.values[header]])),
  }));
  return {
    workbook: workbookPath,
    sheet: input.sheet,
    query: input.query,
    headers,
    selectedHeaders,
    searchedHeaders,
    rows: outputRows,
    matchedRows: outputRows.length,
  };
}
