import path from "node:path";
import ExcelJS, { type CellValue, type Workbook as ExcelWorkbook, type Worksheet } from "exceljs";
import { CHANGELOG_HEADERS, MASTER_HEADERS, PROJECT_INFO_HEADERS, RAW_HEADERS, SOURCE_HEADERS, WRITABLE_HEADERS } from "../constants/headers.js";
import { SHEET_NAMES } from "../constants/sheets.js";
import type { SheetOverview } from "../types/keywordProject.js";
import { withWorkbookWrite } from "./writeQueue.js";

/** ExcelJS 是 CommonJS 包，运行时必须从默认导入对象中取 Workbook。 */
const { Workbook } = ExcelJS;

/** 工作表可读性样式配置。 */
const SHEET_STYLE_CONFIG: Record<string, { tabColor: string; headerFill: string }> = {
  项目信息: { tabColor: "FF2F75B5", headerFill: "FF1F4E78" },
  数据来源: { tabColor: "FF70AD47", headerFill: "FF548235" },
  原始关键词: { tabColor: "FF5B9BD5", headerFill: "FF1F4E78" },
  关键词主表: { tabColor: "FFC55A11", headerFill: "FF9E480E" },
  更新记录: { tabColor: "FF7F7F7F", headerFill: "FF595959" },
};

/** 按字段名定义更适合阅读的列宽。 */
const COLUMN_WIDTHS: Record<string, number> = {
  字段: 24,
  值: 52,
  导入时间: 22,
  来源: 14,
  来源文件: 44,
  归档文件: 58,
  导入行数: 14,
  未映射字段: 28,
  关键词: 36,
  搜索量: 14,
  关键词难度: 14,
  "点击均价(CPC)": 16,
  搜索意图: 18,
  词性: 14,
  分类: 18,
  SERP功能: 48,
  过滤: 14,
  优先级: 12,
  备注: 46,
  筛选时间: 22,
  变更时间: 22,
  操作: 18,
  详情: 70,
};

/** 按字段名定义 Excel 数字格式。 */
const COLUMN_NUMBER_FORMATS: Record<string, string> = {
  搜索量: "#,##0",
  关键词难度: "0.00",
  "点击均价(CPC)": "$0.00",
  导入行数: "#,##0",
};

/** 需要自动换行的长文本字段。 */
const WRAPPED_COLUMNS = new Set(["来源文件", "归档文件", "关键词", "备注", "详情", "值"]);

/** 关键词主表词性列的固定可选值。 */
export const POS_VALUES = ["核心词", "长尾词", "问题词", "场景词"] as const;

/** 关键词主表来源列的固定可选值。 */
export const SOURCE_VALUES = ["导入", "手动"] as const;

/** 新建工作簿的项目元数据。 */
export interface WorkbookProjectInfo {
  /** 项目名称。 */
  name: string;
  /** 目标网站。 */
  site: string;
  /** 目标市场。 */
  market: string;
  /** 目标语言。 */
  language: string;
  /** 核心产品或服务。 */
  product: string;
}

/** 批量写入记录输入。 */
export interface WriteProjectRecordsInput {
  /** 目标人工维护工作表。 */
  sheet: keyof typeof WRITABLE_HEADERS;
  /** 要写入的记录。 */
  records: Array<Record<string, unknown>>;
  /** 写入模式。 */
  mode: "append" | "upsert";
  /** upsert 使用的关键列。 */
  key?: string;
}

/** 批量写入记录结果。 */
export interface WriteProjectRecordsResult {
  /** 追加的记录数。 */
  inserted: number;
  /** 更新的记录数。 */
  updated: number;
  /** 跳过的记录数。 */
  skipped: number;
  /** 输入中存在但工作表没有的字段。 */
  unknownFields: Array<{ recordIndex: number; fields: string[] }>;
  /** 每条输入记录的处理结果。 */
  rows: Array<Record<string, unknown>>;
}

/** 单条更新输入。 */
export interface UpdateRecordInput {
  /** 目标人工维护工作表。 */
  sheet: keyof typeof WRITABLE_HEADERS;
  /** 用于查找的关键列。 */
  key: string;
  /** 用于查找的关键值。 */
  value: string;
  /** 要覆盖的字段。 */
  patch: Record<string, unknown>;
}

/** 删除记录输入。 */
export interface DeleteRecordsInput {
  /** 目标人工维护工作表。 */
  sheet: keyof typeof WRITABLE_HEADERS;
  /** 用于查找的关键列。 */
  key: string;
  /** 要删除的关键值列表。 */
  values: string[];
}

/** 返回统一的本地时间字符串。 */
export function nowText(): string {
  return new Date().toISOString().replace("T", " ").slice(0, 19);
}

/** 把普通 JavaScript 值转换为 ExcelJS 可接受的单元格值。 */
export function toCellValue(value: unknown): CellValue {
  if (value == null) return null;
  if (value instanceof Date) return value;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  return String(value);
}

/** 给工作表第一行应用基础表头样式。 */
export function styleHeaderRow(sheet: Worksheet): void {
  const row = sheet.getRow(1);
  const style = SHEET_STYLE_CONFIG[sheet.name] ?? SHEET_STYLE_CONFIG["原始关键词"];
  row.height = 24;
  row.font = { bold: true, color: { argb: "FFFFFFFF" } };
  row.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: style.headerFill } };
  row.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: style.headerFill } };
    cell.border = {
      top: { style: "thin", color: { argb: "FFD9E2F3" } },
      left: { style: "thin", color: { argb: "FFD9E2F3" } },
      bottom: { style: "thin", color: { argb: "FFD9E2F3" } },
      right: { style: "thin", color: { argb: "FFD9E2F3" } },
    };
  });
  sheet.views = [{ state: "frozen", ySplit: 1 }];
}

/** 把列序号转换为 Excel A1 样式列名。 */
function columnLetter(columnIndex: number): string {
  let value = columnIndex;
  let letter = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    value = Math.floor((value - 1) / 26);
  }
  return letter;
}

/** 根据字段名返回当前列宽。 */
function columnWidth(header: string): number {
  return COLUMN_WIDTHS[header] ?? Math.max(14, Math.min(34, header.length + 8));
}

/** 为关键词主表固定选项列添加下拉验证。 */
function applyMasterDropdownValidations(sheet: Worksheet, headers: readonly string[]): void {
  if (sheet.name !== "关键词主表") return;

  /** 需要下拉验证的列及其可选值、提示信息。 */
  const dropdowns: Record<string, { values: readonly string[]; title: string; message: string }> = {
    词性: { values: POS_VALUES, title: "无效词性", message: "请从下拉列表中选择：核心词、长尾词、问题词、场景词" },
    来源: { values: SOURCE_VALUES, title: "无效来源", message: "请从下拉列表中选择：导入、手动" },
  };

  for (const [field, config] of Object.entries(dropdowns)) {
    const columnIndex = headers.indexOf(field) + 1;
    if (columnIndex <= 0) continue;
    const column = columnLetter(columnIndex);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (sheet as any).dataValidations.add(`${column}2:${column}1048576`, {
      type: "list",
      allowBlank: true,
      formulae: [`"${config.values.join(",")}"`],
      showErrorMessage: true,
      errorTitle: config.title,
      error: config.message,
    });
  }
}

/** 应用工作表级样式、筛选器和列格式。 */
function applyReadableSheetStyle(sheet: Worksheet, headers: readonly string[]): void {
  const style = SHEET_STYLE_CONFIG[sheet.name] ?? SHEET_STYLE_CONFIG["原始关键词"];
  sheet.properties.defaultRowHeight = 20;
  sheet.properties.tabColor = { argb: style.tabColor };
  sheet.autoFilter = `A1:${columnLetter(headers.length)}1`;
  headers.forEach((header, index) => {
    const column = sheet.getColumn(index + 1);
    column.width = columnWidth(header);
    column.alignment = {
      vertical: "top",
      horizontal: COLUMN_NUMBER_FORMATS[header] ? "right" : "left",
      wrapText: WRAPPED_COLUMNS.has(header),
    };
    if (COLUMN_NUMBER_FORMATS[header]) {
      column.numFmt = COLUMN_NUMBER_FORMATS[header];
    }
  });
}

/** 向工作表写入表头并应用基础样式。 */
export function writeHeaders(sheet: Worksheet, headers: readonly string[]): void {
  sheet.addRow([...headers]);
  applyReadableSheetStyle(sheet, headers);
  styleHeaderRow(sheet);
  applyMasterDropdownValidations(sheet, headers);
}

/** 读取工作表第一行表头。 */
export function readHeaders(sheet: Worksheet): string[] {
  const row = sheet.getRow(1);
  const headers: string[] = [];
  for (let index = 1; index <= sheet.columnCount; index += 1) {
    const value = row.getCell(index).value;
    headers.push(value == null ? "" : String(value));
  }
  return headers.filter((header) => header !== "");
}

/** 确保工作表存在并返回工作表对象。 */
export function requireWorksheet(workbook: ExcelWorkbook, sheetName: string): Worksheet {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw new Error(`工作表不存在：${sheetName}`);
  }
  return sheet;
}

/** 把项目元数据写入项目信息工作表。 */
export function writeProjectInfo(sheet: Worksheet, info: WorkbookProjectInfo): void {
  writeHeaders(sheet, PROJECT_INFO_HEADERS);
  const createdAt = nowText();
  sheet.addRows([
    ["项目名称", info.name],
    ["网站", info.site],
    ["目标市场", info.market],
    ["语言", info.language],
    ["核心产品", info.product],
    ["创建时间", createdAt],
    ["最后更新时间", createdAt],
  ]);
  sheet.getColumn(1).width = 24;
  sheet.getColumn(2).width = 48;
}

/** 更新项目信息工作表里的最后更新时间。 */
export function touchProjectInfo(workbook: ExcelWorkbook): void {
  const sheet = requireWorksheet(workbook, "项目信息");
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (String(sheet.getRow(rowNumber).getCell(1).value ?? "") === "最后更新时间") {
      sheet.getRow(rowNumber).getCell(2).value = nowText();
      return;
    }
  }
  sheet.addRow(["最后更新时间", nowText()]);
}

/** 追加一条更新记录。 */
export function appendChangelog(workbook: ExcelWorkbook, action: string, details: string): void {
  const sheet = workbook.getWorksheet("更新记录") ?? workbook.addWorksheet("更新记录");
  if (sheet.rowCount === 0) {
    writeHeaders(sheet, CHANGELOG_HEADERS);
  }
  sheet.addRow([nowText(), action, details]);
}

/** 创建标准关键词项目工作簿。 */
export async function createKeywordWorkbook(workbookPath: string, info: WorkbookProjectInfo): Promise<void> {
  const workbook = new Workbook();
  writeProjectInfo(workbook.addWorksheet(SHEET_NAMES[0]), info);
  writeHeaders(workbook.addWorksheet("数据来源"), SOURCE_HEADERS);
  writeHeaders(workbook.addWorksheet("原始关键词"), RAW_HEADERS);
  writeHeaders(workbook.addWorksheet("关键词主表"), MASTER_HEADERS);
  writeHeaders(workbook.addWorksheet("更新记录"), CHANGELOG_HEADERS);
  appendChangelog(workbook, "创建项目", `创建项目 ${info.name}`);
  await workbook.xlsx.writeFile(workbookPath);
}

/** 获取工作簿概览。 */
export async function getWorkbookOverview(workbookPath: string): Promise<{ workbook: string; projectDir: string; sheets: SheetOverview[] }> {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(workbookPath);
  return {
    workbook: path.resolve(workbookPath),
    projectDir: path.dirname(path.resolve(workbookPath)),
    sheets: workbook.worksheets.map((sheet) => ({
      name: sheet.name,
      maxRow: sheet.rowCount,
      maxColumn: sheet.columnCount,
      dataRows: Math.max(0, sheet.rowCount - 1),
      headers: readHeaders(sheet),
    })),
  };
}

/** 按工作表表头顺序追加多行字典数据，主要用于测试和内部导入。 */
export async function appendRows(workbookPath: string, sheetName: keyof typeof WRITABLE_HEADERS, records: Array<Record<string, unknown>>): Promise<void> {
  await withWorkbookWrite(workbookPath, async () => {
    const workbook = new Workbook();
    await workbook.xlsx.readFile(workbookPath);
    const sheet = requireWorksheet(workbook, sheetName);
    const headers = readHeaders(sheet);
    for (const record of records) {
      sheet.addRow(headers.map((header) => toCellValue(record[header])));
    }
    touchProjectInfo(workbook);
    appendChangelog(workbook, "测试写入", `向 ${sheetName} 写入 ${records.length} 行`);
    await workbook.xlsx.writeFile(workbookPath);
  });
}

/** 返回人工维护表默认 upsert 关键列。 */
function defaultWriteKey(sheetName: keyof typeof WRITABLE_HEADERS): string {
  return "关键词";
}

/** 按关键列查找第一条匹配行。 */
function findRowByKey(sheet: Worksheet, headers: string[], key: string, value: unknown): number | undefined {
  const columnIndex = headers.indexOf(key) + 1;
  if (columnIndex <= 0) return undefined;
  const expected = String(value ?? "").trim();
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (String(sheet.getRow(rowNumber).getCell(columnIndex).value ?? "").trim() === expected) {
      return rowNumber;
    }
  }
  return undefined;
}

/** 把输入记录拆成可写字段和未知字段。 */
function splitRecord(headers: string[], record: Record<string, unknown>): { known: Record<string, unknown>; unknown: string[] } {
  return {
    known: Object.fromEntries(Object.entries(record).filter(([key]) => headers.includes(key))),
    unknown: Object.keys(record).filter((key) => !headers.includes(key)),
  };
}

/** 把记录中的已知字段写入指定 Excel 行。 */
function updateSheetRow(sheet: Worksheet, headers: string[], rowNumber: number, values: Record<string, unknown>): void {
  for (const [field, value] of Object.entries(values)) {
    if (value != null && value !== "") {
      sheet.getRow(rowNumber).getCell(headers.indexOf(field) + 1).value = toCellValue(value);
    }
  }
}

/** 向人工维护表追加或 upsert 记录。 */
export async function writeProjectRecords(workbookPath: string, input: WriteProjectRecordsInput): Promise<WriteProjectRecordsResult> {
  return withWorkbookWrite(workbookPath, async () => {
    const workbook = new Workbook();
    await workbook.xlsx.readFile(workbookPath);
    const sheet = requireWorksheet(workbook, input.sheet);
    const headers = readHeaders(sheet);
    const key = input.key ?? defaultWriteKey(input.sheet);
    let inserted = 0;
    let updated = 0;
    let skipped = 0;
    const unknownFields: Array<{ recordIndex: number; fields: string[] }> = [];
    const rows: Array<Record<string, unknown>> = [];
    input.records.forEach((record, index) => {
      const enriched = { ...record };
      if (input.sheet === "关键词主表" && !("筛选时间" in enriched)) {
        enriched["筛选时间"] = nowText();
      }
      const { known, unknown } = splitRecord(headers, enriched);
      if (unknown.length > 0) unknownFields.push({ recordIndex: index + 1, fields: unknown });
      if (Object.keys(known).length === 0) {
        skipped += 1;
        rows.push({ recordIndex: index + 1, action: "skipped", reason: "没有可写入字段" });
        return;
      }
      if (input.mode === "upsert") {
        const keyValue = known[key];
        if (keyValue == null || keyValue === "") {
          skipped += 1;
          rows.push({ recordIndex: index + 1, action: "skipped", reason: `缺少关键列：${key}` });
          return;
        }
        const rowNumber = findRowByKey(sheet, headers, key, keyValue);
        if (rowNumber) {
          updateSheetRow(sheet, headers, rowNumber, known);
          updated += 1;
          rows.push({ recordIndex: index + 1, action: "updated", excelRow: rowNumber });
          return;
        }
      }
      sheet.addRow(headers.map((header) => toCellValue(known[header])));
      inserted += 1;
      rows.push({ recordIndex: index + 1, action: "inserted", excelRow: sheet.rowCount });
    });
    touchProjectInfo(workbook);
    appendChangelog(workbook, "写入项目", `向 ${input.sheet} 以 ${input.mode} 模式写入 ${inserted} 条，更新 ${updated} 条，跳过 ${skipped} 条`);
    await workbook.xlsx.writeFile(workbookPath);
    return { inserted, updated, skipped, unknownFields, rows };
  });
}

/** 按 key 更新单条人工维护表记录。 */
export async function updateRecord(workbookPath: string, input: UpdateRecordInput): Promise<{ updated: number; unknownFields: string[] }> {
  return withWorkbookWrite(workbookPath, async () => {
    const workbook = new Workbook();
    await workbook.xlsx.readFile(workbookPath);
    const sheet = requireWorksheet(workbook, input.sheet);
    const headers = readHeaders(sheet);
    const rowNumber = findRowByKey(sheet, headers, input.key, input.value);
    const { known, unknown } = splitRecord(headers, input.patch);
    if (!rowNumber) return { updated: 0, unknownFields: unknown };
    updateSheetRow(sheet, headers, rowNumber, known);
    touchProjectInfo(workbook);
    appendChangelog(workbook, "更新记录", `更新 ${input.sheet} 中 ${input.key}=${input.value}`);
    await workbook.xlsx.writeFile(workbookPath);
    return { updated: 1, unknownFields: unknown };
  });
}

/** 按 key 删除人工维护表记录。 */
export async function deleteRecords(workbookPath: string, input: DeleteRecordsInput): Promise<{ deleted: number; notFound: string[] }> {
  return withWorkbookWrite(workbookPath, async () => {
    const workbook = new Workbook();
    await workbook.xlsx.readFile(workbookPath);
    const sheet = requireWorksheet(workbook, input.sheet);
    const headers = readHeaders(sheet);
    const columnIndex = headers.indexOf(input.key) + 1;
    if (columnIndex <= 0) throw new Error(`关键列不存在：${input.key}`);
    const remaining = new Set(input.values.map((value) => value.trim()));
    let deleted = 0;
    for (let rowNumber = sheet.rowCount; rowNumber >= 2; rowNumber -= 1) {
      const value = String(sheet.getRow(rowNumber).getCell(columnIndex).value ?? "").trim();
      if (remaining.has(value)) {
        sheet.spliceRows(rowNumber, 1);
        remaining.delete(value);
        deleted += 1;
      }
    }
    touchProjectInfo(workbook);
    appendChangelog(workbook, "删除记录", `从 ${input.sheet} 删除 ${deleted} 条记录`);
    await workbook.xlsx.writeFile(workbookPath);
    return { deleted, notFound: [...remaining] };
  });
}
