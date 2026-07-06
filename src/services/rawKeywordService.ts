import type { Worksheet } from "exceljs";
import { MASTER_HEADERS, RAW_HEADERS } from "../constants/headers.js";
import { appendChangelog, nowText, readHeaders, requireWorksheet, toCellValue, touchProjectInfo } from "./workbookService.js";
import { withWorkbookWrite } from "./writeQueue.js";

/** ExcelJS 是 CommonJS 包，运行时必须从默认导入对象中取 Workbook。 */
import ExcelJS from "exceljs";

/** ExcelJS 工作簿构造器。 */
const { Workbook: ExcelWorkbook } = ExcelJS;

/** 原始关键词过滤值。 */
export type RawKeywordFilterValue = "否" | "已过滤";

/** 原始关键词定位输入。 */
export interface RawKeywordSelector {
  /** 原始关键词所在 Excel 行号。 */
  rawExcelRows?: number[];
  /** 原始关键词文本。 */
  keywords?: string[];
}

/** 单条原始关键词过滤项。 */
export interface FilterRawKeywordItem {
  /** 原始关键词所在 Excel 行号。 */
  rawExcelRow?: number;
  /** 原始关键词文本。 */
  keyword?: string;
  /** 要写入过滤列的值。 */
  filterValue?: RawKeywordFilterValue;
}

/** 单条原始关键词转移项。 */
export interface TransferRawKeywordItem {
  /** 原始关键词所在 Excel 行号。 */
  rawExcelRow?: number;
  /** 原始关键词文本。 */
  keyword?: string;
  /** 转移到关键词主表时设置的优先级。 */
  priority: "高" | "中" | "低";
  /** 转移到关键词主表时设置的备注。 */
  note: string;
}

/** 设置原始关键词过滤状态的输入。 */
export interface FilterRawKeywordsInput extends RawKeywordSelector {
  /** 要写入过滤列的值。 */
  filterValue?: RawKeywordFilterValue;
  /** 要批量处理的过滤项，每项可独立指定过滤值。 */
  items?: FilterRawKeywordItem[];
}

/** 转移原始关键词到主表的输入。 */
export interface TransferRawKeywordsInput extends RawKeywordSelector {
  /** 转移到关键词主表时设置的优先级，未使用 items 时必填。 */
  priority?: "高" | "中" | "低";
  /** 转移到关键词主表时设置的备注，未使用 items 时必填。 */
  note?: string;
  /** 写入关键词主表的模式。 */
  mode?: "append" | "upsert";
  /** upsert 使用的关键列。 */
  key?: string;
  /** 要批量转移的关键词项，每项可独立设置优先级和备注。 */
  items?: TransferRawKeywordItem[];
}

/** 单条筛选决策。 */
export type RawKeywordScreenDecision =
  | ({ action: "filter" } & FilterRawKeywordItem)
  | ({ action: "transfer" } & TransferRawKeywordItem);

/** 混合筛选决策输入。 */
export interface ScreenRawKeywordsInput {
  /** 一次提交的筛选决策。 */
  decisions: RawKeywordScreenDecision[];
  /** 写入关键词主表的模式。 */
  mode?: "append" | "upsert";
  /** upsert 使用的关键列。 */
  key?: string;
}

/** 原始关键词筛选结果。 */
export interface RawKeywordMutationResult {
  /** 匹配到并处理的行数。 */
  updated: number;
  /** 未匹配到的关键词或行号。 */
  notFound: Array<string | number>;
  /** 被处理的 Excel 行号。 */
  rows: number[];
}

/** 原始关键词转移结果。 */
export interface TransferRawKeywordsResult {
  /** 转移到关键词主表的记录数。 */
  transferred: number;
  /** 更新的关键词主表记录数。 */
  updated: number;
  /** 追加的关键词主表记录数。 */
  inserted: number;
  /** 未匹配到的关键词或行号。 */
  notFound: Array<string | number>;
  /** 被标记为已过滤的原始关键词 Excel 行号。 */
  filteredRows: number[];
}

/** 混合筛选决策结果。 */
export interface ScreenRawKeywordsResult {
  /** 标记为过滤的原始关键词数量。 */
  filtered: number;
  /** 转移到关键词主表的记录数。 */
  transferred: number;
  /** 更新的关键词主表记录数。 */
  updated: number;
  /** 追加的关键词主表记录数。 */
  inserted: number;
  /** 未匹配到的关键词或行号。 */
  notFound: Array<string | number>;
  /** 被处理的原始关键词 Excel 行号。 */
  rows: number[];
}

/** 把 ExcelJS 运行时工作簿类型统一为实例类型。 */
type RuntimeWorkbook = InstanceType<typeof ExcelWorkbook>;

/** 在原始关键词表中查找目标行。 */
function selectRawRows(sheet: Worksheet, headers: string[], selector: RawKeywordSelector): { rowNumbers: number[]; notFound: Array<string | number> } {
  const rowNumbers = new Set<number>();
  const notFound: Array<string | number> = [];
  const keywordColumn = headers.indexOf("关键词") + 1;
  for (const rowNumber of selector.rawExcelRows ?? []) {
    if (rowNumber >= 2 && rowNumber <= sheet.rowCount) rowNumbers.add(rowNumber);
    else notFound.push(rowNumber);
  }
  for (const keyword of selector.keywords ?? []) {
    let matched = false;
    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
      const value = String(sheet.getRow(rowNumber).getCell(keywordColumn).value ?? "").trim();
      if (value === keyword.trim()) {
        rowNumbers.add(rowNumber);
        matched = true;
      }
    }
    if (!matched) notFound.push(keyword);
  }
  return { rowNumbers: [...rowNumbers].sort((left, right) => left - right), notFound };
}

/** 把单条定位项转换为通用选择器。 */
function selectorFromItem(item: { rawExcelRow?: number; keyword?: string }): RawKeywordSelector {
  return {
    rawExcelRows: item.rawExcelRow ? [item.rawExcelRow] : undefined,
    keywords: item.keyword ? [item.keyword] : undefined,
  };
}

/** 把旧版过滤入参展开为逐项过滤列表。 */
function expandFilterItems(input: FilterRawKeywordsInput): FilterRawKeywordItem[] {
  if (input.items?.length) return input.items;
  return [
    ...(input.rawExcelRows ?? []).map((rawExcelRow) => ({ rawExcelRow, filterValue: input.filterValue })),
    ...(input.keywords ?? []).map((keyword) => ({ keyword, filterValue: input.filterValue })),
  ];
}

/** 把旧版转移入参展开为逐项转移列表。 */
function expandTransferItems(input: TransferRawKeywordsInput): TransferRawKeywordItem[] {
  if (input.items?.length) return input.items;
  if (!input.priority || !input.note) {
    throw new Error("转移原始关键词时必须提供 priority 和 note，或使用 items 为每条记录分别提供。");
  }
  return [
    ...(input.rawExcelRows ?? []).map((rawExcelRow) => ({ rawExcelRow, priority: input.priority!, note: input.note! })),
    ...(input.keywords ?? []).map((keyword) => ({ keyword, priority: input.priority!, note: input.note! })),
  ];
}

/** 根据表头读取指定原始关键词行。 */
function rawRowToRecord(sheet: Worksheet, headers: string[], rowNumber: number): Record<string, unknown> {
  const row = sheet.getRow(rowNumber);
  return Object.fromEntries(headers.map((header, index) => [header, row.getCell(index + 1).value]));
}

/** 按关键列查找关键词主表行。 */
function findMasterRow(sheet: Worksheet, headers: string[], key: string, value: unknown): number | undefined {
  const columnIndex = headers.indexOf(key) + 1;
  if (columnIndex <= 0) return undefined;
  const expected = String(value ?? "").trim();
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const actual = String(sheet.getRow(rowNumber).getCell(columnIndex).value ?? "").trim();
    if (actual === expected) return rowNumber;
  }
  return undefined;
}

/** 把主表记录写入指定行。 */
function writeMasterRecord(sheet: Worksheet, headers: string[], rowNumber: number, record: Record<string, unknown>): void {
  for (const [field, value] of Object.entries(record)) {
    const columnIndex = headers.indexOf(field) + 1;
    if (columnIndex > 0) sheet.getRow(rowNumber).getCell(columnIndex).value = toCellValue(value);
  }
}

/** 把原始关键词记录转换为关键词主表记录。 */
function buildMasterRecord(rawRecord: Record<string, unknown>, input: Pick<TransferRawKeywordItem, "priority" | "note">): Record<string, unknown> {
  return {
    关键词: rawRecord["关键词"],
    搜索量: rawRecord["搜索量"],
    关键词难度: rawRecord["关键词难度"],
    "点击均价(CPC)": rawRecord["点击均价(CPC)"],
    搜索意图: rawRecord["搜索意图"],
    优先级: input.priority,
    备注: input.note,
    筛选时间: nowText(),
  };
}

/** 在已打开工作簿中批量设置原始关键词过滤列。 */
function filterRawKeywordItems(rawSheet: Worksheet, headers: string[], items: FilterRawKeywordItem[]): RawKeywordMutationResult {
  const filterColumn = headers.indexOf("过滤") + 1;
  if (filterColumn <= 0) throw new Error("原始关键词缺少过滤列。");
  const notFound: Array<string | number> = [];
  const rows: number[] = [];
  for (const item of items) {
    const selected = selectRawRows(rawSheet, headers, selectorFromItem(item));
    notFound.push(...selected.notFound);
    for (const rowNumber of selected.rowNumbers) {
      rawSheet.getRow(rowNumber).getCell(filterColumn).value = item.filterValue ?? "已过滤";
      rows.push(rowNumber);
    }
  }
  return { updated: rows.length, notFound, rows };
}

/** 在已打开工作簿中批量转移原始关键词到关键词主表。 */
function transferRawKeywordItems(rawSheet: Worksheet, masterSheet: Worksheet, rawHeaders: string[], masterHeaders: string[], items: TransferRawKeywordItem[], mode: "append" | "upsert", key: string): TransferRawKeywordsResult {
  const filterColumn = rawHeaders.indexOf("过滤") + 1;
  if (filterColumn <= 0) throw new Error("原始关键词缺少过滤列。");
  const notFound: Array<string | number> = [];
  const filteredRows: number[] = [];
  let inserted = 0;
  let updated = 0;
  for (const item of items) {
    const selected = selectRawRows(rawSheet, rawHeaders, selectorFromItem(item));
    notFound.push(...selected.notFound);
    for (const rowNumber of selected.rowNumbers) {
      const rawRecord = rawRowToRecord(rawSheet, rawHeaders, rowNumber);
      const masterRecord = buildMasterRecord(rawRecord, item);
      const targetRow = mode === "upsert" ? findMasterRow(masterSheet, masterHeaders, key, masterRecord[key]) : undefined;
      if (targetRow) {
        writeMasterRecord(masterSheet, masterHeaders, targetRow, masterRecord);
        updated += 1;
      } else {
        masterSheet.addRow(MASTER_HEADERS.map((header) => toCellValue(masterRecord[header])));
        inserted += 1;
      }
      rawSheet.getRow(rowNumber).getCell(filterColumn).value = "已过滤";
      filteredRows.push(rowNumber);
    }
  }
  return { transferred: filteredRows.length, updated, inserted, notFound, filteredRows };
}

/** 设置原始关键词过滤列。 */
export async function filterRawKeywords(workbookPath: string, input: FilterRawKeywordsInput): Promise<RawKeywordMutationResult> {
  return withWorkbookWrite(workbookPath, async () => {
    const workbook = new ExcelWorkbook() as RuntimeWorkbook;
    await workbook.xlsx.readFile(workbookPath);
    const rawSheet = requireWorksheet(workbook, "原始关键词");
    const headers = readHeaders(rawSheet);
    const result = filterRawKeywordItems(rawSheet, headers, expandFilterItems(input));
    touchProjectInfo(workbook);
    appendChangelog(workbook, "过滤原始关键词", `设置 ${result.updated} 条原始关键词过滤状态`);
    await workbook.xlsx.writeFile(workbookPath);
    return result;
  });
}

/** 从原始关键词转移记录到关键词主表，并把原始关键词标记为已过滤。 */
export async function transferRawKeywordsToMaster(workbookPath: string, input: TransferRawKeywordsInput): Promise<TransferRawKeywordsResult> {
  return withWorkbookWrite(workbookPath, async () => {
    const workbook = new ExcelWorkbook() as RuntimeWorkbook;
    await workbook.xlsx.readFile(workbookPath);
    const rawSheet = requireWorksheet(workbook, "原始关键词");
    const masterSheet = requireWorksheet(workbook, "关键词主表");
    const rawHeaders = readHeaders(rawSheet);
    const masterHeaders = readHeaders(masterSheet);
    const mode = input.mode ?? "upsert";
    const key = input.key ?? "关键词";
    const result = transferRawKeywordItems(rawSheet, masterSheet, rawHeaders, masterHeaders, expandTransferItems(input), mode, key);
    touchProjectInfo(workbook);
    appendChangelog(workbook, "转移原始关键词", `转移 ${result.transferred} 条原始关键词到关键词主表，并标记为已过滤`);
    await workbook.xlsx.writeFile(workbookPath);
    return result;
  });
}

/** 一次提交混合筛选决策，支持过滤和转移。 */
export async function screenRawKeywords(workbookPath: string, input: ScreenRawKeywordsInput): Promise<ScreenRawKeywordsResult> {
  return withWorkbookWrite(workbookPath, async () => {
    const workbook = new ExcelWorkbook() as RuntimeWorkbook;
    await workbook.xlsx.readFile(workbookPath);
    const rawSheet = requireWorksheet(workbook, "原始关键词");
    const masterSheet = requireWorksheet(workbook, "关键词主表");
    const rawHeaders = readHeaders(rawSheet);
    const masterHeaders = readHeaders(masterSheet);
    const filterItems = input.decisions.filter((decision): decision is Extract<RawKeywordScreenDecision, { action: "filter" }> => decision.action === "filter");
    const transferItems = input.decisions.filter((decision): decision is Extract<RawKeywordScreenDecision, { action: "transfer" }> => decision.action === "transfer");
    const filterResult = filterRawKeywordItems(rawSheet, rawHeaders, filterItems);
    const transferResult = transferRawKeywordItems(rawSheet, masterSheet, rawHeaders, masterHeaders, transferItems, input.mode ?? "upsert", input.key ?? "关键词");
    touchProjectInfo(workbook);
    appendChangelog(workbook, "筛选原始关键词", `过滤 ${filterResult.updated} 条，转移 ${transferResult.transferred} 条`);
    await workbook.xlsx.writeFile(workbookPath);
    return {
      filtered: filterResult.updated,
      transferred: transferResult.transferred,
      updated: transferResult.updated,
      inserted: transferResult.inserted,
      notFound: [...filterResult.notFound, ...transferResult.notFound],
      rows: [...filterResult.rows, ...transferResult.filteredRows],
    };
  });
}
