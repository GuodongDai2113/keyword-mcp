import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS, { type Workbook as ExcelWorkbook } from "exceljs";
import { RAW_HEADERS, SOURCE_HEADERS } from "../constants/headers.js";
import { IGNORED_SOURCE_COLUMNS, SEMRUSH_COLUMN_ALIASES } from "../constants/sourceMappings.js";
import { appendChangelog, nowText, readHeaders, requireWorksheet, toCellValue, touchProjectInfo } from "./workbookService.js";
import { withWorkbookWrite } from "./writeQueue.js";

/** ExcelJS 是 CommonJS 包，运行时必须从默认导入对象中取 Workbook。 */
const { Workbook } = ExcelJS;

/** 来源导入输入。 */
export interface ImportKeywordSourcesInput {
  /** 来源名称，初版只支持 semrush。 */
  source: "semrush";
  /** 要导入的文件；未传入时扫描 data-sources。 */
  filePaths?: string[];
  /** 是否跳过数据来源表中已经记录过的同名文件。 */
  skipExisting: boolean;
}

/** 来源导入结果。 */
export interface ImportKeywordSourcesResult {
  /** 成功导入的文件。 */
  importedFiles: Array<{ file: string; path: string; importedRows: number; unmappedColumns: string[]; storedFile: string }>;
  /** 被跳过的文件。 */
  skippedFiles: Array<{ file: string; path: string; importedRows: number; unmappedColumns: string[]; reason: string }>;
  /** 本次合计导入行数。 */
  totalImportedRows: number;
}

/** 把来源字段名标准化为匹配 key。 */
function normalizeColumn(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "");
}

/** 把常见数字文本转换为数值。 */
function toNumber(value: string | undefined): number | undefined {
  if (value == null || value.trim() === "") return undefined;
  const parsed = Number(value.replace(/,/g, "").replace(/%$/, ""));
  return Number.isFinite(parsed) ? Number(parsed.toFixed(4)) : undefined;
}

/** 把 SEMrush 搜索意图转换为中文。 */
function normalizeIntent(value: string | undefined): string {
  const mapping: Record<string, string> = {
    i: "信息型",
    informational: "信息型",
    n: "导航型",
    navigational: "导航型",
    c: "商业型",
    commercial: "商业型",
    t: "交易型",
    transactional: "交易型",
  };
  const labels: string[] = [];
  for (const part of String(value ?? "").split(/[,;/|\s]+/)) {
    const label = mapping[part.trim().toLowerCase()];
    if (label && !labels.includes(label)) labels.push(label);
  }
  return labels.join("、");
}

/** 把 SEMrush SERP 功能转换为中文展示文本。 */
function normalizeSerpFeatures(value: string | undefined): string {
  const mapping: Record<string, string> = {
    "featured snippet": "精选摘要",
    "people also ask": "其他用户还会问",
    paa: "其他用户还会问",
    sitelinks: "站点链接",
    "image pack": "图片包",
    images: "图片结果",
    video: "视频结果",
    videos: "视频结果",
    "local pack": "本地包",
    reviews: "评价",
    faq: "常见问题",
  };
  const labels: string[] = [];
  for (const part of String(value ?? "").split(/[,;/|]+/)) {
    const text = part.trim();
    if (!text) continue;
    const label = mapping[text.toLowerCase()] ?? text;
    if (!labels.includes(label)) labels.push(label);
  }
  return labels.join(", ");
}

/** 解析一行简单 CSV/TSV，支持双引号包裹字段。 */
function parseDelimitedLine(line: string, delimiter: "," | "\t"): string[] {
  const values: string[] = [];
  let current = "";
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const next = line[index + 1];
    if (char === "\"" && quoted && next === "\"") {
      current += "\"";
      index += 1;
    } else if (char === "\"") {
      quoted = !quoted;
    } else if (char === delimiter && !quoted) {
      values.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }
  values.push(current.trim());
  return values;
}

/** 解析简单 CSV 或 TSV 文本为对象数组。 */
function parseDelimited(text: string, delimiter: "," | "\t"): Array<Record<string, string>> {
  const lines = text.split(/\r?\n/).filter((line) => line.trim() !== "");
  const headers = lines[0] ? parseDelimitedLine(lines[0], delimiter) : [];
  return lines.slice(1).map((line) => {
    const values = parseDelimitedLine(line, delimiter);
    return Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""]));
  });
}

/** 读取来源表格文件为对象数组。 */
async function readSourceRows(filePath: string): Promise<Array<Record<string, string>>> {
  const suffix = path.extname(filePath).toLowerCase();
  if (suffix === ".csv" || suffix === ".tsv") {
    const text = await fs.readFile(filePath, "utf8");
    return parseDelimited(text, suffix === ".tsv" ? "\t" : ",");
  }
  const workbook = new Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.worksheets[0];
  if (!sheet) return [];
  const headers = readHeaders(sheet);
  const rows: Array<Record<string, string>> = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    rows.push(Object.fromEntries(headers.map((header, index) => [header, String(row.getCell(index + 1).value ?? "")])));
  }
  return rows;
}

/** 建立 SEMrush 来源列到内部字段的映射。 */
function mapColumns(columns: string[]): { mapped: Record<string, string>; unmapped: string[] } {
  const aliases = Object.fromEntries(
    Object.entries(SEMRUSH_COLUMN_ALIASES).flatMap(([standard, sourceAliases]) => sourceAliases.map((alias) => [normalizeColumn(alias), standard])),
  );
  const mapped: Record<string, string> = {};
  const unmapped: string[] = [];
  for (const column of columns) {
    const normalized = normalizeColumn(column);
    const standard = aliases[normalized];
    if (standard) mapped[column] = standard;
    else if (!IGNORED_SOURCE_COLUMNS.has(normalized)) unmapped.push(column);
  }
  return { mapped, unmapped };
}

/** 判断数据来源表是否已经记录过同名来源文件。 */
function sourceExists(workbook: ExcelWorkbook, fileName: string): boolean {
  const sheet = workbook.getWorksheet("数据来源");
  if (!sheet) return false;
  const headers = readHeaders(sheet);
  const index = headers.indexOf("来源文件") + 1;
  if (index <= 0) return false;
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (String(sheet.getRow(rowNumber).getCell(index).value ?? "") === fileName) return true;
  }
  return false;
}

/** 把标准字段转换为原始关键词表的中文行。 */
function buildRawKeywordRow(importedAt: string, fileName: string, standard: Record<string, string | undefined>): Record<string, unknown> {
  return {
    导入时间: importedAt,
    来源: "SEMrush",
    来源文件: fileName,
    关键词: String(standard.keyword ?? "").trim(),
    搜索量: toNumber(standard.searchVolume),
    关键词难度: toNumber(standard.keywordDifficulty),
    "点击均价(CPC)": toNumber(standard.cpc),
    搜索意图: normalizeIntent(standard.intent),
    过滤: "否",
  };
}

/** 导入关键词来源文件到原始关键词表。 */
export async function importKeywordSources(workbookPath: string, input: ImportKeywordSourcesInput): Promise<ImportKeywordSourcesResult> {
  return withWorkbookWrite(workbookPath, async () => {
    if (!input.filePaths || input.filePaths.length === 0) {
      throw new Error("未提供要导入的文件路径");
    }
    const filePaths = input.filePaths;
    const importedFiles: ImportKeywordSourcesResult["importedFiles"] = [];
    const skippedFiles: ImportKeywordSourcesResult["skippedFiles"] = [];
    const workbook = new Workbook();
    await workbook.xlsx.readFile(workbookPath);
    const rawSheet = requireWorksheet(workbook, "原始关键词");
    const sourceSheet = requireWorksheet(workbook, "数据来源");
    for (const filePath of filePaths) {
      const fileName = path.basename(filePath);
      if (input.skipExisting && sourceExists(workbook, fileName)) {
        skippedFiles.push({ file: fileName, path: filePath, importedRows: 0, unmappedColumns: [], reason: "导入记录已存在" });
        continue;
      }
      const rows = await readSourceRows(filePath);
      const mapping = mapColumns(Object.keys(rows[0] ?? {}));
      if (!Object.values(mapping.mapped).includes("keyword")) {
        throw new Error("无法识别关键词列");
      }
      const importedAt = nowText();
      let importedRows = 0;
      for (const row of rows) {
        const standard = Object.fromEntries(Object.entries(mapping.mapped).map(([sourceColumn, standardColumn]) => [standardColumn, row[sourceColumn]]));
        const rawRow = buildRawKeywordRow(importedAt, fileName, standard);
        if (!rawRow["关键词"]) continue;
        rawSheet.addRow(RAW_HEADERS.map((header) => toCellValue(rawRow[header])));
        importedRows += 1;
      }
      const storedFile = path.resolve(filePath);
      const sourceRow: Record<string, unknown> = {
        导入时间: importedAt,
        来源: "SEMrush",
        来源文件: fileName,
        归档文件: storedFile,
        导入行数: importedRows,
        未映射字段: mapping.unmapped.join(", "),
        备注: mapping.unmapped.length > 0 ? "存在未映射字段" : "导入完成",
      };
      sourceSheet.addRow(SOURCE_HEADERS.map((header) => toCellValue(sourceRow[header])));
      importedFiles.push({ file: fileName, path: filePath, importedRows, unmappedColumns: mapping.unmapped, storedFile });
    }
    const totalImportedRows = importedFiles.reduce((sum, item) => sum + item.importedRows, 0);
    touchProjectInfo(workbook);
    appendChangelog(workbook, "导入数据源", `导入 ${importedFiles.length} 个文件，合计 ${totalImportedRows} 行`);
    await workbook.xlsx.writeFile(workbookPath);
    return { importedFiles, skippedFiles, totalImportedRows };
  });
}
