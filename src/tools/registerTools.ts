import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { createServerConfig } from "../config.js";
import { SUPPORTED_SOURCE_SUFFIXES } from "../constants/sourceMappings.js";
import { createProject, resolveWorkbookPath } from "../services/projectService.js";
import { deleteRecords, getWorkbookOverview, readHeaders, requireWorksheet, updateRecord, writeProjectRecords } from "../services/workbookService.js";
import { importKeywordSources } from "../services/importService.js";
import { readSheetWindow, searchSheetRows } from "../services/queryService.js";
import { filterRawKeywords, screenRawKeywords, transferRawKeywordsToMaster } from "../services/rawKeywordService.js";
import { createProjectSchema, deleteRecordsSchema, filterRawKeywordsSchema, importSourceSchema, readSheetSchema, screenRawKeywordsSchema, searchSheetSchema, transferRawKeywordsSchema, updateRecordSchema, workbookLocatorSchema, writeRecordsSchema } from "../schemas/tools.js";

/** ExcelJS 是 CommonJS 包，运行时必须从默认导入对象中取 Workbook。 */
const { Workbook } = ExcelJS;

/** MCP 暴露的关键词项目 tool 名称。 */
export const registeredToolNames = [
  "keyword_project_create",
  "keyword_project_overview",
  "keyword_project_import_source",
  "keyword_project_read_sheet",
  "keyword_project_search_sheet",
  "keyword_project_write_records",
  "keyword_project_update_record",
  "keyword_project_delete_records",
  "keyword_project_filter_raw_keywords",
  "keyword_project_transfer_raw_keywords",
  "keyword_project_screen_raw_keywords",
  "keyword_project_list_sources",
] as const;

/** 把任意结果包装为 MCP 文本响应。 */
function jsonResponse(data: unknown) {
  return { content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }] };
}

/** 根据通用定位输入解析工作簿路径。 */
function workbookFromLocator(input: { projectPath?: string; workbookPath?: string }): string {
  const config = createServerConfig();
  return resolveWorkbookPath({ ...input, projectsRoot: config.projectsRoot, workbookName: config.workbookName });
}

/** 列出项目数据源文件和导入记录。 */
async function listSources(workbookPath: string): Promise<{ workbook: string; sourceFiles: Array<Record<string, unknown>>; importRecords: Array<Record<string, unknown>> }> {
  const sourceDir = path.join(path.dirname(workbookPath), "data-sources");
  const workbook = new Workbook();
  await workbook.xlsx.readFile(workbookPath);
  const sourceSheet = requireWorksheet(workbook, "数据来源");
  const headers = readHeaders(sourceSheet);
  const importRecords: Array<Record<string, unknown>> = [];
  for (let rowNumber = 2; rowNumber <= sourceSheet.rowCount; rowNumber += 1) {
    const row = sourceSheet.getRow(rowNumber);
    if (!row.hasValues) continue;
    importRecords.push(Object.fromEntries(headers.map((header, index) => [header, row.getCell(index + 1).value])));
  }
  const importedNames = new Set(importRecords.map((record) => String(record["来源文件"] ?? "")));
  const entries = await fs.readdir(sourceDir, { withFileTypes: true }).catch(() => []);
  const sourceFiles = await Promise.all(
    entries
      .filter((entry) => entry.isFile() && SUPPORTED_SOURCE_SUFFIXES.has(path.extname(entry.name).toLowerCase()) && !entry.name.startsWith("~$"))
      .map(async (entry) => {
        const filePath = path.join(sourceDir, entry.name);
        const stat = await fs.stat(filePath);
        return {
          name: entry.name,
          path: filePath,
          suffix: path.extname(entry.name).toLowerCase(),
          size: stat.size,
          imported: importedNames.has(entry.name),
        };
      }),
  );
  return { workbook: workbookPath, sourceFiles, importRecords };
}

/** 注册所有关键词项目 MCP tools。 */
export function registerKeywordTools(server: McpServer): void {
  server.tool("keyword_project_create", createProjectSchema.shape, async (input) => {
    const parsed = createProjectSchema.parse(input);
    const config = createServerConfig();
    return jsonResponse(
      await createProject({
        root: parsed.root ?? config.projectsRoot,
        name: parsed.name,
        site: parsed.site,
        market: parsed.market ?? config.defaultMarket,
        language: parsed.language ?? config.defaultLanguage,
        product: parsed.product,
      }),
    );
  });

  server.tool("keyword_project_overview", workbookLocatorSchema.shape, async (input) => {
    return jsonResponse(await getWorkbookOverview(workbookFromLocator(workbookLocatorSchema.parse(input))));
  });

  server.tool("keyword_project_import_source", importSourceSchema.shape, async (input) => {
    const parsed = importSourceSchema.parse(input);
    return jsonResponse(await importKeywordSources(workbookFromLocator(parsed), parsed));
  });

  server.tool("keyword_project_read_sheet", readSheetSchema.shape, async (input) => {
    const parsed = readSheetSchema.parse(input);
    return jsonResponse(await readSheetWindow(workbookFromLocator(parsed), parsed));
  });

  server.tool("keyword_project_search_sheet", searchSheetSchema.shape, async (input) => {
    const parsed = searchSheetSchema.parse(input);
    return jsonResponse(await searchSheetRows(workbookFromLocator(parsed), parsed));
  });

  server.tool("keyword_project_write_records", writeRecordsSchema.shape, async (input) => {
    const parsed = writeRecordsSchema.parse(input);
    return jsonResponse(await writeProjectRecords(workbookFromLocator(parsed), parsed));
  });

  server.tool("keyword_project_update_record", updateRecordSchema.shape, async (input) => {
    const parsed = updateRecordSchema.parse(input);
    return jsonResponse(await updateRecord(workbookFromLocator(parsed), parsed));
  });

  server.tool("keyword_project_delete_records", deleteRecordsSchema.shape, async (input) => {
    const parsed = deleteRecordsSchema.parse(input);
    return jsonResponse(await deleteRecords(workbookFromLocator(parsed), parsed));
  });

  server.tool("keyword_project_filter_raw_keywords", filterRawKeywordsSchema.shape, async (input) => {
    const parsed = filterRawKeywordsSchema.parse(input);
    return jsonResponse(await filterRawKeywords(workbookFromLocator(parsed), parsed));
  });

  server.tool("keyword_project_transfer_raw_keywords", transferRawKeywordsSchema.shape, async (input) => {
    const parsed = transferRawKeywordsSchema.parse(input);
    return jsonResponse(await transferRawKeywordsToMaster(workbookFromLocator(parsed), parsed));
  });

  server.tool("keyword_project_screen_raw_keywords", screenRawKeywordsSchema.shape, async (input) => {
    const parsed = screenRawKeywordsSchema.parse(input);
    return jsonResponse(await screenRawKeywords(workbookFromLocator(parsed), parsed));
  });

  server.tool("keyword_project_list_sources", workbookLocatorSchema.shape, async (input) => {
    return jsonResponse(await listSources(workbookFromLocator(workbookLocatorSchema.parse(input))));
  });
}
