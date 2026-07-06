import { z } from "zod";

/** 可写工作表 schema。 */
export const writableSheetSchema = z.enum(["关键词主表"]);

/** 排序规则 schema。 */
export const sortRuleSchema = z.object({
  column: z.string().min(1),
  direction: z.enum(["asc", "desc"]).default("asc"),
});

/** 创建项目 tool 输入 schema。 */
export const createProjectSchema = z.object({
  name: z.string().min(1),
  site: z.string().default(""),
  market: z.string().optional(),
  language: z.string().optional(),
  product: z.string().default(""),
  root: z.string().optional(),
});

/** 定位项目工作簿的通用输入 schema。 */
export const workbookLocatorSchema = z.object({
  projectPath: z.string().optional(),
  workbookPath: z.string().optional(),
});

/** 导入来源 tool 输入 schema。 */
export const importSourceSchema = workbookLocatorSchema.extend({
  source: z.literal("semrush").default("semrush"),
  filePaths: z.array(z.string()).optional(),
  skipExisting: z.boolean().default(true),
});

/** 读取工作表 tool 输入 schema。 */
export const readSheetSchema = workbookLocatorSchema.extend({
  sheet: z.string().min(1),
  start: z.number().int().positive().default(1),
  limit: z.number().int().min(0).max(500).default(20),
  columns: z.array(z.string()).optional(),
  sort: z.array(sortRuleSchema).optional(),
  includeFiltered: z.boolean().default(false),
});

/** 搜索工作表 tool 输入 schema。 */
export const searchSheetSchema = readSheetSchema.extend({
  query: z.string(),
  searchColumns: z.array(z.string()).optional(),
});

/** 批量写入记录 tool 输入 schema。 */
export const writeRecordsSchema = workbookLocatorSchema.extend({
  sheet: writableSheetSchema,
  records: z.array(z.record(z.unknown())),
  mode: z.enum(["append", "upsert"]).default("append"),
  key: z.string().optional(),
});

/** 单条更新记录 tool 输入 schema。 */
export const updateRecordSchema = workbookLocatorSchema.extend({
  sheet: writableSheetSchema,
  key: z.string().min(1),
  value: z.string().min(1),
  patch: z.record(z.unknown()),
});

/** 删除记录 tool 输入 schema。 */
export const deleteRecordsSchema = workbookLocatorSchema.extend({
  sheet: writableSheetSchema,
  key: z.string().min(1),
  values: z.array(z.string().min(1)).min(1),
});

/** 原始关键词定位 schema。 */
export const rawKeywordSelectorSchema = z.object({
  rawExcelRows: z.array(z.number().int().min(2)).optional(),
  keywords: z.array(z.string().min(1)).optional(),
});

/** 单条原始关键词过滤项 schema。 */
export const filterRawKeywordItemSchema = z.object({
  rawExcelRow: z.number().int().min(2).optional(),
  keyword: z.string().min(1).optional(),
  filterValue: z.enum(["否", "已过滤"]).default("已过滤"),
});

/** 单条原始关键词转移项 schema。 */
export const transferRawKeywordItemSchema = z.object({
  rawExcelRow: z.number().int().min(2).optional(),
  keyword: z.string().min(1).optional(),
  priority: z.enum(["高", "中", "低"]),
  note: z.string().min(1),
});

/** 设置原始关键词过滤状态 tool 输入 schema。 */
export const filterRawKeywordsSchema = workbookLocatorSchema.merge(rawKeywordSelectorSchema).extend({
  filterValue: z.enum(["否", "已过滤"]).default("已过滤"),
  items: z.array(filterRawKeywordItemSchema).optional(),
});

/** 转移原始关键词到关键词主表 tool 输入 schema。 */
export const transferRawKeywordsSchema = workbookLocatorSchema.merge(rawKeywordSelectorSchema).extend({
  priority: z.enum(["高", "中", "低"]).optional(),
  note: z.string().min(1).optional(),
  mode: z.enum(["append", "upsert"]).default("upsert"),
  key: z.string().default("关键词"),
  items: z.array(transferRawKeywordItemSchema).optional(),
});

/** 单条混合筛选决策 schema。 */
export const screenRawKeywordDecisionSchema = z.discriminatedUnion("action", [
  filterRawKeywordItemSchema.extend({ action: z.literal("filter") }),
  transferRawKeywordItemSchema.extend({ action: z.literal("transfer") }),
]);

/** 混合筛选决策 tool 输入 schema。 */
export const screenRawKeywordsSchema = workbookLocatorSchema.extend({
  decisions: z.array(screenRawKeywordDecisionSchema).min(1),
  mode: z.enum(["append", "upsert"]).default("upsert"),
  key: z.string().default("关键词"),
});
