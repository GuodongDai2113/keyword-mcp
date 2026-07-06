/** 标准工作表名称，顺序与新建工作簿中的默认顺序保持一致。 */
export const SHEET_NAMES = [
  "项目信息",
  "数据来源",
  "原始关键词",
  "关键词主表",
  "更新记录",
] as const;

/** 允许人工写入、更新和删除的工作表名称。 */
export const WRITABLE_SHEET_NAMES = ["关键词主表"] as const;

/** 标准工作表名称类型。 */
export type SheetName = (typeof SHEET_NAMES)[number];

/** 可人工维护的工作表名称类型。 */
export type WritableSheetName = (typeof WRITABLE_SHEET_NAMES)[number];
