/** 项目信息工作表表头。 */
export const PROJECT_INFO_HEADERS = ["字段", "值"] as const;

/** 数据来源工作表表头。 */
export const SOURCE_HEADERS = ["导入时间", "来源", "来源文件", "归档文件", "导入行数", "未映射字段", "备注"] as const;

/** 原始关键词工作表表头。 */
export const RAW_HEADERS = ["导入时间", "来源", "来源文件", "关键词", "搜索量", "关键词难度", "点击均价(CPC)", "搜索意图", "过滤"] as const;

/** 关键词主表工作表表头。 */
export const MASTER_HEADERS = ["关键词", "搜索量", "关键词难度", "点击均价(CPC)", "搜索意图", "来源", "词性", "分类", "优先级", "备注", "筛选时间"] as const;

/** 更新记录工作表表头。 */
export const CHANGELOG_HEADERS = ["变更时间", "操作", "详情"] as const;

/** 可写工作表到标准表头的映射。 */
export const WRITABLE_HEADERS = {
  关键词主表: MASTER_HEADERS,
} as const;
