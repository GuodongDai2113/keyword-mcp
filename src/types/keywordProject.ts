/** 关键词项目服务配置。 */
export interface ServerConfig {
  /** 允许 MCP 操作的项目根目录绝对路径。 */
  projectsRoot: string;
  /** 每个关键词项目使用的工作簿名。 */
  workbookName: string;
  /** 新建项目默认目标市场。 */
  defaultMarket: string;
  /** 新建项目默认语言。 */
  defaultLanguage: string;
}

/** 工作簿中的一行结构化数据。 */
export interface WorkbookRow {
  /** Excel 中的实际行号。 */
  excelRow: number;
  /** 按表头映射后的单元格值。 */
  values: Record<string, unknown>;
}

/** 工作表概览信息。 */
export interface SheetOverview {
  /** 工作表名称。 */
  name: string;
  /** 工作表最大行数。 */
  maxRow: number;
  /** 工作表最大列数。 */
  maxColumn: number;
  /** 不包含表头的数据行数。 */
  dataRows: number;
  /** 第一行读取到的表头。 */
  headers: string[];
}

/** 排序规则。 */
export interface SortRule {
  /** 用于排序的列名。 */
  column: string;
  /** 排序方向。 */
  direction: "asc" | "desc";
}

/** 文件系统中的数据源文件信息。 */
export interface SourceFileInfo {
  /** 来源文件名。 */
  name: string;
  /** 来源文件绝对路径。 */
  path: string;
  /** 来源文件扩展名。 */
  suffix: string;
  /** 来源文件字节大小。 */
  size: number;
  /** 是否已经存在于数据来源记录中。 */
  imported: boolean;
}
