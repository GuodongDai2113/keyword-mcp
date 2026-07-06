/** SEMrush 来源字段到内部标准字段的映射。 */
export const SEMRUSH_COLUMN_ALIASES = {
  keyword: ["Keyword"],
  intent: ["Intent"],
  searchVolume: ["Volume"],
  keywordDifficulty: ["Keyword Difficulty"],
  cpc: ["CPC (USD)"],
  serpFeatures: ["SERP Features"],
} as const;

/** 导入时识别但不落入关键词工作簿的来源字段。 */
export const IGNORED_SOURCE_COLUMNS = new Set([
  "traffic",
  "organictraffic",
  "impressions",
  "clicks",
  "ctr",
  "clickthroughrate",
  "position",
  "averageposition",
  "avgposition",
  "competition",
  "competitivedensity",
  "com",
]);

/** 支持自动读取的来源文件扩展名。 */
export const SUPPORTED_SOURCE_SUFFIXES = new Set([".csv", ".tsv", ".xlsx", ".xls"]);
