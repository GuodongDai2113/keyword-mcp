# Keyword MCP 设计文档

日期：2026-07-06

## 背景

当前仓库 `J:\project\keyword-mcp` 基本为空，且不是 Git 仓库。参考项目位于 `Q:\ai-playground\.opencode\skills\keyword-screening`，它已经用 Python 实现了关键词项目初始化、SEMrush 数据导入、工作簿读取、搜索、写入等能力。

本项目要把这些能力设计成 TypeScript MCP server。用户已确认继续使用 Excel 工作簿作为事实源：一个关键词项目维护一个 `keyword-plan.xlsx`，导入的关键词数据来自其他表格文件。

## 目标

1. 使用 TypeScript 实现一个 MCP server，向客户端暴露关键词项目维护工具。
2. 一个项目对应一个目录和一个 `keyword-plan.xlsx` 工作簿。
3. 支持项目创建、项目概览、数据源导入、表格读取、搜索、人工维护表写入、更新和删除。
4. 保留现有 Python 参考项目的数据结构和工作边界，避免迁移成本。
5. 所有修改类操作都写入 `更新记录`，便于追溯。

## 非目标

1. 初版不做 Web UI。
2. 初版不引入数据库，`keyword-plan.xlsx` 是唯一事实源。
3. 初版不自动做关键词筛选决策，只提供结构化读写能力；筛选规则由使用 MCP 的智能体或人执行。
4. 初版只支持 SEMrush 导出文件；其他来源通过后续字段映射扩展。
5. 初版不做多人并发协同编辑，只在 MCP 进程内串行化写入。

## 推荐方案

采用 TypeScript 原生 MCP server，使用 `exceljs` 直接维护 Excel 工作簿，使用 `zod` 校验 MCP tool 输入。

这个方案比直接包装 Python 脚本更适合作为长期 MCP 项目：类型定义、schema、错误返回和测试都可以保持在 TypeScript 生态内。同时它仍然保留参考项目的 Excel 事实源，不需要改变现有关键词筛选工作方式。

## 项目文件结构

默认项目根目录为 `keyword-projects/`：

```text
keyword-projects/<project-slug>/
├── data-sources/
├── reports/
└── keyword-plan.xlsx
```

`data-sources/` 平铺存放导入来源文件，不按来源嵌套目录。导入外部文件时，MCP 会复制一份到该目录，便于审计和重复导入判断。

## 工作簿结构

标准工作表：

```text
项目信息
数据来源
原始关键词
关键词主表
内容规划
更新记录
```

`项目信息` 字段：

```text
字段
值
```

默认项：

```text
项目名称
网站
目标市场
语言
核心产品
创建时间
最后更新时间
```

`数据来源` 字段：

```text
导入时间
来源
来源文件
归档文件
导入行数
未映射字段
备注
```

`原始关键词` 字段：

```text
导入时间
来源
来源文件
关键词
搜索量
关键词难度
点击均价(CPC)
搜索意图
SERP功能
```

`关键词主表` 字段：

```text
关键词
搜索量
关键词难度
点击均价(CPC)
搜索意图
SERP功能
优先级
备注
筛选时间
```

`内容规划` 字段：

```text
页面主题
主关键词
次关键词
搜索意图
页面类型
优先级
状态
目标URL
备注
```

`更新记录` 字段：

```text
变更时间
操作
详情
```

## MCP Tools

### keyword_project_create

创建项目目录和标准工作簿。

输入：

```ts
{
  name: string;
  site?: string;
  market?: string;
  language?: string;
  product?: string;
  root?: string;
}
```

行为：

1. 根据 `name` 生成稳定 slug。
2. 创建 `<root>/<slug>/data-sources` 和 `<root>/<slug>/reports`。
3. 若工作簿不存在，创建标准工作表、表头和初始更新记录。
4. 若工作簿已存在，返回既有路径，不覆盖。

### keyword_project_overview

返回项目概览。

输入：

```ts
{
  projectPath?: string;
  workbookPath?: string;
}
```

行为：

1. 定位 `keyword-plan.xlsx`。
2. 返回工作簿路径、项目目录、每张表的数据行数、表头。
3. 返回 `data-sources/` 下可导入文件列表。
4. 返回最近的数据来源记录。

### keyword_project_import_source

导入来源表格到 `原始关键词`。

输入：

```ts
{
  projectPath?: string;
  workbookPath?: string;
  source?: "semrush";
  filePaths?: string[];
  skipExisting?: boolean;
}
```

行为：

1. 默认扫描项目 `data-sources/` 下的 `.csv`、`.tsv`、`.xlsx`、`.xls` 文件。
2. 只支持 SEMrush 字段映射。
3. 复制外部来源文件到 `data-sources/`。
4. 同名 `来源文件` 已存在时默认跳过。
5. 把来源字段转换为中文标准字段后追加到 `原始关键词`。
6. 写入 `数据来源` 和 `更新记录`。

SEMrush 默认字段映射：

```json
{
  "Keyword": "关键词",
  "Intent": "搜索意图",
  "Volume": "搜索量",
  "Keyword Difficulty": "关键词难度",
  "CPC (USD)": "点击均价(CPC)",
  "SERP Features": "SERP功能"
}
```

### keyword_project_read_sheet

分页读取工作表。

输入：

```ts
{
  projectPath?: string;
  workbookPath?: string;
  sheet: string;
  start?: number;
  limit?: number;
  columns?: string[];
  sort?: Array<{ column: string; direction?: "asc" | "desc" }>;
}
```

行为：

1. 表头行不计入 `start`，`start=1` 表示第一条数据行。
2. 默认最多返回 20 行。
3. `columns` 只返回实际存在的列。
4. `sort` 只接受当前表中存在的列。

### keyword_project_search_sheet

在指定工作表中搜索。

输入：

```ts
{
  projectPath?: string;
  workbookPath?: string;
  sheet: string;
  query: string;
  limit?: number;
  columns?: string[];
  searchColumns?: string[];
  sort?: Array<{ column: string; direction?: "asc" | "desc" }>;
}
```

行为：

1. 对指定搜索列做大小写不敏感包含匹配。
2. 未指定搜索列时搜索全部列。
3. 返回有限数量的匹配行和 Excel 行号。

### keyword_project_write_records

向人工维护表写入记录。

输入：

```ts
{
  projectPath?: string;
  workbookPath?: string;
  sheet: "关键词主表" | "内容规划";
  records: Array<Record<string, unknown>>;
  mode?: "append" | "upsert";
  key?: string;
}
```

行为：

1. 只允许写入 `关键词主表` 和 `内容规划`。
2. `append` 直接追加。
3. `upsert` 使用 key 查找既有行，找到则更新，找不到则追加。
4. 默认 key：`关键词主表` 使用 `关键词`，`内容规划` 使用 `主关键词`。
5. 未知字段只报告，不写入工作簿。
6. 空记录、缺少 key 的 upsert 记录会跳过并报告原因。
7. 写入后更新 `项目信息` 的最后更新时间，并追加 `更新记录`。

### keyword_project_update_record

按 key 更新单条记录。

输入：

```ts
{
  projectPath?: string;
  workbookPath?: string;
  sheet: "关键词主表" | "内容规划";
  key: string;
  value: string;
  patch: Record<string, unknown>;
}
```

行为：

1. 只允许更新人工维护表。
2. 找到第一条匹配行后，只更新 `patch` 中存在且属于表头的字段。
3. 未找到时返回 `updated: 0`，不自动追加。

### keyword_project_delete_records

按 key 删除人工维护表记录。

输入：

```ts
{
  projectPath?: string;
  workbookPath?: string;
  sheet: "关键词主表" | "内容规划";
  key: string;
  values: string[];
}
```

行为：

1. 只允许删除 `关键词主表` 和 `内容规划`。
2. 不允许删除 `原始关键词`、`数据来源`、`更新记录`。
3. 从下往上删除匹配行，避免行号偏移。
4. 返回删除数量和未匹配的 key 值。
5. 追加 `更新记录`。

### keyword_project_list_sources

列出数据源文件和导入记录。

输入：

```ts
{
  projectPath?: string;
  workbookPath?: string;
}
```

行为：

1. 返回 `data-sources/` 下可导入文件。
2. 返回 `数据来源` 表中已有导入记录。
3. 标记文件是否已导入。

## 模块设计

```text
src/
├── index.ts
├── config.ts
├── constants/
│   ├── sheets.ts
│   ├── headers.ts
│   └── sourceMappings.ts
├── schemas/
│   └── tools.ts
├── services/
│   ├── projectService.ts
│   ├── workbookService.ts
│   ├── importService.ts
│   └── queryService.ts
├── tools/
│   ├── createProject.ts
│   ├── overview.ts
│   ├── importSource.ts
│   ├── readSheet.ts
│   ├── searchSheet.ts
│   ├── writeRecords.ts
│   ├── updateRecord.ts
│   ├── deleteRecords.ts
│   └── listSources.ts
└── types/
    └── keywordProject.ts
```

模块职责：

1. `index.ts` 注册 MCP server 和 tools。
2. `config.ts` 读取环境变量，提供默认项目根目录。
3. `constants/` 保存工作表名、表头、来源字段映射。
4. `schemas/` 用 `zod` 定义输入校验。
5. `projectService.ts` 负责项目路径、slug、工作簿定位。
6. `workbookService.ts` 负责 Excel 创建、表头维护、读写行、更新记录。
7. `importService.ts` 负责来源文件读取、字段映射、数据标准化。
8. `queryService.ts` 负责分页读取、搜索和排序。
9. `tools/` 只做 MCP 输入解析、调用 service、返回结构化结果。

## 路径和安全边界

1. 默认只允许操作配置的 `projectsRoot` 内部项目。
2. `workbookPath` 和 `projectPath` 解析后必须位于允许根目录内。
3. 导入外部文件时允许读取外部路径，但归档副本必须落到目标项目的 `data-sources/`。
4. 不暴露任意文件删除能力。
5. 不允许通过 sheet 参数写入任意工作表。

## 并发和一致性

初版 MCP 在单进程内对同一工作簿写入串行执行。实现方式可以用内存中的 `Map<string, Promise<void>>` 或轻量队列按工作簿路径排队。

该设计不能防止 Excel 客户端或另一个进程同时编辑同一文件。遇到保存失败、文件被占用或权限错误时，MCP 返回明确错误，不做静默重试覆盖。

## 错误处理

所有 tools 返回结构化结果。可恢复问题作为结果字段返回，例如：

```ts
{
  skipped: 2,
  unknownFields: [...],
  rows: [...]
}
```

不可恢复问题抛出 MCP tool error，例如：

1. 工作簿不存在。
2. sheet 不存在。
3. 来源文件不存在。
4. 来源字段无法识别关键词列。
5. 请求写入不允许写入的工作表。
6. 路径越过允许根目录。

## 测试策略

使用 Vitest。

核心测试：

1. 创建项目时生成标准目录、工作簿、工作表和表头。
2. 重复创建项目不覆盖已有工作簿。
3. 导入 SEMrush CSV 时写入 `原始关键词`、`数据来源`、`更新记录`。
4. 重复导入同名来源文件时跳过。
5. 分页读取能返回指定列和 Excel 行号。
6. 搜索能按指定列匹配并限制数量。
7. `write_records append` 能写入人工维护表并报告未知字段。
8. `write_records upsert` 能更新既有记录并追加新记录。
9. `update_record` 找不到 key 时不追加。
10. `delete_records` 只能删除人工维护表，不能删除原始关键词。
11. 非允许根目录路径被拒绝。
12. 所有修改操作都会追加更新记录。

## 中文注释要求

根据用户的 AGENTS.md 指令，后续 TypeScript 实现中，所有函数、属性、类都必须写完整中文注释。设计中的模块和代码应避免匿名复杂逻辑，公共 API、class、interface、schema 字段和主要 helper 都应有中文说明。

## 实施顺序

1. 初始化 TypeScript MCP 项目和测试框架。
2. 定义常量、类型、schema 和配置。
3. 实现项目创建和工作簿基础服务。
4. 实现读取、搜索和概览 tools。
5. 实现 SEMrush 导入。
6. 实现写入、更新、删除人工维护表。
7. 补齐路径安全、写入串行化和错误返回。
8. 完成 Vitest 覆盖。

## 待用户确认

1. 项目根目录默认使用 `keyword-projects/`。
2. 初版只支持 SEMrush 来源。
3. 删除功能只允许作用于 `关键词主表` 和 `内容规划`。
4. 不引入数据库，`keyword-plan.xlsx` 是唯一事实源。
