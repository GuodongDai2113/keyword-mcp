# Keyword MCP Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a TypeScript MCP server that maintains one `keyword-plan.xlsx` workbook per keyword project, with tools for project creation, import, reading, searching, writing, updating, and deleting.

**Architecture:** The MCP tool layer validates input with `zod` and delegates all behavior to focused services. `exceljs` is the workbook engine, `keyword-plan.xlsx` remains the only source of truth, and write operations are serialized per workbook path.

**Tech Stack:** TypeScript, Node.js ESM, `@modelcontextprotocol/sdk`, `exceljs`, `zod`, `tsx`, `vitest`.

---

## File Structure

- Create: `package.json` - npm scripts, runtime dependencies, dev dependencies.
- Create: `tsconfig.json` - strict TypeScript ESM configuration.
- Create: `vitest.config.ts` - Vitest config for Node tests.
- Create: `src/index.ts` - MCP stdio server entry and tool registration.
- Create: `src/config.ts` - environment-backed configuration and allowed project root resolution.
- Create: `src/constants/sheets.ts` - standard sheet names.
- Create: `src/constants/headers.ts` - standard workbook headers.
- Create: `src/constants/sourceMappings.ts` - SEMrush source field mapping and ignored fields.
- Create: `src/types/keywordProject.ts` - shared TypeScript interfaces with Chinese property comments.
- Create: `src/schemas/tools.ts` - zod schemas for all MCP tool inputs.
- Create: `src/services/projectService.ts` - slugging, path resolution, workbook discovery, path safety.
- Create: `src/services/workbookService.ts` - workbook creation, sheet/header maintenance, row read/write/update/delete, changelog.
- Create: `src/services/queryService.ts` - paginated reads, search, sorting.
- Create: `src/services/importService.ts` - CSV/TSV/XLSX source import and SEMrush normalization.
- Create: `src/services/writeQueue.ts` - per-workbook write serialization.
- Create: `src/tools/registerTools.ts` - MCP tool registration and response formatting.
- Create: `tests/helpers/workspace.ts` - temporary test workspace helpers.
- Create: `tests/projectService.test.ts` - project creation and path safety tests.
- Create: `tests/queryService.test.ts` - read/search behavior tests.
- Create: `tests/importService.test.ts` - SEMrush import behavior tests.
- Create: `tests/writeTools.test.ts` - write/update/delete behavior tests.
- Create: `tests/mcpRegistration.test.ts` - tool registration smoke test.

Current workspace is not a Git repository. Commit steps in this plan are recorded as skipped unless Git is initialized before execution.

---

### Task 1: Project Scaffold

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `vitest.config.ts`

- [ ] **Step 1: Create package manifest**

Create `package.json` with this content:

```json
{
  "name": "keyword-mcp",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "bin": {
    "keyword-mcp": "dist/index.js"
  },
  "scripts": {
    "build": "tsc -p tsconfig.json",
    "dev": "tsx src/index.ts",
    "test": "vitest run",
    "test:watch": "vitest",
    "typecheck": "tsc -p tsconfig.json --noEmit"
  },
  "dependencies": {
    "@modelcontextprotocol/sdk": "^1.17.0",
    "exceljs": "^4.4.0",
    "zod": "^3.25.76"
  },
  "devDependencies": {
    "@types/node": "^24.0.10",
    "tsx": "^4.20.3",
    "typescript": "^5.8.3",
    "vitest": "^3.2.4"
  }
}
```

- [ ] **Step 2: Create TypeScript configuration**

Create `tsconfig.json`:

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "strict": true,
    "esModuleInterop": true,
    "forceConsistentCasingInFileNames": true,
    "skipLibCheck": true,
    "outDir": "dist",
    "rootDir": ".",
    "types": ["node"]
  },
  "include": ["src/**/*.ts", "tests/**/*.ts", "vitest.config.ts"]
}
```

- [ ] **Step 3: Create Vitest configuration**

Create `vitest.config.ts`:

```ts
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["tests/**/*.test.ts"],
    testTimeout: 10000,
  },
});
```

- [ ] **Step 4: Install dependencies**

Run:

```powershell
npm install
```

Expected: `node_modules/` and `package-lock.json` are created, and npm exits with code `0`.

- [ ] **Step 5: Verify scaffold**

Run:

```powershell
npm run typecheck
```

Expected: TypeScript exits with code `0`; no source files exist yet, so the command only verifies config syntax.

---

### Task 2: Constants, Types, Schemas, And Config

**Files:**
- Create: `src/constants/sheets.ts`
- Create: `src/constants/headers.ts`
- Create: `src/constants/sourceMappings.ts`
- Create: `src/types/keywordProject.ts`
- Create: `src/schemas/tools.ts`
- Create: `src/config.ts`
- Test: `tests/projectService.test.ts`

- [ ] **Step 1: Write failing schema and config tests**

Create `tests/projectService.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { createServerConfig } from "../src/config.js";
import { createProjectSchema, writeRecordsSchema } from "../src/schemas/tools.js";
import { MASTER_HEADERS, RAW_HEADERS } from "../src/constants/headers.js";

describe("配置和 schema", () => {
  it("使用当前工作目录下的 keyword-projects 作为默认项目根目录", () => {
    const config = createServerConfig({});

    expect(config.projectsRoot.endsWith("keyword-projects")).toBe(true);
    expect(config.workbookName).toBe("keyword-plan.xlsx");
  });

  it("定义参考项目兼容的中文表头", () => {
    expect(RAW_HEADERS).toEqual([
      "导入时间",
      "来源",
      "来源文件",
      "关键词",
      "搜索量",
      "关键词难度",
      "点击均价(CPC)",
      "搜索意图",
      "SERP功能",
    ]);
    expect(MASTER_HEADERS).toContain("筛选时间");
  });

  it("校验创建项目和写入记录输入", () => {
    expect(createProjectSchema.parse({ name: "Example Product" }).name).toBe("Example Product");
    expect(() => writeRecordsSchema.parse({ sheet: "原始关键词", records: [] })).toThrow();
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
npm test -- tests/projectService.test.ts
```

Expected: FAIL because `src/config.ts`, `src/schemas/tools.ts`, and constants do not exist.

- [ ] **Step 3: Create constants**

Create `src/constants/sheets.ts`:

```ts
/** 标准工作表名称，顺序与新建工作簿中的默认顺序保持一致。 */
export const SHEET_NAMES = [
  "项目信息",
  "数据来源",
  "原始关键词",
  "关键词主表",
  "内容规划",
  "更新记录",
] as const;

/** 允许人工写入、更新和删除的工作表名称。 */
export const WRITABLE_SHEET_NAMES = ["关键词主表", "内容规划"] as const;

/** 标准工作表名称类型。 */
export type SheetName = (typeof SHEET_NAMES)[number];

/** 可人工维护的工作表名称类型。 */
export type WritableSheetName = (typeof WRITABLE_SHEET_NAMES)[number];
```

Create `src/constants/headers.ts`:

```ts
/** 项目信息工作表表头。 */
export const PROJECT_INFO_HEADERS = ["字段", "值"] as const;

/** 数据来源工作表表头。 */
export const SOURCE_HEADERS = ["导入时间", "来源", "来源文件", "归档文件", "导入行数", "未映射字段", "备注"] as const;

/** 原始关键词工作表表头。 */
export const RAW_HEADERS = ["导入时间", "来源", "来源文件", "关键词", "搜索量", "关键词难度", "点击均价(CPC)", "搜索意图", "SERP功能"] as const;

/** 关键词主表工作表表头。 */
export const MASTER_HEADERS = ["关键词", "搜索量", "关键词难度", "点击均价(CPC)", "搜索意图", "SERP功能", "优先级", "备注", "筛选时间"] as const;

/** 内容规划工作表表头。 */
export const CONTENT_PLAN_HEADERS = ["页面主题", "主关键词", "次关键词", "搜索意图", "页面类型", "优先级", "状态", "目标URL", "备注"] as const;

/** 更新记录工作表表头。 */
export const CHANGELOG_HEADERS = ["变更时间", "操作", "详情"] as const;

/** 可写工作表到标准表头的映射。 */
export const WRITABLE_HEADERS = {
  关键词主表: MASTER_HEADERS,
  内容规划: CONTENT_PLAN_HEADERS,
} as const;
```

Create `src/constants/sourceMappings.ts`:

```ts
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
```

- [ ] **Step 4: Create shared types**

Create `src/types/keywordProject.ts`:

```ts
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
```

- [ ] **Step 5: Create config and schemas**

Create `src/config.ts`:

```ts
import path from "node:path";
import type { ServerConfig } from "./types/keywordProject.js";

/** 根据环境变量创建 MCP 服务配置。 */
export function createServerConfig(env: NodeJS.ProcessEnv = process.env): ServerConfig {
  const projectsRoot = path.resolve(env.KEYWORD_MCP_PROJECTS_ROOT ?? "keyword-projects");
  return {
    projectsRoot,
    workbookName: env.KEYWORD_MCP_WORKBOOK_NAME ?? "keyword-plan.xlsx",
    defaultMarket: env.KEYWORD_MCP_DEFAULT_MARKET ?? "global",
    defaultLanguage: env.KEYWORD_MCP_DEFAULT_LANGUAGE ?? "en",
  };
}
```

Create `src/schemas/tools.ts`:

```ts
import { z } from "zod";

/** 可写工作表 schema。 */
export const writableSheetSchema = z.enum(["关键词主表", "内容规划"]);

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
```

- [ ] **Step 6: Run tests**

Run:

```powershell
npm test -- tests/projectService.test.ts
npm run typecheck
```

Expected: both commands pass.

---

### Task 3: Project And Workbook Creation

**Files:**
- Create: `src/services/projectService.ts`
- Create: `src/services/workbookService.ts`
- Modify: `tests/projectService.test.ts`

- [ ] **Step 1: Extend failing tests**

Append these tests to `tests/projectService.test.ts`:

```ts
import fs from "node:fs/promises";
import path from "node:path";
import { Workbook } from "exceljs";
import { createTempWorkspace } from "./helpers/workspace.js";
import { createProject, resolveWorkbookPath } from "../src/services/projectService.js";
import { createKeywordWorkbook } from "../src/services/workbookService.js";

describe("项目创建和工作簿定位", () => {
  it("创建项目目录和标准工作簿", async () => {
    const workspace = await createTempWorkspace();
    const result = await createProject({
      root: workspace.root,
      name: "Example Product Keyword",
      site: "https://example.com",
      market: "us",
      language: "en",
      product: "Example Product",
    });

    expect(result.directory.endsWith(path.join("example-product-keyword"))).toBe(true);
    expect(result.workbook.endsWith("keyword-plan.xlsx")).toBe(true);

    const workbook = new Workbook();
    await workbook.xlsx.readFile(result.workbook);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual([
      "项目信息",
      "数据来源",
      "原始关键词",
      "关键词主表",
      "内容规划",
      "更新记录",
    ]);
    expect(workbook.getWorksheet("关键词主表")?.getRow(1).values).toContain("关键词");
  });

  it("拒绝定位项目根目录外的工作簿", async () => {
    const workspace = await createTempWorkspace();
    const outside = path.resolve(workspace.root, "..", "outside", "keyword-plan.xlsx");

    expect(() => resolveWorkbookPath({ workbookPath: outside, projectsRoot: workspace.root })).toThrow("路径越过允许根目录");
  });

  it("重复创建时不覆盖既有工作簿", async () => {
    const workspace = await createTempWorkspace();
    const first = await createProject({ root: workspace.root, name: "Repeat Project" });
    await fs.writeFile(path.join(path.dirname(first.workbook), "marker.txt"), "keep", "utf8");
    const second = await createProject({ root: workspace.root, name: "Repeat Project" });

    expect(second.workbook).toBe(first.workbook);
    await expect(fs.readFile(path.join(path.dirname(first.workbook), "marker.txt"), "utf8")).resolves.toBe("keep");
  });
});
```

- [ ] **Step 2: Create test helper**

Create `tests/helpers/workspace.ts`:

```ts
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

/** 临时测试工作区。 */
export interface TempWorkspace {
  /** 临时项目根目录。 */
  root: string;
}

/** 创建用于测试的临时项目根目录。 */
export async function createTempWorkspace(): Promise<TempWorkspace> {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "keyword-mcp-"));
  return { root };
}
```

- [ ] **Step 3: Run tests to verify they fail**

Run:

```powershell
npm test -- tests/projectService.test.ts
```

Expected: FAIL because project and workbook services do not exist.

- [ ] **Step 4: Implement project service**

Create `src/services/projectService.ts`:

```ts
import fs from "node:fs/promises";
import path from "node:path";
import { createKeywordWorkbook } from "./workbookService.js";

/** 创建项目所需输入。 */
export interface CreateProjectInput {
  /** 项目根目录。 */
  root: string;
  /** 项目名称。 */
  name: string;
  /** 目标网站。 */
  site?: string;
  /** 目标市场。 */
  market?: string;
  /** 目标语言。 */
  language?: string;
  /** 核心产品或服务。 */
  product?: string;
}

/** 项目创建结果。 */
export interface CreateProjectResult {
  /** 项目目录绝对路径。 */
  directory: string;
  /** 工作簿绝对路径。 */
  workbook: string;
}

/** 把项目名称转换为稳定的英文小写目录名。 */
export function slugify(value: string): string {
  const slug = value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  return slug || "keyword-project";
}

/** 判断目标路径是否位于允许的项目根目录内。 */
export function assertInsideRoot(targetPath: string, projectsRoot: string): void {
  const root = path.resolve(projectsRoot);
  const target = path.resolve(targetPath);
  const relative = path.relative(root, target);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(`路径越过允许根目录：${target}`);
  }
}

/** 根据项目路径或工作簿路径定位 keyword-plan.xlsx。 */
export function resolveWorkbookPath(input: { projectPath?: string; workbookPath?: string; projectsRoot: string; workbookName?: string }): string {
  const workbookName = input.workbookName ?? "keyword-plan.xlsx";
  if (input.workbookPath) {
    const workbookPath = path.resolve(input.workbookPath);
    assertInsideRoot(workbookPath, input.projectsRoot);
    return workbookPath;
  }
  const projectPath = path.resolve(input.projectPath ?? input.projectsRoot);
  assertInsideRoot(projectPath, input.projectsRoot);
  return path.join(projectPath, workbookName);
}

/** 创建关键词项目目录和标准工作簿；既有工作簿不会被覆盖。 */
export async function createProject(input: CreateProjectInput): Promise<CreateProjectResult> {
  const root = path.resolve(input.root);
  const directory = path.join(root, slugify(input.name));
  const workbook = path.join(directory, "keyword-plan.xlsx");
  await fs.mkdir(path.join(directory, "data-sources"), { recursive: true });
  await fs.mkdir(path.join(directory, "reports"), { recursive: true });
  try {
    await fs.access(workbook);
  } catch {
    await createKeywordWorkbook(workbook, {
      name: input.name,
      site: input.site ?? "",
      market: input.market ?? "global",
      language: input.language ?? "en",
      product: input.product ?? "",
    });
  }
  return { directory, workbook };
}
```

- [ ] **Step 5: Implement workbook creation**

Create `src/services/workbookService.ts`:

```ts
import { Workbook, type Worksheet } from "exceljs";
import { CHANGELOG_HEADERS, CONTENT_PLAN_HEADERS, MASTER_HEADERS, PROJECT_INFO_HEADERS, RAW_HEADERS, SOURCE_HEADERS } from "../constants/headers.js";
import { SHEET_NAMES } from "../constants/sheets.js";

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

/** 返回统一的本地时间字符串。 */
export function nowText(): string {
  return new Date().toISOString().replace("T", " ").slice(0, 19);
}

/** 给工作表第一行应用基础表头样式。 */
export function styleHeaderRow(sheet: Worksheet): void {
  const row = sheet.getRow(1);
  row.font = { bold: true, color: { argb: "FFFFFFFF" } };
  row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
  sheet.views = [{ state: "frozen", ySplit: 1 }];
}

/** 向工作表写入表头并应用基础样式。 */
export function writeHeaders(sheet: Worksheet, headers: readonly string[]): void {
  sheet.addRow([...headers]);
  styleHeaderRow(sheet);
  headers.forEach((header, index) => {
    sheet.getColumn(index + 1).width = Math.max(12, Math.min(32, header.length + 6));
  });
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

/** 追加一条更新记录。 */
export function appendChangelog(workbook: Workbook, action: string, details: string): void {
  const sheet = workbook.getWorksheet("更新记录") ?? workbook.addWorksheet("更新记录");
  if (sheet.rowCount === 0) {
    writeHeaders(sheet, CHANGELOG_HEADERS);
  }
  sheet.addRow([nowText(), action, details]);
}

/** 创建标准关键词项目工作簿。 */
export async function createKeywordWorkbook(workbookPath: string, info: WorkbookProjectInfo): Promise<void> {
  const workbook = new Workbook();
  const first = workbook.addWorksheet(SHEET_NAMES[0]);
  writeProjectInfo(first, info);
  writeHeaders(workbook.addWorksheet("数据来源"), SOURCE_HEADERS);
  writeHeaders(workbook.addWorksheet("原始关键词"), RAW_HEADERS);
  writeHeaders(workbook.addWorksheet("关键词主表"), MASTER_HEADERS);
  writeHeaders(workbook.addWorksheet("内容规划"), CONTENT_PLAN_HEADERS);
  writeHeaders(workbook.addWorksheet("更新记录"), CHANGELOG_HEADERS);
  appendChangelog(workbook, "创建项目", `创建项目 ${info.name}`);
  await workbook.xlsx.writeFile(workbookPath);
}
```

- [ ] **Step 6: Run tests**

Run:

```powershell
npm test -- tests/projectService.test.ts
npm run typecheck
```

Expected: both commands pass.

---

### Task 4: Overview, Read, Search, And Sorting

**Files:**
- Modify: `src/services/workbookService.ts`
- Create: `src/services/queryService.ts`
- Modify: `tests/queryService.test.ts`

- [ ] **Step 1: Write failing query tests**

Create `tests/queryService.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { createProject } from "../src/services/projectService.js";
import { appendRows, getWorkbookOverview } from "../src/services/workbookService.js";
import { readSheetWindow, searchSheetRows } from "../src/services/queryService.js";
import { createTempWorkspace } from "./helpers/workspace.js";

describe("工作簿读取和搜索", () => {
  it("返回工作簿概览、表头和数据行数", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Overview Project" });
    await appendRows(project.workbook, "关键词主表", [{ 关键词: "alpha keyword", 搜索量: 100 }]);

    const overview = await getWorkbookOverview(project.workbook);

    expect(overview.sheets.find((sheet) => sheet.name === "关键词主表")?.dataRows).toBe(1);
  });

  it("分页读取指定列并保留 Excel 行号", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Read Project" });
    await appendRows(project.workbook, "关键词主表", [
      { 关键词: "alpha", 搜索量: 10 },
      { 关键词: "beta", 搜索量: 30 },
    ]);

    const result = await readSheetWindow(project.workbook, {
      sheet: "关键词主表",
      start: 1,
      limit: 1,
      columns: ["关键词"],
      sort: [{ column: "搜索量", direction: "desc" }],
    });

    expect(result.rows).toEqual([{ excelRow: 3, values: { 关键词: "beta" } }]);
  });

  it("按指定列搜索关键词", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Search Project" });
    await appendRows(project.workbook, "关键词主表", [
      { 关键词: "alpha software" },
      { 关键词: "beta hardware" },
    ]);

    const result = await searchSheetRows(project.workbook, {
      sheet: "关键词主表",
      query: "soft",
      limit: 10,
      columns: ["关键词"],
      searchColumns: ["关键词"],
    });

    expect(result.rows[0]?.values).toEqual({ 关键词: "alpha software" });
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
npm test -- tests/queryService.test.ts
```

Expected: FAIL because query functions and append helper do not exist.

- [ ] **Step 3: Extend workbook service**

Append to `src/services/workbookService.ts`:

```ts
import { WRITABLE_HEADERS } from "../constants/headers.js";

/** 读取工作表第一行表头。 */
export function readHeaders(sheet: Worksheet): string[] {
  const row = sheet.getRow(1);
  return row.values.slice(1).map((value) => (value == null ? "" : String(value)));
}

/** 确保工作表存在并返回工作表对象。 */
export function requireWorksheet(workbook: Workbook, sheetName: string): Worksheet {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw new Error(`工作表不存在：${sheetName}`);
  }
  return sheet;
}

/** 按工作表表头顺序追加多行字典数据。 */
export async function appendRows(workbookPath: string, sheetName: keyof typeof WRITABLE_HEADERS, records: Array<Record<string, unknown>>): Promise<void> {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(workbookPath);
  const sheet = requireWorksheet(workbook, sheetName);
  const headers = readHeaders(sheet);
  for (const record of records) {
    sheet.addRow(headers.map((header) => record[header]));
  }
  appendChangelog(workbook, "测试写入", `向 ${sheetName} 写入 ${records.length} 行`);
  await workbook.xlsx.writeFile(workbookPath);
}

/** 获取工作簿概览。 */
export async function getWorkbookOverview(workbookPath: string): Promise<{ workbook: string; projectDir: string; sheets: Array<{ name: string; maxRow: number; maxColumn: number; dataRows: number; headers: string[] }> }> {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(workbookPath);
  return {
    workbook: workbookPath,
    projectDir: workbookPath.replace(/[/\\]keyword-plan\.xlsx$/, ""),
    sheets: workbook.worksheets.map((sheet) => ({
      name: sheet.name,
      maxRow: sheet.rowCount,
      maxColumn: sheet.columnCount,
      dataRows: Math.max(0, sheet.rowCount - 1),
      headers: readHeaders(sheet),
    })),
  };
}
```

- [ ] **Step 4: Implement query service**

Create `src/services/queryService.ts`:

```ts
import { Workbook } from "exceljs";
import { readHeaders, requireWorksheet } from "./workbookService.js";
import type { SortRule, WorkbookRow } from "../types/keywordProject.js";

/** 工作表窗口读取参数。 */
export interface ReadSheetInput {
  /** 目标工作表名称。 */
  sheet: string;
  /** 从第几个数据行开始。 */
  start?: number;
  /** 返回行数限制。 */
  limit?: number;
  /** 需要返回的列。 */
  columns?: string[];
  /** 排序规则。 */
  sort?: SortRule[];
}

/** 工作表搜索参数。 */
export interface SearchSheetInput extends ReadSheetInput {
  /** 查询文本。 */
  query: string;
  /** 参与搜索的列。 */
  searchColumns?: string[];
}

/** 把单元格值转换为可排序值。 */
function sortableValue(value: unknown): string | number {
  if (typeof value === "number") {
    return value;
  }
  const text = String(value ?? "").trim();
  const number = Number(text.replace(/,/g, ""));
  return Number.isFinite(number) && text !== "" ? number : text.toLowerCase();
}

/** 按指定列从 Excel 行中取值。 */
function rowToValues(headers: string[], rowValues: unknown[], selectedHeaders: string[]): Record<string, unknown> {
  const result: Record<string, unknown> = {};
  for (const header of selectedHeaders) {
    const index = headers.indexOf(header);
    result[header] = index >= 0 ? rowValues[index + 1] : undefined;
  }
  return result;
}

/** 根据排序规则稳定排序工作簿行。 */
function sortRows(rows: WorkbookRow[], sortRules: SortRule[]): WorkbookRow[] {
  return [...rows].sort((left, right) => {
    for (const rule of sortRules) {
      const leftValue = sortableValue(left.values[rule.column]);
      const rightValue = sortableValue(right.values[rule.column]);
      if (leftValue < rightValue) return rule.direction === "desc" ? 1 : -1;
      if (leftValue > rightValue) return rule.direction === "desc" ? -1 : 1;
    }
    return left.excelRow - right.excelRow;
  });
}

/** 读取指定工作表的一段数据。 */
export async function readSheetWindow(workbookPath: string, input: ReadSheetInput): Promise<{ rows: WorkbookRow[]; headers: string[]; selectedHeaders: string[] }> {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(workbookPath);
  const sheet = requireWorksheet(workbook, input.sheet);
  const headers = readHeaders(sheet);
  const selectedHeaders = input.columns?.filter((column) => headers.includes(column)) ?? headers;
  const sortRules = input.sort?.filter((rule) => headers.includes(rule.column)) ?? [];
  const start = Math.max(1, input.start ?? 1);
  const limit = Math.max(0, input.limit ?? 20);
  const rows: WorkbookRow[] = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (!row.hasValues) continue;
    const values = row.values as unknown[];
    const allValues = rowToValues(headers, values, headers);
    rows.push({ excelRow: rowNumber, values: allValues });
  }
  const orderedRows = sortRules.length > 0 ? sortRows(rows, sortRules) : rows;
  const slicedRows = limit === 0 ? [] : orderedRows.slice(start - 1, start - 1 + limit);
  return {
    headers,
    selectedHeaders,
    rows: slicedRows.map((row) => ({ excelRow: row.excelRow, values: rowToValues(headers, [undefined, ...headers.map((header) => row.values[header])], selectedHeaders) })),
  };
}

/** 搜索指定工作表并返回匹配行。 */
export async function searchSheetRows(workbookPath: string, input: SearchSheetInput): Promise<{ rows: WorkbookRow[]; headers: string[]; selectedHeaders: string[] }> {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(workbookPath);
  const sheet = requireWorksheet(workbook, input.sheet);
  const headers = readHeaders(sheet);
  const selectedHeaders = input.columns?.filter((column) => headers.includes(column)) ?? headers;
  const searchableHeaders = input.searchColumns?.filter((column) => headers.includes(column)) ?? headers;
  const query = input.query.toLowerCase();
  const matched: WorkbookRow[] = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (!row.hasValues) continue;
    const values = row.values as unknown[];
    const fullValues = rowToValues(headers, values, headers);
    const isMatched = searchableHeaders.some((header) => String(fullValues[header] ?? "").toLowerCase().includes(query));
    if (isMatched) {
      matched.push({ excelRow: rowNumber, values: fullValues });
    }
  }
  const sortRules = input.sort?.filter((rule) => headers.includes(rule.column)) ?? [];
  const orderedRows = sortRules.length > 0 ? sortRows(matched, sortRules) : matched;
  const limit = Math.max(0, input.limit ?? 20);
  return {
    headers,
    selectedHeaders,
    rows: orderedRows.slice(0, limit).map((row) => ({ excelRow: row.excelRow, values: Object.fromEntries(selectedHeaders.map((header) => [header, row.values[header]])) })),
  };
}
```

- [ ] **Step 5: Run tests**

Run:

```powershell
npm test -- tests/queryService.test.ts
npm run typecheck
```

Expected: both commands pass. If TypeScript reports duplicate imports in `workbookService.ts`, merge imports at the file top while preserving the same exported functions.

---

### Task 5: SEMrush Source Import

**Files:**
- Create: `src/services/importService.ts`
- Modify: `tests/importService.test.ts`

- [ ] **Step 1: Write failing import tests**

Create `tests/importService.test.ts`:

```ts
import fs from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { createProject } from "../src/services/projectService.js";
import { importKeywordSources } from "../src/services/importService.js";
import { searchSheetRows } from "../src/services/queryService.js";
import { createTempWorkspace } from "./helpers/workspace.js";

describe("SEMrush 来源导入", () => {
  it("导入 CSV 到原始关键词并记录数据来源", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Import Project" });
    const sourceFile = path.join(project.directory, "semrush.csv");
    await fs.writeFile(sourceFile, "Keyword,Intent,Volume,Keyword Difficulty,CPC (USD),SERP Features\nexample keyword,C,100,25,1.2,People also ask\n", "utf8");

    const result = await importKeywordSources(project.workbook, { source: "semrush", filePaths: [sourceFile], skipExisting: true });

    expect(result.totalImportedRows).toBe(1);
    const rows = await searchSheetRows(project.workbook, { sheet: "原始关键词", query: "example", searchColumns: ["关键词"], columns: ["关键词", "搜索意图", "SERP功能"] });
    expect(rows.rows[0]?.values).toEqual({ 关键词: "example keyword", 搜索意图: "商业型", SERP功能: "其他用户还会问" });
  });

  it("重复导入同名来源文件时跳过", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Skip Import Project" });
    const sourceFile = path.join(project.directory, "repeat.csv");
    await fs.writeFile(sourceFile, "Keyword,Volume\nrepeat keyword,10\n", "utf8");

    await importKeywordSources(project.workbook, { source: "semrush", filePaths: [sourceFile], skipExisting: true });
    const second = await importKeywordSources(project.workbook, { source: "semrush", filePaths: [sourceFile], skipExisting: true });

    expect(second.importedFiles).toHaveLength(0);
    expect(second.skippedFiles[0]?.file).toBe("repeat.csv");
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
npm test -- tests/importService.test.ts
```

Expected: FAIL because `importService.ts` does not exist.

- [ ] **Step 3: Implement import service**

Create `src/services/importService.ts`:

```ts
import fs from "node:fs/promises";
import path from "node:path";
import { Workbook } from "exceljs";
import { IGNORED_SOURCE_COLUMNS, SEMRUSH_COLUMN_ALIASES, SUPPORTED_SOURCE_SUFFIXES } from "../constants/sourceMappings.js";
import { RAW_HEADERS, SOURCE_HEADERS } from "../constants/headers.js";
import { appendChangelog, nowText, readHeaders, requireWorksheet } from "./workbookService.js";

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
  const mapping: Record<string, string> = { i: "信息型", informational: "信息型", n: "导航型", navigational: "导航型", c: "商业型", commercial: "商业型", t: "交易型", transactional: "交易型" };
  const labels: string[] = [];
  for (const part of String(value ?? "").split(/[,;/|\s]+/)) {
    const label = mapping[part.trim().toLowerCase()];
    if (label && !labels.includes(label)) labels.push(label);
  }
  return labels.join("、");
}

/** 把 SEMrush SERP 功能转换为中文展示文本。 */
function normalizeSerpFeatures(value: string | undefined): string {
  const mapping: Record<string, string> = { "people also ask": "其他用户还会问", paa: "其他用户还会问", "featured snippet": "精选摘要", faq: "常见问题", images: "图片结果", video: "视频结果", videos: "视频结果" };
  const labels: string[] = [];
  for (const part of String(value ?? "").split(/[,;/|]+/)) {
    const text = part.trim();
    if (!text) continue;
    const label = mapping[text.toLowerCase()] ?? text;
    if (!labels.includes(label)) labels.push(label);
  }
  return labels.join(", ");
}

/** 解析简单 CSV 或 TSV 文本为对象数组。 */
function parseDelimited(text: string, delimiter: "," | "\t"): Array<Record<string, string>> {
  const lines = text.split(/\r?\n/).filter((line) => line.trim() !== "");
  const headers = lines[0]?.split(delimiter).map((value) => value.trim()) ?? [];
  return lines.slice(1).map((line) => {
    const values = line.split(delimiter);
    return Object.fromEntries(headers.map((header, index) => [header, values[index]?.trim() ?? ""]));
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
  const aliases = Object.fromEntries(Object.entries(SEMRUSH_COLUMN_ALIASES).flatMap(([standard, sourceAliases]) => sourceAliases.map((alias) => [normalizeColumn(alias), standard])));
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
function sourceExists(workbook: Workbook, fileName: string): boolean {
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

/** 发现 data-sources 目录下可导入的来源文件。 */
async function discoverSourceFiles(workbookPath: string): Promise<string[]> {
  const sourceDir = path.join(path.dirname(workbookPath), "data-sources");
  const entries = await fs.readdir(sourceDir, { withFileTypes: true }).catch(() => []);
  return entries.filter((entry) => entry.isFile() && SUPPORTED_SOURCE_SUFFIXES.has(path.extname(entry.name).toLowerCase()) && !entry.name.startsWith("~$")).map((entry) => path.join(sourceDir, entry.name)).sort();
}

/** 导入关键词来源文件到原始关键词表。 */
export async function importKeywordSources(workbookPath: string, input: ImportKeywordSourcesInput): Promise<ImportKeywordSourcesResult> {
  const filePaths = input.filePaths ?? (await discoverSourceFiles(workbookPath));
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
    const storedFile = path.join(path.dirname(workbookPath), "data-sources", fileName);
    await fs.mkdir(path.dirname(storedFile), { recursive: true });
    if (path.resolve(filePath) !== path.resolve(storedFile)) {
      await fs.copyFile(filePath, storedFile);
    }
    const importedAt = nowText();
    let importedRows = 0;
    for (const row of rows) {
      const standard = Object.fromEntries(Object.entries(mapping.mapped).map(([sourceColumn, standardColumn]) => [standardColumn, row[sourceColumn]]));
      const keyword = String(standard.keyword ?? "").trim();
      if (!keyword) continue;
      rawSheet.addRow(RAW_HEADERS.map((header) => ({
        导入时间: importedAt,
        来源: "SEMrush",
        来源文件: fileName,
        关键词: keyword,
        搜索量: toNumber(String(standard.searchVolume ?? "")),
        关键词难度: toNumber(String(standard.keywordDifficulty ?? "")),
        "点击均价(CPC)": toNumber(String(standard.cpc ?? "")),
        搜索意图: normalizeIntent(String(standard.intent ?? "")),
        SERP功能: normalizeSerpFeatures(String(standard.serpFeatures ?? "")),
      })[header])));
      importedRows += 1;
    }
    sourceSheet.addRow(SOURCE_HEADERS.map((header) => ({
      导入时间: importedAt,
      来源: "SEMrush",
      来源文件: fileName,
      归档文件: storedFile,
      导入行数: importedRows,
      未映射字段: mapping.unmapped.join(", "),
      备注: mapping.unmapped.length > 0 ? "存在未映射字段" : "导入完成",
    })[header]));
    importedFiles.push({ file: fileName, path: filePath, importedRows, unmappedColumns: mapping.unmapped, storedFile });
  }
  appendChangelog(workbook, "导入数据源", `导入 ${importedFiles.length} 个文件，合计 ${importedFiles.reduce((sum, item) => sum + item.importedRows, 0)} 行`);
  await workbook.xlsx.writeFile(workbookPath);
  return { importedFiles, skippedFiles, totalImportedRows: importedFiles.reduce((sum, item) => sum + item.importedRows, 0) };
}
```

- [ ] **Step 4: Run tests**

Run:

```powershell
npm test -- tests/importService.test.ts
npm run typecheck
```

Expected: both commands pass.

---

### Task 6: Write, Update, Delete, And Write Queue

**Files:**
- Create: `src/services/writeQueue.ts`
- Modify: `src/services/workbookService.ts`
- Create: `tests/writeTools.test.ts`

- [ ] **Step 1: Write failing mutation tests**

Create `tests/writeTools.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { createProject } from "../src/services/projectService.js";
import { deleteRecords, updateRecord, writeProjectRecords } from "../src/services/workbookService.js";
import { searchSheetRows } from "../src/services/queryService.js";
import { withWorkbookWrite } from "../src/services/writeQueue.js";
import { createTempWorkspace } from "./helpers/workspace.js";

describe("人工维护表写入、更新和删除", () => {
  it("append 写入关键词主表并报告未知字段", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Write Project" });

    const result = await writeProjectRecords(project.workbook, { sheet: "关键词主表", records: [{ 关键词: "alpha", 搜索量: 100, 未知字段: "ignored" }], mode: "append" });

    expect(result.inserted).toBe(1);
    expect(result.unknownFields[0]?.fields).toEqual(["未知字段"]);
  });

  it("upsert 更新既有行并追加新行", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Upsert Project" });
    await writeProjectRecords(project.workbook, { sheet: "内容规划", records: [{ 主关键词: "alpha", 状态: "待规划" }], mode: "append" });

    const result = await writeProjectRecords(project.workbook, { sheet: "内容规划", records: [{ 主关键词: "alpha", 状态: "已发布" }, { 主关键词: "beta", 状态: "待规划" }], mode: "upsert" });

    expect(result.updated).toBe(1);
    expect(result.inserted).toBe(1);
  });

  it("update 找不到 key 时不追加", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Update Project" });

    const result = await updateRecord(project.workbook, { sheet: "关键词主表", key: "关键词", value: "missing", patch: { 优先级: "高" } });

    expect(result.updated).toBe(0);
  });

  it("delete 只删除人工维护表记录", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Delete Project" });
    await writeProjectRecords(project.workbook, { sheet: "关键词主表", records: [{ 关键词: "alpha" }, { 关键词: "beta" }], mode: "append" });

    const result = await deleteRecords(project.workbook, { sheet: "关键词主表", key: "关键词", values: ["alpha", "missing"] });
    const rows = await searchSheetRows(project.workbook, { sheet: "关键词主表", query: "alpha", searchColumns: ["关键词"] });

    expect(result.deleted).toBe(1);
    expect(result.notFound).toEqual(["missing"]);
    expect(rows.rows).toHaveLength(0);
  });

  it("同一工作簿写入按队列串行执行", async () => {
    const order: string[] = [];
    await Promise.all([
      withWorkbookWrite("same.xlsx", async () => { order.push("a"); }),
      withWorkbookWrite("same.xlsx", async () => { order.push("b"); }),
    ]);

    expect(order).toEqual(["a", "b"]);
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
npm test -- tests/writeTools.test.ts
```

Expected: FAIL because mutation functions and write queue do not exist.

- [ ] **Step 3: Implement write queue**

Create `src/services/writeQueue.ts`:

```ts
/** 按工作簿路径保存的写入队列。 */
const workbookQueues = new Map<string, Promise<unknown>>();

/** 对同一个工作簿路径串行执行写入操作。 */
export async function withWorkbookWrite<T>(workbookPath: string, operation: () => Promise<T>): Promise<T> {
  const previous = workbookQueues.get(workbookPath) ?? Promise.resolve();
  const current = previous.then(operation, operation);
  workbookQueues.set(workbookPath, current.finally(() => {
    if (workbookQueues.get(workbookPath) === current) {
      workbookQueues.delete(workbookPath);
    }
  }));
  return current;
}
```

- [ ] **Step 4: Implement mutation functions**

Append to `src/services/workbookService.ts`:

```ts
import { withWorkbookWrite } from "./writeQueue.js";

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

/** 返回人工维护表默认 upsert 关键列。 */
function defaultWriteKey(sheetName: keyof typeof WRITABLE_HEADERS): string {
  return sheetName === "关键词主表" ? "关键词" : "主关键词";
}

/** 按关键列查找第一条匹配行。 */
function findRowByKey(sheet: Worksheet, headers: string[], key: string, value: unknown): number | undefined {
  const columnIndex = headers.indexOf(key) + 1;
  if (columnIndex <= 0) return undefined;
  const expected = String(value ?? "").trim();
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    if (String(sheet.getRow(rowNumber).getCell(columnIndex).value ?? "").trim() === expected) return rowNumber;
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

/** 向人工维护表追加或 upsert 记录。 */
export async function writeProjectRecords(workbookPath: string, input: WriteProjectRecordsInput): Promise<{ inserted: number; updated: number; skipped: number; unknownFields: Array<{ recordIndex: number; fields: string[] }> }> {
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
    input.records.forEach((record, index) => {
      const { known, unknown } = splitRecord(headers, record);
      if (unknown.length > 0) unknownFields.push({ recordIndex: index + 1, fields: unknown });
      if (Object.keys(known).length === 0) {
        skipped += 1;
        return;
      }
      if (input.mode === "upsert") {
        const keyValue = known[key];
        if (keyValue == null || keyValue === "") {
          skipped += 1;
          return;
        }
        const rowNumber = findRowByKey(sheet, headers, key, keyValue);
        if (rowNumber) {
          for (const [field, value] of Object.entries(known)) {
            if (value != null && value !== "") sheet.getRow(rowNumber).getCell(headers.indexOf(field) + 1).value = value as string | number | boolean | Date;
          }
          updated += 1;
          return;
        }
      }
      sheet.addRow(headers.map((header) => known[header]));
      inserted += 1;
    });
    appendChangelog(workbook, "写入项目", `向 ${input.sheet} 以 ${input.mode} 模式写入 ${inserted} 条，更新 ${updated} 条，跳过 ${skipped} 条`);
    await workbook.xlsx.writeFile(workbookPath);
    return { inserted, updated, skipped, unknownFields };
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
    for (const [field, value] of Object.entries(known)) {
      if (value != null && value !== "") sheet.getRow(rowNumber).getCell(headers.indexOf(field) + 1).value = value as string | number | boolean | Date;
    }
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
    appendChangelog(workbook, "删除记录", `从 ${input.sheet} 删除 ${deleted} 条记录`);
    await workbook.xlsx.writeFile(workbookPath);
    return { deleted, notFound: [...remaining] };
  });
}
```

- [ ] **Step 5: Run tests**

Run:

```powershell
npm test -- tests/writeTools.test.ts
npm run typecheck
```

Expected: both commands pass. If TypeScript reports import placement errors, consolidate all imports at the top of `src/services/workbookService.ts`.

---

### Task 7: MCP Tool Registration

**Files:**
- Create: `src/tools/registerTools.ts`
- Create: `src/index.ts`
- Create: `tests/mcpRegistration.test.ts`

- [ ] **Step 1: Write failing registration test**

Create `tests/mcpRegistration.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { registeredToolNames } from "../src/tools/registerTools.js";

describe("MCP tool 注册", () => {
  it("暴露关键词项目 MCP tools", () => {
    expect(registeredToolNames).toEqual([
      "keyword_project_create",
      "keyword_project_overview",
      "keyword_project_import_source",
      "keyword_project_read_sheet",
      "keyword_project_search_sheet",
      "keyword_project_write_records",
      "keyword_project_update_record",
      "keyword_project_delete_records",
      "keyword_project_list_sources",
    ]);
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
npm test -- tests/mcpRegistration.test.ts
```

Expected: FAIL because `registerTools.ts` does not exist.

- [ ] **Step 3: Implement MCP tool registration**

Create `src/tools/registerTools.ts`:

```ts
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { createServerConfig } from "../config.js";
import { createProject, resolveWorkbookPath } from "../services/projectService.js";
import { getWorkbookOverview, writeProjectRecords, updateRecord, deleteRecords } from "../services/workbookService.js";
import { importKeywordSources } from "../services/importService.js";
import { readSheetWindow, searchSheetRows } from "../services/queryService.js";
import { createProjectSchema, deleteRecordsSchema, importSourceSchema, readSheetSchema, searchSheetSchema, updateRecordSchema, workbookLocatorSchema, writeRecordsSchema } from "../schemas/tools.js";

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

/** 注册所有关键词项目 MCP tools。 */
export function registerKeywordTools(server: McpServer): void {
  server.tool("keyword_project_create", createProjectSchema.shape, async (input) => {
    const parsed = createProjectSchema.parse(input);
    const config = createServerConfig();
    return jsonResponse(await createProject({
      root: parsed.root ?? config.projectsRoot,
      name: parsed.name,
      site: parsed.site,
      market: parsed.market ?? config.defaultMarket,
      language: parsed.language ?? config.defaultLanguage,
      product: parsed.product,
    }));
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

  server.tool("keyword_project_list_sources", workbookLocatorSchema.shape, async (input) => {
    return jsonResponse(await getWorkbookOverview(workbookFromLocator(workbookLocatorSchema.parse(input))));
  });
}
```

- [ ] **Step 4: Implement MCP entrypoint**

Create `src/index.ts`:

```ts
#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerKeywordTools } from "./tools/registerTools.js";

/** 创建并启动关键词 MCP stdio 服务。 */
async function main(): Promise<void> {
  const server = new McpServer({ name: "keyword-mcp", version: "0.1.0" });
  registerKeywordTools(server);
  await server.connect(new StdioServerTransport());
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
```

- [ ] **Step 5: Run registration and full tests**

Run:

```powershell
npm test
npm run typecheck
npm run build
```

Expected: all commands pass and `dist/index.js` is created.

---

### Task 8: Final Verification And Usage Notes

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Create README**

Create `README.md`:

```md
# keyword-mcp

TypeScript MCP server for maintaining keyword projects. One project maps to one `keyword-plan.xlsx` workbook.

## Development

```powershell
npm install
npm test
npm run typecheck
npm run build
```

## Configuration

- `KEYWORD_MCP_PROJECTS_ROOT`: allowed keyword project root. Default: `keyword-projects`
- `KEYWORD_MCP_WORKBOOK_NAME`: workbook file name. Default: `keyword-plan.xlsx`
- `KEYWORD_MCP_DEFAULT_MARKET`: project default market. Default: `global`
- `KEYWORD_MCP_DEFAULT_LANGUAGE`: project default language. Default: `en`

## Tools

- `keyword_project_create`
- `keyword_project_overview`
- `keyword_project_import_source`
- `keyword_project_read_sheet`
- `keyword_project_search_sheet`
- `keyword_project_write_records`
- `keyword_project_update_record`
- `keyword_project_delete_records`
- `keyword_project_list_sources`
```

- [ ] **Step 2: Run final verification**

Run:

```powershell
npm test
npm run typecheck
npm run build
```

Expected: all commands pass.

- [ ] **Step 3: Check comments requirement**

Run:

```powershell
rg "export (class|interface|function|const)|function " src
```

Expected: every exported class, interface, function, const, and important helper is preceded by a complete Chinese comment.

- [ ] **Step 4: Commit status**

Run:

```powershell
git status --short
```

Expected in the current workspace: command fails with `not a git repository`. Do not run a commit command unless Git has been initialized by the user.

---

## Self-Review

Spec coverage:

1. Excel workbook as sole source of truth is covered by Tasks 3-6.
2. One project per `keyword-plan.xlsx` is covered by Task 3.
3. Standard sheets and headers are covered by Tasks 2-3.
4. Project create, overview, import, read, search, write, update, delete, and list-source tool names are covered by Task 7.
5. SEMrush-only import is covered by Task 5.
6. Write boundaries for `关键词主表` and `内容规划` are covered by Tasks 2 and 6.
7. Deletion restriction is covered by schema and service tests in Task 6.
8. Path safety is covered by Task 3.
9. Per-workbook write serialization is covered by Task 6.
10. Chinese comments requirement is covered by Task 8.

Placeholder scan:

The plan contains no placeholder implementation steps. Each task names exact files, commands, and expected outcomes.

Type consistency:

The plan consistently uses `CreateProjectInput`, `WriteProjectRecordsInput`, `UpdateRecordInput`, `DeleteRecordsInput`, `WorkbookRow`, and `SortRule`. Tool schemas map directly to service inputs.
