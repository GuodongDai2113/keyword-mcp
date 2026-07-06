import fs from "node:fs/promises";
import path from "node:path";
import { Workbook } from "exceljs";
import { describe, expect, it } from "vitest";
import { createServerConfig } from "../src/config.js";
import { MASTER_HEADERS, RAW_HEADERS } from "../src/constants/headers.js";
import { createProjectSchema, writeRecordsSchema } from "../src/schemas/tools.js";
import { createProject, resolveWorkbookPath } from "../src/services/projectService.js";
import { createTempWorkspace } from "./helpers/workspace.js";

describe("配置和 schema", () => {
  it("默认在当前目录创建项目", () => {
    const config = createServerConfig({});

    expect(config.projectsRoot).toBe(path.resolve("."));
    expect(config.defaultMarket).toBe("global");
    expect(config.defaultLanguage).toBe("en");
  });

  it("定义参考项目兼容的中文表头", () => {
    expect(RAW_HEADERS).toEqual(["导入时间", "来源", "来源文件", "关键词", "搜索量", "关键词难度", "点击均价(CPC)", "搜索意图", "过滤"]);
    expect(MASTER_HEADERS).not.toContain("SERP功能");
    expect(MASTER_HEADERS).toContain("筛选时间");
    expect(MASTER_HEADERS).toContain("词性");
    expect(MASTER_HEADERS).toContain("分类");
  });

  it("校验创建项目和写入记录输入", () => {
    expect(createProjectSchema.parse({ name: "Example Product" }).name).toBe("Example Product");
    expect(() => writeRecordsSchema.parse({ sheet: "原始关键词", records: [] })).toThrow();
    expect(() => writeRecordsSchema.parse({ sheet: "内容规划", records: [] })).toThrow();
  });
});

describe("项目创建和工作簿定位", () => {
  it("创建工作簿文件，命名为 {slug}-keywords.xlsx", async () => {
    const workspace = await createTempWorkspace();
    const result = await createProject({
      root: workspace.root,
      name: "Example Product Keyword",
      site: "https://example.com",
      market: "us",
      language: "en",
      product: "Example Product",
    });

    expect(result.workbook.endsWith("example-product-keyword-keywords.xlsx")).toBe(true);

    const workbook = new Workbook();
    await workbook.xlsx.readFile(result.workbook);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["项目信息", "数据来源", "原始关键词", "关键词主表", "更新记录"]);
    expect(workbook.getWorksheet("关键词主表")?.getRow(1).values).toContain("关键词");
  });

  it("拒绝定位项目根目录外的工作簿", async () => {
    const workspace = await createTempWorkspace();
    const outside = path.resolve(workspace.root, "..", "outside", "test-keywords.xlsx");

    expect(() => resolveWorkbookPath({ workbookPath: outside, projectsRoot: workspace.root })).toThrow("路径越过允许根目录");
  });

  it("重复创建时不覆盖既有工作簿", async () => {
    const workspace = await createTempWorkspace();
    const first = await createProject({ root: workspace.root, name: "Repeat Project" });
    const markerPath = path.join(workspace.root, "marker.txt");
    await fs.writeFile(markerPath, "keep", "utf8");
    const second = await createProject({ root: workspace.root, name: "Repeat Project" });

    expect(second.workbook).toBe(first.workbook);
    await expect(fs.readFile(markerPath, "utf8")).resolves.toBe("keep");
  });

  it("新建工作簿时应用便于阅读的表格样式", async () => {
    const workspace = await createTempWorkspace();
    const result = await createProject({ root: workspace.root, name: "Styled Project" });

    const workbook = new Workbook();
    await workbook.xlsx.readFile(result.workbook);
    const rawSheet = workbook.getWorksheet("原始关键词");
    const masterSheet = workbook.getWorksheet("关键词主表");

    expect(rawSheet?.autoFilter).toEqual("A1:I1");
    expect(rawSheet?.views[0]).toMatchObject({ state: "frozen", ySplit: 1 });
    expect(rawSheet?.properties.tabColor?.argb).toBe("FF5B9BD5");
    expect(rawSheet?.getRow(1).height).toBe(24);
    expect(rawSheet?.getRow(1).getCell(1).alignment).toMatchObject({ horizontal: "center", vertical: "middle", wrapText: true });
    expect(rawSheet?.getRow(1).getCell(1).fill).toMatchObject({
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1F4E78" },
    });
    expect(rawSheet?.getColumn(4).width).toBeGreaterThanOrEqual(32);
    expect(rawSheet?.getColumn(5).numFmt).toBe("#,##0");
    expect(rawSheet?.getColumn(7).numFmt).toBe("$0.00");
    expect(rawSheet?.getColumn(9).width).toBeGreaterThanOrEqual(12);

    expect(masterSheet?.properties.tabColor?.argb).toBe("FFC55A11");
    expect(masterSheet?.autoFilter).toEqual("A1:K1");
    expect(masterSheet?.getRow(1).getCell(1).alignment).toMatchObject({ horizontal: "center", vertical: "middle", wrapText: true });
    expect(masterSheet?.getColumn(10).width).toBeGreaterThanOrEqual(42);
    const classificationValidation = (masterSheet as unknown as { dataValidations?: { model?: unknown } })?.dataValidations?.model;
    expect(classificationValidation).toBeDefined();
    expect(masterSheet?.getColumn(1).width).toBeGreaterThanOrEqual(32);
    expect(workbook.getWorksheet("内容规划")).toBeUndefined();
  });
});
