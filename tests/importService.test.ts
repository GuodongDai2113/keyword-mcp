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
    const sourceFile = path.join(workspace.root, "semrush.csv");
    await fs.writeFile(sourceFile, "Keyword,Intent,Volume,Keyword Difficulty,CPC (USD),SERP Features\nexample keyword,C,100,25,1.2,People also ask\n", "utf8");

    const result = await importKeywordSources(project.workbook, { source: "semrush", filePaths: [sourceFile], skipExisting: true });

    expect(result.totalImportedRows).toBe(1);
    const rows = await searchSheetRows(project.workbook, { sheet: "原始关键词", query: "example", searchColumns: ["关键词"], columns: ["关键词", "搜索意图", "过滤"] });
    expect(rows.rows[0]?.values).toEqual({ 关键词: "example keyword", 搜索意图: "商业型", 过滤: "否" });
  });

  it("重复导入同名来源文件时跳过", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Skip Import Project" });
    const sourceFile = path.join(workspace.root, "repeat.csv");
    await fs.writeFile(sourceFile, "Keyword,Volume\nrepeat keyword,10\n", "utf8");

    await importKeywordSources(project.workbook, { source: "semrush", filePaths: [sourceFile], skipExisting: true });
    const second = await importKeywordSources(project.workbook, { source: "semrush", filePaths: [sourceFile], skipExisting: true });

    expect(second.importedFiles).toHaveLength(0);
    expect(second.skippedFiles[0]?.file).toBe("repeat.csv");
  });
});
