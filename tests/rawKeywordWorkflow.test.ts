import fs from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { createProject } from "../src/services/projectService.js";
import { importKeywordSources } from "../src/services/importService.js";
import { filterRawKeywords, screenRawKeywords, transferRawKeywordsToMaster } from "../src/services/rawKeywordService.js";
import { readSheetWindow, searchSheetRows } from "../src/services/queryService.js";
import { createTempWorkspace } from "./helpers/workspace.js";

describe("原始关键词筛选流程", () => {
  it("默认查询原始关键词时忽略已过滤行，并可显式包含已过滤行", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Filter Raw Project" });
    const sourceFile = path.join(project.directory, "filter.csv");
    await fs.writeFile(sourceFile, "Keyword,Intent,Volume,Keyword Difficulty,CPC (USD),SERP Features\nkeep keyword,C,100,10,1.1,People also ask\nskip keyword,I,20,30,0.2,FAQ\n", "utf8");
    await importKeywordSources(project.workbook, { source: "semrush", filePaths: [sourceFile], skipExisting: true });

    const filterResult = await filterRawKeywords(project.workbook, { keywords: ["skip keyword"], filterValue: "已过滤" });
    const defaultSearch = await searchSheetRows(project.workbook, { sheet: "原始关键词", query: "keyword", searchColumns: ["关键词"], columns: ["关键词", "过滤"], limit: 10 });
    const includeFilteredSearch = await searchSheetRows(project.workbook, { sheet: "原始关键词", query: "keyword", searchColumns: ["关键词"], columns: ["关键词", "过滤"], includeFiltered: true, limit: 10 });

    expect(filterResult.updated).toBe(1);
    expect(defaultSearch.rows.map((row) => row.values)).toEqual([{ 关键词: "keep keyword", 过滤: "否" }]);
    expect(includeFilteredSearch.rows.map((row) => row.values)).toEqual([{ 关键词: "keep keyword", 过滤: "否" }, { 关键词: "skip keyword", 过滤: "已过滤" }]);
  });

  it("把原始关键词转移到关键词主表后自动设定原始行过滤", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Transfer Raw Project" });
    const sourceFile = path.join(project.directory, "transfer.csv");
    await fs.writeFile(sourceFile, "Keyword,Intent,Volume,Keyword Difficulty,CPC (USD),SERP Features\ntransfer keyword,C,90,8,0.8,People also ask\n", "utf8");
    await importKeywordSources(project.workbook, { source: "semrush", filePaths: [sourceFile], skipExisting: true });

    const result = await transferRawKeywordsToMaster(project.workbook, {
      keywords: ["transfer keyword"],
      priority: "高",
      note: "高相关商业意图，进入主表",
      mode: "upsert",
    });
    const masterRows = await searchSheetRows(project.workbook, { sheet: "关键词主表", query: "transfer keyword", searchColumns: ["关键词"], columns: ["关键词", "搜索量", "优先级", "备注", "筛选时间", "SERP功能"] });
    const rawRows = await readSheetWindow(project.workbook, { sheet: "原始关键词", start: 1, limit: 5, columns: ["关键词", "过滤"], includeFiltered: true });

    expect(result.transferred).toBe(1);
    expect(masterRows.selectedHeaders).not.toContain("SERP功能");
    expect(masterRows.rows[0]?.values).toMatchObject({ 关键词: "transfer keyword", 搜索量: 90, 优先级: "高", 备注: "高相关商业意图，进入主表" });
    expect(masterRows.rows[0]?.values["筛选时间"]).toEqual(expect.any(String));
    expect(rawRows.rows[0]?.values).toEqual({ 关键词: "transfer keyword", 过滤: "已过滤" });
  });

  it("支持一次提交多条过滤和多条转移，每条可使用独立参数", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Batch Raw Project" });
    const sourceFile = path.join(project.directory, "batch.csv");
    await fs.writeFile(sourceFile, "Keyword,Intent,Volume,Keyword Difficulty,CPC (USD)\nfilter a,I,10,20,0.1\nfilter b,I,20,30,0.2\ntransfer a,C,100,8,1.1\ntransfer b,T,80,12,0.9\n", "utf8");
    await importKeywordSources(project.workbook, { source: "semrush", filePaths: [sourceFile], skipExisting: true });

    const filterResult = await filterRawKeywords(project.workbook, {
      items: [{ keyword: "filter a" }, { keyword: "filter b", filterValue: "已过滤" }],
    });
    const transferResult = await transferRawKeywordsToMaster(project.workbook, {
      items: [
        { keyword: "transfer a", priority: "高", note: "批量转移 A" },
        { keyword: "transfer b", priority: "中", note: "批量转移 B" },
      ],
      mode: "upsert",
    });
    const masterRows = await searchSheetRows(project.workbook, { sheet: "关键词主表", query: "transfer", searchColumns: ["关键词"], columns: ["关键词", "优先级", "备注"], limit: 10 });
    const rawRows = await searchSheetRows(project.workbook, { sheet: "原始关键词", query: "filter", searchColumns: ["关键词"], columns: ["关键词", "过滤"], includeFiltered: true, limit: 10 });

    expect(filterResult.updated).toBe(2);
    expect(transferResult.transferred).toBe(2);
    expect(masterRows.rows.map((row) => row.values)).toEqual([
      { 关键词: "transfer a", 优先级: "高", 备注: "批量转移 A" },
      { 关键词: "transfer b", 优先级: "中", 备注: "批量转移 B" },
    ]);
    expect(rawRows.rows.map((row) => row.values)).toEqual([
      { 关键词: "filter a", 过滤: "已过滤" },
      { 关键词: "filter b", 过滤: "已过滤" },
    ]);
  });

  it("支持一次提交混合筛选决策，过滤不用词并转移保留词", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Screen Raw Project" });
    const sourceFile = path.join(project.directory, "screen.csv");
    await fs.writeFile(sourceFile, "Keyword,Intent,Volume,Keyword Difficulty,CPC (USD)\nscreen filter,I,10,20,0.1\nscreen transfer,C,100,8,1.1\n", "utf8");
    await importKeywordSources(project.workbook, { source: "semrush", filePaths: [sourceFile], skipExisting: true });

    const result = await screenRawKeywords(project.workbook, {
      decisions: [
        { action: "filter", keyword: "screen filter", filterValue: "已过滤" },
        { action: "transfer", keyword: "screen transfer", priority: "高", note: "一次提交筛选决策" },
      ],
      mode: "upsert",
    });
    const defaultRawRows = await searchSheetRows(project.workbook, { sheet: "原始关键词", query: "screen", searchColumns: ["关键词"], columns: ["关键词", "过滤"], limit: 10 });
    const allRawRows = await searchSheetRows(project.workbook, { sheet: "原始关键词", query: "screen", searchColumns: ["关键词"], columns: ["关键词", "过滤"], includeFiltered: true, limit: 10 });
    const masterRows = await searchSheetRows(project.workbook, { sheet: "关键词主表", query: "screen transfer", searchColumns: ["关键词"], columns: ["关键词", "优先级", "备注"], limit: 10 });

    expect(result.filtered).toBe(1);
    expect(result.transferred).toBe(1);
    expect(defaultRawRows.rows).toHaveLength(0);
    expect(allRawRows.rows.map((row) => row.values)).toEqual([
      { 关键词: "screen filter", 过滤: "已过滤" },
      { 关键词: "screen transfer", 过滤: "已过滤" },
    ]);
    expect(masterRows.rows[0]?.values).toEqual({ 关键词: "screen transfer", 优先级: "高", 备注: "一次提交筛选决策" });
  });
});
