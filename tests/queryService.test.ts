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
    await appendRows(project.workbook, "关键词主表", [{ 关键词: "alpha software" }, { 关键词: "beta hardware" }]);

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
