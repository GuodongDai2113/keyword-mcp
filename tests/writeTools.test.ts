import { describe, expect, it } from "vitest";
import { createProject } from "../src/services/projectService.js";
import { deleteRecords, updateRecord, writeProjectRecords } from "../src/services/workbookService.js";
import { readSheetWindow, searchSheetRows } from "../src/services/queryService.js";
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

  it("append 写入时自动填充筛选时间", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Auto Time Project" });

    await writeProjectRecords(project.workbook, { sheet: "关键词主表", records: [{ 关键词: "auto-time", 搜索量: 50 }], mode: "append" });
    const rows = await readSheetWindow(project.workbook, { sheet: "关键词主表", start: 1, limit: 5, columns: ["关键词", "筛选时间"] });

    expect(rows.rows[0]?.values["筛选时间"]).toEqual(expect.any(String));
    expect((rows.rows[0]?.values["筛选时间"] as string).length).toBeGreaterThan(0);
  });

  it("upsert 更新关键词主表既有行并追加新行", async () => {
    const workspace = await createTempWorkspace();
    const project = await createProject({ root: workspace.root, name: "Upsert Project" });
    await writeProjectRecords(project.workbook, { sheet: "关键词主表", records: [{ 关键词: "alpha", 优先级: "低" }], mode: "append" });

    const result = await writeProjectRecords(project.workbook, { sheet: "关键词主表", records: [{ 关键词: "alpha", 优先级: "高" }, { 关键词: "beta", 优先级: "低" }], mode: "upsert" });

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
      withWorkbookWrite("same.xlsx", async () => {
        order.push("a");
      }),
      withWorkbookWrite("same.xlsx", async () => {
        order.push("b");
      }),
    ]);

    expect(order).toEqual(["a", "b"]);
  });
});
