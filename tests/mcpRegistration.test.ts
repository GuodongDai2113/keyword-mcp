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
      "keyword_project_filter_raw_keywords",
      "keyword_project_transfer_raw_keywords",
      "keyword_project_screen_raw_keywords",
      "keyword_project_list_sources",
      "keyword_project_manual_entry",
    ]);
  });
});
