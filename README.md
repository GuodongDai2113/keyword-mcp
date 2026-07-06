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
- `keyword_project_filter_raw_keywords`
- `keyword_project_transfer_raw_keywords`
- `keyword_project_screen_raw_keywords`
- `keyword_project_list_sources`
