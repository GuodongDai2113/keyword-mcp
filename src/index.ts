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
