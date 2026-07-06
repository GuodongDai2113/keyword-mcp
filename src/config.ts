import path from "node:path";
import type { ServerConfig } from "./types/keywordProject.js";

/** 根据环境变量创建 MCP 服务配置。 */
export function createServerConfig(env: NodeJS.ProcessEnv = process.env): ServerConfig {
  const projectsRoot = path.resolve(env.KEYWORD_MCP_PROJECTS_ROOT ?? ".");
  return {
    projectsRoot,
    defaultMarket: env.KEYWORD_MCP_DEFAULT_MARKET ?? "global",
    defaultLanguage: env.KEYWORD_MCP_DEFAULT_LANGUAGE ?? "en",
  };
}
