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
