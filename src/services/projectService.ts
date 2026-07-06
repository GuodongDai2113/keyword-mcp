import fs from "node:fs/promises";
import path from "node:path";
import { createKeywordWorkbook } from "./workbookService.js";

/** 创建项目所需输入。 */
export interface CreateProjectInput {
  /** 项目根目录。 */
  root: string;
  /** 项目名称。 */
  name: string;
  /** 目标网站。 */
  site?: string;
  /** 目标市场。 */
  market?: string;
  /** 目标语言。 */
  language?: string;
  /** 核心产品或服务。 */
  product?: string;
}

/** 项目创建结果。 */
export interface CreateProjectResult {
  /** 项目目录绝对路径。 */
  directory: string;
  /** 工作簿绝对路径。 */
  workbook: string;
}

/** 工作簿定位输入。 */
export interface ResolveWorkbookInput {
  /** 项目目录路径。 */
  projectPath?: string;
  /** 工作簿路径。 */
  workbookPath?: string;
  /** 允许操作的项目根目录。 */
  projectsRoot: string;
  /** 工作簿名。 */
  workbookName?: string;
}

/** 把项目名称转换为稳定的英文小写目录名。 */
export function slugify(value: string): string {
  const slug = value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  return slug || "keyword-project";
}

/** 判断目标路径是否位于允许的项目根目录内。 */
export function assertInsideRoot(targetPath: string, projectsRoot: string): void {
  const root = path.resolve(projectsRoot);
  const target = path.resolve(targetPath);
  const relative = path.relative(root, target);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    throw new Error(`路径越过允许根目录：${target}`);
  }
}

/** 根据项目路径或工作簿路径定位 keyword-plan.xlsx。 */
export function resolveWorkbookPath(input: ResolveWorkbookInput): string {
  const workbookName = input.workbookName ?? "keyword-plan.xlsx";
  if (input.workbookPath) {
    const workbookPath = path.resolve(input.workbookPath);
    assertInsideRoot(workbookPath, input.projectsRoot);
    return workbookPath;
  }
  const projectPath = path.resolve(input.projectPath ?? input.projectsRoot);
  assertInsideRoot(projectPath, input.projectsRoot);
  return path.join(projectPath, workbookName);
}

/** 创建关键词项目目录和标准工作簿；既有工作簿不会被覆盖。 */
export async function createProject(input: CreateProjectInput): Promise<CreateProjectResult> {
  const root = path.resolve(input.root);
  const directory = path.join(root, slugify(input.name));
  const workbook = path.join(directory, "keyword-plan.xlsx");
  await fs.mkdir(path.join(directory, "data-sources"), { recursive: true });
  await fs.mkdir(path.join(directory, "reports"), { recursive: true });
  try {
    await fs.access(workbook);
  } catch {
    await createKeywordWorkbook(workbook, {
      name: input.name,
      site: input.site ?? "",
      market: input.market ?? "global",
      language: input.language ?? "en",
      product: input.product ?? "",
    });
  }
  return { directory, workbook };
}
