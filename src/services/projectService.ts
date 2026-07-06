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
  /** 工作簿绝对路径（项目即工作簿文件，不再创建目录）。 */
  workbook: string;
}

/** 从项目名称生成工作簿文件名 {slug}-keywords.xlsx。 */
export function projectWorkbookName(projectName: string): string {
  return `${slugify(projectName)}-keywords.xlsx`;
}

/** 工作簿定位输入。 */
export interface ResolveWorkbookInput {
  /** 工作簿路径或目录路径。 */
  projectPath?: string;
  /** 工作簿路径。 */
  workbookPath?: string;
  /** 允许操作的项目根目录。 */
  projectsRoot: string;
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

/** 根据工作簿路径或项目路径定位工作簿文件。 */
export function resolveWorkbookPath(input: ResolveWorkbookInput): string {
  if (input.workbookPath) {
    const workbookPath = path.resolve(input.workbookPath);
    assertInsideRoot(workbookPath, input.projectsRoot);
    return workbookPath;
  }
  const resolved = path.resolve(input.projectPath ?? input.projectsRoot);
  assertInsideRoot(resolved, input.projectsRoot);
  if (resolved.endsWith(".xlsx")) return resolved;
  return resolved;
}

/** 创建工作簿文件（不再创建项目子目录）；既有工作簿不会被覆盖。 */
export async function createProject(input: CreateProjectInput): Promise<CreateProjectResult> {
  const root = path.resolve(input.root);
  const fileName = projectWorkbookName(input.name);
  const workbook = path.join(root, fileName);
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
  return { workbook };
}
