import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { requireFile, requireIncludes } from './check-docs-discovery-core.ts'

export interface SmolagentsWorkpaperToolDiscoveryContext {
  readonly repoRoot: string
  readonly docsRoot: string
  readonly index: string
  readonly llms: string
  readonly llmsFull: string
  readonly scopedWorkpaperPackageReadme: string
}

export async function requireSmolagentsWorkpaperToolDiscovery(context: SmolagentsWorkpaperToolDiscoveryContext): Promise<void> {
  const smolagentsWorkpaperTool = await readFile(join(context.docsRoot, 'smolagents-workpaper-tool.md'), 'utf8')

  await Promise.all(
    ['README.md', 'smolagents_workpaper_tool.py', 'scripts/check-smolagents-recipe.py'].map((sourceFile) =>
      requireFile(join(context.repoRoot, 'examples', 'smolagents-workpaper-tool', sourceFile)),
    ),
  )

  requireIncludes(context.index, './smolagents-workpaper-tool.html', 'docs/index.html')
  requireIncludes(smolagentsWorkpaperTool, 'examples/smolagents-workpaper-tool', 'docs/smolagents-workpaper-tool.md')
  requireIncludes(smolagentsWorkpaperTool, 'from smolagents import Tool', 'docs/smolagents-workpaper-tool.md')
  requireIncludes(smolagentsWorkpaperTool, 'verify_workpaper_formula_readback', 'docs/smolagents-workpaper-tool.md')
  requireIncludes(
    smolagentsWorkpaperTool,
    'https://huggingface.co/docs/smolagents/main/en/tutorials/tools',
    'docs/smolagents-workpaper-tool.md',
  )
  requireIncludes(context.llms, 'https://proompteng.github.io/bilig/smolagents-workpaper-tool.html', 'docs/llms.txt')
  requireIncludes(context.llms, 'https://github.com/proompteng/bilig/tree/main/examples/smolagents-workpaper-tool', 'docs/llms.txt')
  requireIncludes(context.llmsFull, 'smolagents WorkPaper Tool', 'docs/llms-full.txt')
  requireIncludes(context.scopedWorkpaperPackageReadme, 'Hugging Face smolagents tool', 'packages/workpaper/README.md')
  requireIncludes(
    context.scopedWorkpaperPackageReadme,
    'https://proompteng.github.io/bilig/smolagents-workpaper-tool.html',
    'packages/workpaper/README.md',
  )
}
