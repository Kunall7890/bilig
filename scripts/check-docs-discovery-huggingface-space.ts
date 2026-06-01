import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { requireFile, requireIncludes } from './check-docs-discovery-core.ts'

export interface HuggingFaceWorkpaperSpaceDiscoveryContext {
  readonly repoRoot: string
  readonly docsRoot: string
  readonly index: string
  readonly llms: string
  readonly llmsFull: string
  readonly readme: string
  readonly scopedWorkpaperPackageReadme: string
}

export async function requireHuggingFaceWorkpaperSpaceDiscovery(context: HuggingFaceWorkpaperSpaceDiscoveryContext): Promise<void> {
  const spaceDoc = await readFile(join(context.docsRoot, 'huggingface-workpaper-space.md'), 'utf8')

  await Promise.all(
    ['README.md', 'Dockerfile', 'app.py', 'workpaper_proof.mjs', 'package.json', 'requirements.txt', 'scripts/check-space.py'].map(
      (sourceFile) => requireFile(join(context.repoRoot, 'examples', 'huggingface-workpaper-space', sourceFile)),
    ),
  )

  requireIncludes(spaceDoc, 'https://huggingface.co/spaces/gregkonush/bilig-workpaper-mcp-readback', 'docs/huggingface-workpaper-space.md')
  requireIncludes(spaceDoc, 'examples/huggingface-workpaper-space', 'docs/huggingface-workpaper-space.md')
  requireIncludes(spaceDoc, 'gradio[mcp]', 'docs/huggingface-workpaper-space.md')
  requireIncludes(spaceDoc, 'Inputs!B3', 'docs/huggingface-workpaper-space.md')
  requireIncludes(spaceDoc, 'Summary!B3', 'docs/huggingface-workpaper-space.md')
  requireIncludes(spaceDoc, 'Deploy Your Own', 'docs/huggingface-workpaper-space.md')
  requireIncludes(spaceDoc, 'bilig-workpaper-mcp', 'docs/huggingface-workpaper-space.md')
  requireIncludes(context.index, './huggingface-workpaper-space.html', 'docs/index.html')
  requireIncludes(context.llms, 'https://proompteng.github.io/bilig/huggingface-workpaper-space.html', 'docs/llms.txt')
  requireIncludes(context.llms, 'https://huggingface.co/spaces/gregkonush/bilig-workpaper-mcp-readback', 'docs/llms.txt')
  requireIncludes(context.llmsFull, 'Hugging Face WorkPaper MCP Space', 'docs/llms-full.txt')
  requireIncludes(context.readme, 'Hugging Face Gradio MCP Space', 'README.md')
  requireIncludes(context.scopedWorkpaperPackageReadme, 'Hugging Face Gradio MCP Space', 'packages/workpaper/README.md')
  requireIncludes(
    context.scopedWorkpaperPackageReadme,
    'https://proompteng.github.io/bilig/huggingface-workpaper-space.html',
    'packages/workpaper/README.md',
  )
}
