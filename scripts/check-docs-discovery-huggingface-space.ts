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
  const workpaperPackageJson = await readFile(join(context.repoRoot, 'packages', 'workpaper', 'package.json'), 'utf8')
  const spacePackageJson = await readFile(join(context.repoRoot, 'examples', 'huggingface-workpaper-space', 'package.json'), 'utf8')
  const spaceReadme = await readFile(join(context.repoRoot, 'examples', 'huggingface-workpaper-space', 'README.md'), 'utf8')
  const spaceProof = await readFile(join(context.repoRoot, 'examples', 'huggingface-workpaper-space', 'workpaper_proof.mjs'), 'utf8')
  const spaceCheck = await readFile(join(context.repoRoot, 'examples', 'huggingface-workpaper-space', 'scripts', 'check-space.py'), 'utf8')
  const workpaperPackageVersion = readStringField(workpaperPackageJson, 'version', 'packages/workpaper/package.json')
  const spaceDependencyVersion = readDependencyVersion(
    spacePackageJson,
    '@bilig/workpaper',
    'examples/huggingface-workpaper-space/package.json',
  )

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
  if (spaceDependencyVersion !== workpaperPackageVersion) {
    throw new Error(
      `examples/huggingface-workpaper-space/package.json must pin @bilig/workpaper to ${workpaperPackageVersion}, found ${spaceDependencyVersion}`,
    )
  }
  requireIncludes(spaceReadme, `@bilig/workpaper@${workpaperPackageVersion}`, 'examples/huggingface-workpaper-space/README.md')
  requireIncludes(spaceReadme, `"packageVersion": "${workpaperPackageVersion}"`, 'examples/huggingface-workpaper-space/README.md')
  requireIncludes(
    spaceProof,
    `const workpaperPackageVersion = '${workpaperPackageVersion}'`,
    'examples/huggingface-workpaper-space/workpaper_proof.mjs',
  )
  requireIncludes(
    spaceCheck,
    `payload.get("packageVersion") != "${workpaperPackageVersion}"`,
    'examples/huggingface-workpaper-space/scripts/check-space.py',
  )
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

function readDependencyVersion(packageJson: string, packageName: string, sourceName: string): string {
  const parsed = parseJsonObject(packageJson, sourceName)
  const dependencies = Reflect.get(parsed, 'dependencies')
  if (typeof dependencies !== 'object' || dependencies === null || Array.isArray(dependencies)) {
    throw new Error(`${sourceName} must define dependencies`)
  }
  const version = Reflect.get(dependencies, packageName)
  if (typeof version !== 'string') {
    throw new Error(`${sourceName} dependencies must include ${packageName}`)
  }
  return version
}

function readStringField(packageJson: string, fieldName: string, sourceName: string): string {
  const parsed = parseJsonObject(packageJson, sourceName)
  const value = Reflect.get(parsed, fieldName)
  if (typeof value !== 'string') {
    throw new Error(`${sourceName} must define string ${fieldName}`)
  }
  return value
}

function parseJsonObject(packageJson: string, sourceName: string): object {
  const parsed: unknown = JSON.parse(packageJson)
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error(`${sourceName} must be a JSON object`)
  }
  return parsed
}
