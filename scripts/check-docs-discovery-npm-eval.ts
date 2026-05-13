import { readFile, stat } from 'node:fs/promises'
import { join } from 'node:path'

function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

async function requireFile(path: string): Promise<void> {
  const info = await stat(path)
  if (!info.isFile()) {
    throw new Error(`${path} is not a file`)
  }
}

export async function requireNpmEvalDiscovery(
  repoRoot: string,
  docsRoot: string,
  readme: string,
  headlessReadme: string,
  headlessExampleReadme: string,
): Promise<void> {
  const sourcePath = join(repoRoot, 'examples', 'headless-workpaper', 'npm-eval.ts')
  const publicMirrorPath = join(docsRoot, 'npm-eval.ts')
  await Promise.all([requireFile(sourcePath), requireFile(publicMirrorPath)])

  const [source, publicMirror, tryPage] = await Promise.all([
    readFile(sourcePath, 'utf8'),
    readFile(publicMirrorPath, 'utf8'),
    readFile(join(docsRoot, 'try-bilig-headless-in-node.md'), 'utf8'),
  ])
  if (source !== publicMirror) {
    throw new Error('docs/npm-eval.ts must match examples/headless-workpaper/npm-eval.ts')
  }

  for (const [path, content] of [
    ['README.md', readme],
    ['packages/headless/README.md', headlessReadme],
    ['docs/try-bilig-headless-in-node.md', tryPage],
  ] as const) {
    requireIncludes(content, 'https://proompteng.github.io/bilig/npm-eval.ts', path)
    requireIncludes(content, 'examples/headless-workpaper/npm-eval.ts', path)
  }

  requireIncludes(headlessExampleReadme, 'npm run npm-eval', 'examples/headless-workpaper/README.md')
  requireIncludes(headlessExampleReadme, '## npm Package Eval', 'examples/headless-workpaper/README.md')
}
