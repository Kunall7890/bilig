import { readdir, readFile } from 'node:fs/promises'
import { join, relative, sep } from 'node:path'

const publicSnippetFiles = [
  'README.md',
  'packages/headless/README.md',
  'docs/agent-workpaper-tool-calling-recipe.md',
  'docs/serverless-workpaper-api-route.md',
  'docs/building-a-revenue-model-with-headless-workpaper.md',
  'docs/unsupported-formula-troubleshooting-recipe.md',
  'docs/dev-to-workbook-apis-post.md',
] as const

const forbiddenPublicSnippetPatterns = [
  {
    pattern: /^```(?:js|javascript)$/gim,
    reason: 'uses a JavaScript code fence where the public example should be TypeScript',
  },
  {
    pattern: /\bCreate `[^`\n]*\.js`/g,
    reason: 'asks readers to create a JavaScript file for a public TypeScript example',
  },
  {
    pattern: /\bnode\s+[^`\n]*\.js\b/g,
    reason: 'runs a JavaScript file directly instead of the maintained TypeScript smoke path',
  },
  {
    pattern: /\.mjs\b/g,
    reason: 'mentions an mjs example path in a public TypeScript-first snippet surface',
  },
] as const

export async function requireTypeScriptFirstPublicSnippets(repoRoot: string): Promise<void> {
  const files = new Set<string>(publicSnippetFiles)
  for (const examplesMarkdownFile of await listMarkdownFiles(join(repoRoot, 'examples'))) {
    files.add(toRepoPath(repoRoot, examplesMarkdownFile))
  }

  await Promise.all(
    [...files].map(async (path) => {
      const content = await readFile(join(repoRoot, path), 'utf8')
      requireNoJavaScriptFirstSnippetMarkers(path, content)
    }),
  )
}

function requireNoJavaScriptFirstSnippetMarkers(path: string, content: string): void {
  for (const { pattern, reason } of forbiddenPublicSnippetPatterns) {
    pattern.lastIndex = 0
    const match = pattern.exec(content)
    if (match === null) {
      continue
    }

    throw new Error(`${path} ${reason}: ${match[0]}`)
  }
}

async function listMarkdownFiles(root: string): Promise<string[]> {
  const entries = await readdir(root, { withFileTypes: true })
  const files: string[] = []
  const childDirectories: string[] = []

  for (const entry of entries) {
    const path = join(root, entry.name)
    if (entry.isDirectory()) {
      if (entry.name === 'node_modules' || entry.name.startsWith('.')) {
        continue
      }
      childDirectories.push(path)
      continue
    }

    if (entry.isFile() && entry.name.endsWith('.md')) {
      files.push(path)
    }
  }

  const childFiles = await Promise.all(childDirectories.map((path) => listMarkdownFiles(path)))
  return [...files, ...childFiles.flat()]
}

function toRepoPath(repoRoot: string, path: string): string {
  return relative(repoRoot, path).split(sep).join('/')
}
