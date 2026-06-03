import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { requireIncludes } from './check-docs-discovery-core.ts'

interface StaleFormulaReadbackChooserDiscoveryContext {
  readonly docsRoot: string
  readonly repoRoot: string
  readonly index: string
  readonly llms: string
  readonly llmsFull: string
  readonly readme: string
}

export async function requireStaleFormulaReadbackChooserDiscovery({
  docsRoot,
  repoRoot,
  index,
  llms,
  llmsFull,
  readme,
}: StaleFormulaReadbackChooserDiscoveryContext): Promise<void> {
  const chooser = await readFile(join(docsRoot, 'stale-formula-readback-chooser.md'), 'utf8')
  const headlessReadme = await readFile(join(repoRoot, 'packages', 'headless', 'README.md'), 'utf8')
  const packageReadmes = await Promise.all(
    [
      'packages/xlsx-formula-recalc/README.md',
      'packages/bilig-xlsx-formula-recalc/README.md',
      'packages/sheetjs-formula-recalc/README.md',
      'packages/bilig-sheetjs-formula-recalc/README.md',
      'packages/exceljs-formula-recalc/README.md',
      'packages/bilig-exceljs-formula-recalc/README.md',
    ].map(async (relativePath) => ({
      relativePath,
      content: await readFile(join(repoRoot, relativePath), 'utf8'),
    })),
  )

  for (const required of [
    'title: Stale formula readback chooser for Node.js',
    'answer public threads with a Bilig link',
    'stale-cache failure is exact',
    'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json',
    'npm exec --package @bilig/sheetjs-formula-recalc@latest -- sheetjs-recalc --demo --json',
    'npm exec --package @bilig/exceljs-formula-recalc@latest -- exceljs-recalc --demo --json',
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'Agent and JavaScript search map',
    'SheetJS xlsx formula result not updating',
    'ExcelJS formula result not updating after Node edits',
    'MCP spreadsheet formula server for coding agents',
    'successful file save is not formula readback.',
    'Last checked on 2026-06-03 from the npm downloads API',
    '2026-05-27 through 2026-06-02',
    '`xlsx` | 10,557,539',
    '`exceljs` | 8,644,921',
    '`@bilig/xlsx-formula-recalc` | 10,613',
    '`@bilig/workpaper` | 8,814',
    'Stack Overflow `63085785`',
    'Stack Overflow `44199441`',
    'Skip the reply when the thread is old',
    'Do not ask for stars in support answers.',
  ] as const) {
    requireIncludes(chooser, required, 'docs/stale-formula-readback-chooser.md')
  }

  requireIncludes(index, './stale-formula-readback-chooser.html', 'docs/index.html')
  requireIncludes(readme, 'docs/stale-formula-readback-chooser.md', 'README.md')
  requireIncludes(headlessReadme, 'docs/stale-formula-readback-chooser.md', 'packages/headless/README.md')
  requireIncludes(llms, 'https://proompteng.github.io/bilig/stale-formula-readback-chooser.html', 'docs/llms.txt')
  requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/stale-formula-readback-chooser.md', 'docs/llms.txt')
  requireIncludes(llmsFull, '## Stale Formula Readback Chooser For Node.js', 'docs/llms-full.txt')

  for (const { relativePath, content } of packageReadmes) {
    requireIncludes(content, 'stale formula readback chooser', relativePath)
    requireIncludes(content, 'https://proompteng.github.io/bilig/stale-formula-readback-chooser.html', relativePath)
  }
}
