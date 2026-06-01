import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { requireIncludes } from './check-docs-discovery-core.ts'

interface XlsxFormulaSupportAnswersDiscoveryContext {
  readonly docsRoot: string
  readonly index: string
  readonly llms: string
  readonly readme: string
}

export async function requireXlsxFormulaSupportAnswersDiscovery({
  docsRoot,
  index,
  llms,
  readme,
}: XlsxFormulaSupportAnswersDiscoveryContext): Promise<void> {
  const xlsxFormulaSupportAnswers = await readFile(join(docsRoot, 'xlsx-formula-support-answers.md'), 'utf8')

  for (const required of [
    'title: XLSX formula support answers for Node.js',
    'Do not paste them\ninto old threads just to mention Bilig.',
    'SheetJS, ExcelJS, `xlsx-populate`, and template\nlibraries can write workbook bytes.',
    'npx --yes --package @bilig/sheetjs-formula-recalc sheetjs-recalc --demo --json',
    'npx --yes --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json',
    'npx --package @bilig/xlsx-formula-recalc xlsx-cache-doctor report.xlsx --json',
    "fail-on-stale: 'false'",
    'Skip the public reply when:',
  ] as const) {
    requireIncludes(xlsxFormulaSupportAnswers, required, 'docs/xlsx-formula-support-answers.md')
  }

  requireIncludes(index, './xlsx-formula-support-answers.html', 'docs/index.html')
  requireIncludes(readme, 'docs/xlsx-formula-support-answers.md', 'README.md')
  requireIncludes(llms, 'https://proompteng.github.io/bilig/xlsx-formula-support-answers.html', 'docs/llms.txt')
  requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/xlsx-formula-support-answers.md', 'docs/llms.txt')
}
