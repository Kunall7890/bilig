import { requireIncludes } from './check-docs-discovery-core.ts'

export function requireMicrosoftGraphExcelBoundaryDiscovery(input: {
  readonly doc: string
  readonly headlessReadme: string
  readonly index: string
  readonly llms: string
  readonly readme: string
}): void {
  const { doc, headlessReadme, index, llms, readme } = input
  for (const required of [
    'title: Microsoft Graph Excel recalculation vs local Node WorkPaper',
    'POST /me/drive/items/{id}/workbook/application/calculate',
    'Files.ReadWrite',
    'application permissions are not supported for that API',
    'Use `@bilig/workpaper` when the workbook is service-owned state',
    'https://learn.microsoft.com/en-us/graph/api/workbookapplication-calculate',
    'https://github.com/proompteng/bilig/stargazers',
  ] as const) {
    requireIncludes(doc, required, 'docs/microsoft-graph-excel-recalculation-node.md')
  }
  requireIncludes(index, './microsoft-graph-excel-recalculation-node.html', 'docs/index.html')
  requireIncludes(readme, 'docs/microsoft-graph-excel-recalculation-node.md', 'README.md')
  requireIncludes(headlessReadme, 'docs/microsoft-graph-excel-recalculation-node.md', 'packages/headless/README.md')
  requireIncludes(llms, 'https://proompteng.github.io/bilig/microsoft-graph-excel-recalculation-node.html', 'docs/llms.txt')
  requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/microsoft-graph-excel-recalculation-node.md', 'docs/llms.txt')
}
