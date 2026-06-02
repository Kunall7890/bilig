import { requireIncludes } from './check-docs-discovery-core.ts'

export function requireGoogleSheetsQuerySortnDiscovery(input: {
  readonly doc: string
  readonly headlessReadme: string
  readonly index: string
  readonly llms: string
  readonly llmsFull: string
  readonly readme: string
}): void {
  const { doc, headlessReadme, index, llms, llmsFull, readme } = input
  for (const required of [
    'title: Google Sheets QUERY and SORTN formulas in Node.js',
    '@bilig/workpaper',
    'QUERY(range, "select ... where ... order by ... limit ... offset ...", headers)',
    'QUERY` `group by` with `sum(column)` and `count(column)`',
    'QUERY` `label` for selected columns and supported aggregate output headers',
    'SORTN(range, n, tie_mode, sort_column_or_range, ascending, ...)',
    'Provider-backed imports such as `IMPORTDATA`, `IMPORTRANGE`, `IMPORTHTML`',
    'blocked result instead of pretending to have',
    'network or account access',
    '"verified": true',
    'https://support.google.com/docs/answer/3093343',
    'https://support.google.com/docs/answer/7354624',
  ] as const) {
    requireIncludes(doc, required, 'docs/google-sheets-query-sortn-node-workpaper.md')
  }

  requireIncludes(index, './google-sheets-query-sortn-node-workpaper.html', 'docs/index.html')
  requireIncludes(readme, 'docs/google-sheets-query-sortn-node-workpaper.md', 'README.md')
  requireIncludes(headlessReadme, 'docs/google-sheets-query-sortn-node-workpaper.md', 'packages/headless/README.md')
  requireIncludes(llms, 'https://proompteng.github.io/bilig/google-sheets-query-sortn-node-workpaper.html', 'docs/llms.txt')
  requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/google-sheets-query-sortn-node-workpaper.md', 'docs/llms.txt')
  requireIncludes(llmsFull, '## Google Sheets QUERY and SORTN in Node.js', 'docs/llms-full.txt')
}
