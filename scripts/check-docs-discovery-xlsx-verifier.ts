function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

export function requireXlsxCorpusVerifierDiscovery(content: string): void {
  for (const required of [
    'title: Do not trust stale XLSX cached formula values',
    'canonical_url: https://proompteng.github.io/bilig/xlsx-corpus-verifier-walkthrough.html',
    '# Do not trust stale XLSX cached formula values',
    'Only the second report is an accuracy verdict.',
    '## Run the Excel oracle harness',
    'pnpm workpaper:xlsx-oracle -- prepare-oracle /path/to/workbooks "$OUT"',
    'cache-diagnostic.json',
    'excel-oracle-report.json',
    'missing_excel_oracle',
    '## What counts as a real mismatch',
    'fresh Excel expected value',
    'Bilig actual value',
    '## Cache diagnostic still has value',
    '## Put it in CI',
    'pnpm workpaper:xlsx-corpus:check -- /path/to/workbooks',
    '## Turn a miss into a contribution',
    'https://github.com/proompteng/bilig/issues/new/choose',
    'https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only',
    'https://github.com/proompteng/bilig/blob/main/packages/headless/README.md',
  ] as const) {
    requireIncludes(content, required, 'docs/xlsx-corpus-verifier-walkthrough.md')
  }
}
