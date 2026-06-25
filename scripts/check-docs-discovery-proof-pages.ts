import type { BenchmarkDiscoveryEvidence } from './check-docs-discovery-benchmark-evidence.ts'

type RequireIncludes = (haystack: string, needle: string, context: string) => void

export function requireFormulaProofDiscovery({
  benchmarkEvidence,
  communityLaunchPack,
  formulaWorkbooksProof,
  headlessReadme,
  index,
  llms,
  readme,
  requireIncludes,
  showHnFormulaWorkbooksProof,
}: {
  readonly benchmarkEvidence: BenchmarkDiscoveryEvidence
  readonly communityLaunchPack: string
  readonly formulaWorkbooksProof: string
  readonly headlessReadme: string
  readonly index: string
  readonly llms: string
  readonly readme: string
  readonly requireIncludes: RequireIncludes
  readonly showHnFormulaWorkbooksProof: string
}): void {
  for (const required of [
    'title: Formula workbooks for Node services and tool integrations',
    'npm install @bilig/workpaper',
    'quote-approval-api.ts',
    '"restoredMatchesAfter": true',
    'bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable',
    'Use HyperFormula first when you need a mature, broad formula engine',
    'Use SheetJS or ExcelJS first when the primary job is reading, writing, styling',
    'Use Google Sheets API first when a shared hosted spreadsheet',
    `The current checked benchmark artifact records \`${benchmarkEvidence.meanWinHeadline}\` comparable`,
    benchmarkEvidence.p95HoldoutWorkload,
    'https://github.com/proompteng/bilig',
    'https://github.com/proompteng/bilig/discussions/new?category=general',
    'concrete implementation-gap discussion',
  ] as const) {
    requireIncludes(formulaWorkbooksProof, required, 'docs/formula-workbooks-node-services-agent-tools.md')
  }

  requireIncludes(readme, 'formula workbooks proof page', 'README.md')
  requireIncludes(readme, 'docs/formula-workbooks-node-services-agent-tools.md', 'README.md')
  requireIncludes(headlessReadme, 'formula workbooks for Node services and tool integrations', 'packages/headless/README.md')
  requireIncludes(headlessReadme, 'docs/formula-workbooks-node-services-agent-tools.md', 'packages/headless/README.md')
  requireIncludes(
    communityLaunchPack,
    'https://proompteng.github.io/bilig/formula-workbooks-node-services-agent-tools.html',
    'internal/growth/community-launch-pack.md',
  )
  requireIncludes(
    communityLaunchPack,
    'Hacker News Submission After The Formula Workbooks Page',
    'internal/growth/community-launch-pack.md',
  )

  for (const required of [
    "title: 'Bilig maintainer note: formula WorkPapers for Node services and tool hosts'",
    'uses the latest published package',
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'npm install @bilig/workpaper',
    '"verified": true',
    `wins \`${benchmarkEvidence.meanWinHeadline}\` comparable`,
    `\`${benchmarkEvidence.meanAndP95Headline}\` on both mean and p95`,
    `\`${benchmarkEvidence.p95HoldoutWorkload}\` is the current worst p95`,
    'Bilig maintainer note: formula WorkPapers for Node services and tool hosts',
    '## Review checklist',
    'bilig-evaluate --door workpaper-service --json',
    'bilig-evaluate --door agent-mcp --json',
  ] as const) {
    requireIncludes(showHnFormulaWorkbooksProof, required, 'docs/show-hn-formula-workbooks-node-services.md')
  }

  requireIncludes(index, './show-hn-formula-workbooks-node-services.html', 'docs/index.html')
  requireIncludes(llms, 'https://proompteng.github.io/bilig/show-hn-formula-workbooks-node-services.html', 'docs/llms.txt')
  requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/show-hn-formula-workbooks-node-services.md', 'docs/llms.txt')
}
