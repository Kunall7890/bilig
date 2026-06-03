import { requireIncludes } from './check-docs-discovery-core.ts'

export function requireXlsxRecalcPublicDiscovery(args: {
  readonly agentXlsxFormulaRecalculationWithoutLibreOffice: string
  readonly headlessReadme: string
  readonly index: string
  readonly liveSheetjsRecalcCli: string
  readonly llms: string
  readonly readme: string
  readonly xlsxFormulaRecalculationNode: string
  readonly xlsxCacheDoctorProofTranscript: string
  readonly xlsxCacheDoctorCli: string
  readonly xlsxRecalcCli: string
}): void {
  requireIncludes(args.readme, 'examples/xlsx-recalculation-node', 'README.md')
  requireIncludes(args.readme, 'examples/recalc-bridge-workflows', 'README.md')
  requireIncludes(args.readme, 'docs/xlsx-formula-recalculation-node.md', 'README.md')
  requireIncludes(args.readme, args.xlsxRecalcCli, 'README.md')
  requireIncludes(args.readme, args.xlsxCacheDoctorCli, 'README.md')
  requireIncludes(args.readme, args.liveSheetjsRecalcCli, 'README.md')
  requireIncludes(args.readme, 'docs/agent-xlsx-formula-recalculation-without-libreoffice.md', 'README.md')
  requireIncludes(args.readme, 'docs/excel-file-calculation-engine-node.md', 'README.md')
  requireIncludes(args.readme, 'docs/exceljs-shared-formula-recalculation-node.md', 'README.md')
  requireIncludes(args.headlessReadme, 'examples/xlsx-recalculation-node', 'packages/headless/README.md')
  requireIncludes(args.headlessReadme, 'docs/xlsx-formula-recalculation-node.md', 'packages/headless/README.md')
  requireIncludes(args.headlessReadme, args.xlsxRecalcCli, 'packages/headless/README.md')
  requireIncludes(args.headlessReadme, args.liveSheetjsRecalcCli, 'packages/headless/README.md')
  requireIncludes(
    args.headlessReadme,
    'https://proompteng.github.io/bilig/agent-xlsx-formula-recalculation-without-libreoffice.html',
    'packages/headless/README.md',
  )
  requireIncludes(args.headlessReadme, 'docs/excel-file-calculation-engine-node.md', 'packages/headless/README.md')
  requireIncludes(args.headlessReadme, 'docs/exceljs-shared-formula-recalculation-node.md', 'packages/headless/README.md')
  requireIncludes(args.index, 'examples/xlsx-recalculation-node', 'docs/index.html')
  requireIncludes(args.index, '<code>@bilig/xlsx-formula-recalc</code> when XLSX is the boundary.', 'docs/index.html')
  requireIncludes(args.index, './xlsx-formula-recalculation-node.html', 'docs/index.html')
  requireIncludes(args.index, './xlsx-cache-doctor-proof-transcript.html', 'docs/index.html')
  requireIncludes(args.index, './xlsx-cache-doctor-github-action.html', 'docs/index.html')
  requireIncludes(args.index, './sheetjs-formula-result-not-updating-node.html', 'docs/index.html')
  requireIncludes(args.llms, 'https://github.com/proompteng/bilig/tree/main/examples/xlsx-recalculation-node', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/xlsx-formula-recalculation-node.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://github.com/proompteng/bilig/blob/main/docs/xlsx-formula-recalculation-node.md', 'docs/llms.txt')
  requireIncludes(args.llms, args.xlsxRecalcCli, 'docs/llms.txt')
  requireIncludes(args.llms, args.xlsxCacheDoctorCli, 'docs/llms.txt')
  requireIncludes(args.llms, args.liveSheetjsRecalcCli, 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/eval-xlsx-cache-doctor.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/xlsx-cache-doctor-proof-transcript.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://github.com/proompteng/bilig/blob/main/docs/xlsx-cache-doctor-proof-transcript.md', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://www.npmjs.com/package/@bilig/xlsx-formula-recalc', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://www.npmjs.com/package/@bilig/sheetjs-formula-recalc', 'docs/llms.txt')
  requireIncludes(args.index, 'xlsx-cache-doctor', 'docs/index.html')
  requireIncludes(args.xlsxFormulaRecalculationNode, 'xlsx-recalc --demo --json', 'docs/xlsx-formula-recalculation-node.md')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/xlsx-cache-doctor-github-action.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://github.com/proompteng/bilig/blob/main/docs/xlsx-cache-doctor-github-action.md', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/xlsx-recalculation-proof.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/xlsx-recalculation-proof.ts', 'docs/llms.txt')
  requireIncludes(args.llms, 'creates an XLSX workbook, edits inputs, recalculates formulas in Node.js', 'docs/llms.txt')
  for (const required of [
    'title: XLSX Cache Doctor proof transcript',
    'canonical_url: https://proompteng.github.io/bilig/xlsx-cache-doctor-proof-transcript.html',
    'npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor --demo --json',
    '"schemaVersion": "xlsx-cache-doctor.v1"',
    '"staleCachedFormulaCount": 1',
    '"target": "Summary!B2"',
    '"cachedValue": 60000',
    '"literalRecalculatedValue": 72000',
    '"target": "Sheet1!B61"',
    '"cachedValue": 999',
    '"literalRecalculatedValue": 600',
    '"excelParity": "not_proven"',
    'fail-on-stale: "false"',
    '[Evaluate stale XLSX formula caches](eval-xlsx-cache-doctor.md)',
  ]) {
    requireIncludes(args.xlsxCacheDoctorProofTranscript, required, 'docs/xlsx-cache-doctor-proof-transcript.md')
  }
  for (const required of [
    'title: Agent XLSX formula recalculation without LibreOffice',
    'canonical_url: https://proompteng.github.io/bilig/agent-xlsx-formula-recalculation-without-libreoffice.html',
    'curl -fsSLO https://proompteng.github.io/bilig/xlsx-recalculation-proof.ts',
    '"formulasSurvivedXlsxRoundTrip": true',
    'verified: true',
    '[MCP spreadsheet tool server](mcp-workpaper-tool-server.md)',
  ]) {
    requireIncludes(
      args.agentXlsxFormulaRecalculationWithoutLibreOffice,
      required,
      'docs/agent-xlsx-formula-recalculation-without-libreoffice.md',
    )
  }
}
