import { requireIncludes } from './check-docs-discovery-core.ts'

export function requireXlsxRecalcPublicDiscovery(args: {
  readonly agentXlsxFormulaRecalculationWithoutLibreOffice: string
  readonly headlessReadme: string
  readonly index: string
  readonly liveSheetjsRecalcCli: string
  readonly llms: string
  readonly readme: string
  readonly workbookCompatibilityReport: string
  readonly workbookCompatibilityReportJson: string
  readonly workbookCompatibilityReportTranscript: string
  readonly xlsxFormulaRecalculationNode: string
  readonly xlsxCacheDoctorProofTranscript: string
  readonly xlsxRecalcCli: string
}): void {
  requireIncludes(args.readme, 'examples/xlsx-recalculation-node', 'README.md')
  requireIncludes(args.readme, 'examples/recalc-bridge-workflows', 'README.md')
  requireIncludes(args.readme, 'docs/workbook-compatibility-report.md', 'README.md')
  requireIncludes(args.readme, 'bilig-evaluate --door workbook-compatibility --json', 'README.md')
  requireIncludes(args.readme, 'workbook-compatibility-report workbook.xlsx --json', 'README.md')
  requireIncludes(args.readme, 'docs/xlsx-formula-recalculation-node.md', 'README.md')
  requireIncludes(args.readme, args.xlsxRecalcCli, 'README.md')
  requireIncludes(args.readme, args.liveSheetjsRecalcCli, 'README.md')
  requireIncludes(args.headlessReadme, 'examples/xlsx-recalculation-node', 'packages/headless/README.md')
  requireIncludes(args.headlessReadme, 'docs/xlsx-formula-recalculation-node.md', 'packages/headless/README.md')
  requireIncludes(args.headlessReadme, 'file boundary is the product contract', 'packages/headless/README.md')
  requireIncludes(args.headlessReadme, 'workbook-compatibility-report.md', 'packages/headless/README.md')
  requireIncludes(args.index, 'examples/xlsx-recalculation-node', 'docs/index.html')
  requireIncludes(args.index, 'Use saved-file packages only', 'docs/index.html')
  requireIncludes(args.index, './workbook-compatibility-report.html', 'docs/index.html')
  requireIncludes(args.index, './workbook-compatibility-report-transcript.html', 'docs/index.html')
  requireIncludes(args.index, './agent-xlsx-risk-preflight.html', 'docs/index.html')
  requireIncludes(args.index, './xlsx-formula-recalculation-node.html', 'docs/index.html')
  requireIncludes(args.index, './sheetjs-formula-result-not-updating-node.html', 'docs/index.html')
  requireIncludes(args.llms, 'https://github.com/proompteng/bilig/tree/main/examples/xlsx-recalculation-node', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/workbook-compatibility-report.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://github.com/proompteng/bilig/blob/main/docs/workbook-compatibility-report.md', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/workbook-compatibility-report-transcript.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/workbook-compatibility-report.json', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/agent-xlsx-risk-preflight.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://github.com/proompteng/bilig/blob/main/docs/agent-xlsx-risk-preflight.md', 'docs/llms.txt')
  requireIncludes(args.llms, 'pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight', 'docs/llms.txt')
  requireIncludes(args.llms, 'bilig-evaluate --door workbook-compatibility --json', 'docs/llms.txt')
  requireIncludes(args.llms, 'workbook-compatibility-report workbook.xlsx --json', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/xlsx-formula-recalculation-node.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://github.com/proompteng/bilig/blob/main/docs/xlsx-formula-recalculation-node.md', 'docs/llms.txt')
  requireIncludes(args.llms, args.xlsxRecalcCli, 'docs/llms.txt')
  requireIncludes(args.llms, args.liveSheetjsRecalcCli, 'docs/llms.txt')
  requireIncludes(args.llms, 'https://www.npmjs.com/package/@bilig/xlsx-formula-recalc', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://www.npmjs.com/package/@bilig/sheetjs-formula-recalc', 'docs/llms.txt')
  requireIncludes(args.xlsxFormulaRecalculationNode, 'xlsx-recalc --demo --json', 'docs/xlsx-formula-recalculation-node.md')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/xlsx-recalculation-proof.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/xlsx-recalculation-proof.ts', 'docs/llms.txt')
  requireIncludes(args.llms, 'creates an XLSX workbook, edits inputs, recalculates formulas in Node.js', 'docs/llms.txt')
  for (const required of [
    'title: Workbook Compatibility Report',
    'canonical_url: https://proompteng.github.io/bilig/workbook-compatibility-report.html',
    'workbook-compatibility-report workbook.xlsx --json',
    'bilig-evaluate --door workbook-compatibility --json',
    '"schemaVersion": "bilig-workbook-compatibility-report.v1"',
    '"unsupportedFunctions": [{ "name": "CUBEVALUE", "count": 1 }]',
    '"volatileFunctions": [{ "name": "NOW", "count": 1 }]',
    '"excelParity": "not_proven"',
    'cacheInspection.uninspectedFormulaCellCount',
    'limited inspection is visible and raises risk',
    'It is not an Excel compatibility certification.',
    'compatibilityScore',
  ]) {
    requireIncludes(args.workbookCompatibilityReport, required, 'docs/workbook-compatibility-report.md')
  }
  for (const required of [
    'title: Workbook Compatibility Report transcript',
    'canonical_url: https://proompteng.github.io/bilig/workbook-compatibility-report-transcript.html',
    'workbook-compatibility-report --demo --json',
    'bilig-evaluate --door workbook-compatibility --json',
    '"schemaVersion": "bilig-workbook-compatibility-report.v1"',
    '"door": "workbook-compatibility"',
    '"riskLevel": "high"',
    'compatibilityScore',
    'excelCompatibilityPercent',
  ]) {
    requireIncludes(args.workbookCompatibilityReportTranscript, required, 'docs/workbook-compatibility-report-transcript.md')
  }
  const parsedCompatibilityReport: unknown = JSON.parse(args.workbookCompatibilityReportJson)
  if (
    typeof parsedCompatibilityReport !== 'object' ||
    parsedCompatibilityReport === null ||
    Array.isArray(parsedCompatibilityReport) ||
    Reflect.get(parsedCompatibilityReport, 'schemaVersion') !== 'bilig-workbook-compatibility-report.v1' ||
    Reflect.get(parsedCompatibilityReport, 'verified') !== true ||
    Reflect.get(parsedCompatibilityReport, 'excelParity') !== 'not_proven'
  ) {
    throw new Error('docs/workbook-compatibility-report.json must contain the checked workbook compatibility report proof')
  }
  const serializedCompatibilityReport = JSON.stringify(parsedCompatibilityReport)
  requireIncludes(serializedCompatibilityReport, '"CUBEVALUE"', 'docs/workbook-compatibility-report.json')
  requireIncludes(serializedCompatibilityReport, '"NOW"', 'docs/workbook-compatibility-report.json')
  requireIncludes(
    serializedCompatibilityReport,
    '"It is not an Excel compatibility certification."',
    'docs/workbook-compatibility-report.json',
  )
  requireIncludes(serializedCompatibilityReport, '"excelParity":"not_proven"', 'docs/workbook-compatibility-report.json')
  if (/compatibilityScore|excelCompatibilityPercent/iu.test(serializedCompatibilityReport)) {
    throw new Error('docs/workbook-compatibility-report.json must not include compatibility score fields')
  }
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
