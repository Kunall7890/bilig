import { requireIncludes } from './check-docs-discovery-core.ts'

export function requireXlsxRecalcPublicDiscovery(args: {
  readonly agentXlsxFormulaRecalculationWithoutLibreOffice: string
  readonly headlessReadme: string
  readonly index: string
  readonly liveSheetjsRecalcCli: string
  readonly llms: string
  readonly readme: string
  readonly xlsxFormulaRecalculationNode: string
  readonly xlsxRecalcCli: string
}): void {
  requireIncludes(args.readme, 'examples/xlsx-recalculation-node', 'README.md')
  requireIncludes(args.readme, 'examples/recalc-bridge-workflows', 'README.md')
  requireIncludes(args.readme, 'docs/xlsx-formula-recalculation-node.md', 'README.md')
  requireIncludes(args.readme, args.xlsxRecalcCli, 'README.md')
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
  requireIncludes(args.index, 'SheetJS-named package before moving to a full workbook runtime', 'docs/index.html')
  requireIncludes(args.index, 'xlsx-populate, and ExcelJS with one workbook', 'docs/index.html')
  requireIncludes(args.index, './xlsx-formula-recalculation-node.html', 'docs/index.html')
  requireIncludes(args.index, './xlsx-recalculation-proof.html', 'docs/index.html')
  requireIncludes(args.index, './agent-xlsx-formula-recalculation-without-libreoffice.html', 'docs/index.html')
  requireIncludes(args.index, './excel-file-calculation-engine-node.html', 'docs/index.html')
  requireIncludes(args.index, './exceljs-shared-formula-recalculation-node.html', 'docs/index.html')
  requireIncludes(args.index, './xlsx-template-formula-recalculation-node.html', 'docs/index.html')
  requireIncludes(args.index, './xlsx-populate-formula-result-node.html', 'docs/index.html')
  requireIncludes(args.index, './sheetjs-formula-result-not-updating-node.html', 'docs/index.html')
  requireIncludes(args.llms, 'https://github.com/proompteng/bilig/tree/main/examples/xlsx-recalculation-node', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/xlsx-formula-recalculation-node.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://github.com/proompteng/bilig/blob/main/docs/xlsx-formula-recalculation-node.md', 'docs/llms.txt')
  requireIncludes(args.llms, args.xlsxRecalcCli, 'docs/llms.txt')
  requireIncludes(args.llms, args.liveSheetjsRecalcCli, 'docs/llms.txt')
  requireIncludes(args.llms, 'https://www.npmjs.com/package/@bilig/xlsx-formula-recalc', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://www.npmjs.com/package/@bilig/sheetjs-formula-recalc', 'docs/llms.txt')
  requireIncludes(args.index, 'xlsx-recalc --demo --json', 'docs/index.html')
  requireIncludes(args.index, 'sheetjs-recalc --demo --json', 'docs/index.html')
  requireIncludes(args.index, 'sheetjs-formula-result-not-updating-node.html', 'docs/index.html')
  requireIncludes(args.xlsxFormulaRecalculationNode, 'xlsx-recalc --demo --json', 'docs/xlsx-formula-recalculation-node.md')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/xlsx-recalculation-proof.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/xlsx-recalculation-proof.ts', 'docs/llms.txt')
  requireIncludes(args.llms, 'creates an XLSX workbook, edits inputs, recalculates formulas in Node.js', 'docs/llms.txt')
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
