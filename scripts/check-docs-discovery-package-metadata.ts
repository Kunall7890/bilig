import { requireIncludes, requirePackageKeywords } from './check-docs-discovery-core.ts'

export function requirePackageMetadataDiscovery(args: {
  readonly headlessPackageJson: string
  readonly scopedWorkpaperPackageJson: string
  readonly scopedWorkpaperPackageReadme: string
}): void {
  requirePackageKeywords(
    args.headlessPackageJson,
    [
      'excel',
      'excel-formulas',
      'formula-recalculation',
      'formula-engine',
      'headless-spreadsheet',
      'hyperformula',
      'mcp',
      'mcp-server',
      'node',
      'spreadsheet-automation',
      'spreadsheet-engine',
      'spreadsheet-formulas',
      'typescript',
      'workbook-api',
      'workpaper',
      'xlsx',
    ],
    'packages/headless/package.json',
  )
  requireIncludes(
    args.headlessPackageJson,
    '"homepage": "https://proompteng.github.io/bilig/try-bilig-headless-in-node.html"',
    'packages/headless/package.json',
  )
  requirePackageKeywords(
    args.scopedWorkpaperPackageJson,
    [
      'bilig-workpaper',
      'excel-formulas',
      'formula-engine',
      'headless-spreadsheet',
      'mcp',
      'mcp-server',
      'model-context-protocol',
      'node-spreadsheet',
      'node-spreadsheet-formulas',
      'server-side-formula-engine',
      'server-side-spreadsheet',
      'spreadsheet-automation',
      'spreadsheet-engine',
      'spreadsheet-formula-engine',
      'spreadsheet-formulas',
      'workbook-api',
      'workpaper',
      'xlsx',
    ],
    'packages/workpaper/package.json',
  )
  requireIncludes(
    args.scopedWorkpaperPackageJson,
    '"description": "Run workbook-shaped business rules in Node services: edit inputs, recalculate formulas, read outputs, and save WorkPaper JSON."',
    'packages/workpaper/package.json',
  )
  requireIncludes(
    args.scopedWorkpaperPackageReadme,
    'Bilig WorkPaper is an API, CLI evaluator, and optional MCP server',
    'packages/workpaper/README.md',
  )
  requireIncludes(
    args.scopedWorkpaperPackageJson,
    '"homepage": "https://proompteng.github.io/bilig/agent-framework-workbook-tools.html"',
    'packages/workpaper/package.json',
  )
}
