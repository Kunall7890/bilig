import { requireIncludes, requirePackageKeywords } from './check-docs-discovery-core.ts'

export function requirePackageMetadataDiscovery(args: {
  readonly headlessPackageJson: string
  readonly scopedWorkpaperPackageJson: string
  readonly scopedWorkpaperPackageReadme: string
}): void {
  requirePackageKeywords(
    args.headlessPackageJson,
    [
      'agent-tools',
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
      'agent-tools',
      'ai-agents',
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
      'spreadsheet-agent',
      'spreadsheet-automation',
      'spreadsheet-engine',
      'spreadsheet-formula-engine',
      'spreadsheet-formulas',
      'workbook-agent',
      'workbook-api',
      'workpaper',
      'xlsx',
    ],
    'packages/workpaper/package.json',
  )
  requireIncludes(
    args.scopedWorkpaperPackageJson,
    '"description": "Bilig WorkPaper API, CLI evaluator, and MCP server for headless spreadsheet formulas in Node.js services and agents."',
    'packages/workpaper/package.json',
  )
  requireIncludes(
    args.scopedWorkpaperPackageReadme,
    'Bilig WorkPaper is an API, CLI evaluator, and MCP server',
    'packages/workpaper/README.md',
  )
  requireIncludes(
    args.scopedWorkpaperPackageJson,
    '"homepage": "https://proompteng.github.io/bilig/agent-framework-workbook-tools.html"',
    'packages/workpaper/package.json',
  )
}
