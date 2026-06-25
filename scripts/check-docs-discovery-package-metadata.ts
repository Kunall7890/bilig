import { requireIncludes, requirePackageKeywords } from './check-docs-discovery-core.ts'

export function requirePackageMetadataDiscovery(args: {
  readonly headlessPackageJson: string
  readonly scopedWorkpaperPackageJson: string
  readonly scopedWorkpaperPackageReadme: string
}): void {
  requirePackageKeywords(
    args.headlessPackageJson,
    [
      'bilig',
      'formula-engine',
      'formula-recalculation',
      'formula-workbook',
      'hyperformula',
      'mcp',
      'mcp-server',
      'mcp-tools',
      'node',
      'node-services',
      'server-side-formula-engine',
      'server-side-formulas',
      'tool-integration',
      'typescript',
      'workbook',
      'workbook-api',
      'workbook-formulas',
      'workbook-runtime',
      'workpaper',
      'workpaper-json',
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
      'bilig',
      'bilig-workpaper',
      'formula-engine',
      'formula-recalculation',
      'formula-workbook',
      'mcp',
      'mcp-server',
      'mcp-tools',
      'model-context-protocol',
      'node',
      'server-side-formula-engine',
      'server-side-formulas',
      'tool-integration',
      'typescript',
      'workbook',
      'workbook-api',
      'workbook-formulas',
      'workbook-runtime',
      'workpaper',
      'workpaper-json',
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
  requireIncludes(args.scopedWorkpaperPackageJson, '"homepage": "https://proompteng.github.io/bilig/"', 'packages/workpaper/package.json')
}
