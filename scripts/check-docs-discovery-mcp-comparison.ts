import { requireIncludes } from './check-docs-discovery-core.ts'

export function requireSpreadsheetMcpServerComparisonDiscovery(input: { readonly spreadsheetMcpServerComparison: string }): void {
  const { spreadsheetMcpServerComparison } = input
  for (const required of [
    '## Named Public Alternatives',
    'https://github.com/henilcalagiya/google-sheets-mcp',
    'https://github.com/dream-num/univer-mcp',
    'https://github.com/GRID-is/claude-mcp',
    'https://github.com/mort-lab/excel-mcp',
    'https://github.com/negokaz/excel-mcp-server',
    'https://github.com/haris-musa/excel-mcp-server',
    'https://mcpservers.org/servers/iheldan/sheetforge-mcp',
    'https://cdn.cdata.com/help/RXK/mcp/pg_excelformula.htm',
    'A file library can preserve formulas without recalculating fresh results in Node',
    'Do not pitch Bilig as "another Google\nSheets MCP server"',
    'A long-running SheetJS issue asks\nwhether a formula value can be refreshed after changing an input cell',
    'ExcelJS discussion describes JSON-driven workbook edits where shared formulas',
    '## Formula Boundary Checklist',
    'The formula was authored; the result still needs a calculation engine.',
    'The value may be a cached value from the file unless recalculation is documented.',
    'The source must stay available and the engine/configuration must be part of proof.',
    'The agent can cite the edited input and dependent readback from the same run.',
  ] as const) {
    requireIncludes(spreadsheetMcpServerComparison, required, 'docs/spreadsheet-mcp-server-comparison.md')
  }
}
