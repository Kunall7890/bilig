import { requireIncludes } from './check-docs-discovery-core.ts'

export function requireSpreadsheetMcpServerComparisonDiscovery(input: { readonly spreadsheetMcpServerComparison: string }): void {
  const { spreadsheetMcpServerComparison } = input
  for (const required of [
    '## Named Public Alternatives',
    'https://cellium.dev/',
    'https://xlsx-for-ai.dev/',
    'https://www.quadratichq.com/ai/mcp/excel',
    'https://github.com/henilcalagiya/google-sheets-mcp',
    'https://github.com/dream-num/univer-mcp',
    'https://github.com/GRID-is/claude-mcp',
    'https://github.com/mort-lab/excel-mcp',
    'https://github.com/negokaz/excel-mcp-server',
    'https://github.com/haris-musa/excel-mcp-server',
    'https://mcpservers.org/servers/iheldan/sheetforge-mcp',
    'https://cdn.cdata.com/help/RXK/mcp/pg_excelformula.htm',
    'Cellium is a live Excel-control layer with API-key and session-pairing\nboundaries.',
    'Quadratic is a hosted spreadsheet workspace after import.',
    'xlsx-for-ai\nis a hosted API/npm MCP path for Excel-file operations.',
    'The MCP client is not the proof.',
    'what the spreadsheet tool proves after a write.',
    'A file library can preserve formulas without recalculating fresh results in Node',
    'Do not pitch Bilig as "another Google\nSheets MCP server," "another Excel file editor," or "a hosted Excel control\nlayer."',
    'Openpyxl-backed MCP servers can write a\nformula string and read cached workbook values, but openpyxl itself does not\ncalculate formulas.',
    'A long-running SheetJS issue asks\nwhether a formula value can be refreshed after changing an input cell',
    'ExcelJS discussion describes JSON-driven workbook edits where shared formulas',
    '## Formula Boundary Checklist',
    'The live service or workspace performed the operation; the account/session/configuration is part of proof.',
    'The formula was authored; the result still needs a calculation engine.',
    'The value may be a cached value from the file unless recalculation is documented.',
    'The source must stay available and the engine/configuration must be part of proof.',
    'The agent can cite the edited input and dependent readback from the same run.',
  ] as const) {
    requireIncludes(spreadsheetMcpServerComparison, required, 'docs/spreadsheet-mcp-server-comparison.md')
  }
}
