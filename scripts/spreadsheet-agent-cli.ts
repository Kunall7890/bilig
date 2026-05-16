export type SpreadsheetAgentCommandName =
  | 'read-range'
  | 'write-cell'
  | 'write-range'
  | 'set-formula'
  | 'set-formulas'
  | 'set-range-style'
  | 'clear-range-style'
  | 'set-range-number-format'
  | 'clear-range-number-format'
  | 'clear-range'
  | 'create-pivot'
  | 'batch'
  | 'get-metrics'
  | 'export-snapshot'

export type SpreadsheetAgentCliOptions = Record<string, string>

const spreadsheetAgentCommandNames = new Set<string>([
  'read-range',
  'write-cell',
  'write-range',
  'set-formula',
  'set-formulas',
  'set-range-style',
  'clear-range-style',
  'set-range-number-format',
  'clear-range-number-format',
  'clear-range',
  'create-pivot',
  'batch',
  'get-metrics',
  'export-snapshot',
])

export function spreadsheetAgentUsageText(): string {
  return `Usage:
  bun scripts/spreadsheet-agent.ts read-range --range Sheet1!A1:B2 [--server URL] [--document ID] [--replica ID]
  bun scripts/spreadsheet-agent.ts write-cell --sheet Sheet1 --addr A1 --value 42 [--server URL] [--document ID] [--replica ID]
  bun scripts/spreadsheet-agent.ts write-range --range Sheet1!A1:B2 --values '[[1,2],[3,4]]'
  bun scripts/spreadsheet-agent.ts set-formula --sheet Sheet1 --addr B1 --formula 'SUM(A1:A10)'
  bun scripts/spreadsheet-agent.ts set-formulas --range Sheet1!B1:B2 --formulas '[["A1*2"],["A2*2"]]'
  bun scripts/spreadsheet-agent.ts set-range-style --range Sheet1!A1:C3 --patch '{"fill":{"backgroundColor":"#fff59d"},"font":{"family":"Georgia"}}'
  bun scripts/spreadsheet-agent.ts clear-range-style --range Sheet1!A1:C3 [--fields '["backgroundColor"]']
  bun scripts/spreadsheet-agent.ts set-range-number-format --range Sheet1!B2:B10 --format '{"kind":"accounting","currency":"USD","decimals":2}'
  bun scripts/spreadsheet-agent.ts clear-range-number-format --range Sheet1!B2:B10
  bun scripts/spreadsheet-agent.ts clear-range --range Sheet1!A1:B2
  bun scripts/spreadsheet-agent.ts create-pivot --name MyPivot --sheet Sheet1 --addr D1 --source Sheet2!A1:C100 --group '["Category"]' --values '[{"sourceColumn":"Amount","summarizeBy":"sum"}]'
  bun scripts/spreadsheet-agent.ts batch --requests @ops.json [--server URL] [--document ID] [--replica ID]
  bun scripts/spreadsheet-agent.ts get-metrics
  bun scripts/spreadsheet-agent.ts export-snapshot

JSON-heavy flags such as --values, --formulas, --group, --requests, and --value accept:
  inline JSON     '[["a","b"]]'
  @file.json      load JSON from a file
  @-              read JSON from stdin
`
}

export function parseSpreadsheetAgentCliOptions(args: readonly string[]): SpreadsheetAgentCliOptions {
  const options: SpreadsheetAgentCliOptions = {}
  for (let index = 0; index < args.length; index += 1) {
    const token = args[index]
    if (!token?.startsWith('--')) {
      throw new Error(`Unexpected argument: ${token ?? ''}`)
    }
    const key = token.slice(2)
    if (key.length === 0) {
      throw new Error('Unexpected empty option name')
    }
    if (Object.hasOwn(options, key)) {
      throw new Error(`Duplicate option: --${key}`)
    }
    const value = args[index + 1]
    if (value === undefined || value.trim().length === 0 || value.startsWith('--')) {
      throw new Error(`Missing value for --${key}`)
    }
    options[key] = value
    index += 1
  }
  return options
}

export function isSpreadsheetAgentCommandName(value: string): value is SpreadsheetAgentCommandName {
  return spreadsheetAgentCommandNames.has(value)
}
