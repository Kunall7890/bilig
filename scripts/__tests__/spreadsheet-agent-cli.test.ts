import { describe, expect, it } from 'vitest'

import { isSpreadsheetAgentCommandName, parseSpreadsheetAgentCliOptions, spreadsheetAgentUsageText } from '../spreadsheet-agent-cli.ts'

describe('spreadsheet agent CLI parser', () => {
  it('parses command options without importing the executable agent', () => {
    expect(
      parseSpreadsheetAgentCliOptions([
        '--range',
        'Sheet1!A1:B2',
        '--server',
        'http://127.0.0.1:4321',
        '--value',
        '{"kind":"number","value":42}',
      ]),
    ).toEqual({
      range: 'Sheet1!A1:B2',
      server: 'http://127.0.0.1:4321',
      value: '{"kind":"number","value":42}',
    })
  })

  it('rejects blank option values before sending workbook mutations', () => {
    expect(() => parseSpreadsheetAgentCliOptions(['--value', '   '])).toThrow('Missing value for --value')
  })

  it('rejects duplicate options instead of silently overriding earlier values', () => {
    expect(() => parseSpreadsheetAgentCliOptions(['--range', 'Sheet1!A1', '--range', 'Sheet1!B2'])).toThrow('Duplicate option: --range')
  })

  it('keeps command validation and usage text side-effect free', () => {
    expect(isSpreadsheetAgentCommandName('write-cell')).toBe(true)
    expect(isSpreadsheetAgentCommandName('write-workbook')).toBe(false)
    expect(spreadsheetAgentUsageText()).toContain('bun scripts/spreadsheet-agent.ts write-cell')
  })
})
