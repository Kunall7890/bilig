import { describe, expect, it } from 'vitest'
import {
  formatRangeToolArgsSchema,
  startWorkflowToolArgsSchema,
  type WorkbookAgentStartWorkflowRequest,
} from './workbook-agent-tool-schemas.js'

describe('workbook agent tool schemas', () => {
  it('parses workflow requests with template-specific arguments', () => {
    const request: WorkbookAgentStartWorkflowRequest = startWorkflowToolArgsSchema.parse({
      workflowTemplate: 'searchWorkbookQuery',
      query: 'revenue',
      sheetName: 'Forecast',
      limit: 10,
    })

    expect(request).toEqual({
      workflowTemplate: 'searchWorkbookQuery',
      query: 'revenue',
      sheetName: 'Forecast',
      limit: 10,
    })
  })

  it('requires one target and one formatting payload for format range requests', () => {
    expect(() => formatRangeToolArgsSchema.parse({ range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' } })).toThrow(
      /patch or numberFormat is required/,
    )
    expect(() =>
      formatRangeToolArgsSchema.parse({
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
        selector: { kind: 'currentSelection' },
        patch: { font: { bold: true } },
      }),
    ).toThrow(/Provide exactly one/)

    expect(
      formatRangeToolArgsSchema.parse({
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
        patch: { font: { bold: true } },
      }),
    ).toEqual({
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
      patch: { font: { bold: true } },
    })
  })
})
