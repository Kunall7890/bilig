import { describe, expect, it } from 'vitest'

import {
  assertWorkPaperMcpDemoOutput,
  buildDemoWorkPaper,
  createWorkPaperMcpDemoOutput,
  createWorkPaperMcpToolServer,
} from '../work-paper-mcp-server.js'

describe('WorkPaper MCP server', () => {
  it('exposes stable tool definitions and structured formula readback', () => {
    const output = createWorkPaperMcpDemoOutput()

    assertWorkPaperMcpDemoOutput(output)
    expect(output.listResponse.result.tools.map((tool) => tool.name)).toEqual(['read_workpaper_summary', 'set_workpaper_input_cell'])
    expect(output.writeResponse.result.structuredContent).toMatchObject({
      editedCell: 'Inputs!B3',
      after: {
        expectedArr: 96000,
        expansionArr: 105600,
        targetGap: 5600,
      },
      checks: {
        formulasPersisted: true,
        restoredMatchesAfter: true,
      },
    })
  })

  it('rejects unknown tool calls instead of returning a misleading success', () => {
    const server = createWorkPaperMcpToolServer(buildDemoWorkPaper())

    expect(() =>
      server.handleJsonRpc({
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/call',
        params: {
          name: 'unknown_tool',
          arguments: {},
        },
      }),
    ).toThrow('Unknown WorkPaper tool')
  })
})
