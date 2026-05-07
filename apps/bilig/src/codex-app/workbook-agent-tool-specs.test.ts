import { describe, expect, it } from 'vitest'
import { WORKBOOK_AGENT_TOOL_NAMES } from '@bilig/agent-api'
import { workbookAgentDynamicToolSpecs } from './workbook-agent-tool-specs.js'

describe('workbook agent dynamic tool specs', () => {
  it('registers core workbook tools with unique names', () => {
    const names = workbookAgentDynamicToolSpecs.map((spec) => spec.name)

    expect(names).toContain(WORKBOOK_AGENT_TOOL_NAMES.getContext)
    expect(names).toContain(WORKBOOK_AGENT_TOOL_NAMES.readWorkbook)
    expect(names).toContain(WORKBOOK_AGENT_TOOL_NAMES.writeRange)
    expect(names).toContain(WORKBOOK_AGENT_TOOL_NAMES.setFormula)
    expect(names).toContain(WORKBOOK_AGENT_TOOL_NAMES.createSheet)
    expect(new Set(names).size).toBe(names.length)
  })
})
