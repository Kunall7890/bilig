import { describe, expect, it } from 'vitest'

import { createWorkPaperAgentTools, workPaperAgentToolSchemas } from '../agent-tools.js'
import { WorkPaper } from '../work-paper.js'

describe('WorkPaper agent tools', () => {
  it('reads a sheet-qualified range with values and serialized formula contracts', () => {
    const workbook = buildAgentWorkbook()
    try {
      const tools = createWorkPaperAgentTools(workbook)
      const readback = tools.readWorkPaperRange({ range: 'Summary!A1:B5' })

      expect(readback.range).toBe('Summary!A1:B5')
      expect(readback.values[1]?.[1]).toEqual({ tag: 1, value: 5 })
      expect(readback.values[2]?.[1]).toEqual({ tag: 1, value: 60000 })
      expect(readback.serialized[2]?.[1]).toBe('=B2*Inputs!B4')
      expect(workPaperAgentToolSchemas.setWorkPaperCell.inputSchema.required).toEqual(['value'])
    } finally {
      workbook.dispose()
    }
  })

  it('keeps writes disabled unless the host explicitly opts in', () => {
    const workbook = buildAgentWorkbook()
    try {
      const tools = createWorkPaperAgentTools(workbook)

      expect(() =>
        tools.setWorkPaperCell({
          target: 'Inputs!B3',
          value: 0.4,
        }),
      ).toThrow('read-only by default')
    } finally {
      workbook.dispose()
    }
  })

  it('returns recalculation and restore proof for allowed writes', () => {
    const workbook = buildAgentWorkbook()
    try {
      const tools = createWorkPaperAgentTools(workbook, {
        allowedInputSheets: ['Inputs'],
        trackedRanges: ['Summary!A1:B5'],
        writable: true,
      })

      const result = tools.setWorkPaperCell({
        sheetName: 'Inputs',
        address: 'B3',
        value: 0.4,
      })

      expect(result.editedCell).toBe('Inputs!B3')
      expect(result.changes.map((change) => change.kind)).toContain('cell')
      expect(result.trackedRanges['Summary!A1:B5']?.before.values[2]?.[1]).toEqual({ tag: 1, value: 60000 })
      expect(result.trackedRanges['Summary!A1:B5']?.after.values[2]?.[1]).toEqual({ tag: 1, value: 96000 })
      expect(result.trackedRanges['Summary!A1:B5']?.restored.values[2]?.[1]).toEqual({ tag: 1, value: 96000 })
      expect(result.checks).toMatchObject({
        previousValue: 0.25,
        newValue: 0.4,
        cellValueChanged: true,
        trackedRangesChanged: true,
        formulasPersisted: true,
        restoredMatchesAfter: true,
      })
      expect(result.checks.serializedBytes).toBeGreaterThan(0)
    } finally {
      workbook.dispose()
    }
  })

  it('rejects writes outside the configured input sheets', () => {
    const workbook = buildAgentWorkbook()
    try {
      const tools = createWorkPaperAgentTools(workbook, {
        allowedInputSheets: ['Inputs'],
        writable: true,
      })

      expect(() =>
        tools.setWorkPaperCell({
          target: 'Summary!B2',
          value: 10,
        }),
      ).toThrow('outside allowed input sheets')
    } finally {
      workbook.dispose()
    }
  })
})

function buildAgentWorkbook() {
  return WorkPaper.buildFromSheets({
    Inputs: [
      ['Metric', 'Value'],
      ['Qualified opportunities', 20],
      ['Win rate', 0.25],
      ['Average ARR', 12000],
      ['Expansion multiplier', 1.1],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Expected customers', '=Inputs!B2*Inputs!B3'],
      ['Expected ARR', '=B2*Inputs!B4'],
      ['Expansion ARR', '=B3*Inputs!B5'],
      ['Target gap', '=B4-100000'],
    ],
  })
}
