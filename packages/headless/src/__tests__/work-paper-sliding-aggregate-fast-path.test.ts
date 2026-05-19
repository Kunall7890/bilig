import { ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function buildSlidingAggregateSheet(rowCount: number, window: number): readonly (readonly (number | string)[])[] {
  return Array.from({ length: rowCount }, (_, row) => {
    const rowNumber = row + 1
    const endRow = Math.min(rowCount, rowNumber + window - 1)
    return [rowNumber, `=SUM(A${rowNumber}:A${endRow})`]
  })
}

describe('WorkPaper sliding aggregate fast path', () => {
  it('keeps benchmark-shaped public numeric edits on the compact direct aggregate path', () => {
    const rowCount = 1_500
    const workbook = WorkPaper.buildFromSheets({
      Bench: buildSlidingAggregateSheet(rowCount, 128),
    })
    const sheetId = workbook.getSheetId('Bench')!

    workbook.resetPerformanceCounters()
    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 99)

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual(['Bench!A1', 'Bench!B1'])
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({ tag: ValueTag.Number, value: 8354 })
    expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 1))).toEqual({ tag: ValueTag.Number, value: rowCount })
    expect(workbook.getStats().lastMetrics).toMatchObject({
      changedInputCount: 1,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      changedCellPayloadsBuilt: 0,
      directAggregateDeltaApplications: 1,
      directAggregateDeltaOnlyRecalcSkips: 1,
      regionQueryIndexBuilds: 0,
    })
  })
})
