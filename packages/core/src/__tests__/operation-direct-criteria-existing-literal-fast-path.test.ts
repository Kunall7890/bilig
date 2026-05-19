import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

describe('direct criteria existing literal mutations', () => {
  it('updates INDEX/MATCH string criteria edits through the direct existing-literal path', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-direct-criteria-existing-literal',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'D1', 'K5')
    for (let row = 1; row <= 10; row += 1) {
      engine.setCellValue('Sheet1', `A${row + 1}`, `K${row}`)
      engine.setCellValue('Sheet1', `B${row + 1}`, row * 10)
    }
    engine.setCellFormula('Sheet1', 'E1', '=INDEX(B2:B11,MATCH(D1,A2:A11,0))')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const criteriaIndex = engine.workbook.getCellIndex('Sheet1', 'D1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'E1')!
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    const result = engine.tryApplyExistingLiteralCellMutationAt({
      sheetId,
      row: 0,
      col: 3,
      cellIndex: criteriaIndex,
      value: 'K9',
      emitTracked: false,
    })

    expect(result).toEqual({
      firstChangedCellIndex: criteriaIndex,
      secondChangedCellIndex: formulaIndex,
      secondChangedRow: 0,
      secondChangedCol: 4,
      secondChangedNumericValue: 90,
      changedCellCount: 2,
      explicitChangedCount: 1,
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({ tag: ValueTag.String, value: 'K9' })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 90 })
    expect(engine.getPerformanceCounters().directFormulaKernelSyncOnlyRecalcSkips).toBe(1)
    expect(engine.getPerformanceCounters().formulasBound).toBe(0)
    expect(engine.getPerformanceCounters().topoRebuilds).toBe(0)
    expect(tracked).not.toHaveBeenCalled()

    expect(engine.undo()).toBe(true)
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({ tag: ValueTag.String, value: 'K5' })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 50 })
    unsubscribe()
  })
})
