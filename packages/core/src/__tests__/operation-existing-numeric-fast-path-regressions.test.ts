import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

describe('existing numeric mutation fast path regressions', () => {
  it('falls back when a numeric input has both aggregate and scalar formula dependents', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'existing-numeric-mixed-dependent-fallback',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 0; row < 5; row += 1) {
      engine.setCellValue('Sheet1', `A${row + 1}`, row * 10 + 1)
    }
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!

    const result = engine.tryApplyExistingNumericCellMutationAt({
      sheetId,
      row: 0,
      col: 0,
      cellIndex: inputIndex,
      value: 0,
      emitTracked: false,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 1,
    })

    expect(result).toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Number, value: 105 })
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 3 })

    engine.setCellValue('Sheet1', 'A1', 0)

    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Number, value: 104 })
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 2 })
  })
})
