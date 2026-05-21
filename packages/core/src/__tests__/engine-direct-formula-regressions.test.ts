import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

describe('engine direct formula regressions', () => {
  it('recalculates scalar formulas when an input also has aggregate dependents', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'direct-scalar-with-aggregate-dependent-regression',
      replicaId: 'direct-scalar-with-aggregate-dependent-regression',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 0; row < 5; row += 1) {
      for (let col = 0; col < 5; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row + 1}`, row * 10 + col + 1)
      }
    }
    engine.setCellFormula('Sheet1', 'A6', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'B6', 'A1+B1')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:B2)')

    engine.setCellValue('Sheet1', 'A1', 0)

    expect(engine.getCellValue('Sheet1', 'A6')).toEqual({ tag: ValueTag.Number, value: 104 })
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 25 })

    engine.insertRows('Sheet1', 0, 1)

    expect(engine.getCell('Sheet1', 'B7').formula).toBe('A2+B2')
    expect(engine.getCellValue('Sheet1', 'B7')).toEqual({ tag: ValueTag.Number, value: 2 })
  })
})
