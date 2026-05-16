import { ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { SpreadsheetEngine } from '../engine.js'

describe('cross-sheet direct aggregate formulas', () => {
  it('updates cross-sheet scalar fanout with direct deltas instead of dirty traversal', async () => {
    const rowCount = 256
    const engine = new SpreadsheetEngine({ workbookName: 'cross-sheet-direct-scalar-fanout' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')

    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Data', `A${row}`, row)
      engine.setCellValue('Data', `B${row}`, row * 10)
      engine.setCellFormula('Summary', `A${row}`, `Data!$A$1+Data!B${row}`)
    }

    const terminalFormulaIndex = engine.workbook.getCellIndex('Summary', `A${rowCount}`)
    expect(engine.state.formulas.get(terminalFormulaIndex)?.directScalar).toMatchObject({ kind: 'binary', operator: '+' })

    engine.resetPerformanceCounters()
    engine.setCellValue('Data', 'A1', 10)

    expect(engine.getCellValue('Summary', `A${rowCount}`)).toEqual({ tag: ValueTag.Number, value: rowCount * 10 + 10 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directScalarDeltaApplications: rowCount,
      directScalarDeltaOnlyRecalcSkips: 1,
    })
  })

  it('updates cross-sheet SUM fanout with direct deltas instead of dirty traversal', async () => {
    const rowCount = 256
    const engine = new SpreadsheetEngine({ workbookName: 'cross-sheet-direct-aggregate-fanout' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')

    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Data', `A${row}`, row)
      engine.setCellFormula('Summary', `A${row}`, `SUM(Data!A1:A${row})`)
    }

    const terminalFormulaIndex = engine.workbook.getCellIndex('Summary', `A${rowCount}`)
    expect(engine.state.formulas.get(terminalFormulaIndex)?.directAggregate).toMatchObject({
      aggregateKind: 'sum',
      sheetName: 'Data',
      rowStart: 0,
      rowEnd: rowCount - 1,
      col: 0,
      colEnd: 0,
    })

    engine.resetPerformanceCounters()
    engine.setCellValue('Data', 'A1', 10)

    expect(engine.getCellValue('Summary', `A${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: (rowCount * (rowCount + 1)) / 2 + 9,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directAggregateDeltaApplications: rowCount,
      directAggregateDeltaOnlyRecalcSkips: 1,
      directAggregateScanEvaluations: 0,
      directAggregateScanCells: 0,
    })
  })
})
