import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import type { EngineExistingNumericCellMutationResult } from '../cell-mutations-at.js'
import { SpreadsheetEngine } from '../engine.js'
import type { OperationDirectAggregateLiteralMutationRequest } from '../engine/services/operation-direct-aggregate-literal-fast-path.js'

function operationHook<Args extends readonly unknown[], Result>(engine: SpreadsheetEngine, name: string): (...args: Args) => Result {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const operations = Reflect.get(runtime, 'operations')
  if (typeof operations !== 'object' || operations === null) {
    throw new TypeError('Expected operation service')
  }
  const hooks = Reflect.get(operations, '__testHooks')
  if (typeof hooks !== 'object' || hooks === null) {
    throw new TypeError('Expected operation hooks')
  }
  const hook = Reflect.get(hooks, name)
  if (typeof hook !== 'function') {
    throw new TypeError(`Expected operation hook ${name}`)
  }
  return (...args: Args): Result => Reflect.apply(hook, hooks, args)
}

describe('existing numeric mutation fast path regressions', () => {
  it('rejects collected direct criteria dependents before aggregate delta application', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'existing-numeric-direct-criteria-not-aggregate-delta',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')
    engine.setCellValue('Data', 'A1', 'Deposit')
    engine.setCellValue('Data', 'B1', 1.235)
    engine.setCellValue('Data', 'A2', 'Deposit')
    engine.setCellValue('Data', 'B2', 0)
    engine.setCellFormula('Summary', 'A1', 'ROUND(SUMIFS(Data!$B$1:$B$2,Data!$A$1:$A$2,"Deposit"),2)')

    const sheet = engine.workbook.getSheet('Data')!
    const inputIndex = engine.workbook.getCellIndex('Data', 'B1')!
    const fastPath = operationHook<[OperationDirectAggregateLiteralMutationRequest], EngineExistingNumericCellMutationResult | null>(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )
    const result = fastPath({
      existingIndex: inputIndex,
      sheetId: sheet.id,
      sheetName: 'Data',
      row: 0,
      col: 1,
      value: 1.236,
      delta: 0.001,
      emitTracked: false,
    })

    expect(result).toBeNull()
    expect(engine.getCellValue('Data', 'B1')).toEqual({ tag: ValueTag.Number, value: 1.235 })
    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 1.24 })
  })

  it('applies terminal aggregate deltas for valid off-column collected dependents', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'existing-numeric-terminal-aggregate-delta',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A2)')

    const sheet = engine.workbook.getSheet('Sheet1')!
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const fastPath = operationHook<[OperationDirectAggregateLiteralMutationRequest], EngineExistingNumericCellMutationResult | null>(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )
    const result = fastPath({
      existingIndex: inputIndex,
      sheetId: sheet.id,
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      value: 5,
      delta: 4,
      emitTracked: false,
    })

    expect(result).toMatchObject({
      firstChangedCellIndex: inputIndex,
      changedCellCount: 2,
      explicitChangedCount: 1,
      secondChangedNumericValue: 7,
      secondChangedRow: 0,
      secondChangedCol: 2,
    })
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 7 })
  })

  it('recalculates every aggregate touching a range-entity input after text replacement', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'existing-numeric-overlapping-aggregate-text-replacement',
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
    engine.insertColumns('Sheet1', 0, 1)
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:B2)')

    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 105 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'C6')).toEqual({ tag: ValueTag.Number, value: 13 })

    engine.setCellValue('Sheet1', 'B2', 'north')

    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 94 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'C6')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

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
