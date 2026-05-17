import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { indexToColumn } from '@bilig/formula'
import { createBatch } from '../replica-state.js'
import { SpreadsheetEngine } from '../engine.js'
import { applyBatchSetCellFormulaOp } from '../engine/services/operation-batch-cell-formula-mutations.js'
import { getOperationService, getReplicaState } from './operation-service-test-helpers.js'

describe('operation batch cell formula mutations', () => {
  it('keeps batch set-formula application in a dedicated module', () => {
    expect(applyBatchSetCellFormulaOp).toBeTypeOf('function')
  })

  it('uses a direct scalar delta root for same-topology generic batch formula replacements', async () => {
    const downstreamCount = 18
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-formula-rewrite-direct-delta-root' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')
    for (let offset = 1; offset <= downstreamCount; offset += 1) {
      const col = 2 + offset
      engine.setCellFormula('Sheet1', `${indexToColumn(col)}1`, `${indexToColumn(col - 1)}1+1`)
    }

    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellFormula', sheetName: 'Sheet1', address: 'C1', formula: 'A1*B1' }])

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', `${indexToColumn(2 + downstreamCount)}1`)).toEqual({
      tag: ValueTag.Number,
      value: 6 + downstreamCount,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      formulasParsed: 1,
      formulasBound: 0,
      directScalarDeltaApplications: downstreamCount,
      directScalarDeltaOnlyRecalcSkips: 1,
    })
  })

  it('binds generic batch formulas over existing literal cells and updates dependents', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-formula-over-literal' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellValue('Sheet1', 'B1', 4)
    engine.setCellValue('Sheet1', 'C1', 9)
    engine.setCellFormula('Sheet1', 'D1', 'C1+1')

    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellFormula', sheetName: 'Sheet1', address: 'C1', formula: 'A1*B1' }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCell('Sheet1', 'C1').formula).toBe('A1*B1')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 13 })
  })

  it('stores invalid generic batch formulas as errors and continues later ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-formula-error-continues' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)

    const batch = createBatch(getReplicaState(engine), [
      { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'B1', formula: 'SUM(' },
      { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'C1', formula: 'A1*2' },
    ])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('keeps aggregate dependents current when generic batch formulas enter their range', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-batch-formula-aggregate-dependent',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')

    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellFormula', sheetName: 'Sheet1', address: 'A2', formula: 'A1+A3' }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })
  })
})
