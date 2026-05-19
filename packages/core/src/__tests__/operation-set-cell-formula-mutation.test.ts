import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { indexToColumn } from '@bilig/formula'
import { createBatch } from '../replica-state.js'
import { SpreadsheetEngine } from '../engine.js'
import { cellMutationRefToEngineOp, type EngineCellMutationRef } from '../cell-mutations-at.js'
import { applySetCellFormulaMutation } from '../engine/services/operation-set-cell-formula-mutation.js'
import { getOperationService, getReplicaState } from './operation-service-test-helpers.js'

describe('operation set-cell-formula mutations', () => {
  it('keeps set-cell-formula mutation application in a dedicated module', () => {
    expect(applySetCellFormulaMutation).toBeTypeOf('function')
  })

  it('uses a direct scalar delta root for same-topology formula replacements', async () => {
    const downstreamCount = 24
    const engine = new SpreadsheetEngine({ workbookName: 'operation-formula-rewrite-direct-delta-root' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')
    for (let offset = 1; offset <= downstreamCount; offset += 1) {
      const col = 2 + offset
      engine.setCellFormula('Sheet1', `${indexToColumn(col)}1`, `${indexToColumn(col - 1)}1+1`)
    }
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const formulaCellIndex = engine.workbook.getCellIndex('Sheet1', 'C1')
    expect(formulaCellIndex).toBeDefined()
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        cellIndex: formulaCellIndex,
        mutation: { kind: 'setCellFormula', row: 0, col: 2, formula: 'A1*B1' },
      },
    ]
    const batch = createBatch(
      getReplicaState(engine),
      refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref)),
    )
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 0))

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', `${indexToColumn(2 + downstreamCount)}1`)).toEqual({
      tag: ValueTag.Number,
      value: 2 + downstreamCount,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0 })
    expect(engine.getPerformanceCounters().formulasParsed).toBe(1)
    expect(engine.getPerformanceCounters().formulasBound).toBe(0)
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(downstreamCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    const changedIndices = Array.from(tracked.mock.calls.at(-1)?.[0].changedCellIndices ?? [])
    expect(changedIndices[0]).toBe(formulaCellIndex)
    expect(changedIndices).toContain(engine.workbook.getCellIndex('Sheet1', `${indexToColumn(2 + downstreamCount)}1`))
    unsubscribe()
  })

  it('keeps no-listener formula replacement direct deltas off changed-cell materialization', async () => {
    const downstreamCount = 12
    const engine = new SpreadsheetEngine({ workbookName: 'operation-formula-rewrite-no-listener-direct-delta-root' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')
    for (let offset = 1; offset <= downstreamCount; offset += 1) {
      const col = 2 + offset
      engine.setCellFormula('Sheet1', `${indexToColumn(col)}1`, `${indexToColumn(col - 1)}1+1`)
    }

    engine.resetPerformanceCounters()
    const notifyCellValueWritten = vi.spyOn(engine.workbook, 'notifyCellValueWritten')
    engine.setCellFormula('Sheet1', 'C1', 'A1*B1')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', `${indexToColumn(2 + downstreamCount)}1`)).toEqual({
      tag: ValueTag.Number,
      value: 2 + downstreamCount,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      formulasParsed: 1,
      formulasBound: 0,
      directScalarDeltaApplications: downstreamCount,
      directScalarDeltaOnlyRecalcSkips: 1,
      changedCellPayloadsBuilt: 0,
    })
    expect(notifyCellValueWritten).not.toHaveBeenCalled()
    notifyCellValueWritten.mockRestore()
  })

  it('rebinds formulas over existing literal cells through mutation refs', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-mutation-formula-over-literal',
      replicaId: 'a',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'setCellFormula', row: 0, col: 0, formula: '2+2' },
      },
    ]
    const batch = createBatch(
      getReplicaState(engine),
      refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref)),
    )

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 1))

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCell('Sheet1', 'A1').formula).toBe('2+2')
  })

  it('stores invalid formula refs as errors and continues applying the batch', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-local-formula-error-ref', replicaId: 'a' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1*2' },
      },
      {
        sheetId,
        mutation: { kind: 'setCellFormula', row: 0, col: 2, formula: 'SUM(' },
      },
    ]
    const batch = createBatch(
      getReplicaState(engine),
      refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref)),
    )

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 2))

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })
})
