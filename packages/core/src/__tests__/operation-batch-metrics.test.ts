import { describe, expect, it } from 'vitest'
import type { RecalcMetrics } from '@bilig/protocol'
import { createOperationBatchMetrics } from '../engine/services/operation-batch-metrics.js'

const previousMetrics: RecalcMetrics = {
  batchId: 10,
  changedInputCount: 99,
  dirtyFormulaCount: 7,
  wasmFormulaCount: 11,
  jsFormulaCount: 13,
  rangeNodeVisits: 17,
  recalcMs: 19,
  compileMs: 23,
}

describe('operation batch metrics', () => {
  it('resets recalculation counters when no recalculation ran', () => {
    expect(
      createOperationBatchMetrics({
        previousMetrics,
        didRunRecalc: false,
        directFormulaMetrics: { wasmFormulaCount: 2, jsFormulaCount: 3 },
        changedInputCount: 4,
        formulaChangedCount: 5,
        compileMs: 6,
      }),
    ).toEqual({
      batchId: 11,
      changedInputCount: 9,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 2,
      jsFormulaCount: 3,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 6,
    })
  })

  it('adds post-recalc direct formula counts when a recalculation ran', () => {
    expect(
      createOperationBatchMetrics({
        previousMetrics,
        didRunRecalc: true,
        directFormulaMetrics: { wasmFormulaCount: 2, jsFormulaCount: 3 },
        changedInputCount: 4,
        formulaChangedCount: 5,
        compileMs: 6,
      }),
    ).toEqual({
      batchId: 11,
      changedInputCount: 9,
      dirtyFormulaCount: 7,
      wasmFormulaCount: 13,
      jsFormulaCount: 16,
      rangeNodeVisits: 17,
      recalcMs: 19,
      compileMs: 6,
    })
  })
})
