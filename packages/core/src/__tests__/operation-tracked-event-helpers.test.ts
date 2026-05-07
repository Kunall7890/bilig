import { describe, expect, it, vi } from 'vitest'
import { createInitialRecalcMetrics } from '../engine/runtime-state.js'
import { createOperationBatchMetrics } from '../engine/services/operation-batch-metrics.js'
import { emitOperationTrackedCellsBatch } from '../engine/services/operation-tracked-event-helpers.js'

describe('operation tracked event helpers', () => {
  it('emits a typed cells batch event with empty structural invalidations', () => {
    const emitTracked = vi.fn()
    const changedCellIndices = Uint32Array.of(1, 3)
    const metrics = createOperationBatchMetrics({
      previousMetrics: createInitialRecalcMetrics(),
      didRunRecalc: false,
      directFormulaMetrics: { wasmFormulaCount: 0, jsFormulaCount: 0 },
      changedInputCount: 1,
      formulaChangedCount: 0,
      compileMs: 0,
    })

    emitOperationTrackedCellsBatch({
      events: { emitTracked },
      changedCellIndices,
      metrics,
      explicitChangedCount: 2,
    })

    expect(emitTracked).toHaveBeenCalledWith({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices,
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics,
      explicitChangedCount: 2,
    })
  })

  it('defaults explicit changed count to one for single-input fast paths', () => {
    const emitTracked = vi.fn()
    const metrics = createInitialRecalcMetrics()

    emitOperationTrackedCellsBatch({
      events: { emitTracked },
      changedCellIndices: Uint32Array.of(4),
      metrics,
    })

    expect(emitTracked.mock.calls[0]?.[0]).toMatchObject({ explicitChangedCount: 1 })
  })
})
