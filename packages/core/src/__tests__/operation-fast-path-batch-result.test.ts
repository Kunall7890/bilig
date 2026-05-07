import { describe, expect, it, vi } from 'vitest'
import type { EngineChangedCell } from '@bilig/protocol'
import { createInitialRecalcMetrics } from '../engine/runtime-state.js'
import {
  emitCellMutationFastPathBatchResult,
  type OperationFastPathBatchResultState,
} from '../engine/services/operation-fast-path-batch-result.js'
import { createBatch, createReplicaState } from '../replica-state.js'

function makeState(): {
  readonly state: OperationFastPathBatchResultState
  readonly emit: OperationFastPathBatchResultState['events']['emit']
  readonly emitTracked: OperationFastPathBatchResultState['events']['emitTracked']
  readonly send: (batch: ReturnType<typeof createBatch>) => void
  readonly setLastMetrics: OperationFastPathBatchResultState['setLastMetrics']
} {
  const emit: OperationFastPathBatchResultState['events']['emit'] = vi.fn()
  const emitTracked: OperationFastPathBatchResultState['events']['emitTracked'] = vi.fn()
  const send = vi.fn((_: ReturnType<typeof createBatch>) => {})
  const setLastMetrics: OperationFastPathBatchResultState['setLastMetrics'] = vi.fn()
  const previousMetrics = {
    ...createInitialRecalcMetrics(),
    batchId: 7,
    dirtyFormulaCount: 12,
    wasmFormulaCount: 4,
    jsFormulaCount: 3,
    rangeNodeVisits: 99,
    recalcMs: 13,
    compileMs: 5,
  }
  return {
    state: {
      events: {
        emit,
        emitTracked,
      },
      workbook: {
        getQualifiedAddress: (cellIndex) => `Sheet1!R${String(cellIndex)}`,
      },
      getLastMetrics: () => previousMetrics,
      setLastMetrics,
      getSyncClientConnection: () => ({
        send,
        disconnect: vi.fn(),
      }),
    },
    emit,
    emitTracked,
    send,
    setLastMetrics,
  }
}

describe('operation fast path batch result helper', () => {
  it('publishes zero-recalc metrics, general events, tracked events, and sync batches', () => {
    const { state, emit, emitTracked, send, setLastMetrics } = makeState()
    const batch = createBatch(createReplicaState('fast-path-test'), [])
    const emitBatch = vi.fn((_: typeof batch) => {})
    const captureChangedCells = vi.fn((_: readonly number[] | Uint32Array): readonly EngineChangedCell[] => [])

    emitCellMutationFastPathBatchResult({
      state,
      changed: Uint32Array.of(2, 3),
      changedInputCount: 2,
      explicitChangedCount: 1,
      hasGeneralEventListeners: true,
      hasTrackedEventListeners: true,
      hasWatchedCellListeners: false,
      captureChangedCells,
      batch,
      emitBatch,
    })

    expect(setLastMetrics).toHaveBeenCalledWith(
      expect.objectContaining({
        batchId: 8,
        changedInputCount: 2,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      }),
    )
    expect(captureChangedCells).toHaveBeenCalledWith(Uint32Array.of(2, 3))
    expect(emit).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        invalidation: 'cells',
        explicitChangedCount: 1,
      }),
      Uint32Array.of(2, 3),
      expect.any(Function),
    )
    expect(emitTracked).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        invalidation: 'cells',
        explicitChangedCount: 1,
      }),
    )
    expect(send).toHaveBeenCalledWith(batch)
    expect(emitBatch).toHaveBeenCalledWith(batch)
  })

  it('notifies watched listeners without materializing general changed cells', () => {
    const { state, emit, emitTracked } = makeState()
    const captureChangedCells = vi.fn((_: readonly number[] | Uint32Array): readonly EngineChangedCell[] => [])

    emitCellMutationFastPathBatchResult({
      state,
      changed: Uint32Array.of(4),
      changedInputCount: 1,
      explicitChangedCount: 1,
      hasGeneralEventListeners: false,
      hasTrackedEventListeners: false,
      hasWatchedCellListeners: true,
      captureChangedCells,
      batch: null,
      emitBatch: vi.fn(),
    })

    expect(captureChangedCells).not.toHaveBeenCalled()
    expect(emit).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCells: [],
        explicitChangedCount: 1,
      }),
      Uint32Array.of(4),
      expect.any(Function),
    )
    expect(emitTracked).not.toHaveBeenCalled()
  })
})
