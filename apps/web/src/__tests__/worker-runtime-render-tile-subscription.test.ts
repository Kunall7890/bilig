import { describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import type { EngineEvent, RecalcMetrics } from '@bilig/protocol'
import { WorkerRuntimeRenderTileDeltaPublisher } from '../worker-runtime-render-tile-subscription.js'

const TEST_METRICS: RecalcMetrics = {
  batchId: 1,
  changedInputCount: 0,
  compileMs: 0,
  dirtyFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  wasmFormulaCount: 0,
}

function createFullEvent(): EngineEvent {
  return {
    kind: 'batch',
    invalidation: 'full',
    changedCellIndices: new Uint32Array(),
    changedCells: [],
    invalidatedColumns: [],
    invalidatedRanges: [],
    invalidatedRows: [],
    metrics: TEST_METRICS,
  }
}

describe('WorkerRuntimeRenderTileDeltaPublisher', () => {
  it('does not publish synthetic tile replacements for a missing sheet', async () => {
    let emit: ((event: EngineEvent) => void) | null = null
    const listener = vi.fn()
    const publisher = new WorkerRuntimeRenderTileDeltaPublisher()
    const engine = new SpreadsheetEngine({ workbookName: 'render-tile-missing-sheet-proof' })
    await engine.ready()
    engine.subscribe = (next) => {
      emit = next
      return () => undefined
    }

    publisher.subscribe({
      getProjectionEngine: async () => engine,
      listener,
      subscription: {
        sheetId: 7,
        sheetName: 'Missing',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
    })

    await vi.waitFor(() => {
      expect(emit).not.toBeNull()
    })

    expect(listener).not.toHaveBeenCalled()

    emit?.(createFullEvent())
    expect(listener).not.toHaveBeenCalled()

    engine.createSheet('Missing')
    emit?.(createFullEvent())

    expect(listener).toHaveBeenCalledTimes(1)
  })
})
