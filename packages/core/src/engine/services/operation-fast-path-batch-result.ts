import type { EngineChangedCell, EngineEvent } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook'
import type { EngineSyncClientConnection, U32 } from '../runtime-state.js'

export interface OperationFastPathBatchResultState {
  readonly events: {
    readonly emit: (
      event: EngineEvent & { explicitChangedCount: number },
      changedCellIndices: readonly number[] | U32,
      resolveAddress?: (cellIndex: number) => string,
    ) => void
    readonly emitTracked: (event: {
      kind: EngineEvent['kind']
      invalidation: EngineEvent['invalidation']
      changedCellIndices: EngineEvent['changedCellIndices']
      invalidatedRanges: EngineEvent['invalidatedRanges']
      invalidatedRows: EngineEvent['invalidatedRows']
      invalidatedColumns: EngineEvent['invalidatedColumns']
      metrics: EngineEvent['metrics']
      explicitChangedCount: number
    }) => void
  }
  readonly workbook: {
    readonly getQualifiedAddress: (cellIndex: number) => string
  }
  readonly getLastMetrics: () => EngineEvent['metrics']
  readonly setLastMetrics: (metrics: EngineEvent['metrics']) => void
  readonly getSyncClientConnection: () => EngineSyncClientConnection | null
}

export interface EmitCellMutationFastPathBatchResultArgs {
  readonly state: OperationFastPathBatchResultState
  readonly changed: U32
  readonly changedInputCount: number
  readonly explicitChangedCount: number
  readonly hasGeneralEventListeners: boolean
  readonly hasTrackedEventListeners: boolean
  readonly hasWatchedCellListeners: boolean
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
  readonly batch: EngineOpBatch | null
  readonly emitBatch: (batch: EngineOpBatch) => void
}

export function emitCellMutationFastPathBatchResult(args: EmitCellMutationFastPathBatchResultArgs): void {
  const previousMetrics = args.state.getLastMetrics()
  const lastMetrics = {
    ...previousMetrics,
    dirtyFormulaCount: 0,
    wasmFormulaCount: 0,
    jsFormulaCount: 0,
    rangeNodeVisits: 0,
    recalcMs: 0,
    batchId: previousMetrics.batchId + 1,
    changedInputCount: args.changedInputCount,
    compileMs: 0,
  }
  args.state.setLastMetrics(lastMetrics)
  if (args.hasGeneralEventListeners || args.hasWatchedCellListeners) {
    const event: EngineEvent & { explicitChangedCount: number } = {
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices: args.changed,
      changedCells: args.hasGeneralEventListeners ? args.captureChangedCells(args.changed) : [],
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics: lastMetrics,
      explicitChangedCount: args.explicitChangedCount,
    }
    args.state.events.emit(event, args.changed, (cellIndex) => args.state.workbook.getQualifiedAddress(cellIndex))
  }
  if (args.hasTrackedEventListeners) {
    args.state.events.emitTracked({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices: args.changed,
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics: lastMetrics,
      explicitChangedCount: args.explicitChangedCount,
    })
  }
  if (args.batch) {
    void args.state.getSyncClientConnection()?.send(args.batch)
    args.emitBatch(args.batch)
  }
}
