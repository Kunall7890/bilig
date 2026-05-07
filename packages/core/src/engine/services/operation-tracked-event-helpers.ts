import type { EngineEvent } from '@bilig/protocol'

export interface OperationTrackedEventEmitter {
  readonly emitTracked: (event: OperationTrackedCellsBatchEvent) => void
}

export interface OperationTrackedCellsBatchEvent {
  readonly kind: EngineEvent['kind']
  readonly invalidation: 'cells'
  readonly changedCellIndices: EngineEvent['changedCellIndices']
  readonly invalidatedRanges: []
  readonly invalidatedRows: []
  readonly invalidatedColumns: []
  readonly metrics: EngineEvent['metrics']
  readonly explicitChangedCount: number
}

export function emitOperationTrackedCellsBatch(args: {
  readonly events: OperationTrackedEventEmitter
  readonly changedCellIndices: EngineEvent['changedCellIndices']
  readonly metrics: EngineEvent['metrics']
  readonly explicitChangedCount?: number
}): void {
  args.events.emitTracked({
    kind: 'batch',
    invalidation: 'cells',
    changedCellIndices: args.changedCellIndices,
    invalidatedRanges: [],
    invalidatedRows: [],
    invalidatedColumns: [],
    metrics: args.metrics,
    explicitChangedCount: args.explicitChangedCount ?? 1,
  })
}
