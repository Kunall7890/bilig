import type { CellRangeRef, EngineEvent, RecalcMetrics } from '@bilig/protocol'
import type { U32 } from '../runtime-state.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

const GENERAL_CHANGED_CELL_PAYLOAD_LIMIT = 512

type OperationEventAccess = Pick<CreateEngineOperationServiceArgs, 'state' | 'captureChangedCells' | 'captureChangedPatches'>

interface OperationChangedEventInvalidations {
  readonly invalidatedRanges: CellRangeRef[]
  readonly invalidatedRows: EngineEvent['invalidatedRows']
  readonly invalidatedColumns: EngineEvent['invalidatedColumns']
}

interface EmitOperationChangedEventsArgs extends OperationChangedEventInvalidations {
  readonly changed: U32
  readonly invalidation: EngineEvent['invalidation']
  readonly metrics: RecalcMetrics
  readonly explicitChangedCount: number
  readonly hasGeneralEventListeners: boolean
  readonly hasTrackedEventListeners: boolean
  readonly hasWatchedCellListeners: boolean
  readonly materializeChangedCells: boolean
}

export function shouldMaterializeOperationChangedCells(args: {
  readonly changedLength: number
  readonly hasGeneralEventListeners: boolean
  readonly invalidation: EngineEvent['invalidation']
  readonly invalidations: OperationChangedEventInvalidations
}): boolean {
  return (
    args.hasGeneralEventListeners &&
    args.invalidation !== 'full' &&
    (args.changedLength <= GENERAL_CHANGED_CELL_PAYLOAD_LIMIT ||
      (args.invalidations.invalidatedRanges.length === 0 &&
        args.invalidations.invalidatedRows.length === 0 &&
        args.invalidations.invalidatedColumns.length === 0))
  )
}

export function emitOperationChangedEvents(access: OperationEventAccess, args: EmitOperationChangedEventsArgs): void {
  if (args.hasGeneralEventListeners || args.hasWatchedCellListeners) {
    const event: EngineEvent & { explicitChangedCount: number } = {
      kind: 'batch',
      invalidation: args.invalidation,
      changedCellIndices: args.changed,
      changedCells: args.materializeChangedCells ? access.captureChangedCells(args.changed) : [],
      invalidatedRanges: args.invalidatedRanges,
      invalidatedRows: args.invalidatedRows,
      invalidatedColumns: args.invalidatedColumns,
      metrics: args.metrics,
      explicitChangedCount: args.explicitChangedCount,
    }
    if (event.invalidation === 'full') {
      access.state.events.emitAllWatched(event)
    } else {
      access.state.events.emit(event, args.changed, (cellIndex) => access.state.workbook.getQualifiedAddress(cellIndex))
    }
  }
  if (!args.hasTrackedEventListeners) {
    return
  }
  const patchRequest = {
    invalidation: args.invalidation,
    invalidatedRanges: args.invalidatedRanges,
    invalidatedRows: args.invalidatedRows,
    invalidatedColumns: args.invalidatedColumns,
  } satisfies Parameters<typeof access.captureChangedPatches>[1]
  const shouldCapturePatches =
    args.changed.length > 0 &&
    (patchRequest.invalidation !== 'cells' ||
      patchRequest.invalidatedRanges.length > 0 ||
      patchRequest.invalidatedRows.length > 0 ||
      patchRequest.invalidatedColumns.length > 0)
  const patches = shouldCapturePatches ? access.captureChangedPatches(args.changed, patchRequest) : undefined
  access.state.events.emitTracked({
    kind: 'batch',
    invalidation: args.invalidation,
    changedCellIndices: args.changed,
    ...(patches ? { patches } : {}),
    invalidatedRanges: args.invalidatedRanges,
    invalidatedRows: args.invalidatedRows,
    invalidatedColumns: args.invalidatedColumns,
    metrics: args.metrics,
    explicitChangedCount: args.explicitChangedCount,
  })
}
