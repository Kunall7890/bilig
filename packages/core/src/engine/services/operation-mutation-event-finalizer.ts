import type { CellRangeRef, EngineEvent } from '@bilig/protocol'
import type { U32 } from '../runtime-state.js'
import { canTrustPhysicalTrackedChangeSplit, tagTrustedPhysicalTrackedChanges } from './operation-change-helpers.js'
import { createOperationBatchMetrics } from './operation-batch-metrics.js'
import { emitOperationChangedEvents } from './operation-event-emission.js'
import type { DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

export interface OperationMutationEventFinalizerInvalidations {
  readonly invalidatedRanges: CellRangeRef[]
  readonly invalidatedRows: EngineEvent['invalidatedRows']
  readonly invalidatedColumns: EngineEvent['invalidatedColumns']
}

export interface FinalizeOperationMutationEventsArgs extends OperationMutationEventFinalizerInvalidations {
  readonly serviceArgs: CreateEngineOperationServiceArgs
  readonly suppressChangedSet: boolean
  readonly canComposeDisjointEventChanges: boolean
  readonly recalculated: U32
  readonly explicitChangedCount: number
  readonly changedInputCount: number
  readonly formulaChangedCount: number
  readonly compileMs: number
  readonly didRunRecalc: boolean
  readonly directFormulaMetrics: DirectFormulaMetricCounts
  readonly invalidation: EngineEvent['invalidation']
  readonly hasGeneralEventListeners: boolean
  readonly hasTrackedEventListeners: boolean
  readonly hasWatchedCellListeners: boolean
  readonly shouldMaterializeChangedCells: (changedLength: number) => boolean
}

export function finalizeOperationMutationEvents(request: FinalizeOperationMutationEventsArgs): void {
  const changed = request.suppressChangedSet
    ? new Uint32Array()
    : request.canComposeDisjointEventChanges
      ? request.serviceArgs.composeDisjointEventChanges(request.recalculated, request.explicitChangedCount)
      : request.serviceArgs.composeEventChanges(request.recalculated, request.explicitChangedCount)
  if (
    request.hasTrackedEventListeners &&
    request.canComposeDisjointEventChanges &&
    changed.length > 4 &&
    request.explicitChangedCount > 0 &&
    request.explicitChangedCount < changed.length
  ) {
    const sheetId = request.serviceArgs.state.workbook.cellStore.sheetIds[changed[0]!]
    if (
      sheetId !== undefined &&
      canTrustPhysicalTrackedChangeSplit(changed, sheetId, request.explicitChangedCount, request.serviceArgs.state.workbook)
    ) {
      tagTrustedPhysicalTrackedChanges(changed, sheetId, request.explicitChangedCount)
    }
  }
  const lastMetrics = createOperationBatchMetrics({
    previousMetrics: request.serviceArgs.state.getLastMetrics(),
    didRunRecalc: request.didRunRecalc,
    directFormulaMetrics: request.directFormulaMetrics,
    changedInputCount: request.changedInputCount,
    formulaChangedCount: request.formulaChangedCount,
    compileMs: request.compileMs,
  })
  request.serviceArgs.state.setLastMetrics(lastMetrics)
  emitOperationChangedEvents(request.serviceArgs, {
    changed,
    invalidation: request.invalidation,
    invalidatedRanges: request.invalidatedRanges,
    invalidatedRows: request.invalidatedRows,
    invalidatedColumns: request.invalidatedColumns,
    metrics: lastMetrics,
    explicitChangedCount: request.explicitChangedCount,
    hasGeneralEventListeners: request.hasGeneralEventListeners,
    hasTrackedEventListeners: request.hasTrackedEventListeners,
    hasWatchedCellListeners: request.hasWatchedCellListeners,
    materializeChangedCells: request.shouldMaterializeChangedCells(changed.length),
  })
}
