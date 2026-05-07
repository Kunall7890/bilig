import type { EngineEvent, LiteralInput } from '@bilig/protocol'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState } from '../runtime-state.js'
import { emitOperationTrackedCellsBatch } from './operation-tracked-event-helpers.js'

interface OperationKernelSyncLiteralFastPathArgs {
  readonly state: Pick<EngineRuntimeState, 'counters' | 'events' | 'setLastMetrics'>
  readonly writeFastPathLiteralToExistingCell: (cellIndex: number, value: LiteralInput) => void
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => EngineEvent['metrics']
}

interface OperationKernelSyncLiteralFastPathRequest {
  readonly existingIndex: number
  readonly value: LiteralInput
  readonly emitTracked: boolean
  readonly afterWrite?: () => void
}

export function tryApplySingleKernelSyncOnlyLiteralMutationFastPath(
  args: OperationKernelSyncLiteralFastPathArgs,
  request: OperationKernelSyncLiteralFastPathRequest,
): boolean {
  args.writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
  request.afterWrite?.()
  addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
  args.deferSingleCellKernelSync(request.existingIndex)
  const lastMetrics = args.makeSingleLiteralSkipMetrics()
  args.state.setLastMetrics(lastMetrics)
  if (request.emitTracked) {
    emitOperationTrackedCellsBatch({
      events: args.state.events,
      changedCellIndices: Uint32Array.of(request.existingIndex),
      metrics: lastMetrics,
    })
  }
  return true
}
