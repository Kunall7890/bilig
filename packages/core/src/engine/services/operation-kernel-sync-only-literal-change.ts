import type { EngineEvent } from '@bilig/protocol'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'

type KernelSyncOnlyLiteralChangeState = Pick<EngineRuntimeState, 'counters' | 'events' | 'setLastMetrics'>

export function recordKernelSyncOnlyLiteralChange(args: {
  readonly state: KernelSyncOnlyLiteralChangeState
  readonly cellIndex: number
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => EngineEvent['metrics']
  readonly emitTracked: boolean
  readonly explicitChangedCount?: number
  readonly changedCellIndices?: U32
}): {
  readonly changedCellIndices: U32
  readonly metrics: EngineEvent['metrics']
} {
  addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
  args.deferSingleCellKernelSync(args.cellIndex)
  const metrics = args.makeSingleLiteralSkipMetrics()
  args.state.setLastMetrics(metrics)
  const changedCellIndices = args.changedCellIndices ?? Uint32Array.of(args.cellIndex)
  if (args.emitTracked) {
    args.state.events.emitTracked({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices,
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics,
      explicitChangedCount: args.explicitChangedCount ?? 1,
    })
  }
  return { changedCellIndices, metrics }
}
