import type { RecalcMetrics } from '@bilig/protocol'
import type { DirectFormulaMetricCounts } from './operation-post-recalc-direct-formulas.js'

export function createOperationBatchMetrics(args: {
  readonly previousMetrics: RecalcMetrics
  readonly didRunRecalc: boolean
  readonly directFormulaMetrics: DirectFormulaMetricCounts
  readonly changedInputCount: number
  readonly formulaChangedCount: number
  readonly compileMs: number
}): RecalcMetrics {
  return {
    ...args.previousMetrics,
    ...(args.didRunRecalc
      ? {
          dirtyFormulaCount: args.previousMetrics.dirtyFormulaCount,
          wasmFormulaCount: args.previousMetrics.wasmFormulaCount + args.directFormulaMetrics.wasmFormulaCount,
          jsFormulaCount: args.previousMetrics.jsFormulaCount + args.directFormulaMetrics.jsFormulaCount,
        }
      : {
          dirtyFormulaCount: 0,
          wasmFormulaCount: args.directFormulaMetrics.wasmFormulaCount,
          jsFormulaCount: args.directFormulaMetrics.jsFormulaCount,
          rangeNodeVisits: 0,
          recalcMs: 0,
        }),
    batchId: args.previousMetrics.batchId + 1,
    changedInputCount: args.changedInputCount + args.formulaChangedCount,
    compileMs: args.compileMs,
  }
}
