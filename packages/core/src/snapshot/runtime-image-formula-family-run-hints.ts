import type { DeferredInitialFormulaFamilyRun } from '../engine/services/formula-initialization-family-runs.js'

type RuntimeFormulaFamilyRunHintSource = {
  readonly length: number
  freshFormulaFamilyRuns?: readonly DeferredInitialFormulaFamilyRun[] | undefined
  freshFormulaFamilyRunFallbackCount?: number
}

export function attachRuntimeFormulaFamilyRunHints(args: {
  readonly refs: RuntimeFormulaFamilyRunHintSource
  readonly restoredRuns:
    | {
        readonly runs: readonly DeferredInitialFormulaFamilyRun[]
        readonly fallbackCount: number
      }
    | undefined
  readonly runtimeRunCount: number
  readonly cellIndicesAreRuntimeAligned: boolean
}): void {
  const { refs, restoredRuns, runtimeRunCount } = args
  if (refs.length === 0 || runtimeRunCount === 0) {
    return
  }
  refs.freshFormulaFamilyRuns = undefined
  if (args.cellIndicesAreRuntimeAligned && restoredRuns !== undefined && restoredRuns.fallbackCount === 0 && restoredRuns.runs.length > 0) {
    refs.freshFormulaFamilyRuns = restoredRuns.runs
    refs.freshFormulaFamilyRunFallbackCount = 0
    return
  }
  refs.freshFormulaFamilyRunFallbackCount = runtimeRunCount
}
