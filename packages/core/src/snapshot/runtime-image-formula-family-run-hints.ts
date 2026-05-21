import type { DeferredInitialFormulaFamilyRun } from '../engine/services/formula-initialization-family-runs.js'
import type { RestoredRuntimeFormulaFamilyRuns } from './runtime-image-formula-family-runs.js'

type RuntimeFormulaFamilyRunHintSource = {
  readonly length: number
  freshFormulaFamilyRuns?: readonly DeferredInitialFormulaFamilyRun[] | undefined
  freshFormulaFamilyRunFallbackCount?: number
}

export function selectRuntimeFormulaFamilyRunHints(args: {
  readonly restoredRuns: RestoredRuntimeFormulaFamilyRuns | undefined
  readonly cellIndices: ArrayLike<number>
}): RestoredRuntimeFormulaFamilyRuns | undefined {
  const restoredRuns = args.restoredRuns
  if (restoredRuns === undefined || restoredRuns.runs.length === 0 || args.cellIndices.length === 0) {
    return undefined
  }
  const availableCellIndices = new Set<number>()
  for (let index = 0; index < args.cellIndices.length; index += 1) {
    availableCellIndices.add(args.cellIndices[index]!)
  }
  const runs = restoredRuns.runs.filter((run) => run.cellIndices.every((cellIndex) => availableCellIndices.has(cellIndex)))
  return runs.length === 0 ? undefined : { runs, fallbackCount: 0 }
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
