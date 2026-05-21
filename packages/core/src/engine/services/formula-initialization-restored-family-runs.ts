import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState } from '../runtime-state.js'
import { readFreshFormulaFamilyRunsFromRefs, type DeferredInitialFormulaFamilyRun } from './formula-initialization-family-runs.js'
import type { InitialFormulaEntryRefSource } from './formula-initialization-refs.js'

export function readAlignedFreshFormulaFamilyRunsFromRefs<Entry>(args: {
  readonly refs: InitialFormulaEntryRefSource<Entry>
  readonly hadExistingFormulas: boolean
  readonly counters: EngineRuntimeState['counters']
}): readonly DeferredInitialFormulaFamilyRun[] | undefined {
  const freshFormulaFamilyRuns = readFreshFormulaFamilyRunsFromRefs(args.refs)
  if (args.hadExistingFormulas && freshFormulaFamilyRuns === undefined) {
    return undefined
  }
  const alignedFreshFormulaFamilyRuns = freshFormulaFamilyRuns?.runs
  if (freshFormulaFamilyRuns?.fallbackCount) {
    addEngineCounter(args.counters, 'formulaFamilyRuntimeRunFallbacks', freshFormulaFamilyRuns.fallbackCount)
  }
  if (alignedFreshFormulaFamilyRuns !== undefined) {
    addEngineCounter(args.counters, 'formulaFamilyRuntimeRunsRestored', alignedFreshFormulaFamilyRuns.length)
    addEngineCounter(
      args.counters,
      'formulaFamilyRuntimeRunMembersRestored',
      alignedFreshFormulaFamilyRuns.reduce((sum, run) => sum + run.cellIndices.length, 0),
    )
  }
  return alignedFreshFormulaFamilyRuns
}

export function flushAlignedFreshFormulaFamilyRuns(args: {
  readonly runs: readonly DeferredInitialFormulaFamilyRun[] | undefined
  readonly shouldDeferFormulaFamilyIndex: boolean
  readonly deferFormulaFamilyIndexRuns?: ((runs: readonly DeferredInitialFormulaFamilyRun[]) => void) | undefined
  readonly deferFormulaFamilyIndexRebuild?: (() => void) | undefined
  readonly registerFormulaFamilyRun: (run: DeferredInitialFormulaFamilyRun) => void
  readonly checkEvaluationBudget?: (() => void) | undefined
}): boolean {
  if (args.runs === undefined) {
    return false
  }
  if (args.shouldDeferFormulaFamilyIndex) {
    if (args.deferFormulaFamilyIndexRuns) {
      args.deferFormulaFamilyIndexRuns(args.runs)
    } else {
      args.deferFormulaFamilyIndexRebuild?.()
    }
    return true
  }
  args.runs.forEach((run) => {
    args.checkEvaluationBudget?.()
    args.registerFormulaFamilyRun(run)
  })
  return true
}
