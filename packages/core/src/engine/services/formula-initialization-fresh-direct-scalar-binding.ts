import type { EngineCounters } from '../../perf/engine-counters.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineFormulaInitializationServiceArgs } from './formula-initialization-service-types.js'
import type { InitialResolvedFormulaEntry } from './formula-initialization-refs.js'
import { rethrowFatalFormulaBindingError } from './formula-binding-error-policy.js'
import { canUseFreshDirectScalarFormulaBinding } from './formula-initialization-hydrated-direct-scalar.js'
import { compiledFormulaRequiresWorkbookMetadataBinding } from './formula-initialization-predicates.js'

const MIN_INITIAL_FRESH_DIRECT_SCALAR_FAST_BINDINGS = 32

export function tryBindInitialFreshDirectScalarFormula(args: {
  readonly bindFreshDirectScalarFormulaRun: EngineFormulaInitializationServiceArgs['bindFreshDirectScalarFormulaRun']
  readonly counters: EngineCounters
  readonly hadExistingFormulas: boolean
  readonly prepared: InitialResolvedFormulaEntry
  readonly refsLength: number
}): boolean {
  const { prepared } = args
  if (
    args.hadExistingFormulas ||
    args.refsLength < MIN_INITIAL_FRESH_DIRECT_SCALAR_FAST_BINDINGS ||
    args.bindFreshDirectScalarFormulaRun === undefined ||
    prepared.templateId === undefined ||
    compiledFormulaRequiresWorkbookMetadataBinding(prepared.compiled) ||
    !canUseFreshDirectScalarFormulaBinding(prepared.compiled)
  ) {
    return false
  }
  try {
    args.bindFreshDirectScalarFormulaRun({
      sheetId: prepared.sheetId,
      ownerSheetName: prepared.ownerSheetName,
      cellIndex: prepared.cellIndex,
      member: {
        row: prepared.row,
        col: prepared.col,
        source: prepared.source,
        compiled: prepared.compiled,
        templateId: prepared.templateId,
      },
    })
    addEngineCounter(args.counters, 'initialFreshDirectScalarFastBindings')
    return true
  } catch (error) {
    rethrowFatalFormulaBindingError(error)
    return false
  }
}
