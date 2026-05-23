import type { EngineFormulaInitializationServiceArgs } from './formula-initialization-service-types.js'
import type { InitialResolvedFormulaEntry } from './formula-initialization-refs.js'
import type { FreshDirectScalarFormulaBindingMember } from './formula-binding-service-types.js'
import { canUseFreshDirectScalarFormulaBinding } from './formula-initialization-hydrated-direct-scalar.js'
import { compiledFormulaRequiresWorkbookMetadataBinding } from './formula-initialization-predicates.js'

const MIN_INITIAL_FRESH_DIRECT_SCALAR_FAST_BINDINGS = 32

export function canBindInitialFreshDirectScalarFormula(args: {
  readonly bindFreshDirectScalarFormulaRun: EngineFormulaInitializationServiceArgs['bindFreshDirectScalarFormulaRun']
  readonly hadExistingFormulas: boolean
  readonly prepared: InitialResolvedFormulaEntry
  readonly refsLength: number
}): boolean {
  return (
    !args.hadExistingFormulas &&
    args.refsLength >= MIN_INITIAL_FRESH_DIRECT_SCALAR_FAST_BINDINGS &&
    args.bindFreshDirectScalarFormulaRun !== undefined &&
    args.prepared.templateId !== undefined &&
    !compiledFormulaRequiresWorkbookMetadataBinding(args.prepared.compiled) &&
    canUseFreshDirectScalarFormulaBinding(args.prepared.compiled)
  )
}

export function createInitialFreshDirectScalarFormulaBindingMember(
  prepared: InitialResolvedFormulaEntry,
): FreshDirectScalarFormulaBindingMember {
  if (prepared.templateId === undefined) {
    throw new Error('Expected initial fresh direct scalar formula template id')
  }
  return {
    row: prepared.row,
    col: prepared.col,
    source: prepared.source,
    compiled: prepared.compiled,
    templateId: prepared.templateId,
  }
}
