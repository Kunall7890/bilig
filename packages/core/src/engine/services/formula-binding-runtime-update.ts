import type { CompiledPlanRecord, RuntimeFormula } from '../runtime-state.js'

export interface FormulaRuntimePlanFieldUpdate {
  readonly source: string
  readonly plan: CompiledPlanRecord
  readonly templateId: number | undefined
  readonly programLength: number
  readonly runtimeProgram?: Uint32Array
  readonly inlineScalarFastPlanKind?: RuntimeFormula['inlineScalarFastPlanKind']
  readonly inlineScalarArithmeticDeltaCoefficients?: RuntimeFormula['inlineScalarArithmeticDeltaCoefficients']
  readonly inlineScalarFastPlanStringIds?: RuntimeFormula['inlineScalarFastPlanStringIds']
  readonly inlineScalarPlanCellIndices?: Uint32Array | undefined
}

export function applyFormulaRuntimePlanFields(formula: RuntimeFormula, update: FormulaRuntimePlanFieldUpdate): void {
  formula.source = update.source
  formula.structuralSourceTransform = undefined
  formula.sourceRenameTransforms = undefined
  delete formula.preserveCachedValueOnFullRecalc
  formula.planId = update.plan.id
  formula.templateId = update.templateId
  formula.compiled = update.plan.compiled
  formula.plan = update.plan
  formula.inlineScalarFastPlanKind = update.inlineScalarFastPlanKind
  if (update.inlineScalarArithmeticDeltaCoefficients !== undefined || formula.inlineScalarArithmeticDeltaCoefficients !== undefined) {
    formula.inlineScalarArithmeticDeltaCoefficients = update.inlineScalarArithmeticDeltaCoefficients
  }
  if (update.inlineScalarFastPlanStringIds !== undefined || formula.inlineScalarFastPlanStringIds !== undefined) {
    formula.inlineScalarFastPlanStringIds = update.inlineScalarFastPlanStringIds
  }
  formula.inlineScalarPlanCellIndices = update.inlineScalarPlanCellIndices
  if (update.runtimeProgram !== undefined) {
    formula.runtimeProgram = update.runtimeProgram
  }
  formula.constants = update.plan.compiled.constants
  formula.programLength = update.programLength
  formula.constNumberLength = update.plan.compiled.constants.length
}
