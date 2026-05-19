import {
  aggregateColumnDependencyKey,
  canEvaluatePostRecalcDirectFormulasWithoutKernel,
  collectTrackedDependents,
  composeSingleDisjointExplicitEventChanges,
  countDirectFormulaDeltaSkip,
  directAggregateNumericContribution,
  directCriteriaTouchesPoint,
  directFormulaChangesAreDisjointFromInputs,
  hasCompleteDirectFormulaDeltas,
  lookupImpactCacheKey,
} from './direct-formula-recalc-helpers.js'
import {
  cellRange,
  makeCompactExistingNumericMutationResult,
  makeExistingNumericMutationResult,
  mergeChangedCellIndices,
  rangesIntersect,
  tagTrustedPhysicalTrackedChanges,
  throwProtectionBlocked,
} from './operation-change-helpers.js'
import { canFinalizeStructuralNoValueMutationWithoutRecalc } from './operation-structural-no-value-finalization.js'

export const operationServiceTestHooks = {
  aggregateColumnDependencyKey,
  canFinalizeStructuralNoValueMutationWithoutRecalc,
  canEvaluatePostRecalcDirectFormulasWithoutKernel,
  cellRange,
  collectTrackedDependents,
  composeSingleDisjointExplicitEventChanges,
  countDirectFormulaDeltaSkip,
  directAggregateNumericContribution,
  directCriteriaTouchesPoint,
  directFormulaChangesAreDisjointFromInputs,
  getConstantDirectFormulaDeltas: hasCompleteDirectFormulaDeltas,
  lookupImpactCacheKey,
  makeCompactExistingNumericMutationResult,
  makeExistingNumericMutationResult,
  mergeChangedCellIndices,
  rangesIntersect,
  tagTrustedPhysicalTrackedChanges,
  throwProtectionBlocked,
}
