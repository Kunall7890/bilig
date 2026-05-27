import type { EngineExistingNumericCellMutationsRef, SpreadsheetEngine } from '@bilig/core/headless-runtime'
import type { WorkPaperCellMutationApplyOptions } from './work-paper-cell-mutation-refs.js'

type ExistingNumericBatchMutationEngine = SpreadsheetEngine & {
  readonly tryApplyExistingNumericCellMutationsAt?: (request: EngineExistingNumericCellMutationsRef) => boolean
}

export function tryApplyExistingNumericCellMutationsAtWithOptions(
  engine: SpreadsheetEngine,
  record: EngineExistingNumericCellMutationsRef,
  options: WorkPaperCellMutationApplyOptions,
): boolean {
  if (
    options.captureUndo !== true ||
    options.source !== 'local' ||
    options.returnUndoOps !== false ||
    (options.potentialNewCells ?? 0) !== 0
  ) {
    return false
  }
  const mutationEngine = engine as ExistingNumericBatchMutationEngine
  const applyExistingNumericCellMutations = mutationEngine.tryApplyExistingNumericCellMutationsAt
  if (typeof applyExistingNumericCellMutations !== 'function') {
    return false
  }
  return applyExistingNumericCellMutations.call(mutationEngine, record)
}
