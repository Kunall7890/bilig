import type { EngineCellMutationRef, SpreadsheetEngine } from '@bilig/core'
import { WorkPaperMutationQueues } from './work-paper-mutation-queues.js'
import type { WorkPaperSheetDimensionCache } from './work-paper-sheet-dimension-cache.js'

export function createWorkPaperRuntimeMutationQueues(args: {
  readonly getEngine: () => SpreadsheetEngine
  readonly getSheetDimensionCache: () => WorkPaperSheetDimensionCache
  readonly updateSheetDimensionsAfterCellMutationRefs: (refs: readonly EngineCellMutationRef[]) => void
}): WorkPaperMutationQueues {
  return new WorkPaperMutationQueues({
    applyCellMutationsAtWithOptions: (refs, options) => {
      args.getEngine().applyCellMutationsAtWithOptions(refs, options)
    },
    canSkipSheetDimensionUpdateAfterLiteralMutationRefs: (refs, potentialNewCells) =>
      args.getSheetDimensionCache().canSkipUpdateAfterLiteralMutationRefs(refs, potentialNewCells),
    updateSheetDimensionsAfterCellMutationRefs: args.updateSheetDimensionsAfterCellMutationRefs,
  })
}
