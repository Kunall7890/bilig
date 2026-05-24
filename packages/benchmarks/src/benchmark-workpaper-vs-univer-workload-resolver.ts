import {
  namedExpressionChangeScenario,
  rebuildAndRecalculateScenario,
  runtimeSnapshotBuildScenario,
  sheetRenameDependencyScenario,
} from './benchmark-workpaper-vs-univer-lifecycle-workloads.js'
import {
  crossSheetAggregateRecalcScenario,
  crossSheetDashboardBuildScenario,
  crossSheetDashboardRecalcScenario,
  crossSheetScalarFanoutRecalcScenario,
  manySheetsBuildScenario,
} from './benchmark-workpaper-vs-univer-multisheet-workloads.js'
import {
  batchMultiColumnScenario,
  batchSingleColumnScenario,
  batchSingleColumnUndoScenario,
  conditionalAggregationCriteriaEditScenario,
  conditionalAggregationMixedCriteriaScenario,
  conditionalAggregationScenario,
  conditionalAggregationSharedCriteriaScenario,
  formulaEditScenario,
  indexedLookupAfterBatchWriteScenario,
  indexedLookupAfterColumnWriteScenario,
  lookupApproximateAfterColumnWriteScenario,
  lookupApproximateDescendingScenario,
  mixedFrontierScenario,
  rangeReadDenseScenario,
  rangeReadFormulaGridScenario,
  rangeReadSparseWideScenario,
  rectangularBatchClearScenario,
  rectangularBatchEditScenario,
  singleEditRecalcScenario,
  suspendedBatchMultiColumnScenario,
  suspendedBatchSingleColumnScenario,
} from './benchmark-workpaper-vs-univer-operation-workloads.js'
import {
  aggregate2dCanonicalScenario,
  approximateDuplicateLookupCanonicalScenario,
  approximateLookupCanonicalScenario,
  denseLiteralBuildScenario,
  exactLookupCanonicalScenario,
  formulaChainRowScenario,
  formulaFanoutRowScenario,
  indexMatchExactCanonicalScenario,
  indexReferenceCanonicalScenario,
  mixedContentBuildScenario,
  overlappingAggregateCanonicalScenario,
  parserCacheMixedTemplateBuildScenario,
  parserCacheRowTemplateBuildScenario,
  parserCacheUniqueFormulaBuildScenario,
  slidingAggregateCanonicalScenario,
  textLookupCanonicalScenario,
} from './benchmark-workpaper-vs-univer-single-sheet-workloads.js'
import {
  appendFormulaRowsScenario,
  structuralDeleteColumnsScenario,
  structuralDeleteRowsScenario,
  structuralInsertColumnsScenario,
  structuralInsertRowsScenario,
  structuralMoveColumnsScenario,
  structuralMoveRowsScenario,
} from './benchmark-workpaper-vs-univer-structural-workloads.js'
import type { WorkPaperUniverScenario, WorkPaperUniverWorkload } from './benchmark-workpaper-vs-univer.js'

export function univerScenario(workload: WorkPaperUniverWorkload): WorkPaperUniverScenario {
  switch (workload) {
    case 'build-from-sheets':
      return denseLiteralBuildScenario(workload, 160, 24)
    case 'build-dense-literals':
      return denseLiteralBuildScenario(workload, 160, 24)
    case 'build-dense-literals-wide':
      return denseLiteralBuildScenario(workload, 96, 96)
    case 'build-dense-literals-tall':
      return denseLiteralBuildScenario(workload, 768, 12)
    case 'build-mixed-content':
      return mixedContentBuildScenario(workload, 750)
    case 'build-mixed-content-small':
      return mixedContentBuildScenario(workload, 250)
    case 'build-mixed-content-large':
      return mixedContentBuildScenario(workload, 1_500)
    case 'build-parser-cache-row-templates':
      return parserCacheRowTemplateBuildScenario(workload, 1_500)
    case 'build-parser-cache-mixed-templates':
      return parserCacheMixedTemplateBuildScenario(workload, 1_500)
    case 'build-parser-cache-unique-formulas':
      return parserCacheUniqueFormulaBuildScenario(workload, 1_500)
    case 'build-many-sheets':
      return manySheetsBuildScenario(workload, 6, 96, 16)
    case 'build-many-sheets-wide':
      return manySheetsBuildScenario(workload, 4, 64, 48)
    case 'build-many-sheets-narrow':
      return manySheetsBuildScenario(workload, 12, 128, 8)
    case 'build-cross-sheet-dashboard':
      return crossSheetDashboardBuildScenario(workload, 4, 500)
    case 'build-cross-sheet-dashboard-small':
      return crossSheetDashboardBuildScenario(workload, 2, 250)
    case 'build-cross-sheet-dashboard-large':
      return crossSheetDashboardBuildScenario(workload, 6, 750)
    case 'rebuild-and-recalculate':
      return rebuildAndRecalculateScenario(workload, 1_500)
    case 'rebuild-and-recalculate-large':
      return rebuildAndRecalculateScenario(workload, 3_000)
    case 'rebuild-runtime-from-snapshot':
      return runtimeSnapshotBuildScenario(workload, 1_500)
    case 'rebuild-runtime-from-snapshot-large':
      return runtimeSnapshotBuildScenario(workload, 3_000)
    case 'sheet-rename-dependencies':
      return sheetRenameDependencyScenario(workload)
    case 'named-expression-change':
      return namedExpressionChangeScenario(workload)
    case 'cross-sheet-scalar-recalc':
      return crossSheetScalarFanoutRecalcScenario(workload, 1_500)
    case 'cross-sheet-aggregate-recalc':
      return crossSheetAggregateRecalcScenario(workload, 1_500)
    case 'cross-sheet-dashboard-recalc':
      return crossSheetDashboardRecalcScenario(workload, 4, 1_000)
    case 'single-edit-chain':
      return formulaChainRowScenario(workload, 2_000)
    case 'single-edit-chain-small':
      return formulaChainRowScenario(workload, 500)
    case 'single-edit-chain-large':
      return formulaChainRowScenario(workload, 3_000)
    case 'single-edit-fanout':
      return formulaFanoutRowScenario(workload, 2_000)
    case 'single-edit-fanout-small':
      return formulaFanoutRowScenario(workload, 500)
    case 'single-edit-fanout-large':
      return formulaFanoutRowScenario(workload, 3_000)
    case 'single-edit-recalc':
      return singleEditRecalcScenario(workload, 2_000)
    case 'partial-recompute-mixed-frontier':
      return mixedFrontierScenario(workload, 1_500)
    case 'single-formula-edit-recalc':
      return formulaEditScenario(workload, 1_500)
    case 'single-formula-edit-recalc-large':
      return formulaEditScenario(workload, 3_000)
    case 'batch-edit-recalc':
      return batchSingleColumnScenario(workload, 500)
    case 'batch-edit-single-column':
      return batchSingleColumnScenario(workload, 500)
    case 'batch-edit-single-column-small':
      return batchSingleColumnScenario(workload, 128)
    case 'batch-edit-single-column-large':
      return batchSingleColumnScenario(workload, 1_000)
    case 'batch-edit-multi-column-small':
      return batchMultiColumnScenario(workload, 128)
    case 'batch-edit-multi-column':
      return batchMultiColumnScenario(workload, 250)
    case 'batch-edit-multi-column-large':
      return batchMultiColumnScenario(workload, 500)
    case 'batch-edit-rectangular-block':
      return rectangularBatchEditScenario(workload, 64, 12)
    case 'batch-edit-rectangular-block-wide':
      return rectangularBatchEditScenario(workload, 96, 16)
    case 'batch-clear-rectangular-block':
      return rectangularBatchClearScenario(workload, 64, 12)
    case 'batch-clear-rectangular-block-wide':
      return rectangularBatchClearScenario(workload, 96, 16)
    case 'batch-edit-single-column-with-undo':
      return batchSingleColumnUndoScenario(workload, 500)
    case 'batch-suspended-single-column':
      return suspendedBatchSingleColumnScenario(workload, 500)
    case 'batch-suspended-multi-column':
      return suspendedBatchMultiColumnScenario(workload, 250)
    case 'structural-insert-rows':
      return structuralInsertRowsScenario(workload, 1_500)
    case 'structural-insert-rows-small':
      return structuralInsertRowsScenario(workload, 500)
    case 'structural-insert-rows-large':
      return structuralInsertRowsScenario(workload, 3_000)
    case 'structural-append-formula-rows':
      return appendFormulaRowsScenario(workload, 750, 6, 250)
    case 'structural-append-formula-rows-small':
      return appendFormulaRowsScenario(workload, 250, 4, 100)
    case 'structural-append-formula-rows-large':
      return appendFormulaRowsScenario(workload, 1_000, 8, 250)
    case 'structural-delete-rows':
      return structuralDeleteRowsScenario(workload, 1_500)
    case 'structural-move-rows':
      return structuralMoveRowsScenario(workload, 1_500)
    case 'structural-insert-columns':
      return structuralInsertColumnsScenario(workload, 1_500)
    case 'structural-insert-columns-small':
      return structuralInsertColumnsScenario(workload, 500)
    case 'structural-insert-columns-large':
      return structuralInsertColumnsScenario(workload, 3_000)
    case 'structural-delete-columns':
      return structuralDeleteColumnsScenario(workload, 1_500)
    case 'structural-delete-columns-large':
      return structuralDeleteColumnsScenario(workload, 3_000)
    case 'structural-move-columns':
      return structuralMoveColumnsScenario(workload, 1_500)
    case 'structural-move-columns-large':
      return structuralMoveColumnsScenario(workload, 3_000)
    case 'range-read':
      return rangeReadDenseScenario(workload, 240, 24)
    case 'range-read-dense':
      return rangeReadDenseScenario(workload, 240, 24)
    case 'range-read-wide':
      return rangeReadDenseScenario(workload, 128, 96)
    case 'range-read-sparse-wide':
      return rangeReadSparseWideScenario(workload, 128, 96)
    case 'range-read-formula-grid':
      return rangeReadFormulaGridScenario(workload, 256, 4, 8)
    case 'range-read-formula-grid-wide':
      return rangeReadFormulaGridScenario(workload, 128, 8, 16)
    case 'aggregate-2d-ranges':
      return aggregate2dCanonicalScenario(workload, 1_500)
    case 'aggregate-2d-ranges-small':
      return aggregate2dCanonicalScenario(workload, 500)
    case 'aggregate-2d-ranges-large':
      return aggregate2dCanonicalScenario(workload, 3_000)
    case 'aggregate-overlapping-ranges':
      return overlappingAggregateCanonicalScenario(workload, 1_500)
    case 'aggregate-overlapping-ranges-small':
      return overlappingAggregateCanonicalScenario(workload, 500)
    case 'aggregate-overlapping-sliding-window':
      return slidingAggregateCanonicalScenario(workload, 1_500, 32)
    case 'aggregate-overlapping-sliding-window-wide':
      return slidingAggregateCanonicalScenario(workload, 1_500, 128)
    case 'lookup-no-column-index':
      return exactLookupCanonicalScenario(workload, 5_000, false)
    case 'lookup-no-column-index-small':
      return exactLookupCanonicalScenario(workload, 1_000, false)
    case 'lookup-with-column-index':
      return exactLookupCanonicalScenario(workload, 5_000, true)
    case 'lookup-with-column-index-large':
      return exactLookupCanonicalScenario(workload, 10_000, true)
    case 'lookup-index-match-exact':
      return indexMatchExactCanonicalScenario(workload, 5_000)
    case 'lookup-index-match-exact-large':
      return indexMatchExactCanonicalScenario(workload, 10_000)
    case 'lookup-index-reference':
      return indexReferenceCanonicalScenario(workload, 5_000)
    case 'lookup-index-reference-large':
      return indexReferenceCanonicalScenario(workload, 10_000)
    case 'lookup-with-column-index-after-column-write':
      return indexedLookupAfterColumnWriteScenario(workload, 5_000)
    case 'lookup-with-column-index-after-batch-write':
      return indexedLookupAfterBatchWriteScenario(workload, 5_000, 256)
    case 'lookup-with-column-index-after-batch-write-large':
      return indexedLookupAfterBatchWriteScenario(workload, 10_000, 512)
    case 'lookup-approximate-sorted':
      return approximateLookupCanonicalScenario(workload, 5_000)
    case 'lookup-approximate-sorted-large':
      return approximateLookupCanonicalScenario(workload, 10_000)
    case 'lookup-approximate-descending':
      return lookupApproximateDescendingScenario(workload, 5_000)
    case 'lookup-approximate-duplicates':
      return approximateDuplicateLookupCanonicalScenario(workload, 5_000)
    case 'lookup-approximate-sorted-after-column-write':
      return lookupApproximateAfterColumnWriteScenario(workload, 5_000)
    case 'lookup-text-exact':
      return textLookupCanonicalScenario(workload, 5_000)
    case 'lookup-text-exact-large':
      return textLookupCanonicalScenario(workload, 10_000)
    case 'conditional-aggregation-reused-ranges':
      return conditionalAggregationScenario(workload, 2_000, 32)
    case 'conditional-aggregation-reused-ranges-large':
      return conditionalAggregationScenario(workload, 3_000, 48)
    case 'conditional-aggregation-criteria-cell-edit':
      return conditionalAggregationCriteriaEditScenario(workload, 2_000, 32)
    case 'conditional-aggregation-shared-criteria':
      return conditionalAggregationSharedCriteriaScenario(workload, 2_000, 32)
    case 'conditional-aggregation-mixed-criteria':
      return conditionalAggregationMixedCriteriaScenario(workload, 2_000, 24)
  }
}
