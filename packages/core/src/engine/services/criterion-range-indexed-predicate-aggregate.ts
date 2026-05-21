import { compileCriteriaMatcher, type CompiledCriteriaMatcher } from '@bilig/formula'
import type { CellValue } from '@bilig/protocol'
import type { CriterionRowSetView } from './criterion-rowset-view.js'
import type { EngineRuntimeColumnStoreService, RuntimeColumnView } from './runtime-column-store-service.js'
import type {
  CriterionCompoundExactAggregateRequest,
  CriterionExactAggregateBucket,
  CriterionExactAggregateRequest,
  CriterionRangeDescriptor,
  SliceFastPredicate,
} from './criterion-range-cache-service.js'

const INDEXED_PREDICATE_AGGREGATE_MAX_ANCHOR_DENSITY = 0.75

export function getOrBuildIndexedPredicateAggregateFromColumnViews(args: {
  readonly request: CriterionCompoundExactAggregateRequest
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  readonly getColumnView: (request: CriterionRangeDescriptor) => RuntimeColumnView
  readonly buildSlicePredicate: (compiled: CompiledCriteriaMatcher) => SliceFastPredicate
  readonly readIndexedEqualityRows: (
    view: RuntimeColumnView,
    predicate: SliceFastPredicate,
    runtimeColumnStore: EngineRuntimeColumnStoreService,
  ) => CriterionRowSetView | undefined
  readonly slicePredicateMatches: (
    predicate: SliceFastPredicate,
    view: RuntimeColumnView,
    offset: number,
    runtimeColumnStore: EngineRuntimeColumnStoreService,
  ) => boolean
  readonly createExactAggregateBucket: () => CriterionExactAggregateBucket
  readonly addExactAggregateMatch: (
    bucket: CriterionExactAggregateBucket,
    aggregateView: RuntimeColumnView | undefined,
    offset: number,
  ) => void
  readonly exactAggregateValueFromBucket: (
    bucket: CriterionExactAggregateBucket | undefined,
    aggregateKind: CriterionExactAggregateRequest['aggregateKind'],
  ) => CellValue
}): CellValue | undefined {
  const { request } = args
  const { criteriaPairs } = request
  if (criteriaPairs.length === 0) {
    return undefined
  }
  const expectedLength = criteriaPairs[0]!.range.length
  if (criteriaPairs.some((pair) => pair.range.length !== expectedLength)) {
    return undefined
  }
  if (request.aggregateKind !== 'count' && request.aggregateRange === undefined) {
    return undefined
  }
  if (request.aggregateRange !== undefined && request.aggregateRange.length !== expectedLength) {
    return undefined
  }

  const predicates = criteriaPairs.map((pair) => args.buildSlicePredicate(compileCriteriaMatcher(pair.criteria)))
  const resolvedPairs = criteriaPairs.map((pair, index) => ({
    view: args.getColumnView(pair.range),
    predicate: predicates[index]!,
  }))

  let limitingPairIndex: number | undefined
  let limitingRows: CriterionRowSetView | undefined
  for (let pairIndex = 0; pairIndex < resolvedPairs.length; pairIndex += 1) {
    const pair = resolvedPairs[pairIndex]!
    const rowset = args.readIndexedEqualityRows(pair.view, pair.predicate, args.runtimeColumnStore)
    if (rowset === undefined) {
      continue
    }
    if (limitingRows === undefined || rowset.cardinality < limitingRows.cardinality) {
      limitingPairIndex = pairIndex
      limitingRows = rowset
    }
  }
  if (limitingRows === undefined || limitingPairIndex === undefined) {
    return undefined
  }
  if (expectedLength > 0 && limitingRows.cardinality / expectedLength > INDEXED_PREDICATE_AGGREGATE_MAX_ANCHOR_DENSITY) {
    return undefined
  }

  const aggregateView = request.aggregateRange === undefined ? undefined : args.getColumnView(request.aggregateRange)
  const bucket = args.createExactAggregateBucket()
  limitingRows.forEachOffset((rowOffset) => {
    for (let pairIndex = 0; pairIndex < resolvedPairs.length; pairIndex += 1) {
      if (pairIndex === limitingPairIndex) {
        continue
      }
      const pair = resolvedPairs[pairIndex]!
      if (!args.slicePredicateMatches(pair.predicate, pair.view, rowOffset, args.runtimeColumnStore)) {
        return
      }
    }
    args.addExactAggregateMatch(bucket, aggregateView, rowOffset)
  })

  return args.exactAggregateValueFromBucket(bucket.count === 0 ? undefined : bucket, request.aggregateKind)
}
