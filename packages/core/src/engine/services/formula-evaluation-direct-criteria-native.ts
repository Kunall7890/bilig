import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, RuntimeDirectCriteriaDescriptor } from '../runtime-state.js'
import type { CriterionRangeDescriptor, CriterionRangeMatch } from './criterion-range-cache-service.js'
import type { EngineRuntimeColumnStoreService } from './runtime-column-store-service.js'

const MIN_NATIVE_DIRECT_CRITERIA_MATCHED_AGGREGATE_ROWS = 64
const NATIVE_DIRECT_AGGREGATE_OP_SUM = 1
const NATIVE_DIRECT_AGGREGATE_OP_AVERAGE = 2
const NATIVE_DIRECT_AGGREGATE_OP_COUNT = 3
const NATIVE_DIRECT_AGGREGATE_OP_MIN = 4
const NATIVE_DIRECT_AGGREGATE_OP_MAX = 5

export function tryEvaluateNativeDirectCriteriaMatchedAggregate(
  args: {
    readonly state: Pick<EngineRuntimeState, 'wasm' | 'counters'>
    readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  },
  input: {
    readonly aggregateKind: RuntimeDirectCriteriaDescriptor['aggregateKind']
    readonly aggregateRange: CriterionRangeDescriptor
    readonly matches: CriterionRangeMatch
  },
): CellValue | undefined {
  const aggregateKind = nativeDirectCriteriaAggregateKind(input.aggregateKind)
  if (
    aggregateKind === undefined ||
    input.matches.length < MIN_NATIVE_DIRECT_CRITERIA_MATCHED_AGGREGATE_ROWS ||
    !args.state.wasm.initSyncIfPossible()
  ) {
    return undefined
  }

  const aggregateSlice = args.runtimeColumnStore.getColumnSlice({
    sheetName: input.aggregateRange.sheetName,
    rowStart: input.aggregateRange.rowStart,
    rowEnd: input.aggregateRange.rowEnd,
    col: input.aggregateRange.col,
  })
  const outTags = new Uint8Array(1)
  const outNumbers = new Float64Array(1)
  const outErrors = new Uint16Array(1)
  const matchedRows =
    input.matches.rows.length === input.matches.length ? input.matches.rows : input.matches.rows.subarray(0, input.matches.length)

  if (
    !args.state.wasm.evalDirectCriteriaMatchedAggregateBatch({
      aggregateKinds: Uint8Array.of(aggregateKind),
      matchStarts: Uint32Array.of(0),
      matchLengths: Uint32Array.of(input.matches.length),
      matchedRows,
      aggregateTags: aggregateSlice.tags,
      aggregateNumbers: aggregateSlice.numbers,
      aggregateErrors: aggregateSlice.errors,
      outTags,
      outNumbers,
      outErrors,
    })
  ) {
    return undefined
  }

  addEngineCounter(args.state.counters, 'nativeDirectCriteriaAggregateEvaluations')
  const tag = (outTags[0] as ValueTag | undefined) ?? ValueTag.Empty
  if (tag === ValueTag.Number) {
    return { tag: ValueTag.Number, value: outNumbers[0] ?? 0 }
  }
  if (tag === ValueTag.Error) {
    return { tag: ValueTag.Error, code: (outErrors[0] as ErrorCode | undefined) ?? ErrorCode.None }
  }
  return undefined
}

function nativeDirectCriteriaAggregateKind(kind: RuntimeDirectCriteriaDescriptor['aggregateKind']): number | undefined {
  switch (kind) {
    case 'sum':
      return NATIVE_DIRECT_AGGREGATE_OP_SUM
    case 'average':
      return NATIVE_DIRECT_AGGREGATE_OP_AVERAGE
    case 'count':
      return NATIVE_DIRECT_AGGREGATE_OP_COUNT
    case 'min':
      return NATIVE_DIRECT_AGGREGATE_OP_MIN
    case 'max':
      return NATIVE_DIRECT_AGGREGATE_OP_MAX
    case 'first':
      return undefined
  }
}
