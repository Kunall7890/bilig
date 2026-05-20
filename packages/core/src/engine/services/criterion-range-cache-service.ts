import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import {
  compileCriteriaMatcher,
  matchesCompiledCriteria,
  normalizeExactLookupNumber,
  type CompiledCriteriaMatcher,
  type CriteriaOperator,
} from '@bilig/formula'
import type { EngineRuntimeColumnStoreService, RuntimeColumnView } from './runtime-column-store-service.js'
import { sliceAbsoluteRowsToRangeView, type CriterionRowSetView } from './criterion-rowset-view.js'
import type { DepPatternStore } from '../../deps/dep-pattern-store.js'
import type { RegionGraph } from '../../deps/region-graph.js'

export interface CriterionRangeDescriptor {
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly col: number
  readonly length: number
}

export interface CriterionRangePair {
  readonly range: CriterionRangeDescriptor
  readonly criteria: CellValue
}

export interface CriterionRangeMatch {
  readonly rows: Uint32Array
  readonly length: number
}

export interface CriterionRangeCacheService {
  readonly getOrBuildMatchingRows: (request: { criteriaPairs: readonly CriterionRangePair[] }) => CriterionRangeMatch | CellValue
  readonly getOrBuildExactCriteriaAggregate: (request: CriterionCompoundExactAggregateRequest) => CellValue | undefined
  readonly getOrBuildExactAggregate: (request: CriterionExactAggregateRequest) => CellValue | undefined
  readonly getOrBuildCompoundExactAggregate: (request: CriterionCompoundExactAggregateRequest) => CellValue | undefined
}

interface CriterionCacheEntry {
  readonly rows: Uint32Array
  readonly length: number
  readonly pairVersions: ReadonlyArray<{
    columnVersion: number
    structureVersion: number
  }>
}

export interface CriterionExactAggregateRequest {
  readonly criteriaPair: CriterionRangePair
  readonly aggregateRange?: CriterionRangeDescriptor
  readonly aggregateKind: 'count' | 'sum' | 'average' | 'min' | 'max'
}

export interface CriterionCompoundExactAggregateRequest {
  readonly criteriaPairs: readonly CriterionRangePair[]
  readonly aggregateRange?: CriterionRangeDescriptor
  readonly aggregateKind: CriterionExactAggregateRequest['aggregateKind']
}

interface CriterionExactAggregateBucket {
  count: number
  sum: number
  numericCount: number
  minimum: number
  maximum: number
  firstError?: CellValue
}

interface CriterionExactAggregateIndex {
  readonly numbers: ReadonlyMap<number, CriterionExactAggregateBucket>
  readonly strings: ReadonlyMap<string, CriterionExactAggregateBucket>
}

interface CriterionEqualityRowIndex {
  readonly numbers: ReadonlyMap<number, Uint32Array>
  readonly strings: ReadonlyMap<string, Uint32Array>
}

type CriterionEqualityIndexKey =
  | {
      readonly kind: 'number'
      readonly value: number
    }
  | {
      readonly kind: 'string'
      readonly value: string
    }

type SliceFastPredicate =
  | {
      kind: 'eq-empty'
      negate: boolean
    }
  | {
      kind: 'eq-bool'
      negate: boolean
      value: boolean
    }
  | {
      kind: 'eq-number'
      negate: boolean
      value: number
    }
  | {
      kind: 'eq-string'
      negate: boolean
      value: string
    }
  | {
      kind: 'cmp-number'
      operator: Exclude<CriteriaOperator, '=' | '<>'>
      value: number
    }
  | {
      kind: 'generic'
      compiled: CompiledCriteriaMatcher
    }

const equalityRowIndexes = new WeakMap<RuntimeColumnView['owner'], CriterionEqualityRowIndex>()

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

function numberValue(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function criteriaCacheKey(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return 'e:'
    case ValueTag.Number:
      return `n:${normalizeExactLookupNumber(value.value)}`
    case ValueTag.Boolean:
      return value.value ? 'b:1' : 'b:0'
    case ValueTag.String:
      return `s:${value.value}`
    case ValueTag.Error:
      return `r:${value.code}`
  }
}

function decodeValueTag(rawTag: number | undefined): ValueTag {
  if (rawTag === undefined) {
    return ValueTag.Empty
  }
  switch (rawTag) {
    case 1:
      return ValueTag.Number
    case 2:
      return ValueTag.Boolean
    case 3:
      return ValueTag.String
    case 4:
      return ValueTag.Error
    case 0:
    default:
      return ValueTag.Empty
  }
}

function normalizeSliceString(runtimeColumnStore: EngineRuntimeColumnStoreService, stringId: number): string {
  return stringId === 0 ? '' : runtimeColumnStore.normalizeStringId(stringId)
}

function materializeSliceValue(view: RuntimeColumnView, offset: number, runtimeColumnStore: EngineRuntimeColumnStoreService): CellValue {
  const tag = decodeValueTag(view.readTagAt(offset))
  switch (tag) {
    case ValueTag.Empty:
      return { tag: ValueTag.Empty }
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: view.readNumberAt(offset) }
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: view.readNumberAt(offset) !== 0 }
    case ValueTag.String: {
      const stringId = view.readStringIdAt(offset)
      return {
        tag: ValueTag.String,
        value: stringId === 0 ? '' : runtimeColumnStore.normalizeStringId(stringId),
        stringId,
      }
    }
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: view.readErrorAt(offset) ?? ErrorCode.None }
  }
}

function buildSlicePredicate(compiled: CompiledCriteriaMatcher): SliceFastPredicate {
  const { operator, operand, wildcardPattern } = compiled
  if (wildcardPattern) {
    return { kind: 'generic', compiled }
  }
  if (operator === '=' || operator === '<>') {
    const negate = operator === '<>'
    switch (operand.tag) {
      case ValueTag.Empty:
        return { kind: 'eq-empty', negate }
      case ValueTag.Boolean:
        return { kind: 'eq-bool', negate, value: operand.value }
      case ValueTag.Number:
        return {
          kind: 'eq-number',
          negate,
          value: normalizeExactLookupNumber(operand.value),
        }
      case ValueTag.String:
        return { kind: 'eq-string', negate, value: operand.value.toUpperCase() }
      case ValueTag.Error:
        return { kind: 'generic', compiled }
    }
  }
  if (operand.tag === ValueTag.Number) {
    return {
      kind: 'cmp-number',
      operator,
      value: normalizeExactLookupNumber(operand.value),
    }
  }
  return { kind: 'generic', compiled }
}

function slicePredicateMatches(
  predicate: SliceFastPredicate,
  view: RuntimeColumnView,
  offset: number,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): boolean {
  switch (predicate.kind) {
    case 'eq-empty': {
      const tag = decodeValueTag(view.readTagAt(offset))
      const matches =
        tag === ValueTag.Empty || (tag === ValueTag.String && normalizeSliceString(runtimeColumnStore, view.readStringIdAt(offset)) === '')
      return predicate.negate ? !matches : matches
    }
    case 'eq-bool': {
      const tag = decodeValueTag(view.readTagAt(offset))
      const numeric = tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.Empty ? view.readNumberAt(offset) : undefined
      const matches = numeric !== undefined && (Object.is(numeric, -0) ? 0 : numeric) === (predicate.value ? 1 : 0)
      return predicate.negate ? !matches : matches
    }
    case 'eq-number': {
      const tag = decodeValueTag(view.readTagAt(offset))
      const numeric = normalizeExactLookupNumber(view.readNumberAt(offset))
      const matches = (tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.Empty) && numeric === predicate.value
      return predicate.negate ? !matches : matches
    }
    case 'eq-string': {
      const tag = decodeValueTag(view.readTagAt(offset))
      const matches =
        (tag === ValueTag.String || tag === ValueTag.Empty) &&
        (tag === ValueTag.Empty ? '' : normalizeSliceString(runtimeColumnStore, view.readStringIdAt(offset))) === predicate.value
      return predicate.negate ? !matches : matches
    }
    case 'cmp-number': {
      const tag = decodeValueTag(view.readTagAt(offset))
      if (tag !== ValueTag.Number && tag !== ValueTag.Boolean) {
        return false
      }
      const numeric = normalizeExactLookupNumber(view.readNumberAt(offset))
      switch (predicate.operator) {
        case '>':
          return numeric > predicate.value
        case '>=':
          return numeric >= predicate.value
        case '<':
          return numeric < predicate.value
        case '<=':
          return numeric <= predicate.value
        default:
          return false
      }
    }
    case 'generic':
      return matchesCompiledCriteria(materializeSliceValue(view, offset, runtimeColumnStore), predicate.compiled)
  }
}

function slicePredicateMatchesEmpty(predicate: SliceFastPredicate): boolean {
  switch (predicate.kind) {
    case 'eq-empty':
      return !predicate.negate
    case 'eq-bool':
      return predicate.negate ? predicate.value : !predicate.value
    case 'eq-number':
      return predicate.negate ? predicate.value !== 0 : predicate.value === 0
    case 'eq-string':
      return predicate.negate ? predicate.value !== '' : predicate.value === ''
    case 'cmp-number':
      return false
    case 'generic':
      return matchesCompiledCriteria({ tag: ValueTag.Empty }, predicate.compiled)
  }
}

function equalityIndexKeyForPredicate(predicate: SliceFastPredicate): CriterionEqualityIndexKey | undefined {
  if (predicate.kind === 'eq-number' && !predicate.negate && predicate.value !== 0) {
    return { kind: 'number', value: predicate.value }
  }
  if (predicate.kind === 'eq-bool' && !predicate.negate && predicate.value) {
    return { kind: 'number', value: 1 }
  }
  if (predicate.kind === 'eq-string' && !predicate.negate && predicate.value !== '') {
    return { kind: 'string', value: predicate.value }
  }
  return undefined
}

function equalityIndexKeyForStoredValue(
  tag: ValueTag,
  number: number,
  stringId: number,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): CriterionEqualityIndexKey | undefined {
  if (tag === ValueTag.Number) {
    const value = normalizeExactLookupNumber(number)
    return value === 0 ? undefined : { kind: 'number', value }
  }
  if (tag === ValueTag.Boolean) {
    return number === 0 ? undefined : { kind: 'number', value: 1 }
  }
  if (tag === ValueTag.String) {
    const value = normalizeSliceString(runtimeColumnStore, stringId)
    return value === '' ? undefined : { kind: 'string', value }
  }
  return undefined
}

function sortedUint32Rows(rows: number[]): Uint32Array {
  if (rows.length === 0) {
    return new Uint32Array(0)
  }
  rows.sort((left, right) => left - right)
  return Uint32Array.from(rows)
}

function getOrBuildEqualityRowIndex(
  view: RuntimeColumnView,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): CriterionEqualityRowIndex {
  const cached = equalityRowIndexes.get(view.owner)
  if (cached !== undefined) {
    return cached
  }

  const numbers = new Map<number, number[]>()
  const strings = new Map<string, number[]>()
  view.owner.pages.forEach((page) => {
    if (page.nonEmptyCount === 0) {
      return
    }
    const rowEnd = page.rowStart + page.tags.length - 1
    for (let row = page.rowStart; row <= rowEnd; row += 1) {
      const localRow = row - page.rowStart
      const tag = decodeValueTag(page.tags[localRow])
      const key = equalityIndexKeyForStoredValue(tag, page.numbers[localRow] ?? 0, page.stringIds[localRow] ?? 0, runtimeColumnStore)
      if (key?.kind === 'number') {
        let rows = numbers.get(key.value)
        if (rows === undefined) {
          rows = []
          numbers.set(key.value, rows)
        }
        rows.push(row)
        continue
      }
      if (key?.kind === 'string') {
        let rows = strings.get(key.value)
        if (rows === undefined) {
          rows = []
          strings.set(key.value, rows)
        }
        rows.push(row)
      }
    }
  })

  const index: CriterionEqualityRowIndex = {
    numbers: new Map([...numbers].map(([value, rows]) => [value, sortedUint32Rows(rows)])),
    strings: new Map([...strings].map(([value, rows]) => [value, sortedUint32Rows(rows)])),
  }
  equalityRowIndexes.set(view.owner, index)
  return index
}

function readIndexedEqualityRows(
  view: RuntimeColumnView,
  predicate: SliceFastPredicate,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): CriterionRowSetView | undefined {
  const key = equalityIndexKeyForPredicate(predicate)
  if (key === undefined) {
    return undefined
  }
  const index = getOrBuildEqualityRowIndex(view, runtimeColumnStore)
  const rows = key.kind === 'number' ? index.numbers.get(key.value) : index.strings.get(key.value)
  return sliceAbsoluteRowsToRangeView(rows ?? new Uint32Array(0), view.rowStart, view.rowEnd)
}

function createExactAggregateBucket(): CriterionExactAggregateBucket {
  return {
    count: 0,
    sum: 0,
    numericCount: 0,
    minimum: Number.POSITIVE_INFINITY,
    maximum: Number.NEGATIVE_INFINITY,
  }
}

function addExactAggregateMatch(bucket: CriterionExactAggregateBucket, aggregateView: RuntimeColumnView | undefined, offset: number): void {
  bucket.count += 1
  if (aggregateView === undefined) {
    return
  }
  const tag = decodeValueTag(aggregateView.readTagAt(offset))
  if (tag === ValueTag.Error) {
    bucket.firstError ??= { tag: ValueTag.Error, code: aggregateView.readErrorAt(offset) ?? ErrorCode.None }
    return
  }
  if (tag === ValueTag.Number) {
    const numeric = aggregateView.readNumberAt(offset)
    bucket.sum += numeric
    bucket.numericCount += 1
    bucket.minimum = Math.min(bucket.minimum, numeric)
    bucket.maximum = Math.max(bucket.maximum, numeric)
    return
  }
  if (tag === ValueTag.Boolean) {
    bucket.sum += aggregateView.readNumberAt(offset) !== 0 ? 1 : 0
    bucket.numericCount += 1
    return
  }
  if (tag === ValueTag.Empty) {
    bucket.numericCount += 1
  }
}

function addExactAggregateBucket(
  index: { numbers: Map<number, CriterionExactAggregateBucket>; strings: Map<string, CriterionExactAggregateBucket> },
  key: CriterionEqualityIndexKey,
  aggregateView: RuntimeColumnView | undefined,
  offset: number,
): void {
  let bucket = key.kind === 'number' ? index.numbers.get(key.value) : index.strings.get(key.value)
  if (bucket === undefined) {
    bucket = createExactAggregateBucket()
    if (key.kind === 'number') {
      index.numbers.set(key.value, bucket)
    } else {
      index.strings.set(key.value, bucket)
    }
  }
  addExactAggregateMatch(bucket, aggregateView, offset)
}

function exactAggregateValueFromBucket(
  bucket: CriterionExactAggregateBucket | undefined,
  aggregateKind: CriterionExactAggregateRequest['aggregateKind'],
): CellValue {
  if (bucket === undefined) {
    return aggregateKind === 'average' ? errorValue(ErrorCode.Div0) : numberValue(0)
  }
  if (aggregateKind === 'count') {
    return numberValue(bucket.count)
  }
  if (bucket.firstError) {
    return bucket.firstError
  }
  if (aggregateKind === 'sum') {
    return numberValue(bucket.sum)
  }
  if (aggregateKind === 'average') {
    return bucket.numericCount === 0 ? errorValue(ErrorCode.Div0) : numberValue(bucket.sum / bucket.numericCount)
  }
  if (aggregateKind === 'min') {
    return numberValue(bucket.minimum === Number.POSITIVE_INFINITY ? 0 : bucket.minimum)
  }
  return numberValue(bucket.maximum === Number.NEGATIVE_INFINITY ? 0 : bucket.maximum)
}

function exactAggregateBucketValue(
  index: CriterionExactAggregateIndex,
  key: CriterionEqualityIndexKey,
  aggregateKind: CriterionExactAggregateRequest['aggregateKind'],
): CellValue {
  return exactAggregateValueFromBucket(key.kind === 'number' ? index.numbers.get(key.value) : index.strings.get(key.value), aggregateKind)
}

function exactAggregateIndexCacheKey(request: {
  readonly criteriaRegionId: number
  readonly criteriaView: RuntimeColumnView
  readonly aggregateRegionId?: number
  readonly aggregateView?: RuntimeColumnView
}): string {
  return [
    request.criteriaRegionId,
    request.criteriaView.columnVersion,
    request.criteriaView.structureVersion,
    request.aggregateRegionId ?? -1,
    request.aggregateView?.columnVersion ?? -1,
    request.aggregateView?.structureVersion ?? -1,
  ].join('\u0001')
}

function exactKeyCachePart(key: CriterionEqualityIndexKey): string {
  return key.kind === 'number' ? `n:${key.value}` : `s:${key.value}`
}

function compoundExactAggregateCacheKey(request: {
  readonly aggregateKind: CriterionExactAggregateRequest['aggregateKind']
  readonly criteriaPairs: readonly {
    readonly regionId: number
    readonly view: RuntimeColumnView
    readonly key: CriterionEqualityIndexKey
  }[]
  readonly aggregateRegionId?: number
  readonly aggregateView?: RuntimeColumnView
}): string {
  return [
    request.aggregateKind,
    request.criteriaPairs
      .map(({ regionId, view, key }) => `${regionId}:${view.columnVersion}:${view.structureVersion}:${exactKeyCachePart(key)}`)
      .join('\u0002'),
    request.aggregateRegionId ?? -1,
    request.aggregateView?.columnVersion ?? -1,
    request.aggregateView?.structureVersion ?? -1,
  ].join('\u0001')
}

function countNonEmptyRowsInView(view: RuntimeColumnView): number {
  let count = 0
  view.owner.pages.forEach((page) => {
    const rowStart = Math.max(view.rowStart, page.rowStart)
    const rowEnd = Math.min(view.rowEnd, page.rowStart + page.tags.length - 1)
    if (rowStart > rowEnd || page.nonEmptyCount === 0) {
      return
    }
    if (rowStart === page.rowStart && rowEnd === page.rowStart + page.tags.length - 1) {
      count += page.nonEmptyCount
      return
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      if (decodeValueTag(page.tags[row - page.rowStart]) !== ValueTag.Empty) {
        count += 1
      }
    }
  })
  return count
}

function forEachNonEmptyRowOffsetInView(view: RuntimeColumnView, fn: (rowOffset: number) => void): void {
  view.owner.pages.forEach((page) => {
    const rowStart = Math.max(view.rowStart, page.rowStart)
    const rowEnd = Math.min(view.rowEnd, page.rowStart + page.tags.length - 1)
    if (rowStart > rowEnd || page.nonEmptyCount === 0) {
      return
    }
    if (rowStart === page.rowStart && rowEnd === page.rowStart + page.tags.length - 1 && page.nonEmptyCount === page.tags.length) {
      for (let row = rowStart; row <= rowEnd; row += 1) {
        fn(row - view.rowStart)
      }
      return
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      if (decodeValueTag(page.tags[row - page.rowStart]) !== ValueTag.Empty) {
        fn(row - view.rowStart)
      }
    }
  })
}

export function createCriterionRangeCacheService(args: {
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  readonly regionGraph: Pick<RegionGraph, 'internSingleColumnRegion'>
  readonly depPatternStore: DepPatternStore
}): CriterionRangeCacheService {
  const exactAggregateIndexes = new Map<string, CriterionExactAggregateIndex>()
  const compoundExactAggregateResults = new Map<string, CellValue>()

  const getColumnView = (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }): RuntimeColumnView => {
    const direct = Reflect.get(args.runtimeColumnStore, 'getColumnView')
    if (typeof direct === 'function') {
      return direct.call(args.runtimeColumnStore, request)
    }
    const slice = args.runtimeColumnStore.getColumnSlice(request)
    return {
      owner: {
        sheetName: slice.sheetName,
        col: slice.col,
        columnVersion: slice.columnVersion,
        structureVersion: slice.structureVersion,
        sheetColumnVersions: slice.sheetColumnVersions,
        pages: new Map(),
      },
      sheetName: slice.sheetName,
      rowStart: slice.rowStart,
      rowEnd: slice.rowEnd,
      col: slice.col,
      length: slice.length,
      columnVersion: slice.columnVersion,
      structureVersion: slice.structureVersion,
      sheetColumnVersions: slice.sheetColumnVersions,
      readTagAt(offset) {
        return slice.tags[offset] ?? ValueTag.Empty
      },
      readNumberAt(offset) {
        return slice.numbers[offset] ?? 0
      },
      readStringIdAt(offset) {
        return slice.stringIds[offset] ?? 0
      },
      readErrorAt(offset) {
        return slice.errors[offset] ?? ErrorCode.None
      },
      readCellValueAt(offset) {
        return materializeSliceValue(this, offset, args.runtimeColumnStore)
      },
    }
  }

  const rangeRegionId = (range: CriterionRangeDescriptor): number =>
    args.regionGraph.internSingleColumnRegion({
      sheetName: range.sheetName,
      rowStart: range.rowStart,
      rowEnd: range.rowEnd,
      col: range.col,
    })

  const buildExactAggregateIndex = (
    criteriaView: RuntimeColumnView,
    aggregateView: RuntimeColumnView | undefined,
  ): CriterionExactAggregateIndex => {
    const buckets = {
      numbers: new Map<number, CriterionExactAggregateBucket>(),
      strings: new Map<string, CriterionExactAggregateBucket>(),
    }
    criteriaView.owner.pages.forEach((page) => {
      const rowStart = Math.max(criteriaView.rowStart, page.rowStart)
      const rowEnd = Math.min(criteriaView.rowEnd, page.rowStart + page.tags.length - 1)
      if (rowStart > rowEnd || page.nonEmptyCount === 0) {
        return
      }
      for (let row = rowStart; row <= rowEnd; row += 1) {
        const localRow = row - page.rowStart
        const key = equalityIndexKeyForStoredValue(
          decodeValueTag(page.tags[localRow]),
          page.numbers[localRow] ?? 0,
          page.stringIds[localRow] ?? 0,
          args.runtimeColumnStore,
        )
        if (key === undefined) {
          continue
        }
        addExactAggregateBucket(buckets, key, aggregateView, row - criteriaView.rowStart)
      }
    })
    return buckets
  }

  const getOrBuildExactAggregate = (request: CriterionExactAggregateRequest): CellValue | undefined => {
    const criteriaPredicate = buildSlicePredicate(compileCriteriaMatcher(request.criteriaPair.criteria))
    const criteriaKey = equalityIndexKeyForPredicate(criteriaPredicate)
    if (criteriaKey === undefined) {
      return undefined
    }
    if (request.aggregateKind !== 'count' && request.aggregateRange === undefined) {
      return undefined
    }
    if (request.aggregateRange !== undefined && request.aggregateRange.length !== request.criteriaPair.range.length) {
      return undefined
    }
    const criteriaRegionId = rangeRegionId(request.criteriaPair.range)
    const criteriaView = getColumnView(request.criteriaPair.range)
    const aggregateRegionId = request.aggregateRange === undefined ? undefined : rangeRegionId(request.aggregateRange)
    const aggregateView = request.aggregateRange === undefined ? undefined : getColumnView(request.aggregateRange)
    const cacheKey = exactAggregateIndexCacheKey({
      criteriaRegionId,
      criteriaView,
      ...(aggregateRegionId === undefined ? {} : { aggregateRegionId }),
      ...(aggregateView === undefined ? {} : { aggregateView }),
    })
    let index = exactAggregateIndexes.get(cacheKey)
    if (index === undefined) {
      index = buildExactAggregateIndex(criteriaView, aggregateView)
      exactAggregateIndexes.set(cacheKey, index)
    }
    return exactAggregateBucketValue(index, criteriaKey, request.aggregateKind)
  }

  const getOrBuildCompoundExactAggregate = (request: CriterionCompoundExactAggregateRequest): CellValue | undefined => {
    const { criteriaPairs } = request
    if (criteriaPairs.length < 2) {
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

    const predicates = criteriaPairs.map((pair) => buildSlicePredicate(compileCriteriaMatcher(pair.criteria)))
    const keys = predicates.map(equalityIndexKeyForPredicate)
    if (keys.some((key) => key === undefined)) {
      return undefined
    }

    const resolvedPairs = criteriaPairs.map((pair, index) => ({
      regionId: rangeRegionId(pair.range),
      view: getColumnView(pair.range),
      predicate: predicates[index]!,
      key: keys[index]!,
    }))
    const aggregateRegionId = request.aggregateRange === undefined ? undefined : rangeRegionId(request.aggregateRange)
    const aggregateView = request.aggregateRange === undefined ? undefined : getColumnView(request.aggregateRange)
    const cacheKey = compoundExactAggregateCacheKey({
      aggregateKind: request.aggregateKind,
      criteriaPairs: resolvedPairs,
      ...(aggregateRegionId === undefined ? {} : { aggregateRegionId }),
      ...(aggregateView === undefined ? {} : { aggregateView }),
    })
    const cached = compoundExactAggregateResults.get(cacheKey)
    if (cached !== undefined) {
      return cached
    }

    const rowsets = resolvedPairs.map(({ view, predicate }) => readIndexedEqualityRows(view, predicate, args.runtimeColumnStore))
    if (rowsets.some((rowset) => rowset === undefined)) {
      return undefined
    }

    let limitingPairIndex = 0
    let limitingRows = rowsets[0]!
    for (let index = 1; index < rowsets.length; index += 1) {
      const rowset = rowsets[index]!
      if (rowset.cardinality < limitingRows.cardinality) {
        limitingPairIndex = index
        limitingRows = rowset
      }
    }

    const bucket = createExactAggregateBucket()
    limitingRows.forEachOffset((rowOffset) => {
      for (let pairIndex = 0; pairIndex < resolvedPairs.length; pairIndex += 1) {
        if (pairIndex === limitingPairIndex) {
          continue
        }
        const pair = resolvedPairs[pairIndex]!
        if (!slicePredicateMatches(pair.predicate, pair.view, rowOffset, args.runtimeColumnStore)) {
          return
        }
      }
      addExactAggregateMatch(bucket, aggregateView, rowOffset)
    })

    const result = exactAggregateValueFromBucket(bucket.count === 0 ? undefined : bucket, request.aggregateKind)
    compoundExactAggregateResults.set(cacheKey, result)
    return result
  }

  const getOrBuildExactCriteriaAggregate = (request: CriterionCompoundExactAggregateRequest): CellValue | undefined => {
    return request.criteriaPairs.length === 1
      ? getOrBuildExactAggregate({
          criteriaPair: request.criteriaPairs[0]!,
          ...(request.aggregateRange === undefined ? {} : { aggregateRange: request.aggregateRange }),
          aggregateKind: request.aggregateKind,
        })
      : getOrBuildCompoundExactAggregate(request)
  }

  const getOrBuildMatchingRows = (request: { criteriaPairs: readonly CriterionRangePair[] }): CriterionRangeMatch | CellValue => {
    const { criteriaPairs } = request
    if (criteriaPairs.length === 0) {
      return errorValue(ErrorCode.Value)
    }
    const expectedLength = criteriaPairs[0]!.range.length
    if (criteriaPairs.some((pair) => pair.range.length !== expectedLength)) {
      return errorValue(ErrorCode.Value)
    }

    const resolvedPairs = criteriaPairs.map((pair) => ({
      regionId: args.regionGraph.internSingleColumnRegion({
        sheetName: pair.range.sheetName,
        rowStart: pair.range.rowStart,
        rowEnd: pair.range.rowEnd,
        col: pair.range.col,
      }),
      view: getColumnView({
        sheetName: pair.range.sheetName,
        rowStart: pair.range.rowStart,
        rowEnd: pair.range.rowEnd,
        col: pair.range.col,
      }),
      criteria: pair.criteria,
    }))
    const versionStamp = resolvedPairs
      .map(({ regionId, view }) => `${regionId}:${view.columnVersion}:${view.structureVersion}`)
      .join('\u0001')
    const existing = args.depPatternStore.getCriteriaPattern({
      regionIds: resolvedPairs.map(({ regionId }) => regionId),
      criteriaKeys: resolvedPairs.map(({ criteria }) => criteriaCacheKey(criteria)),
      versionStamp,
    })
    if (existing) {
      return existing
    }

    const predicates = criteriaPairs.map((pair) => buildSlicePredicate(compileCriteriaMatcher(pair.criteria)))
    const matchingRows: number[] = []
    let limitingPairIndex: number | undefined
    let limitingIndexedRows: CriterionRowSetView | undefined
    let limitingPairNonEmptyRows = Number.POSITIVE_INFINITY
    for (let pairIndex = 0; pairIndex < predicates.length; pairIndex += 1) {
      const indexedRows = readIndexedEqualityRows(resolvedPairs[pairIndex]!.view, predicates[pairIndex]!, args.runtimeColumnStore)
      if (indexedRows !== undefined) {
        if (indexedRows.cardinality < limitingPairNonEmptyRows) {
          limitingPairIndex = undefined
          limitingIndexedRows = indexedRows
          limitingPairNonEmptyRows = indexedRows.cardinality
        }
        continue
      }
      if (slicePredicateMatchesEmpty(predicates[pairIndex]!)) {
        continue
      }
      const view = resolvedPairs[pairIndex]!.view
      if (view.owner.pages.size === 0) {
        continue
      }
      const nonEmptyRows = countNonEmptyRowsInView(view)
      if (nonEmptyRows < limitingPairNonEmptyRows) {
        limitingPairIndex = pairIndex
        limitingIndexedRows = undefined
        limitingPairNonEmptyRows = nonEmptyRows
      }
    }
    const visitCandidate = (rowOffset: number): void => {
      let matches = true
      for (let pairIndex = 0; pairIndex < predicates.length; pairIndex += 1) {
        if (!slicePredicateMatches(predicates[pairIndex]!, resolvedPairs[pairIndex]!.view, rowOffset, args.runtimeColumnStore)) {
          matches = false
          break
        }
      }
      if (matches) {
        matchingRows.push(rowOffset)
      }
    }
    if (limitingIndexedRows !== undefined) {
      limitingIndexedRows.forEachOffset(visitCandidate)
    } else if (limitingPairIndex === undefined) {
      for (let rowOffset = 0; rowOffset < expectedLength; rowOffset += 1) {
        visitCandidate(rowOffset)
      }
    } else {
      forEachNonEmptyRowOffsetInView(resolvedPairs[limitingPairIndex]!.view, visitCandidate)
    }

    const entry: CriterionCacheEntry = {
      rows: Uint32Array.from(matchingRows),
      length: matchingRows.length,
      pairVersions: resolvedPairs.map(({ view }) => ({
        columnVersion: view.columnVersion,
        structureVersion: view.structureVersion,
      })),
    }
    return args.depPatternStore.setCriteriaPattern({
      regionIds: resolvedPairs.map(({ regionId }) => regionId),
      criteriaKeys: resolvedPairs.map(({ criteria }) => criteriaCacheKey(criteria)),
      versionStamp,
      rows: entry.rows,
      length: entry.length,
    })
  }

  return {
    getOrBuildCompoundExactAggregate,
    getOrBuildExactCriteriaAggregate,
    getOrBuildExactAggregate,
    getOrBuildMatchingRows,
  }
}
