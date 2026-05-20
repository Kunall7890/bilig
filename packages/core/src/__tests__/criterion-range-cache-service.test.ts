import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { createDepPatternStore } from '../deps/dep-pattern-store.js'
import { createRegionGraph } from '../deps/region-graph.js'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'
import { createCriterionRangeCacheService } from '../engine/services/criterion-range-cache-service.js'
import { createEngineRuntimeColumnStoreService } from '../engine/services/runtime-column-store-service.js'

function setStoredCellValue(
  workbook: WorkbookStore,
  strings: StringPool,
  sheetName: string,
  address: string,
  value: { tag: ValueTag.String; value: string } | { tag: ValueTag.Number; value: number } | { tag: ValueTag.Boolean; value: boolean },
): void {
  const cellIndex = workbook.ensureCell(sheetName, address)
  workbook.cellStore.setValue(cellIndex, value, value.tag === ValueTag.String ? strings.intern(value.value) : 0)
}

function isErrorValue(value: unknown): value is Extract<CellValue, { tag: ValueTag.Error }> {
  return typeof value === 'object' && value !== null && 'tag' in value && value.tag === ValueTag.Error && 'code' in value
}

describe('createCriterionRangeCacheService', () => {
  it('rejects empty criteria requests', () => {
    const workbook = new WorkbookStore('criteria-cache-empty')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    expect(criterionCache.getOrBuildMatchingRows({ criteriaPairs: [] })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('reuses matching row sets for identical requests and invalidates on source writes', () => {
    const workbook = new WorkbookStore('criteria-cache')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;['A', 'B', 'A', 'B', 'A', 'B'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    const first = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: 'A', stringId: strings.intern('A') },
        },
      ],
    })
    if (isErrorValue(first)) {
      throw new Error(`unexpected criteria cache error: ${first.code}`)
    }
    expect([...first.rows]).toEqual([0, 2, 4])

    const second = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: 'A', stringId: strings.intern('A') },
        },
      ],
    })
    if (isErrorValue(second)) {
      throw new Error(`unexpected criteria cache error: ${second.code}`)
    }
    expect(second.rows).toBe(first.rows)

    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', {
      tag: ValueTag.String,
      value: 'A',
    })

    const third = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: 'A', stringId: strings.intern('A') },
        },
      ],
    })
    if (isErrorValue(third)) {
      throw new Error(`unexpected criteria cache error: ${third.code}`)
    }
    expect(third.rows).not.toBe(first.rows)
    expect([...third.rows]).toEqual([0, 1, 2, 4])
  })

  it('uses the equality row index for uncached exact criteria values', () => {
    const workbook = new WorkbookStore('criteria-cache-equality-index')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;['A', 'B', 'A', 'B', 'A', 'B'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })

    const realRuntimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    let readTagCalls = 0
    const runtimeColumnStore = {
      ...realRuntimeColumnStore,
      getColumnView(request: Parameters<typeof realRuntimeColumnStore.getColumnView>[0]) {
        const view = realRuntimeColumnStore.getColumnView(request)
        return {
          ...view,
          readTagAt(offset: number) {
            readTagCalls += 1
            return view.readTagAt(offset)
          },
        }
      },
    }
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    const first = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: 'A', stringId: strings.intern('A') },
        },
      ],
    })
    if (isErrorValue(first)) {
      throw new Error(`unexpected criteria cache error: ${first.code}`)
    }
    expect([...first.rows]).toEqual([0, 2, 4])

    readTagCalls = 0
    const second = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: 'B', stringId: strings.intern('B') },
        },
      ],
    })
    if (isErrorValue(second)) {
      throw new Error(`unexpected criteria cache error: ${second.code}`)
    }
    expect([...second.rows]).toEqual([1, 3, 5])
    expect(readTagCalls).toBe(3)
  })

  it('uses the most selective equality row index for multi-criteria requests', () => {
    const workbook = new WorkbookStore('criteria-cache-multi-equality-index')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;['A', 'A', 'A', 'A', 'A', 'A'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })
    ;['skip', 'target', 'skip', 'skip', 'skip', 'skip'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `B${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })

    const realRuntimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    let readTagCalls = 0
    const runtimeColumnStore = {
      ...realRuntimeColumnStore,
      getColumnView(request: Parameters<typeof realRuntimeColumnStore.getColumnView>[0]) {
        const view = realRuntimeColumnStore.getColumnView(request)
        return {
          ...view,
          readTagAt(offset: number) {
            readTagCalls += 1
            return view.readTagAt(offset)
          },
        }
      },
    }
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    const matching = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: 'A', stringId: strings.intern('A') },
        },
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 },
          criteria: { tag: ValueTag.String, value: 'target', stringId: strings.intern('target') },
        },
      ],
    })
    if (isErrorValue(matching)) {
      throw new Error(`unexpected criteria cache error: ${matching.code}`)
    }
    expect([...matching.rows]).toEqual([1])
    expect(readTagCalls).toBe(2)
  })

  it('computes exact criteria count and sum aggregates from a grouped column pass', () => {
    const workbook = new WorkbookStore('criteria-cache-exact-aggregate')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;['A', 'B', 'A', 'B', 'A', 'B'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
      setStoredCellValue(workbook, strings, 'Sheet1', `B${index + 1}`, {
        tag: ValueTag.Number,
        value: index + 1,
      })
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })
    const criteriaPair = {
      range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
      criteria: { tag: ValueTag.String, value: 'B', stringId: strings.intern('B') },
    } as const

    expect(
      criterionCache.getOrBuildExactAggregate({
        criteriaPair,
        aggregateKind: 'count',
      }),
    ).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(
      criterionCache.getOrBuildExactAggregate({
        criteriaPair,
        aggregateRange: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 },
        aggregateKind: 'sum',
      }),
    ).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(
      criterionCache.getOrBuildExactAggregate({
        criteriaPair,
        aggregateRange: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 },
        aggregateKind: 'average',
      }),
    ).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(
      criterionCache.getOrBuildExactAggregate({
        criteriaPair,
        aggregateRange: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 },
        aggregateKind: 'min',
      }),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      criterionCache.getOrBuildExactAggregate({
        criteriaPair,
        aggregateRange: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 },
        aggregateKind: 'max',
      }),
    ).toEqual({ tag: ValueTag.Number, value: 6 })

    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', {
      tag: ValueTag.String,
      value: 'A',
    })

    expect(
      criterionCache.getOrBuildExactAggregate({
        criteriaPair,
        aggregateRange: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 },
        aggregateKind: 'sum',
      }),
    ).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(
      criterionCache.getOrBuildExactAggregate({
        criteriaPair,
        aggregateRange: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 },
        aggregateKind: 'average',
      }),
    ).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(
      criterionCache.getOrBuildExactAggregate({
        criteriaPair,
        aggregateRange: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 },
        aggregateKind: 'min',
      }),
    ).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(
      criterionCache.getOrBuildExactAggregate({
        criteriaPair: {
          ...criteriaPair,
          criteria: { tag: ValueTag.String, value: 'missing', stringId: strings.intern('missing') },
        },
        aggregateRange: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 },
        aggregateKind: 'average',
      }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
  })

  it('aggregates compound exact criteria from the most selective equality rowset', () => {
    const workbook = new WorkbookStore('criteria-cache-compound-exact-aggregate')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;['A', 'A', 'A', 'A', 'A', 'A', 'A', 'A'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })
    ;['skip', 'target', 'skip', 'skip', 'skip', 'skip', 'target', 'skip'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `B${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })
    ;[5, 20, 7, 9, 11, 13, 30, 17].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `C${index + 1}`, {
        tag: ValueTag.Number,
        value,
      })
    })

    const realRuntimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    let readTagCalls = 0
    const runtimeColumnStore = {
      ...realRuntimeColumnStore,
      getColumnView(request: Parameters<typeof realRuntimeColumnStore.getColumnView>[0]) {
        const view = realRuntimeColumnStore.getColumnView(request)
        return {
          ...view,
          readTagAt(offset: number) {
            readTagCalls += 1
            return view.readTagAt(offset)
          },
        }
      },
    }
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })
    const criteriaPairs = [
      {
        range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 7, col: 0, length: 8 },
        criteria: { tag: ValueTag.String, value: 'A', stringId: strings.intern('A') },
      },
      {
        range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 7, col: 1, length: 8 },
        criteria: { tag: ValueTag.String, value: 'target', stringId: strings.intern('target') },
      },
    ] as const

    expect(
      criterionCache.getOrBuildCompoundExactAggregate({
        criteriaPairs,
        aggregateRange: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 7, col: 2, length: 8 },
        aggregateKind: 'sum',
      }),
    ).toEqual({ tag: ValueTag.Number, value: 50 })
    expect(readTagCalls).toBeLessThan(8)

    expect(
      criterionCache.getOrBuildCompoundExactAggregate({
        criteriaPairs: [
          criteriaPairs[0],
          {
            ...criteriaPairs[1],
            criteria: { tag: ValueTag.String, value: 'missing', stringId: strings.intern('missing') },
          },
        ],
        aggregateRange: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 7, col: 2, length: 8 },
        aggregateKind: 'average',
      }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
  })

  it('reuses compound exact aggregate buckets across different criteria tuples', () => {
    const workbook = new WorkbookStore('criteria-cache-compound-exact-buckets')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;['East', 'East', 'West', 'West', 'East', 'West'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })
    ;['Retail', 'Wholesale', 'Retail', 'Wholesale', 'Retail', 'Retail'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `B${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })
    ;[10, 20, 30, 40, 50, 60].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `C${index + 1}`, {
        tag: ValueTag.Number,
        value,
      })
    })

    const realRuntimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    let readTagCalls = 0
    const runtimeColumnStore = {
      ...realRuntimeColumnStore,
      getColumnView(request: Parameters<typeof realRuntimeColumnStore.getColumnView>[0]) {
        const view = realRuntimeColumnStore.getColumnView(request)
        return {
          ...view,
          readTagAt(offset: number) {
            readTagCalls += 1
            return view.readTagAt(offset)
          },
        }
      },
    }
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })
    const criteriaRangeA = { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 } as const
    const criteriaRangeB = { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 } as const
    const aggregateRange = { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 2, length: 6 } as const
    const request = (
      first: string,
      second: string,
    ): Parameters<typeof criterionCache.getOrBuildCompoundExactAggregate>[0] & { readonly useCompoundBucketIndex: true } => ({
      criteriaPairs: [
        {
          range: criteriaRangeA,
          criteria: { tag: ValueTag.String, value: first, stringId: strings.intern(first) },
        },
        {
          range: criteriaRangeB,
          criteria: { tag: ValueTag.String, value: second, stringId: strings.intern(second) },
        },
      ],
      aggregateRange,
      aggregateKind: 'sum',
      useCompoundBucketIndex: true,
    })

    expect(criterionCache.getOrBuildCompoundExactAggregate(request('East', 'Retail'))).toEqual({
      tag: ValueTag.Number,
      value: 60,
    })
    const callsAfterFirstTuple = readTagCalls

    expect(criterionCache.getOrBuildCompoundExactAggregate(request('West', 'Retail'))).toEqual({
      tag: ValueTag.Number,
      value: 90,
    })
    expect(readTagCalls).toBe(callsAfterFirstTuple)
  })

  it('supports compiled operator criteria and validates matching range lengths', () => {
    const workbook = new WorkbookStore('criteria-cache-operators')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 2, 3, 4, 5, 6].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.Number,
        value,
      })
    })
    ;['x', 'x', 'y', 'x', 'y', 'x'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `B${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    const matching = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: '>2', stringId: strings.intern('>2') },
        },
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 1, length: 6 },
          criteria: { tag: ValueTag.String, value: 'x', stringId: strings.intern('x') },
        },
      ],
    })
    if (isErrorValue(matching)) {
      throw new Error(`unexpected criteria cache error: ${matching.code}`)
    }
    expect([...matching.rows]).toEqual([3, 5])

    const invalid = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: '>2', stringId: strings.intern('>2') },
        },
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 4, col: 1, length: 5 },
          criteria: { tag: ValueTag.String, value: 'x', stringId: strings.intern('x') },
        },
      ],
    })
    expect(invalid).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers remaining numeric comparator branches and non-numeric rejections', () => {
    const workbook = new WorkbookStore('criteria-cache-comparators')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 2, 3, 4, 5, 6].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.Number,
        value,
      })
    })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A7', {
      tag: ValueTag.String,
      value: 'skip',
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    const gteMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 7, col: 0, length: 8 },
          criteria: { tag: ValueTag.String, value: '>=4', stringId: strings.intern('>=4') },
        },
      ],
    })
    if (isErrorValue(gteMatches)) {
      throw new Error(`unexpected >= criteria error: ${gteMatches.code}`)
    }
    expect([...gteMatches.rows]).toEqual([3, 4, 5])

    const lteMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 7, col: 0, length: 8 },
          criteria: { tag: ValueTag.String, value: '<=2', stringId: strings.intern('<=2') },
        },
      ],
    })
    if (isErrorValue(lteMatches)) {
      throw new Error(`unexpected <= criteria error: ${lteMatches.code}`)
    }
    expect([...lteMatches.rows]).toEqual([0, 1])

    const lessThanMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 7, col: 0, length: 8 },
          criteria: { tag: ValueTag.String, value: '<4', stringId: strings.intern('<4') },
        },
      ],
    })
    if (isErrorValue(lessThanMatches)) {
      throw new Error(`unexpected < criteria error: ${lessThanMatches.code}`)
    }
    expect([...lessThanMatches.rows]).toEqual([0, 1, 2])
  })

  it('matches empty, boolean, wildcard, and error-heavy criteria correctly', () => {
    const workbook = new WorkbookStore('criteria-cache-generic')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    workbook.ensureCell('Sheet1', 'A1')
    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', {
      tag: ValueTag.String,
      value: '',
    })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A3', {
      tag: ValueTag.Boolean,
      value: false,
    })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A4', {
      tag: ValueTag.Number,
      value: 0,
    })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A5', {
      tag: ValueTag.String,
      value: 'northwest',
    })
    const errorCell = workbook.ensureCell('Sheet1', 'A6')
    workbook.cellStore.setValue(errorCell, { tag: ValueTag.Error, code: ErrorCode.Ref })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const criterionCache = createCriterionRangeCacheService({
      runtimeColumnStore,
      regionGraph: createRegionGraph({ workbook }),
      depPatternStore: createDepPatternStore(),
    })

    const emptyMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.Empty },
        },
      ],
    })
    if (isErrorValue(emptyMatches)) {
      throw new Error(`unexpected empty-match error: ${emptyMatches.code}`)
    }
    expect([...emptyMatches.rows]).toEqual([0, 1])

    const falseMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.Boolean, value: false },
        },
      ],
    })
    if (isErrorValue(falseMatches)) {
      throw new Error(`unexpected false-match error: ${falseMatches.code}`)
    }
    expect([...falseMatches.rows]).toEqual([0, 2, 3])

    const wildcardMatches = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.String, value: 'north*', stringId: strings.intern('north*') },
        },
      ],
    })
    if (isErrorValue(wildcardMatches)) {
      throw new Error(`unexpected wildcard-match error: ${wildcardMatches.code}`)
    }
    expect([...wildcardMatches.rows]).toEqual([4])

    const errorCriterion = criterionCache.getOrBuildMatchingRows({
      criteriaPairs: [
        {
          range: { sheetName: 'Sheet1', rowStart: 0, rowEnd: 5, col: 0, length: 6 },
          criteria: { tag: ValueTag.Error, code: ErrorCode.Name },
        },
      ],
    })
    if (isErrorValue(errorCriterion)) {
      throw new Error(`unexpected error-criterion failure: ${errorCriterion.code}`)
    }
    expect([...errorCriterion.rows]).toEqual([])
  })
})
