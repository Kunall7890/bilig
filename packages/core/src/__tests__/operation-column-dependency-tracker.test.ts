import { describe, expect, it } from 'vitest'
import { entityPayload, makeExactLookupColumnEntity, makeSortedLookupColumnEntity } from '../entity-ids.js'
import { aggregateColumnDependencyKey } from '../engine/services/direct-formula-recalc-helpers.js'
import {
  createOperationColumnDependencyTrackerService,
  createOperationTrackedColumnDependencyFlagResolver,
} from '../engine/services/operation-column-dependency-tracker.js'

function createTracker(
  request: {
    exact?: boolean
    sorted?: boolean
    aggregate?: boolean
    regionAt?: boolean | undefined
    regionByName?: boolean
    regionAny?: boolean
    cellSlice?: { ptr: number; len: number; cap: number }
  } = {},
) {
  const exactEdges = new Map()
  if (request.exact) {
    exactEdges.set(entityPayload(makeExactLookupColumnEntity(7, 3)), { ptr: 0, len: 1, cap: 1 })
  }
  const sortedEdges = new Map()
  if (request.sorted) {
    sortedEdges.set(entityPayload(makeSortedLookupColumnEntity(7, 3)), { ptr: 0, len: 1, cap: 1 })
  }
  const aggregateEdges = new Map()
  if (request.aggregate) {
    aggregateEdges.set(aggregateColumnDependencyKey(7, 3), new Set([20]))
  }
  return createOperationColumnDependencyTrackerService({
    reverseState: {
      reverseCellEdges: [request.cellSlice],
      reverseExactLookupColumnEdges: exactEdges,
      reverseSortedLookupColumnEdges: sortedEdges,
      reverseAggregateColumnEdges: aggregateEdges,
    },
    workbook: {
      getSheetNameById: (sheetId: number) => (sheetId === 7 ? 'Sheet1' : undefined),
    },
    hasRegionFormulaSubscriptionsForColumnAt: () => request.regionAt,
    hasRegionFormulaSubscriptionsForColumn: () => request.regionByName ?? false,
    hasRegionFormulaSubscriptions: () => request.regionAny ?? false,
  })
}

describe('operation column dependency tracker', () => {
  it('detects lookup, direct range, and aggregate column dependencies independently', () => {
    expect(createTracker({ exact: true }).hasTrackedExactLookupDependents(7, 3)).toBe(true)
    expect(createTracker({ sorted: true }).hasTrackedSortedLookupDependents(7, 3)).toBe(true)
    expect(createTracker({ aggregate: true }).hasTrackedDirectRangeDependents(7, 3)).toBe(true)
    expect(createTracker({ regionAt: true }).hasTrackedDirectRangeDependents(7, 3)).toBe(true)
    expect(createTracker({ regionAt: undefined, regionByName: true }).hasTrackedDirectRangeDependents(7, 3)).toBe(true)
    expect(createTracker().hasTrackedColumnDependents(7, 3)).toBe(false)
  })

  it('uses aggregate reverse edges before probing region subscriptions', () => {
    let regionProbeCount = 0
    const aggregateEdges = new Map([[aggregateColumnDependencyKey(7, 3), new Set([20])]])
    const tracker = createOperationColumnDependencyTrackerService({
      reverseState: {
        reverseCellEdges: [],
        reverseExactLookupColumnEdges: new Map(),
        reverseSortedLookupColumnEdges: new Map(),
        reverseAggregateColumnEdges: aggregateEdges,
      },
      workbook: {
        getSheetNameById: () => {
          throw new Error('aggregate edges should avoid sheet-name lookup')
        },
      },
      hasRegionFormulaSubscriptionsForColumnAt: () => {
        regionProbeCount += 1
        return false
      },
      hasRegionFormulaSubscriptionsForColumn: () => {
        throw new Error('aggregate edges should avoid region subscription lookup')
      },
      hasRegionFormulaSubscriptions: () => false,
    })

    expect(tracker.hasTrackedDirectRangeDependents(7, 3)).toBe(true)
    expect(regionProbeCount).toBe(0)
  })

  it('detects whether any tracked column dependencies exist globally', () => {
    expect(createTracker().hasTrackedColumnDependentsAnywhere()).toBe(false)
    expect(createTracker({ exact: true }).hasTrackedColumnDependentsAnywhere()).toBe(true)
    expect(createTracker({ sorted: true }).hasTrackedColumnDependentsAnywhere()).toBe(true)
    expect(createTracker({ aggregate: true }).hasTrackedColumnDependentsAnywhere()).toBe(true)
    expect(createTracker({ regionAny: true }).hasTrackedColumnDependentsAnywhere()).toBe(true)
  })

  it('treats missing and empty cell edge slices as no dependents', () => {
    expect(createTracker().hasNoCellDependents(0)).toBe(true)
    expect(createTracker({ cellSlice: { ptr: -1, len: 1, cap: 1 } }).hasNoCellDependents(0)).toBe(true)
    expect(createTracker({ cellSlice: { ptr: 0, len: 0, cap: 1 } }).hasNoCellDependents(0)).toBe(true)
    expect(createTracker({ cellSlice: { ptr: 0, len: 1, cap: 1 } }).hasNoCellDependents(0)).toBe(false)
  })

  it('caches tracked column dependency flags and clears them explicitly', () => {
    let exact = true
    let sorted = false
    let aggregate = true
    const resolver = createOperationTrackedColumnDependencyFlagResolver({
      hasTrackedExactLookupDependents: () => exact,
      hasTrackedSortedLookupDependents: () => sorted,
      hasTrackedDirectRangeDependents: () => aggregate,
    })

    expect(resolver.resolve(7, 3)).toEqual({
      hasExactLookupDependents: true,
      hasSortedLookupDependents: false,
      hasAggregateDependents: true,
      needsLookupValueRead: true,
    })
    exact = false
    sorted = true
    aggregate = false
    expect(resolver.resolve(7, 3)).toEqual({
      hasExactLookupDependents: true,
      hasSortedLookupDependents: false,
      hasAggregateDependents: true,
      needsLookupValueRead: true,
    })
    resolver.clear()
    expect(resolver.resolve(7, 3)).toEqual({
      hasExactLookupDependents: false,
      hasSortedLookupDependents: true,
      hasAggregateDependents: false,
      needsLookupValueRead: true,
    })
  })
})
