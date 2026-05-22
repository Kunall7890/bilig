import { entityPayload, makeExactLookupColumnEntity, makeSortedLookupColumnEntity } from '../../entity-ids.js'
import type { EdgeSlice } from '../../edge-arena.js'
import { aggregateColumnDependencyKey } from './direct-formula-recalc-helpers.js'

export interface OperationColumnDependencyTrackerService {
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedColumnDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedColumnDependentsAnywhere: () => boolean
  readonly hasNoCellDependents: (cellIndex: number) => boolean
}

export interface OperationTrackedColumnDependencyFlags {
  readonly hasExactLookupDependents: boolean
  readonly hasSortedLookupDependents: boolean
  readonly hasAggregateDependents: boolean
  readonly needsLookupValueRead: boolean
}

export interface OperationTrackedColumnDependencyFlagResolver {
  readonly clear: () => void
  readonly resolve: (sheetId: number, col: number) => OperationTrackedColumnDependencyFlags
}

export function createOperationColumnDependencyTrackerService(args: {
  readonly reverseState: {
    readonly reverseCellEdges: Array<EdgeSlice | undefined>
    readonly reverseExactLookupColumnEdges: Map<number, EdgeSlice>
    readonly reverseSortedLookupColumnEdges: Map<number, EdgeSlice>
    readonly reverseAggregateColumnEdges: Map<number, Set<number>>
  }
  readonly workbook: {
    getSheetNameById(sheetId: number): string | undefined
  }
  readonly hasRegionFormulaSubscriptionsForColumn: (sheetName: string, col: number) => boolean
  readonly hasRegionFormulaSubscriptionsForColumnAt?: ((sheetId: number, col: number) => boolean) | undefined
  readonly hasRegionFormulaSubscriptions?: (() => boolean) | undefined
}): OperationColumnDependencyTrackerService {
  const hasTrackedExactLookupDependents = (sheetId: number, col: number): boolean => {
    const exactLookupEdges = args.reverseState.reverseExactLookupColumnEdges
    if (exactLookupEdges.size === 0) {
      return false
    }
    const slice = exactLookupEdges.get(entityPayload(makeExactLookupColumnEntity(sheetId, col)))
    return slice !== undefined && slice.len > 0
  }

  const hasTrackedSortedLookupDependents = (sheetId: number, col: number): boolean => {
    const sortedLookupEdges = args.reverseState.reverseSortedLookupColumnEdges
    if (sortedLookupEdges.size === 0) {
      return false
    }
    const slice = sortedLookupEdges.get(entityPayload(makeSortedLookupColumnEntity(sheetId, col)))
    return slice !== undefined && slice.len > 0
  }

  const hasTrackedDirectRangeDependents = (sheetId: number, col: number): boolean => {
    let hasRegionSubscriptions = args.hasRegionFormulaSubscriptionsForColumnAt?.(sheetId, col)
    if (hasRegionSubscriptions === undefined) {
      const sheetName = args.workbook.getSheetNameById(sheetId)
      hasRegionSubscriptions = sheetName ? args.hasRegionFormulaSubscriptionsForColumn(sheetName, col) : false
    }
    return (
      hasRegionSubscriptions ||
      (args.reverseState.reverseAggregateColumnEdges.get(aggregateColumnDependencyKey(sheetId, col))?.size ?? 0) > 0
    )
  }

  const hasTrackedColumnDependents = (sheetId: number, col: number): boolean =>
    hasTrackedExactLookupDependents(sheetId, col) ||
    hasTrackedSortedLookupDependents(sheetId, col) ||
    hasTrackedDirectRangeDependents(sheetId, col)

  const hasTrackedColumnDependentsAnywhere = (): boolean => {
    return (
      args.reverseState.reverseExactLookupColumnEdges.size > 0 ||
      args.reverseState.reverseSortedLookupColumnEdges.size > 0 ||
      args.reverseState.reverseAggregateColumnEdges.size > 0 ||
      (args.hasRegionFormulaSubscriptions?.() ?? true)
    )
  }

  const hasNoCellDependents = (cellIndex: number): boolean => {
    const slice = args.reverseState.reverseCellEdges[cellIndex]
    return slice === undefined || slice.len === 0 || slice.ptr < 0
  }

  return {
    hasTrackedExactLookupDependents,
    hasTrackedSortedLookupDependents,
    hasTrackedDirectRangeDependents,
    hasTrackedColumnDependents,
    hasTrackedColumnDependentsAnywhere,
    hasNoCellDependents,
  }
}

export function createOperationTrackedColumnDependencyFlagResolver(args: {
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
}): OperationTrackedColumnDependencyFlagResolver {
  const flagsBySheet = new Map<number, Map<number, OperationTrackedColumnDependencyFlags>>()

  return {
    clear() {
      flagsBySheet.clear()
    },
    resolve(sheetId, col) {
      let flagsByColumn = flagsBySheet.get(sheetId)
      if (flagsByColumn === undefined) {
        flagsByColumn = new Map()
        flagsBySheet.set(sheetId, flagsByColumn)
      }
      const cached = flagsByColumn.get(col)
      if (cached !== undefined) {
        return cached
      }
      const hasExactLookupDependents = args.hasTrackedExactLookupDependents(sheetId, col)
      const hasSortedLookupDependents = args.hasTrackedSortedLookupDependents(sheetId, col)
      const hasAggregateDependents = args.hasTrackedDirectRangeDependents(sheetId, col)
      const next = {
        hasExactLookupDependents,
        hasSortedLookupDependents,
        hasAggregateDependents,
        needsLookupValueRead: hasExactLookupDependents || hasSortedLookupDependents || hasAggregateDependents,
      }
      flagsByColumn.set(col, next)
      return next
    },
  }
}
