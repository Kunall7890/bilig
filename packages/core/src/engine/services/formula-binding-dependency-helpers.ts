import type { CompiledFormula } from '@bilig/formula'
import { MAX_COLS } from '@bilig/protocol'
import type { RuntimeDirectAggregateDescriptor, RuntimeDirectCriteriaDescriptor, RuntimeFormula } from '../runtime-state.js'

export type FormulaBindingSourceRenameTransform = {
  readonly oldSheetName: string
  readonly newSheetName: string
}

interface FormulaBindingSourceTransformRecord {
  sourceRenameTransforms?: FormulaBindingSourceRenameTransform[] | undefined
}

type FormulaBindingDirectLookupColumnInfoSource =
  | {
      readonly kind: 'exact' | 'approximate'
      readonly prepared: {
        readonly sheetName: string
        readonly col: number
      }
    }
  | {
      readonly kind: 'exact-uniform-numeric' | 'approximate-uniform-numeric'
      readonly sheetName: string
      readonly col: number
    }

export function appendTrackedReverseEdge<Key extends string | number>(
  registry: Map<Key, Set<number>>,
  key: Key,
  dependentCellIndex: number,
): void {
  const existing = registry.get(key)
  if (existing) {
    existing.add(dependentCellIndex)
    return
  }
  registry.set(key, new Set([dependentCellIndex]))
}

export function removeTrackedReverseEdge<Key extends string | number>(
  registry: Map<Key, Set<number>>,
  key: Key,
  dependentCellIndex: number,
): void {
  const existing = registry.get(key)
  if (!existing) {
    return
  }
  existing.delete(dependentCellIndex)
  if (existing.size === 0) {
    registry.delete(key)
  }
}

export function collectTrackedDependents<Key extends string | number>(registry: Map<Key, Set<number>>, keys: readonly Key[]): number[] {
  if (keys.length === 1) {
    const dependents = registry.get(keys[0]!)
    return dependents === undefined || dependents.size === 0 ? [] : [...dependents]
  }
  const candidates = new Set<number>()
  keys.forEach((key) => {
    registry.get(key)?.forEach((cellIndex) => {
      candidates.add(cellIndex)
    })
  })
  return [...candidates]
}

export function parseQualifiedDependencySheetName(dependency: string): string | undefined {
  const delimiter = dependency.lastIndexOf('!')
  if (delimiter <= 0) {
    return undefined
  }
  const qualifier = dependency.slice(0, delimiter)
  return qualifier.startsWith("'") && qualifier.endsWith("'") ? qualifier.slice(1, -1).replace(/''/g, "'") : qualifier
}

export function hasQualifiedDependencies(compiled: Pick<CompiledFormula, 'deps'>): boolean {
  return compiled.deps.some((dependency) => dependency.includes('!'))
}

export function directRegionIdsForFormula(
  value:
    | Pick<RuntimeFormula, 'directAggregate' | 'directCriteria'>
    | {
        directAggregate: RuntimeDirectAggregateDescriptor | undefined
        directCriteria: RuntimeDirectCriteriaDescriptor | undefined
      },
): readonly number[] {
  if (!value.directAggregate && !value.directCriteria) {
    return []
  }
  const regionIds = new Set<number>()
  if (value.directAggregate) {
    regionIds.add(value.directAggregate.regionId)
  }
  if (value.directCriteria) {
    if (value.directCriteria.aggregateRange) {
      regionIds.add(value.directCriteria.aggregateRange.regionId)
    }
    value.directCriteria.criteriaPairs.forEach((pair) => {
      regionIds.add(pair.range.regionId)
    })
  }
  return [...regionIds]
}

export function appendSheetRenameSourceTransform(
  formula: FormulaBindingSourceTransformRecord,
  oldSheetName: string,
  newSheetName: string,
): void {
  appendSheetRenameSourceTransformRecord(formula, { oldSheetName, newSheetName })
}

export function appendSheetRenameSourceTransformRecord(
  formula: FormulaBindingSourceTransformRecord,
  transform: FormulaBindingSourceRenameTransform,
): void {
  const transforms = formula.sourceRenameTransforms
  if (transforms) {
    transforms.push(transform)
    return
  }
  formula.sourceRenameTransforms = [transform]
}

export function aggregateColumnDependencyKey(sheetId: number, col: number): number {
  return sheetId * MAX_COLS + col
}

type FormulaBindingSheetLookup = {
  getSheet(sheetName: string): { readonly id: number } | undefined
}

interface DirectAggregateColumnOwnerInterval {
  readonly formulaCellIndex: number
  readonly rowStart: number
  readonly rowEnd: number
}

interface DirectAggregateColumnOwnerIndex {
  readonly intervalsByFormulaCellIndex: Map<number, DirectAggregateColumnOwnerInterval>
  readonly intervalsByRowStart: DirectAggregateColumnOwnerInterval[]
  readonly prefixMaxRowEnd: number[]
  readonly unindexedFormulaCellIndices: Set<number>
  orderedByRowStart: boolean
  lastRowStart: number
}

const directAggregateColumnOwnerIndexes = new WeakMap<Set<number>, DirectAggregateColumnOwnerIndex>()

function getOrCreateDirectAggregateColumnOwnerIndex(dependents: Set<number>): DirectAggregateColumnOwnerIndex {
  let index = directAggregateColumnOwnerIndexes.get(dependents)
  if (index === undefined) {
    index = {
      intervalsByFormulaCellIndex: new Map(),
      intervalsByRowStart: [],
      prefixMaxRowEnd: [],
      unindexedFormulaCellIndices: new Set(),
      orderedByRowStart: true,
      lastRowStart: Number.NEGATIVE_INFINITY,
    }
    directAggregateColumnOwnerIndexes.set(dependents, index)
  }
  return index
}

function appendDirectAggregateColumnOwnerInterval(
  dependents: Set<number>,
  formulaCellIndex: number,
  directAggregate: RuntimeDirectAggregateDescriptor,
): void {
  const index = getOrCreateDirectAggregateColumnOwnerIndex(dependents)
  const previous = index.intervalsByFormulaCellIndex.get(formulaCellIndex)
  if (previous !== undefined) {
    if (previous.rowStart === directAggregate.rowStart && previous.rowEnd === directAggregate.rowEnd) {
      return
    }
    index.intervalsByFormulaCellIndex.delete(formulaCellIndex)
  }
  const interval = {
    formulaCellIndex,
    rowStart: directAggregate.rowStart,
    rowEnd: directAggregate.rowEnd,
  }
  index.intervalsByFormulaCellIndex.set(formulaCellIndex, interval)
  index.intervalsByRowStart.push(interval)
  const previousPrefixMax = index.prefixMaxRowEnd[index.prefixMaxRowEnd.length - 1] ?? Number.NEGATIVE_INFINITY
  index.prefixMaxRowEnd.push(Math.max(previousPrefixMax, interval.rowEnd))
  if (interval.rowStart < index.lastRowStart) {
    index.orderedByRowStart = false
  }
  index.lastRowStart = interval.rowStart
}

function removeDirectAggregateColumnOwnerInterval(dependents: Set<number>, formulaCellIndex: number): void {
  directAggregateColumnOwnerIndexes.get(dependents)?.intervalsByFormulaCellIndex.delete(formulaCellIndex)
}

function firstIntervalStartingAfter(intervals: readonly DirectAggregateColumnOwnerInterval[], row: number): number {
  let low = 0
  let high = intervals.length
  while (low < high) {
    const mid = (low + high) >>> 1
    if (intervals[mid]!.rowStart <= row) {
      low = mid + 1
    } else {
      high = mid
    }
  }
  return low
}

export function visitIndexedDirectAggregateColumnDependentsForRow(
  dependents: Set<number>,
  row: number,
  visit: (formulaCellIndex: number) => boolean,
): boolean {
  const index = directAggregateColumnOwnerIndexes.get(dependents)
  if (index === undefined || !index.orderedByRowStart || index.unindexedFormulaCellIndices.size > 0) {
    return false
  }
  const intervals = index.intervalsByRowStart
  let cursor = firstIntervalStartingAfter(intervals, row) - 1
  for (; cursor >= 0; cursor -= 1) {
    if ((index.prefixMaxRowEnd[cursor] ?? Number.NEGATIVE_INFINITY) < row) {
      break
    }
    const interval = intervals[cursor]!
    const current = index.intervalsByFormulaCellIndex.get(interval.formulaCellIndex)
    if (current !== interval || interval.rowEnd < row) {
      continue
    }
    if (!visit(interval.formulaCellIndex)) {
      break
    }
  }
  return true
}

export function collectIndexedDirectAggregateColumnDependentsForRow(dependents: Set<number>, row: number): number[] | undefined {
  const collected: number[] = []
  const usedIndex = visitIndexedDirectAggregateColumnDependentsForRow(dependents, row, (formulaCellIndex) => {
    collected.push(formulaCellIndex)
    return true
  })
  return usedIndex ? collected.toReversed() : undefined
}

function forEachDirectAggregateColumnDependencyKey(
  workbook: FormulaBindingSheetLookup,
  directAggregate: RuntimeDirectAggregateDescriptor | undefined,
  visit: (key: number) => void,
): void {
  if (directAggregate === undefined) {
    return
  }
  const sheet = workbook.getSheet(directAggregate.sheetName)
  if (!sheet) {
    return
  }
  for (let col = directAggregate.col; col <= directAggregate.colEnd; col += 1) {
    visit(aggregateColumnDependencyKey(sheet.id, col))
  }
}

export function appendDirectAggregateColumnReverseEdges(
  registry: Map<number, Set<number>>,
  workbook: FormulaBindingSheetLookup,
  directAggregate: RuntimeDirectAggregateDescriptor | undefined,
  dependentCellIndex: number,
): void {
  forEachDirectAggregateColumnDependencyKey(workbook, directAggregate, (key) => {
    appendTrackedReverseEdge(registry, key, dependentCellIndex)
    const dependents = registry.get(key)
    if (dependents !== undefined && directAggregate !== undefined) {
      appendDirectAggregateColumnOwnerInterval(dependents, dependentCellIndex, directAggregate)
    }
  })
}

export function removeDirectAggregateColumnReverseEdges(
  registry: Map<number, Set<number>>,
  workbook: FormulaBindingSheetLookup,
  directAggregate: RuntimeDirectAggregateDescriptor | undefined,
  dependentCellIndex: number,
): void {
  forEachDirectAggregateColumnDependencyKey(workbook, directAggregate, (key) => {
    const dependents = registry.get(key)
    if (dependents !== undefined) {
      removeDirectAggregateColumnOwnerInterval(dependents, dependentCellIndex)
    }
    removeTrackedReverseEdge(registry, key, dependentCellIndex)
  })
}

export function appendUnindexedAggregateColumnReverseEdge(
  registry: Map<number, Set<number>>,
  key: number,
  dependentCellIndex: number,
): void {
  appendTrackedReverseEdge(registry, key, dependentCellIndex)
  const dependents = registry.get(key)
  if (dependents !== undefined) {
    getOrCreateDirectAggregateColumnOwnerIndex(dependents).unindexedFormulaCellIndices.add(dependentCellIndex)
  }
}

export function removeUnindexedAggregateColumnReverseEdge(
  registry: Map<number, Set<number>>,
  key: number,
  dependentCellIndex: number,
): void {
  const dependents = registry.get(key)
  if (dependents !== undefined) {
    directAggregateColumnOwnerIndexes.get(dependents)?.unindexedFormulaCellIndices.delete(dependentCellIndex)
  }
  removeTrackedReverseEdge(registry, key, dependentCellIndex)
}

export function directCriteriaAggregateColumn(
  value:
    | {
        readonly aggregateRange?:
          | {
              readonly sheetName: string
              readonly col: number
            }
          | undefined
      }
    | undefined,
):
  | {
      readonly sheetName: string
      readonly col: number
    }
  | undefined {
  const aggregateRange = value?.aggregateRange
  return aggregateRange ? { sheetName: aggregateRange.sheetName, col: aggregateRange.col } : undefined
}

export function formulaColumnCountKey(sheetId: number, col: number): number {
  return sheetId * MAX_COLS + col
}

export function directLookupColumnInfo(directLookup: FormulaBindingDirectLookupColumnInfoSource): {
  sheetName: string
  col: number
  isExact: boolean
} {
  switch (directLookup.kind) {
    case 'exact':
      return {
        sheetName: directLookup.prepared.sheetName,
        col: directLookup.prepared.col,
        isExact: true,
      }
    case 'exact-uniform-numeric':
      return {
        sheetName: directLookup.sheetName,
        col: directLookup.col,
        isExact: true,
      }
    case 'approximate':
      return {
        sheetName: directLookup.prepared.sheetName,
        col: directLookup.prepared.col,
        isExact: false,
      }
    case 'approximate-uniform-numeric':
      return {
        sheetName: directLookup.sheetName,
        col: directLookup.col,
        isExact: false,
      }
  }
}
