import type { CompiledFormula } from '@bilig/formula'
import { MAX_COLS } from '@bilig/protocol'
import type { RuntimeDirectAggregateDescriptor, RuntimeDirectCriteriaDescriptor, RuntimeFormula } from '../runtime-state.js'

type FormulaBindingSourceRenameTransform = {
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
): number[] {
  if (value.directAggregate && !value.directCriteria) {
    return value.directAggregate.regionIds ? [...value.directAggregate.regionIds] : [value.directAggregate.regionId]
  }
  if (!value.directAggregate && !value.directCriteria) {
    return []
  }
  const regionIds = new Set<number>()
  if (value.directAggregate) {
    ;(value.directAggregate.regionIds ?? [value.directAggregate.regionId]).forEach((regionId) => {
      regionIds.add(regionId)
    })
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
  const transforms = formula.sourceRenameTransforms
  if (transforms) {
    transforms.push({ oldSheetName, newSheetName })
    return
  }
  formula.sourceRenameTransforms = [{ oldSheetName, newSheetName }]
}

export function aggregateColumnDependencyKey(sheetId: number, col: number): number {
  return sheetId * MAX_COLS + col
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
