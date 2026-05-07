import { describe, expect, it } from 'vitest'
import { MAX_COLS } from '@bilig/protocol'
import {
  aggregateColumnDependencyKey,
  appendSheetRenameSourceTransform,
  appendTrackedReverseEdge,
  collectTrackedDependents,
  directCriteriaAggregateColumn,
  directLookupColumnInfo,
  directRegionIdsForFormula,
  formulaColumnCountKey,
  hasQualifiedDependencies,
  parseQualifiedDependencySheetName,
  removeTrackedReverseEdge,
} from '../engine/services/formula-binding-dependency-helpers.js'

describe('formula binding dependency helpers', () => {
  it('tracks, removes, and collects reverse-edge dependents without duplicates', () => {
    const registry = new Map<string, Set<number>>()
    appendTrackedReverseEdge(registry, 'Sheet1', 10)
    appendTrackedReverseEdge(registry, 'Sheet1', 10)
    appendTrackedReverseEdge(registry, 'Sheet1', 11)
    appendTrackedReverseEdge(registry, 'Sheet2', 11)

    expect(collectTrackedDependents(registry, ['Sheet1', 'Sheet2']).toSorted((left, right) => left - right)).toEqual([10, 11])
    removeTrackedReverseEdge(registry, 'Sheet1', 10)
    expect(collectTrackedDependents(registry, ['Sheet1'])).toEqual([11])
    removeTrackedReverseEdge(registry, 'Sheet1', 11)
    expect(registry.has('Sheet1')).toBe(false)
  })

  it('parses qualified dependency sheet names including escaped quoted names', () => {
    expect(parseQualifiedDependencySheetName('Sheet1!A1')).toBe('Sheet1')
    expect(parseQualifiedDependencySheetName("'Q1 Report'!A1")).toBe('Q1 Report')
    expect(parseQualifiedDependencySheetName("'Owner''s Sheet'!A1")).toBe("Owner's Sheet")
    expect(parseQualifiedDependencySheetName('A1')).toBeUndefined()
  })

  it('detects qualified dependencies from compiled dependency text', () => {
    expect(hasQualifiedDependencies({ deps: ['A1', 'B2'] })).toBe(false)
    expect(hasQualifiedDependencies({ deps: ['Sheet2!A1'] })).toBe(true)
  })

  it('collects direct aggregate and criteria region ids in stable insertion order', () => {
    expect(
      directRegionIdsForFormula({
        directAggregate: { regionId: 1, regionIds: [1, 2] },
        directCriteria: undefined,
      }),
    ).toEqual([1, 2])

    expect(
      directRegionIdsForFormula({
        directAggregate: { regionId: 1, regionIds: [1, 2] },
        directCriteria: {
          aggregateRange: { regionId: 3, sheetName: 'Sheet1', col: 4 },
          criteriaPairs: [{ range: { regionId: 2 } }, { range: { regionId: 5 } }],
        },
      }),
    ).toEqual([1, 2, 3, 5])
  })

  it('appends sheet-rename source transforms', () => {
    const formula: { sourceRenameTransforms?: Array<{ oldSheetName: string; newSheetName: string }> } = {}
    appendSheetRenameSourceTransform(formula, 'Old', 'New')
    appendSheetRenameSourceTransform(formula, 'Older', 'Newest')

    expect(formula.sourceRenameTransforms).toEqual([
      { oldSheetName: 'Old', newSheetName: 'New' },
      { oldSheetName: 'Older', newSheetName: 'Newest' },
    ])
  })

  it('builds column keys and direct lookup column info', () => {
    expect(aggregateColumnDependencyKey(2, 3)).toBe(2 * MAX_COLS + 3)
    expect(formulaColumnCountKey(4, 5)).toBe(4 * MAX_COLS + 5)
    expect(directCriteriaAggregateColumn({ aggregateRange: { sheetName: 'Sheet1', col: 7 } })).toEqual({
      sheetName: 'Sheet1',
      col: 7,
    })
    expect(directCriteriaAggregateColumn(undefined)).toBeUndefined()

    expect(
      directLookupColumnInfo({
        kind: 'exact',
        prepared: { sheetName: 'Lookup', col: 2 },
      }),
    ).toEqual({ sheetName: 'Lookup', col: 2, isExact: true })
    expect(
      directLookupColumnInfo({
        kind: 'approximate-uniform-numeric',
        sheetName: 'Approx',
        col: 3,
      }),
    ).toEqual({ sheetName: 'Approx', col: 3, isExact: false })
  })
})
