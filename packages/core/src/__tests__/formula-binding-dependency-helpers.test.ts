import { describe, expect, it } from 'vitest'
import { MAX_COLS } from '@bilig/protocol'
import {
  aggregateColumnDependencyKey,
  appendDirectAggregateColumnReverseEdges,
  appendSheetRenameSourceTransform,
  appendTrackedReverseEdge,
  appendUnindexedAggregateColumnReverseEdge,
  collectIndexedDirectAggregateColumnDependentsForRow,
  collectTrackedDependents,
  directCriteriaAggregateColumn,
  directLookupColumnInfo,
  directRegionIdsForFormula,
  formulaColumnCountKey,
  hasQualifiedDependencies,
  parseQualifiedDependencySheetName,
  removeDirectAggregateColumnReverseEdges,
  removeTrackedReverseEdge,
  removeUnindexedAggregateColumnReverseEdge,
  visitIndexedDirectAggregateColumnDependentsForRow,
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

  it('tracks direct aggregate dependencies by every touched sheet column', () => {
    const registry = new Map<number, Set<number>>()
    const workbook = {
      getSheet: (sheetName: string) => (sheetName === 'Sheet1' ? { id: 4 } : undefined),
    }
    const descriptor = {
      aggregateKind: 'sum' as const,
      regionId: 1,
      regionIds: [1, 2, 3],
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 9,
      col: 2,
      colEnd: 4,
      length: 30,
    }

    appendDirectAggregateColumnReverseEdges(registry, workbook, descriptor, 20)
    expect(registry.get(aggregateColumnDependencyKey(4, 2))).toEqual(new Set([20]))
    expect(registry.get(aggregateColumnDependencyKey(4, 3))).toEqual(new Set([20]))
    expect(registry.get(aggregateColumnDependencyKey(4, 4))).toEqual(new Set([20]))

    removeDirectAggregateColumnReverseEdges(registry, workbook, descriptor, 20)
    expect(registry.size).toBe(0)
  })

  it('visits indexed direct aggregate column owners by touched row', () => {
    const registry = new Map<number, Set<number>>()
    const workbook = {
      getSheet: (sheetName: string) => (sheetName === 'Sheet1' ? { id: 4 } : undefined),
    }
    for (let index = 0; index < 8; index += 1) {
      appendDirectAggregateColumnReverseEdges(
        registry,
        workbook,
        {
          aggregateKind: 'sum',
          regionId: index + 1,
          sheetName: 'Sheet1',
          rowStart: index,
          rowEnd: index + 2,
          col: 0,
          colEnd: 0,
          length: 3,
        },
        100 + index,
      )
    }
    const dependents = registry.get(aggregateColumnDependencyKey(4, 0))
    expect(dependents).toBeDefined()
    const visited: number[] = []
    expect(
      visitIndexedDirectAggregateColumnDependentsForRow(dependents!, 3, (formulaCellIndex) => {
        visited.push(formulaCellIndex)
        return true
      }),
    ).toBe(true)
    expect(visited.toSorted((left, right) => left - right)).toEqual([101, 102, 103])
    expect(collectIndexedDirectAggregateColumnDependentsForRow(dependents!, 3)?.toSorted((left, right) => left - right)).toEqual([
      101, 102, 103,
    ])

    removeDirectAggregateColumnReverseEdges(
      registry,
      workbook,
      {
        aggregateKind: 'sum',
        regionId: 3,
        sheetName: 'Sheet1',
        rowStart: 2,
        rowEnd: 4,
        col: 0,
        colEnd: 0,
        length: 3,
      },
      102,
    )
    const afterRemoval: number[] = []
    expect(
      visitIndexedDirectAggregateColumnDependentsForRow(dependents!, 3, (formulaCellIndex) => {
        afterRemoval.push(formulaCellIndex)
        return true
      }),
    ).toBe(true)
    expect(afterRemoval.toSorted((left, right) => left - right)).toEqual([101, 103])
  })

  it('tracks unindexed aggregate column reverse edges alongside indexed direct aggregate owners', () => {
    const registry = new Map<number, Set<number>>()
    const workbook = {
      getSheet: (sheetName: string) => (sheetName === 'Sheet1' ? { id: 4 } : undefined),
    }
    const key = aggregateColumnDependencyKey(4, 0)
    appendDirectAggregateColumnReverseEdges(
      registry,
      workbook,
      {
        aggregateKind: 'sum',
        regionId: 1,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 10,
        col: 0,
        colEnd: 0,
        length: 11,
      },
      10,
    )
    expect(
      visitIndexedDirectAggregateColumnDependentsForRow(registry.get(key)!, 3, () => {
        return true
      }),
    ).toBe(true)
    appendUnindexedAggregateColumnReverseEdge(registry, key, 20)
    expect(registry.get(key)).toEqual(new Set([10, 20]))
    expect(
      visitIndexedDirectAggregateColumnDependentsForRow(registry.get(key)!, 3, () => {
        return true
      }),
    ).toBe(false)
    removeUnindexedAggregateColumnReverseEdge(registry, key, 20)
    expect(registry.get(key)).toEqual(new Set([10]))
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

  it('collects direct aggregate and criteria region ids for structural retargeting', () => {
    expect(
      directRegionIdsForFormula({
        directAggregate: { regionId: 1, regionIds: [1, 2] },
        directCriteria: undefined,
      }),
    ).toEqual([1])

    expect(
      directRegionIdsForFormula({
        directAggregate: { regionId: 1, regionIds: [1, 2] },
        directCriteria: {
          aggregateRange: { regionId: 3, sheetName: 'Sheet1', col: 4 },
          criteriaPairs: [{ range: { regionId: 2 } }, { range: { regionId: 5 } }],
        },
      }),
    ).toEqual([1, 3, 2, 5])
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
