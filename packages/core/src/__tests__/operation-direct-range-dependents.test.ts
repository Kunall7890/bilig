import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import type { RuntimeDirectAggregateDescriptor, RuntimeDirectCriteriaDescriptor } from '../engine/runtime-state.js'
import { DirectFormulaIndexCollection } from '../engine/services/direct-formula-index-collection.js'
import { aggregateColumnDependencyKey } from '../engine/services/direct-formula-recalc-helpers.js'
import { appendDirectAggregateColumnReverseEdges } from '../engine/services/formula-binding-dependency-helpers.js'
import { createOperationDirectRangeDependentService } from '../engine/services/operation-direct-range-dependents.js'

function directAggregate(overrides: Partial<RuntimeDirectAggregateDescriptor> = {}): RuntimeDirectAggregateDescriptor {
  return {
    regionId: 1,
    aggregateKind: 'sum',
    sheetName: 'Sheet1',
    rowStart: 1,
    rowEnd: 3,
    col: 0,
    colEnd: 1,
    length: 3,
    ...overrides,
  }
}

function directCriteria(overrides: Partial<RuntimeDirectCriteriaDescriptor> = {}): RuntimeDirectCriteriaDescriptor {
  return {
    aggregateKind: 'sum',
    aggregateRange: { regionId: 1, sheetName: 'Sheet1', rowStart: 1, rowEnd: 3, col: 4, length: 3 },
    criteriaPairs: [
      {
        range: { regionId: 2, sheetName: 'Sheet1', rowStart: 1, rowEnd: 3, col: 0, length: 3 },
        criterion: { kind: 'literal', value: { tag: ValueTag.String, value: 'A' } },
      },
    ],
    ...overrides,
  }
}

function createService(
  request: {
    formulas?: {
      readonly get: (cellIndex: number) =>
        | {
            directAggregate?: RuntimeDirectAggregateDescriptor
            directCriteria?: RuntimeDirectCriteriaDescriptor
            dependencyLength?: number
          }
        | undefined
    }
    regionDependents?: readonly number[]
    singleRegionDependent?: number
    aggregateDependents?: readonly number[]
    indexedAggregateDependents?: readonly { readonly cellIndex: number; readonly aggregate: RuntimeDirectAggregateDescriptor }[]
    hasNoCellDependents?: boolean
    entityDependent?: number
    criteriaDelta?: number | undefined
  } = {},
) {
  const aggregateEdges = new Map<number, Set<number>>()
  if (request.aggregateDependents !== undefined) {
    aggregateEdges.set(aggregateColumnDependencyKey(7, 0), new Set(request.aggregateDependents))
  }
  request.indexedAggregateDependents?.forEach((dependent) => {
    appendDirectAggregateColumnReverseEdges(
      aggregateEdges,
      { getSheet: (sheetName: string) => (sheetName === 'Sheet1' ? { id: 7 } : undefined) },
      dependent.aggregate,
      dependent.cellIndex,
    )
  })
  return createOperationDirectRangeDependentService({
    workbook: {
      getSheet: (sheetName: string) => (sheetName === 'Sheet1' ? { id: 7 } : undefined),
    },
    formulas: {
      get: (cellIndex: number) => {
        const formula = request.formulas?.get(cellIndex)
        return formula === undefined
          ? undefined
          : {
              directAggregate: formula.directAggregate,
              directCriteria: formula.directCriteria,
              dependencyIndices: { length: formula.dependencyLength ?? 0 },
            }
      },
    },
    reverseAggregateColumnEdges: aggregateEdges,
    collectRegionFormulaDependentsForCell: () => Uint32Array.from(request.regionDependents ?? []),
    collectSingleRegionFormulaDependentForCell: () => request.singleRegionDependent ?? -1,
    hasNoCellDependents: () => request.hasNoCellDependents ?? true,
    getSingleEntityDependent: () => request.entityDependent ?? -1,
    markFormulaChanged: (cellIndex, count) => count + cellIndex,
    tryDirectCriteriaSumDelta: () => request.criteriaDelta,
    postRecalcLimit: 16_384,
  })
}

describe('operation direct range dependents', () => {
  it('collects affected direct aggregate and criteria dependents without duplicates', () => {
    const service = createService({
      formulas: new Map([
        [10, { directAggregate: directAggregate() }],
        [11, { directCriteria: directCriteria() }],
        [12, { directAggregate: directAggregate({ rowStart: 9, rowEnd: 10 }) }],
      ]),
      regionDependents: [10, 11],
      aggregateDependents: [10, 12],
    })

    expect(service.collectAffectedDirectRangeDependents({ sheetName: 'Sheet1', row: 2, col: 0 })).toEqual([10, 11])
  })

  it('collects multi-window aggregate dependents from the row index without scanning unrelated windows', () => {
    const service = createService({
      formulas: {
        get(cellIndex: number) {
          if (cellIndex === 13) {
            throw new Error('unrelated aggregate window should not be scanned')
          }
          const formulas = new Map([
            [10, { directAggregate: directAggregate({ rowStart: 1, rowEnd: 3 }) }],
            [11, { directAggregate: directAggregate({ rowStart: 2, rowEnd: 4 }) }],
          ])
          return formulas.get(cellIndex)
        },
      },
      indexedAggregateDependents: [
        { cellIndex: 10, aggregate: directAggregate({ rowStart: 1, rowEnd: 3 }) },
        { cellIndex: 11, aggregate: directAggregate({ rowStart: 2, rowEnd: 4 }) },
        { cellIndex: 13, aggregate: directAggregate({ rowStart: 20, rowEnd: 30 }) },
      ],
    })

    expect(service.collectAffectedDirectRangeDependents({ sheetName: 'Sheet1', row: 2, col: 0 })).toEqual([10, 11])
  })

  it('returns single affected dependents and reports ambiguity from aggregate edges', () => {
    expect(
      createService({
        formulas: new Map([[10, { directAggregate: directAggregate() }]]),
        singleRegionDependent: 10,
      }).collectSingleAffectedDirectRangeDependent({ sheetName: 'Sheet1', row: 2, col: 0 }),
    ).toBe(10)

    expect(
      createService({
        formulas: new Map([
          [10, { directAggregate: directAggregate() }],
          [11, { directAggregate: directAggregate() }],
        ]),
        aggregateDependents: [10, 11],
      }).collectSingleAffectedDirectRangeDependent({ sheetName: 'Sheet1', row: 2, col: 0 }),
    ).toBe(-2)

    expect(createService().collectSingleAffectedDirectRangeDependent({ sheetName: 'Missing', row: 2, col: 0 })).toBe(-1)
  })

  it('filters direct aggregate delta candidates by dependency count, range, and downstream dependents', () => {
    expect(
      createService({
        formulas: new Map([[10, { directAggregate: directAggregate() }]]),
      }).canApplyDirectAggregateLiteralDelta(10),
    ).toBe(true)

    expect(
      createService({
        formulas: new Map([[10, { directAggregate: directAggregate() }]]),
      }).canApplyDirectAggregateLiteralDeltaForRequest(10, { sheetName: 'Sheet1', row: 2, col: 0 }),
    ).toBe(true)

    expect(
      createService({
        formulas: new Map([[10, { directAggregate: directAggregate(), dependencyLength: 1 }]]),
      }).canApplyDirectAggregateLiteralDeltaForRequest(10, { sheetName: 'Sheet1', row: 2, col: 0 }),
    ).toBe(false)

    expect(
      createService({
        formulas: new Map([[10, { directAggregate: directAggregate() }]]),
        hasNoCellDependents: false,
      }).canApplyDirectAggregateLiteralDeltaForRequest(10, { sheetName: 'Sheet1', row: 2, col: 0 }),
    ).toBe(false)
  })

  it('returns the single applicable direct aggregate delta dependent', () => {
    expect(
      createService({
        formulas: new Map([[10, { directAggregate: directAggregate() }]]),
        singleRegionDependent: 10,
      }).collectSingleApplicableDirectAggregateDependent({ sheetName: 'Sheet1', row: 2, col: 0 }),
    ).toBe(10)

    expect(
      createService({
        formulas: new Map([[10, { directAggregate: directAggregate(), dependencyLength: 1 }]]),
        singleRegionDependent: 10,
      }).collectSingleApplicableDirectAggregateDependent({ sheetName: 'Sheet1', row: 2, col: 0 }),
    ).toBe(-2)
  })

  it('marks direct range dependents or records post-recalc aggregate deltas', () => {
    expect(
      createService({
        formulas: new Map([[10, { directAggregate: directAggregate() }]]),
        singleRegionDependent: 10,
        entityDependent: 20,
      }).markAffectedDirectRangeDependents(
        {
          sheetName: 'Sheet1',
          row: 2,
          col: 0,
          oldValue: { tag: ValueTag.Number, value: 1 },
          newValue: { tag: ValueTag.Number, value: 4 },
        },
        5,
      ),
    ).toBe(15)

    const postRecalc = new DirectFormulaIndexCollection()
    expect(
      createService({
        formulas: new Map([[10, { directAggregate: directAggregate() }]]),
        singleRegionDependent: 10,
      }).markAffectedDirectRangeDependents(
        {
          sheetName: 'Sheet1',
          row: 2,
          col: 0,
          oldValue: { tag: ValueTag.Number, value: 1 },
          newValue: { tag: ValueTag.Number, value: 4 },
          inputCellIndex: 3,
        },
        5,
        postRecalc,
      ),
    ).toBe(5)
    expect(postRecalc.getDelta(10)).toBe(3)
    expect(postRecalc.hasCoveredDirectRangeInput(3)).toBe(true)
  })
})
