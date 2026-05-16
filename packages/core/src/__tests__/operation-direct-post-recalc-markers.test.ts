import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { makeCellEntity } from '../entity-ids.js'
import { DirectFormulaIndexCollection } from '../engine/services/direct-formula-index-collection.js'
import { createOperationDirectPostRecalcMarkers } from '../engine/services/operation-direct-post-recalc-markers.js'
import type { RuntimeDirectScalarDescriptor } from '../engine/runtime-state.js'

type MarkerArgs = Parameters<typeof createOperationDirectPostRecalcMarkers>[0]

const numberValue = (value: number): CellValue => ({ tag: ValueTag.Number, value })

function formula(directScalar: RuntimeDirectScalarDescriptor) {
  return {
    compiled: {
      deps: [],
      volatile: false,
      producesSpill: false,
    },
    directAggregate: undefined,
    directCriteria: undefined,
    directLookup: undefined,
    directScalar,
  }
}

function binaryScalar(
  operator: '+' | '-' | '*' | '/',
  leftCellIndex: number,
  right: { kind: 'cell'; cellIndex: number } | { kind: 'literal-number'; value: number },
): RuntimeDirectScalarDescriptor {
  return {
    kind: 'binary',
    operator,
    left: { kind: 'cell', cellIndex: leftCellIndex },
    right,
  }
}

function createMarkers(input: {
  readonly formulas: Map<number, ReturnType<typeof formula>>
  readonly numbers: Map<number, number>
  readonly singleDependents: Map<number, number>
  readonly entityDependents?: Map<number, Uint32Array>
  readonly canSkipDirectFormulaColumnVersion?: (cellIndex: number) => boolean
}) {
  const tags: ValueTag[] = []
  input.numbers.forEach((_value, cellIndex) => {
    tags[cellIndex] = ValueTag.Number
  })
  const state: MarkerArgs['state'] = {
    workbook: {
      cellStore: {
        flags: [],
        tags,
        getValue: (cellIndex: number) => numberValue(input.numbers.get(cellIndex) ?? 0),
      },
    },
    formulas: {
      get: (cellIndex: number) => input.formulas.get(cellIndex),
    },
    strings: {
      get: (id: number) => `s${id}`,
    },
  }

  return createOperationDirectPostRecalcMarkers({
    state,
    getSingleEntityDependent: (entityId) => input.singleDependents.get(entityId) ?? -1,
    getEntityDependents: (entityId) => input.entityDependents?.get(entityId) ?? new Uint32Array(),
    hasNoCellDependents: () => true,
    canSkipDirectFormulaColumnVersion: input.canSkipDirectFormulaColumnVersion ?? (() => true),
    readDirectScalarCellNumber: (cellIndex) => input.numbers.get(cellIndex) ?? 0,
    directScalarCellNumericValue: (cellIndex) => input.numbers.get(cellIndex),
    directScalarCurrentResultMatchesCell: () => false,
    lookupCurrent: {
      canEvaluateDirectUniformLookupCurrentResultFromNumeric: () => false,
      tryDirectExactLookupCurrentResult: () => undefined,
      tryDirectUniformLookupCurrentResult: () => undefined,
      tryDirectUniformLookupCurrentResultFromNumeric: () => undefined,
    },
    scalarDeltaClosureLimit: 32,
  })
}

describe('operation direct post-recalc markers', () => {
  it('marks a linear scalar closure as validated constant deltas', () => {
    const formulas = new Map([
      [20, formula(binaryScalar('+', 10, { kind: 'literal-number', value: 1 }))],
      [30, formula(binaryScalar('+', 20, { kind: 'literal-number', value: 2 }))],
    ])
    const markers = createMarkers({
      formulas,
      numbers: new Map([
        [20, 3],
        [30, 5],
      ]),
      singleDependents: new Map([
        [makeCellEntity(10), 20],
        [makeCellEntity(20), 30],
        [makeCellEntity(30), -1],
      ]),
    })
    const collection = new DirectFormulaIndexCollection()

    expect(markers.tryMarkDirectScalarLinearDeltaClosure(10, numberValue(2), numberValue(5), collection)).toBe(true)

    expect(collection.size).toBe(2)
    expect(collection.getCellIndicesForRead()).toBeInstanceOf(Uint32Array)
    expect(collection.getDelta(20)).toBe(3)
    expect(collection.getDelta(30)).toBe(3)
    expect(collection.hasCompleteScalarDeltas()).toBe(true)
    expect(collection.hasValidatedScalarDeltaCells()).toBe(true)
  })

  it('falls back to graph scalar closure when the dependent path branches', () => {
    const formulas = new Map([
      [20, formula(binaryScalar('+', 10, { kind: 'literal-number', value: 1 }))],
      [30, formula(binaryScalar('+', 20, { kind: 'literal-number', value: 2 }))],
      [40, formula(binaryScalar('*', 20, { kind: 'literal-number', value: 2 }))],
    ])
    const markers = createMarkers({
      formulas,
      numbers: new Map([
        [20, 3],
        [30, 5],
        [40, 6],
      ]),
      singleDependents: new Map([
        [makeCellEntity(10), 20],
        [makeCellEntity(20), -2],
      ]),
      entityDependents: new Map([
        [makeCellEntity(10), new Uint32Array([20])],
        [makeCellEntity(20), new Uint32Array([30, 40])],
      ]),
    })
    const collection = new DirectFormulaIndexCollection()

    markers.markDirectScalarDeltaClosure(10, numberValue(2), numberValue(5), collection)

    expect(collection.size).toBe(3)
    expect(collection.getDelta(20)).toBe(3)
    expect(collection.getDelta(30)).toBe(3)
    expect(collection.getDelta(40)).toBe(6)
  })
})
