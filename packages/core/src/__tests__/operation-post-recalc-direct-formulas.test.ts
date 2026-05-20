import { describe, expect, it, vi } from 'vitest'
import { FormulaMode, ValueTag, type ErrorCode } from '@bilig/protocol'
import { createEngineCounters } from '../perf/engine-counters.js'
import type {
  RuntimeDirectCriteriaDescriptor,
  RuntimeDirectLookupDescriptor,
  RuntimeDirectScalarDescriptor,
  U32,
} from '../engine/runtime-state.js'
import { DirectFormulaIndexCollection } from '../engine/services/direct-formula-index-collection.js'
import {
  applyPostRecalcDirectFormulaChanges,
  countOperationPostRecalcDirectFormulaMetric,
  tryApplySinglePostRecalcDirectFormula,
  type ApplyPostRecalcDirectFormulaChangesArgs,
  type DirectFormulaMetricCounts,
  type OperationPostRecalcFormula,
} from '../engine/services/operation-post-recalc-direct-formulas.js'

const directScalar: RuntimeDirectScalarDescriptor = {
  kind: 'abs',
  operand: { kind: 'literal-number', value: 1 },
}

const directLookup: RuntimeDirectLookupDescriptor = {
  kind: 'exact-uniform-numeric',
  operandCellIndex: 1,
  sheetName: 'Sheet1',
  sheetId: 1,
  rowStart: 0,
  rowEnd: 4,
  col: 0,
  length: 5,
  columnVersion: 1,
  structureVersion: 1,
  sheetColumnVersions: new Uint32Array(),
  start: 1,
  step: 1,
  searchMode: 1,
}

function formula(overrides: Partial<OperationPostRecalcFormula> = {}): OperationPostRecalcFormula {
  return {
    compiled: { producesSpill: false },
    directAggregate: undefined,
    directCriteria: undefined,
    directScalar: undefined,
    ...overrides,
  }
}

function makeState(formulas: Map<number, OperationPostRecalcFormula> = new Map()): ApplyPostRecalcDirectFormulaChangesArgs['state'] {
  return {
    workbook: {
      cellStore: {
        flags: [],
        tags: [],
        numbers: [],
        errors: [],
      },
      withBatchedColumnVersionUpdates: (apply: () => void): void => apply(),
    },
    formulas: {
      get: (cellIndex: number) => formulas.get(cellIndex),
    },
    counters: createEngineCounters(),
  }
}

function makeArgs(overrides: Partial<ApplyPostRecalcDirectFormulaChangesArgs> = {}): ApplyPostRecalcDirectFormulaChangesArgs {
  const collection = new DirectFormulaIndexCollection()
  const metrics: DirectFormulaMetricCounts = { wasmFormulaCount: 0, jsFormulaCount: 0 }
  return {
    state: makeState(),
    collection,
    recalculated: new Uint32Array(),
    didRunRecalc: false,
    metrics,
    applyDirectFormulaCurrentResult: vi.fn(() => true),
    applyDirectFormulaNumericDelta: vi.fn(() => true),
    applyDirectScalarCurrentValue: vi.fn(() => true),
    tryApplyDirectScalarDeltas: vi.fn(() => undefined),
    tryApplyDirectFormulaDeltas: vi.fn(() => undefined),
    countPostRecalcDirectFormulaMetric: vi.fn(),
    evaluateDirectFormula: vi.fn(() => undefined),
    ...overrides,
  }
}

describe('operation post-recalc direct formula helpers', () => {
  it('applies a single current result while allowing callers to suppress captured changes', () => {
    const collection = new DirectFormulaIndexCollection()
    collection.addCurrentResult(7, { kind: 'number', value: 42 })
    const applyCurrent = vi.fn(() => true)
    const args = makeArgs({
      state: makeState(new Map([[7, formula({ directScalar })]])),
      collection,
      applyDirectFormulaCurrentResult: applyCurrent,
    })

    const changed = tryApplySinglePostRecalcDirectFormula(args, false)

    expect(Array.from(changed ?? [])).toEqual([])
    expect(applyCurrent).toHaveBeenCalledWith(7, { kind: 'number', value: 42 })
  })

  it('rejects a single current result when the formula target is stale', () => {
    const collection = new DirectFormulaIndexCollection()
    collection.addCurrentResult(7, { kind: 'number', value: 42 })
    const applyCurrent = vi.fn(() => true)
    const args = makeArgs({
      collection,
      applyDirectFormulaCurrentResult: applyCurrent,
    })

    const changed = tryApplySinglePostRecalcDirectFormula(args, false)

    expect(changed).toBeUndefined()
    expect(applyCurrent).not.toHaveBeenCalled()
  })

  it('uses complete delta application before falling back to per-formula evaluation', () => {
    const collection = new DirectFormulaIndexCollection()
    collection.addDelta(2, 1)
    collection.addDelta(3, 1)
    const tryApplyDirectFormulaDeltas = vi.fn((): U32 => Uint32Array.of(2, 3))
    const evaluateDirectFormula = vi.fn(() => [99])

    const changed = applyPostRecalcDirectFormulaChanges(
      makeArgs({
        collection,
        recalculated: Uint32Array.of(1),
        tryApplyDirectFormulaDeltas,
        evaluateDirectFormula,
      }),
    )

    expect(Array.from(changed)).toEqual([1, 2, 3])
    expect(tryApplyDirectFormulaDeltas).toHaveBeenCalledWith(collection, true)
    expect(evaluateDirectFormula).not.toHaveBeenCalled()
  })

  it('counts post-recalc direct formula metrics by formula mode', () => {
    const counts: DirectFormulaMetricCounts = { wasmFormulaCount: 0, jsFormulaCount: 0 }
    const formulas = new Map<number, OperationPostRecalcFormula>([
      [
        1,
        formula({
          compiled: { mode: FormulaMode.WasmFastPath, producesSpill: false },
          directScalar,
        }),
      ],
      [
        2,
        formula({
          compiled: { mode: FormulaMode.JsOnly, producesSpill: false },
          directAggregate: {},
        }),
      ],
      [
        3,
        formula({
          compiled: { mode: FormulaMode.JsOnly, producesSpill: false },
          directLookup,
        }),
      ],
    ])

    ;[1, 2, 3, 4].forEach((cellIndex) => {
      countOperationPostRecalcDirectFormulaMetric({
        formulas: { get: (lookupCellIndex) => formulas.get(lookupCellIndex) },
        cellIndex,
        counts,
      })
    })

    expect(counts).toEqual({ wasmFormulaCount: 1, jsFormulaCount: 2 })
  })

  it('evaluates remaining direct formulas in a batched column-version update', () => {
    const collection = new DirectFormulaIndexCollection()
    collection.add(2)
    collection.add(3)
    const withBatchedColumnVersionUpdates = vi.fn((apply: () => void): void => apply())
    const formulas = new Map([
      [2, formula({ directScalar })],
      [3, formula({ directScalar })],
    ])
    const evaluateDirectFormula = vi.fn((cellIndex: number) => [cellIndex + 100])

    const changed = applyPostRecalcDirectFormulaChanges(
      makeArgs({
        state: {
          ...makeState(formulas),
          workbook: {
            cellStore: {
              flags: [],
              tags: [],
              numbers: [],
              errors: [],
            },
            withBatchedColumnVersionUpdates,
          },
        },
        collection,
        recalculated: Uint32Array.of(1),
        didRunRecalc: true,
        applyDirectScalarCurrentValue: vi.fn(() => false),
        evaluateDirectFormula,
      }),
    )

    expect(Array.from(changed)).toEqual([1, 2, 3, 102, 103])
    expect(withBatchedColumnVersionUpdates).toHaveBeenCalledOnce()
    expect(evaluateDirectFormula).toHaveBeenCalledWith(2)
    expect(evaluateDirectFormula).toHaveBeenCalledWith(3)
  })

  it('reuses evaluated direct criteria results for copied formula shapes in one post-recalc pass', () => {
    const collection = new DirectFormulaIndexCollection()
    collection.add(2)
    collection.add(3)
    const directCriteria: RuntimeDirectCriteriaDescriptor = {
      aggregateKind: 'sum',
      aggregateRange: {
        regionId: 2,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        col: 1,
        length: 32,
      },
      criteriaPairs: [
        {
          range: {
            regionId: 1,
            sheetName: 'Sheet1',
            rowStart: 0,
            rowEnd: 31,
            col: 0,
            length: 32,
          },
          criterion: { kind: 'cell', cellIndex: 1 },
        },
      ],
    }
    const formulas = new Map([
      [2, formula({ directCriteria })],
      [3, formula({ directCriteria: { ...directCriteria, criteriaPairs: [...directCriteria.criteriaPairs] } })],
    ])
    const tags: Array<ValueTag | undefined> = []
    const numbers: number[] = []
    const errors: ErrorCode[] = []
    const evaluateDirectFormula = vi.fn((cellIndex: number) => {
      tags[cellIndex] = ValueTag.Number
      numbers[cellIndex] = 42
      return undefined
    })
    const applyDirectFormulaCurrentResult = vi.fn(() => true)

    const changed = applyPostRecalcDirectFormulaChanges(
      makeArgs({
        state: {
          ...makeState(formulas),
          workbook: {
            cellStore: {
              flags: [],
              tags,
              numbers,
              errors,
            },
            withBatchedColumnVersionUpdates: (apply: () => void): void => apply(),
          },
        },
        collection,
        applyDirectFormulaCurrentResult,
        evaluateDirectFormula,
      }),
    )

    expect(Array.from(changed)).toEqual([2, 3])
    expect(evaluateDirectFormula).toHaveBeenCalledOnce()
    expect(evaluateDirectFormula).toHaveBeenCalledWith(2)
    expect(applyDirectFormulaCurrentResult).toHaveBeenCalledWith(3, { kind: 'number', value: 42 })
  })

  it('counts aggregate and scalar delta applications on the single-cell delta path', () => {
    const collection = new DirectFormulaIndexCollection()
    collection.addDelta(4, 2)
    const state = makeState(
      new Map([
        [
          4,
          formula({
            directAggregate: {},
            directScalar,
          }),
        ],
      ]),
    )

    const changed = tryApplySinglePostRecalcDirectFormula(
      makeArgs({
        state,
        collection,
      }),
    )

    expect(Array.from(changed ?? [])).toEqual([4])
    expect(state.counters.directAggregateDeltaApplications).toBe(1)
    expect(state.counters.directScalarDeltaApplications).toBe(1)
  })
})
