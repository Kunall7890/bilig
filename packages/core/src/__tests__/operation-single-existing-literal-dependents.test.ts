import { describe, expect, it, vi } from 'vitest'
import { makeRangeEntity } from '../entity-ids.js'
import {
  collectSingleFormulaLeafRangeDependentForSingleExistingLiteral,
  hasKnownDynamicFormulaDependentForSingleExistingLiteral,
  type DynamicFormulaDependentArgs,
  type FormulaLeafRangeDependentArgs,
} from '../engine/services/operation-single-existing-literal-dependents.js'

interface FormulaShape {
  readonly dependencyIndices: Uint32Array
  readonly rangeDependencies: Uint32Array
  readonly graphRangeDependencies: Uint32Array
}

function formula(input: {
  readonly dependencyIndices?: readonly number[]
  readonly rangeDependencies?: readonly number[]
  readonly graphRangeDependencies?: readonly number[]
}): FormulaShape {
  return {
    dependencyIndices: Uint32Array.from(input.dependencyIndices ?? []),
    rangeDependencies: Uint32Array.from(input.rangeDependencies ?? []),
    graphRangeDependencies: Uint32Array.from(input.graphRangeDependencies ?? []),
  }
}

function dynamicArgs(input: {
  readonly formulas?: ReadonlyMap<number, FormulaShape>
  readonly collectAt?: (sheetId: number, row: number, col: number) => number | undefined
  readonly hasDynamic?: (cellIndex: number) => boolean
}): DynamicFormulaDependentArgs {
  return {
    state: { formulas: input.formulas ?? new Map<number, FormulaShape>() },
    collectSingleRegionFormulaDependentForCellAt: input.collectAt,
    hasDynamicFormulaDependents: input.hasDynamic ?? (() => false),
  }
}

function leafArgs(input: {
  readonly formulas?: ReadonlyMap<number, FormulaShape>
  readonly singleEntityDependent?: (entityId: number) => number
  readonly collectAt?: (sheetId: number, row: number, col: number) => number | undefined
  readonly collectByName?: (sheetName: string, row: number, col: number) => number
  readonly collectRegion?: (sheetName: string, row: number, col: number) => readonly number[]
}): FormulaLeafRangeDependentArgs {
  return {
    state: { formulas: input.formulas ?? new Map<number, FormulaShape>() },
    getSingleEntityDependent: input.singleEntityDependent ?? (() => -1),
    collectSingleRegionFormulaDependentForCellAt: input.collectAt,
    collectSingleRegionFormulaDependentForCell: input.collectByName ?? (() => -1),
    collectRegionFormulaDependentsForCell: input.collectRegion ?? (() => []),
  }
}

function dynamicRequest(
  singleExistingCellDependent: number,
): Parameters<typeof hasKnownDynamicFormulaDependentForSingleExistingLiteral>[1] {
  return {
    sheetId: 1,
    row: 2,
    col: 3,
    existingIndex: 42,
    singleExistingCellDependent,
  }
}

function leafRequest(
  singleExistingCellDependent: number,
  overrides: Partial<Parameters<typeof collectSingleFormulaLeafRangeDependentForSingleExistingLiteral>[1]> = {},
): Parameters<typeof collectSingleFormulaLeafRangeDependentForSingleExistingLiteral>[1] {
  return {
    existingIndex: 42,
    singleExistingCellDependent,
    sheetId: 1,
    sheetName: 'Sheet1',
    row: 2,
    col: 3,
    formulaScanLimit: 8,
    ...overrides,
  }
}

describe('single existing literal dependent helpers', () => {
  it('detects dynamic dependents from sentinel and compacted range formulas', () => {
    const hasDynamic = vi.fn((cellIndex: number) => cellIndex === 42)

    expect(hasKnownDynamicFormulaDependentForSingleExistingLiteral(dynamicArgs({ hasDynamic }), dynamicRequest(-2))).toBe(true)
    expect(hasDynamic).toHaveBeenCalledWith(42)

    expect(
      hasKnownDynamicFormulaDependentForSingleExistingLiteral(
        dynamicArgs({
          formulas: new Map([[7, formula({ rangeDependencies: [1], graphRangeDependencies: [] })]]),
        }),
        dynamicRequest(7),
      ),
    ).toBe(true)
  })

  it('checks both direct and indexed formula dependents for compacted range edges', () => {
    const formulas = new Map([
      [7, formula({ rangeDependencies: [1], graphRangeDependencies: [1] })],
      [8, formula({ rangeDependencies: [2], graphRangeDependencies: [] })],
    ])

    expect(
      hasKnownDynamicFormulaDependentForSingleExistingLiteral(
        dynamicArgs({
          formulas,
          collectAt: () => 8,
        }),
        dynamicRequest(7),
      ),
    ).toBe(true)

    expect(
      hasKnownDynamicFormulaDependentForSingleExistingLiteral(
        dynamicArgs({
          formulas: new Map(),
          collectAt: () => makeRangeEntity(2),
        }),
        dynamicRequest(makeRangeEntity(1)),
      ),
    ).toBe(false)
  })

  it('collects direct leaf range dependents from range, indexed, and scanned dependency sources', () => {
    const rangeEntity = makeRangeEntity(11)

    expect(
      collectSingleFormulaLeafRangeDependentForSingleExistingLiteral(
        leafArgs({ singleEntityDependent: (entityId) => (entityId === rangeEntity ? 9 : -1) }),
        leafRequest(rangeEntity),
      ),
    ).toBe(9)

    expect(collectSingleFormulaLeafRangeDependentForSingleExistingLiteral(leafArgs({ collectByName: () => 10 }), leafRequest(-1))).toBe(10)

    expect(
      collectSingleFormulaLeafRangeDependentForSingleExistingLiteral(
        leafArgs({
          formulas: new Map([
            [12, formula({ dependencyIndices: [42] })],
            [13, formula({ dependencyIndices: [77] })],
          ]),
        }),
        leafRequest(-1),
      ),
    ).toBe(12)
  })

  it('returns stable ambiguity sentinels for multiple or unavailable leaf dependents', () => {
    expect(
      collectSingleFormulaLeafRangeDependentForSingleExistingLiteral(
        leafArgs({ collectRegion: () => [21, makeRangeEntity(4), 22] }),
        leafRequest(-1),
      ),
    ).toBe(-2)

    expect(
      collectSingleFormulaLeafRangeDependentForSingleExistingLiteral(
        leafArgs({
          formulas: new Map([
            [23, formula({ dependencyIndices: [42] })],
            [24, formula({ dependencyIndices: [42] })],
          ]),
        }),
        leafRequest(-1),
      ),
    ).toBe(-2)

    expect(collectSingleFormulaLeafRangeDependentForSingleExistingLiteral(leafArgs({ collectByName: () => -2 }), leafRequest(-1))).toBe(-2)
  })

  it('honors the formula scan limit before scanning all formulas', () => {
    expect(
      collectSingleFormulaLeafRangeDependentForSingleExistingLiteral(
        leafArgs({
          formulas: new Map([
            [30, formula({ dependencyIndices: [42] })],
            [31, formula({ dependencyIndices: [42] })],
          ]),
        }),
        leafRequest(-1, { formulaScanLimit: 1 }),
      ),
    ).toBe(-1)
  })
})
