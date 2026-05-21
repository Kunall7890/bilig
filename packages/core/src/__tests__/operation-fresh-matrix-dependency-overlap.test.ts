import { describe, expect, it } from 'vitest'
import { FormulaMode, ValueTag } from '@bilig/protocol'
import type { CompiledFormula, FormulaNode } from '@bilig/formula'
import { RangeRegistry } from '../range-registry.js'
import type { EngineRuntimeState, RuntimeFormula } from '../engine/runtime-state.js'
import {
  freshMatrixOverlapsFormulaDependencies,
  type FreshMatrixDependencyOverlapArgs,
  type FreshMatrixDependencyOverlapBatch,
} from '../engine/services/operation-fresh-matrix-dependency-overlap.js'

const matrix: FreshMatrixDependencyOverlapBatch = {
  sheet: { name: 'Sheet1' },
  sheetId: 7,
  rowStart: 10,
  rowCount: 4,
  colStart: 1,
  formulaCol: 3,
}

const numberAst: FormulaNode = { kind: 'NumberLiteral', value: 0 }

function compiledFormula(): CompiledFormula {
  return {
    id: 0,
    source: '0',
    mode: FormulaMode.JsOnly,
    depsPtr: 0,
    depsLen: 0,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: 0,
    rangeListOffset: 0,
    rangeListLength: 0,
    maxStackDepth: 0,
    ast: numberAst,
    optimizedAst: numberAst,
    deps: [],
    symbolicNames: [],
    symbolicTables: [],
    symbolicSpills: [],
    volatile: false,
    randCallCount: 0,
    producesSpill: false,
    jsPlan: [],
    program: new Uint32Array(),
    constants: new Float64Array(),
    symbolicRefs: [],
    symbolicRanges: [],
    symbolicStrings: [],
  }
}

function formula(overrides: Partial<RuntimeFormula>): RuntimeFormula {
  const compiled = compiledFormula()
  return {
    cellIndex: 0,
    formulaSlotId: 0,
    planId: 0,
    templateId: undefined,
    source: compiled.source,
    compiled,
    plan: { id: 0, source: compiled.source, compiled },
    dependencyIndices: new Uint32Array(),
    dependencyEntities: { ptr: -1, len: 0, cap: 0 },
    directAggregate: undefined,
    directCriteria: undefined,
    rangeDependencies: new Uint32Array(),
    graphRangeDependencies: new Uint32Array(),
    runtimeProgram: new Uint32Array(),
    constants: new Float64Array(),
    structuralSourceTransform: undefined,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: 0,
    rangeListOffset: 0,
    rangeListLength: 0,
    directLookup: undefined,
    directScalar: undefined,
    ...overrides,
  }
}

function registerRange(
  registry: RangeRegistry,
  range: {
    readonly row1?: number
    readonly row2?: number
    readonly col1?: number
    readonly col2?: number
  } = {},
): number {
  return registry.intern(
    matrix.sheetId,
    {
      kind: 'cells',
      start: { row: range.row1 ?? matrix.rowStart, col: range.col1 ?? matrix.colStart, text: 'A1' },
      end: { row: range.row2 ?? matrix.rowStart + matrix.rowCount - 1, col: range.col2 ?? matrix.formulaCol, text: 'B2' },
    },
    {
      ensureCell: () => 0,
      forEachSheetCell: () => {},
    },
  ).rangeIndex
}

function state(
  args: {
    formulas?: readonly RuntimeFormula[]
    ranges?: RangeRegistry
  } = {},
): Pick<EngineRuntimeState, 'formulas' | 'ranges'> {
  return {
    formulas: new Map(args.formulas?.map((runtimeFormula, index) => [index, runtimeFormula]) ?? []),
    ranges: args.ranges ?? new RangeRegistry(),
  }
}

function overlaps(args: Partial<FreshMatrixDependencyOverlapArgs>, batch: FreshMatrixDependencyOverlapBatch = matrix): boolean {
  return freshMatrixOverlapsFormulaDependencies(
    {
      hasRegionFormulaSubscriptionsIntersectingRect: undefined,
      state: state(),
      ...args,
    },
    batch,
  )
}

describe('freshMatrixOverlapsFormulaDependencies', () => {
  it('uses region subscription fast paths to skip impossible overlaps', () => {
    expect(
      overlaps({
        getRegionFormulaSubscriptionCount: () => 1,
        hasRegionFormulaSubscriptionsOverlappingRange: () => false,
        state: state({ formulas: [formula({})] }),
      }),
    ).toBe(false)

    expect(
      overlaps({
        hasRegionFormulaSubscriptionsIntersectingRect: () => false,
        state: state(),
      }),
    ).toBe(false)
  })

  it('detects direct aggregate and criteria dependencies that overlap the fresh matrix', () => {
    expect(
      overlaps({
        state: state({
          formulas: [
            formula({
              directAggregate: {
                regionId: 1,
                aggregateKind: 'sum',
                sheetName: 'Sheet1',
                rowStart: 12,
                rowEnd: 12,
                col: 2,
                colEnd: 2,
                length: 1,
              },
            }),
          ],
        }),
      }),
    ).toBe(true)

    expect(
      overlaps({
        state: state({
          formulas: [
            formula({
              directCriteria: {
                aggregateKind: 'sum',
                aggregateRange: undefined,
                criteriaPairs: [
                  {
                    range: {
                      regionId: 2,
                      sheetName: 'Sheet1',
                      rowStart: 13,
                      rowEnd: 13,
                      col: 3,
                      length: 1,
                    },
                    criterion: { kind: 'literal', value: { tag: ValueTag.Empty } },
                  },
                ],
              },
            }),
          ],
        }),
      }),
    ).toBe(true)
  })

  it('detects runtime range and graph range descriptors that overlap the fresh matrix', () => {
    const rangeDependencyRegistry = new RangeRegistry()
    const rangeDependencyIndex = registerRange(rangeDependencyRegistry, { col1: 1, col2: 1 })
    expect(
      overlaps({
        state: state({
          formulas: [formula({ rangeDependencies: Uint32Array.of(rangeDependencyIndex) })],
          ranges: rangeDependencyRegistry,
        }),
      }),
    ).toBe(true)

    const nonOverlappingGraphRegistry = new RangeRegistry()
    const nonOverlappingGraphIndex = registerRange(nonOverlappingGraphRegistry, { row1: 9, row2: 9 })
    expect(
      overlaps({
        state: state({
          formulas: [formula({ graphRangeDependencies: Uint32Array.of(nonOverlappingGraphIndex) })],
          ranges: nonOverlappingGraphRegistry,
        }),
      }),
    ).toBe(false)

    const overlappingGraphRegistry = new RangeRegistry()
    const overlappingGraphIndex = registerRange(overlappingGraphRegistry, { row1: 11, row2: 11 })
    expect(
      overlaps({
        state: state({
          formulas: [formula({ graphRangeDependencies: Uint32Array.of(overlappingGraphIndex) })],
          ranges: overlappingGraphRegistry,
        }),
      }),
    ).toBe(true)
  })

  it('stops scanning formula dependencies after the first overlap', () => {
    const ranges = new RangeRegistry()
    const rangeIndex = registerRange(ranges, { row1: 10, row2: 10 })
    expect(
      overlaps({
        state: state({
          formulas: [formula({ rangeDependencies: Uint32Array.of(rangeIndex) }), formula({ rangeDependencies: Uint32Array.of(99) })],
          ranges,
        }),
      }),
    ).toBe(true)
  })
})
