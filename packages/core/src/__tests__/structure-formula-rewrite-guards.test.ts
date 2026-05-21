import { compileFormula, type CompiledFormula, type StructuralAxisTransform } from '@bilig/formula'
import { describe, expect, it } from 'vitest'

import type { RuntimeDirectAggregateDescriptor, RuntimeFormula } from '../engine/runtime-state.js'
import {
  dependencyTouchesSheet,
  isStructurallyStableSimpleFormulaNode,
  rangeDependencyAxisAffected,
  runtimeDirectRangeAxisAffected,
  structuralDirectAggregateRewritePreservesValue,
  structuralRewritePreservesBinding,
  structuralRewritePreservesValue,
} from '../engine/services/structure-formula-rewrite-guards.js'

const rowInsert = { kind: 'insert', axis: 'row', start: 1, count: 1 } satisfies StructuralAxisTransform
const rowDelete = { kind: 'delete', axis: 'row', start: 1, count: 1 } satisfies StructuralAxisTransform
const rowMoveExtract = { kind: 'move', axis: 'row', start: 1, count: 1, target: 4 } satisfies StructuralAxisTransform
const rowMoveWholeRange = { kind: 'move', axis: 'row', start: 0, count: 5, target: 6 } satisfies StructuralAxisTransform
const colMoveExtract = { kind: 'move', axis: 'column', start: 0, count: 1, target: 2 } satisfies StructuralAxisTransform

function directAggregate(overrides: Partial<RuntimeDirectAggregateDescriptor> = {}): RuntimeDirectAggregateDescriptor {
  return {
    regionId: 1,
    aggregateKind: 'sum',
    sheetName: 'Sheet1',
    rowStart: 0,
    rowEnd: 4,
    col: 0,
    colEnd: 0,
    length: 5,
    ...overrides,
  }
}

function runtimeFormula(
  source: string,
  overrides: Partial<
    Pick<RuntimeFormula, 'directAggregate' | 'directScalar' | 'directLookup' | 'directCriteria' | 'rangeDependencies'>
  > = {},
): RuntimeFormula {
  const compiled = compileFormula(source)
  return {
    cellIndex: 1,
    formulaSlotId: 1,
    planId: 1,
    templateId: undefined,
    source,
    compiled,
    plan: { programOffset: 0, programLength: 0, constNumberOffset: 0, constNumberLength: 0, rangeListOffset: 0, rangeListLength: 0 },
    dependencyIndices: new Uint32Array(0),
    dependencyEntities: new Uint32Array(0),
    rangeDependencies: new Uint32Array(compiled.symbolicRanges.length),
    graphRangeDependencies: new Uint32Array(compiled.symbolicRanges.length),
    runtimeProgram: new Uint32Array(0),
    constants: new Float64Array(0),
    structuralSourceTransform: undefined,
    programOffset: 0,
    programLength: 0,
    constNumberOffset: 0,
    constNumberLength: 0,
    rangeListOffset: 0,
    rangeListLength: 0,
    directLookup: undefined,
    directAggregate: undefined,
    directScalar: undefined,
    directCriteria: undefined,
    ...overrides,
  }
}

function rewritten(compiled: CompiledFormula, overrides: Partial<CompiledFormula> = {}) {
  return {
    compiled: {
      ...compiled,
      ...overrides,
    },
    reusedProgram: true,
  }
}

describe('structural formula rewrite guards', () => {
  it('detects sheet-qualified dependencies and unaffected ranges precisely', () => {
    expect(dependencyTouchesSheet('Sheet1!A1', 'Sheet1')).toBe(true)
    expect(dependencyTouchesSheet("'Revenue FY'!A1", 'Revenue FY')).toBe(true)
    expect(dependencyTouchesSheet('A1', 'Sheet1')).toBe(false)

    expect(rangeDependencyAxisAffected({ sheetId: 1, row1: 0, row2: 4, col1: 0, col2: 1 }, 2, rowMoveExtract)).toBe(false)
    expect(rangeDependencyAxisAffected({ sheetId: 1, row1: 0, row2: 4, col1: 0, col2: 1 }, 1, rowMoveExtract)).toBe(true)
    expect(rangeDependencyAxisAffected({ sheetId: 1, row1: 0, row2: 4, col1: 0, col2: 1 }, 1, rowMoveWholeRange)).toBe(true)
    expect(rangeDependencyAxisAffected({ sheetId: 1, row1: 0, row2: 4, col1: 0, col2: 1 }, 1, rowDelete)).toBe(true)

    expect(runtimeDirectRangeAxisAffected(undefined, 'Sheet1', colMoveExtract, undefined)).toBe(false)
    expect(
      runtimeDirectRangeAxisAffected(1, 'Sheet1', colMoveExtract, {
        sheetName: 'Other',
        rowStart: 0,
        rowEnd: 4,
        col: 0,
        colEnd: 1,
      }),
    ).toBe(false)
    expect(
      runtimeDirectRangeAxisAffected(1, 'Sheet1', colMoveExtract, {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 4,
        col: 0,
        colEnd: 1,
      }),
    ).toBe(true)
  })

  it('only allows value-preserving rewrites for stable formulas and accepted structural shapes', () => {
    expect(
      isStructurallyStableSimpleFormulaNode({
        kind: 'ArrayConstant',
        rows: [[{ kind: 'NumberLiteral', value: 1 }], [{ kind: 'BooleanLiteral', value: true }]],
      }),
    ).toBe(true)
    expect(isStructurallyStableSimpleFormulaNode({ kind: 'RangeRef', refKind: 'cells', start: 'A1', end: 'A2' })).toBe(false)

    const formula = runtimeFormula('A1+1')
    expect(structuralRewritePreservesValue(formula, rewritten(formula.compiled), rowInsert)).toBe(true)
    expect(structuralRewritePreservesValue(formula, rewritten(formula.compiled), rowDelete)).toBe(false)
    expect(
      structuralRewritePreservesValue(
        runtimeFormula('A1+1', { directScalar: { kind: 'literal-number', value: 4 } }),
        rewritten(formula.compiled),
        rowMoveExtract,
      ),
    ).toBe(false)
  })

  it('keeps binding reuse strict about range layout and direct descriptor invariants', () => {
    const rangeFormula = runtimeFormula('SUM(A1:A5)')
    expect(structuralRewritePreservesBinding(rangeFormula, rewritten(rangeFormula.compiled), true)).toBe(true)
    expect(structuralRewritePreservesBinding(rangeFormula, rewritten(rangeFormula.compiled), false)).toBe(false)

    expect(
      structuralRewritePreservesBinding(
        rangeFormula,
        rewritten(rangeFormula.compiled, {
          parsedSymbolicRanges: [],
        }),
        true,
      ),
    ).toBe(false)
    expect(
      structuralRewritePreservesBinding(
        rangeFormula,
        rewritten(rangeFormula.compiled, {
          parsedSymbolicRanges: rangeFormula.compiled.parsedSymbolicRanges?.map((range) =>
            Object.assign({}, range, { refKind: range.refKind === 'cells' ? 'cols' : 'cells' }),
          ),
        }),
        true,
      ),
    ).toBe(false)

    const aggregateFormula = runtimeFormula('SUM(A1:A5)', { directAggregate: directAggregate() })
    expect(structuralDirectAggregateRewritePreservesValue(aggregateFormula, rewritten(aggregateFormula.compiled), rowInsert)).toBe(true)
    expect(structuralDirectAggregateRewritePreservesValue(aggregateFormula, rewritten(aggregateFormula.compiled), rowMoveExtract)).toBe(
      false,
    )
    expect(
      structuralDirectAggregateRewritePreservesValue(
        runtimeFormula('SUM(A1:A5)', { directAggregate: directAggregate({ rowStart: 0, rowEnd: 4 }) }),
        rewritten(aggregateFormula.compiled),
        rowMoveWholeRange,
      ),
    ).toBe(true)
  })
})
