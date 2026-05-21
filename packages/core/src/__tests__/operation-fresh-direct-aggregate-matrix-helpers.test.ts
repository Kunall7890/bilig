import { describe, expect, it } from 'vitest'
import { ErrorCode, FormulaMode } from '@bilig/protocol'
import type { CompiledFormula, FormulaNode, ParsedRangeReferenceInfo } from '@bilig/formula'
import {
  createFreshMatrixDirectAggregateTemplate,
  evaluateFreshDirectAggregateMatrixRow,
  materializeFreshMatrixAxisIds,
  normalizeFreshMatrixDirectAggregateOffset,
  tryTranslateFreshMatrixDirectAggregateTemplate,
} from '../engine/services/operation-fresh-direct-aggregate-matrix-helpers.js'

const numberAst: FormulaNode = { kind: 'NumberLiteral', value: 0 }

function compiled(): CompiledFormula {
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
    parsedSymbolicRanges: [],
    symbolicRanges: [],
    symbolicStrings: [],
  }
}

function parsedRange(overrides: Partial<ParsedRangeReferenceInfo> = {}): ParsedRangeReferenceInfo {
  return {
    address: 'A2:C2',
    kind: 'range',
    refKind: 'cells',
    startAddress: 'A2',
    endAddress: 'C2',
    startRow: 1,
    endRow: 1,
    startCol: 0,
    endCol: 2,
    ...overrides,
  }
}

describe('fresh direct aggregate matrix helpers', () => {
  it('normalizes result offsets and materializes stable axis id runs', () => {
    expect(normalizeFreshMatrixDirectAggregateOffset(undefined)).toBeUndefined()
    expect(normalizeFreshMatrixDirectAggregateOffset(0)).toBeUndefined()
    expect(normalizeFreshMatrixDirectAggregateOffset(-0)).toBeUndefined()
    expect(normalizeFreshMatrixDirectAggregateOffset(3)).toBe(3)
    expect(materializeFreshMatrixAxisIds(3, 4, (index) => `axis-${index}`)).toEqual(['axis-4', 'axis-5', 'axis-6'])
  })

  it('creates templates only for same-row cell ranges before the formula column', () => {
    const aggregate = { callee: 'SUM', aggregateKind: 'sum' as const, symbolicRangeIndex: 0, resultOffset: 0 }
    const base = {
      aggregate,
      compiled: compiled(),
      formulaCol: 4,
      range: parsedRange(),
      row: 1,
      templateId: 9,
    }

    expect(createFreshMatrixDirectAggregateTemplate(base)).toMatchObject({
      aggregateKind: 'sum',
      formulaCol: 4,
      rangeColStart: 0,
      rangeColEnd: 2,
      resultOffset: undefined,
      row: 1,
      templateId: 9,
    })
    expect(createFreshMatrixDirectAggregateTemplate({ ...base, range: parsedRange({ refKind: 'rows' }) })).toBeUndefined()
    expect(createFreshMatrixDirectAggregateTemplate({ ...base, range: parsedRange({ sheetName: 'Other' }) })).toBeUndefined()
    expect(createFreshMatrixDirectAggregateTemplate({ ...base, range: parsedRange({ endRow: 2 }) })).toBeUndefined()
    expect(createFreshMatrixDirectAggregateTemplate({ ...base, range: parsedRange({ startCol: 3, endCol: 2 }) })).toBeUndefined()
    expect(createFreshMatrixDirectAggregateTemplate({ ...base, range: parsedRange({ endCol: 4 }) })).toBeUndefined()
  })

  it('validates translated sources against the template shape before compiling', () => {
    const template = createFreshMatrixDirectAggregateTemplate({
      aggregate: { callee: 'SUM', aggregateKind: 'sum', symbolicRangeIndex: 0, resultOffset: 2 },
      compiled: compiled(),
      formulaCol: 4,
      range: parsedRange(),
      row: 1,
      templateId: 1,
    })!

    expect(tryTranslateFreshMatrixDirectAggregateTemplate(template, '=SUM(A2:C2)+2', 1, 3)).toBeUndefined()
    expect(tryTranslateFreshMatrixDirectAggregateTemplate(template, '=COUNT(A2:C2)+2', 1, 4)).toBeUndefined()
    expect(tryTranslateFreshMatrixDirectAggregateTemplate(template, '=SUM(A3:C3)+2', 1, 4)).toBeUndefined()
    expect(tryTranslateFreshMatrixDirectAggregateTemplate(template, '=SUM(A2:C2)+3', 1, 4)).toBeUndefined()
  })

  it('evaluates direct aggregate matrix rows with spreadsheet aggregate semantics', () => {
    const values = Float64Array.of(1, 2, 3, 4, 5, 6)
    const base = {
      colStart: 0,
      colEnd: 2,
      inputColCount: 3,
      matrixColStart: 0,
      resultOffset: undefined,
      rowOffset: 1,
      values,
    }

    expect(evaluateFreshDirectAggregateMatrixRow({ ...base, aggregateKind: 'sum' })).toEqual({ kind: 'number', value: 15 })
    expect(evaluateFreshDirectAggregateMatrixRow({ ...base, aggregateKind: 'count' })).toEqual({ kind: 'number', value: 3 })
    expect(evaluateFreshDirectAggregateMatrixRow({ ...base, aggregateKind: 'average', resultOffset: 1 })).toEqual({
      kind: 'number',
      value: 6,
    })
    expect(evaluateFreshDirectAggregateMatrixRow({ ...base, aggregateKind: 'min' })).toEqual({ kind: 'number', value: 4 })
    expect(evaluateFreshDirectAggregateMatrixRow({ ...base, aggregateKind: 'max' })).toEqual({ kind: 'number', value: 6 })
    expect(evaluateFreshDirectAggregateMatrixRow({ ...base, aggregateKind: 'average', colStart: 2, colEnd: 1 })).toEqual({
      kind: 'error',
      code: ErrorCode.Div0,
    })
    expect(evaluateFreshDirectAggregateMatrixRow({ ...base, aggregateKind: 'min', colStart: 2, colEnd: 1 })).toEqual({
      kind: 'number',
      value: 0,
    })
    expect(evaluateFreshDirectAggregateMatrixRow({ ...base, aggregateKind: 'max', colStart: 2, colEnd: 1 })).toEqual({
      kind: 'number',
      value: 0,
    })
  })
})
