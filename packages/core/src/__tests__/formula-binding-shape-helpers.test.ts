import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { ParsedRangeReferenceInfo } from '@bilig/formula'
import type {
  RuntimeDirectAggregateDescriptor,
  RuntimeDirectCriteriaDescriptor,
  RuntimeDirectScalarDescriptor,
} from '../engine/runtime-state.js'
import {
  directAggregateStructureEqual,
  directCriteriaOperandEqual,
  directCriteriaStructureEqual,
  directScalarDependencyCellsEqual,
  floatArrayEqual,
  hasStableSymbolicRangeLayout,
  stringArrayEqual,
  uint32ArrayEqual,
} from '../engine/services/formula-binding-shape-helpers.js'

function range(refKind: ParsedRangeReferenceInfo['refKind']): ParsedRangeReferenceInfo {
  return {
    address: 'A1:A3',
    kind: 'range',
    refKind,
    startAddress: 'A1',
    endAddress: 'A3',
    startRow: 0,
    endRow: 2,
    startCol: 0,
    endCol: 0,
  }
}

describe('formula binding shape helpers', () => {
  it('compares numeric and string dependency buffers by value', () => {
    expect(uint32ArrayEqual(Uint32Array.of(1, 2, 3), [1, 2, 3])).toBe(true)
    expect(uint32ArrayEqual(Uint32Array.of(1, 2, 3), [1, 3, 2])).toBe(false)
    expect(floatArrayEqual(Float64Array.of(1, 2), [1, 2])).toBe(true)
    expect(floatArrayEqual(Float64Array.of(1, 2), [1, 3])).toBe(false)
    expect(floatArrayEqual(Float64Array.of(1), [1, 2])).toBe(false)
    expect(stringArrayEqual(['Name', 'Table'], ['Name', 'Table'])).toBe(true)
    expect(stringArrayEqual(['Name'], ['Name', 'Table'])).toBe(false)
    expect(stringArrayEqual(['Name', 'Table'], ['Name', 'Range'])).toBe(false)
  })

  it('accepts symbolic range layouts only when dependency count and reference kinds are stable', () => {
    const current = {
      symbolicRanges: ['A1:A3'],
      parsedSymbolicRanges: [range('cells')],
    }
    expect(hasStableSymbolicRangeLayout(current, { ...current, parsedSymbolicRanges: [range('cells')] }, 1)).toBe(true)
    expect(hasStableSymbolicRangeLayout(current, { ...current, parsedSymbolicRanges: [range('rows')] }, 1)).toBe(false)
    expect(hasStableSymbolicRangeLayout(current, current, 0)).toBe(false)
    expect(
      hasStableSymbolicRangeLayout({ symbolicRanges: ['A1:A3'], parsedSymbolicRanges: undefined }, { symbolicRanges: ['A1:A3'] }, 1),
    ).toBe(true)
    expect(
      hasStableSymbolicRangeLayout(current, { symbolicRanges: ['A1:A3'], parsedSymbolicRanges: [range('cells'), range('rows')] }, 1),
    ).toBe(false)
  })

  it('compares direct scalar dependencies independent of commutative operand order', () => {
    const left: RuntimeDirectScalarDescriptor = {
      kind: 'binary',
      operator: '+',
      left: { kind: 'cell', cellIndex: 1 },
      right: { kind: 'cell', cellIndex: 2 },
    }
    const right: RuntimeDirectScalarDescriptor = {
      kind: 'binary',
      operator: '+',
      left: { kind: 'cell', cellIndex: 2 },
      right: { kind: 'cell', cellIndex: 1 },
    }

    expect(directScalarDependencyCellsEqual(left, right)).toBe(true)
    expect(directScalarDependencyCellsEqual({ kind: 'abs', operand: { kind: 'cell', cellIndex: 1 } }, left)).toBe(false)
    expect(directScalarDependencyCellsEqual(undefined, left)).toBe(false)
  })

  it('compares direct criteria structures including operands and result transforms', () => {
    const base: RuntimeDirectCriteriaDescriptor = {
      aggregateKind: 'sum',
      firstMatchMode: 'exact-lookup',
      aggregateRange: {
        regionId: 1,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 4,
        col: 1,
        length: 5,
      },
      offsetOperand: { kind: 'literal-number', value: 2 },
      resultTransforms: [
        { kind: 'if-empty-cell', cellIndex: 3, fallback: { tag: ValueTag.Number, value: 0 } },
        { kind: 'if-error', fallback: { tag: ValueTag.Error, code: ErrorCode.NA } },
      ],
      criteriaPairs: [
        {
          range: {
            regionId: 2,
            sheetName: 'Sheet1',
            rowStart: 0,
            rowEnd: 4,
            col: 2,
            length: 5,
          },
          criterion: { kind: 'cell-month-boundary-string-concat', cellIndex: 4, prefix: '>=', suffix: '', offsetMonths: 1 },
        },
      ],
    }

    expect(directCriteriaStructureEqual(base, { ...base })).toBe(true)
    expect(
      directCriteriaStructureEqual(base, {
        ...base,
        resultTransforms: [{ kind: 'if-empty-cell', cellIndex: 4, fallback: { tag: ValueTag.Number, value: 0 } }],
      }),
    ).toBe(false)
    expect(directCriteriaStructureEqual(base, { ...base, offsetOperand: { kind: 'error', code: ErrorCode.Ref } })).toBe(false)
    expect(
      directCriteriaOperandEqual(base.criteriaPairs[0].criterion, {
        kind: 'cell-month-boundary-string-concat',
        cellIndex: 4,
        prefix: '>=',
        suffix: '',
        offsetMonths: 2,
      }),
    ).toBe(false)
  })

  it('compares direct aggregate descriptors by structural fields', () => {
    const base: RuntimeDirectAggregateDescriptor = {
      regionId: 1,
      aggregateKind: 'sum',
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 9,
      col: 1,
      colEnd: 2,
      length: 10,
    }

    expect(directAggregateStructureEqual(base, { ...base })).toBe(true)
    expect(directAggregateStructureEqual(base, { ...base, colEnd: 3 })).toBe(false)
    expect(directAggregateStructureEqual(undefined, base)).toBe(false)
  })
})
