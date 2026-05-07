import { describe, expect, it } from 'vitest'
import type { ParsedRangeReferenceInfo } from '@bilig/formula'
import { hasStableSymbolicRangeLayout, stringArrayEqual, uint32ArrayEqual } from '../engine/services/formula-binding-shape-helpers.js'

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
    expect(stringArrayEqual(['Name', 'Table'], ['Name', 'Table'])).toBe(true)
    expect(stringArrayEqual(['Name'], ['Name', 'Table'])).toBe(false)
  })

  it('accepts symbolic range layouts only when dependency count and reference kinds are stable', () => {
    const current = {
      symbolicRanges: ['A1:A3'],
      parsedSymbolicRanges: [range('cells')],
    }
    expect(hasStableSymbolicRangeLayout(current, { ...current, parsedSymbolicRanges: [range('cells')] }, 1)).toBe(true)
    expect(hasStableSymbolicRangeLayout(current, { ...current, parsedSymbolicRanges: [range('rows')] }, 1)).toBe(false)
    expect(hasStableSymbolicRangeLayout(current, current, 0)).toBe(false)
  })
})
