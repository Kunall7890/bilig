import { describe, expect, it } from 'vitest'

import { compareCellAddresses, readImportedLiteralCellValue, readImportedNumberFormat } from '../xlsx-import-cell-values.js'

describe('xlsx imported cell values', () => {
  it('maps legacy Excel numeric error codes to literal error text', () => {
    expect(readImportedLiteralCellValue({ t: 'e', v: 0 })).toBe('#NULL!')
    expect(readImportedLiteralCellValue({ t: 'e', v: 7 })).toBe('#DIV/0!')
    expect(readImportedLiteralCellValue({ t: 'e', v: 15 })).toBe('#VALUE!')
    expect(readImportedLiteralCellValue({ t: 'e', v: 23 })).toBe('#REF!')
    expect(readImportedLiteralCellValue({ t: 'e', v: 29 })).toBe('#NAME?')
    expect(readImportedLiteralCellValue({ t: 'e', v: 36 })).toBe('#NUM!')
    expect(readImportedLiteralCellValue({ t: 'e', v: 42 })).toBe('#N/A')
    expect(readImportedLiteralCellValue({ t: 'e', v: 43 })).toBe('#GETTING_DATA')
  })

  it('prefers explicit error display text when present', () => {
    expect(readImportedLiteralCellValue({ t: 'e', v: 99, w: '#SPILL!' })).toBe('#SPILL!')
    expect(readImportedLiteralCellValue({ t: 'e', v: '#CALC!' })).toBe('#CALC!')
  })

  it('uses a generic error marker for unrecognized error payloads', () => {
    expect(readImportedLiteralCellValue({ t: 'e', v: 99 })).toBe('#ERROR!')
    expect(readImportedLiteralCellValue({ t: 'e', v: 'not-an-error' })).toBe('#ERROR!')
    expect(readImportedLiteralCellValue({ t: 'e' })).toBeUndefined()
  })

  it('normalizes imported number formats', () => {
    expect(readImportedNumberFormat(undefined)).toBeUndefined()
    expect(readImportedNumberFormat('')).toBeUndefined()
    expect(readImportedNumberFormat('  General  ')).toBeUndefined()
    expect(readImportedNumberFormat('  m/d/yyyy  ')).toBe('m/d/yyyy')
  })

  it('sorts cell addresses in row-major order', () => {
    expect(['B1', 'A2', 'A1'].toSorted(compareCellAddresses)).toEqual(['A1', 'B1', 'A2'])
  })
})
