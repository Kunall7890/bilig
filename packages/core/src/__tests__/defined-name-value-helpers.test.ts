import { describe, expect, it } from 'vitest'
import { isScalarOnlyDefinedNameValue } from '../engine/services/defined-name-value-helpers.js'

describe('defined name value helpers', () => {
  it('treats literals and scalar snapshots as scalar-only', () => {
    expect(isScalarOnlyDefinedNameValue(12)).toBe(true)
    expect(isScalarOnlyDefinedNameValue('plain text')).toBe(true)
    expect(isScalarOnlyDefinedNameValue(true)).toBe(true)
    expect(isScalarOnlyDefinedNameValue(null)).toBe(true)
    expect(isScalarOnlyDefinedNameValue({ kind: 'scalar', value: '=1+1' })).toBe(true)
  })

  it('classifies formula snapshots by whether the parsed expression is scalar-only', () => {
    expect(isScalarOnlyDefinedNameValue('=1')).toBe(true)
    expect(isScalarOnlyDefinedNameValue('="x"')).toBe(true)
    expect(isScalarOnlyDefinedNameValue('=')).toBe(false)
    expect(isScalarOnlyDefinedNameValue({ kind: 'formula', formula: '=TRUE' })).toBe(true)
    expect(isScalarOnlyDefinedNameValue({ kind: 'formula', formula: '=SUM(A1:A2)' })).toBe(false)
    expect(isScalarOnlyDefinedNameValue({ kind: 'formula', formula: '=A1' })).toBe(false)
    expect(isScalarOnlyDefinedNameValue({ kind: 'formula', formula: '=' })).toBe(false)
  })

  it('rejects non-scalar workbook references', () => {
    expect(isScalarOnlyDefinedNameValue({ kind: 'cell-ref', sheetName: 'Sheet1', address: 'A1' })).toBe(false)
    expect(isScalarOnlyDefinedNameValue({ kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A3' })).toBe(false)
    expect(isScalarOnlyDefinedNameValue({ kind: 'structured-ref', tableName: 'Sales', columnName: 'Amount' })).toBe(false)
  })
})
