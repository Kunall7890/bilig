import { describe, expect, it } from 'vitest'
import { tryCompileSimpleDirectScalarFormula } from '../formula/simple-direct-scalar-compile.js'

describe('tryCompileSimpleDirectScalarFormula', () => {
  it('compiles unquoted sheet-qualified binary references on the direct compiler fast path', () => {
    const compiled = tryCompileSimpleDirectScalarFormula('Data1!A250+Data1!B250')

    expect(compiled?.symbolicRefs).toEqual(['Data1!A250', 'Data1!B250'])
    expect(compiled?.parsedSymbolicRefs).toEqual([
      {
        address: 'Data1!A250',
        sheetName: 'Data1',
        explicitSheet: true,
        row: 249,
        col: 0,
        rowAbsolute: false,
        colAbsolute: false,
      },
      {
        address: 'Data1!B250',
        sheetName: 'Data1',
        explicitSheet: true,
        row: 249,
        col: 1,
        rowAbsolute: false,
        colAbsolute: false,
      },
    ])
    expect(compiled?.optimizedAst).toEqual({
      kind: 'BinaryExpr',
      operator: '+',
      left: { kind: 'CellRef', ref: 'A250', sheetName: 'Data1' },
      right: { kind: 'CellRef', ref: 'B250', sheetName: 'Data1' },
    })
  })

  it('rejects unsafe integer row references on the direct compiler fast path', () => {
    const unsafeRow = String(Number.MAX_SAFE_INTEGER + 1)

    expect(tryCompileSimpleDirectScalarFormula(`A${unsafeRow}+1`)).toBeUndefined()
    expect(tryCompileSimpleDirectScalarFormula(`A1+B${unsafeRow}`)).toBeUndefined()
    expect(tryCompileSimpleDirectScalarFormula(`ABS(A${unsafeRow})`)).toBeUndefined()
  })
})
