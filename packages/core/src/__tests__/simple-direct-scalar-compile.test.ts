import { describe, expect, it } from 'vitest'
import { tryCompileSimpleDirectScalarFormula } from '../formula/simple-direct-scalar-compile.js'

describe('tryCompileSimpleDirectScalarFormula', () => {
  it('rejects unsafe integer row references on the direct compiler fast path', () => {
    const unsafeRow = String(Number.MAX_SAFE_INTEGER + 1)

    expect(tryCompileSimpleDirectScalarFormula(`A${unsafeRow}+1`)).toBeUndefined()
    expect(tryCompileSimpleDirectScalarFormula(`A1+B${unsafeRow}`)).toBeUndefined()
    expect(tryCompileSimpleDirectScalarFormula(`ABS(A${unsafeRow})`)).toBeUndefined()
  })
})
