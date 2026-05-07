import { FormulaMode } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { compileFormula } from '../compiler.js'

describe('OFFSET reference compilation', () => {
  it('rewrites static worksheet-reference OFFSET formulas to their target references', () => {
    const scalar = compileFormula('OFFSET(A1,1,1)')

    expect(scalar.mode).toBe(FormulaMode.WasmFastPath)
    expect(scalar.volatile).toBe(true)
    expect(scalar.deps).toEqual(['B2'])
    expect(scalar.optimizedAst).toEqual({ kind: 'CellRef', ref: 'B2' })

    const simple = compileFormula('SUM(OFFSET(B2:C2,4,0))')

    expect(simple.mode).toBe(FormulaMode.WasmFastPath)
    expect(simple.volatile).toBe(true)
    expect(simple.deps).toEqual(['B6:C6'])
    expect(simple.symbolicRanges).toEqual(['B6:C6'])

    const spilling = compileFormula('OFFSET(B2:C2,4,0)')

    expect(spilling.mode).toBe(FormulaMode.JsOnly)
    expect(spilling.volatile).toBe(true)
    expect(spilling.deps).toEqual(['B6:C6'])
    expect(spilling.optimizedAst).toEqual({ kind: 'RangeRef', refKind: 'cells', start: 'B6', end: 'C6' })
  })

  it('keeps dynamic worksheet-reference OFFSET formulas on the context-aware JS path', () => {
    const correlation = compileFormula('CORREL(OFFSET($C$15:$C$18,0,$K3),OFFSET($C$15:$C$18,0,D$11))')

    expect(correlation.mode).toBe(FormulaMode.JsOnly)
    expect(correlation.volatile).toBe(true)
    expect(correlation.deps).toEqual(['C15:C18', '$K3', 'D$11'])
  })
})
