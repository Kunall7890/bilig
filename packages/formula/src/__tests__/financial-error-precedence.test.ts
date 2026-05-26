import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { getBuiltin } from '../builtins.js'

function num(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function err(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code }
}

describe('financial formula error precedence', () => {
  it('preserves incoming cash-flow errors before scalar coercion', () => {
    const CUMIPMT = getBuiltin('CUMIPMT')!
    const FVSCHEDULE = getBuiltin('FVSCHEDULE')!
    const NPV = getBuiltin('NPV')!
    const PV = getBuiltin('PV')!

    expect(NPV(err(ErrorCode.Name), num(100), num(200))).toEqual(err(ErrorCode.Name))
    expect(FVSCHEDULE(num(1000), err(ErrorCode.NA), num(0.1))).toEqual(err(ErrorCode.NA))
    expect(PV(err(ErrorCode.Ref), num(2), num(100))).toEqual(err(ErrorCode.Ref))
    expect(CUMIPMT(num(0.01), num(12), err(ErrorCode.Div0), num(1), num(3), num(0))).toEqual(err(ErrorCode.Div0))
  })

  it('preserves incoming rate-conversion errors before scalar coercion', () => {
    const EFFECT = getBuiltin('EFFECT')!
    const NOMINAL = getBuiltin('NOMINAL')!
    const PDURATION = getBuiltin('PDURATION')!
    const RRI = getBuiltin('RRI')!

    expect(EFFECT(err(ErrorCode.NA), num(12))).toEqual(err(ErrorCode.NA))
    expect(NOMINAL(num(0.1), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(PDURATION(num(0.1), err(ErrorCode.Div0), num(200))).toEqual(err(ErrorCode.Div0))
    expect(RRI(err(ErrorCode.Name), num(100), num(200))).toEqual(err(ErrorCode.Name))
  })
})
